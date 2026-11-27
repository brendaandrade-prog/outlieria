from __future__ import annotations

import io
import html
import gzip
import json
import math
import os
import unicodedata
import zipfile
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from dataclasses import dataclass
from typing import Iterable
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import shapefile
import streamlit as st
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from shapely.geometry import Point, shape as shapely_shape
from cloud_storage import (
    configured as cloud_storage_configured,
    database_configured,
    delete_analysis,
    download_persisted_source,
    load_persisted_analyses,
    missing_settings,
    persist_analysis,
    sha256,
)
SKLEARN_IMPORT_ERROR = ""
try:
    from sklearn.neighbors import LocalOutlierFactor
    from sklearn.preprocessing import RobustScaler
    from sklearn.linear_model import TheilSenRegressor
    LOF_AVAILABLE = True
except Exception as exc:
    LocalOutlierFactor = None
    RobustScaler = None
    TheilSenRegressor = None
    LOF_AVAILABLE = False
    SKLEARN_IMPORT_ERROR = f"{type(exc).__name__}: {exc}"

YEAR_RE = r"^\d{4}$"
CURRENT_YEAR = "2026"
DEFAULT_RADIUS_M = 300.0
MAX_DISTANCE_M = DEFAULT_RADIUS_M  # compatibilidade com análises antigas; o cálculo usa o raio informado.
LOF_THRESHOLD_PERCENTILE = 90.0
ATTENTION_MARGIN = 0.0
MIN_LOF_NEIGHBORS = 3
MIN_REGRESSION_CYCLES = 3
PROCESS_EXECUTOR = ThreadPoolExecutor(max_workers=1, thread_name_prefix="outlieria-process")
# Posições físicas do Excel, em índice zero: A, B, G, H e S–W.
# Isso não significa ignorar um elemento cujo cabeçalho seja "S" (Enxofre).
IGNORED_EXCEL_COLUMN_INDICES = {0, 1, 6, 7, *range(18, 23)}
IGNORED_FROM_INDEX = 28  # AC, zero based
IGNORED_ELEMENTS = {"H+AL", "CTC EFETIVA", "CTC", "SB", "SA", "SK", "RELEVO", "SCA", "SMG", "CA/MG"}
ELEMENT_LABELS = {
    "PH": "pH",
    "SMP": "Índice SMP",
    "P": "Fósforo",
    "K": "Potássio",
    "MO": "Matéria orgânica",
    "AL": "Alumínio",
    "CA": "Cálcio",
    "MG": "Magnésio",
    "H+Al": "Hidrogênio + alumínio",
    "CTC Efetiva": "CTC efetiva",
    "CTC": "Capacidade de troca catiônica",
    "SB": "Soma de bases",
    "SA": "Saturação por alumínio",
    "ZN": "Zinco",
    "CU": "Cobre",
    "B": "Boro",
    "MN": "Manganês",
    "Argila": "Argila",
}
LOF_FEATURES = [
    "value", "regional_mean", "regional_median",
    "regional_std", "neighbor_count", "mean_distance_m",
    "nearest_distance_m", "difference_median", "z_score",
    "historical_percentile", "historical_distribution_shift",
    "current_cohort_shift",
]

@dataclass
class PolygonRecord:
    geometry: object
    area_ha: float
    producer_name: str = ""
    farm_name: str = ""
    field_name: str = ""


@dataclass
class IBGEMunicipality:
    geometry: object
    municipality: str
    state: str
    state_code: str
    bounds: tuple[float, float, float, float]


def haversine(lon1, lat1, lon2, lat2):
    r = 6_371_000.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.asin(math.sqrt(a))


def haversine_vector(lon, lat, longitudes, latitudes):
    """Calcula várias distâncias Haversine sem apply linha a linha."""
    radius = 6_371_000.0
    p1 = np.radians(float(lat))
    p2 = np.radians(np.asarray(latitudes, dtype=float))
    dp = p2 - p1
    dl = np.radians(np.asarray(longitudes, dtype=float) - float(lon))
    a = np.sin(dp / 2) ** 2 + np.cos(p1) * np.cos(p2) * np.sin(dl / 2) ** 2
    return 2 * radius * np.arcsin(np.sqrt(np.clip(a, 0, 1)))


def validate_radius(value) -> float:
    """Converte e valida o raio informado antes de qualquer cálculo espacial."""
    try:
        radius = float(value)
    except (TypeError, ValueError):
        raise ValueError("Informe um raio numérico válido em metros.")
    if not math.isfinite(radius) or radius <= 0:
        raise ValueError("O raio deve ser um número positivo e finito, maior que zero.")
    if radius > 100_000:
        raise ValueError("Informe um raio de até 100.000 metros.")
    return radius


def polygon_area_ha(geometry):
    # SHP coordinates are assumed to be longitude/latitude (WGS84).
    minx, miny, maxx, maxy = geometry.bounds
    lat = (miny + maxy) / 2
    mx = 111_320 * math.cos(math.radians(lat))
    my = 111_320
    area_m2 = 0.0
    for ring in getattr(geometry, "geoms", [geometry]):
        coords = list(ring.exterior.coords) if hasattr(ring, "exterior") else list(ring.coords)
        for (x1, y1), (x2, y2) in zip(coords, coords[1:]):
            area_m2 += (x1 * mx) * (y2 * my) - (x2 * mx) * (y1 * my)
    return abs(area_m2) / 2 / 10_000


def read_polygons(shp_bytes: bytes) -> list[PolygonRecord]:
    archive = zipfile.ZipFile(io.BytesIO(shp_bytes)) if zipfile.is_zipfile(io.BytesIO(shp_bytes)) else None
    if archive:
        members = {name.lower(): name for name in archive.namelist() if not name.endswith("/")}
        shp_member = next((name for key, name in members.items() if key.endswith(".shp")), None)
        if not shp_member:
            raise ValueError("O ZIP não contém um arquivo .shp.")
        stem = shp_member[:-4].lower()
        reader = shapefile.Reader(
            shp=io.BytesIO(archive.read(shp_member)),
            shx=io.BytesIO(archive.read(members[stem + ".shx"])) if stem + ".shx" in members else None,
            dbf=io.BytesIO(archive.read(members[stem + ".dbf"])) if stem + ".dbf" in members else None,
        )
    else:
        reader = shapefile.Reader(shp=io.BytesIO(shp_bytes))
    shapes = reader.shapes()
    field_names: list[str] = []
    attribute_rows: list[list[object]] = []
    try:
        field_names = [shp_attribute_text(field[0]) for field in reader.fields[1:]]
        attribute_rows = reader.records()
    except (shapefile.ShapefileException, ValueError):
        # Um .shp isolado não possui a tabela DBF. A geometria continua útil;
        # os nomes serão obtidos do Excel ou exibidos como não informados.
        field_names = []
        attribute_rows = []
    field_keys = {normalise_header_key(name): name for name in field_names}
    producer_field = next((field_keys[key] for key in ("PRODUTOR", "PRODUCER") if key in field_keys), None)
    farm_field = next((field_keys[key] for key in ("FAZENDA", "FARM", "PROPERTY") if key in field_keys), None)
    # Para o gráfico, somente o atributo CAMPO do SHP é válido.
    field_field = field_keys.get("CAMPO")
    records = []
    for index, shp in enumerate(shapes):
        geometry = shapely_shape(shp.__geo_interface__)
        attributes = attribute_rows[index] if index < len(attribute_rows) else []
        attribute_map = dict(zip(field_names, attributes))
        producer = shp_attribute_text(attribute_map.get(producer_field)) if producer_field else ""
        farm = shp_attribute_text(attribute_map.get(farm_field)) if farm_field else ""
        field = shp_attribute_text(attribute_map.get(field_field)) if field_field else ""
        records.append(PolygonRecord(geometry, polygon_area_ha(geometry), producer, farm, field))
    if archive:
        archive.close()
    return records


def shp_source_names(polygons: list[PolygonRecord]) -> tuple[str, str]:
    """Obtém produtor e fazenda do primeiro atributo preenchido no SHP."""
    producer = next((item.producer_name for item in polygons if str(item.producer_name).strip()), "")
    farm = next((item.farm_name for item in polygons if str(item.farm_name).strip()), "")
    return str(producer).strip(), str(farm).strip()


@st.cache_resource(show_spinner=False)
def load_ibge_municipalities() -> list[IBGEMunicipality]:
    path = os.path.join(os.path.dirname(__file__), "data", "ibge", "ibge_municipios_2025.geojson.gz")
    if not os.path.exists(path):
        return []
    with gzip.open(path, "rt", encoding="utf-8") as source:
        collection = json.load(source)
    boundaries = []
    for feature in collection.get("features", []):
        geometry = shapely_shape(feature["geometry"])
        props = feature.get("properties", {})
        boundaries.append(IBGEMunicipality(
            geometry=geometry,
            municipality=str(props.get("NM_MUN") or "Não localizado"),
            state=str(props.get("NM_UF") or "Não localizado"),
            state_code=str(props.get("SIGLA_UF") or ""),
            bounds=geometry.bounds,
        ))
    return boundaries


def find_header_row(rows: list[list[object]]) -> int:
    for i, row in enumerate(rows[:12]):
        names = {normalise_header_key(x) for x in row}
        if names.intersection({"LONGITUDE", "LONG", "X"}) and names.intersection({"LATITUDE", "LAT", "Y"}):
            return i
    return 0


def normalise_columns(columns: Iterable[object]) -> list[str]:
    result = []
    for col in columns:
        text = str(col or "").strip()
        result.append(text if text else f"coluna_{len(result) + 1}")
    return result


def normalise_header_key(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or "")).encode("ascii", "ignore").decode()
    return "".join(character for character in text.upper().strip() if character.isalnum())


def shp_attribute_text(value: object) -> str:
    """Converte atributos DBF sem descartar valores zero ou bytes."""
    if value is None:
        return ""
    if isinstance(value, bytes):
        # Nomes e valores DBF podem chegar como bytes quando o arquivo não
        # informa claramente a codificação. Latin-1 preserva os bytes sem
        # introduzir o caractere de substituição do UTF-8.
        try:
            value = value.decode("utf-8")
        except UnicodeDecodeError:
            value = value.decode("latin-1", errors="replace")
    text = str(value).replace("\x00", "").strip()
    return "" if text.lower() in {"", "none", "nan"} else text


def canonicalise_coordinate_columns(frame: pd.DataFrame) -> pd.DataFrame:
    """Aceita variações comuns de nomes sem alterar as demais colunas da planilha."""
    aliases = {
        "Longitude": {"LONGITUDE", "LONG", "X", "COORDENADAX", "LON"},
        "Latitude": {"LATITUDE", "LAT", "Y", "COORDENADAY"},
        "Profundidade": {"PROFUNDIDADE", "DEPTH", "CAMADA"},
    }
    rename = {}
    used = set()
    for column in frame.columns:
        key = normalise_header_key(column)
        for canonical, candidates in aliases.items():
            if key in candidates and canonical not in used:
                rename[column] = canonical
                used.add(canonical)
                break
    return frame.rename(columns=rename)


def parse_numeric_values(values: pd.Series) -> pd.Series:
    """Converte números Excel e textos com vírgula decimal."""
    numeric = pd.to_numeric(values, errors="coerce")
    missing = numeric.isna()
    if missing.any():
        text = values.astype(str).str.strip().str.replace("\u00a0", "", regex=False)
        text = text.str.replace(",", ".", regex=False).str.replace(r"[^0-9+\-.eE]", "", regex=True)
        numeric.loc[missing] = pd.to_numeric(text.loc[missing], errors="coerce")
    return numeric


def numeric_element_columns(frame: pd.DataFrame) -> list[str]:
    excluded = {"Registro", "Descrição antiga", "Fazenda", "Longitude", "Latitude", "Área/Lavoura", "Campo", "Pontos", "ID", "Profundidade"}
    elements = []
    for index, name in enumerate(frame.columns):
        normalized_name = str(name).strip().upper()
        if str(name).startswith("__") or index in IGNORED_EXCEL_COLUMN_INDICES or index >= IGNORED_FROM_INDEX or name in excluded or normalized_name in {normalise_header_key(item) for item in excluded} or is_ignored_element(name):
            continue
        numeric = pd.to_numeric(frame[name], errors="coerce")
        if numeric.notna().any():
            elements.append(name.strip())
    return elements


def parse_workbook(data: bytes, polygons: list[PolygonRecord], client_override: str, farm_override: str):
    workbook = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    detected_client = "Cliente não identificado"
    detected_farm = ""
    if "Resumo de Amostragem" in workbook.sheetnames:
        title = str(workbook["Resumo de Amostragem"].cell(1, 1).value or "")
        if " - " in title:
            detected_client, detected_farm = [part.strip() for part in title.split(" - ", 1)]
    annual = [name for name in workbook.sheetnames if str(name).isdigit() and len(str(name)) == 4]
    frames: dict[str, pd.DataFrame] = {}
    for year in annual:
        rows = [list(row) for row in workbook[year].iter_rows(values_only=True)]
        header = find_header_row(rows)
        if header >= len(rows):
            continue
        columns = normalise_columns(rows[header])
        frame = pd.DataFrame(rows[header + 1 :], columns=columns)
        frame = canonicalise_coordinate_columns(frame)
        frame = frame.dropna(how="all")
        frame["__year"] = str(year)
        frame["__excel_row"] = np.arange(header + 2, header + 2 + len(frame))
        frame["__header_row"] = header
        frame["__sheet"] = str(year)
        frames[str(year)] = frame
    if not frames:
        raise ValueError("Nenhuma aba anual YYYY foi encontrada.")
    elements = numeric_element_columns(next(iter(frames.values())))
    all_rows = []
    coordinate_rows = 0
    valid_coordinate_rows = 0
    coordinate_examples = []
    for year, frame in frames.items():
        required = {"Longitude", "Latitude", "Profundidade"}
        if not required.issubset(frame.columns):
            continue
        frame["Longitude"] = parse_numeric_values(frame["Longitude"])
        frame["Latitude"] = parse_numeric_values(frame["Latitude"])
        frame = frame.dropna(subset=["Longitude", "Latitude"])
        coordinate_rows += len(frame)
        for _, row in frame.iterrows():
            longitude = float(row["Longitude"])
            latitude = float(row["Latitude"])
            if not (-180 <= longitude <= 180 and -90 <= latitude <= 90):
                continue
            valid_coordinate_rows += 1
            point = Point(longitude, latitude)
            polygon_index = next((i for i, p in enumerate(polygons) if p.geometry.contains(point) or p.geometry.touches(point)), -1)
            if polygon_index < 0 and -90 <= longitude <= 90 and -180 <= latitude <= 180:
                swapped_point = Point(latitude, longitude)
                swapped_index = next((i for i, p in enumerate(polygons) if p.geometry.contains(swapped_point) or p.geometry.touches(swapped_point)), -1)
                if swapped_index >= 0:
                    longitude, latitude, point, polygon_index = latitude, longitude, swapped_point, swapped_index
            if polygon_index < 0:
                if len(coordinate_examples) < 3:
                    coordinate_examples.append(f"{longitude:.6f}, {latitude:.6f}")
                continue
            campo_column = next((column for column in row.index if normalise_header_key(column) == "CAMPO"), None)
            area_column = next((column for column in row.index if normalise_header_key(column) in {"AREALAVOURA", "LAVOURA", "AREA"}), None)
            crop = str(row.get(campo_column) if campo_column else row.get(area_column) if area_column else "Lavoura não identificada")
            if not crop or crop == "nan":
                crop = "Lavoura não identificada"
            # Mantém no processamento somente os elementos válidos. As colunas
            # S–W e AC em diante continuam existindo apenas no arquivo original.
            analysis_raw = {element: row.get(element) for element in elements}
            all_rows.append({
                "year": year, "source_row": int(row["__excel_row"]), "sheet": year,
                "point_id": str(row.get("ID") or ""), "client": client_override or detected_client,
                "farm": farm_override or str(row.get("Fazenda") or detected_farm or "Fazenda não identificada"),
                "crop": crop, "depth": str(row.get("Profundidade") or "Não informado"),
                "longitude": longitude, "latitude": latitude,
                "polygon_index": polygon_index, "raw": analysis_raw,
            })
    if not all_rows:
        if coordinate_rows == 0:
            raise ValueError("Nenhuma coordenada numérica foi encontrada. Verifique os cabeçalhos Longitude/Latitude e se os valores usam vírgula ou ponto decimal.")
        bounds = ", ".join(f"{value:.6f}" for value in polygons[0].geometry.bounds) if polygons else "SHP vazio"
        examples = "; ".join(coordinate_examples) or "sem exemplos válidos"
        raise ValueError(
            f"Nenhum ponto com coordenadas foi encontrado dentro do SHP. A planilha tem {valid_coordinate_rows} coordenadas válidas, mas 0 ficaram dentro do limite. "
            f"Exemplos fora: {examples}. Faixa do primeiro polígono SHP (minX, minY, maxX, maxY): {bounds}. "
            "Confirme que planilha e SHP usam o mesmo sistema de coordenadas, que o SHP corresponde à mesma fazenda e que o limite está em longitude/latitude WGS84."
        )
    sample = pd.DataFrame(all_rows)
    for element in elements:
        sample[element] = pd.to_numeric(sample["raw"].map(lambda x: x.get(element)), errors="coerce")
    return sample, elements, frames


def classify(sample: pd.DataFrame, elements: list[str], radius_m: float = MAX_DISTANCE_M) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Cria a matriz de candidatos 2026; a classificação vem de LOF + regressão."""
    current = sample[sample.year == CURRENT_YEAR].copy()
    results, evidence, attention = [], [], []
    for idx, row in current.iterrows():
        for element in elements:
            # Segurança adicional: mesmo que uma planilha ou sessão antiga
            # carregue uma dessas colunas, ela nunca entra nos resultados.
            if is_ignored_element(element):
                continue
            value = row.get(element)
            if pd.isna(value):
                continue
            result = {**row.to_dict(), "sample_index": idx, "element": element, "value": float(value), "status": "normal", "direction": ""}
            results.append(result)
    return pd.DataFrame(results), pd.DataFrame(evidence), pd.DataFrame(attention)


def eligible_lof_neighbors(target: pd.Series, reference: pd.DataFrame, element: str, radius_m: float, exclude_index=None) -> pd.DataFrame:
    candidates = reference[
        (reference["crop"].astype(str) == str(target["crop"])) &
        (reference["depth"].astype(str) == str(target["depth"]))
    ].copy()
    if exclude_index is not None:
        candidates = candidates[candidates.index != exclude_index]
    if candidates.empty:
        return candidates
    candidates["distance_m"] = haversine_vector(
        float(target["longitude"]), float(target["latitude"]), candidates["longitude"].to_numpy(), candidates["latitude"].to_numpy()
    )
    candidates = candidates[candidates["distance_m"] <= radius_m].copy()
    candidates[element] = pd.to_numeric(candidates[element], errors="coerce")
    return candidates.dropna(subset=[element])


def build_lof_features(
    target: pd.Series,
    reference: pd.DataFrame,
    element: str,
    radius_m: float,
    exclude_index=None,
    minimum_neighbors: int = MIN_LOF_NEIGHBORS,
    cohort_values=None,
    historical_distribution_values=None,
):
    """Monta o vetor regional de um ponto usando apenas seu histórico local."""
    neighbors = eligible_lof_neighbors(target, reference, element, radius_m, exclude_index)
    values = neighbors[element]
    value = pd.to_numeric(pd.Series([target[element]]), errors="coerce").iloc[0]
    if pd.isna(value) or len(values) < minimum_neighbors:
        return None
    mean = values.mean()
    median = values.median()
    std = values.std(ddof=0)
    percentile = float((values <= value).mean())
    # A distribuição histórica local é a referência principal.  O segundo
    # componente é calculado sobre o lote de 2026, mas nunca usa os valores de
    # 2026 para treinar o modelo: ele mede uma mudança coletiva da região.
    distribution = pd.to_numeric(pd.Series(
        historical_distribution_values if historical_distribution_values is not None else values.tolist()
    ), errors="coerce").dropna()
    distribution_median = float(distribution.median()) if not distribution.empty else float(median)
    distribution_mad = float(np.median(np.abs(distribution - distribution_median))) if not distribution.empty else 0.0
    distribution_scale = max(1.4826 * distribution_mad, float(distribution.std(ddof=0)) if len(distribution) > 1 else 0.0, 1e-6)
    historical_shift = abs(float(value) - distribution_median) / distribution_scale
    cohort_shift = 0.0
    if cohort_values is not None:
        cohort = pd.to_numeric(pd.Series(cohort_values), errors="coerce").dropna()
        if not cohort.empty:
            cohort_shift = abs(float(cohort.median()) - distribution_median) / distribution_scale
    return {
        "value": float(value),
        "regional_mean": float(mean),
        "regional_median": float(median),
        "regional_std": float(std),
        "neighbor_count": int(len(values)),
        "mean_distance_m": float(neighbors["distance_m"].mean()),
        "nearest_distance_m": float(neighbors["distance_m"].min()),
        "difference_median": float(value - median),
        "z_score": float((value - mean) / std) if std > 0 else 0.0,
        "historical_percentile": percentile,
        "historical_distribution_shift": float(historical_shift),
        "current_cohort_shift": float(cohort_shift),
    }


def train_lof_group(historical: pd.DataFrame, element: str, radius_m: float, minimum_neighbors: int = MIN_LOF_NEIGHBORS):
    feature_rows = []
    for index, target in historical.iterrows():
        features = build_lof_features(target, historical, element, radius_m, exclude_index=index, minimum_neighbors=minimum_neighbors)
        if features is not None:
            features["sample_index"] = index
            feature_rows.append(features)
    feature_frame = pd.DataFrame(feature_rows)
    if len(feature_frame) < minimum_neighbors + 1:
        return None
    matrix = feature_frame[LOF_FEATURES].replace([np.inf, -np.inf], np.nan).dropna()
    if len(matrix) < minimum_neighbors + 1:
        return None
    scaler = RobustScaler()
    scaled_matrix = scaler.fit_transform(matrix)
    n_neighbors = min(20, len(matrix) - 1)
    if n_neighbors < minimum_neighbors:
        return None
    model = LocalOutlierFactor(n_neighbors=n_neighbors, contamination="auto", novelty=True)
    model.fit(scaled_matrix)
    historical_scores = -model.score_samples(scaled_matrix)
    return {"model": model, "scaler": scaler, "threshold": float(np.percentile(historical_scores, LOF_THRESHOLD_PERCENTILE)), "n_neighbors": n_neighbors}


def score_lof_current(
    current: pd.DataFrame,
    historical: pd.DataFrame,
    element: str,
    radius_m: float,
    trained: dict,
    minimum_neighbors: int = MIN_LOF_NEIGHBORS,
    cohort_values=None,
    historical_distribution_values=None,
):
    predictions = []
    for index, target in current.iterrows():
        features = build_lof_features(
            target, historical, element, radius_m,
            minimum_neighbors=minimum_neighbors,
            cohort_values=cohort_values,
            historical_distribution_values=historical_distribution_values,
        )
        if features is None:
            continue
        vector = pd.DataFrame([features])[LOF_FEATURES].replace([np.inf, -np.inf], np.nan)
        if vector.isna().any().any():
            continue
        score = float(-trained["model"].score_samples(trained["scaler"].transform(vector))[0])
        collective_shift = float(features["current_cohort_shift"])
        distribution_shift = float(features["historical_distribution_shift"])
        # O LOF tradicional pode considerar vários pontos igualmente baixos
        # como uma região densa e, portanto, normal.  O deslocamento robusto
        # abaixo preserva a decisão do LOF, mas impede que uma mudança coletiva
        # em 2026 seja escondida pela densidade interna do próprio ciclo.
        collective_anomaly = max(collective_shift, distribution_shift) >= 3.0
        is_anomaly = score >= trained["threshold"] or collective_anomaly
        predictions.append({
            "sample_index": index, "element": element, "ml_score": score,
            "ml_status": "anomalia espacial" if is_anomaly else "normal",
            "ml_neighbor_count": features["neighbor_count"],
            "ml_k_effective": trained["n_neighbors"],
            "ml_radius_m": radius_m,
            "ml_rule": f"LOF ≥ percentil {LOF_THRESHOLD_PERCENTILE:g} ou deslocamento histórico robusto ≥ 3; k={trained['n_neighbors']}; raio={radius_m:g} m",
            "ml_reference_scope": "conjunto 2026 comparado à distribuição histórica espacial da mesma região; sem pareamento por ID",
            "ml_mean_distance_m": features["mean_distance_m"],
            "ml_regional_median": features["regional_median"],
            "ml_historical_percentile": features["historical_percentile"],
            "ml_historical_distribution_shift": distribution_shift,
            "ml_current_cohort_shift": collective_shift,
        })
    return pd.DataFrame(predictions)


def run_spatial_lof(sample: pd.DataFrame, elements: list[str], radius_m: float) -> pd.DataFrame:
    """Treina e aplica LOF por lavoura, profundidade e elemento, sem usar 2026 no treino."""
    if not LOF_AVAILABLE:
        # O LOF é opcional para não impedir a importação da planilha quando o
        # ambiente ainda está instalando ou não consegue carregar o sklearn.
        return pd.DataFrame()
    historical = sample[sample["year"].astype(str) != CURRENT_YEAR].copy()
    current = sample[sample["year"].astype(str) == CURRENT_YEAR].copy()
    predictions = []
    for element in elements:
        if is_ignored_element(element):
            continue
        for crop in current["crop"].dropna().unique():
            for depth in current["depth"].dropna().unique():
                historical_group = historical[(historical["crop"].astype(str) == str(crop)) & (historical["depth"].astype(str) == str(depth))].copy()
                current_group = current[(current["crop"].astype(str) == str(crop)) & (current["depth"].astype(str) == str(depth))].copy()
                if current_group.empty:
                    continue
                trained = train_lof_group(historical_group, element, radius_m, MIN_LOF_NEIGHBORS) if not historical_group.empty else None
                fallback_rows = []
                for current_index, target in current_group.iterrows():
                    neighbors = eligible_lof_neighbors(target, historical_group, element, radius_m)
                    neighbor_count = len(neighbors)
                    if neighbor_count < MIN_LOF_NEIGHBORS:
                        fallback_rows.append({"sample_index": current_index, "element": element, "ml_status": "dados insuficientes", "ml_score": np.nan, "ml_neighbor_count": neighbor_count, "ml_k_effective": np.nan, "ml_radius_m": radius_m, "ml_rule": f"LOF não calculado: {neighbor_count} vizinhos elegíveis; mínimo={MIN_LOF_NEIGHBORS}; raio={radius_m:g} m", "ml_mean_distance_m": np.nan, "ml_regional_median": np.nan, "ml_reference_scope": "distribuição histórica espacial da mesma região", "ml_historical_percentile": np.nan, "ml_historical_distribution_shift": np.nan, "ml_current_cohort_shift": np.nan})
                    elif trained is None:
                        fallback_rows.append({"sample_index": current_index, "element": element, "ml_status": "dados insuficientes", "ml_score": np.nan, "ml_neighbor_count": neighbor_count, "ml_k_effective": np.nan, "ml_radius_m": radius_m, "ml_rule": f"LOF não calculado: histórico treinável insuficiente para k≥{MIN_LOF_NEIGHBORS}; raio={radius_m:g} m", "ml_mean_distance_m": np.nan, "ml_regional_median": np.nan, "ml_reference_scope": "distribuição histórica espacial da mesma região", "ml_historical_percentile": np.nan, "ml_historical_distribution_shift": np.nan, "ml_current_cohort_shift": np.nan})
                # Monta uma distribuição histórica espacial única para o lote
                # de 2026. Cada registro histórico entra no conjunto uma vez,
                # mesmo que seja vizinho de mais de um ponto atual.
                historical_distribution = []
                current_values = pd.to_numeric(current_group[element], errors="coerce").dropna().tolist()
                for _, current_target in current_group.iterrows():
                    local = eligible_lof_neighbors(current_target, historical_group, element, radius_m)
                    historical_distribution.extend(local[element].tolist())
                historical_distribution = pd.to_numeric(pd.Series(historical_distribution), errors="coerce").dropna().drop_duplicates().tolist()
                scored = score_lof_current(
                    current_group, historical_group, element, radius_m, trained,
                    MIN_LOF_NEIGHBORS, cohort_values=current_values,
                    historical_distribution_values=historical_distribution,
                ) if trained is not None else pd.DataFrame()
                if fallback_rows:
                    scored = pd.concat([scored, pd.DataFrame(fallback_rows)], ignore_index=True)
                if not scored.empty:
                    scored["crop"] = crop
                    scored["depth"] = depth
                    predictions.append(scored)
    return pd.concat(predictions, ignore_index=True) if predictions else pd.DataFrame()


def combine_detection(row):
    # Os dois critérios precisam apontar a mesma anomalia para confirmar um
    # outlier. A interface exibe somente essa confirmação simultânea.
    ml_outlier = row.get("ml_status") == "anomalia espacial"
    regression_outlier = row.get("reg_status") == "anomalia temporal"
    criteria_count = sum((ml_outlier, regression_outlier))
    if criteria_count == 2:
        return "outlier confirmado"
    if criteria_count == 1 or row.get("status") == "atenção":
        return "atenção"
    return "normal"


def spatial_temporal_reference(target: pd.Series, historical: pd.DataFrame, element: str, radius_m: float) -> list[tuple[int, float, float]]:
    """Reconstrói a série da coordenada usando os registros mais próximos de cada ciclo.

    O ID não participa da seleção. Em cada ano, os até cinco pontos mais próximos
    da coordenada do alvo são combinados por peso inverso da distância.
    """
    candidates = historical[
        (historical["crop"].astype(str) == str(target["crop"])) &
        (historical["depth"].astype(str) == str(target["depth"]))
    ].copy()
    if candidates.empty:
        return []
    candidates["value_numeric"] = pd.to_numeric(candidates[element], errors="coerce")
    candidates = candidates.dropna(subset=["value_numeric", "longitude", "latitude"])
    if candidates.empty:
        return []
    candidates["distance_m"] = haversine_vector(
        float(target["longitude"]), float(target["latitude"]), candidates["longitude"].to_numpy(), candidates["latitude"].to_numpy()
    )
    candidates = candidates[candidates["distance_m"] <= radius_m]
    if candidates.empty:
        return []
    series = []
    for year, group in candidates.groupby(candidates["year"].astype(int)):
        nearest = group.sort_values("distance_m").head(5)
        distances = nearest["distance_m"].to_numpy(dtype=float)
        values = nearest["value_numeric"].to_numpy(dtype=float)
        if np.any(distances == 0):
            estimate = float(values[np.argmin(distances)])
        else:
            weights = 1.0 / distances
            estimate = float(np.average(values, weights=weights))
        series.append((int(year), estimate, float(nearest["distance_m"].min())))
    return sorted(series)


def run_spatial_regression(sample: pd.DataFrame, elements: list[str], radius_m: float) -> pd.DataFrame:
    """Avalia a tendência temporal da localização, sem parear IDs entre ciclos."""
    historical = sample[sample["year"].astype(str) != CURRENT_YEAR].copy()
    current = sample[sample["year"].astype(str) == CURRENT_YEAR].copy()
    predictions = []
    for element in elements:
        if is_ignored_element(element):
            continue
        for index, target in current.iterrows():
            value = pd.to_numeric(pd.Series([target[element]]), errors="coerce").iloc[0]
            if pd.isna(value):
                continue
            series = spatial_temporal_reference(target, historical, element, radius_m)
            if len(series) < MIN_REGRESSION_CYCLES:
                predictions.append({
                    "sample_index": index, "element": element, "reg_status": "dados insuficientes",
                    "reg_prediction": np.nan, "reg_residual": np.nan, "reg_score": np.nan,
                    "reg_cycles": len(series), "reg_rule": f"Regressão não calculada: {len(series)} ciclos espaciais; mínimo={MIN_REGRESSION_CYCLES}; raio={radius_m:g} m",
                })
                continue
            years = np.array([item[0] for item in series], dtype=float)
            values = np.array([item[1] for item in series], dtype=float)
            model = TheilSenRegressor(random_state=42, fit_intercept=True) if TheilSenRegressor is not None else None
            if model is not None:
                model.fit(years.reshape(-1, 1), values)
                predicted = float(model.predict(np.array([[int(CURRENT_YEAR)]]))[0])
                fitted = model.predict(years.reshape(-1, 1))
            else:
                slope, intercept = np.polyfit(years, values, 1)
                predicted = float(slope * int(CURRENT_YEAR) + intercept)
                fitted = slope * years + intercept
            residuals = values - fitted
            mad = float(np.median(np.abs(residuals - np.median(residuals))))
            scale = max(3.0 * mad, float(np.std(residuals)), 1e-6)
            residual = float(value - predicted)
            score = abs(residual) / scale
            # Regra robusta original: só marca quando o resíduo supera três
            # escalas robustas da série temporal espacial.
            is_anomaly = score > 3.0
            predictions.append({
                "sample_index": index, "element": element,
                "reg_status": "anomalia temporal" if is_anomaly else "normal",
                "reg_prediction": predicted, "reg_residual": residual, "reg_score": score,
                "reg_cycles": len(series),
                "reg_rule": f"Theil-Sen espacial; {len(series)} ciclos; resíduo robusto > 3×escala; raio={radius_m:g} m",
            })
    return pd.DataFrame(predictions)


def apply_ph_smp_validation(result: pd.DataFrame) -> pd.DataFrame:
    """Marca o pH quando ele está acima do SMP do mesmo ponto."""
    # Alguns Excel podem carregar cabeçalhos repetidos. Sem esta proteção,
    # frame["value"] retorna uma DataFrame e a comparação abaixo deixa de ser
    # escalar, gerando o ValueError de ambiguidade booleana.
    frame = result.loc[:, ~result.columns.duplicated()].copy()
    frame["ph_smp_status"] = ""
    if frame.empty or "element" not in frame.columns:
        return frame
    if "sample_index" not in frame.columns:
        frame["sample_index"] = frame.get("source_row", pd.Series(frame.index, index=frame.index))
    normalized = frame["element"].astype(str).str.strip().str.upper()
    ph_rows = frame[normalized == "PH"]
    smp_rows = frame[normalized == "SMP"]
    if ph_rows.empty or smp_rows.empty:
        return frame
    smp_by_point = pd.to_numeric(smp_rows.set_index("sample_index")["value"], errors="coerce")
    if isinstance(smp_by_point, pd.DataFrame):
        smp_by_point = smp_by_point.iloc[:, 0]
    smp_by_point = smp_by_point.groupby(level=0).first()
    for index, row in ph_rows.iterrows():
        smp_value = smp_by_point.get(row["sample_index"], np.nan)
        ph_value = pd.to_numeric(pd.Series([row["value"]]), errors="coerce").iloc[0]
        if pd.notna(ph_value) and pd.notna(smp_value) and ph_value > smp_value:
            frame.at[index, "ph_smp_status"] = "anomalia: pH > SMP"
    return frame


def display_classification(row: pd.Series) -> str:
    """Reduz os sinais internos a outlier, anomalia ou normal para a tabela."""
    if row.get("ph_smp_status") == "anomalia: pH > SMP":
        return "anomalia (pH > SMP)"
    if row.get("final_status") == "outlier confirmado":
        return "outlier"
    if row.get("final_status") in {"atenção", "anomalia temporal", "anomalia espacial"}:
        return "anomalia"
    return "normal"


def select_evidence(result: pd.DataFrame) -> pd.DataFrame:
    """Seleciona outliers confirmados e validações determinísticas pH/SMP.

    A validação pH/SMP não substitui LOF ou regressão e não entra no
    treinamento. Ela é uma regra química independente e precisa continuar
    visível/exportável mesmo quando os dois critérios estatísticos não
    confirmam o ponto.
    """
    if result.empty:
        return result.copy()
    confirmed = result.get("final_status", pd.Series(index=result.index, dtype=object)).eq("outlier confirmado")
    ph_smp = result.get("ph_smp_status", pd.Series(index=result.index, dtype=object)).eq("anomalia: pH > SMP")
    return result[confirmed | ph_smp].copy()


def locate_ibge(rows: pd.DataFrame) -> pd.DataFrame:
    result = rows.copy()
    result["municipality"] = "Não localizado"
    result["state"] = "Não localizado"
    boundaries = load_ibge_municipalities()
    if not boundaries:
        return result
    for index, row in result.iterrows():
        point = Point(float(row.longitude), float(row.latitude))
        for boundary in boundaries:
            minx, miny, maxx, maxy = boundary.bounds
            if minx <= point.x <= maxx and miny <= point.y <= maxy and (boundary.geometry.contains(point) or boundary.geometry.touches(point)):
                result.at[index, "municipality"] = boundary.municipality
                result.at[index, "state"] = boundary.state
                break
    return result


@st.cache_data(show_spinner=False)
def export_individual(original: bytes, evidence: pd.DataFrame, radius_m: float = MAX_DISTANCE_M) -> bytes:
    workbook = load_workbook(io.BytesIO(original))
    sheet = workbook[CURRENT_YEAR]
    # Snapshot explícito: as colunas excluídas permanecem com os valores
    # originais, mesmo quando o restante da planilha recebe destaques.
    original_workbook = load_workbook(io.BytesIO(original), data_only=False)
    original_sheet = original_workbook[CURRENT_YEAR]
    excluded_columns = list(range(19, 24)) + list(range(29, original_sheet.max_column + 1))  # S–W e AC+
    original_values = {
        (row, column): original_sheet.cell(row, column).value
        for row in range(1, original_sheet.max_row + 1)
        for column in excluded_columns
    }
    headers = {str(cell.value or "").strip().lower(): cell.column for cell in sheet[1]}
    fill = PatternFill("solid", fgColor="F6B26B")
    ph_smp_fill = PatternFill("solid", fgColor="FFF2CC")
    sheet.cell(1, 1).comment = Comment(
        f"Configuração Outlieria: ano analisado {CURRENT_YEAR}; raio regional usado por ponto {radius_m:g} m; "
        f"LOF por lavoura, elemento e profundidade; mínimo de {MIN_LOF_NEIGHBORS} vizinhos; "
        f"limiar LOF no percentil {LOF_THRESHOLD_PERCENTILE:g} dos scores históricos; regressão temporal espacial Theil-Sen com mínimo de {MIN_REGRESSION_CYCLES} ciclos.",
        "Outlieria",
    )
    evidence = evidence[~evidence.element.map(is_ignored_element)].copy()
    # A profundidade 10–20 cm continua disponível na Visão 2026, mas não é
    # destacada nem considerada na planilha exportada, conforme solicitado.
    if "depth" in evidence.columns:
        depth_key = evidence["depth"].astype(str).str.strip().str.replace("–", "-", regex=False).str.lower()
        evidence = evidence[~depth_key.str.startswith("10-20")].copy()
    for _, item in evidence.iterrows():
        column = headers.get(str(item.element).strip().lower())
        if not column:
            continue
        cell = sheet.cell(int(item.source_row), column)
        is_ph_smp_anomaly = str(item.get("ph_smp_status", "")) == "anomalia: pH > SMP"
        cell.fill = ph_smp_fill if is_ph_smp_anomaly else fill
        cell.font = Font(name=cell.font.name, sz=cell.font.sz, bold=True, color="7A2E00")
        final_status = str(item.get("final_status", "outlier estatístico")).upper()
        ml_note = ""
        if str(item.get("ml_status", "")) == "anomalia espacial":
            ml_note = f" LOF espacial: score {float(item.get('ml_score', 0)):.3f}; {int(item.get('ml_neighbor_count', 0))} vizinhos; k={int(item.get('ml_k_effective', MIN_LOF_NEIGHBORS))}; raio={float(item.get('ml_radius_m', radius_m)):g} m."
        validation_note = " Validação pH/SMP: pH acima do SMP; célula destacada em amarelo."
        if not is_ph_smp_anomaly:
            validation_note = ""
        reg_note = ""
        if str(item.get("reg_status", "")) == "anomalia temporal":
            reg_note = f" Regressão temporal espacial: previsto {float(item.get('reg_prediction', 0)):.3f}; resíduo {float(item.get('reg_residual', 0)):.3f}; score {float(item.get('reg_score', 0)):.3f}."
        cell.comment = Comment(f"{final_status} em 2026 comparado pela coordenada ao histórico espacial de {radius_m:g} m. Valor {item.value}.{ml_note}{reg_note}{validation_note}", "Outlieria")
    for (row, column), value in original_values.items():
        sheet.cell(row, column).value = value
    # As colunas preservadas podem conter fórmulas; solicite ao Excel/LibreOffice
    # o recálculo para que os valores calculados não apareçam vazios ao abrir.
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"
    output = io.BytesIO(); workbook.save(output); return output.getvalue()


def export_consolidated(all_results: list[tuple[pd.DataFrame, list[PolygonRecord]]]) -> bytes:
    rows = []
    coverage = {}
    for analysis_index, (result, polygons) in enumerate(all_results):
        required = {"element", "depth", "value", "municipality", "state", "source_row", "point_id"}
        if result is None or not required.issubset(result.columns):
            continue
        result = result[(~result.element.map(is_ignored_element)) & (result.depth == "0-10 cm") & result.value.notna()].copy()
        for _, row in result.iterrows():
            area = row.municipality if row.municipality != "Não localizado" else "Não localizado"
            state = row.state if row.state != "Não localizado" else "Não localizado"
            rows.append({"municipality": area, "state": state, "element": row.element, "value": row.value})
            key = (state, analysis_index)
            coverage.setdefault(key, {"points": set(), "area": sum(p.area_ha for p in polygons)})["points"].add((row.source_row, row.point_id))
    frame = pd.DataFrame(rows)
    if frame.empty:
        frame = pd.DataFrame(columns=["municipality", "state", "element", "value"])
    def make_matrix(area):
        group = frame.groupby([area, "element"], dropna=False).value.mean().unstack("element")
        group.index.name = "Localidade"; return group.reset_index()
    municipality = make_matrix("municipality")
    state = make_matrix("state")
    coverage_rows = []
    for state_name in sorted({key[0] for key in coverage}):
        items = [value for (name, _), value in coverage.items() if name == state_name]
        coverage_rows.append({"Estado": state_name, "Pontos totais": len(set().union(*(x["points"] for x in items))), "Área total amostrada (ha)": round(sum(x["area"] for x in items), 2)})
    coverage_frame = pd.DataFrame(coverage_rows)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        municipality.to_excel(writer, index=False, sheet_name="Médias Município")
        state.to_excel(writer, index=False, sheet_name="Médias Estado", startrow=0, startcol=0)
        coverage_frame.to_excel(writer, index=False, sheet_name="Médias Estado", startrow=0, startcol=len(state.columns) + 2)
    formatted = load_workbook(io.BytesIO(output.getvalue()))
    for sheet_name in ["Médias Município", "Médias Estado"]:
        sheet = formatted[sheet_name]
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "0.00"
    state_sheet = formatted["Médias Estado"]
    coverage_start = len(state.columns) + 3
    for row in range(2, state_sheet.max_row + 1):
        state_sheet.cell(row, coverage_start + 1).number_format = "0"
        state_sheet.cell(row, coverage_start + 2).number_format = "0.00"
    final_output = io.BytesIO(); formatted.save(final_output)
    return final_output.getvalue()


@st.cache_data(show_spinner=False)
def frame_for_display(result: pd.DataFrame) -> pd.DataFrame:
    columns = ["point_id", "crop", "depth", "element", "value", "display_status"]
    frame = result[[c for c in columns if c in result.columns]].copy()
    if "value" in frame.columns:
        frame["value"] = pd.to_numeric(frame["value"], errors="coerce").round(2)
    if "ml_score" in frame.columns:
        frame["ml_score"] = pd.to_numeric(frame["ml_score"], errors="coerce").round(3)
    for column in ("ml_score", "reg_prediction", "reg_residual", "reg_score", "ml_radius_m"):
        if column in frame.columns:
            frame[column] = pd.to_numeric(frame[column], errors="coerce").round(3)
    if "element" in frame.columns:
        frame["element"] = frame["element"].map(element_label)
    return frame.rename(columns={"point_id": "ponto", "crop": "lavoura", "depth": "profundidade", "element": "elemento", "value": "valor", "display_status": "classificação"})


def ordered_unique(values: pd.Series) -> list[str]:
    """Retorna os valores na ordem em que aparecem na planilha, sem ordenar."""
    return list(dict.fromkeys(str(value) for value in values.dropna().tolist()))


def element_label(value: object) -> str:
    return ELEMENT_LABELS.get(str(value), str(value))


def is_ignored_element(value: object) -> bool:
    normalized = unicodedata.normalize("NFKD", str(value).strip()).encode("ascii", "ignore").decode().upper()
    return normalized in IGNORED_ELEMENTS or normalized.startswith("SATURAC")


def analysis_identity_key(client_name: object, farm_name: object, excel_name: object, shp_name: object) -> str:
    """Chave estável para reconhecer a mesma fonte mesmo quando seu conteúdo mudou."""
    values = [client_name, farm_name, CURRENT_YEAR, excel_name, shp_name]
    return "|".join(str(value or "").strip().casefold() for value in values)


def normalise_result_frame(result: pd.DataFrame, radius_m: float = MAX_DISTANCE_M) -> pd.DataFrame:
    """Completa campos de versões antigas para que o filtro 2026 seja uniforme."""
    frame = result.copy()
    core_defaults = {
        "client": "", "farm": "", "year": CURRENT_YEAR, "sample_index": np.nan,
        "source_row": np.nan, "point_id": "", "crop": "", "depth": "",
        "element": "", "value": np.nan, "municipality": "Não localizado",
        "state": "Não localizado", "longitude": np.nan, "latitude": np.nan,
    }
    for column, default in core_defaults.items():
        if column not in frame.columns:
            frame[column] = default
    if frame.empty:
        return frame
    defaults = {
        "ml_status": "dados insuficientes",
        "ml_score": np.nan,
        "ml_neighbor_count": np.nan,
        "ml_k_effective": np.nan,
        "ml_radius_m": radius_m,
        "ml_rule": "LOF não disponível nesta análise persistida",
        "reg_status": "dados insuficientes",
        "reg_prediction": np.nan,
        "reg_residual": np.nan,
        "reg_score": np.nan,
        "reg_cycles": 0,
        "reg_rule": "Regressão não disponível nesta análise persistida",
    }
    for column, default in defaults.items():
        if column not in frame.columns:
            frame[column] = default
        elif column in {"ml_status", "ml_rule", "reg_status", "reg_rule"}:
            frame[column] = frame[column].fillna(default)
    if "final_status" not in frame.columns:
        frame["final_status"] = frame.apply(combine_detection, axis=1)
    else:
        frame["final_status"] = frame["final_status"].fillna(frame.apply(combine_detection, axis=1))
    frame = apply_ph_smp_validation(frame)
    frame["display_status"] = frame.apply(display_classification, axis=1)
    return frame


def process_source_bytes(original_bytes: bytes, shp_bytes: bytes, radius_m: float) -> dict[str, object]:
    """Processa uma fonte completa sem depender de widgets Streamlit."""
    polygons = read_polygons(shp_bytes)
    shp_producer, shp_farm = shp_source_names(polygons)
    sample, elements, _ = parse_workbook(original_bytes, polygons, shp_producer, shp_farm)
    radius_m = validate_radius(radius_m)
    if not (sample["year"].astype(str) == CURRENT_YEAR).any():
        raise ValueError("A planilha não contém registros do ciclo 2026.")
    if not elements:
        raise ValueError("Nenhum elemento numérico elegível foi encontrado para a análise.")
    result, _, _ = classify(sample, elements, radius_m)
    result = locate_ibge(result)
    lof_result = run_spatial_lof(sample, elements, radius_m)
    if lof_result.empty:
        result["ml_status"] = "dados insuficientes"
        result["ml_score"] = np.nan
        result["ml_neighbor_count"] = np.nan
        result["ml_k_effective"] = np.nan
        result["ml_radius_m"] = radius_m
        result["ml_rule"] = f"LOF não calculado; mínimo={MIN_LOF_NEIGHBORS}; raio={radius_m:g} m"
        result["ml_mean_distance_m"] = np.nan
        result["ml_regional_median"] = np.nan
        result["ml_reference_scope"] = "distribuição histórica espacial da mesma região; sem pareamento por ID"
        result["ml_historical_percentile"] = np.nan
        result["ml_historical_distribution_shift"] = np.nan
        result["ml_current_cohort_shift"] = np.nan
    else:
        lof_columns = ["sample_index", "element", "ml_score", "ml_status", "ml_neighbor_count", "ml_k_effective", "ml_radius_m", "ml_rule", "ml_mean_distance_m", "ml_regional_median", "ml_reference_scope", "ml_historical_percentile", "ml_historical_distribution_shift", "ml_current_cohort_shift"]
        result = result.merge(lof_result[lof_columns], on=["sample_index", "element"], how="left")
        result["ml_status"] = result["ml_status"].fillna("dados insuficientes")
    result["ml_radius_m"] = result["ml_radius_m"].fillna(radius_m)
    result["ml_rule"] = result["ml_rule"].fillna(f"LOF não calculado; mínimo={MIN_LOF_NEIGHBORS}; raio={radius_m:g} m")
    try:
        regression_result = run_spatial_regression(sample, elements, radius_m)
    except Exception:
        regression_result = pd.DataFrame()
    if regression_result.empty:
        result["reg_status"] = "dados insuficientes"
        result["reg_prediction"] = np.nan
        result["reg_residual"] = np.nan
        result["reg_score"] = np.nan
        result["reg_cycles"] = 0
        result["reg_rule"] = f"Regressão não calculada; mínimo={MIN_REGRESSION_CYCLES} ciclos; raio={radius_m:g} m"
    else:
        result = result.merge(regression_result, on=["sample_index", "element"], how="left")
        result["reg_status"] = result["reg_status"].fillna("dados insuficientes")
    result["reg_rule"] = result["reg_rule"].fillna(f"Regressão não calculada; mínimo={MIN_REGRESSION_CYCLES} ciclos; raio={radius_m:g} m")
    result["final_status"] = result.apply(combine_detection, axis=1)
    result = apply_ph_smp_validation(result)
    result["display_status"] = result.apply(display_classification, axis=1)
    client_name = shp_producer or (str(result.client.dropna().iloc[0]) if not result.empty and result.client.notna().any() else "Sem cliente")
    farm_name = shp_farm or (str(result.farm.dropna().iloc[0]) if not result.empty and result.farm.notna().any() else "Sem fazenda")
    result["client"] = client_name
    result["farm"] = farm_name
    shp_fields = {index: polygon.field_name.strip() for index, polygon in enumerate(polygons) if polygon.field_name.strip()}
    if "polygon_index" in result.columns:
        result["graph_crop"] = result["polygon_index"].map(shp_fields).fillna("Campo não informado")
    evidence = select_evidence(result)
    attention = result[result.final_status == "atenção"].copy()
    return {"polygons": polygons, "result": result, "evidence": evidence, "attention": attention, "client_name": client_name, "farm_name": farm_name, "area_ha": sum(p.area_ha for p in polygons), "elements": elements, "radius_m": radius_m}


def build_persisted_analysis(payload: dict[str, object]) -> dict[str, object]:
    raw_result = pd.DataFrame(payload.get("result_rows", []))
    result = normalise_result_frame(raw_result, float(payload.get("radius_m") or MAX_DISTANCE_M))
    client_name = str(payload.get("client_name") or (result.client.dropna().iloc[0] if not result.empty and result.client.notna().any() else "—"))
    farm_name = str(payload.get("farm_name") or (result.farm.dropna().iloc[0] if not result.empty and result.farm.notna().any() else "—"))
    if result.empty:
        # Resumos antigos não tinham analysis_points/analysis_results. Ainda
        # assim, eles devem aparecer como opções de cliente e fazenda.
        result = normalise_result_frame(pd.DataFrame([{
            "client": client_name,
            "farm": farm_name,
            "year": CURRENT_YEAR,
            "element": "",
            "value": np.nan,
            "persisted_summary_only": True,
        }]), float(payload.get("radius_m") or MAX_DISTANCE_M))
    area_ha = float(payload.get("area_ha") or 0)
    polygons = [PolygonRecord(None, area_ha)] if area_ha else []
    name = str(payload.get("name") or f"{farm_name}_{CURRENT_YEAR}.xlsx")
    shp_name = str(payload.get("shp_name") or f"{farm_name}_limite.shp")
    return {
        "analysis_id": payload.get("analysis_id"),
        "source_hash": payload.get("source_hash", ""),
        "analysis_key": analysis_identity_key(client_name, farm_name, name, shp_name),
        "name": name,
        "shp_name": shp_name,
        "client_name": client_name,
        "farm_name": farm_name,
        "original": b"",
        "shp_bytes": b"",
        "radius_m": float(payload.get("radius_m") or MAX_DISTANCE_M),
        "result": result,
        "evidence": select_evidence(result),
        "attention": result[result.final_status == "atenção"].copy() if "final_status" in result.columns else result.iloc[0:0].copy(),
        "polygons": polygons,
        "area_ha": area_ha,
        "elements": [value for value in ordered_unique(result.element) if str(value).strip()] if not result.empty and "element" in result.columns else [],
        "file_urls": payload.get("file_urls", {}),
        "updated_at": payload.get("updated_at"),
        "point_count": payload.get("point_count", 0),
        "persisted": True,
    }


def build_vision_snapshot(analyses: list[dict[str, object]]) -> pd.DataFrame | None:
    frames = [analysis["result"] for analysis in analyses if isinstance(analysis.get("result"), pd.DataFrame) and not analysis["result"].empty]
    return pd.concat(frames, ignore_index=True).copy() if frames else None


def upsert_session_analysis(record: dict[str, object]) -> bool:
    """Substitui a fonte existente por identidade/hash; não duplica reenvios."""
    for index, current in enumerate(st.session_state.analyses):
        same_hash = record.get("source_hash") and current.get("source_hash") == record.get("source_hash")
        same_identity = record.get("analysis_key") and current.get("analysis_key") == record.get("analysis_key")
        if same_hash or same_identity:
            st.session_state.analyses[index] = record
            return True
    st.session_state.analyses.append(record)
    return False


def rehydrate_persisted_analysis(payload: dict[str, object]) -> dict[str, object]:
    """Recupera fontes antigas do Drive e preenche seus pontos/resultados."""
    result_rows = payload.get("result_rows") or []
    # Resultados gravados antes da coluna graph_crop não carregam o CAMPO do
    # SHP. Nessa situação, baixa-se a fonte original e reprocessa-se a análise
    # para reconstruir o nome da lavoura a partir do limite correto.
    has_graph_crop = any(str(row.get("graph_crop") or "").strip() for row in result_rows if isinstance(row, dict))
    if result_rows and has_graph_crop:
        return build_persisted_analysis(payload)
    record = build_persisted_analysis(payload)
    if not cloud_storage_configured() or not payload.get("analysis_id"):
        return record
    source = download_persisted_source(str(payload["analysis_id"]))
    processed = process_source_bytes(source["excel_original"], source["shp"], float(payload.get("radius_m") or MAX_DISTANCE_M))
    reanalysis = export_individual(source["excel_original"], processed["evidence"], processed["radius_m"])
    metadata = persist_analysis(
        client_name=processed["client_name"], farm_name=processed["farm_name"], year=int(CURRENT_YEAR),
        radius_m=processed["radius_m"], original=source["excel_original"], shp_bytes=source["shp"],
        reanalysis=reanalysis, original_name=source.get("excel_original_name") or payload.get("name") or "analise.xlsx",
        shp_name=source.get("shp_name") or payload.get("shp_name") or "limite.shp", area_ha=processed["area_ha"],
        point_count=len(processed["result"]), outlier_count=len(processed["evidence"]),
        attention_count=len(processed["attention"]), result=processed["result"], lof_min_neighbors=MIN_LOF_NEIGHBORS,
    ) or {}
    name = str(source.get("excel_original_name") or payload.get("name") or "analise.xlsx")
    shp_name = str(source.get("shp_name") or payload.get("shp_name") or "limite.shp")
    return {
        "analysis_id": metadata.get("analysis_id") or payload.get("analysis_id"),
        "source_hash": metadata.get("source_hash") or payload.get("source_hash") or sha256(source["excel_original"] + source["shp"]),
        "analysis_key": analysis_identity_key(processed["client_name"], processed["farm_name"], name, shp_name),
        "name": name, "shp_name": shp_name, "client_name": processed["client_name"], "farm_name": processed["farm_name"],
        "original": source["excel_original"], "shp_bytes": source["shp"], "radius_m": processed["radius_m"],
        "result": processed["result"], "evidence": processed["evidence"], "attention": processed["attention"],
        "polygons": processed["polygons"], "area_ha": processed["area_ha"], "elements": processed["elements"],
        "file_urls": metadata.get("file_urls") or payload.get("file_urls", {}),
        "updated_at": metadata.get("updated_at") or payload.get("updated_at"), "point_count": len(processed["result"]), "persisted": True,
    }


def process_uploaded_source_job(original_bytes: bytes, shp_bytes: bytes, radius_m: float, excel_name: str, shp_name: str) -> tuple[dict[str, object], str | None]:
    """Processa e persiste uma fonte em uma tarefa independente da navegação."""
    processed = process_source_bytes(original_bytes, shp_bytes, radius_m)
    reanalysis_bytes = export_individual(original_bytes, processed["evidence"], processed["radius_m"])
    metadata: dict[str, object] = {}
    message = None
    if database_configured():
        try:
            metadata = persist_analysis(
                client_name=processed["client_name"], farm_name=processed["farm_name"], year=int(CURRENT_YEAR),
                radius_m=processed["radius_m"], original=original_bytes, shp_bytes=shp_bytes,
                reanalysis=reanalysis_bytes, original_name=excel_name, shp_name=shp_name,
                area_ha=processed["area_ha"], point_count=len(processed["result"]),
                outlier_count=len(processed["evidence"]), attention_count=len(processed["attention"]),
                result=processed["result"], lof_min_neighbors=MIN_LOF_NEIGHBORS,
            ) or {}
            message = "Análise atualizada no Supabase e no Google Drive." if cloud_storage_configured() else "Análise atualizada no Supabase; configure o Google Drive para armazenar os arquivos."
        except Exception as storage_exc:
            message = f"A análise foi processada nesta sessão, mas não pôde ser sincronizada com a nuvem: {storage_exc}"
    else:
        message = f"A análise foi processada somente nesta sessão. Configure estas Secrets para ativar o Supabase e o Google Drive: {', '.join(missing_settings())}."
    record = {
        "analysis_id": metadata.get("analysis_id"),
        "source_hash": metadata.get("source_hash") or sha256(original_bytes + shp_bytes),
        "analysis_key": analysis_identity_key(processed["client_name"], processed["farm_name"], excel_name, shp_name),
        "name": excel_name, "shp_name": shp_name,
        "client_name": processed["client_name"], "farm_name": processed["farm_name"],
        "original": original_bytes, "shp_bytes": shp_bytes, "radius_m": processed["radius_m"],
        "result": processed["result"], "evidence": processed["evidence"], "attention": processed["attention"],
        "polygons": processed["polygons"], "area_ha": processed["area_ha"], "elements": processed["elements"],
        "file_urls": metadata.get("file_urls", {}),
        "updated_at": metadata.get("updated_at") or datetime.now(ZoneInfo("America/Sao_Paulo")).isoformat(),
        "persisted": bool(metadata.get("analysis_id")),
    }
    return record, message


def render_scatter(result: pd.DataFrame, y_domain: tuple[float, float] | None = None) -> None:
    """Renderiza dispersão em SVG, sem Altair, Plotly ou dependências de gráfico."""
    visible_statuses = {"atenção", "outlier confirmado", "outlier estatístico", "anomalia temporal", "anomalia espacial", "outlier"}
    points = result[result.get("final_status", pd.Series(index=result.index, dtype=object)).isin(visible_statuses)].dropna(subset=["source_row", "value"]).copy()
    if points.empty:
        return
    width, height = 1100, 280
    # Reserva uma coluna própria para a legenda, sem sobreposição com o gráfico.
    left, right, top, bottom = 145, 16, 18, 72
    # O eixo X é organizado em blocos por lavoura, preservando a ordem em que
    # as lavouras aparecem na planilha e delimitando visualmente cada região.
    if "graph_crop" in points.columns:
        crop_order = ordered_unique(points["graph_crop"].fillna("Campo não informado"))
    else:
        crop_order = ["Campo não informado"]
        points["graph_crop"] = "Campo não informado"
    ordered_parts = []
    crop_regions = []
    cursor = 0
    for crop_name in crop_order:
        crop_points = points[points["graph_crop"].astype(str) == str(crop_name)].copy()
        if crop_points.empty:
            continue
        crop_points = crop_points.sort_values("source_row", kind="stable")
        start_x = cursor
        cursor += len(crop_points)
        crop_regions.append((str(crop_name), start_x, cursor - 1))
        ordered_parts.append(crop_points)
        cursor += 1  # espaço visual entre lavouras
    points = pd.concat(ordered_parts, ignore_index=True) if ordered_parts else points.reset_index(drop=True)
    many_crops = len(crop_regions) > 6
    height = 360 if many_crops else 280
    x_values = np.array([index + sum(1 for _, start, end in crop_regions if end < index) for index in range(len(points))], dtype=float)
    y_values = pd.to_numeric(points["value"], errors="coerce").to_numpy(dtype=float)
    raw_x_min, raw_x_max = float(np.nanmin(x_values)), float(np.nanmax(x_values))
    raw_y_min, raw_y_max = y_domain if y_domain else (float(np.nanmin(y_values)), float(np.nanmax(y_values)))
    # Escalas legíveis e consistentes: 5, 10, 50, 100, 500...
    axis_steps = [5 * (10 ** exponent) for exponent in range(-2, 6)]
    def rounded_axis(raw_min, raw_max):
        span = max(raw_max - raw_min, 0.5)
        step = min(axis_steps, key=lambda candidate: abs(candidate - span / 4))
        axis_min = math.floor(raw_min / step) * step
        axis_max = math.ceil(raw_max / step) * step
        if axis_min == axis_max: axis_max = axis_min + step
        return axis_min, axis_max, step
    x_min, x_max = -0.5, max(raw_x_max, 0.5)
    x_step = 1.0
    y_min, y_max, y_step = rounded_axis(raw_y_min, raw_y_max)
    def sx(x): return left + ((x - x_min) / (x_max - x_min)) * (width - left - right)
    def sy(y): return top + ((y_max - y) / (y_max - y_min)) * (height - top - bottom)
    def tick_text(value, step):
        return f"{value:.1f}" if step == 0.5 else f"{value:.0f}"
    colors = {"outlier confirmado": "#c94b35", "outlier estatístico": "#d8793e", "anomalia temporal": "#bd6b37", "anomalia espacial": "#7656a8", "outlier": "#d8793e", "atenção": "#b79b43", "normal": "#487c80"}
    circles = []
    tooltips = []
    for point_index, ((_, row), x, y) in enumerate(zip(points.iterrows(), x_values, y_values)):
        status = str(row.get("final_status", row.get("status", "normal"))); color = colors.get(status, "#487c80")
        px, py = sx(x), sy(y)
        # Mantém o cartão branco acima do marcador, liberando a área de dados
        # para inspeção sem cobrir o próprio ponto.
        tooltip_x, tooltip_y = min(px + 9, width - 158), max(py - 58, 4)
        point_id = html.escape(str(row.get("point_id", row.get("source_row", ""))))
        ml_score = row.get("ml_score")
        score_text = f" · LOF: {float(ml_score):.2f}" if pd.notna(ml_score) else ""
        circles.append(f'<g class="outlieria-point point-{point_index}"><circle cx="{px:.1f}" cy="{py:.1f}" r="4.5" fill="{color}" stroke="#fffdf8" stroke-width="1.5"/></g>')
        tooltips.append(f'<g class="point-tooltip tooltip-{point_index}" style="z-index:100"><rect x="{tooltip_x:.1f}" y="{tooltip_y:.1f}" width="148" height="46" rx="5" fill="#ffffff" fill-opacity="1" stroke="#dce1d1"/><text x="{tooltip_x + 9:.1f}" y="{tooltip_y + 14:.1f}" font-family="IBM Plex Mono, monospace" font-size="9" fill="#173638">Vértice / ID: {point_id}</text><text x="{tooltip_x + 9:.1f}" y="{tooltip_y + 27:.1f}" font-family="IBM Plex Mono, monospace" font-size="9" fill="#5d6d68">Valor: {y:.2f}</text><text x="{tooltip_x + 9:.1f}" y="{tooltip_y + 40:.1f}" font-family="IBM Plex Mono, monospace" font-size="8" fill="#7656a8">{html.escape(status)}{score_text}</text></g>')
    grid_parts = []
    y_ticks = np.arange(y_min, y_max + y_step * 0.01, y_step)
    for y_tick in y_ticks:
        fraction = (y_max - y_tick) / (y_max - y_min)
        gy = top + fraction * (height - top - bottom)
        grid_parts.append(f'<line x1="{left}" x2="{width-right}" y1="{gy:.1f}" y2="{gy:.1f}" stroke="#e5e8df" stroke-width="1"/><text x="{left - 7}" y="{gy + 3:.1f}" text-anchor="end" font-family="IBM Plex Mono, monospace" font-size="9" fill="#7c8880">{tick_text(y_tick, y_step)}</text>')
    region_backgrounds = []
    for region_index, (crop_name, start_x, end_x) in enumerate(crop_regions):
        region_left = start_x - 0.5
        region_right = end_x + 0.5
        if region_index % 2 == 0:
            region_backgrounds.append(f'<rect x="{sx(region_left):.1f}" y="{top}" width="{max(sx(region_right) - sx(region_left), 1):.1f}" height="{height-bottom-top}" fill="#f1f5e8" fill-opacity="0.58"/>')
        if region_index:
            separator_x = sx(region_left)
            grid_parts.append(f'<line x1="{separator_x:.1f}" x2="{separator_x:.1f}" y1="{top}" y2="{height-bottom}" stroke="#cbd4c4" stroke-width="1.2" stroke-dasharray="3 3"/>')
        center_x = sx((start_x + end_x) / 2)
        label = html.escape(crop_name)
        label_y = height - bottom + (50 if many_crops else 17)
        if many_crops:
            grid_parts.append(f'<text x="{center_x:.1f}" y="{label_y}" text-anchor="end" transform="rotate(-45 {center_x:.1f} {label_y})" font-family="IBM Plex Mono, monospace" font-size="9" font-weight="600" fill="#6f7c76">{label}</text>')
        else:
            grid_parts.append(f'<text x="{center_x:.1f}" y="{label_y}" text-anchor="middle" font-family="IBM Plex Mono, monospace" font-size="9" font-weight="600" fill="#6f7c76">{label}</text>')
    grid = "".join(region_backgrounds) + "".join(grid_parts)
    legend = '<g font-family="IBM Plex Mono, monospace" font-size="9"><circle cx="13" cy="92" r="4" fill="#b79b43"/><text x="23" y="95" fill="#6f7c76">atenção</text><circle cx="13" cy="112" r="4" fill="#d8793e"/><text x="23" y="115" fill="#6f7c76">outlier 2026</text><circle cx="13" cy="132" r="4" fill="#7656a8"/><text x="23" y="135" fill="#6f7c76">anomalia</text></g>'
    hover_rules = "".join(f".outlieria-svg:has(.point-{index}:hover) .tooltip-{index}{{opacity:1}}" for index in range(len(circles)))
    st.markdown(f'<div style="background:#fffdf8;border:1px solid #dce1d1;border-radius:10px;padding:4px 8px"><svg class="outlieria-svg" viewBox="0 0 {width} {height}" width="100%" height="{height}" role="img" aria-label="Gráfico de dispersão"><style>.point-tooltip{{opacity:0;pointer-events:none;transition:opacity .12s;isolation:isolate;z-index:100}}{hover_rules}</style>{grid}<line x1="{left}" x2="{left}" y1="{top}" y2="{height-bottom}" stroke="#aeb8ad"/><line x1="{left}" x2="{width-right}" y1="{height-bottom}" y2="{height-bottom}" stroke="#aeb8ad"/>{"".join(circles)}<g class="outlieria-tooltips">{"".join(tooltips)}</g>{legend}</svg></div>', unsafe_allow_html=True)


def render_spatial_map(result: pd.DataFrame) -> None:
    """Renderiza a posição geográfica dos pontos em escala aproximada de metros.

    A longitude é convertida usando a latitude média da seleção e a latitude
    usa o comprimento médio de um grau. A mesma escala é aplicada aos dois
    eixos para preservar a proporção espacial entre os pontos.
    """
    visible_statuses = {"atenção", "outlier confirmado", "outlier estatístico", "anomalia temporal", "anomalia espacial", "outlier"}
    required = {"longitude", "latitude", "value"}
    if not required.issubset(result.columns):
        return
    points = result[result.get("final_status", pd.Series(index=result.index, dtype=object)).isin(visible_statuses)].copy()
    points["longitude"] = pd.to_numeric(points["longitude"], errors="coerce")
    points["latitude"] = pd.to_numeric(points["latitude"], errors="coerce")
    points["value"] = pd.to_numeric(points["value"], errors="coerce")
    points = points.dropna(subset=["longitude", "latitude", "value"])
    if points.empty:
        return

    width, height = 1100, 400
    left, right, top, bottom = 78, 18, 18, 58
    mean_lat = float(points["latitude"].mean())
    lon0 = float(points["longitude"].mean())
    lat0 = float(points["latitude"].mean())
    meters_per_degree_lat = 110_540.0
    meters_per_degree_lon = 111_320.0 * math.cos(math.radians(mean_lat))
    points["x_m"] = (points["longitude"] - lon0) * meters_per_degree_lon
    points["y_m"] = (points["latitude"] - lat0) * meters_per_degree_lat
    x_raw_min, x_raw_max = float(points["x_m"].min()), float(points["x_m"].max())
    y_raw_min, y_raw_max = float(points["y_m"].min()), float(points["y_m"].max())
    x_span = max(x_raw_max - x_raw_min, 1.0)
    y_span = max(y_raw_max - y_raw_min, 1.0)
    # Margem em metros para que os marcadores não fiquem colados na borda.
    margin = max(x_span, y_span) * 0.08
    x_min, x_max = x_raw_min - margin, x_raw_max + margin
    y_min, y_max = y_raw_min - margin, y_raw_max + margin
    plot_width, plot_height = width - left - right, height - top - bottom
    scale = min(plot_width / (x_max - x_min), plot_height / (y_max - y_min))
    map_width, map_height = (x_max - x_min) * scale, (y_max - y_min) * scale
    offset_x, offset_y = left + (plot_width - map_width) / 2, top + (plot_height - map_height) / 2
    def sx(value): return offset_x + (float(value) - x_min) * scale
    def sy(value): return offset_y + (y_max - float(value)) * scale

    colors = {"outlier confirmado": "#c94b35", "outlier estatístico": "#d8793e", "anomalia temporal": "#bd6b37", "anomalia espacial": "#7656a8", "outlier": "#d8793e", "atenção": "#b79b43"}
    grid_parts = [f'<rect x="{offset_x:.1f}" y="{offset_y:.1f}" width="{map_width:.1f}" height="{map_height:.1f}" fill="#f1f5e8" fill-opacity="0.58" stroke="#aeb8ad"/>']
    for fraction in (0.25, 0.5, 0.75):
        gx = offset_x + map_width * fraction
        gy = offset_y + map_height * fraction
        grid_parts.append(f'<line x1="{gx:.1f}" x2="{gx:.1f}" y1="{offset_y:.1f}" y2="{offset_y + map_height:.1f}" stroke="#e5e8df"/>')
        grid_parts.append(f'<line x1="{offset_x:.1f}" x2="{offset_x + map_width:.1f}" y1="{gy:.1f}" y2="{gy:.1f}" stroke="#e5e8df"/>')
    grid_parts.append(f'<text x="{offset_x + map_width / 2:.1f}" y="{height - 12}" text-anchor="middle" font-family="IBM Plex Mono, monospace" font-size="9" fill="#7c8880">posição leste–oeste (m)</text>')
    axis_label_x = max(38.0, offset_x - 28.0)
    grid_parts.append(f'<text x="{axis_label_x:.1f}" y="{offset_y + map_height / 2:.1f}" text-anchor="middle" transform="rotate(-90 {axis_label_x:.1f} {offset_y + map_height / 2:.1f})" font-family="IBM Plex Mono, monospace" font-size="9" fill="#7c8880">posição norte–sul (m)</text>')

    circles, tooltips = [], []
    for point_index, (_, row) in enumerate(points.iterrows()):
        status = str(row.get("final_status", "normal"))
        color = colors.get(status, "#487c80")
        px, py = sx(row["x_m"]), sy(row["y_m"])
        tooltip_x, tooltip_y = min(px + 9, width - 166), max(py - 58, 4)
        point_id = html.escape(str(row.get("point_id", row.get("source_row", ""))))
        crop_name = html.escape(str(row.get("graph_crop", row.get("crop", ""))))
        circles.append(f'<circle class="spatial-point spatial-point-{point_index}" cx="{px:.1f}" cy="{py:.1f}" r="5" fill="{color}" stroke="#fffdf8" stroke-width="1.5"/>')
        tooltips.append(f'<g class="spatial-tooltip spatial-tooltip-{point_index}"><rect x="{tooltip_x:.1f}" y="{tooltip_y:.1f}" width="158" height="62" rx="5" fill="#ffffff" stroke="#dce1d1"/><text x="{tooltip_x + 9:.1f}" y="{tooltip_y + 14:.1f}" font-family="IBM Plex Mono, monospace" font-size="9" fill="#173638">Vértice / ID: {point_id}</text><text x="{tooltip_x + 9:.1f}" y="{tooltip_y + 27:.1f}" font-family="IBM Plex Mono, monospace" font-size="9" fill="#5d6d68">CAMPO: {crop_name}</text><text x="{tooltip_x + 9:.1f}" y="{tooltip_y + 40:.1f}" font-family="IBM Plex Mono, monospace" font-size="9" fill="#5d6d68">Valor: {float(row["value"]):.2f}</text><text x="{tooltip_x + 9:.1f}" y="{tooltip_y + 53:.1f}" font-family="IBM Plex Mono, monospace" font-size="8" fill="#7656a8">{html.escape(status)}</text></g>')
    hover_rules = "".join(f'.outlieria-spatial-svg:has(.spatial-point-{index}:hover) .spatial-tooltip-{index}{{opacity:1}}' for index in range(len(circles)))
    legend = '<g font-family="IBM Plex Mono, monospace" font-size="9"><circle cx="13" cy="28" r="4" fill="#b79b43"/><text x="23" y="31" fill="#6f7c76">atenção</text><circle cx="13" cy="48" r="4" fill="#d8793e"/><text x="23" y="51" fill="#6f7c76">outlier 2026</text><circle cx="13" cy="68" r="4" fill="#7656a8"/><text x="23" y="71" fill="#6f7c76">anomalia</text></g>'
    st.markdown(f'<div style="background:#fffdf8;border:1px solid #dce1d1;border-radius:10px;padding:4px 8px"><svg class="outlieria-spatial-svg" viewBox="0 0 {width} {height}" width="100%" height="{height}" role="img" aria-label="Mapa espacial dos pontos"><style>.spatial-tooltip{{opacity:0;pointer-events:none;transition:opacity .12s;isolation:isolate;z-index:100}}{hover_rules}</style>{"".join(grid_parts)}{"".join(circles)}<g>{"".join(tooltips)}</g>{legend}</svg></div>', unsafe_allow_html=True)


def format_source_timestamp(value: object) -> str:
    """Formata timestamps no horário local de Brasília para a UI."""
    if not value:
        return datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%d/%m/%Y %H:%M")
    try:
        parsed = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=ZoneInfo("America/Sao_Paulo"))
        return parsed.astimezone(ZoneInfo("America/Sao_Paulo")).strftime("%d/%m/%Y %H:%M")
    except (TypeError, ValueError):
        return str(value)


def inject_style():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
    :root { --ink:#173638; --muted:#7b817a; --lime:#c8f169; --deep:#102d30; --paper:#f7f4eb; --line:#dce1d1; --orange:#c97d38; }
    .stApp { background:var(--paper); color:var(--ink); font-family:'DM Sans',sans-serif; }
    [data-testid='stSidebar'] { background:var(--deep); min-width:250px; }
    [data-testid='stSidebar'] * { color:#f7f4eb; }
    [data-testid='stSidebar'] .stRadio > div { gap:2px; }
    [data-testid='stSidebar'] .stRadio label { border-radius:10px; padding:8px 10px; margin:2px 0; font-weight:600; font-size:12px; }
    [data-testid='stSidebar'] .stRadio label:has(input:checked) { background:var(--lime); color:var(--ink); }
    [data-testid='stSidebar'] .stRadio label > div:first-child { display:none; }
    [data-testid='stSidebar'] .stRadio label:hover { background:#23474a; }
    .brand { display:flex; align-items:center; gap:8px; padding:5px 0 18px; color:#f7f4eb; font-weight:700; font-size:16px; }
    .brand-mark { width:20px; height:20px; border:3px solid var(--lime); border-right-color:transparent; border-radius:50%; position:relative; }
    .brand-mark:after { content:''; position:absolute; width:5px; height:5px; background:var(--lime); border-radius:50%; right:-3px; top:1px; }
    .rail-caption { color:#9aa9a0; font:600 10px 'IBM Plex Mono',monospace; letter-spacing:.12em; text-transform:uppercase; margin:8px 0; }
    .scope { border:1px solid #315355; border-radius:9px; padding:9px; margin-top:14px; } .scope strong{display:block;font-size:16px}.scope span{color:#9aa9a0;font:10px 'IBM Plex Mono',monospace;}
    .topbar { border-bottom:1px solid var(--line); padding:2px 0 12px; margin-bottom:18px; display:flex; justify-content:space-between; align-items:center; }
    .tab-loading { display:inline-flex; align-items:center; gap:7px; color:#9b6a3e; font:600 9px 'IBM Plex Mono',monospace; text-transform:uppercase; letter-spacing:.08em; animation:tab-loading-fade 1.35s ease-out forwards; }
    .tab-loading-dot { width:10px; height:10px; border:2px solid #d7dfbe; border-top-color:#9caf4b; border-radius:50%; animation:tab-loading-spin .7s linear infinite; }
    @keyframes tab-loading-spin { to { transform:rotate(360deg); } }
    @keyframes tab-loading-fade { 0%,65% { opacity:1; } 100% { opacity:0; } }
    .eyebrow,.mono { font:600 10px 'IBM Plex Mono',monospace; text-transform:uppercase; letter-spacing:.10em; color:#b16e35; }
    h1 { font-size:42px !important; line-height:.98 !important; letter-spacing:-.055em; margin:10px 0 !important; color:var(--ink); } h1 em { color:#87978d; font-style:normal; }
    .lead { color:#747a73; max-width:560px; font-size:12px; }
    .rule-panel { background:#f0f8de; border:1px solid #d3e3a7; border-radius:15px; padding:14px; margin:15px 0; }
    .rule-title { display:flex; justify-content:space-between; font-weight:700; font-size:16px; margin-bottom:10px; }
    .rule-grid { display:grid; grid-template-columns:repeat(4,1fr); gap:10px; }
    .rule-item { background:#f8faef; border:1px solid #e3e9d4; border-radius:10px; padding:10px; min-height:78px; }
    .rule-num { color:#9aa85f; font:600 11px 'IBM Plex Mono',monospace; } .rule-item strong{display:block;margin:5px 0 3px;font-size:12px}.rule-item p{margin:0;color:#7d8479;font-size:10px;line-height:1.25}
    .stats { display:grid; grid-template-columns:repeat(4,1fr); gap:9px; margin:13px 0; }
    .stat { border:1px solid var(--line); border-radius:12px; padding:11px; min-height:82px; background:#fffdf7; }.stat.blue{background:#e7f3f7}.stat.lime{background:#e8f8bd}.stat.orange{background:#fff0df}.stat.sand{background:#f2eee3}
    .stat-label { font:600 10px 'IBM Plex Mono',monospace; color:#77827a; text-transform:uppercase; }.stat-value{font:700 23px 'IBM Plex Mono',monospace;margin:8px 0 1px}.stat-note{font-size:10px;color:#7c8179}
    .section-head { display:flex; justify-content:space-between; align-items:end; margin:20px 0 7px; }.section-head h2{font-size:20px;margin:4px 0}.section-kicker{font:600 10px 'IBM Plex Mono',monospace;color:#9b6a3e;text-transform:uppercase;letter-spacing:.08em}
    .file-strip { display:flex; gap:9px; align-items:center; background:#fffdf8; border:1px solid var(--line); border-radius:11px; padding:10px; margin:13px 0; }.file-strip strong{font-size:12px}.file-strip span{display:block;color:#8b8e85;font:10px 'IBM Plex Mono',monospace;margin-top:4px}
    .badge { background:#e8f6c5; color:#55701b; border-radius:30px; padding:6px 10px; font:600 10px 'IBM Plex Mono',monospace; margin-left:auto; }
    .chart-card { background:#fffdf8; border:1px solid var(--line); border-radius:13px; padding:10px; }.chart-title{font-weight:700;font-size:15px}.chart-sub{font:10px 'IBM Plex Mono',monospace;color:#8a8e87;text-transform:uppercase}
    .source-card { background:#fffdf8; border:1px solid var(--line); border-radius:13px; padding:13px; margin:7px 0; }
    .source-card h3{margin:0 0 5px}.source-card p{color:#7c8179;font-size:12px}.source-count{color:#849069;font:600 10px 'IBM Plex Mono',monospace;text-transform:uppercase}
    [data-testid='stAppViewContainer'] .main .block-container { max-width:1180px; padding-top:1rem; padding-bottom:1.5rem; } [data-testid='stButton'] button, [data-testid='stDownloadButton'] button { min-height:32px !important; padding:6px 10px !important; border-radius:7px !important; font-size:11px !important; font-weight:600 !important; } [data-testid='stFileUploader'] section { padding:2px 5px !important; min-height:46px !important; height:46px !important; border-radius:7px !important; } [data-testid='stFileUploader'] section > div { min-height:40px !important; padding:0 !important; } [data-testid='stFileUploader'] button { padding:3px 7px !important; font-size:10px !important; } [data-testid='stFileUploader'] small { font-size:9px !important; } [data-testid='stFileUploader'] label { font-size:10px !important; } [data-testid='stFileUploaderDropzoneInstructions'] { font-size:9px !important; line-height:1.1 !important; white-space:nowrap !important; overflow:hidden !important; text-overflow:ellipsis !important; } [data-testid='stNumberInput'] > div { min-height:46px !important; height:46px !important; } [data-testid='stNumberInput'] input { height:40px !important; } [data-testid='stSelectbox'] > div > div { min-height:32px !important; height:32px !important; border-radius:7px !important; background:#fffdf8 !important; border-color:var(--line) !important; } [data-testid='stSelectbox'] label, .stRadio label { font-size:10px !important; } [data-testid='stVerticalBlock'] { gap:0.45rem; } @media (max-width:900px){.rule-grid,.stats{grid-template-columns:repeat(2,1fr)} h1{font-size:34px!important}} @media (max-width:600px){.rule-grid,.stats{grid-template-columns:1fr}}
    </style>
    """, unsafe_allow_html=True)


def html_stat(label, value, note, tone=''):
    return f"<div class='stat {tone}'><div class='stat-label'>{label}</div><div class='stat-value'>{value}</div><div class='stat-note'>{note}</div></div>"


def main():
    st.set_page_config(page_title="Outlieria", page_icon="◌", layout="wide", initial_sidebar_state="expanded")
    inject_style()
    if "analyses" not in st.session_state: st.session_state.analyses = []
    if "persisted_analyses_loaded" not in st.session_state:
        st.session_state.persisted_analyses_loaded = False
    if "persisted_rehydration_errors" not in st.session_state:
        st.session_state.persisted_rehydration_errors = []
    if not st.session_state.persisted_analyses_loaded:
        if database_configured():
            try:
                persisted = []
                for payload in load_persisted_analyses(int(CURRENT_YEAR)):
                    try:
                        persisted.append(rehydrate_persisted_analysis(payload))
                    except Exception as rehydrate_exc:
                        persisted.append(build_persisted_analysis(payload))
                        st.session_state.persisted_rehydration_errors.append(
                            f"{payload.get('client_name') or 'análise'}: {rehydrate_exc}"
                        )
                st.session_state.analyses = persisted
            except Exception as storage_exc:
                st.warning(f"Não foi possível carregar as análises persistidas: {storage_exc}")
        st.session_state.persisted_analyses_loaded = True
    if st.session_state.persisted_rehydration_errors:
        st.info("Algumas análises antigas permaneceram como resumo porque os arquivos originais não puderam ser recuperados do Google Drive nesta sessão.")
    # O Streamlit reexecuta o script ao trocar a navegação. Mantemos um
    # snapshot explícito da visão para que essa reexecução não apague os dados
    # exibidos; ele só é substituído quando uma nova análise é processada.
    if "vision_2026_result" not in st.session_state:
        st.session_state.vision_2026_result = build_vision_snapshot(st.session_state.analyses)
    if "vision_2026_revision" not in st.session_state:
        st.session_state.vision_2026_revision = 0
    active_radius = float(st.session_state.analyses[-1].get("radius_m", MAX_DISTANCE_M)) if st.session_state.analyses else MAX_DISTANCE_M
    with st.sidebar:
        st.markdown("<div class='brand'><span class='brand-mark'></span><span>outlieria</span></div>", unsafe_allow_html=True)
        st.markdown("<div class='rail-caption'>análise agrícola</div>", unsafe_allow_html=True)
        page = st.radio("Navegação", ["Visão 2026", "Fontes"], label_visibility="collapsed")
        st.markdown(f"<div class='rail-caption' style='margin-top:30px'>recorte ativo</div><div class='scope'><strong>2026</strong><span>último ano · {active_radius:g} m</span></div>", unsafe_allow_html=True)
    st.markdown("<div class='topbar'><div class='mono' style='color:#737d76'>ANÁLISE AGRÍCOLA &nbsp;/&nbsp; <b>PAINEL 2026</b></div><span class='tab-loading'><span class='tab-loading-dot'></span>carregando</span><span class='badge'>● ABA 2026</span></div>", unsafe_allow_html=True)
    if page == "Visão 2026":
        st.markdown("<style>.st-key-source_workspace,.st-key-source-workspace{display:none !important}</style>", unsafe_allow_html=True)
    with st.container(key="source_workspace"):
        st.markdown("<div class='eyebrow'>— &nbsp; FONTES DE DADOS</div><h1>Cadastre as fontes<br><em>uma única vez.</em></h1><p class='lead'>Envie a planilha agrícola e o limite SHP. O processamento começa somente depois de clicar em <b>Pronto</b>.</p>", unsafe_allow_html=True)
        c1, c2, c3 = st.columns([1.25, 1.25, 1.25])
        with c1: excel = st.file_uploader("Planilha Excel", type=["xlsx", "xls"], key="excel")
        with c2: shp = st.file_uploader("Limite SHP ou ZIP", type=["shp", "zip"], key="shp")
        if excel:
            st.session_state.uploaded_excel_bytes = excel.getvalue()
            st.session_state.uploaded_excel_name = excel.name
        if shp:
            st.session_state.uploaded_shp_bytes = shp.getvalue()
            st.session_state.uploaded_shp_name = shp.name
        excel_bytes = st.session_state.get("uploaded_excel_bytes", b"")
        shp_bytes_cached = st.session_state.get("uploaded_shp_bytes", b"")
        excel_name = st.session_state.get("uploaded_excel_name", "analise.xlsx")
        shp_name_cached = st.session_state.get("uploaded_shp_name", "limite.shp")
        with c3:
            radius_m = st.number_input("Raio regional (m)", min_value=1.0, max_value=100000.0, value=active_radius, step=10.0, format="%.0f", help="Distância usada para selecionar o histórico regional de cada ponto. Utilize sempre o raio do maior grid para a comparação.", key="radius_m")
        if not LOF_AVAILABLE:
            st.caption("LOF indisponível até a instalação do scikit-learn pelo requirements.txt.")
        else:
            st.caption(f"LOF espacial ativo automaticamente: mínimo de {MIN_LOF_NEIGHBORS} vizinhos, sem misturar lavoura ou profundidade.")
        st.markdown("<p class='lead' style='font-size:12px'>O raio informado neste campo controla dinamicamente os vizinhos de cada ponto. O valor é aplicado separadamente por lavoura, elemento e profundidade e fica registrado na análise. A consolidação usa somente a profundidade 0–10 cm.</p>", unsafe_allow_html=True)
        action_left, action_right = st.columns([1, 1])
        with action_left:
            available = bool(st.session_state.analyses)
            consolidated_data = export_consolidated([(a["result"], a["polygons"]) for a in st.session_state.analyses]) if available else b""
            st.download_button("↓  Baixar médias", consolidated_data, "outlieria_analises_consolidadas_2026.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", disabled=not available, use_container_width=False)
        with action_right:
            ready = st.button("✓  Pronto", type="primary", use_container_width=False)
        if ready:
            if not excel_bytes or not shp_bytes_cached:
                st.error("Selecione a planilha Excel e o limite SHP antes de continuar.")
            elif st.session_state.get("processing_future") and not st.session_state.processing_future.done():
                st.info("Processamento em andamento; você pode trocar de navegação sem interrompê-lo.")
            else:
                st.session_state.processing_future = PROCESS_EXECUTOR.submit(
                    process_uploaded_source_job, excel_bytes, shp_bytes_cached, radius_m, excel_name, shp_name_cached
                )
                st.info("Processamento iniciado. A troca de navegação não interromperá esta análise.")
        processing_future = st.session_state.get("processing_future")
        if processing_future and not processing_future.done() and page == "Fontes":
            st.info("Processamento em andamento. Esta análise continua rodando mesmo ao alternar para a Visão 2026.")
        if processing_future and processing_future.done():
            try:
                record, completion_message = processing_future.result()
                upsert_session_analysis(record)
                st.session_state.vision_2026_result = build_vision_snapshot(st.session_state.analyses)
                st.session_state.vision_2026_revision += 1
                st.session_state.pop("processing_future", None)
                st.success(completion_message or f"Aba 2026 atualizada: {len(record['evidence'])} outliers confirmados exibidos.")
            except Exception as processing_exc:
                st.session_state.pop("processing_future", None)
                st.error(f"Não foi possível concluir o processamento: {processing_exc}")
        st.markdown(f"<div class='section-head'><div><div class='section-kicker'>análises inseridas</div><h2>Arquivos disponíveis</h2></div><span class='source-count'>{len(st.session_state.analyses)} análise{'s' if len(st.session_state.analyses) != 1 else ''}</span></div>", unsafe_allow_html=True)
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        if not st.session_state.analyses:
            st.info("Nenhuma análise processada nesta sessão.")
        else:
            for analysis_index, analysis in enumerate(st.session_state.analyses):
                result = analysis["result"]
                client_name = str(result.client.dropna().iloc[0]) if not result.empty else "—"
                farm_name = str(result.farm.dropna().iloc[0]) if not result.empty else "—"
                area_ha = sum(p.area_ha for p in analysis["polygons"])
                with st.container(border=True):
                    row_main, row_client, row_farm, row_date, row_area, row_download = st.columns([2.5, 1.2, 1.35, 1.0, 1.0, 1.35])
                    with row_main:
                        storage_note = "arquivos e resultados persistidos" if analysis.get("persisted") else "arquivos armazenados nesta sessão"
                        st.markdown(f"**▧  {farm_name} · análise 2026**<br><span class='mono' style='font-size:9px;color:#7c8179'>{storage_note}</span>", unsafe_allow_html=True)
                        persisted_links = analysis.get("file_urls", {})
                        if not analysis.get("original") and persisted_links:
                            links = [f"[{label}]({url})" for label, key in (("Excel", "excel_original"), ("SHP", "shp"), ("reanálise", "reanálise_2026")) if (url := persisted_links.get(key))]
                            if links:
                                st.caption("Arquivos: " + " · ".join(links))
                    with row_client:
                        st.caption("CLIENTE"); st.markdown(f"**{client_name}**")
                    with row_farm:
                        st.caption("FAZENDA"); st.markdown(f"**{farm_name}**")
                    with row_date:
                        st.caption("ATUALIZADO EM"); st.markdown(f"**{format_source_timestamp(analysis.get('updated_at'))}**")
                    with row_area:
                        st.caption("ÁREA 2026"); st.markdown(f"**{area_ha:,.1f} ha**".replace(',', 'X').replace('.', ',').replace('X', '.'))
                    with row_download:
                        st.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)
                        export_rows = analysis["evidence"]
                        if analysis.get("original"):
                            st.download_button("⇩  BAIXAR PLANILHA", export_individual(analysis["original"], export_rows, float(analysis.get("radius_m", MAX_DISTANCE_M))), f"outlieria_reanalise_2026_{analysis['name']}", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key=f"download_{analysis_index}_{analysis['name']}")
                        elif analysis.get("file_urls", {}).get("reanálise_2026"):
                            st.link_button("↗  BAIXAR PLANILHA", analysis["file_urls"]["reanálise_2026"], use_container_width=True)
                        else:
                            st.caption("arquivo original não disponível")
                        if st.button("Excluir arquivos", key=f"delete_{analysis_index}_{analysis.get('analysis_id') or analysis['name']}", use_container_width=True):
                            try:
                                deletion = None
                                if analysis.get("analysis_id"):
                                    if not database_configured():
                                        raise RuntimeError("Supabase não está configurado; a análise não foi excluída do banco.")
                                    deletion = delete_analysis(str(analysis["analysis_id"]))
                                st.session_state.analyses.pop(analysis_index)
                                st.session_state.vision_2026_result = build_vision_snapshot(st.session_state.analyses)
                                st.session_state.vision_2026_revision += 1
                                if deletion and deletion.get("drive_files_pending"):
                                    st.warning("Análise excluída do banco e da interface. Os arquivos do Google Drive permaneceram porque o Drive não está configurado.")
                                else:
                                    st.success("Análise excluída da interface, do Supabase e dos arquivos vinculados no Google Drive.")
                                st.rerun()
                            except Exception as deletion_exc:
                                st.error(f"Não foi possível excluir a análise: {deletion_exc}")
            st.markdown("<div class='section-kicker' style='margin-top:18px'>exportação consolidada · somente 0–10 cm</div>", unsafe_allow_html=True)
    if page == "Visão 2026":
        if not st.session_state.analyses:
            st.markdown("<div class='eyebrow'>— &nbsp; RADAR DE LAVOURAS / RECORTE 2026</div><h1>O que saiu do padrão<br><em>no último ano?</em></h1><p class='lead'>O histórico define a referência. Os sinais exibidos aqui pertencem somente à aba 2026.</p>", unsafe_allow_html=True)
            st.markdown("<div class='rule-panel'><div class='rule-title'><span>Outliers apenas no último ano</span><span class='badge'>● 2026</span></div><p style='color:#707b6b'>Vá para <b>Fontes</b>, envie os arquivos e clique em <b>Pronto</b> para carregar os resultados.</p></div>", unsafe_allow_html=True)
            return
        current = st.session_state.analyses[-1]
        snapshot = st.session_state.get("vision_2026_result")
        if snapshot is None or snapshot.empty or not {"element", "depth", "value", "client", "farm"}.issubset(snapshot.columns):
            st.info("As fontes persistidas foram carregadas, mas ainda não possuem resultados detalhados suficientes para a Visão 2026. Reenvie a planilha e o SHP para atualizar essa fonte.")
            return
        result = snapshot.copy()
        result = result[~result.element.map(is_ignored_element)].copy()
        if "ml_status" not in result.columns:
            result["ml_status"] = "dados insuficientes"
        if "ml_score" not in result.columns:
            result["ml_score"] = np.nan
        if "ml_neighbor_count" not in result.columns:
            result["ml_neighbor_count"] = np.nan
        if "ml_k_effective" not in result.columns:
            result["ml_k_effective"] = np.nan
        if "ml_radius_m" not in result.columns:
            result["ml_radius_m"] = result.get("radius_m", active_radius)
        if "ml_rule" not in result.columns:
            result["ml_rule"] = "LOF não disponível nesta sessão antiga"
        if "final_status" not in result.columns:
            result["final_status"] = result.apply(combine_detection, axis=1)
        result = apply_ph_smp_validation(result)
        result["display_status"] = result.apply(display_classification, axis=1)
        element_order = []
        for analysis in st.session_state.analyses:
            source_elements = analysis.get("elements", ordered_unique(analysis["result"].element))
            for element_name in source_elements:
                if not is_ignored_element(element_name) and str(element_name) not in element_order:
                    element_order.append(str(element_name))
        st.markdown("<div class='eyebrow'>— &nbsp; RADAR DE LAVOURAS / RECORTE 2026</div><h1>O que saiu do padrão<br><em>no último ano?</em></h1><p class='lead'>O histórico define a referência. Os sinais exibidos aqui pertencem somente à aba 2026.</p>", unsafe_allow_html=True)
        filter_col1, filter_col2 = st.columns(2)
        client_options = ["Todos"] + sorted(result.client.dropna().astype(str).unique().tolist())
        saved_client = str(st.session_state.get("vision_client_filter", "Todos"))
        if saved_client not in client_options:
            saved_client = "Todos"
            st.session_state.vision_client_filter = saved_client
            st.session_state.pop("vision_client_widget", None)
        client_index = client_options.index(saved_client)
        with filter_col1:
            selected_client = st.selectbox("Cliente", client_options, index=client_index, key="vision_client_widget")
        client_changed = selected_client != saved_client
        st.session_state.vision_client_filter = selected_client
        if client_changed:
            st.session_state.vision_farm_filter = ""
            st.session_state.pop("vision_farm_widget", None)
        farm_source = result[result.client.astype(str) == selected_client] if selected_client != "Todos" else result.iloc[0:0]
        farm_options = ["Selecione um cliente"] if selected_client == "Todos" else ["Selecione uma fazenda"] + ordered_unique(farm_source.farm)
        saved_farm = str(st.session_state.get("vision_farm_filter", ""))
        if saved_farm not in farm_options:
            saved_farm = farm_options[0]
            st.session_state.vision_farm_filter = saved_farm
            st.session_state.pop("vision_farm_widget", None)
        farm_index = farm_options.index(saved_farm)
        with filter_col2:
            selected_farm = st.selectbox("Fazenda", farm_options, index=farm_index, key="vision_farm_widget", disabled=selected_client == "Todos")
        st.session_state.vision_farm_filter = selected_farm
        if selected_client == "Todos" or selected_farm in {"Selecione um cliente", "Selecione uma fazenda"}:
            st.info("Selecione um cliente e uma fazenda para carregar os dados da Visão 2026.")
            return
        result = result[(result.client.astype(str) == selected_client) & (result.farm.astype(str) == selected_farm)]
        detailed_result = result[result["element"].astype(str).str.strip().ne("") & result["value"].notna()]
        if detailed_result.empty:
            st.info("Cliente e fazenda carregados do banco. Esta fonte possui apenas o resumo da análise; não há resultados detalhados por ponto persistidos para exibir.")
            return
        result = result.copy(); result["status_label"] = result.final_status.map({"outlier confirmado":"outlier confirmado", "outlier estatístico":"outlier estatístico", "anomalia temporal":"anomalia temporal", "anomalia espacial":"anomalia espacial", "atenção":"atenção", "normal":"padrão"})
        points = result[["source_row", "point_id"]].drop_duplicates().shape[0]; outliers = int((result.final_status == "outlier confirmado").sum()); crops = result.crop.nunique()
        st.markdown(f"<div class='rule-panel'><div class='rule-title'><span>Outliers confirmados do último ano</span><span class='badge'>● 2026</span></div><div class='rule-grid'><div class='rule-item'><span class='rule-num'>01</span><strong>LOF espacial</strong><p>Compara a densidade local dos vizinhos da mesma lavoura, elemento e profundidade.</p></div><div class='rule-item'><span class='rule-num'>02</span><strong>Regressão espacial</strong><p>Reconstrói a tendência da coordenada com os pontos mais próximos de cada ciclo, sem usar o ID.</p></div><div class='rule-item'><span class='rule-num'>03</span><strong>Vizinhança</strong><p>O raio informado seleciona os vizinhos e o LOF exige pelo menos 3.</p></div><div class='rule-item'><span class='rule-num'>04</span><strong>Exibição</strong><p>Qualquer diferença temporal espacial detectada gera outlier; o LOF também pode gerar o sinal.</p></div></div></div>", unsafe_allow_html=True)
        st.markdown(f"<div class='stats'>{html_stat('pontos 2026', points, 'pontos com coordenadas', 'blue')}{html_stat('lavouras', crops, 'nomes relacionados ao limite', 'lime')}{html_stat('regra', 'LOF + regressão', 'confirmação simultânea', 'sand')}{html_stat('outliers confirmados', outliers, 'únicos exibidos', 'orange')}</div>", unsafe_allow_html=True)
        c1, c2 = st.columns([3,1])
        with c1:
            st.markdown("<div class='section-head'><div><div class='section-kicker'>01 / distribuição</div><h2>Leitura dos elementos em 2026</h2></div></div>", unsafe_allow_html=True)
            # Mantém a sequência original da planilha; o catálogo global é apenas fallback.
            element_options = ordered_unique(result.element) or element_order or ["Sem dados"]
            depth_options = ordered_unique(result.depth) or ["Sem dados"]
            element = st.selectbox("Elemento", element_options, format_func=element_label, label_visibility="collapsed")
            depth = st.selectbox("Profundidade", depth_options, label_visibility="collapsed")
            chart = result[(result.element == element) & (result.depth == depth)].copy()
            crop_options = ["Todas"] + ordered_unique(chart.get("graph_crop", pd.Series(dtype=object)))
            selected_chart_crop = st.selectbox("Lavoura", crop_options, key="vision_chart_crop", label_visibility="collapsed")
            if selected_chart_crop != "Todas":
                chart = chart[chart.graph_crop.astype(str) == str(selected_chart_crop)].copy()
            if not chart.empty:
                element_values = pd.to_numeric(result.loc[result.element == element, "value"], errors="coerce").dropna()
                y_domain = (float(element_values.min()), float(element_values.max())) if not element_values.empty else None
                render_scatter(chart, y_domain=y_domain)
                st.markdown("<div class='section-kicker' style='margin-top:18px'>02 / posição espacial</div><h3 style='margin-top:4px'>Onde estão os pontos selecionados</h3><p class='lead' style='font-size:12px'>Posição relativa em metros, preservando a proximidade geográfica entre longitude e latitude. A escala é igual nos dois eixos.</p>", unsafe_allow_html=True)
                render_spatial_map(chart)
        with c2:
            analysis_radius = float(current.get("radius_m", active_radius))
            st.markdown(f"<div class='source-card'><div class='section-kicker'>regra aplicada</div><h3>Último ano,<br><em>sem misturar.</em></h3><p>LOF e regressão usam a coordenada, separadamente por lavoura, elemento e profundidade. Os ciclos são relacionados espacialmente, nunca pelo ID. O LOF exige pelo menos {MIN_LOF_NEIGHBORS} vizinhos e a regressão {MIN_REGRESSION_CYCLES} ciclos.</p><hr><b>{analysis_radius:g} m</b><br><span class='mono'>raio espacial ativo · coordenada como referência</span></div>", unsafe_allow_html=True)
        st.markdown("<div class='section-head'><div><div class='section-kicker'>02 / inspeção</div><h2>Pontos da planilha 2026</h2></div></div>", unsafe_allow_html=True)
        # Além dos outliers confirmados, a validação química pH/SMP precisa ser
        # exibida mesmo quando LOF e regressão não confirmam o ponto.
        table = select_evidence(result)
        # Estado visual original da tabela nativa; os três selects permanecem como filtros.
        header_filters = st.columns([1.2, 1.35, 1.2, 1.2, 1.0, 0.9])
        with header_filters[1]:
            selected_crop = st.selectbox("Lavoura", ["Todas"] + sorted(table.crop.dropna().astype(str).unique().tolist()), key="table_filter_crop", label_visibility="collapsed")
        with header_filters[2]:
            selected_depth = st.selectbox("Profundidade", ["Todas"] + sorted(table.depth.dropna().astype(str).unique().tolist()), key="table_filter_depth", label_visibility="collapsed")
        with header_filters[3]:
                selected_element = st.selectbox("Elemento", ["Todos"] + element_order, key="table_filter_element", format_func=lambda value: "Todos" if value == "Todos" else element_label(value), label_visibility="collapsed")
        if selected_crop != "Todas": table = table[table.crop.astype(str) == selected_crop]
        if selected_depth != "Todas": table = table[table.depth.astype(str) == selected_depth]
        if selected_element != "Todos": table = table[table.element.astype(str) == selected_element]
        st.dataframe(frame_for_display(table), column_config={"valor": st.column_config.NumberColumn("valor", format="%.2f")}, use_container_width=True, hide_index=True, height=430)

if __name__ == "__main__":
    main()
