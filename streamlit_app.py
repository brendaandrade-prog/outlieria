from __future__ import annotations

import io
import html
import gzip
import json
import math
import os
import unicodedata
from datetime import datetime
from dataclasses import dataclass
from typing import Iterable

import numpy as np
import pandas as pd
import shapefile
import streamlit as st
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from shapely.geometry import Point, shape as shapely_shape
from cloud_storage import configured as cloud_storage_configured, missing_settings, persist_analysis
try:
    from sklearn.neighbors import LocalOutlierFactor
    from sklearn.preprocessing import RobustScaler
    LOF_AVAILABLE = True
except ImportError:
    LocalOutlierFactor = None
    RobustScaler = None
    LOF_AVAILABLE = False

YEAR_RE = r"^\d{4}$"
CURRENT_YEAR = "2026"
MAX_DISTANCE_M = 300.0
IQR_MULTIPLIER = 1.5
LOF_THRESHOLD_PERCENTILE = 90.0
ATTENTION_MARGIN = 0.0
REGRESSION_CHANGE_THRESHOLD = 0.20
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
    "value", "regional_mean", "regional_median", "regional_iqr",
    "regional_std", "neighbor_count", "mean_distance_m",
    "nearest_distance_m", "difference_median", "z_score",
]

@dataclass
class PolygonRecord:
    geometry: object
    area_ha: float


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
    reader = shapefile.Reader(shp=io.BytesIO(shp_bytes))
    records = []
    for shp in reader.shapes():
        geometry = shapely_shape(shp.__geo_interface__)
        records.append(PolygonRecord(geometry, polygon_area_ha(geometry)))
    return records


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
        names = {str(x or "").strip().lower() for x in row}
        if {"longitude", "latitude"}.issubset(names):
            return i
    return 0


def normalise_columns(columns: Iterable[object]) -> list[str]:
    result = []
    for col in columns:
        text = str(col or "").strip()
        result.append(text if text else f"coluna_{len(result) + 1}")
    return result


def numeric_element_columns(frame: pd.DataFrame) -> list[str]:
    excluded = {"Registro", "Descrição antiga", "Fazenda", "Longitude", "Latitude", "Área/Lavoura", "Pontos", "ID", "Profundidade"}
    elements = []
    for index, name in enumerate(frame.columns):
        normalized_name = str(name).strip().upper()
        if index in IGNORED_EXCEL_COLUMN_INDICES or index >= IGNORED_FROM_INDEX or name in excluded or is_ignored_element(name):
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
    for year, frame in frames.items():
        required = {"Longitude", "Latitude", "Profundidade"}
        if not required.issubset(frame.columns):
            continue
        frame["Longitude"] = pd.to_numeric(frame["Longitude"], errors="coerce")
        frame["Latitude"] = pd.to_numeric(frame["Latitude"], errors="coerce")
        frame = frame.dropna(subset=["Longitude", "Latitude"])
        for _, row in frame.iterrows():
            point = Point(float(row["Longitude"]), float(row["Latitude"]))
            polygon_index = next((i for i, p in enumerate(polygons) if p.geometry.contains(point) or p.geometry.touches(point)), -1)
            if polygon_index < 0:
                continue
            crop = str(row.get("Área/Lavoura") or "Lavoura não identificada")
            # Mantém no processamento somente os elementos válidos. As colunas
            # S–W e AC em diante continuam existindo apenas no arquivo original.
            analysis_raw = {element: row.get(element) for element in elements}
            all_rows.append({
                "year": year, "source_row": int(row["__excel_row"]), "sheet": year,
                "point_id": str(row.get("ID") or ""), "client": client_override or detected_client,
                "farm": farm_override or str(row.get("Fazenda") or detected_farm or "Fazenda não identificada"),
                "crop": crop, "depth": str(row.get("Profundidade") or "Não informado"),
                "longitude": float(row["Longitude"]), "latitude": float(row["Latitude"]),
                "polygon_index": polygon_index, "raw": analysis_raw,
            })
    if not all_rows:
        raise ValueError("Nenhum ponto com coordenadas foi encontrado dentro do SHP.")
    sample = pd.DataFrame(all_rows)
    for element in elements:
        sample[element] = pd.to_numeric(sample["raw"].map(lambda x: x.get(element)), errors="coerce")
    return sample, elements, frames


def classify(sample: pd.DataFrame, elements: list[str], radius_m: float = MAX_DISTANCE_M) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    current = sample[sample.year == CURRENT_YEAR].copy()
    historical = sample[sample.year != CURRENT_YEAR].copy()
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
            candidates = historical[(historical.crop == row.crop) & (historical.depth == row.depth)]
            distances = candidates.apply(lambda h: haversine(row.longitude, row.latitude, h.longitude, h.latitude), axis=1) if not candidates.empty else pd.Series(dtype=float)
            candidates = candidates.loc[distances[distances <= radius_m].index] if not candidates.empty else candidates
            values = pd.to_numeric(candidates[element], errors="coerce").dropna().to_numpy() if not candidates.empty else np.array([])
            if len(values) < 2:
                status, direction, q1, q3, iqr, lower, upper, mean = "normal", "", np.nan, np.nan, np.nan, np.nan, np.nan, np.nan
            else:
                q1, q3 = np.percentile(values, [25, 75]); iqr = q3 - q1
                lower, upper = q1 - IQR_MULTIPLIER * iqr, q3 + IQR_MULTIPLIER * iqr
                mean = float(np.mean(values)); direction = "abaixo" if value < lower else "acima" if value > upper else ""
                status = "outlier" if direction else "normal"
                margin = abs(iqr) * ATTENTION_MARGIN
                if not direction and (abs(value - lower) <= margin or abs(value - upper) <= margin):
                    status = "atenção"; direction = "abaixo" if abs(value - lower) < abs(value - upper) else "acima"
            result = {**row.to_dict(), "sample_index": idx, "element": element, "value": float(value), "status": status, "direction": direction,
                      "historical_mean": mean, "q1": q1, "q3": q3, "iqr": iqr, "lower_fence": lower, "upper_fence": upper}
            results.append(result)
            if status in {"outlier", "atenção"}:
                (evidence if status == "outlier" else attention).append(result)
    return pd.DataFrame(results), pd.DataFrame(evidence), pd.DataFrame(attention)


def build_lof_features(target: pd.Series, reference: pd.DataFrame, element: str, radius_m: float, exclude_index=None, minimum_neighbors: int = 3):
    """Monta o vetor regional de um ponto usando apenas seu histórico local."""
    candidates = reference[
        (reference["crop"].astype(str) == str(target["crop"])) &
        (reference["depth"].astype(str) == str(target["depth"]))
    ].copy()
    if exclude_index is not None:
        candidates = candidates[candidates.index != exclude_index]
    if candidates.empty:
        return None
    candidates["distance_m"] = candidates.apply(
        lambda item: haversine(float(target["longitude"]), float(target["latitude"]), float(item["longitude"]), float(item["latitude"])),
        axis=1,
    )
    neighbors = candidates[candidates["distance_m"] <= radius_m].copy()
    values = pd.to_numeric(neighbors[element], errors="coerce").dropna()
    value = pd.to_numeric(pd.Series([target[element]]), errors="coerce").iloc[0]
    if pd.isna(value) or len(values) < minimum_neighbors:
        return None
    q1, q3 = np.percentile(values, [25, 75])
    iqr = q3 - q1
    mean = values.mean()
    median = values.median()
    std = values.std(ddof=0)
    return {
        "value": float(value),
        "regional_mean": float(mean),
        "regional_median": float(median),
        "regional_iqr": float(iqr),
        "regional_std": float(std),
        "neighbor_count": int(len(values)),
        "mean_distance_m": float(neighbors["distance_m"].mean()),
        "nearest_distance_m": float(neighbors["distance_m"].min()),
        "difference_median": float(value - median),
        "z_score": float((value - mean) / std) if std > 0 else 0.0,
    }


def train_lof_group(historical: pd.DataFrame, element: str, radius_m: float, minimum_neighbors: int = 3):
    feature_rows = []
    for index, target in historical.iterrows():
        features = build_lof_features(target, historical, element, radius_m, exclude_index=index, minimum_neighbors=minimum_neighbors)
        if features is not None:
            features["sample_index"] = index
            feature_rows.append(features)
    feature_frame = pd.DataFrame(feature_rows)
    if len(feature_frame) < 5:
        return None
    matrix = feature_frame[LOF_FEATURES].replace([np.inf, -np.inf], np.nan).dropna()
    if len(matrix) < 5:
        return None
    scaler = RobustScaler()
    scaled_matrix = scaler.fit_transform(matrix)
    n_neighbors = min(20, max(2, len(matrix) - 1))
    model = LocalOutlierFactor(n_neighbors=n_neighbors, contamination="auto", novelty=True)
    model.fit(scaled_matrix)
    historical_scores = -model.score_samples(scaled_matrix)
    return {"model": model, "scaler": scaler, "threshold": float(np.percentile(historical_scores, LOF_THRESHOLD_PERCENTILE)), "n_neighbors": n_neighbors}


def score_lof_current(current: pd.DataFrame, historical: pd.DataFrame, element: str, radius_m: float, trained: dict, minimum_neighbors: int = 3):
    predictions = []
    for index, target in current.iterrows():
        features = build_lof_features(target, historical, element, radius_m, minimum_neighbors=minimum_neighbors)
        if features is None:
            continue
        vector = pd.DataFrame([features])[LOF_FEATURES].replace([np.inf, -np.inf], np.nan)
        if vector.isna().any().any():
            continue
        score = float(-trained["model"].score_samples(trained["scaler"].transform(vector))[0])
        predictions.append({
            "sample_index": index, "element": element, "ml_score": score,
            "ml_status": "anomalia espacial" if score >= trained["threshold"] else "normal",
            "ml_neighbor_count": features["neighbor_count"],
            "ml_mean_distance_m": features["mean_distance_m"],
            "ml_regional_median": features["regional_median"],
            "ml_regional_iqr": features["regional_iqr"],
        })
    return pd.DataFrame(predictions)


def run_spatial_lof(sample: pd.DataFrame, elements: list[str], radius_m: float) -> pd.DataFrame:
    """Treina e aplica LOF por lavoura, profundidade e elemento, sem usar 2026 no treino."""
    if not LOF_AVAILABLE:
        raise RuntimeError("O pacote scikit-learn não está instalado. Aguarde o redeploy ou instale as dependências do requirements.txt.")
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
                if historical_group.empty or current_group.empty:
                    continue
                trained = train_lof_group(historical_group, element, radius_m)
                if trained is None:
                    continue
                scored = score_lof_current(current_group, historical_group, element, radius_m, trained)
                if not scored.empty:
                    scored["crop"] = crop
                    scored["depth"] = depth
                    predictions.append(scored)
    return pd.concat(predictions, ignore_index=True) if predictions else pd.DataFrame()


def run_temporal_regression(sample: pd.DataFrame, elements: list[str], radius_m: float) -> pd.DataFrame:
    """Estima a tendência histórica de cada elemento e avalia o resíduo de 2026."""
    historical = sample[sample["year"].astype(str) != CURRENT_YEAR].copy()
    current = sample[sample["year"].astype(str) == CURRENT_YEAR].copy()
    predictions = []
    for index, target in current.iterrows():
        for element in elements:
            if is_ignored_element(element) or pd.isna(target.get(element)):
                continue
            candidates = historical[
                (historical["crop"].astype(str) == str(target["crop"]))
                & (historical["depth"].astype(str) == str(target["depth"]))
            ].copy()
            if candidates.empty:
                continue
            candidates["distance_m"] = candidates.apply(
                lambda item: haversine(float(target["longitude"]), float(target["latitude"]), float(item["longitude"]), float(item["latitude"])),
                axis=1,
            )
            candidates = candidates[candidates["distance_m"] <= radius_m].copy()
            candidates["numeric_value"] = pd.to_numeric(candidates[element], errors="coerce")
            candidates["numeric_year"] = pd.to_numeric(candidates["year"], errors="coerce")
            candidates = candidates.dropna(subset=["numeric_value", "numeric_year"])
            if len(candidates) < 5 or candidates["numeric_year"].nunique() < 2:
                continue
            x = candidates["numeric_year"].to_numpy(dtype=float)
            y = candidates["numeric_value"].to_numpy(dtype=float)
            slope, intercept = np.polyfit(x, y, 1)
            expected = float(slope * float(CURRENT_YEAR) + intercept)
            residuals = y - (slope * x + intercept)
            q1, q3 = np.percentile(residuals, [25, 75])
            residual_iqr = float(q3 - q1)
            lower, upper = float(q1 - IQR_MULTIPLIER * residual_iqr), float(q3 + IQR_MULTIPLIER * residual_iqr)
            residual = float(float(target[element]) - expected)
            relative_change = residual / abs(expected) if abs(expected) > 1e-12 else (0.0 if abs(residual) <= 1e-12 else np.inf)
            predictions.append({
                "sample_index": index,
                "element": element,
                "reg_expected": expected,
                "reg_residual": residual,
                "reg_change_pct": float(relative_change * 100.0),
                "reg_history_count": int(len(candidates)),
                "reg_status": "anomalia temporal" if relative_change < -(REGRESSION_CHANGE_THRESHOLD + 1e-9) or relative_change > (REGRESSION_CHANGE_THRESHOLD + 1e-9) else "normal",
            })
    return pd.DataFrame(predictions)


def combine_detection(row):
    # A regressão temporal é um indicador independente: variação superior
    # a 10% para qualquer direção já caracteriza outlier.
    iqr_outlier = row.get("status") == "outlier"
    ml_outlier = row.get("ml_status") == "anomalia espacial"
    regression_outlier = row.get("reg_status") == "anomalia temporal"
    if regression_outlier:
        return "outlier confirmado"
    criteria_count = sum((ml_outlier, iqr_outlier))
    if criteria_count == 2:
        return "outlier confirmado"
    if criteria_count == 1:
        return "atenção"
    return "normal"


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
    evidence = evidence[~evidence.element.map(is_ignored_element)].copy()
    for _, item in evidence.iterrows():
        column = headers.get(str(item.element).strip().lower())
        if not column:
            continue
        cell = sheet.cell(int(item.source_row), column)
        cell.fill = fill
        cell.font = Font(name=cell.font.name, sz=cell.font.sz, bold=True, color="7A2E00")
        final_status = str(item.get("final_status", "outlier estatístico")).upper()
        ml_note = ""
        if str(item.get("ml_status", "")) == "anomalia espacial":
            ml_note = f" LOF espacial: score {float(item.get('ml_score', 0)):.3f}; {int(item.get('ml_neighbor_count', 0))} vizinhos."
        regression_note = ""
        if str(item.get("reg_status", "")) == "anomalia temporal":
            regression_note = f" Regressão temporal: esperado {float(item.get('reg_expected', 0)):.3f}; variação {float(item.get('reg_change_pct', 0)):.1f}%; resíduo {float(item.get('reg_residual', 0)):.3f}; {int(item.get('reg_history_count', 0))} históricos."
        cell.comment = Comment(f"{final_status} em 2026 comparado ao histórico regional de {radius_m:g} m. Valor {item.value}; média histórica {item.historical_mean:.3f}; Q1 {item.q1:.3f}; Q3 {item.q3:.3f}; IQR {item.iqr:.3f}; limites {item.lower_fence:.3f} a {item.upper_fence:.3f}.{ml_note}{regression_note}", "Outlieria")
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
    columns = ["point_id", "crop", "depth", "element", "value", "status", "ml_status", "ml_score", "ml_neighbor_count", "reg_status", "reg_expected", "reg_residual", "reg_change_pct", "final_status"]
    frame = result[[c for c in columns if c in result.columns]].copy()
    if "value" in frame.columns:
        frame["value"] = pd.to_numeric(frame["value"], errors="coerce").round(2)
    if "ml_score" in frame.columns:
        frame["ml_score"] = pd.to_numeric(frame["ml_score"], errors="coerce").round(3)
    for column in ("reg_expected", "reg_residual", "reg_change_pct"):
        if column in frame.columns:
            frame[column] = pd.to_numeric(frame[column], errors="coerce").round(3)
    if "element" in frame.columns:
        frame["element"] = frame["element"].map(element_label)
    return frame.rename(columns={"point_id": "ponto", "crop": "lavoura", "depth": "profundidade", "element": "elemento", "value": "valor", "status": "sinal IQR", "ml_status": "sinal LOF", "ml_score": "score LOF", "ml_neighbor_count": "vizinhos", "reg_status": "sinal regressão", "reg_expected": "valor esperado", "reg_residual": "resíduo temporal", "reg_change_pct": "variação regressão (%)", "final_status": "classificação final"})


def ordered_unique(values: pd.Series) -> list[str]:
    """Retorna os valores na ordem em que aparecem na planilha, sem ordenar."""
    return list(dict.fromkeys(str(value) for value in values.dropna().tolist()))


def element_label(value: object) -> str:
    return ELEMENT_LABELS.get(str(value), str(value))


def is_ignored_element(value: object) -> bool:
    normalized = unicodedata.normalize("NFKD", str(value).strip()).encode("ascii", "ignore").decode().upper()
    return normalized in IGNORED_ELEMENTS or normalized.startswith("SATURAC")


def render_scatter(result: pd.DataFrame) -> None:
    """Renderiza dispersão em SVG, sem Altair, Plotly ou dependências de gráfico."""
    points = result.dropna(subset=["source_row", "value"]).copy()
    if points.empty:
        return
    width, height = 820, 280
    left, right, top, bottom = 48, 16, 18, 38
    x_values = pd.to_numeric(points["source_row"], errors="coerce").fillna(0).to_numpy(dtype=float)
    y_values = pd.to_numeric(points["value"], errors="coerce").to_numpy(dtype=float)
    raw_x_min, raw_x_max = float(np.nanmin(x_values)), float(np.nanmax(x_values))
    raw_y_min, raw_y_max = float(np.nanmin(y_values)), float(np.nanmax(y_values))
    axis_steps = [0.5, 1, 3, 5, 10, 25]
    def rounded_axis(raw_min, raw_max):
        span = max(raw_max - raw_min, 0.5)
        step = min(axis_steps, key=lambda candidate: abs(candidate - span / 4))
        axis_min = math.floor(raw_min / step) * step
        axis_max = math.ceil(raw_max / step) * step
        if axis_min == axis_max: axis_max = axis_min + step
        return axis_min, axis_max, step
    x_min, x_max, x_step = rounded_axis(raw_x_min, raw_x_max)
    y_min, y_max, y_step = rounded_axis(raw_y_min, raw_y_max)
    def sx(x): return left + ((x - x_min) / (x_max - x_min)) * (width - left - right)
    def sy(y): return top + ((y_max - y) / (y_max - y_min)) * (height - top - bottom)
    def tick_text(value, step):
        return f"{value:.1f}" if step == 0.5 else f"{value:.0f}"
    colors = {"outlier confirmado": "#c94b35", "outlier estatístico": "#d8793e", "anomalia espacial": "#7656a8", "outlier": "#d8793e", "atenção": "#b79b43", "normal": "#487c80"}
    circles = []
    for (_, row), x, y in zip(points.iterrows(), x_values, y_values):
        status = str(row.get("final_status", row.get("status", "normal"))); color = colors.get(status, "#487c80")
        px, py = sx(x), sy(y)
        tooltip_x, tooltip_y = min(px + 9, width - 158), max(py - 42, 5)
        point_id = html.escape(str(row.get("point_id", row.get("source_row", ""))))
        ml_score = row.get("ml_score")
        score_text = f" · LOF: {float(ml_score):.2f}" if pd.notna(ml_score) else ""
        circles.append(f'<g class="outlieria-point"><circle cx="{px:.1f}" cy="{py:.1f}" r="4.5" fill="{color}" stroke="#fffdf8" stroke-width="1.5"/><g class="point-tooltip"><rect x="{tooltip_x:.1f}" y="{tooltip_y:.1f}" width="148" height="46" rx="5" fill="#ffffff" stroke="#dce1d1"/><text x="{tooltip_x + 9:.1f}" y="{tooltip_y + 14:.1f}" font-family="IBM Plex Mono, monospace" font-size="9" fill="#173638">Vértice / ID: {point_id}</text><text x="{tooltip_x + 9:.1f}" y="{tooltip_y + 27:.1f}" font-family="IBM Plex Mono, monospace" font-size="9" fill="#5d6d68">Valor: {y:.2f}</text><text x="{tooltip_x + 9:.1f}" y="{tooltip_y + 40:.1f}" font-family="IBM Plex Mono, monospace" font-size="8" fill="#7656a8">{html.escape(status)}{score_text}</text></g></g>')
    grid_parts = []
    y_ticks = np.arange(y_min, y_max + y_step * 0.01, y_step)
    for y_tick in y_ticks:
        fraction = (y_max - y_tick) / (y_max - y_min)
        gy = top + fraction * (height - top - bottom)
        grid_parts.append(f'<line x1="{left}" x2="{width-right}" y1="{gy:.1f}" y2="{gy:.1f}" stroke="#e5e8df" stroke-width="1"/><text x="{left - 7}" y="{gy + 3:.1f}" text-anchor="end" font-family="IBM Plex Mono, monospace" font-size="9" fill="#7c8880">{tick_text(y_tick, y_step)}</text>')
    x_ticks = np.arange(x_min, x_max + x_step * 0.01, x_step)
    for x_tick in x_ticks:
        fraction = (x_tick - x_min) / (x_max - x_min)
        gx = left + fraction * (width - left - right)
        grid_parts.append(f'<text x="{gx:.1f}" y="{height - bottom + 17}" text-anchor="middle" font-family="IBM Plex Mono, monospace" font-size="9" fill="#7c8880">{tick_text(x_tick, x_step)}</text>')
    grid = "".join(grid_parts)
    legend = '<g font-family="IBM Plex Mono, monospace" font-size="10"><circle cx="55" cy="265" r="4" fill="#487c80"/><text x="64" y="269" fill="#6f7c76">padrão</text><circle cx="140" cy="265" r="4" fill="#b79b43"/><text x="149" y="269" fill="#6f7c76">atenção</text><circle cx="225" cy="265" r="4" fill="#d8793e"/><text x="234" y="269" fill="#6f7c76">outlier 2026</text></g>'
    st.markdown(f'<div style="background:#fffdf8;border:1px solid #dce1d1;border-radius:10px;padding:4px 8px"><svg viewBox="0 0 {width} {height}" width="100%" height="280" role="img" aria-label="Gráfico de dispersão"><style>.outlieria-point .point-tooltip{{opacity:0;pointer-events:none;transition:opacity .12s}}.outlieria-point:hover .point-tooltip{{opacity:1}}</style>{grid}<line x1="{left}" x2="{left}" y1="{top}" y2="{height-bottom}" stroke="#aeb8ad"/><line x1="{left}" x2="{width-right}" y1="{height-bottom}" y2="{height-bottom}" stroke="#aeb8ad"/>{"".join(circles)}{legend}</svg></div>', unsafe_allow_html=True)


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
    active_radius = float(st.session_state.analyses[-1].get("radius_m", MAX_DISTANCE_M)) if st.session_state.analyses else MAX_DISTANCE_M
    with st.sidebar:
        st.markdown("<div class='brand'><span class='brand-mark'></span><span>outlieria</span></div>", unsafe_allow_html=True)
        st.markdown("<div class='rail-caption'>análise agrícola</div>", unsafe_allow_html=True)
        page = st.radio("Navegação", ["Visão 2026", "Fontes"], label_visibility="collapsed")
        st.markdown(f"<div class='rail-caption' style='margin-top:30px'>recorte ativo</div><div class='scope'><strong>2026</strong><span>último ano · {active_radius:g} m</span></div>", unsafe_allow_html=True)
    st.markdown(f"<div class='topbar'><div class='mono' style='color:#737d76'>ANÁLISE AGRÍCOLA &nbsp;/&nbsp; <b>{page.upper()}</b></div><span class='tab-loading'><span class='tab-loading-dot'></span>carregando</span><span class='badge'>● ABA 2026</span></div>", unsafe_allow_html=True)
    if page == "Fontes":
        st.markdown("<div class='eyebrow'>— &nbsp; FONTES DE DADOS</div><h1>Cadastre as fontes<br><em>uma única vez.</em></h1><p class='lead'>Envie a planilha agrícola e o limite SHP. O processamento começa somente depois de clicar em <b>Pronto</b>.</p>", unsafe_allow_html=True)
        c1, c2, c3 = st.columns([1.25, 1.25, 1.25])
        with c1: excel = st.file_uploader("Planilha Excel", type=["xlsx", "xls"], key="excel")
        with c2: shp = st.file_uploader("Limite SHP", type=["shp"], key="shp")
        with c3:
            radius_m = st.number_input("Raio regional (m)", min_value=1.0, max_value=100000.0, value=active_radius, step=10.0, format="%.0f", help="Distância usada para selecionar o histórico regional de cada ponto. Utilize sempre o raio do maior grid para a comparação.", key="radius_m")
        enable_lof = st.checkbox("Ativar LOF espacial", value=False, disabled=not LOF_AVAILABLE, help="Usa aprendizado não supervisionado para comparar cada ponto com o histórico da mesma lavoura, elemento e profundidade dentro do raio informado.", key="enable_lof")
        if not LOF_AVAILABLE:
            st.caption("LOF indisponível até a instalação do scikit-learn pelo requirements.txt.")
        st.markdown("<p class='lead' style='font-size:12px'>O raio informado é aplicado separadamente a cada ponto, mantendo lavoura e profundidade isoladas. A consolidação usa somente a profundidade 0–10 cm.</p>", unsafe_allow_html=True)
        action_left, action_right = st.columns([1, 1])
        with action_left:
            available = bool(st.session_state.analyses)
            consolidated_data = export_consolidated([(a["result"], a["polygons"]) for a in st.session_state.analyses]) if available else b""
            st.download_button("↓  Baixar médias", consolidated_data, "outlieria_analises_consolidadas_2026.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", disabled=not available, use_container_width=False)
        with action_right:
            ready = st.button("✓  Pronto", type="primary", use_container_width=False)
        if ready:
            if not excel or not shp: st.error("Selecione a planilha Excel e o limite SHP antes de continuar.")
            else:
                with st.spinner("Processando Excel, SHP, histórico e localização..."):
                    try:
                        polygons = read_polygons(shp.getvalue())
                        sample, elements, _ = parse_workbook(excel.getvalue(), polygons, "", "")
                        radius_m = validate_radius(radius_m)
                        result, _, _ = classify(sample, elements, radius_m); result = locate_ibge(result)
                        with st.spinner("Calculando variação temporal por elemento..."):
                            regression_result = run_temporal_regression(sample, elements, radius_m)
                        if regression_result.empty:
                            result["reg_status"] = "dados insuficientes"
                            result["reg_expected"] = np.nan
                            result["reg_residual"] = np.nan
                            result["reg_change_pct"] = np.nan
                            result["reg_history_count"] = np.nan
                        else:
                            result = result.merge(regression_result, on=["sample_index", "element"], how="left")
                            result["reg_status"] = result["reg_status"].fillna("dados insuficientes")
                        if enable_lof:
                            with st.spinner("Executando LOF espacial no histórico regional..."):
                                lof_result = run_spatial_lof(sample, elements, radius_m)
                        else:
                            lof_result = pd.DataFrame()
                        if lof_result.empty:
                            result["ml_status"] = "desativado"
                            result["ml_score"] = np.nan
                            result["ml_neighbor_count"] = np.nan
                            result["ml_mean_distance_m"] = np.nan
                            result["ml_regional_median"] = np.nan
                            result["ml_regional_iqr"] = np.nan
                        else:
                            result = result.merge(lof_result[["sample_index", "element", "ml_score", "ml_status", "ml_neighbor_count", "ml_mean_distance_m", "ml_regional_median", "ml_regional_iqr"]], on=["sample_index", "element"], how="left")
                            result["ml_status"] = result["ml_status"].fillna("dados insuficientes")
                        result["final_status"] = result.apply(combine_detection, axis=1)
                        evidence = result[result.final_status == "outlier confirmado"].copy()
                        attention = result[result.final_status == "atenção"].copy()
                        original_bytes = excel.getvalue()
                        shp_bytes = shp.getvalue()
                        reanalysis_bytes = export_individual(original_bytes, evidence, radius_m)
                        client_name = str(result.client.dropna().iloc[0]) if not result.empty and result.client.notna().any() else "Sem cliente"
                        farm_name = str(result.farm.dropna().iloc[0]) if not result.empty and result.farm.notna().any() else "Sem fazenda"
                        area_ha = sum(p.area_ha for p in polygons)
                        st.session_state.analyses.append({"name": excel.name, "original": original_bytes, "shp_bytes": shp_bytes, "radius_m": radius_m, "result": result, "evidence": evidence, "attention": attention, "polygons": polygons, "elements": elements})
                        if cloud_storage_configured():
                            try:
                                persist_analysis(client_name=client_name, farm_name=farm_name, year=int(CURRENT_YEAR), radius_m=radius_m, original=original_bytes, shp_bytes=shp_bytes, reanalysis=reanalysis_bytes, original_name=excel.name, area_ha=area_ha, point_count=len(result), outlier_count=len(evidence), attention_count=len(attention))
                                st.success("Análise processada e persistida no Supabase e no Google Drive.")
                            except Exception as storage_exc:
                                st.warning(f"A análise foi processada nesta sessão, mas não pôde ser sincronizada com a nuvem: {storage_exc}")
                        else:
                            missing = ", ".join(missing_settings())
                            st.warning(f"A análise foi processada somente nesta sessão. Configure estas Secrets no Streamlit Cloud para ativar o Supabase e o Google Drive: {missing}.")
                        st.success(f"Aba 2026 atualizada: {len(evidence)} outliers e {len(attention)} pontos em atenção.")
                    except Exception as exc: st.exception(exc)
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
                        st.markdown(f"**▧  {farm_name} · análise 2026**<br><span class='mono' style='font-size:9px;color:#7c8179'>planilha + limite SHP armazenados nesta sessão</span>", unsafe_allow_html=True)
                    with row_client:
                        st.caption("CLIENTE"); st.markdown(f"**{client_name}**")
                    with row_farm:
                        st.caption("FAZENDA"); st.markdown(f"**{farm_name}**")
                    with row_date:
                        st.caption("INSERÇÃO"); st.markdown(f"**{datetime.now().strftime('%d %b %Y')}**")
                    with row_area:
                        st.caption("ÁREA 2026"); st.markdown(f"**{area_ha:,.1f} ha**".replace(',', 'X').replace('.', ',').replace('X', '.'))
                    with row_download:
                        st.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)
                        st.download_button("⇩  BAIXAR PLANILHA", export_individual(analysis["original"], analysis["evidence"], float(analysis.get("radius_m", MAX_DISTANCE_M))), f"outlieria_reanalise_2026_{analysis['name']}", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key=f"download_{analysis_index}_{analysis['name']}")
            st.markdown("<div class='section-kicker' style='margin-top:18px'>exportação consolidada · somente 0–10 cm</div>", unsafe_allow_html=True)
        return
    if not st.session_state.analyses:
        st.markdown("<div class='eyebrow'>— &nbsp; RADAR DE LAVOURAS / RECORTE 2026</div><h1>O que saiu do padrão<br><em>no último ano?</em></h1><p class='lead'>O histórico define a referência. Os sinais exibidos aqui pertencem somente à aba 2026.</p>", unsafe_allow_html=True)
        st.markdown("<div class='rule-panel'><div class='rule-title'><span>Outliers apenas no último ano</span><span class='badge'>● 2026</span></div><p style='color:#707b6b'>Vá para <b>Fontes</b>, envie os arquivos e clique em <b>Pronto</b> para carregar os resultados.</p></div>", unsafe_allow_html=True)
        return
    current = st.session_state.analyses[-1]
    result = pd.concat([analysis["result"] for analysis in st.session_state.analyses], ignore_index=True).copy()
    result = result[~result.element.map(is_ignored_element)].copy()
    if "ml_status" not in result.columns:
        result["ml_status"] = "desativado"
    if "ml_score" not in result.columns:
        result["ml_score"] = np.nan
    if "ml_neighbor_count" not in result.columns:
        result["ml_neighbor_count"] = np.nan
    if "final_status" not in result.columns:
        result["final_status"] = result.apply(combine_detection, axis=1)
    element_order = []
    for analysis in st.session_state.analyses:
        source_elements = analysis.get("elements", ordered_unique(analysis["result"].element))
        for element_name in source_elements:
            if not is_ignored_element(element_name) and str(element_name) not in element_order:
                element_order.append(str(element_name))
    st.markdown("<div class='eyebrow'>— &nbsp; RADAR DE LAVOURAS / RECORTE 2026</div><h1>O que saiu do padrão<br><em>no último ano?</em></h1><p class='lead'>O histórico define a referência. Os sinais exibidos aqui pertencem somente à aba 2026.</p>", unsafe_allow_html=True)
    filter_col1, filter_col2 = st.columns(2)
    client_options = ["Todos"] + sorted(result.client.dropna().astype(str).unique().tolist())
    with filter_col1:
        selected_client = st.selectbox("Cliente", client_options, key="analysis_client")
    farm_source = result[result.client.astype(str) == selected_client] if selected_client != "Todos" else result.iloc[0:0]
    farm_options = ["Selecione um cliente"] if selected_client == "Todos" else ["Selecione uma fazenda"] + ordered_unique(farm_source.farm)
    with filter_col2:
        selected_farm = st.selectbox("Fazenda", farm_options, key="analysis_farm", disabled=selected_client == "Todos")
    if selected_client == "Todos" or selected_farm in {"Selecione um cliente", "Selecione uma fazenda"}:
        st.info("Selecione um cliente e uma fazenda para carregar os dados da Visão 2026.")
        return
    result = result[(result.client.astype(str) == selected_client) & (result.farm.astype(str) == selected_farm)]
    result = result.copy(); result["status_label"] = result.final_status.map({"outlier confirmado":"outlier confirmado", "outlier estatístico":"outlier estatístico", "anomalia espacial":"anomalia espacial", "atenção":"atenção", "normal":"padrão"})
    points = result[["source_row", "point_id"]].drop_duplicates().shape[0]; outliers = int(result.final_status.isin(["outlier confirmado", "outlier estatístico", "anomalia espacial"]).sum()); attention = int((result.final_status == "atenção").sum()); crops = result.crop.nunique()
    st.markdown(f"<div class='rule-panel'><div class='rule-title'><span>Outliers apenas no último ano</span><span class='badge'>● 2026</span></div><div class='rule-grid'><div class='rule-item'><span class='rule-num'>01</span><strong>LOF espacial</strong><p>O modelo compara a densidade local dos vizinhos da mesma lavoura.</p></div><div class='rule-item'><span class='rule-num'>02</span><strong>IQR regional</strong><p>O valor é comparado às cercas Q1 − 1,5×IQR e Q3 + 1,5×IQR.</p></div><div class='rule-item'><span class='rule-num'>03</span><strong>Regressão temporal</strong><p>Variação superior a 20% para mais ou para menos frente ao esperado é outlier.</p></div><div class='rule-item'><span class='rule-num'>04</span><strong>Classificação</strong><p>A regressão pode confirmar o outlier; IQR e LOF juntos também confirmam.</p></div></div></div>", unsafe_allow_html=True)
    st.markdown(f"<div class='stats'>{html_stat('pontos 2026', points, 'pontos com coordenadas', 'blue')}{html_stat('lavouras', crops, 'nomes relacionados ao limite', 'lime')}{html_stat('atenção', attention, 'um sinal IQR ou LOF', 'sand')}{html_stat('outliers 2026', outliers, 'regressão ou IQR + LOF', 'orange')}</div>", unsafe_allow_html=True)
    c1, c2 = st.columns([3,1])
    with c1:
        st.markdown("<div class='section-head'><div><div class='section-kicker'>01 / distribuição</div><h2>Leitura dos elementos em 2026</h2></div></div>", unsafe_allow_html=True)
        element_options = element_order or ["Sem dados"]
        depth_options = ordered_unique(result.depth) or ["Sem dados"]
        element = st.selectbox("Elemento", element_options, format_func=element_label, label_visibility="collapsed")
        depth = st.selectbox("Profundidade", depth_options, label_visibility="collapsed")
        chart = result[(result.element == element) & (result.depth == depth)].copy()
        if not chart.empty:
            render_scatter(chart)
    with c2:
        st.markdown(f"<div class='source-card'><div class='section-kicker'>regra aplicada</div><h3>Último ano,<br><em>sem misturar.</em></h3><p>A classificação usa a distribuição histórica regional e mantém cada profundidade isolada.</p><hr><b>{active_radius:g} m</b><br><span class='mono'>raio espacial ativo</span></div>", unsafe_allow_html=True)
    st.markdown("<div class='section-head'><div><div class='section-kicker'>02 / inspeção</div><h2>Pontos da planilha 2026</h2></div></div>", unsafe_allow_html=True)
    table = result[result.final_status.isin(["outlier confirmado", "outlier estatístico", "anomalia espacial"])].copy()
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
