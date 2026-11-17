from __future__ import annotations

import io
import math
import os
from dataclasses import dataclass
from typing import Iterable

import numpy as np
import pandas as pd
import requests
import shapefile
import streamlit as st
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from shapely.geometry import Point, shape as shapely_shape

YEAR_RE = r"^\d{4}$"
CURRENT_YEAR = "2026"
MAX_DISTANCE_M = 300.0
IQR_MULTIPLIER = 1.5
ATTENTION_MARGIN = 0.10
IGNORED_COLUMNS = {"A", "B", "G", "H", "S", "T", "U", "V", "W"}
IGNORED_FROM_INDEX = 28  # AC, zero based

@dataclass
class PolygonRecord:
    geometry: object
    area_ha: float


def haversine(lon1, lat1, lon2, lat2):
    r = 6_371_000.0
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return 2 * r * math.asin(math.sqrt(a))


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
        if index >= IGNORED_FROM_INDEX or name in excluded or name in {"S", "T", "U", "V", "W"}:
            continue
        numeric = pd.to_numeric(frame[name], errors="coerce")
        if numeric.notna().any():
            elements.append(name.strip())
    return elements


def parse_workbook(data: bytes, polygons: list[PolygonRecord], client_override: str, farm_override: str):
    workbook = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
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
            all_rows.append({
                "year": year, "source_row": int(row["__excel_row"]), "sheet": year,
                "point_id": str(row.get("ID") or ""), "client": client_override or "Cliente não identificado",
                "farm": farm_override or str(row.get("Fazenda") or "Fazenda não identificada"),
                "crop": crop, "depth": str(row.get("Profundidade") or "Não informado"),
                "longitude": float(row["Longitude"]), "latitude": float(row["Latitude"]),
                "polygon_index": polygon_index, "raw": row.to_dict(),
            })
    if not all_rows:
        raise ValueError("Nenhum ponto com coordenadas foi encontrado dentro do SHP.")
    sample = pd.DataFrame(all_rows)
    elements = numeric_element_columns(next(iter(frames.values())))
    for element in elements:
        sample[element] = pd.to_numeric(sample["raw"].map(lambda x: x.get(element)), errors="coerce")
    return sample, elements, frames


def classify(sample: pd.DataFrame, elements: list[str]) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    current = sample[sample.year == CURRENT_YEAR].copy()
    historical = sample[sample.year != CURRENT_YEAR].copy()
    results, evidence, attention = [], [], []
    for idx, row in current.iterrows():
        for element in elements:
            value = row.get(element)
            if pd.isna(value):
                continue
            candidates = historical[(historical.crop == row.crop) & (historical.depth == row.depth)]
            distances = candidates.apply(lambda h: haversine(row.longitude, row.latitude, h.longitude, h.latitude), axis=1) if not candidates.empty else pd.Series(dtype=float)
            candidates = candidates.loc[distances[distances <= MAX_DISTANCE_M].index] if not candidates.empty else candidates
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
            result = {**row.to_dict(), "element": element, "value": float(value), "status": status, "direction": direction,
                      "historical_mean": mean, "q1": q1, "q3": q3, "iqr": iqr, "lower_fence": lower, "upper_fence": upper}
            results.append(result)
            if status in {"outlier", "atenção"}:
                (evidence if status == "outlier" else attention).append(result)
    return pd.DataFrame(results), pd.DataFrame(evidence), pd.DataFrame(attention)


def locate_ibge(rows: pd.DataFrame) -> pd.DataFrame:
    result = rows.copy()
    result["municipality"] = "Não localizado"
    result["state"] = "Não localizado"
    url = os.getenv("IBGE_REVERSE_URL")
    if not url:
        return result
    for index, row in result.iterrows():
        try:
            response = requests.get(url, params={"lat": row.latitude, "lon": row.longitude}, timeout=8)
            payload = response.json()
            result.at[index, "municipality"] = payload.get("municipality", payload.get("municipio", "Não localizado"))
            result.at[index, "state"] = payload.get("state", payload.get("estado", "Não localizado"))
        except Exception:
            continue
    return result


def export_individual(original: bytes, evidence: pd.DataFrame) -> bytes:
    workbook = load_workbook(io.BytesIO(original))
    sheet = workbook[CURRENT_YEAR]
    headers = {str(cell.value or "").strip().lower(): cell.column for cell in sheet[1]}
    fill = PatternFill("solid", fgColor="F6B26B")
    for _, item in evidence.iterrows():
        column = headers.get(str(item.element).strip().lower())
        if not column:
            continue
        cell = sheet.cell(int(item.source_row), column)
        cell.fill = fill
        cell.font = Font(name=cell.font.name, sz=cell.font.sz, bold=True, color="7A2E00")
        cell.comment = Comment(f"OUTLIER 2026 comparado ao histórico regional de 300 m. Valor {item.value}; média histórica {item.historical_mean:.3f}; Q1 {item.q1:.3f}; Q3 {item.q3:.3f}; IQR {item.iqr:.3f}; limites {item.lower_fence:.3f} a {item.upper_fence:.3f}.", "Outlieria")
    output = io.BytesIO(); workbook.save(output); return output.getvalue()


def export_consolidated(all_results: list[tuple[pd.DataFrame, list[PolygonRecord]]]) -> bytes:
    rows = []
    coverage = {}
    for analysis_index, (result, polygons) in enumerate(all_results):
        result = result[(result.depth == "0-10 cm") & result.value.notna()].copy()
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
    return output.getvalue()


def frame_for_display(result: pd.DataFrame) -> pd.DataFrame:
    columns = ["source_row", "point_id", "crop", "depth", "element", "value", "municipality", "state", "status", "direction", "historical_mean", "q1", "q3", "iqr", "lower_fence", "upper_fence"]
    return result[[c for c in columns if c in result.columns]].rename(columns={"source_row": "linha", "point_id": "ponto", "crop": "lavoura", "depth": "profundidade", "element": "elemento", "value": "valor", "municipality": "município", "state": "estado", "status": "sinal", "direction": "direção"})


def main():
    st.set_page_config(page_title="Outlieria", page_icon="◌", layout="wide")
    st.title("Outlieria")
    st.caption("Detecção de outliers agrícolas: o histórico define os limites; somente 2026 é sinalizado.")
    if "analyses" not in st.session_state:
        st.session_state.analyses = []
    with st.sidebar:
        st.header("Fontes")
        excel = st.file_uploader("Planilha Excel", type=["xlsx", "xls"], key="excel")
        shp = st.file_uploader("Limite SHP", type=["shp"], key="shp")
        client = st.text_input("Cliente", value="")
        farm = st.text_input("Fazenda", value="")
        ready = st.button("Pronto", type="primary", use_container_width=True)
    if ready:
        if not excel or not shp:
            st.error("Envie a planilha Excel e o limite SHP antes de clicar em Pronto.")
        else:
            with st.spinner("Processando Excel, SHP, histórico e localização..."):
                try:
                    shp_records = read_polygons(shp.getvalue())
                    sample, elements, _ = parse_workbook(excel.getvalue(), shp_records, client, farm)
                    result, evidence, attention = classify(sample, elements)
                    result = locate_ibge(result)
                    evidence = result[result.status == "outlier"].copy()
                    attention = result[result.status == "atenção"].copy()
                    st.session_state.analyses.append({"name": excel.name, "original": excel.getvalue(), "result": result, "evidence": evidence, "attention": attention, "polygons": shp_records})
                    st.success(f"Análise concluída: {len(evidence)} outliers e {len(attention)} pontos em atenção.")
                except Exception as exc:
                    st.exception(exc)
    if not st.session_state.analyses:
        st.info("Envie os dois arquivos na barra lateral e clique em Pronto para iniciar.")
        return
    current = st.session_state.analyses[-1]
    result = current["result"]
    outlier_count = int((result.status == "outlier").sum())
    attention_count = int((result.status == "atenção").sum())
    points = result[["source_row", "point_id"]].drop_duplicates().shape[0]
    tab_view, tab_sources = st.tabs(["Visão 2026", "Fontes"])
    with tab_view:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Pontos 2026", points); c2.metric("Outliers 2026", outlier_count); c3.metric("Atenção", attention_count); c4.metric("Profundidade", "0–10 / 10–20 cm")
        st.subheader("Regra aplicada")
        st.info("O histórico anterior a 2026 calcula Q1, Q3 e IQR em raio de 300 m. Os valores de 2026 são comparados contra essas cercas, sem misturar profundidades.")
        element = st.selectbox("Elemento", sorted(result.element.unique()))
        depth = st.selectbox("Profundidade", sorted(result.depth.unique()))
        view = result[(result.element == element) & (result.depth == depth)].copy()
        st.scatter_chart(view, x="source_row", y="value", color="status")
        mode = st.radio("Inspeção", ["Todos", "Outliers", "Atenção"], horizontal=True)
        if mode == "Outliers": view = view[view.status == "outlier"]
        if mode == "Atenção": view = view[view.status == "atenção"]
        st.dataframe(frame_for_display(view), use_container_width=True, hide_index=True)
    with tab_sources:
        st.subheader("Downloads")
        st.download_button("Baixar planilha original com outliers destacados", export_individual(current["original"], current["evidence"]), "outlieria_reanalise_2026.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        all_results = [(a["result"], a["polygons"]) for a in st.session_state.analyses]
        st.download_button("Baixar análises 2026 — médias consolidadas", export_consolidated(all_results), "outlieria_analises_consolidadas_2026.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.subheader("Análises inseridas")
        for analysis in st.session_state.analyses:
            st.write(f"**{analysis['name']}** — {len(analysis['evidence'])} outliers; {len(analysis['attention'])} em atenção")

if __name__ == "__main__":
    main()
