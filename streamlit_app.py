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
                "point_id": str(row.get("ID") or ""), "client": client_override or detected_client,
                "farm": farm_override or str(row.get("Fazenda") or detected_farm or "Fazenda não identificada"),
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
    [data-testid='stAppViewContainer'] .main .block-container { max-width:1180px; padding-top:1rem; padding-bottom:1.5rem; } [data-testid='stButton'] button, [data-testid='stDownloadButton'] button { min-height:32px !important; padding:6px 10px !important; border-radius:7px !important; font-size:11px !important; font-weight:600 !important; } [data-testid='stFileUploader'] section { padding:2px 5px !important; min-height:46px !important; height:46px !important; border-radius:7px !important; } [data-testid='stFileUploader'] section > div { min-height:40px !important; padding:0 !important; } [data-testid='stFileUploader'] button { padding:3px 7px !important; font-size:10px !important; } [data-testid='stFileUploader'] small { font-size:9px !important; } [data-testid='stFileUploader'] label { font-size:10px !important; } [data-testid='stFileUploaderDropzoneInstructions'] { font-size:9px !important; line-height:1.1 !important; white-space:nowrap !important; overflow:hidden !important; text-overflow:ellipsis !important; } [data-testid='stSelectbox'] > div > div { min-height:32px !important; height:32px !important; border-radius:7px !important; background:#fffdf8 !important; border-color:var(--line) !important; } [data-testid='stSelectbox'] label, .stRadio label { font-size:10px !important; } [data-testid='stVerticalBlock'] { gap:0.45rem; } @media (max-width:900px){.rule-grid,.stats{grid-template-columns:repeat(2,1fr)} h1{font-size:34px!important}} @media (max-width:600px){.rule-grid,.stats{grid-template-columns:1fr}}
    </style>
    """, unsafe_allow_html=True)


def html_stat(label, value, note, tone=''):
    return f"<div class='stat {tone}'><div class='stat-label'>{label}</div><div class='stat-value'>{value}</div><div class='stat-note'>{note}</div></div>"


def main():
    st.set_page_config(page_title="Outlieria", page_icon="◌", layout="wide", initial_sidebar_state="expanded")
    inject_style()
    if "analyses" not in st.session_state: st.session_state.analyses = []
    with st.sidebar:
        st.markdown("<div class='brand'><span class='brand-mark'></span><span>outlieria</span></div>", unsafe_allow_html=True)
        st.markdown("<div class='rail-caption'>análise agrícola</div>", unsafe_allow_html=True)
        page = st.radio("Navegação", ["Visão 2026", "Fontes"], label_visibility="collapsed")
        st.markdown("<div class='rail-caption' style='margin-top:30px'>recorte ativo</div><div class='scope'><strong>2026</strong><span>último ano · 300 m</span></div>", unsafe_allow_html=True)
    st.markdown(f"<div class='topbar'><div class='mono' style='color:#737d76'>ANÁLISE AGRÍCOLA &nbsp;/&nbsp; <b>{page.upper()}</b></div><span class='badge'>● ABA 2026</span></div>", unsafe_allow_html=True)
    if page == "Fontes":
        st.markdown("<div class='eyebrow'>— &nbsp; FONTES DE DADOS</div><h1>Cadastre as fontes<br><em>uma única vez.</em></h1><p class='lead'>Envie a planilha agrícola e o limite SHP. O processamento começa somente depois de clicar em <b>Pronto</b>.</p>", unsafe_allow_html=True)
        c1, c2, _ = st.columns([1.25, 1.25, 1.5])
        with c1: excel = st.file_uploader("Planilha Excel", type=["xlsx", "xls"], key="excel")
        with c2: shp = st.file_uploader("Limite SHP", type=["shp"], key="shp")
        st.markdown("<p class='lead' style='font-size:12px'>A fazenda e o cliente são lidos diretamente da planilha. A consolidação usa somente a profundidade 0–10 cm.</p>", unsafe_allow_html=True)
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
                        result, _, _ = classify(sample, elements); result = locate_ibge(result)
                        evidence = result[result.status == "outlier"].copy(); attention = result[result.status == "atenção"].copy()
                        st.session_state.analyses.append({"name": excel.name, "original": excel.getvalue(), "result": result, "evidence": evidence, "attention": attention, "polygons": polygons})
                        st.success(f"Aba 2026 atualizada: {len(evidence)} outliers e {len(attention)} pontos em atenção.")
                    except Exception as exc: st.exception(exc)
        st.markdown("<div class='section-head'><div><div class='section-kicker'>análises inseridas</div><h2>Arquivos disponíveis</h2></div></div>", unsafe_allow_html=True)
        if not st.session_state.analyses:
            st.info("Nenhuma análise processada nesta sessão.")
        else:
            for analysis in st.session_state.analyses:
                st.markdown(f"<div class='source-card'><div class='source-count'>fonte registrada</div><h3>{analysis['name']}</h3><p>{len(analysis['evidence'])} outliers · {len(analysis['attention'])} pontos em atenção · planilha e SHP processados</p></div>", unsafe_allow_html=True)
            current = st.session_state.analyses[-1]
            st.download_button("Baixar planilha original com outliers destacados", export_individual(current["original"], current["evidence"]), "outlieria_reanalise_2026.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            all_results = [(a["result"], a["polygons"]) for a in st.session_state.analyses]
            st.markdown("<div class='section-kicker' style='margin-top:18px'>exportação consolidada · somente 0–10 cm</div>", unsafe_allow_html=True)
        return
    if not st.session_state.analyses:
        st.markdown("<div class='eyebrow'>— &nbsp; RADAR DE LAVOURAS / RECORTE 2026</div><h1>O que saiu do padrão<br><em>no último ano?</em></h1><p class='lead'>O histórico define a referência. Os sinais exibidos aqui pertencem somente à aba 2026.</p>", unsafe_allow_html=True)
        st.markdown("<div class='rule-panel'><div class='rule-title'><span>Outliers apenas no último ano</span><span class='badge'>● 2026</span></div><p style='color:#707b6b'>Vá para <b>Fontes</b>, envie os arquivos e clique em <b>Pronto</b> para carregar os resultados.</p></div>", unsafe_allow_html=True)
        return
    current = st.session_state.analyses[-1]; result = current["result"].copy()
    points = result[["source_row", "point_id"]].drop_duplicates().shape[0]; outliers = int((result.status == "outlier").sum()); attention = int((result.status == "atenção").sum()); crops = result.crop.nunique()
    st.markdown("<div class='eyebrow'>— &nbsp; RADAR DE LAVOURAS / RECORTE 2026</div><h1>O que saiu do padrão<br><em>no último ano?</em></h1><p class='lead'>O histórico define a referência. Os sinais exibidos aqui pertencem somente à aba 2026.</p>", unsafe_allow_html=True)
    filter_col1, filter_col2 = st.columns(2)
    client_options = ["Todos"] + sorted(result.client.dropna().astype(str).unique().tolist())
    with filter_col1:
        selected_client = st.selectbox("Cliente", client_options, key="analysis_client")
    farm_source = result if selected_client == "Todos" else result[result.client.astype(str) == selected_client]
    farm_options = ["Todos"] + sorted(farm_source.farm.dropna().astype(str).unique().tolist())
    with filter_col2:
        selected_farm = st.selectbox("Fazenda", farm_options, key="analysis_farm")
    if selected_client != "Todos": result = result[result.client.astype(str) == selected_client]
    if selected_farm != "Todos": result = result[result.farm.astype(str) == selected_farm]
    result = result.copy(); result["status_label"] = result.status.map({"outlier":"outlier 2026", "atenção":"atenção", "normal":"padrão"})
    points = result[["source_row", "point_id"]].drop_duplicates().shape[0]; outliers = int((result.status == "outlier").sum()); attention = int((result.status == "atenção").sum()); crops = result.crop.nunique()
    st.markdown(f"<div class='file-strip'><div style='font-size:22px'>▣</div><div><strong>{current['name']}</strong><span>aba analisada: 2026 · {result.state.iloc[0] if 'state' in result else 'localização pendente'}</span></div><span class='badge'>● RECORTE ATUALIZADO</span></div>", unsafe_allow_html=True)
    st.markdown("<div class='rule-panel'><div class='rule-title'><span>Outliers apenas no último ano</span><span class='badge'>● 2026</span></div><div class='rule-grid'><div class='rule-item'><span class='rule-num'>01</span><strong>Média regional</strong><p>O histórico define a distribuição dos pontos da mesma região de 300 m.</p></div><div class='rule-item'><span class='rule-num'>02</span><strong>Limites por IQR</strong><p>Q1 − 1,5×IQR e Q3 + 1,5×IQR definem as cercas.</p></div><div class='rule-item'><span class='rule-num'>03</span><strong>Atenção — perto da cerca</strong><p>Valores aceitos, mas próximos do limite, aparecem em atenção.</p></div><div class='rule-item'><span class='rule-num'>04</span><strong>Profundidade isolada</strong><p>0–10 cm nunca é comparado com 10–20 cm.</p></div></div></div>", unsafe_allow_html=True)
    st.markdown(f"<div class='stats'>{html_stat('pontos 2026', points, 'pontos com coordenadas', 'blue')}{html_stat('lavouras', crops, 'nomes relacionados ao limite', 'lime')}{html_stat('atenção', attention, 'próximos das cercas IQR', 'sand')}{html_stat('outliers 2026', outliers, 'somente no último ano', 'orange')}</div>", unsafe_allow_html=True)
    c1, c2 = st.columns([3,1])
    with c1:
        st.markdown("<div class='section-head'><div><div class='section-kicker'>01 / distribuição</div><h2>Leitura dos elementos em 2026</h2></div></div>", unsafe_allow_html=True)
        element = st.selectbox("Elemento", sorted(result.element.unique()), label_visibility="collapsed")
        depth = st.selectbox("Profundidade", sorted(result.depth.unique()), label_visibility="collapsed")
        chart = result[(result.element == element) & (result.depth == depth)].copy()
        if not chart.empty: st.scatter_chart(chart, x="source_row", y="value", color="status_label", height=280)
    with c2:
        st.markdown("<div class='source-card'><div class='section-kicker'>regra aplicada</div><h3>Último ano,<br><em>sem misturar.</em></h3><p>A classificação usa a distribuição histórica regional e mantém cada profundidade isolada.</p><hr><b>300 m</b><br><span class='mono'>raio espacial</span></div>", unsafe_allow_html=True)
    st.markdown("<div class='section-head'><div><div class='section-kicker'>02 / inspeção</div><h2>Pontos da planilha 2026</h2></div></div>", unsafe_allow_html=True)
    mode = st.radio("Filtrar", ["Todos", "Outliers", "Atenção"], horizontal=True, label_visibility="collapsed")
    table = result.copy()
    if mode == "Outliers": table = table[table.status == "outlier"]
    if mode == "Atenção": table = table[table.status == "atenção"]
    st.dataframe(frame_for_display(table), use_container_width=True, hide_index=True, height=430)


if __name__ == "__main__":
    main()
