# -*- coding: utf-8 -*-
"""Excel 测试数据批量判定工具 —— 精密仪器面板风格 UI

判定规则：
- 上限、下限都有值：下限 <= 实测值 <= 上限 → OK，否则 NG
- 只有下限：实测值 >= 下限 → OK，否则 NG
- 只有上限：实测值 <= 上限 → OK，否则 NG
- 上下限都无值或实测值非数字 → NG
文件内所有数据行均为 OK，则整个文件判定为 OK，否则为 NG。
"""

from __future__ import annotations

import html
import io
import zipfile
from dataclasses import dataclass, field

import pandas as pd
import streamlit as st

REQUIRED_COLUMNS = ["上限", "下限", "实测值"]

st.set_page_config(page_title="质检判定台 · Excel 批量判定", page_icon="◈", layout="wide")

# ---------------- 视觉层 ----------------
st.markdown("""
<style>
:root{
  --paper:#F6F4EF; --ink:#15212C; --ink-soft:#4A5866; --hairline:#DDD8CC;
  --teal:#0E5F56; --ok:#17734E; --ok-bg:#E4F2EA; --ng:#B3362B; --ng-bg:#F9E7E4;
  --warn:#9A6A12; --warn-bg:#F7EEDC;
  --serif:"Songti SC","STSong","SimSun",serif;
  --sans:"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif;
  --mono:"Cascadia Mono","JetBrains Mono",Consolas,monospace;
}
html,body,[data-testid="stAppViewContainer"]{background:var(--paper);color:var(--ink);font-family:var(--sans);}
[data-testid="stAppViewContainer"]{background-image:
  linear-gradient(var(--hairline) 1px,transparent 1px),
  linear-gradient(90deg,var(--hairline) 1px,transparent 1px);
  background-size:56px 56px;background-blend-mode:multiply;}
[data-testid="stHeader"]{background:transparent;}
.block-container{max-width:1080px;padding-top:2.2rem;}

.qc-masthead{border-bottom:3px double var(--ink);padding-bottom:14px;margin-bottom:6px;}
.qc-eyebrow{font-family:var(--mono);font-size:12px;letter-spacing:.32em;color:var(--teal);text-transform:uppercase;}
.qc-title{font-family:var(--serif);font-weight:700;font-size:40px;line-height:1.15;margin:4px 0 2px;color:var(--ink);}
.qc-sub{color:var(--ink-soft);font-size:14px;}
.qc-rule-panel{background:#FFFEFA;border:1px solid var(--hairline);border-left:4px solid var(--teal);
  padding:14px 18px;margin:18px 0 8px;font-size:14px;line-height:1.9;box-shadow:2px 2px 0 var(--hairline);}
.qc-rule-panel b{font-family:var(--mono);}

.qc-metrics{display:flex;gap:14px;margin:18px 0 6px;}
.qc-card{flex:1;background:#FFFEFA;border:1px solid var(--hairline);padding:16px 18px 12px;position:relative;box-shadow:3px 3px 0 var(--hairline);}
.qc-card::before{content:"";position:absolute;top:0;left:0;right:0;height:3px;background:var(--ink);}
.qc-card.ok::before{background:var(--ok);} .qc-card.ng::before{background:var(--ng);} .qc-card.err::before{background:var(--warn);}
.qc-card .lbl{font-family:var(--mono);font-size:11px;letter-spacing:.22em;color:var(--ink-soft);text-transform:uppercase;}
.qc-card .val{font-family:var(--mono);font-size:42px;font-weight:700;line-height:1.05;margin-top:4px;}
.qc-card.ok .val{color:var(--ok);} .qc-card.ng .val{color:var(--ng);} .qc-card.err .val{color:var(--warn);}

.qc-table{width:100%;border-collapse:collapse;background:#FFFEFA;font-size:14px;margin-top:6px;}
.qc-table th{font-family:var(--mono);font-size:11px;letter-spacing:.18em;text-transform:uppercase;
  color:var(--ink-soft);border-bottom:2px solid var(--ink);padding:10px 12px;text-align:left;}
.qc-table td{border-bottom:1px solid var(--hairline);padding:10px 12px;vertical-align:middle;}
.qc-table tr:hover td{background:#F2EFE6;}
.qc-table td.num{font-family:var(--mono);}
.pill{display:inline-block;font-family:var(--mono);font-weight:700;font-size:12px;letter-spacing:.14em;
  padding:3px 12px;border-radius:999px;border:1.5px solid currentColor;}
.pill.ok{color:var(--ok);background:var(--ok-bg);} .pill.ng{color:var(--ng);background:var(--ng-bg);} .pill.err{color:var(--warn);background:var(--warn-bg);}

section[data-testid="stFileUploaderDropzone"]{background:#FFFEFA;border:1.5px dashed var(--ink-soft);}
div[data-testid="stExpander"] details{background:#FFFEFA;border:1px solid var(--hairline);}
div[data-testid="stExpander"] summary{font-family:var(--mono);font-size:13px;}
.stDownloadButton button, .stDownloadButton button:focus{
  background:var(--ink)!important;color:#FFFEFA!important;border:none!important;border-radius:2px!important;
  font-family:var(--mono)!important;letter-spacing:.12em!important;padding:.7rem 1.6rem!important;
  box-shadow:3px 3px 0 var(--hairline)!important;transition:transform .12s ease,box-shadow .12s ease;}
.stDownloadButton button:hover{transform:translate(-2px,-2px);box-shadow:5px 5px 0 var(--teal)!important;}
.qc-foot{margin-top:34px;padding-top:10px;border-top:1px solid var(--hairline);
  font-family:var(--mono);font-size:11px;letter-spacing:.14em;color:var(--ink-soft);}
</style>
""", unsafe_allow_html=True)


# ---------------- 逻辑层（与已验证版本一致） ----------------
@dataclass
class FileResult:
    name: str
    status: str
    total: int = 0
    ok_count: int = 0
    ng_count: int = 0
    message: str = ""
    row_details: list[dict] = field(default_factory=list)
    new_name: str = ""
    data: bytes = b""


def to_number(value):
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if pd.isna(value):
            return None
        return float(value)
    text = str(value).strip().replace(",", "").replace("，", "")
    if text == "" or text.lower() in {"nan", "none", "-", "/"}:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def judge_row(upper, lower, measured):
    if measured is None:
        return False, "实测值为空或不是数字"
    if upper is None and lower is None:
        return False, "上限和下限都为空，无法判定"
    if upper is not None and lower is not None:
        if lower <= measured <= upper:
            return True, ""
        return False, f"实测值 {measured} 不在 [{lower}, {upper}] 区间内"
    if lower is not None:
        if measured >= lower:
            return True, ""
        return False, f"实测值 {measured} 小于下限 {lower}"
    if measured <= upper:
        return True, ""
    return False, f"实测值 {measured} 大于上限 {upper}"


def unique_name(name: str, used: set) -> str:
    if name not in used:
        used.add(name)
        return name
    stem, dot, ext = name.rpartition(".")
    i = 2
    while True:
        candidate = f"{stem}_{i}.{ext}" if dot else f"{name}_{i}"
        if candidate not in used:
            used.add(candidate)
            return candidate
        i += 1


def process_file(uploaded_file) -> FileResult:
    result = FileResult(name=uploaded_file.name, status="OK")
    try:
        df = pd.read_excel(io.BytesIO(uploaded_file.getvalue()))
    except Exception as exc:
        result.status = "ERROR"
        result.message = f"文件读取失败：{exc}"
        return result

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        result.status = "ERROR"
        result.message = f"缺少必需列：{', '.join(missing)}（实际列：{', '.join(map(str, df.columns))}）"
        return result

    for idx, row in df.iterrows():
        upper = to_number(row["上限"])
        lower = to_number(row["下限"])
        measured = to_number(row["实测值"])
        ok, reason = judge_row(upper, lower, measured)
        result.total += 1
        if ok:
            result.ok_count += 1
        else:
            result.ng_count += 1
            result.row_details.append({"行号": idx + 2, "上限": row["上限"], "下限": row["下限"], "实测值": row["实测值"], "原因": reason})

    result.status = "OK" if result.ng_count == 0 else "NG"
    return result


def build_zip(results: list[FileResult]) -> bytes:
    buf = io.BytesIO()
    used_names: set = set()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for r in results:
            if r.status == "ERROR":
                folder = "无法读取"
                new_name = r.name
            else:
                stem, dot, ext = r.name.rpartition(".")
                new_name = f"{stem}_{r.status}{dot}{ext}" if dot else f"{r.name}_{r.status}"
                folder = r.status
                r.new_name = new_name
            arcname = unique_name(f"{folder}/{new_name}", used_names)
            zf.writestr(arcname, r.data)
    buf.seek(0)
    return buf.getvalue()


# ---------------- 页面 ----------------
st.markdown("""
<div class="qc-masthead">
  <div class="qc-eyebrow">Quality Inspection Console</div>
  <div class="qc-title">质检判定台</div>
  <div class="qc-sub">Excel 测试数据批量判定 · 实测值与上下限自动比对 · OK / NG 分类归档</div>
</div>
<div class="qc-rule-panel">
  <b>判定规则</b> —— 逐行校验：<b>下限 ≤ 实测值 ≤ 上限</b>；
  仅有下限时 <b>实测值 ≥ 下限</b>；仅有上限时 <b>实测值 ≤ 上限</b>。
  上下限均缺失或实测值非数字记为 <b>NG</b> 并注明原因。
  文件内全部数据行合格 → 整份文件判 <b>OK</b>，否则判 <b>NG</b>；
  结果按 OK / NG 文件夹打包下载，文件名自动附加 <b>_OK</b> / <b>_NG</b> 后缀。
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "拖入或选择 Excel 文件（.xlsx / .xls，支持批量）",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

if uploaded_files:
    results: list[FileResult] = []
    for uf in uploaded_files:
        data = uf.getvalue()
        r = process_file(uf)
        r.data = data
        results.append(r)

    ok_files = [r for r in results if r.status == "OK"]
    ng_files = [r for r in results if r.status == "NG"]
    err_files = [r for r in results if r.status == "ERROR"]

    st.markdown(f"""
<div class="qc-metrics">
  <div class="qc-card"><div class="lbl">Files · 文件总数</div><div class="val">{len(results)}</div></div>
  <div class="qc-card ok"><div class="lbl">Pass · OK</div><div class="val">{len(ok_files)}</div></div>
  <div class="qc-card ng"><div class="lbl">Fail · NG</div><div class="val">{len(ng_files)}</div></div>
  <div class="qc-card err"><div class="lbl">Unreadable · 无法读取</div><div class="val">{len(err_files)}</div></div>
</div>
""", unsafe_allow_html=True)

    rows_html = []
    for r in results:
        if r.status == "OK":
            pill = '<span class="pill ok">OK</span>'
        elif r.status == "NG":
            pill = '<span class="pill ng">NG</span>'
        else:
            pill = '<span class="pill err">ERROR</span>'
        note = html.escape(r.message) if r.message else "—"
        rows_html.append(
            f"<tr><td>{html.escape(r.name)}</td><td>{pill}</td>"
            f'<td class="num">{r.total}</td><td class="num">{r.ok_count}</td>'
            f'<td class="num">{r.ng_count}</td><td>{note}</td></tr>'
        )
    st.markdown(f"""
<table class="qc-table">
  <thead><tr><th>文件名</th><th>判定</th><th>数据行</th><th>OK</th><th>NG</th><th>备注</th></tr></thead>
  <tbody>{''.join(rows_html)}</tbody>
</table>
""", unsafe_allow_html=True)

    for r in results:
        if r.status == "NG" and r.row_details:
            with st.expander(f"▸ {r.name} — NG 明细（{r.ng_count} 行）"):
                st.dataframe(pd.DataFrame(r.row_details), use_container_width=True, hide_index=True)
        elif r.status == "ERROR":
            with st.expander(f"▸ {r.name} — 处理失败"):
                st.error(r.message)

    st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)
    st.download_button(
        label="▼ 下载分类结果 · ZIP（OK / NG 文件夹）",
        data=build_zip(results),
        file_name="判定结果.zip",
        mime="application/zip",
    )
else:
    st.markdown("""
<div class="qc-rule-panel" style="border-left-color:var(--ink);text-align:center;color:var(--ink-soft);">
  等待上料 —— 请上传一个或多个 Excel 文件开始判定
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="qc-foot">QC CONSOLE · v1.1 · 下限 ≤ 实测值 ≤ 上限</div>', unsafe_allow_html=True)
