# -*- coding: utf-8 -*-
"""Excel 测试数据批量判定工具

判定规则：
- 上限、下限都有值：下限 <= 实测值 <= 上限 → OK，否则 NG
- 只有下限：实测值 >= 下限 → OK，否则 NG
- 只有上限：实测值 <= 上限 → OK，否则 NG
- 上下限都无值：无法判定 → NG（可在页面中查看原因）
- 实测值无法转为数字 → NG
文件内所有数据行均为 OK，则整个文件判定为 OK，否则为 NG。
"""

from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass, field

import pandas as pd
import streamlit as st

REQUIRED_COLUMNS = ["上限", "下限", "实测值"]

st.set_page_config(page_title="Excel 测试数据判定工具", page_icon="📊", layout="wide")


@dataclass
class FileResult:
    name: str
    status: str  # "OK" / "NG" / "ERROR"
    total: int = 0
    ok_count: int = 0
    ng_count: int = 0
    message: str = ""
    row_details: list[dict] = field(default_factory=list)
    new_name: str = ""
    data: bytes = b""


def to_number(value):
    """把单元格内容转成 float，失败返回 None。支持带空格/逗号的数字。"""
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
    """返回 (是否OK, 原因说明)"""
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
    # 只有上限
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
    except Exception as exc:  # 读取失败
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
st.title("📊 Excel 测试数据批量判定工具")
st.markdown(
    """
**判定规则**：逐行判断「实测值」是否落在「下限 ≤ 实测值 ≤ 上限」区间内；
只有下限时要求「实测值 ≥ 下限」，只有上限时要求「实测值 ≤ 上限」。
全部数据行合格 → 文件判定为 **OK**，否则为 **NG**。
处理后按 OK / NG 文件夹分类打包下载，文件名会自动加上 `_OK` / `_NG` 后缀。
"""
)

uploaded_files = st.file_uploader(
    "批量上传 Excel 文件（.xlsx / .xls）",
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

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("文件总数", len(results))
    c2.metric("✅ OK 文件", len(ok_files))
    c3.metric("❌ NG 文件", len(ng_files))
    c4.metric("⚠️ 无法读取", len(err_files))

    st.subheader("判定结果汇总")
    summary_rows = []
    for r in results:
        summary_rows.append({
            "文件名": r.name,
            "判定": r.status,
            "数据行数": r.total,
            "OK 行数": r.ok_count,
            "NG 行数": r.ng_count,
            "备注": r.message,
        })
    summary_df = pd.DataFrame(summary_rows)

    def color_status(val):
        if val == "OK":
            return "color: green; font-weight: bold"
        if val == "NG":
            return "color: red; font-weight: bold"
        return "color: orange; font-weight: bold"

    st.dataframe(summary_df.style.map(color_status, subset=["判定"]), use_container_width=True, hide_index=True)

    for r in results:
        if r.status == "NG" and r.row_details:
            with st.expander(f"❌ {r.name} — NG 明细（{r.ng_count} 行）"):
                st.dataframe(pd.DataFrame(r.row_details), use_container_width=True, hide_index=True)
        elif r.status == "ERROR":
            with st.expander(f"⚠️ {r.name} — 处理失败"):
                st.error(r.message)

    zip_bytes = build_zip(results)
    st.download_button(
        label="📦 下载分类结果（ZIP：OK / NG 文件夹）",
        data=zip_bytes,
        file_name="判定结果.zip",
        mime="application/zip",
        type="primary",
    )
else:
    st.info("请上传一个或多个 Excel 文件开始判定。")
