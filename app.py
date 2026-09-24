#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
亚马逊外箱面单自动化处理 - 极简高颜值 Web 版
部署环境：streamlit.io (Streamlit Cloud)
"""

import os
import sys
import io
import re
import warnings
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple
from collections import defaultdict

import streamlit as st
import pdfplumber
try:
    from PyPDF2 import PdfWriter, PdfReader
except ImportError:
    from pypdf import PdfWriter, PdfReader

from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import pandas as pd

# ==============================================================================
# 0. 极简页面与字体初始化
# ==============================================================================
st.set_page_config(
    page_title="外箱面单处理",
    page_icon="📦",
    layout="centered"
)

# 注入极简美化样式
st.markdown("""
<style>
/* 隐藏 Streamlit 默认顶部及底部杂物 */
#MainMenu {visibility: hidden;}
header {visibility: hidden;}
footer {visibility: hidden;}

/* 容器排版内边距与居中宽度限制 */
.block-container {
    padding-top: 2.2rem;
    padding-bottom: 2rem;
    max-width: 680px;
}

/* 极简状态胶囊样式 */
.status-pill {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    padding: 5px 14px;
    border-radius: 20px;
    font-size: 0.82rem;
    background: rgba(255, 255, 255, 0.05);
    border: 1px solid rgba(255, 255, 255, 0.12);
    color: #94a3b8;
    margin-top: 2px;
    margin-bottom: 12px;
}
.dot {
    width: 7px;
    height: 7px;
    border-radius: 50%;
    display: inline-block;
}
.dot-green { background-color: #10b981; box-shadow: 0 0 6px #10b981; }
.dot-red { background-color: #ef4444; box-shadow: 0 0 6px #ef4444; }

/* 换表格小按钮样式微调 */
div[data-testid="stPopover"] > button {
    border-radius: 16px;
    padding: 2px 10px;
    font-size: 0.8rem;
    height: 30px;
}
</style>
""", unsafe_allow_html=True)

DEFAULT_FONT = "Helvetica"
FONT_SEARCH_PATHS = [
    # Linux (Streamlit Cloud Debian)
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    # 本地仓库自带
    "simhei.ttf",
    "wqy-microhei.ttc",
    # Windows
    r"C:\Windows\Fonts\simhei.ttf",
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\simsun.ttc",
    "/System/Library/Fonts/PingFang.ttc",
]

CHINESE_FONT_REGISTERED = False
for fpath in FONT_SEARCH_PATHS:
    if os.path.exists(fpath):
        try:
            pdfmetrics.registerFont(TTFont("ChineseFont", fpath))
            DEFAULT_FONT = "ChineseFont"
            CHINESE_FONT_REGISTERED = True
            break
        except Exception:
            pass

if not CHINESE_FONT_REGISTERED:
    try:
        pdfmetrics.registerFont(TTFont("ChineseFont", "simhei.ttf"))
        DEFAULT_FONT = "ChineseFont"
    except Exception:
        DEFAULT_FONT = "Helvetica"


# ==============================================================================
# 1. 自动定位最新 Commodities 表格
# ==============================================================================

def find_latest_commodities_file(directory: str = ".") -> Optional[Dict[str, Any]]:
    valid_exts = (".xlsx", ".xls", ".csv")
    candidates = []

    if not os.path.exists(directory):
        return None

    for fname in os.listdir(directory):
        if fname.startswith("~$"):
            continue
        if any(fname.lower().endswith(ext) for ext in valid_exts):
            if "commodit" in fname.lower():
                full_path = os.path.join(directory, fname)
                mtime = os.path.getmtime(full_path)

                date_matches = re.findall(
                    r'\b20\d{2}[-_]?(?:0[1-9]|1[0-2])[-_]?(?:0[1-9]|[12]\d|3[01])\b|\b20\d{6}\b',
                    fname
                )
                date_score = 0
                if date_matches:
                    clean_date = re.sub(r'[-_]', '', date_matches[-1])
                    try:
                        date_score = int(clean_date)
                    except ValueError:
                        date_score = 0

                candidates.append({
                    "path": full_path,
                    "filename": fname,
                    "date_score": date_score,
                    "mtime": mtime,
                })

    if not candidates:
        return None

    candidates.sort(key=lambda x: (x["date_score"], x["mtime"]), reverse=True)
    return candidates[0]


def load_commodities_df(source: Any) -> Optional[pd.DataFrame]:
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            if isinstance(source, str):
                return pd.read_csv(source) if source.lower().endswith(".csv") else pd.read_excel(source)
            else:
                is_csv = getattr(source, "name", "").lower().endswith(".csv")
                return pd.read_csv(source) if is_csv else pd.read_excel(source)
    except Exception:
        return None


def get_sku_info_from_df(sku: str, df: Optional[pd.DataFrame]) -> dict:
    if df is None:
        return {"product": "未找到商品库", "brand": "请检查表格"}

    col_sku = next((c for c in df.columns if str(c).strip().upper() == "SKU"), None)
    col_product = next((c for c in df.columns if "品名" in str(c) or "商品" in str(c)), None)
    col_brand = next((c for c in df.columns if "品牌" in str(c) or "工厂" in str(c)), None)

    if not col_sku or not col_product or not col_brand:
        return {"product": "表格列缺失", "brand": "缺少SKU/品名/品牌"}

    match = df[df[col_sku].astype(str).str.strip() == sku.strip()]
    if match.empty:
        return {"product": "未匹配到SKU", "brand": "请更新商品库"}

    product = str(match[col_product].values[0]) if pd.notna(match[col_product].values[0]) else ""
    brand = str(match[col_brand].values[0]) if pd.notna(match[col_brand].values[0]) else ""
    return {"product": product, "brand": brand}


def extract_sku_from_text(text: str) -> str:
    lines = text.split("\n")
    for line in lines:
        match = re.search(r"SKU\s*[:：]\s*(\S+)", line, re.IGNORECASE)
        if match:
            return match.group(1).strip()

    for i, line in enumerate(lines):
        if "Single SKU" in line and i + 1 < len(lines):
            match = re.search(r"([\w-]+)", lines[i + 1])
            if match:
                return match.group(1).strip()
    return "未知SKU"


def extract_warehouse_from_text(text: str) -> str:
    lines = text.split("\n")
    for line in lines:
        match = re.search(r"FBA STA \(.*\)-([A-Z0-9]{3,5})\b", line)
        if match:
            return match.group(1)
        match = re.search(r"(?<!\d)-([A-Z]{2,}[0-9]{1,3})\b(?!-\d{4})", line)
        if match:
            return match.group(1)
    return "未知仓库"


def add_sku_label_page(
    writer: PdfWriter,
    sku: str,
    count: int,
    warehouse: Optional[str] = None,
    sku_info: Optional[dict] = None,
    readers_cache: Optional[list] = None
) -> None:
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(283.46, 283.46))

    product_name = sku_info.get("product", "") if sku_info else ""
    brand_name = sku_info.get("brand", "") if sku_info else ""

    font_name = DEFAULT_FONT
    font_size = 12
    c.setFont(font_name, font_size)

    x_pos = 20
    y_pos = 180

    c.drawString(x_pos, y_pos, f"SKU: {sku}")

    if warehouse:
        y_pos -= 28
        c.drawString(x_pos, y_pos, f"仓库: {warehouse}")

    sub_font_size = max(10, font_size - 2)
    c.setFont(font_name, sub_font_size)

    if product_name:
        y_pos -= 26
        c.drawString(x_pos, y_pos, f"品名: {product_name[:20]}")

    if brand_name:
        y_pos -= 26
        c.drawString(x_pos, y_pos, f"工厂: {brand_name[:20]}")

    if count is not None:
        y_pos -= 26
        c.drawString(x_pos, y_pos, f"数量: 共 {count} 箱")

    c.save()
    packet.seek(0)
    reader = PdfReader(packet)
    if readers_cache is not None:
        readers_cache.append(reader)
    writer.add_page(reader.pages[0])


def process_pdf_in_memory(
    pdf_file_bytes: bytes,
    commodities_df: Optional[pd.DataFrame],
    progress_callback=None
) -> Tuple[bytes, Dict[str, Any], List[Dict[str, Any]]]:
    sku_counts = defaultdict(int)
    sku_pages = defaultdict(list)
    warehouse_info = defaultdict(str)

    # 1. 页面分析
    with pdfplumber.open(io.BytesIO(pdf_file_bytes)) as pdf:
        total_pages = len(pdf.pages)
        for i, page in enumerate(pdf.pages):
            if progress_callback:
                progress_callback(i + 1, total_pages, "解析面单中...")
            text = page.extract_text() or ""
            if not text:
                continue
            sku = extract_sku_from_text(text)
            if not sku or sku == "未知SKU":
                continue
            warehouse = extract_warehouse_from_text(text)
            if warehouse != "未知仓库":
                warehouse_info[sku] = warehouse
            sku_counts[sku] += 1
            sku_pages[sku].append(i)

    # 2. 面单重组
    writer = PdfWriter()
    original_reader = PdfReader(io.BytesIO(pdf_file_bytes))
    readers_cache = [original_reader]

    total_skus = len(sku_pages)
    current_sku_idx = 0
    table_data = []

    for sku, pages in sku_pages.items():
        current_sku_idx += 1
        if progress_callback:
            progress_callback(current_sku_idx, total_skus, f"重构排版中 ({current_sku_idx}/{total_skus})...")

        count = sku_counts[sku]
        warehouse = warehouse_info.get(sku, None)
        info = get_sku_info_from_df(sku, commodities_df)

        table_data.append({
            "SKU": sku,
            "品名": info["product"],
            "工厂/品牌": info["brand"],
            "数量": f"{count} 箱"
        })

        add_sku_label_page(writer, sku, count, warehouse, info, readers_cache=readers_cache)
        for page_num in pages:
            writer.add_page(original_reader.pages[page_num])
            writer.add_page(original_reader.pages[page_num])
        add_sku_label_page(writer, sku, count, warehouse, info, readers_cache=readers_cache)

    out_buf = io.BytesIO()
    writer.write(out_buf)
    out_buf.seek(0)

    summary = {
        "total_original_pages": total_pages,
        "sku_types_count": len(sku_counts),
        "total_output_pages": len(writer.pages)
    }

    return out_buf.getvalue(), summary, table_data


# ==============================================================================
# 2. 极简页面渲染
# ==============================================================================

def main():
    # 顶部标题与右上角轻量按钮
    col_title, col_opt = st.columns([3.8, 1.2], vertical_alignment="center")
    
    with col_title:
        st.subheader("📦 亚马逊外箱面单处理")

    # 自动检索最新的 Commodities 表格
    auto_commodities = find_latest_commodities_file(".")
    active_df = None

    # 极简气泡弹窗（Popover），平时完全隐藏，点击才展开上传
    custom_file = None
    with col_opt:
        try:
            with st.popover("⚙️ 换表格"):
                custom_file = st.file_uploader(
                    "更换商品列表",
                    type=["xlsx", "xls", "csv"],
                    label_visibility="collapsed",
                    key="pop_uploader"
                )
        except AttributeError:
            with st.expander("⚙️ 换表格"):
                custom_file = st.file_uploader(
                    "更换商品列表",
                    type=["xlsx", "xls", "csv"],
                    label_visibility="collapsed",
                    key="exp_uploader"
                )

    # 状态指示胶囊
    if custom_file is not None:
        active_df = load_commodities_df(custom_file)
        table_label = f"自定义: {custom_file.name}"
        st.markdown(f'<div class="status-pill"><span class="dot dot-green"></span><span>{table_label}</span></div>', unsafe_allow_html=True)
    elif auto_commodities:
        active_df = load_commodities_df(auto_commodities["path"])
        table_label = auto_commodities["filename"]
        st.markdown(f'<div class="status-pill"><span class="dot dot-green"></span><span>{table_label}</span></div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="status-pill"><span class="dot dot-red"></span><span>未检测到商品库 (点击右侧换表格)</span></div>', unsafe_allow_html=True)

    # 主操作区：面单拖拽上传
    uploaded_pdf = st.file_uploader(
        "拖拽或点击上传亚马逊面单 PDF",
        type=["pdf"],
        label_visibility="collapsed",
        key="main_pdf"
    )

    if uploaded_pdf is not None:
        if st.button("开始处理", type="primary", use_container_width=True):
            p_bar = st.progress(0)
            status_txt = st.empty()

            def on_progress(cur, tot, msg):
                pct = int(cur / tot * 100) if tot > 0 else 0
                p_bar.progress(min(pct, 100))
                status_txt.caption(f"{pct}% · {msg}")

            try:
                res_bytes, summary, table_data = process_pdf_in_memory(
                    uploaded_pdf.getvalue(),
                    active_df,
                    progress_callback=on_progress
                )
                p_bar.progress(100)
                status_txt.empty()

                # 下载区域
                out_name = f"{os.path.splitext(uploaded_pdf.name)[0]}-优化.pdf"
                st.download_button(
                    label=f"📥 下载优化面单 ({summary['total_output_pages']} 页)",
                    data=res_bytes,
                    file_name=out_name,
                    mime="application/pdf",
                    type="primary",
                    use_container_width=True
                )

                # 简洁指标卡
                c1, c2, c3 = st.columns(3)
                c1.metric("原箱数", f"{summary['total_original_pages']}")
                c2.metric("SKU 数", f"{summary['sku_types_count']}")
                c3.metric("总页数", f"{summary['total_output_pages']}")

                # 简洁明细表
                if table_data:
                    st.dataframe(pd.DataFrame(table_data), use_container_width=True, hide_index=True)

            except Exception as e:
                st.error(f"处理失败: {e}")


if __name__ == "__main__":
    main()