#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
亚马逊外箱面单处理工具 - Streamlit Web 版
部署目标：streamlit.io
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
# 0. 页面基本配置与字体初始化
# ==============================================================================
st.set_page_config(
    page_title="亚马逊外箱面单优化工具",
    page_icon="📦",
    layout="centered"
)

DEFAULT_FONT = "Helvetica"
FONT_SEARCH_PATHS = [
    # 优先 Linux (Streamlit Cloud Debian) 系统字体
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    # 本地仓库自带字体
    "simhei.ttf",
    "wqy-microhei.ttc",
    # Windows 常见路径 (本地测试用)
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
# 1. 核心业务逻辑：自动检索最新 Commodities 表格
# ==============================================================================

def find_latest_commodities_file(directory: str = ".") -> Optional[Dict[str, Any]]:
    """
    在指定目录中搜索文件名包含 'Commodities' 的表格文件，
    若存在多份，优先按文件名提取的时间日期排序，其次按文件修改时间排序。
    """
    valid_exts = (".xlsx", ".xls", ".csv")
    candidates = []

    if not os.path.exists(directory):
        return None

    for fname in os.listdir(directory):
        if fname.startswith("~$"):  # 忽略临时打开的 Excel 文件
            continue
        if any(fname.lower().endswith(ext) for ext in valid_exts):
            if "commodit" in fname.lower():  # 容错匹配 Commodities / Commoditie
                full_path = os.path.join(directory, fname)
                mtime = os.path.getmtime(full_path)

                # 正则提取文件名中包含的日期格式 (如 2026-09-23, 20260923, 2026_09_23)
                date_matches = re.findall(
                    r'\b20\d{2}[-_]?(?:0[1-9]|1[0-2])[-_]?(?:0[1-9]|[12]\d|3[01])\b|\b20\d{6}\b',
                    fname
                )
                date_score = 0
                date_str = ""
                if date_matches:
                    clean_date = re.sub(r'[-_]', '', date_matches[-1])
                    try:
                        date_score = int(clean_date)
                        date_str = f"{clean_date[:4]}-{clean_date[4:6]}-{clean_date[6:8]}"
                    except ValueError:
                        date_score = 0

                candidates.append({
                    "path": full_path,
                    "filename": fname,
                    "date_score": date_score,
                    "date_str": date_str,
                    "mtime": mtime,
                    "mtime_str": datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
                })

    if not candidates:
        return None

    # 排序：优先按提取出来的日期数字降序，其次按系统修改时间降序
    candidates.sort(key=lambda x: (x["date_score"], x["mtime"]), reverse=True)
    best = candidates[0]
    reason = f"识别到文件名最新日期: {best['date_str']}" if best["date_score"] > 0 else f"文件最新修改时间: {best['mtime_str']}"
    best["reason"] = reason
    best["all_candidates"] = candidates
    return best


def load_commodities_df(excel_source: Any) -> Optional[pd.DataFrame]:
    """读取 Commodities 数据表 (支持路径或 Streamlit 上传的内存文件)"""
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            if isinstance(excel_source, str):
                if excel_source.lower().endswith(".csv"):
                    df = pd.read_csv(excel_source)
                else:
                    df = pd.read_excel(excel_source)
            else:
                # 内存文件对象
                if getattr(excel_source, "name", "").lower().endswith(".csv"):
                    df = pd.read_csv(excel_source)
                else:
                    df = pd.read_excel(excel_source)
            return df
    except Exception:
        return None


def get_sku_info_from_df(sku: str, df: Optional[pd.DataFrame]) -> dict:
    """从给定的 DataFrame 匹配品名与工厂/品牌"""
    if df is None:
        return {"product": "未找到商品列表", "brand": "请检查表格文件"}

    col_sku = next((c for c in df.columns if str(c).strip().upper() == "SKU"), None)
    col_product = next((c for c in df.columns if "品名" in str(c) or "商品" in str(c)), None)
    col_brand = next((c for c in df.columns if "品牌" in str(c) or "工厂" in str(c)), None)

    if not col_sku or not col_product or not col_brand:
        return {"product": "表格格式错误", "brand": "缺少SKU/品名/品牌列"}

    match = df[df[col_sku].astype(str).str.strip() == sku.strip()]
    if match.empty:
        return {"product": "未匹配到SKU", "brand": "请到塞狐下载最新商品列表"}

    product = str(match[col_product].values[0]) if pd.notna(match[col_product].values[0]) else ""
    brand = str(match[col_brand].values[0]) if pd.notna(match[col_brand].values[0]) else ""
    return {"product": product, "brand": brand}


def extract_sku_from_text(text: str) -> str:
    """从面单页面文本中智能提取 SKU 编码"""
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
    """从面单文本中提取 FBA 目标仓库代码"""
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
    """生成 10*10cm (283.46x283.46 pt) 隔页标注卡"""
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
    """在内存中完整处理 PDF 并返回生成文件的二进制字节流"""
    sku_counts = defaultdict(int)
    sku_pages = defaultdict(list)
    warehouse_info = defaultdict(str)

    # 1. 解析阶段
    with pdfplumber.open(io.BytesIO(pdf_file_bytes)) as pdf:
        total_pages = len(pdf.pages)
        for i, page in enumerate(pdf.pages):
            if progress_callback:
                progress_callback(i + 1, total_pages, "正在扫描解析页面...")
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

    # 2. 重构生成阶段
    writer = PdfWriter()
    original_reader = PdfReader(io.BytesIO(pdf_file_bytes))
    readers_cache = [original_reader]

    total_skus = len(sku_pages)
    current_sku_idx = 0

    table_data = []

    for sku, pages in sku_pages.items():
        current_sku_idx += 1
        if progress_callback:
            progress_callback(current_sku_idx, total_skus, f"正在生成 SKU 隔页卡与双份面单 ({current_sku_idx}/{total_skus})...")

        count = sku_counts[sku]
        warehouse = warehouse_info.get(sku, None)
        info = get_sku_info_from_df(sku, commodities_df)

        table_data.append({
            "SKU": sku,
            "品名": info["product"],
            "工厂/品牌": info["brand"],
            "箱数": f"{count} 箱"
        })

        # 前置标牌卡
        add_sku_label_page(writer, sku, count, warehouse, info, readers_cache=readers_cache)
        # 双份原面单
        for page_num in pages:
            writer.add_page(original_reader.pages[page_num])
            writer.add_page(original_reader.pages[page_num])
        # 后置标牌卡
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
# 2. Streamlit 页面渲染与交互
# ==============================================================================

def main():
    st.title("📦 亚马逊外箱面单自动化处理")
    st.markdown("自动解析 SKU 并插入 **10×10 cm 隔页标注卡**，原面单**自动打印双份**。")

    # 检查仓库中自动匹配的 Commodities 表格
    commodities_info = find_latest_commodities_file(".")
    active_df = None

    with st.container():
        st.subheader("1. 商品列表 (Commodities 表格)")
        if commodities_info:
            st.success(f"自动检测到最新商品表格：**`{commodities_info['filename']}`**")
            st.caption(f"判定规则：{commodities_info['reason']}")
            active_df = load_commodities_df(commodities_info["path"])

            if len(commodities_info["all_candidates"]) > 1:
                with st.expander(f"查看同目录下检测到的所有候选表格 ({len(commodities_info['all_candidates'])} 份)"):
                    for c in commodities_info["all_candidates"]:
                        is_selected = " (当前选中)" if c["path"] == commodities_info["path"] else ""
                        st.text(f"• {c['filename']} - 修改时间: {c['mtime_str']}{is_selected}")
        else:
            st.warning("⚠️ 仓库同目录下未找到包含 `Commodities` 的表格文件，请手动上传。")

        # 备选：手动上传商品列表覆盖
        uploaded_excel = st.file_uploader(
            "或者手动上传商品列表 (将优先使用上传的表格)",
            type=["xlsx", "xls", "csv"],
            key="excel_uploader"
        )
        if uploaded_excel is not None:
            active_df = load_commodities_df(uploaded_excel)
            st.info(f"已切换为手动上传的表格：`{uploaded_excel.name}`")

    st.divider()

    # 上传 PDF 面单
    st.subheader("2. 上传外箱面单 PDF")
    uploaded_pdf = st.file_uploader("选择亚马逊原版外箱面单 PDF 文件", type=["pdf"], key="pdf_uploader")

    if uploaded_pdf is not None:
        file_size_mb = len(uploaded_pdf.getvalue()) / (1024 * 1024)
        st.write(f"📄 文件名: **{uploaded_pdf.name}** ({file_size_mb:.2f} MB)")

        if st.button("🚀 开始一键优化处理", type="primary", use_container_width=True):
            progress_bar = st.progress(0)
            status_text = st.empty()

            def update_progress(current, total, message):
                percent = int(current / total * 100) if total > 0 else 0
                progress_bar.progress(min(percent, 100))
                status_text.text(f"[{percent}%] {message}")

            try:
                processed_bytes, summary, table_data = process_pdf_in_memory(
                    uploaded_pdf.getvalue(),
                    active_df,
                    progress_callback=update_progress
                )
                progress_bar.progress(100)
                status_text.text(" 处理完成！")

                st.balloons()
                st.success("🎉 面单优化重组完成！")

                # 指标展示卡
                col1, col2, col3 = st.columns(3)
                col1.metric("原始面单数", f"{summary['total_original_pages']} 箱/页")
                col2.metric("识别 SKU 种类", f"{summary['sku_types_count']} 种")
                col3.metric("生成新 PDF 总页数", f"{summary['total_output_pages']} 页")

                # 下载按钮
                base_name = os.path.splitext(uploaded_pdf.name)[0]
                download_filename = f"{base_name}-优化.pdf"

                st.download_button(
                    label=f" 立即下载优化后的面单 ({download_filename})",
                    data=processed_bytes,
                    file_name=download_filename,
                    mime="application/pdf",
                    type="primary",
                    use_container_width=True
                )

                # SKU 统计明细表
                st.subheader("📊 SKU 箱数明细汇总")
                if table_data:
                    df_display = pd.DataFrame(table_data)
                    st.dataframe(df_display, use_container_width=True, hide_index=True)

            except Exception as e:
                st.error(f"❌ 处理过程中出现异常：{e}")
                import traceback
                st.code(traceback.format_exc())


if __name__ == "__main__":
    main()