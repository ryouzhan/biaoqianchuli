#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
亚马逊外箱面单自动化处理 - 批量处理 + 双引擎自适应极简版
依赖：
  基础依赖：pip install pypdf reportlab pandas openpyxl streamlit
  可选加速：pip install pymupdf (若安装则自动启用极速模式)
"""

import os
import io
import re
import zipfile
import warnings
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple
from collections import defaultdict

import streamlit as st
import pandas as pd

# 检查是否存在 PyMuPDF (fitz)
USE_FITZ = False
try:
    import pymupdf as fitz
    USE_FITZ = True
except ImportError:
    try:
        import fitz
        USE_FITZ = True
    except ImportError:
        USE_FITZ = False

# 基础纯 Python PDF 库兜底
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ==============================================================================
# 0. 极简样式与字体加载
# ==============================================================================
st.set_page_config(
    page_title="外箱面单批量处理",
    page_icon="??",
    layout="centered"
)

st.markdown("""
<style>
#MainMenu {visibility: hidden;}
header {visibility: hidden;}
footer {visibility: hidden;}
.block-container {
    padding-top: 2.2rem;
    padding-bottom: 2rem;
    max-width: 720px;
}
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
.dot { width: 7px; height: 7px; border-radius: 50%; display: inline-block; }
.dot-green { background-color: #10b981; box-shadow: 0 0 6px #10b981; }
.dot-red { background-color: #ef4444; box-shadow: 0 0 6px #ef4444; }
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
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "simhei.ttf",
    "wqy-microhei.ttc",
    r"C:\Windows\Fonts\simhei.ttf",
    r"C:\Windows\Fonts\msyh.ttc",
    r"C:\Windows\Fonts\simsun.ttc",
    "/System/Library/Fonts/PingFang.ttc",
]

CHINESE_FONT_PATH = None
for fpath in FONT_SEARCH_PATHS:
    if os.path.exists(fpath):
        CHINESE_FONT_PATH = fpath
        try:
            pdfmetrics.registerFont(TTFont("ChineseFont", fpath))
            DEFAULT_FONT = "ChineseFont"
            break
        except Exception:
            pass


# ==============================================================================
# 1. 商品表处理与正则提取
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
                date_matches = re.findall(r'\b20\d{6,8}\b|\b20\d{2}[-_]\d{2}[-_]\d{2}\b', fname)
                date_score = int(re.sub(r'[-_]', '', date_matches[-1])) if date_matches else 0

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
    col_product = next((c for c in df.columns if any(k in str(c) for k in ("品名", "商品", "名称"))), None)
    col_brand = next((c for c in df.columns if any(k in str(c) for k in ("品牌", "工厂"))), None)

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


# ==============================================================================
# 2. 核心处理引擎（支持 PyMuPDF 极速模式，自动降级纯 pypdf）
# ==============================================================================

def add_sku_label_page_reportlab(
    writer: PdfWriter,
    sku: str,
    count: int,
    warehouse: Optional[str] = None,
    sku_info: Optional[dict] = None
) -> None:
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(283.46, 283.46))

    # 边框装饰线
    c.setStrokeColorRGB(0.75, 0.75, 0.75)
    c.rect(12, 12, 283.46 - 24, 283.46 - 24)

    product_name = sku_info.get("product", "") if sku_info else ""
    brand_name = sku_info.get("brand", "") if sku_info else ""

    font_name = DEFAULT_FONT
    c.setFont(font_name, 13)

    x_pos = 24
    y_pos = 220
    c.drawString(x_pos, y_pos, f"SKU: {sku}")

    c.setFont(font_name, 11)
    if warehouse:
        y_pos -= 32
        c.drawString(x_pos, y_pos, f"仓库: {warehouse}")

    c.setFont(font_name, 10)
    if product_name:
        y_pos -= 30
        c.drawString(x_pos, y_pos, f"品名: {product_name[:18]}")

    if brand_name:
        y_pos -= 30
        c.drawString(x_pos, y_pos, f"工厂: {brand_name[:18]}")

    if count is not None:
        y_pos -= 30
        c.setFont(font_name, 11)
        c.drawString(x_pos, y_pos, f"数量: 共 {count} 箱")

    c.save()
    packet.seek(0)
    reader = PdfReader(packet)
    writer.add_page(reader.pages[0])


def process_single_pdf_bytes(
    pdf_file_bytes: bytes,
    commodities_df: Optional[pd.DataFrame]
) -> Tuple[bytes, Dict[str, Any], List[Dict[str, Any]]]:
    """处理单个 PDF 字节流，支持根据环境自适应"""
    sku_counts = defaultdict(int)
    sku_pages = defaultdict(list)
    warehouse_info = defaultdict(str)

    # 优先采用 PyMuPDF (如果可用)
    if USE_FITZ:
        src_doc = fitz.open(stream=pdf_file_bytes, filetype="pdf")
        total_pages = len(src_doc)
        for i in range(total_pages):
            text = src_doc[i].get_text() or ""
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

        out_doc = fitz.open()
        table_data = []
        for sku, pages in sku_pages.items():
            count = sku_counts[sku]
            warehouse = warehouse_info.get(sku, None)
            info = get_sku_info_from_df(sku, commodities_df)
            table_data.append({
                "SKU": sku,
                "品名": info["product"],
                "工厂/品牌": info["brand"],
                "数量": f"{count} 箱",
                "箱数数值": count
            })

            # 用 reportlab 生成分隔页，再由 fitz 读取插入
            sep_writer = PdfWriter()
            add_sku_label_page_reportlab(sep_writer, sku, count, warehouse, info)
            sep_bytes = io.BytesIO()
            sep_writer.write(sep_bytes)
            sep_bytes.seek(0)
            sep_fitz = fitz.open(stream=sep_bytes.getvalue(), filetype="pdf")

            out_doc.insert_pdf(sep_fitz)
            for p in pages:
                out_doc.insert_pdf(src_doc, from_page=p, to_page=p)
                out_doc.insert_pdf(src_doc, from_page=p, to_page=p)
            out_doc.insert_pdf(sep_fitz)
            sep_fitz.close()

        out_bytes = out_doc.tobytes(deflate=True, garbage=3)
        out_page_count = len(out_doc)
        src_doc.close()
        out_doc.close()

    else:
        # 降级模式：纯 pypdf + reportlab（无需安装任何 C 库）
        reader = PdfReader(io.BytesIO(pdf_file_bytes))
        total_pages = len(reader.pages)
        for i, page in enumerate(reader.pages):
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

        writer = PdfWriter()
        table_data = []
        for sku, pages in sku_pages.items():
            count = sku_counts[sku]
            warehouse = warehouse_info.get(sku, None)
            info = get_sku_info_from_df(sku, commodities_df)
            table_data.append({
                "SKU": sku,
                "品名": info["product"],
                "工厂/品牌": info["brand"],
                "数量": f"{count} 箱",
                "箱数数值": count
            })

            add_sku_label_page_reportlab(writer, sku, count, warehouse, info)
            for p in pages:
                writer.add_page(reader.pages[p])
                writer.add_page(reader.pages[p])
            add_sku_label_page_reportlab(writer, sku, count, warehouse, info)

        out_buf = io.BytesIO()
        writer.write(out_buf)
        out_buf.seek(0)
        out_bytes = out_buf.getvalue()
        out_page_count = len(writer.pages)

    summary = {
        "original_pages": total_pages,
        "sku_count": len(sku_counts),
        "output_pages": out_page_count
    }
    return out_bytes, summary, table_data


# ==============================================================================
# 3. Streamlit 交互界面
# ==============================================================================

def main():
    col_title, col_opt = st.columns([3.8, 1.2], vertical_alignment="center")
    with col_title:
        st.subheader("?? 亚马逊外箱面单批量处理")

    auto_commodities = find_latest_commodities_file(".")
    active_df = None
    custom_file = None

    with col_opt:
        try:
            with st.popover("?? 换表格"):
                custom_file = st.file_uploader(
                    "更换商品列表",
                    type=["xlsx", "xls", "csv"],
                    label_visibility="collapsed",
                    key="pop_uploader"
                )
        except AttributeError:
            with st.expander("?? 换表格"):
                custom_file = st.file_uploader(
                    "更换商品列表",
                    type=["xlsx", "xls", "csv"],
                    label_visibility="collapsed",
                    key="exp_uploader"
                )

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

    # 关键改动：开启 accept_multiple_files=True
    uploaded_pdfs = st.file_uploader(
        "拖拽或点击上传一个或多个亚马逊面单 PDF",
        type=["pdf"],
        accept_multiple_files=True,
        label_visibility="collapsed",
        key="main_pdfs"
    )

    if uploaded_pdfs:
        num_files = len(uploaded_pdfs)
        btn_label = f"开始批量处理 ({num_files} 个文件)" if num_files > 1 else "开始处理面单"

        if st.button(btn_label, type="primary", use_container_width=True):
            p_bar = st.progress(0)
            status_txt = st.empty()

            processed_results = []
            all_table_data = []
            total_orig_pages = 0
            total_out_pages = 0

            for idx, pdf_file in enumerate(uploaded_pdfs):
                current_num = idx + 1
                status_txt.caption(f"正在处理 [{current_num}/{num_files}]: {pdf_file.name}")
                p_bar.progress(int((idx / num_files) * 100))

                try:
                    res_bytes, summary, table_data = process_single_pdf_bytes(
                        pdf_file.getvalue(),
                        active_df
                    )
                    base_name = os.path.splitext(pdf_file.name)[0]
                    out_filename = f"{base_name}-优化.pdf"

                    processed_results.append({
                        "filename": out_filename,
                        "bytes": res_bytes,
                        "summary": summary
                    })

                    total_orig_pages += summary["original_pages"]
                    total_out_pages += summary["output_pages"]

                    # 汇总表格带上来源文件名
                    for row in table_data:
                        row_copy = dict(row)
                        row_copy["来源面单"] = pdf_file.name
                        all_table_data.append(row_copy)

                except Exception as e:
                    st.error(f"处理文件 {pdf_file.name} 失败: {e}")

            p_bar.progress(100)
            status_txt.empty()

            if processed_results:
                # 1. 下载区域：单文件直出 PDF，多文件打包成 ZIP
                if num_files == 1:
                    single = processed_results[0]
                    st.download_button(
                        label=f"?? 下载优化面单 ({single['summary']['output_pages']} 页)",
                        data=single["bytes"],
                        file_name=single["filename"],
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )
                else:
                    # 打包为 ZIP
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                        for item in processed_results:
                            zf.writestr(item["filename"], item["bytes"])
                    zip_buffer.seek(0)

                    timestamp = datetime.now().strftime("%m%d_%H%M")
                    zip_name = f"优化面单批量包_{timestamp}.zip"

                    st.download_button(
                        label=f"?? 一键下载全部优化面单 (共 {num_files} 个文件 · ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name=zip_name,
                        mime="application/zip",
                        type="primary",
                        use_container_width=True
                    )

                # 2. 统计指标卡
                distinct_skus = len(set(r["SKU"] for r in all_table_data))
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("文件数", f"{num_files}")
                c2.metric("总原箱数", f"{total_orig_pages}")
                c3.metric("总SKU种类", f"{distinct_skus}")
                c4.metric("总生成页数", f"{total_out_pages}")

                # 3. 汇总数据明细表
                if all_table_data:
                    df_display = pd.DataFrame(all_table_data)
                    # 调整展示列顺序
                    cols = ["来源面单", "SKU", "品名", "工厂/品牌", "数量"]
                    display_cols = [c for c in cols if c in df_display.columns]
                    st.dataframe(df_display[display_cols], use_container_width=True, hide_index=True)

                # 4. 多文件时提供每个文件的单独下载入口（折叠卡片）
                if num_files > 1:
                    with st.expander("?? 点击展开单个文件独立下载"):
                        for item in processed_results:
                            st.download_button(
                                label=f"下载 {item['filename']} ({item['summary']['output_pages']} 页)",
                                data=item["bytes"],
                                file_name=item["filename"],
                                mime="application/pdf",
                                use_container_width=True
                            )


if __name__ == "__main__":
    main()