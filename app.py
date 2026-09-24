#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
亚马逊外箱面单自动化处理 - 胶囊一体化极简版
功能：
1. 自动检测最新商品库，点击胶囊可直接呼出隐藏换表弹窗
2. 支持单文件 / 多文件批量拖拽处理
3. 单文件直接下载优化后的 PDF，多文件自动打包为 ZIP 供一键下载
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

# 双兼容导入 PDF 读写库
try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    from PyPDF2 import PdfReader, PdfWriter

from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ==============================================================================
# 0. 极简页面配置与样式定制
# ==============================================================================
st.set_page_config(
    page_title="外箱面单批量处理",
    page_icon="📦",
    layout="centered"
)

# 注入 CSS：将 st.popover 按钮定制为极简圆角状态胶囊
st.markdown("""
<style>
/* 隐藏 Streamlit 默认顶部与页脚 */
#MainMenu {visibility: hidden;}
header {visibility: hidden;}
footer {visibility: hidden;}

/* 内容居中与宽度约束 */
.block-container {
    padding-top: 2.2rem;
    padding-bottom: 2rem;
    max-width: 720px;
}

/* 核心：将 Popover 按钮伪装成极简圆角状态胶囊 */
div[data-testid="stPopover"] > button {
    border-radius: 20px !important;
    padding: 4px 14px !important;
    font-size: 0.82rem !important;
    background: rgba(255, 255, 255, 0.05) !important;
    border: 1px solid rgba(255, 255, 255, 0.12) !important;
    color: #94a3b8 !important;
    height: auto !important;
    margin-top: 2px !important;
    margin-bottom: 16px !important;
    display: inline-flex !important;
    align-items: center !important;
    transition: all 0.2s ease;
}

/* 鼠标悬停时微亮 */
div[data-testid="stPopover"] > button:hover {
    border-color: rgba(255, 255, 255, 0.28) !important;
    background: rgba(255, 255, 255, 0.09) !important;
    color: #e2e8f0 !important;
}

/* 去除默认 popover 按钮内的多余轮廓 */
div[data-testid="stPopover"] > button:focus {
    box-shadow: none !important;
}
</style>
""", unsafe_allow_html=True)

# 中文字体探测与注册
DEFAULT_FONT = "Helvetica"
FONT_SEARCH_PATHS = [
    # Linux (Streamlit Cloud Debian: packages.txt 中安装 fonts-wqy-microhei)
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    # 本地目录
    "simhei.ttf",
    "wqy-microhei.ttc",
    # Windows
    r"C:\Windows\Fonts\simhei.ttf",
    r"C:\Windows\Fonts\msyh.ttc",
    # macOS
    "/System/Library/Fonts/PingFang.ttc",
]

for fpath in FONT_SEARCH_PATHS:
    if os.path.exists(fpath):
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
                if source.lower().endswith(".csv"):
                    try:
                        return pd.read_csv(source, encoding="utf-8-sig")
                    except Exception:
                        return pd.read_csv(source, encoding="gbk")
                return pd.read_excel(source)
            else:
                is_csv = getattr(source, "name", "").lower().endswith(".csv")
                if is_csv:
                    try:
                        return pd.read_csv(source, encoding="utf-8-sig")
                    except Exception:
                        return pd.read_csv(source, encoding="gbk")
                return pd.read_excel(source)
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
# 2. 面单分隔页绘制与排版逻辑
# ==============================================================================

def add_sku_label_page(
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
    sku_counts = defaultdict(int)
    sku_pages = defaultdict(list)
    warehouse_info = defaultdict(str)

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
            "数量": f"{count} 箱"
        })

        add_sku_label_page(writer, sku, count, warehouse, info)
        for p in pages:
            writer.add_page(reader.pages[p])
            writer.add_page(reader.pages[p])
        add_sku_label_page(writer, sku, count, warehouse, info)

    out_buf = io.BytesIO()
    writer.write(out_buf)
    out_buf.seek(0)

    summary = {
        "original_pages": total_pages,
        "sku_count": len(sku_counts),
        "output_pages": len(writer.pages)
    }
    return out_buf.getvalue(), summary, table_data


# ==============================================================================
# 3. Streamlit 主页面
# ==============================================================================

def main():
    # 顶部标题
    st.subheader("📦 亚马逊外箱面单批量处理")

    # 1. 自动检索本地商品表
    auto_commodities = find_latest_commodities_file(".")

    # 2. 计算胶囊展示文案
    custom_uploaded = st.session_state.get("custom_commodities", None)
    if custom_uploaded is not None:
        pill_label = f"🟢 自定义: {custom_uploaded.name} ▾"
    elif auto_commodities:
        pill_label = f"🟢 {auto_commodities['filename']} ▾"
    else:
        pill_label = "🔴 未检测到商品库 (点击上传) ▾"

    # 3. 核心设计：将状态胶囊本身作为 Popover 弹窗入口
    with st.popover(pill_label):
        st.caption("如需临时覆盖或更换商品库，请在此上传：")
        custom_file = st.file_uploader(
            "上传替代商品列表",
            type=["xlsx", "xls", "csv"],
            label_visibility="collapsed",
            key="custom_commodities"
        )
        if custom_uploaded is not None and st.button("恢复使用默认商品库", use_container_width=True):
            del st.session_state["custom_commodities"]
            st.rerun()

    # 4. 确定当前生效的商品库
    active_df = None
    if custom_file is not None:
        active_df = load_commodities_df(custom_file)
    elif auto_commodities:
        active_df = load_commodities_df(auto_commodities["path"])

    # 5. 核心操作区：批量面单拖拽上传
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
                status_txt.caption(f"正在处理 [{idx+1}/{num_files}]: {pdf_file.name}")
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

                    for row in table_data:
                        row_copy = dict(row)
                        row_copy["来源面单"] = pdf_file.name
                        all_table_data.append(row_copy)

                except Exception as e:
                    st.error(f"处理文件 {pdf_file.name} 失败: {e}")

            p_bar.progress(100)
            status_txt.empty()

            if processed_results:
                # 单文件直出 PDF，多文件打包为 ZIP
                if num_files == 1:
                    single = processed_results[0]
                    st.download_button(
                        label=f"📥 下载优化面单 ({single['summary']['output_pages']} 页)",
                        data=single["bytes"],
                        file_name=single["filename"],
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )
                else:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                        for item in processed_results:
                            zf.writestr(item["filename"], item["bytes"])
                    zip_buffer.seek(0)

                    timestamp = datetime.now().strftime("%m%d_%H%M")
                    zip_name = f"优化面单批量包_{timestamp}.zip"

                    st.download_button(
                        label=f"📦 一键下载全部优化面单 (共 {num_files} 个文件 · ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name=zip_name,
                        mime="application/zip",
                        type="primary",
                        use_container_width=True
                    )

                # 全局汇总指标卡
                distinct_skus = len(set(r["SKU"] for r in all_table_data))
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("文件数", f"{num_files}")
                c2.metric("总原箱数", f"{total_orig_pages}")
                c3.metric("总SKU数", f"{distinct_skus}")
                c4.metric("总生成页数", f"{total_out_pages}")

                # 汇总明细表格
                if all_table_data:
                    df_display = pd.DataFrame(all_table_data)
                    cols = ["来源面单", "SKU", "品名", "工厂/品牌", "数量"]
                    display_cols = [c for c in cols if c in df_display.columns]
                    st.dataframe(df_display[display_cols], use_container_width=True, hide_index=True)

                # 多文件时支持展开下载单个面单
                if num_files > 1:
                    with st.expander("📄 展开单独下载某个面单"):
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