#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
亚马逊外箱面单自动化处理 - 密文云端同步 + 极简纯净版
"""

import os
import io
import re
import json
import zipfile
import warnings
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple
from collections import defaultdict

import streamlit as st
import pandas as pd
from cryptography.fernet import Fernet

# 双兼容导入 PDF 读写库
try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    from PyPDF2 import PdfReader, PdfWriter

from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ==============================================================================
# 0. 极简样式与字体加载
# ==============================================================================
st.set_page_config(
    page_title="外箱面单批量处理",
    page_icon="📦",
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

div[data-testid="stPopover"] > button:hover {
    border-color: rgba(255, 255, 255, 0.28) !important;
    background: rgba(255, 255, 255, 0.09) !important;
    color: #e2e8f0 !important;
}

div[data-testid="stPopover"] > button:focus {
    box-shadow: none !important;
}

div[data-testid="stPopoverBody"] {
    min-width: 380px !important;
    max-width: 440px !important;
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
# 1. 特殊项映射字典管理
# ==============================================================================
MAPPING_FILE = "sku_mapping.json"

def load_sku_mapping() -> Dict[str, str]:
    if os.path.exists(MAPPING_FILE):
        try:
            with open(MAPPING_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_sku_mapping(mapping: Dict[str, str]) -> None:
    try:
        with open(MAPPING_FILE, "w", encoding="utf-8") as f:
            json.dump(mapping, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


# ==============================================================================
# 2. 自动解密与商品表加载逻辑
# ==============================================================================

def get_secret_key() -> Optional[bytes]:
    """获取解密密钥：优先 Streamlit Secrets，其次本地 secret.key"""
    if hasattr(st, "secrets") and "COMMODITIES_KEY" in st.secrets:
        return st.secrets["COMMODITIES_KEY"].encode()
    if os.path.exists("secret.key"):
        try:
            with open("secret.key", "rb") as f:
                return f.read().strip()
        except Exception:
            pass
    return None


def parse_raw_table_bytes(raw_bytes: bytes) -> Optional[pd.DataFrame]:
    """从二进制字节流自动识别 Excel 或 CSV 并转为 DataFrame"""
    try:
        # Excel .xlsx 标准 zip 头魔数
        if raw_bytes.startswith(b"PK\x03\x04") or raw_bytes.startswith(b"\xd0\xcf\x11\xe0"):
            return pd.read_excel(io.BytesIO(raw_bytes))
        else:
            try:
                return pd.read_csv(io.BytesIO(raw_bytes), encoding="utf-8-sig")
            except Exception:
                return pd.read_csv(io.BytesIO(raw_bytes), encoding="gbk")
    except Exception:
        return None


def load_active_commodities() -> Tuple[Optional[pd.DataFrame], str]:
    """
    智能定位并加载商品表：
    1. 优先解密云端 commodities.dat
    2. 若无则检索本地明文表格（方便离线单机测试）
    """
    key = get_secret_key()

    # 优先检测加密文件 commodities.dat
    if os.path.exists("commodities.dat"):
        if not key:
            return None, "未配置解密密钥 (请在 Secrets 填入 COMMODITIES_KEY)"
        try:
            with open("commodities.dat", "rb") as f:
                cipher_data = f.read()
            cipher = Fernet(key)
            decrypted_bytes = cipher.decrypt(cipher_data)
            df = parse_raw_table_bytes(decrypted_bytes)
            if df is not None:
                return df, "商品列表 (已加密安全同步)"
        except Exception as e:
            return None, f"解密失败: {e}"

    # 备选：本地明文表格
    valid_exts = (".xlsx", ".xls", ".csv")
    candidates = []
    for fname in os.listdir("."):
        if fname.startswith("~$"):
            continue
        if any(fname.lower().endswith(ext) for ext in valid_exts) and "commodit" in fname.lower():
            full_path = os.path.join(".", fname)
            candidates.append((full_path, fname, os.path.getmtime(full_path)))

    if candidates:
        candidates.sort(key=lambda x: x, reverse=True)
        latest_path, latest_name, _ = candidates[0]
        try:
            with open(latest_path, "rb") as f:
                df = parse_raw_table_bytes(f.read())
            return df, latest_name
        except Exception:
            pass

    return None, "未检测到商品库"


def get_sku_info_from_df(sku: str, df: Optional[pd.DataFrame], sku_mapping: Optional[Dict[str, str]] = None) -> dict:
    lookup_sku = sku
    if sku_mapping and sku in sku_mapping:
        lookup_sku = sku_mapping[sku]

    if df is None:
        return {"product": "未找到商品库", "brand": "请检查表格"}

    col_sku = next((c for c in df.columns if str(c).strip().upper() == "SKU"), None)
    col_product = next((c for c in df.columns if any(k in str(c) for k in ("品名", "商品", "名称"))), None)
    col_brand = next((c for c in df.columns if any(k in str(c) for k in ("品牌", "工厂"))), None)

    if not col_sku or not col_product or not col_brand:
        return {"product": "表格列缺失", "brand": "缺少SKU/品名/品牌"}

    match = df[df[col_sku].astype(str).str.strip() == lookup_sku.strip()]
    if match.empty:
        if lookup_sku != sku:
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
# 3. 分隔页绘制与面单重排
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
    commodities_df: Optional[pd.DataFrame],
    sku_mapping: Optional[Dict[str, str]] = None
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
        info = get_sku_info_from_df(sku, commodities_df, sku_mapping=sku_mapping)

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
# 4. Streamlit 页面与操作流
# ==============================================================================

def main():
    st.subheader("📦 亚马逊外箱面单批量处理")

    # 1. 加载特殊映射字典
    if "sku_mapping" not in st.session_state:
        st.session_state["sku_mapping"] = load_sku_mapping()
    current_mapping = st.session_state["sku_mapping"]

    # 2. 自动解密加载商品库
    active_df, table_label = load_active_commodities()

    # 状态指示
    custom_uploaded = st.session_state.get("custom_commodities", None)
    if custom_uploaded is not None:
        table_pill_label = f"🟢 自定义: {custom_uploaded.name} ▾"
    elif active_df is not None:
        table_pill_label = f"🟢 {table_label} ▾"
    else:
        table_pill_label = f"🔴 {table_label} ▾"

    map_count = len(current_mapping)
    mapping_pill_label = f"⚡ 特殊映射 ({map_count}条) ▾" if map_count > 0 else "⚡ 特殊映射 ▾"

    # 3. 顶部并排双极简胶囊
    col_p1, col_p2, _ = st.columns([1.5, 1.2, 1.3])

    with col_p1:
        with st.popover(table_pill_label):
            st.caption("临时更换商品库（仅本次生效）：")
            custom_file = st.file_uploader(
                "上传替代商品列表",
                type=["xlsx", "xls", "csv"],
                label_visibility="collapsed",
                key="custom_commodities"
            )
            if custom_uploaded is not None and st.button("恢复默认商品库", use_container_width=True):
                del st.session_state["custom_commodities"]
                st.rerun()

    with col_p2:
        with st.popover(mapping_pill_label):
            st.caption("双击编辑，支持从 Excel 复制两列直接粘贴：")
            rows = [{"面单SKU": k, "商品库SKU": v} for k, v in current_mapping.items()]
            if not rows:
                rows = [{"面单SKU": "", "商品库SKU": ""}]
            df_mapping = pd.DataFrame(rows)

            edited_df = st.data_editor(
                df_mapping,
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                height=220,
                column_config={
                    "面单SKU": st.column_config.TextColumn("面单 SKU", required=True),
                    "商品库SKU": st.column_config.TextColumn("商品库 SKU", required=True)
                },
                key="sku_mapping_editor"
            )

            c_btn1, c_btn2 = st.columns(2)
            if c_btn1.button("保存规则", type="primary", use_container_width=True):
                new_map = {}
                for _, r in edited_df.iterrows():
                    src = str(r.get("面单SKU", "")).strip()
                    tgt = str(r.get("商品库SKU", "")).strip()
                    if src and tgt and src != "nan" and tgt != "nan":
                        new_map[src] = tgt
                st.session_state["sku_mapping"] = new_map
                save_sku_mapping(new_map)
                st.rerun()

            if c_btn2.button("清空全部", use_container_width=True):
                st.session_state["sku_mapping"] = {}
                save_sku_mapping({})
                st.rerun()

    # 4. 判定最终使用的商品库
    if custom_uploaded is not None:
        active_df = parse_raw_table_bytes(custom_uploaded.getvalue())

    # 5. 面单批量上传
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
                        active_df,
                        sku_mapping=current_mapping
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
                # 醒目的下载按钮
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

                # 偏好 B：统计数字与大表全部默认折叠隐藏
                with st.expander("📊 查看处理数据与明细 (点击展开)"):
                    distinct_skus = len(set(r["SKU"] for r in all_table_data))
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("文件数", f"{num_files}")
                    c2.metric("总原箱数", f"{total_orig_pages}")
                    c3.metric("总SKU数", f"{distinct_skus}")
                    c4.metric("总生成页数", f"{total_out_pages}")

                    if all_table_data:
                        df_display = pd.DataFrame(all_table_data)
                        cols = ["来源面单", "SKU", "品名", "工厂/品牌", "数量"]
                        display_cols = [c for c in cols if c in df_display.columns]
                        st.dataframe(df_display[display_cols], use_container_width=True, hide_index=True)

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
