#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""海外仓备货单处理工具（Streamlit 界面版）"""

from datetime import datetime

import pandas as pd
import streamlit as st

from shipment_core import convert_df, load_commodities, to_excel_bytes

st.set_page_config(
    page_title="海外仓备货单处理",
    page_icon="📦",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
<style>
    .stApp {
        background-color: #F8FAFC !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
        color: #0F172A !important;
    }
    #MainMenu, footer, header { visibility: hidden; }
    .block-container {
        max-width: 820px !important;
        padding-top: 2.2rem !important;
        padding-bottom: 3.5rem !important;
    }
    .header-box { text-align: center; padding: 0.5rem 0 0.8rem 0; }
    .header-badge {
        display: inline-flex; align-items: center; gap: 5px;
        padding: 4px 12px; font-size: 0.72rem; font-weight: 600;
        letter-spacing: 0.05em; text-transform: uppercase;
        background: #EEF2F6; color: #475569; border-radius: 9999px;
        margin-bottom: 0.6rem; border: 1px solid #E2E8F0;
    }
    .header-title {
        font-size: 2.1rem; font-weight: 800; color: #0F172A !important;
        letter-spacing: -0.03em; margin: 0;
    }
    .header-subtitle {
        font-size: 0.92rem; color: #64748B !important;
        margin-top: 0.45rem; font-weight: 400;
    }
    [data-testid="stFileUploader"] { background: transparent !important; }
    [data-testid="stFileUploader"] section {
        background-color: #FFFFFF !important;
        border: 1.8px dashed #CBD5E1 !important; border-radius: 16px !important;
        padding: 2.2rem 1.5rem !important; text-align: center !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.02) !important;
        transition: all 0.25s ease !important;
    }
    [data-testid="stFileUploader"] section:hover {
        border-color: #2563EB !important; background-color: #F8FAFC !important;
        box-shadow: 0 6px 20px rgba(37, 99, 235, 0.08) !important;
    }
    [data-testid="stFileUploader"] section button {
        background-color: #2563EB !important; color: #FFFFFF !important;
        border-radius: 8px !important; font-weight: 600 !important;
        border: none !important; padding: 0.45rem 1.4rem !important;
        box-shadow: 0 2px 5px rgba(37, 99, 235, 0.2) !important;
    }
    [data-testid="stFileUploader"] section button * { color: #FFFFFF !important; }
    [data-testid="stFileUploader"] section button:hover {
        background-color: #1D4ED8 !important;
        box-shadow: 0 4px 10px rgba(37, 99, 235, 0.3) !important;
    }
    [data-testid="stFileUploader"] label { display: none !important; }
    .metric-grid {
        display: grid; grid-template-columns: repeat(3, 1fr);
        gap: 12px; margin: 1.6rem 0 1.8rem 0;
    }
    .metric-card {
        background: #FFFFFF; padding: 1.1rem 1rem; border-radius: 12px;
        border: 1px solid #E2E8F0; box-shadow: 0 1px 3px rgba(0, 0, 0, 0.02);
        transition: all 0.2s ease;
    }
    .metric-card:hover {
        transform: translateY(-2px); box-shadow: 0 6px 16px rgba(0,0,0,0.05);
        border-color: #CBD5E1;
    }
    .metric-label {
        font-size: 0.76rem; font-weight: 600; color: #64748B;
        text-transform: uppercase; letter-spacing: 0.03em; margin-bottom: 0.35rem;
    }
    .metric-value {
        font-size: 1.55rem; font-weight: 800; color: #0F172A;
        letter-spacing: -0.02em; line-height: 1.15;
    }
    .metric-unit { font-size: 0.8rem; font-weight: 500; color: #94A3B8; margin-left: 0.25rem; }
    .stDownloadButton button {
        background: #0F172A !important; color: #FFFFFF !important;
        font-size: 0.96rem !important; font-weight: 600 !important;
        border-radius: 10px !important; padding: 0.7rem 1.8rem !important;
        border: none !important; box-shadow: 0 4px 12px rgba(15, 23, 42, 0.15) !important;
    }
    .stDownloadButton button:hover {
        background: #1E293B !important;
        box-shadow: 0 6px 18px rgba(15, 23, 42, 0.25) !important;
        transform: translateY(-1px);
    }
</style>
""",
    unsafe_allow_html=True,
)


def main():
    st.markdown(
        """
        <div class="header-box">
            <div class="header-badge">✨ 海外仓备货单处理</div>
            <h1 class="header-title">海外仓备货单处理工具</h1>
            <p class="header-subtitle">上传备货单 Excel，自动生成标准发货单处理结果</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    uploaded_file = st.file_uploader(
        "上传海外仓备货单 Excel 文件",
        type=["xlsx", "xls"],
        help="支持「备货单详情」sheet，自动提取货件编号并计算箱数、重量、体积",
    )

    # 商品库（可选，兜底外箱箱规/供应商/图片）
    with st.expander("🛒 商品库（可选，用于补外箱箱规/供应商/图片）", expanded=False):
        st.caption("外箱箱规、单箱重量、供应商、商品图片均从商品库按 SKU 匹配。不上传则这些字段留空。")
        comm_file = st.file_uploader(
            "上传商品库 Excel（含「商品」sheet）",
            type=["xlsx", "xls"],
            key="commodities",
            label_visibility="collapsed",
        )

    comm_dict = {}
    if comm_file is not None:
        comm_dict = load_commodities(pd.read_excel(comm_file, sheet_name="商品"))

    if uploaded_file is not None:
        try:
            with st.spinner("正在解析备货单并计算物流数据..."):
                df_raw = pd.read_excel(uploaded_file, sheet_name="备货单详情")
                df_detail, df_summary, kpi, missing_skus, missing_fields = convert_df(df_raw, comm_dict)
                excel_bytes = to_excel_bytes(df_detail, df_summary)

            # KPI 卡片
            st.markdown(
                f"""
                <div class="metric-grid">
                    <div class="metric-card">
                        <div class="metric-label">关联备货单号</div>
                        <div class="metric-value" style="font-size:1.15rem; word-break:break-all;">{df_summary['货件编号'].iloc[0] if len(df_summary) else '—'}</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">总装箱量</div>
                        <div class="metric-value">{kpi['total_box']:,}<span class="metric-unit">箱</span></div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">总备货量</div>
                        <div class="metric-value">{kpi['total_sets']:,}<span class="metric-unit">套</span></div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">总 PCS (单件实物)</div>
                        <div class="metric-value">{kpi['total_pcs']:,}<span class="metric-unit">件</span></div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">实重 / 总体积</div>
                        <div class="metric-value" style="font-size:1.15rem;">{kpi['total_weight']:,.2f}<span class="metric-unit">kg</span> / {kpi['total_volume']:,.4f}<span class="metric-unit">m³</span></div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-label">货件数</div>
                        <div class="metric-value">{kpi['orders']}<span class="metric-unit">个</span></div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            if missing_skus:
                st.warning(
                    f"⚠️ 提示：共有 {len(missing_skus)} 个 SKU 未在商品库中匹配到外箱箱规："
                    f" {', '.join(missing_skus[:8])}{'...' if len(missing_skus) > 8 else ''}"
                )

            if missing_fields:
                lines = []
                for sku, fields in missing_fields.items():
                    lines.append(f"- `{sku}` 缺：{'、'.join(fields)}")
                st.warning(
                    "⚠️ 提示：以下 SKU 在商品库中缺少关键字段（会导致体积 / 体积重 / 重量算不出来）：\n\n"
                    + "\n".join(lines[:15])
                    + ("\n\n…" if len(lines) > 15 else "")
                )

            # 导出按钮
            base = uploaded_file.name.rsplit(".", 1)[0]
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_filename = f"{base}_发货单处理结果_{ts}.xlsx"
            st.download_button(
                label="⬇️ 导出标准发货单与汇总表 (.xlsx)",
                data=excel_bytes,
                file_name=out_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                width="stretch",
            )

            st.write("")

            tab1, tab2 = st.tabs(["📝 详细数据", "📊 汇总结果"])
            with tab1:
                st.dataframe(df_detail, width="stretch")
            with tab2:
                st.dataframe(df_summary, width="stretch")

        except Exception as e:
            st.error(f"❌ 数据处理失败：{str(e)}")


if __name__ == "__main__":
    main()
