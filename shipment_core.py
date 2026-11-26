#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""海外仓备货单 → 发货单处理结果 核心转换逻辑（供脚本版与网页版共用）。

外箱箱规来源（重要）：
- 外箱长/宽/高(cm)  ← 商品库「箱规长/宽/高(cm)」
- 单箱重量(kg)      ← 商品库「单箱重量(kg)」
- 单箱数量(pcs)     ← 商品库「单箱数量(pcs)」（输入文件缺失时兜底）
- 供应商            ← 商品库「商品品牌」（商品列表的商品品牌即供应商名称）
- 商品图片          ← 商品库「图片链接」

注意：输入文件「备货单详情」里的「商品包装规格长/宽/高(cm)」「商品包装重量(g)」
是【单个商品的小包装】尺寸，不是外箱箱规，不参与计算。
"""

import io
import re

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

DETAIL_COLUMNS = [
    "SKU", "品名", "商品图片", "发货量", "箱数", "单箱数量",
    "物流中心编码", "供应商", "货件编号", "ReferenceId", "箱号",
    "总箱数编号", "外箱重量(kg)", "外箱总重量(kg)", "外箱长(cm)",
    "外箱宽(cm)", "外箱高(cm)", "外箱体积(m³)", "外箱总体积(m³)",
    "外箱总体积重(kg)", "创建时间", "发货时间", "物流商",
]

SUMMARY_COLUMNS = [
    "物流中心编码", "货件编号", "ReferenceId", "品名", "供应商",
    "总箱数", "外箱总重量", "外箱总体积", "外箱总体积重(kg)",
]


def _num(v, default=np.nan):
    if v is None:
        return default
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    s = str(v).strip()
    if s in ("", "nan", "None", "NaN"):
        return default
    try:
        return float(s)
    except (ValueError, TypeError):
        return default


def extract_order_code(remark, fallback):
    remark = "" if remark is None else str(remark)
    m = re.search(r"货件编号[：:]\s*([A-Za-z0-9\-]+)", remark)
    if m and m.group(1).strip():
        return m.group(1).strip()
    return str(fallback).strip() if fallback is not None else ""


# 云数据库英文字段 → 转换逻辑使用的中文字段名（与 Excel 商品库列名对齐）
_CLOUD_FIELD_MAP = {
    "sku": "SKU",
    "name": "品名",
    "brand": "商品品牌",
    "box_length_cm": "箱规长(cm)",
    "box_width_cm": "箱规宽(cm)",
    "box_height_cm": "箱规高(cm)",
    "box_weight_kg": "单箱重量(kg)",
    "box_qty": "单箱数量(pcs)",
    "image_url": "图片链接",
    "unit_price": "采购单价",
    "supplier": "供应商",
}


def load_commodities(path_or_df):
    """加载商品库，返回 {SKU: dict} 映射。输入可以是路径或 DataFrame。"""
    if isinstance(path_or_df, pd.DataFrame):
        df = path_or_df
    else:
        df = pd.read_excel(path_or_df, sheet_name="商品")
    df = df.copy()
    df["SKU"] = df["SKU"].astype(str).str.strip()
    # 去掉重复 SKU（保留第一条）
    df = df.drop_duplicates(subset="SKU", keep="first")
    return df.set_index("SKU").to_dict("index")


def load_commodities_from_cloud(rows):
    """从云数据库查询结果（list[dict]，英文字段）构建商品库映射 {SKU: dict}。

    rows 来自 cloud.database.from('commodities').select('*') 的 data。
    英文字段会被映射成转换逻辑使用的中文字段名，与本地 Excel 商品库对齐。
    """
    comm_dict = {}
    for r in rows:
        sku = str(r.get("sku") or "").strip()
        if not sku:
            continue
        item = {}
        for en, cn in _CLOUD_FIELD_MAP.items():
            v = r.get(en)
            if v is not None and str(v).strip() != "":
                item[cn] = v
        comm_dict[sku] = item
    return comm_dict


def _match_commodity(sku, comm_dict):
    """按 SKU 匹配商品库，支持 ZF- 前缀容错。"""
    sku = str(sku).strip()
    no_zf = sku[3:] if sku.upper().startswith("ZF-") else sku
    with_zf = f"ZF-{sku}" if not sku.upper().startswith("ZF-") else sku
    return (
        comm_dict.get(sku)
        or comm_dict.get(no_zf)
        or comm_dict.get(with_zf)
        or {}
    )


def extract_pcs_from_text(text):
    """从品名里抠 PCS 数（如「灵芝盘金色 25PC」→25），无则返回 1。"""
    match = re.search(r"(\d+)\s*(?:pc|pcs|PC|PCS|只|件|套)", str(text), re.I)
    return int(match.group(1)) if match else 1


def _get_dim(matched, keys, fallback=np.nan):
    """从商品库匹配结果里按优先级取数值字段。"""
    for k in keys:
        if k in matched:
            v = _num(matched[k])
            if not np.isnan(v) and v > 0:
                return float(v)
    return float(fallback) if not np.isnan(fallback) else np.nan


def convert_df(df, comm_dict=None):
    """把「备货单详情」DataFrame 转成 (df_detail, df_summary, kpi, missing_skus, missing_fields)。

    comm_dict: 商品库映射 {SKU: {...}}，可选；提供则兜底补外箱箱规/供应商/图片。
    missing_skus: 完全没在商品库匹配到的 SKU 列表。
    missing_fields: {SKU: [缺哪些关键字段]}，仅统计箱规长/宽/高、单箱重量这些
                    缺失后会导致体积/体积重/重量算不出来的字段。
    """
    if comm_dict is None:
        comm_dict = {}

    df["_货件编号"] = [
        extract_order_code(r, b)
        for r, b in zip(df.get("单据备注"), df.get("备货单号"))
    ]

    rows = []
    missing_skus = []
    missing_fields = {}  # {SKU: [缺哪些关键字段]}
    grouped = df.groupby("_货件编号", sort=False)

    for order_code, grp in grouped:
        cursor = 0
        total_boxes = int(grp["箱数"].fillna(0).astype(int).sum())
        for _, r in grp.iterrows():
            sku = str(r.get("SKU") or "").strip()
            matched = _match_commodity(sku, comm_dict)
            if not matched:
                missing_skus.append(sku)

            boxes = int(_num(r.get("箱数"), 0) or 0)
            # 单箱数量：输入文件优先，缺失用商品库「单箱数量(pcs)」
            carton = _num(r.get("单箱数量"))
            if np.isnan(carton) or carton <= 0:
                carton = _get_dim(matched, ["单箱数量(pcs)", "单箱数量"], 0)
            carton = int(carton) if not np.isnan(carton) and carton > 0 else 0

            qty = int(_num(r.get("备货量"), 0) or 0)
            if qty <= 0 and boxes > 0 and carton > 0:
                qty = boxes * carton

            # 外箱箱规：一律从商品库取
            l = _get_dim(matched, ["箱规长(cm)", "外箱长(cm)", "长(cm)"])
            w = _get_dim(matched, ["箱规宽(cm)", "外箱宽(cm)", "宽(cm)"])
            h = _get_dim(matched, ["箱规高(cm)", "外箱高(cm)", "高(cm)"])
            # 单箱重量：商品库「单箱重量(kg)」
            single_weight_kg = _get_dim(
                matched, ["单箱重量(kg)", "重量(kg)", "毛重(kg)", "单箱重量"]
            )

            # 检测关键字段缺失（箱规长/宽/高、单箱重量），这些缺了会导致体积/体积重/重量算不出
            if matched:
                field_labels = [
                    ("箱规长", l), ("箱规宽", w), ("箱规高", h),
                    ("单箱重量", single_weight_kg),
                ]
                miss = [name for name, val in field_labels if np.isnan(val) or val <= 0]
                if miss:
                    missing_fields.setdefault(sku, miss)

            single_weight = round(single_weight_kg, 2) if not np.isnan(single_weight_kg) and single_weight_kg > 0 else ""
            single_vol = round(l * w * h / 1e6, 4) if (
                not np.isnan(l) and not np.isnan(w) and not np.isnan(h)
                and l > 0 and w > 0 and h > 0
            ) else ""

            total_weight = round(single_weight * boxes, 2) if (
                isinstance(single_weight, (int, float)) and boxes > 0
            ) else ""
            total_vol = round(single_vol * boxes, 4) if (
                isinstance(single_vol, (int, float)) and boxes > 0
            ) else ""
            total_vol_weight = round(single_vol * boxes * 167, 2) if (
                isinstance(single_vol, (int, float)) and boxes > 0
            ) else ""

            box_range = f"{cursor + 1}-{cursor + boxes}" if boxes > 0 else ""
            cursor += boxes

            # 供应商 / 图片 / 品名：输入优先，商品库兜底
            supplier = matched.get("商品品牌") or matched.get("供应商名称") or matched.get("供应商") or ""
            image = matched.get("图片链接") or matched.get("图片") or ""
            product_name = str(r.get("品名") or "").strip() or (
                matched.get("品名") or matched.get("中文品名") or ""
            )

            rows.append({
                "SKU": sku,
                "品名": str(product_name).strip(),
                "商品图片": "" if pd.isna(image) else str(image).strip(),
                "发货量": qty,
                "箱数": boxes,
                "单箱数量": carton,
                "物流中心编码": str(r.get("收货仓库") or "").strip(),
                "供应商": "" if pd.isna(supplier) else str(supplier).strip(),
                "货件编号": order_code,
                "ReferenceId": "",
                "箱号": box_range,
                "总箱数编号": total_boxes,
                "外箱重量(kg)": single_weight,
                "外箱总重量(kg)": total_weight,
                "外箱长(cm)": round(l, 1) if not np.isnan(l) and l > 0 else "",
                "外箱宽(cm)": round(w, 1) if not np.isnan(w) and w > 0 else "",
                "外箱高(cm)": round(h, 1) if not np.isnan(h) and h > 0 else "",
                "外箱体积(m³)": single_vol,
                "外箱总体积(m³)": total_vol,
                "外箱总体积重(kg)": total_vol_weight,
                "创建时间": str(r.get("创建时间") or "").strip(),
                "发货时间": str(r.get("实际发货时间") or "").strip(),
                "物流商": str(r.get("物流商") or "").strip(),
            })

    df_detail = pd.DataFrame(rows, columns=DETAIL_COLUMNS).astype(object)
    df_detail = df_detail.where(pd.notnull(df_detail), "")
    df_detail = df_detail.replace({"nan": "", "None": ""})

    # 汇总结果
    summary_rows = []
    for order_code, grp in grouped:
        first = grp.iloc[0]
        boxes_sum = int(grp["箱数"].fillna(0).astype(int).sum())
        sw, sv, svw = 0.0, 0.0, 0.0
        first_sku = str(first.get("SKU") or "").strip()
        first_matched = _match_commodity(first_sku, comm_dict)

        for _, r in grp.iterrows():
            b = int(_num(r.get("箱数"), 0) or 0)
            sku = str(r.get("SKU") or "").strip()
            m = _match_commodity(sku, comm_dict)
            l = _get_dim(m, ["箱规长(cm)", "外箱长(cm)", "长(cm)"])
            w = _get_dim(m, ["箱规宽(cm)", "外箱宽(cm)", "宽(cm)"])
            h = _get_dim(m, ["箱规高(cm)", "外箱高(cm)", "高(cm)"])
            g = _get_dim(m, ["单箱重量(kg)", "重量(kg)", "毛重(kg)", "单箱重量"])
            if not np.isnan(g) and g > 0:
                sw += g * b
            if (not np.isnan(l) and not np.isnan(w) and not np.isnan(h)
                    and l > 0 and w > 0 and h > 0):
                vol = l * w * h / 1e6
                sv += vol * b
                svw += vol * b * 167

        supplier = first_matched.get("商品品牌") or first_matched.get("供应商名称") or first_matched.get("供应商") or ""
        product_name = str(first.get("品名") or "").strip() or (
            first_matched.get("品名") or first_matched.get("中文品名") or ""
        )
        summary_rows.append({
            "物流中心编码": str(first.get("收货仓库") or "").strip(),
            "货件编号": order_code,
            "ReferenceId": "",
            "品名": str(product_name).strip(),
            "供应商": "" if pd.isna(supplier) else str(supplier).strip(),
            "总箱数": boxes_sum,
            "外箱总重量": round(sw, 2) if sw > 0 else "",
            "外箱总体积": round(sv, 4) if sv > 0 else "",
            "外箱总体积重(kg)": round(svw, 2) if svw > 0 else "",
        })

    df_summary = pd.DataFrame(summary_rows, columns=SUMMARY_COLUMNS).astype(object)
    df_summary = df_summary.where(pd.notnull(df_summary), "")
    df_summary = df_summary.replace({"nan": "", "None": ""})

    total_pcs = int(sum(
        int(_num(r.get("备货量"), 0) or 0) * extract_pcs_from_text(r.get("品名"))
        for _, r in df.iterrows()
    ))
    kpi = {
        "total_box": int(df_summary["总箱数"].astype(str).str.strip().replace("", "0").astype(int).sum()),
        "total_sets": int(sum(int(_num(r.get("备货量"), 0) or 0) for _, r in df.iterrows())),
        "total_pcs": total_pcs,
        "total_weight": round(sum(
            float(x) for x in df_summary["外箱总重量"] if str(x) not in ("", "nan", "None")
        ), 2),
        "total_volume": round(sum(
            float(x) for x in df_summary["外箱总体积"] if str(x) not in ("", "nan", "None")
        ), 4),
        "orders": len(df_summary),
    }

    return df_detail, df_summary, kpi, list(dict.fromkeys(missing_skus)), missing_fields


def to_excel_bytes(df_detail, df_summary):
    """返回美化后的 Excel 字节流。"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_detail.to_excel(writer, sheet_name="详细数据", index=False)
        df_summary.to_excel(writer, sheet_name="汇总结果", index=False)

    wb = load_workbook(output)
    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
    stripe_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for sheetname in ["详细数据", "汇总结果"]:
        ws = wb[sheetname]
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row)):
            for cell in row:
                cell.border = thin_border
                cell.alignment = center_align
                if i == 0:
                    cell.fill = header_fill
                    cell.font = header_font
                elif i % 2 == 0:
                    cell.fill = stripe_fill
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                v = str(cell.value) if cell.value is not None else ""
                try:
                    length = len(v.encode("gbk"))
                    if length > max_len:
                        max_len = length
                except Exception:
                    pass
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 48)

    final_stream = io.BytesIO()
    wb.save(final_stream)
    return final_stream.getvalue()
