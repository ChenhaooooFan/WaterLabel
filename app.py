"""
NailVesta 水单发货系统（两阶段工作流）
========================================================
阶段 1（生成水单）:
    Lark CSV → 选日期 → 生成水单 CSV → 发给打 Label 的人

阶段 2（拣货 + 发货 + 核对）:
    Lark CSV + 图册 CSV + Label PDF → OCR 核对 + 生成拣货单/发货单
"""

import io
import os
import re
import subprocess
import tempfile
from datetime import datetime
from typing import Optional

import pandas as pd
import streamlit as st

# ============================================================
# 常量
# ============================================================
SENDER_INFO = {
    "name": "NailVesta", "company": None, "phone": "5105089943",
    "country": "US", "state": "CA", "city": "Los Angeles",
    "zip": "90071", "address": "515 S Flower St, Floor 18 & 19, STE 1901",
}
PACKAGE_DEFAULTS = {
    "weight": 0.3, "length": 20, "width": 15, "height": 2,
    "cn_name": "穿戴甲", "en_name": "Press-On Nails",
    "qty": 1, "declare_price": 5, "net_weight": 0.3,
}
SHUIDAN_HEADERS = [
    "客户订单号","物流产品(产品编号)","重量","长","宽","高",
    "发件人姓名","发件人公司","发件人电话","发件人国家","发件人省/州",
    "发件人城市","发件人邮编","发件人地址",
    "收件人姓名","收件人公司","收件人电话","收件人国家","收件人省/州",
    "收件人城市","收件人地址一","收件人地址二","收件人邮编",
    "中文品名1","英文品名1","SKU1","数量1","配货备注1",
    "申报单价1","单位净重(kg)1",
]
SIZE_COL = "Size'"

# ============================================================
# 地址解析
# ============================================================
_US_STATES_FULL = {
    "ALABAMA","ALASKA","ARIZONA","ARKANSAS","CALIFORNIA","COLORADO","CONNECTICUT",
    "DELAWARE","FLORIDA","GEORGIA","HAWAII","IDAHO","ILLINOIS","INDIANA","IOWA",
    "KANSAS","KENTUCKY","LOUISIANA","MAINE","MARYLAND","MASSACHUSETTS","MICHIGAN",
    "MINNESOTA","MISSISSIPPI","MISSOURI","MONTANA","NEBRASKA","NEVADA",
    "NEW HAMPSHIRE","NEW JERSEY","NEW MEXICO","NEW YORK","NORTH CAROLINA",
    "NORTH DAKOTA","OHIO","OKLAHOMA","OREGON","PENNSYLVANIA","RHODE ISLAND",
    "SOUTH CAROLINA","SOUTH DAKOTA","TENNESSEE","TEXAS","UTAH","VERMONT",
    "VIRGINIA","WASHINGTON","WEST VIRGINIA","WISCONSIN","WYOMING",
    "DISTRICT OF COLUMBIA","PUERTO RICO",
}
_STATE_ABBR = {"AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL",
"IN","IA","KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV","NH",
"NJ","NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN","TX","UT","VT",
"VA","WA","WV","WI","WY","DC","PR"}
_STATE_FULL_TO_ABBR = {
    "ALABAMA":"AL","ALASKA":"AK","ARIZONA":"AZ","ARKANSAS":"AR","CALIFORNIA":"CA",
    "COLORADO":"CO","CONNECTICUT":"CT","DELAWARE":"DE","FLORIDA":"FL","GEORGIA":"GA",
    "HAWAII":"HI","IDAHO":"ID","ILLINOIS":"IL","INDIANA":"IN","IOWA":"IA","KANSAS":"KS",
    "KENTUCKY":"KY","LOUISIANA":"LA","MAINE":"ME","MARYLAND":"MD","MASSACHUSETTS":"MA",
    "MICHIGAN":"MI","MINNESOTA":"MN","MISSISSIPPI":"MS","MISSOURI":"MO","MONTANA":"MT",
    "NEBRASKA":"NE","NEVADA":"NV","NEW HAMPSHIRE":"NH","NEW JERSEY":"NJ","NEW MEXICO":"NM",
    "NEW YORK":"NY","NORTH CAROLINA":"NC","NORTH DAKOTA":"ND","OHIO":"OH","OKLAHOMA":"OK",
    "OREGON":"OR","PENNSYLVANIA":"PA","RHODE ISLAND":"RI","SOUTH CAROLINA":"SC",
    "SOUTH DAKOTA":"SD","TENNESSEE":"TN","TEXAS":"TX","UTAH":"UT","VERMONT":"VT",
    "VIRGINIA":"VA","WASHINGTON":"WA","WEST VIRGINIA":"WV","WISCONSIN":"WI","WYOMING":"WY",
    "DISTRICT OF COLUMBIA":"DC","PUERTO RICO":"PR",
}


def state_to_abbr(state: str) -> str:
    if not state: return ""
    s = state.strip().upper()
    if s in _STATE_FULL_TO_ABBR: return _STATE_FULL_TO_ABBR[s]
    if s in _STATE_ABBR: return s
    return state if len(state) == 2 else state


def _is_zip(s): return bool(re.fullmatch(r"\d{5}(-\d{4})?", s.strip()))


def _is_phone_line(s):
    if re.match(r"^[\sa]*\(?\+?1?\)?[\s\-\.]?\(?\d{3}\)?[\s\-\.]?\d{3}[\s\-\.]?\d{4}", s.strip()):
        return True
    return bool(re.match(r"^(Tel|Phone|WhatsApp)\s*[:\uff1a]", s.strip(), re.I))


def _normalize_state(tok):
    t = tok.strip().upper()
    if t in _US_STATES_FULL: return t
    if t in _STATE_ABBR: return t
    return None


def _clean_phone(s):
    digits = re.sub(r"\D", "", s)
    if digits.startswith("1") and len(digits) == 11: digits = digits[1:]
    return digits


def _clean_street(s):
    s = re.sub(r"\(([^)]*)\)", r" \1", s)
    s = re.sub(r"#\s+", "#", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def parse_shipping_info(text):
    if not isinstance(text, str): return {}
    lines = [ln.strip() for ln in text.replace("\r","").split("\n") if ln.strip()]
    while lines and re.search(r"[\u4e00-\u9fff]", lines[0]) and len(lines[0]) > 15:
        lines.pop(0)
    if not lines: return {}
    info = {"name":"","phone":"","street":"","street2":"",
            "city":"","state":"","country":"United States","zip":""}
    zip_idx = None
    for i in range(len(lines)-1,-1,-1):
        if _is_zip(lines[i]):
            zip_idx = i; info["zip"] = lines[i].strip(); break
    phone_idx = None
    for i, ln in enumerate(lines):
        if i == zip_idx: continue
        if _is_phone_line(ln):
            info["phone"] = _clean_phone(ln); phone_idx = i; break
    csc_idx = None
    for i, ln in enumerate(lines):
        if i in (zip_idx, phone_idx): continue
        if "," in ln:
            parts = [p.strip() for p in ln.split(",")]
            if (any("UNITED STATES" in p.upper() or "USA" in p.upper() for p in parts)
                    and len(parts) >= 3):
                st_ = _normalize_state(parts[-2])
                if st_:
                    info["state"] = st_; info["country"] = parts[-1]
                    head = parts[0]; csc_idx = i
                    has_digit = bool(re.search(r"\d", head))
                    m = re.match(r"^(.*?)\s*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)\s*$", head)
                    if has_digit and m and re.search(r"\d", m.group(1)):
                        info["street"] = m.group(1).strip()
                        info["city"] = m.group(2).strip()
                    else:
                        info["city"] = head
                    break
    used = {zip_idx, phone_idx, csc_idx}
    others = [(i, lines[i]) for i in range(len(lines)) if i not in used]
    name_pos = None
    for k, (i, ln) in enumerate(others):
        clean = re.sub(r"^Name\s*[:\uff1a]\s*", "", ln, flags=re.I)
        if not re.search(r"\d", clean) and "," not in clean and ":" not in clean:
            info["name"] = clean; name_pos = k; break
    street_lines = [
        re.sub(r"^(Address|Street|Apt|Suite)\s*[:\uff1a]\s*", "", ln, flags=re.I)
        for k, (i, ln) in enumerate(others) if k != name_pos
    ]
    if not info["street"]:
        if street_lines:
            info["street"] = street_lines[0]
            if len(street_lines) > 1:
                info["street2"] = " ".join(street_lines[1:])
    else:
        if street_lines: info["street2"] = " ".join(street_lines)
    info["street"] = _clean_street(info["street"])
    info["street2"] = _clean_street(info["street2"]) if info["street2"] else ""
    return info


# ============================================================
# 数据加载
# ============================================================
@st.cache_data(show_spinner=False)
def load_lark_data(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_csv(io.BytesIO(file_bytes))
    df["日期"] = df["日期"].astype(str).str.strip()
    return df


@st.cache_data(show_spinner=False)
def load_catalog(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_csv(io.BytesIO(file_bytes))
    keep = ["SKU","中文名称","款式英文名称","甲型","图片","库位","所属系列"]
    keep = [c for c in keep if c in df.columns]
    df = df[keep].copy()
    df["SKU"] = df["SKU"].astype(str).str.strip()
    if "款式英文名称" in df.columns:
        df["款式英文名称"] = df["款式英文名称"].astype(str).str.strip()
    return df


def filter_orders_for_date(df, target_date):
    out = df[
        (df["日期"] == target_date)
        & df["Order ID"].notna()
        & df["Shipping Info"].notna()
    ].copy()
    out["Order ID"] = out["Order ID"].apply(
        lambda x: str(int(x)) if isinstance(x, float) and float(x).is_integer() else str(x)
    )
    return out


# ============================================================
# 组合装拆单 + 图册富集
# ============================================================
def _parse_size(size_raw):
    s = str(size_raw).strip() if size_raw else ""
    if ";" in s:
        mapping = {"_default": s}
        for part in s.split(";"):
            part = part.strip()
            m = re.match(r"^(.+?)\s+(S|M|L|XL|XS|\d+\s*个)$", part, re.I)
            if m: mapping[m.group(1).strip()] = m.group(2).strip()
        return mapping
    return {"_default": s}


def explode_orders(orders, catalog):
    cat_by_name = cat_by_sku = None
    if catalog is not None and not catalog.empty:
        if "款式英文名称" in catalog.columns:
            tmp = catalog.dropna(subset=["款式英文名称"]).copy()
            tmp["_key"] = tmp["款式英文名称"].astype(str).str.strip().str.lower()
            cat_by_name = tmp.set_index("_key")
        cat_by_sku = catalog.set_index("SKU")

    def lookup_by_name(name):
        if cat_by_name is None or not name: return None
        key = name.strip().lower()
        if key in cat_by_name.index:
            row = cat_by_name.loc[key]
            if isinstance(row, pd.DataFrame): row = row.iloc[0]
            return row
        return None

    def lookup_by_sku(sku):
        if cat_by_sku is None or not sku: return None
        if sku in cat_by_sku.index:
            row = cat_by_sku.loc[sku]
            if isinstance(row, pd.DataFrame): row = row.iloc[0]
            return row
        return None

    rows = []
    for _, r in orders.iterrows():
        order_id = r["Order ID"]
        ship_info = r.get("Shipping Info", "")
        lark_loc = str(r.get("库位", "") or "").strip()
        if lark_loc.lower() == "nan": lark_loc = ""

        pn_raw = str(r.get("Product Name", "") or "").strip()
        sku_raw = str(r.get("SKU", "") or "").strip()
        size_raw = str(r.get(SIZE_COL, "") or "").strip()
        if pn_raw.lower() == "nan": pn_raw = ""
        if sku_raw.lower() == "nan": sku_raw = ""
        if size_raw.lower() == "nan": size_raw = ""

        pn_list = [p.strip() for p in pn_raw.split(",") if p.strip()] if pn_raw else []
        sku_list = [s.strip() for s in sku_raw.split(",") if s.strip()] if sku_raw else []
        size_map = _parse_size(size_raw)
        n = max(len(pn_list), len(sku_list), 1)

        for i in range(n):
            name_i = pn_list[i] if i < len(pn_list) else ""
            sku_i = sku_list[i] if i < len(sku_list) else ""
            cat_row = lookup_by_name(name_i) if name_i else None
            if cat_row is None:
                cat_row = lookup_by_sku(sku_i) if sku_i else None
            cn_name = cat_loc = ""
            cat_en_name = name_i
            if cat_row is not None:
                cn_v = cat_row.get("中文名称", "")
                cn_name = "" if pd.isna(cn_v) else str(cn_v).strip()
                loc_v = cat_row.get("库位", "")
                cat_loc = "" if pd.isna(loc_v) else str(loc_v).strip()
                if not name_i:
                    en_v = cat_row.get("款式英文名称", "")
                    cat_en_name = "" if pd.isna(en_v) else str(en_v).strip()
                if not sku_i and cat_row.name:
                    sku_i = str(cat_row.name)
            size_i = size_map.get(name_i, size_map.get("_default", size_raw))
            final_loc = cat_loc or lark_loc
            rows.append({
                "Order ID": order_id, "英文款式": cat_en_name, "SKU": sku_i,
                "中文名": cn_name, "库位": final_loc, "Size": size_i,
                "Shipping Info": ship_info,
                "Full SKU": str(r.get("Full SKU", "") or "").strip(),
            })
    return pd.DataFrame(rows)


# ============================================================
# 文件 1：拣货表（按库位汇总，符合截图格式）
# ============================================================
def build_picking_summary_csv(exploded: pd.DataFrame) -> bytes:
    """
    按库位 + Product Name 汇总，每行: 库位 / Product Name / S / M / L / Total
    """
    if len(exploded) == 0:
        return pd.DataFrame(columns=["库位","Product Name","S","M","L","其他","Total"]).to_csv(index=False).encode("utf-8-sig")

    df = exploded.copy()
    df["Size"] = df["Size"].fillna("").astype(str).str.strip().str.upper()

    # 标准化 Size
    def size_bucket(s):
        if s in ("S","M","L"): return s
        if s in ("XS","XL"): return s
        return "其他"
    df["SizeBucket"] = df["Size"].apply(size_bucket)

    # group
    grouped = df.groupby(["库位", "英文款式", "SizeBucket"]).size().unstack(fill_value=0)
    # 确保 S/M/L 列存在
    for c in ["S","M","L"]:
        if c not in grouped.columns: grouped[c] = 0
    other_cols = [c for c in grouped.columns if c not in ("S","M","L")]
    grouped["其他"] = grouped[other_cols].sum(axis=1) if other_cols else 0
    grouped["Total"] = grouped[["S","M","L","其他"]].sum(axis=1)
    grouped = grouped[["S","M","L","其他","Total"]].reset_index()
    grouped = grouped.rename(columns={"英文款式": "Product Name"})

    # 排序：库位先，空库位最后
    grouped["_loc_sort"] = grouped["库位"].apply(lambda x: (x == "", x))
    grouped = grouped.sort_values(by=["_loc_sort", "Product Name"]).drop(columns=["_loc_sort"]).reset_index(drop=True)

    # 如果"其他"列全是 0，就删掉
    if (grouped["其他"] == 0).all():
        grouped = grouped.drop(columns=["其他"])

    return grouped.to_csv(index=False).encode("utf-8-sig")


# ============================================================
# 文件 2：Packing Slip CSV
# ============================================================
def build_packing_slip_csv(exploded: pd.DataFrame, date_str: str) -> bytes:
    rows = []
    for order_id, grp in exploded.groupby("Order ID", sort=False):
        first = grp.iloc[0]
        ship = parse_shipping_info(first.get("Shipping Info", ""))
        addr_parts = [p for p in [
            ship.get("street",""), ship.get("street2",""),
            f"{ship.get('city','')}, {state_to_abbr(ship.get('state',''))} {ship.get('zip','')}".strip(", "),
            ship.get("country","United States"),
        ] if p.strip(" ,")]
        full_address = " | ".join(addr_parts)
        for _, r in grp.iterrows():
            rows.append({
                "Order ID": order_id, "Date": date_str,
                "Recipient": ship.get("name","").title() if ship.get("name") else "",
                "Phone": ship.get("phone",""), "Address": full_address,
                "SKU": r["SKU"] or "", "Style": r["英文款式"] or "",
                "Chinese Name": r["中文名"] or "", "Size": r["Size"] or "", "Qty": 1,
            })
    return pd.DataFrame(rows).to_csv(index=False).encode("utf-8-sig")


# ============================================================
# 文件 3：水单 CSV
# ============================================================
def build_shuidan_csv(orders: pd.DataFrame) -> bytes:
    rows = []
    for _, r in orders.iterrows():
        ship = parse_shipping_info(r.get("Shipping Info", ""))
        rows.append({
            "客户订单号": r["Order ID"], "物流产品(产品编号)": "",
            "重量": PACKAGE_DEFAULTS["weight"], "长": PACKAGE_DEFAULTS["length"],
            "宽": PACKAGE_DEFAULTS["width"], "高": PACKAGE_DEFAULTS["height"],
            "发件人姓名": SENDER_INFO["name"], "发件人公司": "",
            "发件人电话": SENDER_INFO["phone"], "发件人国家": SENDER_INFO["country"],
            "发件人省/州": SENDER_INFO["state"], "发件人城市": SENDER_INFO["city"],
            "发件人邮编": SENDER_INFO["zip"], "发件人地址": SENDER_INFO["address"],
            "收件人姓名": ship.get("name",""), "收件人公司": "",
            "收件人电话": ship.get("phone",""), "收件人国家": "US",
            "收件人省/州": state_to_abbr(ship.get("state","")),
            "收件人城市": ship.get("city",""),
            "收件人地址一": ship.get("street",""),
            "收件人地址二": ship.get("street2",""),
            "收件人邮编": ship.get("zip",""),
            "中文品名1": PACKAGE_DEFAULTS["cn_name"],
            "英文品名1": PACKAGE_DEFAULTS["en_name"], "SKU1": "",
            "数量1": PACKAGE_DEFAULTS["qty"], "配货备注1": "",
            "申报单价1": PACKAGE_DEFAULTS["declare_price"],
            "单位净重(kg)1": PACKAGE_DEFAULTS["net_weight"],
        })
    return pd.DataFrame(rows, columns=SHUIDAN_HEADERS).to_csv(index=False).encode("utf-8-sig")


# ============================================================
# Label PDF OCR & 核对
# ============================================================
def check_ocr_available() -> tuple:
    """检查 tesseract + poppler 可用"""
    has_tess = subprocess.run(["which","tesseract"], capture_output=True).returncode == 0
    has_popp = subprocess.run(["which","pdftoppm"], capture_output=True).returncode == 0
    return has_tess, has_popp


def parse_label_pdf(pdf_bytes: bytes) -> list:
    """对 PDF 每一页 OCR，返回 list of dict: {page, order_id, tracking, name, address_lines, raw_text}"""
    has_tess, has_popp = check_ocr_available()
    if not (has_tess and has_popp):
        raise RuntimeError("OCR 工具未安装：需要 tesseract + poppler-utils（在 packages.txt 中声明）")

    with tempfile.TemporaryDirectory() as tmp:
        pdf_path = os.path.join(tmp, "labels.pdf")
        with open(pdf_path, "wb") as f:
            f.write(pdf_bytes)
        # 渲染所有页为 jpg
        subprocess.run(
            ["pdftoppm", "-jpeg", "-r", "200", pdf_path, os.path.join(tmp, "lbl")],
            capture_output=True, timeout=120,
        )
        jpgs = sorted([f for f in os.listdir(tmp) if f.startswith("lbl") and f.endswith(".jpg")])

        results = []
        for jpg in jpgs:
            jpg_path = os.path.join(tmp, jpg)
            page_num = int(re.search(r"lbl-(\d+)", jpg).group(1))
            text = subprocess.run(
                ["tesseract", jpg_path, "-"],
                capture_output=True, text=True, timeout=30,
            ).stdout
            results.append(_parse_label_text(text, page_num))
        results.sort(key=lambda x: x["page"])
        return results


def _parse_label_text(text: str, page: int) -> dict:
    """从 OCR 文本提取关键字段"""
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    info = {"page": page, "order_id": "", "tracking": "",
            "name": "", "address_lines": [], "raw_text": text}

    # Order ID：从下往上找 17+ 位纯数字（去空格后）
    for ln in reversed(lines):
        digits = re.sub(r"\s", "", ln)
        m = re.match(r"^(\d{17,20})(-\d+)?$", digits)
        if m:
            info["order_id"] = digits
            break

    # Tracking
    for ln in lines:
        if re.search(r"9\d{3}\s+\d{4}\s+\d{4}\s+\d{4}", ln):
            info["tracking"] = re.sub(r"\s+", "", ln)
            break

    # 收件人 + 地址
    in_addr = False
    for ln in lines:
        if "LOS ANGELES" in ln.upper() and "CA" in ln.upper():
            in_addr = True; continue
        if in_addr:
            if "TRACKING" in ln.upper(): break
            cleaned = re.sub(r"^([a-zA-Z]+:?\s*[=>:;©]+\s*|[=>:;©]+\s*|[a-z]\s+|[a-zA-Z0-9]{1,4}[=:;©]\s*)", "", ln)
            cleaned = re.sub(r"^[a-z]:\s*", "", cleaned)
            cleaned = re.sub(r"^[=>:;]+\s*", "", cleaned).strip()
            if cleaned and re.search(r"[A-Za-z0-9]", cleaned):
                info["address_lines"].append(cleaned)

    if info["address_lines"]:
        info["name"] = info["address_lines"][0].upper().strip()
    return info


def reconcile_labels(orders: pd.DataFrame, labels: list) -> pd.DataFrame:
    """
    把 OCR 的 Label 结果与 Lark 订单核对
    返回每单的核对结果 DataFrame:
      Order ID / 状态 / 收件人 (Lark) / 收件人 (Label) / 城市 (Label) / Tracking / 备注
    """
    # 按 Order ID 索引（去除 -N 后缀）
    label_by_oid = {}
    for lbl in labels:
        oid = lbl["order_id"]
        oid_base = re.sub(r"-\d+$", "", oid)
        if oid_base:
            label_by_oid[oid_base] = lbl

    rows = []
    for _, r in orders.iterrows():
        oid = str(r["Order ID"]).strip()
        ship = parse_shipping_info(r.get("Shipping Info", ""))
        lark_name = ship.get("name", "")
        lark_zip = ship.get("zip", "")
        lark_state = state_to_abbr(ship.get("state", ""))

        lbl = label_by_oid.get(oid)
        if not lbl:
            rows.append({
                "Order ID": oid, "状态": "❌ 缺面单",
                "Lark 收件人": lark_name, "Label 收件人": "—",
                "Lark 邮编": lark_zip, "Label 末行": "—",
                "Tracking": "—", "备注": "PDF 中找不到此订单的面单",
            })
            continue

        # 名字模糊匹配（去标点、空格、转大写）
        def norm(s): return re.sub(r"[^A-Z0-9]", "", s.upper()) if s else ""
        lark_n = norm(lark_name); lbl_n = norm(lbl["name"])
        # 检查 Label 的最末地址行是否含 Lark 的 ZIP
        last_line = lbl["address_lines"][-1] if lbl["address_lines"] else ""
        zip_in_label = lark_zip and lark_zip in last_line

        # 名字相似度（包含关系即视为通过）
        name_match = lark_n and lbl_n and (
            lark_n in lbl_n or lbl_n in lark_n
            or _ratio(lark_n, lbl_n) >= 0.7
        )

        if name_match and zip_in_label:
            status = "✅ 一致"
            note = ""
        elif name_match:
            status = "⚠️ 名字一致但邮编对不上"
            note = f"Lark={lark_zip} | Label={last_line}"
        elif zip_in_label:
            status = "⚠️ 邮编一致但名字 OCR 模糊"
            note = f"Lark={lark_name} | Label={lbl['name']}"
        else:
            status = "❌ 不匹配"
            note = f"Lark={lark_name} {lark_zip} | Label={lbl['name']} {last_line}"

        rows.append({
            "Order ID": oid, "状态": status,
            "Lark 收件人": lark_name, "Label 收件人": lbl["name"],
            "Lark 邮编": lark_zip, "Label 末行": last_line,
            "Tracking": lbl["tracking"], "备注": note,
        })

    # 反向：Label 里有但 Lark 里没有的
    lark_oids = set(str(o).strip() for o in orders["Order ID"])
    for oid_base, lbl in label_by_oid.items():
        if oid_base not in lark_oids:
            rows.append({
                "Order ID": lbl["order_id"], "状态": "❌ 多余面单",
                "Lark 收件人": "—", "Label 收件人": lbl["name"],
                "Lark 邮编": "—",
                "Label 末行": lbl["address_lines"][-1] if lbl["address_lines"] else "",
                "Tracking": lbl["tracking"], "备注": "Label PDF 含此订单但 Lark 当日无此单",
            })

    return pd.DataFrame(rows)


def _ratio(a, b):
    """简单的字符相似度（最长公共子序列长度 / 最长字符串长度）"""
    if not a or not b: return 0
    if a == b: return 1
    # 简化版 LCS
    m, n = len(a), len(b)
    dp = [[0]*(n+1) for _ in range(m+1)]
    for i in range(1, m+1):
        for j in range(1, n+1):
            if a[i-1] == b[j-1]: dp[i][j] = dp[i-1][j-1] + 1
            else: dp[i][j] = max(dp[i-1][j], dp[i][j-1])
    return dp[m][n] / max(m, n)


# ============================================================
# Streamlit UI
# ============================================================
st.set_page_config(page_title="NailVesta 发货系统", page_icon="💅", layout="wide")
st.title("💅 NailVesta 发货系统")

tab1, tab2 = st.tabs([
    "1️⃣ 第一步：生成水单（发给打 Label 的人）",
    "2️⃣ 第二步：拿到 Label PDF 后，核对 + 生成拣货单/发货单",
])

# ============================================================
# Tab 1: 生成水单
# ============================================================
with tab1:
    st.subheader("第 1 步 · 生成水单 CSV")
    st.caption("上传 Lark CSV → 选日期 → 下载水单 → 发给打 Label 的同事")

    lark_file_1 = st.file_uploader("Lark 水单 CSV", type=["csv"], key="lark1")

    if lark_file_1 is None:
        st.info("👆 请上传 Lark CSV")
    else:
        df1 = load_lark_data(lark_file_1.read())
        all_dates_1 = sorted(
            [d for d in df1["日期"].dropna().unique()
             if re.match(r"^\d{4}/\d{1,2}/\d{1,2}$", str(d))],
            key=lambda x: datetime.strptime(x, "%Y/%m/%d"), reverse=True,
        )
        if not all_dates_1:
            st.error("CSV 中找不到合法日期")
        else:
            sel_date_1 = st.selectbox(
                "发货日期（洛杉矶时间，默认最晚）",
                options=all_dates_1, index=0,
                format_func=lambda x: f"{x} {'⬅️ 最新' if x == all_dates_1[0] else ''}",
                key="d1",
            )
            orders_1 = filter_orders_for_date(df1, sel_date_1)
            st.metric("当日订单数", len(orders_1))

            if len(orders_1) == 0:
                st.warning(f"⚠️ {sel_date_1} 没有可发货订单")
            else:
                with st.expander("👀 订单预览", expanded=False):
                    preview_cols = ["Order ID", "Product Name", "Full SKU", SIZE_COL, "Shipping Info"]
                    preview = orders_1[[c for c in preview_cols if c in orders_1.columns]]
                    st.dataframe(preview, use_container_width=True, height=300)

                shuidan_csv = build_shuidan_csv(orders_1)
                date_compact = sel_date_1.replace("/", "")
                st.download_button(
                    "🚚 下载水单 CSV（发给打 Label 的人）",
                    data=shuidan_csv,
                    file_name=f"{date_compact}客人水单.csv",
                    mime="text/csv",
                    type="primary",
                    use_container_width=True,
                )

# ============================================================
# Tab 2: 核对 + 拣货单 + 发货单
# ============================================================
with tab2:
    st.subheader("第 2 步 · 核对面单 + 生成拣货单/发货单")
    st.caption("上传 Lark CSV + 图册 + Label PDF → 自动核对 + 生成拣货单/发货单")

    c_left, c_right = st.columns(2)
    with c_left:
        lark_file_2 = st.file_uploader("Lark 水单 CSV", type=["csv"], key="lark2")
        catalog_file_2 = st.file_uploader("产品图册 CSV", type=["csv"], key="cat2")
    with c_right:
        label_pdf = st.file_uploader("Label PDF（打 Label 后拿到的面单）", type=["pdf"], key="pdf2")

    if lark_file_2 is None or catalog_file_2 is None:
        st.info("👆 请上传 Lark CSV 和产品图册 CSV")
        st.stop()

    df2 = load_lark_data(lark_file_2.read())
    catalog2 = load_catalog(catalog_file_2.read())

    all_dates_2 = sorted(
        [d for d in df2["日期"].dropna().unique()
         if re.match(r"^\d{4}/\d{1,2}/\d{1,2}$", str(d))],
        key=lambda x: datetime.strptime(x, "%Y/%m/%d"), reverse=True,
    )
    if not all_dates_2:
        st.error("CSV 中找不到合法日期")
        st.stop()

    sel_date_2 = st.selectbox(
        "发货日期", options=all_dates_2, index=0,
        format_func=lambda x: f"{x} {'⬅️ 最新' if x == all_dates_2[0] else ''}",
        key="d2",
    )
    orders_2 = filter_orders_for_date(df2, sel_date_2)
    exploded_2 = explode_orders(orders_2, catalog2)

    m1, m2, m3 = st.columns(3)
    m1.metric("订单数", len(orders_2))
    m2.metric("拣货行数", len(exploded_2),
              delta=f"+{len(exploded_2)-len(orders_2)} 组合装" if len(exploded_2) > len(orders_2) else None)
    m3.metric("拣货库位", exploded_2["库位"].replace("", pd.NA).dropna().nunique() if len(exploded_2) else 0)

    if len(orders_2) == 0:
        st.warning(f"⚠️ {sel_date_2} 没有可发货订单")
        st.stop()

    # ---- 面单核对 ----
    if label_pdf is not None:
        st.divider()
        st.subheader("🔍 面单核对结果")
        has_tess, has_popp = check_ocr_available()
        if not (has_tess and has_popp):
            st.error(
                "OCR 未启用：服务器缺少 tesseract / poppler-utils。"
                "本地：`brew install tesseract poppler` 或 `apt install tesseract-ocr poppler-utils`。"
                "Streamlit Cloud：在仓库根目录的 `packages.txt` 加上 `tesseract-ocr` 和 `poppler-utils`。"
            )
        else:
            with st.spinner("OCR 解析 PDF 中（每页约 1-2 秒）..."):
                try:
                    labels = parse_label_pdf(label_pdf.read())
                    rec = reconcile_labels(orders_2, labels)
                except Exception as e:
                    st.error(f"PDF 解析失败：{e}")
                    rec = None

            if rec is not None:
                # 统计
                ok = (rec["状态"] == "✅ 一致").sum()
                warn = rec["状态"].str.startswith("⚠️").sum()
                bad = rec["状态"].str.startswith("❌").sum()
                cc1, cc2, cc3 = st.columns(3)
                cc1.metric("✅ 一致", ok)
                cc2.metric("⚠️ 警告", warn)
                cc3.metric("❌ 不一致", bad)

                # 高亮表格
                def color_status(val):
                    if val.startswith("✅"): return "background-color: #d4edda"
                    if val.startswith("⚠️"): return "background-color: #fff3cd"
                    if val.startswith("❌"): return "background-color: #f8d7da"
                    return ""
                styled = rec.style.map(color_status, subset=["状态"])
                st.dataframe(styled, use_container_width=True, height=400)

                if bad == 0 and warn == 0:
                    st.success(f"🎉 所有 {len(rec)} 单面单核对通过，可以发货！")
                elif bad > 0:
                    st.error(f"⚠️ {bad} 单严重不一致，请人工核对再发货")

    # ---- 生成拣货单 + 发货单 ----
    st.divider()
    st.subheader("📦 生成拣货单 + 发货单")

    if catalog2 is not None:
        missing = exploded_2[(exploded_2["中文名"] == "") & (exploded_2["英文款式"] != "")]
        if len(missing) > 0:
            unmatched = missing["英文款式"].unique().tolist()
            st.warning(
                f"⚠️ {len(missing)} 行未在图册匹配到："
                + ", ".join(unmatched[:10]) + ("..." if len(unmatched) > 10 else "")
            )

    with st.expander("👀 拣货明细预览", expanded=False):
        st.dataframe(
            exploded_2[["Order ID","中文名","英文款式","SKU","Size","库位"]],
            use_container_width=True, height=300,
        )

    if st.button("✨ 生成拣货单 + 发货单", type="primary", use_container_width=True):
        with st.spinner("生成中..."):
            picking_csv = build_picking_summary_csv(exploded_2)
            slip_csv = build_packing_slip_csv(exploded_2, sel_date_2)
        st.success(f"✅ 已生成 {sel_date_2} 的发货文件")

        date_compact = sel_date_2.replace("/", "")
        cc1, cc2 = st.columns(2)
        with cc1:
            st.download_button(
                "📋 拣货单（按库位汇总）",
                data=picking_csv,
                file_name=f"拣货单_{date_compact}.csv",
                mime="text/csv", use_container_width=True,
            )
        with cc2:
            st.download_button(
                "📦 发货单 / Packing Slip",
                data=slip_csv,
                file_name=f"发货单_{date_compact}.csv",
                mime="text/csv", use_container_width=True,
            )
