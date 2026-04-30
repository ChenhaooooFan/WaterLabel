"""
NailVesta 发货系统
-----------------------------------------------------------
- 上传 Lark 水单 CSV、产品图册 CSV
- 选择洛杉矶日期（默认最晚日期）
- 生成 3 个 CSV：
  1. 拣货表（按库位排序，组合装拆成多行）
     列: 拣货顺序 / 库位 / 中文名 / 英文款式 / Size / Order ID / 收件人
  2. Packing Slip（按订单分组，含真实款式 + 收件人地址）
  3. 买 Label 的水单（每单一行，与你原 0430 模板字段一致）
"""

import io
import re
from datetime import datetime
from typing import Optional

import pandas as pd
import streamlit as st

# ============================================================
# 常量
# ============================================================
SENDER_INFO = {
    "name": "NailVesta",
    "company": None,
    "phone": "5105089943",
    "country": "US",
    "state": "CA",
    "city": "Los Angeles",
    "zip": "90071",
    "address": "515 S Flower St, Floor 18 & 19, STE 1901",
}

PACKAGE_DEFAULTS = {
    "weight": 0.3, "length": 20, "width": 15, "height": 2,
    "cn_name": "穿戴甲", "en_name": "Press-On Nails",
    "qty": 1, "declare_price": 5, "net_weight": 0.3,
}

SHUIDAN_HEADERS = [
    "客户订单号", "物流产品(产品编号)", "重量", "长", "宽", "高",
    "发件人姓名", "发件人公司", "发件人电话", "发件人国家", "发件人省/州",
    "发件人城市", "发件人邮编", "发件人地址",
    "收件人姓名", "收件人公司", "收件人电话", "收件人国家", "收件人省/州",
    "收件人城市", "收件人地址一", "收件人地址二", "收件人邮编",
    "中文品名1", "英文品名1", "SKU1", "数量1", "配货备注1",
    "申报单价1", "单位净重(kg)1",
]

SIZE_COL = "Size'"  # Lark 里带撇号的那列才是真 size

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
_STATE_ABBR = {
    "AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN","IA",
    "KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV","NH","NJ",
    "NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN","TX","UT","VT",
    "VA","WA","WV","WI","WY","DC","PR",
}
_STATE_FULL_TO_ABBR = {
    "ALABAMA":"AL","ALASKA":"AK","ARIZONA":"AZ","ARKANSAS":"AR","CALIFORNIA":"CA",
    "COLORADO":"CO","CONNECTICUT":"CT","DELAWARE":"DE","FLORIDA":"FL","GEORGIA":"GA",
    "HAWAII":"HI","IDAHO":"ID","ILLINOIS":"IL","INDIANA":"IN","IOWA":"IA",
    "KANSAS":"KS","KENTUCKY":"KY","LOUISIANA":"LA","MAINE":"ME","MARYLAND":"MD",
    "MASSACHUSETTS":"MA","MICHIGAN":"MI","MINNESOTA":"MN","MISSISSIPPI":"MS",
    "MISSOURI":"MO","MONTANA":"MT","NEBRASKA":"NE","NEVADA":"NV","NEW HAMPSHIRE":"NH",
    "NEW JERSEY":"NJ","NEW MEXICO":"NM","NEW YORK":"NY","NORTH CAROLINA":"NC",
    "NORTH DAKOTA":"ND","OHIO":"OH","OKLAHOMA":"OK","OREGON":"OR","PENNSYLVANIA":"PA",
    "RHODE ISLAND":"RI","SOUTH CAROLINA":"SC","SOUTH DAKOTA":"SD","TENNESSEE":"TN",
    "TEXAS":"TX","UTAH":"UT","VERMONT":"VT","VIRGINIA":"VA","WASHINGTON":"WA",
    "WEST VIRGINIA":"WV","WISCONSIN":"WI","WYOMING":"WY","DISTRICT OF COLUMBIA":"DC",
    "PUERTO RICO":"PR",
}


def state_to_abbr(state: str) -> str:
    if not state:
        return ""
    s = state.strip().upper()
    if s in _STATE_FULL_TO_ABBR:
        return _STATE_FULL_TO_ABBR[s]
    if s in _STATE_ABBR:
        return s
    return state if len(state) == 2 else state


def _is_zip(s: str) -> bool:
    return bool(re.fullmatch(r"\d{5}(-\d{4})?", s.strip()))


def _is_phone_line(s: str) -> bool:
    if re.match(r"^[\sa]*\(?\+?1?\)?[\s\-\.]?\(?\d{3}\)?[\s\-\.]?\d{3}[\s\-\.]?\d{4}", s.strip()):
        return True
    return bool(re.match(r"^(Tel|Phone|WhatsApp)\s*[:\uff1a]", s.strip(), re.I))


def _normalize_state(tok: str) -> Optional[str]:
    t = tok.strip().upper()
    if t in _US_STATES_FULL: return t
    if t in _STATE_ABBR: return t
    return None


def _clean_phone(s: str) -> str:
    digits = re.sub(r"\D", "", s)
    if digits.startswith("1") and len(digits) == 11:
        digits = digits[1:]
    return digits


def _clean_street(s: str) -> str:
    s = re.sub(r"\(([^)]*)\)", r" \1", s)
    s = re.sub(r"#\s+", "#", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def parse_shipping_info(text: str) -> dict:
    """鲁棒解析多种 Shipping Info 格式（4/5/6/7 行）"""
    if not isinstance(text, str): return {}
    lines = [ln.strip() for ln in text.replace("\r","").split("\n") if ln.strip()]
    while lines and re.search(r"[\u4e00-\u9fff]", lines[0]) and len(lines[0]) > 15:
        lines.pop(0)
    if not lines: return {}

    info = {"name":"", "phone":"", "street":"", "street2":"",
            "city":"", "state":"", "country":"United States", "zip":""}

    zip_idx = None
    for i in range(len(lines)-1, -1, -1):
        if _is_zip(lines[i]):
            zip_idx = i
            info["zip"] = lines[i].strip()
            break

    phone_idx = None
    for i, ln in enumerate(lines):
        if i == zip_idx: continue
        if _is_phone_line(ln):
            info["phone"] = _clean_phone(ln)
            phone_idx = i
            break

    csc_idx = None
    for i, ln in enumerate(lines):
        if i in (zip_idx, phone_idx): continue
        if "," in ln:
            parts = [p.strip() for p in ln.split(",")]
            if (any("UNITED STATES" in p.upper() or "USA" in p.upper() for p in parts)
                    and len(parts) >= 3):
                st_ = _normalize_state(parts[-2])
                if st_:
                    info["state"] = st_
                    info["country"] = parts[-1]
                    head = parts[0]
                    csc_idx = i
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
            info["name"] = clean
            name_pos = k
            break

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
        if street_lines:
            info["street2"] = " ".join(street_lines)

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
    keep = ["SKU", "中文名称", "款式英文名称", "甲型", "图片", "库位", "所属系列"]
    keep = [c for c in keep if c in df.columns]
    df = df[keep].copy()
    df["SKU"] = df["SKU"].astype(str).str.strip()
    if "款式英文名称" in df.columns:
        df["款式英文名称"] = df["款式英文名称"].astype(str).str.strip()
    return df


def filter_orders_for_date(df: pd.DataFrame, target_date: str) -> pd.DataFrame:
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
def _parse_size(size_raw: str) -> dict:
    """
    解析 Size' 列，返回 dict:
      {'_default': 'M'}                          标准单 size
      {'Winery':'L', 'Ribbon':'M', '_default':''} 多款式分别 size
    """
    s = str(size_raw).strip() if size_raw else ""
    if ";" in s:
        mapping = {"_default": s}
        for part in s.split(";"):
            part = part.strip()
            m = re.match(r"^(.+?)\s+(S|M|L|XL|XS|\d+\s*个)$", part, re.I)
            if m:
                mapping[m.group(1).strip()] = m.group(2).strip()
        return mapping
    return {"_default": s}


def explode_orders(orders: pd.DataFrame, catalog: Optional[pd.DataFrame]) -> pd.DataFrame:
    """
    按 Product Name 把组合装拆成多行。
    每行: Order ID / 英文款式 / SKU / 中文名 / 库位 / Size / Shipping Info / Full SKU
    """
    cat_by_name = None
    cat_by_sku = None
    if catalog is not None and not catalog.empty:
        if "款式英文名称" in catalog.columns:
            tmp = catalog.dropna(subset=["款式英文名称"]).copy()
            tmp["_key"] = tmp["款式英文名称"].astype(str).str.strip().str.lower()
            cat_by_name = tmp.set_index("_key")
        cat_by_sku = catalog.set_index("SKU")

    def lookup_by_name(name: str):
        if cat_by_name is None or not name: return None
        key = name.strip().lower()
        if key in cat_by_name.index:
            row = cat_by_name.loc[key]
            if isinstance(row, pd.DataFrame): row = row.iloc[0]
            return row
        return None

    def lookup_by_sku(sku: str):
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

            cn_name = ""
            cat_loc = ""
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
                "Order ID": order_id,
                "英文款式": cat_en_name,
                "SKU": sku_i,
                "中文名": cn_name,
                "库位": final_loc,
                "Size": size_i,
                "Shipping Info": ship_info,
                "Full SKU": str(r.get("Full SKU", "") or "").strip(),
            })

    return pd.DataFrame(rows)


# ============================================================
# 文件 1：拣货表 CSV
# ============================================================
def build_picking_csv(exploded: pd.DataFrame) -> bytes:
    """列: 拣货顺序 / 库位 / 中文名 / 英文款式 / Size / Order ID / 收件人"""
    rows = []
    for _, r in exploded.iterrows():
        ship = parse_shipping_info(r.get("Shipping Info", ""))
        rows.append({
            "库位": r["库位"] or "",
            "中文名": r["中文名"] or "",
            "英文款式": r["英文款式"] or "",
            "Size": r["Size"] or "",
            "Order ID": r["Order ID"],
            "收件人": ship.get("name", ""),
        })
    df = pd.DataFrame(rows)
    df["_loc_sort"] = df["库位"].apply(lambda x: (x == "", x))
    df = df.sort_values(by=["_loc_sort", "英文款式"]).drop(columns=["_loc_sort"]).reset_index(drop=True)
    df.insert(0, "拣货顺序", range(1, len(df) + 1))
    return df.to_csv(index=False).encode("utf-8-sig")


# ============================================================
# 文件 2：Packing Slip CSV
# ============================================================
def build_packing_slip_csv(exploded: pd.DataFrame, date_str: str) -> bytes:
    """每个订单一段，含完整收件人 + 所有商品行"""
    rows = []
    for order_id, grp in exploded.groupby("Order ID", sort=False):
        first = grp.iloc[0]
        ship = parse_shipping_info(first.get("Shipping Info", ""))
        addr_parts = [p for p in [
            ship.get("street", ""),
            ship.get("street2", ""),
            f"{ship.get('city','')}, {state_to_abbr(ship.get('state',''))} {ship.get('zip','')}".strip(", "),
            ship.get("country", "United States"),
        ] if p.strip(" ,")]
        full_address = " | ".join(addr_parts)

        for _, r in grp.iterrows():
            rows.append({
                "Order ID": order_id,
                "Date": date_str,
                "Recipient": ship.get("name", "").title() if ship.get("name") else "",
                "Phone": ship.get("phone", ""),
                "Address": full_address,
                "SKU": r["SKU"] or "",
                "Style": r["英文款式"] or "",
                "Chinese Name": r["中文名"] or "",
                "Size": r["Size"] or "",
                "Qty": 1,
            })
    df = pd.DataFrame(rows)
    return df.to_csv(index=False).encode("utf-8-sig")


# ============================================================
# 文件 3：水单 CSV（每订单一行）
# ============================================================
def build_shuidan_csv(orders: pd.DataFrame) -> bytes:
    rows = []
    for _, r in orders.iterrows():
        ship = parse_shipping_info(r.get("Shipping Info", ""))
        rows.append({
            "客户订单号": r["Order ID"],
            "物流产品(产品编号)": "",
            "重量": PACKAGE_DEFAULTS["weight"],
            "长": PACKAGE_DEFAULTS["length"],
            "宽": PACKAGE_DEFAULTS["width"],
            "高": PACKAGE_DEFAULTS["height"],
            "发件人姓名": SENDER_INFO["name"],
            "发件人公司": "",
            "发件人电话": SENDER_INFO["phone"],
            "发件人国家": SENDER_INFO["country"],
            "发件人省/州": SENDER_INFO["state"],
            "发件人城市": SENDER_INFO["city"],
            "发件人邮编": SENDER_INFO["zip"],
            "发件人地址": SENDER_INFO["address"],
            "收件人姓名": ship.get("name", ""),
            "收件人公司": "",
            "收件人电话": ship.get("phone", ""),
            "收件人国家": "US",
            "收件人省/州": state_to_abbr(ship.get("state", "")),
            "收件人城市": ship.get("city", ""),
            "收件人地址一": ship.get("street", ""),
            "收件人地址二": ship.get("street2", ""),
            "收件人邮编": ship.get("zip", ""),
            "中文品名1": PACKAGE_DEFAULTS["cn_name"],
            "英文品名1": PACKAGE_DEFAULTS["en_name"],
            "SKU1": "",
            "数量1": PACKAGE_DEFAULTS["qty"],
            "配货备注1": "",
            "申报单价1": PACKAGE_DEFAULTS["declare_price"],
            "单位净重(kg)1": PACKAGE_DEFAULTS["net_weight"],
        })
    df = pd.DataFrame(rows, columns=SHUIDAN_HEADERS)
    return df.to_csv(index=False).encode("utf-8-sig")


# ============================================================
# Streamlit UI
# ============================================================
st.set_page_config(page_title="NailVesta 发货系统", page_icon="💅", layout="wide")

st.title("💅 NailVesta 发货系统")
st.caption("Lark CSV + 图册 CSV → 一键生成 拣货表 / Packing Slip / 水单（CSV）")

with st.sidebar:
    st.header("📤 文件上传")
    lark_file = st.file_uploader("1️⃣ Lark 水单 CSV", type=["csv"])
    catalog_file = st.file_uploader("2️⃣ 产品图册 CSV", type=["csv"])

    st.divider()
    st.markdown("**📦 发件人信息**")
    st.caption(
        f"{SENDER_INFO['name']}\n{SENDER_INFO['address']}\n"
        f"{SENDER_INFO['city']}, {SENDER_INFO['state']} {SENDER_INFO['zip']}"
    )

if lark_file is None:
    st.info("👈 请先在左侧上传 Lark 水单 CSV")
    st.stop()

df = load_lark_data(lark_file.read())
all_dates = sorted(
    [d for d in df["日期"].dropna().unique()
     if re.match(r"^\d{4}/\d{1,2}/\d{1,2}$", str(d))],
    key=lambda x: datetime.strptime(x, "%Y/%m/%d"),
    reverse=True,
)

if not all_dates:
    st.error("CSV 中找不到合法日期")
    st.stop()

catalog = None
if catalog_file is not None:
    try:
        catalog = load_catalog(catalog_file.read())
        st.sidebar.success(f"✅ 图册已加载（{len(catalog)} 个 SKU）")
    except Exception as e:
        st.sidebar.error(f"图册解析失败: {e}")

col1, col2 = st.columns([1, 2])
with col1:
    st.subheader("📅 选择日期")
    st.caption("洛杉矶时间，默认最晚")
    selected_date = st.selectbox(
        "发货日期",
        options=all_dates,
        index=0,
        format_func=lambda x: f"{x} {'⬅️ 最新' if x == all_dates[0] else ''}",
    )

orders = filter_orders_for_date(df, selected_date)
exploded = explode_orders(orders, catalog)

with col2:
    st.subheader("📊 当日订单概况")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("订单数", len(orders))
    extra = len(exploded) - len(orders)
    m2.metric("拣货行数", len(exploded), delta=f"+{extra} 组合装" if extra > 0 else None)
    m3.metric("SKU 种类", exploded["SKU"].replace("", pd.NA).dropna().nunique() if len(exploded) else 0)
    m4.metric("拣货库位", exploded["库位"].replace("", pd.NA).dropna().nunique() if len(exploded) else 0)

if len(orders) == 0:
    st.warning(f"⚠️ {selected_date} 没有可发货订单")
    st.stop()

if catalog is not None:
    missing = exploded[(exploded["中文名"] == "") & (exploded["英文款式"] != "")]
    if len(missing) > 0:
        unmatched = missing["英文款式"].unique().tolist()
        st.warning(
            f"⚠️ {len(missing)} 行未在图册匹配到："
            + ", ".join(unmatched[:10])
            + ("..." if len(unmatched) > 10 else "")
        )
elif catalog is None:
    st.info("💡 未上传产品图册，无法显示中文名和准确库位")

with st.expander("👀 拆单后的拣货明细预览", expanded=True):
    preview = exploded[["Order ID", "中文名", "英文款式", "SKU", "Size", "库位"]].copy()
    st.dataframe(preview, use_container_width=True, height=350)

st.divider()

st.subheader("🚀 生成发货文件（CSV）")
if st.button("✨ 一键生成 3 个 CSV", type="primary", use_container_width=True):
    with st.spinner("生成中..."):
        date_compact = selected_date.replace("/", "")
        picking_csv = build_picking_csv(exploded)
        slip_csv = build_packing_slip_csv(exploded, selected_date)
        shuidan_csv = build_shuidan_csv(orders)

    st.success(f"✅ 已生成 {selected_date} 的 {len(orders)} 单（拣货 {len(exploded)} 行）")

    c1, c2, c3 = st.columns(3)
    with c1:
        st.download_button(
            "📋 1. 拣货表",
            data=picking_csv,
            file_name=f"拣货表_{date_compact}.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with c2:
        st.download_button(
            "📦 2. Packing Slip",
            data=slip_csv,
            file_name=f"PackingSlip_{date_compact}.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with c3:
        st.download_button(
            "🚚 3. 水单（买 Label）",
            data=shuidan_csv,
            file_name=f"{date_compact}客人水单.csv",
            mime="text/csv",
            use_container_width=True,
        )
