"""The two charge-detail sheets, laid out the way the accountant's workbook lays them out.

`报告明细` is ML's Facturacion export with a working block bolted on the right, and
`贷记明细` is the credit-note export left as ML publishes it. Both are rebuilt from the
downloads every run -- nothing is copied out of anyone's workbook -- so they come out the
same shape for any store.

Every added column is derived, and each derivation was checked against the accountant's
own file for MX-TA02:

``年月``      the charge month (412/412)
``费用项目``  ``Detalle`` mapped to their Chinese category names
``标志``      a short code per category: YJ commission, YF shipping, CC storage, GG ads
``SKU``       the listing's SKU (304/304), via the sales export rather than a hand master
``引用R``     ``Número de venta`` without ML's constant ``2000`` prefix
``引用AB``    ``Número de paquete``, same treatment
``重次``      the charge's position within its order (304/304)
``税金``      ``Total de la venta / 1.16 * 0.105`` (304/304)
``调减费用``  ``Y`` where a credit note voided the charge and it belongs to a product
``抵扣年月``  the month that deduction lands in
"""

from __future__ import annotations

import pandas as pd

from mx_sales.reports.facturacion import CREDIT_NOTE_VOID, categorize_fee

#: Canonical fee key -> the accountant's category name. ``pickup`` folds into storage and
#: reversals of an account-level fee keep the parent's name, which is how they group them.
FEE_CATEGORY_CHINESE = {
    "commission": "销售费",
    "commission_reversal": "取消销售费",
    "shipping": "Mercado Libre 运费",
    "shipping_reversal": "取消 Mercado Libre 运费",
    "return_fee": "退款费用",
    "advertising": "广告费用",
    "storage": "仓储服务费",
    "pickup": "仓储服务费",
    "stock_withdrawal": "退仓费用",
    "maintenance": "页面维护费",
    "maintenance_reversal": "页面维护费",
    "reputation_guarantee": "其他费用",
    "other": "其他费用",
}

#: Category -> the short flag they pivot on.
CATEGORY_FLAG = {
    "销售费": "YJ",
    "取消销售费": "YJ",
    "Mercado Libre 运费": "YF",
    "取消 Mercado Libre 运费": "YF",
    "退款费用": "YF",
    "仓储服务费": "CC",
    "广告费用": "GG",
    "页面维护费": "Y",
    "退仓费用": "CC",
    "其他费用": "",
}

#: ML prefixes every order and pack id with this. Stripping it is what their ``引用``
#: columns do, and it makes the two id families comparable at a glance.
ID_PREFIX = "2000"

ADDED_COLUMNS = [
    "费用项目", "标志", "SKU", "引用R", "引用AB", "重次", "税金", "调减费用", "抵扣年月",
]


def charge_detail(
    charges: pd.DataFrame, listing_to_sku: pd.Series, store: str | None = None
) -> pd.DataFrame:
    """Build `报告明细`: the Facturacion export plus the accountant's working columns.

    ``charges`` must carry ML's own Spanish headers (see
    :func:`mx_sales.reports.facturacion.load_display`), so the sheet reads like the export
    it came from.
    """
    if charges is None or charges.empty:
        return pd.DataFrame(columns=["年月", "店铺", "币别", *ADDED_COLUMNS])

    out = charges.copy().reset_index(drop=True)
    charged = pd.to_datetime(_column(out, "Fecha del cargo"), errors="coerce")

    out.insert(0, "币别", "MXN")
    out.insert(0, "店铺", store or "")
    out.insert(0, "年月", charged.dt.strftime("%Y%m"))

    category = _column(out, "Detalle").map(categorize_fee).map(FEE_CATEGORY_CHINESE)
    out["费用项目"] = category
    out["标志"] = category.map(CATEGORY_FLAG)

    listing = _column(out, "Número de publicación")
    out["SKU"] = listing.map(listing_to_sku) if len(listing_to_sku) else pd.NA

    out["引用R"] = _strip_prefix(_column(out, "Número de venta"))
    out["引用AB"] = _strip_prefix(_column(out, "Número de paquete"))
    out["重次"] = _charge_sequence(out["引用R"])

    sale_total = pd.to_numeric(_column(out, "Total de la venta"), errors="coerce")
    out["税金"] = (sale_total / 1.16 * 0.105).round(2)

    voided = _column(out, "Estado del cargo").eq(CREDIT_NOTE_VOID).fillna(False)
    out["调减费用"] = pd.Series("", index=out.index).mask(voided & out["SKU"].notna(), "Y")
    out["抵扣年月"] = pd.Series(pd.NA, index=out.index, dtype="object").mask(
        voided, out["年月"]
    )
    return out


def credit_note_detail(notes: pd.DataFrame) -> pd.DataFrame:
    """Build `贷记明细`: the credit-note export, untouched.

    The accountant adds nothing to this one -- no period, no store, no categories -- and
    neither do we. It is reference material behind the credit-note totals, and the same
    events already reach the fee logic through ``Estado del cargo`` on the Facturacion
    side, so re-deriving them here would only invite the two copies to disagree.
    """
    if notes is None or notes.empty:
        return pd.DataFrame()
    return notes.copy().reset_index(drop=True)


def _column(frame: pd.DataFrame, name: str) -> pd.Series:
    """A column by ML's Spanish name, or an all-null series when the export omits it."""
    if name in frame.columns:
        return frame[name]
    return pd.Series(pd.NA, index=frame.index, dtype="object")


def _strip_prefix(series: pd.Series) -> pd.Series:
    """Render an id without ML's constant ``2000`` prefix.

    The result is numeric, which also drops the zero that sits behind the prefix --
    ``2000017459121180`` becomes ``17459121180``, matching how the accountant writes it.
    """
    digits = pd.to_numeric(series, errors="coerce").astype("Int64").astype("string")
    trimmed = digits.str.removeprefix(ID_PREFIX)
    return pd.to_numeric(trimmed, errors="coerce").astype("Int64")


def _charge_sequence(order_ref: pd.Series) -> pd.Series:
    """Which charge this is within its order, counting in the order ML printed them."""
    sequence = order_ref.groupby(order_ref).cumcount() + 1
    return sequence.where(order_ref.notna()).astype("Int64")
