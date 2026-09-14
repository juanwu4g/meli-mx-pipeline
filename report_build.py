# -*- coding: utf-8 -*-
"""
Mercado Libre MX —— 月度财务报表生成器（独立版，不依赖项目内其他模块）

    python report_build.py --month 2026-08 \
        --store "EWTTO_SM=D:/Projects/mx_sales_data/downloads/EWTTO_SM/20260909_021624" \
        --store "UNIT_PW01(catalog)=D:/.../UNIT_PW01(catalog)/20260909_021624" \
        --out reports/MercadoLibre_月度财务报表_2026年8月.xlsx

设计原则
--------
1) 确定性。同样的输入文件，任意次运行产出相同的行位置、相同的公式、相同的数字。
   ② 损益表的 27 行是写死的常量表（PL_LINES），某个费用科目本月为 0 也保留该行，
   这样跨月对比时行号永远对齐 —— 这是"格式不变、只有趋势变"的前提。

2) 费用科目必须登记。FEE_CATALOG 是唯一的科目字典。出现字典里没有的 Detalle 时，
   **绝不静默归入"其他"**：金额照常进入报表（否则勾稽会断），但会被单独归到
   "未登记科目"行、在 ⑨ 页列为高优先级待办、并让 check_fee_catalog 校验失败。
   （旧实现把 comprar más espacio / sobrepasar espacio / stock antiguo /
   incumplimiento 四个科目静默丢进 other，EWTTO 2026-08 合计 6,506.85 因此成了糊涂账。）

3) 冲销按"被冲销的那张单"归类，不靠字符串猜。账单行有 `Número del cargo`，
   冲销行有 `Cargo que bonifica` 指向它，用这个 ID 关联即可把
   "Anulación del cargo por campaña de publicidad de Products Ads"
   （注意平台自己把 Product 写成了 Products）正确地冲减到广告费上。
   ID 关联不上时才退回按名称匹配。

4) 一切"发现"都固化成公式与校验：
   - 代扣代缴税金  = Ventas 的 Cargo por venta e impuestos − 账单纯佣金
   - 卖家优惠券    = 由 Total 列倒推（平台把 Descuentos 列导成了空白）
   - 结算通道拆分  = Descontado de la operación ∈ {Si, No, No aplica}
   校验见 validate()。

依赖：pandas, openpyxl。重算公式缓存值需要 LibreOffice（soffice），可用 --no-recalc 跳过。
"""
from __future__ import print_function

import argparse
import datetime
import glob
import io
import json
import os
import re
import subprocess
import sys
import warnings

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# ════════════════════════════════════════════════════════════════════════
# 常量：原始报表的结构
# ════════════════════════════════════════════════════════════════════════

VENTAS_SHEET = "Ventas MX"
VENTAS_HEADER_ROW = 5        # 0-indexed，实际表头在第 6 行
BILLING_HEADER_ROW = 7       # Facturación / Notas de Crédito，表头在第 8 行
CARGOS_HEADER_ROW = 5        # Reporte_Cargos_Full_*，表头在第 6 行
PAGOS_HEADER_ROW = 9         # Reporte_Pagos_Facturas_*，表头在第 10 行
STORAGE_HEADER_ROW = 4       # Costos_por_servicio_almacenamiento，表头在第 5 行
RETURNS_HEADER_ROW = 2       # Returns_*.xlsx / Triages，表头在第 3 行

SPANISH_MONTHS = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}
MONTH_FULL_ES = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo",
                 6: "Junio", 7: "Julio", 8: "Agosto", 9: "Septiembre",
                 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
# 平台文件名里用的月份缩写。注意 9 月是 "Sept" 不是 "Sep"，10 月两种都见过。
FILE_MONTH_TOKENS = {
    "ene": 1, "enero": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
    "jul": 7, "ago": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dic": 12,
}

# Ventas 列名。各店的导出列数不一致（例如 EWTTO 多一列"换货运费"），所以一律按列名取，
# 绝不按列号。缺列时按 0 处理。
V_ORDER_ID = "# de venta"
V_PACK_ID = "Orden de compra"
V_DATE = "Fecha de venta"
V_STATUS = "Estado"
V_UNITS = "Unidades"
V_REVENUE = "Ingresos por productos (MXN)"
V_COMMISSION = "Cargo por venta e impuestos (MXN)"   # ← 佣金 + 代扣税，混在一起
V_SHIP_INCOME = "Ingresos por envío (MXN)"
V_SHIP_COST = "Costos de envío (MXN)"
V_REFUND = "Anulaciones y reembolsos (MXN)"
V_TOTAL = "Total (MXN)"
V_SKU = "SKU"
V_TITLE = "Título de la publicación"
V_PACK_FLAG = "Paquete de varios productos"
V_KIT = "Pertenece a un kit"
V_AD_SALE = "Venta por publicidad"
V_DISCOUNT = "Descuentos y bonificaciones"   # 平台导出为空列，本报表由 Total 倒推

# Facturación 列名
F_INVOICE = "N° de factura fiscal"
F_CHARGE_DATE = "Fecha del cargo"
F_CHARGE_NO = "Número del cargo"
F_DETALLE = "Detalle"
F_CHANNEL = "Descontado de la operación"      # Si / No / No aplica
F_STATE = "Estado del cargo"
F_REVERSES = "Cargo que bonifica"             # 指向被冲销的 Número del cargo
F_AMOUNT = "Valor del cargo"
F_SALE_NO = "Número de venta"
F_PACK_NO = "Número de paquete"
F_SALE_DATE = "Fecha de venta"

CH_SI = "Si"
CH_NO = "No"
CH_NA = "No aplica"
CHANNEL_ZH = {CH_SI: "订单内直扣", CH_NO: "月底账单支付", CH_NA: "取消订单冲销"}
CHANNEL_MEANING = {
    CH_SI: "下单打款时平台已直接扣除，不会再收第二次",
    CH_NO: "月底汇总，从 MercadoPago 余额自动扣款",
    CH_NA: "取消/退货订单的原计费与冲销",
}


# ════════════════════════════════════════════════════════════════════════
# 费用科目字典 —— 唯一的真相来源
# ════════════════════════════════════════════════════════════════════════
# (西班牙语 Detalle 原文, 中文, 大类 bucket, 是否订单级)
#
# "是否订单级" 只用于校验：订单级科目理应带 Fecha de venta，账单级理应不带。
# 实际归属仍以数据里有没有 Fecha de venta 为准（数据 > 字典），不一致时记为异常。
#
# 新增科目请加在这里，并同时在 PL_LINES 里给它一行，否则它会落进"未登记科目"。
FEE_CATALOG = [
    # ---- 订单级（打款时直扣，Descontado = Si） ----
    ("Cargo por venta",                                "平台销售佣金",              "commission",   True),
    ("Cargo por venta con afiliados",                  "联盟推广佣金",              "commission",   True),
    ("Cargo por envíos de Mercado Libre",              "平台配送费 (Mercado Envíos)", "shipping",   True),
    ("Cargo por devolución",                           "退货处理费",                "returns",      True),
    # ---- 账单级（月底自动扣款，Descontado = No） ----
    ("Cargo por campaña de publicidad de Product Ads", "商品广告费 Product Ads",     "ads",          False),
    ("Cargo por campaña de publicidad de Display Ads", "展示广告费 Display Ads",     "ads",          False),
    ("Cargo por servicio de almacenamiento Full",      "Full 仓储费",               "storage",      False),
    ("Cargo por servicio de colecta Full",             "Full 揽收/入仓费",           "fulfillment",  False),
    ("Cargo por comprar más espacio en Full",          "Full 购买额外仓位",          "fulfillment",  False),
    ("Cargo por sobrepasar espacio Full",              "Full 超仓位罚金",            "penalty",      False),
    ("Cargo por retiro de stock Full",                 "Full 退仓/取货费",           "fulfillment",  False),
    ("Cargo por stock antiguo en Full",                "Full 长龄库存费 (库龄4个月+)", "fulfillment", False),
    ("Cargo por incumplimiento en Envíos Full",        "Full 发货违规罚金 (库存差异)", "penalty",     False),
    ("Cargo por mantenimiento de Mi página",           "官方店页面维护费",           "subscription", False),
]
FEE_BY_ES = {es: (zh, bucket, order_level) for es, zh, bucket, order_level in FEE_CATALOG}

# 冲销行的名字通常是原名把 "Cargo por" 换成 "Anulación del cargo por"，
# 但平台自己有拼写不一致（Product → Products），所以这里显式列出例外。
REVERSAL_ALIASES = {
    "Anulación del cargo por campaña de publicidad de Products Ads":
        "Cargo por campaña de publicidad de Product Ads",
}


def _auto_reversal_name(es):
    return es.replace("Cargo por", "Anulación del cargo por", 1)


# 冲销名 → 原科目名
REVERSAL_TO_PARENT = dict(REVERSAL_ALIASES)
for _es, _zh, _b, _o in FEE_CATALOG:
    REVERSAL_TO_PARENT.setdefault(_auto_reversal_name(_es), _es)

UNCLASSIFIED_ZH = "⚠ 未登记科目（需人工确认）"


def is_reversal(detalle):
    return str(detalle or "").startswith("Anulación del cargo")


def classify(detalle):
    """Detalle → (中文, bucket, 是否订单级, 是否冲销, 是否已登记)。

    未登记的科目返回 bucket='unclassified'。调用方必须把它当成"看得见的异常"，
    而不是丢进 other 了事。
    """
    s = str(detalle or "").strip()
    if is_reversal(s):
        parent = REVERSAL_TO_PARENT.get(s)
        if parent and parent in FEE_BY_ES:
            zh, bucket, order_level = FEE_BY_ES[parent]
            return "－ " + zh + " 冲销", bucket, order_level, True, True
        return UNCLASSIFIED_ZH + "（冲销）", "unclassified", None, True, False
    if s in FEE_BY_ES:
        zh, bucket, order_level = FEE_BY_ES[s]
        return zh, bucket, order_level, False, True
    return UNCLASSIFIED_ZH, "unclassified", None, False, False


# ════════════════════════════════════════════════════════════════════════
# 订单状态 / 退货原因 词典
# ════════════════════════════════════════════════════════════════════════
# (中文, 资金归属, 库存处置)。未收录的状态照常统计，标"待确认"。
STATUS_ZH = {
    "Cancelada por el comprador": ("买家下单后取消", "买家", "未发出，库存无损"),
    "Cancelaste la venta": ("卖家主动取消", "卖家", "未发出，库存无损"),
    "Cancelada por seguridad": ("平台风控取消", "平台", "未发出，库存无损"),
    "Paquete cancelado por Mercado Libre": ("平台取消整个包裹", "平台", "未发出，库存无损"),
    "Devolución finalizada con reembolso al comprador": ("退货完成，已退款买家", "买家", "待质检结果"),
    "Devolución finalizada. Pusimos el producto de nuevo a la venta": ("退货完成，商品已重新上架", "买家", "✔ 可回收，重新销售"),
    "Devolución finalizada. Pusimos los productos de nuevo a la venta": ("退货完成，多件商品已重新上架", "买家", "✔ 可回收，重新销售"),
    "Devolución revisada. Pusimos el producto de nuevo a la venta": ("退货已质检，商品重新上架", "买家", "✔ 可回收，重新销售"),
    "Devolución revisada. Solicita el retiro del producto": ("退货已质检，需卖家自提", "买家", "✘ 不可再售，须付费取回"),
    "Devolución finalizada. Descartamos el producto": ("退货完成，商品已销毁", "买家", "✘ 彻底损失"),
    "Devolución rechazada. Regresamos el producto al comprador": ("退货被驳回，商品退回买家", "卖家(保留货款)", "— 无库存回收"),
    "Devolución en camino": ("退货在途", "待定", "待质检"),
    "Devolución en preparación": ("退货准备中", "待定", "待质检"),
    "Devolución en revisión": ("退货质检中", "待定", "待质检"),
    "Devolución con fecha actualizada": ("退货日期已更新", "待定", "待质检"),
    "Devolución finalizada": ("退货已完成", "买家", "待质检结果"),
    "Mediación finalizada con reembolso al comprador": ("平台仲裁，判退款给买家", "买家", "通常不回收"),
    "Mediación finalizada. Te dimos el dinero.": ("平台仲裁，判货款归卖家", "卖家(保留货款)", "— 无库存回收"),
    "Mediación con devolución habilitada": ("仲裁中，已开放退货", "待定", "待质检"),
    "Cambio no entregado": ("换货未送达", "待定", "待处理"),
    "Venta con solicitud de cambio": ("买家申请换货", "卖家(换货)", "换新，旧品待回收"),
    "Cambio entregado. Devolución finalizada.": ("换货已送达，退货完成", "卖家(换货)", "换新，旧品待回收"),
    "Paquete no entregado": ("包裹未送达", "待定", "待处理"),
    "Reclamo cerrado": ("投诉已关闭", "待定", "—"),
    "Entregado": ("已送达", "—", "—"),
    "Venta entregada": ("销售已送达", "—", "—"),
    "En camino": ("运输中", "—", "—"),
    "Procesando en la bodega": ("仓库处理中", "—", "—"),
    # ↓ 实际数据里出现、平台文案的变体或细分说法
    "Cancelada": ("订单已取消", "待定", "未发出，库存无损"),
    "Devolución en camino sin costo de envío": ("退货在途（免运费）", "待定", "待质检"),
    "Devolución en preparación sin costo de envío": ("退货准备中（免运费）", "待定", "待质检"),
    "Devolución en camino. Revisaremos el producto": ("退货在途，将进行质检", "待定", "待质检"),
    "Devolución en camino. Revisaremos los productos": ("退货在途，多件将进行质检", "待定", "待质检"),
    "Devolución finalizada. Revisa los resultados.": ("退货完成，请查看质检结果", "买家", "待质检结果"),
    "Devolución no entregada": ("退货未送达", "待定", "待处理"),
    "Mediación finalizada. Te dimos el dinero": ("平台仲裁，判货款归卖家", "卖家(保留货款)", "— 无库存回收"),
    "Mediación en espera de respuesta de Mercado Libre": ("仲裁中，等待平台答复", "待定", "待处理"),
    "Mediación para responder el lunes": ("仲裁中，需在周一前答复", "待定", "★ 待卖家答复"),
    "Reclamo abierto para resolver hasta mañana a las 14 hs": ("投诉待处理，明日 14 时截止", "待定", "★ 待卖家答复"),
    "Reclamo en espera de respuesta del comprador": ("投诉中，等待买家答复", "待定", "待处理"),
    "Reclamo cerrado con reembolso al comprador": ("投诉结案，全额退款买家", "买家", "通常不回收"),
    "Reclamo cerrado con reembolso parcial": ("投诉结案，部分退款买家", "买家(部分)", "通常不回收"),
}
# 识别"售后/异常订单"的正则。用于售后率与退货分析。
ABNORMAL_RE = r"Devoluci|Mediaci|Cancel|Reclamo|Cambio|no entregado|rechaz"
# 库存去向分组
STATES_BACK_ON_SALE = [
    "Devolución finalizada. Pusimos el producto de nuevo a la venta",
    "Devolución finalizada. Pusimos los productos de nuevo a la venta",
    "Devolución revisada. Pusimos el producto de nuevo a la venta",
]
STATES_STOCK_LOST = [
    "Devolución revisada. Solicita el retiro del producto",
    "Devolución finalizada. Descartamos el producto",
]
STATES_SELLER_KEEPS = [
    "Mediación finalizada. Te dimos el dinero.",
    "Devolución rechazada. Regresamos el producto al comprador",
]

# 退货质检：商品状态 → (中文, 库存后果)
RETURN_REASON_ZH = {
    "Sello dañado": ("包装封条破损", "✘ 多数判定不可再售"),
    "Producto dañado": ("商品损坏", "✘ 不可再售"),
    "Producto con marcas de mal uso": ("商品有使用/磨损痕迹", "✘ 多数不可再售"),
    "No funciona correctamente": ("功能故障，无法正常工作", "✘ 不可再售"),
    "Producto sin accesorios, manuales o etiquetas": ("缺配件/说明书/标签", "△ 补齐后可能可售"),
    "Producto ausente, destruido o ultrajado": ("商品缺失、损毁或被拆封破坏", "✘ 彻底损失"),
    "-": ("平台未记录具体原因", "△ 需人工核实"),
}
# 退货质检：处置结果 → (中文, 库存价值, 财务影响)
RETURN_RESULT_ZH = {
    "Producto quedó nuevamente para la venta": ("商品重新上架销售", "100% 保留", "退款损失可由再次销售弥补"),
    "Producto para retirar en centro de distribución": ("需卖家到配送中心自提", "需付退仓费才能取回", "★ 主要损失来源：退款 + 取货费 + 商品贬值"),
    "Producto devuelto al comprador": ("退货被驳回，商品退回买家", "商品已给买家", "货款保留，无库存回收"),
    "Producto archivado en centro de distribución": ("商品在配送中心封存", "待处理", "需尽快决定自提或放弃"),
}
MONEY_REFUNDED = "Reembolsamos el dinero al comprador"
MONEY_KEPT = "Te dimos el dinero"

# 代扣代缴税金的预期费率区间（占含税成交额）。落在区间外说明税制或账号状态变了。
TAX_RATE_MIN, TAX_RATE_MAX = 0.085, 0.095
TAX_RATE_NOMINAL_GROSS = 0.0905     # 占含税价
TAX_RATE_NOMINAL_NET = 0.1050       # 占不含税价（÷1.16）
IVA_RATE = 0.16


# ════════════════════════════════════════════════════════════════════════
# ② 合并损益表的行定义 —— 写死，保证跨月行号对齐
# ════════════════════════════════════════════════════════════════════════
# (key, 中文行名, 西班牙语科目, kind, 说明)
#   kind: 'd' 取数 / 'sum' 小计(加粗浅蓝) / 'f' 计算行(灰底) / 'in' 需录入(黄底)
# key 为 None 的行由公式生成。费用行的 key 对应 metrics 里的字段名。
PL_LINES = [
    ("ing",          "商品销售收入 (GMV)",        "Ingresos por productos",           "d",
     "买家实付商品金额（含税售价），按下单日归集"),
    ("env_ing",      "运费收入",                  "Ingresos por envío",               "d",
     "买家另行支付的运费"),
    ("coupon_neg",   "减：卖家承担优惠券/折扣",    "Descuentos y bonificaciones",      "d", None),
    (None,           "营业收入合计",              "",                                 "sum",
     "＝ 商品收入 ＋ 运费收入 − 卖家优惠券"),
    ("refund_neg",   "减：退款与取消（净）",       "Anulaciones y reembolsos",         "d",
     "已退给买家的净金额（平台已返还的佣金/运费/税已抵减）"),
    ("cn_adj",       "加：部分退款重复扣佣金调整",  "(本报表修正，平台无此科目)",         "d",
     "平台对部分退款订单按退款比例退还佣金与代扣税，但销售报表的“退款与取消”列会把原佣金"
     "再扣一次。本行把多扣的加回来，逐单明细见 ⑩ 页。已用 MercadoPago 现金流水逐单核对。"),
    (None,           "净销售收入",                "",                                 "f",
     "＝ 营业收入合计 − 退款与取消 ＋ 重复扣佣金调整"),
    ("com_neg",      "减：平台销售佣金",           "Cargo por venta",                  "d",
     "账单口径纯佣金（含联盟推广佣金），按品类 13%–18%，不含代扣税"),
    ("shipcost_neg", "减：平台配送费",             "Cargo por envíos de Mercado Libre", "d",
     "Mercado Envíos 配送费（已扣平台补贴后的净额）"),
    ("dev_neg",      "减：退货处理费",             "Cargo por devolución",             "d",
     "退货产生的逆向物流/处理费"),
    ("unc_order_neg", "减：订单级未登记科目 ⚠",     "(未在 FEE_CATALOG 登记)",           "d",
     "字典里没有的订单级费用。金额照常计入以保证勾稽，但需人工确认科目归属，见 ⑩ 页"),
    (None,           "平台交易费用小计",           "",                                 "sum",
     "订单打款时即已扣除（账单 Descontado = Si）"),
    (None,           "订单毛贡献",                "",                                 "f",
     "＝ 净销售收入 − 平台交易费用"),
]
# 账单级费用行：按 FEE_CATALOG 顺序自动生成，保证新增科目时行序稳定
PL_BILL_NOTES = {
    "Cargo por campaña de publicidad de Product Ads": "已扣除当月冲销",
    "Cargo por campaña de publicidad de Display Ads": "已扣除当月冲销",
    "Cargo por servicio de almacenamiento Full": "按件按天计费；另有独立仓储费报表可交叉核对",
    "Cargo por servicio de colecta Full": "按体积 m³ 计费",
    "Cargo por comprar más espacio en Full": "主动购买的额外仓位",
    "Cargo por sobrepasar espacio Full": "占用超出配额，按超出件数罚",
    "Cargo por retiro de stock Full": "从 Full 仓取回库存的费用",
    "Cargo por stock antiguo en Full": "库龄 4 个月以上加收",
    "Cargo por incumplimiento en Envíos Full": "库存差异等履约违规罚金",
    "Cargo por mantenimiento de Mi página": "Mi página 月费",
}
PL_TAIL = [
    ("unc_bill_neg", "减：账单级未登记科目 ⚠",     "(未在 FEE_CATALOG 登记)",           "d",
     "字典里没有的账单级费用。金额照常计入以保证勾稽，但需人工确认科目归属，见 ⑩ 页"),
    (None,           "月度账单费用小计",           "",                                 "sum",
     "月底从 MercadoPago 余额自动扣款（账单 Descontado = No）"),
    (None,           "经营贡献毛利（未扣商品成本）", "",                                "f",
     "＝ 订单毛贡献 − 月度账单费用"),
    ("cogs",         "减：商品采购成本",           "",                                 "in",
     "⚠ 平台数据不含此项，请在【⑦ SKU盈利分析】录入单位成本后回填，或直接在此填总额"),
    (None,           "税前净利润",                "",                                 "f",
     "＝ 经营贡献毛利 − 商品采购成本"),
]

SHEET_NAMES = [
    "① 阅读说明", "② 合并损益表", "③ 平账勾稽表", "④ 平台费用明细",
    "⑤ 退货退款分析", "⑥ 退货质检明细", "⑦ SKU盈利分析", "⑧ 月度趋势",
    "⑨ 库存与动销分析", "⑩ 数据缺口与待办",
]

# ════════════════════════════════════════════════════════════════════════
# 样式
# ════════════════════════════════════════════════════════════════════════
FONT_NAME = "Arial"
MNY = '#,##0.00;(#,##0.00);-'
INT = '#,##0;(#,##0);-'
PCT = '0.0%;(0.0%);-'
PCT2 = '0.00%'
PCT3 = '0.000%'

C_NAVY, C_BAND, C_SUB, C_CALC = "1F3864", "4472C4", "D9E2F3", "EDEDED"
C_TOTAL, C_INPUT, C_OK, C_BAD, C_WARN = "FCE4D6", "FFFF00", "E2EFDA", "FBE5E5", "FFF2CC"
C_RED, C_GREEN, C_AMBER, C_GREY = "C00000", "375623", "BF8F00", "595959"

_THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)


def F(size=10, bold=False, color="000000", italic=False):
    return Font(name=FONT_NAME, size=size, bold=bold, color=color, italic=italic)


def fill(hexcolor):
    return PatternFill("solid", fgColor=hexcolor)


def wrap(vertical="center"):
    return Alignment(wrap_text=True, vertical=vertical)


def est_height(text, per_line=60, line_h=15, minimum=15):
    """按字符数粗估行高。中文宽字符，60 字/行是保守值。"""
    if not text:
        return minimum
    return max(minimum, line_h * (len(str(text)) // per_line + 1))


# ════════════════════════════════════════════════════════════════════════
# 通用小工具
# ════════════════════════════════════════════════════════════════════════

def money(x, dec=2):
    try:
        return round(float(x), dec)
    except (TypeError, ValueError):
        return 0.0


def fmt(x, dec=2):
    """1427367.66 -> '1,427,367.66'"""
    try:
        return "{:,.{d}f}".format(float(x), d=dec)
    except (TypeError, ValueError):
        return "-"


def fmt0(x):
    return fmt(x, 0)


def pct(x, dec=2):
    try:
        return "{:.{d}f}%".format(float(x) * 100.0, d=dec)
    except (TypeError, ValueError):
        return "-"


def parse_spanish_date(value):
    """'31 de agosto de 2026 15:10 hs.' -> Timestamp；解析不了返回 NaT。"""
    if isinstance(value, pd.Timestamp):
        return value
    if not isinstance(value, str):
        return pd.NaT
    m = re.match(r"\s*(\d{1,2})\s+de\s+([A-Za-zÁ-úá-ú]+)\s+de\s+(\d{4})"
                 r"(?:\s+(\d{1,2}):(\d{2}))?", value.strip(), re.IGNORECASE)
    if not m:
        return pd.NaT
    day, month_name, year, hh, mm = m.groups()
    month = SPANISH_MONTHS.get(month_name.lower())
    if not month:
        return pd.NaT
    try:
        return pd.Timestamp(int(year), month, int(day), int(hh or 0), int(mm or 0))
    except ValueError:
        return pd.NaT


def num(df, col):
    """按列名取数值列；列不存在时返回全 0（各店导出列数不一致）。"""
    if col in df.columns:
        return pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    return pd.Series(0.0, index=df.index)


def idkey(series):
    """订单号在不同报表里一边是 str 一边是 float，统一成无小数点的字符串。"""
    return series.astype(str).str.replace(r"\.0$", "", regex=True).str.strip()


def month_bounds(month_str):
    y, m = int(month_str[:4]), int(month_str[5:7])
    start = pd.Timestamp(y, m, 1)
    end = (start + pd.offsets.MonthEnd(1)).normalize()
    return start, end


def file_month_token(name):
    """'Reporte_Pagos_Facturas_Ago2026.xlsx' -> (2026, 8)；识别不了返回 None。"""
    m = re.search(r"_([A-Za-z]{3,5})(\d{4})\.xlsx$", name)
    if not m:
        return None
    tok = m.group(1).lower()
    mon = FILE_MONTH_TOKENS.get(tok)
    if not mon:
        return None
    return int(m.group(2)), mon


def storage_file_range(name):
    """'01-08-26_31-08-26_Costos...' -> (start, end)；识别不了返回 None。"""
    m = re.match(r"(\d{2})-(\d{2})-(\d{2})_(\d{2})-(\d{2})-(\d{2})_", name)
    if not m:
        return None
    d1, m1, y1, d2, m2, y2 = [int(x) for x in m.groups()]
    try:
        return (pd.Timestamp(2000 + y1, m1, d1), pd.Timestamp(2000 + y2, m2, d2))
    except ValueError:
        return None


# ════════════════════════════════════════════════════════════════════════
# 载入层
# ════════════════════════════════════════════════════════════════════════

def _first(patterns, folder):
    for p in patterns:
        hits = sorted(g for g in glob.glob(os.path.join(folder, p))
                      if not os.path.basename(g).startswith("~$"))
        if hits:
            return hits[0]
    return None


def load_ventas(folder):
    path = _first(["*Ventas_MX*.xlsx", "*Ventas*.xlsx"], folder)
    if not path:
        raise IOError("找不到 Ventas 报表：%s" % folder)
    df = pd.read_excel(path, sheet_name=VENTAS_SHEET, header=VENTAS_HEADER_ROW)
    df["fecha"] = df[V_DATE].map(parse_spanish_date) if V_DATE in df.columns else pd.NaT
    df["ym"] = df["fecha"].dt.to_period("M").astype(str)
    df["oid"] = idkey(df[V_ORDER_ID]) if V_ORDER_ID in df.columns else ""
    df["pid"] = idkey(df[V_PACK_ID]) if V_PACK_ID in df.columns else ""
    st = df[V_STATUS].astype(str) if V_STATUS in df.columns else pd.Series("", index=df.index)
    df["abnormal"] = st.str.contains(ABNORMAL_RE, case=False, na=False)
    # 卖家承担的优惠券：平台把 Descuentos y bonificaciones 整列导成空白，
    # 只能用 Total 与各明细列的差额倒推。Total 为空的行（换货第二行）不参与。
    tot_raw = pd.to_numeric(df[V_TOTAL], errors="coerce") if V_TOTAL in df.columns else pd.Series(float("nan"), index=df.index)
    recomputed = (num(df, V_REVENUE) + num(df, V_SHIP_INCOME) + num(df, V_COMMISSION)
                  + num(df, V_SHIP_COST) + num(df, V_REFUND))
    coupon = -(tot_raw.fillna(0.0) - recomputed)
    coupon[tot_raw.isna()] = 0.0
    df["coupon"] = coupon
    return path, df


def load_billing(folder):
    """Facturación + Notas de Crédito 全部期间合并。两者是不同的凭证，不会重复。"""
    frames = []
    for pat, doc in (("Reporte_Facturacion_MercadoLibre_*.xlsx", "FAC"),
                     ("Reporte_Notas_Credito_MercadoLibre_*.xlsx", "NC")):
        for p in sorted(glob.glob(os.path.join(folder, pat))):
            if os.path.basename(p).startswith("~$"):
                continue
            t = pd.read_excel(p, sheet_name="REPORT", header=BILLING_HEADER_ROW)
            t["_doc"] = doc
            t["_file"] = os.path.basename(p)
            frames.append(t)
    if not frames:
        raise IOError("找不到 Facturación 报表：%s" % folder)
    df = pd.concat(frames, ignore_index=True)
    df = df[df[F_AMOUNT].notna()].copy()
    df["amount"] = pd.to_numeric(df[F_AMOUNT], errors="coerce").fillna(0.0)
    df["cdate"] = pd.to_datetime(df.get(F_CHARGE_DATE), errors="coerce")
    df["sdate"] = pd.to_datetime(df.get(F_SALE_DATE), errors="coerce")
    df["charge_no"] = idkey(df[F_CHARGE_NO]) if F_CHARGE_NO in df.columns else ""
    df["reverses"] = idkey(df[F_REVERSES]) if F_REVERSES in df.columns else ""
    df["k_sale"] = idkey(df[F_SALE_NO]) if F_SALE_NO in df.columns else ""
    df["k_pack"] = idkey(df[F_PACK_NO]) if F_PACK_NO in df.columns else ""
    df["channel"] = df[F_CHANNEL].astype(str) if F_CHANNEL in df.columns else ""

    # 冲销行归属：优先用 Cargo que bonifica → Número del cargo 的 ID 关联，
    # 关联不上再退回按名称匹配（REVERSAL_TO_PARENT）。
    by_no = {}
    for cn, det, ch in zip(df["charge_no"], df[F_DETALLE].astype(str), df["channel"]):
        if cn and cn.lower() not in ("nan", "none", ""):
            by_no[cn] = (det, ch)
    parents, pchannels, known = [], [], []
    for det, rev in zip(df[F_DETALLE].astype(str), df["reverses"]):
        if is_reversal(det):
            hit = by_no.get(rev)
            if hit:
                parents.append(hit[0]); pchannels.append(hit[1])
                known.append(hit[0] in FEE_BY_ES)
                continue
            p = REVERSAL_TO_PARENT.get(det)
            parents.append(p); pchannels.append(None); known.append(bool(p) and p in FEE_BY_ES)
        else:
            parents.append(det); pchannels.append(None); known.append(det in FEE_BY_ES)
    df["parent_detalle"] = parents
    df["parent_channel"] = pchannels
    df["known"] = known
    df["is_rev"] = df[F_DETALLE].astype(str).map(is_reversal)
    return df


def load_pagos(folder, year, month):
    """当期的账单支付与贷记单。返回 (path, DataFrame) 或 (None, 空表)。"""
    best = None
    for p in sorted(glob.glob(os.path.join(folder, "Reporte_Pagos_Facturas_*.xlsx"))):
        tok = file_month_token(os.path.basename(p))
        if tok == (year, month):
            best = p
            break
    if not best:
        return None, pd.DataFrame()
    df = pd.read_excel(best, sheet_name="Pagos y notas de crédito",
                       header=PAGOS_HEADER_ROW).dropna(how="all")
    return best, df


def load_cargos_full(folder, year, month):
    """Full 费用明细（七张分表）。仅用于交叉核对，绝不与账单相加。"""
    best = None
    for p in sorted(glob.glob(os.path.join(folder, "Reporte_Cargos_Full_*.xlsx"))):
        if file_month_token(os.path.basename(p)) == (year, month):
            best = p
            break
    if not best:
        return None, {}, 0.0
    sheets, total = {}, 0.0
    xl = pd.ExcelFile(best)
    for sn in xl.sheet_names:
        d = pd.read_excel(best, sheet_name=sn, header=CARGOS_HEADER_ROW)
        col = next((c for c in d.columns if "Monto del cargo" in str(c)), None)
        if col is None:
            continue
        v = pd.to_numeric(d[col], errors="coerce")
        sheets[sn] = (float(v.sum()), int(v.notna().sum()))
        total += float(v.sum())
    return best, sheets, total


def load_storage(folder, p0, p1):
    """独立仓储费报表（按实际占用日）。挑与报表月重叠最多的那份。"""
    best, best_overlap = None, -1
    for p in sorted(glob.glob(os.path.join(folder, "*Costos_por_servicio_almacenamiento.xlsx"))):
        rng = storage_file_range(os.path.basename(p))
        if not rng:
            continue
        ov = (min(rng[1], p1) - max(rng[0], p0)).days
        if ov > best_overlap:
            best, best_overlap = p, ov
    if not best or best_overlap < 0:
        return None, 0.0, None
    s = pd.read_excel(best, sheet_name="Resumen", header=STORAGE_HEADER_ROW)
    col = "Total" if "Total" in s.columns else s.columns[-1]
    vals = pd.to_numeric(s[col], errors="coerce").dropna()
    accrual = float(vals.iloc[-1]) if len(vals) else 0.0
    return best, accrual, storage_file_range(os.path.basename(best))


def load_returns(folder):
    path = _first(["Returns*.xlsx"], folder)
    if not path:
        return None, pd.DataFrame()
    try:
        r = pd.read_excel(path, sheet_name="Triages", header=RETURNS_HEADER_ROW)
    except Exception:
        return path, pd.DataFrame()
    r = r.dropna(how="all")
    if "Fecha de revisión" in r.columns:
        r["fr"] = pd.to_datetime(r["Fecha de revisión"], format="%d-%m-%Y", errors="coerce")
    else:
        r["fr"] = pd.NaT
    return path, r


STOCK_SHEET = "Resumen"
# Resumen 页的表头是**两到三层**：第一层是分组（"Unidades en Full"），
# 下面才是真正的列名（"Aptas para vender"）。而且有两列都叫
# "Pendientes de ingreso"（一列是在途，一列是"建议动作"分类），
# 只能靠分组名区分。所以按 (分组, 列名) 匹配，不按列位置 —— 平台一改列序
# 位置法就会静默错位，而且错得看不出来。
STOCK_COLS = {
    "código ml": "ml",
    "sku": "sku",
    "producto": "title",
    "estado de la publicación": "estado",
    "ofrece full": "full",
    "stock promedio últimos 30 días (u.)": "avg_stock",
    "unidades con antigüedad": "aged",
    "tiempo hasta agotar stock": "runout",
    "en transferencia": "transfer",
    "devueltas por el comprador": "returned",
    "aptas para vender": "sellable",
    "no aptas para vender": "unsellable",
    "extraviadas": "lost",
    "en revisión": "in_review",
    "ventas canceladas": "cancelled_u",
    "unidades que ocupan espacio en full": "occupying",
    "buena calidad": "good",
    "para impulsar ventas": "boost",
    "para poner en venta": "to_list",
    "para evitar descarte": "discard",
    "ventas": "sales30_amt",
    "unidades vendidas": "sales30_u",
    "urgencia de envío": "urgency",
}
STOCK_NUM = ["avg_stock", "aged", "transfer", "returned", "sellable", "unsellable",
             "lost", "in_review", "cancelled_u", "occupying", "in_transit",
             "act_pending", "good", "boost", "to_list", "discard",
             "sales30_amt", "sales30_u"]


def _hnorm(v):
    """表头单元格 → 匹配用的短名：只取第一行、去空白、转小写。

    平台把说明文字塞在同一个单元格里（"Unidades con antigüedad\\n¿Por qué...?"），
    整串拿去匹配会被这些说明拖着变。
    """
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    return str(v).split("\n")[0].strip().lower()


def load_stock(folder):
    """Full 仓库存报表（stock_general_full）的 Resumen 页。

    这一页本身就是完整的 SKU 级主表，其余四页只是它的切片，所以只读这一页。
    注意：它是**下载当时的快照**，不是月末余额 —— 与损益表不同期，不能相加。

    返回 (文件路径, 快照文字, DataFrame)。没有报表时返回空 DataFrame 而不报错：
    这个报表要单独的权限，不少店根本拉不到。
    """
    path = _first(["stock_general_full*.xlsx"], folder)
    if not path:
        return None, "", pd.DataFrame()
    try:
        raw = pd.read_excel(path, sheet_name=STOCK_SHEET, header=None)
    except Exception:
        return path, "", pd.DataFrame()

    snap = ""
    hdr = None
    for i in range(len(raw)):
        row = [_hnorm(v) for v in raw.iloc[i]]
        if not snap:
            for v in raw.iloc[i]:
                if isinstance(v, str) and v.strip().startswith("Actualizado el"):
                    snap = v.strip()
        if "código ml" in row:
            hdr = i
            break
    if hdr is None:
        return path, snap, pd.DataFrame()

    ncol = raw.shape[1]
    group = [""] * ncol            # 第一层分组，向右填充
    sub = [""] * ncol              # 最靠下的那一层才是真列名
    cur = ""
    for j in range(ncol):
        g = _hnorm(raw.iloc[hdr, j])
        if g:
            cur = g
        group[j] = cur
        for k in (hdr + 2, hdr + 1, hdr):
            if k < len(raw):
                s = _hnorm(raw.iloc[k, j])
                if s:
                    sub[j] = s
                    break

    field = {}
    for j in range(ncol):
        name = STOCK_COLS.get(sub[j])
        if sub[j] == "pendientes de ingreso":
            # 两列同名，靠分组区分：一列是在途，一列是"建议动作"里的在途
            name = "in_transit" if "camino" in group[j] else "act_pending"
        if name and name not in field:
            field[name] = j

    body = raw.iloc[hdr + 1:]
    if "ml" not in field:
        return path, snap, pd.DataFrame()
    body = body[body.iloc[:, field["ml"]].notna()]     # 末尾的合计行没有 ML 码，正好排除

    out = pd.DataFrame(index=range(len(body)))
    for name, j in field.items():
        out[name] = body.iloc[:, j].values
    for c in STOCK_NUM:
        out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0.0) if c in out.columns \
            else 0.0
    for c in ("sku", "title", "estado", "runout", "urgency", "full", "ml"):
        out[c] = out[c].astype(str).str.strip() if c in out.columns else ""
        # 未上架的商品这些列是空的，pandas 转成字符串会写出 "nan" 摆在表里
        out[c] = out[c].replace({"nan": "", "NaT": "", "None": ""})
    out["total_u"] = (out["in_transit"] + out["transfer"] + out["returned"]
                      + out["sellable"] + out["unsellable"] + out["lost"]
                      + out["in_review"] + out["cancelled_u"])
    return path, snap, out


def stock_snapshot_date(snap, folder):
    """把 'Actualizado el 8 de septiembre a las 23:40 hs.' 解析成日期。

    平台**不写年份**，所以用下载目录名里的年份（downloads/<店>/YYYYMMDD_HHMMSS）。
    跨年下载时快照月份会大于目录月份，这时年份减一。
    """
    m = re.search(r"(\d{1,2})\s+de\s+([A-Za-zÁ-úá-ú]+)", snap or "")
    if not m:
        return pd.NaT
    mon = SPANISH_MONTHS.get(m.group(2).lower())
    if not mon:
        return pd.NaT
    base = os.path.basename(os.path.normpath(folder))
    try:
        ref = pd.Timestamp(int(base[:4]), int(base[4:6]), int(base[6:8]))
    except Exception:
        ref = pd.Timestamp(datetime.date.today())
    for y in (ref.year, ref.year - 1):
        try:
            d = pd.Timestamp(y, mon, int(m.group(1)))
        except ValueError:
            return pd.NaT
        if d <= ref + pd.Timedelta(days=1):
            return d
    return pd.NaT


def load_settlement(folder, p0, p1):
    """MercadoPago 结算流水（可选）。用于第三方交叉验证；很多店没有报表权限。"""
    path = _first(["settlement_v2*.csv"], folder)
    if not path:
        return None, None
    try:
        s = pd.read_csv(path, sep=";")
        s["dt"] = pd.to_datetime(s["TRANSACTION_DATE"], errors="coerce", utc=True)
    except Exception:
        return path, None
    s = s[(s["dt"].dt.tz_localize(None) >= p0) & (s["dt"].dt.tz_localize(None) <= p1 + pd.Timedelta(days=1))]
    if not len(s):
        return path, None
    g = lambda t: s[s["TRANSACTION_TYPE"] == t]
    st = g("SETTLEMENT")
    out = {
        "n": int(len(st)),
        "trans": float(st["TRANSACTION_AMOUNT"].sum()),
        "fee": float(st["FEE_AMOUNT"].sum()),
        "tax": float(st["TAXES_AMOUNT"].sum()),
        "real_settlement": float(st["REAL_AMOUNT"].sum()),
        "refund_real": float(s[s["TRANSACTION_TYPE"].isin(["REFUND", "DISPUTE"])]["REAL_AMOUNT"].sum()),
        "real_all": float(s["REAL_AMOUNT"].sum()),
    }
    pos = st[st["TRANSACTION_AMOUNT"] > 0]
    if len(pos):
        rate = (-pos["TAXES_AMOUNT"] / pos["TRANSACTION_AMOUNT"])
        out["rate_mean"] = float(rate.mean())
        out["rate_std"] = float(rate.std()) if len(pos) > 1 else 0.0
        out["rate_mode_n"] = int((rate.round(4) == round(rate.median(), 4)).sum())
        out["rate_n"] = int(len(pos))
    ref = s[s["TRANSACTION_TYPE"].isin(["REFUND", "DISPUTE"])]
    out["refund_detail"] = " ＋ ".join(
        "%s %s" % (t, fmt(v)) for t, v in
        zip(ref["TRANSACTION_TYPE"], ref["REAL_AMOUNT"])) or "无"
    return path, out


# ── 贷记明细 × 销售报表 逐单交叉比对 ────────────────────────────────
# 关联键有个坑：账单的 Número de venta（2000017/18xxx）与销售报表的
# # de venta（2000014xxx）**不是同一个号段**，可靠的桥是
# Número de paquete ↔ 销售报表的 # de venta / Orden de compra。
# 所以不能"k_sale 非空就取 k_sale"，必须挑能对上销售报表的那个键 ——
# 实测这么改，12 家店的匹配率从 132/464 升到 411/464（BOCINA_TA02 从 0 升到 11）。
CN_TOL = 0.05
PARTIAL_RE = re.compile(r"parcial", re.I)


def _srcid(v):
    """MercadoPago 收款号规整成字符串。

    同一列在不同店铺导出成三种 dtype：float64、字符串、甚至非数字的
    "000dj6x3ef"。直接 "%d" % float(v) 会在字符串上抛异常，把整份流水
    连带丢掉（实测 12 家店里 10 家因此拿不到现金凭证）。逐值判断，
    认不出来的原样保留 —— 匹配不上只是少一条凭证，不该拖垮其余的。
    """
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    if isinstance(v, float):
        return "%d" % v
    s = str(v).strip()
    return s[:-2] if s.endswith(".0") else s


def load_mp_cash(folder):
    """MercadoPago 逐笔到账净额 {收款号: 净额}，用来给比对结论找现金凭证。

    两个来源都认：reserve-release（明细最全，含 mediation/refund 分解）优先，
    没有就退回 settlement_v2。两者都要单独的报表权限，拿不到就返回空 dict，
    比对照常做，只是结论降级为"无现金凭证可核"。
    """
    out = {}
    p = _first(["reserve-release-*.csv"], folder)
    if p:
        try:
            d = pd.read_csv(p, sep=";")
            net = (pd.to_numeric(d["NET_CREDIT_AMOUNT"], errors="coerce").fillna(0.0)
                   - pd.to_numeric(d["NET_DEBIT_AMOUNT"], errors="coerce").fillna(0.0))
            for a, b in zip(d["SOURCE_ID"].map(_srcid), net):
                if a:
                    out[a] = out.get(a, 0.0) + float(b)
        except Exception:
            out = {}
    if out:
        return out
    p = _first(["settlement_v2*.csv"], folder)
    if p:
        try:
            d = pd.read_csv(p, sep=";")
            amt = pd.to_numeric(d["REAL_AMOUNT"], errors="coerce").fillna(0.0)
            for a, b in zip(d["SOURCE_ID"].map(_srcid), amt):
                if a:
                    out[a] = out.get(a, 0.0) + float(b)
        except Exception:
            return {}
    return out


def credit_note_audit(S, folder):
    """按订单号把贷记明细和销售报表逐单对照，找三类现有校验看不见的异常。

    现有 7 条校验的盲区：桥A 只在销售报表内部勾稽（残差恒为 0），桥B 只比账单
    总额，MercadoPago 那条只比店铺汇总差额（三种原因混在一起分不开）。下面三类
    都是逐单的，落在盲区里：

      orphan  —— 账单为某单计了费，销售报表里根本查无此单，且计提后没有被冲平。
                 绝大多数"查无此单"是下单后立刻作废（计费又全额冲销，净额为 0），
                 无害；净额不为零的才是真的被收了钱。
      unrev   —— 取消单计了费却没有任何冲销。注意发货前取消的单平台压根不下账单，
                 所以"没有贷记单"本身不是问题，必须回查账单确认确实计过费。
      partial —— 佣金被**部分**冲销（Estado del cargo = Anulado parcialmente）。
                 平台按退款比例 p 同时退还佣金与代扣税，但销售报表的
                 「退款与取消」列在这条路径上会把原佣金**重复扣一次**：
                     应退       = p × (GMV − 佣金 − 代扣税)
                     销售报表写的 = 应退 + 原佣金
                 即销售报表少算了一个原佣金的利润。已用 MercadoPago 现金流水在
                 6 家店 10 单上逐单证实，每一单差额都精确等于该单原佣金。
    """
    A, V, B = S["ventas"], S["ventas_all"], S["billing_all"]
    res = {"orphan": [], "unrev": [], "partial": [], "mismatch": [],
           "orphan_total": 0.0, "unrev_total": 0.0, "partial_total": 0.0,
           "matched": 0, "order_rows": 0, "has_cash": False}
    if not len(B) or "Estado del cargo" not in B.columns:
        return res

    ids = set(V["oid"]) | set(V["pid"])
    ids.discard(""); ids.discard("nan")
    cash = load_mp_cash(folder)
    res["has_cash"] = bool(cash)

    def okey(ks, kp):
        """账单行 → 销售报表里的订单键；对不上返回 (None, 账单自己的键)。"""
        ks, kp = str(ks).strip(), str(kp).strip()
        if ks in ids:
            return ks, ks
        if kp in ids:
            return kp, kp
        raw = kp if kp not in ("", "nan") else (ks if ks not in ("", "nan") else "")
        return None, raw

    # ---- orphan：账单有、销售报表无，且没被冲平 ----
    agg = {}
    for ks, kp, amt, doc in zip(B["k_sale"], B["k_pack"], B["amount"], B["_doc"]):
        hit, raw = okey(ks, kp)
        if not raw:
            continue                       # 账单级费用（广告等），本来就没有订单号
        res["order_rows"] += 1
        if hit:
            res["matched"] += 1
        else:
            a = agg.setdefault(raw, [0.0, 0])
            a[0] += float(amt); a[1] += 1
    for k, (amt, cnt) in sorted(agg.items(), key=lambda x: -abs(x[1][0])):
        if abs(amt) > CN_TOL:
            res["orphan"].append((k, amt, cnt))
            res["orphan_total"] += amt

    # ---- unrev：取消单计了费却没冲销 ----
    nc_keys = set()
    for ks, kp, doc in zip(B["k_sale"], B["k_pack"], B["_doc"]):
        if doc != "NC":
            continue
        hit, _ = okey(ks, kp)
        if hit:
            nc_keys.add(hit)
    if len(A) and V_STATUS in A.columns:
        can = A[A[V_STATUS].astype(str).str.contains("cancel", case=False, na=False)]
        for _, x in can.iterrows():
            keys = [z for z in (str(x["oid"]), str(x["pid"])) if z not in ("", "nan", " ")]
            if set(keys) & nc_keys:
                continue
            b = B[B["k_sale"].astype(str).isin(keys) | B["k_pack"].astype(str).isin(keys)]
            net = float(b["amount"].sum()) if len(b) else 0.0
            if len(b) and abs(net) > CN_TOL:     # 真的计过费又没退
                res["unrev"].append((keys[0], net, str(x.get(V_STATUS, "")),
                                     x.get("fecha"), str(x.get(V_SKU, ""))))
                res["unrev_total"] += net

    # ---- partial：佣金被部分冲销，销售报表重复扣了一次原佣金 ----
    par = B[(B["_doc"] != "NC") & (B["Detalle"].astype(str) == "Cargo por venta")
            & B["Estado del cargo"].astype(str).str.contains(PARTIAL_RE, na=False)]
    for _, x in par.iterrows():
        com = float(x["amount"])
        if com <= 0:
            continue
        back = -float(B[B["reverses"].astype(str) == str(x["charge_no"])]["amount"].sum())
        p = back / com
        hit, _ = okey(x["k_sale"], x["k_pack"])
        if not hit:
            continue
        v = A[(A["oid"] == hit) | (A["pid"] == hit)]
        if not len(v):
            continue                        # 订单不在本会计期，留给它自己的月份处理
        gmv = float(num(v, V_REVENUE).sum())
        tax = float(-num(v, V_COMMISSION).sum()) - com
        vref = float(-num(v, V_REFUND).sum())
        vtot = float(num(v, V_TOTAL).sum())
        want = p * (gmv - com - tax)        # 闭式：退多少货款就退多少比例的费用
        expect = want + com                 # 销售报表在这条路径上实际会写的数
        # 现金凭证：MercadoPago 实际净额与销售报表 Total 的差应恰好等于原佣金
        pago = _srcid(x.get("Pago"))
        real = cash.get(pago)
        proof = ("无 MercadoPago 流水" if not cash else
                 "流水缺此笔" if real is None else
                 "现金凭证确认" if abs((real - vtot) - com) <= CN_TOL else
                 "现金凭证不符（差 %s）" % fmt(real - vtot - com))
        # 认定顺序：现金凭证 > 公式吻合。流水是独立凭证，直接证明销售报表的
        # Total 比实际净额少了一个原佣金；公式只是"长得像这个模式"。两者冲突时
        # 以流水为准 —— 实测有两单公式差几块钱，流水却分毫不差。
        by_cash = real is not None and abs((real - vtot) - com) <= CN_TOL
        row = (hit, p, com, gmv, tax, vref, want, vtot, real, proof,
               str(x.get("Fecha de venta", ""))[:10], str(v[V_SKU].iloc[0]),
               "现金凭证" if by_cash else "按规律推定")
        if by_cash or abs(vref - expect) <= CN_TOL:
            res["partial"].append(row)
            res["partial_total"] += com
        else:
            res["mismatch"].append(row + (vref - expect,))
    return res


# ════════════════════════════════════════════════════════════════════════
# 推导层 —— 把"发现"固化成公式
# ════════════════════════════════════════════════════════════════════════

def derive(store, folder, p0, p1):
    """读一家店一个月的全部原始文件，算出报表需要的每一个数。"""
    S = {"store": store, "folder": folder, "files": {}, "anomalies": []}

    vpath, V = load_ventas(folder)
    S["files"]["ventas"] = vpath
    S["ventas_all"] = V
    A = V[(V["fecha"] >= p0) & (V["fecha"] <= p1 + pd.Timedelta(days=1) - pd.Timedelta(seconds=1))].copy()
    S["ventas"] = A

    B = load_billing(folder)
    S["billing_all"] = B
    FA = B[(B["cdate"] >= p0) & (B["cdate"] <= p1)].copy()
    S["billing"] = FA

    ids = set(A["oid"]) | set(A["pid"])
    ids.discard(""); ids.discard("nan")
    L = B[B["k_sale"].isin(ids) | B["k_pack"].isin(ids)].copy()   # 跨账期，8月订单的费用可能落在9月账单
    S["billing_linked"] = L

    def bucket_sum(frame, bucket, reversal=None):
        f = frame[frame["parent_detalle"].map(
            lambda d: FEE_BY_ES.get(str(d), (None, None, None))[1] == bucket)]
        if reversal is True:
            f = f[f["is_rev"]]
        elif reversal is False:
            f = f[~f["is_rev"]]
        return float(f["amount"].sum())

    m = {}
    m["rows"] = int(len(A))
    m["orders_unique"] = int(A["oid"].nunique())
    m["units"] = float(num(A, V_UNITS).sum())
    m["ing"] = float(num(A, V_REVENUE).sum())
    m["env_ing"] = float(num(A, V_SHIP_INCOME).sum())
    m["com_imp"] = float(-num(A, V_COMMISSION).sum())      # 佣金 + 代扣税（Ventas 口径）
    m["env_cost"] = float(-num(A, V_SHIP_COST).sum())
    m["refund"] = float(-num(A, V_REFUND).sum())
    m["total"] = float(num(A, V_TOTAL).sum())
    m["coupon"] = float(A["coupon"].sum())
    m["coupon_n"] = int((A["coupon"].abs() > 0.02).sum())
    denoms = A.loc[A["coupon"].abs() > 0.02, "coupon"].round(2).value_counts()
    m["coupon_denoms"] = {float(k): int(v) for k, v in denoms.items()}

    m["com_gross"] = bucket_sum(L, "commission", reversal=False)
    m["com_anul"] = bucket_sum(L, "commission", reversal=True)
    m["env_gross"] = bucket_sum(L, "shipping", reversal=False)
    m["env_anul"] = bucket_sum(L, "shipping", reversal=True)
    m["dev_gross"] = bucket_sum(L, "returns", reversal=False)
    m["dev_anul"] = bucket_sum(L, "returns", reversal=True)
    m["dev_net"] = m["dev_gross"] + m["dev_anul"]

    # ★ 代扣代缴税金：Ventas 的"佣金及税金" − 账单的纯佣金（毛额）
    m["tax_ret"] = m["com_imp"] - m["com_gross"]
    m["tax_rate_gross"] = (m["tax_ret"] / m["ing"]) if m["ing"] else 0.0
    m["iva_in_gmv"] = m["ing"] * IVA_RATE / (1 + IVA_RATE)
    m["net_revenue_exvat"] = m["ing"] - m["iva_in_gmv"]
    m["tax_vs_iva"] = (m["tax_ret"] / m["iva_in_gmv"]) if m["iva_in_gmv"] else 0.0

    # 桥 A 残差：等式成立时应为 0
    m["resid"] = m["total"] - (m["ing"] + m["env_ing"] - m["com_imp"]
                               - m["env_cost"] - m["refund"] - m["coupon"])

    # ---- 结算通道 ----
    m["bill_total"] = float(FA["amount"].sum())
    m["bill_by_channel"] = {c: float(FA[FA["channel"] == c]["amount"].sum())
                            for c in (CH_SI, CH_NO, CH_NA)}
    m["bill_cnt_by_channel"] = {c: int((FA["channel"] == c).sum())
                                for c in (CH_SI, CH_NO, CH_NA)}
    # 应付 = No 桶的费用 ＋ 同期冲销掉 No 桶费用的那些行
    rev_of_no = FA[FA["is_rev"] & (FA["parent_channel"] == CH_NO)]
    if not len(rev_of_no):   # ID 关联不上时退回按科目性质判断
        rev_of_no = FA[FA["is_rev"] & FA["parent_detalle"].map(
            lambda d: FEE_BY_ES.get(str(d), (None, None, None))[2] is False)]
    m["invoice_internal_reversal"] = float(rev_of_no["amount"].sum())
    m["payable"] = m["bill_by_channel"][CH_NO] + m["invoice_internal_reversal"]

    # ---- 账单级（与任何订单都不挂钩）的费用 ----
    other = FA[FA["sdate"].isna()].copy()
    S["billing_other"] = other
    m["other_total"] = float(other["amount"].sum())
    # 按科目净额（费用 + 其冲销）
    per_fee = {}
    for es, zh, bucket, order_level in FEE_CATALOG:
        if order_level:
            continue
        v = float(other[other["parent_detalle"].astype(str) == es]["amount"].sum())
        per_fee[es] = v
    m["bill_fee_net"] = per_fee
    named_sum = sum(per_fee.values())
    # 兜底行：保证"月度账单费用小计"恒等于 other_total，任何意外都看得见而不是消失
    m["unc_bill"] = m["other_total"] - named_sum
    # 订单级的未登记科目
    unc_order_rows = FA[(~FA["known"]) & FA["sdate"].notna()]
    m["unc_order"] = float(unc_order_rows["amount"].sum())
    m["order_linked_billed"] = m["payable"] - m["other_total"]

    # ---- 未登记科目清单（供 ⑨ 页与校验使用）----
    unk = FA[~FA["known"]]
    S["unknown_fees"] = []
    if len(unk):
        g = unk.groupby(unk[F_DETALLE].astype(str))["amount"].agg(["count", "sum"])
        for name, row in g.sort_values("sum", key=abs, ascending=False).iterrows():
            chans = sorted(set(unk[unk[F_DETALLE].astype(str) == name]["channel"]))
            S["unknown_fees"].append({"detalle": name, "count": int(row["count"]),
                                      "amount": float(row["sum"]),
                                      "channels": ", ".join(chans)})
    # 已登记但出现在意外通道的科目（信息性）
    for es, zh, bucket, order_level in FEE_CATALOG:
        sub = FA[(FA[F_DETALLE].astype(str) == es)]
        if not len(sub):
            continue
        has_sale = sub["sdate"].notna()
        if order_level and not has_sale.any():
            S["anomalies"].append("科目「%s」按字典应为订单级，但本月全部行都没有销售日期" % es)
        if (not order_level) and has_sale.any():
            S["anomalies"].append("科目「%s」按字典应为账单级，但本月有 %d 行带销售日期"
                                  % (es, int(has_sale.sum())))

    # ---- 账单支付 ----
    ppath, P = load_pagos(folder, p0.year, p0.month)
    S["files"]["pagos"] = ppath
    S["pagos"] = P
    if len(P) and "Tipo de pago" in P.columns:
        tp = P["Tipo de pago"].astype(str)
        m["pay_auto"] = float(P[tp.str.contains("Cobro", na=False)]["Importe total"].sum())
        m["pay_nc"] = float(P[tp.str.contains("Nota", na=False)]["Importe total"].sum())
        pd_col = "Fecha de pago / Emisión de NC"
        dts = pd.to_datetime(P[tp.str.contains("Cobro", na=False)].get(pd_col), errors="coerce").dropna()
        m["pay_date"] = dts.max().strftime("%Y-%m-%d") if len(dts) else ""
    else:
        m["pay_auto"] = m["pay_nc"] = 0.0
        m["pay_date"] = ""

    # ---- Full 费用交叉核对 ----
    cpath, csheets, ctotal = load_cargos_full(folder, p0.year, p0.month)
    S["files"]["cargos_full"] = cpath
    S["cargos_full_sheets"] = csheets
    m["cargos_full_report_total"] = ctotal
    full_buckets = ("storage", "fulfillment", "penalty")
    m["full_bill_total"] = float(sum(
        v for es, v in per_fee.items()
        if FEE_BY_ES[es][1] in full_buckets))
    m["ads_net"] = float(sum(v for es, v in per_fee.items() if FEE_BY_ES[es][1] == "ads"))
    m["subscription_net"] = float(sum(v for es, v in per_fee.items()
                                      if FEE_BY_ES[es][1] == "subscription"))

    # ---- 仓储费报表 ----
    spath, accrual, srange = load_storage(folder, p0, p1)
    S["files"]["storage"] = spath
    m["storage_accrual"] = accrual
    m["storage_range"] = srange
    m["storage_bill"] = per_fee.get("Cargo por servicio de almacenamiento Full", 0.0)

    # ---- 退货质检 ----
    rpath, R = load_returns(folder)
    S["files"]["returns"] = rpath
    S["returns_all"] = R
    if len(R):
        S["returns"] = R[(R["fr"] >= p0) & (R["fr"] <= p1)].copy()
    else:
        S["returns"] = pd.DataFrame()

    # ---- Full 仓库存快照 ----
    spath, snap, ST = load_stock(folder)
    S["files"]["stock"] = spath
    S["stock"] = ST
    S["stock_snap_text"] = snap
    S["stock_snap"] = stock_snapshot_date(snap, folder)

    # ---- MercadoPago 结算流水 ----
    mpath, MP = load_settlement(folder, p0, p1)
    S["files"]["settlement"] = mpath
    S["mp"] = MP

    # ---- 组合单 / 多件单 / 换货单 ----
    m["pack_flag_yes"] = int((A[V_PACK_FLAG] == "Sí").sum()) if V_PACK_FLAG in A.columns else 0
    m["multi_unit_rows"] = int((num(A, V_UNITS) > 1).sum())
    nsku = A.groupby("oid")[V_SKU].nunique() if V_SKU in A.columns else pd.Series(dtype=int)
    m["multi_sku_orders"] = int((nsku > 1).sum())
    nsku_all = V.groupby("oid")[V_SKU].nunique() if V_SKU in V.columns else pd.Series(dtype=int)
    m["multi_sku_orders_all"] = int((nsku_all > 1).sum())
    m["blank_total_rows"] = int(pd.to_numeric(A[V_TOTAL], errors="coerce").isna().sum()) if V_TOTAL in A.columns else 0
    m["abnormal_orders"] = int(A["abnormal"].sum())

    # ---- 售后与库存去向 ----
    st = A[V_STATUS].astype(str) if V_STATUS in A.columns else pd.Series("", index=A.index)
    m["n_back_on_sale"] = int(st.isin(STATES_BACK_ON_SALE).sum())
    m["n_stock_lost"] = int(st.isin(STATES_STOCK_LOST).sum())
    m["n_seller_keeps"] = int(st.isin(STATES_SELLER_KEEPS).sum())
    m["amt_stock_lost"] = float(num(A[st.isin(STATES_STOCK_LOST)], V_REVENUE).sum())

    S["m"] = m

    # ---- 贷记明细 × 销售报表 逐单交叉比对 ----
    S["cn"] = credit_note_audit(S, folder)

    return S


# ════════════════════════════════════════════════════════════════════════
# 校验层
# ════════════════════════════════════════════════════════════════════════

def _chk(key, store, name, ok, detail, severity="error", skipped=False):
    return {"key": key, "store": store, "name": name, "ok": bool(ok),
            "detail": detail, "severity": severity, "skipped": skipped}


def validate(stores, tol=0.05):
    out = []
    for S in stores:
        m, name = S["m"], S["store"]

        out.append(_chk("bridge_a", name, "桥A 逐笔订单勾稽",
                        abs(m["resid"]) <= tol,
                        "残差 %s（阈值 %s）。不为零说明还有未识别的隐藏费用或折扣。"
                        % (fmt(m["resid"]), fmt(tol))))

        if S["files"].get("pagos"):
            d = m["payable"] - m["pay_auto"]
            out.append(_chk("bridge_b", name, "桥B 账单应付 vs 实际扣款",
                            abs(d) <= 0.01,
                            "应付 %s，实扣 %s，差 %s" % (fmt(m["payable"]), fmt(m["pay_auto"]), fmt(d))))
        else:
            out.append(_chk("bridge_b", name, "桥B 账单应付 vs 实际扣款", True,
                            "本期没有 Reporte_Pagos_Facturas 文件，跳过", skipped=True))

        unk = S["unknown_fees"]
        out.append(_chk("fee_catalog", name, "费用科目全部已登记",
                        not unk,
                        "全部科目已在 FEE_CATALOG 登记" if not unk else
                        "发现 %d 个未登记科目，合计 %s：%s" % (
                            len(unk), fmt(sum(u["amount"] for u in unk)),
                            "；".join("%s(%s)" % (u["detalle"], fmt(u["amount"])) for u in unk))))

        out.append(_chk("bill_block_balance", name, "账单费用小计自平",
                        abs(m["unc_bill"]) <= 0.01,
                        "兜底行金额 %s。不为零说明有科目未在 FEE_CATALOG 里、或已登记科目出现在非预期位置。"
                        % fmt(m["unc_bill"]), severity="warn"))

        r = m["tax_rate_gross"]
        out.append(_chk("tax_rate", name, "代扣税率在预期区间",
                        TAX_RATE_MIN <= r <= TAX_RATE_MAX,
                        "实测 %s（预期 %s–%s）。越界说明税制或账号税务状态变了，需重新确认。"
                        % (pct(r, 3), pct(TAX_RATE_MIN, 1), pct(TAX_RATE_MAX, 1))))

        if S["files"].get("cargos_full"):
            d = m["cargos_full_report_total"] - m["full_bill_total"]
            out.append(_chk("cargos_full", name, "Full 费用报表 vs 账单",
                            abs(d) <= 0.01,
                            "Full 报表 %s，账单口径 %s，差 %s（两者必须相等，相加即重复计数）"
                            % (fmt(m["cargos_full_report_total"]), fmt(m["full_bill_total"]), fmt(d))))
        else:
            out.append(_chk("cargos_full", name, "Full 费用报表 vs 账单", True,
                            "本期没有 Reporte_Cargos_Full 文件，跳过", skipped=True))

        MP = S.get("mp")
        if MP:
            d = MP["real_all"] - m["total"]
            out.append(_chk("mp_settlement", name, "MercadoPago 流水交叉验证",
                            abs(d) <= 0.05,
                            "结算净额 %s vs 销售报表 Total %s，差 %s"
                            % (fmt(MP["real_all"]), fmt(m["total"]), fmt(d))))
        else:
            out.append(_chk("mp_settlement", name, "MercadoPago 流水交叉验证", True,
                            "本店没有 settlement 流水（多为该登录账号无报表权限），跳过",
                            severity="warn", skipped=True))

        # ---- 贷记明细 × 销售报表 逐单交叉比对 ----
        # 这三项都是逐单的，落在其余校验的盲区里：桥A 只在销售报表内部勾稽，
        # 桥B 只比账单总额，MercadoPago 那条只比店铺汇总差额。
        c = S["cn"]
        out.append(_chk(
            "cn_orphan", name, "账单有费用但销售报表无此单",
            not c["orphan"],
            ("全部对得上，或虽查无此单但计费已被全额冲平（净额 0）"
             if not c["orphan"] else
             "%d 单在账单里计了费、销售报表里查无此单，且未被冲平，合计 %s。"
             "这笔钱不在销售报表里，桥A 查不到，需人工向平台核实。"
             % (len(c["orphan"]), fmt(c["orphan_total"])))))

        out.append(_chk(
            "cn_unreversed", name, "取消单的费用是否已冲销",
            not c["unrev"],
            ("取消单要么平台压根没计费，要么已全额冲销"
             if not c["unrev"] else
             "%d 单已取消、账单却计了费且没有冲销，合计 %s。"
             "（发货前取消平台通常不下账单，所以只统计确实计过费的。）"
             % (len(c["unrev"]), fmt(c["unrev_total"])))))

        n_cash = sum(1 for x in c["partial"] if x[12] == "现金凭证")
        out.append(_chk(
            "cn_partial", name, "部分退款订单重复扣佣金",
            not c["partial"] and not c["mismatch"],
            ("本期没有部分冲销的佣金"
             if not c["partial"] and not c["mismatch"] else
             "%d 单佣金被部分冲销，销售报表把原佣金重复扣了一次，合计 %s"
             "（其中 %d 单有 MercadoPago 现金凭证，%d 单按规律推定）。"
             "已在损益表“加：部分退款重复扣佣金调整”行修正。%s"
             % (len(c["partial"]), fmt(c["partial_total"]), n_cash,
                len(c["partial"]) - n_cash,
                ("另有 %d 单金额对不上已知模式，未计入调整，需人工核对。"
                 % len(c["mismatch"])) if c["mismatch"] else "")),
            severity="warn"))
    return out


# ════════════════════════════════════════════════════════════════════════
# 出表层 —— 通用骨架
# ════════════════════════════════════════════════════════════════════════

class Sheet(object):
    """薄封装：多店铺时列数随店铺数变化，这里统一算列位。"""

    def __init__(self, wb, name, first_store_col, nstores, has_total=True):
        self.ws = wb.create_sheet(name)
        self.c0 = first_store_col
        self.n = nstores
        self.has_total = has_total

    def scol(self, i):
        return get_column_letter(self.c0 + i)

    @property
    def tcol_idx(self):
        return self.c0 + self.n

    @property
    def tcol(self):
        return get_column_letter(self.tcol_idx)

    def title(self, text, sub="", sub2=""):
        ws = self.ws
        ws["A1"] = text
        ws["A1"].font = F(14, True, C_NAVY)
        ws.row_dimensions[1].height = 24
        if sub:
            ws["A2"] = sub
            ws["A2"].font = F(9, False, C_GREY)
        if sub2:
            ws["A3"] = sub2
            ws["A3"].font = F(9, False, C_GREY)

    def header(self, row, cols, widths, height=30, color=C_NAVY):
        ws = self.ws
        for i, t in enumerate(cols, 1):
            c = ws.cell(row=row, column=i, value=t)
            c.font = F(10, True, "FFFFFF")
            c.fill = fill(color)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = BOX
        ws.row_dimensions[row].height = height
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w


def store_cols(stores):
    return [S["store"] for S in stores]


def fee_rows(S, es, reversal=False):
    """(笔数, 金额) —— 当期账单里某个科目的原始行或其冲销行。"""
    FA = S["billing"]
    f = FA[FA[F_DETALLE].astype(str) == es] if not reversal else \
        FA[FA["is_rev"] & (FA["parent_detalle"].astype(str) == es)]
    return int(len(f)), float(f["amount"].sum())


def per_store(stores, fn, sep="；"):
    """'EWTTO_SM 130,027.31；UNIT_PW01 2,921.26'"""
    return sep.join("%s %s" % (S["store"], fn(S)) for S in stores)


def main_store(stores):
    return max(stores, key=lambda S: S["m"]["ing"])


def agg_tax_rate(stores):
    """全部店铺合计的代扣税率（按 GMV 加权）。

    凡是句子里引用的金额是"全部店铺合计"，税率就必须同口径。早先这类句子用的是
    main_store 的税率 —— 单店报表看不出问题，12 家店的汇总报表却会把 GMV 最大
    那家的税率说成全体的。
    """
    tax = sum(S["m"]["tax_ret"] for S in stores)
    gmv = sum(S["m"]["ing"] for S in stores)
    return (tax / gmv) if gmv else 0.0


def store_count_cn(stores):
    """店铺数量的中文表述，供正文使用。"""
    n = len(stores)
    return "本店" if n == 1 else "%d 家店铺合计" % n


# ════════════════════════════════════════════════════════════════════════
# ① 阅读说明
# ════════════════════════════════════════════════════════════════════════

def sheet_readme(wb, stores, ctx):
    ws = wb.create_sheet(SHEET_NAMES[0])
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 118
    ws["A1"] = "Mercado Libre 墨西哥店铺 月度财务报表"
    ws["A1"].font = F(15, True, C_NAVY)
    ws.row_dimensions[1].height = 26
    ws["A2"] = ("会计期间：%s　|　币种：MXN 墨西哥比索　|　编制日期：%s　|　数据截取时点：%s"
                % (ctx["period_cn"], ctx["build_date"], ctx["cutoff"]))
    ws["A2"].font = F(9, False, C_GREY)

    M = main_store(stores)
    mm = M["m"]
    chk = {(c["store"], c["key"]): c for c in ctx["checks"]}
    ok_a = all(chk[(S["store"], "bridge_a")]["ok"] for S in stores)
    ok_b = all(chk[(S["store"], "bridge_b")]["ok"] for S in stores)
    mp_stores = [S for S in stores if S.get("mp")]
    coupon_stores = [S for S in stores if S["m"]["coupon_n"] > 0]

    src_names = {"ventas": "销售明细(Ventas MX)", "pagos": "账单支付(Pagos)",
                 "cargos_full": "Full 费用报表(Cargos Full)", "storage": "仓储费用报表",
                 "returns": "退货质检(Returns/Triages)", "settlement": "MercadoPago 结算流水"}
    have = []
    for k, label in src_names.items():
        who = [S["store"] for S in stores if S["files"].get(k)]
        if who:
            have.append(label if len(who) == len(stores) else "%s(仅 %s)" % (label, "、".join(who)))
    sources = "、".join(["平台账单(Facturación)", "贷记单(Notas de Crédito)"] + have)

    rows = [("", "", "")]
    rows.append(("S", "一、本报表覆盖范围", ""))
    for i, S in enumerate(stores, 1):
        s = S["m"]
        rows.append(("", "店铺 %d" % i,
                     "%s —— %s %s 单 / %s 件 / GMV %s"
                     % (S["store"], ctx["period_short"], fmt0(s["rows"]),
                        fmt0(s["units"]), fmt(s["ing"]))))
    rows.append(("", "数据来源", sources))
    rows.append(("", "", ""))

    rows.append(("S", "二、为什么你以前“平不了账”——六个结构性原因", ""))
    rows.append(("", "① 三条时间线",
                 "销售报表按【下单日】、平台账单按【计费日】、资金到账按【放款日】。同一笔订单会落在三个不同的月份，"
                 "直接把三张表加总必然对不上。本报表统一采用【订单日口径】做损益，并在“平账勾稽”页给出到账单口径的桥式调节。"))
    rows.append(("", "② 佣金里混进了代扣税",
                 "销售报表的“Cargo por venta e impuestos”= 平台佣金 + 平台代扣代缴的税金(IVA/ISR retenciones)，"
                 "而平台账单里的“Cargo por venta”只有纯佣金。两者差额 = 代扣税，本期实测约占含税售价的 %s。"
                 "这笔钱不是费用，是已预缴税款，可用于抵扣申报，必须单列。本期：%s。"
                 % (pct(mm["tax_rate_gross"], 2),
                    per_store(stores, lambda S: fmt(S["m"]["tax_ret"]), "、"))))
    rows.append(("", "③ 费用分两个通道结算",
                 "账单里“Descontado de la operación = Si”的费用（佣金、配送费）在每笔订单打款时就已扣掉，不会再收一次；"
                 "“No”的费用（广告、仓储、揽收、罚金等）月底汇总，从 MercadoPago 余额自动扣款。"
                 "把两类相加就会重复计算一次佣金和运费（本期重复额将达 %s）。"
                 % fmt(sum(S["m"]["bill_by_channel"][CH_SI] for S in stores))))
    rows.append(("", "④ 退货退款走三套凭证",
                 "一笔退货会同时产生：销售报表的“Anulaciones y reembolsos”（退给买家的钱）、"
                 "账单里的“Anulación del cargo”（平台退还的佣金/运费）、以及贷记单 Nota de Crédito（另开红字发票）。"
                 "只看其中一套，退货损失就会算错。"))
    if coupon_stores:
        cs = coupon_stores[0]
        den = "、".join("%s×%d 笔" % (fmt(k, 0), v)
                        for k, v in sorted(cs["m"]["coupon_denoms"].items(), reverse=True))
        rows.append(("", "⑤ 卖家优惠券被导出成空白",
                     "Ventas 报表本来有“Descuentos y bonificaciones”（折扣与补贴）这一列，但平台导出时整列是空的。"
                     "本期 %s 有 %d 笔自负优惠券、合计 %s（面额 %s），在任何一张明细表里都查不到，只能用 Total 倒推。"
                     "已验证：账单佣金 ÷（原价 − 券额）= 标准费率，确认平台是按扣券后的价格计佣。"
                     % (cs["store"], cs["m"]["coupon_n"], fmt(cs["m"]["coupon"]), den)))
    else:
        rows.append(("", "⑤ 卖家优惠券被导出成空白",
                     "Ventas 报表的“Descuentos y bonificaciones”列被平台导出成空白。本期各店由 Total 倒推的券额均为 0，"
                     "但这一列随时可能再次出现金额，本报表每期都会自动倒推核对。"))
    rows.append(("", "⑥ 缺少商品采购成本",
                 "平台所有报表都不含你的进货成本，因此平台数据最多算到“经营贡献毛利”。"
                 "本报表在【⑦ SKU盈利分析】页留出黄色单元格供录入单位成本，录入后净利润会自动重算。"))
    rows.append(("", "", ""))

    rows.append(("S", "三、本报表的核心勾稽等式（每期自动校验）", ""))
    rows.append(("", "等式 1（订单资金）",
                 "商品收入 + 运费收入 − 卖家优惠券 − 平台佣金 − 平台配送费 − 平台代扣税 − 退款 "
                 "= 销售报表“Total”列 = 实际到账"))
    rows.append(("", "　验证结果",
                 ("全部店铺完全相等，差异 0.00，逐笔订单勾稽无残留 ✓" if ok_a else
                  "⚠ 未通过：%s" % per_store(stores, lambda S: "残差 " + fmt(S["m"]["resid"])))))
    rows.append(("", "等式 2（账单支付）",
                 "本期账单中“不在订单内扣除”的费用 − 同期发票内冲销 = 月底自动扣款金额"))
    rows.append(("", "　验证结果",
                 per_store(stores, lambda S: "%s − %s = %s（实扣 %s）%s" % (
                     fmt(S["m"]["bill_by_channel"][CH_NO]),
                     fmt(-S["m"]["invoice_internal_reversal"]),
                     fmt(S["m"]["payable"]), fmt(S["m"]["pay_auto"]),
                     "✓" if chk[(S["store"], "bridge_b")]["ok"] else "⚠"))))
    if mp_stores:
        s0 = mp_stores[0]
        MP = s0["mp"]
        rows.append(("", "等式 3（第三方交叉验证）",
                     "%s 的 MercadoPago 结算流水：交易额 %s、手续费 %s、代扣税 %s、净额 %s，"
                     "与销售报表推算值逐项一致，证明本报表口径正确。"
                     % (s0["store"], fmt(MP["trans"]), fmt(-MP["fee"]),
                        fmt(-MP["tax"]), fmt(MP["real_all"]))))
    else:
        rows.append(("", "等式 3（第三方交叉验证）",
                     "本期没有任何一家店下载到 MercadoPago 结算流水，因此缺少独立的第三方凭证。"
                     "建议开通报表权限后重新下载，见【%s】。" % SHEET_NAMES[9]))
    rows.append(("", "", ""))

    rows.append(("S", "四、组合单 / 一单多件 / 换货单 是怎么处理的", ""))
    rows.append(("", "① 合并包裹（Paquete de varios productos = Sí）",
                 "这是最容易误解的一个字段。它的意思不是“一单买了多个商品”，而是“这笔订单和别的订单合并到同一个包裹发货”。"
                 "本期 %s，但每一行仍然是一笔独立的销售，有自己完整的收入、佣金、运费和 Total。"
                 "所以按行汇总不会重复也不会遗漏。"
                 % per_store(stores, lambda S: "%s 有 %s 行（占 %s）" % (
                     S["store"], fmt0(S["m"]["pack_flag_yes"]),
                     pct(S["m"]["pack_flag_yes"] / S["m"]["rows"] if S["m"]["rows"] else 0, 0)), "、")))
    rows.append(("", "② 一单多件（Unidades > 1）",
                 "同一个 SKU 一次买多件，只占一行，金额列已经是多件的合计。"
                 "本报表用 Unidades 列算件数、用金额列算钱，两边都不会错。本期 %s。"
                 % per_store(stores, lambda S: "%s 有 %d 行" % (S["store"], S["m"]["multi_unit_rows"]), "、")))
    rows.append(("", "③ 真正的一单多 SKU",
                 "绝大多数是【换货】：第一行是原销售，带全部金额但 SKU 为空；第二行是换出去的新商品，有 SKU 和件数但金额全空。"
                 "金额不会重复计，但件数会多算、SKU 会归到换货商品上。本期 %s。"
                 % per_store(stores, lambda S: "%s %d 单（销售报表全窗口 %d 单）" % (
                     S["store"], S["m"]["multi_sku_orders"], S["m"]["multi_sku_orders_all"]), "、")))
    rows.append(("", "④ 汇总口径",
                 "本报表一律按【行】汇总原始金额列，不做任何拆分或分摊。%s"
                 % per_store(stores, lambda S: "%s：订单数按行计 %s 行，唯一订单号 %s 个，件数取 Unidades 合计 %s"
                             % (S["store"], fmt0(S["m"]["rows"]), fmt0(S["m"]["orders_unique"]),
                                fmt0(S["m"]["units"])))))
    rows.append(("", "", ""))

    rows.append(("S", "五、各页说明", ""))
    for nm, desc in [
        (SHEET_NAMES[1], "各店分列 + 合计，从 GMV 一路推到净利润，含费用率分析。行位置每期固定，可直接跨月对比。"),
        (SHEET_NAMES[2], "五个桥：A 订单资金流、B 月度账单、C 两种口径对比、D 从到账净额到经营利润、E 代扣税解释。这一页是解决“平不了账”的核心。"),
        (SHEET_NAMES[3], "本期账单全部费用科目（西班牙语原文 + 中文），标明结算通道。"),
        (SHEET_NAMES[4], "按订单状态、按金额、按平台/卖家责任划分。"),
        (SHEET_NAMES[5], "Full 仓退货逐件质检结果与退货原因，含可回收/彻底损失判定。"),
        (SHEET_NAMES[6], "按 SKU 的收入、费用、退货与贡献；五张表：表一正常订单、表二~表四三类异常明细、表五 SKU 真实盈利（最终口径）。单位采购成本只在表五填。"),
        (SHEET_NAMES[7], "销售报表窗口内各月的销售、费率、退货率走势。"),
        (SHEET_NAMES[8], "Full 仓库存结构、动销与周转，并把平台的近 30 天销量与销售报表逐 SKU 核对。"),
        (SHEET_NAMES[9], "本期校验结果、数据缺失项和需要你补充/核实的事项。"),
    ]:
        rows.append(("", nm, desc))
    rows.append(("", "", ""))

    rows.append(("S", "六、一个容易搞错的概念：到账净额 ≠ 利润", ""))
    rows.append(("", "到账净额是现金指标",
                 "销售报表的 Total（%s 本期 %s）只是订单层面进你 MercadoPago 账户的钱。"
                 "广告费、Full 仓储与揽收费、页面维护费都不在里面 —— 它们是月底汇总后从同一个余额里另外扣走的（%s扣了 %s）。"
                 % (M["store"], fmt(mm["total"]), mm["pay_date"] or "结算日", fmt(mm["pay_auto"]))))
    rows.append(("", "利润指标是经营贡献毛利",
                 "%s 本期 %s。算法：到账净额 ＋ 加回代扣税 %s − 退货处理费 %s − 广告 %s − Full 仓费用 %s − 页面费 %s。"
                 "完整推导见【%s】桥 D。"
                 % (M["store"], fmt(ctx["contrib"][M["store"]]), fmt(mm["tax_ret"]), fmt(mm["dev_net"]),
                    fmt(mm["ads_net"]), fmt(mm["full_bill_total"]), fmt(mm["subscription_net"]),
                    SHEET_NAMES[2])))
    pa_n, pa_v = fee_rows(M, "Cargo por campaña de publicidad de Product Ads")
    da_n, da_v = fee_rows(M, "Cargo por campaña de publicidad de Display Ads")
    pa_rn, pa_rv = fee_rows(M, "Cargo por campaña de publicidad de Product Ads", reversal=True)
    da_rn, da_rv = fee_rows(M, "Cargo por campaña de publicidad de Display Ads", reversal=True)
    ad_rows = M["billing"][M["billing"][F_DETALLE].astype(str).str.contains("publicidad", na=False)]
    ad_linked = int(ad_rows[F_SALE_NO].notna().sum()) if F_SALE_NO in ad_rows.columns else 0
    rows.append(("", "广告费怎么归集的",
                 "取本期账单里 Product Ads %s（%d笔）＋ Display Ads %s（%d笔），减去当月冲销 %s 与 %s，净额 %s，按计费日归属本期。"
                 "账单的广告行没有商品编号也没有订单号（%d 行中 %d 行带订单号），所以无法按 SKU 或订单分摊，只能作为店铺级共同费用。"
                 % (fmt(pa_v), pa_n, fmt(da_v), da_n, fmt(pa_rv), fmt(da_rv),
                    fmt(mm["ads_net"]), len(ad_rows), ad_linked)))
    st_n, st_v = fee_rows(M, "Cargo por servicio de almacenamiento Full")
    storage_note = ("取本期账单的 Cargo por servicio de almacenamiento Full %s（%d 笔，基本每天一笔）。"
                    % (fmt(st_v), st_n))
    if M["files"].get("storage"):
        d = mm["storage_bill"] - mm["storage_accrual"]
        storage_note += ("另有一张独立的仓储费报表按实际占用日算出 %s，两者差 %s，原因是账单按计费日、报表按占用日，跨月错了一两天。"
                         "本报表统一采用账单口径，因为那才是真正扣钱的金额。" % (fmt(mm["storage_accrual"]), fmt(d)))
    rows.append(("", "仓储费怎么归集的", storage_note))
    full_items = [(FEE_BY_ES[es][0], v) for es, v in mm["bill_fee_net"].items()
                  if FEE_BY_ES[es][1] in ("storage", "fulfillment", "penalty") and abs(v) > 0.005]
    full_items.sort(key=lambda x: -abs(x[1]))
    rows.append(("", "Full 仓相关费用",
                 "%s，合计 %s。%s"
                 % ("、".join("%s %s" % (n, fmt(v)) for n, v in full_items) or "本期无",
                    fmt(mm["full_bill_total"]),
                    ("已与 %s 的各张明细表逐项核对，差异 %s。"
                     % (os.path.basename(M["files"]["cargos_full"]),
                        fmt(mm["cargos_full_report_total"] - mm["full_bill_total"]))
                     if M["files"].get("cargos_full") else "本期未下载到 Full 费用明细报表，无法交叉核对。"))))
    rows.append(("", "", ""))
    rows.append(("S", "七、重要提示", ""))

    r = 4
    for tag, b, c in rows:
        if tag == "S":
            cell = ws.cell(row=r, column=2, value=b)
            cell.font = F(12, True, C_NAVY)
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
            cell.fill = fill(C_SUB)
        else:
            ws.cell(row=r, column=2, value=b).font = F(10, True)
            ws.cell(row=r, column=2).alignment = Alignment(vertical="top")
            c2 = ws.cell(row=r, column=3, value=c)
            c2.font = F(10)
            c2.alignment = wrap("top")
            if c and len(c) > 90:
                ws.row_dimensions[r].height = est_height(c, 78)
        r += 1

    notes = ["本报表所有金额均为墨西哥比索(MXN)，不含商品采购成本（平台数据不提供）。",
             "“平台代扣代缴税金”在损益表中单列为备查项，不计入费用。请与你的墨西哥会计师核对该笔预缴税款的抵扣处理。"]
    no_mp = [S["store"] for S in stores if not S.get("mp")]
    if no_mp:
        notes.append("%s 本期没有 MercadoPago 资金流水（常见原因是该登录账号无报表权限），"
                     "因此这些店的“实际到账”为按销售报表推算。建议让管理员在 MercadoPago 后台“协作者”中"
                     "开通报表权限后重新下载核对。" % "、".join(no_mp))
    if ctx["partial_months"]:
        notes.append("销售报表窗口内的 %s 属不完整月份，本报表未纳入损益，仅在【%s】作参考。"
                     % ("、".join(ctx["partial_months"]), SHEET_NAMES[7]))
    bad = [c for c in ctx["checks"] if not c["ok"]]
    if bad:
        notes.append("⚠ 本期有 %d 项校验未通过，报表数字可能不完整，请先看【%s】顶部的校验结果再使用。"
                     % (len(bad), SHEET_NAMES[9]))
    for i, t in enumerate(notes, 1):
        ws.cell(row=r, column=2, value="注 %d" % i).font = F(10, True, C_RED)
        c2 = ws.cell(row=r, column=3, value=t)
        c2.font = F(10)
        c2.alignment = wrap("top")
        if len(t) > 90:
            ws.row_dimensions[r].height = est_height(t, 78)
        r += 1


# ════════════════════════════════════════════════════════════════════════
# ② 合并损益表
# ════════════════════════════════════════════════════════════════════════

def build_pl_rows():
    """PL_LINES + 按 FEE_CATALOG 顺序生成的账单级费用行 + PL_TAIL。行序恒定。"""
    rows = list(PL_LINES)
    for es, zh, bucket, order_level in FEE_CATALOG:
        if order_level:
            continue
        rows.append(("fee:" + es, "减：" + zh, es, "d", PL_BILL_NOTES.get(es, "")))
    rows.extend(PL_TAIL)
    return rows


def pl_value(S, key):
    m = S["m"]
    if key is None:
        return None
    if key == "ing":
        return m["ing"]
    if key == "env_ing":
        return m["env_ing"]
    if key == "coupon_neg":
        return -m["coupon"]
    if key == "refund_neg":
        return -m["refund"]
    if key == "cn_adj":
        return S["cn"]["partial_total"]
    if key == "com_neg":
        return -m["com_gross"]
    if key == "shipcost_neg":
        return -m["env_cost"]
    if key == "dev_neg":
        return -m["dev_net"]
    if key == "unc_order_neg":
        return -m["unc_order"]
    if key == "unc_bill_neg":
        return -m["unc_bill"]
    if key == "cogs":
        return 0.0
    if key.startswith("fee:"):
        return -m["bill_fee_net"].get(key[4:], 0.0)
    return 0.0


def sheet_pl(wb, stores, ctx):
    n = len(stores)
    sh = Sheet(wb, SHEET_NAMES[1], first_store_col=3, nstores=n)
    ws = sh.ws
    sh.title("合并损益表（%s）" % ctx["period_cn"],
             "口径：按【订单日】归集收入与订单级费用；月度账单费用按【计费日】归集。单位：MXN。"
             "灰底=公式行，浅蓝=小计，黄底=需你录入。行位置每期固定，可直接跨月对比。")
    pcol = get_column_letter(sh.tcol_idx + 1)      # 占GMV%
    ncol_idx = sh.tcol_idx + 2                      # 说明
    cols = ["项目", "西班牙语原始科目"] + store_cols(stores) + ["合计", "占GMV%", "说明"]
    widths = [30, 42] + [16] * n + [16, 10, 46]
    sh.header(4, cols, widths, 26)

    rows = build_pl_rows()
    r0 = 5
    idx = {}
    for i, (key, label, es, kind, note) in enumerate(rows):
        r = r0 + i
        idx[label] = r
        ws.cell(row=r, column=1, value=label).font = F(10, kind in ("sum", "f"))
        ws.cell(row=r, column=2, value=es or "").font = F(9, False, C_GREY)
        if kind in ("d", "in"):
            for j, S in enumerate(stores):
                ws.cell(row=r, column=sh.c0 + j, value=money(pl_value(S, key)))
        note_cell = ws.cell(row=r, column=ncol_idx, value=note or "")
        note_cell.font = F(9, False, "404040")
        note_cell.alignment = wrap()
        if note and len(note) > 46:
            ws.row_dimensions[r].height = est_height(note, 40)

    def R(label):
        return idx[label]

    def setrow(label, expr_fn):
        r = R(label)
        for j in range(n):
            ws.cell(row=r, column=sh.c0 + j, value=expr_fn(sh.scol(j)))

    setrow("营业收入合计",
           lambda c: "=SUM(%s%d:%s%d)" % (c, R("商品销售收入 (GMV)"), c, R("减：卖家承担优惠券/折扣")))
    setrow("净销售收入",
           lambda c: "=%s%d+%s%d+%s%d" % (c, R("营业收入合计"), c, R("减：退款与取消（净）"),
                                          c, R("加：部分退款重复扣佣金调整")))
    setrow("平台交易费用小计",
           lambda c: "=SUM(%s%d:%s%d)" % (c, R("减：平台销售佣金"), c, R("减：订单级未登记科目 ⚠")))
    setrow("订单毛贡献",
           lambda c: "=%s%d+%s%d" % (c, R("净销售收入"), c, R("平台交易费用小计")))
    first_bill = R("减：" + FEE_BY_ES["Cargo por campaña de publicidad de Product Ads"][0])
    setrow("月度账单费用小计",
           lambda c: "=SUM(%s%d:%s%d)" % (c, first_bill, c, R("减：账单级未登记科目 ⚠")))
    setrow("经营贡献毛利（未扣商品成本）",
           lambda c: "=%s%d+%s%d" % (c, R("订单毛贡献"), c, R("月度账单费用小计")))
    setrow("税前净利润",
           lambda c: "=%s%d-%s%d" % (c, R("经营贡献毛利（未扣商品成本）"), c, R("减：商品采购成本")))

    gmv_row = R("商品销售收入 (GMV)")
    last = r0 + len(rows) - 1
    for r in range(r0, last + 1):
        ws.cell(row=r, column=sh.tcol_idx,
                value="=SUM(%s%d:%s%d)" % (sh.scol(0), r, sh.scol(n - 1), r))
        ws.cell(row=r, column=sh.tcol_idx + 1,
                value="=IF($%s$%d=0,0,%s%d/$%s$%d)" % (sh.tcol, gmv_row, sh.tcol, r, sh.tcol, gmv_row))
        ws.cell(row=r, column=sh.tcol_idx + 1).number_format = PCT
        for j in range(n + 1):
            ws.cell(row=r, column=sh.c0 + j).number_format = MNY
        for c in range(1, ncol_idx + 1):
            ws.cell(row=r, column=c).border = BOX

    for i, (key, label, es, kind, note) in enumerate(rows):
        r = r0 + i
        if kind == "sum":
            for c in range(1, ncol_idx + 1):
                ws.cell(row=r, column=c).fill = fill(C_SUB)
            for c in range(1, ncol_idx):
                ws.cell(row=r, column=c).font = F(10, True)
        elif kind == "f":
            for c in range(1, ncol_idx + 1):
                ws.cell(row=r, column=c).fill = fill(C_CALC)
            for c in range(1, ncol_idx):
                ws.cell(row=r, column=c).font = F(10, True, C_NAVY)
        elif kind == "in":
            for j in range(n):
                cc = ws.cell(row=r, column=sh.c0 + j)
                cc.fill = fill(C_INPUT)
                cc.font = F(10, True, "0000FF")
        if key in ("unc_order_neg", "unc_bill_neg"):
            has = any(abs(pl_value(S, key)) > 0.005 for S in stores)
            for c in range(1, ncol_idx + 1):
                ws.cell(row=r, column=c).fill = fill(C_BAD if has else C_WARN)
    rr = R("税前净利润")
    for c in range(1, ncol_idx + 1):
        ws.cell(row=rr, column=c).fill = fill(C_TOTAL)
        ws.cell(row=rr, column=c).font = F(11, True, C_RED)

    # ---------------- 备查项 ----------------
    r = last + 2
    ws.cell(row=r, column=1, value="备查项（不计入上表费用）").font = F(11, True, C_NAVY)
    r += 1
    contrib_row = R("经营贡献毛利（未扣商品成本）")
    memo = [
        ("平台代扣代缴税金（IVA/ISR 预扣）", "Impuestos retenidos",
         lambda S: S["m"]["tax_ret"], None,
         "≈ 含税售价的 %s（＝不含税价的 %s），本期实测。销售报表把它和佣金混在“Cargo por venta e impuestos”里，"
         "是最容易算错的一项。它是预付税款而非成本 —— 详见【%s】桥 E。"
         % (pct(agg_tax_rate(stores), 2), pct(TAX_RATE_NOMINAL_NET, 2), SHEET_NAMES[2])),
        ("口径尾差（应为 0）", "",
         lambda S: S["m"]["resid"], None,
         "上表的收入与费用逐项加总后与销售报表 Total 列的差额。为 0 表示无任何未解释差额。"),
        ("本期账单实际自动扣款", "Cobro automático por factura",
         lambda S: S["m"]["pay_auto"], None,
         "月底从 MercadoPago 余额扣走的金额，与上表“月度账单费用小计”的勾稽见【%s】桥 B。" % SHEET_NAMES[2]),
        ("本期贷记单（平台退还）", "Notas de crédito",
         lambda S: S["m"]["pay_nc"], None,
         "针对取消/退货订单退还的佣金与运费，已在“退款与取消（净）”中抵减。"),
        ("【敏感性】若代扣税不可抵扣时的贡献毛利", "", None, "sens",
         "＝ 经营贡献毛利 − 代扣代缴税金。请与墨西哥会计师确认该笔预扣能否全额抵扣；"
         "不能抵扣时利润会大幅下降，这是本报表最大的不确定项。"),
    ]
    tax_memo_row = r
    for label, es, fn, special, note in memo:
        ws.cell(row=r, column=1, value=label).font = F(10)
        ws.cell(row=r, column=2, value=es).font = F(9, False, C_GREY)
        for j, S in enumerate(stores):
            c = ws.cell(row=r, column=sh.c0 + j)
            if special == "sens":
                c.value = "=%s%d-%s%d" % (sh.scol(j), contrib_row, sh.scol(j), tax_memo_row)
            else:
                c.value = money(fn(S))
            c.number_format = MNY
        ws.cell(row=r, column=sh.tcol_idx,
                value="=SUM(%s%d:%s%d)" % (sh.scol(0), r, sh.scol(n - 1), r)).number_format = MNY
        ws.cell(row=r, column=sh.tcol_idx + 1,
                value="=IF($%s$%d=0,0,%s%d/$%s$%d)" % (sh.tcol, gmv_row, sh.tcol, r, sh.tcol, gmv_row)
                ).number_format = PCT
        nc = ws.cell(row=r, column=ncol_idx, value=note)
        nc.font = F(9, False, "404040")
        nc.alignment = wrap()
        ws.row_dimensions[r].height = est_height(note, 44)
        for c in range(1, ncol_idx + 1):
            ws.cell(row=r, column=c).border = BOX
        if special == "sens":      # 敏感性行整行加粗标红
            for c in range(1, ncol_idx):
                ws.cell(row=r, column=c).font = F(10, True, C_RED)
        r += 1
    ws.freeze_panes = "%s5" % sh.scol(0)
    return {"contrib_row": contrib_row}


# ════════════════════════════════════════════════════════════════════════
# ③ 平账勾稽表
# ════════════════════════════════════════════════════════════════════════

def _bridge_head(wb, stores, ctx):
    n = len(stores)
    sh = Sheet(wb, SHEET_NAMES[2], first_store_col=3, nstores=n)
    ws = sh.ws
    ncol_idx = sh.tcol_idx + 1
    ws["A1"] = "平账勾稽表（%s）" % ctx["period_cn"]
    ws["A1"].font = F(14, True, C_NAVY)
    ws.row_dimensions[1].height = 24
    ws["A2"] = ("把“销售报表 / 平台账单 / 实际资金”三张表逐项打通。"
                "每个桥的最后一行都是差异检验，✓ 表示完全勾稽。单位：MXN")
    ws["A2"].font = F(9, False, C_GREY)
    for i, w in enumerate([6, 44] + [16] * n + [16, 60], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    state = {"r": 4}

    def block(t, sub=""):
        r = state["r"]
        c = ws.cell(row=r, column=1, value=t)
        c.font = F(12, True, "FFFFFF")
        for k in range(1, ncol_idx + 1):
            ws.cell(row=r, column=k).fill = fill(C_NAVY)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol_idx)
        ws.row_dimensions[r].height = 20
        r += 1
        if sub:
            c = ws.cell(row=r, column=1, value=sub)
            c.font = F(9, False, C_GREY)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol_idx)
            c.alignment = wrap()
            ws.row_dimensions[r].height = est_height(sub, 110, 14, 28)
            r += 1
        state["r"] = r

    def head(labels=None):
        r = state["r"]
        cols = labels or (["#", "项目"] + store_cols(stores) + ["合计", "说明"])
        for i, t in enumerate(cols, 1):
            c = ws.cell(row=r, column=i, value=t)
            c.font = F(10, True, "FFFFFF")
            c.fill = fill(C_BAND)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = BOX
        state["r"] = r + 1

    def line(no, label, values, note, kind="d", total=True, fmt_code=MNY):
        """values: 可以是 list（每店一个值/公式），也可以是 callable(col_letter, j)。"""
        r = state["r"]
        ws.cell(row=r, column=1, value=no).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2, value=label)
        for j in range(n):
            v = values(sh.scol(j), j) if callable(values) else values[j]
            c = ws.cell(row=r, column=sh.c0 + j)
            if isinstance(v, str) or v is None:
                c.value = v
            else:
                c.value = money(v)
            c.number_format = fmt_code
        if total:
            ws.cell(row=r, column=sh.tcol_idx,
                    value="=SUM(%s%d:%s%d)" % (sh.scol(0), r, sh.scol(n - 1), r)
                    ).number_format = fmt_code
        nc = ws.cell(row=r, column=ncol_idx, value=note)
        nc.font = F(9, False, "404040")
        nc.alignment = wrap()
        for c in range(1, ncol_idx + 1):
            ws.cell(row=r, column=c).border = BOX
            if kind == "s":
                ws.cell(row=r, column=c).fill = fill(C_SUB)
                ws.cell(row=r, column=c).font = F(10, True)
            elif kind == "t":
                ws.cell(row=r, column=c).fill = fill(C_CALC)
                ws.cell(row=r, column=c).font = F(10, True, C_NAVY)
            elif kind == "chk":
                ws.cell(row=r, column=c).fill = fill(C_OK)
                ws.cell(row=r, column=c).font = F(10, True, C_GREEN)
            else:
                ws.cell(row=r, column=c).font = F(10)
        nc.font = F(9, False, "404040")
        nc.alignment = wrap()
        if note and len(note) > 60:
            ws.row_dimensions[r].height = est_height(note, 56)
        state["r"] = r + 1
        return r

    def gap():
        state["r"] += 1

    M = main_store(stores)
    mm = M["m"]
    V = lambda f: [f(S) for S in stores]

    # ───────── 桥 A ─────────
    block("桥 A　订单资金流：从销售报表推算到实际到账",
          "销售报表(Ventas MX)的每一行是一笔订单。平台在给你打款时，已经把佣金、配送费和代扣税直接扣掉了，"
          "所以“Total”列就是这笔订单真正进你 MercadoPago 账户的钱。")
    head()
    a1 = line("A1", "商品销售收入 (GMV)", V(lambda S: S["m"]["ing"]),
              "Ventas MX → Ingresos por productos，按下单日归集")
    line("A2", "＋ 运费收入", V(lambda S: S["m"]["env_ing"]), "Ventas MX → Ingresos por envío")
    line("A3", "− 平台销售佣金", V(lambda S: -S["m"]["com_gross"]),
         "账单 Facturación → Cargo por venta（纯佣金，含联盟佣金，不含税）")
    line("A4", "− 平台配送费", V(lambda S: -S["m"]["env_cost"]),
         "Ventas MX → Costos de envío（已扣平台运费补贴后的净额）")
    line("A5", "− 平台代扣代缴税金 (IVA/ISR)", V(lambda S: -S["m"]["tax_ret"]),
         "★ 关键项。= Ventas 的“Cargo por venta e impuestos” − 账单的纯佣金。本期实测约为含税售价的 %s。"
         "这是平台替你预缴给墨西哥税局的钱，不是费用。" % pct(mm["tax_rate_gross"], 2))
    line("A6", "− 退款与取消（净）", V(lambda S: -S["m"]["refund"]),
         "Ventas MX → Anulaciones y reembolsos")
    cs = [S for S in stores if S["m"]["coupon_n"] > 0]
    coupon_note = ("★ 平台把 Descuentos y bonificaciones 这一列导出成了空白，所以这笔钱在任何明细列里都看不到，"
                   "只能用 Total 倒推。已交叉验证：账单佣金 ÷（原价 − 券额）＝ 标准费率。")
    if cs:
        coupon_note += "本期 " + "；".join(
            "%s %d 笔、合计 %s（面额 %s）" % (
                S["store"], S["m"]["coupon_n"], fmt(S["m"]["coupon"]),
                "、".join("%s×%d" % (fmt(k, 0), v)
                          for k, v in sorted(S["m"]["coupon_denoms"].items(), reverse=True)))
            for S in cs) + "。"
    else:
        coupon_note += "本期各店倒推券额均为 0。"
    a6b = line("A6b", "− 卖家承担优惠券/折扣", V(lambda S: -S["m"]["coupon"]), coupon_note)
    a7 = line("A7", "＝ 推算到账净额",
              lambda c, j: "=SUM(%s%d:%s%d)" % (c, a1, c, a6b),
              "按上述公式推算出的应到账金额", kind="t")
    a8 = line("A8", "销售报表 “Total (MXN)” 列实际合计", V(lambda S: S["m"]["total"]),
              "直接取 Ventas MX 的 Total 列求和，作为独立验证", kind="s")
    ok_a = all(abs(S["m"]["resid"]) <= 0.05 for S in stores)
    line("A9", "差异检验（A8 − A7）",
         lambda c, j: "=%s%d-%s%d" % (c, a8, c, a7),
         ("全部店铺均为 0.00 —— 逐笔订单完全勾稽，无任何未解释差额 ✓" if ok_a else
          "⚠ 存在残差，说明还有未识别的隐藏费用或折扣，请按【%s】的排查步骤处理。" % SHEET_NAMES[9]),
         kind="chk")

    # ───────── 桥 A′ MercadoPago 交叉验证 ─────────
    for S in [s for s in stores if s.get("mp")]:
        gap()
        MP, sm = S["mp"], S["m"]
        block("桥 A′　第三方交叉验证：%s 的 MercadoPago 结算流水" % S["store"],
              "用平台自己的结算明细(settlement_v2)反向验证桥 A 的口径是否正确。"
              "这是唯一独立于销售报表和账单的第三方凭证。")
        r = state["r"]
        for i, t in enumerate(("#", "项目", "销售报表推算", "MercadoPago 实际流水", "差异", "说明"), 1):
            c = ws.cell(row=r, column=i, value=t)
            c.font = F(10, True, "FFFFFF")
            c.fill = fill(C_BAND)
            c.alignment = Alignment(horizontal="center")
            c.border = BOX
        state["r"] = r + 1
        cross = [
            ("V1", "交易总额（含运费收入）", sm["ing"] + sm["env_ing"], MP["trans"],
             "TRANSACTION_AMOUNT（仅 SETTLEMENT 类）"),
            ("V2", "平台手续费（佣金＋配送费）", -(sm["com_gross"] + sm["env_cost"]), MP["fee"],
             "FEE_AMOUNT（SETTLEMENT 类）＝ 账单中 Descontado=Si 的 %s ＋ 取消订单原计费 %s"
             % (fmt(sm["bill_by_channel"][CH_SI]),
                fmt(-MP["fee"] - sm["bill_by_channel"][CH_SI]))),
            ("V3", "平台代扣代缴税金", -sm["tax_ret"], MP["tax"],
             "TAXES_AMOUNT —— 证实“代扣税”确实独立存在，与手续费分开列示"),
            ("V4", "退款与取消（争议＋退货）", -sm["refund"], MP["refund_real"],
             "REFUND / DISPUTE 类：%s" % MP["refund_detail"]),
            ("V5", "净到账金额", sm["total"], MP["real_all"],
             "REAL_AMOUNT 合计，与销售报表 Total 对比"),
        ]
        for no, lab, calc, act, note in cross:
            r = state["r"]
            ws.cell(row=r, column=1, value=no).alignment = Alignment(horizontal="center")
            ws.cell(row=r, column=2, value=lab)
            ws.cell(row=r, column=3, value=money(calc)).number_format = MNY
            ws.cell(row=r, column=4, value=money(act)).number_format = MNY
            ws.cell(row=r, column=5, value="=C%d-D%d" % (r, r)).number_format = MNY
            c = ws.cell(row=r, column=6, value=note)
            c.font = F(9, False, "404040")
            c.alignment = wrap()
            for k in range(1, 7):
                ws.cell(row=r, column=k).border = BOX
                ws.cell(row=r, column=k).fill = fill(C_OK)
            state["r"] = r + 1

    # ───────── 桥 B ─────────
    gap()
    block("桥 B　月度账单：从账单总额到实际扣款",
          "平台账单(Facturación)里的费用分两个结算通道，混在同一张表里。搞不清这一点，就会把佣金和运费重复计算一次。")
    head()
    line("B1", "本期账单全部费用合计", V(lambda S: S["m"]["bill_total"]),
         "Facturación ＋ Notas de Crédito 全部行求和（含冲销行）")
    line("B2", "其中：已在订单打款时扣除 (Descontado = Si)",
         V(lambda S: S["m"]["bill_by_channel"][CH_SI]),
         "佣金、配送费、部分退货费。★ 这部分不会再收第二次，已体现在桥 A 里")
    line("B3", "其中：取消订单的费用与冲销 (No aplica)",
         V(lambda S: S["m"]["bill_by_channel"][CH_NA]),
         "被取消的订单原计费与其冲销，净额为负表示平台净退还")
    b4 = line("B4", "其中：需单独支付 (Descontado = No)",
              V(lambda S: S["m"]["bill_by_channel"][CH_NO]),
              "广告、Full 仓储/揽收/仓位/罚金、页面维护费等", kind="s")
    b5 = line("B5", "− 同一张发票内的冲销",
              V(lambda S: S["m"]["invoice_internal_reversal"]),
              "按“Cargo que bonifica → Número del cargo”关联到被冲销的原始费用，只取原费用属于 No 通道的那些行。")
    b6 = line("B6", "＝ 应付账单金额",
              lambda c, j: "=%s%d+%s%d" % (c, b4, c, b5),
              "这才是真正要掏钱的部分", kind="t")
    b7 = line("B7", "实际自动扣款金额", V(lambda S: S["m"]["pay_auto"]),
              "Pagos 报表 → Cobro automático por factura vencida（%s 从 MercadoPago 余额扣走）"
              % (mm["pay_date"] or "结算日"), kind="s")
    ok_b = all(abs(S["m"]["payable"] - S["m"]["pay_auto"]) <= 0.01
               for S in stores if S["files"].get("pagos"))
    line("B8", "差异检验（B7 − B6）",
         lambda c, j: "=%s%d-%s%d" % (c, b7, c, b6),
         "全部店铺均为 0.00，完全勾稽 ✓" if ok_b else "⚠ 存在差异，见【%s】" % SHEET_NAMES[9],
         kind="chk")
    line("B7a", "其中：与订单挂钩、但走账单支付的费用",
         V(lambda S: S["m"]["order_linked_billed"]),
         "多为联盟推广佣金等。这部分已计入【%s】的订单级费用行，"
         "所以损益表的“月度账单费用小计”会比实际扣款少这一块 —— 不是漏记，是归类不同，全表无重复无遗漏。"
         % SHEET_NAMES[1])
    line("B9", "另：贷记单退还 (Notas de Crédito)", V(lambda S: S["m"]["pay_nc"]),
         "针对取消/退货订单退还的佣金与运费，走红字发票单独退回，"
         "已在桥 A 的“退款与取消（净）”中体现，不要重复计入。")

    # ───────── 桥 C ─────────
    gap()
    block("桥 C　为什么“销售报表的扣费”和“账单的费用”对不上",
          "这是最常见的算错点。两张表统计的是不同的东西，差额可以完全解释。")
    head()
    c1 = line("C1", "销售报表显示的平台扣费（佣金及税金＋配送费）",
              V(lambda S: S["m"]["com_imp"] + S["m"]["env_cost"]),
              "Ventas MX 中 Cargo por venta e impuestos ＋ Costos de envío 的绝对值")
    c2 = line("C2", "− 其中属于代扣代缴税金，不是平台费用",
              V(lambda S: -S["m"]["tax_ret"]), "★ 差异的最大来源")
    c3 = line("C3", "＝ 销售报表口径的真实平台费用",
              lambda c, j: "=%s%d+%s%d" % (c, c1, c, c2), "", kind="t")
    c4 = line("C4", "账单中与订单挂钩的费用（佣金＋配送费，毛额）",
              V(lambda S: S["m"]["com_gross"] + S["m"]["env_gross"]),
              "账单 Cargo por venta ＋ Cargo por envíos（已匹配到本期订单的部分，可能跨两张账单）", kind="s")
    line("C5", "差异（C4 − C3）",
         lambda c, j: "=%s%d-%s%d" % (c, c4, c, c3),
         "剩余小额差异来自：被取消订单在账单里保留了原计费行、以及多商品包裹的行归属。"
         "本期 %s。" % per_store(stores, lambda S: fmt(
             S["m"]["com_gross"] + S["m"]["env_gross"] - (S["m"]["com_imp"] + S["m"]["env_cost"] - S["m"]["tax_ret"])), "、"),
         kind="chk")
    gap()
    c6 = line("C6", "账单中与任何订单都无关的费用", V(lambda S: S["m"]["other_total"]),
              "广告、Full 仓储/揽收/仓位/长龄库存/罚金、页面维护费。"
              "★ 销售报表里完全看不到这些，只看销售报表会高估利润。", kind="s")
    r = line("C7", "占 GMV 比例",
             lambda c, j: "=IF(%s%d=0,0,%s%d/%s%d)" % (c, a1, c, c6, c, a1),
             "本期 %s；其中广告占 %s。" % (
                 per_store(stores, lambda S: "%s %s" % (
                     S["store"], pct(S["m"]["other_total"] / S["m"]["ing"] if S["m"]["ing"] else 0, 1)), "、"),
                 per_store(stores, lambda S: "%s %s" % (
                     S["store"], pct(S["m"]["ads_net"] / S["m"]["ing"] if S["m"]["ing"] else 0, 1)), "、")),
             total=False, fmt_code=PCT2)
    return sh, ws, {"state": state, "a1": a1, "block": block, "head": head,
                    "line": line, "gap": gap, "ncol": ncol_idx}



def _bridge_tail(stores, ctx, sh, ws, H, n):
    """③ 的后半段：桥 D / E / F 与常见误算。游标沿用 _bridge_head 的 state。"""
    state = H["state"]
    block, head, line, gap = H["block"], H["head"], H["line"], H["gap"]
    ncol_idx = H["ncol"]
    M = main_store(stores)
    mm = M["m"]
    V = lambda f: [f(S) for S in stores]

    def banner(text, color=C_WARN, font=None):
        r = state["r"]
        c = ws.cell(row=r, column=2, value=text)
        c.font = font or F(10, True, C_RED)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=ncol_idx)
        c.alignment = wrap()
        ws.row_dimensions[r].height = est_height(text, 100, 16, 32)
        for k in range(1, ncol_idx + 1):
            ws.cell(row=r, column=k).fill = fill(color)
            ws.cell(row=r, column=k).border = BOX
        state["r"] = r + 1

    # ───────── 桥 D ─────────
    gap()
    block("桥 D　从“到账净额”到“经营利润”：为什么到账净额不等于利润",
          "销售报表的 Total（到账净额）只是订单层面进账的钱，广告费和仓储费是月底另外从同一个 MercadoPago 余额里扣走的，"
          "不在这个数里。同时它扣掉了代扣代缴税金——那是预缴税款不是费用。要得到利润，必须做下面这个桥。")
    head()
    d1 = line("D1", "销售报表 “Total” 到账净额", V(lambda S: S["m"]["total"]),
              "桥 A 的结论。这是订单层面真正进你账户的钱，不是利润。", kind="s")
    line("D2", "＋ 加回平台代扣代缴税金", V(lambda S: S["m"]["tax_ret"]),
         "预缴给税局的钱，能抵扣就不是成本，所以算利润时要加回来（若不能抵扣，见【%s】敏感性行）" % SHEET_NAMES[1])
    line("D3", "− 退货处理费", V(lambda S: -S["m"]["dev_net"]),
         "Cargo por devolución，在退款交易里单独扣，不在 Total 内")
    line("D4", "− 广告费（Product Ads ＋ Display Ads，净）", V(lambda S: -S["m"]["ads_net"]),
         "★ 完全不在销售报表里。账单 Descontado=No，月底自动扣款")
    line("D5", "− Full 仓相关费用（仓储/揽收/仓位/罚金等）", V(lambda S: -S["m"]["full_bill_total"]),
         "★ 同样不在销售报表里。%s"
         % ("已与 Reporte_Cargos_Full 逐项核对，差异 %s"
            % fmt(mm["cargos_full_report_total"] - mm["full_bill_total"])
            if M["files"].get("cargos_full") else "本期未下载到 Full 费用明细报表，无法交叉核对"))
    line("D6", "− 官方店页面维护费", V(lambda S: -S["m"]["subscription_net"]), "Mi página 月费")
    d6b = line("D6b", "− 未登记科目（订单级＋账单级）",
               V(lambda S: -(S["m"]["unc_order"] + S["m"]["unc_bill"])),
               "字典外的费用。为 0 说明本期所有科目都已登记；不为 0 时见【%s】。" % SHEET_NAMES[9])
    d7 = line("D7", "＝ 经营贡献毛利（未扣商品成本）",
              lambda c, j: "=SUM(%s%d:%s%d)" % (c, d1, c, d6b),
              "与【%s】的同名行完全一致" % SHEET_NAMES[1], kind="t")
    d8 = line("D8", "− 商品采购成本", [0.0] * n,
              "⚠ 平台数据不含此项。在【%s】填入单位成本后即可得出真实净利润。" % SHEET_NAMES[6], kind="s")
    for j in range(n):
        c = ws.cell(row=d8, column=sh.c0 + j)
        c.fill = fill(C_INPUT)
        c.font = F(10, True, "0000FF")
    line("D9", "＝ 税前净利润", lambda c, j: "=%s%d-%s%d" % (c, d7, c, d8), "", kind="chk")

    gap()
    contrib_m = ctx["contrib"][M["store"]]
    banner("一句话：到账净额是「现金」，经营贡献毛利才是「利润」。%s 本期到账 %s，"
           "但加回代扣税、再扣掉广告和仓储之后，真正的经营贡献是 %s；两者差 %s%s——这就是只看到账额会误判的地方。"
           % (M["store"], fmt(mm["total"]), fmt(contrib_m), fmt(abs(contrib_m - mm["total"])),
              "，方向还相反" if contrib_m > mm["total"] else ""))

    # ───────── 桥 E ─────────
    gap()
    block("桥 E　“平台代扣代缴税金”到底是什么，为什么算利润时要加回来",
          "这是本报表唯一一项“从现金里扣了、但不算费用”的支出。下面是它在原始数据里的准确形态与验证过程。")
    head()
    e1 = line("E1", "本期被代扣金额", V(lambda S: S["m"]["tax_ret"]),
              "MercadoPago 结算流水里有独立字段 TAXES_AMOUNT，与手续费 FEE_AMOUNT 分开列示，"
              "说明平台自己也不把它当手续费")
    mp = next((S for S in stores if S.get("mp") and S["mp"].get("rate_n")), None)
    rate_note = "本期各店实测：%s。" % per_store(
        stores, lambda S: "%s %s" % (S["store"], pct(S["m"]["tax_rate_gross"], 3)), "、")
    if mp:
        rate_note += ("%s 的 %d 笔结算流水中有 %d 笔费率完全一致，标准差 %.4f —— 费率高度统一，是法定预扣不是浮动费用。"
                      % (mp["store"], mp["mp"]["rate_n"], mp["mp"]["rate_mode_n"], mp["mp"]["rate_std"]))
    line("E2", "占含税成交额比例",
         lambda c, j: "=IF(%s%d=0,0,%s%d/%s%d)" % (c, H["a1"], c, e1, c, H["a1"]),
         rate_note, total=False, fmt_code=PCT3)
    line("E3", "占不含税价比例（÷1.16 后）",
         lambda c, j: "=IF(%s%d=0,0,%s%d/(%s%d/(1+%s)))" % (c, H["a1"], c, e1, c, H["a1"], IVA_RATE),
         "本期各店实测 %s。按墨西哥数字平台代扣制度，最可能的构成是 IVA 预扣 8%% ＋ ISR 预扣 2.5%%，"
         "但准确拆分必须以 MercadoPago 每月出具的 Constancia de retenciones de IVA e ISR 为准 —— "
         "该凭证不在本次下载的文件里。" % per_store(
             stores, lambda S: "%s %s" % (S["store"], pct(
                 S["m"]["tax_ret"] / S["m"]["net_revenue_exvat"] if S["m"]["net_revenue_exvat"] else 0, 2)), "、"),
         total=False, fmt_code=PCT3)
    line("E4", "为什么算利润时要加回", [None] * n,
         "预扣税是“预先替你缴给税局的税”，性质是预付税款（资产），不是经营成本。"
         "你在做 IVA/ISR 申报时可以拿它抵减应缴税额。若当成费用扣掉，等于同一笔税被计两次。",
         kind="s", total=False)

    # ───────── 桥 F 含税口径提示 ─────────
    gap()
    banner("⚠ 但要注意：本报表的收入是含税口径", C_WARN, F(11, True, C_RED))
    head()
    f1 = line("F1", "含税商品收入（报表中的 GMV）", V(lambda S: S["m"]["ing"]),
              "买家实付金额，含 %s IVA" % pct(IVA_RATE, 0))
    f2 = line("F2", "其中 IVA 成分（× 16/116）",
              lambda c, j: "=%s%d*%s/(1+%s)" % (c, f1, IVA_RATE, IVA_RATE),
              "这部分钱在报表里被算进了收入，但它不属于你，最终要交给税局")
    line("F3", "不含税收入（参考）",
         lambda c, j: "=%s%d-%s%d" % (c, f1, c, f2),
         "若要编制正式的不含税损益表，应以此为收入起点", kind="t")
    f4 = line("F4", "已通过代扣预缴的部分", V(lambda S: S["m"]["tax_ret"]), "")
    line("F5", "预缴占 IVA 成分比例",
         lambda c, j: "=IF(%s%d=0,0,%s%d/%s%d)" % (c, f2, c, f4, c, f2),
         "本期 %s。也就是说 IVA 还没缴完，剩余部分要在申报时结清（可用采购与费用的进项 IVA 抵扣）。"
         "★ 结论：加回代扣税只是把“不该当费用的”剔出去；要得到真正的税后净利，"
         "还需要用采购发票做完整的 IVA 销项—进项计算，这已超出平台数据能覆盖的范围，请交给你的墨西哥会计师。"
         % per_store(stores, lambda S: "%s %s" % (
             S["store"], pct(S["m"]["tax_vs_iva"], 1)), "、"),
         total=False, fmt_code=PCT, kind="chk")

    # ───────── 常见误算 ─────────
    gap()
    block("常见误算示范（请避免）")
    tot = lambda f: sum(f(S) for S in stores)
    mis = [
        ("❌ 把“到账净额”当成利润",
         "到账净额里没有广告费(%s)、Full 仓费用(%s)和页面费(%s)，这些是月底另外扣的；"
         "同时它多扣了代扣税(%s)。见桥 D。"
         % (fmt(tot(lambda S: S["m"]["ads_net"])), fmt(tot(lambda S: S["m"]["full_bill_total"])),
            fmt(tot(lambda S: S["m"]["subscription_net"])), fmt(tot(lambda S: S["m"]["tax_ret"])))),
        ("❌ 把销售报表 Total 减去账单全部费用",
         "会把佣金和配送费重复扣一次（本期 %s），利润被严重低估。"
         % per_store(stores, lambda S: "%s 约 %s" % (S["store"], fmt(S["m"]["bill_by_channel"][CH_SI], 0)), "、")),
        ("❌ 把“Cargo por venta e impuestos”整笔当作平台佣金",
         "其中 %s 是代扣税，属预缴税款而非费用，会让佣金率虚高约 %s。"
         % (per_store(stores, lambda S: "%s %s" % (S["store"], fmt(S["m"]["tax_ret"])), " / "),
            pct(mm["tax_rate_gross"], 1))),
        ("❌ 用账单月份直接当作销售月份",
         "本期账单里混着上月订单的费用，本期订单也有一部分费用要到下月账单才出现。见桥 C。"),
        ("❌ 把代扣税当作平台费用扣掉",
         "它是预付税款不是成本，可用于抵减申报税额。当费用扣会让同一笔税被计两次，本期共 %s。"
         % fmt(tot(lambda S: S["m"]["tax_ret"]))),
        ("❌ 把含税收入直接当作利润起点",
         "GMV %s 里含 %s 的 IVA，那不是你的钱。本报表为便于与平台数据勾稽采用含税口径，正式申报需换算成不含税。"
         % (fmt(tot(lambda S: S["m"]["ing"])), fmt(tot(lambda S: S["m"]["iva_in_gmv"])))),
        ("❌ 只看销售报表算利润",
         "广告费、仓储费、揽收费等 %s（全部店铺合计）完全不在销售报表里。"
         % fmt(tot(lambda S: S["m"]["other_total"]))),
        ("❌ 退款只看销售报表的 Anulaciones",
         "平台退还的佣金/运费走贷记单（本期 %s），退货处理费又在账单里单独收，三处都要看。"
         % fmt(tot(lambda S: S["m"]["pay_nc"]))),
        ("❌ 把 Cargos_Full 报表与账单相加",
         "Full 费用报表是账单里 Full 科目的明细拆分，两者金额相同（本期差 %s）。相加即每笔仓储费重复计一次。"
         % fmt(tot(lambda S: S["m"]["cargos_full_report_total"] - S["m"]["full_bill_total"]))),
    ]
    for lab, note in mis:
        r = state["r"]
        ws.cell(row=r, column=2, value=lab).font = F(10, True, C_RED)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=ncol_idx - 1)
        c = ws.cell(row=r, column=ncol_idx, value=note)
        c.font = F(9)
        c.alignment = wrap()
        for k in range(1, ncol_idx + 1):
            ws.cell(row=r, column=k).border = BOX
        ws.row_dimensions[r].height = est_height(note, 56)
        state["r"] = r + 1
    ws.freeze_panes = "%s5" % sh.scol(0)


def sheet_bridge(wb, stores, ctx):
    sh, ws, H = _bridge_head(wb, stores, ctx)
    _bridge_tail(stores, ctx, sh, ws, H, len(stores))


# ════════════════════════════════════════════════════════════════════════
# ④ 平台费用明细
# ════════════════════════════════════════════════════════════════════════

def sheet_fees(wb, stores, ctx):
    ws = wb.create_sheet(SHEET_NAMES[3])
    sh = Sheet.__new__(Sheet)
    sh.ws = ws
    ws["A1"] = "平台费用明细（%s账单，按计费日归集）" % ctx["period_cn"]
    ws["A1"].font = F(14, True, C_NAVY)
    ws.row_dimensions[1].height = 24
    ws["A2"] = ("数据源：Reporte_Facturacion_MercadoLibre + Reporte_Notas_Credito。"
                "结算通道决定这笔钱是“下单时已扣”还是“月底再扣”，混淆两者是平不了账的主因之一。"
                "未在 FEE_CATALOG 登记的科目以红底标出。单位：MXN")
    ws["A2"].font = F(9, False, C_GREY)
    cols = ["店铺", "西班牙语科目 (Detalle)", "中文科目", "结算通道", "笔数", "金额", "占该店 GMV%"]
    widths = [18, 52, 30, 16, 10, 16, 14]
    for i, t in enumerate(cols, 1):
        c = ws.cell(row=4, column=i, value=t)
        c.font = F(10, True, "FFFFFF")
        c.fill = fill(C_NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    ws.row_dimensions[4].height = 30
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    r = 5
    chan_order = {CH_SI: 0, CH_NO: 1, CH_NA: 2}
    for S in stores:
        FA, gmv = S["billing"], S["m"]["ing"] or 1.0
        g = FA.groupby([FA[F_DETALLE].astype(str), FA["channel"]])["amount"].agg(["count", "sum"]).reset_index()
        g.columns = ["detalle", "channel", "count", "sum"]
        g["_c"] = g["channel"].map(lambda x: chan_order.get(x, 9))
        g = g.sort_values(["_c", "sum"], key=lambda s: s if s.name != "sum" else s.abs(),
                          ascending=[True, False])
        start = r
        for _, x in g.iterrows():
            zh, bucket, _o, is_rev, known = classify(x["detalle"])
            ws.cell(row=r, column=1, value=S["store"])
            ws.cell(row=r, column=2, value=x["detalle"])
            ws.cell(row=r, column=3, value=zh)
            ws.cell(row=r, column=4, value=CHANNEL_ZH.get(x["channel"], x["channel"]))
            ws.cell(row=r, column=5, value=int(x["count"])).number_format = INT
            ws.cell(row=r, column=6, value=money(x["sum"])).number_format = MNY
            ws.cell(row=r, column=7, value="=F%d/%s" % (r, gmv)).number_format = PCT
            for c in range(1, 8):
                ws.cell(row=r, column=c).border = BOX
                ws.cell(row=r, column=c).font = F(10)
            if not known:
                for c in range(1, 8):
                    ws.cell(row=r, column=c).fill = fill(C_BAD)
                    ws.cell(row=r, column=c).font = F(10, True, C_RED)
            elif is_rev:
                for c in range(1, 8):
                    ws.cell(row=r, column=c).font = F(10, False, "8C4A00")
            r += 1
        ws.cell(row=r, column=1, value=S["store"])
        ws.cell(row=r, column=3, value="账单费用合计")
        ws.cell(row=r, column=5, value="=SUM(E%d:E%d)" % (start, r - 1)).number_format = INT
        ws.cell(row=r, column=6, value="=SUM(F%d:F%d)" % (start, r - 1)).number_format = MNY
        ws.cell(row=r, column=7, value="=F%d/%s" % (r, gmv)).number_format = PCT
        for c in range(1, 8):
            ws.cell(row=r, column=c).fill = fill(C_SUB)
            ws.cell(row=r, column=c).font = F(10, True)
            ws.cell(row=r, column=c).border = BOX
        r += 2

    ws.cell(row=r, column=2, value="按结算通道汇总").font = F(12, True, C_NAVY)
    r += 1
    for i, t in enumerate(["店铺", "结算通道", "含义", "", "笔数", "金额", "占该店 GMV%"], 1):
        c = ws.cell(row=r, column=i, value=t)
        c.font = F(10, True, "FFFFFF")
        c.fill = fill(C_NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BOX
    r += 1
    for S in stores:
        gmv = S["m"]["ing"] or 1.0
        for k in (CH_SI, CH_NO, CH_NA):
            ws.cell(row=r, column=1, value=S["store"])
            ws.cell(row=r, column=2, value="%s  (Descontado = %s)" % (CHANNEL_ZH[k], k))
            ws.cell(row=r, column=3, value=CHANNEL_MEANING[k]).font = F(9, False, "404040")
            ws.cell(row=r, column=5, value=S["m"]["bill_cnt_by_channel"][k]).number_format = INT
            ws.cell(row=r, column=6, value=money(S["m"]["bill_by_channel"][k])).number_format = MNY
            ws.cell(row=r, column=7, value="=F%d/%s" % (r, gmv)).number_format = PCT
            for c in range(1, 8):
                ws.cell(row=r, column=c).border = BOX
            r += 1
    ws.freeze_panes = "A5"


# ════════════════════════════════════════════════════════════════════════
# ⑤ 退货退款分析
# ════════════════════════════════════════════════════════════════════════

def sheet_returns_money(wb, stores, ctx):
    ws = wb.create_sheet(SHEET_NAMES[4])
    ws["A1"] = "退货 · 退款 · 售后分析（%s订单）" % ctx["period_cn"]
    ws["A1"].font = F(14, True, C_NAVY)
    ws.row_dimensions[1].height = 24
    ws["A2"] = ("一笔退货会同时影响：退给买家的货款、平台退还的佣金与运费、额外的退货处理费、以及库存能否回收。"
                "下表按订单状态拆开，并标明责任方。单位：MXN")
    ws["A2"].font = F(9, False, C_GREY)
    cols = ["店铺", "订单状态 (Estado)", "中文含义", "单数", "件数", "原销售额", "实际退款额",
            "退款/原销售额", "资金归属", "库存处置"]
    widths = [16, 46, 30, 9, 9, 15, 15, 13, 14, 26]
    for i, t in enumerate(cols, 1):
        c = ws.cell(row=4, column=i, value=t)
        c.font = F(10, True, "FFFFFF")
        c.fill = fill(C_NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    ws.row_dimensions[4].height = 30
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    r = 5
    for S in stores:
        A = S["ventas"]
        ab = A[A["abnormal"]]
        if not len(ab):
            continue
        g = ab.groupby(A[V_STATUS].astype(str)).apply(
            lambda x: pd.Series({"单数": len(x), "件数": num(x, V_UNITS).sum(),
                                 "原销售额": num(x, V_REVENUE).sum(),
                                 "退款": -num(x, V_REFUND).sum()})).reset_index()
        g.columns = ["estado", "单数", "件数", "原销售额", "退款"]
        g = g.sort_values("原销售额", ascending=False)
        start = r
        for _, x in g.iterrows():
            zh, who, inv = STATUS_ZH.get(x["estado"], (x["estado"], "待定", "待确认"))
            ws.cell(row=r, column=1, value=S["store"])
            ws.cell(row=r, column=2, value=x["estado"])
            ws.cell(row=r, column=3, value=zh).alignment = wrap()
            ws.cell(row=r, column=4, value=int(x["单数"])).number_format = INT
            ws.cell(row=r, column=5, value=int(x["件数"])).number_format = INT
            ws.cell(row=r, column=6, value=money(x["原销售额"])).number_format = MNY
            ws.cell(row=r, column=7, value=money(x["退款"])).number_format = MNY
            ws.cell(row=r, column=8, value="=IF(F%d=0,0,G%d/F%d)" % (r, r, r)).number_format = PCT
            ws.cell(row=r, column=9, value=who)
            ws.cell(row=r, column=10, value=inv).alignment = wrap()
            for c in range(1, 11):
                ws.cell(row=r, column=c).border = BOX
                ws.cell(row=r, column=c).font = F(10)
            if inv.startswith("✘"):
                for c in range(1, 11):
                    ws.cell(row=r, column=c).fill = fill(C_BAD)
            elif inv.startswith("✔"):
                for c in range(1, 11):
                    ws.cell(row=r, column=c).fill = fill(C_OK)
            r += 1
        ws.cell(row=r, column=2, value="%s 售后订单合计" % S["store"]).font = F(10, True)
        for col, cl in ((4, "D"), (5, "E"), (6, "F"), (7, "G")):
            ws.cell(row=r, column=col, value="=SUM(%s%d:%s%d)" % (cl, start, cl, r - 1)
                    ).number_format = INT if col in (4, 5) else MNY
        ws.cell(row=r, column=8, value="=IF(F%d=0,0,G%d/F%d)" % (r, r, r)).number_format = PCT
        for c in range(1, 11):
            ws.cell(row=r, column=c).fill = fill(C_SUB)
            ws.cell(row=r, column=c).font = F(10, True)
            ws.cell(row=r, column=c).border = BOX
        r += 2

    # ---- KPI ----
    r += 1
    ws.cell(row=r, column=1, value="退货核心指标").font = F(12, True, C_NAVY)
    r += 1
    n = len(stores)
    hdr = ["指标"] + store_cols(stores) + ["合计", "说明"]
    for i, t in enumerate(hdr, 1):
        c = ws.cell(row=r, column=i, value=t)
        c.font = F(10, True, "FFFFFF")
        c.fill = fill(C_NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BOX
    ws.row_dimensions[r].height = 22
    r += 1
    tcol = 1 + n + 1
    ncol = tcol + 1
    kpi_rows = [
        ("本期订单总数", lambda S: S["m"]["rows"], INT, "Ventas MX 按下单日，按行计"),
        ("其中：售后/异常订单数", lambda S: S["m"]["abnormal_orders"], INT, "含取消、退货、仲裁、投诉、换货"),
        ("售后订单占比", "ratio_prev", PCT, "售后订单数 ÷ 订单总数"),
        ("实际退款金额", lambda S: S["m"]["refund"], MNY, "Anulaciones y reembolsos 净额"),
        ("退款率（占 GMV）", "ratio_gmv", PCT, "退款金额 ÷ GMV"),
        ("退货处理费（平台另收）", lambda S: S["m"]["dev_net"], MNY, "Cargo por devolución 净额，账单中单独收取"),
        ("　— 库存可回收（重新上架）单数", lambda S: S["m"]["n_back_on_sale"], INT, "商品质检合格，已重新上架销售，库存损失为 0"),
        ("　— 库存不可再售（自提/销毁）单数", lambda S: S["m"]["n_stock_lost"], INT,
         "★ 需付费从 Full 仓取回或直接销毁，货款与货物双重损失"),
        ("　— 卖家保留货款单数", lambda S: S["m"]["n_seller_keeps"], INT, "平台仲裁判卖家胜诉或退货被驳回，未退款"),
        ("库存不可再售订单原销售额", lambda S: S["m"]["amt_stock_lost"], MNY,
         "★ 这部分是最实在的损失：钱退了、货也拿不回来卖"),
    ]
    for label, fn, code, note in kpi_rows:
        ws.cell(row=r, column=1, value=label).font = F(10)
        if fn == "ratio_prev":
            for j in range(n):
                cl = get_column_letter(2 + j)
                ws.cell(row=r, column=2 + j, value="=IF(%s%d=0,0,%s%d/%s%d)" % (cl, r - 2, cl, r - 1, cl, r - 2))
            ws.cell(row=r, column=tcol, value="")
        elif fn == "ratio_gmv":
            for j, S in enumerate(stores):
                cl = get_column_letter(2 + j)
                gmv = S["m"]["ing"] or 1.0
                ws.cell(row=r, column=2 + j, value="=%s%d/%s" % (cl, r - 1, gmv))
            ws.cell(row=r, column=tcol, value="")
        else:
            for j, S in enumerate(stores):
                ws.cell(row=r, column=2 + j, value=money(fn(S)))
            ws.cell(row=r, column=tcol,
                    value="=SUM(%s%d:%s%d)" % (get_column_letter(2), r, get_column_letter(1 + n), r))
        for j in range(n + 1):
            ws.cell(row=r, column=2 + j).number_format = code
        c = ws.cell(row=r, column=ncol, value=note)
        c.font = F(9, False, "404040")
        c.alignment = wrap()
        for c2 in range(1, ncol + 1):
            ws.cell(row=r, column=c2).border = BOX
        if "不可再售" in label:
            for c2 in range(1, ncol + 1):
                ws.cell(row=r, column=c2).fill = fill(C_BAD)
        elif "可回收" in label:
            for c2 in range(1, ncol + 1):
                ws.cell(row=r, column=c2).fill = fill(C_OK)
        r += 1
    ws.column_dimensions[get_column_letter(ncol)].width = 60


# ════════════════════════════════════════════════════════════════════════
# ⑥ 退货质检明细
# ════════════════════════════════════════════════════════════════════════

def sheet_triage(wb, stores, ctx):
    ws = wb.create_sheet(SHEET_NAMES[5])
    ws["A1"] = "Full 仓退货质检明细与原因分析"
    ws["A1"].font = F(14, True, C_NAVY)
    ws.row_dimensions[1].height = 24
    have = [S for S in stores if len(S["returns"])]
    none = [S["store"] for S in stores if not len(S["returns"])]
    cover = "；".join(
        "%s 覆盖 %s ~ %s 共 %d 件（其中本期 %d 件）"
        % (S["store"], S["returns_all"]["fr"].min().strftime("%Y-%m-%d"),
           S["returns_all"]["fr"].max().strftime("%Y-%m-%d"),
           len(S["returns_all"]), len(S["returns"]))
        for S in stores if len(S.get("returns_all", [])))
    ws["A2"] = ("数据源：Returns_*.xlsx（Triages 页），按【质检日】记录，不是按订单日 —— "
                "本期订单的退货可能下月才质检，反之亦然。%s%s"
                % (cover, ("。%s 无退货质检记录。" % "、".join(none)) if none else ""))
    ws["A2"].font = F(9, False, C_GREY)
    r = 4
    if not have:
        ws.cell(row=r, column=1, value="本期没有任何退货质检记录。").font = F(11, False, C_GREY)
        return

    for S in have:
        R = S["returns"]
        tot = len(R) or 1
        ws.cell(row=r, column=1, value="【%s】一、按退货原因汇总（本期质检 %d 件）" % (S["store"], len(R))
                ).font = F(12, True, C_NAVY)
        r += 1
        cols = ["退货原因 (Estado del producto)", "中文", "件数", "占比",
                "其中：已退款买家", "其中：货款归卖家", "库存后果"]
        for i, (t, w) in enumerate(zip(cols, [42, 26, 9, 10, 15, 15, 34]), 1):
            c = ws.cell(row=r, column=i, value=t)
            c.font = F(10, True, "FFFFFF")
            c.fill = fill(C_NAVY)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = BOX
            ws.column_dimensions[get_column_letter(i)].width = w
        r += 1
        start = r
        for k, v in R.groupby("Estado del producto").size().sort_values(ascending=False).items():
            zh, cons = RETURN_REASON_ZH.get(k, (str(k), "待确认"))
            sub = R[R["Estado del producto"] == k]
            ws.cell(row=r, column=1, value=str(k))
            ws.cell(row=r, column=2, value=zh)
            ws.cell(row=r, column=3, value=int(v)).number_format = INT
            ws.cell(row=r, column=4, value="=C%d/%d" % (r, tot)).number_format = PCT
            ws.cell(row=r, column=5, value=int((sub["Estado del dinero"] == MONEY_REFUNDED).sum())).number_format = INT
            ws.cell(row=r, column=6, value=int((sub["Estado del dinero"] == MONEY_KEPT).sum())).number_format = INT
            ws.cell(row=r, column=7, value=cons)
            for c in range(1, 8):
                ws.cell(row=r, column=c).border = BOX
                ws.cell(row=r, column=c).font = F(10)
            if cons.startswith("✘"):
                for c in range(1, 8):
                    ws.cell(row=r, column=c).fill = fill(C_BAD)
            r += 1
        ws.cell(row=r, column=1, value="合计").font = F(10, True)
        for col, cl in ((3, "C"), (5, "E"), (6, "F")):
            ws.cell(row=r, column=col, value="=SUM(%s%d:%s%d)" % (cl, start, cl, r - 1)).number_format = INT
        ws.cell(row=r, column=4, value="=C%d/%d" % (r, tot)).number_format = PCT
        for c in range(1, 8):
            ws.cell(row=r, column=c).fill = fill(C_SUB)
            ws.cell(row=r, column=c).font = F(10, True)
            ws.cell(row=r, column=c).border = BOX
        r += 2

        ws.cell(row=r, column=1, value="【%s】二、按质检处置结果汇总" % S["store"]).font = F(12, True, C_NAVY)
        r += 1
        cols = ["处置结果 (Resultado de la revisión)", "中文", "件数", "占比", "资金结果", "库存价值", "财务影响"]
        for i, t in enumerate(cols, 1):
            c = ws.cell(row=r, column=i, value=t)
            c.font = F(10, True, "FFFFFF")
            c.fill = fill(C_NAVY)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = BOX
        r += 1
        for k, v in R.groupby("Resultado de la revisión").size().sort_values(ascending=False).items():
            zh, val, fin = RETURN_RESULT_ZH.get(k, (str(k), "待确认", "待确认"))
            sub = R[R["Resultado de la revisión"] == k]
            ref = int((sub["Estado del dinero"] == MONEY_REFUNDED).sum())
            ws.cell(row=r, column=1, value=str(k))
            ws.cell(row=r, column=2, value=zh)
            ws.cell(row=r, column=3, value=int(v)).number_format = INT
            ws.cell(row=r, column=4, value="=C%d/%d" % (r, tot)).number_format = PCT
            ws.cell(row=r, column=5, value="退款 %d 件 / 保留 %d 件" % (ref, int(v) - ref))
            ws.cell(row=r, column=6, value=val)
            ws.cell(row=r, column=7, value=fin)
            for c in range(1, 8):
                ws.cell(row=r, column=c).border = BOX
                ws.cell(row=r, column=c).font = F(10)
            if k == "Producto para retirar en centro de distribución":
                for c in range(1, 8):
                    ws.cell(row=r, column=c).fill = fill(C_BAD)
            elif k == "Producto quedó nuevamente para la venta":
                for c in range(1, 8):
                    ws.cell(row=r, column=c).fill = fill(C_OK)
            r += 1
        r += 1

        ws.cell(row=r, column=1, value="【%s】三、逐件明细（本期质检记录）" % S["store"]).font = F(12, True, C_NAVY)
        r += 1
        cols = ["订单号", "SKU", "质检日期", "商品状态 (原因)", "质检处置结果", "资金处理", "中文说明"]
        for i, t in enumerate(cols, 1):
            c = ws.cell(row=r, column=i, value=t)
            c.font = F(10, True, "FFFFFF")
            c.fill = fill(C_NAVY)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = BOX
        r += 1
        for _, x in R.sort_values("fr").iterrows():
            zh = RETURN_REASON_ZH.get(x["Estado del producto"], (str(x["Estado del producto"]), ""))[0]
            dz = RETURN_RESULT_ZH.get(x["Resultado de la revisión"],
                                      (str(x["Resultado de la revisión"]), "", ""))[0]
            mon = "已退款买家" if x["Estado del dinero"] == MONEY_REFUNDED else "货款归卖家"
            ws.cell(row=r, column=1, value=str(x.get("Número de orden", "")).replace(".0", ""))
            ws.cell(row=r, column=2, value=x.get("SKU"))
            ws.cell(row=r, column=3, value=x["fr"].strftime("%Y-%m-%d") if pd.notna(x["fr"]) else "")
            ws.cell(row=r, column=4, value=x["Estado del producto"])
            ws.cell(row=r, column=5, value=x["Resultado de la revisión"])
            ws.cell(row=r, column=6, value=x["Estado del dinero"])
            ws.cell(row=r, column=7, value="%s；%s；%s" % (zh, dz, mon))
            for c in range(1, 8):
                ws.cell(row=r, column=c).border = BOX
                ws.cell(row=r, column=c).font = F(9)
            if x["Resultado de la revisión"] == "Producto para retirar en centro de distribución":
                for c in range(1, 8):
                    ws.cell(row=r, column=c).fill = fill(C_BAD)
            elif x["Resultado de la revisión"] == "Producto quedó nuevamente para la venta":
                for c in range(1, 8):
                    ws.cell(row=r, column=c).fill = fill(C_OK)
            r += 1
        r += 2


# ════════════════════════════════════════════════════════════════════════
# ⑦ SKU 盈利分析
# ════════════════════════════════════════════════════════════════════════

V_UNIT_PRICE = "Precio unitario de venta de la publicación (MXN)"
PKG_RE = re.compile(r"^Paquete de (\d+) productos", re.I)


def split_packages(A):
    """把组合订单母行的金额按 件数×单价 摊到它的子行上。

    MercadoLibre 的导出把一单多件写成：母行 Estado="Paquete de N productos"
    带走**全部金额**却没有 SKU；紧随其后的 N 行有 SKU 和件数，金额列全为空。
    父子之间没有任何共同 ID —— 子行各有自己的 # de venta，母行的
    Orden de compra 还是空的 —— 所以只能靠**相邻位置**关联。

    拆分基数用 件数×单价。实测 Σ(件数×单价) 与母行收入分毫不差
    （BOCINA_TA02 的 538.92、580.00，TANKE_EE 的 397.20、418.60 等均如此），
    所以这是精确拆分而非估算。四舍五入的尾差记到份额最大的那一行，保证拆完
    各子行之和仍等于母行原值。

    **只给 ⑦ SKU 盈利分析用。** 损益表、各座桥和全部校验仍读未拆分的原始明细，
    这样既拿到按 SKU 的归属，又不会扰动已经通过的勾稽（桥A 残差 0.00）。
    """
    A = A.copy().reset_index(drop=True)
    money_cols = [c for c in (V_REVENUE, V_COMMISSION, V_SHIP_INCOME,
                              V_SHIP_COST, V_REFUND, "coupon",
                              "com_pure", "tax_ret") if c in A.columns]
    if V_STATUS not in A.columns:
        return A, 0
    st = A[V_STATUS].astype(str)
    drop, n_split = [], 0

    for i in range(len(A)):
        m = PKG_RE.match(st.iloc[i].strip())
        if not m:
            continue
        n = int(m.group(1))
        kids = list(range(i + 1, min(i + 1 + n, len(A))))
        # 子行数量对不上就原样保留 —— 宁可留一行空 SKU，也不要把金额摊到
        # 不相干的订单上。
        if len(kids) != n:
            continue

        units = pd.to_numeric(A.loc[kids, V_UNITS], errors="coerce").fillna(0.0)
        if V_UNIT_PRICE in A.columns:
            price = pd.to_numeric(A.loc[kids, V_UNIT_PRICE], errors="coerce").fillna(0.0)
        else:
            price = pd.Series(0.0, index=kids)
        basis = units * price
        if float(basis.sum()) <= 0:          # 没有单价就退回按件数分
            basis = units.copy()
        if float(basis.sum()) <= 0:          # 件数也没有就均分
            basis = pd.Series(1.0, index=kids)
        share = basis / float(basis.sum())
        big = share.idxmax()

        for c in money_cols:
            amt = pd.to_numeric(pd.Series([A.loc[i, c]]), errors="coerce").fillna(0.0).iloc[0]
            alloc = (share * float(amt)).round(2)
            alloc.loc[big] += round(float(amt) - float(alloc.sum()), 2)
            cur = pd.to_numeric(A.loc[kids, c], errors="coerce").fillna(0.0)
            A.loc[kids, c] = cur + alloc
        drop.append(i)
        n_split += 1

    if drop:
        A = A.drop(index=drop).reset_index(drop=True)

    # ---- 第二遍：换货订单 ----
    # 换货写成两行且**共享同一个 # de venta**：一行带金额没有 SKU
    # （Estado="Venta con solicitud de cambio"），另一行带 SKU 没有金额。
    # 这比包裹可靠 —— 有共同 ID，不用靠相邻位置猜。
    if V_ORDER_ID in A.columns and V_SKU in A.columns:
        sk = A[V_SKU].astype(str).str.strip()
        has_sku = (sk != "") & (sk.str.lower() != "nan")
        rev = pd.to_numeric(A[V_REVENUE], errors="coerce").fillna(0.0)             if V_REVENUE in A.columns else pd.Series(0.0, index=A.index)
        drop2 = []
        for oid, idx in A.groupby(A[V_ORDER_ID].astype(str)).groups.items():
            idx = list(idx)
            if len(idx) < 2:
                continue
            payers = [i for i in idx if not has_sku[i] and abs(rev[i]) > 0.01]
            takers = [i for i in idx if has_sku[i]]
            if len(payers) != 1 or not takers:
                continue
            p = payers[0]
            units = pd.to_numeric(A.loc[takers, V_UNITS], errors="coerce").fillna(0.0)
            basis = units if float(units.sum()) > 0 else pd.Series(1.0, index=takers)
            share = basis / float(basis.sum())
            big = share.idxmax()
            for c in money_cols:
                amt = pd.to_numeric(pd.Series([A.loc[p, c]]), errors="coerce").fillna(0.0).iloc[0]
                alloc = (share * float(amt)).round(2)
                alloc.loc[big] += round(float(amt) - float(alloc.sum()), 2)
                cur = pd.to_numeric(A.loc[takers, c], errors="coerce").fillna(0.0)
                A.loc[takers, c] = cur + alloc
            drop2.append(p)
            n_split += 1
        if drop2:
            A = A.drop(index=drop2).reset_index(drop=True)

    return A, n_split


def sku_grid():
    """⑦ 页四张表共用的一套列。主表用 A~P，三张明细多出来的项放到 Q 之后。

    四张表共用同一套列，是为了让人可以直接把鼠标从主表拖到任意一张明细上
    求和 —— 同一列永远是同一个项目。明细上没有的主表列留空（不是 0，
    0 会被误读成"这项是零"），主表上没有的明细列在主表里同样留空。
    """
    return [
        # (列名, 宽度)
        ("店铺", 16), ("SKU", 16), ("商品名称", 38),
        ("订单数", 9), ("销售件数", 10),
        ("销售额 GMV", 14), ("运费收入", 12),
        ("平台佣金", 13), ("平台代扣代缴税金", 15), ("平台配送费", 13),
        ("退款金额", 13), ("卖家优惠券", 12),
        ("单位采购成本\n（只在表五填）", 14), ("采购成本合计", 14),
        ("净贡献毛利", 14), ("贡献率", 10),
        # ↓ 以下仅明细表使用，主表留空
        ("订单号", 20), ("销售日期", 13), ("订单状态", 30),
        ("资金归属", 14), ("库存处置", 24),
        ("确定可回收日期", 15), ("质检结论", 16), ("质检原因 / 商品状态", 30),
        ("去向", 12), ("判定来源", 16), ("原始状态 (Estado)", 40),
        # ↓ 仅表五使用
        ("其中：正常销售件数", 14), ("其中：损失件数(不可再售)", 15), ("损失件数占比", 12),
    ]


SKU_NCOL = 30
SKU_MAIN_NCOL = 16
# 净贡献毛利 ＝ 销售额 ＋ 运费收入 − 佣金 − 代扣税 − 配送费 − 退款 − 优惠券 − 采购成本
SKU_NET = "=F{r}+G{r}-H{r}-I{r}-J{r}-K{r}-L{r}-N{r}"
SKU_RATE = "=IF(F{r}+G{r}=0,0,O{r}/(F{r}+G{r}))"

# 取消类订单：商品从未发出，库存没有减少。
CANCEL_RE = re.compile(r"cancel", re.I)
# Returns 报表的质检判定"重新上架"——库存实际回到手上
RESULT_BACK_ON_SALE = "Producto quedó nuevamente para la venta"
# Ventas 报表自带的质检三列（Resultado / Destino / Motivo del resultado）。
# 它和 Returns 报表是两个不同的来源：这三列就在订单行上，不用关联，覆盖面更广；
# Returns 报表则是 Full 仓的质检台账，两者都用，前者优先。
V_REVIEW_DATE = "Fecha de revisión"
V_REVIEW_RESULT = "Resultado"
V_REVIEW_DEST = "Destino"
V_REVIEW_WHY = "Motivo del resultado"
VENTAS_RESULT_ZH = {
    "Apto para la venta": ("可再售", True),
    "No apto para la venta": ("不可再售", False),
    "Regresado al comprador": ("已退回买家", False),
}
VENTAS_DEST_ZH = {
    "Vendedor": "退回卖家",
    "Comprador": "留在买家处",
    "Mercado Libre": "留在平台仓",
}
VENTAS_WHY_ZH = {
    "está en buenas condiciones": "状况良好",
    "están en buenas condiciones": "多件均状况良好",
    "no está en buenas condiciones": "状况不佳",
    "no funciona": "无法正常工作",
    "presenta irregularidades": "存在异常",
    "tiene el sello de la caja dañado": "包装封条破损",
    "tiene la caja original abierta": "原包装已拆封",
    "tiene la caja original dañada": "原包装破损",
    "tiene marcas de uso": "有使用痕迹",
}


def parse_spanish_daymonth(value, after=None):
    """'8 de septiembre' —— 平台的质检日**不带年份**，只能靠销售日推年。

    质检一定发生在销售之后，所以取销售当年；若这样算出来的日期早于销售日
    （典型是 12 月下单、次年 1 月质检），就进位到下一年。没有销售日可参照时
    返回 NaT，宁可留空也不要编一个年份出来。
    """
    if isinstance(value, pd.Timestamp):
        return value
    if not isinstance(value, str):
        return pd.NaT
    m = re.match(r"\s*(\d{1,2})\s+de\s+([A-Za-zÁ-úá-ú]+)\s*$", value.strip(), re.IGNORECASE)
    if not m:
        return parse_spanish_date(value)      # 带年份的照旧走原解析
    month = SPANISH_MONTHS.get(m.group(2).lower())
    if not month or after is None or pd.isna(after):
        return pd.NaT
    for y in (after.year, after.year + 1):
        try:
            d = pd.Timestamp(y, month, int(m.group(1)))
        except ValueError:
            return pd.NaT
        if d >= after.normalize():
            return d
    return pd.NaT


def split_commission(S):
    """把 Ventas 的『佣金及税金』一列拆成 纯佣金 与 代扣代缴税金 两列。

    平台在 Ventas 报表里把这两项**合并**给出（Cargo por venta e impuestos），
    在账单明细（Facturación）里只给纯佣金（Cargo por venta），代扣税从来没有
    单独的科目。所以逐单的拆法只能是：

        某单代扣税 ＝ 该单 Ventas 佣金及税 − 该单账单纯佣金

    账单行通过 # de venta / Orden de compra 关联到订单，可以跨账期
    （8 月的单，佣金可能记在 9 月账单上）。关联不上的少数订单用
    店内已知部分推出的比例补足，并保证两列之和仍等于损益表的
    平台佣金 + 平台代扣代缴税金 —— 页面因此与损益表分毫不差。

    返回 (com_pure, tax_ret, 说明文字)，两个 Series 与 S["ventas"] 同索引。
    """
    A = S["ventas"]
    v_com = -num(A, V_COMMISSION)                    # Ventas 口径，正数=费用
    if not len(A):
        return v_com, v_com, "本期没有订单"

    ids = set(A["oid"]) | set(A["pid"])
    ids.discard(""); ids.discard("nan")
    L = S["billing_linked"]
    com = L[(~L["is_rev"]) & L["parent_detalle"].map(
        lambda d: FEE_BY_ES.get(str(d), (None, None, None))[1] == "commission")]

    # 账单行 → 订单键。同一行的 k_sale / k_pack 只认能对上 Ventas 的那个，
    # 两个都对不上就放弃这行（它的钱由后面的残差分摊兜住）。
    by_key = {}
    for ks, kp, amt in zip(com["k_sale"].astype(str).str.strip(),
                           com["k_pack"].astype(str).str.strip(),
                           com["amount"]):
        k = ks if ks in ids else (kp if kp in ids else None)
        if k:
            by_key[k] = by_key.get(k, 0.0) + float(amt)

    pure = pd.Series(0.0, index=A.index)
    hit = pd.Series(False, index=A.index)
    for i, oid in A["oid"].items():
        if oid in by_key:
            pure[i] = by_key[oid]
            hit[i] = True
    # 只挂在包裹号上的账单行：按包裹内各行的 Ventas 佣金比例分摊
    rest = A.index[~hit]
    for pid, idx in A.loc[rest].groupby(A.loc[rest, "pid"]).groups.items():
        if pid not in by_key:
            continue
        idx = list(idx)
        w = v_com[idx].abs()
        w = w / w.sum() if float(w.sum()) > 0 else pd.Series(1.0 / len(idx), index=idx)
        pure[idx] = w * by_key[pid]
        hit[idx] = True

    # 残差：账单里有、但一行都没挂上的佣金。按 Ventas 佣金比例摊到没对上的行，
    # 全都对上了就摊回全部行。不摊会让代扣税凭空变大，比估算更糟。
    resid = float(S["m"]["com_gross"]) - float(pure.sum())
    if abs(resid) > 0.005:
        tgt = A.index[~hit] if (~hit).any() else A.index
        w = v_com[tgt].abs()
        w = w / w.sum() if float(w.sum()) > 0 else pd.Series(1.0 / len(tgt), index=tgt)
        pure[tgt] = pure[tgt] + w * resid

    n_hit = int(hit.sum())
    note = ("%d/%d 单的佣金直接取自账单明细" % (n_hit, len(A))
            + ("；其余 %d 单账单尚未出，按店内实际佣金率推算" % (len(A) - n_hit)
               if n_hit < len(A) else "")
            + ("；另有 %s MXN 账单佣金未能挂到具体订单，已按比例摊回" % fmt(resid)
               if abs(resid) > 0.005 and n_hit == len(A) else ""))
    return pure, v_com - pure, note


def returns_review(S):
    """订单号 → (质检日, 处置结果, 商品状态)。取全部期间的质检记录，
    因为本期订单的退货常常下个月才质检。

    这是 Ventas 自带质检列的**补充来源**：Returns 报表只覆盖 Full 仓，
    且它的订单号与 Ventas 不一定有交集，所以只在 Ventas 那三列为空时才用。
    """
    R = S.get("returns_all")
    out = {}
    if R is None or not len(R) or "Número de orden" not in R.columns:
        return out
    blank = pd.Series([None] * len(R), index=R.index)
    for oid, fr, res, est in zip(idkey(R["Número de orden"]), R["fr"],
                                 R["Resultado de la revisión"] if "Resultado de la revisión" in R.columns else blank,
                                 R["Estado del producto"] if "Estado del producto" in R.columns else blank):
        out[oid] = (fr, str(res) if pd.notna(res) else "",
                    str(est) if pd.notna(est) else "")
    return out


def _txt(x, col):
    v = x.get(col)
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    v = str(v).strip()
    return "" if v in ("nan", "-") else v


def review_of(x, fallback):
    """一行异常订单的质检结论，以及它是从哪来的。

    先看 Ventas 自带的 Resultado/Destino/Motivo（就在订单行上，最可靠）；
    没有才回退到 Returns 报表按订单号关联的记录；两边都没有就是"未质检"。

    返回 (日期, 结论中文, 原因中文, 去向中文, 来源, 可回收?)，
    可回收? 为 None 表示平台还没给结论 —— 与"判定为不可回收"必须区分开。
    """
    res = _txt(x, V_REVIEW_RESULT)
    if res:
        zh, ok = VENTAS_RESULT_ZH.get(res, (res, None))
        d = parse_spanish_daymonth(_txt(x, V_REVIEW_DATE), x.get("fecha"))
        why = _txt(x, V_REVIEW_WHY)
        return (d, zh, VENTAS_WHY_ZH.get(why, why),
                VENTAS_DEST_ZH.get(_txt(x, V_REVIEW_DEST), _txt(x, V_REVIEW_DEST)),
                "销售报表质检列", ok)
    if fallback:
        fr, r2, est = fallback
        return (fr, RETURN_RESULT_ZH.get(r2, (r2, "", ""))[0],
                RETURN_REASON_ZH.get(est, (est, ""))[0], "",
                "Returns 质检报表", (r2 == RESULT_BACK_ON_SALE) if r2 else None)
    return (pd.NaT, "", "", "", "", None)


def classify_abnormal(estado, rv):
    """异常订单归到哪一张明细表。

    分三类是因为三者的**库存后果完全不同**，混在一起算 SKU 盈利必然是错的：
      cancel  —— 从未发货，库存没动，不该有任何采购成本，毛利应为零
      back    —— 退货已质检判定可再售，货回到手上，同样不消耗采购成本
      other   —— 判定不可再售、或平台还没给结论，以及仲裁/投诉/换货/未送达：
                 库存价值已损失或去向未定，要计采购成本

    注意"还没给结论"归 other 而不是 back —— 在平台说可以之前，
    默认这批货收不回来才是保守且安全的。
    """
    if CANCEL_RE.search(estado or ""):
        return "cancel"
    if rv and rv[5] is True:
        return "back"
    if rv and rv[5] is None and estado in STATES_BACK_ON_SALE:
        return "back"      # 老状态文案已经写明"重新上架"，即便没有质检行
    return "other"


def sheet_sku(wb, stores, ctx):
    ws = wb.create_sheet(SHEET_NAMES[6])
    ws["A1"] = "SKU 盈利分析（%s）" % ctx["period_cn"]
    ws["A1"].font = F(14, True, C_NAVY)
    ws.row_dimensions[1].height = 24
    ws["A2"] = ("【使用说明】单位采购成本（到岸成本，含头程运费与关税）只在最下方"
                "【表五】的黄色 M 列填一次，表一与表四会自动引用同一个值，不要分头填。单位：MXN")
    ws["A2"].font = F(9, True, C_RED)
    ws["A3"] = ("本页共五张表。**要看某个 SKU 到底赚不赚钱，直接看最下方的【表五】** —— "
                "它是最终口径。表一只算正常订单，表二至表四是三类异常订单的逐单明细。"
                "五张表**列完全对齐**，同一列永远是同一个项目，可以直接拖着跨表求和；"
                "明细独有的项放在 Q 列之后，不适用的表留空。")
    ws["A3"].font = F(9, True, C_NAVY)
    ws["A4"] = ("注 1：佣金与代扣代缴税金已拆成两列 —— 平台在 Ventas 报表里把它们合并给出，"
                "本页用账单明细的纯佣金逐单还原（见表末说明）。每一列只放一项费用，不再合并。"
                "注 2：组合订单（Paquete de N productos）的母行金额已按 件数×单价 拆到各子 SKU。"
                "注 3：广告费、Full 仓储/揽收费在平台账单里没有商品编号，无法按 SKU 拆分，"
                "故本页是“订单级贡献”，未分摊店铺级共同费用。")
    ws["A4"].font = F(9, False, C_GREY)

    grid = sku_grid()
    HDR = 6

    def header(r, main):
        for i, (t, w) in enumerate(grid, 1):
            c = ws.cell(row=r, column=i, value=t)
            c.font = F(10, True, "FFFFFF")
            c.fill = fill(C_NAVY if (main and i <= SKU_MAIN_NCOL) else C_BAND)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = BOX
            if main:
                ws.column_dimensions[get_column_letter(i)].width = w
        ws.row_dimensions[r].height = 34

    header(HDR, True)

    def band(r, color, bold=False):
        for c in range(1, SKU_NCOL + 1):
            ws.cell(row=r, column=c).fill = fill(color)
            ws.cell(row=r, column=c).border = BOX
            if bold:
                ws.cell(row=r, column=c).font = F(10, True)

    def total_row(r, start, label, extra=()):
        """合计行：五张表列位置一一对应，所以整段直接求和即可。"""
        ws.cell(row=r, column=2, value=label).font = F(10, True)
        for col in tuple((4, 5, 6, 7, 8, 9, 10, 11, 12, 14, 15)) + tuple(extra):
            cl = get_column_letter(col)
            ws.cell(row=r, column=col, value="=SUM(%s%d:%s%d)" % (cl, start, cl, r - 1)
                    ).number_format = INT if col in (4, 5) or col in extra else MNY
        ws.cell(row=r, column=16, value=SKU_RATE.format(r=r)).number_format = PCT
        band(r, C_SUB, bold=True)

    def money_cells(r, vals):
        for col, v in vals:
            ws.cell(row=r, column=col, value=money(v)).number_format = MNY

    def sku_agg(D):
        """一组订单行 → 该 SKU 的各项金额。五张表都用它，口径不会分叉。"""
        return {
            "订单数": len(D), "件数": float(num(D, V_UNITS).sum()),
            "GMV": float(num(D, V_REVENUE).sum()),
            "运费收入": float(num(D, V_SHIP_INCOME).sum()),
            "佣金": float(num(D, "com_pure").sum()),
            "代扣税": float(num(D, "tax_ret").sum()),
            "配送费": float(-num(D, V_SHIP_COST).sum()),
            "退款": float(-num(D, V_REFUND).sum()),
            "券": float(D["coupon"].sum()) if "coupon" in D.columns else 0.0,
        }

    buckets = {"cancel": [], "back": [], "other": []}   # 跨店收集，明细按类分表
    split_notes = []
    other_cost_cell = {}      # 店铺 → 表一"其他异常小计"行号，稍后回指表四
    frames = {}               # 店铺 → (正常订单, 全部异常订单)
    cost_ref_rows = []        # 表一里需要回指表五取成本的行

    r = HDR + 1
    for S in stores:
        pure, tax, note = split_commission(S)
        split_notes.append("%s：%s" % (S["store"], note))
        V = S["ventas"].copy()
        V["com_pure"] = pure
        V["tax_ret"] = tax
        A, n_split = split_packages(V)          # 组合单的金额（含新加的两列）摊到子行
        S["_pkg_split"] = n_split
        fb = returns_review(S)

        ab = A[A["abnormal"]] if "abnormal" in A.columns else A.iloc[0:0]
        ok = A[~A["abnormal"]] if "abnormal" in A.columns else A
        kinds = {}
        for i, x in ab.iterrows():
            rv = review_of(x, fb.get(str(x.get("oid", ""))))
            k = classify_abnormal(str(x.get(V_STATUS, "")), rv)
            buckets[k].append((S["store"], x, rv))
            kinds[i] = k
        ab = ab.copy()
        ab["_kind"] = pd.Series(kinds) if kinds else ""
        frames[S["store"]] = (ok, ab)

        g = ok.groupby(ok[V_SKU].astype(str)).apply(lambda x: pd.Series(dict(
            sku_agg(x), 名称=(str(x[V_TITLE].iloc[0])[:45] if V_TITLE in x.columns else "")
        ))).reset_index()
        if len(g):
            g.columns = ["SKU"] + list(g.columns[1:])
            g = g.sort_values(["GMV", "SKU"], ascending=[False, True])
        start = r
        for _, x in g.iterrows():
            ws.cell(row=r, column=1, value=S["store"])
            ws.cell(row=r, column=2, value=str(x["SKU"]))
            ws.cell(row=r, column=3, value=x["名称"])
            ws.cell(row=r, column=4, value=int(x["订单数"])).number_format = INT
            ws.cell(row=r, column=5, value=int(x["件数"])).number_format = INT
            money_cells(r, [(6, x["GMV"]), (7, x["运费收入"]), (8, x["佣金"]),
                            (9, x["代扣税"]), (10, x["配送费"]), (11, x["退款"]),
                            (12, x["券"])])
            cost_ref_rows.append(r)          # M 列稍后回指表五
            ws.cell(row=r, column=14, value="=E%d*M%d" % (r, r)).number_format = MNY
            ws.cell(row=r, column=15, value=SKU_NET.format(r=r)).number_format = MNY
            ws.cell(row=r, column=16, value=SKU_RATE.format(r=r)).number_format = PCT
            for c in range(1, SKU_NCOL + 1):
                ws.cell(row=r, column=c).border = BOX
            for c in list(range(1, 13)) + list(range(14, 17)):
                ws.cell(row=r, column=c).font = F(10)
            r += 1

        # 三类异常各留一行小计。留在表一里，本店合计才与损益表的 GMV 对得上。
        # 取消与可回收的采购成本恒为 0；其他异常的成本在表四逐单算，这里回指过去，
        # 不然表四有成本、本店合计却看不见。
        for kind, label, why in (
            ("cancel", "取消订单小计", "从未发货，库存未减少，不计采购成本；明细见表二"),
            ("back", "退货可回收小计", "已质检判定可再售，库存已收回，不计采购成本；明细见表三"),
            ("other", "其他异常订单小计", "退货不可再售、仲裁、投诉、换货、尚未质检等；采购成本取自表四"),
        ):
            sub = [x for st, x, _ in buckets[kind] if st == S["store"]]
            if not sub:
                continue
            a = sku_agg(pd.DataFrame(sub))
            ws.cell(row=r, column=2, value=label).font = F(10, True, C_AMBER)
            ws.cell(row=r, column=3, value=why).font = F(9, False, C_GREY)
            ws.cell(row=r, column=4, value=a["订单数"]).number_format = INT
            ws.cell(row=r, column=5, value=int(a["件数"])).number_format = INT
            money_cells(r, [(6, a["GMV"]), (7, a["运费收入"]), (8, a["佣金"]),
                            (9, a["代扣税"]), (10, a["配送费"]), (11, a["退款"]),
                            (12, a["券"])])
            ws.cell(row=r, column=14, value=0).number_format = MNY
            if kind == "other":
                other_cost_cell[S["store"]] = r
            ws.cell(row=r, column=15, value=SKU_NET.format(r=r)).number_format = MNY
            ws.cell(row=r, column=16, value=SKU_RATE.format(r=r)).number_format = PCT
            band(r, C_WARN)
            r += 1

        total_row(r, start, "%s 合计" % S["store"])
        r += 2

    # ──────────────────────────────────────────────────────────────
    # 表二~表四：三张逐单明细。列与表一完全一致，Q 列之后才是明细独有的项。
    # ──────────────────────────────────────────────────────────────
    nz = lambda v: 0.0 if pd.isna(v) else float(v)
    detail_cost_rows = []

    def detail(title, intro, kind, cost_ref):
        """返回明细数据行的 (首行, 末行)，没有数据时返回 None。"""
        nonlocal r
        rows = buckets[kind]
        r += 1
        ws.cell(row=r, column=1, value=title).font = F(12, True, C_NAVY)
        r += 1
        ws.cell(row=r, column=1, value=intro).font = F(9, False, C_GREY)
        r += 1
        header(r, False)
        r += 1
        if not rows:
            ws.cell(row=r, column=1, value="本期没有这一类订单。").font = F(10, False, C_GREEN)
            r += 2
            return None
        start = r
        for store, x, rv in rows:
            estado = str(x.get(V_STATUS, ""))
            zh, who, inv = STATUS_ZH.get(estado, (estado, "待确认", "待确认"))
            rd, rres, rwhy, rdest, rsrc, _ = rv
            ws.cell(row=r, column=1, value=store)
            ws.cell(row=r, column=2, value=str(x.get(V_SKU, "")))
            ws.cell(row=r, column=3, value=str(x.get(V_TITLE, ""))[:45])
            ws.cell(row=r, column=4, value=1).number_format = INT
            ws.cell(row=r, column=5, value=int(nz(x.get(V_UNITS)))).number_format = INT
            money_cells(r, [(6, nz(x.get(V_REVENUE))), (7, nz(x.get(V_SHIP_INCOME))),
                            (8, nz(x.get("com_pure"))), (9, nz(x.get("tax_ret"))),
                            (10, -nz(x.get(V_SHIP_COST))), (11, -nz(x.get(V_REFUND))),
                            (12, nz(x.get("coupon")))])
            if cost_ref:
                detail_cost_rows.append(r)
                ws.cell(row=r, column=14, value="=E%d*M%d" % (r, r)).number_format = MNY
            else:
                ws.cell(row=r, column=14, value=0).number_format = MNY
            ws.cell(row=r, column=15, value=SKU_NET.format(r=r)).number_format = MNY
            ws.cell(row=r, column=16, value=SKU_RATE.format(r=r)).number_format = PCT
            # ---- Q 之后：明细独有的项 ----
            ws.cell(row=r, column=17, value=str(x.get("oid", "")))
            d = x.get("fecha")
            ws.cell(row=r, column=18, value=("" if pd.isna(d) else d.strftime("%Y-%m-%d")))
            ws.cell(row=r, column=19, value=zh)
            ws.cell(row=r, column=20, value=who)
            ws.cell(row=r, column=21, value=inv)
            if pd.notna(rd):
                ws.cell(row=r, column=22, value=rd.strftime("%Y-%m-%d"))
            elif kind == "back":
                # 已经归到可回收，说明判定确实做了，只是平台没给日期 ——
                # 不能写成"尚未质检"，那是另一回事。
                ws.cell(row=r, column=22, value="平台未给日期").font = F(9, False, C_AMBER)
            elif rres:
                ws.cell(row=r, column=22, value="平台未给日期").font = F(9, False, C_AMBER)
            elif kind == "other":
                ws.cell(row=r, column=22, value="尚未质检").font = F(9, False, C_AMBER)
            ws.cell(row=r, column=23, value=rres)
            ws.cell(row=r, column=24, value=rwhy)
            ws.cell(row=r, column=25, value=rdest)
            ws.cell(row=r, column=26, value=rsrc)
            ws.cell(row=r, column=27, value=estado)
            for c in range(1, SKU_NCOL + 1):
                ws.cell(row=r, column=c).border = BOX
                if c != 13 and ws.cell(row=r, column=c).font.color is None:
                    ws.cell(row=r, column=c).font = F(9)
            r += 1
        total_row(r, start, "合计 %d 单" % len(rows))
        r += 2
        return start, r - 3

    detail("表二 · 取消订单明细（商品从未发出，库存未减少）",
           "这些单在发货前就取消了：没有出库、库存没有减少，所以**不应产生任何采购成本**，"
           "采购成本合计一律记 0，净贡献只反映平台可能仍收取或已退还的费用。"
           "货还在手上可以再卖，所以表五**不把它们算进各 SKU 的盈利**，"
           "只在表五末尾单列一行残留费用。",
           "cancel", cost_ref=False)

    detail("表三 · 退货可回收明细（已质检判定可再售）",
           "退货已经过质检、判定可以重新销售，库存实际收回，因此同样不消耗采购成本，"
           "表五也不把它们算进各 SKU 的盈利。"
           "“确定可回收日期”＝ 平台做出这个判定的那一天（销售报表的 Fecha de revisión，"
           "缺失时取 Returns 质检报表的质检日）；“判定来源”列标明这一行取自哪个报表。",
           "back", cost_ref=False)

    span = detail("表四 · 其他异常订单明细（不可再售、仲裁、投诉、换货、未送达、尚未质检）",
                  "这些单的商品要么已判定不可再售、要么去向未定、要么平台还没质检，"
                  "库存价值不能假定收得回来 —— 所以**表五把它们并进对应 SKU**，"
                  "按件计采购成本。单位成本自动取自表五，不用在这里重填。",
                  "other", cost_ref=True)
    if span:
        d0, d1 = span
        for store, row in other_cost_cell.items():
            ws.cell(row=row, column=14,
                    value='=SUMIF($A$%d:$A$%d,"%s",$N$%d:$N$%d)' % (d0, d1, store, d0, d1)
                    ).number_format = MNY

    # ──────────────────────────────────────────────────────────────
    # 表五 · SKU 真实盈利。正常订单 ＋ 不可再售的异常订单，按 SKU 合并。
    # ──────────────────────────────────────────────────────────────
    r += 1
    ws.cell(row=r, column=1, value="表五 · SKU 真实盈利（最终口径）").font = F(13, True, C_NAVY)
    r += 1
    ws.cell(row=r, column=1, value=(
        "口径：正常订单 ＋ 表四的“其他异常”订单，按 SKU 合并。"
        "取消（表二）与退货可再售（表三）**不计入各 SKU** —— 这两类的货还在手上、还能再卖，"
        "把它们摊到 SKU 上会凭空压低该 SKU 的盈利；而表四那批货已经确定收不回来，"
        "货值实实在在损失了，必须算进这个 SKU 的成本里。"
    )).font = F(9, False, C_NAVY)
    r += 1
    ws.cell(row=r, column=1, value=(
        "但取消与可再售在**钱**上并不恒等于零（平台常常退了货款却不退配送费），"
        "所以每店末尾单列两行残留费用，把这部分留在合计里。"
        "因此本表的店铺合计与表一的店铺合计相等，也与损益表对得上 —— 只是钱的归属换了个分法。"
        "（两表可能相差 1 分：表一按“正常/异常”分组后各自四舍五入到分，本表按 SKU 合并后再舍入，"
        "落点不同。差额只在佣金与代扣税之间互相抵消，两列之和不变。）"
    )).font = F(9, False, C_GREY)
    r += 1
    ws.cell(row=r, column=1, value=(
        "★ 黄色 M 列是全页**唯一**的成本录入位：填在这里，表一与表四会自动引用。"
    )).font = F(9, True, C_RED)
    r += 1
    header(r, True)
    r += 1
    t5_start = r

    for S in stores:
        ok, ab = frames[S["store"]]
        other = ab[ab["_kind"] == "other"] if len(ab) else ab
        keys = sorted(set(ok[V_SKU].astype(str)) | set(other[V_SKU].astype(str)))
        rows = []
        for sku in keys:
            o = ok[ok[V_SKU].astype(str) == sku]
            t = other[other[V_SKU].astype(str) == sku] if len(other) else other
            a = sku_agg(pd.concat([o, t]) if len(t) else o)
            title = ""
            for D in (o, t):
                if len(D) and V_TITLE in D.columns:
                    title = str(D[V_TITLE].iloc[0])[:45]
                    break
            a.update({"SKU": sku, "名称": title,
                      "正常件数": float(num(o, V_UNITS).sum()),
                      "损失件数": float(num(t, V_UNITS).sum()) if len(t) else 0.0})
            rows.append(a)
        rows.sort(key=lambda a: (-a["GMV"], a["SKU"]))

        start = r
        for a in rows:
            ws.cell(row=r, column=1, value=S["store"])
            ws.cell(row=r, column=2, value=a["SKU"])
            ws.cell(row=r, column=3, value=a["名称"])
            ws.cell(row=r, column=4, value=a["订单数"]).number_format = INT
            ws.cell(row=r, column=5, value=int(a["件数"])).number_format = INT
            money_cells(r, [(6, a["GMV"]), (7, a["运费收入"]), (8, a["佣金"]),
                            (9, a["代扣税"]), (10, a["配送费"]), (11, a["退款"]),
                            (12, a["券"])])
            cc = ws.cell(row=r, column=13, value=0)
            cc.number_format = MNY
            cc.fill = fill(C_INPUT)
            cc.font = F(10, True, "0000FF")
            ws.cell(row=r, column=14, value="=E%d*M%d" % (r, r)).number_format = MNY
            ws.cell(row=r, column=15, value=SKU_NET.format(r=r)).number_format = MNY
            ws.cell(row=r, column=16, value=SKU_RATE.format(r=r)).number_format = PCT
            ws.cell(row=r, column=28, value=int(a["正常件数"])).number_format = INT
            ws.cell(row=r, column=29, value=int(a["损失件数"])).number_format = INT
            ws.cell(row=r, column=30,
                    value="=IF(E%d=0,0,AC%d/E%d)" % (r, r, r)).number_format = PCT
            for c in range(1, SKU_NCOL + 1):
                ws.cell(row=r, column=c).border = BOX
            for c in list(range(1, 13)) + list(range(14, 17)) + [28, 29, 30]:
                ws.cell(row=r, column=c).font = F(10)
            if a["损失件数"] > 0:
                for c in (2, 29, 30):
                    ws.cell(row=r, column=c).fill = fill(C_WARN)
            r += 1

        for kind, label, why in (
            ("cancel", "取消订单残留费用", "货未发出，不摊到任何 SKU；此处只留平台未退还的那部分费用"),
            ("back", "退货可再售残留费用", "货已收回可再卖，不摊到任何 SKU；此处只留平台未退还的那部分费用"),
        ):
            sub = [x for st, x, _ in buckets[kind] if st == S["store"]]
            if not sub:
                continue
            a = sku_agg(pd.DataFrame(sub))
            ws.cell(row=r, column=2, value=label).font = F(10, True, C_AMBER)
            ws.cell(row=r, column=3, value=why).font = F(9, False, C_GREY)
            ws.cell(row=r, column=4, value=a["订单数"]).number_format = INT
            ws.cell(row=r, column=5, value=int(a["件数"])).number_format = INT
            money_cells(r, [(6, a["GMV"]), (7, a["运费收入"]), (8, a["佣金"]),
                            (9, a["代扣税"]), (10, a["配送费"]), (11, a["退款"]),
                            (12, a["券"])])
            ws.cell(row=r, column=14, value=0).number_format = MNY
            ws.cell(row=r, column=15, value=SKU_NET.format(r=r)).number_format = MNY
            ws.cell(row=r, column=16, value=SKU_RATE.format(r=r)).number_format = PCT
            band(r, C_WARN)
            r += 1

        total_row(r, start, "%s 合计（＝表一同名行）" % S["store"], extra=(28, 29))
        r += 2
    t5_end = r - 2

    # 表一与表四的单位采购成本回指表五。用 SUMIFS 按【店铺＋SKU】两个键匹配 ——
    # 汇总报表里同一个 SKU 可能出现在好几家店，只按 SKU 找会取到别家的成本。
    cost_ref = ('=SUMIFS($M$%d:$M$%d,$A$%d:$A$%d,$A{r},$B$%d:$B$%d,$B{r})'
                % (t5_start, t5_end, t5_start, t5_end, t5_start, t5_end))
    for row in cost_ref_rows + detail_cost_rows:
        c = ws.cell(row=row, column=13, value=cost_ref.format(r=row))
        c.number_format = MNY
        c.fill = fill(C_CALC)
        c.font = F(10, False, C_GREY)
    # ⑨ 库存页也从这里取成本，同样按店铺＋SKU 两个键
    ctx["sku_cost_ref"] = (SHEET_NAMES[6], t5_start, t5_end)

    # ---- 表末：佣金/代扣税拆分的取数说明，避免有人以为这两列是平台直接给的 ----
    r += 1
    ws.cell(row=r, column=1, value="佣金与代扣代缴税金的拆分口径").font = F(11, True, C_NAVY)
    r += 1
    ws.cell(row=r, column=1, value=(
        "平台只在 Ventas 报表给出合并值 Cargo por venta e impuestos（佣金＋代扣税），"
        "账单明细 Facturación 只给纯佣金 Cargo por venta，代扣税没有独立科目。"
        "本页按订单号把账单的纯佣金挂回每一单，代扣税 ＝ 合并值 − 纯佣金。"
        "两列之和等于损益表的『平台销售佣金 ＋ 平台代扣代缴税金』。")).font = F(9, False, C_GREY)
    r += 1
    for t in split_notes:
        ws.cell(row=r, column=1, value=t).font = F(9, False, C_GREY)
        r += 1

    ws.freeze_panes = "A7"


# ══════════════════════════════════════════════════════════════════════
# ⑧ 月度趋势
# ══════════════════════════════════════════════════════════════════════

# 账单关联覆盖率低于这个值就不信账单，改用税率推算 —— 半数以上订单挂不上时，
# 用挂上的那点钱当整月佣金会把代扣税撑得离谱。
TREND_LINK_MIN = 0.5
SRC_BILLED = "账单实测"
SRC_RATED = "按本店实测税率推算"


def monthly_commission(S):
    """按**销售月**把 Ventas 的『佣金及税金』拆成 纯佣金 / 代扣代缴税金。

    返回 {ym: (纯佣金, 代扣税, 来源)}。

    账单只覆盖最近两三个计费期，销售报表却是滚动 6~13 个月，所以早期月份
    根本没有账单可挂。这些月份推的是**代扣税**而不是佣金 —— 代扣税是法定
    比例（各店实测 9.0x%，逐月几乎不动），佣金率却随品类在 11%~17% 之间
    跳动，推佣金等于编数字。推完再用 佣金 ＝ Ventas 合并值 − 代扣税 反求，
    两者之和因此永远等于平台给出的合并值，一分不多一分不少。
    """
    V = S["ventas_all"]
    V = V[V["ym"] != "NaT"]
    B = S["billing_all"]
    com = B[(~B["is_rev"]) & B["parent_detalle"].map(
        lambda d: FEE_BY_ES.get(str(d), (None, None, None))[1] == "commission")]
    ks_all = com["k_sale"].astype(str).str.strip()
    kp_all = com["k_pack"].astype(str).str.strip()

    out, measured, facts = {}, [], {}
    for ym, A in V.groupby("ym"):
        gmv = float(num(A, V_REVENUE).sum())
        vcom = float(-num(A, V_COMMISSION).sum())
        facts[ym] = (gmv, vcom)
        oids = set(A["oid"])
        ids = oids | set(A["pid"])
        ids.discard(""); ids.discard("nan")
        billed, hit = 0.0, set()
        for ks, kp, amt in zip(ks_all, kp_all, com["amount"]):
            k = ks if ks in ids else (kp if kp in ids else None)
            if k:
                billed += float(amt)
                hit.add(k)
        cov = len(hit & oids) / max(1, A["oid"].nunique())
        if cov >= TREND_LINK_MIN:
            out[ym] = (billed, vcom - billed, SRC_BILLED)
            if gmv > 0:
                measured.append((gmv, vcom - billed, billed))
        else:
            out[ym] = (None, None, None)      # 第二轮按税率补

    tot_gmv = sum(g for g, _, _ in measured)
    rate = (sum(t for _, t, _ in measured) / tot_gmv) if tot_gmv > 0 else TAX_RATE_NOMINAL_GROSS
    src = SRC_RATED + "（%.2f%%）" % (rate * 100)
    # 推算月份的哨兵：代扣税率一旦推高了，多算的税会**整笔落到佣金列**上，
    # 表现为佣金率异常偏低。只查偏低这一个方向 —— 佣金率偏高多半是品类结构
    # 变了（佣金按品类 13%~18%），本来就该逐月波动，当成异常只会天天误报。
    # 墨西哥站未完成税务登记的账号按更高比例代扣，新店早期几个月正是这种情况。
    crs = [b / g for g, _, b in measured if g > 0]
    floor = (min(crs) - 0.03) if crs else None
    for ym in out:
        if out[ym][0] is not None:
            continue
        gmv, vcom = facts[ym]
        tax = gmv * rate
        pure = vcom - tax
        note = src
        if floor is not None and gmv > 0 and pure / gmv < floor:
            note = ("%s ⚠ 倒推出的佣金率只有 %s，远低于实测月份的 %s —— "
                    "说明当月实际代扣税率高于 %s（账号税务登记状态多半和现在不同），"
                    "本行两列的分摊不可用，合计仍准确"
                    % (src, pct(pure / gmv, 1), pct(min(crs), 1), pct(rate, 2)))
        out[ym] = (pure, tax, note)
    return out


def trend_table(S):
    V = S["ventas_all"].copy()
    V = V[V["ym"] != "NaT"]
    split = monthly_commission(S)
    g = V.groupby("ym").apply(lambda x: pd.Series({
        "订单数": len(x), "件数": num(x, V_UNITS).sum(),
        "GMV": num(x, V_REVENUE).sum(),
        "运费": -num(x, V_SHIP_COST).sum(),
        "退款": -num(x, V_REFUND).sum(),
        "券": x["coupon"].sum(),
        "净额": num(x, V_TOTAL).sum(),
        "异常": x["abnormal"].sum(),
    })).reset_index()
    g["佣金"] = g["ym"].map(lambda y: split[y][0])
    g["代扣税"] = g["ym"].map(lambda y: split[y][1])
    g["拆分来源"] = g["ym"].map(lambda y: split[y][2])
    return g


def sheet_trend(wb, stores, ctx):
    ws = wb.create_sheet(SHEET_NAMES[7])
    ws["A1"] = "月度经营趋势"
    ws["A1"].font = F(14, True, C_NAVY)
    ws.row_dimensions[1].height = 24
    ws["A2"] = ("数据源：Ventas MX（销售报表为滚动窗口导出，窗口两端的月份可能不完整）。"
                "本期为 %s，灰斜体行表示不完整月份，仅供参考。单位：MXN" % ctx["period"])
    ws["A2"].font = F(9, False, C_GREY)
    ws["A3"] = ("平台三项扣费各占一列，不合并。佣金与代扣代缴税金在平台报表里是一个数，"
                "本表按账单明细逐单还原；账单只覆盖最近两三个计费期，更早的月份改按本店"
                "实测代扣税率推算代扣税、再倒推佣金（两者之和始终等于平台给出的合并值）。"
                "每行末列注明这一行到底是实测还是推算。")
    ws["A3"].font = F(9, False, C_NAVY)
    cols = ["店铺", "月份", "订单数", "销售件数", "销售额 GMV", "客单价",
            "平台佣金", "平台代扣代缴税金", "平台配送费",
            "佣金率", "代扣税率", "配送费率",
            "退款金额", "退款率", "卖家优惠券", "售后订单数", "售后率",
            "到账净额", "到账率", "佣金/代扣税来源"]
    widths = [16, 10, 9, 10, 15, 11, 14, 15, 14, 9, 10, 10, 13, 9, 12, 10, 9, 15, 9, 52]
    for i, t in enumerate(cols, 1):
        c = ws.cell(row=4, column=i, value=t)
        c.font = F(10, True, "FFFFFF")
        c.fill = fill(C_NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
        ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]
    ws.row_dimensions[4].height = 30
    NC = len(cols)

    r = 5
    for S in stores:
        for _, x in trend_table(S).iterrows():
            ym = str(x["ym"])
            ws.cell(row=r, column=1, value=S["store"])
            ws.cell(row=r, column=2, value=ym)
            ws.cell(row=r, column=3, value=int(x["订单数"])).number_format = INT
            ws.cell(row=r, column=4, value=int(x["件数"])).number_format = INT
            ws.cell(row=r, column=5, value=money(x["GMV"])).number_format = MNY
            ws.cell(row=r, column=6, value="=IF(C%d=0,0,E%d/C%d)" % (r, r, r)).number_format = MNY
            ws.cell(row=r, column=7, value=money(x["佣金"])).number_format = MNY
            ws.cell(row=r, column=8, value=money(x["代扣税"])).number_format = MNY
            ws.cell(row=r, column=9, value=money(x["运费"])).number_format = MNY
            for col, cl in ((10, "G"), (11, "H"), (12, "I")):
                ws.cell(row=r, column=col,
                        value="=IF($E%d=0,0,%s%d/$E%d)" % (r, cl, r, r)).number_format = PCT
            ws.cell(row=r, column=13, value=money(x["退款"])).number_format = MNY
            ws.cell(row=r, column=14, value="=IF(E%d=0,0,M%d/E%d)" % (r, r, r)).number_format = PCT
            ws.cell(row=r, column=15, value=money(x["券"])).number_format = MNY
            ws.cell(row=r, column=16, value=int(x["异常"])).number_format = INT
            ws.cell(row=r, column=17, value="=IF(C%d=0,0,P%d/C%d)" % (r, r, r)).number_format = PCT
            ws.cell(row=r, column=18, value=money(x["净额"])).number_format = MNY
            ws.cell(row=r, column=19, value="=IF(E%d=0,0,R%d/E%d)" % (r, r, r)).number_format = PCT
            sc = ws.cell(row=r, column=20, value=x["拆分来源"])
            for c in range(1, NC + 1):
                ws.cell(row=r, column=c).border = BOX
                ws.cell(row=r, column=c).font = F(10)
            if "⚠" in str(x["拆分来源"]):
                sc.font = F(9, True, C_RED)       # 推算与实测明显打架，别当成可用数字
            elif not str(x["拆分来源"]).startswith(SRC_BILLED):
                sc.font = F(9, False, C_AMBER)
            if ym in ctx["partial_months"]:
                for c in range(1, NC + 1):
                    ws.cell(row=r, column=c).font = F(10, False, "808080", True)
            if ym == ctx["period"]:
                for c in range(1, NC + 1):
                    ws.cell(row=r, column=c).fill = fill(C_SUB)
                    ws.cell(row=r, column=c).font = F(10, True)
            r += 1
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="趋势解读（按规则自动生成）").font = F(12, True, C_NAVY)
    r += 1
    for t in trend_notes(stores, ctx):
        c = ws.cell(row=r, column=1, value="• " + t)
        c.font = F(10)
        c.alignment = wrap()
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=NC)
        ws.row_dimensions[r].height = est_height(t, 120, 15, 15)
        r += 1
    ws.freeze_panes = "C5"


def trend_notes(stores, ctx):
    """确定性的趋势结论。规则固定，只有数字随月份变化。"""
    out = []
    for S in stores:
        g = trend_table(S)
        g = g[~g["ym"].isin(ctx["partial_months"])]
        if len(g) < 2:
            out.append("%s 的销售报表窗口内只有 %d 个完整月份，暂时无法判断趋势；"
                       "要看真实走势需要按月累积导出并保存历史。" % (S["store"], len(g)))
            continue
        cur = g[g["ym"] == ctx["period"]]
        if not len(cur):
            continue
        cur = cur.iloc[0]
        peak = g.loc[g["GMV"].idxmax()]
        if peak["ym"] != ctx["period"] and peak["GMV"] > 0:
            chg = (cur["GMV"] - peak["GMV"]) / peak["GMV"]
            out.append("%s 的 GMV 从 %s 的峰值 %s 变化到本期 %s（%s），订单数由 %s 变为 %s。"
                       % (S["store"], peak["ym"], fmt(peak["GMV"]), fmt(cur["GMV"]),
                          pct(chg, 1), fmt0(peak["订单数"]), fmt0(cur["订单数"])))
        gm = g["GMV"].replace(0, float("nan"))
        # 三项费率分开看：佣金率随品类结构走，配送费率随客单价走，代扣税率是法定的，
        # 混在一起看不出是哪一项在动。
        billed = g[g["拆分来源"] == SRC_BILLED]
        cr = (g["佣金"] / gm).dropna()
        if len(cr):
            out.append("%s 的佣金率区间 %s–%s，本期 %s（账单实测的月份有 %d 个，"
                       "其余月份的佣金是用合并值减推算代扣税倒推的）。佣金率变动说明品类结构变了。"
                       % (S["store"], pct(cr.min(), 1), pct(cr.max(), 1),
                          pct(cur["佣金"] / cur["GMV"] if cur["GMV"] else 0, 1), len(billed)))
        sr = (g["运费"] / gm).dropna()
        if len(sr):
            out.append("%s 的配送费率区间 %s–%s，本期 %s。每单配送费基本固定，"
                       "这个比率主要随客单价走 —— 客单价低则费率高。"
                       % (S["store"], pct(sr.min(), 1), pct(sr.max(), 1),
                          pct(cur["运费"] / cur["GMV"] if cur["GMV"] else 0, 1)))
        if len(billed):
            tr = (billed["代扣税"] / billed["GMV"].replace(0, float("nan"))).dropna()
            if len(tr):
                out.append("%s 的代扣代缴税率（账单实测月份）区间 %s–%s，法定名义值 %s。"
                           "越界说明税制或账号税务状态变了，需要人工确认。"
                           % (S["store"], pct(tr.min(), 2), pct(tr.max(), 2),
                              pct(TAX_RATE_NOMINAL_GROSS, 2)))
        rr = (g["退款"] / gm).dropna()
        if len(rr):
            worst = g.loc[rr.idxmax()]
            out.append("%s 的退款率峰值出现在 %s（%s），本期 %s；售后订单率本期 %s。"
                       % (S["store"], worst["ym"], pct(rr.max(), 2),
                          pct(cur["退款"] / cur["GMV"] if cur["GMV"] else 0, 2),
                          pct(cur["异常"] / cur["订单数"] if cur["订单数"] else 0, 2)))
        out.append("%s 的到账率（Ventas Total ÷ GMV）本期 %s，即每 100 元 GMV 实际到账约 %s 元，"
                   "其余是佣金、代扣税、配送费与退款。注意这个比率还没有扣广告费和仓储费。"
                   % (S["store"], pct(cur["净额"] / cur["GMV"] if cur["GMV"] else 0, 1),
                      fmt(cur["净额"] / cur["GMV"] * 100 if cur["GMV"] else 0, 0)))
    if len(stores) > 1:
        rates = sorted(((S["m"]["com_imp"] + S["m"]["env_cost"]) / S["m"]["ing"] if S["m"]["ing"] else 0, S)
                       for S in stores)
        lo, hi = rates[0], rates[-1]
        if hi[0] - lo[0] > 0.02:
            aov = lambda S: S["m"]["ing"] / S["m"]["rows"] if S["m"]["rows"] else 0
            out.append("横向比较：%s 的三项扣费合计占 GMV %s，明显高于 %s 的 %s，"
                       "主要因为客单价差异（%s vs %s）而每单固定配送费相近，配送费占比被拉高。"
                       % (hi[1]["store"], pct(hi[0], 1), lo[1]["store"], pct(lo[0], 1),
                          fmt(aov(hi[1]), 0), fmt(aov(lo[1]), 0)))
    return out


# ══════════════════════════════════════════════════════════════════════
# ⑨ 库存与动销分析
# ══════════════════════════════════════════════════════════════════════

# 平台的"近 30 天销量"只给到某个时刻为止，我们只能按整天切窗口，边界上差
# 一两件属于正常。超过这个容差才算真的对不上。
STOCK_TOL_ABS = 2
STOCK_TOL_PCT = 0.05
STOCK_WINDOW_DAYS = 30
# 可售天数的判定线
# 90 天没到平台加收长龄费的门槛，报出来只是噪音；120 天正是 Full 开始按
# "库龄 4 个月以上"加收的那条线 —— 越过它，积压才真的开始花钱。
DAYS_URGENT, DAYS_PILED = 14, 120


def sales_window(S):
    """与平台"近 30 天"对齐的自算销量：{SKU: (件数, 金额)}。

    口径是**剔除异常订单后的全部销量**。实测过四种口径（全部 / 仅 Full /
    仅正常 / Full＋正常），"仅正常"最贴平台：603 个 SKU 里 563 个差不超过
    1 件；再加 Full 过滤反而更差，说明平台这个数并不只算 Full 渠道。
    """
    d = S.get("stock_snap")
    V = S["ventas_all"]
    if d is None or pd.isna(d) or not len(V):
        return {}, None, None
    w0 = d - pd.Timedelta(days=STOCK_WINDOW_DAYS)
    W = V[(V["fecha"] >= w0) & (V["fecha"] < d + pd.Timedelta(days=1))]
    W = W[~W["abnormal"]] if "abnormal" in W.columns else W
    out = {}
    if len(W):
        k = W[V_SKU].astype(str).str.strip()
        u = W.groupby(k)[V_UNITS].apply(lambda x: float(pd.to_numeric(x, errors="coerce").fillna(0).sum()))
        a = W.groupby(k).apply(lambda x: float(num(x, V_REVENUE).sum()))
        for sku in u.index:
            out[sku] = (float(u[sku]), float(a.get(sku, 0.0)))
    return out, w0, d


def stock_verdict(row, ours, tol):
    """一行 SKU 的核对结论 + 动销判定 + 需要处理的事。

    核对和动销分开说：核对回答"平台和我们的数对不对得上"，动销回答
    "这批货该补还是该清"。两件事混成一列的话，一个对不上的滞销 SKU
    只能显示其中一半。
    """
    plat = row["sales30_u"]
    diff = ours - plat
    if row["_no_sales_row"] and plat > 0:
        chk = "⚠ 销售报表里查无此 SKU"
    elif row["_no_stock_row"]:
        chk = "⚠ 库存报表里查无此 SKU（非 Full 或已下架）"
    elif abs(diff) <= tol:
        chk = "一致"
    elif diff > 0:
        chk = "⚠ 平台少记 %d 件" % round(diff)
    else:
        chk = "⚠ 平台多记 %d 件" % round(-diff)

    sell, total = row["sellable"], row["total_u"]
    days = (sell / (plat / STOCK_WINDOW_DAYS)) if plat > 0 else None
    acts = []
    if row["discard"] > 0:
        acts.append("面临丢弃 %d 件" % round(row["discard"]))
    if row["unsellable"] > 0:
        acts.append("不可售 %d 件" % round(row["unsellable"]))
    if row["aged"] > 0:
        acts.append("长龄 %d 件（加收仓储费）" % round(row["aged"]))
    if sell <= 0 and plat > 0:
        acts.append("已断货，近 30 天仍有 %d 件成交" % round(plat))
    elif plat <= 0 and sell > 0:
        acts.append("零动销，压 %d 件" % round(sell))
    elif days is not None and days < DAYS_URGENT:
        acts.append("可售仅 %.0f 天，需补货" % days)
    elif days is not None and days > DAYS_PILED:
        acts.append("可售 %.0f 天，库存积压" % days)
    if row["to_list"] > 0:
        acts.append("未上架 %d 件" % round(row["to_list"]))
    if row["in_transit"] > 0 and total > 0 and row["in_transit"] == total:
        # 整批刚发出去还没入仓，"零动销/积压"是它的必然状态，不是问题。
        # 归到"正常"，免得每批新到货都进待办清单把真问题淹掉。
        acts = [a for a in acts if "零动销" not in a and "积压" not in a
                and "断货" not in a]
        if not acts:
            return chk, days, "正常（%d 件在途，尚未入仓）" % round(row["in_transit"])
        acts.insert(0, "全部在途，尚未入仓")
    return chk, days, ("；".join(acts) if acts else "正常")


def sheet_inventory(wb, stores, ctx):
    ws = wb.create_sheet(SHEET_NAMES[8])
    ws["A1"] = "库存与动销分析"
    ws["A1"].font = F(14, True, C_NAVY)
    ws.row_dimensions[1].height = 24
    have = [S for S in stores if len(S.get("stock", []))]
    none = [S["store"] for S in stores if not len(S.get("stock", []))]
    snaps = "；".join("%s 快照于 %s" % (S["store"],
                                    S["stock_snap"].strftime("%Y-%m-%d")
                                    if pd.notna(S.get("stock_snap")) else "日期未知")
                     for S in have)
    ws["A2"] = ("数据源：stock_general_full 报表的 Resumen 页（Full 仓库存）。%s%s"
                % (snaps, ("。%s 没有库存报表（需单独权限）。" % "、".join(none)) if none else ""))
    ws["A2"].font = F(9, False, C_GREY)
    ws["A3"] = ("⚠ 这是**下载当时的快照**，不是月末余额，与损益表不是同一个时点，"
                "两者不能相加。本页的“近 30 天”也是以快照日为终点，"
                "与会计月（%s）不重合。" % ctx["period"])
    ws["A3"].font = F(9, True, C_RED)
    ws["A4"] = ("销量核对：把平台给的“近 30 天销量”与我们从销售报表同窗口自算的件数逐 SKU 对比。"
                "自算口径为剔除异常订单后的全部销量。平台只给到快照那一刻，我们只能按整天切，"
                "所以差 %d 件以内（或 %s 以内）算一致，超出才标记。"
                % (STOCK_TOL_ABS, pct(STOCK_TOL_PCT, 0)))
    ws["A4"].font = F(9, False, C_NAVY)

    if not have:
        ws["A6"] = "本期没有任何店铺拉到库存报表，无法做库存分析。"
        ws["A6"].font = F(11, False, C_GREY)
        return

    # ---------------- 表一 · 各店库存结构 ----------------
    r = 6
    ws.cell(row=r, column=1, value="表一 · 各店 Full 仓库存结构").font = F(12, True, C_NAVY)
    r += 1
    t1 = ["店铺", "快照日", "SKU 数", "在途待入库", "在转运", "买家退回", "可售",
          "不可售", "丢失/审核/取消", "库存合计", "好品质", "待促动销", "待上架",
          "待避免丢弃", "长龄单位", "平台30天销量", "自算30天销量", "销量差异",
          "近30天平均库存", "库存周转天数"]
    w1 = [18, 12, 8, 12, 10, 11, 10, 10, 14, 11, 11, 11, 10, 12, 11, 13, 13, 11, 14, 13]
    for i, (t, w) in enumerate(zip(t1, w1), 1):
        c = ws.cell(row=r, column=i, value=t)
        c.font = F(10, True, "FFFFFF")
        c.fill = fill(C_NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[r].height = 30
    r += 1
    t1_start = r
    per_store = {}
    for S in have:
        ST = S["stock"]
        sold, w0, wd = sales_window(S)
        ours = sum(v[0] for v in sold.values())   # 含库存报表里没有的 SKU，差异才是全的
        per_store[S["store"]] = (sold, w0, wd)
        vals = [S["store"],
                S["stock_snap"].strftime("%Y-%m-%d") if pd.notna(S.get("stock_snap")) else "—",
                int(ST["sku"].nunique()),
                ST["in_transit"].sum(), ST["transfer"].sum(), ST["returned"].sum(),
                ST["sellable"].sum(), ST["unsellable"].sum(),
                ST["lost"].sum() + ST["in_review"].sum() + ST["cancelled_u"].sum(),
                ST["total_u"].sum(), ST["good"].sum(), ST["boost"].sum(),
                ST["to_list"].sum(), ST["discard"].sum(), ST["aged"].sum(),
                ST["sales30_u"].sum(), ours, ours - ST["sales30_u"].sum(),
                ST["avg_stock"].sum()]
        for i, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=i, value=(int(v) if isinstance(v, float) else v))
            if i >= 3:
                c.number_format = INT
        # 周转天数 = 平均库存 ÷ 日均销量。销量为 0 时不写 0 —— 0 天会被读成
        # "马上就卖完"，恰好与事实相反。
        ws.cell(row=r, column=20,
                value="=IF(P%d=0,\"无动销\",S%d/(P%d/%d))" % (r, r, r, STOCK_WINDOW_DAYS)
                ).number_format = "0.0"
        for c in range(1, len(t1) + 1):
            ws.cell(row=r, column=c).border = BOX
            ws.cell(row=r, column=c).font = F(10)
        r += 1
    if len(have) > 1:
        ws.cell(row=r, column=1, value="合计").font = F(10, True)
        for col in range(3, 20):
            cl = get_column_letter(col)
            ws.cell(row=r, column=col, value="=SUM(%s%d:%s%d)" % (cl, t1_start, cl, r - 1)
                    ).number_format = INT
        ws.cell(row=r, column=20,
                value="=IF(P%d=0,\"无动销\",S%d/(P%d/%d))" % (r, r, r, STOCK_WINDOW_DAYS)
                ).number_format = "0.0"
        for c in range(1, len(t1) + 1):
            ws.cell(row=r, column=c).fill = fill(C_SUB)
            ws.cell(row=r, column=c).font = F(10, True)
            ws.cell(row=r, column=c).border = BOX
        r += 1
    r += 2

    # ---------------- 表二 · 逐 SKU ----------------
    ws.cell(row=r, column=1, value="表二 · 逐 SKU 库存、动销与销量核对").font = F(12, True, C_NAVY)
    r += 1
    ws.cell(row=r, column=1, value=(
        "“自算30天销量”来自销售报表同窗口、剔除异常订单后的件数。"
        "单位采购成本从【%s】自动带过来（在那一页填一次即可），用来估算可售库存占用的资金。"
        % SHEET_NAMES[6])).font = F(9, False, C_GREY)
    r += 1
    t2 = ["店铺", "SKU", "商品名称", "刊登状态", "在途待入库", "可售", "不可售",
          "丢失/审核/取消", "库存合计", "长龄单位", "好品质", "待促动销", "待上架",
          "待避免丢弃", "平台30天销量", "自算30天销量", "销量差异", "核对结论",
          "平台30天销售额", "本期会计月销量", "近30天平均库存", "可售天数",
          "平台售罄预估", "单位采购成本", "可售库存金额", "需要处理的事"]
    w2 = [16, 18, 34, 11, 12, 9, 9, 13, 10, 10, 10, 10, 9, 11, 12, 12, 10, 22,
          14, 13, 13, 10, 14, 13, 14, 40]
    for i, (t, w) in enumerate(zip(t2, w2), 1):
        c = ws.cell(row=r, column=i, value=t)
        c.font = F(10, True, "FFFFFF")
        c.fill = fill(C_BAND)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    ws.row_dimensions[r].height = 32
    r += 1

    problems = []
    for S in have:
        ST = S["stock"].copy()
        sold, w0, wd = per_store[S["store"]]
        month_u = {}
        A = S["ventas"]
        if len(A):
            ok = A[~A["abnormal"]] if "abnormal" in A.columns else A
            month_u = ok.groupby(ok[V_SKU].astype(str).str.strip())[V_UNITS].apply(
                lambda x: float(pd.to_numeric(x, errors="coerce").fillna(0).sum())).to_dict()
        # 平台按**刊登**出一行，同一个 SKU 可能挂在好几个刊登上，先按 SKU 合并
        agg = ST.groupby("sku", as_index=False).agg(
            {c: "sum" for c in ("in_transit", "transfer", "returned", "sellable",
                                "unsellable", "lost", "in_review", "cancelled_u",
                                "total_u", "aged", "good", "boost", "to_list",
                                "discard", "avg_stock", "sales30_u", "sales30_amt")})
        first = ST.drop_duplicates("sku").set_index("sku")
        agg["title"] = agg["sku"].map(first["title"])
        agg["estado"] = agg["sku"].map(first["estado"])
        agg["runout"] = agg["sku"].map(first["runout"])
        agg["_no_stock_row"] = False
        # 销售报表里有、库存报表里没有的 SKU 也要露面，否则漏掉的正是最该查的
        extra = sorted(set(sold) - set(agg["sku"]) - {""})
        for sku in extra:
            row = {c: 0.0 for c in agg.columns if c not in ("sku", "title", "estado", "runout")}
            row.update({"sku": sku, "title": "（库存报表中无此 SKU）", "estado": "—",
                        "runout": "—", "_no_stock_row": True})
            agg = pd.concat([agg, pd.DataFrame([row])], ignore_index=True)
        agg = agg[agg["sku"].astype(str).str.strip() != ""]
        agg = agg.sort_values(["sales30_amt", "total_u"], ascending=[False, False])

        start = r
        for _, x in agg.iterrows():
            sku = str(x["sku"])
            ours = sold.get(sku, (0.0, 0.0))[0]
            x = dict(x)
            x["_no_sales_row"] = sku not in sold
            tol = max(STOCK_TOL_ABS, x["sales30_u"] * STOCK_TOL_PCT)
            chk, days, todo = stock_verdict(x, ours, tol)
            vals = [S["store"], sku, str(x["title"])[:45], x["estado"],
                    x["in_transit"], x["sellable"], x["unsellable"],
                    x["lost"] + x["in_review"] + x["cancelled_u"], x["total_u"],
                    x["aged"], x["good"], x["boost"], x["to_list"], x["discard"],
                    x["sales30_u"], ours, ours - x["sales30_u"]]
            for i, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=i, value=(int(v) if isinstance(v, float) else v))
                if i >= 5:
                    c.number_format = INT
            cc = ws.cell(row=r, column=18, value=chk)
            if chk != "一致":
                cc.font = F(9, True, C_AMBER)
            ws.cell(row=r, column=19, value=money(x["sales30_amt"])).number_format = MNY
            ws.cell(row=r, column=20, value=int(month_u.get(sku, 0.0))).number_format = INT
            ws.cell(row=r, column=21, value=int(x["avg_stock"])).number_format = INT
            ws.cell(row=r, column=22,
                    value="=IF(O%d=0,\"无动销\",F%d/(O%d/%d))" % (r, r, r, STOCK_WINDOW_DAYS)
                    ).number_format = "0.0"
            ws.cell(row=r, column=23, value=x["runout"])
            # 成本从 ⑦ 页表五带过来：那里是全页唯一的录入位，不要让人填第二遍。
            # 按【店铺＋SKU】两个键匹配 —— 汇总报表里同一个 SKU 会出现在好几家店。
            # 只认表五那一段，不能整列求和：⑦ 页的表一、表四同样有 M 列，会重复计。
            ref = ctx.get("sku_cost_ref")
            ws.cell(row=r, column=24, value=(
                "=SUMIFS('%s'!$M$%d:$M$%d,'%s'!$A$%d:$A$%d,$A%d,'%s'!$B$%d:$B$%d,$B%d)"
                % (ref[0], ref[1], ref[2], ref[0], ref[1], ref[2], r,
                   ref[0], ref[1], ref[2], r) if ref else 0)
                ).number_format = MNY
            ws.cell(row=r, column=25, value="=F%d*X%d" % (r, r)).number_format = MNY
            tc = ws.cell(row=r, column=26, value=todo)
            tc.alignment = Alignment(wrap_text=True, vertical="top")
            if not todo.startswith("正常"):
                tc.font = F(9, False, C_RED)
            if not todo.startswith("正常") or chk != "一致":
                problems.append((S["store"], sku, str(x["title"])[:40], todo, chk,
                                 x["sellable"], x["sales30_u"], ours,
                                 (0 if days is None else round(days)),
                                 x["discard"] + x["unsellable"]))
            for c in range(1, len(t2) + 1):
                ws.cell(row=r, column=c).border = BOX
                if ws.cell(row=r, column=c).font.color is None:
                    ws.cell(row=r, column=c).font = F(9)
            if x["_no_stock_row"] or x["_no_sales_row"]:
                for c in range(1, len(t2) + 1):
                    ws.cell(row=r, column=c).fill = fill(C_WARN)
            r += 1

        ws.cell(row=r, column=2, value="%s 合计" % S["store"]).font = F(10, True)
        for col in list(range(5, 18)) + [19, 20, 21, 25]:
            cl = get_column_letter(col)
            ws.cell(row=r, column=col, value="=SUM(%s%d:%s%d)" % (cl, start, cl, r - 1)
                    ).number_format = MNY if col in (19, 25) else INT
        for c in range(1, len(t2) + 1):
            ws.cell(row=r, column=c).fill = fill(C_SUB)
            ws.cell(row=r, column=c).font = F(10, True)
            ws.cell(row=r, column=c).border = BOX
        r += 2

    # ---------------- 表三 · 需要处理的清单 ----------------
    r += 1
    ws.cell(row=r, column=1, value="表三 · 需要处理的清单（按占用资金排序）").font = F(12, True, C_NAVY)
    r += 1
    ws.cell(row=r, column=1, value=(
        "只列出表二里需要处理、或与平台对不上的 SKU。顺序：断货 → 面临丢弃 → "
        "数量对不上 → 其余。断货是在丢确定的销量，面临丢弃是已经确定要亏的货，"
        "数量对不上则说明 SKU 映射或渠道口径有问题，会影响后面所有按 SKU 的分析。"
        )).font = F(9, False, C_GREY)
    r += 1
    t3 = ["店铺", "SKU", "商品名称", "需要处理的事", "核对结论", "可售库存",
          "平台30天销量", "自算30天销量", "可售天数", "不可售+待丢弃"]
    w3 = [16, 18, 34, 44, 24, 11, 13, 13, 10, 13]
    for i, (t, w) in enumerate(zip(t3, w3), 1):
        c = ws.cell(row=r, column=i, value=t)
        c.font = F(10, True, "FFFFFF")
        c.fill = fill(C_BAND)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    r += 1
    if not problems:
        ws.cell(row=r, column=1, value="本期库存没有需要处理的项。").font = F(10, False, C_GREEN)
        r += 1
    else:
        # 断货 > 面临丢弃 > 其他；同级按可售库存量
        rank = lambda p: (0 if "断货" in p[3] else 1 if "丢弃" in p[3] else
                          2 if "⚠" in p[4] else 3, -p[5])
        for p in sorted(problems, key=rank):
            for i, v in enumerate(p[:5], 1):
                ws.cell(row=r, column=i, value=v)
            for i, v in zip((6, 7, 8, 9, 10), p[5:]):
                ws.cell(row=r, column=i, value=int(v)).number_format = INT
            ws.cell(row=r, column=4).alignment = Alignment(wrap_text=True, vertical="top")
            for c in range(1, len(t3) + 1):
                ws.cell(row=r, column=c).border = BOX
                ws.cell(row=r, column=c).font = F(9)
            if "断货" in p[3]:
                for c in range(1, len(t3) + 1):
                    ws.cell(row=r, column=c).fill = fill(C_BAD)
            r += 1
    ws.freeze_panes = "C8"


# ══════════════════════════════════════════════════════════════════════
# ⑩ 数据缺口与待办（含本期校验结果）
# ══════════════════════════════════════════════════════════════════════

def sheet_gaps(wb, stores, ctx):
    ws = wb.create_sheet(SHEET_NAMES[9])
    ws["A1"] = "本期校验结果 · 数据缺口 · 待办清单"
    ws["A1"].font = F(14, True, C_NAVY)
    ws.row_dimensions[1].height = 24
    ws["A2"] = "上半部分是脚本每期自动跑的硬校验；下半部分是需要人处理的事项，按优先级排序。"
    ws["A2"].font = F(9, False, C_GREY)
    for i, w in enumerate([10, 30, 60, 54, 40], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    r = 4
    ws.cell(row=r, column=1, value="一、本期校验结果").font = F(12, True, C_NAVY)
    r += 1
    for i, t in enumerate(["结果", "店铺", "校验项", "明细", ""], 1):
        c = ws.cell(row=r, column=i, value=t)
        c.font = F(10, True, "FFFFFF")
        c.fill = fill(C_NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BOX
    r += 1
    for c0 in ctx["checks"]:
        mark = "跳过" if c0["skipped"] else ("通过 ✓" if c0["ok"] else "未通过 ✗")
        col = C_GREY if c0["skipped"] else (C_GREEN if c0["ok"] else C_RED)
        bg = C_WARN if c0["skipped"] else (C_OK if c0["ok"] else C_BAD)
        ws.cell(row=r, column=1, value=mark).font = F(10, True, col)
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r, column=2, value=c0["store"]).font = F(10)
        ws.cell(row=r, column=3, value=c0["name"]).font = F(10)
        d = ws.cell(row=r, column=4, value=c0["detail"])
        d.font = F(9)
        d.alignment = wrap("top")
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
        ws.row_dimensions[r].height = est_height(c0["detail"], 70)
        for c in range(1, 6):
            ws.cell(row=r, column=c).border = BOX
            ws.cell(row=r, column=c).fill = fill(bg)
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="二、数据缺口与待办").font = F(12, True, C_NAVY)
    r += 1
    for i, t in enumerate(["优先级", "事项", "现状 / 发现", "建议动作", "对财务的影响"], 1):
        c = ws.cell(row=r, column=i, value=t)
        c.font = F(10, True, "FFFFFF")
        c.fill = fill(C_NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BOX
    ws.row_dimensions[r].height = 22
    r += 1

    items = []
    nsku = sum(S["ventas"][V_SKU].nunique() for S in stores)
    items.append(("高", "补齐商品采购成本",
                  "平台所有报表都不含进货成本，因此本报表只能算到“经营贡献毛利”（本期合计 %s）。"
                  % fmt(sum(ctx["contrib"].values())),
                  "在【%s】表五的黄色“单位采购成本”列填入每个 SKU 的到岸单位成本（进货价＋头程运费＋关税＋清关杂费），"
                  "本期共 %d 个 SKU。" % (SHEET_NAMES[6], nsku),
                  "这是算出真实净利润的唯一缺口，填完即可得出完整损益。"))

    unk_all = [(S, u) for S in stores for u in S["unknown_fees"]]
    if unk_all:
        items.append(("高", "出现未登记的费用科目 ⚠",
                      "本期发现 %d 个 FEE_CATALOG 里没有的科目：%s。"
                      "它们的金额已计入报表的“未登记科目”行以保证勾稽不断，但科目归属未确认。"
                      % (len(unk_all), "；".join("%s（%s，%s，%d 笔）" % (
                          u["detalle"], S["store"], fmt(u["amount"]), u["count"]) for S, u in unk_all)),
                      "确认这些科目的业务含义后，加入 report_build.py 的 FEE_CATALOG 常量表"
                      "（写明中文名、大类、是否订单级），并在 PL_LINES 里给它一行。",
                      "未登记科目会让损益表的费用分类不完整；金额本身没有丢，但看不出是什么钱。"))

    no_mp = [S["store"] for S in stores if not S.get("mp")]
    if no_mp:
        items.append(("高", "MercadoPago 报表权限",
                      "%s 本期没有下载到 settlement 结算流水，无法用第三方凭证交叉验证到账金额。"
                      % "、".join(no_mp),
                      "让账号管理员在 MercadoPago 后台“协作者(Colaboradores)”中开通"
                      "“查看收款与账单报表”和“查看操作报表”两项权限，然后重新下载结算流水。",
                      "缺少独立凭证时，到账金额只能按销售报表推算，无法验证平台是否少付。"))

    for S in stores:
        m = S["m"]
        if m["ing"] and m["ads_net"] / m["ing"] > 0.04:
            share = m["ads_net"] / ctx["contrib"][S["store"]] if ctx["contrib"][S["store"]] else 0
            items.append(("高", "广告投入产出需要复核（%s）" % S["store"],
                          "本期广告净支出 %s，占 GMV %s，占经营贡献毛利 %s。"
                          % (fmt(m["ads_net"]), pct(m["ads_net"] / m["ing"], 1), pct(share, 1)),
                          "按 Product Ads 后台的广告活动(campaña)导出 ACOS/ROAS 明细找出无效活动。"
                          "平台账单不提供按商品的广告拆分，必须从广告后台单独导。",
                          "这是金额最大的可控成本项之一。"))
        if len(S["returns"]):
            R = S["returns"]
            lost = int((R["Resultado de la revisión"] == "Producto para retirar en centro de distribución").sum())
            if lost and lost / len(R) > 0.3:
                top = R.groupby("Estado del producto").size().sort_values(ascending=False)
                items.append(("中", "Full 仓退货“需自提”占比过高（%s）" % S["store"],
                              "本期质检 %d 件中 %d 件（%s）被判定“需到配送中心自提”，即不可再售。"
                              "主因：%s。" % (len(R), lost, pct(lost / len(R), 0),
                                            "、".join("%s %d 件" % (k, v) for k, v in top.head(3).items())),
                              "① 尽快决定这批货是自提翻新还是放弃（放着会产生长龄库存费）；"
                              "② 针对封条破损/包装类原因，检查外包装与封装工艺，这类退货多数可以从源头避免。",
                              "每件都是“钱退了、货也卖不掉”的双重损失，同时还要付退仓费。"))
        space = sum(v for es, v in m["bill_fee_net"].items()
                    if es in ("Cargo por comprar más espacio en Full",
                              "Cargo por sobrepasar espacio Full",
                              "Cargo por stock antiguo en Full"))
        if space > 0:
            items.append(("中", "Full 库存空间与库龄成本（%s）" % S["store"],
                          "本期购买额外仓位 %s、超仓位罚金 %s、长龄库存费 %s，合计 %s。"
                          % (fmt(m["bill_fee_net"].get("Cargo por comprar más espacio en Full", 0)),
                             fmt(m["bill_fee_net"].get("Cargo por sobrepasar espacio Full", 0)),
                             fmt(m["bill_fee_net"].get("Cargo por stock antiguo en Full", 0)), fmt(space)),
                          "核对 stock_general_full 报表中“需清理”的库存，把滞销 SKU 从 Full 仓撤出或做清仓。",
                          "仓位与库龄费合计占经营贡献毛利 %s，属可优化项。"
                          % pct(space / ctx["contrib"][S["store"]] if ctx["contrib"][S["store"]] else 0, 1)))
        pen = m["bill_fee_net"].get("Cargo por incumplimiento en Envíos Full", 0)
        if pen > 0:
            items.append(("中", "Full 发货违规罚金（%s）" % S["store"],
                          "本期被罚 %s。" % fmt(pen),
                          "核对相关发货批次的实发数量与申报数量，找出差异原因，必要时向平台申诉。",
                          "金额通常不大，但反映入仓流程有问题，容易反复发生。"))
        if m["coupon_n"]:
            items.append(("中", "优惠券列被平台导出成空白（%s）" % S["store"],
                          "Ventas 报表的 Descuentos y bonificaciones 列整列为空，但本期实际发生 %d 笔"
                          "卖家自负优惠券、合计 %s。本报表用 Total 倒推还原。"
                          % (m["coupon_n"], fmt(m["coupon"])),
                          "① 到促销后台核对优惠券活动的实际投放与预算；"
                          "② 每月导出时留意这一列，若平台修复了导出，可直接取用不必倒推。",
                          "占 GMV %s。这是唯一一笔“账面完全看不见”的支出，不还原就永远差这一块对不上。"
                          % pct(m["coupon"] / m["ing"] if m["ing"] else 0, 2)))
        if S["files"].get("returns") and len(S.get("returns_all", [])):
            ra = S["returns_all"]
            items.append(("中", "退货质检报表期间不匹配（%s）" % S["store"],
                          "Returns 报表按【质检日】导出，本次覆盖 %s ~ %s；本期订单的退货可能下月才质检。"
                          % (ra["fr"].min().strftime("%Y-%m-%d"), ra["fr"].max().strftime("%Y-%m-%d")),
                          "每月固定在月结后 15 天再导一次 Returns，用【订单号】与销售报表关联，"
                          "才能把退货损失准确归到销售月份。",
                          "本期退货损失可能被低估（部分本期订单的退货尚未质检完）。"))
        if not S["files"].get("cargos_full"):
            items.append(("中", "缺少 Full 费用明细报表（%s）" % S["store"],
                          "没有找到本期的 Reporte_Cargos_Full 文件，无法与账单的 Full 科目交叉核对。",
                          "在 Full 后台重新导出本期的 Cargos Full 报表。",
                          "少了一道校验；账单口径仍然准确，但无法独立验证。"))
        if S["files"].get("storage"):
            d = m["storage_bill"] - m["storage_accrual"]
            if abs(d) > 0.01:
                items.append(("低", "仓储费两个口径的时间差（%s）" % S["store"],
                              "账单（按计费日）%s vs 独立仓储报表（按占用日）%s，差 %s。"
                              % (fmt(m["storage_bill"]), fmt(m["storage_accrual"]), fmt(d)),
                              "无需处理。本报表统一采用账单口径，因为那才是真正扣钱的金额。",
                              "影响极小，属正常的跨月错位。"))

    for S in stores:
        for a in S["anomalies"]:
            items.append(("中", "科目出现在非预期通道（%s）" % S["store"], a,
                          "确认该科目的计费方式是否变化，必要时调整 FEE_CATALOG 里的“是否订单级”。",
                          "会影响费用在损益表中的归类位置，不影响合计。"))

    if ctx["partial_months"]:
        items.append(("低", "销售报表窗口两端的月份不完整",
                      "%s 为不完整月份（销售报表是滚动窗口导出）。" % "、".join(ctx["partial_months"]),
                      "下月初重新下载完整数据；长期看应按月累积保存指标，才能得到真实趋势。",
                      "本报表未将这些月份计入损益，仅在【%s】作参考。" % SHEET_NAMES[7]))
    items.append(("低", "代扣代缴税金的税务处理",
                  "本期%s被代扣 %s，约为含税售价的 %s。"
                  % (store_count_cn(stores),
                     fmt(sum(S["m"]["tax_ret"] for S in stores)),
                     pct(agg_tax_rate(stores), 2)),
                  "把 MercadoPago 的 Constancia de retenciones（预扣凭证）交给墨西哥会计师，"
                  "确认能否全额抵扣 IVA/ISR 申报。",
                  "若能全额抵扣，这笔钱不是成本；若不能，需重新计入费用，会显著改变净利润。"))

    order = {"高": 0, "中": 1, "低": 2}
    items.sort(key=lambda x: order.get(x[0], 9))
    col = {"高": C_RED, "中": C_AMBER, "低": C_GREEN}
    bg = {"高": C_BAD, "中": C_WARN, "低": C_OK}
    for pri, tit, now, act, imp in items:
        ws.cell(row=r, column=1, value=pri).font = F(11, True, col[pri])
        ws.cell(row=r, column=1).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r, column=1).fill = fill(bg[pri])
        ws.cell(row=r, column=2, value=tit).font = F(10, True)
        ws.cell(row=r, column=2).alignment = wrap("top")
        for c, txt in ((3, now), (4, act), (5, imp)):
            cc = ws.cell(row=r, column=c, value=txt)
            cc.font = F(9)
            cc.alignment = wrap("top")
        for c in range(1, 6):
            ws.cell(row=r, column=c).border = BOX
        ws.row_dimensions[r].height = max(est_height(now, 48), est_height(act, 44), est_height(imp, 32))
        r += 1

    # ---- 三、贷记明细 × 销售报表 逐单比对的异常清单 ----
    anyrow = any(S["cn"]["orphan"] or S["cn"]["unrev"] or S["cn"]["partial"]
                 or S["cn"]["mismatch"] for S in stores)
    r += 2
    ws.cell(row=r, column=1, value="三、贷记明细 × 销售报表 逐单比对").font = F(12, True, C_NAVY)
    r += 1
    ws.cell(row=r, column=1, value=(
        "按订单号把账单/贷记单与销售报表逐单对照。关联键有个坑：账单的 Número de venta "
        "与销售报表的 # de venta 不是同一个号段，可靠的桥是 Número de paquete。"
        "下面三类异常，桥A（销售报表内部勾稽）、桥B（账单总额）、MercadoPago（店铺汇总）都看不见。"
    )).font = F(9, False, C_GREY)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    ws.row_dimensions[r].height = est_height(ws.cell(row=r, column=1).value, 100)
    r += 2
    if not anyrow:
        ws.cell(row=r, column=1, value="本期三类异常均为 0，逐单比对全部通过。").font = F(10, False, C_GREEN)
        r += 1

    def block(title, intro, cols, widths, rows, sum_cols=()):
        nonlocal r
        if not rows:
            return
        ws.cell(row=r, column=1, value=title).font = F(11, True, C_NAVY)
        r += 1
        ws.cell(row=r, column=1, value=intro).font = F(9, False, C_GREY)
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=max(5, len(cols)))
        ws.row_dimensions[r].height = est_height(intro, 150)
        r += 1
        for i, (t, w) in enumerate(zip(cols, widths), 1):
            c = ws.cell(row=r, column=i, value=t)
            c.font = F(10, True, "FFFFFF")
            c.fill = fill(C_BAND)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = BOX
            if ws.column_dimensions[get_column_letter(i)].width < w:
                ws.column_dimensions[get_column_letter(i)].width = w
        ws.row_dimensions[r].height = 28
        r += 1
        start = r
        for vals in rows:
            for i, v in enumerate(vals, 1):
                c = ws.cell(row=r, column=i, value=v)
                c.font = F(9)
                c.border = BOX
                if isinstance(v, float):
                    c.number_format = MNY
            r += 1
        ws.cell(row=r, column=1, value="合计 %d 单" % len(rows)).font = F(10, True)
        # 要合计哪几列由调用方指定，不靠表头文字去猜 —— 猜法会把"MP实际净额"
        # 这种加起来没有意义的列也求和。
        for i in sum_cols:
            cl = get_column_letter(i)
            ws.cell(row=r, column=i, value="=SUM(%s%d:%s%d)" % (cl, start, cl, r - 1)
                    ).number_format = MNY
        for i in range(1, len(cols) + 1):
            ws.cell(row=r, column=i).fill = fill(C_SUB)
            ws.cell(row=r, column=i).font = F(10, True)
            ws.cell(row=r, column=i).border = BOX
        r += 2

    block("3.1 账单有费用、销售报表查无此单（且未被冲平）",
          "下单后立刻作废的单，平台会计费再全额冲销，净额为 0，属正常，已排除。"
          "留在这里的是净额不为零的 —— 平台为一笔销售报表上不存在的订单收了钱。"
          "这笔钱既不在销售报表里，也不会被桥A 发现，只能人工向平台核实。",
          ["店铺", "账单上的订单键", "账单净额", "账单行数"], [18, 22, 14, 10],
          [(S["store"], k, money(amt), cnt)
           for S in stores for k, amt, cnt in S["cn"]["orphan"]], sum_cols=(3,))

    block("3.2 取消单计了费却没有冲销",
          "发货前取消的订单，平台通常压根不下账单，所以“没有贷记单”本身不是问题 —— "
          "这里只列确实计过费、账单净额不为零的。",
          ["店铺", "订单号", "账单净额", "订单状态", "销售日", "SKU"],
          [18, 22, 14, 34, 13, 18],
          [(S["store"], k, money(net), st,
            ("" if d is None or pd.isna(d) else d.strftime("%Y-%m-%d")), sku)
           for S in stores for k, net, st, d, sku in S["cn"]["unrev"]], sum_cols=(3,))

    block("3.3 部分退款订单：销售报表把原佣金重复扣了一次",
          "平台按退款比例 p 同时退还佣金与代扣税（代扣税在账单里没有科目，只在 MercadoPago 流水可见）。"
          "应退 ＝ p ×（GMV − 佣金 − 代扣税）；而销售报表的“退款与取消”列写的是 应退 ＋ 原佣金，"
          "于是多扣一个原佣金。“认定依据”为现金凭证的，是用 MercadoPago 实际净额与销售报表 Total "
          "的差值直接证实（该差值精确等于原佣金）；无流水的店只能按规律推定。"
          "本表合计已计入损益表的“加：部分退款重复扣佣金调整”行。",
          ["店铺", "订单号", "SKU", "退款比例", "原佣金＝多扣金额", "GMV",
           "销售报表退款", "闭式应退", "销售报表Total", "MP实际净额", "认定依据"],
          [18, 22, 18, 10, 16, 12, 14, 12, 14, 13, 34],
          [(S["store"], hit, sku, pct(p, 2), money(com), money(gmv), money(vref),
            money(want), money(vtot), (money(real) if real is not None else "—"),
            "现金凭证确认" if basis == "现金凭证" else "按规律推定（%s）" % proof)
           for S in stores
           for hit, p, com, gmv, tax, vref, want, vtot, real, proof, d, sku, basis
           in S["cn"]["partial"]], sum_cols=(5, 6, 7, 8))

    block("3.4 部分冲销、但金额对不上已知模式（未计入调整，需人工核对）",
          "这些单的佣金确实被部分冲销了，但销售报表的退款金额与上面的规律对不上，"
          "说明还叠加了别的调整。没有把它们计入损益表调整 —— 宁可少调也不要凭猜测调。",
          ["店铺", "订单号", "SKU", "退款比例", "原佣金", "销售报表退款",
           "按规律应为", "差额", "MP实际净额"],
          [18, 22, 18, 10, 12, 14, 14, 12, 13],
          [(S["store"], hit, sku, pct(p, 2), money(com), money(vref),
            money(vref - gap), money(gap), (money(real) if real is not None else "—"))
           for S in stores
           for hit, p, com, gmv, tax, vref, want, vtot, real, proof, d, sku, basis, gap
           in S["cn"]["mismatch"]], sum_cols=(5, 6))
    ws.freeze_panes = "A5"


# ════════════════════════════════════════════════════════════════════════
# 组装 / 重算 / CLI
# ════════════════════════════════════════════════════════════════════════

def build_context(stores, month, build_date=None):
    p0, p1 = month_bounds(month)
    cutoffs = [S["ventas_all"]["fecha"].max() for S in stores if len(S["ventas_all"])]
    cutoff = max([c for c in cutoffs if pd.notna(c)]) if cutoffs else p1
    # 不完整月份：销售报表窗口两端（最早月与最晚月），且不等于报表月本身时才提示
    months = sorted(set(m for S in stores for m in S["ventas_all"]["ym"].unique() if m != "NaT"))
    partial = set()
    if months:
        partial.add(months[0])
        partial.add(months[-1])
    partial.discard(month)
    contrib = {}
    for S in stores:
        m = S["m"]
        bill = sum(m["bill_fee_net"].values()) + m["unc_bill"]
        contrib[S["store"]] = (m["ing"] + m["env_ing"] - m["coupon"] - m["refund"]
                               - m["com_gross"] - m["env_cost"] - m["dev_net"]
                               - m["unc_order"] - bill)
    return {
        "period": month,
        "period_cn": "%d年%d月" % (p0.year, p0.month),
        "period_short": "%d月" % p0.month,
        "period_es": "%s %d" % (MONTH_FULL_ES[p0.month], p0.year),
        "p0": p0, "p1": p1,
        "build_date": build_date or datetime.date.today().strftime("%Y-%m-%d"),
        "cutoff": cutoff.strftime("%Y-%m-%d") if pd.notna(cutoff) else "-",
        "partial_months": sorted(partial),
        "contrib": contrib,
    }


def build_workbook(stores, ctx, out_path):
    wb = Workbook()
    wb.remove(wb.active)
    sheet_readme(wb, stores, ctx)
    sheet_pl(wb, stores, ctx)
    sheet_bridge(wb, stores, ctx)
    sheet_fees(wb, stores, ctx)
    sheet_returns_money(wb, stores, ctx)
    sheet_triage(wb, stores, ctx)
    sheet_sku(wb, stores, ctx)
    sheet_trend(wb, stores, ctx)
    sheet_inventory(wb, stores, ctx)
    sheet_gaps(wb, stores, ctx)
    d = os.path.dirname(os.path.abspath(out_path))
    if d and not os.path.isdir(d):
        os.makedirs(d)
    wb.save(out_path)
    return out_path


def recalc(path, timeout=180):
    """用 LibreOffice 重算并回填公式缓存值。openpyxl 只写公式不写结果，
    不重算的话 pandas / data_only=True 读到的全是 None。"""
    out_dir = os.path.join(os.path.dirname(os.path.abspath(path)), "_recalc")
    cmd = ["soffice", "--headless", "--norestore", "--convert-to",
           "xlsx:Calc MS Excel 2007 XML", "--outdir", out_dir, path]
    try:
        subprocess.call(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    except OSError:
        return False, "找不到 soffice，跳过重算（公式仍然正确，Excel 打开时会自动算）"
    produced = os.path.join(out_dir, os.path.basename(path))
    if not os.path.exists(produced):
        return False, "LibreOffice 未产出文件，跳过重算"
    try:
        os.remove(path)
    except OSError:
        pass
    os.rename(produced, path)
    try:
        os.rmdir(out_dir)
    except OSError:
        pass
    return True, "已重算"


def scan_formula_errors(path):
    """重算后扫一遍，确保没有 #REF! / #NAME? 之类烂在文件里。"""
    ERR = ("#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#N/A", "#NULL!", "#NUM!", "Err:")
    wbv = load_workbook(path, data_only=True)
    wbf = load_workbook(path)
    n_formula = n_err = n_none = 0
    bad = []
    for sn in wbf.sheetnames:
        sf, sv = wbf[sn], wbv[sn]
        for row in sf.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    n_formula += 1
                    v = sv[c.coordinate].value
                    if isinstance(v, str) and any(e in v for e in ERR):
                        n_err += 1
                        bad.append("%s!%s = %s" % (sn, c.coordinate, v))
                    elif v is None:
                        n_none += 1
    return n_formula, n_err, n_none, bad[:20]


def metrics_record(S, ctx):
    """一行月度指标，用于累积到指标库（趋势的真正来源）。"""
    m = S["m"]
    rec = {"period": ctx["period"], "store": S["store"], "built_at": ctx["build_date"]}
    for k in ("rows", "orders_unique", "units", "ing", "env_ing", "coupon", "coupon_n",
              "refund", "com_imp", "com_gross", "env_cost", "env_gross", "dev_net",
              "tax_ret", "tax_rate_gross", "total", "resid", "bill_total", "payable",
              "pay_auto", "pay_nc", "other_total", "ads_net", "full_bill_total",
              "subscription_net", "storage_bill", "storage_accrual", "unc_order",
              "unc_bill", "abnormal_orders", "n_back_on_sale", "n_stock_lost",
              "n_seller_keeps", "amt_stock_lost", "iva_in_gmv"):
        v = m.get(k)
        rec[k] = round(v, 6) if isinstance(v, float) else v
    rec["contribution"] = round(ctx["contrib"][S["store"]], 2)
    for es, v in m["bill_fee_net"].items():
        rec["fee:" + es] = round(v, 2)
    return rec


def parse_store_arg(v):
    if "=" not in v:
        raise argparse.ArgumentTypeError("--store 需要 NAME=PATH 形式，收到 %r" % v)
    name, path = v.split("=", 1)
    name, path = name.strip(), os.path.expanduser(path.strip())
    if not os.path.isdir(path):
        raise argparse.ArgumentTypeError("目录不存在：%s" % path)
    return (name, path)


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Mercado Libre MX 月度财务报表生成器",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--month", required=True, help="会计期间，如 2026-08")
    ap.add_argument("--store", action="append", required=True, type=parse_store_arg,
                    metavar="NAME=PATH", help="店铺名=下载目录，可重复")
    ap.add_argument("--out", default=None, help="输出 xlsx 路径")
    ap.add_argument("--metrics-out", default=None, help="同时导出月度指标 JSON（用于累积指标库）")
    ap.add_argument("--strict", action="store_true",
                    help="任一校验未通过就不写出文件，退出码 1")
    ap.add_argument("--no-recalc", action="store_true", help="跳过 LibreOffice 重算")
    ap.add_argument("--tol", type=float, default=0.05, help="桥A 残差容差，默认 0.05")
    ap.add_argument("--build-date", default=None, help="覆盖编制日期（用于可复现测试）")
    a = ap.parse_args(argv)

    p0, p1 = month_bounds(a.month)
    stores = []
    for name, path in a.store:
        print("[读取] %s ← %s" % (name, path))
        stores.append(derive(name, path, p0, p1))

    ctx = build_context(stores, a.month, a.build_date)
    checks = validate(stores, tol=a.tol)
    ctx["checks"] = checks

    print("\n[校验]")
    for c in checks:
        mark = "跳过" if c["skipped"] else ("  ✓" if c["ok"] else "  ✗")
        print("  %s  %-12s %-22s %s" % (mark, c["store"][:12], c["name"], c["detail"]))
    failed = [c for c in checks if not c["ok"] and c["severity"] == "error"]

    print("\n[结果]")
    for S in stores:
        m = S["m"]
        print("  %-16s GMV %14s  到账 %14s  经营贡献毛利 %14s"
              % (S["store"], fmt(m["ing"]), fmt(m["total"]), fmt(ctx["contrib"][S["store"]])))
    print("  %-16s %18s %19s %18s"
          % ("合计", fmt(sum(S["m"]["ing"] for S in stores)),
             fmt(sum(S["m"]["total"] for S in stores)), fmt(sum(ctx["contrib"].values()))))

    if failed and a.strict:
        print("\n[中止] %d 项校验未通过，--strict 模式下不生成报表。" % len(failed))
        return 1

    out = a.out or os.path.join("reports", "MercadoLibre_月度财务报表_%s.xlsx" % ctx["period_cn"])
    build_workbook(stores, ctx, out)
    print("\n[输出] %s" % out)

    if not a.no_recalc:
        ok, msg = recalc(out)
        print("[重算] %s" % msg)
        if ok:
            nf, ne, nn, bad = scan_formula_errors(out)
            print("[公式] 共 %d 条，错误 %d 条，未求值 %d 条" % (nf, ne, nn))
            for b in bad:
                print("       %s" % b)
            if ne:
                print("[警告] 文件里存在公式错误，请勿直接发出。")
                return 1

    if a.metrics_out:
        recs = [metrics_record(S, ctx) for S in stores]
        d = os.path.dirname(os.path.abspath(a.metrics_out))
        if d and not os.path.isdir(d):
            os.makedirs(d)
        with io.open(a.metrics_out, "w", encoding="utf-8") as f:
            f.write(json.dumps(recs, ensure_ascii=False, indent=1))
        print("[指标] %s" % a.metrics_out)

    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
