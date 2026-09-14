"""Chinese labels for the output columns.

The workbook is read by both an English-reading engineer and a Chinese-reading
accountant, so every sheet carries the English name on row 1 and the Chinese on row 2.
Columns with no entry here fall back to their English name, which keeps a newly added
column visible rather than blank.
"""

from __future__ import annotations

#: Output column -> Chinese label. Sheet column letters are noted where the column comes
#: straight from the source sheet, so the two can be lined up during review.
CHINESE: dict[str, str] = {
    # --- identity and status ---
    "order_id": "订单号",                              # A
    "sold_at": "销售日期",                             # B
    "order_status": "状态",                            # C
    "order_status_detail": "状态说明",                  # D
    "is_multi_product_package": "是否组合包",            # E
    "belongs_to_kit": "是否属于套装",                    # F
    "units": "数量",                                   # G
    # --- money, columns H..P ---
    "product_revenue_mxn": "产品收入 (H)",
    "sale_fee_and_taxes_mxn": "销售费用及税金 (I)",
    "shipping_revenue_mxn": "运费收入 (J)",
    "shipping_cost_mxn": "运费成本 (K)",
    "shipping_cost_declared_mxn": "按申报尺寸重量的运费 (L)",
    "shipping_dimension_adjustment_mxn": "尺寸重量差异费 (M)",
    "discounts_and_bonuses_mxn": "折扣与补贴 (N)",
    "cancellations_and_refunds_mxn": "取消与退款 (O)",
    "total_mxn": "合计 (P)",
    "net_margin_mxn": "净额",
    "purchase_order": "采购单号",
    # --- listing ---
    "is_advertising_sale": "广告成交",
    "sku": "SKU",
    "listing_id": "商品编号",
    "listing_title": "商品标题",
    "listing_variant": "规格",
    "listing_variant_attributes": "规格明细",
    "unit_price_mxn": "单价 (W)",
    "listing_type": "刊登类型",
    # --- shipping ---
    "shipping_method": "配送方式",
    "shipped_at": "发货时间",
    "delivered_at": "送达时间",
    "shipping_carrier": "承运商",
    "shipping_tracking_number": "运单号",
    "shipping_tracking_url": "物流查询链接",
    "days_to_deliver": "送达天数",
    # --- returns ---
    "returned_units": "退货数量",
    "return_shipping_method": "退货配送方式",
    "return_shipped_at": "退货发出时间",
    "return_delivered_at": "退货送达时间",
    "return_carrier": "退货承运商",
    "return_tracking_number": "退货运单号",
    "return_tracking_url": "退货查询链接",
    "return_reviewed_by_ml": "平台已验货",
    "return_reviewed_at": "验货时间",
    "return_money_favours": "退款归属方",
    "return_inspection_result": "验货结果",
    "return_destination": "退货去向",
    "return_result_reason": "验货结果原因",
    # --- claims ---
    "claim_units": "纠纷数量",
    "has_open_claim": "存在未结纠纷",
    "claims_closed": "已结纠纷数",
    "has_mediation": "经过调解",
    # --- derived by the rules ---
    "tax_withholding_mxn": "代扣税金 (规则5)",
    "platform_fee_mxn": "平台佣金 (规则6)",
    "sale_period": "会计期间 YYYYMM (规则7)",
    "components_sum_mxn": "H至O合计 (规则1)",
    "total_diff_mxn": "与P的差异 (规则1)",
    "total_matches": "是否平账 (规则1)",
    "sale_year": "销售年份",
    "sale_month": "销售月份",
    # --- package handling ---
    "row_role": "行类型",
    "package_id": "组合包编号",
    "package_size": "组合包件数",
    "package_split": "组合包已拆分 (规则2)",
    "package_note": "组合包备注 (规则3)",
    "allocated_from_package": "由组合包分摊 (规则2)",
    "is_returned_item": "组合包内退货商品 (规则3)",
    # --- lineage ---
    "row_index": "原表行号",
    "record_seq": "重复行序号",
    "store": "店铺",
    "seller_id": "卖家编号",
    "source_file": "来源文件",
    "source_modified_at": "来源文件时间",
    "ingested_at": "处理时间",
    # --- sku sheet, worded to match the accountant's 销售 tab ---
    "fee_commission_mxn": "销售费",
    "fee_commission_reversal_mxn": "取消销售费",
    "fee_shipping_mxn": "Mercado Libre 运费",
    "fee_shipping_reversal_mxn": "取消 Mercado Libre 运费",
    "fee_return_mxn": "退款费用",
    "fee_total_mxn": "费用合计",
    "settlement_amount_mxn": "结算金额",
    "settlement_unit_price_mxn": "结算单价",
    "k3_material": "K3物料",
    # --- 报告明细 working columns (already Chinese; keep them as they are) ---
    "年月": "年月", "店铺": "店铺", "币别": "币别", "费用项目": "费用项目", "标志": "标志",
    "引用R": "引用R", "引用AB": "引用AB", "重次": "重次", "税金": "税金",
    "调减费用": "调减费用", "抵扣年月": "抵扣年月",
    "advertising_mxn": "广告费",
    "storage_mxn": "平台仓租",
    "net_after_account_costs_mxn": "扣除广告仓租后",
    # --- rule 10: revenue exclusions ---
    "counts_toward_revenue": "计入收入 (规则10)",
    "revenue_exclusion_reason": "不计入原因 (规则10)",
    "excluded_orders": "不计入订单数",
    "excluded_revenue_mxn": "不计入金额",
    "gross_revenue_mxn": "收入总额(含不计入)",
    # --- costos sheet ---
    "charge_month": "费用月份",
    "charge_detail": "费用说明 (Detalle)",
    "fee_category": "费用类别",
    "charges": "笔数",
    "tied_to_sale": "关联到订单",
    "total_mxn": "金额合计",
    # --- summary sheets ---
    "metric": "项目",
    "source": "原表",
    "output": "输出",
    "difference": "差异",
    "stage": "阶段",
    "rows": "行数",
    "checked": "已核对",
    "not_checked": "无法核对",
    "matching": "平账",
    "differing": "有差异",
    "net_difference_mxn": "差异净额",
    "largest_difference_mxn": "最大差异",
    "orders": "订单数",
    "revenue_mxn": "收入",
    "tax_withheld_mxn": "代扣税金",
    "refunds_mxn": "退款",
    "payout_mxn": "实收",
    "unexplained_mxn": "未解释差异",
}


def chinese(column: str) -> str:
    """The Chinese label for a column, falling back to the English name."""
    return CHINESE.get(column, column)
