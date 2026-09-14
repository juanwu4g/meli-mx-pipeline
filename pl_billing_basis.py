# -*- coding: utf-8 -*-
"""
【验证用·可随时删除】按「计费口径」重算损益表，与 report_build.py 的现行口径对比。

    python pl_billing_basis.py --store EWTTO_SM --month 2026-08
    python pl_billing_basis.py --store EWTTO_SM --month 2026-08 --out 对比.xlsx

为什么单独一个文件
------------------
`report_build.py` 的 `derive()` 产出的 `m` 字典被损益表、勾稽表、SKU 页、趋势页
共用。要在它里面塞第二套口径，就得给每个数加分支，改动面铺开到整个文件 —— 而这
是个验证性实验，结论可能是推翻。所以这里只 **import 复用** 它的
`load_ventas` / `load_billing` / `FEE_CATALOG` / `derive`，绝不修改它；删掉本文件
不会留下任何痕迹。

两种口径的差别
--------------
现行（混合口径）：
  · 收入、运费收入、退款、券、**配送费** ← Ventas 报表，按【订单日】
  · **佣金、退货处理费**                 ← 账单，但按【订单号关联】，跨账期
  · 广告 / 仓储 / 揽收等月度费用          ← 账单，按【计费日】

本文件（纯计费口径）：
  · 费用**全部**来自当月账单（Fecha del cargo 落在本月），按科目分组
  · 收入仍来自 Ventas —— 账单里没有收入，这一侧没有替代来源

所以对比只在**费用侧**有意义，收入行两边必然相同。差额的来源有三类：
  ① 跨账期错位：8 月的订单，费用可能记在 9 月账单（或反之）
  ② 配送费换了源：Ventas 的 Costos de envío vs 账单的 Cargo por envíos
  ③ 代扣税：账单里根本没有这个科目（平台只在 Ventas 里混着给），
     所以纯计费口径**算不出代扣税**，这本身就是一个结论
"""
import argparse
import console  # noqa: F401  中文输出编码保护
import os
import sys

import pandas as pd

import report_build as rb
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter


def billing_basis(store, folder, month):
    """当月账单（计费日）按科目的净额。返回 (科目->金额, 汇总信息)。"""
    p0, p1 = rb.month_bounds(month)
    B = rb.load_billing(folder)
    FA = B[(B["cdate"] >= p0) & (B["cdate"] <= p1)].copy()

    per_fee, seen = {}, set()
    for es, zh, bucket, order_level in rb.FEE_CATALOG:
        v = float(FA[FA["parent_detalle"].astype(str) == es]["amount"].sum())
        if abs(v) > 0.005:
            per_fee[es] = v
        seen.add(es)
    # 未登记科目单列，绝不悄悄并进"其他" —— 与 report_build 的原则一致
    unknown = FA[~FA["parent_detalle"].astype(str).isin(seen)]
    unk_order = unknown[unknown["sdate"].notna()]
    unk_bill = unknown[unknown["sdate"].isna()]
    info = {
        "unknown_order": float(unk_order["amount"].sum()),
        "unknown_bill": float(unk_bill["amount"].sum()),
        "rows": len(FA),
        "total": float(FA["amount"].sum()),
        "order_linked": int(FA["sdate"].notna().sum()),
        "account_level": int(FA["sdate"].isna().sum()),
        "reversals": int(FA["is_rev"].sum()),
        "unknown_total": float(unknown["amount"].sum()),
        "unknown_names": sorted(set(unknown["parent_detalle"].astype(str)))[:8],
        "cn_adj": 0.0,          # 由调用方用现行口径的同名值填入
    }
    return per_fee, info


def current_basis(S):
    """现行口径下，损益表各费用行的金额（正数=费用）。"""
    m = S["m"]
    out = {
        "Cargo por venta": m["com_gross"],
        "Cargo por envíos de Mercado Libre": m["env_cost"],   # ← 来自 Ventas
        "Cargo por devolución": m["dev_net"],
    }
    for es, v in m["bill_fee_net"].items():
        if abs(v) > 0.005:
            out[es] = out.get(es, 0.0) + v
    return out


def compare(store, folder, month):
    p0, p1 = rb.month_bounds(month)
    S = rb.derive(store, folder, p0, p1)
    cur = current_basis(S)
    bil, info = billing_basis(store, folder, month)

    keys = sorted(set(cur) | set(bil),
                  key=lambda k: -abs(bil.get(k, 0.0) or cur.get(k, 0.0)))
    rows = []
    for es in keys:
        zh = rb.FEE_BY_ES.get(es, ("(未登记)", None, None))[0]
        a, b = cur.get(es, 0.0), bil.get(es, 0.0)
        rows.append((zh, es, a, b, b - a))
    return S, rows, info, bil


SRC_NOTE = {
    "Cargo por venta": "现行取自账单但按【订单号关联】跨账期；计费口径按【计费日】",
    "Cargo por envíos de Mercado Libre": "现行取自 Ventas 的 Costos de envío；计费口径取自账单",
    "Cargo por devolución": "现行按【订单号关联】跨账期；计费口径按【计费日】",
}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--store", default="EWTTO_SM")
    ap.add_argument("--month", default="2026-08")
    ap.add_argument("--folder", default=None,
                    help="下载目录；默认自动取最近一个文件齐全的")
    ap.add_argument("--out", default=None, help="同时写出 xlsx 对比表")
    a = ap.parse_args()

    folder = a.folder
    if not folder:
        import financial_report
        folder, note = financial_report.latest_run(a.store)
        if not folder:
            print("找不到可用的下载目录：%s" % note)
            return 1
        print("下载目录：%s" % note)

    S, rows, info, bil = compare(a.store, folder, a.month)
    m = S["m"]

    print("=" * 96)
    print("口径对比 —— %s  %s" % (a.store, a.month))
    print("=" * 96)
    print("  当月账单：%d 行（挂订单 %d，账单级 %d，其中冲销 %d），金额合计 %s"
          % (info["rows"], info["order_linked"], info["account_level"],
             info["reversals"], rb.fmt(info["total"])))
    if abs(info["unknown_total"]) > 0.005:
        print("  ⚠ 未登记科目 %s：%s"
              % (rb.fmt(info["unknown_total"]), "、".join(info["unknown_names"])))
    print()
    print("  %-26s %16s %16s %14s  %s"
          % ("费用科目", "现行(混合口径)", "计费口径", "差额", "说明"))
    print("  " + "-" * 94)
    ta = tb = 0.0
    for zh, es, x, y, d in rows:
        ta += x
        tb += y
        flag = "" if abs(d) < 0.01 else ("  ←" if abs(d) > 1000 else "")
        print("  %-26s %16s %16s %14s%s  %s"
              % (zh[:26], rb.fmt(x), rb.fmt(y), rb.fmt(d), flag,
                 SRC_NOTE.get(es, "")[:34]))
    print("  " + "-" * 94)
    print("  %-26s %16s %16s %14s" % ("费用合计", rb.fmt(ta), rb.fmt(tb), rb.fmt(tb - ta)))
    print()
    print("  收入侧（两种口径相同，账单里没有收入）")
    print("    商品销售收入 GMV %18s" % rb.fmt(m["ing"]))
    print("    运费收入         %18s" % rb.fmt(m["env_ing"]))
    print()
    print("  经营贡献毛利（收入 − 费用）")
    base = m["ing"] + m["env_ing"] - m["refund"] - m["coupon"]
    print("    现行口径 %18s" % rb.fmt(base - ta))
    print("    计费口径 %18s   差 %s" % (rb.fmt(base - tb), rb.fmt(ta - tb)))
    print()
    print("  ⚠ 代扣代缴税金：现行口径推算为 %s（Ventas佣金税合计 − 账单纯佣金）；"
          % rb.fmt(m["tax_ret"]))
    print("    计费口径**算不出来** —— 账单里没有这个科目，平台只在 Ventas 里混着给。")

    if a.out:
        from openpyxl import Workbook
        wb = Workbook()
        wb.remove(wb.active)
        info["cn_adj"] = S["cn"]["partial_total"]
        sheet_pl_billing(wb, a.store, a.month, m, bil, info, current_pl_values(S))
        _compare_sheet(wb, a.store, a.month, rows, m, info)
        wb.save(a.out)
        print("")
        print("  已写出 %s" % os.path.abspath(a.out))
        print("     ② 损益表(计费口径) —— 行结构与现行报表一致，附现行口径与差额两列")
        print("     ③ 科目对比          —— 逐科目差异一览")
    return 0


def _compare_sheet(wb, store, month, rows, m, info):



    ws = wb.create_sheet("③ 科目对比")

    ws["A1"] = "损益表口径对比 —— %s %s" % (store, month)
    ws["A1"].font = rb.F(14, True, rb.C_NAVY)
    ws["A2"] = ("左=现行混合口径（收入与配送费按订单日、佣金按订单号关联、月度费用按计费日）；"
                "右=纯计费口径（费用全部按当月账单的计费日）。收入侧两者相同。")
    ws["A2"].font = rb.F(9, False, rb.C_GREY)
    head = ["费用科目", "西班牙语原始科目", "现行(混合口径)", "计费口径", "差额", "说明"]
    for i, t in enumerate(head, 1):
        c = ws.cell(row=4, column=i, value=t)
        c.font = rb.F(10, True, "FFFFFF")
        c.fill = rb.fill(rb.C_NAVY)
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        c.border = rb.BOX
    for w, i in zip([26, 46, 16, 16, 14, 48], range(1, 7)):
        ws.column_dimensions[chr(64 + i)].width = w
    r = 5
    ta = tb = 0.0
    for zh, es, x, y, d in rows:
        ta += x; tb += y
        ws.cell(row=r, column=1, value=zh)
        ws.cell(row=r, column=2, value=es).font = rb.F(9, False, rb.C_GREY)
        ws.cell(row=r, column=3, value=rb.money(x)).number_format = rb.MNY
        ws.cell(row=r, column=4, value=rb.money(y)).number_format = rb.MNY
        ws.cell(row=r, column=5, value=rb.money(d)).number_format = rb.MNY
        ws.cell(row=r, column=6, value=SRC_NOTE.get(es, "")).font = rb.F(9, False, rb.C_GREY)
        if abs(d) > 0.01:
            for c in range(1, 6):
                ws.cell(row=r, column=c).fill = rb.fill(rb.C_WARN)
        for c in range(1, 7):
            ws.cell(row=r, column=c).border = rb.BOX
        r += 1
    ws.cell(row=r, column=1, value="费用合计").font = rb.F(10, True)
    for col, v in ((3, ta), (4, tb), (5, tb - ta)):
        ws.cell(row=r, column=col, value=rb.money(v)).number_format = rb.MNY
        ws.cell(row=r, column=col).font = rb.F(10, True)
    for c in range(1, 7):
        ws.cell(row=r, column=c).fill = rb.fill(rb.C_SUB)
        ws.cell(row=r, column=c).border = rb.BOX
    r += 2
    for label, v in (("商品销售收入 GMV（两口径相同）", m["ing"]),
                     ("运费收入（两口径相同）", m["env_ing"]),
                     ("代扣代缴税金（仅现行口径可推算）", m["tax_ret"])):
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=3, value=rb.money(v)).number_format = rb.MNY
        r += 1
    return ws


# ════════════════════════════════════════════════════════════════════════
# 计费口径损益表 —— 行结构与现行报表完全一致，只有取数来源不同
# ════════════════════════════════════════════════════════════════════════

# 计费口径下取不到的行。写 None 而不是 0：0 会被误读成"这个月没有"，
# None 显示为「无法取得」，并在说明列讲清楚为什么。
UNAVAILABLE = {
    "com_neg": "账单的纯佣金已在本行；但代扣税无法从账单取得，见表末备查",
}


def billing_pl_value(key, m, bil, info):
    """计费口径下某一行的金额。返回 None 表示该口径下取不到。

    收入侧没有替代来源（账单里只有费用），所以照旧取 Ventas；
    费用侧一律改取当月账单（计费日）。
    """
    if key is None:
        return None
    # ---- 收入侧：账单里没有，只能沿用 Ventas ----
    if key == "ing":
        return m["ing"]
    if key == "env_ing":
        return m["env_ing"]
    if key == "coupon_neg":
        return -m["coupon"]
    if key == "refund_neg":
        return -m["refund"]
    if key == "cn_adj":
        # 这一行修正的是销售报表退款列的错误，与费用归集口径无关，
        # 两种口径都要带上，否则对比列会凭空多出一笔差额。
        return info.get("cn_adj", 0.0)
    if key == "cogs":
        return 0.0
    # ---- 费用侧：全部来自当月账单 ----
    if key == "com_neg":
        return -(bil.get("Cargo por venta", 0.0)
                 + bil.get("Cargo por venta con afiliados", 0.0))
    if key == "shipcost_neg":
        return -bil.get("Cargo por envíos de Mercado Libre", 0.0)
    if key == "dev_neg":
        return -bil.get("Cargo por devolución", 0.0)
    if key == "unc_order_neg":
        return -info.get("unknown_order", 0.0)
    if key == "unc_bill_neg":
        return -info.get("unknown_bill", 0.0)
    if key.startswith("fee:"):
        return -bil.get(key[4:], 0.0)
    return 0.0


def sheet_pl_billing(wb, store, month, m, bil, info, cur_rows=None):
    """按现行报表的行结构出一张计费口径损益表。

    行序、行数、公式结构与 report_build 的 ② 合并损益表**完全一致** —— 直接复用
    它的 build_pl_rows()，这样两份表可以逐行并排对照，不用人工对齐。
    差别只在：C 列的取数换成当月账单（计费日），D 列附上现行口径供对比。
    """
    ws = wb.create_sheet("② 损益表(计费口径)")
    ws["A1"] = "损益表 · 计费口径（%s %s）" % (store, month)
    ws["A1"].font = rb.F(14, True, rb.C_NAVY)
    ws.row_dimensions[1].height = 24
    ws["A2"] = ("口径：**费用全部按当月账单的【计费日】(Fecha del cargo) 归集**；"
                "收入仍取自 Ventas 报表（账单里没有收入，这一侧没有替代来源）。")
    ws["A2"].font = rb.F(9, False, rb.C_RED)
    ws["A3"] = ("与现行报表的差别只在费用侧取数：现行把佣金/退货费按【订单号关联】跨账期抓取、"
                "配送费取自 Ventas 的 Costos de envío；本表一律只认当月账单上的行。"
                "行序与现行 ② 合并损益表完全一致，可逐行并排对照。")
    ws["A3"].font = rb.F(9, False, rb.C_GREY)

    cols = ["项目", "西班牙语原始科目", "计费口径", "现行(混合口径)", "差额", "占GMV%", "说明"]
    widths = [30, 42, 16, 16, 14, 10, 52]
    for i, t in enumerate(cols, 1):
        c = ws.cell(row=4, column=i, value=t)
        c.font = rb.F(10, True, "FFFFFF")
        c.fill = rb.fill(rb.C_NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = rb.BOX
        ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]
    ws.row_dimensions[4].height = 26

    rows = rb.build_pl_rows()
    r0 = 5
    gmv_row = r0            # 第一行就是 GMV，占比公式的分母
    idx = {}
    for i, (key, label, es, kind, note) in enumerate(rows):
        r = r0 + i
        idx[label] = r
        ws.cell(row=r, column=1, value=label).font = rb.F(10, kind in ("sum", "f"))
        ws.cell(row=r, column=2, value=es or "").font = rb.F(9, False, rb.C_GREY)

        if kind in ("d", "in"):
            v = billing_pl_value(key, m, bil, info)
            cell = ws.cell(row=r, column=3)
            if v is None:
                cell.value = "无法取得"
                cell.font = rb.F(10, False, rb.C_RED)
            else:
                cell.value = rb.money(v)
                cell.number_format = rb.MNY
            if kind == "in":
                cell.fill = rb.fill(rb.C_INPUT)
        elif kind in ("sum", "f"):
            ws.cell(row=r, column=3).number_format = rb.MNY   # 公式稍后按标签统一写

        # D 列：现行口径，供逐行对照
        if cur_rows is not None and i < len(cur_rows):
            cv = cur_rows[i]
            if cv is not None:
                ws.cell(row=r, column=4, value=rb.money(cv)).number_format = rb.MNY
                ws.cell(row=r, column=5,
                        value="=IF(OR(ISTEXT(C%d),ISBLANK(D%d)),\"\",C%d-D%d)" % (r, r, r, r)
                        ).number_format = rb.MNY
        ws.cell(row=r, column=6,
                value="=IF(OR(ISTEXT(C%d),$C$%d=0),\"\",C%d/$C$%d)" % (r, gmv_row, r, gmv_row)
                ).number_format = rb.PCT
        ws.cell(row=r, column=7, value=note or "").font = rb.F(9, False, rb.C_GREY)
        ws.cell(row=r, column=7).alignment = Alignment(wrap_text=True, vertical="top")

        if kind == "sum":
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = rb.fill(rb.C_SUB)
        elif kind == "f":
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = rb.fill(rb.C_CALC)
                ws.cell(row=r, column=c).font = rb.F(10, True)
        for c in range(1, 8):
            ws.cell(row=r, column=c).border = rb.BOX

    # 合计/流水行的公式：照搬 report_build.sheet_pl 的写法 —— 按**标签**取行号，
    # 不靠位置推断。行序一旦调整，公式跟着自动对，不会悄悄错行。
    R = lambda label: idx[label]
    put = lambda label, expr: ws.cell(row=R(label), column=3, value=expr)
    put("营业收入合计", "=SUM(C%d:C%d)"
        % (R("商品销售收入 (GMV)"), R("减：卖家承担优惠券/折扣")))
    put("净销售收入", "=C%d+C%d" % (R("营业收入合计"), R("减：退款与取消（净）")))
    put("平台交易费用小计", "=SUM(C%d:C%d)"
        % (R("减：平台销售佣金"), R("减：订单级未登记科目 ⚠")))
    put("订单毛贡献", "=C%d+C%d" % (R("净销售收入"), R("平台交易费用小计")))
    first_bill = R("减：" + rb.FEE_BY_ES["Cargo por campaña de publicidad de Product Ads"][0])
    put("月度账单费用小计", "=SUM(C%d:C%d)" % (first_bill, R("减：账单级未登记科目 ⚠")))
    put("经营贡献毛利（未扣商品成本）", "=C%d+C%d"
        % (R("订单毛贡献"), R("月度账单费用小计")))
    put("税前净利润", "=C%d-C%d"
        % (R("经营贡献毛利（未扣商品成本）"), R("减：商品采购成本")))
    for label in ("营业收入合计", "净销售收入", "平台交易费用小计", "订单毛贡献",
                  "月度账单费用小计", "经营贡献毛利（未扣商品成本）", "税前净利润"):
        ws.cell(row=R(label), column=3).number_format = rb.MNY

    r = r0 + len(rows) + 1
    ws.cell(row=r, column=1, value="备查项（计费口径取不到的数）").font = rb.F(11, True, rb.C_NAVY)
    r += 1
    for label, v, why in (
        ("平台代扣代缴税金 (IVA/ISR)", m["tax_ret"],
         "账单里**没有这个科目** —— 平台只在 Ventas 的 Cargo por venta e impuestos 里"
         "把它和佣金混在一起给。所以纯计费口径推算不出来，此处填的是现行口径的值。"),
        ("当月账单行数", float(info["rows"]),
         "挂订单 %d 行、账单级 %d 行，其中冲销 %d 行" %
         (info["order_linked"], info["account_level"], info["reversals"])),
        ("当月账单金额合计", info["total"],
         "这是平台本月实际计费的总额；上表费用合计应与之相等（未登记科目已单列）"),
    ):
        ws.cell(row=r, column=1, value=label).font = rb.F(10)
        ws.cell(row=r, column=3, value=rb.money(v)).number_format = rb.MNY
        ws.cell(row=r, column=7, value=why).font = rb.F(9, False, rb.C_GREY)
        ws.cell(row=r, column=7).alignment = Alignment(wrap_text=True, vertical="top")
        for c in range(1, 8):
            ws.cell(row=r, column=c).border = rb.BOX
        r += 1

    ws.freeze_panes = "A5"
    return ws


def current_pl_values(S):
    """现行口径下、与 build_pl_rows() 同序的每行金额（明细行才有值）。"""
    out = []
    for key, label, es, kind, note in rb.build_pl_rows():
        out.append(rb.pl_value(S, key) if kind in ("d", "in") else None)
    return out


if __name__ == "__main__":
    sys.exit(main())
