"""微信群订单 → AI 解析 → 核对表 + 智慧记销售单导入文件。

用法：
    python parse_orders.py 2_微信群消息_10-21.txt            # 货品/客户从测试 xls 读取
    python parse_orders.py 消息.txt --db "D:/.../zhjd/data.db"  # 货品/客户直接只读智慧记数据库
"""
import argparse
import json
import os
import re
import sqlite3
import sys
from datetime import date
from pathlib import Path

import xlrd
import xlwt
from openai import OpenAI
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

HERE = Path(__file__).parent
MODEL = "deepseek-flash"

SYSTEM_PROMPT = """你是食品批发商的录单助手。输入是从微信群复制出来的聊天记录，以及系统里的货品清单和客户清单。
请把聊天记录整理成销售单，严格按 JSON 输出。

规则：
1. 一个客户在同一天的所有订货合并成一张单（包括后面「再加」「补」的消息）。
2. 后面的改价、改数量消息要覆盖前面的值，不要重复出行。
3. 每个商品必须从货品清单里选 code；群里叫法是简称、别名、错别字时按意思匹配（如「牛丸」→撒尿牛丸，「金龙鱼油」→金龙鱼大豆油，「鸡腿」→冰鲜鸡腿）。
   有规格区分时（如「小包装」）必须选对应的那个。清单里确实没有的，code 填 null，raw_name 写原话。
4. 数量写阿拉伯数字（「十件」→10，「两箱」→2）。箱、件视为同一单位。
5. 消息里写了单价就填 price；没写填 null（后续由系统按批发价补）。
6. 赠品（送、搭赠、送本品）单独一行，is_gift=true，price=0。
7. customer 用客户清单里的名字；清单里没有就写群里的原名，customer_in_list=false。
8. 送货方式、付款情况、补货说明、地址等写进 remark。
9. 需要人确认的情况写进 warnings（如：货品不在清单、新客户、「按老价」、涉及历史欠货/搭赠、意思不明确）。
10. 「收到」「几点送」等闲聊忽略。日期没写年份按今年。

输出格式：
{"orders":[{"customer":"","customer_in_list":true,"date":"YYYY-MM-DD",
  "items":[{"code":"YR001","raw_name":"群里原话","qty":0,"price":null,"is_gift":false}],
  "remark":"","warnings":[""],"source":"对应的原始消息摘录"}]}
"""


def load_catalog_xls():
    wb = xlrd.open_workbook(HERE / "1_基础资料" / "货品资料_测试.xls")
    sh = wb.sheet_by_index(0)
    products = [{"code": sh.cell_value(r, 0), "name": sh.cell_value(r, 1),
                 "unit": sh.cell_value(r, 4), "trade_prc": float(sh.cell_value(r, 11) or 0)}
                for r in range(1, sh.nrows)]
    wb = xlrd.open_workbook(HERE / "1_基础资料" / "客户资料_测试.xls")
    sh = wb.sheet_by_index(0)
    customers = [sh.cell_value(r, 0) for r in range(1, sh.nrows)]
    return products, customers


def load_catalog_db(db_path):
    # 只读打开，绝不写智慧记数据库
    con = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True)
    products = [{"code": c, "name": n, "unit": u, "trade_prc": p or 0}
                for c, n, u, p in con.execute(
                    "select code, name, unit, trade_prc from products where is_del=0 and is_stop=0")]
    # companies.tye: 客户/供应商类型，按实际账套核对
    customers = [n for (n,) in con.execute("select name from companies where is_del=0 and is_stop=0")]
    return products, customers


def ask_llm(chat_text, products, customers):
    client = OpenAI(api_key=os.environ["DEEPSEEK_API_KEY"], base_url="https://api.deepseek.com")
    catalog = "\n".join(f"{p['code']} | {p['name']} | 批发价{p['trade_prc']}" for p in products)
    user = (f"今天是 {date.today()}。\n\n【货品清单】code | 品名规格 | 批发价\n{catalog}\n\n"
            f"【客户清单】\n{'、'.join(customers)}\n\n【群聊记录】\n{chat_text}")
    resp = client.chat.completions.create(
        model=MODEL,
        messages=[{"role": "system", "content": SYSTEM_PROMPT}, {"role": "user", "content": user}],
        response_format={"type": "json_object"},
        temperature=0,
    )
    return json.loads(resp.choices[0].message.content)["orders"]


def postprocess(orders, products, customers):
    """AI 只负责理解，价格和校验用确定性代码做。"""
    by_code = {p["code"]: p for p in products}
    for o in orders:
        o.setdefault("warnings", [])
        o["warnings"] = [w for w in o["warnings"] if w]
        if o["customer"] not in customers:
            o["customer_in_list"] = False
            _warn(o, f"新客户「{o['customer']}」，需先在智慧记里建档")
        for it in o["items"]:
            p = by_code.get(it.get("code"))
            if not p:
                it["code"] = None
                it["name"] = ""
                _warn(o, f"货品「{it['raw_name']}」不在系统里")
                continue
            it["name"] = p["name"]
            if it.get("is_gift"):
                it["price"] = 0
            elif it.get("price") is None:
                it["price"] = p["trade_prc"]
                it["price_from"] = "批发价"
            it["amount"] = round(it["qty"] * it["price"], 2)
        o["total"] = round(sum(it.get("amount", 0) for it in o["items"]), 2)
    return orders


def _warn(order, msg):
    if msg not in order["warnings"]:
        order["warnings"].append(msg)


def write_review_xlsx(orders, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "核对表"
    hdr = ["单号", "客户", "日期", "货品代码", "品名规格", "群里原话", "数量", "单价", "金额", "赠品", "备注", "需确认"]
    ws.append(hdr)
    for c in ws[1]:
        c.font = Font(bold=True)
    yellow = PatternFill("solid", fgColor="FFF2CC")
    red = PatternFill("solid", fgColor="F8CBAD")
    for i, o in enumerate(orders, 1):
        for j, it in enumerate(o["items"]):
            row = [i, o["customer"], o["date"], it.get("code") or "", it.get("name", ""), it["raw_name"],
                   it["qty"], it.get("price"), it.get("amount"), "是" if it.get("is_gift") else "",
                   o["remark"] if j == 0 else "", "；".join(o["warnings"]) if j == 0 else ""]
            ws.append(row)
            fill = red if not it.get("code") or not o.get("customer_in_list", True) else (
                yellow if o["warnings"] else None)
            if fill:
                for c in ws[ws.max_row]:
                    c.fill = fill
    for col, w in zip("ABCDEFGHIJKL", [5, 16, 11, 9, 26, 22, 6, 8, 9, 5, 28, 46]):
        ws.column_dimensions[col].width = w
    for r in ws.iter_rows(min_row=2):
        r[11].alignment = Alignment(wrap_text=True)
    wb.save(path)


def write_import_xls(order, path):
    """按智慧记「销售资料导入模板.xls」4 列格式，一张单一个文件。"""
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet("Sheet1")
    for c, h in enumerate(["货品代码", "品名规格（必填）", "数量", "原价"]):
        ws.write(0, c, h)
    r = 1
    for it in order["items"]:
        if not it.get("code"):
            continue
        for c, v in enumerate([it["code"], it["name"], it["qty"], it["price"]]):
            ws.write(r, c, v)
        r += 1
    wb.add_sheet("Sheet2")
    wb.add_sheet("Sheet3")
    wb.save(str(path))


def score(orders, expected_path):
    exp = json.loads(Path(expected_path).read_text(encoding="utf-8"))
    got = {o["customer"]: o for o in orders}
    total = ok = 0
    lines = []
    for e in exp:
        g = got.get(e["customer"])
        if not g:
            lines.append(f"  ✗ {e['customer']}: 没解析出来")
            total += len(e["items"])
            continue
        gi = sorted((it.get("code"), it["qty"], 0 if it.get("is_gift") else -1) for it in g["items"] if it.get("code"))
        for ei in e["items"]:
            total += 1
            key = (ei["code"], ei["qty"], 0 if ei["price"] == 0 else -1)
            if key in gi:
                ok += 1
                gi.remove(key)
            else:
                lines.append(f"  ✗ {e['customer']}: 期望 {ei['code']}×{ei['qty']}，实际 "
                             f"{[(it.get('code'), it['qty']) for it in g['items']]}")
        # 显式给了价格的行要核对单价
        for ei in e["items"]:
            if ei["price"] not in (None, 0):
                hit = [it for it in g["items"] if it.get("code") == ei["code"] and not it.get("is_gift")]
                if hit and hit[0].get("price") != ei["price"]:
                    lines.append(f"  ✗ {e['customer']}: {ei['code']} 单价应为 {ei['price']}，实际 {hit[0].get('price')}")
    extra = set(got) - {e["customer"] for e in exp}
    for x in extra:
        lines.append(f"  ✗ 多出来的单：{x}")
    print(f"\n准确率（货品行）：{ok}/{total}")
    print("\n".join(lines) if lines else "  全部正确")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("chat", help="复制/转发出来的群聊文本文件")
    ap.add_argument("--db", help="智慧记 data.db 路径（只读）")
    ap.add_argument("--out", default=str(HERE / "output"))
    args = ap.parse_args()

    products, customers = load_catalog_db(args.db) if args.db else load_catalog_xls()
    chat = Path(args.chat).read_text(encoding="utf-8")
    print(f"货品 {len(products)} 个，客户 {len(customers)} 个，消息 {len(chat)} 字，调用 {MODEL} 解析中…")

    orders = postprocess(ask_llm(chat, products, customers), products, customers)

    out = Path(args.out)
    (out / "导入文件").mkdir(parents=True, exist_ok=True)
    for f in (out / "导入文件").glob("*.xls"):
        f.unlink()
    (out / "解析结果.json").write_text(json.dumps(orders, ensure_ascii=False, indent=2), encoding="utf-8")
    write_review_xlsx(orders, out / "核对表.xlsx")
    for i, o in enumerate(orders, 1):
        if o.get("customer_in_list", True):
            safe = re.sub(r'[\\/:*?"<>|]', "_", o["customer"])
            write_import_xls(o, out / "导入文件" / f"{i:02d}_{safe}_{o['date']}.xls")

    for i, o in enumerate(orders, 1):
        flag = " ⚠" if o["warnings"] else ""
        print(f"\n[{i}] {o['customer']}  {o['date']}  合计 {o['total']}{flag}")
        for it in o["items"]:
            src = f"（{it['price_from']}）" if it.get("price_from") else ""
            print(f"    {it.get('code') or '??????':6} {it.get('name') or it['raw_name']:24} "
                  f"×{it['qty']:<4} 单价 {it.get('price')}{src}{'  [赠品]' if it.get('is_gift') else ''}")
        if o["remark"]:
            print(f"    备注：{o['remark']}")
        for w in o["warnings"]:
            print(f"    ⚠ {w}")

    exp = HERE / "4_标准答案.json"
    if exp.exists() and not args.db:
        score(orders, exp)
    print(f"\n输出：{out}")


if __name__ == "__main__":
    sys.exit(main())
