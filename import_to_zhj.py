"""把 parse_orders.py 的解析结果写进金蝶智慧记（通过它自带的本地服务 zhjd，127.0.0.1:5166）。

走的是软件自己的开单接口（sale.chk 校验 → sale.i 保存），库存、应收款由智慧记自己计算，
不直接写数据库。

用法：
    python import_to_zhj.py output/解析结果.json            # 只校验，不写入
    python import_to_zhj.py output/解析结果.json --commit   # 校验通过后写入
"""
import argparse
import json
import os
import sqlite3
import urllib.request
from pathlib import Path

URL = "http://127.0.0.1:5166/zhjd"
SID = os.getenv("ZHJ_SID", "sid#pc#qweruiop412")   # 智慧记界面调本地服务时用的 sid，从它的运行日志里能看到
DB = os.getenv("ZHJ_DB", "")   # 智慧记安装目录下 zhjd/data.db 的路径，也可以用 --db 传


def call(cmd, **kw):
    body = json.dumps({"cmd": cmd, "sid": SID, "ver": 1, **kw}, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(URL, body, {"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=15) as r:
        return json.loads(r.read().decode("utf-8"))


def load_ids(db):
    # 只读，用来把货品代码/客户名换成智慧记内部 id
    con = sqlite3.connect(f"file:{db}?mode=ro", uri=True)
    products = {code: pid for pid, code in con.execute(
        "select id, code from products where is_del=0 and code<>''")}
    customers = {name: cid for cid, name in con.execute(
        "select id, name from companies where is_del=0 and is_sys=0")}
    uid, = con.execute("select uid from shopusers order by id limit 1").fetchone()
    return products, customers, uid


def build(order, products, customers, uid):
    cid = customers.get(order["customer"])
    if not cid:
        return None, f"客户「{order['customer']}」不在智慧记里"
    items, skipped = [], []
    for it in order["items"]:
        pid = products.get(it.get("code"))
        if not pid:
            skipped.append(it["raw_name"])
            continue
        prc = float(it["price"])
        qty = float(it["qty"])
        items.append({"product_id": pid, "utye": 1, "ratio": 0, "disc": 1,
                      "oprc": prc, "prc": prc, "qty": qty, "amt": round(prc * qty, 2),
                      "mprc": prc, "mqty": qty, "remark": "赠品" if it.get("is_gift") else ""})
    if not items:
        return None, "没有可录入的货品"
    total = round(sum(i["amt"] for i in items), 2)
    remark = order.get("remark", "")
    if skipped:
        remark = (remark + "；" if remark else "") + "未录：" + "、".join(skipped)
    code = call("bill.prep.code", bt="XSD", opt_on=order["date"])["code"]
    bill = {"id": 0, "is_update": 0, "tye": 1, "opt_on": order["date"], "code": code,
            "company_id": cid, "bill_amt": total, "disc_amt": total, "disc": 1,
            "pay_amt": 0, "owe_amt": total, "deduction_amt": 0, "express_amt": 0, "change_amt": 0,
            "acct_id": 0, "trade_src": 1, "trade_type": 1, "sorder_id": 0,
            "user_id": 1, "uid": uid, "username": "", "remark": remark[:200]}
    return {"bill": bill, "items": items}, ("未录货品：" + "、".join(skipped)) if skipped else ""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("orders_json")
    ap.add_argument("--commit", action="store_true", help="校验通过后真正写入")
    ap.add_argument("--db", default=DB)
    args = ap.parse_args()

    if not args.db:
        raise SystemExit("请用 --db 或环境变量 ZHJ_DB 指定智慧记的 data.db 路径")
    if call("sys.ping").get("msg") != "pong":
        raise SystemExit("智慧记没有运行，请先打开智慧记")
    orders = json.loads(Path(args.orders_json).read_text(encoding="utf-8"))
    products, customers, uid = load_ids(args.db)

    for o in orders:
        head = f"{o['customer']} {o['date']}"
        payload, note = build(o, products, customers, uid)
        if not payload:
            print(f"✗ 跳过 {head}：{note}")
            continue
        chk = call("sale.chk", **payload)
        if chk.get("rst") != 1:
            print(f"✗ 校验失败 {head}：{chk.get('msg')}")
            continue
        amt = payload["bill"]["bill_amt"]
        if not args.commit:
            print(f"✓ 校验通过 {head}  {len(payload['items'])} 行  合计 {amt}  {note}")
            continue
        r = call("sale.i", **payload)
        if r.get("rst") == 1:
            print(f"✓ 已录入 {head}  单号 {r['code']}  合计 {amt}  {note}")
        else:
            print(f"✗ 录入失败 {head}：{r.get('msg')}")


if __name__ == "__main__":
    main()
