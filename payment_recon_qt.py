# -*- coding: utf-8 -*-
import csv
import json
import os
import sqlite3
import sys
import traceback
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

try:
    from openpyxl import load_workbook
except Exception:
    load_workbook = None


APP_TITLE = "财务第三方支付核对"
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DB_PATH = DATA_DIR / "reconciliation.sqlite3"

SUMMARY_COLUMNS = [
    "店铺", "月份", "订单实付应结", "平台补贴", "商家补贴", "结算运费", "订单退款",
    "收入净额合计", "已结算佣金", "技术服务费", "支出金额", "结算金额", "提现金额", "期初金额", "结算期末余额",
]

RAW_COLUMNS = [
    "店铺", "月份", "动账时间", "动账流水号", "动账方向", "动账账户", "动账金额",
    "动账摘要", "业务类型", "主订单编号", "子订单编号", "售后单号", "下单时间",
    "商品信息", "商品编码", "售卖类型", "订单实付应结", "平台补贴", "商家补贴",
    "结算运费", "订单退款", "佣金", "技术服务费",
]

ADJUSTABLE_COLUMNS = [
    "备查（不影响汇总）", "订单实付应结", "平台补贴", "商家补贴", "结算运费", "订单退款",
    "已结算佣金", "技术服务费", "提现金额", "期初金额",
]

SUMMARY_EXTRA_COLUMNS = ["店铺期末余额", "差异", "差异原因", "调整说明"]

DEFAULT_FORMULA_NOTE = (
    "收入净额合计=订单实付应结+平台补贴+商家补贴+结算运费+订单退款\n"
    "支出金额=已结算佣金+技术服务费\n"
    "结算金额=收入净额合计+支出金额\n"
    "结算期末余额=期初金额+收入净额合计+支出金额+提现金额"
)

FIELD_KEYS = {
    "订单实付应结": "paid_settlement",
    "平台补贴": "platform_subsidy",
    "商家补贴": "merchant_subsidy",
    "结算运费": "freight",
    "订单退款": "refund",
    "收入净额合计": "income_total",
    "已结算佣金": "commission",
    "技术服务费": "tech_fee",
    "支出金额": "expense_amount",
    "结算金额": "settlement_amount",
    "提现金额": "withdraw_amount",
    "期初金额": "opening_balance",
    "结算期末余额": "ending_balance",
}


def money(value) -> Decimal:
    if value is None or value == "":
        return Decimal("0")
    try:
        return Decimal(str(value).replace(",", "").strip()).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        return Decimal("0")


def money_text(value) -> str:
    return f"{money(value):,.2f}"


def db_float(value) -> float:
    return float(money(value))


def normalize_header(value) -> str:
    return str(value or "").strip().replace("\n", "").replace("\r", "")


HEADER_ALIASES = {
    "店铺": ["店铺"],
    "月份": ["动账时间1", "动帐时间1", "月份", "账期", "帐期"],
    "动账时间": ["动账时间", "动帐时间", "交易时间", "发生时间"],
    "动账流水号": ["动账流水号", "动帐流水号", "流水号", "资金流水号"],
    "动账方向": ["动账方向", "动帐方向", "收支方向"],
    "动账账户": ["动账账户", "动帐账户", "账户"],
    "动账金额": ["动账金额", "动帐金额", "金额"],
    "动账摘要": ["动账摘要", "动帐摘要", "动账场景", "动帐场景", "计费类型", "摘要"],
    "业务类型": ["业务类型", "订单类型", "计费类型"],
    "主订单编号": ["主订单编号", "订单号"],
    "子订单编号": ["子订单编号", "子订单号"],
    "售后单号": ["售后单号", "售后编号"],
    "下单时间": ["下单时间"],
    "商品信息": ["商品信息", "商品名称"],
    "商品编码": ["商品编码", "商品ID"],
    "售卖类型": ["售卖类型", "订单类型"],
}

AMOUNT_ALIASES = {
    "订单实付应结": ["订单实付应结"],
    "平台补贴": ["平台补贴", "实际平台补贴_运费", "实际平台补贴", "其他平台补贴", "政府补贴平台垫资"],
    "商家补贴": ["商家补贴", "实际达人补贴", "实际抖音支付补贴", "实际抖音月付营销补贴", "银行补贴", "以旧换新抵扣"],
    "结算运费": ["结算运费", "运费实付"],
    "订单退款": ["订单退款"],
    "佣金": ["佣金", "服务商佣金", "渠道分成", "招商服务费", "站外推广费", "其他分成"],
    "技术服务费": ["技术服务费", "平台服务费"],
}


def difference_reason(row) -> str:
    if row["account_ending"] is None:
        return "未录入店铺期末余额"
    if row["difference"] == Decimal("0.00"):
        return "一致"
    direction = "结算期末余额大于店铺期末余额" if row["difference"] > 0 else "结算期末余额小于店铺期末余额"
    return f"{direction}；请核对收入净额合计、支出金额、提现金额、期初金额或店铺期末余额录入"


def parse_month_sort(label, transaction_time=None) -> str:
    label = str(label or "").strip()
    digits = "".join(ch for ch in label if ch.isdigit())
    if "月" in label and digits:
        year = "0000"
        if transaction_time:
            text = str(transaction_time)
            if len(text) >= 4 and text[:4].isdigit():
                year = text[:4]
        return f"{year}-{int(digits):02d}"
    if len(digits) == 6:
        return f"{digits[:4]}-{digits[4:]}"
    if transaction_time:
        text = str(transaction_time)
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d"):
            try:
                return datetime.strptime(text[:19], fmt).strftime("%Y-%m")
            except ValueError:
                pass
        if len(text) >= 7 and text[4] in "-/":
            return text[:7].replace("/", "-")
    return label


def normalize_month_token(value) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    digits = "".join(ch for ch in text if ch.isdigit())
    if len(digits) >= 6:
        return f"{digits[:4]}-{digits[4:6]}"
    if len(text) >= 7 and text[4] in "-/":
        return text[:7].replace("/", "-")
    return text


def parse_month_filter(value):
    text = str(value or "").strip()
    if not text:
        return []
    text = text.replace("，", ",").replace("、", ",").replace("至", "-").replace("到", "-")
    if "," in text:
        return [m for m in (normalize_month_token(part) for part in text.split(",")) if m]
    if "-" in text:
        compact = text.replace(" ", "")
        digits = "".join(ch for ch in compact if ch.isdigit())
        if len(digits) == 12:
            start = normalize_month_token(digits[:6])
            end = normalize_month_token(digits[6:])
            return month_range(start, end)
    month = normalize_month_token(text)
    return [month] if month else []


def month_range(start, end):
    start = normalize_month_token(start)
    end = normalize_month_token(end)
    if not start or not end:
        return []
    sy, sm = [int(x) for x in start.split("-")]
    ey, em = [int(x) for x in end.split("-")]
    months = []
    y, m = sy, sm
    while (y, m) <= (ey, em):
        months.append(f"{y:04d}-{m:02d}")
        m += 1
        if m == 13:
            y += 1
            m = 1
    return months


def json_default(value):
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, Decimal):
        return str(value)
    return str(value)


def parse_formula_lines(text):
    formulas = []
    normalized = str(text or "").replace("；", "\n").replace(";", "\n")
    for line in normalized.splitlines():
        line = line.strip().rstrip("。")
        if not line or "=" not in line:
            continue
        left, right = line.split("=", 1)
        left = left.strip()
        right = right.strip()
        if left in FIELD_KEYS and right:
            formulas.append((left, right))
    return formulas


def eval_money_formula(expr, values):
    names = sorted(FIELD_KEYS.keys(), key=len, reverse=True)
    local_vars = {}
    safe_expr = str(expr)
    for idx, name in enumerate(names):
        var = f"v{idx}"
        local_vars[var] = money(values.get(FIELD_KEYS[name], 0))
        safe_expr = safe_expr.replace(name, var)
    allowed = set("0123456789.+-*/() _v")
    if any(ch not in allowed for ch in safe_expr):
        raise ValueError(f"公式包含不支持的字符：{expr}")
    return money(eval(safe_expr, {"__builtins__": {}}, local_vars))


@dataclass
class ImportResult:
    imported: int
    skipped: int
    file_name: str


class Repository:
    def __init__(self, path: Path):
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self.init_db()

    def init_db(self):
        cur = self.conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS stores (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                note TEXT,
                active INTEGER DEFAULT 1,
                created_at TEXT NOT NULL
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS transactions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                source_file TEXT,
                imported_at TEXT,
                store TEXT NOT NULL,
                month_label TEXT NOT NULL,
                month_sort TEXT NOT NULL,
                transaction_time TEXT,
                flow_id TEXT,
                direction TEXT,
                account TEXT,
                amount REAL DEFAULT 0,
                summary TEXT,
                biz_type TEXT,
                main_order TEXT,
                sub_order TEXT,
                after_sale TEXT,
                order_time TEXT,
                product_info TEXT,
                product_code TEXT,
                sale_type TEXT,
                paid_settlement REAL DEFAULT 0,
                platform_subsidy REAL DEFAULT 0,
                merchant_subsidy REAL DEFAULT 0,
                freight REAL DEFAULT 0,
                refund REAL DEFAULT 0,
                commission REAL DEFAULT 0,
                tech_fee REAL DEFAULT 0,
                raw_payload TEXT,
                UNIQUE(store, month_sort, flow_id, sub_order, summary, amount)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS store_configs (
                store TEXT PRIMARY KEY,
                raw_columns TEXT,
                summary_columns TEXT,
                formula_note TEXT,
                updated_at TEXT NOT NULL
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS balances (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                store TEXT NOT NULL,
                month_label TEXT NOT NULL,
                month_sort TEXT NOT NULL,
                opening_balance REAL DEFAULT 0,
                account_ending_balance REAL,
                note TEXT,
                updated_at TEXT NOT NULL,
                UNIQUE(store, month_sort)
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS adjustments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                store TEXT NOT NULL,
                month_label TEXT NOT NULL,
                month_sort TEXT NOT NULL,
                target_column TEXT DEFAULT '备查（不影响汇总）',
                item TEXT NOT NULL,
                amount REAL DEFAULT 0,
                note TEXT,
                updated_at TEXT NOT NULL
            )
        """)
        adj_cols = {row["name"] for row in cur.execute("PRAGMA table_info(adjustments)").fetchall()}
        if "target_column" not in adj_cols:
            cur.execute("ALTER TABLE adjustments ADD COLUMN target_column TEXT DEFAULT '备查（不影响汇总）'")
        tx_cols = {row["name"] for row in cur.execute("PRAGMA table_info(transactions)").fetchall()}
        if "raw_payload" not in tx_cols:
            cur.execute("ALTER TABLE transactions ADD COLUMN raw_payload TEXT")
        cur.execute("""
            INSERT OR IGNORE INTO stores (name, note, active, created_at)
            SELECT DISTINCT store, '', 1, ?
            FROM transactions
            WHERE store IS NOT NULL AND TRIM(store) <> ''
        """, (datetime.now().strftime("%Y-%m-%d %H:%M:%S"),))
        cur.execute("""
            INSERT OR IGNORE INTO store_configs (store, raw_columns, summary_columns, formula_note, updated_at)
            SELECT name, ?, ?, ?, ? FROM stores
        """, (
            json.dumps(RAW_COLUMNS, ensure_ascii=False),
            json.dumps(SUMMARY_COLUMNS + SUMMARY_EXTRA_COLUMNS, ensure_ascii=False),
            DEFAULT_FORMULA_NOTE,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        ))
        self.conn.commit()

    def import_excel(self, path: str, selected_store: str = "") -> ImportResult:
        if load_workbook is None:
            raise RuntimeError("缺少 openpyxl，无法读取 Excel。请使用启动脚本附带的 Python 运行。")

        workbook = load_workbook(path, read_only=True, data_only=True)
        imported = 0
        skipped = 0
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cur = self.conn.cursor()

        for ws in workbook.worksheets:
            rows = ws.iter_rows(values_only=True)
            try:
                headers = [normalize_header(v) for v in next(rows)]
            except StopIteration:
                continue
            idx = {name: i for i, name in enumerate(headers)}
            if selected_store:
                self.merge_store_raw_columns(selected_store, headers)

            def cell(row, name):
                pos = idx.get(name)
                if pos is None or pos >= len(row):
                    return None
                return row[pos]

            def aliased(row, name):
                for candidate in HEADER_ALIASES.get(name, [name]):
                    value = cell(row, candidate)
                    if value not in (None, ""):
                        return value
                return None

            def amount_sum(row, name):
                total = Decimal("0.00")
                found = False
                for candidate in AMOUNT_ALIASES.get(name, [name]):
                    if candidate in idx:
                        total += money(cell(row, candidate))
                        found = True
                return total if found else money(cell(row, name))

            has_store = selected_store or any(name in idx for name in HEADER_ALIASES["店铺"])
            has_month_or_time = any(name in idx for name in HEADER_ALIASES["月份"] + HEADER_ALIASES["动账时间"])
            if not has_store or not has_month_or_time:
                missing = []
                if not has_store:
                    missing.append("店铺")
                if not has_month_or_time:
                    missing.append("月份或动账时间")
                raise ValueError(f"工作表 {ws.title} 无法自动识别关键列：{', '.join(missing)}")

            for row in rows:
                store = selected_store or str(aliased(row, "店铺") or "").strip()
                if store:
                    self.ensure_store_config(store)
                    self.merge_store_raw_columns(store, headers)
                transaction_time = aliased(row, "动账时间")
                month_label = str(aliased(row, "月份") or "").strip()
                month_sort = parse_month_sort(month_label, transaction_time)
                if not month_label:
                    month_label = month_sort
                flow_id = str(aliased(row, "动账流水号") or "").strip()
                if not flow_id:
                    flow_id = f"{Path(path).name}:{ws.title}:{imported + skipped + 2}"
                if not store or not month_sort:
                    skipped += 1
                    continue
                raw_payload = {
                    headers[i]: json_default(row[i]) if i < len(row) else ""
                    for i in range(len(headers))
                    if headers[i]
                }
                values = (
                    Path(path).name, now, store, month_label, month_sort,
                    str(transaction_time or ""), flow_id,
                    str(aliased(row, "动账方向") or ""), str(aliased(row, "动账账户") or ""),
                    db_float(aliased(row, "动账金额")), str(aliased(row, "动账摘要") or ""),
                    str(aliased(row, "业务类型") or ""), str(aliased(row, "主订单编号") or ""),
                    str(aliased(row, "子订单编号") or ""), str(aliased(row, "售后单号") or ""),
                    str(aliased(row, "下单时间") or ""), str(aliased(row, "商品信息") or ""),
                    str(aliased(row, "商品编码") or ""), str(aliased(row, "售卖类型") or ""),
                    db_float(amount_sum(row, "订单实付应结")), db_float(amount_sum(row, "平台补贴")),
                    db_float(amount_sum(row, "商家补贴")), db_float(amount_sum(row, "结算运费")),
                    db_float(amount_sum(row, "订单退款")), db_float(amount_sum(row, "佣金")),
                    db_float(amount_sum(row, "技术服务费")),
                    json.dumps(raw_payload, ensure_ascii=False, default=json_default),
                )
                try:
                    cur.execute("""
                        INSERT INTO transactions (
                            source_file, imported_at, store, month_label, month_sort,
                            transaction_time, flow_id, direction, account, amount,
                            summary, biz_type, main_order, sub_order, after_sale, order_time,
                            product_info, product_code, sale_type, paid_settlement,
                            platform_subsidy, merchant_subsidy, freight, refund,
                            commission, tech_fee, raw_payload
                        ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    """, values)
                    imported += 1
                except sqlite3.IntegrityError:
                    skipped += 1

        self.conn.commit()
        return ImportResult(imported=imported, skipped=skipped, file_name=Path(path).name)

    def configured_stores(self, include_inactive=False):
        if include_inactive:
            rows = self.conn.execute("SELECT * FROM stores ORDER BY active DESC, name").fetchall()
        else:
            rows = self.conn.execute("SELECT * FROM stores WHERE active=1 ORDER BY name").fetchall()
        return rows

    def add_store(self, name, note=""):
        name = str(name or "").strip()
        if not name:
            raise ValueError("请输入店铺名称")
        self.conn.execute("""
            INSERT INTO stores (name, note, active, created_at)
            VALUES (?, ?, 1, ?)
            ON CONFLICT(name) DO UPDATE SET note=excluded.note, active=1
        """, (name, note, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        self.ensure_store_config(name)
        self.conn.commit()

    def delete_store(self, name):
        self.conn.execute("UPDATE stores SET active=0 WHERE name=?", (name,))
        self.conn.commit()

    def stores(self):
        rows = self.conn.execute("""
            SELECT name AS store FROM stores WHERE active=1
            UNION
            SELECT DISTINCT store FROM transactions
            UNION
            SELECT DISTINCT store FROM balances
            ORDER BY store
        """).fetchall()
        return [r[0] for r in rows if r[0]]

    def ensure_store_config(self, store):
        if not store:
            return
        self.conn.execute("""
            INSERT OR IGNORE INTO store_configs (store, raw_columns, summary_columns, formula_note, updated_at)
            VALUES (?, ?, ?, ?, ?)
        """, (
            store,
            json.dumps(RAW_COLUMNS, ensure_ascii=False),
            json.dumps(SUMMARY_COLUMNS + SUMMARY_EXTRA_COLUMNS, ensure_ascii=False),
            DEFAULT_FORMULA_NOTE,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        ))

    def store_config(self, store):
        self.ensure_store_config(store)
        row = self.conn.execute("SELECT * FROM store_configs WHERE store=?", (store,)).fetchone()
        if not row:
            return {
                "store": store,
                "raw_columns": RAW_COLUMNS,
                "summary_columns": SUMMARY_COLUMNS + SUMMARY_EXTRA_COLUMNS,
                "formula_note": DEFAULT_FORMULA_NOTE,
            }
        return {
            "store": store,
            "raw_columns": self._loads_columns(row["raw_columns"], RAW_COLUMNS),
            "summary_columns": self._loads_columns(row["summary_columns"], SUMMARY_COLUMNS + SUMMARY_EXTRA_COLUMNS),
            "formula_note": row["formula_note"] or DEFAULT_FORMULA_NOTE,
        }

    def _loads_columns(self, text, fallback):
        try:
            values = json.loads(text or "[]")
            return [str(v).strip() for v in values if str(v).strip()] or list(fallback)
        except Exception:
            return list(fallback)

    def save_store_config(self, store, raw_columns, summary_columns, formula_note):
        self.ensure_store_config(store)
        self.conn.execute("""
            UPDATE store_configs
            SET raw_columns=?, summary_columns=?, formula_note=?, updated_at=?
            WHERE store=?
        """, (
            json.dumps(raw_columns, ensure_ascii=False),
            json.dumps(summary_columns, ensure_ascii=False),
            formula_note,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            store,
        ))
        self.conn.commit()

    def merge_store_raw_columns(self, store, headers):
        if not store:
            return
        config = self.store_config(store)
        columns = list(config["raw_columns"])
        changed = False
        for header in headers:
            header = str(header or "").strip()
            if header and header not in columns:
                columns.append(header)
                changed = True
        if changed:
            self.save_store_config(store, columns, config["summary_columns"], config["formula_note"])

    def months(self):
        rows = self.conn.execute("SELECT DISTINCT month_sort, month_label FROM transactions UNION SELECT DISTINCT month_sort, month_label FROM balances ORDER BY month_sort").fetchall()
        return [{"month_sort": r["month_sort"], "month_label": r["month_label"]} for r in rows]

    def upsert_balance(self, store, month_label, month_sort, opening, account_ending, note):
        self.conn.execute("""
            INSERT INTO balances (store, month_label, month_sort, opening_balance, account_ending_balance, note, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(store, month_sort) DO UPDATE SET
                month_label=excluded.month_label,
                opening_balance=excluded.opening_balance,
                account_ending_balance=excluded.account_ending_balance,
                note=excluded.note,
                updated_at=excluded.updated_at
        """, (store, month_label, month_sort, db_float(opening), db_float(account_ending), note, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        self.conn.commit()

    def delete_month(self, store, month_sort, delete_balance=False):
        self.conn.execute("DELETE FROM transactions WHERE store=? AND month_sort=?", (store, month_sort))
        self.conn.execute("DELETE FROM adjustments WHERE store=? AND month_sort=?", (store, month_sort))
        if delete_balance:
            self.conn.execute("DELETE FROM balances WHERE store=? AND month_sort=?", (store, month_sort))
        self.conn.commit()

    def transaction_by_id(self, tx_id):
        return self.conn.execute("SELECT * FROM transactions WHERE id=?", (tx_id,)).fetchone()

    def update_transaction(self, tx_id, data):
        self.conn.execute("""
            UPDATE transactions SET
                store=?, month_label=?, month_sort=?, transaction_time=?,
                direction=?, account=?, amount=?, summary=?, biz_type=?
            WHERE id=?
        """, (
            data["店铺"], data["月份显示"], data["月份排序(YYYY-MM)"], data["动账时间"],
            data["动账方向"], data["动账账户"], db_float(data["动账金额"]),
            data["动账摘要"], data["业务类型"], tx_id,
        ))
        self.conn.commit()

    def delete_transaction(self, tx_id):
        self.conn.execute("DELETE FROM transactions WHERE id=?", (tx_id,))
        self.conn.commit()

    def add_adjustment(self, store, month_label, month_sort, target_column, item, amount, note):
        self.conn.execute("""
            INSERT INTO adjustments (store, month_label, month_sort, target_column, item, amount, note, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (store, month_label, month_sort, target_column, item, db_float(amount), note, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        self.conn.commit()

    def delete_adjustment(self, adj_id):
        self.conn.execute("DELETE FROM adjustments WHERE id=?", (adj_id,))
        self.conn.commit()

    def monthly_summaries(self, store_filter="", month_filter="", aggregate=True):
        months = parse_month_filter(month_filter)
        tx_rows = self._monthly_base_rows(store_filter, months)
        month_rows = [self._build_month_summary(row) for row in tx_rows]
        month_rows.extend(self._balance_only_rows(store_filter, months))
        month_rows = sorted(month_rows, key=lambda x: (x["store"], x["month_sort"]))
        if not aggregate or len(months) <= 1:
            return month_rows
        return self._aggregate_rows(month_rows, months)

    def _month_where(self, where, params, months):
        if months:
            where.append("month_sort IN (%s)" % ",".join("?" for _ in months))
            params.extend(months)

    def _monthly_base_rows(self, store_filter="", months=None):
        months = months or []
        params = []
        where = []
        if store_filter:
            where.append("store LIKE ?")
            params.append(f"%{store_filter}%")
        self._month_where(where, params, months)
        where_sql = "WHERE " + " AND ".join(where) if where else ""
        return self.conn.execute(f"""
            SELECT store, month_label, month_sort,
                   SUM(CASE WHEN summary='账户余额提现' THEN amount ELSE 0 END) AS withdraw_amount,
                   SUM(paid_settlement) AS paid_settlement,
                   SUM(platform_subsidy) AS platform_subsidy,
                   SUM(merchant_subsidy) AS merchant_subsidy,
                   SUM(freight) AS freight,
                   SUM(refund) AS refund,
                   SUM(paid_settlement + platform_subsidy + merchant_subsidy + freight + refund) AS income_total,
                   SUM(commission) AS commission,
                   SUM(tech_fee) AS tech_fee,
                   SUM(commission + tech_fee) AS expense_amount,
                   SUM(paid_settlement + platform_subsidy + merchant_subsidy + freight + refund + commission + tech_fee) AS settlement_amount
            FROM transactions
            {where_sql}
            GROUP BY store, month_sort
        """, params).fetchall()

    def _build_month_summary(self, row):
        balance = self.balance_for(row["store"], row["month_sort"])
        adjustments_by_target = self.adjustment_totals_by_target(row["store"], row["month_sort"])
        paid_settlement = money(row["paid_settlement"]) + adjustments_by_target.get("订单实付应结", Decimal("0.00"))
        platform_subsidy = money(row["platform_subsidy"]) + adjustments_by_target.get("平台补贴", Decimal("0.00"))
        merchant_subsidy = money(row["merchant_subsidy"]) + adjustments_by_target.get("商家补贴", Decimal("0.00"))
        freight = money(row["freight"]) + adjustments_by_target.get("结算运费", Decimal("0.00"))
        refund = money(row["refund"]) + adjustments_by_target.get("订单退款", Decimal("0.00"))
        commission = money(row["commission"]) + adjustments_by_target.get("已结算佣金", Decimal("0.00"))
        tech_fee = money(row["tech_fee"]) + adjustments_by_target.get("技术服务费", Decimal("0.00"))
        withdraw_amount = money(row["withdraw_amount"]) + adjustments_by_target.get("提现金额", Decimal("0.00"))
        opening = money(balance["opening_balance"] if balance else 0) + adjustments_by_target.get("期初金额", Decimal("0.00"))
        income_total = paid_settlement + platform_subsidy + merchant_subsidy + freight + refund
        expense_amount = commission + tech_fee
        settlement_amount = income_total + expense_amount
        ending = opening + income_total + expense_amount + withdraw_amount
        account_ending = None if not balance or balance["account_ending_balance"] is None else money(balance["account_ending_balance"])
        diff = None if account_ending is None else ending - account_ending
        summary = {
            "store": row["store"], "month_label": row["month_label"], "month_sort": row["month_sort"],
            "month_sorts": [row["month_sort"]], "is_aggregate": False,
            "paid_settlement": paid_settlement, "platform_subsidy": platform_subsidy,
            "merchant_subsidy": merchant_subsidy, "freight": freight, "refund": refund,
            "income_total": income_total, "commission": commission, "tech_fee": tech_fee,
            "expense_amount": expense_amount, "settlement_amount": settlement_amount,
            "withdraw_amount": withdraw_amount, "opening_balance": opening,
            "adjustments": self.adjustment_total(row["store"], row["month_sort"]),
            "adjustments_by_target": adjustments_by_target,
            "adjustment_notes": self.adjustment_notes(row["store"], [row["month_sort"]]),
            "ending_balance": ending, "account_ending": account_ending, "difference": diff,
        }
        self.apply_store_formulas(row["store"], summary)
        summary["difference"] = None if summary["account_ending"] is None else summary["ending_balance"] - summary["account_ending"]
        return summary

    def _balance_only_rows(self, store_filter="", months=None):
        months = months or []
        result = []
        rows = self.conn.execute("""
            SELECT b.* FROM balances b
            LEFT JOIN transactions t ON t.store=b.store AND t.month_sort=b.month_sort
            WHERE t.id IS NULL
            ORDER BY b.store, b.month_sort
        """).fetchall()
        for row in rows:
            if store_filter and store_filter not in row["store"]:
                continue
            if months and row["month_sort"] not in months:
                continue
            opening = money(row["opening_balance"])
            adjustments_by_target = self.adjustment_totals_by_target(row["store"], row["month_sort"])
            paid_settlement = adjustments_by_target.get("订单实付应结", Decimal("0.00"))
            platform_subsidy = adjustments_by_target.get("平台补贴", Decimal("0.00"))
            merchant_subsidy = adjustments_by_target.get("商家补贴", Decimal("0.00"))
            freight = adjustments_by_target.get("结算运费", Decimal("0.00"))
            refund = adjustments_by_target.get("订单退款", Decimal("0.00"))
            commission = adjustments_by_target.get("已结算佣金", Decimal("0.00"))
            tech_fee = adjustments_by_target.get("技术服务费", Decimal("0.00"))
            withdraw_amount = adjustments_by_target.get("提现金额", Decimal("0.00"))
            opening = opening + adjustments_by_target.get("期初金额", Decimal("0.00"))
            income_total = paid_settlement + platform_subsidy + merchant_subsidy + freight + refund
            expense_amount = commission + tech_fee
            settlement_amount = income_total + expense_amount
            ending = opening + income_total + expense_amount + withdraw_amount
            account_ending = None if row["account_ending_balance"] is None else money(row["account_ending_balance"])
            summary = {
                "store": row["store"], "month_label": row["month_label"], "month_sort": row["month_sort"],
                "month_sorts": [row["month_sort"]], "is_aggregate": False,
                "paid_settlement": paid_settlement, "platform_subsidy": platform_subsidy, "merchant_subsidy": merchant_subsidy,
                "freight": freight, "refund": refund, "income_total": income_total,
                "commission": commission, "tech_fee": tech_fee, "expense_amount": expense_amount,
                "settlement_amount": settlement_amount, "withdraw_amount": withdraw_amount,
                "opening_balance": opening, "adjustments": self.adjustment_total(row["store"], row["month_sort"]),
                "adjustments_by_target": adjustments_by_target,
                "adjustment_notes": self.adjustment_notes(row["store"], [row["month_sort"]]),
                "ending_balance": ending, "account_ending": account_ending,
                "difference": None if account_ending is None else ending - account_ending,
            }
            self.apply_store_formulas(row["store"], summary)
            summary["difference"] = None if summary["account_ending"] is None else summary["ending_balance"] - summary["account_ending"]
            result.append(summary)
        return result

    def _aggregate_rows(self, rows, months):
        result = []
        by_store = {}
        for row in rows:
            by_store.setdefault(row["store"], []).append(row)
        numeric_sum_fields = [
            "paid_settlement", "platform_subsidy", "merchant_subsidy", "freight", "refund",
            "income_total", "commission", "tech_fee", "expense_amount", "settlement_amount",
            "withdraw_amount", "adjustments",
        ]
        for store, store_rows in by_store.items():
            store_rows = sorted(store_rows, key=lambda r: r["month_sort"])
            first_month = store_rows[0]["month_sort"]
            last_month = store_rows[-1]["month_sort"]
            first_balance = self.balance_for(store, first_month)
            last_balance = self.balance_for(store, last_month)
            aggregate = {
                "store": store,
                "month_label": f"{months[0]}至{months[-1]}",
                "month_sort": f"{months[0]}~{months[-1]}",
                "month_sorts": [r["month_sort"] for r in store_rows],
                "is_aggregate": True,
                "opening_balance": money(first_balance["opening_balance"] if first_balance else store_rows[0]["opening_balance"]),
                "account_ending": None if not last_balance or last_balance["account_ending_balance"] is None else money(last_balance["account_ending_balance"]),
                "adjustment_notes": self.adjustment_notes(store, [r["month_sort"] for r in store_rows]),
            }
            for field in numeric_sum_fields:
                aggregate[field] = sum((money(r[field]) for r in store_rows), Decimal("0.00"))
            aggregate["ending_balance"] = (
                aggregate["opening_balance"] + aggregate["income_total"] +
                aggregate["expense_amount"] + aggregate["withdraw_amount"]
            )
            self.apply_store_formulas(store, aggregate)
            aggregate["difference"] = None if aggregate["account_ending"] is None else aggregate["ending_balance"] - aggregate["account_ending"]
            result.append(aggregate)
            result.extend(store_rows)
        return result

    def apply_store_formulas(self, store, summary):
        formulas = parse_formula_lines(self.store_config(store)["formula_note"])
        for target, expr in formulas:
            key = FIELD_KEYS[target]
            try:
                summary[key] = eval_money_formula(expr, summary)
            except Exception:
                continue

    def balance_for(self, store, month_sort):
        return self.conn.execute("SELECT * FROM balances WHERE store=? AND month_sort=?", (store, month_sort)).fetchone()

    def adjustment_total(self, store, month_sort) -> Decimal:
        row = self.conn.execute("SELECT SUM(amount) AS total FROM adjustments WHERE store=? AND month_sort=?", (store, month_sort)).fetchone()
        return money(row["total"] if row else 0)

    def adjustment_totals_by_target(self, store, month_sort):
        rows = self.conn.execute("""
            SELECT COALESCE(target_column, '备查（不影响汇总）') AS target_column, SUM(amount) AS total
            FROM adjustments
            WHERE store=? AND month_sort=?
            GROUP BY COALESCE(target_column, '备查（不影响汇总）')
        """, (store, month_sort)).fetchall()
        return {row["target_column"]: money(row["total"]) for row in rows}

    def adjustment_notes(self, store, months):
        if not months:
            return ""
        rows = self.conn.execute(f"""
            SELECT month_sort, target_column, item, amount, note
            FROM adjustments
            WHERE store=? AND month_sort IN ({",".join("?" for _ in months)})
            ORDER BY month_sort, id
        """, [store] + list(months)).fetchall()
        return "；".join(
            f"{r['month_sort']} [{r['target_column'] or '备查'}] {r['item']} {money_text(r['amount'])} {r['note'] or ''}".strip()
            for r in rows
        )

    def details(self, store, month_sort=None, months=None, sort_order="ASC"):
        months = months or ([month_sort] if month_sort else [])
        if not store or not months:
            return []
        order = "DESC" if str(sort_order).upper() == "DESC" else "ASC"
        return self.conn.execute(f"""
            SELECT * FROM transactions WHERE store=? AND month_sort IN ({",".join("?" for _ in months)})
            ORDER BY transaction_time {order}, id {order}
        """, [store] + list(months)).fetchall()

    def raw_rows_for_export(self, store_filter="", month_filter=""):
        months = parse_month_filter(month_filter)
        params = []
        where = []
        if store_filter:
            where.append("store LIKE ?")
            params.append(f"%{store_filter}%")
        self._month_where(where, params, months)
        where_sql = "WHERE " + " AND ".join(where) if where else ""
        rows = self.conn.execute(f"""
            SELECT * FROM transactions {where_sql}
            ORDER BY store, month_sort, transaction_time, id
        """, params).fetchall()
        columns = []
        payloads = []
        for row in rows:
            try:
                payload = json.loads(row["raw_payload"] or "{}")
            except Exception:
                payload = {}
            if not payload:
                payload = {col: row[col] if col in row.keys() else "" for col in RAW_COLUMNS}
            for key in payload.keys():
                if key not in columns:
                    columns.append(key)
            payloads.append((row, payload))
        return columns, payloads

    def difference_groups(self, store, month_sort=None, months=None):
        months = months or ([month_sort] if month_sort else [])
        rows = self.conn.execute("""
            SELECT direction, summary, COUNT(*) AS count, SUM(amount) AS amount,
                   SUM(paid_settlement) AS paid_settlement,
                   SUM(platform_subsidy) AS platform_subsidy,
                   SUM(merchant_subsidy) AS merchant_subsidy,
                   SUM(freight) AS freight,
                   SUM(refund) AS refund,
                   SUM(paid_settlement + platform_subsidy + merchant_subsidy + freight + refund) AS income_total,
                   SUM(commission) AS commission,
                   SUM(tech_fee) AS tech_fee,
                   SUM(commission + tech_fee) AS expense_amount
            FROM transactions
            WHERE store=? AND month_sort IN (%s)
            GROUP BY direction, summary
            ORDER BY direction, summary
        """ % ",".join("?" for _ in months), [store] + list(months)).fetchall()
        adjs = self.conn.execute(f"""
            SELECT * FROM adjustments WHERE store=? AND month_sort IN ({",".join("?" for _ in months)})
            ORDER BY month_sort, id
        """, [store] + list(months)).fetchall()
        return rows, adjs


class BalanceDialog(simpledialog.Dialog):
    def __init__(self, parent, title, initial=None):
        self.initial = initial or {}
        self.result = None
        super().__init__(parent, title)

    def body(self, master):
        labels = ["店铺", "月份显示", "月份排序(YYYY-MM)", "期初金额", "店铺期末余额", "备注"]
        self.entries = {}
        for i, label in enumerate(labels):
            ttk.Label(master, text=label).grid(row=i, column=0, sticky="e", padx=8, pady=6)
            entry = ttk.Entry(master, width=34)
            entry.grid(row=i, column=1, sticky="ew", padx=8, pady=6)
            self.entries[label] = entry
        values = {
            "店铺": self.initial.get("store", ""),
            "月份显示": self.initial.get("month_label", ""),
            "月份排序(YYYY-MM)": self.initial.get("month_sort", ""),
            "期初金额": self.initial.get("opening_balance", "0"),
            "店铺期末余额": self.initial.get("account_ending", "0"),
            "备注": self.initial.get("note", ""),
        }
        for label, value in values.items():
            self.entries[label].insert(0, "" if value is None else str(value))
        return self.entries["店铺"]

    def validate(self):
        try:
            if not self.entries["店铺"].get().strip():
                raise ValueError("请输入店铺")
            if not self.entries["月份排序(YYYY-MM)"].get().strip():
                raise ValueError("请输入月份排序，例如 2026-06")
            money(self.entries["期初金额"].get())
            money(self.entries["店铺期末余额"].get())
            return True
        except Exception as exc:
            messagebox.showerror("输入有误", str(exc), parent=self)
            return False

    def apply(self):
        self.result = {
            "store": self.entries["店铺"].get().strip(),
            "month_label": self.entries["月份显示"].get().strip() or self.entries["月份排序(YYYY-MM)"].get().strip(),
            "month_sort": self.entries["月份排序(YYYY-MM)"].get().strip(),
            "opening_balance": self.entries["期初金额"].get().strip(),
            "account_ending": self.entries["店铺期末余额"].get().strip(),
            "note": self.entries["备注"].get().strip(),
        }


class AdjustmentDialog(simpledialog.Dialog):
    def __init__(self, parent, title, initial):
        self.initial = initial
        self.result = None
        super().__init__(parent, title)

    def body(self, master):
        ttk.Label(master, text="调整列").grid(row=0, column=0, sticky="e", padx=8, pady=6)
        ttk.Label(master, text="调整事项").grid(row=1, column=0, sticky="e", padx=8, pady=6)
        ttk.Label(master, text="金额").grid(row=2, column=0, sticky="e", padx=8, pady=6)
        ttk.Label(master, text="说明").grid(row=3, column=0, sticky="e", padx=8, pady=6)
        self.target_var = tk.StringVar(value=ADJUSTABLE_COLUMNS[0])
        self.target = ttk.Combobox(master, textvariable=self.target_var, values=ADJUSTABLE_COLUMNS, width=32, state="readonly")
        self.item = ttk.Entry(master, width=34)
        self.amount = ttk.Entry(master, width=34)
        self.note = ttk.Entry(master, width=34)
        self.target.grid(row=0, column=1, padx=8, pady=6)
        self.item.grid(row=1, column=1, padx=8, pady=6)
        self.amount.grid(row=2, column=1, padx=8, pady=6)
        self.note.grid(row=3, column=1, padx=8, pady=6)
        self.item.insert(0, "手工差异调整")
        self.amount.insert(0, "0")
        return self.target

    def validate(self):
        try:
            if not self.item.get().strip():
                raise ValueError("请输入调整事项")
            money(self.amount.get())
            return True
        except Exception as exc:
            messagebox.showerror("输入有误", str(exc), parent=self)
            return False

    def apply(self):
        self.result = {
            "target_column": self.target_var.get().strip(),
            "item": self.item.get().strip(),
            "amount": self.amount.get().strip(),
            "note": self.note.get().strip(),
        }


class TransactionDialog(simpledialog.Dialog):
    def __init__(self, parent, title, initial):
        self.initial = initial
        self.result = None
        super().__init__(parent, title)

    def body(self, master):
        fields = ["店铺", "月份显示", "月份排序(YYYY-MM)", "动账时间", "动账方向", "动账账户", "动账金额", "动账摘要", "业务类型"]
        self.entries = {}
        for i, field in enumerate(fields):
            ttk.Label(master, text=field).grid(row=i, column=0, sticky="e", padx=8, pady=5)
            entry = ttk.Entry(master, width=42)
            entry.grid(row=i, column=1, sticky="ew", padx=8, pady=5)
            entry.insert(0, str(self.initial.get(field, "") or ""))
            self.entries[field] = entry
        return self.entries["动账金额"]

    def validate(self):
        try:
            if not self.entries["店铺"].get().strip():
                raise ValueError("请输入店铺")
            if not self.entries["月份排序(YYYY-MM)"].get().strip():
                raise ValueError("请输入月份排序")
            money(self.entries["动账金额"].get())
            return True
        except Exception as exc:
            messagebox.showerror("输入有误", str(exc), parent=self)
            return False

    def apply(self):
        self.result = {key: entry.get().strip() for key, entry in self.entries.items()}


class ImportStoreDialog(simpledialog.Dialog):
    def __init__(self, parent, stores):
        self.stores = stores
        self.result = None
        super().__init__(parent, "选择导入店铺")

    def body(self, master):
        ttk.Label(master, text="导入归属店铺").grid(row=0, column=0, sticky="e", padx=8, pady=8)
        self.store_var = tk.StringVar()
        self.store_box = ttk.Combobox(master, textvariable=self.store_var, values=self.stores, width=34, state="readonly")
        self.store_box.grid(row=0, column=1, sticky="ew", padx=8, pady=8)
        if self.stores:
            self.store_box.current(0)
        ttk.Label(master, text="说明：选择后，本次导入的流水全部按该店铺处理；Excel 没有“店铺”列也可以导入。").grid(
            row=1, column=0, columnspan=2, sticky="w", padx=8, pady=(0, 8)
        )
        return self.store_box

    def validate(self):
        if not self.store_var.get().strip():
            messagebox.showerror("请选择店铺", "请先选择一个已配置店铺。", parent=self)
            return False
        return True

    def apply(self):
        self.result = self.store_var.get().strip()


class StoreManageDialog(tk.Toplevel):
    def __init__(self, parent, repo, on_changed):
        super().__init__(parent)
        self.repo = repo
        self.on_changed = on_changed
        self.title("店铺配置")
        self.geometry("520x420")
        self.transient(parent)
        self.create_widgets()
        self.refresh()

    def create_widgets(self):
        form = ttk.Frame(self, padding=10)
        form.pack(fill="x")
        ttk.Label(form, text="店铺名称").grid(row=0, column=0, padx=4, pady=6, sticky="e")
        self.name_entry = ttk.Entry(form, width=28)
        self.name_entry.grid(row=0, column=1, padx=4, pady=6, sticky="ew")
        ttk.Label(form, text="备注").grid(row=1, column=0, padx=4, pady=6, sticky="e")
        self.note_entry = ttk.Entry(form, width=28)
        self.note_entry.grid(row=1, column=1, padx=4, pady=6, sticky="ew")
        ttk.Button(form, text="新增/启用店铺", command=self.add_store).grid(row=0, column=2, rowspan=2, padx=8, pady=6)
        form.columnconfigure(1, weight=1)

        self.tree = ttk.Treeview(self, columns=["店铺", "备注", "状态"], show="headings")
        self.tree.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        for col in ["店铺", "备注", "状态"]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor="center")

        actions = ttk.Frame(self, padding=(10, 0, 10, 10))
        actions.pack(fill="x")
        ttk.Button(actions, text="停用选中店铺", command=self.delete_store).pack(side="left", padx=4)
        ttk.Button(actions, text="关闭", command=self.destroy).pack(side="right", padx=4)

    def refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for row in self.repo.configured_stores(include_inactive=True):
            status = "启用" if row["active"] else "停用"
            self.tree.insert("", "end", values=[row["name"], row["note"] or "", status], iid=row["name"])

    def add_store(self):
        try:
            self.repo.add_store(self.name_entry.get(), self.note_entry.get())
            self.name_entry.delete(0, "end")
            self.note_entry.delete(0, "end")
            self.refresh()
            self.on_changed()
        except Exception as exc:
            messagebox.showerror("保存失败", str(exc), parent=self)

    def delete_store(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("未选择", "请先选择一个店铺。", parent=self)
            return
        name = selection[0]
        if not messagebox.askyesno("确认停用", f"确定停用店铺“{name}”吗？历史流水不会删除。", parent=self):
            return
        self.repo.delete_store(name)
        self.refresh()
        self.on_changed()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1280x760")
        self.minsize(1080, 680)
        self.repo = Repository(DB_PATH)
        self.selected_summary = None
        self.create_widgets()
        self.refresh_all()

    def create_widgets(self):
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        self.configure(bg="#f5f7fb")
        style.configure(".", font=("Microsoft YaHei UI", 9))
        style.configure("TFrame", background="#f5f7fb")
        style.configure("TLabel", background="#f5f7fb", foreground="#243042")
        style.configure("App.TFrame", background="#f5f7fb")
        style.configure("Header.TFrame", background="#17233c")
        style.configure("Surface.TLabelframe", background="#f5f7fb", bordercolor="#d8dee9", relief="solid")
        style.configure("Surface.TLabelframe.Label", background="#f5f7fb", foreground="#283447", font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("Treeview", rowheight=28, background="#ffffff", fieldbackground="#ffffff", foreground="#243042", bordercolor="#d8dee9")
        style.configure("Treeview.Heading", background="#eef2f7", foreground="#1f2a3d", font=("Microsoft YaHei UI", 9, "bold"), relief="flat")
        style.map("Treeview", background=[("selected", "#dbeafe")], foreground=[("selected", "#0f172a")])
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 15, "bold"), background="#17233c", foreground="#ffffff")
        style.configure("Subtitle.TLabel", font=("Microsoft YaHei UI", 9), background="#17233c", foreground="#cbd5e1")
        style.configure("Muted.TLabel", background="#f5f7fb", foreground="#64748b")
        style.configure("Primary.TButton", font=("Microsoft YaHei UI", 9, "bold"), padding=(12, 7), background="#2563eb", foreground="#ffffff")
        style.map("Primary.TButton", background=[("active", "#1d4ed8")], foreground=[("disabled", "#dbeafe")])
        style.configure("Action.TButton", padding=(10, 6), background="#ffffff", foreground="#1f2a3d")
        style.configure("Danger.TButton", padding=(10, 6), background="#fff1f2", foreground="#be123c")
        style.configure("TEntry", padding=(4, 3), fieldbackground="#ffffff")

        top = ttk.Frame(self, padding=(16, 12), style="Header.TFrame")
        top.pack(fill="x")
        title_box = ttk.Frame(top, style="Header.TFrame")
        title_box.pack(side="left", fill="x", expand=True)
        ttk.Label(title_box, text=APP_TITLE, style="Title.TLabel").pack(anchor="w")
        ttk.Label(title_box, text="第三方支付流水导入、月度汇总、余额核对与差异追踪", style="Subtitle.TLabel").pack(anchor="w", pady=(2, 0))
        ttk.Button(top, text="导入Excel流水", command=self.import_excel, style="Primary.TButton").pack(side="right", padx=(8, 0))
        ttk.Button(top, text="店铺配置", command=self.manage_stores, style="Action.TButton").pack(side="right", padx=4)
        ttk.Button(top, text="录入/修改余额", command=self.edit_balance, style="Action.TButton").pack(side="right", padx=4)
        ttk.Button(top, text="导出汇总CSV", command=self.export_summary, style="Action.TButton").pack(side="right", padx=4)

        filters = ttk.LabelFrame(self, text="查询条件", style="Surface.TLabelframe")
        filters.pack(fill="x", padx=14, pady=(12, 8))
        ttk.Label(filters, text="店铺").pack(side="left", padx=(10, 4), pady=8)
        self.store_filter = ttk.Entry(filters, width=20)
        self.store_filter.pack(side="left", padx=4)
        ttk.Label(filters, text="月份").pack(side="left", padx=(12, 4))
        self.month_filter = ttk.Entry(filters, width=16)
        self.month_filter.pack(side="left", padx=4)
        ttk.Button(filters, text="查询", command=self.refresh_all, style="Primary.TButton").pack(side="left", padx=6)
        ttk.Button(filters, text="清空", command=self.clear_filters, style="Action.TButton").pack(side="left", padx=4)
        self.status = ttk.Label(filters, text=f"数据库：{DB_PATH}", style="Muted.TLabel")
        self.status.pack(side="right", padx=10)

        panes = ttk.PanedWindow(self, orient="vertical")
        panes.pack(fill="both", expand=True, padx=14, pady=(0, 12))

        summary_frame = ttk.LabelFrame(panes, text="按店铺月份汇总", style="Surface.TLabelframe")
        panes.add(summary_frame, weight=2)
        self.summary_frozen_tree, self.summary_tree = self.make_summary_tree(summary_frame, SUMMARY_COLUMNS, ["店铺期末余额", "差异", "差异原因"], height=10)
        self.summary_frozen_tree.bind("<<TreeviewSelect>>", self.on_summary_select)
        self.summary_frozen_tree.bind("<Double-1>", lambda _e: self.edit_balance())
        self.summary_tree.bind("<<TreeviewSelect>>", self.on_summary_select)
        self.summary_tree.bind("<Double-1>", lambda _e: self.edit_balance())

        actions = ttk.Frame(summary_frame)
        actions.pack(fill="x", padx=8, pady=6)
        ttk.Button(actions, text="查看差异明细", command=self.show_difference, style="Action.TButton").pack(side="left", padx=4)
        ttk.Button(actions, text="新增调整", command=self.add_adjustment, style="Primary.TButton").pack(side="left", padx=4)
        ttk.Button(actions, text="删除选中月份流水", command=self.delete_month, style="Danger.TButton").pack(side="left", padx=4)

        detail_frame = ttk.LabelFrame(panes, text="明细流水", style="Surface.TLabelframe")
        panes.add(detail_frame, weight=3)
        self.detail_tree = self.make_tree(detail_frame, ["ID"] + RAW_COLUMNS, height=14)
        detail_actions = ttk.Frame(detail_frame)
        detail_actions.pack(fill="x", padx=8, pady=(0, 6))
        ttk.Button(detail_actions, text="编辑选中流水", command=self.edit_transaction, style="Action.TButton").pack(side="left", padx=4)
        ttk.Button(detail_actions, text="删除选中流水", command=self.delete_transaction, style="Danger.TButton").pack(side="left", padx=4)

        bottom = ttk.Frame(self, padding=(14, 0, 14, 10), style="App.TFrame")
        bottom.pack(fill="x")
        ttk.Label(bottom, text="计算口径：收入净额合计=订单实付应结+平台补贴+商家补贴+结算运费+订单退款；支出金额=已结算佣金+技术服务费；结算期末余额=期初金额+收入净额合计+支出金额+提现金额。", style="Muted.TLabel").pack(side="left")

    def make_tree(self, parent, columns, height):
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True, padx=8, pady=8)
        tree = ttk.Treeview(frame, columns=columns, show="headings", height=height)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        for col in columns:
            tree.heading(col, text=col)
            width = 130
            if col in ("商品信息",):
                width = 220
            if col in ("ID", "月份"):
                width = 70
            tree.column(col, width=width, minwidth=60, anchor="center")
        return tree

    def make_summary_tree(self, parent, base_columns, extra_columns, height):
        frozen_columns = base_columns[:2]
        scroll_columns = base_columns[2:] + extra_columns
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True, padx=8, pady=8)
        frozen = ttk.Treeview(frame, columns=frozen_columns, show="headings", height=height, selectmode="browse")
        scroll = ttk.Treeview(frame, columns=scroll_columns, show="headings", height=height, selectmode="browse")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=lambda *args: self.sync_summary_yview(frozen, scroll, *args))
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=scroll.xview)
        frozen.configure(yscrollcommand=lambda first, last: vsb.set(first, last))
        scroll.configure(yscrollcommand=lambda first, last: vsb.set(first, last), xscrollcommand=hsb.set)
        frozen.grid(row=0, column=0, sticky="nsw")
        scroll.grid(row=0, column=1, sticky="nsew")
        vsb.grid(row=0, column=2, sticky="ns")
        hsb.grid(row=1, column=1, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        for col in frozen_columns:
            frozen.heading(col, text=col)
            frozen.column(col, width=90, minwidth=70, anchor="center", stretch=False)
        for col in scroll_columns:
            scroll.heading(col, text=col)
            scroll.column(col, width=130, minwidth=80, anchor="center")
        frozen.bind("<MouseWheel>", lambda event: self.on_summary_mousewheel(frozen, scroll, event))
        scroll.bind("<MouseWheel>", lambda event: self.on_summary_mousewheel(frozen, scroll, event))
        return frozen, scroll

    def sync_summary_yview(self, frozen, scroll, *args):
        frozen.yview(*args)
        scroll.yview(*args)

    def on_summary_mousewheel(self, frozen, scroll, event):
        units = -1 if event.delta > 0 else 1
        frozen.yview_scroll(units, "units")
        scroll.yview_scroll(units, "units")
        return "break"

    def auto_fit_tree_columns(self, tree, max_width=260):
        for col in tree["columns"]:
            values = [str(tree.set(item, col)) for item in tree.get_children("")]
            longest = max([str(col)] + values, key=lambda value: len(value))
            width = min(max(len(longest) * 9 + 24, 70), max_width)
            if col in ("差异原因", "商品信息"):
                width = min(max(width, 180), 360)
            tree.column(col, width=width)

    def refresh_all(self):
        self.refresh_summary()
        self.refresh_details(None)

    def clear_filters(self):
        self.store_filter.delete(0, "end")
        self.month_filter.delete(0, "end")
        self.refresh_all()

    def manage_stores(self):
        StoreManageDialog(self, self.repo, self.refresh_all)

    def import_excel(self):
        store_names = [row["name"] for row in self.repo.configured_stores()]
        if not store_names:
            if messagebox.askyesno("需要配置店铺", "还没有配置店铺。是否现在添加店铺？"):
                self.manage_stores()
            return
        dialog = ImportStoreDialog(self, store_names)
        if not dialog.result:
            return
        selected_store = dialog.result
        paths = filedialog.askopenfilenames(
            title="选择资金流水 Excel",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm"), ("所有文件", "*.*")]
        )
        if not paths:
            return
        try:
            total_imported = total_skipped = 0
            for path in paths:
                result = self.repo.import_excel(path, selected_store=selected_store)
                total_imported += result.imported
                total_skipped += result.skipped
            self.refresh_all()
            messagebox.showinfo("导入完成", f"店铺：{selected_store}\n新增 {total_imported} 条，跳过重复/空行 {total_skipped} 条。")
        except Exception as exc:
            messagebox.showerror("导入失败", f"{exc}\n\n{traceback.format_exc(limit=2)}")

    def refresh_summary(self):
        for item in self.summary_frozen_tree.get_children():
            self.summary_frozen_tree.delete(item)
        for item in self.summary_tree.get_children():
            self.summary_tree.delete(item)
        rows = self.repo.monthly_summaries(self.store_filter.get().strip(), self.month_filter.get().strip())
        for row in rows:
            frozen_values = [row["store"], row["month_label"]]
            scroll_values = [
                money_text(row["paid_settlement"]),
                money_text(row["platform_subsidy"]), money_text(row["merchant_subsidy"]),
                money_text(row["freight"]), money_text(row["refund"]), money_text(row["income_total"]),
                money_text(row["commission"]), money_text(row["tech_fee"]),
                money_text(row["expense_amount"]), money_text(row["settlement_amount"]),
                money_text(row["withdraw_amount"]), money_text(row["opening_balance"]), money_text(row["ending_balance"]),
                "" if row["account_ending"] is None else money_text(row["account_ending"]),
                "" if row["difference"] is None else money_text(row["difference"]),
                difference_reason(row),
            ]
            tags = ("diff",) if row["difference"] is not None and row["difference"] != Decimal("0.00") else ()
            iid = f"{row['store']}|{row['month_sort']}"
            self.summary_frozen_tree.insert("", "end", values=frozen_values, tags=tags, iid=iid)
            self.summary_tree.insert("", "end", values=scroll_values, tags=tags, iid=iid)
        self.summary_frozen_tree.tag_configure("diff", background="#fff1f0")
        self.summary_tree.tag_configure("diff", background="#fff1f0")
        self.auto_fit_tree_columns(self.summary_frozen_tree, max_width=160)
        self.auto_fit_tree_columns(self.summary_tree)
        self.status.configure(text=f"数据库：{DB_PATH}    汇总 {len(rows)} 条")

    def refresh_details(self, key):
        for item in self.detail_tree.get_children():
            self.detail_tree.delete(item)
        if not key:
            return
        store, month_sort = key
        for row in self.repo.details(store, month_sort):
            values = [
                row["id"], row["store"], row["month_label"], row["transaction_time"], row["flow_id"],
                row["direction"], row["account"], money_text(row["amount"]), row["summary"],
                row["biz_type"], row["main_order"], row["sub_order"], row["after_sale"],
                row["order_time"], row["product_info"], row["product_code"], row["sale_type"],
                money_text(row["paid_settlement"]), money_text(row["platform_subsidy"]),
                money_text(row["merchant_subsidy"]), money_text(row["freight"]),
                money_text(row["refund"]), money_text(row["commission"]), money_text(row["tech_fee"]),
            ]
            self.detail_tree.insert("", "end", values=values)
        self.auto_fit_tree_columns(self.detail_tree, max_width=280)

    def current_key(self):
        selection = self.summary_tree.selection() or self.summary_frozen_tree.selection()
        if not selection:
            return None
        iid = selection[0]
        if "|" not in iid:
            return None
        return tuple(iid.split("|", 1))

    def selected_row_dict(self):
        key = self.current_key()
        if not key:
            return None
        for row in self.repo.monthly_summaries():
            if row["store"] == key[0] and row["month_sort"] == key[1]:
                return row
        return None

    def on_summary_select(self, _event=None):
        key = self.current_key()
        if key:
            iid = "|".join(key)
            if self.summary_tree.exists(iid):
                self.summary_tree.selection_set(iid)
            if self.summary_frozen_tree.exists(iid):
                self.summary_frozen_tree.selection_set(iid)
        self.refresh_details(key)

    def edit_balance(self):
        row = self.selected_row_dict()
        initial = {}
        if row:
            bal = self.repo.balance_for(row["store"], row["month_sort"])
            initial = {
                "store": row["store"],
                "month_label": row["month_label"],
                "month_sort": row["month_sort"],
                "opening_balance": row["opening_balance"],
                "account_ending": row["account_ending"] if row["account_ending"] is not None else "0",
                "note": bal["note"] if bal else "",
            }
        dialog = BalanceDialog(self, "录入/修改余额", initial)
        if dialog.result:
            self.repo.upsert_balance(
                dialog.result["store"],
                dialog.result["month_label"],
                dialog.result["month_sort"],
                dialog.result["opening_balance"],
                dialog.result["account_ending"],
                dialog.result["note"],
            )
            self.refresh_all()
            messagebox.showinfo("保存成功", "期初金额和店铺期末余额已保存。")

    def add_adjustment(self):
        row = self.selected_row_dict()
        if not row:
            messagebox.showwarning("未选择", "请先选择一个店铺月份。")
            return
        dialog = AdjustmentDialog(self, "新增差异调整", row)
        if dialog.result:
            self.repo.add_adjustment(row["store"], row["month_label"], row["month_sort"], dialog.result["target_column"], dialog.result["item"], dialog.result["amount"], dialog.result["note"])
            self.refresh_all()

    def current_transaction_id(self):
        selection = self.detail_tree.selection()
        if not selection:
            return None
        values = self.detail_tree.item(selection[0], "values")
        if not values:
            return None
        try:
            return int(values[0])
        except (TypeError, ValueError):
            return None

    def edit_transaction(self):
        tx_id = self.current_transaction_id()
        if not tx_id:
            messagebox.showwarning("未选择", "请先在明细流水中选择一条记录。")
            return
        tx = self.repo.transaction_by_id(tx_id)
        if not tx:
            messagebox.showerror("未找到", "这条流水记录不存在，可能已被删除。")
            self.refresh_all()
            return
        initial = {
            "店铺": tx["store"],
            "月份显示": tx["month_label"],
            "月份排序(YYYY-MM)": tx["month_sort"],
            "动账时间": tx["transaction_time"],
            "动账方向": tx["direction"],
            "动账账户": tx["account"],
            "动账金额": money_text(tx["amount"]).replace(",", ""),
            "动账摘要": tx["summary"],
            "业务类型": tx["biz_type"],
        }
        dialog = TransactionDialog(self, "编辑流水", initial)
        if dialog.result:
            self.repo.update_transaction(tx_id, dialog.result)
            self.refresh_all()

    def delete_transaction(self):
        tx_id = self.current_transaction_id()
        if not tx_id:
            messagebox.showwarning("未选择", "请先在明细流水中选择一条记录。")
            return
        if not messagebox.askyesno("确认删除", f"确定删除流水 ID {tx_id} 吗？"):
            return
        self.repo.delete_transaction(tx_id)
        self.refresh_all()

    def show_difference(self):
        row = self.selected_row_dict()
        if not row:
            messagebox.showwarning("未选择", "请先选择一个店铺月份。")
            return
        rows, adjs = self.repo.difference_groups(row["store"], row["month_sort"])
        win = tk.Toplevel(self)
        win.title(f"差异明细 - {row['store']} {row['month_label']}")
        win.geometry("920x560")
        text = tk.Text(win, wrap="none", font=("Consolas", 10))
        text.pack(fill="both", expand=True)
        text.insert("end", f"店铺：{row['store']}    月份：{row['month_label']} ({row['month_sort']})\n")
        text.insert("end", f"订单实付应结：{money_text(row['paid_settlement'])}\n")
        text.insert("end", f"平台补贴：{money_text(row['platform_subsidy'])}\n")
        text.insert("end", f"商家补贴：{money_text(row['merchant_subsidy'])}\n")
        text.insert("end", f"结算运费：{money_text(row['freight'])}\n")
        text.insert("end", f"订单退款：{money_text(row['refund'])}\n")
        text.insert("end", f"收入净额合计：{money_text(row['income_total'])}\n")
        text.insert("end", f"已结算佣金：{money_text(row['commission'])}\n")
        text.insert("end", f"技术服务费：{money_text(row['tech_fee'])}\n")
        text.insert("end", f"支出金额：{money_text(row['expense_amount'])}\n")
        text.insert("end", f"结算金额：{money_text(row['settlement_amount'])}\n")
        text.insert("end", f"提现金额：{money_text(row['withdraw_amount'])}\n")
        text.insert("end", f"期初金额：{money_text(row['opening_balance'])}\n")
        text.insert("end", f"结算期末余额：{money_text(row['ending_balance'])}\n")
        text.insert("end", f"店铺期末余额：{'' if row['account_ending'] is None else money_text(row['account_ending'])}\n")
        text.insert("end", f"差异：{'' if row['difference'] is None else money_text(row['difference'])}\n\n")
        text.insert("end", "差异原因说明：\n")
        text.insert("end", f"{difference_reason(row)}\n")
        text.insert("end", "核对公式：差异 = 结算期末余额 - 店铺期末余额；结算期末余额 = 期初金额 + 收入净额合计 + 支出金额 + 提现金额。\n\n")
        text.insert("end", "按动账方向/摘要汇总：\n")
        text.insert("end", "方向\t摘要\t笔数\t动账金额\t订单实付应结\t平台补贴\t商家补贴\t结算运费\t订单退款\t收入净额合计\t已结算佣金\t技术服务费\t支出金额\n")
        for r in rows:
            text.insert("end", f"{r['direction']}\t{r['summary']}\t{r['count']}\t{money_text(r['amount'])}\t{money_text(r['paid_settlement'])}\t{money_text(r['platform_subsidy'])}\t{money_text(r['merchant_subsidy'])}\t{money_text(r['freight'])}\t{money_text(r['refund'])}\t{money_text(r['income_total'])}\t{money_text(r['commission'])}\t{money_text(r['tech_fee'])}\t{money_text(r['expense_amount'])}\n")
        text.insert("end", "\n手工调整记录（选择具体列会参与汇总；选择备查不影响汇总）：\n")
        if adjs:
            for a in adjs:
                text.insert("end", f"#{a['id']} [{a['target_column'] or '备查（不影响汇总）'}] {a['item']}\t{money_text(a['amount'])}\t{a['note'] or ''}\n")
        else:
            text.insert("end", "无\n")
        text.configure(state="disabled")

    def delete_month(self):
        row = self.selected_row_dict()
        if not row:
            messagebox.showwarning("未选择", "请先选择要删除的店铺月份。")
            return
        if not messagebox.askyesno("确认删除", f"确定删除 {row['store']} {row['month_label']} 的全部流水和手工调整吗？余额记录默认保留。"):
            return
        self.repo.delete_month(row["store"], row["month_sort"], delete_balance=False)
        self.refresh_all()

    def export_summary(self):
        path = filedialog.asksaveasfilename(
            title="保存汇总CSV",
            defaultextension=".csv",
            filetypes=[("CSV 文件", "*.csv")]
        )
        if not path:
            return
        rows = self.repo.monthly_summaries(self.store_filter.get().strip(), self.month_filter.get().strip())
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(SUMMARY_COLUMNS + ["店铺期末余额", "差异", "差异原因"])
            for row in rows:
                writer.writerow([
                    row["store"], row["month_label"], row["paid_settlement"],
                    row["platform_subsidy"], row["merchant_subsidy"], row["freight"], row["refund"],
                    row["income_total"], row["commission"], row["tech_fee"], row["expense_amount"],
                    row["settlement_amount"], row["withdraw_amount"], row["opening_balance"], row["ending_balance"],
                    "" if row["account_ending"] is None else row["account_ending"],
                    "" if row["difference"] is None else row["difference"],
                    difference_reason(row),
                ])
        messagebox.showinfo("导出完成", f"已保存：{path}")


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
