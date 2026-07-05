# -*- coding: utf-8 -*-
import csv
import sys
import traceback
from decimal import Decimal
from pathlib import Path

LOCAL_PYSIDE = Path(__file__).resolve().parents[2] / "work" / "pyside6_pkg"
if LOCAL_PYSIDE.exists():
    sys.path.insert(0, str(LOCAL_PYSIDE))

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QColor, QFont
from PySide6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QToolBar,
    QVBoxLayout,
    QWidget,
)

from payment_recon_qt import (
    ADJUSTABLE_COLUMNS,
    DB_PATH,
    RAW_COLUMNS,
    SUMMARY_COLUMNS,
    Repository,
    difference_reason,
    money,
    money_text,
)


APP_TITLE = "财务第三方支付核对 - Qt版"


def table_item(value, align=Qt.AlignCenter):
    item = QTableWidgetItem(str(value))
    item.setTextAlignment(align)
    item.setFlags(item.flags() ^ Qt.ItemIsEditable)
    return item


class BalanceDialog(QDialog):
    def __init__(self, parent, initial):
        super().__init__(parent)
        self.setWindowTitle("录入/修改余额")
        self.resize(460, 280)
        self.result_data = None

        self.store = QLineEdit(str(initial.get("store", "")))
        self.month_label = QLineEdit(str(initial.get("month_label", "")))
        self.month_sort = QLineEdit(str(initial.get("month_sort", "")))
        self.opening = QLineEdit(str(initial.get("opening_balance", "0")))
        self.account_ending = QLineEdit(str(initial.get("account_ending", "0")))
        self.note = QLineEdit(str(initial.get("note", "")))

        form = QFormLayout()
        form.addRow("店铺", self.store)
        form.addRow("月份显示", self.month_label)
        form.addRow("月份排序(YYYY-MM)", self.month_sort)
        form.addRow("期初金额", self.opening)
        form.addRow("店铺期末余额", self.account_ending)
        form.addRow("备注", self.note)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(buttons)

    def accept(self):
        if not self.store.text().strip() or not self.month_sort.text().strip():
            QMessageBox.warning(self, "请补充信息", "店铺和月份排序不能为空。")
            return
        self.result_data = {
            "store": self.store.text().strip(),
            "month_label": self.month_label.text().strip() or self.month_sort.text().strip(),
            "month_sort": self.month_sort.text().strip(),
            "opening_balance": self.opening.text().strip() or "0",
            "account_ending": self.account_ending.text().strip() or "0",
            "note": self.note.text().strip(),
        }
        super().accept()


class AdjustmentDialog(QDialog):
    def __init__(self, parent, row):
        super().__init__(parent)
        self.setWindowTitle("新增差异调整")
        self.resize(480, 300)
        self.result_data = None

        self.target = QComboBox()
        self.target.addItems(ADJUSTABLE_COLUMNS)
        self.item = QLineEdit("差异调整")
        self.amount = QLineEdit("0")
        self.note = QLineEdit("")
        info = QLabel(f"{row['store']} / {row['month_label']}    当前差异：{'' if row['difference'] is None else money_text(row['difference'])}")
        info.setObjectName("mutedLabel")

        form = QFormLayout()
        form.addRow("调整列", self.target)
        form.addRow("调整事项", self.item)
        form.addRow("调整金额", self.amount)
        form.addRow("说明", self.note)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addWidget(info)
        layout.addLayout(form)
        layout.addWidget(buttons)

    def accept(self):
        if not self.item.text().strip():
            QMessageBox.warning(self, "请补充信息", "调整事项不能为空。")
            return
        self.result_data = {
            "target_column": self.target.currentText(),
            "item": self.item.text().strip(),
            "amount": self.amount.text().strip() or "0",
            "note": self.note.text().strip(),
        }
        super().accept()


class TransactionDialog(QDialog):
    FIELDS = ["店铺", "月份显示", "月份排序(YYYY-MM)", "动账时间", "动账方向", "动账账户", "动账金额", "动账摘要", "业务类型"]

    def __init__(self, parent, initial):
        super().__init__(parent)
        self.setWindowTitle("编辑流水")
        self.resize(520, 420)
        self.result_data = None
        self.edits = {}

        form = QFormLayout()
        for field in self.FIELDS:
            edit = QLineEdit(str(initial.get(field, "")))
            self.edits[field] = edit
            form.addRow(field, edit)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(buttons)

    def accept(self):
        if not self.edits["店铺"].text().strip() or not self.edits["月份排序(YYYY-MM)"].text().strip():
            QMessageBox.warning(self, "请补充信息", "店铺和月份排序不能为空。")
            return
        self.result_data = {field: edit.text().strip() for field, edit in self.edits.items()}
        super().accept()


class StoreDialog(QDialog):
    def __init__(self, parent, repo):
        super().__init__(parent)
        self.repo = repo
        self.setWindowTitle("店铺配置")
        self.resize(560, 420)

        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["店铺", "备注", "状态"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)

        add_btn = QPushButton("新增/启用店铺")
        del_btn = QPushButton("停用选中店铺")
        add_btn.clicked.connect(self.add_store)
        del_btn.clicked.connect(self.disable_store)

        bar = QHBoxLayout()
        bar.addWidget(add_btn)
        bar.addWidget(del_btn)
        bar.addStretch()

        layout = QVBoxLayout(self)
        layout.addWidget(self.table)
        layout.addLayout(bar)
        self.refresh()

    def refresh(self):
        rows = self.repo.configured_stores(include_inactive=True)
        self.table.setRowCount(len(rows))
        for r, row in enumerate(rows):
            self.table.setItem(r, 0, table_item(row["name"]))
            self.table.setItem(r, 1, table_item(row["note"] or "", Qt.AlignLeft | Qt.AlignVCenter))
            self.table.setItem(r, 2, table_item("启用" if row["active"] else "停用"))
        self.table.resizeColumnsToContents()

    def add_store(self):
        name, ok = QInputDialog.getText(self, "新增店铺", "店铺名称")
        if not ok or not name.strip():
            return
        note, _ = QInputDialog.getText(self, "店铺备注", "备注（可为空）")
        try:
            self.repo.add_store(name.strip(), note.strip())
            self.refresh()
        except Exception as exc:
            QMessageBox.critical(self, "保存失败", str(exc))

    def disable_store(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "未选择", "请先选择一个店铺。")
            return
        name = self.table.item(row, 0).text()
        if QMessageBox.question(self, "确认停用", f"确定停用店铺：{name}？") != QMessageBox.Yes:
            return
        self.repo.delete_store(name)
        self.refresh()


class DifferenceDialog(QDialog):
    def __init__(self, parent, repo, row):
        super().__init__(parent)
        self.setWindowTitle(f"差异明细 - {row['store']} {row['month_label']}")
        self.resize(1060, 680)

        groups, adjustments = repo.difference_groups(row["store"], row["month_sort"])
        text = QPlainTextEdit()
        text.setReadOnly(True)
        text.setFont(QFont("Consolas", 10))
        lines = [
            f"店铺：{row['store']}    月份：{row['month_label']} ({row['month_sort']})",
            f"订单实付应结：{money_text(row['paid_settlement'])}",
            f"平台补贴：{money_text(row['platform_subsidy'])}",
            f"商家补贴：{money_text(row['merchant_subsidy'])}",
            f"结算运费：{money_text(row['freight'])}",
            f"订单退款：{money_text(row['refund'])}",
            f"收入净额合计：{money_text(row['income_total'])}",
            f"已结算佣金：{money_text(row['commission'])}",
            f"技术服务费：{money_text(row['tech_fee'])}",
            f"支出金额：{money_text(row['expense_amount'])}",
            f"结算金额：{money_text(row['settlement_amount'])}",
            f"提现金额：{money_text(row['withdraw_amount'])}",
            f"期初金额：{money_text(row['opening_balance'])}",
            f"结算期末余额：{money_text(row['ending_balance'])}",
            f"店铺期末余额：{'' if row['account_ending'] is None else money_text(row['account_ending'])}",
            f"差异：{'' if row['difference'] is None else money_text(row['difference'])}",
            "",
            "差异原因说明：",
            difference_reason(row),
            "核对公式：差异 = 结算期末余额 - 店铺期末余额；结算期末余额 = 期初金额 + 收入净额合计 + 支出金额 + 提现金额。",
            "",
            "按动账方向/摘要汇总：",
            "方向\t摘要\t笔数\t动账金额\t订单实付应结\t平台补贴\t商家补贴\t结算运费\t订单退款\t收入净额合计\t已结算佣金\t技术服务费\t支出金额",
        ]
        for r in groups:
            lines.append(
                f"{r['direction']}\t{r['summary']}\t{r['count']}\t{money_text(r['amount'])}\t"
                f"{money_text(r['paid_settlement'])}\t{money_text(r['platform_subsidy'])}\t{money_text(r['merchant_subsidy'])}\t"
                f"{money_text(r['freight'])}\t{money_text(r['refund'])}\t{money_text(r['income_total'])}\t"
                f"{money_text(r['commission'])}\t{money_text(r['tech_fee'])}\t{money_text(r['expense_amount'])}"
            )
        lines.extend(["", "手工调整记录（选择具体列会参与汇总；选择备查不影响汇总）："])
        if adjustments:
            for a in adjustments:
                lines.append(f"#{a['id']} [{a['target_column'] or '备查（不影响汇总）'}] {a['item']}\t{money_text(a['amount'])}\t{a['note'] or ''}")
        else:
            lines.append("无")
        text.setPlainText("\n".join(lines))

        layout = QVBoxLayout(self)
        layout.addWidget(text)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.repo = Repository(DB_PATH)
        self.summary_rows = []
        self.current_summary_key = None
        self.setWindowTitle(APP_TITLE)
        self.resize(1480, 920)
        self.build_ui()
        self.apply_style()
        self.refresh_all()

    def build_ui(self):
        toolbar = QToolBar("主操作")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)
        for text, handler in [
            ("导入Excel流水", self.import_excel),
            ("店铺配置", self.manage_stores),
            ("录入/修改余额", self.edit_balance),
            ("新增调整", self.add_adjustment),
            ("查看差异明细", self.show_difference),
            ("导出汇总CSV", self.export_summary),
        ]:
            action = QAction(text, self)
            action.triggered.connect(handler)
            toolbar.addAction(action)

        root = QWidget()
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(14, 12, 14, 12)
        root_layout.setSpacing(10)

        title = QLabel("财务第三方支付核对")
        title.setObjectName("titleLabel")
        subtitle = QLabel("Qt 版桌面客户端 · SQLite 落盘 · 店铺/月度核对 · 手工调整留痕")
        subtitle.setObjectName("subtitleLabel")
        root_layout.addWidget(title)
        root_layout.addWidget(subtitle)

        filter_bar = QFrame()
        filter_bar.setObjectName("filterBar")
        filter_layout = QHBoxLayout(filter_bar)
        self.store_filter = QLineEdit()
        self.store_filter.setPlaceholderText("按店铺筛选")
        self.month_filter = QLineEdit()
        self.month_filter.setPlaceholderText("按月份筛选，如 2026-06")
        search_btn = QPushButton("查询")
        clear_btn = QPushButton("清空")
        search_btn.clicked.connect(self.refresh_all)
        clear_btn.clicked.connect(self.clear_filters)
        filter_layout.addWidget(QLabel("店铺"))
        filter_layout.addWidget(self.store_filter, 1)
        filter_layout.addWidget(QLabel("月份"))
        filter_layout.addWidget(self.month_filter, 1)
        filter_layout.addWidget(search_btn)
        filter_layout.addWidget(clear_btn)
        root_layout.addWidget(filter_bar)

        splitter = QSplitter(Qt.Vertical)
        root_layout.addWidget(splitter, 1)

        summary_widget = QWidget()
        summary_layout = QVBoxLayout(summary_widget)
        summary_layout.setContentsMargins(0, 0, 0, 0)
        summary_title = QLabel("按店铺月份汇总（冻结店铺、月份两列）")
        summary_title.setObjectName("sectionTitle")
        summary_layout.addWidget(summary_title)
        summary_tables = QHBoxLayout()
        self.summary_fixed = QTableWidget(0, 2)
        self.summary_scroll = QTableWidget(0, len(SUMMARY_COLUMNS[2:]) + 3)
        self.summary_fixed.setHorizontalHeaderLabels(SUMMARY_COLUMNS[:2])
        self.summary_scroll.setHorizontalHeaderLabels(SUMMARY_COLUMNS[2:] + ["店铺期末余额", "差异", "差异原因"])
        self.configure_table(self.summary_fixed)
        self.configure_table(self.summary_scroll)
        self.summary_fixed.setFixedWidth(220)
        self.summary_fixed.verticalScrollBar().valueChanged.connect(self.summary_scroll.verticalScrollBar().setValue)
        self.summary_scroll.verticalScrollBar().valueChanged.connect(self.summary_fixed.verticalScrollBar().setValue)
        self.summary_fixed.itemSelectionChanged.connect(self.on_fixed_selection)
        self.summary_scroll.itemSelectionChanged.connect(self.on_scroll_selection)
        self.summary_fixed.cellDoubleClicked.connect(lambda _r, _c: self.edit_balance())
        self.summary_scroll.cellDoubleClicked.connect(lambda _r, _c: self.edit_balance())
        summary_tables.addWidget(self.summary_fixed)
        summary_tables.addWidget(self.summary_scroll, 1)
        summary_layout.addLayout(summary_tables, 1)

        detail_widget = QWidget()
        detail_layout = QVBoxLayout(detail_widget)
        detail_layout.setContentsMargins(0, 0, 0, 0)
        detail_title = QLabel("明细流水")
        detail_title.setObjectName("sectionTitle")
        detail_layout.addWidget(detail_title)
        self.detail_table = QTableWidget(0, len(["ID"] + RAW_COLUMNS))
        self.detail_table.setHorizontalHeaderLabels(["ID"] + RAW_COLUMNS)
        self.configure_table(self.detail_table)
        detail_layout.addWidget(self.detail_table, 1)
        detail_bar = QHBoxLayout()
        edit_tx = QPushButton("编辑选中流水")
        del_tx = QPushButton("删除选中流水")
        del_month = QPushButton("删除选中月份流水")
        edit_tx.clicked.connect(self.edit_transaction)
        del_tx.clicked.connect(self.delete_transaction)
        del_month.clicked.connect(self.delete_month)
        detail_bar.addWidget(edit_tx)
        detail_bar.addWidget(del_tx)
        detail_bar.addWidget(del_month)
        detail_bar.addStretch()
        detail_layout.addLayout(detail_bar)

        splitter.addWidget(summary_widget)
        splitter.addWidget(detail_widget)
        splitter.setSizes([390, 430])

        self.status = QLabel("")
        self.status.setObjectName("statusLabel")
        root_layout.addWidget(self.status)

        formula = QLabel("口径：收入净额合计=订单实付应结+平台补贴+商家补贴+结算运费+订单退款；支出金额=已结算佣金+技术服务费；结算期末余额=期初金额+收入净额合计+支出金额+提现金额。")
        formula.setObjectName("formulaLabel")
        formula.setWordWrap(True)
        root_layout.addWidget(formula)
        self.setCentralWidget(root)

    def configure_table(self, table):
        table.setAlternatingRowColors(True)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setSelectionMode(QTableWidget.SingleSelection)
        table.verticalHeader().setVisible(False)
        table.horizontalHeader().setStretchLastSection(False)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)

    def apply_style(self):
        self.setStyleSheet("""
            QMainWindow, QWidget { background: #f5f7fb; color: #1f2937; font-family: "Microsoft YaHei"; font-size: 12px; }
            QToolBar { background: #ffffff; border-bottom: 1px solid #d9e2ef; spacing: 8px; padding: 8px; }
            QToolButton { background: #2563eb; color: #ffffff; border: none; border-radius: 6px; padding: 7px 12px; }
            QToolButton:hover { background: #1d4ed8; }
            QPushButton { background: #ffffff; border: 1px solid #cbd5e1; border-radius: 6px; padding: 7px 12px; }
            QPushButton:hover { background: #eff6ff; border-color: #93c5fd; }
            QLineEdit, QComboBox { background: #ffffff; border: 1px solid #cbd5e1; border-radius: 6px; padding: 7px 9px; }
            QTableWidget { background: #ffffff; alternate-background-color: #f8fafc; gridline-color: #e2e8f0; border: 1px solid #d9e2ef; border-radius: 6px; }
            QHeaderView::section { background: #eef2f7; color: #334155; padding: 7px; border: 0; border-right: 1px solid #d9e2ef; font-weight: 600; }
            QLabel#titleLabel { font-size: 24px; font-weight: 700; color: #0f172a; }
            QLabel#subtitleLabel, QLabel#statusLabel, QLabel#formulaLabel, QLabel#mutedLabel { color: #64748b; }
            QLabel#sectionTitle { font-size: 14px; font-weight: 700; color: #0f172a; padding: 4px 0; }
            QFrame#filterBar { background: #ffffff; border: 1px solid #d9e2ef; border-radius: 8px; }
        """)

    def clear_filters(self):
        self.store_filter.clear()
        self.month_filter.clear()
        self.refresh_all()

    def refresh_all(self):
        self.refresh_summary()
        self.refresh_details(None)

    def refresh_summary(self):
        self.summary_rows = self.repo.monthly_summaries(self.store_filter.text().strip(), self.month_filter.text().strip())
        self.summary_fixed.setRowCount(len(self.summary_rows))
        self.summary_scroll.setRowCount(len(self.summary_rows))
        for row_idx, row in enumerate(self.summary_rows):
            fixed_values = [row["store"], row["month_label"]]
            scroll_values = [
                money_text(row["paid_settlement"]),
                money_text(row["platform_subsidy"]),
                money_text(row["merchant_subsidy"]),
                money_text(row["freight"]),
                money_text(row["refund"]),
                money_text(row["income_total"]),
                money_text(row["commission"]),
                money_text(row["tech_fee"]),
                money_text(row["expense_amount"]),
                money_text(row["settlement_amount"]),
                money_text(row["withdraw_amount"]),
                money_text(row["opening_balance"]),
                money_text(row["ending_balance"]),
                "" if row["account_ending"] is None else money_text(row["account_ending"]),
                "" if row["difference"] is None else money_text(row["difference"]),
                difference_reason(row),
            ]
            has_diff = row["difference"] is not None and row["difference"] != Decimal("0.00")
            bg = QColor("#fff1f0") if has_diff else None
            for col, value in enumerate(fixed_values):
                item = table_item(value)
                if bg:
                    item.setBackground(bg)
                self.summary_fixed.setItem(row_idx, col, item)
            for col, value in enumerate(scroll_values):
                item = table_item(value, Qt.AlignLeft | Qt.AlignVCenter if col == len(scroll_values) - 1 else Qt.AlignCenter)
                if bg:
                    item.setBackground(bg)
                self.summary_scroll.setItem(row_idx, col, item)
        self.auto_fit(self.summary_fixed, max_width=130)
        self.auto_fit(self.summary_scroll, max_width=260)
        self.status.setText(f"数据库：{DB_PATH}    汇总 {len(self.summary_rows)} 条")

    def auto_fit(self, table, max_width=280):
        table.resizeColumnsToContents()
        for col in range(table.columnCount()):
            width = max(72, min(table.columnWidth(col) + 20, max_width))
            table.setColumnWidth(col, width)

    def on_fixed_selection(self):
        row = self.summary_fixed.currentRow()
        if row >= 0 and self.summary_scroll.currentRow() != row:
            self.summary_scroll.selectRow(row)
        self.select_summary_row(row)

    def on_scroll_selection(self):
        row = self.summary_scroll.currentRow()
        if row >= 0 and self.summary_fixed.currentRow() != row:
            self.summary_fixed.selectRow(row)
        self.select_summary_row(row)

    def select_summary_row(self, row):
        if row < 0 or row >= len(self.summary_rows):
            self.current_summary_key = None
            self.refresh_details(None)
            return
        selected = self.summary_rows[row]
        key = (selected["store"], selected["month_sort"])
        if key != self.current_summary_key:
            self.current_summary_key = key
            self.refresh_details(key)

    def selected_summary(self):
        row = self.summary_scroll.currentRow()
        if row < 0:
            row = self.summary_fixed.currentRow()
        if 0 <= row < len(self.summary_rows):
            return self.summary_rows[row]
        return None

    def refresh_details(self, key):
        self.detail_table.setRowCount(0)
        if not key:
            return
        rows = self.repo.details(key[0], key[1])
        self.detail_table.setRowCount(len(rows))
        for r, row in enumerate(rows):
            values = [
                row["id"], row["store"], row["month_label"], row["transaction_time"], row["flow_id"],
                row["direction"], row["account"], money_text(row["amount"]), row["summary"],
                row["biz_type"], row["main_order"], row["sub_order"], row["after_sale"],
                row["order_time"], row["product_info"], row["product_code"], row["sale_type"],
                money_text(row["paid_settlement"]), money_text(row["platform_subsidy"]),
                money_text(row["merchant_subsidy"]), money_text(row["freight"]),
                money_text(row["refund"]), money_text(row["commission"]), money_text(row["tech_fee"]),
            ]
            for c, value in enumerate(values):
                align = Qt.AlignLeft | Qt.AlignVCenter if c in (8, 13) else Qt.AlignCenter
                self.detail_table.setItem(r, c, table_item(value, align))
        self.auto_fit(self.detail_table, max_width=280)

    def manage_stores(self):
        dialog = StoreDialog(self, self.repo)
        dialog.exec()
        self.refresh_all()

    def import_excel(self):
        stores = [row["name"] for row in self.repo.configured_stores()]
        if not stores:
            if QMessageBox.question(self, "需要配置店铺", "还没有配置店铺。是否现在添加店铺？") == QMessageBox.Yes:
                self.manage_stores()
            return
        store, ok = QInputDialog.getItem(self, "选择导入店铺", "本次 Excel 归属店铺", stores, 0, False)
        if not ok or not store:
            return
        paths, _ = QFileDialog.getOpenFileNames(self, "选择资金流水 Excel", "", "Excel 文件 (*.xlsx *.xlsm);;所有文件 (*.*)")
        if not paths:
            return
        try:
            total_imported = 0
            total_skipped = 0
            for path in paths:
                result = self.repo.import_excel(path, selected_store=store)
                total_imported += result.imported
                total_skipped += result.skipped
            self.refresh_all()
            QMessageBox.information(self, "导入完成", f"店铺：{store}\n新增 {total_imported} 条，跳过重复/空行 {total_skipped} 条。")
        except Exception as exc:
            QMessageBox.critical(self, "导入失败", f"{exc}\n\n{traceback.format_exc(limit=2)}")

    def edit_balance(self):
        row = self.selected_summary()
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
        dialog = BalanceDialog(self, initial)
        if dialog.exec() == QDialog.Accepted:
            data = dialog.result_data
            self.repo.upsert_balance(data["store"], data["month_label"], data["month_sort"], data["opening_balance"], data["account_ending"], data["note"])
            self.refresh_all()
            QMessageBox.information(self, "保存成功", "期初金额和店铺期末余额已保存。")

    def add_adjustment(self):
        row = self.selected_summary()
        if not row:
            QMessageBox.warning(self, "未选择", "请先选择一个店铺月份。")
            return
        dialog = AdjustmentDialog(self, row)
        if dialog.exec() == QDialog.Accepted:
            data = dialog.result_data
            self.repo.add_adjustment(row["store"], row["month_label"], row["month_sort"], data["target_column"], data["item"], data["amount"], data["note"])
            self.refresh_all()

    def show_difference(self):
        row = self.selected_summary()
        if not row:
            QMessageBox.warning(self, "未选择", "请先选择一个店铺月份。")
            return
        DifferenceDialog(self, self.repo, row).exec()

    def current_transaction_id(self):
        row = self.detail_table.currentRow()
        if row < 0 or not self.detail_table.item(row, 0):
            return None
        try:
            return int(self.detail_table.item(row, 0).text())
        except ValueError:
            return None

    def edit_transaction(self):
        tx_id = self.current_transaction_id()
        if not tx_id:
            QMessageBox.warning(self, "未选择", "请先在明细流水中选择一条记录。")
            return
        tx = self.repo.transaction_by_id(tx_id)
        if not tx:
            QMessageBox.warning(self, "未找到", "这条流水记录不存在，可能已被删除。")
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
        dialog = TransactionDialog(self, initial)
        if dialog.exec() == QDialog.Accepted:
            self.repo.update_transaction(tx_id, dialog.result_data)
            self.refresh_all()

    def delete_transaction(self):
        tx_id = self.current_transaction_id()
        if not tx_id:
            QMessageBox.warning(self, "未选择", "请先在明细流水中选择一条记录。")
            return
        if QMessageBox.question(self, "确认删除", f"确定删除流水 ID {tx_id} 吗？") != QMessageBox.Yes:
            return
        self.repo.delete_transaction(tx_id)
        self.refresh_all()

    def delete_month(self):
        row = self.selected_summary()
        if not row:
            QMessageBox.warning(self, "未选择", "请先选择要删除的店铺月份。")
            return
        if QMessageBox.question(self, "确认删除", f"确定删除 {row['store']} {row['month_label']} 的全部流水和手工调整吗？余额记录默认保留。") != QMessageBox.Yes:
            return
        self.repo.delete_month(row["store"], row["month_sort"], delete_balance=False)
        self.refresh_all()

    def export_summary(self):
        path, _ = QFileDialog.getSaveFileName(self, "保存汇总CSV", "支付核对汇总.csv", "CSV 文件 (*.csv)")
        if not path:
            return
        rows = self.repo.monthly_summaries(self.store_filter.text().strip(), self.month_filter.text().strip())
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
        QMessageBox.information(self, "导出完成", f"已保存：{path}")


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
