from datetime import datetime

from PySide6.QtCore import QDate, Qt
from PySide6.QtGui import QStandardItem, QStandardItemModel
from PySide6.QtWidgets import (
    QComboBox,
    QDateEdit,
    QDoubleSpinBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QPushButton,
    QSpinBox,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from data_gen import generate_rows
from excel_export import export_to_xlsx, HEADERS


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("RPA Demo Desktop App - IFS Simulator")
        self.rows = []

        self.dp_from = QDateEdit()
        self.dp_from.setObjectName("dpFrom")
        self.dp_from.setCalendarPopup(True)
        self.dp_from.setDate(QDate.currentDate().addDays(-7))

        self.dp_to = QDateEdit()
        self.dp_to.setObjectName("dpTo")
        self.dp_to.setCalendarPopup(True)
        self.dp_to.setDate(QDate.currentDate())

        self.cb_category = QComboBox()
        self.cb_category.setObjectName("cbCategory")
        self.cb_category.addItems(["Sales", "Production", "Logistics"])

        self.sp_min_amount = QDoubleSpinBox()
        self.sp_min_amount.setObjectName("spMinAmount")
        self.sp_min_amount.setMinimum(0.0)
        self.sp_min_amount.setMaximum(100000.0)
        self.sp_min_amount.setDecimals(2)
        self.sp_min_amount.setValue(0.0)
        self.sp_min_amount.setSingleStep(1.0)

        self.sp_records = QSpinBox()
        self.sp_records.setObjectName("spRecords")
        self.sp_records.setRange(1, 5000)
        self.sp_records.setValue(50)

        self.btn_generate = QPushButton("Generate")
        self.btn_generate.setObjectName("btnGenerate")
        self.btn_generate.clicked.connect(self.on_generate)

        self.btn_export = QPushButton("Export")
        self.btn_export.setObjectName("btnExport")
        self.btn_export.clicked.connect(self.on_export)

        self.lbl_status = QLabel("Rows: 0 | Last generated: -")
        self.lbl_status.setObjectName("lblStatus")

        self.tbl_data = QTableView()
        self.tbl_data.setObjectName("tblData")
        self.tbl_data.setAlternatingRowColors(True)
        self.tbl_data.setSelectionBehavior(QTableView.SelectRows)
        self.tbl_data.setEditTriggers(QTableView.NoEditTriggers)

        self.model = QStandardItemModel(0, len(HEADERS))
        self.model.setHorizontalHeaderLabels(HEADERS)
        self.tbl_data.setModel(self.model)
        self.tbl_data.horizontalHeader().setStretchLastSection(True)

        controls = QHBoxLayout()
        controls.addWidget(QLabel("From"))
        controls.addWidget(self.dp_from)
        controls.addWidget(QLabel("To"))
        controls.addWidget(self.dp_to)
        controls.addWidget(QLabel("Category"))
        controls.addWidget(self.cb_category)
        controls.addWidget(QLabel("Min Amount"))
        controls.addWidget(self.sp_min_amount)
        controls.addWidget(QLabel("Records"))
        controls.addWidget(self.sp_records)
        controls.addWidget(self.btn_generate)
        controls.addWidget(self.btn_export)

        layout = QVBoxLayout()
        layout.addLayout(controls)
        layout.addWidget(self.tbl_data)
        layout.addWidget(self.lbl_status)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def on_generate(self) -> None:
        count = self.sp_records.value()
        date_from = self.dp_from.date().toPython()
        date_to = self.dp_to.date().toPython()
        dt_from = datetime(date_from.year, date_from.month, date_from.day, 0, 0, 0)
        dt_to = datetime(date_to.year, date_to.month, date_to.day, 23, 59, 59)
        category = self.cb_category.currentText()
        min_amount = self.sp_min_amount.value()

        self.rows = generate_rows(count, dt_from, dt_to, category, min_amount)
        self._populate_table(self.rows)

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.lbl_status.setText(f"Rows: {len(self.rows)} | Last generated: {now}")

    def on_export(self) -> None:
        if not self.rows:
            self.lbl_status.setText("Rows: 0 | Last generated: -")
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Excel",
            "raw.xlsx",
            "Excel Workbook (*.xlsx)",
        )
        if not path:
            return

        if not path.lower().endswith(".xlsx"):
            path = f"{path}.xlsx"

        export_to_xlsx(path, self.rows)

    def _populate_table(self, rows) -> None:
        self.model.setRowCount(0)
        for row in rows:
            items = []
            for header in HEADERS:
                value = row[header]
                if header == "Amount":
                    text = f"{value:.2f}"
                else:
                    text = str(value)
                item = QStandardItem(text)
                if header in {"Id", "Amount"}:
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                items.append(item)
            self.model.appendRow(items)
