import sys
import subprocess
from pathlib import Path
from typing import Optional

import pandas as pd
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
    QFileDialog,
    QListWidget,
    QListWidgetItem,
    QTextEdit,
    QFormLayout,
    QGroupBox,
    QComboBox,
)


class DropArea(QLabel):
    def __init__(self, on_file_dropped):
        super().__init__("Drop PDF here\n\nor click 'Open PDF'")
        self.on_file_dropped = on_file_dropped
        self.setAcceptDrops(True)
        self.setAlignment(Qt.AlignCenter)
        self.setStyleSheet(
            "QLabel { border: 2px dashed #888; padding: 24px; font-size: 16px; min-height: 150px; }"
        )

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.toLocalFile().lower().endswith(".pdf"):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(".pdf"):
                self.on_file_dropped(Path(path))
                event.acceptProposedAction()
                return
        event.ignore()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Warehouse Parser Dashboard")
        self.resize(1400, 850)

        self.current_pdf: Optional[Path] = None
        self.current_output_xlsx: Optional[Path] = None
        self.current_unknown_csv: Optional[Path] = None
        self.current_audit_csv: Optional[Path] = None
        self.current_unknown_df = pd.DataFrame()
        self.current_audit_df = pd.DataFrame()

        self.base_dir = Path(__file__).resolve().parent
        self.rules_dir = self.base_dir / "Rules"
        self.parser_path = self.base_dir / "parser.py"

        self._build_ui()
        self.refresh_rule_counts()

    def _build_ui(self):
        root = QWidget()
        root_layout = QVBoxLayout(root)

        # Top controls
        top_bar = QHBoxLayout()
        self.pdf_path_edit = QLineEdit()
        self.pdf_path_edit.setPlaceholderText("Select or drop a PDF...")
        self.pdf_path_edit.setReadOnly(True)

        open_btn = QPushButton("Open PDF")
        open_btn.clicked.connect(self.choose_pdf)

        run_btn = QPushButton("Run Parser")
        run_btn.clicked.connect(self.run_parser)

        refresh_btn = QPushButton("Reload Unknowns")
        refresh_btn.clicked.connect(self.reload_outputs)

        top_bar.addWidget(self.pdf_path_edit)
        top_bar.addWidget(open_btn)
        top_bar.addWidget(run_btn)
        top_bar.addWidget(refresh_btn)

        root_layout.addLayout(top_bar)

        splitter = QSplitter()
        root_layout.addWidget(splitter)

        # Left panel
        left = QWidget()
        left_layout = QVBoxLayout(left)
        self.drop_area = DropArea(self.set_pdf)
        left_layout.addWidget(self.drop_area)

        status_group = QGroupBox("Status")
        status_layout = QFormLayout(status_group)
        self.status_pdf = QLabel("—")
        self.status_rows = QLabel("—")
        self.status_unknowns = QLabel("—")
        self.status_valid = QLabel("—")
        self.status_corrections = QLabel("—")
        self.status_conf_high = QLabel("—")
        self.status_conf_medium = QLabel("—")
        self.status_conf_low = QLabel("—")
        status_layout.addRow("Current PDF:", self.status_pdf)
        status_layout.addRow("Parsed rows:", self.status_rows)
        status_layout.addRow("Unknown parts:", self.status_unknowns)
        status_layout.addRow("Valid parts:", self.status_valid)
        status_layout.addRow("Corrections:", self.status_corrections)
        status_layout.addRow("Confidence high:", self.status_conf_high)
        status_layout.addRow("Confidence medium:", self.status_conf_medium)
        status_layout.addRow("Confidence low:", self.status_conf_low)
        left_layout.addWidget(status_group)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        left_layout.addWidget(self.log_box, stretch=1)

        splitter.addWidget(left)

        # Middle panel
        middle = QWidget()
        middle_layout = QVBoxLayout(middle)
        middle_layout.addWidget(QLabel("Parsed Output Preview"))
        self.output_table = QTableWidget()
        middle_layout.addWidget(self.output_table)
        splitter.addWidget(middle)

        # Right panel
        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.addWidget(QLabel("Unknown Parts Review"))

        self.unknown_list = QListWidget()
        self.unknown_list.currentRowChanged.connect(self.show_unknown_details)
        right_layout.addWidget(self.unknown_list)

        details_group = QGroupBox("Selected Unknown")
        details_layout = QFormLayout(details_group)
        self.detail_before = QLineEdit()
        self.detail_after = QLineEdit()
        self.detail_display = QLineEdit()
        self.detail_po = QLineEdit()
        self.detail_source = QLineEdit()
        for w in [self.detail_before, self.detail_after, self.detail_display, self.detail_po, self.detail_source]:
            w.setReadOnly(True)
        self.detail_raw = QTextEdit()
        self.detail_raw.setReadOnly(True)

        details_layout.addRow("Before correction:", self.detail_before)
        details_layout.addRow("Current normalized:", self.detail_after)
        details_layout.addRow("Display part:", self.detail_display)
        details_layout.addRow("PO:", self.detail_po)
        details_layout.addRow("Source PDF:", self.detail_source)
        details_layout.addRow("Raw OCR line:", self.detail_raw)
        right_layout.addWidget(details_group)

        actions_group = QGroupBox("Rule Actions")
        actions_layout = QVBoxLayout(actions_group)

        self.add_valid_btn = QPushButton("Add current normalized to valid parts")
        self.add_valid_btn.clicked.connect(self.add_selected_to_valid_parts)

        self.correction_target_edit = QLineEdit()
        self.correction_target_edit.setPlaceholderText("Correct part number, e.g. A-H135423")

        self.add_correction_btn = QPushButton("Add correction (before → target)")
        self.add_correction_btn.clicked.connect(self.add_selected_correction)

        actions_layout.addWidget(self.add_valid_btn)
        actions_layout.addWidget(self.correction_target_edit)
        actions_layout.addWidget(self.add_correction_btn)

        duplicate_group = QGroupBox("Duplicate Rule Editor")
        duplicate_layout = QVBoxLayout(duplicate_group)

        self.duplicate_part_edit = QLineEdit()
        self.duplicate_part_edit.setPlaceholderText("Part number for duplicate rule")

        self.duplicate_type_combo = QComboBox()
        self.duplicate_type_combo.addItems(["B", "TAG", "BL", "LBL", "K1", "K2", "K3"])

        self.add_duplicate_btn = QPushButton("Add duplicate type for selected part")
        self.add_duplicate_btn.clicked.connect(self.add_selected_duplicate_rule)

        self.remove_duplicate_btn = QPushButton("Remove all manual duplicate rules for selected part")
        self.remove_duplicate_btn.clicked.connect(self.remove_selected_duplicate_rules)

        duplicate_layout.addWidget(self.duplicate_part_edit)
        duplicate_layout.addWidget(self.duplicate_type_combo)
        duplicate_layout.addWidget(self.add_duplicate_btn)
        duplicate_layout.addWidget(self.remove_duplicate_btn)

        actions_layout.addWidget(duplicate_group)
        right_layout.addWidget(actions_group)

        splitter.addWidget(right)
        splitter.setSizes([300, 650, 450])

        self.setCentralWidget(root)

    def log(self, text: str):
        self.log_box.append(text)

    def choose_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select PDF", str(self.base_dir), "PDF Files (*.pdf)")
        if path:
            self.set_pdf(Path(path))

    def set_pdf(self, path: Path):
        self.current_pdf = path
        self.pdf_path_edit.setText(str(path))
        self.status_pdf.setText(path.name)
        self.log(f"Selected PDF: {path}")

    def run_parser(self):
        if not self.current_pdf:
            QMessageBox.warning(self, "No PDF selected", "Select or drop a PDF first.")
            return

        if not self.parser_path.exists():
            QMessageBox.critical(self, "Missing parser.py", f"Could not find parser.py at:\n{self.parser_path}")
            return

        self.log("Running parser...")
        try:
            result = subprocess.run(
                [sys.executable, str(self.parser_path), str(self.current_pdf)],
                cwd=str(self.base_dir),
                capture_output=True,
                text=True,
                check=False,
            )
        except Exception as e:
            QMessageBox.critical(self, "Parser failed to launch", str(e))
            return

        if result.stdout:
            self.log(result.stdout.strip())
        if result.stderr:
            self.log(result.stderr.strip())

        if result.returncode != 0:
            QMessageBox.critical(self, "Parser error", f"Parser exited with code {result.returncode}.")
            return

        self.reload_outputs()

    def reload_outputs(self):
        if not self.current_pdf:
            return

        self.current_output_xlsx = self.current_pdf.with_name(self.current_pdf.stem + "_output.xlsx")
        self.current_unknown_csv = self.current_pdf.with_name(self.current_pdf.stem + "_unknown_parts.csv")
        self.current_audit_csv = self.current_pdf.with_name(self.current_pdf.stem + "_correction_audit.csv")

        if self.current_audit_csv.exists():
            self.log(f"Found correction audit: {self.current_audit_csv.name}")

        self.load_output_preview()
        self.load_unknowns()
        self.load_correction_audit_summary()
        self.refresh_rule_counts()

    def load_output_preview(self):
        self.output_table.clear()
        if not self.current_output_xlsx or not self.current_output_xlsx.exists():
            self.status_rows.setText("0")
            return

        try:
            df = pd.read_excel(self.current_output_xlsx)
        except Exception as e:
            self.log(f"Failed reading output workbook: {e}")
            return

        self.status_rows.setText(str(len(df)))
        self.output_table.setRowCount(len(df))
        self.output_table.setColumnCount(len(df.columns))
        self.output_table.setHorizontalHeaderLabels([str(c) for c in df.columns])

        preview = df.head(500)
        self.output_table.setRowCount(len(preview))
        for r_idx, (_, row) in enumerate(preview.iterrows()):
            for c_idx, value in enumerate(row.tolist()):
                self.output_table.setItem(r_idx, c_idx, QTableWidgetItem("" if pd.isna(value) else str(value)))

        self.output_table.resizeColumnsToContents()

    def load_unknowns(self):
        self.unknown_list.clear()
        self.current_unknown_df = pd.DataFrame()

        if not self.current_unknown_csv or not self.current_unknown_csv.exists():
            self.status_unknowns.setText("0")
            self.clear_unknown_details()
            return

        try:
            df = pd.read_csv(self.current_unknown_csv)
        except Exception as e:
            self.log(f"Failed reading unknown parts CSV: {e}")
            self.clear_unknown_details()
            return

        self.current_unknown_df = df.fillna("")
        self.status_unknowns.setText(str(len(self.current_unknown_df)))

        for _, row in self.current_unknown_df.iterrows():
            label = str(row.get("part_number_display") or row.get("part_number_norm") or "(unknown)")
            item = QListWidgetItem(label)
            self.unknown_list.addItem(item)

        if len(self.current_unknown_df) > 0:
            self.unknown_list.setCurrentRow(0)
        else:
            self.clear_unknown_details()

    def load_correction_audit_summary(self):
        self.current_audit_df = pd.DataFrame()
        self.status_conf_high.setText("0")
        self.status_conf_medium.setText("0")
        self.status_conf_low.setText("0")

        if not self.current_audit_csv or not self.current_audit_csv.exists():
            return

        try:
            df = pd.read_csv(self.current_audit_csv)
        except Exception as e:
            self.log(f"Failed reading correction audit CSV: {e}")
            return

        self.current_audit_df = df.fillna("")
        if "confidence" not in self.current_audit_df.columns:
            self.log("Correction audit CSV missing 'confidence' column.")
            return

        confidence_counts = (
            self.current_audit_df["confidence"]
            .astype(str)
            .str.strip()
            .str.lower()
            .value_counts()
        )

        self.status_conf_high.setText(str(int(confidence_counts.get("high", 0))))
        self.status_conf_medium.setText(str(int(confidence_counts.get("medium", 0))))
        self.status_conf_low.setText(str(int(confidence_counts.get("low", 0))))
            

    def clear_unknown_details(self):
        self.detail_before.clear()
        self.detail_after.clear()
        self.detail_display.clear()
        self.detail_po.clear()
        self.detail_source.clear()
        self.detail_raw.clear()
        self.correction_target_edit.clear()

    def show_unknown_details(self, row_index: int):
        if row_index < 0 or self.current_unknown_df.empty or row_index >= len(self.current_unknown_df):
            self.clear_unknown_details()
            return

        row = self.current_unknown_df.iloc[row_index]
        before = str(row.get("part_number_before_correction", ""))
        after = str(row.get("part_number_norm", ""))

        self.detail_before.setText(before)
        self.detail_after.setText(after)
        self.detail_display.setText(str(row.get("part_number_display", "")))
        self.detail_po.setText(str(row.get("po", "")))
        self.detail_source.setText(str(row.get("source_pdf", "")))
        self.detail_raw.setPlainText(str(row.get("raw_line", "")))
        self.correction_target_edit.setText(after)
        self.duplicate_part_edit.setText(after)

    def refresh_rule_counts(self):
        valid_path = self.rules_dir / "valid_part_numbers.csv"
        corrections_path = self.rules_dir / "part_corrections.csv"

        valid_count = self.safe_count_rows(valid_path)
        corrections_count = self.safe_count_rows(corrections_path)

        self.status_valid.setText(str(valid_count) if valid_count is not None else "missing")
        self.status_corrections.setText(str(corrections_count) if corrections_count is not None else "missing")

    def safe_count_rows(self, path: Path) -> Optional[int]:
        if not path.exists():
            return None
        try:
            df = pd.read_csv(path)
            return len(df)
        except Exception:
            return None

    def add_selected_to_valid_parts(self):
        idx = self.unknown_list.currentRow()
        if idx < 0 or self.current_unknown_df.empty:
            QMessageBox.information(self, "No selection", "Select an unknown part first.")
            return

        value = str(self.current_unknown_df.iloc[idx].get("part_number_norm", "")).strip()
        if not value:
            QMessageBox.warning(self, "Missing part", "Selected row has no normalized part number.")
            return

        path = self.rules_dir / "valid_part_numbers.csv"
        self.rules_dir.mkdir(parents=True, exist_ok=True)

        if path.exists():
            df = pd.read_csv(path)
            if "part_number" not in df.columns:
                QMessageBox.critical(self, "Bad CSV", f"{path.name} must contain a 'part_number' column.")
                return
        else:
            df = pd.DataFrame(columns=["part_number"])

        existing = set(df["part_number"].astype(str).str.strip().str.upper())
        if value.upper() not in existing:
            df.loc[len(df)] = {"part_number": value}
            df = df.drop_duplicates(subset=["part_number"])
            df = df.sort_values("part_number")
            df.to_csv(path, index=False)
            self.log(f"Added to valid parts: {value}")
        else:
            self.log(f"Already in valid parts: {value}")

        self.refresh_rule_counts()

    def add_selected_correction(self):
        idx = self.unknown_list.currentRow()
        if idx < 0 or self.current_unknown_df.empty:
            QMessageBox.information(self, "No selection", "Select an unknown part first.")
            return

        target = self.correction_target_edit.text().strip().upper()
        before = str(self.current_unknown_df.iloc[idx].get("part_number_before_correction", "")).strip().upper()
        if not before:
            before = str(self.current_unknown_df.iloc[idx].get("part_number_norm", "")).strip().upper()

        if not before or not target:
            QMessageBox.warning(self, "Missing values", "Need both a source part and a correction target.")
            return

        path = self.rules_dir / "part_corrections.csv"
        self.rules_dir.mkdir(parents=True, exist_ok=True)

        if path.exists():
            df = pd.read_csv(path)
            required = {"bad_part", "good_part"}
            if not required.issubset(df.columns):
                QMessageBox.critical(self, "Bad CSV", f"{path.name} must contain bad_part,good_part columns.")
                return
        else:
            df = pd.DataFrame(columns=["bad_part", "good_part"])

        mask = df["bad_part"].astype(str).str.strip().str.upper() == before
        if mask.any():
            df.loc[mask, "good_part"] = target
            self.log(f"Updated correction: {before} -> {target}")
        else:
            df.loc[len(df)] = {"bad_part": before, "good_part": target}
            self.log(f"Added correction: {before} -> {target}")

        df = df.drop_duplicates(subset=["bad_part"], keep="last")
        df = df.sort_values("bad_part")
        df.to_csv(path, index=False)

        self.refresh_rule_counts()

    def add_duplicate_rule(self, part_number: str, rule_type: str):
        path = self.rules_dir / "duplicate_parts_manual.csv"
        self.rules_dir.mkdir(parents=True, exist_ok=True)

        if path.exists():
            df = pd.read_csv(path)
            required = {"part_number", "type"}
            if not required.issubset(df.columns):
                QMessageBox.critical(self, "Bad CSV", f"{path.name} must contain part_number,type columns.")
                return
        else:
            df = pd.DataFrame(columns=["part_number", "type"])

        part_number = str(part_number).strip().upper()
        rule_type = str(rule_type).strip().upper()

        mask = (
            df["part_number"].astype(str).str.strip().str.upper().eq(part_number)
            & df["type"].astype(str).str.strip().str.upper().eq(rule_type)
        )

        if mask.any():
            self.log(f"Duplicate rule already exists: {part_number} -> {rule_type}")
            return

        df.loc[len(df)] = {"part_number": part_number, "type": rule_type}
        df = df.drop_duplicates(subset=["part_number", "type"])
        df = df.sort_values(["part_number", "type"])
        df.to_csv(path, index=False)
        self.log(f"Added duplicate rule: {part_number} -> {rule_type}")

    def add_selected_duplicate_rule(self):
        part_number = self.duplicate_part_edit.text().strip().upper()
        rule_type = self.duplicate_type_combo.currentText().strip().upper()

        if not part_number:
            QMessageBox.warning(self, "Missing part", "Select an unknown part or type a part number first.")
            return

        self.add_duplicate_rule(part_number, rule_type)

    def remove_selected_duplicate_rules(self):
        part_number = self.duplicate_part_edit.text().strip().upper()
        if not part_number:
            QMessageBox.warning(self, "Missing part", "Select an unknown part or type a part number first.")
            return

        path = self.rules_dir / "duplicate_parts_manual.csv"
        if not path.exists():
            self.log("No duplicate_parts_manual.csv found.")
            return

        df = pd.read_csv(path)
        required = {"part_number", "type"}
        if not required.issubset(df.columns):
            QMessageBox.critical(self, "Bad CSV", f"{path.name} must contain part_number,type columns.")
            return

        before = len(df)
        df = df[~df["part_number"].astype(str).str.strip().str.upper().eq(part_number)]
        after = len(df)

        if after == before:
            self.log(f"No manual duplicate rules found for: {part_number}")
            return

        df.to_csv(path, index=False)
        self.log(f"Removed manual duplicate rules for: {part_number}")

        self.refresh_rule_counts()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
