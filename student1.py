# fee_manager_strict.py
# fee_manager_strict_final.py
"""
Strict Student Fee Manager (PyQt5)
- Enforces exact columns:
  Name | Mobile Number | Year | Dept | Fee Amount | Fee Paid | Balance | Due Date | Email | Fee Paid On
- Only files created by this app (exact headers & order) can be opened.
- Auto-calculates Balance, auto-fills Fee Paid On when balance==0, clears Due Date when fully paid.
- GUI: Create/Open/Save, Add/Edit/Delete rows, Search, Date filter, Export PDF, Send email reminders.
- Improved SMTP error reporting and brief App Password hint.
"""

import sys
import os
from datetime import datetime, date
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTableWidget,
    QTableWidgetItem, QFileDialog, QLineEdit, QLabel, QMessageBox, QInputDialog,
    QDialog, QFormLayout, QDialogButtonBox, QDateEdit, QSpinBox, QComboBox
)
from PyQt5.QtCore import QDate
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
import smtplib
from email.mime.text import MIMEText

# Strict columns (exact order & spelling)
STRICT_COLUMNS = [
    "Name", "Mobile Number", "Year", "Dept",
    "Fee Amount", "Fee Paid", "Balance",
    "Due Date", "Email", "Fee Paid On"
]

DATE_FORMAT = "%Y-%m-%d"  # YYYY-MM-DD


def today_str():
    return date.today().strftime(DATE_FORMAT)


class StudentDialog(QDialog):
    """Dialog to Add / Edit a row using strict columns."""
    def __init__(self, data=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Student Record")
        self.data = data or {}
        self.widgets = {}
        form = QFormLayout()

        # Name
        w = QLineEdit(self.data.get("Name", ""))
        form.addRow("Name:", w); self.widgets["Name"] = w

        # Mobile Number (text)
        w = QLineEdit(self.data.get("Mobile Number", ""))
        form.addRow("Mobile Number:", w); self.widgets["Mobile Number"] = w

        # Year (text)
        w = QLineEdit(self.data.get("Year", ""))
        form.addRow("Year:", w); self.widgets["Year"] = w

        # Dept (text)
        w = QLineEdit(self.data.get("Dept", ""))
        form.addRow("Dept:", w); self.widgets["Dept"] = w

        # Fee Amount (integer)
        w = QSpinBox(); w.setMaximum(10_000_000)
        try: w.setValue(int(self.data.get("Fee Amount", 0) or 0))
        except Exception: w.setValue(0)
        form.addRow("Fee Amount:", w); self.widgets["Fee Amount"] = w

        # Fee Paid (integer)
        w = QSpinBox(); w.setMaximum(10_000_000)
        try: w.setValue(int(self.data.get("Fee Paid", 0) or 0))
        except Exception: w.setValue(0)
        form.addRow("Fee Paid:", w); self.widgets["Fee Paid"] = w

        # Balance (read-only; computed)
        bal_widget = QLineEdit(str(self.data.get("Balance", "")))
        bal_widget.setReadOnly(True)
        form.addRow("Balance (auto):", bal_widget); self.widgets["Balance"] = bal_widget

        # Due Date (date)
        due_widget = QDateEdit()
        due_widget.setCalendarPopup(True)
        due_val = self.data.get("Due Date", "")
        if due_val:
            try:
                dt = pd.to_datetime(due_val, errors='coerce')
                if pd.notna(dt):
                    due_widget.setDate(QDate(dt.year, dt.month, dt.day))
                else:
                    due_widget.setDate(QDate.currentDate())
            except Exception:
                due_widget.setDate(QDate.currentDate())
        else:
            due_widget.setDate(QDate.currentDate())
        form.addRow("Due Date:", due_widget); self.widgets["Due Date"] = due_widget

        # Email
        w = QLineEdit(self.data.get("Email", ""))
        form.addRow("Email:", w); self.widgets["Email"] = w

        # Fee Paid On (read-only, auto)
        paid_on_widget = QLineEdit(self.data.get("Fee Paid On", ""))
        paid_on_widget.setReadOnly(True)
        form.addRow("Fee Paid On (auto):", paid_on_widget); self.widgets["Fee Paid On"] = paid_on_widget

        buttons = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        dlg_layout = QVBoxLayout()
        dlg_layout.addLayout(form)
        dlg_layout.addWidget(buttons)
        self.setLayout(dlg_layout)

        # connect to update balance when fee fields change
        self.widgets["Fee Amount"].valueChanged.connect(self._update_balance_field)
        self.widgets["Fee Paid"].valueChanged.connect(self._update_balance_field)
        # initialize balance field
        self._update_balance_field()

    def _update_balance_field(self):
        amt = int(self.widgets["Fee Amount"].value())
        paid = int(self.widgets["Fee Paid"].value())
        bal = max(0, amt - paid)
        self.widgets["Balance"].setText(str(bal))
        # set Fee Paid On / Due Date based on balance
        if bal == 0:
            self.widgets["Fee Paid On"].setText(today_str())
        else:
            self.widgets["Fee Paid On"].setText("")

    def get_data(self):
        # Return dict with all strict columns
        out = {}
        out["Name"] = self.widgets["Name"].text().strip()
        out["Mobile Number"] = self.widgets["Mobile Number"].text().strip()
        out["Year"] = self.widgets["Year"].text().strip()
        out["Dept"] = self.widgets["Dept"].text().strip()
        out["Fee Amount"] = int(self.widgets["Fee Amount"].value())
        out["Fee Paid"] = int(self.widgets["Fee Paid"].value())
        # Balance computed
        out["Balance"] = int(out["Fee Amount"]) - int(out["Fee Paid"])
        # Due Date: get string YYYY-MM-DD
        due_qdate = self.widgets["Due Date"].date()
        out["Due Date"] = due_qdate.toString("yyyy-MM-dd")
        out["Email"] = self.widgets["Email"].text().strip()
        # Fee Paid On: if balance==0 set today else blank
        out["Fee Paid On"] = today_str() if out["Balance"] == 0 else ""
        return out


class FeeManagerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Student Fee Manager (Strict)")
        self.resize(1150, 700)

        self.files = {}  # display name -> (full_path, dataframe)
        self.selected_name = None
        self.current_df_view = None

        self._build_ui()

    def _build_ui(self):
        main = QVBoxLayout()

        # Top controls
        top = QHBoxLayout()
        btn_create = QPushButton("Create New Excel (strict)")
        btn_create.clicked.connect(self.create_new_file)
        top.addWidget(btn_create)

        btn_open = QPushButton("Open Excel (strict)")
        btn_open.clicked.connect(self.open_file)
        top.addWidget(btn_open)

        btn_save = QPushButton("Save Current")
        btn_save.clicked.connect(self.save_current)
        top.addWidget(btn_save)

        top.addWidget(QLabel(" | Select File: "))
        self.combo = QComboBox()
        self.combo.currentTextChanged.connect(self.on_select_file)
        top.addWidget(self.combo)

        top.addWidget(QLabel("Search:"))
        self.search = QLineEdit(); self.search.setPlaceholderText("Search across all columns")
        self.search.textChanged.connect(self.apply_search)
        top.addWidget(self.search)

        # Date filter
        top.addWidget(QLabel("Filter by Month:"))
        self.month_combo = QComboBox()
        self.month_combo.addItem("All")
        for m in range(1, 13):
            self.month_combo.addItem(f"{m:02d}")
        top.addWidget(self.month_combo)

        top.addWidget(QLabel("Year:"))
        self.filter_year = QLineEdit(); self.filter_year.setFixedWidth(70)
        top.addWidget(self.filter_year)

        btn_apply = QPushButton("Apply Date Filter")
        btn_apply.clicked.connect(self.apply_date_filter)
        top.addWidget(btn_apply)

        btn_clear = QPushButton("Clear Filters")
        btn_clear.clicked.connect(self.clear_filters)
        top.addWidget(btn_clear)

        main.addLayout(top)

        # Table
        self.table = QTableWidget()
        self.table.setColumnCount(len(STRICT_COLUMNS))
        self.table.setHorizontalHeaderLabels(STRICT_COLUMNS)
        main.addWidget(self.table)

        # Buttons row
        row = QHBoxLayout()
        btn_add = QPushButton("Add Row")
        btn_add.clicked.connect(self.add_row)
        row.addWidget(btn_add)

        btn_edit = QPushButton("Edit Selected")
        btn_edit.clicked.connect(self.edit_selected)
        row.addWidget(btn_edit)

        btn_delete = QPushButton("Delete Selected")
        btn_delete.clicked.connect(self.delete_selected)
        row.addWidget(btn_delete)

        btn_pdf = QPushButton("Export PDF (current view)")
        btn_pdf.clicked.connect(self.export_pdf)
        row.addWidget(btn_pdf)

        btn_email = QPushButton("Send Reminders (selected/all due)")
        btn_email.clicked.connect(self.send_reminders)
        row.addWidget(btn_email)

        main.addLayout(row)
        self.setLayout(main)

    # ---------- file operations ----------
    def create_new_file(self):
        path, _ = QFileDialog.getSaveFileName(self, "Create new strict Excel file", "students_fees_strict.xlsx", "Excel files (*.xlsx)")
        if not path:
            return
        # create df with strict columns exactly in order
        df = pd.DataFrame(columns=STRICT_COLUMNS)
        try:
            df.to_excel(path, index=False)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to create file: {e}")
            return
        name = os.path.basename(path)
        self.files[name] = (path, df)
        if self.combo.findText(name) == -1:
            self.combo.addItem(name)
        self.combo.setCurrentText(name)
        QMessageBox.information(self, "Created", f"Created strict file: {name}")

    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open strict Excel", "", "Excel files (*.xlsx)")
        if not path:
            return
        try:
            df = pd.read_excel(path)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to read Excel: {e}")
            return
        # validate strict header equality (order & names)
        if list(df.columns) != STRICT_COLUMNS:
            QMessageBox.critical(self, "Invalid format",
                                 "This file is not in the strict format.\n"
                                 "Only files created by this application with exact headers can be opened.")
            return
        # compute balance to ensure correctness
        df = self._recalculate_business_rules(df)
        name = os.path.basename(path)
        self.files[name] = (path, df)
        if self.combo.findText(name) == -1:
            self.combo.addItem(name)
        self.combo.setCurrentText(name)
        QMessageBox.information(self, "Loaded", f"Loaded {name}")

    def save_current(self):
        if not self.selected_name:
            QMessageBox.warning(self, "No file", "No file selected")
            return
        path, df = self.files[self.selected_name]
        df = self._recalculate_business_rules(df)
        try:
            df.to_excel(path, index=False)
            self.files[self.selected_name] = (path, df)
            QMessageBox.information(self, "Saved", f"Saved {self.selected_name}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save: {e}")

    def on_select_file(self, name):
        if not name:
            return
        self.selected_name = name
        path, df = self.files[name]
        self.current_df_view = df.copy()
        self.refresh_table(self.current_df_view)

    # ---------- table / CRUD ----------
    def refresh_table(self, df: pd.DataFrame):
        self.table.setRowCount(0)
        if df is None or df.empty:
            return
        self.table.setRowCount(len(df))
        for i, (_, row) in enumerate(df.iterrows()):
            for j, col in enumerate(STRICT_COLUMNS):
                val = row.get(col, "")
                display = "" if pd.isna(val) else str(val)
                self.table.setItem(i, j, QTableWidgetItem(display))

    def add_row(self):
        if not self.selected_name:
            QMessageBox.warning(self, "No file", "Create or open a strict file first")
            return
        dlg = StudentDialog(parent=self)
        if dlg.exec_():
            rec = dlg.get_data()
            path, df = self.files[self.selected_name]
            row_df = pd.DataFrame([rec], columns=STRICT_COLUMNS)
            df = pd.concat([df, row_df], ignore_index=True)
            df = self._recalculate_business_rules(df)
            self.files[self.selected_name] = (path, df)
            self.current_df_view = df.copy()
            self.refresh_table(self.current_df_view)

    def edit_selected(self):
        if not self.selected_name:
            QMessageBox.warning(self, "No file", "Open a strict file first")
            return
        row_idx = self.table.currentRow()
        if row_idx < 0:
            QMessageBox.warning(self, "Select", "Select a row to edit")
            return
        path, df = self.files[self.selected_name]
        existing = df.iloc[row_idx].to_dict()
        dlg = StudentDialog(data=existing, parent=self)
        if dlg.exec_():
            updated = dlg.get_data()
            for col in STRICT_COLUMNS:
                df.at[row_idx, col] = updated.get(col, "")
            df = self._recalculate_business_rules(df)
            self.files[self.selected_name] = (path, df)
            self.current_df_view = df.copy()
            self.refresh_table(self.current_df_view)

    def delete_selected(self):
        if not self.selected_name:
            QMessageBox.warning(self, "No file", "Open a strict file first")
            return
        row_idx = self.table.currentRow()
        if row_idx < 0:
            QMessageBox.warning(self, "Select", "Select row to delete")
            return
        path, df = self.files[self.selected_name]
        df = df.drop(df.index[row_idx]).reset_index(drop=True)
        df = self._recalculate_business_rules(df)
        self.files[self.selected_name] = (path, df)
        self.current_df_view = df.copy()
        self.refresh_table(self.current_df_view)

    # ---------- search & filter ----------
    def apply_search(self):
        if not self.selected_name:
            return
        keyword = self.search.text().strip().lower()
        path, df = self.files[self.selected_name]
        if keyword == "":
            self.current_df_view = df.copy()
        else:
            mask = pd.Series(False, index=df.index)
            for col in STRICT_COLUMNS:
                mask = mask | df[col].astype(str).str.lower().str.contains(keyword, na=False)
            self.current_df_view = df[mask].copy()
        self.refresh_table(self.current_df_view)

    def apply_date_filter(self):
        if not self.selected_name:
            return
        mon = self.month_combo.currentText()
        yr_text = self.filter_year.text().strip()
        path, df = self.files[self.selected_name]
        try:
            parsed = pd.to_datetime(df["Due Date"], errors='coerce')
        except Exception:
            QMessageBox.warning(self, "Date parse", "Could not parse Due Date column")
            return
        mask = pd.Series(True, index=df.index)
        if mon != "All":
            mask = mask & (parsed.dt.month == int(mon))
        if yr_text:
            try:
                y = int(yr_text)
                mask = mask & (parsed.dt.year == y)
            except ValueError:
                QMessageBox.warning(self, "Year", "Year must be numeric")
                return
        self.current_df_view = df[mask].copy()
        self.refresh_table(self.current_df_view)

    def clear_filters(self):
        if not self.selected_name:
            return
        path, df = self.files[self.selected_name]
        self.current_df_view = df.copy()
        self.refresh_table(self.current_df_view)

    # ---------- business rules ----------
    def _recalculate_business_rules(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        for c in STRICT_COLUMNS:
            if c not in df.columns:
                df[c] = ""
        df = df[STRICT_COLUMNS]
        df["Fee Amount"] = pd.to_numeric(df["Fee Amount"], errors='coerce').fillna(0).astype(int)
        df["Fee Paid"] = pd.to_numeric(df["Fee Paid"], errors='coerce').fillna(0).astype(int)
        df["Balance"] = (df["Fee Amount"] - df["Fee Paid"]).clip(lower=0).astype(int)
        for idx, row in df.iterrows():
            bal = int(row["Balance"])
            if bal == 0:
                if not row.get("Fee Paid On"):
                    df.at[idx, "Fee Paid On"] = today_str()
                df.at[idx, "Due Date"] = ""
            else:
                df.at[idx, "Fee Paid On"] = ""
        return df

    # ---------- PDF export ----------
    def export_pdf(self):
        if self.current_df_view is None or self.current_df_view.empty:
            QMessageBox.warning(self, "No data", "No data to export")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "fee_report.pdf", "PDF files (*.pdf)")
        if not path:
            return
        try:
            self._build_pdf(self.current_df_view, path)
            QMessageBox.information(self, "Exported", f"Saved PDF to {path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to export PDF: {e}")

    def _build_pdf(self, df: pd.DataFrame, out_path: str):
        from reportlab.lib.pagesizes import A4
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib import colors
        from reportlab.lib.styles import getSampleStyleSheet

        doc = SimpleDocTemplate(out_path, pagesize=A4)
        styles = getSampleStyleSheet()
        elems = [Paragraph("Student Fee Report", styles["Title"]), Spacer(1, 12)]

    # Columns you want in the PDF
        pdf_columns = ["Name", "Mobile Number", "Balance", "Due Date"]

    # Filter DataFrame to only those columns (ignore if missing)
        available_cols = [col for col in pdf_columns if col in df.columns]
        data = [available_cols]  # header row

    # Add the rows with only these columns, converting NaNs to empty strings
        for _, row in df.iterrows():
            row_data = [str(row[col]) if pd.notna(row[col]) else "" for col in available_cols]
            data.append(row_data)

        table = Table(data, repeatRows=1)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#4F81BD")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ]))
        elems.append(table)
        doc.build(elems)

    # ---------- email reminders ----------
    def send_reminders(self):
        if not self.selected_name:
            QMessageBox.warning(self, "No file", "Open a strict file first")
            return
        choice, ok = QInputDialog.getItem(self, "Reminders", "Send to:", ["All with dues", "Selected row only"], 0, False)
        if not ok:
            return
        path, df = self.files[self.selected_name]
        df = self._recalculate_business_rules(df)
        recipients = []
        if choice == "Selected row only":
            r = self.table.currentRow()
            if r < 0:
                QMessageBox.warning(self, "Select", "Select a row first")
                return
            row = df.iloc[r]
            if int(row["Balance"]) <= 0:
                QMessageBox.information(self, "No due", "Selected row has no due balance")
                return
            recipients.append(row)
        else:
            for _, row in df.iterrows():
                if int(row["Balance"]) > 0:
                    recipients.append(row)
        if not recipients:
            QMessageBox.information(self, "No recipients", "No students with dues found")
            return

        # Hint about App Password
        QMessageBox.information(self, "SMTP info",
                                "If using Gmail, you must enable 2-step verification and create an App Password.\n"
                                "Use that App Password (16 chars) as the 'password' below. Regular account passwords will be blocked by Google.")

        smtp_server, ok = QInputDialog.getText(self, "SMTP", "SMTP server:", text="smtp.gmail.com")
        if not ok:
            return
        smtp_port, ok = QInputDialog.getInt(self, "SMTP", "Port:", value=587)
        if not ok:
            return
        sender, ok = QInputDialog.getText(self, "Sender", "Sender email:")
        if not ok:
            return
        password, ok = QInputDialog.getText(self, "Password", "Sender password (or app password):", echo=QLineEdit.Password)
        if not ok:
            return

        failed = []
        sent = 0
        for row in recipients:
            to = row["Email"]
            if not to or str(to).strip() == "":
                failed.append(("(no email)", "missing email"))
                continue
            subject = "Fee Payment Reminder"
            body = f"Hello {row['Name']},\n\nThis is a gentle reminder that your pending fee balance is{row['Balance']}.\n"
            if row.get("Due Date"):
                body += f"Due Date: {row['Due Date']}\n"
            body += f"Kindly make the payment at the earliest to avoid any late charges.\nIf you have already completed the payment, please disregard this message.\nThank you for your prompt attention.\nPlease pay at the earliest.\n\nRegards,\nPyLinX Hub"
            msg = MIMEText(body)
            msg["Subject"] = subject
            msg["From"] = sender
            msg["To"] = to
            try:
                server = smtplib.SMTP(smtp_server, smtp_port, timeout=20)
                server.ehlo()
                server.starttls()
                server.ehlo()
                server.login(sender, password)
                server.sendmail(sender, [to], msg.as_string())
                server.quit()
                sent += 1
            except Exception as e:
                # record exact exception string so user can see what's wrong
                failed.append((to, str(e)))

        info = f"Sent: {sent}"
        if failed:
            info += f", Failed: {len(failed)}"
        # show details if any failed
        if failed:
            details = "\n".join([f"{t}: {err}" for t, err in failed[:10]])
            QMessageBox.warning(self, "Email result", info + "\n\nSome failures:\n" + details)
        else:
            QMessageBox.information(self, "Email result", info)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = FeeManagerApp()
    w.show()
    sys.exit(app.exec_())
