import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from datetime import datetime
import pandas as pd
import shutil


# ==========================================
# CONFIGURATION
# ==========================================

PATIENT_COLUMN = "PATIENT NAME"
PATIENT_RESPONSIBILITY_COLUMN = "PATIENT RESPONSIBILITY"
CHECKDATE_COLUMN = "PAYMENT/CHECK DATE"

MASTER_PATIENT_COLUMN = "Patient Name"
MASTER_PATIENT_RESPONSIBILITY_COLUMN = "Patient Responsiblity"
MASTER_CHECKDATE_COLUMN = "Checkdate"

# Master file location
APP_DIRECTORY = (
    Path.home() / "Documents" / "PatientReconciliation"
)

MASTER_FILE = APP_DIRECTORY / "Master.xlsx"

BACKUP_DIRECTORY = APP_DIRECTORY / "Backups"


# ==========================================
# DATA PROCESSING
# ==========================================

def load_excel_file(file_path):

    df = pd.read_excel(file_path, skiprows=1)

    # Ensure required columns exist
    required_columns = [PATIENT_COLUMN, PATIENT_RESPONSIBILITY_COLUMN, CHECKDATE_COLUMN]

    for column in required_columns:
        if column not in df.columns:
            raise ValueError(
                f"Missing required column: {column}"
            )

    # Keep only necessary columns
    df = df[required_columns].copy()

    # Remove empty client names
    df = df.dropna(subset=[PATIENT_COLUMN])

    # Normalize client names
    df[PATIENT_COLUMN] = (
        df[PATIENT_COLUMN]
        .astype(str)
        .str.strip()
    )

    # Remove blank names
    df = df[df[PATIENT_COLUMN] != ""]

    # Convert values to numeric
    df[PATIENT_RESPONSIBILITY_COLUMN] = pd.to_numeric(
        df[PATIENT_RESPONSIBILITY_COLUMN],
        errors="coerce"
    )

    if df[PATIENT_RESPONSIBILITY_COLUMN].isna().any():
        raise ValueError(
            f"One or more clients have missing or "
            f"non-numeric values in {PATIENT_RESPONSIBILITY_COLUMN}."
        )

    # Create normalized matching key
    df["_client_key"] = (
        df[PATIENT_COLUMN].str.casefold()
    )

    # Remove duplicate clients
    df = df.drop_duplicates(
        subset=["_client_key"],
        keep="first"
    )

    return df

def load_master_excel_file(file_path):

    df = pd.read_excel(file_path)

    # Ensure required columns exist
    required_columns = [PATIENT_COLUMN, PATIENT_RESPONSIBILITY_COLUMN, CHECKDATE_COLUMN]

    for column in required_columns:
        if column not in df.columns:
            raise ValueError(
                f"Missing required column: {column}"
            )

    # Keep only necessary columns
    df = df[required_columns].copy()

    # Remove empty client names
    df = df.dropna(subset=[PATIENT_COLUMN])

    # Normalize client names
    df[PATIENT_COLUMN] = (
        df[PATIENT_COLUMN]
        .astype(str)
        .str.strip()
    )

    # Remove blank names
    df = df[df[PATIENT_COLUMN] != ""]

    # Convert values to numeric
    df[PATIENT_RESPONSIBILITY_COLUMN] = pd.to_numeric(
        df[PATIENT_RESPONSIBILITY_COLUMN],
        errors="coerce"
    )

    if df[PATIENT_RESPONSIBILITY_COLUMN].isna().any():
        raise ValueError(
            f"One or more clients have missing or "
            f"non-numeric values in {PATIENT_RESPONSIBILITY_COLUMN}."
        )

    # Create normalized matching key
    df["_client_key"] = (
        df[PATIENT_COLUMN].str.casefold()
    )

    # Remove duplicate clients
    df = df.drop_duplicates(
        subset=["_client_key"],
        keep="first"
    )

    return df


def create_master_if_missing():

    APP_DIRECTORY.mkdir(
        parents=True,
        exist_ok=True
    )

    BACKUP_DIRECTORY.mkdir(
        parents=True,
        exist_ok=True
    )

    if not MASTER_FILE.exists():

        empty_master = pd.DataFrame(
            columns=[PATIENT_COLUMN, PATIENT_RESPONSIBILITY_COLUMN, CHECKDATE_COLUMN]
        )

        empty_master.to_excel(
            MASTER_FILE,
            index=False
        )


def backup_master():

    if MASTER_FILE.exists():

        timestamp = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )

        backup_file = (
            BACKUP_DIRECTORY /
            f"Master_{timestamp}.xlsx"
        )

        shutil.copy2(
            MASTER_FILE,
            backup_file
        )


def save_master(master):

    output = master[
        [PATIENT_COLUMN, PATIENT_RESPONSIBILITY_COLUMN, CHECKDATE_COLUMN]
    ].copy()

    output.to_excel(
        MASTER_FILE,
        index=False
    )


def process_files(old_file, new_file):

    create_master_if_missing()

    # Load existing master
    master = load_master_excel_file(MASTER_FILE)

    master_lookup = {
        row["_client_key"]: {
            "patient_name": row[PATIENT_COLUMN],
            "patient_responsibility": row[PATIENT_RESPONSIBILITY_COLUMN],
            "checkdate": row[CHECKDATE_COLUMN]
        }
        for _, row in master.iterrows()
    }

    # ======================================
    # STEP 1: PROCESS OLD FILE IF PROVIDED
    # ======================================

    if old_file:

        old_data = load_excel_file(old_file)

        for _, row in old_data.iterrows():

            key = row["_client_key"]

            master_lookup[key] = {
                "patient_name": row[PATIENT_COLUMN],
                "patient_responsibility": row[PATIENT_RESPONSIBILITY_COLUMN],
                "checkdate": row[CHECKDATE_COLUMN]
            }

    # ======================================
    # STEP 2: LOAD NEW FILE
    # ======================================

    new_data = load_excel_file(new_file)

    # ======================================
    # STEP 3: COMPARE NEW DATA AGAINST MASTER
    # ======================================

    changes = []

    for _, row in new_data.iterrows():

        key = row["_client_key"]

        patient_name = row[PATIENT_COLUMN]

        new_patient_responsibility = row[PATIENT_RESPONSIBILITY_COLUMN]

        new_checkdate = row[CHECKDATE_COLUMN]

        # Client does not exist in master
        if key not in master_lookup:

            changes.append({
                "Patient Name": patient_name,
                "Previous Value": None,
                "New Patient Responsibility": new_patient_responsibility,
                "Difference": None,
                "Status": "NEW"
            })

        else:

            old_value = master_lookup[key]["patient_responsibility"]

            # Check whether value changed
            if old_value != new_patient_responsibility:

                changes.append({
                    "Patient Name": patient_name,
                    "Previous Patient Responsibility": old_value,
                    "New Patient Responsibility": new_patient_responsibility,
                    "Difference": new_patient_responsibility - old_value,
                    "Status": "CHANGED"
                })

        # Update master with newest value
        master_lookup[key] = {
            "patient_name": patient_name,
            "patient_responsibility": new_patient_responsibility,
            "checkdate": new_checkdate

        }

    # ======================================
    # STEP 4: GENERATE RESULTS
    # ======================================

    results = pd.DataFrame(
        changes,
        columns=[
            "Patient Name",
            "Previous Patient Responsibility",
            "New Patient Responsibility",
            "Difference",
            "Status"
        ]
    )

    # ======================================
    # STEP 5: UPDATE MASTER
    # ======================================

    updated_master = pd.DataFrame([
        {
            PATIENT_COLUMN: record["patient_name"],
            PATIENT_RESPONSIBILITY_COLUMN: record["patient_responsibility"],
            CHECKDATE_COLUMN: record["checkdate"]
        }
        for record in master_lookup.values()
    ])

    return results, updated_master


# ==========================================
# TKINTER APPLICATION
# ==========================================

class ClientReconciliationApp:

    def __init__(self, root):

        self.root = root

        self.root.title(
            "Patient Reconciliation Tool"
        )

        self.root.geometry("600x400")

        self.old_file = None
        self.new_file = None

        self.build_interface()

    def build_interface(self):

        # Title
        tk.Label(
            self.root,
            text="Patient Reconciliation Tool",
            font=("Arial", 18, "bold")
        ).pack(pady=20)

        # Description
        tk.Label(
            self.root,
            text="Compare most recent patient responsibility values against the master file."
        ).pack(pady=5)

        # Old file selection
        tk.Button(
            self.root,
            text="Select Old Excel File (Only select this the first time using this application)",
            width=60,
            command=self.select_old_file
        ).pack(pady=10)

        self.old_file_label = tk.Label(
            self.root,
            text="No old file selected"
        )

        self.old_file_label.pack()

        # New file selection
        tk.Button(
            self.root,
            text="Select New Excel File",
            width=35,
            command=self.select_new_file
        ).pack(pady=10)

        self.new_file_label = tk.Label(
            self.root,
            text="No new file selected"
        )

        self.new_file_label.pack()

        # Process button
        tk.Button(
            self.root,
            text="Process Files",
            width=25,
            height=2,
            command=self.process
        ).pack(pady=25)

        # Status
        self.status_label = tk.Label(
            self.root,
            text="Ready",
            font=("Arial", 10)
        )

        self.status_label.pack(pady=10)

        # Master location
        tk.Label(
            self.root,
            text=f"Master File: {MASTER_FILE}",
            font=("Arial", 8),
            wraplength=550
        ).pack(side=tk.BOTTOM, pady=10)

    def select_old_file(self):

        file_path = filedialog.askopenfilename(
            title="Select Old Excel File",
            filetypes=[
                ("Excel Files", "*.xlsx *.xls")
            ]
        )

        if file_path:

            self.old_file = file_path

            self.old_file_label.config(
                text=Path(file_path).name
            )

    def select_new_file(self):

        file_path = filedialog.askopenfilename(
            title="Select New Excel File",
            filetypes=[
                ("Excel Files", "*.xlsx *.xls")
            ]
        )

        if file_path:

            self.new_file = file_path

            self.new_file_label.config(
                text=Path(file_path).name
            )

    def process(self):

        if not self.new_file:

            messagebox.showerror(
                "Missing File",
                "Please select a new Excel file."
            )

            return

        self.status_label.config(
            text="Processing..."
        )

        self.root.update_idletasks()

        try:

            # Process data
            results, updated_master = process_files(
                self.old_file,
                self.new_file
            )

            # Ask where to save report
            report_path = filedialog.asksaveasfilename(
                title="Save Comparison Report",
                defaultextension=".xlsx",
                initialfile="Comparison_Report.xlsx",
                filetypes=[
                    ("Excel Files", "*.xlsx")
                ]
            )

            if not report_path:

                self.status_label.config(
                    text="Processing cancelled. Master unchanged."
                )

                return

            # Back up existing master
            # backup_master()

            # Save updated master
            save_master(updated_master)

            # Save comparison report
            results.to_excel(
                report_path,
                index=False
            )

            # Calculate statistics
            changed_count = len(
                results[results["Status"] == "CHANGED"]
            )

            new_count = len(
                results[results["Status"] == "NEW"]
            )

            self.status_label.config(
                text="Processing completed successfully!"
            )

            messagebox.showinfo(
                "Processing Complete",
                f"Comparison completed successfully!\n\n"
                f"Clients with changed values: {changed_count}\n"
                f"New clients added: {new_count}\n\n"
                f"Total clients in master: {len(updated_master)}\n\n"
                f"Master file updated successfully."
            )

        except Exception as error:

            self.status_label.config(
                text="An error occurred."
            )

            messagebox.showerror(
                "Processing Error",
                str(error)
            )


# ==========================================
# APPLICATION ENTRY POINT
# ==========================================

if __name__ == "__main__":

    root = tk.Tk()

    app = ClientReconciliationApp(root)

    root.mainloop()