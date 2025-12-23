"""
Template Generator GUI
----------------------
Standalone Python GUI application using tkinter (built-in).
Generates template files, JSONs, and specimen PDFs from .docx and rules.xlsx files.

To create EXE:
    pip install pyinstaller openpyxl pandas reportlab pypdf docx2pdf
    pyinstaller --onefile --windowed --name="TemplateGenerator" template_generator_gui.py
"""

import json
import os
import re
import shutil
import uuid
from dataclasses import dataclass
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from tkinter import (
    Tk,
    Frame,
    Label,
    Button,
    Listbox,
    Scrollbar,
    Text,
    filedialog,
    messagebox,
    BOTH,
    LEFT,
    RIGHT,
    Y,
    END,
    VERTICAL,
    HORIZONTAL,
)
from tkinter import ttk
import tkinter as tk
from typing import List, Sequence, Iterable

import pandas as pd
from docx2pdf import convert as docx2pdf_convert
from pypdf import PdfReader, PdfWriter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas


@dataclass
class TemplateContext:
    template_path: Path
    template_guid: uuid.UUID
    document_guid: uuid.UUID
    matching_rule: dict
    copied_docx: Path


# ============================================================================
# UTILITY FUNCTIONS (same as original script)
# ============================================================================


def slug(text: str) -> str:
    return re.sub(r"[\s_\-]+", "", text or "").lower()


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def parse_list_field(value) -> List[str]:
    if pd.isna(value):
        return []
    text = str(value).strip()
    if not text:
        return []
    return [item.strip() for item in text.split(",") if item.strip()]


def now_timestamp() -> tuple[str, float]:
    dt = datetime.now(timezone.utc)
    iso_value = dt.isoformat(timespec="milliseconds").replace("+00:00", "Z")
    return iso_value, dt.timestamp()


def format_date(value) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.strftime("%Y-%m-%d")
    text = str(value)
    return text[:10]


def load_rules(rules_path: Path) -> List[dict]:
    required_columns = [
        "FormNumber",
        "FormTitle",
        "EditionDate",
        "EffectiveDate",
        "ExpirationDate",
        "DisplaySequence",
        "USStateCode",
        "FormRequired",
        "FormType",
        "Product",
        "LegalEntity",
        "TransactionStatus",
    ]
    rules_df = pd.read_excel(
        rules_path, parse_dates=["EffectiveDate", "ExpirationDate"]
    )
    missing = [col for col in required_columns if col not in rules_df.columns]
    if missing:
        raise ValueError(f"Missing columns in rules.xlsx: {', '.join(missing)}")

    rules_df["EffectiveDate"] = rules_df["EffectiveDate"].apply(format_date)
    rules_df["ExpirationDate"] = rules_df["ExpirationDate"].apply(format_date)
    rules_df["TransactionStatus"] = rules_df["TransactionStatus"].apply(
        parse_list_field
    )
    rules_df["LegalEntity"] = rules_df["LegalEntity"].apply(parse_list_field)
    rules_df["Product"] = rules_df["Product"].apply(parse_list_field)

    return json.loads(rules_df.to_json(orient="records", default_handler=str))


def find_matching_rule(template_name: str, rules: Sequence[dict]) -> dict:
    sanitized = slug(template_name)
    for rule in rules:
        form_number = slug(str(rule.get("FormNumber", "")))
        if form_number and form_number in sanitized:
            return rule
    raise ValueError(f"No matching rule found for template '{template_name}'")


def copy_template_docx(
    source: Path, destination_dir: Path, template_guid: uuid.UUID
) -> Path:
    ensure_dir(destination_dir)
    destination = destination_dir / f"{template_guid}{source.suffix}"
    shutil.copy2(source, destination)
    return destination


def generate_file_json(
    template_guid: uuid.UUID, form_title: str, template_name: str, destination_dir: Path
) -> Path:
    # Embedded FILE_TEMPLATE
    data = {
        "id": "",
        "Partition": "fileTemplates",
        "SystemInfo": {
            "CreatedOn": {"Timestamp": {"Value": "", "Epoch": 0}, "UserName": None},
            "UpdatedOn": None,
            "DocumentId": "",
            "DocumentVersionId": "",
            "DocumentPartition": "fileTemplates",
            "DocumentOrganization": "hudsoninsgroup",
        },
        "Content": {
            "FileName": "",
            "FileSize": 32768,
            "AzureBlobStorageContainer": "templates",
            "AzureBlobStorageFileName": "",
        },
        "Organization": "hudsoninsgroup",
    }

    guid_text = str(template_guid)
    data["id"] = guid_text
    data["SystemInfo"]["DocumentId"] = guid_text
    data["SystemInfo"]["DocumentVersionId"] = guid_text
    data["Content"]["FileName"] = f"{template_name}.docx"
    data["Content"][
        "AzureBlobStorageFileName"
    ] = f"hudsoninsgroup0001/01/01/{guid_text}.docx"

    now_iso, now_epoch = now_timestamp()
    data["SystemInfo"]["CreatedOn"]["Timestamp"]["Value"] = now_iso
    data["SystemInfo"]["CreatedOn"]["Timestamp"]["Epoch"] = int(now_epoch)

    out_name = f"File_{form_title}_{guid_text}.json".replace(" ", "_")
    ensure_dir(destination_dir)
    out_path = destination_dir / out_name
    with out_path.open("w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)
    return out_path


def generate_document_json(
    template_guid: uuid.UUID,
    document_guid: uuid.UUID,
    rule: dict,
    destination_dir: Path,
    all_states: Iterable[str],
    profile_data: dict = None,
) -> Path:
    # Embedded DOCUMENT_TEMPLATE
    data = {
        "id": "",
        "Partition": "documentTemplates",
        "SystemInfo": {
            "DocumentType": "DocumentTemplate",
            "SchemaId": "",
            "PreviousVersions": None,
            "Version": 1,
            "CreatedOn": {
                "Timestamp": {"Value": "", "Epoch": 0},
                "UserName": "ODYSSEYRE\\NZamorin",
            },
            "UpdatedOn": None,
            "DocumentId": "",
            "DocumentVersionId": "",
            "DocumentPartition": "documentTemplates",
            "DocumentOrganization": "hudsoninsgroup",
        },
        "Content": {
            "LinkedGUID": "",
            "CommonGUID": "",
            "IsCurrentVersion": "True",
            "SchemaID": "",
            "sysCreatedBy": "ODYSSEYRE\\NZamorin",
            "InRuleOnSave": "8992AFD0-0893-4053-BBE1-52EE54A7BE5B",
            "TemplateFile": "",
            "DocumentSchema": "45E8B1EC-5CCF-4021-A411-1DC0E415995F",
            "TextPolicyFormNumber": "",
            "TextOutputTemplateTitle": "",
            "TemplateCriteria": [
                {
                    "TransactionStatus": [],
                    "TemplateStartDate": "",
                    "TemplateEndDate": "",
                    "LegalEntity": [],
                    "Product": [],
                    "PolicyState": [],
                    "ProgramID": [],
                    "ScribeProcessID": "1",
                    "OutputTemplatePrintOrder": "",
                    "OutputTemplateMandatory": "Optional",
                    "TextOutputTemplateSpecimenUrl": "",
                    "OutputTemplateFormType": "Static",
                    "TextFormCategoryConcatList": "-1",
                }
            ],
            "sysCreatedTS": "",
        },
        "Organization": "hudsoninsgroup",
    }

        # ============================================
   
    # APPLY PROFILE DATA (ProgramID, SchemaID, etc)
    # ============================================
    if profile_data:
        # SystemInfo fields
        data["SystemInfo"]["SchemaId"] = profile_data.get("SchemaId", "")
        data["SystemInfo"]["InRuleOnSave"] = profile_data.get("InRuleOnSave", "")
        data["SystemInfo"]["DocumentSchema"] = profile_data.get("DocumentSchema", "")

        # Content fields
        data["Content"]["ProgramID"] = profile_data.get("ProgramID", "")
        data["Content"]["LinkedGUID"] = profile_data.get("LinkedGUID", "")
        data["Content"]["SchemaID"] = profile_data.get("SchemaId", "")
        data["Content"]["InRuleOnSave"] = profile_data.get("InRuleOnSave", "")
        data["Content"]["DocumentSchema"] = profile_data.get("DocumentSchema", "")

    template_guid_text = str(template_guid)
    document_guid_text = str(document_guid)
    form_number = str(rule.get("FormNumber", ""))
    form_title = str(rule.get("FormTitle", ""))
    edition_date = str(rule.get("EditionDate", "")).strip()
    policy_state_raw = rule.get("USStateCode", "All")
    policy_states = (
        list(all_states)
        if str(policy_state_raw).lower() == "all"
        else [policy_state_raw]
    )
    output_template_mandatory = (
        "Mandatory" if str(rule.get("FormRequired", "0")) == "1" else "Optional"
    )
    form_type = str(rule.get("FormType", "S"))
    output_template_form_type = "Dynamic" if form_type.upper() == "D" else "Static"
    print_order = str(rule.get("DisplaySequence", ""))

    data["id"] = document_guid_text
    data["SystemInfo"]["DocumentId"] = document_guid_text
    data["SystemInfo"]["DocumentVersionId"] = document_guid_text
    data["Content"]["CommonGUID"] = document_guid_text
    data["Content"]["TemplateFile"] = template_guid_text
    data["Content"]["TextPolicyFormNumber"] = f"{form_number} {edition_date}".strip()
    data["Content"]["TextOutputTemplateTitle"] = form_title

    now_iso, now_epoch = now_timestamp()
    data["SystemInfo"]["CreatedOn"]["Timestamp"]["Value"] = now_iso
    data["SystemInfo"]["CreatedOn"]["Timestamp"]["Epoch"] = int(now_epoch)
    data["Content"]["sysCreatedTS"] = now_iso

    criteria = data["Content"]["TemplateCriteria"][0]
    criteria["TemplateStartDate"] = format_date(rule.get("EffectiveDate", ""))
    criteria["TemplateEndDate"] = format_date(rule.get("ExpirationDate", ""))
    criteria["PolicyState"] = policy_states
    criteria["OutputTemplatePrintOrder"] = print_order
    criteria["OutputTemplateMandatory"] = output_template_mandatory
    criteria["OutputTemplateFormType"] = output_template_form_type
    criteria["TextOutputTemplateSpecimenUrl"] = (
        f"https://hudsonfiles.hudsonportal.com/BL/{template_guid_text}.pdf"
    )
    criteria["TransactionStatus"] = rule.get("TransactionStatus", [])
    criteria["LegalEntity"] = rule.get("LegalEntity", [])
    criteria["Product"] = rule.get("Product", [])

    out_name = f"Document_{form_title}_{document_guid_text}.json".replace(" ", "_")
    ensure_dir(destination_dir)
    out_path = destination_dir / out_name
    with out_path.open("w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)
    return out_path


def register_arial_font() -> str:
    font_name = "Arial"
    try:
        font_paths = [
            Path("C:/Windows/Fonts/arial.ttf"),
            Path("/usr/share/fonts/truetype/msttcorefonts/Arial.ttf"),
            Path("/Library/Fonts/Arial.ttf"),
        ]
        for path in font_paths:
            if path.exists():
                pdfmetrics.registerFont(TTFont(font_name, str(path)))
                return font_name
    except Exception:
        pass
    return "Helvetica"


def build_watermark_page(text: str, width: float, height: float) -> BytesIO:
    buffer = BytesIO()
    font_name = register_arial_font()
    c = canvas.Canvas(buffer, pagesize=(width, height))
    c.saveState()
    c.translate(width / 2, height / 2)
    c.rotate(45)
    c.setFont(font_name, 125)
    c.setFillColorRGB(0.85, 0.85, 0.85)
    c.drawCentredString(0, 0, text)
    c.restoreState()
    c.save()
    buffer.seek(0)
    return buffer


def apply_watermark(
    base_pdf: Path, output_pdf: Path, watermark_text: str = "Specimen"
) -> None:
    reader = PdfReader(str(base_pdf))
    writer = PdfWriter()
    for page in reader.pages:
        width = float(page.mediabox.width)
        height = float(page.mediabox.height)
        wm_pdf = PdfReader(build_watermark_page(watermark_text, width, height))
        watermark_page = wm_pdf.pages[0]
        watermark_page.merge_page(page)
        writer.add_page(watermark_page)
    with output_pdf.open("wb") as f:
        writer.write(f)


def generate_specimen(
    docx_path: Path, output_pdf: Path, watermark_text: str = "Specimen"
) -> None:
    with TemporaryDirectory() as tmp_dir:
        tmp_dir_path = Path(tmp_dir)
        docx2pdf_convert(str(docx_path), str(tmp_dir_path))
        base_pdf = tmp_dir_path / f"{docx_path.stem}.pdf"
        if not base_pdf.exists():
            raise FileNotFoundError(
                f"docx2pdf did not produce expected PDF at {base_pdf}"
            )
        apply_watermark(base_pdf, output_pdf, watermark_text=watermark_text)


def build_specimen_report(destination_dir: Path) -> tuple[Path, Path]:
    ensure_dir(destination_dir)
    records = []
    for docx_file in sorted(destination_dir.glob("*.docx")):
        name = docx_file.stem
        records.append(
            {
                "FileName": name + ".docx",
                "SpecimenURL": f"https://hudsonfiles.hudsonportal.com/BL/{name}.pdf",
            }
        )
    json_path = destination_dir / "Specimen_Report.json"
    xlsx_path = destination_dir / "Specimen_Report.xlsx"
    with json_path.open("w", encoding="utf-8") as f:
        json.dump(records, f, indent=4)
    pd.DataFrame(records).to_excel(xlsx_path, index=False)
    return json_path, xlsx_path


# Embedded AllStates
ALL_STATES = [
    "AA",
    "AE",
    "AK",
    "AL",
    "AP",
    "AR",
    "AS",
    "AZ",
    "DC",
    "DE",
    "FL",
    "FM",
    "GA",
    "GU",
    "HI",
    "IA",
    "ID",
    "IL",
    "IN",
    "KS",
    "KY",
    "LA",
    "MA",
    "MD",
    "ME",
    "MH",
    "MI",
    "MN",
    "MO",
    "MP",
    "MS",
    "MT",
    "NC",
    "ND",
    "NE",
    "NH",
    "NJ",
    "NV",
    "NY",
    "OH",
    "OK",
    "OR",
    "PA",
    "PR",
    "PW",
    "RI",
    "NM",
    "SC",
    "SD",
    "TN",
    "TX",
    "UT",
    "VA",
    "VI",
    "VT",
    "WA",
    "WI",
    "WV",
    "WY",
    "CA",
    "CO",
    "CT",
]


# ============================================================================
# GUI APPLICATION
# ============================================================================


class TemplateGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Template Generator")
        self.root.geometry("800x600")

        self.docx_files = []
        self.rules_file = None
        self.output_dir = None

        # Profile management
        self.profiles_file = Path("profiles.json")
        self.load_profiles()

        self.setup_ui()

    def setup_ui(self):
        # Title
        title = Label(self.root, text="Template Generator", font=("Arial", 16, "bold"))
        title.pack(pady=10)

        # Profile Section
        frame_profile = Frame(self.root, relief="solid", borderwidth=1)
        frame_profile.pack(fill="x", padx=10, pady=10)

        Label(frame_profile, text="Profile:", font=("Arial", 10, "bold")).pack(
            side=LEFT, padx=5
        )

        self.profile_var = tk.StringVar()
        self.profile_dropdown = ttk.Combobox(
            frame_profile, textvariable=self.profile_var, state="readonly", width=30
        )
        self.profile_dropdown.pack(side=LEFT, padx=5)
        self.refresh_profile_dropdown()

        Button(
            frame_profile,
            text="Create",
            command=self.create_profile,
            bg="#4CAF50",
            fg="white",
        ).pack(side=LEFT, padx=2)
        Button(
            frame_profile,
            text="Edit",
            command=self.edit_profile,
            bg="#2196F3",
            fg="white",
        ).pack(side=LEFT, padx=2)
        Button(
            frame_profile,
            text="Delete",
            command=self.delete_profile,
            bg="#f44336",
            fg="white",
        ).pack(side=LEFT, padx=2)

        # DOCX Files Section
        frame_docx = Frame(self.root)
        frame_docx.pack(fill=BOTH, expand=True, padx=10, pady=5)

        Label(frame_docx, text="DOCX Templates:", font=("Arial", 10, "bold")).pack(
            anchor="w"
        )

        btn_frame = Frame(frame_docx)
        btn_frame.pack(fill="x", pady=5)

        Button(btn_frame, text="Add DOCX Files", command=self.add_docx_files).pack(
            side=LEFT, padx=5
        )
        Button(btn_frame, text="Clear", command=self.clear_docx_files).pack(side=LEFT)

        list_frame = Frame(frame_docx)
        list_frame.pack(fill=BOTH, expand=True)

        scrollbar = Scrollbar(list_frame, orient=VERTICAL)
        self.docx_listbox = Listbox(list_frame, yscrollcommand=scrollbar.set, height=8)
        scrollbar.config(command=self.docx_listbox.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.docx_listbox.pack(side=LEFT, fill=BOTH, expand=True)

        # Rules File Section
        frame_rules = Frame(self.root)
        frame_rules.pack(fill="x", padx=10, pady=5)

        Label(frame_rules, text="Rules File (Excel):", font=("Arial", 10, "bold")).pack(
            anchor="w"
        )

        rules_btn_frame = Frame(frame_rules)
        rules_btn_frame.pack(fill="x", pady=5)

        Button(
            rules_btn_frame, text="Select Rules.xlsx", command=self.select_rules_file
        ).pack(side=LEFT, padx=5)
        self.rules_label = Label(rules_btn_frame, text="No file selected", fg="gray")
        self.rules_label.pack(side=LEFT, padx=10)

        # Output Directory Section
        frame_output = Frame(self.root)
        frame_output.pack(fill="x", padx=10, pady=5)

        Label(frame_output, text="Output Directory:", font=("Arial", 10, "bold")).pack(
            anchor="w"
        )

        output_btn_frame = Frame(frame_output)
        output_btn_frame.pack(fill="x", pady=5)

        Button(
            output_btn_frame,
            text="Select Output Folder",
            command=self.select_output_dir,
        ).pack(side=LEFT, padx=5)
        self.output_label = Label(
            output_btn_frame, text="No folder selected", fg="gray"
        )
        self.output_label.pack(side=LEFT, padx=10)

        # Generate Button
        Button(
            self.root,
            text="Generate Templates",
            command=self.generate_templates,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 12, "bold"),
            height=2,
        ).pack(pady=15, padx=10, fill="x")

        # Log Area
        Label(self.root, text="Log:", font=("Arial", 10, "bold")).pack(
            anchor="w", padx=10
        )

        log_frame = Frame(self.root)
        log_frame.pack(fill=BOTH, expand=True, padx=10, pady=5)

        log_scrollbar = Scrollbar(log_frame)
        log_scrollbar.pack(side=RIGHT, fill=Y)

        self.log_text = Text(
            log_frame, height=8, yscrollcommand=log_scrollbar.set, state="disabled"
        )
        self.log_text.pack(side=LEFT, fill=BOTH, expand=True)
        log_scrollbar.config(command=self.log_text.yview)

    def load_profiles(self):
        """Load profiles from JSON file"""
        if self.profiles_file.exists():
            with open(self.profiles_file, "r") as f:
                self.profiles = json.load(f)
        else:
            # Default profile
            self.profiles = {
                "Default": {
                    "SchemaId": "8e6f9af8-4ce2-423c-bf87-380eee1f41e3",
                    "ProgramID": 212,
                    "LinkedGUID": "8E6F9AF8-4CE2-423C-BF87-380EEE1F41E3",
                    "DocumentSchema": "45E8B1EC-5CCF-4021-A411-1DC0E415995F",
                    "InRuleOnSave": "8992AFD0-0893-4053-BBE1-52EE54A7BE5B",
                }
            }
            self.save_profiles()

    def save_profiles(self):
        """Save profiles to JSON file"""
        with open(self.profiles_file, "w") as f:
            json.dump(self.profiles, f, indent=4)

    def refresh_profile_dropdown(self):
        """Refresh the profile dropdown with current profiles"""
        profile_names = list(self.profiles.keys())
        self.profile_dropdown["values"] = profile_names
        if profile_names:
            self.profile_dropdown.current(0)

    def create_profile(self):
        """Open dialog to create a new profile"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Create Profile")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()

        Label(dialog, text="Profile Name:", font=("Arial", 10, "bold")).grid(
            row=0, column=0, sticky="w", padx=10, pady=5
        )
        name_entry = tk.Entry(dialog, width=40)
        name_entry.grid(row=0, column=1, padx=10, pady=5)

        Label(dialog, text="SchemaId:", font=("Arial", 10)).grid(
            row=1, column=0, sticky="w", padx=10, pady=5
        )
        schema_entry = tk.Entry(dialog, width=40)
        schema_entry.insert(0, "8e6f9af8-4ce2-423c-bf87-380eee1f41e3")
        schema_entry.grid(row=1, column=1, padx=10, pady=5)

        Label(dialog, text="ProgramID:", font=("Arial", 10)).grid(
            row=2, column=0, sticky="w", padx=10, pady=5
        )
        program_entry = tk.Entry(dialog, width=40)
        program_entry.insert(0, "212")
        program_entry.grid(row=2, column=1, padx=10, pady=5)

        Label(dialog, text="LinkedGUID:", font=("Arial", 10)).grid(
            row=3, column=0, sticky="w", padx=10, pady=5
        )
        linked_entry = tk.Entry(dialog, width=40)
        linked_entry.insert(0, "8E6F9AF8-4CE2-423C-BF87-380EEE1F41E3")
        linked_entry.grid(row=3, column=1, padx=10, pady=5)

        Label(dialog, text="DocumentSchema:", font=("Arial", 10)).grid(
            row=4, column=0, sticky="w", padx=10, pady=5
        )
        doc_schema_entry = tk.Entry(dialog, width=40)
        doc_schema_entry.insert(0, "45E8B1EC-5CCF-4021-A411-1DC0E415995F")
        doc_schema_entry.grid(row=4, column=1, padx=10, pady=5)

        Label(dialog, text="InRuleOnSave:", font=("Arial", 10)).grid(
            row=5, column=0, sticky="w", padx=10, pady=5
        )
        rule_entry = tk.Entry(dialog, width=40)
        rule_entry.insert(0, "8992AFD0-0893-4053-BBE1-52EE54A7BE5B")
        rule_entry.grid(row=5, column=1, padx=10, pady=5)

        def save_profile():
            name = name_entry.get().strip()
            if not name:
                messagebox.showerror("Error", "Profile name cannot be empty")
                return
            if name in self.profiles:
                messagebox.showerror("Error", "Profile name already exists")
                return

            try:
                program_id = int(program_entry.get().strip())
            except ValueError:
                messagebox.showerror("Error", "ProgramID must be a number")
                return

            self.profiles[name] = {
                "SchemaId": schema_entry.get().strip(),
                "ProgramID": program_id,
                "LinkedGUID": linked_entry.get().strip(),
                "DocumentSchema": doc_schema_entry.get().strip(),
                "InRuleOnSave": rule_entry.get().strip(),
            }
            self.save_profiles()
            self.refresh_profile_dropdown()
            dialog.destroy()
            messagebox.showinfo("Success", f"Profile '{name}' created successfully")

        Button(
            dialog, text="Save", command=save_profile, bg="#4CAF50", fg="white"
        ).grid(row=6, column=0, columnspan=2, pady=20)

    def edit_profile(self):
        """Open dialog to edit the selected profile"""
        selected_profile = self.profile_var.get()
        if not selected_profile:
            messagebox.showwarning("Warning", "Please select a profile to edit")
            return

        if selected_profile not in self.profiles:
            messagebox.showerror("Error", "Selected profile not found")
            return

        profile_data = self.profiles[selected_profile]

        dialog = tk.Toplevel(self.root)
        dialog.title(f"Edit Profile: {selected_profile}")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()

        Label(
            dialog, text=f"Editing: {selected_profile}", font=("Arial", 12, "bold")
        ).grid(row=0, column=0, columnspan=2, pady=10)

        Label(dialog, text="SchemaId:", font=("Arial", 10)).grid(
            row=1, column=0, sticky="w", padx=10, pady=5
        )
        schema_entry = tk.Entry(dialog, width=40)
        schema_entry.insert(0, profile_data.get("SchemaId", ""))
        schema_entry.grid(row=1, column=1, padx=10, pady=5)

        Label(dialog, text="ProgramID:", font=("Arial", 10)).grid(
            row=2, column=0, sticky="w", padx=10, pady=5
        )
        program_entry = tk.Entry(dialog, width=40)
        program_entry.insert(0, str(profile_data.get("ProgramID", "")))
        program_entry.grid(row=2, column=1, padx=10, pady=5)

        Label(dialog, text="LinkedGUID:", font=("Arial", 10)).grid(
            row=3, column=0, sticky="w", padx=10, pady=5
        )
        linked_entry = tk.Entry(dialog, width=40)
        linked_entry.insert(0, profile_data.get("LinkedGUID", ""))
        linked_entry.grid(row=3, column=1, padx=10, pady=5)

        Label(dialog, text="DocumentSchema:", font=("Arial", 10)).grid(
            row=4, column=0, sticky="w", padx=10, pady=5
        )
        doc_schema_entry = tk.Entry(dialog, width=40)
        doc_schema_entry.insert(0, profile_data.get("DocumentSchema", ""))
        doc_schema_entry.grid(row=4, column=1, padx=10, pady=5)

        Label(dialog, text="InRuleOnSave:", font=("Arial", 10)).grid(
            row=5, column=0, sticky="w", padx=10, pady=5
        )
        rule_entry = tk.Entry(dialog, width=40)
        rule_entry.insert(0, profile_data.get("InRuleOnSave", ""))
        rule_entry.grid(row=5, column=1, padx=10, pady=5)

        def save_changes():
            try:
                program_id = int(program_entry.get().strip())
            except ValueError:
                messagebox.showerror("Error", "ProgramID must be a number")
                return

            self.profiles[selected_profile] = {
                "SchemaId": schema_entry.get().strip(),
                "ProgramID": program_id,
                "LinkedGUID": linked_entry.get().strip(),
                "DocumentSchema": doc_schema_entry.get().strip(),
                "InRuleOnSave": rule_entry.get().strip(),
            }
            self.save_profiles()
            dialog.destroy()
            messagebox.showinfo(
                "Success", f"Profile '{selected_profile}' updated successfully"
            )

        Button(
            dialog, text="Save Changes", command=save_changes, bg="#2196F3", fg="white"
        ).grid(row=6, column=0, columnspan=2, pady=20)

    def delete_profile(self):
        """Delete the selected profile"""
        selected_profile = self.profile_var.get()
        if not selected_profile:
            messagebox.showwarning("Warning", "Please select a profile to delete")
            return

        if len(self.profiles) == 1:
            messagebox.showerror("Error", "Cannot delete the last profile")
            return

        confirm = messagebox.askyesno(
            "Confirm Delete",
            f"Are you sure you want to delete profile '{selected_profile}'?",
        )
        if confirm:
            del self.profiles[selected_profile]
            self.save_profiles()
            self.refresh_profile_dropdown()
            messagebox.showinfo(
                "Success", f"Profile '{selected_profile}' deleted successfully"
            )

    def log(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert(END, message + "\n")
        self.log_text.see(END)
        self.log_text.config(state="disabled")
        self.root.update()

    def add_docx_files(self):
        files = filedialog.askopenfilenames(
            title="Select DOCX Templates",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
        )
        for file in files:
            if file not in self.docx_files:
                self.docx_files.append(file)
                self.docx_listbox.insert(END, Path(file).name)

    def clear_docx_files(self):
        self.docx_files.clear()
        self.docx_listbox.delete(0, END)

    def select_rules_file(self):
        file = filedialog.askopenfilename(
            title="Select Rules Excel File",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
        )
        if file:
            self.rules_file = file
            self.rules_label.config(text=Path(file).name, fg="black")

    def select_output_dir(self):
        folder = filedialog.askdirectory(title="Select Output Directory")
        if folder:
            self.output_dir = folder
            self.output_label.config(text=folder, fg="black")

    def generate_templates(self):
        # Validation
        if not self.docx_files:
            messagebox.showerror("Error", "Please add at least one DOCX template")
            return
        if not self.rules_file:
            messagebox.showerror("Error", "Please select a rules.xlsx file")
            return
        if not self.output_dir:
            messagebox.showerror("Error", "Please select an output directory")
            return

        self.log_text.config(state="normal")
        self.log_text.delete(1.0, END)
        self.log_text.config(state="disabled")

        try:
            self.log("Starting template generation...")
            self.log(f"Output directory: {self.output_dir}")

            output_path = Path(self.output_dir)
            ensure_dir(output_path)

            # Load rules
            self.log("Loading rules from Excel...")
            rules = load_rules(Path(self.rules_file))
            self.log(f"Loaded {len(rules)} rules")

            # Process each template
            contexts: list[TemplateContext] = []

            for docx_file in self.docx_files:
                docx_path = Path(docx_file)
                self.log(f"\nProcessing: {docx_path.name}")

                try:
                    rule = find_matching_rule(docx_path.stem, rules)
                    self.log(
                        f"  Matched rule: {rule.get('FormNumber')} - {rule.get('FormTitle')}"
                    )
                except ValueError as exc:
                    self.log(f"  ERROR: {exc}")
                    continue

                template_guid = uuid.uuid4()
                document_guid = uuid.uuid4()

                # Copy DOCX
                copied = copy_template_docx(docx_path, output_path, template_guid)
                self.log(f"  Copied DOCX: {copied.name}")

                contexts.append(
                    TemplateContext(
                        template_path=docx_path,
                        template_guid=template_guid,
                        document_guid=document_guid,
                        matching_rule=rule,
                        copied_docx=copied,
                    )
                )

            # Generate JSONs and PDFs
            self.log("\nGenerating File JSONs...")
            for context in contexts:
                rule = context.matching_rule
                file_json = generate_file_json(
                    context.template_guid,
                    form_title=str(rule.get("FormTitle", "")),
                    template_name=context.template_path.stem,
                    destination_dir=output_path,
                )
                self.log(f"  Generated: {file_json.name}")

            self.log("\nGenerating Document JSONs...")

            # Load selected profile
            selected_profile_name = self.profile_var.get()
            profile_data = self.profiles.get(selected_profile_name, {})

            for context in contexts:
                rule = context.matching_rule
                doc_json = generate_document_json(
                    context.template_guid,
                    context.document_guid,
                    rule=rule,
                    destination_dir=output_path,
                    all_states=ALL_STATES,
                    profile_data=profile_data,   # <-- AHORA PASAMOS EL PERFIL
                )
                self.log(f"  Generated: {doc_json.name}")

            self.log("\nGenerating Specimen PDFs (this may take a while)...")
            for context in contexts:
                pdf_path = output_path / f"{context.template_guid}.pdf"
                generate_specimen(
                    context.copied_docx, pdf_path, watermark_text="Specimen"
                )
                self.log(f"  Generated: {pdf_path.name}")

            self.log("\nGenerating Specimen Report...")
            json_report, xlsx_report = build_specimen_report(output_path)
            self.log(f"  Generated: {json_report.name}")
            self.log(f"  Generated: {xlsx_report.name}")

            self.log("\n✓ Generation completed successfully!")
            messagebox.showinfo(
                "Success",
                f"Templates generated successfully!\n\nOutput: {self.output_dir}",
            )

        except Exception as e:
            self.log(f"\n✗ ERROR: {str(e)}")
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")


def main():
    root = Tk()
    app = TemplateGeneratorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
