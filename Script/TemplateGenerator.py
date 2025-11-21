"""
TemplateGenerator
-----------------
CLI script that replaces the previous generators by producing:
- Template docx copies named with a UUID
- File_*.json and Document_*.json populated from Input/File_.json and Input/Document_.json
- Specimen PDFs with a diagonal Arial watermark
- Specimen report (JSON + XLSX)

Subcommands:
  template  Generate UUID-named docx copies
  file      Generate file template JSON (also copies docx)
  document  Generate document template JSON (also copies docx + file JSON)
  specimen  Generate specimen PDFs with watermark (also copies docx)
  all       Run everything above
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import uuid
from dataclasses import dataclass
from datetime import date, datetime
from io import BytesIO
from pathlib import Path
from typing import Iterable, List, Sequence

import pandas as pd
from docx2pdf import convert as docx2pdf_convert
from pypdf import PdfReader, PdfWriter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas


ROOT_DIR = Path(__file__).resolve().parent.parent
INPUT_DIR = ROOT_DIR / "Input"
TEMPLATES_DIR = ROOT_DIR / "Templates"
OUTPUT_DIR = ROOT_DIR / "Output"
RULES_PATH = ROOT_DIR / "Rules" / "rules.xlsx"
ALL_STATES_PATH = ROOT_DIR / "Assets" / "AllStates.json"
FILE_TEMPLATE_PATH = INPUT_DIR / "File_.json"
DOCUMENT_TEMPLATE_PATH = INPUT_DIR / "Document_.json"


@dataclass
class TemplateContext:
    template_path: Path
    template_guid: uuid.UUID
    document_guid: uuid.UUID
    matching_rule: dict
    copied_docx: Path


def slug(text: str) -> str:
    return re.sub(r"[\s_\-]+", "", text or "").lower()


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def empty_output_folder(destination: Path) -> None:
    if not destination.exists():
        destination.mkdir(parents=True, exist_ok=True)
        return
    for child in destination.iterdir():
        if child.is_file() or child.is_symlink():
            child.unlink()
        elif child.is_dir():
            shutil.rmtree(child)


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
    ]
    rules_df = pd.read_excel(rules_path, parse_dates=["EffectiveDate", "ExpirationDate"])
    missing = [col for col in required_columns if col not in rules_df.columns]
    if missing:
        raise ValueError(f"Missing expected columns in rules.xlsx: {', '.join(missing)}")
    rules_df["EffectiveDate"] = rules_df["EffectiveDate"].apply(format_date)
    rules_df["ExpirationDate"] = rules_df["ExpirationDate"].apply(format_date)
    return json.loads(
        rules_df.to_json(orient="records", default_handler=str)
    )  # type: ignore[arg-type]


def load_all_states(all_states_path: Path) -> List[str]:
    with all_states_path.open("r", encoding="utf-8") as f:
        return json.load(f)


def format_date(value) -> str:
    if pd.isna(value):
        return ""
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return value.strftime("%Y-%m-%d")
    text = str(value)
    return text[:10]


def find_matching_rule(template_name: str, rules: Sequence[dict]) -> dict:
    sanitized = slug(template_name)
    for rule in rules:
        form_number = slug(str(rule.get("FormNumber", "")))
        if not form_number:
            continue
        if form_number in sanitized:
            return rule
    raise ValueError(f"No matching rule found for template '{template_name}'")


def copy_template_docx(source: Path, destination_dir: Path, template_guid: uuid.UUID) -> Path:
    ensure_dir(destination_dir)
    destination = destination_dir / f"{template_guid}{source.suffix}"
    shutil.copy2(source, destination)
    return destination


def generate_file_json(
    template_guid: uuid.UUID,
    form_title: str,
    template_name: str,
    template: Path,
    destination_dir: Path,
) -> Path:
    with template.open("r", encoding="utf-8") as f:
        data = json.load(f)
    guid_text = str(template_guid)
    data["id"] = guid_text
    data["SystemInfo"]["DocumentId"] = guid_text
    data["SystemInfo"]["DocumentVersionId"] = guid_text
    data["Content"]["FileName"] = f"{template_name}.docx"
    data["Content"]["AzureBlobStorageFileName"] = f"hudsoninsgroup0001/01/01/{guid_text}.docx"
    base_name = template.stem
    out_name = f"{base_name}{form_title}_{guid_text}{template.suffix}".replace(" ", "_")
    ensure_dir(destination_dir)
    out_path = destination_dir / out_name
    with out_path.open("w", encoding="utf-8") as f:
        json.dump(data, f, indent=4)
    return out_path


def generate_document_json(
    template_guid: uuid.UUID,
    document_guid: uuid.UUID,
    rule: dict,
    template: Path,
    destination_dir: Path,
    all_states: Iterable[str],
) -> Path:
    with template.open("r", encoding="utf-8") as f:
        data = json.load(f)

    template_guid_text = str(template_guid)
    document_guid_text = str(document_guid)
    form_number = str(rule.get("FormNumber", ""))
    form_title = str(rule.get("FormTitle", ""))
    edition_date = str(rule.get("EditionDate", "")).strip()
    policy_state_raw = rule.get("USStateCode", "All")
    policy_states = list(all_states) if str(policy_state_raw).lower() == "all" else [policy_state_raw]
    output_template_mandatory = "Mandatory" if str(rule.get("FormRequired", "0")) == "1" else "Optional"
    form_type = str(rule.get("FormType", "S"))
    output_template_form_type = "Dynamic" if form_type.upper() == "D" else "Static"
    print_order = str(rule.get("DisplaySequence", ""))

    data["id"] = document_guid_text
    sys_info = data.get("SystemInfo", {})
    sys_info["DocumentId"] = document_guid_text
    sys_info["DocumentVersionId"] = document_guid_text
    data["SystemInfo"] = sys_info
    content = data.get("Content", {})
    content["CommonGUID"] = document_guid_text
    content["TemplateFile"] = template_guid_text
    text_policy_form_number = f"{form_number} {edition_date}".strip()
    content["TextPolicyFormNumber"] = text_policy_form_number
    content["TextOutputTemplateTitle"] = form_title
    criteria_list = content.get("TemplateCriteria", [])
    if criteria_list:
        criteria = criteria_list[0]
        criteria["TemplateStartDate"] = format_date(rule.get("EffectiveDate", ""))
        criteria["TemplateEndDate"] = format_date(rule.get("ExpirationDate", ""))
        criteria["PolicyState"] = policy_states
        criteria["OutputTemplatePrintOrder"] = print_order
        criteria["OutputTemplateMandatory"] = output_template_mandatory
        criteria["OutputTemplateFormType"] = output_template_form_type
        criteria["TextOutputTemplateSpecimenUrl"] = f"https://hudsonfiles.hudsonportal.com/BL/{template_guid_text}.pdf"
        content["TemplateCriteria"][0] = criteria
    data["Content"] = content

    base_name = template.stem
    out_name = f"{base_name}{form_title}_{document_guid_text}{template.suffix}".replace(" ", "_")
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
    c.setFont(font_name, 64)
    c.setFillColorRGB(0.8, 0.8, 0.8)
    c.drawCentredString(0, 0, text)
    c.restoreState()
    c.save()
    buffer.seek(0)
    return buffer


def apply_watermark(base_pdf: Path, output_pdf: Path, watermark_text: str = "SPECIMEN") -> None:
    reader = PdfReader(str(base_pdf))
    writer = PdfWriter()
    for page in reader.pages:
        width = float(page.mediabox.width)
        height = float(page.mediabox.height)
        wm_pdf = PdfReader(build_watermark_page(watermark_text, width, height))
        watermark_page = wm_pdf.pages[0]
        page.merge_page(watermark_page)
        writer.add_page(page)
    with output_pdf.open("wb") as f:
        writer.write(f)


def generate_specimen(docx_path: Path, output_pdf: Path, watermark_text: str = "SPECIMEN") -> None:
    from tempfile import TemporaryDirectory

    with TemporaryDirectory() as tmp_dir:
        tmp_dir_path = Path(tmp_dir)
        docx2pdf_convert(str(docx_path), str(tmp_dir_path))
        base_pdf = tmp_dir_path / f"{docx_path.stem}.pdf"
        if not base_pdf.exists():
            raise FileNotFoundError(f"docx2pdf did not produce expected PDF at {base_pdf}")
        apply_watermark(base_pdf, output_pdf, watermark_text=watermark_text)


def build_specimen_report(destination_dir: Path) -> tuple[Path, Path]:
    ensure_dir(destination_dir)
    records = []
    for docx_file in sorted(destination_dir.glob("*.docx")):
        name = docx_file.stem
        records.append(
            {
                "FileName": name,
                "SpecimenURL": f"https://hudsonfiles.hudsonportal.com/BL/{name}.pdf",
            }
        )
    json_path = destination_dir / "Specimen Report.json"
    xlsx_path = destination_dir / "Specimen Report.xlsx"
    with json_path.open("w", encoding="utf-8") as f:
        json.dump(records, f, indent=4)
    pd.DataFrame(records).to_excel(xlsx_path, index=False)
    return json_path, xlsx_path


def process_templates(actions: set[str], clean: bool) -> None:
    ensure_dir(TEMPLATES_DIR)
    ensure_dir(OUTPUT_DIR)
    if clean:
        empty_output_folder(OUTPUT_DIR)

    rules = load_rules(RULES_PATH)
    all_states = load_all_states(ALL_STATES_PATH)
    template_files = sorted(TEMPLATES_DIR.glob("*.docx"))
    if not template_files:
        print(f"No templates found in {TEMPLATES_DIR}. Add .docx files and re-run.")
        return

    contexts: list[TemplateContext] = []
    for path in template_files:
        try:
            rule = find_matching_rule(path.stem, rules)
        except ValueError as exc:
            print(exc)
            continue
        template_guid = uuid.uuid4()
        document_guid = uuid.uuid4()
        copied = copy_template_docx(path, OUTPUT_DIR, template_guid)
        contexts.append(
            TemplateContext(
                template_path=path,
                template_guid=template_guid,
                document_guid=document_guid,
                matching_rule=rule,
                copied_docx=copied,
            )
        )

    for context in contexts:
        rule = context.matching_rule
        if "file" in actions or "document" in actions:
            generate_file_json(
                context.template_guid,
                form_title=str(rule.get("FormTitle", "")),
                template_name=context.template_path.stem,
                template=FILE_TEMPLATE_PATH,
                destination_dir=OUTPUT_DIR,
            )
        if "document" in actions:
            generate_document_json(
                context.template_guid,
                context.document_guid,
                rule=rule,
                template=DOCUMENT_TEMPLATE_PATH,
                destination_dir=OUTPUT_DIR,
                all_states=all_states,
            )
        if "specimen" in actions:
            pdf_path = OUTPUT_DIR / f"{context.template_guid}.pdf"
            generate_specimen(context.copied_docx, pdf_path, watermark_text="SPECIMEN")

    if "report" in actions:
        build_specimen_report(OUTPUT_DIR)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Template generator with specimen support.")
    parser.add_argument(
        "command",
        choices=["template", "file", "document", "specimen", "all"],
        help="Which step to run.",
    )
    parser.add_argument(
        "--no-clean",
        action="store_true",
        help="Do not empty the Output folder before running.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    command = args.command
    clean = not args.no_clean

    # Command → actions mapping
    actions_map = {
        "template": {"template"},
        "file": {"template", "file"},
        "document": {"template", "file", "document"},
        "specimen": {"template", "specimen", "report"},
        "all": {"template", "file", "document", "specimen", "report"},
    }
    actions = actions_map[command]
    process_templates(actions=actions, clean=clean)


if __name__ == "__main__":
    main()
