
import os
import re
import tempfile
import unicodedata
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd
from flask import Flask, render_template, request, send_file, flash, redirect, url_for, jsonify
from werkzeug.utils import secure_filename
from openpyxl import Workbook
from openpyxl.styles import Font
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

APP_TITLE = "E-SOCIAL EVELLYN"
ALLOWED_EXTENSIONS = {".xls", ".xlsx", ".html", ".htm"}

app = Flask(__name__)
app.secret_key = "troque-esta-chave-em-producao"


def normalize_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip().upper()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"\s+", " ", text)
    return text


def sanitize_filename(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|]+', "_", str(name))
    name = re.sub(r"\s+", " ", name).strip()
    return name or "arquivo"


def unique_path(path: str) -> str:
    base, ext = os.path.splitext(path)
    if not os.path.exists(path):
        return path
    counter = 2
    while True:
        new_path = f"{base} ({counter}){ext}"
        if not os.path.exists(new_path):
            return new_path
        counter += 1


def is_allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


def extract_cnpj(text: str) -> str:
    text = "" if text is None else str(text)
    digits = re.sub(r"\D", "", text)
    if len(digits) >= 14:
        return digits[:14]
    return ""


def score_dataframe(df: pd.DataFrame) -> int:
    score = 0
    cols = [normalize_text(c) for c in df.columns]
    for wanted in ["FUNCIONARIO", "TIPO DE EXAME", "DEPOSITANTE", "SETOR", "NOME", "TIPO", "EMPRESA"]:
        if wanted in cols:
            score += 10
    score += min(len(df), 50)
    return score


def list_sheets(path: str):
    suffix = Path(path).suffix.lower()
    if suffix in {".html", ".htm"}:
        return ["Planilha principal"]
    try:
        xl = pd.ExcelFile(path)
        if xl.sheet_names:
            return xl.sheet_names
    except Exception:
        pass
    try:
        tables = pd.read_html(path)
        if tables:
            return ["Planilha principal"]
    except Exception:
        pass
    return ["Planilha principal"]


def read_spreadsheet(path: str, selected_sheet: str | None = None) -> pd.DataFrame:
    suffix = Path(path).suffix.lower()

    if selected_sheet and selected_sheet != "Planilha principal":
        try:
            return pd.read_excel(path, sheet_name=selected_sheet)
        except Exception as exc:
            raise RuntimeError(
                f"Não foi possível ler a aba '{selected_sheet}' do arquivo {os.path.basename(path)}. Erro: {exc}"
            ) from exc

    if suffix == ".xls":
        try:
            tables = pd.read_html(path)
            if tables:
                return tables[0]
        except Exception:
            pass

    try:
        xl = pd.ExcelFile(path)
        best_df = None
        best_score = -1
        for sheet in xl.sheet_names:
            try:
                df = pd.read_excel(path, sheet_name=sheet)
            except Exception:
                continue
            score = score_dataframe(df)
            if score > best_score:
                best_df = df
                best_score = score
        if best_df is not None:
            return best_df
    except Exception:
        pass

    try:
        return pd.read_excel(path)
    except Exception:
        pass

    try:
        tables = pd.read_html(path)
        if tables:
            return tables[0]
    except Exception as exc:
        raise RuntimeError(f"Não foi possível ler o arquivo {os.path.basename(path)}. Erro: {exc}") from exc

    raise RuntimeError(f"Não foi possível ler o arquivo {os.path.basename(path)}.")


def find_column(df: pd.DataFrame, expected_names: list[str]) -> str:
    normalized = {normalize_text(col): col for col in df.columns}
    for name in expected_names:
        norm = normalize_text(name)
        if norm in normalized:
            return normalized[norm]
    raise KeyError(
        f"Coluna não encontrada. Esperado um destes nomes: {expected_names}. "
        f"Colunas encontradas: {list(df.columns)}"
    )


def prepare_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df.dropna(axis=1, how="all")
    df = df.dropna(axis=0, how="all")
    df.columns = [str(c).strip() for c in df.columns]

    for col in df.columns:
        if normalize_text(col) == "DATA":
            try:
                original = df[col]
                dt = pd.to_datetime(df[col], errors="coerce", dayfirst=True)
                formatted = dt.dt.strftime("%d/%m/%Y")
                df[col] = formatted.where(~formatted.isna(), original.astype(str))
            except Exception:
                pass
    return df


def build_key_series(name_series: pd.Series, type_series: pd.Series) -> pd.Series:
    return name_series.map(normalize_text) + "||" + type_series.map(normalize_text)


def make_paragraph(text: str, style: ParagraphStyle) -> Paragraph:
    text = "" if pd.isna(text) else str(text)
    text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    text = text.replace("\n", "<br/>")
    return Paragraph(text, style)


def build_pdf(df: pd.DataFrame, pdf_path: str, title: str):
    if df.empty:
        raise ValueError("A tabela filtrada ficou vazia. Não há dados para gerar o PDF.")

    page_width, _ = landscape(A4)
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=landscape(A4),
        leftMargin=10 * mm,
        rightMargin=10 * mm,
        topMargin=10 * mm,
        bottomMargin=10 * mm,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "TitleCustom",
        parent=styles["Title"],
        alignment=TA_CENTER,
        fontName="Helvetica-Bold",
        fontSize=12,
        spaceAfter=6,
        textColor=colors.black,
    )
    cell_style = ParagraphStyle(
        "Cell",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=8.0,
        leading=9.5,
        alignment=TA_CENTER,
        spaceAfter=0,
        spaceBefore=0,
    )
    header_style = ParagraphStyle(
        "Header",
        parent=cell_style,
        fontName="Helvetica-Bold",
        textColor=colors.white,
    )

    headers = [str(col) for col in df.columns]
    data = [[make_paragraph(h, header_style) for h in headers]]
    for _, row in df.iterrows():
        data.append([make_paragraph(row[col], cell_style) for col in df.columns])

    total_width = page_width - doc.leftMargin - doc.rightMargin
    preferred = {
        "EVENTO": 14, "EMPRESA": 34, "UNIDADE": 30, "NOME": 28,
        "CPF": 16, "TIPO": 18, "STATUS": 18, "DATA": 15,
        "RECIBO ESOCIAL": 28, "RECIBO E-SOCIAL": 28, "RECIBO SEFAZ": 30,
    }
    weights = [preferred.get(normalize_text(col), 18) for col in headers]
    weight_sum = sum(weights)
    col_widths = [total_width * w / weight_sum for w in weights]

    table = Table(data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#3A3A3A")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.75, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))

    doc.build([Paragraph(title, title_style), Spacer(1, 2 * mm), table])


def create_output_folder(base_output_dir: str) -> str:
    timestamp = datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    folder_path = os.path.join(base_output_dir, f"RESULTADO FINAL - {timestamp}")
    os.makedirs(folder_path, exist_ok=True)
    return folder_path


def create_structure(base_folder: str):
    pdf_folder = os.path.join(base_folder, "PDFs")
    log_folder = os.path.join(base_folder, "Logs")
    os.makedirs(pdf_folder, exist_ok=True)
    os.makedirs(log_folder, exist_ok=True)
    return pdf_folder, log_folder


def create_zip_from_folder(folder: str) -> str:
    zip_path = unique_path(folder.rstrip("/\\") + ".zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(folder):
            for filename in files:
                if filename.lower().endswith(".zip"):
                    continue
                full_path = os.path.join(root, filename)
                arcname = os.path.relpath(full_path, folder)
                zf.write(full_path, arcname)
    return zip_path


def export_summary_excel(rows: list[dict], excel_path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumo"

    headers = [
        "EMPRESA",
        "CNPJ",
        "STATUS",
        "TOTAL BASE EMPRESA",
        "TOTAL ENCONTRADO NO SISTEMA",
        "MOTIVO",
        "PDF GERADO",
    ]
    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True)

    for row in rows:
        ws.append([
            row.get("empresa", ""),
            row.get("cnpj", ""),
            row.get("status", ""),
            row.get("total_base", 0),
            row.get("total_encontrado", 0),
            row.get("motivo", ""),
            row.get("pdf", ""),
        ])

    widths = {
        "A": 45, "B": 18, "C": 18, "D": 20, "E": 24, "F": 100, "G": 55
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    wb.save(excel_path)


def get_company_fields_system(system_df: pd.DataFrame):
    normalized_map = {normalize_text(col): col for col in system_df.columns}
    company_col = None
    for wanted in ["EMPRESA", "SETOR", "UNIDADE"]:
        if wanted in normalized_map:
            company_col = normalized_map[wanted]
            break
    if not company_col:
        raise KeyError("Não foi possível identificar a coluna da empresa na planilha do sistema.")
    values = system_df[company_col].dropna().astype(str).str.strip()
    company_text = values.iloc[0] if not values.empty else "EMPRESA"
    company_cnpj = extract_cnpj(company_text)
    return company_col, company_text, company_cnpj


def get_company_fields_base(base_df: pd.DataFrame):
    normalized_map = {normalize_text(col): col for col in base_df.columns}
    company_col = None
    for wanted in ["SETOR", "EMPRESA", "UNIDADE"]:
        if wanted in normalized_map:
            company_col = normalized_map[wanted]
            break
    if not company_col:
        raise KeyError("Não foi possível identificar a coluna da empresa na planilha base.")
    return company_col


def filter_base_company(base_df: pd.DataFrame, company_text: str, company_cnpj: str, company_col: str):
    work = base_df.copy()
    work["__EMPRESA_TXT__"] = work[company_col].astype(str).str.strip()
    work["__EMPRESA_NORM__"] = work["__EMPRESA_TXT__"].map(normalize_text)
    work["__EMPRESA_CNPJ__"] = work["__EMPRESA_TXT__"].map(extract_cnpj)

    if company_cnpj:
        filtered = work[work["__EMPRESA_CNPJ__"] == company_cnpj].copy()
        if not filtered.empty:
            return filtered

    company_norm = normalize_text(company_text)
    return work[work["__EMPRESA_NORM__"] == company_norm].copy()


def run_company_process(system_file: str, base_file: str, pdf_folder: str, log_folder: str, base_sheet: str | None = None):
    system_df = prepare_dataframe(read_spreadsheet(system_file))
    base_df = prepare_dataframe(read_spreadsheet(base_file, selected_sheet=base_sheet))

    system_nome = find_column(system_df, ["NOME"])
    system_tipo = find_column(system_df, ["TIPO"])
    base_nome = find_column(base_df, ["FUNCIONARIO", "FUNCIONÁRIO"])
    base_tipo = find_column(base_df, ["TIPO DE EXAME"])
    base_status = find_column(base_df, ["DEPOSITANTE"])

    _, company_text, company_cnpj = get_company_fields_system(system_df)
    base_company_col = get_company_fields_base(base_df)
    base_company_df = filter_base_company(base_df, company_text, company_cnpj, base_company_col)

    if base_company_df.empty:
        return {
            "empresa": company_text,
            "cnpj": company_cnpj,
            "status": "NÃO GERADO",
            "total_base": 0,
            "total_encontrado": 0,
            "motivo": "Empresa não encontrada na planilha base.",
            "pdf": "",
        }

    base_company_df["__STATUS_OK__"] = base_company_df[base_status].map(normalize_text)
    invalid_status = base_company_df[base_company_df["__STATUS_OK__"] != "OK E-SOCIAL"].copy()

    system_df["__CHAVE__"] = build_key_series(system_df[system_nome], system_df[system_tipo])
    base_company_df["__CHAVE__"] = build_key_series(base_company_df[base_nome], base_company_df[base_tipo])

    expected_keys = set(base_company_df["__CHAVE__"].dropna().tolist())
    filtered_system = system_df[system_df["__CHAVE__"].isin(expected_keys)].copy()
    filtered_system = filtered_system[[c for c in system_df.columns if c != "__CHAVE__"]].copy()

    missing_keys = sorted(expected_keys - set(system_df["__CHAVE__"].dropna().tolist()))

    reasons = []
    if not invalid_status.empty:
        for _, row in invalid_status.iterrows():
            reasons.append(
                f"{row.get(base_nome, '')} | {row.get(base_tipo, '')} | {row.get(base_status, '')}"
            )

    if missing_keys:
        for key in missing_keys:
            name, exam = key.split("||", 1)
            reasons.append(f"NÃO ENCONTRADO NO SISTEMA | {name} | {exam}")

    log_path = unique_path(os.path.join(log_folder, sanitize_filename(company_text) + " - LOG.txt"))
    with open(log_path, "w", encoding="utf-8") as f:
        f.write("RESUMO DO PROCESSAMENTO\n")
        f.write("=" * 80 + "\n")
        f.write(f"Planilha do sistema: {system_file}\n")
        f.write(f"Planilha base: {base_file}\n")
        f.write(f"Aba base usada: {base_sheet or 'Detecção automática'}\n")
        f.write(f"Empresa: {company_text}\n")
        f.write(f"CNPJ: {company_cnpj}\n")
        f.write(f"Total base empresa: {len(base_company_df)}\n")
        f.write(f"Total encontrado no sistema: {len(filtered_system)}\n\n")

        if reasons:
            f.write("MOTIVOS PARA NÃO GERAR PDF\n")
            f.write("-" * 80 + "\n")
            for item in reasons:
                f.write(f"- {item}\n")
        else:
            f.write("Todos os funcionários da empresa estão com OK E-SOCIAL e foram encontrados no sistema.\n")

    if reasons:
        return {
            "empresa": company_text,
            "cnpj": company_cnpj,
            "status": "NÃO GERADO",
            "total_base": len(base_company_df),
            "total_encontrado": len(filtered_system),
            "motivo": " | ".join(reasons),
            "pdf": "",
        }

    pdf_path = unique_path(os.path.join(pdf_folder, sanitize_filename(company_text) + ".pdf"))
    build_pdf(filtered_system, pdf_path, title=company_text)

    return {
        "empresa": company_text,
        "cnpj": company_cnpj,
        "status": "GERADO",
        "total_base": len(base_company_df),
        "total_encontrado": len(filtered_system),
        "motivo": "OK",
        "pdf": pdf_path,
    }


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html", title=APP_TITLE)


@app.route("/abas-base", methods=["POST"])
def abas_base():
    base_file = request.files.get("base_file")
    if not base_file or not base_file.filename:
        return jsonify({"ok": False, "error": "Nenhuma planilha base enviada."}), 400

    temp_root = Path(tempfile.mkdtemp(prefix="esocial_abas_"))
    try:
        base_path = temp_root / secure_filename(base_file.filename)
        base_file.save(base_path)
        sheets = list_sheets(str(base_path))
        return jsonify({"ok": True, "sheets": sheets})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 500


@app.route("/processar", methods=["POST"])
def processar():
    base_file = request.files.get("base_file")
    rel_files = request.files.getlist("rel_files")
    base_sheet = request.form.get("base_sheet", "").strip() or None

    if not base_file or not base_file.filename:
        flash("Selecione a planilha base.")
        return redirect(url_for("index"))

    valid_rel_files = [f for f in rel_files if f and f.filename and is_allowed_file(f.filename)]
    if not valid_rel_files:
        flash("Selecione um ou mais arquivos RELFUNCGERAL válidos.")
        return redirect(url_for("index"))

    temp_root = Path(tempfile.mkdtemp(prefix="esocial_web_"))
    upload_dir = temp_root / "uploads"
    output_root = temp_root / "saida"
    upload_dir.mkdir(parents=True, exist_ok=True)
    output_root.mkdir(parents=True, exist_ok=True)

    try:
        base_path = upload_dir / secure_filename(base_file.filename)
        base_file.save(base_path)

        rel_paths = []
        for index, rel in enumerate(valid_rel_files, start=1):
            filename = secure_filename(Path(rel.filename).name)
            path = upload_dir / f"{index:03d}_{filename}"
            rel.save(path)
            rel_paths.append(path)

        general_output_folder = Path(create_output_folder(str(output_root)))
        pdf_folder, log_folder = create_structure(str(general_output_folder))

        summary_rows = []
        for rel_path in rel_paths:
            try:
                summary_rows.append(
                    run_company_process(
                        str(rel_path),
                        str(base_path),
                        pdf_folder,
                        log_folder,
                        base_sheet=base_sheet,
                    )
                )
            except Exception as exc:
                summary_rows.append({
                    "empresa": rel_path.name,
                    "cnpj": "",
                    "status": "NÃO GERADO",
                    "total_base": 0,
                    "total_encontrado": 0,
                    "motivo": str(exc),
                    "pdf": "",
                })

        resumo_excel = str(Path(general_output_folder) / "RESUMO_FINAL.xlsx")
        export_summary_excel(summary_rows, resumo_excel)

        resumo_txt = Path(log_folder) / "RESUMO_GERAL.txt"
        with open(resumo_txt, "w", encoding="utf-8") as f:
            f.write("RESUMO GERAL DO PROCESSAMENTO\n")
            f.write("=" * 80 + "\n\n")
            f.write(f"Planilha base: {base_path.name}\n")
            f.write(f"Aba base usada: {base_sheet or 'Detecção automática'}\n")
            f.write(f"Total de empresas processadas: {len(summary_rows)}\n")
            f.write(f"PDFs gerados: {sum(1 for r in summary_rows if r['status'] == 'GERADO')}\n")
            f.write(f"PDFs não gerados: {sum(1 for r in summary_rows if r['status'] != 'GERADO')}\n")
            f.write(f"Resumo Excel: {resumo_excel}\n")

        zip_path = Path(create_zip_from_folder(str(general_output_folder)))
        return send_file(zip_path, as_attachment=True, download_name=zip_path.name, mimetype="application/zip")

    except Exception as exc:
        flash(f"Erro ao processar: {exc}")
        return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
