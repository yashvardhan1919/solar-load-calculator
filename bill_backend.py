from io import BytesIO
from pathlib import Path
import re
import shutil
from datetime import datetime

import openpyxl
from PIL import Image, ImageOps
import pytesseract


tesseract_path = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")
if tesseract_path.exists():
    pytesseract.pytesseract.tesseract_cmd = str(tesseract_path)


def image_to_text(uploaded_file):
    image = Image.open(uploaded_file)
    image = ImageOps.grayscale(image)
    image = image.resize((image.width * 2, image.height * 2))
    try:
        text = pytesseract.image_to_string(image, lang="eng+mar")
    except pytesseract.TesseractError:
        text = pytesseract.image_to_string(image)
    return text


def get_lines(text):
    return [line.strip() for line in text.splitlines() if line.strip()]


def get_consumer_number(text):
    lines = get_lines(text)
    for line in lines:
        lower_line = line.lower()
        if "ग्राहक क्रमांक" in lower_line or "consumer no" in lower_line or "consumer number" in lower_line or "ग्राहक" in lower_line:
            match = re.search(r"\d{10,15}", line)
            if match:
                return match.group()
    match = re.search(r"\d{10,15}", text)
    if match:
        return match.group()
    return ""


def get_name(text, consumer_number):
    lines = get_lines(text)
    for i, line in enumerate(lines):
        if consumer_number and consumer_number in line:
            for next_line in lines[i + 1:i + 5]:
                if re.search(r"[A-Za-z]{3,}", next_line) and not re.search(r"\d{4,}", next_line):
                    return clean_name(next_line)
    return ""


def clean_name(name):
    words = name.split()
    new_words = []
    for word in words:
        only_letters = re.sub(r"[^A-Za-z]", "", word)
        if only_letters and only_letters == only_letters.upper():
            new_words.append(word)
        else:
            break
    return " ".join(new_words)


def get_units(text):
    lines = get_lines(text)
    for line in lines:
        numbers = re.findall(r"\d+(?:\.\d+)?", line)
        if len(numbers) >= 5 and "1.00" in line:
            return numbers[-1]
    for line in lines:
        if "वापर" in line.lower() or "units" in line.lower() or "unit" in line.lower():
            numbers = re.findall(r"\d+(?:\.\d+)?", line)
            if numbers:
                return numbers[-1]
    return ""


def get_amount(text):
    matches = re.findall(r"Rs\.?\s*(\d+(?:\.\d+)?)", text, re.IGNORECASE)
    if matches:
        return matches[0]
    lines = get_lines(text)
    for line in lines:
        if "देय रक्कम" in line.lower() or "amount" in line.lower() or "payable" in line.lower():
            numbers = re.findall(r"\d+(?:\.\d+)?", line)
            if numbers:
                return numbers[-1]
    return ""


def extract_data(text):
    consumer_number = get_consumer_number(text)
    name = get_name(text, consumer_number)
    units = get_units(text)
    amount = get_amount(text)
    load = get_load(text)
    connection_type = get_connection_type(text)
    fixed_charges = get_fixed_charges(text)
    bill_month = get_bill_month(text)
    return {
        "Consumer Number": consumer_number,
        "Name": name,
        "Units": units,
        "Amount": amount,
        "Load": load,
        "Connection Type": connection_type,
        "Fixed Charges": fixed_charges,
        "Bill Month": bill_month,
    }


def get_load(text):
    lines = get_lines(text)
    for line in lines:
        if "kw" in line.lower():
            match = re.search(r"\d+(?:\.\d+)?\s*kw", line, re.IGNORECASE)
            if match:
                return match.group().replace(" ", "")
    return ""


def get_connection_type(text):
    lines = get_lines(text)
    for line in lines:
        if "phase" in line.lower() or "lt" in line.lower():
            match = re.search(r"\d+\s*/?\s*LT.*?Phase", line, re.IGNORECASE)
            if match:
                return match.group().replace("  ", " ").strip()
    return ""


def get_fixed_charges(text):
    lines = get_lines(text)
    for line in lines:
        lower_line = line.lower()
        if "fixed charges" in lower_line:
            numbers = re.findall(r"\d+(?:\.\d+)?", line)
            if numbers:
                return numbers[-1]
    return "130"


def get_bill_month(text):
    matches = re.findall(r"(\d{2})-(\d{2})-(\d{4})", text)
    if matches:
        dates = []
        for day, month, year in matches:
            try:
                dates.append(datetime(int(year), int(month), int(day)))
            except ValueError:
                continue
        if dates:
            latest_date = max(dates)
            return latest_date.replace(day=1)
    return datetime.today().replace(day=1)


def get_template_path():
    download_template = Path(r"C:\Users\ADMIN\Downloads\Copy of Pranay HOME E-Bill Analysis.xlsx")
    if download_template.exists():
        return download_template
    local_template = Path("template.xlsx")
    if local_template.exists():
        return local_template
    return local_template


def get_history_rows(consumer_number):
    template_path = get_template_path()
    if not template_path.exists():
        return []
    workbook = openpyxl.load_workbook(template_path)
    sheet = workbook.active
    if str(sheet["D2"].value).split(".")[0] == consumer_number:
        month_col, unit_col, amount_col = "C", "D", "E"
    elif str(sheet["H2"].value).split(".")[0] == consumer_number:
        month_col, unit_col, amount_col = "G", "H", "I"
    else:
        return []
    rows = []
    for row in range(9, 21):
        rows.append(
            {
                "month": sheet[f"{month_col}{row}"].value,
                "units": sheet[f"{unit_col}{row}"].value,
                "amount": sheet[f"{amount_col}{row}"].value,
            }
        )
    return rows


def clear_single_bill_section(sheet):
    for cell in ["D1", "D2", "D3", "D4", "D5"]:
        sheet[cell] = None
    for cell in ["H1", "H2", "H3", "H4", "H5", "H22", "H23", "H24", "H25", "H26", "I22", "I23", "I24", "I25", "I26", "J22", "J23", "J24", "J25", "J26"]:
        sheet[cell] = None
    for row in range(9, 22):
        sheet[f"C{row}"] = None
        sheet[f"D{row}"] = None
        sheet[f"E{row}"] = None
        sheet[f"G{row}"] = None
        sheet[f"H{row}"] = None
        sheet[f"I{row}"] = None
    sheet["B9"] = None
    for row in range(10, 22):
        sheet[f"B{row}"] = row - 7


def build_history_rows(data):
    rows = get_history_rows(data["Consumer Number"])
    if rows:
        if rows[-1]["month"] is None:
            rows[-1]["month"] = data["Bill Month"]
        if data["Units"]:
            rows[-1]["units"] = float(data["Units"])
        return rows
    rows = []
    start_month = datetime(data["Bill Month"].year, data["Bill Month"].month, 1)
    for offset in range(11, -1, -1):
        year = start_month.year
        month = start_month.month - offset
        while month <= 0:
            month += 12
            year -= 1
        month_date = datetime(year, month, 1)
        units = float(data["Units"]) if offset == 0 and data["Units"] else None
        amount = float(data["Amount"]) if offset == 0 and data["Amount"] else None
        rows.append({"month": month_date, "units": units, "amount": amount})
    return rows


def fill_history_rows(sheet, rows):
    for index, row_number in enumerate(range(9, 21)):
        row_data = rows[index] if index < len(rows) else {}
        sheet[f"C{row_number}"] = row_data.get("month")
        sheet[f"D{row_number}"] = row_data.get("units")
        sheet[f"E{row_number}"] = row_data.get("amount")


def clean_layout(sheet):
    sheet.column_dimensions["A"].hidden = True
    for column in ["G", "H", "I", "J"]:
        sheet.column_dimensions[column].hidden = True
    for row in range(9, 21):
        is_empty = sheet[f"C{row}"].value is None and sheet[f"D{row}"].value is None and sheet[f"E{row}"].value is None
        sheet.row_dimensions[row].hidden = is_empty


def fill_left_section(sheet, data):
    sheet["D1"] = data["Name"]
    if data["Consumer Number"]:
        sheet["D2"] = int(data["Consumer Number"])
    if data["Fixed Charges"]:
        sheet["D3"] = float(data["Fixed Charges"])
    sheet["D4"] = data["Load"]
    sheet["D5"] = data["Connection Type"]
    history_rows = build_history_rows(data)
    fill_history_rows(sheet, history_rows)


def fill_excel(data):
    template_path = get_template_path()
    if template_path.name != "template.xlsx":
        shutil.copy(template_path, "template.xlsx")
    workbook = openpyxl.load_workbook("template.xlsx")
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    sheet = workbook.active
    clear_single_bill_section(sheet)
    fill_left_section(sheet, data)
    clean_layout(sheet)
    workbook.save("output.xlsx")
    buffer = BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def process_bill(uploaded_file):
    text = image_to_text(uploaded_file)
    data = extract_data(text)
    output_bytes = fill_excel(data)
    return data, output_bytes
