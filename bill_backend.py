from io import BytesIO
from pathlib import Path
import re

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
    return {
        "Consumer Number": consumer_number,
        "Name": name,
        "Units": units,
        "Amount": amount,
    }


def create_template_if_missing():
    template_path = Path("template.xlsx")
    if not template_path.exists():
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet["A1"] = "Consumer Number"
        sheet["A2"] = "Name"
        sheet["A3"] = "Units"
        sheet["A4"] = "Amount"
        workbook.save(template_path)


def fill_excel(data):
    create_template_if_missing()
    workbook = openpyxl.load_workbook("template.xlsx")
    sheet = workbook.active
    sheet["B1"] = data["Consumer Number"]
    sheet["B2"] = data["Name"]
    sheet["B3"] = data["Units"]
    sheet["B4"] = data["Amount"]
    workbook.save("output.xlsx")
    buffer = BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def process_bill(uploaded_file):
    text = image_to_text(uploaded_file)
    data = extract_data(text)
    output_bytes = fill_excel(data)
    return data, output_bytes
