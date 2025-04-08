from fastapi import FastAPI, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
import google.generativeai as genai
import os
import json
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from datetime import datetime
import re
import base64
import tempfile
import os
from PIL import Image
from io import BytesIO

app = FastAPI()
# как апи, на вход - инпут, на выход - документ

app.mount("/static", StaticFiles(directory="static"), name="static")

REPORTS_DIR = "reports"
os.makedirs(REPORTS_DIR, exist_ok=True)

# genai.configure(api_key="")

json_file_path = "input.json"
json_data = {}


def load_json_data():
    global json_data
    if not os.path.exists(json_file_path):
        raise HTTPException(status_code=500, detail="JSON-файл не найден.")

    with open(json_file_path, "r", encoding="utf-8") as f:
        json_data = json.load(f)


def set_document_styles(doc):
    sections = doc.sections
    for section in sections:
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(1.5)
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)

    styles = doc.styles

    normal_style = styles["Normal"]
    normal_style.font.name = "Times New Roman"
    normal_style.font.size = Pt(12)
    normal_style.font.color.rgb = RGBColor(30, 30, 147)
    normal_style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    normal_style.paragraph_format.first_line_indent = Cm(1.25)
    normal_style.paragraph_format.line_spacing = 1.5
    normal_style.paragraph_format.space_after = Pt(0)

    heading1 = styles["Heading 1"]
    heading1.font.name = "Times New Roman"
    heading1.font.size = Pt(16)
    heading1.font.bold = True
    heading1.font.color.rgb = RGBColor(30, 30, 147)
    heading1.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    heading1.paragraph_format.space_after = Pt(24)

    heading2 = styles["Heading 2"]
    heading2.font.name = "Times New Roman"
    heading2.font.size = Pt(14)
    heading2.font.bold = True
    heading2.font.color.rgb = RGBColor(30, 30, 147)
    heading2.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    heading2.paragraph_format.space_before = Pt(24)
    heading2.paragraph_format.space_after = Pt(12)


def add_page_numbers(doc):
    for section in doc.sections:
        footer = section.footer
        paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph(
        )
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        fldSimple = OxmlElement("w:fldSimple")
        fldSimple.set(qn("w:instr"), "PAGE")

        run = paragraph.add_run()
        run._r.append(fldSimple)


def format_text(doc, text):
    text = text.strip()
    if not text:
        return

    if text.startswith("# "):
        paragraph = doc.add_paragraph(text[2:], style="Heading 1")

    elif text.startswith("## "):
        paragraph = doc.add_paragraph(text[3:], style="Heading 2")

    elif text.startswith("* "):
        paragraph = doc.add_paragraph(style="ListBullet")
        remaining_text = text[2:]

        bold_parts = re.split(r"(\*\*.*?\*\*)", remaining_text)

        for part in bold_parts:
            run = paragraph.add_run()
            if part.startswith("**") and part.endswith("**"):
                run.text = part[2:-2]
                run.bold = True
            else:
                run.text = part
            run.font.color.rgb = RGBColor(30, 30, 147)

    else:
        paragraph = doc.add_paragraph(style="Normal")
        paragraph.paragraph_format.space_after = Pt(0)
        paragraph.paragraph_format.space_before = Pt(0)
        bold_parts = re.split(r"(\*\*.*?\*\*)", text)

        for part in bold_parts:
            run = paragraph.add_run()
            if part.startswith("**") and part.endswith("**"):
                run.text = part[2:-2]
                run.bold = True
            else:
                run.text = part
            run.font.color.rgb = RGBColor(30, 30, 147)


def decode_image_base64(base64_str):
    try:
        if ',' in base64_str:
            base64_data = base64_str.split(",")[1]
        else:
            base64_data = base64_str

        base64_data = base64_data.strip().replace('\n', '').replace('\r', '')

        missing_padding = len(base64_data) % 4
        if missing_padding:
            base64_data += '=' * (4 - missing_padding)

        image_data = base64.b64decode(base64_data)

        image = Image.open(BytesIO(image_data))
        image.load()

        with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as temp_file:
            image.save(temp_file, format="PNG")
            return temp_file.name

    except Exception as e:
        print(f"[decode_image_base64] Ошибка: {e}")
        raise ValueError("Ошибка при декодировании изображения из base64")


def generate_text_gemini(prompt: str) -> str:
    try:
        model = genai.GenerativeModel("gemini-2.0-flash")
        response = model.generate_content(prompt)
        return response.text.strip() if response.text else "Ошибка генерации текста"
    except Exception as e:
        return f"Ошибка при генерации текста: {str(e)}"


@app.post("/generate_report/")
async def generate_report():
    load_json_data()

    city = json_data.get("Общие характеристики", {}).get("Город")
    if not city:
        raise HTTPException(
            status_code=500, detail="Город не найден в JSON-файле.")

    doc = Document()
    set_document_styles(doc)

    current_date = datetime.today().strftime('%d.%m.%Y')

    prompts = [
        f"""Составь подробный отчет по анализу общих сведений об объекте оценки. Для этого проанализируй данные: {json_data.get('Общие характеристики', 'Нет данных')}. Опиши их связным текстом. Для составления отчета используй следующую структуру (пункты необходимо заключить в "#", а подпункты в "##"):
        # 1. Общие сведения об объекте оценки #""",

        f"""Составь подробный отчет по анализу влияния общей политической и социально-экономической обстановки в России на рынок недвижимости. Укажи достоверные цифры на {current_date} и приведи конкретные примеры. Для анализа использовать публикации только надежных источников. Укажи ссылки на используемые публикации. Для составления отчета используй следующую структуру (пункты необходимо заключить в "#", а подпункты в "##"):
        # 2. Анализ влияния общей политической и социально-экономической обстановки в России на рынок недвижимости #
            ## 2.1. Социально-экономическая обстановка в РФ в 2024 году ##
            ## 2.2. Основные экономические факторы ##
            ## 2.3. Основные экономические показатели ##
            ## 2.4. Продажа недвижимости ##
            ## 2.5. Аренда недвижимости ##
            ## 2.6. Ставки доходности ##
            ## 2.7. Окупаемость инвестиций ##
            ## 2.8. Тенденции и перспективы ##
            ## 2.9. Прогноз на 2025 год ##
            ## Заключение ##""",

        f"""Составь подробный отчет по анализу влияния общей политической и социально-экономической обстановки в городе {city} на рынок недвижимости. Укажи достоверные цифры на {current_date} и приведи конкретные примеры. Для анализа использовать публикации только надежных источников. Укажи ссылки на используемые публикации. Для составления отчета используй следующую структуру (пункты необходимо заключить в "#", а подпункты в "##"): 
        # 3. Анализ влияния общей политической и социально-экономической обстановки на рынок недвижимости #
            ## 3.1. Продажа недвижимости ##
            ## 3.2. Динамика спроса и предложения ##
            ## 3.3. Аренда недвижимости ##
            ## 3.4. Тенденции и перспективы ##
            ## Заключение ##""",

        f"""Составь подробный отчет по анализу дома, в котором находится оцениваемая квартира. Для этого проанализируй данные: {json_data.get('Характеристики здания', 'Нет данных')}. Опиши их связным текстом. Для составления отчета используй следующую структуру (пункты необходимо заключить в "#", а подпункты в "##"):
        # 4. Информация о доме, в котором находится оцениваемая квартира #""",

        f"""Составь подробный отчет по анализу местоположения объекта оценки и ближайшего окружения. Для составления отчета используй следующую структуру (пункты необходимо заключить в "#", а подпункты в "##"):
        # 5. Анализ местоположения объекта оценки. Ближайшее окружение #
            ## 5.1. Территориально-функциональная зона ## (В данном подпункте проанализируй территориально-функциональную зону: {json_data.get('Зона', 'Нет данных')}, в которой находится объект оценки),
            ## 5.2. Ближайшее окружение ## (В данном подпункте проанализируй данные: {json_data.get('Анализ местоположения', 'Нет данных')} и {json_data.get('Количестов объектов в радиусе 900м', 'Нет данных')}. Опиши их связным текстом),
            ## 5.3. Зона местоположения ## (В данном подпункте проанализируй данные: {json_data.get('Рейтинг зоны местонахождения', 'Нет данных')}. Опиши их связным текстом),
            ## Заключение ##""",

        f"""Составь подробный отчет по анализу фактических данных о ценах и арендных ставках в городе {city}. Для этого проанализируй данные : {json_data.get('Удельная арендная ставка', 'Нет данных')}, {json_data.get('Удельная цена', 'Нет данных')}. Опиши их связным текстом. Для составления отчета используй следующую структуру (пункты необходимо заключить в "#", а подпункты в "##"):
        # 6. Анализ фактических данных о ценах и арендных ставках # """,

        f"""Составь подробный отчет по анализу основных факторов, влияющих на цены недвижимости в городе {city}. Укажи достоверные цифры на {current_date} и приведи конкретные примеры. Для анализа использовать публикации только надежных источников. Укажи ссылки на используемые публикации. Для составления отчета используй следующую структуру (пункты необходимо заключить в "#", а подпункты в "##"):
        # 7. Анализ основных факторов, влияющих на цены недвижимости #
            ## 7.1. Перечень основных факторов ##
            ## 7.2. Характер влияния факторов ##
            ## Заключение ##""",

        f"""Составь подробный отчет по анализу предложений на продажу квартир в городе {city}. Для этого проанализируй данные: {json_data.get('Общая площадь', 'Нет данных')}, {json_data.get('Состояние отделки', 'Нет данных')}, {json_data.get('Количество комнат', 'Нет данных')}. Опиши их связным текстом. Для составления отчета используй следующую структуру (пункты необходимо заключить в "#", а подпункты в "##"):
        # 8. Анализ предложений на продажу квартир #""",

        f"""Составь подробный отчет по анализу ликвидности квартир в городе {city}. Для составления отчета используй следующую структуру (пункты необходимо заключить в "#", а подпункты в "##"):
        # 9. Анализ ликвидности квартир #""",

        f"""Составь подробный отчет по анализу спроса на квартиры в городе {city}. Для составления отчета используй следующую структуру (пункты необходимо заключить в "#", а подпункты в "##"):
        # 10. Анализ спроса на квартиры #""",

        f"""Составь подробный отчет по основным выводам относительно рынка недвижимости в городе {city}. Для этого проанализируй данные: {json_data.get('INF-Оценка', 'Нет данных')}. Опиши их связным текстом. Укажи достоверные цифры на {current_date} и приведи конкретные примеры. Для анализа использовать публикации только надежных источников. Укажи ссылки на используемые публикации. Для составления отчета используй следующую структуры (пункты необходимо заключить в "#", а подпункты в "##"):
        # 11. Основные выводы относительно рынка недвижимости #
            ## 11.1. Динамика рынка ##
            ## 11.2. Спрос и предложение ##
            ## 11.3. Объем продаж и емкость рынка ##
            ## 11.4. Мотивация покупателей и продавцов ##
            ## 11.5. Ликвидность ##
            ## 11.6. Колебания цен ##
            ## 11.7. Другие выводы ##"""
    ]

    for prompt in prompts:
        text = generate_text_gemini(prompt)
        lines = text.split("\n")

        image_inserted = False

        for line in lines:
            if line.startswith("#") or line.startswith("##"):
                format_text(doc, line)

                section_data = json_data.get(list(json_data.keys())[
                                             prompts.index(prompt)], {})
                image_path = section_data.get("img_src") if isinstance(
                    section_data, dict) else None

                if image_path and not image_inserted:
                    try:
                        decoded_image_path = decode_image_base64(image_path)
                        doc.add_picture(decoded_image_path, width=Inches(5.5))
                        doc.add_paragraph(
                            f"Рис: {os.path.basename(decoded_image_path)}")
                        image_inserted = True
                        os.remove(decoded_image_path)
                    except Exception as e:
                        print(f"Ошибка вставки изображения: {e}")

            else:
                format_text(doc, line)

    report_path = os.path.join(REPORTS_DIR, f"report_{city}.docx")
    doc.save(report_path)

    return {"report_path": report_path, "city": city}


@app.get("/download_report/")
async def download_report(city: str):
    report_path = os.path.join(REPORTS_DIR, f"report_{city}.docx")
    if os.path.exists(report_path):
        return FileResponse(report_path, filename=f"report_{city}.docx")
    raise HTTPException(status_code=404, detail="Отчет не найден")


@app.get("/")
async def serve_home():
    return FileResponse("static/index.html")
