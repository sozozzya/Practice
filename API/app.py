import os
import openai
import logging
import time
import pandas as pd
from flask import Flask, request, jsonify, render_template, send_from_directory
from docx import Document

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = "uploads"
app.config["REPORTS_FOLDER"] = "reports"
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
os.makedirs(app.config["REPORTS_FOLDER"], exist_ok=True)

#openai.api_key = ""

logging.basicConfig(level=logging.INFO)

REPORT_SECTIONS = {
    "1. Общие сведения об объекте оценки": [],
    "2. Анализ влияния общей политической и социально-экономической обстановки в России на рынок недвижимости": [
        "2.1. Социально-экономическая обстановка в РФ в 2024 году",
        "2.2. Основные экономические факторы",
        "2.3. Основные экономические показатели",
        "2.4. Продажа недвижимости",
        "2.5. Аренда недвижимости",
        "2.6. Ставки доходности",
        "2.7. Окупаемость инвестиций",
        "2.8. Тенденции и перспективы",
        "2.9. Прогноз на 2025 год",
        "Заключение"
    ],
    "3. Анализ влияния общей политической и социально-экономической обстановки в городе {city} на рынок недвижимости": [
        "3.1. Продажа недвижимости",
        "3.2. Динамика спроса и предложения",
        "3.3. Аренда недвижимости",
        "3.4. Тенденции и перспективы",
        "Заключение"
    ],
    "4. Информация о доме, в котором находится оцениваемая квартира": [],
    "5. Анализ местоположения объекта оценки. Ближайшее окружение": [
        "5.1. Территориально-функциональная зона",
        "5.2. Ближайшее окружение",
        "5.3. Зона местоположения",
        "Заключение"
    ],
    "6. Анализ фактических данных о ценах и арендных ставках": [
        "6.1. Фактические данные",
        "6.2. Предварительный анализ фактических данных",
        "Заключение"
    ],
    "7. Анализ основных факторов, влияющих на цены недвижимости": [
        "7.1. Перечень основных факторов",
        "7.2. Характер влияния факторов",
        "Заключение"
    ],
    "8. Анализ предложений на продажу квартир": [],
    "9. Анализ ликвидности квартир в городе {city}": [],
    "10. Анализ спроса на квартиры в городе {city}": [],
    "11. Основные выводы относительно рынка недвижимости в городе {city}": [
        "11.1. Динамика рынка",
        "11.2. Спрос и предложение",
        "11.3. Объем продаж и емкость рынка",
        "11.4. Мотивация покупателей и продавцов",
        "11.5. Ликвидность",
        "11.6. Колебания цен",
        "11.7. Другие выводы"
    ]
}

def extract_table_data(file_path):
    """Извлекает данные из загруженного файла (CSV/Excel)."""
    try:
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        elif file_path.endswith(('.xls', '.xlsx')):
            df = pd.read_excel(file_path)
        else:
            return "Формат файла не поддерживается."

        return df.to_string(index=False)
    except Exception as e:
        return f"Ошибка обработки файла: {e}"

def get_chatgpt_response(prompt):
    """Запрос к ChatGPT с обработкой ошибок."""
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}]
        )
        return response['choices'][0]['message']['content'].strip()
    except openai.error.OpenAIError as e:
        logging.error(f"Ошибка API OpenAI: {e}")
        return f"Ошибка API: {str(e)}"

@app.route('/')
def index():
    """Отображает главную страницу."""
    return render_template('index.html')

@app.route('/upload_data', methods=['POST'])
def upload_data():
    """Обрабатывает загрузку файла с данными."""
    if 'file' not in request.files:
        return jsonify({'error': 'Файл не загружен'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Файл не выбран'}), 400
    
    file_path = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
    file.save(file_path)

    extracted_data = extract_table_data(file_path)
    
    return jsonify({'message': 'Файл загружен', 'data': extracted_data})

@app.route('/get_real_estate_info', methods=['POST'])
def get_real_estate_info():
    """Генерирует отчет, используя загруженную таблицу данных."""
    data = request.json
    city = data.get('city')
    table_data = data.get('table_data', "фактические данные")

    if not city:
        return jsonify({'error': 'Город обязателен'}), 400

    report_content = []
    for section, subsections in REPORT_SECTIONS.items():
        section_title = section.format(city=city)

        full_prompt = f"{section_title}\n"
        if subsections:
            full_prompt += "\n".join([f"- {sub.format(city=city)}" for sub in subsections])

        full_prompt += f"\n\nДля анализа используйте фактические данные из таблицы \"фактические данные\"."

        response_text = get_chatgpt_response(full_prompt)
        report_content.append(f"{section_title}\n\n{response_text}\n")

        time.sleep(1)

    full_report = "\n\n".join(report_content)
    report_filename = f"report_{city}.docx"
    report_path = os.path.join(app.config["REPORTS_FOLDER"], report_filename)

    doc = Document()
    doc.add_heading(f"Анализ рынка недвижимости: {city}", level=1)
    for content in report_content:
        doc.add_paragraph(content)
    doc.save(report_path)

    return jsonify({'info': full_report, 'report_url': f"/reports/{report_filename}"})

@app.route("/reports/<filename>")
def download_report(filename):
    """Позволяет скачивать отчеты."""
    return send_from_directory(app.config["REPORTS_FOLDER"], filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)