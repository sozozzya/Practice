import os
import openai
import logging
from flask import Flask, request, jsonify, render_template, send_from_directory
from werkzeug.utils import secure_filename
from docx import Document

app = Flask(__name__)

#openai.api_key = ""

UPLOAD_FOLDER = "uploads"
REPORTS_FOLDER = "reports"
ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "gif", "webp", "csv", "xlsx"}

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["REPORTS_FOLDER"] = REPORTS_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORTS_FOLDER, exist_ok=True)

logging.basicConfig(level=logging.INFO)


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/get_real_estate_info', methods=['POST'])
def get_real_estate_info():
    data = request.json
    city = data.get('city')

    if not city:
        return jsonify({'error': 'Город обязателен для ввода'}), 400

    try:
        response = openai.ChatCompletion.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "Вы эксперт по анализу рынка недвижимости."},
                {"role": "user", "content": f"""
                    Составь отчет по анализу рынка недвижимости в городе {city}. Для отчета используй следующую структуру:
                        1. Общие сведения об объекте оценки

                        2. Анализ влияния общей политической и социально-экономической обстановки в России на рынок продажи и аренды недвижимости, в том числе тенденций, наметившихся на рынке за 2024 год
                            2.1. Социально-экономическая обстановка в РФ в 2024 году
                            2.2. Основные экономические факторы
                            2.3. Основные экономические показатели
                            2.4. Продажа недвижимости
                            2.5. Аренда недвижимости
                            2.6. Ставки доходности
                            2.7. Окупаемость инвестиций
                            2.8. Тенденции и перспективы
                            2.9. Прогноз на 2025 год
                            Заключение

                        3. Анализ влияния общей политической и социально-экономической обстановки в городе {city} на рынок продажи и аренды недвижимости, в том числе тенденций, наметившихся на рынке за 2024 год
                            3.1. Продажа недвижимости
                            3.2. Динамика спроса и предложения
                            3.3. Аренда недвижимости
                            3.4. Тенденции и перспективы
                            Заключение

                        4. Информация о доме, в котором находится оцениваемая квартира

                        5. Анализ местоположения объекта оценки. Ближайшее окружение
                            5.1. Территориально функциональная зона, в которой находится объект оценки
                            5.2. Ближайшее окружение
                            5.3. Зона местоположения
                            Заключение

                        6. Анализ фактических данных о ценах, предложений и арендных ставок с объектами из сегмента рынка, к которому может быть оцениваемый объект
                            6.1. Фактические данные
                            6.2. Предварительный анализ фактических данных
                            Заключение

                        7. Анализ основных факторов, влияющих на цены и (или) арендные ставки сопоставимых с оцениваемым объектов недвижимости
                            7.1. Перечень основных факторов, влияющих на цены квартир
                            7.2. Характер влияния на цену квартиры основных факторов ценообразования
                            Заключение

                        8. Анализ предложений на продажу квартир. Распределение выставленных на продажу квартир по основным характеристикам

                        9. Анализ ликвидности квартир. Сроки экспозиции

                        10. Анализ спроса на квартиры в сегменте

                        11. Основные выводы относительно рынка недвижимости в сегменте
                            11.1. Динамика рынка
                            11.2. Спрос и предложение
                            11.3. Объем продаж и емкость рынка
                            11.4. Мотивация покупателей и продавцов
                            11.5. Ликвидность
                            11.6. Колебания цен
                            11.7. Другие выводы

                        Использованные публикации
                """}
            ]
        )
        real_estate_info = response['choices'][0]['message']['content'].strip()

        report_filename = f"report_{city}.docx"
        report_path = os.path.join(
            app.config["REPORTS_FOLDER"], report_filename)

        doc = Document()
        doc.add_heading(f"Анализ рынка недвижимости: {city}", level=1)
        doc.add_paragraph(real_estate_info)

        doc.save(report_path)

        return jsonify({'info': real_estate_info, 'report_url': f"/reports/{report_filename}"})

    except openai.error.OpenAIError as e:
        return jsonify({'error': str(e)}), 500


@app.route("/upload_files", methods=["POST"])
def upload_files():
    files = request.files.getlist("files")
    saved_files = []

    for file in files:
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            file.save(file_path)
            file_url = f"/uploads/{filename}"
            file_type = "image" if filename.split(
                ".")[-1] in ["png", "jpg", "jpeg", "gif", "webp"] else "file"
            saved_files.append(
                {"name": filename, "url": file_url, "type": file_type})

    return jsonify({"files": saved_files})


@app.route("/uploads/<filename>")
def uploaded_file(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename)


@app.route("/reports/<filename>")
def download_report(filename):
    return send_from_directory(app.config["REPORTS_FOLDER"], filename, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)
