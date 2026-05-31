"""
Веб-приложение для анализа строительных смет.
Запуск: python app.py
"""
import io
import os
import traceback

from flask import Flask, render_template, request, send_file, jsonify

from processor.reader import read_smeta_file
from processor.claude_api import (
    extract_structure,
    get_material_market_prices,
    get_work_market_rates,
    get_market_analysis,
)
from processor.excel_gen import generate_excel_book

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024
os.makedirs('uploads', exist_ok=True)


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/process', methods=['POST'])
def process():
    if 'file' not in request.files:
        return jsonify({'error': 'Файл не выбран'}), 400
    file = request.files['file']
    if not file.filename:
        return jsonify({'error': 'Файл не выбран'}), 400
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in ('.docx', '.xlsx', '.xls'):
        return jsonify({'error': 'Поддерживаются .docx и .xlsx'}), 400

    try:
        file_bytes = file.read()

        # Шаг 1: Читаем и сжимаем файл
        text = read_smeta_file(file_bytes, ext)

        # Шаг 2: Claude извлекает полную структуру (все позиции + материалы)
        smeta_data = extract_structure(text)

        # Собираем плоский список всех материалов и позиций
        all_materials = [
            m for sec in smeta_data['sections']
            for m in sec.get('materials', [])
        ]
        all_positions = [
            p for sec in smeta_data['sections']
            for p in sec.get('positions', [])
        ]

        # Шаг 3: Рыночные цены на материалы (построчно)
        mat_prices = get_material_market_prices(all_materials)

        # Шаг 4: Рыночные расценки на работы (построчно)
        work_rates = get_work_market_rates(all_positions)

        # Шаг 5: Собираем рыночный анализ (математика, без API)
        full_data = get_market_analysis(smeta_data, mat_prices, work_rates)

        # Шаг 6: Генерируем Excel
        excel_bytes = generate_excel_book(full_data)

        project_name = full_data.get('project', {}).get('name', 'Проект')
        safe_name = ''.join(c for c in project_name if c.isalnum() or c in ' _-').strip()[:40]
        filename = f"{safe_name or 'Смета'}_Книга_проекта.xlsx"

        return send_file(
            io.BytesIO(excel_bytes),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route('/health')
def health():
    return jsonify({'status': 'ok'})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
