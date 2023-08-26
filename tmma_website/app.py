import sys,os

from flask import Flask, render_template, request, send_from_directory

sys.path.append('/home/didi/myproject/tmma/')
from main import ExcelAgendaEngine

app = Flask(__name__)

@app.route('/')
def index():
    with open('user_input_template.txt', 'r') as f:
        default_text = f.read()
    return render_template('index.html', default_text=default_text, message=None)

@app.route('/save', methods=['POST'])
def save_text():
    text_content = request.form.get('text_content')
    if text_content:
        filename = 'real_user_input.txt'
        with open(filename, 'w') as file:
            file.write(text_content)
        user_input_file_path = os.path.join(app.root_path, filename)
        output_excel_file_path = os.path.join(app.root_path, 'generated_agenda.xlsx')
        engine = ExcelAgendaEngine(user_input_file_path, output_excel_file_path)
        engine.write()

        excel_generated_link = output_excel_file_path
        return render_template('index.html', default_text=text_content, message='Text saved successfully.', excel_generated=excel_generated_link)
    else:
        return render_template('index.html', default_text='', message='No text to save.')

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.root_path, filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)