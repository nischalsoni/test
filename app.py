from flask import Flask, request, send_file, render_template
import openpyxl
import os

app = Flask(__name__)

# Serve the HTML form
@app.route('/')
def index():
    return render_template('index.html')

# Handle form submission
@app.route('/run_script', methods=['POST'])
def run_script():
    num1 = int(request.form['num1'])
    num2 = int(request.form['num2'])
    result = num1 + num2

    # Create a new Excel file
    filename = 'result.xlsx'
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet['A1'] = 'Number 1'
    sheet['B1'] = 'Number 2'
    sheet['C1'] = 'Sum'
    sheet['A2'] = num1
    sheet['B2'] = num2
    sheet['C2'] = result
    workbook.save(filename)

    return {'message': 'Script ran successfully'}

@app.route('/download_results', methods=['GET'])
def download_results():
    filepath = 'result.xlsx'
    return send_file(filepath, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
