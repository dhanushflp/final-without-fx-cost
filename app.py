import os
from flask import Flask, render_template, request, jsonify
from werkzeug.utils import secure_filename
import subprocess
import sys
import traceback

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def run_script(script_path, inputs):
    """Run a Python script with given inputs."""
    try:
        process = subprocess.Popen(
            [sys.executable, script_path],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        
        # Simulate user inputs
        inputs_str = '\n'.join(inputs) + '\n'
        stdout, stderr = process.communicate(input=inputs_str)
        
        return {
            'success': process.returncode == 0,
            'stdout': stdout,
            'stderr': stderr
        }
    except Exception as e:
        return {
            'success': False,
            'error': str(e),
            'traceback': traceback.format_exc()
        }

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    try:
        # Collect uploaded files
        csv_file = request.files.get('csvFile')
        landing_plan_file = request.files.get('landingPlanFile')
        mg_file = request.files.get('mgFile')
        rates_file = request.files.get('ratesFile')
        
        # Save uploaded files
        files = {
            'csv': os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(csv_file.filename)),
            'landing_plan': os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(landing_plan_file.filename)),
            'mg': os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(mg_file.filename)),
            'rates': os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(rates_file.filename))
        }
        
        csv_file.save(files['csv'])
        landing_plan_file.save(files['landing_plan'])
        mg_file.save(files['mg'])
        rates_file.save(files['rates'])
        
        # Additional inputs
        vendor_name = request.form.get('vendorName', '')
        month_name = request.form.get('monthName', '')
        year_name = request.form.get('yearName', '')
        
        # Run processing scripts sequentially
        scripts_and_inputs = [
            ('scripts/first_block.py', [files['csv'], vendor_name, month_name, year_name]),
            ('scripts/second_block.py', [f"{vendor_name}_{month_name}_MIS_Summary_{year_name}.xlsx", files['landing_plan']]),
            ('scripts/third_block.py', [f"{vendor_name}_{month_name}_MIS_Summary_{year_name}_Updated.xlsx", files['mg']]),
            ('scripts/fourth_block.py', [f"{vendor_name}_{month_name}_MIS_Summary_{year_name}_Updated_updated.xlsx", files['rates']])
        ]

        results = []
        for script, inputs in scripts_and_inputs:
            result = run_script(script, inputs)
            results.append(result)
            
            if not result['success']:
                return jsonify({
                    'success': False,
                    'error': result.get('stderr', result.get('error', 'Unknown error')),
                    'traceback': result.get('traceback', '')
                })

        return jsonify({
            'success': True,
            'results': results,
            'finalFile': f"{vendor_name}_{month_name}_MIS_Summary_{year_name}_Updated_updated.xlsx"
        })
    
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e),
            'traceback': traceback.format_exc()
        })

if __name__ == '__main__':
    app.run(debug=True)
