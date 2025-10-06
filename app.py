from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from docx import Document
from docx.shared import Inches
from docx.oxml.ns import qn
from docx.enum.text import WD_COLOR_INDEX
import os
import subprocess
from io import BytesIO
import tempfile
import shutil
import re
import time
from werkzeug.utils import secure_filename
from datetime import datetime


app = Flask(__name__)
app.secret_key = 'solar_doc_gen_secret_key_2025'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

TEMPLATES_CONFIG = {
    'net_metering': {
        'name': 'Net Metering Agreement',
        'icon': 'lightning-charge',
        'file': 'static/templates/NET.docx',
        'route': '/net-metering',
        'variables': {
            'KUSUMBE': 'location',
            '1 st day of SEP , 2025': 'agreement_date',
            '1 st day of SEP, 2025': 'agreement_date',
            'BHARTI PARAS PARDHI': 'consumer_name',
            'BNO. 05 PNO.18 GNO.383 SWAMISAMARTH JAL GAON JALGAON KUSUMBE Kh 425001': 'premises_address',
            '1.10363099084e+11': 'consumer_number',
            '110363099084': 'consumer_number',
            'MSEDCL , JALGAON DIST. JALGAON': 'distribution_licensee',
            'MSEDCL ,JALGAON DIST. JALGAON': 'distribution_licensee',
            '3 .00 KW': 'system_capacity',
            '3.00 KW': 'system_capacity',
            '3 .00 kw': 'system_capacity',
            '3.00 kw': 'system_capacity'
        },
        'form_fields': [
            {'name': 'location', 'label': 'Location', 'type': 'text', 'placeholder': 'e.g., KUSUMBE'},
            {'name': 'agreement_date', 'label': 'Agreement Date', 'type': 'text', 'placeholder': 'e.g., 1st day of OCT, 2025'},
            {'name': 'consumer_name', 'label': 'Consumer Name', 'type': 'text', 'placeholder': 'e.g., JOHN DOE'},
            {'name': 'premises_address', 'label': 'Premises Address', 'type': 'textarea', 'placeholder': 'Complete address'},
            {'name': 'consumer_number', 'label': 'Consumer Number', 'type': 'text', 'placeholder': 'e.g., 110363099084'},
            {'name': 'distribution_licensee', 'label': 'Distribution Licensee', 'type': 'text', 'placeholder': 'e.g., MSEDCL, JALGAON DIST. JALGAON'},
            {'name': 'system_capacity', 'label': 'Solar System Capacity', 'type': 'text', 'placeholder': 'e.g., 3.00'},
        ],
        'has_image': False
    },
    'wcr': {
        'name': 'Work Completion Report (WCR)',
        'icon': 'file-earmark-check',
        'file': 'static/templates/WCR.docx',
        'route': '/wcr',
        'variables': {
            'BHARTI PARAS PARDHI': 'consumer_name',
            'BHARTI PARAS PARDH0049': 'consumer_identity',
            '1.10363099084e+11': 'consumer_number',
            '110363099084': 'consumer_number',
            'Addressvariable': 'premises_address',
            '3655 9447 0923': 'aadhar_number',
            'WAAREE 560-144M (560 Wp )': 'module_model',
            'WAAREE B05F502049EA222': 'inverter_model',
            '3.3 KW': 'total_capacity',
            '3 .3 KW': 'total_capacity',
            '56 0 Wp': 'module_wattage',
            '560 Wp': 'module_wattage',
            'WAAREE POWER': 'module_make',
            '3 KW': 'installed_capacity',
            '5 KW': 'inverter_capacity',
            '5 .3': 'inverter_rating',
            '5.3': 'inverter_rating',
            '06': 'num_modules'
        },
        'form_fields': [
            {'name': 'consumer_name', 'label': 'Consumer Name', 'type': 'text', 'placeholder': 'Full Name'},
            {'name': 'consumer_identity', 'label': 'Consumer Name (Identity)', 'type': 'text', 'placeholder': 'Name as per ID'},
            {'name': 'consumer_number', 'label': 'Consumer Number', 'type': 'text', 'placeholder': '12-digit number'},
            {'name': 'premises_address', 'label': 'Installation Address', 'type': 'textarea', 'placeholder': 'Complete address'},
            {'name': 'aadhar_number', 'label': 'Aadhar Number', 'type': 'text', 'placeholder': 'XXXX XXXX XXXX'},
            {'name': 'aadhar_image', 'label': 'Aadhar Card Image', 'type': 'file', 'placeholder': 'Upload Aadhar card'},
            {'name': 'signature_image', 'label': 'Signature Image', 'type': 'file', 'placeholder': 'Upload signature'},
            {'name': 'installed_capacity', 'label': 'Installed Capacity', 'type': 'text', 'placeholder': 'e.g., 3 KW'},
            {'name': 'module_make', 'label': 'Solar Module Make', 'type': 'text', 'placeholder': 'e.g., WAAREE POWER'},
            {'name': 'module_model', 'label': 'Module Model', 'type': 'text', 'placeholder': 'Complete model number'},
            {'name': 'module_wattage', 'label': 'Module Wattage', 'type': 'text', 'placeholder': 'e.g., 560 Wp'},
            {'name': 'num_modules', 'label': 'Number of Modules', 'type': 'text', 'placeholder': 'e.g., 06'},
            {'name': 'total_capacity', 'label': 'Total Capacity (KWP)', 'type': 'text', 'placeholder': 'e.g., 3.3 KW'},
            {'name': 'inverter_model', 'label': 'Inverter Model', 'type': 'text', 'placeholder': 'Complete model number'},
            {'name': 'inverter_rating', 'label': 'Inverter Rating', 'type': 'text', 'placeholder': 'e.g., 5.3'},
            {'name': 'inverter_capacity', 'label': 'Inverter Capacity', 'type': 'text', 'placeholder': 'e.g., 5 KW'},
        ],
        'has_image': True
    },
    'model_agreement': {
        'name': 'Model Agreement',
        'icon': 'file-earmark-text',
        'file': 'static/templates/Model-Agreement-2.docx',
        'route': '/model-agreement',
        'variables': {
            'datevariable': 'agreement_date',
            'namevariable': 'consumer_name',
            'consumernumbervariable': 'consumer_number',
            'addressvariable': 'premises_address',
            'totalcapacityvariable': 'system_capacity',
            'solarmodulevariable': 'module_maker',
            'modelnamevariable': 'module_model',
            'capacityvariable': 'module_capacity',
            'companynamevariable': 'inverter_maker',
            'ratedcapacityvariable': 'inverter_capacity',
            'amountvariable': 'total_cost',
        },
        'form_fields': [
            {'name': 'agreement_date', 'label': 'Agreement Date', 'type': 'text', 'placeholder': 'e.g., 5th August 2024'},
            {'name': 'consumer_name', 'label': 'Consumer Name', 'type': 'text', 'placeholder': 'Full Name'},
            {'name': 'consumer_number', 'label': 'Consumer Number', 'type': 'text', 'placeholder': '12-digit number'},
            {'name': 'premises_address', 'label': 'Premises Address', 'type': 'textarea', 'placeholder': 'Complete address'},
            {'name': 'system_capacity', 'label': 'System Capacity', 'type': 'text', 'placeholder': 'e.g., 2.0 kWp'},
            {'name': 'module_maker', 'label': 'Module Make', 'type': 'text', 'placeholder': 'e.g., Waaree'},
            {'name': 'module_model', 'label': 'Module Model', 'type': 'text', 'placeholder': 'Complete model number'},
            {'name': 'module_capacity', 'label': 'Module Capacity (Wp)', 'type': 'text', 'placeholder': 'e.g., 900'},
            {'name': 'inverter_maker', 'label': 'Inverter Make', 'type': 'text', 'placeholder': 'e.g., TATA Power'},
            {'name': 'inverter_capacity', 'label': 'Inverter Capacity', 'type': 'text', 'placeholder': 'e.g., 4 Kw'},
            {'name': 'total_cost', 'label': 'Total Cost (₹)', 'type': 'text', 'placeholder': 'e.g., 100000'},
        ],
        'has_image': False,
        'has_special_cost': False
    },
    'proforma_a': {
        'name': 'Proforma-A (Commissioning Report)',
        'icon': 'file-earmark-ruled',
        'file': 'static/templates/2.-Annexure-I-Profarma-A-3.docx',
        'route': '/proforma-a',
        'variables': {
            'BHARTI PARAS PARDHI': 'consumer_name',
            '110363099084': 'consumer_number',
            '1.10363099084e+11': 'consumer_number',
            '9096917057': 'mobile_number',
            '9.096917057e+09': 'mobile_number',
            'addressline': 'installation_address',
            'district variable': 'district',
            '3.3 KW': 'module_total_capacity',
            '3 .3 KW': 'module_total_capacity',
            '3 KW': 'installed_capacity',
            '5 KW': 'sanctioned_capacity',
            '06': 'num_modules',
            'WAAREE SOLAR': 'inverter_make'
        },
        'form_fields': [
            {'name': 'consumer_name', 'label': 'Consumer Name', 'type': 'text', 'placeholder': 'Full Name'},
            {'name': 'consumer_number', 'label': 'Consumer Number', 'type': 'text', 'placeholder': '12-digit number'},
            {'name': 'mobile_number', 'label': 'Mobile Number', 'type': 'text', 'placeholder': '10-digit mobile'},
            {'name': 'installation_address', 'label': 'Installation Address', 'type': 'textarea', 'placeholder': 'Complete installation address'},
            {'name': 'sanctioned_capacity', 'label': 'Sanctioned Capacity', 'type': 'text', 'placeholder': 'e.g., 5 KW'},
            {'name': 'installed_capacity', 'label': 'Installed Capacity', 'type': 'text', 'placeholder': 'e.g., 3 KW'},
            {'name': 'module_total_capacity', 'label': 'Module Total Capacity', 'type': 'text', 'placeholder': 'e.g., 3.3 KW'},
            {'name': 'installation_date', 'label': 'Installation Date', 'type': 'date', 'placeholder': ''},
            {'name': 'performance_check_date', 'label': 'Performance Check Date', 'type': 'date', 'placeholder': ''},
            {'name': 'district', 'label': 'District', 'type': 'text', 'placeholder': 'e.g., Jalgaon'},
            {'name': 'num_modules', 'label': 'Number of PV Modules', 'type': 'text', 'placeholder': 'e.g., 06'},
            {'name': 'inverter_make', 'label': 'Inverter Make', 'type': 'text', 'placeholder': 'e.g., WAAREE SOLAR'},
        ],
        'has_image': False,
        'has_manual_dates': True
    }
}

def replace_text_in_paragraph(paragraph, search_text, replace_text):
    """Replace text handling split runs and whitespace variations"""
    search_normalized = ' '.join(search_text.split())
    para_normalized = ' '.join(paragraph.text.split())
    
    if search_normalized not in para_normalized:
        return False
    
    runs = paragraph.runs
    char_to_run = []
    
    for run_idx, run in enumerate(runs):
        for char in run.text:
            char_to_run.append((run_idx, char))
    
    full_text = ''.join(char for _, char in char_to_run)
    
    pattern_parts = []
    for word in search_text.split():
        escaped = re.escape(word)
        pattern_parts.append(escaped)
    
    pattern = r'[\s,]*'.join(pattern_parts)
    match = re.search(pattern, full_text, re.IGNORECASE)
    
    if not match:
        return False
    
    start_pos = match.start()
    end_pos = match.end()
    
    affected_runs = set()
    for pos in range(start_pos, end_pos):
        if pos < len(char_to_run):
            run_idx, _ = char_to_run[pos]
            affected_runs.add(run_idx)
    
    if not affected_runs:
        return False
    
    affected_runs = sorted(affected_runs)
    first_run = affected_runs[0]
    last_run = affected_runs[-1]
    
    chars_before_match = sum(len(runs[i].text) for i in range(first_run))
    start_in_first_run = start_pos - chars_before_match
    
    for run_idx in range(len(runs)):
        if run_idx < first_run:
            continue
        elif run_idx == first_run:
            prefix = runs[run_idx].text[:start_in_first_run]
            
            if run_idx == last_run:
                chars_before_this_run = sum(len(runs[i].text) for i in range(run_idx))
                suffix_start = end_pos - chars_before_this_run
                if suffix_start < len(runs[run_idx].text):
                    suffix = runs[run_idx].text[suffix_start:]
                    runs[run_idx].text = prefix + replace_text + suffix
                else:
                    runs[run_idx].text = prefix + replace_text
            else:
                runs[run_idx].text = prefix + replace_text
        elif run_idx in affected_runs and run_idx < last_run:
            runs[run_idx].text = ''
        elif run_idx == last_run and run_idx != first_run:
            chars_before_this_run = sum(len(runs[i].text) for i in range(run_idx))
            suffix_start = end_pos - chars_before_this_run
            if suffix_start < len(runs[run_idx].text):
                runs[run_idx].text = runs[run_idx].text[suffix_start:]
            else:
                runs[run_idx].text = ''
    
    return True

def replace_image_in_wcr(doc, new_image_path):
    """Replace existing Aadhar image in WCR document Para 23 - FIXED VERSION"""
    try:
        para = doc.paragraphs[23]
        
        # Remove all runs in REVERSE order to avoid index issues
        for i in range(len(para.runs) - 1, -1, -1):
            r = para.runs[i]
            r._element.getparent().remove(r._element)
        
        # Double check paragraph is empty
        if len(para.runs) > 0:
            para.clear()
        
        # Add new image
        run = para.add_run()
        run.add_picture(new_image_path, width=Inches(3.0))
        
        return True
    except Exception as e:
        print(f"Error replacing image: {str(e)}")
        
        # Fallback method
        try:
            para.clear()
            run = para.add_run()
            run.add_picture(new_image_path, width=Inches(3.0))
            return True
        except:
            return False

def replace_dates_in_proforma(doc, installation_date, performance_date):
    """Replace dates in Proforma-A - both in table AND in paragraphs"""
    try:
        if installation_date:
            date_obj = datetime.strptime(installation_date, '%Y-%m-%d')
            formatted_date = date_obj.strftime('%d/%m/%Y')
        else:
            formatted_date = ''
        
        table = doc.tables[0]
        cell = table.rows[15].cells[0]
        
        for para in cell.paragraphs:
            runs = para.runs
            date_found = False
            
            for i, run in enumerate(runs):
                if run.font.highlight_color:
                    if '1/09/' in run.text or '/' in run.text:
                        run.text = formatted_date
                        date_found = True
                    elif date_found and '2025' in run.text:
                        run.text = ''
                        break
        
        for para in doc.paragraphs:
            if '1/09/2025' in para.text or 'On 1/09/' in para.text:
                for run in para.runs:
                    if '1/09/2025' in run.text:
                        run.text = run.text.replace('1/09/2025', formatted_date)
                    elif '1/09/' in run.text:
                        run.text = run.text.replace('1/09/', formatted_date.split('/')[0] + '/' + formatted_date.split('/')[1] + '/')
        
        if performance_date:
            perf_date_obj = datetime.strptime(performance_date, '%Y-%m-%d')
            formatted_perf_date = perf_date_obj.strftime('%d/%m/%Y')
            
            for para in doc.paragraphs:
                if 'PerformanceCheckDatePlaceholder' in para.text:
                    for run in para.runs:
                        if 'PerformanceCheckDatePlaceholder' in run.text:
                            run.text = run.text.replace('PerformanceCheckDatePlaceholder', formatted_perf_date)
            
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for para in cell.paragraphs:
                            if 'PerformanceCheckDatePlaceholder' in para.text:
                                for run in para.runs:
                                    if 'PerformanceCheckDatePlaceholder' in run.text:
                                        run.text = run.text.replace('PerformanceCheckDatePlaceholder', formatted_perf_date)
        
        return True
    except Exception as e:
        print(f"Error replacing dates: {str(e)}")
        return False

def remove_all_highlighting(doc):
    """Remove yellow highlighting from entire document"""
    for para in doc.paragraphs:
        for run in para.runs:
            if run.font.highlight_color:
                run.font.highlight_color = None
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        if run.font.highlight_color:
                            run.font.highlight_color = None

def docx_replace_robust(doc, replacements):
    """Replace text throughout document with longest-first sorting"""
    all_paragraphs = list(doc.paragraphs)
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                all_paragraphs.extend(cell.paragraphs)
    
    sorted_replacements = sorted(replacements.items(), key=lambda x: len(x[0]), reverse=True)
    
    for search_text, replace_text in sorted_replacements:
        for paragraph in all_paragraphs:
            replace_text_in_paragraph(paragraph, search_text, replace_text)

def add_images_to_wcr(doc, aadhar_path, signature_path):
    """
    Add Aadhar and Signature images to WCR document
    Both images go AFTER Para 18 (Aadhar Number line)
    - Aadhar image first
    - Signature image below it
    """
    try:
        # Add Aadhar image at Para 23 (right after Aadhar Number text)
        if aadhar_path:
            para_aadhar = doc.paragraphs[23]
            run = para_aadhar.add_run()
            run.add_picture(aadhar_path, width=Inches(3.0))
            
            # Add a line break after Aadhar
            para_aadhar.add_run('\n\n')
        
        # Add Signature image right after Aadhar (same paragraph or next)
        if signature_path:
            if aadhar_path:
                # Add to same paragraph, below Aadhar
                run = para_aadhar.add_run()
                run.add_picture(signature_path, width=Inches(2.0))
            else:
                # If no Aadhar, add signature to Para 23
                para_signature = doc.paragraphs[23]
                run = para_signature.add_run()
                run.add_picture(signature_path, width=Inches(2.0))
        
        return True
    except Exception as e:
        print(f"Error adding images to WCR: {str(e)}")
        return False



def convert_docx_to_pdf(docx_path, output_dir):
    """Convert DOCX to PDF using LibreOffice"""
    try:
        libreoffice_paths = [
            '/usr/local/bin/soffice',
            '/Applications/LibreOffice.app/Contents/MacOS/soffice',
            'soffice',
            '/usr/bin/soffice'
        ]
        
        soffice_cmd = None
        for path in libreoffice_paths:
            if os.path.exists(path) or shutil.which(path):
                soffice_cmd = path
                break
        
        if not soffice_cmd:
            raise Exception("LibreOffice not found. Install: brew install --cask libreoffice")
        
        cmd = [soffice_cmd, '--headless', '--convert-to', 'pdf', '--outdir', output_dir, docx_path]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        
        if result.returncode != 0:
            raise Exception(f"LibreOffice error: {result.stderr}")
        
        pdf_filename = os.path.splitext(os.path.basename(docx_path))[0] + '.pdf'
        pdf_path = os.path.join(output_dir, pdf_filename)
        
        time.sleep(1)
        
        if not os.path.exists(pdf_path):
            raise Exception(f"PDF not created at {pdf_path}")
        
        return pdf_path
    except subprocess.TimeoutExpired:
        raise Exception("PDF conversion timeout (60s)")
    except Exception as e:
        raise Exception(f"PDF conversion failed: {str(e)}")


def process_template(template_key):
    """Generic template processor with special handling"""
    config = TEMPLATES_CONFIG[template_key]
    
    if request.method == 'POST':
        form_data = {}
        uploaded_aadhar_path = None
        uploaded_signature_path = None
        
        for field in config['form_fields']:
            if field['type'] != 'file':
                form_data[field['name']] = request.form.get(field['name'], '').strip()
        
        # Handle file uploads for WCR
        if config.get('has_image', False):
            if 'aadhar_image' in request.files:
                file = request.files['aadhar_image']
                if file and file.filename:
                    filename = secure_filename(file.filename)
                    uploaded_aadhar_path = os.path.join(app.config['UPLOAD_FOLDER'], 'aadhar_' + filename)
                    file.save(uploaded_aadhar_path)
            
            if 'signature_image' in request.files:
                file = request.files['signature_image']
                if file and file.filename:
                    filename = secure_filename(file.filename)
                    uploaded_signature_path = os.path.join(app.config['UPLOAD_FOLDER'], 'sig_' + filename)
                    file.save(uploaded_signature_path)
        
        replacements = {}
        sorted_vars = sorted(config['variables'].items(), key=lambda x: len(x[0]), reverse=True)
        
        for search_text, field_name in sorted_vars:
            if field_name in form_data and form_data[field_name]:
                replacements[search_text] = form_data[field_name]
        
        try:
            doc = Document(config['file'])
            
            docx_replace_robust(doc, replacements)
            
            if template_key == 'proforma_a' and config.get('has_manual_dates', False):
                if 'installation_date' in form_data or 'performance_check_date' in form_data:
                    dates_replaced = replace_dates_in_proforma(
                        doc, 
                        form_data.get('installation_date', ''),
                        form_data.get('performance_check_date', '')
                    )
                    if not dates_replaced:
                        flash('Warning: Could not replace dates', 'warning')
            
            # WCR: Add both images (no replacement, just add)
            if template_key == 'wcr':
                if uploaded_aadhar_path or uploaded_signature_path:
                    images_added = add_images_to_wcr(doc, uploaded_aadhar_path, uploaded_signature_path)
                    if not images_added:
                        flash('Warning: Could not add images', 'warning')
            
            remove_all_highlighting(doc)
            
            output_format = request.form.get('output_format', 'docx')
            doc_name = form_data.get('consumer_name', 'document').replace(' ', '_')
            
            if output_format == 'pdf':
                temp_dir = tempfile.mkdtemp()
                temp_docx = os.path.join(temp_dir, 'temp.docx')
                
                try:
                    doc.save(temp_docx)
                    pdf_path = convert_docx_to_pdf(temp_docx, temp_dir)
                    
                    with open(pdf_path, 'rb') as f:
                        pdf_data = BytesIO(f.read())
                    
                    pdf_data.seek(0)
                    return send_file(pdf_data, as_attachment=True,
                                   download_name=f'{template_key}_{doc_name}.pdf',
                                   mimetype='application/pdf')
                except Exception as e:
                    flash(f'PDF error: {str(e)}', 'error')
                    return redirect(url_for(template_key))
                finally:
                    try:
                        shutil.rmtree(temp_dir)
                        if uploaded_aadhar_path:
                            os.remove(uploaded_aadhar_path)
                        if uploaded_signature_path:
                            os.remove(uploaded_signature_path)
                    except:
                        pass
            else:
                stream = BytesIO()
                doc.save(stream)
                stream.seek(0)
                
                if uploaded_aadhar_path:
                    try:
                        os.remove(uploaded_aadhar_path)
                    except:
                        pass
                
                if uploaded_signature_path:
                    try:
                        os.remove(uploaded_signature_path)
                    except:
                        pass
                
                return send_file(stream, as_attachment=True,
                               download_name=f'{template_key}_{doc_name}.docx',
                               mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        
        except Exception as e:
            flash(f'Error: {str(e)}', 'error')
            if uploaded_aadhar_path:
                try:
                    os.remove(uploaded_aadhar_path)
                except:
                    pass
            if uploaded_signature_path:
                try:
                    os.remove(uploaded_signature_path)
                except:
                    pass
            return redirect(url_for(template_key))
    
    return render_template(f'{template_key}.html', config=config)


@app.route('/')
def index():
    return render_template('index.html', templates=TEMPLATES_CONFIG)

@app.route('/net-metering', methods=['GET', 'POST'])
def net_metering():
    return process_template('net_metering')

@app.route('/wcr', methods=['GET', 'POST'])
def wcr():
    return process_template('wcr')

@app.route('/model-agreement', methods=['GET', 'POST'])
def model_agreement():
    return process_template('model_agreement')

@app.route('/proforma-a', methods=['GET', 'POST'])
def proforma_a():
    return process_template('proforma_a')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)