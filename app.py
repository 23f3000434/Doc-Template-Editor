from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from docx import Document
from docx.shared import Inches
import os
import zipfile
from io import BytesIO
import tempfile
import traceback
from werkzeug.utils import secure_filename
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'solar_unified_doc_generator_2025'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Complete variable mapping
VARIABLE_MAPPING = {
    'name_variable': 'consumer_name',
    'consumer_number_variable': 'consumer_number',
    'consumer_variable': 'consumer_number',
    'address_variable': 'address',
    'sanctioned_capacity_variable': 'sanctioned_capacity',
    'reinstalled_capacity_variable': 'installed_capacity',
    'module_make_variable': 'module_make',
    'inverter_capacity_variable': 'inverter_capacity',
    'module_capacity_variable': 'module_capacity',
    'number_of_pv_modules_variable': 'number_of_modules',
    'district_variable': 'district',
    'installation_date_variable': 'installation_date',
    'distribution_license_variable': 'distribution_licensee',
    'model_number_variable': 'model_number',
    'wattage_variable': 'wattage',
    'model_number_inverter_variable': 'model_number_inverter',
    'rating_variable': 'rating',
    'aadhar_number_variable': 'aadhar_number',
    'executed_date_variable': 'agreement_date',
    'module_number': 'model_number',
    'model_capacity': 'model_capacity',
    'sanctioned_caacity_variable': 'sanctioned_capacity',
    'cost_of_rts_variable': 'total_cost',
    'mobile_number_variable': 'mobile_number',
    'email_address_variable': 'email',
    'system_checkdate_variable': 'performance_check_date',
    'todays_date_variable': 'todays_date',
}

DOCUMENT_TEMPLATES = {
    'NET': 'static/templates/NET.docx',
    'WCR': 'static/templates/WCR.docx',
    'Model-Agreement': 'static/templates/Model-Agreement.docx',
    'Proforma-A': 'static/templates/2.-Annexure-I-Profarma-A.docx'
}

def replace_text_in_paragraph(paragraph, search_text, replace_text):
    """Replace text preserving formatting"""
    if search_text in paragraph.text:
        inline = paragraph.runs
        for run in inline:
            if search_text in run.text:
                run.text = run.text.replace(search_text, replace_text)
                run.font.highlight_color = None
        return True
    return False

def remove_all_highlighting(doc):
    """Remove highlighting"""
    for para in doc.paragraphs:
        for run in para.runs:
            run.font.highlight_color = None
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.highlight_color = None

def docx_replace_robust(doc, form_data):
    """Replace variables"""
    for para in doc.paragraphs:
        for variable, field_name in VARIABLE_MAPPING.items():
            if field_name in form_data and form_data[field_name]:
                replace_text_in_paragraph(para, variable, form_data[field_name])
    
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for variable, field_name in VARIABLE_MAPPING.items():
                        if field_name in form_data and form_data[field_name]:
                            replace_text_in_paragraph(para, variable, form_data[field_name])
    
    remove_all_highlighting(doc)

def add_images_to_wcr(doc, aadhar_path, signature_path):
    """FIXED: Safely add images to WCR"""
    try:
        # Add images at the END of document (safest approach)
        if aadhar_path and os.path.exists(aadhar_path):
            try:
                para = doc.add_paragraph()
                para.add_run().add_picture(aadhar_path, width=Inches(3.0))
                print("✓ Aadhar image added to WCR")
            except Exception as e:
                print(f"✗ Aadhar image failed: {e}")
        
        if signature_path and os.path.exists(signature_path):
            try:
                para = doc.add_paragraph()
                para.add_run().add_picture(signature_path, width=Inches(1.5))
                print("✓ Signature image added to WCR")
            except Exception as e:
                print(f"✗ Signature image failed: {e}")
        
        return True
    except Exception as e:
        print(f"✗ Image addition error: {e}")
        traceback.print_exc()
        # Return True anyway - don't fail WCR just because images failed
        return True

@app.route('/')
def index():
    return render_template('unified_form.html')

@app.route('/generate_documents', methods=['POST'])
def generate_documents():
    print("=== GENERATE DOCUMENTS CALLED ===")
    
    try:
        # Collect form data
        form_data = {}
        for key in request.form:
            value = request.form.get(key, '').strip()
            form_data[key] = value
            print(f"Form field: {key} = {value[:50] if value else 'empty'}")
        
        # Add today's date
        form_data['todays_date'] = datetime.now().strftime('%d/%m/%Y')
        
        # Handle file uploads
        uploaded_files = {}
        for key in ['aadhar_image', 'signature_image']:
            if key in request.files:
                file = request.files[key]
                if file and file.filename:
                    filename = secure_filename(file.filename)
                    temp_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    file.save(temp_path)
                    uploaded_files[key] = temp_path
                    print(f"Uploaded: {key} -> {temp_path}")
        
        # Create temp directory
        tmpdir = tempfile.mkdtemp()
        print(f"Temp dir: {tmpdir}")
        
        generated_files = []
        
        # Process each document
        for doc_name, template_path in DOCUMENT_TEMPLATES.items():
            print(f"\n--- Processing {doc_name} ---")
            print(f"Template path: {template_path}")
            
            if not os.path.exists(template_path):
                print(f"ERROR: Template not found: {template_path}")
                flash(f'Template not found: {template_path}', 'error')
                continue
            
            try:
                doc = Document(template_path)
                print(f"Document loaded: {len(doc.paragraphs)} paragraphs")
                
                # Replace variables
                docx_replace_robust(doc, form_data)
                print("Variables replaced")
                
                # Add images to WCR
                if doc_name == 'WCR':
                    aadhar = uploaded_files.get('aadhar_image')
                    sig = uploaded_files.get('signature_image')
                    if aadhar or sig:
                        add_images_to_wcr(doc, aadhar, sig)
                        print("Images added to WCR")
                
                # Save
                output_file = os.path.join(tmpdir, f"{doc_name}.docx")
                doc.save(output_file)
                generated_files.append(output_file)
                print(f"Saved: {output_file}")
                
            except Exception as e:
                print(f"ERROR processing {doc_name}: {e}")
                traceback.print_exc()
        
        if not generated_files:
            flash('No documents were generated', 'error')
            return redirect(url_for('index'))
        
        # Create ZIP
        consumer_name = form_data.get('consumer_name', 'client').replace(' ', '_')
        zip_filename = f"Solar_Documents_{consumer_name}.zip"
        zip_path = os.path.join(tmpdir, zip_filename)
        
        print(f"\n--- Creating ZIP: {zip_path} ---")
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in generated_files:
                zipf.write(file_path, arcname=os.path.basename(file_path))
                print(f"Added to ZIP: {os.path.basename(file_path)}")
        
        # Read ZIP into memory
        with open(zip_path, 'rb') as f:
            zip_data = BytesIO(f.read())
        
        zip_data.seek(0)
        
        # Clean up
        for file_path in uploaded_files.values():
            try:
                os.remove(file_path)
            except:
                pass
        
        print(f"Sending ZIP: {zip_filename}")
        
        # Send file
        return send_file(
            zip_data,
            as_attachment=True,
            download_name=zip_filename,
            mimetype='application/zip'
        )
    
    except Exception as e:
        print(f"\n=== FATAL ERROR ===")
        print(f"Error: {e}")
        traceback.print_exc()
        flash(f'Error generating documents: {str(e)}', 'error')
        return redirect(url_for('index'))

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)  # Debug=True for testing
