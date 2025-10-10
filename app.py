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

def replace_in_runs(runs, replacements):
    """FIXED: Handle variables split across runs"""
    full_text = ''.join(run.text for run in runs)
    
    modified = False
    for var_name, var_value in replacements.items():
        if var_name in full_text:
            full_text = full_text.replace(var_name, str(var_value))
            modified = True
    
    if modified:
        for run in runs[1:]:
            run.text = ''
        if runs:
            runs[0].text = full_text
            runs[0].font.highlight_color = None

def docx_replace_robust(doc, form_data):
    """Replace variables using run-based approach"""
    replacements = {}
    for variable, field_name in VARIABLE_MAPPING.items():
        if field_name in form_data and form_data[field_name]:
            replacements[variable] = form_data[field_name]
    
    # Process paragraphs
    for para in doc.paragraphs:
        replace_in_runs(para.runs, replacements)
    
    # Process tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    replace_in_runs(para.runs, replacements)
    
    # Process headers/footers
    for section in doc.sections:
        try:
            for para in section.header.paragraphs:
                replace_in_runs(para.runs, replacements)
        except:
            pass
        try:
            for para in section.footer.paragraphs:
                replace_in_runs(para.runs, replacements)
        except:
            pass

def add_images_to_wcr(doc, aadhar_path, signature_path):
    """Find image variable placeholders and replace with actual images"""
    try:
        # Look for paragraphs with image variable names
        for para in doc.paragraphs:
            para_text = para.text
            
            # Replace signature_image_variable with actual signature
            if 'signature_image_variable' in para_text:
                # Clear the paragraph text
                for run in para.runs:
                    run.text = run.text.replace('signature_image_variable', '')
                
                # Add signature image
                if signature_path and os.path.exists(signature_path):
                    para.runs[0].add_picture(signature_path, width=Inches(1.5))
                    print(f"  ✓ Added signature image")
            
            # Replace aadhar placeholder
            if 'aadhar_image_variable' in para_text or 'aadhar_image' in para_text:
                for run in para.runs:
                    run.text = run.text.replace('aadhar_image_variable', '').replace('aadhar_image', '')
                
                if aadhar_path and os.path.exists(aadhar_path):
                    para.runs[0].add_picture(aadhar_path, width=Inches(3.0))
                    print(f"  ✓ Added Aadhar image")
        
        # Also check tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        para_text = para.text
                        
                        if 'signature_image_variable' in para_text:
                            for run in para.runs:
                                run.text = run.text.replace('signature_image_variable', '')
                            if signature_path and os.path.exists(signature_path):
                                para.runs[0].add_picture(signature_path, width=Inches(1.5))
                                print(f"  ✓ Added signature in table")
                        
                        if 'aadhar_image_variable' in para_text or 'aadhar_image' in para_text:
                            for run in para.runs:
                                run.text = run.text.replace('aadhar_image_variable', '').replace('aadhar_image', '')
                            if aadhar_path and os.path.exists(aadhar_path):
                                para.runs[0].add_picture(aadhar_path, width=Inches(3.0))
                                print(f"  ✓ Added Aadhar in table")
        
        return True
        
    except Exception as e:
        print(f"  ✗ Error adding images: {e}")
        import traceback
        traceback.print_exc()
        return True


@app.route('/')
def index():
    return render_template('unified_form.html')

@app.route('/generate_documents', methods=['POST'])
def generate_documents():
    print("\n" + "="*80)
    print("STARTING DOCUMENT GENERATION")
    print("="*80 + "\n")
    
    try:
        # Collect form data
        form_data = {}
        for key in request.form:
            value = request.form.get(key, '').strip()
            form_data[key] = value
        
        form_data['todays_date'] = datetime.now().strftime('%d/%m/%Y')
        print(f"✓ Collected {len(form_data)} form fields")
        
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
                    print(f"✓ Uploaded: {key} -> {os.path.basename(temp_path)}")
        
        # Create temp directory
        tmpdir = tempfile.mkdtemp()
        print(f"✓ Created temp dir: {tmpdir}\n")
        
        generated_files = []
        errors = []
        
        # Process each document
        for doc_name, template_path in DOCUMENT_TEMPLATES.items():
            print(f"{'='*60}")
            print(f"Processing: {doc_name}")
            print(f"{'='*60}")
            
            if not os.path.exists(template_path):
                error_msg = f"Template not found: {template_path}"
                print(f"✗ {error_msg}")
                errors.append(error_msg)
                continue
            
            try:
                # Load document
                doc = Document(template_path)
                print(f"  ✓ Loaded ({len(doc.paragraphs)} paragraphs, {len(doc.tables)} tables)")
                
                # Replace variables
                docx_replace_robust(doc, form_data)
                print(f"  ✓ Variables replaced")
                
                # Add images to WCR ONLY
                if doc_name == 'WCR':
                    aadhar = uploaded_files.get('aadhar_image')
                    sig = uploaded_files.get('signature_image')
                    if aadhar or sig:
                        print(f"  → Adding images to WCR...")
                        add_images_to_wcr(doc, aadhar, sig)
                
                # Save document
                output_file = os.path.join(tmpdir, f"{doc_name}.docx")
                doc.save(output_file)
                
                # VERIFY FILE WAS CREATED
                if os.path.exists(output_file):
                    file_size = os.path.getsize(output_file)
                    print(f"  ✓ Saved successfully ({file_size:,} bytes)")
                    generated_files.append(output_file)
                else:
                    error_msg = f"{doc_name}: File not created after save!"
                    print(f"  ✗ {error_msg}")
                    errors.append(error_msg)
                
            except Exception as e:
                error_msg = f"{doc_name}: {str(e)}"
                print(f"  ✗ ERROR: {error_msg}")
                traceback.print_exc()
                errors.append(error_msg)
            
            print()  # Blank line
        
        print(f"{'='*60}")
        print(f"SUMMARY: Generated {len(generated_files)}/{len(DOCUMENT_TEMPLATES)} documents")
        if errors:
            print(f"Errors: {len(errors)}")
            for err in errors:
                print(f"  - {err}")
        print(f"{'='*60}\n")
        
        if not generated_files:
            flash('No documents were generated! Check console for errors.', 'error')
            return redirect(url_for('index'))
        
        # Create ZIP
        consumer_name = form_data.get('consumer_name', 'client').replace(' ', '_')
        zip_filename = f"Solar_Documents_{consumer_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        zip_path = os.path.join(tmpdir, zip_filename)
        
        print(f"Creating ZIP: {zip_filename}")
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in generated_files:
                arcname = os.path.basename(file_path)
                zipf.write(file_path, arcname=arcname)
                print(f"  ✓ Added: {arcname}")
        
        # Verify ZIP
        with zipfile.ZipFile(zip_path, 'r') as zipf:
            files_in_zip = zipf.namelist()
            print(f"\n✓ ZIP contains {len(files_in_zip)} files: {files_in_zip}")
        
        # Read ZIP into memory
        with open(zip_path, 'rb') as f:
            zip_data = BytesIO(f.read())
        
        zip_data.seek(0)
        
        # Clean up uploaded files
        for file_path in uploaded_files.values():
            try:
                os.remove(file_path)
            except:
                pass
        
        print(f"\n✓ Sending ZIP file to user\n")
        print("="*80 + "\n")
        
        return send_file(
            zip_data,
            as_attachment=True,
            download_name=zip_filename,
            mimetype='application/zip'
        )
    
    except Exception as e:
        print(f"\n{'='*80}")
        print(f"FATAL ERROR:")
        print(f"{'='*80}")
        print(f"{str(e)}")
        traceback.print_exc()
        print(f"{'='*80}\n")
        flash(f'Error: {str(e)}', 'error')
        return redirect(url_for('index'))

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
