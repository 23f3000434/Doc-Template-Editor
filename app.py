from flask import Flask, render_template, request, send_file
from docx import Document
from docx2pdf import convert
import os
from datetime import datetime
from io import BytesIO
import tempfile

app = Flask(__name__)

def replace_text_in_runs(paragraph, old_text, new_text):
    """
    Paragraphs ke andar runs ke level pe text replace karta hai
    without losing any formatting like bold, italic, color etc.
    """
    if old_text in paragraph.text:
        inline = paragraph.runs
        # Pehle check karo ki text mil raha hai
        for i in range(len(inline)):
            if old_text in inline[i].text:
                inline[i].text = inline[i].text.replace(old_text, new_text)

def replace_text_in_tables(tables, old_text, new_text):
    """
    Tables ke andar text replace karta hai formatting preserve karke
    """
    for table in tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_text_in_runs(paragraph, old_text, new_text)

def replace_in_document(doc_path, replacements):
    """
    Document load karke sabhi replacements apply karta hai
    """
    doc = Document(doc_path)
    
    # Paragraphs mein replace karo
    for paragraph in doc.paragraphs:
        for old_text, new_text in replacements.items():
            replace_text_in_runs(paragraph, old_text, new_text)
    
    # Tables mein replace karo
    for old_text, new_text in replacements.items():
        replace_text_in_tables(doc.tables, old_text, new_text)
    
    return doc

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Form se data lo
        replacements = {
            'KUSUMBE': request.form['location'],
            '1 ^st^ **day of SEP, 2025**': request.form['agreement_date'],
            'BHARTI PARAS PARDHI': request.form['consumer_name'],
            'BNO. 05 PNO.18 GNO.383 SWAMISAMARTH JAL GAON JALGAON KUSUMBE Kh 425001': request.form['premises_address'],
            '110363099084': request.form['consumer_number'],
            'MSEDCL ,JALGAON DIST. JALGAON': request.form['distribution_licensee'],
            '3.00 KW': request.form['system_capacity']
        }
        
        # Original template ka path
        template_path = os.path.join('static', 'NET.docx')
        
        # Document process karo
        doc = replace_in_document(template_path, replacements)
        
        # Format ke basis pe output do
        output_format = request.form.get('output_format', 'docx')
        
        if output_format == 'pdf':
            # Temporary files banao
            with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_docx:
                doc.save(temp_docx.name)
                temp_docx_path = temp_docx.name
            
            temp_pdf_path = temp_docx_path.replace('.docx', '.pdf')
            
            try:
                # DOCX ko PDF mein convert karo
                convert(temp_docx_path, temp_pdf_path)
                
                # PDF file ko memory mein load karo
                with open(temp_pdf_path, 'rb') as pdf_file:
                    pdf_data = BytesIO(pdf_file.read())
                
                # Temporary files delete karo
                os.remove(temp_docx_path)
                os.remove(temp_pdf_path)
                
                pdf_data.seek(0)
                return send_file(
                    pdf_data,
                    as_attachment=True,
                    download_name=f'Agreement_{request.form["consumer_name"]}.pdf',
                    mimetype='application/pdf'
                )
            except Exception as e:
                # Agar PDF conversion fail ho to error handle karo
                return f"PDF conversion failed: {str(e)}. Make sure Microsoft Word is installed."
        
        else:  # docx format
            # DOCX ko memory mein save karo
            file_stream = BytesIO()
            doc.save(file_stream)
            file_stream.seek(0)
            
            return send_file(
                file_stream,
                as_attachment=True,
                download_name=f'Agreement_{request.form["consumer_name"]}.docx',
                mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
    
    return render_template('form.html')

if __name__ == '__main__':
    app.run(debug=True)
