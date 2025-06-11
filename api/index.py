import os
import shutil
import zipfile
from lxml import etree as ET
from datetime import datetime
from flask import Flask, render_template, request, send_file, url_for, redirect, session, abort, jsonify
from werkzeug.utils import secure_filename
import tempfile
import secrets

app = Flask(__name__)
app.secret_key = secrets.token_hex(16)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xlsm'}

EXCEL_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NSMAP = {'main': EXCEL_NS}

REQUIRED_XLSX_FILES = {'[Content_Types].xml', 'xl/workbook.xml'}

# Vercel: templates directory is at ../../templates relative to this file
TEMPLATE_DIR = os.path.join(os.path.dirname(__file__), '../../templates')

app.template_folder = TEMPLATE_DIR

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def is_valid_excel(file_path):
    # Check ZIP header
    with open(file_path, 'rb') as f:
        header = f.read(4)
        if header != b'PK\x03\x04':
            return False
    # Check required files in ZIP
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            names = set(zip_ref.namelist())
            for req in REQUIRED_XLSX_FILES:
                if req not in names:
                    return False
    except Exception:
        return False
    return True

def remove_nodes(tree, xpath):
    for node in tree.xpath(xpath, namespaces=NSMAP):
        parent = node.getparent()
        if parent is not None:
            parent.remove(node)

def remove_protection(file_path):
    with tempfile.TemporaryDirectory() as temp_dir:
        # Extract all files
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
            all_files = zip_ref.namelist()
            compression = zip_ref.compression

        # Edit worksheet XML files
        worksheets_dir = os.path.join(temp_dir, 'xl', 'worksheets')
        if os.path.exists(worksheets_dir):
            for xml_file in os.listdir(worksheets_dir):
                if xml_file.endswith('.xml'):
                    xml_path = os.path.join(worksheets_dir, xml_file)
                    parser = ET.XMLParser(remove_blank_text=False)
                    tree = ET.parse(xml_path, parser)
                    remove_nodes(tree, '//main:sheetProtection')
                    tree.write(xml_path, encoding='UTF-8', xml_declaration=True, standalone='yes')

        # Edit workbook.xml
        workbook_path = os.path.join(temp_dir, 'xl', 'workbook.xml')
        if os.path.exists(workbook_path):
            parser = ET.XMLParser(remove_blank_text=False)
            tree = ET.parse(workbook_path, parser)
            remove_nodes(tree, '//main:workbookProtection')
            remove_nodes(tree, '//main:fileSharing')
            tree.write(workbook_path, encoding='UTF-8', xml_declaration=True, standalone='yes')

        # Create new ZIP in a temp file
        temp_out = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
        temp_out.close()
        with zipfile.ZipFile(temp_out.name, 'w', compression=compression) as zipf:
            for file in all_files:
                abs_path = os.path.join(temp_dir, *file.split('/'))
                if os.path.isfile(abs_path):
                    zipf.write(abs_path, file)
        return temp_out.name

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'files[]' not in request.files:
        return jsonify({'error': 'No files selected'}), 400
    files = request.files.getlist('files[]')
    if not files:
        return jsonify({'error': 'No files selected'}), 400
    download_links = []
    errors = []
    for file in files:
        filename = secure_filename(file.filename)
        if not allowed_file(filename):
            errors.append(f"{filename}: Invalid file type.")
            continue
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_in:
            file.save(temp_in.name)
            temp_in_path = temp_in.name
        if not is_valid_excel(temp_in_path):
            os.remove(temp_in_path)
            errors.append(f"{filename}: Not a valid Excel file.")
            continue
        try:
            processed_file = remove_protection(temp_in_path)
            os.remove(temp_in_path)
            # Change download filename to original name + _unlocked.xlsx
            base, _ = os.path.splitext(filename)
            download_filename = f'{base}_unlocked.xlsx'
            # Session isolation: store processed_file in session
            if 'user_files' not in session:
                session['user_files'] = []
            session['user_files'].append(processed_file)
            session.modified = True
            download_url = url_for('download_file', temp=processed_file, original_name=download_filename)
            download_links.append({'filename': filename, 'download_url': download_url, 'download_filename': download_filename})
        except Exception as e:
            if os.path.exists(temp_in_path):
                os.remove(temp_in_path)
            errors.append(f"{filename}: Error processing file.")
    return jsonify({'downloads': download_links, 'errors': errors})

@app.route('/download')
def download_file():
    temp = request.args.get('temp')
    original_name = request.args.get('original_name', 'unlocked.xlsx')
    # Session isolation: only allow download if file is in user's session
    user_files = session.get('user_files', [])
    if temp and os.path.exists(temp) and temp in user_files:
        response = send_file(temp, as_attachment=True, download_name=original_name)
        @response.call_on_close
        def cleanup():
            try:
                os.remove(temp)
                # Remove from session
                user_files = session.get('user_files', [])
                if temp in user_files:
                    user_files.remove(temp)
                    session['user_files'] = user_files
                    session.modified = True
            except:
                pass
        return response
    abort(403)

# Vercel entry point
# Export 'app' as the handler
handler = app 