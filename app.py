import os
import json
import uuid
import tempfile
import traceback
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, session, jsonify
from web.services.clr_parser import parse_clr, extract_itk_summary
from web.services.transfer_engine import parse_template_itks, transfer_data

app = Flask(__name__, template_folder='web/templates', static_folder='web/static')
app.secret_key = os.environ.get('SECRET_KEY', 'clr-transfer-tool-dev-key-2026')
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB
app.config['UPLOAD_FOLDER'] = os.environ.get('UPLOAD_FOLDER', tempfile.mkdtemp())
app.config['SESSION_COOKIE_SECURE'] = os.environ.get('RAILWAY_ENVIRONMENT') is not None
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)


def _get_job_dir():
    """Get or create a unique job directory for the current session."""
    job_id = session.get('job_id')
    if not job_id:
        job_id = str(uuid.uuid4())[:8]
        session['job_id'] = job_id
    job_dir = os.path.join(app.config['UPLOAD_FOLDER'], f'job_{job_id}')
    os.makedirs(job_dir, exist_ok=True)
    return job_dir


def _save_job_data(key, data):
    """Save large data to a JSON file in the job directory instead of the session cookie."""
    job_dir = _get_job_dir()
    filepath = os.path.join(job_dir, f'{key}.json')
    with open(filepath, 'w') as f:
        json.dump(data, f)


def _load_job_data(key, default=None):
    """Load data from a JSON file in the job directory."""
    job_dir = _get_job_dir()
    filepath = os.path.join(job_dir, f'{key}.json')
    if os.path.exists(filepath):
        with open(filepath) as f:
            return json.load(f)
    return default


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/analyze', methods=['POST'])
def analyze():
    if 'clr_file' not in request.files:
        flash('No file uploaded.', 'error')
        return redirect(url_for('index'))

    clr_file = request.files['clr_file']
    if clr_file.filename == '':
        flash('No file selected.', 'error')
        return redirect(url_for('index'))

    ext = os.path.splitext(clr_file.filename)[1].lower()
    if ext not in ('.xlsx', '.xlsm'):
        flash('Only .xlsx and .xlsm files are supported.', 'error')
        return redirect(url_for('index'))

    # Reset job for fresh upload
    session.pop('job_id', None)
    job_dir = _get_job_dir()

    # Save CLR file
    clr_path = os.path.join(job_dir, 'clr_' + clr_file.filename)
    clr_file.save(clr_path)
    session['clr_path'] = clr_path
    session['clr_filename'] = clr_file.filename

    try:
        clr_data = parse_clr(clr_path)
        itk_summary = extract_itk_summary(clr_data)
        # Store large data in files, not session cookie
        _save_job_data('itk_summary', itk_summary)
        session['total_products'] = clr_data['total_products']
        session['total_parents'] = clr_data['total_parents']
        session['total_children'] = clr_data['total_children']
        session['total_standalone'] = clr_data['total_standalone']
    except Exception as e:
        app.logger.error(f'Error parsing CLR: {traceback.format_exc()}')
        flash(f'Error parsing CLR file: {str(e)}', 'error')
        return redirect(url_for('index'))

    return redirect(url_for('step1_results'))


@app.route('/step1')
def step1_results():
    itk_summary = _load_job_data('itk_summary')
    if not itk_summary:
        flash('Please upload a CLR file first.', 'warning')
        return redirect(url_for('index'))

    # Ensure ITK summary is in the new 3-tuple format (display, raw, count)
    if itk_summary and isinstance(itk_summary[0], (list, tuple)) and len(itk_summary[0]) == 2:
        itk_summary = [(item[0], item[0], item[1]) for item in itk_summary]
        _save_job_data('itk_summary', itk_summary)

    return render_template('step1_results.html',
                           itk_summary=itk_summary,
                           clr_filename=session.get('clr_filename', ''),
                           total_products=session.get('total_products', 0),
                           total_parents=session.get('total_parents', 0),
                           total_children=session.get('total_children', 0),
                           total_standalone=session.get('total_standalone', 0))


@app.route('/step2')
def step2_upload():
    if not session.get('clr_path'):
        flash('Please upload a CLR file first.', 'warning')
        return redirect(url_for('index'))

    itk_summary = _load_job_data('itk_summary', [])
    if itk_summary and isinstance(itk_summary[0], (list, tuple)) and len(itk_summary[0]) == 2:
        itk_summary = [(item[0], item[0], item[1]) for item in itk_summary]
        _save_job_data('itk_summary', itk_summary)

    return render_template('step2_upload.html',
                           clr_filename=session.get('clr_filename', ''),
                           itk_summary=itk_summary)


@app.route('/transfer', methods=['POST'])
def transfer():
    clr_path = session.get('clr_path')
    if not clr_path or not os.path.exists(clr_path):
        flash('CLR file not found. Please start over.', 'error')
        return redirect(url_for('index'))

    template_files = request.files.getlist('template_files')
    if not template_files or all(f.filename == '' for f in template_files):
        flash('Please upload at least one category template file.', 'error')
        return redirect(url_for('step2_upload'))

    # Save template files
    job_dir = _get_job_dir()
    template_paths = []
    for tf in template_files:
        if tf.filename:
            ext = os.path.splitext(tf.filename)[1].lower()
            if ext not in ('.xlsx', '.xlsm'):
                continue
            tp = os.path.join(job_dir, 'tmpl_' + tf.filename)
            tf.save(tp)
            template_paths.append(tp)

    if not template_paths:
        flash('No valid template files (.xlsx/.xlsm) found.', 'error')
        return redirect(url_for('step2_upload'))

    try:
        results = transfer_data(clr_path, template_paths)
        # Store large results in file, not session cookie
        _save_job_data('transfer_results', results['summary'])
        _save_job_data('output_files', results['output_files'])
    except Exception as e:
        app.logger.error(f'Error during transfer: {traceback.format_exc()}')
        flash(f'Error during transfer: {str(e)}', 'error')
        return redirect(url_for('step2_upload'))

    return redirect(url_for('results'))


@app.route('/results')
def results():
    transfer_results = _load_job_data('transfer_results')
    if not transfer_results:
        flash('No transfer results available.', 'warning')
        return redirect(url_for('index'))

    return render_template('results.html',
                           results=transfer_results,
                           clr_filename=session.get('clr_filename', ''))


@app.route('/download/<int:file_index>')
def download_file(file_index):
    output_files = _load_job_data('output_files', [])
    if file_index < 0 or file_index >= len(output_files):
        flash('File not found.', 'error')
        return redirect(url_for('results'))

    file_info = output_files[file_index]
    return send_file(file_info['path'],
                     as_attachment=True,
                     download_name=file_info['filename'])


@app.route('/download-all')
def download_all():
    import zipfile
    output_files = _load_job_data('output_files', [])
    if not output_files:
        flash('No files to download.', 'error')
        return redirect(url_for('results'))

    job_dir = _get_job_dir()
    zip_path = os.path.join(job_dir, 'transferred_templates.zip')
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for fi in output_files:
            zf.write(fi['path'], fi['filename'])

    return send_file(zip_path, as_attachment=True, download_name='transferred_templates.zip')


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5001))
    debug = os.environ.get('RAILWAY_ENVIRONMENT') is None
    app.run(debug=debug, host='0.0.0.0', port=port)
