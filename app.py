import os
import json
import uuid
import tempfile
import traceback
from flask import Flask, render_template, request, redirect, url_for, flash, session
from web.services.clr_parser import parse_clr, extract_itk_summary

app = Flask(__name__, template_folder='web/templates', static_folder='web/static')
app.secret_key = os.environ.get('SECRET_KEY', 'clr-transfer-tool-dev-key-2026')
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB
app.config['UPLOAD_FOLDER'] = os.environ.get('UPLOAD_FOLDER', tempfile.mkdtemp())
app.config['SESSION_COOKIE_SECURE'] = os.environ.get('RAILWAY_ENVIRONMENT') is not None
app.config['SESSION_COOKIE_SAMESITE'] = 'Lax'

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
    job_dir = _get_job_dir()
    with open(os.path.join(job_dir, f'{key}.json'), 'w') as f:
        json.dump(data, f)


def _load_job_data(key, default=None):
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
    session['clr_filename'] = clr_file.filename

    try:
        clr_data = parse_clr(clr_path)
        itk_summary = extract_itk_summary(clr_data)
        _save_job_data('itk_summary', itk_summary)
        session['total_products'] = clr_data['total_products']
        session['total_parents'] = clr_data['total_parents']
        session['total_children'] = clr_data['total_children']
        session['total_standalone'] = clr_data['total_standalone']
    except Exception as e:
        app.logger.error(f'Error parsing CLR: {traceback.format_exc()}')
        flash(f'Error parsing CLR file: {str(e)}', 'error')
        return redirect(url_for('index'))

    return redirect(url_for('results'))


@app.route('/results')
def results():
    itk_summary = _load_job_data('itk_summary')
    if not itk_summary:
        flash('Please upload a CLR file first.', 'warning')
        return redirect(url_for('index'))

    return render_template('results.html',
                           itk_summary=itk_summary,
                           clr_filename=session.get('clr_filename', ''),
                           total_products=session.get('total_products', 0),
                           total_parents=session.get('total_parents', 0),
                           total_children=session.get('total_children', 0),
                           total_standalone=session.get('total_standalone', 0))


@app.errorhandler(500)
def internal_error(e):
    app.logger.error(f'Internal Server Error: {e}')
    return render_template('error.html', error=str(e)), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5001))
    debug = os.environ.get('RAILWAY_ENVIRONMENT') is None
    app.run(debug=debug, host='0.0.0.0', port=port)
