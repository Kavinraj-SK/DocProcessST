"""Doc Processor: local web front end for doc_to_json.py and json_to_db.py.

Run it next to the two scripts:   python app.py     (then open http://127.0.0.1:5055)

It does not change how the scripts work. It calls their own functions one file at a time,
so it can show "file k of N", and it asks RefreshImageUrls for fresh links to show the images.
"""
import base64
import json
import os
import re
import sys
import threading
import time
import webbrowser
from collections import Counter, deque
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

import requests
from flask import Flask, jsonify, request, send_from_directory

import doc_to_json as d
import json_to_db as j

app = Flask(__name__, static_folder=str(HERE / 'static'), static_url_path='/static')
app.config['MAX_CONTENT_LENGTH'] = 1024 * 1024 * 1024

SETTINGS_PATH = HERE / 'dashboard_settings.state'
STATE_NAME = '_dashboard_state.state'   # not .json, so json_to_db.py never mistakes it for a question file
USER_AGENT = j.USER_AGENT

# ----------------------------------------------------------------------------- settings
DEFAULT_SETTINGS = {
    'mode': 'docx',                      # 'docx' = Word documents, 'csv' = CSV / Excel files
    'folder': os.getenv('DOCPROC_FOLDER') or '',
    'csv_folder': os.getenv('CSV_FOLDER') or '',
    'api_url': j.API_URL,
    'exam_id': j.EXAM_ID_DEFAULT,
    'qb_config_id': j.QB_CONFIG_ID,
    'mark': j.QUESTION_MARK,
    'ques_type_id': j.QUES_TYPE_ID,
    'delay': 0.6,
    'subj_ids': os.getenv('SUBJ_IDS', ''),
}
SETTINGS = dict(DEFAULT_SETTINGS)
if SETTINGS_PATH.exists():
    try:
        SETTINGS.update(json.loads(SETTINGS_PATH.read_text(encoding='utf-8')))
    except Exception:
        pass

TOKEN = {'value': os.getenv('ques_TOKEN') or ''}


def _num(value, cast, name):
    try:
        return cast(value)
    except (TypeError, ValueError):
        raise ValueError(f"{name} must be a number")


def apply_settings():
    s = SETTINGS
    j.API_URL = str(s['api_url']).strip()
    base = j.API_URL.rsplit('/', 1)[0]
    if not os.getenv('POLY_SUBJECTS_URL'):
        j.SUBJECTS_URL = base + '/GetAllPracticalSubjects'
    if not os.getenv('POLY_UPLOAD_URL'):
        d.POLY_UPLOAD_URL = base + '/UploadQuestionImages'
    if not os.getenv('POLY_REFRESH_URL'):
        d.POLY_REFRESH_URL = base + '/RefreshImageUrls'
    j.EXAM_ID_DEFAULT = int(s['exam_id'])
    j.QB_CONFIG_ID = int(s['qb_config_id'])
    j.QUES_TYPE_ID = int(s['ques_type_id'])
    m = float(s['mark'])
    j.QUESTION_MARK = int(m) if m == int(m) else m
    ids = {}
    for pair in str(s.get('subj_ids') or '').replace('\n', ',').split(','):
        if '=' in pair:
            code, sid = pair.split('=', 1)
            if code.strip() and sid.strip().isdigit():
                ids[code.strip()] = int(sid.strip())
    j.ENV_SUBJ_IDS.clear()
    j.ENV_SUBJ_IDS.update(ids)


def apply_token():
    value = TOKEN['value'].strip()
    d.POLY_TOKEN = value or None
    j.QUES_TOKEN = value or None


def token_info():
    value = TOKEN['value'].strip()
    info = {'present': bool(value), 'exp': None, 'seconds_left': None, 'readable': False}
    if not value:
        return info
    try:
        payload = value.split('.')[1]
        payload += '=' * (-len(payload) % 4)
        data = json.loads(base64.urlsafe_b64decode(payload))
        exp = data.get('exp')
        info['readable'] = True
        if exp:
            info['exp'] = int(exp)
            info['seconds_left'] = int(exp - time.time())
    except Exception:
        pass
    return info


def save_settings():
    SETTINGS_PATH.write_text(json.dumps(SETTINGS, indent=2), encoding='utf-8')


apply_settings()
apply_token()

# ----------------------------------------------------------------------------- files and state
INFO_CACHE = {}      # json path -> (mtime_ns, info, media)
STATE = {}           # stem -> {'convert': {...}, 'save': {...}, 'images': {...}, 'messages': [...]}
STATE_LOCK = threading.Lock()
LOADED_STATE_FOR = {'folder': None}


def current_mode():
    return 'csv' if SETTINGS.get('mode') == 'csv' else 'docx'


def folder_path():
    """The folder chosen in the front end for the current mode, or None while none is chosen."""
    raw = str(SETTINGS.get('csv_folder' if current_mode() == 'csv' else 'folder') or '').strip().strip('"')
    return Path(raw) if raw else None


def state_file():
    return folder_path() / STATE_NAME


def load_state():
    folder = str(folder_path())
    if folder_path() is None:
        return
    if LOADED_STATE_FOR['folder'] == folder:
        return
    with STATE_LOCK:
        STATE.clear()
        try:
            STATE.update(json.loads(state_file().read_text(encoding='utf-8')))
        except Exception:
            pass
        LOADED_STATE_FOR['folder'] = folder


def persist_state():
    try:
        with STATE_LOCK:
            state_file().write_text(json.dumps(STATE, ensure_ascii=False), encoding='utf-8')
    except Exception:
        pass


def scan(folder):
    """{stem: {'docx': Path|None, 'json': Path|None}} for the working folder."""
    files = {}
    if not folder.is_dir():
        return files
    if current_mode() == 'csv':
        # CSV / Excel mode: the 'docx' slot holds the .csv or .xlsx source (same shape, so the table is shared)
        for p in sorted(folder.iterdir()):
            if p.is_file() and p.suffix.lower() in SOURCE_EXTENSIONS and not p.name.startswith('~$'):
                files.setdefault(p.stem, {'docx': None, 'json': None})['docx'] = p
        for p in sorted(folder.glob('*.json')):
            if p.stem in files:
                files[p.stem]['json'] = p
        return files
    for p in sorted(folder.rglob('*.docx')):
        if p.name.startswith('~$'):
            continue
        files.setdefault(p.stem, {'docx': None, 'json': None})['docx'] = p
    for p in sorted(folder.glob('*.json')):
        files.setdefault(p.stem, {'docx': None, 'json': None})['json'] = p
    return files


def count_questions(node):
    n = 1 if node.get('Description') is not None else 0
    for child in (node.get('SubDivisions') or []) + (node.get('SubQuestions') or []):
        n += count_questions(child)
    return n


def json_info(path):
    """Summary of one converted JSON file (cached by modification time)."""
    mtime = path.stat().st_mtime_ns
    cached = INFO_CACHE.get(str(path))
    if cached and cached[0] == mtime:
        return cached[1], cached[2]
    try:
        data = json.loads(path.read_text(encoding='utf-8'))
    except Exception as e:
        info = {'error': f'unreadable JSON: {e}'}
        INFO_CACHE[str(path)] = (mtime, info, [])
        return info, []
    if isinstance(data, list):       # CSV / Excel mode: one row = one question
        code = ''
        for row in data:
            if isinstance(row, dict):
                code = next((str(row[k]).strip() for k in SUBCODE_KEYS if row.get(k) not in (None, '')), '')
                if code:
                    break
        m = re.search(r'\d{8,}', path.name)
        code = code or (m.group(0) if m else '')
        info = {'subj_code': code, 'questions': len(data), 'images_total': 0, 'images_uploaded': 0,
                'placeholders': 0, 'flags': 0}
        INFO_CACHE[str(path)] = (mtime, info, [])
        return info, []
    questions = data.get('questions') or []
    media = [m for m in (data.get('_embedded_media') or []) if isinstance(m, dict)]
    uploaded = [m for m in media if m.get('filePath')]
    text = json.dumps(questions, ensure_ascii=False)
    info = {
        'subj_code': data.get('subj_code', ''),
        'questions': sum(count_questions(q) for q in questions),
        'images_total': len(media),
        'images_uploaded': len(uploaded),
        'placeholders': text.count('----media/'),
        'flags': len(data.get('_review_flags') or []),
    }
    INFO_CACHE[str(path)] = (mtime, info, uploaded)
    return info, uploaded


def human_size(n):
    for unit in ('B', 'KB', 'MB', 'GB'):
        if n < 1024 or unit == 'GB':
            return f"{n:.0f} {unit}" if unit == 'B' else f"{n:.2f} {unit}"
        n /= 1024


# ----------------------------------------------------------------------------- activity log
LOG = deque(maxlen=1500)
LOG_SEQ = [0]
CURRENT = {'stem': ''}
AUTH_FAIL = threading.Event()


def log(*args, **kwargs):
    """Stands in for print() inside doc_to_json and json_to_db while a job runs."""
    msg = ' '.join(str(a) for a in args).rstrip()
    if not msg.strip() or msg.startswith('[DRY RUN]'):
        return
    low = msg.lower()
    if 'failed' in low or '[error]' in low:
        level = 'error'
    elif '[warn]' in low or 'duplicate' in low or 'skipping' in low or '[retry]' in low:
        level = 'warn'
    else:
        level = 'info'
    if 'auth error 401' in low or 'auth error 403' in low or '401 unauthorized' in low:
        AUTH_FAIL.set()
    LOG_SEQ[0] += 1
    LOG.append({'n': LOG_SEQ[0], 't': datetime.now().strftime('%H:%M:%S'), 'stem': CURRENT['stem'], 'level': level, 'msg': msg.strip()})
    stem = CURRENT['stem']
    if stem:
        with STATE_LOCK:
            msgs = STATE.setdefault(stem, {}).setdefault('messages', [])
            if len(msgs) < 80:
                msgs.append(msg.strip())


d.print = log
j.print = log

# ----------------------------------------------------------------------------- your CSV / Excel scripts
# Your own csv_to_json.py and MCQ uploader are used as they are: they are never edited. Put both in
# this folder. The page finds them by what they contain, so their file names do not matter.
SOURCE_EXTENSIONS = ('.csv', '.xlsx')
SUBCODE_KEYS = ('SUBCODE', 'Subcode', 'SubCode', 'Subject Code')
OWN_FILES = {'app.py', 'doc_to_json.py', 'json_to_db.py'}
CSV_COUNT = Counter()      # per-file counters filled by the wrapped uploader functions
UPLOADER = {'path': None, 'mtime': None, 'mod': None}


def find_script(*needles, prefer=()):
    candidates = []
    for p in sorted(HERE.glob('*.py')):
        if p.name in OWN_FILES:
            continue
        try:
            text = p.read_text(encoding='utf-8', errors='replace')
        except OSError:
            continue
        if all(n in text for n in needles):
            candidates.append(p)
    for name in prefer:
        for p in candidates:
            if p.name == name:
                return p
    return candidates[0] if candidates else None


def csv_converter_path():
    return find_script('csv.DictReader', 'folder_path', prefer=('csv_to_json.py',))


def csv_uploader_path():
    return find_script('def run_one_file', 'def save_question')


def csv_scripts_error(need_convert, need_save):
    if need_convert and csv_converter_path() is None:
        return 'Put your csv_to_json.py in the Doc Processor folder (next to app.py), then try again.'
    if need_save and csv_uploader_path() is None:
        return 'Put your CSV uploader script (the one with run_one_file) in the Doc Processor folder (next to app.py), then try again.'
    return None


def run_csv_converter(folder):
    """Run your csv_to_json.py on the chosen folder. Its source is run as written, only its folder line is pointed at the page's folder."""
    path = csv_converter_path()
    source = path.read_text(encoding='utf-8')
    source, n = re.subn(r'(?m)^(\s*)folder_path\s*=.*$', lambda m: f"{m.group(1)}folder_path={str(folder)!r}", source, count=1)
    if not n:
        raise RuntimeError(f'Could not find the "folder_path = ..." line in {path.name}.')
    namespace = {'__name__': 'csv_to_json_run', '__file__': str(path), 'print': log}
    exec(compile(source, str(path), 'exec'), namespace)
    repair_csv_json(folder)


def _unmojibake(text):
    """Undo UTF-8 text that was read as latin-1 (e.g. a BOM turning the first header into 'ï»¿SUBCODE'). Real latin-1 text is left alone."""
    if isinstance(text, str) and any(0x80 <= ord(ch) <= 0xFF for ch in text):
        try:
            fixed = text.encode('latin-1').decode('utf-8')
            return fixed.lstrip('\ufeff')
        except (UnicodeEncodeError, UnicodeDecodeError):
            return text
    return text


def repair_csv_json(folder):
    """Your csv_to_json.py reads CSV files as latin-1. Fix what that does to UTF-8 CSVs, in the JSON it wrote."""
    for src in folder.glob('*.csv'):
        jp = folder / f"{src.stem}.json"
        if not jp.exists():
            continue
        try:
            rows = json.loads(jp.read_text(encoding='utf-8'))
            fixed = [{_unmojibake(k): _unmojibake(v) for k, v in row.items()} if isinstance(row, dict) else row for row in rows]
            if fixed != rows:
                jp.write_text(json.dumps(fixed, indent=4, ensure_ascii=False), encoding='utf-8')
                log(f"{src.name}: repaired UTF-8 characters that were read as latin-1")
        except Exception as e:
            log(f"[WARN] {src.name}: could not check the converted JSON ({e})")


def load_uploader():
    """Import your uploader and point it at this page's settings without editing the file."""
    path = csv_uploader_path()
    mtime = path.stat().st_mtime_ns
    if UPLOADER['path'] != path or UPLOADER['mtime'] != mtime:
        import importlib.util
        spec = importlib.util.spec_from_file_location('mcq_uploader_user', str(path))
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        missing = [n for n in ('run_one_file', 'build_payload', 'save_question') if not hasattr(mod, n)]
        if missing:
            raise RuntimeError(f"{path.name} has no function named {', '.join(missing)}.")
        original_build, original_save = mod.build_payload, mod.save_question

        def build_payload(*args, **kwargs):
            payload = original_build(*args, **kwargs)
            CSV_COUNT['built'] += 1
            code = str(payload.get('subj_Code') or '').strip()
            if code in j.SUBJ_ID_MAP:
                payload['subj_Id'] = j.SUBJ_ID_MAP[code]
            else:
                j.UNMAPPED_CODES.add(code)
            return payload

        def save_question(session, payload):
            if AUTH_FAIL.is_set():
                raise RuntimeError('401 Unauthorized -- stopped, the token was rejected')
            try:
                return original_save(session, payload)
            finally:
                if float(SETTINGS['delay']) > 0:
                    time.sleep(float(SETTINGS['delay']))

        def counting_print(*args, **kwargs):
            msg = ' '.join(str(a) for a in args)
            if re.match(r'\s*\[row \d+\] FAILED', msg):
                CSV_COUNT['failed'] += 1
            log(*args)

        mod.build_payload, mod.save_question, mod.print = build_payload, save_question, counting_print
        UPLOADER.update(path=path, mtime=mtime, mod=mod)
    mod = UPLOADER['mod']
    mod.API_BASE_URL = j.API_URL
    mod.EXAM_ID_DEFAULT = int(SETTINGS['exam_id'])
    return mod


# ----------------------------------------------------------------------------- the ingestion job
JOB_LOCK = threading.Lock()
CANCEL = threading.Event()


def new_job():
    return {'status': 'idle', 'stage': '', 'current': '', 'current_size': '', 'index': 0, 'total_files': 0,
            'files_done': 0, 'steps_done': 0, 'steps_total': 0, 'percent': 0.0, 'started': None, 'finished': None,
            'eta_seconds': None, 'files_per_min': 0.0, 'message': '', 'notes': [], 'options': {}}


JOB = new_job()


def job_update(**kw):
    with JOB_LOCK:
        JOB.update(kw)
        if JOB['steps_total']:
            JOB['percent'] = round(100.0 * JOB['steps_done'] / JOB['steps_total'], 1)
        if JOB['started'] and JOB['status'] == 'running':
            elapsed = max(time.time() - JOB['started'], 0.001)
            JOB['files_per_min'] = round(JOB['files_done'] / elapsed * 60, 1)
            if JOB['steps_done'] > 0:
                JOB['eta_seconds'] = int(elapsed / JOB['steps_done'] * (JOB['steps_total'] - JOB['steps_done']))


def run_job(stems, opts):
    folder = folder_path()
    files = scan(folder)
    do_convert, do_save, dry = opts['convert'], opts['save'], opts['dry_run']
    csv_mode = current_mode() == 'csv'
    stems = [s for s in stems if s in files]
    per_file = int(do_convert) + int(do_save)
    CANCEL.clear()
    AUTH_FAIL.clear()
    j.UNMAPPED_CODES.clear()
    with JOB_LOCK:
        JOB.clear()
        JOB.update(new_job())
        JOB.update(status='running', total_files=len(stems), steps_total=len(stems) * per_file,
                   started=time.time(), options=opts, message='Starting')
    session = requests.Session()
    bad_codes = set()
    final_status, final_message = 'done', ''
    try:
        csv_uploader = None
        csv_status = {}
        if csv_mode:
            if do_convert:
                job_update(stage='Converting')
                t0 = time.time() - 2
                run_csv_converter(folder)
                for s_ in stems:
                    jp = folder / f"{s_}.json"
                    csv_status[s_] = 'ok' if jp.exists() and jp.stat().st_mtime >= t0 else 'error'
            if do_save:
                csv_uploader = load_uploader()
                session.headers.update({'Accept': '*/*', 'Authorization': f"Bearer {TOKEN['value'].strip()}",
                                        'Content-Type': 'application/json', 'User-Agent': USER_AGENT})
        if do_save and not dry:
            job_update(stage='Loading subject list')
            j.SUBJ_ID_MAP.clear()
            j.SUBJ_ID_MAP.update(j.fetch_subject_map(session))
            j.SUBJ_ID_MAP.update(j.ENV_SUBJ_IDS)
        for i, stem in enumerate(stems):
            if CANCEL.is_set():
                final_status, final_message = 'cancelled', 'Stopped before the next file'
                break
            entry = files[stem]
            CURRENT['stem'] = stem
            source = entry['docx'] or entry['json']
            size = human_size(source.stat().st_size) if source and source.exists() else ''
            with STATE_LOCK:
                STATE.setdefault(stem, {})['messages'] = []
            job_update(index=i + 1, current=(entry['docx'] or entry['json']).name, current_size=size)
            convert_ok = True
            if do_convert:
                job_update(stage='Converting')
                if entry['docx'] is None:
                    log(f"[WARN] {stem}: no .docx to convert; using the existing JSON")
                else:
                    try:
                        status = csv_status.get(stem, 'error') if csv_mode else d.process_file(entry['docx'], folder)
                    except Exception as e:
                        status = 'error'
                        log(f"[ERROR] {stem}: conversion failed: {e}")
                    convert_ok = status == 'ok'
                    with STATE_LOCK:
                        STATE.setdefault(stem, {})['convert'] = {'status': status, 'at': time.time()}
                job_update(steps_done=JOB['steps_done'] + 1)
            json_path = folder / f"{stem}.json"
            if do_save:
                job_update(stage='Saving to question bank' if not dry else 'Checking (dry run)')
                if not convert_ok or not json_path.exists():
                    with STATE_LOCK:
                        STATE.setdefault(stem, {})['save'] = {'skipped': 'no converted JSON', 'at': time.time()}
                else:
                    stats = Counter()
                    try:
                        if csv_mode:
                            CSV_COUNT.clear()
                            try:
                                ok_n, dup_n, errs = csv_uploader.run_one_file(session, json_path, dry, 1, None)
                            except Exception:
                                stats['error'] += max(0, CSV_COUNT['failed'] - 1)
                                raise
                            if dry:
                                stats['saved'], stats['error'] = CSV_COUNT['built'], CSV_COUNT['failed']
                            else:
                                stats['saved'], stats['duplicate'], stats['error'] = ok_n, dup_n, len(errs)
                        else:
                            j.process_file(json_path, session, stats, dry, False, float(SETTINGS['delay']), bad_codes)
                    except Exception as e:
                        stats['error'] += 1
                        log(f"[ERROR] {stem}: save failed: {e}")
                    with STATE_LOCK:
                        STATE.setdefault(stem, {})['save'] = {
                            'saved': stats['saved'], 'duplicate': stats['duplicate'], 'error': stats['error'],
                            'children_skipped': stats['children skipped (parent duplicate)'],
                            'subject_missing': stats['files skipped (subject not found)'],
                            'dry_run': dry, 'at': time.time()}
                job_update(steps_done=JOB['steps_done'] + 1)
            job_update(files_done=JOB['files_done'] + 1)
            persist_state()
            if AUTH_FAIL.is_set():
                final_status = 'error'
                final_message = 'The server rejected the token (401/403). Paste a fresh token in Settings and run again.'
                break
    except Exception as e:
        final_status, final_message = 'error', f'Stopped by an unexpected error: {e}'
        log(f"[ERROR] {final_message}")
    finally:
        CURRENT['stem'] = ''
        persist_state()
        notes = []
        if j.UNMAPPED_CODES:
            notes.append('No numeric subject id for: ' + ', '.join(sorted(c or '(empty)' for c in j.UNMAPPED_CODES)[:20])
                         + '. Those files were sent as id 0 plus code, so the duplicate check may not see them.')
        with JOB_LOCK:
            JOB['status'] = final_status
            JOB['finished'] = time.time()
            JOB['stage'] = ''
            JOB['message'] = final_message
            JOB['notes'] = notes
            JOB['eta_seconds'] = 0
            if final_status == 'done':
                JOB['percent'] = 100.0


# ----------------------------------------------------------------------------- image links
URL_CACHE = {}   # filePath -> (url, time)


def refresh_urls(paths):
    """Ask RefreshImageUrls for fresh signed links. Returns ({path: url or None}, error or None)."""
    out, error = {}, None
    now = time.time()
    todo = []
    for p in dict.fromkeys(paths):
        cached = URL_CACHE.get(p)
        if cached and now - cached[1] < 600:
            out[p] = cached[0]
        else:
            todo.append(p)
    headers = {'accept': '*/*', 'Authorization': f"Bearer {TOKEN['value']}", 'User-Agent': USER_AGENT}
    for i in range(0, len(todo), 25):
        chunk = todo[i:i + 25]
        try:
            resp = requests.post(d.POLY_REFRESH_URL, headers=headers, json=chunk, timeout=60)
            body = resp.json()
        except Exception as e:
            error = f"Could not reach RefreshImageUrls: {e}"
            for p in chunk:
                out[p] = None
            continue
        if resp.status_code in (401, 403):
            error = 'The server rejected the token. Paste a fresh token in Settings.'
        got = {}
        for img in (body.get('images') if isinstance(body, dict) else None) or []:
            fp, url = img.get('filePath'), img.get('imageUrl')
            if fp:
                got[fp] = url if isinstance(url, str) and url.lower().startswith('http') else None
        for p in chunk:
            url = got.get(p)
            out[p] = url
            if url:
                URL_CACHE[p] = (url, now)
    return out, error


def collect_images(stems=None):
    folder = folder_path()
    items = []
    if folder is None:
        return items
    for stem, entry in scan(folder).items():
        if stems and stem not in stems:
            continue
        if not entry['json']:
            continue
        _, uploaded = json_info(entry['json'])
        for m in uploaded:
            items.append({'stem': stem, 'fileName': m.get('fileName') or Path(m['filePath']).name, 'filePath': m['filePath']})
    return items


# ----------------------------------------------------------------------------- routes
@app.get('/')
def index():
    return send_from_directory(str(HERE / 'static'), 'index.html')


@app.get('/api/settings')
def api_get_settings():
    return jsonify({'settings': SETTINGS, 'token': token_info()})


@app.post('/api/settings')
def api_set_settings():
    body = request.get_json(silent=True) or {}
    new = dict(SETTINGS)
    try:
        for key in ('folder', 'csv_folder', 'api_url', 'subj_ids'):
            if key in body:
                new[key] = str(body[key]).strip().strip('"')
        for key, cast in (('exam_id', int), ('qb_config_id', int), ('ques_type_id', int), ('mark', float), ('delay', float)):
            if key in body:
                new[key] = _num(body[key], cast, key.replace('_', ' '))
        if not new['api_url'].lower().startswith('http'):
            raise ValueError('API URL must start with http')
        for fkey in ('folder', 'csv_folder'):
            if new[fkey] and not Path(new[fkey]).is_dir():
                raise ValueError('That folder does not exist. Check the path and try again.')
    except (ValueError, OSError) as e:
        return jsonify({'error': str(e)}), 400
    if JOB['status'] == 'running':
        return jsonify({'error': 'Wait for the running ingestion to finish before changing settings.'}), 409
    SETTINGS.update(new)
    try:
        apply_settings()
    except (ValueError, TypeError) as e:
        return jsonify({'error': f'Invalid setting: {e}'}), 400
    save_settings()
    load_state()
    return jsonify({'settings': SETTINGS})


@app.post('/api/token')
def api_set_token():
    body = request.get_json(silent=True) or {}
    token = str(body.get('token', '')).strip()
    if token.lower().startswith('bearer '):
        token = token[7:].strip()
    TOKEN['value'] = token
    apply_token()
    URL_CACHE.clear()
    if body.get('save_env'):
        env_path = HERE / '.env'
        lines = env_path.read_text(encoding='utf-8').splitlines() if env_path.exists() else []
        lines = [ln for ln in lines if not ln.startswith('ques_TOKEN=')] + [f'ques_TOKEN={token}']
        env_path.write_text('\n'.join(lines) + '\n', encoding='utf-8')
    return jsonify({'token': token_info()})


@app.get('/api/token')
def api_token():
    return jsonify({'token': token_info()})


@app.post('/api/mode')
def api_set_mode():
    mode = (request.get_json(silent=True) or {}).get('mode')
    if mode not in ('docx', 'csv'):
        return jsonify({'error': 'Mode must be docx or csv.'}), 400
    if JOB['status'] == 'running':
        return jsonify({'error': 'Wait for the running ingestion to finish before switching.'}), 409
    SETTINGS['mode'] = mode
    save_settings()
    load_state()
    return jsonify({'settings': SETTINGS})


@app.get('/api/rows')
def api_rows():
    load_state()
    folder = folder_path()
    if folder is None:
        return jsonify({'folder': '', 'folder_exists': False, 'rows': [], 'totals': {}})
    files = scan(folder)
    rows = []
    totals = Counter()
    for stem, entry in files.items():
        row = {'stem': stem, 'docx': entry['docx'].name if entry['docx'] else None,
               'size': human_size(entry['docx'].stat().st_size) if entry['docx'] else '',
               'has_json': bool(entry['json']), 'info': None, 'updated': None}
        if entry['json']:
            row['info'], _ = json_info(entry['json'])
            row['updated'] = entry['json'].stat().st_mtime
        with STATE_LOCK:
            row['state'] = json.loads(json.dumps(STATE.get(stem, {})))
        rows.append(row)
        totals['files'] += 1
        if entry['json'] and row['info'] and 'error' not in row['info']:
            totals['converted'] += 1
            totals['questions'] += row['info']['questions']
            totals['images'] += row['info']['images_uploaded']
        save = row['state'].get('save') or {}
        if not save.get('dry_run'):
            totals['saved'] += save.get('saved', 0)
        totals['duplicates'] += save.get('duplicate', 0)
        totals['errors'] += save.get('error', 0)
        imgs = row['state'].get('images') or {}
        totals['images_ok'] += imgs.get('ok', 0)
        totals['images_missing'] += imgs.get('missing', 0)
    return jsonify({'folder': str(folder), 'folder_exists': folder.is_dir(), 'rows': rows, 'totals': dict(totals)})


@app.post('/api/upload')
def api_upload():
    folder = folder_path()
    if folder is None or not folder.is_dir():
        return jsonify({'error': 'Choose a folder first.'}), 400
    saved, skipped = [], []
    for f in request.files.getlist('files'):
        name = Path(f.filename or '').name
        wanted = SOURCE_EXTENSIONS if current_mode() == 'csv' else ('.docx',)
        if not name.lower().endswith(wanted) or name.startswith('~$'):
            skipped.append(name or '(unnamed)')
            continue
        f.save(str(folder / name))
        saved.append(name)
    return jsonify({'saved': saved, 'skipped': skipped})


@app.post('/api/pick-folder')
def api_pick_folder():
    """Open the operating system's folder dialog on this computer and return the chosen path."""
    import subprocess
    title = 'Select the folder with your CSV / Excel files' if current_mode() == 'csv' else 'Select the folder with your .docx files'
    code = ("import tkinter as tk\nfrom tkinter import filedialog\nr = tk.Tk(); r.withdraw(); r.attributes('-topmost', True)\n"
            f"print(filedialog.askdirectory(title={title!r}) or '')")
    try:
        out = subprocess.run([sys.executable, '-c', code], capture_output=True, text=True, timeout=300)
    except Exception as e:
        return jsonify({'error': f'Could not open the folder dialog: {e}'}), 500
    if out.returncode != 0:
        return jsonify({'error': 'The folder dialog is not available here. Type the folder path instead.'}), 500
    return jsonify({'folder': out.stdout.strip()})


@app.post('/api/run')
def api_run():
    body = request.get_json(silent=True) or {}
    if JOB['status'] == 'running':
        return jsonify({'error': 'An ingestion is already running.'}), 409
    do_convert, do_save = bool(body.get('convert', True)), bool(body.get('save', True))
    dry = bool(body.get('dry_run', False))
    if not (do_convert or do_save):
        return jsonify({'error': 'Choose at least one step: convert or save.'}), 400
    needs_token = (do_save and not dry) or (do_convert and current_mode() != 'csv')   # Word conversion uploads images
    if needs_token:
        info = token_info()
        if not info['present']:
            return jsonify({'error': 'Paste your token in Settings first.'}), 400
        if info['seconds_left'] is not None and info['seconds_left'] <= 0:
            return jsonify({'error': 'The token has expired. Paste a fresh token in Settings.'}), 400
    if folder_path() is None or not folder_path().is_dir():
        return jsonify({'error': 'Choose a folder first.'}), 400
    if current_mode() == 'csv':
        problem = csv_scripts_error(do_convert, do_save)
        if problem:
            return jsonify({'error': problem}), 400
    files = scan(folder_path())
    stems = body.get('stems') or list(files)
    stems = [s for s in stems if s in files]
    if body.get('limit'):
        try:
            stems = stems[:max(1, int(body['limit']))]
        except (TypeError, ValueError):
            return jsonify({'error': 'Limit must be a whole number.'}), 400
    if not stems:
        return jsonify({'error': 'No files to process in the folder.'}), 400
    opts = {'convert': do_convert, 'save': do_save, 'dry_run': dry}
    threading.Thread(target=run_job, args=(stems, opts), daemon=True).start()
    time.sleep(0.05)
    return jsonify({'started': True, 'files': len(stems)})


@app.post('/api/stop')
def api_stop():
    CANCEL.set()
    return jsonify({'stopping': True})


@app.get('/api/job')
def api_job():
    with JOB_LOCK:
        job = dict(JOB)
    job['now'] = time.time()
    return jsonify(job)


@app.get('/api/log')
def api_log():
    since = int(request.args.get('since', 0) or 0)
    return jsonify({'lines': [ln for ln in list(LOG) if ln['n'] > since], 'last': LOG_SEQ[0]})


@app.get('/api/images')
def api_images():
    stem = request.args.get('stem')
    items = collect_images([stem] if stem else None)
    folders = sorted({str(Path(i['filePath']).parent).replace('\\', '/') for i in items})
    return jsonify({'items': items, 'folders': folders})


@app.post('/api/images/refresh')
def api_images_refresh():
    paths = [p for p in ((request.get_json(silent=True) or {}).get('paths') or []) if isinstance(p, str)]
    if not TOKEN['value']:
        return jsonify({'error': 'Paste your token in Settings first.'}), 400
    urls, error = refresh_urls(paths[:200])
    return jsonify({'urls': urls, 'error': error})


@app.post('/api/images/check')
def api_images_check():
    """Check every image in the folder (or one file) and remember how many have a working link."""
    load_state()
    body = request.get_json(silent=True) or {}
    stem = body.get('stem')
    if not TOKEN['value']:
        return jsonify({'error': 'Paste your token in Settings first.'}), 400
    items = collect_images([stem] if stem else None)
    URL_CACHE.clear()
    urls, error = refresh_urls([i['filePath'] for i in items])
    per, missing_paths = {}, []
    for i in items:
        slot = per.setdefault(i['stem'], {'ok': 0, 'missing': 0})
        if urls.get(i['filePath']):
            slot['ok'] += 1
        else:
            slot['missing'] += 1
            missing_paths.append(i['filePath'])
    with STATE_LOCK:
        for s, c in per.items():
            STATE.setdefault(s, {})['images'] = {**c, 'checked_at': time.time()}
    persist_state()
    return jsonify({'checked': len(items), 'ok': sum(c['ok'] for c in per.values()),
                    'missing': sum(c['missing'] for c in per.values()), 'missing_paths': missing_paths, 'error': error})


def main():
    port = int(os.getenv('PORT', '5055'))
    url = f'http://127.0.0.1:{port}'
    print(f'Doc Processor running at {url}  (choose the folder in the page)')
    if os.getenv('NO_BROWSER') != '1':
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    app.run(host='127.0.0.1', port=port, threaded=True, debug=False)


if __name__ == '__main__':
    main()