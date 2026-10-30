import argparse, json, os, re, sys, time
from collections import Counter
from pathlib import Path
import requests
from dotenv import load_dotenv

load_dotenv()

API_URL = os.getenv('API_BASE_URL', 'https://polyexam.tnedu.in/Dote/api/QuestionBank/SaveQuestion')
EXAM_ID_DEFAULT = int(os.getenv('EXAM_ID_DEFAULT', '1'))
QB_CONFIG_ID = int(os.getenv('QB_CONFIG_ID', '17'))
_qm = float(os.getenv('QUESTION_MARK', '7'))
QUESTION_MARK = int(_qm) if _qm == int(_qm) else _qm
QUES_TYPE_ID = int(os.getenv('QUES_TYPE_ID', '10'))
ADD_DEFAULT_ANSWER = os.getenv('ADD_DEFAULT_ANSWER', '1') == '1'
# Numeric subject ids, like the frontend sends. Two sources, merged at the start of a run:
#   1. GetAllPracticalSubjects (one call per run), the same list the frontend dropdown uses
#   2. optional .env override, which wins:  SUBJ_IDS=MA232431=2827,SU232460=1234
# A subject code with no id falls back to subj_Id 0 + subj_Code (the old behaviour).
ENV_SUBJ_IDS = {}
for _pair in os.getenv('SUBJ_IDS', '').split(','):
    if '=' in _pair:
        _code, _sid = _pair.split('=', 1)
        if _code.strip() and _sid.strip().isdigit():
            ENV_SUBJ_IDS[_code.strip()] = int(_sid.strip())
SUBJ_ID_MAP = dict(ENV_SUBJ_IDS)
SUBJECTS_URL = os.getenv('POLY_SUBJECTS_URL', API_URL.rsplit('/', 1)[0] + '/GetAllPracticalSubjects')
IMG_STYLE = 'max-width: 350px; max-height: 300px; width: auto; height: auto; object-fit: contain; display: block; margin: 8px auto;'
DEFAULT_FOLDER = 'C:\\Users\\Administrator\\Downloads\\23PT17\\zDocProcessor-main\\R2023 All Sem Missing Practicum'
QUES_TOKEN = os.getenv('ques_TOKEN')
SAVE_TIMEOUT = (float(os.getenv('SAVE_CONNECT_TIMEOUT', '10')), float(os.getenv('SAVE_READ_TIMEOUT', '180')))
USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'
IMAGE_MARKER_RE = re.compile(r'----media/([A-Za-z0-9_.-]+)----')
SAFE_TAG_RE = re.compile(r'<(?!/?(?:p|br|span|table|tbody|thead|tr|td|th|img|sup|sub|b|i|u|strong|em)\b)')


def fetch_subject_map(session):
    """One GET to the same list the frontend dropdown uses. Returns {subject code: numeric id}."""
    headers = {'accept': '*/*', 'Authorization': f"Bearer {QUES_TOKEN}", 'User-Agent': USER_AGENT}
    try:
        resp = session.get(SUBJECTS_URL, params={'examId': EXAM_ID_DEFAULT}, headers=headers, timeout=SAVE_TIMEOUT)
        body = resp.json()
    except Exception as e:
        print(f"[WARN] could not load the subject list ({e}); using SUBJ_IDS from .env only")
        return {}
    rows = (body.get('result') or body.get('data') or []) if isinstance(body, dict) else []
    found = {}
    for r in rows:
        if not isinstance(r, dict):
            continue
        low = {str(k).lower(): v for k, v in r.items()}
        code = low.get('subj_code') or low.get('subjcode')
        sid = low.get('id') or low.get('subj_id') or low.get('subjid')
        if code and isinstance(sid, int):
            found[str(code).strip()] = sid
    print(f"Subject list: {len(found)} subject(s) from the server")
    return found


def sync_image_urls(data):
    patterns = []
    for m in data.get('_embedded_media') or []:
        fp = m.get('filePath')
        if not fp:
            continue
        base = Path(fp).name
        alt = m.get('fileName') or base
        tag = f'<img src="{m.get("imageUrl", "")}" data-file-path="{fp}" alt="{alt}" style="{IMG_STYLE}" />'
        patterns.append((re.compile(r'<img\b[^>]*' + re.escape(base) + r'[^>]*>'), tag))
    if not patterns:
        return

    def walk(node):
        desc = node.get('Description')
        if desc is not None:
            for pat, tag in patterns:
                desc = pat.sub(lambda _m, t=tag: t, desc)
            node['Description'] = desc
        for child in (node.get('SubDivisions') or []) + (node.get('SubQuestions') or []):
            walk(child)

    for q in data.get('questions', []):
        walk(q)


def to_katex_html(text):
    return SAFE_TAG_RE.sub('&lt;', text)


_TABLE_RE = re.compile(r'(<table\b.*?</table>)', re.S)


def format_ques_content(description):
    text = (description or '').replace('\r\n', '\n').strip()
    text = to_katex_html(text)
    parts = []
    for seg in _TABLE_RE.split(text):
        if seg.startswith('<table'):
            parts.append(seg)
        else:
            parts.extend(f'<p>{ln.strip()}</p>' for ln in seg.split('\n') if ln.strip())
    return ''.join(parts) or '<p></p>'


def build_payload(subj_code, content_html, parent_id, disp_order):
    # hash_Ques and every enc_*/orig_* field are computed by the server in SaveQuestionAsync
    # (it overwrites whatever the client sends), so they are sent as null.
    answers = [{'ans_Content': '', 'is_Correct': True, 'display_Order': 1, 'stepMarks': []}] if ADD_DEFAULT_ANSWER else []
    # Same as the frontend: numeric subj_Id and subj_Code null when the id is known.
    subj_id = SUBJ_ID_MAP.get(subj_code, 0)
    return {
        'ques_Id': None,
        'qbConfig_ID': QB_CONFIG_ID,
        'exam_ID': EXAM_ID_DEFAULT,
        'subj_Id': subj_id,
        'subj_Code': None if subj_id else subj_code,
        'unit_Id': None,
        'co_Id': None,
        'topic_Id': None,
        'mark': QUESTION_MARK,
        'parntQus_Id': parent_id,
        'ques_Lvl': None,
        'disp_Order': disp_order,
        'quesType_Id': QUES_TYPE_ID,
        'bloomLvl_Id': None,
        'diffLvl_Id': None,
        'ques_Content': content_html,
        'hash_Ques': None,
        'enc_Text_Ques': None,
        'enc_Key_Ques': None,
        'nonce_Ques': None,
        'enc_Tag_Ques': None,
        'orig_Enc_Text': None,
        'orig_Enc_Key': None,
        'orig_Nonce': None,
        'orig_Enc_Tag': None,
        'answers': answers,
        'stepMarks': [],
    }


def extract_new_id(response_json):
    if response_json is None:
        return None
    if isinstance(response_json, int):
        return response_json
    if isinstance(response_json, dict):
        for key in ('ques_Id', 'Ques_Id', 'quesId', 'id', 'Id', 'ID', 'data'):
            if key in response_json:
                val = response_json[key]
                if isinstance(val, dict):
                    return extract_new_id(val)
                if isinstance(val, int):
                    return val
    return None


TAG_RE = re.compile(r'<[^>]+>')


def _dup_details(body):
    existing = body.get('existingQuestion') or body.get('existingQuestions') or body.get('duplicateQuestions') or []
    if isinstance(existing, dict):
        existing = [existing]
    ids, texts = [], []
    for e in existing if isinstance(existing, list) else []:
        if not isinstance(e, dict):
            continue
        for k in ('ques_Id', 'quesId', 'QuesId', 'dup_QuesId', 'id', 'Id'):
            if e.get(k) is not None:
                ids.append(str(e[k]))
                break
        raw = e.get('ques_Content') or e.get('quesContent') or e.get('decryptedQuesContent') or e.get('text') or ''
        texts.append(TAG_RE.sub(' ', str(raw)).strip()[:80])
    return ids, texts


def post_question(session, payload, dry_run, test_mode):
    if dry_run:
        print(f"[DRY RUN] Would POST: subj_Id={payload['subj_Id']} subj_Code={payload['subj_Code']} parent={payload['parntQus_Id']} content={payload['ques_Content'][:60]!r}")
        return True, None, 'dry-run'
    headers = {'accept': '*/*', 'Authorization': f"Bearer {QUES_TOKEN}", 'Content-Type': 'application/json', 'User-Agent': USER_AGENT}
    payload_kb = len(json.dumps(payload).encode('utf-8')) / 1024
    if test_mode or payload_kb > 500:
        print(f"  payload size: {payload_kb:.0f} KB")
    for attempt in range(3):
        try:
            resp = session.post(API_URL, headers=headers, json=payload, timeout=SAVE_TIMEOUT)
        except requests.RequestException as e:
            if attempt == 2:
                return False, None, f"request error (payload {payload_kb:.0f} KB): {e}"
            time.sleep(5 * (attempt + 1))
            continue
        if resp.status_code in (200, 201):
            try:
                resp_json = resp.json()
            except ValueError:
                resp_json = None
            if test_mode:
                print(f"  -> status {resp.status_code}, raw response: {resp.text[:500]}")
            return True, extract_new_id(resp_json), 'ok'
        try:
            body = resp.json()
        except ValueError:
            body = None
        if isinstance(body, dict):
            message = str(body.get('message', '')).lower()
            if 'duplicate' in message or 'similar questions already exist' in message:
                pct = body.get('similarityPercentage')
                ids, texts = _dup_details(body)
                if test_mode:
                    print(f"  -> FULL DUP RESPONSE: {resp.text[:2000]}")
                match = ', matches ques_Id ' + '/'.join(ids) if ids else ''
                return False, None, f"duplicate ({pct}% similar{match}): {str(body.get('message', ''))[:120]}"
        if resp.status_code in (401, 403):
            return False, None, f"auth error {resp.status_code}: {resp.text[:200]}"
        if resp.status_code == 400:
            return False, None, f"http 400: {resp.text[:1500]}"
        if resp.status_code == 405:
            chain = ' -> '.join(h.url for h in resp.history) + (' -> ' + resp.url if resp.history else '')
            allow = resp.headers.get('Allow', '<not sent>')
            detail = f"http 405 | final_url={resp.url} | redirects={len(resp.history)} ({chain if resp.history else 'none'}) | Allow_header={allow} | body={resp.text[:150]}"
            return False, None, detail
        if resp.status_code >= 500 and attempt < 2:
            time.sleep(5 * (attempt + 1))
            continue
        return False, None, f"http {resp.status_code}: {resp.text[:200]}"
    return False, None, 'failed after retries'


UNMAPPED_CODES = set()


class SubjectNotFoundError(Exception):
    pass


class MissingTokenError(Exception):
    pass


def _count_descendants(node):
    kids = list(node.get('SubDivisions', [])) + list(node.get('SubQuestions', []))
    return len(kids) + sum(_count_descendants(k) for k in kids)


def process_node(node, subj_code, qno_path, parent_id, disp_order, session, stats, dry_run, test_mode, delay):
    description = node.get('Description')
    current_id = parent_id
    if description is not None:
        if IMAGE_MARKER_RE.search(description):
            stats['error'] += 1
            print(f"  FAILED {qno_path}: image was not uploaded (placeholder still in text). Re-run doc_to_json.py for this file; re-running this script is safe.")
            return
        content = format_ques_content(description)
      
        payload = build_payload(subj_code, content, parent_id, disp_order)
        ok, new_id, status = post_question(session, payload, dry_run, test_mode)
        is_duplicate = not ok and status.startswith('duplicate')
        is_subject_missing = not ok and not is_duplicate and 'subj_id' in status.lower() and 'null' in status.lower()
        if not dry_run:
            time.sleep(delay)
        if ok:
            stats['saved'] += 1
            if new_id is not None:
                current_id = new_id
        elif is_subject_missing:
            raise SubjectNotFoundError(subj_code)
        elif is_duplicate:
            stats['duplicate'] += 1
            skipped = _count_descendants(node)
            stats['children skipped (parent duplicate)'] += skipped
            note = f" (+{skipped} sub-question(s) not sent)" if skipped else ''
            print(f"  DUPLICATE {qno_path}: {status}{note}")
            return
        else:
            stats['error'] += 1
            print(f"  FAILED {qno_path}: {status}")
    children = list(node.get('SubDivisions', [])) + list(node.get('SubQuestions', []))
    for i, child in enumerate(children):
        process_node(child, subj_code, f"{qno_path}.{child.get('Qno')}", current_id, i + 1, session, stats, dry_run, test_mode, delay)


def process_file(path, session, stats, dry_run, test_mode, delay, bad_subj_codes):
    with open(path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    sync_image_urls(data)
    subj_code = data.get('subj_code', '')
    questions = data.get('questions', [])
    if subj_code not in SUBJ_ID_MAP:
        UNMAPPED_CODES.add(subj_code)
    print(f"Processing {path.name} ({len(questions)} top-level questions)")
    if subj_code in bad_subj_codes:
        print(f"  SKIPPING (subj_code {subj_code!r} not found earlier in this run) -- {path.name}")
        return
    for i, q in enumerate(questions):
        try:
            process_node(q, subj_code, str(q.get('Qno')), None, i + 1, session, stats, dry_run, test_mode, delay)
        except SubjectNotFoundError:
            print(f"  SKIPPING REST OF FILE: subj_code {subj_code!r} not found on server (NULL Subj_Id) -- {path.name}")
            bad_subj_codes.add(subj_code)
            stats['files skipped (subject not found)'] += 1
            break


def run_upload(folder, dry_run=False, test_mode=False, limit=None, delay=.6):
    if not dry_run and not QUES_TOKEN:
        raise MissingTokenError('ques_TOKEN not found. Put it in a .env file as ques_TOKEN=...')
    folder = Path(folder)
    files = sorted(folder.glob('*.json'))
    if not files:
        raise FileNotFoundError(f"No .json files found in {folder}")
    if test_mode:
        files = files[:1]
        print(f"TEST MODE: processing only {files[0].name}")
    elif limit:
        files = files[:limit]
    session = requests.Session()
    if not dry_run:
        SUBJ_ID_MAP.clear()
        SUBJ_ID_MAP.update(fetch_subject_map(session))
        SUBJ_ID_MAP.update(ENV_SUBJ_IDS)
    stats = Counter()
    bad_subj_codes = set()
    for path in files:
        try:
            process_file(path, session, stats, dry_run, test_mode, delay, bad_subj_codes)
        except Exception as e:
            stats['files with errors'] += 1
            print(f"[ERROR] processing {path.name}: {e}")
    print(f"\nDone. {len(files)} file(s). " + ', '.join(f"{k}: {v}" for k, v in sorted(stats.items())))
    if UNMAPPED_CODES:
        shown = ', '.join(sorted(UNMAPPED_CODES)[:30]) + (' ...' if len(UNMAPPED_CODES) > 30 else '')
        print(f"\n{len(UNMAPPED_CODES)} subject code(s) had no numeric id and were sent as id 0 + code "
              f"(the server's duplicate check may not see them): {shown}")
    return dict(stats)


def main():
    parser = argparse.ArgumentParser(description='Upload question-bank JSON files to SaveQuestion API')
    parser.add_argument('--folder', default=DEFAULT_FOLDER, help=f"Folder containing the .json files (default: {DEFAULT_FOLDER})")
    parser.add_argument('--dry-run', action='store_true', help="Don't call the API, just print payloads")
    parser.add_argument('--test', action='store_true', help='Process only the first file and print raw responses')
    parser.add_argument('--limit', type=int, default=None, help='Only process this many files')
    parser.add_argument('--delay', type=float, default=.6, help='Seconds to sleep between real API calls')
    args = parser.parse_args()
    try:
        run_upload(args.folder, dry_run=args.dry_run, test_mode=args.test, limit=args.limit, delay=args.delay)
    except MissingTokenError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)
    except FileNotFoundError as e:
        print(str(e), file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()