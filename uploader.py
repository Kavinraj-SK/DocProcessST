import argparse, json, re, sys, time, base64
from pathlib import Path
import os, requests
from dotenv import load_dotenv

load_dotenv()

JSON_FOLDER = 'C:\\Users\\Administrator\\Downloads\\23PT17\\Ques(csv)\\csv'
PROCESSED_LOG_FILENAME = 'mcq_processed_log.json'
API_BASE_URL = os.getenv('API_BASE_URL', 'https://campus.psgtech.ac.in/Dote/api/QuestionBank/SaveQuestion')
AUTH_TOKEN = os.getenv('AUTH_TOKEN')
EXAM_ID_DEFAULT = 1

# Q_MARKS attrval_code mapping (attr_id 82): row MARK -> code sent in the payload
#   1 = 1 MARK, 2 = 2 MARKS, 3 = 5 MARKS, 4 = 10 MARKS, 5 = 12 MARKS,
#   6 = 15 MARKS, 7 = 100 MARKS, 8 = 3 MARKS, 9 = 50 MARKS
MARK_TO_ATTRVAL_CODE = {1: 1, 2: 2, 3: 8}


def check_token_expiry(token: str):
    try:
        payload_b64 = token.split('.')[1]
        padding = '=' * (-len(payload_b64) % 4)
        payload = json.loads(base64.urlsafe_b64decode(payload_b64 + padding))
        exp = payload.get('exp')
        if exp is None:
            return
        remaining = exp - time.time()
        if remaining <= 0:
            print('WARNING: AUTH_TOKEN appears to already be expired. Requests will likely fail with 401.\n')
        elif remaining < 3600:
            print(f"WARNING: AUTH_TOKEN expires in about {int(remaining / 60)} minute(s).\n")
    except Exception:
        pass


CANONICAL_KEY_ALIASES = {
    'SUBCODE': ['SUBCODE', 'Subcode', 'SubCode', 'Subject Code'],
    'MARK': ['MARK', 'Mark', 'Marks'],
    'QUESTION': ['QUESTION', 'Question', 'Ques'],
    'ANSWER': ['ANSWER', 'Answer'],
    'BLOOMS_LEVEL': ['BLOOMS_LEVEL', 'Blooms Level', 'BloomsLevel', 'Bloom Level'],
    'OPTION1': ['OPTION1', 'Option 1', 'Option1'],
    'OPTION2': ['OPTION2', 'Option 2', 'Option2'],
    'OPTION3': ['OPTION3', 'Option 3', 'Option3'],
    'OPTION4': ['OPTION4', 'Option 4', 'Option4'],
}


def normalize_row(row: dict) -> dict:
    normalized = dict(row)
    for canonical, aliases in CANONICAL_KEY_ALIASES.items():
        if normalized.get(canonical) not in (None, ''):
            continue
        for alias in aliases:
            if row.get(alias) not in (None, ''):
                normalized[canonical] = row[alias]
                break
    return normalized


def resolve_correct_index(row: dict) -> int:
    raw = row.get('ANSWER')
    if raw is None or str(raw).strip() == '':
        raise ValueError('missing ANSWER')
    raw_str = str(raw).strip()
    if raw_str.isdigit():
        n = int(raw_str)
        if 1 <= n <= 4:
            return n
        raise ValueError(f"ANSWER {raw_str!r} is a number but not in range 1-4")

    def norm(s):
        return ' '.join(str(s).strip().casefold().split())

    target = norm(raw_str)
    matches = []
    for i in range(1, 5):
        option_text = row.get(f"OPTION{i}")
        if option_text is not None and norm(option_text) == target:
            matches.append(i)
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        raise ValueError(f"ANSWER text {raw_str!r} matches more than one option ({matches}) -- ambiguous, skipping")
    raise ValueError(f"ANSWER {raw_str!r} is not a number 1-4 and doesn't match any OPTION1-4 text -- can't resolve correct answer")


def make_answer(content: str, is_correct: bool, display_order: int) -> dict:
    return {
        'ans_Content': content,
        'is_Correct': is_correct,
        'display_Order': display_order,
        'hash_Ans': '',
        'enc_Text_Ans': '',
        'enc_Key_Ans': '',
        'nonce_Ans': '',
        'enc_Tag_Ans': '',
        'stepMarks': [],
    }


def build_answers(row: dict, correct_index: int) -> list:
    answers = []
    for i in range(1, 5):
        option_text = row.get(f"OPTION{i}")
        if option_text is None or option_text == '':
            continue
        answers.append(make_answer(str(option_text).strip(), i == correct_index, i))
    return answers


def build_fill_blank_answers(correct_text: str) -> list:
    """Fill in the blank: only the question and the single correct answer, no other slots."""
    return [make_answer(correct_text, True, 1)]


def capitalize_first(text) -> str:
    if text is None:
        return text
    s = str(text).strip()
    if not s:
        return s
    return s[0].upper() + s[1:]


def parse_blooms_level(raw) -> int:
    if raw is None or str(raw).strip() == '':
        return 0
    s = str(raw).strip()
    m = re.match(r'^[Ll]\s*(\d+)$', s)
    if m:
        return int(m.group(1))
    if s.isdigit():
        return int(s)
    return 0


def has_options(row: dict) -> bool:
    return any(row.get(f"OPTION{i}") not in (None, '') for i in range(1, 5))


QUES_TYPE_MCQ = 3
QUES_TYPE_FILL_IN_BLANK = 3
QUES_TYPE_SHORT_ANSWER = 1
QUES_TYPE_DESCRIPTIVE = 2


def determine_type_and_answers(row: dict, mark_val: int, row_num: int):
    if mark_val == 1:
        if has_options(row):
            correct_index = resolve_correct_index(row)
            return QUES_TYPE_MCQ, build_answers(row, correct_index)
        print(f"  [row {row_num}] FILL IN THE BLANK detected (MARK=1, no options filled)")
        raw_answer = row.get('ANSWER')
        if raw_answer is None or str(raw_answer).strip() == '':
            raise ValueError('missing ANSWER (fill in the blank, MARK=1)')
        return QUES_TYPE_FILL_IN_BLANK, build_fill_blank_answers(str(raw_answer).strip())
    elif mark_val == 2:
        return QUES_TYPE_SHORT_ANSWER, []
    elif mark_val == 3:
        return QUES_TYPE_DESCRIPTIVE, []
    raise ValueError(f"unsupported MARK value: {mark_val!r}")


def build_payload(row: dict, quesType_Id: int, answers: list, mark_val: int) -> dict:
    return {
        'ques_Id': 0,
        'qbConfig_ID': 0,
        'exam_ID': EXAM_ID_DEFAULT,
        'subj_Id': 0,
        'subj_Code': str(row.get('SUBCODE')),
        'unit_Id': 0,
        'co_Id': 0,
        'topic_Id': 0,
        'mark': MARK_TO_ATTRVAL_CODE.get(mark_val, mark_val),
        'parntQus_Id': None,
        'ques_Lvl': 0,
        'disp_Order': int(row['QUESTION_ID']) if row.get('QUESTION_ID') not in (None, '') else 0,
        'quesType_Id': quesType_Id,
        'bloomLvl_Id': parse_blooms_level(row.get('BLOOMS_LEVEL')),
        'diffLvl_Id': 0,
        'ques_Content': f"<p>{capitalize_first(row.get('QUESTION'))}</p>",
        'hash_Ques': '',
        'enc_Text_Ques': '',
        'enc_Key_Ques': '',
        'nonce_Ques': '',
        'enc_Tag_Ques': '',
        'orig_Enc_Text': '',
        'orig_Enc_Key': '',
        'orig_Nonce': '',
        'orig_Enc_Tag': '',
        'answers': answers,
        'skippedDuplicates': [],
    }


def extract_duplicate_ids(body: dict) -> list:
    candidates = body.get('duplicateQuestions') or body.get('existingQuestions') or []
    if isinstance(candidates, dict):
        candidates = [candidates]
    ids = []
    for entry in candidates:
        if isinstance(entry, dict):
            for key in ('existingQuesId', 'Ques_Id', 'ques_Id', 'id', 'Id', 'quesId'):
                if key in entry and entry[key] is not None:
                    ids.append(entry[key])
                    break
        elif isinstance(entry, (int, str)):
            ids.append(entry)
    return ids


class DuplicateQuestionError(RuntimeError):
    pass


def save_question(session: requests.Session, payload: dict):
    response = session.post(API_BASE_URL, json=payload, timeout=30)
    if response.status_code == 401:
        raise RuntimeError('401 Unauthorized -- AUTH_TOKEN is missing, invalid, or expired.')
    if not response.ok:
        detail = f"{response.status_code} {response.reason}: {response.text[:300]}"
        try:
            body = response.json()
        except ValueError:
            body = None
        if isinstance(body, dict):
            message = str(body.get('message', '')).lower()
            if 'duplicate' in message and 'similar' not in message:
                raise DuplicateQuestionError(body.get('message', detail))
            if 'similar questions already exist' in message:
                dup_ids = extract_duplicate_ids(body)
                pct = body.get('similarityPercentage')
                print(f"    -> near-duplicate ({pct}% similar to existing Ques_Id {dup_ids or '?'}); treating as already saved, skipping")
                raise DuplicateQuestionError(body.get('message', detail))
        allow_header = response.headers.get('Allow')
        if allow_header:
            detail += f" | server says Allow: {allow_header}"
        raise RuntimeError(detail)
    return response


def extract_subcode_from_filename(filename: str):
    m = re.search(r'\d{8,}', filename)
    return m.group(0) if m else None


def run_one_file(session: requests.Session, path: Path, dry_run: bool, dry_run_limit: int, limit_rows: int = None):
    with open(path, 'r', encoding='utf-8-sig') as f:
        rows = json.load(f)
    if not isinstance(rows, list):
        raise RuntimeError(f"Expected a JSON list of rows in {path.name}, got {type(rows).__name__}")
    non_dict_count = sum(1 for r in rows if not isinstance(r, dict))
    if non_dict_count:
        print(f"  -> WARNING: {non_dict_count} of {len(rows)} entries in this file are not row-objects (see per-row detail below)")
    print(f"  -> rows in file: {len(rows)}")
    if limit_rows is not None and limit_rows < len(rows):
        rows = rows[:limit_rows]
        print(f"  -> --limit-rows active: only processing the first {limit_rows} row(s) of this file")
    filename_subcode = extract_subcode_from_filename(path.name)
    if filename_subcode:
        print(f"  -> fallback SUBCODE from filename (used only for rows missing SUBCODE): {filename_subcode}")
    ok_count = 0
    duplicate_count = 0
    file_errors = []
    for i, row in enumerate(rows, start=1):
        if not isinstance(row, dict):
            file_errors.append((i, f"row is not an object, it's a {type(row).__name__}"))
            print(f"  [row {i}] FAILED: row is not an object, it's a {type(row).__name__}")
            print(f"    raw value: {row!r}")
            continue
        row = normalize_row(row)
        if not row.get('SUBCODE'):
            if filename_subcode:
                row = {**row, 'SUBCODE': filename_subcode}
                print(f"  [row {i}] SUBCODE missing -- using code from filename: {filename_subcode}")
            else:
                file_errors.append((i, 'missing SUBCODE (and none could be extracted from the filename)'))
                print(f"  [row {i}] FAILED: missing SUBCODE (and none could be extracted from the filename)")
                continue
        try:
            mark_val = int(row['MARK']) if row.get('MARK') not in (None, '') else 0
        except (KeyError, ValueError):
            file_errors.append((i, f"bad MARK value: {row.get('MARK')!r}"))
            print(f"  [row {i}] FAILED: bad MARK value: {row.get('MARK')!r}")
            continue
        try:
            quesType_Id, answers = determine_type_and_answers(row, mark_val, i)
        except ValueError as exc:
            file_errors.append((i, str(exc)))
            print(f"  [row {i}] FAILED: {exc}")
            continue
        try:
            payload = build_payload(row, quesType_Id, answers, mark_val)
        except (KeyError, ValueError) as exc:
            file_errors.append((i, f"bad row data: {exc}"))
            print(f"  [row {i}] FAILED: bad row data: {exc}")
            continue
        if dry_run:
            if i <= dry_run_limit:
                print(f"  [row {i}] DRY RUN payload:")
                print(json.dumps(payload, indent=2, ensure_ascii=False))
            continue
        try:
            save_question(session, payload)
            ok_count += 1
        except DuplicateQuestionError:
            duplicate_count += 1
            print(f"  [row {i}] ALREADY EXISTS (duplicate, skipped)")
        except Exception as exc:
            file_errors.append((i, str(exc)))
            print(f"  [row {i}] FAILED: {exc}")
    if dry_run:
        print(f"  -> dry run: showed {min(dry_run_limit, len(rows))} payload(s), sent nothing")
        return 0, 0, []
    print(f"  -> saved {ok_count} new, {duplicate_count} already existed, {len(file_errors)} genuinely failed (of {len(rows)} total)")
    if file_errors and ok_count == 0 and duplicate_count == 0:
        raise RuntimeError(f"All {len(file_errors)} row(s) in this file failed to save.")
    return ok_count, duplicate_count, file_errors


def load_processed_log(log_path: Path) -> dict:
    if not log_path.exists():
        return {}
    try:
        with open(log_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (ValueError, OSError):
        return {}


def save_processed_log(log_path: Path, log: dict):
    with open(log_path, 'w', encoding='utf-8') as f:
        json.dump(log, f, indent=2, ensure_ascii=False)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--dry-run', action='store_true',
                        help='Build and print payloads without sending anything to the API.')
    parser.add_argument('--dry-run-limit', type=int, default=2,
                        help='How many payloads per file to print in dry-run mode (default 2).')
    parser.add_argument('--limit-rows', type=int, default=None,
                        help='Only process this many rows total, from the first file(s) processed. '
                             'Useful for a small live test before committing to a full run. '
                             'Rows/files are still marked processed in the log as normal.')
    args = parser.parse_args()

    if not AUTH_TOKEN:
        print("AUTH_TOKEN not set. Put it in .env as AUTH_TOKEN=<bearer token, no 'Bearer ' prefix>.")
        sys.exit(1)
    check_token_expiry(AUTH_TOKEN)

    folder = Path(JSON_FOLDER)
    log_path = folder / PROCESSED_LOG_FILENAME
    processed_log = load_processed_log(log_path)
    json_files = sorted(p for p in folder.rglob('*.json') if p.name != PROCESSED_LOG_FILENAME)
    if not json_files:
        print(f"No .json files found in {folder}")
        sys.exit(1)
    print(f"Found {len(json_files)} JSON file(s) in {folder} (including subfolders)\n")

    session = requests.Session()
    session.headers.update({
        'Accept': '*/*',
        'Authorization': f"Bearer {AUTH_TOKEN}",
        'Content-Type': 'application/json',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
    })

    succeeded = []
    already_up_to_date = []
    partially_failed = []
    failed = []
    skipped_already_processed = []
    total_saved = 0
    total_duplicates = 0
    rows_processed_so_far = 0

    if args.limit_rows is not None:
        print(f"--limit-rows {args.limit_rows} active: will stop after {args.limit_rows} row(s) total have been attempted, across file(s), then exit early.\n")

    for path in json_files:
        rel_key = str(path.relative_to(folder))
        if not args.dry_run and rel_key in processed_log:
            print(f"Skipping {rel_key} -- already processed (logged {processed_log[rel_key]})\n")
            skipped_already_processed.append(rel_key)
            continue
        if args.limit_rows is not None and rows_processed_so_far >= args.limit_rows:
            print(f"Reached --limit-rows {args.limit_rows}. Stopping before {rel_key}.\n")
            break
        file_row_limit = None
        if args.limit_rows is not None:
            file_row_limit = args.limit_rows - rows_processed_so_far
        print(f"Processing {rel_key} ...")
        try:
            ok_count, duplicate_count, file_errors = run_one_file(
                session, path, args.dry_run, args.dry_run_limit, limit_rows=file_row_limit)
            if args.limit_rows is not None:
                rows_processed_so_far += ok_count + duplicate_count + len(file_errors)
            total_saved += ok_count
            total_duplicates += duplicate_count
            if args.dry_run:
                print()
                continue
            if file_errors:
                partially_failed.append((rel_key, len(file_errors)))
                print(f"  PARTIAL: {rel_key} ({len(file_errors)} row(s) genuinely failed)\n")
            elif ok_count == 0 and duplicate_count > 0:
                already_up_to_date.append(rel_key)
                print(f"  ALREADY UP TO DATE: {rel_key} ({duplicate_count} row(s) already saved)\n")
                if args.limit_rows is None:
                    processed_log[rel_key] = time.strftime('%Y-%m-%d %H:%M:%S')
                    save_processed_log(log_path, processed_log)
                else:
                    print("    (NOT marking as processed in log -- this was a --limit-rows test run)\n")
            else:
                succeeded.append(rel_key)
                print(f"  OK: {rel_key}\n")
                if args.limit_rows is None:
                    processed_log[rel_key] = time.strftime('%Y-%m-%d %H:%M:%S')
                    save_processed_log(log_path, processed_log)
                else:
                    print("    (NOT marking as processed in log -- this was a --limit-rows test run)\n")
        except Exception as exc:
            failed.append((rel_key, str(exc)))
            print(f"  FAILED: {rel_key} -> {exc}\n")

    if args.dry_run:
        print('Dry run complete -- nothing was sent to the API.')
        return

    print('=' * 60)
    print(f"Done. {len(succeeded)} file(s) with new questions saved, {len(already_up_to_date)} already up to date, "
          f"{len(skipped_already_processed)} skipped (logged from a prior run), {len(partially_failed)} partially failed, "
          f"{len(failed)} fully failed.\n{total_saved} new question(s) saved, {total_duplicates} already existed.")
    if partially_failed:
        print('\nFiles with some rows genuinely failing:')
        for name, count in partially_failed:
            print(f"  - {name}: {count} row(s) failed")
    if failed:
        print('\nFully failed files:')
        for name, err in failed:
            print(f"  - {name}: {err}")

if __name__ == '__main__':
    main()