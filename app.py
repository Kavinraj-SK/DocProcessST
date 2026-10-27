import queue
import sys
import tempfile
import threading
import traceback
import uuid
from pathlib import Path

from flask import Flask, Response, jsonify, render_template, request

import doc_to_json as pipeline
import json_to_db as db_pipeline

app = Flask(__name__)

# The real, on-disk folder doc_to_json.py itself is configured to write
# into -- results land here directly, same as running the CLI, instead
# of behind a temp-folder + zip-download round trip.
RESULT_PATH = Path(pipeline.OUTPUT_PATH).expanduser().resolve()


@app.errorhandler(Exception)
def _json_on_any_error(e):
    """
    Belt-and-suspenders: whatever goes wrong, anywhere, always send JSON
    back instead of Flask's default HTML error page. The frontend's
    resp.json() call has no way to parse HTML, so any endpoint that can
    return non-JSON on failure turns into a silent hang in the browser.
    The real traceback still goes to this terminal so it's not lost.
    """
    traceback.print_exc()
    from werkzeug.exceptions import HTTPException
    status = e.code if isinstance(e, HTTPException) else 500
    return jsonify({"error": str(e) or e.__class__.__name__}), status

JOBS_ROOT = Path(tempfile.gettempdir()) / "doc_to_json_webapp"
JOBS_ROOT.mkdir(exist_ok=True)

# job_id -> {"queue": Queue, "status": "running"|"done"|"error",
#            "summary": dict|None, "error": str|None, "kind": "convert"|"refresh"}
# Note: there's no per-job "output_dir" any more -- results are written
# straight into RESULT_PATH (shared, real, on-disk), not somewhere
# scoped to the job.
_jobs = {}
_jobs_lock = threading.Lock()


class _ThreadTee:
    def __init__(self, real_stdout):
        self._real = real_stdout
        self._local = threading.local()

    def register(self, q: "queue.Queue"):
        self._local.queue = q

    def unregister(self):
        self._local.queue = None

    def write(self, text):
        self._real.write(text)
        q = getattr(self._local, "queue", None)
        if q is not None and text.strip():
            for line in text.splitlines():
                if line.strip():
                    q.put(line)

    def flush(self):
        self._real.flush()

    def isatty(self):
        return False


_tee = _ThreadTee(sys.stdout)
sys.stdout = _tee


def _new_job(kind: str) -> str:
    job_id = uuid.uuid4().hex[:12]  # shorter than the full 32 chars -- leaves
    # more of Windows' 260-char MAX_PATH budget for the uploaded folder's own
    # (sometimes long) relative paths and filenames
    job_dir = JOBS_ROOT / job_id
    if kind in ("convert", "upload_db"):
        (job_dir / "input").mkdir(parents=True)
    with _jobs_lock:
        _jobs[job_id] = {
            "queue": queue.Queue(),
            "status": "running",
            "summary": None,
            "error": None,
            "kind": kind,
        }
    return job_id


def _safe_relpath(raw_path: str) -> Path:
    """
    Turns a browser-supplied relative path (from webkitRelativePath) into
    a safe path with no ".." segments or absolute roots, so an uploaded
    folder can never write outside its own job directory.
    """
    parts = [p for p in raw_path.replace("\\", "/").split("/") if p not in ("", ".", "..")]
    if not parts:
        parts = ["file"]
    return Path(*parts)


def _save_uploaded_folder(job_dir: Path) -> int:
    """Saves every uploaded file under job_dir/input, preserving the
    folder structure the browser sent via webkitRelativePath. Returns
    how many files were saved."""
    count = 0
    for f in request.files.getlist("files"):
        if not f.filename:
            continue
        rel = _safe_relpath(f.filename)
        dest = job_dir / "input" / rel
        dest.parent.mkdir(parents=True, exist_ok=True)
        f.save(dest)
        count += 1
    return count


def _run_convert_job(job_id: str):
    job = _jobs[job_id]
    _tee.register(job["queue"])
    try:
        input_dir = JOBS_ROOT / job_id / "input"
        print(f"Writing results to {RESULT_PATH}")
        summary = pipeline.run_pipeline(input_dir, RESULT_PATH)
        job["summary"] = summary
        job["status"] = "done"
    except Exception as e:
        print(f"[ERROR] {e}")
        job["error"] = str(e)
        job["status"] = "error"
    finally:
        _tee.unregister()
        job["queue"].put(None)  # sentinel: stream is over


def _run_refresh_job(job_id: str):
    job = _jobs[job_id]
    _tee.register(job["queue"])
    try:
        print(f"Refreshing image URLs in {RESULT_PATH}")
        json_files = sorted(RESULT_PATH.glob("*.json"))
        if not json_files:
            raise FileNotFoundError(f"No .json files found in {RESULT_PATH}.")
        print(f"Refreshing image URLs in {len(json_files)} file(s)...")
        # refresh_json_file() rewrites each file in place -- this IS the
        # "current path inside doc_to_json.py" (RESULT_PATH), so nothing
        # needs to be copied anywhere first.
        ok = sum(1 for jf in json_files if pipeline.refresh_json_file(jf))
        print(f"\nDone. {ok}/{len(json_files)} file(s) refreshed.")
        job["summary"] = {"total": len(json_files), "refreshed": ok}
        job["status"] = "done"
    except Exception as e:
        print(f"[ERROR] {e}")
        job["error"] = str(e)
        job["status"] = "error"
    finally:
        _tee.unregister()
        job["queue"].put(None)


def _run_upload_db_job(job_id: str, dry_run: bool):
    job = _jobs[job_id]
    _tee.register(job["queue"])
    try:
        input_dir = JOBS_ROOT / job_id / "input"
        log_path = JOBS_ROOT / job_id / "upload_log.csv"
        # bad_subj_codes.json is shared across every upload job (not put in
        # the per-job folder), so a subj_code found bad in one run is
        # skipped instantly in the next, same as the CLI's behaviour when
        # you keep re-running it against the same output folder.
        shared_bad_subj_cache = str(JOBS_ROOT / "bad_subj_codes.json")
        print(f"{'DRY RUN -- ' if dry_run else ''}Uploading question JSON from {input_dir}")
        summary = db_pipeline.run_upload(
            input_dir,
            log=str(log_path),
            dry_run=dry_run,
            bad_subj_cache_path=shared_bad_subj_cache,
        )
        job["summary"] = summary
        job["status"] = "done"
    except db_pipeline.MissingTokenError as e:
        print(f"[ERROR] {e}")
        job["error"] = str(e)
        job["status"] = "error"
    except Exception as e:
        print(f"[ERROR] {e}")
        job["error"] = str(e)
        job["status"] = "error"
    finally:
        _tee.unregister()
        job["queue"].put(None)


@app.route("/")
def index():
    return render_template(
        "index.html",
        bearer_configured=bool(pipeline.BEARER_TOKEN),
        ques_token_configured=bool(db_pipeline.QUES_TOKEN),
        result_path=str(RESULT_PATH),
    )


@app.route("/convert", methods=["POST"])
def convert():
    job_id = _new_job("convert")
    job_dir = JOBS_ROOT / job_id
    try:
        n = _save_uploaded_folder(job_dir)
    except OSError as e:
        with _jobs_lock:
            del _jobs[job_id]
        # Common cause: Windows MAX_PATH (260 chars) exceeded once the temp
        # dir + job id + relative folder + filename are all concatenated.
        return jsonify({"error": f"Could not save an uploaded file ({e}). "
                                  f"If this is a long/nested filename, try "
                                  f"shortening it or uploading a shallower folder."}), 500
    if n == 0:
        with _jobs_lock:
            del _jobs[job_id]
        return jsonify({"error": "No files were uploaded."}), 400
    threading.Thread(target=_run_convert_job, args=(job_id,), daemon=True).start()
    return jsonify({"job_id": job_id, "files_received": n})


@app.route("/refresh", methods=["POST"])
def refresh():
    job_id = _new_job("refresh")
    threading.Thread(target=_run_refresh_job, args=(job_id,), daemon=True).start()
    return jsonify({"job_id": job_id})


@app.route("/upload_db", methods=["POST"])
def upload_db():
    job_id = _new_job("upload_db")
    job_dir = JOBS_ROOT / job_id
    try:
        n = _save_uploaded_folder(job_dir)
    except OSError as e:
        with _jobs_lock:
            del _jobs[job_id]
        return jsonify({"error": f"Could not save an uploaded file ({e})."}), 500
    if n == 0:
        with _jobs_lock:
            del _jobs[job_id]
        return jsonify({"error": "No files were uploaded."}), 400
    # Reject anything that isn't .json up front, so a wrong drag-and-drop
    # (e.g. the .docx folder) fails fast with a clear reason instead of
    # silently finding "no .json files" deep inside run_upload().
    input_dir = job_dir / "input"
    non_json = [str(p.relative_to(input_dir)) for p in input_dir.rglob("*")
                if p.is_file() and p.suffix.lower() != ".json"]
    if non_json:
        with _jobs_lock:
            del _jobs[job_id]
        sample = ", ".join(non_json[:5])
        return jsonify({"error": f"Only .json files are accepted here ({sample}...)."}), 400
    # run_upload() only looks at the immediate top level of the folder it's
    # given (matching the CLI's behaviour), so if a whole folder got
    # dragged in rather than loose files, flatten everything up to
    # input_dir's root first.
    for p in list(input_dir.rglob("*.json")):
        if p.parent != input_dir:
            dest = input_dir / p.name
            i = 1
            while dest.exists():
                dest = input_dir / f"{p.stem}_{i}{p.suffix}"
                i += 1
            p.rename(dest)
    dry_run = request.form.get("dry_run") == "0"
    threading.Thread(target=_run_upload_db_job, args=(job_id, dry_run), daemon=True).start()
    return jsonify({"job_id": job_id, "files_received": n})


@app.route("/stream/<job_id>")
def stream(job_id):
    job = _jobs.get(job_id)
    if job is None:
        return "unknown job", 404

    def gen():
        q = job["queue"]
        while True:
            line = q.get()
            if line is None: 
                payload = {
                    "status": job["status"],
                    "summary": job["summary"],
                    "error": job["error"],
                }
                yield f"event: done\ndata: {_json_dumps(payload)}\n\n"
                return
            yield f"data: {_json_dumps({'line': line})}\n\n"

    return Response(gen(), mimetype="text/event-stream", headers={
        "Cache-Control": "no-cache",
        "X-Accel-Buffering": "no",
    })


def _json_dumps(obj):
    import json
    return json.dumps(obj)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, threaded=True, debug=False)