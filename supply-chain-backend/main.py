import os, uuid, json, asyncio
from pathlib import Path
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, StreamingResponse
import aiofiles

from db import get_pool, close_pool

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173"],
    allow_methods=["*"],
    allow_headers=["*"],
)

JOBS_DIR = Path("jobs")
JOBS_DIR.mkdir(exist_ok=True)

# In-memory event buffers for active SSE streams only — NOT the source of truth.
# Postgres ("jobs" table) is the source of truth and is what survives restarts.
LIVE_EVENTS = {}   # job_id -> list[event]
LIVE_STATUS = {}   # job_id -> 'queued' | 'running' | 'done' | 'error'


@app.on_event("shutdown")
async def shutdown():
    await close_pool()


async def new_job(kind, month_label=None):
    job_id = str(uuid.uuid4())[:8]
    pool = await get_pool()
    await pool.execute(
        "INSERT INTO jobs (job_id, kind, status, month_label) VALUES ($1, $2, 'queued', $3)",
        job_id, kind, month_label,
    )
    LIVE_EVENTS[job_id] = []
    LIVE_STATUS[job_id] = "queued"
    return job_id


def push_event(job_id, event):
    LIVE_EVENTS[job_id].append(event)


async def finalize_job(job_id, status, output_path=None, summary=None):
    LIVE_STATUS[job_id] = status
    pool = await get_pool()
    await pool.execute(
        """UPDATE jobs SET status = $1, output_path = $2, summary = $3,
           events = $4::jsonb WHERE job_id = $5""",
        status, output_path,
        json.dumps(summary) if summary else None,
        json.dumps(LIVE_EVENTS.get(job_id, [])),
        job_id,
    )


async def event_stream(job_id):
    sent = 0
    while True:
        events = LIVE_EVENTS.get(job_id, [])
        while sent < len(events):
            yield f"data: {json.dumps(events[sent])}\n\n"
            sent += 1
        if LIVE_STATUS.get(job_id) in ("done", "error"):
            break
        await asyncio.sleep(0.3)


@app.get("/api/health")
async def health():
    return {"status": "ok"}


@app.post("/api/run")
async def run_pipeline(
    forecast:      UploadFile = File(...),
    bom:           UploadFile = File(...),
    component_soh: UploadFile = File(...),
    prices:        UploadFile = File(...),
    vendor:        UploadFile = File(...),
    month_label:   str        = Form(...),
):
    job_id  = await new_job("run", month_label)
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir()

    files = {
        "forecast":      forecast,
        "bom":           bom,
        "component_soh": component_soh,
        "prices":        prices,
        "vendor":        vendor,
    }
    saved = {}
    for key, upload in files.items():
        dest = job_dir / upload.filename
        async with aiofiles.open(dest, "wb") as f:
            await f.write(await upload.read())
        saved[key] = str(dest)

    asyncio.create_task(_run_job(job_id, saved, month_label))
    return {"job_id": job_id}


@app.get("/api/status/{job_id}")
async def stream_status(job_id: str):
    if job_id not in LIVE_STATUS:
        raise HTTPException(404, "Job not found")
    return StreamingResponse(
        event_stream(job_id),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.get("/api/download/{job_id}")
async def download_output(job_id: str):
    pool = await get_pool()
    row = await pool.fetchrow("SELECT output_path FROM jobs WHERE job_id = $1", job_id)
    if not row or not row["output_path"] or not Path(row["output_path"]).exists():
        raise HTTPException(404, "Output not ready")
    return FileResponse(
        row["output_path"],
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="Integrated_Supply_Plan.xlsx",
    )


@app.get("/api/checks")
async def get_checks():
    pool = await get_pool()
    rows = await pool.fetch("SELECT check_num, label, description, threshold, fail_type FROM checks ORDER BY check_num")
    return {"checks": [dict(r) for r in rows]}


@app.post("/api/add-check")
async def add_check(
    description: str = Form(...),
    threshold:   str = Form(...),
    fail_type:   str = Form("business_alert"),
    output_file: UploadFile = File(...),
):
    pool = await get_pool()
    row = await pool.fetchrow("SELECT COALESCE(MAX(check_num), 15) + 1 AS next_num FROM checks")
    next_num = row["next_num"]

    job_id = await new_job("add_check")
    job_dir = JOBS_DIR / job_id
    job_dir.mkdir()
    dest = job_dir / output_file.filename
    async with aiofiles.open(dest, "wb") as f:
        await f.write(await output_file.read())

    check = {
        "check_num": next_num,
        "label": description[:80],
        "description": description,
        "threshold": threshold,
        "fail_type": fail_type,
    }

    asyncio.create_task(_run_add_check(job_id, check, str(dest)))
    return {"job_id": job_id}


@app.get("/api/history")
async def get_history():
    pool = await get_pool()
    rows = await pool.fetch(
        "SELECT job_id, kind, status, month_label, summary, created_at FROM jobs WHERE status = 'done' ORDER BY created_at DESC LIMIT 50"
    )
    return {"history": [dict(r) for r in rows]}


async def _run_job(job_id, files, month_label):
    LIVE_STATUS[job_id] = "running"
    pool = await get_pool()
    await pool.execute("UPDATE jobs SET status = 'running' WHERE job_id = $1", job_id)

    def emit(event):
        push_event(job_id, event)

    try:
        from runner import PipelineRunner
        runner = PipelineRunner(job_id=job_id, push_event=emit)
        output_path, summary = await asyncio.to_thread(runner.run, files, month_label)
        emit({"type": "done", "summary": summary})
        await finalize_job(job_id, "done", output_path=output_path, summary=summary)
    except Exception as e:
        emit({"type": "error", "message": str(e)})
        await finalize_job(job_id, "error")


async def _run_add_check(job_id, check, output_file_path):
    LIVE_STATUS[job_id] = "running"
    pool = await get_pool()
    await pool.execute("UPDATE jobs SET status = 'running' WHERE job_id = $1", job_id)

    def emit(event):
        push_event(job_id, event)

    try:
        from claude_agent import ClaudeAgent
        from pathlib import Path as _P
        engine_path = _P(__file__).parent / "checklist_engine.py"

        agent = ClaudeAgent(push_event=emit)
        code, check_num = agent.write_new_check(
            description=check["description"],
            threshold=check["threshold"],
            fail_type=check["fail_type"],
            engine_source=engine_path.read_text(),
        )
        check["check_num"] = check_num

        # TODO: actually splice `code` into checklist_engine.py at the right spot.
        # Stubbed here — agent.write_new_check currently just returns code text.
        emit({"type": "claude", "message": f"Generated code for check #{check_num} (not yet auto-inserted)"})

        await pool.execute(
            """INSERT INTO checks (check_num, label, description, threshold, fail_type)
               VALUES ($1, $2, $3, $4, $5)""",
            check["check_num"], check["label"], check["description"], check["threshold"], check["fail_type"],
        )

        emit({"type": "done", "summary": {"check_num": check_num}})
        await finalize_job(job_id, "done", output_path=output_file_path, summary={"check_num": check_num})
    except Exception as e:
        emit({"type": "error", "message": str(e)})
        await finalize_job(job_id, "error")
