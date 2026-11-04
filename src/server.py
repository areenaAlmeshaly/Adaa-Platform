
from pathlib import Path
import json
from datetime import datetime

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

PROJECT_ROOT = Path(__file__).resolve().parent.parent
RUNTIME_DIR = PROJECT_ROOT / "runtime"
DATA_DIR = PROJECT_ROOT / "data"

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


def rj(filename, directory=RUNTIME_DIR):
    file_path = directory / filename

    try:
        with open(file_path, "r", encoding="utf-8") as fp:
            return json.load(fp)
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return None


@app.get("/api/today")
def today():
    data = rj("data.json") or {}
    summary = rj("summary.json") or {}

    completed = data.get("completed_tasks", [])
    current = data.get("current_task")
    ai_sess = data.get("ai_sessions", [])
    snapshots = data.get("snapshots", [])
    prod = data.get("productivity", 0)

    cells_done = sum(
        max(0, t.get("cells_added", 0))
        for t in completed
    )
    cells_current = max(0, current.get("cells_added", 0)) if current else 0
    total_cells = cells_done + cells_current

    saves_done = sum(t.get("saves", 0) for t in completed)
    saves_current = current.get("saves", 0) if current else 0
    total_saves = saves_done + saves_current

    active_min = round(
        sum(1 for s in snapshots if s.get("excel")) * (15 / 60), 1
    )
    idle_min = round(
        sum(1 for s in snapshots if not s.get("excel")) * (15 / 60), 1
    )
    ai_min = round(
        sum(s.get("duration_minutes", 0) for s in ai_sess), 1
    )

    last = snapshots[-1] if snapshots else {}

    if last.get("excel"):
        emp_status = f"🟢 نشطة — {last.get('task', 'Excel')}"
    elif last.get("ai"):
        emp_status = "🤖 تستخدم AI"
    elif snapshots:
        emp_status = "🟡 خارج Excel"
    else:
        emp_status = "⚪ في انتظار البيانات..."

    timeline = []

    for s in snapshots[-15:]:
        if s.get("excel"):
            icon, label, color = "🟢", f"Excel — {s.get('task', '')}", "green"
        elif s.get("ai"):
            icon, label, color = "🤖", "تستخدم AI", "purple"
        else:
            icon, label, color = "🟡", "خارج Excel", "yellow"

        timeline.append({
            "time": s.get("time", ""),
            "icon": icon,
            "label": label,
            "color": color,
            "saves": s.get("saves", 0),
            "filled": s.get("filled", 0),
            "progress": s.get("progress", 0),
        })

    analyzed_tasks = summary.get("tasks", completed)
    tasks = list(analyzed_tasks)

    if current:
        tasks.append(current)

    return {
        "productivity": prod,
        "active_minutes": active_min,
        "idle_minutes": idle_min,
        "total_ai_minutes": ai_min,
        "total_ai_sessions": len(ai_sess),
        "total_cells_added": total_cells,
        "total_saves": total_saves,
        "emp_status": emp_status,
        "current_task": current,
        "tasks": tasks,
        "alerts": summary.get("alerts", []),
        "energy_curve": summary.get("energy_curve", []),
        "peak_hour": summary.get("peak_hour", "--"),
        "ai_analysis": summary.get("ai_analysis", []),
        "comparison": summary.get("comparison"),
        "daily_report": summary.get("daily_report"),
        "timeline": timeline,
        "total_completed": len(completed),
        "last_update": datetime.now().strftime("%H:%M:%S"),
    }


@app.get("/api/database")
def database():
    db = rj("database.json", DATA_DIR)

    if db is not None:
        return {"status": "ok", "data": db}

    return {"status": "empty", "data": {}}


@app.get("/api/status")
def status():
    return {
        "agent": (RUNTIME_DIR / "data.json").exists(),
        "analyzer": (RUNTIME_DIR / "summary.json").exists(),
        "database": (DATA_DIR / "database.json").exists(),
        "time": datetime.now().strftime("%H:%M:%S"),
    }


@app.get("/")
def home():
    dashboard_path = PROJECT_ROOT / "frontend" / "dashboard.html"
    return FileResponse(dashboard_path, headers={"Cache-Control": "no-store"})


if __name__ == "__main__":
    import uvicorn

    print("🚀 http://127.0.0.1:8000")
    uvicorn.run(app, host="127.0.0.1", port=8000)
