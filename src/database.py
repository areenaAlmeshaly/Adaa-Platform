
import json
from pathlib import Path
from datetime import datetime

PROJECT_ROOT = Path(__file__).resolve().parent.parent
RUNTIME_DIR = PROJECT_ROOT / "runtime"
DATA_DIR = PROJECT_ROOT / "data"

DATA_FILE = RUNTIME_DIR / "data.json"
SUMMARY_FILE = RUNTIME_DIR / "summary.json"
LIVE_FILE = RUNTIME_DIR / "live_status.json"
DB_FILE = DATA_DIR / "database.json"
TASKS_FILE = DATA_DIR / "completed_tasks.json"


def load_json(filepath):
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return {}


def save_json(filepath, data):
    filepath.parent.mkdir(parents=True, exist_ok=True)
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def init_db():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    RUNTIME_DIR.mkdir(parents=True, exist_ok=True)

    if not DB_FILE.exists():
        save_json(DB_FILE, {
            "employees": {},
            "created_at": datetime.now().strftime("%Y-%m-%d")
        })

    if not TASKS_FILE.exists():
        save_json(TASKS_FILE, [])


def save_analysis(data: dict):
    init_db()
    data["timestamp"] = datetime.now().isoformat()
    save_json(LIVE_FILE, data)

    db = load_json(DB_FILE)
    today = datetime.now().strftime("%Y-%m-%d")
    emp = db.setdefault("employees", {}).setdefault("emp_001", {
        "name": "أرينا المشعلي", "days": {}
    })
    day = emp["days"].setdefault(today, {"snapshots": [], "tasks_completed": 0})

    day["snapshots"].append({
        "time": data["timestamp"],
        "status": data.get("status"),
        "progress": data.get("progress", 0),
        "activity": data.get("activity", ""),
        "active_app": data.get("active_app", ""),
    })

    save_json(DB_FILE, db)


def save_completed_task(started_at: str, duration_minutes: int, activity: str):
    init_db()
    tasks = load_json(TASKS_FILE)

    if not isinstance(tasks, list):
        tasks = []

    tasks.append({
        "started_at": started_at,
        "completed_at": datetime.now().isoformat(),
        "duration_minutes": duration_minutes,
        "activity": activity,
    })

    save_json(TASKS_FILE, tasks)

    db = load_json(DB_FILE)
    today = datetime.now().strftime("%Y-%m-%d")

    try:
        day = db["employees"]["emp_001"]["days"][today]
        day["tasks_completed"] = day.get("tasks_completed", 0) + 1
        save_json(DB_FILE, db)
    except (KeyError, TypeError):
        pass


def get_latest_analysis() -> dict | None:
    if LIVE_FILE.exists():
        data = load_json(LIVE_FILE)
        if data:
            return data

    if not DATA_FILE.exists():
        return None

    raw = load_json(DATA_FILE)
    if not raw:
        return None

    current = raw.get("current_task")
    completed = raw.get("completed_tasks", [])
    productivity = raw.get("productivity", 0)
    active_ai = raw.get("active_ai")

    if current:
        status = "in_progress"
        progress = min(30 + current.get("saves", 0) * 15, 90)
        activity = f"يعمل على: {current.get('name','')} — {current.get('saves',0)} حفظ"
    elif productivity > 0:
        status = "in_progress"
        progress = productivity
        activity = f"نشاط مسجّل — إنتاجية {productivity}%"
    else:
        status = "not_started"
        progress = 0
        activity = "لا يوجد نشاط على Excel"

    if active_ai:
        activity += f" | يستخدم {active_ai.get('tool','AI')}"

    insight = f"أنجز {len(completed)} مهمة اليوم"

    if completed:
        last = completed[-1]
        insight += f" — آخرها: {last.get('name','')} ({last.get('duration','')})"

    return {
        "status": status,
        "progress": progress,
        "activity": activity,
        "insight": insight,
        "confidence": 85,
        "active_app": "Microsoft Excel",
        "timestamp": datetime.now().isoformat(),
    }


def get_today_stats() -> dict:
    if DATA_FILE.exists():
        raw = load_json(DATA_FILE)

        if raw:
            total_cells = sum(
                t.get("cells_added", 0)
                for t in raw.get("completed_tasks", [])
            )

            return {
                "snapshots": len(raw.get("snapshots", [])),
                "tasks_done": len(raw.get("completed_tasks", [])),
                "avg_progress_today": raw.get("productivity", 0),
                "total_cells_added": total_cells,
                "work_hours": 8,
            }

    return {
        "snapshots": 0,
        "tasks_done": 0,
        "avg_progress_today": 0,
        "total_cells_added": 0,
        "work_hours": 8
    }


def get_progress_history(limit: int = 20) -> list:
    init_db()
    today = datetime.now().strftime("%Y-%m-%d")
    db = load_json(DB_FILE)

    try:
        snaps = db["employees"]["emp_001"]["days"][today]["snapshots"]
    except (KeyError, TypeError):
        snaps = []

    return snaps[-limit:]


def get_employees() -> list:
    init_db()
    db = load_json(DB_FILE)
    today = datetime.now().strftime("%Y-%m-%d")
    out = []

    for emp_id, emp_data in db.get("employees", {}).items():
        day = emp_data.get("days", {}).get(today, {})
        snaps = day.get("snapshots", [])
        tasks_done = day.get("tasks_completed", 0)
        avg_prog = round(sum(s["progress"] for s in snaps) / len(snaps)) if snaps else 0

        out.append({
            "id": emp_id,
            "name": emp_data.get("name", ""),
            "tasks_done": tasks_done,
            "snapshots": len(snaps),
            "avg_progress": avg_prog,
            "total_cells_added": day.get("total_cells_added", 0),
        })

    return out


def get_completed_tasks() -> list:
    init_db()
    tasks = load_json(TASKS_FILE)

    if not isinstance(tasks, list):
        return []

    return list(reversed(tasks[-50:]))


def save_daily_summary(employee_id, employee_name, summary):
    init_db()
    db = load_json(DB_FILE)
    today = datetime.now().strftime("%Y-%m-%d")

    emp = db.setdefault("employees", {}).setdefault(employee_id, {
        "name": employee_name, "days": {}
    })

    emp["name"] = employee_name

    emp["days"][today] = {
        "date": today,
        "productivity_percent": summary.get("productivity", summary.get("productivity_percent", 0)),
        "completed_tasks": summary.get("total_completed", summary.get("completed_tasks", 0)),
        "total_cells_added": summary.get("total_cells_added", 0),
        "total_ai_minutes": summary.get("total_ai_minutes", 0),
        "tasks": summary.get("tasks", []),
        "tasks_completed": summary.get("total_completed", 0),
        "snapshots": [],
    }

    save_json(DB_FILE, db)
    print(f"تم حفظ ملخص {employee_name} ليوم {today}")


def compare_performance(employee_id):
    emp = get_employee_data(employee_id)

    if not emp or not emp.get("days"):
        return None

    days = emp["days"]
    dates = sorted(days.keys(), reverse=True)

    today_prod = days[dates[0]].get("productivity_percent", 0) if dates else 0
    yesterday_prod = days[dates[1]].get("productivity_percent", 0) if len(dates) > 1 else 0

    week_days = list(days.values())[-7:]
    week_avg = round(
        sum(d.get("productivity_percent", 0) for d in week_days) / len(week_days)
    ) if week_days else 0

    trend = "⬆️ ارتفع" if today_prod > yesterday_prod else (
        "⬇️ انخفض" if today_prod < yesterday_prod else "➡️ ثابت"
    )

    return {
        "today": today_prod,
        "yesterday": yesterday_prod,
        "week_average": week_avg,
        "trend": trend,
    }


def get_employee_data(employee_id):
    return load_json(DB_FILE).get("employees", {}).get(employee_id)


def get_today_summary(employee_id):
    emp = get_employee_data(employee_id)
    today = datetime.now().strftime("%Y-%m-%d")
    return emp["days"].get(today) if emp else None


def get_all_employees():
    return load_json(DB_FILE).get("employees", {})


def run_database():
    init_db()

    if not SUMMARY_FILE.exists():
        print("ما فيه summary.json — شغّل analyzer.py أول")
        return

    summary = load_json(SUMMARY_FILE)
    save_daily_summary("emp_001", "أرينا المشعلي", summary)

    comparison = compare_performance("emp_001")

    if comparison:
        print(f"\nالإنتاجية اليوم: {comparison['today']}%")
        print(f"الإنتاجية أمس: {comparison['yesterday']}%")
        print(f"متوسط الأسبوع: {comparison['week_average']}%")
        print(f"الاتجاه: {comparison['trend']}")


if __name__ == "__main__":
    run_database()
