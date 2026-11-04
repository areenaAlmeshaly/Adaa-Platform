
import json
import time
import requests
import os
from pathlib import Path
from dotenv import load_dotenv
from datetime import datetime

PROJECT_ROOT = Path(__file__).resolve().parent.parent
RUNTIME_DIR = PROJECT_ROOT / "runtime"
DATA_DIR = PROJECT_ROOT / "data"

RUNTIME_DIR.mkdir(exist_ok=True)
DATA_DIR.mkdir(exist_ok=True)

load_dotenv(PROJECT_ROOT / ".env")

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent"

ANALYZE_EVERY = 60


def ask_gemini(prompt):
    try:
        if not GEMINI_API_KEY:
            raise ValueError("GEMINI_API_KEY غير موجود في ملف .env")

        payload = {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {
                "temperature": 0.3,
                "maxOutputTokens": 1500
            }
        }

        r = requests.post(
            GEMINI_URL,
            headers={
                "Content-Type": "application/json",
                "x-goog-api-key": GEMINI_API_KEY
            },
            json=payload,
            timeout=20
        )

        r.raise_for_status()
        return r.json()["candidates"][0]["content"]["parts"][0]["text"]

    except Exception as e:
        print(f" ❌ Gemini: {e}")
        return None


def gemini_analyze_task(task):
    sheets_info = "\n".join(
        f"- ورقة '{s['name']}': ({s['filled']} خانة)"
        for s in task.get("sheets", [])
    ) or "لا تفاصيل"

    prompt = f"""محلل أداء وظيفي. حلّل بيانات العمل التالية بالعربية باختصار مهني:
اسم الملف : {task.get('name','')}
المدة : {task.get('duration','')} | الحفظات : {task.get('saves',0)}
الخانات قبل: {task.get('initial_filled', 0)} | بعد : {task.get('final_filled',0)}
نسبة اكتمال الشيت : {task.get('completion_pct',0)}%
الأوراق :
{sheets_info}

أجب على 4 نقاط فقط:
1. **نوع العمل المحتمل**: ما نوع البيانات؟
2. **تقييم الإنتاجية**: هل الإنجاز جيد مقارنة بالوقت؟
3. **ملاحظة مهمة**: شيء إيجابي أو سلبي
4. **توصية للمشرف**: جملة واحدة
"""

    resp = ask_gemini(prompt)

    return {
        "gemini_analysis": resp,
        "analyzed_at": datetime.now().strftime("%H:%M:%S")
    }


def gemini_daily_report(data):
    completed = data.get("completed_tasks", [])

    if not completed:
        return None

    tasks_txt = "\n".join(
        f"مهمة {i}: {t.get('name','')} ({t.get('duration','')})"
        for i, t in enumerate(completed, 1)
    )

    ai_txt = "\n".join(
        f"- {s.get('tool','')}: {s.get('duration_minutes', 0)} دقيقة"
        for s in data.get("ai_sessions", [])
    ) or "لا يوجد"

    prompt = f"""أنت محلل أداء في جهة حكومية. قيم أداء الموظف:
الموظف : {data.get('employee','')}
الإنتاجية : {data.get('productivity', 0)}% | المهام : {len(completed)}
المهام المنجزة:
{tasks_txt}
استخدام AI:
{ai_txt}
اكتب تقريراً رسمياً في 5 نقاط:
1. **التقييم العام** (ممتاز / جيد / مقبول / ضعيف) مع مبرر
2. **أبرز إنجاز اليوم**
3. **نقطة تحتاج تحسين**
4. **تقييم استخدام AI**
5. **توصية للمشرف** (جملة واحدة حازمة)
. أسلوب رسمي مختصر مناسب للتقارير الحكومية
"""

    resp = ask_gemini(prompt)

    return {
        "daily_report": resp,
        "generated_at": datetime.now().strftime("%H:%M:%S")
    }


def analyze():
    try:
        with open(RUNTIME_DIR / "data.json", "r", encoding="utf-8") as f:
            data = json.load(f)

    except (FileNotFoundError, json.JSONDecodeError, OSError) as e:
        print(f" ❌ تعذر قراءة data.json: {e}")
        return

    completed = data.get("completed_tasks", [])
    current = data.get("current_task")
    ai_sess = data.get("ai_sessions", [])
    snapshots = data.get("snapshots", [])
    prod = data.get("productivity", 0)

    cells_done = sum(
        max(0, t.get("cells_added", 0))
        for t in completed
    )

    cells_current = (
        max(0, current.get("cells_added", 0))
        if current else 0
    )

    total_cells = cells_done + cells_current

    saves_done = sum(
        t.get("saves", 0)
        for t in completed
    )

    saves_current = (
        current.get("saves", 0)
        if current else 0
    )

    total_saves = saves_done + saves_current

    analyzed = []

    for task in completed:
        start = task.get("start", "")
        end = task.get("end", "")
        dur = task.get("duration_min", 1)

        ai_dur = [
            s for s in ai_sess
            if start <= s.get("start", "") <= end
        ]

        ai_min = sum(
            s.get("duration_minutes", 0)
            for s in ai_dur
        )

        ai_pct = round((ai_min / max(dur, 1)) * 100)

        task["ai_minutes"] = ai_min
        task["ai_percent"] = ai_pct
        task["ai_verdict"] = _verdict(ai_pct, ai_min)

        if not task.get("gemini_analysis"):
            print(f" 🤖 Gemini يحلل: {task.get('name','')}")

            g = gemini_analyze_task(task)

            if g and g.get("gemini_analysis"):
                task.update(g)
            else:
                task["gemini_analysis"] = "تعذر التحليل"

        analyzed.append(task)

    hours = {}

    for s in snapshots:
        h = s.get("time", "")[:2]

        if not h:
            continue

        hours.setdefault(
            h,
            {"total": 0, "excel": 0}
        )

        hours[h]["total"] += 1

        if s.get("excel"):
            hours[h]["excel"] += 1

    energy = [
        {
            "hour": h + ":00",
            "productivity": round(
                d["excel"] / max(d["total"], 1) * 100
            )
        }
        for h, d in sorted(hours.items())
    ]

    peak = (
        max(energy, key=lambda x: x["productivity"])["hour"]
        if energy else "-"
    )

    ai_analysis = []

    for s in ai_sess:
        end_ai = s.get("end", "")

        after = any(
            t.get("start", "") >= end_ai
            for t in completed
        )

        ai_analysis.append({
            "tool": s.get("tool", "AI"),
            "start": s.get("start"),
            "end": end_ai,
            "minutes": s.get("duration_minutes", 0),
            "verdict": (
                "استخدام منتج"
                if after
                else "استخدام بدون إنجاز بعده!"
            ),
            "productive": after
        })

    alerts = []

    consec = 0
    mx = 0

    for s in snapshots:
        if not s.get("excel"):
            consec += 1
            mx = max(mx, consec)
        else:
            consec = 0

    idle_min = round(mx * (15 / 60), 1)

    if idle_min >= 20:
        alerts.append({
            "type": "warning",
            "icon": "⚠️",
            "title": "خمول متواصل",
            "msg": f"{idle_min} دقيقة بدون عمل"
        })

    if prod < 50 and len(snapshots) > 12:
        alerts.append({
            "type": "warning",
            "icon": "📉",
            "title": "إنتاجية منخفضة",
            "msg": f"{prod}% فقط"
        })

    nonprod = [
        a for a in ai_analysis
        if not a["productive"] and a["minutes"] > 10
    ]

    if nonprod:
        alerts.append({
            "type": "danger",
            "icon": "🛑",
            "title": "استخدام AI غير منتج",
            "msg": f"{len(nonprod)} جلسات بدون إنجاز"
        })

    if total_cells >= 30 and prod >= 60:
        alerts.append({
            "type": "success",
            "icon": "🌟",
            "title": "أداء ممتاز",
            "msg": f"{total_cells} خلية و {prod}% إنتاجية"
        })

    if current:
        alerts.append({
            "type": "info",
            "icon": "🔄",
            "title": f"جاري: {current.get('name','')}",
            "msg": f"أضافت {cells_current} خانة حتى الآن"
        })

    comparison = None

    try:
        with open(DATA_DIR / "database.json", "r", encoding="utf-8") as f:
            db = json.load(f)

        days = (
            db.get("employees", {})
            .get("emp_001", {})
            .get("days", {})
        )

        sd = sorted(days.keys(), reverse=True)

        if len(sd) >= 2:
            td, yd = days[sd[0]], days[sd[1]]

            diff = (
                td.get("productivity_percent", 0)
                - yd.get("productivity_percent", 0)
            )

            comparison = {
                "today": td.get("productivity_percent", 0),
                "yesterday": yd.get("productivity_percent", 0),
                "diff": diff,
                "trend": (
                    "↑" if diff > 0
                    else "↓" if diff < 0
                    else "-"
                ),
                "color": (
                    "green" if diff > 0
                    else "red" if diff < 0
                    else "neutral"
                ),
                "today_tasks": td.get("completed_tasks", 0),
                "yesterday_tasks": td.get("completed_tasks", 0),
                "today_cells": td.get("total_cells_added", 0),
                "yesterday_cells": yd.get("total_cells_added", 0)
            }

    except (FileNotFoundError, json.JSONDecodeError, OSError, AttributeError):
        pass

    daily_report = None

    if completed:
        print(" 📝 Gemini يكتب التقرير اليومي...")

        daily_report = gemini_daily_report({
            "employee": data.get("employee", ""),
            "productivity": prod,
            "completed_tasks": completed,
            "ai_sessions": ai_sess
        })

    result = {
        "date": datetime.now().strftime("%Y-%m-%d"),
        "generated_at": datetime.now().strftime("%H:%M:%S"),
        "productivity": prod,
        "tasks": analyzed,
        "current_task": current,
        "total_completed": len(completed),
        "total_cells_added": total_cells,
        "total_saves": total_saves,
        "ai_analysis": ai_analysis,
        "total_ai_minutes": sum(
            s.get("duration_minutes", 0)
            for s in ai_sess
        ),
        "energy_curve": energy,
        "peak_hour": peak,
        "alerts": alerts,
        "comparison": comparison,
        "daily_report": daily_report
    }

    with open(RUNTIME_DIR / "summary.json", "w", encoding="utf-8") as f:
        json.dump(
            result,
            f,
            ensure_ascii=False,
            indent=2
        )

    print(
        f" ✅ [{result['generated_at']}] "
        f"مهام: {len(completed)} | "
        f"خانات: {total_cells} "
        f"(منها {cells_current} جارية)"
    )

    return result


def _verdict(pct, mn):
    if mn == 0:
        return None

    if pct < 20:
        return {
            "level": "مستقل",
            "icon": "🟢",
            "color": "green"
        }

    elif pct < 50:
        return {
            "level": "متوازن",
            "icon": "🔵",
            "color": "blue"
        }

    elif pct < 80:
        return {
            "level": "اعتماد عالي",
            "icon": "🟡",
            "color": "yellow"
        }

    else:
        return {
            "level": "اعتماد كلي",
            "icon": "🔴",
            "color": "red"
        }


if __name__ == "__main__":
    print("=== Analyzer + Gemini | كل 60 ثانية + فوري عند إغلاق Excel ===")

    last_run = 0
    flag_path = RUNTIME_DIR / "analyze_now.flag"

    while True:
        try:
            now_t = time.time()

            if flag_path.exists():
                flag_path.unlink()

                print(" ⚡ تحليل فوري — تم إغلاق Excel")

                analyze()
                last_run = time.time()

            elif now_t - last_run >= ANALYZE_EVERY:
                analyze()
                last_run = time.time()

        except KeyboardInterrupt:
            break

        except Exception as e:
            print(f"X {e}")

        time.sleep(5)
