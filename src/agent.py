
import subprocess
import glob
import psutil
import openpyxl
import shutil
import os
import time
import json
import tempfile
from pathlib import Path
from datetime import datetime

PROJECT_ROOT = Path(__file__).resolve().parent.parent
RUNTIME_DIR = PROJECT_ROOT / "runtime"
RUNTIME_DIR.mkdir(exist_ok=True)

CHECK_INTERVAL = 15
RETURN_WINDOW = 30
EMPLOYEE_NAME = "أرينا المشعلي"
EMPLOYEE_ID = "emp_001"

AI_TITLES = [
    'chatgpt', 'chat.openai', 'claude.ai', 'anthropic',
    'gemini', 'copilot', 'bing.com/chat', 'perplexity',
    'grok', 'poe.com', 'you.com', 'bard'
]

FRONT_WINDOW_SCRIPT = '''
tell application "System Events"
    set p to first application process whose frontmost is true
    set appName to name of p
    set appPid to unix id of p
    try
        set winTitle to name of front window of p
    on error
        set winTitle to ""
    end try
    return appName & "|||" & appPid & "|||" & winTitle
end tell
'''

EXCEL_JXA = r'''
function run(argv) {
  var target = (argv[0] || "").toLowerCase().trim();
  function base(n) { return n.replace(/\.(xlsx|xlsm|xlsb|xls)$/i, "").toLowerCase().trim(); }
  try {
    var xl = Application("Microsoft Excel");
    if (!xl.running()) return "null";
    var wb = null;
    if (target === "") {
      wb = xl.activeWorkbook;
    } else {
      var wbs = xl.workbooks();
      for (var i = 0; i < wbs.length; i++) {
        var b = base(wbs[i].name());
        if (b === target || b.indexOf(target) >= 0 || target.indexOf(b) >= 0) { wb = wbs[i]; break; }
      }
    }
    if (!wb) return "null";
    var sheets = wb.worksheets();
    var out = [], total = 0;
    for (var s = 0; s < sheets.length; s++) {
      var filled = 0, rows = 0, cols = 0;
      try {
        var v = sheets[s].usedRange.value();
        if (Array.isArray(v)) {
          rows = v.length;
          for (var r = 0; r < v.length; r++) {
            var row = v[r];
            if (!Array.isArray(row)) row = [row];
            cols = Math.max(cols, row.length);
            for (var c = 0; c < row.length; c++) {
              var x = row[c];
              if (x !== null && x !== undefined && String(x).trim() !== "") filled++;
            }
          }
        } else if (v !== null && v !== undefined && String(v).trim() !== "") {
          filled = 1; rows = 1; cols = 1;
        }
      } catch (e) {}
      out.push({name: sheets[s].name(), filled: filled, rows: rows, cols: cols});
      total += filled;
    }
    var path = ""; try { path = wb.fullName(); } catch (e) {}
    return JSON.stringify({name: wb.name(), path: path, saved: wb.saved(), filled: total, sheets: out});
  } catch (e) { return "null"; }
}
'''


def _osascript(args, timeout=20):
    try:
        r = subprocess.run(
            ["osascript"] + args,
            capture_output=True,
            text=True,
            timeout=timeout
        )
        return (r.stdout or r.stderr or "").strip()
    except Exception:
        return ""


def get_active_window():
    try:
        out = _osascript(["-e", FRONT_WINDOW_SCRIPT], timeout=5)
        app, pid, title = (out.split("|||") + ["", "", ""])[:3]
        return {
            "title": title.strip(),
            "exe": app.strip().lower(),
            "pid": int(pid) if pid.isdigit() else 0
        }
    except Exception:
        return {"title": "", "exe": "", "pid": 0}


def is_excel_process_running():
    for proc in psutil.process_iter(['name']):
        try:
            if 'excel' in (proc.info['name'] or '').lower():
                return True
        except Exception:
            pass
    return False


def is_excel(w):
    return "microsoft excel" in w["exe"]


def is_ai(w):
    t = w["title"].lower()

    for k in AI_TITLES:
        if k in t:
            label = {
                "chatgpt": "ChatGPT",
                "chat.openai": "ChatGPT",
                "claude.ai": "Claude",
                "anthropic": "Claude",
                "gemini": "Gemini",
                "copilot": "Copilot",
                "bing.com/chat": "Copilot",
                "perplexity": "Perplexity",
                "grok": "Grok",
                "poe.com": "Poe",
                "you.com": "You.com",
                "bard": "Bard"
            }.get(k, k.capitalize())

            return True, label

    return False, None


def strip_ext(name):
    base, ext = os.path.splitext(name)
    return base if ext.lower() in (".xlsx", ".xlsm", ".xlsb", ".xls") else name


def get_excel_state(file_name=None):
    out = _osascript(
        ["-l", "JavaScript", "-e", EXCEL_JXA, file_name or ""]
    )

    if not out or out == "null":
        return None

    try:
        d = json.loads(out)
        d["name"] = strip_ext(d["name"])
        return d
    except Exception:
        return None


def find_excel_file_fallback(file_name):
    extensions = [".xlsx", ".xlsm", ".xls", ".xlsb"]

    search_dirs = [
        os.path.expanduser("~/Desktop"),
        os.path.expanduser("~/Documents"),
        os.path.expanduser("~/Downloads"),
    ] + glob.glob(os.path.expanduser("~/Library/CloudStorage/*"))

    for directory in search_dirs:
        if not os.path.exists(directory):
            continue

        for ext in extensions:
            path = os.path.join(directory, file_name + ext)
            if os.path.exists(path):
                return path

    for directory in search_dirs:
        if not os.path.exists(directory):
            continue

        for root, dirs, files in os.walk(directory):
            depth = root.replace(directory, '').count(os.sep)

            if depth > 3:
                dirs[:] = []
                continue

            dirs[:] = [
                x for x in dirs
                if not x.startswith(('.', '$', 'System'))
            ]

            for filename in files:
                if os.path.splitext(filename)[0].lower() == file_name.lower():
                    return os.path.join(root, filename)

    return None


def get_excel_path(file_name, state=None):
    if state and state.get("path") and os.path.exists(state["path"]):
        return state["path"]

    return find_excel_file_fallback(file_name)


def get_cells_via_file(filepath):
    if not filepath or not os.path.exists(filepath):
        return None

    tmp_path = None

    try:
        tmp_path = os.path.join(
            tempfile.gettempdir(),
            f"tr_{os.path.basename(filepath)}"
        )

        shutil.copy2(filepath, tmp_path)

        wb = openpyxl.load_workbook(
            tmp_path,
            read_only=True,
            data_only=True
        )

        total_filled = 0
        sheets_info = []

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]

            filled = sum(
                1
                for row in ws.iter_rows()
                for cell in row
                if cell.value is not None
                and str(cell.value).strip() not in ("", "None")
            )

            sheets_info.append({
                "name": sheet_name,
                "filled": filled
            })

            total_filled += filled

        wb.close()

        return {
            "filled": total_filled,
            "sheets": sheets_info
        }

    except Exception:
        return None

    finally:
        if tmp_path and os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass


def get_file_mtime(filepath):
    try:
        return os.path.getmtime(filepath)
    except Exception:
        return 0


class ExcelTracker:

    def __init__(self):
        self.completed = []
        self.ai_sessions = []
        self.current = None
        self.current_ai = None
        self.away_since = None
        self.snapshots = []
        self.excel_seconds = 0
        self.total_seconds = 0
        self.last_closed_file = None
        self.last_closed_task = None

    def _new_task(self, file_name, now, state, is_unsaved):
        com_data = state

        if not com_data or com_data["filled"] == 0:
            fp = get_excel_path(file_name, state)
            fd = get_cells_via_file(fp) if fp else None

            if fd and fd["filled"] > 0:
                com_data = fd
                print(
                    f"[{now}] ⚠️ قراءة Excel فشلت — "
                    f"قرأت من الملف مباشرة: {fd['filled']} خانة"
                )

        filled = com_data["filled"] if com_data else 0
        file_path = get_excel_path(file_name, state)

        return {
            "name": file_name,
            "path": file_path,
            "start": now,
            "start_epoch": time.time(),
            "saves": 0,
            "is_unsaved": is_unsaved,
            "last_mtime": get_file_mtime(file_path) if file_path else 0,
            "com_was_saved": True,
            "initial_filled": filled,
            "current_filled": filled,
            "sheets": com_data["sheets"] if com_data else [],
        }

    def process(self, window):
        now = datetime.now().strftime("%H:%M:%S")
        excel = is_excel(window)
        ai_flag, ai_name = is_ai(window)

        self.total_seconds += CHECK_INTERVAL

        if excel:
            self.excel_seconds += CHECK_INTERVAL

        if ai_flag:
            if self.current_ai is None:
                self.current_ai = {
                    "tool": ai_name,
                    "start": now,
                    "start_epoch": time.time()
                }

        elif self.current_ai is not None:
            dur = round(
                (time.time() - self.current_ai["start_epoch"]) / 60,
                1
            )

            self.current_ai.update({
                "end": now,
                "duration_minutes": dur
            })

            self.ai_sessions.append(self.current_ai)
            self.current_ai = None

        snap = {
            "time": now,
            "excel": excel,
            "ai": ai_flag,
            "title": window["title"][:80],
            "task": None,
            "saves": 0,
            "filled": 0,
            "progress": 0,
            "cells_added": 0
        }

        if excel:
            state = get_excel_state()

            if not state:
                self.snapshots.append(snap)
                return

            file_name = state["name"]
            is_unsaved = not state["saved"]
            self.away_since = None
            new_filled = state["filled"]

            if self.current is None:
                if self.last_closed_file and self._same_name(
                    file_name,
                    self.last_closed_file
                ):
                    print(f"[{now}] رجعت لنفس الملف — إلغاء التاسك")

                    if self.last_closed_task and self.last_closed_task in self.completed:
                        self.completed.pop()

                    self.last_closed_file = None
                    self.last_closed_task = None

                self.current = self._new_task(
                    file_name,
                    now,
                    state,
                    is_unsaved
                )

                print(
                    f"[{now}] بدأ: {file_name} | "
                    f"{self.current['initial_filled']} خانة"
                )

            elif self._same_name(file_name, self.current["name"]):
                if (
                    new_filled >= self.current["current_filled"]
                    and new_filled > 0
                ):
                    delta = new_filled - self.current["current_filled"]
                    self.current["current_filled"] = new_filled
                    self.current["sheets"] = state["sheets"]

                    if delta > 0:
                        print(
                            f"[{now}] +{delta} خانة "
                            f"({new_filled} إجمالي)"
                        )

                com_saved = bool(state["saved"])
                prev_saved = self.current.get("com_was_saved", True)

                if com_saved and prev_saved is False:
                    self.current["saves"] += 1
                    print(f"[{now}] 💾 حفظ #{self.current['saves']}")

                self.current["com_was_saved"] = com_saved

                fp = self.current.get("path")

                if fp and os.path.exists(fp):
                    new_mtime = get_file_mtime(fp)

                    if new_mtime > self.current["last_mtime"] + 0.5:
                        self.current["saves"] += 1
                        self.current["last_mtime"] = new_mtime
                        print(
                            f"[{now}] 💾 حفظ mtime "
                            f"#{self.current['saves']}"
                        )

                self.current["is_unsaved"] = is_unsaved
                snap["saves"] = self.current["saves"]

            else:
                done = self._finalize(now, f"انتقل لـ {file_name}")

                if done:
                    print(
                        f"[{now}] ✅ {done['name']} | "
                        f"+{done['cells_added']}"
                    )

                self.current = self._new_task(
                    file_name,
                    now,
                    state,
                    is_unsaved
                )

        else:
            if self.current is not None:
                if self.away_since is None:
                    self.away_since = time.time()
                    print(f"[{now}] غادر Excel")

                away = time.time() - self.away_since

                if away >= RETURN_WINDOW:
                    if not is_excel_process_running():
                        fp = self.current.get("path")
                        fd = get_cells_via_file(fp)

                        if (
                            fd
                            and fd["filled"] > self.current["current_filled"]
                        ):
                            self.current["current_filled"] = fd["filled"]
                            self.current["sheets"] = fd["sheets"]

                    if (
                        self.current["current_filled"]
                        != self.current["initial_filled"]
                    ):
                        self.current["saves"] += 1
                        print(
                            f"[{now}] غياب {away:.0f}ث = "
                            f"حفظة #{self.current['saves']}"
                        )

                    done = self._finalize(now, f"غادر {away:.0f}ث")

                    if done:
                        self.last_closed_file = done["name"]
                        self.last_closed_task = done

                        print(
                            f"[{now}] ✅ تاسك منجز: {done['name']} | "
                            f"+{done['cells_added']}"
                        )

        if self.current is not None:
            cells_added = (
                self.current["current_filled"]
                - self.current["initial_filled"]
            )

            snap["task"] = self.current["name"]
            snap["filled"] = self.current["current_filled"]
            snap["saves"] = self.current["saves"]
            snap["cells_added"] = cells_added

            snap["progress"] = (
                round(
                    cells_added / self.current["current_filled"] * 100,
                    1
                )
                if self.current["current_filled"] > 0
                else 0
            )

        self.snapshots.append(snap)

    def _same_name(self, a, b):
        a = a.lower().strip()
        b = b.lower().strip()
        return a == b or a in b or b in a

    def _finalize(self, end_time, reason):
        if not self.current:
            return None

        c = self.current

        duration = round(
            (time.time() - c["start_epoch"]) / 60,
            1
        )

        cells_added = c["current_filled"] - c["initial_filled"]

        if c["saves"] == 0 and cells_added <= 0:
            fp = c.get("path") or get_excel_path(c["name"])

            if fp:
                fd = get_cells_via_file(fp)

                if fd and fd["filled"] > 0:
                    print(
                        f"[{end_time}] 🔄 fallback من الملف: "
                        f"{fd['filled']} خانة"
                    )

                    c["current_filled"] = fd["filled"]
                    c["sheets"] = fd["sheets"]
                    cells_added = (
                        c["current_filled"] - c["initial_filled"]
                    )

        if c["saves"] == 0 and cells_added <= 0:
            print(f"[{end_time}] أهمل: {c['name']}")
            self.current = self.away_since = None
            return None

        task = {
            "name": c["name"],
            "path": c["path"],
            "start": c["start"],
            "end": end_time,
            "duration": f"{duration} دقيقة",
            "duration_min": duration,
            "saves": c["saves"],
            "initial_filled": c["initial_filled"],
            "final_filled": c["current_filled"],
            "cells_added": cells_added,
            "completion_pct": (
                round(
                    cells_added / c["current_filled"] * 100,
                    1
                )
                if c["current_filled"] > 0
                else 0
            ),
            "sheets": c["sheets"],
            "reason": reason,
            "status": "مكتمل",
            "quality": self._quality(
                c["saves"],
                duration,
                cells_added
            ),
        }

        self.completed.append(task)

        try:
            with open(
                RUNTIME_DIR / "analyze_now.flag",
                "w",
                encoding="utf-8"
            ) as f:
                f.write(end_time)
        except OSError:
            pass

        self.current = self.away_since = None
        return task

    def _quality(self, saves, duration, cells_added):
        if cells_added >= 50 or (saves >= 3 and duration >= 30):
            return {
                "label": "عمل مكثف",
                "icon": "🟢",
                "color": "green"
            }

        elif cells_added >= 20 or (saves >= 2 and duration >= 15):
            return {
                "label": "عمل جيد",
                "icon": "🟢",
                "color": "green"
            }

        elif cells_added >= 5 or saves >= 1:
            return {
                "label": "تعديل معقول",
                "icon": "🔵",
                "color": "blue"
            }

        return {
            "label": "تعديل بسيط",
            "icon": "🟡",
            "color": "yellow"
        }

    def summary(self):
        prod = (
            round((self.excel_seconds / self.total_seconds) * 100)
            if self.total_seconds > 0
            else 0
        )

        current_info = None

        if self.current:
            c = self.current
            cells_added = (
                c["current_filled"] - c["initial_filled"]
            )

            current_info = {
                "name": c["name"],
                "start": c["start"],
                "saves": c["saves"],
                "initial_filled": c["initial_filled"],
                "current_filled": c["current_filled"],
                "cells_added": cells_added,
                "completion_pct": (
                    round(
                        cells_added / c["current_filled"] * 100,
                        1
                    )
                    if c["current_filled"] > 0
                    else 0
                ),
                "sheets": c["sheets"],
                "status": "جاري",
            }

        cells_done = sum(
            max(0, t.get("cells_added", 0))
            for t in self.completed
        )

        cells_current = (
            max(
                0,
                self.current["current_filled"]
                - self.current["initial_filled"]
            )
            if self.current
            else 0
        )

        return {
            "employee": EMPLOYEE_NAME,
            "employee_id": EMPLOYEE_ID,
            "productivity": prod,
            "completed_tasks": self.completed,
            "total_completed": len(self.completed),
            "current_task": current_info,
            "ai_sessions": self.ai_sessions,
            "active_ai": (
                {
                    "tool": self.current_ai["tool"],
                    "start": self.current_ai["start"]
                }
                if self.current_ai
                else None
            ),
            "snapshots": self.snapshots[-120:],
            "total_cells_added": cells_done + cells_current,
            "excel_seconds": self.excel_seconds,
            "total_seconds": self.total_seconds,
        }


def run():
    print("=" * 60)
    print(f"🟢 Excel Tracker v5 (macOS) — {EMPLOYEE_NAME}")
    print(
        f"⏱️  فحص كل {CHECK_INTERVAL}ث | "
        f"غياب {RETURN_WINDOW}ث = تاسك منجز"
    )
    print("🛑 Ctrl+C للإيقاف")
    print("=" * 60)

    tracker = ExcelTracker()

    while True:
        try:
            window = get_active_window()
            tracker.process(window)

            with open(
                RUNTIME_DIR / "data.json",
                "w",
                encoding="utf-8"
            ) as f:
                json.dump(
                    tracker.summary(),
                    f,
                    ensure_ascii=False,
                    indent=2
                )

            time.sleep(CHECK_INTERVAL)

        except KeyboardInterrupt:
            print("\n🛑 إيقاف...")
            break

        except Exception as e:
            print(f"⚠️ {e}")
            time.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    run()
