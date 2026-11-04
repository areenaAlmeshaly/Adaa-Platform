
import subprocess
import sys
import time
import webbrowser
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
RUNTIME_DIR = PROJECT_ROOT / "runtime"
RUNTIME_DIR.mkdir(exist_ok=True)

scripts = [
    PROJECT_ROOT / "src" / "agent.py",
    PROJECT_ROOT / "src" / "analyzer.py",
    PROJECT_ROOT / "src" / "server.py"
]

temporary_files = [
    "data.json",
    "summary.json",
    "live_status.json",
    "analyze_now.flag"
]

procs = []

try:
    for script in scripts:
        procs.append(
            subprocess.Popen(
                [sys.executable, str(script)],
                cwd=PROJECT_ROOT
            )
        )

    time.sleep(3)
    webbrowser.open("http://127.0.0.1:8000")

    print("كل شيء يعمل. اضغطي Ctrl+C لإيقاف البرنامج.")

    while True:
        time.sleep(1)

except KeyboardInterrupt:
    print("\nجاري إيقاف العمليات...")

finally:
    for process in procs:
        if process.poll() is None:
            process.terminate()

    for process in procs:
        process.wait()

    for filename in temporary_files:
        file_path = RUNTIME_DIR / filename

        if file_path.exists():
            file_path.unlink()
            print(f"تم حذف {filename}")

    print("تم إيقاف العمليات وتنظيف الملفات المؤقتة.")
