import os
import subprocess
import sys
import re

TASK_NAME = "DailyReportAutomation_MailDispatcher"
OLD_TASK_NAME = "DailyReportAutomation_7PM"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
BAT_PATH = os.path.join(BASE_DIR, "run_scheduled_task.bat")

def get_task_info() -> tuple[bool, str]:
    """Checks if the Windows Task is registered and extracts its scheduled next run time or status."""
    if sys.platform != "win32":
        return False, ""
    for t_name in [TASK_NAME, OLD_TASK_NAME]:
        try:
            cmd = f'schtasks /query /tn "{t_name}" /fo LIST'
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            if result.returncode == 0 and t_name in result.stdout:
                # Extract Next Run Time line
                match = re.search(r"Next Run Time:\s*(.+)", result.stdout)
                next_run = match.group(1).strip() if match else "Registered"
                return True, next_run
        except Exception:
            pass
    return False, ""

def is_task_scheduled() -> bool:
    """Checks if either the new or old Windows Task is currently registered."""
    if sys.platform != "win32":
        return False
    ok, _ = get_task_info()
    return ok

def register_task(time_str: str = "19:00") -> tuple[bool, str]:
    """Registers a daily task in Windows Task Scheduler at any custom time (HH:MM in 24-hr format)."""
    if sys.platform != "win32":
        return False, "Windows Task Scheduler is only available when running locally on Windows."
    # Remove old task if it existed under the old name
    try:
        subprocess.run(f'schtasks /delete /tn "{OLD_TASK_NAME}" /f', shell=True, capture_output=True)
    except Exception:
        pass

    # Ensure batch file exists
    if not os.path.exists(BAT_PATH):
        with open(BAT_PATH, "w", encoding="utf-8") as f:
            f.write('@echo off\ncd /d "%~dp0"\nset "PATH=%LOCALAPPDATA%\\Programs\\Python\\Python312;%LOCALAPPDATA%\\Programs\\Python\\Python312\\Scripts;%PATH%"\npython send_today_report.py >> logs\\scheduler_bat.log 2>&1\n')

    cmd = f'schtasks /create /tn "{TASK_NAME}" /tr "\"{BAT_PATH}\"" /sc daily /st {time_str} /f'
    try:
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        if result.returncode == 0:
            return True, f"Successfully scheduled daily task '{TASK_NAME}' at {time_str}!"
        else:
            return False, result.stderr or result.stdout
    except Exception as e:
        return False, str(e)

def unregister_task() -> tuple[bool, str]:
    """Unregisters both new and old tasks from Windows Task Scheduler."""
    if sys.platform != "win32":
        return False, "Windows Task Scheduler is only available when running locally on Windows."
    deleted_any = False
    err_msgs = []
    for t_name in [TASK_NAME, OLD_TASK_NAME]:
        cmd = f'schtasks /delete /tn "{t_name}" /f'
        try:
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            if result.returncode == 0:
                deleted_any = True
            elif "cannot find" not in (result.stderr + result.stdout).lower():
                err_msgs.append(result.stderr or result.stdout)
        except Exception as e:
            err_msgs.append(str(e))
    
    if deleted_any:
        return True, "Successfully removed scheduled task."
    elif not err_msgs:
        return True, "No scheduled task found to remove."
    return False, " ".join(err_msgs)

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "unregister":
        ok, msg = unregister_task()
        print(msg)
    elif len(sys.argv) > 1:
        custom_time = sys.argv[1]
        ok, msg = register_task(custom_time)
        print(msg)
    else:
        ok, msg = register_task("19:00")
        print(msg)
