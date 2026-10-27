import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import uuid
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
from flask import Flask, abort, flash, jsonify, redirect, render_template, request, send_from_directory, url_for

BASE_DIR = Path(__file__).resolve().parent
SCRIPTS_DIR = BASE_DIR / "scripts"
DATA_DIR = BASE_DIR / "data"
LOGS_DIR = BASE_DIR / "logs"

FOLDERS = {
    "word": DATA_DIR / "word",
    "connection": DATA_DIR / "connection",
    "csm": DATA_DIR / "csm",
    "corrections": DATA_DIR / "corrections",
    "report": DATA_DIR / "report",
}

for folder in [SCRIPTS_DIR, DATA_DIR, LOGS_DIR, *FOLDERS.values()]:
    folder.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get("WEB_APP_SECRET", "local-statistics-app")
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024

SCRIPTS = {
    "word_excel": {
        "title": "Word → Excel",
        "description": "Обрабатывает все .docx в папке операции и создаёт .xlsx рядом с ними.",
        "script": "script_conferter_word_excel.py",
        "folder_key": "word",
        "stdin": lambda form: None,
        "output_extensions": {".xlsx"},
    },
    "connection": {
        "title": "Объединение Excel",
        "description": "Объединяет все .xlsx из папки операции в один файл с диапазоном дат.",
        "script": "script_connection_excel.py",
        "folder_key": "connection",
        "stdin": lambda form: None,
        "output_extensions": {".xlsx"},
    },
    "csm": {
        "title": "CSM лог → Excel",
        "description": "Преобразует выбранный .txt-лог в Excel.",
        "script": "script_convertor_csm.py",
        "folder_key": "csm",
        "stdin": lambda form: f"{Path(form['input_file']).stem}\n{form.get('output_name', '')}\n",
        "output_extensions": {".xlsx"},
    },
    "delete_prints": {
        "title": "Исправление ФИО в Excel",
        "description": "Вносит предусмотренные исходным скриптом замены непосредственно в выбранный Excel-файл.",
        "script": "script_for_delete_prints.py",
        "folder_key": "corrections",
        "stdin": lambda form: f"{form['input_file']}\n",
        "output_extensions": {".xlsx"},
    },
    "report_by_date": {
        "title": "Отчёт по группам за период",
        "description": "Формирует отчёт по группам, ресурсу и комментарию за выбранный период.",
        "script": "script_statistic_to_date_filter.py",
        "folder_key": "report",
        "stdin": lambda form: (
            f"{form['serial_mode']}\n"
            f"{form['groups_file']}\n"
            f"{form['results_file']}\n"
            f"{form['output_name']}\n"
            f"{form['period_start']}\n"
            f"{form['period_end']}\n"
        ),
        "output_extensions": {".xlsx"},
    },
}

RUNS = {}
RUNS_LOCK = threading.Lock()


def list_files(folder_key, extensions=None):
    folder = FOLDERS[folder_key]
    result = []
    for path in sorted(folder.iterdir(), key=lambda p: p.name.lower()):
        if not path.is_file():
            continue
        if extensions and path.suffix.lower() not in extensions:
            continue
        stat = path.stat()
        result.append({
            "name": path.name,
            "size": stat.st_size,
            "mtime": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
        })
    return result


def format_size(size):
    units = ["Б", "КБ", "МБ", "ГБ"]
    value = float(size)
    for unit in units:
        if value < 1024 or unit == units[-1]:
            return f"{value:.1f} {unit}"
        value /= 1024


def log_path(run_id):
    return LOGS_DIR / f"{run_id}.log"


def write_log(path, text):
    with path.open("a", encoding="utf-8", buffering=1) as f:
        f.write(text)
        f.flush()


def run_script(run_id, config, folder, stdin_data):
    script_path = SCRIPTS_DIR / config["script"]
    log_file = log_path(run_id)

    env = os.environ.copy()
    # Принудительно оставляем UTF-8 только для обмена сообщениями веб-приложения
    # и дочернего процесса. Обработку данных скриптов это не меняет.
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"

    command = [sys.executable, "-u", str(script_path)]
    started = datetime.now()
    with log_file.open("w", encoding="utf-8", buffering=1) as f:
        f.write(f"[{started:%H:%M:%S}] Запуск: {config['title']}\n")
        f.write(f"Рабочая папка: {folder}\n")
        f.write(f"Скрипт: {script_path}\n")
        f.write("-" * 70 + "\n")

    try:
        proc = subprocess.Popen(
            command,
            cwd=str(folder),
            stdin=subprocess.PIPE if stdin_data is not None else None,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
            bufsize=1,
        )
        with RUNS_LOCK:
            RUNS[run_id]["pid"] = proc.pid
            RUNS[run_id]["process"] = proc

        if stdin_data is not None:
            try:
                proc.stdin.write(stdin_data)
                proc.stdin.close()
            except Exception as exc:
                write_log(log_file, f"\n[Ошибка ввода] {exc}\n")

        for line in proc.stdout:
            write_log(log_file, line)
        proc.wait()
        returncode = proc.returncode
        write_log(log_file, "-" * 70 + f"\n[{datetime.now():%H:%M:%S}] Завершено. Код: {returncode}\n")
    except Exception as exc:
        returncode = -1
        write_log(log_file, f"\n[ОШИБКА ЗАПУСКА] {exc}\n")
    finally:
        with RUNS_LOCK:
            RUNS[run_id]["returncode"] = returncode
            RUNS[run_id]["finished"] = datetime.now().isoformat()
            RUNS[run_id].pop("process", None)


def get_operation(key):
    if key not in SCRIPTS:
        abort(404)
    return SCRIPTS[key]


def safe_filename(name):
    # Сохраняем русские и другие Unicode-символы в имени файла.
    # Защищаем только от выхода из рабочей папки.
    name = str(name).replace("\\", "/").split("/")[-1].strip()
    if not name or name in {".", ".."} or any(ord(ch) < 32 for ch in name):
        abort(400)
    return name


@app.context_processor
def inject_helpers():
    return {"format_size": format_size}


def previous_month_range(today=None):
    today = today or date.today()
    first_this_month = today.replace(day=1)
    last_previous_month = first_this_month - timedelta(days=1)
    first_previous_month = last_previous_month.replace(day=1)
    return first_previous_month.isoformat(), last_previous_month.isoformat(), first_previous_month.strftime("%Y-%m")


def statistics_period(folder_key):
    """Определяет фактический диапазон дат по доступным Excel-файлам, не изменяя их."""
    folder = FOLDERS[folder_key]
    date_columns = {
        "начало счёта", "начало счета", "дата начала", "дата", "начало", "дата/время начала",
        "start date", "start", "date"
    }
    found_dates = []
    files_count = 0
    for path in folder.glob("*.xlsx"):
        try:
            frame = pd.read_excel(path, sheet_name=0, nrows=100000)
        except Exception:
            continue
        files_count += 1
        for column in frame.columns:
            normalized = str(column).strip().lower()
            if normalized in date_columns or ("начало" in normalized and ("счёт" in normalized or "счет" in normalized or "дата" in normalized)):
                values = pd.to_datetime(frame[column], errors="coerce", dayfirst=False).dropna()
                if not values.empty:
                    found_dates.extend([values.min(), values.max()])
        # Если нашли период по файлу, не перебираем дополнительные колонки в нём.
    if not found_dates:
        return {"label": "Период пока не определён", "files_count": files_count, "has_period": False}
    start, end = min(found_dates), max(found_dates)
    return {
        "label": f"{start:%d.%m.%Y} — {end:%d.%m.%Y}",
        "files_count": files_count,
        "has_period": True,
    }


@app.route("/")
def index():
    periods = {
        "vckp": statistics_period("connection"),
        "csm": statistics_period("csm"),
    }
    month_start, month_end, month_label = previous_month_range()
    return render_template("index.html", scripts=SCRIPTS, periods=periods,
                           month_start=month_start, month_end=month_end, month_label=month_label)


@app.route("/script/<key>", methods=["GET", "POST"])
def script_page(key):
    config = get_operation(key)
    folder_key = config["folder_key"]
    folder = FOLDERS[folder_key]

    if request.method == "POST":
        form = {k: request.form.get(k, "").strip() for k in request.form}

        required = []
        if key == "csm":
            required = ["input_file"]
        elif key == "delete_prints":
            required = ["input_file"]
        elif key == "report_by_date":
            required = ["groups_file", "results_file", "output_name", "period_start", "period_end"]

        missing = [field for field in required if not form.get(field)]
        if missing:
            flash("Заполните все обязательные поля.", "error")
            return redirect(url_for("script_page", key=key))

        run_id = uuid.uuid4().hex[:12]
        run = {
            "id": run_id,
            "key": key,
            "title": config["title"],
            "folder_key": folder_key,
            "started": datetime.now().isoformat(),
            "finished": None,
            "returncode": None,
        }
        with RUNS_LOCK:
            RUNS[run_id] = run

        stdin_data = config["stdin"](form)
        thread = threading.Thread(target=run_script, args=(run_id, config, folder, stdin_data), daemon=True)
        thread.start()
        return redirect(url_for("run_view", run_id=run_id))

    files = list_files(folder_key)
    if key in {"csm", "delete_prints", "report_by_date"}:
        extensions = {".txt", ".xlsx"}
        selectable = [f for f in files if Path(f["name"]).suffix.lower() in extensions]
    else:
        selectable = []
    return render_template("script.html", key=key, info=config, files=files, selectable=selectable)


def _workflow_log(log_file, message):
    write_log(log_file, message.rstrip() + "\n")


def _run_step(script_name, workdir, stdin_text, log_file, title):
    script_path = SCRIPTS_DIR / script_name
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    _workflow_log(log_file, f"\n▶ {title}\nКоманда: {script_path}\nРабочая папка: {workdir}\n")
    proc = subprocess.Popen(
        [sys.executable, "-X", "utf8", "-u", str(script_path)],
        cwd=str(workdir), stdin=subprocess.PIPE if stdin_text is not None else None,
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
        encoding="utf-8", errors="replace", env=env, bufsize=1,
    )
    if stdin_text is not None:
        try:
            proc.stdin.write(stdin_text)
            proc.stdin.close()
        except Exception as exc:
            _workflow_log(log_file, f"[Ошибка передачи ввода] {exc}")
    for line in proc.stdout:
        _workflow_log(log_file, line.rstrip("\n"))
    proc.wait()
    _workflow_log(log_file, f"Код завершения шага: {proc.returncode}")
    if proc.returncode != 0:
        raise RuntimeError(f"Шаг «{title}» завершился с кодом {proc.returncode}. Проверьте журнал запуска.")


def run_monthly_workflow(run_id, source, selected_name, period_start, period_end, month_label):
    """Оркестрация существующих скриптов. Сами скрипты обработки не меняются."""
    log_file = log_path(run_id)
    report_folder = FOLDERS["report"]
    temp_root = DATA_DIR / "outputs"
    temp_root.mkdir(parents=True, exist_ok=True)
    workdir = Path(tempfile.mkdtemp(prefix=f"monthly_{run_id}_", dir=str(temp_root)))
    returncode = -1
    try:
        write_log(log_file, f"Запуск быстрого отчёта за {month_label}\nИсточник: {source}\n")
        group_file = report_folder / "group_file_origin.xlsx"
        if not group_file.is_file():
            raise FileNotFoundError("Не найден data/report/group_file_origin.xlsx. Сначала добавьте файл групп.")

        if source == "vckp":
            source_path = FOLDERS["word"] / Path(selected_name).name
            if not source_path.is_file() or source_path.suffix.lower() != ".docx":
                raise FileNotFoundError("Выбранный файл ВЦКП не найден в папке data/word.")
            shutil.copy2(source_path, workdir / source_path.name)
            _run_step("script_conferter_word_excel.py", workdir, None, log_file, "1/3 · Word → Excel")
            produced_xlsx = [p for p in workdir.glob("*.xlsx") if p.name != group_file.name]
            if not produced_xlsx:
                raise RuntimeError("Скрипт Word → Excel не создал Excel-файл. Проверьте выбранный DOCX и журнал.")
            _run_step("script_connection_excel.py", workdir, None, log_file, "2/3 · Объединение Excel")
            aggregates = [p for p in workdir.glob("statistic_*.xlsx") if p.name != group_file.name]
            if not aggregates:
                raise RuntimeError("Скрипт объединения не создал statistic_*.xlsx.")
            statistics_file = max(aggregates, key=lambda p: p.stat().st_mtime)
            output_stem = f"Отчёт_ВЦКП_{month_label}"
            source_title = "ВЦКП"
        elif source == "csm":
            source_path = FOLDERS["csm"] / Path(selected_name).name
            if not source_path.is_file() or source_path.suffix.lower() != ".txt":
                raise FileNotFoundError("Выбранный TXT-файл ЦСМ не найден в папке data/csm.")
            shutil.copy2(source_path, workdir / source_path.name)
            csm_output_stem = f"csm_statistic_{month_label}"
            _run_step("script_convertor_csm.py", workdir,
                      f"{source_path.stem}\n{csm_output_stem}\n", log_file, "1/2 · ЦСМ TXT → Excel")
            statistics_file = workdir / f"{csm_output_stem}.xlsx"
            if not statistics_file.is_file():
                # Скрипт может не создать файл, если не нашёл записей.
                raise RuntimeError("Скрипт ЦСМ не создал Excel-файл. Проверьте исходный TXT и журнал.")
            output_stem = f"Отчёт_ЦСМ_{month_label}"
            source_title = "ЦСМ"
        else:
            raise ValueError("Неизвестный источник статистики.")

        # Файл групп добавляем только перед отчётом, чтобы скрипт объединения Excel
        # не включил его в статистику ВЦКП.
        shutil.copy2(group_file, workdir / group_file.name)
        _run_step(
            "script_statistic_to_date_filter.py", workdir,
            f"нет\n{group_file.name}\n{statistics_file.name}\n{output_stem}\n{period_start}\n{period_end}\n",
            log_file, ("3/3" if source == "vckp" else "2/2") + " · Отчёт по группам и датам",
        )
        report_path = workdir / f"{output_stem}.xlsx"
        if not report_path.is_file():
            candidates = [p for p in workdir.glob("*.xlsx") if p.name.startswith(output_stem)]
            if not candidates:
                raise RuntimeError("Отчёт не найден после завершения скрипта фильтрации.")
            report_path = max(candidates, key=lambda p: p.stat().st_mtime)
        final_path = report_folder / report_path.name
        # Не перезаписываем старый отчёт без предупреждения: имя уникализируется временем.
        if final_path.exists():
            final_path = report_folder / f"{report_path.stem}_{datetime.now():%H%M%S}{report_path.suffix}"
        shutil.copy2(report_path, final_path)
        _workflow_log(log_file, f"\n✓ Готово. Итоговый отчёт: {final_path.name}")
        returncode = 0
    except Exception as exc:
        _workflow_log(log_file, f"\n[ОШИБКА] {exc}")
    finally:
        with RUNS_LOCK:
            RUNS[run_id]["returncode"] = returncode
            RUNS[run_id]["finished"] = datetime.now().isoformat()
        # Оставляем временную папку при ошибке для диагностики; удаляем после успеха.
        if returncode == 0:
            shutil.rmtree(workdir, ignore_errors=True)


@app.route("/monthly", methods=["GET", "POST"])
def monthly_report():
    month_start, month_end, month_label = previous_month_range()
    if request.method == "POST":
        source = request.form.get("source", "").strip()
        selected_name = request.form.get("source_file", "").strip()
        if source not in {"vckp", "csm"}:
            flash("Выберите источник статистики.", "error")
            return redirect(url_for("monthly_report"))
        if not selected_name:
            flash("Выберите исходный файл.", "error")
            return redirect(url_for("monthly_report"))
        if not (FOLDERS["report"] / "group_file_origin.xlsx").is_file():
            flash("Не найден файл групп. Добавьте group_file_origin.xlsx в раздел «Отчёт». ", "error")
            return redirect(url_for("monthly_report"))
        run_id = uuid.uuid4().hex[:12]
        with RUNS_LOCK:
            RUNS[run_id] = {
                "id": run_id, "key": "monthly", "title": f"Отчёт {source.upper()} за {month_label}",
                "folder_key": "report", "started": datetime.now().isoformat(),
                "finished": None, "returncode": None,
            }
        thread = threading.Thread(
            target=run_monthly_workflow,
            args=(run_id, source, selected_name, month_start, month_end, month_label), daemon=True,
        )
        thread.start()
        return redirect(url_for("run_view", run_id=run_id))

    word_files = [f for f in list_files("word") if Path(f["name"]).suffix.lower() == ".docx"]
    csm_files = [f for f in list_files("csm") if Path(f["name"]).suffix.lower() == ".txt"]
    return render_template("monthly.html", word_files=word_files, csm_files=csm_files,
                           month_start=month_start, month_end=month_end, month_label=month_label,
                           has_groups=(FOLDERS["report"] / "group_file_origin.xlsx").is_file())


@app.route("/run/<run_id>")
def run_view(run_id):
    with RUNS_LOCK:
        run = RUNS.get(run_id)
    if not run:
        abort(404)
    return render_template("run.html", run=run)


@app.route("/api/run/<run_id>")
def api_run(run_id):
    with RUNS_LOCK:
        run = RUNS.get(run_id)
    if not run:
        abort(404)

    path = log_path(run_id)
    log = path.read_text(encoding="utf-8", errors="replace") if path.exists() else ""
    folder = FOLDERS[run["folder_key"]]
    files = []
    if run.get("finished"):
        files = list_files(run["folder_key"])

    return jsonify({
        "finished": bool(run.get("finished")),
        "returncode": run.get("returncode"),
        "log": log,
        "files": files,
    })


@app.route("/files/<folder_key>")
def files_page(folder_key):
    if folder_key not in FOLDERS:
        abort(404)
    return render_template("files.html", folder_key=folder_key, files=list_files(folder_key))


@app.route("/upload/<folder_key>", methods=["POST"])
def upload(folder_key):
    if folder_key not in FOLDERS:
        abort(404)
    uploaded = request.files.getlist("file")
    saved = 0
    for file in uploaded:
        if not file or not file.filename:
            continue
        name = safe_filename(file.filename)
        file.save(FOLDERS[folder_key] / name)
        saved += 1
    flash(f"Загружено файлов: {saved}", "ok")
    return redirect(url_for("files_page", folder_key=folder_key))


@app.route("/download/<folder_key>/<path:filename>")
def download(folder_key, filename):
    if folder_key not in FOLDERS:
        abort(404)
    filename = Path(filename).name
    path = FOLDERS[folder_key] / filename
    if not path.is_file():
        abort(404)
    # Flask до 2.0 использует attachment_filename, новые версии — download_name.
    # Пробуем новый параметр и делаем совместимый fallback для старых установок.
    try:
        return send_from_directory(str(FOLDERS[folder_key]), filename,
                                   as_attachment=True, download_name=filename)
    except TypeError as exc:
        if "download_name" not in str(exc):
            raise
        return send_from_directory(str(FOLDERS[folder_key]), filename,
                                   as_attachment=True, attachment_filename=filename)


@app.route("/delete/<folder_key>/<path:filename>", methods=["POST"])
def delete_file(folder_key, filename):
    if folder_key not in FOLDERS:
        abort(404)
    path = FOLDERS[folder_key] / Path(filename).name
    if path.is_file():
        path.unlink()
        flash(f"Удалён файл: {path.name}", "ok")
    return redirect(url_for("files_page", folder_key=folder_key))


@app.route("/groups", methods=["GET", "POST"])
def groups_editor():
    path = FOLDERS["report"] / "group_file_origin.xlsx"
    if request.method == "POST":
        try:
            rows = json.loads(request.form.get("data", "[]"))
            df = pd.DataFrame(rows)
            columns = ["Наименование группы", "Участники", "Ресурс", "Коментарий"]
            for col in columns:
                if col not in df.columns:
                    df[col] = ""
            df = df[columns]
            df.to_excel(path, index=False)
            flash("Файл групп сохранён.", "ok")
        except Exception as exc:
            flash(f"Не удалось сохранить файл групп: {exc}", "error")
        return redirect(url_for("groups_editor"))

    if path.exists():
        df = pd.read_excel(path).fillna("")
        rows = df.astype(str).to_dict(orient="records")
    else:
        rows = []
    return render_template("groups.html", rows=rows)


@app.errorhandler(413)
def too_large(_):
    return render_template("error.html", message="Файл слишком большой. Максимальный размер — 200 МБ."), 413


@app.errorhandler(404)
def not_found(_):
    return render_template("error.html", message="Страница или файл не найдены."), 404


if __name__ == "__main__":
    print("Веб-интерфейс статистики: http://127.0.0.1:5000")
    app.run(host="127.0.0.1", port=5000, debug=False, threaded=True)
