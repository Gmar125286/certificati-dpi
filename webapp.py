from __future__ import annotations

import json
import os
import secrets
from functools import wraps
from pathlib import Path

from flask import (
    Flask,
    Response,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)

from app import (
    ASSETS_OUTPUT_DIR,
    DEFAULT_ADMIN_USERNAME,
    DB_PATH,
    DOCX_OUTPUT_DIR,
    ICON_PATH,
    IRUDEK_NORMS_PATH,
    LOG_PATH,
    OUTPUT_DIR,
    PDF_OUTPUT_DIR,
    SETTINGS_PATH,
    USERS_PATH,
    CertificateGenerator,
    IrudekNormCatalog,
    RevisionRecord,
    RevisionRepository,
    UserStore,
    build_irudek_norm_catalog,
    current_datetime_display,
    date_only_display,
    default_settings,
    extract_checklist_items,
    normalize_catalog_code,
    normalize_username,
    template_files,
)

app = Flask(__name__)

repository = RevisionRepository(DB_PATH)
user_store = UserStore(USERS_PATH)
norm_catalog = IrudekNormCatalog(IRUDEK_NORMS_PATH)
generator = CertificateGenerator()


def template_lookup() -> dict[str, Path]:
    return {path.name: path for path in template_files()}


def load_settings() -> dict:
    if SETTINGS_PATH.exists():
        try:
            return default_settings() | json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
        except Exception:
            return default_settings()
    return default_settings()


def save_settings(payload: dict) -> None:
    SETTINGS_PATH.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def ensure_web_secret() -> str:
    settings = load_settings()
    secret = str(settings.get("web_secret_key", "")).strip()
    if secret:
        return secret
    secret = secrets.token_hex(32)
    settings["web_secret_key"] = secret
    save_settings(settings)
    return secret


app.secret_key = ensure_web_secret()


def current_user() -> str:
    return str(session.get("username", "")).strip()


def is_admin() -> bool:
    return normalize_username(current_user()) == DEFAULT_ADMIN_USERNAME


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not current_user():
            return redirect(url_for("login"))
        return view(*args, **kwargs)

    return wrapped


def admin_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not current_user():
            return redirect(url_for("login"))
        if not is_admin():
            flash("Solo admin puo' creare utenti.", "error")
            return redirect(url_for("dashboard"))
        return view(*args, **kwargs)

    return wrapped


def write_error_log(message: str) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    with LOG_PATH.open("a", encoding="utf-8") as handle:
        from datetime import datetime

        handle.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")


def form_defaults() -> dict[str, str | bool]:
    templates = template_files()
    first_template = templates[0].name if templates else ""
    return {
        "cliente": "",
        "template_name": first_template,
        "modello_dispositivo": first_template.replace("CERTIFICATO DI REVISIONE - ", "").replace(".docx", "").strip()
        if first_template
        else "",
        "norma": "EN 365",
        "codice_sap": "",
        "lotto": "",
        "serie": "",
        "anno_fabbricazione": "",
        "data_revisione": current_datetime_display(),
        "data_prossima_revisione": "",
        "tecnico": "",
        "osservazioni": "",
        "esito": "Positivo",
        "export_pdf": True,
    }


def form_from_request() -> dict[str, str | bool]:
    payload = {
        "cliente": request.form.get("cliente", "").strip(),
        "template_name": request.form.get("template_name", "").strip(),
        "modello_dispositivo": request.form.get("modello_dispositivo", "").strip(),
        "norma": request.form.get("norma", "").strip(),
        "codice_sap": request.form.get("codice_sap", "").strip(),
        "lotto": request.form.get("lotto", "").strip(),
        "serie": request.form.get("serie", "").strip(),
        "anno_fabbricazione": request.form.get("anno_fabbricazione", "").strip(),
        "data_revisione": request.form.get("data_revisione", "").strip(),
        "data_prossima_revisione": request.form.get("data_prossima_revisione", "").strip(),
        "tecnico": request.form.get("tecnico", "").strip(),
        "osservazioni": request.form.get("osservazioni", "").strip(),
        "esito": request.form.get("esito", "Positivo").strip() or "Positivo",
        "export_pdf": request.form.get("export_pdf") == "on",
    }
    if payload["template_name"] and not payload["modello_dispositivo"]:
        payload["modello_dispositivo"] = (
            str(payload["template_name"]).replace("CERTIFICATO DI REVISIONE - ", "").replace(".docx", "").strip()
        )
    return payload


def build_record(payload: dict[str, str | bool]) -> RevisionRecord:
    return RevisionRecord(
        cliente=str(payload["cliente"]),
        template_name=str(payload["template_name"]),
        modello_dispositivo=str(payload["modello_dispositivo"]),
        norma=str(payload["norma"]),
        codice_sap=str(payload["codice_sap"]),
        lotto=str(payload["lotto"]),
        serie=str(payload["serie"]),
        anno_fabbricazione=str(payload["anno_fabbricazione"]),
        data_revisione=str(payload["data_revisione"]),
        data_prossima_revisione=str(payload["data_prossima_revisione"]),
        tecnico=str(payload["tecnico"]),
        osservazioni=str(payload["osservazioni"]),
        esito=str(payload["esito"]),
    )


def validate_record(record: RevisionRecord) -> list[str]:
    missing = []
    for label, value in [
        ("Modello certificato", record.template_name),
        ("Cliente", record.cliente),
        ("Modello dispositivo", record.modello_dispositivo),
        ("Data revisione", record.data_revisione),
        ("Tecnico", record.tecnico),
    ]:
        if not value:
            missing.append(label)
    return missing


def available_checklist(template_name: str, esito: str) -> tuple[list[str], list[str]]:
    lookup = template_lookup()
    template_path = lookup.get(template_name)
    if not template_path:
        return [], []
    items = extract_checklist_items(template_path)
    default_value = "Positivo" if esito != "Negativo" else "Negativo"
    return items, [default_value for _ in items]


def checklist_from_request(template_name: str, esito: str) -> list[str]:
    items, defaults = available_checklist(template_name, esito)
    values: list[str] = []
    for idx, default in enumerate(defaults):
        value = request.form.get(f"check_{idx}", default)
        values.append("Negativo" if value == "Negativo" else "Positivo")
    return values


def record_to_form(row) -> dict[str, str | bool]:
    return {
        "cliente": row["cliente"] or "",
        "template_name": row["template_name"] or "",
        "modello_dispositivo": row["modello_dispositivo"] or "",
        "norma": row["norma"] or "",
        "codice_sap": row["codice_sap"] or "",
        "lotto": row["lotto"] or "",
        "serie": row["serie"] or "",
        "anno_fabbricazione": row["anno_fabbricazione"] or "",
        "data_revisione": row["data_revisione"] or "",
        "data_prossima_revisione": row["data_prossima_revisione"] or "",
        "tecnico": row["tecnico"] or "",
        "osservazioni": row["osservazioni"] or "",
        "esito": row["esito"] or "Positivo",
        "export_pdf": bool(row["pdf_path"]),
    }


def build_history(search_query: str) -> list[dict]:
    rows = []
    for row in repository.search_records(search_query):
        rows.append(
            {
                "id": row["id"],
                "created_at": row["created_at"],
                "updated_at": row["updated_at"] or "",
                "cliente": row["cliente"],
                "modello": row["modello_dispositivo"],
                "seriale": row["serie"],
                "data_revisione": row["data_revisione"],
                "esito": row["esito"],
                "has_docx": bool(row["docx_path"]),
                "has_pdf": bool(row["pdf_path"]),
            }
        )
    return rows


def resolve_file_for_record(record_id: int, kind: str) -> Path | None:
    row = repository.get_record_by_id(record_id)
    if not row:
        return None
    raw_path = row["pdf_path"] if kind == "pdf" else row["docx_path"]
    if not raw_path:
        return None
    path = Path(raw_path)
    return path if path.exists() else None


@app.context_processor
def inject_globals():
    settings = load_settings()
    return {
        "app_title": "Gestione Certificati DPI Web",
        "current_user": current_user(),
        "is_admin": is_admin(),
        "logo_path": settings.get("logo_path", ""),
    }


@app.route("/login", methods=["GET", "POST"])
def login():
    if current_user():
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        if user_store.verify_credentials(username, password):
            session["username"] = username
            flash("Accesso eseguito correttamente.", "success")
            return redirect(url_for("dashboard"))
        flash("Credenziali non valide.", "error")
    return render_template("login.html")


@app.route("/logout", methods=["POST"])
@login_required
def logout():
    session.clear()
    flash("Sessione chiusa.", "success")
    return redirect(url_for("login"))


@app.route("/", methods=["GET"])
@login_required
def dashboard():
    settings = load_settings()
    search_query = request.args.get("q", "").strip()
    edit_id = request.args.get("edit", "").strip()
    form_data = form_defaults()
    if edit_id.isdigit():
        row = repository.get_record_by_id(int(edit_id))
        if row:
            form_data = record_to_form(row)
    templates = [path.name for path in template_files()]
    checklist_items, checklist_values = available_checklist(
        str(form_data["template_name"]), str(form_data["esito"])
    )
    recent_clients = repository.distinct_clients()
    return render_template(
        "dashboard.html",
        templates=templates,
        settings=settings,
        form_data=form_data,
        edit_id=edit_id if edit_id.isdigit() else "",
        history_rows=build_history(search_query),
        search_query=search_query,
        recent_clients=recent_clients,
        checklist_items=checklist_items,
        checklist_values=checklist_values,
        catalog_generated_at=norm_catalog.generated_at,
        catalog_codes=len(norm_catalog.items),
    )


@app.route("/save", methods=["POST"])
@login_required
def save_revision():
    settings = load_settings()
    payload = form_from_request()
    edit_id_raw = request.form.get("edit_id", "").strip()
    action = request.form.get("action_name", "generate")
    record = build_record(payload)
    missing = validate_record(record)
    if missing:
        flash("Compila questi campi: " + ", ".join(missing), "error")
        return redirect(url_for("dashboard", edit=edit_id_raw or None))

    lookup = template_lookup()
    template_path = lookup.get(record.template_name)
    if not template_path:
        flash("Il modello selezionato non e disponibile.", "error")
        return redirect(url_for("dashboard", edit=edit_id_raw or None))

    logo_value = str(settings.get("logo_path", "")).strip()
    logo_path = Path(logo_value) if logo_value else None
    if logo_path and not logo_path.exists():
        flash("Il logo configurato non esiste piu.", "error")
        return redirect(url_for("dashboard", edit=edit_id_raw or None))

    checklist_results = checklist_from_request(record.template_name, record.esito)
    try:
        docx_path, pdf_path, pdf_error = generator.generate(
            template_path,
            record,
            export_pdf=bool(payload["export_pdf"]),
            logo_path=logo_path,
            checklist_results=checklist_results,
        )
        print_error = None
        if action == "generate_print":
            try:
                generator.print_document(docx_path, copies=2)
            except Exception as exc:
                print_error = str(exc)

        if edit_id_raw.isdigit():
            record_id = int(edit_id_raw)
            previous = repository.get_record_by_id(record_id)
            repository.update_record(record_id, record, docx_path, pdf_path)
            if previous:
                for previous_path in [previous["docx_path"], previous["pdf_path"]]:
                    if not previous_path:
                        continue
                    path = Path(previous_path)
                    if path not in [docx_path, pdf_path] and path.exists():
                        try:
                            path.unlink()
                        except Exception as exc:
                            write_error_log(f"Cleanup warning for {path}: {exc}")
            flash(f"Revisione aggiornata. ID storico {record_id}.", "success")
        else:
            record_id = repository.insert_record(record, docx_path, pdf_path)
            flash(f"Certificato creato. ID storico {record_id}.", "success")

        if pdf_error:
            write_error_log(f"PDF export warning for {docx_path.name}: {pdf_error}")
            flash(f"PDF non creato: {pdf_error}", "warning")
        if print_error:
            write_error_log(f"Print warning for {docx_path.name}: {print_error}")
            flash(f"Stampa non riuscita: {print_error}", "warning")
        elif action == "generate_print":
            flash("Stampa 2 copie inviata.", "success")
    except Exception as exc:
        write_error_log(str(exc))
        flash(f"Generazione non riuscita: {exc}", "error")

    return redirect(url_for("dashboard"))


@app.route("/delete/<int:record_id>", methods=["POST"])
@login_required
def delete_revision(record_id: int):
    row = repository.get_record_by_id(record_id)
    if not row:
        flash("Revisione non trovata.", "error")
        return redirect(url_for("dashboard"))

    errors: list[str] = []
    for key in ["docx_path", "pdf_path"]:
        raw_path = row[key]
        if not raw_path:
            continue
        path = Path(raw_path)
        if not path.exists():
            continue
        try:
            path.unlink()
        except Exception as exc:
            errors.append(f"{path.name}: {exc}")

    if errors:
        flash("Eliminazione non completata: " + " | ".join(errors), "error")
        return redirect(url_for("dashboard"))

    repository.delete_record(record_id)
    flash(f"Revisione {record_id} eliminata.", "success")
    return redirect(url_for("dashboard"))


@app.route("/download/<string:file_kind>/<int:record_id>")
@login_required
def download_record_file(file_kind: str, record_id: int):
    if file_kind not in {"docx", "pdf"}:
        return Response(status=404)
    file_path = resolve_file_for_record(record_id, file_kind)
    if not file_path:
        flash("File non trovato per questa revisione.", "error")
        return redirect(url_for("dashboard"))
    mimetype = "application/pdf" if file_kind == "pdf" else (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    return send_file(file_path, as_attachment=False, download_name=file_path.name, mimetype=mimetype)


@app.route("/api/checklist")
@login_required
def api_checklist():
    template_name = request.args.get("template_name", "").strip()
    esito = request.args.get("esito", "Positivo").strip() or "Positivo"
    items, defaults = available_checklist(template_name, esito)
    return jsonify({"items": items, "defaults": defaults})


@app.route("/api/sap")
@login_required
def api_lookup_sap():
    code = request.args.get("code", "").strip()
    entry = norm_catalog.lookup(code)
    if not entry:
        return jsonify({"found": False})
    return jsonify(
        {
            "found": True,
            "code": normalize_catalog_code(code),
            "standard": str(entry.get("standard", "")).strip(),
            "title": str(entry.get("title", "")).strip(),
        }
    )


@app.route("/api/product-search")
@login_required
def api_product_search():
    query = request.args.get("q", "").strip()
    matches = norm_catalog.search_by_name(query)
    return jsonify(
        {
            "results": [
                {
                    "code": code,
                    "standard": str(entry.get("standard", "")).strip(),
                    "title": str(entry.get("title", "")).strip(),
                }
                for code, entry in matches
            ]
        }
    )


@app.route("/settings/logo", methods=["POST"])
@login_required
def update_logo():
    settings = load_settings()
    logo_path = request.form.get("logo_path", "").strip()
    if logo_path and not Path(logo_path).exists():
        flash("Il percorso logo non esiste.", "error")
        return redirect(url_for("dashboard"))
    settings["logo_path"] = logo_path
    save_settings(settings)
    flash("Logo aggiornato.", "success")
    return redirect(url_for("dashboard"))


@app.route("/admin/users", methods=["POST"])
@admin_required
def create_user():
    username = request.form.get("username", "").strip()
    password = request.form.get("password", "")
    confirm_password = request.form.get("confirm_password", "")
    if password != confirm_password:
        flash("Le password non coincidono.", "error")
        return redirect(url_for("dashboard"))
    try:
        created = user_store.create_user(username, password)
        flash(f"Utente creato: {created}", "success")
    except ValueError as exc:
        flash(str(exc), "error")
    return redirect(url_for("dashboard"))


@app.route("/catalog/update", methods=["POST"])
@login_required
def update_catalog():
    try:
        total_codes, conflicts = build_irudek_norm_catalog(IRUDEK_NORMS_PATH)
        norm_catalog.load()
        flash(
            f"Catalogo aggiornato. Codici SAP: {total_codes}. Conflitti ignorati: {conflicts}.",
            "success",
        )
    except Exception as exc:
        write_error_log(str(exc))
        flash(f"Aggiornamento catalogo non riuscito: {exc}", "error")
    return redirect(url_for("dashboard"))


@app.route("/favicon.ico")
def favicon():
    if ICON_PATH.exists():
        return send_file(ICON_PATH)
    return Response(status=404)


@app.route("/healthz")
def healthz():
    return jsonify({"status": "ok"})


def ensure_runtime_dirs() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    DOCX_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    PDF_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    ASSETS_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


if __name__ == "__main__":
    ensure_runtime_dirs()
    host = os.getenv("HOST", "0.0.0.0")
    port = int(os.getenv("PORT", "3000"))
    app.run(host=host, port=port, debug=False)
