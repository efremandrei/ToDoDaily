"""
Daily To-Do Tracker Web Application.

This Flask application provides:
- A very simple "login" via username selection (no passwords).
- Per-user, per-weekday configuration of to-do list presets.
- A daily view that shows today's list, clickable to cross / un-cross items.
- State is remembered per day, and a new day's state starts uncrossed.
- Admin user management and Excel export (per user and for all users).
- Optional email reminders if all today's items are still un-crossed
  at a configurable time.

The app is designed for use on a smartphone:
- Responsive layout via Bootstrap.
- Large touch-friendly buttons.
"""

from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    session,
    flash,
    send_file,
)
from flask_sqlalchemy import SQLAlchemy
from functools import wraps
import os
import threading
import time
from datetime import datetime, date
from io import BytesIO
import smtplib
from email.mime.text import MIMEText
from openpyxl import Workbook

# ---------------------------------------------------------------------------
# Flask and database setup
# ---------------------------------------------------------------------------

app = Flask(__name__)

# Secret key used for sessions; override in production via environment variable.
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-secret-key")

# SQLite database in the container / local folder.
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///todo.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

# SQLAlchemy database instance.
db = SQLAlchemy(app)


# ---------------------------------------------------------------------------
# Database models
# ---------------------------------------------------------------------------

class User(db.Model):
    """Represents a logical user in the system.

    Each user has:
    - A unique name used for "login" selection.
    - An email for optional notifications.
    - A flag that marks the first user as admin.
    - Flags that control if and when reminder emails are sent.
    """
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(80), unique=True, nullable=False)
    email = db.Column(db.String(120), nullable=True)
    is_admin = db.Column(db.Boolean, default=False)

    # True if we should send email reminders to this user.
    email_notifications_enabled = db.Column(db.Boolean, default=False)

    # Time of day to send reminder emails in "HH:MM" 24h format (e.g. "23:59").
    email_time = db.Column(db.String(5), default="23:59")


class TodoTemplate(db.Model):
    """Represents a 'preset' task line for a specific weekday and user.

    This is the configuration that defines what should appear for each weekday.
    Weekday index:
        0 = Sunday
        1 = Monday
        2 = Tuesday
        3 = Wednesday
        4 = Thursday
        5 = Friday
        6 = Saturday
    """
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    weekday = db.Column(db.Integer, nullable=False)  # 0=Sunday, 5=Friday
    order_index = db.Column(db.Integer, nullable=False, default=0)
    text = db.Column(db.String(255), nullable=False)


class TodoState(db.Model):
    """Represents the daily state of a todo item for a specific date and user.

    The items are created from TodoTemplate when the user first opens
    their list for that day. Each TodoState can be toggled crossed / uncrossed.
    """
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)

    # Date of this instance, in "YYYY-MM-DD".
    date = db.Column(db.String(10), nullable=False)

    # Weekday index (0=Sunday ... 6=Saturday), stored for convenience.
    weekday = db.Column(db.Integer, nullable=False)

    # Used to keep items in the same order as in the template.
    order_index = db.Column(db.Integer, nullable=False, default=0)

    # Snapshot of the text from the template at the time of creation.
    text = db.Column(db.String(255), nullable=False)

    # True if the item is crossed / completed.
    checked = db.Column(db.Boolean, default=False)


class EmailLog(db.Model):
    """Stores information about reminder emails that were already sent.

    Used to avoid sending duplicate emails on the same day.
    """
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    date = db.Column(db.String(10), nullable=False)  # "YYYY-MM-DD"
    created_at = db.Column(db.DateTime, default=datetime.utcnow)


# ---------------------------------------------------------------------------
# Initialization helpers
# ---------------------------------------------------------------------------

def init_db():
    """Create all database tables and ensure an admin user exists.

    The first time the application runs, it will create an 'Admin' user
    automatically so the system always has at least one administrator.
    """
    db.create_all()

    # If the database is empty, seed it with a default admin user.
    if User.query.count() == 0:
        admin = User(name="Admin", email="", is_admin=True)
        db.session.add(admin)
        db.session.commit()


def get_current_user():
    """Return the currently logged-in user object, or None if not logged in."""
    user_id = session.get("user_id")
    if user_id is None:
        return None
    return User.query.get(user_id)


@app.context_processor
def inject_current_user():
    """Make the current user available as 'current_user' in all templates."""
    return {"current_user": get_current_user()}


# ---------------------------------------------------------------------------
# Auth decorators
# ---------------------------------------------------------------------------

def login_required(f):
    """Decorator to enforce that a user must be logged in.

    If no user is in the session, the caller is redirected to the login page.
    """
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if get_current_user() is None:
            flash("Please choose a user to continue.", "warning")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function


def admin_required(f):
    """Decorator to enforce that the current user is an admin."""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        user = get_current_user()
        if user is None or not user.is_admin:
            flash("Admin access required.", "danger")
            return redirect(url_for("today"))
        return f(*args, **kwargs)
    return decorated_function


# ---------------------------------------------------------------------------
# Weekday utilities
# ---------------------------------------------------------------------------

def get_weekday_index(d: date) -> int:
    """Return custom weekday index for a date where Sunday=0, Monday=1, ..., Friday=5.

    Python's datetime.weekday() returns Monday=0 ... Sunday=6.
    We rotate this so that Sunday becomes 0:

        python_weekday = d.weekday()  # Monday=0 ... Sunday=6
        sunday_based = (python_weekday + 1) % 7

    This function returns that "sunday_based" index.
    """
    return (d.weekday() + 1) % 7


def get_weekday_name(index: int) -> str:
    """Map our custom weekday index to a human-friendly name."""
    names = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    return names[index]


# ---------------------------------------------------------------------------
# Daily state management
# ---------------------------------------------------------------------------

def ensure_states_for_today(user: User):
    """Ensure that TodoState entries exist for the current day for a user.

    If they don't exist yet, the function creates them from the user's templates
    for the corresponding weekday. It returns a list of TodoState objects
    ordered by their order_index.
    """
    today = date.today()
    weekday_index = get_weekday_index(today)

    # Only Sunday (0) through Saturday (6) are considered valid working days.
    if weekday_index > 6:
        return []

    date_str = today.isoformat()

    # Check if we already created the TodoState rows for this date.
    states = (
        TodoState.query
        .filter_by(user_id=user.id, date=date_str)
        .order_by(TodoState.order_index)
        .all()
    )
    if states:
        return states

    # If no state rows exist yet for today, create them from templates.
    templates = (
        TodoTemplate.query
        .filter_by(user_id=user.id, weekday=weekday_index)
        .order_by(TodoTemplate.order_index)
        .all()
    )

    for tmpl in templates:
        state = TodoState(
            user_id=user.id,
            date=date_str,
            weekday=weekday_index,
            order_index=tmpl.order_index,
            text=tmpl.text,
            checked=False,
        )
        db.session.add(state)
    db.session.commit()

    # Return the freshly created states.
    states = (
        TodoState.query
        .filter_by(user_id=user.id, date=date_str)
        .order_by(TodoState.order_index)
        .all()
    )
    return states


# ---------------------------------------------------------------------------
# Email sending + reminder logic
# ---------------------------------------------------------------------------

def send_email(to_address: str, subject: str, body: str):
    """Send a plain-text email using SMTP configuration from environment variables.

    Required environment variables:
        SMTP_SERVER
        SMTP_PORT
        SMTP_USERNAME
        SMTP_PASSWORD
        EMAIL_FROM

    If something fails, the exception is printed and swallowed so the
    background thread does not crash.
    """
    smtp_server = os.environ.get("SMTP_SERVER")
    smtp_port = int(os.environ.get("SMTP_PORT", "587"))
    smtp_username = os.environ.get("SMTP_USERNAME")
    smtp_password = os.environ.get("SMTP_PASSWORD")
    email_from = os.environ.get("EMAIL_FROM")

    # If configuration is incomplete, we silently skip sending emails.
    if not all([smtp_server, smtp_username, smtp_password, email_from]):
        print("Email configuration incomplete; skipping email send.")
        return

    message = MIMEText(body)
    message["Subject"] = subject
    message["From"] = email_from
    message["To"] = to_address

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_username, smtp_password)
            server.send_message(message)
        print(f"Sent reminder email to {to_address}")
    except Exception as exc:
        # We do not raise here to avoid killing the background thread.
        print(f"Error sending email to {to_address}: {exc}")


def check_and_send_reminders():
    """Check all users and send reminder emails if conditions are met.

    A reminder is sent if:
    - The user has email notifications enabled.
    - The current local time matches the user's configured time (HH:MM).
    - There is a list for today (Sunday-Saturday).
    - All items for today are unchecked.
    - No email was already sent today.
    """
    now = datetime.now()
    today = now.date()
    weekday_index = get_weekday_index(today)

    # No reminders on Saturday by definition of this app.
    if weekday_index > 5:
        return

    current_hhmm = now.strftime("%H:%M")
    date_str = today.isoformat()

    users = User.query.all()
    for user in users:
        # Skip users without notifications enabled or without email.
        if not user.email_notifications_enabled:
            continue
        if not user.email:
            continue

        # Only act at the exact configured time (HH:MM).
        if user.email_time != current_hhmm:
            continue

        # Check if we already sent an email to this user today.
        already_sent = EmailLog.query.filter_by(user_id=user.id, date=date_str).first()
        if already_sent:
            continue

        # Load today's state. If there is no state, we do not send anything.
        states = TodoState.query.filter_by(user_id=user.id, date=date_str).all()
        if not states:
            continue

        # Check if all items are unchecked.
        any_checked = any(state.checked for state in states)
        if any_checked:
            continue

        # All items are unchecked => send an email reminder.
        subject = f"Reminder - Your {get_weekday_name(weekday_index)} tasks are still open"
        body_lines = [
            f"Hello {user.name},",
            "",
            f"It is {now.strftime('%Y-%m-%d %H:%M')} and all your tasks for today "
            f"({get_weekday_name(weekday_index)}) are still unchecked.",
            "This is a friendly reminder to review your to-do list.",
            "",
            "Best regards,",
            "Your Daily To-Do Tracker",
        ]
        body = "\n".join(body_lines)
        send_email(user.email, subject, body)

        # Record that we sent an email today for this user.
        log = EmailLog(user_id=user.id, date=date_str)
        db.session.add(log)
        db.session.commit()


# Flag used to ensure we only start one background thread per process.
email_thread_started = False


def start_email_thread():
    """Start a background thread that periodically checks and sends reminders.

    The thread runs an infinite loop, sleeping 60 seconds between checks.

    NOTE:
        This simple approach is fine for a single-process deployment.
        In a multi-process setup (e.g. gunicorn with multiple workers),
        you'd want a more robust scheduler that runs only once.
    """
    global email_thread_started
    if email_thread_started:
        return

    def worker():
        """Worker loop executed in a dedicated thread."""
        # We need an application context to access the database.
        with app.app_context():
            while True:
                check_and_send_reminders()
                time.sleep(60)

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()
    email_thread_started = True


# ---------------------------------------------------------------------------
# App setup (Flask 3 compatible)
# ---------------------------------------------------------------------------

def setup_app() -> None:
    """Initialize database and start the email thread once at startup.

    This is called at import time so it works both when running with
    `python app.py` and when running under a WSGI server / `flask run`.
    """
    with app.app_context():
        init_db()
        start_email_thread()

# Run setup immediately on import
setup_app()



# ---------------------------------------------------------------------------
# Routes: general navigation and login
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    """Home route simply redirects to today's list or to login."""
    user = get_current_user()
    if user:
        return redirect(url_for("today"))
    return redirect(url_for("login"))


@app.route("/login", methods=["GET", "POST"])
def login():
    """Very simple 'login' that lets the user pick a name from a drop-down.

    No passwords are used here; this is meant for a small internal tool.
    """
    users = User.query.order_by(User.name).all()

    if request.method == "POST":
        user_id = request.form.get("user_id")
        if not user_id:
            flash("Please choose a user.", "warning")
            return redirect(url_for("login"))

        user = User.query.get(int(user_id))
        if user is None:
            flash("User not found.", "danger")
            return redirect(url_for("login"))

        session["user_id"] = user.id
        flash(f"Logged in as {user.name}", "success")
        return redirect(url_for("today"))

    return render_template("login.html", users=users)


@app.route("/logout")
def logout():
    """Clear the session so the user is logged out."""
    session.clear()
    flash("You have been logged out.", "info")
    return redirect(url_for("login"))


# ---------------------------------------------------------------------------
# Routes: today's list
# ---------------------------------------------------------------------------

@app.route("/today")
@login_required
def today():
    """Show today's to-do list for the currently selected user."""
    user = get_current_user()
    today_date = date.today()
    weekday_index = get_weekday_index(today_date)
    weekday_name = get_weekday_name(weekday_index)

    # For Saturday, we simply show a friendly message and no list.
    if weekday_index > 5:
        return render_template(
            "today.html",
            user=user,
            weekday_name=weekday_name,
            date=today_date,
            states=[],
            is_saturday=True,
        )

    states = ensure_states_for_today(user)

    return render_template(
        "today.html",
        user=user,
        weekday_name=weekday_name,
        date=today_date,
        states=states,
        is_saturday=False,
    )


@app.route("/toggle/<int:state_id>", methods=["POST"])
@login_required
def toggle(state_id: int):
    """Toggle the checked state of a TodoState entry belonging to the current user."""
    user = get_current_user()
    state = TodoState.query.get_or_404(state_id)

    # Security check: users may not change each other's items.
    if state.user_id != user.id:
        flash("You cannot modify another user's items.", "danger")
        return redirect(url_for("today"))

    state.checked = not state.checked
    db.session.commit()
    return redirect(url_for("today"))


# ---------------------------------------------------------------------------
# Routes: configuration (per user, per weekday)
# ---------------------------------------------------------------------------

@app.route("/config")
@login_required
def config_redirect():
    """Redirect '/config' to the configuration page of the current user."""
    user = get_current_user()
    return redirect(url_for("config_user", user_id=user.id))


@app.route("/config/<int:user_id>", methods=["GET", "POST"])
@login_required
def config_user(user_id: int):
    """User (or admin) configuration main screen.

    This screen:
    - Lets the current user (or admin) edit the user's email and notification settings.
    - Shows links to configure each weekday's preset list.
    - Provides a link to export that user's presets to Excel.
    """
    current = get_current_user()

    # You can only edit your own settings unless you're admin.
    if current.id != user_id and not current.is_admin:
        flash("You can only edit your own settings unless you are admin.", "danger")
        return redirect(url_for("today"))

    user = User.query.get_or_404(user_id)

    if request.method == "POST":
        # Update email and notification settings.
        user.email = request.form.get("email", "").strip()
        user.email_notifications_enabled = bool(
            request.form.get("email_notifications_enabled")
        )

        email_time = request.form.get("email_time", "23:59").strip()

        # Very simple validation: HH:MM format with 5 characters.
        if len(email_time) != 5 or ":" not in email_time:
            flash("Invalid time format. Use HH:MM (24h).", "danger")
        else:
            user.email_time = email_time

        db.session.commit()
        flash("User configuration saved.", "success")
        return redirect(url_for("config_user", user_id=user.id))

    weekday_names = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday","Saturday"]

    return render_template(
        "config_user.html",
        current=current,
        user=user,
        weekday_names=weekday_names,
    )


@app.route("/config/<int:user_id>/day/<int:weekday>", methods=["GET", "POST"])
@login_required
def config_day(user_id: int, weekday: int):
    """Edit the preset list of a specific weekday for a given user.

    The UI:
    - Always shows at least one empty row at the beginning.
    - Has an 'Add item' button to add more empty rows.
    On save:
    - We delete the old presets for that day and recreate from the
      posted list of non-empty items, preserving their order.
    """
    current = get_current_user()

    # Security: non-admin users can only edit their own lists.
    if current.id != user_id and not current.is_admin:
        flash("You can only edit your own lists unless you are admin.", "danger")
        return redirect(url_for("today"))

    # Weekday must be in Sunday (0) .. Saturday (6).
    if weekday < 0 or weekday > 6:
        flash("Weekday must be between Sunday (0) and Saturday (6).", "danger")
        return redirect(url_for("config_user", user_id=user_id))

    user = User.query.get_or_404(user_id)
    weekday_name = get_weekday_name(weekday)

    if request.method == "POST":
        # Remove all existing templates for this day/user combination.
        TodoTemplate.query.filter_by(user_id=user.id, weekday=weekday).delete()

        # Recreate from submitted items in order.
        # We name all inputs "items[]" so Flask gives us a list.
        items = request.form.getlist("items[]")
        order = 0
        for text in items:
            clean = text.strip()
            if not clean:
                # Empty lines are simply ignored (this also acts as "delete").
                continue
            tmpl = TodoTemplate(
                user_id=user.id,
                weekday=weekday,
                order_index=order,
                text=clean,
            )
            db.session.add(tmpl)
            order += 1

        db.session.commit()
        flash(f"Preset list for {weekday_name} saved.", "success")
        return redirect(url_for("config_day", user_id=user.id, weekday=weekday))

    # For GET: load existing templates and show them.
    templates = (
        TodoTemplate.query
        .filter_by(user_id=user.id, weekday=weekday)
        .order_by(TodoTemplate.order_index)
        .all()
    )

    return render_template(
        "config_day.html",
        current=current,
        user=user,
        weekday=weekday,
        weekday_name=weekday_name,
        templates=templates,
    )


# ---------------------------------------------------------------------------
# Routes: user management (admin only)
# ---------------------------------------------------------------------------

@app.route("/users", methods=["GET", "POST"])
@admin_required
def users():
    """Admin-only user management screen.

    Admin can:
    - See all users.
    - Create new users.
    - Jump to configuration of a particular user.
    """
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        email = request.form.get("email", "").strip()

        if not name:
            flash("User name is required.", "danger")
            return redirect(url_for("users"))

        existing = User.query.filter_by(name=name).first()
        if existing:
            flash("User with that name already exists.", "danger")
            return redirect(url_for("users"))

        user = User(name=name, email=email, is_admin=False)
        db.session.add(user)
        db.session.commit()
        flash(f"User '{name}' created.", "success")
        return redirect(url_for("users"))

    users = User.query.order_by(User.id).all()
    return render_template("users.html", users=users)


# ---------------------------------------------------------------------------
# Routes: Excel export
# ---------------------------------------------------------------------------

def build_user_workbook(user: User) -> Workbook:
    """Create an in-memory Excel workbook for a single user's presets.

    Sheet columns:
      - Weekday (text, e.g. 'Sunday')
      - Order (numeric order index)
      - Item (task text)
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "ToDo Presets"

    # Header row.
    ws.append(["Weekday", "Order", "Item"])

    # Add all templates grouped by weekday, then ordered by order_index.
    templates = (
        TodoTemplate.query
        .filter_by(user_id=user.id)
        .order_by(TodoTemplate.weekday, TodoTemplate.order_index)
        .all()
    )
    for tmpl in templates:
        ws.append([get_weekday_name(tmpl.weekday), tmpl.order_index, tmpl.text])

    return wb


@app.route("/export/user/<int:user_id>")
@login_required
def export_user(user_id: int):
    """Export a single user's presets to an Excel file."""
    current = get_current_user()

    if current.id != user_id and not current.is_admin:
        flash("You can only export your own data unless you are admin.", "danger")
        return redirect(url_for("today"))

    user = User.query.get_or_404(user_id)
    wb = build_user_workbook(user)

    # Save workbook to an in-memory buffer.
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    filename = f"{user.name.replace(' ', '_')}_todo_presets.xlsx"
    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype="application/"
                 "vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/export/all")
@admin_required
def export_all():
    """Export all users' presets into a single Excel workbook.

    Each user gets their own sheet.
    """
    # Create a workbook and remove the default sheet; we'll add per-user sheets.
    master_wb = Workbook()
    default_ws = master_wb.active
    master_wb.remove(default_ws)

    users = User.query.order_by(User.id).all()
    for user in users:
        # Build a sheet for the user.
        # Excel sheet names are limited to 31 characters.
        ws = master_wb.create_sheet(title=user.name[:31])
        ws.append(["Weekday", "Order", "Item"])

        templates = (
            TodoTemplate.query
            .filter_by(user_id=user.id)
            .order_by(TodoTemplate.weekday, TodoTemplate.order_index)
            .all()
        )
        for tmpl in templates:
            ws.append([get_weekday_name(tmpl.weekday), tmpl.order_index, tmpl.text])

    buffer = BytesIO()
    master_wb.save(buffer)
    buffer.seek(0)

    filename = "all_users_todo_presets.xlsx"
    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype="application/"
                 "vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------------------------------------------------------------------------
# Entrypoint (development)
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    # When running directly with `python app.py`, just run the server.
    # setup_app() has already initialized the DB and email thread.
    app.run(host="0.0.0.0", port=5000, debug=True)
