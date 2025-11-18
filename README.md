# Daily To-Do Tracker (Flask, Jinja, Docker)

A small, touch-friendly **Python Flask web application** that manages **per-user, per-day to-do lists**, designed primarily for use from a **smartphone** (vertical layout, big tap targets).

Each user gets:

- Their own weekly **preset to-do lists** (Sunday–Friday)
- A **“Today”** page that shows the list for the current weekday
- Tap-to-cross / tap-to-uncross items (they stay in the list)
- A reset of “crossed/uncrossed” state each day
- Optional **email reminders** if a day ends with all items still un-crossed
- Ability to **export their presets as an Excel file**

The **first user** created is **Admin**, who can manage other users and export all user data.

---

## Features

Mapped to the original requirements:

1. **Docker-ready**  
   - The project includes a `Dockerfile` so the app can run as a container.

2. **Mobile / smartphone friendly UI**  
   - Uses Bootstrap and a narrow layout for **vertical / touch** usage.

3. **Day-of-week aware**  
   - The app computes the current weekday **with Sunday as the first day** (index 0).
   - Saturday has no working list (treated as a “day off”).

4. **Daily To-Do List behaviour**
   - Each day, the relevant **pre-set items for that weekday** appear in the list.
   - Items:
     - Do **not disappear** when clicked.
     - Are simply **crossed out** (line-through) when clicked.
     - Toggle cross/uncross when clicked again.
   - Day state:
     - The crossed/uncrossed state is **remembered for that specific date**.
     - A new day loads from presets and starts with everything **uncrossed**.

5. **Configuration screen for presets**
   - A per-user config screen with:
     - Email / notifications settings.
     - Links to edit presets for Sunday–Friday.

6. **Week starts on Sunday, ends on Friday**
   - Internally, weekdays are indexed with **Sunday=0 ... Friday=5**, Saturday not used.

7. **Day editing – new items**
   - When editing a day’s list:
     - One **empty entry is always at the top**.
     - An **“Add item”** button adds new empty lines for additional items.

8. **Day editing – delete items**
   - Individual lines can be removed.
   - Empty lines are ignored on save (effectively “delete”).

9. **Save configuration**
   - Each day-edit view has a **“Save … list”** button.
   - Saving replaces the old preset list for that user+weekday with the new one.

10. **Per-user databases**
    - Each user has their **own set of presets** for the week.
    - The app uses a single SQLite DB but keeps lists separate by `user_id`.

11. **Admin user**
    - On first run, an `Admin` user is auto-created.
    - Admin can:
      - Create new users.
      - Edit other users’ presets.
      - Export all users’ presets at once.

12. **Excel export**
    - Any user:
      - Can export **their own presets** to an `.xlsx` file.
    - Admin:
      - Can export **all users** into a single `.xlsx` file with one sheet per user.
    - Uses `openpyxl`.

13. **Per-user email field**
    - Each user has an **e-mail field**.
    - Editable in the configuration screen.

14. **Optional email notifications (off by default)**
    - For each user:
      - A switch to enable/disable **email notifications**.
      - If enabled:
        - At the configured time, if **all items in that day’s list are still un-crossed**, an email is sent.
    - Implemented via a simple background thread that checks every minute.

15. **Configurable email send time (default 23:59)**
    - Time is stored per user in `HH:MM` 24h format.
    - Default is **23:59**.

All core code comes with **docstrings and extensive comments**.

---

## Tech Stack

- **Backend:** Python, Flask, Flask-SQLAlchemy, SQLite
- **Templates:** Jinja2, Bootstrap 5
- **Excel:** openpyxl
- **Container:** Docker (Python slim image)

---

## Project structure (first version)

```text
.
├─ app.py               # Flask app (routes + models + email scheduler)
├─ requirements.txt     # Flask + SQLAlchemy + openpyxl
├─ Dockerfile           # Container definition
└─ templates/
   ├─ base.html         # Layout (navbar, mobile container)
   ├─ login.html        # User dropdown login
   ├─ today.html        # Today's list (touch-friendly)
   ├─ config_user.html  # User-level configuration (email, reminders, day links)
   ├─ config_day.html   # Edit presets for a specific weekday
   └─ users.html        # Admin user management
