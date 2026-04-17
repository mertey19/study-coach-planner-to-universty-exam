# Study Coach Planner

A desktop planning app for weekly study scheduling, TYT/AYT exam tracking, and progress analytics.

Built with Python + Tkinter, with JSON-based local storage and optional Excel/PDF workflows.

## Highlights

- Weekly planning by day/hour (`00:00` to `23:00`)
- Student-based profiles with week management (`Week 1`, copy week, quick switching)
- TYT/AYT exam tracking with subject nets and track-aware AYT total net calculation
- Right-side analytics panel with:
  - Daily planned vs completed study bars
  - Weekly completion rate bars
  - Exam total net chart
  - Subject net trend chart
  - Type/time filters (`All`, `TYT`, `AYT`, `Last 30 days`, `Last 3 months`)
- Quality-of-life tools:
  - Auto text generation from selected course/topic/source
  - Per-student remembered selections
  - Quick suggestions from past entries
  - One-click auto-plan (adds 3 entries to selected day)
- Data safety:
  - Automatic backup
  - Timestamped backups (`backups/`)
  - JSON export/import
  - Restore from backup
  - Full data reset (with backup)

## Project Structure

```text
excel_organiser/
?? main.py
?? excel_organizer.py
?? constants.py
?? requirements.txt
?? CHANGELOG.md
?? EXE_KURULUM.md
?? config/
?? storage/
?? services/
?? data/
```

## Requirements

- Python 3.8+
- Windows recommended for `.exe` packaging

Install dependencies:

```powershell
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

## Run the App

```powershell
.\.venv\Scripts\activate
python main.py
```

## Packaging as `.exe` (Windows)

Quick one-file build:

```powershell
.\.venv\Scripts\activate
pip install pyinstaller
pyinstaller --noconfirm --clean --onefile --windowed --name "StudyCoachPlanner" main.py
```

Output:

- `dist\StudyCoachPlanner.exe`

For a full distribution guide (onefile/onedir, troubleshooting), see `EXE_KURULUM.md`.

## Data Files

App data is stored locally in the app folder:

- `program_kayitlari.json` (main data)
- `config/` (settings)
- `backups/` (timestamped backup files)

## Keyboard Shortcuts

- `Ctrl + F`: focus list filter
- `Ctrl + Shift + A`: add 3 automatic entries to selected day

## Changelog

See `CHANGELOG.md` for release notes (`v1.0.0` and later).

## License

No license file is currently defined.  
If you want, I can add an MIT license and update this section.
