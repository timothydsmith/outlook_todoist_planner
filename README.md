# Planner: Outlook calendar + Todoist in one view

A small app that shows your Outlook meetings and your scheduled Todoist tasks in the same week (or day) view. It runs only on your own computer and opens in its own window. Nothing gets installed, and it doesn't need an Azure app registration or admin consent.

What you can do in it:

- See meetings (solid blocks) and Todoist tasks (outlined blocks, with a colour edge for priority) side by side.
- Drag a task onto the calendar to give it a time, or onto the all-day row to give it a date. Drag the bottom edge of a task to change its length.
- Drag overdue or undated tasks from the left-hand list onto the calendar.
- Tick a task's circle to complete it (with Undo).
- Double-click an empty slot to add a task at that time. It goes to your Todoist inbox.
- Click a meeting to see its details and a **Join meeting** button when it has a Teams or Zoom link.

Keyboard: ← and → move between weeks, **T** goes to today, **D** and **W** switch between day and week view, and **R** refreshes.

## What you need

- **Python 3.9 or newer.** On Windows you can install it without admin rights, either from the Microsoft Store ("Python 3.12" or later) or with the python.org installer with "Install for all users" left unticked. On a Mac, use the python.org installer.
- Your **Todoist API token**: in Todoist, open Settings › Integrations › Developer and copy the token.
- A way to read your **Outlook calendar**. Choose one:
  1. **Outlook on this PC (Windows; the default there).** Needs classic Outlook to be *installed*, but you don't need to open it or use it: Planner runs it in the background with no window, and closes it again when you close Planner. You can keep using the new Outlook as normal. Nothing is published.
  2. **Published calendar link (Windows or Mac).** In Outlook on the web, open Settings › Calendar › Shared calendars › Publish a calendar. Pick your calendar and "Can view titles and locations", select Publish, then copy the **ICS** link (not the HTML one).

## Running it

**Windows:** double-click `Start Planner.bat`. Or open a terminal in the folder and run:

```
py -m planner
```

**Mac:** open Terminal in the folder and run `python3 -m planner`.

Keep the whole folder together: the app's code lives in the `planner` folder. (`python outlook_todoist_planner.py` still works too.)

On first launch, the Settings screen opens. Paste your Todoist token, choose how to read Outlook, and select Save.

To try it with sample data first, run `py -m planner --demo`.

When you close the window, the app stops on its own. For a desktop shortcut on Windows, right-click `Start Planner.bat` › Send to › Desktop.

## For developers

- `planner/` holds the Python code, one module per area: `todoist.py`, `ics.py` (calendar files), `outlook/` (classic Outlook through a long-running PowerShell helper, `helper.ps1`), `sources.py` (where events and tasks come from), `app.py`, `server.py` (the local web server), and `window.py`.
- `planner/ui/` holds the page itself: `index.html`, `app.js` and `styles.css`. The server reads them on each request, so a browser refresh picks up any edits.
- Run the tests with `py -m unittest` (Mac: `python3 -m unittest`). They use only the standard library and need no accounts.

## Good to know

- **Your data stays local.** The app talks only to Todoist and to Outlook's calendar link. Settings, including your Todoist token, are saved in `%APPDATA%\OutlookTodoistPlanner\settings.json` on Windows or `~/Library/Application Support/OutlookTodoistPlanner/settings.json` on a Mac.
- **The published link is unauthenticated.** Anyone who has it can see what you published, so keep it to yourself, and check it sits within your organisation's IT policy. You can unpublish it at any time from the same Outlook settings page. If you'd rather publish nothing, the "Outlook on this PC" option avoids it.
- **Outlook is read-only here.** Change meetings in Outlook. A published calendar can take a few minutes to reflect changes.
- **Recurring tasks** show on their next date but can't be dragged, because that would break their repeat pattern. Change those in Todoist.
- **Show work tasks only:** in Settings, add a Todoist filter such as `#Work` or `@work`.
- **Time zones:** Outlook's published calendar includes its own time-zone details, so it needs nothing extra. If a meeting or task uses a time zone the app can't work out, the yellow bar names it and shows it in your own time zone. On Windows this usually means Python's time-zone data is missing: run `py -m pip install tzdata` once (it doesn't need admin rights), then restart the app.

## If something goes wrong

The app shows a yellow bar at the top with a plain explanation and an **Open settings** link. Common cases:

| What you see | What to do |
|---|---|
| "Todoist didn't accept the API token" | Copy the token again from Todoist and paste it in Settings. |
| "Outlook returned HTTP 404" | The calendar was unpublished or re-published. Copy the new ICS link. |
| "Publish a calendar" is missing in Outlook | Your organisation has turned publishing off. On Windows, use "Outlook on this PC" instead. |
| "Couldn't read the Outlook desktop app" | Classic Outlook needs to be installed (it doesn't need to be open). If the details mention a class not being registered, it isn't installed. |
| "Classic Outlook took too long to answer" | Classic Outlook may be waiting on a dialog box, such as a sign-in prompt. Open classic Outlook once, deal with the prompt, close it, and refresh. |
| "Outlook hasn't connected to the server yet" | Classic Outlook was still connecting when Planner read it, so the meetings shown come from its last saved copy. The next refresh tries again. |
| "Some items use time zones this computer's Python doesn't know" | Close the app, run `py -m pip install tzdata` (Mac: `python3 -m pip install tzdata`), and start it again. |
| Certificate errors on a Mac | Run "Install Certificates.command" from your Python folder in Applications. |
| `python` isn't recognised (Windows) | Install Python as described above, then try again. |

Other options: `--browser` opens the app in your normal browser rather than its own window, `--port 9000` uses a different local port, and `--keep-running` keeps it running after you close the window.
