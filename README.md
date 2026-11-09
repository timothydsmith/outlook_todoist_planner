# Planner: Outlook calendar + Todoist in one view

A small app that shows your Outlook meetings and your scheduled Todoist tasks in the same week (or day) view. It runs only on your own computer and opens in its own window. Nothing gets installed, and it doesn't need an Azure app registration or admin consent.

What you can do in it:

- See meetings (solid blocks, in their Outlook category colours) and Todoist tasks (outlined blocks, with a colour edge for priority) side by side. Category colours need "Outlook on this PC"; a published calendar link doesn't include them.
- Drag a task onto the calendar to give it a time, or onto the all-day row to give it a date. Drag the bottom edge of a task to change its length.
- Drag overdue or undated tasks from the left-hand list onto the calendar.
- Tick a task's circle to complete it (with Undo).
- Double-click an empty slot to add a task at that time. It goes to your Todoist inbox.
- Click a meeting to see its details and a **Join meeting** button when it has a Teams or Zoom link.
- Optionally, put your timed tasks into your Outlook calendar as events, and move them from there (Windows, with "Outlook on this PC").

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

## Tasks in Outlook

With "Outlook on this PC" selected, tick **Add timed tasks to my Outlook calendar** in Settings. Each Todoist task that has a time then gets its own event in your calendar, which you'll also see in the new Outlook and on your phone.

- **Which tasks:** tasks with a time, from today up to the number of weeks ahead you choose. Tasks with only a date don't get an event. Repeating tasks get one for their next occurrence.
- **Both directions:** moving a task in Planner or Todoist moves its event. Moving or resizing the event in Outlook moves the task in Todoist, and deleting the event removes the task's time (it keeps its date). If both were changed between syncs, Todoist wins and the yellow bar says so. Repeating tasks are always moved in Todoist; if you move one in Outlook, Planner puts it back.
- **Completing a task** removes its upcoming event. An event whose time has already started stays, as a record.
- **What's in an event:** the task name, its description and a link to it in Todoist. It shows you as Busy (Settings can change this to Tentative or Free) and has no reminder, because Todoist does reminders.
- **When it happens:** a couple of seconds after each change in Planner, and at each refresh. Changes you make in Outlook or Todoist while Planner is closed are picked up next time it runs.
- **Safety:** Planner tags its events and only ever changes or deletes those; your meetings are never touched. Deleted events go to Deleted Items. If Todoist can't be read or Outlook hasn't connected yet, nothing is changed, and if one sync would remove more than 10 events, it removes none and tells you why.
- **Turning it off** removes the upcoming events Planner added (to Deleted Items). Switching away from "Outlook on this PC" does the same.

Planner remembers which event belongs to which task in `outlook_sync.json`, next to its settings. If that file is lost, it finds its events again by their tags.

## For developers

- `planner/` holds the Python code, one module per area: `todoist.py`, `ics.py` (calendar files), `outlook/` (classic Outlook through a long-running PowerShell helper, `helper.ps1`, and the two-way task sync, `sync.py`), `sources.py` (where events and tasks come from), `app.py`, `server.py` (the local web server), and `window.py`.
- `planner/ui/` holds the page itself: `index.html`, `app.js` and `styles.css`. The server reads them on each request, so a browser refresh picks up any edits.
- Run the tests with `py -m unittest` (Mac: `python3 -m unittest`). They use only the standard library and need no accounts.

### Versions and branches

Planner uses [semantic versioning](https://semver.org): `MAJOR.MINOR.PATCH`. The version lives in `planner/__init__.py` (`VERSION`), shows at the bottom of Settings, and each release is listed in [CHANGELOG.md](CHANGELOG.md).

- **PATCH** (1.1.0 → 1.1.1): bug fixes only.
- **MINOR** (1.1.0 → 1.2.0): new features that keep existing settings and behaviour working.
- **MAJOR** (1.1.0 → 2.0.0): changes that break something you rely on, such as settings that have to be entered again or a removed feature.

Branches:

- **`main`** holds released versions only. Each release is tagged `vX.Y.Z`.
- **`dev`** is where work happens. Commit there, and note each change under "Unreleased" in the changelog.

To release from `dev`:

1. Choose the new version number. Move the "Unreleased" notes under a new heading for it in `CHANGELOG.md`, and set `VERSION` to match. A test checks that the two agree.
2. Commit, then run the tests.
3. Merge into `main` with `git switch main`, then `git merge --no-ff dev`.
4. Tag the merge with `git tag -a vX.Y.Z -m "Planner X.Y.Z"`.
5. Push with `git push origin main dev --follow-tags`, then `git switch dev`.

## Good to know

- **Your data stays local.** The app talks only to Todoist and to Outlook's calendar link. Settings, including your Todoist token, are saved in `%APPDATA%\OutlookTodoistPlanner\settings.json` on Windows or `~/Library/Application Support/OutlookTodoistPlanner/settings.json` on a Mac.
- **The published link is unauthenticated.** Anyone who has it can see what you published, so keep it to yourself, and check it sits within your organisation's IT policy. You can unpublish it at any time from the same Outlook settings page. If you'd rather publish nothing, the "Outlook on this PC" option avoids it.
- **Meetings are read-only here.** Change them in Outlook. A published calendar can take a few minutes to reflect changes. The only events Planner writes are its own task events (see Tasks in Outlook).
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
| "Planner would have removed N task events from Outlook in one go" | A lot of tasks disappeared at once, often because the Todoist filter changed. If you meant it, turn "Add timed tasks to my Outlook calendar" off (which removes them) and on again. |
| "… was moved in both Outlook and Todoist" | Planner kept the Todoist time. Move it again wherever you like. |
| "Couldn't update your Outlook calendar" | The details say what classic Outlook reported. Planner tries again at the next refresh. |
| "Some items use time zones this computer's Python doesn't know" | Close the app, run `py -m pip install tzdata` (Mac: `python3 -m pip install tzdata`), and start it again. |
| Certificate errors on a Mac | Run "Install Certificates.command" from your Python folder in Applications. |
| `python` isn't recognised (Windows) | Install Python as described above, then try again. |

Other options: `--browser` opens the app in your normal browser rather than its own window, `--port 9000` uses a different local port, and `--keep-running` keeps it running after you close the window.
