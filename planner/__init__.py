"""
Planner: your Outlook calendar and Todoist tasks in one week view.

It runs a small web app on your own computer (http://127.0.0.1:8765) and opens it
in its own window. Nothing gets installed and no admin consent is needed: it uses
only Python 3.9+ and its standard library.

    py -m planner           start the app
    py -m planner --demo    try it with sample data
    py -m planner --help    other options

Where the data comes from
  Todoist   your personal API token (Todoist > Settings > Integrations > Developer)
  Outlook   either the ICS link from "Publish a calendar" in Outlook on the web,
            or (Windows only) classic Outlook, which the app runs in the background with no window

Settings are saved in a small JSON file in your user profile (see config_path()).
Your Todoist token is stored there in plain text, readable only by your account.
"""

APP_NAME = "Planner"
VERSION = "1.0"
DEFAULT_PORT = 8765
