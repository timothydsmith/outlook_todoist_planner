# Changelog

All notable changes to Planner are listed here. The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and versions follow [semantic versioning](https://semver.org).

## [Unreleased]

### Added
- Meetings show their Outlook category colour, as they do in Outlook. Tentative and free meetings show a lighter shade with a coloured edge. A meeting's categories are listed in its details. Colours need "Outlook on this PC", because published calendar links don't include them.

## [1.1.0] - 2026-11-08

### Added
- **Tasks in Outlook**, for Windows with "Outlook on this PC". When you tick it in Settings, each timed Todoist task gets its own event in your Outlook calendar. Moving or resizing the event in Outlook moves the task in Todoist. Deleting the event removes the task's time.
- New settings for the task events: show as Busy, Tentative or Free, and how far ahead to add them.
- The version number now shows at the bottom of Settings.

### Changed
- "Outlook on this PC" now keeps one background helper running instead of starting PowerShell for every refresh. Refreshes are faster, and the background Outlook closes when Planner closes.
- The code is now a `planner` package. Start it with `py -m planner`. `outlook_todoist_planner.py` and `Start Planner.bat` still work.
- New sample data for demo mode.

### Fixed
- Edits made in the new Outlook or on your phone are picked up straight away, because Planner asks classic Outlook to send and receive before reading.

## [1.0.0] - 2026-11-06

### Added
- The first version: Outlook meetings and scheduled Todoist tasks in one week or day view.
- To schedule a task, drag it onto the calendar. Drag a task's bottom edge to change its length.
- Complete tasks with Undo, and double-click an empty slot to add one.
- Outlook is read from a published ICS link or from classic Outlook on this PC.
- A test suite that uses only the standard library.

[Unreleased]: https://github.com/timothydsmith/outlook_todoist_planner/compare/v1.1.0...HEAD
[1.1.0]: https://github.com/timothydsmith/outlook_todoist_planner/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/timothydsmith/outlook_todoist_planner/releases/tag/v1.0.0
