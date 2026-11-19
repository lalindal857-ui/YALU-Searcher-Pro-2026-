YALU SEARCHER PRO - NEON COMPACT EDITION

1. Extract this ZIP.
2. Push the extracted project to your GitHub repository (replace the old files).
3. Open Actions > Build YALU Searcher EXE > Run workflow.
4. Download the generated EXE artifact.

Features:
- Compact 283 x 114 floating window.
- Always stays above other windows.
- Drag the top header to move it around the screen.
- Search and copy journal descriptions from the Excel report.
- The full Entry No. (example: JE-FY2026-27528) is appended at the end
  of each copied description.
- Excel search logic and workbook column names are retained.

Excel sheet required: Journal Entries
Required columns:
- Entry No. / Line
- Posting Date / Cost Center
- Description


TASKBAR ICON:
- The EXE is built with yalu_searcher.ico and the app sets a stable Windows
  AppUserModelID plus the window icon.
- After downloading the new EXE, unpin the old YALU shortcut from the taskbar
  (if pinned), run the new EXE once, then right-click its taskbar icon and
  choose Pin to taskbar. Windows may cache an older pinned icon.


NEW APP ICON:
- Updated to a blue feather inside a magnifying glass.
- Used as the EXE icon and Windows taskbar/window icon.
- Rebuild the EXE with GitHub Actions.
- If the old icon remains pinned, unpin it and pin the newly built EXE.

ALL-FEATURES BUILD:
- Includes the feather magnifying-glass icon for EXE, window and taskbar.
- Excel report loading and Entry No. search.
- Copies each description with the full JE Entry No. appended at the end.
- Visible Cost Center and copied status in the compact window.
- Visible “Prepared By Lakshan” credit in the footer and copy notification.
- Always-on-top compact window and draggable custom header.
- Keyboard shortcuts: Ctrl+O opens an Excel report; Ctrl+L shows the last copied description.
- Small copied-success notification.
