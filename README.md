# Audition Reviewer — Studio Tiny Tuko

A Google Sheets sidebar for reviewing voice auditions, with an optional voice remote.

## What's in this repo

| File | What it is | Where it runs |
|---|---|---|
| **voice.html** | The voice remote window | GitHub Pages: https://davidjnunezf.github.io/voice-audition-tester/voice.html |
| **apps-script/AuditionReviewer.gs** | The sheet's script (menu, ratings, highlights) | Each audition sheet → Extensions → Apps Script, as **Code.gs** |
| **apps-script/AuditionSidebar.html** | The sidebar (player, buttons, keyboard shortcuts) | Same Apps Script project, as an HTML file named exactly **AuditionSidebar** |

The two apps-script files are a backup of what's running in the TAMA SAMPALOSA sheet. Google blocks the microphone inside Sheets sidebars, so voice runs in voice.html, which is hosted here and opened from the sidebar's 🎤 button.

## Using it

1. Open the sheet in desktop Chrome or Edge → **🎙️ Auditions → Open reviewer**.
2. Click a row to start from there (otherwise it starts at the first unlistened audition).
3. Optional: click **🎤 Open voice remote → Start listening**, allow the mic, and pick your mic in the menu. Wear headphones.

| Say | Key | What happens |
|---|---|---|
| next | N | General Rating 1, Listened? Yes, opens the next audition |
| good | G | 3, yellow row, next |
| peak | P | 5, green row, next |
| custom 3.5 (or just "three point five") | C, then type | Any score 0–5. Below 3 no colour, 3+ yellow, 4+ green |
| N/A | X | Listened? No, Role + General Rating N/A, red row, next |
| notes … | Enter in the notes box | Adds "- …" to Notes |
| back | B | Previous audition, rating unchanged |
| undo | U | Reverts the last rating |
| descending / ascending | ↓ / ↑ buttons | Work top → bottom / bottom → top |
| top / bottom | | Jump to the first unlistened audition from that end |
| replay / pause / play | R / Space | Audio controls |

## Adding it to another role sheet

1. Open that sheet → Extensions → Apps Script (if it fails with "unable to open the file", open the sheet in a Chrome profile where studiotinytuko is the only Google account).
2. Paste **apps-script/AuditionReviewer.gs** into Code.gs.
3. Add an HTML file named **AuditionSidebar** and paste **apps-script/AuditionSidebar.html**.
4. Save, reload the sheet, run **🎙️ Auditions → Open reviewer** and approve the permissions.

Columns are found by header name (Full Name, Listened?, General Rating / Rating, Role Rating, Notes, Submit Your Audition…, Submitting with a link…), so no edits are needed.

## Who can use it

Anyone with **edit** access to the sheet, in desktop Chrome or Edge, with access to the audition upload folder in Drive. Their sheet account should be the only (or first) Google account in that Chrome profile.

## If voice doesn't hear anything

The voice window shows a live mic-level bar and explains most problems in red. Common fixes: use desktop Chrome or Edge (not Brave, Arc, Firefox or a phone); allow the mic in the address bar; on Mac allow Chrome under System Settings → Privacy & Security → Microphone; pick the right mic in the Microphone menu; turn off VPNs.

## Changing things

- Ratings, colours and thresholds: **AR_CONFIG.commands** and **AR_CONFIG.colorBands** in the script.
- Spoken words for each command: **SYNONYMS** in voice.html.
