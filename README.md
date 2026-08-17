# AD2L Moderator Report Generator

Generates a color-coded Excel report of every team roster in the current AD2L season, enriched with OpenDota rank/MMR stats and league-eligibility flags (previous champions, alt accounts, players on multiple teams, low recent activity).

## What it does

The pipeline runs in three steps (`ad2l.py`):

1. **Fetch players** — pulls the current season's rosters from the AD2L MySQL database, along with previous playoff winners (for eligibility checks) and approved alt-account relationships. Results are cached to `databasePlayers.json`.
2. **Fetch OpenDota stats** — for every player (and any off-roster alts), fetches current rank, leaderboard position, and ranked-games-played-in-90-days from the OpenDota API, using a thread pool of concurrent workers. Already-cached players are skipped, so this step is safe to re-run.
3. **Build Excel report** — writes `ad2l_report.xlsx`, one worksheet per league, with a formatted block per team. Rows are color-coded by rank tier and flagged for:
   - Ineligible players (playing below the league they previously won)
   - Previous champions
   - Low ranked-game count in the last 90 days
   - Players rostered on multiple teams
   - Alt accounts (both on-roster and off-roster)

## Requirements

- Python 3.10+
- Dependencies:
  ```
  pip install -r requirements.txt
  ```

## Configuration

Copy `.env.example` to `.env.local` in this directory and fill in your values:

```
DATABASE_URL=
DATABASE_USERNAME=
DATABASE_PASSWORD=
DATABASE_NAME=ad2l_production
DATABASE_PORT=3306

OPEN_DOTA_API_KEY=
```

`DATABASE_NAME` and `DATABASE_PORT` default to `ad2l_production` and `3306` if omitted. `.env.local` is gitignored — never commit it.

## Usage

```
python ad2l.py
```

This runs all three steps in order and produces `ad2l_report.xlsx` in the current directory.

To re-run just the Excel formatting step after manually editing `databasePlayers.json` (e.g. fixing a missing grand-final score), you can call `step3_create_excel()` directly from a Python shell instead of re-running the full pipeline.

## Notes

- **Current season detection**: the current season is auto-detected as the highest-numbered non-playoff season prefix (e.g. `S47`) in the `seasons` table — no manual configuration needed each season.
- **Missing grand-final scores**: if a playoff grand-final match has no score entered in the database, the run aborts with a message listing the affected matches — enter the scores and re-run.
- **Caching**: `databasePlayers.json` persists between runs. Delete it (or the `APId` fields inside it) to force a fresh fetch.
- **Rate limiting**: OpenDota requests automatically back off and retry on HTTP 429 responses.

## Project structure

```
ad2l.py               Entry point — orchestrates the 3-step pipeline
requirements.txt      Python dependencies
.env.example          Template for .env.local
src/
  config.py           Loads settings from .env.local
  database.py         MySQL queries: rosters, playoff winners, alt accounts
  opendota.py         OpenDota API client with retry handling
  excel.py            Excel report generation (xlsxwriter)
  utils.py            Shared helpers: rank conversion, JSON I/O, league ranks
```

Both `databasePlayers.json` (cache) and `ad2l_report.xlsx` (output) are created on first run and are gitignored.
