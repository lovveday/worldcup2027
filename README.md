# World Cup 2026 Stats Tracker

Automatically collects match, team, and player stats for every World Cup 2026
fixture and keeps them updated in `data/world_cup_2026.xlsx` — no manual copy-pasting.

Data source: [API-Football](https://www.api-football.com/) (api-sports.io), a
licensed sports data API. This does **not** scrape SofaScore or any other site.

## What you get

The workbook has 5 sheets, all auto-updated:

- **Matches** — every fixture, score, stage, venue, status
- **Team Stats** — shots, possession, corners, cards, passes, xG, per team per match
- **Player Stats** — minutes, rating, goals, assists, shots, passes, tackles, duels, cards, per player per match
- **Match Events** — goals, cards, substitutions with minute and player
- **Update Log** — a row added every time the automation runs, so you can see it's alive

The script only spends API calls on matches that have **just finished** and that
it hasn't already recorded — so it stays well within the free tier's 100
requests/day even covering the entire 48-team tournament.

## One-time setup (15 minutes)

### 1. Get a free API-Football key
1. Go to https://dashboard.api-football.com/register and sign up (free, no card required)
2. Once logged in, copy your API key from the dashboard

### 2. Create the GitHub repo
1. Create a new **private** repo on GitHub (private keeps your workbook off public view — you can make it public later if you want)
2. Push these files to it:
   ```
   git init
   git add .
   git commit -m "Initial setup"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<repo-name>.git
   git push -u origin main
   ```

### 3. Add your API key as a GitHub Secret
1. In your repo: **Settings → Secrets and variables → Actions → New repository secret**
2. Name: `API_FOOTBALL_KEY`
3. Value: paste the key from step 1
4. Save

### 4. Turn it on
- The workflow (`.github/workflows/update-stats.yml`) is already scheduled to run twice a day automatically.
- To run it right now instead of waiting: go to the **Actions** tab → **Update World Cup 2026 Stats** → **Run workflow**.

### 5. Get your data
- After each run, check `data/world_cup_2026.xlsx` in the repo — it'll have committed updates.
- Download it any time, or sync the repo locally with `git pull` to always have the latest copy in a real Excel app.

## Adjusting the schedule

Edit the `cron` lines in `.github/workflows/update-stats.yml`. Cron times are
UTC. Two runs a day is plenty since most World Cup days have a handful of
matches finishing within a few hours of each other.

## Notes & limits

- Free tier = 100 requests/day. Each finished match costs ~3 requests (stats,
  players, events) plus 1 for the fixture list refresh. That's roughly 30
  matches/day worth of headroom — far more than any single World Cup day will produce.
- If you ever see `429` warnings in the Action logs, you've hit the daily cap;
  it'll simply pick up the rest on the next scheduled run.
- All times in the workbook are UTC.
- Per API-Football's own World Cup guide, **data coverage can be thinner for
  early-tournament matches** (e.g. some stats fields blank) and fills in more
  reliably as the tournament progresses. You can check what's supported via
  `GET /leagues?id=1&season=2026` and inspecting the `coverage` object — not
  something the script needs to do automatically, just useful to know if a
  cell looks empty.
