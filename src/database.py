import pymysql
from .config import DATABASE_URL, DATABASE_USER, DATABASE_PASSWORD, DATABASE_NAME, DATABASE_PORT
from .utils import LEAGUE_RANK, get_league_name


def _connect():
    return pymysql.connect(
        host=DATABASE_URL,
        user=DATABASE_USER,
        password=DATABASE_PASSWORD,
        database=DATABASE_NAME,
        port=DATABASE_PORT,
    )


def get_current_season_ids() -> list[int]:
    """
    Automatically finds the most recent set of regular-season IDs.
    Looks for the highest season-number prefix (e.g. 'S46') among
    non-playoff seasons and returns all season IDs with that prefix.
    """
    conn = _connect()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT id, SUBSTRING_INDEX(title, ' ', 1) AS prefix
            FROM seasons
            WHERE title NOT LIKE '%Playoffs%'
              AND sport_id = (
                  SELECT sport_id FROM seasons
                  WHERE title NOT LIKE '%Playoffs%'
                  ORDER BY id DESC LIMIT 1
              )
            ORDER BY id DESC
        """)
        rows = cursor.fetchall()
    finally:
        conn.close()

    if not rows:
        raise RuntimeError("No regular seasons found in the database.")

    current_prefix = rows[0][1]  # e.g. "S46"
    ids = [row[0] for row in rows if row[1] == current_prefix]
    return ids


def get_previous_playoff_ids(num_prev_seasons: int = 2) -> list[int]:
    """
    Returns playoff season IDs for the N most recently completed seasons
    (i.e. the season groups before the current one).
    Used to check if players won a championship and may be ineligible.
    """
    conn = _connect()
    try:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT id, SUBSTRING_INDEX(title, ' ', 1) AS prefix
            FROM seasons
            WHERE title LIKE '%Playoffs%'
              AND sport_id = (
                  SELECT sport_id FROM seasons
                  WHERE title NOT LIKE '%Playoffs%'
                  ORDER BY id DESC LIMIT 1
              )
            ORDER BY id DESC
        """)
        rows = cursor.fetchall()
    finally:
        conn.close()

    # Group IDs by prefix, ordered most-recent first
    seen_prefixes = []
    groups: dict[str, list[int]] = {}
    for sid, prefix in rows:
        if prefix not in groups:
            groups[prefix] = []
            seen_prefixes.append(prefix)
        groups[prefix].append(sid)

    # Skip the current season's own playoffs (if they exist), take the next N
    # The current season prefix is the one with the highest IDs among non-playoffs
    current_prefix = seen_prefixes[0] if seen_prefixes else None

    prev_ids = []
    count = 0
    for prefix in seen_prefixes:
        if prefix == current_prefix:
            continue  # skip current season's playoffs
        prev_ids.extend(groups[prefix])
        count += 1
        if count >= num_prev_seasons:
            break

    return prev_ids


def fetch_players(season_ids: list[int]) -> list[dict]:
    season_str = ", ".join(str(s) for s in season_ids)
    conn = _connect()
    try:
        cursor = conn.cursor()
        cursor.execute(f"""
            SELECT
                s.title,
                t.name  AS team_name,
                p.name  AS player_name,
                p.steam32id AS steam_id,
                p.id    AS player_id
            FROM seasons s
            JOIN team_seasons ts ON ts.season_id = s.id
            JOIN teams t         ON ts.participant_id = t.id
            JOIN players_teams pt ON pt.team_id = t.id
            JOIN players p        ON pt.player_id = p.id
            WHERE ts.season_id IN ({season_str})
            GROUP BY s.title, t.name, p.name, p.steam32id, p.id
        """)
        results = cursor.fetchall()
    finally:
        conn.close()

    return [
        {"season": row[0], "team": row[1], "name": row[2], "steam_id": row[3], "player_id": row[4]}
        for row in results
    ]


def fetch_previous_winners(playoff_season_ids: list[int]) -> dict[str, dict]:
    """
    Returns a dict keyed by steam_id with winner info for each player
    who was on a grand-final-winning team in the given playoff seasons.

    Shape: { "steam_id": { "won_league": "Heroic", "won_season": "S45",
                            "won_rank": 6 } }
    """
    id_str = ", ".join(str(s) for s in playoff_season_ids)
    conn = _connect()
    try:
        cursor = conn.cursor()

        # Find the grand-final match per playoff season (last week, no winner_match_id)
        cursor.execute(f"""
            SELECT s.title,
                   m.id,
                   m.home_participant_id, m.away_participant_id,
                   m.home_score,         m.away_score
            FROM matches m
            JOIN seasons s ON s.id = m.season_id
            WHERE m.season_id IN ({id_str})
              AND m.winner_match_id IS NULL
              AND m.week = (
                  SELECT MAX(m2.week) FROM matches m2
                  WHERE m2.season_id = m.season_id
              )
        """)
        all_finals = cursor.fetchall()

        # Check for finals with no scores — cannot determine winner, must be resolved manually
        missing = [(title, match_id) for title, match_id, home_id, away_id, hs, aws in all_finals
                   if hs is None and aws is None]
        if missing:
            lines = [
                "",
                "=" * 60,
                "  ACTION REQUIRED — Missing grand-final scores",
                "=" * 60,
                "  The following playoff finals have no scores entered.",
                "  Winner cannot be determined automatically — players",
                "  from these seasons will NOT be flagged as prev_winner.",
                "  Please enter the correct scores in the database before",
                "  running this report.",
                "",
            ]
            for title, match_id in missing:
                lines.append(f"  • {title}")
                lines.append(f"    https://dota.playon.gg/matches/{match_id}")
            lines += ["=" * 60, ""]
            raise RuntimeError("\n".join(lines))

        finals = [(t, mid, h, a, hs, aws) for t, mid, h, a, hs, aws in all_finals]

        # Determine winning team_id per final
        winning_team_ids = []
        for title, _match_id, home_id, away_id, hs, aws in finals:
            league = get_league_name(title)
            season_prefix = title.split()[0]  # e.g. "S44"
            if hs is not None and aws is not None:
                winner_id = home_id if hs > aws else away_id
            elif hs is None:
                winner_id = away_id
            else:
                winner_id = home_id
            winning_team_ids.append((winner_id, league, season_prefix))

        # Get rosters for every winning team
        winners = {}
        for team_id, league, season_prefix in winning_team_ids:
            cursor.execute("""
                SELECT p.steam32id
                FROM players_teams pt
                JOIN players p ON p.id = pt.player_id
                WHERE pt.team_id = %s
                  AND p.steam32id IS NOT NULL AND p.steam32id != ''
            """, (team_id,))
            for (steam_id,) in cursor.fetchall():
                sid = str(steam_id)
                # Keep the highest-ranked win if a player won multiple leagues
                existing = winners.get(sid)
                new_rank = LEAGUE_RANK.get(league, 0)
                if existing is None or new_rank > existing["won_rank"]:
                    winners[sid] = {
                        "won_league":  league,
                        "won_season":  season_prefix,
                        "won_rank":    new_rank,
                    }
    finally:
        conn.close()

    return winners


def group_by_season_and_team(players: list[dict]) -> dict:
    seasons = {}
    for player in players:
        (seasons
         .setdefault(player["season"], {})
         .setdefault(player["team"],   [])
         .append(player))
    return seasons


def fetch_alt_accounts(player_ids: list[int]) -> dict[int, dict]:
    """
    Given a list of internal player IDs (p.id), returns approved alt-account
    relationships that involve any of those players.

    For players who are playing on an ALT account (alt_id in roster):
      { player_id: { "is_alt": True, "main_player_id": int,
                     "main_player_name": str, "main_steam_id": str } }

    For players who are a MAIN account with registered alts:
      { player_id: { "is_alt": False,
                     "alt_accounts": [
                       { "player_id": int, "name": str, "steam_id": str,
                         "on_roster": bool }  # on_roster=True if alt is also playing
                     ] } }
    """
    if not player_ids:
        return {}

    player_id_set = set(player_ids)
    id_str = ", ".join(str(i) for i in player_ids)
    conn = _connect()
    try:
        cursor = conn.cursor()
        cursor.execute(f"""
            SELECT pa.main_id, pa.alt_id,
                   pm.name        AS main_name,
                   pm.steam32id   AS main_steam_id,
                   pa_p.name      AS alt_name,
                   pa_p.steam32id AS alt_steam_id
            FROM player_alts pa
            JOIN players pm  ON pm.id  = pa.main_id
            JOIN players pa_p ON pa_p.id = pa.alt_id
            WHERE pa.approved = 1
              AND (pa.main_id IN ({id_str}) OR pa.alt_id IN ({id_str}))
        """)
        rows = cursor.fetchall()
    finally:
        conn.close()

    result: dict[int, dict] = {}
    for main_id, alt_id, main_name, main_steam_id, alt_name, alt_steam_id in rows:
        # Player is on the roster playing on their alt account
        if alt_id in player_id_set:
            result[alt_id] = {
                "is_alt":           True,
                "main_player_id":   main_id,
                "main_player_name": main_name,
                "main_steam_id":    str(main_steam_id) if main_steam_id else None,
            }
        # Player is a main account on the roster — collect ALL their registered alts
        if main_id in player_id_set:
            # May already exist from the branch above (player is an alt AND a main),
            # in which case the entry has no "alt_accounts" key yet.
            entry = result.setdefault(main_id, {"is_alt": False})
            entry.setdefault("alt_accounts", []).append({
                "player_id": alt_id,
                "name":      alt_name,
                "steam_id":  str(alt_steam_id) if alt_steam_id else None,
                "on_roster": alt_id in player_id_set,
            })

    return result


def annotate_alt_flags(seasons: dict, alt_map: dict[int, dict]) -> None:
    """
    Adds alt-account fields directly onto each player dict in-place:
      - is_alt           : bool
      - main_player_id   : int or None       (set when is_alt is True)
      - main_player_name : str or None
      - main_steam_id    : str or None
      - alt_accounts     : list of dicts     (set when player is a main with alts)
          Each entry: { player_id, name, steam_id, on_roster, APId (added later) }
    """
    for teams in seasons.values():
        for players in teams.values():
            for player in players:
                pid = player.get("player_id")
                info = alt_map.get(pid) if pid is not None else None
                if info and info["is_alt"]:
                    player["is_alt"]           = True
                    player["main_player_id"]   = info["main_player_id"]
                    player["main_player_name"] = info["main_player_name"]
                    player["main_steam_id"]    = info["main_steam_id"]
                    player["alt_accounts"]     = []
                else:
                    player["is_alt"]           = False
                    player["main_player_id"]   = None
                    player["main_player_name"] = None
                    player["main_steam_id"]    = None
                    player["alt_accounts"]     = info["alt_accounts"] if info else []


def annotate_multi_team_flags(seasons: dict) -> None:
    """
    Flags players who appear on more than one team across any league.
    Adds 'other_teams': [(season_title, team_name), ...] to every player.
    """
    appearances: dict[str, list[tuple[str, str]]] = {}
    for season_title, teams in seasons.items():
        for team_name, players in teams.items():
            for player in players:
                sid = str(player.get("steam_id", ""))
                if sid:
                    appearances.setdefault(sid, []).append((season_title, team_name))

    for season_title, teams in seasons.items():
        for team_name, players in teams.items():
            for player in players:
                sid = str(player.get("steam_id", ""))
                locs = appearances.get(sid, [])
                player["other_teams"] = [
                    (s, t) for s, t in locs
                    if not (s == season_title and t == team_name)
                ]


def annotate_winner_flags(seasons: dict, previous_winners: dict) -> None:
    """
    Adds flag fields directly onto each player dict in-place:
      - prev_winner  : bool
      - won_league   : str or None
      - won_season   : str or None
      - ineligible   : bool  (playing in a lower league than they won)
    """
    for season_title, teams in seasons.items():
        current_league = get_league_name(season_title)
        current_rank   = LEAGUE_RANK.get(current_league, 0)

        for players in teams.values():
            for player in players:
                info = previous_winners.get(str(player["steam_id"]))
                if info:
                    player["prev_winner"] = True
                    player["won_league"]  = info["won_league"]
                    player["won_season"]  = info["won_season"]
                    player["ineligible"]  = current_rank < info["won_rank"]
                else:
                    player["prev_winner"] = False
                    player["won_league"]  = None
                    player["won_season"]  = None
                    player["ineligible"]  = False
