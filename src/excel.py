import re

import xlsxwriter
from .utils import read_json, convert_rank, get_league_name

# Excel rejects [ ] : * ? / \ in worksheet names and caps them at 31 chars.
_INVALID_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")


def _safe_sheet_name(name: str) -> str:
    """Make a season title usable as an Excel worksheet name."""
    return _INVALID_SHEET_CHARS.sub("-", name).strip("'")[:31] or "Sheet"

# --- Column definitions ---
COLUMNS = [
    {"header": "Player Name",          "width": 22},
    {"header": "Current Rank",         "width": 14},
    {"header": "Leaderboard",          "width": 13},
    {"header": "Ranked (90d)",          "width": 12},
    {"header": "OpenDota",             "width": 10},
    {"header": "Dotabuff",             "width": 10},
    {"header": "Flag",                 "width": 45},
    {"header": "Other Teams",          "width": 35},
    {"header": "MMR",                  "width": 10},
    {"header": "MMR Screenshot",       "width": 18},
]

# --- Color palette ---
COLOR_TEAM_BG       = "#1A3A5C"   # dark navy  — team header
COLOR_TEAM_TEXT     = "#FFFFFF"
COLOR_COL_BG        = "#2E75B6"   # medium blue — column headers
COLOR_COL_TEXT      = "#FFFFFF"
COLOR_ROW_ODD       = "#EBF5FB"   # very light blue
COLOR_ROW_EVEN      = "#FFFFFF"
COLOR_PRIVATE       = "#AAAAAA"   # grey for private accounts
COLOR_BORDER        = "#BDD7EE"   # soft blue border
COLOR_LINK          = "#1155CC"   # hyperlink blue
COLOR_INELIGIBLE    = "#FF0000"   # red text — player cannot play here
COLOR_INELIG_BG     = "#FFE0E0"   # light red background
COLOR_WINNER_BG     = "#FFF3CD"   # light amber background
COLOR_LOW_GAMES_BG  = "#FFE5CC"   # light orange — low ranked game count
COLOR_MULTI_TEAM_BG = "#D6EEF0"   # soft teal — player on multiple teams
COLOR_ALT_BG        = "#F3E5F5"   # light purple — alt account

MIN_RANKED_GAMES   = 15

# Dota 2 rank tier colors (bg, text) — tiers 1-8 matching Herald→Immortal
RANK_COLORS = {
    1: ("#D6EDCB", "#3A7A28"),  # Herald   — green
    2: ("#E8D9C0", "#7A5A2A"),  # Guardian — brown
    3: ("#C8E8E6", "#1A8A80"),  # Crusader — teal
    4: ("#FFF3CC", "#9A7A00"),  # Archon   — gold
    5: ("#FFE8D0", "#9A6020"),  # Legend   — orange-gold
    6: ("#EDD5F0", "#7A2A9A"),  # Ancient  — purple
    7: ("#CCE4F7", "#2A70AA"),  # Divine   — light blue
    8: ("#FFD0B8", "#AA3A10"),  # Immortal — orange-red
}


def _make_formats(workbook: xlsxwriter.Workbook) -> dict:
    base = {"font_name": "Calibri", "font_size": 11, "border": 1,
            "border_color": COLOR_BORDER}

    def fmt(overrides):
        return workbook.add_format({**base, **overrides})

    fmts = {
        "team_header": fmt({
            "bold": True, "font_size": 13, "align": "center", "valign": "vcenter",
            "bg_color": COLOR_TEAM_BG, "font_color": COLOR_TEAM_TEXT,
            "border": 2, "border_color": COLOR_TEAM_BG,
        }),
        "col_header": fmt({
            "bold": True, "align": "center", "valign": "vcenter",
            "bg_color": COLOR_COL_BG, "font_color": COLOR_COL_TEXT,
            "text_wrap": True,
        }),
        "row_odd": fmt({
            "align": "left", "valign": "vcenter",
            "bg_color": COLOR_ROW_ODD,
        }),
        "row_even": fmt({
            "align": "left", "valign": "vcenter",
            "bg_color": COLOR_ROW_EVEN,
        }),
        "row_odd_center": fmt({
            "align": "center", "valign": "vcenter",
            "bg_color": COLOR_ROW_ODD,
        }),
        "row_even_center": fmt({
            "align": "center", "valign": "vcenter",
            "bg_color": COLOR_ROW_EVEN,
        }),
        "private_odd": fmt({
            "italic": True, "font_color": COLOR_PRIVATE,
            "align": "center", "valign": "vcenter",
            "bg_color": COLOR_ROW_ODD,
        }),
        "private_even": fmt({
            "italic": True, "font_color": COLOR_PRIVATE,
            "align": "center", "valign": "vcenter",
            "bg_color": COLOR_ROW_EVEN,
        }),
        "link_odd": fmt({
            "align": "center", "valign": "vcenter",
            "font_color": COLOR_LINK, "underline": True,
            "bg_color": COLOR_ROW_ODD,
        }),
        "link_even": fmt({
            "align": "center", "valign": "vcenter",
            "font_color": COLOR_LINK, "underline": True,
            "bg_color": COLOR_ROW_EVEN,
        }),
        "flag_ineligible": fmt({
            "bold": True, "align": "center", "valign": "vcenter",
            "font_color": COLOR_INELIGIBLE, "bg_color": COLOR_INELIG_BG,
        }),
        "flag_winner": fmt({
            "align": "center", "valign": "vcenter",
            "bg_color": COLOR_WINNER_BG,
        }),
        "flag_low_games": fmt({
            "align": "center", "valign": "vcenter",
            "bg_color": COLOR_LOW_GAMES_BG,
        }),
        "flag_none_odd": fmt({
            "align": "center", "valign": "vcenter",
            "bg_color": COLOR_ROW_ODD,
        }),
        "flag_none_even": fmt({
            "align": "center", "valign": "vcenter",
            "bg_color": COLOR_ROW_EVEN,
        }),
        "other_teams": fmt({
            "align": "left", "valign": "vcenter",
            "bg_color": COLOR_MULTI_TEAM_BG,
        }),
        "other_teams_blank_odd": fmt({
            "bg_color": COLOR_ROW_ODD,
        }),
        "other_teams_blank_even": fmt({
            "bg_color": COLOR_ROW_EVEN,
        }),
        # Alt account row formats
        "row_alt": fmt({
            "align": "left", "valign": "vcenter",
            "bg_color": COLOR_ALT_BG,
        }),
        "row_alt_center": fmt({
            "align": "center", "valign": "vcenter",
            "bg_color": COLOR_ALT_BG,
        }),
        "private_alt": fmt({
            "italic": True, "font_color": COLOR_PRIVATE,
            "align": "center", "valign": "vcenter",
            "bg_color": COLOR_ALT_BG,
        }),
        "link_alt": fmt({
            "align": "center", "valign": "vcenter",
            "font_color": COLOR_LINK, "underline": True,
            "bg_color": COLOR_ALT_BG,
        }),
        "flag_alt": fmt({
            "align": "center", "valign": "vcenter",
            "bg_color": COLOR_ALT_BG,
        }),
        "flag_alt_ineligible": fmt({
            "bold": True, "align": "center", "valign": "vcenter",
            "font_color": COLOR_INELIGIBLE, "bg_color": COLOR_ALT_BG,
        }),
        "other_teams_alt": fmt({
            "align": "left", "valign": "vcenter",
            "bg_color": COLOR_ALT_BG,
        }),
    }

    # Rank tier formats (one per Dota 2 rank, Herald=1 … Immortal=8)
    for tier, (bg, fg) in RANK_COLORS.items():
        fmts[f"rank_{tier}"] = workbook.add_format({
            **base,
            "align": "center", "valign": "vcenter",
            "bg_color": bg, "font_color": fg, "bold": True,
        })

    return fmts


def _rank_fmt(fmts: dict, rank_num, fallback: str) -> object:
    """Return the rank-colored format for a given raw rank number, or fallback."""
    if rank_num is None:
        return fmts[fallback]
    tier = 8 if int(rank_num) >= 80 else int(rank_num) // 10
    return fmts.get(f"rank_{tier}", fmts[fallback])


def _sort_with_alts(players: list) -> list:
    """
    Reorder players so each main account is immediately followed by any
    on-roster alt accounts (i.e. an alt that is also a registered player
    on this team). Off-roster alts are rendered as sub-rows in _write_team_block
    and don't need sorting here.
    """
    by_pid = {p["player_id"]: p for p in players if p.get("player_id") is not None}
    team_pids = set(by_pid)
    result = []
    added: set = set()

    for p in players:
        pid = p.get("player_id")
        if pid in added:
            continue
        # Defer on-roster alts whose main is also on this team
        if p.get("is_alt") and p.get("main_player_id") in team_pids:
            continue
        result.append(p)
        if pid is not None:
            added.add(pid)
        # Insert on-roster alts of this player right after
        for alt in p.get("alt_accounts", []):
            alt_pid = alt.get("player_id")
            if alt.get("on_roster") and alt_pid in team_pids and alt_pid not in added:
                result.append(by_pid[alt_pid])
                added.add(alt_pid)

    # Append any remaining on-roster alts whose main wasn't on this team
    for p in players:
        pid = p.get("player_id")
        if pid not in added:
            result.append(p)
            if pid is not None:
                added.add(pid)

    return result


def _write_team_block(worksheet, y: int, team_name: str, players: list,
                      fmts: dict) -> int:
    num_cols = len(COLUMNS) - 1

    # Team header row
    worksheet.set_row(y, 20)
    worksheet.merge_range(y, 0, y, num_cols, team_name, fmts["team_header"])
    y += 1

    # Column header row
    worksheet.set_row(y, 18)
    for col_idx, col in enumerate(COLUMNS):
        worksheet.write(y, col_idx, col["header"], fmts["col_header"])
    y += 1

    # Sort so on-roster alts appear immediately after their main
    players = _sort_with_alts(players)
    team_pids = {p.get("player_id") for p in players}

    # Player rows (odd/even index only counts non-alt roster rows)
    main_row_index = 0
    for player in players:
        is_alt = player.get("is_alt", False)
        # Alt rows always use the purple alt palette; others alternate odd/even
        if is_alt:
            row_fmt    = fmts["row_alt"]
            center_fmt = fmts["row_alt_center"]
            priv_fmt   = fmts["private_alt"]
            link_fmt   = fmts["link_alt"]
            other_teams_blank_fmt = fmts["other_teams_alt"]
        else:
            is_odd     = main_row_index % 2 == 0
            main_row_index += 1
            row_fmt    = fmts["row_odd"]        if is_odd else fmts["row_even"]
            center_fmt = fmts["row_odd_center"] if is_odd else fmts["row_even_center"]
            priv_fmt   = fmts["private_odd"]    if is_odd else fmts["private_even"]
            link_fmt   = fmts["link_odd"]       if is_odd else fmts["link_even"]
            other_teams_blank_fmt = fmts["other_teams_blank_odd"] if is_odd else fmts["other_teams_blank_even"]

        steam_id = player["steam_id"]
        stats    = player["APId"]
        worksheet.set_row(y, 16)

        # Player name — indent on-roster alt rows whose main is also on this team
        name = player["name"]
        if is_alt and player.get("main_player_id") in team_pids:
            name = "  ↳ " + name   # "  ↳ PlayerName"
        worksheet.write(y, 0, name, row_fmt)

        if stats == "private":
            worksheet.write(y, 1, "Private", priv_fmt)
            worksheet.write(y, 2, "—",       priv_fmt)
            worksheet.write(y, 3, "—",       priv_fmt)
        else:
            rank_num = stats["rank"]
            rank_str = convert_rank(rank_num) if rank_num else "—"
            lb       = stats["leaderboard"] or "—"
            games    = stats["total_ranked_games_in_3_months"]
            worksheet.write(y, 1, rank_str, _rank_fmt(fmts, rank_num, "row_odd_center" if not is_alt else "row_alt_center"))
            worksheet.write(y, 2, lb,       center_fmt)
            worksheet.write(y, 3, games,    center_fmt)

        # Hyperlinks
        worksheet.write_url(y, 4, f"https://www.opendota.com/players/{steam_id}", link_fmt, "OpenDota")
        worksheet.write_url(y, 5, f"https://www.dotabuff.com/players/{steam_id}", link_fmt, "Dotabuff")

        # Flag column (no multi-team info here)
        flags = []
        if player.get("ineligible"):
            flags.append(f"INELIGIBLE — won {player['won_league']} ({player['won_season']})")
        elif player.get("prev_winner"):
            flags.append(f"Prev winner — {player['won_league']} ({player['won_season']})")

        games = stats.get("total_ranked_games_in_3_months") if isinstance(stats, dict) else None
        if games is not None and games < MIN_RANKED_GAMES:
            flags.append(f"Low games — {games} ranked (90d)")

        if is_alt:
            flags.append(f"Alt of: {player.get('main_player_name', 'unknown')}")

        if player.get("ineligible"):
            flag_fmt = fmts["flag_alt_ineligible"] if is_alt else fmts["flag_ineligible"]
        elif is_alt:
            flag_fmt = fmts["flag_alt"]
        elif games is not None and games < MIN_RANKED_GAMES:
            flag_fmt = fmts["flag_low_games"]
        elif player.get("prev_winner"):
            flag_fmt = fmts["flag_winner"]
        else:
            flag_fmt = fmts["flag_none_odd"] if (main_row_index % 2 == 1) else fmts["flag_none_even"]

        worksheet.write(y, 6, " | ".join(flags), flag_fmt)

        # Other Teams column
        other_teams = player.get("other_teams", [])
        if other_teams:
            text = " | ".join(f"{t} ({get_league_name(s) or s})" for s, t in other_teams)
            worksheet.write(y, 7, text, fmts["other_teams"])
        else:
            worksheet.write(y, 7, "", other_teams_blank_fmt)

        worksheet.write(y, 8, "", center_fmt)
        worksheet.write(y, 9, "", center_fmt)
        y += 1

        # Render off-roster alt accounts as sub-rows with full stats
        for alt in player.get("alt_accounts", []):
            if alt.get("on_roster"):
                continue  # this alt is a proper player row, already sorted above
            alt_steam_id = alt.get("steam_id", "")
            alt_stats    = alt.get("APId")
            worksheet.set_row(y, 16)
            worksheet.write(y, 0, "  ↳ " + alt.get("name", "?"), fmts["row_alt"])

            if alt_stats is None:
                for col in range(1, 4):
                    worksheet.write(y, col, "—", fmts["row_alt_center"])
            elif alt_stats == "private":
                for col in range(1, 4):
                    worksheet.write(y, col, "Private" if col == 1 else "—", fmts["private_alt"])
            else:
                alt_rank_num = alt_stats.get("rank")
                rank_str  = convert_rank(alt_rank_num) if alt_rank_num else "—"
                lb        = alt_stats.get("leaderboard") or "—"
                alt_games = alt_stats.get("total_ranked_games_in_3_months", "—")
                worksheet.write(y, 1, rank_str,  _rank_fmt(fmts, alt_rank_num, "row_alt_center"))
                worksheet.write(y, 2, lb,         fmts["row_alt_center"])
                worksheet.write(y, 3, alt_games,  fmts["row_alt_center"])

            if alt_steam_id:
                worksheet.write_url(y, 4, f"https://www.opendota.com/players/{alt_steam_id}", fmts["link_alt"], "OpenDota")
                worksheet.write_url(y, 5, f"https://www.dotabuff.com/players/{alt_steam_id}", fmts["link_alt"], "Dotabuff")
            else:
                worksheet.write(y, 4, "—", fmts["row_alt_center"])
                worksheet.write(y, 5, "—", fmts["row_alt_center"])

            worksheet.write(y, 6, f"Alt of: {player['name']}", fmts["flag_alt"])
            worksheet.write(y, 7, "", fmts["other_teams_alt"])
            worksheet.write(y, 8, "", fmts["row_alt_center"])
            worksheet.write(y, 9, "", fmts["row_alt_center"])
            y += 1

    return y + 2  # 2-row gap between teams


def create_excel(output_path: str = "excel.xlsx", data_path: str = "databasePlayers.json") -> None:
    seasons = read_json(data_path)

    workbook = xlsxwriter.Workbook(output_path, {"strings_to_urls": False})
    fmts = _make_formats(workbook)

    for season_name, teams in seasons.items():
        # Strip Excel-illegal characters and truncate to the 31-char limit
        sheet_name = _safe_sheet_name(season_name)
        worksheet = workbook.add_worksheet(name=sheet_name)

        # Set column widths once per sheet
        for col_idx, col in enumerate(COLUMNS):
            worksheet.set_column(col_idx, col_idx, col["width"])

        y = 0
        for team_name, players in teams.items():
            if not players:
                continue
            y = _write_team_block(worksheet, y, team_name, players, fmts)

    workbook.close()
    print(f"Saved: {output_path}")
