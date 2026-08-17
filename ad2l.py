from concurrent.futures import ThreadPoolExecutor, as_completed
from src import database, opendota
from src.excel import create_excel
from src.utils import read_json, write_json

OUTPUT_FILE     = "ad2l_report.xlsx"
DATA_CACHE_FILE = "databasePlayers.json"


def step1_fetch_players() -> None:
    print("Step 1/3 — Fetching player rosters from database...")

    current_ids  = database.get_current_season_ids()
    prev_playoff_ids = database.get_previous_playoff_ids(num_prev_seasons=2)

    players          = database.fetch_players(current_ids)
    seasons          = database.group_by_season_and_team(players)
    previous_winners = database.fetch_previous_winners(prev_playoff_ids)
    database.annotate_winner_flags(seasons, previous_winners)

    all_player_ids = [p["player_id"] for p in players if p.get("player_id") is not None]
    alt_map = database.fetch_alt_accounts(all_player_ids)
    database.annotate_alt_flags(seasons, alt_map)
    database.annotate_multi_team_flags(seasons)
    alt_count = sum(1 for p in players if p.get("is_alt"))

    write_json(seasons, DATA_CACHE_FILE)
    print(f"           Season IDs: {sorted(current_ids)}")
    print(f"           {len(players)} players across {len(seasons)} leagues saved.")
    print(f"           {alt_count} alt account(s) detected.\n")


OPENDOTA_WORKERS = 20  # concurrent requests (safe well under 3000/min premium limit)


def step2_fetch_opendota_stats() -> None:
    print("Step 2/3 — Fetching OpenDota stats...")
    seasons = read_json(DATA_CACHE_FILE)

    # Collect roster players and off-roster alt accounts that still need fetching
    pending_players = []
    pending_alts    = []
    skipped = 0
    for teams in seasons.values():
        for players in teams.values():
            for player in players:
                if "APId" in player:
                    skipped += 1
                else:
                    pending_players.append(player)
                # Off-roster alts attached to this player
                for alt in player.get("alt_accounts", []):
                    if not alt.get("on_roster") and "APId" not in alt and alt.get("steam_id"):
                        pending_alts.append(alt)

    pending = pending_players + pending_alts
    if not pending:
        print(f"           All {skipped} players already cached.\n")
        return

    print(f"           {len(pending_players)} roster + {len(pending_alts)} alt accounts to fetch, "
          f"{skipped} cached — using {OPENDOTA_WORKERS} workers...")
    done = 0

    with ThreadPoolExecutor(max_workers=OPENDOTA_WORKERS) as pool:
        futures = {pool.submit(opendota.get_player_stats, p["steam_id"]): p for p in pending}
        for future in as_completed(futures):
            futures[future]["APId"] = future.result()
            done += 1
            print(f"\r           {done}/{len(pending)} fetched", end="", flush=True)

    print()
    write_json(seasons, DATA_CACHE_FILE)
    print(f"           Done.\n")


def step3_create_excel() -> None:
    print(f"Step 3/3 — Building Excel report...")
    create_excel(output_path=OUTPUT_FILE, data_path=DATA_CACHE_FILE)
    print(f"           Report saved to: {OUTPUT_FILE}\n")


if __name__ == "__main__":
    print("=" * 50)
    print("  AD2L Moderator Report Generator")
    print("=" * 50 + "\n")

    step1_fetch_players()
    step2_fetch_opendota_stats()
    step3_create_excel()

    print("Done!")
