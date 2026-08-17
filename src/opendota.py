import time
import requests
from .config import OPEN_DOTA_API_KEY

LOBBY_RANKED           = 7
RANKED_DATE_WINDOW_DAYS = 90
MAX_RETRIES            = 5
DEFAULT_RETRY_WAIT     = 60  # seconds to wait if no Retry-After header


def _get(url: str) -> dict | list:
    """GET with automatic retry on 429 rate-limit responses."""
    for attempt in range(MAX_RETRIES):
        response = requests.get(url)

        if response.status_code == 429:
            wait = int(response.headers.get("Retry-After", DEFAULT_RETRY_WAIT))
            print(f"\n  [OpenDota] Rate limited — waiting {wait}s before retrying...")
            time.sleep(wait)
            continue

        response.raise_for_status()
        return response.json()

    raise RuntimeError(f"Failed after {MAX_RETRIES} retries: {url}")


def _get_player_profile(steam_id) -> dict | str:
    url = f"https://api.opendota.com/api/players/{steam_id}?api_key={OPEN_DOTA_API_KEY}"
    try:
        data = _get(url)
        if data["profile"]["fh_unavailable"]:
            return "private"
        return {
            "name":        data["profile"]["personaname"],
            "rank":        data["rank_tier"],
            "leaderboard": data["leaderboard_rank"],
        }
    except Exception:
        return "private"


def get_player_stats(steam_id) -> dict | str:
    player = _get_player_profile(steam_id)
    if player == "private":
        return "private"

    url = (
        f"https://api.opendota.com/api/players/{steam_id}/matches"
        f"?lobby_type={LOBBY_RANKED}&date={RANKED_DATE_WINDOW_DAYS}&api_key={OPEN_DOTA_API_KEY}"
    )
    try:
        matches = _get(url)
        player["total_ranked_games_in_3_months"] = len(matches)
    except Exception:
        player["total_ranked_games_in_3_months"] = 0

    return player
