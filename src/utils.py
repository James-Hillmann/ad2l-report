import json

# League skill order — higher number = higher skill tier
LEAGUE_RANK = {
    "Explorer": 1, "Voyager": 2, "Challenger": 3, "Warrior": 4,
    "Conqueror": 5, "Champion": 6, "Heroic": 7, "Aegis": 8,
}

def get_league_name(season_title: str) -> str | None:
    for name in LEAGUE_RANK:
        if name in season_title:
            return name
    return None

RANKS = {
    1: "Herald",
    2: "Guardian",
    3: "Crusader",
    4: "Archon",
    5: "Legend",
    6: "Ancient",
    7: "Divine",
    8: "Immortal",
}

DATA_FILE = "databasePlayers.json"


def convert_rank(rank_num: int) -> str:
    rank_num = int(rank_num)
    if rank_num >= 80:
        return "Immortal"
    rank = rank_num // 10
    sub_level = rank_num % 10
    return f"{RANKS[rank]} {sub_level}"


def read_json(path: str = DATA_FILE) -> dict:
    with open(path, "r") as f:
        return json.load(f)


def write_json(data: dict, path: str = DATA_FILE) -> None:
    with open(path, "w") as f:
        json.dump(data, f, indent=4)
