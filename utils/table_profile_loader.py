import json
import os


class TableProfileLoader:
    @staticmethod
    def load_table_profile():
        # ВАЖНО: папка у тебя называется table_standarts
        path = "data/table_standarts/table_format.json"
        if not os.path.exists(path):
            raise FileNotFoundError(f"Профиль таблиц не найден: {path}")
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
