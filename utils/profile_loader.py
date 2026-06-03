import json
import os

class ProfileLoader:
    @staticmethod
    def load_format_profile(faculty):
        path = f"data/profiles/{faculty}_format.json"
        if not os.path.exists(path):
            raise FileNotFoundError(f"Форматный профиль не найден: {path}")
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)

    @staticmethod
    def load_structure_profile(faculty):
        path = f"data/profiles/{faculty}_structure.json"
        if not os.path.exists(path):
            raise FileNotFoundError(f"Структурный профиль не найден: {path}")
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
