import re
from utils.profile_loader import ProfileLoader


class StructureChecker:
    """
    Проверка структуры курсовой работы.
    """

    LITERATURE_VARIANTS = [
        "литература",
        "список литературы",
        "список использованной литературы",
        "список используемой литературы",
        "список использованных источников",
        "список используемых источников",
        "библиографический список",
        "библиография",
        "использованная литература",
        "использованные источники",
        "список источников",
        "список литературы и источников",
    ]

    APPENDIX_PATTERN = re.compile(
        r"^приложени[ея]\s*(?:№?\s*[а-яa-z0-9])?\s*$",
        re.IGNORECASE
    )

    def __init__(self, found_dict):
        self.found = found_dict
        profile = ProfileLoader.load_structure_profile("fmfhi")

        self.REQUIRED = [s.lower() for s in profile["required_sections"]]
        self.OPTIONAL = [s.lower() for s in profile["optional_sections"]]

    @staticmethod
    def normalize(text: str) -> str:
        return " ".join(text.strip().lower().split())

    def is_application_section(self, text: str) -> bool:
        return bool(self.APPENDIX_PATTERN.match(text.strip().lower()))

    def is_literature_section(self, text: str) -> bool:
        t = self.normalize(text)
        if t in self.LITERATURE_VARIANTS:
            return True
        return any(k in t for k in ["литератур", "источник", "библиограф"])

    def generate_report(self) -> str:
        found_sections = []
        missing_required = []

        for sec in self.REQUIRED:
            if self.found.get(sec, False):
                found_sections.append(sec)
            else:
                missing_required.append(sec)

        for sec in self.OPTIONAL:
            if self.found.get(sec, False):
                found_sections.append(sec)

        report = "📑 ОТЧЁТ ПРОВЕРКИ СТРУКТУРЫ\n"
        report += "=" * 50 + "\n\n"

        report += f"✅ Найденные разделы ({len(found_sections)}):\n"
        for sec in found_sections:
            report += f"  • {sec.capitalize()}\n"

        if missing_required:
            report += f"\n❌ Отсутствующие обязательные разделы ({len(missing_required)}):\n"
            for sec in missing_required:
                report += f"  • {sec.capitalize()}\n"
        else:
            report += "\n✅ Все обязательные разделы присутствуют!\n"

        return report
