import re
from docx import Document


class DocxParser:
    """
    Парсер документов Word (.docx) для проверки курсовых работ.
    Извлекает структуру документа и ключевые разделы.
    """

    # Варианты названий списка литературы
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

    # Паттерн главы
    CHAPTER_PATTERN = re.compile(
        r"^глава\s*№?\s*\d+[\.\:\-\s]*.*$",
        re.IGNORECASE
    )

    # Подглавы вида 1.1, 1.2, 2.3, 2.3.1
    SUBCHAPTER_PATTERN = re.compile(
        r"^\s*(?:§\s*)?\d+\.\d+(?:\.\d+)?(?:\s|\.|:|-|$)",
        re.IGNORECASE
    )

    # Приложения
    APPENDIX_PATTERN = re.compile(
        r"^приложени[ея]\s*(?:№?\s*[а-яa-z0-9])?\s*$",
        re.IGNORECASE
    )

    def __init__(self, file_path):
        self.doc = Document(file_path)

    def extract_full_paragraphs(self):
        return [p.text.strip() for p in self.doc.paragraphs]

    def extract_paragraphs_from_introduction(self):
        paragraphs = self.extract_full_paragraphs()
        lower = [p.lower() for p in paragraphs]

        for i, text in enumerate(lower):
            if text.strip() == "введение":
                return paragraphs[i:]

        return paragraphs

    def extract_key_headings(self):
        paragraphs = self.extract_paragraphs_from_introduction()
        lower = [p.lower() for p in paragraphs]

        found = {
            "введение": False,
            "глава 1": False,
            "глава 2": False,
            "глава 3": False,
            "глава 4": False,
            "глава 5": False,
            "заключение": False,
            "литература": False,
            "приложения": False,
        }

        for text in lower:
            clean = text.strip()

            # Введение
            if clean == "введение":
                found["введение"] = True

            # Главы
            if self.CHAPTER_PATTERN.match(clean):
                m = re.search(r"\d+", clean)
                if m:
                    num = m.group(0)
                    key = f"глава {num}"
                    if key in found:
                        found[key] = True

            # Заключение
            if clean == "заключение":
                found["заключение"] = True

            # Литература
            if clean in self.LITERATURE_VARIANTS:
                found["литература"] = True

            # Приложения
            if self.APPENDIX_PATTERN.match(clean):
                found["приложения"] = True

        return found, paragraphs
