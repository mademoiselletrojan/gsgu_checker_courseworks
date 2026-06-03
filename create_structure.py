import os

OUTPUT_FILE = "project_structure.txt"


def build_tree(start_path, prefix=""):
    """
    Рекурсивно строит дерево каталогов в стиле tree /F.
    """
    entries = sorted(os.listdir(start_path))
    tree_lines = []

    for i, entry in enumerate(entries):
        path = os.path.join(start_path, entry)
        connector = "└── " if i == len(entries) - 1 else "├── "

        tree_lines.append(prefix + connector + entry)

        if os.path.isdir(path):
            extension = "    " if i == len(entries) - 1 else "│   "
            tree_lines.extend(build_tree(path, prefix + extension))

    return tree_lines


def save_structure(base_path):
    """
    Создаёт текстовый файл с полной структурой проекта.
    """
    tree = build_tree(base_path)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(base_path + "\n")
        f.write("\n".join(tree))

    print(f"Структура проекта сохранена в {OUTPUT_FILE}")


if __name__ == "__main__":
    base = os.getcwd()
    save_structure(base)
