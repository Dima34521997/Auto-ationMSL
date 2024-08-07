def only_dir(path: str) -> str:
    """Из пути к файлу берёт только путь без имени"""
    return path[:path.rfind('\\') + 1]


def only_name(path: str) -> str:
    """Из пути к файлу берёт только имя файла"""
    return path[path.rfind('/') + 1:]


def path_for_win(path: str) -> str:
    """Заменяет слэши в строке с адресом на бекслэши"""
    path = path.replace('/', '\\')
    return path
