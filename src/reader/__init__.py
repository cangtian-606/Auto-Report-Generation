from .xlsx import ExcelDataReader
from .yaml import YamlDataReader


def create_reader(path: str):
    if path.endswith(('.yaml', '.yml')):
        return YamlDataReader(path)
    return ExcelDataReader(path)


__all__ = ['ExcelDataReader', 'YamlDataReader', 'create_reader']
