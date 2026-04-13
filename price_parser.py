"""
price_parser.py - Поиск цен ресурсов в Excel-таблице (Сплит-форма)
"""

import pandas as pd
import json
from pathlib import Path
from typing import Optional, Dict, Any


class PriceParser:
    """
    Парсер цен ресурсов из Excel-файла "Сплит-форма".
    Ищет базисные и текущие цены, а также индексы по коду ресурса.
    """

    def __init__(self, excel_path: str):
        """
        Инициализация парсера.

        Args:
            excel_path: Путь к файлу Сплит-форма.xlsx
        """
        self.excel_path = Path(excel_path)
        self.dataframe = None
        self._load()

    def _load(self) -> None:
        """Загрузить Excel-файл и подготовить данные для поиска"""
        if not self.excel_path.exists():
            raise FileNotFoundError(f"Файл не найден: {self.excel_path}")

        # Загружаем весь файл
        self.dataframe = pd.read_excel(self.excel_path, header=None)

        # Строка с заголовками находится на индексе 17
        HEADER_ROW = 17

        # Устанавливаем заголовки
        self.dataframe.columns = self.dataframe.iloc[HEADER_ROW].values
        self.dataframe = self.dataframe.iloc[HEADER_ROW + 1:].reset_index(drop=True)

        # Создаём колонку 'code' для поиска (колонка 0)
        self.dataframe['code'] = self.dataframe.iloc[:, 0].astype(str).str.strip()

        # Базисная цена - колонка 4
        self.dataframe['base_price'] = pd.to_numeric(self.dataframe.iloc[:, 4], errors='coerce')

        # Текущая цена - колонка 7
        self.dataframe['current_price'] = pd.to_numeric(self.dataframe.iloc[:, 7], errors='coerce')

        # Индекс - колонка 8
        self.dataframe['index'] = pd.to_numeric(self.dataframe.iloc[:, 8], errors='coerce')

        # Удаляем строки с пустыми кодами
        self.dataframe = self.dataframe[self.dataframe['code'].notna() & (self.dataframe['code'] != 'nan')]

        print(f"Загружено {len(self.dataframe)} ресурсов из {self.excel_path.name}")

    def lookup(self, code: str) -> Optional[Dict[str, Any]]:
        """
        Найти цены для ресурса по его коду.

        Args:
            code: Код ресурса

        Returns:
            Словарь с ценами {'base_price': ..., 'current_price': ..., 'index': ...}
            или None, если ресурс не найден
        """
        code = str(code).strip()

        if self.dataframe is None or self.dataframe.empty:
            return None

        # Точный поиск
        matches = self.dataframe[self.dataframe['code'] == code]

        # Если нет точного совпадения, ищем частичное (начинается с кода)
        if matches.empty:
            matches = self.dataframe[self.dataframe['code'].str.startswith(code, na=False)]

        if matches.empty:
            return None

        row = matches.iloc[0]

        result = {}
        if pd.notna(row['base_price']):
            result['base_price'] = float(row['base_price'])
        if pd.notna(row['current_price']):
            result['current_price'] = float(row['current_price'])
        if pd.notna(row['index']):
            result['index'] = float(row['index'])

        return result

    def lookup_json(self, code: str) -> str:
        """
        Найти цены и вернуть в формате JSON.

        Args:
            code: Код ресурса

        Returns:
            JSON-строка с ценами или сообщением об ошибке
        """
        result = self.lookup(code)
        if result is None:
            return json.dumps({"error": f"Код {code} не найден", "code": code})
        return json.dumps(result, ensure_ascii=False, indent=2)
