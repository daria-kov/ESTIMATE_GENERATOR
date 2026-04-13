"""
normative_rates_parser.py - Парсер файла с нормативами НР и СП
(Рекомендуемая привязка нормативов накладных расходов и сметной прибыли)
"""

import pandas as pd
from pathlib import Path
from decimal import Decimal
from typing import Dict, Optional, List, Tuple
from dataclasses import dataclass
import re


@dataclass
class OverheadRate:
    """Норматив накладных расходов (НР) для разных регионов"""
    code: str
    name: str
    territory: Decimal   # Для территории РФ
    mprks: Decimal       # Для МПРКС (районы Крайнего Севера)
    rks: Decimal         # Для РКС (приравненные к Крайнему Северу)

    def get_rate(self, region_type: str = "territory") -> Decimal:
        """Получить норматив для указанного типа региона"""
        if region_type == "territory":
            return self.territory
        elif region_type == "mprks":
            return self.mprks
        elif region_type == "rks":
            return self.rks
        return self.territory


@dataclass
class ProfitRate:
    """Норматив сметной прибыли (СП)"""
    code: str
    name: str
    percent: Decimal


@dataclass
class WorkTypeMapping:
    """Привязка вида работ к нормативам НР и СП"""
    code: str
    name: str
    gesn_sections: List[str]   # Коды сборников ГЭСН
    overhead: Optional[OverheadRate]
    profit: Optional[ProfitRate]


class NormativeRatesParser:
    """
    Парсер нормативов НР и СП из Excel-файла ФГИС ЦС.
    Позволяет получить проценты НР и СП по шифру норматива.
    """

    def __init__(self, excel_path: str):
        """
        Инициализация парсера.

        Args:
            excel_path: Путь к файлу с нормативами
        """
        self.excel_path = Path(excel_path)
        self.dataframe = None
        self.overhead_rates: Dict[str, OverheadRate] = {}
        self.profit_rates: Dict[str, ProfitRate] = {}
        self.mappings: List[WorkTypeMapping] = []
        self._load()

    def _parse_percent(self, value_str: str) -> Decimal:
        """
        Преобразует строку с процентом в Decimal.
        Обрабатывает случаи:
        - "102" -> 102
        - "102/113" -> 102 (берёт первое значение)
        - "nan" -> 0
        - "" -> 0

        Args:
            value_str: Строковое представление процента

        Returns:
            Decimal значение процента
        """
        if pd.isna(value_str):
            return Decimal("0")

        value_str = str(value_str).strip()

        if not value_str or value_str == "nan" or value_str == "":
            return Decimal("0")

        # Если есть дробь, берём первое значение (для территории)
        if '/' in value_str:
            parts = value_str.split('/')
            value_str = parts[0].strip()

        # Убираем всё кроме цифр и точки
        value_str = re.sub(r'[^\d.]', '', value_str)

        if not value_str:
            return Decimal("0")

        try:
            return Decimal(value_str)
        except:
            return Decimal("0")

    def _load(self) -> None:
        """Загрузить и распарсить Excel-файл"""
        if not self.excel_path.exists():
            print(f"Файл с нормативами НР/СП не найден: {self.excel_path}")
            return

        self.dataframe = pd.read_excel(self.excel_path, header=None)

        # Ищем строку с заголовками (обычно до 30 строки)
        header_row = None
        for idx, row in self.dataframe.iterrows():
            if idx < 30:
                first_cell = str(row.iloc[0]) if pd.notna(row.iloc[0]) else ""
                if "№ п/п" in first_cell or "Шифр" in first_cell:
                    header_row = idx
                    break

        if header_row is None:
            print("Не удалось найти заголовки в файле")
            return

        # Устанавливаем заголовки
        self.dataframe.columns = self.dataframe.iloc[header_row].values
        self.dataframe = self.dataframe.iloc[header_row + 1:].reset_index(drop=True)
        self.dataframe = self.dataframe.fillna("")

        for idx, row in self.dataframe.iterrows():
            code = str(row.iloc[1]).strip() if len(row) > 1 else ""
            name = str(row.iloc[2]).strip() if len(row) > 2 else ""
            gesn_info = str(row.iloc[3]).strip() if len(row) > 3 else ""

            # Пропускаем разделители и пустые строки
            if not code or code in ["I", "II", "III", "IV", "V", "VI", "VII"]:
                continue

            territory_str = str(row.iloc[4]).strip() if len(row) > 4 else "0"
            mprks_str = str(row.iloc[5]).strip() if len(row) > 5 else "0"
            rks_str = str(row.iloc[6]).strip() if len(row) > 6 else "0"
            profit_str = str(row.iloc[7]).strip() if len(row) > 7 else "0"

            territory = self._parse_percent(territory_str)
            mprks = self._parse_percent(mprks_str)
            rks = self._parse_percent(rks_str)
            profit = self._parse_percent(profit_str)

            overhead = None
            if territory != 0 or mprks != 0 or rks != 0:
                overhead = OverheadRate(
                    code=code,
                    name=name,
                    territory=territory,
                    mprks=mprks,
                    rks=rks,
                )
                self.overhead_rates[code] = overhead

            profit_rate = None
            if profit != 0:
                profit_rate = ProfitRate(
                    code=code,
                    name=name,
                    percent=profit,
                )
                self.profit_rates[code] = profit_rate

            if overhead or profit_rate:
                mapping = WorkTypeMapping(
                    code=code,
                    name=name,
                    gesn_sections=self._parse_gesn_sections(gesn_info),
                    overhead=overhead,
                    profit=profit_rate,
                )
                self.mappings.append(mapping)

        print(f"Загружено {len(self.overhead_rates)} нормативов НР и {len(self.profit_rates)} нормативов СП")

    def _parse_gesn_sections(self, gesn_info: str) -> List[str]:
        """
        Парсит информацию о сборниках ГЭСН в список кодов.

        Args:
            gesn_info: Строка с информацией о сборниках

        Returns:
            Список кодов сборников
        """
        if not gesn_info or gesn_info == "nan":
            return []

        sections = []
        for part in gesn_info.replace('\n', ',').split(','):
            part = part.strip()
            if part and part != "nan":
                sections.append(part)
        return sections

    def get_by_code(self, code: str, region_type: str = "territory") -> Tuple[Decimal, Decimal]:
        """
        Получить НР и СП по шифру норматива.

        Args:
            code: Шифр норматива (например, "006.0")
            region_type: Тип региона ("territory", "mprks", "rks")

        Returns:
            Кортеж (НР_процент, СП_процент)
        """
        overhead = self.overhead_rates.get(code)
        profit = self.profit_rates.get(code)

        nr = overhead.get_rate(region_type) if overhead else Decimal("0")
        sp = profit.percent if profit else Decimal("0")

        return nr, sp

    def find_by_gesn_code(self, gesn_code: str) -> Optional[WorkTypeMapping]:
        """
        Найти вид работ по коду ГЭСН.

        Args:
            gesn_code: Код ГЭСН

        Returns:
            Mapping или None
        """
        collection = gesn_code[:2] if gesn_code else ""

        for mapping in self.mappings:
            for section in mapping.gesn_sections:
                if collection and collection in section:
                    return mapping
                if gesn_code in section:
                    return mapping
        return None

    def get_nrsp_for_work(self, work_code: str, region_type: str = "territory") -> Tuple[Decimal, Decimal, str]:
        """
        Получить проценты НР и СП для кода работы.

        Args:
            work_code: Код работы ГЭСН
            region_type: Тип региона

        Returns:
            Кортеж (НР_процент, СП_процент, шифр_норматива)
        """
        mapping = self.find_by_gesn_code(work_code)
        if mapping:
            nr = mapping.overhead.get_rate(region_type) if mapping.overhead else Decimal("0")
            sp = mapping.profit.percent if mapping.profit else Decimal("0")
            return nr, sp, mapping.code
        return Decimal("0"), Decimal("0"), ""

    def print_info(self) -> None:
        """Вывести информацию о загруженных нормативах (для отладки)"""
        print("\n" + "="*60)
        print("НОРМАТИВЫ НАКЛАДНЫХ РАСХОДОВ (НР):")
        print("="*60)
        for code, rate in list(self.overhead_rates.items())[:10]:
            print(f"  {code}: {rate.name[:50]}...")
            print(f"      Территория: {rate.territory}%, МПРКС: {rate.mprks}%, РКС: {rate.rks}%")

        print("\n" + "="*60)
        print("НОРМАТИВЫ СМЕТНОЙ ПРИБЫЛИ (СП):")
        print("="*60)
        for code, rate in list(self.profit_rates.items())[:10]:
            print(f"  {code}: {rate.name[:50]}... → {rate.percent}%")


if __name__ == "__main__":
    excel_path = Path(__file__).parent / "ФСНБ-2022  Привязка НР (812_пр) и СП (774_пр) ФСНБ-2022 Доп.17.xlsx"

    if excel_path.exists():
        parser = NormativeRatesParser(str(excel_path))
        parser.print_info()
    else:
        print(f"Файл не найден: {excel_path}")
