"""
resource_classifier.py - Классификация ресурсов по данным ФСБЦ
Определяет, является ли ресурс машиной, материалом или трудовым ресурсом,
а также устанавливает связи между машинами и машинистами.
"""

import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional, Set, Dict, Tuple


class ResourceClassifier:
    """
    Классификатор ресурсов на основе файлов ФСБЦ.
    Позволяет определить тип ресурса (машина/материал/рабочий)
    и получить информацию о машинисте для машин.
    """

    def __init__(self, machines_path: str = None, materials_path: str = None):
        """
        Инициализация классификатора.

        Args:
            machines_path: Путь к файлу ФСБЦ_Маш.xml (классификация машин)
            materials_path: Путь к файлу ФСБЦ_Мат&Оборуд.xml (классификация материалов)
        """
        self.machines_path = Path(machines_path) if machines_path else None
        self.materials_path = Path(materials_path) if materials_path else None

        # Множества кодов для быстрого поиска
        self.machine_codes: Set[str] = set()      # Коды машин
        self.material_codes: Set[str] = set()     # Коды материалов

        # Связи машина → машинист
        self.machine_to_driver: Dict[str, str] = {}   # код машины → код машиниста
        self.driver_rank: Dict[str, int] = {}         # код машиниста → разряд

        self._load()

    def _load(self) -> None:
        """Загрузить и распарсить оба файла ФСБЦ"""
        if self.machines_path and self.machines_path.exists():
            self._load_machines()
        elif self.machines_path:
            print(f"Файл с машинами не найден: {self.machines_path}")

        if self.materials_path and self.materials_path.exists():
            self._load_materials()
        elif self.materials_path:
            print(f"Файл с материалами не найден: {self.materials_path}")

        if self.machine_codes or self.material_codes:
            print(f"ФСБЦ: загружено {len(self.machine_codes)} машин и {len(self.material_codes)} материалов")

    def _load_machines(self) -> None:
        """Загрузить файл ФСБЦ_Маш.xml с классификацией машин"""
        tree = ET.parse(self.machines_path)
        root = tree.getroot()

        for resource in root.findall(".//Resource"):
            code = resource.get("Code", "")
            if not code:
                continue

            self.machine_codes.add(code)

            # Ищем информацию о машинисте
            price_elem = resource.find(".//Price")
            if price_elem is not None:
                driver_code = price_elem.get("DriverCode")
                if driver_code:
                    self.machine_to_driver[code] = driver_code

                    # Извлекаем разряд машиниста
                    machinist_cat = price_elem.get("MachinistCategory", "")
                    if machinist_cat:
                        try:
                            rank = int(float(machinist_cat))
                            self.driver_rank[driver_code] = rank
                        except ValueError:
                            pass

    def _load_materials(self) -> None:
        """Загрузить файл ФСБЦ_Мат&Оборуд.xml с классификацией материалов"""
        tree = ET.parse(self.materials_path)
        root = tree.getroot()

        for resource in root.findall(".//Resource"):
            code = resource.get("Code", "")
            if code:
                self.material_codes.add(code)

    def classify(self, code: str) -> Tuple[str, Optional[str], Optional[int]]:
        """
        Определить тип ресурса по его коду.

        Args:
            code: Код ресурса

        Returns:
            Кортеж (тип, код_машиниста, разряд_машиниста)
            тип может быть: 'machine', 'material', 'labor'
        """
        if code in self.machine_codes:
            driver_code = self.machine_to_driver.get(code)
            rank = self.driver_rank.get(driver_code) if driver_code else None
            return ('machine', driver_code, rank)
        elif code in self.material_codes:
            return ('material', None, None)
        return ('labor', None, None)

    def get_driver_info(self, machine_code: str) -> Tuple[Optional[str], Optional[int]]:
        """
        Получить информацию о машинисте для машины.

        Args:
            machine_code: Код машины

        Returns:
            Кортеж (код_машиниста, разряд_машиниста)
        """
        driver_code = self.machine_to_driver.get(machine_code)
        rank = self.driver_rank.get(driver_code) if driver_code else None
        return driver_code, rank

    # Методы для обратной совместимости
    def is_machine(self, code: str) -> bool:
        return code in self.machine_codes

    def is_material(self, code: str) -> bool:
        return code in self.material_codes

    def get_driver_code(self, machine_code: str) -> Optional[str]:
        return self.machine_to_driver.get(machine_code)

    def get_driver_rank(self, driver_code: str) -> Optional[int]:
        return self.driver_rank.get(driver_code)
