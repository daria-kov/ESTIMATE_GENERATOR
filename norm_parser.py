"""
norm_parser.py - Парсер XML-файлов с нормативами ГЭСН
Поддерживает ГЭСН, ГЭСНр, ГЭСНм, ГЭСНмр, ГЭСНп
"""

import xml.etree.ElementTree as ET
from decimal import Decimal
from typing import List, Optional, Dict, Tuple, Union
from pathlib import Path
import re

from classes import WorkItem, Resource, ResourceType, Coefficient, AbstractResource
from price_parser import PriceParser
from resource_classifier import ResourceClassifier
from utils import apply_price_data

def _safe_decimal(value: str, default: Decimal = Decimal("0")) -> Decimal:
    """Безопасное преобразование строки в Decimal с обработкой запятых и пробелов"""
    if not value or not str(value).strip():
        return default

    value_str = str(value).strip()

    # Заменяем запятую на точку (российский формат чисел)
    value_str = value_str.replace(',', '.')

    # Удаляем пробелы внутри числа
    value_str = value_str.replace(' ', '')

    # Если строка пустая после очистки
    if not value_str:
        return default

    try:
        return Decimal(value_str)
    except:
        # Если преобразование не удалось, выводим предупреждение
        print(f"  Предупреждение: не удалось преобразовать '{value}' в число, используем 0")
        return default


class NormParser:
    """
    Парсер нормативов ГЭСН из XML-файлов.
    Загружает нормы, ресурсы, абстрактные ресурсы и ссылки на НР/СП.
    """

    # Соответствие префикса кода нормативу и имени файла
    PREFIX_TO_FILE = {
        "ГЭСН": "ГЭСН.xml",
        "ГЭСНр": "ГЭСНр.xml",
        "ГЭСНм": "ГЭСНм.xml",
        "ГЭСНмр": "ГЭСНмр.xml",
        "ГЭСНп": "ГЭСНп.xml",
    }

    # Коды служебных ресурсов, которые не нужно показывать
    SKIP_RESOURCE_CODES = {'1', '2', '4'}

    def __init__(self, xml_folder: Union[str, Path], price_db_path: str = None,
                 fsbc_mash_path: str = None, fsbc_materials_path: str = None):
        """
        Инициализация парсера.

        Args:
            xml_folder: Путь к папке с XML-файлами ГЭСН
            price_db_path: Путь к файлу Сплит-форма.xlsx
            fsbc_mash_path: Путь к файлу ФСБЦ_Маш.xml
            fsbc_materials_path: Путь к файлу ФСБЦ_Мат&Оборуд.xml
        """
        self.xml_folder = Path(xml_folder)
        self.price_parser = PriceParser(price_db_path) if price_db_path else None
        self.classifier = None

        if fsbc_mash_path or fsbc_materials_path:
            self.classifier = ResourceClassifier(fsbc_mash_path, fsbc_materials_path)

        self.roots: Dict[str, ET.Element] = {}
        self._load_all_xml()

    def _load_all_xml(self) -> None:
        """Загрузить все XML-файлы ГЭСН из указанной папки"""
        if not self.xml_folder.exists():
            raise FileNotFoundError(f"Папка с XML не найдена: {self.xml_folder}")

        loaded_count = 0
        for prefix, filename in self.PREFIX_TO_FILE.items():
            filepath = self.xml_folder / filename
            if filepath.exists():
                try:
                    tree = ET.parse(filepath)
                    self.roots[prefix] = tree.getroot()
                    loaded_count += 1
                    print(f"Загружен {filename}")
                except Exception as e:
                    print(f"Ошибка загрузки {filename}: {e}")

        if loaded_count == 0:
            raise FileNotFoundError(f"Не найдено ни одного файла ГЭСН в {self.xml_folder}")

        print(f"Загружено {loaded_count} файлов ГЭСН")

    def _extract_prefix_and_code(self, work_code: str) -> Tuple[str, str]:
        """
        Извлечь префикс и чистый код из полного кода работы.

        Args:
            work_code: Полный код (например, "ГЭСНр57-01-002-01")

        Returns:
            Кортеж (префикс, чистый_код)
        """
        work_code = work_code.strip()
        # Сортируем префиксы по длине, чтобы "ГЭСНр" обработался раньше "ГЭСН"
        sorted_prefixes = sorted(self.PREFIX_TO_FILE.keys(), key=len, reverse=True)

        for prefix in sorted_prefixes:
            if work_code.startswith(prefix):
                return prefix, work_code[len(prefix):]

        return "ГЭСН", work_code

    def _find_work_element(self, work_code: str) -> Tuple[Optional[ET.Element], Optional[str], Optional[str]]:
        """
        Найти элемент Work в XML по коду, а также BeginName из родительского NameGroup.

        Args:
            work_code: Полный код работы

        Returns:
            Кортеж (XML-элемент Work, префикс, begin_name)
        """
        prefix, clean_code = self._extract_prefix_and_code(work_code)

        # Если файл с таким префиксом не загружен, пробуем ГЭСН
        if prefix not in self.roots:
            if "ГЭСН" in self.roots:
                prefix = "ГЭСН"
            else:
                print(f"Файл для префикса {prefix} не загружен. Доступны: {list(self.roots.keys())}")
                return None, None, None

        root = self.roots[prefix]

        # Ищем Work элемент
        for work_elem in root.iter('Work'):
            if work_elem.get('Code') == clean_code:
                # Ищем BeginName в родительском NameGroup
                begin_name = ""
                # Получаем родительский элемент (работает в Python 3.8+)
                parent = None
                for elem in root.iter():
                    if work_elem in elem:
                        parent = elem
                        break
                # Ищем NameGroup среди предков
                if parent is not None:
                    # Поднимаемся вверх до NameGroup
                    current = parent
                    while current is not None:
                        if current.tag == 'NameGroup':
                            begin_name = current.get('BeginName', '')
                            break
                        # Ищем родителя другим способом
                        parent_found = None
                        for elem in root.iter():
                            if current in elem and elem != current:
                                parent_found = elem
                                break
                        current = parent_found
                return work_elem, prefix, begin_name

        return None, None, None

    def _create_driver_resource(self, machine_resource: Resource, quantity: Decimal) -> Optional[Resource]:
        """
        Создать ресурс машиниста для данной машины.

        Args:
            machine_resource: Ресурс-машина
            quantity: Количество машино-часов на единицу

        Returns:
            Ресурс машиниста или None, если машинист не найден
        """
        if not self.classifier:
            return None

        driver_code, rank = self.classifier.get_driver_info(machine_resource.code)
        if not driver_code:
            return None

        driver_name = f"ОТм(ЗТм) Средний разряд машинистов {rank}" if rank else "ОТм(ЗТм)"

        driver_resource = Resource(
            code=driver_code,
            name=driver_name,
            resource_type=ResourceType.LABOR_OPERATOR,
            unit="чел.-ч",
            quantity_per_unit=quantity,
            coefficient=Decimal("1.0"),
            quantity_total=Decimal("0"),
            base_price=None,
            index=Decimal("1.0"),
            current_price=Decimal("0"),
            total_cost=Decimal("0"),
            driver_code=None,
            rank=rank,
        )

        if self.price_parser:
            price_data = self.price_parser.lookup(driver_code)
            apply_price_data(driver_resource, price_data)

        return driver_resource

    def _parse_resources(self, work_element: ET.Element) -> List[Resource]:
        """
        Распарсить ресурсы из XML-элемента Work.

        Args:
            work_element: XML-элемент Work

        Returns:
            Список ресурсов
        """
        resources = []
        resources_elem = work_element.find('Resources')

        if resources_elem is None:
            return resources

        for res_elem in resources_elem.findall('Resource'):
            code = res_elem.get('Code', '')

            # Логируем проблемный ресурс
            quantity_raw = res_elem.get('Quantity', '0')
            if quantity_raw == 'П':
                print(f"  Найден ресурс с Quantity='П': code={code}, name={res_elem.get('EndName', '')}")

            # Пропускаем служебные ресурсы (групповые строки ОТ, ЭМ, М)
            if code in self.SKIP_RESOURCE_CODES:
                continue

            name = res_elem.get('EndName', '')

            # Безопасное преобразование Quantity
            quantity = _safe_decimal(quantity_raw)

            unit = res_elem.get('MeasureUnit', None)

            # Определяем тип ресурса (машина/материал/рабочий)
            resource_type = ResourceType.LABOR
            if self.classifier:
                type_code, _, rank = self.classifier.classify(code)
                if type_code == 'machine':
                    resource_type = ResourceType.EQUIPMENT
                elif type_code == 'material':
                    resource_type = ResourceType.MATERIAL

            # Определяем единицу измерения, если не задана
            if not unit:
                units_map = {
                    ResourceType.LABOR: "чел.-ч",
                    ResourceType.LABOR_OPERATOR: "чел.-ч",
                    ResourceType.EQUIPMENT: "маш.-ч",
                    ResourceType.MATERIAL: "ед.",
                }
                unit = units_map.get(resource_type, "ед.")

            # Извлекаем разряд рабочего из названия
            rank = None
            rank_match = re.search(r'разряд\s+работы?\s*(\d+)', name, re.IGNORECASE)
            if rank_match:
                rank = int(rank_match.group(1))

            resource = Resource(
                code=code,
                name=name,
                resource_type=resource_type,
                unit=unit,
                quantity_per_unit=quantity,
                coefficient=Decimal("1.0"),
                quantity_total=Decimal("0"),
                base_price=None,
                index=Decimal("1.0"),
                current_price=Decimal("0"),
                total_cost=Decimal("0"),
                driver_code=None,
                rank=rank,
            )

            # Загружаем цены из сплит-формы
            if self.price_parser:
                price_data = self.price_parser.lookup(code)
                apply_price_data(resource, price_data)

            resources.append(resource)

            # Для машин создаём ресурс машиниста и добавляем его следом
            if resource_type == ResourceType.EQUIPMENT:
                driver = self._create_driver_resource(resource, quantity)
                if driver:
                    resources.append(driver)

        return resources

    def _parse_abstract_resources(self, work_element: ET.Element,
                               mapping: Dict[str, dict] = None,
                               work_quantity: Decimal = Decimal("1")) -> List[AbstractResource]:
        """
        Распарсить абстрактные ресурсы (бетон, арматура и т.д.)
        Только те, которые указаны в mapping из output.json

        Args:
            work_element: XML-элемент Work
            mapping: Словарь из output.json (AbstractResource)
            work_quantity: Объём работ
        """
        abstract_resources = []
        resources_elem = work_element.find('Resources')

        if resources_elem is None:
            return abstract_resources

        # Если mapping не передан или пустой, не добавляем ни одного абстрактного ресурса
        if not mapping:
            print(f"  Нет маппинга для AbstractResource, пропускаем все абстрактные ресурсы")
            return abstract_resources

        for abs_elem in resources_elem.findall('AbstractResource'):
            code = abs_elem.get('Code', '')
            quantity_raw = abs_elem.get('Quantity', '0')

            # Логируем проблемный ресурс
            if quantity_raw == 'П':
                print(f"  Найден AbstractResource с Quantity='П': code={code}, name={abs_elem.get('Name', '')}")

            name_from_gesn = abs_elem.get('Name', '')
            unit = abs_elem.get('MeasureUnit', '')
            tech_groups = abs_elem.get('TechnologyGroups', '')

            # Пропускаем пустые коды
            if not code:
                continue

            # ✅ ПРОВЕРКА: если код абстрактного ресурса НЕ указан в mapping, пропускаем его
            if code not in mapping:
                print(f"  AbstractResource {code} не найден в output.json, пропускаем")
                continue

            # Пропускаем, если quantity = 0 или 'П'
            if quantity_raw == '0' or quantity_raw == 'П':
                print(f"  AbstractResource {code} имеет quantity='{quantity_raw}', пропускаем")
                continue

            # Безопасное преобразование Quantity
            qty_per_unit = _safe_decimal(quantity_raw)

            # ИТОГОВОЕ КОЛИЧЕСТВО = норма × объём работ
            total_quantity = qty_per_unit * work_quantity

            # Уточнённый код для поиска цены в сплит-форме
            actual_code = code
            refined_codes = mapping[code].get("code", [])
            if refined_codes:
                actual_code = refined_codes[0]
                print(f"  AbstractResource: {code} → {actual_code} (норма {qty_per_unit} × {work_quantity} = {total_quantity} {unit})")
            else:
                print(f"  AbstractResource: {code} не имеет уточнённого кода в mapping, используем исходный {code}")

            # Ищем цену и название в сплит-форме
            resource_name = name_from_gesn  # по умолчанию имя из ГЭСН
            base_price = None
            current_price = None
            price_index = Decimal("1.0")

            if self.price_parser:
                price_data = self.price_parser.lookup(actual_code)
                # Если не нашли, пробуем добавить типовой префикс
                if not price_data:
                    for prefix in ['01.', '02.', '03.', '04.', '08.', '11.']:
                        test_code = f"{prefix}{actual_code}"
                        price_data = self.price_parser.lookup(test_code)
                        if price_data:
                            print(f"Нашёл цену по коду {test_code}")
                            break

                if price_data:
                    # Берём название из сплит-формы, если оно там есть
                    if 'name' in price_data and price_data['name']:
                        resource_name = price_data['name']
                    if 'base_price' in price_data:
                        base_price = Decimal(str(price_data['base_price']))
                    if 'current_price' in price_data:
                        current_price = Decimal(str(price_data['current_price']))
                    if 'index' in price_data:
                        price_index = Decimal(str(price_data['index']))
                else:
                    print(f"  Цена не найдена для {actual_code}")

            abstract_res = AbstractResource(
                code=actual_code,
                name=resource_name,
                unit=unit,
                quantity=total_quantity,
                technology_groups=tech_groups,
                base_price=base_price,
                index=price_index,
                current_price=current_price or (base_price * price_index if base_price else Decimal("0")),
                total_cost=Decimal("0"),
            )

            # Рассчитываем общую стоимость
            if abstract_res.current_price:
                abstract_res.total_cost = total_quantity * abstract_res.current_price

            abstract_resources.append(abstract_res)

        return abstract_resources

    def _parse_nrsp(self, work_element: ET.Element) -> Tuple[Optional[str], Optional[str]]:
        """
        Распарсить ссылки на нормативы НР и СП.

        Args:
            work_element: XML-элемент Work

        Returns:
            Кортеж (ссылка_на_НР, ссылка_на_СП)
        """
        nrsp_elem = work_element.find('NrSp')
        if nrsp_elem is None:
            return None, None
        reason_item = nrsp_elem.find('ReasonItem')
        if reason_item is None:
            return None, None
        return reason_item.get('Nr'), reason_item.get('Sp')

    def _extract_code_from_reason(self, reason: str) -> Optional[str]:
        """
        Извлечь числовой код из ссылки на норматив.

        Args:
            reason: Ссылка вида "Пр/812-006.0-1"

        Returns:
            Числовой код (например, "006.0")
        """
        if not reason:
            return None
        match = re.search(r'(\d+\.\d+)', reason)
        return match.group(1) if match else None

    def _calculate_components_base(self, resources: List[Resource]) -> Dict[str, Decimal]:
        """
        Рассчитать базовые компоненты прямых затрат (ОТ, ЭМ, ОТм, М).

        Args:
            resources: Список ресурсов

        Returns:
            Словарь с суммами по типам ресурсов
        """
        components = {"OT": Decimal("0"), "EM": Decimal("0"), "OTm": Decimal("0"), "M": Decimal("0")}
        type_map = {
            ResourceType.LABOR: "OT",
            ResourceType.EQUIPMENT: "EM",
            ResourceType.LABOR_OPERATOR: "OTm",
            ResourceType.MATERIAL: "M",
        }
        for res in resources:
            key = type_map.get(res.resource_type)
            if key:
                components[key] += res.quantity_per_unit
        return components

    def parse_work_by_code(self, work_code: str, quantity: Decimal = Decimal("1"),
                           abstract_resource_mapping: Dict[str, dict] = None) -> Optional[WorkItem]:
        """
        Основной метод: загрузить норматив по коду и создать WorkItem.

        Args:
            work_code: Полный код работы (например, "ГЭСНр57-01-002-01")
            quantity: Объём работ
            abstract_resource_mapping: Словарь уточнённых кодов абстрактных ресурсов
                                       из output.json (общий_код -> {code: [уточнённый_код]})

        Returns:
            WorkItem или None, если норматив не найден
        """
        work_element, file_prefix, begin_name = self._find_work_element(work_code)

        if work_element is None:
            print(f"Норма не найдена: {work_code}")
            return None

        _, clean_code = self._extract_prefix_and_code(work_code)

        # Получаем EndName
        end_name = work_element.get("EndName", "")

        # Формируем полное наименование: BeginName + EndName
        # Если BeginName не пустой, объединяем с EndName
        if begin_name and end_name:
            # Убираем лишние пробелы и правильно форматируем
            begin_name = begin_name.rstrip()
            end_name = end_name.lstrip()
            if begin_name.endswith(':'):
                full_name = f"{begin_name} {end_name}"
            elif begin_name.endswith(' ') or end_name.startswith(' '):
                full_name = f"{begin_name}{end_name}"
            else:
                full_name = f"{begin_name} {end_name}"
        elif begin_name:
            full_name = begin_name
        else:
            full_name = end_name

        work_item = WorkItem(
            row_num=0,
            code=f"{file_prefix}{clean_code}",
            name=full_name,
            unit=work_element.get("MeasureUnit", "ед."),
            quantity=quantity,
        )

        # Парсим ресурсы и абстрактные ресурсы (с передачей mapping)
        work_item.resources = self._parse_resources(work_element)
        work_item.abstract_resources = self._parse_abstract_resources(
            work_element,
            abstract_resource_mapping,
            quantity  # ← добавляем объём работ
        )

        # Парсим ссылки на НР/СП
        overhead_ref, profit_ref = self._parse_nrsp(work_element)
        work_item.overhead_ref = overhead_ref
        work_item.profit_ref = profit_ref
        work_item.nr_code = self._extract_code_from_reason(overhead_ref)
        work_item.sp_code = self._extract_code_from_reason(profit_ref)
        work_item.components_base = self._calculate_components_base(work_item.resources)

        print(f"Загружена норма: {work_item.code} - {work_item.name[:50]}...")

        return work_item

    def get_all_work_codes(self) -> List[str]:
        """
        Получить список всех кодов работ во всех загруженных файлах.

        Returns:
            Список кодов
        """
        all_codes = []
        for prefix, root in self.roots.items():
            for work_elem in root.iter('Work'):
                code = work_elem.get('Code')
                if code:
                    all_codes.append(f"{prefix}{code}")
        return sorted(all_codes)
