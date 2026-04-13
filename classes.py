from dataclasses import dataclass, field
from decimal import Decimal
from datetime import date
from typing import List, Optional, Dict, Any
from enum import Enum


class ResourceType(Enum):
    """Типы ресурсов"""
    LABOR = "OT"           # Оплата труда рабочих (ЗТ)
    EQUIPMENT = "EM"       # Эксплуатация машин
    LABOR_OPERATOR = "OTm" # Оплата труда машинистов (ЗТм)
    MATERIAL = "M"         # Материалы


class RowType(Enum):
    """Типы строк в таблице сметы"""
    SECTION_HEADER = "section_header"
    WORK_MAIN = "work_main"
    COEFFICIENT = "coefficient"
    RESOURCE_GROUP = "resource_group"      # ОТ, ЭМ, ОТм, М
    RESOURCE_DETAIL = "resource_detail"    # Конкретный ресурс (код + название)
    SUBTOTAL_DIRECT = "subtotal_direct"
    SUBTOTAL_FOT = "subtotal_fot"
    OVERHEAD = "overhead"
    PROFIT = "profit"
    POSITION_TOTAL = "position_total"


@dataclass
class EstimateMetadata:
    """Метаданные сметы (шапка Excel)"""
    software_name: str = "Пример ЛС БИМ для ФГИС ЦС"

    # Нормативные редакции
    normative_editions: List[str] = field(default_factory=lambda: [
        "Приказ Минстроя России от 26.12.2019 № 876/пр",
        "Приказ Минстроя России от 04.08.2020 № 421/пр",
        "Приказ Минстроя России от 21.12.2020 № 812/пр",
        "Приказ Минстроя России от 11.12.2020 № 774/пр",
    ])

    # Дополнения и изменения
    amendments: List[str] = field(default_factory=lambda: [
        "Приказ Минстроя России от 30.03.2020 № 172/пр",
        "Приказ Минстроя России от 01.06.2020 № 294/пр",
        "Приказ Минстроя России от 30.06.2020 № 352/пр",
        "Приказ Минстроя России от 20.10.2020 № 636/пр",
        "Приказ Минстроя России от 09.02.2021 № 51/пр",
        "Приказ Минстроя России от 24.05.2021 № 321/пр",
        "Приказ Минстроя России от 24.06.2021 № 408/пр",
        "Приказ Минстроя России от 07.07.2022 № 557/пр",
        "Приказ Минстроя России от 22.04.2022 № 317/пр",
    ])

    # Индексы
    indices_letter: str = "Письмо Минстроя России от ХХ.ХХ.ХХХХ № ХХХХХХХХХ"

    # Оплата труда
    labor_act: str = "Постановление Правительства ХХХХХХХ области от ХХ.ХХ.ХХХХ № ХХХ"

    # Территория
    region: str = "ХХХХХХХ"
    zone: str = "ХХХХХХХ"

    # Объект
    project_address: str = "Строительство объекта капитального строительства по адресу: г. N……"
    project_name: str = "(наименование стройки)"
    object_name: str = "Объект капитального строительства"
    object_detail: str = "(наименование объекта капитального строительства)"

    # Смета
    estimate_number: str = "ЛС-02-01-01"
    estimate_title: str = "Пример локального сметного расчета (сметы)_базисно-индексный метод с применением индексов по элементам прямых затрат"
    work_description: str = "(наименование работ и затрат)"

    # Метод
    method: str = "базисно-индексный"
    basis: str = "ведомость объемов работ № ВОР 02-01-01"
    basis_description: str = "(проектная и (или) иная техническая документация)"

    # Уровни цен
    price_level_current: str = "II кв. 2021 г."
    price_level_base: str = "01.01.2000"

    # Индексы пересчета
    indices: Dict[str, Decimal] = field(default_factory=lambda: {
        "OT": Decimal("23.57"),
        "EM": Decimal("1.0"),
        "OTm": Decimal("23.57"),
        "M": Decimal("1.0"),
    })


@dataclass
class Coefficient:
    """Коэффициент к расценке (421/пр и др.)"""
    code: str                    # "421/пр_2020_прил.10_т.5_п.3_гр.3"
    description: str             # Полное описание условия
    values: Dict[str, Decimal]   # {"ОЗП": 1.15, "ЭМ": 1.15, "ЗПМ": 1.15, "ТЗ": 1.15, "ТЗМ": 1.15}


@dataclass
class Resource:
    """Ресурс в составе расценки"""
    code: str                    # "1-100-30", "91.05.05-015"
    name: str                    # "Средний разряд работы 3,0"
    resource_type: ResourceType  # OT, EM, OTm, M
    unit: str                    # "чел.-ч", "маш.-ч", "м3", "кг"

    # Нормы на единицу
    quantity_per_unit: Decimal = Decimal("0")
    coefficient: Decimal = Decimal("1.0")
    quantity_total: Decimal = Decimal("0")  # С учётом коэффициентов и объёма

    # Цены
    base_price: Optional[Decimal] = None  # Базисная цена
    index: Decimal = Decimal("1.0")       # Индекс пересчета
    current_price: Decimal = Decimal("0") # Текущая цена
    total_cost: Decimal = Decimal("0")    # Стоимость строки

    # Для машин — код соответствующего машиниста (если есть)
    driver_code: Optional[str] = None

    # Разряд рабочего или машиниста
    rank: Optional[int] = None


@dataclass
class AbstractResource:
    """Абстрактный ресурс из ГЭСН (бетон, арматура, опалубка)"""
    code: str                    # "04.1.02.05"
    name: str                    # "Смеси бетонные тяжелого бетона"
    unit: str                    # "м3"
    quantity: Decimal            # 1.015
    technology_groups: str = ""  # "17.01.022"

    # Цены и расчёты (как у Resource)
    base_price: Optional[Decimal] = None
    index: Decimal = Decimal("1.0")
    current_price: Decimal = Decimal("0")
    total_cost: Decimal = Decimal("0")

    # Для экспорта — номер позиции
    row_num: int = 0


@dataclass
class WorkItem:
    """Позиция сметы (ГЭСН/ФСБЦ)"""
    row_num: int
    code: str                    # "ГЭСН06-01-004-04"
    name: str                    # Наименование
    unit: str                    # Единица измерения
    quantity: Decimal            # Объём работ

    # Компоненты прямых затрат (базисный уровень, на единицу)
    components_base: Dict[str, Decimal] = field(default_factory=dict)

    # Ресурсы детально
    resources: List[Resource] = field(default_factory=list)

    # Абстрактные ресурсы (опалубка, бетон, арматура)
    abstract_resources: List[AbstractResource] = field(default_factory=list)

    # Коэффициенты
    coefficients: List[Coefficient] = field(default_factory=list)

    # Ссылки на НР и СП
    overhead_ref: Optional[str] = None  # "Пр/812-006.0-1"
    profit_ref: Optional[str] = None    # "Пр/774-006.0"

    # Расчетные итоги
    direct_costs: Decimal = Decimal("0")
    fot: Decimal = Decimal("0")
    overhead_amount: Decimal = Decimal("0")
    profit_amount: Decimal = Decimal("0")
    position_total: Decimal = Decimal("0")

    # Для ФСБЦ (упрощённая позиция)
    is_fsbc: bool = False
    fsbc_base_price: Optional[Decimal] = None
    fsbc_index: Decimal = Decimal("1.0")


@dataclass
class SectionSummary:
    """Итоговая строка раздела"""
    row_type: RowType
    name: str
    reference: Optional[str] = None
    percentage: Optional[Decimal] = None
    quantity_base: Optional[Decimal] = None
    coefficient: Optional[Decimal] = None
    base_price: Optional[Decimal] = None
    index: Optional[Decimal] = None
    current_price: Optional[Decimal] = None
    total_cost: Decimal = Decimal("0")


@dataclass
class Section:
    """Раздел сметы"""
    id: int
    name: str
    note: Optional[str] = None
    items: List[WorkItem] = field(default_factory=list)
    summaries: List[SectionSummary] = field(default_factory=list)

    @property
    def full_name(self) -> str:
        return f"Раздел {self.id}. {self.name}"


@dataclass
class EstimateSummary:
    """Сводные показатели сметы"""
    # Стоимость (тыс. руб.)
    total_cost: Decimal = Decimal("0")
    construction_works: Decimal = Decimal("0")
    installation_works: Decimal = Decimal("0")
    equipment: Decimal = Decimal("0")
    other_costs: Decimal = Decimal("0")

    # Оплата труда (тыс. руб.)
    labor_cost_workers: Decimal = Decimal("0")

    # Трудовые ресурсы (чел.-ч)
    labor_hours_workers: Decimal = Decimal("0")
    labor_hours_operators: Decimal = Decimal("0")


@dataclass
class Estimate:
    """Корневой объект сметы"""
    metadata: EstimateMetadata = field(default_factory=EstimateMetadata)
    summary: EstimateSummary = field(default_factory=EstimateSummary)
    sections: List[Section] = field(default_factory=list)

    def add_section(self, section: Section) -> None:
        section.id = len(self.sections) + 1
        self.sections.append(section)
