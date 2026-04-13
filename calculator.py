"""
calculator.py - Расчёт всех сумм в смете
"""

from decimal import Decimal
from typing import Dict, List
from classes import WorkItem, Resource, ResourceType, Section, SectionSummary, RowType


class Calculator:
    """
    Калькулятор сметы.
    Выполняет расчёт прямых затрат, ФОТ, НР и СП для каждой позиции и раздела.
    """

    def __init__(self, indices: Dict[str, Decimal] = None):
        """
        Инициализация калькулятора с индексами пересчёта.

        Args:
            indices: Словарь с индексами для разных типов ресурсов
                     (OT, EM, OTm, M). По умолчанию используются значения из примера.
        """
        self.indices = indices or {
            "OT": Decimal("23.57"),
            "EM": Decimal("1.0"),
            "OTm": Decimal("23.57"),
            "M": Decimal("1.0"),
        }

    def calculate_section(self, section: Section) -> None:
        """
        Рассчитать все позиции в разделе и заполнить итоговые строки.

        Args:
            section: Раздел сметы для расчёта
        """
        # Рассчитываем каждую позицию в разделе
        for item in section.items:
            self._calculate_work_item(item)

        # Собираем итоги по всему разделу
        direct_total = sum(item.direct_costs for item in section.items)
        fot_total = sum(item.fot for item in section.items)

        # Заполняем итоговые строки раздела
        for summary in section.summaries:
            if summary.row_type == RowType.SUBTOTAL_DIRECT:
                summary.total_cost = direct_total
            elif summary.row_type == RowType.SUBTOTAL_FOT:
                summary.total_cost = fot_total
            elif summary.row_type == RowType.OVERHEAD and summary.percentage:
                summary.total_cost = fot_total * (summary.percentage / Decimal("100"))
            elif summary.row_type == RowType.PROFIT and summary.percentage:
                summary.total_cost = fot_total * (summary.percentage / Decimal("100"))
            elif summary.row_type == RowType.POSITION_TOTAL:
                overhead = next((s.total_cost for s in section.summaries if s.row_type == RowType.OVERHEAD), Decimal("0"))
                profit = next((s.total_cost for s in section.summaries if s.row_type == RowType.PROFIT), Decimal("0"))
                summary.total_cost = direct_total + overhead + profit

    def _calculate_work_item(self, item: WorkItem) -> None:
        """
        Рассчитать одну позицию сметы (WorkItem).

        Алгоритм:
        1. Применить коэффициенты к каждому ресурсу
        2. Рассчитать quantity_total = норма × объём × коэффициент
        3. Рассчитать current_price = base_price × index (если не задана)
        4. Рассчитать total_cost = quantity_total × current_price
        5. Суммировать прямые затраты и ФОТ

        Args:
            item: Позиция сметы для расчёта
        """
        fot_total = Decimal("0")
        direct_total = Decimal("0")

        for res in item.resources:
            # ШАГ 1: Применяем коэффициенты к ресурсу
            self._apply_coefficients_to_resource(res, item)

            # ШАГ 2: Рассчитываем количество с учётом коэффициента
            res.quantity_total = res.quantity_per_unit * item.quantity * res.coefficient

            # ШАГ 3: Рассчитываем текущую цену (если ещё не задана)
            self._calculate_current_price(res)

            # ШАГ 4: Рассчитываем общую стоимость
            if res.current_price != Decimal("0"):
                res.total_cost = res.quantity_total * res.current_price

            # ШАГ 5: Собираем итоги по позиции
            direct_total += res.total_cost
            if res.resource_type in (ResourceType.LABOR, ResourceType.LABOR_OPERATOR):
                fot_total += res.total_cost

        item.direct_costs = direct_total
        item.fot = fot_total
        item.position_total = direct_total

    def _apply_coefficients_to_resource(self, resource: Resource, item: WorkItem) -> None:
        """
        Применить коэффициенты из WorkItem к ресурсу.
        Коэффициенты перемножаются: 1,15 × 1,15 = 1,3225.

        Args:
            resource: Ресурс для применения коэффициентов
            item: Позиция сметы, содержащая список коэффициентов
        """
        total_coef = Decimal("1.0")

        for coef in item.coefficients:
            # Определяем, какой ключ использовать в зависимости от типа ресурса
            coef_key = {
                ResourceType.LABOR: "ТЗ",           # Для рабочих - трудозатраты
                ResourceType.EQUIPMENT: "ЭМ",       # Для машин - эксплуатация
                ResourceType.LABOR_OPERATOR: "ТЗМ", # Для машинистов - трудозатраты машинистов
            }.get(resource.resource_type)

            if coef_key:
                total_coef *= coef.values.get(coef_key, Decimal("1.0"))

        resource.coefficient = total_coef

    def _calculate_current_price(self, resource: Resource) -> None:
        """
        Рассчитать текущую цену ресурса, если она ещё не задана.

        Текущая цена = базовая цена × индекс.
        Если индекс не задан, используется индекс из общих настроек по типу ресурса.

        Args:
            resource: Ресурс для расчёта цены
        """
        if resource.current_price == Decimal("0") and resource.base_price:
            # Если индекс не задан, берём из общих индексов
            if resource.index == Decimal("1.0"):
                index_key = {
                    ResourceType.LABOR: "OT",
                    ResourceType.EQUIPMENT: "EM",
                    ResourceType.LABOR_OPERATOR: "OTm",
                    ResourceType.MATERIAL: "M",
                }.get(resource.resource_type)

                if index_key:
                    resource.index = self.indices.get(index_key, Decimal("1.0"))

            resource.current_price = resource.base_price * resource.index
