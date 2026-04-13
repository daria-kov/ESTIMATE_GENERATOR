from decimal import Decimal
from pathlib import Path

from classes import Estimate, EstimateMetadata, Section, SectionSummary, RowType, Coefficient
from export import EstimateExcelExporter
from norm_parser import NormParser
from calculator import Calculator
from normative_rates_parser import NormativeRatesParser


def create_estimate_with_gesn_parser() -> Estimate:
    """Создать смету с использованием парсера ГЭСН и калькулятора"""

    metadata = EstimateMetadata(
        software_name="Пример ЛС БИМ для ФГИС ЦС",
        estimate_number="ЛС-02-01-01",
        estimate_title="Пример локального сметного расчета (сметы)_базисно-индексный метод с применением индексов по элементам прямых затрат",
        region="г. Москва",
        zone="Центральная",
    )

    estimate = Estimate(metadata=metadata)

    # Пути к файлам
    xml_folder = Path(__file__).parent
    price_file = Path(__file__).parent / "Сплит-форма Республика Татарстан на 4 квартал 2025 года.xlsx"
    fsbc_machines = Path(__file__).parent / "ФСБЦ_Маш.xml"
    fsbc_materials = Path(__file__).parent / "ФСБЦ_Мат&Оборуд.xml"
    nrsp_file = Path(__file__).parent / "ФСНБ-2022  Привязка НР (812_пр) и СП (774_пр) ФСНБ-2022 Доп.17.xlsx"

    if not xml_folder.exists():
        print(f"Папка с XML не найдена: {xml_folder}")
        return estimate

    # Создаём парсер ГЭСН
    parser = NormParser(
        str(xml_folder),
        price_db_path=str(price_file) if price_file.exists() else None,
        fsbc_mash_path=str(fsbc_machines) if fsbc_machines.exists() else None,
        fsbc_materials_path=str(fsbc_materials) if fsbc_materials.exists() else None
    )

    # Загружаем парсер НР/СП
    nrsp_parser = None
    if nrsp_file.exists():
        try:
            nrsp_parser = NormativeRatesParser(str(nrsp_file))
            print(f"Загружен файл с нормативами НР/СП: {nrsp_file.name}")
        except Exception as e:
            print(f"Ошибка при загрузке файла НР/СП: {e}")

    # Входные данные: список кодов работ и объёмов
    user_input = [
        {"code": "ГЭСН06-01-004-04", "quantity": Decimal("9")},
        {"code": "ГЭСН11-01-011-03", "quantity": Decimal("0.43")}
    ]

    # Создаём раздел
    section = Section(id=0, name="Общестроительные работы")

    # Общие коэффициенты (применяются ко всем позициям)
    common_coefficients = [
        Coefficient(
            code="421/пр_2020_прил.10_т.5_п.3_гр.3",
            description="Производство ремонтно-строительных работ... ОЗП=1,15; ЭМ=1,15; ЗПМ=1,15; ТЗ=1,15; ТЗМ=1,15",
            values={"ОЗП": Decimal("1.15"), "ЭМ": Decimal("1.15"), "ЗПМ": Decimal("1.15"), "ТЗ": Decimal("1.15"), "ТЗМ": Decimal("1.15")},
        ),
        Coefficient(
            code="421/пр_2020_п.58_пп.б",
            description="При применении сметных норм... ОЗП=1,15; ЭМ=1,25; ЗПМ=1,25; ТЗ=1,15; ТЗМ=1,25",
            values={"ОЗП": Decimal("1.15"), "ЭМ": Decimal("1.25"), "ЗПМ": Decimal("1.25"), "ТЗ": Decimal("1.15"), "ТЗМ": Decimal("1.25")},
        ),
    ]

    for item_data in user_input:
        work_item = parser.parse_work_by_code(
            work_code=item_data["code"],
            quantity=item_data["quantity"],
        )

        if work_item is None:
            print(f"Норма не найдена: {item_data['code']}")
            continue

        # Применяем коэффициенты
        work_item.coefficients = common_coefficients.copy()

        # Применяем нормативы НР/СП
        if nrsp_parser and work_item.nr_code:
            base_nr, base_sp = nrsp_parser.get_by_code(work_item.nr_code)

            reduction_nr = Decimal("0.9")
            reduction_sp = Decimal("0.85")

            work_item.base_nr_percent = base_nr
            work_item.base_sp_percent = base_sp
            work_item.nr_reduction = reduction_nr
            work_item.sp_reduction = reduction_sp
            work_item.nr_percent = base_nr * reduction_nr
            work_item.sp_percent = base_sp * reduction_sp

        section.items.append(work_item)

    # Назначаем номера позиций (начиная с 12, как в образце)
    for idx, item in enumerate(section.items):
        item.row_num = idx

    estimate.add_section(section)

    # Выполняем расчёт
    calculator = Calculator(indices=metadata.indices)
    calculator.calculate_section(section)

    return estimate


if __name__ == "__main__":
    estimate = create_estimate_with_gesn_parser()

    output_dir = Path(__file__).parent / "output"
    output_dir.mkdir(exist_ok=True)

    filepath = output_dir / "смета_гэсн_пример.xlsx"

    exporter = EstimateExcelExporter(estimate)
    exporter.export(str(filepath))

    print(f"\nСмета экспортирована: {filepath}")
