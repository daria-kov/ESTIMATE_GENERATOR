from decimal import Decimal
from pathlib import Path
import json

from classes import Estimate, EstimateMetadata, Section, Coefficient, WorkItem
from export import EstimateExcelExporter
from norm_parser import NormParser
from calculator import Calculator
from normative_rates_parser import NormativeRatesParser


def create_estimate_from_json(json_path: Path) -> Estimate:
    """Создать смету из JSON файла (output.json)"""

    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # Метаданные из JSON
    metadata = EstimateMetadata(
        software_name=data.get("software_name", ""),
        estimate_number=data.get("estimate_number", ""),
        estimate_title=data.get("estimate_title", ""),
        region=data.get("region", ""),
        zone=data.get("zone", ""),
        method="ресурсно-индексным",
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

    # Загружаем парсер НР/СП из Excel
    nrsp_parser = None
    if nrsp_file.exists():
        try:
            nrsp_parser = NormativeRatesParser(str(nrsp_file))
            print(f"Загружен файл с нормативами НР/СП: {nrsp_file.name}")
        except Exception as e:
            print(f"Ошибка при загрузке файла НР/СП: {e}")

    # Проход по разделам из JSON
    for section_name, work_items_data in data["smeta_level"].items():
        section = Section(id=0, name=section_name)

        for item_data in work_items_data:
            # Объём из gesn_unit.value
            quantity = Decimal(str(item_data["gesn_unit"]["value"]))

            # Парсим норму с передачей mapping для AbstractResource
            work_item = parser.parse_work_by_code(
                work_code=item_data["code"],
                quantity=quantity,
                abstract_resource_mapping=item_data.get("AbstractResource", {})
            )

            if work_item is None:
                print(f"Норма не найдена: {item_data['code']}")
                continue

            conversion = item_data.get("conversion", {})
            work_item.conversion_string = conversion.get("translation_string")

            # Применяем коэффициенты
            work_item.coefficients = []
            for coef_data in item_data.get("coefficients", []):
                work_item.coefficients.append(Coefficient(
                    code=coef_data["code"],
                    description=coef_data["description"],
                    values={k: Decimal(str(v)) for k, v in coef_data["values"].items()}
                ))

            # Понижающие коэффициенты для НР и СП
            work_item.nr_reduction = Decimal(str(item_data.get("reduction_nr", 1.0)))
            work_item.sp_reduction = Decimal(str(item_data.get("reduction_sp", 1.0)))

            # НР/СП из Excel (по nr_code из XML)
            if nrsp_parser and work_item.nr_code:
                base_nr, base_sp = nrsp_parser.get_by_code(work_item.nr_code, region_type="territory")

                work_item.base_nr_percent = base_nr
                work_item.base_sp_percent = base_sp
                work_item.nr_percent = base_nr * work_item.nr_reduction
                work_item.sp_percent = base_sp * work_item.sp_reduction

            section.items.append(work_item)

        estimate.add_section(section)

    return estimate


if __name__ == "__main__":
    input_json = Path(__file__).parent / "output.json"

    if not input_json.exists():
        print(f"Файл {input_json} не найден!")
        exit(1)

    estimate = create_estimate_from_json(input_json)

    # Выполняем расчёт для каждого раздела
    calculator = Calculator(indices=estimate.metadata.indices)
    for section in estimate.sections:
        calculator.calculate_section(section)

    output_dir = Path(__file__).parent / "output"
    output_dir.mkdir(exist_ok=True)

    filepath = output_dir / "Пример_сметы.xlsx"

    exporter = EstimateExcelExporter(estimate)
    exporter.export(str(filepath))

    print(f"\nСмета экспортирована: {filepath}")
