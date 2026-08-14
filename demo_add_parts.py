"""
Demo: simula duas pecas "extraidas do CAD" sendo adicionadas ao Cost Report
real da equipe. Os dados abaixo sao ficticios (representam o formato que a
extracao do SolidWorks vai entregar futuramente) -- servem so para validar
se a aba gerada e o registro na BOM ficam no formato correto.

Como rodar:
    python demo_add_parts.py --template caminho\\para\\template.xlsx --out caminho\\para\\saida.xlsx
"""

import argparse

import openpyxl
from bom_builder import create_part_sheet, add_part_to_bom


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--template", required=True, help="Cost Report .xlsx base")
    parser.add_argument("--out", required=True, help="Arquivo de saida .xlsx")
    args = parser.parse_args()

    wb = openpyxl.load_workbook(args.template, data_only=False)

    # -----------------------------------------------------------------
    # Peca ficticia 1: um suporte simples de freio (Brake System)
    # -----------------------------------------------------------------
    part_1 = {
        "university": "SENAI CIMATEC",
        "system": "Brake System",
        "assembly": "Brake Line Assembly",
        "part_name": "Rear Brake Line Bracket",
        "pn_base": "FSAEB-26-24-BR99",
        "suffix": "AA",
        "details": "Suporte de fixacao da linha de freio traseira, chapa de aco.",
        "qty": 2,
        "materials": [
            {
                # Extraido do SolidWorks: massa da peca = 0.180 kg
                "material": "Steel, Mild (per kg)",
                "use": "Corpo do suporte",
                "size1": 0.180, "unit1": "kg",
                "size2": None, "unit2": "",
                "area_name": "", "area": None,
                "length": None, "density": None,
                "quantity": 1,
            },
        ],
        "processes": [
            {"process": "Laser Cut", "use": "Corte da chapa", "quantity": 40},
            {"process": "Sheet metal bends", "use": "Dobra do suporte", "quantity": 2},
            {"process": "Drilled holes < 25.4 mm dia.", "use": "Furos de fixacao", "quantity": 2},
        ],
    }

    subtotals_1 = create_part_sheet(wb, template_sheet_name="BR01",
                                     new_sheet_name="BR99", part=part_1)
    add_part_to_bom(wb, part_1, subtotals_1)
    print("Peca 1 (BR99) adicionada, subtotals:", subtotals_1)

    # -----------------------------------------------------------------
    # Peca ficticia 2: um espacador de suspensao (Suspension)
    # -----------------------------------------------------------------
    part_2 = {
        "university": "SENAI CIMATEC",
        "system": "Suspension",
        "assembly": "Front Suspension Assembly",
        "part_name": "A-Arm Spacer 3",
        "pn_base": "FSAEB-26-24-SU99",
        "suffix": "AA",
        "details": "Espacador adicional entre A-arm e upright, aluminio.",
        "qty": 4,
        "materials": [
            {
                # Extraido do SolidWorks: massa da peca = 0.032 kg
                "material": "Aluminum, Normal (per kg)",
                "use": "Espacador",
                "size1": 0.032, "unit1": "kg",
                "size2": None, "unit2": "",
                "area_name": "", "area": None,
                "length": None, "density": None,
                "quantity": 1,
            },
        ],
        "processes": [
            {"process": "Machining", "use": "Usinagem do espacador", "quantity": 3},
        ],
    }

    subtotals_2 = create_part_sheet(wb, template_sheet_name="BR01",
                                     new_sheet_name="SU99", part=part_2)
    add_part_to_bom(wb, part_2, subtotals_2)
    print("Peca 2 (SU99) adicionada, subtotals:", subtotals_2)

    wb.save(args.out)
    print("Salvo em:", args.out)


if __name__ == "__main__":
    main()
