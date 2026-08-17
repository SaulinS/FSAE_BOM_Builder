"""
sw_extract/translator.py

Converte os dados brutos extraidos do SolidWorks (massa, material nativo,
custom properties) no formato de dicionario `part` que o bom_builder.py
espera (ver create_part_sheet em bom_builder.py).
"""

from decimal import Decimal

from sw_extract.extractor import (
    get_custom_properties, parse_fsae_properties,
    get_native_material, get_mass_properties, MASS_UNIT_TO_KG,
)
from sw_extract.material_matcher import suggest_material_matches


# Unidade de massa que a planilha do Cost Report espera na coluna Size1.
# CONFIRMADO (2026-08-16) contra um cost report real publicado (University
# of Delaware, FSAE Lincoln 2017, Car #67): a linha de material
# "Aluminum, Normal | $4.20 | 0.014 | kg | ... | 2712.0 kg/m3 | 1 | $0.06"
# mostra Size1 em QUILOGRAMAS e o custo unitario por kg. Se um ano o
# template mudar para gramas, basta trocar esta constante -- a formula de
# conversao escrita na planilha se ajusta sozinha.
BOM_MASS_UNIT = "kg"

# Quantos digitos significativos manter ao escrever um numero numa formula.
_SIG_DIGITS = 12


def _plain_number(value, sig_digits=_SIG_DIGITS):
    """
    Formata um numero em notacao decimal simples, NUNCA cientifica, com
    `sig_digits` algarismos significativos e sem zeros sobrando no fim.

    Existe porque os formatos padrao do Python ("%g", repr) mudam para
    notacao cientifica sozinhos em numeros pequenos -- 0.000001 vira
    "1e-06" -- e uma formula com notacao cientifica e escrita pelo openpyxl
    sem erro nenhum, mas o Excel pode recusar na hora de abrir. Como o
    openpyxl nao avalia formula, esse tipo de erro so apareceria na frente
    do juiz de custo.
    """
    dec = Decimal(f"%.{sig_digits}g" % float(value))
    text = format(dec, "f")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"


def _mass_size1(mass_props):
    """
    Monta o valor da coluna Size1 (massa) como uma FORMULA do Excel, no
    formato `=<valor_como_o_documento_exibe>*<fator>`, em vez de um numero
    ja convertido em Python.

    Por que formula e nao numero pronto: a conversao fica visivel na
    planilha. O primeiro termo e exatamente o numero que aparece no dialogo
    de Mass Properties do SolidWorks (nas unidades do documento), entao da
    para conferir a olho -- "o SolidWorks diz 14 g" contra "=14*0.001" na
    celula -- sem abrir o codigo. Um erro de unidade deixa de ser silencioso
    e passa a ser algo que qualquer pessoa revisando a planilha percebe.

    Retorna (valor_para_size1, unidade, warnings).
    """
    warnings = []
    mass_kg = mass_props["mass_kg"]
    doc_value = mass_props.get("doc_mass_value")
    doc_unit = mass_props.get("doc_mass_unit")
    doc_factor = mass_props.get("doc_unit_factor_kg")
    target_factor = MASS_UNIT_TO_KG[BOM_MASS_UNIT]

    if not mass_props.get("used_system_units", False):
        warnings.append(
            "Nao foi possivel forcar UseSystemUnits=True ao ler a massa -- "
            "o valor pode estar nas unidades do documento, nao em kg. "
            "CONFERIR a massa desta peca manualmente antes de submeter."
        )

    if doc_value is None or doc_factor is None:
        # Nao deu para descobrir a unidade do documento: cai para o valor SI
        # (que ja veio de UseSystemUnits=True) e avisa, para ninguem supor
        # que a conversao foi verificada.
        warnings.append(
            f"Nao foi possivel determinar a unidade de massa do documento "
            f"(unidade detectada: {doc_unit!r}). Size1 foi preenchido com o "
            f"valor em kg lido no modo de unidades de sistema, sem formula "
            f"de conversao -- conferir contra a massa real da peca."
        )
        return round(mass_kg / target_factor, 6), BOM_MASS_UNIT, warnings

    multiplier = doc_factor / target_factor
    # Nao arredondamos cedo -- quem arredonda e o Excel. Mas os dois termos
    # PRECISAM sair em notacao decimal simples (ver _plain_number): uma peca
    # leve o bastante, ou o fator de miligramas, sairiam como "1e-06" no
    # formato %g padrao, e formula em notacao cientifica e justamente o tipo
    # de coisa que o openpyxl escreve sem reclamar e o Excel recusa depois.
    return (f"={_plain_number(doc_value)}*{_plain_number(multiplier)}",
            BOM_MASS_UNIT, warnings)


def build_part_dict(component_entry, catalog_material_names, university="",
                     material_name_override=None):
    """
    Recebe uma entrada de traverse_assembly() (com .component e .quantity)
    e devolve:
        (part_dict, warnings)

    `part_dict` esta pronto para ser passado a bom_builder.create_part_sheet,
    exceto pelos campos system/assembly/pn_base/suffix/details/processes,
    que dependem das custom properties FSAE_* -- se ainda nao tiverem sido
    preenchidas na peca (via a GUI de padronizacao), esses campos voltam
    vazios e isso e reportado como warning, NAO como erro (a peca ainda e
    gerada, so precisa de revisao manual antes de submeter).

    `material_name_override`: se fornecido, pula o fuzzy match e usa esse
    nome de material diretamente (usado quando uma pessoa ja confirmou a
    sugestao na tela do corretor).
    """
    warnings = []
    component = component_entry["component"]
    model_doc = component.GetModelDoc2()

    if model_doc is None:
        warnings.append(
            f"Componente '{component_entry['path']}' nao pode ser aberto "
            f"(arquivo ausente ou nao carregado) -- pulado."
        )
        return None, warnings

    # --- Propriedades FSAE_* (system/assembly/pn_base/suffix/details/processos)
    raw_props = get_custom_properties(model_doc, component_entry.get("config", ""))
    fsae = parse_fsae_properties(raw_props)

    if not fsae["system"]:
        warnings.append(
            f"Peca '{model_doc.GetTitle()}' sem FSAE_System preenchido -- "
            f"revisar antes de gerar a BOM final."
        )
    if not fsae["processes"]:
        warnings.append(
            f"Peca '{model_doc.GetTitle()}' sem nenhum FSAE_Process_N "
            f"preenchido -- bloco de Processos sera gerado vazio."
        )

    # --- Massa ------------------------------------------------------------
    # A leitura ja vem forcada para unidades de sistema (kg) pelo extractor,
    # e acompanhada da unidade do proprio documento -- ver _mass_size1, que
    # transforma isso na formula de conversao escrita na planilha.
    mass_props = get_mass_properties(model_doc)
    size1_value, size1_unit, mass_warnings = _mass_size1(mass_props)
    for w in mass_warnings:
        warnings.append(f"Peca '{model_doc.GetTitle()}': {w}")

    # --- Material nativo + fuzzy match -----------------------------------
    material_entry = None
    if material_name_override:
        material_entry = {
            "material": material_name_override,
            "use": "Corpo da peca",
            "size1": size1_value,
            "unit1": size1_unit,
            "size2": None, "unit2": "",
            "area_name": "", "area": None,
            "length": None, "density": None,
            "quantity": 1,
        }
    else:
        native_material, _db = get_native_material(model_doc, component_entry.get("config", ""))
        if native_material:
            # Por design (ver material_matcher.py), o corretor fuzzy so
            # SUGERE candidatos -- nunca aplica um automaticamente aqui.
            # A peca e gerada com o bloco de Material vazio; uma pessoa
            # revisa as sugestoes e reroda com material_name_override.
            suggestions = suggest_material_matches(native_material, catalog_material_names)
            if suggestions:
                warnings.append(
                    f"Peca '{model_doc.GetTitle()}': material nativo "
                    f"'{native_material}' sem correspondencia confirmada -- "
                    f"bloco de Material sera gerado vazio. Sugestoes do "
                    f"corretor (revisar e confirmar com "
                    f"material_name_override): {suggestions}."
                )
            else:
                warnings.append(
                    f"Peca '{model_doc.GetTitle()}': material nativo "
                    f"'{native_material}' sem correspondencia no catalogo -- "
                    f"bloco de Material sera gerado vazio. Revisar "
                    f"manualmente e informar material_name_override."
                )
        else:
            warnings.append(
                f"Peca '{model_doc.GetTitle()}' sem material atribuido no "
                f"SolidWorks -- bloco de Material sera gerado vazio."
            )

    part = {
        "university": university,
        "system": fsae["system"],
        "assembly": fsae["assembly"],
        "part_name": model_doc.GetTitle(),
        "pn_base": fsae["pn_base"],
        "suffix": fsae["suffix"] or "AA",
        "details": fsae["details"],
        "qty": component_entry["quantity"],
        "materials": [material_entry] if material_entry else [],
        "processes": fsae["processes"],
    }
    return part, warnings
