"""
sw_extract/translator.py

Converte os dados brutos extraidos do SolidWorks (massa, material nativo,
custom properties) no formato de dicionario `part` que o bom_builder.py
espera (ver create_part_sheet em bom_builder.py).
"""

from sw_extract.extractor import (
    get_custom_properties, parse_fsae_properties,
    get_native_material, get_mass_properties,
)
from sw_extract.material_matcher import suggest_material_matches


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

    # --- Massa (kg) -----------------------------------------------------
    mass_props = get_mass_properties(model_doc)
    mass_kg = mass_props["mass_kg"]

    # --- Material nativo + fuzzy match -----------------------------------
    material_entry = None
    if material_name_override:
        material_entry = {
            "material": material_name_override,
            "use": "Corpo da peca",
            "size1": round(mass_kg, 4),
            "unit1": "kg",
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
