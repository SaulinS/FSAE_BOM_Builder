"""
sw_extract/extractor.py

Funcoes de extracao de dados de uma peca/montagem aberta no SolidWorks:
- propriedades customizadas (FSAE_System, FSAE_Process_1, etc.)
- material nativo aplicado a peca (para o corretor fuzzy casar depois
  com o catalogo Materials)
- propriedades de massa (massa em kg -- **VERIFICAR PRIMEIRO**: ver aviso
  detalhado em get_mass_properties, o valor pode nao vir em kg dependendo
  do sistema de unidades do documento)
- travessia da arvore de montagem, agrupando componentes repetidos e
  contando quantidade, ignorando suprimidos e itens marcados como
  "excluir da BOM"

TESTADO PARCIALMENTE ao vivo em 2026-08-12 (SOLIDWORKS 2025, peca simples,
sem montagem aberta) -- ver aviso detalhado em connector.py sobre metodos
sem parametros precisarem ser acessados SEM parenteses nesta instalacao
(GetType, GetTitle, GetPathName, GetNames, CreateMassProperty -- todos
confirmados). Parametros de SAIDA (ByRef) como o 2o parametro de
GetMaterialPropertyName2 tambem precisam ser um
win32com.client.VARIANT(pythoncom.VT_BYREF | ..., valor_inicial) em vez de
uma string/bool literal -- confirmado em get_native_material. O mesmo
padrao foi aplicado por analogia (ainda NAO confirmado ao vivo) em Get5
(get_custom_properties, precisa de uma peca com FSAE_* de verdade pra
testar) e em GetChildren/IsSuppressed/GetModelDoc2 (traverse_assembly,
precisa de uma montagem aberta pra testar).
"""

import pythoncom
import win32com.client

from sw_extract.connector import SW_DOC_PART, SW_DOC_ASSEMBLY

# Quantos slots de "Processo" o schema de custom properties suporta.
# Aumentar aqui se alguma peca precisar de mais de 8 processos.
MAX_PROCESS_SLOTS = 8


# ---------------------------------------------------------------------------
# Propriedades customizadas
# ---------------------------------------------------------------------------

def get_custom_properties(model_doc, config_name=""):
    """
    Le todas as custom properties do documento (na configuracao indicada,
    ou na configuracao ativa se config_name="").

    OBS: a assinatura exata de CustomPropertyManager.Get5 pode variar por
    versao do SW. Se der erro de assinatura, tentem Get4 ou Get3 -- a API
    do SolidWorks manteve varias versoes desse metodo por compatibilidade.
    O CustomPropertyManager(config_name) em si (metodo com parametro
    obrigatorio) e o GetNames sem parenteses ja foram confirmados ao vivo;
    o Get5 abaixo ainda nao (a peca de teste usada nao tinha nenhuma custom
    property) -- testem contra uma peca com FSAE_* preenchido (Passo 2 do
    README) antes de confiar nos valores lidos.
    """
    cust_prop_mgr = model_doc.Extension.CustomPropertyManager(config_name)
    names = cust_prop_mgr.GetNames  # sem parenteses -- ver aviso no topo do arquivo
    props = {}
    if not names:
        return props
    for name in names:
        try:
            # Get5(FieldName, UseCached, ValOut, ResolvedValOut, WasResolved)
            # -- os 3 ultimos sao parametros de SAIDA (ByRef) e precisam ser
            # VARIANTs de verdade, nao string/bool literais (o win32com nao
            # escreve de volta numa string Python comum).
            val_out = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BSTR, "")
            resolved_out = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BSTR, "")
            was_resolved_out = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BOOL, False)
            cust_prop_mgr.Get5(name, False, val_out, resolved_out, was_resolved_out)
            props[name] = resolved_out.value
        except Exception as exc:
            props[name] = None
            print(f"[aviso] Falha ao ler custom property '{name}': {exc}")
    return props


def parse_fsae_properties(raw_props):
    """
    Converte o dicionario bruto de custom properties (nomes FSAE_*) no
    formato usado pelo bom_builder: system, assembly, pn_base, suffix,
    details, processes (lista de dicts).
    """
    parsed = {
        "system": raw_props.get("FSAE_System", ""),
        "assembly": raw_props.get("FSAE_Assembly", ""),
        "pn_base": raw_props.get("FSAE_PN_Base", ""),
        "suffix": raw_props.get("FSAE_Suffix", ""),
        "details": raw_props.get("FSAE_Details", ""),
        "processes": [],
    }
    for i in range(1, MAX_PROCESS_SLOTS + 1):
        proc_name = raw_props.get(f"FSAE_Process_{i}")
        if not proc_name:
            continue
        parsed["processes"].append({
            "process": proc_name,
            "use": raw_props.get(f"FSAE_Process_{i}_Use", ""),
            "quantity": _to_number(raw_props.get(f"FSAE_Process_{i}_Qty", 1)),
        })
    return parsed


def _to_number(val, default=1):
    try:
        return float(val)
    except (TypeError, ValueError):
        return default


# ---------------------------------------------------------------------------
# Material nativo do SolidWorks
# ---------------------------------------------------------------------------

def get_native_material(model_doc, config_name=""):
    """
    Le o material atribuido nativamente na peca (Material Editor do SW),
    nao uma custom property. Retorna (nome_material, nome_biblioteca) ou
    (None, None) se nenhum material foi atribuido.

    So se aplica a documentos do tipo Part (SW_DOC_PART); montagens nao
    tem material proprio.

    CONFIRMADO ao vivo (2026-08-12): GetMaterialPropertyName2 retorna o
    nome do material diretamente (uma string, nao uma lista/tupla como a
    primeira versao deste codigo supunha); o 2o parametro (nome do banco
    de dados) e um parametro de SAIDA (ByRef) que precisa ser um
    win32com.client.VARIANT, senao a chamada da erro de "tipo nao
    correspondente".
    """
    if model_doc.GetType != SW_DOC_PART:  # sem parenteses -- ver topo do arquivo
        return None, None
    try:
        db_out = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BSTR, "")
        material_name = model_doc.GetMaterialPropertyName2(config_name, db_out)
        database_name = db_out.value
        if not material_name:
            return None, None
        return material_name, database_name
    except Exception as exc:
        print(f"[aviso] Falha ao ler material nativo: {exc}")
        return None, None


# ---------------------------------------------------------------------------
# Propriedades de massa
# ---------------------------------------------------------------------------

# Fatores de conversao para kg, por unidade de massa. Usados tanto para
# classificar a unidade do documento (ver _classify_mass_unit) quanto para
# montar a formula de conversao na planilha (ver translator.py).
MASS_UNIT_TO_KG = {
    "kg": 1.0,
    "g": 0.001,
    "mg": 0.000001,
    "lb": 0.45359237,
    "oz": 0.028349523125,
}


def _classify_mass_unit(ratio, tolerance=0.01):
    """
    Recebe a razao (massa em unidades do documento / massa em unidades de
    sistema) e devolve o nome da unidade do documento, ou None se a razao
    nao corresponder a nenhuma unidade conhecida.

    Ex.: um documento MMGS devolve gramas, entao a razao e ~1000 e a
    unidade e "g" (1 kg = 1000 g => fator para kg = 0.001).
    """
    if not ratio or ratio <= 0:
        return None
    for unit, factor_to_kg in MASS_UNIT_TO_KG.items():
        # massa_doc = massa_kg / factor_to_kg  =>  ratio esperado = 1/factor
        expected = 1.0 / factor_to_kg
        if abs(ratio - expected) <= tolerance * expected:
            return unit
    return None


def get_mass_properties(model_doc):
    """
    Retorna massa, volume e area de superficie em unidades SI (kg, m^3,
    m^2), FORCANDO o modo de unidades de sistema em vez de supor que a API
    ja devolve SI.

    Como a ambiguidade de unidade foi resolvida (documentacao oficial da
    API, confirmada em 2026-08-16): IMassProperty.UseSystemUnits aceita
    True para unidades de SISTEMA (metros, radianos, quilogramas) e False
    para as unidades do DOCUMENTO; o default e True. Como o default nao e
    garantido em toda versao/instalacao, este codigo seta a propriedade
    explicitamente em vez de confiar nele -- e dai vem a garantia de kg.

    Alem de forcar SI, a funcao SONDA a unidade do proprio documento: le a
    massa nos dois modos e compara. A razao entre as duas leituras revela
    empiricamente em que unidade o documento esta (razao ~1000 => gramas,
    ~1 => kg, ~2.2046 => libras). Isso e feito por medicao, sem depender de
    constantes de enum da API (swUnitsMassPropMass), cujos valores
    numericos variam entre versoes e seriam mais um ponto de erro
    silencioso. O resultado alimenta o aviso do translator e a formula de
    conversao escrita na planilha.

    Chaves retornadas:
        mass_kg              massa em kg (o valor a usar na BOM)
        volume_m3            volume em m^3
        surface_area_m2      area de superficie em m^2
        doc_mass_unit        unidade do documento ("kg", "g", ...) ou None
        doc_mass_value       massa como o documento a exibe
        doc_unit_factor_kg   fator que converte doc_mass_value -> kg
        used_system_units    se conseguimos forcar o modo SI
    """
    mass_prop = model_doc.Extension.CreateMassProperty  # sem parenteses

    used_system_units = True
    try:
        mass_prop.UseSystemUnits = True
    except Exception as exc:
        used_system_units = False
        print(f"[aviso] Nao foi possivel forcar UseSystemUnits=True: {exc}")

    mass_si = mass_prop.Mass
    volume_si = mass_prop.Volume
    area_si = mass_prop.SurfaceArea

    # --- Sonda a unidade do documento ------------------------------------
    doc_mass_value = None
    doc_mass_unit = None
    try:
        mass_prop.UseSystemUnits = False
        doc_mass_value = mass_prop.Mass
        if mass_si:
            doc_mass_unit = _classify_mass_unit(doc_mass_value / mass_si)
    except Exception as exc:
        print(f"[aviso] Nao foi possivel sondar a unidade do documento: {exc}")
    finally:
        # Restaura o modo SI -- o objeto pode ser reutilizado pelo chamador.
        try:
            mass_prop.UseSystemUnits = True
        except Exception:
            pass

    return {
        "mass_kg": mass_si,
        "volume_m3": volume_si,
        "surface_area_m2": area_si,
        "doc_mass_unit": doc_mass_unit,
        "doc_mass_value": doc_mass_value,
        "doc_unit_factor_kg": MASS_UNIT_TO_KG.get(doc_mass_unit),
        "used_system_units": used_system_units,
    }


# ---------------------------------------------------------------------------
# Travessia da arvore de montagem
# ---------------------------------------------------------------------------

def traverse_assembly(assembly_doc):
    """
    Percorre a arvore de componentes de uma montagem (recursivamente,
    incluindo sub-montagens) e retorna uma lista de entradas, uma por
    componente UNICO (agrupando instancias repetidas por quantidade),
    ignorando componentes suprimidos ou marcados como "excluir da BOM"
    -- exatamente como a BOM nativa do SolidWorks faz.

    Cada entrada da lista:
        {
            "path": caminho do arquivo do componente,
            "config": configuracao referenciada,
            "quantity": quantidade de instancias na montagem,
            "component": referencia ao objeto Component2 (para extrair
                         propriedades/massa depois),
            "is_assembly": True/False,
        }
    """
    root = assembly_doc.ConfigurationManager.ActiveConfiguration.GetRootComponent3(True)
    grouped = {}
    _walk_component_tree(root, grouped, is_root=True)
    return list(grouped.values())


def _walk_component_tree(component, grouped, is_root=False):
    # GetChildren/IsSuppressed/GetModelDoc2 acessados sem parenteses por
    # analogia com GetType/GetTitle/GetPathName/GetNames/CreateMassProperty
    # (todos confirmados ao vivo -- ver aviso no topo do arquivo), mas isso
    # especificamente ainda NAO foi testado contra uma montagem de verdade
    # (a sessao de teste so tinha uma peca aberta). Confirmem no Passo 3 do
    # README antes de confiar na travessia numa montagem grande.
    children = component.GetChildren
    if not children:
        return
    for child in children:
        if child is None:
            continue
        if child.IsSuppressed:
            continue
        try:
            # getattr(..., default) so protege contra AttributeError, mas o
            # dispatch tardio do win32com pode levantar outro tipo de erro
            # (pywintypes.com_error) pra uma propriedade COM desconhecida --
            # entao capturamos Exception explicitamente, sem deixar a
            # travessia inteira cair por causa de uma unica propriedade.
            exclude_from_bom = bool(child.ExcludeFromBOM)
        except Exception:
            exclude_from_bom = False
        if exclude_from_bom:
            continue

        key = (child.GetPathName, child.ReferencedConfiguration)
        if key not in grouped:
            child_doc = child.GetModelDoc2
            grouped[key] = {
                "path": key[0],
                "config": key[1],
                "quantity": 0,
                "component": child,
                "is_assembly": (child_doc.GetType == SW_DOC_ASSEMBLY) if child_doc else False,
            }
        grouped[key]["quantity"] += 1

        # Recursao: entra em sub-montagens tambem (a BOM real do time trata
        # cada nivel -- montagem E pecas filhas -- como linhas separadas,
        # entao continuamos descendo independentemente do tipo).
        _walk_component_tree(child, grouped, is_root=False)
