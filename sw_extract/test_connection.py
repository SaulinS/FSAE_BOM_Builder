"""
sw_extract/test_connection.py

RODEM ISSO PRIMEIRO, com uma peca simples aberta no SolidWorks (nao uma
montagem inteira ainda). O objetivo e so confirmar que a conexao COM
funciona e que os metodos da API tem a assinatura esperada nessa versao
do SolidWorks instalada na maquina de voces.

O MAIS IMPORTANTE a conferir nesta etapa: peguem uma peca de massa
CONHECIDA (ex.: ja pesada ou calculada a mao) e comparem com o valor de
"mass_kg" impresso abaixo. get_mass_properties() assume que o SolidWorks
sempre devolve a massa em kg, mas isso pode nao ser verdade dependendo do
sistema de unidades do template da peca (MMGS entrega gramas, por
exemplo) -- ver aviso detalhado em sw_extract/extractor.py. Se esse valor
nao bater, TODO custo de material auto-extraido depois fica errado, entao
nao sigam pra proxima etapa sem confirmar isso.

Como rodar (a partir da pasta QUE CONTEM sw_extract\\, nao de dentro dela --
e um import relativo, so funciona rodado como modulo):
    pip install pywin32
    python -m sw_extract.test_connection
"""

from sw_extract.connector import connect_to_solidworks, get_active_document, describe_document
from sw_extract.extractor import get_custom_properties, get_native_material, get_mass_properties


def main():
    print("Conectando ao SolidWorks...")
    sw_app = connect_to_solidworks()
    print("Conectado. Versao:", sw_app.RevisionNumber)

    model = get_active_document(sw_app)
    info = describe_document(model)
    print("\nDocumento ativo:")
    for k, v in info.items():
        print(f"  {k}: {v}")

    print("\n--- Custom Properties ---")
    props = get_custom_properties(model)
    if not props:
        print("  (nenhuma custom property encontrada)")
    for k, v in props.items():
        print(f"  {k} = {v!r}")

    print("\n--- Material nativo ---")
    material, db = get_native_material(model)
    print(f"  Material: {material!r}  (biblioteca: {db!r})")

    print("\n--- Propriedades de massa ---")
    mass = get_mass_properties(model)
    for k, v in mass.items():
        print(f"  {k}: {v}")

    print("\n--- Diagnostico de unidade de massa ---")
    if not mass["used_system_units"]:
        print("  !! NAO foi possivel forcar unidades de sistema. O valor de "
              "mass_kg acima pode NAO estar em kg -- confiram manualmente "
              "antes de usar qualquer custo extraido.")
    elif mass["doc_mass_unit"] is None:
        print("  !! Unidade do documento nao identificada. mass_kg veio do "
              "modo de unidades de sistema (deveria ser kg), mas a sonda "
              "nao conseguiu confirmar comparando os dois modos.")
        print("     Comparem mass_kg com a massa real da peca antes de seguir.")
    else:
        print(f"  Documento exibe a massa em: {mass['doc_mass_unit']} "
              f"({mass['doc_mass_value']})")
        print(f"  Convertido para kg: {mass['mass_kg']}")
        print(f"  Na planilha, Size1 sai como: "
              f"={mass['doc_mass_value']:.10g}*{mass['doc_unit_factor_kg']:.10g}")
        if mass["doc_mass_unit"] != "kg":
            print(f"  (o template desta peca NAO esta em kg -- a conversao "
                  f"acima e justamente o que evita o erro de "
                  f"{1 / mass['doc_unit_factor_kg']:.0f}x)")

    print("\n>> CONFERIR AGORA: o valor de 'doc_mass_value' acima tem que "
          "bater com o que o SolidWorks mostra em Mass Properties para esta "
          "peca, e 'mass_kg' tem que bater com a massa real em kg. Se as "
          "duas coisas baterem, a extracao de massa esta validada.")

    print("\nSe os valores acima baterem com o que voces esperam da peca "
          "aberta, a conexao esta funcionando. Proximo passo: testar "
          "traverse_assembly() numa montagem pequena.")


if __name__ == "__main__":
    main()
