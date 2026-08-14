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

    print("\nSe os valores acima baterem com o que voces esperam da peca "
          "aberta, a conexao esta funcionando. Proximo passo: testar "
          "traverse_assembly() numa montagem pequena.")


if __name__ == "__main__":
    main()
