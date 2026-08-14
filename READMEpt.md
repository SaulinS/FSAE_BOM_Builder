# FSAE BOM Builder

🇧🇷 Português | [🇺🇸 English](README.md)

Gerador automatizado de BOM (lista de materiais) para um carro FSAE (Formula
SAE), que extrai dados das peças diretamente do **SolidWorks** via sua API
COM — custom properties, material nativo e massa — e monta uma planilha de
BOM padronizada.

Criado para eliminar a transcrição manual e sujeita a erro da BOM a partir
do CAD: em vez de copiar nome, material e massa das peças à mão para uma
planilha, o pipeline lê a árvore de montagem do SolidWorks diretamente e
gera uma saída estruturada e padronizada.

## ⚠️ Status: não testado numa máquina real com SolidWorks

Este código foi escrito seguindo a API oficial do SolidWorks, mas **não foi
testado em uma máquina real** (o ambiente onde foi escrito é Linux, sem
SolidWorks). Esperem precisar de ajustes ao rodar pela primeira vez —
principalmente em nomes/assinaturas de método, que podem variar entre
versões do SolidWorks (veja as notas de troubleshooting abaixo).

## O que ainda depende de trabalho manual (por design)

- Preencher `FSAE_System`, `FSAE_Process_N`, etc. em cada peça no
  SolidWorks (uma GUI de padronização para facilitar isso está planejada,
  mas ainda não foi construída).
- Confirmar as sugestões do corretor fuzzy de material — o script nunca
  aplica uma correspondência sozinho sem sinalizar como sugestão para
  revisão.
- Fixadores e ferramental continuam sendo preenchidos manualmente na aba
  gerada.

## Stack técnica

| Camada | Tecnologia |
|---|---|
| Integração CAD | API COM do SolidWorks via `pywin32` |
| Processamento de dados | Python |
| Saída | Planilha de BOM em Excel (`.xlsx`) |
| Ambiente | Windows (SolidWorks necessário), Python 3 |

## Estrutura do projeto

```
files/
├── sw_extract/
│   ├── connector.py          # conexão COM com o SolidWorks
│   ├── extractor.py          # leitura de custom properties, material nativo, massa, e travessia da árvore de montagem
│   ├── material_matcher.py   # corretor fuzzy (material do SW -> nome exato do catálogo)
│   ├── translator.py         # converte dados extraídos no formato que o bom_builder.py espera
│   └── test_connection.py    # script de teste de conexão isolado (rodar primeiro)
├── bom_builder.py            # monta a planilha de BOM padronizada
├── demo_add_parts.py         # exemplo/demo de adição de peças a uma BOM
└── extract_and_build.py      # pipeline completo (rodar por último)
```

## Como rodar

**Pré-requisitos:** Windows com SolidWorks instalado.

```bash
pip install pywin32
```

### Ordem de teste recomendada (não pulem etapas)

**1. Teste de conexão numa peça simples**

Abram uma peça de massa **conhecida** (já pesada ou calculada à mão — não
uma montagem ainda) no SolidWorks. Depois, a partir desta pasta (`files\`,
a pasta que *contém* `sw_extract\` — não entrem nela com `cd`, é um import
relativo, só funciona rodado como módulo a partir de fora):

```bash
python -m sw_extract.test_connection
```

Isso deve imprimir: título do documento, custom properties, material
nativo, massa. Se der erro aqui, é mais fácil debugar numa peça simples do
que numa montagem inteira — resolvam isso antes de ir para o próximo passo.

**O mais importante de conferir nesta etapa:** o valor de `mass_kg`
impresso precisa bater com a massa real da peça. `get_mass_properties()`
assume que o SolidWorks sempre devolve a massa em kg, mas isso depende do
sistema de unidades do template usado (um template MMGS, por exemplo,
devolveria gramas) — isso ainda não foi confirmado numa máquina real. Se
esse valor vier errado e ninguém notar aqui, todo custo de material
auto-extraído dali em diante fica silenciosamente errado (potencialmente
1000x, se for kg vs. g).

Erros esperáveis nessa etapa e o que costumam significar:
- `AttributeError` em `Get5` → a versão do SW usa uma assinatura diferente
  do método. Tentem `Get4` ou `Get3` (ver comentário em `extractor.py`).
- Material nativo vem `None` → a peça provavelmente não tem material
  atribuído ainda no SolidWorks (Material Editor).
- `GetActiveObject` falha → o SolidWorks não está aberto, ou está aberto
  mas sem nenhum documento carregado.

**2. Teste com propriedades `FSAE_*` preenchidas manualmente**

Adicionem manualmente (via SolidWorks: File > Properties > Custom) as
propriedades `FSAE_System`, `FSAE_PN_Base`, `FSAE_Suffix`, `FSAE_Details`,
`FSAE_Process_1`, `FSAE_Process_1_Use`, `FSAE_Process_1_Qty` numa peça de
teste, e rodem `python -m sw_extract.test_connection` de novo (a partir de
`files\`) — confirmem que os valores aparecem certinho na saída.

**3. Teste de travessia numa montagem pequena**

Peguem uma sub-montagem pequena e conhecida (ex: só o sistema de freio) e
testem `traverse_assembly()` isoladamente antes de ir para o carro
inteiro — é mais fácil conferir se a contagem de quantidade e a lista de
peças batem com o que vocês esperam manualmente.

**4. Pipeline completo**

Só depois dos passos acima funcionarem, rodem:

```bash
python extract_and_build.py --template caminho\para\template.xlsx --out caminho\para\saida.xlsx
```

## Autor

Construído como parte do ferramental de uma equipe FSAE, por um estudante
de Engenharia da Computação.
