# Extracao CAD -> BOM (SolidWorks)

## Antes de rodar qualquer coisa

Este codigo foi escrito seguindo a API oficial do SolidWorks, mas **nao foi
testado em uma maquina real** (o ambiente onde foi escrito e Linux, sem
SolidWorks). Esperem precisar de ajustes ao rodar pela primeira vez --
principalmente em nomes/assinaturas de metodo que podem variar entre
versoes do SolidWorks.

## Setup

Na maquina Windows com SolidWorks instalado:

```
pip install pywin32
```

## Ordem de teste recomendada (NAO pulem etapas)

### 1. Teste de conexao numa peca simples

Abram uma peca de massa CONHECIDA (ja pesada ou calculada a mao -- nao uma
montagem ainda) no SolidWorks. Depois, a partir desta pasta (`files\`, a
pasta QUE CONTEM `sw_extract\` -- nao entrem nela com `cd`, e um import
relativo, so funciona rodado como modulo a partir de fora):

```
pip install pywin32
python -m sw_extract.test_connection
```

Isso deve imprimir: titulo do documento, custom properties, material
nativo, massa. Se der erro aqui, e mais facil de debugar numa peca simples
do que numa montagem inteira -- resolvam isso antes de ir pro proximo
passo.

**O mais importante de conferir nesta etapa:** o valor de `mass_kg`
impresso precisa bater com a massa real da peca. `get_mass_properties()`
assume que o SolidWorks sempre devolve a massa em kg, mas isso depende do
sistema de unidades do template usado (um template MMGS, por exemplo,
devolveria gramas) -- isso NAO foi confirmado numa maquina real ainda. Se
esse valor vier errado e ninguem notar aqui, todo custo de material
auto-extraido dai em diante fica silenciosamente errado (potencialmente
1000x, se for kg vs. g).

**Erros esperaveis nessa etapa e o que costumam significar:**
- `AttributeError` em `Get5` -> a versao do SW usa uma assinatura diferente
  do metodo. Tentem `Get4` ou `Get3` (ver comentario em `extractor.py`).
- Material nativo vem `None` -> a peca provavelmente nao tem material
  atribuido ainda no SolidWorks (Material Editor).
- `GetActiveObject` falha -> o SolidWorks nao esta aberto, ou esta aberto
  mas sem nenhum documento carregado.

### 2. Teste com propriedades FSAE_* preenchidas manualmente

Adicionem manualmente (via SolidWorks: File > Properties > Custom) as
propriedades `FSAE_System`, `FSAE_PN_Base`, `FSAE_Suffix`, `FSAE_Details`,
`FSAE_Process_1`, `FSAE_Process_1_Use`, `FSAE_Process_1_Qty` numa peca de
teste, e rodem `python -m sw_extract.test_connection` de novo (a partir de
`files\`) -- confirmem que os valores aparecem certinho na saida.

### 3. Teste de travessia numa montagem pequena

Peguem uma sub-montagem pequena e conhecida (ex: so o sistema de freio) e
testem `traverse_assembly()` isoladamente antes de ir para o carro
inteiro -- e mais facil conferir se a contagem de quantidade e a lista de
pecas batem com o que voces esperam manualmente.

### 4. Pipeline completo

So depois dos passos acima funcionarem, rodem:

```
python extract_and_build.py --template caminho\para\template.xlsx --out caminho\para\saida.xlsx
```

## O que ainda depende de trabalho manual (por design, combinado com o time)

- Preencher `FSAE_System`, `FSAE_Process_N` etc. em cada peca (isso e
  exatamente o que a GUI de padronizacao, ainda por construir, vai
  facilitar -- por enquanto e preenchimento manual via
  File > Properties > Custom no SolidWorks).
- Confirmar as sugestoes de material do corretor fuzzy -- o script nunca
  aplica uma correspondencia sozinho sem sinalizar como sugestao.
- Fixadores e Ferramental continuam sendo preenchidos manualmente na aba
  gerada (mesmo escopo combinado desde o bom_builder.py).

## Arquivos

- `sw_extract/connector.py` -- conexao COM com o SolidWorks
- `sw_extract/extractor.py` -- leitura de custom properties, material
  nativo, massa, e travessia da arvore de montagem
- `sw_extract/material_matcher.py` -- corretor fuzzy (material do SW ->
  nome exato do catalogo)
- `sw_extract/translator.py` -- converte dados extraidos no formato que o
  `bom_builder.py` espera
- `sw_extract/test_connection.py` -- script de teste isolado (rodar
  primeiro)
- `extract_and_build.py` -- pipeline completo (rodar por ultimo)
