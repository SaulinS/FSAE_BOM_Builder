"""
sw_extract/connector.py

Conexao com uma instancia do SolidWorks ja aberta (via COM/win32com).

IMPORTANTE (leia antes de rodar):
- Este modulo so funciona no Windows, com SolidWorks instalado.
- Precisa do pacote pywin32:  pip install pywin32
- O SolidWorks PRECISA estar aberto (com o arquivo/montagem carregado) antes
  de rodar o script -- este modulo conecta na instancia ja em execucao, nao
  abre o SolidWorks sozinho por padrao (ha uma opcao para isso, comentada
  abaixo, mas abrir programaticamente é mais lento e mais fragil).
- TESTADO ao vivo pela primeira vez em 2026-08-12, contra SOLIDWORKS 2025
  (RevisionNumber 33.5.0) com uma peca simples aberta -- achado confirmado:
  metodos da API SEM parametros (GetType, GetTitle, GetPathName, etc.) sao
  expostos pelo dispatch dinamico do win32com como PROPRIEDADES, nao
  metodos -- chamar com "()" da TypeError tipo "'int' object is not
  callable". A correcao (usada neste arquivo e em extractor.py) e acessar
  sem parenteses: `model.GetType`, nao `model.GetType()`. Isso NAO foi
  testado ainda numa montagem (so numa peca), entao os metodos usados em
  traverse_assembly() (extractor.py) ainda podem ter surpresas -- rodem o
  script de teste (test_connection.py) numa peca simples primeiro (ja
  validado), depois test_connection numa peca com as propriedades FSAE_*
  preenchidas (Get5 ainda nao foi confirmado contra uma propriedade real),
  e so depois traverse_assembly() numa montagem pequena.
"""

import win32com.client


# Constantes da API do SolidWorks (swDocumentTypes_e)
SW_DOC_PART = 1
SW_DOC_ASSEMBLY = 2
SW_DOC_DRAWING = 3


def connect_to_solidworks(start_if_not_running=False):
    """
    Conecta a uma instancia do SolidWorks ja aberta. Se `start_if_not_running`
    for True e nao houver instancia aberta, tenta iniciar uma nova (mais
    lento, e o SolidWorks pode levar 10-30s para carregar).
    """
    try:
        sw_app = win32com.client.GetActiveObject("SldWorks.Application")
        return sw_app
    except Exception:
        if not start_if_not_running:
            raise RuntimeError(
                "Nao foi possivel conectar a uma instancia do SolidWorks em "
                "execucao. Abra o SolidWorks com o arquivo desejado antes de "
                "rodar este script, ou chame connect_to_solidworks(True) "
                "para tentar abrir uma instancia nova."
            )
        sw_app = win32com.client.Dispatch("SldWorks.Application")
        sw_app.Visible = True
        return sw_app


def get_active_document(sw_app):
    """Retorna o documento atualmente ativo (o que esta em foco no SW)."""
    model = sw_app.ActiveDoc
    if model is None:
        raise RuntimeError(
            "Nenhum documento aberto/ativo no SolidWorks. Abra a peca ou "
            "montagem desejada e deixe a janela em foco."
        )
    return model


def describe_document(model):
    """Info basica de diagnostico -- util pro script de teste."""
    doc_type = model.GetType  # sem parenteses -- ver aviso no topo do arquivo
    type_name = {
        SW_DOC_PART: "Peca (Part)",
        SW_DOC_ASSEMBLY: "Montagem (Assembly)",
        SW_DOC_DRAWING: "Desenho (Drawing)",
    }.get(doc_type, f"Desconhecido ({doc_type})")
    return {
        "title": model.GetTitle,
        "path": model.GetPathName,
        "type": type_name,
        "type_code": doc_type,
    }
