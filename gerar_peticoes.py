#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gerador Automático de Petições — Tema 20/TST — PREVI/Banco do Brasil
Versão 2.0

Gera petições .docx individualizadas a partir de modelos Word e planilha de dados.
"""

import os
import sys
import re
import copy
import datetime
import traceback
from pathlib import Path

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from docx import Document


# ──────────────────────────────────────────────
# Constantes
# ──────────────────────────────────────────────

PASTA_ENTRADA = Path("entrada")
PASTA_SAIDA = Path("saida")

ARQUIVO_PLANILHA = PASTA_ENTRADA / "dados_clientes.xlsx"
ABA_DADOS = "Dados Clientes"

MODELOS = {
    ("M", 1): PASTA_ENTRADA / "modelo_masculino_1RT.docx",
    ("M", 2): PASTA_ENTRADA / "modelo_masculino_2RT.docx",
    ("F", 1): PASTA_ENTRADA / "modelo_feminino_1RT.docx",
    ("F", 2): PASTA_ENTRADA / "modelo_feminino_2RT.docx",
}

COLUNAS_OBRIGATORIAS = [
    "Nome", "Genero", "QualificacaoAutor", "DataDesligamento",
    "ValorBeneficio", "RT_Verbas_1", "RT_Indenizatoria", "TrechoPrescricao",
]

# Colunas usadas na substituição de variáveis
COLUNAS_PETICAO = [
    "Nome", "QualificacaoAutor", "DataDesligamento", "ValorBeneficio",
    "RT_Verbas_1", "RT_Verbas_2", "RT_Indenizatoria",
    "PreservacaoSP", "TrechoPrescricao",
]

# Colunas informacionais — ignoradas pelo sistema
COLUNAS_IGNORADAS = [
    "Verbas_1", "TransitoJulgado_1", "Verbas_2", "TransitoJulgado_2",
    "Grupo", "DataReferenciaSTJ",
]

MESES_PT = [
    "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro",
]

CARACTERES_PROIBIDOS = re.compile(r'[<>:"/\\|?*]')

MARCADOR_INICIO = "{SE:PreservacaoSP}"
MARCADOR_FIM = "{FIM_SE:PreservacaoSP}"


# ──────────────────────────────────────────────
# Logger
# ──────────────────────────────────────────────

class Logger:
    """Grava em arquivo e imprime no terminal simultaneamente."""

    def __init__(self, caminho: Path):
        self.arquivo = open(caminho, "w", encoding="utf-8")

    def log(self, msg: str = ""):
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        linha = f"[{ts}] {msg}"
        print(linha)
        self.arquivo.write(linha + "\n")
        self.arquivo.flush()

    def close(self):
        self.arquivo.close()


# ──────────────────────────────────────────────
# Data por extenso
# ──────────────────────────────────────────────

def data_por_extenso(dt: datetime.date) -> str:
    """Retorna '14 de março de 2026'."""
    return f"{dt.day} de {MESES_PT[dt.month - 1]} de {dt.year}"


# ──────────────────────────────────────────────
# Leitura da planilha
# ──────────────────────────────────────────────

def ler_planilha(caminho: Path, logger: Logger):
    """Lê a aba 'Dados Clientes' e retorna (cabeçalhos, lista_de_dicts)."""
    wb = openpyxl.load_workbook(caminho, data_only=True)
    if ABA_DADOS not in wb.sheetnames:
        logger.log(f"ERRO: aba '{ABA_DADOS}' não encontrada na planilha.")
        sys.exit(1)
    ws = wb[ABA_DADOS]

    # Cabeçalhos na linha 1
    cabecalhos = []
    for cell in ws[1]:
        val = cell.value
        if val is not None:
            cabecalhos.append(str(val).strip())
        else:
            cabecalhos.append("")

    logger.log(f"Colunas encontradas: {cabecalhos}")

    # Validar colunas obrigatórias
    faltantes = [c for c in COLUNAS_OBRIGATORIAS if c not in cabecalhos]
    if faltantes:
        logger.log(f"ERRO: colunas obrigatórias ausentes: {faltantes}")
        sys.exit(1)

    # Ler dados — tudo como texto
    clientes = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        # Pular linhas completamente vazias
        if all(v is None for v in row):
            continue
        registro = {}
        for i, cab in enumerate(cabecalhos):
            if not cab:
                continue
            val = row[i] if i < len(row) else None
            if val is None:
                registro[cab] = ""
            elif isinstance(val, datetime.datetime):
                # Formatar data como texto dd/mm/yyyy para preservar formato
                registro[cab] = val.strftime("%-d/%m/%Y").replace("/0", "/").lstrip("0") if False else val.strftime("%d/%m/%Y").lstrip("0").replace("/0", "/")
            else:
                registro[cab] = str(val).strip()
        clientes.append(registro)

    wb.close()
    return cabecalhos, clientes


# ──────────────────────────────────────────────
# Substituição em documentos Word
# ──────────────────────────────────────────────

def substituir_variaveis_paragrafo(paragrafo, variaveis: dict) -> list:
    """
    Substitui variáveis {xxx} num parágrafo tratando fragmentação de runs.
    Retorna lista de variáveis não encontradas (residuais).
    """
    # Concatenar texto de todos os runs
    texto_completo = "".join(run.text for run in paragrafo.runs)

    if "{" not in texto_completo:
        return []

    texto_novo = texto_completo
    for var_nome, var_valor in variaveis.items():
        placeholder = "{" + var_nome + "}"
        texto_novo = texto_novo.replace(placeholder, var_valor)

    if texto_novo == texto_completo:
        return []

    # Colocar texto resultante no primeiro run, esvaziar os demais
    if paragrafo.runs:
        paragrafo.runs[0].text = texto_novo
        for run in paragrafo.runs[1:]:
            run.text = ""

    return []


def formatar_nome_no_paragrafo(paragrafo, nome: str):
    """Aplica negrito e versalete ao nome do cliente no parágrafo."""
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from copy import deepcopy

    if not nome or not paragrafo.runs:
        return

    run0 = paragrafo.runs[0]
    texto = run0.text

    if nome not in texto:
        return

    idx = texto.index(nome)
    antes = texto[:idx]
    depois = texto[idx + len(nome):]

    def criar_run(texto_novo, rPr_base=None):
        """Cria um w:r limpo com texto, copiando apenas o w:rPr do run base."""
        r = OxmlElement("w:r")
        if rPr_base is not None:
            rPr_orig = rPr_base.find(qn("w:rPr"))
            if rPr_orig is not None:
                r.append(deepcopy(rPr_orig))
        t = OxmlElement("w:t")
        t.text = texto_novo
        if texto_novo and (texto_novo[0] == " " or texto_novo[-1] == " "):
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        r.append(t)
        return r

    # Run original fica só com o texto antes do nome
    run0.text = antes

    # Run do nome: copia só o rPr base e adiciona bold + smallCaps
    run_nome_el = criar_run(nome, rPr_base=run0._r)
    rPr = run_nome_el.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        run_nome_el.insert(0, rPr)
    rPr.append(OxmlElement("w:b"))
    rPr.append(OxmlElement("w:smallCaps"))

    run0._r.addnext(run_nome_el)

    # Run do texto depois do nome (formatação normal)
    if depois:
        run_depois_el = criar_run(depois, rPr_base=run0._r)
        run_nome_el.addnext(run_depois_el)


def aplicar_formatacao_nome_documento(doc, nome: str):
    """Aplica negrito e versalete ao nome em todo o documento."""
    for paragrafo in doc.paragraphs:
        formatar_nome_no_paragrafo(paragrafo, nome)
    for tabela in doc.tables:
        for row in tabela.rows:
            for cell in row.cells:
                for paragrafo in cell.paragraphs:
                    formatar_nome_no_paragrafo(paragrafo, nome)
    for secao in doc.sections:
        for hf in [secao.header, secao.footer, secao.first_page_header,
                   secao.first_page_footer, secao.even_page_header, secao.even_page_footer]:
            if hf:
                try:
                    for paragrafo in hf.paragraphs:
                        formatar_nome_no_paragrafo(paragrafo, nome)
                except Exception:
                    pass


def encontrar_residuais_paragrafo(paragrafo) -> list:
    """Retorna lista de variáveis {xxx} residuais no parágrafo."""
    texto = "".join(run.text for run in paragrafo.runs)
    return re.findall(r"\{[A-Za-z_][A-Za-z0-9_]*\}", texto)


def processar_bloco_condicional(doc, preservacao_vazia: bool):
    """
    Remove ou mantém o bloco entre {SE:PreservacaoSP} e {FIM_SE:PreservacaoSP}.
    Opera no corpo principal do documento.
    """
    _processar_bloco_em_paragrafos(doc.element.body, preservacao_vazia)


def _processar_bloco_em_paragrafos(parent_element, preservacao_vazia: bool):
    """Processa bloco condicional nos parágrafos filhos de um elemento XML."""
    from docx.oxml.ns import qn

    paragrafos = parent_element.findall(qn("w:p"))
    idx_inicio = None
    idx_fim = None

    for i, p in enumerate(paragrafos):
        texto = "".join(node.text or "" for node in p.iter(qn("w:t")))
        if MARCADOR_INICIO in texto:
            idx_inicio = i
        if MARCADOR_FIM in texto:
            idx_fim = i

    if idx_inicio is None or idx_fim is None:
        return  # Marcadores não encontrados — nada a fazer

    if preservacao_vazia:
        # Remover TUDO do marcador início ao marcador fim (inclusive)
        for i in range(idx_fim, idx_inicio - 1, -1):
            parent_element.remove(paragrafos[i])
    else:
        # Manter conteúdo, remover apenas os parágrafos dos marcadores
        parent_element.remove(paragrafos[idx_fim])
        parent_element.remove(paragrafos[idx_inicio])


def substituir_variaveis_documento(doc, variaveis: dict) -> list:
    """
    Substitui variáveis em todo o documento: corpo, tabelas, cabeçalhos e rodapés.
    Retorna lista de variáveis residuais.
    """
    residuais = []

    # Corpo principal
    for paragrafo in doc.paragraphs:
        substituir_variaveis_paragrafo(paragrafo, variaveis)

    # Tabelas
    for tabela in doc.tables:
        for row in tabela.rows:
            for cell in row.cells:
                for paragrafo in cell.paragraphs:
                    substituir_variaveis_paragrafo(paragrafo, variaveis)

    # Cabeçalhos e rodapés de todas as seções
    for secao in doc.sections:
        for header in [secao.header, secao.first_page_header, secao.even_page_header]:
            if header and header.is_linked_to_previous is False or header:
                try:
                    for paragrafo in header.paragraphs:
                        substituir_variaveis_paragrafo(paragrafo, variaveis)
                    for tabela in header.tables:
                        for row in tabela.rows:
                            for cell in row.cells:
                                for paragrafo in cell.paragraphs:
                                    substituir_variaveis_paragrafo(paragrafo, variaveis)
                except Exception:
                    pass

        for footer in [secao.footer, secao.first_page_footer, secao.even_page_footer]:
            if footer:
                try:
                    for paragrafo in footer.paragraphs:
                        substituir_variaveis_paragrafo(paragrafo, variaveis)
                    for tabela in footer.tables:
                        for row in tabela.rows:
                            for cell in row.cells:
                                for paragrafo in cell.paragraphs:
                                    substituir_variaveis_paragrafo(paragrafo, variaveis)
                except Exception:
                    pass

    # Verificar residuais
    for paragrafo in doc.paragraphs:
        residuais.extend(encontrar_residuais_paragrafo(paragrafo))
    for tabela in doc.tables:
        for row in tabela.rows:
            for cell in row.cells:
                for paragrafo in cell.paragraphs:
                    residuais.extend(encontrar_residuais_paragrafo(paragrafo))
    for secao in doc.sections:
        for hf in [secao.header, secao.footer, secao.first_page_header,
                    secao.first_page_footer, secao.even_page_header, secao.even_page_footer]:
            if hf:
                try:
                    for paragrafo in hf.paragraphs:
                        residuais.extend(encontrar_residuais_paragrafo(paragrafo))
                except Exception:
                    pass

    return residuais


# ──────────────────────────────────────────────
# Relatório de conferência
# ──────────────────────────────────────────────

def gerar_relatorio(resultados: list, caminho: Path):
    """Gera relatorio_conferencia.xlsx com cores por status."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Relatório de Conferência"

    # Cabeçalhos
    headers = ["Nº", "Nome", "Gênero", "Modelo Usado", "Arquivo Gerado",
               "Status", "Observação"]
    ws.append(headers)

    # Estilo dos cabeçalhos
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    # Cores por status
    fill_ok = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    fill_erro = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    fill_aviso = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")

    for i, r in enumerate(resultados, start=1):
        row = [
            i, r.get("nome", ""), r.get("genero", ""), r.get("modelo", ""),
            r.get("arquivo", ""), r.get("status", ""), r.get("observacao", ""),
        ]
        ws.append(row)

        # Colorir linha
        status = r.get("status", "")
        if status == "OK":
            fill = fill_ok
        elif status == "ERRO":
            fill = fill_erro
        else:
            fill = fill_aviso

        for cell in ws[i + 1]:
            cell.fill = fill

    # Ajustar largura das colunas
    for col_letter in ["A", "B", "C", "D", "E", "F", "G"]:
        ws.column_dimensions[col_letter].width = 20
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["E"].width = 50
    ws.column_dimensions["G"].width = 50

    # Filtro automático
    ws.auto_filter.ref = ws.dimensions

    # Resumo numérico
    total = len(resultados)
    ok = sum(1 for r in resultados if r["status"] == "OK")
    erros = sum(1 for r in resultados if r["status"] == "ERRO")
    avisos = sum(1 for r in resultados if r["status"] == "AVISO")

    ws.append([])
    ws.append(["Resumo:"])
    ws.append(["Total", total])
    ws.append(["OK", ok])
    ws.append(["ERRO", erros])
    ws.append(["AVISO", avisos])

    wb.save(caminho)


# ──────────────────────────────────────────────
# Função principal
# ──────────────────────────────────────────────

def main():
    # 1. Verificar arquivos de entrada
    arquivos_necessarios = [ARQUIVO_PLANILHA] + list(MODELOS.values())
    for arq in arquivos_necessarios:
        if not arq.exists():
            print(f"ERRO: arquivo de entrada não encontrado: {arq}")
            sys.exit(1)

    # 2. Criar pasta de saída e logger
    PASTA_SAIDA.mkdir(exist_ok=True)
    logger = Logger(PASTA_SAIDA / "log_execucao.txt")
    logger.log("=" * 60)
    logger.log("INÍCIO DA EXECUÇÃO — Gerador de Petições Tema 20/TST")
    logger.log("=" * 60)

    # 3. Ler planilha
    logger.log("Lendo planilha de dados...")
    cabecalhos, clientes = ler_planilha(ARQUIVO_PLANILHA, logger)
    logger.log(f"Total de registros encontrados: {len(clientes)}")

    # 4. Data da petição
    data_peticao = data_por_extenso(datetime.date.today())
    logger.log(f"Data da petição: {data_peticao}")

    # 5. Processar cada cliente
    resultados = []

    for idx, cliente in enumerate(clientes, start=1):
        nome = cliente.get("Nome", "").strip()
        genero = cliente.get("Genero", "").strip().upper()
        logger.log("-" * 40)
        logger.log(f"Cliente {idx}: {nome}")

        resultado = {
            "nome": nome,
            "genero": genero,
            "modelo": "",
            "arquivo": "",
            "status": "",
            "observacao": "",
        }

        # 5a. Validar dados mínimos
        if not nome:
            logger.log("  ERRO: Nome vazio. Pulando cliente.")
            resultado["status"] = "ERRO"
            resultado["observacao"] = "Nome vazio"
            resultados.append(resultado)
            continue

        if genero not in ("M", "F"):
            logger.log(f"  AVISO: Gênero inválido '{genero}'. Pulando cliente.")
            resultado["status"] = "AVISO"
            resultado["observacao"] = f"Gênero inválido: {genero}"
            resultados.append(resultado)
            continue

        # 5b. Ler QtdRTs
        qtd_rts_raw = cliente.get("QtdRTs", "").strip()
        try:
            qtd_rts = int(float(qtd_rts_raw))
        except (ValueError, TypeError):
            logger.log(f"  AVISO: QtdRTs inválido '{qtd_rts_raw}'. Pulando cliente.")
            resultado["status"] = "AVISO"
            resultado["observacao"] = f"QtdRTs inválido: {qtd_rts_raw}"
            resultados.append(resultado)
            continue

        if qtd_rts not in (1, 2):
            logger.log(f"  AVISO: QtdRTs = {qtd_rts}, esperado 1 ou 2. Pulando cliente.")
            resultado["status"] = "AVISO"
            resultado["observacao"] = f"QtdRTs inválido: {qtd_rts}"
            resultados.append(resultado)
            continue

        # 5c. Selecionar modelo
        chave_modelo = (genero, qtd_rts)
        caminho_modelo = MODELOS[chave_modelo]
        nome_modelo = caminho_modelo.name
        resultado["modelo"] = nome_modelo
        logger.log(f"  Modelo: {nome_modelo}")

        try:
            # 5d. Abrir cópia do modelo
            doc = Document(str(caminho_modelo))

            # 5e. Bloco condicional
            preservacao = cliente.get("PreservacaoSP", "").strip()
            preservacao_vazia = (preservacao == "")
            processar_bloco_condicional(doc, preservacao_vazia)
            if preservacao_vazia:
                logger.log("  Bloco 4.2 (preservação): REMOVIDO")
            else:
                logger.log("  Bloco 4.2 (preservação): MANTIDO")

            # 5f. Montar dicionário de variáveis
            variaveis = {}
            for col in COLUNAS_PETICAO:
                if col in cliente:
                    variaveis[col] = cliente[col]
            variaveis["DataPeticao"] = data_peticao

            # Substituir
            residuais = substituir_variaveis_documento(doc, variaveis)

            # Formatar nome em negrito e versalete
            aplicar_formatacao_nome_documento(doc, nome)

            # 5g. Verificar variáveis residuais
            observacoes = []
            if residuais:
                obs = f"Variáveis residuais: {', '.join(set(residuais))}"
                logger.log(f"  AVISO: {obs}")
                observacoes.append(obs)

            # 5h. Salvar petição
            nome_arquivo = CARACTERES_PROIBIDOS.sub("", nome)
            nome_arquivo = f"Peticao - {nome_arquivo}.docx"
            caminho_saida = PASTA_SAIDA / nome_arquivo
            doc.save(str(caminho_saida))
            resultado["arquivo"] = nome_arquivo

            # 5i. Status
            if residuais:
                resultado["status"] = "AVISO"
                resultado["observacao"] = "; ".join(observacoes)
            else:
                resultado["status"] = "OK"

            logger.log(f"  Salvo: {nome_arquivo} — {resultado['status']}")

        except Exception as e:
            logger.log(f"  ERRO ao gerar petição: {e}")
            logger.log(f"  {traceback.format_exc()}")
            resultado["status"] = "ERRO"
            resultado["observacao"] = str(e)

        resultados.append(resultado)

    # 6. Relatório de conferência
    logger.log("-" * 40)
    caminho_relatorio = PASTA_SAIDA / "relatorio_conferencia.xlsx"
    gerar_relatorio(resultados, caminho_relatorio)
    logger.log(f"Relatório de conferência gerado: {caminho_relatorio}")

    # Resumo final
    total = len(resultados)
    ok = sum(1 for r in resultados if r["status"] == "OK")
    erros = sum(1 for r in resultados if r["status"] == "ERRO")
    avisos = sum(1 for r in resultados if r["status"] == "AVISO")

    logger.log("=" * 60)
    logger.log("RESUMO FINAL")
    logger.log(f"  Total de clientes processados: {total}")
    logger.log(f"  OK:    {ok}")
    logger.log(f"  ERRO:  {erros}")
    logger.log(f"  AVISO: {avisos}")
    logger.log("=" * 60)
    logger.log("FIM DA EXECUÇÃO")
    logger.close()


if __name__ == "__main__":
    main()
