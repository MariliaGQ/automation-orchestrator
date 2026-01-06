"""
Módulo de Utilitários para Agendamento de Processos
====================================================

Este módulo fornece funções utilitárias para manipulação de datas/horários
e verificação de agendamentos de processos no sistema de orquestração.

Funcionalidades principais:
    - Obtenção das partes do tempo atual (ano, mês, semana, dia, hora, minuto)
    - Normalização de valores de células do Excel
    - Verificação se um processo deve ser executado em determinado momento
    - Conversão de linhas da planilha em itens de fila de processamento

Autor: Sistema de Orquestração
Versão: 1.0
"""

from __future__ import annotations

import math
import locale
import time
from dataclasses import dataclass
from typing import Any, Mapping


# =============================================================================
# CONSTANTES E VARIÁVEIS GLOBAIS
# =============================================================================

# Flag para controlar se o locale já foi inicializado
# Evita múltiplas chamadas desnecessárias a locale.setlocale()
_LOCALE_INITIALIZED = False


# =============================================================================
# CLASSES DE DADOS
# =============================================================================

@dataclass(frozen=True, slots=True)
class NowParts:
    """
    Classe de dados que representa as partes do tempo atual.
    
    Esta classe é utilizada para avaliar agendamentos de processos,
    comparando cada dimensão temporal com os critérios de agendamento
    definidos na planilha de configuração.
    
    Atributos:
        year (str): Ano atual no formato 'YYYY' (ex: '2026').
        month_name (str): Nome do mês atual no idioma do sistema (ex: 'Janeiro').
        week_of_month (str): Semana do mês atual, de '1' a '5'.
        weekday_name (str): Nome do dia da semana no idioma do sistema (ex: 'Segunda-feira').
        day (str): Dia do mês com dois dígitos (ex: '06').
        hour (str): Hora atual com dois dígitos, formato 24h (ex: '14').
        minute (str): Minuto atual com dois dígitos (ex: '30').
    
    Observações:
        - A classe é imutável (frozen=True) para garantir integridade dos dados.
        - Usa slots para otimização de memória.
        - Os nomes de mês e dia da semana dependem do locale do sistema operacional
          (ex: 'December' em inglês vs 'Dezembro' em português).
    
    Exemplo de uso:
        >>> partes = get_now_parts()
        >>> print(partes.year)
        '2026'
        >>> print(partes.month_name)
        'Janeiro'
    """

    year: str           # Ano no formato 'YYYY'
    month_name: str     # Nome do mês (depende do locale)
    week_of_month: str  # Semana do mês (1 a 5)
    weekday_name: str   # Nome do dia da semana (depende do locale)
    day: str            # Dia do mês (01 a 31)
    hour: str           # Hora (00 a 23)
    minute: str         # Minuto (00 a 59)


# =============================================================================
# FUNÇÕES DE OBTENÇÃO DE TEMPO
# =============================================================================

def get_now_parts() -> NowParts:
    """
    Obtém as partes do tempo atual como strings formatadas.
    
    Esta função retorna um objeto NowParts contendo todas as dimensões
    temporais necessárias para avaliar se um processo deve ser executado.
    
    O locale do sistema é configurado na primeira chamada para garantir
    que os nomes de meses e dias da semana sejam retornados no idioma
    correto do sistema operacional (português no Windows Brasil).
    
    Returns:
        NowParts: Objeto contendo ano, nome do mês, semana do mês,
                  dia da semana, dia, hora e minuto atuais.
    
    Observações:
        - A semana do mês é calculada dividindo o dia por 7 e somando 1.
        - O locale só é configurado uma vez para evitar overhead.
        - Em caso de falha na configuração do locale, usa o padrão do sistema.
    
    Exemplo de uso:
        >>> agora = get_now_parts()
        >>> print(f"Estamos em {agora.month_name} de {agora.year}")
        'Estamos em Janeiro de 2026'
    """
    global _LOCALE_INITIALIZED
    
    # Inicializa o locale apenas uma vez para performance
    if not _LOCALE_INITIALIZED:
        try:
            # Configura o locale para o padrão do sistema operacional
            # Isso garante que %B (mês) e %A (dia da semana) retornem
            # nomes no idioma correto (ex: 'Janeiro' em vez de 'January')
            locale.setlocale(locale.LC_TIME, "")
        except Exception:
            # Se falhar, mantém o locale padrão
            pass
        _LOCALE_INITIALIZED = True

    # Obtém o dia atual como inteiro para calcular a semana do mês
    dia_atual = int(time.strftime("%d"))
    
    # Calcula a semana do mês (1 a 5)
    # Fórmula: (dia - 1) // 7 + 1
    # Exemplo: dia 1-7 = semana 1, dia 8-14 = semana 2, etc.
    semana_do_mes = (dia_atual - 1) // 7 + 1

    # Retorna o objeto NowParts com todas as dimensões temporais
    return NowParts(
        year=time.strftime("%Y"),           # Ano com 4 dígitos
        month_name=time.strftime("%B"),     # Nome completo do mês
        week_of_month=str(semana_do_mes),   # Semana do mês calculada
        weekday_name=time.strftime("%A"),   # Nome completo do dia da semana
        day=time.strftime("%d"),            # Dia com 2 dígitos
        hour=time.strftime("%H"),           # Hora em formato 24h
        minute=time.strftime("%M"),         # Minuto com 2 dígitos
    )


# =============================================================================
# FUNÇÕES AUXILIARES DE VALIDAÇÃO E NORMALIZAÇÃO
# =============================================================================

def _is_nan(value: Any) -> bool:
    """
    Verifica se um valor é NaN (Not a Number).
    
    Esta função é utilizada para tratar valores inválidos que podem
    vir da leitura de planilhas Excel, onde células vazias podem
    ser interpretadas como float NaN.
    
    Args:
        value (Any): Valor a ser verificado.
    
    Returns:
        bool: True se o valor for float NaN, False caso contrário.
    
    Observações:
        - Apenas valores do tipo float podem ser NaN.
        - Em caso de exceção, retorna False por segurança.
    
    Exemplo de uso:
        >>> import math
        >>> _is_nan(math.nan)
        True
        >>> _is_nan(42)
        False
        >>> _is_nan(None)
        False
    """
    try:
        return isinstance(value, float) and math.isnan(value)
    except Exception:
        return False


def _normalize_cell(value: Any) -> str:
    """
    Normaliza o valor de uma célula para uma string comparável.
    
    Esta função trata os diversos tipos de dados que podem vir
    da leitura de planilhas Excel, convertendo-os para strings
    padronizadas que podem ser comparadas consistentemente.
    
    Args:
        value (Any): Valor da célula a ser normalizado.
    
    Returns:
        str: Valor normalizado como string.
    
    Regras de normalização:
        - None ou NaN → string vazia ''
        - Float inteiro (ex: 7.0) → string do inteiro '7'
        - Demais tipos → string com espaços removidos (strip)
    
    Observações:
        - Essa normalização evita divergências comuns na leitura do Excel,
          como horas e dias vindo como float em vez de inteiro.
        - O strip() remove espaços em branco no início e fim.
    
    Exemplo de uso:
        >>> _normalize_cell(7.0)
        '7'
        >>> _normalize_cell('  texto  ')
        'texto'
        >>> _normalize_cell(None)
        ''
    """
    # Trata valores nulos e NaN
    if value is None or _is_nan(value):
        return ""
    
    # Converte floats inteiros para string sem casas decimais
    # Exemplo: 7.0 → '7' em vez de '7.0'
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    
    # Para demais tipos, converte para string e remove espaços
    return str(value).strip()


def _split_values(raw: str) -> list[str]:
    """
    Divide uma string que pode conter múltiplos valores separados.
    
    Esta função permite que campos de agendamento aceitem múltiplos
    valores em uma única célula, separados por diferentes delimitadores.
    
    Args:
        raw (str): String contendo um ou mais valores.
    
    Returns:
        list[str]: Lista de valores individuais, sem espaços em branco.
    
    Separadores suportados:
        - Vírgula (,)
        - Ponto e vírgula (;)
        - Pipe (|)
    
    Observações:
        - Valores vazios após o split são ignorados.
        - Espaços em branco são removidos de cada valor.
        - Retorna lista vazia se a string de entrada for vazia.
    
    Exemplo de uso:
        >>> _split_values('07, 08, 09')
        ['07', '08', '09']
        >>> _split_values('Segunda;Terça|Quarta')
        ['Segunda', 'Terça', 'Quarta']
        >>> _split_values('')
        []
    """
    # Retorna lista vazia para strings vazias
    if not raw:
        return []

    # Inicia com a string original como único token
    tokens = [raw]
    
    # Itera sobre cada separador e divide os tokens existentes
    for separador in (";", "|", ","):
        novos_tokens: list[str] = []
        for token in tokens:
            novos_tokens.extend(token.split(separador))
        tokens = novos_tokens

    # Remove espaços e filtra tokens vazios
    return [t.strip() for t in tokens if t.strip()]


# =============================================================================
# FUNÇÕES DE COMPARAÇÃO DE AGENDAMENTO
# =============================================================================

def _matches(field_value: Any, now_value: str) -> bool:
    """
    Verifica se um campo de agendamento corresponde ao valor atual.
    
    Esta função é o núcleo da lógica de agendamento, determinando
    se um valor configurado na planilha corresponde ao momento atual.
    
    Args:
        field_value (Any): Valor do campo de agendamento da planilha.
        now_value (str): Valor atual da dimensão temporal correspondente.
    
    Returns:
        bool: True se o campo corresponde ao valor atual, False caso contrário.
    
    Padrões aceitos:
        - 'Todos' (case-insensitive): Sempre corresponde (coringa).
        - Vazio/NaN: Nunca corresponde (obriga configuração explícita).
        - Valor único (ex: '07' ou '7'): Compara por equivalência numérica
          ou por substring para compatibilidade.
        - Múltiplos valores (ex: '07,08' ou '07;08'): Corresponde se
          qualquer um dos valores corresponder.
    
    Observações:
        - A comparação numérica permite que '7' corresponda a '07'.
        - A comparação por substring mantém compatibilidade com o
          comportamento original do sistema.
    
    Exemplo de uso:
        >>> _matches('Todos', '14')
        True
        >>> _matches('07,08,09', '08')
        True
        >>> _matches('7', '07')
        True
        >>> _matches('', '14')
        False
    """
    # Normaliza o valor do campo
    raw = _normalize_cell(field_value)
    
    # Campo vazio não corresponde a nada
    if not raw:
        return False
    
    # 'Todos' é um coringa que sempre corresponde
    if raw.lower() == "todos":
        return True

    def token_matches(token: str, current: str) -> bool:
        """
        Função interna que verifica se um token individual corresponde.
        
        Args:
            token (str): Token do campo de agendamento.
            current (str): Valor atual a ser comparado.
        
        Returns:
            bool: True se o token corresponde ao valor atual.
        """
        token = token.strip()
        current = current.strip()
        
        # Tokens ou valores vazios não correspondem
        if not token or not current:
            return False

        # Para campos numéricos, aceita equivalência numérica
        # Isso permite que '7' corresponda a '07'
        if token.isdigit() and current.isdigit():
            try:
                return int(token) == int(current)
            except Exception:
                pass

        # Para campos não numéricos, verifica se o valor atual
        # está contido no token (compatibilidade com comportamento original)
        return current in token

    # Divide o campo em múltiplos valores e verifica cada um
    candidatos = _split_values(raw)
    
    if candidatos:
        # Retorna True se qualquer candidato corresponder
        return any(token_matches(token, now_value) for token in candidatos)
    
    # Se não houver candidatos após o split, compara diretamente
    return token_matches(raw, now_value)


# =============================================================================
# FUNÇÕES PRINCIPAIS DE AGENDAMENTO
# =============================================================================

def should_enqueue(row: Mapping[str, Any], now: NowParts) -> bool:
    """
    Determina se um processo deve ser enfileirado para execução.
    
    Esta função avalia todas as dimensões de agendamento de uma linha
    da planilha e determina se o processo correspondente deve ser
    executado no momento atual.
    
    Args:
        row (Mapping[str, Any]): Linha da planilha contendo os campos
                                  de agendamento (dict-like).
        now (NowParts): Objeto com as partes do tempo atual.
    
    Returns:
        bool: True se TODAS as dimensões de agendamento corresponderem
              ao momento atual, False caso contrário.
    
    Campos de agendamento verificados:
        - ano: Ano de execução
        - meses_do_ano: Meses em que o processo deve executar
        - semanas_do_mes: Semanas do mês (1 a 5)
        - dias_da_semana: Dias da semana (Segunda, Terça, etc.)
        - dia: Dia do mês (1 a 31)
        - hora: Hora de execução (0 a 23)
        - minuto: Minuto de execução (0 a 59)
    
    Observações:
        - O processo só é enfileirado se TODAS as dimensões corresponderem.
        - Campos com valor 'Todos' sempre correspondem.
        - Campos vazios nunca correspondem.
    
    Exemplo de uso:
        >>> row = {'ano': '2026', 'hora': '14', 'minuto': '30', ...}
        >>> agora = get_now_parts()
        >>> if should_enqueue(row, agora):
        ...     print("Processo deve ser executado agora!")
    """
    # Verifica cada dimensão de agendamento
    # Todas devem corresponder para o processo ser enfileirado
    return (
        _matches(row.get("ano"), now.year)
        and _matches(row.get("meses_do_ano"), now.month_name)
        and _matches(row.get("semanas_do_mes"), now.week_of_month)
        and _matches(row.get("dias_da_semana"), now.weekday_name)
        and _matches(row.get("dia"), now.day)
        and _matches(row.get("hora"), now.hour)
        and _matches(row.get("minuto"), now.minute)
    )


# =============================================================================
# FUNÇÕES DE CONVERSÃO DE DADOS
# =============================================================================

def to_process_item(row: Mapping[str, Any]) -> dict[str, str]:
    """
    Converte uma linha da planilha em um item de fila de processamento.
    
    Esta função extrai os campos relevantes de uma linha da planilha
    e os normaliza em um dicionário padronizado que pode ser utilizado
    pela fila de processamento do orquestrador.
    
    Args:
        row (Mapping[str, Any]): Linha da planilha contendo os dados
                                  do processo (dict-like).
    
    Returns:
        dict[str, str]: Dicionário com as seguintes chaves:
            - processo: Nome identificador do processo
            - ferramenta: Ferramenta utilizada para execução
            - caminho: Caminho do arquivo/script a ser executado
    
    Observações:
        - Os valores são normalizados usando _normalize_cell().
        - Campos ausentes resultam em strings vazias.
        - As chaves do dicionário de saída são sempre minúsculas.
    
    Campos esperados na linha de entrada:
        - Nome_Processo: Nome identificador do processo
        - Ferramenta: Ferramenta de execução (ex: Python, Batch)
        - Caminho: Caminho completo do arquivo a ser executado
    
    Exemplo de uso:
        >>> row = {
        ...     'Nome_Processo': 'Processo de Backup',
        ...     'Ferramenta': 'Python',
        ...     'Caminho': 'C:/scripts/backup.py'
        ... }
        >>> item = to_process_item(row)
        >>> print(item)
        {
            'processo': 'Processo de Backup',
            'ferramenta': 'Python',
            'caminho': 'C:/scripts/backup.py'
        }
    """
    return {
        "processo": _normalize_cell(row.get("Nome_Processo")),
        "ferramenta": _normalize_cell(row.get("Ferramenta")),
        "caminho": _normalize_cell(row.get("Caminho")),
    }


