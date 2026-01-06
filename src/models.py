from __future__ import annotations

"""Modelos de dados do Orchestrador.

Este módulo concentra as estruturas de dados (dataclasses) usadas pelo
orquestrador, GUI e camada de persistência.

Objetivos:
    - Definir contratos simples e estáveis entre camadas.
    - Facilitar serialização (ex.: para SQLite) e exibição na interface.
"""

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class ProcessConfig:
    """Configuração de um processo cadastrável/executável.

    A classe reflete a estrutura persistida na tabela `processes`.

    Observações:
        - Os campos de agenda são strings para manter compatibilidade com a
          entrada do usuário e com a lógica de comparação do scheduler.
        - `enabled` controla se o processo está ativo.
    """

    id: int | None
    Nome_Processo: str
    Ferramenta: str
    Caminho: str
    ano: str
    meses_do_ano: str
    semanas_do_mes: str
    dias_da_semana: str
    dia: str
    hora: str
    minuto: str
    enabled: bool = True


@dataclass(frozen=True, slots=True)
class LogEntry:
    """Entrada de log persistida no banco.

    `stream` diferencia o tipo do log (ex.: 'log', 'stdout', 'stderr').
    """

    id: int | None
    ts_iso: str
    process_id: int | None
    stream: str
    message: str
