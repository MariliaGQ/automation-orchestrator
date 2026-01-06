from __future__ import annotations

"""Scheduler e fila do Orchestrador.

Este módulo contém a lógica de agendamento (scheduler) que consulta o banco,
avalia se processos estão "na hora" e insere itens em uma fila.

Regras do scheduler (mantidas):
    - Opera entre 07:00 e 18:00.
    - Verifica a cada `poll_seconds`.
    - Evita duplicidade do mesmo processo no mesmo minuto.
"""

import logging
import queue as queue_mod
import time
from collections import deque
from typing import Any, Deque

from db import OrchestratorDB, process_to_schedule_row
from util import NowParts, get_now_parts, should_enqueue, to_process_item


class InMemoryQueue:
    """Fila FIFO simples para uso na GUI (single-thread/event-loop).

    A implementação usa `collections.deque` por ser leve e eficiente.
    """

    def __init__(self) -> None:
        self._dq: Deque[dict[str, Any]] = deque()

    def put(self, item: dict[str, Any]) -> None:
        """Insere um item no fim da fila."""
        self._dq.append(item)

    def get(self) -> dict[str, Any]:
        """Remove e retorna o item do início da fila.

        Raises:
            queue.Empty: Se a fila estiver vazia.
        """
        if not self._dq:
            raise queue_mod.Empty
        return self._dq.popleft()

    def empty(self) -> bool:
        """Indica se a fila está vazia."""
        return not self._dq

    def __len__(self) -> int:
        return len(self._dq)

    def snapshot(self) -> list[dict[str, Any]]:
        """Retorna uma cópia do estado atual da fila (para UI/debug)."""
        return list(self._dq)


def poll_due_processes(db: OrchestratorDB, now_parts: NowParts) -> list[dict[str, str]]:
    """Consulta o banco e retorna itens prontos para enfileirar.

    Args:
        db: Instância do banco.
        now_parts: Partes do tempo atual usadas na comparação.

    Returns:
        Lista de itens (dict) já normalizados para execução.
    """

    due: list[dict[str, str]] = []
    for proc in db.list_processes(enabled_only=True):
        row = process_to_schedule_row(proc)
        if should_enqueue(row, now_parts):
            due.append(to_process_item(row))
    return due


class DuplicateGuard:
    """Evita enfileirar o mesmo processo mais de uma vez no mesmo minuto."""

    def __init__(self) -> None:
        self._last_minute_key: str | None = None
        self._seen: set[str] = set()

    def reset_if_new_minute(self, now_parts: NowParts) -> None:
        """Reseta o conjunto de itens vistos quando o minuto muda."""
        minute_key = f"{now_parts.year}{now_parts.month_name}{now_parts.day}{now_parts.hour}{now_parts.minute}"
        if minute_key != self._last_minute_key:
            self._last_minute_key = minute_key
            self._seen.clear()

    def allow(self, item: dict[str, Any], now_parts: NowParts) -> bool:
        """Indica se o item pode ser enfileirado (deduplicação por minuto)."""
        self.reset_if_new_minute(now_parts)
        key = f"{item.get('processo','')}|{item.get('ferramenta','')}|{item.get('caminho','')}"
        if key in self._seen:
            return False
        self._seen.add(key)
        return True


class DBBackedScheduler:
    """Scheduler baseado em DB.

    Regras iguais ao projeto atual:
    - roda entre 07:00 e 18:00
    - verifica a cada N segundos
    """

    def __init__(
        self,
        db: OrchestratorDB,
        out_queue: Any,
        *,
        poll_seconds: int = 60,
        log_to_db: bool = True,
    ) -> None:
        self.db = db
        self.out_queue = out_queue
        self.poll_seconds = poll_seconds
        self.log_to_db = log_to_db
        self._guard = DuplicateGuard()

    def tick_once(self) -> None:
        """Executa um ciclo de verificação e enfileiramento (uma "batida")."""
        now_parts = get_now_parts()
        for item in poll_due_processes(self.db, now_parts):
            if not self._guard.allow(item, now_parts):
                continue
            logging.info("SCHEDULER - Inserindo na fila: %s", item.get("processo"))
            if self.log_to_db:
                self.db.append_log(
                    stream="log",
                    message=f"Enfileirado: {item.get('processo')} ({item.get('ferramenta')})",
                )
            self.out_queue.put(item)

    def loop(self) -> None:
        """Loop bloqueante (útil para modo CLI/serviço, se necessário)."""
        while 7 <= int(time.strftime("%H")) < 18:
            self.tick_once()
            time.sleep(self.poll_seconds)


def item_as_dict(item: dict[str, Any]) -> dict[str, Any]:
    """Normalização defensiva de item (útil para logs/debug).

    Converte `None` para string vazia para evitar ruídos na exibição.
    """

    return {k: ("" if v is None else v) for k, v in item.items()}
