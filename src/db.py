from __future__ import annotations

"""Camada de persistência (SQLite) do Orchestrador.

Este módulo encapsula:
    - Inicialização do schema SQLite
    - CRUD de processos
    - Persistência e consulta de logs

Notas de design:
    - O banco é um arquivo SQLite local (por padrão `orch.sqlite3`).
    - As operações são feitas com context manager para garantir commit/close.
"""

import os
import sqlite3
from contextlib import contextmanager
from dataclasses import asdict
from datetime import datetime, timezone
from typing import Iterable, Iterator

from models import LogEntry, ProcessConfig


SCHEMA_SQL = """
PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS processes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    Nome_Processo TEXT NOT NULL,
    Ferramenta TEXT NOT NULL,
    Caminho TEXT NOT NULL,

    ano TEXT NOT NULL DEFAULT 'Todos',
    meses_do_ano TEXT NOT NULL DEFAULT 'Todos',
    semanas_do_mes TEXT NOT NULL DEFAULT 'Todos',
    dias_da_semana TEXT NOT NULL DEFAULT 'Todos',
    dia TEXT NOT NULL DEFAULT 'Todos',
    hora TEXT NOT NULL DEFAULT 'Todos',
    minuto TEXT NOT NULL DEFAULT 'Todos',

    enabled INTEGER NOT NULL DEFAULT 1
);

CREATE INDEX IF NOT EXISTS idx_processes_enabled ON processes(enabled);

CREATE TABLE IF NOT EXISTS logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ts_iso TEXT NOT NULL,
    process_id INTEGER NULL,
    stream TEXT NOT NULL,
    message TEXT NOT NULL,
    FOREIGN KEY(process_id) REFERENCES processes(id) ON DELETE SET NULL
);

CREATE INDEX IF NOT EXISTS idx_logs_ts ON logs(ts_iso);
CREATE INDEX IF NOT EXISTS idx_logs_process_id ON logs(process_id);
"""


def default_db_path() -> str:
    """Resolve o caminho padrão do banco.

    Prioriza a variável de ambiente `ORCH_DB_PATH`. Caso não exista, usa
    `orch.sqlite3` no diretório de trabalho atual.
    """
    return os.getenv("ORCH_DB_PATH") or os.path.join(os.getcwd(), "orch.sqlite3")


class OrchestratorDB:
    """Acesso ao banco SQLite do Orchestrador.

    Responsável por:
        - Criar/garantir o schema
        - CRUD de processos
        - Inserção e leitura de logs
    """

    def __init__(self, db_path: str | None = None) -> None:
        self.db_path = db_path or default_db_path()
        self._init_schema()

    @contextmanager
    def connect(self) -> Iterator[sqlite3.Connection]:
        """Abre uma conexão SQLite e garante commit e close.

        Yields:
            sqlite3.Connection: Conexão com `row_factory` configurado.
        """
        os.makedirs(os.path.dirname(self.db_path) or ".", exist_ok=True)
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
            conn.commit()
        finally:
            conn.close()

    def _init_schema(self) -> None:
        """Garante que o schema do banco exista (idempotente)."""
        with self.connect() as conn:
            conn.executescript(SCHEMA_SQL)

    def list_processes(self, *, enabled_only: bool = False) -> list[ProcessConfig]:
        """Lista processos cadastrados.

        Args:
            enabled_only: Quando True, retorna apenas processos ativos.

        Returns:
            Lista de `ProcessConfig` ordenada por nome (case-insensitive).
        """
        sql = "SELECT * FROM processes"
        params: tuple[object, ...] = ()
        if enabled_only:
            sql += " WHERE enabled = 1"
        sql += " ORDER BY Nome_Processo COLLATE NOCASE"

        with self.connect() as conn:
            rows = conn.execute(sql, params).fetchall()
        return [self._row_to_process(r) for r in rows]

    def add_process(self, proc: ProcessConfig) -> int:
        """Insere um novo processo e retorna o id gerado."""
        payload = asdict(proc)
        payload.pop("id", None)

        cols = ",".join(payload.keys())
        placeholders = ",".join(["?"] * len(payload))
        values = list(payload.values())

        with self.connect() as conn:
            cur = conn.execute(
                f"INSERT INTO processes ({cols}) VALUES ({placeholders})",
                values,
            )
            return int(cur.lastrowid)

    def get_process(self, process_id: int) -> ProcessConfig | None:
        """Busca um processo por id.

        Returns:
            `ProcessConfig` se encontrado; caso contrário None.
        """
        with self.connect() as conn:
            row = conn.execute("SELECT * FROM processes WHERE id = ?", (process_id,)).fetchone()
        if row is None:
            return None
        return self._row_to_process(row)

    def update_process(self, proc: ProcessConfig) -> None:
        """Atualiza um processo existente.

        Raises:
            ValueError: Se `proc.id` for None.
        """
        if proc.id is None:
            raise ValueError("ProcessConfig.id é obrigatório para update")

        payload = asdict(proc)
        process_id = int(payload.pop("id"))

        assignments = ",".join([f"{col} = ?" for col in payload.keys()])
        values = list(payload.values())
        values.append(process_id)

        with self.connect() as conn:
            conn.execute(
                f"UPDATE processes SET {assignments} WHERE id = ?",
                values,
            )

    def delete_process(self, process_id: int) -> None:
        """Remove um processo do banco (delete físico)."""
        with self.connect() as conn:
            conn.execute("DELETE FROM processes WHERE id = ?", (process_id,))

    def set_enabled(self, process_id: int, enabled: bool) -> None:
        """Habilita/desabilita um processo cadastrado."""
        with self.connect() as conn:
            conn.execute(
                "UPDATE processes SET enabled = ? WHERE id = ?",
                (1 if enabled else 0, process_id),
            )

    def append_log(
        self,
        *,
        message: str,
        stream: str = "log",
        process_id: int | None = None,
        ts_iso: str | None = None,
    ) -> int:
        """Insere uma linha de log.

        Args:
            message: Mensagem a persistir.
            stream: Tipo/origem do log (ex.: 'log', 'stdout', 'stderr').
            process_id: Id do processo relacionado (quando aplicável).
            ts_iso: Timestamp ISO-8601. Se None, usa UTC now().

        Returns:
            Id do log inserido.
        """
        ts_iso = ts_iso or datetime.now(timezone.utc).isoformat()
        with self.connect() as conn:
            cur = conn.execute(
                "INSERT INTO logs (ts_iso, process_id, stream, message) VALUES (?, ?, ?, ?)",
                (ts_iso, process_id, stream, message),
            )
            return int(cur.lastrowid)

    def list_logs(self, *, limit: int = 1000) -> list[LogEntry]:
        """Lista logs mais recentes (ordem decrescente por id)."""
        with self.connect() as conn:
            rows = conn.execute(
                "SELECT id, ts_iso, process_id, stream, message FROM logs ORDER BY id DESC LIMIT ?",
                (limit,),
            ).fetchall()
        entries = [LogEntry(None, "", None, "", "")]
        entries.clear()
        for r in rows:
            entries.append(
                LogEntry(
                    id=int(r["id"]),
                    ts_iso=str(r["ts_iso"]),
                    process_id=(int(r["process_id"]) if r["process_id"] is not None else None),
                    stream=str(r["stream"]),
                    message=str(r["message"]),
                )
            )
        return entries

    def list_logs_between(
        self,
        start_ts_iso: str,
        end_ts_iso: str,
        *,
        limit: int | None = None,
    ) -> list[LogEntry]:
        """Lista logs dentro de um intervalo (inclusive).

        Args:
            start_ts_iso: Timestamp inicial (ISO-8601).
            end_ts_iso: Timestamp final (ISO-8601).
            limit: Limite máximo opcional.
        """
        sql = (
            "SELECT id, ts_iso, process_id, stream, message "
            "FROM logs "
            "WHERE ts_iso >= ? AND ts_iso <= ? "
            "ORDER BY ts_iso ASC, id ASC"
        )
        params: tuple[object, ...]
        if limit is None:
            params = (start_ts_iso, end_ts_iso)
        else:
            sql += " LIMIT ?"
            params = (start_ts_iso, end_ts_iso, int(limit))

        with self.connect() as conn:
            rows = conn.execute(sql, params).fetchall()

        entries: list[LogEntry] = []
        for r in rows:
            entries.append(
                LogEntry(
                    id=int(r["id"]),
                    ts_iso=str(r["ts_iso"]),
                    process_id=(int(r["process_id"]) if r["process_id"] is not None else None),
                    stream=str(r["stream"]),
                    message=str(r["message"]),
                )
            )
        return entries

    @staticmethod
    def _row_to_process(row: sqlite3.Row) -> ProcessConfig:
        """Converte sqlite3.Row em ProcessConfig."""
        return ProcessConfig(
            id=int(row["id"]),
            Nome_Processo=str(row["Nome_Processo"]),
            Ferramenta=str(row["Ferramenta"]),
            Caminho=str(row["Caminho"]),
            ano=str(row["ano"]),
            meses_do_ano=str(row["meses_do_ano"]),
            semanas_do_mes=str(row["semanas_do_mes"]),
            dias_da_semana=str(row["dias_da_semana"]),
            dia=str(row["dia"]),
            hora=str(row["hora"]),
            minuto=str(row["minuto"]),
            enabled=bool(int(row["enabled"])),
        )


def process_to_schedule_row(proc: ProcessConfig) -> dict[str, str]:
    """Converte `ProcessConfig` para o formato esperado pelo scheduler.

    O scheduler trabalha com um "row" (dict) com as mesmas chaves usadas na
    lógica de agendamento em `util.should_enqueue` e na conversão em item de fila
    em `util.to_process_item`.
    """

    return {
        "Nome_Processo": proc.Nome_Processo,
        "Ferramenta": proc.Ferramenta,
        "Caminho": proc.Caminho,
        "ano": proc.ano,
        "meses_do_ano": proc.meses_do_ano,
        "semanas_do_mes": proc.semanas_do_mes,
        "dias_da_semana": proc.dias_da_semana,
        "dia": proc.dia,
        "hora": proc.hora,
        "minuto": proc.minuto,
    }
