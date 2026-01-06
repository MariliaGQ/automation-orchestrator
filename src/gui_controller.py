from __future__ import annotations

"""Controller da GUI do Orchestrador.

Este módulo implementa a camada de controle (no padrão MVC-ish) responsável por:
    - CRUD de processos via `OrchestratorDB`
    - Controle do scheduler (liga/desliga e tick)
    - Execução de itens usando `QProcess`
    - Emissão de sinais para atualizar a interface (logs, status, fila)

Observações:
    - O controller tenta ser resiliente: falhas de DB/log não devem derrubar a GUI.
    - A execução é feita via `QProcess` para integração com o event loop do Qt.
"""

import os
import time
import re
import locale
from typing import Any

from PySide6 import QtCore

from orchestrator import DBBackedScheduler, InMemoryQueue
from db import OrchestratorDB
from executor import build_subprocess_command
from db import process_to_schedule_row
from models import LogEntry, ProcessConfig
from util import NowParts, should_enqueue


class OrchestratorController(QtCore.QObject):
    """Orquestra operações da GUI.

    Expõe sinais para a interface e métodos para:
        - Gerenciar processos
        - Gerenciar o scheduler
        - Enfileirar e executar itens
        - Consultar logs e agenda do dia
    """

    console_text = QtCore.Signal(str)   # stdout/stderr do processo em execução
    log_text = QtCore.Signal(str)       # mensagens persistidas no DB
    status_text = QtCore.Signal(str)    # status curto para a UI
    processes_changed = QtCore.Signal() # avisar para recarregar lista
    scheduler_state_changed = QtCore.Signal(bool)  # True=ligado, False=desligado
    process_running_changed = QtCore.Signal(bool)  # True=executando RPA, False=idle
    queue_changed = QtCore.Signal(object)          # list[dict]
    running_item_changed = QtCore.Signal(object)   # dict|None

    def __init__(self, db_path: str | None = None, parent: QtCore.QObject | None = None) -> None:
        """Inicializa controller, DB, fila, scheduler e QProcess.

        Args:
            db_path: Caminho do SQLite (opcional). Se None, usa padrão do DB.
            parent: QObject pai (Qt).
        """
        super().__init__(parent)
        self.db = OrchestratorDB(db_path)
        self.queue = InMemoryQueue()

        self._scheduler = DBBackedScheduler(self.db, self.queue, poll_seconds=60, log_to_db=True)
        self._scheduler_timer = QtCore.QTimer(self)
        self._scheduler_timer.setInterval(60_000)
        self._scheduler_timer.timeout.connect(self._on_scheduler_tick)

        self._process = QtCore.QProcess(self)
        self._process.started.connect(self._on_started)
        self._process.readyReadStandardOutput.connect(self._on_stdout)
        self._process.readyReadStandardError.connect(self._on_stderr)
        self._process.finished.connect(self._on_finished)
        self._process.errorOccurred.connect(self._on_error)

        self._running_scheduler = False
        self._running_item: dict[str, Any] | None = None

        self._ansi_escape_re = re.compile(r"\x1b\[[0-?]*[ -/]*[@-~]")

    def _decode_process_output(self, data: bytes) -> str:
        """Decodifica bytes do `QProcess` em texto legível.

        Motivo:
            Diferentes ferramentas podem escrever em UTF-8, UTF-16 ou codepages
            do Windows. Essa função aplica heurísticas e remove sequências ANSI.

        Args:
            data: Bytes brutos do stdout/stderr.

        Returns:
            Texto normalizado, sem caracteres de controle indesejados.
        """
        if not data:
            return ""

        # Alguns programas (e principalmente PowerShell) podem cuspir UTF-16.
        if data.startswith((b"\xff\xfe", b"\xfe\xff")):
            try:
                text = data.decode("utf-16")
                text = text.replace("\r\n", "\n").replace("\r", "\n")
                text = self._ansi_escape_re.sub("", text)
                text = "".join(ch for ch in text if ch in ("\n", "\t") or ord(ch) >= 32)
                return text
            except Exception:
                pass

        # Heurística: se muitos bytes nulos, provavelmente UTF-16 LE.
        if len(data) >= 8 and data[1::2].count(0) > (len(data) // 6):
            try:
                text = data.decode("utf-16-le")
                text = text.replace("\r\n", "\n").replace("\r", "\n")
                text = self._ansi_escape_re.sub("", text)
                text = "".join(ch for ch in text if ch in ("\n", "\t") or ord(ch) >= 32)
                return text
            except Exception:
                pass

        # 1) Tenta UTF-8 (comum em Python/PowerShell moderno)
        try:
            text = data.decode("utf-8")
        except UnicodeDecodeError:
            # 2) Fallback para codepages do Windows (programas nativos variam entre OEM/ANSI)
            oem_enc = None
            acp_enc = None
            try:
                import ctypes

                oem_enc = f"cp{ctypes.windll.kernel32.GetOEMCP()}"
                acp_enc = f"cp{ctypes.windll.kernel32.GetACP()}"
            except Exception:
                pass

            candidates: list[str] = []
            # Preferir ACP (cp1252) ajuda em muitos outputs de apps Windows.
            if acp_enc:
                candidates.append(acp_enc)
            if oem_enc and oem_enc not in candidates:
                candidates.append(oem_enc)

            pref = locale.getpreferredencoding(False)
            if pref and pref not in candidates:
                candidates.append(pref)

            candidates.extend(["cp1252", "cp850", "latin-1"])

            # Decodificações podem "funcionar" mas gerar mojibake.
            # Escolhe a melhor entre as candidatas usando uma heurística simples.
            suspects = set("ßÝÚ═þÿ")
            accents = set("áàâãäéèêëíìîïóòôõöúùûüçÁÀÂÃÄÉÈÊËÍÌÎÏÓÒÔÕÖÚÙÛÜÇ")

            decoded: list[tuple[str, str]] = []
            for enc in candidates:
                try:
                    decoded.append((enc, data.decode(enc)))
                except Exception:
                    continue

            if not decoded:
                text = data.decode(errors="replace")
            else:
                def score(s: str) -> int:
                    suspect_count = sum(1 for ch in s if ch in suspects)
                    accent_count = sum(1 for ch in s if ch in accents)
                    return accent_count * 3 - suspect_count * 8

                text = max(decoded, key=lambda it: score(it[1]))[1]

        # Normaliza newlines e remove artefatos que aparecem como 'caracteres especiais'
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        text = self._ansi_escape_re.sub("", text)
        text = "".join(ch for ch in text if ch in ("\n", "\t") or ord(ch) >= 32)
        return text

    # -------------------- Processos (CRUD) --------------------
    def list_processes(self) -> list[ProcessConfig]:
        """Retorna todos os processos cadastrados (ativos e inativos)."""
        return self.db.list_processes(enabled_only=False)

    def save_process(self, proc: ProcessConfig) -> int:
        """Cria ou atualiza um processo.

        Args:
            proc: Configuração do processo.

        Returns:
            Id do processo (novo ou existente).
        """
        if proc.id is None:
            new_id = self.db.add_process(proc)
            self.processes_changed.emit()
            return new_id

        self.db.update_process(proc)
        self.processes_changed.emit()
        return int(proc.id)

    def delete_process(self, process_id: int) -> None:
        """Exclui um processo e notifica a UI."""
        self.db.delete_process(process_id)
        self.processes_changed.emit()

    # -------------------- Scheduler --------------------
    def start_scheduler(self) -> None:
        """Liga o scheduler e executa um tick imediato."""
        if self._running_scheduler:
            return
        self._running_scheduler = True
        self._scheduler_timer.start()
        self.scheduler_state_changed.emit(True)
        self._emit_queue_changed()
        self.status_text.emit("Scheduler: ligado")
        self._on_scheduler_tick()

    def stop_scheduler(self) -> None:
        """Desliga o scheduler."""
        if not self._running_scheduler:
            return
        self._running_scheduler = False
        self._scheduler_timer.stop()
        self.scheduler_state_changed.emit(False)
        self._emit_queue_changed()
        self.status_text.emit("Scheduler: desligado")

    def is_scheduler_running(self) -> bool:
        """Indica se o scheduler está ligado."""
        return self._running_scheduler

    def _on_scheduler_tick(self) -> None:
        """Tick do scheduler (chamado pelo QTimer).

        Regras mantidas:
            - Fora do horário 07-18: não executa.
            - Em caso de erro: registra log e segue.
        """
        hour = int(time.strftime("%H"))
        if not (7 <= hour < 18):
            self.status_text.emit("Fora do horário (07-18): scheduler em espera")
            return

        try:
            self._scheduler.tick_once()
        except Exception as exc:
            self._append_log(f"Erro no scheduler: {exc}")

        self._emit_queue_changed()

        self._drain_queue_if_idle()

    # -------------------- Execução --------------------
    def enqueue_manual(self, item: dict[str, Any]) -> None:
        """Enfileira um item manualmente e inicia execução se estiver idle."""
        self.queue.put(item)
        self._emit_queue_changed()
        self._append_log(f"Enfileirado manual: {item.get('processo')} ({item.get('ferramenta')})")
        self._drain_queue_if_idle()

    def stop_current_process(self) -> None:
        """Solicita cancelamento do processo em execução (quando permitido)."""
        if self._process.state() == QtCore.QProcess.ProcessState.NotRunning:
            return

        if not self.can_cancel_current_process():
            self._append_log("Cancelamento disponível apenas para execuções Python.")
            return

        self._append_log("Cancelando execução Python...")
        self._process.kill()

    def is_process_running(self) -> bool:
        """Indica se há um processo em execução via QProcess."""
        return self._process.state() != QtCore.QProcess.ProcessState.NotRunning

    def can_cancel_current_process(self) -> bool:
        """Indica se o botão 'Cancelar execução' deve estar habilitado.

        Requisito: cancelar só deve funcionar para Python.
        """

        if self._process.state() == QtCore.QProcess.ProcessState.NotRunning:
            return False

        item = self._running_item
        if not isinstance(item, dict) or not item:
            return False

        tool = str(item.get("ferramenta") or "").strip().lower()
        if tool == "python":
            return True

        path = str(item.get("caminho") or "").strip()
        ext = os.path.splitext(path)[1].lower()
        return ext in {".py", ".pyw"}

    def _drain_queue_if_idle(self) -> None:
        """Se não houver execução em andamento, consome a fila e inicia o próximo."""
        if self._process.state() != QtCore.QProcess.ProcessState.NotRunning:
            return
        if self.queue.empty():
            self._emit_queue_changed()
            self.status_text.emit("Fila vazia")
            return

        try:
            item = self.queue.get()
        except Exception:
            return

        self._emit_queue_changed()

        self._running_item = item
        self.running_item_changed.emit(item)
        self._start_item(item)

    def _start_item(self, item: dict[str, Any]) -> None:
        """Inicia a execução do item via QProcess."""
        cmd = build_subprocess_command(item)
        if not cmd:
            self._append_log("Comando vazio (não executado)")
            self._running_item = None
            return

        program, args = cmd[0], cmd[1:]
        self.console_text.emit(f"\n$ {program} {' '.join(args)}\n")
        self._append_log(f"Executando: {item.get('processo')} ({item.get('ferramenta')})")

        self._process.setProgram(program)
        self._process.setArguments([str(a) for a in args])
        self._process.start()

        self.status_text.emit(f"Executando: {item.get('processo')}")

    def _on_started(self) -> None:
        """Slot chamado quando o QProcess inicia."""
        self.process_running_changed.emit(True)

    def _on_stdout(self) -> None:
        """Slot chamado quando há dados em stdout."""
        data = self._decode_process_output(bytes(self._process.readAllStandardOutput()))
        if data:
            self.console_text.emit(data)
            self._append_log(data.rstrip("\n"), stream="stdout", persist_raw=False)

    def _on_stderr(self) -> None:
        """Slot chamado quando há dados em stderr."""
        data = self._decode_process_output(bytes(self._process.readAllStandardError()))
        if data:
            self.console_text.emit(data)
            self._append_log(data.rstrip("\n"), stream="stderr", persist_raw=False)

    def _on_error(self, _err: QtCore.QProcess.ProcessError) -> None:
        """Slot chamado quando o QProcess reporta erro de inicialização/execução."""
        self._append_log("Falha ao iniciar o processo (QProcess)")
        self.process_running_changed.emit(False)
        self._running_item = None
        self.running_item_changed.emit(None)

    def _on_finished(self, exit_code: int, _exit_status: QtCore.QProcess.ExitStatus) -> None:
        """Slot chamado quando o processo termina."""
        item = self._running_item
        self._running_item = None

        self.process_running_changed.emit(False)
        self.running_item_changed.emit(None)

        if exit_code == 0:
            self._append_log(f"Finalizado com sucesso: {item.get('processo') if item else ''}")
        else:
            self._append_log(f"Finalizado com erro (code={exit_code}): {item.get('processo') if item else ''}")

        self._drain_queue_if_idle()

    def _emit_queue_changed(self) -> None:
        """Emite snapshot da fila para a UI, sem deixar exceções vazarem."""
        try:
            self.queue_changed.emit(self.queue.snapshot())
        except Exception:
            # nunca quebra a UI por falha de observabilidade
            pass

    def get_running_item(self) -> dict[str, Any] | None:
        """Retorna o item atualmente em execução (ou None)."""
        return self._running_item

    def get_queue_snapshot(self) -> list[dict[str, Any]]:
        """Retorna um snapshot da fila atual (para exibição na UI)."""
        return self.queue.snapshot()

    def list_today_schedule(self) -> list[dict[str, str]]:
        """Lista execuções previstas hoje (07:00-18:00) em granularidade de 1 minuto."""

        lt = time.localtime()
        y, m, d = lt.tm_year, lt.tm_mon, lt.tm_mday
        start = int(time.mktime((y, m, d, 7, 0, 0, 0, 0, -1)))
        end = int(time.mktime((y, m, d, 18, 0, 0, 0, 0, -1)))

        procs = self.db.list_processes(enabled_only=True)
        proc_rows = [(p, process_to_schedule_row(p)) for p in procs]

        out: list[dict[str, str]] = []
        for ts in range(start, end, 60):
            tl = time.localtime(ts)
            day_int = int(time.strftime("%d", tl))
            week_of_month = (day_int - 1) // 7 + 1
            now_parts = NowParts(
                year=time.strftime("%Y", tl),
                month_name=time.strftime("%B", tl),
                week_of_month=str(week_of_month),
                weekday_name=time.strftime("%A", tl),
                day=time.strftime("%d", tl),
                hour=time.strftime("%H", tl),
                minute=time.strftime("%M", tl),
            )

            for proc, row in proc_rows:
                if should_enqueue(row, now_parts):
                    out.append(
                        {
                            "hora": f"{now_parts.hour}:{now_parts.minute}",
                            "processo": proc.Nome_Processo,
                            "ferramenta": proc.Ferramenta,
                            "caminho": proc.Caminho,
                        }
                    )

        out.sort(key=lambda it: (it["hora"], it["processo"].casefold()))
        return out

    # -------------------- Logs --------------------
    def list_logs_text(self, limit: int = 500) -> list[str]:
        """Lista logs como strings prontas para exibição.

        Args:
            limit: Quantidade máxima de linhas (mais recentes).
        """
        # logs vêm DESC; devolve em ordem cronológica pra exibir melhor
        entries = list(reversed(self.db.list_logs(limit=limit)))
        return [f"{e.ts_iso} [{e.stream}] {e.message}" for e in entries]

    def list_logs_entries_between(self, start_ts_iso: str, end_ts_iso: str, *, limit: int | None = None) -> list[LogEntry]:
        """Lista logs entre dois timestamps ISO, retornando objetos LogEntry."""
        return self.db.list_logs_between(start_ts_iso, end_ts_iso, limit=limit)

    def _append_log(self, message: str, *, stream: str = "log", persist_raw: bool = True) -> None:
        """Persiste (quando possível) e emite uma mensagem de log para a UI."""
        msg = message if persist_raw else message
        try:
            self.db.append_log(message=msg, stream=stream)
        except Exception:
            # não deixa a UI quebrar por falha no DB
            pass
        self.log_text.emit(message)
