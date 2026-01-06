"""Testes do módulo de execução (executor).

Valida:
    - Montagem do comando UiPath
    - Comportamento de `run_item` para diferentes ferramentas
"""

import subprocess
import time

import pytest

from executor import build_uipath_command, run_item


def test_build_uipath_command():
    """Garante que o comando UiPath mantenha o contrato atual."""
    assert build_uipath_command("robot.exe", "Proc") == [
        "robot.exe",
        "execute",
        "--process-name",
        "Proc",
    ]


def test_run_item_uipath_calls_subprocess(monkeypatch):
    """Garante que UiPath execute com captura de saída e espera."""
    calls = []

    def fake_run(cmd, check, capture_output=False, text=False):
        calls.append((cmd, check, capture_output, text))
        return subprocess.CompletedProcess(cmd, 0)

    monkeypatch.setattr(subprocess, "run", fake_run)
    monkeypatch.setattr(time, "sleep", lambda *_args, **_kwargs: None)

    item = {
        "processo": "Proc",
        "ferramenta": "Uipath",
        "caminho": "robot.exe",
    }
    run_item(item)

    assert calls
    assert calls[0][0] == ["robot.exe", "execute", "--process-name", "Proc"]
    assert calls[0][1] is True


def test_run_item_python_calls_subprocess(monkeypatch):
    """Garante que itens não-UiPath usem execução direta."""
    calls = []

    def fake_run(cmd, check, capture_output=False, text=False):
        calls.append((cmd, check, capture_output, text))
        return subprocess.CompletedProcess(cmd, 0)

    monkeypatch.setattr(subprocess, "run", fake_run)

    item = {
        "processo": "Proc",
        "ferramenta": "Python",
        "caminho": "script.exe",
    }
    run_item(item)

    assert calls
    assert calls[0][0] == ["script.exe"]
    assert calls[0][1] is True


def test_run_item_invalid_item_raises():
    """Garante que itens incompletos gerem ValueError."""
    with pytest.raises(ValueError):
        run_item({"ferramenta": "Python"})
