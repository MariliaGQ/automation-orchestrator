from __future__ import annotations

"""Execução de itens (processos) via subprocess no Windows.

Este módulo centraliza a montagem e execução de comandos para diferentes
ferramentas/formatos de arquivos.

Regras atuais:
    - UiPath: usa `robot.exe execute --process-name <nome>` e aguarda 20s.
    - .py/.pyw: executa com o `sys.executable` atual.
    - .bat/.cmd: executa com `cmd.exe /c`.
    - .ps1: executa com PowerShell (Bypass).
    - .lnk: usa `Start-Process -Wait` via PowerShell.
"""

import os
import shlex
import subprocess
import sys
import time
from typing import Any


def build_uipath_command(robot_path: str, process_name: str) -> list[str]:
    """Monta o comando de execução do UiPath.

    Args:
        robot_path: Caminho do executável do Robot (ex.: robot.exe).
        process_name: Nome do processo publicado no UiPath.

    Returns:
        Lista de argumentos para uso com `subprocess.run`.
    """

    return [robot_path, "execute", "--process-name", process_name]


def _split_command_windows(raw: str) -> list[str]:
    """Divide um comando bruto em tokens no Windows.

    Motivo:
        Caminhos podem conter espaços e vir sem aspas. Esta função tenta
        reconstruir o executável (token 0) quando necessário.

    Args:
        raw: Texto do comando/caminho.

    Returns:
        Lista de tokens, onde o primeiro deve ser o executável/arquivo.
    """
    raw = raw.strip()
    if not raw:
        return []

    if os.path.exists(raw):
        return [raw]

    try:
        tokens = shlex.split(raw, posix=False)
    except ValueError:
        tokens = [raw]

    if len(tokens) > 1 and not raw.lstrip().startswith(("\"", "'")):
        for i in range(len(tokens), 0, -1):
            candidate = " ".join(tokens[:i])
            if os.path.isfile(candidate):
                return [candidate, *tokens[i:]]

        for i in range(len(tokens), 0, -1):
            candidate = " ".join(tokens[:i])
            if os.path.exists(candidate):
                return [candidate, *tokens[i:]]

    return tokens


def build_subprocess_command(item: dict[str, Any]) -> list[str]:
    """Converte um item de fila em um comando executável (lista de tokens).

    Args:
        item: Dicionário com chaves esperadas: processo, ferramenta, caminho.

    Returns:
        Lista de tokens para `subprocess.run`.

    Raises:
        ValueError: Se o item estiver incompleto ou o comando for inválido.
    """
    tool = str(item.get("ferramenta") or "").strip()
    path = str(item.get("caminho") or "").strip()
    process_name = str(item.get("processo") or "").strip()

    if not tool or not path:
        raise ValueError("Item inválido: ferramenta/caminho ausentes")

    if tool.lower() == "uipath":
        return build_uipath_command(path, process_name)

    tokens = _split_command_windows(path)
    if not tokens:
        raise ValueError("Item inválido: comando vazio")

    entry = tokens[0]
    ext = os.path.splitext(entry)[1].lower()

    if ext in (".py", ".pyw"):
        return [sys.executable, *tokens]

    if ext in (".bat", ".cmd"):
        return ["cmd.exe", "/c", *tokens]

    if ext == ".ps1":
        return [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            entry,
            *tokens[1:],
        ]

    if ext == ".lnk":
        def _ps_quote(value: str) -> str:
            return "'" + value.replace("'", "''") + "'"

        arg_list = "@(" + ",".join(_ps_quote(a) for a in tokens[1:]) + ")" if tokens[1:] else "@()"
        return [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            f"Start-Process -FilePath {_ps_quote(entry)} -ArgumentList {arg_list} -Wait",
        ]

    if os.path.exists(entry) and os.path.isfile(entry) and ext not in (".exe", ".com"):
        raise ValueError(
            f"Caminho aponta para arquivo não-executável no Windows: '{entry}' (extensão '{ext}'). "
            "Use .exe/.bat/.cmd/.ps1/.py ou um atalho .lnk para um executável."
        )

    return tokens


def run_item(item: dict[str, Any]) -> None:
    """Executa um item de fila.

    Mantém o contrato atual do projeto:
        - UiPath: captura saída e espera um tempo adicional.
        - Demais comandos: execução direta (check=True).

    Args:
        item: Dicionário com chaves esperadas: processo, ferramenta, caminho.

    Raises:
        subprocess.CalledProcessError: Se o comando retornar código != 0.
        ValueError: Se o item estiver inválido.
    """

    tool = str(item.get("ferramenta") or "").strip()
    cmd = build_subprocess_command(item)

    if tool.lower() == "uipath":
        subprocess.run(cmd, check=True, capture_output=True, text=True)
        # Mantém o comportamento atual: aguarda para o Robot terminar de liberar recursos.
        time.sleep(20)
        return

    subprocess.run(cmd, check=True)
