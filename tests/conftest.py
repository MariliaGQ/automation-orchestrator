"""Configuração de testes (pytest).

Este arquivo ajusta o `sys.path` para que os módulos em `src/` possam ser
importados nos testes sem instalação do pacote.
"""

import sys
from pathlib import Path


# Garante que imports (ex.: `executor`, `util`) apontem para src/
PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"

if str(SRC_DIR) not in sys.path:
    sys.path.insert(0, str(SRC_DIR))
