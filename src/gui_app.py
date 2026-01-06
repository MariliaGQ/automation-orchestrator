from __future__ import annotations

"""Interface gráfica (PySide6) do Orquestrador.

Este módulo implementa a GUI principal e a tela de cadastro/gerenciamento de
processos. Ele se integra ao `OrchestratorController`, que encapsula a lógica
de execução, scheduler e persistência.

Componentes principais:
    - `MainWindow`: dashboard (agenda do dia, fila, console e logs)
    - `ProcessManagerWindow`: CRUD de processos + ações manuais
    - `MultiSelectComboBox`: combobox multi-seleção com checkboxes (agenda)

Observações:
    - O estilo visual é aplicado via stylesheet (Qt Fusion + CSS).
    - Nomes de meses/dias da semana podem depender do locale do sistema.
"""

import locale
import os
import sys
from datetime import datetime, timedelta, timezone
from collections.abc import Callable
from typing import Any, cast

from PySide6 import QtCore, QtGui, QtWidgets

from gui_controller import OrchestratorController
from models import ProcessConfig


def _apply_app_style(app: QtWidgets.QApplication) -> None:
    """Aplica o estilo padrão do aplicativo.

    Define:
        - Style "Fusion" (boa consistência no Windows)
        - Fonte padrão
        - Stylesheet com regras de layout e cores

    Args:
        app: Instância da aplicação Qt.
    """
    # Visual baseado na pasta example/ (gui_styles.py), com extensões para os widgets extras do app.
    app.setStyle("Fusion")
    app.setFont(QtGui.QFont("Segoe UI", 10))
    app.setStyleSheet(
        """
        * { font-size: 13px; color: #334155; }
        QMainWindow { background-color: #F8FAFC; }

        QFrame#AppBar { background-color: #FFFFFF; border-bottom: 1px solid #E2E8F0; }
        QLabel#AppTitle { font-size: 18px; font-weight: 700; color: #0F172A; }
        QLabel#AppSubtitle { font-size: 12px; color: #64748B; }

        QLabel[role="section-title"] {
            font-size: 13px;
            font-weight: 600;
            color: #0F172A;
            padding: 2px 0px;
        }

        QPushButton {
            padding: 8px 14px;
            border-radius: 6px;
            font-weight: 600;
            border: 1px solid #CBD5E1;
            background-color: #E2E8F0;
            min-height: 34px;
        }
        QPushButton:hover { background-color: #CBD5E1; border-color: #94A3B8; }
        QPushButton:pressed { background-color: #94A3B8; }
        QPushButton:disabled { color: #94A3B8; background: #F8FAFC; }

        QPushButton#PrimaryBtn, QPushButton[variant="primary"] {
            background-color: #22C55E;
            color: #FFFFFF;
            border: none;
        }
        QPushButton#PrimaryBtn:hover, QPushButton[variant="primary"]:hover { background-color: #16A34A; }
        QPushButton#PrimaryBtn:pressed, QPushButton[variant="primary"]:pressed { background-color: #15803D; }
        QPushButton#PrimaryBtn:disabled, QPushButton[variant="primary"]:disabled {
            background-color: #F8FAFC;
            color: #94A3B8;
            border: 1px solid #E2E8F0;
        }

        QPushButton#DangerBtn, QPushButton[variant="danger"] {
            background-color: #F87171;
            color: #FFFFFF;
            border: none;
        }
        QPushButton#DangerBtn:hover, QPushButton[variant="danger"]:hover { background-color: #EF4444; }
        QPushButton#DangerBtn:pressed, QPushButton[variant="danger"]:pressed { background-color: #DC2626; }
        QPushButton#DangerBtn:disabled, QPushButton[variant="danger"]:disabled {
            background-color: #F8FAFC;
            color: #94A3B8;
            border: 1px solid #E2E8F0;
        }

        QGroupBox {
            font-weight: 700;
            border: 1px solid #E2E8F0;
            border-radius: 8px;
            margin-top: 12px;
            padding-top: 16px;
            background-color: #FFFFFF;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 6px;
            color: #334155;
            font-weight: 700;
        }

        QTableWidget, QListWidget, QTextEdit {
            background-color: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-radius: 8px;
        }
        QHeaderView::section {
            background-color: #F8FAFC;
            padding: 8px;
            border: none;
            border-bottom: 1px solid #E2E8F0;
            font-weight: 600;
        }

        QLineEdit {
            padding: 8px;
            border: 1px solid #E2E8F0;
            border-radius: 6px;
            background: #FFFFFF;
        }
        QComboBox {
            padding: 8px 10px;
            border: 1px solid #E2E8F0;
            border-radius: 6px;
            background: #FFFFFF;
        }
        QComboBox:disabled { color: #94A3B8; background: #F8FAFC; }
        QDateTimeEdit {
            padding: 8px 10px;
            border: 1px solid #E2E8F0;
            border-radius: 6px;
            background: #FFFFFF;
        }

        QTabWidget::pane {
            border: 1px solid #E2E8F0;
            border-radius: 8px;
            background: #FFFFFF;
        }
        QTabBar::tab {
            padding: 8px 14px;
            border: 1px solid #E2E8F0;
            border-bottom: none;
            border-top-left-radius: 8px;
            border-top-right-radius: 8px;
            background: #F8FAFC;
            margin-right: 4px;
        }
        QTabBar::tab:selected { background: #FFFFFF; }

        QLabel#StatusText { color: #64748B; font-size: 12px; padding: 4px 0px; }

        QFrame#StatusDot {
            min-width: 12px;
            max-width: 12px;
            min-height: 12px;
            max-height: 12px;
            border-radius: 6px;
            border: 1px solid #E2E8F0;
            background: #94A3B8;
        }
        QFrame#StatusDot[state="stopped"] { background: #94A3B8; }
        QFrame#StatusDot[state="idle"] { background: #10B981; border: 1px solid #059669; }
        QFrame#StatusDot[state="busy"] { background: #F59E0B; border: 1px solid #D97706; }
        """
    )


def _set_button_icon(button: QtWidgets.QAbstractButton, standard_pixmap: QtWidgets.QStyle.StandardPixmap) -> None:
    """Configura um ícone padrão do Qt em um botão.

    Args:
        button: Botão alvo.
        standard_pixmap: Ícone padrão (enum do Qt).
    """
    icon = button.style().standardIcon(standard_pixmap)
    button.setIcon(icon)
    button.setIconSize(QtCore.QSize(16, 16))


def _make_section_label(text: str) -> QtWidgets.QLabel:
    """Cria um label com estilo de "título de seção".

    Args:
        text: Texto do label.

    Returns:
        QLabel configurado com a propriedade `role="section-title"`.
    """
    lbl = QtWidgets.QLabel(text)
    lbl.setProperty("role", "section-title")
    return lbl


def _install_qt_message_filter() -> None:
    """Silencia mensagens ruidosas específicas do Qt no Windows.

    Em algumas máquinas o Qt emite warnings de DirectWrite para fontes legadas
    (ex.: 8514oem). Isso não quebra a GUI, mas polui o terminal.
    """

    def handler(mode: QtCore.QtMsgType, context: QtCore.QMessageLogContext, message: str) -> None:  # type: ignore[name-defined]
        if "DirectWrite: CreateFontFaceFromHDC() failed" in message:
            return
        sys.stderr.write(message + "\n")

    QtCore.qInstallMessageHandler(handler)


def _split_multi_values(raw: str) -> list[str]:
    """Divide um texto em múltiplos valores.

    Suporta separadores: vírgula (,), ponto e vírgula (;) e pipe (|).

    Observação:
        Esta função é usada apenas na GUI para ler valores já salvos no banco
        no formato "Todos" ou "v1,v2,v3".
    """

    raw = str(raw or "").strip()
    if not raw:
        return []
    tokens = [raw]
    for sep in (";", "|", ","):
        next_tokens: list[str] = []
        for t in tokens:
            next_tokens.extend(t.split(sep))
        tokens = next_tokens
    return [t.strip() for t in tokens if t.strip()]


def _localized_strftime_options(fmt: str, dates: list[datetime]) -> list[str]:
    """Gera opções usando strftime respeitando o locale atual do Python.

    Isso ajuda a evitar erro em meses/dias da semana quando o Windows está em PT/EN.
    """

    # Em Windows, o locale padrão do Python pode não seguir o idioma do SO.
    # Aqui tentamos forçar o locale padrão do SO (quando disponível) apenas
    # para gerar listas de opções coerentes com o scheduler.
    try:
        old_locale = locale.setlocale(locale.LC_TIME)
    except Exception:
        old_locale = None

    try:
        try:
            # Ajusta para o locale padrão do SO (quando disponível)
            locale.setlocale(locale.LC_TIME, "")
        except Exception:
            pass

        seen: set[str] = set()
        out: list[str] = []
        for dt in dates:
            val = dt.strftime(fmt)
            if val and val not in seen:
                seen.add(val)
                out.append(val)
        return out
    finally:
        if old_locale:
            try:
                locale.setlocale(locale.LC_TIME, old_locale)
            except Exception:
                pass


def _month_name_options() -> list[str]:
    """Lista nomes de meses no idioma do sistema (quando possível)."""

    base = datetime(2000, 1, 1)
    dates = [datetime(base.year, m, 1) for m in range(1, 13)]
    return _localized_strftime_options("%B", dates)


def _weekday_name_options() -> list[str]:
    """Lista nomes dos dias da semana no idioma do sistema (quando possível)."""

    # 2000-01-03 foi uma segunda-feira
    base = datetime(2000, 1, 3)
    dates = [base + timedelta(days=i) for i in range(7)]
    return _localized_strftime_options("%A", dates)


class _NoAutoCheckDelegate(QtWidgets.QStyledItemDelegate):
    """Impede o toggle automático do checkbox pelo delegate.

    Em alguns temas/versões do Qt, um item `ItemIsUserCheckable` pode ter o estado
    alternado automaticamente ao clicar. Como o `MultiSelectComboBox` também alterna
    manualmente (para manter regras como "Todos" exclusivo), isso pode gerar "toggle
    duplo" e dar a sensação de que a opção "não marca".
    """

    def editorEvent(
        self,
        event: QtCore.QEvent,
        model: QtCore.QAbstractItemModel,
        option: QtWidgets.QStyleOptionViewItem,
        index: QtCore.QModelIndex,
    ) -> bool:
        if event.type() in {
            QtCore.QEvent.Type.MouseButtonPress,
            QtCore.QEvent.Type.MouseButtonRelease,
            QtCore.QEvent.Type.MouseButtonDblClick,
            QtCore.QEvent.Type.KeyPress,
            QtCore.QEvent.Type.KeyRelease,
        }:
            return False
        return super().editorEvent(event, model, option, index)


class MultiSelectComboBox(QtWidgets.QComboBox):
    """ComboBox com multi-seleção via checkboxes.

    Objetivo:
        Evitar erros de digitação em campos de agendamento.

    Formato armazenado/retornado:
        - Quando "Todos" está selecionado (ou nada selecionado) -> "Todos".
        - Quando há seleções -> "v1,v2,v3".
    """

    def __init__(
        self,
        options: list[str],
        all_label: str = "Todos",
        parent: QtWidgets.QWidget | None = None,
        normalize_token: Callable[[str], str] | None = None,
    ) -> None:
        super().__init__(parent)
        self.setEditable(True)
        if self.lineEdit() is not None:
            self.lineEdit().setReadOnly(True)
            self.lineEdit().setPlaceholderText(all_label)

        self._all_label = all_label
        self._normalize_token = normalize_token
        self._skip_next_hide = False
        self.setModel(QtGui.QStandardItemModel(self))
        self.view().setItemDelegate(_NoAutoCheckDelegate(self.view()))
        # Intercepta cliques no popup para impedir o comportamento padrão do QComboBox
        # (seleção única/fechamento) e controlar o toggle via checkboxes.
        self.view().setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.NoSelection)
        self.view().viewport().installEventFilter(self)

        self._add_option(all_label, checked=True)
        for opt in options:
            opt_norm = self._norm(opt)
            if opt_norm and opt_norm != all_label:
                self._add_option(opt_norm, checked=False)
        self._update_display_text()

    def eventFilter(self, watched: QtCore.QObject, event: QtCore.QEvent) -> bool:
        if watched is self.view().viewport():
            if event.type() == QtCore.QEvent.Type.MouseButtonRelease:
                mouse_event = cast(QtGui.QMouseEvent, event)
                index = self.view().indexAt(mouse_event.position().toPoint())
                if index.isValid():
                    self._on_item_pressed(index)
                return True
            if event.type() == QtCore.QEvent.Type.MouseButtonDblClick:
                return True
        return super().eventFilter(watched, event)

    def _norm(self, text: str) -> str:
        t = str(text or "").strip()
        if not t:
            return ""
        if self._normalize_token is None:
            return t
        try:
            return str(self._normalize_token(t)).strip()
        except Exception:
            return t

    def _add_option(self, text: str, checked: bool) -> None:
        model = self.model()
        if not isinstance(model, QtGui.QStandardItemModel):
            return
        item = QtGui.QStandardItem(text)
        item.setFlags(
            QtCore.Qt.ItemFlag.ItemIsEnabled
            | QtCore.Qt.ItemFlag.ItemIsSelectable
            | QtCore.Qt.ItemFlag.ItemIsUserCheckable
        )
        item.setData(
            QtCore.Qt.CheckState.Checked if checked else QtCore.Qt.CheckState.Unchecked,
            QtCore.Qt.ItemDataRole.CheckStateRole,
        )
        model.appendRow(item)

    def _ensure_option(self, text: str) -> None:
        text = self._norm(text)
        if not text:
            return
        model = self.model()
        if not isinstance(model, QtGui.QStandardItemModel):
            return
        for i in range(model.rowCount()):
            it = model.item(i)
            if it is not None and it.text() == text:
                return
        self._add_option(text, checked=False)

    def _on_item_pressed(self, index: QtCore.QModelIndex) -> None:
        model = self.model()
        if not isinstance(model, QtGui.QStandardItemModel):
            return
        item = model.itemFromIndex(index)
        if item is None:
            return

        self._skip_next_hide = True

        new_state = (
            QtCore.Qt.CheckState.Unchecked
            if item.checkState() == QtCore.Qt.CheckState.Checked
            else QtCore.Qt.CheckState.Checked
        )
        item.setCheckState(new_state)

        all_label_norm = self._all_label.strip().lower()

        # Regra: "Todos" é exclusivo
        if item.text().strip().lower() == all_label_norm:
            if new_state == QtCore.Qt.CheckState.Checked:
                for i in range(model.rowCount()):
                    it = model.item(i)
                    if it is not None and it.text() != item.text():
                        it.setCheckState(QtCore.Qt.CheckState.Unchecked)
        else:
            # Se qualquer outro foi marcado, desmarca "Todos"
            for i in range(model.rowCount()):
                it = model.item(i)
                if it is not None and it.text().strip().lower() == all_label_norm:
                    it.setCheckState(QtCore.Qt.CheckState.Unchecked)
                    break

        # Evita estado vazio: se nada ficou marcado, volta para "Todos".
        if not self._checked_values():
            for i in range(model.rowCount()):
                it = model.item(i)
                if it is not None and it.text().strip().lower() == all_label_norm:
                    it.setCheckState(QtCore.Qt.CheckState.Checked)
                    break

        self._update_display_text()

    def hidePopup(self) -> None:
        """Evita fechar o popup ao clicar em um item (multi-seleção)."""

        if self._skip_next_hide:
            self._skip_next_hide = False
            return
        super().hidePopup()

    def _checked_values(self) -> list[str]:
        model = self.model()
        if not isinstance(model, QtGui.QStandardItemModel):
            return []
        values: list[str] = []
        for i in range(model.rowCount()):
            it = model.item(i)
            if it is None:
                continue
            if it.checkState() == QtCore.Qt.CheckState.Checked:
                values.append(it.text())
        return values

    def value_text(self) -> str:
        """Retorna o valor em formato persistível ("Todos" ou lista)."""

        values = self._checked_values()
        if not values:
            return self._all_label
        if any(v.strip().lower() == self._all_label.lower() for v in values):
            return self._all_label
        return ",".join(values)

    def set_value_text(self, raw: str) -> None:
        """Carrega o estado a partir de um texto persistido ("Todos" ou lista)."""

        raw = str(raw or "").strip()
        model = self.model()
        if not isinstance(model, QtGui.QStandardItemModel):
            return

        if not raw or raw.lower() == self._all_label.lower():
            for i in range(model.rowCount()):
                it = model.item(i)
                if it is None:
                    continue
                it.setCheckState(
                    QtCore.Qt.CheckState.Checked
                    if it.text().strip().lower() == self._all_label.lower()
                    else QtCore.Qt.CheckState.Unchecked
                )
            self._update_display_text()
            return

        tokens_raw = _split_multi_values(raw)
        tokens = [self._norm(t) for t in tokens_raw if self._norm(t)]
        for t in tokens:
            self._ensure_option(t)

        for i in range(model.rowCount()):
            it = model.item(i)
            if it is None:
                continue
            if it.text().strip().lower() == self._all_label.lower():
                it.setCheckState(QtCore.Qt.CheckState.Unchecked)
                continue
            it.setCheckState(
                QtCore.Qt.CheckState.Checked
                if any(tok.lower() == it.text().lower() for tok in tokens)
                else QtCore.Qt.CheckState.Unchecked
            )

        self._update_display_text()

    def _update_display_text(self) -> None:
        if self.lineEdit() is None:
            return
        text = self.value_text()
        self.lineEdit().setText(text)


class StatusDot(QtWidgets.QFrame):
    """Indicador visual de status do orquestrador.

    Estados:
        - stopped: scheduler desligado
        - idle: scheduler ligado e sem execução
        - busy: execução em andamento
    """

    def __init__(self, parent: QtWidgets.QWidget | None = None) -> None:
        """Inicializa o componente e animação de pulso (quando busy)."""
        super().__init__(parent)
        self.setObjectName("StatusDot")
        self.setProperty("state", "stopped")
        self.setToolTip("Orchestrador parado")

        # Feedback visual sutil quando está executando (sem mudar a UX)
        self._opacity = QtWidgets.QGraphicsOpacityEffect(self)
        self._opacity.setOpacity(1.0)
        self.setGraphicsEffect(self._opacity)

        self._pulse = QtCore.QPropertyAnimation(self._opacity, b"opacity", self)
        self._pulse.setDuration(900)
        self._pulse.setStartValue(1.0)
        self._pulse.setEndValue(0.55)
        self._pulse.setEasingCurve(QtCore.QEasingCurve.Type.InOutSine)
        self._pulse.setLoopCount(-1)

    def set_state(self, state: str) -> None:
        """Atualiza o estado visual do indicador."""
        state = state.strip().lower()
        if state not in {"stopped", "idle", "busy"}:
            state = "stopped"

        self.setProperty("state", state)
        if state == "busy":
            self.setToolTip("Executando RPA")
            if self._pulse.state() != QtCore.QAbstractAnimation.State.Running:
                self._pulse.start()
        elif state == "idle":
            self.setToolTip("Orchestrador livre")
            if self._pulse.state() == QtCore.QAbstractAnimation.State.Running:
                self._pulse.stop()
            self._opacity.setOpacity(1.0)
        else:
            self.setToolTip("Orchestrador parado")
            if self._pulse.state() == QtCore.QAbstractAnimation.State.Running:
                self._pulse.stop()
            self._opacity.setOpacity(1.0)

        self.style().unpolish(self)
        self.style().polish(self)
        self.update()


class ProcessManagerWindow(QtWidgets.QMainWindow):
    """Janela de gerenciamento (CRUD) de processos.

    Permite:
        - Cadastrar/editar/remover processos
        - Habilitar/desabilitar processos
        - Enfileirar execução manual de um processo selecionado
    """

    def __init__(self, controller: OrchestratorController) -> None:
        """Monta a janela e conecta sinais/ações ao controller."""
        super().__init__()
        self.setWindowTitle("Cadastro de processos")
        self.resize(1100, 680)

        self.controller = controller
        self.controller.processes_changed.connect(self._reload_processes)

        central = QtWidgets.QWidget(self)
        self.setCentralWidget(central)
        root = QtWidgets.QVBoxLayout(central)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(12)

        actions = QtWidgets.QHBoxLayout()
        actions.setSpacing(10)
        self.btn_run = QtWidgets.QPushButton("Executar selecionado")
        self.btn_refresh = QtWidgets.QPushButton("Recarregar")

        # Sem ícones (conforme solicitado)
        actions.addWidget(self.btn_run)
        actions.addWidget(self.btn_refresh)
        actions.addStretch(1)
        root.addLayout(actions)

        splitter_h = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        root.addWidget(splitter_h, 1)

        # Tabela de processos
        self.table = QtWidgets.QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["ID", "Ativo", "Nome", "Ferramenta", "Caminho"])
        self.table.setSelectionBehavior(QtWidgets.QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QtWidgets.QTableWidget.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.setColumnHidden(0, True)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setStretchLastSection(True)
        splitter_h.addWidget(self.table)

        # Formulário
        form_container = QtWidgets.QWidget()
        form_layout = QtWidgets.QVBoxLayout(form_container)

        self.form_enabled = QtWidgets.QCheckBox("Ativo")
        self.form_name = QtWidgets.QLineEdit()
        self.form_tool = QtWidgets.QLineEdit()
        self.form_path = QtWidgets.QLineEdit()

        self.form_name.setPlaceholderText("Ex.: Conciliação diária")
        self.form_tool.setPlaceholderText("Ex.: UiPath")
        self.form_path.setPlaceholderText("Caminho do executável/script/atalho")

        # Seletores para reduzir erros de digitação (multi-seleção)
        current_year = int(datetime.now().strftime("%Y"))
        year_options = [str(y) for y in range(current_year - 1, current_year + 6)]
        self.form_year = MultiSelectComboBox(year_options, all_label="Todos")
        self.form_month = MultiSelectComboBox(_month_name_options(), all_label="Todos")
        self.form_week = MultiSelectComboBox(["1", "2", "3", "4", "5"], all_label="Todos")
        self.form_weekday = MultiSelectComboBox(_weekday_name_options(), all_label="Todos")

        def _pad2(token: str) -> str:
            t = str(token or "").strip()
            if not t:
                return t
            if t.lower() == "todos":
                return "Todos"
            if t.isdigit():
                return t.zfill(2)
            return t

        day_options = [f"{d:02d}" for d in range(1, 32)]
        hour_options = [f"{h:02d}" for h in range(0, 24)]
        minute_options = [f"{m:02d}" for m in range(0, 60)]

        self.form_day = MultiSelectComboBox(day_options, all_label="Todos", normalize_token=_pad2)
        self.form_hour = MultiSelectComboBox(hour_options, all_label="Todos", normalize_token=_pad2)
        self.form_minute = MultiSelectComboBox(minute_options, all_label="Todos", normalize_token=_pad2)

        self._current_id: int | None = None

        form = QtWidgets.QFormLayout()
        form.addRow(self.form_enabled)
        form.addRow("Nome do processo", self.form_name)
        form.addRow("Ferramenta", self.form_tool)
        form.addRow("Caminho do executável", self.form_path)
        form.addRow("Ano", self.form_year)
        form.addRow("Meses do ano", self.form_month)
        form.addRow("Semanas do mês", self.form_week)
        form.addRow("Dias da semana", self.form_weekday)
        form.addRow("Dia", self.form_day)
        form.addRow("Hora", self.form_hour)
        form.addRow("Minuto", self.form_minute)
        form_layout.addLayout(form)

        buttons = QtWidgets.QHBoxLayout()
        self.btn_new = QtWidgets.QPushButton("Novo")
        self.btn_save = QtWidgets.QPushButton("Salvar")
        self.btn_delete = QtWidgets.QPushButton("Excluir")

        # Sem ícones (conforme solicitado)
        self.btn_save.setProperty("variant", "primary")
        self.btn_delete.setProperty("variant", "danger")
        self.btn_save.setObjectName("PrimaryBtn")
        self.btn_delete.setObjectName("DangerBtn")
        buttons.addWidget(self.btn_new)
        buttons.addWidget(self.btn_save)
        buttons.addWidget(self.btn_delete)
        buttons.addStretch(1)
        form_layout.addLayout(buttons)

        helper = QtWidgets.QLabel(
            "Agendamento aceita 'Todos' ou listas (ex.: 07,08).\n"
            "Mês e dia da semana dependem do idioma do Windows (ex.: December/Friday)."
        )
        helper.setWordWrap(True)
        form_layout.addWidget(helper)

        splitter_h.addWidget(form_container)
        splitter_h.setStretchFactor(0, 2)
        splitter_h.setStretchFactor(1, 3)

        # Wiring
        self.btn_refresh.clicked.connect(self._reload_processes)
        self.btn_new.clicked.connect(self._new_form)
        self.btn_save.clicked.connect(self._save_form)
        self.btn_delete.clicked.connect(self._delete_selected)
        self.btn_run.clicked.connect(self._run_selected)
        self.table.itemSelectionChanged.connect(self._on_table_selection_changed)

        self._reload_processes()

    def _reload_processes(self) -> None:
        keep_id = self._current_id
        procs = self.controller.list_processes()
        self.table.setRowCount(0)
        for proc in procs:
            row = self.table.rowCount()
            self.table.insertRow(row)

            id_item = QtWidgets.QTableWidgetItem(str(proc.id or ""))
            enabled_item = QtWidgets.QTableWidgetItem("Sim" if proc.enabled else "Não")
            name_item = QtWidgets.QTableWidgetItem(proc.Nome_Processo)
            tool_item = QtWidgets.QTableWidgetItem(proc.Ferramenta)
            path_item = QtWidgets.QTableWidgetItem(proc.Caminho)

            self.table.setItem(row, 0, id_item)
            self.table.setItem(row, 1, enabled_item)
            self.table.setItem(row, 2, name_item)
            self.table.setItem(row, 3, tool_item)
            self.table.setItem(row, 4, path_item)

        self.table.resizeColumnsToContents()

        if keep_id is not None:
            self._select_row_by_id(keep_id)

    def _select_row_by_id(self, process_id: int) -> None:
        for row in range(self.table.rowCount()):
            id_item = self.table.item(row, 0)
            if id_item is None or not id_item.text().strip():
                continue
            try:
                if int(id_item.text()) == int(process_id):
                    self.table.selectRow(row)
                    self.table.scrollToItem(self.table.item(row, 2))
                    return
            except Exception:
                continue

    def _on_table_selection_changed(self) -> None:
        selected = self.table.selectedItems()
        if not selected:
            return

        row = selected[0].row()
        id_item = self.table.item(row, 0)
        if id_item is None or not id_item.text().strip():
            return

        process_id = int(id_item.text())
        proc = self.controller.db.get_process(process_id)
        if proc is None:
            return

        self._current_id = proc.id
        self.form_enabled.setChecked(proc.enabled)
        self.form_name.setText(proc.Nome_Processo)
        self.form_tool.setText(proc.Ferramenta)
        self.form_path.setText(proc.Caminho)
        self.form_year.set_value_text(proc.ano)
        self.form_month.set_value_text(proc.meses_do_ano)
        self.form_week.set_value_text(proc.semanas_do_mes)
        self.form_weekday.set_value_text(proc.dias_da_semana)
        self.form_day.set_value_text(proc.dia)
        self.form_hour.set_value_text(proc.hora)
        self.form_minute.set_value_text(proc.minuto)

    def _new_form(self) -> None:
        self._current_id = None
        self.form_enabled.setChecked(True)
        self.form_name.setText("")
        self.form_tool.setText("")
        self.form_path.setText("")
        self.form_year.set_value_text("Todos")
        self.form_month.set_value_text("Todos")
        self.form_week.set_value_text("Todos")
        self.form_weekday.set_value_text("Todos")
        self.form_day.set_value_text("Todos")
        self.form_hour.set_value_text("Todos")
        self.form_minute.set_value_text("Todos")

    def _save_form(self) -> None:
        name = self.form_name.text().strip()
        tool = self.form_tool.text().strip()
        path = self.form_path.text().strip()

        if not name or not tool or not path:
            QtWidgets.QMessageBox.warning(
                self,
                "Campos obrigatórios",
                "Nome, Ferramenta e Caminho são obrigatórios.",
            )
            return

        proc = ProcessConfig(
            id=self._current_id,
            Nome_Processo=name,
            Ferramenta=tool,
            Caminho=path,
            ano=self.form_year.value_text().strip() or "Todos",
            meses_do_ano=self.form_month.value_text().strip() or "Todos",
            semanas_do_mes=self.form_week.value_text().strip() or "Todos",
            dias_da_semana=self.form_weekday.value_text().strip() or "Todos",
            dia=self.form_day.value_text().strip() or "Todos",
            hora=self.form_hour.value_text().strip() or "Todos",
            minuto=self.form_minute.value_text().strip() or "Todos",
            enabled=self.form_enabled.isChecked(),
        )

        try:
            new_id = self.controller.save_process(proc)
            self._current_id = new_id
            self._reload_processes()
            # Reaplica no form o que foi persistido (evita sensação de não atualizar)
            self._select_row_by_id(new_id)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Erro", f"Falha ao salvar: {exc}")

    def _delete_selected(self) -> None:
        selected = self.table.selectedItems()
        if not selected:
            return
        row = selected[0].row()
        id_item = self.table.item(row, 0)
        if id_item is None or not id_item.text().strip():
            return
        process_id = int(id_item.text())

        ok = QtWidgets.QMessageBox.question(self, "Confirmar", "Excluir o processo selecionado?")
        if ok != QtWidgets.QMessageBox.StandardButton.Yes:
            return

        try:
            self.controller.delete_process(process_id)
            self._new_form()
            self._reload_processes()
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Erro", f"Falha ao excluir: {exc}")

    def _run_selected(self) -> None:
        selected = self.table.selectedItems()
        if not selected:
            return

        row = selected[0].row()
        id_item = self.table.item(row, 0)
        if id_item is None or not id_item.text().strip():
            return

        process_id = int(id_item.text())
        proc = self.controller.db.get_process(process_id)
        if proc is None:
            return

        item: dict[str, Any] = {
            "processo": proc.Nome_Processo,
            "ferramenta": proc.Ferramenta,
            "caminho": proc.Caminho,
        }
        self.controller.enqueue_manual(item)


class MainWindow(QtWidgets.QMainWindow):
    """Janela principal (dashboard) do Orquestrador."""

    def __init__(self, db_path: str | None = None) -> None:
        """Inicializa a GUI principal e conecta sinais do controller."""
        super().__init__()
        self.setWindowTitle("Orquestrador RPA")
        self.resize(1200, 760)

        self.controller = OrchestratorController(db_path)
        self.controller.console_text.connect(self._append_console)
        self.controller.log_text.connect(self._append_log)
        self.controller.status_text.connect(self._set_status)
        self.controller.scheduler_state_changed.connect(self._on_scheduler_state_changed)
        self.controller.process_running_changed.connect(self._on_process_running_changed)
        self.controller.queue_changed.connect(self._on_queue_changed)
        self.controller.running_item_changed.connect(self._on_running_item_changed)
        self.controller.processes_changed.connect(self._reload_today_schedule)

        self._process_manager: ProcessManagerWindow | None = None

        central = QtWidgets.QWidget(self)
        self.setCentralWidget(central)
        root = QtWidgets.QVBoxLayout(central)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(12)

        # App bar (visual moderno): título + status + ações
        appbar = QtWidgets.QFrame()
        appbar.setObjectName("AppBar")
        appbar_layout = QtWidgets.QVBoxLayout(appbar)
        appbar_layout.setContentsMargins(12, 12, 12, 12)
        appbar_layout.setSpacing(10)

        title_row = QtWidgets.QHBoxLayout()
        title_row.setSpacing(10)

        title_col = QtWidgets.QVBoxLayout()
        title_col.setSpacing(2)
        title_lbl = QtWidgets.QLabel("Orquestrador RPA")
        title_lbl.setObjectName("AppTitle")
        subtitle_lbl = QtWidgets.QLabel("Dashboard")
        subtitle_lbl.setObjectName("AppSubtitle")
        title_col.addWidget(title_lbl)
        title_col.addWidget(subtitle_lbl)
        title_row.addLayout(title_col)

        title_row.addStretch(1)

        self.status_dot = StatusDot()
        self.status_label = QtWidgets.QLabel("Pronto")
        self.status_label.setObjectName("StatusText")
        status_wrap = QtWidgets.QHBoxLayout()
        status_wrap.setSpacing(8)
        status_wrap.addWidget(self.status_dot)
        status_wrap.addWidget(self.status_label)
        status_widget = QtWidgets.QWidget()
        status_widget.setLayout(status_wrap)
        title_row.addWidget(status_widget)

        appbar_layout.addLayout(title_row)

        actions = QtWidgets.QHBoxLayout()
        actions.setSpacing(10)

        self.btn_processes = QtWidgets.QPushButton("Gerenciar processos")
        self.btn_start = QtWidgets.QPushButton("Iniciar")
        self.btn_stop = QtWidgets.QPushButton("Parar")
        self.btn_cancel = QtWidgets.QPushButton("Cancelar execução")
        self.btn_refresh = QtWidgets.QPushButton("Recarregar")

        # Sem ícones (conforme solicitado)
        for b in (self.btn_processes, self.btn_start, self.btn_stop, self.btn_cancel, self.btn_refresh):
            b.setIcon(QtGui.QIcon())
            b.setMinimumWidth(170)

        self.btn_start.setProperty("variant", "primary")
        self.btn_stop.setProperty("variant", "danger")
        self.btn_cancel.setProperty("variant", "danger")

        # Compatibilidade visual com example/ (selectors #PrimaryBtn/#DangerBtn)
        self.btn_start.setObjectName("PrimaryBtn")
        self.btn_stop.setObjectName("DangerBtn")
        self.btn_cancel.setObjectName("DangerBtn")

        self.btn_start.setToolTip("Inicia o scheduler")
        self.btn_stop.setToolTip("Para o scheduler")
        self.btn_cancel.setToolTip("Interrompe a execução atual")
        self.btn_refresh.setToolTip("Atualiza agenda, fila e logs")
        self.btn_processes.setToolTip("Abre o cadastro de processos")

        # Atalhos (pequenos ganhos de usabilidade)
        self.btn_refresh.setShortcut(QtGui.QKeySequence.Refresh)
        self.btn_processes.setShortcut(QtGui.QKeySequence("Ctrl+P"))

        for b in (self.btn_processes, self.btn_start, self.btn_stop, self.btn_cancel, self.btn_refresh):
            actions.addWidget(b)

        actions.addStretch(1)
        appbar_layout.addLayout(actions)
        root.addWidget(appbar)

        # Painel principal: esquerda (agenda do dia), direita (fila + executando)
        splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        root.addWidget(splitter, 1)

        left_group = QtWidgets.QGroupBox("Agenda de hoje")
        left_layout = QtWidgets.QVBoxLayout(left_group)
        left_layout.setContentsMargins(12, 12, 12, 12)
        left_layout.setSpacing(8)
        left_layout.addWidget(_make_section_label("Automações previstas (07:00–18:00)"))
        self.today_table = QtWidgets.QTableWidget(0, 3)
        self.today_table.setHorizontalHeaderLabels(["Horário", "Processo", "Ferramenta"])
        self.today_table.setSelectionBehavior(QtWidgets.QTableWidget.SelectionBehavior.SelectRows)
        self.today_table.setEditTriggers(QtWidgets.QTableWidget.EditTrigger.NoEditTriggers)
        self.today_table.verticalHeader().setVisible(False)
        self.today_table.setAlternatingRowColors(True)
        self.today_table.horizontalHeader().setStretchLastSection(True)
        left_layout.addWidget(self.today_table, 1)
        splitter.addWidget(left_group)

        right_group = QtWidgets.QGroupBox("Execução")
        right_layout = QtWidgets.QVBoxLayout(right_group)
        right_layout.setContentsMargins(12, 12, 12, 12)
        right_layout.setSpacing(8)

        right_layout.addWidget(_make_section_label("Executando agora"))
        self.running_label = QtWidgets.QLabel("Nenhum")
        self.running_label.setWordWrap(True)
        right_layout.addWidget(self.running_label)

        right_layout.addWidget(_make_section_label("Fila"))
        self.queue_list = QtWidgets.QListWidget()
        right_layout.addWidget(self.queue_list, 1)
        splitter.addWidget(right_group)

        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)

        # Console + Logs
        tabs = QtWidgets.QTabWidget()
        root.addWidget(tabs, 1)

        self.console = QtWidgets.QTextEdit()
        self.console.setReadOnly(True)
        # Console escuro (como em example/gui_windows.py)
        self.console.setStyleSheet(
            "background-color: #0F172A; color: #F8FAFC; font-family: 'Consolas', monospace;"
        )
        tabs.addTab(self.console, "Console")

        self.logs = QtWidgets.QTextEdit()
        self.logs.setReadOnly(True)

        logs_tab = QtWidgets.QWidget()
        logs_layout = QtWidgets.QVBoxLayout(logs_tab)
        logs_layout.setContentsMargins(0, 0, 0, 0)
        logs_layout.setSpacing(10)

        export_bar = QtWidgets.QHBoxLayout()
        export_bar.setSpacing(10)

        export_bar.addWidget(QtWidgets.QLabel("De:"))
        self.logs_from = QtWidgets.QDateTimeEdit()
        self.logs_from.setCalendarPopup(True)
        self.logs_from.setDisplayFormat("dd/MM/yyyy HH:mm")
        self.logs_from.setTimeSpec(QtCore.Qt.TimeSpec.LocalTime)
        export_bar.addWidget(self.logs_from)

        export_bar.addWidget(QtWidgets.QLabel("Até:"))
        self.logs_to = QtWidgets.QDateTimeEdit()
        self.logs_to.setCalendarPopup(True)
        self.logs_to.setDisplayFormat("dd/MM/yyyy HH:mm")
        self.logs_to.setTimeSpec(QtCore.Qt.TimeSpec.LocalTime)
        export_bar.addWidget(self.logs_to)

        self.btn_export_logs = QtWidgets.QPushButton("Exportar logs")
        # Sem ícone (conforme solicitado)
        self.btn_export_logs.setIcon(QtGui.QIcon())
        self.btn_export_logs.setMinimumWidth(170)
        self.btn_export_logs.setShortcut(QtGui.QKeySequence("Ctrl+E"))
        export_bar.addWidget(self.btn_export_logs)
        export_bar.addStretch(1)

        export_widget = QtWidgets.QWidget()
        export_widget.setLayout(export_bar)
        logs_layout.addWidget(export_widget)
        logs_layout.addWidget(self.logs, 1)
        tabs.addTab(logs_tab, "Logs")

        # Wiring
        self.btn_processes.clicked.connect(self._open_process_manager)
        self.btn_start.clicked.connect(self.controller.start_scheduler)
        self.btn_stop.clicked.connect(self.controller.stop_scheduler)
        self.btn_cancel.clicked.connect(self.controller.stop_current_process)
        self.btn_refresh.clicked.connect(self._refresh_dashboard)
        self.btn_export_logs.clicked.connect(self._export_logs)

        # Estado inicial
        fixed_font = QtGui.QFont("Consolas", 10)
        fixed_font.setStyleHint(QtGui.QFont.StyleHint.Monospace)
        self.console.setFont(fixed_font)
        self.logs.setFont(fixed_font)

        now = QtCore.QDateTime.currentDateTime()
        start_of_day = QtCore.QDateTime(now.date(), QtCore.QTime(0, 0))
        self.logs_from.setDateTime(start_of_day)
        self.logs_to.setDateTime(now)

        self._scheduler_running = self.controller.is_scheduler_running()
        self._process_running = self.controller.is_process_running()
        self._update_action_buttons()
        self._update_status_dot()

        self._refresh_dashboard()

    def _open_process_manager(self) -> None:
        """Abre (ou traz para frente) a janela de gerenciamento de processos."""
        if self._process_manager is None:
            self._process_manager = ProcessManagerWindow(self.controller)
        self._process_manager.show()
        self._process_manager.raise_()
        self._process_manager.activateWindow()

    def _refresh_dashboard(self) -> None:
        """Recarrega agenda, logs e estados atuais (fila e executando)."""
        self._reload_today_schedule()
        self._reload_logs()
        self._on_queue_changed(self.controller.get_queue_snapshot())
        self._on_running_item_changed(self.controller.get_running_item())

    def _reload_today_schedule(self) -> None:
        """Carrega a agenda do dia e popula a tabela."""
        items = self.controller.list_today_schedule()
        self.today_table.setRowCount(0)
        if not items:
            self.today_table.setRowCount(1)
            empty = QtWidgets.QTableWidgetItem("Nenhuma automação prevista para hoje")
            empty.setFlags(QtCore.Qt.ItemFlag.NoItemFlags)
            empty.setForeground(QtGui.QBrush(QtGui.QColor("#6B7280")))
            self.today_table.setItem(0, 0, empty)
            self.today_table.setSpan(0, 0, 1, 3)
            self.today_table.resizeColumnsToContents()
            return
        for it in items:
            row = self.today_table.rowCount()
            self.today_table.insertRow(row)
            self.today_table.setItem(row, 0, QtWidgets.QTableWidgetItem(it["hora"]))
            self.today_table.setItem(row, 1, QtWidgets.QTableWidgetItem(it["processo"]))
            self.today_table.setItem(row, 2, QtWidgets.QTableWidgetItem(it["ferramenta"]))
        self.today_table.resizeColumnsToContents()

    def _on_queue_changed(self, items: object) -> None:
        """Atualiza a lista de fila a partir do snapshot emitido pelo controller."""
        self.queue_list.clear()
        if not isinstance(items, list):
            return
        if not items:
            empty = QtWidgets.QListWidgetItem("Fila vazia")
            empty.setFlags(QtCore.Qt.ItemFlag.NoItemFlags)
            empty.setForeground(QtGui.QBrush(QtGui.QColor("#6B7280")))
            self.queue_list.addItem(empty)
            return
        for it in items:
            if not isinstance(it, dict):
                continue
            txt = f"{it.get('processo','')} ({it.get('ferramenta','')})"
            self.queue_list.addItem(txt)

    def _on_running_item_changed(self, item: object) -> None:
        """Atualiza o texto de "executando agora" e botões de ação."""
        if not isinstance(item, dict) or not item:
            self.running_label.setText("Nenhum")
            self._update_action_buttons()
            return
        self.running_label.setText(f"{item.get('processo','')} ({item.get('ferramenta','')})")
        self._update_action_buttons()

    # ---------------- UI helpers ----------------
    def _set_status(self, text: str) -> None:
        """Atualiza o texto de status na barra superior."""
        self.status_label.setText(text)

    def _update_action_buttons(self) -> None:
        """Habilita/desabilita botões conforme estado do scheduler/execução."""
        # Iniciar: apenas quando scheduler NÃO está rodando
        self.btn_start.setEnabled(not self._scheduler_running)
        # Parar: apenas quando scheduler está rodando
        self.btn_stop.setEnabled(self._scheduler_running)

        # Cancelar execução: apenas quando há algo executando E for Python
        can_cancel = False
        try:
            can_cancel = bool(self._process_running and self.controller.can_cancel_current_process())
        except Exception:
            can_cancel = False
        self.btn_cancel.setEnabled(can_cancel)

    def _on_scheduler_state_changed(self, running: bool) -> None:
        """Slot para mudança de estado do scheduler (liga/desliga)."""
        self._scheduler_running = bool(running)
        self._update_action_buttons()
        self._update_status_dot()

    def _on_process_running_changed(self, running: bool) -> None:
        """Slot para mudança de estado de execução (idle/busy)."""
        self._process_running = bool(running)
        self._update_action_buttons()
        self._update_status_dot()

    def _update_status_dot(self) -> None:
        """Atualiza o StatusDot conforme estados atuais."""
        if not self._scheduler_running:
            self.status_dot.set_state("stopped")
            return

        if self._process_running:
            self.status_dot.set_state("busy")
        else:
            self.status_dot.set_state("idle")

    def _append_console(self, text: str) -> None:
        """Acrescenta texto no console da UI."""
        self.console.moveCursor(QtGui.QTextCursor.MoveOperation.End)
        self.console.insertPlainText(text)
        self.console.moveCursor(QtGui.QTextCursor.MoveOperation.End)

    def _append_log(self, text: str) -> None:
        """Acrescenta uma linha no painel de logs da UI."""
        self.logs.moveCursor(QtGui.QTextCursor.MoveOperation.End)
        self.logs.insertPlainText(text + "\n")
        self.logs.moveCursor(QtGui.QTextCursor.MoveOperation.End)

    def _reload_logs(self) -> None:
        """Recarrega o painel de logs (texto) a partir do banco."""
        self.logs.clear()
        for line in self.controller.list_logs_text(limit=500):
            self.logs.append(line)

    def _export_logs(self) -> None:
        """Exporta logs do período selecionado para CSV ou TXT."""
        start_dt = self.logs_from.dateTime()
        end_dt = self.logs_to.dateTime()

        if end_dt < start_dt:
            QtWidgets.QMessageBox.warning(self, "Período inválido", "A data 'Até' deve ser maior ou igual à data 'De'.")
            return

        start_utc = datetime.fromtimestamp(start_dt.toSecsSinceEpoch(), tz=timezone.utc).isoformat()
        end_utc = datetime.fromtimestamp(end_dt.toSecsSinceEpoch(), tz=timezone.utc).isoformat()

        entries = self.controller.list_logs_entries_between(start_utc, end_utc, limit=None)

        path, selected_filter = QtWidgets.QFileDialog.getSaveFileName(
            self,
            "Exportar logs",
            "logs.csv",
            "CSV (*.csv);;Texto (*.txt)",
        )
        if not path:
            return

        use_csv = path.lower().endswith(".csv") or "CSV" in selected_filter
        if use_csv and not path.lower().endswith(".csv"):
            path += ".csv"
        if (not use_csv) and (not path.lower().endswith(".txt")):
            path += ".txt"

        try:
            if use_csv:
                import csv

                with open(path, "w", encoding="utf-8", newline="") as f:
                    w = csv.writer(f)
                    w.writerow(["ts_iso", "stream", "process_id", "message"])
                    for e in entries:
                        w.writerow([e.ts_iso, e.stream, e.process_id if e.process_id is not None else "", e.message])
            else:
                with open(path, "w", encoding="utf-8") as f:
                    for e in entries:
                        f.write(f"{e.ts_iso} [{e.stream}] {e.message}\n")
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Erro", f"Falha ao exportar logs: {exc}")
            return

        QtWidgets.QMessageBox.information(self, "Exportação concluída", f"{len(entries)} linhas exportadas.\nArquivo: {path}")



def main() -> None:
    """Ponto de entrada da aplicação GUI."""
    _install_qt_message_filter()
    app = QtWidgets.QApplication(sys.argv)
    _apply_app_style(app)

    db_path = os.getenv("ORCH_DB_PATH")
    win = MainWindow(db_path)
    win.show()

    # Auto-iniciar scheduler ao abrir (como solicitado)
    QtCore.QTimer.singleShot(0, win.controller.start_scheduler)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
