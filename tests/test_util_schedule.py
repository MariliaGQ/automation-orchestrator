"""Testes do módulo util (agendamento).

Cobertura principal:
    - Comparação 'Todos'
    - Campos numéricos com e sem zero à esquerda
    - Campos multi-valor com separadores
    - Normalização de item de processo
"""

import pytest

from util import NowParts, should_enqueue, to_process_item


def make_row(**overrides):
    """Cria uma linha base de planilha com overrides para testes."""
    base = {
        "Nome_Processo": "MeuProcesso",
        "Ferramenta": "Uipath",
        "Caminho": r"C:\UiPath\robot.exe",
        "ano": "Todos",
        "meses_do_ano": "Todos",
        "semanas_do_mes": "Todos",
        "dias_da_semana": "Todos",
        "dia": "Todos",
        "hora": "Todos",
        "minuto": "Todos",
    }
    base.update(overrides)
    return base


def test_should_enqueue_all_todos_matches():
    """Quando tudo é 'Todos', sempre deve enfileirar."""
    now = NowParts(
        year="2025",
        month_name="December",
        week_of_month="4",
        weekday_name="Friday",
        day="26",
        hour="07",
        minute="30",
    )
    assert should_enqueue(make_row(), now) is True


def test_should_enqueue_specific_hour_and_minute_matches():
    """Hora/minuto específicos devem casar quando iguais."""
    now = NowParts(
        year="2025",
        month_name="December",
        week_of_month="4",
        weekday_name="Friday",
        day="26",
        hour="07",
        minute="05",
    )

    row = make_row(hora="07", minuto="05")
    assert should_enqueue(row, now) is True


def test_should_enqueue_portuguese_month_and_weekday_matches():
    """Mês e dia da semana em PT devem casar quando iguais."""
    now = NowParts(
        year="2025",
        month_name="Dezembro",
        week_of_month="4",
        weekday_name="sexta-feira",
        day="26",
        hour="07",
        minute="30",
    )

    row = make_row(meses_do_ano="Dezembro", dias_da_semana="sexta-feira")
    assert should_enqueue(row, now) is True


def test_should_enqueue_numeric_fields_accept_unpadded_values():
    """Campos numéricos devem aceitar valores sem padding (ex.: '7' == '07')."""
    now = NowParts(
        year="2025",
        month_name="December",
        week_of_month="4",
        weekday_name="Friday",
        day="07",
        hour="07",
        minute="05",
    )

    # Valores sem zero à esquerda devem casar com os valores atuais (strftime usa '07', '05', etc.)
    row = make_row(dia="7", hora="7", minuto="5")
    assert should_enqueue(row, now) is True


def test_should_enqueue_mixed_padded_and_unpadded_list_matches():
    """Lista mista de tokens deve casar por equivalência numérica."""
    now = NowParts(
        year="2025",
        month_name="December",
        week_of_month="4",
        weekday_name="Friday",
        day="26",
        hour="07",
        minute="00",
    )

    # Lista mista deve funcionar: 7 == 07 e 08 continua válido.
    row = make_row(hora="7,08", minuto="00")
    assert should_enqueue(row, now) is True


def test_should_enqueue_specific_hour_and_minute_does_not_match():
    """Hora/minuto diferentes não devem casar."""
    now = NowParts(
        year="2025",
        month_name="December",
        week_of_month="4",
        weekday_name="Friday",
        day="26",
        hour="07",
        minute="06",
    )

    row = make_row(hora="07", minuto="05")
    assert should_enqueue(row, now) is False


@pytest.mark.parametrize(
    ("field_value", "now_value"),
    [
        ("07,08", "07"),
        ("07;08", "08"),
        ("07|08", "07"),
        ("07, 08", "08"),
    ],
)
def test_multi_value_fields_match_via_separators(field_value, now_value):
    """Separadores ',', ';' e '|' devem funcionar para multi-valores."""
    now = NowParts(
        year="2025",
        month_name="December",
        week_of_month="4",
        weekday_name="Friday",
        day="26",
        hour=now_value,
        minute="00",
    )

    row = make_row(hora=field_value, minuto="00")
    assert should_enqueue(row, now) is True


def test_to_process_item_normalizes_numeric_and_strips():
    """`to_process_item` deve normalizar e remover espaços."""
    row = make_row(Nome_Processo="  X ", Ferramenta="  Uipath ", Caminho="  c ")
    item = to_process_item(row)
    assert item == {"processo": "X", "ferramenta": "Uipath", "caminho": "c"}
