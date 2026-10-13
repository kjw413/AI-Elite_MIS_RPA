"""UIA 메뉴 선택 검증: 실제 MIS 조작/수집 파일 쓰기 없이 실행한다."""
from types import SimpleNamespace
from unittest.mock import Mock, patch

import pytest
from pywinauto.uia_defines import NoPatternInterfaceError

from _common import open_mis_screen
import production_daily_rpa as production
import utility_daily_rpa as utility
import wip_daily_rpa as wip


def menu(name, auto_id="", *, control_type="TreeItem", parent=None):
    item = Mock()
    item.element_info = SimpleNamespace(
        name=name, automation_id=auto_id, control_type=control_type,
    )
    item.parent.return_value = parent
    item.is_visible.return_value = True
    item.is_enabled.return_value = True
    return item


def window_with(*items):
    window = Mock()
    window.descendants.side_effect = lambda **kwargs: [
        item for item in items
        if item.element_info.control_type == kwargs["control_type"]
    ]
    return window


def test_selects_exact_menu_despite_tab_and_similar_menu():
    target = menu("기간별 제품 생산실적(일/월)")
    similar = menu("기간별 제품 생산실적(일/월) 상세")
    tab = menu(target.element_info.name, control_type="TabItem")
    open_mis_screen(window_with(similar, tab, target), target.element_info.name)
    target.double_click_input.assert_called_once_with()
    similar.double_click_input.assert_not_called()
    tab.double_click_input.assert_not_called()


def test_normalizes_spacing_only():
    target = menu("생산계획 대비 실적현황 (완제품 / 재공품)")
    open_mis_screen(window_with(target), "생산계획 대비 실적현황(완제품/재공품)")
    target.double_click_input.assert_called_once()


def test_automation_id_disambiguates_duplicate_names():
    first, target = menu("동일 이름", "first"), menu("동일 이름", "target")
    open_mis_screen(window_with(first, target), {"title": "동일 이름", "auto_id": "target"})
    target.double_click_input.assert_called_once()
    first.double_click_input.assert_not_called()


def test_can_select_by_object_id_and_control_type_without_name():
    target = menu("이름 변경됨", "menu42", control_type="MenuItem")
    open_mis_screen(window_with(target), {"auto_id": "menu42", "control_type": "MenuItem"})
    target.double_click_input.assert_called_once()


def test_ambiguous_menu_stops_without_any_click():
    items = [menu("동일 이름"), menu("동일 이름")]
    with pytest.raises(RuntimeError, match="2개 객체"):
        open_mis_screen(window_with(*items), "동일 이름", timeout=0)
    for item in items:
        item.double_click_input.assert_not_called()


def test_missing_menu_reports_available_names_without_click():
    other = menu("다른 메뉴")
    with pytest.raises(RuntimeError, match="다른 메뉴"):
        open_mis_screen(window_with(other), "없는 메뉴", timeout=0)
    other.double_click_input.assert_not_called()


@pytest.mark.parametrize("property_name", ["is_visible", "is_enabled"])
def test_unavailable_menu_is_not_clicked(property_name):
    target = menu("대상")
    getattr(target, property_name).return_value = False
    with pytest.raises(RuntimeError, match="활성화되지"):
        open_mis_screen(window_with(target), "대상", timeout=0)
    target.double_click_input.assert_not_called()


def test_retries_while_menu_is_loading():
    target = menu("대상")
    window = Mock()
    window.descendants.side_effect = [[], [target]]
    with patch("_common.time.sleep") as sleep:
        open_mis_screen(window, "대상")
    sleep.assert_called_once_with(0.2)
    target.double_click_input.assert_called_once()


def test_expands_parents_in_order_before_scrolling_and_clicking():
    top = menu("상위")
    middle = menu("중간", parent=top)
    target = menu("대상", parent=middle)
    calls = Mock()
    calls.attach_mock(top.expand, "top_expand")
    calls.attach_mock(middle.expand, "middle_expand")
    calls.attach_mock(target.ensure_visible, "scroll")
    calls.attach_mock(target.double_click_input, "click")
    open_mis_screen(window_with(target), "대상")
    assert [call[0] for call in calls.mock_calls] == ["top_expand", "middle_expand", "scroll", "click"]


def test_visible_menu_works_without_scroll_pattern():
    target = menu("대상")
    target.ensure_visible.side_effect = NoPatternInterfaceError()
    open_mis_screen(window_with(target), "대상")
    target.double_click_input.assert_called_once()


@pytest.mark.parametrize("selector", [None, {}, "", {"title": 123}, {"auto_id": " "}, {"title": "대상", "control_type": None}])
def test_invalid_selector_fails_before_search(selector):
    window = Mock()
    with pytest.raises(ValueError):
        open_mis_screen(window, selector)
    window.descendants.assert_not_called()


@pytest.mark.parametrize("module,cls,method,args,expected", [
    (production, production.MISProductionRPA, "navigate_to_production_screen", (), "기간별 제품 생산실적(일/월)"),
    (utility, utility.MISUtilityRPA, "navigate_to_screen", (utility.SCREEN_UNIT_INPUT,), "원단위 실적입력(일단위)"),
    (wip, wip.MISWIPRPA, "navigate_to_wip_screen", (), "생산계획 대비 실적현황(완제품/재공품)"),
])
def test_each_rpa_uses_configured_menu_and_no_coordinate_click(module, cls, method, args, expected):
    runner = cls.__new__(cls)
    runner.coords = runner._load_coords()
    runner.main_window = Mock()
    with patch.object(module, "open_mis_screen") as select, patch.object(module, "fast_click") as click, patch.object(module.time, "sleep"):
        getattr(runner, method)(*args)
    select.assert_called_once_with(runner.main_window, expected)
    click.assert_not_called()


def test_utility_coordinate_validation_does_not_require_menu_coordinates():
    runner = utility.MISUtilityRPA.__new__(utility.MISUtilityRPA)
    runner.coords = runner._load_coords()
    assert "tree_menu" not in runner.coords[utility.SCREEN_UNIT_INPUT]
    runner._validate_coords()
