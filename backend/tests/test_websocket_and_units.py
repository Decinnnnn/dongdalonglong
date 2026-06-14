from decimal import Decimal

import pytest
from fastapi import WebSocketDisconnect

from app.models import Courier
from app.service_utils import STATUS_TEXT, build_order_out
from app.utils import (
    create_access_token, decode_access_token,
    hash_password, verify_password, parse_token_optional,
)


def test_password_hash_and_jwt_round_trip():
    hashed = hash_password("pass1234")
    assert hashed != "pass1234"
    assert verify_password("pass1234", hashed)
    assert not verify_password("wrong", hashed)

    token = create_access_token(7, "customer", "tester")
    payload = decode_access_token(token)
    assert payload["sub"] == "7"
    assert payload["role"] == "customer"
    assert payload["name"] == "tester"


def test_password_hash_handles_special_characters():
    """边界：特殊字符密码也能正常哈希与验证。"""
    for pw in ["", "a", "你好世界", "p@ss w0rd!😀", "a" * 128]:
        h = hash_password(pw)
        assert verify_password(pw, h)
        assert not verify_password(pw + "x", h)


def test_password_hash_is_deterministic_in_verification():
    """同一密码两次哈希结果不同（salt），但都能通过验证。"""
    h1 = hash_password("same")
    h2 = hash_password("same")
    assert h1 != h2
    assert verify_password("same", h1)
    assert verify_password("same", h2)


def test_jwt_payload_includes_role_and_name():
    token = create_access_token(1, "merchant", "商家A")
    payload = decode_access_token(token)
    assert payload["role"] == "merchant"
    assert payload["name"] == "商家A"
    assert "exp" in payload


# ── parse_token_optional 覆盖（此前完全未测）──────────────
def test_parse_token_optional_returns_tuple_for_valid_token():
    token = create_access_token(42, "courier", "骑手王")
    result = parse_token_optional(token)
    assert result == (42, "courier", "骑手王")


def test_parse_token_optional_returns_none_for_garbage():
    assert parse_token_optional("not.a.jwt") is None
    assert parse_token_optional("") is None


def test_parse_token_optional_returns_none_for_tampered_token():
    token = create_access_token(1, "customer", "x")
    # 篡改 payload 部分
    parts = token.split(".")
    tampered = parts[0] + "." + "Zm9vYmFy" + "." + parts[2]
    assert parse_token_optional(tampered) is None


# ── STATUS_TEXT 完整性 ─────────────────────────────────────
def test_all_order_statuses_have_labels():
    assert STATUS_TEXT == {
        0: "待处理",
        1: "已接单",
        2: "派送中",
        3: "已完成",
        4: "客户已取消",
        5: "商家已拒单",
    }


def test_status_text_covers_all_defined_states():
    """0-5 状态码均在 STATUS_TEXT 中有映射。"""
    for code in range(6):
        assert code in STATUS_TEXT
        assert isinstance(STATUS_TEXT[code], str)
        assert len(STATUS_TEXT[code]) > 0


# ── build_order_out 单元测试 ───────────────────────────────
def test_build_order_out_unknown_status_defaults_to_unknown():
    """未定义的状态码应回退为 '未知'。"""
    from unittest.mock import MagicMock

    order = MagicMock()
    order.id = 1
    order.customer_id = 1
    order.customer = MagicMock()
    order.customer.name = "test"
    order.merchant_id = 1
    order.merchant = MagicMock()
    order.merchant.name = "test"
    order.shipping_id = 1
    order.shipping = MagicMock()
    order.shipping.address = "addr"
    order.shipping.phone = "123"
    order.courier_id = None
    order.courier = None
    order.total = 0
    order.status = 99  # 不存在
    order.order_time = None
    order.order_dishes = []

    result = build_order_out(order)
    assert result["status_text"] == "未知"


@pytest.mark.websocket
def test_websocket_accepts_matching_identity(client, api_helpers):
    api_helpers["register"](client, "merchant", "ws-merchant")
    data = api_helpers["login"](client, "merchant", "ws-merchant")
    with client.websocket_connect(
        f"/api/merchant/ws/merchant/{data['user_id']}?token={data['access_token']}"
    ) as websocket:
        websocket.send_text("ping")


@pytest.mark.websocket
def test_websocket_rejects_missing_or_mismatched_token(client, api_helpers):
    with pytest.raises(WebSocketDisconnect) as missing:
        with client.websocket_connect("/api/merchant/ws/merchant/1"):
            pass
    assert missing.value.code == 4001

    api_helpers["register"](client, "customer", "ws-customer")
    data = api_helpers["login"](client, "customer", "ws-customer")
    with pytest.raises(WebSocketDisconnect) as mismatch:
        with client.websocket_connect(
            f"/api/merchant/ws/merchant/{data['user_id']}?token={data['access_token']}"
        ):
            pass
    assert mismatch.value.code == 4003
