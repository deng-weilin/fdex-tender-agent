#!/usr/bin/env python3
"""M12 精确十进制报价计算（绿地模块）。

任务 M12-01：按**给定规则**计算报价、检查税费口径与限价。

设计要点（与验收器契约对齐）：
- 全程 Decimal，金额从字符串构造，绝不碰 float。
- 舍入先分项（net）再税费/总价；含税（inclusive）用"差额法"：
  gross = 数量×单价（含税总价），net = gross/(1+税率)，tax = gross − net，
  保证 net + tax == gross 恒等式，审计可对账。
- 缺税率或含税口径：保留未知（tax/gross=null），标记 TAX_RULE_MISSING，绝不默认零税。
- 超限（gross > ceiling）→ CEILING_EXCEEDED；声明总价（declared_total）与计算总价不一致
  → TOTAL_MISMATCH。两类都属硬阻断 → status=BLOCKED。
- 单价变化使旧批准失效：比对 previous_lines 与当前 lines 的内容，不一致 → approval_status=STALE。

门禁（fail-closed，坏数据/不支持参数进不来）：
- 数值"缺失"与"非法"严格区分：缺失（缺字段/空串）→ LINE_INCOMPLETE / TAX_RULE_MISSING；
  非法（不可解析、NaN/Infinity、负数越界）→ 对应 *_INVALID 码，一律不计算、不透出。
- 舍入规则只支持任务要求的 ROUND_HALF_UP；其它任何声明 → ROUNDING_RULE_UNSUPPORTED，
  **绝不静默降级**；digits 必须是非负整数，否则 ROUNDING_DIGITS_INVALID。
- 非法 ceiling / declared_total 绝不当作"未提供"而跳过校验 → CEILING_INVALID / DECLARED_TOTAL_INVALID。
- 批准必须"四证齐备"才 CURRENT：status==APPROVED、摘要绑定 previous_lines[0]、
  scope 显式覆盖全部当前报价行、当前 lines 与 previous_lines 一致；否则 INVALID/STALE。

本模块只做计算与状态判定，不读 expected.json、不自行定价、不真实审批、不生成 Word。
对外输出字典符合 bid-training.acceptance-result.v1（不含 input_hashes/evidence，那些由适配器补）。
"""
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
import hashlib
import json

RESULT_VERSION = "bid-training.acceptance-result.v1"

# 本任务仅要求 ROUND_HALF_UP；其它任何值都必须明确报"不支持"，绝不静默降级改规则。
_ROUNDING = {"ROUND_HALF_UP": ROUND_HALF_UP}

# 阻断类问题（任一出现 → status=BLOCKED）：坏数据/不支持参数/硬约束违反。
_BLOCKING = {
    "LINE_INCOMPLETE",
    "CEILING_EXCEEDED",
    "TOTAL_MISMATCH",
    "QUANTITY_INVALID",
    "UNIT_PRICE_INVALID",
    "TAX_RATE_INVALID",
    "CEILING_INVALID",
    "DECLARED_TOTAL_INVALID",
    "ROUNDING_RULE_UNSUPPORTED",
    "ROUNDING_DIGITS_INVALID",
}
# 待核对类问题（无阻断项时出现 → status=REVIEW）：规则缺口/批准状态需人判。
_REVIEW = {"TAX_RULE_MISSING", "APPROVAL_STALE", "APPROVAL_INVALID"}

_VALID_MODES = ("exclusive", "inclusive")


def _parse_amount(raw):
    """金额/数量/税率 → (Decimal 或 None, state)。

    state:
      "missing" —— 未提供（None / 空白串），语义上"没给"；
      "invalid" —— 提供了但非法（不可解析、NaN、Infinity、bool 之类），"给错了"；
      "ok"      —— 合法有限值。

    关键：绝不把"非法"降级成"缺失"，否则非法 ceiling/declared_total 会静默关闭校验。
    """
    if raw is None:
        return None, "missing"
    if isinstance(raw, bool):
        return None, "invalid"
    if isinstance(raw, str) and raw.strip() == "":
        return None, "missing"
    try:
        value = Decimal(str(raw).strip())
    except (InvalidOperation, ValueError, TypeError, ArithmeticError):
        return None, "invalid"
    if not value.is_finite():
        return None, "invalid"
    return value, "ok"


def _parse_digits(raw):
    """digits → (int 或 None, state)。默认 2；必须是非负整数。"""
    if raw is None:
        return 2, "ok"
    if isinstance(raw, bool):
        return None, "invalid"
    if isinstance(raw, int):
        value = raw
    elif isinstance(raw, str):
        text = raw.strip()
        if text.lstrip("+-").isdigit() and text not in ("+", "-"):
            value = int(text)
        else:
            return None, "invalid"
    else:
        return None, "invalid"
    if value < 0:
        return None, "invalid"
    return value, "ok"


def _round(value, digits, rounding):
    return value.quantize(Decimal(1).scaleb(-digits), rounding=rounding)


def _quote(line):
    """计算一条报价行。返回 (quote_dict, [issue_code, ...])。"""
    qty, qty_state = _parse_amount(line.get("quantity"))
    price, price_state = _parse_amount(line.get("unit_price"))
    rate, rate_state = _parse_amount(line.get("tax_rate"))
    ceiling, ceiling_state = _parse_amount(line.get("ceiling"))
    mode = line.get("tax_mode")

    quote = {}
    invalid = []

    # 非法 / 非负范围（数量、金额、税率、限价都必须有限且非负）。
    if qty_state == "invalid" or (qty is not None and qty < 0):
        invalid.append("QUANTITY_INVALID")
    if price_state == "invalid" or (price is not None and price < 0):
        invalid.append("UNIT_PRICE_INVALID")
    if rate_state == "invalid" or (rate is not None and rate < 0):
        invalid.append("TAX_RATE_INVALID")
    if ceiling_state == "invalid" or (ceiling is not None and ceiling < 0):
        invalid.append("CEILING_INVALID")

    digits, digits_state = _parse_digits(line.get("digits", 2))
    if digits_state != "ok":
        invalid.append("ROUNDING_DIGITS_INVALID")
    rule_name = line.get("rounding", "ROUND_HALF_UP")
    rounding = _ROUNDING.get(rule_name) if isinstance(rule_name, str) else None
    if rounding is None:
        invalid.append("ROUNDING_RULE_UNSUPPORTED")

    if invalid:
        # 坏数据/不支持参数一律不计算、不透出数值。
        quote["net"] = None
        quote["tax"] = None
        quote["gross"] = None
        return quote, invalid

    # 缺失（非非法）：必需的数量/单价没给 → 行不完整。
    if qty_state == "missing" or price_state == "missing":
        quote["net"] = None
        quote["tax"] = None
        quote["gross"] = None
        return quote, ["LINE_INCOMPLETE"]

    if mode not in _VALID_MODES or rate_state == "missing":
        # 缺税率或未知含税口径：不能当零税，保留未知（net 仍可算）。
        quote["net"] = str(_round(qty * price, digits, rounding))
        quote["tax"] = None
        quote["gross"] = None
        return quote, ["TAX_RULE_MISSING"]

    if mode == "exclusive":
        net = _round(qty * price, digits, rounding)
        tax = _round(net * rate, digits, rounding)
        gross = _round(net + tax, digits, rounding)
    else:  # inclusive：含税总价 = 数量×单价；净价反算；税额取差额（保证 net+tax==gross）。
        gross = _round(qty * price, digits, rounding)
        net = _round(gross / (Decimal(1) + rate), digits, rounding)
        tax = _round(gross - net, digits, rounding)

    quote["net"], quote["tax"], quote["gross"] = str(net), str(tax), str(gross)

    # 超限检查：ceiling 合法且已提供才判（非法 ceiling 已在上面阻断）。
    if ceiling is not None:
        over = gross > ceiling
        quote["over_ceiling"] = bool(over)
        if over:
            return quote, ["CEILING_EXCEEDED"]
    return quote, []


def _lines_keyed(lines):
    return {x.get("id"): json.dumps(x, sort_keys=True, ensure_ascii=False) for x in lines}


def _scope_ids(scope):
    """批准范围 → 覆盖的报价行 id 集合；无法解析 → None（缺范围=不能确认覆盖）。"""
    if scope is None:
        return None
    if isinstance(scope, str):
        parts = [p.strip() for p in scope.replace(";", ",").split(",")]
        return {p for p in parts if p}
    if isinstance(scope, (list, tuple, set)):
        return {str(s).strip() for s in scope if str(s).strip()}
    return None


def _approval_status(payload):
    """返回 (approval_status_or_None, [issue_code, ...])。

    只有"四证齐备"才 CURRENT：明示 APPROVED + 摘要绑定 previous_lines[0] +
    scope 显式覆盖全部当前行 + 当前 lines 与 previous_lines 一致。
    任何一项不满足都不得标成 CURRENT（未批准/被驳回/范围不符/内容漂移）。
    """
    approval = payload.get("approval")
    previous = payload.get("previous_lines") or []
    lines = payload.get("lines", []) or []
    if not approval:
        return None, []

    # 1) 批准状态必须明示 APPROVED；REJECTED/PENDING/未知一律不得 CURRENT。
    if approval.get("status") != "APPROVED":
        return "INVALID", ["APPROVAL_INVALID"]

    # 2) 摘要必须真正绑定 previous_lines[0]（防伪造/过期引用）；无 previous 无法验证。
    if not previous:
        return "INVALID", ["APPROVAL_INVALID"]
    expected_digest = hashlib.sha256(
        json.dumps(previous[0], sort_keys=True, ensure_ascii=False).encode()
    ).hexdigest()
    if expected_digest != approval.get("input_digest"):
        return "INVALID", ["APPROVAL_INVALID"]

    # 3) 批准范围必须显式覆盖全部当前报价行；缺 scope 或未覆盖 → 该行无有效批准。
    scope_ids = _scope_ids(approval.get("scope"))
    current_ids = [line.get("id") for line in lines]
    if scope_ids is None or not set(current_ids).issubset(scope_ids):
        return "INVALID", ["APPROVAL_INVALID"]

    # 4) 批准内容与当前行一致才持续有效；不一致 → 旧批准失效。
    if _lines_keyed(lines) != _lines_keyed(previous):
        return "STALE", ["APPROVAL_STALE"]

    return "CURRENT", []


def compute_pricing(request):
    """request = input.json 解析后的 dict（含 payload）。返回结果字典。"""
    payload = request["payload"]
    lines = payload.get("lines", []) or []

    raw_issues = []
    quotes = {}
    computed_total = Decimal("0")
    for line in lines:
        quote, issues = _quote(line)
        quotes[line["id"]] = quote
        raw_issues += issues
        if quote.get("gross") is not None:
            computed_total += Decimal(quote["gross"])

    # 声明总价：缺失跳过；非法（不可解析/NaN/Inf/负数）必须拒绝，绝不静默跳过校验。
    declared_total, declared_state = _parse_amount(payload.get("declared_total"))
    if declared_state == "invalid" or (declared_total is not None and declared_total < 0):
        raw_issues.append("DECLARED_TOTAL_INVALID")
    elif declared_total is not None and computed_total != declared_total:
        raw_issues.append("TOTAL_MISMATCH")

    approval_status, approval_issues = _approval_status(payload)
    raw_issues += approval_issues

    # 去重并保持稳定顺序。
    seen, issue_codes = set(), []
    for code in raw_issues:
        if code not in seen:
            seen.add(code)
            issue_codes.append(code)

    if seen & _BLOCKING:
        status = "BLOCKED"
    elif seen & _REVIEW:
        status = "REVIEW"
    else:
        status = "COMPLETE"

    facts = {"quotes": quotes, "release_allowed": False}
    if approval_status is not None:
        facts["approval_status"] = approval_status

    return {
        "schema_version": RESULT_VERSION,
        "module": request.get("module", "M12"),
        "case_id": request.get("case_id"),
        "source_mode": request.get("source_mode"),
        "status": status,
        "facts": facts,
        "issue_codes": issue_codes,
        "real_submission_authorized": False,
    }
