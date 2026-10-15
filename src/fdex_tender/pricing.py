#!/usr/bin/env python3
"""M12 精确十进制报价计算（绿地模块）。

任务 M12-01：按给定规则计算报价、检查税费口径与限价。

设计要点（与验收器契约对齐）：
- 全程 Decimal，金额从字符串构造，绝不碰 float。
- 舍入先分项（net）再税费/总价；含税（inclusive）用"差额法"：
  gross = 数量×单价（含税总价），net = gross/(1+税率)，tax = gross − net，
  保证 net + tax == gross 恒等式，审计可对账。
- 缺税率或含税口径：保留未知（tax/gross=null），标记 TAX_RULE_MISSING，绝不默认零税。
- 超限（gross > ceiling）→ CEILING_EXCEEDED；声明总价（declared_total）与计算总价不一致
  → TOTAL_MISMATCH。两类都属硬阻断 → status=BLOCKED。
- 单价变化使旧批准失效：比对 previous_lines 与当前 lines 的内容，不一致 → approval_status=STALE
  且 issue_code=APPROVAL_STALE → status=REVIEW。

本模块只做计算与状态判定，不读 expected.json、不自行定价、不真实审批、不生成 Word。
对外输出字典符合 bid-training.acceptance-result.v1（不含 input_hashes/evidence，那些由适配器补）。
"""
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
import hashlib
import json

RESULT_VERSION = "bid-training.acceptance-result.v1"

# 本任务只要求 ROUND_HALF_UP；以后要扩模式再补映射。
_ROUNDING = {"ROUND_HALF_UP": ROUND_HALF_UP}

# 阻断类问题（任一出现 → status=BLOCKED）
_BLOCKING = {"CEILING_EXCEEDED", "TOTAL_MISMATCH", "LINE_INCOMPLETE"}
# 待核对类问题（无阻断项时出现 → status=REVIEW）
_REVIEW = {"TAX_RULE_MISSING", "APPROVAL_STALE", "APPROVAL_INVALID"}


def _dec(value, default=None):
    """输入字符串/数字 → Decimal；None/空 → default。"""
    if value is None or value == "":
        return default
    try:
        return Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError):
        return default


def _round(value, digits, rounding):
    return value.quantize(Decimal(1).scaleb(-digits), rounding=rounding)


def _quote(line, ceiling):
    """计算一条报价行。返回 (quote_dict, [issue_code, ...])。"""
    qty = _dec(line.get("quantity"))
    price = _dec(line.get("unit_price"))
    rate = _dec(line.get("tax_rate"))
    mode = line.get("tax_mode")
    digits = int(line.get("digits", 2))
    rounding = _ROUNDING.get(line.get("rounding", "ROUND_HALF_UP"), ROUND_HALF_UP)

    quote = {}
    if qty is None or price is None:
        quote["net"] = None
        quote["tax"] = None
        quote["gross"] = None
        return quote, ["LINE_INCOMPLETE"]

    net = _round(qty * price, digits, rounding)
    quote["net"] = str(net)

    # 缺税率或含税口径：不能当零税，保留未知。
    if rate is None or not mode:
        quote["tax"] = None
        quote["gross"] = None
        return quote, ["TAX_RULE_MISSING"]

    if mode == "exclusive":
        tax = _round(net * rate, digits, rounding)
        gross = _round(net + tax, digits, rounding)
        quote["net"], quote["tax"], quote["gross"] = str(net), str(tax), str(gross)
    elif mode == "inclusive":
        # 单价已含税：含税总价 = 数量×单价；净价反算；税额取差额。
        gross = _round(qty * price, digits, rounding)
        net_inc = _round(gross / (Decimal(1) + rate), digits, rounding)
        tax = _round(gross - net_inc, digits, rounding)
        quote["net"], quote["tax"], quote["gross"] = str(net_inc), str(tax), str(gross)
    else:
        # 未知含税口径，按缺规则处理。
        quote["tax"] = None
        quote["gross"] = None
        return quote, ["TAX_RULE_MISSING"]

    # 超限检查：gross 已知才判。
    if ceiling is not None:
        over = gross > ceiling
        quote["over_ceiling"] = bool(over)
        if over:
            return quote, ["CEILING_EXCEEDED"]
    return quote, []


def _lines_keyed(lines):
    return {x.get("id"): json.dumps(x, sort_keys=True, ensure_ascii=False) for x in lines}


def _approval_status(payload):
    """返回 (approval_status_or_None, [issue_code, ...])。"""
    approval = payload.get("approval")
    previous = payload.get("previous_lines") or []
    if not approval:
        return None, []
    # 校验登记时的 input_digest 是否真的对应 previous_lines[0]（防伪造/过期引用）。
    if previous:
        expected_digest = hashlib.sha256(
            json.dumps(previous[0], sort_keys=True, ensure_ascii=False).encode()
        ).hexdigest()
        if expected_digest != approval.get("input_digest"):
            return "INVALID", ["APPROVAL_INVALID"]
    # 当前 lines 与批准时的 previous_lines 是否一致；不一致 → 旧批准失效。
    if _lines_keyed(payload.get("lines", [])) != _lines_keyed(previous):
        return "STALE", ["APPROVAL_STALE"]
    return "CURRENT", []


def compute_pricing(request):
    """request = input.json 解析后的 dict（含 payload）。返回结果字典。"""
    payload = request["payload"]
    lines = payload.get("lines", []) or []
    declared_total = _dec(payload.get("declared_total"))

    quotes = {}
    raw_issues = []
    computed_total = Decimal("0")
    for line in lines:
        ceiling = _dec(line.get("ceiling"))
        quote, issues = _quote(line, ceiling)
        quotes[line["id"]] = quote
        raw_issues += issues
        if quote.get("gross") is not None:
            computed_total += Decimal(quote["gross"])

    # 声明总价与计算总价不一致 → TOTAL_MISMATCH。
    if declared_total is not None and computed_total != declared_total:
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
