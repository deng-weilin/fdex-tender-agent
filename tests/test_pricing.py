#!/usr/bin/env python3
"""M12 报价计算模块的专项回归（自包含，不依赖 samples/，供隔离 CI 实际执行）。

覆盖维护者复审指出的三类门禁缺口，以及 4 个验收用例的核心算术：
- P1 批准状态：REJECTED / scope 不符 / 摘要不符 一律不得 CURRENT；多组报价的批准
  因单条摘要无法覆盖每行基线，明确不支持（APPROVAL_MULTI_LINE_UNSUPPORTED），
  包括"把某行新价同步写进 lines 与 previous_lines 伪造基线"的攻击；
- P2 舍入：未支持规则必须明确报错、绝不静默降级；digits 非法不得崩溃；
- P3 数值：缺失与非法严格区分，非法金额/限价/总价/NaN/Infinity/负数一律拒绝。
外加正常路径与单行批准（CURRENT / STALE）对照，防止"修门禁修坏正常路径"。
"""
import copy
import hashlib
import json
import unittest
from decimal import Decimal

from fdex_tender.pricing import compute_pricing

HALF_UP = "ROUND_HALF_UP"


def line(**over):
    base = {
        "id": "PRICE-01", "quantity": "3", "unit_price": "100.00", "tax_rate": "0.06",
        "tax_mode": "exclusive", "rounding": HALF_UP, "digits": 2, "ceiling": "400.00",
    }
    base.update(over)
    return base


def request(lines, **payload_over):
    payload = {
        "lines": lines, "rounding_stage": "LINE_THEN_TOTAL",
        "independent_quotes": True, "approval": None,
    }
    payload.update(payload_over)
    return {
        "schema_version": "bid-training.acceptance-input.v1", "module": "M12",
        "case_id": "UNIT", "source_mode": "DEMO", "payload": payload,
    }


def line_digest(line_obj):
    return hashlib.sha256(json.dumps(line_obj, sort_keys=True, ensure_ascii=False).encode()).hexdigest()


class NormalPathTests(unittest.TestCase):
    def test_exclusive_half_up_boundary(self):
        # 2.5 × 99.99 = 249.975 → HALF_UP 249.98（float 会错成 249.97）。
        result = compute_pricing(request([line(quantity="2.5", unit_price="99.99", tax_rate="0.13")]))
        quote = result["facts"]["quotes"]["PRICE-01"]
        self.assertEqual(result["status"], "COMPLETE")
        self.assertEqual(result["issue_codes"], [])
        self.assertEqual(quote["net"], "249.98")
        self.assertEqual(quote["tax"], "32.50")
        self.assertEqual(quote["gross"], "282.48")

    def test_inclusive_difference_method_keeps_identity(self):
        result = compute_pricing(request([line(quantity="10", unit_price="12.50", tax_rate="0.13",
                                             tax_mode="inclusive", ceiling="150.00")]))
        quote = result["facts"]["quotes"]["PRICE-01"]
        self.assertEqual(result["status"], "COMPLETE")
        self.assertEqual(quote["net"], "110.62")
        self.assertEqual(quote["tax"], "14.38")
        self.assertEqual(quote["gross"], "125.00")
        self.assertEqual(Decimal(quote["net"]) + Decimal(quote["tax"]), Decimal(quote["gross"]))

    def test_valid_control_is_complete(self):
        result = compute_pricing(request([line()]))
        self.assertEqual(result["status"], "COMPLETE")
        self.assertEqual(result["issue_codes"], [])
        self.assertFalse(result["real_submission_authorized"])
        self.assertFalse(result["facts"]["release_allowed"])

    def test_missing_tax_rate_is_never_zero_tax(self):
        result = compute_pricing(request([line(tax_rate=None, tax_mode=None)]))
        quote = result["facts"]["quotes"]["PRICE-01"]
        self.assertEqual(result["status"], "REVIEW")
        self.assertIn("TAX_RULE_MISSING", result["issue_codes"])
        self.assertEqual(quote["net"], "300.00")
        self.assertIsNone(quote["tax"])
        self.assertIsNone(quote["gross"])

    def test_ceiling_exceeded_is_blocked(self):
        result = compute_pricing(request([line(ceiling="300.00")]))
        self.assertEqual(result["status"], "BLOCKED")
        self.assertIn("CEILING_EXCEEDED", result["issue_codes"])
        self.assertTrue(result["facts"]["quotes"]["PRICE-01"]["over_ceiling"])

    def test_declared_total_mismatch_is_blocked(self):
        result = compute_pricing(request([line()], declared_total="317.00"))
        self.assertEqual(result["status"], "BLOCKED")
        self.assertIn("TOTAL_MISMATCH", result["issue_codes"])


class ApprovalGateTests(unittest.TestCase):
    """P1：批准必须"四证齐备"才 CURRENT。"""

    def _approved(self, prev_price="100.00", cur_price="100.00", scope="PRICE-01", status="APPROVED"):
        previous = line(unit_price=prev_price)
        approval = {"status": status, "scope": scope, "input_digest": line_digest(previous)}
        return request([line(unit_price=cur_price)], approval=approval, previous_lines=[previous])

    def test_approved_and_unchanged_is_current(self):
        result = compute_pricing(self._approved())
        self.assertEqual(result["facts"]["approval_status"], "CURRENT")
        self.assertEqual(result["status"], "COMPLETE")

    def test_changed_price_makes_approval_stale(self):
        result = compute_pricing(self._approved(prev_price="100.00", cur_price="101.00"))
        self.assertEqual(result["facts"]["approval_status"], "STALE")
        self.assertIn("APPROVAL_STALE", result["issue_codes"])
        self.assertEqual(result["status"], "REVIEW")

    def test_rejected_approval_is_not_current(self):
        # 批准被驳回，但摘要与内容都"看起来"匹配——不得因此标成 CURRENT。
        result = compute_pricing(self._approved(status="REJECTED"))
        self.assertNotEqual(result["facts"]["approval_status"], "CURRENT")
        self.assertIn("APPROVAL_INVALID", result["issue_codes"])
        self.assertEqual(result["status"], "REVIEW")

    def test_scope_pointing_elsewhere_is_not_current(self):
        result = compute_pricing(self._approved(scope="LOT-999-OTHER"))
        self.assertNotEqual(result["facts"]["approval_status"], "CURRENT")
        self.assertIn("APPROVAL_INVALID", result["issue_codes"])

    def test_digest_mismatch_is_not_current(self):
        request_obj = self._approved()
        request_obj["payload"]["approval"]["input_digest"] = "0" * 64
        result = compute_pricing(request_obj)
        self.assertNotEqual(result["facts"]["approval_status"], "CURRENT")
        self.assertIn("APPROVAL_INVALID", result["issue_codes"])

    def test_scope_must_cover_every_line(self):
        # 多组独立报价：单条摘要无法覆盖每行，批准一律不支持（不是"已获批准"）。
        line1 = line(unit_price="100.00")
        line2 = line(id="PRICE-02", quantity="5", unit_price="200.00", ceiling="4000.00")
        approval = {"status": "APPROVED", "scope": "PRICE-01", "input_digest": line_digest(line1)}
        result = compute_pricing(request([line1, line2], approval=approval, previous_lines=[line1, line2]))
        self.assertNotEqual(result["facts"]["approval_status"], "CURRENT")
        self.assertIn("APPROVAL_MULTI_LINE_UNSUPPORTED", result["issue_codes"])

    def test_multiline_forged_baseline_stays_not_current(self):
        """第三轮复审核心指控：多组报价时把第二行新价**同时**写进 lines 与
        previous_lines（伪造"变更前基线"），scope 覆盖两行，input_digest 仍绑
        第一行旧价——摘要只覆盖第一行，第二行新价从未被批准，绝不能 CURRENT。
        """
        line1_old = line(id="PRICE-01", unit_price="100.00")
        line2_new = line(id="PRICE-02", quantity="5", unit_price="250.00", ceiling="4000.00")
        approval = {"status": "APPROVED", "scope": "PRICE-01,PRICE-02",
                    "input_digest": line_digest(line1_old)}
        result = compute_pricing(request(
            [copy.deepcopy(line1_old), copy.deepcopy(line2_new)],
            approval=approval,
            previous_lines=[copy.deepcopy(line1_old), copy.deepcopy(line2_new)],
        ))
        self.assertNotEqual(result["facts"]["approval_status"], "CURRENT")
        self.assertIn("APPROVAL_MULTI_LINE_UNSUPPORTED", result["issue_codes"])
        self.assertEqual(result["status"], "REVIEW")

    def test_multiline_with_too_few_baselines_not_current(self):
        # 两行报价却只有一条基线：同样不得 CURRENT（无法逐行绑定）。
        line1 = line(id="PRICE-01", unit_price="100.00")
        line2 = line(id="PRICE-02", quantity="5", unit_price="200.00", ceiling="4000.00")
        approval = {"status": "APPROVED", "scope": "PRICE-01,PRICE-02", "input_digest": line_digest(line1)}
        result = compute_pricing(request([line1, line2], approval=approval, previous_lines=[line1]))
        self.assertNotEqual(result["facts"]["approval_status"], "CURRENT")
        self.assertIn("APPROVAL_MULTI_LINE_UNSUPPORTED", result["issue_codes"])

    def test_single_line_gates_unaffected_by_multiline_rule(self):
        # 对照：多组不支持不得误伤单行——单行有效仍 CURRENT，单行旧价仍 STALE。
        self.assertEqual(compute_pricing(self._approved())["facts"]["approval_status"], "CURRENT")
        stale = compute_pricing(self._approved(prev_price="100.00", cur_price="101.00"))
        self.assertEqual(stale["facts"]["approval_status"], "STALE")
        self.assertIn("APPROVAL_STALE", stale["issue_codes"])

    def test_digest_convention_matches_kit_change_case(self):
        # 约定：input_digest = previous_lines[0] 按 sort_keys 序列化后的 sha256（非文件哈希）。
        previous = line()
        self.assertEqual(line_digest(previous),
                         "78247d010cedc5d82787a95345306c1fc30aa7b219c91f5b0f80c1903715ebbd")


class RoundingGateTests(unittest.TestCase):
    """P2：未支持规则必须报错，绝不静默降级；digits 非法不得崩溃。"""

    def test_unsupported_rounding_not_silently_downgraded(self):
        result = compute_pricing(request([line(quantity="2.5", unit_price="99.99", rounding="ROUND_DOWN")]))
        self.assertEqual(result["status"], "BLOCKED")
        self.assertIn("ROUNDING_RULE_UNSUPPORTED", result["issue_codes"])
        # 绝不静默按 HALF_UP 算出 249.98。
        self.assertIsNone(result["facts"]["quotes"]["PRICE-01"]["net"])

    def test_unknown_rounding_rule_blocked(self):
        result = compute_pricing(request([line(rounding="TOTALLY_MADE_UP")]))
        self.assertEqual(result["status"], "BLOCKED")
        self.assertIn("ROUNDING_RULE_UNSUPPORTED", result["issue_codes"])

    def test_non_integer_digits_does_not_crash(self):
        result = compute_pricing(request([line(digits="abc")]))
        self.assertEqual(result["status"], "BLOCKED")
        self.assertIn("ROUNDING_DIGITS_INVALID", result["issue_codes"])

    def test_negative_digits_blocked(self):
        result = compute_pricing(request([line(digits=-1)]))
        self.assertEqual(result["status"], "BLOCKED")
        self.assertIn("ROUNDING_DIGITS_INVALID", result["issue_codes"])


class NumericGateTests(unittest.TestCase):
    """P3：缺失 vs 非法严格区分；非法数值/限价/总价一律拒绝，不得静默跳过或崩溃。"""

    def test_invalid_ceiling_is_not_treated_as_absent(self):
        result = compute_pricing(request([line(quantity="100", unit_price="100.00", ceiling="abc")]))
        self.assertEqual(result["status"], "BLOCKED")
        self.assertIn("CEILING_INVALID", result["issue_codes"])

    def test_invalid_declared_total_is_not_treated_as_absent(self):
        result = compute_pricing(request([line()], declared_total="abc"))
        self.assertEqual(result["status"], "BLOCKED")
        self.assertIn("DECLARED_TOTAL_INVALID", result["issue_codes"])

    def test_negative_quantity_blocked(self):
        result = compute_pricing(request([line(quantity="-5")]))
        self.assertEqual(result["status"], "BLOCKED")
        self.assertIn("QUANTITY_INVALID", result["issue_codes"])

    def test_negative_unit_price_blocked(self):
        result = compute_pricing(request([line(unit_price="-0.01")]))
        self.assertEqual(result["status"], "BLOCKED")
        self.assertIn("UNIT_PRICE_INVALID", result["issue_codes"])

    def test_nan_unit_price_rejected_without_crash(self):
        result = compute_pricing(request([line(unit_price="NaN")]))
        self.assertEqual(result["status"], "BLOCKED")
        self.assertIn("UNIT_PRICE_INVALID", result["issue_codes"])

    def test_infinite_tax_rate_rejected_without_crash(self):
        result = compute_pricing(request([line(tax_rate="Infinity")]))
        self.assertEqual(result["status"], "BLOCKED")
        self.assertIn("TAX_RATE_INVALID", result["issue_codes"])

    def test_absent_ceiling_still_computes(self):
        # 真正的"未提供"（缺 ceiling）应当照常计算，不误伤。
        result = compute_pricing(request([line(ceiling=None)]))
        self.assertEqual(result["status"], "COMPLETE")
        self.assertNotIn("over_ceiling", result["facts"]["quotes"]["PRICE-01"])


if __name__ == "__main__":
    unittest.main()
