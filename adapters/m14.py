#!/usr/bin/env python3
"""M14 实际位置与证据审查的验收适配器。

读取测试请求 input.json，调用生产模块 `fdex_tender.documents.review_document`，
把原始信封、中文摘要和验收投影 result.json 写进输出目录。

约定（docs/testing.md）：
- 只读 input.json 与其中登记的输入文件，不读 expected.json；
- 不按 case_id 返回固定答案，全部结论来自本次实际打开的那份 Word；
- 路径按仓库根目录解析，因为检查器以仓库根为工作目录启动本适配器。
"""
import argparse
import json
import shutil
import uuid
from pathlib import Path

from PIL import Image

from fdex_tender.core import digest_file, envelope, load_context, read_json
from fdex_tender.documents import review_document

RESULT_VERSION = "bid-training.acceptance-result.v1"
ROOT = Path.cwd()


def resolve(relative):
    path = Path(relative)
    return path if path.is_absolute() else ROOT / path


def build_evidence(payload, index, upstream_root, trace, selected):
    """证据只引用本次登记过的输入文件，且逐条可回查。"""
    evidence = []
    if trace:
        trace_relative = payload["trace"]
        lines = resolve(trace_relative).read_text(encoding="utf-8").splitlines()
        for entry in trace:
            needle = '"person_id": "{}"'.format(entry.get("person_id"))
            for number, line in enumerate(lines, 1):
                if needle in line:
                    evidence.append({"path": trace_relative, "kind": "line", "line": number, "quote": line})
                    break
    selected_ids = {p["person_id"] for p in selected}
    for asset in index["data"]["assets"]:
        if asset["person_id"] not in selected_ids or not asset.get("relative_path"):
            continue
        image_relative = (Path(payload["upstream_input_root"]) / asset["relative_path"]).as_posix()
        try:
            with Image.open(resolve(image_relative)) as picture:
                box = [0, 0, picture.width, picture.height]
        except (OSError, ValueError):
            continue
        evidence.append({"path": image_relative, "kind": "image_region", "box": box, "field": "certificate_image"})
    return evidence


def attach_requirement_mapping(trace, brief):
    """trace 未声明要求映射时，按生产装配的同一条规则补出映射，并回报补过哪些人。

    `fdex_tender.documents.assemble` 写 trace 时用的规则是：人员资质章节覆盖
    `tender_brief` 中 `in_personnel_scope` 为真的全部要求。这里用同一份登记输入、
    同一条规则还原 M13→M14 交接时应有的内容，不是凭空构造覆盖。

    仅在 key 完全不存在时补齐。只要 key 已存在——哪怕值是 null、字符串、数字或
    空数组——一律原样交给生产审查：适配层不替生产门禁"修正"声明，非法值必须以
    它本来的面目到达 wanted/actual 比对，该 REVIEW 就 REVIEW、该阻断就阻断。
    （第三轮评审：此前按 isinstance(..., list) 判断，null/字符串/数字会被覆盖成
    全量清单，非法声明被洗成正常通过。）
    """
    in_scope = [r["requirement_id"] for r in brief["data"]["requirements"] if r.get("in_personnel_scope")]
    filled = []
    for entry in trace:
        if "requirement_ids" in entry:
            continue
        entry["requirement_ids"] = list(in_scope)
        filled.append(entry.get("person_id"))
    return filled


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    request_path = resolve(args.input)
    request = read_json(request_path)
    payload = request["payload"]
    output = Path(args.output)
    output.mkdir(parents=True, exist_ok=True)

    case = read_json(request_path.parent / "case.json")
    upstream_root = resolve(payload["upstream_input_root"])
    ctx = load_context(upstream_root, "RUN-" + uuid.uuid4().hex[:12])

    upstream = {}
    for relative in payload["upstream_files"]:
        value = read_json(resolve(relative))
        upstream[value["module"]] = value
    matrix, index, matched = upstream["M07"], upstream["M08"], upstream["M09"]

    document_path = resolve(payload["document"])
    document_name = document_path.name
    shutil.copyfile(document_path, output / document_name)
    trace = read_json(resolve(payload["trace"])) if payload.get("trace") else []
    mapping_filled = attach_requirement_mapping(trace, upstream["M04"]) if trace else []

    assembled = envelope(ctx, "M13",
                         {"artifact_status": "PARTIAL_DRAFT",
                          "files": [{"relative_path": document_name, "sha256": digest_file(document_path), "editable": True}],
                          "trace": trace,
                          "selected_person_ids": [p["person_id"] for p in matched["data"]["selected"]],
                          "selection_digest": matched["data"]["selection_digest"]},
                         matrix["input_refs"] + index["input_refs"])

    review = review_document(ctx, matrix, index, matched, assembled, output, release=payload.get("release"))
    (output / "review_result.json").write_text(json.dumps(review, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    data = review["data"]
    summary = ["# M14 实际位置与证据审查摘要", "",
               "处理状态：{}".format(review["processing_status"]),
               "受审文件：{}（SHA256 {}）".format(document_name, data.get("inspected_document_sha256")),
               "范围：人员资质章节；全本标书、技术方案、商务报价与正式递交不在本次范围。", "",
               "## 定位到的问题"]
    summary += ["- {}：{}；{}接手；{}".format(i["code"], i["message"], i["owner_role"], i["next_action"])
                for i in review["issues"]] or ["本次声明范围内未发现已知阻断问题；业务确认与放行另行记录。"]
    if data.get("paragraph_index_drift"):
        summary += ["", "## 仅提示不改判的段落漂移",
                    "- " + "; ".join("{} 声称第{}段、实际第{}段".format(d["person_id"], d["claimed_paragraph_index"], d["actual_paragraph_index"])
                                     for d in data["paragraph_index_drift"])]
    if mapping_filled:
        summary += ["", "## 要求映射的来源",
                    "- 传入 trace 未声明 requirement_ids：{}。本适配器按生产装配的同一规则"
                    "（tender_brief.json 中 in_personnel_scope 为真的要求）补出映射后再送审；"
                    "覆盖比对仍由生产审查执行，未跳过。".format("、".join(str(p) for p in mapping_filled))]
    summary += ["", "机器可读详情：review_result.json"]
    (output / "review_result.md").write_text("\n".join(summary) + "\n", encoding="utf-8")

    facts = {"scope_total": data["coverage"]["scope_total"],
             "full_total": data["coverage"]["full_total"],
             "full_tender_complete": data["coverage"]["full_tender_complete"],
             "visual_review": data.get("visual_review"),
             "location_mismatches": data.get("location_mismatches", []),
             "inspected_document_sha256": data.get("inspected_document_sha256"),
             "located_issue_kinds": data.get("located_issue_kinds", []),
             "release_allowed": data.get("release_allowed")}
    if "release_hash_matches" in data:
        facts["release_hash_matches"] = data["release_hash_matches"]
    if mapping_filled:
        facts["requirement_mapping_source"] = "tender_brief.json:in_personnel_scope"
        facts["requirement_mapping_filled_person_ids"] = mapping_filled

    result = {"schema_version": RESULT_VERSION,
              "module": "M14",
              "case_id": request["case_id"],
              "source_mode": request["source_mode"],
              "status": review["processing_status"],
              "facts": facts,
              "issue_codes": [i["code"] for i in review["issues"]],
              "input_hashes": {item["path"]: digest_file(resolve(item["path"])) for item in case["inputs"]},
              "evidence": build_evidence(payload, index, upstream_root, trace, matched["data"]["selected"]),
              "artifacts": {},
              "real_submission_authorized": False}
    (output / "result.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"case_id": request["case_id"], "status": review["processing_status"],
                      "issue_codes": result["issue_codes"]}, ensure_ascii=False))
    return 0 if review["processing_status"] == "COMPLETE" else 2


if __name__ == "__main__":
    raise SystemExit(main())
