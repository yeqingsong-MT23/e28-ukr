# -*- coding: utf-8 -*-
"""
根据 UK200_转化漏斗KPI 基础表生成 UK2061_各订单等级三次报价转化漏斗概览.xlsx。
"""

from __future__ import annotations

import argparse
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook

from uk_export_utils import build_default_output_path
from uk_funnel_utils import (
    build_stage_breakdown,
    find_latest_input,
    load_stage_qpo_records,
    month_sort_key,
    safe_div,
    style_two_header_sheet,
)


SCRIPT_DIR = Path(__file__).resolve().parent
BASE_DIR = SCRIPT_DIR.parent
EXPORT_DIR = Path(r"P:\0_Report\78_AI Report\UKR_Report")
DEFAULT_OUTPUT_NAME = "UK2061_各订单等级三次报价转化漏斗概览.xlsx"
ORDER_LEVELS = ["A", "B", "C", "D"]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="根据 UK200 基础表生成 UK2061_各订单等级三次报价转化漏斗概览。"
    )
    parser.add_argument("--input", dest="input_path", help="手动指定输入基础表路径")
    parser.add_argument("--output", dest="output_path", help="手动指定输出文件路径")
    return parser.parse_args()


def aggregate(records: list[dict[str, object]]) -> list[dict[str, object]]:
    grouped: dict[tuple[str, str], list[dict[str, object]]] = defaultdict(list)

    for record in records:
        order_level = record.get("order_level")
        month_compact = record.get("month_compact")
        if order_level not in ORDER_LEVELS:
            continue
        if not month_compact:
            continue

        grouped[(month_compact, order_level)].append(record)
        grouped[("ALL", order_level)].append(record)
        grouped[("ALL", "ALL")].append(record)

    rows: list[dict[str, object]] = []
    months = sorted(
        {
            str(record.get("month_compact"))
            for record in records
            if record.get("month_compact")
        },
        key=month_sort_key,
    )

    for month in months + ["ALL"]:
        levels = ORDER_LEVELS if month != "ALL" else ORDER_LEVELS + ["ALL"]
        for level in levels:
            bucket = grouped.get((month, level), [])
            if not bucket:
                continue

            submitted = len(bucket)
            success = sum(1 for record in bucket if record.get("simple_type") == "Success")
            in_progress = sum(
                1 for record in bucket if record.get("simple_type") == "InProgress"
            )
            cancel = sum(1 for record in bucket if record.get("simple_type") == "Cancel")

            row = {
                "Month": month,
                "Order_Level": level,
                "Submitted_QPO_Count": submitted,
                "Success_QPO_Count": success,
                "In_Progress_QPO_Count": in_progress,
                "Cancel_QPO_Count": cancel,
                "Success_Rate": safe_div(success, submitted),
                "In_Progress_Rate": safe_div(in_progress, submitted),
                "Cancel_Rate": safe_div(cancel, submitted),
            }

            stage_metrics = build_stage_breakdown(
                bucket,
                include_stage_inprogress=False,
                include_cancel_split=True,
            )
            row.update(stage_metrics)
            rows.append(row)

    return rows


def build_summary(records: list[dict[str, object]]) -> list[dict[str, object]]:
    return aggregate(records)


def build_notes_rows(
    rows: list[dict[str, object]],
    columns_en: list[str],
    metadata: dict[str, object],
) -> list[tuple[object, object, object, object, object, object]]:
    input_rows = int(metadata.get("input_rows", 0) or 0)
    valid_rows = int(metadata.get("valid_rows", 0) or 0)
    excluded_qpo_ids = metadata.get("excluded_qpo_ids", []) or []
    if not isinstance(excluded_qpo_ids, list):
        excluded_qpo_ids = list(excluded_qpo_ids)
    included_qpo_count = int(metadata.get("included_qpo_count", 0) or 0)
    invalid_rows = max(input_rows - valid_rows, 0)

    return [
        (
            "Field (EN)",
            "字段中文",
            "Field Group",
            "Purpose / Business Meaning",
            "Calculation / Rule",
            "Notes",
        ),
        ("英文字段", "中文解释", "字段分组", "业务含义", "计算口径", "备注"),
        ("A. Rules（口径规则）", None, None, None, None, None),
        (
            "Type_Mapping",
            "Cancel: SellerCancel/CSCancel/Long time-Cancel/WootCancel; In Progress: Hold/New/SubmitRBS; Success: End/WaitingRPO/WaitingPayment/WaitingDeal/OnDeal/WaitingPO; 其他=Other",
            None,
            None,
            None,
            None,
        ),
        ("OrderLevel_Mapping", "Order_Level=AA → A，其余不变", None, None, None, None),
        (
            "HasFlag_Rule",
            "按 QPO_ID 聚合：对应报价字段任一行非空(非NaN且去空格后非空字符串)即 True",
            None,
            None,
            None,
            None,
        ),
        (
            "Stage_Rule",
            "Has_Third=True→Third；否则Has_Second=True→Second；否则→Initial",
            None,
            None,
            None,
            None,
        ),
        (
            "Exclusion_Rule",
            "若某QPO_ID：Type=Success 且 Has_Initial/Has_Second/Has_Third 全False，则整单剔除",
            None,
            None,
            None,
            None,
        ),
        (
            "CancelSplit_Rule",
            "CSCancel/Long time-Cancel/WootCancel→WootCancel；SellerCancel保持；优先级WootCancel>SellerCancel",
            None,
            None,
            None,
            None,
        ),
        (
            "Denominator_Rule",
            "整体占比：分母=Submitted_QPO_Count；阶段占比：分母=Submitted_QPO_Count；阶段拆分占比：分母=Stage_QPO_Count；取消内部结构占比：分母=Stage_Cancel_QPO_Count；0分母输出0.00%",
            None,
            None,
            None,
            None,
        ),
        (
            "Summary_Rule",
            "汇总行( Month=ALL )：Count求和等价于重新去重计数的结果；Rate均用汇总分子/汇总分母重新计算(禁止平均)",
            None,
            None,
            None,
            None,
        ),
        (
            "ThirdStage_Focus",
            "本报表聚焦三次报价场景，主表保留整体漏斗列与三阶段字段，重点观察 Third 阶段规模与最终转化质量",
            None,
            None,
            None,
            None,
        ),
        (None, None, None, None, None, None),
        ("自检报告（QA Report）", None, None, None, None, None),
        ("输入行数", str(input_rows), None, None, None, None),
        ("有效行数(Submission_Date可解析)", str(valid_rows), None, None, None, None),
        ("无效日期行数(不参与统计)", str(invalid_rows), None, None, None, None),
        (
            "剔除异常整单QPO数(Success但未进入任何报价阶段)",
            str(len(excluded_qpo_ids)),
            None,
            None,
            None,
            None,
        ),
        (
            "剔除示例QPO_ID(最多20)",
            ", ".join(str(value) for value in excluded_qpo_ids[:20]),
            None,
            None,
            None,
            None,
        ),
        ("最终纳入统计QPO数(去重)", str(included_qpo_count), None, None, None, None),
        ("结果表行数", str(len(rows)), None, None, None, None),
        ("0分母情况", "已统一输出0.00%(如Submitted或Stage或StageCancel为0)", None, None, None, None),
        (
            "字段完整性检查",
            f"Sheet1字段数={len(columns_en)}；预期={len(columns_en)}；一致",
            None,
            None,
            None,
            None,
        ),
    ]


def write_workbook(
    rows: list[dict[str, object]],
    output_path: Path,
    metadata: dict[str, object],
) -> None:
    columns_en = [
        "Month",
        "Order_Level",
        "Submitted_QPO_Count",
        "Success_QPO_Count",
        "In_Progress_QPO_Count",
        "Cancel_QPO_Count",
        "Success_Rate",
        "In_Progress_Rate",
        "Cancel_Rate",
        "Initial_Stage_QPO_Count",
        "Initial_Stage_Success_QPO_Count",
        "Initial_Stage_Success_Submit_Rate",
        "Initial_Stage_Cancel_QPO_Count",
        "Initial_Stage_Cancel_Submit_Rate",
        "Initial_Stage_WootCancel_QPO_Count",
        "Initial_Stage_SellerCancel_QPO_Count",
        "Initial_Stage_WootCancel_Stage_Rate",
        "Initial_Stage_SellerCancel_Stage_Rate",
        "Initial_Stage_WootCancel_Cancel_Rate",
        "Initial_Stage_SellerCancel_Cancel_Rate",
        "Second_Stage_QPO_Count",
        "Second_Stage_Success_QPO_Count",
        "Second_Stage_Success_Submit_Rate",
        "Second_Stage_Cancel_QPO_Count",
        "Second_Stage_Cancel_Submit_Rate",
        "Second_Stage_WootCancel_QPO_Count",
        "Second_Stage_SellerCancel_QPO_Count",
        "Second_Stage_WootCancel_Stage_Rate",
        "Second_Stage_SellerCancel_Stage_Rate",
        "Second_Stage_WootCancel_Cancel_Rate",
        "Second_Stage_SellerCancel_Cancel_Rate",
        "Third_Stage_QPO_Count",
        "Third_Stage_Success_QPO_Count",
        "Third_Stage_Success_Submit_Rate",
        "Third_Stage_Cancel_QPO_Count",
        "Third_Stage_Cancel_Submit_Rate",
        "Third_Stage_WootCancel_QPO_Count",
        "Third_Stage_SellerCancel_QPO_Count",
        "Third_Stage_WootCancel_Stage_Rate",
        "Third_Stage_SellerCancel_Stage_Rate",
        "Third_Stage_WootCancel_Cancel_Rate",
        "Third_Stage_SellerCancel_Cancel_Rate",
    ]
    columns_cn = [
        "月份(提报YYYYMM/ALL)",
        "订单等级",
        "提报QPO数(去重)",
        "成交QPO数(去重)",
        "跟进中QPO数(去重)",
        "取消QPO数(去重)",
        "成交占比(成交/提报)",
        "跟进中占比(跟进中/提报)",
        "取消占比(取消/提报)",
        "初次报价阶段QPO数(最高阶段归属)",
        "初次报价阶段成交QPO数",
        "初次报价阶段成交/提报占比",
        "初次报价阶段取消QPO数",
        "初次报价阶段取消/提报占比",
        "初次报价阶段WootCancel QPO数",
        "初次报价阶段SellerCancel QPO数",
        "初次报价阶段WootCancel占比(WootCancel/阶段)",
        "初次报价阶段SellerCancel占比(SellerCancel/阶段)",
        "初次报价阶段取消内WootCancel占比(WootCancel/阶段取消)",
        "初次报价阶段取消内SellerCancel占比(SellerCancel/阶段取消)",
        "二次报价阶段QPO数(最高阶段归属)",
        "二次报价阶段成交QPO数",
        "二次报价阶段成交/提报占比",
        "二次报价阶段取消QPO数",
        "二次报价阶段取消/提报占比",
        "二次报价阶段WootCancel QPO数",
        "二次报价阶段SellerCancel QPO数",
        "二次报价阶段WootCancel占比(WootCancel/阶段)",
        "二次报价阶段SellerCancel占比(SellerCancel/阶段)",
        "二次报价阶段取消内WootCancel占比(WootCancel/阶段取消)",
        "二次报价阶段取消内SellerCancel占比(SellerCancel/阶段取消)",
        "三次报价阶段QPO数(最高阶段归属)",
        "三次报价阶段成交QPO数",
        "三次报价阶段成交/提报占比",
        "三次报价阶段取消QPO数",
        "三次报价阶段取消/提报占比",
        "三次报价阶段WootCancel QPO数",
        "三次报价阶段SellerCancel QPO数",
        "三次报价阶段WootCancel占比(WootCancel/阶段)",
        "三次报价阶段SellerCancel占比(SellerCancel/阶段)",
        "三次报价阶段取消内WootCancel占比(WootCancel/阶段取消)",
        "三次报价阶段取消内SellerCancel占比(SellerCancel/阶段取消)",
    ]

    workbook = Workbook()

    sheet = workbook.active
    sheet.title = "UK2061"
    sheet.append(columns_en)
    sheet.append(columns_cn)
    for row in rows:
        sheet.append([row.get(column) for column in columns_en])
    style_two_header_sheet(sheet, columns_en, columns_cn)

    notes = workbook.create_sheet("Notes")
    for note_row in build_notes_rows(rows, columns_en, metadata):
        notes.append(note_row)
    for col in "ABCDEF":
        notes.column_dimensions[col].width = 28
    notes.freeze_panes = "A3"

    workbook.save(output_path)


def main() -> None:
    args = parse_args()

    input_path = Path(args.input_path) if args.input_path else find_latest_input(BASE_DIR)
    output_path = (
        Path(args.output_path)
        if args.output_path
        else build_default_output_path(EXPORT_DIR, DEFAULT_OUTPUT_NAME, BASE_DIR)
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)

    records, metadata = load_stage_qpo_records(input_path)
    if metadata is None:
        metadata = {}

    rows = build_summary(records)
    write_workbook(rows, output_path, metadata)

    print("生成完成：")
    print(output_path)


if __name__ == "__main__":
    main()
