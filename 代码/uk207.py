# -*- coding: utf-8 -*-
"""
根据 UK200_转化漏斗KPI 基础表生成 UK207｜不同报价方式转化漏斗.xlsx。
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
DEFAULT_OUTPUT_NAME = "UK207｜不同报价方式转化漏斗.xlsx"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="根据 UK200 基础表生成 UK207｜不同报价方式转化漏斗。"
    )
    parser.add_argument("--input", dest="input_path", help="指定输入基础表路径")
    parser.add_argument("--output", dest="output_path", help="指定输出 Excel 路径")
    return parser.parse_args()


def validate_metadata(metadata: dict[str, object]) -> None:
    required_keys = [
        "input_rows",
        "valid_rows",
        "excluded_qpo_ids",
        "included_qpo_count",
    ]
    missing = [key for key in required_keys if key not in metadata]
    if missing:
        raise ValueError(f"load_stage_qpo_records 返回的 metadata 缺少必要字段: {', '.join(missing)}")


def aggregate(records: list[dict[str, object]]) -> tuple[list[dict[str, object]], list[str]]:
    grouped: dict[tuple[str, str], list[dict[str, object]]] = defaultdict(list)
    pricing_modes = sorted(
        {
            str(record.get("pricing_mode") or "").strip()
            for record in records
            if str(record.get("pricing_mode") or "").strip()
        }
    )

    for record in records:
        pricing_mode = str(record.get("pricing_mode") or "").strip()
        if not pricing_mode:
            continue
        month = str(record.get("month_dash") or "").strip()
        if not month:
            continue
        grouped[(month, pricing_mode)].append(record)
        grouped[("ALL", pricing_mode)].append(record)
        grouped[("ALL", "ALL")].append(record)

    months = sorted(
        {
            str(record.get("month_dash") or "").strip()
            for record in records
            if str(record.get("month_dash") or "").strip()
        },
        key=month_sort_key,
    )

    rows: list[dict[str, object]] = []
    for month in months + ["ALL"]:
        modes = pricing_modes if month != "ALL" else pricing_modes + ["ALL"]
        for pricing_mode in modes:
            bucket = grouped.get((month, pricing_mode), [])
            if not bucket:
                continue

            submitted = len(bucket)
            success = sum(1 for record in bucket if record.get("simple_type") == "Success")
            in_progress = sum(1 for record in bucket if record.get("simple_type") == "InProgress")
            cancel = sum(1 for record in bucket if record.get("simple_type") == "Cancel")

            row = {
                "Month": month,
                "Pricing_Mode": pricing_mode,
                "Submitted_QPO_Count": submitted,
                "Success_QPO_Count": success,
                "In_Progress_QPO_Count": in_progress,
                "Cancel_QPO_Count": cancel,
                "Success_Rate": safe_div(success, submitted),
                "In_Progress_Rate": safe_div(in_progress, submitted),
                "Cancel_Rate": safe_div(cancel, submitted),
            }
            row.update(
                build_stage_breakdown(
                    bucket,
                    include_stage_inprogress=True,
                    include_cancel_split=True,
                )
            )
            rows.append(row)

    return rows, pricing_modes


def build_notes_rows(metadata: dict[str, object], pricing_modes: list[str]) -> list[tuple[object, ...]]:
    validate_metadata(metadata)

    pricing_mode_text = "；".join(pricing_modes) if pricing_modes else "无有效报价方式"
    excluded_qpo_ids = metadata.get("excluded_qpo_ids") or []
    if isinstance(excluded_qpo_ids, (list, tuple, set)):
        excluded_sample = "；".join(str(x) for x in list(excluded_qpo_ids)[:20]) or None
    else:
        excluded_sample = str(excluded_qpo_ids)

    return [
        (
            "Field (EN)",
            "字段中文",
            "Field Group",
            "Business Meaning",
            "Calculation / Rule",
            "Notes",
        ),
        (
            "字段英文",
            "字段中文",
            "字段分组",
            "业务含义",
            "计算方式 / 口径",
            "补充说明",
        ),
        (
            "Pricing_Mode_Filter",
            "报价方式过滤规则",
            "Rule",
            "仅统计存在有效报价方式的QPO",
            "Pricing_Mode 为空、空字符串或仅空格的记录不进入结果表",
            "主表仅输出 Pricing_Mode 非空数据",
        ),
        (
            "Exclusion_Rule",
            "异常成交剔除",
            "Rule",
            "剔除未进入任何报价阶段却标记为成交的异常QPO，避免漏斗失真",
            "在QPO粒度：Type=Success 且 Has_Initial/Has_Second/Has_Third 均为False，则剔除该QPO全部记录",
            "先剔除再判定Stage与汇总",
        ),
        (
            "Type_Mapping",
            "状态映射",
            "Rule",
            "将QPO_Status映射到统一漏斗状态Type",
            "Success=End/WaitingRPO/WaitingPayment/WaitingDeal/OnDeal/WaitingPO；In Progress=Hold/New/SubmitRBS；WootCancel=CSCancel/Long time-Cancel/WootCancel；SellerCancel=SellerCancel；其他=Other",
            "Cancel = WootCancel + SellerCancel；Other不计入Success/InProgress/Cancel",
        ),
        (
            "Cancel_Split_Rule",
            "取消拆分规则",
            "Rule",
            "区分取消来源，支持取消结构分析",
            "Cancel总口径=Type in {WootCancel,SellerCancel}；并分别统计WootCancel与SellerCancel",
            "阶段内同样按该口径拆分；Cancel = WootCancel + SellerCancel",
        ),
        (
            "HasFlag_Rule",
            "报价进入标记规则",
            "Rule",
            "判断QPO是否进入过初审/二审/三审报价流程",
            "在QPO粒度：任一行 Initial_Quote/Second_Quote/MVM_Approved_Price 非空(非NaN/非空字符串/非纯空格)即为True",
            "不要求值必须为Yes，只要有值即可",
        ),
        (
            "Stage_Rule",
            "阶段归属规则",
            "Rule",
            "每个QPO只归属到最高进入的报价阶段",
            "Has_Third=True→Third；否则Has_Second=True→Second；否则→Initial",
            "Initial阶段不要求Initial_Quote必须有值",
        ),
        (
            "Denominator_Rule",
            "分母统一规则",
            "Rule",
            "统一总体与阶段提报占比分母，便于横向对比",
            "Success_Rate / In_Progress_Rate / Cancel_Rate 分母=Submitted_QPO_Count；所有 *_Submit_Rate 分母也为 Submitted_QPO_Count",
            "分母为0时结果记为0",
        ),
        (
            "StageRate_Rule",
            "阶段内占比规则",
            "Rule",
            "定义阶段内取消来源占比的分母",
            "X_Stage_Rate 分母=该Stage的QPO总数；X_Cancel_Rate 分母=该Stage的Cancel QPO数",
            "分母为0时结果记为0",
        ),
        (
            "Summary_Rows",
            "汇总行规则",
            "Rule",
            "提供按报价方式及全量汇总，便于总览",
            "主表包含 Month=ALL, Pricing_Mode=各报价方式 以及 Month=ALL, Pricing_Mode=ALL；汇总行Count=明细求和；Rate=汇总分子/汇总分母重新计算",
            "禁止对Rate做简单平均",
        ),
        (
            "Detected_Pricing_Modes",
            "识别到的报价方式",
            "QA",
            "用于检查Pricing_Mode脏值或枚举异常",
            pricing_mode_text,
            f"识别数量={len(pricing_modes)}",
        ),
        (
            "QA_Input_Rows",
            "输入行数",
            "QA",
            "原始输入行数",
            str(metadata["input_rows"]),
            None,
        ),
        (
            "QA_Valid_Rows",
            "有效日期行数",
            "QA",
            "Submission_Date可解析行数",
            str(metadata["valid_rows"]),
            None,
        ),
        (
            "QA_Excluded_QPOs",
            "剔除异常成交QPO数",
            "QA",
            "未进入任何报价阶段却标记为成交的QPO",
            str(len(metadata["excluded_qpo_ids"])),
            excluded_sample,
        ),
        (
            "QA_Final_QPOs",
            "最终纳入统计QPO数",
            "QA",
            "去重后最终统计QPO数",
            str(metadata["included_qpo_count"]),
            None,
        ),
    ]


def write_workbook(
    rows: list[dict[str, object]],
    output_path: Path,
    metadata: dict[str, object],
    pricing_modes: list[str],
) -> None:
    columns_en = [
        "Month",
        "Pricing_Mode",
        "Submitted_QPO_Count",
        "Success_QPO_Count",
        "In_Progress_QPO_Count",
        "Cancel_QPO_Count",
        "Success_Rate",
        "In_Progress_Rate",
        "Cancel_Rate",
        "Initial_Stage_QPO_Count",
        "Initial_Stage_Success_QPO_Count",
        "Initial_Stage_Cancel_QPO_Count",
        "Initial_Stage_In_Progress_QPO_Count",
        "Initial_Stage_Success_Submit_Rate",
        "Initial_Stage_Cancel_Submit_Rate",
        "Initial_Stage_In_Progress_Submit_Rate",
        "Initial_Stage_WootCancel_QPO_Count",
        "Initial_Stage_SellerCancel_QPO_Count",
        "Initial_Stage_WootCancel_Stage_Rate",
        "Initial_Stage_WootCancel_Cancel_Rate",
        "Initial_Stage_SellerCancel_Stage_Rate",
        "Initial_Stage_SellerCancel_Cancel_Rate",
        "Second_Stage_QPO_Count",
        "Second_Stage_Success_QPO_Count",
        "Second_Stage_Cancel_QPO_Count",
        "Second_Stage_In_Progress_QPO_Count",
        "Second_Stage_Success_Submit_Rate",
        "Second_Stage_Cancel_Submit_Rate",
        "Second_Stage_In_Progress_Submit_Rate",
        "Second_Stage_WootCancel_QPO_Count",
        "Second_Stage_SellerCancel_QPO_Count",
        "Second_Stage_WootCancel_Stage_Rate",
        "Second_Stage_WootCancel_Cancel_Rate",
        "Second_Stage_SellerCancel_Stage_Rate",
        "Second_Stage_SellerCancel_Cancel_Rate",
        "Third_Stage_QPO_Count",
        "Third_Stage_Success_QPO_Count",
        "Third_Stage_Cancel_QPO_Count",
        "Third_Stage_In_Progress_QPO_Count",
        "Third_Stage_Success_Submit_Rate",
        "Third_Stage_Cancel_Submit_Rate",
        "Third_Stage_In_Progress_Submit_Rate",
        "Third_Stage_WootCancel_QPO_Count",
        "Third_Stage_SellerCancel_QPO_Count",
        "Third_Stage_WootCancel_Stage_Rate",
        "Third_Stage_WootCancel_Cancel_Rate",
        "Third_Stage_SellerCancel_Stage_Rate",
        "Third_Stage_SellerCancel_Cancel_Rate",
    ]
    columns_cn = [
        "月份",
        "报价方式",
        "提报QPO数",
        "成交QPO数",
        "跟进中QPO数",
        "取消QPO数",
        "成交率(占提报)",
        "跟进中占比(占提报)",
        "取消率(占提报)",
        "初审阶段QPO数",
        "初审阶段成交QPO数",
        "初审阶段取消QPO数",
        "初审阶段跟进中QPO数",
        "初审阶段成交占提报",
        "初审阶段取消占提报",
        "初审阶段跟进中占提报",
        "初审阶段Woot取消QPO数",
        "初审阶段卖家取消QPO数",
        "初审阶段Woot取消占阶段",
        "初审阶段Woot取消占取消",
        "初审阶段卖家取消占阶段",
        "初审阶段卖家取消占取消",
        "二审阶段QPO数",
        "二审阶段成交QPO数",
        "二审阶段取消QPO数",
        "二审阶段跟进中QPO数",
        "二审阶段成交占提报",
        "二审阶段取消占提报",
        "二审阶段跟进中占提报",
        "二审阶段Woot取消QPO数",
        "二审阶段卖家取消QPO数",
        "二审阶段Woot取消占阶段",
        "二审阶段Woot取消占取消",
        "二审阶段卖家取消占阶段",
        "二审阶段卖家取消占取消",
        "三审阶段QPO数",
        "三审阶段成交QPO数",
        "三审阶段取消QPO数",
        "三审阶段跟进中QPO数",
        "三审阶段成交占提报",
        "三审阶段取消占提报",
        "三审阶段跟进中占提报",
        "三审阶段Woot取消QPO数",
        "三审阶段卖家取消QPO数",
        "三审阶段Woot取消占阶段",
        "三审阶段Woot取消占取消",
        "三审阶段卖家取消占阶段",
        "三审阶段卖家取消占取消",
    ]

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "UK207_PricingMode_Funnel"
    sheet.append(columns_en)
    sheet.append(columns_cn)
    for row in rows:
        sheet.append([row.get(column) for column in columns_en])
    style_two_header_sheet(sheet, columns_en, columns_cn)

    notes = workbook.create_sheet("Notes")
    for note_row in build_notes_rows(metadata, pricing_modes):
        notes.append(list(note_row))
    for col in "ABCDEF":
        notes.column_dimensions[col].width = 28
    notes.freeze_panes = "A3"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def main() -> None:
    args = parse_args()

    input_path = Path(args.input_path) if args.input_path else find_latest_input(BASE_DIR)
    if not input_path.exists():
        raise FileNotFoundError(f"未找到输入文件: {input_path}")

    output_path = (
        Path(args.output_path)
        if args.output_path
        else build_default_output_path(EXPORT_DIR, DEFAULT_OUTPUT_NAME, BASE_DIR)
    )

    records, metadata = load_stage_qpo_records(input_path)
    validate_metadata(metadata)
    rows, pricing_modes = aggregate(records)
    write_workbook(rows, output_path, metadata, pricing_modes)

    print("生成完成：")
    print(output_path)


if __name__ == "__main__":
    main()
