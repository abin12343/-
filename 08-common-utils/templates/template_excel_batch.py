# -*- coding: utf-8 -*-
"""
【模板】Excel / 表格批处理脚本

用途：<一句话说明这个脚本做什么>
依赖：Python 3 + openpyxl
运行：python template_excel_batch.py [--input 文件或目录] [--dry-run]

怎么用这个模板：
  1. 复制本文件到对应功能目录（如 02-data-processing/），改个业务化的名字
  2. 改 process_row() —— 那是唯一需要你写业务逻辑的地方
  3. 复制 config.example.json 为 config.json，按需调整列别名
  4. 命令行先跑 --dry-run 看预览，确认无误再正式执行

骨架已经处理好的事（不用重复写）：
  - 弹窗选文件 + 选错自动校验重弹
  - 配置外置、日志统一写 logs/
  - 输出文件名带时间戳，不会覆盖历史
  - --dry-run 只预览不落盘
  - 保存时文件被 Excel 占用自动降级
  - 结束推送企业微信 / 写金山文档运行记录（失败不影响结果）
"""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from autolib import config as cfgmod  # noqa: E402
from autolib import dialog, excel, log, notify, paths  # noqa: E402

APP_NAME = "Excel 批处理模板"


# ============ 只需要改这里 ============

def column_aliases() -> dict:
    """列别名 -> 逻辑名。表格加列/改列名都不会错位。"""
    return {
        "no": ["客户单号", "系统单号", "单号"],
        "fee": ["费用名称", "翻译", "项目"],
        "amount": ["费用金额", "金额", "金额USD"],
    }


def process_row(row: dict, context: dict) -> dict | None:
    """处理一行数据。

    :param row: 该行的 dict，键是 column_aliases() 里的逻辑名
    :param context: 全局上下文（配置、计数器等），可用来累计统计
    :return: 有问题的行返回 {"no":..., "fee":..., "amount":..., "remark":"说明"}，
             正常行返回 None
    """
    amount = row.get("amount")
    if amount is None:
        context["skipped"] += 1
        return None
    # 示例规则：金额大于阈值的算问题，按你的业务改写
    threshold = context["config"].get("run", {}).get("amount_threshold", 100)
    try:
        if float(amount) > float(threshold):
            context["problems"] += 1
            return {
                "no": row.get("no"),
                "fee": row.get("fee"),
                "amount": amount,
                "remark": f"金额 {amount} 超过阈值 {threshold}",
            }
    except (TypeError, ValueError):
        context["skipped"] += 1
    return None


# ============ 以下通常不用动 ============


def parse_args():
    ap = argparse.ArgumentParser(description=APP_NAME)
    ap.add_argument("--input", default=None, help="输入 xlsx；不给则弹窗选择")
    ap.add_argument("--config", default=None, help="配置文件路径；缺省取脚本同目录 config.json")
    ap.add_argument("--sheet", default=None, help="工作表名；缺省第一个 sheet")
    ap.add_argument("--dry-run", action="store_true", help="只预览，不写文件")
    ap.add_argument("--no-notify", action="store_true", help="不发送通知")
    return ap.parse_args()


def main() -> int:
    args = parse_args()
    conf = cfgmod.load_config(args.config)
    logger = log.setup_logger(Path(__file__).stem, conf)
    timer = log.Timer()
    start_ts = time.time()
    status, note = "成功", ""

    logger.info("=" * 60)
    logger.info(f"{APP_NAME} 开始" + ("（dry-run 预览）" if args.dry_run else ""))

    # 1) 选输入文件，选错自动重弹
    src = args.input
    if not src:
        initial = cfgmod.get(conf, "paths.input_dir")
        src = dialog.pick_until_valid(
            title="请选择要处理的 Excel 文件",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
            initialdir=initial,
            validator=excel.validate_bill_like,
        )
    if not src:
        logger.warning("未选择文件，流程取消。")
        return 1
    src = Path(src)
    logger.info(f"输入文件: {src}")

    # 2) 读表
    openpyxl = excel._require_openpyxl()
    wb = openpyxl.load_workbook(src, data_only=True)
    ws = wb[args.sheet] if args.sheet and args.sheet in wb.sheetnames else wb[wb.sheetnames[0]]
    rows = excel.read_rows(ws, header_row=1, aliases=column_aliases())
    logger.info(f"读取 {len(rows)} 行，工作表: {ws.title}")

    # 3) 逐行处理
    context = {"config": conf, "problems": 0, "skipped": 0}
    problems = []
    for row in rows:
        item = process_row(row, context)
        if item:
            problems.append(item)

    logger.info(f"处理完成：问题 {len(problems)} 条，跳过 {context['skipped']} 行")

    # 4) 输出
    out_dir = paths.output_dir(conf)
    paths.ensure_dir(out_dir)
    if args.dry_run:
        print(f"[dry-run] 共 {len(problems)} 条待写入，以下为前 20 条预览：")
        for it in problems[:20]:
            print(f"  {it.get('no')} | {it.get('fee')} | {it.get('amount')} | {it.get('remark')}")
        print("[dry-run] 未写入任何文件。")
    else:
        out_xlsx = out_dir / paths.stamped_name(src.stem, "问题清单", "xlsx")
        out_wb = openpyxl.Workbook()
        out_ws = out_wb.active
        out_ws.title = "问题清单"
        headers = ["单号", "费用名称", "金额", "备注"]
        out_ws.append(headers)
        for it in problems:
            out_ws.append([it.get("no"), it.get("fee"), it.get("amount"), it.get("remark")])
        saved, degraded = excel.safe_save(out_wb, out_xlsx)
        logger.info(f"已输出: {saved}" + ("（原文件被占用，已改名保存）" if degraded else ""))

        out_txt = out_dir / paths.stamped_name(src.stem, "报告", "txt")
        out_txt.write_text(
            "\n".join([
                f"输入: {src}",
                f"总行数: {len(rows)}",
                f"问题数: {len(problems)}",
                f"跳过数: {context['skipped']}",
                f"耗时: {timer.elapsed}s",
            ]),
            encoding="utf-8",
        )
        logger.info(f"已输出: {out_txt}")

    logger.info(f"{APP_NAME} 结束，耗时 {timer.elapsed}s")

    if not args.no_notify:
        notify.notify_all(
            conf,
            notify.build_summary(
                conf.get("notify", {}).get("app_name", APP_NAME),
                status, start_ts, note=note,
                extra_lines=[f"问题 {len(problems)} 条 / 共 {len(rows)} 行"],
            ),
            status=status, start_ts=start_ts,
        )
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:  # noqa: BLE001
        print(f"[异常] {exc}")
        raise
