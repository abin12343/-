# -*- coding: utf-8 -*-
"""
【模板】文件批量处理脚本（整理 / 重命名 / 移动 / 清理）

用途：<一句话说明这个脚本做什么>
依赖：Python 3 标准库
运行：python template_file_batch.py --dir <目录> [--apply]

安全设计（务必保留）：
  - 默认只扫描并打印预览，**不做任何改动**；必须显式加 --apply 才执行
  - --apply 前会弹窗二次确认
  - 移动/重命名前先备份到 <目录>/_备份_<时间戳>/
  - 每一步都记日志，出问题可回溯

怎么用这个模板：
  1. 复制到 01-file-management/，改个业务化的名字
  2. 改 plan_actions() —— 只返回"计划动作"，不真的执行；执行由骨架统一负责
  3. 永远先不加 --apply 跑一遍，看预览确认无误
"""

from __future__ import annotations

import argparse
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from autolib import config as cfgmod  # noqa: E402
from autolib import dialog, log, notify, paths  # noqa: E402

APP_NAME = "文件批处理模板"

# 危险目录黑名单：命中直接拒绝执行，防止误伤
FORBIDDEN = {
    "C:\\", "C:\\Windows", "C:\\Program Files", "C:\\Program Files (x86)",
    "C:\\Users", str(Path.home()),
}


# ============ 只需要改这里 ============

def plan_actions(files: list[Path], context: dict) -> list[dict]:
    """根据扫描到的文件，返回计划动作列表。本函数**不得修改任何文件**。

    支持的动作：
      {"kind": "rename", "src": Path, "dst": Path}
      {"kind": "move",   "src": Path, "dst": Path}
      {"kind": "copy",   "src": Path, "dst": Path}
      {"kind": "delete", "src": Path}
    """
    actions = []
    for f in files:
        # 示例：把文件名里的空格替换成下划线
        new_name = f.name.replace(" ", "_")
        if new_name != f.name:
            actions.append({"kind": "rename", "src": f, "dst": f.with_name(new_name)})
    return actions


# ============ 以下通常不用动 ============


def parse_args():
    ap = argparse.ArgumentParser(description=APP_NAME)
    ap.add_argument("--dir", default=None, help="目标目录；不给则弹窗选择")
    ap.add_argument("--config", default=None)
    ap.add_argument("--recursive", action="store_true", help="递归子目录")
    ap.add_argument("--pattern", default="*", help="文件通配符，如 *.jpg")
    ap.add_argument("--apply", action="store_true", help="真正执行（不加则只预览）")
    ap.add_argument("--no-backup", action="store_true", help="不备份（不建议）")
    ap.add_argument("--max-files", type=int, default=5000, help="最多处理多少个文件")
    ap.add_argument("--no-notify", action="store_true")
    return ap.parse_args()


def scan(directory: Path, pattern: str, recursive: bool, limit: int) -> list[Path]:
    it = directory.rglob(pattern) if recursive else directory.glob(pattern)
    return sorted([p for p in it if p.is_file()])[:limit]


def is_forbidden(directory: Path) -> bool:
    try:
        resolved = str(directory.resolve())
    except Exception:  # noqa: BLE001
        return True
    return any(resolved.lower().startswith(f.lower()) for f in FORBIDDEN)


def execute(actions: list[dict], backup_dir: Path | None, logger) -> tuple[int, int]:
    done = failed = 0
    for act in actions:
        src: Path = act["src"]
        dst: Path | None = act.get("dst")
        try:
            if backup_dir and act["kind"] != "copy" and src.exists():
                shutil.copy2(src, backup_dir / src.name)
            if act["kind"] == "rename":
                src.rename(dst)
            elif act["kind"] == "move":
                dst.parent.mkdir(parents=True, exist_ok=True)
                shutil.move(str(src), str(dst))
            elif act["kind"] == "copy":
                dst.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src, dst)
            elif act["kind"] == "delete":
                src.unlink()
            done += 1
        except Exception as exc:  # noqa: BLE001
            failed += 1
            logger.warning(f"失败 {act['kind']} {src}: {exc}")
    return done, failed


def main() -> int:
    args = parse_args()
    conf = cfgmod.load_config(args.config)
    logger = log.setup_logger(Path(__file__).stem, conf)
    timer = log.Timer()
    start_ts = time.time()
    status, note = "成功", ""

    directory = Path(args.dir) if args.dir else None
    if directory is None:
        picked = dialog.ask_directory("请选择要处理的目录")
        if not picked:
            logger.warning("未选择目录，流程取消。")
            return 1
        directory = Path(picked)

    if not directory.is_dir():
        logger.error(f"目录不存在: {directory}")
        return 2
    if is_forbidden(directory):
        logger.error(f"拒绝在受保护目录执行: {directory}（请指定更具体的子目录）")
        return 2

    files = scan(directory, args.pattern, args.recursive, args.max_files)
    logger.info(f"扫描 {directory} -> {len(files)} 个文件")

    actions = plan_actions(files, {"config": conf, "root": directory})
    logger.info(f"生成 {len(actions)} 个计划动作")

    if not actions:
        logger.info("没有需要处理的动作，结束。")
        return 0

    # 预览
    print(f"\n{'=' * 60}")
    print(f"预览：共 {len(actions)} 个动作（{'执行模式' if args.apply else '预览模式，未做任何改动'}）")
    print("=" * 60)
    for act in actions[:30]:
        arrow = "->" if act.get("dst") else "(删除)"
        print(f"  [{act['kind']}] {act['src'].name} {arrow} "
              f"{act['dst'].name if act.get('dst') else ''}")
    if len(actions) > 30:
        print(f"  ... 另有 {len(actions) - 30} 个")

    if not args.apply:
        print("\n确认无误后加 --apply 执行。")
        return 0

    # 二次确认
    if not dialog.confirm(
        f"即将对 {len(actions)} 个文件执行操作（目录：{directory}）。\n"
        f"操作前会自动备份到 _备份_<时间戳>/ 目录。是否继续？"
    ):
        logger.info("用户取消，未执行任何操作。")
        return 0

    backup_dir = None
    if not args.no_backup:
        backup_dir = directory / f"_备份_{datetime.now():%Y%m%d_%H%M%S}"
        backup_dir.mkdir(parents=True, exist_ok=True)
        logger.info(f"备份目录: {backup_dir}")

    done, failed = execute(actions, backup_dir, logger)
    logger.info(f"执行完成：成功 {done}，失败 {failed}，耗时 {timer.elapsed}s")
    if failed:
        status, note = "失败", f"{failed} 个动作失败"
    if backup_dir:
        logger.info(f"原始文件已备份于: {backup_dir}")

    if not args.no_notify:
        notify.notify_all(
            conf,
            notify.build_summary(
                APP_NAME, status, start_ts, note=note,
                extra_lines=[f"成功 {done} / 失败 {failed}"],
            ),
            status=status, start_ts=start_ts,
        )
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:  # noqa: BLE001
        print(f"[异常] {exc}")
        raise
