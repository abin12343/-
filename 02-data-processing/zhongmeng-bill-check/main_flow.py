# -*- coding: utf-8 -*-
"""中盟总流程：批量加工账单、天图核验、问题统计、运行记录和企微通知。"""
from __future__ import annotations
import argparse, csv, json, sys, time
from pathlib import Path

# 开发环境取源码目录；PyInstaller 环境中 __file__ 位于 _internal，
# 配置、输出和随包引擎都位于发布根目录（exe 所在目录的上一级）。
if getattr(sys, "frozen", False):
    HERE=Path(sys.executable).resolve().parent.parent
    for _stream in (sys.stdout, sys.stderr):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass
else:
    HERE=Path(__file__).resolve().parent
sys.path.insert(0,str(HERE)); sys.path.insert(0,str(HERE.parent.parent/"08-common-utils"))
from process import load_zone_index, process, process_data_list, write_checklist, write_report

def _merge(base,over):
    out=dict(base)
    for k,v in (over or {}).items(): out[k]=_merge(out[k],v) if isinstance(v,dict) and isinstance(out.get(k),dict) else v
    return out
def cfg_load(path=None):
    p=Path(path or HERE/"config.json"); cfg=json.loads(p.read_text(encoding="utf-8")); local=p.with_name("config.local.json")
    if local.exists(): cfg=_merge(cfg,json.loads(local.read_text(encoding="utf-8")))
    notify_local=p.with_name("notify.local.json")
    if notify_local.exists(): cfg["notify"]=_merge(cfg.get("notify",{}),json.loads(notify_local.read_text(encoding="utf-8")))
    return cfg
def _engine_cfg(cfg):
    tc=dict(cfg.get("tiantu",{})); p=Path(tc.get("engine", "")); tc["engine"]=str((HERE/p).resolve()) if not p.is_absolute() else str(p); return tc
def _append_stats_items(stats_file,sheet,items,log=print):
    if not stats_file or not Path(stats_file).is_file() or not items: return 0
    from openpyxl import load_workbook
    wb=load_workbook(stats_file); ws=wb[sheet] if sheet in wb.sheetnames else wb.create_sheet(sheet)
    from collections import Counter
    existing=Counter((str(r[0] or ""),str(r[1] or ""),str(r[2] or "")) for r in ws.iter_rows(min_row=2,values_only=True)); seen=Counter(); n=0
    for x in items:
        item=(str(x.get("fee", "")),str(x.get("amount", "")),str(x.get("no", "")))
        seen[item]+=1
        if seen[item]>existing[item]: ws.append(item); n+=1
    wb.save(stats_file); log(f"[统计表] 写入 {n} 条"); return n
def missing_to_stats(stats_file,sheet,result_csv,checklist=None,log=print):
    """兼容现有天图结果字段；按中盟模板列序写入并去重。"""
    if not stats_file or not Path(stats_file).is_file() or not result_csv or not Path(result_csv).is_file(): return 0
    amounts={}
    if checklist and Path(checklist).is_file():
        with Path(checklist).open(encoding="utf-8-sig",newline="") as f:
            for x in csv.DictReader(f): amounts[(x.get("客户单号(去后缀)",""),x.get("费用名称",""),x.get("期望标记",""))]=x.get("金额USD","")
    items=[]
    with Path(result_csv).open(encoding="utf-8-sig",newline="") as f:
        for row in csv.DictReader(f):
            if str(row.get("是否出现","")).strip()=="是": continue
            no=row.get("单号") or row.get("客户单号(去后缀)") or ""; fee=row.get("费用名称",""); mark=row.get("期望标记","")
            items.append({"no":no,"fee":fee,"amount":row.get("金额USD") or amounts.get((no,fee,mark),"")})
    return _append_stats_items(stats_file,sheet,items,log)
def _bills(input_dir,explicit):
    if explicit: return [Path(x) for x in explicit]
    root=Path(input_dir)
    # 中盟实际账单通常放在 input_dir/中盟账单；递归查找仍保留同一批次内的文件名规则。
    return sorted(p for p in root.rglob("US*.xlsx") if p.is_file() and "原始备份" not in p.name)

def _pick_folder(initial):
    """按 SOP 弹出文件夹选择；无图形环境时返回空值，由调用方继续使用配置目录。"""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root=tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
        selected=filedialog.askdirectory(initialdir=str(initial), title="选择中盟账单文件夹")
        root.destroy()
        return Path(selected) if selected else None
    except Exception:
        return None
def _pick_data_list(initial):
    """按 SOP 由使用者选择本批次对应的数据列表文件。"""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root=tk.Tk(); root.withdraw(); root.attributes("-topmost", True)
        root.update()
        selected=filedialog.askopenfilename(
            initialdir=str(initial), title="选择中盟数据列表文件",
            filetypes=(("Excel 文件", "*.xlsx"), ("所有文件", "*.*")),
            parent=root,
        )
        root.destroy()
        return Path(selected) if selected else None
    except Exception:
        return None


def _pick_quote(initial):
    """报价表属于每台电脑/每个批次的输入，配置路径失效时重新选择。"""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root=tk.Tk(); root.withdraw(); root.attributes("-topmost", True); root.update()
        selected=filedialog.askopenfilename(
            initialdir=str(initial), title="请选择中盟尾程报价表 xlsx",
            filetypes=(("Excel 文件", "*.xlsx"), ("所有文件", "*.*")), parent=root,
        )
        root.destroy()
        return Path(selected) if selected else None
    except Exception:
        return None
def main(argv=None):
    ap=argparse.ArgumentParser(description="中盟账单数据整理")
    ap.add_argument("--bill",action="append",help="账单，可重复传入；不传则递归批量处理 input_dir 下的 US*.xlsx")
    ap.add_argument("--input-dir",help="账单文件夹；覆盖配置中的 input_dir")
    ap.add_argument("--pick-folder",action="store_true",help="弹窗选择账单文件夹（默认行为，保留此参数用于兼容）")
    ap.add_argument("--no-dialog",action="store_true",help="不弹窗，直接使用 --input-dir 或配置目录")
    ap.add_argument("--quote"); ap.add_argument("--data-list",help="数据列表 xlsx（C列运单、E列分区）")
    ap.add_argument("--stats"); ap.add_argument("--config"); ap.add_argument("--dry-run",action="store_true")
    ap.add_argument("--skip-tiantu",action="store_true"); ap.add_argument("--selftest",action="store_true"); ap.add_argument("--tiantu-limit",type=int,default=0)
    args=ap.parse_args(argv); cfg=cfg_load(args.config); paths=cfg.get("paths",{}); input_dir=Path(args.input_dir or paths.get("input_dir","."))
    # SOP：系统导出步骤无需脚本操作；正常启动时只让使用者选择装有全部账单的文件夹。
    should_pick=args.pick_folder or (not args.no_dialog and not args.bill and not args.input_dir)
    if should_pick:
        picked=_pick_folder(input_dir)
        if not picked:
            print("[取消] 未选择中盟账单文件夹，程序已退出。")
            return 2
        input_dir=picked
    bills=_bills(input_dir,args.bill)
    print(f"[选择] 已选择账单文件夹：{input_dir}；发现账单 {len(bills)} 份")
    quote=Path(args.quote or paths.get("quote_file", ""))
    if not quote.is_file() and not args.no_dialog and not args.quote:
        print(f"[提示] 配置中的报价表不存在：{quote}")
        print("[下一步] 请在随后弹出的窗口中选择本批次的中盟尾程报价表.xlsx")
        picked_quote=_pick_quote(input_dir.parent if input_dir.name == "中盟账单" else input_dir)
        if picked_quote: quote=picked_quote
    out=Path(paths.get("output_dir",HERE/"outputs")); out.mkdir(parents=True,exist_ok=True)
    logdir=Path(paths.get("log_dir",HERE/"logs")); logdir.mkdir(parents=True,exist_ok=True); runlog=logdir/f"zhongmeng_{time.strftime('%Y%m%d_%H%M%S')}.log"; start=time.time()
    def emit(msg):
        print(msg)
        try: runlog.open("a",encoding="utf-8").write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {msg}\n")
        except OSError: pass
    if not bills or not quote.is_file(): emit(f"[失败] 账单数={len(bills)}，报价表={quote}"); return 2
    data_list=args.data_list
    if not data_list and not args.no_dialog:
        print("[下一步] 请在随后弹出的窗口中选择本批次的数据列表.xlsx")
        picked=_pick_data_list(input_dir.parent if input_dir.name == "中盟账单" else input_dir)
        if picked: data_list=str(picked)
        else:
            print("[取消] 未选择中盟数据列表文件，程序已退出。")
            return 2
    if not data_list:
        # 无界面自动化时仍支持从所选文件夹或其上级目录自动识别。
        found=next((p for p in input_dir.rglob("*数据列表*.xlsx") if not p.name.startswith("~$")),None)
        if not found and input_dir.parent != input_dir:
            found=next((p for p in input_dir.parent.glob("*数据列表*.xlsx") if not p.name.startswith("~$")),None)
        data_list=str(found) if found else None
    zone_index=load_zone_index(data_list) if data_list else {}; all_tasks=[]; force_items=[]; totals={}; stages=[]
    if not zone_index:
        emit("[失败] 未找到可用的数据列表（需 C 列客户单号、E 列分区）。请将数据列表放入所选账单文件夹或其上级目录。")
        return 2
    emit(f"[数据列表] 已加载 {Path(data_list).name}，分区索引 {len(zone_index)} 条")
    try:
        t=time.time()
        for bill in bills:
            if not bill.is_file(): raise FileNotFoundError(bill)
            stat,tasks,_=process(bill,quote,cfg,log=emit,dry_run=args.dry_run,zone_index=zone_index)
            force_items.extend(stat.get("强制登记",[]))
            for k,v in stat.items():
                if isinstance(v,(int,float)): totals[k]=totals.get(k,0)+v
            for task in tasks: task["note"]=f"{bill.name} / {task['note']}"
            all_tasks.extend(tasks); emit(f"[加工] {bill.name}：{stat}")
        if data_list:
            data_count=process_data_list(data_list,quote,dry_run=args.dry_run)
            emit(f"[数据列表] 运费2处理 {data_count} 行")
        stages.append({"name":"加工报价","secs":round(time.time()-t,2),"ok":True})
        checklist=write_checklist(all_tasks,out) if all_tasks else None; result=None
        if checklist and args.selftest:
            from selftest_tiantu import check
            cc=dict(cfg); cc["tiantu"]=_engine_cfg(cfg); problems=check(cc,checklist)
            if problems:
                for p in problems: emit(f"[天图自检失败] {p}")
                return 1
            emit("[天图自检] 契约检查通过")
        if checklist and not args.skip_tiantu:
            emit("[天图] 开始核验；若未配置凭据，将等待浏览器手动登录")
            from autolib.tiantu import run
            result=run(checklist,out/"天图核验结果.csv",args.tiantu_limit,config=_engine_cfg(cfg))
        stats_path=args.stats or paths.get("stats_file"); stats_sheet=paths.get("stats_sheet","中盟")
        if not args.dry_run and force_items: _append_stats_items(stats_path,stats_sheet,force_items,emit)
        if result and not args.dry_run: missing_to_stats(stats_path,stats_sheet,result,checklist,emit)
        report=write_report(out/f"中盟运行报告_{time.strftime('%Y%m%d_%H%M%S')}.txt",totals,all_tasks,checklist); emit(f"[报告] {report}")
        try:
            from autolib.notify import build_summary,notify_all
            msg=build_summary(cfg.get("notify",{}).get("app_name","中盟账单数据整理"),"成功",start,stages,extra_lines=[f"账单 {len(bills)} 份；已报价 {totals.get('已报价',0)} 行；未匹配 {totals.get('未匹配报价',0)} 行；天图任务 {len(all_tasks)} 条"])
            notify_all(cfg,msg,"成功",start_ts=start,stages=stages)
        except Exception as exc: emit(f"[通知] 跳过：{exc}")
        return 0
    except Exception as exc:
        emit(f"[失败] {exc}")
        try:
            from autolib.notify import notify_all; notify_all(cfg,f"中盟账单数据整理\n状态：失败\n原因：{exc}","失败",start_ts=start)
        except Exception: pass
        return 1
if __name__=="__main__": raise SystemExit(main())
