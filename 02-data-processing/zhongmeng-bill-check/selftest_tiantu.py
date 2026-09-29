# -*- coding: utf-8 -*-
"""天图自检：离线契约检查 + 可选联网双跑一致性。"""
from __future__ import annotations
import csv,json,sys
from pathlib import Path
HERE=Path(__file__).resolve().parent
sys.path.insert(0,str(HERE.parent.parent/"08-common-utils"))
KNOWN={"偏远","超偏远","超长","超重","住宅私人","address","pod"}
def _engine(config):
    p=Path((config.get("tiantu") or {}).get("engine", "")); return p if p.is_absolute() else (HERE/p).resolve()
def check(config,checklist):
    problems=[]; p=Path(checklist)
    if not p.is_file(): return [f"清单不存在：{p}"]
    with p.open(encoding="utf-8-sig",newline="") as f: rows=list(csv.DictReader(f)); fields=rows[0].keys() if rows else []
    required={"客户单号(去后缀)","费用名称","期望标记"}; problems += [f"缺少字段：{x}" for x in required-set(fields)]
    problems += [f"未知天图标记：{x}" for x in sorted({r.get('期望标记','') for r in rows}-KNOWN)]
    engine=_engine(config)
    if not engine.is_file(): problems.append(f"引擎不存在：{engine}")
    cred=engine.parent/"credentials.local.json"
    if not cred.is_file(): problems.append(f"天图凭据不存在：{cred}")
    else:
        try:
            d=json.loads(cred.read_text(encoding="utf-8"))
            if not (d.get("tiantu_user") or d.get("user")): problems.append("天图凭据缺少账号")
            if not (d.get("tiantu_pass") or d.get("pass")): problems.append("天图凭据缺少密码")
        except Exception as exc: problems.append(f"天图凭据无法读取：{exc}")
    return problems
def _read_result(path):
    with Path(path).open(encoding="utf-8-sig",newline="") as f:
        rows=list(csv.DictReader(f))
    return [(r.get("单号") or r.get("客户单号(去后缀)"),r.get("期望标记"),r.get("是否出现")) for r in rows]
def online_double_run(config,checklist,limit=6):
    from autolib.tiantu import run
    tc=dict(config.get("tiantu",{})); tc["engine"]=str(_engine(config))
    first=run(checklist,limit=limit,config=tc); a=_read_result(first)
    second=run(checklist,limit=limit,config=tc); b=_read_result(second)
    flips=[(x,y) for x,y in zip(a,b) if x!=y]
    if len(a)!=len(b): flips.append((f"第一次{len(a)}条",f"第二次{len(b)}条"))
    return first,second,flips
def main(argv=None):
    import argparse
    ap=argparse.ArgumentParser(); ap.add_argument("checklist"); ap.add_argument("--config",default=str(HERE/"config.json")); ap.add_argument("--online",action="store_true"); ap.add_argument("--limit",type=int,default=6); a=ap.parse_args(argv)
    c=json.loads(Path(a.config).read_text(encoding="utf-8")); problems=check(c,a.checklist)
    if problems:
        print("\n".join("[失败] "+x for x in problems)); return 1
    print("天图契约自检通过")
    if a.online:
        first,second,flips=online_double_run(c,a.checklist,a.limit); print(f"第一次：{first}\n第二次：{second}\n翻转：{len(flips)}")
        for x,y in flips: print("  ",x,"=>",y)
        return 1 if flips else 0
    return 0
if __name__=="__main__": raise SystemExit(main())
