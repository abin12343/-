# -*- coding: utf-8 -*-
"""配置加载。

规则：
  1. 默认读取脚本同目录的 config.json
  2. 若存在 config.local.json（本地私密配置，已 gitignore），逐键覆盖
  3. 环境变量优先级最高：cfg 中形如 "${VAR_NAME}" 的字符串会被替换

这样换机器只需要改 local 文件或设环境变量，不动代码。
"""

from __future__ import annotations

import json
import os
import re
from pathlib import Path

from . import paths

_ENV_PATTERN = re.compile(r"^\$\{([A-Za-z_][A-Za-z0-9_]*)\}$")


def _expand(value, default=None):
    """把 "${VAR}" 展开为环境变量；未设置时返回 default。"""
    if isinstance(value, str):
        m = _ENV_PATTERN.match(value.strip())
        if m:
            return os.environ.get(m.group(1), default)
        return value
    if isinstance(value, dict):
        return {k: _expand(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_expand(v) for v in value]
    return value


def _deep_merge(base: dict, override: dict) -> dict:
    out = dict(base)
    for k, v in (override or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def load_config(path=None, base_dir=None, local=True, expand_env=True) -> dict:
    """加载配置。

    :param path: 主配置文件路径；缺省取 base_dir/config.json
    :param base_dir: 基准目录；缺省为脚本所在目录
    :param local: 是否用 config.local.json 覆盖
    :param expand_env: 是否展开 "${VAR}" 占位符
    :return: 合并后的配置字典；文件不存在时返回 {}
    """
    bd = Path(base_dir) if base_dir else paths.base_dir()
    main_path = Path(path) if path else bd / "config.json"
    cfg: dict = {}
    if main_path.exists():
        try:
            cfg = json.loads(main_path.read_text(encoding="utf-8"))
        except Exception as exc:  # noqa: BLE001
            raise ValueError(f"配置文件解析失败 {main_path}: {exc}") from exc

    if local:
        local_path = main_path.with_name(
            main_path.stem + ".local" + main_path.suffix
        )
        if local_path.exists():
            try:
                cfg = _deep_merge(cfg, json.loads(local_path.read_text(encoding="utf-8")))
            except Exception as exc:  # noqa: BLE001
                raise ValueError(f"本地配置文件解析失败 {local_path}: {exc}") from exc

    return _expand(cfg) if expand_env else cfg


def get(cfg: dict, dotted: str, default=None):
    """按点号路径取值，如 get(cfg, "paths.output_dir")。"""
    cur = cfg
    for key in dotted.split("."):
        if not isinstance(cur, dict) or key not in cur:
            return default
        cur = cur[key]
    return cur


def require(cfg: dict, dotted: str):
    """按点号路径取值，缺失时抛错（用于必须配置项的早失败）。"""
    val = get(cfg, dotted)
    if val in (None, "", [], {}):
        raise KeyError(f"缺少必需配置项: {dotted}（请检查 config.json / config.local.json）")
    return val


def load_credentials(cfg: dict, prefix: str, env_user: str = "", env_pass: str = ""):
    """读取账号密码：环境变量 > config.credentials > <prefix> 字段。

    返回 (user, password)，取不到时为 ("", "")。
    注意：不要把真实密码写进 config.json，用 config.local.json 或环境变量。
    """
    if env_user and os.environ.get(env_user):
        return os.environ.get(env_user, ""), os.environ.get(env_pass, "")
    creds = get(cfg, "credentials", {}) or {}
    user = creds.get(f"{prefix}_user") or creds.get("user") or ""
    pwd = creds.get(f"{prefix}_pass") or creds.get("pass") or ""
    return user, pwd
