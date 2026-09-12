#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
build_review_html.py —— 把内容 JSON 注入模板，产出单文件可交互 HTML 复习页。

用法:
    python build_review_html.py --content review_content.json --out 期末复习.html
                                [--template ../assets/template.html] [--title "..."]

行为:
    1. 读取并校验内容 JSON（缺字段时报明确错误，并给出可修复提示）
    2. 自动补齐 meta/科目颜色/icon/缺失的小节
    3. 把 JSON 注入模板占位符 __REVIEW_DATA__，输出单个自包含 HTML（无外网依赖）
    4. 打印内容体检报告：每科目要点数/速记数/题目数/易错点数，并对过瘦科目告警
"""

from __future__ import annotations

import argparse
import json
import os
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:  # pragma: no cover
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TEMPLATE = os.path.join(HERE, "..", "assets", "template.html")
PLACEHOLDER = "__REVIEW_DATA__"

PALETTE = [
    "#4f6ef7", "#0ea5a4", "#f97316", "#a855f7",
    "#ef4444", "#0891b2", "#65a30d", "#d946a0",
]
LEVELS = {"core", "key", "basic"}
QUIZ_TYPES = {"单选", "多选", "判断", "填空", "名词解释", "简答", "计算", "论述", "案例分析"}


def fail(msg: str) -> "None":
    print(f"[内容校验失败] {msg}")
    sys.exit(1)


def validate(data: dict) -> dict:
    if not isinstance(data, dict):
        fail("根节点必须是 JSON 对象")
    subjects = data.get("subjects")
    if not isinstance(subjects, list) or not subjects:
        fail("缺少 subjects 数组，或为空。至少需要一个科目。")

    meta = data.setdefault("meta", {})
    meta.setdefault("title", "期末复习冲刺包")
    meta.setdefault("subtitle", "知识要点 · 速背速记 · 模拟题解析")
    meta.setdefault("generated_at", "")
    meta.setdefault("note", "")

    used_ids = set()
    for i, s in enumerate(subjects):
        if not isinstance(s, dict):
            fail(f"subjects[{i}] 必须是对象")
        s.setdefault("id", f"subject-{i + 1}")
        if s["id"] in used_ids:
            s["id"] = f"{s['id']}-{i + 1}"
        used_ids.add(s["id"])
        if not s.get("name"):
            fail(f"subjects[{i}] 缺少 name")
        s.setdefault("short", s["name"][:4])
        s.setdefault("icon", s["name"][:1])
        s.setdefault("color", PALETTE[i % len(PALETTE)])
        s.setdefault("overview", "")
        s.setdefault("stats", [])
        s.setdefault("key_points", [])
        s.setdefault("memorize", [])
        s.setdefault("quiz", [])
        s.setdefault("pitfalls", [])
        s.setdefault("checklist", [])

        for ci, ch in enumerate(s["key_points"]):
            ch.setdefault("chapter", f"未命名章节 {ci + 1}")
            ch.setdefault("weight", "")
            for pi, p in enumerate(ch.get("items", [])):
                if not p.get("term") and not p.get("body"):
                    fail(f"科目「{s['name']}」key_points[{ci}].items[{pi}] 缺少 term/body")
                p.setdefault("term", f"要点 {pi + 1}")
                p.setdefault("body", "")
                p.setdefault("note", "")
                p.setdefault("tags", [])
                lvl = p.get("level", "key")
                p["level"] = lvl if lvl in LEVELS else "key"
            ch.setdefault("items", [])

        for mi, m in enumerate(s["memorize"]):
            m.setdefault("type", "记忆卡")
            m.setdefault("title", f"速记 {mi + 1}")
            m.setdefault("content", "")
            m.setdefault("mask", False)

        for qi, q in enumerate(s["quiz"]):
            if not q.get("stem"):
                fail(f"科目「{s['name']}」quiz[{qi}] 缺少题干 stem")
            q.setdefault("type", "单选")
            if q["type"] not in QUIZ_TYPES:
                q["type"] = "单选"
            q.setdefault("options", [])
            q.setdefault("answer", "")
            q.setdefault("analysis", "")
            q.setdefault("level", "中")
            q.setdefault("source", "")
            if q["type"] in {"单选", "多选", "判断"} and not q["options"]:
                fail(f"科目「{s['name']}」quiz[{qi}]（{q['type']}）缺少 options，或题型标注有误")
            if not q["answer"]:
                fail(f"科目「{s['name']}」quiz[{qi}] 缺少 answer")

        for pi_, pf in enumerate(s["pitfalls"]):
            if isinstance(pf, str):
                s["pitfalls"][pi_] = {"title": pf, "detail": ""}
    return data


def report(data: dict) -> None:
    total_q = 0
    total_p = 0
    print("\n内容体检报告")
    print("-" * 62)
    for s in data["subjects"]:
        n_ch = len(s["key_points"])
        n_p = sum(len(c["items"]) for c in s["key_points"])
        n_m = len(s["memorize"])
        n_q = len(s["quiz"])
        n_f = len(s["pitfalls"])
        total_q += n_q
        total_p += n_p
        flag = ""
        if n_p < 8:
            flag += " ⚠ 要点偏少"
        if n_q < 5:
            flag += " ⚠ 题目偏少"
        if n_m < 2:
            flag += " ⚠ 速记偏少"
        print(f"  {s['name']:<10} 章节 {n_ch:>2}  要点 {n_p:>3}  速记 {n_m:>2}  题目 {n_q:>3}  易错 {n_f:>2}{flag}")
    print("-" * 62)
    print(f"  合计：{len(data['subjects'])} 个科目 / {total_p} 个知识要点 / {total_q} 道模拟题")


def main() -> int:
    ap = argparse.ArgumentParser(description="生成单文件 HTML 复习页")
    ap.add_argument("--content", required=True, help="内容 JSON 路径")
    ap.add_argument("--out", required=True, help="输出 HTML 路径")
    ap.add_argument("--template", default=DEFAULT_TEMPLATE, help="HTML 模板路径")
    ap.add_argument("--title", default=None, help="覆盖 meta.title")
    args = ap.parse_args()

    with open(args.content, "r", encoding="utf-8") as f:
        try:
            data = json.load(f)
        except json.JSONDecodeError as e:
            fail(f"{args.content} 不是合法 JSON：{e}")
    if args.title:
        data.setdefault("meta", {})["title"] = args.title
    data = validate(data)

    if not os.path.exists(args.template):
        fail(f"找不到模板文件：{args.template}")
    with open(args.template, "r", encoding="utf-8") as f:
        html = f.read()
    if PLACEHOLDER not in html:
        fail(f"模板中缺少占位符 {PLACEHOLDER}")

    payload = json.dumps(data, ensure_ascii=False)
    payload = payload.replace("</", "<\\/").replace("<!--", "<\\!--")

    html = html.replace(PLACEHOLDER, payload)

    out_dir = os.path.dirname(os.path.abspath(args.out))
    os.makedirs(out_dir, exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(html)

    report(data)
    size_kb = os.path.getsize(args.out) / 1024
    print(f"\n✅ 已生成：{os.path.abspath(args.out)}  ({size_kb:.1f} KB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
