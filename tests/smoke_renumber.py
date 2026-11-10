# -*- coding: utf-8 -*-
"""序列重排真机冒烟验收：DCOM 构造断序文档 -> AutomationRenumberSequences -> 回读断言。"""
import os, sys, time
import win32com.client as wc

EXPECT = [
    # (原文, 期望文, 说明)
    ("一、选择题",          "一、选择题",        "健康中文序列不动"),
    ("1. 题目甲",           "1. 题目甲",         "1. 健康"),
    ("2. 题目乙",           "2. 题目乙",         "2. 健康"),
    ("4. 题目丙",           "3. 题目丙",         "缺号 3 → 补齐"),
    ("（1）子题一",         "（1）子题一",       "（1）健康"),
    ("（2）子题二",         "（2）子题二",       "（2）健康"),
    ("1.5 米的绳子",        "1.5 米的绳子",      "小数不误伤"),
    ("①命题甲",             "①命题甲",           "① 健康"),
    ("②命题乙",             "②命题乙",           "② 健康"),
    ("④命题丁",             "③命题丁",           "③ 缺号 → 补齐"),
    ("2026 年的大事",       "2026 年的大事",     "年份不误伤"),
    ("二、填空题",          "二、填空题",        "第二大题健康"),
    ("1、概念题",           "1、概念题",         "顿号序列健康"),
    ("1、计算题",           "1、计算题",         "第二大题顿号重新起号（健康）"),
]

def main():
    word = wc.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Add()
        for i, (src, _, _) in enumerate(EXPECT):
            if i > 0: doc.Paragraphs.Add()
            r = doc.Paragraphs(i + 1).Range
            r.Text = src + "\r"

        try:
            addin = word.COMAddIns.Item("TypeRight.Connect").Object
        except Exception as ex:
            print("FAIL: 未找到 TypeRight COMAddIn: " + str(ex)); return 2
        if addin is None:
            print("FAIL: TypeRight 未加载（Object 为空）"); return 2

        try:
            summary = addin.AutomationRenumberSequences   # pywin32 动态绑定：带返回值方法按属性取即已调用
        except Exception as ex:
            summary = "INVOKE-EXC: " + str(ex)
        print("summary =", summary)

        text = doc.Range().Text
        lines = [l for l in text.replace("\x07", "").split("\r") if l.strip()]
        fails = 0
        for (src, want, note) in EXPECT:
            key = tail_of(src)
            got = next((l.strip() for l in lines if key in l), None)
            ok = (got == want)
            print(("PASS " if ok else "FAIL ") + src + " -> " + repr(got) + ("  (" + note + ")"))
            if not ok: fails += 1

        # 撤销整体性：一次 doc.Undo() 应还原全部改动
        undone = doc.Undo()
        text2 = doc.Range().Text
        all_restored = all((src + "\r") in text2 or src in text2 for (src, _, _) in EXPECT)
        print(("PASS " if (undone and all_restored) else "WARN ") + "一次 Undo 还原全部: " + str(all_restored))

        doc.Close(False)
        print("RESULT:", "FAIL" if fails else "ALL PASS")
        return 1 if fails else 0
    finally:
        word.Quit()

def tail_of(s):
    """首个 CJK 字符起的正文尾（各行唯一，避开改号后前缀失配）。"""
    for i, ch in enumerate(s):
        if '一' <= ch <= '鿿':
            return s[i:]
    return s

if __name__ == "__main__":
    sys.exit(main())
