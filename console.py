# -*- coding: utf-8 -*-
"""控制台输出编码保护。

程序的提示文案是中文，店铺名本身也可能是中文（清货1_PW03、清货2_EE02）。
Windows 控制台默认编码是 cp936/cp1252，一旦把输出重定向到文件或管道，
Python 会按该编码写入并抛 UnicodeEncodeError —— 整个批次会因为一句 print
而中断（GUIDE.md 6.4 记录过这个坑）。

入口脚本导入本模块即可，不需要每次都在命令行前面加 PYTHONIOENCODING=utf-8。
"""
import sys


def force_utf8():
    """把 stdout/stderr 切到 UTF-8，失败也不影响程序继续跑。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            # errors="replace"：宁可打印成问号，也不要因为一个字符终止批次
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            # Python 3.6 及以下没有 reconfigure；被重定向成非文本流时也会失败
            pass


force_utf8()
