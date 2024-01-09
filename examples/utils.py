#!/usr/bin/env python3.10
# coding: utf-8
# @carl9527


import os, sys
from functools import lru_cache


@lru_cache()
def try_or(func, default=None, expected_exc=(Exception,)):
    try:
        return func()
    except expected_exc:
        return default

@lru_cache()
def stylize_df(s):
    return "font-weight: normal; text-align: center; vertical-align: middle;"

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.environ.get("_MEIPASS2", os.path.abspath("."))

    return os.path.join(base_path, relative_path)
