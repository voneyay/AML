#!/usr/bin/env python
# coding: utf-8
# @carl9527


import sys
from textrank4zh import TextRank4Keyword, TextRank4Sentence
from utils import resource_path


STOPWORDS = resource_path('stop_words.txt')


class TextRankSummarization():
    _instance = None
    def __new__(cls, *args, **kwargs):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        try:
            import importlib
            importlib.reload(sys)
        except:
            pass

    def keywords(self, content:str=None, count:int=20):
        if not content:
            return []

        tr4w = TextRank4Keyword(stop_words_file=STOPWORDS)
        # python2中text必須是utf8編碼的str或者unicode對象，python3中必須是utf8編碼的bytes或者str對象
        tr4w.analyze(text=content, lower=False, window=3)#2)
        # return a list of dict
        return tr4w.get_keywords(count, word_min_len=2)

    def keyphrases(self, content:str=None, count:int=20):
        if not content:
            return []

        tr4w = TextRank4Keyword(stop_words_file=STOPWORDS)
        # python2中text必須是utf8編碼的str或者unicode對象，python3中必須是utf8編碼的bytes或者str對象
        tr4w.analyze(text=content, lower=False, window=3)#2)
        # return a list
        return tr4w.get_keyphrases(keywords_num=count, min_occur_num=3)#2)

    def keysentences(self, content:str=None, count:int=5):
        if not content:
            return []

        tr4s = TextRank4Sentence(stop_words_file=STOPWORDS)
        tr4s.analyze(text=content, lower=False, source = 'all_filters')
        # return a list of dict
        return tr4s.get_key_sentences(num=count)
