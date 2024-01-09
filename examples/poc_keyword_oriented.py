#!/usr/bin/env python3.10
# coding: utf-8
# @carl9527


import os.path as path
from loguru import logger
import copy
import re
import pandas as pd
import numpy as np
from txtrank_summary import TextRankSummarization
from utils import try_or, stylize_df
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity


# Global variables
# ----------------
source_file = 'source-tool7.xlsm'
# Simple reference local file for version 1
pd_path = source_file # path.abspath(path.join(path.dirname(__file__), source_file))

# 20240108, 0105 討論中，BU表示 工具七 中的 "原始資料" 有謬誤，希望從PDF裡取得原始資料。
# 所以這一區域測試從 PDF 裏面取得正確的 原始資料。
# ------------
import tabula
# 這個檔案就是 "許多備註.pdf" or "客戶備註很多.pdf"
source2_file = 'source-tool7.pdf'
# 使用Evon的方式取出所有的page
# Simple reference local file for version 1
pd2_path = source2_file # path.abspath(path.join(path.dirname(__file__), source2_file))
src2_dfs = tabula.read_pdf(pd2_path, area=[120, 5, 800, 1200], pages="all")

# 這裡的 "支出" 與 "存入" 用 "Out" 以及 "In"取代，目的是跟工具七的欄位一致
src2_collect_dfs = {
    '交易日期':[],'帳務日期':[],'交易代號':[],'交易時間':[],'交易分行':[],'交易櫃員':[],
    '摘要':[],'Out':[],'In':[],'餘額':[],
    '轉出入帳號':[],'合作機構':[],'金資序號':[],'票號':[],'備註':[],'註記':[]}

src3_df = pd.DataFrame.from_dict( src2_collect_dfs )

def check_wrong_idx(wrong_idx: list=[], check_from: int=0, check_to: int=-1) -> list:
    results = list()
    if len(wrong_idx) <= 0: return results
    for item in wrong_idx:
        if (item >= check_from) and (item <= check_to):
            results.append(item)

    return results
        
for src_df in src2_dfs:
    src_df = src_df.replace(np.nan, '', regex=True)
    
    headers = src_df.columns.values.tolist()
    wrong_idx = [i for i, item in enumerate(headers) if re.search('^Unnamed:', item)]
    (rows, cols) = src_df.shape

    for row in range(rows):
        if row <= 0:
            continue

        refine_list = []
        value_list = src_df.iloc[row].tolist()

        should_ignore = 0
        for col in range(cols):
            col_name = headers[col]

            if col < 4: # before 4, only a wrong col_name '帳務日期 交易代號', but you can ignore it
                refine_list.append(f'{value_list[col]}')
            if col == 4: # this comes a wrong col_name '交易分行 交易櫃員'
                (col4_val, col5_val) = value_list[col].split(' ')
                refine_list.append(f'{col4_val}')
                refine_list.append(f'{col5_val}')
            if (col >= 5) and (col <= 9): # from '摘要' to '轉出入帳號'
                # check this range
                wlist = check_wrong_idx(wrong_idx, 5, 9)
                if len(wlist) <= 0:
                    refine_list.append(f'{value_list[col]}')
                else:
                    for w in wlist:
                        if col == w:
                            continue
                        else:
                            refine_list.append(f'{value_list[col]}')
            if col == 10:
                refine_list.append(f'{value_list[col]}')
            if col >= 11:
                # check this range
                wlist = check_wrong_idx(wrong_idx, 10, len(value_list))
                if len(wlist) <= 0:
                    refine_list.append(f'{value_list[col]}')
                else:
                    for w in wlist:
                        if (col == w) and (col != should_ignore):
                            should_ignore = col + 1
                            refine_list.append(f'{value_list[col]}')
                        elif should_ignore > 0:
                            should_ignore = 0
                            continue
                        else:
                            refine_list.append(f'{value_list[col]}')

        src3_df.loc[len(src3_df.index)] = refine_list
# -----------

select_year = '2022'
rawdata_sheet = '原始資料'
atm_info_sheet = 'Report_MachineManage'

# Read sheets to dataframes
# -------------------------
'''
Read and clean column headers
Like this: 代碼(記事本)\n=9&B2
'''
srcdata = src3_df.copy(deep=True)
if False:
    # 這個是從 工具七 裡讀出 "原始資料" 頁
    srcdata = pd.read_excel(pd_path, sheet_name=rawdata_sheet, skiprows=7)
    srcdata.columns = srcdata.columns.str.split('\\n').str[0]
logger.debug(srcdata)

rawdata_csv = 'rawdata-tool8.xlsx'
with pd.ExcelWriter(rawdata_csv) as writer:
    srcdata.to_excel(writer, sheet_name="Rawdata", index=False)

# 這個是從 工具七 裡讀出 "Report_MachineManage" 頁
rmdata = pd.read_excel(pd_path, sheet_name=atm_info_sheet, skiprows=-1)
rmdata.columns = rmdata.columns.str.split('\\n').str[0]

# Filter raw data
# ---------------
'''
Condition
'''
rawdata = srcdata[
        #((srcdata['交易日期'] >= f'{select_year}-01-01') & (srcdata['交易日期'] <= f'{select_year}-12-31')) &
        (((srcdata['Out'] == srcdata['Out']) & (srcdata['Out'] != '')) |
        ((srcdata['In'] == srcdata['In']) & (srcdata['In'] != '')))
        ]


# Formatting column values
# ------------------------
'''
Formatting everything we want
'''
def atm_formatting(s):
    id_code = try_or(lambda: f"{s['ID1']:08d}", default=f"{s['ID1']}")
    deal_code = try_or(lambda: f"{s['剖析']:04d}", default=f"{s['剖析']}")
    deal_teller = try_or(lambda: f"{s['代碼(記事本)']:05d}", default=f"{s['代碼(記事本)']}")
    return id_code, deal_code, deal_teller


rmdata['ID1'], rmdata['剖析'], rmdata['代碼(記事本)'] = zip(*rmdata.apply(atm_formatting, axis=1))
rmdata = rmdata.replace(np.nan, '', regex=True)
logger.debug(rmdata)


def rawdata_formatting(s):
    deal_date = try_or(lambda: f"{s['交易日期'].year:04d}/{s['交易日期'].month:02d}/{s['交易日期'].day:02d}", default=f"{s['交易日期']}")
    acc_date = try_or(lambda: f"{s['帳務日期'].year:04d}/{s['帳務日期'].month:02d}/{s['帳務日期'].day:02d}", default=f"{s['帳務日期']}")
    deal_code = try_or(lambda: f"{s['交易代號']:04d}", default=f"{s['交易代號']}")
    deal_time = try_or(lambda: f"{s['交易時間'].hour:02d}:{s['交易時間'].minute:02d}:{s['交易時間'].second:02d}", default=f"{s['交易時間']}")
    deal_branch = try_or(lambda: f"{s['交易分行']:03d}", default=f"{s['交易分行']}")
    deal_teller = try_or(lambda: f"{s['交易櫃員']:05d}", default=f"{s['交易櫃員']}")
    summary = try_or(lambda: f"{s['摘要']}".strip(), default=f"{s['摘要']}")
    m_out = try_or(lambda: f"{s['Out']:,.2f}", default=f"{s['Out']}")
    m_in = try_or(lambda: f"{s['In']:,.2f}", default=f"{s['In']}")
    m_balance = try_or(lambda: f"{s['餘額']:,.2f}", default=f"{s['餘額']}")
    tr_acc = try_or(lambda: f"{s['轉出入帳號']}".strip(), default=f"{s['轉出入帳號']}")
    tr_infra = try_or(lambda: f"{s['合作機構']}".strip(), default=f"{s['合作機構']}")

    comment = f"{s['備註']}".strip()
    if False:
        # 工具七裏面 幾乎備註都在 '票號/備註' 裡
        comment = f"{s['備註']}".strip() if len(f"{s['備註']}".strip()) > 0 else f"{s['票號/備註']}".strip()

    return deal_date, acc_date, deal_code, deal_time, deal_branch, deal_teller, \
            summary, m_out, m_in, m_balance, tr_acc, tr_infra, comment


src_dtls = {
        '交易日期':[],'帳務日期':[],'交易代號':[],'交易時間':[],
        '交易分行':[],'交易櫃員':[],'摘要':[],'支出':[],'存入':[],
        '餘額':[],'轉出入帳號':[],'合作機構':[],'備註':[]}

rawdata = rawdata.replace(np.nan, '', regex=True)
src_df = pd.DataFrame(copy.deepcopy(src_dtls))
src_df['交易日期'], src_df['帳務日期'], src_df['交易代號'],\
        src_df['交易時間'], src_df['交易分行'], src_df['交易櫃員'],\
        src_df['摘要'], src_df['支出'], src_df['存入'],\
        src_df['餘額'], src_df['轉出入帳號'], src_df['合作機構'],\
        src_df['備註'] \
        = zip(*rawdata.apply(rawdata_formatting, axis=1))

logger.debug(src_df)


# Create intermediate product
# ------------------------
'''
Text4Rank
'''
g_metadata = TextRankSummarization()
g_similar_max = float(1)
g_similar_min = float(0)
g_similar_threshold = float(0.44) # 相似閥值拉在 0.44
g_tfidf_vectorizer = TfidfVectorizer(analyzer="char")

def intermediate_builder(s, rm_df, kw_dict):
    comment = try_or(lambda: f"{s['備註']}".strip(), default=f"{s['備註']}")

    kitems = g_metadata.keywords(comment, count=1)
    keyword = try_or(lambda: f"{kitems[0]['word'].strip().lower()}", default='')

    summary = try_or(lambda: f"{s['摘要']}".strip(), default=f"{s['摘要']}")
    deal_date = try_or(lambda: f"{s['交易日期']}".strip(), default=f"{s['交易日期']}")
    deal_time = try_or(lambda: f"{s['交易時間']}".strip(), default=f"{s['交易時間']}")

    deal_code = try_or(lambda: f"{s['交易代號']}".strip(), default=f"{s['交易代號']}")

    deal_branch = try_or(lambda: f"{s['交易分行']}".strip(), default=f"{s['交易分行']}")
    deal_teller = try_or(lambda: f"{s['交易櫃員']}".strip(), default=f"{s['交易櫃員']}")
    atm_addr = try_or(lambda: f"{rm_df[rm_df['代碼(記事本)'] == deal_teller]['地址-區域'].to_list()[0]}", default=deal_teller)

    m_out = try_or(lambda: f"{s['支出']}".strip(), default=f"{s['支出']}")
    m_in = try_or(lambda: f"{s['存入']}".strip(), default=f"{s['存入']}")

    sort_idx = f'{deal_date}={deal_time}={deal_branch}={deal_teller}'
    if len(keyword) > 0:
        if keyword not in list(kw_dict.keys()):
            kw_dict[keyword] = list()
        kw_dict[keyword].append(sort_idx)

    return keyword, comment, summary, deal_date, deal_time,\
            deal_branch, deal_teller, atm_addr, m_out, m_in


mid1_df = src_df[src_df['備註'] != '']

mid_dtls = {
        '關鍵字':[],'備註':[],'摘要':[],'交易日期':[],'交易時間':[],
        '交易分行':[],'交易櫃員':[],'ATM機台據點':[],'支出/Out':[],'存入/In':[]
        }

logger.debug(mid1_df)

mid2_kw = dict()
mid2_df = pd.DataFrame.from_dict( mid_dtls )

mid2_df['關鍵字'], mid2_df['備註'], mid2_df['摘要'], mid2_df['交易日期'],\
        mid2_df['交易時間'], mid2_df['交易分行'],\
        mid2_df['交易櫃員'], mid2_df['ATM機台據點'], mid2_df['支出/Out'], mid2_df['存入/In'] \
        = zip(*src_df.apply(lambda x: intermediate_builder(x, rmdata, mid2_kw), axis=1))

mid2_kw_keys = [k for _, k in sorted(zip(map(len, mid2_kw.values()), mid2_kw.keys()), reverse=True)]
mid3_df = pd.DataFrame.from_dict( mid_dtls )
for kw in mid2_kw_keys:
    cl = mid2_kw[kw]
    for cl_idx in cl:
        deal_date, deal_time, deal_branch, deal_teller = '', '', '', ''
        [deal_date, deal_time, deal_branch, deal_teller] = cl_idx.split('=')
        temp_df = mid2_df[(mid2_df['交易日期'] == deal_date) & \
                    (mid2_df['交易時間'] == deal_time) & \
                    (mid2_df['交易分行'] == deal_branch) & \
                    (mid2_df['交易櫃員'] == deal_teller)]
        mid3_df.loc[len(mid3_df.index)] = temp_df.iloc[0].tolist()


mid2_df = mid2_df[mid2_df['關鍵字'] != '']
logger.debug(mid2_df)
logger.debug(mid3_df)

middle_csv = 'middle-tool8.xlsx'

# Convert to a excel as intermediate product
with pd.ExcelWriter(middle_csv) as writer:
    mid3_df.to_excel(writer, sheet_name="Keywords", index=False)


# Create final product
# ------------------------
final_csv = 'result-tool8.xlsx'

final1_df = mid3_df.copy(deep=True)

old_str = ''
col_list = final1_df.columns.values.tolist()
col_idx = 0
col_name = final1_df.columns[col_idx]
old_list = final1_df[col_name].tolist()
new_list = list()
for st in old_list:
    new_st = st
    if st != old_str:
        old_str = st
    else:
        new_st = ''

    new_list.append(new_st)

final1_df.drop(col_name, axis = 1, inplace = True)
final1_df.insert(col_idx, col_name, new_list) 


(
    pd.DataFrame([final1_df.to_dict('list')])
        .apply(pd.Series.explode)
        .pivot_table(index=col_list, sort=False)
        .style.applymap_index(stylize_df)
        .to_excel(final_csv, startrow=-1)
)

