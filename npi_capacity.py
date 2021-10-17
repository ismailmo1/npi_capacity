#!/usr/bin/env python
# coding: utf-8

import math
import re

import numpy as np
import pandas as pd
from pandas.tseries.offsets import BDay

## read in excel reports


wtl_df = pd.read_excel("TSS_ShopFloorWorkToList.xls", header=8)
ordrbk_df = pd.read_excel(
    "TSS_Assembly_Silastics_MachinedSeals_Report.xls", header=0
)

wtl_items = set(wtl_df["Item Number"].unique())
ordrbk_items = set(ordrbk_df["COItem"].unique())
not_released = ordrbk_items.difference(wtl_items)
no_orders = wtl_items.difference(ordrbk_items)
item_list = ",".join(list(ordrbk_items))
print(
    str(len(no_orders)) + " items without orders: \n",
    wtl_items.difference(ordrbk_items),
    end="\n\n",
)
print(
    str(len(not_released)) + " items not released: \n",
    ordrbk_items.difference(wtl_items),
    end="\n\n",
)


df_sql = pd.read_excel("npi_capacity_refresh.xlsm", engine="openpyxl").dropna(
    how="all", axis=1
)
# con = pyodbc.connect('Driver={SQL Server};'
#                       'Server=tssstrfsh006;'
#                       'Database=fsdbtt;'
#                       'Trusted_Connection=yes;')

# tool_qry = f'''SELECT TOP 200 Bills.ComponentItemNumber,
# Bills.RequiredQuantity,
# Bills.ParentItemKey,
# Bills.ComponentItemKey,
# Items.ItemNumber As ParentItemNum,
# Items.ItemDescription As ParentItemDesc,
# Items2.ItemDescription as CompDesc

# FROM
# 	FS_BillOfMaterial Bills JOIN FS_Item Items
# 		ON Bills.ParentItemKey = Items.ItemKey
# 	JOIN FS_Item Items2
# 		ON Bills.ComponentItemKey = Items2.ItemKey

# WHERE (Bills.ComponentItemNumber IN {item_list}) AND (Items.ItemNumber LIKE '___-___-%') AND
# (OR Bills.ComponentItemNumber LIKE 'WC[[]R]___-___%'
# OR Bills.ComponentItemNumber LIKE 'WC[[]R]%Y[0-9]%')'''

# tool_qry = f'''SELECT TOP 200 Bills.ComponentItemNumber,
# Bills.RequiredQuantity,
# Bills.ParentItemKey,
# Bills.ComponentItemKey,
# Items.ItemNumber As ParentItemNum,
# Items.ItemDescription As ParentItemDesc,
# Items2.ItemDescription as CompDesc

# FROM
# 	FS_BillOfMaterial Bills JOIN FS_Item Items
# 		ON Bills.ParentItemKey = Items.ItemKey
# 	JOIN FS_Item Items2
# 		ON Bills.ComponentItemKey = Items2.ItemKey

# WHERE (Bills.ComponentItemNumber IN {item_list}) AND (Items.ItemNumber LIKE '___-___-%') AND
# (Bills.ComponentItemNumber LIKE 'WC[[]R]MOULD%'
# OR Bills.ComponentItemNumber LIKE 'WC[[]R]1010%'
# OR Bills.ComponentItemNumber LIKE 'WC[[]R]1012%')'''

# df_tool = pd.read_sql(tool_qry, con)
# df_mould = pd.read_sql(mould_qry, con)


def match_tool_desc(rw):
    """returns true if tool desc matches regex pattern 'WC\[R][0-9]{3}-[0-9]{3}"""
    return bool(re.match("WC\[R][0-9]{3}-[0-9]{3}", rw["ComponentItemNumber"]))


tool_mask = df_sql.apply(lambda x: match_tool_desc(x), axis=1)

df_tool = df_sql.loc[tool_mask]
df_mould = df_sql.loc[~tool_mask]

# clean tool desc to get cavs out (reusing from alternate tool finder)

# regex looks mad complicated but use  pyregex.com to help test patterns
# use findall to return list of matches
df_tool.loc[:, "CAVS"] = df_tool.loc[:, "CompDesc"].apply(
    lambda x: re.findall("[,]?\s?[S/.0-9]{1,3}\s*(?:CAV[S]?|C|INS)[,]?", x)
)

# function to extract numbers from cavs list
def getcavs(cavlist):

    # create mini func to convert stuff to numbers and sort out single cavs
    def cav_to_int(cavstr):
        return int(re.sub("S[/.]?", "1", cavstr))

    num = []
    # loop through list of numbers to find numbers or 'S/',
    # this will look for first match and return it, otherwise no return
    for x in cavlist:
        m = re.search("S[/.]|[0-9]{1,3}", x)
        if m:
            num.append(m.group(0))

    # if theres only one match then return that
    if len(num) == 1:
        cavs = cav_to_int(num[0])
        # assume 200 cavs is incorrect detection
        if cavs > 200:
            cavs = 1

        return cavs
    # if theres two matches..
    elif len(num) == 2:
        # and both are the same, add the first element of that list together and return it
        if num[0] == num[1]:
            return sum((cav_to_int(num[0]), cav_to_int(num[1])))
        # if theyre not the same then return the full list (need to convert to number to use max())
        else:
            return max((cav_to_int(num[0]), cav_to_int(num[1])))
    else:
        # return 1 cav as fallback
        return 1


df_tool.loc[:, "ToolCavs"] = df_tool.loc[:, "CAVS"].apply(getcavs)

# get rid of initial cav column
df_tool.drop("CAVS", axis=1, inplace=True)

df_tool = df_tool.loc[:, ("ParentItemNum", "ToolCavs")]

df_mould.loc[:, "MouldTimeMin"] = df_mould.loc[:, "RequiredQuantity"] * 60

df_mould = df_mould.loc[:, ("ParentItemNum", "MouldTimeMin")]

df_cycle_times = df_tool.merge(
    df_mould, on="ParentItemNum", how="outer", indicator=True
)


def findCycleTime(rw):
    """use finalCavs to find cycle time - threshold: 15min - 5min(we're trying to guess what the accountants used as the basis for mould time calc)"""
    try:
        cycle_time = round(rw["ToolCavs"] * rw["MouldTimeMin"])
    except ValueError:
        return None
    if cycle_time <= 5:
        return 5
    elif cycle_time > 15:
        return 15
    return cycle_time


df_cycle_times["ExpectedCycleTime"] = df_cycle_times.apply(
    lambda x: findCycleTime(x), axis=1
)

df_cycle_times = df_cycle_times.loc[
    :, ("ParentItemNum", "ToolCavs", "ExpectedCycleTime")
]


df_cycle_times = pd.read_csv("mock_cycle_times.csv")


# ## get current stock levels


ordrbk_df = ordrbk_df[
    [
        "COItem",
        "Family",
        "Trans LT",
        "Req Ship",
        "COQty",
        "Line Value",
        "Stock",
        "MONumber",
        "Scan Point",
        "Last Scan",
        "COLExt Text",
    ]
]


def get_curr_stock(stock: str):
    """extract total stock from stock location identifier
    ignores 99 and TODUMP locations
    i.e. get_curr_stock("11 FAIR 4, 11 S33B13 20, 99 FAIR 4, 11 TODUMP 1") -> 24

    """

    total_qty = 0
    try:
        stock_locs = stock.split(", ")
    except:
        return None
    stock_qty = 0

    for s in stock_locs:
        stock_split = s.split(" ")
        if stock_split[1] == "TODUMP" or stock_split[0] == "99":
            stock_qty = 0
        else:
            stock_qty = int(stock_split[2])

        total_qty += stock_qty

    return total_qty


ordrbk_df["curr_stock"] = ordrbk_df["Stock"].apply(lambda x: get_curr_stock(x))

ordrbk_df["req_week"] = ordrbk_df["Req Ship"].apply(lambda x: x.week)


# # add document requirements (3.1, FAIR etc)


def detect_quality_docs(
    ext_text: str,
    keywords: list = ["PSQ", "3.1", "FAIR", "ISIR", "LICENSE", "LICENCE"],
) -> str:
    """detects presence of keywords in ext_text"""
    if type(ext_text) == str:
        for word in keywords:
            if word in ext_text.upper():
                return word
    return None


ordrbk_df["extra_docs"] = ordrbk_df["COLExt Text"].apply(detect_quality_docs)
ordrbk_df["extra_docs_days"] = ordrbk_df["extra_docs"].apply(
    lambda x: 4 if x else 0
)


# ## calculate demand based off stock and CO qty


# curr_stock_df = ordrbk
curr_stock_df = ordrbk_df.dropna(subset=["curr_stock"])[
    ["COItem", "curr_stock"]
]

ordrbk_stock_df = ordrbk_df[
    ["COItem", "COQty", "req_week", "Req Ship", "extra_docs"]
].sort_values("req_week")
ordrbk_stock_df["cumsum"] = ordrbk_stock_df.groupby(["COItem"])[
    "COQty"
].transform(pd.Series.cumsum)
# ordrbk_stock_df['curr_stock'] = ordrbk_stock_df.groupby(['COItem'])['curr_stock'].bfill()

curr_stock_df.drop_duplicates(subset="COItem", inplace=True)

ordrbk_stock_df = ordrbk_stock_df.merge(curr_stock_df, on="COItem", how="left")

ordrbk_stock_df["cumul_demand"] = (
    ordrbk_stock_df["cumsum"] - ordrbk_stock_df["curr_stock"]
)


def get_week_demand(rw):
    """returns qty demand for week
    returns co qty if no stock in place,
    else returns lower qty discounted due to that week's stock"""
    if rw["cumul_demand"] < 0:
        return 0
    if rw["COQty"] > rw["cumul_demand"]:
        return rw["cumul_demand"]
    else:
        return rw["COQty"]


ordrbk_stock_df["demand_qty"] = ordrbk_stock_df.apply(
    lambda x: get_week_demand(x), axis=1
)
ordrbk_stock_df = ordrbk_stock_df[
    [
        "COItem",
        "req_week",
        "Req Ship",
        "COQty",
        "curr_stock",
        "demand_qty",
        "extra_docs",
    ]
]


mould_setup_mins = 90
pre_mould_days = 2
post_mould_days = 9
doc_days = 4
total_shift_mins = 360

lead_time_df = ordrbk_stock_df.merge(
    df_cycle_times, left_on="COItem", right_on="ParentItemNum", how="left"
)
lead_time_df["ToolCavs"].fillna(1, inplace=True)
lead_time_df["ExpectedCycleTime"].fillna(12, inplace=True)
# calculate num of lifts required
lead_time_df["lifts"] = lead_time_df["demand_qty"] / lead_time_df["ToolCavs"]
lead_time_df["lifts"] = lead_time_df["lifts"].apply(np.ceil)

# use num of lifts as boolean mask to remove process days for parts with no manufacturing required)
process_needed_mask = lead_time_df["lifts"].astype(bool)


lead_time_df["mould_mins"] = (
    lead_time_df["lifts"] * lead_time_df["ExpectedCycleTime"]
)
lead_time_df["mould_setup_mins"] = process_needed_mask * mould_setup_mins

lead_time_df["total_mould_mins"] = (
    lead_time_df["mould_setup_mins"] + lead_time_df["mould_mins"]
)
lead_time_df["total_mould_days"] = (
    lead_time_df["total_mould_mins"] / total_shift_mins
)

lead_time_df["process_time_days"] = process_needed_mask * (
    pre_mould_days
    + post_mould_days
    + lead_time_df["total_mould_days"]
    + ((lead_time_df["extra_docs"]).astype(bool) * doc_days)
)
lead_time_df["process_time_days"] = lead_time_df["process_time_days"].apply(
    np.ceil
)

lead_time_df["process_time_weeks"] = lead_time_df["process_time_days"] / 5
lead_time_df["process_time_weeks"] = lead_time_df["process_time_weeks"].apply(
    np.ceil
)


lead_time_df["start_week"] = (
    lead_time_df["req_week"] - lead_time_df["process_time_weeks"]
)


lead_time_df["mould_start_day"] = lead_time_df.apply(
    lambda x: x["Req Ship"]
    - BDay(post_mould_days)
    - BDay(math.ceil(x["total_mould_days"])),
    axis=1,
)
lead_time_df["mould_end_day"] = lead_time_df.apply(
    lambda x: x["mould_start_day"]
    + BDay(math.ceil(x["total_mould_days"]))
    - BDay(1),
    axis=1,
)

lead_time_df["moulding_dates"] = lead_time_df.apply(
    lambda x: pd.bdate_range(x["mould_start_day"], x["mould_end_day"]), axis=1
)


capacity_split_df = lead_time_df.explode("moulding_dates")


capacity_split_df["daily_mould_mins"] = capacity_split_df[
    "total_mould_mins"
] / capacity_split_df["total_mould_days"].apply(np.ceil)

capacity_split_df["moulding_week"] = capacity_split_df["moulding_dates"].apply(
    lambda x: x.week
)


capacity_split_df.to_csv("capacity_output.csv", index=False)
