import pandas as pd
import openpyxl

df = pd.read_excel("config.xlsx", header=None)
d = dict(zip(df[0], df[1]))

# It is assumed that the question groups are divided into two batches.
# Each batch contribute one group of questions for each category, MxxA, OxxA.
# If not, the "get_ques_list" function in draw_ques.py needs to be modified
first_group = d['first group']
mid_group = d['mid group']
last_group = d['last group']
first_category = d['first category']
last_category = d['last category']
ques_per_cat_str = str(d['questions per category'])
ques_per_cat_list = [int(item) for item in ques_per_cat_str.split(',')]

file_ques_bank = d['question bank']
file_testpaper = d['test paper']
file_marksheet = d['marksheet']
