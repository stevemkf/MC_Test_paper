import pandas as pd
import openpyxl
import random


# Parse the question bank Excel file and build up a 2D question position list, based on Group and Category
class DrawQuestions():
    def __init__(self, file_ques_bank, first_group, last_group, first_category, last_category):
        self.df = pd.read_excel(file_ques_bank)

        # Question group: M, N, O, P
        # Question category: A to H
        total_groups = ord(last_group) - ord(first_group) + 1
        total_cats = ord(last_category) - ord(first_category) + 1
        # Create a 2D question position list.  Position means Excel dataframe row number
        self.question_pos_lists = [[[] for col in range(total_cats)] for row in range(total_groups)]

        # Parse all rows in the question worksheet and fill in the 2D question position lists
        for index_df, row in self.df.iterrows():
            # question_num: e.g. M01A, P04H
            question_num = row['no']
            group = ord(question_num[0]) - ord(first_group)
            cat = ord(question_num[len(question_num) - 1]) - ord(first_category)
            # defensive programming - avoid index out of range
            if (group <= ord(last_group) - ord(first_group)) and \
                    (cat <= ord(last_category) - ord(first_category)):
                self.question_pos_lists[group][cat].append(index_df)


    # Randomly draw questions for one test paper.  Return a list of indexes for the dataframe
    def get_ques_list(self, first_group, mid_group, last_group, ques_per_cat_list):
        index_df_list = []
        # The questions will follow the category orders, i.e. A to H
        # defensive programming - take care of excessive entries in config.sys
        num_cat = len(self.question_pos_lists[0])
        for index_cat, num_ques_cat in enumerate(ques_per_cat_list[0:num_cat]):
            # Choose either Group M or N + either Group O or P for each category of questions
            group1_end = ord(mid_group) - ord(first_group)
            group1 = random.randint(0, group1_end)
            cat_ques_list = self.question_pos_lists[group1][index_cat]
            # one group of two groups of questions?
            if last_group != mid_group:
                group2_end = ord(last_group) - ord(first_group)
                group2 = random.randint(group1_end + 1, group2_end)
                cat_ques_list = cat_ques_list + self.question_pos_lists[group2][index_cat]
            cat_ques_drawn = random.sample(cat_ques_list, num_ques_cat)
            # Sort the questions so that it is easier to check the results
            cat_ques_drawn.sort()
            # Combine the questions to form the test paper
            index_df_list = index_df_list + cat_ques_drawn
        return index_df_list


    def get_ques_num_ans_list(self, index_df_list):
        ques_num_list = []
        ques_ans_list = []
        for index_df in index_df_list:
            row_content = self.df.iloc[index_df]
            ques_num_list.append(row_content['no'])
            ques_ans_list.append(str(row_content['ans']))
        return ques_num_list, ques_ans_list


    # Return full data of a particular question in the format of Python dictionary
    def get_question(self, index_df):
        row_content = self.df.iloc[index_df]
        ques = dict()
        ques["question_num"] = row_content['no']    # e.g. e.g. M01A, P04H
        ques["question"] = row_content['question']
        ques["choice_1"] = str(row_content['cho1'])
        ques["choice_2"] = str(row_content['cho2'])
        ques["choice_3"] = str(row_content['cho3'])
        ques["choice_4"] = str(row_content['cho4'])
        ques["answer"] = row_content['ans']
        image = row_content['pic']
        # empty cell return NaN which should be replaced with empty string
        if image != image:
            image = ""
        ques["image"] = image
        return ques
