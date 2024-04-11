from draw_ques import DrawQuestions
import openpyxl
from genPDF import ExamPDF
from read_config import *


# Create mark sheet as Excel file
def write_marksheet(marksheet):
    workbook = openpyxl.Workbook()
    # Select the default sheet (usually named 'Sheet')
    sheet = workbook.active
    for item in marksheet:
        sheet.append(item)
    workbook.save(file_marksheet)


q = DrawQuestions(file_ques_bank, first_group, last_group, first_category, last_category)
index_df_list = q.get_ques_list(ques_per_cat_list)
exam_pdf = ExamPDF(file_testpaper)
marksheet = []
marksheet.append(['no', 'question', 'correct_ans'])


# Retrieve the question contents from dataframe, based on their position numbers
ques_num_paper = 1
for index_df in index_df_list:
    ques = q.get_question(index_df)
    question_num = ques["question_num"]
    question = ques["question"]
    choice_1 = ques["choice_1"]
    choice_2 = ques["choice_2"]
    choice_3 = ques["choice_3"]
    choice_4 = ques["choice_4"]
    answer = ques["answer"]
    image = ques["image"]

    # build the mark sheet
    marksheet.append([ques_num_paper, question_num, answer])
    # built the pdf question paper
    exam_pdf.write_question(ques_num_paper, question, choice_1, choice_2, choice_3, choice_4, image)
    ques_num_paper = ques_num_paper + 1


# Save the mark sheet as an Excel file
write_marksheet(marksheet)
exam_pdf.close()





