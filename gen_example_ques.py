import openpyxl
import random

questions = [['no', 'question', 'cho1', 'cho2', 'cho3', 'cho4', 'ans', 'pic']]
groups = ['M', 'N', 'O', 'P']
cats = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']

for group in groups:
    for cat in cats:
        # generate 10 questions per each group-cat
        for x in range (1, 11):
            ques_num = group + f"{x:02}" + cat
            question = f"This is question {ques_num}."
            answer = random.randint(1, 4)
            pic_no = random.randint(1, 30)
            if pic_no <= 7:
                pic = f"{pic_no}.jpg"
            else:
                pic = ""
            questions.append([ques_num, question, "choice 1", "choice 2", "choice 3", "choice 4", answer, pic])

workbook = openpyxl.Workbook()
# Select the default sheet (usually named 'Sheet')
sheet = workbook.active
for item in questions:
    sheet.append(item)
workbook.save("example_ques.xlsx")
