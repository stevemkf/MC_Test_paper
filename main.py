import tkinter as tk
from tkinter import filedialog
import os
import pandas as pd
import openpyxl
from draw_ques import DrawQuestions
from genPDF import ExamPDF

# global variables
folder_name = ""


# Create mark sheet as Excel file
def write_marksheet(excel_rows, paper_no):
    global output_path, config_dict

    file_marksheet = config_dict['marksheet']
    workbook = openpyxl.Workbook()
    # Select the default sheet (usually named 'Sheet')
    sheet = workbook.active
    for item in excel_rows:
        sheet.append(item)
    filename = f"{file_marksheet}-{paper_no:02d}.xlsx"
    workbook.save(os.path.join(output_path, filename))


def gen_paper(paper_no):
    global q, output_path, config_dict

    # retrieve configuration file contents from a global dictionary
    file_ques_bank = config_dict['question bank']
    first_group = config_dict['first group']
    mid_group = config_dict['mid group']
    last_group = config_dict['last group']
    first_category = config_dict['first category']
    last_category = config_dict['last category']
    ques_per_cat_str = str(config_dict['questions per category'])
    ques_per_cat_list = [int(item) for item in ques_per_cat_str.split(',')]
    language = config_dict['language']
    file_testpaper = config_dict['test paper']

    # one question dataframe is sufficient for several test papers
    if paper_no == 1:
        q = DrawQuestions(file_ques_bank, first_group, last_group, first_category, last_category)

    # draw different sets of questions for each paper
    index_df_list = q.get_ques_list(first_group, mid_group, last_group, ques_per_cat_list)
    # assign different filenames for the test papers
    filename = f"{file_testpaper}-{paper_no:02d}.pdf"
    exam_pdf = ExamPDF(os.path.join(output_path, filename), language)
    # prepare Excel mark sheets as well
    excel_rows = [['no', 'question', 'correct_ans']]

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

        # built the pdf question paper
        exam_pdf.write_question(ques_num_paper, question, choice_1, choice_2, choice_3, choice_4, image)
        # build the mark sheet
        excel_rows.append([ques_num_paper, question_num, answer])
        ques_num_paper = ques_num_paper + 1

    # save the test paper as a pdf file
    exam_pdf.close()
    # save the mark sheet as an Excel file
    write_marksheet(excel_rows, paper_no)


def gen_papers():
    global folder_name, output_path, config_dict

    msg_box.delete(0.0, tk.END)
    # conduct preliminary check on user inputs
    if folder_name == "":
        message = "Please load the configuration file first!"
        msg_box.insert(tk.END, message)
    elif 'question bank' not in config_dict:
        message = "Please load a correct configuration file first!"
        msg_box.insert(tk.END, message)
    else:
        input_str = num_papers.get()
        if input_str.isnumeric():

            # prepare the output folder for test papers and mark sheets
            output_path = os.path.join(folder_name, "output")
            if not os.path.exists(output_path):
                os.makedirs(output_path)
            files = [os.path.join(output_path, f) for f in os.listdir(output_path)]
            for f in files:
                os.remove(f)

            # create the test papers and mark sheets one by one
            total_papers = int(input_str)
            for paper_no in range(1, total_papers + 1):
                gen_paper(paper_no)
                # give feedback on the progress
                message = f"Paper {paper_no} was created.\n"
                msg_box.insert(tk.END, message)
                window.update()
        else:
            message = "Please specify how many papers!"
            msg_box.insert(tk.END, message)


def select_config():
    global folder_name, config_dict

    title = 'Select your config.xlsx'
    home_dir = os.path.expanduser("~")
    file_path = filedialog.askopenfilename(title=title, initialdir=home_dir, filetypes=[("Excel files", "*.xlsx"), ])
    folder_name, file_name = os.path.split(file_path)
    msg_box.delete(0.0, tk.END)
    # conduct preliminary checks
    if file_name != "config.xlsx":
        message = "Incorrect file!"
        msg_box.insert(tk.END, message)
    else:
        config_path.delete(0, tk.END)
        config_path.insert(0, file_path)
        df = pd.read_excel(file_path, header=None)
        config_dict = dict(zip(df[0], df[1]))
        # make a simple check on the configuration file contents
        if 'question bank' in config_dict:
            message = "Configuration file read!"
        else:
            message = "Configuration file contents incorrect!"
        msg_box.insert(tk.END, message)


# main loop providing the user interface
window = tk.Tk()
window.geometry("800x480")
window.title('Generate MC Exam Paper')

btn_config = tk.Button(text="Select configuration file", command=select_config)
btn_config.grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
config_path = tk.Entry(width=60)
config_path.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
label_2 = tk.Label(text="Generate how many exam papers?")
label_2.grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
num_papers = tk.Entry()
num_papers.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
btn_papers = tk.Button(text="Generate exam papers", command=gen_papers)
btn_papers.grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
msg_box = tk.Text(width=80)
msg_box.grid(row=3, column=0, columnspan=2, sticky=tk.W, padx=5, pady=5)

window.mainloop()






