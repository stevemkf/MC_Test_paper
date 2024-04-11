from reportlab.platypus import SimpleDocTemplate, Table, Paragraph, ListFlowable, Image, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus.tables import TableStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.units import inch

# the following two lines are required to write Chinese characters onto pdf file
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont


class ExamPDF():
    def __init__(self, file_testpaper):
        # the outermost container for the document
        self.doc = SimpleDocTemplate(file_testpaper,
                                rightMargin=24,
                                leftMargin=48,
                                topMargin=24,
                                bottomMargin=24,
                                allowSplitting=False)        # a question will not be splitted on two pages

        # from Adobe's Asian Language Packs
        font_Chinese = "STSong-Light"
        pdfmetrics.registerFont(UnicodeCIDFont(font_Chinese))

        # specifications for layouts of pages of various kinds
        styles = getSampleStyleSheet()
        styles['Normal'].fontName = font_Chinese
        styles['Normal'].leading = 24                       # = line height
        self.normal = styles['Normal']
        styles['Heading2'].fontName = font_Chinese
        styles['Heading2'].alignment = TA_CENTER
        self.big_char = styles['Heading2']

        # pdf document contents will be built up in story
        self.story = []


    def write_question(self, ques_num_paper, question, choice_1, choice_2, choice_3, choice_4, image):
        # the reportlab table coordinates are (col, row)
        cell_00 = Paragraph("Q" + str(ques_num_paper), self.normal)
        cell_10 = Paragraph(question, self.normal)
        # the 4 answers are added as a numbered list
        cell_11 = ListFlowable([Paragraph(s, self.normal) for s in [choice_1, choice_2, choice_3, choice_4]],
                                bulletType="a", bulletFormat="%s)")
        if image:                                   # is there a picture in the questions?
            try:                                    # yes
                file1 = open(f"image/{image}")      # check if the image file can be found or not
                file1.close()
                # the image dimension in Electrician trade written test is 4:3
                cell_12 = Image(f"image/{image}", width=3.2 * inch, height=2.4 * inch)
            except:
                cell_12 = Paragraph("找不到圖片!", self.big_char)    # image file not found
                print(f"image/{image} file not found!")                  # give error message
        else:
            cell_12 = ""
        # create a table for each question
        data = [[cell_00, cell_10, ""], ["", cell_11, cell_12]]
        t = Table(data, colWidths=(0.5*inch, 3*inch, 3.5*inch))
        t.setStyle(TableStyle([('ALIGN', (0, 0), (2, 1), 'LEFT'),
                               ('VALIGN', (0, 0), (2, 1), 'TOP'),
                               ('SPAN', (1, 0), (2, 0)),])
                   )
        # append table to the file contents
        self.story.append(t)
        # add vertical space after each question
        self.story.append(Spacer(1, 0.3*inch))


    def close(self):

        # specify the page footer
        def myPages(canvas, doc):
            canvas.saveState()
            canvas.setFont('Times-Roman', 10)
            canvas.drawString(4 * inch, 0.5 * inch,  "p. %d" % (doc.page))
            canvas.restoreState()

        # write the whole question paper to PDF file
        self.story.append(Paragraph("--- 完 ---", self.big_char))
        self.doc.build(self.story, onFirstPage=myPages, onLaterPages=myPages)
