from docx import Document
import Paragraph
import PreSet
import Word_to_pdf

document = Document('Example.docx')
PreSet.Fontset(document)
time = Paragraph.paragraph(document)
document.save('技术文章：《每天30秒电力热讯 ' + time + '》.docx')
Word_to_pdf.Word_to_Pdf('技术文章：《每天30秒电力热讯 ' + time + '》.docx')
