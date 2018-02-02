import win32com
from win32com import client as wc

word = wc.Dispatch('Word.Application')


# 目标文件路径可以自由改动，大家注意SaveAs方法中的参数，好多啊，别写错了word = wc.Dispatch('Word.Application')
def do_save():
    doc = word.Documents.Open('F:/PycharmProjects/docx_demo/A/01-220kV母联开关由检修改冷备用.doc')  # 目标路径下的文件
    doc.SaveAs('F:/PycharmProjects/docx_demo/A/01-220kV母联开关由检修改冷备用.docx', 12, False, "", True, "", False, False, False,
               False)  # 转化后路径下的文件
    doc.Close()
    word.Quit()  # 转化为.docx文件后，在处理.docx文件，一路畅通无阻，网上很多解决方案，这里我就不详细说了，有问题，可以给我留言哟
