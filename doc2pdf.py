import os
from win32com import client as wc  # 导入模块
import PyPDF2

word = wc.Dispatch("Word.Application")  # 打开word应用程序


def saveaspdf(source_file_list):
    failed = []
    success = 0
    total = len(source_file_list)
    for file in source_file_list:
        print(f"\r {success + 1}/{total}\t tranfering ：{file}", end="")
        try:
            doc = word.Documents.Open(file)  # 打开word文件
            pdfFileName = "{}.pdf".format(file.replace(".docx", '').replace(".doc", ''))
            doc.SaveAs(pdfFileName, 17)  # 另存为后缀为".docx"的文件，其中参数12指docx文件
            doc.Close()  # 关闭原来word文件
            success += 1
        except Exception as e:
            failed.append(file)
    word.Quit()
    print(f"\r ALL WORD FILES: {total}, SUCCESS: {success}, FAILED : {len(failed)}, {failed}。", end='')


def mergePdfs(pdf_list):
    if "Merged.pdf" in pdf_list:
        raise Exception("Error! Merged.pdf are EXIST now! please check whether you have the merged files already")
        os.remove("./" + 'Merged.pdf')
        pdf_list.remove("Merged.pdf")
    else:
        opened_file = [open("./" + file_name, 'rb') for file_name in pdf_list]
        pdfFM = PyPDF2.PdfFileMerger()
        for file in opened_file:
            pdfFM.append(file)
        # output the file.
        with open(".\\Merged.pdf", 'wb') as write_out_file:
            pdfFM.write(write_out_file)

        # close all the input files.
        for file in opened_file:
            file.close()
        print('all pdf files merged successfully, and saved as merged.pdf')


if __name__ == '__main__':
    # current_path = input("Input MS-WORD file path:")
    current_path = "C:/Users/HP/Desktop/互联网撰写材料+"
    if current_path == "":
        print("You did not enter the path, using current directory as default.")
        # 原文件夹路径
        current_path = os.getcwd() + "/"
    else:
        current_path += "/"
    print(f"current_path is {current_path}")
    source_file_list = [x for x in os.listdir(current_path) if x.endswith(".docx")]
    print(source_file_list)
    # 获取源文件夹内文件列表
    source_file_list_all = []
    for file in source_file_list:
        source_file_list_all.append(current_path + file)
    saveaspdf(source_file_list_all)
    mergeOrNot = input(
        'Do you need merge all pdf files to single pdf？notice: You may get the wrong file order. If you need merge all pdf file, press "Enter" to cotinue. otherwise you can close this window directly. ')
    pdf_list = sorted([x for x in os.listdir(current_path) if x.endswith(".pdf")])
    print(pdf_list)
    mergePdfs(pdf_list)
    delSinglePDF = input(
        'Do you need every single PDF files which already merged in Merged.pdf? DELETE all the single pdf files ? (y/n): ')
    if str.upper(delSinglePDF) == "Y":
        for pdf in pdf_list:
            os.remove(pdf)
        print("all the single pdf files DELETEDE.")