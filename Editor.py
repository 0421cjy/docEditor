import os
from docx import Document
from docx.shared import RGBColor
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import win32com.client as win32
from shutil import copyfile

def readTxtFile(path, dic):
    file = open(path, "r", encoding='UTF-8')
    
    while True:
        line = file.readline()
        if not line:
            break
        
        raw_res = line.strip()
        pair = raw_res.split(':')
        
        if pair[0] == '':
            continue
        
        # print(pair[0], pair[1])
        dic[pair[0]] = pair[1]

    file.close()
    
def readExcelFile(path, dic):
    file = open(path, "r", encoding='UTF-8')
    
    while True:
        line = file.readline()
        if not line:
            break
        
        raw_res = line.strip()
        pair = raw_res.split(':')
        
        # print(pair[0], pair[1])
        dic[pair[0]] = pair[1]

    file.close()

def editHwpFiles(name, dic):
    filenames = os.listdir(os.getcwd())
    fileformat = '.hwp'

    for filename in filenames:
        full_filename = os.path.join(name, filename)
        ext = os.path.splitext(full_filename)[-1]
        
        if (ext == '.hwp'):
            # print(full_filename)
            
            hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            
            editname = full_filename.replace(fileformat, '')
            resultname = editname + '_수정본' + fileformat
            
            copyfile(full_filename, resultname)
            
            hwp.Open(resultname)
            hwp.InitScan()
            
            for i in dic:
                hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
                option=hwp.HParameterSet.HFindReplace
                option.FindString = i
                option.ReplaceString = dic[i]
                option.ReplaceCharShape.TextColor = hwp.RGBColor(255, 0, 0)
                #option.ReplaceCharShape.Bold = 1
                option.IgnoreMessage = 1
                hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
                    
            hwp.ReleaseScan()
            hwp.Clear(3)
            hwp.Quit()
    
def editWordFiles(name, dic):
    filenames = os.listdir(os.getcwd())
    fileformat = '.docx'

    for filename in filenames:
        full_filename = os.path.join(name, filename)
        ext = os.path.splitext(full_filename)[-1]
        
        if (ext == fileformat):
            # print(full_filename)

            document = Document(full_filename)

            for p in document.paragraphs:
                for run in p.runs:
                    for i in dic:
                        if run.text.find(i) >= 0:
                            text = run.text.replace(i, dic[i])
                            run.text = text
                            run.font.color.rgb = RGBColor(255, 0, 0)

            editname = filename.replace(fileformat, '')
            resultname = editname + '_수정본' + fileformat
            
            document.save(resultname)
            
def editRtfFiles(name, dic):
    filenames = os.listdir(os.getcwd())
    fileformat = '.rtf'

    for filename in filenames:
        full_filename = os.path.join(name, filename)
        ext = os.path.splitext(full_filename)[-1]
        
        if (ext == fileformat):
            word = win32.Dispatch("Word.Application")
            doc = word.Documents.Open(full_filename)
            
            editname = full_filename.replace(fileformat, '')
            resultname = editname + '_수정본' + '.docx'
            
            wdFormatDocumentDefault = 16
            
            doc.SaveAs(resultname, FileFormat=wdFormatDocumentDefault)
            doc.Close()
            word.Quit()
            
            document = Document(resultname)
            
            for p in document.paragraphs:
                for run in p.runs:
                    for i in dic:
                        if run.text.find(i) >= 0:
                            text = run.text.replace(i, dic[i])
                            run.text = text
                            run.font.color.rgb = RGBColor(255, 0, 0)
            
            document.save(resultname)
            os.remove(resultname)
            
            editname = full_filename.replace(fileformat, '')
            resultname = editname + '_수정본' + '.rtf'
            document.save(resultname)
    
if __name__ == "__main__":
    
    myDictionary = {}
    dirname = os.getcwd()
    notepath = dirname + '\\note.txt'
    
    readTxtFile(notepath, myDictionary)
    #readExcelFile(notepath, myDictionary)
    editWordFiles(dirname, myDictionary)
    editHwpFiles(dirname, myDictionary)
    editRtfFiles(dirname, myDictionary)