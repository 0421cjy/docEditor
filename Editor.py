import os
from docx import Document
from docx.shared import RGBColor
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import win32com.client as win32
from shutil import copyfile

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
                option.ReplaceCharShape.Bold = 1
                option.IgnoreMessage = 1
                hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
                    
            hwp.ReleaseScan()
            
            hwp.Clear(3)
            hwp.Quit()
            
def readTxtFile(path, dic):
    file = open(path, "r", encoding='UTF-8')
    while True:
        line = file.readline()
        if not line:
            break
        
        raw_res = line.strip()
        pair = raw_res.split(':')
        
        ## print(pair[0], pair[1])
        dic[pair[0]] = pair[1]

    file.close()
    
def editWordFiles(name, dic):
    filenames = os.listdir(os.getcwd())
    fileformat = '.docx'

    for filename in filenames:
        full_filename = os.path.join(name, filename)
        ext = os.path.splitext(full_filename)[-1]
        
        if (ext == fileformat):
            # print(full_filename)

            document = Document(full_filename)

            for i in dic:
                for p in document.paragraphs:
                    if p.text.find(i) >= 0:
                        index = p.text.index(i)
                        end_index = index + len(i)
                        
                        # print('index : ', index)
                        # print('end_index : ', end_index)
                        
                        if index == 0:
                            rest_context = p.text.replace(i, '')
                            
                            p.text = ''
                        
                            new_run = p.add_run(dic[i])
                            font = new_run.font
                            font.color.rgb = RGBColor(255, 0, 0)
                            font.size = Pt(12)
                            
                            origin_run = p.add_run(rest_context)
                            font = origin_run.font
                            font.color.rgb = RGBColor(0, 0, 0)
                            font.size = Pt(12)
                            
                        if index != 0:
                            first_context = p.text[0:index]
                            middle_context = p.text[index:end_index]
                            last_context = p.text[end_index:]
                            
                            p.text = ''
                            
                            first_run = p.add_run(first_context)
                            font = first_run.font
                            font.color.rgb = RGBColor(0, 0, 0)
                            font.size = Pt(12)
                            
                            middle_context = middle_context.replace(i, dic[i])
                            
                            middle_run = p.add_run(middle_context)
                            font = middle_run.font
                            font.color.rgb = RGBColor(255, 0, 0)
                            font.size = Pt(12)
                            
                            last_run = p.add_run(last_context)
                            font = last_run.font
                            font.color.rgb = RGBColor(0, 0, 0)
                            font.size = Pt(12)

            editname = filename.replace(fileformat, '')
            resultname = editname + '_수정본' + fileformat
            
            document.save(resultname)
    
if __name__ == "__main__":
    
    myDictionary = {}
    dirname = os.getcwd()
    notepath = dirname + '\\note.txt'
    
    readTxtFile(notepath, myDictionary)
    editWordFiles(dirname, myDictionary)
    editHwpFiles(dirname, myDictionary)