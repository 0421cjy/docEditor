import os
import olefile
import zlib
import struct
from docx import Document
from docx.shared import RGBColor
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def editHwpFiles(name, dic):
    filenames = os.listdir(os.getcwd())

    for filename in filenames:
        full_filename = os.path.join(name, filename)
        ext = os.path.splitext(full_filename)[-1]
        
        if (ext == '.hwp'):
            # print(full_filename)
            
            hwp = olefile.OleFileIO(full_filename)
            dirs = hwp.listdir()
            
            header = hwp.openstream("Fileheader")
            header_data = header.read()
            is_compressed = (header_data[36] & 1) == 1
            
            nums = []
            for d in dirs:
                if d[0] == "BodyText":
                    nums.append(int(d[1][len("Section"):]))
                    
            sections = ["BodyText/Section"+str(x) for x in sorted(nums)]
            
            text = ""
            
            for section in sections:
                bodytext = hwp.openstream(section)
                data = bodytext.read()
                if is_compressed:
                    unpacked_data = zlib.decompress(data, -15)
                else:
                    unpacked_data = data
            
                # 각 Section 내 text 추출    
                section_text = ""
                i = 0
                size = len(unpacked_data)
                while i < size:
                    header = struct.unpack_from("<I", unpacked_data, i)[0]
                    rec_type = header & 0x3ff
                    rec_len = (header >> 20) & 0xfff

                    if rec_type in [67]:
                        rec_data = unpacked_data[i+4:i+4+rec_len]
                        section_text += rec_data.decode('utf-16')
                        section_text += "\n"

                    i += 4 + rec_len

                text += section_text
                text += "\n"

            print(text)

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

    for filename in filenames:
        full_filename = os.path.join(name, filename)
        ext = os.path.splitext(full_filename)[-1]
        
        if (ext == '.docx'):
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

            editname = filename.replace('.docx', '')
            resultname = editname + '_수정본.docx'
            
            document.save(resultname)
    
if __name__ == "__main__":
    
    myDictionary = {}
    dirname = os.getcwd()
    notepath = dirname + '\\note.txt'
    
    readTxtFile(notepath, myDictionary)
    editWordFiles(dirname, myDictionary)
    editHwpFiles(dirname, myDictionary)