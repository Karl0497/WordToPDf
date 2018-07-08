import sys
import os
import win32com.client
def validate(full_path):

    dot = full_path.rfind(".")
    if dot==-1: return None
    extension = full_path[dot+1:]
    if extension not in ["docx","doc"]: return None
    slash = full_path.rfind("\\")
    return full_path[slash+1:dot], full_path[:slash+1]

def main():
    wdFormatPDF = 17
    full_path =  os.path.abspath(sys.argv[1])
    print(full_path)
    res = validate(full_path)
    
    if not res:
        print("Invalid file")
        return
    file_name,path=res
    word = win32com.client.Dispatch('Word.Application')
    out_file = (path+file_name+".pdf")
    print(out_file)
    doc = word.Documents.Open(full_path)
    doc.SaveAs(out_file, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()
    
try:
    main()
    print("done")
    input()
except Exception as e:
    print(e)
    input()
