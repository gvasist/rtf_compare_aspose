#This is the working version of RTF Compare.
#This program generates rtf files with compare differences.
#GUI to select folders (folder1, folder2 and compare)

#Enhancements Tried in this version:
#Use ASPOSE word to compare differences b/w the files - This works and generates wonderful rtf outputs
#Next step is to build a GUI, count the number of differences, and give user options to exclude header/footer.


import difflib
import os
import re
from difflib import Differ
from pathlib import Path
from difflib import HtmlDiff
from tkinter import *
from tkinter import filedialog
import aspose.words as aw
import datetime

root = Tk()

compare_options = aw.comparing.CompareOptions()
compare_options.ignore_formatting = False
compare_options.ignore_case_changes = False
compare_options.ignore_comments = False
compare_options.ignore_tables = False
compare_options.ignore_fields = False
compare_options.ignore_footnotes = False
compare_options.ignore_textboxes = False
compare_options.ignore_headers_and_footers = False
compare_options.target = aw.comparing.ComparisonTargetType.NEW

def read_folder():
    root.foldername = filedialog.askdirectory(initialdir="C:/Users/gokul.vasist/MyDocs/Official/Python/Programs")
    # my_label = Label(root,text=root.filename).pack()
    return root.foldername

def read_rtf_files(input_file):
    with open(os.path.join(input_file).replace("\\", "/")) as file_input:
        file_data  = file_input.read()
    return file_data


def strip_rtf_content(file_data):
    try:
        import striprtf
    except:
        get_ipython().system('pip install striprtf')

    from striprtf.striprtf import rtf_to_text
    text = rtf_to_text(file_data)
    return text

firstfolder = read_folder()
print(firstfolder)

secondfolder = read_folder()
print(secondfolder)

comparefolder = read_folder()
print(comparefolder)


fol_1 = []
fol_2 = []
for fname in os.listdir(path=firstfolder):
    fol_1.append(fname)
for fname in os.listdir(path=secondfolder):
    fol_2.append(fname)
for i in fol_1:
    for j in fol_2:
        print("First file is: ", i)
        print("Second file is: ", j)
        print("First folder and file location is: ", os.path.join(firstfolder, i).replace("\\", "/"))

        if(i == j):
           filename = i[0:-4]
           docA = aw.Document(os.path.join(firstfolder,i))
           docB = aw.Document(os.path.join(secondfolder,j))

           docA.accept_all_revisions()
           docB.accept_all_revisions()

           docA.compare(docB, "Yourself",datetime.time())
           #docA.save("Output.rtf")
           docA.save(os.path.join(comparefolder, filename+'_compare.rtf'))

