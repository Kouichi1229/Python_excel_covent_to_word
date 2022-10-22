import openpyxl as xl, win32com.client, PIL, os, sys, random, datetime, matplotlib.pyplot as plt
from openpyxl.chart import LineChart, Reference
from PIL import ImageGrab, Image
from docx.shared import Cm
from docx import Document
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Inches, Mm, Emu

import Collage_Paste_action





File_name = '台大院級1028(轉檔範例).xlsx' # 輸入檔案名稱



for i in range(0,7):
        if i==0:
            Collage_Paste_action.table_one(File_name,Collage_Paste_action.collage_table_name[i])
        elif i==1:
            Collage_Paste_action.table_two(File_name,Collage_Paste_action.collage_table_name[i])
        elif i==2:
            Collage_Paste_action.table_three(File_name,Collage_Paste_action.collage_table_name[i])
        elif i==3:
            Collage_Paste_action.table_four(File_name,Collage_Paste_action.collage_table_name[i])
        elif i==4:
            Collage_Paste_action.table_five(File_name,Collage_Paste_action.collage_table_name[i])
        elif i==5:
            Collage_Paste_action.table_six(File_name,Collage_Paste_action.collage_table_name[i])
        else :
            Collage_Paste_action.table_seven(File_name,Collage_Paste_action.collage_table_name[i])

