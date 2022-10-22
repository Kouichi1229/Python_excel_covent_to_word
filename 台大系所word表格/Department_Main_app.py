import openpyxl as xl, win32com.client, PIL, os, sys, random, datetime, matplotlib.pyplot as plt
from openpyxl.chart import LineChart, Reference
from PIL import ImageGrab, Image
from docx.shared import Cm
from docx import Document
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Inches, Mm, Emu
import Department_Paste_action



File_name = '台大系所1028(轉檔範例)拿掉公式.xlsx' # 輸入檔案名稱


for i in range(0,5):
        if i==0:
            Department_Paste_action.table_one(File_name,Department_Paste_action.collage_table_name[i])
        elif i==1:
            Department_Paste_action.table_two(File_name,Department_Paste_action.collage_table_name[i])
        elif i==2:
            Department_Paste_action.table_three(File_name,Department_Paste_action.collage_table_name[i])
        elif i==3:
            Department_Paste_action.table_four(File_name,Department_Paste_action.collage_table_name[i])
        else:
            Department_Paste_action.table_five(File_name,Department_Paste_action.collage_table_name[i])
        
        
