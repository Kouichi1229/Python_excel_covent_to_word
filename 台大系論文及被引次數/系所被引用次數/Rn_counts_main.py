import openpyxl as xl, win32com.client, PIL, os, sys, random, datetime, matplotlib.pyplot as plt
from openpyxl.chart import LineChart, Reference
from PIL import ImageGrab, Image
from docx.shared import Cm
from docx import Document
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Inches, Mm, Emu
import paste_action 

File_name = '各系所被引次數統計.xlsx'

for i in range(0,16):
    paste_action.table_two(File_name,paste_action.table_name[i],paste_action.table_name[i])