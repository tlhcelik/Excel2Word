# Program to extract a particular row value
#python-docx, xlrd, python-tk
#_*_ coding:utf-8 _*_
import xlrd
from docx import Document
import Tkinter as tk
import tkMessageBox
import tkFileDialog
import sys
reload(sys)
sys.setdefaultencoding('utf8')

def readExcelAndWriteDoc(path):

  loc = (path)
   
  wb = xlrd.open_workbook(loc)
  status_label['text'] = "Excel dosyasi okunuyor"
  sheet = wb.sheet_by_index(0)
  sheet.cell_value(0, 0)

  document = Document()

  juriSayisi = txt_juri_sayisi.get(1.0, 2.0)
  juriSayisi.replace(" ", "")
  
  if (not juriSayisi):
    tkMessageBox.showinfo("Hata", "Juri Sayisi Giriniz.")
    return
  
  intJuriSayisi = int(juriSayisi)
  intJuriSayisi = intJuriSayisi + 2
  counter = 0
  baskan_cell = 1
  juri_cell = 1

  table = document.add_table(rows=sheet.nrows + juri_cell + baskan_cell + 2, cols=sheet.ncols + baskan_cell +intJuriSayisi + 2)

  for i in range(0, sheet.nrows):
    for j in range(0, sheet.ncols):
      table.cell(i, j).text = str(sheet.cell_value(i,j))


  table.cell(0, sheet.ncols + 1 ).text = str("Baskan")
  for j in range (sheet.ncols + 2, sheet.ncols + intJuriSayisi):
    counter += 1
    table.cell(0, j).text = str("Juri " + str(counter))

  counter += 2
  table.cell(0, sheet.ncols + counter ).text = "Toplam"
  table.cell(0, sheet.ncols + counter + 1).text = "Ortalama"

  table.cell(sheet.nrows + 1 , 1).text = str("Baskan")
  for i in range(2, intJuriSayisi):
    table.cell(sheet.nrows + 1, i).text = str("Juri " + str(i-1))



  save_path = loc[0:len(loc)-5]
  save_path += ".docx"
  status_label['text'] = "Word dosyasi {0} konumuna kaydedildi.".format(save_path)
  document.save(save_path)



def select_file():

  top.filename = tkFileDialog.askopenfilename(initialdir = "/",title = "Dosya Sec",filetypes = (("Excel dosyasi","*.xlsx"),("Excel","*.xlsx")))
  XLSX_FILE_PATH_AND_NAME = top.filename
  if (XLSX_FILE_PATH_AND_NAME != ""):
    status_label['text'] = XLSX_FILE_PATH_AND_NAME
    readExcelAndWriteDoc(XLSX_FILE_PATH_AND_NAME)
  else:
    status_label['text'] = "Dosya Secilmedi"



top = tk.Tk()
top.title("Excel2Word")
XLSX_FILE_PATH_AND_NAME = ""

btn_open = tk.Button(top, text ="Dosya Sec", command = select_file)
status_label = tk.Label(top, text="Excel Dosyasi Seciniz")

lbl_juri = tk.Label(top, text="Juri Sayisini Girin:")
txt_juri_sayisi = tk.Text(top, height=1, width=10)

lbl_juri.pack()
txt_juri_sayisi.pack()
status_label.pack()
btn_open.pack()
top.mainloop()


