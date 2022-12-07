
import pandas as pd
from docxtpl import DocxTemplate

from win32com import client
import time
import os

word_app = client.Dispatch("Word.Application")
data_frame = pd.read_excel('studentdata.xlsx');
for r_index, row in data_frame.iterrows():
    student_name= row['Name']
    
    # create doc file from excel data
    
    tpl = DocxTemplate("MyTemplate.docx")
    df_to_docT = data_frame.to_dict()
    x = data_frame.to_dict(orient='records')
    context = x
    tpl.render(context[r_index])
    tpl.save('Doc\\'+student_name+".docx")
    
    time.sleep(1)
    ROOT_DIR =os.path.dirname(os.path.abspath(__file__))
    print(ROOT_DIR)
    #create pdf from word
    
    doc= word_app.Documents.Open(ROOT_DIR+'\\Doc\\'+student_name+'.docx')
    print('Exporting pdf')
    doc.SaveAs(ROOT_DIR+'\\PDF\\'+student_name+'.pdf', FileFormat=17)
    
    
    
word_app.Quit()