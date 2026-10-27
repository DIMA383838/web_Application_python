import pandas as pd
import docx
import os
from glob import glob

 
def to_hours(v):
    if not v or v.strip()=='':
        return 0.0
    v=v.strip().replace(',','.')
    if ':' in v:
        p=v.split(':')
        h=float(p[0])+(float(p[1])/60)+(float(p[2])/3600 if len(p)>2 else 0)
        return round(h,2)
    return float(v)
    

def process_doc_file(input_file):
 name_output_file=os.path.splitext(input_file)[0]+'.xlsx'

 if not name_output_file:
    name_output_file="Our format"

   
 doc=docx.Document(input_file)

 total_tables=len(doc.tables)

 if total_tables==7:
    table_indices=[1,3,5]

 if total_tables==5:
    table_indices=[1,3]

 if total_tables==3:
    table_indices=[1]

 if not table_indices:
    table_indices=[1]
    print(f"Будет использован индекс апо умолчанию: {table_indices}")



 data = []

 for table_index in table_indices:
    table = doc.tables[table_index]

    headers=[]

    for cell in table.rows[0].cells:
        headers.append(cell.text.strip())

    index={}

    for i, name in enumerate(headers):
        index[name]=i



    for row in table.rows[1:]:
    
        
        cells = row.cells
    
        time_column= None
        for col_name in index.keys():
            if 'Время ЦП' in col_name:
                time_column=col_name
                break
    
        fio = cells[index.get('Исполнитель')].text.strip() if 'Исполнитель' in index else ' '
    
        metodika = cells[index.get('Методика')].text.strip() if 'Методика' in index else ' '
    
        start = cells[index.get('Начало счета')].text.strip() if 'Начало счета' in index else ' '
    
        end = cells[index.get('Окончание счета')].text.strip() if 'Окончание счета' in index else ' '
    
        yadra = cells[index.get('Число проц.ядер')].text.strip() if 'Число проц.ядер' in index else ' '
    
        #yadrahours = to_hours(cells[index.get('Время ЦП, проц.ядер*ч.')].text.strip() if 'Время ЦП, проц.ядер*ч.' in index else ' ') 
        yadrahours = to_hours(cells[index[time_column]].text.strip() if time_column in index else ' ') 
    
        comment = cells[index.get('Комментарий')].text.strip() if 'Комментарий' in index else ' '
        
        if table_index==1:
            resourse ="VCKP_3"
        if table_index==3:
            resourse="VCKP_4"
    
        data.append([metodika,fio,start,end,yadra,yadrahours,comment,resourse])
    
        df=pd.DataFrame(data, columns=['Методика','ФИО','Начало счёта','Окончание счёта','Число проц.ядер','Время ЦП, проц.ядер/ч','Коментарий','Ресурс'])

 df.to_excel(name_output_file, index=False)
 print("Готово, Файл сохранён как: "+name_output_file)
        
def main():
    #current_dir=os.getcwd() # Использовать, если нужна дерикторая, в которой лежит скрипт
    current_dir=os.getcwd()  # Рабочая папка задаётся приложением/текущей директорией
    doc_files=glob(os.path.join(current_dir,"*.docx"))
    doc_files=sorted(set(doc_files))
    print(f"Найдено файлов для обработки: {len(doc_files)}")
    print("Убедитесь, что все файлы .docx закрыты")
    for doc_file in doc_files:
        print(f"Началась обработка файла: {doc_file}")
        process_doc_file(doc_file)
    print("Все файлы .docx в директории обработаны. Работа завершена.")
        
    

if __name__=="__main__":
    main()
   
    
    