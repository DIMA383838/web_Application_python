import pandas as pd

file_name = input('Введите имя Excel-файла для исправления (с расширением): ').strip()
if not file_name:
    raise SystemExit('Имя файла не указано.')
df =pd.read_excel(file_name)
df['ФИО']=df['ФИО'].str.replace('КОРНЕЕВ А.В.','КОРНЕВ А.В.',regex=False)
df['ФИО']=df['ФИО'].str.replace('ГУДУЛИН Д.А.','ГУДИЛИН Д.А.',regex=False)
df['ФИО']=df['ФИО'].str.replace('Останко','ОСТАНКО Д.А.',regex=False)
df.to_excel(file_name,index=False)
print("Готово")
