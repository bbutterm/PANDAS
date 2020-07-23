import pandas as pd
import xlrd as rd
import xlsxwriter
data = pd.read_excel('file.xlsx')
data_conso = data[data['Changed Using'].str.contains('CONSO')]
data_pg_prefilter = data[data['Changed Using'].str.contains('TO PRODUCT GROUP')]
data_pg = data_pg_prefilter[data_pg_prefilter['New Dispatch Flag'].str.contains('NOT DISPATCH')]
writer = pd.ExcelWriter('edited.xlsx', engine='xlsxwriter')
data_conso.to_excel(writer, 'conso')
data_pg.to_excel(writer, 'pg')
writer.save()