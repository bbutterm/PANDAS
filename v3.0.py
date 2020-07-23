import pandas as pd
import xlrd as rd
import xlsxwriter
p_file = input("input file name:")
data = pd.read_csv(p_file, sep=',', error_bad_lines=False, index_col=False, dtype='unicode')
writer = pd.ExcelWriter('EIM.xlsx', engine='xlsxwriter')
#data.to_excel(writer,'data')
#writer.save()
#data = pd.read_excel('file.xlsx')
audit = pd.read_excel('audit.xlsx')
data_conso = data[data['Changed Using'].str.contains('CONSO')]
data_pg_prefilter = data[data['Changed Using'].str.contains('TO PRODUCT GROUP')]
data_pg = data_pg_prefilter[data_pg_prefilter['New Dispatch Flag'].str.contains('NOT DISPATCH')]
audited = data_pg.merge(audit,on='New Product Group Description')
data_conso.to_excel(writer, 'conso')
data_pg.to_excel(writer, 'pg')
audited.to_excel(writer, 'audited')
writer.save()
