# import xlwings as xw
# import os, sys
# from docxtpl import DocxTemplate

# os.chdir(sys.path[0])

# # def main():
# #     wb = xw.Book('Aidas US Sales Datasets.xlsx')
# #     sht_panel = wb.sheets['final']
# #     doc = DocxTemplate('Laporan_penjualan_template.docx')

# #     context = sht_panel.range('A2').options(dict, expand='table', numbers=int).value

# #     output_name = f'Laporan_Penjualan_{context["tahun"]}.docx'
# #     doc.render(context=context)
# #     doc.save(output_name)
