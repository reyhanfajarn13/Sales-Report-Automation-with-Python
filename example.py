import xlwings as xw
import os, sys
from docxtpl import DocxTemplate
import win32com.client as win32
import pandas as pd
import matplotlib.pyplot as plt
import smtplib
from docx2pdf import convert
from dotenv import load_dotenv
from email_data import *

from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
load_dotenv()

os.chdir(sys.path[0])


def create_barchart(df, barchart_output):
    """Group DataFrame by sub-category, plot barchart, save plot as PNG"""
    top_products = df.groupby(by=df["Region"]).sum()[["Total Sales"]]
    top_products = top_products.sort_values(by="Total Sales")
    plt.rcParams["figure.dpi"] = 500
    plot = top_products.plot(kind="barh")
    fig = plot.get_figure()
    fig.savefig(barchart_output, bbox_inches="tight")
    return None

def create_piechart(df, piechart_output):
    top_retail = df.groupby(by=df["Retailer"]).sum()[["Total Sales"]]
    top_retail = top_retail.sort_values(by="Total Sales")
    
    
    df_top_retail = top_retail.reset_index()
    plt.figure(figsize=(8, 8))
    plt.subplots(subplot_kw=dict(aspect="equal"))
    colors = plt.cm.Paired(range(len(df)))

    plt.pie(df_top_retail['Total Sales'], labels=df_top_retail['Retailer'], autopct='%1.1f%%', startangle=140, colors=colors, pctdistance=0.85)
    centre_circle = plt.Circle((0,0),0.70,fc='white')
    fig = plt.gcf()
    fig.gca().add_artist(centre_circle)
    plt.title('Sales Distribution by Region')
    plt.axis('equal')
    

    # plt.legend(df_top_retail['Retailer'], title='Retailer',loc="upper right", fontsize=8, bbox_to_anchor=(1, 0.5))
    plt.savefig(piechart_output, bbox_inches="tight")


def convert_to_pdf(doc):

    convert(doc,f"{doc}.pdf")
    # word = win32.DispatchEx("Word.Application")
    # full_path = os.path.abspath(doc)

    # if not os.path.exists(full_path):
    #     print(f"Document not found: {full_path}")
    #     return None

    # new_name = full_path.replace(".docx", ".pdf")
    # worddoc = word.Documents.Open(full_path)
    # worddoc.Close()
    # worddoc.SaveAs(new_name, FileFormat=17)
   
    return None

def send_email(file):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    smtp_username = 'reyhanfajarnst13@gmail.com'
    smtp_password = os.environ['SMTP_PASSWORD']

    from_email = 'reyhanfajarnst13@gmail.com'
    to_email = 'reyhanfajarn13@gmail.com'
    subject = email_subject()
    body = email_body()

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body))

    with open(f'{file}.pdf', 'rb') as f:
        attachment = MIMEApplication(f.read(), _subtype='pdf')
        attachment.add_header('Content-Disposition', 'attachment', filename=f'{file}.pdf')
        msg.attach(attachment)

    with smtplib.SMTP(smtp_server, smtp_port) as smtp:
        smtp.starttls()
        smtp.login(smtp_username, smtp_password)
        smtp.send_message(msg)
    
    return None

wb = xw.Book('Adidas US Sales Datasets.xlsx')
sht_panel = wb.sheets

workingsheet = sht_panel[1]
base_data = sht_panel[3]

context = workingsheet.range('A1').options(dict, expand='table', numbers=int).value
df = base_data.range("B5").options(pd.DataFrame, index=False, expand="table").value


barchart_name = "sales_by_region"
barchart_output = f"{barchart_name}.png"
create_barchart(df, barchart_output)

piechart_name = "sales_by_retail"
piechart_output = f"{piechart_name}.png"
create_piechart(df, piechart_output)

doc = DocxTemplate('Laporan_penjualan_template.docx')

doc.replace_pic("Placeholder", barchart_output)
doc.replace_pic("Placeholder_2", piechart_output)

output_name = f'Laporan_Penjualan_{context["tahun"]}.docx'
doc.render(context=context)
doc.save(output_name)

convert_to_pdf(str(output_name))

send_email(str(output_name))

    
