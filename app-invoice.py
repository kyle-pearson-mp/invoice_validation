# -*- coding: utf-8 -*-
"""
Created on Fri Jun 27 14:54:19 2025

@author: KylePearson
"""

import streamlit as st
import pandas as pd
import pdfplumber
import os
import re
import zipfile
import tempfile
#%%

st.title('Xcel Invoice Extraction')

df = pd.DataFrame()
folder_path = r"C:\Users\KylePearson\Downloads\Fw_ XCL515 Invoicing Review - May"
def add_row(df, contract,month, invoice_number, total_amount, po_num, po_num2, filename_po):
    new_row = {'Invoice': invoice_number,
               'PO': po_num,
               'Contract': contract,
               'Month': month,
               'Total $': total_amount,
               'Month2': month,
               'PO2': po_num2,
               'Filename PO': filename_po}
    new_row = pd.DataFrame([new_row])
    df = pd.concat([new_row, df])
    return df

def extract_invoice_data(pdf, df):
    print(pdf)
    contract = " ".join(pdf.split(' ')[2:-2])
    month = pdf.split(' ')[-2]
    filename_po = re.search(r'PO\s*No\.\s*(\d+)', pdf, re.IGNORECASE).group(1).strip()
    print('Filename PO: ', filename_po)
    with pdfplumber.open(pdf) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                match = re.search(r'\b\d{1,2}/\d{1,2}/\d{4}\s+(\d{5})\b', text)
                if match:
                    invoice_number = match.group(1)
                    #st.write("Invoice Number:", invoice_number)
    
                match = re.search(r'TOTAL\s*\$\s*([\d,]+\.\d{2})', text, re.IGNORECASE)
                if match:
                    total_amount = match.group(1)
                    #print("Extracted amount:", total_amount)
    
                match = re.search(r"Project\s*Title[:\s]*(.+)", text, re.IGNORECASE)
                    
                match = re.search(r'PO\s*No\.\s*(\d+)', text, re.IGNORECASE)
                if match:
                    #data["PO_Number"] = match.group(1).strip()
                    po_num = match.group(1).strip()
                    #print("PO Number: ", po_num)
                match = re.search(r'\b\d{1,2}/\d{1,2}/\d{4}\s+(\d{10})\b', text)
                if match:
                    #data["PO_Number2"] = match.group(1).strip()
                    po_num2 = match.group(1).strip()
                    #print("PO Number2: ", po_num2)
                if filename_po != po_num or filename_po != po_num2 or po_num != po_num2:
                    print('Mismatch PO numbers')
                    st.write(f'Filename PO: {filename_po}, First PO: {po_num}, Second PO: {po_num2}')
    try:
        df = add_row(df, contract, month, invoice_number, total_amount, po_num, po_num2, filename_po)
    except:
        st.write(text)
        df = add_row(df, contract, month, invoice_number, None, po_num, po_num2, filename_po)
    return df

st.title("Upload a ZIP of Invoice PDFs")
uploaded_zips = st.file_uploader("Upload ZIP file", type="zip", accept_multiple_files=True)

if uploaded_zips:
    for uploaded_zip in uploaded_zips:
        st.subheader(f"Processing: {uploaded_zip.name}")

        with tempfile.TemporaryDirectory() as tmpdir:
            zip_path = os.path.join(tmpdir, "uploaded.zip")
            
            # Save uploaded zip to temp directory
            with open(zip_path, "wb") as f:
                f.write(uploaded_zip.read())
    
            # Extract zip contents
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(tmpdir)
                st.success("ZIP file extracted!")      
                pdf_files = []
                for root, dirs, files in os.walk(tmpdir):
                    for file in files:
                        if file.lower().endswith(".pdf"):
                            pdf_files.append(os.path.join(root, file))
                # List PDF files
                #pdf_files = [f for f in os.listdir(tmpdir) if f.lower().endswith(".pdf")]
                #st.write("Found PDF files:")
                for pdf_path in pdf_files:
                    #pdf_path = os.path.join(tmpdir, pdf)
                    df = extract_invoice_data(pdf_path, df)
                st.success("Invoice Data extracted!")
                

st.subheader('Invoice data')
st.write(df)