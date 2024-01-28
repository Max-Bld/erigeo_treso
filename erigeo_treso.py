#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Jan 27 23:01:49 2024

@author: maxbld
"""

import os
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

cwd=os.getcwd()
fname_exc = "Tresorerie-mensuelle-2024.xlsx"

#INIT
init_sheet = pd.read_excel(f"{cwd}/{fname_exc}", sheet_name="Init")
init = init_sheet[init_sheet.columns[1]][0]
  

# DEPENSES
#%% columns name cleaning

depense = pd.read_excel(f"{cwd}/{fname_exc}", sheet_name="Dépenses")

col_names = list(depense.columns)

for n in range(len(col_names)):
    col_names[n]=col_names[n].lower()
    col_names[n]=col_names[n].replace(" €", "").replace("montant ", "").replace("réf. ", "").replace("é","e")
    
depense.columns = col_names

print(depense.columns)

#%% rows cleaning

depense = depense.drop(0)
depense = depense.reset_index(drop=True)

#%% Processing

depense["date"]=pd.to_datetime(depense["date"])

depense[["ht", "ttc", "tva"]] = - depense[["ht", "ttc", "tva"]]

# CREDITS
#%% columns name cleaning


credit = pd.read_excel(f"{cwd}/{fname_exc}", sheet_name="Crédits")

col_names = list(credit.columns)

for n in range(len(col_names)):
    col_names[n]=col_names[n].lower()
    col_names[n]=col_names[n].replace(" €", "").replace("montant ", "").replace("réf. ", "").replace("é","e")
    
credit.columns = col_names

print(credit.columns)

#%% rows cleaning

credit = credit.drop(0)
credit = credit.reset_index(drop=True)

#%% Processing

credit["date"]=pd.to_datetime(credit["date"])

for n in ['ht', 'ttc', 'tva']:
    try:
        credit[n] = credit[n].str.replace(',', '')
        credit[n] = credit[n].astype(float)
    except:
        continue
    
# TRESORERIE
#%% Merging cols

tresorerie = pd.concat([depense, credit])
tresorerie["nature"]=tresorerie["debit"].fillna("")+tresorerie["credit"].fillna("")
tresorerie.drop(['debit', 'credit'], axis=1, inplace=True)
tresorerie["client_fourn"]=tresorerie["fournisseur"].fillna("")+tresorerie["client"].fillna("")
tresorerie.drop(['client', 'fournisseur'], axis=1, inplace=True)

#%% Sorting by date

tresorerie.sort_values(by="date", inplace=True)
tresorerie = tresorerie.reset_index(drop=True)

#%%  Tresorerie col

tresorerie["tresorerie"] = tresorerie["ttc"].cumsum() + init

# DATA VIZ
#%%
def export_to_png(data, date_limite="2024-02-01", png_export_file_name="treso_evo_janv_2024.png"):
    specific_date = data[tresorerie["date"]<date_limite]

    plt.figure(figsize=(10,5))
    plt.suptitle("Évolution de la trésorerie (Janvier 2024)", fontsize=20)
    plt.plot(specific_date["date"], specific_date["tresorerie"], linewidth=4)
    plt.axhline(y = 0, color = 'grey', linestyle = '--') 
    plt.xlabel("Date")
    plt.ylabel("Montant de la trésorerie (€)")
    plt.grid(True)
    plt.savefig(png_export_file_name, dpi=300)
    
    return png_export_file_name
    
# specific_date = tresorerie[tresorerie["date"]<"2024-02-01"]

# plt.figure(figsize=(10,5))
# plt.suptitle("Évolution de la trésorerie (Janvier 2024)", fontsize=20)
# plt.plot(specific_date["date"], specific_date["tresorerie"], linewidth=4)
# plt.axhline(y = 0, color = 'grey', linestyle = '--') 
# plt.xlabel("Date")
# plt.ylabel("Montant de la trésorerie (€)")
# plt.grid(True)
# plt.savefig("treso_evo_janv_2024.png", dpi=300)

# EXPORT
#%%

# tresorerie["date"]=tresorerie["date"].dt.date
# depense["date"]=depense["date"].dt.date
# credit["date"]=credit["date"].dt.date

def export_to_excel(export_file_name="tresorerie_erigeo_python.xlsx", sheet1='Init', sheet2='Dépenses', sheet3='Crédits', sheet4='Concat_python'):
    with pd.ExcelWriter(export_file_name) as writer:  
        init_sheet.to_excel(writer, sheet_name=sheet1, index=False)
        depense.to_excel(writer, sheet_name=sheet2, index=False)
        credit.to_excel(writer, sheet_name=sheet3, index=False)
        tresorerie.to_excel(writer, sheet_name=sheet4, index=False)
        
    return export_file_name

# WEB APP
#%%
st.title("Trésorerie Erigéo")
st.header("Données", divider='blue')
if st.checkbox('Afficher les dépenses'):
    st.write(depense)
if st.checkbox('Afficher les crédits'):
    st.write(credit)
if st.checkbox('Afficher la concaténation'):
    st.write(tresorerie)
st.header("Graphique", divider='blue')
if st.checkbox('Afficher le graphique'):
    st.line_chart(data=tresorerie, x="date", y=["ttc", "ht", "tva", "tresorerie"])
st.header("Export", divider='blue')
if st.button("Export"):
    export_file_name=export_to_excel()
    st.write(f"Export du fichier *{export_file_name}* terminé.")
    png_export_file_name=export_to_png(tresorerie)
    st.write(f"Export du fichier *{png_export_file_name}* terminé.")
