# -*- coding: utf-8 -*-
"""
Created on Sun Sep  5 22:23:33 2021

@author: lenovo
"""

from __future__ import print_function
import streamlit as st
from mailmerge import MailMerge
import datetime
from datetime import timedelta
import os
from docxcompose.composer import Composer
from docx import Document

name = ["STADTKIND QUINOA", "TRÜFFEL-TRAUM", "SOMMER 21", "SRIRACHA NUDEL", "KICHERERBSEN CAESAR", "HUMMUS PESTO", "GOOD BALLS"]
weight = ['350 g', '370 g','310 g', '380 g', '290 g', '400 g', '280 g']
price = ["€ 5,95", "€ 5,95", "€ 5,95", "€ 4,95", "€ 4,95", "€ 5,95", "€ 3,45"]
per100 = ["€ 1,70 / 100g", "€ 1,86 / 100g", "€ 2,05 / 100g", "€ 1,46 / 100g", "€ 2,11 / 100g", "€ 1,57 / 100g", "€ 1,23 / 100g"]
#quantities = [1,0,0,3,0,0,1]

path = os.getcwd()
template_list = [path + r"\template_{}.docx".format(bowlname) for bowlname in name]
datestr = str(datetime.date.today())
date2 = datetime.datetime.strptime(datestr, '%Y-%m-%d').date()

st.sidebar.markdown('#### Enter quantities for each bowl')
quantities=[0,0,0,0,0,0,0]
for quant in range(len(quantities)):
    quantities[quant] = st.sidebar.number_input(name[quant])

for indice,nombre in enumerate(name):
 if nombre == "GOOD BALLS" and quantities[indice] != 0:
     document = MailMerge(template_list[indice])
     document.merge(
         MHD = str(date2 + timedelta(days=3)),
         NAME = nombre,
         PRICE = price[indice],
         PRICE100 = per100[indice],
         WEIGHT = weight[indice])
     quantity = quantities[indice]
     document.write(path + r'\tempfiles\{}.docx'.format(nombre))
 elif quantities[indice] !=0:
     document = MailMerge(template_list[indice])
     document.merge(
         NAME = nombre,
         MHD = str(date2 + timedelta(days=2)),
         PRICE = price[indice],
         PRICE100 = per100[indice],
         WEIGHT = weight[indice])
     quantity = quantities[indice]
     document.write(path + r'\tempfiles\{}.docx'.format(nombre))
 else:
     pass

typeprimero = quantities.index(next(filter(lambda x: x!=0, quantities)))
master = Composer(Document(path + r"\tempfiles\{}.docx".format(name[typeprimero])))
quantities[typeprimero]=quantities[typeprimero]-1
if quantities[typeprimero] == 0:
    os.remove(path + r"\tempfiles\{}.docx".format(name[typeprimero]))
for indice,quant in enumerate(quantities):
    if quant !=0:
        label = Document(path + r"\tempfiles\{}.docx".format(name[indice]))
        for q in range(quant):
            master.append(label)
        os.remove(path + r"\tempfiles\{}.docx".format(name[indice]))
master.save(path + r"\LABELS_TO_PRINT_{}.docx".format(date2))
