# -*- coding: utf-8 -*-
"""
Created on Mon Nov 23 09:13:04 2020

@author: Collabcap
"""


import os
ma_pos=os.getcwd()
output_path = os.path.join(ma_pos,"json")

import pymongo
from pymongo import MongoClient

import json
import numpy as np
import pandas as pd


import swifter

import argparse

import codecs
import re

import unicodedata

from stop_words import get_stop_words

import warnings
warnings.simplefilter('ignore', UserWarning)

import time


#####################"

from multiprocessing import Pool, current_process



########################################"




DF = pd.ExcelFile( "list_mot_niv5.xlsx" )
tab_naf_code = pd.read_csv( "naf2008_5_niveaux.csv")

list_code_naf_niv5 = tab_naf_code.NIV5.str.replace(".","").tolist()
tab_naf_code["NIV5"] = list_code_naf_niv5
tab_naf_code["niveau5"] = list_code_naf_niv5
tab_naf_code["niveau4"] = tab_naf_code.niveau5.apply(lambda x : x[0:4])
tab_naf_code["niveau3"] = tab_naf_code.niveau5.apply(lambda x : x[0:3])
tab_naf_code["niveau2"] = tab_naf_code.niveau5.apply(lambda x : x[0:2])

tab_naf_code=[]

###################################"

#writer = pd.ExcelWriter( "list_mot_niv4.xlsx" , engine='xlsxwriter')

writer = pd.ExcelWriter( "list_mot_niv4.xlsx" , engine='xlsxwriter')

list_etudie = tab_naf_code["niveau4"].unique().tolist() 

for niv4 in list_etudie :
    DF_temp = pd.DataFrame(None,columns = ["word","occurence"])
    list_niveau5 = tab_naf_code[ tab_naf_code.niveau4 == niv4]["niveau5"].tolist()
    for niv5 in list_niveau5 :
        temp = DF.parse( niv5 )
        DF_temp = DF_temp.append(temp )
    DF_excel = DF_temp.groupby( [ "word"] ).sum()
    DF_excel.to_excel(writer, sheet_name= niv4 , index = True)
    for column in DF_excel:
                column_length = max(DF_excel[column].astype(str).map(len).max(), len(column)) #ajusté a la plus grande cellue
                col_idx = DF_excel.columns.get_loc(column)
                writer.sheets[ niv4 ].set_column(col_idx, col_idx, min(column_length,30) )
writer.save()
        
writer = pd.ExcelWriter( "list_mot_niv3.xlsx" , engine='xlsxwriter')

list_etudie = tab_naf_code["niveau3"].unique().tolist() 

for niv3 in list_etudie :
    DF_temp = pd.DataFrame(None,columns = ["word","occurence"])
    list_niveau5 = tab_naf_code[ tab_naf_code.niveau3 == niv3]["niveau5"].tolist()
    for niv5 in list_niveau5 :
        temp = DF.parse( niv5 )
        DF_temp = DF_temp.append(temp )
    DF_excel = DF_temp.groupby( [ "word"] ).sum()
    DF_excel.to_excel(writer, sheet_name= niv3 , index = True)
    for column in DF_excel:
                column_length = max(DF_excel[column].astype(str).map(len).max(), len(column)) #ajusté a la plus grande cellue
                col_idx = DF_excel.columns.get_loc(column)
                writer.sheets[ niv3 ].set_column(col_idx, col_idx, min(column_length,30) )
writer.save()
        
    
       



import pandas

df1 = pandas.DataFrame( { 
    "word" : ["Alice", "Bob", "Mallory", "Mallory", "Bob" , "Mallory"] , 
    "occurence" : [1,5, 2, 4, 2, 7] } )

g1 = df1.groupby( [ "word"] ).sum()








