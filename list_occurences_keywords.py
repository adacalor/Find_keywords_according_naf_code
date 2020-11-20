# -*- coding: utf-8 -*-
"""
Created on Fri Nov 20 11:27:43 2020

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

from sklearn.metrics.pairwise import cosine_similarity

from sklearn.neighbors import DistanceMetric

from sklearn.feature_extraction.text import CountVectorizer

from stop_words import get_stop_words

import warnings
warnings.simplefilter('ignore', UserWarning)

import coloredlogs, logging
coloredlogs.install()

import time


#####################"

from multiprocessing import Pool, current_process



########################################"

client = MongoClient( 'localhost', 27017)
db = client['enterprises']
collection = db['companies']


list_stop_words= get_stop_words("french") + get_stop_words("english") +["nan"]

####################################

codenaff = "4531Z"

def select_by_naf_code(codenaff) :
    search = { "organization.activity.ape_code" :  {"$regex" : str(codenaff)+".*"} }
    resp_form ={"id" : 1, 'web_infos.list_normalized_description' :1
        }
    cursor = collection.find( search, resp_form )
    
    list_id_text = [ {"id_item" : item["id"], "list_descr" : item["web_infos"]['list_normalized_description'] } for item in cursor ]                                      
    print( "taille pour le code naf {} ".format(list_id_text ))
    return( list_id_text )    


def list_occurences_from_list_strings ( text) :
    l_text = [ word for word in text if not( word in list_stop_words )]
    dict_occurences = dict((x,l_text.count(x)) for x in set(l_text))
    list_dict_occurences = [ { "word" : word, "occurence" : value  } for word,value in dict_occurences.items()  ]
    
    return( list_dict_occurences )
    

def occurrence_global(list_list_string,code_naf ) :
    DF = pd.DataFrame(None,columns = ["word","occurence"])
    if list_list_string == [] or list_list_string == None :
        pass
    else:
        big_list = sum(list_list_string, [] )
        DF = DF.append( list_occurences_from_list_strings ( big_list) )
        DF = DF.sort_values(by = ['occurence'], ascending=False )
    return(DF)


def global_mission( codenaff) :
    t1 = time.time()
    list_element = select_by_naf_code(codenaff)
    list_list_text = [ item["list_descr"] for item in  list_element ]
    DF = occurrence_global(list_list_text,codenaff )
    t2 = time.time()
    print(" finish " + codenaff+" en {} minutes".format( (t2-t1)/60 ))
    return(DF)

def list_occurence_by_code_naf( list_code_naf,addresse) :
    writer = pd.ExcelWriter( list_code_naf , engine='xlsxwriter')
    for code_naf in list_code_naf :
        DF = global_mission( code_naf)
        DF.to_excel(writer, sheet_name= code_naf , index = False)
        for column in DF:
                column_length = max(DF[column].astype(str).map(len).max(), len(column)) #ajusté a la plus grande cellue
                col_idx = DF.columns.get_loc(column)
                writer.sheets[ code_naf ].set_column(col_idx, col_idx, min(column_length,30) )
    writer.save()
                

list_occurence_by_code_naf( ["01","02","45", "62"],"01_02_45_62.xlsx")


list_occurence_by_code_naf( ["1051D","2013A","4110A", "5110Z"],"1051D_2013A_4110A_5110Z.xlsx")























