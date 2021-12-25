import pandas as pd
import os
from neware import NDA_Processor as n
import new_nda, old_nda, nda_version_8_0
# change user filepath based on local computer name
# this is for WSL 
#user_filepath = r'C:\\Users\\xiang\\'
user_filepath = r"/mnt/c/Users/xiang/"
data_filepath = "Documents/libedp/test_data/E-chem data (auto backup) - Copy2"
#data_filepath = 'Box/PNE - R&D (Xiaofang)/R&D Administration/Projects/E-chem/E-chem data (auto backup) - Copy2'
# data_filepath = 'Box\PNE - R&D (Xiaofang)\R&D Adminstration\Projects\E-chem\E-chem data (auto backup)'
# C:\\Users\\xiang\\Box\PNE - R&D (Xiaofang)\R&D Adminstration\Projects\E-chem\E-chem data (auto backup)

backup_files_folder = user_filepath + data_filepath
outpath = backup_files_folder 

nda_files = os.listdir(outpath)
for item in nda_files:
    inpath = backup_files_folder + "/" + item
    n.nda_to_csv(inpath, inpath)