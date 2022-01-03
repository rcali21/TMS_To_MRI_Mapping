# Note that some of these packages might need to be installed, as they are not on the Martinos server natively...
import os
import nibabel as nb
from glob import glob
import numpy as np
from os import listdir
from os.path import isfile, join
import sys
import pandas as pd
import mne
from mne.transforms import apply_trans
import openpyxl 
import re
import xlsxwriter
import shutil
from pathlib import Path



#-----------MRI portion------------

# Specify input path to T1w image as such: /autofs/space/nicc_001/data/TMS-EEG/sourcedata/sub-ou1neuroc003/MRI/ou1neuroc003/MEMPRAGE_4e_p2_1mm_iso/007/MPRAGE.nii 

user_in = input('Specify path to MRI data (Must be isotropic):')
if os.path.exists(user_in):
    print('')
    print('File path found.')
    
else:
    print('ERROR: File path not found.')
    
    
nii = nb.load(user_in)

img = nii.get_fdata()

examine_header = nii.header


voxel_dims = (nii.header["pixdim"])[1:4]
print('')
print(voxel_dims)
isotropic_target = [1, 1, 1]

if (voxel_dims==isotropic_target).all():
    
    print('Isotropic volume found.')
    

else:
    
    print('')
    print('ERROR: This program expects an isotropic image (e.g., 1x1x1 dimensions).')
    sys.exit()

matrix_dims = (nii.header["dim"])[1:4]
MRI_x_dims = matrix_dims[0]
MRI_y_dims = matrix_dims[1]
MRI_z_dims = matrix_dims[2]

print("Image dimensions:",[MRI_x_dims, MRI_y_dims, MRI_z_dims])


#-----------TMS portion------------

# Specify input path to TMS-EEG sessions spreadsheet file as such: /autofs/space/nicc_001/data/TMS-EEG/sourcedata/sub-ou1neuroc003/ses-01/ou1neuroc003_ses01.xlsx 
session = input('Specify path to TMS session spreadsheet data:')
if os.path.exists(session):
    print('')
    print('File path found.')
    
else:
    print('ERROR: File path not found.')
    sys.exit()



print(session)




#-----------CONVERSION PORTION FOR TMS-TO-MRI COORDINATES------------
# NO DATA IS OVERWRITTEN. ALL SPREADSHEETS AND DATA REMAIN INTACT AND A NEW SPREADSHEET WILL BE GENERATED AFTER RUNNING. THIS WILL CONVERT EACH TMS-EEG SESSION TO MULTIPLE SHEETS
# IN ONE .xlsx file
 
pd.set_option('display.float_format', lambda x: '%.5f' % x)
x_col = [19]
y_col = [20]
z_col = [21]





target_cols = x_col + y_col + z_col

wb = openpyxl.load_workbook(session)

pd.set_option('display.max_rows', None, 'display.max_columns', None)




writer = pd.ExcelWriter('TMS_MRI_Coordinates.xlsx', engine='xlsxwriter')


for sheet in wb.sheetnames:
    if re.findall('Session.+', sheet):
        df = pd.read_excel(session, sheet, header=None)
        
        df1 = df[df.columns[target_cols]]
        x_targ = df1[df.columns[19]]
        y_targ = df1[df.columns[20]]
        z_targ = df1[df.columns[21]]
        
        
        df1 = pd.to_numeric(x_targ,
                                     errors = 'coerce')
        df2 = pd.to_numeric(y_targ,
                                     errors = 'coerce')
        df3 = pd.to_numeric(z_targ,
                                     errors = 'coerce')
        
        df1 = MRI_x_dims - df1 / 1
        df2 = MRI_z_dims - df2 / 1
        df3 = MRI_y_dims - df3 / 1
        
        
        
        
        df['x_val'] = df1
        df['y_val'] = df2
        df['z_val'] = df3
        
        
        
        MRI_coords = [24, 25, 26]
        
        df_new = df[df.columns[MRI_coords]]
        
        x = apply_trans(nii.affine, df_new)
        
        
        a = pd.DataFrame(x)
        
        r = a.round(3)
        
        r.to_excel(writer, sheet_name=sheet, header=['RAS_x', 'RAS_y', 'RAS_z'])  
      
    
writer.save() # Add something before this line if you wish to have the writer write to a specific directory, otherwise a new .xlsx spreadsheet will be written to your current wd
# where you are running the ipynb. 





b = os.getcwd()
a = os.path.dirname(session)


for src_file in Path(b).glob('*TMS_MRI_Coordinates*'):
    shutil.copy(src_file, a)
    

final_dir = os.path.join(a, src_file)    
print('Success!')
print('')
print('MRI transforms written to:', final_dir)
        


