import os, sys, time
import pandas as pd
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm, Inches, Mm, Emu
from docx import Document
import win32com.client
import olefile
import warnings

#to directly currend path
os.chdir(sys.path[0])




#to read the excel
with warnings.catch_warnings():
    warnings.filterwarnings("ignore", category=UserWarning)
    try:
        input_data = pd.read_excel('Input_Data.xlsx', 'MAIN').to_numpy()
    except Exception as message:
        print(message)

#For choose information device BGP peer as it area
def informationDeviceCN(area, subregion):
    area = str(area)
    if area == '22':
        return [['BDG-BBT2-CN1-C9910','114.1.186.1', '0.0.0.22','ASR9K'],['BDG-BBT2-CN2-C9910', '114.1.186.2', '0.0.0.22', 'ASR9K']]
    elif area == '24':
        return [
            ['BDG-BBT2-CN1-C9910','114.1.186.1','0.0.0.24',	'ASR9K'],
            ['BDG-BBT2-CN1-C9910','114.1.186.2','0.0.0.24',	'ASR9K'],
            ['SUK-SYK-CN1-C9910','114.1.191.1',	'0.0.0.24',	'ASR9K'],
            ['SUK-SYK-CN2-C9910','114.1.191.2',	'0.0.0.24',	'ASR9K']]
    elif area == '30':
        return [['BDG-BBT2-CN1-C9910','114.1.186.1', '0.0.0.22','ASR9K'],['BDG-BBT2-CN2-C9910', '114.1.186.2', '0.0.0.22', 'ASR9K']]
    elif area == '32':
        return [
            ['BDG-BBT2-CN1-C9910','114.1.186.1','0.0.0.32','ASR9K'],
            ['CRB-PLBN-CN1-C9910','114.1.200.1','0.0.0.32','ASR9K'],
            ['CRB-PLBN-CN2-C9910','114.1.200.2','0.0.0.32','ASR9K'],
            ['TGL-PSKN-CN1-C9910','114.1.214.1','0.0.0.32','ASR9K'],]
    elif area == '33':
        return [
            ['SKA-BHS-CN1-C9910','114.1.206.1','0.0.0.33','ASR9K'],
            ['SKA-BHS-CN2-C9910','114.1.206.2','0.0.0.33','ASR9K'],
        ]
    elif area == '34':
        return [
            ['TGL-PSKN-CN1-C9910','114.1.214.1','0.0.0.34','ASR9K'],
            ['SMG-GBL-CN1-C9922','114.1.220.1','0.0.0.34','ASR9922'],
            ['SMG-GBL-CN2-C9922','114.1.220.2','0.0.0.34','ASR9922'],
        ]
    elif area == '35':
        return [
            ['TGL-PSKN-CN1-C9910','114.1.214.1','0.0.0.35','ASR9K'],
            ['TGL-PSKN-CN2-C9910','114.1.214.2','0.0.0.35','ASR9K'],
            ['PWT-BTRD-CN1-C9910','114.1.203.1','0.0.0.35','ASR9K'],
            ['PWT-BTRD-CN2-C9910','114.1.203.2','0.0.0.35','ASR9K'],
        ]
    elif area == '37' or area == '137':
        if subregion == 'SMG':
            return [
                ['SMG-GBL-CN1-C9922','114.1.220.1','0.0.0.34','ASR9922'],
                ['SMG-GBL-CN2-C9922','114.1.220.2','0.0.0.34','ASR9922'],
            ]
        elif subregion == 'YOG':
            return [
                ['YOG-CDCT-CN1-C9910','114.1.210.1','0.0.0.37'],
                ['YOG-CDCT-CN2-C9910','114.1.210.2','0.0.0.37'],
            ]
        elif subregion == 'PWT':
            return [
                ['PWT-BTRD-CN1-C9910','114.1.203.1','0.0.0.35','ASR9K'],
                ['PWT-BTRD-CN2-C9910','114.1.203.2','0.0.0.35','ASR9K'],
            ]
    elif area == '50':
        return [
            ['TSM-SKNG-CN1-C9910','114.1.195.1','0.0.0.50','ASR9k'],
            ['TSM-SKNG-CN2-C9910','114.1.195.2','0.0.0.50','ASR9k'],
            ['PWT-BTRD-CN2-C9910','114.1.203.2','0.0.0.50','ASR9k'],
        ]


#for creating document or MoP
def createDocx(context, device, area):
    
    #For check device and choose template 
    if device == '7210 SASSX':
        template = DocxTemplate('template/Template_SASSX.docx')
    elif device == '7250 IXR-R6':
        template = DocxTemplate('template/Template_IXR-R6.docx')
    
    # print(context)
    # for data in context['info_device_cn']:
    #     print(data)
    # For create the document MoP
    try :
        template.render(context)
        template.save(f"result/NOKIA_R3_Method of Procedure Upgrade TiMOS {context['ring_id']} v1.docx")
        print(f"Document {context['ring_id']} has been created!!")
    except Exception as message :
        print(message)
        

#for defined and get each data from input excel
for data in input_data:
    data_ring_id = pd.read_excel('Input_Data.xlsx', data[0]).to_numpy()
    # print(data[0])
    
    #the data need to input at docx
    context = {
        'ring_id': data[0],
        'all_data_site' : data_ring_id,
        'date' : data[3],
        'info_device_cn': informationDeviceCN(data[2], data[4])
    }
    createDocx(context, data[1], data[2])
    
