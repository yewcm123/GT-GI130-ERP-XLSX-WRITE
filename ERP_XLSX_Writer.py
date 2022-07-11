"""
XLSX_ERP_GENERATOR

Description
Generates 5 ERP files according to input part number and updated with next day's date.
Created on Sat Jul  9 18:46:31 2022

@author: Yew Choon Min
"""

import xlsxwriter
import datetime

PartMaster_Header = "Company,PartNum,Part Description,MfrPartNum_c,TypeCode," \
                     "UOMClassID,IUM,SalesUM,PUM,ProdCode,ClassID,PartBrand_c,Commodity Code," \
                     "NetWeight,NetWeightUOM,GrossWeight,GrossWeightUOM,CostMethod,NonStock,"  \
                     "QtyBearing,TrackSerialNum,RplPartNum_c,BuyToOrder"

PartRev_Header = "Company,PartNum,RevisionNum,RevShortDesc,EffectiveDate," \
                  "AltMethod,PartAudit#ChangeDescription,DrawNum"

PartRevAttach_Header = "Company,PartNum,RevisionNum,FileName,DrawDesc"

BOO_Header = "Company,Plant,PartNum,RevisionNum,OprSeq,OpCode,ECOGroupID,"\
              "dFormat,ProdStandard,QtyPer"

BOM_Header = "Company,Plant,PartNum,RevisionNum,MtlSeq,MtlPartNum,"\
              "QtyPer,UOMCode,RelatedOperation,ECOGroupID,ViewAsAsm,PullAsAsm,"\
              "PlanAsAsm"
"""Header of each type of files"""

PartMaster_Data = "GTI001,GI130-1047-G00000,LABEL PRESENCE DETECTION - MARCHESINI PEN,, M,"\
                   "COUNT,SET,SET,SET,ASY,APC,GREATECH,,,,,,F,TRUE,TRUE,FALSE,,FALSE"

PartRev_Data = "GTI001,	GI130-1047-G00000,0,	Initial Design,07-Jul-22	,,Initial Design,"

PartRevAttach_Data = "GTI001,GI130-1047-G00000,0,\\gti-erp-apps\Drawing\BU Medical\PDF\GI130-1047-G00000-R0.PDF,GI130-1047-G00000-R0"

BOO_Data = "GTI001,MfgSys,GI130-1047-G00000,0,10,ASY,202200000877,PH	,0,1"

BOM_Data = "GTI001,MfgSys,GI130-1047-G00000,0,10,GI130-1047-E00000,1	,SET,10,202200000877	,1,0,0"
"""Data of each type of files"""



class Part_Data:
    """Store new parts datas"""
    def __init__(self,partnum,part_desc,ECO):
        self._partnum = partnum
        self._part_desc = part_desc
        self._ECO = ECO
    
    def partnum(self):
        return self._partnum
    
    def part_desc(self):
        return self._part_desc
    
    def ECO(self):
        return self._ECO
    
    PART_MAP = {        #Mapping each part type to an integer
    0 : "G",
    1 : "E",
    2 : "U",
    3 : "V",
    4 : "T",
    5 : "M"
    }

class File_Data:
    """Store each datas' variables and functions"""
    def __init__(self, file_type, file_name, header, data):
        self._file_type = file_type
        self._file_name = file_name
        self._header = header.split(",")
        self._data = data.split(",")
    
    def file_type(self):
        return self._file_type
    
    def file_name(self):
        return self._file_name
    
    def header(self):
        return self._header
    
    def data(self):
        return self._data
    
def User_Input():
    """Inquire user for input for part number, part description and ECO number
        
        Args:
            partnum: New part number
            part_desc: New part desciption
            ECO: New part ECO
    """
    partnum = input("Please Input part number(eg:GI130-1046): ")    #Prompt user for part number, part description and ECO number
    part_desc = input("Please input part description in CAPITAL LETTERS: ")
    ECO = input("Please input ECO Number: ")
    New_Part = Part_Data(partnum,part_desc,ECO) #Create new object based on user input
    return New_Part
    
def Print_xlsx(file):
    """Create a xlsx file and fill in required data into xlsx file and finally save the file.
    
        Args:
            file: file type (eg: Part Master file)
    """    
    workbook = xlsxwriter.Workbook(file.file_name())
    worksheet = workbook.add_worksheet()
    row = 0     #set header printing location to (A1)
    col = 0
    for header in file.header():    #print header according to file type
        worksheet.write(row,col,header)
        col+=1
    
    end_cond = 6    #Set ending row to print    
    row = 1         #Set data printing location to A2
    
    if file.file_type() == 4:   #SPECIAL CONDITION: Due to BOM file data referenced in parent and child format, this will not write
        end_cond-=1             #parent part type into part in child section
    while row <= end_cond:
        col = 0
        for data in file.data():
            worksheet.write(row,col,data)   #print data that does not required edit/change
            variable_change(worksheet,row,file)   #print data that required modification (eg:part type, date,MtlSeq)
            col+=1
        row += 1

    workbook.close()

def variable_change(worksheet,row,file):
    """Change variable to be written into file based on type of file
    
        Args:
            worksheet: current excel xlsx worksheet
            row: current row to be written on
            file: file data (eg: Part Master file data)             
    """     
    if file.file_type() == 0:       #Part Master file type
        worksheet.write(row,1,part_symbol(row-1))
        worksheet.write(row,2,New_Part.part_desc())
    elif file.file_type() == 1:     #Part Revision file type
        worksheet.write(row,1,part_symbol(row-1))
        worksheet.write(row,4,date_tomorrow.strftime("%d-%b-%y"))
    elif file.file_type() == 2:     #Part Revision with Attachment file type
        worksheet.write(row,1,part_symbol(row-1))
        worksheet.write(row,3,Attach_hyperlink(row-1))
        worksheet.write(row,4,part_symbol(row-1)+"-R0")
    elif file.file_type() == 3:     #BOO file type
        worksheet.write(row,2,part_symbol(row-1))
        worksheet.write(row,6,New_Part.ECO())
    elif file.file_type() == 4:     #BOM file type
        worksheet.write(row,2,part_symbol(0))
        worksheet.write(row,4,row*10)
        worksheet.write(row,5,part_symbol(row))
        worksheet.write(row,9,New_Part.ECO())
        
    else :
        raise ValueError

def part_symbol(types):
    """Write part file with its required part type
    
        Args:
            types: an integer to signify which part type to be written according to the part type mapping (PART_MAP)
    """
    return (New_Part.partnum()+"-"+New_Part.PART_MAP[types]+"00000")  

def Attach_hyperlink(types):
    """Write part file attachment document address hyperlink
    
        Args:
            types: an integer to signify which part type to be written according to the part type mapping (PART_MAP)
    """
    return ("\\\\gti-erp-apps\\Drawing\\BU Medical\\PDF\\"+part_symbol(types)+"-R0.PDF")

date_tomorrow = datetime.datetime.now() + datetime.timedelta(days=1)    #Set tomorrow's date as ERP update date

New_Part = User_Input()  #Ask user to enter new part datas

#Init each file data with respective name, file name, part name, header data and part data
PartMaster_File = File_Data(0,"1.Part Master.xlsx",PartMaster_Header,PartMaster_Data) 
PartRev_File = File_Data(1,"2.Part Revision.xlsx",PartRev_Header,PartRev_Data) 
PartRevAttach_File = File_Data(2,"3.Part Revision with attachment.xlsx",PartRevAttach_Header,PartRevAttach_Data) 
BOO_File = File_Data(3,"4.BOO.xlsx",BOO_Header,BOO_Data) 
BOM_File = File_Data(4,"5.BOM.xlsx",BOM_Header,BOM_Data) 

Print_xlsx(PartMaster_File)
Print_xlsx(PartRev_File)
Print_xlsx(PartRevAttach_File)
Print_xlsx(BOO_File)
Print_xlsx(BOM_File)


