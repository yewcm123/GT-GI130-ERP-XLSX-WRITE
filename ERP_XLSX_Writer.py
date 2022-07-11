import xlsxwriter

class filetype:
    def __init__(self,filename,partnum,header):
        self._filename = filename
        self._partnum = partnum
        self._header = header
                
    def filename(self):
        return self._filename
    def partnum(self):
        return self._partnum
    def header(self):
        return self._header
    
    
#Header for each file type
PartMaster_header = ["Company","PartNum","Part Description","MfrPartNum_c","TypeCode",
          "UOMClassID","IUM","SalesUM","PUM","ProdCode","ClassID","PartBrand_c","Commodity Code",
          "NetWeight","NetWeightUOM","GrossWeight","GrossWeightUOM","CostMethod","NonStock"	,
          "QtyBearing","TrackSerialNum","RplPartNum_c","BuyToOrder"]

PartRev_header = ["Company	PartNum","RevisionNum","RevShortDesc","EffectiveDate",
                       "AltMethod","PartAudit#ChangeDescription","DrawNum"]

PartRevAttach_header = ["Company","PartNum","RevisionNum","FileName","DrawDesc"]

boo_header = ["Company","Plant","PartNum","RevisionNum","OprSeq","OpCode","ECOGroupID",
              "dFormat","ProdStandard","QtyPer"]

bom_header = ["Company","Plant","PartNum","RevisionNum","MtlSeq","MtlPartNum",
              "QtyPer","UOMCode","RelatedOperation","ECOGroupID","ViewAsAsm","PullAsAsm",
              "PlanAsAsm"]

def inquire_data():
    partnum = input("What is the Part Number?")
    return partnum

def open_file(filename):
    workbook = xlsxwriter.Workbook(filename)
    return workbook

def init_filetype():
    filetype("Part Master.xlsx","GI130-1047-G00000",PartMaster_header)
    filetype("Part Rev.xlsx","GI130-1047-G00000",PartMaster_header)
    filetype("Part Rev Attach.xlsx","GI130-1047-G00000",PartMaster_header)
    filetype("BOO.xlsx","GI130-1047-G00000",PartMaster_header)
    filetype("BOM.xlsx","GI130-1047-G00000",PartMaster_header)
        
def write_data(partnum, header_list):
    row = 0
    col = 0
    
    for x in header_list:
        worksheet.write(row,col,x)
        col+=1
    row += 1
    col = 0
    for y in header_list:
        worksheet.write(row, col, partnum)
        col+=1

partnum=inquire_data()

workbook = open_file("abc.xlsx")
worksheet = workbook.add_worksheet()
write_data(partnum, PartRevAttach_header)
workbook.close()
    
    


