import pypyodbc

inAGS = True

pathForPre2017BMPpdfs = r'\\sparkssvm\DelDotPics\BMP_reports'
bmpInspOverall = ''
bmpInspActionScores = []

connection = pypyodbc.connect('Driver={SQL Server};' 'Server=vm-deldotsql\DELDOT;' 'Database=deldot_migration2;' 'uid=deldot;pwd=deldot#123')
connectionCL = pypyodbc.connect('Driver={SQL Server};' 'Server=vm-deldotsql\DELDOT;' 'Database=DelDOTbase;' 'uid=deldot;pwd=deldot#123')
connectionCloud = pypyodbc.connect('Driver={SQL Server};' 'Server=vm-deldotsql\DELDOT;' 'Database=devDeldot;' 'uid=deldot;pwd=deldot#123')
cloudTableName = "[devDeldot].[dbo].[File_Attachments_Cloud]"
workingPath = r"C:\Projects\DelDOT_NPDES_Reports"
imagePath = workingPath + r"\images"
mxdPath = workingPath + r"\mxds"
connFolder = workingPath + r"\connections"
GDBconn =connFolder + r"\gdb.sde"
GDBconn_CL = connFolder + r"\gdb.sde" 
flowLineConnStr =  ('Driver={SQL Server};' 'Server=vm-deldotsql\DELDOT;' 'Database=deldotBase;' 'uid=deldot;pwd=deldot#123') #new
#DBmapImageRoot = r"\\spsvm\P055\DelDotPics\BMP_DATABASE_MAPS"
DBmapImageRoot = r"\\sparkssvm\DelDotPics\BMP_DATABASE_MAPS"
DBmapImageBaseName = "database_map_image.jpg"
fileAttachRootURL = 'http://deldot103.kci.com/deldotwebviewer/api/Task/DownloadDoc?finename='
fileAttachRootURL = 'http://deldot103.kci.com/deldotwebviewer/api/Task/DownloadDoc?finename='
KCIprjs = 'KCI Projects 17151749C / 17161749N2'
dicAreas={'Laurel':'AREA 1 - LAUREL','Seaford':'AREA 2 - SEAFORD','Ellendale':'AREA 3 - ELLENDALE','Gravel Hill':'AREA 4 - GRAVEL HILL','Dagsboro':'AREA 5 - DAGSBORO','Harrington':'AREA 6 - HARRINGTON','Magnolia':'AREA 7 - MAGNOLIA','Cheswold':'AREA 8 - CHESWOLD','Middletown':'AREA 9 - MIDDLETOWN','Bear':'AREA 10 - BEAR','Kiamensi':'AREA 11 - KIAMENSI','Talley':'AREA 12 - TALLEY','Expressways':'AREA 14 - EXPRESSWAYS'}
#change this in the reportsharedfunctions module
connStr = "'Driver={SQL Server};' 'Server=vm-deldotsql\DELDOT;' 'Database=deldot_migration2;' 'uid=deldot;pwd=deldot#123'"


dicInspScoreVSactionName = { 'Access':'ACCESS_SCORE',
 'BMPVege':'BMP_VEG_SCORE',
 'ConveyCond':'CONV_COND_SCORE',
 'DoStrmEMCvr':'DOWNSTRM_COND_SCORE',
 'DStrmCond':'DOWNSTRM_EMB_CVR_SCORE',
 'EmgSpillway':'EMERG_SPILLWAY_SCORE',
 'Fence':'FENCE_SCORE',
 'InfCond':'INFLOW_COND_SCORE',
 'Invasive':'OV_INVASIVE_SCORE',
 'OSStab':'OV_SITE_STAB_SCORE',
 'PermanentTool':'PERM_POOL_SCORE',
 'Ponding':'PONDING_SCORE',
 'PreTreat':'PRE_TREATMNT_SCORE',
 'PrclSpill':'PRINCIPAL_SPILLWAY_SCORE',
 'PubHazards':'PUB_HAZARDS_SCORE',
 'RISAccumu':'RSR_ACCUM_SCORE',
 'RISLowflow':'RSR_LOWFLOW_SCORE',
 'RISOpening':'RSR_OPENING_SCORE',
 'RISStructure':'RSR_STRUCTURE_SCORE',
 'Seepage':'SEEPAGE_SCORE',
 'Spillway':'SPILLWAY_OUTFALL_SCORE',
 'UpStrmEMCvr':'UPSTRM_EMB_CVR_SCORE',
 'WaterQC':'WQ_CONTAM_SCORE',
 'WaterQT':'WQ_TREATMENT_SCORE'}