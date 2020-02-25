#copy "C:\Python27\ArcGIS10.3\Lib\site-packages\Desktop10.3.pth" to C:\ProgramData\Miniconda2\Lib\site-packages
import arcpy
import os, sys, traceback ,  shutil,uuid
from reportlab.lib.pagesizes import letter, inch
from reportlab.platypus import *
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.rl_config import defaultPageSize
from reportlab.graphics import shapes
from reportlab.lib import colors
from reportlab.graphics.charts.textlabels import Label
import datetime, time
import ReportSharedFunctionsES2M ,appConfigES2M ,IRF

workingPath = appConfigES2M.workingPath

PAGE_HEIGHT=defaultPageSize[1]; PAGE_WIDTH=defaultPageSize[0]

styles = getSampleStyleSheet()
HeaderStyle = styles["Heading2"]
ParaStyle = styles["Normal"]
PreStyle = styles["Code"]
styleN = styles['Normal']

pageinfo = datetime.datetime.now().strftime("%B %d, %Y  %I:%M %p")

headFootTop=0.65
headFootInfo = ''

def myFirstPage(canvas, doc):
    canvas.saveState()   
    canvas.setFont('Helvetica',9)
    canvas.drawString(inch, 0.75 * inch, pageinfo)
    canvas.restoreState()

def myLaterPages(canvas, doc):
    canvas.saveState()

    canvas.setFont("Helvetica", 8)
    canvas.drawString(inch, headFootTop*inch, "Page %d -- %s" % (doc.page , pageinfo))
    canvas.restoreState()


#######################################################################################



#######################################################################################
try:    
    appConfigES2M.inAGS =  arcpy.env.scratchWorkspace != None

    if appConfigES2M.inAGS:
        reportType = arcpy.GetParameterAsText(0)
        form_ID = arcpy.GetParameterAsText(1)
        userName = arcpy.GetParameterAsText(2)
        workingFolder = arcpy.env.scratchWorkspace
    else:
        reportType = '1'
        form_ID = '13'
        userName = 'gerry.kelly'
        workingFolder = appConfigES2M.workingPath

    if reportType == '1':
        elements = IRF.returnIRF(form_ID)
        RptName = "\\DelDOT_ES2M_Inspection_Ratings_Form_"    +  form_ID + "_report.pdf"
    elif reportType == '2':
        pass
    elif reportType == '3':
        pass


    
    DocPath = workingFolder + RptName 

    doc = SimpleDocTemplate(DocPath, pagesize=(8.5*inch, 11*inch), allowSplitting=1)  
    doc.rightMargin = .5*inch
    doc.leftMargin =  .5*inch
    doc.topMargin = 30
    doc.bottom = 20
    doc.allowSplitting = 1

    doc.build(elements, onFirstPage=myFirstPage, onLaterPages=myLaterPages)

    arcpy.SetParameterAsText(2,DocPath)

except:
    tb = sys.exc_info()[2]
    tbinfo = traceback.format_tb(tb)[0]
    pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
            str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
    print pymsg; arcpy.AddError(pymsg)
    try:
        appConfigES2M.connectionES2M.close()
    except:
        pass
    pass