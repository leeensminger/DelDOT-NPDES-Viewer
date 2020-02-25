from reportlab.lib.pagesizes import letter, inch
from reportlab.platypus import *
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.rl_config import defaultPageSize
from reportlab.graphics import shapes
from reportlab.lib import colors
from cgi import escape
	
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
 
 
import  os, sys, traceback ,io
import datetime, time
import appConfig, ReportSharedFunctions
import pypyodbc,PIL#
import arcpy 
import urllib

pdfmetrics.registerFont(TTFont('Tahoma', 'Tahoma.ttf'))

DocPath = '' 
pageinfo = "Date Generated: " + datetime.datetime.now().strftime("%#m/%d/%Y  %#I:%M %p")
isMemo = False
memoHead = ''
pdfs = []


def myFirstPage(canvas, doc):
    canvas.saveState()   
    canvas.setFont('Helvetica',9)
    canvas.drawString(inch, 0.75 * inch, pageinfo)
    canvas.restoreState()

def myLaterPages(canvas, doc):
    canvas.saveState()

    canvas.setFont("Helvetica", 8)
    if isMemo:
        canvas.drawString(inch, 0.75 * inch, pageinfo)
    else:
        canvas.drawString(inch, 0.75 * inch, pageinfo) #canvas.drawString(inch, 0.65 * inch, "Page %d -- %s" % (doc.page , pageinfo))
    canvas.restoreState()

styles = getSampleStyleSheet()
ParaStyle = styles["Normal"]


box = Paragraph('<para alignment= "LEFT"><font name=Tahoma size=10>&#9677;</font>', styles["Normal"]) 
boxFilled = Paragraph('<para alignment= "CENTER"><font name=Tahoma size=12>&#x25A0</font>', styles["Normal"])

def coverPDF(val, num, formNo, folder):

    attTy = ReportSharedFunctions.valFromDomCode("D_AttachmentType",val)
    if num !=0:
        if num == 1:
            p="PDF"
        else:
            p="PDFs"
        
        numbers = ["", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten"]
        if num < 11:
            num = numbers[num].title()

        t = str(num) + " " + p + " related to this investigation\nwhere attachment type = " + attTy + " to follow."
    else:
        t = "There are no PDFs related to this investigation where\nattachment type = " + attTy

    elements = []
    content =[['Form '+ str(formNo) ]]
    content.append([''])
    content.append([t])

    tbl = Table(content, [7*inch])
    tbl.setStyle(TableStyle([ ('SIZE', (0,0), (0, 0), 16),('SIZE', (0,2), (0, 2), 12),  ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'),
                                    ('ALIGN', (0,0),  (-1,-1), 'CENTER')  
                                ]))
    
    if not t.startswith('There'):
        elements.append(tbl)                  
       
    DocPath = folder + "\\covPDFs" + str(formNo) + ".pdf"
    pdfs.append(DocPath)
    docWQ = SimpleDocTemplate(DocPath, pagesize=(8.5*inch, 11*inch), allowSplitting=1)
    docWQ.rightMargin = .75*inch
    docWQ.leftMargin =  .75*inch
    docWQ.topMargin = 144
    docWQ.bottom = 20
    docWQ.allowSplitting = 1         
    docWQ.build(elements, onFirstPage=myFirstPage, onLaterPages=myLaterPages) 
     

def downnLoadPDFs(glob, aType, pref, folder,formNo):
    sql = ("SELECT [cloud_file_name] FROM " + appConfig.cloudTableName + " where [FEATURE_ID] = '{0}' and [ATTACHMENT_TYPE] = '{1}'").format(glob, aType)
    docs = ReportSharedFunctions.returnDataODBCcloud(sql)
    cntr = 1
    
    #coverPDF(aType, len(docs),formNo,folder)

    for doc in docs:
        cntr += 1
        filename = doc[0]   
        url = appConfig.fileAttachRootURL + filename
        try:
            path = folder + "\\" + pref + "_" + str(cntr) + ".pdf"
            DLfile = urllib.URLopener()
            DLfile.retrieve(url, path)    
            pdfs.append(path)
        except:
            tb = sys.exc_info()[2]
            tbinfo = traceback.format_tb(tb)[0]
            pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                    str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
            if "406" in pymsg:
                arcpy.AddWarning("http error 406 for download of file " + filename )
            else:
                arcpy.AddWarning("error for download of file " + filename + ".  " +  pymsg)
            print pymsg; arcpy.AddError(pymsg)
            #raise
            pass

    pass

def unitNoBold(attriCommaUnit):
    au = attriCommaUnit.split(',')
    attri=au[0]
    unit=au[1]
    return Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=10>'+attri+'&nbsp </font><font name=Helvetica size=10>'+unit+'</font>', styles["Normal"]) 


def frm8Make(globalid,INCIDENT_ID,ty,FEATURE_ID,ftrNum, OBJECTID): 
    try:
        WQfields_NewforSep2019 =  ReportSharedFunctions.returnDataODBC("Select [SETTING] ,[ODOR] ,[NEAREST_STREAM]  ,[SETTING_OTHER_DESC] ,[ODOR_OTHER_DESC] FROM [deldot_migration2].[dbo].[WQ_INVESTIGATIONS_evw] where [globalid] = '" + globalid + "'")[0]
        WQfields_NewforSep2019 = ['' if x == None else x for x in WQfields_NewforSep2019] 
        SETTING_NewforSep2019 = WQfields_NewforSep2019[0]
        ODOR_NewforSep2019 = WQfields_NewforSep2019[1]
        NEAREST_STREAM_NewforSep2019 = WQfields_NewforSep2019[2]
        SETTING_OTHER_DESC_NewforSep2019 = WQfields_NewforSep2019[3]
        ODOR_OTHER_DESC_NewforSep2019 = WQfields_NewforSep2019[4]

        tblStyle = [('SIZE', (0,0), (-1, -1), 10),  ('TOPPADDING', (0,0), (-1, -1), 1),  ('BOTTOMPADDING', (0,0), (-1, -1), 1), ('VALIGN', (0,0), (-1, -1),'MIDDLE') , ('ALIGN', (2,0), (2, -1),'CENTER'),   
                                ('BOX', (0,0), (-1,-1),0.5, colors.darkgray) ,('INNERGRID', (0,0), (-1,-1), 0.5, colors.darkgray), ('FONTNAME', (0,0), (1,-1), 'Helvetica-Bold')  
                                ,('SPAN', (0,19), (0, 21)) ,('SPAN', (0,22), (0, 24)) ,('SPAN', (0,25), (0, 26)),('SPAN', (0,0), (1, 0)),('SPAN', (0,1), (1, 1)),('SPAN', (0,2), (1, 2)),  
                                ('SPAN', (0,3), (1, 3)),('SPAN', (0,4), (1, 4)),('SPAN', (0,5), (1, 5)),('SPAN', (0,6), (1, 6)),('SPAN', (0,7), (1, 7)),('SPAN', (0,8), (1, 8)),  
                                ('SPAN', (0,9), (1, 9)),('SPAN', (0,10), (1, 10)),('SPAN', (0,11), (1, 11)),('SPAN', (0,12), (1, 12)),('SPAN', (0,13), (1, 13)),('SPAN', (0,14), (1, 14)),  
                                ('SPAN', (0,15), (1, 15)),('SPAN', (0,16), (1, 16)),('SPAN', (0,17), (1, 17)),('SPAN', (0,18), (1, 18)) ]   
        elements = []

 
        if ty == "STRUCTURE":
            table = "STRUCTURES_evw"
        elif ty == "CONVEYANCE":
            table = "CONVEYANCES_evw"
        elif ty == "PID_POINT":
            table = "pid_points_evw"

        featureFields = ReportSharedFunctions.returnDataODBC("SELECT [COUNTY], [ADDRESS_LINE1] ,[CITY] ,COALESCE(SUBDIVISION_NAME_FIELD, SUBDIVISION_NAME_GIS, 'N/A') as Subdivision, [WATERSHED] , 'pending' as Stream ,[ZIP]  FROM [deldot_migration2].[dbo]." + table + " where globalid = '" + FEATURE_ID + "'" )[0]
        featureFields = ['' if x == None else x for x in featureFields] 
        COUNTY = ReportSharedFunctions.valFromDomCode("D_County",featureFields[0])
        ADDRESS_LINE1 = featureFields[1].title()
        Subdivision = featureFields[3].title()
        incNo = INCIDENT_ID if INCIDENT_ID != None else "N/A"

        content = [['ILLICIT DISCHARGE DETECTION & ELIMINATION\nFIELD SHEET']]
        content.append([ty.title() + ' Number:',ftrNum])
        content.append(['Incident ID #:',incNo])
        content.append(['County:',COUNTY])
        content.append(['Subdivision:',Subdivision])
        content.append(['Address/Location',ADDRESS_LINE1])
        content.append(['',''])

        reportGrid = Table(content, [3.5*inch,3.5*inch])
        reportGrid.setStyle(TableStyle([                                      
                                        ('SIZE', (0,0), (1, 0), 12),('SIZE', (0,1), (-1, -1), 11),  ('FONTNAME', (0,0), (1,0), 'Helvetica-Bold'),
                                        ('FONTNAME', (1,0), (1,-1), 'Helvetica-Bold') , ('ALIGN', (0,0), (0,0), 'CENTER') , ('ALIGN', (0,1), (0,-1), 'RIGHT'),
                                        ('SPAN', (0,0), (1, 0)) , ('TOPPADDING', (0,0), (-1, -1), 1),  ('BOTTOMPADDING', (0,0), (-1, -1), 2) 
                                    ]))

        elements.append(reportGrid)

        wqFields = ReportSharedFunctions.returnDataODBC("SELECT q.[PERSONNEL], format(q.[CREATE_DATE],'M/d/yyyy') as d, format(q.[CREATE_DATE],'h:mm tt') as t, DETERMINATION,TYPE_DISCHARGE FROM [deldot_migration2].[dbo].[WQ_INVESTIGATIONS_evw] q where q.objectid = " + OBJECTID )[0]
        wqFields = ['' if x == None else x for x in wqFields]
         
         #GK 9/19 WQI_FIELD_SHEET was (case ODOR when 'Other' then 'O - ' + isnull(ODOR_Other_DESC,'') else ODOR end)  > WQIsheet[10]
        WQIsheet = ReportSharedFunctions.returnDataODBC("SELECT  format([AIR_TEMP],'###') ,(case when LAST_RAIN_DATE is null then '' else FORMAT([LAST_RAIN_DATE], 'M/d/yyyy') end ), " +
                          " (case reason when 'Sample Collected' then 'Yes' when 'Not Collected - Stream' then 'Yes-Stream' when 'Not Collected - Cannot Access/Cannot Locate' then 'CA/CL' when 'Not Collected - No Flow' then 'No' else ' ' end),  " + 
                          "(case LAND_USE when 'Other' then 'O - ' + isnull(LAND_USE_Other_DESC,'') else LAND_USE end), " +
                          " (case STRUCTURAL_COND when 'Other' then 'O - ' + STRUCTURAL_COND_Other_DESC else STRUCTURAL_COND end), erosion,ALGAE_GROWTH, " +
                          " (case VEG_COND when 'Other' then 'O - ' + isnull(VEG_COND_Other_DESC,'') else VEG_COND end),globalid,STAINING_PRESENT,'x' dummyOdor_OderDesc  " +
                          ", (case DEPOSITS when 'Other' then 'O - ' + isnull(DEPOSITS_Other_DESC,'') else DEPOSITS end)   FROM [deldot_migration2].[dbo].[WQI_FIELD_SHEET_evw] where [WQ_INVESTIGATION_ID] = '" + globalid + "'")[0]
        WQIsheet = ['' if x == None else x for x in WQIsheet]
        WQsheetGlobalid = WQIsheet[8]

        Persnl = wqFields[0]; 	DateOf = wqFields[1]; 	TimeOf = wqFields[2]; 	AirTemp = WQIsheet[0]; 	
        Determ =  wqFields[3]  
        if Determ != "Evidence of Illicit Discharge":
            Determ = ReportSharedFunctions.valFromDomCode("D_Determination",Determ)
        else:
            Determ = ReportSharedFunctions.valFromDomCode("D_TypeofIllicitDischarge",wqFields[4])
         
        PhotoYN = "No" if ReportSharedFunctions.returnDataODBC("SELECT count(*) from [deldot_migration2].[dbo].[WQI_INVESTIGATION_PHOTOS_evw] where WQ_INVESTIGATION_ID = '" + globalid + "'")[0][0] == 0  else "Yes"	
        DatLastRain = WQIsheet[1];  
       
        if ty == "STRUCTURE":
            strucTyGID = ReportSharedFunctions.returnDataODBC("SELECT  STRUCTURE_TYPE, globalid FROM [deldot_migration2].[dbo]." + table + " where globalid = '" + FEATURE_ID + "'" )[0] 
            strucTy = strucTyGID[0]
            strucGID = strucTyGID[1]
            if strucTy == 'Pipe End':
                sql = "SELECT p.PIPE_SHAPE, case when p.WIDTH  is  not null then format(round(p.WIDTH,0),'###') else '' end,  case when p.HEIGHT is not  null then format(round(p.HEIGHT,0),'###') else '' end,p.MATERIAL  FROM [deldot_migration2].[dbo].[STRUCTURES_evw] s join [deldot_migration2].[dbo].[CONVEYANCES_evw] c on c.DOWNSTREAM_STRUCTURE_ID = s.globalid join [deldot_migration2].[dbo].[PIPE_SEGMENTS_evw] p on p.conveyance_id = c.globalid where s.globalid = '" + FEATURE_ID + "'"
                pipeend = ReportSharedFunctions.returnDataODBC(sql)[0] 
                pipeend = ['' if x == None else x for x in pipeend]
                OutfallShape = pipeend[0]
                if OutfallShape == "Round":
                    OutfallDim = pipeend[1]
                else:
                    OutfallDim  =pipeend[1] + " X " + pipeend[2]
                OutfallShape = ReportSharedFunctions.valFromDomCode("D_PipeShape",OutfallShape) 
                
                OutfallTy = ReportSharedFunctions.valFromDomCode("D_PipeMaterial",pipeend[3])                 
            else:
                strucTy = ReportSharedFunctions.valFromDomCode("D_StructureTypes_1",strucTy) 
                OutfallDim = "N/A - " + strucTy
                OutfallShape = 'N/A'
                OutfallTy = 'N/A'

        elif ty == "CONVEYANCE":
            OutfallDim = "N/A - Conveyance"
            OutfallShape = "N/A - Conveyance"
            OutfallTy = "N/A - Conveyance"
        elif ty == "PID_POINT":
            OutfallDim = "N/A - PID POINT"
            OutfallShape = "N/A - PID POINT"
            OutfallTy = "N/A - PID POINT"

        FlowObs = WQIsheet[2]; 	LU = ReportSharedFunctions.valFromDomCode("D_LandUse",WQIsheet[3]); 	StrucCond =  ReportSharedFunctions.valFromDomCode("D_StructureCondition",WQIsheet[4])
        ErosArea = ReportSharedFunctions.valFromDomCode("D_Erosion",WQIsheet[5]) ; 	AlJgrowth = ReportSharedFunctions.valFromDomCode("D_Boolean",WQIsheet[6])
        VegCond = ReportSharedFunctions.valFromDomCode("D_Vegetation",WQIsheet[7]) ; 	
        if ODOR_NewforSep2019 == 'Other':
            Odor = 'O - ' + ODOR_OTHER_DESC_NewforSep2019
        else:
            Odor = ODOR_NewforSep2019

        Odor =  ReportSharedFunctions.valFromDomCode("D_Odor",Odor); 	#was WQIsheet[10] 9/19
        Deposits = ReportSharedFunctions.valFromDomCode("D_Deposits",WQIsheet[11]); 	
        StainPres=ReportSharedFunctions.valFromDomCode("D_Boolean",WQIsheet[9])

        SurfLab = "N/A"; 	SurfFollow = "N/A"; 		AmmonLab = "N/A"; 	 AmmonFollow = "N/A"; 	PotasLab = "N/A"; 	PotasFollow = "N/A";        
             
        FloRate = ""; WatTemp = ""; 	pH = ""; 	Turbidity = ""; 	SurfField = ""; AmmonField = ""; Color = ""; 	Float = "";
        hasFlow = False
        hasLab = False
        sql = ("SELECT  (case FLOW_RATE when 0 then 'Unable to measure' else format(round(FLOW_RATE,5),'##0.#####') end) ,format([WATER_TEMP],'##0'),format([PH],'###.00'),format([TURBIDITY_ID],'##0.00'), " +
        " format([SURFACTANTS],'##0.00##'), format([ammonia],'##0.00##'), (case COLOR when 'Other' then 'O - ' +  isnull(COLOR_OTHER_DESC,'')  else COLOR end), " + 
        " (case FLOATABLES when 'Other' then 'O - ' +  isnull(FLOATABLES_OTHER_DESC,'')  else FLOATABLES end),globalid FROM [deldot_migration2].[dbo].[FLOW_CHAR_evw] where flow_id = '"  + WQsheetGlobalid + "'" )
        FlowChar = ReportSharedFunctions.returnDataODBC(sql)
        if len(FlowChar) > 0:
            hasFlow = True
            FlowChar=FlowChar[0]    
            FlowChar = ['' if x == None else x for x in FlowChar]         
        
        if hasFlow:
            FloRate= FlowChar[0]
            WatTemp = FlowChar[1]
            pH = FlowChar[2]
            Turbidity = FlowChar[3]
            SurfField = FlowChar[4]
            AmmonField = FlowChar[5]
            Color = ReportSharedFunctions.valFromDomCode("D_Color",FlowChar[6])   
            Float = ReportSharedFunctions.valFromDomCode("D_Floatables",FlowChar[7])  
            FloID = FlowChar[8]


 
            # now see if lab
        
            cnt =  ReportSharedFunctions.returnDataODBC("select count(*)  FROM [deldot_migration2].[dbo].[LAB_RESULTS_evw] where [FLOW_CHAR_ID] = '" + FloID + "'")[0][0]
            if cnt > 0:
                hasLab = True
                labGlobIDs = ReportSharedFunctions.returnDataODBC("select globalid  FROM [deldot_migration2].[dbo].[LAB_RESULTS_evw] where [FLOW_CHAR_ID] = '" + FloID + "' order by [SENT_TO_LAB_DATE] desc")
                labGlobIDlist = []
                for l in labGlobIDs:
                    labGlobIDlist.append(l[0])
                
                labRes = ReportSharedFunctions.returnDataODBC("select SURFACTANTS FROM [deldot_migration2].[dbo].[LAB_RESULTS_evw] where [FLOW_CHAR_ID] = '" + FloID + "' and SURFACTANTS is not null order by [SENT_TO_LAB_DATE]")
                if len(labRes)>0:
                    SurfLab = labRes[0][0]
                if len(labRes)>1:
                    SurfFollow =labRes[len(labRes) - 1][0]
 
                abRes = ReportSharedFunctions.returnDataODBC("select AMMONIA FROM [deldot_migration2].[dbo].[LAB_RESULTS_evw] where [FLOW_CHAR_ID] = '" + FloID + "' and AMMONIA is not null order by [SENT_TO_LAB_DATE]")
                if len(labRes)>0:
                    AmmonLab = labRes[0][0]
                if len(labRes)>1:
                    AmmonFollow =labRes[len(labRes) - 1][0]

                labRes = ReportSharedFunctions.returnDataODBC("select POTASSIUM FROM [deldot_migration2].[dbo].[LAB_RESULTS_evw] where [FLOW_CHAR_ID] = '" + FloID + "' and POTASSIUM is not null order by [SENT_TO_LAB_DATE]")
                if len(labRes)>0:
                    PotasLab = labRes[0][0]
                if len(labRes)>1:
                    PotasFollow =labRes[len(labRes) - 1][0]
#10/2018
        FloRate = 'N/A' if FloRate == '' else FloRate
        WatTemp = 'N/A' if WatTemp == ''  else WatTemp
        pH = 'N/A' if pH== '' else pH
        Turbidity = 'N/A' if Turbidity== '' else Turbidity
        SurfField = 'N/A' if SurfField== '' else SurfField
        AmmonField = 'N/A' if AmmonField== '' else AmmonField
        Color = 'N/A' if Color== '' else Color
        Float = 'N/A' if Float== '' else Float

        content  = [['Personnel',"",Persnl]]
        content.append(['Date',"",DateOf])
        content.append(['Time',"",TimeOf])
        content.append([unitNoBold("Air Temperature,(F)"),"",AirTemp])
        content.append(['Photograph',"",PhotoYN])
        content.append(['Date Last Rain',"",DatLastRain])
        content.append([unitNoBold("Outfall Dimensions,(inches)"),"",OutfallDim])
        content.append(['Outfall Shape',"",OutfallShape])
        content.append(['Outfall Type',"",OutfallTy])
        content.append(['Flow Observed',"",FlowObs])
        content.append(['Land Use',"",LU])
        content.append(['Structural Condition',"",StrucCond])
        content.append([unitNoBold("Erosion,(Outfall Area)"),"",ErosArea])
        content.append(['Algae Growth',"",AlJgrowth])
        content.append([unitNoBold("Vegetative Condition,(Outfall Area)"),"",VegCond])
        content.append([unitNoBold("Flow Rate,(cfs)"),"",FloRate])
        content.append([unitNoBold("Water Temperature,(F)"),"",WatTemp])
        content.append([unitNoBold("pH,(units)"),"",pH])
        content.append([unitNoBold("Turbidity,(ntu)"),"",Turbidity])
        content.append([unitNoBold("Surfactants,(mg/L)"),"Field Tested:",SurfField])
        content.append(['',"Lab Tested:",SurfLab])
        content.append(['',"Follow Up Lab Tested: ",SurfFollow])
        content.append([unitNoBold("Ammonia,(mg/L)"),"Field Tested:",AmmonField])
        content.append(['',"Lab Tested:",AmmonLab])
        content.append(['',"Follow Up Lab Tested:",AmmonFollow])
        content.append([unitNoBold("Potassium,(mg/L)"),"Lab Tested:",PotasLab])
        content.append(['',"Follow Up Lab Tested:",PotasFollow])
                                                       
        ############################################
        #now the conditional adding of line and spanning
        lineLogicIndex = len(content) -1

        DieselLab = ""; 	GasLab = ""; OilGreaseLab = "";	  	
        DieselFollow = ""; 	GasFollow = "";  OilGreaseFollow = "";	 
        if hasLab:
            labRes = ReportSharedFunctions.returnDataODBC("select DIESEL_ORGANICS FROM [deldot_migration2].[dbo].[LAB_RESULTS_evw] where [FLOW_CHAR_ID] = '" + FloID + "' and DIESEL_ORGANICS is not null order by [SENT_TO_LAB_DATE]")
            if len(labRes)>0:
                DieselLab = labRes[0][0]
                content.append([unitNoBold("Diesel Range Organics,(mg/L)"),"Lab Tested:",DieselLab])
                lineLogicIndex +=1
            if len(labRes)>1:
                DieselFollow =labRes[len(labRes) - 1][0]
                content.append(['',"Follow Up Lab Tested:",DieselFollow])
                tblStyle.append(('SPAN', (0,lineLogicIndex), (0, lineLogicIndex+1)))
                lineLogicIndex +=1

            labRes = ReportSharedFunctions.returnDataODBC("select [GAS_ORGANICS] FROM [deldot_migration2].[dbo].[LAB_RESULTS_evw] where [FLOW_CHAR_ID] = '" + FloID + "' and GAS_ORGANICS is not null order by [SENT_TO_LAB_DATE]")
            if len(labRes)>0:
                GasLab = labRes[0][0]
                content.append([unitNoBold("Gasoline Range Organics,(mg/L)"),"Lab Tested:",GasLab])
                lineLogicIndex +=1
            if len(labRes)>1:
                GasFollow =labRes[len(labRes) - 1][0]
                content.append([unitNoBold("Gasoline Range Organics,(mg/L)"),"Follow Up Lab Tested:",GasFollow])
                tblStyle.append(('SPAN', (0,lineLogicIndex), (0, lineLogicIndex+1)))           
            

            labRes = ReportSharedFunctions.returnDataODBC("select [OIL_GREASE] FROM [deldot_migration2].[dbo].[LAB_RESULTS_evw] where [FLOW_CHAR_ID] = '" + FloID + "' and [OIL_GREASE] is not null order by [SENT_TO_LAB_DATE]")
            if len(labRes)>0:
                OilGreaseLab = labRes[0][0]
                content.append([unitNoBold("Oil & Grease,(mg/L)"),"Lab Tested:",OilGreaseLab])
                lineLogicIndex +=1
            if len(labRes)>1:
                OilGreaseFollow =labRes[len(labRes) - 1][0]
                content.append([unitNoBold("Gasoline Range Organics,(mg/L)"),"Follow Up Lab Tested:",OilGreaseFollow])
                tblStyle.append(('SPAN', (0,lineLogicIndex), (0, lineLogicIndex+1)))     


            # now if lab-other
            lstTests = []
            for lab in labGlobIDlist:  #span the line                
                labReses = ReportSharedFunctions.returnDataODBC(" SELECT [TEST_DESCRIPTION] ,[TEST_VALUE]  FROM [deldot_migration2].[dbo].[LAB_RESULTS_OTHERS_evw] where [LAB_RESULT_ID] = '" + lab + "'")
                for labRes in labReses:
                    labRes = ['' if x == None else x for x in labRes]
                    if not labRes[0] in lstTests:
                        lstTests.append(labRes[0])
                        content.append([labRes[0],"Lab Tested:",labRes[1]])         

        content.append(['Odor',"",Odor])
        content.append(['Deposits/Stains',"",Deposits])
        content.append(['Color',"",Color])
        content.append(['Floatables',"",Float])
        content.append(['Determination (From IDDE Flowchart)',"",Determ])

        ln = len(content)
        for x in range(1,6):
            tblStyle.append(('SPAN', (0,ln-x), (1, ln-x)))

        reportGrid = Table(content, [2.5*inch,2.5*inch,3*inch])
        reportGrid.setStyle(TableStyle(tblStyle ))

        elements.append(reportGrid)
        return elements;

    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass


def frm3Make_revised1017(globalid,INCIDENT_ID,ty,FEATURE_ID,ftrNum, OBJECTID,DocPath,workingFolder):
    try:
        elements = []

        if ty == "STRUCTURE":
            table = "STRUCTURES_evw"
        elif ty == "CONVEYANCE":
            table = "CONVEYANCES_evw"
        elif ty == "PID_POINT":
            table = "pid_points_evw"

        memo1 = ReportSharedFunctions.returnDataODBC("SELECT TOP 1 MEMO_TO, MEMO_from, [MEMO_COMMENT],objectid FROM [deldot_migration2].[dbo].[WQI_MEMOS_evw] where [WQ_INVESTIGATION_ID] = '" + globalid + "' order by objectid desc;")[0]
        memo1 = ['' if x == None else x for x in memo1]
        
       
        featureFields = ReportSharedFunctions.returnDataODBC("SELECT [COUNTY], [ADDRESS_LINE1] ,[CITY] ,COALESCE(SUBDIVISION_NAME_FIELD, SUBDIVISION_NAME_GIS, 'N/A') as Subdivision, [WATERSHED] , 'pending' as Stream ,[ZIP]  FROM [deldot_migration2].[dbo]." + table + " where globalid = '" + FEATURE_ID + "'" )[0]
        featureFields = ['' if x == None else x for x in featureFields] 
        COUNTY = ReportSharedFunctions.valFromDomCode("D_County",featureFields[0])
        ADDRESS_LINE1 = featureFields[1].title()
        CITY = featureFields[2]
        Subdivision = featureFields[3].title()
        WATERSHED = ReportSharedFunctions.valFromDomCode("D_Watershed",featureFields[4])

        to = memo1[0].replace(";", "\n")
        memFrom = memo1[1]


        comment = memo1[2]
        comments = comment.split("@@")

        commentContents = []

        for c in comments:
            cList = c.split("||")
            cList = [x.strip() for x in cList]
            commentContents.append(cList)

        memDate = commentContents[0][0].split(";")[0]
        try:
            memDate = ReportSharedFunctions.returnDataODBC("select FORMAT(cast('"+ memDate +"' as datetime), 'MMMM d, yyyy')")[0][0]
        except:
            pass
        if len(commentContents) > 1:
            memoDate2 = commentContents[len(commentContents)-1][0].split(";")[0]
            try:
                memoDate2 = ReportSharedFunctions.returnDataODBC("select FORMAT(cast('"+ memoDate2 +"' as datetime), 'MMMM d, yyyy')")[0][0]
            except:
                pass
            memDate += " - " + memoDate2

        incNo = INCIDENT_ID if INCIDENT_ID != None else "N/A"

        subject = ("Potential Illicit Discharge Investigation\nIncident ID No. " + incNo + "\n" + ADDRESS_LINE1 + " / " + Subdivision + "\n" 
        +  ty.title()  + " No. " + str(ftrNum) + "\n" + WATERSHED +   " Watershed / " + COUNTY + " County"  + "\n"  )#removed 10/2018 + appConfig.KCIprjs + '\n'

        msg = Paragraph('<para alignment= "LEFT"><font name=Helvetica size=12>The purpose of this Memorandum is to summarize the KCI field investigation regarding a Potential Illicit Discharge (PID) at ' + ADDRESS_LINE1 + ' in ' + COUNTY + ' County.  Photos are attached on the following pages.</font>', styles["Normal"]) 

        kciLogo = Image(appConfig.imagePath + r"\KCIlogo.jpg", width = .67*inch, height = .85*inch)
        content = [[kciLogo,"                                       MEMORANDUM"]]
        content.append(['',''])
        content.append(['TO:',to])
        content.append(['',''])
        content.append(['FROM:',memFrom])
        content.append(['',''])
        content.append(['DATE:',memDate])
        content.append(['',''])
        content.append(['SUBJECT:',subject])
        content.append(['',''])
        content.append([msg,''])
        content.append(['',''])
        reportGrid = Table(content, [1*inch,6*inch])
        reportGrid.setStyle(TableStyle([                                      
                                        ('VALIGN', (0,0), (-1, 0), 'MIDDLE'), ('VALIGN', (0,1), (-1, -1), 'TOP'),  
                                        ('BOTTOMPADDING', (0,1), (-1, -1), 0), ('SIZE', (0,0), (-1, -1), 12),  
                                        ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold') ,('FONTNAME', (1,0), (1,0), 'Helvetica-Bold') ,
                                        ('SPAN', (0,10), (-1, 10)),('LINEBELOW', (0,8),(-1,8),1,colors.black) 
                                    ]))
        elements.append(reportGrid)

#comments       
        for line in commentContents:
            commentAry = []
            pastFirst = False
            lineDateDesc = line[0].split(";",1)
            lineDate = lineDateDesc[0]
            try:
                lineDate = ReportSharedFunctions.returnDataODBC("select FORMAT(cast('"+ lineDate +"' as datetime), 'MMMM d, yyyy')")[0][0]
            except:
                pass
            lineDesc = lineDateDesc[1]
			
            content = [[lineDate]]
            descAry = []
            ary = lineDesc.split("\n")
            for a in ary:
                if a.strip() != '': descAry.append(a.strip())
            if len(descAry) > 0:
                for c in descAry:
                    content.append([ Paragraph('<para alignment= "LEFT"><font name=Helvetica size=12>' + escape(c) + '</font>', styles["Normal"]) ])#u'\u25CF'
            else:
                content.append('','No Description')

   
            innerContentCnt=0                     
            for part in line:                
                if pastFirst == True:
                    commentAry = []
                    ary = part.split("\n")
                    for a in ary:
                        if a.strip() != '': commentAry.append(a.strip())
                    if len(commentAry) > 0:
                        for c in commentAry:
                            row= [u'\u25CF', Paragraph('<para alignment= "LEFT"><font name=Helvetica size=12>' + escape(c) + '</font>', styles["Normal"]) ]
                            if innerContentCnt == 0:
                                innerContent = [row]
                            else:
                                innerContent.append(row) 
                        innerContentCnt += 1
                else:
                    pastFirst = True

            if pastFirst == True and innerContentCnt > 0:
                reportGridInner = Table(innerContent, [.25*inch,6.75*inch])
                reportGridInner.setStyle(TableStyle([ ('VALIGN', (0,0), (-1, -1), 'TOP'),  ('SIZE', (0,0), (0, 0), 12),('TOPMARGIN', (0,0), (-1, -1), 0),  ('SIZE', (0,0), (-1, -1), 7)
                                                 ]))
                content.append([reportGridInner,''])


            content.append(['',''])
            reportGrid = Table(content, [7*inch,0.05*inch])
            reportGrid.setStyle(TableStyle([ ('VALIGN', (0,0), (-1, -1), 'TOP'),  ('SIZE', (0,0), (0, 0), 12),('TOPMARGIN', (0,0), (-1, -1), 0),  ('SIZE', (0,1), (0, -1), 7),  
                                            ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold') ,('FONTNAME', (1,0), (1,0), 'Helvetica-Bold') ,
                                            ('SPAN', (0,0), (-1, 0))            ]))
                     


            elements.append(reportGrid)
        elements.append(PageBreak())

# photos frm3
        sql1 = 'SELECT atch.[data], ph.WQI_PHOTO_COMMENTS FROM [deldot_migration2].[dbo].[WQI_INVESTIGATION_PHOTOS__ATTACH_evw] atch ' 
        sql1 += 'join [deldot_migration2].[dbo].[WQI_INVESTIGATION_PHOTOS_evw] ph on ph.GlobalID = atch.[REL_GLOBALID] '
        sql1 += "join [deldot_migration2].[dbo].WQ_INVESTIGATIONS_evw wq on wq.GlobalID=ph.WQ_INVESTIGATION_ID where ph.WQI_PHOTO_CATEGORY = '"
        sql2 = "' and atch.DATA_SIZE>0 and wq.OBJECTID = " + OBJECTID 
        titles = ["Landscape", "Structure","Other"] #10/2018
        imgs = ReportSharedFunctions.returnPhotoTableWQ(titles, sql1, sql2, True, 4.25, workingFolder)

        content = [["Incident ID No. " + incNo]]
        content.append([ADDRESS_LINE1 + "/ " + Subdivision + "\n" + WATERSHED +   " Watershed/ " + COUNTY + " County"  + "\n"])
        if len(imgs)>0:
            for img in imgs:
                pass
                content.append([img])
        else:
            content.append(["No photo"])

        content.append([''])
        reportGrid = Table(content, [4.25*inch], hAlign="CENTER", repeatRows=2)
        reportGrid.setStyle(TableStyle([ ('ALIGN', (0,0), (-1, -1), 'CENTER'), ('VALIGN', (0,1), (-1, -1), 'TOP'),  ('SIZE', (0,0), (0, 0), 12),
                                        ('FONTNAME', (0,0), (0,0), 'Helvetica-Bold')  ,
                                        ('SPAN', (0,0), (-1, 0))            ]))
        elements.append(reportGrid)

        global isMemo; isMemo = True
        global memoHead; memoHead = 'Potential Illicit Discharge Investigation - Incident ID No. ' + incNo + "  " + datetime.date.today().strftime("%m/%d%Y")
        docWQ = SimpleDocTemplate(DocPath, pagesize=(8.5*inch, 11*inch), allowSplitting=1)
        docWQ.rightMargin = .75*inch
        docWQ.leftMargin =  .75*inch
        docWQ.topMargin = 50
        docWQ.bottom = 20
        docWQ.allowSplitting = 1 
        
        docWQ.build(elements, onFirstPage=myFirstPage, onLaterPages=myLaterPages)

      
    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass		
		
		
def WQI_Location(FCwqi, OBJECTID,workingPath,FEATURE_ID,INCIDENT_ID,layerOperation):
    try:
        mxd = appConfig.mxdPath  + "\\WQIdev.mxd"

        mxdDoc = arcpy.mapping.MapDocument(mxd)  
        df = arcpy.mapping.ListDataFrames(mxdDoc)[0] 
        lyr = arcpy.mapping.ListLayers(mxd, FCwqi, df)[0]

        arcpy.SelectLayerByAttribute_management(lyr, "NEW_SELECTION", ' "globalid" = ' + "'" + FEATURE_ID + "'")
        df.zoomToSelectedFeatures()

        lyrParc = arcpy.mapping.ListLayers(mxd, 'State Parcels', df)[0]#layerOperation new 1/2019
        if layerOperation == 'Nothing':
            lyrParc.visible=True 
            lyrParc.showLabels=True 	
        elif layerOperation == 'PARCEL':
            lyrParc.visible=False
        elif layerOperation == 'PARCEL_LABEL':
            lyrParc.showLabels=False		

        df.scale = 1000
        arcpy.RefreshActiveView()
        #mxdDoc.save()
        arcpy.mapping.ExportToJPEG(mxdDoc, workingPath + r"\WQIlocQQQ.jpg")		
        arcpy.mapping.ExportToJPEG(mxdDoc, workingPath + r"\WQIloc.jpg",'PAGE_LAYOUT', 
                                   df_export_width=600,df_export_height=600,world_file=False,resolution=120)
        
        
        img =  Image(workingPath + r"\WQIloc.jpg", width = 7*inch, height = 9*inch)

        content = [[INCIDENT_ID]]
        content.append([img])
        reportGrid = Table(content, [7*inch])
        reportGrid.setStyle(TableStyle([('ALIGN', (0,0), (-1, -1), 'CENTER'), 
                                      
                                        ('VALIGN', (0,0), (-1, -1), 'MIDDLE'),  
                                        ('SIZE', (0,0), (-1, -1), 11),  
                                        ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold') ,
                                        ('BOX', (0,0), (-1,-1),1, colors.black)  ,
                                                 ('BOTTOMPADDING', (0,1), (-1, -1), 0),
                                                ('TOPPADDING', (0,1), (-1, -1), 0),
                                                ('LEFTPADDING', (0,0), (-1, -1), 0),
                                                ('RIGHTPADDING', (0,0), (-1, -1), 0)
                                    ]))
        return reportGrid
    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        arcpy.AddMessage("WQI_Location " + pymsg) 
        raise
        pass

# all changes marked #9/10/19 dev update only waiting for prod
def frm1Make(OBJECTID,ty,ftrNum, INCIDENT_ID,firstMemoDate,lastMemoDate,globalid,FEATURE_ID, workingFolder,hasMemo):
    try: #http://www.fileformat.info/info/unicode/block/geometric_shapes/list.htm data.append(['','',box,u'\u25A0'])
 
        WQfields_NewforSep2019 =  ReportSharedFunctions.returnDataODBC("Select [SETTING] ,[ODOR] ,[NEAREST_STREAM]  ,[SETTING_OTHER_DESC] ,[ODOR_OTHER_DESC] FROM [deldot_migration2].[dbo].[WQ_INVESTIGATIONS_evw] where [globalid] = '" + globalid + "'")[0]
        WQfields_NewforSep2019 = ['' if x == None else x for x in WQfields_NewforSep2019] 
        SETTING_NewforSep2019 = WQfields_NewforSep2019[0]
        ODOR_NewforSep2019 = WQfields_NewforSep2019[1]
        NEAREST_STREAM_NewforSep2019 = WQfields_NewforSep2019[2]
        SETTING_OTHER_DESC_NewforSep2019 = WQfields_NewforSep2019[3]
        ODOR_OTHER_DESC_NewforSep2019 = WQfields_NewforSep2019[4]

        if ty == "STRUCTURE":
            table = "STRUCTURES_evw"
        elif ty == "CONVEYANCE":
            table = "CONVEYANCES_evw"
        elif ty == "PID_POINT":
            table = "pid_points_evw"


        WQfields =  ReportSharedFunctions.returnDataODBC("Select Determination FROM [deldot_migration2].[dbo].[WQ_INVESTIGATIONS_evw] where [globalid] = '" + globalid + "'")[0]
        WQfields = ['' if x == None else x for x in WQfields] 
        
        Determination = WQfields[0]

        elements = []

        data = [['TRACKING FORM']]
        if hasMemo:
            if firstMemoDate == lastMemoDate:
                dateTF = lastMemoDate
            else:
                dateTF = firstMemoDate + " to " + lastMemoDate

            data.append(["Incident ID No.",INCIDENT_ID,'Date:',dateTF])
        else:
            WQfields =  ReportSharedFunctions.returnDataODBC("Select (case when [CREATE_DATE] is null then '' else FORMAT([CREATE_DATE], 'MM-dd-yy') end ) as d FROM [deldot_migration2].[dbo].[WQ_INVESTIGATIONS_evw] where [globalid] = '" + globalid + "'")[0]
            data.append(["Incident ID No.",INCIDENT_ID,'Date:',WQfields[0]])
        data.append([ty + " No.",str(ftrNum),'',''])
        
        tblHead = Table(data, [1.4*inch,2.8*inch,0.8*inch,2*inch ], hAlign="CENTER")

        tblHead.setStyle(TableStyle([('ALIGN', (0,0), (-1, -1), 'CENTER'),  ('ALIGN', (0,1), (-1, -1), 'LEFT'),
                                      ('SPAN', (0,0), (-1, 0)),                                                                                          
                                    ('FONTNAME', (0,0), (-1,2), 'Helvetica-Bold'), ('FONTNAME', (1,1), (1,2), 'Helvetica'), ('FONTNAME', (3,1), (3,2), 'Helvetica'), 
                                       ('TEXTCOLOR',(0,0),(-1,2),colors.black)  ,('LINEBELOW', (0,2),(-1,2),.5,colors.darkgray) , 
                                       ('BOTTOMMARGIN', (1,2), (-1, 2), 8)  , ('TOPPADDING', (0,0), (-1, -1), 0),
                                                ('BOTTOMPADDING', (0,0), (-1, -1), 2)   
                                       # ('BOX', (0,0), (-1,-1),1, colors.black)               
                                    ])) 
        
        elements.append(tblHead)
        elements.append(Spacer(1, 0.2*inch))



        evBox1 = box  ; evBox2= box  ; evBox3=box  
        if Determination == "Evidence of Illicit Discharge":
            evBox1 = u'\u25A0' 
        elif Determination == "No Evidence of Illicit Discharge":
            evBox2= u'\u25A0' 
        elif Determination == "To Be Determined":
            evBox3= u'\u25A0' 


        data = [['EVIDENCE OF ILLICIT DISCHARGE:', evBox1,'YES', evBox2, 'NO',evBox3,'TBD']]
        tblEvid = Table(data, [3*inch,.2*inch,.8*inch,.2*inch,.8*inch,.2*inch,1.8*inch ], hAlign="CENTER")
        tblEvid.setStyle(TableStyle([ ('ALIGN', (0,0), (-1, -1), 'LEFT'),                                                
                                    ('FONTNAME', (0,0), (0,0), 'Helvetica-Bold'),  
                                     ('LINEBELOW', (0,0),(-1,0),.5,colors.darkgray)       , ('TOPPADDING', (0,0), (-1, -1), 0),
                                                ('BOTTOMPADDING', (0,0), (-1, -1), 2)             
                                    ])) 


        elements.append(tblEvid)
        elements.append(Spacer(1, 0.1*inch))


        featureFields = ReportSharedFunctions.returnDataODBC("SELECT [COUNTY], [ADDRESS_LINE1] ,[CITY] ,COALESCE(SUBDIVISION_NAME_FIELD, SUBDIVISION_NAME_GIS, 'N/A') as Subdivision, [WATERSHED] , 'pending' as Stream ,[ZIP]  FROM [deldot_migration2].[dbo]." + table + " where globalid = '" + FEATURE_ID + "'" )[0]
        featureFields = ['' if x == None else x for x in featureFields] 
        COUNTY = featureFields[0]
        COUNTY = ReportSharedFunctions.valFromDomCode("D_County",COUNTY)
        ADDRESS_LINE1 = featureFields[1]
        CITY = featureFields[2]
        Subdivision = featureFields[3]
        WATERSHED = ReportSharedFunctions.valFromDomCode("D_Watershed",featureFields[4])
        Stream = NEAREST_STREAM_NewforSep2019
        ZIP = featureFields[6]

        data = [['LOCATION:']]
        data.append(['County:',COUNTY,'Subdivision:',Subdivision])
        data.append(['Address:',ADDRESS_LINE1,'Watershed:',WATERSHED])
        data.append(['City:',CITY,'Stream:',Stream])
        data.append(['Zip Code:',ZIP,'',''])
        tblLoc = Table(data, [1*inch,2.5*inch, 1*inch,2.5*inch], hAlign="CENTER")
        tblLoc.setStyle(TableStyle([ ('ALIGN', (0,0), (-1, -1), 'LEFT'),                                                
                                    ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),  ('FONTNAME', (2,0), (2,-1), 'Helvetica-Bold'),
                                     ('LINEBELOW', (0,4),(-1,4),.5,colors.darkgray)         , ('TOPPADDING', (0,0), (-1, -1), 0),
                                                ('BOTTOMPADDING', (0,0), (-1, -1), 2)           
                                    ])) 	
        elements.append(tblLoc)
        elements.append(Spacer(1, 0.1*inch))


#setting 9/19 no longer care about fieldsheet or dump for setting
        SETTING = SETTING_NewforSep2019 
        SETTING_OTHER_DESC =  SETTING_OTHER_DESC_NewforSep2019  
        boxSD =  box;	boxINSTREAM =  box;	boxSWMPOND =  box;	boxOUTFALL =  box;	boxALONGBANK =  box;	boxROADWAY =  box;	boxMANHOLE =  box;	boxOTHER =  box;
        other = ''
        if SETTING == 'Storm Drain':
            boxSD =  u'\u25A0' 
        elif SETTING == 'In Stream':
            boxINSTREAM =  u'\u25A0' 
        elif SETTING == 'Stormwater Pond':
            boxSWMPOND = u'\u25A0' 
        elif SETTING == 'Outfall':
            boxOUTFALL = u'\u25A0' 
        elif SETTING == 'Along Bank':
            boxALONGBANK =  u'\u25A0' 
        elif SETTING == 'Roadway':
            boxROADWAY = u'\u25A0' 
        elif SETTING == 'Manhole':
            boxMANHOLE =  u'\u25A0' 
        elif SETTING == 'Other':                
            boxOTHER = u'\u25A0' 
            if SETTING_OTHER_DESC != '':
                other = SETTING_OTHER_DESC
            else:   
                other = SETTING_NewforSep2019              
            if other != '': 
                other = Paragraph('<para alignment= "LEFT"><font name=Helvetica size=10>Other&nbsp(' + other + ')</font>', styles["Normal"])  
             
        try: 
            if other == '': other = "Other" #GK added indent here BC of error
        except:
            pass	  

        hasSheet = False
        hasDump  = False
        #[SETTING_OTHER_DESC],[SETTING]  dropped from WQI_PURPOSE_DUMPING 9/19
        DRsql = "SELECT 'a' as dummy ,[OTHER_DESCRIPTION],[DUMPING_COMMENTS],[TBD],[OTHER],[ANTIFREEZE_TRANSFLUID],[COOKING_GREASE],[DETERGENT],[EXCESSIVE_DIRT],[HOME_IMPROVEMENT_WASTE],[MOTOR_OIL_FILTERS],[PAINT],[PESTICIDES_FERTILIZERS],[PET_WASTE],[SOLVENT_DEGREASER],[TRASH],[YARD_WASTE],[IMMEDIATE_ACTION_NEEDED] FROM [deldot_migration2].[dbo].[WQI_PURPOSE_DUMPING_evw] where [WQ_INVESTIGATION_ID] = '" + globalid + "'"
        arcpy.AddMessage(DRsql) 
        dumpResult = ReportSharedFunctions.returnDataODBC(DRsql)         
        if len(dumpResult) > 0:
            dumpResult = dumpResult[0]
            dumpResult = ['' if x == None else x for x in dumpResult] 
            hasDump = True

        # commented out 9/19 boxSD =  box;	boxINSTREAM =  box;	boxSWMPOND =  box;	boxOUTFALL =  box;	boxALONGBANK =  box;	boxROADWAY =  box;	boxMANHOLE =  box;	boxOTHER =  box; other=''
        #next from WQI_FIELD_SHEET_evw
        #9/19 drop [SETTING_OTHER_DESC]  [ODOR_OTHER_DESC], SETTING,[ODOR]
        QL = "SELECT  [AIR_TEMP]  , 'a' as [dummySETTING_OTHER_DESC] ,[LAND_USE_OTHER_DESC] ,'b' as [dummyODOR_OTHER_DESC] ,[DEPOSITS_OTHER_DESC] ,[STRUCTURAL_COND_OTHER_DESC] ,[VEG_COND_OTHER_DESC] ,[FLOW_COMMENTS] ,[EROSION] ,[VEG_COND] ,[DEPOSITS] ,'c' as [dummyODOR] ,[PRECIP_72HRS] ,[ALGAE_GROWTH] ,[STAINING_PRESENT] ,[LAST_RAIN_DATE] ,'d' as [dummySETTING] ,[STRUCTURAL_COND] ,[LAND_USE]  ,[FLOW_OBSERVED] ,[GlobalID] ,[REASON]   FROM [deldot_migration2].[dbo].[WQI_FIELD_SHEET_evw] where [WQ_INVESTIGATION_ID] = '" + globalid + "'"
        arcpy.AddMessage(QL)
        WQIsheet = ReportSharedFunctions.returnDataODBC(QL) 
        if WQIsheet != True and  len(WQIsheet) > 0:
            hasSheet = True
            WQIsheet = WQIsheet[0]
            WQIsheet = ['' if x == None else x for x in WQIsheet] 
            WQsheetGlobalid = WQIsheet[20]
            
            #SETTING = SETTING_NewforSep2019
            #SETTING_OTHER_DESC =  SETTING_OTHER_DESC_NewforSep2019
            #boxSD =  box;	boxINSTREAM =  box;	boxSWMPOND =  box;	boxOUTFALL =  box;	boxALONGBANK =  box;	boxROADWAY =  box;	boxMANHOLE =  box;	boxOTHER =  box;
            #other = ''
            #if SETTING == 'Storm Drain':
            #    boxSD =  u'\u25A0' 
            #elif SETTING == 'In Stream':
            #    boxINSTREAM =  u'\u25A0' 
            #elif SETTING == 'Stormwater Pond':
            #    boxSWMPOND = u'\u25A0' 
            #elif SETTING == 'Outfall':
            #    boxOUTFALL = u'\u25A0' 
            #elif SETTING == 'Along Bank':
            #    boxALONGBANK =  u'\u25A0' 
            #elif SETTING == 'Roadway':
            #    boxROADWAY = u'\u25A0' 
            #elif SETTING == 'Manhole':
            #    boxMANHOLE =  u'\u25A0' 
            #elif SETTING == 'Other':                
            #    boxOTHER = u'\u25A0' 
            #    if SETTING_OTHER_DESC != '':
            #        other = SETTING_OTHER_DESC
            #    else:
            #        #if hasDump == True: change 9/19
            #        #    if dumpResult[0] != None: other = dumpResult[0] change 9/19
            #        other=   SETTING_OTHER_DESC   # change 9/19           
            #    if other != '': 
            #        other = Paragraph('<para alignment= "LEFT"><font name=Helvetica size=10>Other&nbsp(' + other + ')</font>', styles["Normal"])  
					
            #try: 
            #    if other == '': other = "Other" #GK added indent here BC of error
            #except:
            #    pass			

        elif hasDump:

            SETTING = SETTING_NewforSep2019 
            SETTING_OTHER_DESC =  SETTING_OTHER_DESC_NewforSep2019  
            boxSD =  box;	boxINSTREAM =  box;	boxSWMPOND =  box;	boxOUTFALL =  box;	boxALONGBANK =  box;	boxROADWAY =  box;	boxMANHOLE =  box;	boxOTHER =  box;
            other = ''
            #if SETTING == 'Storm Drain':
            #    boxSD =  u'\u25A0' 
            #elif SETTING == 'In Stream':
            #    boxINSTREAM =  u'\u25A0' 
            #elif SETTING == 'Stormwater Pond':
            #    boxSWMPOND = u'\u25A0' 
            #elif SETTING == 'Outfall':
            #    boxOUTFALL = u'\u25A0' 
            #elif SETTING == 'Along Bank':
            #    boxALONGBANK =  u'\u25A0' 
            #elif SETTING == 'Roadway':
            #    boxROADWAY = u'\u25A0' 
            #elif SETTING == 'Manhole':
            #    boxMANHOLE =  u'\u25A0' 
            #elif SETTING == 'Other':                
            #    boxOTHER = u'\u25A0' 
            #    if SETTING_OTHER_DESC != '':
            #        other = SETTING_OTHER_DESC
            #    else:
            #        #if hasDump == True: change 9/19
            #        #    if dumpResult[0] != None: other = dumpResult[0]    
            #        other = SETTING_NewforSep2019              
            #    if other != '': 
            #        other = Paragraph('<para alignment= "LEFT"><font name=Helvetica size=10>Other&nbsp(' + other + ')</font>', styles["Normal"])  
             
            #try: 
            #    if other == '': other = "Other" #GK added indent here BC of error
            #except:
            #    pass	  
####################################################################################################################        
        data = [['SETTING:']]
        data.append([boxSD,'Storm Drain',boxOUTFALL,'Outfall',boxROADWAY,'Roadway'])
        data.append([boxINSTREAM,'In Stream',boxALONGBANK,'Along Bank',boxMANHOLE,'Manhole'])
        data.append([boxSWMPOND,'Stormwater Pond',boxOTHER, other])
        tbl = Table(data, [.5*inch,1.5*inch, .5*inch,2*inch, .5*inch,2*inch], hAlign="CENTER")
        tbl.setStyle(TableStyle([ ('ALIGN', (0,0), (-1, -1), 'LEFT'),                                                
                                    ('FONTNAME', (0,0), (0,-0), 'Helvetica-Bold') ,
                                     ('LINEBELOW', (0,3),(-1,3),.5,colors.darkgray) , ('SPAN',(3,3),(5,3))      , ('TOPPADDING', (0,0), (-1, -1), 0),
                                                ('BOTTOMPADDING', (0,0), (-1, -1), 2)             
                                    ])) 	
        elements.append(tbl)
        elements.append(Spacer(1, 0.1*inch))



#VISUAL
        Flowbox = box;	Stainingbox = box;	Oilbox = box;	Otherbox = box;	Floatablesbox = box;	Precipbox = box;	YardWastebox = box;	Soapbox = box;	Cloudybox = box;	Algaebox = box;

        hasFlow = False
        if hasSheet:
            sql = "SELECT  [COLOR]  ,[FLOATABLES], [COLOR_OTHER_DESC],[FLOATABLES_OTHER_DESC]  FROM [deldot_migration2].[dbo].[FLOW_CHAR_evw] where flow_id = '"  + WQsheetGlobalid + "'" 
            FlowChar = ReportSharedFunctions.returnDataODBC(sql)
            if len(FlowChar) > 0:
                hasFlow = True
                FlowChar=FlowChar[0]    
                FlowChar = ['' if x == None else x for x in FlowChar] 
                #revise dom codes > '1' to 'Yes' OTHER > Other
            other = ''    
            if hasSheet:
                if WQIsheet[21] == 'Sample Collected':
                    Flowbox = u'\u25A0'
                if WQIsheet[14] == 'Yes':
                    Stainingbox = u'\u25A0'
                if WQIsheet[13] == 'Yes':
                    Algaebox = u'\u25A0'
                if WQIsheet[12] == 'Yes':
                    Precipbox = u'\u25A0'
                if WQIsheet[10] == 'Other':
                    Otherbox = u'\u25A0'
                    if WQIsheet[4].strip() != '': other += WQIsheet[4].strip() + ", "
                
            if hasFlow:
                if FlowChar[1] == 'Suds':
                    Soapbox = u'\u25A0'
                if FlowChar[1] == 'Oil Sheen':
                    Oilbox = u'\u25A0'
                if FlowChar[1] in ['Other','Sewage','Trash']:
                    Floatablesbox = u'\u25A0'
                if FlowChar[0] == 'Cloudy':
                    Cloudybox = u'\u25A0'
                if FlowChar[1] == 'Other' :
                    Otherbox = u'\u25A0'
                    if FlowChar[3].strip() != '': other += FlowChar[3].strip() + ", "
                if FlowChar[0] == 'Other':
                    Otherbox = u'\u25A0'
                    if FlowChar[2].strip() != '': other += FlowChar[2].strip() + ", "

        if other != '':	other += ', ' #just a guess with the 9/19 change to see if since, I think, hasDump does not matter
        if hasDump:
            #if other != '':	other += ', ' change 9/19
            if dumpResult[16] == 'Yes':
                YardWastebox = u'\u25A0'
            if dumpResult[5] == 'Yes': Otherbox = u'\u25A0'; other +=  "Antifreeze/Transmission Fluid" + ", "
            if dumpResult[6] == 'Yes': Otherbox = u'\u25A0'; other +=   "Cooking Grease" + ", "
            if dumpResult[7] == 'Yes': Otherbox = u'\u25A0'; other +=   "Detergent" + ", "
            if dumpResult[8] == 'Yes': Otherbox = u'\u25A0'; other +=   "Excessive Dirt" + ", "
            if dumpResult[9] == 'Yes': Otherbox = u'\u25A0'; other +=   "Home Improvement Waste" + ", "
            if dumpResult[10] == 'Yes': Otherbox = u'\u25A0'; other +=   "Motor Oil/Oil Filter" + ", "
            if dumpResult[11] == 'Yes': Otherbox = u'\u25A0'; other +=   "Paint" + ", "
            if dumpResult[12] == 'Yes': Otherbox = u'\u25A0'; other +=   "Fertilizer" + ", "
            if dumpResult[13] == 'Yes': Otherbox = u'\u25A0'; other +=   "Pet Waste" + ", "
            if dumpResult[14] == 'Yes': Otherbox = u'\u25A0'; other +=   "Solvent/Degreaser" + ", "
            if dumpResult[15] == 'Yes': Otherbox = u'\u25A0'; other +=  "Trash"  + ", "

 
        if other != '':
            if other.endswith(", "): #10/2018 fix elsewhere too
                other = other[:-2] 
            other=other.replace(', ,',',')	
            other=other.replace(', ,',',')	
            other=other.replace(',,',',')
            other=other.replace(',,',',')	
            other = Paragraph('<para alignment= "LEFT"><font name=Helvetica size=10>Other&nbsp(' + other + ')</font>', styles["Normal"]) 
        else:
            other = 'Other'

        data = [['VISUAL:']]
        data.append([Flowbox,'Flow',Soapbox,'Soap',Cloudybox,'Cloudy'])
        data.append([Stainingbox,'Staining',Floatablesbox,'Floatables (toilet paper, etc)',Algaebox,'Algae'])
        data.append([Oilbox,'Oil / Oil Sheen',Precipbox,'Precip w/in 72 hrs ',YardWastebox,'Yard Waste'])
        data.append([Otherbox, other])
        tbl = Table(data, [.5*inch,1.5*inch, .5*inch,2*inch, .5*inch,2*inch], hAlign="CENTER")
        tbl.setStyle(TableStyle([ ('ALIGN', (0,0), (-1, -1), 'LEFT'),                                                
                                    ('VALIGN', (0,1), (-1, -1), 'TOP'), ('FONTNAME', (0,0), (0,-0), 'Helvetica-Bold') ,
                                     ('LINEBELOW', (0,4),(-1,4),.5,colors.darkgray) , ('SPAN',(1,4),(5,4))    
                                     , ('TOPPADDING', (0,0), (-1, -1), 0),
                                                ('BOTTOMPADDING', (0,0), (-1, -1), 2)          
                                    ])) 	
        elements.append(tbl)
        elements.append(Spacer(1, 0.1*inch))


#ODOR
        other = 'Other' #10/2018
        Nonebox = box;	Sulfidebox  = box;	Gas_Oilbox = box;	Sewagebox = box;	Otherbox = box;

        # if hasSheet: commented out 9/19
        ODOR = ODOR_NewforSep2019 # WQIsheet[11] 9/19
        ODOR_OTHER_DESC = ODOR_OTHER_DESC_NewforSep2019 # WQIsheet[3]
        if ODOR == 'None':
            Nonebox = u'\u25A0' 
        elif ODOR == 'Sewage':
            Sewagebox = u'\u25A0'
        elif ODOR == 'Sulfur':
            Sulfidebox = u'\u25A0'
        elif ODOR == 'Gasoline' or ODOR == 'Oil':
            Gas_Oilbox = u'\u25A0'
        elif ODOR == 'Other':
            Otherbox = u'\u25A0'
            if ODOR_OTHER_DESC != '':
                other = Paragraph('<para alignment= "LEFT"><font name=Helvetica size=10>Other&nbsp(' + ODOR_OTHER_DESC + ')</font>', styles["Normal"]) 
            else:
                other = Paragraph('<para alignment= "LEFT"><font name=Helvetica size=10>Other&nbsp(' + 'No description' + ')</font>', styles["Normal"]) 
        else:
                other = 'Other'
        #else:
        #    other = 'Other'
        #    Nonebox = u'\u25A0' #11/2018 move to prod
			
        data = [['ODOR:']]
        data.append([Nonebox,'None',Sulfidebox,'Sulfide ("rotten egg")',Gas_Oilbox,'Gas/Oil'])
        data.append([Sewagebox, 'Sewage', Otherbox, other])
        tbl = Table(data, [.5*inch,1.5*inch, .5*inch,2*inch, .5*inch,2*inch], hAlign="CENTER")
        tbl.setStyle(TableStyle([ ('ALIGN', (0,0), (-1, -1), 'LEFT'),                                                
                                    ('FONTNAME', (0,0), (0,-0), 'Helvetica-Bold') ,
                                     ('LINEBELOW', (0,2),(-1,2),.2,colors.darkgray) , ('SPAN',(3,2),(5,2))    
                                     , ('TOPPADDING', (0,0), (-1, -1), 0),
                                                ('BOTTOMPADDING', (0,0), (-1, -1), 2)          
                                    ])) 	
        elements.append(tbl)
        elements.append(Spacer(1, 0.1*inch))


        #IDDE INVESTIGATION SUMMARY  
        ReferredBy = '';   Determine = ''; 
        sql = "SELECT [DESCRIPTION] ,[DETERMINATION], [INVESTIGATION_REASON],[EXTERNAL_REPORT_TYPE] ,[REPORTED_BY], TYPE_DISCHARGE FROM [deldot_migration2].[dbo].[WQ_INVESTIGATIONS_evw] where OBJECTID = " + OBJECTID 
        res = ReportSharedFunctions.returnDataODBC(sql)[0]
        res = ['' if x == None else x for x in res]

        #old logic replaced 5/18/18 gk
        #if res[2] == 'Targeted':
        #    ReferredBy = 'Targeted'
        #elif res[4] == 'MS4 Field Crew':
        #    ReferredBy = 'MS4 Field Crew'
        #elif res[3] == 'Homeowner':
        #    ReferredBy = 'Citizen Report'
        #elif res[3] == 'DelDOT':
        #    ReferredBy = 'DelDOT'       
        #elif res[3] == 'NCCo':
        #    ReferredBy = 'NCC'
        #elif res[3] == 'Municipality':
        #    ReferredBy = 'Municipality'
        #elif res[3] == 'STOPPIT Campaign':
        #    ReferredBy = 'STOPPIT Hotline'

        #new logic   5/18/18 gk
        sql = "SELECT (case when [INVESTIGATION_REASON] = 'Targeted' then 'Targeted' when [INVESTIGATION_REASON] = 'Report' and [REPORTED_BY] <> 'External Reports' then [REPORTED_BY] when [INVESTIGATION_REASON] = 'Report' and [REPORTED_BY] = 'External Reports' then [EXTERNAL_REPORT_TYPE] else ' '  end ) "
        sql += " FROM [deldot_migration2].[dbo].[WQ_INVESTIGATIONS_evw]   where objectid  = " + OBJECTID 
        ReferredBy = ReportSharedFunctions.returnDataODBC(sql)[0][0]

        IssueInv = Paragraph('<para alignment= "LEFT"><font name=Helvetica size=10>' + res[0] + '</font>', styles["Normal"]) 

        if Determination == "Evidence of Illicit Discharge":
            DetermineTy = ReportSharedFunctions.valFromDomCode("D_TypeofIllicitDischarge",res[5])
            Determine = Paragraph('<para alignment= "LEFT"><font name=Helvetica size=10>' + Determination + "; " + DetermineTy + '</font>', styles["Normal"])
        else:
            Determine = Paragraph('<para alignment= "LEFT"><font name=Helvetica size=10>' + Determination + '</font>', styles["Normal"])
			
        sql1 = 'SELECT atch.[data], ph.WQI_PHOTO_COMMENTS FROM [deldot_migration2].[dbo].[WQI_INVESTIGATION_PHOTOS__ATTACH_evw] atch ' 
        sql1 += 'join [deldot_migration2].[dbo].[WQI_INVESTIGATION_PHOTOS_evw] ph on ph.GlobalID = atch.[REL_GLOBALID] '
        sql1 += "join [deldot_migration2].[dbo].WQ_INVESTIGATIONS_evw wq on wq.GlobalID=ph.WQ_INVESTIGATION_ID where ph.WQI_PHOTO_CATEGORY = '"
        sql2 = "' and atch.DATA_SIZE>0 and wq.OBJECTID = " + OBJECTID 

        titles = ["Landscape"]
        imgs = ReportSharedFunctions.returnPhotoTableWQ(titles, sql1, sql2, False, 2.4,workingFolder)

        if len(imgs)>0:
            photo = imgs[0]
        else:
            photo = "No photo"

        data = [['IDDE INVESTIGATION SUMMARY:','','']]
        data.append(['Referred By:',ReferredBy,photo])
        data.append(['Issue/Field Investigation:',IssueInv,''])
        data.append(['Determination:',Determine,''])

        tbl = Table(data, [2*inch,2*inch, 3*inch], hAlign="CENTER")
        tbl.setStyle(TableStyle([ ('ALIGN', (0,0), (-1, -1), 'LEFT'),   ('VALIGN', (0,0), (-1, -1), 'TOP'),                                              
                                    ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold') ,
                                     ('LINEBELOW', (0,3),(-1,3),.2,colors.darkgray) , ('SPAN',(2,1),(2,3))    
                                     , ('TOPPADDING', (0,0), (-1, -1), 0),
                                                ('BOTTOMPADDING', (0,0), (-1, -1), 2)          
                                    ])) 	
        elements.append(tbl)
        elements.append(Spacer(1, 0.1*inch))


        lbbox = u'\u25A0'; smpbox = box; fdsbox = box; ldbox=box;hdbox=box; npidbox=box; obox= box;

        cnt = ReportSharedFunctions.returnDataODBC("SELECT count(*) FROM [deldot_migration2].[dbo].[WQI_MEMOS_evw] where [WQ_INVESTIGATION_ID] = '" + globalid + "';")[0][0]
        if cnt>0:smpbox = u'\u25A0'
        
        val = ReportSharedFunctions.returnDataODBC("SELECT [FIELD_SHEET]  FROM [deldot_migration2].[dbo].[WQ_INVESTIGATIONS_evw] where globalID = '" + globalid + "';")[0][0]
        if val == 'Yes':fdsbox = u'\u25A0'

        cnt = ReportSharedFunctions.returnDataODBCcloud("SELECT count(*) FROM " + appConfig.cloudTableName + " where ATTACHMENT_TYPE = 'Lab Result' and [FEATURE_ID] = '" + globalid + "';")[0][0]
        if cnt>0:ldbox = u'\u25A0'
        
        cnt = ReportSharedFunctions.returnDataODBCcloud("SELECT count(*) FROM " + appConfig.cloudTableName + " where ATTACHMENT_TYPE in ('Door Hanger Distribution Map','Door Hanger') and [FEATURE_ID] = '" + globalid + "';")[0][0]
        if cnt>0:hdbox = u'\u25A0'

        cnt = ReportSharedFunctions.returnDataODBCcloud("SELECT count(*) FROM " + appConfig.cloudTableName + " where ATTACHMENT_TYPE in ('Notice of Violation','Notice of Potential Illegal Discharge') and [FEATURE_ID] = '" + globalid + "';")[0][0]
        if cnt>0:  npidbox = u'\u25A0'
                
        cnt = ReportSharedFunctions.returnDataODBCcloud("SELECT count(*) FROM " + appConfig.cloudTableName + " where not ATTACHMENT_TYPE in ('Lab Result','Notice of Violation','Notice of Potential Illegal Discharge','Door Hanger Distribution Map','Door Hanger') and [FEATURE_ID] = '" + globalid + "';")[0][0]
        if cnt>0:
            obox = u'\u25A0'    
            sql = "SELECT ATTACHMENT_TYPE FROM [deldot_migration2].[dbo].[FILE_ATTACHMENT_evw] where not ATTACHMENT_TYPE in ('lab result','Notice of Violation','Notice of Potential Illegal Discharge','Door Hanger Distribution Map','Door Hanger') and [FEATURE_ID] = '" + globalid + "';"          
            res = ReportSharedFunctions.returnDataODBC(sql)
            other = ''
            for r in res:
                if r[0]!=None: other += r[0] + ", "
            if other != '':
                other = other[:-2]
                other = Paragraph('<para alignment= "LEFT"><font name=Helvetica size=10>Other&nbsp(' + other + ')</font>', styles["Normal"]) 
            else:
                other = 'Other'
        else:
            other='Other' #move to prod 
                                               
        titles = ["Structure"]
        imgs = ReportSharedFunctions.returnPhotoTableWQ(titles, sql1, sql2, False,2.4,workingFolder)

        if len(imgs)>0:
            photo = imgs[0]
        else:
            photo = "No photo"



        data = [['DOCUMENTATION:','','']]
        data.append([lbbox,'Location Map',photo])
        data.append([smpbox,'Summary Memorandum with Photographs',''])
        data.append([fdsbox,'Field Data Sheet',''])
        data.append([ldbox,'Laboratory Data',''])
        data.append([hdbox,'Door Hanger and Distribution Map',''])
        data.append([npidbox,'Notice of Potential Illicit Discharge',''])
        data.append([obox,other,''])

        tbl = Table(data, [.2*inch,3.8*inch, 3*inch], hAlign="CENTER")
        tbl.setStyle(TableStyle([ ('ALIGN', (0,0), (-1, -1), 'LEFT'),   ('VALIGN', (0,0), (-1, -1), 'TOP'),                                              
                                    ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold') ,
                                     ('LINEBELOW', (0,4),(-1,3),.2,colors.darkgray) , ('SPAN',(2,1),(2,7))    
                                     , ('TOPPADDING', (0,0), (-1, -1), 0),
                                                ('BOTTOMPADDING', (0,0), (-1, -1), 2)          
                                    ])) 	
        elements.append(tbl)
        elements.append(Spacer(1, 0.1*inch))
        elements.append(PageBreak())

        return elements
       
        pass
    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        arcpy.AddMessage( pymsg)
        raise
        pass

def allProcess(OBJECTID,ty,ftrNum, INCIDENT_ID, globalid,FEATURE_ID, FCwqi, workingPath, workingFolder,layerOperation): # ,layerOperation new 1/2019
    try:        	
        RptName = "\\DelDOT_NPDES_WQ_INVESTIGATION_" + ty  + "_" +  str(ftrNum) + "_report.pdf"
             
        hasMemo = False
        cnt = ReportSharedFunctions.returnDataODBC("SELECT count(*) FROM [deldot_migration2].[dbo].[WQI_MEMOS_evw] where [WQ_INVESTIGATION_ID] = '" + globalid + "';")[0][0]
        if cnt > 0:
            hasMemo = True
            #firstMemoDate = ReportSharedFunctions.returnDataODBC("SELECT TOP 1 (case when [DATE] is null then '' else FORMAT([DATE], 'MM-dd-yy') end ) as d FROM [deldot_migration2].[dbo].[WQI_MEMOS_evw] where [WQ_INVESTIGATION_ID] = '" + globalid + "' order by [DATE];")[0][0]
            #lastMemoDate = ReportSharedFunctions.returnDataODBC("SELECT TOP 1 (case when [DATE] is null then '' else FORMAT([DATE], 'MM-dd-yy') end ) as d FROM [deldot_migration2].[dbo].[WQI_MEMOS_evw] where [WQ_INVESTIGATION_ID] = '" + globalid + "' order by [DATE] desc;")[0][0]
#10/2018 updated logic
            memoComments = ReportSharedFunctions.returnDataODBC("SELECT TOP 1 isnull([MEMO_COMMENT],'') FROM [deldot_migration2].[dbo].[WQI_MEMOS_evw] where [WQ_INVESTIGATION_ID] = '" + globalid + "' order by objectid desc;")[0][0]
            comments = memoComments.split("@@")

            commentContents = []

            for c in comments:
                cList = c.split("||")
                cList = [x.strip() for x in cList]
                commentContents.append(cList)

            firstMemoDate = commentContents[0][0].split(";")[0]
            try:
                firstMemoDate = firstMemoDate[:6] + firstMemoDate[-2:]
            except:
                pass
            if len(commentContents) > 1:
                lastMemoDate = commentContents[len(commentContents)-1][0].split(";")[0]
                try:
                    lastMemoDate = lastMemoDate[:6] + lastMemoDate[-2:]
                except:
                    pass
            else:
                lastMemoDate = firstMemoDate
        else:
            firstMemoDate = 'N/A'
            lastMemoDate = 'N/A'

        elements = []

        if INCIDENT_ID != None:
           arcpy.AddMessage("call form1")
           frm1 = frm1Make(OBJECTID,ty,ftrNum, INCIDENT_ID,firstMemoDate,lastMemoDate,globalid,FEATURE_ID,  workingFolder,hasMemo)        

           for el in frm1:
               elements.append(el)
  
           elements.append(WQI_Location(FCwqi, OBJECTID,workingPath,FEATURE_ID,INCIDENT_ID,layerOperation))
           elements.append(PageBreak())
        
        if len(elements) >0:
           arcpy.AddMessage(str(elements.count))		
           DocPath = workingFolder + "\WQ1_2.pdf"
           pdfs.append(DocPath)
           docWQ = SimpleDocTemplate(DocPath, pagesize=(8.5*inch, 11*inch), allowSplitting=1)
           docWQ.rightMargin = .75*inch
           docWQ.leftMargin =  .75*inch
           docWQ.topMargin = 30
           docWQ.bottom = 20
           docWQ.allowSplitting = 1              
           docWQ.build(elements, onFirstPage=myFirstPage, onLaterPages=myLaterPages) 
             
        if hasMemo:
           DocPath = workingFolder + "\WQ3.pdf"
           pdfs.append(DocPath)
           frm3 = frm3Make_revised1017(globalid,INCIDENT_ID,ty,FEATURE_ID,ftrNum ,OBJECTID,DocPath,workingFolder)             

        isMemo = True #GK revised for domain code update
        downnLoadPDFs(globalid,'Notice of Violation', "WQ4",workingFolder,4) #VIOLATION
        downnLoadPDFs(globalid,'Door Hanger', "WQ5",workingFolder,5)#DOORHANGER
        downnLoadPDFs(globalid,'Door Hanger Distribution Map', "WQ6",workingFolder,6)#DRHANGRMAP
        downnLoadPDFs(globalid,'lab result', "WQ7",workingFolder,7)#LABRESULT
        isMemo = False

        elements = []
        if ReportSharedFunctions.returnDataODBC("SELECT [FIELD_SHEET] from [WQ_INVESTIGATIONS_evw] where objectid = " + OBJECTID)[0][0] == 'Yes':
           cnt = ReportSharedFunctions.returnDataODBC("SELECT count(*) FROM [deldot_migration2].[dbo].[WQI_FIELD_SHEET_evw] where [WQ_INVESTIGATION_ID] = '" + globalid + "';")[0][0]
           if cnt == 0:
               tbl = Table([['FIELD_SHEET = Yes, but no WQI_FIELD_SHEET record exists']], [7*inch])
               elements.append(tbl)
           else:
               frm8 = frm8Make(globalid,INCIDENT_ID,ty,FEATURE_ID,ftrNum ,OBJECTID)
               for el in frm8:
                   elements.append(el)
            
           isMemo = False   
           DocPath = workingFolder + "\WQ8.pdf"
           pdfs.append(DocPath)
           docWQ = SimpleDocTemplate(DocPath, pagesize=(8.5*inch, 11*inch), allowSplitting=1)
           docWQ.rightMargin = .75*inch
           docWQ.leftMargin =  .75*inch
           docWQ.topMargin = 30
           docWQ.bottom = 20
           docWQ.allowSplitting = 1         
           docWQ.build(elements, onFirstPage=myFirstPage, onLaterPages=myLaterPages)         
        
        isMemo = True
        downnLoadPDFs(globalid,'Miscellaneous Water Quality Investigation Document', "WQ9",workingFolder,9)
        isMemo = False
                
        try:
            os.remove(workingFolder +  RptName)
        except OSError:
            pass  
                                      
        pdfDoc = arcpy.mapping.PDFDocumentCreate(workingFolder +  RptName)
        for pdf in pdfs:
            try: 
                pdfDoc.appendPages(pdf)
            except:
                arcpy.AddWarning("File " + pdf + " failed to append to PDF document.  Only valid PDFs can be included in this report.")
        pdfDoc.saveAndClose()             
        
    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass