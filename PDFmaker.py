#copy "C:\Python27\ArcGIS10.3\Lib\site-packages\Desktop10.3.pth" to C:\ProgramData\Miniconda2\Lib\site-packages
import arcpy 
from reportlab.pdfbase import pdfmetrics
import os, sys, traceback ,  shutil,uuid
from reportlab.lib.pagesizes import letter, inch
from reportlab.platypus import *
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.rl_config import defaultPageSize
from reportlab.graphics import shapes
from reportlab.lib import colors
from reportlab.graphics.charts.textlabels import Label
import datetime, time
import ReportSharedFunctions ,appConfig 
import ReturnReportInlet

workingPath = appConfig.workingPath

PAGE_HEIGHT=defaultPageSize[1]; PAGE_WIDTH=defaultPageSize[0]

styles = getSampleStyleSheet()
HeaderStyle = styles["Heading2"]
ParaStyle = styles["Normal"]
PreStyle = styles["Code"]
styleN = styles['Normal']

pageinfo = datetime.datetime.now().strftime("%B %d, %Y  %I:%M %p")

def myFirstPage(canvas, doc):
    canvas.saveState()   
    canvas.setFont('Helvetica',9)
    canvas.drawString(inch, 0.75 * inch, pageinfo)
    canvas.restoreState()

def myLaterPages(canvas, doc):
    canvas.saveState()

    canvas.setFont("Helvetica", 8)
    canvas.drawString(inch, 0.65 * inch, "Page %d -- %s" % (doc.page , pageinfo))
    canvas.restoreState()

def tableTitle(title):
    tblTitle = Table([[title]], [7.5*inch])
    tblTitle.setStyle(TableStyle([('ALIGN', (0,0), (-1, -1), 'LEFT'),  ('BOTTOMPADDING', (0,0), (-1, -1), 8),    
                                   ('SIZE', (0,0), (-1, -1), 14),  
                                   ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold') ,
                                   ('BOX', (0,0), (-1,-1),1, colors.black) 
                                ]))
    return tblTitle


arcpy.AddMessage("start")
############# ############################################
FC = arcpy.GetParameterAsText(0)
arcpy.AddMessage(FC)
OBJECTID = arcpy.GetParameterAsText(1)
arcpy.AddMessage(OBJECTID)
workingFolder =    arcpy.env.scratchWorkspace
arcpy.env.overwriteOutput = True
arcpy.AddMessage(workingFolder)
ImagePath = appConfig.imagePath

try:

    elements = []

    if FC == "STRUCTURES":
        ftrNumNtyGID = ReportSharedFunctions.returnDataODBC("select [STRUCTURE_NUM], [STRUCTURE_TYPE], globalid from structures_evw where OBJECTID = " + OBJECTID)[0]
        ftrNum=ftrNumNtyGID[0]
        ty = "STRUCTURE "
        dom = "D_StructureTypes_1"
    elif FC == "WQ_INVESTIGATIONS": #WQ
        valu = OBJECTID.split(",")
        lenLst = len(valu)
        OBJECTID = valu[0]
        if lenLst>1:
            layerOperation = valu[1]
        else:
            layerOperation = 'Nothing'		
	
        sql = "select [FEATURE_ID]  ,  keyID, ty, [INCIDENT_ID], globalid, FC from (SELECT Q.objectID , Q.[FEATURE_ID], C.CONVEYANCE_NUM as keyID,  'CONVEYANCE' as ty, [INCIDENT_ID], Q.globalid , 'CONVEYANCES' as FC fROM [deldot_migration2].[dbo].[WQ_INVESTIGATIONS_evw]  Q join [deldot_migration2].[dbo].CONVEYANCES_evw C on c.GlobalID=Q.FEATURE_ID union all SELECT Q.objectID , Q.[FEATURE_ID], S.STRUCTURE_NUM as keyID, 'STRUCTURE' as ty, [INCIDENT_ID], Q.globalid, 'STRUCTURES' as FC  fROM [deldot_migration2].[dbo].[WQ_INVESTIGATIONS_evw] Q join [deldot_migration2].[dbo].STRUCTURES_evw S  on S.GlobalID=Q.FEATURE_ID union all SELECT Q.objectID , Q.[FEATURE_ID], convert(varchar(25),P.objectid) as keyID, 'PID_POINT' as ty, [INCIDENT_ID], Q.globalid, 'PID_POINTS' as FC   fROM [deldot_migration2].[dbo].[WQ_INVESTIGATIONS_evw] Q join [deldot_migration2].[dbo].PID_POINTS_evw P on P.GlobalID=Q.FEATURE_ID) x where  x.objectID = " + OBJECTID
        arcpy.AddMessage(sql)
        ftrNumNtyGID = ReportSharedFunctions.returnDataODBC(sql)[0]

        FEATURE_ID = ftrNumNtyGID[0]
        ftrNum=ftrNumNtyGID[1]
        ty =ftrNumNtyGID[2]
        INCIDENT_ID =ftrNumNtyGID[3]
        globalid  = ftrNumNtyGID[4]
        FCwqi   = ftrNumNtyGID[5]
        import ReturnReportWQI
        
        RptName = "\\DelDOT_NPDES_WQ_INVESTIGATION_" + ty  + "_" +  str(ftrNum) + "_report.pdf"
        DocPath = workingFolder + RptName 
        arcpy.AddMessage("WF here " + workingFolder)
        WF = workingFolder
        ReturnReportWQI.allProcess(OBJECTID,ty,ftrNum, INCIDENT_ID,globalid,FEATURE_ID,FCwqi,workingPath,WF,layerOperation)
		
    elif FC == "CONVEYANCES":  
        ftrNumNtyGID = ReportSharedFunctions.returnDataODBC("select [CONVEYANCE_NUM], [CONV_TYPE], globalid from [deldot_migration2].[dbo].CONVEYANCES_evw where OBJECTID = " + OBJECTID)[0]
        ftrNum=ftrNumNtyGID[0]
        ty = "CONVEYANCE "
        dom = "D_Conv_Type"   
    elif FC == "BMPS": #new
        objYear = OBJECTID.split(",")
        OBJECTID = objYear[0].strip()  
        yearInspc = objYear[1].strip() 
        ftrNumNtyGID = ReportSharedFunctions.returnDataODBC("select [BMP_NUM], globalid, bmp_type from [deldot_migration2].[dbo].BMPS_evw where OBJECTID = " + OBJECTID)[0]
        ftrNum=ftrNumNtyGID[0]
        ty = "BMP "   
        bmpGID = ftrNumNtyGID[1]

        bmp_type =  ftrNumNtyGID[2];
        if bmp_type != "Sand Filter":
            import ReturnReportBMP
            elements = ReturnReportBMP.allProcess(OBJECTID,yearInspc,ftrNum, bmpGID,workingFolder)                 
        else:
            import ReturnReportBMPsandFilter
            elements = ReturnReportBMPsandFilter.allProcess(OBJECTID,yearInspc,ftrNum, bmpGID,workingFolder)		
		
    #####    

    if FC == "BMPS":
        RptName =  "\\DelDOT_NPDES_" + ty  + "_" +  str(ftrNum) + "_YR" + yearInspc +  "_report.pdf"
        DocPath = workingFolder + RptName 
        doc = SimpleDocTemplate(DocPath, pagesize=(8.5*inch, 11*inch), allowSplitting=1)  
        doc.rightMargin = .5*inch
        doc.leftMargin =  .5*inch
        doc.topMargin = 30
        doc.bottom = 20
        doc.allowSplitting = 1
 

    elif FC == "CONVEYANCES" or FC ==  'STRUCTURES': #new
        subTy = ReportSharedFunctions.valFromDomCode(dom,ftrNumNtyGID[1])
        ftrGlobalID = ftrNumNtyGID[2]
###########################################################
        RptName = "\\DelDOT_NPDES_" + ty  + "_" + subTy + "_" + ftrNum + "_report.pdf"
        DocPath = workingFolder + RptName 
        doc = SimpleDocTemplate(DocPath, pagesize=(8.5*inch, 11*inch), allowSplitting=1)  
        doc.rightMargin = .5*inch
        doc.leftMargin =  .5*inch
        doc.topMargin = 36
        doc.bottom = 20
        doc.allowSplitting = 1
    ###########################################################

        tt1='DELAWARE DEPARTMENT OF TRANSPORTATION'
        tt2='NPDES DATABASE'
        ItemTitle = ty + ftrNum + " - " + subTy
        teamDdotLogo = Image(ImagePath + r"\teamDdotLogo.jpg", width = 1.8*inch, height = .78*inch)
        onlyRainLogo = Image(ImagePath + r"\onlyRainLogo.jpg", width = .92*inch, height = .92*inch)
        data = [[teamDdotLogo, tt1,onlyRainLogo]] 
        data.append(['',tt2,'']) 
        data.append(['',ItemTitle,''])


        tblHead = Table(data, [1.8*inch, 4.5*inch,  1.2*inch])

        tblHead.setStyle(TableStyle([('ALIGN', (0,0), (-1, -1), 'CENTER'),   
                                     ('ALIGN', (0,0), (0, 0), 'LEFT'),
                                     ('ALIGN', (2,0), (0, 2), 'RIGHT'), 
                                      ('SIZE', (1,0), (1, -1), 12),  
                                       ('SIZE', (1,2), (1, 2), 13),                                        
                                    ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'),  
                                    ('SPAN',(0,0),(0,2)),
                                    ('SPAN',(2,0),(2,2)),
                                    ('TOPPADDING', (0,0), (-1, -1), 1),
                                        ('BOTTOMPADDING', (0,0), (-1, -1), 1),
                                        ('BOTTOMPADDING', (0,2), (-1, 2), 6),                                   
                                        ('LINEBELOW',(0,2),(-1,2),1,colors.black)                  
                                    ]))

 

    if FC == "CONVEYANCES": #new
        import conveyances 
        titlesNelemenstAdded = conveyances.returnelements(OBJECTID,subTy)
        for titleNel in titlesNelemenstAdded:
            elements.append(tblHead)
            elements.append(Spacer(1, 0.15*inch))
            elements.append(tableTitle(titleNel[0]))
            elements.append(titleNel[1])
            elements.append(PageBreak())

        imgs = conveyances.returnPhotoelement(OBJECTID,subTy,workingFolder)
        if len(imgs)>0:
            elements.append(tblHead)
            elements.append(Spacer(1, 0.1*inch))    
            num=1
            for img in imgs:
                elements.append(img)
                if num == 1:
                    elements.append(Spacer(1, 0.2*inch))
                    num +=1
            if len(imgs)==1:#GK nov 2018 fix 
                elements.append(PageBreak())
#defect pics
        imgs = ReportSharedFunctions.returnDefectPhotoTable("conveyance", OBJECTID,workingFolder)
        if len(imgs)>0:
            elements.append(tblHead)
            elements.append(Spacer(1, 0.1*inch))    
            num=1
            for img in imgs:
                elements.append(img)
                elements.append(Spacer(1, 0.1*inch))
                num +=1

				
    elif FC == "STRUCTURES":
        elements.append(tblHead)	
        arcpy.AddMessage("STRUCTURES")	
        import structureLocation
        reportGrid = structureLocation.returnLocation_Information(OBJECTID)
        elements.append(tableTitle("Location Information"))
        elements.append(reportGrid)
        elements.append(PageBreak()) 

        if subTy == "Inlet":
		
            elements.append(tblHead)
            elements.append(Spacer(1, 0.15*inch)) 
            reportGrid = ReturnReportInlet.returnInventory_and_Inspection_Information(OBJECTID) 
            elements.append(tableTitle("Inventory and Inspection Information"))  
            elements.append(reportGrid)
            elements.append(PageBreak()) 

            arcpy.AddMessage("returnComponent_Information")			
            elements.append(tblHead)
            elements.append(Spacer(1, 0.15*inch))  
            reportGrid = ReturnReportInlet.returnComponent_Information(OBJECTID)    
            elements.append(tableTitle("Component Information"))  
            elements.append(reportGrid)
            elements.append(PageBreak()) 

            arcpy.AddMessage("OUTFALLS_evw")			
            result =  (ReportSharedFunctions.returnDataODBC("select count(*) from OUTFALLS_evw where STRUCTURE_ID  = '" + ftrGlobalID + "'")[0])
            if result[0]  > 0:
                elements.append(tblHead)
                elements.append(Spacer(1, 0.15*inch))  
                reportGrid = ReturnReportInlet.returnOutfall_Information(OBJECTID)    
                elements.append(tableTitle("Outfall Information"))  
                elements.append(reportGrid)
                elements.append(PageBreak()) 

            arcpy.AddMessage("WORK_ORDERS_evw")
            result =  (ReportSharedFunctions.returnDataODBC("select count(*) from WORK_ORDERS_evw where [STATUS] <> 'Cancelled' and  FEATURE_ID  = '" + ftrGlobalID + "'")[0])
            if result[0]  > 0:    
                elements.append(tblHead)
                elements.append(Spacer(1, 0.15*inch))  
                reportGrid = ReturnReportInlet.returnWorkOrder_Information(OBJECTID)     
                elements.append(tableTitle("Work Order Information"))  
                elements.append(reportGrid)
                elements.append(PageBreak())        

            arcpy.AddMessage("titles")
            titles = ["Landscape","Structure"]
            sql1 = ("SELECT top 1 atch.DATA  " + # pei.INSPECTION_DATE, sip.PHOTO_CATEGORY, atch.DATA, atch.DATA_SIZE
                "from [deldot_migration2].dbo.STRUCTURES_evw s " +
                "join [deldot_migration2].[dbo].inlets_evw pe on pe.STRUCTURE_ID=s.GlobalID  " +
                "join [deldot_migration2].[dbo].[inlet_INSPECTIONS_evw] pei on pei.inlet_ID = pe.GlobalID "+ 
                "left join [deldot_migration2].dbo.STRUCTURE_INSP_PHOTOS_evw  sip on sip.STRUCTURE_INSPECTION_ID = pei.GlobalID "+
                "left join [deldot_migration2].dbo.STRUCTURE_INSP_PHOTOS__ATTACH_evw atch on atch.REL_GLOBALID=sip.GlobalID  " +
                "where sip.PHOTO_CATEGORY = '")

            sql2 = ("' and atch.DATA_SIZE>0 and s.OBJECTID = " + OBJECTID + " order by pei.INSPECTION_DATE desc")

            imgs = ReportSharedFunctions.returnPhotoTable(titles, sql1, sql2, workingFolder)
            if len(imgs)>0:
                elements.append(tblHead)
                elements.append(Spacer(1, 0.1*inch))    
                num=1
                for img in imgs:
                    arcpy.AddMessage("appending images")				
                    elements.append(img)
                    if num == 1:
                        elements.append(Spacer(1, 0.2*inch))
                        num +=1
            
            arcpy.AddMessage("done inlet")

        elif subTy == "Pipe End":
            import ReturnReportPipeEnd
            elements.append(tblHead)
            elements.append(Spacer(1, 0.15*inch)) 
            reportGrid = ReturnReportPipeEnd.returnInventory_and_Inspection_Information(OBJECTID) 
            elements.append(tableTitle("Inventory and Inspection Information"))  
            elements.append(reportGrid)
            elements.append(PageBreak()) 



            result =  (ReportSharedFunctions.returnDataODBC("select count(*) from OUTFALLS_evw where STRUCTURE_ID  = '" + ftrGlobalID + "'")[0])
            if result[0]  > 0:
                elements.append(tblHead)
                elements.append(Spacer(1, 0.15*inch))  
                reportGrid = ReturnReportPipeEnd.returnOutfall_Information(OBJECTID)    
                elements.append(tableTitle("Outfall Information"))  
                elements.append(reportGrid)
                elements.append(PageBreak()) 


            result =  (ReportSharedFunctions.returnDataODBC("select count(*) from WORK_ORDERS_evw where [STATUS] <> 'Cancelled' and  FEATURE_ID  = '" + ftrGlobalID + "'")[0])
            if result[0]  > 0:    
                elements.append(tblHead)
                elements.append(Spacer(1, 0.15*inch))  
                reportGrid = ReturnReportInlet.returnWorkOrder_Information(OBJECTID)     
                elements.append(tableTitle("Work Order Information"))  
                elements.append(reportGrid)
                elements.append(PageBreak())        


            titles = ["Landscape","Structure"]
            sql1 = ("SELECT top 1 atch.DATA  " + # pei.INSPECTION_DATE, sip.PHOTO_CATEGORY, atch.DATA, atch.DATA_SIZE
                "from [deldot_migration2].dbo.STRUCTURES_evw s " +
                "join [deldot_migration2].[dbo].[PIPE_END_evw] pe on pe.STRUCTURE_ID=s.GlobalID " +
                "join [deldot_migration2].[dbo].[PIPE_END_INSPECTIONS_evw] pei on pei.PIPE_END_ID = pe.GlobalID "+ 
                "left join [deldot_migration2].dbo.STRUCTURE_INSP_PHOTOS_evw  sip on sip.STRUCTURE_INSPECTION_ID = pei.GlobalID "+
                "left join [deldot_migration2].dbo.STRUCTURE_INSP_PHOTOS__ATTACH_evw atch on atch.REL_GLOBALID=sip.GlobalID " +
                "where sip.PHOTO_CATEGORY = '")

            sql2 = ("' and atch.DATA_SIZE>0 and s.OBJECTID = " + OBJECTID + " order by pei.INSPECTION_DATE desc")

            imgs = ReportSharedFunctions.returnPhotoTable(titles, sql1, sql2, workingFolder)
            if len(imgs)>0:
                elements.append(tblHead)
                elements.append(Spacer(1, 0.1*inch))    
                num=1
                for img in imgs:
                    elements.append(img)
                    if num == 1:
                        elements.append(Spacer(1, 0.2*inch))
                        num +=1

        elif subTy == "Manhole":
            import ReturnReportManhole
            elements.append(tblHead)
            elements.append(Spacer(1, 0.15*inch)) 
            reportGrid = ReturnReportManhole.returnInventory_and_Inspection_Information(OBJECTID) 
            elements.append(tableTitle("Inventory and Inspection Information"))  
            elements.append(reportGrid)
            elements.append(PageBreak()) 
            
            
            elements.append(tblHead)
            elements.append(Spacer(1, 0.15*inch))  
            reportGrid = ReturnReportManhole.returnComponent_Information(OBJECTID)    
            elements.append(tableTitle("Component Information"))  
            elements.append(reportGrid)
            elements.append(PageBreak())

            result =  (ReportSharedFunctions.returnDataODBC("select count(*) from [deldot_migration2].[dbo].OUTFALLS_evw where STRUCTURE_ID  = '" + ftrGlobalID + "'")[0])
            if result[0]  > 0: 
                elements.append(tblHead)
                elements.append(Spacer(1, 0.15*inch))  
                reportGrid = ReturnReportManhole.returnOutfall_Information(OBJECTID)    
                elements.append(tableTitle("Outfall Information"))  
                elements.append(reportGrid)
                elements.append(PageBreak()) 


            result =  (ReportSharedFunctions.returnDataODBC("select count(*) from WORK_ORDERS_evw where [STATUS] <> 'Cancelled' and  FEATURE_ID  = '" + ftrGlobalID + "'")[0])
            if result[0]  > 0:    
                elements.append(tblHead)
                elements.append(Spacer(1, 0.15*inch))  
                reportGrid = ReturnReportInlet.returnWorkOrder_Information(OBJECTID)     
                elements.append(tableTitle("Work Order Information"))  
                elements.append(reportGrid)
                elements.append(PageBreak())        

            titles = ["Landscape","Structure"]
            sql1 = ("SELECT top 1 atch.DATA  " + # pei.INSPECTION_DATE, sip.PHOTO_CATEGORY, atch.DATA, atch.DATA_SIZE
                "from [deldot_migration2].dbo.STRUCTURES_evw s " +
                "join [deldot_migration2].[dbo].manholes_evw pe on pe.STRUCTURE_ID=s.GlobalID  " +
                "join [deldot_migration2].[dbo].[manhole_INSPECTIONS_evw] pei on pei.manhole_ID = pe.GlobalID "+ 
                "left join [deldot_migration2].dbo.STRUCTURE_INSP_PHOTOS_evw  sip on sip.STRUCTURE_INSPECTION_ID = pei.GlobalID "+
                "left join [deldot_migration2].dbo.STRUCTURE_INSP_PHOTOS__ATTACH_evw atch on atch.REL_GLOBALID=sip.GlobalID  " +
                "where sip.PHOTO_CATEGORY = '")

            sql2 = ("' and atch.DATA_SIZE>0 and s.OBJECTID = " + OBJECTID + " order by pei.INSPECTION_DATE desc")

            imgs = ReportSharedFunctions.returnPhotoTable(titles, sql1, sql2,workingFolder)
            if len(imgs)>0:
                elements.append(tblHead)
                elements.append(Spacer(1, 0.1*inch))    
                num=1
                for img in imgs:
                    elements.append(img)
                    if num == 1:
                        elements.append(Spacer(1, 0.2*inch))
                        num +=1

        elif subTy == "Swale Point":
            import ReturnReportSwalePoint
            elements.append(tblHead)
            elements.append(Spacer(1, 0.15*inch)) 
            reportGrid = ReturnReportSwalePoint.returnInventory_and_Inspection_Information(OBJECTID) 
            elements.append(tableTitle("Inventory and Inspection Information"))  
            elements.append(reportGrid)
            elements.append(PageBreak()) 

            result =  (ReportSharedFunctions.returnDataODBC("select count(*) from [deldot_migration2].[dbo].OUTFALLS_evw where STRUCTURE_ID  = '" + ftrGlobalID + "'")[0])
            if result[0]  > 0: 
                elements.append(tblHead)
                elements.append(Spacer(1, 0.15*inch))  
                reportGrid = ReturnReportSwalePoint.returnOutfall_Information(OBJECTID)    
                elements.append(tableTitle("Outfall Information"))  
                elements.append(reportGrid)
                elements.append(PageBreak()) 


            result =  (ReportSharedFunctions.returnDataODBC("select count(*) from WORK_ORDERS_evw where [STATUS] <> 'Cancelled' and  FEATURE_ID  = '" + ftrGlobalID + "'")[0])
            if result[0]  > 0:    
                elements.append(tblHead)
                elements.append(Spacer(1, 0.15*inch))  
                reportGrid = ReturnReportInlet.returnWorkOrder_Information(OBJECTID)     
                elements.append(tableTitle("Work Order Information"))  
                elements.append(reportGrid)
                elements.append(PageBreak())        

            titles = ["Landscape","Structure"]
            sql1 = ("SELECT top 1 atch.DATA  " + # pei.INSPECTION_DATE, sip.PHOTO_CATEGORY, atch.DATA, atch.DATA_SIZE
                "from [deldot_migration2].dbo.STRUCTURES_evw s join [deldot_migration2].[dbo].SWALE_CONNECTIONS_evw pe on pe.STRUCTURE_ID=s.GlobalID " +  
                "join [deldot_migration2].[dbo].SWALE_CONN_INSPECTIONS_evw pei on pei.SWALE_CONNECTION_ID = pe.GlobalID " +
                "left join [deldot_migration2].dbo.STRUCTURE_INSP_PHOTOS_evw  sip on sip.STRUCTURE_INSPECTION_ID = pei.GlobalID " +
                "left join [deldot_migration2].dbo.STRUCTURE_INSP_PHOTOS__ATTACH_evw atch on atch.REL_GLOBALID=sip.GlobalID    " +
                "where sip.PHOTO_CATEGORY = '")

            sql2 = ("' and atch.DATA_SIZE>0 and s.OBJECTID = " + OBJECTID + " order by pei.INSPECTION_DATE desc")

            imgs = ReportSharedFunctions.returnPhotoTable(titles, sql1, sql2,workingFolder )
            if len(imgs)>0:
                elements.append(tblHead)
                elements.append(Spacer(1, 0.1*inch))    
                num=1
                for img in imgs:
                    elements.append(img)
                    if num == 1:
                        elements.append(Spacer(1, 0.2*inch))
                        num +=1						

        elif subTy == "Control Structure":
            import ReturnReportControlStructure
            elements.append(tblHead)
            elements.append(Spacer(1, 0.15*inch)) 
            reportGrid = ReturnReportControlStructure.returnInventory_and_Inspection_Information(OBJECTID) 
            elements.append(tableTitle("Inventory and Inspection Information"))  
            elements.append(reportGrid)
            elements.append(PageBreak()) 

            result =  (ReportSharedFunctions.returnDataODBC("select count(*) from [deldot_migration2].[dbo].OUTFALLS_evw where STRUCTURE_ID  = '" + ftrGlobalID + "'")[0])
            if result[0] > 0: 
                elements.append(tblHead)
                elements.append(Spacer(1, 0.15*inch))  
                reportGrid = ReturnReportControlStructure.returnOutfall_Information(OBJECTID)    
                elements.append(tableTitle("Outfall Information"))  
                elements.append(reportGrid)
                elements.append(PageBreak()) 


            result =  (ReportSharedFunctions.returnDataODBC("select count(*) from WORK_ORDERS_evw where [STATUS] <> 'Cancelled' and  FEATURE_ID  = '" + ftrGlobalID + "'")[0])
            if result[0]  > 0:    
                elements.append(tblHead)
                elements.append(Spacer(1, 0.15*inch))  
                reportGrid = ReturnReportInlet.returnWorkOrder_Information(OBJECTID) #ok use inlet for all these     
                elements.append(tableTitle("Work Order Information"))  
                elements.append(reportGrid)
                elements.append(PageBreak())        

            titles = ["Landscape","Structure"]
            sql1 = ("SELECT top 1 atch.DATA  " + # pei.INSPECTION_DATE, sip.PHOTO_CATEGORY, atch.DATA, atch.DATA_SIZE
                "from [deldot_migration2].dbo.STRUCTURES_evw s "+ 
                "join [deldot_migration2].[dbo].CONTROL_STRUCTURES pe on pe.STRUCTURE_ID=s.GlobalID "+
                "join [deldot_migration2].[dbo].CONTROL_STRUCT_INSPECTIONS_evw pei on pei.CONTROL_STRUCTURE_ID = pe.GlobalID  " +
                "left join [deldot_migration2].dbo.STRUCTURE_INSP_PHOTOS_evw  sip on sip.STRUCTURE_INSPECTION_ID = pei.GlobalID " +
                "left join [deldot_migration2].dbo.STRUCTURE_INSP_PHOTOS__ATTACH_evw atch on atch.REL_GLOBALID=sip.GlobalID    " +
                "where sip.PHOTO_CATEGORY = '")

            sql2 = ("' and atch.DATA_SIZE>0 and s.OBJECTID = " + OBJECTID + " order by pei.INSPECTION_DATE desc")

            imgs = ReportSharedFunctions.returnPhotoTable(titles, sql1, sql2,workingFolder)
            if len(imgs)>0:
                elements.append(tblHead)
                elements.append(Spacer(1, 0.1*inch))    
                num=1
                for img in imgs:
                    elements.append(img)
                    if num == 1:
                        elements.append(Spacer(1, 0.2*inch))
                        num +=1

                        
        elif subTy == "Culvert Point":
            import ReturnReportCulvertPoint
            elements.append(tblHead)
            elements.append(Spacer(1, 0.15*inch)) 
            reportGrid = ReturnReportCulvertPoint.returnInventory_and_Inspection_Information(OBJECTID) 
            elements.append(tableTitle("Inventory and Inspection Information"))  
            elements.append(reportGrid)
            elements.append(PageBreak()) 


            result =  (ReportSharedFunctions.returnDataODBC("select count(*) from WORK_ORDERS_evw where [STATUS] <> 'Cancelled' and  FEATURE_ID  = '" + ftrGlobalID + "'")[0])
            if result[0]  > 0:    
                elements.append(tblHead)
                elements.append(Spacer(1, 0.15*inch))  
                reportGrid = ReturnReportInlet.returnWorkOrder_Information(OBJECTID) #ok use inlet for all these     
                elements.append(tableTitle("Work Order Information"))  
                elements.append(reportGrid)
                elements.append(PageBreak())        

            titles = ["Landscape","Structure"]
            sql1 = ("SELECT top 1 atch.DATA  " + # pei.INSPECTION_DATE, sip.PHOTO_CATEGORY, atch.DATA, atch.DATA_SIZE
                "from [deldot_migration2].dbo.STRUCTURES_evw s "+ 
                "join [deldot_migration2].[dbo].CULVERT_POINTS pe on pe.STRUCTURE_ID=s.GlobalID "+
                "join [deldot_migration2].[dbo].CULVERT_PT_INSPECTIONS_evw pei on pei.CULVERT_POINT_ID = pe.GlobalID  " +
                "left join [deldot_migration2].dbo.STRUCTURE_INSP_PHOTOS_evw  sip on sip.STRUCTURE_INSPECTION_ID = pei.GlobalID " +
                "left join [deldot_migration2].dbo.STRUCTURE_INSP_PHOTOS__ATTACH_evw atch on atch.REL_GLOBALID=sip.GlobalID    " +
                "where sip.PHOTO_CATEGORY = '")

            sql2 = ("' and atch.DATA_SIZE>0 and s.OBJECTID = " + OBJECTID + " order by pei.INSPECTION_DATE desc")

            imgs = ReportSharedFunctions.returnPhotoTable(titles, sql1, sql2,workingFolder)
            if len(imgs)>0:
                elements.append(tblHead)
                elements.append(Spacer(1, 0.1*inch))    
                num=1
                for img in imgs:
                    elements.append(img)
                    if num == 1:
                        elements.append(Spacer(1, 0.2*inch))
                        num +=1
                        
        elif subTy == "Junction Box":

            result =  (ReportSharedFunctions.returnDataODBC("select count(*) from WORK_ORDERS_evw where [STATUS] <> 'Cancelled' and  FEATURE_ID  = '" + ftrGlobalID + "'")[0])
            if result[0]  > 0:    
                elements.append(tblHead)
                elements.append(Spacer(1, 0.15*inch))  
                reportGrid = ReturnReportInlet.returnWorkOrder_Information(OBJECTID) #ok use inlet for all these     
                elements.append(tableTitle("Work Order Information"))  
                elements.append(reportGrid)
                elements.append(PageBreak())      
						
            pass

#defect pics
        arcpy.AddMessage("before defect")
        imgs = ReportSharedFunctions.returnDefectPhotoTable("structure", OBJECTID, workingFolder)
        if len(imgs)>0:
            elements.append(tblHead)
            elements.append(Spacer(1, 0.1*inch))    
            num=1
            for img in imgs:
                elements.append(img)
                elements.append(Spacer(1, 0.1*inch))
                num +=1
        arcpy.AddMessage("after defect")
#end if structure


    appConfig.connection.close()

		
    if elements == "NA":
        arcpy.SetParameterAsText(2,DocPath)
    elif FC != 'WQ_INVESTIGATIONS':
        doc.build(elements, onFirstPage=myFirstPage, onLaterPages=myLaterPages)
        arcpy.SetParameterAsText(2,DocPath)
    else:
        if os.path.isfile(DocPath):
            arcpy.SetParameterAsText(2,DocPath)

        else:
            doc = SimpleDocTemplate(DocPath, pagesize=(8.5*inch, 11*inch), allowSplitting=1)  
            doc.rightMargin = .5*inch
            doc.leftMargin =  .5*inch
            doc.topMargin = 36
            doc.bottom = 20
            doc.allowSplitting = 1
            err=Paragraph('<para alignment= "LEFT"><font name=Helvetica size=11>' + "There was no content available to populate this report.\n  If you believe this to be an error please raise it as an issue.\n(WQ Investigation -- OBJECTID = " + OBJECTID + ')</font>', styles["Normal"])
            content = []
            content.append([err])
            element = Table(content, [7.5*inch])
            elements = [element]
            doc.build(elements, onFirstPage=myFirstPage, onLaterPages=myLaterPages)
            arcpy.SetParameterAsText(2,DocPath)


except:
    tb = sys.exc_info()[2]
    tbinfo = traceback.format_tb(tb)[0]
    pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
            str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
    arcpy.AddError(pymsg)
    try:
        appConfig.connection.close()
    except:
        pass	
    pass
