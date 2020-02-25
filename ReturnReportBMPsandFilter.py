from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, inch
from reportlab.platypus import *
from reportlab.lib.styles import getSampleStyleSheet
 
import  os, sys, traceback ,io,shutil
import datetime, time
import appConfig, ReportSharedFunctions
import pypyodbc,PIL
import arcpy 

styles = getSampleStyleSheet()
ParaStyle = styles["Normal"]

def INSPECTION_DATA_SUMMARYhead():
    content = [['INSPECTION DATA SUMMARY','','']]

    reportGrid = Table(content, [2.6*inch, .96*inch, 3.25*inch])
    reportGrid.setStyle(TableStyle([('ALIGN', (0,0), (-1, -1), 'CENTER'),
                                    ('ALIGN', (0,0), (0, 0), 'LEFT'), 
                                    ('SPAN', (1,1), (2, 1)),   ('BOTTOMPADDING', (0,0), (0, 0), 10),
                                    ('VALIGN', (0,0), (-1, -1), 'MIDDLE'),  
                                    ('SIZE', (0,0), (-1, -1), 10),  
                                    ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold') ,
                                    ('BOX', (0,1), (-1,-1),1, colors.black) ,  ('INNERGRID', (0,1), (-1,-1), 0.25, colors.black)
                                ]))

    return reportGrid

    pass


def INSPECTION_DATA_SUMMARY(inspOBJECTID):
    try:     
        sql = ("SELECT [SED_INSPECTED],[SED_WATER_PRESENT],[SED_WATER_DEPTH],[SED_SEDIMENT_DEPTH],[SED_OIL_GREASE_PRESENT],[SED_COMMENTS] " +
                " ,[SND_INSPECTED],[SND_WATER_PRESENT],[SND_WATER_DEPTH],[SND_CLOGGING],[SND_DEBRIS_COVERAGE],[SND_DEBRIS_DEPTH],[SND_GRAVEL_DEPTH],[SND_SAND_DEPTH],[SND_DISCOLORATION_DEPTH],[SND_COMMENTS]  " +
                "  ,[DA_STABILIZED],[GEN_COMMENTS] " +
                " ,globalid , (case when MAINT_STATUS = 'Wet Weather Inspection Required' then 'Yes' else 'No' end) as wetInspReq  " +
                " FROM [deldot_migration2].[dbo].[BMP_DSF_INSPECTION_DRY_evw]  where OBJECTID= " + inspOBJECTID )
 
        result = ReportSharedFunctions.returnDataODBC(sql)[0]
        inspGID =  result[18] 
        Result = [' ' if x == None else x for x in result] 


        content = [['SEDIMENTATION CHAMBER']]
        content.append(['# of Chamber Inspected ','L to R', Result[0]])
        content.append(['Water Present','Y/N', Result[1]])
        if Result[1] == 'Yes':
            Result[2] = str(round(Result[2],2))
        else:
            Result[2] = ' '

        if Result[7] == 'Yes':
            Result[8] = str(round(Result[8],2))
        else:
            Result[8] = ' '

        content.append(['Depth of Water','ft', Result[2]])
        try:
            Result[3] = str(round(Result[3],2))
        except:
            pass

        content.append(['Depth of Sediment','ft', Result[3]])
        content.append(['Oil or Grease Present','Y/N', Result[4]])
        content.append(['Comments','', Paragraph('<para alignment= "LEFT"><font name=Helvetica size=10>' +Result[5] + '</font></para>', styles["Normal"])])
        content.append(['SAND CHAMBER'])
        content.append(['# of Chamber Inspected ','L to R', Result[6]])
        content.append(['Water Present ','Y/N', Result[7]])
        content.append(['Depth of Water','ft', Result[8]])

        cntWetInsp = ReportSharedFunctions.returnDataODBC("Select count(*) from [deldot_migration2].[dbo].[BMP_DSF_INSPECTION_WET_evw] where [DSF_INSPECTION_DRY_ID] = '" + inspGID + "'")[0][0]

        if cntWetInsp != 0:
            evidClog = ReportSharedFunctions.returnDataODBC("Select EVIDENCE_CLOGGING from [deldot_migration2].[dbo].[BMP_DSF_INSPECTION_WET_evw] where [DSF_INSPECTION_DRY_ID] = '" + inspGID + "'")[0][0]
            if evidClog == None:
                evidClog = ' '
        else:
            evidClog = Result[9]

        content.append(['Evidence of Clogging','Y/N', str(evidClog)])   #follow with wet inspection search

        try:
            Result[10] = str(round(Result[10],2))
        except:
            pass

        try:
            Result[11] = str(round(Result[11],2))
        except:
            pass

        try:
            Result[12] = str(round(Result[12],2))
        except:
            pass

        try:
            Result[13] = str(round(Result[13],2))
        except:
            pass

        try:
            Result[14] = str(round(Result[14],2))
        except:
            pass

        content.append(['Debris Coverage','%', Result[10]])



        content.append(['Depth of Debris','ft', Result[11]])
        content.append(['Depth of Gravel','ft', Result[12]])
        content.append(['Depth of Sand','ft', Result[13]])
        content.append(['Depth of Discoloration','ft', Result[14]])

        content.append(['Oil or Grease Present','Y/N',Result[4]]) #'??Source for This ???'
        content.append(['Comments','', Paragraph('<para alignment= "LEFT"><font name=Helvetica size=10>' +Result[15] + '</font></para>', styles["Normal"])]) 

        content.append(['Wet Weather Inspection','Y/N',Result[19]])

        content.append(['DRAINAGE AREA'])
        content.append(['Stabilized','Y/N', Result[16]])
        content.append(['OVERALL COMMENTS'])
        content.append(['','', Paragraph('<para alignment= "LEFT"><font name=Helvetica size=10>' +Result[17] + '</font></para>', styles["Normal"])]) 

        reportGrid = Table(content, [2.2*inch, .8*inch, 3.75*inch])
        reportGrid.setStyle(TableStyle([('ALIGN', (0,0), (-1, -1), 'LEFT'),  ('VALIGN', (0,0), (-1, -1), 'TOP'),
                                        ('SPAN', (0,0), (2, 0)),('SPAN', (0,7), (2, 7)),('SPAN', (0,20), (2, 20)), ('SPAN',(0,22), (2, 22)),
										 ('SPAN',(0,23), (2, 23)),										 
										('SIZE', (0,0), (-1, -1), 10),  
                                        ('FONTNAME', (0,0), (2,0), 'Helvetica-Bold') ,
                                        ('FONTNAME', (0,7), (2,7), 'Helvetica-Bold') ,
                                        ('FONTNAME', (0,20), (2,20), 'Helvetica-Bold') ,
                                        ('FONTNAME', (0,22), (2,22), 'Helvetica-Bold') ,
                                        ('BOX', (0,0), (-1,-1),1, colors.black) ,  ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black)
                                    ]))

        return reportGrid

        pass
    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass


def bmpCentoid(OBJECTID):
    try:
        for row in arcpy.da.SearchCursor(appConfig.GDBconn + r"\BMPS_evw" , ["SHAPE@"], 'OBJECTID='+ OBJECTID):
            centroid = row[0].centroid
            return [centroid.X, centroid.Y]
        pass
    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass



def startTOlocationPhoto(OBJECTID, inspYR, bmpCent,workingFolder,inspOBJECTID):
    try:# GK revise
        prevInspYearTxt="No Prior"
        prevInspYear = 1899
        sql = "select max(year(insp.[INSPECTION_DATE]))  as prevInspYear from [dbo].[BMPS_evw] bmps left join [dbo].[BMP_DSF_INSPECTION_DRY_evw]  insp on insp.BMP_ID = bmps.GlobalID where bmps.OBJECTID= " + OBJECTID + " and year(insp.[INSPECTION_DATE]) < " + inspYR
        res = ReportSharedFunctions.returnDataODBC(sql) 		
        if res != True and len(res) > 0 and res[0][0] != None:
            prevInspYear = int(res[0][0])
            prevInspYearTxt = str(prevInspYear)
        arcpy.AddMessage("PY=" +prevInspYearTxt)	

        sql = ("SELECT (case when insp.[INSPECTION_DATE] is null then '' else FORMAT(insp.[INSPECTION_DATE], 'MM/dd/yyyy' ) end) as d, insp.INSPECTION_TEAM, bmps.DISTRICT, BMPS.MAINT_AREA,BMPS.BMP_TYPE, 'ROAD_NO' as 'ROAD_NO', " +
                "insp.PERFORMANCE_RATING,  " +
                "(SELECT insp.PERFORMANCE_RATING  " +
                "  FROM [deldot_migration2].[dbo].[BMPS_evw] bmps " +
                "  join   [deldot_migration2].[dbo].[BMP_DSF_INSPECTION_DRY_evw]  insp on insp.BMP_ID = bmps.GlobalID " +
                "  where year(insp.[INSPECTION_DATE]) = {2} and bmps.OBJECTID = {0}) as previous_year, " +
                "  da.DRAINAGE_AREA_AC, '' as MAINTENANCE_WORK_ORDER,  " + 
                "(case when insp.Performance_Rating = 'B' then 'X' else '' end) as MAINTENANCE_WORK_ORDER , " +                 
                " (case when insp.PERFORMANCE_RATING = 'C' then 'X' else '' end) as CONTRACTED_WORK , " +
                " (case when insp.PERFORMANCE_RATING = 'D' then 'X' else '' end) as RETROFIT   , isnull(insp.MAINT_AREA_WORK,' ') MAINT_AREA_WORK " +
                "  FROM [deldot_migration2].[dbo].[BMPS_evw] bmps " +
                "  join   [deldot_migration2].[dbo].[BMP_DSF_INSPECTION_DRY_evw] insp on insp.BMP_ID = bmps.GlobalID " +
                "  left join [deldot_migration2].[dbo].BMP_DRAINAGE_AREA_evw da on da.BMP_ID=bmps.GlobalID " +
                "  where year(insp.[INSPECTION_DATE]) = {1} and bmps.OBJECTID = {0}; ").format(OBJECTID, inspYR, str(prevInspYear))
 

        result = ReportSharedFunctions.returnDataODBC(sql)[0]
        result = ['' if x == None else x for x in result] 
    
       
        Inspection_Date = result[0]
        Inspection_Team = result[1]
        District = result[2]
        District = District.upper() + ' DISTRICT' #GK 7/10/17
        Maintenance_Area = result[3]; 
        Maintenance_Area=Maintenance_Area.upper()		
       #move to to prod (when ... is this as intended?
        # try: 
            # Maintenance_Area = appConfig.dicAreas[Maintenance_Area]
        # except:
            # pass

            
                  
        BMP_Type = result[4] ; BMP_Type = ReportSharedFunctions.valFromDomCode("D_BMP_Type",BMP_Type);BMP_Type = BMP_Type.upper()
        BMP_Type = Paragraph('<para alignment= "CENTER"><font name=Helvetica size=10>'+ BMP_Type + '</font></para>', styles["Normal"])
        Maintenance_Road_Number = result[5]
        ThisYear_Performance_Rating_ = result[6] 
        LastYear_Performance_Rating = result[7] 
        if ThisYear_Performance_Rating_ == '': ThisYear_Performance_Rating_ = 'N/A'
        if LastYear_Performance_Rating == '': LastYear_Performance_Rating = 'N/A'
        
        if ThisYear_Performance_Rating_.startswith('A'):
            ThisYear_Performance_Rating_ = 'A - No Performance\nIssues'
        elif ThisYear_Performance_Rating_.startswith('B'):
            ThisYear_Performance_Rating_ = 'B - Minor\nMaintenance'
        elif ThisYear_Performance_Rating_.startswith('C'):
            ThisYear_Performance_Rating_ = 'C - Major\nMaintenance' 
 
        if LastYear_Performance_Rating.startswith('A'):
            LastYear_Performance_Rating = 'A - No Performance\nIssues'
        elif LastYear_Performance_Rating.startswith('B'):
            LastYear_Performance_Rating = 'B - Minor\nMaintenance'
        elif LastYear_Performance_Rating.startswith('C'):
            LastYear_Performance_Rating = 'C - Major\nMaintenance' 
			
                       
        ThisYear_Performance_Rating_ = ThisYear_Performance_Rating_.upper()   
        LastYear_Performance_Rating = LastYear_Performance_Rating.upper()    
 

        Drainage_Area_Ac = result[8]

        if Drainage_Area_Ac != '': Drainage_Area_Ac = str(round(Drainage_Area_Ac,2))

        ROAD_NO = ReportSharedFunctions.findNearestRoad(str(bmpCent[0]),str(bmpCent[1]), "centerlines")[0] 
        Maintenance_Road_Number =   ROAD_NO[0] 

        sql = ("select top 1 a.DATA from bmps_evw b join BMP_DSF_INSPECTION_DRY_evw bi on bi.BMP_ID = b.globalid " + 
            " join BMP_DSF_INSP_PHOTOS_evw p on p.BMP_DSF_INSP_ID = bi.GlobalID " +
            " join BMP_DSF_INSP_PHOTOS__ATTACH_evw a on a.REL_GLOBALID = p.GlobalID " +
            "where bi.OBJECTID = " + str(inspOBJECTID) + " and p.PHOTO_TYPE = 'Overall Landscape' " +
            " order by bi.INSPECTION_DATE desc")

        data = ReportSharedFunctions.returnDataODBC(sql)
        if len(data) == 0:
            bmpImage = Image(appConfig.imagePath + r"\bmpPlaceholder.jpg", width = 3.25*inch, height = 3.4*inch)
        else:
            try: 
                if arcpy.env.scratchWorkspace == None:
                    image = PIL.Image.open(io.BytesIO(data[0][0]))
                    image.save("c:\\temp\\bmpWaterQT.jpg")                
                    ht = (float(image.size[1])/float(image.size[0])) * 3.25
                    bmpImage = Image("c:\\temp\\bmpWaterQT.jpg", width = 3.25*inch, height = ht*inch)
                else:
                    image = PIL.Image.open(io.BytesIO(data[0][0]))
                    image.save(workingFolder + r"\bmpWaterQT.jpg")                
                    ht = (float(image.size[1])/float(image.size[0])) * 3.25
                    bmpImage = Image(workingFolder + r"\bmpWaterQT.jpg", width = 3.25*inch, height = ht*inch)
            except:
                tb = sys.exc_info()[2]
                tbinfo = traceback.format_tb(tb)[0]
                pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                        str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
                print pymsg; arcpy.AddError(pymsg)
                arcpy.AddMessage("Error retrieving bmpWaterQT photo .."  + pymsg)  
                bmpImage = "JPEG data corrupt?"    



        MAINT_AREA_WORK =  result[13]
        content = [['Inspection Date',Inspection_Date,bmpImage]]
        content.append(['Inspection Team',Inspection_Team,''])
        content.append(['District',ReportSharedFunctions.valFromDomCode("D_District",District),''])
        content.append(['Maintenance Area\n(Location)',Maintenance_Area,''])
        content.append(['Maintenance Area\n(Work Order)',MAINT_AREA_WORK,''])
        content.append(['BMP Type',BMP_Type,''])
        content.append(['Maintenance Road'+ '\n' + 'Number',Maintenance_Road_Number,''])
        content.append([str(inspYR) + '\n' + 'Performance Rating',ThisYear_Performance_Rating_,''])
        content.append([prevInspYearTxt + '\n' + 'Performance Rating',LastYear_Performance_Rating,''])
       # content.append(['Last Spillway' + '\n' + 'Inspection ',Last_Spillway_Inspection_,''])
        content.append(['Drainage Area (Ac)',str(Drainage_Area_Ac),''])

        content.append(['Action Item (s)',result[9],'MAINTENANCE WORK ORDER'])

        content.append(['',result[10],'INVASIVE SPECIES SPRAY LIST'])
        content.append(['',result[11],'CONTRACTED WORK'])
        content.append(['',result[12],'RETROFIT'])

        reportGrid = Table(content, [1.79*inch, 1.77*inch, 3.25*inch])
        reportGrid.setStyle(TableStyle([('ALIGN', (0,0), (-1, -1), 'CENTER'), 
                                        ('SPAN', (2,0), (2, 9)), 
                                        ('SPAN', (0,10), (0, 13)), 
                                        ('VALIGN', (0,0), (-1, -1), 'MIDDLE'),  
                                        ('SIZE', (0,0), (-1, -1), 10),  
                                        ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold') ,
                                        ('BOX', (0,0), (-1,-1),1, colors.black) ,  ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black)
                                        ,       ('BOTTOMPADDING', (2,0), (2, 10), 0),
                                                ('TOPPADDING', (2,0), (2, 9), 10),
                                                ('LEFTPADDING', (2,0), (2, 9), 10),
                                                ('RIGHTPADDING', (2,0), (2, 9), 10) 
                                    ]))

        return reportGrid

    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg); arcpy.AddError(pymsg) 
        raise
        pass


def BMP_Location(OBJECTID,workingFolder):
    try:
        mxd = appConfig.mxdPath  + "\\BMPS.mxd"
        #shutil.copyfile( mxdTocopyFolder + "\\mosaicSections.mxd", mxd)

        mxdDoc = arcpy.mapping.MapDocument(mxd)  
        df = arcpy.mapping.ListDataFrames(mxdDoc)[0] 
        lyr = arcpy.mapping.ListLayers(mxd, "BMPS", df)[0]

        arcpy.SelectLayerByAttribute_management(lyr, "NEW_SELECTION", ' "objectid" = ' + OBJECTID)
        df.zoomToSelectedFeatures()
        #arcpy.SelectLayerByAttribute_management(lyr, "CLEAR_SELECTION")
        df.scale = 1500
        arcpy.RefreshActiveView()
    
        arcpy.mapping.ExportToJPEG(mxdDoc, workingFolder + r"\BMPloc.jpg",'PAGE_LAYOUT', 
                                   df_export_width=600,df_export_height=600,world_file=False,resolution=120)
        
        
        img =  Image(workingFolder + r"\BMPloc.jpg", width = 5*inch, height = 4.5*inch)

        content = [["BMP Location",img,'']]

        reportGrid = Table(content, [1.81*inch, 1.75*inch, 3.25*inch])
        reportGrid.setStyle(TableStyle([('ALIGN', (0,0), (-1, -1), 'CENTER'), 
                                        ('SPAN', (1,0), (2, 0)), 
                                        ('VALIGN', (0,0), (-1, -1), 'MIDDLE'),  
                                        ('SIZE', (0,0), (-1, -1), 10),  
                                        ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold') ,
                                        ('BOX', (0,0), (-1,-1),1, colors.black) ,  ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                                                 ('BOTTOMPADDING', (0,0), (-1, -1), 0),
                                                ('TOPPADDING', (0,0), (-1, -1), 0),
                                                ('LEFTPADDING', (0,0), (-1, -1), 0),
                                                ('RIGHTPADDING', (0,0), (-1, -1), 0)
                                    ]))
        return reportGrid
    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg); arcpy.AddError(pymsg)
        raise
        pass



def allProcess(OBJECTID, inspYR,ftrNum,bmpGID,workingFolder):
    try:
        if  int(inspYR) < 2017:

            filename = "BMP_" + str(ftrNum) + ".pdf"
            rootPath = appConfig.pathForPre2017BMPpdfs
            folderByYearPath = '\\' + str(inspYR)
            filePathAndName = rootPath + folderByYearPath+'\\'+filename
            arcpy.AddMessage("This is a pre-2017 inspection and the PDF report will be a copy of "+filePathAndName)

            RptName = "\\DelDOT_NPDES_" + "BMP "  + "_" +  str(ftrNum) + "_YR" + str(inspYR) +  "_report.pdf"
            DocPath = workingFolder + RptName 

            dest = shutil.copyfile(filePathAndName, DocPath) 

            return "NA" 
        	
        inspOBJECTID = 0
        hasDryInsp = False
        sql = "SELECT objectid FROM deldot_migration2.DBO.BMP_DSF_INSPECTION_DRY_evw  where bmp_id = '" + bmpGID + "' and YEAR(inspection_date) = " + inspYR
        cursor = appConfig.connection.cursor()  
        cursor.execute(sql) 
        row = cursor.fetchone()
        if row:
            hasDryInsp = True
            inspOBJECTID = row[0]


        elements = []

        tblHeader = tblHead(inspYR,ftrNum)

        bmpCent = bmpCentoid(OBJECTID)

        elements.append(tblHeader)
        elements.append(Spacer(1, 0.05*inch))
        elements.append(startTOlocationPhoto(OBJECTID, inspYR, bmpCent,workingFolder,inspOBJECTID))
   
        elements.append(BMP_Location(OBJECTID, workingFolder))
        elements.append(PageBreak())

        hasDryInsp = False
        sql = "SELECT objectid FROM deldot_migration2.DBO.BMP_DSF_INSPECTION_DRY_evw  where bmp_id = '" + bmpGID + "' and YEAR(inspection_date) = " + inspYR
        cursor = appConfig.connection.cursor()  
        cursor.execute(sql) 
        row = cursor.fetchone()
        if row:
            hasDryInsp = True
            inspOBJECTID = row[0]
                
            elements.append(tblHeader)
            elements.append(Spacer(1, 0.05*inch))
            elements.append(INSPECTION_DATA_SUMMARYhead())
            elements.append(INSPECTION_DATA_SUMMARY(str(inspOBJECTID)))
            elements.append(PageBreak())

        #elements.append(tblHeader)
        #elements.append(Spacer(1, 0.1*inch))
        #elements.append(dbMapImage(inspYR,ftrNum))

    

        return elements   
        pass

    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg); arcpy.AddError(pymsg)
        raise
        pass



def tblHead(inspYR,ftrNum):
        tt1='DELAWARE DEPARTMENT OF TRANSPORTATION'
        tt2= inspYR + " BMP INSPECTIONS"
        ItemTitle = "BMP # " + str(ftrNum)
      
        data = [[tt1,tt2]]
        data.append([ItemTitle,''])


        tblHead = Table(data, [3.5*inch, 3.5*inch])

        tblHead.setStyle(TableStyle([('ALIGN', (0,0), (-1, -1), 'CENTER'),   
                                     ('ALIGN', (0,0), (0, 0), 'LEFT'),
                                     ('ALIGN', (1,0), (1, 0), 'RIGHT'), 
                                      ('SIZE', (1,0), (1, -1), 10),  
                                       ('SIZE', (1,1), (1, 1), 13),                                        
                                    ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'),  
                                    ('SPAN',(0,1),(1,1)),
                                    ('TOPPADDING', (0,0), (-1, -1), 1),
                                        ('BOTTOMPADDING', (0,0), (-1, -1), 1),
                                        ('BOTTOMPADDING', (0,1), (-1, 1), 1),                                   
                                       ('TEXTCOLOR',(0,0),(1,-1),colors.darkgray)                
                                    ])) # ('LINEBELOW',(0,1),(-1,1),1,colors.black)  ,
        return tblHead


def dbMapImage(inspYR,ftrNum):
    try:
        filePath = appConfig.DBmapImageRoot +  "\\BMP_NUM_" + str(ftrNum) + "_YEAR_" + str(inspYR) + ".jpg" #+ appConfig.DBmapImageBaseName
        if os.path.isfile(filePath):
            pass
        else:
            filePath = appConfig.imagePath  + "\\database_map_image_missing.jpg"  
          
        image = Image(filePath)
        ht = (float(image.imageHeight)/float(image.imageWidth )) * 6.69

        dbImage = Image(filePath, width = 6.69*inch, height = ht*inch)

        content = [['ACTION ITEM SUMMARY MAP']]
        content.append([dbImage])
            
        reportGrid = Table(content, [6.69*inch], hAlign="CENTER")
        reportGrid.setStyle(TableStyle([('ALIGN', (0,0), (-1, -1), 'LEFT'), 
                                        ('VALIGN', (0,0), (-1, -1), 'MIDDLE'),  
                                        ('SIZE', (0,0), (-1, -1), 11),  
                                        ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold') ,
                                        ('BOX', (0,1), (-1,-1),1, colors.black)  
                                        ,       ('BOTTOMPADDING', (0,1), (0, 1), 0),
                                                ('TOPPADDING', (0,1), (0, 1), 0),
                                                ('LEFTPADDING', (0,1), (0, 1), 0),
                                                ('RIGHTPADDING', (0,1), (0, 1), 0)
                                    ]))

        return reportGrid

    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        


