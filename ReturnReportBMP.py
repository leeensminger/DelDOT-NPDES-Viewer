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
    content = [['INSPECTION DATA SUMMARY','','']] #
    pra = Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=10>SCORING EVALUATION:</font><br/><font name=Helvetica size=10>0 = Not Applicable<br/>1 = No Issues<br/>2 = Minor Issues Not Affecting Performance<br/>3 = Maintenance Work Order/Invasive Spray List<br/>4 = Contracted Work<br/>5 = Retrofit</font></para>', styles["Normal"])
    content.append(['PARAMETER', pra])


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


def action_types(inspGID, aType):
    sql = ("Select  ACTION_TYPE FROM [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] " +
         "  where BMP_INSPECTION_ID = '{0}' and  [INSPECTION_PARAMETER] =  '{1}'").format(inspGID, aType)
    
    sText = ''
    result = ReportSharedFunctions.returnDataODBC(sql)
    for r in result:
        sCode = r[0] 
        sText += ReportSharedFunctions.valFromDomCode("D_Action_Type",sCode) + ', '#GK 7/10/17
         
    if sText != '':         
        sText = sText[:-2]#GK 7/10/17
        para = Paragraph('<para alignment= "LEFT"><font name=Helvetica size=10>' + sText.upper()  + '</font></para>', styles["Normal"])
        return para
    else:
        return ''

def INSPECTION_DATA_SUMMARY(inspOBJECTID):
    try:     
        sql = ("SELECT bi.ACCESS_SCORE, bi.FENCE_SCORE, bi.OV_INVASIVE_SCORE, bi.PUB_HAZARDS_SCORE, bi.PRE_TREATMNT_SCORE, " +
                " bi.INFLOW_COND_SCORE, bi.CONV_COND_SCORE, bi.DOWNSTRM_COND_SCORE, bi.PONDING_SCORE, bi.WQ_CONTAM_SCORE, bi.WQ_TREATMENT_SCORE,  " +
                " bi.OV_SITE_STAB_SCORE, bi.PERM_POOL_SCORE, bi.BMP_VEG_SCORE, bi.UPSTRM_EMB_CVR_SCORE, bi.DOWNSTRM_EMB_CVR_SCORE,  " +
                " bi.SEEPAGE_SCORE, bi.RSR_OPENING_SCORE, bi.RSR_LOWFLOW_SCORE, bi.RSR_ACCUM_SCORE, bi.RSR_STRUCTURE_SCORE, bi.PRINCIPAL_SPILLWAY_SCORE, " +
                " bi.SPILLWAY_OUTFALL_SCORE, bi.EMERG_SPILLWAY_SCORE, bi.globalid" +
                " FROM [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] bi where bi.OBJECTID= " + inspOBJECTID )

#GK revise
        #arcpy.AddMessage("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!      " +sql)
        result = ReportSharedFunctions.returnDataODBC(sql)[0]
        inspGID =  result[24]#
        Result = [' ' if x == None else x for x in result] 
        Result = [x[:1]  for x in result] #

        Access = action_types(inspGID, 'Access')
        Fence = action_types(inspGID, 'Fence')

        sql = ("SELECT  inv.INVASIVE_TYPE, inv.INVASIVE_AREA FROM [deldot_migration2].[dbo].BMP_INSPECTION_INVASIVES_evw inv " +
        "join [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] bi on bi.GlobalID  = inv.BMP_INSPECTION_ID " +
        "where bi.OBJECTID = {0}  order by 1").format(inspOBJECTID)

        try:
            Invasive= ''
            Invasives = ReportSharedFunctions.returnDataODBC(sql) 
            for inv in Invasives:
                Invasive += inv[0] + ', '
            Invasive = Invasive[:-2]#GK 7/10/17
            Invasive = Paragraph('<para alignment= "LEFT"><font name=Helvetica size=10>' + Invasive.upper()  + '</font></para>', styles["Normal"])
        except:
            raise
            Invasive= ''
			
        PubHazards = action_types(inspGID, 'PubHazards')
        PreTreat = action_types(inspGID, 'PreTreat')
        InfCond = action_types(inspGID, 'InfCond')
        ConveyCond = action_types(inspGID, 'ConveyCond')
        DStrmCond = action_types(inspGID, 'DStrmCond')
        Ponding = action_types(inspGID, 'Ponding')
        WaterQC = action_types(inspGID, 'WaterQC')
        WaterQT = action_types(inspGID, 'WaterQT')
        OSStab = action_types(inspGID, 'OSStab')
        PermanentTool = action_types(inspGID, 'PermanentTool')
        BMPVege = action_types(inspGID, 'BMPVege')
        UpStrmEMCvr = action_types(inspGID, 'UpStrmEMCvr')
        DoStrmEMCvr = action_types(inspGID, 'DoStrmEMCvr')
        Seepage = action_types(inspGID, 'Seepage')
        RISOpening = action_types(inspGID, 'RISOpening')
        RISLowflow = action_types(inspGID, 'RISLowflow')
        RISAccumu = action_types(inspGID, 'RISAccumu')
        RISStructure = action_types(inspGID, 'RISStructure')
        PrclSpill = action_types(inspGID, 'PrclSpill')
        Spillway = action_types(inspGID, 'Spillway')
        EmgSpillway = action_types(inspGID, 'EmgSpillway')



        AccessScore =  Result[0]
        FenceScore =  Result[1]
        InvasiveScore =  Result[2]
        PubHazardsScore =  Result[3]
        PreTreatScore =  Result[4]
        InfCondScore =  Result[5]
        ConveyCondScore =  Result[6]
        DStrmCondScore =  Result[7]
        PondingScore =  Result[8]
        WaterQCScore =  Result[9]
        WaterQTScore =  Result[10]
        OSStabScore =  Result[11]
        PermanentToolScore =  Result[12]
        BMPVegeScore =  Result[13]
        UpStrmEMCvrScore =  Result[14]
        DoStrmEMCvrScore =  Result[15]
        SeepageScore =  Result[16]
        RISOpeningScore =  Result[17]
        RISLowflowScore =  Result[18]
        RISAccumuScore =  Result[19]
        RISStructureScore =  Result[20]
        PrclSpillScore =  Result[21]
        SpillwayScore =  Result[22]
        EmgSpillwayScore =  Result[23]



        content = [['SITE CONDITIONS OVERVIEW:','SCORE','COMMENT']]
        content.append(['  1.     Access',AccessScore,Access])
        content.append(['  2.     Fence',FenceScore,Fence])
        content.append(['  3.     Invasive Vegetation',InvasiveScore,Invasive])
        content.append(['  4.     Public Hazards',PubHazardsScore,PubHazards])
        content.append([Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=10>WATER QUALITY OVERVIEW</font></para>', styles["Normal"]),'SCORE','COMMENT'])
        content.append(['  5.     Pre-Treatment',PreTreatScore,PreTreat])
        content.append(['  6.     Inflow Condition',InfCondScore,InfCond])
        content.append(['  7.     Conveyance Condition',ConveyCondScore,ConveyCond])
        content.append(['  8.     Downstream Condition',DStrmCondScore,DStrmCond])
        content.append(['  9.     Ponding',PondingScore,Ponding])
        content.append(['  10.   Water Quality Contamination',WaterQCScore,WaterQC])		
        content.append(['  11.   Water Quality Treatment',WaterQTScore,WaterQT])
        content.append(['  12.   Overall Site Stability',OSStabScore,OSStab])
        content.append(['  13.   Permanent Pool',PermanentToolScore,PermanentTool])
        content.append(['  14.   BMP Vegetation',BMPVegeScore,BMPVege])
        content.append([Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=10>EMBANKMENT INSPECTION</font></para>', styles["Normal"]),'SCORE','COMMENT'])
        content.append(['  15.   Upstream Embankment Cover',UpStrmEMCvrScore,UpStrmEMCvr])
        content.append(['  16.   Downstream Embankment Cover',DoStrmEMCvrScore,DoStrmEMCvr])
        content.append(['  17.   Seepage',SeepageScore,Seepage])
        content.append([Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=10>OUTLET INSPECTION</font></para>', styles["Normal"]),'SCORE','COMMENT'])
        content.append(['  18.   Riser - Opening ',RISOpeningScore,RISOpening])
        content.append(['  19.   Riser - Low Flow ',RISLowflowScore,RISLowflow])
        content.append(['  20.   Riser - Accumulation ',RISAccumuScore,RISAccumu])
        content.append(['  21.   Riser - Structure ',RISStructureScore,RISStructure])
        content.append(['  22.   Principal Spillway',PrclSpillScore,PrclSpill])
        content.append(['  23.   Spillway Outfall',SpillwayScore,Spillway])
        content.append(['  24.   Emergency Spillway',EmgSpillwayScore,EmgSpillway])


        reportGrid = Table(content, [2.6*inch, .96*inch, 3.25*inch])
        reportGrid.setStyle(TableStyle([('ALIGN', (0,0), (-1, -1), 'LEFT'), ('ALIGN', (1,0), (1, -1), 'CENTER'), ('VALIGN', (0,0), (-1, -1), 'TOP'),
                                        ('SIZE', (0,0), (-1, -1), 10),  
                                        ('FONTNAME', (0,0), (2,0), 'Helvetica-Bold') ,
                                        ('FONTNAME', (0,5), (2,5), 'Helvetica-Bold') ,
								('FONTNAME', (0,16), (2,16), 'Helvetica-Bold') ,
                                        ('FONTNAME', (0,20), (2,20), 'Helvetica-Bold') ,
                                   ('BOTTOMPADDING', (0,0), (-1, -1), 3),
                                                ('TOPPADDING', (0,0), (-1, -1), 2),
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



def startTOlocationPhoto(OBJECTID, inspYR, bmpCent,inspOBJECTID,workingFolder):
    try:# GK revise 10/2018 replace sql
        prevInspYearTxt="No Prior"
        prevInspYear = 1899
        sql = "select max(year(insp.[INSPECTION_DATE]))  as prevInspYear from [dbo].[BMPS_evw] bmps left join [dbo].[BMP_INSPECTIONS_evw]  insp on insp.BMP_ID = bmps.GlobalID where bmps.OBJECTID= " + OBJECTID + " and year(insp.[INSPECTION_DATE]) < " + inspYR
        res = ReportSharedFunctions.returnDataODBC(sql) 		
        if res != True and len(res) > 0 and res[0][0] != None:
            prevInspYear = int(res[0][0])
            prevInspYearTxt = str(prevInspYear)
        arcpy.AddMessage("PY=" +prevInspYearTxt)	
			
        sql = ("SELECT top 1 (case when insp.[INSPECTION_DATE] is null then '' else FORMAT(insp.[INSPECTION_DATE], 'MM/dd/yyyy' ) end) as d, insp.PERSONNEL, bmps.DISTRICT, BMPS.MAINT_AREA,BMPS.BMP_TYPE, 'ROAD_NO' as 'ROAD_NO', " +
                "insp.OVERALL_CONDITION_SCORE,  " +
                "(SELECT insp.OVERALL_CONDITION_SCORE  " +
                "  FROM [deldot_migration2].[dbo].[BMPS_evw] bmps " +
                "  join   [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw]  insp on insp.BMP_ID = bmps.GlobalID " +
                "  where year(insp.[INSPECTION_DATE]) = {2} and bmps.OBJECTID = {0}) as previous_year, " +
                "insp.SPILLWAY_COAS_INSP, da.DRAINAGE_AREA_AC, " +
                "(case when insp.OVERALL_CONDITION_SCORE = 'B - Minor Maintenance' then 'X' else '' end) as MAINTENANCE_WORK_ORDER , " +
                " (case when inv.INVASIVE_VEG_SCORE = '3 - Maintenance Work Order / Invasive Spray' or  insp.OV_INVASIVE_SCORE in ('4 - Contracted Work','5 - Reftrofit') then 'X' else '' end) as INVASIVE_SPECIES_SPRAY_LIST, " + #no good need all
                " (case when insp.OVERALL_CONDITION_SCORE = 'C - Major Maintenance' then 'X' else '' end) as CONTRACTED_WORK , " +
                " (case when insp.OVERALL_CONDITION_SCORE = 'D - Retrofit' then 'X' else '' end) as RETROFIT , inv.INVASIVE_VEG_SCORE  " +
                " , (case when inv.INVASIVE_VEG_COMMENTS like '%TO BE CUT%' then 'X' else '' end) as INV_CUT    , isnull(insp.MAINT_AREA_WORK,' ') MAINT_AREA_WORK " +  
                "  FROM [deldot_migration2].[dbo].[BMPS_evw] bmps " +
                "  join   [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] insp on insp.BMP_ID = bmps.GlobalID " +
                "  left join [deldot_migration2].[dbo].BMP_DRAINAGE_AREA_evw da on da.BMP_ID=bmps.GlobalID " +
                "  left join [deldot_migration2].[dbo].BMP_INSPECTION_INVASIVES_evw inv on inv.BMP_INSPECTION_ID = insp.GlobalID " +
                "  where year(insp.[INSPECTION_DATE]) = {1} and bmps.OBJECTID = {0} order by inv.INVASIVE_VEG_SCORE desc; ").format(OBJECTID, inspYR, str(prevInspYear))
        sqlRes = sql
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
        ThisYear_Performance_Rating_ = result[6]; ThisYear_Performance_Rating_ = ReportSharedFunctions.valFromDomCode("D_Insp_Rating",ThisYear_Performance_Rating_)
        LastYear_Performance_Rating = result[7]; LastYear_Performance_Rating = ReportSharedFunctions.valFromDomCode("D_Insp_Rating",LastYear_Performance_Rating)
        if ThisYear_Performance_Rating_ == '': ThisYear_Performance_Rating_ = 'N/A'
        if LastYear_Performance_Rating == '': LastYear_Performance_Rating = 'N/A'
        
        if ThisYear_Performance_Rating_ == 'A - No Performance Issues':
            ThisYear_Performance_Rating_ = 'A - No Performance\nIssues'
        elif ThisYear_Performance_Rating_ == 'B - Minor Maintenance':
            ThisYear_Performance_Rating_ = 'B - Minor\nMaintenance'
        elif ThisYear_Performance_Rating_ == 'C - Major Maintenance':
            ThisYear_Performance_Rating_ = 'C - Major\nMaintenance' 
 
        if LastYear_Performance_Rating == 'A - No Performance Issues':
            LastYear_Performance_Rating = 'A - No Performance\nIssues'
        elif LastYear_Performance_Rating == 'B - Minor Maintenance':
            LastYear_Performance_Rating = 'B - Minor\nMaintenance'
        elif LastYear_Performance_Rating == 'C - Major Maintenance':
            LastYear_Performance_Rating = 'C - Major\nMaintenance' 

                       
        ThisYear_Performance_Rating_ = ThisYear_Performance_Rating_.upper()   
        LastYear_Performance_Rating = LastYear_Performance_Rating.upper()
     
        Last_Spillway_Inspection_ = result[8]; Last_Spillway_Inspection_ =   ReportSharedFunctions.valFromDomCode("D_BooleanNA",Last_Spillway_Inspection_) 
        if Last_Spillway_Inspection_ == "N/A":
            pass
        elif Last_Spillway_Inspection_ == "Yes":
            Last_Spillway_Inspection_ = Inspection_Date
        elif Last_Spillway_Inspection_ == "No":
            sql = ("SELECT top 1 (case when insp.[INSPECTION_DATE] is null then '' else FORMAT(insp.[INSPECTION_DATE], 'MM/dd/yyyy' ) end) as d FROM [deldot_migration2].[dbo].[BMPS_evw] bmps   join   [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] insp on insp.BMP_ID = bmps.GlobalID  left join [deldot_migration2].[dbo].BMP_DRAINAGE_AREA_evw da on da.BMP_ID=bmps.GlobalID "
              + " where insp.SPILLWAY_COAS_INSP = 'Yes' and bmps.OBJECTID = {0} and year(insp.[INSPECTION_DATE])  < {1} order by insp.[INSPECTION_DATE] desc; ").format(OBJECTID, inspYR)
            res = ReportSharedFunctions.returnDataODBC(sql)
            if res!= True and len(res) > 0:
                Last_Spillway_Inspection_ = res[0][0]
            else:
                Last_Spillway_Inspection_ = 'None'

        Drainage_Area_Ac = result[9]

        if Drainage_Area_Ac != '': Drainage_Area_Ac = str(round(Drainage_Area_Ac,2))

        ROAD_NO = ReportSharedFunctions.findNearestRoad(str(bmpCent[0]),str(bmpCent[1]), "centerlines")[0] 
        Maintenance_Road_Number =   ROAD_NO[0] 

        sql = ("select top 1 a.DATA from bmps_evw b join BMP_INSPECTIONS_evw bi on bi.BMP_ID = b.globalid " + 
            " join BMP_INSP_PHOTOS_evw p on p.BMP_INSPECTION_ID = bi.GlobalID " +
            " join BMP_INSP_PHOTOS__ATTACH_evw a on a.REL_GLOBALID = p.GlobalID " +
            "where b.OBJECTID = " + OBJECTID + " and p.PHOTO_PARAMETER_ID = 'WaterQT' and year(bi.INSPECTION_DATE) <= " + inspYR +
            " order by bi.INSPECTION_DATE desc")

        data = ReportSharedFunctions.returnDataODBC(sql)
        if len(data) == 0:
            bmpImage = Image(appConfig.imagePath + r"\bmpPlaceholder.jpg", width = 3.25*inch, height = 3.4*inch)
        else:
            try: 
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

        MAINT_AREA_WORK =  result[16]
        content = [['Inspection Date',Inspection_Date,bmpImage]]
        content.append(['Inspection Team',Inspection_Team,''])
        content.append(['District',ReportSharedFunctions.valFromDomCode("D_District",District),''])
        content.append(['Maintenance Area\n(Location)',Maintenance_Area,''])
        content.append(['Maintenance Area\n(Work Order)',MAINT_AREA_WORK,''])
        content.append(['BMP Type',BMP_Type,''])
        content.append(['Maintenance Road'+ '\n' + 'Number',Maintenance_Road_Number,''])
        content.append([str(inspYR) + '\n' + 'Performance Rating',ThisYear_Performance_Rating_,''])
        content.append([prevInspYearTxt + '\n' + 'Performance Rating',LastYear_Performance_Rating,''])
        content.append(['Last Spillway' + '\n' + 'Inspection ',Last_Spillway_Inspection_,''])
        content.append(['Drainage Area (Ac)',str(Drainage_Area_Ac),''])

        content.append(['Action Item (s)',result[10],'MAINTENANCE WORK ORDER'])

        isl = ''
        icl = ''
        if hasInvasiveSpray(inspOBJECTID):
            isl="X"
        if hasInvasiveCut(inspOBJECTID):
            icl = "X"
        content.append(['',icl,'INVASIVE SPECIES CUT LIST'])#!!!!!!!!!!!!!!!! 11 fix
        content.append(['',isl,'INVASIVE SPECIES SPRAY LIST'])
        content.append(['',result[12],'CONTRACTED WORK'])
        content.append(['',result[13],'RETROFIT'])

        reportGrid = Table(content, [1.79*inch, 1.77*inch, 3.25*inch])
        reportGrid.setStyle(TableStyle([('ALIGN', (0,0), (-1, -1), 'CENTER'), 
                                        ('SPAN', (2,0), (2, 5)), 
                                        ('SPAN', (0,11), (0, 15)), 
                                        ('VALIGN', (0,0), (-1, -1), 'MIDDLE'),  
                                        ('SIZE', (0,0), (-1, -1), 10),  
                                        ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold') ,
                                        ('BOX', (0,0), (-1,-1),1, colors.black) ,  ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black)
                                        ,       ('BOTTOMPADDING', (2,0), (2, 11), 0),
                                                ('TOPPADDING', (2,0), (2, 11), 0),
                                                ('LEFTPADDING', (2,0), (2, 11), 0),
                                                ('RIGHTPADDING', (2,0), (2, 11), 0) 
                                    ]))

        return reportGrid


    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass


def BMP_Location(OBJECTID,workingFolder):
    try:
        mxd = appConfig.mxdPath  + "\\BMPS.mxd"


        mxdDoc = arcpy.mapping.MapDocument(mxd)  
        df = arcpy.mapping.ListDataFrames(mxdDoc)[0] 
        lyr = arcpy.mapping.ListLayers(mxd, "BMPS", df)[0]

        arcpy.SelectLayerByAttribute_management(lyr, "NEW_SELECTION", ' "objectid" = ' + OBJECTID)
        df.zoomToSelectedFeatures()
        #arcpy.SelectLayerByAttribute_management(lyr, "CLEAR_SELECTION")
        df.scale = 4500
        arcpy.RefreshActiveView()
    
        arcpy.mapping.ExportToJPEG(mxdDoc, workingFolder + r"\BMPloc.jpg",'PAGE_LAYOUT', 
                                   df_export_width=600,df_export_height=600,world_file=False,resolution=120)
        
        
        img =  Image(workingFolder + r"\BMPloc.jpg", width = 3.5*inch, height = 3.15*inch)#GK nov 2018 fix 

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
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass

def getMinors(minor,minorInv):
    try:

        content = [["Minor Issue(s) Not Affecting Performance"]]
        
        invNum = 0

        for r in minorInv:#move to prod
            invNum += 1
            invTy = r[0]  
            invTy = ReportSharedFunctions.valFromDomCode("D_Invasive",invTy).upper()
            invArea = r[1]
            cmt=r[2]
            if cmt.strip()=='':
                punc = '.'
            else:
                punc = '; '
				
            arcpy.AddMessage(str(invNum) + '.&nbsp ' +invTy + "; " + invArea + punc + cmt)	
            val = str(invNum) + '.&nbsp ' +invTy + "; " + invArea + punc + cmt
            val=val.encode()			
            para = Paragraph('<para alignment= "LEFT"><font name=Helvetica size=10>'+ val +'</font></<para>', styles["Normal"])
            content.append([para])
			
        for r in minor:
            invNum += 1
            actTy = r[0]  
            actTy = ReportSharedFunctions.valFromDomCode("D_Action_Type",actTy).upper()
            act = r[1]
            act  = ReportSharedFunctions.valFromDomCode("D_Action",act).upper()
            para = Paragraph('<para alignment= "LEFT"><font name=Helvetica size=10>'+ str(invNum) + ".&nbsp "  + actTy + ';&nbsp ' + act + ';&nbsp ' + r[2]+'</font></para>', styles["Normal"])
            content.append([para])



        reportGrid = Table(content, [6.81*inch])
        reportGrid.setStyle(TableStyle([('ALIGN', (0,0), (-1, -1), 'LEFT'),('ALIGN', (0,0), (0, -0), 'CENTER'),  
                                         ('BACKGROUND',(0,0),(0,0),colors.lightgrey),
                                        ('SIZE', (0,0), (-1, -1), 10), 
                                        ('SIZE', (0,0), (0, 0), 14),
                                        ('FONTNAME', (0,0), (0,0), 'Helvetica-Bold') ,                              
                                        ('BOX', (0,0), (-1,-1),1, colors.black) ,('BOX', (0,0), (0,0),1, colors.black),  #('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                                        ('BOTTOMPADDING', (0,0), (0, 0), 6), ('TOPPADDING', (0,0), (0, -1), 6),
                                 ('BOTTOMPADDING', (0,1), (0, -1), 3), ('TOPPADDING', (0,0), (0, -1), 3)   ]))

        return reportGrid

    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass



def InvasiveSpeciesSprayList(inspOBJECTID):
    try:

        content = [["Invasive Species Spray List"]]

        sql = ("SELECT  inv.INVASIVE_TYPE, inv.INVASIVE_AREA, isnull(inv.INVASIVE_VEG_COMMENTS,'')  FROM [deldot_migration2].[dbo].BMP_INSPECTION_INVASIVES_evw inv " +
        "join [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] bi on bi.GlobalID  = inv.BMP_INSPECTION_ID " +
        "where left(isnull(inv.INVASIVE_VEG_SCORE,'x'),1) in ('3','4','5') and not isnull(inv.[INVASIVE_VEG_COMMENTS],'x')  like '%TO BE CUT%'  and bi.OBJECTID = {0} ").format(inspOBJECTID)

        result = ReportSharedFunctions.returnDataODBC(sql)
        Result = ['' if x == None else x for x in result] 
        
        invNum = 0
        for r in result:
            invNum += 1
            invTy = r[0]  
            invTy = ReportSharedFunctions.valFromDomCode("D_Invasive",invTy).upper()
            invArea = r[1]  
            invCom = r[2]
            if invCom.strip()=='':
                punc = '.'
            else:
                punc = '; '			
            if invArea == '': invArea = 0
            invArea = str(int(invArea))
            para = invTy + "<br/><br/>" +  str(invNum) + ".  " + invTy + "; AREA = " + invArea + " SQ. FT" + punc + invCom

            para = Paragraph('<para alignment= "LEFT"><font name=Helvetica size=10>' +para +  '</font></para>', styles["Normal"])
            content.append([para])
 
        reportGrid = Table(content, [6.81*inch])
        reportGrid.setStyle(TableStyle([('ALIGN', (0,0), (-1, -1), 'LEFT'),('ALIGN', (0,0), (0, -0), 'CENTER'),  
                                         ('BACKGROUND',(0,0),(0,0),colors.lightgrey),
                                        ('SIZE', (0,0), (-1, -1), 10), 
                                        ('SIZE', (0,0), (0, 0), 14),
                                        ('FONTNAME', (0,0), (0,0), 'Helvetica-Bold') ,                              
                                        ('BOX', (0,0), (-1,-1),1, colors.black) ,('BOX', (0,0), (0,0),1, colors.black) ,  #('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                                 ('BOTTOMPADDING', (0,0), (0, -1), 6),
                                                ('TOPPADDING', (0,0), (0, -1), 6)
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


def InvasiveSpeciesCutList(inspOBJECTID):
    try:

        content = [["Invasive Species Cut List"]]

        sql = ("SELECT  inv.INVASIVE_TYPE, inv.INVASIVE_AREA, isnull(inv.INVASIVE_VEG_COMMENTS,'')  FROM [deldot_migration2].[dbo].BMP_INSPECTION_INVASIVES_evw inv " +
        "join [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] bi on bi.GlobalID  = inv.BMP_INSPECTION_ID " +
        "where  inv.[INVASIVE_VEG_COMMENTS] like '%TO BE CUT%'  and bi.OBJECTID = {0} ").format(inspOBJECTID)

        result = ReportSharedFunctions.returnDataODBC(sql)
        Result = ['' if x == None else x for x in result] 
        
        invNum = 0
        for r in result:
            invNum += 1
            invTy = r[0]  
            invTy = ReportSharedFunctions.valFromDomCode("D_Invasive",invTy).upper()
            invArea = r[1]  
            invCom = r[2]
            if invCom.strip()=='':
                punc = '.'
            else:
                punc = '; '				
            if invArea == '': invArea = 0
            invArea = str(int(invArea))
            para = invTy + "<br/><br/>" +  str(invNum) + ".  " + invTy + "; AREA = " + invArea + " SQ. FT" + punc + invCom

            para = Paragraph('<para alignment= "LEFT"><font name=Helvetica size=10>' +para +  '</font></para>', styles["Normal"])
            content.append([para])
 
        reportGrid = Table(content, [6.81*inch])
        reportGrid.setStyle(TableStyle([('ALIGN', (0,0), (-1, -1), 'LEFT'),('ALIGN', (0,0), (0, -0), 'CENTER'),  
                                         ('BACKGROUND',(0,0),(0,0),colors.lightgrey),
                                        ('SIZE', (0,0), (-1, -1), 10), 
                                        ('SIZE', (0,0), (0, 0), 14),
                                        ('FONTNAME', (0,0), (0,0), 'Helvetica-Bold') ,                              
                                        ('BOX', (0,0), (-1,-1),1, colors.black) ,('BOX', (0,0), (0,0),1, colors.black)  ,  #('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                                 ('BOTTOMPADDING', (0,0), (0, -1), 6),
                                                ('TOPPADDING', (0,0), (0, -1), 6)
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
        else:
            arcpy.AddMessage("This is a post-2016 inspection and the PDF report will be generated dynamically.... ") 	
	
        hasInsp = False
        sql = "SELECT objectid FROM [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw]  where bmp_id = '" + bmpGID + "' and YEAR(inspection_date) = " + inspYR
        cursor = appConfig.connection.cursor()  
        cursor.execute(sql) 
        row = cursor.fetchone()
        if row:
            hasInsp = True
            inspOBJECTID = row[0]

        elements = []

        tblHeader = tblHead(inspYR,ftrNum)

        bmpCent = bmpCentoid(OBJECTID)

        elements.append(tblHeader)
        elements.append(Spacer(1, 0.05*inch))
        elements.append(startTOlocationPhoto(OBJECTID, inspYR, bmpCent,inspOBJECTID,workingFolder))
   
        elements.append(BMP_Location(OBJECTID,workingFolder))
        elements.append(PageBreak())

        if row:
            hasInsp = True
                
            elements.append(tblHeader)
            elements.append(Spacer(1, 0.05*inch))
            elements.append(INSPECTION_DATA_SUMMARYhead())
            elements.append(INSPECTION_DATA_SUMMARY(str(inspOBJECTID)))
            elements.append(PageBreak())

        elements.append(tblHeader)
        elements.append(Spacer(1, 0.1*inch))
        elements.append(dbMapImage(inspYR,ftrNum))

        hasBadScoresTF = hasBadScores(inspOBJECTID) 
        hasActionsTF =  hasActions(inspOBJECTID)
		
        hasInvasivesSpray = hasInvasiveSpray(inspOBJECTID)
        hasInvasivesCut = hasInvasiveCut(inspOBJECTID)

        hasInvasiveTF=hasInvasivesSpray or hasInvasivesCut


        sql = ("SELECT globalid FROM [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] where OBJECTID= " + str(inspOBJECTID) )
        inspGID =  ReportSharedFunctions.returnDataODBC(sql)[0][0]
        actionsWithScoresPopulate_bmpInspActionScores(inspGID)#5/14/2018 gk

        Minors = hasMinors(inspOBJECTID)
        minor = Minors[0]
        minorInv = Minors[1]

        hasMinor = len(minor)>0 or len(minorInv)>0

        if (hasInsp and hasBadScoresTF and hasActionsTF) or hasInvasiveTF or hasMinor:
            elements.append(Spacer(1, 0.05*inch))
            res = ACTION_ITEM_SUMMARY(str(inspOBJECTID),tblHeader)
            brk = True
            if res != None:
                elements.append(res)
            else:
                brk = False

            if hasInvasivesSpray or hasInvasivesCut  or hasMinor:

                if brk:
                    elements.append(PageBreak())

                elements.append(tblHeader)
                elements.append(Spacer(1, 0.35*inch))

                if hasInvasivesCut:
                    elements.append(InvasiveSpeciesCutList(inspOBJECTID))
                    elements.append(Spacer(1, 0.1*inch))
                    pass
                                
                if hasInvasivesSpray:
                    elements.append(InvasiveSpeciesSprayList(inspOBJECTID))
                    elements.append(Spacer(1, 0.1*inch))
                    pass

                
                if hasMinor:
                    elements.append(getMinors(minor,minorInv))
                    pass      
                        
        if len(appConfig.bmpInspActionScores) != 0:    
            score = max(appConfig.bmpInspActionScores, key=lambda x: x[0])[0]           
            elements.append(PageBreak())          
            res = ACTION_newTables(tblHeader,score)
            elements.append(res)

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


def hasMinors(inspOBJECTID):
    try:
        minorInv = []
        minor = []
#revise for prod delete . after FT
        sql = ("SELECT  rtrim(n.INVASIVE_TYPE), 'AREA = ' + format(isnull(n.INVASIVE_AREA,0), '#,###') + ' SQ. FT' as r, isnull(n.INVASIVE_VEG_COMMENTS,'') FROM [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_INVASIVES_evw] n on i.GlobalID = n.BMP_INSPECTION_ID  where [INVASIVE_VEG_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and i.OBJECTID = {0} order by 1").format(inspOBJECTID)
        result = ReportSharedFunctions.returnDataODBC(sql)
        for r in result:
            minorInv.append(r)

        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where [ACCESS_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and a.INSPECTION_PARAMETER =  'Access' and i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: minor.append(result[0])
        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where [FENCE_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and a.INSPECTION_PARAMETER =  'Fence' and i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: minor.append(result[0])
        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where [OV_INVASIVE_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and a.INSPECTION_PARAMETER =  'Invasive' and i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: minor.append(result[0])
        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where [PUB_HAZARDS_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and a.INSPECTION_PARAMETER =  'PubHazards' and i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: minor.append(result[0])
        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where [PRE_TREATMNT_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and a.INSPECTION_PARAMETER =  'PreTreat' and i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: minor.append(result[0])
        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where [INFLOW_COND_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and a.INSPECTION_PARAMETER =  'InfCond' and i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: minor.append(result[0])
        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where [CONV_COND_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and a.INSPECTION_PARAMETER =  'ConveyCond' and i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: minor.append(result[0])
        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where [DOWNSTRM_COND_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and a.INSPECTION_PARAMETER =  'DStrmCond' and i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: minor.append(result[0])
        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where [PONDING_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and a.INSPECTION_PARAMETER =  'Ponding' and i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: minor.append(result[0])
        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where [WQ_CONTAM_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and a.INSPECTION_PARAMETER =  'WaterQC' and i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: minor.append(result[0])
        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where [WQ_TREATMENT_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and a.INSPECTION_PARAMETER =  'WaterQT' and i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: minor.append(result[0])
        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where [OV_SITE_STAB_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and a.INSPECTION_PARAMETER =  'OSStab' and i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: minor.append(result[0])
        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where [PERM_POOL_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and a.INSPECTION_PARAMETER =  'PermanentTool' and i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: minor.append(result[0])
        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where [BMP_VEG_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and a.INSPECTION_PARAMETER =  'BMPVege' and i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: minor.append(result[0])
        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where [UPSTRM_EMB_CVR_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and a.INSPECTION_PARAMETER =  'UpStrmEMCvr' and i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: minor.append(result[0])
        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where [DOWNSTRM_EMB_CVR_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and a.INSPECTION_PARAMETER =  'DoStrmEMCvr' and i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: minor.append(result[0])
        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where [SEEPAGE_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and a.INSPECTION_PARAMETER =  'Seepage' and i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: minor.append(result[0])
        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where [RSR_OPENING_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and a.INSPECTION_PARAMETER =  'RISOpening' and i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: minor.append(result[0])
        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where [RSR_LOWFLOW_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and a.INSPECTION_PARAMETER =  'RISLowflow' and i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: minor.append(result[0])
        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where [RSR_ACCUM_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and a.INSPECTION_PARAMETER =  'RISAccumu' and i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: minor.append(result[0])       
        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where [RSR_STRUCTURE_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and a.INSPECTION_PARAMETER =  'RISStructure' and i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: minor.append(result[0])        
        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where [PRINCIPAL_SPILLWAY_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and a.INSPECTION_PARAMETER =  'PrclSpill' and i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: minor.append(result[0])       
        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where [SPILLWAY_OUTFALL_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and a.INSPECTION_PARAMETER =  'Spillway' and i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: minor.append(result[0])        
        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where [EMERG_SPILLWAY_SCORE] in ('1 - No Issues','2 - Minor Issues not Affecting Performance') and a.INSPECTION_PARAMETER =  'EmgSpillway' and i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: minor.append(result[0])
        
        minorsAll = []
        #need to remove the # score
        for m in minor:
            sComm = m[2]  
            if len(sComm)> 0:
                ScoreAndcomment = actionScoreAndComment(sComm)  
                score = ScoreAndcomment[0]
                sComm = ScoreAndcomment[1]
                m =  [m[0],m[1],sComm]
            minorsAll.append(m)


        #make a seperate for minors (need to change sequenciing for non-minor??)
        result = ReportSharedFunctions.returnDataODBC(("SELECT isnull(a.ACTION_TYPE,''), isnull(a.[ACTION],''), isnull(a.INSP_ACTION_COMMENT ,'') from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] a on a.BMP_INSPECTION_ID = i.GlobalID where  i.OBJECTID = {0}").format(inspOBJECTID))
        if len(result)>0: 
            for rec in result:
                sComm = rec[2]  
                if len(sComm)> 0:
                    ScoreAndcomment = actionScoreAndComment(sComm)  
                    score = ScoreAndcomment[0]
                    sComm = ScoreAndcomment[1]
                    if score == '2':
                        rec =  [rec[0],rec[1],sComm]
                        if not rec in minorsAll:
                            minorsAll.append(rec)

        return [minorsAll,minorInv]

    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg); arcpy.AddError(pymsg)
        pass



def hasInvasiveSpray(inspOBJECTID):
    try:#revise
        sql = ("SELECT  count(*) FROM [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_INVASIVES_evw] n on i.GlobalID = n.BMP_INSPECTION_ID  where [INVASIVE_VEG_SCORE] in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit') and not isnull([INVASIVE_VEG_COMMENTS],'x') like '%TO Be CUT%' and i.OBJECTID = {0}").format(inspOBJECTID)

        result = ReportSharedFunctions.returnDataODBC(sql)[0][0]

        if result > 0:
            return True
        else:
            return False

    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg); arcpy.AddError(pymsg)
        pass


def hasInvasiveCut(inspOBJECTID):
    try:#revise
        sql = ("SELECT  count(*) FROM [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] i join [deldot_migration2].[dbo].[BMP_INSPECTION_INVASIVES_evw] n on i.GlobalID = n.BMP_INSPECTION_ID  where [INVASIVE_VEG_COMMENTS] like '%TO Be CUT%' and i.OBJECTID = {0}").format(inspOBJECTID)

        result = ReportSharedFunctions.returnDataODBC(sql)[0][0]

        if result > 0:
            return True
        else:
            return False

    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg); arcpy.AddError(pymsg)
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
        

def hasBadScores(inspOBJECTID): #revise
    sql = ("SELECT count(*) from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] where  " +
        "OBJECTID = {0} and  " +
        "(ACCESS_SCORE  in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit') or  " +
        "FENCE_SCORE  in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit') or  " +
        "OV_INVASIVE_SCORE  in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit') or  " +
        "PUB_HAZARDS_SCORE  in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit') or  " +
        "PRE_TREATMNT_SCORE   in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit') or  " +
        " INFLOW_COND_SCORE   in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit') or  " +
        " CONV_COND_SCORE   in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit') or  " +
        " DOWNSTRM_COND_SCORE   in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit') or  " +
        " PONDING_SCORE   in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit') or  " +
        " WQ_CONTAM_SCORE   in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit') or  " +
        " WQ_TREATMENT_SCORE   in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit') or  " +
        " OV_SITE_STAB_SCORE   in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit') or  " +
        " PERM_POOL_SCORE   in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit') or  " +
        " BMP_VEG_SCORE   in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit') or  " +
        " UPSTRM_EMB_CVR_SCORE   in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit') or  " +
        " DOWNSTRM_EMB_CVR_SCORE   in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit') or  " +
        " SEEPAGE_SCORE   in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit') or  " +
        " RSR_OPENING_SCORE   in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit') or  " +
        " RSR_LOWFLOW_SCORE   in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit') or  " +
        " RSR_ACCUM_SCORE   in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit') or  " +
        " RSR_STRUCTURE_SCORE   in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit') or  " +
        "  PRINCIPAL_SPILLWAY_SCORE   in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit') or  " +
        " SPILLWAY_OUTFALL_SCORE   in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit') or  " +
        " EMERG_SPILLWAY_SCORE   in ('3 - Maintenance Work Order / Invasive Spray','4 - Contracted Work','5 - Reftrofit'))  " ).format(inspOBJECTID)

    result = ReportSharedFunctions.returnDataODBC(sql)[0][0]

    if result > 0:
        return True
    else:
        return False

pnum = 0
#from reportlab.lib.utils import ImageReader
def returnActionImage(AcGID):
    try:
        workingFolder =    arcpy.env.scratchWorkspace
        global pnum
        pnum +=1
        sql = ("Select  atch.DATA FROM [deldot_migration2].[dbo].BMP_INSP_ACTION_PHOTOS_evw  ap " +
        "join [deldot_migration2].[dbo].BMP_INSP_ACTION_PHOTOS__ATTACH_evw atch  on atch.REL_GLOBALID = ap.GlobalID  " +
        "  where ap.BMP_INSPECTION_ACTION_ID = '{0}'").format(AcGID)

        cursor = appConfig.connection.cursor()  
        cursor.execute(sql) 
        row = cursor.fetchone()
        if row:
            img = row[0]

            image = PIL.Image.open(io.BytesIO(img))

            image.save(workingFolder + "\\P" + str(pnum) + "ac.jpg")
            ht = (float(image.size[1])/float(image.size[0])) * 3.4

            #reportlab_pil_img = ImageReader(image)
            #img2 = Image(reportlab_pil_img, 3.4*inch, ht*inch)
            img = Image(workingFolder + "\\P" + str(pnum) + "ac.jpg", 3.4*inch, ht*inch)

            return img
        #sample image was 3.12w X 2.34h, resized to 3.4w = 2.55h
        #let's try just resize for width each photo and see what we get

        pass

    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
 






def hasActions(inspOBJECTID):
    sql = ("SELECT count(*) from [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw]  a   " +
    "join [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] bi on bi.GlobalID=a.BMP_INSPECTION_ID "
    "where bi.OBJECTID = {0}  and a.INSPECTION_PARAMETER in ('Access', 'Fence', 'Invasive', 'PubHazards', 'PreTreat', 'InfCond', 'ConveyCond', 'DStrmCond', " + 
    "'Ponding', 'WaterQC', 'WaterQT', 'OSStab', 'PermanentTool', 'BMPVege', 'UpStrmEMCvr', 'DoStrmEMCvr', 'Seepage', 'RISOpening', 'RISLowflow', 'RISAccumu', " + 
    " 'RISStructure', 'PrclSpill', 'Spillway', 'EmgSpillway') "  ).format(inspOBJECTID)

    result = ReportSharedFunctions.returnDataODBC(sql)[0][0]

    if result > 0:
        return True
    else:
        return False


def actionScoreAndComment(ScoreAndcomment):  #gk
    score = 'N/A'
    comment = ''
    lstScores = ['1','2','3','4','5']

    if ScoreAndcomment.startswith('@'):
        firstTest = ScoreAndcomment.find('@')
        at2ndidx = ScoreAndcomment.find('@', firstTest + 1)

        if at2ndidx == -1:
            comment = ScoreAndcomment
        else:
            lst = list(ScoreAndcomment[1:at2ndidx])
            lst.reverse
            for s in lst:
                if s in lstScores:
                    score = s
                    break

            if len(ScoreAndcomment) > at2ndidx+1:
                comment = ScoreAndcomment[- (len(ScoreAndcomment) - at2ndidx - 1):]
    else:
        comment = ScoreAndcomment

    return [score,comment]


def actionsWithScoresPopulate_bmpInspActionScores(inspGID):  
    sql = ("Select  ACTION_TYPE, ACTION, INSP_ACTION_COMMENT, globalid FROM [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] " +
         "  where BMP_INSPECTION_ID = '{0}' ").format(inspGID)

    result = ReportSharedFunctions.returnDataODBC(sql)

    lst= []
 
    for r in result:
        r = ['' if x == None else x for x in r]
        sCode = r[0] 
        sAty = ReportSharedFunctions.valFromDomCode("D_Action_Type",sCode).upper()  
        sCode = r[1] 
        sAc = ReportSharedFunctions.valFromDomCode("D_Action",sCode).upper()
        sComm = r[2]  
        if len(sComm)> 0:
            ScoreAndcomment = actionScoreAndComment(sComm)  
            score = ScoreAndcomment[0]
            sComm = ScoreAndcomment[1]

        #aTyNum = str(seqNum) + ".  " #need aTyNum?
        para = Paragraph('<para alignment= "LEFT"><font name=Helvetica size=10>' + sAty + '<br/><br/>' + sAc  + '<br/><br/>' + " - "  + sComm.upper() +  '</font></para>', styles["Normal"])
        pic = 'No Picture'
        #now try to get the pic
        AcGID = r[3]
        img = returnActionImage(AcGID)
        if  img != None:
            pic = img
        else:
            pic = 'No Picture'


        lst.append([para, pic])
        if score in ['3','4','5']: # and score != bmpInspOverall: 5/14/2018 removed '2', to just have major 
            # score is from the comment
            #10/2018 
            inspcScore='3'
            try:
                aType = r[4]
                inspcScoreField = appConfig.dicInspScoreVSactionName[aType]
                sqlForInspScore = ("Select  left({1},1) from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw]  where globalid = '{0}'").format(inspGID, inspcScoreField)
                inspcScore=ReportSharedFunctions.returnDataODBC(sqlForInspScore)[0][0]
            except:
                pass

            #10/2018 now from issue 216 I need to only add this if the related-field inspection score was not 3,4,5 to avoid dups
            if not inspcScore in ['3','4','5']:
                appConfig.bmpInspActionScores.append([score,[para, pic]])

        pass




def action_types_for_Summary(inspGID, aType, aTyNum): #gk
    sql = ("Select  ACTION_TYPE, ACTION, INSP_ACTION_COMMENT, globalid FROM [deldot_migration2].[dbo].[BMP_INSPECTION_ACTION_evw] " +
         "  where BMP_INSPECTION_ID = '{0}' and  [INSPECTION_PARAMETER] =  '{1}'").format(inspGID, aType)

    result = ReportSharedFunctions.returnDataODBC(sql)

    lst= []
 
    for r in result:
        global seqNum
        seqNum +=1
        r = ['' if x == None else x for x in r]
        sCode = r[0] 
        sAty = ReportSharedFunctions.valFromDomCode("D_Action_Type",sCode).upper()  
        sCode = r[1] 
        sAc = ReportSharedFunctions.valFromDomCode("D_Action",sCode).upper()

        score = '0'
        hasScoreFromComment=False
        sComm = r[2]  
        if len(sComm)> 0:
            ScoreAndcomment = actionScoreAndComment(sComm) #need to do same thing for minors
            score = ScoreAndcomment[0]
            sComm = ScoreAndcomment[1]
            if score != 'N/A':
                hasScoreFromComment=True

        aTyNum = str(seqNum) + ".  "
        para = Paragraph('<para alignment= "LEFT"><font name=Helvetica size=10>' + sAty + '<br/><br/>' + sAc  + '<br/><br/>' + aTyNum  + sComm.upper() +  '</font></para>', styles["Normal"])
        pic = 'No Picture'
        #now try to get the pic
        AcGID = r[3]
        img = returnActionImage(AcGID)
        if  img != None:
            pic = img
        else:
            pic = 'No Picture'

            #10/2018 
        inspcScore='3'
        try:
            inspcScoreField = appConfig.dicInspScoreVSactionName[aType]
            sqlForInspScore = ("Select  left({1},1) from [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw]  where globalid = '{0}'").format(inspGID, inspcScoreField)
            inspcScore=ReportSharedFunctions.returnDataODBC(sqlForInspScore)[0][0]
        except:
            pass

        if (score in ['3','4','5'] and hasScoreFromComment) or (inspcScore in ['3','4','5'] and not hasScoreFromComment) : #10/2018 restrict
            #arcpy.AddMessage(sAty + '<br/><br/>' + sAc  + '<br/><br/>' + aTyNum  + sComm.upper() + ": " + "")		
            lst.append([para, pic])
        else:
            seqNum=seqNum-1

    return lst

def ACTION_ITEM_SUMMARY(inspOBJECTID,tblHeader):
    global seqNum
    seqNum = 0
    try:     
        sql = ("SELECT  bi.globalid, bi.OVERALL_CONDITION_SCORE" +
                " FROM [deldot_migration2].[dbo].[BMP_INSPECTIONS_evw] bi where bi.OBJECTID= " + inspOBJECTID )

        result = ReportSharedFunctions.returnDataODBC(sql)[0]
        Result = ['' if x == None else x for x in result]       


        inspGID =  Result[0]

        Access = action_types_for_Summary(inspGID, 'Access','1')
        Fence = action_types_for_Summary(inspGID, 'Fence','2')
        Invasive = action_types_for_Summary(inspGID, 'Invasive','3')
        PubHazards = action_types_for_Summary(inspGID, 'PubHazards','4')
        PreTreat = action_types_for_Summary(inspGID, 'PreTreat','5')
        InfCond = action_types_for_Summary(inspGID, 'InfCond','6')
        ConveyCond = action_types_for_Summary(inspGID, 'ConveyCond','7')
        DStrmCond = action_types_for_Summary(inspGID, 'DStrmCond','8')
        Ponding = action_types_for_Summary(inspGID, 'Ponding','9')
        WaterQC = action_types_for_Summary(inspGID, 'WaterQC','10')
        WaterQT = action_types_for_Summary(inspGID, 'WaterQT','11')
        OSStab = action_types_for_Summary(inspGID, 'OSStab','12')
        PermanentTool = action_types_for_Summary(inspGID, 'PermanentTool','13')
        BMPVege = action_types_for_Summary(inspGID, 'BMPVege','14')
        UpStrmEMCvr = action_types_for_Summary(inspGID, 'UpStrmEMCvr','15')
        DoStrmEMCvr = action_types_for_Summary(inspGID, 'DoStrmEMCvr','16')
        Seepage = action_types_for_Summary(inspGID, 'Seepage','17')
        RISOpening = action_types_for_Summary(inspGID, 'RISOpening','18')
        RISLowflow = action_types_for_Summary(inspGID, 'RISLowflow','19')
        RISAccumu = action_types_for_Summary(inspGID, 'RISAccumu','20')
        RISStructure = action_types_for_Summary(inspGID, 'RISStructure','21')
        PrclSpill = action_types_for_Summary(inspGID, 'PrclSpill','22')
        Spillway = action_types_for_Summary(inspGID, 'Spillway','23')
        EmgSpillway = action_types_for_Summary(inspGID, 'EmgSpillway','24')
   #revise
        OVERALL_CONDITION_SCORE =  Result[1]
        appConfig.bmpInspOverall = OVERALL_CONDITION_SCORE #gk
        tblName=''
        if OVERALL_CONDITION_SCORE[:1] == 'B':
            appConfig.bmpInspOverall = '3' #gk
            tblName='Maintenance Work Order'
        elif OVERALL_CONDITION_SCORE[:1] == 'C':
            appConfig.bmpInspOverall = '4' #gk
            tblName='Contracted Work'
        elif OVERALL_CONDITION_SCORE[:1] == 'D':
            appConfig.bmpInspOverall = '5' #gk
            tblName='Retrofit'

        if OVERALL_CONDITION_SCORE[:1] == 'A':
            return None


        content=[[tblHeader,'']]
        content.append(['ACTION ITEM SUMMARY',''])
        content.append([tblName,''])


        for l in Access: content.append(l)
        for l in Fence: content.append(l)
        for l in Invasive: content.append(l)
        for l in PubHazards: content.append(l)
        for l in PreTreat: content.append(l)
        for l in InfCond: content.append(l)
        for l in ConveyCond: content.append(l)
        for l in DStrmCond: content.append(l)
        for l in Ponding: content.append(l)
        for l in WaterQC: content.append(l)
        for l in WaterQT: content.append(l)
        for l in OSStab: content.append(l)
        for l in PermanentTool: content.append(l)
        for l in BMPVege: content.append(l)
        for l in UpStrmEMCvr: content.append(l)
        for l in DoStrmEMCvr: content.append(l)
        for l in Seepage: content.append(l)
        for l in RISOpening: content.append(l)
        for l in RISLowflow: content.append(l)
        for l in RISAccumu: content.append(l)
        for l in RISStructure: content.append(l)
        for l in PrclSpill: content.append(l)
        for l in Spillway: content.append(l)
        for l in EmgSpillway: content.append(l)


        reportGrid = Table(content, [3.41*inch,3.4*inch], repeatRows=3)
        reportGrid.setStyle(TableStyle([('ALIGN', (0,0), (-1, -1), 'LEFT'), ('ALIGN', (0,2), (1, 2), 'CENTER'), ('VALIGN', (0,0), (-1, -1), 'MIDDLE'),
                                     ('SPAN', (0,1), (-1, 1)),   ('SPAN', (0,2), (-1, 2)), ('BACKGROUND',(0,2),(1,2),colors.lightgrey),
                                        ('SIZE', (0,0), (-1, -1), 10), 
                                        ('SIZE', (0,1), (1, 1), 11),
                                        ('SIZE', (0,2), (-1, 2), 14) ,
                                        ('FONTNAME', (0,1), (1,2), 'Helvetica-Bold') ,
                                         ('BOTTOMPADDING', (0,1), (1, 1), 8),('BOTTOMPADDING', (0,2), (1, 2), 10),
                                               ('BOTTOMPADDING', (1,3), (1, -1), 0),
                                                ('TOPPADDING', (1,3), (1, -1), 0),
                                                ('LEFTPADDING', (1,3), (1, -1), 0),
                                                ('RIGHTPADDING', (1,3), (1, -1), 0),
                                        ('BOX', (0,2), (-1,-1),1, colors.black) ,  ('INNERGRID', (0,2), (-1,-1), 0.25, colors.black)
                                    ]))

        return reportGrid


    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
		
def ACTION_newTables(tblHeader,score):
    try:     

        if score == '3':
            tblName='Maintenance Work Order'
        elif score == '4':
            tblName='Contracted Work'
        elif score == '5':
            tblName='Retrofit'
        

        content=[[tblHeader,'']]
        content.append(['ACTION ITEM SUMMARY' + " --- from individual action scores",''])
        content.append([tblName,''])


        thisList = []

        for l in appConfig.bmpInspActionScores: 
            #if l[0] == score:
            thisList.append(l[1])

        for l in thisList:
            content.append(l)
                

        reportGrid = Table(content, [3.41*inch,3.4*inch], repeatRows=3)
        reportGrid.setStyle(TableStyle([('ALIGN', (0,0), (-1, -1), 'LEFT'), ('ALIGN', (0,2), (1, 2), 'CENTER'), ('VALIGN', (0,0), (-1, -1), 'MIDDLE'),
                                     ('SPAN', (0,1), (-1, 1)),   ('SPAN', (0,2), (-1, 2)), ('BACKGROUND',(0,2),(1,2),colors.lightgrey),
                                        ('SIZE', (0,0), (-1, -1), 10), 
                                        ('SIZE', (0,1), (1, 1), 11),
                                        ('SIZE', (0,2), (-1, 2), 14) ,
                                        ('FONTNAME', (0,1), (1,2), 'Helvetica-Bold') ,
                                         ('BOTTOMPADDING', (0,1), (1, 1), 8),('BOTTOMPADDING', (0,2), (1, 2), 10),
                                               ('BOTTOMPADDING', (1,3), (1, -1), 0),
                                                ('TOPPADDING', (1,3), (1, -1), 0),
                                                ('LEFTPADDING', (1,3), (1, -1), 0),
                                                ('RIGHTPADDING', (1,3), (1, -1), 0),
                                        ('BOX', (0,2), (-1,-1),1, colors.black) ,  ('INNERGRID', (0,2), (-1,-1), 0.25, colors.black)
                                    ]))

        return reportGrid


    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise