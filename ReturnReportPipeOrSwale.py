from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, inch
from reportlab.platypus import *
from reportlab.lib.styles import getSampleStyleSheet
 
import  os, sys, traceback ,arcpy
import datetime, time
import appConfig, ReportSharedFunctions
import pypyodbc

styles = getSampleStyleSheet()
HeaderStyle = styles["Heading2"]
ParaStyle = styles["Normal"]
PreStyle = styles["Code"]
styleN = styles['Normal']


def returnSwaleInventory_and_Inspection_Information(OBJECTID):
    try:

        sql = ("select top 1 s.MATERIAL, s.SWALE_SHAPE, s.SIDE_SLOPE, s.TOP_WIDTH, s.BOTTOM_WIDTH, s.DEPTH,  " +
            "(case when  SI.INSPECTION_DATE is null then '' else FORMAT( SI.INSPECTION_DATE, 'M/d/yyyy' ) end) as d,  " +
            "SI.INSPECTED_REASON, SI.INSPECTED_REASON_DESCRIPTION, SI.OVERALL_RATING, SI.DEBRIS, SI.EROSION, SI.swale_insp_comments  " +
            "from [deldot_migration2].[dbo].CONVEYANCES_evw c " +
            "left join [deldot_migration2].[dbo].swales_evw s on s.CONVEYANCE_ID=c.GlobalID  " +
            "left join [deldot_migration2].[dbo].SWALE_INSPECTIONS_evw si on si.SWALE_ID = s.GlobalID " +

                "where c.OBJECTID = {0} order by SI.INSPECTION_DATE desc " ).format(OBJECTID)    
 
        result = ReportSharedFunctions.returnDataODBC(sql)[0]
        result = ['' if x == None else x for x in result] 

        Material = result[0]
        Swale_Shape = result[1]
        Side_Slope = result[2]
        Top_Width_in = result[3]
        Bottom_Width_in = result[4]
        Depth_in = result[5]
        Last_Inspected_Date = result[6]
        Not_Inspected_Reason = result[7]
        Reason_Description = result[8]
        Overall_Rating = result[9]
        Debris_pct = result[10]
        Erosion = result[11]
        Inspection_Comments = result[12]
        
        Material = ReportSharedFunctions.valFromDomCode("D_DitchMaterial",Material)  
        Swale_Shape = ReportSharedFunctions.valFromDomCode("D_DitchShape",Swale_Shape) 
        Side_Slope = ReportSharedFunctions.valFromDomCode("D_Slope",Side_Slope) 
        Inspection_Comments = Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11>' + Inspection_Comments + '</font></para>', styles["Normal"])
        Not_Inspected_Reason = ReportSharedFunctions.valFromDomCode("D_ReasonConveyance", Not_Inspected_Reason)
        Reason_Description = Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11>' + Reason_Description + '</font></para>', styles["Normal"])
        Debris_pct = ReportSharedFunctions.valFromDomCode("D_Percent",Debris_pct)
        Erosion = ReportSharedFunctions.valFromDomCode("D_Erosion",Erosion)
        Overall_Rating = ReportSharedFunctions.valFromDomCode("D_Condition",Overall_Rating)




        if Top_Width_in != '': 
            if float(Top_Width_in).is_integer() :
                Top_Width_in = str(int(Top_Width_in))
            else:
                Top_Width_in = str(round(Top_Width_in,2))

        if Bottom_Width_in != '': 
            if round(float(Bottom_Width_in),0).is_integer() :
                Bottom_Width_in = str(int(Bottom_Width_in))
            else:
                Bottom_Width_in = str(round(Bottom_Width_in,2))

        if Depth_in != '': 
            if round(float(Depth_in),0).is_integer() :
                Depth_in = str(int(Depth_in))
            else:
                Depth_in = str(round(Depth_in,2))

        content = [['Material:',Material]]
        content.append(['Swale Shape:',Swale_Shape])
        content.append(['Side Slope:',Side_Slope])
        content.append(['Top Width (in):',Top_Width_in])
        content.append(['Bottom Width (in):',Bottom_Width_in])
        content.append(['Depth (in):',Depth_in])
        content.append(['Last Inspected Date:',Last_Inspected_Date])
        content.append(['Not Inspected Reason:',Not_Inspected_Reason])
        content.append(['Reason Description:',Reason_Description])
        content.append(['Overall Rating:',Overall_Rating])
        content.append(['Debris (%):',Debris_pct])
        content.append(['Erosion:',Erosion])
        content.append(['Inspection Comments:',Inspection_Comments])

        reportGrid = Table(content, [2.5*inch, 5*inch])
        reportGrid.setStyle(TableStyle([('ALIGN', (0,0), (0, -1), 'RIGHT'), ('ALIGN', (1,0), (1, -1), 'LEFT'),   ('VALIGN', (0,0), (0, -1), 'TOP'),
                                       ('SIZE', (0,0), (-1, -1), 11),  
                                       ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold') ,
                                       ('BOX', (0,0), (-1,-1),1, colors.black) ,  ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black)
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
         


def returnPipeInventory_and_Inspection_Information(OBJECTID):
    try:

        sql = ("select top 1 P.CONVEYS_STREAM, P.GNIS_ID, P.CONV_STREAM_COMMENTS, P.MATERIAL, P.PIPE_SHAPE, P.WIDTH, P.HEIGHT, P.CROSS_SECTIONAL_AREA,  " +
                "(case when insp.INSPECTION_DATE is null then '' else FORMAT(insp.INSPECTION_DATE, 'M/d/yyyy' ) end) as d,  " +
                "INSP.INSPECTED_REASON, INSP.INSPECTED_REASON_DESCRIPTION, PIR.OVERALL_RATING, PIR.DEBRIS, " +
                "PIR.UPSTREAM_SEAL, PIR.DOWNSTREAM_SEAL, PIR.BARREL_INSP_COMMENTS, c.SHAPE.STLength() " +
                "from [deldot_migration2].dbo.CONVEYANCES_evw c " +
                "left join [deldot_migration2].dbo.PIPE_SEGMENTS_evw P on P.CONVEYANCE_ID = c.GlobalID   " +
                "left join [deldot_migration2].dbo.PIPE_INSPECTIONS_evw INSP on INSP.PIPE_FEATURE_ID= P.GlobalID " +
                "left join [deldot_migration2].dbo.PIPE_INSPECTION_RATINGS_evw pir on pir.PIPE_INSPECTION_ID = insp.GlobalID " +
                "where c.OBJECTID = {0} order by insp.INSPECTION_DATE desc " ).format(OBJECTID)    
 
        result = ReportSharedFunctions.returnDataODBC(sql)[0]
        result = ['' if x == None else x for x in result] 

        Conveys_Stream = result[0]
        GNIS_ID = result[1]
        GNIS_Name = result[2]

 
        if GNIS_ID != '':
            flowLineConnection = pypyodbc.connect(appConfig.flowLineConnStr)
            sql = "SELECT top 1 [GNIS_NAME] FROM [DelDOTBase].[dbo].[MAJOR_RIVERS_AND_STREAMS] where [GNIS_ID] = '" + GNIS_ID + "'"
            cursor = appConfig.connection.cursor()  
            cursor.execute(sql) 
            row = cursor.fetchone()
            if row:
                GNIS_Name = row[0]

            flowLineConnection.close()

        Material = result[3]
        Shape = result[4]
        Width_in = result[5]
        Height_in = result[6]
        Cross_Sectional_Area_sq_ft = result[7]
        Last_Inspected_Date = result[8]
        Not_Inspected_Reason = result[9]
        Reason_Description = result[10]
        Overall_Rating = result[11]
        Debris_pct = result[12]
        Upstream_Seal_Rating = result[13]
        Downstream_Seal_Rating = result[14]
        Inspection_Comments = result[15]
        pipeLen = str(int(round(result[16]))) 

        Material = ReportSharedFunctions.valFromDomCode("D_PipeMaterial",Material)
        Shape = ReportSharedFunctions.valFromDomCode("D_PipeShape",Shape)
        Not_Inspected_Reason = ReportSharedFunctions.valFromDomCode("D_ReasonConveyance", Not_Inspected_Reason)
        Reason_Description = Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11>' + Reason_Description + '</font></para>', styles["Normal"])

        Debris_pct = ReportSharedFunctions.valFromDomCode("D_Percent",Debris_pct)
        Overall_Rating = ReportSharedFunctions.valFromDomCode("D_Condition",Overall_Rating)
        Upstream_Seal_Rating = ReportSharedFunctions.valFromDomCode("D_Condition",Upstream_Seal_Rating)
        Downstream_Seal_Rating = ReportSharedFunctions.valFromDomCode("D_Condition",Downstream_Seal_Rating)
        Inspection_Comments = Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11>' + Inspection_Comments + '</font></para>', styles["Normal"])


        if Width_in != '': 
            if float(Width_in).is_integer() :
                Width_in = str(int(Width_in))
            else:
                Width_in = str(round(Width_in,2))

        if Height_in != '': 
            if round(float(Height_in),0).is_integer() :
                Height_in = str(int(Height_in))
            else:
                Height_in = str(round(Height_in,2))

        if Cross_Sectional_Area_sq_ft != '': 
            Cross_Sectional_Area_sq_ft = str(round(Cross_Sectional_Area_sq_ft,2))


        content = [['Conveys Stream:',Conveys_Stream]]
        content.append(['GNIS ID:',GNIS_ID])
        content.append(['GNIS Name:',GNIS_Name])
        content.append(['Material:',Material])
        content.append(['Shape:',Shape])
        content.append(['Length (ft):',pipeLen])
        content.append(['Width (in):',Width_in])
        content.append(['Height (in):',Height_in])
        content.append(['Cross Sectional Area (sq. ft.):',Cross_Sectional_Area_sq_ft])
        content.append(['Last Inspected Date:',Last_Inspected_Date])
        content.append(['Not Inspected Reason:',Not_Inspected_Reason])
        content.append(['Reason Description:',Reason_Description])
        content.append(['Overall Rating:',Overall_Rating])
        content.append(['Debris (%):',Debris_pct])
        content.append(['Upstream Seal Rating:',Upstream_Seal_Rating])
        content.append(['Downstream Seal Rating:',Downstream_Seal_Rating])
        content.append(['Inspection Comments:',Inspection_Comments])





        reportGrid = Table(content, [2.5*inch, 5*inch])
        reportGrid.setStyle(TableStyle([('ALIGN', (0,0), (0, -1), 'RIGHT'), ('ALIGN', (1,0), (1, -1), 'LEFT'),   ('VALIGN', (0,0), (0, -1), 'TOP'),
                                       ('SIZE', (0,0), (-1, -1), 11),  
                                       ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold') ,
                                       ('BOX', (0,0), (-1,-1),1, colors.black) ,  ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black)
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
def returnWorkOrder_Information(OBJECTID):
    try:
        sql = ("SELECT top 1  w.[STATUS],w.[DISTRICT],w.[SUPERVISOR], w.[IMMEDIATE_ACTION] ,w.[WORK_ORDER_COMMENTS], w.GlobalID " + 
        "FROM [deldot_migration2].[dbo].[WORK_ORDERS_evw] w join [deldot_migration2].[dbo].[conveyances_evw] s on s.GlobalID = w.[FEATURE_ID] " + 
        "where w.[STATUS] <> 'Cancelled' and s.OBJECTID=" + OBJECTID + " order by w.LAST_EDIT_DATE desc")     

        result = ReportSharedFunctions.returnDataODBC(sql)[0]
        result = ['' if x == None else x for x in result]   

        Status = result[0]
        District = result[1]
        Supervisor = result[2]
        Immediate_Action = ReportSharedFunctions.boolVal(result[3])
        Long_Description =  Paragraph('<para alignment= "LEFT"><font size=11>' + result[4] +'</font></para>', styles["Normal"])

        WOglobID =  result[5]  

        content = [['Status:',ReportSharedFunctions.valFromDomCode("D_WorkOrderStatus",Status)]]
        content.append(['District:',ReportSharedFunctions.valFromDomCode("D_District",District)])
        content.append(['Supervisor:',Supervisor])
        content.append(['Immediate Action:',Immediate_Action])   
        content.append(['Long Description:',Long_Description])



#now need to add all defects, contacts and edit tableStyle
        sql = ("SELECT  [PRIORITY] ,[FAILURE_CLASS] ,[PROBLEM_CODE]  ,FUNCTION_CODE  ,[GlobalID]FROM [deldot_migration2].[dbo].[WORK_ORDER_PROBLEM_CODES_evw] " + 
         "where [WORK_ORDER_ID] = '" + WOglobID +"'") 

        hasDefects = False
        iDefects=1
        defectContentS = [[]]
        defectResults = ReportSharedFunctions.returnDataODBC(sql)
        for result in defectResults:
            hasDefects = True
            result = ['' if x == None else x for x in result] 
            Priority = result[0]
            Failure_Class = result[1]
            Problem_Code = result[2]
            Function_Code = result[3]
            DefectGlobID = result[4]
            #get the photos into a list here??
            if iDefects==1:
                defectContent = [[Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11><u>Defect ' + str(iDefects) + ':</u></font></para>', styles["Normal"]),'']]
            else:
                defectContent.append([Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11><u>Defect ' + str(iDefects) + ':</u></font></para>', styles["Normal"]),''])
                                
            defectContent.append(['Priority:',Priority])
            defectContent.append(['Failure Class:', ReportSharedFunctions.valFromDomCode("D_FailureClass",Failure_Class)])
            defectContent.append(['Problem Code:',ReportSharedFunctions.valFromDomCode("D_ProblemCode",Problem_Code)])
            defectContent.append(['Function Code:',ReportSharedFunctions.valFromDomCode("D_FunctionCode",Function_Code) ])
            iDefects+=1


            defectGrid = Table(defectContent, [2.5*inch, 5*inch])
            defectGrid.setStyle(TableStyle([('ALIGN', (0,0), (0, -1), 'RIGHT'), ('ALIGN', (1,0), (1, -1), 'LEFT'),
                                               ('VALIGN', (0,0), (0, -1), 'TOP'),
                                           ('SIZE', (0,0), (-1, -1), 11),  
                                           ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold') ,
                                            ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black)
                                        ]))



        
     #contacts   
#############################################################################################################################################
        sql = ("SELECT [PERSON_CONTACTED] , (case when DATE_CONTACTED is null then '' else FORMAT(DATE_CONTACTED, 'M/d/yyyy' ) end) ,[CONTACT_PERSON_DISTRICT]  ,[CONTACT_METHOD] FROM [deldot_migration2].[dbo].[WORK_ORDER_CONTACTS_evw] " + 
         "where [WORK_ORDER_ID] = '" + WOglobID +"'") 

        hascontacts = False
        icontacts=1
        contactContentS = [[]]
        contactResults = ReportSharedFunctions.returnDataODBC(sql)
        for result in contactResults:
            hascontacts = True
            result = ['' if x == None else x for x in result] 
            Person_Contacted = result[0]
            Contacted_Date = result[1]
            District = result[2]
            Contact_Method = result[3]

            if icontacts==1:
                contactContent = [[Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11><u>Contact ' + str(icontacts) + ':</u></font></para>', styles["Normal"]),'']]
            else:
                contactContent.append([Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11><u>Contact ' + str(icontacts) + ':</u></font></para>', styles["Normal"]),''])
                                
            contactContent.append(['Person Contacted:',Person_Contacted])
            contactContent.append(['Contacted Date:',Contacted_Date])
            contactContent.append(['District:',ReportSharedFunctions.valFromDomCode("D_District",District)])
            contactContent.append(['Contact Method:',ReportSharedFunctions.valFromDomCode("D_ContactMethod",Contact_Method)])
            icontacts+=1


            contactGrid = Table(contactContent, [2.5*inch, 5*inch])
            contactGrid.setStyle(TableStyle([('ALIGN', (0,0), (0, -1), 'RIGHT'), ('ALIGN', (1,0), (1, -1), 'LEFT'),
                                           ('VALIGN', (0,0), (0, -1), 'TOP'),
                                       ('SIZE', (0,0), (-1, -1), 11),  
                                       ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold') ,
                                        ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black)
                                    ]))


###########################################################################################################        
        vTableStyle = [('ALIGN', (0,0), (0, -1), 'RIGHT'), ('ALIGN', (1,0), (1, -1), 'LEFT'),
                                    ('VALIGN', (0,0), (0, -1), 'TOP'),
                                    ('SIZE', (0,0), (-1, -1), 11),  
                                    ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold') ,
                                    ('BOX', (0,0), (-1,-1),1, colors.black) ,  ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black)
                                    ]

#######################################################################
        
        if hasDefects:
            content.append([defectGrid])
            vTableStyle.append(('SPAN',(0,5),(-1,5)))
            vTableStyle.append(('LEFTPADDING',(0,5),(-1,5),0))
            vTableStyle.append(('RIGHTPADDING',(0,5),(-1,5),0))
            vTableStyle.append(('TOPPADDING',(0,5),(-1,5),0))
            vTableStyle.append(('BOTTOMPADDING',(0,5),(-1,5),0))

        if hascontacts:
            r= 6 if hasDefects else 5
            content.append([contactGrid])
            vTableStyle.append(('SPAN',(0,r),(-1,r)))
            vTableStyle.append(('LEFTPADDING',(0,r),(-1,r),0))
            vTableStyle.append(('RIGHTPADDING',(0,r),(-1,r),0))
            vTableStyle.append(('TOPPADDING',(0,r),(-1,r),0))
            vTableStyle.append(('BOTTOMPADDING',(0,r),(-1,r),0))

        reportGrid = Table(content, [2.5*inch, 5*inch])
        reportGrid.setStyle(TableStyle(vTableStyle))
                    
        return  reportGrid

    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass