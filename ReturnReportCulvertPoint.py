from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, inch
from reportlab.platypus import *
from reportlab.lib.styles import getSampleStyleSheet
 
import  os, sys, traceback 
import datetime, time
import ReportSharedFunctions


styles = getSampleStyleSheet()
HeaderStyle = styles["Heading2"]
ParaStyle = styles["Normal"]
PreStyle = styles["Code"]
styleN = styles['Normal']





def returnInventory_and_Inspection_Information(OBJECTID):
    try:

        sql = ("select (case when insp.INSPECTION_DATE is null then '' else FORMAT(insp.INSPECTION_DATE, 'M/d/yyyy' ) end) as d , " +
                "insp.INSPECTED_REASON, insp.INSPECTED_REASON_DESCRIPTION, insp.OVERALL_RATING "+ 
                "from deldot_migration2.dbo.STRUCTURES_evw  s  "+
                "join deldot_migration2.dbo.CULVERT_POINTS_evw cs on cs.STRUCTURE_ID=s.GlobalID " + 
                "left join deldot_migration2.dbo.CULVERT_PT_INSPECTIONS_evw insp on insp.CULVERT_POINT_ID =cs.GlobalID  where s.OBJECTID = " + OBJECTID + " order by insp.INSPECTION_DATE desc ")     
 
        result = ReportSharedFunctions.returnDataODBC(sql)[0]
        result = ['' if x == None else x for x in result] 


        Last_Inspected_Date = result[0]    

        Not_Inspected_Reason = result[1] 
        if Not_Inspected_Reason != '': Not_Inspected_Reason = ReportSharedFunctions.valFromDomCode("D_ReasonStructure", Not_Inspected_Reason)

        Reason_Description = result[2]

        Overall_Rating = result[3]
        Overall_Rating = ReportSharedFunctions.valFromDomCode("D_Condition",Overall_Rating)

        content = ([['Last Inspected Date:',Last_Inspected_Date]])
        content.append(['Not Inspected Reason:',Not_Inspected_Reason])
        content.append(['Reason Description:',Paragraph('<para alignment= "LEFT"><font size=11>' + Reason_Description+'</font></para>', styles["Normal"])])
        content.append(['Overall Rating:',Overall_Rating])

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
        "FROM [deldot_migration2].[dbo].[WORK_ORDERS_evw] w join [deldot_migration2].[dbo].[STRUCTURES_evw] s on s.GlobalID = w.[FEATURE_ID] " + 
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
                contactContent = [[Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11><u>Contact ' + str(icontacts) + ':</u></font>', styles["Normal"]),'']]
            else:
                contactContent.append([Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11><u>Contact ' + str(icontacts) + ':</u></font>', styles["Normal"]),''])
                                
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