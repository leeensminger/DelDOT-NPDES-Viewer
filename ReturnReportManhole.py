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

        sql = ("SELECT MH.MANHOLE_TYPE,MH.TIDE_GATE_PRESENT,MH.MODIFIED,MH.STORAGE,MH.STORAGE_DEPTH, " + 
                "(case when MHI.INSPECTION_DATE is null then '' else FORMAT(MHI.INSPECTION_DATE, 'M/d/yyyy' ) end) as d , " +
                "MHI.INSPECTED_REASON,MHI.INSPECTED_REASON_DESCRIPTION,MHI.OVERALL_RATING,MHI.INFILTRATION,MHI.DEBRIS   " +
                "FROM [deldot_migration2].[dbo].[STRUCTURES_evw] s join  [deldot_migration2].[dbo].MANHOLES_evw MH on MH.STRUCTURE_ID=s.GlobalID " +
                "left join [deldot_migration2].[dbo].MANHOLE_INSPECTIONS_evw MHI on MHI.MANHOLE_ID=MH.GlobalID " +
                "where s.OBJECTID = " + OBJECTID + " order by MHI.INSPECTION_DATE desc ")     
 
        result = ReportSharedFunctions.returnDataODBC(sql)[0]
        result = ['' if x == None else x for x in result] 

        Manhole_Type = result[0]
        Manhole_Type = ReportSharedFunctions.valFromDomCode("D_MH_Desc", Manhole_Type)

        Tide_Gate_Present = result[1]
        if Tide_Gate_Present != "": Tide_Gate_Present = ReportSharedFunctions.boolVal(Tide_Gate_Present)

        Modified = result[2]
        if Modified != "": Modified = ReportSharedFunctions.boolVal(Modified)

        Storage = result[3]
        if Storage != "": Storage = ReportSharedFunctions.boolVal(Storage)

        Storage_Depth = result[4]
        if Storage_Depth != '':
            Storage_Depth = "{0:.1f}".format(Storage_Depth)
            if Storage_Depth == '0.0': Storage_Depth = '0'
			
        Last_Inspected_Date = result[5]   

        Not_Inspected_Reason = result[6] 
        if Not_Inspected_Reason != '': Not_Inspected_Reason = ReportSharedFunctions.valFromDomCode("D_ReasonStructure", Not_Inspected_Reason)


        Reason_Description = result[7]
        Overall_Rating = result[8]
        Overall_Rating = ReportSharedFunctions.valFromDomCode("D_Condition",Overall_Rating)

        Infiltration = result[9]
        Infiltration = ReportSharedFunctions.valFromDomCode("D_Infiltration",Infiltration)

        Debris_pct = result[10]
        Debris_pct = ReportSharedFunctions.valFromDomCode("D_Percent",Debris_pct)
        content = [['Manhole Type:',Manhole_Type]]
        content.append(['Tide Gate Present:',Tide_Gate_Present])
        content.append(['Modified:',Modified])
        content.append(['Storage:',Storage])
        content.append(['Storage Depth (in):',Storage_Depth])
        content.append(['Last Inspected Date:',Last_Inspected_Date])
        content.append(['Not Inspected Reason:',Not_Inspected_Reason])
        content.append(['Reason Description:',Paragraph('<para alignment= "LEFT"><font size=11>' + Reason_Description+'</font></para>', styles["Normal"])])
        content.append(['Overall Rating:',Overall_Rating])
        content.append(['Infiltration:',Infiltration])
        content.append(['Debris (%):',Debris_pct])

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
         

    
def returnComponent_Information(OBJECTID):
    try: 
        sql = ("SELECT top 1 i.FRAME_OF_COVER_MATERIAL,i.FRAME_OF_COVER_LENGTH,i.FRAME_OF_COVER_WIDTH,i.FRAME_OF_COVER_HEIGHT,insp.FRAME_OF_COVER_CONDITION, " +   
                "i.COVER_MATERIAL,i.COVER_LENGTH,i.COVER_WIDTH,i.COVER_HEIGHT,insp.COVER_CONDITION,i.WALL_MATERIAL_LU,i.WALL_LENGTH,i.WALL_WIDTH,i.WALL_HEIGHT,   " +  
                "insp.WALL_CONDITION,i.ADJUST_RING_MATERIAL,i.ADJUST_RING_COUNT,i.ADJUST_RING_HEIGHT,i.STEP_MATERIAL,i.STEP_COUNT,i.LOW_FLOW_MATERIAL,  " +  
                "s.[OBJECTID],s.[GlobalID]FROM [deldot_migration2].[dbo].[STRUCTURES_evw] s  " +  
                "left join  [deldot_migration2].[dbo].[MANHOLES_evw] i on i.STRUCTURE_ID = s.[GlobalID]   " +  
                "left join [deldot_migration2].[dbo].[MANHOLE_INSPECTIONS_evw] insp on insp.MANHOLE_ID = i.GlobalID " +  
                "where s.[OBJECTID] =   " + OBJECTID + "   order by insp.INSPECTION_DATE desc;" )     


        result = ReportSharedFunctions.returnDataODBC(sql)[0]
        result = ['' if x == None else x for x in result] 
        result = ['{:.0f}'.format(x) if (str(type(x))).find("Decimal") > -1 else x for x in result]   

        focMaterial = result[0]; focMaterial = ReportSharedFunctions.valFromDomCode("D_ComponentMaterial",focMaterial)
        focLength = result[1]
        focWidth = result[2]
        focHeight = result[3]
        focRating = result[4]; focRating = ReportSharedFunctions.valFromDomCode("D_Condition",focRating)

        cvMaterial = result[5]; cvMaterial = ReportSharedFunctions.valFromDomCode("D_ComponentMaterial",cvMaterial)
        cvLength = result[6]
        cvWidth = result[7]
        cvHeight = result[8]
        cvRating = result[9]; cvRating = ReportSharedFunctions.valFromDomCode("D_Condition",cvRating)

        wMaterial = result[10]; wMaterial = ReportSharedFunctions.valFromDomCode("D_ComponentMaterial",wMaterial)
        wLength = result[11]
        wWidth = result[12]
        wHeight = result[13]
        wRating = result[14]; wRating = ReportSharedFunctions.valFromDomCode("D_Condition",wRating)

        arMaterial = result[15]; arMaterial = ReportSharedFunctions.valFromDomCode("D_ComponentMaterial",arMaterial)
        arCount = result[16]
        arHeight = result[17]

        spMaterial = result[18]; spMaterial = ReportSharedFunctions.valFromDomCode("D_ComponentMaterial",spMaterial)
        spCount = result[19]

        lfMaterial = result[20]; lfMaterial = ReportSharedFunctions.valFromDomCode("D_ComponentMaterial",lfMaterial)

        content = [[Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11><u>Frame of Cover:</u></font></para>', styles["Normal"]),'']]
        content.append(['Material:',focMaterial])
        content.append(['Length (in):',focLength])
        content.append(['Width (in):',focWidth])
        content.append(['Height (in):',focHeight])
        content.append(['Rating:',focRating])
        content.append([Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11><u>Cover:</u></font></para>', styles["Normal"]),''])
        content.append(['Material:',cvMaterial])
        content.append(['Length (in):',cvLength])
        content.append(['Width (in):',cvWidth])
        content.append(['Height (in):',cvHeight])
        content.append(['Rating:',cvRating])
        content.append([Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11><u>Wall:</u></font></para>', styles["Normal"]),''])
        content.append(['Material:',wMaterial])
        content.append(['Length (in):',wLength])
        content.append(['Width (in):',wWidth])
        content.append(['Height (in):',wHeight])
        content.append(['Rating:',wRating])
        content.append([Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11><u>Adjustment Ring:</u></font></para>', styles["Normal"]),''])
        content.append(['Material:',arMaterial])
        content.append(['Count:',arCount])
        content.append(['Height (in):',arHeight])
        content.append([Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11><u>Step:</u></font></para>', styles["Normal"]),''])
        content.append(['Material:',spMaterial])
        content.append(['Count:',spCount])
        content.append([Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11><u>Low Flow Channel:</u></font></para>', styles["Normal"]),''])
        content.append(['Material:',lfMaterial])

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
         

def returnOutfall_Information(OBJECTID):
    try:
        sql = ("SELECT top 1 ofs.NPDES,ofs.STREAM_DISTANCE,ofs.GNIS_NAME,ofs.GNIS_ID,oda.DRAINAGE_AREA_AC,insp.OVERALL_RATING,ofs.OF_PROTECTION_MATERIAL,"+
               "insp.OF_PROTECTION_CONDITION,ofs.DEFINED_OF_CHANNEL_MATERIAL,insp.DOC_CHANNEL_CONDITION,insp.DOC_CHANNEL_BED_ERO,insp.DOC_CHANNEL_BANK_ERO"+
               ",ofs.UNDEFINED_OF_CHANNEL_MATERIAL,insp.UDOC_CHANNEL_CONDITION,ofs. OF_END_TREATMENT_TYPE,ofs.OF_END_TREATMENT_MATERIAL,"+
               "insp.OF_END_TREATMENT_CONDITION,s.[OBJECTID],s.[GlobalID] FROM [deldot_migration2].[dbo].[STRUCTURES_evw] s " + 
               "join [deldot_migration2].[dbo].OUTFALLS_evw  ofs on ofs.STRUCTURE_ID = s.[GlobalID] " + 
               "join  [deldot_migration2].[dbo].[manholes_evw] pe on pe.STRUCTURE_ID = s.[GlobalID] " + 
               " left join  [deldot_migration2].[dbo].[manhole_inspections_evw] pei on pei.manhole_ID = pe.[GlobalID] "+ 
               "left join  [deldot_migration2].[dbo].[OUTFALL_INSPECTIONS_evw] insp on insp.STRUCT_INSP_ID = pei.GlobalID " + 
               "left join  [deldot_migration2].[dbo].OUTFALL_DRAINAGE_AREAS_evw oda on oda.OUTFALL_ID = ofs.GlobalID " + 
               "where s.OBJECTID = " + OBJECTID + " order by pei.INSPECTION_DATE desc")

        result = ReportSharedFunctions.returnDataODBC(sql)[0]
        result = ['' if x == None else x for x in result] 
        result = ['{:.0f}'.format(x) if (str(type(x))).find("Decimal") > -1 else x for x in result]   

        NPDES = ReportSharedFunctions.boolVal(result[0])
        aStream_Distance_ft = result[1]
        aGNIS_Name = result[2]
        aGNIS_ID = result[3]
        aDrainage_Area = result[4]
        aOverall_Rating = result[5]
        aOverall_Rating = ReportSharedFunctions.valFromDomCode("D_Condition",aOverall_Rating)
        bMaterial = ReportSharedFunctions.valFromDomCode("D_OutfallMaterial",result[6]) 
        bCondition = ReportSharedFunctions.valFromDomCode("D_Condition",result[7]) 

        cMaterial =  ReportSharedFunctions.valFromDomCode("D_OutfallMaterial",result[8])  
        cChannel__Condition = ReportSharedFunctions.valFromDomCode("D_Condition",result[9]) 
        cChannel_Bed_Erosion =ReportSharedFunctions.valFromDomCode("D_Erosion",result[10]) 
        cChannel_Bank_Erosion = ReportSharedFunctions.valFromDomCode("D_Erosion",result[11]) 

        dMaterial = ReportSharedFunctions.valFromDomCode("D_OutfallMaterial",result[12]) 
        dCondition =ReportSharedFunctions.valFromDomCode("D_Condition",result[13]) 

        eType = ReportSharedFunctions.valFromDomCode("D_EndTreatmentType",result[14])  
        eMaterial = ReportSharedFunctions.valFromDomCode("D_OutfallMaterial",result[15]) 
        eCondition = ReportSharedFunctions.valFromDomCode("D_Condition",result[16]) 


        content = [['NPDES:',NPDES]]
        content.append(['Stream Distance (ft):',aStream_Distance_ft])
        content.append(['GNIS Name:',aGNIS_Name])
        content.append(['GNIS ID:',aGNIS_ID])
        content.append(['Drainage Area:',aDrainage_Area])
        content.append(['Overall Rating:',aOverall_Rating])
        content.append([Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11><u>Outfall Protection:</u></font></para>', styles["Normal"]),''])
        content.append(['Material:',bMaterial])
        content.append(['Condition:',bCondition])
        content.append([Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11><u>Defined Outfall Channel:</u></font></para>', styles["Normal"]),''])
        content.append(['Material:',cMaterial])
        content.append(['Channel  Condition:',cChannel__Condition])
        content.append(['Channel Bed Erosion:',cChannel_Bed_Erosion])
        content.append(['Channel Bank Erosion:',cChannel_Bank_Erosion])
        content.append([Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11><u>Undefined Outfall Channel:</u></font></para>', styles["Normal"]),''])
        content.append(['Material:',dMaterial])
        content.append(['Condition:',dCondition])
        content.append([Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11><u>Outfall End Treatment:</u></font>', styles["Normal"]),''])
        content.append(['Type:',eType])
        content.append(['Material:',eMaterial])
        content.append(['Condition:',eCondition])


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



def returnOutfall_InformationOLD(OBJECTID):
    try:
        sql = ("")     

        #result = ReportSharedFunctions.returnDataODBC(sql)[0]
        #result = ['' if x == None else x for x in result] 
        #result = ['{:.0f}'.format(x) if (str(type(x))).find("Decimal") > -1 else x for x in result]   

        content = [['Outfall Information:',"TEST"]]
        content = [['NPDES:',"TEST"]]
        content.append(['Stream Distance (ft):',"TEST"])
        content.append(['GNIS Name:',"TEST"])
        content.append(['GNIS ID:',"TEST"])
        content.append(['Drainage Area:',"TEST"])
        content.append(['Overall Rating:',"TEST"])
        content.append([Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11><u>Outfall Protection:</u></font></para>', styles["Normal"]),''])
        content.append(['Material:',"TEST"])
        content.append(['Condition:',"TEST"])
        content.append([Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11><u>Defined Outfall Channel:</u></font></para>', styles["Normal"]),''])
        content.append(['Material:',"TEST"])
        content.append(['Channel  Condition:',"TEST"])
        content.append(['Channel Bed Erosion:',"TEST"])
        content.append(['Channel Bank Erosion:',"TEST"])
        content.append([Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11><u>Undefined Outfall Channel:</u></font></para>', styles["Normal"]),''])
        content.append(['Material:',"TEST"])
        content.append(['Condition:',"TEST"])
        content.append([Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11><u>Outfall End Treatment:</u></font></para>', styles["Normal"]),''])
        content.append(['Type:',"TEST"])
        content.append(['Material:',"TEST"])
        content.append(['Condition:',"TEST"])

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