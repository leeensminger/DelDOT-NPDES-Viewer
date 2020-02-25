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



def returnOutfall_Information(OBJECTID):
    try:
        sql = ("SELECT top 1 ofs.NPDES,ofs.STREAM_DISTANCE,ofs.GNIS_NAME,ofs.GNIS_ID,oda.DRAINAGE_AREA_AC,insp.OVERALL_RATING,ofs.OF_PROTECTION_MATERIAL,"+
               "insp.OF_PROTECTION_CONDITION,ofs.DEFINED_OF_CHANNEL_MATERIAL,insp.DOC_CHANNEL_CONDITION,insp.DOC_CHANNEL_BED_ERO,insp.DOC_CHANNEL_BANK_ERO"+
               ",ofs.UNDEFINED_OF_CHANNEL_MATERIAL,insp.UDOC_CHANNEL_CONDITION,ofs.OF_END_TREATMENT_TYPE,ofs.OF_END_TREATMENT_MATERIAL,"+
               "insp.OF_END_TREATMENT_CONDITION,s.[OBJECTID],s.[GlobalID] FROM [deldot_migration2].[dbo].[STRUCTURES_evw] s " + 
               "join [deldot_migration2].[dbo].OUTFALLS_evw  ofs on ofs.STRUCTURE_ID = s.[GlobalID] " + 
               "join  [deldot_migration2].[dbo].[pipe_end_evw] pe on pe.STRUCTURE_ID = s.[GlobalID] " + 
               " left join  [deldot_migration2].[dbo].[pipe_end_inspections_evw] pei on pei.PIPE_END_ID = pe.[GlobalID] "+ 
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
        content.append([Paragraph('<para alignment= "LEFT"><font name=Helvetica-Bold size=11><u>Outfall End Treatment:</u></font></para>', styles["Normal"]),''])
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




def returnInventory_and_Inspection_Information(OBJECTID):
    try:

        sql = ("select top 1 pe.PIPE_END_TYPE,pe.TIDE_GATE_PRESENT,pe.IS_OUTLET,pe.DISCHARGE_DSTRM_TYPE,pe.OWNERSHIP_DOWNSTRM, " +
                "(case when pei.INSPECTION_DATE is null then '' else FORMAT(pei.INSPECTION_DATE, 'M/d/yyyy' ) end) as d ,pei.INSPECTED_REASON,pei.INSPECTED_REASON_DESCRIPTION,pei.OVERALL_RATING " +
                "from deldot_migration2.[dbo].[STRUCTURES_evw] s  " +
                "left join deldot_migration2.dbo.PIPE_END_evw pe on pe.STRUCTURE_ID=s.GlobalID " +
                "left join deldot_migration2.dbo.PIPE_END_INSPECTIONS_evw pei on pei.PIPE_END_ID=pe.GlobalID " +
                "where s.OBJECTID = " + OBJECTID  +
                "order by pei.INSPECTION_DATE desc " )     

 
        result = ReportSharedFunctions.returnDataODBC(sql)[0]
        result = ['' if x == None else x for x in result] 

        PIPE_END_Type = result[0]
        PIPE_END_Type = ReportSharedFunctions.valFromDomCode("D_EndTreatmentType", PIPE_END_Type)

        Tide_Gate_Present = result[1]
        if Tide_Gate_Present != "": Tide_Gate_Present = ReportSharedFunctions.boolVal(Tide_Gate_Present)

        IS_OUTLET = result[2]
        if IS_OUTLET != "": IS_OUTLET = ReportSharedFunctions.boolVal(IS_OUTLET)

        DISCHARGE_DSTRM_TYPE = result[3]
        DISCHARGE_DSTRM_TYPE = ReportSharedFunctions.valFromDomCode("D_Discharge_Dstrm_Type", DISCHARGE_DSTRM_TYPE)

        OWNERSHIP_DOWNSTRM = result[4]
        OWNERSHIP_DOWNSTRM = ReportSharedFunctions.valFromDomCode("D_Ownership", OWNERSHIP_DOWNSTRM)

        Last_Inspected_Date = result[5]    

        Not_Inspected_Reason = result[6] 
        if Not_Inspected_Reason != '': Not_Inspected_Reason = ReportSharedFunctions.valFromDomCode("D_ReasonStructure", Not_Inspected_Reason)

        Reason_Description = result[7]

        Overall_Rating = result[8]
        Overall_Rating = ReportSharedFunctions.valFromDomCode("D_Condition",Overall_Rating)



        content = [['Pipe End Type:',PIPE_END_Type]]
        content.append(['Tide Gate Present:',Tide_Gate_Present])
        content.append(['Is Outlet:',IS_OUTLET])
        content.append(['Discharge Downstream Type:',DISCHARGE_DSTRM_TYPE])
        content.append(['Ownership Downstream:',OWNERSHIP_DOWNSTRM])
        content.append(['Last Inspected Date:',Last_Inspected_Date])
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
         
