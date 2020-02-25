from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, inch
from reportlab.platypus import *
from reportlab.lib.styles import getSampleStyleSheet
import arcpy 
import  os, sys, traceback , appConfig
import datetime, time
import ReportSharedFunctions

styles = getSampleStyleSheet()
ParaStyle = styles["Normal"]
styleN = styles['Normal']

def returnLocation_Information(OBJECTID, midPtXY,tbl):
    try:
        if tbl == "PIPE_SEGMENTS_evw":
            comfld = "PIPES_COMMENTS"
        else:
            comfld = "SWALES_COMMENTS"

        sql = ("select top 1 c.ADDRESS_LINE1,c.CITY,c.ZIP,c.SUBDIVISION_NAME_GIS,c.WATERSHED, " +
                "c.COUNTY,c.DISTRICT,c.MAINT_AREA,(case when c.ACCEPTANCE_DATE is null then '' else FORMAT(c.ACCEPTANCE_DATE, 'M/d/yyyy' ) end) as d, " +
                "c.LOCATION_CATEGORY,c.[ROLE],c.[OWNER],c.LEGACY_CONV_NUM,c.CONTRACT_NUM, " +
                "s1.STRUCTURE_NUM,s2.STRUCTURE_NUM,sub.{1} , isnull(w.MAINT_AREA_WORK,' ') MAINT_AREA_WORK  " +
                "from [deldot_migration2].dbo.CONVEYANCES_evw c " +
                "left join [deldot_migration2].dbo.{0} sub on sub.CONVEYANCE_ID = c.GlobalID  " +
                "left join [deldot_migration2].dbo.STRUCTURES_evw s1 on s1.GlobalID=c.UPSTREAM_STRUCTURE_ID  " +
                "left join [deldot_migration2].dbo.STRUCTURES_evw s2 on s2.GlobalID=c.DOWNSTREAM_STRUCTURE_ID  " +
                " left join [deldot_migration2].dbo.WORK_ORDERS_evw w on w.FEATURE_ID = c.GlobalID    " +
                "where c.OBJECTID =   " + OBJECTID ).format(tbl,comfld)
        
        result = ReportSharedFunctions.returnDataODBC(sql)[0]
        result = ['' if x == None else x for x in result] 

        Address = result[0]
        City = result[1]
        Zip_Code = result[2]
        Development_Name = result[3]
        Watershed = result[4]
        Watershed = ReportSharedFunctions.valFromDomCode("D_Watershed",Watershed)

        County = result[5]
        County = ReportSharedFunctions.valFromDomCode("D_County",County)
        Maintenance_District = result[6]
        Maintenance_District = ReportSharedFunctions.valFromDomCode("D_District",Maintenance_District)
        Maintenance_Area = result[7] 
        Maintenance_Area = ReportSharedFunctions.valFromDomCode("D_Maintenance_Areas",Maintenance_Area)
        Acceptance_Date = result[8]
        Conveyance_Location = result[9] 
        Conveyance_Location = ReportSharedFunctions.valFromDomCode("D_FeatureLoc",Conveyance_Location)
        Role = result[10] 
        Role = ReportSharedFunctions.valFromDomCode("D_Role",Role)

        Ownership = result[11]
        Ownership = ReportSharedFunctions.valFromDomCode("D_Ownership",Ownership)

        Legacy_Conveyance_Number = result[12].replace(".00000000","")

        Contract_Number = result[13]

        UpStruc = result[14]
        DownStruc = result[15]

        Comments = result[16] #CONVEYANCES_evw
        MAINT_AREA_WORK =  result[17]

        #was replaced 10/2018 but then reverted 11/2018
        #sqlROAD_NO = "select isnull(roadway_id,0) from CONVEYANCES_evw where OBJECTID =   " + OBJECTID 
        #arcpy.AddMessage("sqlROAD_NO = " + sqlROAD_NO)	
        #ROAD_NO = str(ReportSharedFunctions.returnDataODBC(sqlROAD_NO)[0][0])
        #arcpy.AddMessage("ROAD_NO = " + ROAD_NO)
        #if ROAD_NO=='0':ROAD_NO=''
        #Maintenance_Road_Number =   ROAD_NO
        ROAD_NO = ReportSharedFunctions.findNearestRoad(str(midPtXY[0]),str(midPtXY[1]), "centerlines")[0] 
        Maintenance_Road_Number =   ROAD_NO[0]  

        content = [['Address:',Address]]
        content.append(['City:',City])
        content.append(['Zip Code:',Zip_Code])
        content.append(['Development Name:',Development_Name])
        content.append(['Watershed:',Watershed])
        content.append(['Maintenance Road Number:',Maintenance_Road_Number])
        content.append(['County:',County])
        content.append(['Maintenance District:',Maintenance_District])
        content.append(['Maintenance Area\n(Location)',Maintenance_Area])
        content.append(['Maintenance Area\n(Work Order)',MAINT_AREA_WORK])
        content.append(['Acceptance Date:',Acceptance_Date])
        content.append(['Conveyance Location:',Conveyance_Location])
        content.append(['Role:',Role])
        content.append(['Ownership:',Ownership])
        content.append(['Legacy Conveyance Number:',Legacy_Conveyance_Number])
        content.append(['Contract Number:',Contract_Number])
        content.append(['Upstream Structure:',UpStruc])
        content.append(['Downstream Structure:',DownStruc])
        content.append(['Comments:',Paragraph('<para alignment= "LEFT"><font size=11>' + Comments+'</font></para>', styles["Normal"])])



        reportGrid = Table(content, [2.5*inch, 5*inch])
        reportGrid.setStyle(TableStyle([('ALIGN', (0,0), (0, -1), 'RIGHT'), ('ALIGN', (1,0), (1, -1), 'LEFT'), ('VALIGN', (0,0), (0, -1), 'TOP'),  
                                       ('SIZE', (0,0), (-1, -1), 11),    ('VALIGN', (1,8), (1, 9), 'MIDDLE'), 
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