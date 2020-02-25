from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, inch
from reportlab.platypus import *
from reportlab.lib.styles import getSampleStyleSheet
 
import  os, sys, traceback , appConfig, arcpy
import datetime, time
import ReportSharedFunctions

styles = getSampleStyleSheet()
ParaStyle = styles["Normal"]
styleN = styles['Normal']

def returnLocation_Information(OBJECTID):
    try:

        sql = ("select top 1 ADDRESS_LINE1 ,CITY,ZIP,SUBDIVISION_NAME_GIS,WATERSHED,COUNTY,s.DISTRICT,MAINT_AREA, (case when ACCEPTANCE_DATE is null then '' else FORMAT(ACCEPTANCE_DATE, 'M/d/yyyy' ) end) as d ,LOCATION_CATEGORY,STRUCTURE_CLASSIFICATION, " + 
               "OWNER,LEGACY_STRUCT_NUM,CONTRACT_NUMBER,X_COORD,Y_COORD,STRUCTURE_COMMENTS , isnull(w.MAINT_AREA_WORK,' ') MAINT_AREA_WORK from structures_evw s left join WORK_ORDERS_evw w on w.FEATURE_ID = s.GlobalID  where s.OBJECTID = " + OBJECTID )
        
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
        Structure_Location = result[9] 
        Structure_Location = ReportSharedFunctions.valFromDomCode("D_FeatureLoc",Structure_Location)
        Structure_Classification = result[10] 
        Structure_Classification = ReportSharedFunctions.valFromDomCode("D_StructureClassification",Structure_Classification)

        Ownership = result[11]
        Ownership = ReportSharedFunctions.valFromDomCode("D_Ownership",Ownership)

        Legacy_Structure_Number = result[12].replace(".00000000","")


        Contract_Number = result[13]

        Point_X = result[14]
        Point_Y = result[15]
        Comments = result[16]
        MAINT_AREA_WORK =  result[17]
        #ROAD_NO = ReportSharedFunctions.findNearestRoad(str(bmpCent[0]),str(bmpCent[1]), "centerlines")[0] 
        #replaced 10/2018 then reverted 11/18
        #sqlROAD_NO = "select isnull(roadway_id,0) from structures_evw where OBJECTID =   " + OBJECTID 
        #arcpy.AddMessage("sqlROAD_NO = " + sqlROAD_NO)	
        #ROAD_NO = str(ReportSharedFunctions.returnDataODBC(sqlROAD_NO)[0][0])
        #arcpy.AddMessage("ROAD_NO = " + ROAD_NO)
        #if ROAD_NO=='0':ROAD_NO=''
        #Maintenance_Road_Number =   ROAD_NO
        ROAD_NO = ReportSharedFunctions.findNearestRoad(str(Point_X),str(Point_Y), "centerlines")[0] 
        Maintenance_Road_Number = ROAD_NO[0] 
        
        LL = ReportSharedFunctions.projectToLatLong(float(Point_X),float(Point_Y))      

        Latitude = str(round(LL[1],4))
        Longitude = str(round(LL[0],4))


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
        content.append(['Structure Location:',Structure_Location])
        content.append(['Structure Classification:',Structure_Classification])
        content.append(['Ownership:',Ownership])
        content.append(['Legacy Structure Number:',Legacy_Structure_Number])
        content.append(['Contract Number:',Contract_Number])
        content.append(['Latitude:',Latitude])
        content.append(['Longitude:',Longitude])
        content.append(['Point_X:',Point_X])
        content.append(['Point_Y:',Point_Y])
        content.append(['Comments:',Paragraph('<para alignment= "LEFT"><font size=11>' + Comments+'</font></para>', styles["Normal"])])



        reportGrid = Table(content, [2.5*inch, 5*inch])
        reportGrid.setStyle(TableStyle([('ALIGN', (0,0), (0, -1), 'RIGHT'), ('ALIGN', (1,0), (1, -1), 'LEFT'), ('VALIGN', (0,0), (0, -1), 'TOP'),  
                                       ('SIZE', (0,0), (-1, -1), 11),  ('VALIGN', (1,8), (1, 9), 'MIDDLE'),
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