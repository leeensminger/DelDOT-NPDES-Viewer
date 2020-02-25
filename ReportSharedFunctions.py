import  os, sys, traceback, arcpy ,pypyodbc, appConfig
import datetime, time
import PIL,io
from reportlab.platypus import *
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, inch

from reportlab.lib.styles import getSampleStyleSheet

styles = getSampleStyleSheet()
ParaStyle = styles["Normal"]

pypyodbc.connect('Driver={SQL Server};' 'Server=vm-deldotsql\DELDOT;' 'Database=deldot_migration2;' 'uid=deldot;pwd=deldot#123')
connFolder = r"C:\Projects\DelDOT_NPDES_Reports\connections"

def returnDefectPhotoTable(typeOf, objectID,workingFolder):
    try:
        if typeOf == "structure":
            sql = ("SELECT wpa.DATA FROM [deldot_migration2].[dbo].[WORK_ORDERS_evw] w  " +
                "join [deldot_migration2].[dbo].[STRUCTURES_evw] s on s.GlobalID = w.[FEATURE_ID] " +
                "join [deldot_migration2].[dbo].WORK_ORDER_PHOTOS_evw wop on wop.WORK_ORDER_ID=w.GlobalID " +
                "join [deldot_migration2].[dbo].WORK_ORDER_PHOTOS__ATTACH_evw wpa on wpa.REL_GLOBALID = wop.GlobalID " +		  
                "where w.[STATUS] <> 'Cancelled' and s.OBJECTID=  " + objectID + " order by w.LAST_EDIT_DATE desc")
        elif typeOf == "conveyance":
            sql = ("SELECT wpa.DATA FROM [deldot_migration2].[dbo].[WORK_ORDERS_evw] w  " +
                "join [deldot_migration2].[dbo].[conveyanceS_evw] s on s.GlobalID = w.[FEATURE_ID] " +
                "join [deldot_migration2].[dbo].WORK_ORDER_PHOTOS_evw wop on wop.WORK_ORDER_ID=w.GlobalID " +
                "join [deldot_migration2].[dbo].WORK_ORDER_PHOTOS__ATTACH_evw wpa on wpa.REL_GLOBALID = wop.GlobalID " +		  
                "where w.[STATUS] <> 'Cancelled' and s.OBJECTID=  " + objectID + " order by w.LAST_EDIT_DATE desc")

        reportGrids = []

        cursor =  appConfig.connection.cursor()              
        cursor.execute(sql) 
        results = cursor.fetchall()

        picNum = 0
        for result in results:
            picNum += 1
            try:
                image = PIL.Image.open(io.BytesIO(result[0]))
                image.save(workingFolder + "\\" + str(picNum) + ".jpg")
                
                ht = (float(image.size[1])/float(image.size[0])) * 5
                pic = Image(workingFolder + "\\" + str(picNum) + ".jpg", width = 5*inch, height = ht*inch)
            except:
                tb = sys.exc_info()[2]
                tbinfo = traceback.format_tb(tb)[0]
                pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                        str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
                print pymsg; arcpy.AddError(pymsg)
                arcpy.AddMessage("Error retrieving " + "defect" + " photo .."  + pymsg)  
                pic = "JPEG data corrupt?"    
  


            content = [["Defect"  + " Photo"]]
            content.append([pic])
            reportGrid = Table(content, [ 5*inch])
            reportGrid.setStyle(TableStyle([
                                            ('SIZE', (0,0), (-1, -1), 11),  
                                            ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold') ,
                                            ('ALIGN', (0,0), (-1, -1), 'CENTER'),
                                            ('BOTTOMPADDING', (0,1), (-1, -1), 0),
                                            ('TOPPADDING', (0,1), (-1, -1), 0),
                                            ('LEFTPADDING', (0,1), (-1, -1), 0),
                                            ('RIGHTPADDING', (0,1), (-1, -1), 0),
                                            ('BOX', (0,1), (-1,-1),1, colors.black) 
                                        ]))
            reportGrids.append(reportGrid)

        return reportGrids
    
       
    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        arcpy.AddMessage(pymsg)   
        raise      
        pass        


def returnPhotoTableWQ(titles, sql1,sql2, withComment, width,workingFolder):
    try:
        reportGrids = []
        for title in titles: 
            sql = sql1+ title + sql2
            cursor = appConfig.connection.cursor()  
            
            cursor.execute(sql) 
            result = cursor.fetchall()

            if len(result)>0: #10/2018 replace block
                i=0 ###############BUT in AGS the pics are in the job folder!!!
                for r in result:
                    i+=1
                    try: 
                        comment = '' if r[1] == None or r[1].strip() == '' else Paragraph('<para alignment= "CENTER"><font name=Helvetica-Bold size=12>' +  r[1].strip() + '<br/></font>', styles["Normal"])  

                        image = PIL.Image.open(io.BytesIO(r[0]))
                        otherNum = ''
                        image.save(workingFolder + title + str(i) + "WQ.jpg") #workingFolder = "c:\\temp\\" 
                
                        ht = (float(image.size[1])/float(image.size[0])) * width
                        pic = Image(workingFolder + title + str(i) + "WQ.jpg", width = width*inch, height = ht*inch)

                    except:
                        tb = sys.exc_info()[2]
                        tbinfo = traceback.format_tb(tb)[0]
                        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
                        print pymsg; arcpy.AddError(pymsg)
                        arcpy.AddMessage("Error retrieving " + title + " photo .."  + pymsg)  
                        pic = "JPEG data corrupt?"    
  

                    content = [[pic]]
                    if withComment: 
                        content.append([comment])
                        content.append([''])
                    w = 2.8 if width ==2.4 else width
                    reportGrid = Table(content, [ w*inch])
                    reportGrid.setStyle(TableStyle([
                                               
                                                    ('ALIGN', (0,0), (-1, -1), 'CENTER'),('VALIGN', (0,0), (-1, -1), 'MIDDLE'),
                                                    ('BOTTOMPADDING', (0,1), (-1, -1), 0),
                                                    ('TOPPADDING', (0,1), (-1, -1), 0),
                                                    ('LEFTPADDING', (0,1), (-1, -1), 0),
                                                    ('RIGHTPADDING', (0,1), (-1, -1), 0)  
                                                ]))
                    reportGrids.append(reportGrid)

        return reportGrids
    
       
    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        arcpy.AddMessage(pymsg)   
        raise      
        pass       




def returnPhotoTable(titles, sql1,sql2,workingFolder):
    try:
        reportGrids = []
        for title in titles: 
            sql = sql1+ title + sql2
            cursor = appConfig.connection.cursor()  
            
            cursor.execute(sql) 
            result = cursor.fetchall()

            if len(result)>0:
                try: #problem when running same time can this go in the job folder?
                    image = PIL.Image.open(io.BytesIO(result[0][0]))
                    image.save(workingFolder + "\\" + title + ".jpg")
                
                    ht = (float(image.size[1])/float(image.size[0])) * 5
                    pic = Image(workingFolder + "\\" + title + ".jpg", width = 5*inch, height = ht*inch)
                except:
                    tb = sys.exc_info()[2]
                    tbinfo = traceback.format_tb(tb)[0]
                    pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                            str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
                    print pymsg; arcpy.AddError(pymsg)
                    arcpy.AddMessage("Error retrieving " + title + " photo .."  + pymsg)  
                    pic = "JPEG data corrupt?"    
  
 

                content = [[title + " Photo"]]
                content.append([pic])
                reportGrid = Table(content, [ 5*inch])
                reportGrid.setStyle(TableStyle([
                                                ('SIZE', (0,0), (-1, -1), 11),  
                                               ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold') ,
                                                ('ALIGN', (0,0), (-1, -1), 'CENTER'),
                                                ('BOTTOMPADDING', (0,1), (-1, -1), 0),
                                                ('TOPPADDING', (0,1), (-1, -1), 0),
                                                ('LEFTPADDING', (0,1), (-1, -1), 0),
                                                ('RIGHTPADDING', (0,1), (-1, -1), 0),
                                               ('BOX', (0,1), (-1,-1),1, colors.black) 
                                            ]))
                reportGrids.append(reportGrid)

        return reportGrids
    
       
    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        arcpy.AddMessage(pymsg)   
        raise      
        pass        


def boolVal(val):
    if val == '1':
        return "Yes"
    elif val == "2":
        return "No"
    elif val == "0":
        return "Unknown"

    return val


def valFromDomCode(dom, code):
    domains = arcpy.da.ListDomains(appConfig.GDBconn)
    for domain in domains:
        if domain.name == dom:
            coded_values = domain.codedValues
            for val, desc in coded_values.iteritems():
                if val == code:
                    return desc
    return code
    pass

def projectToLatLong(x,y):
    statePlane = arcpy.SpatialReference(2893) 
    WGS84 = arcpy.SpatialReference(4326)

    pt = None
    point = arcpy.PointGeometry(arcpy.Point(x,y),statePlane)

    new_point = point.projectAs(WGS84)

    pt=(new_point.firstPoint.X,new_point.firstPoint.Y)

    return pt

    pass


      

def findNearestRoad(x,y, FCviewTarget):
    try:
        SQL = (r"SELECT TOP 1 ROAD_NO, Shape.STDistance( geometry::STGeomFromText('POINT(" + x + " " + y + ")', 2235)) as dst FROM "   + FCviewTarget  +
               " WHERE Shape.STDistance( geometry::STGeomFromText('POINT(" + x + " " + y + ")', 2235)) IS NOT NULL " +
              " ORDER BY Shape.STDistance( geometry::STGeomFromText('POINT(" + x + " " + y + ")', 2235)); ")

        #GDBconnSQL = arcpy.ArcSDESQLExecute(appConfig.GDBconn_CL)
        results = returnDataODBCcl(SQL)

        #results = GDBconnSQL.execute(SQL)
        return results
    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass

def returnData(sql):
    try:
        GDBconnSQL = arcpy.ArcSDESQLExecute(appConfig.GDBconn)
        results = GDBconnSQL.execute(sql)
        return results

    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass

def returnDataODBCcloud(sql):
    try:

        cursor = appConfig.connectionCloud.cursor()  
        cursor.execute(sql) 
        return cursor.fetchall()
           

    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass      

		
		
def returnDataODBCcl(sql):
    try:

        cursor = appConfig.connectionCL.cursor()  
        cursor.execute(sql) 
        return cursor.fetchall()
           

    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass       
		
		
		
def returnDataODBC(sql):
    try:

        cursor = appConfig.connection.cursor()  
        cursor.execute(sql) 
        return cursor.fetchall()
           

    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass         