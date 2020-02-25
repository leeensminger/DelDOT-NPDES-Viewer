import  os, sys, traceback, arcpy ,pypyodbc, appConfig, conveyanceLocation,ReportSharedFunctions
import ReturnReportPipeOrSwale

def returnMidpointXY(OBJECTID,):
    try:
        ftrs = arcpy.da.SearchCursor(appConfig.GDBconn + "\\conveyances_evw", ["SHAPE@"], "objectid = " + OBJECTID)
        for ftr in ftrs:
            Midpoint = ftr[0].positionAlongLine(0.50,True).firstPoint
            return [Midpoint.X, Midpoint.Y]
            pass


    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass



def returnelements(OBJECTID,subTy):
    try:
   
        if subTy == "Swale":
            tbl = "SWALES_evw"
        else:
            tbl = "PIPE_SEGMENTS_evw"


        MidpointXY = returnMidpointXY(OBJECTID)
        elementPairs =    ([["Location Information" , conveyanceLocation.returnLocation_Information(OBJECTID,MidpointXY,tbl)]])

        if subTy == "Swale":             
            elementPairs.append(["Inventory and Inspection Information", ReturnReportPipeOrSwale.returnSwaleInventory_and_Inspection_Information(OBJECTID)])
        else:
            elementPairs.append(["Inventory and Inspection Information", ReturnReportPipeOrSwale.returnPipeInventory_and_Inspection_Information(OBJECTID)])

        result =  (ReportSharedFunctions.returnDataODBC("select count(*) FROM [deldot_migration2].[dbo].[WORK_ORDERS_evw] w join [deldot_migration2].[dbo].[conveyances_evw] s on s.GlobalID = w.[FEATURE_ID] " + 
            "where w.[STATUS] <> 'Cancelled' and s.OBJECTID=" + OBJECTID )[0])
        if result[0]  > 0:  
            elementPairs.append(["Work Order Information", ReturnReportPipeOrSwale.returnWorkOrder_Information(OBJECTID)])

        return elementPairs

    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass


def returnPhotoelement(OBJECTID,subTy,workingFolder):
    try:

        if subTy != "Swale":
            sqlInGID = "SELECT top 1   insp.GlobalID from [deldot_migration2].[dbo].CONVEYANCES_evw c join [deldot_migration2].[dbo].PIPE_SEGMENTS_evw p on p.CONVEYANCE_ID=c.GlobalID  join [deldot_migration2].[dbo].PIPE_INSPECTIONS_evw insp on insp.PIPE_FEATURE_ID=p.GlobalID   where  c.OBJECTID = " + OBJECTID  + " order by insp.INSPECTION_DATE desc;"
            inspGID = ReportSharedFunctions.returnDataODBC(sqlInGID)[0][0]
            
            titles = ["Upstream","Downstream"]
            sql1 = ("SELECT top 1 atch.DATA  " +
                    "from [deldot_migration2].[dbo].CONVEYANCES_evw c " +
                    "join [deldot_migration2].[dbo].PIPE_SEGMENTS_evw p on p.CONVEYANCE_ID=c.GlobalID  " +
                    "join [deldot_migration2].[dbo].PIPE_INSPECTIONS_evw insp on insp.PIPE_FEATURE_ID=p.GlobalID  " +
                    "join [deldot_migration2].[dbo].PIPE_INSP_PHOTOS_evw pip on  pip.PIPE_INSPECTION_ID = insp.GlobalID  " +
                    "join [deldot_migration2].[dbo].PIPE_INSP_PHOTOS__ATTACH_evw atch on atch.REL_GLOBALID=pip.GlobalID  " +
                      "where pip.PIPE_PHOTO_CAT = '")
            sql2 = ("' and atch.DATA_SIZE>0 and c.OBJECTID = " + OBJECTID + "  and insp.GlobalID= '" + inspGID + "'" )						  
        else:
            sqlInGID = "SELECT top 1   insp.GlobalID from [deldot_migration2].[dbo].CONVEYANCES_evw c join [deldot_migration2].[dbo].SWALES_evw p on p.CONVEYANCE_ID=c.GlobalID  join [deldot_migration2].[dbo].SWALE_INSPECTIONS_evw insp on insp.SWALE_ID=p.GlobalID   where  c.OBJECTID = " + OBJECTID  + " order by insp.INSPECTION_DATE desc;"
            inspGID = ReportSharedFunctions.returnDataODBC(sqlInGID)[0][0]
                        
            titles = ["General"]
            sql1 = ("SELECT top 1 atch.DATA  " +
                    "from [deldot_migration2].[dbo].CONVEYANCES_evw c " +
                    "join [deldot_migration2].[dbo].swales_evw s on s.CONVEYANCE_ID=c.GlobalID  " +
                    "join [deldot_migration2].[dbo].SWALE_INSPECTIONS_evw si on si.SWALE_ID = s.GlobalID " +
                    "join [deldot_migration2].[dbo].SWALE_INSP_PHOTOS_evw sip on sip.SWALE_INSPECTION_ID = si.GlobalID " +
                    "join  [deldot_migration2].[dbo].SWALE_INSP_PHOTOS__ATTACH_evw atch on atch.REL_GLOBALID=sip.GlobalID " +
                      "where sip.SWALE_PHOTO_CATEGORY  =  '")
            sql2 = ("' and atch.DATA_SIZE>0 and c.OBJECTID = " + OBJECTID + "  and si.GlobalID= '" + inspGID + "'" )

        imgs = ReportSharedFunctions.returnPhotoTable(titles, sql1, sql2, workingFolder)
  
        return imgs

    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass