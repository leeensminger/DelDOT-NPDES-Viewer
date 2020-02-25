from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, inch
from reportlab.platypus import *
from reportlab.lib.styles import getSampleStyleSheet
 
import  os, sys, traceback 
import datetime, time
import ReportSharedFunctionsES2M,appConfigES2M

styles = getSampleStyleSheet()

def returnIRFpageOneHdr(IRF_ID):
    contentHead = [[]]

    for h in appConfigES2M.p1Head:
        contentHead.append([Paragraph('<para alignment= "CENTER"><font name=Helvetica-Bold size=12>' + h + '</font>', styles["Normal"])])

    head = Table(contentHead, [7.4*inch])
    head.setStyle(TableStyle([  ('ALIGN', (0,0), (-1, -1), 'CENTER'),('BOTTOMPADDING', (0,0), (-1,-1), 0)       ]))

    return head



def returnIRFpageOneBody(IRF_ID):
    try:

        sql = "SELECT DELDOT_CONTRACT_NUM, [Contractor_Name],[LOCATION_DESC],[ADDRESS],[CITY],[ZIP],[PROJECT_NAME],[INSPECTION_DATE],[TIME_IN],[TIME_OUT],[LAST_INSPECTION_DATE],[PLAN_EXPIRATION_DATE],[CURRENT_STATUS],[REASON_FOR_INSPECTION] "
        sql += " ,[NOI_NUM],[IRF_FORM_TYPE_ID],[FormTypem],[INSPECTION_RATING_FORM_ID],[PRESENT_DURING_INSPECTION]  FROM [es2m].[dbo].[V_INSPECTION_RATING_FORM_P1] "
        sql += " where [INSPECTION_RATING_FORM_ID] = " + IRF_ID

        res = ReportSharedFunctionsES2M.returnDataODBC(sql)[0]

        DELDOT_CONTRACT_NUM =  res[0]
        Contractor_Name =  res[1]
        Contractor_Name=Paragraph('<para alignment= "RIGHT"><font name=Helvetica size=11>' + Contractor_Name + '</font>', styles["Normal"])
        LOCATION_DESC =  res[2]
        ADDRESS =  res[3]
        CITY =  res[4]
        ZIP =  res[5]
        PROJECT_NAME =  res[6]
        PROJECT_NAME =Paragraph('<para alignment= "LEFT"><font name=Helvetica size=11>' + PROJECT_NAME + '</font>', styles["Normal"])
        INSPECTION_DATE =  res[7]
        TIME_IN =  res[8]
        TIME_OUT =  res[9]
        LAST_INSPECTION_DATE =  res[10]
        PLAN_EXPIRATION_DATE =  res[11]
        CURRENT_STATUS =  res[12]
        REASON_FOR_INSPECTION =  res[13]
        NOI_NUM =  res[14]
        IRF_FORM_TYPE_ID =  res[15]
        FormTypem =  res[16]
        INSPECTION_RATING_FORM_ID =  res[17]
        PRESENT_DURING_INSPECTION =  res[18]

        content = [['To:','To name placeheloder','','Date:',INSPECTION_DATE,'Contract No:',DELDOT_CONTRACT_NUM]]
        content.append(['','Firm/agency placeheloder'])
        content.append(['','','','Contract Name:',PROJECT_NAME])
        content.append(['From:','From name placeheloder','','Contractor:',Contractor_Name])
        content.append(['','Firm/agency placeheloder','',Paragraph('<para alignment= "RIGHT"><font name=Helvetica size=11>' + 'Present During Inspection:' + '</font>', styles["Normal"]),
                        Paragraph('<para alignment= "LEFT"><font name=Helvetica size=11>' + PRESENT_DURING_INSPECTION + '</font>', styles["Normal"])])
        content.append(['Time In:',TIME_IN])              
        content.append(['Time Out:',TIME_OUT,'',Paragraph('<para alignment= "RIGHT"><font name=Helvetica size=11>' + 'Date of Last Inspection:' + '</font>', styles["Normal"]),LAST_INSPECTION_DATE
                        ,Paragraph('<para alignment= "RIGHT"><font name=Helvetica size=11>' + 'Plan Expiration:' + '</font>', styles["Normal"]),PLAN_EXPIRATION_DATE]) 
        content.append(['Status:',CURRENT_STATUS,'','NOI#:',NOI_NUM])
        content.append([Paragraph('<para alignment= "RIGHT"><font name=Helvetica size=11>' + 'Reason For Inspection:' + '</font>', styles["Normal"]), REASON_FOR_INSPECTION  ])
        content.append([''])

        reportElement = Table(content, [1.1*inch,1.5*inch,0.65*inch,1.25*inch,1.25*inch,0.75*inch,1.0*inch])
        reportElement.setStyle(TableStyle([   ('VALIGN', (0,0), (-1, -1), 'TOP'),('ALIGN', (0,0), (0, -1), 'RIGHT'),
                                           ('SPAN',(4,2),(6,2)),
                                           ('ALIGN', (1,0), (1, -1), 'LEFT'),('ALIGN', (2,0), (2, -1), 'RIGHT'),
                                           ('ALIGN', (3,0), (3, -1), 'RIGHT'),('ALIGN', (4,0), (4, -1), 'LEFT'),
                                           ('ALIGN', (5,0), (5, -1), 'RIGHT'), ('ALIGN', (6,0), (6, -1), 'LEFT'),
                                       ('SIZE', (0,0), (-1, -1), 11) ,  ('BOTTOMPADDING', (0,0), (-1, -1), 0)]))                                          
                                      # ('BOX', (0,0), (-1,-1),1, colors.black) ,  ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black)        ])) #('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold') ,

        return reportElement

        pass
 
    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass    

        






def returnIRFpageOneTable(IRF_ID):
    try:

        content = ([['','Section',  Paragraph('<para alignment= "CENTER"><font name=Helvetica size=11>' +	'Number of Points Awarded (X)'+ '</font>', styles["Normal"]),	
                     Paragraph('<para alignment= "CENTER"><font name=Helvetica size=11>' +'Number of Points Available (Y)' + '</font>', styles["Normal"]),	
                     Paragraph('<para alignment= "CENTER"><font name=Helvetica size=11>' +'Percent Awarded for Section (X/Y)x100' + '</font>', styles["Normal"]),'']])
        content.append(['','1'])
        content.append(['','2'])
        content.append(['','3'])
        content.append(['',Paragraph('<para alignment= "CENTER"><font name=Helvetica size=11>' +'Weighted Average Percentage (1+2+3) x 70% divided by the number of sections rated' + '</font>', styles["Normal"])])
        content.append(['','4'])
        content.append(['','Section 4 x 20%'])
        content.append(['','5'])
        content.append(['','Section 5 x 10%'])
        #content.append([''])
        content.append(['','Total Rating Percentage (Add Box A, B, and C)= Total','','','75'])
        content.append(['Rating','100-90','89.9-80','79.9-70','69.9-60','<=59.9'])

        hiLiteScore = 75

        if hiLiteScore >= 90:
            hlCol = 1
        elif hiLiteScore >= 80:
            hlCol = 2
        elif hiLiteScore >= 70:
            hlCol = 3
        elif hiLiteScore >= 60:
            hlCol = 4
        elif hiLiteScore >= 0:
            hlCol = 5

        reportElement = Table(content, [1*inch,0.9*inch,1.7*inch,1.5*inch,1.5*inch,0.9*inch])
        reportElement.setStyle(TableStyle([   ('VALIGN', (0,0), (-1, -1), 'MIDDLE'),('ALIGN', (0,0), (-1, -1), 'CENTER'),     
                                       ('SIZE', (0,0), (-1, -1), 11)  ,('SPAN',(1,4),(3,4)),('SPAN',(1,6),(3,6)),('SPAN',(1,8),(3,8)),('SPAN',(1,9),(3,9)),                                          
                                      ('BOX', (1,0), (4,9),1, colors.black) ,  ('INNERGRID', (1,0), (4,9), 0.25, colors.black),
                               ('BACKGROUND',(hlCol,10),(hlCol,10),colors.lightgrey),('FONTNAME', (hlCol,10), (hlCol,10), 'Helvetica-Bold')]))


        return reportElement
 
    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass    


def returnIRFpageOneCC(IRF_ID):
    try:

        content = ([['CC:']])
        content.append(['AE','Vernon Lawton'])
        content.append(['CE','Javier Torrijos'])
        content.append(['AD','Art Wessell'])
        content.append(['URS','Jerry Katzmire'])
        content.append(['URS','David Lafferty'])
        content.append(['ES','Carol Sullivan'])
        content.append(['SWE','Vince Davis'])
        content.append(['ES2M','Mary Hamilton'])


      

        reportElement = Table(content, [.75*inch,1.2*inch],hAlign="LEFT")
        reportElement.setStyle(TableStyle([   ('VALIGN', (0,0), (-1, -1), 'MIDDLE'),('ALIGN', (0,1), (0, -1), 'RIGHT'),     
                                       ('SIZE', (0,0), (-1, -1), 10)  ,('SPAN',(0,0),(1,0)),                                    
                                      ('BOX', (0,0), (-1,-1),1, colors.black) ,  ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
                                      ('BOTTOMPADDING', (0,0), (-1,-1),1), ('TOPPADDING', (0,0), (-1,-1),1)
                            ]))


        return reportElement
 
    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass    


def returnIRFTableAfterPage1(IRF_ID):
    try:
        sql = "SELECT [SPEC_SET_NAME],[SPEC_NUMBER],[SPEC_STATEMENT],[POINTS_AVAILABLE],[DEFICIENCY_STATEMENT],[FIX_STATEMENT],[REGULATORY_TEXT],[INSPECTION_RESULT],  "
        sql += " [POINTS_AWARDED],[STATION_LOCATION],[DUE_DATE],[SPEC_CATEGORY],[DEFICIENT],[INSP_RATING_FRM_ID],[DEFAULT_DUE_DATE],[VISIBLE],[PARENT_ID]  FROM [es2m].[dbo].[V_INSPECTION_RATING_FORM_P2] "
        sql += " where [INSP_RATING_FRM_ID] =   " + IRF_ID + " order by [SPEC_NUMBER]"

        res = ReportSharedFunctionsES2M.returnDataODBC(sql)

        resArray = []
        for r in res:
            resArray.append(['' if x == None else x for x in r] )

        content = ([['Section 1: Is the Project within the scope of the approved Plan?']])
        content.append(['(*An automatic Rating <= 59.9% is assessed)'])
        content.append(['Point\nValue', 	'Spec',	'Statement',	'Y',	'N',	'Pts. *\nExcluded'	,'Points\nAwarded'])

        totPtsAvail = 0
        totPtsAwarded = 0
        for r in resArray:
            SPEC_NUMBER =  r[1]
            SPEC_STATEMENT =  r[2]
            SPEC_STATEMENT=Paragraph('<para alignment= "LEFT"><font name=Helvetica size=11>' + SPEC_STATEMENT +  '</font>', styles["Normal"])
            POINTS_AVAILABLE =  str(r[3])
            DEFICIENCY_STATEMENT =  r[4]
            FIX_STATEMENT =  r[5]
            REGULATORY_TEXT =  r[6]
            INSPECTION_RESULT =  r[7]
            POINTS_AWARDED =  str(r[8])
            STATION_LOCATION =  r[9]
            SPEC_CATEGORY =  r[11]
            DEFICIENT =  r[12]
            DEFAULT_DUE_DATE =  r[14]          

            if POINTS_AVAILABLE == '' or POINTS_AWARDED == '':
                pointsExcluded = ''
            else:
                pointsExcluded = r[3] - r[8]
                pointsExcluded = str(pointsExcluded)

            content.append([POINTS_AVAILABLE,SPEC_NUMBER,SPEC_STATEMENT,'?','?',pointsExcluded,POINTS_AWARDED])

            if POINTS_AVAILABLE != '' :
                totPtsAvail += int(POINTS_AVAILABLE)
            if POINTS_AWARDED != '' :
                totPtsAwarded += int(POINTS_AWARDED)

        totPtsExcuded = totPtsAvail - totPtsAwarded

        reportElement = Table(content, [0.64*inch,1*inch,3.5*inch,0.35*inch,0.35*inch,0.83*inch,0.83*inch])
        reportElement.setStyle(TableStyle([ ('SPAN',(0,0),(6,0)),('SPAN',(0,1),(6,1)),  ('VALIGN', (0,3), (-1, -1), 'TOP'), ('ALIGN', (0,0), (-1, -1), 'CENTER'),
                                           ('FONTNAME', (0,0), (0,0), 'Helvetica-Bold') ,
                                       ('SIZE', (0,0), (-1, -1), 11),    ('BOX', (0,0), (6,2),1, colors.black)   ,                                    
                                       ('BOX', (0,2), (-1,-1),1, colors.black) ,  ('INNERGRID', (0,2), (-1,-1), 0.25, colors.black)
                                    ]))

        return reportElement

        pass
 
    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass    


def returnNotes(IRF_ID):
    try:
        sql = "SELECT [IRF_ID],[NOTE_STATUS_ID],[INSPECTION_SPECIFICATION_ID] ,[NOTE_NUM],[CREATE_DATE],[CLOSED_DATE],[GEN_COM] FROM [dbo].[IRF_NOTES]   "        
        sql += " where [IRF_ID] =   " + IRF_ID + " order by [NOTE_NUM]"

        res = ReportSharedFunctionsES2M.returnDataODBC(sql)

        resArray = []
        for r in res:
            resArray.append(['' if x == None else x for x in r] )

        content = ([['NOTES']])
      
        i = 0
        for r in resArray:
            i +=1
            Note =  str(i) + ".  " + r[6]
            Note = Paragraph('<para alignment= "LEFT"><font name=Helvetica size=11>' + Note + '</font>', styles["Normal"])
      

            content.append([Note])


        reportElement = Table(content, [7.5*inch])
        reportElement.setStyle(TableStyle([   ('VALIGN', (0,0), (0, -1), 'TOP'),
                                       ('SIZE', (0,0), (-1, -1), 11),  
                                       ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold') ,
                                       ('BOX', (0,0), (-1,-1),1, colors.black) ,  ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black)
                                    ]))

        return reportElement

        pass
 
    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass    




def returnDeficiencies(IRF_ID):
    try:
        sql = "SELECT  [INSPECTION_RESULT] ,[STATION_LOCATION] ,[SPEC_CATEGORY] ,[SPEC_NUMBER] ,[DEFICIENCY_STATEMENT] ,[FIX_STATEMENT] ,[DUE_DATE] ,[INSPECTION_SPECIFICATION_ID]  FROM [es2m].[dbo].[V_DEFICIENCIES]  "        
        sql += " where [INSP_RATING_FRM_ID] =   " + IRF_ID  # current vs. previous, how?  order?

        res = ReportSharedFunctionsES2M.returnDataODBC(sql)

        sql = "SELECT [INSPECTION_SPECIFICATION_ID] ,[CLOSED_DATE] ,[CREATE_DATE] ,[GEN_COM] ,[NOTE_STATUS_ID]   FROM [es2m].[dbo].[V_DEFICIENCIES_NOTES] "
        sql += " where [INSP_RATING_FRM_ID] =   " + IRF_ID 
        resNotes = ReportSharedFunctionsES2M.returnDataODBC(sql)
        resNotes = ([['' if x == None else x for x in resNote] for resNote in resNotes])

        sql = "SELECT [PHOTO_NUMBER]  ,[STATION_LOCATION] ,[INSPECTION_SPEC_ID] ,[DEFICIENT]    FROM [dbo].[V_IRF_PHOTOS] where [INSPECTION_RATING_FORM_ID] = "  + IRF_ID 
        resPics = ReportSharedFunctionsES2M.returnDataODBC(sql)
        resPics = [['' if x == None else x for x in resPic] for resPic in resPics]
        resDificientPics = [x  for x in resPics if x[3] == 1]   

        resArray = []
        for r in res:
            resArray.append(['' if x == None else x for x in r] )
      
        i = 0
        for r in resArray:
            
            INSPECTION_RESULT =  r[0]
            STATION_LOCATION =  r[1]
            SPEC_CATEGORY =  r[2]
            SPEC_NUMBER =  r[3]
            DEFICIENCY_STATEMENT =  r[4]
            FIX_STATEMENT =  r[5]
            DUE_DATE =  r[6]
            INSPECTION_SPECIFICATION_ID =  r[7]

            assocNotes = [t for t in resNotes if t[0] == INSPECTION_SPECIFICATION_ID]
            resArray[i].append(assocNotes)

            assocPics = [t for t in resDificientPics if t[2] == INSPECTION_SPECIFICATION_ID]
            resArray[i].append(assocPics)            

            i +=1

            pass

            
 
      

        reportElement = ''

        return reportElement

        pass
 
    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass    


def returnPhotos(IRF_ID):
    try:
        hd1='Photographic Log'
        hd2 = ['Client Name:','Site Location:','Project No.']
        hd3 = ['Delaware Department of Transportation',appConfigES2M.SiteLoc, appConfigES2M.ContractNo]

        sql = "SELECT   [INSPECTION_PHOTO_ID] ,[PHOTO_NUMBER] ,[PHOTO_NUM] ,[NOTE_ID] ,[PHOTO_DATE] ,[STATION_LOCATION] ,[INSPECTION_SPEC_ID] ,[DEFICIENT] ,[CLOUD_FILE_NAME] ,[Attachment_Type_ID] ,[GEN_COM] ,[NAME]   FROM [dbo].[V_IRF_PHOTOS] where [INSPECTION_RATING_FORM_ID] = "  + IRF_ID 
        resPics = ReportSharedFunctionsES2M.returnDataODBC(sql)
        resPics = [['' if x == None else x for x in resPic] for resPic in resPics]

        content = [[]]
        for r in resPics:
            INSPECTION_PHOTO_ID =  r[0]
            PHOTO_NUMBER =  r[1]
            PHOTO_NUM =  r[2]
            NOTE_ID =  r[3]
            PHOTO_DATE =  r[4]
            STATION_LOCATION =  r[5]
            INSPECTION_SPEC_ID =  r[6]
            DEFICIENT =  r[7]
            CLOUD_FILE_NAME =  r[8]
            Attachment_Type_ID =  r[9]
            GEN_COM =  r[10]
            NAME =  r[11]
            pic = ReportSharedFunctionsES2M.returnPhoto(CLOUD_FILE_NAME)


            content.append(['','Photo'])
            content.append(['',pic])

        reportElement = Table(content, [ 2*inch, 5*inch]) #put this part into a def that returns the full table for a one-pic table and then for the content append and repeat header row
        reportElement.setStyle(TableStyle([
                                        ('SIZE', (0,0), (-1, -1), 11),  
                                        ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold') ,
                                        ('ALIGN', (0,0), (-1, -1), 'CENTER'),
                                        ('BOTTOMPADDING', (0,1), (-1, -1), 0),
                                        ('TOPPADDING', (0,1), (-1, -1), 0),
                                        ('LEFTPADDING', (0,1), (-1, -1), 0),
                                        ('RIGHTPADDING', (0,1), (-1, -1), 0),
                                        ('BOX', (0,1), (-1,-1),1, colors.black) 
                                    ]))


            



        return reportElement

        pass
 
    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass    


def returnIRF(IRF_ID):
    try:
        elements = []

        page1 = returnIRFpageOneHdr(IRF_ID)
        elements.append(page1)
        elements.append(Spacer(1, 0.2*inch))

        page1 = returnIRFpageOneBody(IRF_ID)
        elements.append(page1)
        elements.append(Spacer(1, 0.1*inch))

        page1 = returnIRFpageOneTable(IRF_ID)
        elements.append(page1)
        elements.append(Spacer(1, 0.2*inch))

        page1 = returnIRFpageOneCC(IRF_ID)
        elements.append(page1)
        elements.append(PageBreak())

        part2 = returnIRFTableAfterPage1(IRF_ID)
        elements.append(part2)
        elements.append(PageBreak())

        Notes = returnNotes(IRF_ID)
        elements.append(Notes)
        elements.append(PageBreak())

        Deficiencies = returnDeficiencies(IRF_ID)


        pics = returnPhotos(IRF_ID)
        elements.append(pics)

        return elements
 
    except:
        tb = sys.exc_info()[2]
        tbinfo = traceback.format_tb(tb)[0]
        pymsg = "PYTHON ERRORS:\nTraceback Info:\n" + tbinfo + "\nError Info:\n    " + \
                str(sys.exc_type)+ ": " + str(sys.exc_value) + "\n"
        print pymsg; arcpy.AddError(pymsg)
        raise
        pass         