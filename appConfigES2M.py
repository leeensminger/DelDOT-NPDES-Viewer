import pypyodbc

connectionES2M = pypyodbc.connect('Driver={SQL Server};' 'Server=vm-deldotsql\DELDOT;' 'Database=ES2M;' 'uid=deldot;pwd=deldot#123')

workingPath = r"C:\Projects\DelDOT_ES2M_Reports"

fileAttachRootURL = 'http://deldot103.kci.com/deldotwebviewer/api/Task/DownloadDoc?finename='

#constants
p1Head = ['ES2M Inspection Rating Form','URS 4051 Ogletown Rd, Ste. 300 - Newark, DE','Ph: 302.781.590   Fax 302.781.5901']


#shared variables set at RT
inAGS = True

ContractNo = 'ContractNo'
SiteLoc = 'SiteLoc'