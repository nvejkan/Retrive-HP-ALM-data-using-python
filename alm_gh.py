import pywintypes
import win32com.client as w32c
from win32com.client import gencache, DispatchWithEvents, constants

def connect_server(qc, server):
        '''Connect to QC server
        input = str(http adress)
        output = bool(connected) TRUE/FALSE  '''
        try:
            qc.InitConnectionEx(server); 
        except:
            text = "Unable connect to Quality Center database: '%s'"%(server); 
        return qc.Connected;

def connect_login(qc, username, password):
    '''Login to QC server
    input = str(UserName), str(Password)
    output = bool(Logged) TRUE/FALSE  '''
    try:
        qc.Login(username, password);
    except pywintypes.com_error:
        text = unicode(err[2][2]);
    return qc.LoggedIn;

def connect_project(qc, domainname, projectname):
    '''Connect to Project in QC server
    input = str(DomainName), str(ProjectName)
    output = bool(ProjectConnected) TRUE/FALSE  '''

    try:
        qc.Connect(domainname, projectname)
    except pywintypes.com_error:
        text = "Repository of project '%s' in domain '%s' doesn't exist or is not accessible. Please contact your Site Administrator"%(projectname, domainname); 
    return qc.ProjectConnected;


def qc_instance():
        '''Create QualityServer instance under variable qc
        input = None
        output = bool(True/False)'''
        qc= None;
        try:
            qc = w32c.Dispatch("TDApiole80.TDConnection.1");
            text = "DLL QualityCenter file correctly Dispatched"
            return True, qc;
        except:
            return False, qc;

def qcConnect(server, username, password, domainname, projectname):
    print("Getting QC running files") ;

    status, qc = qc_instance();
    if status:
        print("Connecting to QC server");
        if connect_server(qc, server):
            ##connected to server
            print ("Checking username and password") ;
            if connect_login(qc, username, password):
                print ("Connecting to QC domain and project");
                if connect_project(qc, domainname, projectname):
                    text = "Connected"
                    connected = True;
                    return connected, text , qc;
                else:
                    text = "Not connected to Project in QC server.\nPlease, correct DomainName and/or ProjectName";
                    connected = False;
                    return connected, text;
            else:
                text = "Not logged to QC server.\nPlease, correct UserName and/or Password";
                connected = False;
                return connected, text;
        else:
            text = "Not connected to QC server.\nPlease, correct server http address"; 
            connected = False;
            return connected, text;
    else:
        connected = False;
        text = "Unable to find QualityCenter installation files.\nPlease connect first to QualityCenter by web page to install needed files" 
        return connected, text;
def get_bugs(qcConn):
    '''just following boiler plate from vbscript
    PS the SetFilter is not in QTA API, it uses Filter.  
    But due to the workarounds in 
    the very brilliant pythoncom code it supplies a virtual wrapper class
    called SetFilter - this is one of those gotchas '''

    BugFactory = qcConn.BugFactory
    BugFilter = BugFactory.Filter

    BugFilter.SetFilter("BG_STATUS", "Not Closed") 
    #NB - a lot of fields in QC are malleable - and vary from site to site. 
    #COntact your admins for a real list of fields you can adjust
    buglist = BugFilter.NewList()
    return buglist

def get_bugs_by_wt(qcConn,wt,status):
        '''just following boiler plate from vbscript
        PS the SetFilter is not in QTA API, it uses Filter.  
        But due to the workarounds in 
        the very brilliant pythoncom code it supplies a virtual wrapper class
        called SetFilter - this is one of those gotchas '''

        BugFactory = qcConn.BugFactory
        BugFilter = BugFactory.Filter

        if status is not None:
                BugFilter.SetFilter("BG_STATUS", status)
        #NB - a lot of fields in QC are malleable - and vary from site to site. 
        #COntact your admins for a real list of fields you can adjust

        BugFilter.SetFilter("BG_SUMMARY", "*{0}*".format(wt))
        buglist = BugFilter.NewList()
        return buglist  


server= r"https://alm.saas.hp.com/qcbin"
username= "xxx"
password= "xxx"
domainname= "xxx"
projectname= "xxx" 

connection_status, text,qc  = qcConnect(server, username, password, domainname, projectname);
print("connection_status:", connection_status)

if connection_status :
        bug_list = get_bugs(qc)
        for i in bug_list:
                print(i.summary)
        bug_254502 = get_bugs_by_wt(qc,254502,None) #any status
        for i in bug_254502:
                print(i.summary)

     
