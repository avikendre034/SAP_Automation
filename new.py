
from turtle import st
import win32com.client
import os
import subprocess
import datetime
import time
import pandas as pd
from datetime import date
from pandas.tseries.offsets import DateOffset

def sapscript() :
    path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
    subprocess.Popen(path)
    time.sleep(10)  
    SapGuiAuto = win32com.client.GetObject('SAPGUI')

    if not type(SapGuiAuto) == win32com.client.CDispatch:
        return
    application = SapGuiAuto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return
    connection = application.OpenConnection("104. PSE-EBS-SCM Production Operations", True)
    if not type(connection) == win32com.client.CDispatch:
        application = None
        SapGuiAuto = None
        return
    session = connection.Children(0)
    if not type(session) == win32com.client.CDispatch:
        connection = None
        application = None
        SapGuiAuto = None
        return
    
    start_date = (date.today() - pd.offsets.MonthEnd(n=3)).date().strftime('%d.%m.%y')
    end_date = (date.today() + pd.offsets.MonthEnd(n=4)).date().strftime('%d.%m.%y')

    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "Yadavas"
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "Asmi222@"
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 7
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "/n/sapapo/rrp1"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtSV_SIMID").text = "000"
    session.findById("wnd[0]/usr/ctxtSV_SIMID").caretPosition = 3
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtSV_DTSTA").text = start_date
    session.findById("wnd[0]/usr/ctxtSV_DTEND").text = end_date
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/usr/btn%_SO_LOCNO_%_APP_%-VALU_PUSH").press
    session.findById("wnd[0]/usr/btn%_SO_LOCNO_%_APP_%-VALU_PUSH").press
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/btn%_SO_LOCNO_%_APP_%-VALU_PUSH").press
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/usr/ctxtSO_LOCNO-LOW").text = ""
    session.findById("wnd[0]/usr/ctxtSO_LOCNO-HIGH").setFocus
    session.findById("wnd[0]/usr/ctxtSO_LOCNO-HIGH").caretPosition = 0
    session.findById("wnd[0]/usr/btn%_SO_LOCNO_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[0]/usr/btn%_SO_LOCNO_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").caretPosition = 3
    session.findById("wnd[1]").sendVKey (0)
    session.findById("wnd[2]").close
    session.findById("wnd[0]/usr/btn%_SO_LOCNO_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[0]/usr/btn%_SO_LOCNO_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").caretPosition = 3
    session.findById("wnd[1]").sendVKey (2)
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/usr/btn%_SO_LOCNO_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[0]/usr/btn%_SO_LOCNO_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").caretPosition = 3
    session.findById("wnd[1]").sendVKey (2)
    session.findById("wnd[2]/tbar[0]/btn[12]").press()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "CA*"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "US*"
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSELSCR1/ssub%_SUBSCREEN_SELBLOCK:/SAPAPO/SAPLRRP_PT_ENTRY:2010/btn%_SO_MATNR_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSELSCR1/ssub%_SUBSCREEN_SELBLOCK:/SAPAPO/SAPLRRP_PT_ENTRY:2010/btn%_SO_MATNR_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL").select
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-ILOW_I[1,0]").text = "6100000"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,0]").text = "6199999"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,0]").setFocus
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,0]").caretPosition = 7
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    ##session.findById("wnd[2]/usr/cntlOPTION_CONTAINER/shellcont/shell").firstVisibleColumn = "TEXT"
    ##session.findById("wnd[0]").sendVKey (0)
    #session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "CA*"
    #session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "US*"
    #session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus
    #session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 3
    #session.findById("wnd[1]/tbar[0]/btn[8]").press()
    #session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL").select
    #session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-ILOW_I[1,0]").text = "6100000"
    #session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,0]").text = "6199999"
    ##session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,0]").setFocus
    ##session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL_255-IHIGH_I[2,0]").caretPosition = 7
    ##session.findById("wnd[1]/tbar[0]/btn[8]").press()
    ##session.findById("wnd[0]/tbar[1]/btn[8]").press()
    ##session.findById("wnd[0]/tbar[1]/btn[8]").press()
    ##session.findById("wnd[0]").sendVKey (0)
    ##session.findById("wnd[0]").sendVKey (0)
    ##session.findById("wnd[0]/tbar[1]/btn[8]").press()
    ##session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").pressToolbarButton="ORGRID_TOOLBAR_EXPAND"
    ##session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").pressToolbarContextButton="&MB_VARIANT"
    ##session.findById("wnd[0]/tbar[0]/btn[15]").press()
    ##session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").selectContextMenuItem="&LOAD"
    ##session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 10
    ##session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 1
    ##session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "10"
    ##session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
    ##session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").pressToolbarContextButton="&MB_EXPORT"
    ##session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").selectContextMenuItem="&XXL"
    ##session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").firstVisibleRow = 1
    ##session.findById("wnd[0]/usr/subREQMTS:/SAPAPO/SAPLRRP_REQMTS:3000/cntlALV_GRID_REQMTS/shellcont/shell").firstVisibleRow = 0
    ##session.findById("wnd[0]/tbar[0]/btn[15]").press()

    
sapscript()
