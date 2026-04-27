# sap.py
import win32com.client
import time
import sys

def connect_sap():
  
    # Conecta ao SAP GUI Scripting e retorna a sessão ativa
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine

    connection = application.Children(0)
    session = connection.Children(0)
    time.sleep(1)
    session.CreateSession()
    time.sleep(4)
    
    new_session=connection.Children(connection.Children.Count - 1)
    return new_session
 
def sap_transaction(new_session, cliente):

    new_session.findById("wnd[0]/tbar[0]/okcd").text = "ZSD_AGROTIS_MANAGER"
    new_session.findById("wnd[0]").sendVKey(0)
    
    # Abre transação da api e preenche a seção produtor
    new_session.findById("wnd[0]/usr/cmbP_API").key = "1"
    new_session.findById("wnd[0]/usr/chkP_EXEC").selected = True
    new_session.findById("wnd[0]/usr/ctxtP_TDATE").text = "31129999"
    new_session.findById("wnd[0]/usr/ctxtSO_PART-LOW").text = cliente
    new_session.findById("wnd[0]/usr/ctxtSO_PART-LOW").setFocus
    new_session.findById("wnd[0]/usr/ctxtSO_PART-LOW").caretPosition = 8
    new_session.findById("wnd[0]/tbar[1]/btn[8]").press()
    time.sleep(0.3)
    message = new_session.findById("wnd[0]/sbar").Text
    print("Status:", message)
    new_session.findById("wnd[0]/tbar[0]/btn[15]").press()
    time.sleep(0.3)
    
    # Abre e preenche a seção propriedades
    new_session.findById("wnd[0]/tbar[0]/okcd").text = "ZSD_AGROTIS_MANAGER"
    new_session.findById("wnd[0]").sendVKey (0)
    new_session.findById("wnd[0]/usr/cmbP_API").key = "2"
    new_session.findById("wnd[0]/usr/chkP_EXEC").selected = True
    new_session.findById("wnd[0]/usr/ctxtP_TDATE").text = "31129999"
    new_session.findById("wnd[0]/usr/ctxtSO_PART-LOW").text = cliente
    new_session.findById("wnd[0]/usr/ctxtP_TDATE").setFocus
    new_session.findById("wnd[0]/usr/ctxtP_TDATE").caretPosition = 8
    new_session.findById("wnd[0]/tbar[1]/btn[8]").press()
    time.sleep(0.3)
    message = new_session.findById("wnd[0]/sbar").Text
    print("Status:", message)
    new_session.findById("wnd[0]/tbar[0]/btn[15]").press()

    # Abre o BP
    new_session.findById("wnd[0]/tbar[0]/okcd").text = "bp"
    new_session.findById("wnd[0]").sendVKey (0)
    new_session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2240/subSCREEN_1010_LEFT_AREA:SAPLBUS_LOCATOR:3100/tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3202/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3211/subSCREEN_3200_SEARCH_FIELDS_AREA:SAPLBUPA_DIALOG_SEARCH:2100/txtBUS_JOEL_SEARCH-PARTNER_NUMBER").text = cliente
    new_session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2240/subSCREEN_1010_LEFT_AREA:SAPLBUS_LOCATOR:3100/tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3202/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3211/subSCREEN_3200_SEARCH_FIELDS_AREA:SAPLBUPA_DIALOG_SEARCH:2100/txtBUS_JOEL_SEARCH-PARTNER_NUMBER").setFocus
    new_session.findById("wnd[0]/usr/subSCREEN_3000_RESIZING_AREA:SAPLBUS_LOCATOR:2240/subSCREEN_1010_LEFT_AREA:SAPLBUS_LOCATOR:3100/tabsGS_SCREEN_3100_TABSTRIP/tabpBUS_LOCATOR_TAB_02/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3202/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3211/subSCREEN_3200_SEARCH_FIELDS_AREA:SAPLBUPA_DIALOG_SEARCH:2100/txtBUS_JOEL_SEARCH-PARTNER_NUMBER").caretPosition = 8
    new_session.findById("wnd[0]").sendVKey (0)
