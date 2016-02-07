If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/okcd").text = "ZMMR194A"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtP_WERKS").text = "206E"
session.findById("wnd[0]/usr/ctxtP_LGNUM").text = "CDB"
session.findById("wnd[0]/usr/ctxtP_PRVBE").text = "LINIA"
session.findById("wnd[0]/usr/ctxtP_MATNR").text = "MATERIAL"
session.findById("wnd[0]/usr/ctxtP_MATNR").setFocus
session.findById("wnd[0]/usr/ctxtP_MATNR").caretPosition = 8
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
