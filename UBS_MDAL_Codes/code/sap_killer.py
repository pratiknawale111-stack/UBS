import os
import time
import win32com.client


def close_sap():
    print("🛑 Logging out from SAP...")

    try:
        sap_gui = win32com.client.GetObject("SAPGUI")
        app = sap_gui.GetScriptingEngine

        # Loop through all connections & sessions
        for conn in app.Children:
            for session in conn.Children:
                try:
                    session.findById("wnd[0]/tbar[0]/btn[15]").press()  # Log off
                    time.sleep(1)

                    # Confirm logoff popup
                    try:
                        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                    except:
                        pass

                except Exception:
                    pass

        time.sleep(2)

    except Exception:
        print("⚠ SAP scripting not available, forcing close.")

    print("🛑 Closing SAP GUI...")
    os.system("taskkill /IM saplogon.exe /F")
    os.system("taskkill /IM sapgui.exe /F")

    print("✅ SAP logged out and closed.")
