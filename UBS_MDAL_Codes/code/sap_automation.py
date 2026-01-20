import os
import time
import shutil
import subprocess
from time import sleep
from PIL import ImageGrab
import win32com.client
from datetime import datetime
from cryptography.fernet import Fernet


# =========================================
# LOGGING
# =========================================
def log_message(log_file, filename, status, message):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"{ts} | {filename} | {status} | {message}\n"
    os.makedirs(os.path.dirname(log_file), exist_ok=True)
    with open(log_file, "a", encoding="utf-8") as f:
        f.write(line)


# =========================================
# POPUP HANDLER
# =========================================
def handle_all_popups(session, current_file, log_file, screenshot_dir):
    popup_index = 0
    messages = []

    while True:
        sleep(1)
        try:
            popup = session.findById("wnd[1]")
            popup_index += 1

            text_parts = []
            try:
                for c in popup.findById("usr").Children:
                    if hasattr(c, "Text") and c.Text.strip():
                        text_parts.append(c.Text.strip())
            except:
                pass

            message = " | ".join(text_parts) if text_parts else popup.Text
            messages.append(message)

            log_message(log_file, current_file, "POPUP", message)

            try:
                os.makedirs(screenshot_dir, exist_ok=True)
                base = os.path.splitext(os.path.basename(current_file))[0]
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                ImageGrab.grab().save(
                    os.path.join(
                        screenshot_dir,
                        f"{base}_popup_{popup_index}_{ts}.png"
                    )
                )
            except:
                pass

            try:
                popup.sendVKey(0)
            except:
                try:
                    popup.findById("tbar[0]/btn[0]").press()
                except:
                    pass

        except:
            break

    return messages


# =========================================
# LOAD CREDS
# =========================================

def load_creds_encrypted(enc_path, key_path):
    with open(key_path, "rb") as f:
        key = f.read()

    fernet = Fernet(key)

    with open(enc_path, "rb") as f:
        decrypted = fernet.decrypt(f.read()).decode("utf-8")

    creds = {}
    for line in decrypted.splitlines():
        if "=" in line and not line.strip().startswith("#"):
            k, v = line.split("=", 1)
            creds[k.strip()] = v.strip()

    return creds



# =========================================
# SAP LOGIN
# =========================================
def sap_login(creds):
    subprocess.Popen(creds["SAP_LOGON_PATH"])
    time.sleep(5)

    sap_gui = win32com.client.GetObject("SAPGUI")
    app = sap_gui.GetScriptingEngine
    app.OpenConnection(creds["SAP_CONNECTION_NAME"], True)
    time.sleep(3)

    session = app.Children(0).Children(0)

    session.findById("wnd[0]/usr/txtRSYST-MANDT").text = creds["CLIENT"]
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = creds["USER"]
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = creds["PASSWORD"]
    session.findById("wnd[0]/usr/txtRSYST-LANGU").text = creds.get("LANGUAGE", "EN")
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(3)

    return session


# =========================================
# UBS MDAL
# =========================================
def run_ubs_mdal(session, fil1, fil2, fil3):
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZFI_RFEBDK00"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(2)

    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    time.sleep(2)

    alv = session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell")
    alv.currentCellRow = 5
    alv.selectedRows = "5"
    alv.doubleClickCurrentCell()
    time.sleep(2)

    session.findById("wnd[0]/usr/ctxtPAR_FIL1").text = fil1
    session.findById("wnd[0]/usr/ctxtPAR_FIL2").text = fil2
    session.findById("wnd[0]/usr/ctxtPAR_FIL3").text = fil3

    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    time.sleep(3)


# =========================================
# FF_5 UPLOAD
# =========================================
def run_ff5_and_upload(session, creds, header, position):

    session.findById("wnd[0]").sendVKey(3)
    time.sleep(1)

    session.findById("wnd[0]/tbar[0]/okcd").text = "FF_5"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(2)

    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    time.sleep(1)

    session.findById("wnd[1]/usr/txtV-LOW").text = creds["VARIANT_NAME"]
    session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    time.sleep(1)

    session.findById("wnd[0]/usr/ctxtAUSZFILE").text = header
    session.findById("wnd[0]/usr/ctxtUMSFILE").text = position

    session.findById("wnd[0]").sendVKey(8)
    time.sleep(1)

    handle_all_popups(
        session,
        header,
        creds["LOG_FILE"],
        creds["SCREENSHOT_DIR"]
    )

    status = session.findById("wnd[0]/sbar").Text

    if "RFEKA200" in status or "error" in status.lower():
        session.findById("wnd[0]/tbar[0]/btn[12]").press()
        return False, status

    for _ in range(4):
        session.findById("wnd[0]/tbar[0]/btn[12]").press()
        time.sleep(1)

    return True, status


# =========================================
# MAIN ENTRY FOR GMAIL
# =========================================
def run_sap_upload(creds_path):
    """
    Returns:
        (success: bool, error_summary: str)
    """
    try:
        creds = load_creds_encrypted(
            enc_path=r"D:\UBS MDAL\code\creds.enc",
            key_path=r"D:\UBS MDAL\code\creds.key"
        )

        session = sap_login(creds)
        os.makedirs(creds["PROCESSED_DIR"], exist_ok=True)

        processed_any = False
        any_failed = False
        failure_details = []   # 🔴 COLLECT FILE ERRORS

        while True:
            files = [
                f for f in os.listdir(creds["INVOICE_DIR"])
                if os.path.isfile(os.path.join(creds["INVOICE_DIR"], f))
            ]
            if not files:
                break

            invoice = files[0]
            processed_any = True

            name, ext = os.path.splitext(invoice)
            fil1 = os.path.join(creds["INVOICE_DIR"], invoice)
            fil2 = os.path.join(creds["HP_DIR"], f"{name}h{ext}")
            fil3 = os.path.join(creds["HP_DIR"], f"{name}p{ext}")
            inv_p = os.path.join(creds["PROCESSED_DIR"], invoice)
            head_p = os.path.join(creds["PROCESSED_DIR"], f"{name}h{ext}")
            pos_p = os.path.join(creds["PROCESSED_DIR"], f"{name}p{ext}")

            try:
                run_ubs_mdal(session, fil1, fil2, fil3)

                shutil.move(fil2, head_p)
                shutil.move(fil3, pos_p)

                ok, msg = run_ff5_and_upload(session, creds, head_p, pos_p)
                shutil.move(fil1, inv_p)

                if not ok:
                    any_failed = True
                    failure_details.append(f"{invoice} → {msg}")

                log_message(
                    creds["LOG_FILE"],
                    invoice,
                    "SUCCESS" if ok else "FAILED",
                    msg
                )

            except Exception as e:
                any_failed = True
                failure_details.append(f"{invoice} → {str(e)}")

                try:
                    shutil.move(fil1, inv_p)
                except:
                    pass

                log_message(
                    creds["LOG_FILE"],
                    invoice,
                    "FAILED",
                    str(e)
                )

        # 🧠 FINAL RESULT
        if not processed_any:
            return False, "No invoice files found."

        if any_failed:
            return False, "\n".join(failure_details)

        return True, "All files uploaded successfully."

    except Exception as e:
        return False, f"SAP fatal error: {str(e)}"


