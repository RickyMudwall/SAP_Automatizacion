import pyautogui
import os
import time
import logging
import pdb
import win32com.client
import subprocess

session = None
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


class BaseSteps:
    def __init__(self):
        pass

    def sendkeys(self, *keys):
        time.sleep(1)
        self.screenshot_evidence()
        pyautogui.hotkey(*keys)
        time.sleep(1)
        self.screenshot_evidence()

    def screenshot_evidence(self):
        logging.info("Captura de pantalla")
        screenshot = pyautogui.screenshot()
        timestamp = int(time.time())
        screenshot_filename = f"screenshot_{timestamp}.png"
        screenshot_path = os.path.join(os.getcwd(), screenshot_filename)
        screenshot.save(screenshot_path)

    def set_text_sap(self, session, screen, field, text):

        time.sleep(1)
        self.screenshot_evidence()
        path_screen = "wnd[" + screen + "]/usr/" + field
        session.findById("wnd[" + screen + "]/usr/" + field).setFocus()
        session.findById("wnd[" + screen + "]/usr/" + field).text = text
        time.sleep(1)
        self.screenshot_evidence()

    def select_field_sap(self, session, screen, field):
        time.sleep(1)
        self.screenshot_evidence()
        path_screen = "wnd[" + screen + "]/usr/" + field
        session.findById(path_screen).select()
        time.sleep(1)
        self.screenshot_evidence()

    def press_field_sap(self, session, screen, field):
        time.sleep(1)
        self.screenshot_evidence()
        path_screen = "wnd[" + screen + "]/" + field
        session.findById(path_screen).press()
        time.sleep(1)
        self.screenshot_evidence()

    def checkbox_sap(self, session, screen, field):
        time.sleep(1)
        self.screenshot_evidence()
        path_screen = "wnd[" + screen + "]/usr/" + field
        session.findById(path_screen).selected = True
        #time.sleep(1)
        self.screenshot_evidence()



    def select_doc_sap(self, session, text):
        session.findById("wnd[0]/usr/cntlCONTAINER_0111/shellcont/shell").setCurrentCell(0, "EXCEPT2")
        time.sleep(1)
        # Hacer doble clic en la celda actual ppara desplegar toda la tabla
        session.findById("wnd[0]/usr/cntlCONTAINER_0111/shellcont/shell").doubleClickCurrentCell()

        time.sleep(1)
        self.screenshot_evidence()  # Asume que esta función captura correctamente la pantalla.

        session.findById("wnd[0]/usr/cntlCONTAINER_0112/shellcont/shell").pressToolbarButton("&FIND")
        session.findById("wnd[1]/usr/chkGS_SEARCH-EXACT_WORD").selected = True
        session.findById("wnd[1]/usr/txtGS_SEARCH-VALUE").setFocus()
        session.findById("wnd[1]/usr/txtGS_SEARCH-VALUE").text = text
        session.findById("wnd[1]/usr/cmbGS_SEARCH-SEARCH_ORDER").key = "0"
        session.findById("wnd[1]/usr/chkGS_SEARCH-EXACT_WORD").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/tbar[0]/btn[12]").press()
        session.findById("wnd[0]/tbar[1]/btn[5]").press()

        time.sleep(1)
        self.screenshot_evidence()  # Captura final para evidencia



    def waitforelement(self, session, element_id, timeout):
        time.sleep(1)
        end_time = time.time() + timeout

        while time.time() < end_time:
            try:
                element = session.findById(element_id)
                if element:
                    return True
            except Exception as e:
                print(f"Error: {e}")
                pass
            time.sleep(0.5)

        return False
