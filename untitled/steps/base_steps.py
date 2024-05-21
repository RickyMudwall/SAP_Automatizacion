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
        # Convertir todos los argumentos a cadenas
        keys = [str(key) for key in keys]
        time.sleep(1)
        self.screenshot_evidence()
        pyautogui.hotkey(*keys)
        time.sleep(1)
        self.screenshot_evidence()

    def read_file_and_check(self, filename):
        try:
            return_lines = []
            with open(filename, "r") as file:
                for line in file:
                    data = line.strip().split(';')
                    if data and data[-1] == '':
                        return_lines = data
                        return return_lines
            return []
        except FileNotFoundError:
            print(f"The file '{filename}' was not found.")
            return []
        except Exception as e:
            print(f"Error while reading the file '{filename}': {e}")
            return []

    def update_file_and_rewrite(filename, search_value):
        try:
            updated_lines = []
            with open(filename, "r") as file:
                for line in file:
                    data = line.strip().split(';')
                    if data and data[0] == search_value:
                        data[-1] = "PAGADO"
                    updated_lines.append(';'.join(data) + '\n')

            with open(filename, "w") as file:
                file.writelines(updated_lines)
        except FileNotFoundError:
            print(f"The file '{filename}' was not found.")
        except Exception as e:
            print(f"Error while updating the file '{filename}': {e}")

    def screenshot_evidence(self):
        logging.info("Captura de pantalla")
        screenshot = pyautogui.screenshot()
        timestamp = int(time.time())
        screenshot_filename = f'screenshot_{timestamp}.png'

        # Define the evidence directory
        evidence_dir = os.path.join(os.getcwd(), 'evidence')

        # Create the evidence directory if it doesn't exist
        if not os.path.exists(evidence_dir):
            os.makedirs(evidence_dir)

        # Define the full path for the screenshot
        screenshot_path = os.path.join(evidence_dir, screenshot_filename)

        # Save the screenshot
        screenshot.save(screenshot_path)

    def set_text_sap(self, session, screen, field, text):

        time.sleep(1)
        self.screenshot_evidence()
        path_screen = "wnd[" + screen + "]/usr/" + field
        session.findById(path_screen).setFocus()
        session.findById(path_screen).text = text
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
        time.sleep(1)
        session.findById("wnd[0]/usr/cntlCONTAINER_0111/shellcont/shell").selectedRows = "0"
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

