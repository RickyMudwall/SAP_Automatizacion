import pyautogui
import os
import time
import logging
import pdb
import win32com.client
import subprocess


logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


class BaseSteps:
    def __init__(self):
        pass

    def clickelement(self, carpeta_imagenes, imagen_objetivo, tiempo_espera):
        screen_width, screen_height = pyautogui.size()
        #print("Tamaño de la pantalla:", screen_width, screen_height)

        current_mouse_x, current_mouse_y = pyautogui.position()
        #print("Posición actual del mouse:", current_mouse_x, current_mouse_y)

        path_actual = os.getcwd()
        #print("Directorio actual:", path_actual)

        path_completo = os.path.join(path_actual, "img", carpeta_imagenes, imagen_objetivo)
        #print("Ruta completa de la imagen:", path_completo)

        time.sleep(tiempo_espera)
        self.screenshotevidencia()
        ubicacion = pyautogui.locateOnScreen(path_completo, confidence=0.9)
        #print(ubicacion)

        if ubicacion is not None:
            x, y = pyautogui.center(ubicacion)
            time.sleep(1)
            pyautogui.moveTo(x, y)
            pyautogui.click(x, y)
        else:
            print(f"No se pudo encontrar la imagen '{imagen_objetivo}' en la pantalla.")

    def sendkeys(self, *keys):
        time.sleep(1)
        self.screenshotevidencia()
        pyautogui.hotkey(*keys)
        time.sleep(1)
        self.screenshotevidencia()

    def screenshotevidencia(self):
        logging.info("Captura de pantalla")
        screenshot = pyautogui.screenshot()
        timestamp = int(time.time())
        screenshot_filename = f"screenshot_{timestamp}.png"
        screenshot_path = os.path.join(os.getcwd(), screenshot_filename)
        screenshot.save(screenshot_path)

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
