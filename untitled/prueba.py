import sys
import time

import win32com.client
import subprocess
from datetime import datetime


def waitForElement(session, element_id, timeout=10):
    end_time = time.time() + timeout

    while time.time() < end_time:
        try:
            element = session.findById(element_id)
            if element:
                return True
        except:
            pass
        time.sleep(0.5)  # Espera medio segundo antes de intentar nuevamente

    return False


def Main():
    try:
        path = r"C:\Program Files\SAP\FrontEnd\SAPGUI\saplogon.exe"
        subprocess.Popen(path)
        time.sleep(3)

        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not type(SapGuiAuto) == win32com.client.CDispatch:
            return

        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return

        connection = application.OpenConnection("QAS", True)
        if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return

        time.sleep(5)

        print("Connection:", connection)
        for i in range(connection.Children.Count):
            print(i, connection.Children(i))

        session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return

        time.sleep(1)
        assert waitForElement(session, "wnd[0]")
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "CNS_VERITY"
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "verity.2024"
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus()
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 12
        session.findById("wnd[0]").sendVKey(0)

        time.sleep(1)
        if waitForElement(session, "wnd[1]/usr/radMULTI_LOGON_OPT2"):
            print("El elemento está disponible!")
            session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
            session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus()
            session.findById("wnd[1]").sendVKey(0)
        else:
            print("Tiempo de espera agotado. El elemento no está disponible.")

        assert waitForElement(session, "wnd[0]/usr/btnBTN_CNC")
        session.findById("wnd[0]/usr/btnBTN_CNC").press()

        time.sleep(1)
        assert waitForElement(session, "wnd[0]/tbar[0]/okcd")
        session.findById("wnd[0]/tbar[0]/okcd").text = "FPL9"
        session.findById("wnd[0]").sendVKey(0)

        time.sleep(1)
        assert waitForElement(session, "wnd[0]/usr/ctxtFKKL1-GPART")
        session.findById("wnd[0]/usr/ctxtFKKL1-GPART").text = "212719471"
        session.findById("wnd[0]/usr/cmbFKKL1-LSTYP").key = "ALL"
        session.findById("wnd[0]/usr/cmbFKKL1-LSTYP").setFocus()
        session.findById("wnd[0]").sendVKey(0)

        time.sleep(1)
        assert waitForElement(session, "wnd[0]/usr/lbl[30,11]")
        session.findById("wnd[0]/usr/lbl[30,11]").setFocus()
        session.findById("wnd[0]/usr/lbl[30,11]").caretPosition = 3
        session.findById("wnd[0]/tbar[1]/btn[5]").press()
        session.findById("wnd[1]/usr/subBLOCK1:SAPLFKL0:0413/sub:SAPLFKL0:0413/ctxtRFKL0-VONSL[0,0]").text = "3000"
        session.findById("wnd[1]/usr/subBLOCK1:SAPLFKL0:0413/sub:SAPLFKL0:0413/ctxtRFKL0-VONSL[0,0]").caretPosition = 4
        session.findById("wnd[1]").sendVKey(0)

        time.sleep(1)
        # matriz_elementos = session.findById("wnd[0]/usr").Children

        # Recorre todos los elementos en la matriz e imprime sus textos
        # for elemento in matriz_elementos:
        #    texto = elemento.Text
        #    print("Texto del elemento:", texto)
        # texto = session.findById("wnd[0]/usr/lbl[6,13]").Text
        # print("Texto del componente: ", texto)

        assert waitForElement(session, "wnd[0]/usr/lbl[6,11]")
        session.findById("wnd[0]/usr/lbl[6,11]").setFocus()
        session.findById("wnd[0]/usr/lbl[6,11]").caretPosition = 10
        session.findById("wnd[0]/tbar[1]/btn[5]").press()
        session.findById(
            "wnd[1]/usr/subBLOCK1:SAPLFKL0:0413/sub:SAPLFKL0:0413/ctxtRFKL0-VONSL[0,0]").text = "590000186535"
        session.findById("wnd[1]/usr/subBLOCK1:SAPLFKL0:0413/sub:SAPLFKL0:0413/ctxtRFKL0-VONSL[0,0]").caretPosition = 12
        session.findById("wnd[1]").sendVKey(0)

        time.sleep(1)
        assert waitForElement(session, "wnd[0]/usr/lbl[6,13]")
        session.findById("wnd[0]/usr/lbl[6,13]").setFocus()
        session.findById("wnd[0]/usr/lbl[6,13]").caretPosition = 9
        session.findById("wnd[0]").sendVKey(2)

        time.sleep(1)
        assert waitForElement(session, "wnd[0]/mbar/menu[4]/menu[7]")
        session.findById("wnd[0]/mbar/menu[4]/menu[7]").select()

        matriz_elementos = session.findById("wnd[1]/usr/tblSAPLFKDRDEFREV_DISPLAY").Children
        cont = 0

        for elemento in matriz_elementos:
            cont = cont + 1

        for i in range(cont//38):
            texto = session.findById("wnd[1]/usr/tblSAPLFKDRDEFREV_DISPLAY/txtT_ALL_ITEMS-PDATE[4," + str(i) + "]").Text
            print(texto)

        fecha_inicio = datetime.strptime("01.03.2023", "%d.%m.%Y").date()
        fecha_fin = datetime.strptime("31.10.2023", "%d.%m.%Y").date()

        # Obtén la fecha actual
        fecha_actual = datetime.now().date()

        # Calcula la diferencia en meses
        if fecha_inicio <= fecha_actual <= fecha_fin:
            diferencia_fin = (fecha_fin.year - fecha_actual.year) * 12 + fecha_fin.month - fecha_actual.month + 1

        print(f"Diferencia en meses entre la fecha actual y 31.10.2023: {diferencia_fin} meses")



    except Exception as e:
        print("Error:", str(e))

    finally:
        session = None
        connection = None
        application = None
        SapGuiAuto = None


# -Main------------------------------------------------------------------
if __name__ == "__main__":
    Main()

# -End-------------------------------------------------------------------
