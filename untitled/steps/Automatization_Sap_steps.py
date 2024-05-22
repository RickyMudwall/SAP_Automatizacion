import allure
import pyperclip
from behave import given, when, then
from base_steps import BaseSteps
from resumme_steps import GenerationSummary
import logging
import time
import os
import win32com.client
import subprocess
import sys
import pdb
from datetime import datetime, timedelta



logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
inst = BaseSteps()
carpeta_imagenes_home = "prueba"
session = None

now = datetime.now()
fecha_hoy = f"{now.day}.{now.month}.{now.year}"  # Construye la fecha manualmente

hora_ahora = datetime.now().strftime("%d%m%Y%H%M%S")  # Formato 'DDMMYYYYHHMMSS'
ultimos_cinco_digitos = hora_ahora[-5:]  # Extrae los últimos 5 dígitos


# Calcular la fecha de mañana
fecha_manana = datetime.now() + timedelta(days=1)

# Formatear la fecha para SAP GUI
fecha_manana_text = fecha_manana.strftime("%d.%m.%Y")  # Formato 'DD.MM.YYYY'
num_doc = None
cod_deudor = None
via_pago = None
cod_acreedor = None
tipo_pago = None
banco = None
cuenta = None



class MySteps:


    @staticmethod
    def inicialization_variable(filename):
        global num_doc, cod_deudor, via_pago, cod_acreedor,tipo_pago, banco, cuenta

        lines_csv = inst.read_file_and_check(filename)

        num_doc = lines_csv[0]
        cod_deudor = lines_csv[1]
        via_pago = lines_csv[2]
        cod_acreedor = lines_csv[3]
        tipo_pago = lines_csv[4]
        banco = lines_csv[5]
        cuenta = lines_csv[6]

        print("num_doc: ", num_doc)
        print("cod_deudor: ", cod_deudor)
        print("via_pago: ", via_pago)
        print("cod_acreedor: ", cod_acreedor)
        print("tipo_pago: ", tipo_pago)
        print("banco: ", banco)
        print("cuenta: ", cuenta)

    @staticmethod
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


    @given('se ingresa a SAP')
    def step_impl(context):
        global session



        path = r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe"
        subprocess.Popen(path)
        time.sleep(5)

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

        session = connection.Children(0)
        if not type(session) == win32com.client.CDispatch:
            connection = None
            application = None
            SapGuiAuto = None
            return

    @when('se logea con el usuario "{usuario}" y contraseña "{password}"')
    def step_impl(context, usuario, password):
        with allure.step("Comprobar Ventana"):
            allure.attach("Comprobar Ventana", name="Comprobar Ventana", attachment_type=allure.attachment_type.TEXT)
            assert inst.waitforelement(session, "wnd[0]", 10)
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = usuario
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus()
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 12
        session.findById("wnd[0]").sendVKey(0)

        #breakpoint()

        if inst.waitforelement(session, "wnd[1]/usr/radMULTI_LOGON_OPT2", 5):
            print("El elemento está disponible!")
            session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
            session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").setFocus()
            session.findById("wnd[1]").sendVKey(0)
        else:
            print("Tiempo de espera agotado. El elemento no está disponible.")

        assert inst.waitforelement(session, "wnd[0]/usr/btnBTN_CNC", 10)
        session.findById("wnd[0]/usr/btnBTN_CNC").press()

    @then('se ingresa a la transaccion "{trx}"')
    def step_impl(context, trx):
        with allure.step("Se comprueba la barra de transacciones"):
            allure.attach("Se comprueba la barra de transacciones", name="Se comprueba la barra de transacciones", attachment_type=allure.attachment_type.TEXT)
            assert inst.waitforelement(session, "wnd[0]/tbar[0]/okcd", 10)
        session.findById("wnd[0]/tbar[0]/okcd").text = trx
        session.findById("wnd[0]").sendVKey(0)


        MySteps.inicialization_variable("facturas.txt")



    @then('se cierra sap')
    def step_impl(context):
        assert inst.waitforelement(session, "wnd[0]/tbar[0]/btn[15]", 10)

        session.findById("wnd[0]/tbar[0]/btn[15]").press()

        assert inst.waitforelement(session, "wnd[0]", 10)


        session.findById("wnd[0]").close()

        assert inst.waitforelement(session, "wnd[1]/usr/btnSPOP-OPTION1", 10)

        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()



    @then('se ingresan los datos fecha e identificador')
    def step_impl(context):
        # Llamar a screenshot_evidence con la sección adecuada
        section = 'fecha e identificador'

        assert inst.waitforelement(session, "wnd[0]/usr/ctxtF110V-LAUFD", 10)

        inst.set_text_sap(session, "0", "ctxtF110V-LAUFD", fecha_hoy, section)
        inst.set_text_sap(session,"0", "ctxtF110V-LAUFI", ultimos_cinco_digitos, section)
        inst.select_field_sap(session, "0", "tabsF110_TABSTRIP/tabpPAR", section)

        inst.set_text_sap(session, "0", "tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/tblSAPF110VCTRL_FKTTAB/txtF110V-BUKLS[0,0]", cod_deudor, section)
        inst.set_text_sap(session, "0", "tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/tblSAPF110VCTRL_FKTTAB/ctxtF110V-ZWELS[1,0]", via_pago, section)
        inst.set_text_sap(session, "0", "tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/tblSAPF110VCTRL_FKTTAB/ctxtF110V-NEDAT[2,0]", fecha_manana_text, section)
        inst.set_text_sap(session, "0", "tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/subSUBSCR_SEL:SAPF110V:7004/ctxtR_LIFNR-LOW", cod_acreedor, section)


    @then('se configura proceso 1 de pago')
    def step_impl(context):
        # Llamar a screenshot_evidence con la sección adecuada
        section = 'se configura proceso 1 de pago'

        inst.select_field_sap(session, "0", "tabsF110_TABSTRIP/tabpSEL", section)
        inst.select_field_sap(session, "0", "tabsF110_TABSTRIP/tabpLOG", section)

        inst.checkbox_sap(session,"0","tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/chkF110V-XTRFA", section)
        inst.checkbox_sap(session,"0","tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/chkF110V-XTRZW", section)
        inst.checkbox_sap(session,"0","tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/chkF110V-XTRBL", section)
        inst.checkbox_sap(session,"0","tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/chkF110V-XTRBL", section)

        inst.select_field_sap(session, "0", "tabsF110_TABSTRIP/tabpPRI", section)
        inst.set_text_sap(session, "0", "tabsF110_TABSTRIP/tabpPRI/ssubSUBSCREEN_BODY:SAPF110V:0205/tblSAPF110VCTRL_DRPTAB/ctxtF110V-VARI1[1,4]", "TRAN_SDER_IP01", section)
        inst.select_field_sap(session, "0", "tabsF110_TABSTRIP/tabpSTA", section)
        inst.press_field_sap(session, "1", "usr/btnSPOP-OPTION1", section)
        inst.press_field_sap(session, "0", "tbar[1]/btn[13]", section)

        inst.checkbox_sap(session,"1","chkF110V-XSTRF", section)

        inst.press_field_sap(session, "1", "tbar[0]/btn[0]", section)
        inst.press_field_sap(session, "0", "tbar[1]/btn[14]", section)
        inst.press_field_sap(session, "0", "tbar[1]/btn[16]", section)
        inst.press_field_sap(session, "1", "tbar[0]/btn[13]", section)


    @then('se selecciona el documento a pagar')
    def step_impl(context):
        # Llamar a screenshot_evidence con la sección adecuada
        section = 'se selecciona el documento a pagar'


        inst.select_doc_sap(session, num_doc, section)
        inst.press_field_sap(session, "1", "tbar[0]/btn[6]", section)


    @then('se configura via de pago')
    def step_impl(context):
        # Llamar a screenshot_evidence con la sección adecuada
        section = 'se configura via de pago'

        inst.set_text_sap(session, "2", "ctxtREGUH-RZAWE", tipo_pago, section)
        inst.set_text_sap(session, "2", "ctxtREGUH-HBKID", banco, section)
        inst.set_text_sap(session, "2", "ctxtREGUH-HKTID", cuenta, section)
        inst.press_field_sap(session, "2", "tbar[0]/btn[13]", section)


    @then('Se ejecuta el pago')
    def step_impl(context):
        # Llamar a screenshot_evidence con la sección adecuada
        section = 'Se ejecuta el pago'

        inst.press_field_sap(session, "0", "tbar[0]/btn[11]", section)
        inst.press_field_sap(session, "0", "tbar[0]/btn[3]", section)
        inst.press_field_sap(session, "0", "tbar[0]/btn[3]", section)
        inst.press_field_sap(session, "0", "tbar[1]/btn[7]", section)
        inst.press_field_sap(session, "1", "tbar[0]/btn[0]", section)
        inst.press_field_sap(session, "0", "tbar[1]/btn[14]", section)

        #inst.update_file_and_rewrite("facturas.txt", num_doc)
        MySteps.update_file_and_rewrite("facturas.txt", num_doc)


    @then('se genera el reporte final')
    def step_impl(context):

        summary = GenerationSummary()  # Crea una instancia de GenerationSummary
        report_path = summary.generation_summary_report()  # Genera el informe
        print(f"Reporte generado en: {report_path}")  # Imprime la ruta del informe generado
