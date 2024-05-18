import allure
import pyperclip
from behave import given, when, then
from base_steps import BaseSteps
import logging
import time

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


class MySteps:



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

    @then('se busca el rut "{rut}"')
    def step_impl(context, rut):
        with allure.step("Se comprueba el campo rut"):
            allure.attach("Se comprueba el campo rut", name="Se comprueba el campo rut", attachment_type=allure.attachment_type.TEXT)
            assert inst.waitforelement(session, "wnd[0]/usr/ctxtFKKL1-GPART", 10)
        session.findById("wnd[0]/usr/ctxtFKKL1-GPART").text = rut
        session.findById("wnd[0]/usr/cmbFKKL1-LSTYP").key = "ALL"
        session.findById("wnd[0]/usr/cmbFKKL1-LSTYP").setFocus()
        session.findById("wnd[0]").sendVKey(0)

    @then('se selecciona el documento "{numDocumento}"')
    def step_impl(context, numDocumento):
        with allure.step("Se comprueba la ventana"):
            allure.attach("Se comprueba la ventana", name="Se comprueba la ventana", attachment_type=allure.attachment_type.TEXT)
            assert inst.waitforelement(session, "wnd[0]/usr/lbl[30,11]", 10)
        session.findById("wnd[0]/usr/lbl[30,11]").setFocus()
        session.findById("wnd[0]/usr/lbl[30,11]").caretPosition = 3
        session.findById("wnd[0]/tbar[1]/btn[5]").press()
        session.findById("wnd[1]/usr/subBLOCK1:SAPLFKL0:0413/sub:SAPLFKL0:0413/ctxtRFKL0-VONSL[0,0]").text = "3000"
        session.findById("wnd[1]/usr/subBLOCK1:SAPLFKL0:0413/sub:SAPLFKL0:0413/ctxtRFKL0-VONSL[0,0]").caretPosition = 4
        session.findById("wnd[1]").sendVKey(0)

        assert inst.waitforelement(session, "wnd[0]/usr/lbl[6,11]", 10)
        session.findById("wnd[0]/usr/lbl[6,11]").setFocus()
        session.findById("wnd[0]/usr/lbl[6,11]").caretPosition = 10
        session.findById("wnd[0]/tbar[1]/btn[5]").press()
        session.findById(
            "wnd[1]/usr/subBLOCK1:SAPLFKL0:0413/sub:SAPLFKL0:0413/ctxtRFKL0-VONSL[0,0]").text = numDocumento
        session.findById("wnd[1]/usr/subBLOCK1:SAPLFKL0:0413/sub:SAPLFKL0:0413/ctxtRFKL0-VONSL[0,0]").caretPosition = 12
        session.findById("wnd[1]").sendVKey(0)

    @then('se valida el despliegue de informacion del diferido')
    def step_impl(context):
        lista_pagos = [row['fecha_pago'] for row in context.table]

        assert inst.waitforelement(session, "wnd[0]/usr/lbl[6,13]", 10)
        session.findById("wnd[0]/usr/lbl[6,13]").setFocus()
        session.findById("wnd[0]/usr/lbl[6,13]").caretPosition = 9
        session.findById("wnd[0]").sendVKey(2)

        time.sleep(1)
        assert inst.waitforelement(session, "wnd[0]/mbar/menu[4]/menu[7]", 10)
        session.findById("wnd[0]/mbar/menu[4]/menu[7]").select()

        matriz_elementos = session.findById("wnd[1]/usr/tblSAPLFKDRDEFREV_DISPLAY").Children
        cont = 0

        for elemento in matriz_elementos:
            cont = cont + 1

        for i in range(cont//38):
            texto = session.findById("wnd[1]/usr/tblSAPLFKDRDEFREV_DISPLAY/txtT_ALL_ITEMS-PDATE[4," + str(i) + "]").Text
            print(texto)
            assert texto == lista_pagos[i]

        assert inst.waitforelement(session, "wnd[1]/tbar[0]/btn[0]", 10)
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

    @then('se ingresan los datos para inscripcion del alumno')
    def step_impl(context):
        for row in context.table:
            sociedad = row['sociedad']
            int_comercial = row['int_comercial']
            clasific_inscripcion = row['clasific_inscripcion']
            tp_objeto = row['tp_objeto']
            id_objeto = row['id_objeto']
            monto_descuento = row['monto_descuento']
            año_academico = row['año_academico']
            periodo_academico = row['periodo_academico']

        assert inst.waitforelement(session, "wnd[0]/usr/ctxtI_BUKRS", 10)
        session.findById("wnd[0]/usr/ctxtI_BUKRS").text = sociedad
        session.findById("wnd[0]/usr/ctxtI_PARTN").text = int_comercial
        session.findById("wnd[0]/usr/ctxtI_REGCLA").text = clasific_inscripcion
        session.findById("wnd[0]/usr/ctxtI_OTYPE").text = tp_objeto
        session.findById("wnd[0]/usr/txtI_OBJID").text = id_objeto
        session.findById("wnd[0]/usr/txtI_DESCUE").text = monto_descuento
        session.findById("wnd[0]/usr/ctxtI_PERYR").text = año_academico
        session.findById("wnd[0]/usr/ctxtI_PERID").text = periodo_academico
        session.findById("wnd[0]/usr/ctxtI_PERID").setFocus()
        session.findById("wnd[0]/usr/ctxtI_PERID").caretPosition = 3
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

    @then('se valida la visualizacion de los documentos')
    def step_impl(context):
        lista_documentos = [row['documentos'] for row in context.table]
        lista_codigo_documentos = [row['codigo_documento'] for row in context.table]

        assert inst.waitforelement(session, "wnd[0]/usr/cntlGRID1/shellcont/shell", 60)
        assert inst.waitforelement(session, "wnd[0]/tbar[1]/btn[45]", 60)

        session.findById("wnd[0]/tbar[1]/btn[45]").press()

        assert inst.waitforelement(session, "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]", 60)
        assert inst.waitforelement(session, "wnd[1]/tbar[0]/btn[0]", 60)

        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        assert inst.waitforelement(session, "wnd[0]/usr/cntlGRID1/shellcont/shell", 60)

        datos_del_portapapeles = pyperclip.paste()

        print(datos_del_portapapeles)

        lineas = datos_del_portapapeles.split('\n')
        filas_con_S = []

        for linea in lineas:
            if linea.startswith("|S"):
                campos = linea.split("|")
                campos_limpios = [campo.strip() for campo in campos]
                filas_con_S.append(campos_limpios)
        cont = 0
        for fila in filas_con_S:
            if len(fila) >= 10:
                campo_5 = fila[4]
                campo_8 = fila[7]
                campo_9 = fila[8]

                assert "Contabilizado: Doc. No.: " + lista_codigo_documentos[cont] in campo_5, f"Campo 5 no contiene '{lista_codigo_documentos[cont]}': {campo_5}"
                assert lista_documentos[cont] in campo_8, f"Campo 8 no contiene '{lista_documentos[cont]}': {campo_8}"
                assert lista_codigo_documentos[cont] in campo_9, f"Campo 9 no contiene '{lista_codigo_documentos[cont]}': {campo_9}"
            cont += 1

    @then('se cierra sap')
    def step_impl(context):
        assert inst.waitforelement(session, "wnd[0]/tbar[0]/btn[15]", 10)
        session.findById("wnd[0]/tbar[0]/btn[15]").press()

        assert inst.waitforelement(session, "wnd[0]", 10)
        session.findById("wnd[0]").close()

        assert inst.waitforelement(session, "wnd[1]/usr/btnSPOP-OPTION1", 10)
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()

    @then('se ingresan los datos para caja de pago rut "{rut}" y division "{division}"')
    def step_impl(context, rut, division):
        global rut_validation
        rut_validation = rut

        assert inst.waitforelement(session, "wnd[0]/usr/radR_2", 10)
        session.findById("wnd[0]/usr/radR_2").setFocus()
        session.findById("wnd[0]/usr/radR_2").select()

        assert inst.waitforelement(session, "wnd[0]/usr/txtP_GPART", 10)
        session.findById("wnd[0]/usr/txtP_GPART").text = rut
        session.findById("wnd[0]/usr/txtP_GSBER").text = division
        session.findById("wnd[0]/usr/txtP_GSBER").setFocus()
        session.findById("wnd[0]/usr/txtP_GSBER").caretPosition = 4
        session.findById("wnd[0]").sendVKey(0)

        assert inst.waitforelement(session, "wnd[0]/tbar[1]/btn[8]", 10)
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

    @then('se selecciona documento a pagar')
    def step_impl(context):
        for row in context.table:
            documento = row['documentos']

            assert inst.waitforelement(session, "wnd[0]/usr/cntlCONT_ALV/shellcont/shell", 10)
            session.findById("wnd[0]/usr/cntlCONT_ALV/shellcont/shell").setCurrentCell(-1, "OPBEL")
            session.findById("wnd[0]/usr/cntlCONT_ALV/shellcont/shell").firstVisibleRow = 31
            session.findById("wnd[0]/usr/cntlCONT_ALV/shellcont/shell").selectColumn("OPBEL")
            session.findById("wnd[0]/usr/cntlCONT_ALV/shellcont/shell").pressToolbarButton("&MB_FILTER")
            session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = documento
            session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 12
            session.findById("wnd[1]/tbar[0]/btn[0]").press()

            try:
                assert inst.waitforelement(session, "wnd[0]/usr/cntlCONT_ALV/shellcont/shell", 10)
                session.findById("wnd[0]/usr/cntlCONT_ALV/shellcont/shell").modifyCheckbox(0, "CHECKBOX", True)
                break
            except Exception as e:
                session.findById("wnd[0]/usr/cntlCONT_ALV/shellcont/shell").pressToolbarContextButton("&MB_FILTER")
                session.findById("wnd[0]/usr/cntlCONT_ALV/shellcont/shell").selectContextMenuItem("&DELETE_FILTER")
                pass

        session.findById("wnd[0]/usr/cntlCONT_ALV/shellcont/shell").triggerModified()
        assert inst.waitforelement(session, "wnd[0]/usr/btnCALCULAR", 10)
        session.findById("wnd[0]/usr/btnCALCULAR").press()

    @then('se contabiliza el documento con medio de pago "{medio_pago}" y monto "{monto}"')
    def step_impl(context, medio_pago, monto):
        global var_medio_pago
        var_medio_pago = medio_pago
        assert inst.waitforelement(session, "wnd[0]/usr/cmbV_FORMA_PAGO", 10)
        session.findById("wnd[0]/usr/cmbV_FORMA_PAGO").key = medio_pago
        session.findById("wnd[0]/usr/txtZFICA_FORMAS_PAGO-MONTO_PAGO").text = monto
        session.findById("wnd[0]/usr/txtZFICA_FORMAS_PAGO-MONTO_PAGO").setFocus()
        session.findById("wnd[0]/usr/txtZFICA_FORMAS_PAGO-MONTO_PAGO").caretPosition = 6
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/usr/btnAGREGAR_REGISTRO").press()

        assert inst.waitforelement(session, "wnd[0]/usr/cntlCONT_ALV2/shellcont/shell", 10)
        session.findById("wnd[0]/usr/cntlCONT_ALV2/shellcont/shell").modifyCheckbox(0, "CHECKBOX", True)
        session.findById("wnd[0]/usr/cntlCONT_ALV2/shellcont/shell").triggerModified()

        assert inst.waitforelement(session, "wnd[0]/usr/cntlCONT_ALV2/shellcont/shell", 10)
        assert inst.waitforelement(session, "wnd[0]/tbar[1]/btn[8]", 10)
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        assert inst.waitforelement(session, "wnd[1]/tbar[0]/btn[12]", 60)
        session.findById("wnd[1]/tbar[0]/btn[12]").press()

    @then('se valida el log del proceso de pago')
    def step_impl(context):
        assert inst.waitforelement(session, "wnd[0]/usr/cntlCONT_ALV4/shellcont/shell", 10)
        session.findById("wnd[0]/usr/cntlCONT_ALV4/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")

        assert inst.waitforelement(session, "wnd[0]/usr/cntlCONT_ALV4/shellcont/shell", 10)
        session.findById("wnd[0]/usr/cntlCONT_ALV4/shellcont/shell").selectContextMenuItem("&PC")

        assert inst.waitforelement(session, "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]", 60)
        assert inst.waitforelement(session, "wnd[1]/tbar[0]/btn[0]", 60)
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[4,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        assert inst.waitforelement(session, "wnd[0]/usr/cntlCONT_ALV4/shellcont/shell", 60)

        datos_del_portapapeles = pyperclip.paste()

        print(datos_del_portapapeles)

        lineas = datos_del_portapapeles.split('\n')
        filas_con_S = []

        for linea in lineas:
            if linea.startswith("|"+rut_validation):
                campos = linea.split("|")
                campos_limpios = [campo.strip() for campo in campos]
                filas_con_S.append(campos_limpios)
        cont = 0
        for fila in filas_con_S:
            if len(fila) >= 1:
                campo_1 = fila[1]
                campo_2 = fila[2]
                campo_3 = fila[3]

                #print("Texto: " + campo_1)
                #print("Texto: " + campo_2)
                #print("Texto: " + campo_3)

                assert rut_validation in campo_1, f"Campo 1 no contiene '{rut_validation}': {campo_1}"
                assert "5800" in campo_2, f"Campo 8 no contiene '5800': {campo_2}"
                assert "Registro de pago" + var_medio_pago + " correcto" in campo_3, f"Campo 3 no contiene '{var_medio_pago}': {campo_3}"




    @then('se ingresan los datos fecha e identificador "{cod_deudor}" "{via_pago}" "{cod_acreedor}"')
    def step_impl(context):

        assert inst.waitforelement(session, "wnd[0]/usr/ctxtF110V-LAUFD", 10)

        inst.set_text_sap(session, "0", "ctxtF110V-LAUFD", fecha_hoy)
        inst.set_text_sap(session,"0", "ctxtF110V-LAUFI", ultimos_cinco_digitos)
        inst.select_field_sap(session, "0", "tabsF110_TABSTRIP/tabpPAR")
        inst.set_text_sap(session, "0", "tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/tblSAPF110VCTRL_FKTTAB/txtF110V-BUKLS[0,0]", "IP01")
        inst.set_text_sap(session, "0", "tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/tblSAPF110VCTRL_FKTTAB/ctxtF110V-ZWELS[1,0]", "TEC")
        inst.set_text_sap(session, "0", "tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/tblSAPF110VCTRL_FKTTAB/ctxtF110V-NEDAT[2,0]", fecha_manana_text)
        inst.set_text_sap(session, "0", "tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/subSUBSCR_SEL:SAPF110V:7004/ctxtR_LIFNR-LOW", "35")


    @then('se configura proceso 1 de pago')
    def step_impl(context):
        inst.select_field_sap(session, "0", "tabsF110_TABSTRIP/tabpSEL")
        inst.select_field_sap(session, "0", "tabsF110_TABSTRIP/tabpLOG")

        inst.checkbox_sap(session,"0","tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/chkF110V-XTRFA")
        inst.checkbox_sap(session,"0","tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/chkF110V-XTRZW")
        inst.checkbox_sap(session,"0","tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/chkF110V-XTRBL")
        inst.checkbox_sap(session,"0","tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/chkF110V-XTRBL")

        inst.select_field_sap(session, "0", "tabsF110_TABSTRIP/tabpPRI")
        inst.set_text_sap(session, "0", "tabsF110_TABSTRIP/tabpPRI/ssubSUBSCREEN_BODY:SAPF110V:0205/tblSAPF110VCTRL_DRPTAB/ctxtF110V-VARI1[1,4]", "TRAN_SDER_IP01")
        inst.select_field_sap(session, "0", "tabsF110_TABSTRIP/tabpSTA")
        inst.press_field_sap(session, "1", "usr/btnSPOP-OPTION1")
        inst.press_field_sap(session, "0", "tbar[1]/btn[13]")

        inst.checkbox_sap(session,"1","chkF110V-XSTRF")

        inst.press_field_sap(session, "1", "tbar[0]/btn[0]")
        inst.press_field_sap(session, "0", "tbar[1]/btn[14]")
        inst.press_field_sap(session, "0", "tbar[1]/btn[16]")
        inst.press_field_sap(session, "1", "tbar[0]/btn[13]")


    @then('se selecciona el documento a pagar "{documento_pago}"')
    def step_impl(context, documento_pago):
        # Seleccion del documento a pagar
        inst.select_doc_sap(session, documento_pago)
        inst.press_field_sap(session, "1", "tbar[0]/btn[6]")


    @then('se configura via de pago "{tipo_pago}" "{banco}" "{cuenta}"')
    def step_impl(context, tipo_pago, banco, cuenta):

        inst.set_text_sap(session, "2", "ctxtREGUH-RZAWE", tipo_pago)
        inst.set_text_sap(session, "2", "ctxtREGUH-HBKID", banco)
        inst.set_text_sap(session, "2", "ctxtREGUH-HKTID", cuenta)
        inst.press_field_sap(session, "2", "tbar[0]/btn[13]")


    @then('Se ejecuta el pago')
    def step_impl(context):
        #Seccion para ejecutar el pago
        inst.press_field_sap(session, "0", "tbar[0]/btn[11]")
        inst.press_field_sap(session, "0", "tbar[0]/btn[3]")
        inst.press_field_sap(session, "0", "tbar[0]/btn[3]")
        inst.press_field_sap(session, "0", "tbar[1]/btn[7]")
        inst.press_field_sap(session, "1", "tbar[0]/btn[0]")
        inst.press_field_sap(session, "0", "tbar[1]/btn[14]")
