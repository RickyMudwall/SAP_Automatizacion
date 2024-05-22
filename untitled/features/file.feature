#noinspection CucumberUndefinedStep
Feature: Realizar una acción



  @setup_sap
  Scenario: Configurar el acceso a SAP
  Given se ingresa a SAP
  When se logea con el usuario "CNS_VERITY" y contraseña "Verity.2025"

  @pago_proveedores
  Scenario: Configurar el acceso a SAP
  Given se ingresa a SAP
  When se logea con el usuario "CNS_VERITY" y contraseña "Verity.2025"
  Then se ingresa a la transaccion "f110"
  And se ingresan los datos fecha e identificador
  And se configura proceso 1 de pago
  And se selecciona el documento a pagar
  And se configura via de pago
  And Se ejecuta el pago
  And se cierra sap
  And se genera el reporte final





