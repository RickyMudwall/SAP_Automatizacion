#noinspection CucumberUndefinedStep
Feature: Realizar una acción



  @setup_sap
  Scenario: Configurar el acceso a SAP
  Given se ingresa a SAP
  When se logea con el usuario "CNS_VERITY" y contraseña "Verity.2025"

  @pago_proveedores
  Scenario: Configurar el acceso a SAP
  Given Se ingresa a SAP
  When Se logea con el usuario "CNS_VERITY" y contraseña "Verity.2025"
  Then Se ingresa a la transaccion "f110"
  And Se ingresan fecha e identificadores de txn
  And Se configura proceso previo de pago
  And Se selecciona el documento a pagar
  And Se configura via de pago
  And Se ejecuta el pago
  And Se cierra sap
  And Se genera el reporte final





