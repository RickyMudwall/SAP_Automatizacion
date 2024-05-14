#noinspection CucumberUndefinedStep
Feature: Realizar una acción

  @ZFICA0034
  Scenario: Proceso de contabilización PNCT
    Given se ingresa a SAP
    When se logea con el usuario "CNS_VERITY" y contraseña "Verity.2025"
    Then se ingresa a la transaccion "ZFICA0034"
    And se ingresan los datos para inscripcion del alumno
      | sociedad	| int_comercial	| clasific_inscripcion	| tp_objeto | id_objeto | monto_descuento	| año_academico | periodo_academico |
      | IP01		| 212719471		| 30					| SE 		| 24002993	| 50000				| 2023			| 400				|
    And se valida la visualizacion de los documentos
      | documentos |  codigo_documento   |
      | COMPENSACION |  180 |
      | DESCUENTO | 290  |
      | ARANCEL | 590  |
    And se cierra sap

  @FPL9 @Regresion
  Scenario: Ver diferido
    Given se ingresa a SAP
    When se logea con el usuario "CNS_VERITY" y contraseña "Verity.2025"
    Then se ingresa a la transaccion "FPL9"
    And se busca el rut "212971057"
    And se selecciona el documento "590000186588"
    And se valida el despliegue de informacion del diferido
      | fecha_pago |
      | 06.10.2023 |
      | 31.10.2023 |
      | 30.11.2023 |
      | 31.12.2023 |
      | 31.01.2024 |
      | 29.02.2024 |
      | 31.03.2024 |
      | 30.04.2024 |
      | 31.05.2024 |
      | 30.06.2024 |
      | 31.07.2024 |
    And se cierra sap

  @ZFICAFPCJ @Regresion
  Scenario: Flujo Caja
    Given se ingresa a SAP
    When se logea con el usuario "CNS_VERITY" y contraseña "Verity.2025"
    Then se ingresa a la transaccion "ZFICAFPCJ"
    And se ingresan los datos para caja de pago rut "212719471" y division "3000"
    And se selecciona documento a pagar
      | documentos |
      | 590000186592 |
      | 590000186582 |
      | 590000186581 |
      | 590000186580 |
      | 590000186579 |
      | 590000186578 |
    And se contabiliza el documento con medio de pago "EFECTIVO/VALE VISTA" y monto "650000"
    And se valida el log del proceso de pago
    And se cierra sap

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

  @prueba_txn
  Scenario: Flujo Caja
  Given se ingresa a SAP
  When se logea con el usuario "CNS_VERITY" y contraseña "Verity.2025"
  Then se ingresa a la transaccion "ZFICAFPCJ"
  And se ingresan los datos para caja de pago rut "212719471" y division "3000"
