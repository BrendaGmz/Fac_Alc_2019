Attribute VB_Name = "Mod_CFDs"
Option Explicit
'Estructura de datos Generales del comprobante digital
Private Type Generales
    Fecha As String
    Fecha_Timbrado As String
    Fecha_Vencimiento As String
    Folio As String
    Forma_Pago As String
    Forma_Pago_Credito_Contado As String
    Condiciones_Pago As String
    Descuento As Double
    SubTotal As Double
    Total As Double
    Importe_Letra As String
    Cadena_Original As String
    Tipo_Comprobante As String
    No_Aprobacion As String
    Año_Aprobacion As Double
    No_Certificado As String
    Certificado As String
    Tipo_Moneda As String
    Tipo_Cambio As Double
    Sello As String
    Serie As String
    Version As String
    PUR As String
    AAB As String
    ZZZ As String
    Total_Prefactura As Double
    Tasa_Impuestos As Double
    Impuestos As Double
    Cantidad_Total_Partidas As Double
    Conteo_Partidas_Pedidos As Double
    Orden_Compra As String
    IVA_Desglosado As Boolean
    Tipo_Documento As String
    Tipo_Factura As String
    Imagen_BMP As String
    Metodo_Pago As String
    Cuenta_Pago As String
    Factura_ID As String
    No_Salida As String
    Relacionado As String
    Tipo_Relacion As String
    UUID_Relacion As String
    Uso_CFDI As String
    Plazo As String
End Type


'Estructura de datos del Emisor del comprobante digital
Private Type Emisor
    RFC As String
    Nombre As String
    Calle As String
    No_Exterior As String
    No_Interior As String
    Colonia As String
    Localidad As String
    Referencia As String
    Municipio As String
    Estado  As String
    Pais As String
    cp As String
    Expedido_En As String
    GLN As String
    Codigo_Proveedor As String
    Regimen_Fiscal As String
End Type

'Estrucutura de datos del ExpedidoEn
Private Type ExpedidoEn
    Calle As String
    No_Exterior As String
    No_Interior As String
    Colonia As String
    Ciudad As String
    Estado As String
    Pais As String
    cp As String
    ExpedidoEn As String
End Type

'Estrucutura de datos del Receptor del comprobante digital
Private Type Receptor
    RFC As String
    Nombre As String
    Calle As String
    No_Exterior As String
    No_Interior As String
    Colonia As String
    Localidad As String
    Referencia As String
    Municipio As String
    Estado  As String
    Pais As String
    cp As String
    Addenda As Boolean
    Formato_Addenda As String
    GLN As String
    Ship_To As String
    Buyer As String
    Condiciones_Pago As String
    Referencia_OC As String
    Referencia_OC_Fecha As String
    Nombre_Cliente As String
    Direccion_Cliente As String
    Ciudad_Edo_Cliente As String
    CP_Cliente As String
    Dias_Pago As Integer
    Tipo_Referencia_Adicional As String
    Referencia_Adicional As String
    Uso_CFDI As String
End Type

'Estructura de datos de las partidas o conceptos del comprobante digital
Private Type Conceptos
    Cantidad As Double
    Unidad As String
    No_Identificacion As String
    No_Identificacion_Codigo As String
    Descripcion As String
    Valor_Unitario As Double
    Importe As Double
    GTIN As String
    Descripcion_PreFactura As String
    Unidad_PreFactura As String
    Precio_PreFactura As String
    Orden_Compra As String
    No_Factura_Interno As String
    Linea_Factura_Interno As Double
    Planta_Receptora As String
    No_Remision As String
    No_Release As String
    No_Receipt As String
    Tasa_Impuesto As Double
    Impuesto As Double
    SubTotal_PreFactura As Double
    Aplica_IVA As String
    Cod_prod As String
    Unidad_Medida As String
    IVA_Producto As Boolean
    Posicion_OC As Long
End Type

'Estructura de datos de los impuestos generados por el comprobante digital
Private Type Impuestos
    Impuesto As String
    Tasa As String
    Importe As Double
End Type

Private Type Retenciones
    Impuesto As String
    Tasa As String
    Importe As Double
End Type

Private Type Locales
    Impuesto As String
    Tasa As String
    Importe As Double
    Tipo As String
End Type

Private Type Pagos_DR
    Fecha_Pago As String
    Forma_Pago_Pagos As String
    Moneda_Pago As String
    Tipo_Cambio_Pago As String
    Monto As Double
    Num_Operacion As String
    RfcEmisorCtaOrd As String
    NomBancoOrdExt As String
    CtaOrdenante As String
    RfcEmisorCtaBen As String
    CtaBeneficiario As String
    TipoCadPago As String
    CertPago As String
    CadPago As String
    Sello_Pago As String
    ID_Doc As String
    Serie As String
    Folio As String
    Moneda_DR As String
    Tipo_Cambio_DR As String
    Metodo_Pago_DR As String
    No_Parcialidad As String
    Saldo_Anterior As Double
    Importe_Pagado As Double
    Saldo_Insoluto As Double
End Type

Private Type Pagos
    Fecha_Pago As String
    Forma_Pago_Pagos As String
    Moneda_Pago As String
    Tipo_Cambio_Pago As String
    Monto As Double
    Num_Operacion As String
    RfcEmisorCtaOrd As String
    NomBancoOrdExt As String
    CtaOrdenante As String
    RfcEmisorCtaBen As String
    CtaBeneficiario As String
    TipoCadPago As String
    CertPago As String
    CadPago As String
    Sello_Pago As String
    ID_Doc As String
    Serie As String
    Folio As String
    Moneda_DR As String
    Tipo_Cambio_DR As String
    Metodo_Pago_DR As String
    No_Parcialidad As String
    Saldo_Anterior As Double
    Importe_Pagado As Double
    Saldo_Insoluto As Double
End Type

Private Type Relacionados
    Existe As Boolean
    Relacionados As String
    UUID_Relacionados As String
End Type

Private Type Relacionados_concepto
    Folio As Boolean
    Serie As String
    UUID As String
End Type

'Variables privadas de la informacion del CFD
Public CFD_Documento As DOMDocument
Public CFD_Generales As Generales
Public CFD_Emisor As Emisor
Public CFD_ExpedidoEn As ExpedidoEn
Public CFD_Receptor As Receptor
Public CFD_Conceptos() As Conceptos
Public CFD_Impuestos() As Impuestos
Public CFD_Impuestos_Retenidos() As Retenciones
Public CFD_Impuestos_Locales() As Locales
Public cp As String
Public IVA_EXENTO As Boolean
Public CFD_Pagos As Pagos
Public CFD_Pagos_DR() As Pagos_DR
Public CFD_Relacionados As Relacionados
Public CFD_Relacionados_Conceptos() As Relacionados_concepto

'Variable del nodo principal
Public Nodo_Principal As IXMLDOMElement

Private Conexion_Web_Services As DOMDocument

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: CFD_Addenda_AMC7_1
'DESCRIPCIÓN: Crea la addenda de acuerdo al formato dado
'PARÁMETROS:
'CREO: Ismael Prieto Sánchez
'FECHA_CREO: 09/Sep/2009 4:10pm
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Private Sub CFD_Addenda_AMC7_1()
Dim Nodo_Addenda As IXMLDOMElement
Dim Nodo_requestForPayment As IXMLDOMElement
Dim Nodo_requestForPaymentIdentification As IXMLDOMElement
Dim Nodo_specialInstruction As IXMLDOMElement
Dim Nodo_orderIdentification As IXMLDOMElement
Dim Nodo_AdditionalInformation As IXMLDOMElement
Dim Nodo_buyer As IXMLDOMElement
Dim Nodo_seller As IXMLDOMElement
Dim Nodo_shipTo As IXMLDOMElement
Dim Nodo_InvoiceCreator As IXMLDOMElement
Dim Nodo_currency As IXMLDOMElement
Dim Nodo_paymentTerms As IXMLDOMElement
Dim Nodo_netPayment As IXMLDOMElement
Dim Nodo_paymentTimePeriod As IXMLDOMElement
Dim Nodo_lineItem As IXMLDOMElement
Dim Nodo_totalAmount As IXMLDOMElement
Dim Nodo_baseAmount As IXMLDOMElement
Dim Nodo_tax As IXMLDOMElement
Dim Nodo_payableAmount As IXMLDOMElement
Dim Nodo_Elemento As IXMLDOMElement
Dim Nodo_Elemento1 As IXMLDOMElement
Dim Nodo_Elemento2 As IXMLDOMElement
Dim Nodo_Elemento3 As IXMLDOMElement
Dim Conteo_Estructura As Integer
    
    '*****************************************************************************************
    'ENCABEZADO DE LA ADDENDA
    '*****************************************************************************************
    'CREA NODO ADDENDA
    '*****************************************************************************************
    'Agrega el elemento de la Addenda
    Set Nodo_Addenda = CFD_Documento.createElement("Addenda")
    'Agrega un salto de linea
    Nodo_Addenda.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    '*****************************************************************************************
    'CREA NODO requestForPayment
    '*****************************************************************************************
    'Agrega el elemento requestForPayment
    Set Nodo_requestForPayment = CFD_Documento.createElement("requestForPayment")
    'Agrega un salto de linea
    Nodo_requestForPayment.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega los atributos al nodo requestForPayment
    Nodo_requestForPayment.setAttribute "type", "SimpleInvoiceType"
    Nodo_requestForPayment.setAttribute "contentVersion", "1.3.1"
    Nodo_requestForPayment.setAttribute "documentStructureVersion", "AMC7.1"
    Nodo_requestForPayment.setAttribute "documentStatus", "ORIGINAL"
    Nodo_requestForPayment.setAttribute "DeliveryDate", Format(CDate(Mid(CFD_Generales.Fecha, 1, 10)), "yyyy-MM-dd")
    
    'Asigna el nodo requestForPayment al nodo Addenda
    Nodo_Addenda.appendChild Nodo_requestForPayment
    
    '*****************************************************************************************
    'CREA NODO requestForPaymentIdentification
    '*****************************************************************************************
    'Agrega el elemento requestForPaymentIdentification
    Set Nodo_requestForPaymentIdentification = CFD_Documento.createElement("requestForPaymentIdentification")
    'Agrega un salto de linea
    Nodo_requestForPaymentIdentification.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega el elemento entityType
    Set Nodo_Elemento = CFD_Documento.createElement("entityType")
    Nodo_Elemento.nodeTypedValue = "INVOICE"
    Nodo_requestForPaymentIdentification.appendChild Nodo_Elemento
    Set Nodo_Elemento = Nothing
    'Agrega un salto de linea
    Nodo_requestForPaymentIdentification.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega el elemento uniqueCreatorIdentification
    Set Nodo_Elemento = CFD_Documento.createElement("uniqueCreatorIdentification")
    Nodo_Elemento.nodeTypedValue = Trim(CFD_Generales.Serie & CFD_Generales.Folio)
    Nodo_requestForPaymentIdentification.appendChild Nodo_Elemento
    Set Nodo_Elemento = Nothing
    'Agrega un salto de linea
    Nodo_requestForPaymentIdentification.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Asigna el nodo requestForPaymentIdentification al nodo requestForPayment
    Nodo_requestForPayment.appendChild Nodo_requestForPaymentIdentification
    
    'Agrega un salto de linea
    Nodo_requestForPayment.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    '*****************************************************************************************
    'CREA NODO specialInstruction
    '*****************************************************************************************
    'Agrega el elemento specialInstruction PUR
    Set Nodo_specialInstruction = CFD_Documento.createElement("specialInstruction")
    'Agrega un salto de linea
    Nodo_specialInstruction.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega los atributos al nodo specialInstruction
    Nodo_specialInstruction.setAttribute "code", "PUR"

    'Agrega el elemento text
    Set Nodo_Elemento = CFD_Documento.createElement("text")
    Nodo_Elemento.nodeTypedValue = Trim(CFD_Generales.PUR)
    Nodo_specialInstruction.appendChild Nodo_Elemento
    Set Nodo_Elemento = Nothing
    'Agrega un salto de linea
    Nodo_specialInstruction.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Asgna el nodo specialInstruction al nodo requestForPayment
    Nodo_requestForPayment.appendChild Nodo_specialInstruction
    'Agrega un salto de linea
    Nodo_requestForPayment.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Seta el nodo a vacio
    Set Nodo_specialInstruction = Nothing
    
    'Agrega el elemento specialInstruction AAB
    Set Nodo_specialInstruction = CFD_Documento.createElement("specialInstruction")
    'Agrega un salto de linea
    Nodo_specialInstruction.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega los atributos al nodo specialInstruction
    Nodo_specialInstruction.setAttribute "code", "AAB"

    'Agrega el elemento text
    Set Nodo_Elemento = CFD_Documento.createElement("text")
    Nodo_Elemento.nodeTypedValue = Trim(CFD_Generales.AAB)
    Nodo_specialInstruction.appendChild Nodo_Elemento
    Set Nodo_Elemento = Nothing
    'Agrega un salto de linea
    Nodo_specialInstruction.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Asgna el nodo specialInstruction al nodo requestForPayment
    Nodo_requestForPayment.appendChild Nodo_specialInstruction
    'Agrega un salto de linea
    Nodo_requestForPayment.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Seta el nodo a vacio
    Set Nodo_specialInstruction = Nothing

    'Agrega el elemento specialInstruction ZZZ
    Set Nodo_specialInstruction = CFD_Documento.createElement("specialInstruction")
    'Agrega un salto de linea
    Nodo_specialInstruction.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega los atributos al nodo specialInstruction
    Nodo_specialInstruction.setAttribute "code", "ZZZ"

    'Agrega el elemento text
    Set Nodo_Elemento = CFD_Documento.createElement("text")
    Nodo_Elemento.nodeTypedValue = Trim(CFD_Generales.ZZZ)
    Nodo_specialInstruction.appendChild Nodo_Elemento
    Set Nodo_Elemento = Nothing
    'Agrega un salto de linea
    Nodo_specialInstruction.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Asgna el nodo specialInstruction al nodo requestForPayment
    Nodo_requestForPayment.appendChild Nodo_specialInstruction
    'Agrega un salto de linea
    Nodo_requestForPayment.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Seta el nodo a vacio
    Set Nodo_specialInstruction = Nothing
    
    '*****************************************************************************************
    'CREA NODO orderIdentification
    '*****************************************************************************************
    'Agrega el elemento orderIdentification
    Set Nodo_orderIdentification = CFD_Documento.createElement("orderIdentification")
    'Agrega un salto de linea
    Nodo_orderIdentification.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega el elemento referenceIdentification
    Set Nodo_Elemento = CFD_Documento.createElement("referenceIdentification")
    Nodo_Elemento.setAttribute "type", "ON"
    Nodo_Elemento.nodeTypedValue = Trim(CFD_Receptor.Referencia_OC)
    Nodo_orderIdentification.appendChild Nodo_Elemento
    Set Nodo_Elemento = Nothing
    'Agrega un salto de linea
    Nodo_orderIdentification.appendChild CFD_Documento.createTextNode(vbCrLf)
   
    'Agrega el elemento ReferenceDate
    Set Nodo_Elemento = CFD_Documento.createElement("ReferenceDate")
    Nodo_Elemento.nodeTypedValue = Format(Trim(CFD_Receptor.Referencia_OC_Fecha), "yyyy-MM-dd")
    Nodo_orderIdentification.appendChild Nodo_Elemento
    Set Nodo_Elemento = Nothing
    'Agrega un salto de linea
    Nodo_orderIdentification.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Asgna el nodo orderIdentification al nodo requestForPayment
    Nodo_requestForPayment.appendChild Nodo_orderIdentification
    'Agrega un salto de linea
    Nodo_requestForPayment.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    '*****************************************************************************************
    'CREA NODO AdditionalInformation
    '*****************************************************************************************
    'Agrega el elemento AdditionalInformation
    Set Nodo_AdditionalInformation = CFD_Documento.createElement("AdditionalInformation")
    'Agrega un salto de linea
    Nodo_AdditionalInformation.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega el elemento referenceIdentification
    Set Nodo_Elemento = CFD_Documento.createElement("referenceIdentification")
    Nodo_Elemento.setAttribute "type", CFD_Receptor.Tipo_Referencia_Adicional
    Nodo_Elemento.nodeTypedValue = Trim(CFD_Receptor.Referencia_Adicional)
    Nodo_AdditionalInformation.appendChild Nodo_Elemento
    Set Nodo_Elemento = Nothing
    'Agrega un salto de linea
    Nodo_AdditionalInformation.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Asgna el nodo orderIdentification al nodo requestForPayment
    Nodo_requestForPayment.appendChild Nodo_AdditionalInformation
    'Agrega un salto de linea
    Nodo_requestForPayment.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    '*****************************************************************************************
    'CREA NODO buyer
    '*****************************************************************************************
    'Agrega el elemento buyer
    Set Nodo_buyer = CFD_Documento.createElement("buyer")
    'Agrega un salto de linea
    Nodo_buyer.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega el elemento gln
    Set Nodo_Elemento = CFD_Documento.createElement("gln")
    Nodo_Elemento.nodeTypedValue = Trim(CFD_Receptor.GLN)
    Nodo_buyer.appendChild Nodo_Elemento
    Set Nodo_Elemento = Nothing
    'Agrega un salto de linea
    Nodo_buyer.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Asgna el nodo buyer al nodo requestForPayment
    Nodo_requestForPayment.appendChild Nodo_buyer
    'Agrega un salto de linea
    Nodo_requestForPayment.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    '*****************************************************************************************
    'CREA NODO seller
    '*****************************************************************************************
    'Agrega el elemento seller
    Set Nodo_seller = CFD_Documento.createElement("seller")
    'Agrega un salto de linea
    Nodo_seller.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega el elemento gln
    Set Nodo_Elemento = CFD_Documento.createElement("gln")
    Nodo_Elemento.nodeTypedValue = Trim(CFD_Emisor.GLN)
    Nodo_seller.appendChild Nodo_Elemento
    Set Nodo_Elemento = Nothing
    'Agrega un salto de linea
    Nodo_seller.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega el elemento alternatePartyIdentification
    Set Nodo_Elemento = CFD_Documento.createElement("alternatePartyIdentification")
    Nodo_Elemento.setAttribute "type", "SELLER_ASSIGNED_IDENTIFIER_FOR_A_PARTY"
    Nodo_Elemento.nodeTypedValue = Trim(CFD_Emisor.Codigo_Proveedor)
    Nodo_seller.appendChild Nodo_Elemento
    Set Nodo_Elemento = Nothing
    'Agrega un salto de linea
    Nodo_seller.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Asgna el nodo seller al nodo requestForPayment
    Nodo_requestForPayment.appendChild Nodo_seller
    'Agrega un salto de linea
    Nodo_requestForPayment.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    '*****************************************************************************************
    'CREA NODO shipTo
    '*****************************************************************************************
    'Agrega el elemento shipTo
    Set Nodo_shipTo = CFD_Documento.createElement("shipTo")
    'Agrega un salto de linea
    Nodo_shipTo.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega el elemento gln
    Set Nodo_Elemento = CFD_Documento.createElement("gln")
    Nodo_Elemento.nodeTypedValue = Trim(CFD_Receptor.GLN)
    Nodo_shipTo.appendChild Nodo_Elemento
    Set Nodo_Elemento = Nothing
    'Agrega un salto de linea
    Nodo_shipTo.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega el elemento nameAndAddress
    Set Nodo_Elemento = CFD_Documento.createElement("nameAndAddress")
    'Agrega un salto de linea
    Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega el elmento hijo name
    Set Nodo_Elemento1 = CFD_Documento.createElement("name")
    Nodo_Elemento1.nodeTypedValue = Trim(CFD_Receptor.Nombre_Cliente)
    Nodo_Elemento.appendChild Nodo_Elemento1
    Set Nodo_Elemento1 = Nothing
    'Agrega un salto de linea
    Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega el elmento hijo streetAddressOne
    Set Nodo_Elemento1 = CFD_Documento.createElement("streetAddressOne")
    Nodo_Elemento1.nodeTypedValue = Trim(CFD_Receptor.Direccion_Cliente)
    Nodo_Elemento.appendChild Nodo_Elemento1
    Set Nodo_Elemento1 = Nothing
    'Agrega un salto de linea
    Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega el elmento hijo city
    Set Nodo_Elemento1 = CFD_Documento.createElement("city")
    Nodo_Elemento1.nodeTypedValue = Trim(CFD_Receptor.Ciudad_Edo_Cliente)
    Nodo_Elemento.appendChild Nodo_Elemento1
    Set Nodo_Elemento1 = Nothing
    'Agrega un salto de linea
    Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega el elmento hijo postalCode
    Set Nodo_Elemento1 = CFD_Documento.createElement("postalCode")
    Nodo_Elemento1.nodeTypedValue = Trim(CFD_Receptor.CP_Cliente)
    Nodo_Elemento.appendChild Nodo_Elemento1
    Set Nodo_Elemento1 = Nothing
    'Agrega un salto de linea
    Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Asigna el nodo nameAndAddress al nodo shipTo
    Nodo_shipTo.appendChild Nodo_Elemento
    Set Nodo_Elemento = Nothing
    'Agrega un salto de linea
    Nodo_shipTo.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Asgna el nodo shipTo al nodo requestForPayment
    Nodo_requestForPayment.appendChild Nodo_shipTo
    'Agrega un salto de linea
    Nodo_requestForPayment.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    '*****************************************************************************************
    'CREA NODO InvoiceCreator
    '*****************************************************************************************
    'Agrega el elemento InvoiceCreator
    Set Nodo_InvoiceCreator = CFD_Documento.createElement("InvoiceCreator")
    'Agrega un salto de linea
    Nodo_InvoiceCreator.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega el elemento gln
    Set Nodo_Elemento = CFD_Documento.createElement("gln")
    Nodo_Elemento.nodeTypedValue = Trim(CFD_Emisor.GLN)
    Nodo_InvoiceCreator.appendChild Nodo_Elemento
    Set Nodo_Elemento = Nothing
    'Agrega un salto de linea
    Nodo_InvoiceCreator.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega el elemento alternatePartyIdentification
    Set Nodo_Elemento = CFD_Documento.createElement("alternatePartyIdentification")
    Nodo_Elemento.setAttribute "type", "VA"
    Nodo_Elemento.nodeTypedValue = Trim(CFD_Emisor.RFC)
    Nodo_InvoiceCreator.appendChild Nodo_Elemento
    Set Nodo_Elemento = Nothing
    'Agrega un salto de linea
    Nodo_InvoiceCreator.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega el elemento nameAndAddress
    Set Nodo_Elemento = CFD_Documento.createElement("nameAndAddress")
    'Agrega un salto de linea
    Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega el elmento hijo name
    Set Nodo_Elemento1 = CFD_Documento.createElement("name")
    Nodo_Elemento1.nodeTypedValue = Trim(CFD_Emisor.Nombre)
    Nodo_Elemento.appendChild Nodo_Elemento1
    Set Nodo_Elemento1 = Nothing
    'Agrega un salto de linea
    Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega el elmento hijo streetAddressOne
    Set Nodo_Elemento1 = CFD_Documento.createElement("streetAddressOne")
    Nodo_Elemento1.nodeTypedValue = Trim(CFD_Emisor.Calle) & " No. " & CFD_Emisor.No_Exterior & " Col. " & CFD_Emisor.Colonia
    Nodo_Elemento.appendChild Nodo_Elemento1
    Set Nodo_Elemento1 = Nothing
    'Agrega un salto de linea
    Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega el elmento hijo city
    Set Nodo_Elemento1 = CFD_Documento.createElement("city")
    Nodo_Elemento1.nodeTypedValue = Trim(CFD_Emisor.Localidad & "," & CFD_Emisor.Estado)
    Nodo_Elemento.appendChild Nodo_Elemento1
    Set Nodo_Elemento1 = Nothing
    'Agrega un salto de linea
    Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega el elmento hijo postalCode
    Set Nodo_Elemento1 = CFD_Documento.createElement("postalCode")
    Nodo_Elemento1.nodeTypedValue = Trim(CFD_Emisor.cp)
    Nodo_Elemento.appendChild Nodo_Elemento1
    Set Nodo_Elemento1 = Nothing
    'Agrega un salto de linea
    Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Asigna el nodo nameAndAddress al nodo InvoiceCreator
    Nodo_InvoiceCreator.appendChild Nodo_Elemento
    Set Nodo_Elemento = Nothing
    'Agrega un salto de linea
    Nodo_InvoiceCreator.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Asgna el nodo InvoiceCreator al nodo requestForPayment
    Nodo_requestForPayment.appendChild Nodo_InvoiceCreator
    'Agrega un salto de linea
    Nodo_requestForPayment.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    '*****************************************************************************************
    'CREA NODO currency
    '*****************************************************************************************
    'Agrega el elemento currency
    Set Nodo_currency = CFD_Documento.createElement("currency")
    'Agrega atributos
    Nodo_currency.setAttribute "currencyISOCode", "MXP"
    'Agrega un salto de linea
    Nodo_currency.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega el elemento currencyFunction
    Set Nodo_Elemento = CFD_Documento.createElement("currencyFunction")
    Nodo_Elemento.nodeTypedValue = "BILLING_CURRENCY"
    Nodo_currency.appendChild Nodo_Elemento
    Set Nodo_Elemento = Nothing
    'Agrega un salto de linea
    Nodo_currency.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega el elemento rateOfChange
    Set Nodo_Elemento = CFD_Documento.createElement("rateOfChange")
    Nodo_Elemento.nodeTypedValue = "1"
    Nodo_currency.appendChild Nodo_Elemento
    Set Nodo_Elemento = Nothing
    'Agrega un salto de linea
    Nodo_currency.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Asgna el nodo currency al nodo requestForPayment
    Nodo_requestForPayment.appendChild Nodo_currency
    'Agrega un salto de linea
    Nodo_requestForPayment.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    '*****************************************************************************************
    'CREA NODO paymentTerms
    '*****************************************************************************************
    'Agrega el elemento paymentTerms
    Set Nodo_paymentTerms = CFD_Documento.createElement("paymentTerms")
    Nodo_paymentTerms.setAttribute "paymentTermsEvent", "DATE_OF_INVOICE" 'Format(CDate(Mid(CFD_Generales.Fecha, 1, 10)), "yyyyMMdd")
    Dim Fecha_Pago As Date
    Fecha_Pago = DateAdd("d", CFD_Receptor.Dias_Pago, CDate(Mid(CFD_Generales.Fecha, 1, 10)))
    Nodo_paymentTerms.setAttribute "PaymentTermsRelationTime", "REFERENCE_AFTER" 'Format(Fecha_Pago, "yyyyMMdd")
    'Agrega un salto de linea
    Nodo_paymentTerms.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega el elemento netPayment
    Set Nodo_Elemento = CFD_Documento.createElement("netPayment")
    Nodo_Elemento.setAttribute "netPaymentTermsType", "BASIC_NET"
    'Agrega un salto de linea
    Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
        'Agrega el elemento paymentTimePeriod
        Set Nodo_Elemento1 = CFD_Documento.createElement("paymentTimePeriod")
        'Agrega un salto de linea
        Nodo_Elemento1.appendChild CFD_Documento.createTextNode(vbCrLf)
            'Agrega el elemento timePeriodDue
            Set Nodo_Elemento2 = CFD_Documento.createElement("timePeriodDue")
            Nodo_Elemento2.setAttribute "timePeriod", "DAYS"
            'Agrega un salto de linea
            Nodo_Elemento2.appendChild CFD_Documento.createTextNode(vbCrLf)
                'Agrega el elemento value
                Set Nodo_Elemento3 = CFD_Documento.createElement("value")
                Nodo_Elemento3.nodeTypedValue = CFD_Receptor.Dias_Pago
                'Asgna el nodo
                Nodo_Elemento2.appendChild Nodo_Elemento3
                Set Nodo_Elemento3 = Nothing
                'Agrega un salto de linea
                Nodo_Elemento2.appendChild CFD_Documento.createTextNode(vbCrLf)
        'Asgna el nodo
        Nodo_Elemento1.appendChild Nodo_Elemento2
        Set Nodo_Elemento2 = Nothing
        'Agrega un salto de linea
        Nodo_Elemento1.appendChild CFD_Documento.createTextNode(vbCrLf)
    'Asgna el nodo
    Nodo_Elemento.appendChild Nodo_Elemento1
    Set Nodo_Elemento1 = Nothing
    'Agrega un salto de linea
    Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
    'Asgna el nodo netPayment al nodo paymentTerms
    Nodo_paymentTerms.appendChild Nodo_Elemento
    Set Nodo_Elemento = Nothing
    'Asgna el nodo paymentTerms al nodo requestForPayment
    Nodo_requestForPayment.appendChild Nodo_paymentTerms
    'Agrega un salto de linea
    Nodo_paymentTerms.appendChild CFD_Documento.createTextNode(vbCrLf)
    'Agrega un salto de linea
    Nodo_requestForPayment.appendChild CFD_Documento.createTextNode(vbCrLf)
    
'''    'Agrega el elemento paymentTermsEvent
'''    Set Nodo_Elemento = CFD_Documento.createElement("paymentTermsEvent")
'''    Nodo_Elemento.setAttribute "type", "DATE_OF_INVOICE"
'''    Nodo_Elemento.nodeTypedValue = Format(CDate(Mid(CFD_Generales.Fecha, 1, 10)), "yyyyMMdd")
'''    Nodo_paymentTerms.appendChild Nodo_Elemento
'''    Set Nodo_Elemento = Nothing
'''    'Agrega un salto de linea
'''    Nodo_paymentTerms.appendChild CFD_Documento.createTextNode(vbCrLf)
'''
'''    'Agrega el elemento paymentTermsEvent
'''    Set Nodo_Elemento = CFD_Documento.createElement("PaymentTermsRelationTime")
'''    Nodo_Elemento.setAttribute "type", "REFERENCE_AFTER"
'''    'Dim Fecha_Pago As Date
'''    Fecha_Pago = DateAdd("d", CFD_Receptor.Dias_Pago, CDate(Mid(CFD_Generales.Fecha, 1, 10)))
'''    Nodo_Elemento.nodeTypedValue = Format(Fecha_Pago, "yyyyMMdd")
'''    Nodo_paymentTerms.appendChild Nodo_Elemento
'''    Set Nodo_Elemento = Nothing
'''    'Agrega un salto de linea
'''    Nodo_paymentTerms.appendChild CFD_Documento.createTextNode(vbCrLf)
'''
'''    'Asgna el nodo paymentTerms al nodo requestForPayment
'''    Nodo_requestForPayment.appendChild Nodo_paymentTerms
'''    'Agrega un salto de linea
'''    Nodo_requestForPayment.appendChild CFD_Documento.createTextNode(vbCrLf)
    
'''''    '*****************************************************************************************
'''''    'CREA NODO netPayment
'''''    '*****************************************************************************************
'''''    'Agrega el elemento netPayment
'''''    Set Nodo_netPayment = CFD_Documento.createElement("netPayment")
'''''    Nodo_netPayment.setAttribute "netPaymentTermsType", "BASIC_NET"
'''''    'Agrega un salto de linea
'''''    Nodo_netPayment.appendChild CFD_Documento.createTextNode(vbCrLf)
'''''
'''''    'Agrega el elemento netPaymentTermsType
'''''    Set Nodo_Elemento = CFD_Documento.createElement("netPaymentTermsType")
'''''    Nodo_Elemento.nodeTypedValue = "BASIC_NET"
'''''    Nodo_netPayment.appendChild Nodo_Elemento
'''''    Set Nodo_Elemento = Nothing
'''''    'Agrega un salto de linea
'''''    Nodo_netPayment.appendChild CFD_Documento.createTextNode(vbCrLf)
'''''
'''''    'Asgna el nodo netPayment al nodo requestForPayment
'''''    Nodo_paymentTerms.appendChild Nodo_netPayment
'''''    'Agrega un salto de linea
'''''    Nodo_paymentTerms.appendChild CFD_Documento.createTextNode(vbCrLf)
'''''
'''''    '*****************************************************************************************
'''''    'CREA NODO paymentTimePeriod
'''''    '*****************************************************************************************
'''''    'Agrega el elemento paymentTimePeriod
'''''    Set Nodo_paymentTimePeriod = CFD_Documento.createElement("paymentTimePeriod")
'''''    'Agrega un salto de linea
'''''    Nodo_paymentTimePeriod.appendChild CFD_Documento.createTextNode(vbCrLf)
'''''
'''''    'Asigna los atributos
'''''    Nodo_paymentTimePeriod.setAttribute "timePeriod", "DAYS"
'''''    Nodo_paymentTimePeriod.nodeTypedValue = CFD_Receptor.Dias_Pago
'''''
'''''    'Asgna el nodo paymentTerms al nodo requestForPayment
'''''    Nodo_requestForPayment.appendChild Nodo_paymentTerms
'''''    'Agrega un salto de linea
'''''    Nodo_requestForPayment.appendChild CFD_Documento.createTextNode(vbCrLf)
'''''
'''''    'Asgna el nodo paymentTimePeriod al nodo requestForPayment
'''''    Nodo_requestForPayment.appendChild Nodo_paymentTimePeriod
'''''    'Agrega un salto de linea
'''''    Nodo_requestForPayment.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    '*****************************************************************************************
    'DETALLES DE LA ADDENDA
    '*****************************************************************************************
    'CREA NODO lineItem
    '*****************************************************************************************
    'Agrega los items
    For Conteo_Estructura = 1 To UBound(CFD_Conceptos)
        'Agrega el elemento lineItem
        Set Nodo_lineItem = CFD_Documento.createElement("lineItem")
        'Agrega un salto de linea
        Nodo_lineItem.appendChild CFD_Documento.createTextNode(vbCrLf)
    
        'Asigna los atributos
        Nodo_lineItem.setAttribute "type", "SimpleInvoiceLineItemType"
        Nodo_lineItem.setAttribute "number", Conteo_Estructura
        
        'Agrega el elemento tradeItemIdentification
        Set Nodo_Elemento = CFD_Documento.createElement("tradeItemIdentification")
        'Agrega un salto de linea
        Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
            'Agrega el elemento gtin
            Set Nodo_Elemento1 = CFD_Documento.createElement("gtin")
            Nodo_Elemento1.nodeTypedValue = CFD_Conceptos(Conteo_Estructura).GTIN
        Nodo_Elemento.appendChild Nodo_Elemento1
        'Agrega un salto de linea
        Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
        Nodo_lineItem.appendChild Nodo_Elemento
        'Agrega un salto de linea
        Nodo_lineItem.appendChild CFD_Documento.createTextNode(vbCrLf)
        Set Nodo_Elemento = Nothing
        Set Nodo_Elemento1 = Nothing
        
        'Agrega el elemento alternateTradeItemIdentification
        Set Nodo_Elemento = CFD_Documento.createElement("alternateTradeItemIdentification")
        'Asigna los atributos
        Nodo_Elemento.setAttribute "type", "SUPPLIER_ASSIGNED"
        Nodo_Elemento.nodeTypedValue = CFD_Conceptos(Conteo_Estructura).No_Identificacion_Codigo
        Nodo_lineItem.appendChild Nodo_Elemento
        Set Nodo_Elemento = Nothing
        Set Nodo_Elemento1 = Nothing
        'Agrega un salto de linea
        Nodo_lineItem.appendChild CFD_Documento.createTextNode(vbCrLf)
        
        'Agrega el elemento tradeItemDescriptionInformation
        Set Nodo_Elemento = CFD_Documento.createElement("tradeItemDescriptionInformation")
        'Agrega un salto de linea
        Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
        'Asigna los atributos
        Nodo_Elemento.setAttribute "language", "ES"
            'Agrega el elemento longText
            Set Nodo_Elemento1 = CFD_Documento.createElement("longText")
            Nodo_Elemento1.nodeTypedValue = CFD_Conceptos(Conteo_Estructura).Descripcion_PreFactura
        Nodo_Elemento.appendChild Nodo_Elemento1
        Nodo_lineItem.appendChild Nodo_Elemento
        'Agrega un salto de linea
        Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
        Set Nodo_Elemento = Nothing
        Set Nodo_Elemento1 = Nothing
        'Agrega un salto de linea
        Nodo_lineItem.appendChild CFD_Documento.createTextNode(vbCrLf)
        
        'Agrega el elemento invoicedQuantity
        Set Nodo_Elemento = CFD_Documento.createElement("invoicedQuantity")
        'Asigna los atributos
        Nodo_Elemento.setAttribute "unitOfMeasure", CFD_Conceptos(Conteo_Estructura).Unidad_PreFactura
        Nodo_Elemento.nodeTypedValue = CFD_Conceptos(Conteo_Estructura).Cantidad
        Nodo_lineItem.appendChild Nodo_Elemento
        Set Nodo_Elemento = Nothing
        'Agrega un salto de linea
        Nodo_lineItem.appendChild CFD_Documento.createTextNode(vbCrLf)
        
        'Agrega el elemento grossPrice
        Set Nodo_Elemento = CFD_Documento.createElement("grossPrice")
        'Agrega un salto de linea
        Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
            'Agrega el elemento Amount
            Set Nodo_Elemento1 = CFD_Documento.createElement("Amount")
            Nodo_Elemento1.nodeTypedValue = Format(CFD_Conceptos(Conteo_Estructura).Precio_PreFactura, "#0.00")
        Nodo_Elemento.appendChild Nodo_Elemento1
        'Agrega un salto de linea
        Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
        Nodo_lineItem.appendChild Nodo_Elemento
        Set Nodo_Elemento = Nothing
        Set Nodo_Elemento1 = Nothing
        'Agrega un salto de linea
        Nodo_lineItem.appendChild CFD_Documento.createTextNode(vbCrLf)
        
        'Agrega el elemento netPrice
        Set Nodo_Elemento = CFD_Documento.createElement("netPrice")
        'Agrega un salto de linea
        Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
            'Agrega el elemento Amount
            Set Nodo_Elemento1 = CFD_Documento.createElement("Amount")
            Nodo_Elemento1.nodeTypedValue = Format(CFD_Conceptos(Conteo_Estructura).Valor_Unitario, "#0.00")
        Nodo_Elemento.appendChild Nodo_Elemento1
        'Agrega un salto de linea
        Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
        Nodo_lineItem.appendChild Nodo_Elemento
        Set Nodo_Elemento = Nothing
        Set Nodo_Elemento1 = Nothing
        'Agrega un salto de linea
        Nodo_lineItem.appendChild CFD_Documento.createTextNode(vbCrLf)
        
        'Agrega el elemento AdditionalInformation
        Set Nodo_Elemento = CFD_Documento.createElement("AdditionalInformation")
        If CFD_Conceptos(Conteo_Estructura).Orden_Compra <> "" Then
            'Agrega un salto de linea
            Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
                'Agrega el elemento referenceIdentification ON
                Set Nodo_Elemento1 = CFD_Documento.createElement("referenceIdentification")
                Nodo_Elemento1.setAttribute "type", "ON"
                Nodo_Elemento1.nodeTypedValue = CFD_Conceptos(Conteo_Estructura).Orden_Compra
            Nodo_Elemento.appendChild Nodo_Elemento1
        End If
        If CFD_Conceptos(Conteo_Estructura).No_Factura_Interno <> "" Then
            'Agrega un salto de linea
            Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
            Set Nodo_Elemento1 = Nothing
                'Agrega el elemento referenceIdentification IF
                Set Nodo_Elemento1 = CFD_Documento.createElement("referenceIdentification")
                Nodo_Elemento1.setAttribute "type", "IF"
                Nodo_Elemento1.nodeTypedValue = CFD_Conceptos(Conteo_Estructura).No_Factura_Interno
            Nodo_Elemento.appendChild Nodo_Elemento1
            'Agrega un salto de linea
            Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
            Set Nodo_Elemento1 = Nothing
        End If
        If CFD_Conceptos(Conteo_Estructura).Linea_Factura_Interno <> 0 Then
            'Agrega el elemento referenceIdentification IL
            Set Nodo_Elemento1 = CFD_Documento.createElement("referenceIdentification")
            Nodo_Elemento1.setAttribute "type", "IL"
            Nodo_Elemento1.nodeTypedValue = CFD_Conceptos(Conteo_Estructura).Linea_Factura_Interno
        Nodo_Elemento.appendChild Nodo_Elemento1
        'Agrega un salto de linea
        Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
        Set Nodo_Elemento1 = Nothing
        End If
        If CFD_Conceptos(Conteo_Estructura).Planta_Receptora <> "" Then
            'Agrega el elemento referenceIdentification PL
            Set Nodo_Elemento1 = CFD_Documento.createElement("referenceIdentification")
            Nodo_Elemento1.setAttribute "type", "PL"
            Nodo_Elemento1.nodeTypedValue = CFD_Conceptos(Conteo_Estructura).Planta_Receptora
        Nodo_Elemento.appendChild Nodo_Elemento1
        'Agrega un salto de linea
        Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
        Set Nodo_Elemento1 = Nothing
        End If
        If CFD_Conceptos(Conteo_Estructura).No_Remision <> "" Then
            'Agrega el elemento referenceIdentification RE
            Set Nodo_Elemento1 = CFD_Documento.createElement("referenceIdentification")
            Nodo_Elemento1.setAttribute "type", "RE"
            Nodo_Elemento1.nodeTypedValue = CFD_Conceptos(Conteo_Estructura).No_Remision
        Nodo_Elemento.appendChild Nodo_Elemento1
        'Agrega un salto de linea
        Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
        Set Nodo_Elemento1 = Nothing
        End If
        If CFD_Conceptos(Conteo_Estructura).No_Release <> "" Then
            'Agrega el elemento referenceIdentification RL
            Set Nodo_Elemento1 = CFD_Documento.createElement("referenceIdentification")
            Nodo_Elemento1.setAttribute "type", "RL"
            Nodo_Elemento1.nodeTypedValue = CFD_Conceptos(Conteo_Estructura).No_Release
        Nodo_Elemento.appendChild Nodo_Elemento1
        'Agrega un salto de linea
        Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
        Set Nodo_Elemento1 = Nothing
        End If
        If CFD_Conceptos(Conteo_Estructura).No_Receipt <> "" Then
            'Agrega el elemento referenceIdentification RC
            Set Nodo_Elemento1 = CFD_Documento.createElement("referenceIdentification")
            Nodo_Elemento1.setAttribute "type", "RC"
            Nodo_Elemento1.nodeTypedValue = CFD_Conceptos(Conteo_Estructura).No_Receipt
        Nodo_Elemento.appendChild Nodo_Elemento1
        'Agrega un salto de linea
        Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
        Set Nodo_Elemento1 = Nothing
        End If
        'Asgna el nodo hijo al lineItem
        Nodo_lineItem.appendChild Nodo_Elemento
        'Agrega un salto de linea
        Nodo_lineItem.appendChild CFD_Documento.createTextNode(vbCrLf)
        
        'Agrega el elemento tradeItemTaxInformation
        Set Nodo_Elemento = CFD_Documento.createElement("tradeItemTaxInformation")
        'Agrega un salto de linea
        Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
            'Agrega el elemento taxTypeDescription
            Set Nodo_Elemento1 = CFD_Documento.createElement("taxTypeDescription")
            Nodo_Elemento1.nodeTypedValue = "VAT"
        Nodo_Elemento.appendChild Nodo_Elemento1
        'Agrega un salto de linea
        Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
        Set Nodo_Elemento1 = Nothing
            'Agrega el elemento tradeItemTaxAmount
            Set Nodo_Elemento1 = CFD_Documento.createElement("tradeItemTaxAmount")
            'Agrega un salto de linea
            Nodo_Elemento1.appendChild CFD_Documento.createTextNode(vbCrLf)
                'Agrega el elemento tradeItemTaxAmount
                Set Nodo_Elemento2 = CFD_Documento.createElement("taxPercentage")
                Nodo_Elemento2.nodeTypedValue = CFD_Conceptos(Conteo_Estructura).Tasa_Impuesto
            Nodo_Elemento1.appendChild Nodo_Elemento2
            'Agrega un salto de linea
            Nodo_Elemento1.appendChild CFD_Documento.createTextNode(vbCrLf)
            Set Nodo_Elemento2 = Nothing
                'Agrega el elemento taxAmount
                Set Nodo_Elemento2 = CFD_Documento.createElement("taxAmount")
                Nodo_Elemento2.nodeTypedValue = CFD_Conceptos(Conteo_Estructura).Impuesto
            Nodo_Elemento1.appendChild Nodo_Elemento2
            'Agrega un salto de linea
            Nodo_Elemento1.appendChild CFD_Documento.createTextNode(vbCrLf)
            Set Nodo_Elemento2 = Nothing
        Nodo_Elemento.appendChild Nodo_Elemento1
        'Agrega un salto de linea
        Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
        Set Nodo_Elemento1 = Nothing
            'Agrega el elemento taxCategory
            Set Nodo_Elemento1 = CFD_Documento.createElement("taxCategory")
            Nodo_Elemento1.nodeTypedValue = "TRANSFERIDO"
        Nodo_Elemento.appendChild Nodo_Elemento1
        'Agrega un salto de linea
        Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
        Set Nodo_Elemento1 = Nothing
        'Asgna el nodo hijo al lineItem
        Nodo_lineItem.appendChild Nodo_Elemento
        'Agrega un salto de linea
        Nodo_lineItem.appendChild CFD_Documento.createTextNode(vbCrLf)
            
        'Agrega el elemento totalLineAmount
        Set Nodo_Elemento = CFD_Documento.createElement("totalLineAmount")
        'Agrega un salto de linea
        Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
            'Agrega el elemento grossAmount
            Set Nodo_Elemento1 = CFD_Documento.createElement("grossAmount")
            'Agrega un salto de linea
            Nodo_Elemento1.appendChild CFD_Documento.createTextNode(vbCrLf)
                'Agrega el elemento Amount
                Set Nodo_Elemento2 = CFD_Documento.createElement("Amount")
                Nodo_Elemento2.nodeTypedValue = Format(CFD_Conceptos(Conteo_Estructura).SubTotal_PreFactura, "#0.00")
            Nodo_Elemento1.appendChild Nodo_Elemento2
            'Agrega un salto de linea
            Nodo_Elemento1.appendChild CFD_Documento.createTextNode(vbCrLf)
            Set Nodo_Elemento2 = Nothing
        Nodo_Elemento.appendChild Nodo_Elemento1
        'Agrega un salto de linea
        Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
        Set Nodo_Elemento1 = Nothing
            'Agrega el elemento netAmount
            Set Nodo_Elemento1 = CFD_Documento.createElement("netAmount")
            'Agrega un salto de linea
            Nodo_Elemento1.appendChild CFD_Documento.createTextNode(vbCrLf)
                'Agrega el elemento Amount
                Set Nodo_Elemento2 = CFD_Documento.createElement("Amount")
                Nodo_Elemento2.nodeTypedValue = Format(CFD_Conceptos(Conteo_Estructura).Importe, "#0.00")
            Nodo_Elemento1.appendChild Nodo_Elemento2
            'Agrega un salto de linea
            Nodo_Elemento1.appendChild CFD_Documento.createTextNode(vbCrLf)
            Set Nodo_Elemento2 = Nothing
        Nodo_Elemento.appendChild Nodo_Elemento1
        'Agrega un salto de linea
        Nodo_Elemento.appendChild CFD_Documento.createTextNode(vbCrLf)
        Set Nodo_Elemento1 = Nothing
        'Asgna el nodo hijo al lineItem
        Nodo_lineItem.appendChild Nodo_Elemento
        'Agrega un salto de linea
        Nodo_lineItem.appendChild CFD_Documento.createTextNode(vbCrLf)
        Set Nodo_Elemento = Nothing
        
        'Asgna el nodo lineItem al nodo requestForPayment
        Nodo_requestForPayment.appendChild Nodo_lineItem
        'Agrega un salto de linea
        Nodo_requestForPayment.appendChild CFD_Documento.createTextNode(vbCrLf)
        
        'Limpia el nodo
        Set Nodo_lineItem = Nothing
    Next Conteo_Estructura
    
    '*****************************************************************************************
    'TOTALES DE LA ADDENDA
    '*****************************************************************************************
    'CREA NODO totalAmount
    '*****************************************************************************************
    'Agrega el elemento totalAmount
    Set Nodo_totalAmount = CFD_Documento.createElement("totalAmount")
    'Agrega un salto de linea
    Nodo_totalAmount.appendChild CFD_Documento.createTextNode(vbCrLf)
        'Agrega el elemento Amount
        Set Nodo_Elemento = CFD_Documento.createElement("Amount")
        Nodo_Elemento.nodeTypedValue = Format(CFD_Generales.Total, "#0.00")
        Nodo_totalAmount.appendChild Nodo_Elemento
    Set Nodo_Elemento = Nothing
    'Agrega un salto de linea
    Nodo_totalAmount.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Asgna el nodo totalAmount al nodo requestForPayment
    Nodo_requestForPayment.appendChild Nodo_totalAmount
    'Agrega un salto de linea
    Nodo_requestForPayment.appendChild CFD_Documento.createTextNode(vbCrLf)
        
    '*****************************************************************************************
    'CREA NODO baseAmount
    '*****************************************************************************************
    'Agrega el elemento baseAmount
    Set Nodo_baseAmount = CFD_Documento.createElement("baseAmount")
    'Agrega un salto de linea
    Nodo_baseAmount.appendChild CFD_Documento.createTextNode(vbCrLf)
        'Agrega el elemento Amount
        Set Nodo_Elemento = CFD_Documento.createElement("Amount")
        Nodo_Elemento.nodeTypedValue = Format(CFD_Generales.Total_Prefactura, "#0.00")
        Nodo_baseAmount.appendChild Nodo_Elemento
    Set Nodo_Elemento = Nothing
    'Agrega un salto de linea
    Nodo_baseAmount.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Asgna el nodo baseAmount al nodo requestForPayment
    Nodo_requestForPayment.appendChild Nodo_baseAmount
    'Agrega un salto de linea
    Nodo_requestForPayment.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    '*****************************************************************************************
    'CREA NODO tax
    '*****************************************************************************************
    'Agrega el elemento tax
    Set Nodo_tax = CFD_Documento.createElement("tax")
    'Agrega un salto de linea
    Nodo_tax.appendChild CFD_Documento.createTextNode(vbCrLf)
        'Agrega los atributos
        Nodo_tax.setAttribute "type", "VAT"
            'Agrega el elemento taxPercentage
            Set Nodo_Elemento = CFD_Documento.createElement("taxPercentage")
            Nodo_Elemento.nodeTypedValue = CFD_Generales.Tasa_Impuestos
            Nodo_tax.appendChild Nodo_Elemento
            'Agrega un salto de linea
            Nodo_tax.appendChild CFD_Documento.createTextNode(vbCrLf)
            Set Nodo_Elemento = Nothing
            'Agrega el elemento taxAmount
            Set Nodo_Elemento = CFD_Documento.createElement("taxAmount")
            Nodo_Elemento.nodeTypedValue = Format(CFD_Generales.Impuestos, "#0.00")
            Nodo_tax.appendChild Nodo_Elemento
            'Agrega un salto de linea
            Nodo_tax.appendChild CFD_Documento.createTextNode(vbCrLf)
            Set Nodo_Elemento = Nothing
            'Agrega el elemento taxCategory
            Set Nodo_Elemento = CFD_Documento.createElement("taxCategory")
            Nodo_Elemento.nodeTypedValue = "TRANSFERIDO"
            Nodo_tax.appendChild Nodo_Elemento
            'Agrega un salto de linea
            Nodo_tax.appendChild CFD_Documento.createTextNode(vbCrLf)
            Set Nodo_Elemento = Nothing
    
    'Asgna el nodo tax al nodo requestForPayment
    Nodo_requestForPayment.appendChild Nodo_tax
    'Agrega un salto de linea
    Nodo_requestForPayment.appendChild CFD_Documento.createTextNode(vbCrLf)
            
    '*****************************************************************************************
    'CREA NODO payableAmount
    '*****************************************************************************************
    'Agrega el elemento payableAmount
    Set Nodo_payableAmount = CFD_Documento.createElement("payableAmount")
    'Agrega un salto de linea
    Nodo_payableAmount.appendChild CFD_Documento.createTextNode(vbCrLf)
        'Agrega el elemento taxPercentage
        Set Nodo_Elemento = CFD_Documento.createElement("Amount")
        Nodo_Elemento.nodeTypedValue = Format(CFD_Generales.Total, "#0.00")
        Nodo_payableAmount.appendChild Nodo_Elemento
        Set Nodo_Elemento = Nothing
    'Agrega un salto de linea
    Nodo_payableAmount.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Asgna el nodo payableAmount al nodo requestForPayment
    Nodo_requestForPayment.appendChild Nodo_payableAmount
    'Agrega un salto de linea
    Nodo_requestForPayment.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Agrega un salto de linea
    Nodo_Addenda.appendChild CFD_Documento.createTextNode(vbCrLf)
    
    'Asigna el nodo Addenda al nodo principal
    Nodo_Principal.appendChild Nodo_Addenda
    'Agrega un salto de linea
    Nodo_Principal.appendChild CFD_Documento.createTextNode(vbCrLf)
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: CFD_Cadena_Original
'DESCRIPCIÓN: Crea la cadena original del CFDI de acuerdo a los datos asignados
'PARÁMETROS: Devuelve la cadena compuesta en la función
'CREO:       Sergio Godínez Banda
'FECHA_CREO: 25-Mayo-2012
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function CFD_Cadena_Original_2(Tipo_Factura As String)
Dim Cadena_Original As String
Dim cadena As String
Dim Conteo_Conceptos As Integer
Dim Conteo_Estructura2 As Integer
Dim Conteo_Impuestos As Integer
Dim Grupo_Cadena() As String
Dim Suma_Retenciones As Double
Dim Suma_Locales_Retenciones As Double
Dim Suma_Locales_Trasladados As Double
Dim Retenciones_Locales As String
Dim Trasladados_Locales As String
Dim Suma_Trasladados As Double
Dim i As Long
    'Datos Generales
    Cadena_Original = ""
    Cadena_Original = "||" & CFD_Generales.Version
   ' Cadena_Original = Cadena_Original & "|" & CFD_Generales.Version
    Cadena_Original = Cadena_Original & "|" & CFD_Generales.Serie
    Cadena_Original = Cadena_Original & "|" & CFD_Generales.Folio
    Cadena_Original = Cadena_Original & "|" & CFD_Generales.Fecha
    'Cadena_Original = Cadena_Original & "|" & CFD_Generales.Fecha
    'Cadena_Original = Cadena_Original & "|" & CFD_Generales.Sello
    Cadena_Original = Cadena_Original & "|" & Mid(CFD_Generales.Metodo_Pago, 1, 2)
    Cadena_Original = Cadena_Original & "|" & CFD_Generales.No_Certificado
    'Cadena_Original = Cadena_Original & "|" & CFD_Generales.Condiciones_Pago
    
    'Cadena_Original = Cadena_Original & "|" & CFD_Generales.Tipo_Comprobante
    'Cadena_Original = Cadena_Original & "|" & CFD_Generales.Forma_Pago 'LCase(CFD_Generales.Forma_Pago)
    If CFD_Generales.Condiciones_Pago <> "" Then
        Cadena_Original = Cadena_Original & "|" & CFD_Generales.Condiciones_Pago
    End If
    Cadena_Original = Cadena_Original & "|" & CDbl(Format(CFD_Generales.SubTotal, "#0.00"))
    If CFD_Generales.Descuento > 0 Then
        Cadena_Original = Cadena_Original & "|" & Format(CFD_Generales.Descuento, "#0.00")
    End If
    Cadena_Original = Cadena_Original & "|" & CFD_Generales.Tipo_Moneda
    If CFD_Generales.Tipo_Moneda <> "MXN" And CFD_Generales.Tipo_Moneda <> "XXX" Then
        Cadena_Original = Cadena_Original & "|" & CFD_Generales.Tipo_Cambio
    End If
    'Cadena_Original = Cadena_Original & "|" & CFD_Generales.Tipo_Moneda
    
    Cadena_Original = Cadena_Original & "|" & CDbl(Format(CFD_Generales.Total, "#0.00"))
    Cadena_Original = Cadena_Original & "|" & CFD_Generales.Tipo_Comprobante
    Cadena_Original = Cadena_Original & "|" & CFD_Generales.Forma_Pago
    'Cadena_Original = Cadena_Original & "|" & cp
    'If CFD_Generales.No_Cuenta_Pago <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & CFD_Generales.No_Cuenta_Pago
    'End If
    If CFD_Generales.Relacionado = "S" Then
            Cadena_Original = Cadena_Original & "|" & Mid(CFD_Generales.Tipo_Relacion, 1, 2)
            Cadena_Original = Cadena_Original & "|" & CFD_Generales.UUID_Relacion
    End If
    'Datos del Emisor
    
    Cadena_Original = Cadena_Original & "|" & CFD_Emisor.RFC
    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Emisor.Nombre)
    'Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Emisor.Regimen_Fiscal)
    'If CFD_Emisor.No_Exterior <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & CFD_Emisor.No_Exterior
    'End If
    'If CFD_Emisor.No_Interior <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & CFD_Emisor.No_Interior
    'End If
    'If CFD_Emisor.Colonia <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Emisor.Colonia)
    'End If
    'Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Emisor.Localidad)
    'If CFD_Emisor.Referencia <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Emisor.Referencia)
    'End If
    'Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Emisor.Municipio)
    'Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Emisor.Estado)
    'Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Emisor.Pais)
    'Cadena_Original = Cadena_Original & "|" & CFD_Emisor.CP
    
    'Datos de Expedido En
    'If CFD_ExpedidoEn.Calle <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_ExpedidoEn.Calle)
    'End If
    'If CFD_ExpedidoEn.No_Exterior <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & CFD_ExpedidoEn.No_Exterior
    'End If
    'If CFD_ExpedidoEn.No_Interior <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & CFD_ExpedidoEn.No_Interior
    'End If
    'If CFD_ExpedidoEn.Colonia <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_ExpedidoEn.Colonia)
    'End If
'    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Emisor_Expedido_En.Localidad)
'    If CFD_Emisor_Expedido_En.Referencia <> "" Then
'        Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Emisor_Expedido_En.Referencia)
'    End If
    'If CFD_ExpedidoEn.Ciudad <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_ExpedidoEn.Ciudad)
    'End If
    'If CFD_ExpedidoEn.Estado <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_ExpedidoEn.Estado)
    'End If
    'If CFD_ExpedidoEn.Pais <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_ExpedidoEn.Pais)
    'End If
    'If CFD_ExpedidoEn.CP <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & CFD_ExpedidoEn.CP
    'End If
    'Cadena_Original = Cadena_Original & "|" & Regimen_Fiscal
    'If Regimen_Fiscal_2 <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & Regimen_Fiscal_2
    'End If
    'Datos del Receptor
    
    Cadena_Original = Cadena_Original & "|" & CFD_Receptor.RFC
    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Receptor.Nombre)
    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(Mid(CFD_Receptor.Uso_CFDI, 1, 3))
    
    'Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Receptor.Calle)
    'If CFD_Receptor.No_Exterior <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & CFD_Receptor.No_Exterior
    'End If
    'If CFD_Receptor.No_Interior <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & CFD_Receptor.No_Interior
    'End If
    'If CFD_Receptor.Colonia <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Receptor.Colonia)
    'End If
    'If CFD_Receptor.Localidad <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Receptor.Localidad)
    'End If
    'If CFD_Receptor.Referencia <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales((CFD_Receptor.Referencia))
    'End If
    'Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Receptor.Municipio)
    'Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Receptor.Estado)
    'Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Receptor.Pais)
    'Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Receptor.CP)
    
    'Datos de los Conceptos
    ReDim Valores(UBound(CFD_Conceptos))
   For Conteo_Conceptos = 1 To UBound(CFD_Conceptos)
        Cadena_Original = Cadena_Original & "|" & Format(CFD_Conceptos(Conteo_Conceptos).No_Identificacion)
        If CFD_Conceptos(Conteo_Conceptos).No_Identificacion <> "" Then
            Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Conceptos(Conteo_Conceptos).No_Identificacion)
        End If
        Cadena_Original = Cadena_Original & "|" & Format(CFD_Conceptos(Conteo_Conceptos).Cantidad, "#0.00")
        Cadena_Original = Cadena_Original & "|" & CFD_Conceptos(Conteo_Conceptos).Unidad
        If CFD_Conceptos(Conteo_Conceptos).Unidad <> "" Then
            Cadena_Original = Cadena_Original & "|" & CFD_Conceptos(Conteo_Conceptos).Unidad
        End If
        Cadena_Original = Cadena_Original & "|" & CFD_Conceptos(Conteo_Conceptos).Descripcion
        Cadena_Original = Cadena_Original & "|" & CDbl(Format(CFD_Conceptos(Conteo_Conceptos).Valor_Unitario, "#0.00"))
        Cadena_Original = Cadena_Original & "|" & CDbl(Format(CFD_Conceptos(Conteo_Conceptos).Importe, "#0.00"))
        If CFD_Generales.Descuento > 0 Then
                            Cadena_Original = Cadena_Original & "|" & Format(CFD_Conceptos(Conteo_Conceptos).Importe * CFD_Generales.Descuento, "#0.00")
                        End If
        'If CFD_Conceptos(Conteo_Conceptos).IVA_Producto <> "" And CFD_Conceptos(Conteo_Conceptos).IVA_Producto <> 0 Then
        '
        '    'Cadena_Original = Cadena_Original & "|" & CDbl(Format(CFD_Conceptos(Conteo_Conceptos).IVA_Producto, "#0.00"))
        '    Cadena_Original = Cadena_Original & "|" & CDbl(Format(CFD_Conceptos(Conteo_Conceptos).Importe, "#0.00"))
        '    Cadena_Original = Cadena_Original & "|" & "002"
        '    Cadena_Original = Cadena_Original & "|" & "Tasa"
        '    Cadena_Original = Cadena_Original & "|" & Val(CFD_Generales.Tasa_IVA / 100)
        '    Cadena_Original = Cadena_Original & "|" & CDbl(Format(CFD_Conceptos(Conteo_Conceptos).IVA_Producto, "#0.00"))
            
        '    End If
        'If CFD_Conceptos(Conteo_Conceptos).IEPS_Producto <> "" And CFD_Conceptos(Conteo_Conceptos).IEPS_Producto <> 0 Then
        '    Cadena_Original = Cadena_Original & "|" & CDbl(Format(CFD_Conceptos(Conteo_Conceptos).Importe, "#0.00"))
        '    Cadena_Original = Cadena_Original & "|" & "003"
        '    Cadena_Original = Cadena_Original & "|" & CFD_IEPS(1).Factor
        '    Cadena_Original = Cadena_Original & "|" & Val(CFD_IEPS(1).Tasa_IEPS / 100)
        '    'Cadena_Original = Cadena_Original & "|" & CFD_Conceptos.Impuesto_IEPS
        '    Cadena_Original = Cadena_Original & "|" & CDbl(Format(CFD_IEPS(Conteo_Conceptos).Importe_IEPS, "#0.00"))
        '    'Cadena_Original = Cadena_Original & "|" & CDbl(Format(Consulta_Impuestos("Importe_IEPS", "Porcentaje_IEPS", CFD_IEPS(Conteo_Estructura).Tasa_IEPS), "#0.00"))
        'End If
        If CFD_Generales.Tipo_Factura = "Arrendamiento" Then
                'If CFD_Conceptos(Conteo_Conceptos).No_Predial <> "" And CFD_Conceptos(Conteo_Conceptos).No_Predial <> 0 Then
                '    Cadena_Original = Cadena_Original & "|" & CFD_Conceptos(Conteo_Conceptos).No_Predial
                ' End If
                End If
    Next Conteo_Conceptos
    
    If UBound(CFD_Impuestos_Retenidos) > 0 Then
        For Conteo_Estructura2 = 1 To UBound(CFD_Impuestos_Retenidos)
                                Cadena_Original = Cadena_Original & "|" & Format(CFD_Generales.SubTotal, "#0.00")
                                Cadena_Original = Cadena_Original & "|" & CFD_Impuestos_Retenidos(Conteo_Estructura2).Impuesto
                                Cadena_Original = Cadena_Original & "|" & "Tasa"
                                If CFD_Impuestos_Retenidos(Conteo_Estructura2).Impuesto = "002" Then
                                    Cadena_Original = Cadena_Original & "|" & Val(CFD_Impuestos_Retenidos(1).Tasa)
                                Else
                                    Cadena_Original = Cadena_Original & "|" & Val(CFD_Impuestos_Retenidos(2).Tasa)
                                End If
                                Cadena_Original = Cadena_Original & "|" & Format(CFD_Impuestos_Retenidos(Conteo_Estructura2).Importe, "#0.00")
                            Next Conteo_Estructura2
    End If
    
    'Datos de Retencion
    If UBound(CFD_Impuestos_Retenidos) > 0 Then
        Suma_Retenciones = 0
        For Conteo_Conceptos = 1 To UBound(CFD_Impuestos_Retenidos)
            'Cadena_Original = Cadena_Original & "|" & Format(Val(CFD_Impuestos_Retenidos(Conteo_Conceptos).Impuesto), "#0.00") & "|" & Format(Val(CFD_Impuestos_Retenidos(Conteo_Conceptos).Importe), "#0.00") '& "|"
            Suma_Retenciones = Suma_Retenciones + Format(Val(CFD_Impuestos_Retenidos(Conteo_Conceptos).Importe), "#0.00")
            
            'Cadena_Original = Cadena_Original & "|" & Format(CFD_Generales.Subtotal, "#0.00")
            'Cadena_Original = Cadena_Original & "|" & CFD_Impuestos_Retenidos(Conteo_Conceptos).Impuesto
            'Cadena_Original = Cadena_Original & "|" & "Tasa"
            cadena = cadena & "|" & Format(CFD_Generales.SubTotal, "#0.00")
            cadena = cadena & "|" & CFD_Impuestos_Retenidos(Conteo_Conceptos).Impuesto
            cadena = cadena & "|" & "Tasa"
            If CFD_Impuestos_Retenidos(Conteo_Conceptos).Impuesto = "002" Then
               'Cadena_Original = Cadena_Original & "|" & Format(Val(CFD_Impuestos_Retenidos(1).Tasa / 100), "#0.00")
               cadena = cadena & "|" & Val(CFD_Impuestos_Retenidos(1).Tasa)
                Else
                    'Cadena_Original = Cadena_Original & "|" & Format(Val(CFD_Impuestos_Retenidos(2).Tasa / 100), "#0.00")
                    cadena = cadena & "|" & Val(CFD_Impuestos_Retenidos(2).Tasa)
                End If
            'Cadena_Original = Cadena_Original & "|" & Format(CFD_Impuestos_Retenidos(Conteo_Conceptos).Importe, "#0.00")
            cadena = cadena & "|" & Format(CFD_Impuestos_Retenidos(Conteo_Conceptos).Importe, "#0.00")
            
        Next
        Cadena_Original = Cadena_Original & "|" & Format(Suma_Retenciones, "#0.00")
        'Cadena_Original = Cadena_Original & "|" & Format(CFD_Generales.Total_Retenciones, "#0.00")
        Else
            Cadena_Original = Cadena_Original & "|" & Format(0, "#0.00")
    End If
   
    'Datos de los Impuestos
   If UBound(CFD_Impuestos) > 0 Then
        'Cadena_Original = Cadena_Original & "|" '& "Tasa"
        For Conteo_Conceptos = 1 To UBound(CFD_Impuestos)
            Suma_Trasladados = Suma_Trasladados + CDbl(CFD_Impuestos(Conteo_Conceptos).Importe)
            'If UBound(CFD_Impuestos) = Conteo_Conceptos Then
                
                'Cadena_Original = Cadena_Original & "|"
                'Cadena_Original = Cadena_Original & "|" & CDbl(Format(CFD_Impuestos(Conteo_Conceptos).Importe, "#0.00"))
                'Cadena_Original = Cadena_Original & "|" & CDbl(Format(Val(CFD_Generales.Subtotal), "#0.00"))
                'Cadena_Original = Cadena_Original & "|" & CFD_Impuestos(Conteo_Conceptos).Impuesto
                'Cadena_Original = Cadena_Original & "|" & "Tasa"
                'Cadena_Original = Cadena_Original & "|" & Format(Val(CFD_Generales.Tasa_IVA / 100), "#0.00")
                'Cadena_Original = Cadena_Original & "|" & CDbl(Format(Val(CFD_Impuestos(Conteo_Conceptos).Importe), "#0.00"))
                If CFD_Impuestos(Conteo_Conceptos).Impuesto = "002" Then
                cadena = cadena & "|" & CDbl(Format(Val(CFD_Generales.SubTotal), "#0.00"))
                cadena = cadena & "|" & CFD_Impuestos(Conteo_Conceptos).Impuesto
                cadena = cadena & "|" & "Tasa"
                cadena = cadena & "|" & Format(Val(CFD_Impuestos(Conteo_Conceptos).Tasa / 100), "#0.00")
                cadena = cadena & "|" & CDbl(Format(Val(CFD_Impuestos(Conteo_Conceptos).Importe), "#0.00"))
                End If
                
                'If CFD_Impuestos(Conteo_Conceptos).Impuesto = "003" Then
                '             'If CFD_Impuestos(2).Importe > 0 Then
                '                Cadena = Cadena & "|" & CDbl(Format(CFD_Generales.SubTotal, "#0.00"))
                '                Cadena = Cadena & "|" & CFD_Impuestos(Conteo_Conceptos).Impuesto
                '                Cadena = Cadena & "|" & CFD_IEPS(1).Factor
                '                Cadena = Cadena & "|" & Val(CFD_IEPS(1).Tasa_IEPS / 100) 'double
                '                Cadena = Cadena & "|" & CDbl(Format(CFD_Impuestos(2).Importe, "#0.00"))
                '            End If
                '& "|" & CDbl(Format(Suma_Trasladados, "#0.00")) 'Format(CFD_Impuestos(Conteo_Conceptos).Importe, "#0.00")
            'Else
                'Cadena_Original = Cadena_Original & "|" & CFD_Impuestos(Conteo_Conceptos).Impuesto & "|" & CDbl(Format(CFD_Impuestos(Conteo_Conceptos).Tasa, "#0.00")) & "|" & CDbl(Format(CFD_Impuestos(Conteo_Conceptos).Importe, "#0.00"))
            'End If
        Next Conteo_Conceptos
        Cadena_Original = Cadena_Original & "|" & CDbl(Format(Suma_Trasladados, "#0.00"))
        Cadena_Original = Cadena_Original & cadena
    End If
    
    'Datos de Retencion de impuestos locales
    If UBound(CFD_Impuestos_Locales) > 0 Then
        Suma_Locales_Retenciones = 0
        Suma_Locales_Trasladados = 0
        Retenciones_Locales = ""
        Trasladados_Locales = ""
        'Consulta el importe de cada retenciones y trasladados
        'For Conteo_Conceptos = 1 To UBound(CFD_Impuestos_Locales)
        '    If CFD_Impuestos_Locales(Conteo_Conceptos).Tipo = "R" Then
        '        Suma_Locales_Retenciones = Suma_Locales_Retenciones + CDbl(CFD_Impuestos_Locales(Conteo_Conceptos).Importe)
        '        Retenciones_Locales = Retenciones_Locales & CFD_Impuestos_Locales(Conteo_Conceptos).Impuesto & "|" & CDbl(CFD_Impuestos_Locales(Conteo_Conceptos).Tasa) & "|" & CDbl(CFD_Impuestos_Locales(Conteo_Conceptos).Importe)
        '    Else
        '        Suma_Locales_Trasladados = Suma_Locales_Trasladados + CDbl(CFD_Impuestos_Locales(Conteo_Conceptos).Importe)
        '        Trasladados_Locales = Trasladados_Locales & CFD_Impuestos_Locales(Conteo_Conceptos).Impuesto & "|" & CDbl(CFD_Impuestos_Locales(Conteo_Conceptos).Tasa) & "|" & CDbl(CFD_Impuestos_Locales(Conteo_Conceptos).Importe)
        '    End If
        'Next Conteo_Conceptos
        'If CFD_Conceptos(Conteo_Conceptos).No_Predial <> "" Then
         '   Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Conceptos(Conteo_Conceptos).No_Predial)
        'End If
        'Estructura = version|total de retenciones|total de traslados|impuesto retenido|tasa de retencion|importe|impuesto trasladado|tasa de traslado|importe
        Cadena_Original = Cadena_Original & "|" & "1.0" & "|" & Format(Suma_Locales_Retenciones, "#0.00") & "|" & Format(Suma_Locales_Trasladados, "#0.00")
        Debug.Print Cadena_Original
        If Retenciones_Locales <> "" Then
            Cadena_Original = Cadena_Original & "|" & Retenciones_Locales
        End If
        If Trasladados_Locales <> "" Then
            Cadena_Original = Cadena_Original & "|" & Trasladados_Locales
        End If
        Cadena_Original = Cadena_Original & "||"
    Else
        'Complemento de pagos
    If CFD_Generales.Tipo_Factura = "Pagos" Then
            Cadena_Original = Cadena_Original & "|" & "1.0"
            Cadena_Original = Cadena_Original & "|" & CFD_Pagos.Fecha_Pago
            Cadena_Original = Cadena_Original & "|" & Mid(CFD_Pagos.Forma_Pago_Pagos, 1, 2)
            Cadena_Original = Cadena_Original & "|" & CFD_Pagos.Moneda_Pago
            If CFD_Pagos.Tipo_Cambio_Pago <> "" And Val(CFD_Pagos.Tipo_Cambio_Pago) > 0 Then Cadena_Original = Cadena_Original & "|" & CFD_Pagos.Tipo_Cambio_Pago
                Cadena_Original = Cadena_Original & "|" & CFD_Pagos.Monto
                If Mid(CFD_Pagos.Forma_Pago_Pagos, 1, 2) <> "01" Then
                    If CFD_Pagos.Num_Operacion <> "" Then Cadena_Original = Cadena_Original & "|" & CFD_Pagos.Num_Operacion
                    If CFD_Pagos.RfcEmisorCtaOrd <> "" Then Cadena_Original = Cadena_Original & "|" & CFD_Pagos.RfcEmisorCtaOrd
                    If CFD_Pagos.NomBancoOrdExt <> "" Then Cadena_Original = Cadena_Original & "|" & CFD_Pagos.NomBancoOrdExt
                    If CFD_Pagos.CtaOrdenante <> "" Then Cadena_Original = Cadena_Original & "|" & CFD_Pagos.CtaOrdenante
                    If CFD_Pagos.RfcEmisorCtaBen <> "" Then Cadena_Original = Cadena_Original & "|" & CFD_Pagos.RfcEmisorCtaBen
                    If CFD_Pagos.CtaBeneficiario <> "" Then Cadena_Original = Cadena_Original & "|" & CFD_Pagos.CtaBeneficiario
                    If CFD_Pagos.TipoCadPago <> "" Then Cadena_Original = Cadena_Original & "|" & Mid(CFD_Pagos.TipoCadPago, 1, 2)
                    If CFD_Pagos.CertPago <> "" Then Cadena_Original = Cadena_Original & "|" & CFD_Pagos.CertPago
                    If CFD_Pagos.CadPago <> "" Then Cadena_Original = Cadena_Original & "|" & CFD_Pagos.CadPago
                    If CFD_Pagos.Sello_Pago <> "" Then Cadena_Original = Cadena_Original & "|" & CFD_Pagos.Sello_Pago
                End If
                For i = 1 To UBound(CFD_Pagos_DR)
                    Cadena_Original = Cadena_Original & "|" & CFD_Pagos_DR(i).ID_Doc
                    If CFD_Pagos_DR(i).Serie <> "" Then Cadena_Original = Cadena_Original & "|" & CFD_Pagos_DR(i).Serie
                    Cadena_Original = Cadena_Original & "|" & CFD_Pagos_DR(i).Folio
                    Cadena_Original = Cadena_Original & "|" & CFD_Pagos_DR(i).Moneda_DR
                    If CFD_Pagos_DR(i).Tipo_Cambio_DR <> "" And Val(CFD_Pagos_DR(i).Tipo_Cambio_DR) > 0 Then Cadena_Original = Cadena_Original & "|" & CDbl(Val(CFD_Pagos_DR(i).Tipo_Cambio_DR))
                    Cadena_Original = Cadena_Original & "|" & Mid(CFD_Pagos_DR(i).Metodo_Pago_DR, 1, 3)
                    Cadena_Original = Cadena_Original & "|" & CFD_Pagos_DR(i).No_Parcialidad
                    Cadena_Original = Cadena_Original & "|" & CFD_Pagos_DR(i).Saldo_Anterior
                    Cadena_Original = Cadena_Original & "|" & CFD_Pagos_DR(i).Importe_Pagado
                    Cadena_Original = Cadena_Original & "|" & CFD_Pagos_DR(i).Saldo_Insoluto
                Next
           End If
        If Tipo_Factura = "DONATIVOS" Then
            'Estructura = version|número de autorizacíon para emitir recibos de donación|fecha de autorización para emitir recibos de donación|leyenda de donación
            'Cadena_Original = Cadena_Original & "|" & "1.1" & "|" & No_Autorizacion_Donacion & "|" & Fecha_Autorizacion_Donacion & "|" & Leyenda_Donacion & "||"
        Else
            Cadena_Original = Cadena_Original & "||"
        End If
    End If

    'Regresa la cadena original
    Debug.Print (Cadena_Original)
'    CFD_Cadena_Original = Cadena_Original
    
End Function


Public Function CFD_Crea_PDF_Pagos() As rdoResultset
    Dim report As New CRAXDDRT.report
    Dim sql As String
    Dim crapp As New CRAXDDRT.Application
    Dim tbl As CRAXDDRT.DatabaseTable
    Dim Rs_Consulta As rdoResultset
    Dim crxReport As CRAXDDRT.report

    'I modified the query for the report by specifying the WHERE clause
    sql = "SELECT No_Certificado, Timbre_NoCertificadoSA, Timbre_FechaTimbrado, Timbre_UUID, Direccion, Colonia, Localidad, Tipo_Relacion, RFC, Codigo_Postal, Fecha, Forma_Pago, Num_Operacion, RFC_Emisor_Cta_Ord, Cliente, No_Factura, Serie, Saldo_Actual, Monto_Pagado, Saldo_Anterior, Timbre_Version, Timbre_SelloCFD, Cta_Beneficiario, Tipo_Cad_Pago, UUID_Relacionado, No_Parcialidad, No_Factura_Ref"
    sql = sql & " FROM   Complemento_Pago"
    sql = sql & " WHERE  No_Factura='" & Format(CFD_Generales.Folio, "#0000000000") & "' AND Serie='" & CFD_Generales.Serie & "'"
    Set Rs_Consulta = Conectar_Ayudante.Recordset_Consultar(sql)
    Set CFD_Crea_PDF_Pagos = Rs_Consulta
    
    'Rs_Consulta.Close
End Function

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: CFD_Elimina_Espacios
'DESCRIPCIÓN: Eliminas los n espacios entre palabras
'PARÁMETROS:
'               1. Cadena, cadena a eliminar los espacios
'CREO: Ismael Prieto Sánchez
'FECHA_CREO: 16/Oct/2009 12:30 pm
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function CFD_Elimina_Espacios(cadena As String) As String
Dim Grupo_Cadena() As String
Dim Cuenta_Cadena As Integer
Dim Cadena_Nueva As String

    'Divide la cadena por los espacios
    Grupo_Cadena = Split(cadena, " ")
    
    'Valida la longitud
    If UBound(Grupo_Cadena) > 0 Then
        Cadena_Nueva = ""
        For Cuenta_Cadena = 0 To UBound(Grupo_Cadena)
            If Trim(Grupo_Cadena(Cuenta_Cadena)) <> "" Then
                Cadena_Nueva = Cadena_Nueva & " " & Trim(Grupo_Cadena(Cuenta_Cadena))
            End If
        Next Cuenta_Cadena
        If Trim(Cadena_Nueva) <> "" Then
            CFD_Elimina_Espacios = Trim(CFD_Valida_Caracteres_UTF(Trim(Cadena_Nueva)))
        Else
            CFD_Elimina_Espacios = Trim(Cadena_Nueva)
        End If
    Else
        If Trim(cadena) <> "" Then
            CFD_Elimina_Espacios = Trim(CFD_Valida_Caracteres_UTF(Trim(cadena)))
        Else
            CFD_Elimina_Espacios = Trim(cadena)
        End If
    End If
End Function


'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: CFD_Genera_MD5
'DESCRIPCIÓN: Genera la cadena con encripción MD5
'PARÁMETROS:
'               1. Cadena_UTF, para pasarle la cadena formateada a UTF-8
'               2. Regresa el valor de la cadena encriptada en MD5
'CREO: Ismael Prieto Sánchez
'FECHA_CREO: 21/Oct/2006 11:00 am
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function CFD_Genera_MD5(Cadena_UTF As String) As String
Dim strDigest As String
Dim strDataUTF8 As String
Dim nRet As String

    'Asigna el valor a la variable temporal
    strDataUTF8 = Cadena_UTF

    'Crea el Hash con la primera cadena recibida
    strDigest = String(32, " ")
    nRet = HASH_HexFromString(strDigest, Len(strDigest), strDataUTF8, Len(strDataUTF8), PKI_HASH_MD5)
    
    'Regresa el valor a la funcion
    CFD_Genera_MD5 = strDigest
End Function

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Valida_Caracteres_Especiales
'DESCRIPCIÓN: Valida los caracteres especiales en una cadena cambiandolo por uno valido
'PARÁMETROS:
'             1. Cadena
'             2. Regresa la nueva cadena a la función
'CREO: Ismael Prieto Sánchez
'FECHA_CREO: 05/Oct/2006 2:00 pm
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Private Function Valida_Caracteres_Especiales(ByVal cadena As String) As String
Dim Caracteres_Especiales(5) As String
Dim Caracteres_Escape(5) As String
Dim Cuenta_Caracteres As Integer
Dim Posicion_Caracter As Integer

''''    'Crea la lista de caracteres especiales
''''    Caracteres_Especiales(0) = "&"
''''    Caracteres_Especiales(1) = """"
''''    Caracteres_Especiales(2) = "<"
''''    Caracteres_Especiales(3) = ">"
''''    Caracteres_Especiales(4) = "'"
''''
''''    'Crea la lista de caracteres de escape
''''    Caracteres_Escape(0) = "&amp;"
''''    Caracteres_Escape(1) = "&quot;"
''''    Caracteres_Escape(2) = "&lt;"
''''    Caracteres_Escape(3) = "&gt;"
''''    Caracteres_Escape(4) = "&#39"
''''
''''    'Comienza el analisis de la cadena de caracteres
''''    For Cuenta_Caracteres = 0 To 4
''''        Cadena = Replace(Cadena, Caracteres_Especiales(Cuenta_Caracteres), Caracteres_Escape(Cuenta_Caracteres), 1, , vbTextCompare)
''''    Next Cuenta_Caracteres
    
    'Regresa el nuevo valor a la función
    Valida_Caracteres_Especiales = cadena
End Function

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: CFD_Crea_Xml
'DESCRIPCIÓN: Crea el archivo xml de acuerdo a los datos para generar el CFD
'PARÁMETROS:
'               1. Nombre_CFD, toma el nombre del archivo xml
'CREO: Ismael Prieto Sánchez
'FECHA_CREO: 06/Oct/2006 5:40 pm
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Sub CFD_Crea_Xml_OLD(Nombre_CFD As String)
Dim Xml_Identificacion As IXMLDOMProcessingInstruction
Dim Nodo_Raiz As IXMLDOMElement
Dim Nodo_Emisor As IXMLDOMElement
Dim Nodo_Emisor_Domicilio As IXMLDOMElement
Dim Nodo_ExpedidoEn As IXMLDOMElement
Dim Nodo_Receptor As IXMLDOMElement
Dim Nodo_Receptor_Domicilio As IXMLDOMElement
Dim Nodo_Conceptos As IXMLDOMElement
Dim Nodo_Concepto As IXMLDOMElement
Dim Nodo_Impuestos As IXMLDOMElement
Dim Nodo_Traslados As IXMLDOMElement
Dim Nodo_Traslado As IXMLDOMElement
Dim Nodo_Retenciones As IXMLDOMElement
Dim Nodo_Retencion As IXMLDOMElement
Dim Nodo_Complemento As IXMLDOMElement
Dim Nodo_ImpuestoLocal As IXMLDOMElement
Dim Nodo_Elemento As IXMLDOMElement
Dim Nodo_Timbrado As IXMLDOMElement
Dim Nodo_Regimen As IXMLDOMElement
Dim Impuestos_Trasladados As Double
Dim Impuestos_Retenidos As Double
Dim Impuestos_Locales As Double
Dim Conteo_Estructura As Integer
Dim Cadena_XML As String
Dim Genera_XML As String
Dim Proceso_Web_Services As String
Dim Respuesta_Web_Services As String
Dim Respuesta As IXMLDOMNode
Dim CFDI As String
Dim XML As String
Dim Timbrado As New ContelTimbrado_33.Cls_Fe_Timbrado
Dim Codigo_Bidimensional As String

    '*****************************************************************************************
    'CREA DOCUMENTO XML
    '*****************************************************************************************
    'Crea el documento xml
    Set CFD_Documento = New DOMDocument
    
    'Identifica la version y codificacion del xml
    Set Xml_Identificacion = CFD_Documento.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
    CFD_Documento.appendChild Xml_Identificacion
    
    'Crea el documento xml
    Set Nodo_Raiz = CFD_Documento.createElement("cfdi:Comprobante")
    Set CFD_Documento.documentElement = Nodo_Raiz
    
        '*****************************************************************************************
        'CREA NODO PRINCIPAL
        '*****************************************************************************************
        'Crea el elemento principal
        Set Nodo_Principal = CFD_Documento.documentElement
            'Agrega los atributos a la raiz del xml
            Nodo_Principal.setAttribute "xsi:schemaLocation", "http://www.sat.gob.mx/cfd/3 http://www.sat.gob.mx/sitio_internet/cfd/3/cfdv32.xsd"
            Nodo_Principal.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
            Nodo_Principal.setAttribute "xmlns:cfdi", "http://www.sat.gob.mx/cfd/3"
            Nodo_Principal.setAttribute "version", CFD_Generales.Version
            If CFD_Generales.Serie <> "" Then
                Nodo_Principal.setAttribute "serie", CFD_Generales.Serie
            End If
            Nodo_Principal.setAttribute "folio", CFD_Generales.Folio
            Nodo_Principal.setAttribute "fecha", CFD_Generales.Fecha
            Nodo_Principal.setAttribute "sello", CFD_Generales.Sello
            Nodo_Principal.setAttribute "formaDePago", LCase(CFD_Generales.Forma_Pago)
            Nodo_Principal.setAttribute "noCertificado", CFD_Generales.No_Certificado
            Nodo_Principal.setAttribute "certificado", CFD_Generales.Certificado
            If CFD_Generales.Condiciones_Pago <> "" Then
                Nodo_Principal.setAttribute "condicionesDePago", CFD_Generales.Condiciones_Pago
            End If
            Nodo_Principal.setAttribute "subTotal", Format(CFD_Generales.SubTotal, "#0.00")
            If CFD_Generales.Descuento > 0 Then
                Nodo_Principal.setAttribute "descuento", Format(CFD_Generales.Descuento, "#0.00")
            End If
            If CFD_Generales.Tipo_Moneda = "DOLARES" Then
                Nodo_Principal.setAttribute "Tipo_Cambio", CFD_Generales.Tipo_Cambio
            End If
            Nodo_Principal.setAttribute "Moneda", CFD_Generales.Tipo_Moneda
            Nodo_Principal.setAttribute "total", Format(CFD_Generales.Total, "#0.00")
            Nodo_Principal.setAttribute "tipoDeComprobante", CFD_Generales.Tipo_Comprobante
            Nodo_Principal.setAttribute "metodoDePago", CFD_Generales.Metodo_Pago
            Nodo_Principal.setAttribute "LugarExpedicion", CFD_Emisor.Expedido_En
            If CFD_Generales.Cuenta_Pago <> "" Then
                Nodo_Principal.setAttribute "NumCtaPago", CFD_Generales.Cuenta_Pago
            End If
        'Agrega un salto de linea
        Nodo_Principal.appendChild CFD_Documento.createTextNode(vbCrLf)

        '*****************************************************************************************
        'CREA NODO EMISOR
        '*****************************************************************************************
        'Agrega el elemento del Emisor
        Set Nodo_Emisor = CFD_Documento.createElement("cfdi:Emisor")
            'Agrega los atributos al nodo Emisor
            Nodo_Emisor.setAttribute "rfc", CFD_Emisor.RFC
            If CFD_Emisor.Nombre <> "" Then
                Nodo_Emisor.setAttribute "nombre", CFD_Emisor.Nombre
            End If
            'Agrega un salto de linea
            Nodo_Emisor.appendChild CFD_Documento.createTextNode(vbCrLf)
            
            'Agrega los elementos al Emisor
            Set Nodo_Emisor_Domicilio = CFD_Documento.createElement("cfdi:DomicilioFiscal")
                'Agrega los atributos del domicilio al nodo Emisor
                Nodo_Emisor_Domicilio.setAttribute "calle", CFD_Emisor.Calle
                If CFD_Emisor.No_Exterior <> "" Then
                    Nodo_Emisor_Domicilio.setAttribute "noExterior", CFD_Emisor.No_Exterior
                End If
                If CFD_Emisor.No_Interior <> "" Then
                    Nodo_Emisor_Domicilio.setAttribute "noInterior", CFD_Emisor.No_Interior
                End If
                If CFD_Emisor.Colonia <> "" Then
                    Nodo_Emisor_Domicilio.setAttribute "colonia", CFD_Emisor.Colonia
                End If
                If CFD_Emisor.Localidad <> "" Then
                    Nodo_Emisor_Domicilio.setAttribute "localidad", CFD_Emisor.Localidad
                End If
                Nodo_Emisor_Domicilio.setAttribute "municipio", CFD_Emisor.Municipio
                Nodo_Emisor_Domicilio.setAttribute "estado", CFD_Emisor.Estado
                Nodo_Emisor_Domicilio.setAttribute "pais", CFD_Emisor.Pais
                Nodo_Emisor_Domicilio.setAttribute "codigoPostal", CFD_Emisor.cp
            'Asigna el nodo del domicilio al nodo emisor
            Nodo_Emisor.appendChild Nodo_Emisor_Domicilio
        'cierra el nodo
        Nodo_Emisor.appendChild CFD_Documento.createTextNode(vbCrLf)
    
        '*****************************************************************************************
        'CREA NODO REGIMEN FISCAL
        '*****************************************************************************************
        'Agrega los elementos al Emisor
        Set Nodo_Regimen = CFD_Documento.createElement("cfdi:RegimenFiscal")
            'Agrega los atributos del domicilio al nodo Emisor
            Nodo_Regimen.setAttribute "Regimen", Regimen_Emisor
            'Asigna el nodo del domicilio al nodo emisor
            Nodo_Emisor.appendChild Nodo_Regimen
        'Agrega un salto de linea
        Nodo_Emisor.appendChild CFD_Documento.createTextNode(vbCrLf)
                
        'Asigna el nodo Emisor al nodo principal
        Nodo_Principal.appendChild Nodo_Emisor
        'Agrega un salto de linea
        Nodo_Principal.appendChild CFD_Documento.createTextNode(vbCrLf)
            
        '*****************************************************************************************
        'CREA NODO RECEPTOR
        '*****************************************************************************************
        'Agrega el elemento del Receptor
        Set Nodo_Receptor = CFD_Documento.createElement("cfdi:Receptor")
        
            'Agrega los atributos al nodo Emisor
            Nodo_Receptor.setAttribute "rfc", CFD_Receptor.RFC
            If CFD_Receptor.Nombre <> "" Then
                Nodo_Receptor.setAttribute "nombre", CFD_Receptor.Nombre
            End If
            
            'Agrega un salto de linea
            Nodo_Receptor.appendChild CFD_Documento.createTextNode(vbCrLf)
                'Agrega los elementos al Receptor
                Set Nodo_Receptor_Domicilio = CFD_Documento.createElement("cfdi:Domicilio")
                    'Agrega los atributos del domicilio al nodo Receptor
                    Nodo_Receptor_Domicilio.setAttribute "calle", CFD_Receptor.Calle
                    If CFD_Receptor.No_Exterior <> "" Then
                        Nodo_Receptor_Domicilio.setAttribute "noExterior", CFD_Receptor.No_Exterior
                    End If
                    If CFD_Receptor.No_Interior <> "" Then
                        Nodo_Receptor_Domicilio.setAttribute "noInterior", CFD_Receptor.No_Interior
                    End If
                    If CFD_Receptor.Colonia <> "" Then
                        Nodo_Receptor_Domicilio.setAttribute "colonia", CFD_Receptor.Colonia
                    End If
                    If CFD_Receptor.Localidad <> "" Then
                        Nodo_Receptor_Domicilio.setAttribute "localidad", CFD_Receptor.Localidad
                    End If
                    Nodo_Receptor_Domicilio.setAttribute "municipio", CFD_Receptor.Municipio
                    Nodo_Receptor_Domicilio.setAttribute "estado", CFD_Receptor.Estado
                    Nodo_Receptor_Domicilio.setAttribute "pais", CFD_Receptor.Pais
                    Nodo_Receptor_Domicilio.setAttribute "codigoPostal", CFD_Receptor.cp
                'Asigna el nodo del domicilio al nodo Receptor
                Nodo_Receptor.appendChild Nodo_Receptor_Domicilio
            'Agrega un salto de linea
            Nodo_Receptor.appendChild CFD_Documento.createTextNode(vbCrLf)
        
        'Asigna el nodo Receptor al nodo principal
        Nodo_Principal.appendChild Nodo_Receptor
        'Agrega un salto de linea
        Nodo_Principal.appendChild CFD_Documento.createTextNode(vbCrLf)
    
        '*****************************************************************************************
        'CREA NODO CONCEPTOS
        '*****************************************************************************************
        'Agrega el elemento de Conceptos
        Set Nodo_Conceptos = CFD_Documento.createElement("cfdi:Conceptos")
        
            'Agrega los atributos del Concepto al nodo Conceptos
            For Conteo_Estructura = 1 To UBound(CFD_Conceptos)
                'Agrega un salto de linea
                Nodo_Conceptos.appendChild CFD_Documento.createTextNode(vbCrLf)
                'Agrega los elementos al Concepto
                Set Nodo_Concepto = CFD_Documento.createElement("cfdi:Concepto")
                    'Asigna los atributos
                    Nodo_Concepto.setAttribute "cantidad", Format(CFD_Conceptos(Conteo_Estructura).Cantidad, "#0.00")
                    Nodo_Concepto.setAttribute "unidad", CFD_Conceptos(Conteo_Estructura).Unidad
                    If CFD_Conceptos(Conteo_Estructura).No_Identificacion <> "" Then
                        Nodo_Concepto.setAttribute "noIdentificacion", CFD_Conceptos(Conteo_Estructura).No_Identificacion
                    End If
                    Nodo_Concepto.setAttribute "descripcion", CFD_Conceptos(Conteo_Estructura).Descripcion
                    Nodo_Concepto.setAttribute "valorUnitario", Format(CFD_Conceptos(Conteo_Estructura).Valor_Unitario, "#0.00")
                    Nodo_Concepto.setAttribute "importe", Format(CFD_Conceptos(Conteo_Estructura).Importe, "#0.00")
                'Asigna el nodo de Conceptos al nodo Conceptos
                Nodo_Conceptos.appendChild Nodo_Concepto
            Next Conteo_Estructura
            'Agrega un salto de linea
            Nodo_Conceptos.appendChild CFD_Documento.createTextNode(vbCrLf)
        
        'Asigna el nodo Conceptos al nodo principal
        Nodo_Principal.appendChild Nodo_Conceptos
        'Agrega un salto de linea
        Nodo_Principal.appendChild CFD_Documento.createTextNode(vbCrLf)
    
        '*****************************************************************************************
        'CREA NODO IMPUESTOS
        '*****************************************************************************************
        'Agrega el elemento de Impuestos
        Set Nodo_Impuestos = CFD_Documento.createElement("cfdi:Impuestos")
            
            'Inicializa las variables
            Impuestos_Retenidos = 0
            Impuestos_Trasladados = 0
            
            'Verifica si existe monto de retenciones
            For Conteo_Estructura = 1 To UBound(CFD_Impuestos_Retenidos)
                Impuestos_Retenidos = Impuestos_Retenidos + CFD_Impuestos_Retenidos(Conteo_Estructura).Importe
            Next Conteo_Estructura
            
            'Verifica si existe monto de trasladados
            For Conteo_Estructura = 1 To UBound(CFD_Impuestos)
                Impuestos_Trasladados = Impuestos_Trasladados + CFD_Impuestos(Conteo_Estructura).Importe
            Next Conteo_Estructura
                           
            'si hay impuestos retenidos los toma en cuenta para mostrarlos en el xml
            If Impuestos_Retenidos > 0 Then
                'Agrega los atributos del nodo Impuestos
                Nodo_Impuestos.setAttribute "totalImpuestosRetenidos", Format(Impuestos_Retenidos, "#0.00")
                Nodo_Impuestos.setAttribute "totalImpuestosTrasladados", Format(Impuestos_Trasladados, "#0.00")
                'Agrega los elementos de retenidos
                Set Nodo_Retenciones = CFD_Documento.createElement("cfdi:Retenciones")
                    'Agrega un salto de linea
                    Nodo_Impuestos.appendChild CFD_Documento.createTextNode(vbCrLf)
                    'Agrega un salto de linea
                    Nodo_Retenciones.appendChild CFD_Documento.createTextNode(vbCrLf)
                    'Agrega los atributos del Concepto al nodo Conceptos
                    For Conteo_Estructura = 1 To UBound(CFD_Impuestos_Retenidos)
                        'Agrega los elementos de Retenciones
                        Set Nodo_Retencion = CFD_Documento.createElement("cfdi:Retencion")
                        'Agrega los atributos de la retencion al nodo Retenciones
                        Nodo_Retencion.setAttribute "impuesto", CFD_Impuestos_Retenidos(Conteo_Estructura).Impuesto
                        Nodo_Retencion.setAttribute "importe", Format(CFD_Impuestos_Retenidos(Conteo_Estructura).Importe, "#0.00")
                        'Asigna el nodo de Traslado al nodo Retenciones
                        Nodo_Retenciones.appendChild Nodo_Retencion
                        'Agrega un salto de linea
                        Nodo_Retenciones.appendChild CFD_Documento.createTextNode(vbCrLf)
                    Next Conteo_Estructura
                'Cierra el nodo de Retenciones al nodo Impuestos
                Nodo_Impuestos.appendChild Nodo_Retenciones
            Else
                'Si no hubo retenciones agrega el atributo de impuestos trasladados
                Nodo_Impuestos.setAttribute "totalImpuestosTrasladados", Format(Impuestos_Trasladados, "#0.00")
            End If
                
            'Agrega los elementos de impuestos Traslados
            Set Nodo_Traslados = CFD_Documento.createElement("cfdi:Traslados")
                'Agrega un salto de linea
                Nodo_Traslados.appendChild CFD_Documento.createTextNode(vbCrLf)
                'Agrega los atributos del Concepto al nodo Conceptos
                For Conteo_Estructura = 1 To UBound(CFD_Impuestos)
                    'Agrega los elementos de Traslado
                    Set Nodo_Traslado = CFD_Documento.createElement("cfdi:Traslado")
                        'Agrega los atributos del Traslado al nodo Traslados
                        Nodo_Traslado.setAttribute "impuesto", CFD_Impuestos(Conteo_Estructura).Impuesto
                        Nodo_Traslado.setAttribute "tasa", Format(CFD_Impuestos(Conteo_Estructura).Tasa, "#0.00")
                        Nodo_Traslado.setAttribute "importe", Format(CFD_Impuestos(Conteo_Estructura).Importe, "#0.00")
                    'Asigna el nodo de Traslado al nodo Traslados
                    Nodo_Traslados.appendChild Nodo_Traslado
                    'Agrega un salto de linea
                    Nodo_Traslados.appendChild CFD_Documento.createTextNode(vbCrLf)
                Next Conteo_Estructura
            'Cierra el nodo de Traslados al nodo Impuestos
            Nodo_Impuestos.appendChild Nodo_Traslados
            'Agrega un salto de linea
            Nodo_Impuestos.appendChild CFD_Documento.createTextNode(vbCrLf)
        'Cierra el nodo Impuestos al nodo principal
        Nodo_Principal.appendChild Nodo_Impuestos
    
        '*****************************************************************************************
        'CREA NODO COMPLEMENTO
        '*****************************************************************************************
        Nodo_Principal.appendChild CFD_Documento.createTextNode(vbCrLf)
        'Agrega el elemento de la complemento
        Set Nodo_Complemento = CFD_Documento.createElement("cfdi:Complemento")
            'Agrega un salto de linea
            Nodo_Complemento.appendChild CFD_Documento.createTextNode(vbCrLf)
                
            'Inicializa las variables
            Impuestos_Locales = 0
                
            'Verifica si existe monto de impuestos locales
            For Conteo_Estructura = 1 To UBound(CFD_Impuestos_Locales)
                Impuestos_Locales = Impuestos_Locales + CFD_Impuestos_Locales(Conteo_Estructura).Importe
            Next Conteo_Estructura
            
            'Si el valor de impuestos locales es mayor a cero, agrega el nodo Complemento
            If Impuestos_Locales > 0 Then
                'Agrega el elemento Detalle
                Set Nodo_ImpuestoLocal = CFD_Documento.createElement("implocal:ImpuestosLocales")
                    'Agrega el elemento No Proveedor
                    Nodo_ImpuestoLocal.setAttribute "version", "1.0"
                    Nodo_ImpuestoLocal.setAttribute "TotaldeRetenciones", Format(Impuestos_Locales, "#0.00")
                    Nodo_ImpuestoLocal.setAttribute "TotaldeTraslados", Format(0, "#0.00")
                    Nodo_ImpuestoLocal.setAttribute "xmlns:implocal", "http://www.sat.gob.mx/implocal"
                'Cierra el nodo Encabezado
                Nodo_Complemento.appendChild Nodo_ImpuestoLocal
                'Agrega un salto de linea
                Nodo_ImpuestoLocal.appendChild CFD_Documento.createTextNode(vbCrLf)
                
                'Agrega los atributos del impuesto local retenido
                For Conteo_Estructura = 1 To UBound(CFD_Impuestos_Locales)
                    'Agrega los elementos al Concepto
                    Set Nodo_Elemento = CFD_Documento.createElement("implocal:RetencionesLocales")
                        'Asigna los atributos
                        Nodo_Elemento.setAttribute "ImpLocRetenido", CFD_Impuestos_Locales(Conteo_Estructura).Impuesto
                        Nodo_Elemento.setAttribute "TasadeRetencion", CFD_Impuestos_Locales(Conteo_Estructura).Tasa
                        Nodo_Elemento.setAttribute "Importe", CFD_Impuestos_Locales(Conteo_Estructura).Importe
                    'Asigna el nodo de Conceptos al nodo Conceptos
                    Nodo_ImpuestoLocal.appendChild Nodo_Elemento
                Next Conteo_Estructura
                'Agrega un salto de linea
                Nodo_ImpuestoLocal.appendChild CFD_Documento.createTextNode(vbCrLf)
                'Asigna el nodo Conceptos al nodo principal
                Nodo_Complemento.appendChild Nodo_ImpuestoLocal
            End If
        'Cierra el nodo Addenda al nodo principal
        Nodo_Principal.appendChild Nodo_Complemento
        'Agrega un salto de linea
        Nodo_Complemento.appendChild CFD_Documento.createTextNode(vbCrLf)
    'Cierra el nodo principal
    Nodo_Principal.appendChild CFD_Documento.createTextNode(vbCrLf)
      
    'Limpia las variables del timbre
    Timbrado_VersionSat = ""
    Timbrado_UUID = ""
    Timbrado_FechaTimbrado = ""
    Timbrado_selloCFD = ""
    Timbrado_noCertificadoSAT = ""
    Timbrado_selloSAT = ""
               
    'Realiza la petición al webservice
    If Timbrado_Ambiente = "PRUEBA" Then
        XML = Timbrado.GenerarTimbrado(False, Timbrado_Version, Timbrado_Codigo_Usuario_Proveedor, Timbrado_Codigo_Usuario, Timbrado_ID_Sucursal, CStr(CFD_Documento.XML))
    Else
        XML = Timbrado.GenerarTimbrado(True, Timbrado_Version, Timbrado_Codigo_Usuario_Proveedor, Timbrado_Codigo_Usuario, Timbrado_ID_Sucursal, CStr(CFD_Documento.XML))
    End If

    'Pone los datos en el analizador de XML
    Set Conexion_Web_Services = New DOMDocument
    
    'Parametros de validacion
    Conexion_Web_Services.resolveExternals = True

    Conexion_Web_Services.validateOnParse = True
    
    Conexion_Web_Services.async = False
   
    'Carga la respuesta del webservice
     If Conexion_Web_Services.loadXML(XML) Then  'Si es valida la respuesta
        'Obtiene la respuesta
        Set Respuesta = Conexion_Web_Services.SelectSingleNode("//Timbre")
        If Not Respuesta Is Nothing Then
            'Obtiene los parametros del resultado
            Timbrado_VersionSat = Respuesta.Attributes(0).text
            Timbrado_UUID = Respuesta.Attributes(1).text
            Timbrado_FechaTimbrado = Respuesta.Attributes(2).text
            Timbrado_selloCFD = Respuesta.Attributes(3).text
            Timbrado_noCertificadoSAT = Respuesta.Attributes(4).text
            Timbrado_selloSAT = Respuesta.Attributes(5).text
            'Agrega el elemento del timbrado
            Set Nodo_Timbrado = CFD_Documento.createElement("tfd:TimbreFiscalDigital")
                'Agrega los atributos
                Nodo_Timbrado.setAttribute "xsi:schemaLocation", "http://www.sat.gob.mx/TimbreFiscalDigital http://www.sat.gob.mx/TimbreFiscalDigital/TimbreFiscalDigital.xsd"
                Nodo_Timbrado.setAttribute "xmlns:tfd", "http://www.sat.gob.mx/TimbreFiscalDigital"
                Nodo_Timbrado.setAttribute "version", Timbrado_VersionSat
                Nodo_Timbrado.setAttribute "UUID", Timbrado_UUID
                Nodo_Timbrado.setAttribute "FechaTimbrado", Timbrado_FechaTimbrado
                Nodo_Timbrado.setAttribute "selloCFD", Timbrado_selloCFD
                Nodo_Timbrado.setAttribute "noCertificadoSAT", Timbrado_noCertificadoSAT
                Nodo_Timbrado.setAttribute "selloSAT", Timbrado_selloSAT
            'Asigna el nodo timbrado
            Nodo_Complemento.appendChild Nodo_Timbrado
            CFD_Generales.Fecha_Timbrado = Timbrado_FechaTimbrado
        Else 'Si no es valida regresa el error
            'Obtiene la respuesta
            Set Respuesta = Conexion_Web_Services.SelectSingleNode("//Error")
            If Not Respuesta Is Nothing Then
                MsgBox "Error " & Respuesta.Attributes(0).text & Chr(13) & Respuesta.Attributes(1).text
                Err.Raise 7777, "CFD_Crea_Xml", "Error: " & Respuesta.Attributes(0).text & " " & Respuesta.Attributes(1).text
            Else
                MsgBox "Error" & Chr(13) & "No es posible encontrar los parametros de respuesta"
                Err.Raise 7777, "CFD_Crea_Xml", "Error: No es posible encontrar los parametros de respuesta"
            End If
        End If
    Else 'Si no es valida regresa el error
        MsgBox "El procedimiento no pudo ser realizado" & Chr(13) & "      Favor de intentarlo nuevamente", vbCritical
        MsgBox XML
        Err.Raise 7777, "CFD_Crea_Xml", XML
    End If
                
    
    'Guarda el archivo
    If CFD_Generales.Tipo_Comprobante = "ingreso" Then
        'Valida exista la carpeta donde se gueardar el archivo, de lo contrario la crea
        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Ruta_Xmls, "CARPETA") = False Then
            MkDir Ruta_Xmls & "\"
        End If
        CFD_Documento.Save Ruta_Xmls & "\" & Nombre_CFD & ".xml"
    Else
        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Ruta_NC, "CARPETA") = False Then
            MkDir Ruta_NC & "\"
        End If
        CFD_Documento.Save Ruta_NC & "\" & Nombre_CFD & ".xml"
    End If
    'Destruye el documento
    Set CFD_Documento = Nothing
    
    'Genera el codigo bidimensinal
    Codigo_Bidimensional = "?re=" & CFD_Emisor.RFC & "&rr=" & CFD_Receptor.RFC & "&tt=" & Format(CFD_Generales.Total, "0000000000.000000") & "&id=" & Timbrado_UUID
    If CFD_Generales.Tipo_Comprobante = "ingreso" Then
        Call GenerateFile(Codigo_Bidimensional, Ruta_Pdfs & "\CFDI_" & Trim(CFD_Generales.Serie & "_" & CFD_Generales.Folio) & ".bmp")
    Else
        Call GenerateFile(Codigo_Bidimensional, Ruta_NC & "\CFDI_" & Trim(CFD_Generales.Serie & "_" & CFD_Generales.Folio) & ".bmp")
    End If
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: CFD_Crea_PDF
'DESCRIPCIÓN: Crea el archivo PDF de acuerdo a los datos para generar el CFD
'PARÁMETROS:
'               1. Nombre_CFD, toma el nombre del archivo PDF
'               2. Documento, asigna el tipo de document factura o nota de credito
'CREO: Ismael Prieto Sánchez
'FECHA_CREO: 20/Oct/2006 11:50 am
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Sub CFD_Crea_PDF(Nombre_CFD As String, Documento As String, Tipo As String, Año As Integer)
Dim crxApplication As New CRAXDDRT.Application
Dim crxReport As CRAXDDRT.report
Dim crxDatabase As CRAXDDRT.Database
Dim crxDatabaseTables As CRAXDDRT.DatabaseTables
Dim crxDatabaseTable As CRAXDDRT.DatabaseTable
Dim crxSections As CRAXDDRT.Sections
Dim crxSection As CRAXDDRT.Section
Dim crxSubreport As CRAXDDRT.report
Dim crxSubreportObject As SubreportObject
Dim crParamDefs As CRAXDDRT.ParameterFieldDefinitions
Dim crParamDef As CRAXDDRT.ParameterFieldDefinition
Dim Cuenta_Tablas As Integer
Dim Numero As String
Dim Flete As String
Dim Nota As String
    
On Error GoTo handler
    'Asigna el formato de la factura a la variable
If Documento = "PAGOS" Then
    Set crxReport = crxApplication.OpenReport(App.Path & "\REP\Rpt_Formato_Complemento_Pagos.rpt")
Else
    If Documento = "FACTURA" Then
        Set crxReport = crxApplication.OpenReport(App.Path & "\REP\Rpt_Formato_Factura.rpt")
    Else
        If Documento = "REMISION" Then
            Set crxReport = crxApplication.OpenReport(App.Path & "\REP\Rpt_Formato_Remision.rpt")
        Else
            If Documento = "NOTA CARGO" Then
                Set crxReport = crxApplication.OpenReport(App.Path & "\REP\Rpt_Formato_Nota_Cargo.rpt")
            Else
                Set crxReport = crxApplication.OpenReport(App.Path & "\REP\Rpt_Formato_Nota_Credito.rpt")
            End If
        End If
    End If
End If
    'No guarda los datos en el reporte
    crxReport.DiscardSavedData
    
    'Asigna los datos de conexion de la base de datos
    With crxReport
        For Cuenta_Tablas = 1 To .Database.Tables.Count
            Select Case Replace(.Database.Tables(Cuenta_Tablas).DllName, ".dll", "")
                Case "pdsodbc", "crdb_odbc", "crdb_dao"
                    'Primero es el nombre del ODBC y despues el nombre de la base de datos
                     .Database.Tables(Cuenta_Tablas).SetLogOnInfo "Alcoholera", Nombre_BD, Usuario_BD, Password_BD
            End Select
        Next
    End With
    
    'Asigna los datos a los parametros
    Set crParamDefs = crxReport.ParameterFields
    For Each crParamDef In crParamDefs
    If Documento = "PAGOS" Then
        Select Case crParamDef.ParameterFieldName
            Case "No_Factura"
                crParamDef.AddCurrentValue (Format(Val(CFD_Generales.Folio), "0000000000"))
            Case "Serie"
                crParamDef.AddCurrentValue (CFD_Generales.Serie)
            Case "Importe_Letra"
                crParamDef.AddCurrentValue (CFD_Generales.Importe_Letra)
            Case "Cadena_Original"
                crParamDef.AddCurrentValue (CFD_Generales.Cadena_Original)
            Case "Sello_Digital"
                crParamDef.AddCurrentValue (CFD_Generales.Sello)
            Case "Mensaje_Factura"

               crParamDef.AddCurrentValue ""
            Case "Expedido_En"
                crParamDef.AddCurrentValue (CFD_Emisor.Expedido_En)
            Case "Empresa"
                crParamDef.AddCurrentValue (CFD_Emisor.Nombre)
            Case "RFC"
                crParamDef.AddCurrentValue (CFD_Emisor.RFC)
            Case "Direccion"
                crParamDef.AddCurrentValue (CFD_Emisor.Calle)
            Case "Numero_Exterior"
                crParamDef.AddCurrentValue (CFD_Emisor.No_Exterior)
            Case "No_Interior"
                crParamDef.AddCurrentValue (CFD_Emisor.No_Interior)
            Case "Colonia"
                crParamDef.AddCurrentValue (CFD_Emisor.Colonia)
            Case "CP"
                crParamDef.AddCurrentValue (CFD_Emisor.cp)
            Case "Ciudad"
                crParamDef.AddCurrentValue (CFD_Emisor.Municipio)
            Case "Estado"
                crParamDef.AddCurrentValue (CFD_Emisor.Estado)
            Case "Leyenda_Factura"
                crParamDef.AddCurrentValue "ESTE DOCUMENTO ES UNA REPRESENTACION IMPRESA DE UN CFDI"
            Case "Fecha_Timbrado"
                crParamDef.AddCurrentValue (CFD_Generales.Fecha_Timbrado)
            Case "Imagen_Codigo"
                crParamDef.AddCurrentValue (App.Path & "\CFDS\CFDI_" & CFD_Generales.Serie & "_" & CFD_Generales.Folio & ".bmp")
            Case "Forma_Pago"
                crParamDef.AddCurrentValue CFD_Pagos.Forma_Pago_Pagos
            Case "Regimen_Fiscal"
                crParamDef.AddCurrentValue Trim(Regimen_Emisor)
            Case "Ruta_Codigo_CBB"
                crParamDef.AddCurrentValue (App.Path & "\CFDS\CFDI_" & CFD_Generales.Serie & "_" & Val(CFD_Generales.Folio) & ".bmp")
            Case "Factura_Origen"
                crParamDef.AddCurrentValue (CFD_Pagos_DR(1).Folio)
            Case "Informacion_Adicional"
                crParamDef.AddCurrentValue ("")
            Case "No_Interior"
                crParamDef.AddCurrentValue ("")
            Case "Comentarios"
                crParamDef.AddCurrentValue Trim(CFD_Generales.Condiciones_Pago)
             End Select
    Else
        If Documento = "FACTURA" Or Documento = "NOTA CARGO" Then
            If Tipo = "NORMAL" Then
                Select Case crParamDef.ParameterFieldName
                    Case "No_Factura"
                        crParamDef.AddCurrentValue (CFD_Generales.Factura_ID)
                    Case "Serie"
                        crParamDef.AddCurrentValue (CFD_Generales.Serie)
                    Case "Importe_Letra"
                        crParamDef.AddCurrentValue (CFD_Generales.Importe_Letra)
                    Case "Cadena_Original"
                        crParamDef.AddCurrentValue (CFD_Generales.Cadena_Original)
                    Case "Sello_Digital"
                        crParamDef.AddCurrentValue (CFD_Generales.Sello)
                    Case "Mensaje_Factura"
                        crParamDef.AddCurrentValue (Texto_Factura)
                    Case "Expedido_En"
                        crParamDef.AddCurrentValue (Expedida_En)
                    Case "Interes_Moratorio"
                        crParamDef.AddCurrentValue (Tasa_Interes_Pagare)
                    Case "No_Documento"
                        Numero = CFD_Generales.Serie & " " & (CFD_Generales.Folio)
                        crParamDef.AddCurrentValue Numero
                    Case "Fecha_Vencimiento"
                        crParamDef.AddCurrentValue (CFD_Generales.Fecha_Vencimiento)
                    Case "Empresa"
                        crParamDef.AddCurrentValue (Nombre_Emisor)
                    Case "RFC"
                        crParamDef.AddCurrentValue (RFC_Emisor)
                    Case "Direccion"
                        crParamDef.AddCurrentValue (Calle_Emisor)
                    Case "Numero_Exterior"
                        crParamDef.AddCurrentValue (No_Exterior_Emisor)
                    Case "No_Interior"
                        crParamDef.AddCurrentValue (No_Interior_Emisor)
                    Case "Colonia"
                        crParamDef.AddCurrentValue (Colonia_Emisor)
                    Case "CP"
                        crParamDef.AddCurrentValue (Codigo_Postal_Emisor)
                    Case "Ciudad"
                        crParamDef.AddCurrentValue (Municipio_Emisor)
                    Case "Estado"
                        crParamDef.AddCurrentValue (Estado_Emisor)
                    Case "IVA"
                        crParamDef.AddCurrentValue Trim(PG_Retencion_IVA * 100)
                    Case "Leyenda_Pago"
                        crParamDef.AddCurrentValue UCase(CFD_Generales.Forma_Pago)
                    Case "Leyenda_Factura"
                        crParamDef.AddCurrentValue "ESTE DOCUMENTO ES UNA REPRESENTACION IMPRESA DE UN CFDI"
                    Case "Orden_Compra"
                        crParamDef.AddCurrentValue (CFD_Generales.Orden_Compra)
                    Case "IVA_Desglosado"
                        crParamDef.AddCurrentValue (CFD_Generales.IVA_Desglosado)
                    Case "Fecha_Timbrado"
                        crParamDef.AddCurrentValue (CFD_Generales.Fecha_Timbrado)
                    Case "Imagen_Codigo"
                        crParamDef.AddCurrentValue (CFD_Generales.Imagen_BMP)
                    Case "Forma_Pago"
                        crParamDef.AddCurrentValue CFD_Generales.Forma_Pago
                    Case "Metodo_Pago"
                        crParamDef.AddCurrentValue CFD_Generales.Metodo_Pago
                    Case "Regimen_Fiscal"
                        crParamDef.AddCurrentValue (Regimen_Emisor)
                    
                End Select
            End If
        Else
            If Documento = "REMISION" Then
                Select Case crParamDef.ParameterFieldName
                    Case "No_Salida"
                        crParamDef.AddCurrentValue (CFD_Generales.No_Salida)
                    Case "Empresa"
                        crParamDef.AddCurrentValue (Nombre_Emisor)
                    Case "RFC"
                        crParamDef.AddCurrentValue (RFC_Emisor)
                    Case "Direccion"
                        crParamDef.AddCurrentValue (Calle_Emisor)
                    Case "Numero_Exterior"
                        crParamDef.AddCurrentValue (No_Exterior_Emisor)
                    Case "No_Interior"
                        crParamDef.AddCurrentValue (No_Interior_Emisor)
                    Case "Colonia"
                        crParamDef.AddCurrentValue (Colonia_Emisor)
                    Case "CP"
                        crParamDef.AddCurrentValue (Codigo_Postal_Emisor)
                    Case "Ciudad"
                        crParamDef.AddCurrentValue (Municipio_Emisor)
                    Case "Estado"
                        crParamDef.AddCurrentValue (Estado_Emisor)
                    Case "Regimen_Fiscal"
                        crParamDef.AddCurrentValue (Regimen_Emisor)
                    Case "Almacenista"
                        crParamDef.AddCurrentValue (Frm_Adm_Clientes_Facturas.Txt_Nombre_Almacenista.text)
                    Case "Recibe_Mercancia"
                        crParamDef.AddCurrentValue (Frm_Adm_Clientes_Facturas.Txt_Recibe.text)
                End Select
            Else
                'Es una nota de credito
                Select Case crParamDef.ParameterFieldName
                    Case "No_Nota_Credito"
                        crParamDef.AddCurrentValue (Format(Val(CFD_Generales.Folio), "0000000000"))
                    Case "Serie"
                        crParamDef.AddCurrentValue (CFD_Generales.Serie)
                    Case "Importe_Letra"
                        crParamDef.AddCurrentValue (CFD_Generales.Importe_Letra)
                    Case "Cadena_Original"
                        crParamDef.AddCurrentValue (CFD_Generales.Cadena_Original)
                    Case "Sello_Digital"
                        crParamDef.AddCurrentValue (CFD_Generales.Sello)
                    Case "Nota"
                        Nota = (CFD_Generales.Serie) & " " & (Val(CFD_Generales.Folio))
                        crParamDef.AddCurrentValue Nota
                    Case "Expedido_En"
                        crParamDef.AddCurrentValue (Expedida_En)
                    Case "Empresa"
                        crParamDef.AddCurrentValue (Nombre_Emisor)
                    Case "RFC"
                        crParamDef.AddCurrentValue (RFC_Emisor)
                    Case "Direccion"
                        crParamDef.AddCurrentValue (Calle_Emisor)
                    Case "Numero_Exterior"
                        crParamDef.AddCurrentValue (No_Exterior_Emisor)
                    Case "No_Interior"
                        crParamDef.AddCurrentValue (No_Interior_Emisor)
                    Case "Colonia"
                        crParamDef.AddCurrentValue (Colonia_Emisor)
                    Case "CP"
                        crParamDef.AddCurrentValue (Codigo_Postal_Emisor)
                    Case "Ciudad"
                        crParamDef.AddCurrentValue (Municipio_Emisor)
                    Case "Estado"
                        crParamDef.AddCurrentValue (Estado_Emisor)
                    Case "Leyenda_Factura"
                        crParamDef.AddCurrentValue "ESTE DOCUMENTO ES UNA REPRESENTACION IMPRESA DE UN CFDI"
                    Case "IVA"
                        crParamDef.AddCurrentValue Trim(CFD_Generales.Impuestos)
                    Case "Subtotal"
                        crParamDef.AddCurrentValue (CFD_Generales.SubTotal)
                    Case "Total"
                        crParamDef.AddCurrentValue (CFD_Generales.Total)
                    Case "Fecha_Timbrado"
                        crParamDef.AddCurrentValue (CFD_Generales.Fecha_Timbrado)
                    Case "Regimen_Fiscal"
                        crParamDef.AddCurrentValue (Regimen_Emisor)
                    Case "Imagen_Codigo"
                        crParamDef.AddCurrentValue (CFD_Generales.Imagen_BMP)
                    Case "Leyenda_Pago"
                        crParamDef.AddCurrentValue UCase(CFD_Generales.Forma_Pago)
                End Select
            End If
        End If
    End If
    Next
   
    
    'Asigna los datos de exportación
    crxReport.ExportOptions.DestinationType = crEDTDiskFile
    If Documento = "FACTURA" Or Documento = "NOTA CARGO" Then
        'Valida que exista la carpeta donde se almacenaran los pdf, si no existe la crea
        If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Ruta_Pdfs, "CARPETA") = False Then
            MkDir Ruta_Pdfs & "\"
        End If
        crxReport.ExportOptions.DiskFileName = Ruta_Pdfs & "\" & Nombre_CFD & ".pdf"
    Else
        If Documento = "REMISION" Then
            If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Ruta_Remisiones, "CARPETA") = False Then
                MkDir Ruta_Remisiones & "\"
            End If
            crxReport.ExportOptions.DiskFileName = Ruta_Remisiones & "\" & Nombre_CFD & ".pdf"
        Else
            'Es un pago
            If Documento = "PAGOS" Then
                If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Ruta_Pdfs, "CARPETA") = False Then
                    MkDir Ruta_Pdfs & "\"
                End If
                crxReport.ExportOptions.DiskFileName = Ruta_Pdfs & "\" & Nombre_CFD & ".pdf"
            Else
            'Es una nota de crédito
            If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Ruta_NC, "CARPETA") = False Then
                MkDir Ruta_NC & "\"
            End If
            crxReport.ExportOptions.DiskFileName = Ruta_NC & "\" & Nombre_CFD & ".pdf"
         End If
        End If
    End If
    crxReport.ExportOptions.FormatType = crEFTPortableDocFormat
    crxReport.ExportOptions.PDFExportAllPages = True
   
    'Oculta el progreso de la exportacion
    crxReport.DisplayProgressDialog = False
    
    'Genera la exportación del documento
    crxReport.Export (False)
    
    'Destruye el documento
    Set crxReport = Nothing
Exit Sub
handler:
    For Each Er In rdoErrors
        MsgBox Er.Description
Next Er
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: CFD_Valida_Caracteres_UTF
'DESCRIPCIÓN: Realiza la validación de los caracteres para convertirlos a UTF-8
'PARÁMETROS:
'               1. Cadena_Original, para pasarle la cadena original
'               2. Regresa el valor de la cadena original convertida
'CREO: Ismael Prieto Sánchez
'FECHA_CREO: 21/Oct/2006 11:00 am
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function CFD_Valida_Caracteres_UTF(Cadena_Original As String) As String
Dim strData As String
Dim strDataUTF8 As String
Dim strDigest As String
Dim nRet As Long
Dim nLen As Long
    
    'Asigna en forma temporal la cadena original
    strData = Cadena_Original
    
    'Convierte a UTF-8
    nLen = CNV_UTF8FromLatin1("", 0, strData)
    If nLen < 0 Then
        Err.Raise 7777, "CFD_Valida_Caracteres_UTF", "No se puede convertir la cadena a UTF-8, favor de verificarlo"
    End If
    strDataUTF8 = String(nLen, " ")
    nLen = CNV_UTF8FromLatin1(strDataUTF8, nLen, strData)
    
    strDataUTF8 = strData
    
    'Regresa la cadena convertida a la función
    CFD_Valida_Caracteres_UTF = strDataUTF8
End Function


'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: CFD_Genera_Sello
'DESCRIPCIÓN: Genera el sello digital de la factura electronica
'PARÁMETROS:
'               1. Cadena_MD5, para pasarle la cadena encriptada con MD5
'               2. Nombre_Llave, para pasarle el valor de la llave privada para la encripción del sello
'               3. Regresa el valor del sello digital
'CREO: Ismael Prieto Sánchez
'FECHA_CREO: 21/Oct/2006 11:00 am
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function CFD_Genera_Sello(Cadena_MD5 As String, Nombre_Llave As String) As String
    Const MD5_BYTES_LEN As Long = 16
    Dim strKeyFile As String
    Dim strPrivateKey As String
    Dim strFileName As String
    Dim abBlock() As Byte
    Dim nBlockLen As Long
    Dim nLen As Long
    Dim mLen As Long
    Dim nRet As Long
    Dim strBase64 As String
    Dim strData As String
    Dim strDataUTF8 As String
    Dim abMessage() As Byte
    Dim strPassword As String
    
    'Lee la llave privada encriptada en el archivo
    strKeyFile = Nombre_Llave
    strPassword = Password_Llave
    strPrivateKey = rsaReadPrivateKey(strKeyFile, strPassword)
    If Len(strPrivateKey) = 0 Then
        Err.Raise 7777, "CFD_Genera_Sello", "No se puede leer la llave privada, favor de verificarlo"
    Else
        Debug.Print "Key length is " & RSA_KeyBits(strPrivateKey) & " bits/" & RSA_KeyBytes(strPrivateKey) & " bytes"
    End If
    
    'Asigna la cadena en MD5
    strData = Cadena_MD5
    
    'Convierte la cadena a UTF-8
    nLen = CNV_UTF8FromLatin1("", 0, strData)
    If nLen < 0 Then
        Err.Raise 7777, "CFD_Genera_Sello", "No se puede convertir la cadena a UTF-8, favor de verificarlo"
    End If
    strDataUTF8 = String(nLen, " ")
    nLen = CNV_UTF8FromLatin1(strDataUTF8, nLen, strData)
    
    'Convierte la cadena en un arreglo para trabajar con VB
    abMessage = StrConv(strDataUTF8, vbFromUnicode)
    mLen = UBound(abMessage) - LBound(abMessage) + 1
    
    ' Encode ready for signing with `Encoded Message for Signature' block using PKCS#1 v1.5 method
    ' and MD5 o SHA1
    nBlockLen = RSA_KeyBytes(strPrivateKey)
    ReDim abBlock(nBlockLen - 1)
     'Obtiene la cadena en MD5
'    If Año <= 2010 Then
'        nLen = RSA_EncodeMsg(abBlock(0), nBlockLen, abMessage(0), mLen, PKI_EMSIG_DEFAULT + PKI_HASH_MD5)
'    Else
    'Obtiene la cadena en SHA1
    nLen = RSA_EncodeMsg(abBlock(0), nBlockLen, abMessage(0), mLen, PKI_EMSIG_PKCSV1_5 + PKI_HASH_SHA256)
'    End If

    If nLen < 0 Then
        Err.Raise 7777, "CFD_Genera_Sello", "No se puede codificar la cadena, favor de verificarlo"
    End If
    
    ' Now sign using the RSA private key
    nRet = RSA_RawPrivate(abBlock(0), nBlockLen, strPrivateKey, 0)
    
    'Convierte la cadena a base64
    strBase64 = cnvB64StrFromBytes(abBlock)
    
    'Limpia los datos sensibles para que no haya errores
    Call WIPE_String(strPassword, Len(strPassword))
    strPassword = ""
    Call WIPE_String(strPrivateKey, Len(strPrivateKey))
    strPrivateKey = ""
    
    'Devuelve el valor a la funcion
    CFD_Genera_Sello = strBase64
End Function

Public Function CFD_Genera_Sello_Nuevo(Cadena_UTF8 As String, Nombre_Llave As String) As String
Dim strDataFile As String
Dim strKeyFile As String
Dim strPassword As String
Dim strPrivateKey As String
Dim abDigest() As Byte
Dim abBlock() As Byte
Dim nBlockLen As Long
Dim nLen As Long
Dim nRet As Long
Dim strBase64 As String

    ' INPUT: File containing piped-string formed from XML doc in UTF-8 format (NOTE: no Unicode markers in the file),
    '        private key file and its secret password(!)
    strDataFile = Cadena_UTF8
    strKeyFile = Nombre_Llave
    ' Test password - CAUTION: DO NOT hardcode production passwords!
    strPassword = Password_Llave
    
    ' 1. Form the message digest hash of the piped-string directly from the file
    
    ReDim abDigest(PKI_MD5_BYTES - 1)
    nRet = HASH_File(abDigest(0), PKI_MD5_BYTES, strDataFile, PKI_HASH_MD5)
    Debug.Print "HASH_File returns " & nRet
    If nRet <= 0 Then
        Err.Raise 7777, "CFD_Genera_Sello", "No se puede crear el archivo hash, favor de verificarlo."
    End If
    ' Display in hex
    Debug.Print "Digest=" & cnvHexStrFromBytes(abDigest)
    
    ' 2. Sign the message digest using the private key
    
    ' 2.1 Read in private key from encrypted .key file
    strPrivateKey = rsaReadPrivateKey(strKeyFile, strPassword)
    If Len(strPrivateKey) = 0 Then
        Err.Raise 7777, "CFD_Genera_Sello", "No se puede leer la llave privada, favor de verificarlo."
    End If
    ' -- show we got something
    Debug.Print "Private key is " & RSA_KeyBits(strPrivateKey) & " bits long"
    
    ' 2.2 Encode the digest ready for signing with `Encoded Message for Signature' block using PKCS#1 v1.5 method
    nBlockLen = RSA_KeyBytes(strPrivateKey)
    ReDim abBlock(nBlockLen - 1)
    nLen = RSA_EncodeMsg(abBlock(0), nBlockLen, abDigest(0), PKI_MD5_BYTES, PKI_EMSIG_DEFAULT + PKI_EMSIG_DIGESTONLY + PKI_HASH_MD5)
    If nLen < 0 Then
        Err.Raise 7777, "CFD_Genera_Sello", "Error en la codificación RSA, favor de verificarlo."
    End If
    Debug.Print "INPUT BLOCK= " & cnvHexStrFromBytes(abBlock)
    
    ' 2.3 Sign using the RSA private key
    nRet = RSA_RawPrivate(abBlock(0), nBlockLen, strPrivateKey, 0)
    ' Display in hex
    Debug.Print "OUTPUT BLOCK=" & cnvHexStrFromBytes(abBlock)
    
    ' 2.4 Clean up
    Call WIPE_String(strPassword, Len(strPassword))
    strPassword = ""
    Call WIPE_String(strPrivateKey, Len(strPrivateKey))
    strPrivateKey = ""
    
    ' 3. Convert to base64 and output result
    strBase64 = cnvB64StrFromBytes(abBlock)
    Debug.Print "SIGNATURE VALUE=" & strBase64

    'Devuelve el valor a la funcion
    CFD_Genera_Sello_Nuevo = strBase64
End Function


'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: CFD_Consulta_Serie_Certificado
'DESCRIPCIÓN: Consulta el no. de serie del certificado
'PARÁMETROS:
'               1. Nombre_Certificado, para pasarle el nombre del certificado
'               2. Regresa el valor del número de serie
'CREO: Ismael Prieto Sánchez
'FECHA_CREO: 21/Oct/2006 11:00 am
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function CFD_Consulta_Serie_Certificado(Nombre_Certificado As String) As String
Dim nLen As Long
Dim strCertFile As String
Dim strSerialNumber As String
Dim strSerialSAT As String

    'Extrae el numero de serie del certificado
    strCertFile = Nombre_Certificado
    nLen = X509_CertSerialNumber(strCertFile, "", 0, 0)
    If (nLen <= 0) Then
        Err.Raise 7777, "CFD_Consulta_Serie_Certificado", "No se puede leer el certificado, favor de verificarlo."
    End If
    strSerialNumber = String(nLen, " ")
    nLen = X509_CertSerialNumber(strCertFile, strSerialNumber, Len(strSerialNumber), 0)
    
    'Decodifica de hex-encoded integer a cadena ASCII de digitos
    strSerialSAT = StrConv(cnvBytesFromHexStr(strSerialNumber), vbUnicode)

    'Regresa el valor a la funcion
    CFD_Consulta_Serie_Certificado = strSerialSAT
End Function

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: CFD_Consulta_Certificado
'DESCRIPCIÓN: Consulta el no. de serie del certificado
'PARÁMETROS:
'               1. Nombre_Certificado, para pasarle el nombre del certificado
'               2. Regresa el valor del número de serie
'CREO: Ismael Prieto Sánchez
'FECHA_CREO: 21/Oct/2006 11:00 am
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function CFD_Consulta_Certificado(Nombre_Certificado As String) As String
Dim nRet As Long
Dim strCertString As String
Dim strCertFile As String
Dim strThumb1 As String
Dim strThumb2 As String
    
    'Asigna el nombre del certificado
    strCertFile = Nombre_Certificado
    
    'Lee el certificado en forma de cadena
    nRet = X509_ReadStringFromFile("", 0, strCertFile, 0)
    Debug.Print "X509_ReadStringFromFile returns " & nRet
    If nRet <= 0 Then
        Err.Raise 7777, "CFD_Consulta_Certificado", "ERROR: No es posible leer el certificado, favor de verificarlo, no. error: " & nRet
    End If
    strCertString = String(nRet, " ")
    nRet = X509_ReadStringFromFile(strCertString, Len(strCertString), strCertFile, 0)
    Debug.Print "For certificate '" & strCertFile & "':"
    Debug.Print strCertString
    
    'Checa si la version es valida
    strThumb1 = String(PKI_SHA1_CHARS, " ")
    strThumb2 = String(PKI_SHA1_CHARS, " ")
    nRet = X509_CertThumb(strCertFile, strThumb1, Len(strThumb1), 0)
    nRet = X509_CertThumb(strCertString, strThumb2, Len(strThumb2), 0)
    Debug.Print "SHA-1(file)  =" & strThumb1
    Debug.Print "SHA-1(string)=" & strThumb2
    If strThumb1 <> strThumb2 Then
        Err.Raise 7777, "CFD_Consulta_Certificado", "ERROR: El certificado esta alterado favor de verificarlo."
    End If

    'Regresa el valor a la funcion
    CFD_Consulta_Certificado = strCertString
End Function


'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: CFD_Consulta_Certificado
'DESCRIPCIÓN: Consulta el no. de serie del certificado
'PARÁMETROS:
'               1. Nombre_Certificado, para pasarle el nombre del certificado
'               2. Regresa el valor del número de serie
'CREO: Ismael Prieto Sánchez
'FECHA_CREO: 21/Oct/2006 11:00 am
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function CFD_Consulta_Certificado_Net(Nombre_Certificado As String) As String
Dim nRet As Long
Dim strCertString As String
Dim strCertFile As String
Dim strThumb1 As String
Dim strThumb2 As String
Dim Obj As Object

    'Asigna el nombre del certificado
    strCertFile = Nombre_Certificado
    
    Set Obj = CreateObject("FacturaElectronicaMetodos")
    
    'Lee el certificado en forma de cadena
    strCertString = Net_Certificado_No_Serie(Nombre_Certificado)

    'Regresa el valor a la funcion
    CFD_Consulta_Certificado_Net = strCertString
End Function



'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: CFD_Consulta_Vigencia_Desde
'DESCRIPCIÓN: Consulta la vigencia inicial del certificado
'PARÁMETROS:
'               1. Nombre_Certificado, para pasarle el nombre del certificado
'               2. Regresa el valor de la vigencia
'CREO: Ismael Prieto Sánchez
'FECHA_CREO: 01/Sep/2009 10:10 am
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function CFD_Consulta_Vigencia_Desde(Nombre_Certificado As String) As Date
Dim nLen As Long
Dim strCertFile As String
Dim strVigencia As String
Dim strVigenciaSAT As String
Dim Pos As Integer

    strCertFile = Nombre_Certificado
    'Extrae la fecha inicial de vigencia
    nLen = X509_CertIssuedOn(strCertFile, "", 0, 0)
    If (nLen <= 0) Then
        Err.Raise 7777, "CFD_Consulta_Vigencia_Desde", "No se puede leer el certificado, favor de verificarlo."
    End If
    strVigencia = String(nLen, " ")
    nLen = X509_CertIssuedOn(strCertFile, strVigencia, Len(strVigencia), 0)
    
    'Formatea la vigencia
    Pos = InStr(1, strVigencia, "T")
    
    strVigenciaSAT = Mid(strVigencia, 1, Pos - 1) & " " & Mid(strVigencia, Pos + 1, Len(strVigencia))
    strVigenciaSAT = Mid(strVigenciaSAT, 1, Len(strVigenciaSAT) - 1)
    
    'Regresa el valor a la funcion
    CFD_Consulta_Vigencia_Desde = CDate(strVigenciaSAT)
End Function

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: CFD_Consulta_Vigencia_Hasta
'DESCRIPCIÓN: Consulta la vigencia final del certificado
'PARÁMETROS:
'               1. Nombre_Certificado, para pasarle el nombre del certificado
'               2. Regresa el valor de la vigencia
'CREO: Ismael Prieto Sánchez
'FECHA_CREO: 01/Sep/2009 10:10 am
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function CFD_Consulta_Vigencia_Hasta(Nombre_Certificado As String) As Date
Dim nLen As Long
Dim strCertFile As String
Dim strVigencia As String
Dim strVigenciaSAT As String
Dim Pos As Integer

    strCertFile = Nombre_Certificado
    'Extrae la fecha inicial de vigencia
    nLen = X509_CertExpiresOn(strCertFile, "", 0, 0)
    If (nLen <= 0) Then
        Err.Raise 7777, "CFD_Consulta_Vigencia_Hasta", "No se puede leer el certificado, favor de verificarlo."
    End If
    strVigencia = String(nLen, " ")
    nLen = X509_CertExpiresOn(strCertFile, strVigencia, Len(strVigencia), 0)
    
    'Formatea la vigencia
    Pos = InStr(1, strVigencia, "T")
    
    strVigenciaSAT = Mid(strVigencia, 1, Pos - 1) & " " & Mid(strVigencia, Pos + 1, Len(strVigencia))
    strVigenciaSAT = Mid(strVigenciaSAT, 1, Len(strVigenciaSAT) - 1)
    
    'Regresa el valor a la funcion
    CFD_Consulta_Vigencia_Hasta = CDate(strVigenciaSAT)
End Function


''Public Sub CFD_Valida_Esquema(Nombre_CFD As String)
''Dim xmlschema As MSXML2.XMLSchemaCache40
''
''  Set xmlschema = New MSXML2.XMLSchemaCache40
''  xmlschema.Add "http://www.sat.gob.mx/cfd/2", App.Path & "\comprobante.xsd"
''
''  'Create an XML DOMDocument object.
''  Dim xmldom As MSXML2.DOMDocument40
''  Set xmldom = New MSXML2.DOMDocument40
''
''  'Assign the schema cache to the DOM document.
''  'schemas collection.
''  Set xmldom.schemas = xmlschema
''
''  'Load books.xml as the DOM document.
''  xmldom.async = False
''  xmldom.Load App.Path & "\" & Nombre_CFD
''
''  'Return validation results in message to the user.
''  If xmldom.parseError.errorCode <> 0 Then
''     MsgBox xmldom.parseError.errorCode & " " & _
''     xmldom.parseError.reason
''  Else
''     MsgBox "No Error"
''  End If
''End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: CFD_Crea_Xml
'DESCRIPCIÓN: Crea el archivo xml a la estructura solicitada por el SAT
'PARÁMETROS:  Nombre_CFD - Nombre con que se guardará el XML
'CREO:        Sergio Godínez Banda
'FECHA_CREO:  25-Mayo-2012
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Sub CFD_Crea_Xml(Nombre_CFD As String, Tipo_Documento As String, Tipo_Adenda As String) '(Nombre_CFD As String) ', Tipo_Documento As String) ', Tipo_Adenda As String)
Dim Xml_Identificacion As IXMLDOMProcessingInstruction
Dim Nodo_Raiz As IXMLDOMElement
Dim Nodo_Emisor As IXMLDOMElement
Dim Nodo_Emisor_Domicilio As IXMLDOMElement
Dim Nodo_Relacionados As IXMLDOMElement
Dim Nodo_Relacionados_UUID As IXMLDOMElement
Dim Nodo_ExpedidoEn As IXMLDOMElement
Dim Nodo_Regimen As IXMLDOMElement
Dim Nodo_Receptor As IXMLDOMElement
Dim Nodo_Receptor_Domicilio As IXMLDOMElement
Dim Nodo_Conceptos As IXMLDOMElement
Dim Nodo_Concepto As IXMLDOMElement
Dim Nodo_Impuestos As IXMLDOMElement
Dim Nodo_Traslados As IXMLDOMElement
Dim Nodo_Traslado As IXMLDOMElement
Dim Nodo_Retenciones As IXMLDOMElement
Dim Nodo_Retencion As IXMLDOMElement
Dim Nodo_Complemento As IXMLDOMElement
Dim Nodo_ImpuestoLocal As IXMLDOMElement
Dim Nodo_Elemento As IXMLDOMElement
Dim Nodo_Timbrado As IXMLDOMElement
Dim Nodo_Cuenta_Predial As IXMLDOMElement
Dim Nodo_Donaciones As IXMLDOMElement
Dim Nodo_Impuesto_Concepto As IXMLDOMElement
Dim Nodo_Traslado_Concepto As IXMLDOMElement
Dim Nodo_Traslado_Conceptos As IXMLDOMElement
Dim Nodo_Retencion_Concepto As IXMLDOMElement
Dim Nodo_Retencion_Conceptos As IXMLDOMElement
Dim Nodo_Pago As IXMLDOMElement
Dim Nodo_Pagos As IXMLDOMElement
Dim Nodo_Pagos_Relacionados As IXMLDOMElement
Dim Nodo_Aduanera As IXMLDOMElement
Dim Impuestos_Trasladados As Double
Dim Impuestos_Retenidos As Double
Dim Impuestos_Locales As Double
Dim Impuestos_Locales_Retenidos As Double
Dim Impuestos_Locales_Trasladados As Double
Dim Conteo_Estructura As Integer
Dim Conteo_Estructura2 As Integer
Dim Cadena_XML As String
Dim Genera_XML As String
Dim Proceso_Web_Services As String
Dim Respuesta_Web_Services As String
Dim Respuesta As IXMLDOMNode
Dim CFDI As String
Dim XML As String
Dim i As Long
Dim Nodo_Complemento_Comercio_Exterior As IXMLDOMElement
Dim Nodo_Complemento_Comercio_Exterior_Mercancias As IXMLDOMElement
Dim Nodo_Complemento_Comercio_Exterior_Mercancia As IXMLDOMElement
Dim Nodo_Complemento_Comercio_Exterior_Receptor As IXMLDOMElement
Dim Nodo_Complemento_Comercio_Exterior_Emisor As IXMLDOMElement
Dim Nodo_Complemento_Comercio_Exterior_Destinatario As IXMLDOMElement
Dim Nodo_Complemento_Comercio_Exterior_Domicilio As IXMLDOMElement
'Dim Timbrado As New ContelTimbrado.GenerarTimbrado
Dim Timbrado33 As ICls_Fe_Timbrado
Set Timbrado33 = New Cls_Fe_Timbrado
'Timbrado33.
Dim Codigo_Bidimensional As String
Dim Respuesta_Timbrado As String
Dim Importe_IEPS As String
    '*****************************************************************************************
    'CREA DOCUMENTO XML
    '*****************************************************************************************
    'Crea el documento xml
    Set CFD_Documento = New DOMDocument
    'Identifica la version y codificacion del xml
    Set Xml_Identificacion = CFD_Documento.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
    CFD_Documento.appendChild Xml_Identificacion
    'Crea el documento xml
    Set Nodo_Raiz = CFD_Documento.createElement("cfdi:Comprobante")
    Set CFD_Documento.documentElement = Nodo_Raiz
        '*****************************************************************************************
        'CREA NODO PRINCIPAL
        '*****************************************************************************************
        'Crea el elemento principal
        Set Nodo_Principal = CFD_Documento.documentElement
            'Agrega los atributos a la raiz del xml
            Nodo_Principal.setAttribute "xsi:schemaLocation", "http://www.sat.gob.mx/cfd/3 http://www.sat.gob.mx/sitio_internet/cfd/3/cfdv33.xsd http://www.sat.gob.mx/implocal http://www.sat.gob.mx/sitio_internet/cfd/implocal/implocal.xsd http://www.sat.gob.mx/donat http://www.sat.gob.mx/sitio_internet/cfd/donat/donat11.xsd http://www.sat.gob.mx/ComercioExterior11 http://www.sat.gob.mx/sitio_internet/cfd/catalogos/ComExt/catComExt.xsd http://www.sat.gob.mx/Pagos http://www.sat.gob.mx/sitio_internet/cfd/Pagos/Pagos10.xsd"
''            Nodo_Principal.setAttribute "xsi:schemaLocation", "http://www.sat.gob.mx/cfd/3 http://www.sat.gob.mx/sitio_internet/cfd/3/cfdv32.xsd http://www.sat.gob.mx/implocal http://www.sat.gob.mx/sitio_internet/cfd/implocal/implocal.xsd
            Nodo_Principal.setAttribute "xmlns:implocal", "http://www.sat.gob.mx/implocal"
            Nodo_Principal.setAttribute "xmlns:donat", "http://www.sat.gob.mx/donat"
            Nodo_Principal.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
            Nodo_Principal.setAttribute "xmlns:cfdi", "http://www.sat.gob.mx/cfd/3"
            Nodo_Principal.setAttribute "xmlns:pago10", "http://www.sat.gob.mx/Pagos"
            Nodo_Principal.setAttribute "Version", CFD_Generales.Version
            If CFD_Generales.Serie <> "" Then
                Nodo_Principal.setAttribute "Serie", CFD_Generales.Serie
            End If
            Nodo_Principal.setAttribute "Folio", CFD_Generales.Folio
            Nodo_Principal.setAttribute "Fecha", CFD_Generales.Fecha
            Nodo_Principal.setAttribute "Sello", CFD_Generales.Sello
            If CFD_Generales.Forma_Pago <> "" Then
                Nodo_Principal.setAttribute "FormaPago", Mid(CFD_Generales.Forma_Pago, 1, 2) 'LCase(CFD_Generales.Forma_Pago)
            End If
            Nodo_Principal.setAttribute "NoCertificado", CFD_Generales.No_Certificado
            Nodo_Principal.setAttribute "Certificado", CFD_Generales.Certificado
            If CFD_Generales.Condiciones_Pago <> "" Then
                Nodo_Principal.setAttribute "CondicionesDePago", CFD_Generales.Condiciones_Pago
            End If
            Nodo_Principal.setAttribute "SubTotal", CDbl(Format(CFD_Generales.SubTotal, "#0.00"))
            If CFD_Generales.Descuento > 0 Then
                Nodo_Principal.setAttribute "Descuento", Format(CFD_Generales.Descuento, "#0.00")
            End If
            If CFD_Generales.Tipo_Moneda = "MXN" Or CFD_Generales.Tipo_Moneda = "XXX" Then
                Nodo_Principal.setAttribute "Moneda", CFD_Generales.Tipo_Moneda
            Else
                Nodo_Principal.setAttribute "TipoCambio", CFD_Generales.Tipo_Cambio
                Nodo_Principal.setAttribute "Moneda", CFD_Generales.Tipo_Moneda
            End If
            Nodo_Principal.setAttribute "Total", CDbl(Format(CFD_Generales.Total, "#0.00"))
            Nodo_Principal.setAttribute "TipoDeComprobante", CFD_Generales.Tipo_Comprobante
            'If CFD_Generales.Forma_Pago = "Pago en una sola exhibicion" Then
            '    Cadena_Original = Cadena_Original & "|" & "PUE"
            'Else
            '    Cadena_Original = Cadena_Original & "|" & "PPD"
            'End If
            If CFD_Generales.Metodo_Pago <> "" Then Nodo_Principal.setAttribute "MetodoPago", Mid(CFD_Generales.Metodo_Pago, 1, 3)
'            If CFD_Generales.Forma_Pago <> "" Then
'                If CFD_Generales.Forma_Pago = "Pago en una sola exhibicion" Then ' UNA SOLA EXHIBICION" Then
'                 Nodo_Principal.setAttribute "MetodoPago", "PUE"
'             Else
'                Nodo_Principal.setAttribute "MetodoPago", "PPD"
'            End If
'            End If
            Nodo_Principal.setAttribute "LugarExpedicion", CFD_Emisor.cp
            'If CFD_Generales.No_Cuenta_Pago <> "" Then
             '   Nodo_Principal.setAttribute "NumCtaPago", CFD_Generales.No_Cuenta_Pago
            'End If
        'Agrega un salto de linea
        Nodo_Principal.appendChild CFD_Documento.createTextNode(vbCrLf)
        '*****************************************************************************************
        'CREA NODO Comprobantes Relacionados
        '*****************************************************************************************
        'Agrega el elemento del relacionados
        
        If CFD_Relacionados.Existe = True And UBound(CFD_Relacionados_Conceptos) > 0 Then
            Set Nodo_Relacionados = CFD_Documento.createElement("cfdi:CfdiRelacionados")
            Nodo_Relacionados.setAttribute "TipoRelacion", Mid(CFD_Relacionados.Relacionados, 1, 2)
            For i = 1 To UBound(CFD_Relacionados_Conceptos)
                Set Nodo_Relacionados_UUID = CFD_Documento.createElement("cfdi:CfdiRelacionado")
                Nodo_Relacionados_UUID.setAttribute "UUID", CFD_Relacionados_Conceptos(i).UUID
                Nodo_Relacionados.appendChild Nodo_Relacionados_UUID
                Nodo_Principal.appendChild Nodo_Relacionados
                Nodo_Relacionados.appendChild CFD_Documento.createTextNode(vbCrLf)
            Next i
        ElseIf CFD_Relacionados.Existe = True Then
            Set Nodo_Relacionados = CFD_Documento.createElement("cfdi:CfdiRelacionados")
            Nodo_Relacionados.setAttribute "TipoRelacion", Mid(CFD_Relacionados.Relacionados, 1, 2)
            Set Nodo_Relacionados_UUID = CFD_Documento.createElement("cfdi:CfdiRelacionado")
            Nodo_Relacionados_UUID.setAttribute "UUID", CFD_Relacionados.UUID_Relacionados
            Nodo_Relacionados.appendChild Nodo_Relacionados_UUID
            Nodo_Principal.appendChild Nodo_Relacionados
            Nodo_Relacionados.appendChild CFD_Documento.createTextNode(vbCrLf)
        End If
        
        
        '*****************************************************************************************
        'CREA NODO EMISOR
        '*****************************************************************************************
        'Agrega el elemento del Emisor
        Set Nodo_Emisor = CFD_Documento.createElement("cfdi:Emisor")
            'Agrega los atributos al nodo Emisor
            Nodo_Emisor.setAttribute "Rfc", CFD_Emisor.RFC
            If CFD_Emisor.Nombre <> "" Then
                Nodo_Emisor.setAttribute "Nombre", CFD_Emisor.Nombre
            End If
            Nodo_Emisor.setAttribute "RegimenFiscal", Mid(CFD_Emisor.Regimen_Fiscal, 1, 3)
            'Agrega un salto de linea
            Nodo_Principal.appendChild Nodo_Emisor
'            Nodo_Emisor.appendChild CFD_Documento.createTextNode(vbCrLf)
            
            'Agrega los elementos al Emisor
            'Set Nodo_Emisor_Domicilio = CFD_Documento.createElement("cfdi:DomicilioFiscal")
                'Agrega los atributos del domicilio al nodo Emisor
              '  Nodo_Emisor_Domicilio.setAttribute "calle", CFD_Emisor.Calle
              '  If CFD_Emisor.No_Exterior <> "" Then
              '      Nodo_Emisor_Domicilio.setAttribute "noExterior", CFD_Emisor.No_Exterior
              '  End If
              '  If CFD_Emisor.No_Interior <> "" Then
              '      Nodo_Emisor_Domicilio.setAttribute "noInterior", CFD_Emisor.No_Interior
              '  End If
                'If CFD_Emisor.Colonia <> "" Then
              '      Nodo_Emisor_Domicilio.setAttribute "colonia", CFD_Emisor.Colonia
              '  End If
              '  If CFD_Emisor.Localidad <> "" Then
              '      Nodo_Emisor_Domicilio.setAttribute "localidad", CFD_Emisor.Localidad
              '  End If
              '  Nodo_Emisor_Domicilio.setAttribute "municipio", CFD_Emisor.Municipio
              '  Nodo_Emisor_Domicilio.setAttribute "estado", CFD_Emisor.Estado
              '  Nodo_Emisor_Domicilio.setAttribute "pais", CFD_Emisor.Pais
              '  Nodo_Emisor_Domicilio.setAttribute "codigoPostal", CFD_Emisor.CP
            'Asigna el nodo del domicilio al nodo emisor
          '  Nodo_Emisor.appendChild Nodo_Emisor_Domicilio
        'Agrega un salto de linea
       ' Nodo_Emisor.appendChild CFD_Documento.createTextNode(vbCrLf)
    
        '*****************************************************************************************
        'CREA NODO EXPEDIDOEN
        '*****************************************************************************************
        'If Sucursal = True Then
            'Agrega los elementos al Emisor
         '   Set Nodo_ExpedidoEn = CFD_Documento.createElement("cfdi:ExpedidoEn")
                'Agrega los atributos del domicilio al nodo Emisor
          '      Nodo_ExpedidoEn.setAttribute "calle", CFD_ExpedidoEn.Calle
           '     If CFD_ExpedidoEn.No_Exterior <> "" Then
            '        Nodo_ExpedidoEn.setAttribute "noExterior", CFD_ExpedidoEn.No_Exterior
             '   End If
             '   If CFD_ExpedidoEn.No_Interior <> "" Then
             '       Nodo_ExpedidoEn.setAttribute "noInterior", CFD_ExpedidoEn.No_Interior
             '   End If
             '   If CFD_ExpedidoEn.Colonia <> "" Then
             '       Nodo_ExpedidoEn.setAttribute "colonia", CFD_ExpedidoEn.Colonia
             '   End If
             '   If CFD_ExpedidoEn.Ciudad <> "" Then
             '       Nodo_ExpedidoEn.setAttribute "municipio", CFD_ExpedidoEn.Ciudad
             '   End If
             '   If CFD_ExpedidoEn.Estado <> "" Then
             '      Nodo_ExpedidoEn.setAttribute "estado", CFD_ExpedidoEn.Estado
             '   End If
             '   If CFD_ExpedidoEn.Pais <> "" Then
             '       Nodo_ExpedidoEn.setAttribute "pais", CFD_ExpedidoEn.Pais
             '   End If
             '   If CFD_ExpedidoEn.CP <> "" Then
             '       Nodo_ExpedidoEn.setAttribute "codigoPostal", CFD_ExpedidoEn.CP
             '   End If
               'Asigna el nodo del domicilio al nodo emisor
             '   Nodo_Emisor.appendChild Nodo_ExpedidoEn
            'Agrega un salto de linea
           ' Nodo_Emisor.appendChild CFD_Documento.createTextNode(vbCrLf)
        'End If
        
        '*****************************************************************************************
        'CREA NODO REGIMEN FISCAL
        '*****************************************************************************************
        'Agrega los elementos al Emisor
        'Set Nodo_Regimen = CFD_Documento.createElement("cfdi:RegimenFiscal")
            'Agrega los atributos del domicilio al nodo Emisor
        '    Nodo_Regimen.setAttribute "Regimen", Regimen_Fiscal
            'Asigna el nodo del domicilio al nodo emisor
        '    Nodo_Emisor.appendChild Nodo_Regimen
        'Agrega un salto de linea
        'Nodo_Emisor.appendChild CFD_Documento.createTextNode(vbCrLf)
                      
        'Asigna el nodo Emisor al nodo principal
        'Nodo_Principal.appendChild Nodo_Emisor
        'Agrega un salto de linea
        'Nodo_Principal.appendChild CFD_Documento.createTextNode(vbCrLf)
            
        '*****************************************************************************************
        'CREA NODO RECEPTOR
        '*****************************************************************************************
        'Agrega el elemento del Receptor
        Set Nodo_Receptor = CFD_Documento.createElement("cfdi:Receptor")
        
            'Agrega los atributos al nodo Emisor
            Nodo_Receptor.setAttribute "Rfc", CFD_Receptor.RFC
            If CFD_Receptor.Nombre <> "" Then
                Nodo_Receptor.setAttribute "Nombre", Valida_Caracteres_Especiales(CFD_Receptor.Nombre)
            End If
           
            Nodo_Receptor.setAttribute "UsoCFDI", Mid(CFD_Receptor.Uso_CFDI, 1, 3)
            'Agrega un salto de linea
'            Nodo_Receptor.appendChild CFD_Documento.createTextNode(vbCrLf)
                'Agrega los elementos al Receptor
               ' Set Nodo_Receptor_Domicilio = CFD_Documento.createElement("cfdi:Domicilio")
                
                'Agrega los atributos del domicilio al nodo Receptor
               ' Nodo_Receptor_Domicilio.setAttribute "calle", CFD_Receptor.Calle
               ' If CFD_Receptor.No_Exterior <> "" Then
               '     Nodo_Receptor_Domicilio.setAttribute "noExterior", CFD_Receptor.No_Exterior
               ' End If
               ' If CFD_Receptor.No_Interior <> "" Then
               '     Nodo_Receptor_Domicilio.setAttribute "noInterior", CFD_Receptor.No_Interior
               ' End If
               ' If CFD_Receptor.Colonia <> "" Then
               '     Nodo_Receptor_Domicilio.setAttribute "colonia", CFD_Receptor.Colonia
               ' End If
               ' If CFD_Receptor.Localidad <> "" Then
               '     Nodo_Receptor_Domicilio.setAttribute "localidad", CFD_Receptor.Localidad
               ' End If
               ' Nodo_Receptor_Domicilio.setAttribute "municipio", CFD_Receptor.Municipio
               ' Nodo_Receptor_Domicilio.setAttribute "estado", CFD_Receptor.Estado
               ' Nodo_Receptor_Domicilio.setAttribute "pais", CFD_Receptor.Pais
               ' Nodo_Receptor_Domicilio.setAttribute "codigoPostal", CFD_Receptor.CP
                'Asigna el nodo del domicilio al nodo Receptor
               ' Nodo_Receptor.appendChild Nodo_Receptor_Domicilio
            'Agrega un salto de linea
           ' Nodo_Receptor.appendChild CFD_Documento.createTextNode(vbCrLf)
        
        'Asigna el nodo Receptor al nodo principal
        Nodo_Principal.appendChild Nodo_Receptor
        'Agrega un salto de linea
       ' Nodo_Principal.appendChild CFD_Documento.createTextNode(vbCrLf)
    
        '*****************************************************************************************
        'CREA NODO CONCEPTOS
        '*****************************************************************************************
        'Agrega el elemento de Conceptos
        Set Nodo_Conceptos = CFD_Documento.createElement("cfdi:Conceptos")
            'Inicializa las variables
            Impuestos_Retenidos = 0
            Impuestos_Trasladados = 0
            'Agrega los atributos del Concepto al nodo Conceptos
            For Conteo_Estructura = 1 To UBound(CFD_Conceptos)
''                'Valida que el concepto tenga cantidad e importe para agregarlo al xml
''                If CFD_Conceptos(Conteo_Estructura).Cantidad > 0 And CFD_Conceptos(Conteo_Estructura).Importe > 0 Then
                    'Agrega un salto de linea
                    Nodo_Conceptos.appendChild CFD_Documento.createTextNode(vbCrLf)
                    'Agrega los elementos al Concepto
                    Set Nodo_Concepto = CFD_Documento.createElement("cfdi:Concepto")
                        'Asigna los atributos
                        Nodo_Concepto.setAttribute "ClaveProdServ", Mid(CFD_Conceptos(Conteo_Estructura).Cod_prod, 1, 8) 'Código del producto
                        If CFD_Conceptos(Conteo_Estructura).No_Identificacion <> "" Then
                            Nodo_Concepto.setAttribute "NoIdentificacion", CFD_Conceptos(Conteo_Estructura).No_Identificacion
                        End If
                        Nodo_Concepto.setAttribute "Cantidad", Format(CFD_Conceptos(Conteo_Estructura).Cantidad, "#0.00")
                        Nodo_Concepto.setAttribute "ClaveUnidad", CFD_Conceptos(Conteo_Estructura).Unidad_Medida
                        If CFD_Conceptos(Conteo_Estructura).Unidad <> "" Then Nodo_Concepto.setAttribute "Unidad", Trim(Mid(CFD_Conceptos(Conteo_Estructura).Unidad, 1, 20))
                        Nodo_Concepto.setAttribute "Descripcion", CFD_Conceptos(Conteo_Estructura).Descripcion
                        Nodo_Concepto.setAttribute "ValorUnitario", CDbl(Format(CFD_Conceptos(Conteo_Estructura).Valor_Unitario, "#0.00"))
                        Nodo_Concepto.setAttribute "Importe", CDbl(Format(CFD_Conceptos(Conteo_Estructura).Importe, "#0.00"))
                        If CFD_Generales.Descuento > 0 Then
                            Nodo_Concepto.setAttribute "Descuento", CFD_Generales.Descuento / UBound(CFD_Conceptos) 'Format(CFD_Conceptos(Conteo_Estructura).Importe * CFD_Generales.Descuento, "#0.00")
                        End If
                        'Nodo_Concepto.setAttribute "Impuesto", CDbl(Format(CFD_Conceptos(Conteo_Estructura).Importe, "#0.00"))
                        If CFD_Conceptos(Conteo_Estructura).IVA_Producto = True Or UBound(CFD_Impuestos_Retenidos) > 0 Then
                            Set Nodo_Impuesto_Concepto = CFD_Documento.createElement("cfdi:Impuestos")
                            If CFD_Conceptos(Conteo_Estructura).IVA_Producto = True Then
                            
                            Set Nodo_Traslado_Concepto = CFD_Documento.createElement("cfdi:Traslados")
                            Set Nodo_Traslado_Conceptos = CFD_Documento.createElement("cfdi:Traslado")
                                Nodo_Traslado_Conceptos.setAttribute "Base", CDbl(Format(CFD_Conceptos(Conteo_Estructura).Importe, "#0.00"))
                                Nodo_Traslado_Conceptos.setAttribute "Impuesto", "002"
                                Nodo_Traslado_Conceptos.setAttribute "TipoFactor", "Tasa"
                                Nodo_Traslado_Conceptos.setAttribute "TasaOCuota", Format(Val(PG_Retencion_IVA), "#0.000000") 'Val(CFD_Generales.Tasa_IVA / 100) 'double
                                Nodo_Traslado_Conceptos.setAttribute "Importe", Format(CFD_Conceptos(Conteo_Estructura).Importe * Val(PG_Retencion_IVA), "#0.00")   'CFD_Impuestos(1).Tasa 'Val(CFD_Generales.Tasa_IVA / 100),"#0.00") 'CDbl(Format(CFD_Conceptos(Conteo_Estructura).IVA_Producto, "#0.00"))  'double
                                Impuestos_Trasladados = Impuestos_Trasladados + Format(CFD_Conceptos(Conteo_Estructura).Importe * Val(PG_Retencion_IVA), "#0.00")
                                
                                Nodo_Traslado_Concepto.appendChild Nodo_Traslado_Conceptos
                            
                             'Nodo_Traslado_Concepto.appendChild Nodo_Traslado_Conceptos
                             Nodo_Impuesto_Concepto.appendChild Nodo_Traslado_Concepto
                        End If
                        
                        If UBound(CFD_Impuestos_Retenidos) > 0 Then
                            Set Nodo_Retencion_Concepto = CFD_Documento.createElement("cfdi:Retenciones")
                            For Conteo_Estructura2 = 1 To UBound(CFD_Impuestos_Retenidos)
                                'Agrega los elementos de Retenciones
                                If CFD_Impuestos_Retenidos(Conteo_Estructura2).Impuesto = "002" Or CFD_Impuestos_Retenidos(Conteo_Estructura2).Impuesto = "001" Then
                                Set Nodo_Retencion_Conceptos = CFD_Documento.createElement("cfdi:Retencion")
                                Nodo_Retencion_Conceptos.setAttribute "Base", CDbl(Format(CFD_Conceptos(Conteo_Estructura).Importe, "#0.00"))
                                Nodo_Retencion_Conceptos.setAttribute "Impuesto", CFD_Impuestos_Retenidos(Conteo_Estructura2).Impuesto
                                Nodo_Retencion_Conceptos.setAttribute "TipoFactor", "Tasa"
                                If CFD_Impuestos_Retenidos(Conteo_Estructura2).Impuesto = "002" Then
                                    Nodo_Retencion_Conceptos.setAttribute "TasaOCuota", Format(Val(CFD_Impuestos_Retenidos(1).Tasa), "#0.000000")
                                    Nodo_Retencion_Conceptos.setAttribute "Importe", Val(CFD_Impuestos_Retenidos(1).Tasa) * CDbl(Format(CFD_Conceptos(Conteo_Estructura).Importe, "#0.00"))
                                ElseIf CFD_Impuestos_Retenidos(Conteo_Estructura2).Impuesto = "001" Then
                                    Nodo_Retencion_Conceptos.setAttribute "TasaOCuota", Format(Val(CFD_Impuestos_Retenidos(2).Tasa), "#0.000000")
                                    Nodo_Retencion_Conceptos.setAttribute "Importe", CDbl(Format(CFD_Conceptos(Conteo_Estructura).Importe, "#0.00")) * Val(CFD_Impuestos_Retenidos(2).Tasa)
                                End If
                                ' Nodo_Retencion_Conceptos.setAttribute "Importe", Format(CFD_Impuestos_Retenidos(Conteo_Estructura2).Importe, "#0.00")
                                'Asigna el nodo de Traslado al nodo Retenciones
                                Nodo_Retencion_Concepto.appendChild Nodo_Retencion_Conceptos
                                End If
                                'Agrega un salto de linea
                                Nodo_Retencion_Concepto.appendChild CFD_Documento.createTextNode(vbCrLf)
                            Next Conteo_Estructura2
                        'Cierra el nodo de Retenciones al nodo Impuestos
                         Nodo_Impuesto_Concepto.appendChild Nodo_Retencion_Concepto
                        End If
                        Nodo_Concepto.appendChild Nodo_Impuesto_Concepto
                       End If
                       
                       'If CFD_Conceptos(Conteo_Estructura).No_Pedimento <> "" Then
                           ' Set Nodo_Aduanera = CFD_Documento.createElement("cfdi:InformacionAduanera")
                            'Nodo_Aduanera.setAttribute "NumeroPedimento", CFD_Conceptos(Conteo_Estructura).No_Pedimento
                            'Nodo_Concepto.appendChild Nodo_Aduanera
                       'End If
                       
                        Nodo_Conceptos.appendChild Nodo_Concepto
                        'Valida si hay informacion de la cuenta predial que ingresar
                        
                    'Asigna el nodo de Conceptos al nodo Conceptos
                    Nodo_Conceptos.appendChild Nodo_Concepto
''                End If
            Next Conteo_Estructura
            'Agrega un salto de linea
            Nodo_Conceptos.appendChild CFD_Documento.createTextNode(vbCrLf)
        
        'Asigna el nodo Conceptos al nodo principal
        Nodo_Principal.appendChild Nodo_Conceptos
        'Agrega un salto de linea
        Nodo_Principal.appendChild CFD_Documento.createTextNode(vbCrLf)
    
        '*****************************************************************************************
        'CREA NODO IMPUESTOS
        '*****************************************************************************************
       
        'Agrega el elemento de Impuestos
        If UBound(CFD_Impuestos) > 0 Or UBound(CFD_Impuestos_Retenidos) > 0 Then
        Set Nodo_Impuestos = CFD_Documento.createElement("cfdi:Impuestos")
    
            If IVA_EXENTO = False Then
            
''                'Inicializa las variables
'                Impuestos_Retenidos = 0
'                Impuestos_Trasladados = 0
'
'                'Verifica si existe monto de retenciones
'                For Conteo_Estructura = 1 To UBound(CFD_Impuestos_Retenidos)
'                    Impuestos_Retenidos = Impuestos_Retenidos + CFD_Impuestos_Retenidos(Conteo_Estructura).Importe
'                Next Conteo_Estructura
'
'                'Verifica si existe monto de trasladados
'                For Conteo_Estructura = 1 To UBound(CFD_Impuestos)
'                    Impuestos_Trasladados = Impuestos_Trasladados + CFD_Impuestos(Conteo_Estructura).Importe
'                Next Conteo_Estructura
                
''                'Agrega los atributos del nodo Impuestos
                Nodo_Impuestos.setAttribute "TotalImpuestosTrasladados", Format(Impuestos_Trasladados, "#0.00")
                
                'si hay impuestos retenidos los toma en cuenta para mostrarlos en el xml
                  
                    'Nodo_Impuestos.setAttribute "TotalImpuestosRetenidos", Format(CFD_Generales.Total_Retenciones, "#0.00")
                    Nodo_Impuestos.setAttribute "TotalImpuestosRetenidos", Format(Impuestos_Retenidos, "#0.00")
                    If Impuestos_Retenidos > 0 Then
                    'Nodo_Impuestos.setAttribute "TotalImpuestosTrasladados", Format(Impuestos_Trasladados, "#0.00")
                    'Agrega los elementos de retenidos
                    Set Nodo_Retenciones = CFD_Documento.createElement("cfdi:Retenciones")
                        'Agrega un salto de linea
                        Nodo_Impuestos.appendChild CFD_Documento.createTextNode(vbCrLf)
                        'Agrega un salto de linea
                        Nodo_Retenciones.appendChild CFD_Documento.createTextNode(vbCrLf)
                        'Agrega los atributos del Concepto al nodo Conceptos
                        For Conteo_Estructura = 1 To UBound(CFD_Impuestos_Retenidos)
                            'Agrega los elementos de Retenciones
                           If CFD_Impuestos_Retenidos(Conteo_Estructura).Impuesto = "002" Or CFD_Impuestos_Retenidos(Conteo_Estructura).Impuesto = "001" Then
                            Set Nodo_Retencion = CFD_Documento.createElement("cfdi:Retencion")
                            'Agrega los atributos de la retencion al nodo Retenciones
                            'Nodo_Retencion.setAttribute "Base", Format(CFD_Generales.SubTotal, "#0.00")
                            Nodo_Retencion.setAttribute "Impuesto", CFD_Impuestos_Retenidos(Conteo_Estructura).Impuesto
                            'Nodo_Retencion.setAttribute "TipoFactor", "Tasa"
                            If CFD_Impuestos_Retenidos(Conteo_Estructura).Impuesto = "002" Then
                            '    Nodo_Retencion.setAttribute "TasaOCuota", Val(CFD_Impuestos_Retenidos(1).Tasa)
                                Nodo_Retencion.setAttribute "Importe", Format(CFD_Impuestos_Retenidos(Conteo_Estructura).Importe, "#0.00")
                             ElseIf CFD_Impuestos_Retenidos(Conteo_Estructura).Impuesto = "001" Then
                              '      Nodo_Retencion.setAttribute "TasaOCuota", Val(CFD_Impuestos_Retenidos(2).Tasa)
                                    'Nodo_Retencion.setAttribute "Importe", CDbl(Format(CFD_Impuestos_Retenidos(2).Tasa, "#0.00") * CFD_Generales.Subtotal)
                                    Nodo_Retencion.setAttribute "Importe", CDbl(Format(CFD_Impuestos_Retenidos(Conteo_Estructura).Importe, "#0.00"))
                             End If
                            'Nodo_Retencion.setAttribute "Importe", 'CDbl(Format(CFD_Impuestos_Retenidos(Conteo_Estructura).Importe, "#0.00"))
                            'Asigna el nodo de Traslado al nodo Retenciones
                            Nodo_Retenciones.appendChild Nodo_Retencion
                            'Agrega un salto de linea
                            Nodo_Retenciones.appendChild CFD_Documento.createTextNode(vbCrLf)
                            End If
                        Next Conteo_Estructura
                    'Cierra el nodo de Retenciones al nodo Impuestos
                    Nodo_Impuestos.appendChild Nodo_Retenciones
                  End If
                Else
                    'Si no hubo retenciones agrega el atributo de impuestos trasladados
                    Nodo_Impuestos.setAttribute "TotalImpuestosTrasladados", CDbl(Format(Impuestos_Trasladados, "#0.00")) 'double
                End If
                
                'Agrega los elementos de impuestos Traslados
                Set Nodo_Traslados = CFD_Documento.createElement("cfdi:Traslados")
                    'Agrega un salto de linea
                    Nodo_Traslados.appendChild CFD_Documento.createTextNode(vbCrLf)
                    'Agrega los atributos del Concepto al nodo Conceptos
                    For Conteo_Estructura = 1 To UBound(CFD_Impuestos)
                        'Agrega los elementos de Traslado
                        Set Nodo_Traslado = CFD_Documento.createElement("cfdi:Traslado")
                            'Agrega los atributos del Traslado al nodo Traslados
                            'Nodo_Traslado.setAttribute "Base", CDbl(Format(CFD_Impuestos(Conteo_Estructura).Importe, "#0.00")) 'double
                        If CFD_Impuestos(Conteo_Estructura).Impuesto = "002" Then
                            'Nodo_Traslado.setAttribute "Base", CDbl(Format(CFD_Generales.SubTotal, "#0.00"))
                            Nodo_Traslado.setAttribute "Impuesto", CFD_Impuestos(Conteo_Estructura).Impuesto
                            Nodo_Traslado.setAttribute "TipoFactor", "Tasa"
                            Nodo_Traslado.setAttribute "TasaOCuota", Format(Val(PG_Retencion_IVA), "#0.000000") 'Format(CFD_Impuestos(1).Tasa, "#0.00") 'Format(Val(CFD_Generales.Tasa_IVA / 100), "#0.00") 'double
                            Nodo_Traslado.setAttribute "Importe", Format(Impuestos_Trasladados, "#0.00") 'Format(CFD_Impuestos(Conteo_Estructura).Importe, "#0.00")
                            
                          End If
'                             If CFD_Impuestos(Conteo_Estructura).Impuesto = "003" Then
'                             'If CFD_Impuestos(2).Importe > 0 Then
'                                Nodo_Traslado.setAttribute "Base", CDbl(Format(CFD_Generales.SubTotal, "#0.00"))
'                                Nodo_Traslado.setAttribute "Impuesto", CFD_Impuestos(Conteo_Estructura).Impuesto
'
'                                Nodo_Traslado.setAttribute "TipoFactor", CFD_IEPS(1).factor
'                                Nodo_Traslado.setAttribute "TasaOCuota", Val(CFD_IEPS(1).Tasa_IEPS / 100) 'double
'                                Nodo_Traslado.setAttribute "Importe", CDbl(Format(CFD_Impuestos(2).Importe, "#0.00"))
'                           End If

                        'Asigna el nodo de Traslado al nodo Traslados
                        Nodo_Traslados.appendChild Nodo_Traslado
                        'Agrega un salto de linea
                        Nodo_Traslados.appendChild CFD_Documento.createTextNode(vbCrLf)
                    Next Conteo_Estructura
                'Cierra el nodo de Traslados al nodo Impuestos
                Nodo_Impuestos.appendChild Nodo_Traslados
                'Agrega un salto de linea
                Nodo_Impuestos.appendChild CFD_Documento.createTextNode(vbCrLf)
            'End If
        'Cierra el nodo Impuestos al nodo principal
        Nodo_Principal.appendChild Nodo_Impuestos
    End If
        '*****************************************************************************************
        'CREA NODO COMPLEMENTO
        '*****************************************************************************************
        Nodo_Principal.appendChild CFD_Documento.createTextNode(vbCrLf)
        'Agrega el elemento de la complemento
        Set Nodo_Complemento = CFD_Documento.createElement("cfdi:Complemento")
            'Agrega un salto de linea
'            Nodo_Complemento.appendChild CFD_Documento.createTextNode(vbCrLf)
           If CFD_Generales.Tipo_Factura = "Pagos" Then
            Set Nodo_Pagos = CFD_Documento.createElement("pago10:Pagos")
                Nodo_Pagos.setAttribute "Version", "1.0"
'                Nodo_Pagos.appendChild CFD_Documento.createTextNode(vbCrLf)
            Set Nodo_Pago = CFD_Documento.createElement("pago10:Pago")
                Nodo_Pago.setAttribute "FechaPago", CFD_Pagos.Fecha_Pago
                Nodo_Pago.setAttribute "FormaDePagoP", Mid(CFD_Pagos.Forma_Pago_Pagos, 1, 2)
                Nodo_Pago.setAttribute "MonedaP", CFD_Pagos.Moneda_Pago
                If CFD_Pagos.Tipo_Cambio_Pago <> "" And Val(CFD_Pagos.Tipo_Cambio_Pago) > 0 Then Nodo_Pago.setAttribute "TipoCambioP", CFD_Pagos.Tipo_Cambio_Pago
                Nodo_Pago.setAttribute "Monto", CFD_Pagos.Monto
                If CFD_Pagos.Num_Operacion <> "" Then Nodo_Pago.setAttribute "NumOperacion", CFD_Pagos.Num_Operacion
                
                If Mid(CFD_Pagos.Forma_Pago_Pagos, 1, 2) <> "01" Then
                    If CFD_Pagos.RfcEmisorCtaOrd <> "" Then Nodo_Pago.setAttribute "RfcEmisorCtaOrd", CFD_Pagos.RfcEmisorCtaOrd
                    If CFD_Pagos.NomBancoOrdExt <> "" Then Nodo_Pago.setAttribute "NomBancoOrdExt", CFD_Pagos.NomBancoOrdExt
                    If CFD_Pagos.CtaOrdenante <> "" Then Nodo_Pago.setAttribute "CtaOrdenante", CFD_Pagos.CtaOrdenante
                    If CFD_Pagos.RfcEmisorCtaBen <> "" Then Nodo_Pago.setAttribute "RfcEmisorCtaBen", CFD_Pagos.RfcEmisorCtaBen
                    If CFD_Pagos.CtaBeneficiario <> "" Then Nodo_Pago.setAttribute "CtaBeneficiario", CFD_Pagos.CtaBeneficiario
                    If CFD_Pagos.TipoCadPago <> "" Then Nodo_Pago.setAttribute "TipoCadPago", Mid(CFD_Pagos.TipoCadPago, 1, 2)
                    If CFD_Pagos.CertPago <> "" Then Nodo_Pago.setAttribute "CertPago", CFD_Pagos.CertPago
                    If CFD_Pagos.CadPago <> "" Then Nodo_Pago.setAttribute "CadPago", CFD_Pagos.CadPago
                    If CFD_Pagos.Sello_Pago <> "" Then Nodo_Pago.setAttribute "SelloPago", CFD_Pagos.Sello_Pago
                End If
'                Nodo_Pago.appendChild CFD_Documento.createTextNode(vbCrLf)
                For i = 1 To UBound(CFD_Pagos_DR)
                    Set Nodo_Pagos_Relacionados = CFD_Documento.createElement("pago10:DoctoRelacionado")
                    Nodo_Pagos_Relacionados.setAttribute "IdDocumento", CFD_Pagos_DR(i).ID_Doc
                    If CFD_Pagos_DR(i).Serie <> "" Then Nodo_Pagos_Relacionados.setAttribute "Serie", CFD_Pagos_DR(i).Serie
                    Nodo_Pagos_Relacionados.setAttribute "Folio", CFD_Pagos_DR(i).Folio
                    Nodo_Pagos_Relacionados.setAttribute "MonedaDR", CFD_Pagos_DR(i).Moneda_DR
                    If CFD_Pagos_DR(i).Moneda_DR <> CFD_Pagos.Moneda_Pago And CFD_Pagos_DR(i).Tipo_Cambio_DR = "" Then
                        Nodo_Pagos_Relacionados.setAttribute "TipoCambioDR", 1
                    ElseIf CFD_Pagos_DR(i).Tipo_Cambio_DR <> "" And Val(CFD_Pagos_DR(i).Tipo_Cambio_DR) > 0 Then
                        Nodo_Pagos_Relacionados.setAttribute "TipoCambioDR", CDbl(Val(CFD_Pagos_DR(i).Tipo_Cambio_DR))
                    End If
                    Nodo_Pagos_Relacionados.setAttribute "MetodoDePagoDR", Mid(CFD_Pagos_DR(i).Metodo_Pago_DR, 1, 3)
                    Nodo_Pagos_Relacionados.setAttribute "NumParcialidad", CFD_Pagos_DR(i).No_Parcialidad
                    Nodo_Pagos_Relacionados.setAttribute "ImpSaldoAnt", CFD_Pagos_DR(i).Saldo_Anterior
                    Nodo_Pagos_Relacionados.setAttribute "ImpPagado", CFD_Pagos_DR(i).Importe_Pagado 'CFD_Pagos.Importe_Pagado
                    Nodo_Pagos_Relacionados.setAttribute "ImpSaldoInsoluto", CFD_Pagos_DR(i).Saldo_Insoluto
'                    Nodo_Pagos_Relacionados.appendChild CFD_Documento.createTextNode(vbCrLf)
                 'Asigna el nodo de relacionados a pago
                        Nodo_Pago.appendChild Nodo_Pagos_Relacionados
                Next
                        'Agrega un salto de linea
'                        Nodo_Pago.appendChild CFD_Documento.createTextNode(vbCrLf)
                'Cierra el nodo de Traslados al nodo Impuestos
               
                Nodo_Pagos.appendChild Nodo_Pago
'                Nodo_Pagos.appendChild CFD_Documento.createTextNode(vbCrLf)
                Nodo_Complemento.appendChild Nodo_Pagos
           End If
           
        'Cierra el nodo Addenda al nodo principal
        Nodo_Principal.appendChild Nodo_Complemento
        
    'Agrega un salto de linea
'    Nodo_Complemento.appendChild CFD_Documento.createTextNode(vbCrLf)
        
    'Cierra el nodo principal
'    Nodo_Principal.appendChild CFD_Documento.createTextNode(vbCrLf)
      
    'Limpia las variables del timbre
    
    
    Timbrado_VersionSat = ""
    Timbrado_UUID = ""
    Timbrado_FechaTimbrado = ""
    Timbrado_selloCFD = ""
    Timbrado_noCertificadoSAT = ""
    Timbrado_selloSAT = ""
   
    'Realiza la petición al webservice
   
    If Timbrado_Ambiente = "PRUEBA" Then
       Call Timbrado33.GenerarTimbrado(False, "demo", "123456", 0, "", "", "", "", CStr(CFD_Documento.XML))
    Else
        Call Timbrado33.GenerarTimbrado(True, Timbrado_Codigo_Usuario, Timbrado_Codigo_Usuario_Proveedor, 0, "", "", Timbrado_ID_Sucursal, "", CStr(CFD_Documento.XML))
    End If
    If Timbrado33.EstatusRespuesta = False Then
        Err.Raise &HFFFFFF01, "Error", Timbrado33.MensajeValidacion
    End If
    
    XML = Timbrado33.XMLTimbrado
        

    'Pone los datos en el analizador de XML
    Set Conexion_Web_Services = New DOMDocument
    'Parametros de validacion
    Conexion_Web_Services.resolveExternals = True

    Conexion_Web_Services.validateOnParse = True
    
    Conexion_Web_Services.async = False
   
    'Carga la respuesta del webservice
    If Conexion_Web_Services.loadXML(XML) Then  'Si es valida la respuesta
        'Obtiene la respuesta
        Set Respuesta = Conexion_Web_Services.SelectSingleNode("//tfd:TimbreFiscalDigital")
        If Not Respuesta Is Nothing Then
            'Obtiene los parametros del resultado
            Timbrado_VersionSat = Respuesta.Attributes(2).text
            Timbrado_UUID = Respuesta.Attributes(3).text
            Timbrado_FechaTimbrado = Respuesta.Attributes(4).text
            Timbrado_selloCFD = Respuesta.Attributes(6).text
            Timbrado_noCertificadoSAT = Respuesta.Attributes(7).text
            Timbrado_selloSAT = Respuesta.Attributes(8).text
            Timbrado_RFCTimbra = Respuesta.Attributes(5).text
            'Agrega el elemento del timbrado
            Set Nodo_Timbrado = CFD_Documento.createElement("tfd:TimbreFiscalDigital")
                'Agrega los atributos
                Nodo_Timbrado.setAttribute "xsi:schemaLocation", "http://www.sat.gob.mx/TimbreFiscalDigital http://www.sat.gob.mx/sitio_internet/cfd/TimbreFiscalDigital/TimbreFiscalDigitalv11.xsd"
                Nodo_Timbrado.setAttribute "xmlns:tfd", "http://www.sat.gob.mx/TimbreFiscalDigital"
                Nodo_Timbrado.setAttribute "Version", Timbrado_VersionSat
                Nodo_Timbrado.setAttribute "UUID", Timbrado_UUID
                Nodo_Timbrado.setAttribute "FechaTimbrado", Timbrado_FechaTimbrado
                Nodo_Timbrado.setAttribute "RfcProvCertif", Timbrado_RFCTimbra
                Nodo_Timbrado.setAttribute "SelloCFD", Timbrado_selloCFD
                Nodo_Timbrado.setAttribute "NoCertificadoSAT", Timbrado_noCertificadoSAT
                Nodo_Timbrado.setAttribute "SelloSAT", Timbrado_selloSAT
            'Asigna el nodo timbrado
            Nodo_Complemento.appendChild Nodo_Timbrado
            CFD_Generales.Fecha_Timbrado = Timbrado_FechaTimbrado
        Else 'Si no es valida regresa el error
            'Obtiene la respuesta
            Set Respuesta = Conexion_Web_Services.SelectSingleNode("//Error")
            If Not Respuesta Is Nothing Then
                Err.Raise 7777, "CFD_Crea_Xml", "Error: " & Respuesta.Attributes(0).text & " " & Respuesta.Attributes(1).text
            Else
                Err.Raise 7777, "CFD_Crea_Xml", "Error: No es posible encontrar los parametros de respuesta"
            End If
        End If
    Else 'Si no es valida regresa el error
        Err.Raise 7777, "CFD_Crea_Xml", XML
    End If
    'Valida si hay alguna adenda que agregar
     Select Case Tipo_Adenda
         Case "Adenda"
             CFD_Addenda
         Case "Adenda_Nadro"
             CFD_Addenda_Nadro
     End Select
    'If Valida_Existe_Archivo_Carpeta(Ruta_Xmls, "CARPETA") = False Then
     '   MkDir Ruta_Xmls & "\"
    'End If
    'Guarda el archivo
    CFD_Documento.Save App.Path & "\XML\" & Nombre_CFD & ".xml"
    'CFD_Documento.save Ruta_Xmls & "\" & Nombre_CFD & ".xml"
    'Destruye el documento
    Set CFD_Documento = Nothing
    'Genera el codigo bidimensinal
    Codigo_Bidimensional = "?re=" & CFD_Emisor.RFC & "&rr=" & CFD_Receptor.RFC & "&tt=" & Format(CFD_Generales.Total, "0000000000.000000") & "&id=" & Timbrado_UUID
    If CFD_Generales.Tipo_Comprobante = "I" Or CFD_Generales.Tipo_Comprobante = "P" Then
        Call GenerateFile(Codigo_Bidimensional, Ruta_Pdfs & "\CFDI_" & Trim(CFD_Generales.Serie & "_" & CFD_Generales.Folio) & ".bmp")
    Else
        Call GenerateFile(Codigo_Bidimensional, Ruta_NC & "\CFDI_" & Trim(CFD_Generales.Serie & "_" & CFD_Generales.Folio) & ".bmp")
    End If
End Sub
Private Sub CFD_Addenda_Nadro()
Dim Nodo_Addenda As IXMLDOMElement
Dim Nodo_Datos_Nadro As IXMLDOMElement
Dim Nodo_Orden_Compra As IXMLDOMElement
Dim Nodo_Plazo As IXMLDOMElement
Dim Nodo_Entrega_Entrante As IXMLDOMElement
Dim Nodo_PosicionOC As IXMLDOMElement
Dim Nodo_TotalOC As IXMLDOMElement
Dim Nodo_CodigoEAN As IXMLDOMElement
Dim Nodo_Elemento As IXMLDOMElement
Dim i As Long
Dim cadena() As String

    '*****************************************************************************************
    'ENCABEZADO DE LA ADDENDA
    '*****************************************************************************************
    'CREA NODO ADDENDA
    '*****************************************************************************************
    'Agrega el elemento de la Addenda
    Set Nodo_Addenda = CFD_Documento.createElement("cfdi:Addenda")
    'Agrega un salto de linea
    For i = 1 To UBound(CFD_Conceptos)
        'Crea nodo de datos Nadro
        Set Nodo_Datos_Nadro = CFD_Documento.createElement("DatosNadro")
        'Crea los nodo hijo orden de compra
        Set Nodo_Orden_Compra = CFD_Documento.createElement("Orden")
        Nodo_Orden_Compra.text = (Trim(CFD_Generales.Orden_Compra))
        Nodo_Datos_Nadro.appendChild Nodo_Orden_Compra
        'crea nodo hijo plazo
        Set Nodo_Plazo = CFD_Documento.createElement("Plazo")
        Nodo_Plazo.text = (Trim(CFD_Generales.Plazo))
        Nodo_Datos_Nadro.appendChild Nodo_Plazo
        'crea hijo entrega entrante
        Set Nodo_Entrega_Entrante = CFD_Documento.createElement("EntregaEntrante")
        Nodo_Entrega_Entrante.text = 0
        Nodo_Datos_Nadro.appendChild Nodo_Entrega_Entrante
        'crea hijo posicion en la orden de compra
        Set Nodo_PosicionOC = CFD_Documento.createElement("PosicionOC")
        Nodo_PosicionOC.text = CFD_Conceptos(i).Posicion_OC
        Nodo_Datos_Nadro.appendChild Nodo_PosicionOC
        'crea nodo hijo TotalOC
        Set Nodo_TotalOC = CFD_Documento.createElement("TotalOC")
        Nodo_TotalOC.text = CFD_Conceptos(i).Importe
        Nodo_Datos_Nadro.appendChild Nodo_TotalOC
        'crea nodo hijo código del producto
        cadena = Split(CFD_Conceptos(i).Descripcion, " ")
        Set Nodo_CodigoEAN = CFD_Documento.createElement("CodEAN")
        Nodo_CodigoEAN.text = Trim(cadena(0))
        Nodo_Datos_Nadro.appendChild Nodo_CodigoEAN
        
        Nodo_Addenda.appendChild Nodo_Datos_Nadro
    Next i
    'Agrega un salto de linea
    Nodo_Addenda.appendChild CFD_Documento.createTextNode(vbCrLf)
    'Cierra el nodo Addenda al nodo principal
    Nodo_Principal.appendChild Nodo_Addenda
    'Agrega un salto de linea
    Nodo_Addenda.appendChild CFD_Documento.createTextNode(vbCrLf)
    'Agrega un salto de linea
    Nodo_Principal.appendChild CFD_Documento.createTextNode(vbCrLf)
End Sub
Private Sub CFD_Addenda()
Dim Nodo_Addenda As IXMLDOMElement
Dim Nodo_Orden_Compra As IXMLDOMElement
Dim Nodo_Elemento As IXMLDOMElement

    '*****************************************************************************************
    'ENCABEZADO DE LA ADDENDA
    '*****************************************************************************************
    'CREA NODO ADDENDA
    '*****************************************************************************************
    'Agrega el elemento de la Addenda
    Set Nodo_Addenda = CFD_Documento.createElement("cfdi:Addenda")
    'Agrega un salto de linea
    Nodo_Addenda.appendChild CFD_Documento.createTextNode(vbCrLf)
        '*****************************************************************************************
        'CREA NODO ORDENCOMPRA
        '*****************************************************************************************
        Set Nodo_Orden_Compra = CFD_Documento.createElement("InfoAdicional")
        Nodo_Orden_Compra.setAttribute "OrdenCompra", Trim(CFD_Generales.Orden_Compra)
        Nodo_Addenda.appendChild Nodo_Orden_Compra
        Set Nodo_Orden_Compra = Nothing
        'Agrega un salto de linea
        Nodo_Addenda.appendChild CFD_Documento.createTextNode(vbCrLf)
    'Cierra el nodo Addenda al nodo principal
    Nodo_Principal.appendChild Nodo_Addenda
    'Agrega un salto de linea
    Nodo_Addenda.appendChild CFD_Documento.createTextNode(vbCrLf)
    'Agrega un salto de linea
    Nodo_Principal.appendChild CFD_Documento.createTextNode(vbCrLf)
End Sub
'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: CFD_Crea_Xml
'DESCRIPCIÓN: Crea el archivo xml a la estructura solicitada por el SAT
'PARÁMETROS:  Nombre_CFD - Nombre con que se guardará el XML
'CREO:        Sergio Godínez Banda
'FECHA_CREO:  25-Mayo-2012
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Sub CFD_Crea_Xml_2(Nombre_CFD As String)
Dim Xml_Identificacion As IXMLDOMProcessingInstruction
Dim Nodo_Raiz As IXMLDOMElement
Dim Nodo_Emisor As IXMLDOMElement
Dim Nodo_Emisor_Domicilio As IXMLDOMElement
Dim Nodo_Relacionados As IXMLDOMElement
Dim Nodo_Relacionados_UUID As IXMLDOMElement
Dim Nodo_ExpedidoEn As IXMLDOMElement
Dim Nodo_Regimen As IXMLDOMElement
Dim Nodo_Receptor As IXMLDOMElement
Dim Nodo_Receptor_Domicilio As IXMLDOMElement
Dim Nodo_Conceptos As IXMLDOMElement
Dim Nodo_Concepto As IXMLDOMElement
Dim Nodo_Impuestos As IXMLDOMElement
Dim Nodo_Traslados As IXMLDOMElement
Dim Nodo_Traslado As IXMLDOMElement
Dim Nodo_Retenciones As IXMLDOMElement
Dim Nodo_Retencion As IXMLDOMElement
Dim Nodo_Complemento As IXMLDOMElement
Dim Nodo_ImpuestoLocal As IXMLDOMElement
Dim Nodo_Elemento As IXMLDOMElement
Dim Nodo_Timbrado As IXMLDOMElement
Dim Nodo_Cuenta_Predial As IXMLDOMElement
Dim Nodo_Donaciones As IXMLDOMElement
Dim Nodo_Impuesto_Concepto As IXMLDOMElement
Dim Nodo_Traslado_Concepto As IXMLDOMElement
Dim Nodo_Traslado_Conceptos As IXMLDOMElement
Dim Nodo_Retencion_Concepto As IXMLDOMElement
Dim Nodo_Retencion_Conceptos As IXMLDOMElement
Dim Impuestos_Trasladados As Double
Dim Impuestos_Retenidos As Double
Dim Impuestos_Locales As Double
Dim Impuestos_Locales_Retenidos As Double
Dim Impuestos_Locales_Trasladados As Double
Dim Conteo_Estructura As Integer
Dim Conteo_Estructura2 As Integer
Dim Cadena_XML As String
Dim Genera_XML As String
Dim Proceso_Web_Services As String
Dim Respuesta_Web_Services As String
Dim Respuesta As IXMLDOMNode
Dim CFDI As String
Dim XML As String
Dim Nodo_Pago As IXMLDOMElement
Dim Nodo_Pagos As IXMLDOMElement
Dim Nodo_Pagos_Relacionados As IXMLDOMElement
'Dim Timbrado As New ContelTimbrado.GenerarTimbrado
Dim Timbrado33 As ICls_Fe_Timbrado
Set Timbrado33 = New Cls_Fe_Timbrado
'Timbrado33.
Dim Codigo_Bidimensional As String
Dim Respuesta_Timbrado As String
Dim Importe_IEPS As String
Dim i As Long
    '*****************************************************************************************
    'CREA DOCUMENTO XML
    '*****************************************************************************************
    'Crea el documento xml
    Set CFD_Documento = New DOMDocument
    
    'Identifica la version y codificacion del xml
    Set Xml_Identificacion = CFD_Documento.createProcessingInstruction("xml", "version=""1.0"" encoding=""UTF-8""")
    CFD_Documento.appendChild Xml_Identificacion
    
    'Crea el documento xml
    Set Nodo_Raiz = CFD_Documento.createElement("cfdi:Comprobante")
    Set CFD_Documento.documentElement = Nodo_Raiz
    
        '*****************************************************************************************
        'CREA NODO PRINCIPAL
        '*****************************************************************************************
        'Crea el elemento principal
        Set Nodo_Principal = CFD_Documento.documentElement
            'Agrega los atributos a la raiz del xml
            Nodo_Principal.setAttribute "xsi:schemaLocation", "http://www.sat.gob.mx/cfd/3 http://www.sat.gob.mx/sitio_internet/cfd/3/cfdv33.xsd http://www.sat.gob.mx/implocal http://www.sat.gob.mx/sitio_internet/cfd/implocal/implocal.xsd http://www.sat.gob.mx/donat http://www.sat.gob.mx/sitio_internet/cfd/donat/donat11.xsd http://www.sat.gob.mx/Pagos http://www.sat.gob.mx/sitio_internet/cfd/Pagos/Pagos10.xsd"
''            Nodo_Principal.setAttribute "xsi:schemaLocation", "http://www.sat.gob.mx/cfd/3 http://www.sat.gob.mx/sitio_internet/cfd/3/cfdv32.xsd http://www.sat.gob.mx/implocal http://www.sat.gob.mx/sitio_internet/cfd/implocal/implocal.xsd
            Nodo_Principal.setAttribute "xmlns:implocal", "http://www.sat.gob.mx/implocal"
            Nodo_Principal.setAttribute "xmlns:donat", "http://www.sat.gob.mx/donat"
            Nodo_Principal.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
            Nodo_Principal.setAttribute "xmlns:cfdi", "http://www.sat.gob.mx/cfd/3"
            Nodo_Principal.setAttribute "xmlns:pago10", "http://www.sat.gob.mx/Pagos"
            Nodo_Principal.setAttribute "Version", CFD_Generales.Version
            If CFD_Generales.Serie <> "" Then
                Nodo_Principal.setAttribute "Serie", CFD_Generales.Serie
            End If
            Nodo_Principal.setAttribute "Folio", CFD_Generales.Folio
            Nodo_Principal.setAttribute "Fecha", CFD_Generales.Fecha
            Nodo_Principal.setAttribute "Sello", CFD_Generales.Sello
            If CFD_Generales.Forma_Pago <> "" Then Nodo_Principal.setAttribute "FormaPago", Format(Mid(CFD_Generales.Forma_Pago, 1, 2), "00") 'LCase(CFD_Generales.Forma_Pago)
            Nodo_Principal.setAttribute "NoCertificado", CFD_Generales.No_Certificado
            Nodo_Principal.setAttribute "Certificado", CFD_Generales.Certificado
            If CFD_Generales.Condiciones_Pago <> "" Then
                Nodo_Principal.setAttribute "CondicionesDePago", CFD_Generales.Condiciones_Pago
            End If
            Nodo_Principal.setAttribute "SubTotal", CDbl(Format(CFD_Generales.SubTotal, "#0.00"))
            If CFD_Generales.Descuento > 0 Then
                Nodo_Principal.setAttribute "Descuento", Format(CFD_Generales.Descuento, "#0.00")
            End If
            If CFD_Generales.Tipo_Moneda = "MXN" Or CFD_Generales.Tipo_Moneda = "XXX" Then
                Nodo_Principal.setAttribute "Moneda", CFD_Generales.Tipo_Moneda
            Else
                Nodo_Principal.setAttribute "TipoCambio", CFD_Generales.Tipo_Cambio
                Nodo_Principal.setAttribute "Moneda", CFD_Generales.Tipo_Moneda
            End If
            Nodo_Principal.setAttribute "Total", CDbl(Format(CFD_Generales.Total, "#0.00"))
            Nodo_Principal.setAttribute "TipoDeComprobante", CFD_Generales.Tipo_Comprobante
            If CFD_Generales.Metodo_Pago <> "" Then Nodo_Principal.setAttribute "MetodoPago", CFD_Generales.Metodo_Pago
            Nodo_Principal.setAttribute "LugarExpedicion", cp
            'If CFD_Generales.No_Cuenta_Pago <> "" Then
             '   Nodo_Principal.setAttribute "NumCtaPago", CFD_Generales.No_Cuenta_Pago
            'End If
        'Agrega un salto de linea
        Nodo_Principal.appendChild CFD_Documento.createTextNode(vbCrLf)
        '*****************************************************************************************
        'CREA NODO Comprobantes Relacionados
        '*****************************************************************************************
        'Agrega el elemento del relacionados
        If CFD_Generales.Relacionado = "S" Then
            Set Nodo_Relacionados = CFD_Documento.createElement("cfdi:CfdiRelacionados")
            Nodo_Relacionados.setAttribute "TipoRelacion", Mid(CFD_Generales.Tipo_Relacion, 1, 2)
            Set Nodo_Relacionados_UUID = CFD_Documento.createElement("cfdi:CfdiRelacionado")
            Nodo_Relacionados_UUID.setAttribute "UUID", CFD_Generales.UUID_Relacion
            Nodo_Relacionados.appendChild Nodo_Relacionados_UUID
            Nodo_Principal.appendChild Nodo_Relacionados
            Nodo_Relacionados.appendChild CFD_Documento.createTextNode(vbCrLf)
        End If
        
        
        '*****************************************************************************************
        'CREA NODO EMISOR
        '*****************************************************************************************
        'Agrega el elemento del Emisor
        Set Nodo_Emisor = CFD_Documento.createElement("cfdi:Emisor")
            'Agrega los atributos al nodo Emisor
            Nodo_Emisor.setAttribute "Rfc", CFD_Emisor.RFC
            If CFD_Emisor.Nombre <> "" Then
                Nodo_Emisor.setAttribute "Nombre", CFD_Emisor.Nombre
            End If
            Nodo_Emisor.setAttribute "RegimenFiscal", Mid(CFD_Emisor.Regimen_Fiscal, 1, 3)
            'Agrega un salto de linea
            Nodo_Principal.appendChild Nodo_Emisor
            Nodo_Emisor.appendChild CFD_Documento.createTextNode(vbCrLf)
            
            'Agrega los elementos al Emisor
            'Set Nodo_Emisor_Domicilio = CFD_Documento.createElement("cfdi:DomicilioFiscal")
                'Agrega los atributos del domicilio al nodo Emisor
              '  Nodo_Emisor_Domicilio.setAttribute "calle", CFD_Emisor.Calle
              '  If CFD_Emisor.No_Exterior <> "" Then
              '      Nodo_Emisor_Domicilio.setAttribute "noExterior", CFD_Emisor.No_Exterior
              '  End If
              '  If CFD_Emisor.No_Interior <> "" Then
              '      Nodo_Emisor_Domicilio.setAttribute "noInterior", CFD_Emisor.No_Interior
              '  End If
                'If CFD_Emisor.Colonia <> "" Then
              '      Nodo_Emisor_Domicilio.setAttribute "colonia", CFD_Emisor.Colonia
              '  End If
              '  If CFD_Emisor.Localidad <> "" Then
              '      Nodo_Emisor_Domicilio.setAttribute "localidad", CFD_Emisor.Localidad
              '  End If
              '  Nodo_Emisor_Domicilio.setAttribute "municipio", CFD_Emisor.Municipio
              '  Nodo_Emisor_Domicilio.setAttribute "estado", CFD_Emisor.Estado
              '  Nodo_Emisor_Domicilio.setAttribute "pais", CFD_Emisor.Pais
              '  Nodo_Emisor_Domicilio.setAttribute "codigoPostal", CFD_Emisor.CP
            'Asigna el nodo del domicilio al nodo emisor
          '  Nodo_Emisor.appendChild Nodo_Emisor_Domicilio
        'Agrega un salto de linea
       ' Nodo_Emisor.appendChild CFD_Documento.createTextNode(vbCrLf)
    
        '*****************************************************************************************
        'CREA NODO EXPEDIDOEN
        '*****************************************************************************************
        'If Sucursal = True Then
            'Agrega los elementos al Emisor
         '   Set Nodo_ExpedidoEn = CFD_Documento.createElement("cfdi:ExpedidoEn")
                'Agrega los atributos del domicilio al nodo Emisor
          '      Nodo_ExpedidoEn.setAttribute "calle", CFD_ExpedidoEn.Calle
           '     If CFD_ExpedidoEn.No_Exterior <> "" Then
            '        Nodo_ExpedidoEn.setAttribute "noExterior", CFD_ExpedidoEn.No_Exterior
             '   End If
             '   If CFD_ExpedidoEn.No_Interior <> "" Then
             '       Nodo_ExpedidoEn.setAttribute "noInterior", CFD_ExpedidoEn.No_Interior
             '   End If
             '   If CFD_ExpedidoEn.Colonia <> "" Then
             '       Nodo_ExpedidoEn.setAttribute "colonia", CFD_ExpedidoEn.Colonia
             '   End If
             '   If CFD_ExpedidoEn.Ciudad <> "" Then
             '       Nodo_ExpedidoEn.setAttribute "municipio", CFD_ExpedidoEn.Ciudad
             '   End If
             '   If CFD_ExpedidoEn.Estado <> "" Then
             '      Nodo_ExpedidoEn.setAttribute "estado", CFD_ExpedidoEn.Estado
             '   End If
             '   If CFD_ExpedidoEn.Pais <> "" Then
             '       Nodo_ExpedidoEn.setAttribute "pais", CFD_ExpedidoEn.Pais
             '   End If
             '   If CFD_ExpedidoEn.CP <> "" Then
             '       Nodo_ExpedidoEn.setAttribute "codigoPostal", CFD_ExpedidoEn.CP
             '   End If
               'Asigna el nodo del domicilio al nodo emisor
             '   Nodo_Emisor.appendChild Nodo_ExpedidoEn
            'Agrega un salto de linea
           ' Nodo_Emisor.appendChild CFD_Documento.createTextNode(vbCrLf)
        'End If
        
        '*****************************************************************************************
        'CREA NODO REGIMEN FISCAL
        '*****************************************************************************************
        'Agrega los elementos al Emisor
        Set Nodo_Regimen = CFD_Documento.createElement("cfdi:RegimenFiscal")
            'Agrega los atributos del domicilio al nodo Emisor
            Nodo_Regimen.setAttribute "Regimen", CFD_Emisor.Regimen_Fiscal
            'Asigna el nodo del domicilio al nodo emisor
            Nodo_Emisor.appendChild Nodo_Regimen
        'Agrega un salto de linea
        Nodo_Emisor.appendChild CFD_Documento.createTextNode(vbCrLf)
                      
        'Asigna el nodo Emisor al nodo principal
        Nodo_Principal.appendChild Nodo_Emisor
        'Agrega un salto de linea
        Nodo_Principal.appendChild CFD_Documento.createTextNode(vbCrLf)
            
        '*****************************************************************************************
        'CREA NODO RECEPTOR
        '*****************************************************************************************
        'Agrega el elemento del Receptor
        Set Nodo_Receptor = CFD_Documento.createElement("cfdi:Receptor")
        
            'Agrega los atributos al nodo Emisor
            Nodo_Receptor.setAttribute "Rfc", CFD_Receptor.RFC
            If CFD_Receptor.Nombre <> "" Then
                Nodo_Receptor.setAttribute "Nombre", CFD_Receptor.Nombre
            End If
            Nodo_Receptor.setAttribute "UsoCFDI", Mid(CFD_Receptor.Uso_CFDI, 1, 3)
            'Agrega un salto de linea
            Nodo_Receptor.appendChild CFD_Documento.createTextNode(vbCrLf)
                'Agrega los elementos al Receptor
               ' Set Nodo_Receptor_Domicilio = CFD_Documento.createElement("cfdi:Domicilio")
                
                'Agrega los atributos del domicilio al nodo Receptor
               ' Nodo_Receptor_Domicilio.setAttribute "calle", CFD_Receptor.Calle
               ' If CFD_Receptor.No_Exterior <> "" Then
               '     Nodo_Receptor_Domicilio.setAttribute "noExterior", CFD_Receptor.No_Exterior
               ' End If
               ' If CFD_Receptor.No_Interior <> "" Then
               '     Nodo_Receptor_Domicilio.setAttribute "noInterior", CFD_Receptor.No_Interior
               ' End If
               ' If CFD_Receptor.Colonia <> "" Then
               '     Nodo_Receptor_Domicilio.setAttribute "colonia", CFD_Receptor.Colonia
               ' End If
               ' If CFD_Receptor.Localidad <> "" Then
               '     Nodo_Receptor_Domicilio.setAttribute "localidad", CFD_Receptor.Localidad
               ' End If
               ' Nodo_Receptor_Domicilio.setAttribute "municipio", CFD_Receptor.Municipio
               ' Nodo_Receptor_Domicilio.setAttribute "estado", CFD_Receptor.Estado
               ' Nodo_Receptor_Domicilio.setAttribute "pais", CFD_Receptor.Pais
               ' Nodo_Receptor_Domicilio.setAttribute "codigoPostal", CFD_Receptor.CP
                'Asigna el nodo del domicilio al nodo Receptor
               ' Nodo_Receptor.appendChild Nodo_Receptor_Domicilio
            'Agrega un salto de linea
           ' Nodo_Receptor.appendChild CFD_Documento.createTextNode(vbCrLf)
        
        'Asigna el nodo Receptor al nodo principal
        Nodo_Principal.appendChild Nodo_Receptor
        'Agrega un salto de linea
       ' Nodo_Principal.appendChild CFD_Documento.createTextNode(vbCrLf)
    
        '*****************************************************************************************
        'CREA NODO CONCEPTOS
        '*****************************************************************************************
        'Agrega el elemento de Conceptos
        Set Nodo_Conceptos = CFD_Documento.createElement("cfdi:Conceptos")
        
            'Agrega los atributos del Concepto al nodo Conceptos
            For Conteo_Estructura = 1 To UBound(CFD_Conceptos)
''                'Valida que el concepto tenga cantidad e importe para agregarlo al xml
''                If CFD_Conceptos(Conteo_Estructura).Cantidad > 0 And CFD_Conceptos(Conteo_Estructura).Importe > 0 Then
                    'Agrega un salto de linea
                    Nodo_Conceptos.appendChild CFD_Documento.createTextNode(vbCrLf)
                    'Agrega los elementos al Concepto
                    Set Nodo_Concepto = CFD_Documento.createElement("cfdi:Concepto")
                        'Asigna los atributos
                        Nodo_Concepto.setAttribute "ClaveProdServ", CFD_Conceptos(Conteo_Estructura).No_Identificacion 'Código del producto
                        If CFD_Conceptos(Conteo_Estructura).No_Identificacion <> "" Then
                            Nodo_Concepto.setAttribute "NoIdentificacion", CFD_Conceptos(Conteo_Estructura).No_Identificacion
                        End If
                        Nodo_Concepto.setAttribute "Cantidad", Format(CFD_Conceptos(Conteo_Estructura).Cantidad, "#0.00")
                        Nodo_Concepto.setAttribute "ClaveUnidad", CFD_Conceptos(Conteo_Estructura).Unidad
                        If CFD_Conceptos(Conteo_Estructura).Unidad <> "" Then Nodo_Concepto.setAttribute "Unidad", CFD_Conceptos(Conteo_Estructura).Unidad
                        Nodo_Concepto.setAttribute "Descripcion", CFD_Conceptos(Conteo_Estructura).Descripcion
                        Nodo_Concepto.setAttribute "ValorUnitario", CDbl(Format(CFD_Conceptos(Conteo_Estructura).Valor_Unitario, "#0.00"))
                        Nodo_Concepto.setAttribute "Importe", CDbl(Format(CFD_Conceptos(Conteo_Estructura).Importe, "#0.00"))
                        If CFD_Generales.Descuento > 0 Then
                            Nodo_Concepto.setAttribute "Descuento", Format(CFD_Conceptos(Conteo_Estructura).Importe * CFD_Generales.Descuento, "#0.00")
                        End If
                        
                        Set Nodo_Impuesto_Concepto = CFD_Documento.createElement("cfdi:Impuestos")
                        Set Nodo_Traslado_Concepto = CFD_Documento.createElement("cfdi:Traslados")
                        'Nodo_Concepto.setAttribute "Impuesto", CDbl(Format(CFD_Conceptos(Conteo_Estructura).Importe, "#0.00"))
                        If CFD_Conceptos(Conteo_Estructura).Aplica_IVA <> "" And CFD_Conceptos(Conteo_Estructura).Aplica_IVA <> "NO" Then
                        
                            Set Nodo_Traslado_Conceptos = CFD_Documento.createElement("cfdi:Traslado")
                            If CFD_Conceptos(Conteo_Estructura).Aplica_IVA <> "" Then
                                Nodo_Traslado_Conceptos.setAttribute "Base", CDbl(Format(CFD_Conceptos(Conteo_Estructura).Importe, "#0.00"))
                                Nodo_Traslado_Conceptos.setAttribute "Impuesto", "002"
                                Nodo_Traslado_Conceptos.setAttribute "TipoFactor", "Tasa"
                                Nodo_Traslado_Conceptos.setAttribute "TasaOCuota", Format(Val(CFD_Impuestos(1).Tasa / 100), "#0.000000") 'double
                                Nodo_Traslado_Conceptos.setAttribute "Importe", CDbl(Format(CFD_Conceptos(Conteo_Estructura).Impuesto, "#0.00"))  'double
                                Nodo_Traslado_Concepto.appendChild Nodo_Traslado_Conceptos
                            End If
                        '    If CFD_Conceptos(Conteo_Estructura).IEPS_Producto <> "" And CFD_Conceptos(Conteo_Estructura).IEPS_Producto <> 0 Then
                        '        Set Nodo_Traslado_Conceptos = CFD_Documento.createElement("cfdi:Traslado")
                        '        Nodo_Traslado_Conceptos.setAttribute "Base", CDbl(Format(CFD_Conceptos(Conteo_Estructura).Importe, "#0.00"))
                        '        Nodo_Traslado_Conceptos.setAttribute "Impuesto", "003"
                        '        Nodo_Traslado_Conceptos.setAttribute "TipoFactor", CFD_IEPS(1).Factor
                        '        Nodo_Traslado_Conceptos.setAttribute "TasaOCuota", Val(CFD_IEPS(1).Tasa_IEPS / 100) 'double
                        '        'Nodo_Traslado_Conceptos.setAttribute "Importe", CDbl(Format(CFD_IEPS(Conteo_Estructura).Importe_IEPS, "#0.00"))
                        '        'Importe_IEPS = Consulta_Impuestos("Importe_IEPS", "Porcentaje_IEPS", CFD_IEPS(Conteo_Estructura).Tasa_IEPS)
                        '        Nodo_Traslado_Conceptos.setAttribute "Importe", CDbl(Format(CFD_IEPS(Conteo_Estructura).Importe_IEPS, "#0.00"))
                        '        Nodo_Traslado_Concepto.appendChild Nodo_Traslado_Conceptos
                        '    End If
                        '     'Nodo_Traslado_Concepto.appendChild Nodo_Traslado_Conceptos
                             Nodo_Impuesto_Concepto.appendChild Nodo_Traslado_Concepto
                        End If
                        
                        If UBound(CFD_Impuestos_Retenidos) > 0 Then
                            Set Nodo_Retencion_Concepto = CFD_Documento.createElement("cfdi:Retenciones")
                            For Conteo_Estructura2 = 1 To UBound(CFD_Impuestos_Retenidos)
                                'Agrega los elementos de Retenciones
                                Set Nodo_Retencion_Conceptos = CFD_Documento.createElement("cfdi:Retencion")
                                Nodo_Retencion_Conceptos.setAttribute "Base", CDbl(Format(CFD_Conceptos(Conteo_Estructura).Importe, "#0.00"))
                                Nodo_Retencion_Conceptos.setAttribute "Impuesto", CFD_Impuestos_Retenidos(Conteo_Estructura2).Impuesto
                                Nodo_Retencion_Conceptos.setAttribute "TipoFactor", "Tasa"
                                If CFD_Impuestos_Retenidos(Conteo_Estructura2).Impuesto = "002" Then
                                    Nodo_Retencion_Conceptos.setAttribute "TasaOCuota", Format(Val(CFD_Impuestos_Retenidos(1).Tasa), "#0.000000")
                                    Nodo_Retencion_Conceptos.setAttribute "Importe", Val(CFD_Impuestos_Retenidos(1).Tasa) * CDbl(Format(CFD_Conceptos(Conteo_Estructura).Importe, "#0.00"))
                                Else
                                    Nodo_Retencion_Conceptos.setAttribute "TasaOCuota", Format(Val(CFD_Impuestos_Retenidos(2).Tasa), "#0.000000")
                                    Nodo_Retencion_Conceptos.setAttribute "Importe", CDbl(Format(CFD_Conceptos(Conteo_Estructura).Importe, "#0.00")) * Val(CFD_Impuestos_Retenidos(2).Tasa)
                                End If
                                ' Nodo_Retencion_Conceptos.setAttribute "Importe", Format(CFD_Impuestos_Retenidos(Conteo_Estructura2).Importe, "#0.00")
                                'Asigna el nodo de Traslado al nodo Retenciones
                                Nodo_Retencion_Concepto.appendChild Nodo_Retencion_Conceptos
                                'Agrega un salto de linea
                                Nodo_Retencion_Concepto.appendChild CFD_Documento.createTextNode(vbCrLf)
                            Next Conteo_Estructura2
                        'Cierra el nodo de Retenciones al nodo Impuestos
                        Nodo_Impuesto_Concepto.appendChild Nodo_Retencion_Concepto
                        End If
                        
                        Nodo_Concepto.appendChild Nodo_Impuesto_Concepto
                        Nodo_Conceptos.appendChild Nodo_Concepto
                        'Valida si hay informacion de la cuenta predial que ingresar
                        If CFD_Generales.Tipo_Factura = "ARRENDAMIENTO" Then
                            'If CFD_Conceptos(Conteo_Estructura).No_Predial <> "" And CFD_Conceptos(Conteo_Estructura).No_Predial <> 0 Then
                            'Agrega un salto de linea
                            Nodo_Concepto.appendChild CFD_Documento.createTextNode(vbCrLf)
                            'Agrega los elementos al Concepto
                            'Set Nodo_Cuenta_Predial = CFD_Documento.createElement("cfdi:CuentaPredial")
                                'Asigna los atributos
                            '    Nodo_Cuenta_Predial.setAttribute "Numero", CFD_Conceptos(Conteo_Estructura).No_Predial
                            'Asigna el nodo de Conceptos al nodo Conceptos
                            Nodo_Concepto.appendChild Nodo_Cuenta_Predial
                        End If
                    'Asigna el nodo de Conceptos al nodo Conceptos
                    Nodo_Conceptos.appendChild Nodo_Concepto
''                End If
            Next Conteo_Estructura
            'Agrega un salto de linea
            Nodo_Conceptos.appendChild CFD_Documento.createTextNode(vbCrLf)
        
        'Asigna el nodo Conceptos al nodo principal
        Nodo_Principal.appendChild Nodo_Conceptos
        'Agrega un salto de linea
        Nodo_Principal.appendChild CFD_Documento.createTextNode(vbCrLf)
    
        '*****************************************************************************************
        'CREA NODO IMPUESTOS
        '*****************************************************************************************

        'Agrega el elemento de Impuestos
        Set Nodo_Impuestos = CFD_Documento.createElement("cfdi:Impuestos")
            If IVA_EXENTO = False Then
            
                'Inicializa las variables
                Impuestos_Retenidos = 0
                Impuestos_Trasladados = 0
                
                'Verifica si existe monto de retenciones
                For Conteo_Estructura = 1 To UBound(CFD_Impuestos_Retenidos)
                    Impuestos_Retenidos = Impuestos_Retenidos + CFD_Impuestos_Retenidos(Conteo_Estructura).Importe
                Next Conteo_Estructura
                
                'Verifica si existe monto de trasladados
                For Conteo_Estructura = 1 To UBound(CFD_Impuestos)
                    Impuestos_Trasladados = Impuestos_Trasladados + CFD_Impuestos(Conteo_Estructura).Importe
                Next Conteo_Estructura
                
''                'Agrega los atributos del nodo Impuestos
                Nodo_Impuestos.setAttribute "TotalImpuestosTrasladados", Format(Impuestos_Trasladados, "#0.00")
                
                'si hay impuestos retenidos los toma en cuenta para mostrarlos en el xml
                    
                    'Nodo_Impuestos.setAttribute "TotalImpuestosRetenidos", Format(CFD_Generales.Total_Retenciones, "#0.00")
                    Nodo_Impuestos.setAttribute "TotalImpuestosRetenidos", Format(Impuestos_Retenidos, "#0.00")
                    'Nodo_Impuestos.setAttribute "TotalImpuestosTrasladados", Format(Impuestos_Trasladados, "#0.00")
                    'Agrega los elementos de retenidos
                    Set Nodo_Retenciones = CFD_Documento.createElement("cfdi:Retenciones")
                        'Agrega un salto de linea
                        Nodo_Impuestos.appendChild CFD_Documento.createTextNode(vbCrLf)
                        'Agrega un salto de linea
                        Nodo_Retenciones.appendChild CFD_Documento.createTextNode(vbCrLf)
                        'Agrega los atributos del Concepto al nodo Conceptos
                        For Conteo_Estructura = 1 To UBound(CFD_Impuestos_Retenidos)
                            'Agrega los elementos de Retenciones
                            Set Nodo_Retencion = CFD_Documento.createElement("cfdi:Retencion")
                            'Agrega los atributos de la retencion al nodo Retenciones
                            Nodo_Retencion.setAttribute "Base", Format(CFD_Generales.SubTotal, "#0.00")
                            Nodo_Retencion.setAttribute "Impuesto", CFD_Impuestos_Retenidos(Conteo_Estructura).Impuesto
                            Nodo_Retencion.setAttribute "TipoFactor", "Tasa"
                            If CFD_Impuestos_Retenidos(Conteo_Estructura).Impuesto = "002" Then
                                Nodo_Retencion.setAttribute "TasaOCuota", Format(Val(CFD_Impuestos_Retenidos(1).Tasa), "#0.000000")
                                Nodo_Retencion.setAttribute "Importe", CDbl(Format(CFD_Impuestos_Retenidos(Conteo_Estructura).Importe, "#0.00"))
                             Else
                                    Nodo_Retencion.setAttribute "TasaOCuota", Format(Val(CFD_Impuestos_Retenidos(2).Tasa), "#0.000000")
                                    'Nodo_Retencion.setAttribute "Importe", CDbl(Format(CFD_Impuestos_Retenidos(2).Tasa, "#0.00") * CFD_Generales.Subtotal)
                                    Nodo_Retencion.setAttribute "Importe", CDbl(Format(CFD_Impuestos_Retenidos(Conteo_Estructura).Importe, "#0.00"))
                             End If
                            'Nodo_Retencion.setAttribute "Importe", 'CDbl(Format(CFD_Impuestos_Retenidos(Conteo_Estructura).Importe, "#0.00"))
                            'Asigna el nodo de Traslado al nodo Retenciones
                            Nodo_Retenciones.appendChild Nodo_Retencion
                            'Agrega un salto de linea
                            Nodo_Retenciones.appendChild CFD_Documento.createTextNode(vbCrLf)
                        Next Conteo_Estructura
                    'Cierra el nodo de Retenciones al nodo Impuestos
                    Nodo_Impuestos.appendChild Nodo_Retenciones
                Else
                    'Si no hubo retenciones agrega el atributo de impuestos trasladados
                    Nodo_Impuestos.setAttribute "TotalImpuestosTrasladados", CDbl(Format(Impuestos_Trasladados, "#0.00")) 'double
                End If
                
                'Agrega los elementos de impuestos Traslados
                Set Nodo_Traslados = CFD_Documento.createElement("cfdi:Traslados")
                    'Agrega un salto de linea
                    Nodo_Traslados.appendChild CFD_Documento.createTextNode(vbCrLf)
                    'Agrega los atributos del Concepto al nodo Conceptos
                    For Conteo_Estructura = 1 To UBound(CFD_Impuestos)
                        'Agrega los elementos de Traslado
                        Set Nodo_Traslado = CFD_Documento.createElement("cfdi:Traslado")
                            'Agrega los atributos del Traslado al nodo Traslados
                            'Nodo_Traslado.setAttribute "Base", CDbl(Format(CFD_Impuestos(Conteo_Estructura).Importe, "#0.00")) 'double
                        If CFD_Impuestos(Conteo_Estructura).Impuesto = "002" Then
                            Nodo_Traslado.setAttribute "Base", CDbl(Format(CFD_Generales.SubTotal, "#0.00"))
                            Nodo_Traslado.setAttribute "Impuesto", CFD_Impuestos(Conteo_Estructura).Impuesto
                            Nodo_Traslado.setAttribute "TipoFactor", "Tasa"
                            Nodo_Traslado.setAttribute "TasaOCuota", Format(Val(CFD_Impuestos(1).Tasa / 100), "#0.000000") 'double
                            Nodo_Traslado.setAttribute "Importe", CDbl(Format(CFD_Impuestos(Conteo_Estructura).Importe, "#0.00")) 'double
                          End If
                             'If CFD_Impuestos(Conteo_Estructura).Impuesto = "003" Then
                             'If CFD_Impuestos(2).Importe > 0 Then
                             '   Nodo_Traslado.setAttribute "Base", CDbl(Format(CFD_Generales.SubTotal, "#0.00"))
                             '   Nodo_Traslado.setAttribute "Impuesto", CFD_Impuestos(Conteo_Estructura).Impuesto
                             '
                             '   Nodo_Traslado.setAttribute "TipoFactor", CFD_IEPS(1).Factor
                             '   Nodo_Traslado.setAttribute "TasaOCuota", Val(CFD_IEPS(1).Tasa_IEPS / 100) 'double
                             '   Nodo_Traslado.setAttribute "Importe", CDbl(Format(CFD_Impuestos(2).Importe, "#0.00"))
                           'End If
                            
                        'Asigna el nodo de Traslado al nodo Traslados
                        Nodo_Traslados.appendChild Nodo_Traslado
                        'Agrega un salto de linea
                        Nodo_Traslados.appendChild CFD_Documento.createTextNode(vbCrLf)
                    Next Conteo_Estructura
                'Cierra el nodo de Traslados al nodo Impuestos
                Nodo_Impuestos.appendChild Nodo_Traslados
                'Agrega un salto de linea
                Nodo_Impuestos.appendChild CFD_Documento.createTextNode(vbCrLf)
            'End If
        'Cierra el nodo Impuestos al nodo principal
        Nodo_Principal.appendChild Nodo_Impuestos
    
        '*****************************************************************************************
        'CREA NODO COMPLEMENTO
        '*****************************************************************************************
        Nodo_Principal.appendChild CFD_Documento.createTextNode(vbCrLf)
        'Agrega el elemento de la complemento
        Set Nodo_Complemento = CFD_Documento.createElement("cfdi:Complemento")
            'Agrega un salto de linea
            Nodo_Complemento.appendChild CFD_Documento.createTextNode(vbCrLf)
                
            'Inicializa las variables
            Impuestos_Locales_Retenidos = 0
            Impuestos_Locales_Retenidos = 0
                
            'Verifica si existe monto de impuestos locales
            For Conteo_Estructura = 1 To UBound(CFD_Impuestos_Locales)
                If CFD_Impuestos_Locales(Conteo_Estructura).Tipo = "T" Then
                    Impuestos_Locales_Trasladados = Impuestos_Locales + CFD_Impuestos_Locales(Conteo_Estructura).Importe
                Else
                    Impuestos_Locales_Retenidos = Impuestos_Locales_Retenidos + CFD_Impuestos_Locales(Conteo_Estructura).Importe
                End If
            Next Conteo_Estructura
            
            'Si el valor de impuestos locales es mayor a cero, agrega el nodo Complemento
            If Impuestos_Locales_Retenidos > 0 Or Impuestos_Locales_Trasladados > 0 Then
                'Agrega el elemento Detalle
                Set Nodo_ImpuestoLocal = CFD_Documento.createElement("implocal:ImpuestosLocales")
                    'Agrega el elemento No Proveedor
                    Nodo_ImpuestoLocal.setAttribute "version", "1.0"
                    Nodo_ImpuestoLocal.setAttribute "TotaldeRetenciones", Format(Impuestos_Locales_Retenidos, "#0.00")
                    Nodo_ImpuestoLocal.setAttribute "TotaldeTraslados", Format(Impuestos_Locales_Trasladados, "#0.00")
                    Nodo_ImpuestoLocal.setAttribute "xmlns:implocal", "http://www.sat.gob.mx/implocal"
                'Cierra el nodo Encabezado
                Nodo_Complemento.appendChild Nodo_ImpuestoLocal
                'Agrega un salto de linea
                Nodo_ImpuestoLocal.appendChild CFD_Documento.createTextNode(vbCrLf)
                
                'Agrega los atributos del impuesto local retenido
                For Conteo_Estructura = 1 To UBound(CFD_Impuestos_Locales)
                    'Agrega los elementos al Concepto
                    If CFD_Impuestos_Locales(Conteo_Estructura).Tipo = "R" Then
                        Set Nodo_Elemento = CFD_Documento.createElement("implocal:RetencionesLocales")
                            'Asigna los atributos
                            Nodo_Elemento.setAttribute "ImpLocRetenido", CFD_Impuestos_Locales(Conteo_Estructura).Impuesto
                            Nodo_Elemento.setAttribute "TasadeRetencion", CFD_Impuestos_Locales(Conteo_Estructura).Tasa
                            Nodo_Elemento.setAttribute "Importe", CFD_Impuestos_Locales(Conteo_Estructura).Importe
                        'Asigna el nodo de Conceptos al nodo Conceptos
                        Nodo_ImpuestoLocal.appendChild Nodo_Elemento
                    Else
                        'Agrega los elementos al Concepto
                        Set Nodo_Elemento = CFD_Documento.createElement("implocal:TrasladosLocales")
                            'Asigna los atributos
                            Nodo_Elemento.setAttribute "ImpLocTrasladado", CFD_Impuestos_Locales(Conteo_Estructura).Impuesto
                            Nodo_Elemento.setAttribute "TasadeTraslado", CFD_Impuestos_Locales(Conteo_Estructura).Tasa
                            Nodo_Elemento.setAttribute "Importe", CFD_Impuestos_Locales(Conteo_Estructura).Importe
                        'Asigna el nodo de Conceptos al nodo Conceptos
                        Nodo_ImpuestoLocal.appendChild Nodo_Elemento
                    End If
                Next Conteo_Estructura
                'Agrega un salto de linea
                Nodo_ImpuestoLocal.appendChild CFD_Documento.createTextNode(vbCrLf)
                'Asigna el nodo Conceptos al nodo principal
                Nodo_Complemento.appendChild Nodo_ImpuestoLocal
            End If
            
             If CFD_Generales.Tipo_Factura = "Pagos" Then
            Set Nodo_Pagos = CFD_Documento.createElement("pago10:Pagos")
                Nodo_Pagos.setAttribute "Version", "1.0"
                Nodo_Pagos.appendChild CFD_Documento.createTextNode(vbCrLf)
            Set Nodo_Pago = CFD_Documento.createElement("pago10:Pago")
                Nodo_Pago.setAttribute "FechaPago", CFD_Pagos.Fecha_Pago
                Nodo_Pago.setAttribute "FormaDePagoP", Mid(CFD_Pagos.Forma_Pago_Pagos, 1, 2)
                Nodo_Pago.setAttribute "MonedaP", CFD_Pagos.Moneda_Pago
                If CFD_Pagos.Tipo_Cambio_Pago <> "" And Val(CFD_Pagos.Tipo_Cambio_Pago) > 0 Then Nodo_Pago.setAttribute "TipoCambioP", CFD_Pagos.Tipo_Cambio_Pago
                Nodo_Pago.setAttribute "Monto", CFD_Pagos.Monto
                If Mid(CFD_Pagos.Forma_Pago_Pagos, 1, 2) <> "01" Then
                    If CFD_Pagos.Num_Operacion <> "" Then Nodo_Pago.setAttribute "NumOperacion", CFD_Pagos.Num_Operacion
                    If CFD_Pagos.RfcEmisorCtaOrd <> "" Then Nodo_Pago.setAttribute "RfcEmisorCtaOrd", CFD_Pagos.RfcEmisorCtaOrd
                    If CFD_Pagos.NomBancoOrdExt <> "" Then Nodo_Pago.setAttribute "NomBancoOrdExt", CFD_Pagos.NomBancoOrdExt
                    If CFD_Pagos.CtaOrdenante <> "" Then Nodo_Pago.setAttribute "CtaOrdenante", CFD_Pagos.CtaOrdenante
                    If CFD_Pagos.RfcEmisorCtaBen <> "" Then Nodo_Pago.setAttribute "RfcEmisorCtaBen", CFD_Pagos.RfcEmisorCtaBen
                    If CFD_Pagos.CtaBeneficiario <> "" Then Nodo_Pago.setAttribute "CtaBeneficiario", CFD_Pagos.CtaBeneficiario
                    If CFD_Pagos.TipoCadPago <> "" Then Nodo_Pago.setAttribute "TipoCadPago", Mid(CFD_Pagos.TipoCadPago, 1, 2)
                    If CFD_Pagos.CertPago <> "" Then Nodo_Pago.setAttribute "CertPago", CFD_Pagos.CertPago
                    If CFD_Pagos.CadPago <> "" Then Nodo_Pago.setAttribute "CadPago", CFD_Pagos.CadPago
                    If CFD_Pagos.Sello_Pago <> "" Then Nodo_Pago.setAttribute "SelloPago", CFD_Pagos.Sello_Pago
                End If
                Nodo_Pago.appendChild CFD_Documento.createTextNode(vbCrLf)
                For i = 1 To UBound(CFD_Pagos_DR)
                    Set Nodo_Pagos_Relacionados = CFD_Documento.createElement("pago10:DoctoRelacionado")
                    Nodo_Pagos_Relacionados.setAttribute "IdDocumento", CFD_Pagos_DR(i).ID_Doc
                    If CFD_Pagos_DR(i).Serie <> "" Then Nodo_Pagos_Relacionados.setAttribute "Serie", CFD_Pagos_DR(i).Serie
                    Nodo_Pagos_Relacionados.setAttribute "Folio", CFD_Pagos_DR(i).Folio
                    Nodo_Pagos_Relacionados.setAttribute "MonedaDR", CFD_Pagos_DR(i).Moneda_DR
                    If CFD_Pagos_DR(i).Tipo_Cambio_DR <> "" And Val(CFD_Pagos_DR(i).Tipo_Cambio_DR) > 0 Then Nodo_Pagos_Relacionados.setAttribute "TipoCambioDR", CDbl(Val(CFD_Pagos_DR(i).Tipo_Cambio_DR))
                    Nodo_Pagos_Relacionados.setAttribute "MetodoDePagoDR", Mid(CFD_Pagos_DR(i).Metodo_Pago_DR, 1, 3)
                    Nodo_Pagos_Relacionados.setAttribute "NumParcialidad", CFD_Pagos_DR(i).No_Parcialidad
                    Nodo_Pagos_Relacionados.setAttribute "ImpSaldoAnt", CFD_Pagos_DR(i).Saldo_Anterior
                    Nodo_Pagos_Relacionados.setAttribute "ImpPagado", CFD_Pagos_DR(i).Importe_Pagado 'CFD_Pagos.Importe_Pagado
                    Nodo_Pagos_Relacionados.setAttribute "ImpSaldoInsoluto", CFD_Pagos_DR(i).Saldo_Insoluto
                    Nodo_Pagos_Relacionados.appendChild CFD_Documento.createTextNode(vbCrLf)
                 'Asigna el nodo de relacionados a pago
                        Nodo_Pago.appendChild Nodo_Pagos_Relacionados
                Next
                        'Agrega un salto de linea
                        Nodo_Pago.appendChild CFD_Documento.createTextNode(vbCrLf)
                'Cierra el nodo de Traslados al nodo Impuestos
               
                Nodo_Pagos.appendChild Nodo_Pago
                Nodo_Pagos.appendChild CFD_Documento.createTextNode(vbCrLf)
                Nodo_Complemento.appendChild Nodo_Pagos
           End If
           
            'Si el documento es un documento de donaciones agrega el nodo
            'If Tipo_Documento = "DONATIVOS" Then
            '    Set Nodo_Donaciones = CFD_Documento.createElement("donat:Donatarias")
            '        'Agrega el elemento No Proveedor
            '        Nodo_Donaciones.setAttribute "fechaAutorizacion", Fecha_Autorizacion_Donacion
            '        Nodo_Donaciones.setAttribute "noAutorizacion", No_Autorizacion_Donacion
            '        Nodo_Donaciones.setAttribute "version", "1.1"
            '        Nodo_Donaciones.setAttribute "leyenda", Leyenda_Donacion
            '        Nodo_Donaciones.setAttribute "xmlns:donat", "http://www.sat.gob.mx/donat"
                'Cierra el nodo Encabezado
            '    Nodo_Complemento.appendChild Nodo_Donaciones
                'Agrega un salto de linea
            '    Nodo_Complemento.appendChild CFD_Documento.createTextNode(vbCrLf)
            'End If
        'Cierra el nodo Addenda al nodo principal
        Nodo_Principal.appendChild Nodo_Complemento
        
    'Agrega un salto de linea
    Nodo_Complemento.appendChild CFD_Documento.createTextNode(vbCrLf)
        
    'Cierra el nodo principal
    Nodo_Principal.appendChild CFD_Documento.createTextNode(vbCrLf)
      
    'Limpia las variables del timbre
    
    
    Timbrado_VersionSat = ""
    Timbrado_UUID = ""
    Timbrado_FechaTimbrado = ""
    Timbrado_selloCFD = ""
    Timbrado_noCertificadoSAT = ""
    Timbrado_selloSAT = ""
   
    'Realiza la petición al webservice
   
    If Timbrado_Ambiente = "PRUEBA" Then
       Call Timbrado33.GenerarTimbrado(False, "demo", "12345678", 0, "", "", "", "", CStr(CFD_Documento.XML))
       'Debug.Print CStr(CFD_Documento.XML)
'    Else
'        Call Timbrado33.GenerarTimbrado(True, Timbrado_Codigo_Usuario_Proveedor, Timbrado_Codigo_Usuario, Timbrado_ID_Sucursal, CStr(CFD_Documento.XML))
    End If
    If Timbrado33.EstatusRespuesta = False Then
        MsgBox Timbrado33.MensajeValidacion
        Exit Sub
    End If
    
    XML = Timbrado33.XMLTimbrado
        

    'Pone los datos en el analizador de XML
    Set Conexion_Web_Services = New DOMDocument
    'Parametros de validacion
    Conexion_Web_Services.resolveExternals = True

    Conexion_Web_Services.validateOnParse = True
    
    Conexion_Web_Services.async = False
   
    'Carga la respuesta del webservice
    If Conexion_Web_Services.loadXML(XML) Then  'Si es valida la respuesta
        'Obtiene la respuesta
        Set Respuesta = Conexion_Web_Services.SelectSingleNode("//tfd:TimbreFiscalDigital")
        If Not Respuesta Is Nothing Then
            'Obtiene los parametros del resultado
            Timbrado_VersionSat = Respuesta.Attributes(2).text
            Timbrado_UUID = Respuesta.Attributes(3).text
            Timbrado_FechaTimbrado = Respuesta.Attributes(4).text
            Timbrado_selloCFD = Respuesta.Attributes(6).text
            Timbrado_noCertificadoSAT = Respuesta.Attributes(7).text
            Timbrado_selloSAT = Respuesta.Attributes(8).text
            'Agrega el elemento del timbrado
            Set Nodo_Timbrado = CFD_Documento.createElement("tfd:TimbreFiscalDigital")
                'Agrega los atributos
                Nodo_Timbrado.setAttribute "xsi:schemaLocation", "http://www.sat.gob.mx/TimbreFiscalDigital http://www.sat.gob.mx/TimbreFiscalDigital/TimbreFiscalDigital.xsd"
                Nodo_Timbrado.setAttribute "xmlns:tfd", "http://www.sat.gob.mx/TimbreFiscalDigital"
                Nodo_Timbrado.setAttribute "version", Timbrado_VersionSat
                Nodo_Timbrado.setAttribute "UUID", Timbrado_UUID
                Nodo_Timbrado.setAttribute "FechaTimbrado", Timbrado_FechaTimbrado
                Nodo_Timbrado.setAttribute "selloCFD", Timbrado_selloCFD
                Nodo_Timbrado.setAttribute "noCertificadoSAT", Timbrado_noCertificadoSAT
                Nodo_Timbrado.setAttribute "selloSAT", Timbrado_selloSAT
            'Asigna el nodo timbrado
            Nodo_Complemento.appendChild Nodo_Timbrado
            CFD_Generales.Fecha_Timbrado = Timbrado_FechaTimbrado
        Else 'Si no es valida regresa el error
            'Obtiene la respuesta
            Set Respuesta = Conexion_Web_Services.SelectSingleNode("//Error")
            If Not Respuesta Is Nothing Then
                Err.Raise 7777, "CFD_Crea_Xml", "Error: " & Respuesta.Attributes(0).text & " " & Respuesta.Attributes(1).text
            Else
                Err.Raise 7777, "CFD_Crea_Xml", "Error: No es posible encontrar los parametros de respuesta"
            End If
        End If
    Else 'Si no es valida regresa el error
        Err.Raise 7777, "CFD_Crea_Xml", XML
    End If
    
    'Valida si hay alguna adenda que agregar
    'Select Case Tipo_Adenda
    '    Case "GIGANTE VERDE"
    '        CFD_Addenda_Gigante_Verde
    '    Case "SANTANDER"
    '        CFD_Addenda_Banco_Santander
    '    Case "CINEPOLIS"
    '        CFD_Addenda_Cinepolis
    '    Case "COSTCO"
    '        CFD_Addenda_Costco
    '    Case "HONDA"
    '        CFD_Addenda_Honda
    '    Case "PILGRIMS"
    '        CFD_Addenda_Pilgrims_Pride
    '    Case "AAM"
    '        CFD_Addenda_AAM_Gastos_Indirectos
    'End Select
    
    If Conectar_Ayudante.Valida_Existe_Archivo_Carpeta(Ruta_Xmls, "CARPETA") = False Then
        MkDir Ruta_Xmls & "\"
    End If
    'Guarda el archivo
    CFD_Documento.Save Ruta_Xmls & "\" & Nombre_CFD & ".xml"
    'Destruye el documento
    Set CFD_Documento = Nothing
    'Genera el codigo bidimensinal
    Codigo_Bidimensional = "?re=" & CFD_Emisor.RFC & "&rr=" & CFD_Receptor.RFC & "&tt=" & Format(CFD_Generales.Total, "0000000000.000000") & "&id=" & Timbrado_UUID
    Call GenerateFile(Codigo_Bidimensional, Ruta_Pdfs & "\CFDI_" & Trim(CFD_Generales.Serie & " " & CFD_Generales.Folio) & ".bmp")
End Sub

'''*******************************************************************************
'''NOMBRE_FUNCION: CFD_Addenda_Gigante_Verde
'''DESCRIPCION: Crea la addenda de acuerdo al formato solitado por Gigante Verde
'''PARAMETROS :
'''CREO       : Sergio Godínez Banda
'''FECHA_CREO : 10-Enero-2011
'''MODIFICO   :
'''FECHA_MODIFICO:
'''CAUSA_MODIFICACION:
'''*******************************************************************************
''Private Sub CFD_Addenda_Gigante_Verde()
''Dim Nodo_Addenda As IXMLDOMElement
''Dim Nodo_Orden_Compra As IXMLDOMElement
''Dim Nodo_Elemento As IXMLDOMElement
''
''    '*****************************************************************************************
''    'ENCABEZADO DE LA ADDENDA
''    '*****************************************************************************************
''    'CREA NODO ADDENDA
''    '*****************************************************************************************
''    'Agrega el elemento de la Addenda
''    Set Nodo_Addenda = CFD_Documento.createElement("Addenda")
''    'Agrega un salto de linea
''    Nodo_Addenda.appendChild CFD_Documento.createTextNode(vbCrLf)
''
''        '*****************************************************************************************
''        'CREA NODO ORDENCOMPRA
''        '*****************************************************************************************
''        Set Nodo_Orden_Compra = CFD_Documento.createElement("ordendecompra")
''        Nodo_Orden_Compra.nodeTypedValue = Trim(Orden_Compra_GV)
''        Nodo_Addenda.appendChild Nodo_Orden_Compra
''        Set Nodo_Orden_Compra = Nothing
''        'Agrega un salto de linea
''        Nodo_Addenda.appendChild CFD_Documento.createTextNode(vbCrLf)
''
''    'Cierra el nodo Addenda al nodo principal
''    Nodo_Principal.appendChild Nodo_Addenda
''    'Agrega un salto de linea
''    Nodo_Addenda.appendChild CFD_Documento.createTextNode(vbCrLf)
''    'Agrega un salto de linea
''    Nodo_Principal.appendChild CFD_Documento.createTextNode(vbCrLf)
''End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: CFD_Cancela_Xml
'DESCRIPCIÓN: Crea la esrtuctura xml de acuerdo a los datos para cancelar el CFD
'PARÁMETROS:  Codigo_UUID, pasa el folios fiscal a cancelar
'CREO:        Ismael Prieto Sánchez
'FECHA_CREO:  5/Nov/2006 9:30 am
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function CFD_Cancela_Xml(Codigo_UUID As String, Consulta As Boolean)
Dim XML As String
Dim Respuesta As IXMLDOMNode
Dim Timbrado33 As ICls_Fe_Timbrado
Set Timbrado33 = New Cls_Fe_Timbrado

    'Limpia las variables del timbre
    Timbrado_VersionSat = ""
    Timbrado_UUID = ""
    Timbrado_FechaTimbrado = ""
    Timbrado_selloCFD = ""
    Timbrado_noCertificadoSAT = ""
    Timbrado_selloSAT = ""
      
    'Realiza la petición al webservice
     If Consulta Then
        If Timbrado_Ambiente = "PRUEBA" Then
            Call Timbrado33.ValidaEstatusCFDI(False, Timbrado_Codigo_Usuario, Timbrado_Codigo_Usuario_Proveedor, RFC_Emisor, Codigo_UUID)
        Else
            Call Timbrado33.ValidaEstatusCFDI(True, Timbrado_Codigo_Usuario, Timbrado_Codigo_Usuario_Proveedor, RFC_Emisor, Codigo_UUID)
        End If
    Else
        If Timbrado_Ambiente = "PRUEBA" Then
           Call Timbrado33.CancelaTimbrado(False, Timbrado_Codigo_Usuario, Timbrado_Codigo_Usuario_Proveedor, RFC_Emisor, Codigo_UUID)
        Else
            Call Timbrado33.CancelaTimbrado(True, Timbrado_Codigo_Usuario, Timbrado_Codigo_Usuario_Proveedor, RFC_Emisor, Codigo_UUID)
        End If
    End If
    If Timbrado33.EstatusRespuesta = False Then
        Err.Raise &HFFFFFF01, "Error", Timbrado33.MensajeValidacion
    Else
'        MsgBox Timbrado33.EstatusRespuesta
        CFD_Cancela_Xml = Timbrado33.MensajeValidacion
    End If
End Function

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: Valida_Termino_Folios_Activos
'DESCRIPCIÓN: Revisa si el folio utilizado es el último del rango para cambiar el estatus a terminado
'PARÁMETROS :
'CREO       : Sergio Godínez Banda
'FECHA_CREO : 21-abril-2012
'MODIFICO          :
'FECHA_MODIFICO    :
'CAUSA_MODIFICACIÓN:
'******************************************************************************'
Public Sub Valida_Termino_Folios_Activos(Documento As String, Serie As String, No_Documento As String)
Dim Rs_Consulta_Parametros As rdoResultset
Dim Rs_Consulta_Factura_Folio As rdoResultset  'Manejo de la tabla de folios de facturación
Dim Rs_Modifica_Rango As rdoResultset
Dim Folio_Factura As Double
Dim Folio_Final As Double

On Error GoTo errorHandler

    MDIFrm_Apl_Principal.MousePointer = 11
    Conexion_Base.BeginTrans
    'Consulta el parámetro de aviso de folios de terminación
    Mi_SQL = "SELECT Serie, Folio_Final, Estatus FROM Cat_Parametros_Factura_Electronica_Folios"
    Mi_SQL = Mi_SQL & " WHERE Estatus = 'ACTIVO'"
    Mi_SQL = Mi_SQL & " AND Tipo = '" & Documento & "' AND Serie = '" & Serie & "'"
    Set Rs_Consulta_Parametros = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
        With Rs_Consulta_Parametros
            If Not .EOF Then
                Folio_Final = Val(.rdoColumns("Folio_Final"))
            Else
                Folio_Final = 0
            End If
        End With
    Rs_Consulta_Parametros.Close
    
    If Documento = "FACTURA" Then
        Mi_SQL = "SELECT MAX(No_Factura_Electronica) AS Factura FROM Adm_Clientes_Facturas"
        Set Rs_Consulta_Factura_Folio = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_Consulta_Factura_Folio
                If Not .EOF Then
                    If Not IsNull(.rdoColumns("Factura")) Then
                        Folio_Factura = Val(.rdoColumns("Factura"))
                    Else
                        Folio_Factura = 0
                    End If
                Else
                    Folio_Factura = 0
                End If
            End With
        Rs_Consulta_Factura_Folio.Close
    Else
        Mi_SQL = "SELECT MAX(No_Nota_Credito) AS NC FROM Adm_Notas_Credito"
        Set Rs_Consulta_Factura_Folio = Conectar_Ayudante.Recordset_Consultar(Mi_SQL)
            With Rs_Consulta_Factura_Folio
                If Not .EOF Then
                    If Not IsNull(.rdoColumns("NC")) Then
                        Folio_Factura = Val(.rdoColumns("NC"))
                    Else
                        Folio_Factura = 0
                    End If
                Else
                    Folio_Factura = 0
                End If
            End With
        Rs_Consulta_Factura_Folio.Close
    End If
    
    'Comapara si el folio de la factura actual es el último del rango actualmente activo
    If Folio_Final = Folio_Factura Then
        'Cambia de estatus la serie actual
        Mi_SQL = "SELECT * FROM Cat_Parametros_Factura_Electronica_Folios"
        Mi_SQL = Mi_SQL & " WHERE Estatus = 'ACTIVO'"
        Mi_SQL = Mi_SQL & " AND Tipo = '" & Documento & "' AND Serie = '" & Serie & "'"
        Mi_SQL = Mi_SQL & " AND Folio_Final = " & Folio_Factura & ""
        Set Rs_Modifica_Rango = Conectar_Ayudante.Recordset_Editar(Mi_SQL)
            With Rs_Modifica_Rango
                If Not .EOF Then
                    .Edit
                        .rdoColumns("Estatus") = "TERMINADO"
                    .Update
                End If
            End With
        Rs_Modifica_Rango.Close
        Conexion_Base.CommitTrans
        MsgBox "El último folio del rango activo actual ha sido utilizado" & Chr(13) & _
               "    Favor de asignar un nuevo rango de folios activos", vbExclamation
               
    End If
    MDIFrm_Apl_Principal.MousePointer = 0
    Exit Sub
errorHandler:
    MDIFrm_Apl_Principal.MousePointer = 0
    Conexion_Base.RollbackTrans
    MsgBox "Ocurrió un error durante el cambio de estatus de la serie actual, favor de dar aviso al administrador", vbExclamation
End Sub

'*******************************************************************************
'NOMBRE DE LA FUNCIÓN: CFD_Cadena_Original
'DESCRIPCIÓN: Crea la cadena original del CFDI de acuerdo a los datos asignados
'PARÁMETROS: Devuelve la cadena compuesta en la función
'CREO:       Sergio Godínez Banda
'FECHA_CREO: 25-Mayo-2012
'MODIFICO:
'FECHA_MODIFICO
'CAUSA_MODIFICACIÓN
'*******************************************************************************
Public Function CFD_Cadena_Original(Tipo_Factura As String)
Dim Cadena_Original As String
Dim cadena As String
Dim Conteo_Conceptos As Integer
Dim Conteo_Estructura2 As Integer
Dim Conteo_Impuestos As Integer
Dim Grupo_Cadena() As String
Dim Suma_Retenciones As Double
Dim Suma_Locales_Retenciones As Double
Dim Suma_Locales_Trasladados As Double
Dim Retenciones_Locales As String
Dim Trasladados_Locales As String
Dim Suma_Trasladados As Double
Dim i As Long

    'Datos Generales
    Cadena_Original = ""
    Cadena_Original = "||" & CFD_Generales.Version
   ' Cadena_Original = Cadena_Original & "|" & CFD_Generales.Version
    If CFD_Generales.Serie <> "" Then Cadena_Original = Cadena_Original & "|" & CFD_Generales.Serie
    Cadena_Original = Cadena_Original & "|" & CFD_Generales.Folio
    Cadena_Original = Cadena_Original & "|" & CFD_Generales.Fecha
    'Cadena_Original = Cadena_Original & "|" & CFD_Generales.Fecha
    'Cadena_Original = Cadena_Original & "|" & CFD_Generales.Sello
    If CFD_Generales.Metodo_Pago <> "" Then Cadena_Original = Cadena_Original & "|" & Mid(CFD_Generales.Forma_Pago, 1, 2)
   Cadena_Original = Cadena_Original & "|" & CFD_Generales.No_Certificado
    'Cadena_Original = Cadena_Original & "|" & CFD_Generales.Condiciones_Pago
    
    'Cadena_Original = Cadena_Original & "|" & CFD_Generales.Tipo_Comprobante
    'Cadena_Original = Cadena_Original & "|" & CFD_Generales.Forma_Pago 'LCase(CFD_Generales.Forma_Pago)
    If CFD_Generales.Condiciones_Pago <> "" Then
        Cadena_Original = Cadena_Original & "|" & CFD_Generales.Condiciones_Pago
    End If
    Cadena_Original = Cadena_Original & "|" & CDbl(Format(CFD_Generales.SubTotal, "#0.00"))
    If CFD_Generales.Descuento > 0 Then
        Cadena_Original = Cadena_Original & "|" & Format(CFD_Generales.Descuento, "#0.00")
    End If
    Cadena_Original = Cadena_Original & "|" & CFD_Generales.Tipo_Moneda
    If CFD_Generales.Tipo_Moneda <> "MXN" And CFD_Generales.Tipo_Moneda <> "XXX" Then
        Cadena_Original = Cadena_Original & "|" & CFD_Generales.Tipo_Cambio
    End If
    'Cadena_Original = Cadena_Original & "|" & CFD_Generales.Tipo_Moneda
    
    Cadena_Original = Cadena_Original & "|" & CDbl(Format(CFD_Generales.Total, "#0.00"))
    Cadena_Original = Cadena_Original & "|" & CFD_Generales.Tipo_Comprobante
    If CFD_Generales.Metodo_Pago <> "" Then Cadena_Original = Cadena_Original & "|" & Mid(CFD_Generales.Metodo_Pago, 1, 3)

    'Cadena_Original = Cadena_Original & "|" & CFD_Generales.Forma_Pago
    Cadena_Original = Cadena_Original & "|" & CFD_Emisor.cp
    'If CFD_Generales.No_Cuenta_Pago <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & CFD_Generales.No_Cuenta_Pago
    'End If
    If CFD_Relacionados.Existe = True And UBound(CFD_Relacionados_Conceptos) > 0 Then
            Cadena_Original = Cadena_Original & "|" & Mid(CFD_Relacionados.Relacionados, 1, 2)
            For i = 1 To UBound(CFD_Relacionados_Conceptos)
                Cadena_Original = Cadena_Original & "|" & CFD_Relacionados_Conceptos(i).UUID
            Next i
    ElseIf CFD_Relacionados.Existe = True Then
            Cadena_Original = Cadena_Original & "|" & Mid(CFD_Relacionados.Relacionados, 1, 2)
            Cadena_Original = Cadena_Original & "|" & CFD_Relacionados.UUID_Relacionados
    End If
    'Datos del Emisor
    
    Cadena_Original = Cadena_Original & "|" & CFD_Emisor.RFC
    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Emisor.Nombre)
    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(Trim(Mid(CFD_Emisor.Regimen_Fiscal, 1, 4)))
    'If CFD_Emisor.No_Exterior <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & CFD_Emisor.No_Exterior
    'End If
    'If CFD_Emisor.No_Interior <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & CFD_Emisor.No_Interior
    'End If
    'If CFD_Emisor.Colonia <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Emisor.Colonia)
    'End If
    'Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Emisor.Localidad)
    'If CFD_Emisor.Referencia <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Emisor.Referencia)
    'End If
    'Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Emisor.Municipio)
    'Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Emisor.Estado)
    'Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Emisor.Pais)
    'Cadena_Original = Cadena_Original & "|" & CFD_Emisor.CP
    
    'Datos de Expedido En
    'If CFD_ExpedidoEn.Calle <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_ExpedidoEn.Calle)
    'End If
    'If CFD_ExpedidoEn.No_Exterior <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & CFD_ExpedidoEn.No_Exterior
    'End If
    'If CFD_ExpedidoEn.No_Interior <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & CFD_ExpedidoEn.No_Interior
    'End If
    'If CFD_ExpedidoEn.Colonia <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_ExpedidoEn.Colonia)
    'End If
'    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Emisor_Expedido_En.Localidad)
'    If CFD_Emisor_Expedido_En.Referencia <> "" Then
'        Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Emisor_Expedido_En.Referencia)
'    End If
    'If CFD_ExpedidoEn.Ciudad <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_ExpedidoEn.Ciudad)
    'End If
    'If CFD_ExpedidoEn.Estado <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_ExpedidoEn.Estado)
    'End If
    'If CFD_ExpedidoEn.Pais <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_ExpedidoEn.Pais)
    'End If
    'If CFD_ExpedidoEn.CP <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & CFD_ExpedidoEn.CP
    'End If
    'Cadena_Original = Cadena_Original & "|" & Regimen_Fiscal
    'If Regimen_Fiscal_2 <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & Regimen_Fiscal_2
    'End If
    'Datos del Receptor
    
    Cadena_Original = Cadena_Original & "|" & CFD_Receptor.RFC
    Cadena_Original = Cadena_Original & "|" & Trim(CFD_Receptor.Nombre)
    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(Mid(CFD_Receptor.Uso_CFDI, 1, 3))
    
    'Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Receptor.Calle)
    'If CFD_Receptor.No_Exterior <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & CFD_Receptor.No_Exterior
    'End If
    'If CFD_Receptor.No_Interior <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & CFD_Receptor.No_Interior
    'End If
    'If CFD_Receptor.Colonia <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Receptor.Colonia)
    'End If
    'If CFD_Receptor.Localidad <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Receptor.Localidad)
    'End If
    'If CFD_Receptor.Referencia <> "" Then
    '    Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales((CFD_Receptor.Referencia))
    'End If
    'Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Receptor.Municipio)
    'Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Receptor.Estado)
    'Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Receptor.Pais)
    'Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Receptor.CP)
    
    'Datos de los Conceptos
    ReDim Valores(UBound(CFD_Conceptos))
    Suma_Trasladados = 0
   For Conteo_Conceptos = 1 To UBound(CFD_Conceptos)
        Cadena_Original = Cadena_Original & "|" & Mid(CFD_Conceptos(Conteo_Conceptos).Cod_prod, 1, 8)
        If CFD_Conceptos(Conteo_Conceptos).No_Identificacion <> "" Then
            Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Conceptos(Conteo_Conceptos).No_Identificacion)
        End If
        Cadena_Original = Cadena_Original & "|" & Format(CFD_Conceptos(Conteo_Conceptos).Cantidad, "#0.00")
        Cadena_Original = Cadena_Original & "|" & CFD_Conceptos(Conteo_Conceptos).Unidad_Medida
        If CFD_Conceptos(Conteo_Conceptos).Unidad <> "" Then
            Cadena_Original = Cadena_Original & "|" & Trim(Mid(CFD_Conceptos(Conteo_Conceptos).Unidad, 1, 20))
        End If
        Cadena_Original = Cadena_Original & "|" & CFD_Conceptos(Conteo_Conceptos).Descripcion
        Cadena_Original = Cadena_Original & "|" & CDbl(Format(CFD_Conceptos(Conteo_Conceptos).Valor_Unitario, "#0.00"))
        Cadena_Original = Cadena_Original & "|" & CDbl(Format(CFD_Conceptos(Conteo_Conceptos).Importe, "#0.00"))
        If CFD_Generales.Descuento > 0 Then
                            Cadena_Original = Cadena_Original & "|" & CFD_Generales.Descuento / UBound(CFD_Conceptos) 'Format(CFD_Conceptos(Conteo_Conceptos).Importe * CFD_Generales.Descuento, "#0.00")
                        End If
        
        If CFD_Conceptos(Conteo_Conceptos).IVA_Producto = True And IVA_EXENTO = False Then
            
            'Cadena_Original = Cadena_Original & "|" & CDbl(Format(CFD_Conceptos(Conteo_Conceptos).IVA_Producto, "#0.00"))
            Cadena_Original = Cadena_Original & "|" & CDbl(Format(CFD_Conceptos(Conteo_Conceptos).Importe, "#0.00"))
            Cadena_Original = Cadena_Original & "|" & "002"
            Cadena_Original = Cadena_Original & "|" & "Tasa"
            Cadena_Original = Cadena_Original & "|" & Format(CFD_Impuestos(1).Tasa, "#0.000000") 'Val(CFD_Generales.Tasa_IVA / 100)
            Cadena_Original = Cadena_Original & "|" & Format(Format(CFD_Conceptos(Conteo_Conceptos).Importe, "#0.00") * CFD_Impuestos(1).Tasa, "#0.00") 'Val(CFD_Generales.Tasa_IVA / 100)
            Suma_Trasladados = Suma_Trasladados + Format(Format(CFD_Conceptos(Conteo_Conceptos).Importe, "#0.00") * CFD_Impuestos(1).Tasa, "#0.00")
            End If
        
    Next Conteo_Conceptos
    
    If UBound(CFD_Impuestos_Retenidos) > 0 Then
        For Conteo_Estructura2 = 1 To UBound(CFD_Impuestos_Retenidos)
                                Cadena_Original = Cadena_Original & "|" & Val(CFD_Generales.SubTotal)
                                Cadena_Original = Cadena_Original & "|" & CFD_Impuestos_Retenidos(Conteo_Estructura2).Impuesto
                                Cadena_Original = Cadena_Original & "|" & "Tasa"
                                If CFD_Impuestos_Retenidos(Conteo_Estructura2).Impuesto = "002" Then
                                    Cadena_Original = Cadena_Original & "|" & Format(Val(CFD_Impuestos_Retenidos(1).Tasa), "#0.000000")
                                Else
                                    Cadena_Original = Cadena_Original & "|" & Format(Val(CFD_Impuestos_Retenidos(2).Tasa), "#0.000000")
                                End If
                                Cadena_Original = Cadena_Original & "|" & Val(Format(CFD_Impuestos_Retenidos(Conteo_Estructura2).Importe, "#0.00"))
                            Next Conteo_Estructura2
    End If
    
    'Datos de Retencion
    If UBound(CFD_Impuestos_Retenidos) > 0 Then
'        Suma_Retenciones = 0
        For Conteo_Conceptos = 1 To UBound(CFD_Impuestos_Retenidos)
            'Cadena_Original = Cadena_Original & "|" & Format(Val(CFD_Impuestos_Retenidos(Conteo_Conceptos).Impuesto), "#0.00") & "|" & Format(Val(CFD_Impuestos_Retenidos(Conteo_Conceptos).Importe), "#0.00") '& "|"
'            Suma_Retenciones = Suma_Retenciones + Format(Val(CFD_Impuestos_Retenidos(Conteo_Conceptos).Importe), "#0.00")
            
            'Cadena_Original = Cadena_Original & "|" & Format(CFD_Generales.Subtotal, "#0.00")
            'Cadena_Original = Cadena_Original & "|" & CFD_Impuestos_Retenidos(Conteo_Conceptos).Impuesto
            'Cadena_Original = Cadena_Original & "|" & "Tasa"
            'Cadena = Cadena & "|" & Format(CFD_Generales.SubTotal, "#0.00")
            cadena = cadena & "|" & CFD_Impuestos_Retenidos(Conteo_Conceptos).Impuesto
            'Cadena = Cadena & "|" & "Tasa"
'            If CFD_Impuestos_Retenidos(Conteo_Conceptos).impuesto = "002" Then
'               'Cadena_Original = Cadena_Original & "|" & Format(Val(CFD_Impuestos_Retenidos(1).Tasa / 100), "#0.00")
'               Cadena = Cadena & "|" & Val(CFD_Impuestos_Retenidos(1).Tasa)
'                Else
'                    'Cadena_Original = Cadena_Original & "|" & Format(Val(CFD_Impuestos_Retenidos(2).Tasa / 100), "#0.00")
'                    Cadena = Cadena & "|" & Val(CFD_Impuestos_Retenidos(2).Tasa)
'                End If
            'Cadena_Original = Cadena_Original & "|" & Format(CFD_Impuestos_Retenidos(Conteo_Conceptos).Importe, "#0.00")
            cadena = cadena & "|" & Val(Format(CFD_Impuestos_Retenidos(Conteo_Conceptos).Importe, "#0.00"))
            
        Next
        Cadena_Original = Cadena_Original & cadena & "|" & Format(Suma_Retenciones, "#0.00")
        'Cadena_Original = Cadena_Original & "|" & Format(CFD_Generales.Total_Retenciones, "#0.00")
        ElseIf UBound(CFD_Impuestos) > 0 Then
            Cadena_Original = Cadena_Original & "|" & Format(0, "#0.00")
    End If
   cadena = ""
    'Datos de los Impuestos
   If UBound(CFD_Impuestos) > 0 Then
        'Cadena_Original = Cadena_Original & "|" '& "Tasa"
        For Conteo_Conceptos = 1 To UBound(CFD_Impuestos)
'            Suma_Trasladados = Suma_Trasladados + CDbl(CFD_Impuestos(Conteo_Conceptos).Importe)
            'If UBound(CFD_Impuestos) = Conteo_Conceptos Then
                
                'Cadena_Original = Cadena_Original & "|"
                'Cadena_Original = Cadena_Original & "|" & CDbl(Format(CFD_Impuestos(Conteo_Conceptos).Importe, "#0.00"))
                'Cadena_Original = Cadena_Original & "|" & CDbl(Format(Val(CFD_Generales.Subtotal), "#0.00"))
                'Cadena_Original = Cadena_Original & "|" & CFD_Impuestos(Conteo_Conceptos).Impuesto
                'Cadena_Original = Cadena_Original & "|" & "Tasa"
                'Cadena_Original = Cadena_Original & "|" & Format(Val(CFD_Generales.Tasa_IVA / 100), "#0.00")
                'Cadena_Original = Cadena_Original & "|" & CDbl(Format(Val(CFD_Impuestos(Conteo_Conceptos).Importe), "#0.00"))
                If CFD_Impuestos(Conteo_Conceptos).Impuesto = "002" Then
                'Cadena = Cadena & "|" & CDbl(Format(Val(CFD_Generales.SubTotal), "#0.00"))
                cadena = cadena & "|" & CFD_Impuestos(Conteo_Conceptos).Impuesto
                cadena = cadena & "|" & "Tasa"
                cadena = cadena & "|" & Format(CFD_Impuestos(1).Tasa, "#0.000000") '   Format(Val(CFD_Generales.Tasa_IVA / 100), "#0.00")
                cadena = cadena & "|" & Format(Suma_Trasladados, "#0.00") 'Format(Val(CFD_Impuestos(Conteo_Conceptos).Importe), "#0.00")
                End If
                
'                If CFD_Impuestos(Conteo_Conceptos).Impuesto = "003" Then
'                             'If CFD_Impuestos(2).Importe > 0 Then
'                                Cadena = Cadena & "|" & CDbl(Format(CFD_Generales.SubTotal, "#0.00"))
'                                Cadena = Cadena & "|" & CFD_Impuestos(Conteo_Conceptos).Impuesto
'                                Cadena = Cadena & "|" & CFD_IEPS(1).factor
'                                Cadena = Cadena & "|" & Val(CFD_IEPS(1).Tasa_IEPS / 100) 'double
'                                Cadena = Cadena & "|" & CDbl(Format(CFD_Impuestos(2).Importe, "#0.00"))
'                            End If
                '& "|" & CDbl(Format(Suma_Trasladados, "#0.00")) 'Format(CFD_Impuestos(Conteo_Conceptos).Importe, "#0.00")
            'Else
                'Cadena_Original = Cadena_Original & "|" & CFD_Impuestos(Conteo_Conceptos).Impuesto & "|" & CDbl(Format(CFD_Impuestos(Conteo_Conceptos).Tasa, "#0.00")) & "|" & CDbl(Format(CFD_Impuestos(Conteo_Conceptos).Importe, "#0.00"))
            'End If
        Next Conteo_Conceptos
        Cadena_Original = Cadena_Original & cadena & "|" & Format(Suma_Trasladados, "#0.00") '& Cadena
       ' Cadena_Original = Cadena_Original & Cadena
    End If
    
    'Complemento de pagos
    If CFD_Generales.Tipo_Factura = "Pagos" Then
            Cadena_Original = Cadena_Original & "|" & "1.0"
            Cadena_Original = Cadena_Original & "|" & CFD_Pagos.Fecha_Pago
            Cadena_Original = Cadena_Original & "|" & Mid(CFD_Pagos.Forma_Pago_Pagos, 1, 2)
            Cadena_Original = Cadena_Original & "|" & CFD_Pagos.Moneda_Pago
            If CFD_Pagos.Tipo_Cambio_Pago <> "" And Val(CFD_Pagos.Tipo_Cambio_Pago) > 0 Then Cadena_Original = Cadena_Original & "|" & CFD_Pagos.Tipo_Cambio_Pago
                Cadena_Original = Cadena_Original & "|" & CFD_Pagos.Monto
                If CFD_Pagos.Num_Operacion <> "" Then Cadena_Original = Cadena_Original & "|" & CFD_Pagos.Num_Operacion
                If Mid(CFD_Pagos.Forma_Pago_Pagos, 1, 2) <> "01" Then
                    If CFD_Pagos.RfcEmisorCtaOrd <> "" Then Cadena_Original = Cadena_Original & "|" & CFD_Pagos.RfcEmisorCtaOrd
                    If CFD_Pagos.NomBancoOrdExt <> "" Then Cadena_Original = Cadena_Original & "|" & CFD_Pagos.NomBancoOrdExt
                    If CFD_Pagos.CtaOrdenante <> "" Then Cadena_Original = Cadena_Original & "|" & CFD_Pagos.CtaOrdenante
                    If CFD_Pagos.RfcEmisorCtaBen <> "" Then Cadena_Original = Cadena_Original & "|" & CFD_Pagos.RfcEmisorCtaBen
                    If CFD_Pagos.CtaBeneficiario <> "" Then Cadena_Original = Cadena_Original & "|" & CFD_Pagos.CtaBeneficiario
                    If CFD_Pagos.TipoCadPago <> "" Then Cadena_Original = Cadena_Original & "|" & Mid(CFD_Pagos.TipoCadPago, 1, 2)
                    If CFD_Pagos.CertPago <> "" Then Cadena_Original = Cadena_Original & "|" & CFD_Pagos.CertPago
                    If CFD_Pagos.CadPago <> "" Then Cadena_Original = Cadena_Original & "|" & CFD_Pagos.CadPago
                    If CFD_Pagos.Sello_Pago <> "" Then Cadena_Original = Cadena_Original & "|" & CFD_Pagos.Sello_Pago
                End If
                For i = 1 To UBound(CFD_Pagos_DR)
                    Cadena_Original = Cadena_Original & "|" & CFD_Pagos_DR(i).ID_Doc
                    If CFD_Pagos_DR(i).Serie <> "" Then Cadena_Original = Cadena_Original & "|" & CFD_Pagos_DR(i).Serie
                    Cadena_Original = Cadena_Original & "|" & CFD_Pagos_DR(i).Folio
                    Cadena_Original = Cadena_Original & "|" & CFD_Pagos_DR(i).Moneda_DR
                    If CFD_Pagos_DR(i).Tipo_Cambio_DR <> "" And Val(CFD_Pagos_DR(i).Tipo_Cambio_DR) > 0 Then Cadena_Original = Cadena_Original & "|" & CDbl(Val(CFD_Pagos_DR(i).Tipo_Cambio_DR))
                    Cadena_Original = Cadena_Original & "|" & Mid(CFD_Pagos_DR(i).Metodo_Pago_DR, 1, 3)
                    Cadena_Original = Cadena_Original & "|" & CFD_Pagos_DR(i).No_Parcialidad
                    Cadena_Original = Cadena_Original & "|" & CFD_Pagos_DR(i).Saldo_Anterior
                    Cadena_Original = Cadena_Original & "|" & CFD_Pagos_DR(i).Importe_Pagado
                    Cadena_Original = Cadena_Original & "|" & CFD_Pagos_DR(i).Saldo_Insoluto
                Next
           End If
'     End If
'    Datos de Retencion de impuestos locales
'    If UBound(CFD_Impuestos_Locales) > 0 Then
'        Suma_Locales_Retenciones = 0
'        Suma_Locales_Trasladados = 0
'        Retenciones_Locales = ""
'        Trasladados_Locales = ""
'        'Consulta el importe de cada retenciones y trasladados
'        For Conteo_Conceptos = 1 To UBound(CFD_Impuestos_Locales)
'            If CFD_Impuestos_Locales(Conteo_Conceptos).Tipo = "R" Then
'                Suma_Locales_Retenciones = Suma_Locales_Retenciones + CDbl(CFD_Impuestos_Locales(Conteo_Conceptos).Importe)
'                Retenciones_Locales = Retenciones_Locales & CFD_Impuestos_Locales(Conteo_Conceptos).impuesto & "|" & CDbl(CFD_Impuestos_Locales(Conteo_Conceptos).Tasa) & "|" & CDbl(CFD_Impuestos_Locales(Conteo_Conceptos).Importe)
'            Else
'                Suma_Locales_Trasladados = Suma_Locales_Trasladados + CDbl(CFD_Impuestos_Locales(Conteo_Conceptos).Importe)
'                Trasladados_Locales = Trasladados_Locales & CFD_Impuestos_Locales(Conteo_Conceptos).impuesto & "|" & CDbl(CFD_Impuestos_Locales(Conteo_Conceptos).Tasa) & "|" & CDbl(CFD_Impuestos_Locales(Conteo_Conceptos).Importe)
'            End If
'        Next Conteo_Conceptos
'        'If CFD_Conceptos(Conteo_Conceptos).No_Predial <> "" Then
'         '   Cadena_Original = Cadena_Original & "|" & Valida_Caracteres_Especiales(CFD_Conceptos(Conteo_Conceptos).No_Predial)
'        'End If
'        'Estructura = version|total de retenciones|total de traslados|impuesto retenido|tasa de retencion|importe|impuesto trasladado|tasa de traslado|importe
'        Cadena_Original = Cadena_Original & "|" & "1.0" & "|" & Format(Suma_Locales_Retenciones, "#0.00") & "|" & Format(Suma_Locales_Trasladados, "#0.00")
'        Debug.Print Cadena_Original
'        If Retenciones_Locales <> "" Then
'            Cadena_Original = Cadena_Original & "|" & Retenciones_Locales
'        End If
'        If Trasladados_Locales <> "" Then
'            Cadena_Original = Cadena_Original & "|" & Trasladados_Locales
'        End If
'        Cadena_Original = Cadena_Original & "||"
'    Else
'        If Tipo_Factura = "DONATIVOS" Then
'            'Estructura = version|número de autorizacíon para emitir recibos de donación|fecha de autorización para emitir recibos de donación|leyenda de donación
'            Cadena_Original = Cadena_Original & "|" & "1.1" & "|" & No_Autorizacion_Donacion & "|" & Fecha_Autorizacion_Donacion & "|" & Leyenda_Donacion & "||"
'        Else
            Cadena_Original = Cadena_Original & "||"
'        End If
'    End If

    'Regresa la cadena original
    Debug.Print (Cadena_Original)
    CFD_Cadena_Original = Cadena_Original
    
End Function

