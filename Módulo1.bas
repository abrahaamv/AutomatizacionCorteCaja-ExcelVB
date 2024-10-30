Attribute VB_Name = "Módulo1"
' Variable global que verifica si se cobró en dólares
Dim cobroEnDolares As Boolean

' Variables globales que verifican si hay diferencia en cheques
Dim diferenciaEnCheques As Boolean
Dim totalChequesEnReporte As Double

' Variables globales que verifican si hay diferencia en tarjetas
Dim diferenciaEnTarjetas As Boolean
Dim totalTarjetasEnReporte As Double

' Variables globales que verifican si hay diferencia en depósitos
Dim diferenciaEnDepositos As Boolean
Dim totalDepositosEnReporte As Double

' Variables globales que verifican si se cobró un recibo en Quetzales con moneda de Dólares
Dim cobroReciboQuetzalesEnDolares As Boolean



' Ejecuta todos los procedimientos para llenar el Corte de caja

Sub CorteDeCaja()
    ' Mostrar un mensaje de confirmación
    Dim respuesta As VbMsgBoxResult
   
    respuesta = MsgBox("¡Estás a punto de generar un Corte de Caja!" & vbCrLf & "" & vbCrLf & "Asegurate de haber llenado antes manualmente los siguientes datos:" & vbCrLf & "" & vbCrLf & "- Cheques" & vbCrLf & "- Tarjetas" & vbCrLf & "- Depósitos" & vbCrLf & "- Recibos por abono a facturas anuladas" & vbCrLf & "" & vbCrLf & "Asegúrate también de haber exportado el corte SAP como Libro de Excel (.xlsx) en la misma ubicación que este archivo, sin cambiarle el nombre" & vbCrLf & "" & vbCrLf & "¿Estás seguro de que deseas continuar?", vbQuestion + vbYesNo, "Confirmación")
    
    ' Comprobar la respuesta del usuario
    If respuesta = vbYes Then
        
    Dim wbCorteCaja As Workbook
    Dim wbReporte As Workbook
    Dim wsCorteCaja As Worksheet
    Dim wsReporte As Worksheet

    ' Abre el Reporte de Corte de Caja
    Set wbReporte = Workbooks.Open(ThisWorkbook.Path & "\Reporte de Corte de Caja.xlsx")
   
    
    ' Especifica las hojas de trabajo
    Set wsCorteCaja = ThisWorkbook.Sheets("CORTE CANELLA")
    Set wsReporte = wbReporte.Sheets("Sheet1")
    
    'Establece la fecha del reporte en el Corte de Caja
     
    Dim fechaEnReporte As String
    Dim fechaActual As Date
    

    ' Leer la fecha y hora en el reporte
    fechaEnReporte = wsReporte.Range("O1").Value
  


    ' Dividir la cadena usando los dos puntos como separador
    Dim partes() As String
        partes = Split(fechaEnReporte, ":")
    
    If UBound(partes) >= 1 Then
        ' Quita el texto de la fecha
        fechaActual = CDate(Trim(partes(1)))
        
        Else
            ' Si el proceso da error, se usará la fecha actual
            fechaActual = Date
    End If


    
     wsCorteCaja.Range("C8").Value = Day(fechaActual)
     wsCorteCaja.Range("G8").Value = Year(fechaActual)


     Dim Mes As String
     Dim MesNumero As Integer
     MesNumero = Month(fechaActual)
      Select Case MesNumero
        Case 1
            Mes = "ENERO"
        Case 2
            Mes = "FEBRERO"
        Case 3
            Mes = "MARZO"
        Case 4
            Mes = "ABRIL"
        Case 5
            Mes = "MAYO"
        Case 6
            Mes = "JUNIO"
        Case 7
            Mes = "JULIO"
        Case 8
            Mes = "AGOSTO"
        Case 9
            Mes = "SEPTIEMBRE"
        Case 10
            Mes = "OCTUBRE"
        Case 11
            Mes = "NOVIEMBRE"
        Case 12
            Mes = "DICIEMBRE"
    End Select
    wsCorteCaja.Range("E8").Value = Mes
    
    ' Calcula el total de Efectivo en Facturas Contado y Recibos
    totalEfectivo wsCorteCaja, wsReporte
    
    ' Calcula el total de cobro con cheques
    totalCheques wsCorteCaja, wsReporte
    
    ' Calcula el total de Facturas al Contado
    totalFacturasContado wsCorteCaja, wsReporte
    
    ' Calcula el total de Facturas a Crédito
    totalFacturasCredito wsCorteCaja, wsReporte
    
    ' Calcula el total de Recibos de Caja en Quetzales y Dólares
    totalRecibosDeCaja wsCorteCaja, wsReporte
    
    ' Calcula el total de Facturas Anuladas Credito del día
    totalFacturasAnuladasCreditoDelDia wsCorteCaja, wsReporte
    
    ' Calcula el total de cobro con tarjetas
    totalTarjetas wsCorteCaja, wsReporte
    
    ' Calcula el total de Facturas Anuladas de otros días
    totalFacturasAnuladasOtrosDias wsCorteCaja, wsReporte
    
    ' Calcula el total de Facturas Anuladas Contado del día
    totalFacturasAnuladasContadoDelDia wsCorteCaja, wsReporte
    
    ' Calcula el total de cobro con tarjetas
    totalDepositos wsCorteCaja, wsReporte
    
    ' Calcula el total de Excenciones
    totalExcenciones wsCorteCaja, wsReporte
    
    ' Comprueba si el cajero cobró con Dólares
    cobroDolares wsCorteCaja, wsReporte
    
    ' Cierra el reporte sin guardar cambios
    wbReporte.Close SaveChanges:=False
    
    ' Notifica el final del procedimiento
    MsgBox "¡Se han copiado los datos del Reporte de SAP!" & vbCrLf & "" & vbCrLf & "Aunque se te informa si hay algún descuadre, por favor asegúrate de revisarlo antes de enviarlo"
    
    ' Notifica si hay diferencia en cheques
    notificarDiferenciaCheques wsCorteCaja, wsReporte
    
    ' Notifica si hay diferencia en tarjetas
    notificarDiferenciaTarjetas wsCorteCaja, wsReporte
    
    ' Notifica si hay diferencia en depositos
    notificarDiferenciaDepositos wsCorteCaja, wsReporte
    
    ' Notifica si el cajero cobró con Dólares
    notificarCobroDolares wsCorteCaja, wsReporte

    ' Notifica si el cajero cobró un Recibo de caja en Quetzales con Dólares
    notificarCobroDolaresEnRecibosQuetzales wsCorteCaja, wsReporte
    
    End If
End Sub


Sub totalEfectivo(wsCorteCaja As Worksheet, wsReporte As Worksheet)
    Dim celdaBusquedaFacturas As Range
    Dim celdaBusquedaRecibos As Range
    Dim celdaTotalesFacturas As Range
    Dim celdaTotalesRecibos As Range
    Dim totalEfectivo As Double

    ' Busca la celda "Facturas de Contado - Quetzales" en el reporte
    Set celdaBusquedaFacturas = wsReporte.Cells.Find("Facturas de Contado - Quetzales")
    Set celdaBusquedaRecibos = wsReporte.Cells.Find("Recibos de Caja - STOD - Quetzales")
    
    If Not celdaBusquedaFacturas Is Nothing Or Not celdaBusquedaRecibos Is Nothing Then
        ' Se desplaza hacia abajo en la columna hasta encontrar los totales
        Set celdaTotalesFacturas = wsReporte.Range(celdaBusquedaFacturas.Offset(1, 0), wsReporte.Cells(wsReporte.Rows.Count, celdaBusquedaFacturas.Column).End(xlUp)).Find("Totales")
        Set celdaTotalesRecibos = wsReporte.Range(celdaBusquedaRecibos.Offset(1, 0), wsReporte.Cells(wsReporte.Rows.Count, celdaBusquedaRecibos.Column).End(xlUp)).Find("Totales")
       
        ' Verifica si se encontraron las celdas "Totales"
        If Not celdaTotalesFacturas Is Nothing Or Not celdaTotalesRecibos Is Nothing Then

            ' Obtiene el total de Efectivo en Facturas y Recibos
            totalEfectivo = CDbl(celdaTotalesFacturas.Offset(0, 2).Value) + CDbl(celdaTotalesRecibos.Offset(0, 2).Value)
            
            ' Multiplica el total por 100 para trabajar con centavos
            totalEfectivo = totalEfectivo * 100

            ' Define la cantidad de billetes de cada denominación
            Dim billetes200 As Integer
            Dim billetes100 As Integer
            Dim billetes50 As Integer
            Dim billetes20 As Integer
            Dim billetes10 As Integer
            Dim billetes5 As Integer
            Dim billetes1 As Integer
            Dim monedas50c As Integer
            Dim monedas25c As Integer
            Dim monedas10c As Integer
            Dim monedas5c As Integer
            Dim monedas1c As Integer

            ' Distribuye el Efectivo por denominacion sin redondear el total
            billetes200 = totalEfectivo \ (200 * 100)
            totalEfectivo = totalEfectivo Mod (200 * 100)

            billetes100 = totalEfectivo \ (100 * 100)
            totalEfectivo = totalEfectivo Mod (100 * 100)

            billetes50 = totalEfectivo \ (50 * 100)
            totalEfectivo = totalEfectivo Mod (50 * 100)

            billetes20 = totalEfectivo \ (20 * 100)
            totalEfectivo = totalEfectivo Mod (20 * 100)

            billetes10 = totalEfectivo \ (10 * 100)
            totalEfectivo = totalEfectivo Mod (10 * 100)

            billetes5 = totalEfectivo \ (5 * 100)
            totalEfectivo = totalEfectivo Mod (5 * 100)

            billetes1 = totalEfectivo \ (1 * 100)
            totalEfectivo = totalEfectivo Mod (1 * 100)

            monedas50c = totalEfectivo \ 50
            totalEfectivo = totalEfectivo Mod 50

            monedas25c = totalEfectivo \ 25
            totalEfectivo = totalEfectivo Mod 25

            monedas10c = totalEfectivo \ 10
            totalEfectivo = totalEfectivo Mod 10

            monedas5c = totalEfectivo \ 5
            totalEfectivo = totalEfectivo Mod 5

            monedas1c = totalEfectivo

            ' Divide el resultado entre 100 para obtener el monto con centavos
            totalEfectivo = totalEfectivo / 100

            ' Copia los valores en las celdas correspondientes
            wsCorteCaja.Range("B14").Value = IIf(billetes200 = 0, "", billetes200)
            wsCorteCaja.Range("B15").Value = IIf(billetes100 = 0, "", billetes100)
            wsCorteCaja.Range("B16").Value = IIf(billetes50 = 0, "", billetes50)
            wsCorteCaja.Range("B17").Value = IIf(billetes20 = 0, "", billetes20)
            wsCorteCaja.Range("B18").Value = IIf(billetes10 = 0, "", billetes10)
            wsCorteCaja.Range("B19").Value = IIf(billetes5 = 0, "", billetes5)
            wsCorteCaja.Range("B22").Value = IIf(billetes1 = 0, "", billetes1)
            wsCorteCaja.Range("B23").Value = IIf(monedas50c = 0, "", monedas50c)
            wsCorteCaja.Range("B24").Value = IIf(monedas25c = 0, "", monedas25c)
            wsCorteCaja.Range("B25").Value = IIf(monedas10c = 0, "", monedas10c)
            wsCorteCaja.Range("B26").Value = IIf(monedas5c = 0, "", monedas5c)
            wsCorteCaja.Range("B27").Value = IIf(monedas1c = 0, "", monedas1c)
        End If
    End If
End Sub

Sub totalCheques(wsCorteCaja As Worksheet, wsReporte As Worksheet)
    Dim celdaBusquedaFacturas As Range
    Dim celdaBusquedaRecibos As Range
    Dim celdaTotalesFacturas As Range
    Dim celdaTotalesRecibos As Range
    Dim celdaCantidadPropios As Range
    Dim celdaCantidadTerceros As Range
    Dim celdaCantidadRecibos As Range
    Dim totalCheques As Double
    Dim totalChequesEnCorte As Double
    diferenciaEnCheques = False

    
    ' Busca la celda "Facturas Contado -STOD - Quetzales" en el reporte
    Set celdaBusquedaFacturas = wsReporte.Cells.Find("Facturas de Contado - Quetzales")

    ' Busca la celda "Recibos de Caja -STOD - Dólares" en el reporte
    Set celdaBusquedaRecibos = wsReporte.Cells.Find("Recibos de Caja - STOD - Quetzales")

    ' Verifica si se encontraron las celdas
    If Not celdaBusquedaFacturas Is Nothing Or Not celdaBusquedaRecibos Is Nothing Then

        ' Se desplaza hacia abajo en la columna hasta encontrar la celda "Totales"
        Set celdaTotalesFacturas = wsReporte.Range(celdaBusquedaFacturas.Offset(1, 0), wsReporte.Cells(wsReporte.Rows.Count, celdaBusquedaFacturas.Column).End(xlUp)).Find("Totales")
        Set celdaTotalesRecibos = wsReporte.Range(celdaBusquedaRecibos.Offset(1, 0), wsReporte.Cells(wsReporte.Rows.Count, celdaBusquedaRecibos.Column).End(xlUp)).Find("Totales")
        
        ' Verifica si se encontró la celda "Totales"
        If Not celdaTotalesFacturas Is Nothing Or Not celdaTotalesRecibos Is Nothing Then

            ' Se desplaza hacia la derecha para obtener los totales de cheques
            Set celdaCantidadPropios = celdaTotalesFacturas.Offset(0, 4)
            Set celdaCantidadTerceros = celdaTotalesFacturas.Offset(0, 5)
            Set celdaCantidadRecibos = celdaTotalesRecibos.Offset(0, 4)
            
            ' Verifica si se encontro alguna cantidad en los totales
            If Not celdaCantidadPropios Is Nothing Or Not celdaCantidadTerceros Is Nothing Or Not celdaCantidadRecibos Is Nothing Then
                
                ' Obtiene el total de cheques propios y de terceros en facturas y recibos
                totalCheques = Val(celdaCantidadPropios.Value) + Val(celdaCantidadTerceros.Value) + Val(celdaCantidadRecibos.Value)
        
                ' Obtiene el total de cheques ingresados por el cajero al Corte de Caja
                 totalChequesEnCorte = wsCorteCaja.Range("K32")
                 
                ' Guarda el total de cheques en Reporte de Caja en una variable global
                totalChequesEnReporte = totalCheques
                
                If Not totalChequesEnCorte = totalCheques Then
                diferenciaEnCheques = True
                End If
           
            End If
        End If
    End If
End Sub

Sub totalFacturasContado(wsCorteCaja As Worksheet, wsReporte As Worksheet)
    Dim celdaBusqueda As Range
    Dim celdaBusquedaDolares As Range
    Dim celdaTotales As Range
    Dim celdaTotalesDolares As Range
    Dim celdaCantidad As Range
    Dim celdaCantidadDolares As Range
    Dim totalFacturasContado As Double
    
    ' Busca las facturas de Contado - Quetzales y Dolares en el reporte
    Set celdaBusqueda = wsReporte.Cells.Find("Facturas de Contado - Quetzales")
    Set celdaBusquedaDolares = wsReporte.Cells.Find("Facturas de Contado - Dólares")
    
    

    ' Verifica si se encontró la celda
    If Not celdaBusqueda Is Nothing Or Not celdaBusquedaDolares Is Nothing Then
        ' Se desplaza hacia abajo en la columna hasta encontrar la celda que contiene los "Totales"
        Set celdaTotales = wsReporte.Range(celdaBusqueda.Offset(1, 0), wsReporte.Cells(wsReporte.Rows.Count, celdaBusqueda.Column).End(xlUp)).Find("Totales")
        Set celdaTotalesDolares = wsReporte.Range(celdaBusquedaDolares.Offset(1, 0), wsReporte.Cells(wsReporte.Rows.Count, celdaBusquedaDolares.Column).End(xlUp)).Find("Totales")
        ' Verifica si se encontró la celda "Totales"
        If Not celdaTotales Is Nothing Or Not celdaTotalesDolares Is Nothing Then

            ' Se desplaza hacia la derecha para obtener los datos a un lado de la celda "Totales"
            Set celdaCantidad = celdaTotales.Offset(0, 1)
            Set celdaCantidadDolares = celdaTotalesDolares.Offset(0, 1)
            ' Verifica si hay alguna cantidad en los totales
            If Not celdaCantidad Is Nothing Or celdaCantidadDolares Then
                ' Obtiene el valor de la cantidad de facturas al contado
                totalFacturasContado = Val(celdaCantidad.Value) + Val(celdaCantidadDolares.Value)
                
                ' Llena la celda I36 del corte de caja con el total de facturas al contado
                wsCorteCaja.Range("I36").Value = totalFacturasContado
            End If
        End If
    End If
End Sub

Sub totalFacturasCredito(wsCorteCaja As Worksheet, wsReporte As Worksheet)
    Dim celdaBusqueda As Range
    Dim celdaTotales As Range
    Dim celdaCantidad As Range
    Dim totalFacturasCredito As Double
    
    ' Busca la celda "Facturas Crédito" en el reporte
    Set celdaBusqueda = wsReporte.Cells.Find("Facturas Crédito")
    
    ' Verifica si se encontró la celda
    If Not celdaBusqueda Is Nothing Then

        ' Se desplaza hacia abajo en la columna hasta encontrar la celda "Totales"
        Set celdaTotales = wsReporte.Range(celdaBusqueda.Offset(1, 0), wsReporte.Cells(wsReporte.Rows.Count, celdaBusqueda.Column).End(xlUp)).Find("Totales")
        
        ' Verifica si se encontró la celda "Totales"
        If Not celdaTotales Is Nothing Then

            ' Se desplaza hacia la derecha para obtener los totales de Facturas Crédito
            Set celdaCantidad = celdaTotales.Offset(0, 1)
            
            ' Verifica si se encontró alguna cantidad
            If Not celdaCantidad Is Nothing Then

                ' Obtiene el valor de la cantidad de total facturas a crédito
                totalFacturasCredito = Val(celdaCantidad.Value)
                
                ' Llena la celda I37 del corte de caja con el total de facturas a crédito
                wsCorteCaja.Range("I37").Value = totalFacturasCredito
            End If
        End If
    End If
End Sub


Sub totalRecibosDeCaja(wsCorteCaja As Worksheet, wsReporte As Worksheet)
    Dim celdaBusqueda As Range
    Dim celdaBusquedaDolares As Range
    Dim celdaTotales As Range
    Dim celdaTotalesDolares As Range
    Dim celdaCantidad As Range
    Dim celdaCantidadDolares As Range
    Dim celdaCantidadQuetzalesEnDolares As Range
    Dim valorQuetzalesEnDolares As Integer
    Dim totalRecibosDeCaja As Double
    cobroReciboQuetzalesEnDolares = False

    
    ' Busca la celda "Recibos de Caja -STOD - Quetzales" en el reporte
    Set celdaBusqueda = wsReporte.Cells.Find("Recibos de Caja - STOD - Quetzales")

    ' Busca la celda "Recibos de Caja -STOD - Dólares" en el reporte
    Set celdaBusquedaDolares = wsReporte.Cells.Find("Recibos de Caja - STOD - Dólares")

    ' Verifica si se encontraron las celdas
    If Not celdaBusqueda Is Nothing Or Not celdaBusquedaDolares Is Nothing Then

        ' Se desplaza hacia abajo en la columna hasta encontrar la celda "Totales"
        Set celdaTotales = wsReporte.Range(celdaBusqueda.Offset(1, 0), wsReporte.Cells(wsReporte.Rows.Count, celdaBusqueda.Column).End(xlUp)).Find("Totales")
        Set celdaTotalesDolares = wsReporte.Range(celdaBusquedaDolares.Offset(1, 0), wsReporte.Cells(wsReporte.Rows.Count, celdaBusquedaDolares.Column).End(xlUp)).Find("Totales")
        
        ' Verifica si se encontró la celda "Totales"
        If Not celdaTotales Is Nothing Or Not celdaTotalesDolares Is Nothing Then

            ' Se desplaza hacia la derecha para obtener los Totales
            Set celdaCantidad = celdaTotales.Offset(0, 1)
            Set celdaCantidadDolares = celdaTotalesDolares.Offset(0, 2)
            Set celdaCantidadQuetzalesEnDolares = celdaTotales.Offset(0, 3)
             
             
              
                    
            ' Verifica si se encontro alguna cantidad en los totales
            If Not celdaCantidad Is Nothing Or Not celdaCantidadDolares Is Nothing Then

                ' Obtiene el total de recibos de caja en Quetzales y Dólares
                totalRecibosDeCaja = Val(celdaCantidad.Value) + Val(celdaCantidadDolares.Value)
                valorQuetzalesEnDolares = Val(celdaCantidadQuetzalesEnDolares.Value)
        
                ' Llena la celda I37 del corte de caja con el total de Recibos de caja cobrados
                wsCorteCaja.Range("I38").Value = totalRecibosDeCaja
                
                If Not valorQuetzalesEnDolares = 0 Then
                
                 cobroReciboQuetzalesEnDolares = True
                
                End If
                
                   
            End If
            
        End If
       
    End If
End Sub

Sub totalFacturasAnuladasCreditoDelDia(wsCorteCaja As Worksheet, wsReporte As Worksheet)
    Dim celdaBusqueda As Range
    Dim celdaCredito As Range
    Dim celdaFecha As Range
    Dim celdaMonto As Range
    Dim totalFacturasAnuladasCreditoDelDia As Double
    Dim fechaEnReporte As String
    Dim fechaActual As Date
    

    ' Leer la fecha en el reporte
    fechaEnReporte = wsReporte.Range("O1").Value

    ' Dividir la cadena usando los dos puntos como separador
    Dim partes() As String
        partes = Split(fechaEnReporte, ":")
    
    If UBound(partes) >= 1 Then
        ' Quita el texto de la fecha
        fechaActual = CDate(Trim(partes(1)))
        
        Else
            ' Si el proceso da error, se usará la fecha actual
            fechaActual = Date
    End If
    
    
    ' Coloca a 0 el valor en el Corte de Caja
     wsCorteCaja.Range("I39").Value = 0
    
    ' Busca la celda "Notas de Crédito" en el reporte
    Set celdaBusqueda = wsReporte.Cells.Find("Notas de Crédito")
    
    ' Verifica si se encontró la celda "Notas de Crédito"
    If Not celdaBusqueda Is Nothing Then

        ' Busca la celda "Crédito" después de encontrar "Notas de Crédito"
        Set celdaCredito = wsReporte.Range(celdaBusqueda.Offset(1, 0), wsReporte.Cells(wsReporte.Rows.Count, celdaBusqueda.Column).End(xlDown)).Find("CREDITO")
        
        ' Verifica si se encontró la celda "Crédito"
        If Not celdaCredito Is Nothing Then

            ' Recorre las celdas desde la posición de "Crédito" hacia abajo
            Set celdaFecha = celdaCredito.Offset(1, 0)

            ' Se ejecuta buscando datos hasta llegar a la celda "Totales"
            Do While Not UCase(celdaFecha.Value) = "Totales"
                
                ' Avanza a las celdas que contienen las fechas
                Set celdaFecha = celdaFecha.Offset(0, 6)
                
                ' Verifica si la celda contiene una fecha válida y si es igual a la fecha actual
                If celdaFecha.Value = fechaActual Then

                    ' Obtiene el monto que corresponde a las facturas anuladas de fecha igual a la del reporte
                    Set celdaMonto = celdaFecha.Offset(0, -2)

                    ' Guarda el dato en el Corte Excel en cada iteración
                    wsCorteCaja.Range("I39").Value = wsCorteCaja.Range("I39").Value + celdaMonto.Value
                End If

                   ' Si no es la fecha actual, regresa para iniciar el procedimiento en la siguiente fila
                    Set celdaFecha = celdaFecha.Offset(0, -6).Offset(1, 0)
                    
                    ' Cuando se acaban los datos finaliza el loop
                    If celdaFecha.Value = "Totales" Then
                    Exit Do
                End If
            Loop
        End If
    End If
End Sub

Sub totalTarjetas(wsCorteCaja As Worksheet, wsReporte As Worksheet)
    Dim celdaBusquedaFacturas As Range
    Dim celdaBusquedaRecibos As Range
    Dim celdaTotalesFacturas As Range
    Dim celdaTotalesRecibos As Range
    Dim celdaCantidadFacturas As Range
    Dim celdaCantidadRecibos As Range
    Dim totalTarjetas As Double
    Dim totalTarjetasEnCorte As Double
    diferenciaEnTarjetas = False

    
    ' Busca la celda "Facturas Contado -STOD - Quetzales" en el reporte
    Set celdaBusquedaFacturas = wsReporte.Cells.Find("Facturas de Contado - Quetzales")

    ' Busca la celda "Recibos de Caja -STOD - Dólares" en el reporte
    Set celdaBusquedaRecibos = wsReporte.Cells.Find("Recibos de Caja - STOD - Quetzales")

    ' Verifica si se encontraron las celdas
    If Not celdaBusquedaFacturas Is Nothing Or Not celdaBusquedaRecibos Is Nothing Then

        ' Se desplaza hacia abajo en la columna hasta encontrar la celda "Totales"
        Set celdaTotalesFacturas = wsReporte.Range(celdaBusquedaFacturas.Offset(1, 0), wsReporte.Cells(wsReporte.Rows.Count, celdaBusquedaFacturas.Column).End(xlUp)).Find("Totales")
        Set celdaTotalesRecibos = wsReporte.Range(celdaBusquedaRecibos.Offset(1, 0), wsReporte.Cells(wsReporte.Rows.Count, celdaBusquedaRecibos.Column).End(xlUp)).Find("Totales")
        
        ' Verifica si se encontró la celda "Totales"
        If Not celdaTotalesFacturas Is Nothing Or Not celdaTotalesRecibos Is Nothing Then

            ' Se desplaza hacia la derecha para obtener los totales de tarjetas
            Set celdaCantidadFacturas = celdaTotalesFacturas.Offset(0, 7)
            Set celdaCantidadRecibos = celdaTotalesRecibos.Offset(0, 6)
            
            ' Verifica si se encontro alguna cantidad en los totales
            If Not celdaCantidadFacturas Is Nothing Or Not celdaCantidadRecibos Is Nothing Then
                
                ' Obtiene el total de tarjetas en facturas y recibos
                totalTarjetas = Val(celdaCantidadFacturas.Value) + Val(celdaCantidadRecibos.Value)
        
                ' Obtiene el total de tarjetas ingresados por el cajero al Corte de Caja
                 totalTarjetasEnCorte = wsCorteCaja.Range("I45")
                 
                ' Guarda el total de tarjetas en Reporte de Caja en una variable global
                totalTarjetasEnReporte = totalTarjetas
                
                If Not totalTarjetasEnCorte = totalTarjetas Then
                diferenciaEnTarjetas = True
                End If
           
            End If
        End If
    End If
End Sub

Sub totalFacturasAnuladasOtrosDias(wsCorteCaja As Worksheet, wsReporte As Worksheet)
    Dim celdaBusqueda As Range
    Dim celdaNotasdeCredito As Range
    Dim celdaFecha As Range
    Dim celdaMonto As Range
    Dim totalFacturasAnuladasOtrosDias As Double
    Dim fechaEnReporte As String
    Dim fechaActual As Date
    

    ' Leer la fecha en el reporte
    fechaEnReporte = wsReporte.Range("O1").Value

    ' Dividir la cadena usando los dos puntos como separador
    Dim partes() As String
        partes = Split(fechaEnReporte, ":")
    
    If UBound(partes) >= 1 Then
        ' Quita el texto de la fecha
        fechaActual = CDate(Trim(partes(1)))
        
        Else
            ' Si el proceso da error, se usará la fecha actual
            fechaActual = Date
    End If
    
    ' Coloca a 0 el valor en el Corte de Caja
     wsCorteCaja.Range("I48").Value = 0
    
    ' Busca la celda "Notas de Crédito" en el reporte
    Set celdaBusqueda = wsReporte.Cells.Find("Notas de Crédito")
    
    ' Verifica si se encontró la celda "Notas de Crédito"
    If Not celdaBusqueda Is Nothing Then

        ' Establece la variable que se ubica en "Notas de Crédito"
        Set celdaNotasdeCredito = wsReporte.Cells.Find("Notas de Crédito")
        
        ' Verifica si existe la celda "Notas de Crédito"
        If Not celdaNotasdeCredito Is Nothing Then

            ' Recorre las celdas desde la posición de "Notas de Crédito" hacia abajo
            Set celdaFecha = celdaNotasdeCredito.Offset(1, 0)
            
             ' Se ejecuta iterando los datos hasta que llega a "Totales Notas de Credito"
            Do While Not UCase(celdaFecha.Value) = "Totales Notas Credito"

                ' Avanza hacia las celdas que contienen las fechas
                Set celdaFecha = celdaFecha.Offset(0, 6)
                
                ' Verifica si la celda contiene una fecha válida y si es diferente a la fecha actual
                If Not celdaFecha.Value = fechaActual Then
                    
                    ' Obtiene el monto correspondiente a las facturas anuladas de otros días
                    Set celdaMonto = celdaFecha.Offset(0, -2)
                
                    ' Guarda el dato en el Corte Excel en cada iteración
                    wsCorteCaja.Range("I48").Value = wsCorteCaja.Range("I48").Value + celdaMonto.Value

                End If

                   ' Regresa al inicio para repetir el bucle
                    Set celdaFecha = celdaFecha.Offset(0, -6).Offset(1, 0)
                    
                    ' Cuando se acaban los datos finaliza el loop
                    If celdaFecha.Value = "Totales Notas Credito" Then
                        Exit Do
                    End If
            Loop
        End If
    End If
End Sub

Sub totalFacturasAnuladasContadoDelDia(wsCorteCaja As Worksheet, wsReporte As Worksheet)
    Dim celdaBusqueda As Range
    Dim celdaContado As Range
    Dim celdaFecha As Range
    Dim celdaMonto As Range
    Dim totalFacturasAnuladasContadoDelDia As Double
    Dim fechaEnReporte As String
    Dim fechaActual As Date
    

    ' Leer la fecha en el reporte
    fechaEnReporte = wsReporte.Range("O1").Value

    ' Dividir la cadena usando los dos puntos como separador
    Dim partes() As String
        partes = Split(fechaEnReporte, ":")
    
    If UBound(partes) >= 1 Then
        ' Quita el texto de la fecha
        fechaActual = CDate(Trim(partes(1)))
        
        Else
            ' Si el proceso da error, se usará la fecha actual
            fechaActual = Date
    End If
    
    ' Coloca a 0 el valor en el Corte de Caja
     wsCorteCaja.Range("I49").Value = 0
    
    ' Busca la celda "Notas de Crédito" en el reporte
    Set celdaBusqueda = wsReporte.Cells.Find("Notas de Crédito")
    
    ' Verifica si se encontró la celda "Notas de Crédito"
    If Not celdaBusqueda Is Nothing Then
        ' Busca la celda "CONTADO" después de encontrar "Notas de Crédito"
        Set celdaContado = wsReporte.Range(celdaBusqueda.Offset(1, 0), wsReporte.Cells(wsReporte.Rows.Count, celdaBusqueda.Column).End(xlDown)).Find("CONTADO")
        
        ' Verifica si se encontró la celda "CONTADO"
        If Not celdaContado Is Nothing Then

            ' Recorre las celdas desde la posición de "CONTADO" hacia abajo
            Set celdaFecha = celdaContado.Offset(1, 0)
            
            ' Se ejecuta buscando datos hasta llegar a la celda "Totales"
            Do While Not UCase(celdaFecha.Value) = "Totales"
        
                ' Avanza a las celdas que contienen las fechas
                Set celdaFecha = celdaFecha.Offset(0, 6)
                
                ' Verifica si la celda contiene una fecha válida y si es igual a la fecha actual
                If celdaFecha.Value = fechaActual Then

                    ' Obtiene el monto de las facturas del dia anuladas
                    Set celdaMonto = celdaFecha.Offset(0, -2)
                
                    ' Guarda el dato en el Corte Excel en cada iteración
                    wsCorteCaja.Range("I49").Value = wsCorteCaja.Range("I49").Value + celdaMonto.Value

                End If

                   ' Si no es la fecha actual, continua en la siguiente fila
                    Set celdaFecha = celdaFecha.Offset(0, -6).Offset(1, 0)
                    
                    ' Cuando se acaban los datos finaliza el loop
                    If celdaFecha.Value = "Totales" Then
                        Exit Do
                    End If
            Loop
        End If
    End If
End Sub

Sub totalDepositos(wsCorteCaja As Worksheet, wsReporte As Worksheet)
    Dim celdaBusquedaFacturas As Range
    Dim celdaBusquedaRecibos As Range
    Dim celdaTotalesFacturas As Range
    Dim celdaTotalesRecibos As Range
    Dim celdaCantidadFacturas As Range
    Dim celdaCantidadRecibos As Range
    Dim totalDepositos As Double
    Dim totalDepositosEnCorte As Double
    diferenciaEnDepositos = False

    
    ' Busca la celda "Facturas Contado -STOD - Quetzales" en el reporte
    Set celdaBusquedaFacturas = wsReporte.Cells.Find("Facturas de Contado - Quetzales")

    ' Busca la celda "Recibos de Caja -STOD - Dólares" en el reporte
    Set celdaBusquedaRecibos = wsReporte.Cells.Find("Recibos de Caja - STOD - Quetzales")

    ' Verifica si se encontraron las celdas
    If Not celdaBusquedaFacturas Is Nothing Or Not celdaBusquedaRecibos Is Nothing Then

        ' Se desplaza hacia abajo en la columna hasta encontrar la celda "Totales"
        Set celdaTotalesFacturas = wsReporte.Range(celdaBusquedaFacturas.Offset(1, 0), wsReporte.Cells(wsReporte.Rows.Count, celdaBusquedaFacturas.Column).End(xlUp)).Find("Totales")
        Set celdaTotalesRecibos = wsReporte.Range(celdaBusquedaRecibos.Offset(1, 0), wsReporte.Cells(wsReporte.Rows.Count, celdaBusquedaRecibos.Column).End(xlUp)).Find("Totales")
        
        ' Verifica si se encontró la celda "Totales"
        If Not celdaTotalesFacturas Is Nothing Or Not celdaTotalesRecibos Is Nothing Then

            ' Se desplaza hacia la derecha para obtener los totales de depositos
            Set celdaCantidadFacturas = celdaTotalesFacturas.Offset(0, 6)
            Set celdaCantidadRecibos = celdaTotalesRecibos.Offset(0, 5)
            
            ' Verifica si se encontro alguna cantidad en los totales
            If Not celdaCantidadFacturas Is Nothing Or Not celdaCantidadRecibos Is Nothing Then
                
                ' Obtiene el total de depositos en facturas y recibos
                totalDepositos = Val(celdaCantidadFacturas.Value) + Val(celdaCantidadRecibos.Value)
        
                ' Obtiene el total de depóstios ingresados por el cajero al Corte de Caja
                 totalDepositosEnCorte = wsCorteCaja.Range("I54")
                 
                ' Guarda el total de depósitos en Reporte de Caja en una variable global
                totalDepositosEnReporte = totalDepositos
                
                If Not totalDepositosEnCorte = totalDepositos Then
                diferenciaEnDepositos = True
                End If
           
            End If
        End If
    End If
End Sub

Sub totalExcenciones(wsCorteCaja As Worksheet, wsReporte As Worksheet)
    Dim celdaBusqueda As Range
    Dim celdaTotales As Range
    Dim celdaCantidad As Range
    Dim totalExcenciones As Double
    
    ' Busca la celda "Facturas Contado" en el reporte
    Set celdaBusqueda = wsReporte.Cells.Find("Facturas de Contado - Quetzales")
    
    ' Verifica si se encontró la celda
    If Not celdaBusqueda Is Nothing Then
        ' Se desplaza hacia abajo en la columna hasta encontrar la celda "Totales"
        Set celdaTotales = wsReporte.Range(celdaBusqueda.Offset(1, 0), wsReporte.Cells(wsReporte.Rows.Count, celdaBusqueda.Column).End(xlUp)).Find("Totales")
        
        ' Verifica si se encontró la celda "Totales"
        If Not celdaTotales Is Nothing Then
            ' Se desplaza hacia la derecha para obtener el total de excenciones
            Set celdaCantidad = celdaTotales.Offset(0, 8)
            
            ' Verifica si se encontró alguna cantidad
            If Not celdaCantidad Is Nothing Then
                ' Obtiene el valor de la cantidad de total excenciones
                totalExcenciones = Val(celdaCantidad.Value)
                
                ' Llena la celda I52 del corte de caja con el total de Excenciones
                wsCorteCaja.Range("I52").Value = totalExcenciones
            End If
        End If
    End If
End Sub



Sub cobroDolares(wsCorteCaja As Worksheet, wsReporte As Worksheet)
    Dim celdaBusquedaFacturas As Range
    Dim celdaBusquedaRecibos As Range
    Dim celdaFacturas As Range
    Dim celdaRecibos As Range
    Dim celdaCantidadFacturas As Range
    Dim celdaCantidadRecibos As Range
    Dim totalDolares As Double
    cobroEnDolares = False
    
    ' Busca la celda "Facturas de Contado -STOD - Quetzales" en el reporte
    Set celdaBusquedaFacturas = wsReporte.Cells.Find("Facturas de Contado - Quetzales")

    ' Busca la celda "Recibos de Caja -STOD - Dolares" en el reporte
    Set celdaBusquedaRecibos = wsReporte.Cells.Find("Recibos de Caja - STOD - Dólares")

    ' Verifica si se encontraron las celdas
    If Not celdaBusquedaFacturas Is Nothing Or Not celdaBusquedaRecibos Is Nothing Then

        ' Se desplaza hacia abajo en la columna hasta encontrar la celda "Totales"
        Set celdaFacturas = wsReporte.Range(celdaBusquedaFacturas.Offset(1, 0), wsReporte.Cells(wsReporte.Rows.Count, celdaBusquedaFacturas.Column).End(xlUp)).Find("Totales")
        Set celdaRecibos = wsReporte.Range(celdaBusquedaRecibos.Offset(1, 0), wsReporte.Cells(wsReporte.Rows.Count, celdaBusquedaRecibos.Column).End(xlUp)).Find("Totales")
        
        ' Verifica si se encontraron las celdas "Totales"
        If Not celdaFacturas Is Nothing Or Not celdaRecibos Is Nothing Then

            ' Se desplaza hacia la derecha para obtener los totales de Dólares
            Set celdaCantidadFacturas = celdaFacturas.Offset(0, 3)
            Set celdaCantidadRecibos = celdaRecibos.Offset(0, 1)

            ' Verifica si se encontró alguna cantidad
            If Not celdaCantidadFacturas Is Nothing Or Not celdaCantidadRecibos Is Nothing Then
            
                ' Suma las cantidades de Dólares en Facturas Contado y Recibos
                totalDolares = Val(celdaCantidadFacturas.Value) + Val(celdaCantidadRecibos.Value)
            
                ' Si se cobró en dólares se actualiza una variable global
                ' Para notificar al cajero cuando se termine de importar los datos del reporte
                If Not totalDolares = 0 Then
                    cobroEnDolares = True
                End If
            End If
        End If
    End If
    
End Sub

Sub notificarDiferenciaCheques(wsCorteCaja As Worksheet, wsReporte As Worksheet)
 
    ' Obtiene los datos ingresados en el corte
    totalChequesEnCorte = wsCorteCaja.Range("K32").Value
    
    ' Calcula la diferencia entre los datos del reporte y los datos del corte
    Dim diferencia As Double
    diferencia = totalChequesEnCorte - totalChequesEnReporte

    
    If diferenciaEnCheques = True Then
    
        ' Notifica la diferencia al cajero
        MsgBox "¡ATENCIÓN, DESCUADRE EN CHEQUES!" & vbCrLf & "" & vbCrLf & "" & vbCrLf & "Total cheques en Reporte SAP:   Q" & Format(totalChequesEnReporte, "0.00") & vbCrLf & "Total cheques en Corte de Caja: Q" & Format(totalChequesEnCorte, "0.00") & vbCrLf & "" & vbCrLf & "La diferencia es: Q" & Format(diferencia, "0.00")
  End If
End Sub


Sub notificarDiferenciaTarjetas(wsCorteCaja As Worksheet, wsReporte As Worksheet)
 
    ' Obtiene los datos ingresados en el corte
    totalTarjetasEnCorte = wsCorteCaja.Range("I45").Value
    
    ' Calcula la diferencia entre los datos del reporte y los datos del corte
    Dim diferencia As Double
    diferencia = totalTarjetasEnCorte - totalTarjetasEnReporte

    
    If diferenciaEnTarjetas = True Then
    
        ' Notifica la diferencia al cajero
        MsgBox "¡ATENCIÓN, DESCUADRE EN TARJETAS!" & vbCrLf & "" & vbCrLf & "" & vbCrLf & "Total tarjetas en Reporte SAP:   Q" & Format(totalTarjetasEnReporte, "0.00") & vbCrLf & "Total tarjetas en Corte de Caja: Q" & Format(totalTarjetasEnCorte, "0.00") & vbCrLf & "" & vbCrLf & "La diferencia es: Q" & Format(diferencia, "0.00")
  End If
End Sub
  
Sub notificarDiferenciaDepositos(wsCorteCaja As Worksheet, wsReporte As Worksheet)
 
    ' Obtiene los datos ingresados en el corte
    totalDepositosEnCorte = wsCorteCaja.Range("I54").Value
    
    ' Calcula la diferencia entre los datos del reporte y los datos del corte
    Dim diferencia As Double
    diferencia = totalDepositosEnCorte - totalDepositosEnReporte

    
    If diferenciaEnDepositos = True Then
    
        ' Notifica la diferencia al cajero
        MsgBox "¡ATENCIÓN, DESCUADRE EN DEPÓSITOS!" & vbCrLf & "" & vbCrLf & "" & vbCrLf & "Total depósitos en Reporte SAP:   Q" & Format(totalDepositosEnReporte, "0.00") & vbCrLf & "Total depósitos en Corte de Caja: Q" & Format(totalDepositosEnCorte, "0.00") & vbCrLf & "" & vbCrLf & "La diferencia es: Q" & Format(diferencia, "0.00")
  End If
End Sub
  
Sub notificarCobroDolares(wsCorteCaja As Worksheet, wsReporte As Worksheet)
    ' Si se cobró en dólares se notifica al cajero
  
    If cobroEnDolares Then
        MsgBox "¡ATENCIÓN!" & vbCrLf & "" & vbCrLf & "Recuerda que facturaste en dólares, actualiza el tipo de cambio e ingrésalos en el cuadro inferior izquierdo según su forma:" & vbCrLf & "" & vbCrLf & "- Efectivo " & vbCrLf & "- Cheque" & vbCrLf & "- Depósito"
    End If
End Sub

Sub notificarCobroDolaresEnRecibosQuetzales(wsCorteCaja As Worksheet, wsReporte As Worksheet)
    ' Si se cobró en Recibo de Caja - Quezales en dólares se notifica al cajero
    
    
    If cobroReciboQuetzalesEnDolares Then
        MsgBox "¡ATENCIÓN!" & vbCrLf & "" & vbCrLf & "Cobraste un Recibo de Caja en Quetzales con moneda de Dólares, esto es incorrecto." & vbCrLf & "" & vbCrLf & "Anula el Recibo en Quetzales y hazlo en Recibo en Dólares "
    End If
End Sub

