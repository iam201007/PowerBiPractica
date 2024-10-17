Attribute VB_Name = "Módulo1"
Option Explicit
Option Base 1

Sub GenerarCSV()
    Dim fso As Object
    Dim fso1 As Object
    Dim ts As Object
    Dim ts1 As Object
    Dim i, k As Long
    Dim j As Integer
    Dim numFilas As Long
    Dim fechaInicio As Date
    Dim fechaFin As Date
    
    Dim cliente As Integer
    
    Dim fechaAleatoria As Date
    Dim numTicket As Long
    Dim articulos As Integer
    Dim codigos As Integer
    Dim cantidad As Integer
    Dim precio As Double
    
    Dim lista() As Long
    Dim indice As Long
    Dim sucursal As Integer

    ReDim lista(1)
    
    ' Configura la cantidad de filas, fechas de inicio y fin
    numFilas = 2000 ' Cambia este valor para ajustar la cantidad de filas
    fechaInicio = DateSerial(2022, 1, 1) ' Fecha de inicio de los datos
    fechaFin = DateSerial(2023, 12, 31) ' Fecha de fin de los datos
    indice = 0

    ' Crea un objeto para manejar archivos
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Crea un archivo CSV
    Set ts = fso.CreateTextFile("C:\borrar3\TicketsDetalle.csv", True) ' Cambia la ruta por la deseada

    ' Escribe el encabezado del CSV
    ts.writeline "Fecha|NúmeroTicket|Codigos|Cantidad|Precio"

    ' Genera las filas con datos aleatorios
    For i = 1 To numFilas
        ' Genera una fecha aleatoria entre las fechas especificadas
        fechaAleatoria = #1/1/1900# + Int((fechaFin - fechaInicio + 1) * Rnd) + fechaInicio
        ' Genera números aleatorios para ticket, cantidad y precio
        numTicket = Int((99999 - 10000 + 1) * Rnd + 10000)
        articulos = Int((10 - 1 + 1) * Rnd + 1)
        
        For j = 1 To articulos
            
            cantidad = Int((10 - 1 + 1) * Rnd + 1)
            codigos = Int((1000 - 1 + 1) * Rnd + 1)
                
            precio = Round(Rnd * 100, 2)
            
            ' Escribe la fila en el CSV
            ts.writeline fechaAleatoria & "|" & numTicket & "|" & codigo & "|" & cantidad & "|" & precio
        
            
            If indice > 0 Then
            
                If Not (VerificarArticuloEnLista(lista(), numTicket, indice)) Then
                
                    indice = indice + 1
                    ReDim Preserve lista(indice) As Long
                    lista(indice) = numTicket
                
                Else
                End If
                
            Else
                    
                    indice = indice + 1
                    ReDim Preserve lista(indice)
                    lista(indice) = numTicket
            
            
            End If
        
        Next j
        
    Next i

    ' Cierra el archivo
    ts.Close
    Set ts = Nothing
    Set fso = Nothing

    
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    
    Set ts1 = fso1.CreateTextFile("C:\borrar3\TicketsClientesSucursal.csv", True) ' Cambia la ruta por la deseada
    
    ts1.writeline "NúmeroTicket|Cliente|Sucursal"
        
    For k = 1 To indice
        
        cliente = Int((5 - 1 + 1) * Rnd + 1)
    
        sucursal = Int((10 - 1 + 1) * Rnd + 1)
        
        ts1.writeline lista(k) & "|" & cliente & "|" & sucursal
    
    Next k
    
    ts1.Close
    Set ts1 = Nothing
    
    Set fso1 = Nothing


    MsgBox "Archivo CSV generado correctamente."
    
End Sub

Private Function VerificarArticuloEnLista(lista() As Long, articulo, indice) As Boolean

Dim i As Long

    VerificarArticuloEnLista = False

    For i = 1 To indice
    
        If articulo = lista(i) Then
            VerificarArticuloEnLista = True
            Exit For
        Else
        End If
    
    Next i

End Function
