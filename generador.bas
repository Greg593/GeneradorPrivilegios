Attribute VB_Name = "generador"
Sub generaPrivilegios()
       Const formula  As String = "=BUSCAR(ALEATORIO.ENTRE(MIN(Miembros!$A$3:$A$31),MAX(Miembros!$A$3:$A$31)),Miembros!$A$3:$A$31,Miembros!$B$3:$B$31)"
       Const Predicadores As String = "=BUSCAR(ALEATORIO.ENTRE(MIN(Miembros!$A$3:$A$31),MAX(Miembros!$A$3:$A$31)),Miembros!$A$3:$A$31,Miembros!$B$3:$B$31)"
       Const PastorGeneral As String = "=Miembros!$B$2"
       Const PastorJovenes As String = "=Miembros!$B$12"
       rPredicador = Range("Miembros!B2:D31")
       
       ActiveSheet.Protect DrawingObjects:=False, Contents:=False, _
          Scenarios:=False
       Application.ScreenUpdating = False
       On Error GoTo MyErrorTrap
       ' Limpio el documento
       Range("a1:h34").Clear
       ' Pido mes a trabajar
       fechaEntrada = InputBox("Ingrese Mes y A�o del Calendario (Ej: Enero 2018)")
       ' Termina el programa si no  ingresa informaci�n
       If fechaEntrada = "" Then Exit Sub
       ' Primer d�a del mes
       StartDay = DateValue(fechaEntrada)
       If Day(StartDay) <> 1 Then
           StartDay = DateValue(Month(StartDay) & "/1/" & _
               Year(StartDay))
       End If
       ' Prepara formato para colocar el t�tulo del mes + a�o
       Range("a1").NumberFormat = "mmmm yyyy"
       ' Formato para el T�tulo del Mes
       With Range("a1:e1")
           .HorizontalAlignment = xlCenterAcrossSelection
           .VerticalAlignment = xlCenter
           .Font.Size = 20
           .Font.Bold = True
           .RowHeight = 35
       End With
       ' Formato para los t�tulos
       With Range("a2:e2")
           .ColumnWidth = 15
           .VerticalAlignment = xlCenter
           .HorizontalAlignment = xlCenter
           .VerticalAlignment = xlCenter
           .Orientation = xlHorizontal
           .Font.Size = 12
           .Font.Bold = True
           .RowHeight = 25
       End With
       ' T�tulos
       Range("a2") = "D�a"
       Range("b2") = "Actividad"
       Range("c2") = "Direcci�n"
       ' Range("d2") = "Lectura"
       Range("d2") = "Ofrenda"
       Range("e2") = "Predica"

       ' t�tulo del Mes + a�o
       Range("a1").Value = Application.Text(fechaEntrada, "mmmm yyyy")
       ' Qu� d�a de la semana
       diaSemana = Weekday(StartDay)
       ' Identifico A�o y Mes por separado
       CurYear = Year(StartDay)
       CurMonth = Month(StartDay)
       ' Calculo el ultimo d�a del mes
       FinalDay = DateSerial(CurYear, CurMonth + 1, 1)
      
      ' Formato para d�a de servicios
       With Range("a3:a30")
           .HorizontalAlignment = xlLeft
           .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = True
           .RowHeight = 26
           .NumberFormat = "dddd dd"
       End With
       ' Se coloca el primer d�a en la celda
       Contador = 0
       Do While diaSemana <> 0
        If diaSemana = 1 Or diaSemana = 3 Or diaSema = 4 Or diaSemana = 5 Or diaSemana = 7 Then
            Range("a3").Value = StartDay + Contador
            Exit Do
        End If
        diaSemana = diaSemana + 1
        Contador = Contador + 1
       Loop

       ' Genera Fechas de Servicio
        For Each cell In Range("a3:a30")
          ' No hace nada si es la primera celda
          If cell.Column = 1 And cell.Row = 3 Then
          ' Si la celda actual no es la primera
            'Si es domingo se repite nuevamente la fecha
            seraDomingo = Weekday(cell.Value)
            seRepiteDia = 0
            If seraDomingo = 1 Then
              seRepiteDia = 1
            End If
          ElseIf cell.Row <> 1 And seRepiteDia = 0 Then
            diaAnterior = Weekday(cell.Offset(-1, 0).Value)
            If diaAnterior = 1 Or diaAnterior = 5 Then
              cell.Value = cell.Offset(-1, 0).Value + 2
            Else
              cell.Value = cell.Offset(-1, 0).Value + 1
            End If
            'Si es domingo se repite nuevamente la fecha
            seraDomingo = Weekday(cell.Value)
            seRepiteDia = 0
            If seraDomingo = 1 Then
              seRepiteDia = 1
            End If
          ElseIf cell.Row <> 1 And seRepiteDia = 1 Then
            cell.Value = cell.Offset(-1, 0).Value
            seRepiteDia = 0
          End If
          ' Se sale del ciclo si el d�a es mayor al del mes
          If cell.Value >= FinalDay Then
            cell.Value = ""
            Exit For
          End If
        Next

      ' Formato para Actividades
       With Range("b3:b30")
           .HorizontalAlignment = xlLeft
           .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
           .RowHeight = 21
           .ColumnWidth = 25
       End With
        'Genera Tipo de Actividades
        noDomingos = 0
        For Each cell In Range("b3:b30")
          If cell.Offset(0, -1).Value <> "" Then
            fecha = Weekday(cell.Offset(0, -1).Value)
            Select Case fecha
              Case 1
                If domingoMa�ana = 0 Then
                  cell.Value = "Escuela Dominical"
                  domingoMa�ana = 1
                Else
                  If noDomingos <> 1 Then
                    cell.Value = "Servicio Evangel�stico"
                  Else
                    cell.Value = "Santa Cena"
                  End If
                  domingoMa�ana = 0
                End If
                noDomingos = noDomingos + 1
              Case 3
                cell.Value = "Indagando las Escrituras"
              Case 4
                cell.Value = "C�lula"
              Case 5
                cell.Value = "Ense�anza B�blica"
              Case 7
                cell.Value = "Adoraci�n Juvenil"
            End Select
          Else
            Exit For
          End If
        Next

      ' Formato para Privilegios
      With Range("c3:e30")
           .HorizontalAlignment = xlLeft
           .VerticalAlignment = xlCenter
           .Font.Size = 12
           .Font.Bold = False
           .RowHeight = 21
           .ColumnWidth = 25
      End With
      'Genera Privilegios
      For Each cell In Range("c3:e30")

        'Para la fila 3
        If cell.Row = 3 Then ' Direcci�n
          If cell.Column = 3 Then
            cell.FormulaLocal = formula
          ElseIf cell.Column = 4 Then ' Ofrenda
            resultado = 0
            Do While resultado = 0
              cell.FormulaLocal = formula
              queTiene = cell.Value
              If queTiene <> "" Then
                If cell.Offset(0, -1).Value <> queTiene Then
                      resultado = 1
                      fecha = cell.Offset(0, -3).Value
                      celda = cell.Offset(0, 0)
                      Call validaEspecial(fecha, celda, resultado)
                End If
              End If
            Loop
          End If
        End If

        'Para la fila 4
        If cell.Row = 4 Then
          If cell.Column = 3 Then ' Direcci�n
            resultado = 0
            Do While resultado = 0
              cell.FormulaLocal = formula
              queTiene = cell.Value
              If queTiene <> "" Then
                If cell.Offset(-1, 0).Value <> queTiene And _
                   cell.Offset(-1, 1).Value <> queTiene Then
                      resultado = 1
                      fecha = cell.Offset(0, -2).Value
                      celda = cell.Offset(0, 0)
                      Call validaEspecial(fecha, celda, resultado)
                End If
              End If
            Loop
          ElseIf cell.Column = 4 Then ' Ofrenda
           resultado = 0
            Do While resultado = 0
              cell.FormulaLocal = formula
              queTiene = cell.Value
              If queTiene <> "" Then
                If cell.Offset(-1, -1).Value <> queTiene And _
                   cell.Offset(-1, 0).Value <> queTiene And _
                   cell.Offset(0, -1).Value <> queTiene Then
                      resultado = 1
                      fecha = cell.Offset(0, -3).Value
                      celda = cell.Offset(0, 0)
                      Call validaEspecial(fecha, celda, resultado)
                End If
              End If
            Loop
          End If
        End If

        'Para la fila >4
        If cell.Row > 4 Then
          If cell.Column = 3 Then ' Direcci�n
            resultado = 0
            Do While resultado = 0
              cell.FormulaLocal = formula
              queTiene = cell.Value
              If queTiene <> "" Then
                If cell.Offset(-2, 2).Value <> queTiene And _
                   cell.Offset(-2, 1).Value <> queTiene And _
                   cell.Offset(-2, 0).Value <> queTiene And _
                   cell.Offset(-1, 2).Value <> queTiene And _
                   cell.Offset(-1, 1).Value <> queTiene And _
                   cell.Offset(-1, 0).Value <> queTiene Then
                      resultado = 1
                      fecha = cell.Offset(0, -2).Value
                      celda = cell.Offset(0, 0)
                      Call validaEspecial(fecha, celda, resultado)
                End If
              End If
            Loop
          ElseIf cell.Column = 4 Then ' Ofrenda
           resultado = 0
            Do While resultado = 0
              cell.FormulaLocal = formula
              queTiene = cell.Value
              If queTiene <> "" Then
                If cell.Offset(-2, -1).Value <> queTiene And _
                   cell.Offset(-2, 0).Value <> queTiene And _
                   cell.Offset(-1, -1).Value <> queTiene And _
                   cell.Offset(-1, 0).Value <> queTiene And _
                   cell.Offset(0, -1).Value <> queTiene Then
                      resultado = 1
                      fecha = cell.Offset(0, -3).Value
                      celda = cell.Offset(0, 0)
                      Call validaEspecial(fecha, celda, resultado)
                End If
              End If
            Loop
          End If
        End If


        'Convierte F�rmula en Valor
        If cell.HasFormula Then
          cell.Value = cell.Value
        End If
        'Si no tiene fecha lo elimina
        If cell.Offset(0, -1).Value = "" Then
          cell.Value = ""
        End If
      Next
 

      'Genera Predicadores
      yaPaso = 0
      For Each cell In Range("e3:e30")
        If cell.Offset(0, -4).Value <> "" Then
          noDiafecha = Weekday(cell.Offset(0, -4).Value)
          fecha = cell.Offset(0, -4).Value
          Select Case noDiafecha
            Case 1
                noSemana = Format(fecha, "ww")
                residuo = noSemana Mod 2
                If residuo <> 0 Then
                    If yaPaso = 1 Then
                        cell.FormulaLocal = PastorGeneral
                        yaPaso = 0
                    Else
                        cell.FormulaLocal = PastorJovenes
                        yaPaso = 1
                    End If
                Else
                    If yaPaso = 1 Then
                        cell.FormulaLocal = PastorJovenes
                        yaPaso = 0
                    Else
                        cell.FormulaLocal = PastorGeneral
                        yaPaso = 1
                    End If
                End If
            Case 3
              cell.FormulaLocal = PastorGeneral
            Case 5
          End Select
        Else
          Exit For
        End If

        'Convierte F�rmula en Valor
        If cell.HasFormula Then
          cell.Value = cell.Value
        End If
        'Si no tiene fecha lo elimina
        If cell.Offset(0, -1).Value = "" Then
          cell.Value = ""
        End If
      Next
 

        'Otros atributos
       ActiveWindow.DisplayGridlines = False
       ActiveSheet.Protect DrawingObjects:=True, Contents:=True, _
          Scenarios:=True
       ActiveWindow.WindowState = xlMaximized
       ActiveWindow.ScrollRow = 1
       Application.ScreenUpdating = True
       Exit Sub
MyErrorTrap:
       MsgBox "NO ingresaste correctamente la fecha o hubo un Error." _
           & Chr(13) & "Digita correctamente el mes" _
           & " (o utiliza un nombre abreviado por 3 letras)" _
           & Chr(13) & "y 4 digitos para el a�o"
       fechaEntrada = InputBox("Ingrese Mes y A�o del Calendario (Ejemplo: Enero 2018)")
       If fechaEntrada = "" Then Exit Sub
       Resume
   End Sub

Sub validaEspecial(fecha, celda, resultado)
    rOfrenda = Range("Miembros!B2:E31")
    rRegular = Range("Miembros!B2:G31")
    rJoven = Range("Miembros!B2:H31")
    result = resultado
    
    nOfrenda = Application.WorksheetFunction.VLookup(celda, rOfrenda, 4, 0)
    If nOfrenda <> "S" Then
      result = 0
    End If
    
    noDiafecha = Weekday(fecha)
    Select Case noDiafecha
        Case 1
        Case 3
            nRegular = Application.WorksheetFunction.VLookup(celda, rRegular, 6, 0)
            If nRegular <> "S" Then
              result = 0
            End If
        Case 5
            nRegular = Application.WorksheetFunction.VLookup(celda, rRegular, 6, 0)
            If nRegular <> "S" Then
              result = 0
            End If
        Case 7
            nJoven = Application.WorksheetFunction.VLookup(celda, rJoven, 7, 0)
            If nJoven <> "S" Then
              result = 0
            End If
    End Select
    resultado = result
End Sub

