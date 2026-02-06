Attribute VB_Name = "Módulo2"
Option Explicit

Public Sub DepurarPorBlacklist()
    Dim rngData As Range, rngBlacklist As Range
    Dim wsData As Worksheet

    Dim dEmailsBL As Object, dDominiosBL As Object
    Dim dEliminadosUnicos As Object, dFreq As Object

    Dim i As Long, firstRow As Long, lastRow As Long, colData As Long
    Dim email As String, dominio As String, matchBL As Boolean

    '--- Pedir rangos
    Set rngData = Application.InputBox( _
        "Selecciona el rango de CORREOS a depurar (una sola columna, sin encabezado).", _
        "Rango a depurar", Type:=8)
    If rngData Is Nothing Then Exit Sub

    Set rngBlacklist = Application.InputBox( _
        "Selecciona el rango de BLACKLIST (correos completos y/o dominios con @dominio).", _
        "Rango Blacklist", Type:=8)
    If rngBlacklist Is Nothing Then Exit Sub

    Set wsData = rngData.Worksheet

    '--- Diccionarios
    Set dEmailsBL = CreateObject("Scripting.Dictionary"): dEmailsBL.CompareMode = vbTextCompare
    Set dDominiosBL = CreateObject("Scripting.Dictionary"): dDominiosBL.CompareMode = vbTextCompare
    Set dEliminadosUnicos = CreateObject("Scripting.Dictionary"): dEliminadosUnicos.CompareMode = vbTextCompare
    Set dFreq = CreateObject("Scripting.Dictionary"): dFreq.CompareMode = vbTextCompare

    '--- Cargar blacklist
    CargarBlacklist dEmailsBL, dDominiosBL, rngBlacklist

    '--- Preparar datos
    colData = rngData.Column
    firstRow = rngData.Row
    lastRow = rngData.Row + rngData.Rows.Count - 1

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    '--- Crear libro de salida
    Dim wbOut As Workbook
    Dim wsOut As Worksheet, wsResumen As Worksheet
    Set wbOut = Workbooks.Add(xlWBATWorksheet) ' 1 hoja
    Set wsOut = wbOut.Worksheets(1)
    wsOut.Name = "Depuracion"
    Set wsResumen = wbOut.Worksheets.Add(After:=wsOut)
    wsResumen.Name = "Resumen_Depuracion"

    '--- Encabezado en salida
    wsOut.Range("A1").Value = "EMAIL"
    wsOut.Rows(1).Font.Bold = True

    Dim outRow As Long
    outRow = 2

    Dim totalRegistros As Long, totalEliminados As Long
    totalRegistros = (lastRow - firstRow + 1)

    '--- Recorrer y copiar solo los que NO están en blacklist
    For i = firstRow To lastRow
        email = NormalizarEmail(wsData.Cells(i, colData).Value)
        If Len(email) = 0 Then GoTo Siguiente

        dominio = ExtraerDominio(email)

        matchBL = False
        If dEmailsBL.Exists(email) Then matchBL = True
        If dominio <> "" Then
            If dDominiosBL.Exists("@" & dominio) Then matchBL = True
        End If

        If matchBL Then
            totalEliminados = totalEliminados + 1

            If Not dEliminadosUnicos.Exists(email) Then dEliminadosUnicos.Add email, True

            If dFreq.Exists(email) Then
                dFreq(email) = CLng(dFreq(email)) + 1
            Else
                dFreq.Add email, 1
            End If
        Else
            wsOut.Cells(outRow, 1).Value = email
            outRow = outRow + 1
        End If

Siguiente:
    Next i

    '--- Escribir resumen
    EscribirResumen_Blacklist_NuevoLibro wsResumen, totalRegistros, totalEliminados, dEliminadosUnicos, dFreq

    wsOut.Columns("A:A").AutoFit
    wsResumen.Columns("A:E").AutoFit

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Listo. Se creó un nuevo libro con Depuracion y Resumen_Depuracion." & vbCrLf & _
           "Registros totales: " & totalRegistros & vbCrLf & _
           "Eliminados: " & totalEliminados, vbInformation
End Sub

'========================
' Blacklist loader: correos completos + dominios (@dominio)
'========================
Private Sub CargarBlacklist(ByVal dEmails As Object, ByVal dDominios As Object, ByVal rng As Range)
    Dim c As Range, k As String
    For Each c In rng.Cells
        k = NormalizarEmail(c.Value)
        If Len(k) > 0 Then
            If Left$(k, 1) = "@" Then
                If Not dDominios.Exists(k) Then dDominios.Add k, True
            Else
                If Not dEmails.Exists(k) Then dEmails.Add k, True
            End If
        End If
    Next c
End Sub

Private Function NormalizarEmail(ByVal v As Variant) As String
    Dim s As String
    s = Trim$(CStr(v))
    If Len(s) = 0 Then
        NormalizarEmail = ""
    Else
        NormalizarEmail = LCase$(s)
    End If
End Function

Private Function ExtraerDominio(ByVal email As String) As String
    Dim p As Long
    p = InStr(1, email, "@", vbTextCompare)
    If p = 0 Then
        ExtraerDominio = ""
    Else
        ExtraerDominio = Mid$(email, p + 1)
    End If
End Function

'========================
' Resumen en libro nuevo
'========================
Private Sub EscribirResumen_Blacklist_NuevoLibro(ByVal ws As Worksheet, _
                                                ByVal totalRegistros As Long, _
                                                ByVal totalEliminados As Long, _
                                                ByVal dEliminadosUnicos As Object, _
                                                ByVal dFreq As Object)

    Dim k As Variant, r As Long

    ws.Range("A1").Value = "Resumen de depuración"
    ws.Range("A3").Value = "Registros evaluados:"
    ws.Range("B3").Value = totalRegistros

    ws.Range("A4").Value = "Registros eliminados:"
    ws.Range("B4").Value = totalEliminados

    ws.Range("A5").Value = "Emails eliminados:"
    ws.Range("B5").Value = dEliminadosUnicos.Count

    ws.Range("A6").Value = "Registros repetidos eliminados (filas extra por duplicados):"
    ws.Range("B6").Value = totalEliminados - dEliminadosUnicos.Count

    ' Lista de eliminados únicos
    r = 8
    ws.Cells(r, 1).Value = "Correos eliminados (únicos)"
    ws.Cells(r + 1, 1).Value = "Email"

    Dim rr As Long
    rr = r + 2
    For Each k In dEliminadosUnicos.Keys
        ws.Cells(rr, 1).Value = k
        rr = rr + 1
    Next k

    ' TOP 20 repetidos
    ws.Range("D8").Value = "TOP repetidos (filas eliminadas por email)"
    ws.Range("D9").Value = "Email"
    ws.Range("E9").Value = "Repeticiones"

    If dFreq.Count > 0 Then
        Dim arrKeys As Variant, arrCounts As Variant
        Dim i As Long, j As Long, tmpK As Variant, tmpC As Variant, maxN As Long

        arrKeys = dFreq.Keys
        arrCounts = dFreq.Items

        ' Orden burbuja desc
        For i = LBound(arrCounts) To UBound(arrCounts) - 1
            For j = i + 1 To UBound(arrCounts)
                If CLng(arrCounts(j)) > CLng(arrCounts(i)) Then
                    tmpC = arrCounts(i): arrCounts(i) = arrCounts(j): arrCounts(j) = tmpC
                    tmpK = arrKeys(i): arrKeys(i) = arrKeys(j): arrKeys(j) = tmpK
                End If
            Next j
        Next i

        maxN = WorksheetFunction.Min(20, dFreq.Count)
        For i = 0 To maxN - 1
            ws.Cells(10 + i, 4).Value = arrKeys(i)
            ws.Cells(10 + i, 5).Value = arrCounts(i)
        Next i
    Else
        ws.Range("D10").Value = "(sin repetidos)"
    End If

    ws.Columns("A:E").AutoFit
End Sub


