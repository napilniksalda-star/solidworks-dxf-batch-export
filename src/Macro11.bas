Attribute VB_Name = "Macro11"
Option Explicit

' ====== НАСТРОЙКИ (поменять при других именах размеров) ======
Private Const LENGTH_DIM As String = "D1@Эскиз1"
Private Const WIDTH_DIM As String = "D2@Эскиз1"
Private Const QTY_PROP As String = "Quantity"
Private Const FILE_PREFIX As String = "Plate_"
' ============================================================

Dim swApp As SldWorks.SldWorks
Dim swModel As ModelDoc2
Dim swPart As PartDoc

Sub main()

    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then
        MsgBox "Открой деталь!"
        Exit Sub
    End If
    If swModel.GetType <> swDocPART Then
        MsgBox "Это не деталь!"
        Exit Sub
    End If
    Set swPart = swModel

    Dim modelPath As String
    modelPath = swModel.GetPathName
    If modelPath = "" Then
        MsgBox "Сначала сохрани файл!"
        Exit Sub
    End If

    ' Папка для DXF
    Dim outFolder As String
    outFolder = GetFolder(Left(modelPath, InStrRev(modelPath, "\")))
    If outFolder = "" Then
        MsgBox "Отменено."
        Exit Sub
    End If
    If Right(outFolder, 1) <> "\" Then outFolder = outFolder & "\"

    ' Запомнить исходную конфигурацию
    Dim origConf As String
    origConf = swModel.ConfigurationManager.ActiveConfiguration.Name

    Dim vConfNames As Variant
    vConfNames = swModel.GetConfigurationNames

    Dim okCount As Long, failCount As Long, skipCount As Long
    Dim report As String
    report = ""

    Dim i As Integer
    For i = 0 To UBound(vConfNames)
        Dim confName As String
        confName = vConfNames(i)

        swModel.ShowConfiguration2 confName
        swModel.EditRebuild3

        ' Толщина (из листового металла)
        Dim okT As Boolean, tMM As Double
        tMM = ReadThicknessMM(swModel, okT)
        If Not okT Then
            skipCount = skipCount + 1
            report = report & "ПРОПУСК (нет листового металла): " & confName & vbCrLf
            GoTo ContinueLoop
        End If

        ' Длина и ширина (из размеров)
        Dim okL As Boolean, okW As Boolean, lMM As Double, wMM As Double
        lMM = ReadDimMM(swModel, LENGTH_DIM, okL)
        wMM = ReadDimMM(swModel, WIDTH_DIM, okW)
        If (Not okL) Or (Not okW) Then
            failCount = failCount + 1
            report = report & "ОШИБКА (нет размера " & LENGTH_DIM & "/" & WIDTH_DIM & "): " & confName & vbCrLf
            GoTo ContinueLoop
        End If

        ' Количество (из свойства)
        Dim qty As String
        qty = ReadConfigProp(swModel, confName, QTY_PROP)
        If Trim(qty) = "" Then qty = "NA"

        ' Имя файла + защита от перезаписи
        Dim fullPath As String
        fullPath = UniqueFilePath(outFolder, BuildBaseName(lMM, wMM, tMM, qty), confName)

        ' Экспорт развёртки
        Dim status As Boolean
        status = swPart.ExportToDWG2(fullPath, modelPath, _
            swExportToDWG_ExportSheetMetal, True, Nothing, False, False, 0, Nothing)

        If status Then
            okCount = okCount + 1
            Debug.Print "OK: " & fullPath
        Else
            failCount = failCount + 1
            report = report & "ОШИБКА экспорта: " & confName & vbCrLf
        End If

ContinueLoop:
    Next i

    ' Вернуть исходную конфигурацию
    swModel.ShowConfiguration2 origConf
    swModel.EditRebuild3

    ' Отчёт
    Dim msg As String
    msg = "Готово!" & vbCrLf & _
          "Выгружено: " & okCount & vbCrLf & _
          "Пропущено: " & skipCount & vbCrLf & _
          "Ошибок: " & failCount
    If Len(report) > 0 Then msg = msg & vbCrLf & vbCrLf & "Подробности:" & vbCrLf & report
    MsgBox msg

End Sub


' ===================== ПОМОЩНИКИ =====================

' Число (мм) -> строка: точка-разделитель, без хвостовых нулей и без хвостовой точки.
' Str$ не зависит от локали (всегда ".") и не ставит лишний разделитель.
Private Function FormatNum(ByVal mm As Double) As String
    Dim s As String
    s = Trim$(Str$(Round(mm, 2)))
    If InStr(s, ".") > 0 Then
        Do While Right$(s, 1) = "0"
            s = Left$(s, Len(s) - 1)
        Loop
        If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)
    End If
    FormatNum = s
End Function

' Убрать символы, недопустимые в имени файла
Private Function SanitizeName(ByVal s As String) As String
    Dim bad As Variant, i As Integer
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, bad(i), "_")
    Next i
    SanitizeName = s
End Function

' Собрать имя файла: Plate_{Д}x{Ш}x{Т}_Qty_{К}.dxf
Private Function BuildBaseName(ByVal lMM As Double, ByVal wMM As Double, _
                              ByVal tMM As Double, ByVal qty As String) As String
    BuildBaseName = FILE_PREFIX & FormatNum(lMM) & "x" & FormatNum(wMM) & "x" & _
                    FormatNum(tMM) & "_Qty_" & SanitizeName(qty) & ".dxf"
End Function

' Уникальный путь: при совпадении добавляет имя конфигурации, затем счётчик
Private Function UniqueFilePath(ByVal folder As String, ByVal baseName As String, _
                                ByVal confName As String) As String
    Const ext As String = ".dxf"
    Dim p As String
    p = folder & baseName
    If Dir(p) = "" Then
        UniqueFilePath = p
        Exit Function
    End If

    Dim stem As String
    stem = Left(baseName, Len(baseName) - Len(ext))

    Dim cand As String
    cand = folder & stem & "_" & SanitizeName(confName) & ext
    If Dir(cand) = "" Then
        UniqueFilePath = cand
        Exit Function
    End If

    Dim n As Integer
    n = 2
    Do
        cand = folder & stem & "_" & SanitizeName(confName) & "_" & n & ext
        If Dir(cand) = "" Then
            UniqueFilePath = cand
            Exit Function
        End If
        n = n + 1
    Loop
End Function

' Диалог выбора папки (стартует в startPath); "" при отмене
Private Function GetFolder(ByVal startPath As String) As String
    Dim shell As Object, folder As Object
    Set shell = CreateObject("Shell.Application")
    ' &H1 = BIF_RETURNONLYFSDIRS (только папки файловой системы)
    Set folder = shell.BrowseForFolder(0&, "Выберите папку для сохранения DXF", &H1, startPath)
    If folder Is Nothing Then
        GetFolder = ""
    Else
        GetFolder = folder.Self.Path
    End If
End Function

' Значение размера активной конфигурации в мм; ok=False если размер не найден
Private Function ReadDimMM(swModel As ModelDoc2, ByVal dimName As String, _
                           ByRef ok As Boolean) As Double
    Dim swDim As Object
    Set swDim = swModel.Parameter(dimName)
    If swDim Is Nothing Then
        ok = False
        ReadDimMM = 0
    Else
        ok = True
        ReadDimMM = swDim.SystemValue * 1000#   ' метры -> мм
    End If
End Function

' Толщина листового металла активной конфигурации в мм; ok=False если нет Sheet-Metal
Private Function ReadThicknessMM(swModel As ModelDoc2, ByRef ok As Boolean) As Double
    Dim swFeat As Feature
    Set swFeat = swModel.FirstFeature
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "SheetMetal" Then
            Dim swSM As Object
            Set swSM = swFeat.GetDefinition
            If Not swSM Is Nothing Then
                ok = True
                ReadThicknessMM = swSM.Thickness * 1000#
                Exit Function
            End If
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
    ok = False
    ReadThicknessMM = 0
End Function

' Разрешённое значение свойства конфигурации (или "")
Private Function ReadConfigProp(swModel As ModelDoc2, ByVal confName As String, _
                                ByVal propName As String) As String
    Dim swConf As Object
    Set swConf = swModel.GetConfigurationByName(confName)
    If swConf Is Nothing Then
        ReadConfigProp = ""
        Exit Function
    End If
    Dim propMgr As Object
    Set propMgr = swConf.CustomPropertyManager
    Dim valOut As String, resolvedVal As String
    propMgr.Get4 propName, False, valOut, resolvedVal
    ReadConfigProp = resolvedVal
End Function
