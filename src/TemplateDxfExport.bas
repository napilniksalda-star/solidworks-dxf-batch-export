' TemplateDxfExport — экспорт развёрток в DXF по шаблону имени.
' Вставляй копипастом в модуль макроса SolidWorks (строка Attribute не нужна).
Option Explicit

' ====== НАСТРОЙКА ======
Private Const TEMPLATE_PROP As String = "Шаблон_имени"   ' свойство файла с шаблоном имени
' =======================

Dim swApp As SldWorks.SldWorks
Dim swModel As ModelDoc2
Dim swPart As PartDoc

Sub main()
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then MsgBox "Открой деталь!": Exit Sub
    If swModel.GetType <> swDocPART Then MsgBox "Это не деталь!": Exit Sub
    Set swPart = swModel

    Dim modelPath As String
    modelPath = swModel.GetPathName
    If modelPath = "" Then MsgBox "Сначала сохрани файл!": Exit Sub

    ' Шаблон имени из свойства файла
    Dim template As String
    template = ReadFileProp(swModel, TEMPLATE_PROP)
    If Trim(template) = "" Then
        MsgBox "Добавь в свойства файла (вкладка «Настройка») свойство '" & TEMPLATE_PROP & "'" & vbCrLf & _
               "с шаблоном имени, например:" & vbCrLf & _
               "уголок_{D1@Эскиз1}x{D2@Эскиз1}_{Длина}_Qty_{Количество}"
        Exit Sub
    End If

    ' Парность скобок
    If CountChar(template, "{") <> CountChar(template, "}") Then
        MsgBox "В шаблоне непарные скобки { }:" & vbCrLf & template
        Exit Sub
    End If

    ' Папка для DXF
    Dim outFolder As String
    outFolder = GetFolder(Left(modelPath, InStrRev(modelPath, "\")))
    If outFolder = "" Then MsgBox "Отменено.": Exit Sub
    If Right(outFolder, 1) <> "\" Then outFolder = outFolder & "\"

    ' Исходная конфигурация
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

        If Not HasSheetMetal(swModel) Then
            skipCount = skipCount + 1
            report = report & "ПРОПУСК (нет листового металла): " & confName & vbCrLf
            GoTo ContinueLoop
        End If

        Dim okName As Boolean, baseName As String
        baseName = BuildNameFromTemplate(swModel, confName, template, okName)
        If Not okName Then
            failCount = failCount + 1
            report = report & "ОШИБКА (шаблон/размер): " & confName & vbCrLf
            GoTo ContinueLoop
        End If

        Dim fullPath As String
        fullPath = UniqueFilePath(outFolder, SanitizeName(baseName) & ".dxf", confName)

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
    msg = "Готово!" & vbCrLf & "Выгружено: " & okCount & vbCrLf & _
          "Пропущено: " & skipCount & vbCrLf & "Ошибок: " & failCount
    If Len(report) > 0 Then msg = msg & vbCrLf & vbCrLf & "Подробности:" & vbCrLf & report
    MsgBox msg
End Sub


' ===== ПОДСТАНОВКА ШАБЛОНА =====

' Разрешить один плейсхолдер: с "@" -> размер (мм), иначе -> свойство (конфиг, затем файл)
Private Function ResolvePlaceholder(swModel As ModelDoc2, ByVal confName As String, _
                                    ByVal token As String, ByRef ok As Boolean) As String
    If InStr(token, "@") > 0 Then
        Dim swDim As Object
        Set swDim = swModel.Parameter(token)
        If swDim Is Nothing Then
            ok = False
            ResolvePlaceholder = ""
        Else
            ok = True
            ResolvePlaceholder = FormatNum(swDim.SystemValue * 1000#)
        End If
    Else
        ok = True
        Dim v As String
        v = ReadConfigProp(swModel, confName, token)
        If v = "" Then v = ReadFileProp(swModel, token)
        ResolvePlaceholder = v
    End If
End Function

' Собрать базовое имя из шаблона (без расширения). ok=False при непарной скобке или ненайденном размере
Private Function BuildNameFromTemplate(swModel As ModelDoc2, ByVal confName As String, _
                                       ByVal template As String, ByRef ok As Boolean) As String
    ok = True
    Dim result As String
    result = template
    Dim guard As Integer
    Do
        Dim p1 As Long, p2 As Long
        p1 = InStr(result, "{")
        If p1 = 0 Then Exit Do
        p2 = InStr(p1 + 1, result, "}")
        If p2 = 0 Then ok = False: BuildNameFromTemplate = "": Exit Function
        Dim token As String
        token = Mid(result, p1 + 1, p2 - p1 - 1)
        Dim okOne As Boolean, val As String
        val = ResolvePlaceholder(swModel, confName, token, okOne)
        If Not okOne Then ok = False: BuildNameFromTemplate = "": Exit Function
        result = Left(result, p1 - 1) & val & Mid(result, p2 + 1)
        guard = guard + 1
        If guard > 200 Then Exit Do
    Loop
    BuildNameFromTemplate = result
End Function


' ===== ЧТЕНИЕ СВОЙСТВ / ГЕОМЕТРИИ =====

' Значение свойства уровня ФАЙЛА
Private Function ReadFileProp(swModel As ModelDoc2, ByVal propName As String) As String
    Dim cpm As Object
    Set cpm = swModel.Extension.CustomPropertyManager("")
    Dim valOut As String, resolved As String
    cpm.Get4 propName, False, valOut, resolved
    ReadFileProp = resolved
End Function

' Значение свойства КОНФИГУРАЦИИ (или "")
Private Function ReadConfigProp(swModel As ModelDoc2, ByVal confName As String, _
                                ByVal propName As String) As String
    Dim swConf As Object
    Set swConf = swModel.GetConfigurationByName(confName)
    If swConf Is Nothing Then ReadConfigProp = "": Exit Function
    Dim cpm As Object
    Set cpm = swConf.CustomPropertyManager
    Dim valOut As String, resolved As String
    cpm.Get4 propName, False, valOut, resolved
    ReadConfigProp = resolved
End Function

' Есть ли в детали элемент листового металла
Private Function HasSheetMetal(swModel As ModelDoc2) As Boolean
    Dim swFeat As Feature
    Set swFeat = swModel.FirstFeature
    Do While Not swFeat Is Nothing
        If swFeat.GetTypeName2 = "SheetMetal" Then
            HasSheetMetal = True
            Exit Function
        End If
        Set swFeat = swFeat.GetNextFeature
    Loop
    HasSheetMetal = False
End Function


' ===== ОБЩИЕ ПОМОЩНИКИ (как в макросе пластин) =====

' Число (мм) -> строка: точка-разделитель, без хвостовых нулей и без хвостовой точки
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

' Убрать недопустимые в имени файла символы
Private Function SanitizeName(ByVal s As String) As String
    Dim bad As Variant, i As Integer
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, bad(i), "_")
    Next i
    SanitizeName = s
End Function

' Сколько раз символ ch встречается в s
Private Function CountChar(ByVal s As String, ByVal ch As String) As Long
    CountChar = Len(s) - Len(Replace(s, ch, ""))
End Function

' Уникальный путь: при совпадении добавляет имя конфигурации, затем счётчик
Private Function UniqueFilePath(ByVal folder As String, ByVal baseName As String, _
                                ByVal confName As String) As String
    Const ext As String = ".dxf"
    Dim p As String
    p = folder & baseName
    If Dir(p) = "" Then UniqueFilePath = p: Exit Function

    Dim stem As String
    stem = Left(baseName, Len(baseName) - Len(ext))

    Dim cand As String
    cand = folder & stem & "_" & SanitizeName(confName) & ext
    If Dir(cand) = "" Then UniqueFilePath = cand: Exit Function

    Dim n As Integer
    n = 2
    Do
        cand = folder & stem & "_" & SanitizeName(confName) & "_" & n & ext
        If Dir(cand) = "" Then UniqueFilePath = cand: Exit Function
        n = n + 1
    Loop
End Function

' Диалог выбора папки (стартует в startPath); "" при отмене
Private Function GetFolder(ByVal startPath As String) As String
    Dim shell As Object, folder As Object
    Set shell = CreateObject("Shell.Application")
    Set folder = shell.BrowseForFolder(0&, "Выберите папку для сохранения DXF", &H1, startPath)
    If folder Is Nothing Then
        GetFolder = ""
    Else
        GetFolder = folder.Self.Path
    End If
End Function
