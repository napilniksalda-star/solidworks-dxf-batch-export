# Гибкий экспорт DXF по шаблону — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Новый отдельный SolidWorks VBA-макрос, который именует DXF-развёртки по шаблону, заданному в свойстве детали `Шаблон_имени` — работает для деталей любого типа.

**Architecture:** Один новый модуль `src/TemplateDxfExport.bas`. `Sub main()` — оркестратор; имя строится подстановкой `{…}` из шаблона: плейсхолдеры с `@` читаются как размеры геометрии, без `@` — как свойства. Общие помощники копируются из проверенного макроса пластин.

**Tech Stack:** SolidWorks API (sldworks.tlb, swconst.tlb), VBA, `Shell.Application` (поздняя привязка).

**Spec:** `docs/superpowers/specs/2026-06-06-template-dxf-export-design.md`

---

## Как «тестировать» VBA-макрос

Автоматического раннера нет. Разработку ведём в редакторе VBA нового макроса. **Чистые** функции (`FormatNum`, `SanitizeName`, `CountChar`) проверяем в окне Immediate (`?Func(...)`). Функции с API/шаблоном — временным `Sub` на открытой детали. Весь макрос — ручным протоколом (Задача 6). Текстовый исходник — `src/TemplateDxfExport.bas` (экспорт модуля из редактора).

> **Кодировка:** код содержит кириллицу — загружай его в редактор **копипастом**, не через `File → Import`.

---

## File Structure

| Файл | Ответственность |
|---|---|
| `src/TemplateDxfExport.bas` | Новый модуль макроса (создаётся в этом плане). |
| `README.md` | Дополняется разделом про второй макрос (Задача 7). |
| `src/Macro11.bas` | **Не трогаем** — макрос пластин. |

Внутри модуля — функции с одной ответственностью: `main` (оркестратор), `BuildNameFromTemplate` + `ResolvePlaceholder` (шаблон), `ReadFileProp`/`ReadConfigProp`/`HasSheetMetal` (чтение), и общие `FormatNum`/`SanitizeName`/`CountChar`/`UniqueFilePath`/`GetFolder`.

---

## Task 1: Каркас модуля + проверки входа

**Files:**
- Create: `src/TemplateDxfExport.bas`

- [ ] **Step 1: Создать файл с каркасом**

```vba
Attribute VB_Name = "TemplateDxfExport"
Option Explicit

' ====== НАСТРОЙКА ======
Private Const TEMPLATE_PROP As String = "Шаблон_имени"   ' свойство файла с шаблоном
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

    MsgBox "Проверки пройдены (заглушка)."
End Sub
```

- [ ] **Step 2: Загрузить в новый макрос и скомпилировать**

В SolidWorks: `Сервис → Макрос → Создать…`, сохрани (напр. `TemplateDxfExport.swp`). В редакторе VBA вставь код в модуль (копипастом). `Debug → Compile` → без ошибок.

- [ ] **Step 3: Проверить заглушку**

Без открытых документов `F5` → «Открой деталь!». С открытой сохранённой деталью → «Проверки пройдены (заглушка)».

- [ ] **Step 4: Commit**

Экспортируй модуль в `src/TemplateDxfExport.bas` (правый клик → `Export File…`), затем:
```bash
git add src/TemplateDxfExport.bas
git commit -m "feat: template macro skeleton with entry guards"
```

---

## Task 2: Общие помощники (из макроса пластин)

**Files:**
- Modify: `src/TemplateDxfExport.bas`

- [ ] **Step 1: Добавить функции в конец модуля**

```vba
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
```

- [ ] **Step 2: Скомпилировать**

`Debug → Compile` → без ошибок.

- [ ] **Step 3: Проверить чистые функции в Immediate (`Ctrl+G`)**

```
?FormatNum(100)            -> 100
?FormatNum(1.5)            -> 1.5
?SanitizeName("a/b:c")     -> a_b_c
?CountChar("a{b}{c}", "{") -> 2
?CountChar("a{b}{c}", "}") -> 2
```

- [ ] **Step 4: Commit**

```bash
git add src/TemplateDxfExport.bas
git commit -m "feat: shared helpers (FormatNum, SanitizeName, CountChar, UniqueFilePath, GetFolder)"
```

---

## Task 3: Чтение свойств/геометрии + чтение шаблона в main

**Files:**
- Modify: `src/TemplateDxfExport.bas`

- [ ] **Step 1: Добавить функции чтения**

```vba
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
```

- [ ] **Step 2: Заменить заглушку в `main` на чтение шаблона**

Замени строку `MsgBox "Проверки пройдены (заглушка)."` на:
```vba
    Dim template As String
    template = ReadFileProp(swModel, TEMPLATE_PROP)
    If Trim(template) = "" Then
        MsgBox "Добавь в свойства файла (вкладка «Настройка») свойство '" & TEMPLATE_PROP & "'" & vbCrLf & _
               "с шаблоном имени, например:" & vbCrLf & _
               "уголок_{D1@Эскиз1}x{D2@Эскиз1}_{Длина}_Qty_{Количество}"
        Exit Sub
    End If
    MsgBox "Шаблон: " & template   ' временно, для проверки
```

- [ ] **Step 3: Скомпилировать**

`Debug → Compile` → без ошибок.

- [ ] **Step 4: Проверить чтение шаблона на детали**

Открой деталь, добавь свойство файла `Шаблон_имени` = `тест_{Количество}` (`Файл → Свойства → Настройка`). `F5`.
Expected: окно «Шаблон: тест_{Количество}». Убери свойство → `F5` → подсказка добавить свойство.

- [ ] **Step 5: Commit**

```bash
git add src/TemplateDxfExport.bas
git commit -m "feat: property/geometry readers + template read in main"
```

---

## Task 4: Подстановка шаблона

**Files:**
- Modify: `src/TemplateDxfExport.bas`

- [ ] **Step 1: Ожидаемое поведение**

Для шаблона `уголок_{D1@Эскиз1}x{D2@Эскиз1}_{Количество}` и конфигурации с `D1@Эскиз1`=100мм, `D2@Эскиз1`=100мм, свойством `Количество`=8 результат: `уголок_100x100_8`. Если размер не найден или скобки непарные — `ok=False`.

- [ ] **Step 2: Добавить функции**

```vba
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
```

- [ ] **Step 3: Скомпилировать**

`Debug → Compile` → без ошибок.

- [ ] **Step 4: Проверить временным `Sub` на открытой детали**

Подставь реальные имена размеров своей детали. Запусти `t_tmpl`, сверь Immediate, потом удали `Sub`:
```vba
Sub t_tmpl()
    Dim m As ModelDoc2
    Set m = Application.SldWorks.ActiveDoc
    Dim cfg As String, okN As Boolean
    cfg = m.ConfigurationManager.ActiveConfiguration.Name
    Debug.Print BuildNameFromTemplate(m, cfg, "test_{D1@Эскиз1}_{Количество}", okN); " ok="; okN
    Debug.Print BuildNameFromTemplate(m, cfg, "bad_{НетТакого@Эскиз999}", okN); " ok="; okN
End Sub
```
Expected: первая строка — имя с реальными значениями и `ok=True`; вторая — пустое имя и `ok=False` (размер не найден).

- [ ] **Step 5: Commit**

```bash
git add src/TemplateDxfExport.bas
git commit -m "feat: template placeholder substitution"
```

---

## Task 5: Полный `Sub main()`

**Files:**
- Modify: `src/TemplateDxfExport.bas` (заменить тело `main`)

- [ ] **Step 1: Заменить `Sub main()` целиком**

```vba
Sub main()
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc

    If swModel Is Nothing Then MsgBox "Открой деталь!": Exit Sub
    If swModel.GetType <> swDocPART Then MsgBox "Это не деталь!": Exit Sub
    Set swPart = swModel

    Dim modelPath As String
    modelPath = swModel.GetPathName
    If modelPath = "" Then MsgBox "Сначала сохрани файл!": Exit Sub

    Dim template As String
    template = ReadFileProp(swModel, TEMPLATE_PROP)
    If Trim(template) = "" Then
        MsgBox "Добавь в свойства файла (вкладка «Настройка») свойство '" & TEMPLATE_PROP & "'" & vbCrLf & _
               "с шаблоном имени, например:" & vbCrLf & _
               "уголок_{D1@Эскиз1}x{D2@Эскиз1}_{Длина}_Qty_{Количество}"
        Exit Sub
    End If
    If CountChar(template, "{") <> CountChar(template, "}") Then
        MsgBox "В шаблоне непарные скобки { }:" & vbCrLf & template
        Exit Sub
    End If

    Dim outFolder As String
    outFolder = GetFolder(Left(modelPath, InStrRev(modelPath, "\")))
    If outFolder = "" Then MsgBox "Отменено.": Exit Sub
    If Right(outFolder, 1) <> "\" Then outFolder = outFolder & "\"

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

    swModel.ShowConfiguration2 origConf
    swModel.EditRebuild3

    Dim msg As String
    msg = "Готово!" & vbCrLf & "Выгружено: " & okCount & vbCrLf & _
          "Пропущено: " & skipCount & vbCrLf & "Ошибок: " & failCount
    If Len(report) > 0 Then msg = msg & vbCrLf & vbCrLf & "Подробности:" & vbCrLf & report
    MsgBox msg
End Sub
```

- [ ] **Step 2: Скомпилировать**

`Debug → Compile` → без ошибок.

- [ ] **Step 3: Commit**

```bash
git add src/TemplateDxfExport.bas
git commit -m "feat: full template-driven export loop"
```

---

## Task 6: Ручная приёмка в SolidWorks

**Files:** (проверка, без правок кода)

- [ ] **Step 1: Подготовить деталь**

Деталь листового металла с **3+ конфигурациями** (разные значения размеров/полок). Добавь свойство файла `Шаблон_имени`, например:
`уголок_{D1@Эскиз1}x{D2@Эскиз1}_{Длина}_Qty_{Количество}`
(подставь реальные имена размеров; `Длина`, `Количество` — свойства конфигураций, заполни их). Сделай две конфигурации с одинаковыми значениями (проверка перезаписи). Сохрани.

- [ ] **Step 2: Запустить и проверить по чек-листу**

`F5` → выбери папку. Проверь:
- [ ] на каждую конфигурацию с листовым металлом создан DXF;
- [ ] имена — результат подстановки шаблона, размеры (`{…@…}`) совпадают с реальными значениями конфигурации;
- [ ] свойства (`{…}`) подставлены из свойств;
- [ ] две одинаковые конфигурации дали **два** файла (второй с суффиксом), без перезаписи;
- [ ] нет свойства `Шаблон_имени` → понятная подсказка, без падения;
- [ ] непарные скобки в шаблоне → понятное сообщение;
- [ ] после прогона деталь на **исходной** конфигурации;
- [ ] итог: Выгружено/Пропущено/Ошибок корректны.

- [ ] **Step 3: Зафиксировать актуальный исходник**

Если правил код в редакторе — перезапиши `src/TemplateDxfExport.bas` (`Export File…`) и:
```bash
git add src/TemplateDxfExport.bas
git commit -m "fix: adjustments from manual acceptance"
```
(если правок не было — пропусти)

---

## Task 7: README + публикация

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Переименовать существующий раздел про пластины**

В `README.md` после заголовка-описания добавь подзаголовок для первого макроса. Найди строку заголовка установки и преобразуй документ в два раздела. Конкретно — добавь перед разделом «## Требования» строку:
```markdown
## Макрос 1 (пластины): имя из геометрии — `src/Macro11.bas`
```

- [ ] **Step 2: Добавить раздел про второй макрос**

В конец `README.md` (перед «## Состав репозитория») добавь:
```markdown
## Макрос 2 (универсальный): имя по шаблону — `src/TemplateDxfExport.bas`

Для деталей любого типа (уголки, профили и т.п.), где имя нужно собирать из меняющихся значений. Имя DXF задаётся **шаблоном** в самой детали — один макрос на все типы деталей.

### Шаблон
В свойствах файла детали (`Файл → Свойства → Настройка`) заведи свойство **`Шаблон_имени`**, например:

```
уголок_{D1@Эскиз1}x{D2@Эскиз1}_{Длина}_Qty_{Количество}
```

Правило подстановки:
- `{…@…}` (есть `@`) → значение **размера** из геометрии (в мм). Те же числа, что задают деталь, попадают в имя — дублировать не нужно.
- `{…}` (без `@`) → значение **свойства** (сначала конфигурации, потом файла).

Результат: `уголок_100x100_300_Qty_8.dxf`.

### Установка и запуск
Как Макрос 1, но вставляешь `src/TemplateDxfExport.bas` в новый макрос. Перед запуском убедись, что у детали задано свойство `Шаблон_имени`. Если его нет — макрос подскажет.
```

- [ ] **Step 3: Запушить всё**

```bash
git add README.md
git commit -m "docs: document template-based macro (Macro 2)"
git push
```
Expected: `git push` отправляет коммиты (новый макрос, спека, план, README) в `origin/main`.

- [ ] **Step 4: Проверить на GitHub**

Открой https://github.com/napilniksalda-star/solidworks-dxf-batch-export — в `src/` два `.bas`, README описывает оба макроса. Убедись, что клиентских файлов (`.SLDPRT`, `.xlsx`) нет.

---

## Appendix: Полный модуль `TemplateDxfExport.bas`

Итоговый файл — конкатенация в порядке: `Attribute`/`Option Explicit`/`Private Const`/`Dim` (Задача 1) → `Sub main` (Задача 5) → `ResolvePlaceholder`, `BuildNameFromTemplate` (Задача 4) → `ReadFileProp`, `ReadConfigProp`, `HasSheetMetal` (Задача 3) → `FormatNum`, `SanitizeName`, `CountChar`, `UniqueFilePath`, `GetFolder` (Задача 2). Порядок функций после `Sub main` не важен.

---

## Self-Review (выполнено при написании плана)

**1. Покрытие спецификации:**
- Шаблон в свойстве файла `Шаблон_имени` → `ReadFileProp` + main (Tasks 3, 5). ✓
- Подстановка `{…}`, правило `@` → `ResolvePlaceholder` + `BuildNameFromTemplate` (Task 4). ✓
- Размер из геометрии (мм, формат) → `ResolvePlaceholder` + `FormatNum` (Tasks 2, 4). ✓
- Свойство: конфиг, затем файл → `ResolvePlaceholder` + `ReadConfigProp`/`ReadFileProp` (Tasks 3, 4). ✓
- Парность скобок, нет шаблона, отмена папки → main (Task 5). ✓
- Пропуск без листового металла → `HasSheetMetal` + main (Tasks 3, 5). ✓
- Защита от перезаписи → `UniqueFilePath` (Task 2). ✓
- Возврат конфигурации, отчёт → main (Task 5). ✓
- Не трогаем макрос пластин; доставка в репозиторий + README → Task 7. ✓

**2. Плейсхолдеры:** не найдено — в каждом шаге полный код/команды/ожидаемый результат. ✓

**3. Согласованность имён:** `ReadFileProp`, `ReadConfigProp`, `HasSheetMetal`, `ResolvePlaceholder`, `BuildNameFromTemplate`, `FormatNum`, `SanitizeName`, `CountChar`, `UniqueFilePath`, `GetFolder`, константа `TEMPLATE_PROP` — совпадают между задачами и вызовами в `main`. `UniqueFilePath` получает имя уже с `.dxf` (`SanitizeName(baseName) & ".dxf"`) — соответствует его логике отрезания расширения. ✓
