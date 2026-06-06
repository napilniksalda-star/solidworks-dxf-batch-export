# Пакетный экспорт развёрток в DXF — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Доработать SolidWorks VBA-макрос так, чтобы он выгружал развёртку каждой конфигурации в DXF, выбирая папку и беря длину/ширину/толщину из геометрии (не из статичных свойств).

**Architecture:** Один VBA-модуль `Macro11` внутри `Macro1.swp`. `Sub main()` — оркестратор: проверки → выбор папки → цикл по конфигурациям → возврат исходной конфигурации → отчёт. Вся «хрупкая» логика вынесена в маленькие чистые функции-помощники, которые проверяются по отдельности.

**Tech Stack:** SolidWorks API (sldworks.tlb, swconst.tlb), VBA, `Shell.Application` (поздняя привязка) для диалога выбора папки.

**Spec:** `docs/superpowers/specs/2026-06-06-solidworks-dxf-batch-export-design.md`

---

## Как «тестировать» VBA-макрос (важно)

Автоматического раннера (pytest) для VBA внутри SolidWorks тут нет. Поэтому:

- **Разработку ведём в редакторе VBA** самого `Macro1.swp` (`Сервис → Макрос → Изменить…`). Там доступны: **компиляция** (`Debug → Compile`), **окно Immediate** (`Ctrl+G`) и **запуск** (`F5`).
- **Чистые функции** проверяем в Immediate: набираем `?ИмяФункции(аргументы)` и сверяем с ожидаемым.
- **Функции с API/файлами** проверяем маленьким временным `Sub`, который печатает результат (`Debug.Print` / `MsgBox`). После проверки временный `Sub` удаляем.
- **Весь макрос** проверяем вручную по протоколу в Задаче 9 на реальной детали с несколькими конфигурациями.
- **Текстовая копия** модуля хранится в `src/Macro11.bas` (экспорт из редактора: правый клик по модулю → `Export File…`). Это и есть артефакт для чтения/diff/коммита.

> **Кодировка:** в коде есть кириллица (`Эскиз1`, русские сообщения). Загружай код в редактор **копированием-вставкой** из `src/Macro11.bas` (открытого в Unicode-редакторе), а не через `File → Import` — так кириллица не испортится.

> **Git необязателен:** папка — не репозиторий. Задача 1 (опциональная) его заводит. Если её пропустить — пропускай и шаги «Commit».

---

## File Structure

| Файл | Ответственность |
|---|---|
| `src/Macro11.bas` | Текстовый исходник VBA-модуля (единственный модуль макроса). Создаётся и наполняется по задачам. |
| `Macro1.swp` | Бинарный макрос SolidWorks. В него вставляется итоговый код модуля `Macro11` (Задача 9). |
| `docs/superpowers/specs/2026-06-06-solidworks-dxf-batch-export-design.md` | Спецификация (уже есть). |

Модуль один (так устроены макросы SolidWorks), но внутри он разбит на функции с одной ответственностью каждая: `GetFolder`, `ReadDimMM`, `ReadThicknessMM`, `ReadConfigProp`, `FormatNum`, `SanitizeName`, `BuildBaseName`, `UniqueFilePath`, и оркестратор `main`.

---

## Task 1 (опционально): Инициализировать git и базовый коммит

**Files:**
- Create: `.gitignore`

- [ ] **Step 1: Инициализировать репозиторий**

Run:
```bash
git init
```

- [ ] **Step 2: Создать `.gitignore`**

Создать `.gitignore` со следующим содержимым:
```gitignore
# Временные/системные
*.swp.bak
~$*
Thumbs.db
_*.py
_*.txt
```

> Примечание: `*.swp` — это макрос SolidWorks (нужный файл), НЕ vim-своп. Не добавляй `*.swp` в ignore.

- [ ] **Step 3: Базовый коммит**

Run:
```bash
git add -A
git commit -m "chore: baseline (existing macro, order example, spec)"
```
Expected: коммит создан, в нём `Macro1.swp`, `ЗАКАЗ 46 - для чертежей.xlsx`, `docs/...design.md`.

---

## Task 2: Каркас модуля + проверки входа

Создаём модуль с константами и `Sub main()`, который пока только проверяет вход (открыта ли сохранённая деталь). Это даёт компилируемую основу.

**Files:**
- Create: `src/Macro11.bas`

- [ ] **Step 1: Создать `src/Macro11.bas` с каркасом**

```vba
' Macro11 — вставляй копипастом в модуль макроса (строка Attribute не нужна)
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

    MsgBox "Проверки пройдены (заглушка)."

End Sub
```

- [ ] **Step 2: Загрузить в редактор и скомпилировать**

Открой `Macro1.swp` (`Сервис → Макрос → Изменить…`), вставь код в модуль `Macro11` (замени содержимое). В редакторе: `Debug → Compile`.
Expected: компиляция без ошибок.

- [ ] **Step 3: Проверить заглушку**

Закрой все документы в SolidWorks, нажми `F5`.
Expected: появляется «Открой деталь!». Открой сохранённую деталь, снова `F5` → «Проверки пройдены (заглушка)».

- [ ] **Step 4: Commit**

Экспортируй модуль в `src/Macro11.bas` (правый клик по модулю → `Export File…`), затем:
```bash
git add src/Macro11.bas
git commit -m "feat: module skeleton with entry guards"
```

---

## Task 3: `FormatNum` — числа для имени файла

Преобразует мм в строку: округление до 2 знаков, точка как разделитель, без хвостовых нулей.

**Files:**
- Modify: `src/Macro11.bas`

- [ ] **Step 1: Задать ожидаемое поведение (таблица примеров)**

| Вход | Ожидаемый выход |
|---|---|
| `150` | `150` |
| `52` | `52` |
| `1.5` | `1.5` |
| `2` | `2` |
| `1.499999` | `1.5` |
| `176.004` | `176` |

- [ ] **Step 2: Добавить функцию в модуль**

```vba
' Str$ не зависит от локали (всегда "."), не оставляет хвостовую точку
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
```

- [ ] **Step 3: Скомпилировать**

`Debug → Compile`. Expected: без ошибок.

- [ ] **Step 4: Проверить в Immediate (`Ctrl+G`)**

Набери по очереди и сверь:
```
?FormatNum(150)        -> 150
?FormatNum(1.5)        -> 1.5
?FormatNum(2)          -> 2
?FormatNum(1.499999)   -> 1.5
?FormatNum(176.004)    -> 176
```
Expected: вывод совпадает с таблицей Step 1.

- [ ] **Step 5: Commit**

```bash
git add src/Macro11.bas
git commit -m "feat: FormatNum helper for filename numbers"
```

---

## Task 4: `SanitizeName` + `BuildBaseName` — имя файла

`SanitizeName` убирает недопустимые в имени файла символы. `BuildBaseName` собирает полное имя.

**Files:**
- Modify: `src/Macro11.bas`

- [ ] **Step 1: Ожидаемое поведение**

| Функция | Вход | Выход |
|---|---|---|
| `SanitizeName` | `a/b:c*?` | `a_b_c__` |
| `BuildBaseName` | `(150, 52, 1.5, "2")` | `Plate_150x52x1.5_Qty_2.dxf` |
| `BuildBaseName` | `(50, 176, 2, "150")` | `Plate_50x176x2_Qty_150.dxf` |

- [ ] **Step 2: Добавить функции**

```vba
Private Function SanitizeName(ByVal s As String) As String
    Dim bad As Variant, i As Integer
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, bad(i), "_")
    Next i
    SanitizeName = s
End Function

Private Function BuildBaseName(ByVal lMM As Double, ByVal wMM As Double, _
                              ByVal tMM As Double, ByVal qty As String) As String
    BuildBaseName = FILE_PREFIX & FormatNum(lMM) & "x" & FormatNum(wMM) & "x" & _
                    FormatNum(tMM) & "_Qty_" & SanitizeName(qty) & ".dxf"
End Function
```

- [ ] **Step 3: Скомпилировать**

`Debug → Compile`. Expected: без ошибок.

- [ ] **Step 4: Проверить в Immediate**

```
?SanitizeName("a/b:c*?")        -> a_b_c__
?BuildBaseName(150,52,1.5,"2")  -> Plate_150x52x1.5_Qty_2.dxf
?BuildBaseName(50,176,2,"150")  -> Plate_50x176x2_Qty_150.dxf
```
Expected: совпадает с таблицей Step 1.

- [ ] **Step 5: Commit**

```bash
git add src/Macro11.bas
git commit -m "feat: filename builder and sanitizer"
```

---

## Task 5: `UniqueFilePath` — защита от перезаписи

Если файл существует, добавляет имя конфигурации, затем счётчик `_2`, `_3`…

**Files:**
- Modify: `src/Macro11.bas`

- [ ] **Step 1: Ожидаемое поведение**

Папка пуста → возвращает `folder & baseName`. Если такой файл уже есть → `..._{conf}.dxf`. Если и он есть → `..._{conf}_2.dxf`, и т.д.

- [ ] **Step 2: Добавить функцию**

```vba
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
    stem = Left(baseName, Len(baseName) - Len(ext))   ' имя без ".dxf"

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
```

- [ ] **Step 3: Скомпилировать**

`Debug → Compile`. Expected: без ошибок.

- [ ] **Step 4: Проверить временным `Sub`**

Вставь временно, запусти (`F5` на `t_unique`), сверь окно Immediate, потом удали этот `Sub`:
```vba
Sub t_unique()
    Dim fld As String
    fld = Environ$("TEMP") & "\t_unique\"
    If Dir(fld, vbDirectory) = "" Then MkDir fld
    ' очистка
    On Error Resume Next: Kill fld & "*.dxf": On Error GoTo 0

    Debug.Print UniqueFilePath(fld, "Plate_200x52x1.5_Qty_2.dxf", "46-7")   ' нет файла -> базовое имя
    ' создаём базовый файл и проверяем суффикс конфигурации
    Open fld & "Plate_200x52x1.5_Qty_2.dxf" For Output As #1: Close #1
    Debug.Print UniqueFilePath(fld, "Plate_200x52x1.5_Qty_2.dxf", "46-7")   ' -> ..._46-7.dxf
    ' создаём и его, проверяем счётчик
    Open fld & "Plate_200x52x1.5_Qty_2_46-7.dxf" For Output As #1: Close #1
    Debug.Print UniqueFilePath(fld, "Plate_200x52x1.5_Qty_2.dxf", "46-7")   ' -> ..._46-7_2.dxf
End Sub
```
Expected (три строки):
```
...\Plate_200x52x1.5_Qty_2.dxf
...\Plate_200x52x1.5_Qty_2_46-7.dxf
...\Plate_200x52x1.5_Qty_2_46-7_2.dxf
```

- [ ] **Step 5: Commit**

```bash
git add src/Macro11.bas
git commit -m "feat: UniqueFilePath collision protection"
```

---

## Task 6: `GetFolder` — диалог выбора папки

**Files:**
- Modify: `src/Macro11.bas`

- [ ] **Step 1: Ожидаемое поведение**

Открывает системный диалог выбора папки со стартом в `startPath`. Возвращает выбранный путь или `""` при отмене.

- [ ] **Step 2: Добавить функцию**

```vba
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
```

- [ ] **Step 3: Скомпилировать**

`Debug → Compile`. Expected: без ошибок.

- [ ] **Step 4: Проверить временным `Sub`**

```vba
Sub t_folder()
    MsgBox "[" & GetFolder("C:\") & "]"
End Sub
```
Запусти `t_folder`. Expected: появляется диалог, стартует с `C:\`; выбор папки → MsgBox с её путём; «Отмена» → MsgBox `[]`. Затем удали `t_folder`.

- [ ] **Step 5: Commit**

```bash
git add src/Macro11.bas
git commit -m "feat: GetFolder browse dialog"
```

---

## Task 7: `ReadDimMM`, `ReadThicknessMM`, `ReadConfigProp` — чтение из модели

**Files:**
- Modify: `src/Macro11.bas`

- [ ] **Step 1: Ожидаемое поведение**

- `ReadDimMM` возвращает значение размера в мм для активной конфигурации; `ok=False`, если размер не найден.
- `ReadThicknessMM` возвращает толщину листового металла в мм; `ok=False`, если элемента Sheet-Metal нет.
- `ReadConfigProp` возвращает разрешённое значение свойства конфигурации (или `""`).

- [ ] **Step 2: Добавить функции**

```vba
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
```

- [ ] **Step 3: Скомпилировать**

`Debug → Compile`. Expected: без ошибок.

- [ ] **Step 4: Проверить на открытой детали временным `Sub`**

Открой реальную деталь-пластину (листовой металл, со свойством `Quantity` в активной конфигурации):
```vba
Sub t_reads()
    Dim m As ModelDoc2
    Set m = Application.SldWorks.ActiveDoc
    Dim okL As Boolean, okW As Boolean, okT As Boolean
    Debug.Print "L="; ReadDimMM(m, LENGTH_DIM, okL); " ok="; okL
    Debug.Print "W="; ReadDimMM(m, WIDTH_DIM, okW); " ok="; okW
    Debug.Print "T="; ReadThicknessMM(m, okT); " ok="; okT
    Debug.Print "Qty="; ReadConfigProp(m, m.ConfigurationManager.ActiveConfiguration.Name, QTY_PROP)
End Sub
```
Expected: L/W/T — реальные размеры активной конфигурации в мм, все `ok=True`; `Qty` — значение свойства. Если `okL/okW=False` — проверь имена размеров в константах. Затем удали `t_reads`.

- [ ] **Step 5: Commit**

```bash
git add src/Macro11.bas
git commit -m "feat: geometry/property readers"
```

---

## Task 8: Собрать `Sub main()` — полный цикл

Заменяем заглушку `Sub main()` на полную логику: выбор папки, цикл по конфигурациям, экспорт, возврат конфигурации, отчёт.

**Files:**
- Modify: `src/Macro11.bas` (заменить тело `Sub main()` из Задачи 2)

- [ ] **Step 1: Заменить `Sub main()` целиком**

```vba
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

        Dim okT As Boolean, tMM As Double
        tMM = ReadThicknessMM(swModel, okT)
        If Not okT Then
            skipCount = skipCount + 1
            report = report & "ПРОПУСК (нет листового металла): " & confName & vbCrLf
            GoTo ContinueLoop
        End If

        Dim okL As Boolean, okW As Boolean, lMM As Double, wMM As Double
        lMM = ReadDimMM(swModel, LENGTH_DIM, okL)
        wMM = ReadDimMM(swModel, WIDTH_DIM, okW)
        If (Not okL) Or (Not okW) Then
            failCount = failCount + 1
            report = report & "ОШИБКА (нет размера " & LENGTH_DIM & "/" & WIDTH_DIM & "): " & confName & vbCrLf
            GoTo ContinueLoop
        End If

        Dim qty As String
        qty = ReadConfigProp(swModel, confName, QTY_PROP)
        If Trim(qty) = "" Then qty = "NA"

        Dim fullPath As String
        fullPath = UniqueFilePath(outFolder, BuildBaseName(lMM, wMM, tMM, qty), confName)

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
```

- [ ] **Step 2: Скомпилировать**

`Debug → Compile`. Expected: без ошибок.

- [ ] **Step 3: Commit**

```bash
git add src/Macro11.bas
git commit -m "feat: full export loop in main()"
```

---

## Task 9: Ручная приёмка в SolidWorks + установка в макрос

Это главный тест: проверяем поведение на реальной детали и фиксируем код в `Macro1.swp`.

**Files:**
- Modify: `Macro1.swp` (вставить итоговый модуль)

- [ ] **Step 1: Подготовить тестовую деталь**

Деталь листового металла с **3+ конфигурациями**, где различаются длина, ширина и толщина; в одной конфигурации намеренно сделай те же размеры и `Quantity`, что и в другой (проверка перезаписи). Заполни `Quantity` в свойствах конфигураций. Сохрани деталь.

- [ ] **Step 2: Установить код в макрос**

Открой `Macro1.swp` в редакторе VBA, открой `src/Macro11.bas` в Unicode-редакторе, **скопируй-вставь** весь код в модуль `Macro11` (замени содержимое). `Debug → Compile` → без ошибок.

- [ ] **Step 3: Запустить и проверить по чек-листу**

Открой тестовую деталь, `F5`, выбери папку. Проверь:

- [ ] Появился диалог выбора папки; «Отмена» даёт «Отменено» и выход.
- [ ] В выбранной папке создан один DXF на каждую конфигурацию с листовым металлом.
- [ ] Имена вида `Plate_{Д}x{Ш}x{Т}_Qty_{К}.dxf` и **отражают реальные размеры** (разные ширины/толщины — разные, бага «всегда 52» нет).
- [ ] Две конфигурации с одинаковыми размерами/кол-вом дали **два разных файла** (второй с суффиксом конфигурации), перезаписи нет.
- [ ] Открой пару DXF в любом просмотрщике — контур развёртки корректный.
- [ ] После завершения деталь осталась на **исходной** конфигурации (та, что была активна до запуска).
- [ ] Итоговое сообщение показывает корректные счётчики (Выгружено/Пропущено/Ошибок).
- [ ] Конфигурация без листового металла (если есть) попала в «Пропущено», а не уронила макрос.

Если пункт не прошёл — вернись к соответствующей задаче-функции, исправь, перекомпилируй, повтори.

- [ ] **Step 4: Сохранить и закоммитить**

Сохрани `Macro1.swp` в SolidWorks (сохранится бинарь с новым кодом). Затем:
```bash
git add Macro1.swp src/Macro11.bas
git commit -m "feat: install updated DXF batch export macro"
```

---

## Appendix: Полный модуль `Macro11` (для копирования)

Итоговый `src/Macro11.bas` — это конкатенация: строка `Attribute`, `Option Explicit`, блок констант и `Dim` (Задача 2) → `Sub main()` (Задача 8) → функции-помощники (Задачи 3–7): `FormatNum`, `SanitizeName`, `BuildBaseName`, `UniqueFilePath`, `GetFolder`, `ReadDimMM`, `ReadThicknessMM`, `ReadConfigProp`. Порядок функций после `Sub main()` не важен.

---

## Self-Review (выполнено при написании плана)

**1. Покрытие спецификации:**
- Выбор папки → Task 6 + main (Task 8). ✓
- Длина/ширина из геометрии → `ReadDimMM` (Task 7), используется в main. ✓
- Толщина из Sheet-Metal → `ReadThicknessMM` (Task 7). ✓
- Количество из свойства → `ReadConfigProp` (Task 7). ✓
- Имя `Plate_ДхШхТ_Qty_К`, точка-разделитель, без хвостовых нулей → `FormatNum`/`BuildBaseName` (Tasks 3–4). ✓
- Защита от перезаписи → `UniqueFilePath` (Task 5). ✓
- Возврат исходной конфигурации → main (Task 8). ✓
- Пропуск конфигураций без листового металла + отчёт → main (Task 8). ✓
- Итоговый отчёт со счётчиками → main (Task 8). ✓
- Краевые случаи (нет/не сохранён/не деталь, отмена, пустой Quantity=NA) → Tasks 2, 8. ✓

**2. Плейсхолдеры:** не найдено — в каждом шаге полный код/команды/ожидаемый результат. ✓

**3. Согласованность имён:** `FormatNum`, `SanitizeName`, `BuildBaseName`, `UniqueFilePath`, `GetFolder`, `ReadDimMM`, `ReadThicknessMM`, `ReadConfigProp`, константы `LENGTH_DIM/WIDTH_DIM/QTY_PROP/FILE_PREFIX` — совпадают между задачами-определениями и вызовами в `main`. ✓
