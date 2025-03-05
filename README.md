# Поиск файлов Excel с макросами

## Вариант 1 - PowerShell

Выполнить в среде PowerShell ISE:
```
$searchPath = "I:\ОИВТ\02-Документы отдела\21 - Сменные инженеры\04 - Личные папки\12 - Борисов Н.А"
$extensions = @("*.xls", "*.xlsx", "*.xlsm", "*.xltm")
$results = @()

Get-ChildItem -Path $searchPath -Recurse -Include $extensions -File | ForEach-Object {
    $filePath = $_.FullName
    
    if ($filePath -match "\.xls$|\.xlsx$|\.xlsm$") {
        $excel = $null
        $workbook = $null
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $workbook = $excel.Workbooks.Open($filePath)
            if ($workbook.HasVBProject) {
                $results += "Excel файл с макросом: $filePath"
            }
        }
        catch {
            $results += "Ошибка при обработке: $filePath - $_"
        }
        finally {
            if ($workbook) {
                $workbook.Close()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
            }
            if ($excel) {
                $excel.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            }
        }
    }
}

$outputFile = "I:\ОИВТ\02-Документы отдела\21 - Сменные инженеры\04 - Личные папки\12 - Борисов Н.А\FilesMacros.txt"
$results | Out-File -FilePath $outputFile -Encoding UTF8
Write-Host
```

## Вариант 2 - CMD
```
dir I:\ОИВТ\02-Документы отдела\21 - Сменные инженеры\04 - Личные папки\12 - Борисов Н.А\*.xls* /S /B > I:\ОИВТ\02-Документы отдела\21 - Сменные инженеры\04 - Личные папки\12 - Борисов Н.А\FilesList.txt
```
## Вариант 3 - VBScript

Создать файл и сохранить с расширение .vbs
```
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder("I:\ОИВТ\02-Документы отдела\21 - Сменные инженеры\04 - Личные папки\12 - Борисов Н.А")
Set objOutputFile = objFSO.CreateTextFile("I:\ОИВТ\02-Документы отдела\21 - Сменные инженеры\04 - Личные папки\12 - Борисов Н.А\FilesWithMacros.txt", True)

SearchFolder objFolder

Sub SearchFolder(Folder)
    For Each objFile in Folder.Files
        If InStr(objFile.Name, ".xls") Or InStr(objFile.Name, ".xlsx") Or InStr(objFile.Name, ".xlsm") Then
            objOutputFile.WriteLine "Файл найден: " & objFile.Path
        End If
    Next
    For Each objSubFolder in Folder.SubFolders
        SearchFolder objSubFolder
    Next
End Sub

objOutputFile.Close
```

## Вариант 4 - Python

Запустить скрипт .py
```
import os

directory = r"I:\ОИВТ\02-Документы отдела\21 - Сменные инженеры\04 - Личные папки\12 - Борисов Н.А"
output_file = r"I:\ОИВТ\02-Документы отдела\21 - Сменные инженеры\04 - Личные папки\12 - Борисов Н.А\files_with_macros.txt"

with open(output_file, 'w') as f:
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith(".xlsm"):
                f.write(os.path.join(root, file) + '\n')
```
