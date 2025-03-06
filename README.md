# Поиск файлов Excel с макросами

## Вариант 1 - PowerShell

Выполнить в среде PowerShell ISE:
```
$searchPath = "I:\ОИВТ\02-Документы отдела\21 - Сменные инженеры\04 - Личные папки\12 - Борисов Н.А"
$extensions = @("*.xls", "*.xlsx", "*.xlsm", "*.xltm")
$results = @()

Get-ChildItem -Path $searchPath -Recurse -Include $extensions -File | ForEach-Object {
    $filePath = $_.FullName
    
    if ($filePath -match "\.xls$|\.xltm$|\.xlsm$") {
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
## Вариант 5 - Ansible

```
---
- name: Find Excel files with macros
  hosts: удалённый_сервер
  tasks:
    - name: Find .xlsm files
      find:
        paths: I:\ОИВТ\02-Документы отдела\21 - Сменные инженеры\04 - Личные папки\12 - Борисов Н.А
        patterns: '*.xlsm'
      register: excel_files

    - name: Write file paths to text file
      copy:
        content: |
          {% for file in excel_files.files %}
          {{ file.path }}
          {% endfor %}
        dest: I:\ОИВТ\02-Документы отдела\21 - Сменные инженеры\04 - Личные папки\12 - Борисов Н.А\FilesList.txt
```
# Вариант от ПАО
```

##Requires -Modules NTFSSecurity
#[CmdletBinding()]
#param(
#[string]$InFile,
#[string]$OutSuff
#)
#begin
#{
    $TIME_STAMP    = Get-Date -Format "yyyyMMdd-HHmmss"
    $ScriptPath     = $MyInvocation.MyCommand.Path
    $ScriptDir     = split-path -parent $ScriptPath
    $prg=$myinvocation.mycommand.name
    
    $UsernameMB = "MB-TN\PFO-ADREADER-SRV"
    $CredentialMB = Get-Credential -Credential $UsernameMB

    $UsernameDVA = "DVA\svc_reader"
    $CredentialDVA = Get-Credential -Credential $UsernameDVA

    $DVA_Domain_SID="S-1-5-21-3552983670-3413165326-445540045"
    $MB_Domain_SID="S-1-5-21-3129793653-3252460176-3566091208"

    $_LExclude=".snapshot
.TemporaryItems"

$_ArExclude = $_LExclude -split [System.Environment]::NewLine

    If (-not (Test-Path "$ScriptDir\Reports")) {mkdir "$ScriptDir\Reports"}

    function ADSI-Search-By-SID {
    param(
    [ValidateNotNullOrEmpty()]
    [string]$TargetDomanFQDN,
    [ValidateNotNullOrEmpty()]
    [string]$TargetObjectSID,
    [System.Management.Automation.PSCredential] $Credential = $null
    )
        
    $adsisearcher		= [adsisearcher]""
    $DomanFQDN = $TargetDomanFQDN
    $ObjectSID = $TargetObjectSID
    $Root		= "DC=" + ($DomanFQDN.split(".") -join ",DC=")
    $Base	= "LDAP://" + $Root
    $SAN = $null
    if($Credential -ne $null)
    {
        $adsisearcher.SearchRoot	=  New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList $Base, $($Credential.UserName), $($Credential.GetNetworkCredential().password)
    }
    else
    {
        $adsisearcher.SearchRoot	=  New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList $Base
    }
    $adsisearcher.filter	= "(objectsid=$ObjectSID)"
        $AdsiObject = $null
        $AdsiObject	= $adsisearcher.FindOne()
        if($AdsiObject -ne $null)
        {
            $SAN = $($AdsiObject.Properties['samaccountname'] | select -Last 1)
        }

    $adsisearcher = $null
    return $SAN
    }



    $OPError = [ordered]@{
        Path = ""
        Error = ""
    }
    function AddToError
    {
    Param(
	[Parameter(Position=0,Mandatory=$True)]
	[ValidateNotNullOrEmpty()]
	#[string]$FQDN,                # 
    [string]$Path = "",
    [string]$Error = ""
	)
        $OError = New-Object -TypeName PSObject -Property $OPError

        $OError.Path = $Path
        $OError.Error = $Error

        $Script:AL_Error.Add($OError) | out-null
        
        #$Error.Clear()
    }

    $OPResult = [ordered]@{
        owner = ""
        write = ""
        access = ""
        hash = ""
        name = ""
        path = ""
    }
    function AddToResult
    {
    Param(
	[Parameter(Position=0,Mandatory=$True)]
	[ValidateNotNullOrEmpty()]
	[string]$owner = "",
    [string]$write = "",
    [string]$access = "",
    [string]$hash = "",
    [string]$name = "",
    [string]$path = ""
	)
        $OResult = New-Object -TypeName PSObject -Property $OPResult

        $OResult.owner = $owner
        $OResult.write = $write
        $OResult.access = $access
        $OResult.hash = $hash
        $OResult.name = $name
        $OResult.path = $path

        $Script:AL_ScanResult.Add($OResult) | out-null
    }

    function _ScanFilesRec
    {
    Param(
	[Parameter(Position=0,Mandatory=$True)]
	[ValidateNotNullOrEmpty()]
	[string]$_SrcPath = ""
    )
        Get-ChildItem  $_SrcPath -File -Include ("*.vsd", "*.vsdx") -Recurse | % {
        
            $_Owner = "-"
            $_LWTime = ""
            $_LATime = ""
            $_Hash = ""
            $_FName = ""
            $_Path = ""
            $FullPath = $_.FullName
            $ACL = $null

            try
            {
                Write-Host $_.FullName

                #$_Owner = ((Get-Acl -LiteralPath $FullPath).Owner).ToUpper()
                $ACL_ = Get-NTFSOwner -Path $FullPath -ErrorAction Stop
                if ($ACL_ -eq $null)
                {
                    throw "Get Owner"
                }
                if($ACL_.Owner.LastError -ne $null)
                {
                    $OBJ = $null
                    $SSID = $ACL_.Owner.Sid.ToString()
                    $_Owner = $SSID
                    if($SSID -like $("$DVA_Domain_SID*"))
                    {
                        $OBJ = ADSI-Search-By-SID -TargetDomanFQDN "dva.tn.corp" -TargetObjectSID $SSID -Credential $CredentialDVA
                        $_Owner = "DVA\$OBJ"
                    }
                    if($SSID -like $("$MB_Domain_SID*"))
                    {
                        $OBJ = ADSI-Search-By-SID -TargetDomanFQDN "mb.tn.corp" -TargetObjectSID $SSID -Credential $CredentialMB
                        $_Owner = "MB-TN\$OBJ"
                    }
                }
                else
                {
                    $_Owner = $ACL_.Owner.AccountName
                }
                $_LWTime = ($_.LastWriteTime).ToString("yyyyMMdd-HHmmss")
                $_LATime = ($_.LastAccessTime).ToString("yyyyMMdd-HHmmss")
                $_Hash = ((Get-FileHash -algorithm md5 $FullPath).hash).ToUpper()
                $_FName = ($_.Name).ToUpper()
                $_Path = ($_.directoryname).ToUpper()

                AddToResult -owner $_Owner -write $_LWTime -access $_LATime -hash $_Hash -name $_FName -path $_Path
            }
            catch
            {
                AddToError -Path $FullPath -Error $Error[0].ToString()
                $Error.Clear()
            }
           
        }
    }


    function _ScanFiles
    {
    Param(
	[Parameter(Position=0,Mandatory=$True)]
	[ValidateNotNullOrEmpty()]
	[string]$_SrcPath = ""
    )
        Get-ChildItem  $_SrcPath -File -Include ("*.vsd", "*.vsdx") | % {
        
            $_Owner = "-"
            $_LWTime = ""
            $_LATime = ""
            $_Hash = ""
            $_FName = ""
            $_Path = ""
            $FullPath = $_.FullName
            $ACL = $null

            try
            {
                Write-Host $_.FullName

                #$_Owner = ((Get-Acl -LiteralPath $FullPath).Owner).ToUpper()
                $ACL_ = Get-NTFSOwner -Path $FullPath -ErrorAction Stop
                if ($ACL_ -eq $null)
                {
                    throw "Get Owner"
                }
                if($ACL_.Owner.LastError -ne $null)
                {
                    $OBJ = $null
                    $SSID = $ACL_.Owner.Sid.ToString()
                    $_Owner = $SSID
                    if($SSID -like $("$DVA_Domain_SID*"))
                    {
                        $OBJ = ADSI-Search-By-SID -TargetDomanFQDN "dva.tn.corp" -TargetObjectSID $SSID -Credential $CredentialDVA
                        $_Owner = "DVA\$OBJ"
                    }
                    if($SSID -like $("$MB_Domain_SID*"))
                    {
                        $OBJ = ADSI-Search-By-SID -TargetDomanFQDN "mb.tn.corp" -TargetObjectSID $SSID -Credential $CredentialMB
                        $_Owner = "MB-TN\$OBJ"
                    }
                }
                else
                {
                    $_Owner = $ACL_.Owner.AccountName
                }
                $_LWTime = ($_.LastWriteTime).ToString("yyyyMMdd-HHmmss")
                $_LATime = ($_.LastAccessTime).ToString("yyyyMMdd-HHmmss")
                $_Hash = ((Get-FileHash -algorithm md5 $FullPath).hash).ToUpper()
                $_FName = ($_.Name).ToUpper()
                $_Path = ($_.directoryname).ToUpper()

                AddToResult -owner $_Owner -write $_LWTime -access $_LATime -hash $_Hash -name $_FName -path $_Path
            }
            catch
            {
                AddToError -Path $FullPath -Error $Error[0].ToString()
                $Error.Clear()
            }
           
        }
    }
    Import-CSV -Delimiter ";" -Path "\\tn.tngrp.ru\df\USR\CR\EfimkinDV\My Documents\EfimkinDV\Скрипты\IN\IN_FRes_MakrosINDocuments.csv" | ForEach-Object {
        $Error.Clear()
        $OutBFileName = $($_."ИД").Trim()
#        $OutFileName = "$OutBFileName.csv"
        $SrcPath = $($_."ПУТЬ").Trim()
	    
        $AL_Error = new-object System.Collections.ArrayList
        $ErrorLog =  "$ScriptDir\Reports\$OutBFileName"+"_ErrorLog.csv"

        try
        {
            if (-not (Test-Path -Path $SrcPath))
            {
                throw "Not Exists"
            }

            $AL_ScanResult = new-object System.Collections.ArrayList
        	
            $ShareResultLog = "$ScriptDir\Reports\$OutBFileName.csv"
            if(Test-Path -Path $ShareResultLog)
            {
                Remove-Item -Path $ShareResultLog -Force
            }
            New-Item -Path $ShareResultLog -ItemType File

            _ScanFiles -_SrcPath $SrcPath
            Get-ChildItem  $SrcPath | % {
                if ($_.Name -notin $_ArExclude)
                {
                    _ScanFilesRec -_SrcPath $_.FullName
                }
            }
            if ($AL_ScanResult.Count -gt 0) {
                $AL_ScanResult | export-csv -Encoding UTF8 -Delimiter ";" -NoTypeInformation -Path $ShareResultLog -Append
            }
            $AL_ScanResult = $null
        }
        catch
        {
            AddToError -Path $SrcPath -Error $Error[0].Exception.Message.ToString()
            $Error.Clear()
        }
        if ($AL_Error.Count -gt 0) {
            $AL_Error | export-csv -Encoding UTF8 -Delimiter ";" -NoTypeInformation -Path $ErrorLog -Append
        }
        $AL_Error = $null
    
    }
    
#}
#process
#{
#    $Error.Clear()
#}
#end
#{
#}
```
