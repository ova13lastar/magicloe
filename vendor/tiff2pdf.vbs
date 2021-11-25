' tiff2pdf.vbs script
' Part of PDFCreator
' License: GPL
' Homepage: http://www.pdfforge.org/products/pdfcreator
' Windows Scripting Host version: 5.1
' Version: 1.2
' Date: 25/11/2021
' Author: Yann DANIEL

Option Explicit

Const WshRunning = 0
Const WshFailed = 1
Const sImages2PdfCPath = "C:\Program Files (x86)\PDFCreator\Images2PDF\Images2PdfC.exe"
Const sSelectCLOEPrinterPath = "C:\APPLILOC\MAGICLOE\vendor\select_cloe_printer.exe"
Const sPdfSuffix = "_magicloe"
Const sAcrobatReaderRegPathx32 = "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\AcroRd32.exe\Path"
Const sAcrobatReaderRegValx32 = "AcroRd32.exe"
Const sAcrobatReaderRegPathx64 = "HKLM\Software\Microsoft\Windows\CurrentVersion\App Paths\Acrobat.exe\Path"
Const sAcrobatReaderRegValx64 = "Acrobat.exe"

Dim objArgs, sInputTiffName, sOutputPdfName, fso, i, AppTitle, Scriptname, ScriptBasename, sAcrobatReaderPath
Dim WshShell, oExec, sCommand

Set fso = CreateObject("Scripting.FileSystemObject")
Scriptname = fso.GetFileName(Wscript.ScriptFullname)
ScriptBasename = fso.GetFileName(Wscript.ScriptFullname)
AppTitle = "PDFCreator - " & ScriptBaseName

'--- On verifie que la version de Vbscript est compatible
ExitIfVbsNotCompatible()

'--- On verifie que PdfCreator (et Images2Pdfc) est bien installé
ExitIfExecutableNotInstalled(sImages2PdfCPath)

'--- On récupère les arguments
Set objArgs = WScript.Arguments

'--- On vérifie que l'argument est bien passé
ExitIfWrongNumberOfArgs(1)

'--- On initialise le Shell
Set WshShell = WScript.CreateObject("WScript.Shell")

'--- On verifie que Acrobat Reader est bien accessible 
If Not RegReadValue(sAcrobatReaderRegPathx64, sAcrobatReaderPath) Then
  If Not RegReadValue(sAcrobatReaderRegPathx32, sAcrobatReaderPath) Then
    MsgBox "Chemin Acrobat Reader non trouve dans le registre !", vbExclamation + vbSystemModal, AppTitle
    WScript.Quit
  Else
    '--- Acrobat Reader x32 trouve
    sAcrobatReaderPath = sAcrobatReaderPath & sAcrobatReaderRegValx32
  End If
Else
  '--- Acrobat Reader x64 trouve
  sAcrobatReaderPath = sAcrobatReaderPath & sAcrobatReaderRegValx64
End if

'--- On boucle sur les arguments
For i = 0 to objArgs.Count - 1
    '--- On recupere le path du fichier en entree
    sInputTiffName = objArgs(i)
    sOutputPdfName = GetFilenameWithoutExtension(sInputTiffName) & sPdfSuffix & ".pdf"
    '--- On lance la conversion TIFF => PDF
    sCommand = """" & sImages2PdfCPath & """ /i """ & sInputTiffName & """ /e """ & sOutputPdfName & """"
    Set oExec = WshShell.Exec(sCommand)
    While oExec.Status = WshRunning
        WScript.Sleep 50
    Wend
    '--- On lance l'impression du pdf aplati
    PrintPdf("""" & sOutputPdfName & """")
    '--- On choisi CLOE par defaut
    WScript.Sleep 2000
    Set oExec = WshShell.Exec("""" & sSelectCLOEPrinterPath & """")
    While oExec.Status = WshRunning
        WScript.Sleep 50
    Wend
    set WshShell = Nothing 
Next

' -----------------------------------------

Function GetFilenameWithoutExtension(ByVal FileName)
  Dim Result, i
  Result = FileName
  i = InStrRev(FileName, ".")
  If ( i > 0 ) Then
    Result = Mid(FileName, 1, i - 1)
  End If
  GetFilenameWithoutExtension = Result
End Function

Sub ExitIfVbsNotCompatible()
  If CDbl(Replace(WScript.Version,".",",")) < 5.1 then
    MsgBox "Vous devez utiliser ""Windows Scripting Host"" version 5.1 ou supérieure !", vbCritical + vbSystemModal, AppTitle
    Wscript.Quit
  End if
End Sub

Sub ExitIfExecutableNotInstalled(sExecutablePath)
  If Not FileExists(sExecutablePath) Then
    MsgBox "L'executable principal n'est pas correctement installé. Absence du fichier : " & sExecutablePath, vbExclamation + vbSystemModal, AppTitle
    WScript.Quit
  End If
End Sub

Sub ExitIfWrongNumberOfArgs(iNumberOfArgs)
  Dim sArgsPluriel
  sArgsPluriel = ""
  If iNumberOfArgs > 1 Then
    sArgsPluriel = "s"
  End If
  If objArgs.Count <> iNumberOfArgs Then
    MsgBox "Ce script attend " & iNumberOfArgs & " argument" & sArgsPluriel, vbExclamation + vbSystemModal, AppTitle
    WScript.Quit
  End If
End Sub

Function RegReadValue(valuePath, outValue)
    On Error Resume Next
    Err.Clear
    RegReadValue = False
    outValue = WshShell.RegRead(valuePath)
    ' msgbox "outValue:" & outValue
    If Err.Number=0 Then
      RegReadValue = True
    End If
    On Error Goto 0
End Function

Sub OpenPdf(filename, page)
    Set wshShell = WScript.CreateObject("WSCript.shell")
    wshShell.Run """" & sAcrobatReaderPath & """ /A ""page=" & page & """ " & fileName
End Sub

Sub PrintPdf(filename)
    Set wshShell = WScript.CreateObject("WSCript.shell")
    wshShell.run """" & sAcrobatReaderPath & """ /p " & fileName,,false
End Sub

Function FileExists(FilePath)
  Set fso = CreateObject("Scripting.FileSystemObject")
  If fso.FileExists(FilePath) Then
    FileExists=CBool(1)
  Else
    FileExists=CBool(0)
  End If
End Function