#include <Array.au3>
#include <File.au3>
#include <Word.au3>
#include <Excel.au3>

Global $WdApp,$XlApp


msgbox(48,"Warning","please close all Word and Excel Files before using this application")




myarr()


Func _GetFilesList (ByRef $Path, ByRef $Type)

	Local $FPath = $Path
if $Type = "" Then $Type = "*"

	if $FPath <> "" Then

		Local $FileList = _FileListToArrayRec($FPath,$Type,1,1,Default,1)

		Return $FileList
	Else
		Return 0

		EndIf

EndFunc

Func myarr()

$FolderPath = InputBox ("Folder Path","Please Insert Folder Path")
$FileType = InputBox ("File Type","Please Insert File Type like '*'"&@CRLF &"*.docx")

$TheFiles = _GetFilesList($FolderPath,$FileType)
$filesNum = UBound($TheFiles)

if $filesNum > 0 Then
$WdApp = _Word_Create(False)
$XlApp =_Excel_Open(False,False,False,False,False)
For $i = 1 To $filesNum -1
ConvertToPDF($FolderPath & "\" & $TheFiles[$i])
Next
_Word_Quit($WdApp)
_Excel_Close($XlApp)
MsgBox(64,"Done","Done Converting the files.")
Else
	MsgBox(16,"Error","Unknown Folder Path")
EndIf

EndFunc


Func ConvertToPDF($file)
$Sp = StringSplit($file,".")

if $Sp[Ubound($Sp)-1] = "docx" Then

$wdfile = _Word_DocOpen($WdApp,$file)
_Word_DocExport($wdfile,StringLeft($file,StringLen($file)-4) & "pdf")
_Word_DocClose($wdfile)

ElseIf  $Sp[Ubound($Sp)-1] = "xlsx" Then
$xlfile = _Excel_BookOpen($XlApp, $file)
_Excel_Export($XlApp, $xlfile, StringLeft($file,StringLen($file)-4) & "pdf")
_Excel_BookClose($xlfile)
EndIf
EndFunc