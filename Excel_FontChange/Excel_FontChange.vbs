Option Explicit

' �Ǘ��Ҍ����m�F
Sub CheckAdmin()

    Dim Args
    Dim UacFlag '�Ǘ��Ҍ����t���O�i�p�����[�^��uac���܂ގ�True�j
    Set Args = WScript.Arguments

    dim i
    For i = 0 To Args.Count - 1
      if Args(i) = "uac" then UacFlag = true
    Next

    ' �Ǘ��Ҍ����ɏ��i
    'Dim WScript    'VBE�ł̃R�[�h�`�F�b�N�p
    Dim Param
    Do While UacFlag = false And WScript.Version >= 5.7

      '���݂̃p�����[�^���X�y�[�X��؂�ɕϊ�
      Param = ""
      For i = 0 To Args.Count - 1
        Param = Param & " " & Args(i)
      Next
      
      ' Check WScript5.7~ and Vista~
      Dim os, wmi, Value
      Set wmi = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\.\root\cimv2")
      Set os = wmi.ExecQuery("SELECT *FROM Win32_OperatingSystem")
      For Each Value In os
        If Left(Value.Version, 3) < 6.0 Then Exit Do   'Exit if not vista
      Next
       
      ' Run this script as admin.
      Dim sha
      Set sha = CreateObject("Shell.Application")
      sha.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ uac" & Param, "", "runas"
       
      WScript.Quit
    Loop

End Sub

Call CheckAdmin

'----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------
'�����ݒ�

Dim fso
Dim wsh

Dim gScriptName
Dim gScriptFullName
Dim gScriptPath
Dim PATH_TEMPLATE
Dim PATH_NEW

Const APP_NAME = "Excel_FontChange"
Const APP_PATH = "C:\Program Files\Excel_FontChange\"
Const VBS_PATH = "C:\Program Files\Excel_FontChange\Excel_FontChange.vbs"
Const EXCEL_PATH_FILE = "ExcelPath.txt"

Const FILE_XLTX = "Book.xltx"

Sub VBA_Main()
    Dim PGF

'#If VBA7 Then
'#Else
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wsh = CreateObject("WScript.Shell")
    
    gScriptName = WScript.ScriptName
    gScriptFullName = WScript.ScriptFullName
    gScriptPath = fso.GetParentFolderName(gScriptFullName) & "\"
'#End If

'   ' Excel.EXE�̑��݂���ProgramFiles�̃p�X�����
'   If fso.FileExists("C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE") Then
'       PGF = "Program Files (x86)"
'   ElseIf fso.FileExists("C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE") Then
'       PGF = "Program Files"
'   Else
'       MsgBox "�Ή����Ă���o�[�W������Excel���C���X�g�[������Ă��܂���B"
'       Exit Sub
'   End If
'
'   ' �e���v���[�g�t�@�C���ۑ��p�X
'   PATH_TEMPLATE = "C:\" & PGF & "\Microsoft Office\root\Office16\XLSTART\"
'   ' �V�K�쐬�̃t�@�C���ۑ��p�X
'   PATH_NEW = "C:\" & PGF & "\Microsoft Office\root\VFS\Windows\SHELLNEW\"
    
    If Not fso.FileExists(APP_PATH & EXCEL_PATH_FILE) Then
        MsgBox "Excel Path�t�@�C����������܂���B"
        WScript.Quit
        Exit Sub
    End If
    
    Dim ExcelPathFile: Set ExcelPathFile = fso.OpenTextFile(APP_PATH & EXCEL_PATH_FILE, 1, False)
    Dim RootPath: RootPath = Trim( ExcelPathFile.ReadLine )
    ExcelPathFile.Close
    
    ' �e���v���[�g�t�@�C���ۑ��p�X
    PATH_TEMPLATE = RootPath & "\Office16\XLSTART\"
    ' �V�K�쐬�̃t�@�C���ۑ��p�X
    PATH_NEW = RootPath & "\VFS\Windows\SHELLNEW\"
End Sub

Call VBA_Main

'----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------
' todo:

Dim CurPath
CurPath = gScriptPath

'�e���v���[�g���R�s�[
Dim XlStartPath
XlStartPath = PATH_TEMPLATE

fso.CopyFile gScriptPath & FILE_XLTX, XlStartPath & FILE_XLTX, True

'�V�K�쐬�t�H���_�ɃR�s�[
Dim ShellNewPath
ShellNewPath = PATH_NEW

Dim NewFiles
NewFiles = Array("EXCEL8.XLS", "EXCEL12.XLSM", "EXCEL12.XLSX")

Dim file
For Each file In NewFiles
    fso.CopyFile CurPath & file, ShellNewPath & file, True
Next

'�㏈��
Dim i
Dim IsHide
Dim Args

Set Args = WScript.Arguments
For i = 0 To Args.Count - 1
    If Args(i) = "/hide" Then IsHide = True
Next

If IsHide = False Then
    MsgBox "����",, "Excel �t�H���g�����ύX"
End If
