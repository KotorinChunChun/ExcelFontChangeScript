Option Explicit

' �Ǘ��Ҍ����m�F
Sub CheckAdmin()

    Dim Args
    Dim IsExecutedUAC '�Ǘ��Ҍ����t���O�i�p�����[�^��uac���܂ގ�True�j
    Set Args = WScript.Arguments

    dim i
    For i = 0 To Args.Count - 1
      if Args(i) = "uac" then IsExecutedUAC = true
    Next

    ' �Ǘ��Ҍ����ɏ��i
    'Dim WScript    'VBE�ł̃R�[�h�`�F�b�N�p
    Dim Param
    Do While IsExecutedUAC = false And WScript.Version >= 5.7

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

Sub SearchFiles(Filename, Folder, FoundFiles, fso)
    If Not IsObject(Folder) Or fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set Folder = fso.GetFolder(Folder)
        FoundFiles = Array()
    End If
    
    Dim CurrentFile
    Dim Filepath
    Dim Subfolder
    
    For Each Filepath In Folder.Files
        Set CurrentFile = fso.GetFile(Filepath)
        If UCase(CurrentFile.Name) = UCase(Filename) Then
            If UBound(FoundFiles) < 0 Then
                ReDim FoundFiles(1)
            Else
                ReDim Preserve FoundFiles(UBound(FoundFiles) + 1)
            End If
            FoundFiles(UBound(FoundFiles)) = Filepath
        End If
    Next
    
    For Each Subfolder in Folder.SubFolders
        On Error Resume Next
        Call SearchFiles(Filename, Subfolder, FoundFiles, fso)
        On Error Goto 0
    Next
End Sub


Function GetOffice16RootPath()
    Dim TargetFolder
    Dim FoundFiles
    Dim ExePath
    Dim RootPath
    Dim Reg: Set Reg = New RegExp
    
    GetOffice16RootPath = ""
    Reg.pattern = "\\Office16\\.*$"
    
    ' TODO:
    '   - wsh.ExpandEnvironmentStrings("%ProgramFiles(x86)%"), wsh.ExpandEnvironmentStrings("%ProgramFiles%") �Ŏ擾�����ق����x�^�[
    '   - Path �̑g�ݗ��Ă� fso.BuildPath(Folder, Filename) ���g���ق����x�^�[
    For Each TargetFolder In Array( "C:\Program Files (x86)\Microsoft Office", "C:\Program Files\Microsoft Office", "C:\Program Files\WindowsApps" )
        Set FoundFiles = Nothing
        On Error Resume Next
        Call SearchFiles("EXCEL.EXE", TargetFolder, FoundFiles, Nothing)
        On Error Goto 0
        If IsArray(FoundFiles) Then
            For Each ExePath in FoundFiles
                RootPath = Reg.Replace(ExePath, "")
                If RootPath <> ExePath Then
                    GetOffice16RootPath = RootPath
                    Exit Function
                End If
            Next
        End If
    Next
End Function


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

'    ' Excel.EXE�̑��݂���ProgramFiles�̃p�X�����
'    If fso.FileExists("C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE") Then
'        PGF = "Program Files (x86)"
'    ElseIf fso.FileExists("C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE") Then
'        PGF = "Program Files"
'    Else
'        MsgBox "�Ή����Ă���o�[�W������Excel���C���X�g�[������Ă��܂���B"
'        Exit Sub
'    End If
'
'    ' �e���v���[�g�t�@�C���ۑ��p�X
'    PATH_TEMPLATE = "C:\" & PGF & "\Microsoft Office\root\Office16\XLSTART\"
'    ' �V�K�쐬�̃t�@�C���ۑ��p�X
'    PATH_NEW = "C:\" & PGF & "\Microsoft Office\root\VFS\Windows\SHELLNEW\"
    
    ' EXCEL.EXE�̑��݂���ProgramFiles�̃p�X�����
    Dim RootPath: RootPath = GetOffice16RootPath()
    
    If RootPath = "" Then
        MsgBox "�Ή����Ă���o�[�W������Excel���C���X�g�[������Ă��܂���B��Ƃ𒆎~���܂��B"
        WScript.Quit
        Exit Sub
    End If
    
    Dim ExcelPathFile: Set ExcelPathFile = fso.OpenTextFile(gScriptPath & APP_NAME & "\" & EXCEL_PATH_FILE, 2, True)
    ExcelPathFile.WriteLine RootPath
    ExcelPathFile.Close
    
    ' �e���v���[�g�t�@�C���ۑ��p�X
    PATH_TEMPLATE = RootPath & "\Office16\XLSTART\"
    ' �V�K�쐬�̃t�@�C���ۑ��p�X
    PATH_NEW = RootPath & "\VFS\Windows\SHELLNEW\"
    
    ' ����쓮�̌����ƂȂ�댯�����邽�߁A�R�����g�A�E�g
    ' wsh.exec "takeown /F """ & RootPath & """ /R /A"
    ' wsh.exec "icacls """ & RootPath & """ /grant:r Administrators:(OI)(CI)(F) /T /C /Q"
End Sub


'----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------
' todo:

Select Case MsgBox(" �͂� �F�C���X�g�[��" & vbLf & "�������F�A���C���X�g�[��", vbYesNoCancel, "Excel ���S�V�b�N�폜�c�[��")
    Case vbYes : Call Install
    Case vbNo  : Call UnInstall
    Case Else  :  MsgBox "�L�����Z������܂����B"
End Select

Sub Install()

'#If VBA7 Then
    Call VBA_Main
'#End If

    ' Excel�̃I�v�V�����̕ύX
    wsh.RegWrite "HKCU\Software\Microsoft\Office\16.0\Excel\Options\Font", "�l�r �o�S�V�b�N,11", "REG_SZ"

    ' �e���v���[�g�^�V�K�쐬�̃t�@�C�����O����
    ' If Not fso.FolderExists(APP_PATH) Then
    '     fso.DeleteFolder Left(APP_PATH, Len(APP_PATH) - 1)
    fso.CopyFolder gScriptPath & APP_NAME, Left(APP_PATH, Len(APP_PATH) - 1), True
    ' End If

    ' �e���v���[�g�uBook.xltx�v�̎����X�V�v���O�������d����
    ' wsh.exec "schtasks /create /tn " & APP_NAME & " /tr ""wscript aaa"" /sc onlogon /rl highest /F"
    
    ' �N����(���O�C����)�Ɏ��s������@
    ' wsh.exec "schtasks /create /tn " & APP_NAME & " /tr ""wscript \""" & VBS_PATH & "\"" \""/hide\"""" /sc onlogon /rl highest /F"
    
    ' 1�������Ɏ��s������@
    ' schtasks /create /tn AUTO_BUILD /tr c:\test.vbs /sc minute /mo 1 
    wsh.exec "schtasks /create /tn " & APP_NAME & " /tr ""wscript \""" & VBS_PATH & "\"" \""/hide\"""" /sc minute /mo 1 /rl highest /F"
    
    
    ' ' /hide�̑O��WQ������Ȃ����ǉ��̂�����
    ' ' wsh.exec "schtasks /create /tn " & APP_NAME & " /tr ""wscript \""" & VBS_PATH & "\"" /hide\"""" /sc onlogon /rl highest /F"
    ' ' schtasks /create /tn AAA /tr "wscript \"BBB\" /hide\"" /sc onlogon /rl highest /F
    
    ' ' �p�����[�^�����Ȃ�OK
    ' ' wsh.exec "schtasks /create /tn " & APP_NAME & " /tr ""wscript \""" & VBS_PATH & "\"""" /sc onlogon /rl highest /F"
    ' ' schtasks /create /tn AAA /tr "wscript \"BBB\"" /sc onlogon /rl highest /F
    
    ' ' �o�b�`�t�@�C���̌��{
    ' '       schtasks /create /tn "Excel2016FontChange" /tr "wscript \"C:\Program Files (x86)\aaa.vbs\"" /sc onlogon /rl highest /F
    
    ' �Ƃ肠����������1����s���Ă���
    ' �쐬����͎��s����Ȃ��炵���̂Œx��������B
    WScript.Sleep 1000
    wsh.exec "schtasks /run /tn " & APP_NAME & ""
    
    ' �V�K�쐬�̃o�b�N�A�b�v
    If fso.FileExists(PATH_NEW & "EXCEL12.XLSX") Then
        If Not fso.FileExists(PATH_NEW & "EXCEL12_base.XLSX") Then
            fso.CopyFile PATH_NEW & "EXCEL12.XLSX", PATH_NEW & "EXCEL12_base.XLSX"
        End If
    End If

    ' �V�K�쐬�̃��W�X�g���ǉ�
    wsh.RegWrite "HKCR\.xlsm\Excel.SheetMacroEnabled.12\ShellNew\FileName", PATH_NEW & "EXCEL12.XLSM", "REG_SZ"
    ' ���L�͌��ʂȂ��B Excel.Sheet.8��xls_auto_file�̒�`���Ȃ����炾�Ǝv����
    ' REG ADD "HKEY_CLASSES_ROOT\.xls\Excel.Sheet.8\ShellNew" /v "FileName" /t REG_SZ /d "C:\Program Files (x86)\Microsoft Office\Root\VFS\Windows\ShellNew\EXCEL8.XLS" /f
    wsh.RegWrite "HKCR\.xls\ShellNew\FileName", PATH_NEW & "EXCEL8.XLS", "REG_SZ"
    ' [�t�@�C���̎��]�̒�`�B�R�R���󗓂���ShellNew�ɓo�^���Ă����j���[�ɑ����Ȃ��B
    wsh.RegWrite "HKCU\Software\Classes\xls_auto_file", "Microsoft Excel 97-2003 �݊��u�b�N", "REG_SZ"
    ' ����l��\�ŏI��
    wsh.RegWrite "HKCR\xls_auto_file\", "Microsoft Excel 97-2003 �݊��u�b�N", "REG_SZ"

    MsgBox "����", vbOKOnly, "�C���X�g�[��"
    
End Sub

Sub Uninstall()

'#If VBA7 Then
    Call VBA_Main
'#End If
    If fso.FolderExists(APP_PATH) Then
        fso.DeleteFolder Left(APP_PATH, Len(APP_PATH) - 1)
    End If

    wsh.exec "schtasks /delete /tn """ & APP_NAME & """ /F"

    ' Excel�̃I�v�V�����̃t�H���g�ύX
    wsh.RegWrite "HKCU\Software\Microsoft\Office\16.0\Excel\Options\Font", "���S�V�b�N,11", "REG_SZ"

    ' �e���v���[�g�uBook.xltx�v�̍폜
    If fso.FileExists(PATH_TEMPLATE & FILE_XLTX) Then
        fso.DeleteFile PATH_TEMPLATE & FILE_XLTX
    End If

    ' �V�K�쐬�̃t�@�C���폜
    If fso.FolderExists(PATH_NEW) Then
        On Error Resume Next
        fso.DeleteFile PATH_NEW & "EXCEL8.XLS"
        fso.DeleteFile PATH_NEW & "EXCEL12.XLSM"
        fso.DeleteFile PATH_NEW & "EXCEL12.XLSX"
        ' fso.CopyFile PATH_NEW & "EXCEL12_base.XLSX", PATH_NEW & "EXCEL12.XLSX", True
        fso.MoveFile PATH_NEW & "EXCEL12_base.XLSX", PATH_NEW & "EXCEL12.XLSX"
        On Error GoTo 0
    End If
    
    ' �V�K�쐬�̃��W�X�g���폜
    On Error Resume Next
    wsh.RegDelete "HKCR\.xlsm\Excel.SheetMacroEnabled.12\ShellNew\FileName"
    ' Excel.Sheet.8�͕s�v�Ȃ͂������ꉞ
    wsh.RegDelete "HKCR\.xls\Excel.Sheet.8\ShellNew\FileName"
    wsh.RegDelete "HKCR\.xls\ShellNew\FileName"
    On Error GoTo 0
    
    MsgBox "����", vbOKOnly, "�A���C���X�g�[��"

End Sub
