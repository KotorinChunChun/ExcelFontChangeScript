Option Explicit

' 管理者権限確認
Sub CheckAdmin()

    Dim Args
    Dim IsExecutedUAC '管理者権限フラグ（パラメータにuacを含む時True）
    Set Args = WScript.Arguments

    dim i
    For i = 0 To Args.Count - 1
      if Args(i) = "uac" then IsExecutedUAC = true
    Next

    ' 管理者権限に昇格
    'Dim WScript    'VBEでのコードチェック用
    Dim Param
    Do While IsExecutedUAC = false And WScript.Version >= 5.7

      '現在のパラメータをスペース区切りに変換
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
'初期設定

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
    '   - wsh.ExpandEnvironmentStrings("%ProgramFiles(x86)%"), wsh.ExpandEnvironmentStrings("%ProgramFiles%") で取得したほうがベター
    '   - Path の組み立ても fso.BuildPath(Folder, Filename) を使うほうがベター
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

'    ' Excel.EXEの存在からProgramFilesのパスを特定
'    If fso.FileExists("C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE") Then
'        PGF = "Program Files (x86)"
'    ElseIf fso.FileExists("C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE") Then
'        PGF = "Program Files"
'    Else
'        MsgBox "対応しているバージョンのExcelがインストールされていません。"
'        Exit Sub
'    End If
'
'    ' テンプレートファイル保存パス
'    PATH_TEMPLATE = "C:\" & PGF & "\Microsoft Office\root\Office16\XLSTART\"
'    ' 新規作成のファイル保存パス
'    PATH_NEW = "C:\" & PGF & "\Microsoft Office\root\VFS\Windows\SHELLNEW\"
    
    ' EXCEL.EXEの存在からProgramFilesのパスを特定
    Dim RootPath: RootPath = GetOffice16RootPath()
    
    If RootPath = "" Then
        MsgBox "対応しているバージョンのExcelがインストールされていません。作業を中止します。"
        WScript.Quit
        Exit Sub
    End If
    
    Dim ExcelPathFile: Set ExcelPathFile = fso.OpenTextFile(gScriptPath & APP_NAME & "\" & EXCEL_PATH_FILE, 2, True)
    ExcelPathFile.WriteLine RootPath
    ExcelPathFile.Close
    
    ' テンプレートファイル保存パス
    PATH_TEMPLATE = RootPath & "\Office16\XLSTART\"
    ' 新規作成のファイル保存パス
    PATH_NEW = RootPath & "\VFS\Windows\SHELLNEW\"
    
    ' ↓誤作動の原因となる危険があるため、コメントアウト
    ' wsh.exec "takeown /F """ & RootPath & """ /R /A"
    ' wsh.exec "icacls """ & RootPath & """ /grant:r Administrators:(OI)(CI)(F) /T /C /Q"
End Sub


'----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------
' todo:

Select Case MsgBox(" はい ：インストール" & vbLf & "いいえ：アンインストール", vbYesNoCancel, "Excel 游ゴシック削除ツール")
    Case vbYes : Call Install
    Case vbNo  : Call UnInstall
    Case Else  :  MsgBox "キャンセルされました。"
End Select

Sub Install()

'#If VBA7 Then
    Call VBA_Main
'#End If

    ' Excelのオプションの変更
    wsh.RegWrite "HKCU\Software\Microsoft\Office\16.0\Excel\Options\Font", "ＭＳ Ｐゴシック,11", "REG_SZ"

    ' テンプレート／新規作成のファイル事前準備
    ' If Not fso.FolderExists(APP_PATH) Then
    '     fso.DeleteFolder Left(APP_PATH, Len(APP_PATH) - 1)
    fso.CopyFolder gScriptPath & APP_NAME, Left(APP_PATH, Len(APP_PATH) - 1), True
    ' End If

    ' テンプレート「Book.xltx」の自動更新プログラムを仕込む
    ' wsh.exec "schtasks /create /tn " & APP_NAME & " /tr ""wscript aaa"" /sc onlogon /rl highest /F"
    
    ' 起動時(ログイン時)に実行する方法
    ' wsh.exec "schtasks /create /tn " & APP_NAME & " /tr ""wscript \""" & VBS_PATH & "\"" \""/hide\"""" /sc onlogon /rl highest /F"
    
    ' 1分おきに実行する方法
    ' schtasks /create /tn AUTO_BUILD /tr c:\test.vbs /sc minute /mo 1 
    wsh.exec "schtasks /create /tn " & APP_NAME & " /tr ""wscript \""" & VBS_PATH & "\"" \""/hide\"""" /sc minute /mo 1 /rl highest /F"
    
    
    ' ' /hideの前にWQが足りないけど何故か動く
    ' ' wsh.exec "schtasks /create /tn " & APP_NAME & " /tr ""wscript \""" & VBS_PATH & "\"" /hide\"""" /sc onlogon /rl highest /F"
    ' ' schtasks /create /tn AAA /tr "wscript \"BBB\" /hide\"" /sc onlogon /rl highest /F
    
    ' ' パラメータ無しならOK
    ' ' wsh.exec "schtasks /create /tn " & APP_NAME & " /tr ""wscript \""" & VBS_PATH & "\"""" /sc onlogon /rl highest /F"
    ' ' schtasks /create /tn AAA /tr "wscript \"BBB\"" /sc onlogon /rl highest /F
    
    ' ' バッチファイルの見本
    ' '       schtasks /create /tn "Excel2016FontChange" /tr "wscript \"C:\Program Files (x86)\aaa.vbs\"" /sc onlogon /rl highest /F
    
    ' とりあえず今すぐ1回実行しておく
    ' 作成直後は実行されないらしいので遅延させる。
    WScript.Sleep 1000
    wsh.exec "schtasks /run /tn " & APP_NAME & ""
    
    ' 新規作成のバックアップ
    If fso.FileExists(PATH_NEW & "EXCEL12.XLSX") Then
        If Not fso.FileExists(PATH_NEW & "EXCEL12_base.XLSX") Then
            fso.CopyFile PATH_NEW & "EXCEL12.XLSX", PATH_NEW & "EXCEL12_base.XLSX"
        End If
    End If

    ' 新規作成のレジストリ追加
    wsh.RegWrite "HKCR\.xlsm\Excel.SheetMacroEnabled.12\ShellNew\FileName", PATH_NEW & "EXCEL12.XLSM", "REG_SZ"
    ' 下記は効果なし。 Excel.Sheet.8はxls_auto_fileの定義がないからだと思われる
    ' REG ADD "HKEY_CLASSES_ROOT\.xls\Excel.Sheet.8\ShellNew" /v "FileName" /t REG_SZ /d "C:\Program Files (x86)\Microsoft Office\Root\VFS\Windows\ShellNew\EXCEL8.XLS" /f
    wsh.RegWrite "HKCR\.xls\ShellNew\FileName", PATH_NEW & "EXCEL8.XLS", "REG_SZ"
    ' [ファイルの種類]の定義。ココが空欄だとShellNewに登録してもメニューに増えない。
    wsh.RegWrite "HKCU\Software\Classes\xls_auto_file", "Microsoft Excel 97-2003 互換ブック", "REG_SZ"
    ' 既定値は\で終了
    wsh.RegWrite "HKCR\xls_auto_file\", "Microsoft Excel 97-2003 互換ブック", "REG_SZ"

    MsgBox "完了", vbOKOnly, "インストール"
    
End Sub

Sub Uninstall()

'#If VBA7 Then
    Call VBA_Main
'#End If
    If fso.FolderExists(APP_PATH) Then
        fso.DeleteFolder Left(APP_PATH, Len(APP_PATH) - 1)
    End If

    wsh.exec "schtasks /delete /tn """ & APP_NAME & """ /F"

    ' Excelのオプションのフォント変更
    wsh.RegWrite "HKCU\Software\Microsoft\Office\16.0\Excel\Options\Font", "游ゴシック,11", "REG_SZ"

    ' テンプレート「Book.xltx」の削除
    If fso.FileExists(PATH_TEMPLATE & FILE_XLTX) Then
        fso.DeleteFile PATH_TEMPLATE & FILE_XLTX
    End If

    ' 新規作成のファイル削除
    If fso.FolderExists(PATH_NEW) Then
        On Error Resume Next
        fso.DeleteFile PATH_NEW & "EXCEL8.XLS"
        fso.DeleteFile PATH_NEW & "EXCEL12.XLSM"
        fso.DeleteFile PATH_NEW & "EXCEL12.XLSX"
        ' fso.CopyFile PATH_NEW & "EXCEL12_base.XLSX", PATH_NEW & "EXCEL12.XLSX", True
        fso.MoveFile PATH_NEW & "EXCEL12_base.XLSX", PATH_NEW & "EXCEL12.XLSX"
        On Error GoTo 0
    End If
    
    ' 新規作成のレジストリ削除
    On Error Resume Next
    wsh.RegDelete "HKCR\.xlsm\Excel.SheetMacroEnabled.12\ShellNew\FileName"
    ' Excel.Sheet.8は不要なはずだが一応
    wsh.RegDelete "HKCR\.xls\Excel.Sheet.8\ShellNew\FileName"
    wsh.RegDelete "HKCR\.xls\ShellNew\FileName"
    On Error GoTo 0
    
    MsgBox "完了", vbOKOnly, "アンインストール"

End Sub
