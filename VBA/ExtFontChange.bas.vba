Attribute VB_Name = "ExtFontChange"
Rem --------------------------------------------------------------------------------
Rem
Rem  @module        ExtFontChange
Rem
Rem  @description   Excelブックのフォントを統一するマクロ
Rem
Rem  @update        2020/08/07
Rem
Rem  @author        @KotorinChunChun (GitHub / Twitter)
Rem
Rem  @license       MIT (http://www.opensource.org/licenses/mit-license.php)
Rem
Rem --------------------------------------------------------------------------------
Rem  @references
Rem    Microsoft Scripting Runtime
Rem
Rem --------------------------------------------------------------------------------
Rem  @history
Rem     2019/07/21 : ブログ掲載
Rem     2020/08/23 : GitHub掲載
Rem
Rem --------------------------------------------------------------------------------
Rem  @note
Rem   [公開先]
Rem     えくせるちゅんちゅん - Excelから游ゴシック体を徹底的に駆逐する Part3
Rem      https://www.excel-chunchun.com/entry/FontChange3
Rem
Rem   [対応しているフォント]
Rem     ・游ゴシック
Rem     ・游明朝
Rem     ・ＭＳ Ｐゴシック
Rem     ・ＭＳ Ｐ明朝
Rem     ・メイリオ　　　※メイリオは列幅も変動するため図形変形対策非対応
Rem
Rem --------------------------------------------------------------------------------

Option Explicit

Rem これをリボンに登録したりF8で実行
Public Sub アクティブブックのフォントをMSPゴシックに統一()
    Call ブックのフォントを指定フォントに変更(ActiveWorkbook, "ＭＳ Ｐゴシック", "ＭＳ Ｐ明朝")
End Sub

Public Sub アクティブブックのフォントを游ゴシックに統一()
    Call ブックのフォントを指定フォントに変更(ActiveWorkbook, "游ゴシック", "游明朝")
End Sub

Public Sub アクティブブックのフォントをメイリオに統一()
'    Call ブックのフォントを指定フォントに変更(ActiveWorkbook, "メイリオ", "メイリオ")
    Call ブックのフォントを指定フォントに変更(ActiveWorkbook, "Meiryo UI", "Meiryo UI")
End Sub

Rem 変換マクロ本体
Rem 基本は游ゴシック→AfterGothic、游明朝→AfterMinchoだが、
Rem AfterGothicが既定のフォントとなる
Public Sub ブックのフォントを指定フォントに変更(wb As Workbook, AfterGothic As String, AfterMincho As String)
    
    'フォント変換テーブルを作成
    Dim fonts As Dictionary: Set fonts = New Dictionary
    Dim item As Variant
    
    '同一のフォント名が複数存在することがある為、全通り登録
    For Each item In GetFonts("游ゴシック*", "YuGothic*", "Yu Gothic*", "ＭＳ Ｐゴシック*", "Meiryo*", "メイリオ*")
        If Not fonts.Exists(item) Then
            fonts.Add item, AfterGothic
        End If
    Next
    For Each item In GetFonts("游明朝*", "YuMincho*", "Yu Mincho*", "ＭＳ Ｐ明朝*")
        If Not fonts.Exists(item) Then
            fonts.Add item, AfterMincho
        End If
    Next
    
    'フォントの変換処理を実行
    Dim ws As Worksheet
    
    For Each ws In wb.Worksheets
        Call 行の高さを固定に変更(ws)
    Next
    
    Call セルのスタイルのフォントを変更(wb, fonts)
    Call Officeテーマフォントを変更(wb, "" & fonts.Items(1))
    
    For Each ws In wb.Worksheets
        Call セルのフォントを変更(ws, fonts)
        Call 図のフォントを変更(ws, fonts)
        Call ヘッダフッタのフォントを変更(ws, fonts)
    Next
    
End Sub

'指定した単語から始まるフォントをリスト化
Public Function GetFonts(ParamArray FontNameFiltes() As Variant) As Collection
    Set GetFonts = New Collection
    
    If TypeName(Excel.Selection) <> "Range" Then ActiveWindow.RangeSelection.Select
    
    With Application.CommandBars("Formatting").Controls(1)
        Dim i As Long
        For i = 1 To .ListCount
            Dim fnf As Variant
            For Each fnf In FontNameFiltes
                If .List(i) Like fnf Then
'                    Debug.Print .List(i)
                    GetFonts.Add .List(i)
                End If
            Next
        Next
    End With
End Function

Sub Test_GetFonts()
    Dim item
    For Each item In GetFonts("游ゴシック*", "YuGothic*", "Yu Gothic*")
        Debug.Print item
    Next
End Sub

Public Sub 行の高さを固定に変更(ws As Worksheet)
    With ws
        Dim rng As Range
        For Each rng In .Range(.Cells(1, 1), .UsedRange).EntireRow
            rng.RowHeight = rng.RowHeight
        Next
    End With
End Sub

Public Sub セルのスタイルのフォントを変更(wb As Workbook, fonts As Dictionary)
    Dim st As Style
    For Each st In wb.Styles
        With st.Font
            Dim fts
            For Each fts In fonts.Keys
                If .Name = fts Then .Name = fonts(fts)
            Next
        End With
    Next
End Sub

Sub Test_Officeテーマフォントを変更()
    Call Officeテーマフォントを変更(ActiveWorkbook, "MSPゴシック")
    Call Officeテーマフォントを変更(ActiveWorkbook, "MSP明朝")
    Call Officeテーマフォントを変更(ActiveWorkbook, "MS Pゴシック")
    Call Officeテーマフォントを変更(ActiveWorkbook, "MS P ゴシック")
    Call Officeテーマフォントを変更(ActiveWorkbook, "游ゴシック")
End Sub

Public Sub Officeテーマフォントを変更(wb As Workbook, select_font As String)
    '※パスはExcel 2016/Office365/32bit/64bitでしか検証していない。
    Const FP1 = "C:\Program Files\Microsoft Office\Root\Document Themes 16\"
    Const FP2 = "C:\Program Files (x86)\Microsoft Office\Root\Document Themes 16\"
    
    Dim fnt As String
    fnt = select_font
    fnt = Replace(Replace(Replace(fnt, "游", "yu"), "明朝", "mincho"), "ゴシック", "gothic")
    fnt = Replace(fnt, "メイリオ", "meiryo")
    fnt = Replace(StrConv(fnt, vbLowerCase + vbNarrow), " ", "")
    
    Dim fp As String
    If fnt Like "yugothic" Then
        fp = "Office Theme.thmx"
        fp = ChooseExistsFile(FP1 & fp, FP2 & fp)
        Debug.Print fp
        Call wb.ApplyTheme(fp)
    Else
        Select Case True
            Case fnt Like "mspgothic*": fp = "Office 2007 - 2010.xml"
            Case fnt Like "mspmincho*": fp = "Century Schoolbook.xml"
            Case fnt Like "meiryo*": fp = "Calibri.xml"
        End Select
        fp = "Theme Fonts\" & fp
        fp = ChooseExistsFile(FP1 & fp, FP2 & fp)
        Debug.Print fp
        Call wb.Theme.ThemeFontScheme.Load(fp)
    End If
End Sub

Rem 指定したファイルパス一覧から最初にファイルが実在したパスを返す
Rem
Rem @param arr_filepath ファイルのフルパスの可変長引数
Rem
Rem @return As String   ファイルパス
Rem                     存在しない場合は空欄""を返す
Public Function ChooseExistsFile(ParamArray arr_filepath()) As String
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim fp
    For Each fp In arr_filepath
        If fso.FileExists(fp) Then
            ChooseExistsFile = fp
            Exit Function
        End If
    Next
    Stop
End Function

Public Sub セルのフォントを変更(ws As Worksheet, fonts As Dictionary)
    Dim rng As Range
    For Each rng In ws.UsedRange
        With rng.Font
            Dim fts
            For Each fts In fonts.Keys
                If .Name = fts Then .Name = fonts(fts)
            Next
        End With
    Next
End Sub

Sub Test_セルのフォントを変更()
    Dim fonts As Dictionary: Set fonts = New Dictionary
    Dim item
    For Each item In GetFonts("游ゴシック*", "YuGothic*", "Yu Gothic*")
        If Not fonts.Exists(item) Then
            fonts.Add item, "ＭＳ Ｐゴシック"
        End If
    Next
    Call セルのフォントを変更(ActiveSheet, fonts)
End Sub

Public Sub 図のフォントを変更(ws As Worksheet, fonts As Dictionary)
    Dim shp As Shape
    For Each shp In ws.Shapes
        '※TextFrame2はExcel 2007以降
        If shp.TextFrame2.HasText Then
            With shp.TextFrame2.TextRange.Font
                Dim fts
                For Each fts In fonts.Keys
                    If .Name = fts Then .Name = fonts(fts)
                    If .NameFarEast = fts Then .NameFarEast = fonts(fts)
                Next
            End With
        End If
    Next
End Sub

Public Sub ヘッダフッタのフォントを変更(ws As Worksheet, fonts As Dictionary)
    Dim ps As PageSetup: Set ps = ws.PageSetup
    With ps
        Dim fts
        For Each fts In fonts.Keys
            'Replaceだと游ゴシックというヘッダフッタ文字列を置換してしまうバグあり
            If HeaderFooterFont(.LeftHeader) = fts Then
                .LeftFooter = Replace(.LeftFooter, fts, fonts(fts))
            End If
            If HeaderFooterFont(.CenterHeader) = fts Then
                .CenterHeader = Replace(.CenterHeader, fts, fonts(fts))
            End If
            If HeaderFooterFont(.RightHeader) = fts Then
                .RightHeader = Replace(.RightHeader, fts, fonts(fts))
            End If
            
            If HeaderFooterFont(.LeftFooter) = fts Then
                .LeftFooter = Replace(.LeftFooter, fts, fonts(fts))
            End If
            If HeaderFooterFont(.CenterFooter) = fts Then
                .CenterFooter = Replace(.CenterFooter, fts, fonts(fts))
            End If
            If HeaderFooterFont(.RightFooter) = fts Then
                .RightFooter = Replace(.RightFooter, fts, fonts(fts))
            End If
        Next
    End With
End Sub

Rem ヘッダ・フッタのフォント文字列を返す
Rem 文字データのみの時　　：　bbbb
Rem フォント設定時　　　　：　&"HGP創英角ﾎﾟｯﾌﾟ体,標準"dddd
Rem フォント、サイズ設定時：　&"游ゴシック,標準"&16ああああ
Rem フォント、太字設定時　：　&"Yu Gothic UI Light,太字"ああああ
Rem フォント、斜体設定時　：　&"-,斜体"ああああ
Rem フォント、下線設定示　：　&"-,標準"&Uああああ
Rem フォント、取消線設定時：　&"-,標準"&Sああああ
Rem 太字・斜体・下線設定時：　&"-,太字 斜体"&Uああああ
Rem 登録して消した通常時　：　&"+,標準"ああああ
Public Function HeaderFooterFont(hfValue As String) As String
    Dim S() As String
    S = Split(hfValue, """")
    If UBound(S, 1) > 0 Then
        HeaderFooterFont = Split(S(1), ",")(0)
    End If
End Function

Public Sub Test_HeaderFooterFont()
    Debug.Print HeaderFooterFont("")
    Debug.Print HeaderFooterFont("&""-,斜体""&Uああああ")
    Debug.Print HeaderFooterFont("&""游ゴシック,標準""&16ああああ")
    Debug.Print HeaderFooterFont("&""-,太字 斜体""&Uああああ")
    Debug.Print HeaderFooterFont("&""Yu Gothic UI Light,太字""ああああ")
End Sub

'これはブックではなくアプリの設定を書き換えるため対象外
'Public Sub Excelアプリケーションの標準フォントを変更()
'    Application.StandardFont = "Meiryo UI"
'    Application.StandardFontSize = 11
'End Sub
