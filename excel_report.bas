Attribute VB_Name = "excel_report"
'---------------------------------------------------------------------------------------
' Module    : Excel関連処理
' Purpose   : Excelに関わる処理をまとめたモジュール
'             Microsoft Excel xx Object Library への参照設定が必要
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


'共通変数の定義
Private xlapp As Excel.Application
'Public xlsBook As Excel.Workbooks
Private wb As Excel.Workbook
Private ws As Excel.Worksheet

Private myCn As New ADODB.Connection
'Public myCn As ADODB.Connection
Private myRs As New ADODB.Recordset

'---------------------------------------------------------------------------------------
' Exceを起動
'---------------------------------------------------------------------------------------
Private Function xlappOpen()
    
'    Set xlapp = CreateObject("Excel.Application")
    Set xlapp = New Excel.Application
    xlapp.UserControl = True

End Function

'---------------------------------------------------------------------------------------
' Excelを終了
'---------------------------------------------------------------------------------------
Private Function xlappClose()

    xlapp.Quit: Set xlapp = Nothing

End Function

'---------------------------------------------------------------------------------------
' ADO Recordsetを作成
'---------------------------------------------------------------------------------------
Private Function Create_ADO_Recordset(SQL As String)
    Set myCn = CurrentProject.Connection
    myRs.Open SQL, myCn, adOpenStatic, adLockReadOnly
End Function

'    Set myCn = CurrentProject.Connection
'    Dim myRs            As New ADODB.Recordset
'    myRs.Open QERY, myCn, adOpenForwardOnly, adLockReadOnly

'---------------------------------------------------------------------------------------
' ADO Recordsetを終了
'---------------------------------------------------------------------------------------
Private Function ADO_Close() As Boolean
    ADO_Close = False
    Set myRs = Nothing: Close
    Set myCn = Nothing: Close
    ADO_Close = True
End Function


Sub test_判定_会計年度()

    Dim TestDate As String

Debug.Print "test_cmn判定_会計年度"
Debug.Print "設定値　:戻り値"

    TestDate = "20190331": GoSub Result
    TestDate = "20190401": GoSub Result
    TestDate = "20190930": GoSub Result
    TestDate = "20191001": GoSub Result
    TestDate = "20200331": GoSub Result
    TestDate = "20200401": GoSub Result

    Exit Sub
    
Result:
    Debug.Print TestDate & ":" & cmn判定_会計年度(CDate(Format(TestDate, "0000/00/00")))
Return

End Sub

'---------------------------------------------------------------------------------------
Public Function cmn判定_会計年度(dtmDate As Date) As String

    cmn判定_会計年度 = CStr(Year(DateAdd("m", -3, dtmDate)))
    
End Function

'---------------------------------------------------------------------------------------
Sub test_判定_上期下期()

    Dim TestDate As String

Debug.Print "test_cmn判定_上期下期"
Debug.Print "設定値　:戻り値"

    TestDate = "20190331": GoSub Result
    TestDate = "20190401": GoSub Result
    TestDate = "20190930": GoSub Result
    TestDate = "20191001": GoSub Result
    TestDate = "20200331": GoSub Result
    TestDate = "20200401": GoSub Result
    
    Exit Sub
    
Result:
    Debug.Print TestDate & ":" & cmn判定_上期下期(CDate(Format(TestDate, "0000/00/00")))
Return

End Sub

'---------------------------------------------------------------------------------------
'上期下期
'---------------------------------------------------------------------------------------
Public Function cmn判定_上期下期(dtmDate As Date) As String

    Select Case Month(dtmDate)
                    
        Case 4, 5, 6, 7, 8, 9
            cmn判定_上期下期 = "上期"
        Case 10, 11, 12, 1, 2, 3
            cmn判定_上期下期 = "下期"
    
    End Select

End Function

'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
Sub main()

    Const 帳票名 = "hoge帳票"
    Const QERY As String = "クエリ1"
    Dim strSQL          As String
    strSQL = "SELECT * FROM " & QERY & ";"
    
    Call 汎用_レコードセットのセル転記(帳票名, 1, strSQL)
    
End Sub

Private Sub 汎用_レコードセットのセル転記(帳票名 As String, ヘッダ行 As Long, strSQL As String)

    '時間計測用
    Dim StartTime As Single
    StartTime = Timer

    '// Exceを起動
    Call xlappOpen

    Call Create_ADO_Recordset(strSQL)

    With xlapp
        .Visible = True
'        .UserControl = True

        Set wb = .Workbooks.Add(Template:=xlWBATWorksheet) ' ワークブックを作成
        Dim ws As Worksheet
        Set ws = wb.Worksheets(1)
        ws.Name = 帳票名 & "_" & Format(date, "yyyymmdd")
        
        'ヘッダ行作成
        Dim i As Integer 'myRs.Fields.Count
        For i = 0 To myRs.Fields.Count - 1
            ws.Cells(ヘッダ行, i + 1) = myRs.Fields(i).Name
        Next

        'レコードセットのセル転記
        Dim FieldsCount As Long: FieldsCount = myRs.Fields.Count - 1
        Dim tableArray As Variant
        ReDim tableArray(myRs.RecordCount - 1, FieldsCount)
        Dim j As Long: j = 0 'myRs.RecordCount
        Do Until myRs.EOF
            For i = 0 To FieldsCount
                tableArray(j, i) = myRs(i).Value
            Next
            j = j + 1: myRs.MoveNext
        Loop
        ws.Range(ws.Cells(ヘッダ行 + 1, 1), ws.Cells(myRs.RecordCount + 1, FieldsCount)).Value = tableArray

        Call 全体罫線(ws)
        Call 帳票仕上げ(ws, 帳票名)
        
    End With


    Call ADO_Close
Debug.Print Timer - StartTime

    Call xlappClose


End Sub

'単純表向けの全体罫線処理
Private Sub 全体罫線(ws As Worksheet)

    With ws.Range("A1").CurrentRegion
    
        .Columns.AutoFit
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
    
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    
    End With

End Sub

Private Sub 帳票仕上げ(ws As Worksheet, 帳票名 As String)

    'ヘッダ行色付け
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column)).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With

    '印刷ヘッダ
    With ws.PageSetup
'            .LeftHeader = "left"
        .CenterHeader = 帳票名
        .RightHeader = Format(date, "yyyy/m/d")
        .CenterFooter = "&P/&N"
    End With

End Sub


'---------------------------------------------------------------------------------------
'列項目中の行で、連続する文言が先頭行を除きブランク埋めする処理
'   arg1:対象のワークシート
'   arg2:対象列
'   arg3:開始行
'---------------------------------------------------------------------------------------
Private Sub fnc重複消し(ws As Worksheet, 対象列 As Long, 開始行 As Long)

    Dim xlLastRow   As Long                             'Excel自体の最終行
    Dim LastRow     As Long                             '最終行
 
    xlLastRow = Cells(Rows.Count, 1).row
    LastRow = Cells(xlLastRow, 対象列).End(xlUp).row

    Dim tmp As String: tmp = ""
    Dim i As Long
    For i = 開始行 To LastRow
        If tmp = ws.Cells(i, 対象列).Value Then
            ws.Cells(i, 対象列).Value = ""
        Else
            tmp = ws.Cells(i, 対象列).Value
        End If
    Next

End Sub










