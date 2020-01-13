Attribute VB_Name = "excel_report"
'---------------------------------------------------------------------------------------
' Module    : Excel�֘A����
' Purpose   : Excel�Ɋւ�鏈�����܂Ƃ߂����W���[��
'             Microsoft Excel xx Object Library �ւ̎Q�Ɛݒ肪�K�v
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


'���ʕϐ��̒�`
Private xlapp As Excel.Application
'Public xlsBook As Excel.Workbooks
Private wb As Excel.Workbook
Private ws As Excel.Worksheet

Private myCn As New ADODB.Connection
'Public myCn As ADODB.Connection
Private myRs As New ADODB.Recordset

'---------------------------------------------------------------------------------------
' Exce���N��
'---------------------------------------------------------------------------------------
Private Function xlappOpen()
    
'    Set xlapp = CreateObject("Excel.Application")
    Set xlapp = New Excel.Application
    xlapp.UserControl = True

End Function

'---------------------------------------------------------------------------------------
' Excel���I��
'---------------------------------------------------------------------------------------
Private Function xlappClose()

    xlapp.Quit: Set xlapp = Nothing

End Function

'---------------------------------------------------------------------------------------
' ADO Recordset���쐬
'---------------------------------------------------------------------------------------
Private Function Create_ADO_Recordset(SQL As String)
    Set myCn = CurrentProject.Connection
    myRs.Open SQL, myCn, adOpenStatic, adLockReadOnly
End Function

'    Set myCn = CurrentProject.Connection
'    Dim myRs            As New ADODB.Recordset
'    myRs.Open QERY, myCn, adOpenForwardOnly, adLockReadOnly

'---------------------------------------------------------------------------------------
' ADO Recordset���I��
'---------------------------------------------------------------------------------------
Private Function ADO_Close() As Boolean
    ADO_Close = False
    Set myRs = Nothing: Close
    Set myCn = Nothing: Close
    ADO_Close = True
End Function


Sub test_����_��v�N�x()

    Dim TestDate As String

Debug.Print "test_cmn����_��v�N�x"
Debug.Print "�ݒ�l�@:�߂�l"

    TestDate = "20190331": GoSub Result
    TestDate = "20190401": GoSub Result
    TestDate = "20190930": GoSub Result
    TestDate = "20191001": GoSub Result
    TestDate = "20200331": GoSub Result
    TestDate = "20200401": GoSub Result

    Exit Sub
    
Result:
    Debug.Print TestDate & ":" & cmn����_��v�N�x(CDate(Format(TestDate, "0000/00/00")))
Return

End Sub

'---------------------------------------------------------------------------------------
Public Function cmn����_��v�N�x(dtmDate As Date) As String

    cmn����_��v�N�x = CStr(Year(DateAdd("m", -3, dtmDate)))
    
End Function

'---------------------------------------------------------------------------------------
Sub test_����_�������()

    Dim TestDate As String

Debug.Print "test_cmn����_�������"
Debug.Print "�ݒ�l�@:�߂�l"

    TestDate = "20190331": GoSub Result
    TestDate = "20190401": GoSub Result
    TestDate = "20190930": GoSub Result
    TestDate = "20191001": GoSub Result
    TestDate = "20200331": GoSub Result
    TestDate = "20200401": GoSub Result
    
    Exit Sub
    
Result:
    Debug.Print TestDate & ":" & cmn����_�������(CDate(Format(TestDate, "0000/00/00")))
Return

End Sub

'---------------------------------------------------------------------------------------
'�������
'---------------------------------------------------------------------------------------
Public Function cmn����_�������(dtmDate As Date) As String

    Select Case Month(dtmDate)
                    
        Case 4, 5, 6, 7, 8, 9
            cmn����_������� = "���"
        Case 10, 11, 12, 1, 2, 3
            cmn����_������� = "����"
    
    End Select

End Function

'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
Sub main()

    Const ���[�� = "hoge���["
    Const QERY As String = "�N�G��1"
    Dim strSQL          As String
    strSQL = "SELECT * FROM " & QERY & ";"
    
    Call �ėp_���R�[�h�Z�b�g�̃Z���]�L(���[��, 1, strSQL)
    
End Sub

Private Sub �ėp_���R�[�h�Z�b�g�̃Z���]�L(���[�� As String, �w�b�_�s As Long, strSQL As String)

    '���Ԍv���p
    Dim StartTime As Single
    StartTime = Timer

    '// Exce���N��
    Call xlappOpen

    Call Create_ADO_Recordset(strSQL)

    With xlapp
        .Visible = True
'        .UserControl = True

        Set wb = .Workbooks.Add(Template:=xlWBATWorksheet) ' ���[�N�u�b�N���쐬
        Dim ws As Worksheet
        Set ws = wb.Worksheets(1)
        ws.Name = ���[�� & "_" & Format(date, "yyyymmdd")
        
        '�w�b�_�s�쐬
        Dim i As Integer 'myRs.Fields.Count
        For i = 0 To myRs.Fields.Count - 1
            ws.Cells(�w�b�_�s, i + 1) = myRs.Fields(i).Name
        Next

        '���R�[�h�Z�b�g�̃Z���]�L
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
        ws.Range(ws.Cells(�w�b�_�s + 1, 1), ws.Cells(myRs.RecordCount + 1, FieldsCount)).Value = tableArray

        Call �S�̌r��(ws)
        Call ���[�d�グ(ws, ���[��)
        
    End With


    Call ADO_Close
Debug.Print Timer - StartTime

    Call xlappClose


End Sub

'�P���\�����̑S�̌r������
Private Sub �S�̌r��(ws As Worksheet)

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

Private Sub ���[�d�グ(ws As Worksheet, ���[�� As String)

    '�w�b�_�s�F�t��
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column)).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With

    '����w�b�_
    With ws.PageSetup
'            .LeftHeader = "left"
        .CenterHeader = ���[��
        .RightHeader = Format(date, "yyyy/m/d")
        .CenterFooter = "&P/&N"
    End With

End Sub
