Attribute VB_Name = "Module1"
Option Compare Database
Option Explicit

Public Function IsHoliday(dDate As Variant) As Boolean
  On Error GoTo Err_Trap
'日付型の引数の場合はFalseを返す
  If IsDate(dDate) = False Then
    IsHoliday = False
    Exit Function
  End If
  
  If Weekday(dDate) = 1 Or Weekday(dDate) = 7 Or DCount("*", "M_休日", "休日=#" & dDate & "#") Then
    IsHoliday = True
  Else
    IsHoliday = False
  End If
  
Exit Function

Err_Trap:
  'エラー発生時はFalseを返す
  IsHoliday = False
  Exit Function
End Function

'---------------------------------------------------------------------------------------
Public Function WorkDay(dStartDate As Date, nWeight As Long)
  Dim dDate As Date
  Dim i As Long
  dDate = dStartDate
  i = 0
  
  If nWeight > 0 Then
    Do Until i = nWeight
      dDate = dDate + 1
      If IsHoliday(dDate) = False Then i = i + 1
    Loop
        
  Else
    Do Until i = nWeight
      dDate = dDate - 1
      If IsHoliday(dDate) = False Then i = i - 1
    Loop
  End If
  
  WorkDay = dDate

End Function

'---------------------------------------------------------------------------------------
Sub test_cmnGetEndOfMonth()

    Dim TestDate As String

Debug.Print "test_cmnGetEndOfMonth"
Debug.Print "設定値　:戻り値"

    TestDate = "20200101": GoSub Result
    TestDate = "20200201": GoSub Result
    TestDate = "20200601": GoSub Result
    
    Exit Sub
    
Result:
    Debug.Print TestDate & ":" & cmnGetEndOfMonth(Format(TestDate, "0000/00/00"))
Return

End Sub

'---------------------------------------------------------------------------------------
'指定した年月日に対応する月末日を取得
'---------------------------------------------------------------------------------------
Public Function cmnGetEndOfMonth(sDate As Date) As String

    cmnGetEndOfMonth = Format(DateSerial(Year(sDate), Month(sDate) + 1, 0), "yyyy/mm/dd")

End Function

