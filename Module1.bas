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


'---------------------------------------------------------------------------------------
Public Function CheckHoliday(dt As Date) As Boolean
Dim flg As Boolean
    'holiday(祝祭日テーブル)テーブルを検索し、引数として受け取った日付が祝祭日に
    'あたるかどうか確認する
    If IsNull(DLookup("holiday", "holiday", "holiday = #" & Format(dt, "yyyy/mm/dd") & "#")) Then
        '祝祭日に該当しない場合は、土曜日か日曜日かをチェック
        '土日が休みでない場合は、Caseに指定する数値を該当の曜日を表す数値に変更する。
        Select Case Weekday(dt, vbSunday) '日曜日が1、土曜日が7になる
            Case 1
                CheckHoliday = True
            Case 7
                CheckHoliday = True
            Case Else
                CheckHoliday = False
        End Select
    Else
        '引数に指定した日付がholiday(祝祭日テーブル)テーブルの日付に該当、つまり祝祭日
        CheckHoliday = True
    End If
End Function



Sub test_func月間営業日数()

    Debug.Print func月間営業日数("2020/2/22")

End Sub

Public Function func月間営業日数(ByVal strDate As String) As Long

On Error GoTo Err_Proc
    
    Dim 月初日 As Date, 月末日 As Date
    
    Dim db              As ADODB.Connection
    Dim rs              As ADODB.Recordset
    
    '// カレンダーマスタ上の祝祭日休業日および土日営業日の日数を取得
    Dim strSQL As String: strSQL = ""
    strSQL = "SELECT count(*) FROM Mカレンダー AS a "
    strSQL = strSQL & "WHERE (a.日付) Between #" & 月初日 & "# And #" & 月末日 & "# AND a.休業区分 = '1'"


    月初日 = DateSerial(Year(strDate), Month(strDate), 1)
    月末日 = DateSerial(Year(strDate), Month(strDate) + 1, 0)

    '// 月間日数から土日を除いた日数を算出
    Dim dtTmp As Date: dtTmp = 月初日
    Dim 営業日数 As Long: 営業日数 = 0
    Do While (dtTmp <= 月末日)

        If ((Weekday(dtTmp) <> vbSunday) And _
            (Weekday(dtTmp) <> vbSaturday)) Then
            営業日数 = 営業日数 + 1
        End If

        dtTmp = dtTmp + 1
    Loop
    
'    Set db = CurrentDb
'    Set rs = db.OpenRecordset(strSQL)
'    With rs
'        If Not .EOF Then
'            .Move (N - 1)
'            Hiduke = .Fields("日付")
'        End If
'    End With
'    rs.Close
'    db.Close
    
Exit_func月間営業日数:
    func月間営業日数 = 営業日数
    Exit Function

Err_Proc:
    func月間営業日数 = 0
    Resume Exit_func月間営業日数

End Function

'SELECT Max(日付) AS 最新日 
'FROM TBL_日付 
'HAVING DateSerial(Left([日付],4),Mid([日付],5,2),Right([日付],2))<=#2019/12/30#

SELECT T1.列1, T1.列2
FROM 
 (SELECT T2.列1, T2.列2, T2.列3, T2.列4
  FROM(
   SELECT DISTINCT 表1.列1,表1.列2,表2.列3, 表2.列4 
   FROM 表1, 表2 
   WHERE 表1.列1=表2.列3
  ) T2
 ORDER BY T2.列4
)T1;
                    
              
              
              
              
