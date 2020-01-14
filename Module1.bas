Attribute VB_Name = "Module1"
Option Compare Database
Option Explicit

Public Function IsHoliday(dDate As Variant) As Boolean
  On Error GoTo Err_Trap
'���t�^�̈����̏ꍇ��False��Ԃ�
  If IsDate(dDate) = False Then
    IsHoliday = False
    Exit Function
  End If
  
  If Weekday(dDate) = 1 Or Weekday(dDate) = 7 Or DCount("*", "M_�x��", "�x��=#" & dDate & "#") Then
    IsHoliday = True
  Else
    IsHoliday = False
  End If
  
Exit Function

Err_Trap:
  '�G���[��������False��Ԃ�
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
Debug.Print "�ݒ�l�@:�߂�l"

    TestDate = "20200101": GoSub Result
    TestDate = "20200201": GoSub Result
    TestDate = "20200601": GoSub Result
    
    Exit Sub
    
Result:
    Debug.Print TestDate & ":" & cmnGetEndOfMonth(Format(TestDate, "0000/00/00"))
Return

End Sub

'---------------------------------------------------------------------------------------
'�w�肵���N�����ɑΉ����錎�������擾
'---------------------------------------------------------------------------------------
Public Function cmnGetEndOfMonth(sDate As Date) As String

    cmnGetEndOfMonth = Format(DateSerial(Year(sDate), Month(sDate) + 1, 0), "yyyy/mm/dd")

End Function


'---------------------------------------------------------------------------------------
Public Function CheckHoliday(dt As Date) As Boolean
Dim flg As Boolean
    'holiday(�j�Փ��e�[�u��)�e�[�u�����������A�����Ƃ��Ď󂯎�������t���j�Փ���
    '�����邩�ǂ����m�F����
    If IsNull(DLookup("holiday", "holiday", "holiday = #" & Format(dt, "yyyy/mm/dd") & "#")) Then
        '�j�Փ��ɊY�����Ȃ��ꍇ�́A�y�j�������j�������`�F�b�N
        '�y�����x�݂łȂ��ꍇ�́ACase�Ɏw�肷�鐔�l���Y���̗j����\�����l�ɕύX����B
        Select Case Weekday(dt, vbSunday) '���j����1�A�y�j����7�ɂȂ�
            Case 1
                CheckHoliday = True
            Case 7
                CheckHoliday = True
            Case Else
                CheckHoliday = False
        End Select
    Else
        '�����Ɏw�肵�����t��holiday(�j�Փ��e�[�u��)�e�[�u���̓��t�ɊY���A�܂�j�Փ�
        CheckHoliday = True
    End If
End Function
