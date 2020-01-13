VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stopwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private mlngStat As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub StartTimer()
    mlngStat = GetTickCount
End Sub

Public Function EndTimer() As Long
    EndTimer = (GetTickCount - mlngStat)
End Function
