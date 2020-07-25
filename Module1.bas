Attribute VB_Name = "Module1"
Option Explicit

Public Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Private Declare Function SetSystemTime Lib "kernel32" (lpSystemTime _
  As SYSTEMTIME) As Long

Public Function SetDate(mDate As SYSTEMTIME)
  Dim lReturn As Long
  lReturn = SetSystemTime(mDate)
End Function

