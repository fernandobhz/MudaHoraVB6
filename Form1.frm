VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Dim ndate As SYSTEMTIME
    ndate.wDay = 5
    ndate.wDayOfWeek = 4
    ndate.wMonth = 3
    ndate.wYear = 2013
    ndate.wHour = 6 + 3
    ndate.wMinute = 0
    ndate.wSecond = 0
    ndate.wMilliseconds = 0
    
    SetDate ndate
    
    Unload Me
End Sub
