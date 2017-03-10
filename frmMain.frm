VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "proxy_exe"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_Load()
    Dim sExecute As String
    Dim sOutput As String
    Dim lMetaCodeCount As Long
    Dim lCount As Long
    
    sExecute = App.Path & "\scip.exe"
    
    On Error Resume Next
    Open App.Path & "\proxy_exe.log" For Output As #1
        Print #1, Date & " " & Time & " " & sExecute
    Close
    
    Call ShellExecute(frmMain.hwnd, "Open", sExecute, "", App.Path, 1)
    
    Unload Me
End Sub
