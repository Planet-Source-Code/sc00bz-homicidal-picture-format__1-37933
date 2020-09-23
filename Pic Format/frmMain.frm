VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pic Format"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   312
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   531
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFileName 
      Height          =   315
      Left            =   1965
      TabIndex        =   4
      Text            =   "C:\Windows\Desktop\Test.hom"
      Top             =   4305
      Width           =   4035
   End
   Begin VB.PictureBox picTest1 
      AutoRedraw      =   -1  'True
      Height          =   3900
      Left            =   60
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   3
      Top             =   60
      Width           =   3900
   End
   Begin VB.PictureBox picTest2 
      AutoRedraw      =   -1  'True
      Height          =   3900
      Left            =   4005
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   2
      Top             =   60
      Width           =   3900
   End
   Begin VB.CommandButton cmdConvert 
      Height          =   255
      Left            =   4605
      TabIndex        =   1
      Top             =   4005
      Width           =   1395
   End
   Begin VB.HScrollBar hsbQuality 
      Height          =   255
      Left            =   1965
      Max             =   7
      Min             =   1
      TabIndex        =   0
      Top             =   4005
      Value           =   3
      Width           =   2595
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsPicFormat As New clsPicFormat
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Sub Fade()
    'Draws A Pic
    Dim A As Long
    
    For A = 0 To 255
        picTest1.Line (0, A)-(256, A), A
        picTest1.Line (A, 0)-(256, A), A
    Next
End Sub
Private Sub cmdConvert_Click()
    'Tests Saving And Loading
    Dim T1 As Long, T2 As Long, T3 As Long
    
    T1 = timeGetTime
    clsPicFormat.SavePic txtFileName.Text, picTest1.hDC, picTest1.ScaleWidth, picTest1.ScaleHeight, hsbQuality.Value
    T2 = timeGetTime
    clsPicFormat.LoadPic txtFileName.Text, picTest2.hDC
    picTest2.Refresh
    T3 = timeGetTime
    MsgBox "Took " & T2 - T1 & "ms to Save." & vbNewLine & "Took " & T3 - T2 & "ms to Load.", 0, ""
End Sub
Private Sub Form_Load()
    Fade
    hsbQuality_Change
End Sub
Private Sub hsbQuality_Change()
    If hsbQuality.Value = 1 Then
        cmdConvert.Caption = "Best"
    ElseIf hsbQuality.Value = 2 Then
        cmdConvert.Caption = "Excellent"
    ElseIf hsbQuality.Value = 3 Then
        cmdConvert.Caption = "Recommended"
    ElseIf hsbQuality.Value = 4 Then
        cmdConvert.Caption = "Good"
    ElseIf hsbQuality.Value = 5 Then
        cmdConvert.Caption = "Poor"
    ElseIf hsbQuality.Value = 6 Then
        cmdConvert.Caption = "Bad"
    ElseIf hsbQuality.Value = 7 Then
        cmdConvert.Caption = "Terrible"
    End If
End Sub
Private Sub hsbQuality_Scroll()
    hsbQuality_Change
End Sub
