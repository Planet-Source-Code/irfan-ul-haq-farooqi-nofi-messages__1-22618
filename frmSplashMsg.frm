VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmSplashMsg 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   Icon            =   "frmSplashMsg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSplash 
      Interval        =   1000
      Left            =   3840
      Top             =   3240
   End
   Begin MSForms.Image Image1 
      Height          =   3855
      Left            =   0
      Top             =   0
      Width           =   4455
      BorderStyle     =   0
      SizeMode        =   3
      SpecialEffect   =   1
      Size            =   "7858;6800"
      Picture         =   "frmSplashMsg.frx":0CCA
   End
End
Attribute VB_Name = "frmSplashMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DelaySplash As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
        Unload Me
End Sub

Private Sub Form_Load()
    DelaySplash = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMessage.Show
End Sub

Private Sub Image1_Click()
        Unload Me
End Sub

Private Sub tmrSplash_Timer()
    If DelaySplash = 2 Then
        Unload Me
    Else
        DelaySplash = DelaySplash + 1
    End If
End Sub

