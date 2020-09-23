VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMessage 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NOFI Messages"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9405
   Icon            =   "Nofi Messages.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   9405
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTemp 
      Height          =   615
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   39
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComDlg.CommonDialog cmnDlg 
      Left            =   5400
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   375
      Left            =   7560
      TabIndex        =   35
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Frame fmType 
      Caption         =   "Message Box Type"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   5400
      TabIndex        =   26
      Top             =   0
      Width           =   3735
      Begin VB.ComboBox cmbVarType 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1800
         Width           =   1095
      End
      Begin VB.ComboBox cmbCondition 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtVarName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   29
         Text            =   "msgResponse"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.OptionButton optFunction 
         Caption         =   "Function"
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optStatement 
         Caption         =   "Statement"
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CheckBox chkVariable 
         Caption         =   "Declare Variable"
         Enabled         =   0   'False
         Height          =   255
         Left            =   840
         TabIndex        =   32
         Top             =   1830
         Width           =   1695
      End
      Begin VB.Label lblCondition 
         AutoSize        =   -1  'True
         Caption         =   "Conditional Code"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1680
         TabIndex        =   34
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label lblVarName 
         AutoSize        =   -1  'True
         Caption         =   "Variable Name"
         Enabled         =   0   'False
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   1080
         Width           =   1035
      End
   End
   Begin VB.Frame fmSpecial 
      Caption         =   "Special Options"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   5400
      TabIndex        =   19
      Top             =   2400
      Width           =   3735
      Begin VB.TextBox txtHelpFile 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "No Help File Selected"
         Top             =   2520
         Width           =   2655
      End
      Begin VB.TextBox txtHelpID 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   37
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton cmdOpenHelp 
         Caption         =   "Open Help File"
         Enabled         =   0   'False
         Height          =   285
         Left            =   480
         TabIndex        =   36
         Top             =   2160
         Width           =   1455
      End
      Begin VB.OptionButton optSysModel 
         Caption         =   "System Modal"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         ToolTipText     =   "All applications are suspended until the user responds to the message box"
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton optAppModel 
         Caption         =   "Application Modal"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         ToolTipText     =   "User must respond to the message box before continuing work in the current application"
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.CheckBox chkForeground 
         Caption         =   "Sets Foreground"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         ToolTipText     =   "Specifies the message box window as the foreground window"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CheckBox chkRight 
         Caption         =   "Text Right Aligned"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         ToolTipText     =   "Text is right aligned"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CheckBox chkHelp 
         Caption         =   "Help Button"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         ToolTipText     =   "Adds Help button to the message box"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblHelpID 
         AutoSize        =   -1  'True
         Caption         =   "Context ID:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2040
         TabIndex        =   38
         Top             =   1920
         Width           =   795
      End
      Begin VB.Label lblSpecial 
         Caption         =   "O P T I O N A L"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1695
         Left            =   3360
         TabIndex        =   25
         Top             =   480
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview"
      Default         =   -1  'True
      Height          =   375
      Left            =   6000
      TabIndex        =   12
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Frame fmIcon 
      Caption         =   "Icons"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   10
      Top             =   4920
      Width           =   4935
      Begin MSComctlLib.Toolbar tlbIcons 
         Height          =   630
         Left            =   1320
         TabIndex        =   18
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   1111
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         ImageList       =   "imlIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "None"
               ImageIndex      =   1
               Style           =   1
               Value           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Info"
               ImageIndex      =   2
               Style           =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Query"
               ImageIndex      =   3
               Style           =   1
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Exclamation"
               ImageIndex      =   4
               Style           =   1
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Critical"
               ImageIndex      =   5
               Style           =   1
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imlIcons 
         Left            =   240
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Nofi Messages.frx":0CCA
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Nofi Messages.frx":0FEE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Nofi Messages.frx":1442
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Nofi Messages.frx":1896
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Nofi Messages.frx":1CF2
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fmButton 
      Caption         =   "Buttons"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   255
      TabIndex        =   4
      Top             =   2160
      Width           =   4920
      Begin VB.Frame fmDefault 
         Caption         =   "Default Button"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   2640
         TabIndex        =   13
         Top             =   480
         Width           =   2055
         Begin VB.OptionButton optDefault4 
            Caption         =   "Fourth Button"
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   17
            Top             =   1440
            Width           =   1575
         End
         Begin VB.OptionButton optDefault3 
            Caption         =   "Third Button"
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   16
            Top             =   1080
            Width           =   1575
         End
         Begin VB.OptionButton optDefault2 
            Caption         =   "Second Button"
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   15
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton optDefault1 
            Caption         =   "First Button"
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.OptionButton optRC 
         Caption         =   "Retry and Cancel"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   2160
         Width           =   2295
      End
      Begin VB.OptionButton optYN 
         Caption         =   "Yes and No"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   2295
      End
      Begin VB.OptionButton optYNC 
         Caption         =   "Yes, No and Cancel"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   2295
      End
      Begin VB.OptionButton optARI 
         Caption         =   "Abort, Retry and Ignore"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   2415
      End
      Begin VB.OptionButton optOC 
         Caption         =   "OK and Cancel"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   2295
      End
      Begin VB.OptionButton optOk 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   2295
      End
   End
   Begin VB.TextBox txtMessage 
      Height          =   1455
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   600
      Width           =   3855
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   2
      Text            =   "Your Title Here ..."
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblMessage 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Message:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   255
      TabIndex        =   1
      Top             =   600
      Width           =   960
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Title:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   495
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu smnNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu smnSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu smnOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu smnFD1 
         Caption         =   "-"
      End
      Begin VB.Menu smnPreview 
         Caption         =   "&Preview"
         Shortcut        =   {F9}
      End
      Begin VB.Menu smnCopy 
         Caption         =   "Copy to &Clipboard"
         Shortcut        =   ^C
      End
      Begin VB.Menu smnFD2 
         Caption         =   "-"
      End
      Begin VB.Menu smnExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
      Begin VB.Menu smnAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msgButton As Integer
Dim msgIcon As Integer
Dim msgBtnDefault As Integer
Dim msgModel As Integer
Dim msgForeground As Long
Dim msgRight As Long
Dim msgHelp As Long

Dim cpmsgButton As String
Dim cpmsgIcon As String
Dim cpmsgBtnDefault As String
Dim cpmsgModel As String
Dim cpmsgForeground As String
Dim cpmsgRight As String
Dim cpmsgHelp As String

Dim HelpResponse As Integer
Dim HelpFileName As String

Dim FSO As FileSystemObject
Dim txtStream As TextStream

Private Sub chkForeground_Click()
    If chkForeground Then
        cpmsgForeground = " + vbMsgBoxSetForeground"
    Else
        cpmsgForeground = ""
    End If
End Sub

Private Sub chkHelp_Click()
    If chkHelp.Value = 1 Then
        cmdOpenHelp.Enabled = True
        lblHelpID.Enabled = True
        txtHelpID.Enabled = True
        txtHelpFile.Enabled = True
        
        If optOk Then
            optDefault2.Enabled = True
        ElseIf optOC Or optYN Or optRC Then
            optDefault3.Enabled = True
        ElseIf optARI Or optYNC Then
            optDefault4.Enabled = True
        End If
        
        cpmsgHelp = " + vbMsgBoxHelpButton"
    Else
        cmdOpenHelp.Enabled = False
        lblHelpID.Enabled = False
        txtHelpID.Enabled = False
        txtHelpFile.Enabled = False
        
        If optOk Then
            optDefault2.Enabled = False
        ElseIf optOC Or optYN Or optRC Then
            optDefault3.Enabled = False
        ElseIf optARI Or optYNC Then
            optDefault4.Enabled = False
        End If
    End If
End Sub

Private Sub chkRight_Click()
    If chkRight Then
        cpmsgRight = " + vbMsgBoxRight"
    Else
        cpmsgRight = ""
    End If
End Sub

Private Sub chkVariable_Click()
    If chkVariable.Value = 1 Then
        cmbVarType.Enabled = True
    Else
        cmbVarType.Enabled = False
    End If
End Sub

Private Sub cmdCopy_Click()
    Clipboard.Clear
    
    If chkHelp Then
        If HelpFileName = "" Or Right(HelpFileName, 3) <> "hlp" Or txtHelpID.Text = "" Then
            HelpResponse = MsgBox("You must provide the Help File and a valid ContextID for the Help Button", vbRetryCancel + vbQuestion + vbDefaultButton2, "Invalid Help File")
            If HelpResponse = vbCancel Then
                chkHelp.Value = 0
            End If
        Else
            txtTemp.Text = "MsgBox " & Chr(34) & txtMessage.Text & Chr(34) & ", _" & vbCrLf & vbTab & cpmsgButton & cpmsgIcon & cpmsgBtnDefault & cpmsgModel & cpmsgForeground & cpmsgRight & cpmsgHelp & ", " & Chr(34) & txtTitle.Text & Chr(34) & ", " & Chr(34) & HelpFileName & Chr(34) & ", " & Val(txtHelpID.Text) & vbCrLf
        End If
    Else
        txtTemp.Text = "MsgBox " & Chr(34) & txtMessage.Text & Chr(34) & ", _" & vbCrLf & vbTab & cpmsgButton & cpmsgIcon & cpmsgBtnDefault & cpmsgModel & cpmsgForeground & cpmsgRight & cpmsgHelp & ", " & Chr(34) & txtTitle.Text & Chr(34) & vbCrLf
    End If
    
    If optFunction Then
        Set txtStream = FSO.CreateTextFile(App.Path & "\message.tmp", True)
        
        If chkVariable Then
            txtStream.WriteLine "'"
            txtStream.WriteLine "' Please put this declaration in your declarations section!"
            txtStream.WriteLine cmbVarType.Text & " " & txtVarName.Text & " As Integer"
            txtStream.WriteBlankLines 1
        End If
        
        txtStream.Write txtVarName & " = "
        If chkHelp Then
            txtStream.WriteLine "MsgBox (" & Chr(34) & txtMessage.Text & Chr(34) & ", _" & vbCrLf & vbTab & vbTab & cpmsgButton & cpmsgIcon & cpmsgBtnDefault & cpmsgModel & cpmsgForeground & cpmsgRight & cpmsgHelp & ", " & Chr(34) & txtTitle.Text & Chr(34) & ", " & Chr(34) & HelpFileName & Chr(34) & ", " & txtHelpID.Text & ")"
        Else
            txtStream.WriteLine "MsgBox (" & Chr(34) & txtMessage.Text & Chr(34) & ", _" & vbCrLf & vbTab & vbTab & cpmsgButton & cpmsgIcon & cpmsgBtnDefault & cpmsgModel & cpmsgForeground & cpmsgRight & cpmsgHelp & ", " & Chr(34) & txtTitle.Text & Chr(34) & ")"
        End If
        txtStream.WriteBlankLines 1
        
        Select Case cmbCondition.Text
        Case "If..Then..Else"
            If optOk Then
                txtStream.WriteLine "If " & txtVarName & " = vbOK Then"
                txtStream.WriteLine vbTab
            ElseIf optOC Then
                txtStream.WriteLine "If " & txtVarName & " = vbOK Then"
                txtStream.WriteLine vbTab
                txtStream.WriteLine "ElseIf " & txtVarName & " = vbCancel Then"
                txtStream.WriteLine vbTab
            ElseIf optARI Then
                txtStream.WriteLine "If " & txtVarName & " = vbAbort Then"
                txtStream.WriteLine vbTab
                txtStream.WriteLine "ElseIf " & txtVarName & " = vbRetry Then"
                txtStream.WriteLine vbTab
                txtStream.WriteLine "ElseIf " & txtVarName & " = vbIgnore Then"
                txtStream.WriteLine vbTab
            ElseIf optYNC Then
                txtStream.WriteLine "If " & txtVarName & " = vbYes Then"
                txtStream.WriteLine vbTab
                txtStream.WriteLine "ElseIf " & txtVarName & " = vbNo Then"
                txtStream.WriteLine vbTab
                txtStream.WriteLine "ElseIf " & txtVarName & " = vbCancel Then"
                txtStream.WriteLine vbTab
            ElseIf optYN Then
                txtStream.WriteLine "If " & txtVarName & " = vbYes Then"
                txtStream.WriteLine vbTab
                txtStream.WriteLine "ElseIf " & txtVarName & " = vbNo Then"
                txtStream.WriteLine vbTab
            ElseIf optRC Then
                txtStream.WriteLine "If " & txtVarName & " = vbRetry Then"
                txtStream.WriteLine vbTab
                txtStream.WriteLine "ElseIf " & txtVarName & " = vbCancel Then"
                txtStream.WriteLine vbTab
            End If
            txtStream.WriteLine "Endif"
        Case "Select Case"
            txtStream.WriteLine "Select Case " & txtVarName
            If optOk Then
                txtStream.WriteLine "Case vbOK"
                txtStream.WriteLine vbTab
            ElseIf optOC Then
                txtStream.WriteLine "Case vbOK"
                txtStream.WriteLine vbTab
                txtStream.WriteLine "Case vbCancel"
                txtStream.WriteLine vbTab
            ElseIf optARI Then
                txtStream.WriteLine "Case vbAbort"
                txtStream.WriteLine vbTab
                txtStream.WriteLine "Case vbRetry"
                txtStream.WriteLine vbTab
                txtStream.WriteLine "Case vbIgnore"
                txtStream.WriteLine vbTab
            ElseIf optYNC Then
                txtStream.WriteLine "Case vbYes"
                txtStream.WriteLine vbTab
                txtStream.WriteLine "Case vbNo"
                txtStream.WriteLine vbTab
                txtStream.WriteLine "Case vbCancel"
                txtStream.WriteLine vbTab
            ElseIf optYN Then
                txtStream.WriteLine "Case vbYes"
                txtStream.WriteLine vbTab
                txtStream.WriteLine "Case vbNo"
                txtStream.WriteLine vbTab
            ElseIf optRC Then
                txtStream.WriteLine "Case vbRetry"
                txtStream.WriteLine vbTab
                txtStream.WriteLine "Case vbCancel"
                txtStream.WriteLine vbTab
            End If
            txtStream.WriteLine "End Select"
        End Select

        Set txtStream = FSO.OpenTextFile(App.Path & "\message.tmp", ForReading)
        txtTemp.Text = txtStream.ReadAll
        txtStream.Close
    End If

    Clipboard.SetText txtTemp.Text
    MsgBox "Your Syntex for this Message Box is copied to the clipboard successfully", vbInformation, "Copy to Clipboard"
    Me.WindowState = 1
End Sub

Private Sub cmdOpenHelp_Click()
    On Error GoTo OpenHelpError
    
    With cmnDlg
        .DialogTitle = "Open Help File"
        .InitDir = CurDir
        .FileName = ""
        .Filter = "Help File (*.hlp)|*.hlp"
        .Flags = &H1000 + &H4
        .ShowOpen
    End With
    HelpFileName = cmnDlg.FileName
    txtHelpFile.Text = cmnDlg.FileName
    chkHelp.SetFocus
    
    Exit Sub
    
OpenHelpError:
    If Err.Number = 32755 Then
    Else
        MsgBox Err.Description, vbCritical, "Error"
    End If
End Sub

Private Sub cmdPreview_Click()
    If optOk Then
        msgButton = vbOKOnly
    ElseIf optOC Then
        msgButton = vbOKCancel
    ElseIf optARI Then
        msgButton = vbAbortRetryIgnore
    ElseIf optYNC Then
        msgButton = vbYesNoCancel
    ElseIf optYN Then
        msgButton = vbYesNo
    ElseIf optRC Then
        msgButton = vbRetryCancel
    End If
    
    If optDefault1 Then
        msgBtnDefault = vbDefaultButton1
    ElseIf optDefault2 Then
        msgBtnDefault = vbDefaultButton2
    ElseIf optDefault3 Then
        msgBtnDefault = vbDefaultButton3
    ElseIf optDefault4 Then
        msgBtnDefault = vbDefaultButton4
    End If
    
    If tlbIcons.Buttons.Item(1).Value = tbrPressed Then
        msgIcon = 0
    ElseIf tlbIcons.Buttons.Item(3).Value = tbrPressed Then
        msgIcon = vbInformation
    ElseIf tlbIcons.Buttons.Item(5).Value = tbrPressed Then
        msgIcon = vbQuestion
    ElseIf tlbIcons.Buttons.Item(7).Value = tbrPressed Then
        msgIcon = vbExclamation
    ElseIf tlbIcons.Buttons.Item(9).Value = tbrPressed Then
        msgIcon = vbCritical
    End If
        
    If optAppModel Then
        msgModel = vbApplicationModal
    ElseIf optSysModel Then
        msgModel = vbSystemModal
    End If
    
    If chkForeground Then
        msgForeground = vbMsgBoxSetForeground
    End If
    
    If chkRight Then
        msgRight = vbMsgBoxRight
    End If
    
    If chkHelp Then
        msgHelp = vbMsgBoxHelpButton
    End If
    
    MsgBox txtMessage.Text, msgButton + msgIcon + msgBtnDefault + msgModel + msgForeground + msgRight + msgHelp, txtTitle.Text, HelpFileName, Val(txtHelpID.Text)
End Sub

Private Sub Form_Load()
    msgButton = vbOKOnly
    msgIcon = 0
    msgBtnDefault = 0
    msgModel = 0
    msgForeground = 0
    msgRight = 0
    msgHelp = 0
    
    cpmsgButton = "vbOKOnly"
    cpmsgIcon = ""
    cpmsgBtnDefault = ""
    cpmsgModel = ""
    cpmsgForeground = ""
    cpmsgRight = ""
    cpmsgHelp = ""
    
    cmbCondition.AddItem "If..Then..Else", 0
    cmbCondition.AddItem "Select Case", 1
    cmbCondition.ListIndex = 0
    
    cmbVarType.AddItem "Dim", 0
    cmbVarType.AddItem "Private", 1
    cmbVarType.AddItem "Public", 2
    cmbVarType.ListIndex = 0
    
    HelpFileName = ""
    txtHelpID = ""
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
End Sub


Private Sub optAppModel_Click()
    cpmsgModel = ""
    chkHelp.Enabled = True
End Sub

Private Sub optARI_Click()
    optDefault1.Value = True
    optDefault1.Enabled = True
    optDefault2.Enabled = True
    optDefault3.Enabled = True
    If chkHelp Then
        optDefault4.Enabled = True
    Else
        optDefault4.Enabled = False
    End If
    
    optOk.FontBold = False
    optOk.ForeColor = vbButtonText
    optOC.FontBold = False
    optOC.ForeColor = vbButtonText
    optARI.FontBold = True
    optARI.ForeColor = vbHighlight
    optYNC.FontBold = False
    optYNC.ForeColor = vbButtonText
    optYN.FontBold = False
    optYN.ForeColor = vbButtonText
    optRC.FontBold = False
    optRC.ForeColor = vbButtonText
    
    cpmsgButton = "vbAbortRetryIgnore"
End Sub

Private Sub optDefault1_Click()
    cpmsgBtnDefault = ""
End Sub

Private Sub optDefault2_Click()
    cpmsgBtnDefault = " + vbDefaultButton2"
End Sub

Private Sub optDefault3_Click()
    cpmsgBtnDefault = " + vbDefaultButton3"
End Sub

Private Sub optDefault4_Click()
    cpmsgBtnDefault = " + vbDefaultButton4"
End Sub

Private Sub optFunction_Click()
    lblVarName.Enabled = True
    txtVarName.Enabled = True
    lblCondition.Enabled = True
    cmbCondition.Enabled = True
    chkVariable.Enabled = True
    If chkVariable Then
        cmbVarType.Enabled = True
    Else
        cmbVarType.Enabled = False
    End If
End Sub

Private Sub optOC_Click()
    optDefault1.Value = True
    optDefault1.Enabled = True
    optDefault2.Enabled = True
    If chkHelp Then
        optDefault3.Enabled = True
    Else
        optDefault3.Enabled = False
    End If
    optDefault4.Enabled = False
    
    optOC.FontBold = True
    optOC.ForeColor = vbHighlight
    optOk.FontBold = False
    optOk.ForeColor = vbButtonText
    optARI.FontBold = False
    optARI.ForeColor = vbButtonText
    optYNC.FontBold = False
    optYNC.ForeColor = vbButtonText
    optYN.FontBold = False
    optYN.ForeColor = vbButtonText
    optRC.FontBold = False
    optRC.ForeColor = vbButtonText
    
    cpmsgButton = "vbOKCancel"
End Sub

Private Sub optOk_Click()
    optDefault1.Value = True
    optDefault1.Enabled = True
    If chkHelp Then
        optDefault2.Enabled = True
    Else
        optDefault2.Enabled = False
    End If
    optDefault3.Enabled = False
    optDefault4.Enabled = False
    
    optOk.FontBold = True
    optOk.ForeColor = vbHighlight
    optOC.FontBold = False
    optOC.ForeColor = vbButtonText
    optARI.FontBold = False
    optARI.ForeColor = vbButtonText
    optYNC.FontBold = False
    optYNC.ForeColor = vbButtonText
    optYN.FontBold = False
    optYN.ForeColor = vbButtonText
    optRC.FontBold = False
    optRC.ForeColor = vbButtonText
    
    cpmsgButton = "vbOKOnly"
End Sub

Private Sub optRC_Click()
    optDefault1.Value = True
    optDefault1.Enabled = True
    optDefault2.Enabled = True
    If chkHelp Then
        optDefault3.Enabled = True
    Else
        optDefault3.Enabled = False
    End If
    optDefault4.Enabled = False
    
    optOk.FontBold = False
    optOk.ForeColor = vbButtonText
    optOC.FontBold = False
    optOC.ForeColor = vbButtonText
    optARI.FontBold = False
    optARI.ForeColor = vbButtonText
    optYNC.FontBold = False
    optYNC.ForeColor = vbButtonText
    optYN.FontBold = False
    optYN.ForeColor = vbButtonText
    optRC.FontBold = True
    optRC.ForeColor = vbHighlight
    
    cpmsgButton = "vbRetryCancel"
End Sub

Private Sub optStatement_Click()
    lblVarName.Enabled = False
    txtVarName.Enabled = False
    lblCondition.Enabled = False
    cmbCondition.Enabled = False
    chkVariable.Enabled = False
    cmbVarType.Enabled = False
End Sub

Private Sub optSysModel_Click()
    cpmsgModel = " + vbSystemModal"
    chkHelp.Value = 0
    chkHelp.Enabled = False
End Sub

Private Sub optYN_Click()
    optDefault1.Value = True
    optDefault1.Enabled = True
    optDefault2.Enabled = True
    If chkHelp Then
        optDefault3.Enabled = True
    Else
        optDefault3.Enabled = False
    End If
    optDefault4.Enabled = False
    
    optOk.FontBold = False
    optOk.ForeColor = vbButtonText
    optOC.FontBold = False
    optOC.ForeColor = vbButtonText
    optARI.FontBold = False
    optARI.ForeColor = vbButtonText
    optYNC.FontBold = False
    optYNC.ForeColor = vbButtonText
    optYN.FontBold = True
    optYN.ForeColor = vbHighlight
    optRC.FontBold = False
    optRC.ForeColor = vbButtonText
    
    cpmsgButton = "vbYesNo"
End Sub

Private Sub optYNC_Click()
    optDefault1.Value = True
    optDefault1.Enabled = True
    optDefault2.Enabled = True
    optDefault3.Enabled = True
    If chkHelp Then
        optDefault4.Enabled = True
    Else
        optDefault4.Enabled = False
    End If
    
    optOk.FontBold = False
    optOk.ForeColor = vbButtonText
    optOC.FontBold = False
    optOC.ForeColor = vbButtonText
    optARI.FontBold = False
    optARI.ForeColor = vbButtonText
    optYNC.FontBold = True
    optYNC.ForeColor = vbHighlight
    optYN.FontBold = False
    optYN.ForeColor = vbButtonText
    optRC.FontBold = False
    optRC.ForeColor = vbButtonText
    
    cpmsgButton = "vbYesNoCancel"
End Sub

Private Sub smnAbout_Click()
    frmAboutMsg.Show
End Sub

Private Sub smnCopy_Click()
    Call cmdCopy_Click
End Sub

Private Sub smnExit_Click()
    End
End Sub

Private Sub smnExitTray_Click()
    End
End Sub

Private Sub smnNew_Click()
    txtTitle.Text = ""
    txtMessage.Text = ""
    optOk.Value = True
    optStatement.Value = True
    optAppModel.Value = True
    chkForeground.Value = 0
    chkRight.Value = 0
    chkHelp.Value = 0
    HelpFileName = ""
    txtHelpID = ""
    txtHelpFile = "No Help File Selected"
    ButtonValue (1)
    chkVariable.Value = 0
    cmbCondition.ListIndex = 0
    cmbVarType.ListIndex = 0
End Sub

Private Sub smnOpen_Click()
    On Error GoTo OpenError
    
    With cmnDlg
        .DialogTitle = "Open Message Box"
        .InitDir = App.Path
        .FileName = ""
        .Filter = "MessageBox Syntex File (*.mgb)|*.mgb"
        .DefaultExt = "mgb"
        .Flags = &H1000 + &H4
        .ShowOpen
    End With
    
    If cmnDlg.FileName = "" Then
    Else
        Set txtStream = FSO.OpenTextFile(cmnDlg.FileName, ForReading)
        
        txtTitle.Text = txtStream.ReadLine
        txtMessage.Text = txtStream.ReadLine
        
        Select Case txtStream.ReadLine
        Case "vbOKOnly"
            optOk.Value = True
        Case "vbOKCancel"
            optOC.Value = True
        Case "vbAbortRetryIgnore"
            optARI.Value = True
        Case "vbYesNoCancel"
            optYNC.Value = True
        Case "vbYesNo"
            optYN.Value = True
        Case "vbRetryCancel"
            optRC.Value = True
        End Select
        
        Select Case Right(txtStream.ReadLine, 1)
        Case "1"
            optDefault1.Value = True
        Case "2"
            optDefault2.Value = True
        Case "3"
            optDefault3.Value = True
        Case "4"
            optDefault4.Value = True
        End Select
        
        Select Case txtStream.ReadLine
        Case ""
            ButtonValue (1)
        Case "vbInformation"
            ButtonValue (3)
        Case "vbQuestion"
            ButtonValue (5)
        Case "vbExclamation"
            ButtonValue (7)
        Case "vbCritical"
            ButtonValue (9)
        End Select
        
        Select Case txtStream.ReadLine
        Case ""
            optAppModel.Value = True
        Case "vbSystemModal"
            optSysModel.Value = True
        End Select
        
        If txtStream.ReadLine = "vbMsgBoxSetForeground" Then
            chkForeground.Value = 1
        Else
            chkForeground.Value = 0
        End If
        
        If txtStream.ReadLine = "vbMsgBoxRight" Then
            chkRight.Value = 1
        Else
            chkRight.Value = 0
        End If
        
        If txtStream.ReadLine = "vbMsgBoxHelpButton" Then
            chkHelp.Value = 1
            HelpFileName = txtStream.ReadLine
            txtHelpFile.Text = HelpFileName
            txtHelpID.Text = txtStream.ReadLine
        Else
            chkHelp.Value = 0
            txtStream.SkipLine
            txtStream.SkipLine
        End If
        
        If txtStream.ReadLine = "Function" Then
            optFunction.Value = True
            txtVarName.Text = txtStream.ReadLine
            Select Case txtStream.ReadLine
            Case "If..Then..Else"
                cmbCondition.ListIndex = 0
            Case "Select Case"
                cmbCondition.ListIndex = 1
            End Select
            
            Var = txtStream.ReadLine
            If Var <> "" Then
                chkVariable.Value = 1
                Select Case Var
                Case "Dim"
                    cmbVarType.ListIndex = 0
                Case "Private"
                    cmbVarType.ListIndex = 1
                Case "Public"
                    cmbVarType.ListIndex = 2
                End Select
            End If
        Else
            optStatement.Value = True
        End If
        
    End If
    
    Exit Sub

OpenError:
    If Err.Number = 32755 Then
    Else
        MsgBox Err.Description, vbCritical, "Error"
    End If
End Sub

Private Sub smnPreview_Click()
    Call cmdPreview_Click
End Sub

Private Sub smnRestore_Click()
    TrayIcon.Restore
End Sub

Private Sub smnSave_Click()
    On Error GoTo SaveError

    If chkHelp Then
        If HelpFileName = "" Or Right(HelpFileName, 3) <> "hlp" Or txtHelpID.Text = "" Then
            HelpResponse = MsgBox("You must provide the Help File and a valid ContextID for the Help Button. Click Cancel for no Help Button.", vbRetryCancel + vbQuestion + vbDefaultButton2, "Invalid Help File")
            If HelpResponse = vbCancel Then
                chkHelp.Value = 0
            ElseIf HelpResponse = vbRetry Then
                Exit Sub
            End If
        End If
    Else
        txtTemp.Text = "MsgBox " & Chr(34) & txtMessage.Text & Chr(34) & ", _" & vbCrLf & vbTab & cpmsgButton & cpmsgIcon & cpmsgBtnDefault & cpmsgModel & cpmsgForeground & cpmsgRight & cpmsgHelp & ", " & Chr(34) & txtTitle.Text & Chr(34) & vbCrLf
    End If
    
    With cmnDlg
        .DialogTitle = "Save Message Box"
        .InitDir = App.Path
        .FileName = ""
        .Filter = "MessageBox Syntex File (*.mgb)|*.mgb"
        .Flags = &H2 + &H4
        .ShowSave
    End With
    
    If cmnDlg.FileName = "" Then
    Else
        Set txtStream = FSO.CreateTextFile(cmnDlg.FileName, False)
        Call SaveData
    End If
        
    Exit Sub

SaveError:
    If Err.Number = 58 Then
    ElseIf Err.Number = 32755 Then
    Else
        MsgBox Err.Description, vbCritical, "Error"
    End If
End Sub

Private Sub tlbIcons_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Key = "None" Then
        msgIcon = 0
        cpmsgIcon = ""
        ButtonValue (1)
    ElseIf Button.Key = "Info" Then
        msgIcon = vbInformation
        cpmsgIcon = " + vbInformation"
        ButtonValue (3)
    ElseIf Button.Key = "Query" Then
        msgIcon = vbQuestion
        cpmsgIcon = " + vbQuestion"
        ButtonValue (5)
    ElseIf Button.Key = "Exclamation" Then
        msgIcon = vbExclamation
        cpmsgIcon = " + vbExclamation"
        ButtonValue (7)
    ElseIf Button.Key = "Critical" Then
        msgIcon = vbCritical
        cpmsgIcon = " + vbCritical"
        ButtonValue (9)
    End If
End Sub

Private Sub ButtonValue(ButtonIndex As Integer)
    tlbIcons.Buttons.Item(1).Value = tbrUnpressed
    tlbIcons.Buttons.Item(3).Value = tbrUnpressed
    tlbIcons.Buttons.Item(5).Value = tbrUnpressed
    tlbIcons.Buttons.Item(7).Value = tbrUnpressed
    tlbIcons.Buttons.Item(9).Value = tbrUnpressed
    tlbIcons.Buttons.Item(ButtonIndex).Value = tbrPressed
End Sub

Private Sub SaveData()
    txtStream.WriteLine txtTitle.Text
    txtStream.WriteLine txtMessage.Text
    txtStream.WriteLine cpmsgButton
    txtStream.WriteLine Mid(cpmsgBtnDefault, 4)
    txtStream.WriteLine Mid(cpmsgIcon, 4)
    txtStream.WriteLine Mid(cpmsgModel, 4)
    txtStream.WriteLine Mid(cpmsgForeground, 4)
    txtStream.WriteLine Mid(cpmsgRight, 4)
    txtStream.WriteLine Mid(cpmsgHelp, 4)
    txtStream.WriteLine HelpFileName
    txtStream.WriteLine txtHelpID
    If optFunction Then
        txtStream.WriteLine "Function"
        txtStream.WriteLine txtVarName.Text
        txtStream.WriteLine cmbCondition.Text
        If chkVariable Then
            txtStream.WriteLine cmbVarType.Text
        Else
            txtStream.WriteLine ""
        End If
    Else
        txtStream.WriteLine "Statement"
    End If
        
    txtStream.Close
End Sub
