VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3615
   ClientLeft      =   3930
   ClientTop       =   3540
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4680
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkBorder 
      BackColor       =   &H00000000&
      Caption         =   "Show Border"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.PictureBox picOpen 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   3600
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   390
      ScaleWidth      =   750
      TabIndex        =   3
      ToolTipText     =   "Click Here To Open A File"
      Top             =   2640
      Width           =   750
   End
   Begin VB.PictureBox picSave 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   240
      Picture         =   "frmMain.frx":07E1
      ScaleHeight     =   390
      ScaleWidth      =   750
      TabIndex        =   1
      ToolTipText     =   "Click Here To Save Your Data"
      Top             =   2640
      Width           =   750
   End
   Begin VB.PictureBox picMin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3960
      Picture         =   "frmMain.frx":0FAE
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Minimize SimpleText"
      Top             =   0
      Width           =   285
   End
   Begin VB.PictureBox picClose 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4320
      Picture         =   "frmMain.frx":1342
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Close SimpleText"
      Top             =   0
      Width           =   285
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1815
      Left            =   360
      TabIndex        =   0
      ToolTipText     =   "Type your document here"
      Top             =   720
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   3201
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":1605
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   375
      Left            =   360
      TabIndex        =   7
      ToolTipText     =   "Program Title - What Else?"
      Top             =   3120
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Simple Text Saver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   360
      TabIndex        =   6
      ToolTipText     =   "Program Title - What Else?"
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkBorder_Click()
    If chkBorder.Value = 1 Then
        Me.Caption = "SimpleText"
    Else
        Me.Caption = ""
    End If
End Sub

Private Sub Form_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    RichTextBox1.BackColor = &HFFFFFF
    chkBorder.Value = 0
    lblStatus.Caption = "No Text or File Loaded"
End Sub

Private Sub picClose_Click()
 Dim Msg, Response   ' Declare variables.
   Msg = "Are you sure you want to exit?"
   Msg = Msg + vbCrLf + "(Save your work first.)"
   Response = MsgBox(Msg, vbQuestion + vbOKCancel, "Exit SimpleText")
   Select Case Response
      Case vbCancel   ' Don't allow close.
         Cancel = -1
         RichTextBox1.SetFocus
      Case vbOK
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        Unload Me
   End Select
End Sub

Private Sub picMin_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub picOpen_Click()
   Dim sFile As String
    With dlgCommonDialog
        .DialogTitle = "Open RTF File"
        .CancelError = False
        ' Set the options (flags) for the dialog box.
        .Flags = cdlOFNHideReadOnly
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "Rich Text Files (*.rtf)|*.rtf"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    RichTextBox1.LoadFile sFile, rtfRTF
    lblStatus = "~ File Loaded, Hasn't Changed ~"
    RichTextBox1.SetFocus
End Sub

Private Sub picOpen_GotFocus()
    picOpen.BorderStyle = 1
End Sub

Private Sub picOpen_LostFocus()
    picOpen.BorderStyle = 0
End Sub

Private Sub picSave_Click()
    On Error GoTo err
    Dim sFile As String
        With dlgCommonDialog
            .DialogTitle = "Save RTF File"
            .CancelError = True
            ' Set the options (flags) for the dialog box.
            .Flags = cdlOFNHideReadOnly
            'ToDo: set the flags and attributes of the common dialog control
            .Filter = "Rich Text Files (*.rtf)|*.rtf"
            .ShowSave
            If Len(.FileName) = 0 Then
                Exit Sub
            End If
            sFile = .FileName
        End With
        RichTextBox1.SaveFile sFile, rtfRTF
        lblStatus.Caption = "~ File Has Been Saved ~"
        RichTextBox1.SetFocus
    Exit Sub
err:
    Exit Sub
End Sub

Private Sub picSave_GotFocus()
    picSave.BorderStyle = 1
End Sub

Private Sub picSave_LostFocus()
    picSave.BorderStyle = 0
End Sub

Private Sub RichTextBox1_Change()
    RichTextBox1.BackColor = &HFFFFFF
    lblStatus.Caption = "~ Current File Not Saved ~"
End Sub

Private Sub RichTextBox1_GotFocus()
    RichTextBox1.BackColor = &HFFFFFF
End Sub

Private Sub RichTextBox1_LostFocus()
    RichTextBox1.BackColor = &HE0E0E0
End Sub
