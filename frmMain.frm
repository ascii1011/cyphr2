VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Cyphr"
   ClientHeight    =   6195
   ClientLeft      =   1815
   ClientTop       =   2040
   ClientWidth     =   10095
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin MSComctlLib.StatusBar sbInfoDisplay 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5820
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin MSComctlLib.ImageList ilToolbar 
      Left            =   180
      Top             =   660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2BF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5BDFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5BF0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C01E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C178
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C2D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C3E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C4F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":60570
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":60682
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":646FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6480E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":68888
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6899A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6CA14
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6CB26
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6CC38
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6DEBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":71F34
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":72386
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":727D8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ilToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   23
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "New"
            Object.ToolTipText     =   "New record"
            Object.Tag             =   "mnuEdit_ADDNEW"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Save"
            Object.ToolTipText     =   "Save current record"
            Object.Tag             =   "mnuEdit_SAVE"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Print"
            Object.ToolTipText     =   "Print record"
            Object.Tag             =   "mnuFile_Print"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            Object.Tag             =   "mnuEdit_REFRESH"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Employee List"
            Object.ToolTipText     =   "Employee Listing"
            Object.Tag             =   "mnuEdit_Employee"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "First"
            Object.ToolTipText     =   "Employee Listing"
            Object.Tag             =   "mnuEdit_EMPLOYEE"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Previous"
            Object.ToolTipText     =   "Move to the previous record"
            Object.Tag             =   "mnuEdit_PREV"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Next"
            Object.ToolTipText     =   "Move to the next record"
            Object.Tag             =   "mnuEdit_NEXT"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "Last"
            Object.ToolTipText     =   "Move to the last record"
            Object.Tag             =   "mnuEdit_LAST"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "TechnicianList"
            Object.ToolTipText     =   "Technician List"
            Object.Tag             =   "mnuEdit_TechList"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "DispositionList"
            Object.ToolTipText     =   "Disposition List"
            Object.Tag             =   "mnuEdit_DispositionList"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "ConditionReceivedList"
            Object.ToolTipText     =   "Condition Received List"
            Object.Tag             =   "mnuEdit_ConditionReceivedList"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Void"
            Object.ToolTipText     =   "Void record"
            Object.Tag             =   "mnuEdit_VOID"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Table"
            Object.ToolTipText     =   "Table view"
            Object.Tag             =   "mnuHELP_TABLEVIEW"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Find"
            Object.ToolTipText     =   "Launch finder"
            Object.Tag             =   "mnuHELP_LAUNCHFIND"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Help"
            Object.ToolTipText     =   "Launch help file"
            Object.Tag             =   "mnuHELP_START"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLogout 
         Caption         =   "&Logout"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuspace3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintSetup 
         Caption         =   "Printer &Setup"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuspace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit_ADDNEW 
         Caption         =   "&Add"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEdit_SAVE 
         Caption         =   "&Save"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEdit_VOID 
         Caption         =   "&Void"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditSpace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_REFRESH 
         Caption         =   "Ref&resh"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditSpaceI 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_FIND 
         Caption         =   "Fin&d"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuEdit_Encryption 
         Caption         =   "&Encryption"
      End
      Begin VB.Menu mnuEdit_Decryption 
         Caption         =   "&Decryption"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuArrangeWindows 
         Caption         =   "&Arrange Windows"
      End
      Begin VB.Menu mnuWindowspace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAboutthisApplication 
         Caption         =   "&About this Application"
      End
      Begin VB.Menu mnuViewSplashScreen 
         Caption         =   "&view splash screen"
      End
      Begin VB.Menu mnuAbilities 
         Caption         =   "&Abilities"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strUser              As String
Public strPassword          As String
Public strServerConn        As String
Public strDatabaseConn      As String
Public cncyphr               As New ADODB.Connection

Public MyDateTextObject     As PictureBox
Public MyDateLabel          As Label
Public strcalender          As Integer
Public UserName             As String



Private Sub MDIForm_Load()
Dim cmdSysCommand   As New ADODB.Command
Dim parParameter    As New ADODB.Parameter
Dim rsSysValues     As New ADODB.Recordset

On Error Resume Next
    Dim ServerName As String, DatabaseName As String, _
    Password As String
   
    Me.Height = 6900
    Me.Width = 11300
   
    'Put text box values into connection variables.
    ServerName = "chips"
    DatabaseName = "dialer"
    UserName = "sa"
    Password = "123456"
    
    ' Specify the OLE DB provider.
    ''frmMain.cnDIAL.Provider = "sqloledb"
    ''frmMain.cnDIAL.CursorLocation = adUseClient
    
    ' Set SQLOLEDB connection properties.
    ''frmMain.cnDIAL.Properties("Data Source").Value = ServerName
    ''frmMain.cnDIAL.Properties("Initial Catalog").Value = DatabaseName
    ''frmMain.cnDIAL.Properties("User ID").Value = UserName
    ''frmMain.cnDIAL.Properties("Password").Value = Password

    ' Open the database.
    ''frmMain.cnDIAL.Open
    
    frmSplash.Show
    ''frmDialerEmployee.Show

End Sub

Function usrDelayProgram(nSeconds As Integer)
Dim nStoptime As Single

nStoptime = Timer + nSeconds

Do While Timer <= nStoptime: DoEvents: Loop

End Function

Private Sub mnuAboutthisApplication_Click()
    frmAbout.Show
End Sub

Private Sub mnuArrangeWindows_Click()
   
    MousePointer = vbHourglass
    Arrange vbArrangeIcons
    MousePointer = vbDefault

End Sub

Private Sub mnuEdit_Decryption_Click()
    frmDecrypt.Show
End Sub

Private Sub mnuEdit_Encryption_Click()
    frmEncrypt.Show
End Sub

Private Sub mnuEdit_REFRESH_Click()
'Form1.Show
'frmOrderAcknowledgement.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub




Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Response

On Error GoTo errHandle
    
    Select Case (Button.Key)
        Case ("First"):     ActiveForm.prcMoveFirst
        Case ("Next"):      ActiveForm.prcMoveNext
        Case ("Previous"):  ActiveForm.prcMovePrevious
        Case ("Last"):      ActiveForm.prcMoveLast
    End Select
    Exit Sub

errHandle:
    Select Case (Err.Number)
        Case (91):
            Resume Next
        Case (438):
            MsgBox ActiveForm.Caption & " does not support this operation", vbInformation, "SYSTEM"
        Case Else
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Login run time error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select

End Sub



Private Sub mnuTileHorizontal_Click()
    MousePointer = vbHourglass
    Arrange vbTileHorizontal
    MousePointer = vbDefault
End Sub

Private Sub mnuTileVertical_Click()
   MousePointer = vbHourglass
    Arrange vbTileVertical
    MousePointer = vbDefault

End Sub

Private Sub mnuViewSplashScreen_Click()
    frmSplash.Show
End Sub

Private Sub tbToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Response

On Error GoTo errHandle
    
    Select Case (Button.Key)
        Case ("New"):       ActiveForm.prcAddNew
        Case ("Save"):      ActiveForm.prcSave
        Case ("First"):     ActiveForm.prcMoveFirst
        Case ("Next"):      ActiveForm.prcMoveNext
        Case ("Previous"):  ActiveForm.prcMovePrevious
        Case ("Last"):      ActiveForm.prcMoveLast
        Case ("Print"):     ActiveForm.prcPrint
        Case ("Refresh"):   ActiveForm.prcrefresh
        Case ("Void"):      ActiveForm.prcVoid
    End Select
    Exit Sub

errHandle:
    Select Case (Err.Number)
        Case (91):
            Resume Next
        Case (438):
            MsgBox ActiveForm.Caption & " does not support this operation", vbInformation, "SYSTEM"
        Case Else
            Response = MsgBox(Err.Description & vbNewLine & "Try again?", vbExclamation + vbYesNo, "Login run time error")
            If Response = vbYes Then Resume Else Exit Sub
    End Select
End Sub

