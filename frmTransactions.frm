VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmTransactions 
   Caption         =   "CyPhr"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMessagecrypt 
      Enabled         =   0   'False
      Height          =   1995
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   2820
      Width           =   4995
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   6900
      TabIndex        =   11
      Top             =   1980
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   1515
      Left            =   5160
      Picture         =   "frmTransactions.frx":0000
      ScaleHeight     =   1455
      ScaleWidth      =   1395
      TabIndex        =   8
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   5160
      TabIndex        =   1
      Top             =   780
      Width           =   1755
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   375
      Left            =   7020
      TabIndex        =   2
      Top             =   780
      Width           =   1215
   End
   Begin VB.TextBox txtMessage 
      Height          =   2055
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   300
      Width           =   4995
   End
   Begin VB.CommandButton cmdDecipher 
      Caption         =   "&Decipher"
      Height          =   495
      Left            =   6960
      TabIndex        =   5
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "&Encrypt"
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin MSMask.MaskEdBox mskDate 
      Height          =   285
      Left            =   6300
      TabIndex        =   9
      Top             =   60
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label6 
      Caption         =   "Encrypted message"
      Height          =   195
      Left            =   60
      TabIndex        =   14
      Top             =   2580
      Width           =   1515
   End
   Begin VB.Label Label5 
      Caption         =   "Enter message here:"
      Height          =   195
      Left            =   60
      TabIndex        =   13
      Top             =   60
      Width           =   1515
   End
   Begin VB.Label Label4 
      Caption         =   "Today's Date:"
      Height          =   255
      Left            =   5160
      TabIndex        =   10
      Top             =   60
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "Browse to find a file.  Once you have found a file press the decipher button"
      Height          =   795
      Left            =   5160
      TabIndex        =   7
      Top             =   4020
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Type what you would like to encrypt and then Press the Encrypt button."
      Height          =   795
      Left            =   5160
      TabIndex        =   6
      Top             =   2940
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Use the browse button to find a file "
      Height          =   255
      Left            =   5160
      TabIndex        =   4
      Top             =   420
      Width           =   2535
   End
End
Attribute VB_Name = "frmTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private itsConnectionString     As String

Dim Calgorithm  As String
Dim Cname       As String

Private Sub cmdDecipher_Click()

    Dim strMessage As String
    
    If txtMessage <> "" Then
        strMessage = txtMessage.Text
        prcEncrypt strMessage
    Else
        MsgBox "You must type something in the big " & vbNewLine & _
                "text box, for this to work"
    End If
End Sub

Private Sub Command2_Click()
    prcGrabEachChar txtMessage.Text
End Sub

Private Sub Form_Load()
    mskDate.Text = Format(Now, "mm/dd/yy")
    
    ' Open MDB
    Dim strPath As String
    strPath = App.Path
    If Right(strPath, 1) <> "/" Then
        strPath = strPath & "/"
    End If

    strPath = strPath & "cyphr.mdb"
    itsConnectionString = "DRIVER=Microsoft Access Driver (*.mdb);" & _
                          "DBQ=" & strPath & ";"
End Sub

Sub prcEncrypt(strMsg As String)
    Dim strAlg              As Integer
    Dim strReplacedMessage  As String
    
    strAlg = prcEncryptData(1)
    
    strReplacedMessage = Replace(strMsg, "a", "*&")
    strReplacedMessage = Replace(strReplacedMessage, "s", "@#")
    strReplacedMessage = Replace(strReplacedMessage, " ", "%$")
    MsgBox strReplacedMessage
    
End Sub


Function prcEncryptData(iEncryptType As Integer) As Integer
    Dim cnnConnection   As ADODB.Connection
    Dim strQry          As String

    'On Error GoTo VBError
    strQry = "SELECT * FROM cyphr_type " & _
                "WHERE cyphr_id = " & Trim$(CStr(iEncryptType))
                
    'On Error GoTo ADOError
    Set cnnConnection = New Connection
    cnnConnection.ConnectionString = itsConnectionString
    cnnConnection.Open
    
    Dim rsCyphr As Recordset
    Set rsCyphr = GetRecordSet(cnnConnection, strQry)
    
    If rsCyphr.EOF = True And rsCyphr.BOF = True Then
        ' Empty
    Else
        'logfname is the information coming from the db
        'itslogfname is where the information is being stored
        
        'dluser elements
        If Not IsNull(rsCyphr!cyphr_name) Then
            Cname = rsCyphr!cyphr_name
        End If
        If Not IsNull(rsCyphr!cyphr_algorithm) Then
            Calgorithm = rsCyphr!cyphr_algorithm
            prcEncryptData = rsCyphr!cyphr_algorithm
        End If
        
            
    End If
    'MsgBox Cname & ":" & Calgorithm
    'DisplayRecord
    'On Error GoTo VBError

    rsCyphr.Close
    
    cnnConnection.Close
'Done:
    ' Cleanup
    'Set rsCyphr = Nothing
    'Set cnnConnection = Nothing
'Exit Sub
'ADOError:   ' ADO error handler
    'DisplayADOErrors cnnConnection
'VBError:    ' Non-ADO error handler
    'DisplayVBError
    'GoTo Done
End Function


Private Function GetRecordSet(cnnConnection As ADODB.Connection, sQry As String) As ADODB.Recordset
    Dim rsCyphr As Recordset
    Set rsCyphr = New Recordset
    
    rsCyphr.CursorType = adOpenKeyset
    rsCyphr.LockType = adLockOptimistic
    rsCyphr.CursorLocation = adUseClient
    rsCyphr.Source = sQry
    Set rsCyphr.ActiveConnection = cnnConnection
    rsCyphr.Open

    Set GetRecordSet = rsCyphr
End Function

'letters to be manipulated
Function fncAlphaSet(iIndex As Integer) As Integer

    Dim strKeySet() As String
    
    strKeySet(65) = vbKeyA
    strKeySet(66) = vbKeyB
    strKeySet(67) = vbKeyC
    strKeySet(68) = vbKeyD
    strKeySet(69) = vbKeyE
    strKeySet(70) = vbKeyF
    strKeySet(71) = vbKeyG
    strKeySet(72) = vbKeyH
    strKeySet(73) = vbKeyI
    strKeySet(74) = vbKeyJ
    strKeySet(75) = vbKeyK
    strKeySet(76) = vbKeyL
    strKeySet(77) = vbKeyM
    strKeySet(78) = vbKeyN
    strKeySet(79) = vbKeyO
    strKeySet(80) = vbKeyP
    strKeySet(81) = vbKeyQ
    strKeySet(82) = vbKeyR
    strKeySet(83) = vbKeyS
    strKeySet(84) = vbKeyT
    strKeySet(85) = vbKeyU
    strKeySet(86) = vbKeyV
    strKeySet(87) = vbKeyW
    strKeySet(88) = vbKeyX
    strKeySet(89) = vbKeyY
    strKeySet(90) = vbKeyZ
    
    strKeySet(48) = vbKey0
    strKeySet(49) = vbKey1
    strKeySet(50) = vbKey2
    strKeySet(51) = vbKey3
    strKeySet(52) = vbKey4
    strKeySet(53) = vbKey5
    strKeySet(54) = vbKey6
    strKeySet(55) = vbKey7
    strKeySet(56) = vbKey8
    strKeySet(57) = vbKey9
    
    strKeySet(96) = vbKeyNumpad0
    strKeySet(97) = vbKeyNumpad1
    strKeySet(98) = vbKeyNumpad2
    strKeySet(99) = vbKeyNumpad3
    strKeySet(100) = vbKeyNumpad4
    strKeySet(101) = vbKeyNumpad5
    strKeySet(102) = vbKeyNumpad6
    strKeySet(103) = vbKeyNumpad7
    strKeySet(104) = vbKeyNumpad8
    strKeySet(105) = vbKeyNumpad9
    strKeySet(106) = vbKeyMultiply
    strKeySet(107) = vbKeyAdd
    strKeySet(108) = vbKeySeparator
    strKeySet(109) = vbKeySubtract
    strKeySet(110) = vbKeyDecimal
    strKeySet(111) = vbKeyDivide
    
    strKeySet(32) = vbKeySpace
    
End Function

Sub prcGrabEachChar(message As String)
    'Dim message As String,
    Dim tempmessage As String
    Dim strlen As String, strabc As String, strNewChar As String
    Dim iKeyReturned As Integer, iCyphrdKey As Integer, itmp As Integer
    
    'message = "9z/"
    
    strabc = "Basic encryption:" & vbNewLine
    
    While message <> ""
        strlen = Len(message) 'get the length of the string
        tempmessage = Left(message, 1) 'pick the first char off of the string
        'strabc = fntReturnChar
        iKeyReturned = fntReturnKey2(UCase(tempmessage)) 'get key from char
        itmp = iKeyReturned
        iCyphrdKey = fntGetNewKey(itmp) 'get the new key
        strNewChar = fntReturnChar(iCyphrdKey)
        message = Right(message, strlen - 1) 'get the message - the char just evaluated
        'strabc = strabc & strlen & ": " & iKeyReturned & "=" & tempmessage & " and " & iCyphrdKey & "=" & strNewChar & vbNewLine
        strabc = strabc & strNewChar
    Wend
    
    txtMessagecrypt.Text = strabc
End Sub


Function fntReturnChar(iChar As Integer) As String

    Select Case iChar
        Case (65): fntReturnChar = "A"
        Case (66): fntReturnChar = "B"
        Case (67): fntReturnChar = "C"
        Case (68): fntReturnChar = "D"
        Case (69): fntReturnChar = "E"
        Case (70): fntReturnChar = "F"
        Case (71): fntReturnChar = "G"
        Case (72): fntReturnChar = "h"
        Case (73): fntReturnChar = "i"
        Case (74): fntReturnChar = "J"
        Case (75): fntReturnChar = "K"
        Case (76): fntReturnChar = "L"
        Case (77): fntReturnChar = "M"
        Case (78): fntReturnChar = "N"
        Case (79): fntReturnChar = "O"
        Case (80): fntReturnChar = "P"
        Case (81): fntReturnChar = "Q"
        Case (82): fntReturnChar = "R"
        Case (83): fntReturnChar = "S"
        Case (84): fntReturnChar = "T"
        Case (85): fntReturnChar = "U"
        Case (86): fntReturnChar = "V"
        Case (87): fntReturnChar = "W"
        Case (88): fntReturnChar = "X"
        Case (89): fntReturnChar = "Y"
        Case (90): fntReturnChar = "Z"
    
        Case (48): fntReturnChar = "0"
        Case (49): fntReturnChar = "1"
        Case (50): fntReturnChar = "2"
        Case (51): fntReturnChar = "3"
        Case (52): fntReturnChar = "4"
        Case (53): fntReturnChar = "5"
        Case (54): fntReturnChar = "6"
        Case (55): fntReturnChar = "7"
        Case (56): fntReturnChar = "8"
        Case (57): fntReturnChar = "9"
    
        Case (96): fntReturnChar = "0"
        Case (97): fntReturnChar = "1"
        Case (98): fntReturnChar = "2"
        Case (99): fntReturnChar = "3"
        Case (100): fntReturnChar = "4"
        Case (101): fntReturnChar = "5"
        Case (102): fntReturnChar = "6"
        Case (103): fntReturnChar = "7"
        Case (104): fntReturnChar = "8"
        Case (105): fntReturnChar = "9"
        Case (106): fntReturnChar = "*"
        Case (107): fntReturnChar = "+"
        Case (108): fntReturnChar = " "
        Case (109): fntReturnChar = "-"
        Case (110): fntReturnChar = "."
        Case (111): fntReturnChar = "/"
        
        'Case (13): fntReturnChar = vbCharReturn
        Case (32): fntReturnChar = " "
       
    End Select
End Function


Function fntReturnKey(iChar As String) As Integer

    Select Case iChar
    
        Case (vbKeyReturn): fntReturnKey = 13
        Case (vbKeySpace): fntReturnKey = 32
                
        Case (vbKey0): fntReturnKey = 48
        Case (vbKey1): fntReturnKey = 49
        Case (vbKey2): fntReturnKey = 50
        Case (vbKey3): fntReturnKey = 51
        Case (vbKey4): fntReturnKey = 52
        Case (vbKey5): fntReturnKey = 53
        Case (vbKey6): fntReturnKey = 54
        Case (vbKey7): fntReturnKey = 55
        Case (vbKey8): fntReturnKey = 56
        Case (vbKey9): fntReturnKey = 57
    
        Case (vbKeyA): fntReturnKey = 65
        Case (vbKeyB): fntReturnKey = 66
        Case (vbKeyC): fntReturnKey = 67
        Case (vbKeyD): fntReturnKey = 68
        Case (vbKeyE): fntReturnKey = 69
        Case (vbKeyF): fntReturnKey = 70
        Case (vbKeyG): fntReturnKey = 71
        Case (vbKeyH): fntReturnKey = 72
        Case (vbKeyI): fntReturnKey = 73
        Case (vbKeyJ): fntReturnKey = 74
        Case (vbKeyK): fntReturnKey = 75
        Case (vbKeyL): fntReturnKey = 76
        Case (vbKeyM): fntReturnKey = 77
        Case (vbKeyN): fntReturnKey = 78
        Case (vbKeyO): fntReturnKey = 79
        Case (vbKeyP): fntReturnKey = 80
        Case (vbKeyQ): fntReturnKey = 81
        Case (vbKeyR): fntReturnKey = 82
        Case (vbKeyS): fntReturnKey = 83
        Case (vbKeyT): fntReturnKey = 84
        Case (vbKeyU): fntReturnKey = 85
        Case (vbKeyV): fntReturnKey = 86
        Case (vbKeyW): fntReturnKey = 87
        Case (vbKeyX): fntReturnKey = 88
        Case (vbKeyY): fntReturnKey = 89
        Case (vbKeyZ): fntReturnKey = 90
    
        Case (vbKeyNumpad0): fntReturnKey = 96
        Case (vbKeyNumpad1): fntReturnKey = 97
        Case (vbKeyNumpad2): fntReturnKey = 98
        Case (vbKeyNumpad3): fntReturnKey = 99
        Case (vbKeyNumpad4): fntReturnKey = 100
        Case (vbKeyNumpad5): fntReturnKey = 101
        Case (vbKeyNumpad6): fntReturnKey = 102
        Case (vbKeyNumpad7): fntReturnKey = 103
        Case (vbKeyNumpad8): fntReturnKey = 104
        Case (vbKeyNumpad9): fntReturnKey = 105
        Case (vbKeyMultiply): fntReturnKey = 106
        Case (vbKeyAdd): fntReturnKey = 107
        Case (vbKeySeparator): fntReturnKey = 108
        Case (vbKeySubtract): fntReturnKey = 109
        Case (vbKeyDecimal): fntReturnKey = 110
        Case (vbKeyDivide): fntReturnKey = 111
        
    End Select
End Function


Function fntReturnKey2(iChar As String) As Integer

    Select Case iChar
    
        'Case (vbKeyReturn): fntReturnKey2 = 13
        Case (" "): fntReturnKey2 = 32
                
        Case ("0"): fntReturnKey2 = 48
        Case ("1"): fntReturnKey2 = 49
        Case ("2"): fntReturnKey2 = 50
        Case ("3"): fntReturnKey2 = 51
        Case ("4"): fntReturnKey2 = 52
        Case ("5"): fntReturnKey2 = 53
        Case ("6"): fntReturnKey2 = 54
        Case ("7"): fntReturnKey2 = 55
        Case ("8"): fntReturnKey2 = 56
        Case ("9"): fntReturnKey2 = 57
    
        Case ("A"): fntReturnKey2 = 65
        Case ("B"): fntReturnKey2 = 66
        Case ("C"): fntReturnKey2 = 67
        Case ("D"): fntReturnKey2 = 68
        Case ("E"): fntReturnKey2 = 69
        Case ("F"): fntReturnKey2 = 70
        Case ("G"): fntReturnKey2 = 71
        Case ("H"): fntReturnKey2 = 72
        Case ("I"): fntReturnKey2 = 73
        Case ("J"): fntReturnKey2 = 74
        Case ("K"): fntReturnKey2 = 75
        Case ("L"): fntReturnKey2 = 76
        Case ("M"): fntReturnKey2 = 77
        Case ("N"): fntReturnKey2 = 78
        Case ("O"): fntReturnKey2 = 79
        Case ("P"): fntReturnKey2 = 80
        Case ("Q"): fntReturnKey2 = 81
        Case ("R"): fntReturnKey2 = 82
        Case ("S"): fntReturnKey2 = 83
        Case ("T"): fntReturnKey2 = 84
        Case ("U"): fntReturnKey2 = 85
        Case ("V"): fntReturnKey2 = 86
        Case ("W"): fntReturnKey2 = 87
        Case ("X"): fntReturnKey2 = 88
        Case ("Y"): fntReturnKey2 = 89
        Case ("Z"): fntReturnKey2 = 90
    
        'Case (vbKeyNumpad0): fntReturnKey2 = 96
        'Case (vbKeyNumpad1): fntReturnKey2 = 97
        'Case (vbKeyNumpad2): fntReturnKey2 = 98
        'Case (vbKeyNumpad3): fntReturnKey2 = 99
        'Case (vbKeyNumpad4): fntReturnKey2 = 100
        'Case (vbKeyNumpad5): fntReturnKey2 = 101
        'Case (vbKeyNumpad6): fntReturnKey2 = 102
        'Case (vbKeyNumpad7): fntReturnKey2 = 103
        'Case (vbKeyNumpad8): fntReturnKey2 = 104
        'Case (vbKeyNumpad9): fntReturnKey2 = 105
        Case ("*"): fntReturnKey2 = 106
        Case ("+"): fntReturnKey2 = 107
        'Case (vbKeySeparator): fntReturnKey2 = 108
        Case ("-"): fntReturnKey2 = 109
        Case ("."): fntReturnKey2 = 110
        Case ("/"): fntReturnKey2 = 111
        
    End Select
End Function

Function fntGetNewKey(iOldKey As Integer) As Integer
    'Dim icountUp As Integer
    Dim icountdown As Integer
    
    icountdown = 4
    'icountUp = 4
    
    While icountdown <> 0
        
        Select Case iOldKey
            Case (32):  iOldKey = 48
            Case (57):  iOldKey = 67
            Case (90):  iOldKey = 96
            Case (111): iOldKey = 32
            Case Else:  iOldKey = iOldKey + 1
                        'fntGetNewKey = iOldKey
                        'Exit Function
        End Select
        'iOldKey = iOldKey + 1
        icountdown = icountdown - 1
    Wend
    fntGetNewKey = iOldKey
End Function
