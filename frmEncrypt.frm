VERSION 5.00
Begin VB.Form frmEncrypt 
   Caption         =   "Encrypt"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   Icon            =   "frmEncrypt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2925
   ScaleWidth      =   6870
   Begin VB.TextBox txtMessagecrypt 
      Height          =   975
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1800
      Width           =   4995
   End
   Begin VB.TextBox txtMessage 
      Height          =   1095
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   300
      Width           =   4995
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "&Encrypt"
      Height          =   435
      Left            =   5280
      TabIndex        =   1
      Top             =   2460
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Encrypted message: "
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   1560
      Width           =   4995
   End
   Begin VB.Label Label5 
      Caption         =   "Enter message here:"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   1515
   End
   Begin VB.Label Label2 
      Caption         =   "Type what you would like to encrypt and then Press the Encrypt button."
      Height          =   795
      Left            =   5160
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
End
Attribute VB_Name = "frmEncrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private itsConnectionString     As String

Dim Calgorithm  As String
Dim Cname       As String
Dim iEncryptKey As Integer

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


Private Sub cmdEncrypt_Click()
    prcGrabEachChar txtMessage.Text
End Sub

Private Sub Form_Load()
    Me.Width = 6990
    Me.Height = 3330
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

    rsCyphr.Close
    cnnConnection.Close
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


Sub prcGrabEachChar(message As String)
    Dim tempmessage As String
    Dim strlen As String, strabc As String, strNewChar As String
    Dim iKeyReturned As Integer, iCyphrdKey As Integer, itmp As Integer
    Dim strByte
    
    Label6 = "Encrypted message: "
    If Trim(txtMessage.Text) = "" Or IsNull(txtMessage.Text) Then
        Label6 = "Encrypted message: Nothing was entered in the top text box."
    Else
        Label6 = "Encrypted message: Basic encryption:"
    End If
    
    strabc = ""
    
    While message <> ""
        iEncryptKey = 0
        strlen = Len(message) 'get the length of the string
        tempmessage = Left(message, 1) 'pick the first char off of the string
        
        iKeyReturned = fntReturnKey2(UCase(tempmessage)) 'get key from char
        If iKeyReturned = 600 Then
            strByte = "ее"
        Else
            itmp = iKeyReturned
            iCyphrdKey = fntGetNewKey(itmp, iEncryptKey) 'get the new key
            strByte = Hex(iCyphrdKey) 'convert new key to hex
        End If
        
        message = Right(message, strlen - 1) 'get the message - the char just evaluated
        strabc = strabc & strByte
    Wend
    
    txtMessagecrypt.Text = strabc 'Hex(iEncryptKey) & strabc & Hex(iEncryptKey)
End Sub

Function fntReturnKey2(iChar As String) As Integer

    Select Case iChar
    
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
    
        Case ("*"): fntReturnKey2 = 106
        Case ("+"): fntReturnKey2 = 107
        Case ("-"): fntReturnKey2 = 109
        Case ("."): fntReturnKey2 = 110
        Case ("/"): fntReturnKey2 = 111
        Case (","): fntReturnKey2 = 44
        
        Case ("!"): fntReturnKey2 = 33
        Case ("#"): fntReturnKey2 = 35
        Case ("$"): fntReturnKey2 = 36
        Case ("%"): fntReturnKey2 = 37
        Case ("&"): fntReturnKey2 = 38
        Case ("'"): fntReturnKey2 = 39
        Case ("("): fntReturnKey2 = 40
        Case (")"): fntReturnKey2 = 41
        Case (":"): fntReturnKey2 = 58
        Case (";"): fntReturnKey2 = 59
        Case ("<"): fntReturnKey2 = 60
        Case ("="): fntReturnKey2 = 61
        Case (">"): fntReturnKey2 = 62
        Case ("?"): fntReturnKey2 = 63
        Case ("@"): fntReturnKey2 = 64
        
        Case ("["): fntReturnKey2 = 91
        Case ("\"): fntReturnKey2 = 92
        Case ("]"): fntReturnKey2 = 93
        Case ("^"): fntReturnKey2 = 94
        Case ("_"): fntReturnKey2 = 95
        Case ("`"): fntReturnKey2 = 96
        Case ("{"): fntReturnKey2 = 123
        Case ("|"): fntReturnKey2 = 124
        Case ("}"): fntReturnKey2 = 125
        Case ("~"): fntReturnKey2 = 126
        Case (vbLf): fntReturnKey2 = 127
        Case (vbKeyTab): fntReturnKey2 = 9
        Case Else:  fntReturnKey2 = 600
        
    End Select
    
End Function

Function fntGetNewKey(iOldKey As Integer, idisplace As Integer) As Integer
    
    While idisplace <> 0
        
        Select Case iOldKey
            Case (32):  iOldKey = 48
            Case (57):  iOldKey = 67
            Case (90):  iOldKey = 96
            Case (111): iOldKey = 32
            Case Else:  iOldKey = iOldKey + 1
            
        End Select
        
        idisplace = idisplace - 1
    Wend
    fntGetNewKey = iOldKey
End Function
