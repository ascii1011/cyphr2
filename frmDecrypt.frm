VERSION 5.00
Begin VB.Form frmDecrypt 
   Caption         =   "Decrypt"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   Icon            =   "frmDecrypt.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2985
   ScaleWidth      =   6870
   Begin VB.TextBox txtMessage 
      Height          =   1095
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   240
      Width           =   4995
   End
   Begin VB.TextBox txtMessagecrypt 
      Height          =   1095
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1740
      Width           =   4995
   End
   Begin VB.CommandButton cmdDecipher 
      Caption         =   "&Decipher"
      Height          =   495
      Left            =   5220
      TabIndex        =   0
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Enter Encrypted message here:"
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4995
   End
   Begin VB.Label Label6 
      Caption         =   "Decrypted message"
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   1500
      Width           =   4995
   End
   Begin VB.Label Label3 
      Caption         =   "Browse to find a file.  Once you have found a file press the decipher button"
      Height          =   795
      Left            =   5160
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmDecrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub prcGrabEachHex(message As String)
    Dim tempmessage As String
    Dim strlen As Integer, strabc As String, strNewChar As String
    Dim strKeyReturned As String, iCyphrdKey As Integer, itmp As Integer
    Dim strByte
    Dim strEncryptKey As String
    Dim itemp As Integer
    
    Label6 = "Encrypted message: "
    
    If Trim(txtMessage.Text) = "" Or IsNull(txtMessage.Text) Then
        Label6 = "Decrypted message: Nothing was entered in the top text box."
    Else
        Label6 = "Decrypted message: Basic encryption:"
    End If
    
    strlen = "0"
    strabc = ""
    strEncryptKey = ""
    tempmessage = ""
    strKeyReturned = ""
    
    While message <> ""
        strlen = Len(message) 'get the length of the string
        itemp = strlen Mod 2
        If itemp <> 0 Then
            MsgBox "There was an error with the data you entered"
            message = ""
            Exit Sub
        End If
        
        tempmessage = Left(message, 2) 'pick the first char off of the string
        
        strKeyReturned = prcReturnMessage(tempmessage, strEncryptKey)
        
        message = Right(message, strlen - 2) 'get the message - the char just evaluated
        
        strabc = strabc & strKeyReturned 'strByte
    Wend
        
    txtMessagecrypt.Text = strabc
End Sub


Private Sub cmdDecipher_Click()
    prcGrabEachHex txtMessage.Text
End Sub

Function prcReturnMessage(strCharacter As String, strKey As String) As String
    Select Case strCharacter
        Case ("41"):    prcReturnMessage = "a"
        Case ("42"):    prcReturnMessage = "b"
        Case ("43"):    prcReturnMessage = "c"
        Case ("44"):    prcReturnMessage = "d"
        Case ("45"):    prcReturnMessage = "e"
        Case ("46"):    prcReturnMessage = "f"
        Case ("47"):    prcReturnMessage = "g"
        Case ("48"):    prcReturnMessage = "h"
        Case ("49"):    prcReturnMessage = "i"
        Case ("4A"):    prcReturnMessage = "j"
        Case ("4B"):    prcReturnMessage = "k"
        Case ("4C"):    prcReturnMessage = "l"
        Case ("4D"):    prcReturnMessage = "m"
        Case ("4E"):    prcReturnMessage = "n"
        Case ("4F"):    prcReturnMessage = "o"
        Case ("50"):    prcReturnMessage = "p"
        Case ("51"):    prcReturnMessage = "q"
        Case ("52"):    prcReturnMessage = "r"
        Case ("53"):    prcReturnMessage = "s"
        Case ("54"):    prcReturnMessage = "t"
        Case ("55"):    prcReturnMessage = "u"
        Case ("56"):    prcReturnMessage = "v"
        Case ("57"):    prcReturnMessage = "w"
        Case ("58"):    prcReturnMessage = "x"
        Case ("59"):    prcReturnMessage = "y"
        Case ("5A"):    prcReturnMessage = "z"
        Case ("20"):    prcReturnMessage = " "
        Case ("6A"):    prcReturnMessage = "*"
        Case ("6B"):    prcReturnMessage = "+"
        Case ("6D"):    prcReturnMessage = "-"
        Case ("6E"):    prcReturnMessage = "."
        Case ("6F"):    prcReturnMessage = "/"
        Case ("2C"):    prcReturnMessage = ","
        
        Case ("31"):    prcReturnMessage = "1"
        Case ("32"):    prcReturnMessage = "2"
        Case ("33"):    prcReturnMessage = "3"
        Case ("34"):    prcReturnMessage = "4"
        Case ("35"):    prcReturnMessage = "5"
        Case ("36"):    prcReturnMessage = "6"
        Case ("37"):    prcReturnMessage = "7"
        Case ("38"):    prcReturnMessage = "8"
        Case ("39"):    prcReturnMessage = "9"
        Case ("30"):    prcReturnMessage = "0"
        
        Case ("21"):    prcReturnMessage = "!"
        Case ("23"):    prcReturnMessage = "#"
        Case ("24"):    prcReturnMessage = "$"
        Case ("25"):    prcReturnMessage = "%"
        Case ("26"):    prcReturnMessage = "&"
        Case ("27"):    prcReturnMessage = "'"
        Case ("28"):    prcReturnMessage = "("
        Case ("29"):    prcReturnMessage = ")"
        Case ("3A"):    prcReturnMessage = ":"
        Case ("3B"):    prcReturnMessage = ";"
        Case ("3C"):    prcReturnMessage = "<"
        Case ("3D"):    prcReturnMessage = "="
        Case ("3E"):    prcReturnMessage = ">"
        Case ("3F"):    prcReturnMessage = "?"
        Case ("40"):    prcReturnMessage = "@"
        
        Case ("5B"):    prcReturnMessage = "["
        Case ("5C"):    prcReturnMessage = "\"
        Case ("5D"):    prcReturnMessage = "]"
        Case ("5E"):    prcReturnMessage = "^"
        Case ("7B"):    prcReturnMessage = "{"
        Case ("7C"):    prcReturnMessage = "|"
        Case ("7D"):    prcReturnMessage = "}"
        Case ("7E"):    prcReturnMessage = "~"
        Case ("5F"):    prcReturnMessage = "_"
        Case ("60"):    prcReturnMessage = "`"
        
        Case ("7F"):    prcReturnMessage = vbNewLine
        Case ("00"):    prcReturnMessage = vbNewLine
        Case Else:      prcReturnMessage = ""
     End Select
End Function

Private Sub Form_Load()
    Me.Width = 6990
    Me.Height = 3330
End Sub
