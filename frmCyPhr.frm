VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCyPhr 
   Caption         =   "CyPhr"
   ClientHeight    =   3360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   Icon            =   "frmCyPhr.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1515
      Left            =   120
      Picture         =   "frmCyPhr.frx":0442
      ScaleHeight     =   1455
      ScaleWidth      =   1395
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdDecipher 
      Caption         =   "&Decrypt a message"
      Height          =   495
      Left            =   4800
      TabIndex        =   1
      Top             =   2700
      Width           =   1575
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "&Encrypt a Message"
      Height          =   495
      Left            =   4800
      TabIndex        =   0
      Top             =   1680
      Width           =   1575
   End
   Begin MSMask.MaskEdBox mskDate 
      Height          =   285
      Left            =   5580
      TabIndex        =   3
      Top             =   60
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   8
      Mask            =   "##/##/##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label4 
      Caption         =   "Today's Date:"
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   60
      Width           =   1035
   End
End
Attribute VB_Name = "frmCyPhr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDecipher_Click()
    frmDecrypt.Show
End Sub

Private Sub cmdEncrypt_Click()
    frmEncrypt.Show
End Sub

Private Sub Form_Load()
    
    mskDate.Text = Format(Now, "mm/dd/yy")
End Sub
