VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sniper.ps"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Access ≈·Ï Excel „‰"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sniper.ps  ’„Ì„ "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   3480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XL As ADODB.Connection
Dim RS As ADODB.Recordset
Private Sub Command1_Click()
If RS.State = adStateOpen Then RS.Close
RS.Open "Insert Into Table1 In'" & App.Path & "\Access.mdb" & "'" & " Select * From [Sheet1$]", XL, adOpenDynamic, adLockOptimistic
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation, "Sniper.ps"
End Sub

Private Sub Form_Load()
Set XL = New ADODB.Connection
Set RS = New ADODB.Recordset
XL.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Excel.xls" & ";Extended Properties=Excel 8.0;"
XL.CursorLocation = adUseClient
End Sub
