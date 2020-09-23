VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Form"
   ClientHeight    =   6195
   ClientLeft      =   1920
   ClientTop       =   3870
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefreshList 
      Caption         =   "Refresh List"
      Height          =   375
      Left            =   7320
      TabIndex        =   23
      Top             =   1680
      Width           =   3615
   End
   Begin VB.CommandButton cmdClearList 
      Caption         =   "Clear List"
      Height          =   375
      Left            =   7320
      TabIndex        =   22
      Top             =   1320
      Width           =   3615
   End
   Begin VB.CommandButton cmdClearFile 
      Caption         =   "Clear File"
      Height          =   375
      Left            =   7320
      TabIndex        =   21
      Top             =   120
      Width           =   3615
   End
   Begin VB.CommandButton cmdRemoveFile 
      Caption         =   "Remove File"
      Height          =   375
      Left            =   7320
      TabIndex        =   20
      Top             =   480
      Width           =   3615
   End
   Begin VB.CommandButton cmdAddFile 
      Caption         =   "Add File"
      Height          =   375
      Left            =   7320
      TabIndex        =   19
      Top             =   840
      Width           =   3615
   End
   Begin VB.ListBox lstFiles 
      Height          =   3375
      ItemData        =   "frmTest.frx":0000
      Left            =   7320
      List            =   "frmTest.frx":0002
      TabIndex        =   18
      Top             =   2640
      Width           =   3735
   End
   Begin VB.CommandButton cmdDecodeBase64 
      Caption         =   "Decode Base64"
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton cmdEncodeUUEncode 
      Caption         =   "Encode UUEncode"
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmddecode 
      Caption         =   "Decode UUEncode"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton cmdEncode 
      Caption         =   "Encode Base64"
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   1560
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog dlgEncode 
      Left            =   6720
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Orientation     =   2
   End
   Begin MSWinsockLib.Winsock socMail 
      Left            =   6240
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtSubject 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   4575
   End
   Begin VB.TextBox txtMessage 
      Height          =   2835
      Left            =   120
      TabIndex        =   5
      Top             =   2880
      Width           =   7095
   End
   Begin VB.TextBox txtTo 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   4575
   End
   Begin VB.TextBox txtMailFrom 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   4575
   End
   Begin VB.TextBox txtPort 
      Height          =   315
      Left            =   4920
      TabIndex        =   2
      Text            =   "25"
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox txtSMTPHostname 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4575
   End
   Begin VB.CommandButton cmdSendMail 
      Caption         =   "SendMail"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5760
      Width           =   7095
   End
   Begin VB.Label lblBody 
      Caption         =   "Message"
      Height          =   280
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   3975
   End
   Begin VB.Label lblPort 
      Caption         =   "Port"
      Height          =   285
      Left            =   4920
      TabIndex        =   15
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblSMTPHostname 
      Caption         =   "SMTP Hostname"
      Height          =   285
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label lblFrom 
      Caption         =   "Sender Address:"
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label lblSubject 
      Caption         =   "Subject"
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label lblTo 
      Caption         =   "Receiver Address"
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label lblFiles 
      Caption         =   "Attached Files"
      Height          =   285
      Left            =   7320
      TabIndex        =   17
      Top             =   2400
      Width           =   3735
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents myclass As clsEmail
Attribute myclass.VB_VarHelpID = -1




Private Sub cmdAddFile_Click()

On Error GoTo Erh:
dlgEncode.Filter = "All Files (*.*)|*.*"
dlgEncode.ShowOpen

myclass.AddFiles dlgEncode.FileName
lstFiles.AddItem dlgEncode.FileName

Exit Sub

Erh:


End Sub

Private Sub cmdClearFile_Click()

'Clear the queue
myclass.ClearFiles

'Clear the file list
lstFiles.Clear


End Sub

Private Sub cmdClearList_Click()

'Clear the list. (Not the actual queue)
lstFiles.Clear

End Sub

Private Sub cmddecode_Click()
On Error GoTo Erh:
dlgEncode.Filter = "All Files (*.*)|*.*"
dlgEncode.ShowOpen

myclass.UUDecode dlgEncode.FileName

Exit Sub

Erh:

End Sub

Private Sub cmdDecodeBase64_Click()

On Error GoTo Erh:
dlgEncode.Filter = "All Files (*.*)|*.*"
dlgEncode.ShowOpen

myclass.Base64Decode dlgEncode.FileName

Exit Sub

Erh:
End Sub

Private Sub cmdEncode_Click()

On Error GoTo Erh:
dlgEncode.Filter = "All Files (*.*)|*.*"
dlgEncode.ShowOpen

myclass.Base64Encode dlgEncode.FileName, False

Exit Sub

Erh:

End Sub

Private Sub cmdEncodeUUEncode_Click()



On Error GoTo Erh:
dlgEncode.Filter = "All Files (*.*)|*.*"
dlgEncode.ShowOpen

myclass.UUEncode dlgEncode.FileName, False

Exit Sub

Erh:

End Sub







Private Sub cmdRefreshList_Click()



'Clear it first. This is better as to match up the index between list index n queue index
lstFiles.Clear

'Refresh/Obtain the list of the queue
myclass.ListFiles

End Sub

Private Sub cmdRemoveFile_Click()

If lstFiles.ListCount <> 0 Then

    'Remove the file from the queue
    
    myclass.RemoveFile lstFiles.ListIndex
    
    'Remove the file from the queue
    lstFiles.RemoveItem lstFiles.ListIndex

End If

End Sub

Private Sub cmdSendMail_Click()

With myclass
.SMTPHostname = txtSMTPHostname
.SMTPPort = txtPort
.SenderAddress = txtMailFrom
.RecepientAddress = txtTo
.Subject = txtSubject
.MessageBody = txtMessage
.SMTPSocket = socMail
.SendMail

End With




End Sub

Private Sub Form_Load()

Set myclass = New clsEmail


End Sub

Private Sub Form_Terminate()
    myclass.Terminate
End Sub


Private Sub myclass_FilesInQueue(strFilePath As String, intIndex As Integer)

Dim strFileTitle    As String
Dim intCounter      As Integer

For intCounter = Len(strFilePath) To 1 Step -1
    If Mid$(strFilePath, intCounter, 1) = "\" Then
    
        'Or also can use the optional index by the event(Not the case if it is already cleared ;) )
        lstFiles.AddItem Mid$(strFilePath, intCounter + 1)
        
        Exit For
        
    End If
Next intCounter





End Sub

Private Sub myclass_SentSuccessful()
    MsgBox "Sent successfully"
End Sub

Private Sub myclass_UUCodeError(intErrorNumber As Integer, strDescription As String)
    MsgBox "An Error"
End Sub


