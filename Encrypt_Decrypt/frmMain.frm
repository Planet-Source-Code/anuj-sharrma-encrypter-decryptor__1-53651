VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdMain 
      Left            =   255
      Top             =   225
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   390
      Left            =   6690
      TabIndex        =   10
      Top             =   1725
      Width           =   1140
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "&Decrypt"
      Height          =   390
      Left            =   5520
      TabIndex        =   9
      Top             =   1725
      Width           =   1140
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "&Encrypt"
      Height          =   390
      Left            =   4335
      TabIndex        =   8
      Top             =   1725
      Width           =   1140
   End
   Begin VB.CommandButton cmdFileToEncrypt 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7395
      TabIndex        =   5
      Top             =   735
      Width           =   345
   End
   Begin VB.TextBox txtDestination 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   330
      Left            =   2670
      TabIndex        =   3
      Top             =   1125
      Width           =   4680
   End
   Begin VB.TextBox txtSource 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   330
      Left            =   2670
      TabIndex        =   2
      Top             =   750
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      Height          =   1125
      Left            =   60
      TabIndex        =   1
      Top             =   510
      Width           =   7770
      Begin VB.CommandButton cmdFileForEncrypt 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7320
         TabIndex        =   6
         Top             =   600
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Select file destnation pathh"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   75
         TabIndex        =   7
         Top             =   690
         Width           =   2160
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Select a file to encrypt/decrypt"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   255
         Left            =   90
         TabIndex        =   4
         Top             =   270
         Width           =   2445
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   2265
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "5:31 ANUJ"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
            Object.ToolTipText     =   "Developed By: Anuj sharma"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "05/08/2004"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Encrpetr/Decrypter"
      BeginProperty Font 
         Name            =   "Garamond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   360
      Left            =   2580
      TabIndex        =   0
      Top             =   150
      Width           =   2715
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------
'---------------------Developer    Anuj sharma------------------
'---------------------Date         08/May/2004------------------
'---------------------User         Begineer/Intermid------------
'---------------------E-mail       Anujsharrma@yahoo.com--------
'---------------------------------------------------------------
Option Explicit

Private Sub cmdDecrypt_Click()
On Error GoTo Err_Handler
Dim sFileName As String
Dim iFileNo As Integer
Dim sData As String
Dim sDecryptData  As String
Dim sDestinationFile As String
Dim iDestinationFileNo As Integer
    
    txtSource.Enabled = False
    txtDestination.Enabled = False
    cmdEncrypt.Enabled = False
    cmdDecrypt.Enabled = False
    
    sDestinationFile = txtDestination.Text
    iDestinationFileNo = FreeFile
    
    Open sDestinationFile For Output As #iDestinationFileNo
        sFileName = txtSource.Text
        iFileNo = FreeFile
        Open sFileName For Input As #iFileNo
            Do While Not EOF(iFileNo)
                Line Input #iFileNo, sData
                sDecryptData = TextToDecrypt(sData)
                Print #iDestinationFileNo, sDecryptData
                DoEvents
            Loop
        Close #iDestinationFileNo
        MsgBox "Decryption compleated.", vbInformation + vbOKOnly, App.Title
    Close #iFileNo
    
    txtSource.Enabled = True
    txtDestination.Enabled = True
    cmdEncrypt.Enabled = True
    cmdDecrypt.Enabled = True
    txtSource.Text = ""
    txtDestination.Text = ""
Exit Sub
Err_Handler:
    MsgBox Err.Description, vbInformation + vbOKOnly, App.Title
End Sub

Private Sub cmdEncrypt_Click()
On Error GoTo Err_Handler
Dim sFileName As String
Dim iFileNo As Integer
Dim sData As String
Dim sEncryptData As String
Dim sDestinationFile As String
Dim iDestinationFileNo As Integer
    
    txtSource.Enabled = False
    txtDestination.Enabled = False
    cmdEncrypt.Enabled = False
    cmdDecrypt.Enabled = False
    
    sDestinationFile = txtDestination.Text
    iDestinationFileNo = FreeFile
    
    Open sDestinationFile For Output As #iDestinationFileNo
        sFileName = txtSource.Text
        iFileNo = FreeFile
        Open sFileName For Input As #iFileNo
            Do While Not EOF(iFileNo)
                Line Input #iFileNo, sData
                sEncryptData = TextToEncrypt(sData)
                Print #iDestinationFileNo, sEncryptData
                DoEvents
            Loop
        Close #iDestinationFileNo
        MsgBox "Encryption compleated.", vbInformation + vbOKOnly, App.Title
    Close #iFileNo
    
    txtSource.Enabled = True
    txtDestination.Enabled = True
    cmdEncrypt.Enabled = True
    cmdDecrypt.Enabled = True
    txtSource.Text = ""
    txtDestination.Text = ""
Exit Sub
Err_Handler:
    MsgBox Err.Description, vbInformation + vbOKOnly, App.Title
End Sub

Private Function TextToEncrypt(ByVal InputString As String) As String
On Error GoTo Err_Handler
Dim iCharLen As Integer
Dim sMid As String * 1
Dim sReturnString As String
Dim sTotalValue As String
    
    sReturnString = ""
    For iCharLen = 1 To Len(InputString)
        sMid = Mid(InputString, iCharLen, 1)
                sReturnString = sReturnString & Chr(Asc(sMid) + 35)
                DoEvents
    Next iCharLen
 TextToEncrypt = sReturnString
Exit Function
Err_Handler:
    MsgBox Err.Description, vbInformation + vbOKOnly, App.Title
End Function

Private Function TextToDecrypt(ByVal InputString As String) As String
On Error GoTo Err_Handler
Dim iCharLen As Integer
Dim sMid As String * 1
Dim sReturnString As String
Dim sTotalValue As String
    
    sReturnString = ""
      
    For iCharLen = 1 To Len(InputString)
        sMid = Mid(InputString, iCharLen, 1)
                sReturnString = sReturnString & Chr(Asc(sMid) - 35)
                DoEvents
    Next iCharLen
 TextToDecrypt = sReturnString
Exit Function
Err_Handler:
    MsgBox Err.Description, vbInformation + vbOKOnly, App.Title
End Function

Private Sub cmdExit_Click()
    Unload Me
    End
End Sub

Private Sub cmdFileForEncrypt_Click()
Dim sDestination As String

    With cdMain
        .DialogTitle = App.ProductName & "-Select a file."
        .Filter = "All files."
        .ShowSave
    End With
    sDestination = Trim(cdMain.FileName)
    txtDestination.Text = sDestination
    
End Sub

Private Sub cmdFileToEncrypt_Click()
Dim sSourcePath As String

    With cdMain
        .DialogTitle = App.ProductName & "- Select a file."
        .Filter = "All files"
        .ShowOpen
    End With
    sSourcePath = Trim(cdMain.FileName)
    txtSource.Text = sSourcePath
    
End Sub
