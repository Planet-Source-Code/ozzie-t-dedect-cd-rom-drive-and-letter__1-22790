VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "- CD-ROM Dedect - by Ozzie T"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3930
   Icon            =   "Start.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3930
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1780
      Left            =   840
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1750
      Left            =   120
      Top             =   120
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
      Min             =   40
   End
   Begin VB.Label Label1 
      Caption         =   "Please wait program collects the drives..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const DRIVE_CDROM = 5
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long


Private Sub Form_Load()
Timer2.Enabled = False
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 15
Timer2.Enabled = True

End Sub

Private Sub Dedect()
Dim rst As Integer
Dim rstStr As String
Dim Drives As String
Dim Count As Integer
Dim DriveLtrrs As String
Drives = Space(255)


ret& = GetLogicalDriveStrings(Len(Drives), Drives)
For rst = 1 To ret& Step 4
 rstStr = Mid(Drives, rst, 3)
 If GetDriveType(rstStr) = DRIVE_CDROM Then
  Count = Count + 1
  DriveLtrrs = DriveLtrrs & Left(rstStr, 1) & "  "
 End If
Next rst
If Count Then
 MsgBox "Number of CD-ROM(s): " & Count & vbNewLine & vbNewLine & "Drive Letters: " & UCase(DriveLtrrs), vbInformation + vbOKOnly, "Dedected CD-ROM(s)"
Unload Me
Else
 MsgBox "Can not dedect CD-ROM Drive", vbCritical + vbOKOnly
 Unload Me
End If
End Sub

Private Sub Timer2_Timer()
ProgressBar1.Value = 100
Dedect

End Sub
