VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remove Duplicated Chars"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pbrFilter 
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.TextBox txtAfter 
      Height          =   1845
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3720
      Width           =   4455
   End
   Begin VB.CheckBox chkSensitive 
      Caption         =   "Case Sensitive"
      Height          =   195
      Left            =   3120
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdFilter 
      Caption         =   "Filter"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   4455
   End
   Begin VB.TextBox txtBefore 
      Height          =   1845
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmMain.frx":000C
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Before :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "After :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   4455
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4560
      Y1              =   3240
      Y2              =   3240
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function RemoveRepeatedChar(sString As String, bCaseSensitive As Boolean) As String
Dim lString As Long
Dim sFilter As String

If Len(sString) > 0 Then
    Do
        lString = lString + 1
        If bCaseSensitive = False Then
            'Search the string for repeated characters
            If InStr(1, sFilter, Mid(sString, lString, 1)) = 0 Then
                'If there is a new one...
                sFilter = sFilter & Mid(sString, lString, 1)
            End If
        Else
            'You can also use LCase instead of UCase
            If InStr(1, sFilter, UCase(Mid(sString, lString, 1))) = 0 Then
                sFilter = sFilter & UCase(Mid(sString, lString, 1))
            End If
        End If
        'I added this to calculate the progress, you can remove it if you think it's unneccesary
        pbrFilter.Value = lString / Len(sString) * 100
        'Keep everything running smooth
        DoEvents
    Loop Until lString >= Len(sString)
End If
'Success!
RemoveRepeatedChar = sFilter

End Function

Private Sub cmdFilter_Click()
pbrFilter.Value = 0
txtAfter = RemoveRepeatedChar(txtBefore, CBool(chkSensitive.Value))

End Sub

