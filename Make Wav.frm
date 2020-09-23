VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Make Wave"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Sample Rate"
      Height          =   1335
      Left            =   1320
      TabIndex        =   1
      Top             =   165
      Width           =   1695
      Begin VB.OptionButton Option4 
         Caption         =   "22050 Hz"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "5000 Hz"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open File"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make wave!"
      Height          =   855
      Left            =   1320
      TabIndex        =   0
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "For full compability, rename your file to *.wav"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Caption         =   "Caution! Your file will be edited and it cannot be undone!"
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   1590
      Width           =   4005
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Dim SampleRate As Integer
    Dim Sample As Integer
    Dim Time As Integer

Private Sub Command1_Click()

    Open CommonDialog1.FileName For Binary Access Write As #1

        Time = (LOF(1) - 44) / SampleRate
        frmMain.Caption = LOF(1) & " bytes " & Time & " and seconds long"

        'Riff-chunk (+ "fmt ")
        Put #1, 1, "RIFF"       '"RIFF"
        Put #1, 5, LOF(1) - 8   'Length of package to follow (LOF - "RIFF" & theese 4 bytes)
        Put #1, 9, "WAVEfmt "   '"WAVEfmt "

        'Format-chunk (- "fmt ")
        Put #1, 17, 16          'length of format-chunk
        Put #1, 18, 0
        Put #1, 19, 0
        Put #1, 20, 0
        Put #1, 21, 1           'allways 01 00
        Put #1, 22, 0
        Put #1, 23, 1           '01 00 = mono, 02 00 = stereo
        Put #1, 24, 0
        Put #1, 25, SampleRate  'Samplerate e.g. 22050
        Put #1, 27, 0
        Put #1, 28, 0
        Put #1, 29, SampleRate  'Bytes per second, same as samplerate when 8-bit mono
        Put #1, 31, 0
        Put #1, 32, 0
        Put #1, 33, 1           'Bytes per sample, 01 00 = 8-bit mono, 02 00 = 8-bit stereo or 16-bit mono, 04 00 = 16-bit stereo
        Put #1, 34, 0
        Put #1, 35, 8           'Bits per sample, 08 00 or 10 00
        Put #1, 36, 0           'a byte for above

        'Data-chunk
        Put #1, 37, "data"
        Put #1, 41, LOF(1) - 44

    Close #1

End Sub

Private Sub Form_Load()

    SampleRate = 5000

    On Error GoTo ErrHandler

    CommonDialog1.ShowOpen

ErrHandler:
End Sub


Private Sub Option3_Click()

    SampleRate = 5000

End Sub

Private Sub Option4_Click()
    
    SampleRate = 22050

End Sub
