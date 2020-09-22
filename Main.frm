VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jim's Mixer"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9585
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   3480
      Top             =   4800
   End
   Begin VB.Frame Frame4 
      Caption         =   "Other Options"
      Height          =   2535
      Left            =   4200
      TabIndex        =   35
      Top             =   2040
      Width           =   5295
      Begin VB.CommandButton Command1 
         Caption         =   "Quit"
         Height          =   495
         Left            =   1800
         TabIndex        =   44
         Top             =   1800
         Width           =   975
      End
      Begin VB.Frame Frame5 
         Caption         =   "BPM"
         Height          =   1335
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   2775
         Begin VB.CommandButton Command8 
            Caption         =   "&Reset"
            Height          =   375
            Left            =   1560
            TabIndex        =   41
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton Command7 
            Cancel          =   -1  'True
            Caption         =   "&BEAT"
            Default         =   -1  'True
            Height          =   615
            Left            =   1560
            MaskColor       =   &H8000000F&
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackColor       =   &H00101010&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   13.5
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "BPM counter"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   705
            Left            =   120
            TabIndex        =   42
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.CommandButton Show_File_Finder 
         Caption         =   "Show File Finder"
         Height          =   495
         Left            =   720
         TabIndex        =   37
         Top             =   1800
         Width           =   930
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   120
         Picture         =   "Main.frx":030A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   36
         Top             =   1800
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Mixer"
      Height          =   2535
      Left            =   0
      TabIndex        =   17
      Top             =   2040
      Width           =   4095
      Begin VB.Timer Timer3 
         Interval        =   1
         Left            =   1200
         Top             =   960
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Reset Fader"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         TabIndex        =   38
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   255
         Left            =   3720
         TabIndex        =   32
         Top             =   2880
         Width           =   255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   2880
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Automatic Mixing"
         Height          =   375
         Left            =   720
         TabIndex        =   30
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Timer Timer2 
         Interval        =   1
         Left            =   1560
         Top             =   2040
      End
      Begin ComctlLib.ProgressBar ProgressBar2 
         Height          =   255
         Left            =   2280
         TabIndex        =   27
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   327680
         Appearance      =   1
         MouseIcon       =   "Main.frx":0614
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   720
         TabIndex        =   26
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   327680
         Appearance      =   1
         MouseIcon       =   "Main.frx":0630
      End
      Begin VB.Timer timerfade2 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   2040
         Top             =   2040
      End
      Begin VB.Timer timerfade1 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   1080
         Top             =   2040
      End
      Begin VB.CommandButton Command3 
         Caption         =   "> X"
         Height          =   375
         Left            =   3360
         TabIndex        =   24
         Top             =   2040
         Width           =   495
      End
      Begin VB.HScrollBar Cross_Fader 
         Height          =   375
         Left            =   720
         Max             =   100
         Min             =   -100
         TabIndex        =   18
         Top             =   2040
         Value           =   100
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         Caption         =   "X <"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   2040
         Width           =   495
      End
      Begin VB.VScrollBar Deck1_Volume 
         Height          =   1350
         Left            =   120
         Max             =   -100
         TabIndex        =   20
         Top             =   480
         Width           =   240
      End
      Begin VB.VScrollBar Deck2_Volume 
         Height          =   1335
         Left            =   3720
         Max             =   -100
         TabIndex        =   19
         Top             =   480
         Width           =   240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         X1              =   2040
         X2              =   2040
         Y1              =   2460
         Y2              =   1920
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "0%       50%     100%"
         Height          =   255
         Left            =   2160
         TabIndex        =   29
         Top             =   600
         Width           =   1695
      End
      Begin VB.Line Line9 
         X1              =   2760
         X2              =   2760
         Y1              =   600
         Y2              =   480
      End
      Begin VB.Line Line8 
         X1              =   3360
         X2              =   3360
         Y1              =   600
         Y2              =   480
      End
      Begin VB.Line Line7 
         X1              =   2280
         X2              =   3360
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line6 
         X1              =   2280
         X2              =   2280
         Y1              =   480
         Y2              =   600
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "0%       50%     100%"
         Height          =   255
         Left            =   600
         TabIndex        =   28
         Top             =   600
         Width           =   1695
      End
      Begin VB.Line Line5 
         X1              =   1200
         X2              =   1200
         Y1              =   600
         Y2              =   480
      End
      Begin VB.Line Line4 
         X1              =   1800
         X2              =   1800
         Y1              =   600
         Y2              =   480
      End
      Begin VB.Line Line3 
         X1              =   720
         X2              =   1800
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line2 
         X1              =   720
         X2              =   720
         Y1              =   480
         Y2              =   600
      End
      Begin VB.Label Label4 
         Caption         =   "Deck2"
         Height          =   255
         Left            =   3480
         TabIndex        =   25
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Crossfader"
         Height          =   255
         Left            =   960
         TabIndex        =   22
         Top             =   1750
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Deck1"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Deck1"
      Height          =   1935
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4695
      Begin ComctlLib.ProgressBar ProgressBar3 
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   960
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   327680
         Appearance      =   1
         MouseIcon       =   "Main.frx":064C
      End
      Begin VB.CommandButton Deck1_Open 
         Caption         =   "Open"
         Height          =   495
         Left            =   1560
         TabIndex        =   8
         Top             =   1320
         Width           =   645
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Autoplay"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CheckBox Deck1_Mute 
         Caption         =   "Mute Deck 1"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   1410
      End
      Begin VB.Shape Shape1 
         Height          =   675
         Left            =   2400
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Deck1_Time 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   13.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2205
      End
      Begin VB.Label Deck1_Remain 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   13.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2220
      End
      Begin WMPLibCtl.WindowsMediaPlayer Deck1 
         Height          =   675
         Left            =   2640
         TabIndex        =   3
         Top             =   1080
         Width           =   1665
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   100
         mute            =   0   'False
         uiMode          =   "full"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   2937
         _cy             =   1191
      End
      Begin VB.Label Deck1_File 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<NO FILE>"
         ForeColor       =   &H00FFFFFF&
         Height          =   765
         Left            =   2400
         OLEDropMode     =   1  'Manual
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1320
      Top             =   5520
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   1080
      Top             =   4215
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Deck2"
      Height          =   1935
      Left            =   4800
      TabIndex        =   9
      Top             =   0
      Width           =   4695
      Begin ComctlLib.ProgressBar ProgressBar4 
         Height          =   255
         Left            =   2400
         TabIndex        =   34
         Top             =   960
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   327680
         Appearance      =   1
         MouseIcon       =   "Main.frx":0668
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Autoplay"
         Height          =   255
         Left            =   3240
         TabIndex        =   16
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CheckBox Deck2_Mute 
         Caption         =   "Mute Deck 2"
         Height          =   255
         Left            =   3240
         TabIndex        =   15
         Top             =   1320
         Width           =   1410
      End
      Begin VB.CommandButton Deck2_Open 
         Caption         =   "Open"
         Height          =   495
         Left            =   2520
         TabIndex        =   14
         Top             =   1320
         Width           =   645
      End
      Begin VB.Shape Shape2 
         Height          =   675
         Left            =   120
         Top             =   1080
         Width           =   2175
      End
      Begin WMPLibCtl.WindowsMediaPlayer Deck2 
         Height          =   675
         Left            =   360
         TabIndex        =   13
         Top             =   1080
         Width           =   1665
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   100
         mute            =   0   'False
         uiMode          =   "full"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   2937
         _cy             =   1191
      End
      Begin VB.Label Deck2_Remain 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   13.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2400
         TabIndex        =   12
         Top             =   600
         Width           =   2205
      End
      Begin VB.Label Deck2_Time 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   13.5
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2400
         TabIndex        =   11
         Top             =   240
         Width           =   2205
      End
      Begin VB.Label Deck2_File 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<NO FILE>"
         ForeColor       =   &H00FFFFFF&
         Height          =   765
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   10
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   5760
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private BPMArray(14) As Single
Private LastBPM
Private MaxBPMs
Private BPM As Integer
Private OldBPM As Integer

Private OutBuffer As String
Private Sub Check1_Click()
Deck1.settings.autoStart = Not Deck1.settings.autoStart
End Sub

Private Sub Check2_Click()
Deck2.settings.autoStart = Not Deck2.settings.autoStart
End Sub

Private Sub Check3_Click()
Command6.Enabled = False
Timer3.Enabled = Not Timer3.Enabled
Command3.Enabled = Not Command3.Enabled
Command2.Enabled = Not Command2.Enabled
End Sub

'-------------------------
'|     Jim's Mixer       |
'-------------------------
'This version has the following additions:
'
' - BPM Counter
' - Drag and drop ability (i.e. Windows Explorer)
' - File find form to help changing tracks
' - Time elapsed and time remaining display
'
'Anyway, hope you like it, and don't forget to drop me a line!


Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
On Error Resume Next
If timerfade.Value = 1 Then
timerfade2.Enabled = False
timerfade1.Enabled = True
Else
Cross_Fader.Value = Cross_Fader.Value - 5
timerfade2.Enabled = False
timerfade1.Enabled = False
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
If timerfade.Value = 1 Then
timerfade2.Enabled = True
timerfade1.Enabled = False
Else
Cross_Fader.Value = Cross_Fader.Value + 5
timerfade2.Enabled = False
timerfade1.Enabled = False
End If
End Sub

Private Sub Command4_Click()
On Error Resume Next
If timerfade.Value = 1 Then
timerfade2.Enabled = False
timerfade1.Enabled = True
Else
Cross_Fader.Value = Cross_Fader.Value - 5
timerfade2.Enabled = False
timerfade1.Enabled = False
End If
End Sub

Private Sub Command5_Click()
On Error Resume Next
If timerfade.Value = 1 Then
timerfade2.Enabled = True
timerfade1.Enabled = False
Else
Cross_Fader.Value = Cross_Fader.Value + 5
timerfade2.Enabled = False
timerfade1.Enabled = False
End If
End Sub

Private Sub Command6_Click()
timerfade1.Enabled = False
timerfade2.Enabled = False
Cross_Fader.Value = 0
Command6.Enabled = False
End Sub

Private Sub Command8_Click()
  For I = 0 To 14
        BPMArray(I) = 0
    Next
    MaxBPMs = 0
Label8.Caption = Format(OldBPM / 100, "##0.00")
Command7.SetFocus
End Sub

Private Sub Cross_Fader_Change()
' Right, this is where the crossfading is done, 2 lines of code! Simple!
If Cross_Fader.Value > 0 Then Deck1_Volume.Value = (100 - Cross_Fader.Value) - 100
If Cross_Fader.Value < 0 Then Deck2_Volume.Value = Cross_Fader.Value
End Sub

Private Sub Cross_Fader_Scroll()
Cross_Fader_Change
timerfade1.Enabled = False
timerfade2.Enabled = False
End Sub



Private Sub Deck1_File_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim File
On Error GoTo Error
For Each File In Data.Files
Deck1.URL = File
Deck1_File.Caption = Mid(File, InStrRevVB5(File, "\") + 1, Len(File))
Next File
Exit Sub

Error:
MsgBox "Not a valid file!", vbCritical, "Error"
End Sub

Private Sub Deck1_Mute_Click()
If Deck1_Mute.Value = 1 Then Deck1.settings.mute = True
If Deck1_Mute.Value = 0 Then Deck1.settings.mute = False
End Sub



Private Sub Deck1_Open_Click()
On Error GoTo Error
Dialog.CancelError = True 'This is to stop the track resetting when playing
                          'if cancel is pressed
Dialog.Filter = "All supported files |*.wav;*.wma;*.mp3;*.mid|MP3 Files *.mp3|*.mp3|Wave Files *.wav|*.wav|Midi Files *.mid|*.mid"
Dialog.ShowOpen

Deck1.URL = Dialog.filename
'Visual basic 6 users may want to get rid of the module...since it is a feature
'that is already on VB6 (InStrRev)
Deck1_File.Caption = Mid(Dialog.filename, InStrRevVB5(Dialog.filename, "\") + 1, Len(Dialog.filename))
If Dialog.filename = "" Then Deck1_File.Caption = "<NO FILE>"
Exit Sub

Error:
If Err.Number <> 32755 Then
MsgBox "Error loading file - " & Err.Number & " : " & Err.Description
Else
End If
End Sub

Private Sub Deck1_Volume_Change()
Deck1.settings.volume = Deck1_Volume.Value + 100
End Sub

Private Sub Deck1_Volume_Scroll()
Deck1_Volume_Change
End Sub

Private Sub Deck2_File_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim File
On Error GoTo Error
For Each File In Data.Files
Deck2.URL = File
Deck2_File.Caption = Mid(File, InStrRevVB5(File, "\") + 1, Len(File))
Next File
Exit Sub

Error:
MsgBox "Not a valid file!", vbCritical, "Error"
End Sub

Private Sub Deck2_Mute_Click()
If Deck2_Mute.Value = 1 Then Deck2.settings.mute = True
If Deck2_Mute.Value = 0 Then Deck2.settings.mute = False
End Sub

Private Sub Deck2_Open_Click()
On Error GoTo Error
Dialog.CancelError = True
Dialog.Filter = "All supported files |*.wav;*.wma;*.mp3;*.mid|MP3 Files *.mp3|*.mp3|Wave Files *.wav|*.wav|Midi Files *.mid|*.mid"
Dialog.ShowOpen

Deck2.URL = Dialog.filename
'Visual basic 6 users may want to get rid of the module...since it is a feature
'that is already on VB6 (InStrRev)
Deck2_File.Caption = Mid(Dialog.filename, InStrRevVB5(Dialog.filename, "\") + 1, Len(Dialog.filename))
If Dialog.filename = "" Then Deck2_File.Caption = "<NO FILE>"
Exit Sub

Error:
If Err.Number <> 32755 Then ' Cancel was pressed?
MsgBox "Error loading file - " & Err.Number & " : " & Err.Description
Else
End If
End Sub

Private Sub Deck2_Volume_Change()
Deck2.settings.volume = Deck2_Volume.Value + 100
Label2.Caption = Deck2_Volume.Value + 100
End Sub

Private Sub Deck2_Volume_Scroll()
Deck2_Volume_Change
End Sub

Private Sub Form_Load()
Cross_Fader.Value = 0
Me.Caption = "Jim's Mixer v" & App.Major & "." & App.Minor & App.Revision & " Remixed By DJ Stuntmaster"
Deck1.settings.autoStart = False
Deck2.settings.autoStart = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
'unload any other form
   
For Each Form In Forms
      
DoEvents

If Form.Name <> "Form1" Then Unload Form
   Next
   
DoEvents
Set Form = Nothing
End Sub

Private Sub Show_File_Finder_Click()
On Error Resume Next
Form2.Show
Form2.Top = Form1.Top + Form1.Height
Form2.Left = Form1.Left
End Sub



Private Sub Timer1_Timer()
' Show time

' > DECK 1
On Error Resume Next
If Deck1.Controls.currentPosition > 0 Then
Deck1_Time.Caption = TimeSerial(0, 0, Int(Deck1.Controls.currentPosition))
End If
'Remaining time
Deck1_Remain.Caption = "" & TimeSerial(0, 0, Int(Deck1.currentMedia.duration) - Int(Deck1.Controls.currentPosition)) & ""

' > DECK 2
On Error Resume Next
If Deck2.Controls.currentPosition > 0 Then
Deck2_Time.Caption = TimeSerial(0, 0, Int(Deck2.Controls.currentPosition))
End If
'Remaining time
Deck2_Remain.Caption = "" & TimeSerial(0, 0, Int(Deck2.currentMedia.duration) - Int(Deck2.Controls.currentPosition)) & ""
' Turn mp3 name to red if 20 seconds or less left in track

'DECK 1

If Deck1.Controls.currentPosition >= (Deck1.currentMedia.duration - 20) Then
Deck1_File.ForeColor = vbRed
Else
Deck1_File.ForeColor = vbWhite
End If

'DECK 2

If Deck2.Controls.currentPosition >= (Deck2.currentMedia.duration - 20) Then
Deck2_File.ForeColor = vbRed
Else
Deck2_File.ForeColor = vbWhite
End If


ProgressBar3.Max = Deck1.currentMedia.duration
ProgressBar3.Value = Deck1.Controls.currentPosition
ProgressBar4.Max = Deck2.currentMedia.duration
ProgressBar4.Value = Deck2.Controls.currentPosition
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
ProgressBar1.Value = Deck1_Volume.Value + 100
ProgressBar2.Value = Deck2_Volume.Value + 100




If Check3.Value = 1 And Deck1_Remain.Caption = "00:00:10" Then
timerfade2.Enabled = True
timerfade1.Enabled = False
Deck2.Controls.Play
Else
End If

If Check3.Value = 1 And Deck2_Remain.Caption = "00:00:10" Then
timerfade2.Enabled = False
timerfade1.Enabled = True
Deck1.Controls.Play
Else
End If
End Sub

Private Sub Timer3_Timer()
If Cross_Fader.Value <> 0 Then
Command6.Enabled = True
Else
Command6.Enabled = False
End If
End Sub

Private Sub Timer4_Timer()
Show_File_Finder_Click
Timer4.Enabled = False
End Sub

Private Sub timerfade1_Timer()
On Error Resume Next
Cross_Fader.Value = Cross_Fader.Value - 5
End Sub

Private Sub timerfade2_Timer()
On Error Resume Next
Cross_Fader.Value = Cross_Fader.Value + 5
End Sub
Sub ProcessStream(Stream As String)
OldBPM = Asc(Left$(Stream, 1)) + Asc(Right$(Stream, 1)) * 256
End Sub
Function GetLatest() As String
GetLatest = Chr$(Val(Label2.Caption)) + OutBuffer + Chr$(0)
OutBuffer = ""
End Function
Function InitializeDevice(id As Byte) As String
Dim TempID As Long
TempID = Device_Channel + Channel_ChannelID + Channel_Commands + Channel_BPM
'Debug.Print TempID
InitializeDevice = Chr$(TempID And &HFF&) + Chr$((TempID And &HFF00&) / &H100&) + Chr$((TempID And &HFF0000) / &H10000) + Chr$(0)

End Function
Private Sub Command7_Click()
On Error Resume Next
Static LastClick
If LastClick <> 0 And LastClick < Timer Then
    LastBPM = (LastBPM + 1) Mod 15
    If MaxBPMs < 15 Then MaxBPMs = MaxBPMs + 1
    BPMArray(LastBPM) = 60 / (Timer - LastClick)
    For I = 0 To 14
        cBPM = cBPM + BPMArray(I)
    Next
    Label8.Caption = Format(cBPM / MaxBPMs, "##0.00")
    BPM = (cBPM / MaxBPMs) * 100
End If

LastClick = Timer
End Sub

Private Sub Command7_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = o Then Command7.BackColor = &HFF
End Sub

Private Sub Command7_KeyUp(KeyCode As Integer, Shift As Integer)
    Command7.BackColor = &H8000000F
    If Shift = 0 And KeyCode <> 32 Then Command7_Click
End Sub

Private Sub Command7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command7.BackColor = &HFF
End Sub

Private Sub Command7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command7.BackColor = &H8000000F
End Sub
