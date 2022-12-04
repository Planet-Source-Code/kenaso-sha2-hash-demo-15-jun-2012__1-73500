VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5505
   ClientLeft      =   1860
   ClientTop       =   2400
   ClientWidth     =   5130
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3799.649
   ScaleMode       =   0  'User
   ScaleWidth      =   4817.334
   Begin VB.Frame fraThanks 
      Height          =   3735
      Left            =   45
      TabIndex        =   5
      Top             =   855
      Width           =   5025
      Begin VB.PictureBox picURL 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   525
         Index           =   0
         Left            =   105
         Picture         =   "frmAbout.frx":12FA
         ScaleHeight     =   525
         ScaleWidth      =   4845
         TabIndex        =   12
         Top             =   495
         Width           =   4845
         Begin VB.Label lblURL 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "VBNet API code snippets for VB6"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   0
            Left            =   1695
            TabIndex        =   13
            Top             =   180
            Width           =   2685
         End
      End
      Begin VB.PictureBox picURL 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Index           =   1
         Left            =   105
         Picture         =   "frmAbout.frx":2804
         ScaleHeight     =   465
         ScaleWidth      =   4845
         TabIndex        =   10
         Top             =   1165
         Width           =   4845
         Begin VB.Label lblURL 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Planet Source Code for Visual Basic"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   1
            Left            =   1695
            TabIndex        =   11
            Top             =   105
            Width           =   2970
         End
      End
      Begin VB.PictureBox picURL 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   645
         Index           =   2
         Left            =   105
         Picture         =   "frmAbout.frx":3FC2
         ScaleHeight     =   645
         ScaleWidth      =   4845
         TabIndex        =   8
         Top             =   1775
         Width           =   4845
         Begin VB.Label lblURL 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Classic VB code snippets"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   2
            Left            =   1695
            TabIndex        =   9
            Top             =   225
            Width           =   2115
         End
      End
      Begin VB.PictureBox picURL 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Index           =   3
         Left            =   105
         Picture         =   "frmAbout.frx":6C8C
         ScaleHeight     =   555
         ScaleWidth      =   4845
         TabIndex        =   6
         Top             =   2565
         Width           =   4845
         Begin VB.Label lblURL 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "All API Network - VB6 reference "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   3
            Left            =   1695
            TabIndex        =   7
            Top             =   105
            Width           =   2580
         End
      End
      Begin VB.Label lblThankYou 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Acknowledgements"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1425
         TabIndex        =   15
         Top             =   150
         Width           =   2295
      End
      Begin VB.Label lblOperSysInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1800
         TabIndex        =   14
         Top             =   3240
         Width           =   3090
      End
   End
   Begin VB.CommandButton cmdChoice 
      BackColor       =   &H00E0E0E0&
      Height          =   690
      Index           =   1
      Left            =   4290
      Picture         =   "frmAbout.frx":89CE
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Return to main screen"
      Top             =   4695
      Width           =   690
   End
   Begin VB.CommandButton cmdChoice 
      BackColor       =   &H00E0E0E0&
      Height          =   690
      Index           =   0
      Left            =   3540
      Picture         =   "frmAbout.frx":8CD8
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "System Information"
      Top             =   4695
      Width           =   690
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   1455
      Picture         =   "frmAbout.frx":911A
      ScaleHeight     =   435
      ScaleWidth      =   2220
      TabIndex        =   0
      Top             =   120
      Width           =   2220
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   165
      TabIndex        =   4
      Top             =   4710
      Width           =   2550
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Kenneth Ives"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1695
      TabIndex        =   1
      Top             =   585
      Width           =   1680
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ***************************************************************************
' Module:        frmAbout
'
' Description:   This form displays the sites I would like to give thanks.
'                For their code or other information on how to accomplish
'                something.
'
' ===========================================================================
'    DATE      NAME / DESCRIPTION
' -----------  --------------------------------------------------------------
' 19-FEB-2004  Kenneth Ives  kenaso@tx.rr.com
'              Original
' 01-Nov-2008  Kenneth Ives  kenaso@tx.rr.com
'              Verified screen would not be displayed during initial load
' 05-Jul-2011  Kenneth Ives  kenaso@tx.rr.com
'              Updated hyperlinks and user interface links
' ***************************************************************************
Option Explicit

' ***************************************************************************
' Constants
' ***************************************************************************
  Private Const SW_SHOWNORMAL    As Long = 1

' ***************************************************************************
' API Declares
' ***************************************************************************
  ' The ShellExecute function opens or prints a specified file.  The file
  ' can be an executable file or a document file.
  Private Declare Function ShellExecute Lib "shell32.dll" _
          Alias "ShellExecuteA" (ByVal hwnd As Long, _
          ByVal lpOperation As String, ByVal lpFile As String, _
          ByVal lpParameters As String, ByVal lpDirectory As String, _
          ByVal nShowCmd As Long) As Long

  ' The GetDesktopWindow function returns a handle to the desktop window.
  ' The desktop window covers the entire screen. The desktop window is
  ' the area on top of which other windows are painted.
  Private Declare Function GetDesktopWindow Lib "user32" () As Long


Private Sub Form_Initialize()

    ' Make sure this form is hidden during initial load
    frmAbout.Hide
    DoEvents
    
End Sub

Private Sub Form_Load()
          
    Dim strVersion As String
    Dim objOperSys As New cOperSystem  ' Define and instantiate classes
    Dim objKeyEdit As New cKeyEdit
    
    DisableX frmAbout   ' Disable "X" in upper right corner of form
    
    ' Capture information about
    ' this operating system
    With objOperSys
        strVersion = .VersionName & vbNewLine & _
                     "Ver " & .VersionNumber & _
                     "." & .BuildNumber & _
                     "  " & .ServicePack
    End With
    
    ' Hide this form
    With frmAbout
        .Caption = "About - " & PGM_NAME
        .lblOperSysInfo.Caption = strVersion
        .lblDisclaimer.Caption = "This is a freeware product." & vbNewLine & _
                                 "No warranties or guarantees implied or intended."
    
        ' center form on screen
        .Move (Screen.Width - frmAbout.Width) \ 2, (Screen.Height - frmAbout.Height) \ 2
        .Hide
    End With
    
    objKeyEdit.CenterCaption frmAbout   ' Center form window caption
    
    Set objOperSys = Nothing  ' Free class objects form memory
    Set objKeyEdit = Nothing

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    ' Based on the unload code the system passes, we determine what to do.
    '
    ' Unloadmode codes
    '     0 - Close from the control-menu box or Upper right "X"
    '     1 - Unload method from code elsewhere in the application
    '     2 - Windows Session is ending
    '     3 - Task Manager is closing the application
    '     4 - MDI Parent is closing
    Select Case UnloadMode
           Case 0    ' return to main form
                frmMain.Show
                frmAbout.Hide
    
           Case Else
                ' Fall thru. Something else is shutting us down.
    End Select

End Sub

Private Sub lblAuthor_Click()
    SendEmail
End Sub

Private Sub cmdChoice_Click(Index As Integer)

    Dim lngHwnd As Long

    Select Case Index
             
           Case 0  ' Search for System information window
                lngHwnd = FindProcessByCaption("System Info")
                
                ' MSInfo32.exe already active
                If lngHwnd <> 0 Then
                    ' MSInfo32 window may be hidden.
                    ' This is why it will be stopped
                    ' and then restarted.
                    StopProcessByHandle lngHwnd  ' Shutdown MSInfo32 application
                    DoEvents                     ' Allow for system update
                End If
                
                DisplaySysInfo   ' Display System Information (MSInfo32.exe)
    
           Case 1  ' Return to main form
                frmMain.Show
                frmAbout.Hide
    End Select
  
End Sub

Private Sub DisplaySysInfo()

    Dim strPath As String
    
    strPath = Environ$("CommonProgramFiles")                    ' Capture DOS environment variable
    strPath = QualifyPath(TrimStr(strPath))                     ' Prepare base path
    strPath = strPath & "Microsoft Shared\MSInfo\msinfo32.exe"  ' Prepare complete path and filename
    
    ' Verify file does exist
    If IsPathValid(strPath) Then
        Shell strPath, vbNormalFocus   ' Display system information
    End If
    
End Sub

Public Sub DisplayAbout()

    ' Called from frmMain to display this form
    With frmAbout
         .Move (Screen.Width - frmAbout.Width) \ 2, (Screen.Height - frmAbout.Height) \ 2
         .Show vbModeless
         .Refresh
    End With
    
End Sub

Private Sub lblURL_Click(Index As Integer)

    Dim strURL As String
    
    ' Identify URL to be executed
    Select Case Index
           Case 0: strURL = "http://vbnet.mvps.org/index.html"
           Case 1: strURL = "http://www.Planet-Source-Code.com/vb/"
           Case 2: strURL = "http://vb.mvps.org/"
           Case 3: strURL = "http://allapi.mentalis.org/apilist/apilist.php"
           Case Else: Exit Sub
    End Select
           
    RunShellExecute strURL   ' Make hyperlink call
   
End Sub

Private Sub picURL_Click(Index As Integer)

    Dim strURL As String
    
    ' Identify URL to be executed
    Select Case Index
           Case 0: strURL = "http://vbnet.mvps.org/index.html"
           Case 1: strURL = "http://www.Planet-Source-Code.com/vb/"
           Case 2: strURL = "http://vb.mvps.org/"
           Case 3: strURL = "http://allapi.mentalis.org/apilist/apilist.php"
           Case Else: Exit Sub
    End Select
           
    RunShellExecute strURL   ' Make hyperlink call
   
End Sub

Private Sub RunShellExecute(ByVal strURL As String)

    ' Called by lblURL_Click()
    '           picURL_Click()
        
    ShellExecute GetDesktopWindow(), _
                 "open", strURL, 0&, 0&, SW_SHOWNORMAL
End Sub


