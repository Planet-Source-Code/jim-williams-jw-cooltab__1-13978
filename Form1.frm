VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "*\AJW_CoolTabs.vbp"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Demo of JW-Cooltab - ©ADMAX 2000"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   3915
   StartUpPosition =   2  'CenterScreen
   Begin Project1.JW_CoolTabs JW_CoolTabs1 
      Height          =   2655
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   4683
      CaptionTab1     =   "Intro"
      CaptionTab2     =   "Files"
      CaptionTab3     =   "Image"
      CaptionTab4     =   "Date"
   End
   Begin VB.Frame tabFrame 
      BackColor       =   &H00E0E0E0&
      ClipControls    =   0   'False
      Height          =   2235
      Index           =   7
      Left            =   2040
      TabIndex        =   18
      Top             =   2940
      Width           =   3795
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cooltab 8 - for use..."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   3555
      End
   End
   Begin VB.Frame tabFrame 
      BackColor       =   &H00E0E0E0&
      ClipControls    =   0   'False
      Height          =   2235
      Index           =   6
      Left            =   1860
      TabIndex        =   16
      Top             =   2640
      Width           =   3795
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cooltab 7 - for use..."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   3555
      End
   End
   Begin VB.Frame tabFrame 
      BackColor       =   &H00E0E0E0&
      ClipControls    =   0   'False
      Height          =   2235
      Index           =   5
      Left            =   1680
      TabIndex        =   14
      Top             =   2400
      Width           =   3795
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cooltab 6 - for use..."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   3555
      End
   End
   Begin VB.Frame tabFrame 
      BackColor       =   &H00E0E0E0&
      ClipControls    =   0   'False
      Height          =   2235
      Index           =   4
      Left            =   1560
      TabIndex        =   12
      Top             =   2160
      Width           =   3795
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cooltab 5 - for use..."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   3555
      End
   End
   Begin VB.Frame tabFrame 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   2235
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   3795
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "I know I have to do some more job on this control, so that tab´s can be added and so on, but this will be done another day..."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   540
         TabIndex        =   11
         Top             =   1260
         Width           =   3015
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Email: jim@admax.se"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   540
         TabIndex        =   10
         Top             =   1860
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form1.frx":0000
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   540
         TabIndex        =   9
         Top             =   420
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "The Cooltab from ADMAX"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   540
         TabIndex        =   8
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame tabFrame 
      BackColor       =   &H00E0E0E0&
      ClipControls    =   0   'False
      Height          =   2235
      Index           =   3
      Left            =   1200
      TabIndex        =   3
      Top             =   1320
      Width           =   3795
      Visible         =   0   'False
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Weeks"
         Height          =   555
         Left            =   2700
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1560
         Width           =   975
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2010
         Left            =   60
         TabIndex        =   4
         Top             =   120
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   3545
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   14737632
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         StartOfWeek     =   24510466
         CurrentDate     =   36887
      End
   End
   Begin VB.Frame tabFrame 
      BackColor       =   &H00E0E0E0&
      ClipControls    =   0   'False
      Height          =   2235
      Index           =   2
      Left            =   900
      TabIndex        =   2
      Top             =   1080
      Width           =   3795
      Visible         =   0   'False
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "You know the toaster - right..."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   600
         Width           =   2355
      End
      Begin VB.Image Image1 
         Height          =   1110
         Left            =   120
         Picture         =   "Form1.frx":00AD
         Top             =   180
         Width           =   990
      End
   End
   Begin VB.Frame tabFrame 
      BackColor       =   &H00E0E0E0&
      ClipControls    =   0   'False
      Height          =   2235
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   3795
      Visible         =   0   'False
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   60
         TabIndex        =   7
         Top             =   120
         Width           =   3675
      End
   End
   Begin Project1.JW_CoolTabs JW_CoolTabs2 
      Height          =   2655
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   4683
      CaptionTab1     =   "Tab 5"
      CaptionTab2     =   "Tab 6"
      CaptionTab3     =   "Tab 7"
      CaptionTab4     =   "Tab8"
   End
   Begin VB.Menu myPopmenu 
      Caption         =   "myPopmenu"
      Visible         =   0   'False
      Begin VB.Menu Demo 
         Caption         =   "About..."
      End
      Begin VB.Menu L1 
         Caption         =   "-"
      End
      Begin VB.Menu Tab2 
         Caption         =   "Second tabset..."
      End
   End
   Begin VB.Menu myPopmenu2 
      Caption         =   "myPopmenu2"
      Visible         =   0   'False
      Begin VB.Menu Demo2 
         Caption         =   "About..."
      End
      Begin VB.Menu L2 
         Caption         =   "-"
      End
      Begin VB.Menu Tab1 
         Caption         =   "First tabset..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' This code was written by Jim Williams (jim@admax.se)
' in order to try to get as close as possible to
' the simular control that exists in Photoshop.
'
' I think it turned out pretty close :-)
'
' Use it if you like it, and if you do - all I ask from you is
' that you send me an email (jim@admax.se) and tell me that
' you will use my code.
'
' (I have wrote other cool controls also, but not released them yet)
'
' Enjoy this free code...

Dim lastT1 As Integer
Dim lastT2 As Integer




Private Sub Check1_Click()

If Check1.Value = 1 Then
    MonthView1.ShowWeekNumbers = True
Else
    MonthView1.ShowWeekNumbers = False
End If

End Sub

Private Sub Close_Click()
Unload Me
End Sub

Private Sub Demo_Click()
MsgBox "If you are tired of Microsofts old and ugly Tab-control, then this Cooltab-control is the stuff for you. It mimmicks the 'look & feel' from the Tab´s from Photoshop." & vbCrLf & vbCrLf & "This control was borned because I was bored one day during christmas, so - thank´s Santa :-)" & vbCrLf & vbCrLf & "Jim Williams (jim@admax.se)", 64, "About JW_Cooltab"
End Sub

Private Sub Demo2_Click()
Demo_Click
End Sub

Private Sub JW_CoolTabs1_RoundBtnClick()
PopupMenu myPopmenu
End Sub





Private Sub JW_CoolTabs1_TabClick(Index As Integer)
lastT1 = Index

tabFrame(Index).Move 60, 360
With tabFrame(Index)
    .BorderStyle = 0
    .ZOrder
    .Visible = True
End With
End Sub


Private Sub JW_CoolTabs2_RoundBtnClick()
PopupMenu myPopmenu2
End Sub

Private Sub JW_CoolTabs2_TabClick(Index As Integer)
lastT2 = Index

tabFrame(Index + 4).Move 60, 360
With tabFrame(Index + 4)
    .BorderStyle = 0
    .ZOrder
    .Visible = True
End With
End Sub


Private Sub Tab1_Click()
JW_CoolTabs1.ZOrder
DoEvents
JW_CoolTabs1_TabClick lastT1

End Sub

Private Sub Tab2_Click()
JW_CoolTabs2.ZOrder
DoEvents
JW_CoolTabs2_TabClick lastT2

End Sub


