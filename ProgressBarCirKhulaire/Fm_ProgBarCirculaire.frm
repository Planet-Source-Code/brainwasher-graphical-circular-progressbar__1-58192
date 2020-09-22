VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Fm_CircularProgressBar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Graphical circular progress bar ..."
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   335
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   549
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cmd_BothBars 
      Caption         =   "Test both progressbars"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   5175
   End
   Begin VB.Frame Frm_ProgressionType 
      Caption         =   "Progression style ..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4560
      TabIndex        =   14
      Top             =   2400
      Width           =   3495
      Begin VB.OptionButton Opt_Common 
         Caption         =   "Common"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1920
         TabIndex        =   1
         Top             =   480
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Opt_Separated 
         Caption         =   "Separated"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Timer Timer2 
      Left            =   1800
      Top             =   3120
   End
   Begin VB.CommandButton Cmd_ClassicalPB 
      Caption         =   "Test classical progressbar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Frame Frm_ProgressBar 
      Caption         =   "Classical progress bar..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   7935
      Begin MSComctlLib.ProgressBar Int_ClassicalProgressbar 
         Height          =   255
         Left            =   3240
         TabIndex        =   10
         Top             =   360
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar Ext_ClassicalProgressbar 
         Height          =   255
         Left            =   3240
         TabIndex        =   11
         Top             =   720
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label ILbl_ExtProgressBar 
         AutoSize        =   -1  'True
         Caption         =   "External progress bar simulation :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   2925
      End
      Begin VB.Label ILbl_IntProgressBar 
         AutoSize        =   -1  'True
         Caption         =   "Internal progress bar simulation :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Timer Timer1 
      Left            =   1320
      Top             =   3120
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   3600
      Picture         =   "Fm_ProgBarCirculaire.frx":0000
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   8
      Top             =   2640
      Width           =   540
   End
   Begin VB.CommandButton Cmd_End 
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5520
      TabIndex        =   5
      Top             =   3720
      Width           =   2535
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   720
      Picture         =   "Fm_ProgBarCirculaire.frx":0E7A
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   120
      Picture         =   "Fm_ProgBarCirculaire.frx":1CF4
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   34
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.CommandButton Cmd_CircularPB 
      Caption         =   "Test circular progressbar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Lbl_Info 
      BackStyle       =   0  'Transparent
      Caption         =   "Here you can see the result of the circular progress bar                 -->"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   120
      Top             =   2520
      Width           =   4335
   End
   Begin VB.Label Lbl_Title 
      BackStyle       =   0  'Transparent
      Caption         =   "This tutorial will show you how to create such a circular progress bar. The image can be changed."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   960
      TabIndex        =   16
      Top             =   480
      Width           =   7215
   End
   Begin VB.Label Lbl_Title 
      BackStyle       =   0  'Transparent
      Caption         =   "Create a graphical circular progress bar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   15
      Top             =   120
      Width           =   7215
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   750
      Left            =   120
      Picture         =   "Fm_ProgBarCirculaire.frx":2B6E
      Top             =   120
      Width           =   750
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   -120
      Top             =   0
      Width           =   8415
   End
   Begin VB.Line Line1 
      X1              =   8
      X2              =   544
      Y1              =   232
      Y2              =   232
   End
End
Attribute VB_Name = "Fm_CircularProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'___________________________________________________________________________
' Program name      : CircPbar.
' Description       : A simple graphical circular progress bar.
' Company           : MELANTECH
' Authors           : Weitten Pascal
'___________________________________________________________________________
'
' Date              : (c) 2000.
' Version N°        : V0.01
' Customer          : Internal stuff.
'
' Last Modification : 2005.01.06
'___________________________________________________________________________
' TODO :
'       -
'       -
'___________________________________________________________________________
'

'Variables definition.
Dim LastInternalAngle As Double, LastExternalAngle As Double
Dim InternalAngle As Double, ExternalAngle As Double
Dim SeparatedProgression As Boolean

'Constants
Const Pi = 3.141592654
Const MaxAngle_Degrees = 360
Const DrawLineBeforeAngle = True    'True=Draw a little line before the angle
                                    'data. Just here for fun.
                                    'False=no line :-)
Const LineAngle = 4
Const AngleStep = 45                'Circular PB test angle. This would be
                                    'the percentage value in a real case.
                                    'Change this value to test it.

Const MaxValueNormalPB = 360        'Just here to synchro test the classical PB.
Const NormalPB_Step = 45

Private Sub Cmd_End_Click()
    End
End Sub

Private Sub Cmd_CircularPB_Click()
    Call EnableButtons(False)
    Call Init_CircularBar
End Sub

Private Sub Cmd_BothBars_Click()
    Call EnableButtons(False)
    Call Init_ProgressBar
    Call Init_CircularBar
End Sub

Private Sub Cmd_ClassicalPB_Click()
    Call EnableButtons(False)
    Call Init_ProgressBar
End Sub

Function Draw_CircularPB(InternalAngle As Double, ExternalAngle As Double)
    Dim X As Integer, Y As Integer
    Dim X2 As Double, Y2 As Double
    Dim Radius As Double
    Dim i As Long, j As Double, k As Double
    Dim CosX As Double, CosY As Double, SinX As Double, SinY As Double
    Dim FillCircle_Colour As Long
    
    
    On Error Resume Next
    'Define the center of the circle position: X,Y.
    X = (Picture1.Width / 2)
    Y = (Picture1.Height / 2)
    
    'Defines the radius of the circle. Here X=Y=Radius.
    Radius = Picture1.Width / 2
    
    'Use of internal circle.
    If InternalAngle > -1 Then
        For i = LastInternalAngle To InternalAngle
            
            If DrawLineBeforeAngle Then
                'Draw a little line before
                'drawing the new angle position data.
                
                'Convert the angle: radians
                j = ((i + LineAngle) * Pi) / 180
                
                'For some precision reasons we only keep 5 digits.
                CosX = Format(Cos(j), "0.00000")
                SinX = Format(Sin(j), "0.00000")
                For k = 0 To X / 2
                    X2 = X + (k * CosX)
                    Y2 = Y + (k * SinX)
                    Picture3.PSet (X2, Y2), RGB(0, 0, 0)
                Next k
            End If
            
            'Draw the angle data.
            j = (i * Pi) / 180
            CosX = Format(Cos(j), "0.00000")
            SinX = Format(Sin(j), "0.00000")
            
            For k = 0 To X / 2
                X2 = X + (k * CosX)
                Y2 = Y + (k * SinX)
                FillCircle_Colour = Picture2.Point(X2, Y2)
                Picture3.PSet (X2, Y2), FillCircle_Colour
            Next k
        Next i
    End If
    
    
    'Use of external circle.
    If ExternalAngle > -1 Then
        For i = LastExternalAngle To ExternalAngle
            
            If DrawLineBeforeAngle Then
                'Draw a little line before
                'drawing the new angle position data.
                
                'Convert the angle: radians
                j = ((i + LineAngle) * Pi) / 180
                
                'For some precision reasons we only keep 5 digits.
                CosX = Format(Cos(j), "0.00000")
                SinX = Format(Sin(j), "0.00000")
                For k = X / 2 To X
                    X2 = X + (k * CosX)
                    Y2 = Y + (k * SinX)
                    Picture3.PSet (X2, Y2), RGB(0, 0, 0)
                Next k
            End If
        
            'Draw the angle data.
            j = (i * Pi) / 180
    
            CosX = Format(Cos(j), "0.00000")
            SinX = Format(Sin(j), "0.00000")
    
            For k = X / 2 To X
                X2 = X + (k * CosX)
                Y2 = Y + (k * SinX)
                FillCircle_Colour = Picture2.Point(X2, Y2)
                Picture3.PSet (X2, Y2), FillCircle_Colour
            Next k
        Next i
    End If
End Function

Function Restore_InternalCircle()
    Dim X As Integer, Y As Integer
    Dim X2 As Double, Y2 As Double
    Dim Radius As Double
    Dim i As Long, j As Double, k As Double
    Dim CosX As Double, CosY As Double, SinX As Double, SinY As Double
    Dim FillCircle_Colour As Long
    
    
    On Error Resume Next
    'Define the center of the circle position: X,Y.
    X = (Picture1.Width / 2)
    Y = (Picture1.Height / 2)
    
    'Defines the radius of the circle. Here X=Y=Radius.
    Radius = Picture1.Width / 2
    
    'Use the internal circle.
    For i = 0 To MaxAngle_Degrees
        'Convert the angle: radians
        j = (i * Pi) / 180
        
        'For some precision reasons we only keep 5 digits.
        CosX = Format(Cos(j), "0.00000")
        SinX = Format(Sin(j), "0.00000")
        
        For k = 0 To X / 2
            X2 = X + (k * CosX)
            Y2 = Y + (k * SinX)
            FillCircle_Colour = Picture1.Point(X2, Y2)
            Picture3.PSet (X2, Y2), FillCircle_Colour
        Next k
    Next i
End Function

Private Sub Timer1_Timer()
    On Error Resume Next
    If Ext_ClassicalProgressbar.Value <> Ext_ClassicalProgressbar.Max Then
        If Opt_Separated.Value = True Then
            If Int_ClassicalProgressbar.Value = Int_ClassicalProgressbar.Max Then
                Int_ClassicalProgressbar.Value = 0
                Ext_ClassicalProgressbar.Value = Ext_ClassicalProgressbar.Value + NormalPB_Step
            Else
                Int_ClassicalProgressbar.Value = Int_ClassicalProgressbar.Value + NormalPB_Step
            End If
        Else
            Int_ClassicalProgressbar.Value = Int_ClassicalProgressbar.Value + NormalPB_Step
            Ext_ClassicalProgressbar.Value = Ext_ClassicalProgressbar.Value + NormalPB_Step
        End If
    Else
        'Int_ClassicalProgressbar.Value = 0
        'Ext_ClassicalProgressbar.Value = 0
        Timer1.Interval = 0
        Call EnableButtons(True)
    End If
End Sub

Private Sub Timer2_Timer()
    On Error Resume Next
    If SeparatedProgression = True Then
        If ExternalAngle <= (MaxAngle_Degrees - AngleStep) Then
            If InternalAngle >= MaxAngle_Degrees Then
                InternalAngle = 0
                LastInternalAngle = InternalAngle
                ExternalAngle = ExternalAngle + AngleStep
                Call Draw_CircularPB(-1, ExternalAngle)
                LastExternalAngle = ExternalAngle
                If ExternalAngle <= (MaxAngle_Degrees - AngleStep) Then
                    Call Restore_InternalCircle
                End If
            Else
                InternalAngle = InternalAngle + AngleStep
                Call Draw_CircularPB(InternalAngle, -1)
                LastInternalAngle = InternalAngle
            End If
        Else
            InternalAngle = 0
            ExternalAngle = 0
            LastInternalAngle = 0
            LastExternalAngle = 0
            Timer2.Interval = 0
            'Picture3.Picture = Picture1.Picture
            Frm_ProgressionType.Enabled = True
            Call EnableButtons(True)
        End If
    Else
        If ExternalAngle <= (MaxAngle_Degrees - AngleStep) Then
            InternalAngle = InternalAngle + AngleStep
            ExternalAngle = ExternalAngle + AngleStep
            Call Draw_CircularPB(InternalAngle, ExternalAngle)
            LastInternalAngle = InternalAngle
            LastExternalAngle = ExternalAngle
       Else
            InternalAngle = 0
            ExternalAngle = 0
            LastInternalAngle = 0
            LastExternalAngle = 0
            Timer2.Interval = 0
            'Picture3.Picture = Picture1.Picture
            Frm_ProgressionType.Enabled = True
            Call EnableButtons(True)
        End If
    End If
End Sub

Sub Init_ProgressBar()
    On Error Resume Next
    Int_ClassicalProgressbar.Value = 0
    Ext_ClassicalProgressbar.Value = 0
    Int_ClassicalProgressbar.Max = MaxValueNormalPB
    Ext_ClassicalProgressbar.Max = MaxValueNormalPB
    Timer1.Interval = 100
End Sub

Sub Init_CircularBar()
    On Error Resume Next
    'Affecte l'image source à l'image résultat
    Picture3.Picture = Picture1.Picture
    
    LastInternalAngle = 0
    LastExternalAngle = 0
    InternalAngle = 0
    ExternalAngle = 0
    If Opt_Separated.Value = True Then
        SeparatedProgression = True
    ElseIf Opt_Common.Value = True Then
        SeparatedProgression = False
    End If
    Frm_ProgressionType.Enabled = False
    Timer2.Interval = 100
End Sub

Sub EnableButtons(EnableB As Boolean)
    On Error Resume Next
    Cmd_BothBars.Enabled = EnableB
    Cmd_ClassicalPB.Enabled = EnableB
    Cmd_CircularPB.Enabled = EnableB
    Cmd_End.Enabled = EnableB
End Sub
