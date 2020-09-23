VERSION 5.00
Begin VB.Form frmCalculations 
   Caption         =   "Projectile Calculations Version 2"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Launch Path"
      Height          =   4575
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   9495
      Begin VB.PictureBox picDisplay 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   4095
         Left            =   120
         ScaleHeight     =   4035
         ScaleWidth      =   9195
         TabIndex        =   9
         Top             =   360
         Width           =   9255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "General Information"
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   5655
      Begin VB.PictureBox picInfo 
         AutoRedraw      =   -1  'True
         Height          =   1455
         Left            =   240
         ScaleHeight     =   1395
         ScaleWidth      =   5115
         TabIndex        =   6
         Top             =   360
         Width           =   5175
      End
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calculate"
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtAngle 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txtVelocity 
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   3015
      Left            =   6000
      TabIndex        =   7
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Launch Angle"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Velocity"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmCalculations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCalculate_Click()
    Dim RangeOfProjectile As Double             'I beleive that the names of these
    Dim TimeTillImpact As Double                'variables are self explanatory
    Dim TrajectoryApex As Double
    Dim ProjectileVelocity As Double            'The arrays are for holding point
    Dim LaunchAngle As Double                   'Coords for the graphing code
    Dim XcordArray(1 To 100) As Double
    Dim YcordArray(1 To 100) As Double          'i is just a counter variable
    Dim i As Integer
    Dim FractionOfTime As Double
    
    ProjectileVelocity = txtVelocity.Text
    LaunchAngle = Deg2Rad(txtAngle.Text)        'Convert degrees to radians for calculations
    
    'See modprojectiles for more info on these functions
    TimeTillImpact = TotalTime(ProjectileVelocity, LaunchAngle)
    RangeOfProjectile = GetRange(ProjectileVelocity, LaunchAngle, TimeTillImpact)
    TrajectoryApex = Trajectory_Apex(ProjectileVelocity, LaunchAngle)
    FractionOfTime = TimeTillImpact / 100

    picInfo.Cls
    picInfo.Print "Launch Velocity : " & ProjectileVelocity & " feet per second"
    picInfo.Print "Launch Angle : " & LaunchAngle & " Radians"
    picInfo.Print "Range : " & RangeOfProjectile & " feet"
    picInfo.Print "Will Impact in : " & TimeTillImpact & " seconds"
    picInfo.Print "Apex of : " & TrajectoryApex & " feet"
    
    'If ProjectileVelocity = 3000 Then
        'The graph looks screwed up if a greater number is used (picbox is too small)
        'I will later add a function for keeping the picture the same size regardless
        'Of what speed you put in, and only change it for angles
        For i = 1 To 100
            XcordArray(i) = GetXPos(ProjectileVelocity, LaunchAngle, (FractionOfTime * i))
            YcordArray(i) = GetYPos(ProjectileVelocity, LaunchAngle, (FractionOfTime * i))
        Next i
    
        picDisplay.Cls
        picDisplay.Line (200, TrajectoryApex + 200)-(RangeOfProjectile, TrajectoryApex + 200)
        picDisplay.Line ((RangeOfProjectile + 200) / 2, TrajectoryApex + 200)-((RangeOfProjectile + 200) / 2, 200)

        'picCoord.Cls
        'For i = 1 To 100                                   'This section was used for
            'picCoord.Print XcordArray(i), YcordArray(i)    'Debugging
        'Next i

        For i = 1 To 100            'Plot the flight path of the projectile
            picDisplay.PSet (XcordArray(i), ((TrajectoryApex + 200) - YcordArray(i))), vbRed
        Next i
    'End If
End Sub

Private Sub txtAngle_Validate(Cancel As Boolean)
If IsNumeric(txtAngle.Text) = False Then
    MsgBox "Please enter only numeric values in the feilds", , "Error"
    Cancel = True
End If
End Sub

Private Sub txtVelocity_Validate(Cancel As Boolean)
If IsNumeric(txtVelocity.Text) = False Then
    MsgBox "Please enter only numeric values in the feilds", , "Error"
    Cancel = True
End If
End Sub
