VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00020202&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Particle System"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   273
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   479
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOpt7 
      Height          =   285
      Left            =   5880
      TabIndex        =   22
      Text            =   "30"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtOpt8 
      Height          =   285
      Left            =   5880
      TabIndex        =   21
      Text            =   "40"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox txtOpt6 
      Height          =   285
      Left            =   5880
      TabIndex        =   19
      Text            =   "45"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtOpt5 
      Height          =   285
      Left            =   5880
      TabIndex        =   17
      Text            =   "30"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtOpt4 
      Height          =   285
      Left            =   5880
      TabIndex        =   14
      Text            =   "3"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtOpt3 
      Height          =   285
      Left            =   5880
      TabIndex        =   13
      Text            =   "0"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox txtOpt2 
      Height          =   285
      Left            =   5880
      TabIndex        =   9
      Text            =   "0.01"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtOpt1 
      Height          =   285
      Left            =   5880
      TabIndex        =   8
      Text            =   "0"
      Top             =   840
      Width           =   1095
   End
   Begin VB.PictureBox picPart 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00020202&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   4440
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   5
      Top             =   720
      Width           =   240
   End
   Begin VB.PictureBox picPart 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00020202&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   4440
      Picture         =   "frmMain.frx":0342
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   4
      Top             =   480
      Width           =   240
   End
   Begin VB.PictureBox picPart 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00020202&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   4440
      Picture         =   "frmMain.frx":0684
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   3
      Top             =   240
      Width           =   240
   End
   Begin VB.PictureBox picPart 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00020202&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   4440
      Picture         =   "frmMain.frx":09C6
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   2
      Top             =   0
      Width           =   240
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      Max             =   10
      TabIndex        =   1
      Top             =   3840
      Value           =   3
      Width           =   4695
   End
   Begin VB.Timer tmrStep 
      Interval        =   1
      Left            =   4200
      Top             =   3120
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00020202&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3840
      Left            =   0
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   312
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      Begin VB.Timer tmrFPS 
         Interval        =   1000
         Left            =   4200
         Top             =   2640
      End
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   26
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Random"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   25
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Life"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   24
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Assurance"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   23
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "P-Age 3"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   20
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "P-Age 2"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   18
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "P-Age"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   16
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "P-Age 1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   15
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Death Offset"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   12
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Death"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   11
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Gravity Offset"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PType"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "P-Update"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00E0E0E0&
      Height          =   3855
      Left            =   4680
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EmitX As Integer
Dim EmitY As Integer

Dim Running As Boolean
Dim FPSCount As Integer
Dim FPS As Integer

Dim Opt1 As Integer
Dim GOffset As Variant
Dim DOffset As Integer
Dim Page1 As Integer
Dim Page2 As Integer
Dim Page3 As Integer
Dim LifeRnd As Integer
Dim LifeAss As Integer

Private Sub Form_Load()
    
    Me.Show

    EmitX = 100
    EmitY = 100

    GOffset = 0.01
    
    Page1 = 3
    Page2 = 30
    Page3 = 45
    
    LifeRnd = 30
    LifeAss = 40

    For i = 0 To UBound(Particle)
        Particle(i).X = EmitX
        Particle(i).Y = EmitY
        Particle(i).Life = Int(Rnd * 30) + 40
        Particle(i).Gravity = (Rnd * 1) + 0.5
        Particle(i).ImageIndex = 0
    Next i
    
    Running = True
    Call StepLoop

End Sub

Private Sub Form_Unload(Cancel As Integer)
Running = False
End
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        EmitX = X
        EmitY = Y
    End If

End Sub


Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picScreen_MouseDown Button, Shift, X, Y
End Sub

Public Sub StepLoop()
Dim FrameSkip As Integer
Dim A As Integer
    Do Until Running = False
        If FrameSkip = HScroll1.Value Then
            picScreen.Cls
            For i = 0 To UBound(Particle)
            
            'Update
                If Not Particle(i).Dead = True Then
                    If Opt1 = 0 Then
                        Particle(i).X = Particle(i).X + Sin(Rnd * 360)
                        Particle(i).Y = Particle(i).Y - (PSPEED + Particle(i).Gravity)
                    ElseIf Opt1 = 1 Then
                        Particle(i).X = Particle(i).X + Cos(Rnd * 360)
                        Particle(i).Y = Particle(i).Y - (PSPEED + Particle(i).Gravity)
                    ElseIf Opt1 = 2 Then
                        Particle(i).X = Particle(i).X + Tan(Rnd * 360)
                        Particle(i).Y = Particle(i).Y - (PSPEED + Particle(i).Gravity)
                    ElseIf Opt1 = 3 Then
                        Particle(i).X = Particle(i).X + 5
                        Particle(i).Y = Particle(i).Y - (PSPEED + Particle(i).Gravity)
                    ElseIf Opt1 = 4 Then
                        Particle(i).X = Particle(i).X - 5
                        Particle(i).Y = Particle(i).Y - (PSPEED + Particle(i).Gravity)
                    ElseIf Opt1 = 5 Then
                        A = Int(Rnd * 360)
                        Particle(i).X = Particle(i).X - Sin(Particle(i).Y)
                        Particle(i).Y = Particle(i).Y + (Sin(A) + Particle(i).Gravity)
                    ElseIf Opt1 = 6 Then
                        A = Rnd * 360
                        Particle(i).X = Particle(i).X - Cos(A)
                        Particle(i).Y = Particle(i).Y + Sin(A)
                    End If
                    
                    Particle(i).Gravity = Particle(i).Gravity - GOffset
                    Particle(i).Age = Particle(i).Age + 1
                End If
            
            'Check For Death
                If Particle(i).Life <= DOffset Then
                    Particle(i).Dead = True
                End If
            
            'P-Age
                If Particle(i).Age > Page1 Then Particle(i).ImageIndex = 1
                If Particle(i).Age > Page2 Then Particle(i).ImageIndex = 2
                If Particle(i).Age > Page3 Then Particle(i).ImageIndex = 3
            
            'Draw
                If Not Particle(i).Dead = True Then
                        Particle(i).Life = Particle(i).Life - 1
                        BitBlt picScreen.hDC, Particle(i).X - 8, Particle(i).Y - 8, 16, 16, picPart(Particle(i).ImageIndex).hDC, 0, 0, vbSrcPaint
                Else
            'Re Life
                    Particle(i).X = EmitX
                    Particle(i).Y = EmitY
                    Particle(i).Life = Int(Rnd * LifeRnd) + LifeAss
                    Particle(i).Dead = False
                    Particle(i).Gravity = (Rnd * 2) + 0.5
                    Particle(i).ImageIndex = 0
                    Particle(i).Age = 0
                End If
                
            Next i
            FrameSkip = -1
        End If
        
        FPSCount = FPSCount + 1
        FrameSkip = FrameSkip + 1
        
        picScreen.CurrentX = 0
        picScreen.CurrentY = 0
        picScreen.Print "FPS: " & FPS
        picScreen.Refresh
        
        DoEvents
    Loop

End Sub

Private Sub tmrFPS_Timer()
    FPS = FPSCount
    FPSCount = -1
End Sub

Private Sub txtOpt1_Change()
    Opt1 = Val(txtOpt1.Text)
End Sub

Private Sub txtOpt2_Change()
    GOffset = Val(txtOpt2.Text)
End Sub

Private Sub txtOpt3_Change()
    DOffset = Val(txtOpt3.Text)
End Sub

Private Sub txtOpt4_Change()
    Page1 = Val(txtOpt4.Text)
End Sub

Private Sub txtOpt5_Change()
    Page2 = Val(txtOpt5.Text)
End Sub

Private Sub txtOpt6_Change()
    Page3 = Val(txtOpt6.Text)
End Sub

Private Sub txtOpt7_Change()
    LifeRnd = Val(txtOpt7.Text)
End Sub

Private Sub txtOpt8_Change()
    LifeAss = Val(txtOpt8.Text)
End Sub
