VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scheduling Algorithms Test"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2040
      TabIndex        =   30
      Text            =   "10"
      Top             =   1755
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   2040
      List            =   "Form1.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   1185
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Reset"
      Height          =   255
      Left            =   7320
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reset"
      Height          =   255
      Left            =   6120
      TabIndex        =   13
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Top             =   630
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   285
      Left            =   2040
      TabIndex        =   2
      Top             =   630
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "3"
      Top             =   630
      Width           =   375
   End
   Begin VB.Line Line4 
      Index           =   24
      Visible         =   0   'False
      X1              =   2160
      X2              =   2160
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line4 
      Index           =   23
      Visible         =   0   'False
      X1              =   840
      X2              =   840
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line4 
      Index           =   22
      Visible         =   0   'False
      X1              =   7800
      X2              =   7800
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line4 
      Index           =   21
      Visible         =   0   'False
      X1              =   1680
      X2              =   1680
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line4 
      Index           =   20
      Visible         =   0   'False
      X1              =   360
      X2              =   360
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line4 
      Index           =   19
      Visible         =   0   'False
      X1              =   720
      X2              =   720
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line4 
      Index           =   18
      Visible         =   0   'False
      X1              =   1440
      X2              =   1440
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line4 
      Index           =   17
      Visible         =   0   'False
      X1              =   1080
      X2              =   1080
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line4 
      Index           =   16
      Visible         =   0   'False
      X1              =   7440
      X2              =   7440
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line4 
      Index           =   15
      Visible         =   0   'False
      X1              =   7080
      X2              =   7080
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line4 
      Index           =   14
      Visible         =   0   'False
      X1              =   6720
      X2              =   6720
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line4 
      Index           =   13
      Visible         =   0   'False
      X1              =   6360
      X2              =   6360
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line4 
      Index           =   12
      Visible         =   0   'False
      X1              =   6000
      X2              =   6000
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line4 
      Index           =   11
      Visible         =   0   'False
      X1              =   5640
      X2              =   5640
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line4 
      Index           =   10
      Visible         =   0   'False
      X1              =   5280
      X2              =   5280
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line4 
      Index           =   9
      Visible         =   0   'False
      X1              =   4920
      X2              =   4920
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line4 
      Index           =   8
      Visible         =   0   'False
      X1              =   4560
      X2              =   4560
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line4 
      Index           =   7
      Visible         =   0   'False
      X1              =   4200
      X2              =   4200
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line4 
      Index           =   6
      Visible         =   0   'False
      X1              =   3840
      X2              =   3840
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line4 
      Index           =   5
      Visible         =   0   'False
      X1              =   3480
      X2              =   3480
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line4 
      Index           =   4
      Visible         =   0   'False
      X1              =   3120
      X2              =   3120
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line4 
      Index           =   3
      Visible         =   0   'False
      X1              =   2760
      X2              =   2760
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line4 
      Index           =   2
      X1              =   2400
      X2              =   2400
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line4 
      Index           =   1
      X1              =   2040
      X2              =   2040
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Line Line4 
      Index           =   0
      Visible         =   0   'False
      X1              =   120
      X2              =   120
      Y1              =   4560
      Y2              =   4920
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFF80&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   120
      Tag             =   "0"
      Top             =   4560
      Width           =   8055
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   "0.00"
      Height          =   255
      Left            =   7200
      TabIndex        =   55
      Tag             =   "0"
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "0.00"
      Height          =   255
      Left            =   7200
      TabIndex        =   54
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "Average response time :"
      Height          =   255
      Left            =   5040
      TabIndex        =   53
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label10 
      Caption         =   "Average waiting time :"
      Height          =   255
      Left            =   5040
      TabIndex        =   52
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Index           =   6
      Left            =   3480
      TabIndex        =   51
      Tag             =   "0"
      Top             =   4200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   50
      Tag             =   "0"
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Index           =   4
      Left            =   3480
      TabIndex        =   49
      Tag             =   "0"
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Index           =   3
      Left            =   3480
      TabIndex        =   48
      Tag             =   "0"
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   47
      Tag             =   "0"
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   46
      Tag             =   "0"
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Index           =   6
      Left            =   1800
      TabIndex        =   45
      Tag             =   "0"
      Top             =   4200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   44
      Tag             =   "0"
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   43
      Tag             =   "0"
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Index           =   3
      Left            =   1800
      TabIndex        =   42
      Tag             =   "0"
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   41
      Tag             =   "0"
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   40
      Tag             =   "0"
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   39
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   38
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   37
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   36
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   35
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   34
      Top             =   3000
      Width           =   375
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   4560
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label9 
      Caption         =   "Response time"
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   33
      Tag             =   "0"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Waiting time"
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   32
      Tag             =   "0"
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Process"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   31
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Quantum :"
      Height          =   255
      Left            =   1080
      TabIndex        =   29
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Scheduling Algorithm :"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   1245
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   6
      Left            =   7440
      TabIndex        =   26
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   5
      Left            =   7440
      TabIndex        =   25
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   4
      Left            =   7440
      TabIndex        =   24
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   3
      Left            =   7440
      TabIndex        =   23
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   2
      Left            =   7440
      TabIndex        =   22
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   1
      Left            =   7440
      TabIndex        =   21
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Index           =   6
      Left            =   6240
      TabIndex        =   20
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Index           =   5
      Left            =   6240
      TabIndex        =   19
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Index           =   4
      Left            =   6240
      TabIndex        =   18
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Index           =   3
      Left            =   6240
      TabIndex        =   17
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Index           =   2
      Left            =   6240
      TabIndex        =   16
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Index           =   1
      Left            =   6240
      TabIndex        =   15
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "P6"
      Height          =   255
      Index           =   6
      Left            =   5040
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "P5"
      Height          =   255
      Index           =   5
      Left            =   5040
      TabIndex        =   11
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "P4"
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "P3"
      Height          =   255
      Index           =   3
      Left            =   5040
      TabIndex        =   9
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "P2"
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   8
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "P1"
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   7
      Top             =   840
      Width           =   375
   End
   Begin VB.Line Line2 
      Visible         =   0   'False
      X1              =   7080
      X2              =   8160
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      X1              =   4680
      X2              =   7080
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Priority"
      Height          =   255
      Index           =   0
      Left            =   7320
      TabIndex        =   6
      Tag             =   "0"
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Brust Time"
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   5
      Tag             =   "0"
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Process"
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Processes Number :"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   675
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call Randomize
    Call Set_Brust
    Combo1.ListIndex = 0
End Sub

Private Sub Text1_Change()
    Label2(Text1.Text).Visible = True
    Label3(Text1.Text).Visible = True
    Label7(Text1.Text).Visible = True
    Label8(Text1.Text).Visible = True
    Label9(Text1.Text).Visible = True
    Line4(Text1.Text - 1).Visible = True
    If Combo1.ListIndex = 2 Then
        Label4(Text1.Text).Visible = True
        Call Set_Priority
    End If
    If Text1.Text + 1 <= 6 Then
        Label2(Text1.Text + 1).Visible = False
        Label3(Text1.Text + 1).Visible = False
        Label7(Text1.Text + 1).Visible = False
        Label8(Text1.Text + 1).Visible = False
        Label9(Text1.Text + 1).Visible = False
        Line4(Text1.Text).Visible = False
        If Combo1.ListIndex = 2 Then Label4(Text1.Text + 1).Visible = False
    End If
    Call Show_result
End Sub

Private Sub Command1_Click()
    Text1.Text = Text1.Text + 1
    If Text1.Text >= 6 Then Command1.Enabled = False
    If Text1.Text > 2 Then Command2.Enabled = True
End Sub

Private Sub Command2_Click()
    Text1.Text = Text1.Text - 1
    If Text1.Text <= 2 Then Command2.Enabled = False
    If Text1.Text < 6 Then Command1.Enabled = True
End Sub

Private Sub Command3_Click()
    Call Set_Brust
End Sub

Private Sub Command4_Click()
    Call Set_Priority
End Sub

Private Sub Combo1_Click()
    If Combo1.ListIndex = 2 Then
        Call Set_Priority
        For i = 0 To Text1.Text
            Label4(i).Visible = True
        Next
        Command4.Visible = True
        Line2.Visible = True
     Else
        For i = 0 To 6
            Label4(i).Visible = False
        Next
        Command4.Visible = False
        Line2.Visible = False
    End If
    If Combo1.ListIndex = 3 Then
        Label6.Visible = True
        Text2.Visible = True
    Else
        Label6.Visible = False
        Text2.Visible = False
    End If
    Call Show_result
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
    If Text2.Text < 10 Then Text2.Text = 10
    If Text2.Text > 100 Then Text2.Text = 100
    Call Show_result
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    If Not Chr(KeyAscii) Like "[0-9]" Then
        KeyAscii = 0
    End If
End Sub

Public Sub Set_Brust()
    For i = 1 To 6
        Label3(i).Tag = Int(Rnd() * 28 + 3)
        Label3(i).Caption = Label3(i).Tag
    Next
    Call Show_result
End Sub

Public Sub Set_Priority()
    For i = 1 To Text1.Text
        Label4(i).Caption = Int(Rnd() * Text1.Text + 1)
    Next
    Dim Exi As Boolean
    Dim Cut As Boolean
    For Finding = 1 To Text1.Text
        Exi = False
        For itemNow = 1 To Text1.Text
            If Label4(itemNow).Caption = Finding Then
                Exi = True
                Exit For
            End If
        Next
        If Exi = False Then
            Cut = False
            For i = 1 To Text1.Text
                If Label4(i).Caption > Finding Then
                    Label4(i).Caption = Label4(i).Caption - 1
                    Cut = True
                End If
            Next
            If Cut = True Then Finding = Finding - 1
        End If
        If Exi = False And Cut = False Then Exit For
    Next
    Call Show_result
End Sub

Public Sub Show_result()
    For i = 0 To 24
        Line4(i).Visible = False
    Next
    Select Case Combo1.ListIndex
        Case 0
            For i = 1 To Text1.Text
                Label7(i).Caption = Label2(i).Caption
                Label8(i).Tag = Int(Label9(i - 1).Tag)
                Label9(i).Tag = Int(Label8(i).Tag) + Int(Label3(i).Tag)
                Label8(i).Caption = Label8(i).Tag
                Label9(i).Caption = Label9(i).Tag
            Next
            Label12.Caption = 0
            Label13.Caption = 0
            For i = 1 To Text1.Text
                Label12.Caption = Label12.Caption + Int(Label8(i).Tag)
                Label13.Caption = Label13.Caption + Int(Label9(i).Tag)
            Next
            For i = 1 To Text1.Text - 1
                Line4(i).X1 = Label9(i).Tag / Label9(Text1.Text).Tag * 8055 + 120
                Line4(i).X2 = Line4(i).X1
                Line4(i).Visible = True
            Next
        Case 1
            For i = 1 To Text1.Text
                Label7(i).Caption = Label2(i).Caption
                Label8(i).Tag = Label3(i).Tag
            Next
            For i = 1 To Text1.Text - 1
                For j = i + 1 To Text1.Text
                    If Int(Label8(i).Tag) > Int(Label8(j).Tag) Then
                        Temp = Label8(i).Tag
                        Label8(i).Tag = Label8(j).Tag
                        Label8(j).Tag = Temp
                        Temp = Label7(i).Caption
                        Label7(i).Caption = Label7(j).Caption
                        Label7(j).Caption = Temp
                    End If
                Next
            Next
            For i = 1 To Text1.Text - 1
                If Label8(i).Tag = Label8(i + 1).Tag And Label7(i).Caption > Label7(i + 1).Caption Then
                    Temp = Label7(i).Caption
                    Label7(i).Caption = Label7(i + 1).Caption
                    Label7(i + 1).Caption = Temp
                End If
            Next
            For i = 1 To Text1.Text
                Label9(i).Tag = Int(Label8(i).Tag) + Int(Label9(i - 1).Tag)
                Label8(i).Tag = Label9(i - 1).Tag
                Label8(i).Caption = Label8(i).Tag
                Label9(i).Caption = Label9(i).Tag
            Next
            Label12.Caption = 0
            Label13.Caption = 0
            For i = 1 To Text1.Text
                Label12.Caption = Label12.Caption + Int(Label8(i).Tag)
                Label13.Caption = Label13.Caption + Int(Label9(i).Tag)
            Next
            
            For i = 1 To Text1.Text - 1
                Line4(i).X1 = Label9(i).Tag / Label9(Text1.Text).Tag * 8055 + 120
                Line4(i).X2 = Line4(i).X1
                Line4(i).Visible = True
            Next
        Case 2
            For i = 1 To Text1.Text
                Label7(i).Caption = Label2(i).Caption
                Label7(i).Tag = Label4(i).Caption
                Label8(i).Tag = Label3(i).Tag
            Next
            For i = Text1.Text To 2 Step -1
                For j = 1 To i - 1
                    If Label7(j).Tag > Label7(j + 1).Tag Then
                        Temp = Label7(j).Tag
                        Label7(j).Tag = Label7(j + 1).Tag
                        Label7(j + 1).Tag = Temp
                        Temp = Label7(j).Caption
                        Label7(j).Caption = Label7(j + 1).Caption
                        Label7(j + 1).Caption = Temp
                        Temp = Label8(j).Tag
                        Label8(j).Tag = Label8(j + 1).Tag
                        Label8(j + 1).Tag = Temp
                    End If
                Next
            Next
            For i = 1 To Text1.Text
                Label9(i).Tag = Int(Label8(i).Tag) + Int(Label9(i - 1).Tag)
                Label8(i).Tag = Label9(i - 1).Tag
                Label8(i).Caption = Label8(i).Tag
                Label9(i).Caption = Label9(i).Tag
            Next
            Label12.Caption = 0
            Label13.Caption = 0
            For i = 1 To Text1.Text
                Label12.Caption = Label12.Caption + Int(Label8(i).Tag)
                Label13.Caption = Label13.Caption + Int(Label9(i).Tag)
            Next
            For i = 1 To Text1.Text - 1
                Line4(i).X1 = Label9(i).Tag / Label9(Text1.Text).Tag * 8055 + 120
                Line4(i).X2 = Line4(i).X1
                Line4(i).Visible = True
            Next
        Case 3
            Label13.Tag = 0
            For i = 1 To Text1.Text
                Label7(i).Caption = Label2(i).Caption
                Label8(i).Tag = Label3(i).Tag
                Label9(i).Tag = 0
                Label8(i).Caption = 0
                Label9(i).Caption = 0
                Label13.Tag = Label13.Tag + Int(Label3(i).Tag)
            Next
            Shape1.Tag = 0
            For i = 0 To Int(30 / Text2.Text)
                For j = 1 To Text1.Text
                    If Int(Label8(j).Tag) > Int(Text2.Text) Then
                        Label8(j).Caption = Int(Label8(j).Caption) + Int(Shape1.Tag) - Int(Label9(j).Caption)
                        Label9(j).Caption = Int(Shape1.Tag) + Int(Text2.Text)
                        Shape1.Tag = Label9(j).Caption
                        Label8(j).Tag = Int(Label8(j).Tag) - Int(Text2.Text)
                    ElseIf Int(Label8(j).Tag) > 0 Then
                        Label8(j).Caption = Int(Label8(j).Caption) + Int(Shape1.Tag) - Int(Label9(j).Caption)
                        Label9(j).Caption = Int(Shape1.Tag) + Int(Label8(j).Tag)
                        Shape1.Tag = Label9(j).Caption
                        Label8(j).Tag = 0
                    End If
                    Line4(i * Text1.Text + j).X1 = Shape1.Tag / Label13.Tag * 8045 + 120
                    Line4(i * Text1.Text + j).X2 = Line4(i * Text1.Text + j).X1
                    Line4(i * Text1.Text + j).Visible = True
                Next
            Next
            Label12.Caption = 0
            Label13.Caption = 0
            For i = 1 To Text1.Text
                Label12.Caption = Label12.Caption + Int(Label8(i).Caption)
                Label13.Caption = Label13.Caption + Int(Label9(i).Caption)
            Next
    End Select
    Label12.Caption = Int(Label12.Caption / Text1.Text * 100) / 100
    Label13.Caption = Int(Label13.Caption / Text1.Text * 100) / 100
End Sub
