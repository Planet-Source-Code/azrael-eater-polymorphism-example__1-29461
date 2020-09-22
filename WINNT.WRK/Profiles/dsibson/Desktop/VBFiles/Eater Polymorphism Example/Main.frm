VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLog 
      Height          =   5535
      Left            =   2760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   6135
   End
   Begin VB.Timer tmrEat 
      Interval        =   1
      Left            =   7320
      Top             =   4560
   End
   Begin VB.TextBox txtAnimals 
      Height          =   5535
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Main.frx":0000
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private colAnimals As Collection

Private Const ANIMAL_MAX As Long = 1000

Private Sub Form_Load()
    Dim Template As clsBeast
    
    Dim x As Long
    
    Randomize Timer
    
    Set colAnimals = New Collection
    
    For x = 1 To ANIMAL_MAX
        Select Case Int(Rnd * 3) + 1
            Case 1 'Dog
                Set Template = New DOG
            Case 2 'Cat
                Set Template = New CAT
            Case 3 'Mouse
                Set Template = New MOUSE
        End Select
        
        colAnimals.Add Template
        Log "New " & colAnimals(x).Species & " (" & colAnimals(x).Name & ")"
    Next x
    
    Record
End Sub

Private Sub Log(LogText As String)
    Me.txtLog = LogText & vbCrLf & Me.txtLog
End Sub

Private Sub Record()
    Dim x As Long
    
    Me.txtAnimals = vbNullString
    
    For x = 1 To colAnimals.Count
        Me.txtAnimals = Me.txtAnimals & colAnimals(x).Name & " the " & _
            BeastString(colAnimals(x)) & vbCrLf
    Next x
End Sub

Private Sub tmrEat_Timer()
    Dim EatingBeast As clsBeast
    Dim VictimBeast As clsBeast
    
    Dim Eater As Long
    Dim Victim As Long
    
    If colAnimals.Count > 1 Then
        Eater = Int(Rnd * colAnimals.Count) + 1
        Do
            Victim = Int(Rnd * colAnimals.Count) + 1
        Loop Until Victim <> Eater
        
        Set EatingBeast = colAnimals(Eater)
        Set VictimBeast = colAnimals(Victim)
        
        If EatingBeast.Eat(VictimBeast) Then
            Log EatingBeast.Name & " the " & BeastString(EatingBeast) & _
                " ate " & VictimBeast.Name & " the " & BeastString(VictimBeast)
            colAnimals.Remove Victim
        Else
            Log EatingBeast.Name & " the " & BeastString(EatingBeast) & _
                " failed to eat " & VictimBeast.Name & " the " & _
                BeastString(VictimBeast)
        End If
        
        'Record
    End If
End Sub

Private Function BeastString(Beast As clsBeast) As String
    BeastString = Beast.Species & " (" & Beast.Size & ")"
End Function
