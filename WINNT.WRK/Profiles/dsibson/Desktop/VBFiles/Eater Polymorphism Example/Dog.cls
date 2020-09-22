VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsBeast

Private mLngSize As Long
Private mName As String

Private Const DOG_SIZE As Long = 3
Private Const DOG As String = "Dog"

Private Sub Class_Initialize()
    mLngSize = DOG_SIZE
    Select Case Int(Rnd * 3)
        Case 0
            mName = "Ratchet"
        Case 1
            mName = "Sasha"
        Case 2
            mName = "Buck"
    End Select
End Sub

Private Function clsBeast_Eat(Victim As clsBeast) As Boolean
    If Victim.Size < mLngSize Then
        mLngSize = mLngSize + Victim.Size
        clsBeast_Eat = True
    Else
        clsBeast_Eat = False
    End If
End Function

Private Property Let clsBeast_Size(NewSize As Long)
    mLngSize = NewSize
End Property

Private Property Get clsBeast_Size() As Long
    clsBeast_Size = mLngSize
End Property

Private Property Get clsBeast_Species() As String
    clsBeast_Species = DOG
End Property

Private Property Get clsBeast_Name() As String
     clsBeast_Name = mName
End Property
