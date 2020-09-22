VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsBeast

Private mLngSize As Long
Private mName As String

Private Const MOUSE_SIZE As Long = 1
Private Const MOUSE As String = "Mouse"

Private Function clsBeast_Eat(Victim As clsBeast) As Boolean
    'Mice can eat victims of their own size!
    If Victim.Size <= mLngSize Then
        mLngSize = mLngSize + Victim.Size
        clsBeast_Eat = True
    Else
        clsBeast_Eat = False
    End If
End Function

Private Property Get clsBeast_Name() As String
    clsBeast_Name = mName
End Property

Private Property Let clsBeast_Size(NewSize As Long)
    mLngSize = NewSize
End Property

Private Property Get clsBeast_Size() As Long
    clsBeast_Size = mLngSize
End Property

Private Sub Class_Initialize()
    mLngSize = MOUSE_SIZE
    Select Case Int(Rnd * 3)
        Case 0
            mName = "Mickey"
        Case 1
            mName = "Jerry"
        Case 2
            mName = "Eric"
    End Select
End Sub

Private Property Get clsBeast_Species() As String
    clsBeast_Species = MOUSE
End Property
