VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Test ObjPtr"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   3015
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnWeakObjPtrTestAssignSwap 
      Caption         =   "Weak ObjPtr Test AssignSwap"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton BtnWeakObjPtrTest1 
      Caption         =   "Weak ObjPtr Test1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Obj1 As Class1
'

Private Sub BtnWeakObjPtrTest1_Click()
    
    Set m_Obj1 = MNew.Class1(123.456)
    MsgBox "Creating a new object m_Obj1 of Class1 with value: " & m_Obj1.ToStr
    
    Dim pObj1 As LongPtr: pObj1 = ObjPtr(m_Obj1)
    
    MsgBox "The ObjPtr of m_Obj1 is pObj1, it's value is: " & pObj1
    
    MPtr.ZeroObject m_Obj1
    MsgBox "ZeroObject m_Obj1, now m_Obj1 is: ..."
    
    If m_Obj1 Is Nothing Then
        MsgBox "m_Obj1 is nothing"
    Else
        MsgBox "m_Obj1 is*not* nothing, the value of m_Obj1 is: " & m_Obj1.ToStr
    End If
    
    Set m_Obj1 = MPtr.PtrToObject(pObj1)
    MsgBox "Now we write pObj1, remember pObj1 is: " & pObj1 & vbCrLf & _
           "back to the object m_Obj1, now m_Obj1 is: ..."
    
    If m_Obj1 Is Nothing Then
        MsgBox "m_Obj1 is nothing"
    Else
        MsgBox "m_Obj1 is*not* nothing, the value of m_Obj1 is: " & m_Obj1.ToStr
    End If
    
End Sub

Private Sub BtnWeakObjPtrTestAssignSwap_Click()
    
    Dim Obj1 As Class1: Set Obj1 = MNew.Class1(123.456)
    Dim Obj2 As Class1: Set Obj2 = MNew.Class1(456.789)
    
    MsgBox "We created 2 objects Obj1 and Obj2: " & vbCrLf & _
           "Obj1.Value = " & Obj1.Value & vbCrLf & _
           "Obj2.Value = " & Obj2.Value
           
    MPtr.AssignSwap Obj1, Obj2
    
    MsgBox "After AssignSwap Obj1, Obj2: " & vbCrLf & _
           "Obj1.Value = " & Obj1.Value & vbCrLf & _
           "Obj2.Value = " & Obj2.Value
End Sub
