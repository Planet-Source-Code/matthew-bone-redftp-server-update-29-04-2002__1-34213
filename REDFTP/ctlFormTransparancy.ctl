VERSION 5.00
Begin VB.UserControl FormTransparancy 
   ClientHeight    =   1095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1350
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ctlFormTransparancy.ctx":0000
   ScaleHeight     =   1095
   ScaleWidth      =   1350
   ToolboxBitmap   =   "ctlFormTransparancy.ctx":02AA
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   840
      Top             =   600
   End
End
Attribute VB_Name = "FormTransparancy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const m_def_TransparencyLevel = 0
Const m_def_TransparencyDirection = 0

Dim m_TransparencyLevel As Integer
Dim m_TransparencyDirection As Integer

Private Sub UserControl_Initialize()
    m_TransparencyLevel = 0
    m_TransparencyDirection = 0
End Sub

Private Sub UserControl_InitProperties()
    m_TransparencyLevel = m_def_TransparencyLevel
    m_TransparencyDirection = m_def_TransparencyDirection
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_TransparencyLevel = PropBag.ReadProperty("TransparencyLevel", m_def_TransparencyLevel)
    m_TransparencyDirection = PropBag.ReadProperty("TransparencyDirection", m_def_TransparencyDirection)
      MakeTransparent UserControl.Parent.hWnd, m_TransparencyLevel

End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("TransparencyLevel", m_TransparencyLevel, m_def_TransparencyLevel)
    Call PropBag.WriteProperty("TransparencyDirection", m_TransparencyDirection, m_def_TransparencyDirection)
End Sub

Private Sub Timer1_Timer()

  If m_TransparencyDirection <> 0 Then
    If MakeTransparent(UserControl.Parent.hWnd, m_TransparencyLevel) = 1 Then
      If m_TransparencyDirection < 0 Then UserControl.Parent.Visible = False
    End If
    m_TransparencyLevel = m_TransparencyLevel + m_TransparencyDirection
    If m_TransparencyLevel < Abs(m_TransparencyDirection) Then
      m_TransparencyDirection = 0
      m_TransparencyLevel = 0
      MakeTransparent UserControl.Parent.hWnd, m_TransparencyLevel
      Unload UserControl.Parent
    End If
    If m_TransparencyLevel > (255 - Abs(m_TransparencyDirection)) Then
      m_TransparencyDirection = 0
      m_TransparencyLevel = 255
    End If
  End If
  
End Sub


Public Property Get TransparencyDirection() As Long
    TransparencyDirection = m_TransparencyDirection
End Property

Public Property Let TransparencyDirection(ByVal New_TransparencyDirection As Long)
    m_TransparencyDirection = New_TransparencyDirection
    PropertyChanged "TransparencyDirection"
End Property


Public Property Get TransparencyLevel() As Long
    TransparencyLevel = m_TransparencyLevel
End Property

Public Property Let TransparencyLevel(ByVal New_TransparencyLevel As Long)
    m_TransparencyLevel = New_TransparencyLevel
    PropertyChanged "TransparencyLevel"
End Property

Public Function MakeVisible() As Variant
    m_TransparencyLevel = 0
    m_TransparencyDirection = 4
    MakeTransparent UserControl.Parent.hWnd, m_TransparencyLevel
    UserControl.Parent.Visible = True
    UserControl.Parent.SetFocus
End Function

Public Function MakeInVisible() As Variant
    m_TransparencyLevel = 255
    m_TransparencyDirection = -4
    MakeTransparent UserControl.Parent.hWnd, m_TransparencyLevel
End Function

