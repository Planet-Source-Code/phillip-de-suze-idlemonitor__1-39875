VERSION 5.00
Begin VB.UserControl IdleMon 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   InvisibleAtRuntime=   -1  'True
   Picture         =   "IdleMon.ctx":0000
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   32
   ToolboxBitmap   =   "IdleMon.ctx":2372
   Begin VB.Timer tmrStateMonitor 
      Interval        =   1
      Left            =   840
      Top             =   0
   End
   Begin VB.Timer tmrPeriod 
      Interval        =   1
      Left            =   480
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   2
      Left            =   240
      Picture         =   "IdleMon.ctx":2684
      Top             =   240
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   0
      Picture         =   "IdleMon.ctx":49F6
      Top             =   240
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   240
      Picture         =   "IdleMon.ctx":6D68
      Top             =   0
      Width           =   240
   End
End
Attribute VB_Name = "IdleMon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Dim IsIdle As Boolean 'True when idling or While in idle-state
Dim MousePos As POINTAPI 'holds mouse position
Dim startOfIdle As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

'Default Property Values:
Const m_def_Interval = 60
'Property Variables:
Dim m_Interval As Variant
'Event Declarations:
Event IdleStateDisengaged(ByVal IdleStopTime As Long)
Event IdleStateEngaged(ByVal IdleStartTime As Long)
Private Sub tmrPeriod_Timer()
    If IsIdle Then
       If Timer - startOfIdle >= Interval Then
            RaiseEvent IdleStateEngaged(Timer)
            'important: set the values
            startOfIdle = Timer
            IsIdle = True
        End If
    Else
        RaiseEvent IdleStateDisengaged(Timer)
    End If
End Sub
Private Sub tmrStateMonitor_Timer()
    Dim state As Integer
    Dim tmpPos As POINTAPI
    Dim ret As Long 'simply holds the return value of the API
   
    Dim IdleFound As Boolean
    Dim i As Integer 'the counter uses by the For Loop
    IdleFound = False
   
    For i = 1 To 256
        state = GetAsyncKeyState(i)
        If state = -32767 Then
            IdleFound = True 'means that something is withholding the computer of idling
            IsIdle = False 'thus, it is Not idling, so Set the value
        End If
        DoEvents
    Next

    ret = GetCursorPos(tmpPos)

    If tmpPos.x <> MousePos.x Or tmpPos.y <> MousePos.y Then
        IsIdle = False 'set the...
        IdleFound = True 'values
        MousePos.x = tmpPos.x
        MousePos.y = tmpPos.y
    End If

    If Not IdleFound Then
        If Not IsIdle Then
            IsIdle = True
            startOfIdle = Timer
        End If
    End If
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,60
Public Property Get Interval() As Variant
    Interval = m_Interval
End Property
Public Property Let Interval(ByVal New_Interval As Variant)
    m_Interval = New_Interval
    PropertyChanged "Interval"
End Property
Private Sub UserControl_Initialize()
    UserControl.Height = 480
    UserControl.Width = 480
End Sub
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Interval = m_def_Interval
End Sub
'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Interval = PropBag.ReadProperty("Interval", m_def_Interval)
End Sub
Private Sub UserControl_Resize()
    UserControl.Height = 480
    UserControl.Width = 480
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Interval", m_Interval, m_def_Interval)
End Sub

