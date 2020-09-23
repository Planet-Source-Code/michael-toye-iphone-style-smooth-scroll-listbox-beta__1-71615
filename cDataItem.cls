VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDataItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mKey As String
Private mValue As String
Private mValue2 As String
Private mIcon As cDIBSection
Private mTicked As Boolean
Private mDeleted As Boolean
Private mPrevSortItem As Long
Private mNextSortItem As Long
Private mIsRoot As Boolean

Public Property Let Deleted(v As Boolean)
mDeleted = v
End Property
Public Property Get Deleted() As Boolean
Deleted = mDeleted
End Property


Public Property Let IsRoot(v As Boolean)
mIsRoot = v
End Property
Public Property Get IsRoot() As Boolean
IsRoot = mIsRoot
End Property
Public Property Let PrevSortItem(sV As Long)
    mPrevSortItem = sV
End Property
Public Property Get PrevSortItem() As Long
    PrevSortItem = mPrevSortItem
End Property
Public Property Let NextSortItem(sV As Long)
    mNextSortItem = sV
End Property
Public Property Get NextSortItem() As Long
    NextSortItem = mNextSortItem
End Property

Public Property Let Key(sV As String)
    mKey = sV
End Property
Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Value(sV As String)
    mValue = sV
End Property
Public Property Get Value() As String
    Value = mValue
End Property
Public Property Let Ticked(sV As Boolean)
    mTicked = sV
End Property
Public Property Get Ticked() As Boolean
    Ticked = mTicked
End Property
Public Property Set Key(sV As Object)
    Set mIcon = sV
End Property
Public Property Let Value2(sV As String)
    mValue2 = sV
End Property
Public Property Get Value2() As String
    Value2 = mValue2
End Property
