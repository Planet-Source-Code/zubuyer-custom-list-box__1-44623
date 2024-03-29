VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Private StrListArray() As String
Private StrEXListArray() As String
Private lstEXSelection() As Integer
Private List_is_Empty As Boolean

Public Sub Clear()
       ReDim StrListArray(0 To 0)
       ReDim StrEXListArray(0 To 0)
       ReDim lstEXSelection(0 To 0)
       List_is_Empty = True
End Sub


Public Sub RemoveItem(Index As Long)
Const Err_Subscript_Out_Of_Range = 9
Dim IndexID As Long
If UBound(StrListArray) = 0 Then
    StrListArray(0) = ""
    StrEXListArray(0) = ""
    lstEXSelection(0) = vbNull
    List_is_Empty = True
Else
    Select Case Index
    Case Is > ItemCount - 1, Is < 0
        Err.Raise Err_Subscript_Out_Of_Range
    Case Else
        
        For IndexID = Index To UBound(StrListArray) - 1
            StrListArray(IndexID) = StrListArray(IndexID + 1)
            StrEXListArray(IndexID) = StrEXListArray(IndexID + 1)
            lstEXSelection(IndexID) = lstEXSelection(IndexID + 1)
        Next IndexID
    
        ReDim Preserve StrListArray(LBound(StrListArray) To UBound(StrListArray) - 1)
        ReDim Preserve StrEXListArray(LBound(StrEXListArray) To UBound(StrEXListArray) - 1)
        ReDim Preserve lstEXSelection(LBound(lstEXSelection) To UBound(lstEXSelection) - 1)
    End Select
End If
End Sub

Public Sub AddItem(TextLine As String, exTextLine As String, Optional EXSelection As Integer = 0)
If List_is_Empty = False Then
    ReDim Preserve StrListArray(LBound(StrListArray) To UBound(StrListArray) + 1)
    ReDim Preserve StrEXListArray(LBound(StrEXListArray) To UBound(StrEXListArray) + 1)
    ReDim Preserve lstEXSelection(LBound(lstEXSelection) To UBound(lstEXSelection) + 1)

End If
StrEXListArray(UBound(StrListArray)) = exTextLine
StrListArray(UBound(StrListArray)) = TextLine
lstEXSelection(UBound(lstEXSelection)) = EXSelection
List_is_Empty = False
End Sub


Public Sub Sort()
'The fastets sort algorithm!
Dim sVal1 As String, sVal2 As String

Dim Row As Long
Dim MaxRow As Long
Dim MinRow As Long
Dim Swtch As Long
Dim Limit As Long
Dim Offset As Long

MaxRow = UBound(StrListArray)
MinRow = LBound(StrListArray)
Offset = MaxRow \ 2

Do While Offset > 0
      Limit = MaxRow - Offset
      Do
         Swtch = False         ' Assume no switches at this offset.

         ' Compare elements and switch ones out of order:
         
         For Row = MinRow To Limit
                sVal1 = LCase(StrListArray(Row))
                sVal2 = LCase(StrListArray(Row + Offset))
            If sVal1 > sVal2 Then
               strSwap StrListArray(Row), StrListArray(Row + Offset)
               strSwap StrEXListArray(Row), StrEXListArray(Row + Offset)
               intSwap lstEXSelection(Row), lstEXSelection(Row + Offset)
               Swtch = Row
            End If
         Next Row

         ' Sort on next pass only to where last switch was made:
         Limit = Swtch - Offset
      Loop While Swtch

      ' No switches at last offset, try one half as big:
      Offset = Offset \ 2
   Loop
End Sub

Public Property Get ItemCount() As Long
If List_is_Empty Then
    ItemCount = 0
Else
    ItemCount = UBound(StrListArray) + 1
End If
End Property

Public Function Item(ByVal IndexID As Long) As String
    Item = StrListArray(IndexID)
End Function

Public Function exItem(ByVal IndexID As Long) As String
    exItem = StrEXListArray(IndexID)
End Function

Public Function ChangeItem(ByVal IndexID As Long, NewData As String)
    StrListArray(IndexID) = NewData
End Function

Public Function ChangeEXItem(ByVal IndexID As Long, NewData As String)
    StrEXListArray(IndexID) = NewData
End Function

Public Function EXSelection(ByVal IndexID As Long) As Integer
    EXSelection = lstEXSelection(IndexID)
End Function

Public Function SetEXSelection(ByVal IndexID As Long, nEXSelection As Integer) As Integer
    lstEXSelection(IndexID) = nEXSelection
End Function

Private Sub Class_Initialize()
   ReDim StrListArray(0 To 0)
   ReDim StrEXListArray(0 To 0)
   ReDim lstEXSelection(0 To 0)
   
   List_is_Empty = True
End Sub

Private Sub strSwap(ByRef var1 As String, ByRef var2 As String)
    Dim X As String
    X = var1
    var1 = var2
    var2 = X
End Sub

Private Sub intSwap(ByRef var1 As Integer, ByRef var2 As Integer)
    Dim X As Integer
    X = var1
    var1 = var2
    var2 = X
End Sub

