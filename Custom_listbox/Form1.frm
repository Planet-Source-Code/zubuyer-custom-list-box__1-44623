VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom List Box"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   ClipControls    =   0   'False
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3390
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   465
      Width           =   1830
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00F4F5F7&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   60
      ScaleHeight     =   166
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   220
      TabIndex        =   7
      Top             =   45
      Width           =   3330
      Begin VB.VScrollBar ScrollBar 
         Height          =   1635
         Left            =   2235
         Max             =   0
         TabIndex        =   8
         Top             =   180
         Width           =   270
      End
   End
   Begin VB.CommandButton cmdShort 
      Caption         =   "&Short"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3390
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   45
      Width           =   1830
   End
   Begin VB.CommandButton cmdRemove 
      Cancel          =   -1  'True
      Caption         =   "&Remove"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3390
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   900
      Width           =   1830
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3390
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   1830
   End
   Begin VB.TextBox txtAdd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3390
      TabIndex        =   0
      Top             =   1740
      Width           =   1830
   End
   Begin VB.Label lblTest 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AaGgWwi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3630
      TabIndex        =   11
      Top             =   3285
      Width           =   1290
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3390
      TabIndex        =   9
      Top             =   2820
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3390
      TabIndex        =   4
      Top             =   2385
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3390
      TabIndex        =   3
      Top             =   2595
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3390
      TabIndex        =   2
      Top             =   2175
      Width           =   90
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hightest list item = 32768 [this is way enough](isn't it?)
Option Explicit
'Variable Declarations
Dim cList As New cList
Dim NormalSelection
Dim ExtendedSelection
Dim PreviousSelection
Dim ListRange                           'Range of list that will be displayed
Dim LstTextHeight
'List Settings
Const ItemSwap As Boolean = False       'Determines if the list items can be moved by dragging or not
'Color Presets
Const NormalText = &H808080             'Color of normal text
Const SelectedText = &HFFFFFF           'Text color of normal selection
Const ExtendedSelectionText = &H808080  'Text color of a list item when selected by double clicking
Const SelectedBG = &HD2D8D9             'The background color of a selection

Private Sub cmdAdd_Click()
cList.AddItem txtAdd.Text
ReinitializeList
Label2.Caption = "ListCount: " & cList.ItemCount
End Sub

Private Sub cmdClear_Click()
cList.Clear
ReinitializeList
End Sub

Private Sub cmdRemove_Click()
Dim i
Dim j As String
cList.RemoveItem (NormalSelection)
If NormalSelection = cList.ItemCount Then NormalSelection = NormalSelection - 1
ReinitializeList
Label2.Caption = "ListCount: " & cList.ItemCount
End Sub

Private Sub cmdShort_Click()
cList.Sort
ReinitializeList
End Sub

Private Sub Form_DblClick()
End
End Sub

Private Sub Form_Load()
LstTextHeight = picList.TextHeight("")
Debug.Print LstTextHeight
Dim i
For i = 0 To Screen.FontCount - 1
cList.AddItem Screen.Fonts(i), Screen.Fonts(i)
Next

Label2.Caption = "ListCount: " & cList.ItemCount
NormalSelection = -1
ExtendedSelection = -1
End Sub

Private Sub Form_Resize()
picList.Move 0, 0, picList.Width, Me.ScaleHeight
ScrollBar.Move picList.ScaleWidth - ScrollBar.Width, 0, ScrollBar.Width, picList.ScaleHeight
ListRange = Fix(picList.ScaleHeight / LstTextHeight)
ReinitializeList
End Sub

Private Sub picList_Click()
lblTest.FontName = cList.Item(NormalSelection)
End Sub

Private Sub picList_DblClick()
On Error Resume Next
cList.SetEXSelection ExtendedSelection, 0
ExtendedSelection = NormalSelection
Debug.Print NormalSelection
cList.SetEXSelection NormalSelection, 1
ReinitializeList
Label4.Caption = cList.Item(ExtendedSelection)
End Sub

Private Sub picList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Fix(Y / LstTextHeight) + ScrollBar.Value >= 0 And Fix(Y / LstTextHeight) + ScrollBar.Value < cList.ItemCount Then
    NormalSelection = Fix(Y / LstTextHeight) + ScrollBar.Value
    If cList.ItemCount > 0 Then
        Label1.Caption = NormalSelection
        Label3.Caption = cList.Item(NormalSelection) & cList.exItem(NormalSelection)
    ReinitializeList
    End If
    PreviousSelection = NormalSelection
Else
NormalSelection = -1
ReinitializeList
End If
End Sub

Private Sub picList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ItemBuffer As String
Dim EXSelectionBuffer As Integer

If Button And Fix(Y / LstTextHeight) + ScrollBar.Value >= 0 And Fix(Y / LstTextHeight) + ScrollBar.Value < cList.ItemCount Then
    NormalSelection = Fix(Y / LstTextHeight) + ScrollBar.Value
       
    If NormalSelection > (ScrollBar.Value + (ListRange - 1)) Then
        ScrollBar.Value = ScrollBar.Value + 1
    ElseIf NormalSelection < ScrollBar.Value Then
        ScrollBar.Value = ScrollBar.Value - 1
    End If

    If ItemSwap Then
    ItemBuffer = cList.Item(NormalSelection)
    EXSelectionBuffer = cList.EXSelection(NormalSelection)
    
    cList.ChangeItem NormalSelection, cList.Item(PreviousSelection)
    cList.SetEXSelection NormalSelection, cList.EXSelection(PreviousSelection)
    
    cList.ChangeItem PreviousSelection, ItemBuffer
    cList.SetEXSelection PreviousSelection, EXSelectionBuffer
    
    PreviousSelection = NormalSelection
    End If
    
    If cList.ItemCount > 0 Then
        Label1.Caption = NormalSelection
        Label3.Caption = cList.Item(NormalSelection) & cList.exItem(NormalSelection)
        ReinitializeList
    End If
End If
End Sub

Private Sub ScrollBar_Change()
ReinitializeList
End Sub

Private Sub ScrollBar_Scroll()
ReinitializeList
End Sub

Sub ReinitializeList()
Dim i
'Sets the scrollbar max value as the list is changed
'so the list items canbe scrolled
If cList.ItemCount > ListRange Then
    ScrollBar.Max = cList.ItemCount - ListRange 'The list items are too much to _
                                                 display we need to scroll the list
Else
    ScrollBar.Max = 0                           'There is no need for scrolling
End If

picList.Cls
'Gradient
If cList.ItemCount <= ListRange Then
    For i = 0 To cList.ItemCount - 1
        If i = NormalSelection Then
            DrawSelection i
            picList.CurrentX = 0
            picList.CurrentY = NormalSelection * LstTextHeight
            
            If cList.EXSelection(i) = 1 Then
            picList.ForeColor = ExtendedSelectionText
            picList.FontBold = True
            Else
            picList.ForeColor = SelectedText
            End If
            
            picList.Print " " & i + 1 & ". " & cList.Item(i)
            picList.FontBold = False
        ElseIf cList.EXSelection(i) = 1 Then
            picList.CurrentX = 0
            picList.CurrentY = i * LstTextHeight
            picList.ForeColor = ExtendedSelectionText
            picList.FontBold = True
            picList.Print " " & i + 1 & ". " & cList.Item(i)
            picList.FontBold = False
            ExtendedSelection = i
        Else
            picList.CurrentX = 0
            picList.CurrentY = i * LstTextHeight
            picList.ForeColor = NormalText
            picList.Print " " & i + 1 & ". " & cList.Item(i)
        End If
    Next
Else
    For i = ScrollBar.Value To ScrollBar.Value + (ListRange - 1)
        If i = NormalSelection Then
            DrawSelection (i - ScrollBar.Value)
            picList.CurrentX = 0
            picList.CurrentY = (i - ScrollBar.Value) * LstTextHeight
            
            If cList.EXSelection(i) = 1 Then
            picList.ForeColor = ExtendedSelectionText
            picList.FontBold = True
            Else
            picList.ForeColor = SelectedText
            End If
            
            picList.Print " " & i + 1 & ". " & cList.Item(i)
            picList.FontBold = False
        ElseIf cList.EXSelection(i) = 1 Then
            picList.CurrentX = 0
            picList.CurrentY = (i - ScrollBar.Value) * LstTextHeight
            picList.ForeColor = ExtendedSelectionText
            picList.FontBold = True
            picList.Print " " & i + 1 & ". " & cList.Item(i)
            picList.FontBold = False
            ExtendedSelection = i
        Else
            picList.CurrentX = 0
            picList.CurrentY = (i - ScrollBar.Value) * LstTextHeight
            picList.ForeColor = NormalText
            picList.Print " " & i + 1 & ". " & cList.Item(i)
        End If
    Next
End If
End Sub

Sub DrawSelection(Y)
Dim i
For i = 0 To LstTextHeight
 picList.Line (0, i + (Y * LstTextHeight))-(picList.ScaleWidth - 19, i + (Y * LstTextHeight)), RGB(((LstTextHeight - i) * 3) + 183, ((LstTextHeight - i) * 3) + 182, ((LstTextHeight - i) * 3) + 180)
Next
picList.Line (0, (Y * LstTextHeight))-(picList.ScaleWidth - 19, (Y * LstTextHeight) + LstTextHeight), &H929EA3, B
End Sub

Sub Gradient()
Dim i
For i = 0 To picList.ScaleHeight Step LstTextHeight / 2
 picList.Line (0, i)-(picList.ScaleWidth - 18, i + LstTextHeight / 2), RGB(i / 6 + 202, i / 6 + 201, i / 6 + 200), BF
Next
End Sub

