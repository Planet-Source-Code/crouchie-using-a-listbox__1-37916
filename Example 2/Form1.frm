VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Listbox Example"
   ClientHeight    =   2445
   ClientLeft      =   5460
   ClientTop       =   3600
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   3930
   Begin VB.CommandButton Command4 
      Caption         =   "<<"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">>"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin VB.ListBox List2 
      Height          =   1620
      Left            =   2400
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Don't forget to vote"
      Height          =   255
      Left            =   1238
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Email: crouchie1998@hotmail.com"
      Height          =   255
      Left            =   705
      TabIndex        =   6
      Top             =   2160
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This project shows how to use Listboxes
'Copyright (C) 1999 - 2002 Crouchman & Son Computer System
'Email crouchie1998@hotmail.com
'Don't forget to vote!

Option Explicit
Dim intlist1 As Integer
Dim intItem As Integer

Private Sub Command1_Click()
'Check to see if the user has selected an item. If so, move
'it to List2, if they have.
If List1.ListIndex = -1 Then Exit Sub
List2.AddItem (List1.Text)
List1.RemoveItem (List1.ListIndex)
End Sub

Private Sub Command2_Click()
'Check to see if the user has selected an item. If so, move
'it to List1, if they have.
If List2.ListIndex = -1 Then Exit Sub
List1.AddItem (List2.Text)
List2.RemoveItem (List2.ListIndex)
End Sub

Private Sub Command3_Click()
'Move the contents of List1 to List2 & clear the contents
'of List1
For intlist1 = 0 To List1.ListCount - 1
    List2.AddItem (List1.List(intlist1))
Next
List1.Clear
End Sub

Private Sub Command4_Click()
'Move the contents of List2 to List1 & clear the contents
'of List2
For intlist1 = 0 To List2.ListCount - 1
    List1.AddItem (List2.List(intlist1))
Next
List2.Clear
End Sub

Private Sub Form_Load()
'Populate List1 with items
List1.Clear
For intItem = 1 To 8
    List1.AddItem "Item " & intItem
Next
End Sub
