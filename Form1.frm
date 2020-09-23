VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Files Example"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   3585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Delete"
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Search"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear Pointer"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "EMail:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "ID:"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    RAFClear
End Sub

Private Sub Command2_Click()
    If Not RAFAdd(Text1.Text, Text2.Text, Text3.Text) Then
        MsgBox "Failed." & vbCrLf & "ID already in use or pointer file full."
    Else
        Text1.Text = UActive
        Text2.Text = ""
        Text3.Text = ""
    End If
    
End Sub

Private Sub Command3_Click()
    Dim SearchUser As User

    SearchUser = RAFSearch(Text1.Text)
    If SearchUser.ID <> -1 Then
        Text1.Text = RTrim(SearchUser.ID)
        Text2.Text = RTrim(SearchUser.EMail)
        Text3.Text = RTrim(SearchUser.Password)
    Else
        MsgBox "Record not found"
    End If
End Sub

Private Sub Command4_Click()
    If Not RAFDelete(Text1.Text) Then
        MsgBox "Record not found!"
    Else
        MsgBox "Record deleted"
    End If
End Sub

Private Sub Form_Load()
    RAFStartup
End Sub

