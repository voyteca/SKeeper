VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SKUF 
   Caption         =   "SKeeper - keep your selection"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4200
   OleObjectBlob   =   "SKUF.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SKUF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********  SKeeper (Selection Keeper)  ************
'
'
'Author: Wojciech Czubak
'
'Date: 25/07/2011
'
'Description: Simple macro for saving selections
'
'
'*****************************************

Private Sub addObjectB_Click()
    If groupListLB.ListIndex = -1 Then Exit Sub
    Dim gName As String
    gName = groupListLB.List(groupListLB.ListIndex)
    addShapeToGroup ActiveSelection.Shapes.All, gName
End Sub

Private Sub deleteGroupB_Click()
    If groupListLB.ListCount = 0 Or groupListLB.ListIndex = -1 Then Exit Sub
    Dim gName As String
    gName = groupListLB.List(groupListLB.ListIndex)
    deleteGroup gName
End Sub

Private Sub groupListLB_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' if "Alt" key is pressed then no action is processed on drawing area,
    ' this allows to change listBox selection without changing object selection.
    ' Right mouse button might casue problems so exits Sub if it is detected
    If Shift = 4 Or Button = 2 Or groupListLB.ListIndex = -1 Then Exit Sub
    Dim gName As String
    Dim add As Boolean
    gName = groupListLB.List(groupListLB.ListIndex)
    If Shift = 1 Then add = True
    selectGroup gName, add
End Sub


Private Sub newGroupB_Click()
    createNewGroup
End Sub

Private Sub removeObjectB_Click()
    If groupListLB.ListIndex = -1 Then Exit Sub
    Dim gName As String
    gName = groupListLB.List(groupListLB.ListIndex)
    removeShapeFromGroup ActiveSelection.Shapes.All, gName
End Sub

Private Sub UserForm_Initialize()
    startUp
End Sub

