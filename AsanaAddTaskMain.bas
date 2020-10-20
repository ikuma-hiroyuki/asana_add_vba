Attribute VB_Name = "AsanaAddTaskMain"
Option Explicit

Private Sub AddAsanaTaskTest()
    Dim asana As AsanaConnect: Set asana = New AsanaConnect

    With asana
        .TaskName = "test"
        .TaskNote = "hoge"
        .Assignee = "ikuma-hiroyuki@sknc.co.jp"
        .DueDate = Date
        .AddTask
    End With
End Sub

Public Sub AddAsanaTask()
    Dim asana As AsanaConnect: Set asana = New AsanaConnect

    Dim rg As Range
    For Each rg In Selection
        With asana
            .TaskName = Cells(rg.row, 1)
            .TaskNote = Cells(rg.row, 2)
            .Assignee = Cells(rg.row, 3)
            .DueDate = Cells(rg.row, 4)
            .AddTask
        End With
    Next
End Sub
