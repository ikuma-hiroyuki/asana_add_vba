VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AsanaConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Register a task in asana (duplication of task)
'See github below for the registration program itself
'https://github.com/ikuma-hiroyuki/dpulicate_asana_python

Option Explicit

'Private shell As Object

Private taskName_ As String
Private taskNote_ As String
Private dueDate_ As String
Private assignee_ As String

Public Sub AddTask()
    Dim asanaPy As String: asanaPy = "python C:\Users\ikuma\Python\asana_tasks\duplicate_task_asana.py"
    Dim args As Variant: args = Join(Array(Me.TaskName, Me.TaskNote, Me.DueDate, Me.Assignee), " ")
    Dim commandLineString As String: commandLineString = asanaPy & " " & args
    shell commandLineString, vbNormalFocus 'îÒìØä˙èàóù
End Sub

Public Property Get TaskName() As String
    TaskName = taskName_
End Property

Public Property Let TaskName(ByVal addTaskName As String)
    taskName_ = IIf(addTaskName <> vbNullString, "-t " & addTaskName, vbNullString)
End Property

Public Property Get TaskNote() As String
    TaskNote = taskNote_
End Property

Public Property Let TaskNote(ByVal addNote As String)
    taskNote_ = IIf(addNote <> vbNullString, "-n " & addNote, vbNullString)
End Property

Public Property Get DueDate() As String
    DueDate = dueDate_
End Property

Public Property Let DueDate(ByVal addDueDate As String)
    If addDueDate <> vbNullString Then
        If IsDate(addDueDate) Then
            dueDate_ = "-d " & Format(addDueDate, "yyyy-mm-dd")
        Else
            Err.Raise 999, , "Please enter the date"
        End If
    End If
End Property

Public Property Get Assignee() As String
    Assignee = assignee_
End Property

Public Property Let Assignee(ByVal email As String)
    assignee_ = IIf(email <> vbNullString, "-a " & email, vbNullString)
End Property
