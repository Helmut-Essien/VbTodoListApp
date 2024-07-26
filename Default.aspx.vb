Imports System
Imports System.Collections.Generic
Imports System.IO
Imports OfficeOpenXml

Public Class _Default
    Inherits System.Web.UI.Page

    Private Property Tasks As List(Of Task)
        Get
            If Session("Tasks") Is Nothing Then
                Session("Tasks") = New List(Of Task)()
            End If
            Return CType(Session("Tasks"), List(Of Task))
        End Get
        Set(value As List(Of Task))
            Session("Tasks") = value
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        If Not IsPostBack Then
            BindGrid()
        End If
    End Sub

    Protected Sub btnAddTask_Click(sender As Object, e As EventArgs)
        Dim newTask As New Task() With {
            .TaskId = If(Tasks.Count > 0, Tasks.Max(Function(t) t.TaskId) + 1, 1),
            .TaskName = txtTask.Text
        }
        Tasks.Add(newTask)
        BindGrid()
        txtTask.Text = String.Empty
    End Sub

    Private Sub BindGrid()
        gvTasks.DataSource = Tasks
        gvTasks.DataBind()
    End Sub

    Protected Sub gvTasks_RowEditing(sender As Object, e As GridViewEditEventArgs)
        gvTasks.EditIndex = e.NewEditIndex
        BindGrid()
    End Sub

    Protected Sub gvTasks_RowCancelingEdit(sender As Object, e As GridViewCancelEditEventArgs)
        gvTasks.EditIndex = -1
        BindGrid()
    End Sub

    Protected Sub gvTasks_RowUpdating(sender As Object, e As GridViewUpdateEventArgs)
        Dim taskId As Integer = Convert.ToInt32(gvTasks.DataKeys(e.RowIndex).Value)
        Dim task = Tasks.FirstOrDefault(Function(t) t.TaskId = taskId)
        If task IsNot Nothing Then
            task.TaskName = CType(gvTasks.Rows(e.RowIndex).FindControl("txtEditTask"), TextBox).Text
        End If
        gvTasks.EditIndex = -1
        BindGrid()
    End Sub

    Protected Sub gvTasks_RowDeleting(sender As Object, e As GridViewDeleteEventArgs)
        Dim taskId As Integer = Convert.ToInt32(gvTasks.DataKeys(e.RowIndex).Value)
        Dim task = Tasks.FirstOrDefault(Function(t) t.TaskId = taskId)
        If task IsNot Nothing Then
            Tasks.Remove(task)
        End If
        BindGrid()
    End Sub

    ' Button click event to save data as an Excel workbook
    Protected Sub btnSaveExcel_Click(sender As Object, e As EventArgs)
        ' Set the EPPlus license context
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        ' Create a new Excel package
        Using package As New ExcelPackage()
            ' Add a worksheet for tasks
            Dim tasksSheet = package.Workbook.Worksheets.Add("Tasks")
            tasksSheet.Cells(1, 1).Value = "Task ID"
            tasksSheet.Cells(1, 2).Value = "Task Name"

            ' Add tasks data to the worksheet
            For i As Integer = 0 To Tasks.Count - 1
                tasksSheet.Cells(i + 2, 1).Value = Tasks(i).TaskId
                tasksSheet.Cells(i + 2, 2).Value = Tasks(i).TaskName
            Next

            ' Additional sheets can be added similarly
            Dim anotherSheet = package.Workbook.Worksheets.Add("AnotherSheet")

            ' Save the Excel package to a memory stream
            Dim stream As New MemoryStream()
            package.SaveAs(stream)

            ' Write the memory stream to the response
            Response.Clear()
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            Response.AddHeader("content-disposition", "attachment; filename=TodoList.xlsx")
            Response.BinaryWrite(stream.ToArray())
            Response.End()
        End Using
    End Sub
End Class
