<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Default.aspx.vb" Inherits="VbTodoListApp._Default" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>To-Do List</title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <h2>To-Do List</h2>
            <asp:TextBox ID="txtTask" runat="server" Width="300px" />
            <asp:Button ID="btnAddTask" runat="server" Text="Add Task" OnClick="btnAddTask_Click" />
            <asp:GridView ID="gvTasks" runat="server" AutoGenerateColumns="False" DataKeyNames="TaskId"
                          OnRowEditing="gvTasks_RowEditing" OnRowDeleting="gvTasks_RowDeleting"
                          OnRowCancelingEdit="gvTasks_RowCancelingEdit" OnRowUpdating="gvTasks_RowUpdating">
                <Columns>
                    <asp:BoundField DataField="TaskId" HeaderText="Task ID" ReadOnly="True" />
                    <asp:BoundField DataField="TaskName" HeaderText="Task Name" />
                    <asp:TemplateField HeaderText="Actions">
                        <ItemTemplate>
                            <asp:LinkButton ID="lnkEdit" runat="server" CommandName="Edit" Text="Edit" />
                            <asp:LinkButton ID="lnkDelete" runat="server" CommandName="Delete" Text="Delete" />
                        </ItemTemplate>
                        <EditItemTemplate>
                            <asp:TextBox ID="txtEditTask" runat="server" Text='<%# Bind("TaskName") %>' />
                            <asp:LinkButton ID="lnkUpdate" runat="server" CommandName="Update" Text="Update" />
                            <asp:LinkButton ID="lnkCancel" runat="server" CommandName="Cancel" Text="Cancel" />
                        </EditItemTemplate>
                    </asp:TemplateField>
                </Columns>
            </asp:GridView>
            <!-- Button to save data as an Excel workbook -->
            <asp:Button ID="btnSaveExcel" runat="server" Text="Save as Excel" OnClick="btnSaveExcel_Click" />
        </div>
    </form>
</body>
</html>
