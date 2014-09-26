<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="Core.OneDriveUploadSelectorWeb.Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>OneDrive Upload Selector</title>
    <script type="text/javascript" src="../Scripts/jquery-1.9.1.js"></script>
    <script type="text/javascript" src="../Scripts/app.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnableCdn="True" />
    <div id="divSPChrome"></div>
    <div style="left: 40px; position: absolute;">
        <h1>Add/update the ribbon and custom actions</h1>
        <p>View the button implemented on the <asp:HyperLink runat="server" ID="DocumentsLink" Text="Documents" Font-Bold="true" Target="_blank" /> doclib.</p>
        <asp:Button runat="server" ID="InitializeButton" OnClick="InitializeButton_Click" Text="Add/update Ribbon" />
        
        <h1>Remove the ribbon and custom actions</h1>
        <asp:Button runat="server" ID="RemoveButton" OnClick="RemoveButton_Click" Text="Remove Ribbon" />
    </div>
    </form>
</body>
</html>
