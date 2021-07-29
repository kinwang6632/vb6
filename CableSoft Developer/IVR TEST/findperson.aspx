<%@ page language="C#" autoeventwireup="true" inherits="Default2, App_Web_q9iwivb0" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table style="width: 508px; height: 58px">
            <tr>
                <td align="center" style="width: 151px; height: 6px;">
                </td>
                <td align="center" style="width: 250px; height: 6px;">
                    通訊錄查找演示</td>
                <td align="center" style="width: 111px; height: 6px">
                </td>
            </tr>
            <tr>
                <td align="center" style="width: 151px; height: 1px">
                    請輸入用戶姓名：</td>
                <td align="center" style="width: 250px; height: 1px">
                    <asp:TextBox ID="name" runat="server"></asp:TextBox></td>
                <td align="center" style="height: 1px; width: 111px;">
                    <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="查找" /></td>
            </tr>
        </table>
    
    </div>
    </form>
</body>
</html>
