<%@ page language="C#" autoeventwireup="true" inherits="_Default, App_Web_xi5eptua" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
<script language="javascript" type="text/javascript">
// <!CDATA[

function window_onload() {

}

// ]]>
</script>
</head>
<body onload="return window_onload()">
    <form id="form1" runat="server" action="Default.aspx">
    <div>
        &nbsp;
        <asp:Panel ID="Panel1" runat="server" Height="36px" Width="460px" BorderStyle="Dotted" HorizontalAlign="Center">
            &nbsp; &nbsp;&nbsp;&nbsp;<br />
            <table style="width: 439px" border="1" >
                <tr>
                    <td align="center" colspan="2" style="height: 20px">
                        <font color="green" size="8">�q�T���t��</font></td>
                </tr>
                <tr>
                    <td align="right" colspan="" style="width: 128px">
                        �m�W</td>
                    <td style="width: 178px">
            <asp:TextBox ID="name" runat="server" BorderStyle="Double"></asp:TextBox></td>
                </tr>
                <tr>
                    <td align="right" style="width: 128px">
                        �q��</td>
                    <td style="width: 178px">
                        <asp:TextBox ID="telphone"
                runat="server" BorderStyle="Double"></asp:TextBox></td>
                </tr>
                <tr>
                    <td align="right" style="width: 128px">
                        QQ</td>
                    <td style="width: 178px">
                        <asp:TextBox ID="qq"
                runat="server" BorderStyle="Double"></asp:TextBox></td>
                </tr>
                <tr>
                    <td align="right" style="width: 128px">
                        MSN</td>
                    <td style="width: 178px">
                        <asp:TextBox ID="msn"
                runat="server" OnTextChanged="TextBox4_TextChanged" BorderStyle="Double"></asp:TextBox></td>
                </tr>
                <tr>
                    <td align="right" style="width: 128px">
                        ���</td>
                    <td style="width: 178px">
                        <asp:TextBox ID="cellphone"
                runat="server" BorderStyle="Double"></asp:TextBox></td>
                </tr>
                <tr>
                    <td align="right" style="width: 128px">
                        �u�@���</td>
                    <td style="width: 178px">
            <asp:TextBox ID="work" runat="server" BorderStyle="Double"></asp:TextBox></td>
                </tr>
                <tr>
                    <td align="right" style="width: 128px">
                        �a�}</td>
                    <td style="width: 178px">
            <asp:TextBox ID="address" runat="server" BorderStyle="Double"></asp:TextBox></td>
                </tr>
                <tr>
                    <td align="right" style="width: 128px">
                        Email</td>
                    <td style="width: 178px">
            <asp:TextBox ID="email" runat="server" AutoCompleteType="Email" BorderStyle="Double"></asp:TextBox></td>
                </tr>
                <tr>
                    <td align="right" style="width: 128px">
                        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click"
                Text="�O�s" BorderStyle="Double" />
            <asp:Button ID="Button2" runat="server" Text="���m" /></td>
                    <td style="width: 178px">
            <asp:Button ID="Button3" runat="server" OnClick="Button3_Click" Text="�d�ݳq�T��" />
                        <asp:Button ID="Button4" runat="server" OnClick="Button4_Click" Text="�d��" /></td>
                </tr>
            </table>
            &nbsp;&nbsp;
        </asp:Panel>
    
    </div>
    </form>
</body>
</html>
