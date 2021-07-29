<%@ Page language="c#" Codebehind="Login.aspx.cs" AutoEventWireup="false" Inherits="prjEmcQueryDB.Login" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD>
		<title>
			<%=ConfigurationSettings.AppSettings["ModuleName"] %>
		</title>
		<script language="vbscript">
	           sub setFocus
	               document.frmLogin.edtuserID.focus()	               
	           end sub	
		</script>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="C#" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body onload="setFocus()">
		<form id="frmLogin" method="post" runat="server">
			<FONT face="新細明體"></FONT><FONT face="新細明體"></FONT><FONT face="新細明體"></FONT><FONT face="新細明體">
			</FONT><FONT face="新細明體"></FONT><FONT face="新細明體"></FONT><FONT face="新細明體"></FONT>
			<FONT face="新細明體"></FONT><FONT face="新細明體"></FONT><FONT face="新細明體"></FONT><FONT face="新細明體">
			</FONT><FONT face="新細明體"></FONT><FONT face="新細明體"></FONT><FONT face="新細明體"></FONT>
			<FONT face="新細明體"></FONT><FONT face="新細明體"></FONT><FONT face="新細明體"></FONT><FONT face="新細明體">
			</FONT><FONT face="新細明體"></FONT><FONT face="新細明體"></FONT><FONT face="新細明體"></FONT>
			<FONT face="新細明體"></FONT>
			<br>
			<br>
			<br>
			<br>
			<table align="center" width="90%">
				<tr align="center">
					<td align="center" colspan="3">
						<asp:Label id="lblModuleName" runat="server" ForeColor="SteelBlue" Font-Size="X-Large" Font-Bold="True"></asp:Label>
					</td>
				</tr>
				<tr>
					<td></td>
				</tr>
				<tr>
					<td></td>
				</tr>
				<tr align="center">
					<td>
						<asp:Label id="Label1" runat="server">使用者帳號</asp:Label>
						<asp:TextBox id="edtUserID" runat="server"></asp:TextBox>
						<asp:Button id="btnLogin" accessKey="U" tabIndex="2" runat="server" Text="登入"></asp:Button>
					</td>
				</tr>
				<tr align="center">
					<td>
						<asp:Label id="Label2" runat="server">使用者密碼</asp:Label>
						<asp:TextBox id="edtUserPasswd" runat="server" TextMode="Password"></asp:TextBox><INPUT tabIndex="1" type="reset" value="重設">
					</td>
				</tr>
			</table>
			<p align="center"><FONT size="3">
					<asp:Label id="lblMessage" runat="server" ForeColor="Red"></asp:Label></FONT></p>
			<HR width="539" color="#990000" SIZE="2" align="center">
			<BR>
			<font color="royalblue" size="4">
				<MARQUEE width="90%" style="LEFT: 5%; POSITION: relative" BGCOLOR="navajowhite" DIRECTION="left"
					BEHAVIOR="scroll" SCROLLAMOUNT="10" SCROLLDELAY='<%= ConfigurationSettings.AppSettings["ScrollDelay"]%>'
					onMouseOver="JavaScript:stop();" onMouseOut="JavaScript:start();" >
					<P><%= ConfigurationSettings.AppSettings["marquee"] %></P>
				</MARQUEE>
			</font>
		</form>
	</body>
</HTML>
