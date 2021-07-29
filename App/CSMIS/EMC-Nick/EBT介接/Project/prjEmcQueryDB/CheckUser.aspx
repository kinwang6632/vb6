<%@ Page language="c#" Codebehind="CheckUser.aspx.cs" AutoEventWireup="false" Inherits="prjEmcQueryDB.CheckUser" %>
<%@ Reference Page = "Login.aspx" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD>
		<title>
			<%=ConfigurationSettings.AppSettings["ModuleName"] %>
		</title>
		<meta name="vs_snapToGrid" content="True">
		<meta name="vs_showGrid" content="True">
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="C#" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body>
		<form id="Form1" method="post" runat="server">
			<asp:label id="lblMessage" style="Z-INDEX: 101; LEFT: 208px; POSITION: absolute; TOP: 32px"
				runat="server">
			  Label
			</asp:label>
		</form>
	</body>
</HTML>
