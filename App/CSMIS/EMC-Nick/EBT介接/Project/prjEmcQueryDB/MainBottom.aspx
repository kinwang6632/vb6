<%@ Page language="c#" Codebehind="MainBottom.aspx.cs" AutoEventWireup="false" Inherits="prjEmcQueryDB.MainBottom" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD>
		<title>
			<%=ConfigurationSettings.AppSettings["ModuleName"] %>
		</title>
		<meta name="GENERATOR" Content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" Content="C#">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<asp:Label id="lblModuleName" style="Z-INDEX: 101; LEFT: 304px; POSITION: absolute; TOP: 16px"
				runat="server" ForeColor="SlateBlue">
				Label</asp:Label>
		</form>
	</body>
</HTML>
