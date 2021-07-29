<%@ Page CodeBehind="MainTop.aspx.cs" Language="c#" AutoEventWireup="false" Inherits="prjEmcQueryDB.MainTop" %>
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
	<body>
		<form id="frmTop" method="post" runat="server">
			<DIV style="DISPLAY: inline; Z-INDEX: 101; LEFT: 128px; WIDTH: 152px; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
				ms_positioning="FlowLayout" align="left"><font color="#993333" size="4">使用者：<% =GetCookie("UserName") %></font></DIV>
			<FONT face="新細明體"></FONT>
			<asp:Label id="lblFun1" style="Z-INDEX: 104; LEFT: 320px; POSITION: absolute; TOP: 8px" runat="server">
				<a href="QueryEmcCust.aspx" target="frmBottom">客戶狀態資料查詢</a></asp:Label>
			<IMG style="Z-INDEX: 102; LEFT: 8px; WIDTH: 72px; POSITION: absolute; TOP: 8px; HEIGHT: 24px"
				height="24" src="Images/logo.gif" width="72">
			<asp:Label id="lblFun2" style="Z-INDEX: 103; LEFT: 480px; POSITION: absolute; TOP: 8px" runat="server">
				<a href="MaintainUserData.aspx" target="frmBottom">使用者資料維護</a></asp:Label>
			<asp:Label id="lblFun4" style="Z-INDEX: 105; LEFT: 640px; POSITION: absolute; TOP: 8px" runat="server">
				<a href="MaintainGroupData.aspx" target="frmBottom">群組資料維護</a></asp:Label>
		</form>
	</body>
</HTML>
