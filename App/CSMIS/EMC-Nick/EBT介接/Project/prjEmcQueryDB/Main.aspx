<%@ Page CodeBehind="Main.aspx.cs" Language="c#" AutoEventWireup="false" Inherits="prjEmcQueryDB.Main" %>
<HTML>
	<HEAD>
		<TITLE>
			<%=ConfigurationSettings.AppSettings["ModuleName"] %>
		</TITLE>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="C#" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<frameset rows="12%,88%">
		<frame name="frmTop" src="MainTop.aspx" height="10" scrolling="no" noresize>
		<frame src="MainBottom.aspx" name="frmBottom" noresize>
	</frameset>
</HTML>
