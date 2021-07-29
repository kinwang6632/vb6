<%@ Page language="c#" Codebehind="AddrHint.aspx.cs" AutoEventWireup="false" Inherits="prjEmcQueryDB.AddrHint" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD>
		<title>
			<%=ConfigurationSettings.AppSettings["ModuleName"] %>
		</title>
		<SCRIPT language="JScript">		
	    function closeWindow()     
	    {
	 
	    
	      window.close();
    }

		</SCRIPT>
		<meta name="GENERATOR" Content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" Content="C#">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:label id="lblNote1" style="Z-INDEX: 105; LEFT: 80px; POSITION: absolute; TOP: 56px" runat="server"
					Font-Size="Small" ForeColor="#804000" Width="256px">[1. 號必須要輸入，不能空白]</asp:label>
				<asp:label id="lblNote2" style="Z-INDEX: 101; LEFT: 80px; POSITION: absolute; TOP: 88px" runat="server"
					Font-Size="Small" ForeColor="#804000" Width="304px">[2. 鄰只能輸入數字，不能輸入文字]</asp:label>
				<asp:label id="lblNote5" style="Z-INDEX: 103; LEFT: 80px; POSITION: absolute; TOP: 120px" runat="server"
					Font-Size="Small" ForeColor="#804000" Width="536px"> [3.鄰段巷弄衖號等單位字，都不需要輸入；而其餘的單位字則請您自行輸入</asp:label>
				<DIV style="DISPLAY: inline; Z-INDEX: 107; LEFT: 64px; WIDTH: 144px; POSITION: absolute; TOP: 24px; HEIGHT: 24px"
					ms_positioning="FlowLayout"><font color="blue">地址條件輸入說明 :</font></DIV>
				<INPUT style="Z-INDEX: 108; LEFT: 248px; POSITION: absolute; TOP: 24px" type="button" value="關閉視窗"
					onclick="closeWindow()"> </FONT>
		</form>
	</body>
</HTML>
