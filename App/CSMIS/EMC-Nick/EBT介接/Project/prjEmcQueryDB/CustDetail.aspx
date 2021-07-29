<%@ Page language="c#" Codebehind="CustDetail.aspx.cs" AutoEventWireup="false" Inherits="prjEmcQueryDB.CustDetail" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD>
		<title>
			<%=ConfigurationSettings.AppSettings["ModuleName"] %>
		</title>
		<meta name="vs_showGrid" content="True">
		<meta name="GENERATOR" Content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" Content="C#">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body>
		<form id="Form1" method="post" runat="server" style="LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 100%">
			<FONT face="新細明體">
				<asp:DataGrid id="dbgEmc" style="Z-INDEX: 101; LEFT: 16px; POSITION: absolute; TOP: 32px" runat="server"
					BackColor="White" BorderColor="#999999" BorderStyle="None" BorderWidth="1px" CellPadding="3"
					GridLines="Vertical" Width="99%" Font-Size="Smaller">
					<SelectedItemStyle Font-Bold="True" ForeColor="White" BackColor="#008A8C"></SelectedItemStyle>
					<AlternatingItemStyle BackColor="Gainsboro"></AlternatingItemStyle>
					<ItemStyle ForeColor="Black" BackColor="#EEEEEE"></ItemStyle>
					<HeaderStyle Font-Bold="True" ForeColor="White" BackColor="#000084"></HeaderStyle>
					<FooterStyle ForeColor="Black" BackColor="#CCCCCC"></FooterStyle>
					<PagerStyle HorizontalAlign="Center" ForeColor="Black" BackColor="#999999" Mode="NumericPages"></PagerStyle>
				</asp:DataGrid>
				<asp:DataGrid id="dbgEbt" style="Z-INDEX: 104; LEFT: 16px; POSITION: absolute; TOP: 184px" runat="server"
					BackColor="White" BorderColor="#999999" BorderStyle="None" BorderWidth="1px" CellPadding="3"
					GridLines="Vertical" Width="99%" Font-Size="Smaller">
					<SelectedItemStyle Font-Bold="True" ForeColor="White" BackColor="#008A8C"></SelectedItemStyle>
					<AlternatingItemStyle BackColor="Gainsboro"></AlternatingItemStyle>
					<ItemStyle ForeColor="Black" BackColor="#EEEEEE"></ItemStyle>
					<HeaderStyle Font-Bold="True" ForeColor="White" BackColor="#000084"></HeaderStyle>
					<FooterStyle ForeColor="Black" BackColor="#CCCCCC"></FooterStyle>
					<PagerStyle HorizontalAlign="Center" ForeColor="Black" BackColor="#999999" Mode="NumericPages"></PagerStyle>
				</asp:DataGrid>
				<asp:Label id="lblEbtMessage" style="Z-INDEX: 103; LEFT: 16px; POSITION: absolute; TOP: 160px"
					runat="server" ForeColor="DarkGreen" Font-Bold="True">CM/CT 相關資料:</asp:Label>
				<asp:Label id="lblQueryResult" style="Z-INDEX: 102; LEFT: 16px; POSITION: absolute; TOP: 8px"
					runat="server" ForeColor="DarkGreen" Font-Bold="True">CATV 相關資料:</asp:Label>
			</FONT>
		</form>
	</body>
</HTML>
