<%@ Reference Page = "QueryEmcCust.aspx" %>
<%@ Page language="c#" Codebehind="QueryEmcCustR.aspx.cs" AutoEventWireup="false" Inherits="prjEmcQueryDB.QueryEmcCustR" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD>
		<title>
			<%=ConfigurationSettings.AppSettings["ModuleName"] %>
		</title>
		<meta name="vs_showGrid" content="True">
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="C#" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<style type="text/css">
		  .Link { CURSOR: hand; TEXT-DECORATION: underline }
		</style>
		<script language="javascript">
			function CustIdClick(aId,aCode)
			{
				var aWidth = screen.availwidth - 10 + "px";
				var aHeight = screen.availheight - 50 + "px";
				var aFeature = 
					"top=0px, left=0px, height="+ aHeight + "," + " width=" + aWidth + "," +
					"toolbar=no, location=no, directories=no, status=yes, menubar=no, scrollbars=no, resizable=yes";				
				window.open( "CustDetail.aspx?CUSTID="+aId+"&COMPCODE="+aCode, null, aFeature );
					
			}				
		</script>
	</HEAD>
	<body>
		<form id="Form1" method="post" runat="server">
			<asp:datagrid id="dbgComp1" style="Z-INDEX: 101; LEFT: 24px; POSITION: absolute; TOP: 120px" runat="server"
				AllowCustomPaging="True" BackColor="White" BorderColor="#999999" BorderStyle="None" BorderWidth="1px"
				CellPadding="3" GridLines="Vertical" Width="100%" Font-Size="Smaller">
				<SelectedItemStyle Font-Bold="True" ForeColor="White" BackColor="#008A8C"></SelectedItemStyle>
				<AlternatingItemStyle BackColor="Gainsboro"></AlternatingItemStyle>
				<ItemStyle ForeColor="Black" BackColor="#EEEEEE"></ItemStyle>
				<HeaderStyle Font-Bold="True" ForeColor="White" BackColor="#000084"></HeaderStyle>
				<FooterStyle ForeColor="Black" BackColor="White"></FooterStyle>
				<PagerStyle NextPageText="下一頁" PrevPageText="上一頁" HorizontalAlign="Center" ForeColor="Black"
					Position="Top" BackColor="#999999"></PagerStyle>
			</asp:datagrid><asp:label id="lblQueryResult" style="Z-INDEX: 103; LEFT: 16px; POSITION: absolute; TOP: 64px"
				runat="server" ForeColor="DarkGreen" Font-Bold="True"></asp:label><asp:label id="lblQueryCond" style="Z-INDEX: 102; LEFT: 96px; POSITION: absolute; TOP: 16px"
				runat="server" ForeColor="MediumBlue" Font-Bold="True"></asp:label>
			<DIV style="DISPLAY: inline; Z-INDEX: 104; LEFT: 16px; WIDTH: 70px; POSITION: absolute; TOP: 16px; HEIGHT: 15px"
				ms_positioning="FlowLayout"><A href="QueryEmcCust.aspx">重新查詢</A></DIV>
			<asp:button id="btnPreviousPage" style="Z-INDEX: 105; LEFT: 104px; POSITION: absolute; TOP: 88px"
				runat="server" Text="上一頁" CommandName="上一頁"></asp:button><asp:button id="btnNextPage" style="Z-INDEX: 106; LEFT: 176px; POSITION: absolute; TOP: 88px"
				runat="server" Text="下一頁" CommandName="下一頁"></asp:button><asp:label id="lblPageNum" style="Z-INDEX: 107; LEFT: 344px; POSITION: absolute; TOP: 88px"
				runat="server">頁次:</asp:label>
			<asp:Button id="btnFirstPage" style="Z-INDEX: 108; LEFT: 32px; POSITION: absolute; TOP: 88px"
				runat="server" Text="第一頁" CommandName="第一頁"></asp:Button>
			<asp:Button id="btnLastPage" style="Z-INDEX: 109; LEFT: 248px; POSITION: absolute; TOP: 88px"
				runat="server" Text="最後一頁" CommandName="最後一頁"></asp:Button></form>
		</TR></TBODY></TABLE>
	</body>
</HTML>
