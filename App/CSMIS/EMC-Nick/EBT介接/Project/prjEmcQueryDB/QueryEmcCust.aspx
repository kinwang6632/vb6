<%@ Page language="c#" Codebehind="QueryEmcCust.aspx.cs" AutoEventWireup="false" Inherits="prjEmcQueryDB.QueryEmcCust" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD>
		<title>
			<%=ConfigurationSettings.AppSettings["ModuleName"] %>
		</title>
		<META http-equiv="Content-Type" content="text/html; charset=big5">
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="C#" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body>
		<form id="Form1" method="post" runat="server">
			<DIV style="DISPLAY: inline; Z-INDEX: 100; LEFT: 50px; WIDTH: 700px; POSITION: absolute; TOP: 16px; HEIGHT: 20px"
				ms_positioning="FlowLayout"><font color="royalblue" size="4">
					<marquee width="700" height="10" bgcolor= navajowhite behavior="scroll" direction="left" scrolldelay='<%=ConfigurationSettings.AppSettings["ScrollDelay"] %>' scrollamount="10" onMouseOver="stop();" onMouseOut="start();">
						<p><%=ConfigurationSettings.AppSettings["marquee"] %></p>
					</marquee>
				</font>
			</DIV>
			<asp:label id="lblAddrField10" style="Z-INDEX: 137; LEFT: 712px; POSITION: absolute; TOP: 248px"
				runat="server" ForeColor="Maroon" Width="46px">(其他)</asp:label><asp:label id="lblAddrField6" style="Z-INDEX: 122; LEFT: 552px; POSITION: absolute; TOP: 248px"
				runat="server" ForeColor="Teal">衖</asp:label><asp:label id="lblAddrField5" style="Z-INDEX: 121; LEFT: 480px; POSITION: absolute; TOP: 248px"
				runat="server" ForeColor="Teal">弄</asp:label><asp:label id="lblAddrField1" style="Z-INDEX: 120; LEFT: 544px; POSITION: absolute; TOP: 216px"
				runat="server" ForeColor="Maroon" Width="192px">(村,里,莊,埔,路,街,其他)</asp:label><asp:label id="lblAddrField72" style="Z-INDEX: 119; LEFT: 648px; POSITION: absolute; TOP: 248px"
				runat="server" ForeColor="Teal">號</asp:label><asp:label id="lblAddrField3" style="Z-INDEX: 117; LEFT: 248px; POSITION: absolute; TOP: 248px"
				runat="server" ForeColor="Teal">段</asp:label><asp:textbox id="edtAddrField7" style="Z-INDEX: 115; LEFT: 576px; POSITION: absolute; TOP: 248px"
				tabIndex="18" runat="server" Width="72px"></asp:textbox><asp:textbox id="edtAddrField3" style="Z-INDEX: 111; LEFT: 184px; POSITION: absolute; TOP: 248px"
				tabIndex="14" runat="server" Width="60px"></asp:textbox>
			<DIV style="DISPLAY: inline; Z-INDEX: 106; LEFT: 24px; WIDTH: 88px; POSITION: absolute; TOP: 216px; HEIGHT: 24px"
				ms_positioning="FlowLayout"><font color="#cc0066" size="4">查詢條件:</font></DIV>
			<asp:label id="lblMessage" style="Z-INDEX: 105; LEFT: 24px; POSITION: absolute; TOP: 368px"
				runat="server" ForeColor="Red" Width="280px"></asp:label>
			<DIV style="DISPLAY: inline; Z-INDEX: 101; LEFT: 24px; WIDTH: 224px; POSITION: absolute; TOP: 40px; HEIGHT: 24px"
				ms_positioning="FlowLayout"><font color="#cc0066" size="4">系統台別 (資料更新日期)：</font></DIV>
			<HR style="Z-INDEX: 102; LEFT: 16px; WIDTH: 762px; POSITION: absolute; TOP: 392px; HEIGHT: 2px"
				width="762" color="#990000" SIZE="2">
			<asp:button id="btnQuery" style="Z-INDEX: 103; LEFT: 324px; POSITION: absolute; TOP: 360px"
				tabIndex="8" runat="server" Text="查詢" Width="56px"></asp:button><asp:checkboxlist id="chbCompList" style="Z-INDEX: 104; LEFT: 80px; POSITION: absolute; TOP: 80px"
				runat="server" Width="664px" RepeatColumns="5" Height="112px"></asp:checkboxlist><asp:radiobuttonlist id="rdgQueryCond" style="Z-INDEX: 107; LEFT: 80px; POSITION: absolute; TOP: 280px"
				runat="server" Width="676px" RepeatColumns="3" Height="72px" AutoPostBack="True">
				<asp:ListItem Value="1">客戶名稱</asp:ListItem>
				<asp:ListItem Value="2">聯絡電話</asp:ListItem>
				<asp:ListItem Value="3" Selected="True">裝機地址</asp:ListItem>
				<asp:ListItem Value="4">CATV客編</asp:ListItem>
				<asp:ListItem Value="5">CM/CT客編</asp:ListItem>
				<asp:ListItem Value="6">CM/CT合約編號</asp:ListItem>
			</asp:radiobuttonlist><asp:textbox id="edtKeyValue" style="Z-INDEX: 108; LEFT: 112px; POSITION: absolute; TOP: 216px"
				runat="server" Width="643px"></asp:textbox><asp:textbox id="edtAddrField1" style="Z-INDEX: 109; LEFT: 376px; POSITION: absolute; TOP: 216px"
				tabIndex="12" runat="server" Width="168px"></asp:textbox><asp:textbox id="edtAddrField2" style="Z-INDEX: 110; LEFT: 112px; POSITION: absolute; TOP: 248px"
				tabIndex="13" runat="server" Width="48px"></asp:textbox><asp:textbox id="edtAddrField4" style="Z-INDEX: 112; LEFT: 280px; POSITION: absolute; TOP: 248px"
				tabIndex="15" runat="server" Width="88px"></asp:textbox><asp:textbox id="edtAddrField5" style="Z-INDEX: 113; LEFT: 400px; POSITION: absolute; TOP: 248px"
				tabIndex="16" runat="server" Width="80px"></asp:textbox><asp:textbox id="edtAddrField6" style="Z-INDEX: 114; LEFT: 504px; POSITION: absolute; TOP: 248px"
				tabIndex="17" runat="server" Width="49px"></asp:textbox><asp:label id="lblAddrField2" style="Z-INDEX: 116; LEFT: 160px; POSITION: absolute; TOP: 248px"
				runat="server" ForeColor="Teal">鄰</asp:label><asp:label id="lblAddrField4" style="Z-INDEX: 123; LEFT: 368px; POSITION: absolute; TOP: 248px"
				runat="server" ForeColor="Teal">巷</asp:label><asp:textbox id="edtAddrField41" style="Z-INDEX: 124; LEFT: 16px; POSITION: absolute; TOP: 88px"
				runat="server" Width="32px" Visible="False"></asp:textbox><asp:textbox id="edtAddrField51" style="Z-INDEX: 125; LEFT: 16px; POSITION: absolute; TOP: 120px"
				runat="server" Width="32px" Visible="False"></asp:textbox><asp:textbox id="edtAddrField61" style="Z-INDEX: 126; LEFT: 16px; POSITION: absolute; TOP: 152px"
				runat="server" Width="32px" Visible="False"></asp:textbox><asp:label id="lblAddrHint" style="Z-INDEX: 127; LEFT: 528px; POSITION: absolute; TOP: 362px"
				runat="server" Width="160px">
				<a target="_blank" href="AddrHint.mht">顯示地址輸入提示畫面</a>
			</asp:label><asp:textbox id="edtAddrField8" style="Z-INDEX: 128; LEFT: 232px; POSITION: absolute; TOP: 216px"
				tabIndex="11" runat="server" Width="80px"></asp:textbox><asp:label id="lblAddrField8" style="Z-INDEX: 129; LEFT: 312px; POSITION: absolute; TOP: 216px"
				runat="server" ForeColor="Maroon">(鄉鎮市)</asp:label><asp:textbox id="edtAddrField9" style="Z-INDEX: 130; LEFT: 112px; POSITION: absolute; TOP: 216px"
				tabIndex="10" runat="server" Width="72px"></asp:textbox><asp:label id="lblAddrField9" style="Z-INDEX: 131; LEFT: 184px; POSITION: absolute; TOP: 216px"
				runat="server" ForeColor="Maroon">(縣市)</asp:label><asp:textbox id="edtAddrField21" style="Z-INDEX: 132; LEFT: 16px; POSITION: absolute; TOP: 184px"
				runat="server" Width="32px" Visible="False"></asp:textbox><asp:textbox id="edtAddrField31" style="Z-INDEX: 133; LEFT: 16px; POSITION: absolute; TOP: 248px"
				runat="server" Width="32px" Visible="False"></asp:textbox><asp:textbox id="edtAddrField71" style="Z-INDEX: 134; LEFT: 16px; POSITION: absolute; TOP: 280px"
				runat="server" Width="24px" Visible="False"></asp:textbox><asp:button id="btnReset" style="Z-INDEX: 135; LEFT: 424px; POSITION: absolute; TOP: 360px"
				runat="server" Text="重設" Width="56px"></asp:button>
			<asp:TextBox id="edtAddrField10" style="Z-INDEX: 136; LEFT: 672px; POSITION: absolute; TOP: 248px"
				runat="server" Width="40px"></asp:TextBox></form>
	</body>
</HTML>
