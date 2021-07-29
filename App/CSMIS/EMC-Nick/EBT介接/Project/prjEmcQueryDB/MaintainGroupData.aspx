<%@ Page language="c#" Codebehind="MaintainGroupData.aspx.cs" AutoEventWireup="false" Inherits="prjEmcQueryDB.MaintainGroupData" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >
<HTML>
	<HEAD>
		<title>客戶狀態資料查詢系統</title> <!-- #include virtual="./include/common_function.inc" -->
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="C#" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<asp:checkboxlist id="chbCompList" style="Z-INDEX: 110; LEFT: 16px; POSITION: absolute; TOP: 232px"
				runat="server" Width="208px" Height="112px" RepeatColumns="2">
				<asp:ListItem Value="9">陽明山</asp:ListItem>
				<asp:ListItem Value="10">新台北</asp:ListItem>
				<asp:ListItem Value="12">大安文山</asp:ListItem>
				<asp:ListItem Value="11">金頻道</asp:ListItem>
				<asp:ListItem Value="8">全聯</asp:ListItem>
				<asp:ListItem Value="13">新唐城</asp:ListItem>
				<asp:ListItem Value="14">大新店</asp:ListItem>
				<asp:ListItem Value="16">北桃園</asp:ListItem>
				<asp:ListItem Value="7">振道</asp:ListItem>
				<asp:ListItem Value="5">新頻道</asp:ListItem>
				<asp:ListItem Value="6">豐盟</asp:ListItem>
				<asp:ListItem Value="3">南天</asp:ListItem>
				<asp:ListItem Value="1">觀昇</asp:ListItem>
				<asp:ListItem Value="2">屏南</asp:ListItem>
			</asp:checkboxlist>
			<br>
			<table cellSpacing="0" cellPadding="0" width="100%">
				<tr>
					<td class="CONTENTTITLE" align="center" width="100%" bgColor="#d3c9c7" colSpan="2">群組資料維護
					</td>
				</tr>
				<tr>
					<td><span class="MESSAGE" id="Message" runat="server"></span></td>
				</tr>
				<tr>
					<td style="FONT: 10pt verdana" width="30%" bgColor="#aaaadd">新增</td>
					<td style="FONT: 10pt verdana" width="70%" bgColor="#ffccdd">修改/刪除</td>
				</tr>
			</table>
			<BR>
			&nbsp;
			<DIV style="DISPLAY: inline; FONT-WEIGHT: normal; FONT-SIZE: 10pt; Z-INDEX: 102; LEFT: 16px; WIDTH: 96px; LINE-HEIGHT: normal; FONT-STYLE: normal; POSITION: absolute; TOP: 80px; FONT-VARIANT: normal"
				ms_positioning="FlowLayout">群組代號</DIV>
			<asp:textbox id="edtCodeNo" style="Z-INDEX: 105; LEFT: 88px; POSITION: absolute; TOP: 80px" runat="server"
				Height="24px" Font-Size="10pt" Width="128px"></asp:textbox><BR>
			&nbsp;&nbsp;<BR>
			&nbsp;
			<DIV style="DISPLAY: inline; FONT-SIZE: 10pt; Z-INDEX: 103; LEFT: 16px; WIDTH: 70px; POSITION: absolute; TOP: 120px; HEIGHT: 15px"
				ms_positioning="FlowLayout">名稱</DIV>
			<BR>
			&nbsp; <input id="Submit1" style="Z-INDEX: 101; LEFT: 120px; WIDTH: 96px; POSITION: absolute; TOP: 184px; HEIGHT: 24px"
				type="submit" value="新增" name="Submit1" runat="server" OnServerClick="Add_Click">&nbsp;
			<BR>
			</FONT></TD>&nbsp;
			<asp:textbox id="edtDesc" style="Z-INDEX: 106; LEFT: 88px; POSITION: absolute; TOP: 120px" runat="server"
				Font-Size="10pt" Width="128px"></asp:textbox><br>
			<DIV style="DISPLAY: inline; FONT-SIZE: 10pt; Z-INDEX: 104; LEFT: 16px; WIDTH: 64px; POSITION: absolute; TOP: 160px; HEIGHT: 24px"
				ms_positioning="FlowLayout">是否停用</DIV>
			<br>
			<asp:radiobuttonlist id="rdbStopFlagForInsert" style="Z-INDEX: 107; LEFT: 88px; POSITION: absolute; TOP: 152px"
				runat="server" Font-Size="10pt" RepeatColumns="2">
				<asp:ListItem Value="1">是</asp:ListItem>
				<asp:ListItem Value="0" Selected="True">否</asp:ListItem>
			</asp:radiobuttonlist>
			<asp:Label id="Label2" style="Z-INDEX: 111; LEFT: 24px; POSITION: absolute; TOP: 216px" runat="server"
				ForeColor="#0000C0">所屬系統台</asp:Label><br>
			<br>
			<asp:label id="lblInsertMsg" style="Z-INDEX: 108; LEFT: 16px; POSITION: absolute; TOP: 8px"
				runat="server" Font-Bold="True" ForeColor="#C00000"></asp:label><br>
			<br>
			<br>
			<ASP:DATAGRID id="MyDataGrid" style="Z-INDEX: 109; LEFT: 240px; POSITION: absolute; TOP: 80px"
				runat="server" Font-Size="Smaller" AutoGenerateColumns="False" DataKeyField="CodeNo" OnDeleteCommand="MyDataGrid_Delete"
				OnUpdateCommand="MyDataGrid_Update" OnCancelCommand="MyDataGrid_Cancel" OnEditCommand="MyDataGrid_Edit"
				HeaderStyle-BackColor="lightblue" Font-Name="Verdana" CellPadding="3" BorderColor="Black"
				BackColor="#F4FFF4" Font-Names="Verdana">
				<HeaderStyle BackColor="LightBlue"></HeaderStyle>
				<Columns>
					<asp:EditCommandColumn ButtonType="PushButton" UpdateText="儲存" CancelText="取消" EditText="修改">
						<ItemStyle Wrap="False"></ItemStyle>
					</asp:EditCommandColumn>
					<asp:ButtonColumn Text="刪除" ButtonType="PushButton" CommandName="Delete"></asp:ButtonColumn>
					<asp:BoundColumn DataField="CODENO" SortExpression="CODENO" ReadOnly="True" HeaderText="群組代號">
						<ItemStyle Wrap="False"></ItemStyle>
					</asp:BoundColumn>
					<asp:TemplateColumn SortExpression="DESCRIPTION" HeaderText="名稱">
						<ItemTemplate>
							<asp:Label id=Label3 runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "DESCRIPTION") %>'>
							</asp:Label>
						</ItemTemplate>
						<EditItemTemplate>
							<asp:TextBox id=edtDescForUpdate runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "DESCRIPTION") %>'>
							</asp:TextBox>
							<asp:RequiredFieldValidator id="Requiredfieldvalidator1" runat="server" Font-Size="12" Font-Name="Verdana" ErrorMessage="*:不可空白"
								Display="Dynamic" ControlToValidate="edtDescForUpdate"></asp:RequiredFieldValidator>
						</EditItemTemplate>
					</asp:TemplateColumn>
					<asp:TemplateColumn SortExpression="soname" HeaderText="所屬系統台">
						<ItemTemplate>
							<asp:Label id="Label1" runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "SONAME") %>'>
							</asp:Label>
						</ItemTemplate>
						<EditItemTemplate>
							<asp:Label id="Label11" runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "SONAME") %>'>
							</asp:Label>
							<asp:RequiredFieldValidator id="Requiredfieldvalidator2" runat="server" Font-Size="12" Font-Name="Verdana" ErrorMessage="*:不可空白"
								Display="Dynamic" ControlToValidate="edtDescForUpdate"></asp:RequiredFieldValidator>
						</EditItemTemplate>
					</asp:TemplateColumn>
					<asp:TemplateColumn SortExpression="STOPFLAG" HeaderText="是否停用">
						<ItemTemplate>
							<asp:Label id=Label5 runat="server" Text='<%# getIsSysOpLableCaption(DataBinder.Eval(Container.DataItem, "STOPFLAG").ToString()) %>'>
							</asp:Label>
						</ItemTemplate>
						<EditItemTemplate>
							<NOBR>
								<asp:RadioButtonList id=rdbStopFlag runat="server" RepeatColumns="2" SelectedIndex='<%# getIsSysOpIndex(DataBinder.Eval(Container.DataItem,"STOPFLAG").ToString()) %>'>
									<asp:ListItem Value="1">是</asp:ListItem>
									<asp:ListItem Value="0">否</asp:ListItem>
								</asp:RadioButtonList></NOBR>
						</EditItemTemplate>
					</asp:TemplateColumn>
				</Columns>
			</ASP:DATAGRID></form>
	</body>
</HTML>
