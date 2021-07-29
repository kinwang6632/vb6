<%@ Page language="c#" Codebehind="MaintainUserData.aspx.cs" AutoEventWireup="false" Inherits="prjEmcQueryDB.MaintainUserData" %>
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
			<asp:radiobuttonlist id="rdbStopFlagForInsert" style="Z-INDEX: 114; LEFT: 120px; POSITION: absolute; TOP: 264px"
				runat="server" Font-Size="10pt" RepeatColumns="2">
				<asp:ListItem Value="1">是</asp:ListItem>
				<asp:ListItem Value="0" Selected="True">否</asp:ListItem>
			</asp:radiobuttonlist><asp:radiobuttonlist id="rdbIsSysOpForInsert" style="Z-INDEX: 112; LEFT: 120px; POSITION: absolute; TOP: 232px"
				runat="server" Font-Size="10pt" RepeatColumns="2">
				<asp:ListItem Value="1">是</asp:ListItem>
				<asp:ListItem Value="0" Selected="True">否</asp:ListItem>
			</asp:radiobuttonlist><br>
			<table cellSpacing="0" cellPadding="0" width="100%">
				<tr>
					<td class="CONTENTTITLE" align="center" width="100%" bgColor="#d3c9c7" colSpan="2">使用者資料維護
					</td>
				</tr>
				<tr>
					<td style="WIDTH: 273px"><span class="MESSAGE" id="Message" runat="server"></span></td>
				</tr>
				<tr>
					<td style="FONT: 10pt verdana; WIDTH: 273px" width="273" bgColor="#aaaadd">新增</td>
					<td style="FONT: 10pt verdana" width="65%" bgColor="#ffccdd">修改/刪除</td>
				</tr>
			</table>
			<BR>
			&nbsp;
			<DIV style="DISPLAY: inline; FONT-WEIGHT: normal; FONT-SIZE: 10pt; Z-INDEX: 101; LEFT: 16px; WIDTH: 56px; LINE-HEIGHT: normal; FONT-STYLE: normal; POSITION: absolute; TOP: 80px; HEIGHT: 16px; FONT-VARIANT: normal"
				ms_positioning="FlowLayout">帳號</DIV>
			<asp:textbox id="edtUserID" style="Z-INDEX: 105; LEFT: 48px; POSITION: absolute; TOP: 80px" runat="server"
				Font-Size="10pt" Width="128px" Height="24px"></asp:textbox><BR>
			&nbsp;
			<DIV style="DISPLAY: inline; FONT-SIZE: 10pt; Z-INDEX: 102; LEFT: 16px; WIDTH: 56px; POSITION: absolute; TOP: 112px; HEIGHT: 16px"
				ms_positioning="FlowLayout">密碼</DIV>
			<asp:textbox id="edtUserPasswd" style="Z-INDEX: 106; LEFT: 48px; POSITION: absolute; TOP: 112px"
				runat="server" Font-Size="10pt" Width="128px" Height="22px" TextMode="Password"></asp:textbox><asp:label id="lblIsSysOpForInsert" style="Z-INDEX: 115; LEFT: 24px; POSITION: absolute; TOP: 240px"
				runat="server" Font-Size="Smaller">是否系統管理者</asp:label><BR>
			&nbsp;
			<DIV style="DISPLAY: inline; FONT-SIZE: 10pt; Z-INDEX: 104; LEFT: 16px; WIDTH: 48px; POSITION: absolute; TOP: 144px; HEIGHT: 16px"
				ms_positioning="FlowLayout">名稱</DIV>
			<BR>
			&nbsp; <input id="btnAdd" style="Z-INDEX: 113; LEFT: 144px; WIDTH: 96px; POSITION: absolute; TOP: 296px; HEIGHT: 24px"
				type="submit" value="新增" name="Submit1" runat="server" OnServerClick="Add_Click">&nbsp;
			<BR>
			</FONT></TD>&nbsp;
			<asp:textbox id="edtUserName" style="Z-INDEX: 107; LEFT: 48px; POSITION: absolute; TOP: 144px"
				runat="server" Font-Size="10pt" Width="128px"></asp:textbox><br>
			<asp:dropdownlist id="lstUserGroupForInsert" style="Z-INDEX: 110; LEFT: 48px; POSITION: absolute; TOP: 176px"
				runat="server" Font-Size="10pt" Width="216px"></asp:dropdownlist><asp:label id="lblStopFlagForInsert" style="Z-INDEX: 117; LEFT: 64px; POSITION: absolute; TOP: 272px"
				runat="server" Font-Size="Smaller">是否停用</asp:label><br>
			<asp:radiobuttonlist id="rdbIsSupervisorForInsert" style="Z-INDEX: 108; LEFT: 120px; POSITION: absolute; TOP: 200px"
				runat="server" Font-Size="10pt" RepeatColumns="2">
				<asp:ListItem Value="1">是</asp:ListItem>
				<asp:ListItem Value="0" Selected="True">否</asp:ListItem>
			</asp:radiobuttonlist><asp:label id="lblIsSupervisorForInsert" style="Z-INDEX: 116; LEFT: 24px; POSITION: absolute; TOP: 208px"
				runat="server" Font-Size="Smaller">是否群組管理者</asp:label><asp:label id="lblUserGroupForInsert" style="Z-INDEX: 118; LEFT: 16px; POSITION: absolute; TOP: 176px"
				runat="server" Font-Size="Smaller">群組</asp:label><br>
			<br>
			<asp:label id="lblInsertMsg" style="Z-INDEX: 109; LEFT: 16px; POSITION: absolute; TOP: 8px"
				runat="server" ForeColor="#C00000" Font-Bold="True"></asp:label><br>
			<br>
			<br>
			<ASP:DATAGRID id="MyDataGrid" style="Z-INDEX: 111; LEFT: 308px; POSITION: absolute; TOP: 104px"
				runat="server" Font-Size="Smaller" Font-Names="Verdana" BackColor="#F4FFF4" BorderColor="Black"
				CellPadding="3" Font-Name="Verdana" HeaderStyle-BackColor="lightblue" OnEditCommand="MyDataGrid_Edit"
				OnCancelCommand="MyDataGrid_Cancel" OnUpdateCommand="MyDataGrid_Update" OnDeleteCommand="MyDataGrid_Delete"
				DataKeyField="UserID" AutoGenerateColumns="False">
				<HeaderStyle BackColor="LightBlue"></HeaderStyle>
				<Columns>
					<asp:EditCommandColumn ButtonType="PushButton" UpdateText="儲存" CancelText="取消" EditText="修改">
						<ItemStyle Wrap="False"></ItemStyle>
					</asp:EditCommandColumn>
					<asp:ButtonColumn Text="刪除" ButtonType="PushButton" CommandName="Delete"></asp:ButtonColumn>
					<asp:BoundColumn DataField="USERID" SortExpression="USERID" ReadOnly="True" HeaderText="帳號">
						<ItemStyle Wrap="False"></ItemStyle>
					</asp:BoundColumn>
					<asp:TemplateColumn SortExpression="USERPASSWD" HeaderText="密碼">
						<ItemTemplate>
							<asp:Label id=Label2 runat="server" Text='<%# transPassword(DataBinder.Eval(Container.DataItem, "USERPASSWD").ToString()) %>'>
							</asp:Label>
						</ItemTemplate>
						<EditItemTemplate>
							<asp:TextBox id=edtUserPasswdForUpdate runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "UserPasswd") %>'>
							</asp:TextBox>
							<asp:RequiredFieldValidator id="UserPasswd_ReqVal" runat="server" Font-Size="12" Font-Name="Verdana" ControlToValidate="edtUserPasswdForUpdate"
								Display="Dynamic"></asp:RequiredFieldValidator>
						</EditItemTemplate>
					</asp:TemplateColumn>
					<asp:TemplateColumn SortExpression="USERNAME" HeaderText="名稱">
						<ItemTemplate>
							<asp:Label id=Label3 runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "UserName") %>'>
							</asp:Label>
						</ItemTemplate>
						<EditItemTemplate>
							<asp:TextBox id=edtUserNameForUpdate runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "UserName") %>' ReadOnly="True">
							</asp:TextBox>
							<asp:RequiredFieldValidator id="Requiredfieldvalidator1" runat="server" Font-Size="12" Font-Name="Verdana" ErrorMessage="*:不可空白"
								Display="Dynamic" ControlToValidate="edtUserNameForUpdate"></asp:RequiredFieldValidator>
						</EditItemTemplate>
					</asp:TemplateColumn>
					<asp:TemplateColumn SortExpression="USERGROUP" HeaderText="群組">
						<ItemTemplate>
							<asp:Label id=Label6 runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "UserGroup") %>'>
							</asp:Label>
						</ItemTemplate>
						<EditItemTemplate>
							<NOBR>
								<asp:Label id="Label4" runat="server" Text='<%# DataBinder.Eval(Container.DataItem, "UserGroup") %>'>
								</asp:Label></NOBR>
						</EditItemTemplate>
					</asp:TemplateColumn>
					<asp:TemplateColumn SortExpression="ISSUPERVISOR" HeaderText="是否群組管理者">
						<ItemTemplate>
							<asp:Label id=Label5 runat="server" Text='<%# getIsSupervisorLableCaption(DataBinder.Eval(Container.DataItem, "ISSUPERVISOR").ToString()) %>'>
							</asp:Label>
						</ItemTemplate>
						<EditItemTemplate>
							<NOBR>
								<asp:RadioButtonList id=rdbIsSupervisor runat="server" RepeatColumns="2" SelectedIndex='<%# getIsSupervisorIndex(DataBinder.Eval(Container.DataItem,"ISSUPERVISOR").ToString()) %>'>
									<asp:ListItem Value="1">是</asp:ListItem>
									<asp:ListItem Value="0">否</asp:ListItem>
								</asp:RadioButtonList></NOBR>
						</EditItemTemplate>
					</asp:TemplateColumn>
					<asp:TemplateColumn SortExpression="ISSUPERVISOR" HeaderText="是否系統管理者">
						<ItemTemplate>
							<asp:Label id=Label1 runat="server" Text='<%# getIsSysOpLableCaption(DataBinder.Eval(Container.DataItem, "ISSYSOP").ToString()) %>'>
							</asp:Label>
						</ItemTemplate>
						<EditItemTemplate>
							<NOBR>
								<asp:RadioButtonList id=rdbIsSysOp runat="server" RepeatColumns="2" SelectedIndex='<%# getIsSysOpIndex(DataBinder.Eval(Container.DataItem,"ISSYSOP").ToString()) %>'>
									<asp:ListItem Value="1">是</asp:ListItem>
									<asp:ListItem Value="0">否</asp:ListItem>
								</asp:RadioButtonList></NOBR>
						</EditItemTemplate>
					</asp:TemplateColumn>
					<asp:TemplateColumn SortExpression="STOPFLAG" HeaderText="是否停用">
						<ItemTemplate>
							<asp:Label id=Label7 runat="server" Text='<%# getStopFlagCaption(DataBinder.Eval(Container.DataItem, "STOPFLAG").ToString()) %>'>
							</asp:Label>
						</ItemTemplate>
						<EditItemTemplate>
							<NOBR>
								<asp:RadioButtonList id=rdbStopFlag runat="server" RepeatColumns="2" SelectedIndex='<%# getStopFlagIndex(DataBinder.Eval(Container.DataItem,"STOPFLAG").ToString()) %>'>
									<asp:ListItem Value="1">是</asp:ListItem>
									<asp:ListItem Value="0">否</asp:ListItem>
								</asp:RadioButtonList></NOBR>
						</EditItemTemplate>
					</asp:TemplateColumn>
				</Columns>
			</ASP:DATAGRID></form>
	</body>
</HTML>
