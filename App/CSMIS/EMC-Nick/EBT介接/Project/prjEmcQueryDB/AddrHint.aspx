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
			<FONT face="�s�ө���">
				<asp:label id="lblNote1" style="Z-INDEX: 105; LEFT: 80px; POSITION: absolute; TOP: 56px" runat="server"
					Font-Size="Small" ForeColor="#804000" Width="256px">[1. �������n��J�A����ť�]</asp:label>
				<asp:label id="lblNote2" style="Z-INDEX: 101; LEFT: 80px; POSITION: absolute; TOP: 88px" runat="server"
					Font-Size="Small" ForeColor="#804000" Width="304px">[2. �F�u���J�Ʀr�A�����J��r]</asp:label>
				<asp:label id="lblNote5" style="Z-INDEX: 103; LEFT: 80px; POSITION: absolute; TOP: 120px" runat="server"
					Font-Size="Small" ForeColor="#804000" Width="536px"> [3.�F�q�ѧ����������r�A�����ݭn��J�F�Ө�l�����r�h�бz�ۦ��J</asp:label>
				<DIV style="DISPLAY: inline; Z-INDEX: 107; LEFT: 64px; WIDTH: 144px; POSITION: absolute; TOP: 24px; HEIGHT: 24px"
					ms_positioning="FlowLayout"><font color="blue">�a�}�����J���� :</font></DIV>
				<INPUT style="Z-INDEX: 108; LEFT: 248px; POSITION: absolute; TOP: 24px" type="button" value="��������"
					onclick="closeWindow()"> </FONT>
		</form>
	</body>
</HTML>
