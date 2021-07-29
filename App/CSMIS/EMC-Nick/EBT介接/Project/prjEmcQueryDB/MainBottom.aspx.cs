using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Web;
using System.Web.SessionState;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using System.Configuration;

namespace prjEmcQueryDB
{
	/// <summary>
	/// MainBottom ���K�n�y�z�C
	/// </summary>
	public class MainBottom : System.Web.UI.Page
	{
		protected System.Web.UI.WebControls.Label lblModuleName;
	
		private void Page_Load(object sender, System.EventArgs e)
		{
			// �b�o�̩�m�ϥΪ̵{���X�H��l�ƺ���
			lblModuleName.Text = ConfigurationSettings.AppSettings["ModuleName"]  ;
			CommFun L_CommFun = new CommFun();
			String sL_IsSupervisor = L_CommFun.GetCookie(Request, "ISSUPERVISOR");
			String sL_IsSysOp = L_CommFun.GetCookie(Request, "ISSYSOP");
			String sL_IfChangePasswd = Session["IFCHANGEPASSWD"].ToString();

			if (sL_IfChangePasswd=="N")
			    Server.Transfer("MaintainUserData.aspx");
			else if ((sL_IsSupervisor=="0") && (sL_IsSysOp=="0"))
				Server.Transfer("QueryEmcCust.aspx");
			  

		}

		#region Web Form �]�p�u�㲣�ͪ��{���X
		override protected void OnInit(EventArgs e)
		{
			//
			// CODEGEN: ���� ASP.NET Web Form �]�p�u��һݪ��I�s�C
			//
			InitializeComponent();
			base.OnInit(e);
		}
		
		/// <summary>
		/// �����]�p�u��䴩�ҥ�������k - �ФŨϥε{���X�s�边�ק�
		/// �o�Ӥ�k�����e�C
		/// </summary>
		private void InitializeComponent()
		{    
			this.Load += new System.EventHandler(this.Page_Load);

		}
		#endregion
	}
}
