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
	/// MainBottom 的摘要描述。
	/// </summary>
	public class MainBottom : System.Web.UI.Page
	{
		protected System.Web.UI.WebControls.Label lblModuleName;
	
		private void Page_Load(object sender, System.EventArgs e)
		{
			// 在這裡放置使用者程式碼以初始化網頁
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

		#region Web Form 設計工具產生的程式碼
		override protected void OnInit(EventArgs e)
		{
			//
			// CODEGEN: 此為 ASP.NET Web Form 設計工具所需的呼叫。
			//
			InitializeComponent();
			base.OnInit(e);
		}
		
		/// <summary>
		/// 此為設計工具支援所必須的方法 - 請勿使用程式碼編輯器修改
		/// 這個方法的內容。
		/// </summary>
		private void InitializeComponent()
		{    
			this.Load += new System.EventHandler(this.Page_Load);

		}
		#endregion
	}
}
