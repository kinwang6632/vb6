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
using System.Text;
using System.Configuration;
using System.Web.Security;

namespace prjEmcQueryDB
{
	/// <summary>
	/// WebForm1 的摘要描述。
	/// </summary>
	public class Login : System.Web.UI.Page
	{
		protected System.Web.UI.WebControls.Button btnLogin;
		protected System.Web.UI.WebControls.Label lblMessage;
		protected System.Web.UI.WebControls.Label lblModuleName;
		protected System.Web.UI.WebControls.Label Label1;
		protected System.Web.UI.WebControls.TextBox edtUserID;
		protected System.Web.UI.WebControls.TextBox edtUserPasswd;
		protected System.Web.UI.WebControls.Label Label2;
	
		private void Page_Load(object sender, System.EventArgs e)
		{
			// 在這裡放置使用者程式碼以初始化網頁
			lblModuleName.Text = "<p align=center>" + ConfigurationSettings.AppSettings["ModuleName"]  + "</p>";			
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
            this.btnLogin.Click += new System.EventHandler(this.btnLogin_Click);
            this.Load += new System.EventHandler(this.Page_Load);

        }
		#endregion
		public TextBox edtUerIDRef

		{
			get { return edtUserID; }
		}

		public TextBox edtUserPasswdRef

		{
			get { return edtUserPasswd; }
		}
		
		private void btnLogin_Click(object sender, System.EventArgs e)
		{
			if ( ( edtUserID.Text.Trim() != "" ) && ( edtUserPasswd.Text.Trim() != "" ) )
			  Server.Transfer( "CheckUser.aspx" );				
			else				
			  lblMessage.Text = String.Format( "請務必輸入使用者帳號與密碼! [ {0} ]", 		
  			    CommFun.GetNormalDateTimeStr() );
		}
		
	}
}
