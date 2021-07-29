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
	/// OracleConnection 的摘要描述。
	/// </summary>
	public class OracleConnection : System.Web.UI.Page
	{
		protected System.Data.OracleClient.OracleCommand oracleCommandForEbtUser;
		protected System.Data.OracleClient.OracleCommand oracleCommandForEbtUser2;
		protected System.Data.OracleClient.OracleConnection oracleConnectionForEbtUser;
	
		public bool IsDbConnected()
		{
			return ( oracleConnectionForEbtUser.State == ConnectionState.Open ) ? true : false;
		}

		public void CloseDbConn()
		{
			if ( IsDbConnected() ) oracleConnectionForEbtUser.Close();
		}

		public void ConnToDB()
		{
			if ( !IsDbConnected() )
			{
				if ( oracleConnectionForEbtUser.ConnectionString == "" )
  				  oracleConnectionForEbtUser.ConnectionString = 
					  ConfigurationSettings.AppSettings["dbConnStr"];

				oracleConnectionForEbtUser.Open();
			}
		}

		private void Page_Load(object sender, System.EventArgs e)
		{
			// 在這裡放置使用者程式碼以初始化網頁
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
			this.oracleConnectionForEbtUser = new System.Data.OracleClient.OracleConnection();
			this.oracleCommandForEbtUser = new System.Data.OracleClient.OracleCommand();
			this.oracleCommandForEbtUser2 = new System.Data.OracleClient.OracleCommand();
			// 
			// oracleCommandForEbtUser
			// 
			this.oracleCommandForEbtUser.Connection = this.oracleConnectionForEbtUser;
			// 
			// oracleCommandForEbtUser2
			// 
			this.oracleCommandForEbtUser2.Connection = this.oracleConnectionForEbtUser;
			this.Load += new System.EventHandler(this.Page_Load);

		}
		#endregion
	}
}
