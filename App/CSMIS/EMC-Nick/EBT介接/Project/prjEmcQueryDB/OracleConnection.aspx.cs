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
	/// OracleConnection ���K�n�y�z�C
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
			// �b�o�̩�m�ϥΪ̵{���X�H��l�ƺ���
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
