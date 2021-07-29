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
	/// WebForm1 ���K�n�y�z�C
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
			// �b�o�̩�m�ϥΪ̵{���X�H��l�ƺ���
			lblModuleName.Text = "<p align=center>" + ConfigurationSettings.AppSettings["ModuleName"]  + "</p>";			
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
			  lblMessage.Text = String.Format( "�аȥ���J�ϥΪ̱b���P�K�X! [ {0} ]", 		
  			    CommFun.GetNormalDateTimeStr() );
		}
		
	}
}
