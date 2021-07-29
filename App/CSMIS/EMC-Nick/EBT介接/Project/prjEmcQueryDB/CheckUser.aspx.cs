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
	/// CheckUser ���K�n�y�z�C
	/// </summary>
	public class CheckUser : BasePage
	{

		protected System.Web.UI.WebControls.Label lblMessage;
		
		private void WriteLoginLog(string aUserID, string aIp, string aStatus, string aNote)
		{
			if ( ConfigurationSettings.AppSettings["IfWriteLogOfUserLogin"] == "Y" )
			{
				oracleCommandForEbtUser2.CommandText = string.Format( 
					" INSERT INTO EMC_EBT004 ( USERID, FROMIP, LOGINTIME, STATUS, NOTE ) " + 
               " VALUES ( '{0}', '{1}', {2}, '{3}', '{4}' ) ", aUserID, aIp, CommFun.getOracleDateTimeStr(),
					aStatus, aNote );
				try
				{
					oracleCommandForEbtUser2.ExecuteNonQuery();
				}
				catch
				{
					// nothing to do ...
				}
			}
		}

		private void Page_Load(object sender, System.EventArgs e)
		{
			// �b�o�̩�m�ϥΪ̵{���X�H��l�ƺ���
			if ( !IsPostBack )
			{
				prjEmcQueryDB.Login SourcePage = (Login)Context.Handler;

				ConnToDB( string.Empty );

				try
				{
					Session.Clear();

					string aUserID = SourcePage.edtUerIDRef.Text;
					string aPasswd = SourcePage.edtUserPasswdRef.Text;
					string aIp = Request.UserHostAddress.ToString();

					if ( IsDbConnected() )
					{
						string aStatus = ConfigurationSettings.AppSettings["LoginStatusForSuccess"];
						string aNote = "";
					
						oracleCommandForEbtUser.CommandText = string.Format(
							" SELECT * FROM EMC_EBT002 " + 
							"  WHERE STOPFLAG = 0      " + 
							"    AND UPPER( USERID )=  UPPER( '{0}' ) " +
							"    AND UPPER( USERPASSWD ) = UPPER( '{1}' )  ", aUserID, aPasswd );

						System.Data.OracleClient.OracleDataReader 
							aReader = oracleCommandForEbtUser.ExecuteReader();

						if ( aReader.HasRows )
						{
							aReader.Read();
						
							HttpCookie L_MyCookie = new HttpCookie( "USERINFO" );
							L_MyCookie.Values.Add( "USERID", aUserID );
							L_MyCookie.Values.Add( "USERNAME", aReader["USERNAME"].ToString() );
							L_MyCookie.Values.Add( "USERGROUP", aReader["USERGROUP"].ToString() );
							L_MyCookie.Values.Add( "ISSUPERVISOR", aReader["ISSUPERVISOR"].ToString() );
							L_MyCookie.Values.Add( "ISSYSOP", aReader["ISSYSOP"].ToString() );
						
							Response.AppendCookie( L_MyCookie );
							
							Session.Add( "USERID", aUserID );
							Session.Add( "USERNAME", aReader["USERNAME"].ToString() );							
							Session.Add( "USERGROUP", aReader["USERGROUP"].ToString() );
							Session.Add( "ISSUPERVISOR", aReader["ISSUPERVISOR"].ToString() );
							Session.Add( "ISSYSOP", aReader["ISSYSOP"].ToString() );
							Session.Add( "IFCHANGEPASSWD", aReader["IFCHANGEPASSWD"].ToString() );

						}
						else
						{
							aStatus = ConfigurationSettings.AppSettings["LoginStatusForFail"];
							aNote = string.Format( "�K�X:{0}", aPasswd );						
						}

						aReader.Close();
						WriteLoginLog( aUserID, aIp, aStatus, aNote );

					}
				}
				finally
				{
            	CloseDbConn();	
				}

				if ( Session.Count > 0 )
					Server.Transfer( "Main.aspx" );
				else
					lblMessage.Text = 
						"<p align=center><a href='Login.aspx'>�z��J���ϥΪ̱b��/�K�X�����T. �Э��s��J.</a></p>";
			
			}
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
