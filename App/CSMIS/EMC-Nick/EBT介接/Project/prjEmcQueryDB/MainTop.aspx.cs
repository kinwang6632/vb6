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
	  /// MainTop ���K�n�y�z�C
	  /// </summary>
	  /// 
  	
		public class MainTop : System.Web.UI.Page
	  {
			protected System.Web.UI.WebControls.Label lblFun1;
			protected System.Web.UI.WebControls.Label lblFun4;
			protected System.Web.UI.WebControls.Label lblFun2;
		
			protected string GetCookie(string Key)
			{
				CommFun L_CommFun = new CommFun();
				return L_CommFun.GetCookie( Request, Key );
			}		
		    
		   private void Page_Load(object sender, System.EventArgs e)
		   {
					// �b�o�̩�m�ϥΪ̵{���X�H��l�ƺ���
				CommFun aCommFun = new CommFun();
				aCommFun.RedirectToLogin( this );

				if ( !IsPostBack )
				{
					string sL_SupervisorFlag = aCommFun.GetCookie( Request, "ISSUPERVISOR" );
					string sL_SysOpFlag = aCommFun.GetCookie( Request, "ISSYSOP" );

					lblFun1.Visible = true;
					lblFun2.Visible = true;
					lblFun4.Visible = false;

					// ��ܬO�t�κ޲z��
					if (sL_SysOpFlag == "1" ) lblFun4.Visible = true;
						
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
