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
	  /// MainTop 的摘要描述。
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
					// 在這裡放置使用者程式碼以初始化網頁
				CommFun aCommFun = new CommFun();
				aCommFun.RedirectToLogin( this );

				if ( !IsPostBack )
				{
					string sL_SupervisorFlag = aCommFun.GetCookie( Request, "ISSUPERVISOR" );
					string sL_SysOpFlag = aCommFun.GetCookie( Request, "ISSYSOP" );

					lblFun1.Visible = true;
					lblFun2.Visible = true;
					lblFun4.Visible = false;

					// 表示是系統管理員
					if (sL_SysOpFlag == "1" ) lblFun4.Visible = true;
						
				}
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
