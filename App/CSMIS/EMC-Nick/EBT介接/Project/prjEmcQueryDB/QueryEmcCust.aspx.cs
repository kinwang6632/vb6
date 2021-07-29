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
	/// QueryEmcCust 的摘要描述。
	/// </summary>
	public class QueryEmcCust : BasePage
	{
		protected System.Web.UI.WebControls.TextBox edtEmcCustId;
		protected System.Web.UI.WebControls.TextBox edtEbtCustId;
		protected System.Web.UI.WebControls.TextBox edtContractNo;
		protected System.Web.UI.WebControls.TextBox edtCustName;
		protected System.Web.UI.WebControls.TextBox edtCustAddr;
		protected System.Web.UI.WebControls.TextBox edtCustTel;
		protected System.Web.UI.WebControls.CheckBoxList chbCompList;
		protected System.Web.UI.WebControls.Label lblMessage;
		protected System.Web.UI.WebControls.TextBox edtKeyValue;
		protected System.Web.UI.WebControls.RadioButtonList rdgQueryCond;
		protected System.Web.UI.WebControls.TextBox edtAddrField1;
		protected System.Web.UI.WebControls.TextBox edtAddrField3;
		protected System.Web.UI.WebControls.TextBox edtAddrField2;
		protected System.Web.UI.WebControls.TextBox edtAddrField4;
		protected System.Web.UI.WebControls.TextBox edtAddrField5;
		protected System.Web.UI.WebControls.TextBox edtAddrField6;
		protected System.Web.UI.WebControls.Label lblAddrField2;
		protected System.Web.UI.WebControls.Label lblAddrField3;
		protected System.Web.UI.WebControls.Label lblAddrField1;
		protected System.Web.UI.WebControls.Label Label1;
		protected System.Web.UI.WebControls.Label Label2;
		protected System.Web.UI.WebControls.Label Label3;
		protected System.Web.UI.WebControls.Label lblAddrField5;
		protected System.Web.UI.WebControls.Label lblAddrField6;
		protected System.Web.UI.WebControls.Label lblAddrField4;
		protected System.Web.UI.WebControls.Label lblAddrField72;
		protected System.Web.UI.WebControls.TextBox edtAddrField41;
		protected System.Web.UI.WebControls.TextBox edtAddrField51;
		protected System.Web.UI.WebControls.TextBox edtAddrField61;
		protected System.Web.UI.WebControls.Label lblAddrHint;
		protected System.Web.UI.WebControls.TextBox edtAddrField8;
		protected System.Web.UI.WebControls.Label lblAddrField8;
		protected System.Web.UI.WebControls.TextBox edtAddrField9;
		protected System.Web.UI.WebControls.Label lblAddrField9;
		protected System.Web.UI.WebControls.TextBox edtAddrField21;
		protected System.Web.UI.WebControls.TextBox edtAddrField31;
		protected System.Web.UI.WebControls.TextBox edtAddrField7;
		protected System.Web.UI.WebControls.TextBox edtAddrField71;
		protected System.Web.UI.WebControls.Label Label4;
		protected System.Web.UI.WebControls.TextBox edtAddrField10;
		protected System.Web.UI.WebControls.Label lblAddrField10;
		protected System.Web.UI.WebControls.Button btnReset;
		protected System.Web.UI.WebControls.Button btnQuery;

		private void setDefaultCond()
		{
			//rdgQueryCond.SelectedIndex = 2; // 2004/06/07 取消此一功能
		}


		private void SetDefaultQueryValue()
		{
			CommFun L_CommFun = new CommFun();
			String sL_QK1="", sL_QK2="", sL_QK3="", sL_QK4="", sL_QK5="", sL_QK6="", sL_QK7="", sL_QK8="", sL_QK9="", sL_QK10="", sL_QK11="", sL_QK12="";
			String sL_QC1="0", sL_QC2="";
			int nL_QC2Length=0;

			sL_QK1 = L_CommFun.GetCookie(Request, "QK1");
			sL_QK2 = L_CommFun.GetCookie(Request, "QK2");
			sL_QK3 = L_CommFun.GetCookie(Request, "QK3");
			sL_QK4 = L_CommFun.GetCookie(Request, "QK4");
			sL_QK5 = L_CommFun.GetCookie(Request, "QK5");
			sL_QK6 = L_CommFun.GetCookie(Request, "QK6");
			sL_QK7 = L_CommFun.GetCookie(Request, "QK7");
			sL_QK8 = L_CommFun.GetCookie(Request, "QK8");
			sL_QK9 = L_CommFun.GetCookie(Request, "QK9");
			sL_QK10 = L_CommFun.GetCookie(Request, "QK10");
			sL_QK11 = L_CommFun.GetCookie(Request, "QK11");
			sL_QK12 = L_CommFun.GetCookie(Request, "QK12");
			sL_QC1 = L_CommFun.GetCookie(Request, "QC1");
			sL_QC2 = L_CommFun.GetCookie(Request, "QC2");

			edtKeyValue.Text = sL_QK1;
			edtAddrField1.Text = sL_QK2;
			edtAddrField2.Text = sL_QK3;
			edtAddrField3.Text = sL_QK4;
			edtAddrField4.Text = sL_QK5;
			edtAddrField5.Text = sL_QK6;
			edtAddrField6.Text = sL_QK7;
			edtAddrField7.Text = sL_QK8;
			edtAddrField8.Text = sL_QK10;
			edtAddrField9.Text = sL_QK11;
			edtAddrField10.Text = sL_QK12;
			
			if ((sL_QC1=="") || (sL_QC1==null))
				sL_QC1 = "0";
			rdgQueryCond.SelectedIndex = System.Convert.ToInt32(sL_QC1);

			if ((sL_QC2!="") && (sL_QC2!=null))
			{
				
				nL_QC2Length = sL_QC2.Length;
				for (int i=0; i<=nL_QC2Length-1; i++)
				{
					if (sL_QC2.Substring(i,1)=="1")
					  chbCompList.Items[i].Selected = true;
					else
					  chbCompList.Items[i].Selected = false;

				}
			}
		}

		private string GetSnapshotDateString(string aViewName)
		{
			oracleCommandForEbtUser.CommandText = string.Format( 
				" SELECT TO_CHAR( MAX(COMPUTEDATE),'YYYY/MM/DD') COMPUTEDATE " +
				"  FROM {0} " + 
				" WHERE TABLECODE = 'SO001' ", aViewName );
			string aResult = "";
			try
			{
				System.Data.OracleClient.OracleDataReader aDataReader = 
					oracleCommandForEbtUser.ExecuteReader();
						
				if ( aDataReader.HasRows )
				{
					aDataReader.Read();
					aResult = aDataReader["COMPUTEDATE"].ToString();
				}
				aDataReader.Close();
			}
			catch
			{
			   aResult = "";
			}
			return aResult;
		}

		private string getSoInfo(String sI_UserGroup)
		{
			String sL_Result ="", sL_QueryStr="";
			sL_QueryStr = "select SOCODE from EMC_EBT003  where CODENO=" + sI_UserGroup;

			System.Data.OracleClient.OracleDataReader L_DataReader;

			oracleCommandForEbtUser.CommandText = sL_QueryStr;
			try
			{
				L_DataReader = oracleCommandForEbtUser.ExecuteReader();
						
				if (L_DataReader.HasRows)
				{
					L_DataReader.Read();
					sL_Result = L_DataReader["SOCODE"].ToString();
				}
				else
					sL_Result="";
				L_DataReader.Close();
			}
			catch
			{
				sL_Result="";
			}


			return  sL_Result;
		}

		private void Page_Load(object sender, System.EventArgs e)
		{
			
			CommFun aCommFun = new CommFun();
			aCommFun.RedirectToLogin( this );
			
			if ( !IsPostBack )
			{
				
				int aIndex = 0;
				
				string sL_CompCode="", 
					    sL_ViewName="", 
					    sL_CompName="", 
					    sL_UserGroup="", 
					    sL_SoCode="",
				       sL_Tmp="";

				if ( !IsDbConnected() ) ConnToDB( string.Empty );					

				if ( chbCompList.Items.Count <= 0 )
				{
					char[] v_Sep= new char[1];
				
					v_Sep[0]= ConfigurationSettings.AppSettings["SoSep"].ToCharArray()[0];
					
					sL_UserGroup = Session["USERGROUP"].ToString();

					sL_SoCode = getSoInfo(sL_UserGroup);

					string[] v_SoCode = new string[20];
					v_SoCode = sL_SoCode.Split(v_Sep);

					chbCompList.Items.Clear();

					ListItem aItem = null;

					for ( aIndex = 0; aIndex <= v_SoCode.Length - 1; aIndex++ )
					{
						sL_CompCode = v_SoCode[aIndex];
						sL_CompName = CommFun.GetCompChineseName( byte.Parse( sL_CompCode ) );
						sL_ViewName = CommFun.GetQueryViewName( byte.Parse( sL_CompCode ), 2 );

						string aSnapshotDateString = string.Empty;

						if ( CommFun.IsEbtDataArea( sL_CompCode ) )
						{
							if ( sL_ViewName != string.Empty )
								aSnapshotDateString = GetSnapshotDateString( sL_ViewName );
						}


						if ( aSnapshotDateString != string.Empty )
							sL_Tmp = string.Format( "{0}<br>({1})", sL_CompName, aSnapshotDateString );
						else
							sL_Tmp = sL_CompName + "<br>";

						aItem = new ListItem( sL_Tmp, sL_CompCode );
						if ( !CommFun.IsEbtDataArea( sL_CompCode ) )
						{
							aItem.Text = string.Format( "<font color=Blue>{0}", aItem.Text );
						}
					
						chbCompList.Items.Add( aItem );
					}
				}

				//down, 取出 cookie...
				SetDefaultQueryValue();
				
				//up, 取出 cookie
				SetFieldVisibleStatus();

			
				setDefaultCond();
				SetFieldVisibleStatus();
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
			this.btnQuery.Click += new System.EventHandler(this.btnQuery_Click);
			this.rdgQueryCond.SelectedIndexChanged += new System.EventHandler(this.rdgQueryCond_SelectedIndexChanged);
			this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
			this.Load += new System.EventHandler(this.Page_Load);

		}
		#endregion

		private void SetFieldVisibleStatus()
		{
			bool bL_KeyFieldVisible = true;
			if (rdgQueryCond.SelectedItem.Text=="裝機地址")
				bL_KeyFieldVisible = false;


			edtKeyValue.Visible = bL_KeyFieldVisible;
			edtAddrField1.Visible = !bL_KeyFieldVisible;
			edtAddrField2.Visible = !bL_KeyFieldVisible;
			edtAddrField3.Visible = !bL_KeyFieldVisible;
			edtAddrField4.Visible = !bL_KeyFieldVisible;
			edtAddrField5.Visible = !bL_KeyFieldVisible;
			edtAddrField6.Visible = !bL_KeyFieldVisible;
			edtAddrField7.Visible = !bL_KeyFieldVisible;
			edtAddrField8.Visible = !bL_KeyFieldVisible;
			edtAddrField9.Visible = !bL_KeyFieldVisible;
			edtAddrField10.Visible = !bL_KeyFieldVisible;
			lblAddrField1.Visible = !bL_KeyFieldVisible;
			lblAddrField2.Visible = !bL_KeyFieldVisible;
			lblAddrField3.Visible = !bL_KeyFieldVisible;
			
			lblAddrField4.Visible = !bL_KeyFieldVisible;
			lblAddrField5.Visible = !bL_KeyFieldVisible;
			lblAddrField6.Visible = !bL_KeyFieldVisible;
			lblAddrField72.Visible = !bL_KeyFieldVisible;
			lblAddrField8.Visible = !bL_KeyFieldVisible;
			lblAddrField9.Visible = !bL_KeyFieldVisible;
			lblAddrField10.Visible = !bL_KeyFieldVisible;
			lblAddrHint.Visible = !bL_KeyFieldVisible;
			
		}

		public RadioButtonList rdgQueryCondRef
		{
			get 
			{
				return rdgQueryCond;
			}
		}

		public TextBox edtKeyValueRef
		{
			get 
			{
				return edtKeyValue;
			}
		}

		public CheckBoxList chbCompListRef

		{
			get 
			{
				return chbCompList;
			}
		}

		public TextBox edtCustTelRef

		{
			get 
			{
				return edtCustTel;
			}
		}

		public TextBox edtCustAddrRef

		{
			get 
			{
				return edtCustAddr;
			}
		}

		public TextBox edtCustNameRef

		{
			get 
			{
				return edtCustName;
			}
		}


		public TextBox edtContractNoRef

		{
			get 
			{
				return edtContractNo;
			}
		}

		public TextBox edtEbtCustIdRef

		{
			get 
			{
				return edtEbtCustId;
			}
		}

		public TextBox edtEmcCustIdRef

		{
			get 
			{
				return edtEmcCustId;
			}
		}

		public TextBox edtAddrField1Ref

		{
			get 
			{
				return edtAddrField1;
			}
		}
		
		public TextBox edtAddrField2Ref

		{
			get 
			{
				return edtAddrField2;
			}
		}

		public TextBox edtAddrField21Ref

		{
			get 
			{
				return edtAddrField21;
			}
		}
		public TextBox edtAddrField3Ref

		{
			get 
			{
				return edtAddrField3;
			}
		}

		public TextBox edtAddrField31Ref

		{
			get 
			{
				return edtAddrField31;
			}
		}

		public TextBox edtAddrField4Ref

		{
			get 
			{
				return edtAddrField4;
			}
		}

		public TextBox edtAddrField41Ref

		{
			get 
			{
				return edtAddrField41;
			}
		}

		public TextBox edtAddrField5Ref

		{
			get 
			{
				return edtAddrField5;
			}
		}

		public TextBox edtAddrField51Ref

		{
			get 
			{
				return edtAddrField51;
			}
		}

		public TextBox edtAddrField6Ref

		{
			get 
			{
				return edtAddrField6;
			}
		}

		public TextBox edtAddrField61Ref

		{
			get 
			{
				return edtAddrField61;
			}
		}

		public TextBox edtAddrField7Ref

		{
			get 
			{
				return edtAddrField7;
			}
		}

		public TextBox edtAddrField71Ref

		{
			get 
			{
				return edtAddrField71;
			}
		}
		

		public TextBox edtAddrField8Ref

		{
			get 
			{
				return edtAddrField8;
			}
		}

		public TextBox edtAddrField9Ref

		{
			get 
			{
				return edtAddrField9;
			}
		}

		public TextBox edtAddrField10Ref

		{
			get 
			{
				return edtAddrField10;
			}
		}

		private void btnQuery_Click(object sender, System.EventArgs e)
		{
			lblMessage.Text = "";			
			string sL_Msg = "";
			int nL_SelectedCount = 0, 
			    nL_Tmp;

			string sL_MaxSoCount = ConfigurationSettings.AppSettings["MaxSoQueryCount"];
			

			for (int aIndex = 0; aIndex < chbCompList.Items.Count; aIndex++ )
			{
				if ( chbCompList.Items[aIndex].Selected ) nL_SelectedCount++;					  
			}


			if ( nL_SelectedCount == 0 ) sL_Msg = "請選取您想要查詢的系統台";
			   

			if (rdgQueryCond.SelectedItem.Text=="裝機地址")
			{
				if ( edtAddrField7.Text.Trim() == "" )
					sL_Msg ="請輸入號";
				else
					edtAddrField71.Text = edtAddrField7.Text.Trim()	+ "號";
					     
				try
				{
				     if ( edtAddrField2.Text.Trim() != "" )
				       nL_Tmp = System.Convert.ToInt32( edtAddrField2.Text.Trim() );
				     edtAddrField21.Text = edtAddrField2.Text.Trim() + "鄰";
				 }
			   	catch 
			   	{
			   	   sL_Msg = "'鄰'必須要輸入數字"; 
			   	 }

				//down, 處理段之輸入...
				edtAddrField31.Text = "";
				if ( edtAddrField3.Text.Trim() != "" ) edtAddrField31.Text = edtAddrField3.Text.Trim() + "段";					    
					    
				//down, 處理巷之輸入...
				edtAddrField41.Text = "";
				if ( edtAddrField4.Text.Trim()!= "" ) edtAddrField41.Text = edtAddrField4.Text.Trim() + "巷";					    
					     

				//down, 處理弄之輸入...
				edtAddrField51.Text = "";
				if ( edtAddrField5.Text.Trim () != "" ) edtAddrField51.Text = edtAddrField5.Text.Trim() + "弄";
   					
				//down, 處理衖之輸入...
				edtAddrField61.Text = "";
				if (edtAddrField6.Text.Trim()!="") edtAddrField61.Text = edtAddrField6.Text.Trim() + "衖";		

			}
			else if ( rdgQueryCond.SelectedItem.Text == "聯絡電話" )
			{
  				if ( edtKeyValue.Text.Trim().Length < 7 ) sL_Msg ="聯絡電話最少要輸入7碼";		  			
			}
			else 
			{
            if ( edtKeyValue.Text.Trim() =="" ) sL_Msg = "請輸入您的查詢條件值";					
			}
		
			
			if ( nL_SelectedCount > System.Convert.ToInt32( sL_MaxSoCount ) )
			  	sL_Msg = "每次最多只能查詢 " + sL_MaxSoCount + " 家系統台";
			
			if ( sL_Msg == "" )			
				Server.Transfer( "QueryEmcCustR.aspx" );
			else
				lblMessage.Text = sL_Msg;
		}
		

		private void rdgQueryCond_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			SetFieldVisibleStatus();
		}

		private void btnReset_Click(object sender, System.EventArgs e)
		{
			edtAddrField1.Text = "";
			edtAddrField2.Text = "";
			edtAddrField3.Text = "";
			edtAddrField4.Text = "";
			edtAddrField5.Text = "";
			edtAddrField6.Text = "";
			edtAddrField7.Text = "";
			edtAddrField8.Text = "";
			edtAddrField9.Text = "";
			edtAddrField10.Text = "";
			edtKeyValue.Text = "";

			for ( int aIndex = 0; aIndex <= chbCompList.Items.Count - 1; aIndex++ )
				chbCompList.Items[aIndex].Selected = false;
				  
			SetFieldVisibleStatus();

		}
	}
}
