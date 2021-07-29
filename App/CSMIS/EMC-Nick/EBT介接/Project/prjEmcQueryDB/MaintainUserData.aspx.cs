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
using System.Data.OracleClient;


namespace prjEmcQueryDB
{
	/// <summary>
	/// MaintainUserData1 的摘要描述。
	/// </summary>
	public class MaintainUserData : BasePage
	{
		protected System.Web.UI.WebControls.DataGrid MyDataGrid;
		protected System.Web.UI.WebControls.TextBox edtUserID;
		protected System.Web.UI.WebControls.TextBox edtUserPasswd;
		protected System.Web.UI.WebControls.TextBox edtUserName;
		protected System.Web.UI.WebControls.RadioButtonList rdbIsSupervisorForInsert;
		protected System.Web.UI.WebControls.Label lblInsertMsg;
		protected System.Web.UI.WebControls.DropDownList lstUserGroupForInsert;
		protected System.Web.UI.HtmlControls.HtmlGenericControl Message;
		protected System.Web.UI.WebControls.RadioButtonList rdbIsSysOpForInsert;
		protected System.Web.UI.HtmlControls.HtmlInputButton btnAdd;
		protected System.Web.UI.WebControls.RadioButtonList rdbStopFlagForInsert;
		protected System.Web.UI.WebControls.Label lblIsSysOpForInsert;
		protected System.Web.UI.WebControls.Label lblIsSupervisorForInsert;
		protected System.Web.UI.WebControls.Label lblStopFlagForInsert;
		protected System.Web.UI.WebControls.Label lblUserGroupForInsert;
		
		protected System.Data.DataSet dsUserGroup;

		

		private string getGroupDesc(String sI_CodeNo)
		{
			String sL_Result="";
			sL_Result = "abc";
			return sL_Result;

		}

		public int getStopFlagIndex(string sI_Value)
		{
			int nL_Result=0;

			if (sI_Value=="1")
				nL_Result = 0;
			else if (sI_Value=="0")
				nL_Result = 1;
			return nL_Result;
		}

		public String getStopFlagCaption(string aValue)
		{
			string aResult = "";
			if ( aValue == "1" )
				aResult = "是";
			else if ( aValue == "0" )
				aResult = "否";
			return aResult;
		}

		private void Page_Load(object sender, System.EventArgs e)
		{
			// 在這裡放置使用者程式碼以初始化網頁
			
			System.Data.OracleClient.OracleDataReader L_DataReaderForUserGroup;
			CommFun aCommFun = new CommFun();
			aCommFun.RedirectToLogin( this );

			ConnToDB( string.Empty );//不可以放在  if (!IsPostBack) 裡面, 否則會有錯誤..

			if ( !IsPostBack )
			{
				oracleCommandForEbtUser.CommandText = 
               "SELECT CODENO, DESCRIPTION FROM EMC_EBT003 WHERE STOPFLAG = 0";
				
				L_DataReaderForUserGroup = oracleCommandForEbtUser.ExecuteReader();

				lstUserGroupForInsert.DataSource = L_DataReaderForUserGroup;				
				lstUserGroupForInsert.DataTextField = "DESCRIPTION";
				lstUserGroupForInsert.DataValueField = "CODENO";
				lstUserGroupForInsert.DataBind();		
				lstUserGroupForInsert.SelectedIndex = 0;
				L_DataReaderForUserGroup.Close();
				BindGrid();
			}

			CloseDbConn();
			
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
			this.dsUserGroup = new System.Data.DataSet();
			((System.ComponentModel.ISupportInitialize)(this.dsUserGroup)).BeginInit();
			this.btnAdd.ServerClick += new System.EventHandler(this.Submit1_ServerClick);
			// 
			// dsUserGroup
			// 
			this.dsUserGroup.DataSetName = "NewDataSet";
			this.dsUserGroup.Locale = new System.Globalization.CultureInfo("zh-TW");
			this.Load += new System.EventHandler(this.Page_Load);
			((System.ComponentModel.ISupportInitialize)(this.dsUserGroup)).EndInit();

		}
		#endregion

		public String transPassword(String sI_Password)
		{
			
			String sL_Result ="";
			/*
			for (int i=0; i<sI_Password.Length; i++)
			  sL_Result = sL_Result + "*";
			*/

			sL_Result = "***";
			return sL_Result;
		}

		public String getIsSupervisorLableCaption(string sI_Value)
		{
			String sL_Result="";

			if (sI_Value=="1")
				sL_Result = "是";
			else if (sI_Value=="0")
				sL_Result = "否";
			return sL_Result;
		}


		public int getIsSupervisorIndex(string sI_Value)
		{
			int nL_Result=0;

			if (sI_Value=="1")
				nL_Result = 0;
			else if (sI_Value=="0")
				nL_Result = 1;
			return nL_Result;
		}


		public String getIsSysOpLableCaption(string sI_Value)
		{
			String sL_Result="";

			if (sI_Value=="1")
				sL_Result = "是";
			else if (sI_Value=="0")
				sL_Result = "否";
			return sL_Result;
		}


		public int getIsSysOpIndex(string sI_Value)
		{
			int nL_Result=0;

			if (sI_Value=="1")
				nL_Result = 0;
			else if (sI_Value=="0")
				nL_Result = 1;
			return nL_Result;
		}


		public void Add_Click(Object sender, EventArgs E) 
		{
			try
			{
				ConnToDB( string.Empty );
				string sL_NewUserID = edtUserID.Text.Trim();
				string sL_NewUserName = edtUserName.Text.Trim();
				string sL_NewUserPasswd = edtUserPasswd.Text.Trim();
				if( (sL_NewUserID!="") && (sL_NewUserName!="") && (sL_NewUserPasswd!="") )
				{
					string sL_NewUserGroup = lstUserGroupForInsert.SelectedValue;
					string sL_NewUserIsSupervisor = rdbIsSupervisorForInsert.SelectedValue;
					string sL_NewUserIsSysOp = rdbIsSysOpForInsert.SelectedValue;
					string sL_NewStopFlag = rdbStopFlagForInsert.SelectedValue;
					string sL_OracleDateTimeStr = CommFun.getOracleDateTimeStr();

					CommFun L_CommFun = new CommFun();
					string sL_UpdEn = L_CommFun.GetCookie( Request, "USERNAME" );
					string sL_InsertCmd = "";

					oracleCommandForEbtUser.CommandText = string.Format( 
						" SELECT USERGROUP FROM EMC_EBT002 " +						
						"  WHERE USERID = '{0}' ", sL_NewUserID );
					
					System.Data.OracleClient.OracleDataReader L_DataReader;

					L_DataReader = oracleCommandForEbtUser.ExecuteReader();
				
					
					if (!L_DataReader.HasRows)
					{

						sL_InsertCmd = "insert into EMC_EBT002 (USERID,USERNAME,USERPASSWD,USERGROUP,ISSUPERVISOR, ISSYSOP,STOPFLAG, UPDTIME,UPDEN, IFCHANGEPASSWD) values ('" + sL_NewUserID + "','" + sL_NewUserName + "','" + sL_NewUserPasswd + "'," + sL_NewUserGroup + "," + sL_NewUserIsSupervisor + "," + sL_NewUserIsSysOp +  "," + sL_NewStopFlag + "," + sL_OracleDateTimeStr + ",'" + sL_UpdEn +"','N'" +  ")";				
						oracleCommandForEbtUser.CommandText = sL_InsertCmd;
						try 
						{
						
							oracleCommandForEbtUser.ExecuteNonQuery();
					
						}
						catch
						{
							//OracleErrorHandler(e);
						}
				
						BindGrid();
					}
					else
					{
						L_DataReader.Read();			
						lblInsertMsg.Text = "使用者帳號(" + sL_NewUserID + ")已經存在! 所屬群組:" +  L_DataReader["USERGROUP"].ToString();
					}
					L_DataReader.Close();
				}
				else
				{
				  lblInsertMsg.Text = "使用者帳號,密碼,名稱都不可以空白";
				}

				CloseDbConn();
				
			}
			catch
			{
				//				ErrorHandler(e.ToString());
			}

		}


		public void MyDataGrid_Delete(Object sender, DataGridCommandEventArgs E) 
		{
			try
			{
				String sL_IsSupervisor = Session["ISSUPERVISOR"].ToString();
				String sL_IsSysOp = Session["ISSYSOP"].ToString();

				String sL_Msg = "";
				if ((sL_IsSysOp!="1") && (sL_IsSupervisor!="1")) //表示此user只是一般使用者
					sL_Msg = "無法刪除自己";
				
				if (sL_Msg=="")
				{
					ConnToDB( string.Empty );
					string sL_UserID = MyDataGrid.DataKeys[E.Item.ItemIndex].ToString();
					string sL_DeleteCmd = "delete EMC_EBT002 where UserId = '" + sL_UserID + "'";
				
					oracleCommandForEbtUser.CommandText = sL_DeleteCmd;			
					try 
					{
						oracleCommandForEbtUser.ExecuteNonQuery();
					}
					catch
					{					
					}
				
					BindGrid();
					CloseDbConn();
				}
				else
					lblInsertMsg.Text = sL_Msg;
			}
			catch
			{
				
			}
		}

		public void MyDataGrid_Edit(Object sender, DataGridCommandEventArgs E) 
		{
			try
			{
				MyDataGrid.EditItemIndex = E.Item.ItemIndex;
				BindGrid();

				/*
				String sL_Tmp  ="";

				sL_Tmp = E.Item.Cells[5].Controls[0].GetType().ToString();
				sL_Tmp = sL_Tmp + "-" + E.Item.Cells[5].Controls[1].GetType().ToString();
				sL_Tmp = sL_Tmp + "-" + E.Item.Cells[5].Controls[2].GetType().ToString();
				Label4.Text = sL_Tmp;
				
				MyDataGrid.EditItemIndex = E.Item.Cells[0].Controls.Count;
				MyDataGrid.EditItemIndex = E.Item.Cells[1].Controls.Count;
				MyDataGrid.EditItemIndex = E.Item.Cells[2].Controls.Count;
				MyDataGrid.EditItemIndex = E.Item.Cells[3].Controls.Count;
				MyDataGrid.EditItemIndex = E.Item.Cells[4].Controls.Count;
				MyDataGrid.EditItemIndex = E.Item.Cells[5].Controls.Count;
				*/
				/*
				System.Web.UI.WebControls.DropDownList L_DropDownList;		
				L_DropDownList = new DropDownList();
				System.Data.OracleClient.OracleDataReader L_DataReaderForUserGroup;
				String sL_Sql = "select CODENO, DESCRIPTION from EMC_EBT003 where STOPFLAG=0";
				oracleCommandForEbtUser.CommandText = sL_Sql;
				L_DataReaderForUserGroup = oracleCommandForEbtUser.ExecuteReader();

				L_DropDownList.DataSource = L_DataReaderForUserGroup;
				L_DropDownList.DataTextField = "DESCRIPTION";
				L_DropDownList.DataValueField = "CODENO";
				L_DropDownList.DataBind();		
				L_DropDownList.SelectedIndex = 0;
				
				
				
				E.Item.Cells[5].Controls.Add(L_DropDownList);
				*/
				

				/*
					ListControl L_DropDownList;
				//DropDownList L_DropDownList;
				L_DropDownList = (ListControl)E.Item.Cells[5].Controls[2];

				L_DropDownList.DataSource = G_DataReaderForUserGroup;
				L_DropDownList.DataTextField = "DESCRIPTION";
				L_DropDownList.DataValueField = "CODENO";
				L_DropDownList.DataBind();		
				L_DropDownList.SelectedIndex = 0;
				*/
/*
				DropDownList L_DropDownList;
				L_DropDownList = (DropDownList)E.Item.FindControl("lstUserGroup");

				L_DropDownList.DataSource = G_DataReaderForUserGroup;
				L_DropDownList.DataTextField = "DESCRIPTION";
				L_DropDownList.DataValueField = "CODENO";
				L_DropDownList.DataBind();		
				L_DropDownList.SelectedIndex = 0;

*/
				
				
			}
			catch
			{
				
				//				ErrorHandler(e.ToString());
			}
		}

		public void MyDataGrid_Cancel(Object sender, DataGridCommandEventArgs E) 
		{
			try
			{
				MyDataGrid.EditItemIndex = -1;
				BindGrid();
			}
			catch
			{
				//				ErrorHandler(e.ToString());
			}
		}

		public void MyDataGrid_Update(Object sender, DataGridCommandEventArgs E) 
		{
			try
			{
				ConnToDB( string.Empty );
				if (Page.IsValid) 
				{
					String sL_Msg = "";
					String sL_UserID = MyDataGrid.DataKeys[E.Item.ItemIndex].ToString();
					String sL_NewUserName = ((TextBox)E.Item.FindControl("edtUserNameForUpdate")).Text; 
					String sL_NewUserPasswd = ((TextBox)E.Item.FindControl("edtUserPasswdForUpdate")).Text; 
					//String sL_NewUserGroup = ((TextBox)E.Item.FindControl("edtUserGroupForUpdate")).Text; 
					//String sL_NewUserGroup = "1";
//lstUserGroupForInsert
					//String sL_NewUserIsSupervisor = ((TextBox)E.Item.FindControl("edtUserIsSupervisorForUpdate")).Text; 
					String sL_NewUserIsSupervisor = "";
					String sL_NewUserIsSysOp ="";
					String sL_NewStopFlag="";					


					String sL_OracleDateTimeStr = CommFun.getOracleDateTimeStr();
					CommFun L_CommFun = new CommFun();
					String sL_UpdEn = L_CommFun.GetCookie(Request, "USERNAME");


					//down, 找出 是否管理者 的值 ...
					
					RadioButtonList L_RadioButtonList;
					L_RadioButtonList = (RadioButtonList)E.Item.FindControl("rdbIsSupervisor");
					if (L_RadioButtonList.SelectedIndex==0)
						sL_NewUserIsSupervisor="1";
					else if (L_RadioButtonList.SelectedIndex==1)
						sL_NewUserIsSupervisor="0";
					
					//up, 找出 是否管理者 的值 ...


					//down, 找出 是否系統管理者 的值 ...
					
//					RadioButtonList L_RadioButtonList;
					L_RadioButtonList = (RadioButtonList)E.Item.FindControl("rdbIsSysOp");
					if (L_RadioButtonList.SelectedIndex==0)
						sL_NewUserIsSysOp="1";
					else if (L_RadioButtonList.SelectedIndex==1)
						sL_NewUserIsSysOp="0";
					
					//up, 找出 是否系統管理者 的值 ...

					//down, 找出 是否停用 的值 ...
					
					RadioButtonList L_RadioButtonListForStopFlag;
					L_RadioButtonListForStopFlag = (RadioButtonList)E.Item.FindControl("rdbStopFlag");
					if (L_RadioButtonListForStopFlag.SelectedIndex==0)
						sL_NewStopFlag="1";
					else if (L_RadioButtonListForStopFlag.SelectedIndex==1)
						sL_NewStopFlag="0";
					
					//up, 找出 是否停用 的值 ...

					//down, 檢查是否有權限可以更新...
					
					String sL_IsSupervisor = Session["ISSUPERVISOR"].ToString();
					String sL_IsSysOp = Session["ISSYSOP"].ToString();
					if ((sL_IsSysOp!="1") && (sL_IsSupervisor!="1")) //表示此user只是一般使用者
					{
						if (sL_NewUserIsSysOp=="1")
							sL_Msg = "不可以修改系統管理者權限";
						if (sL_NewUserIsSupervisor=="1")
							sL_Msg = "不可以修改群組管理者權限";
						if (sL_NewStopFlag=="1")
							sL_Msg = "不可以使自己停用";
					}
					else if ((sL_IsSysOp!="1") && (sL_IsSupervisor=="1")) //表示此user是群組管理者, 但不是系統管理員
					{
						if (sL_NewUserIsSysOp=="1")
							sL_Msg = "不可以修改系統管理者權限";
						
					}
					else if ((sL_IsSysOp=="1") && (sL_IsSupervisor=="1")) //表示此user是群組管理者, 也是系統管理員
					{
					}
					else if ((sL_IsSysOp=="1") && (sL_IsSupervisor=="0")) //表示此user不是群組管理者, 是系統管理員
					{
						if (sL_NewUserIsSupervisor=="1")
							sL_Msg = "不可以修改群組管理者權限";
					}
				
					//up, 檢查是否有權限可以更新...

					//string sL_UpdateCmd = "update EMC_EBT002 set USERNAME ='" +sL_NewUserName + "', USERPASSWD='"  + sL_NewUserPasswd + "', USERGROUP=" + sL_NewUserGroup + ", ISSUPERVISOR=" + sL_NewUserIsSupervisor  + ", ISSYSOP= "+ sL_NewUserIsSysOp  + ", UPDTIME=" + sL_OracleDateTimeStr + ", UPDEN='" + sL_UpdEn + "'  where USERID ='" + sL_UserID + "'";


					if (sL_Msg=="")
					{
						string sL_UpdateCmd = "";
						if (sL_UserID==Session["USERID"].ToString())
						  sL_UpdateCmd = "update EMC_EBT002 set IFCHANGEPASSWD='Y', USERNAME ='" +sL_NewUserName + "', USERPASSWD='"  + sL_NewUserPasswd +  "', ISSUPERVISOR=" + sL_NewUserIsSupervisor  + ", ISSYSOP= "+ sL_NewUserIsSysOp  + ",STOPFLAG=" + sL_NewStopFlag + ", UPDTIME=" + sL_OracleDateTimeStr + ", UPDEN='" + sL_UpdEn  + "'  where USERID ='" + sL_UserID + "'";// 不 update usergroup 的 sql...
						else
						  sL_UpdateCmd = "update EMC_EBT002 set USERNAME ='" +sL_NewUserName + "', USERPASSWD='"  + sL_NewUserPasswd +  "', ISSUPERVISOR=" + sL_NewUserIsSupervisor  + ", ISSYSOP= "+ sL_NewUserIsSysOp  + ",STOPFLAG=" + sL_NewStopFlag + ", UPDTIME=" + sL_OracleDateTimeStr + ", UPDEN='" + sL_UpdEn  + "'  where USERID ='" + sL_UserID + "'";// 不 update usergroup 的 sql...
						oracleCommandForEbtUser.CommandText = sL_UpdateCmd;
					}
					else
						lblInsertMsg.Text = sL_Msg;

					//Label1.Text = sL_UpdateCmd;

					/*
					string sL_UpdateCmd = "update EMC_EBT002 set   USERNAME = @USERNAME, USERPASSWD=@USERPASSWD, USERGROUP=@USERGROUP, ISSUPERVISOR=@ISSUPERVISOR, UPDTIME=@UPDTIME, UPDEN=@UPDEN where USERID = @USERID";
					

					oracleCommandForEbtUser.CommandText = sL_UpdateCmd;
					
					oracleCommandForEbtUser.Parameters.Add(new OracleParameter("@USERID", OracleType.VarChar  , 8));
					oracleCommandForEbtUser.Parameters.Add(new OracleParameter("@USERNAME", OracleType.VarChar  , 30));
					oracleCommandForEbtUser.Parameters.Add(new OracleParameter("@USERPASSWD", OracleType.VarChar  , 10));
					oracleCommandForEbtUser.Parameters.Add(new OracleParameter("@USERGROUP", OracleType.Number  , 2));
					oracleCommandForEbtUser.Parameters.Add(new OracleParameter("@ISSUPERVISOR", OracleType.Number  , 1));
					oracleCommandForEbtUser.Parameters.Add(new OracleParameter("@UPDTIME", OracleType.DateTime));
					oracleCommandForEbtUser.Parameters.Add(new OracleParameter("@UPDEN", OracleType.VarChar  , 20));

				

					oracleCommandForEbtUser.Parameters["@USERID"].Value = MyDataGrid.DataKeys[E.Item.ItemIndex];
					oracleCommandForEbtUser.Parameters["@USERNAME"].Value = ((TextBox)E.Item.FindControl("edtUserNameForUpdate")).Text; 

					oracleCommandForEbtUser.Parameters["@USERPASSWD"].Value = ((TextBox)E.Item.FindControl("edtUserPasswdForUpdate")).Text; 
					oracleCommandForEbtUser.Parameters["@USERGROUP"].Value = ((TextBox)E.Item.FindControl("edtUserGroupForUpdate")).Text; 
					oracleCommandForEbtUser.Parameters["@ISSUPERVISOR"].Value = ((TextBox)E.Item.FindControl("edtUserIsSupervisorForUpdate")).Text; 
					oracleCommandForEbtUser.Parameters["@UPDTIME"].Value = DateTime.Now;
					
					CommFun L_CommFun = new CommFun();
					oracleCommandForEbtUser.Parameters["@UPDEN"].Value = L_CommFun.GetCookie(Request, "USERNAME");
					*/

					try 
					{
						oracleCommandForEbtUser.ExecuteNonQuery();
						//MessageSQLUptDone();
						MyDataGrid.EditItemIndex = -1;
					}
					catch
					{
						//SQLErrorHandler(e);
	   
					}
					
					BindGrid();
				}
				CloseDbConn();
			}
			catch
			{
				//				ErrorHandler(e.ToString());
			}
		}
		/*
				private void InitializeComponent()
				{

				}
		*/
		private void BindGrid() 
		{
			try
			{
				int i=0;
				
				String sL_SQL = "";
				String sL_UserID = Session["USERID"].ToString();
				String sL_UserGroup = Session["USERGROUP"].ToString();
				String sL_IsSupervisor = Session["ISSUPERVISOR"].ToString();
				String sL_IsSysOp = Session["ISSYSOP"].ToString();
				btnAdd.Visible= true;
				rdbIsSysOpForInsert.Enabled = true;
				rdbIsSupervisorForInsert.Enabled = true;
				rdbStopFlagForInsert.Enabled = true;
				edtUserID.Enabled = true;
				edtUserPasswd.Enabled = true;
				edtUserName.Enabled = true;
				lstUserGroupForInsert.Enabled = true;

				

				if ((sL_IsSysOp!="1") && (sL_IsSupervisor!="1")) //表示此user只是一般使用者
				{
					sL_SQL = "select a.USERID,a.USERNAME,a.USERPASSWD, a.STOPFLAG,b.DESCRIPTION USERGROUP,b.CODENO, a.ISSUPERVISOR,a.ISSYSOP from EMC_EBT002 a, EMC_EBT003 b where a.UserID='" + sL_UserID + "' and a.USERGROUP=b.CODENO  and b.CODENO=" + sL_UserGroup ;
					edtUserID.Enabled = false;
					edtUserPasswd.Enabled = false;
					edtUserName.Enabled = false;

					
					btnAdd.Visible= false;
					/*
					lstUserGroupForInsert.Enabled = false;					  
					rdbIsSysOpForInsert.Enabled = false;
					rdbIsSupervisorForInsert.Enabled = false;
					rdbStopFlagForInsert.Enabled = false;
					*/
					lstUserGroupForInsert.Visible = false;
					lblUserGroupForInsert.Visible = false;
					rdbIsSysOpForInsert.Visible = false;
					rdbIsSupervisorForInsert.Visible = false;
					rdbStopFlagForInsert.Visible = false;

					lblIsSysOpForInsert.Visible = false;
					lblIsSupervisorForInsert.Visible = false;
					lblStopFlagForInsert.Visible = false;


				}
				else if ((sL_IsSysOp!="1") && (sL_IsSupervisor=="1")) //表示此user是群組管理者, 但不是系統管理員
				{
					sL_SQL = "select a.USERID,a.USERNAME,a.USERPASSWD, a.STOPFLAG,b.DESCRIPTION USERGROUP,b.CODENO, a.ISSUPERVISOR,a.ISSYSOP from EMC_EBT002 a, EMC_EBT003 b where  a.USERGROUP=b.CODENO  and b.CODENO=" + sL_UserGroup  + " order by UserID";
					rdbIsSysOpForInsert.Enabled = false;
					rdbIsSupervisorForInsert.Enabled = true;
					rdbStopFlagForInsert.Enabled = true;
					lstUserGroupForInsert.Enabled = false;

				}
				else if ((sL_IsSysOp=="1") && (sL_IsSupervisor=="1")) //表示此user是群組管理者, 也是系統管理員
					sL_SQL = "select a.USERID,a.USERNAME,a.USERPASSWD, a.STOPFLAG,b.DESCRIPTION USERGROUP,b.CODENO, a.ISSUPERVISOR,a.ISSYSOP from EMC_EBT002 a, EMC_EBT003 b where  a.USERGROUP=b.CODENO order by UserID";
				else if ((sL_IsSysOp=="1") && (sL_IsSupervisor=="0")) //表示此user不是群組管理者, 是系統管理員
					sL_SQL = "select a.USERID,a.USERNAME,a.USERPASSWD, a.STOPFLAG,b.DESCRIPTION USERGROUP,b.CODENO, a.ISSUPERVISOR,a.ISSYSOP from EMC_EBT002 a, EMC_EBT003 b where  a.USERGROUP=b.CODENO " + " and b.CODENO=" + sL_UserGroup +"　order by UserID";						

				ConnToDB( string.Empty );
				/*
				CommFun L_ConnFun = new CommFun();
				String sL_UserGroup = L_ConnFun.GetCookie(Request, "USERGROUP");
				*/
				//OracleDataAdapter L_MyCommand = new OracleDataAdapter("select * from EMC_EBT002 where USERGROUP=" + sL_UserGroup  + " order by UserID", oracleConnectionForEbtUser);
				OracleDataAdapter L_DataAdapter  = new OracleDataAdapter(sL_SQL, oracleConnectionForEbtUser);
				
				DataSet L_DataSet = new DataSet();
				L_DataAdapter.Fill(L_DataSet);

				
					
				
					String sL_FilterExpression="CODENO="+sL_UserGroup;
					DataRow[] L_FoundRows;
					L_FoundRows = L_DataSet.Tables[0].Select(sL_FilterExpression);
					sL_UserGroup = L_FoundRows[0][3].ToString();

					for (i=0; i<=lstUserGroupForInsert.Items.Count-1; i++)
					{
						if (lstUserGroupForInsert.Items[i].Text==sL_UserGroup)
						{
							lstUserGroupForInsert.SelectedIndex = i;
							
							break;
						}

					}
				
				

				MyDataGrid.DataSource=L_DataSet;
				/*
				L_MyCommand.Fill(L_DataSet, "USERDATA");
				MyDataGrid.DataSource=L_DataSet.Tables["USERDATA"].DefaultView;
				*/
				MyDataGrid.DataBind();
				CloseDbConn();
			}
			catch
			{
				//				ErrorHandler(e.ToString());
			}
		}

		private void Submit1_ServerClick(object sender, System.EventArgs e)
		{
		
		}

		
	}
}
