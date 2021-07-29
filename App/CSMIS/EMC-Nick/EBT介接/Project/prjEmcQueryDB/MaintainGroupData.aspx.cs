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
using System.Configuration;

namespace prjEmcQueryDB
{
	/// <summary>
	/// MaintainUserData1 的摘要描述。
	/// </summary>
	public class MaintainGroupData : BasePage
	{
		protected System.Web.UI.WebControls.DataGrid MyDataGrid;
		protected System.Web.UI.HtmlControls.HtmlInputButton Submit1;
		protected System.Web.UI.WebControls.Label lblInsertMsg;
		protected System.Web.UI.HtmlControls.HtmlGenericControl Message;
		protected System.Web.UI.WebControls.TextBox edtCodeNo;
		protected System.Web.UI.WebControls.RadioButtonList rdbStopFlagForInsert;
		protected System.Web.UI.WebControls.CheckBoxList chbCompList;
		protected System.Web.UI.WebControls.Label Label2;
		protected System.Web.UI.WebControls.TextBox edtDesc;
		
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

		
		private void Page_Load(object sender, System.EventArgs e)
		{
			// 在這裡放置使用者程式碼以初始化網頁
			
			CommFun aCommFun = new CommFun();
			aCommFun.RedirectToLogin( this );

			ConnToDB( string.Empty ); //不可以放在  if (!IsPostBack) 裡面, 否則會有錯誤..

			if ( !IsPostBack ) BindGrid();
			
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
			this.Submit1.ServerClick += new System.EventHandler(this.Submit1_ServerClick);
			this.Load += new System.EventHandler(this.Page_Load);

		}
		#endregion

		
		public void Add_Click(Object sender, EventArgs E) 
		{
			try
			{
				int i=0;
				String sL_NewCodeNo = edtCodeNo.Text.Trim();
				String sL_NewDesc = edtDesc.Text.Trim();

				String sL_SoName="", sL_SoCode="";
				for (i=0; i<=chbCompList.Items.Count-1; i++)
				{
					if (chbCompList.Items[i].Selected)
					{
						if (sL_SoCode=="")
						  sL_SoCode = chbCompList.Items[i].Value.ToString();
						else
					      sL_SoCode = sL_SoCode + ConfigurationSettings.AppSettings["SoSep"] + chbCompList.Items[i].Value.ToString();

						if (sL_SoName=="")
							sL_SoName = chbCompList.Items[i].Text.ToString();
						else
							sL_SoName = sL_SoName + ConfigurationSettings.AppSettings["SoSep"] + chbCompList.Items[i].Text.ToString();
					}
				}

				if( (sL_NewCodeNo!="") && (sL_NewDesc!="") && (sL_SoCode!="")  && (sL_SoName!="") )
				{
					//String sL_NewUserGroup = edtUserGroup.Text.Trim();
					//String sL_NewUserIsSupervisor = edtIsSupervisor.Text.Trim();
				
					
					String sL_NewStopFlag = rdbStopFlagForInsert.SelectedValue;

					String sL_OracleDateTimeStr = CommFun.getOracleDateTimeStr();
					CommFun L_CommFun = new CommFun();
					String sL_UpdEn = L_CommFun.GetCookie(Request, "USERNAME");

					String sL_Sql = "select CODENO from EMC_EBT003 where CODENO=" +  sL_NewCodeNo + "" ;
					string sL_InsertCmd = "";


					oracleCommandForEbtUser.CommandText = sL_Sql;
					//lblMessage.Text = oracleCommandForEbtUser.CommandText;
					System.Data.OracleClient.OracleDataReader L_DataReader;

					L_DataReader = oracleCommandForEbtUser.ExecuteReader();
				
					//lblMessage.Text = L_DataReader.GetOracleString(1).ToString();
					if (!L_DataReader.HasRows)
					{

						sL_InsertCmd = "insert into EMC_EBT003 (CODENO,DESCRIPTION,SONAME, SOCODE, STOPFLAG,UPDTIME,UPDEN) values (" + sL_NewCodeNo + ",'" + sL_NewDesc + "','" + sL_SoName + "','" + sL_SoCode + "'," +  sL_NewStopFlag + "," + sL_OracleDateTimeStr + ",'" + sL_UpdEn + "')";				
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
						lblInsertMsg.Text = "群組代號(" + sL_NewCodeNo + ")已經存在!" ;
					}
					L_DataReader.Close();
				}
				else
				{
					lblInsertMsg.Text = "群組代號,名稱, 所屬系統台都不可以空白";
				}
				//				Label1.Text = sL_InsertCmd;

				/*
				string sL_InsertCmd = "insert into EMC_EBT002 (USERID,USERNAME,USERPASSWD,USERGROUP,ISSUPERVISOR,UPDTIME,UPDEN) values (@USERID,@USERNAME,@USERPASSWD,@USERGROUP,@ISSUPERVISOR,@UPDTIME,@UPDEN)";
				//SqlCommand myCommand = new SqlCommand(insertCmd, myConnection);

				oracleCommandForEbtUser.CommandText = sL_InsertCmd;
				oracleCommandForEbtUser.Parameters.Add(new OracleParameter("@USERID", OracleType.VarChar  , 8));
				oracleCommandForEbtUser.Parameters.Add(new OracleParameter("@USERNAME", OracleType.VarChar  , 30));
				oracleCommandForEbtUser.Parameters.Add(new OracleParameter("@USERPASSWD", OracleType.VarChar  , 10));
				oracleCommandForEbtUser.Parameters.Add(new OracleParameter("@USERGROUP", OracleType.Number  , 2));
				oracleCommandForEbtUser.Parameters.Add(new OracleParameter("@ISSUPERVISOR", OracleType.Number  , 1));
				oracleCommandForEbtUser.Parameters.Add(new OracleParameter("@UPDTIME", OracleType.DateTime));
				oracleCommandForEbtUser.Parameters.Add(new OracleParameter("@UPDEN", OracleType.VarChar  , 20));




				oracleCommandForEbtUser.Parameters["@USERID"].Value = edtUserID.Text.Trim();        
				oracleCommandForEbtUser.Parameters["@USERNAME"].Value = edtUserName.Text.Trim();

				oracleCommandForEbtUser.Parameters["@USERPASSWD"].Value = edtUserPasswd.Text.Trim();
				oracleCommandForEbtUser.Parameters["@USERGROUP"].Value = edtUserGroup.Text.Trim();
				oracleCommandForEbtUser.Parameters["@ISSUPERVISOR"].Value = edtIsSupervisor.Text.Trim();
				oracleCommandForEbtUser.Parameters["@UPDTIME"].Value = DateTime.Now;

				CommFun L_CommFun = new CommFun();
				oracleCommandForEbtUser.Parameters["@UPDEN"].Value = L_CommFun.GetCookie(Request, "USERNAME");
				*/
		
        

				
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
				
				

				String sL_CodeNo = MyDataGrid.DataKeys[E.Item.ItemIndex].ToString();
				String sL_DeleteCmd = "delete EMC_EBT003 where CodeNo = " + sL_CodeNo + "";
				
				oracleCommandForEbtUser.CommandText = sL_DeleteCmd;
				/*
				oracleCommandForEbtUser.Parameters.Add(new OracleParameter("@USERID", OracleType.VarChar  , 8));;
				oracleCommandForEbtUser.Parameters["@USERID"].Value = MyDataGrid.DataKeys[E.Item.ItemIndex];
				*/
				
				try 
				{
					oracleCommandForEbtUser.ExecuteNonQuery();
					//MessageSQLDelDone();
				}
				catch
				{
					//SQLErrorHandler(e);
				}
				
				BindGrid();
			}
			catch
			{
				//				ErrorHandler(e.ToString());
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
				if (Page.IsValid) 
				{
					String sL_CodeNo = MyDataGrid.DataKeys[E.Item.ItemIndex].ToString();
					String sL_NewDesc = ((TextBox)E.Item.FindControl("edtDescForUpdate")).Text; 
					String sL_NewStopFlag="";					

					String sL_OracleDateTimeStr = CommFun.getOracleDateTimeStr();
					CommFun L_CommFun = new CommFun();
					String sL_UpdEn = L_CommFun.GetCookie(Request, "USERNAME");


					//down, 找出 是否停用 的值 ...
					
					RadioButtonList L_RadioButtonList;
					L_RadioButtonList = (RadioButtonList)E.Item.FindControl("rdbStopFlag");
					if (L_RadioButtonList.SelectedIndex==0)
						sL_NewStopFlag="1";
					else if (L_RadioButtonList.SelectedIndex==1)
						sL_NewStopFlag="0";
					
					//up, 找出 是否停用 的值 ...

					string sL_UpdateCmd = "update EMC_EBT003 set DESCRIPTION='"  + sL_NewDesc + "', STOPFLAG=" + sL_NewStopFlag + ", UPDTIME=" + sL_OracleDateTimeStr + ", UPDEN='" + sL_UpdEn + "'  where CODENO ='" + sL_CodeNo + "'";
					oracleCommandForEbtUser.CommandText = sL_UpdateCmd;

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
				
				
				OracleDataAdapter L_MyCommand = new OracleDataAdapter("select * from EMC_EBT003 order by CODENO", oracleConnectionForEbtUser);
				DataSet L_DataSet = new DataSet();
				L_MyCommand.Fill(L_DataSet);
				
				
				MyDataGrid.DataSource=L_DataSet;
				/*
				L_MyCommand.Fill(L_DataSet, "USERDATA");
				MyDataGrid.DataSource=L_DataSet.Tables["USERDATA"].DefaultView;
				*/
				MyDataGrid.DataBind();
			}
			catch
			{
				//				ErrorHandler(e.ToString());
			}
		}

		private void edtUserName_ServerChange(object sender, System.EventArgs e)
		{
		
		}

		private void MyDataGrid_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		
		}

		private void Submit1_ServerClick(object sender, System.EventArgs e)
		{
		
		}
	}
}
