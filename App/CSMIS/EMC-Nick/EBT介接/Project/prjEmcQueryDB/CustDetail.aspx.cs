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
	/// CustDetail 的摘要描述。
	/// </summary>
	public class CustDetail :BasePage
	{
		protected System.Web.UI.WebControls.DataGrid dbgEmc;
		protected System.Web.UI.WebControls.Label lblEbtMessage;
		protected System.Web.UI.WebControls.Label lblQueryResult;
		protected System.Web.UI.WebControls.DataGrid dbgEbt;


		private string GetServiceTypeColor(string aServiceType)
		{
			string aResult = "red";
			if ( ( aServiceType == ConfigurationSettings.AppSettings["CmKeyWord"] ) ||
				  ( aServiceType == ConfigurationSettings.AppSettings["CpKeyWord"] ) )
			{
				aResult = "blue";
			}
			return aResult;
		}

		private int setEmcDataColumn(DataTable I_DataTable)
		{
						
			I_DataTable.Columns.Add(new DataColumn("系統台",typeof(string)));
			I_DataTable.Columns.Add(new DataColumn("客編",typeof(string)));
			I_DataTable.Columns.Add(new DataColumn("客戶名稱<br>裝機地址",typeof(string)));
			I_DataTable.Columns.Add(new DataColumn("裝機日",typeof(string)));
			I_DataTable.Columns.Add(new DataColumn("拆機日",typeof(string)));
			I_DataTable.Columns.Add(new DataColumn("服務別",typeof(string)));
			I_DataTable.Columns.Add(new DataColumn("繳費週期(月)",typeof(string)));
			I_DataTable.Columns.Add(new DataColumn("客戶狀態",typeof(string)));
			I_DataTable.Columns.Add(new DataColumn("電話",typeof(string)));	
			I_DataTable.Columns.Add(new DataColumn("大樓名稱",typeof(string)));
			I_DataTable.Columns.Add(new DataColumn("建築型態",typeof(string)));
			I_DataTable.Columns.Add(new DataColumn("管線類別",typeof(string)));
			I_DataTable.Columns.Add(new DataColumn("Fiber Node",typeof(string)));

			return 0;
		}

		private int setEbtDataColumn(DataTable I_DataTable)
		{
			
			I_DataTable.Columns.Add(new DataColumn("服務別",typeof(string)));						
			I_DataTable.Columns.Add(new DataColumn("客戶資料",typeof(string)));
			I_DataTable.Columns.Add(new DataColumn("合約資料",typeof(string)));
			I_DataTable.Columns.Add(new DataColumn("公司資料",typeof(string)));
			I_DataTable.Columns.Add(new DataColumn("技術聯絡人資料",typeof(string)));
			I_DataTable.Columns.Add(new DataColumn("地址資料",typeof(string)));
			I_DataTable.Columns.Add(new DataColumn("代理人資料",typeof(string)));
			
			return 0;
		}

		private void Page_Load(object sender, System.EventArgs e)
		{
			// 在這裡放置使用者程式碼以初始化網頁
			
			if (!IsPostBack)
			{
				CommFun aCommFun = new CommFun();
				aCommFun.RedirectToLogin( this );

				string aCustId = Request.QueryString["CUSTID"] ;
				string aCompCode = Request.QueryString["COMPCODE"] ;
				string aIfMoveToOtherSo = string.Empty; 
				string aDataRow0 = string.Empty;
				
				string aCompName = CommFun.GetCompChineseName( byte.Parse( aCompCode ) );
				bool aHasHeader = false; 
				bool aHasEbtData = false;

				string sL_EmcCustName = "", sL_EmcCustStatusName = "";
				string sL_EmcClassName1 = "", sL_EmcClassName2 = "",sL_EmcClassName3 = "", sL_EmcInstTime = "", sL_EmcPrTime="", sL_EmcTel1 = "",sL_EmcTel2 = "";
				string sL_EmcTel3 = "", sL_EmcInstAddr = "", sL_EmcMduName = "";
				string sL_EmcBtName = "", sL_EmcPipelineName = "", sL_EmcNodeNo = "", sL_EmcPeriod;

				string sL_EbtServiceType ="", sL_EbtCustID ="", sL_EbtCustCName="", sL_EbtCustContactPhone="";
				string sL_EbtCustContactMobile="", sL_EbtContraceNo="", sL_EbtContractBDate ="", sL_EbtContractEate="";
				string sL_EbtContractStatusCode="", sL_EbtContractStatusDesc="", sL_EbtFeePeriodCode="", sL_EbtFeePeriodDesc ="";
				string sL_EbtCompOwnerName ="", sL_EbtContactPhone =""; 
				string sL_EbtItContactName ="", sL_EbtItContactPhone="", sL_EbtItContactMobile="";
				string sL_EbtItEMail="", sL_EbtInstAddr="", sL_EbtCustAddr ="", sL_EbtBillAddr ="";
				string sL_EbtAgentName="",sL_EbtAgentPhone="", sL_EbtAgentAddress="";			


				System.Data.OracleClient.OracleDataReader aReader;
				System.Data.OracleClient.OracleDataReader aReader2;

				string sL_ViewName = CommFun.GetQueryViewName( byte.Parse( aCompCode ), 1 );

				DataTable L_EmcDataTable;
				DataRow  L_EmcDataRow;
			
			
				DataTable aEbtDataTable;
				DataRow  aEbtDataRow;
			
				ConnToDB( aCompCode );				

				L_EmcDataTable = new DataTable();
				setEmcDataColumn( L_EmcDataTable );

				aEbtDataTable = new DataTable();
				setEbtDataColumn( aEbtDataTable );

				if ( CommFun.IsEbtDataArea( aCompCode ) )
				{
					oracleCommandForEbtUser.CommandText = string.Format( 
						"SELECT EMCCUSTNAME,                                       " +
						"       EMCCUSTSTATUSCODE,                                 " +
						"       EMCCUSTSTATUSNAME,                                 " + 
						"       TO_CHAR(EMCINSTTIME,'YYYY/MM/DD') EMCINSTTIME,     " + 
						"       TO_CHAR(EMCPRTIME,'YYYY/MM/DD') EMCPRTIME,         " +
						"       EMCTEL1, " + 
						"       EMCTEL2, " + 
						"       EMCTEL3, " + 
						"       EMCINSTADDRESS,   " + 
						"       EMCCLASSNAME1,    " + 
						"       EMCCLASSNAME2,    " + 
						"       EMCCLASSNAME3,    " + 
						"       EMCMDUNAME,       " + 
						"       EMCBTNAME,        " +
						"       EMCPIPELINENAME,  " + 
						"       EMCNODENO,        " + 
						"       EMCPERIOD,        " + 
						"       EBTCONTRACTNO,    " + 
						"       EBTCUSTID,        " + 
						"       TO_CHAR(EBTCONTRACTBDATE,'YYYY/MM/DD') EBTCONTRACTBDATE, " + 
						"       TO_CHAR(EBTCONTRACTEDATE,'YYYY/MM/DD') EBTCONTRACTEDATE, " +
						"       EBTCUSTCNAME,                         " +
						"       EBTCUSTCONTACTPHONE,                  " +
						"       EBTCUSTCONTACTMOBILE,                 " +
						"       EBTCOMPOWNERNAME,                     " +
						"       EBTCONTACTPHONE,                      " +
						"       EBTITCONTACTNAME,                     " +
						"       EBTITCONTACTPHONE,                    " + 
						"       EBTITCONTACTMOBILE,                   " + 
						"       EBTITEMAIL,                           " + 
						"       EBTINSTADDR,                          " + 
						"       EBTCUSTADDR,                          " +
						"       EBTBILLADDR,                          " +
						"       EBTCONTRACTSTATUSCODE,                " + 
						"       EBTCONTRACTSTATUSDESC,                " + 
						"       EBTFEEPERIODCODE,                     " +
						"       EBTFEEPERIODDESC,                     " + 
						"       EBTSERVICETYPE,                       " + 
						"       EBTAGENTNAME,                         " + 
						"       EBTAGENTPHONE,                        " + 
						"       EBTAGENTADDRESS,                      " + 
						"       IFMOVETOOTHERSO                       " +
						"  FROM {0}                                   " + 
						" WHERE EMCCOMPCODE = {1}                     " +
						"   AND EMCCUSTID = {2}                       " +
						" ORDER BY EBTSERVICETYPE, EMCCUSTSTATUSCODE  ", sL_ViewName, aCompCode, aCustId  );
				}
				else
				{
					oracleCommandForEbtUser.CommandText = string.Format(
						" SELECT A.CUSTNAME AS EMCCUSTNAME,                             " +
						"        B.CUSTSTATUSCODE AS EMCCUSTSTATUSCODE,                 " +
						"        B.CUSTSTATUSNAME AS EMCCUSTSTATUSNAME,                 " + 
						"        TO_CHAR( B.INSTTIME, 'YYYY/MM/DD' ) EMCINSTTIME,       " + 
						"        TO_CHAR( B.PRTIME,'YYYY/MM/DD') EMCPRTIME,             " +
						"        A.TEL1 AS EMCTEL1,                                     " +
						"        A.TEL2 AS EMCTEL2,                                     " +
						"        A.TEL3 AS EMCTEL3,                                     " +
						"        A.INSTADDRESS AS EMCINSTADDRESS,                       " + 
					   "        A.CLASSNAME1 AS EMCCLASSNAME1,                         " +
					   "        A.CLASSNAME2 AS EMCCLASSNAME2,                         " +
					   "        A.CLASSNAME3 AS EMCCLASSNAME3,                         " +
						"        D.MDUNAME AS EMCMDUNAME,                               " + 
						"        D.BTNAME AS EMCBTNAME,                                 " +
						"        D.PIPELINENAME AS EMCPIPELINENAME,                     " + 
						"        D.NODENO AS EMCNODENO,                                 " + 
						"        E.PERIOD AS EMCPERIOD                                  " + 
						"    FROM SO001 A , SO002 B, SO014 D, SO003 E                   " + 
					   "   WHERE A.CUSTID = B.CUSTID                                   " + 
					   "     AND A.COMPCODE = B.COMPCODE                               " + 
					   "     AND B.SERVICETYPE = 'C'                                   " + 
					   "     AND A.COMPCODE = D.COMPCODE                               " + 
					   "     AND A.INSTADDRNO = D.ADDRNO                               " + 
					   "     AND A.CUSTID = E.CUSTID(+)                                " + 
						"     AND E.SERVICETYPE(+) = 'C'                                " +
						"     AND A.CUSTID = '{0}'                                      "	+
						"   ORDER BY B.CUSTSTATUSCODE                                   ", aCustId );

				}

				ConnToDB( aCompCode );

				aReader = oracleCommandForEbtUser.ExecuteReader();
				while ( aReader.Read() )
				{
					/* CATV 資料 */

					if ( !aHasHeader )
					{
						
						sL_EmcCustName = aReader["EMCCUSTNAME"].ToString().Trim();
						if ( aReader["EMCCUSTSTATUSCODE"].ToString() == "1" ) 
							sL_EmcCustStatusName = string.Format( "<font color=RED>{0}</font>", aReader["EMCCUSTSTATUSNAME"].ToString().Trim() );
						else
							sL_EmcCustStatusName = aReader["EMCCUSTSTATUSNAME"].ToString().Trim();

						sL_EmcClassName1 = aReader["EMCCLASSNAME1"].ToString().Trim();
						sL_EmcClassName2 = aReader["EMCCLASSNAME2"].ToString().Trim();
						sL_EmcClassName3 = aReader["EMCCLASSNAME3"].ToString().Trim();
						sL_EmcInstTime = aReader["EMCINSTTIME"].ToString().Trim();
						sL_EmcPrTime = aReader["EMCPRTIME"].ToString().Trim();

						sL_EmcTel1 = aReader["EMCTEL1"].ToString().Trim();
						sL_EmcTel2 = aReader["EMCTEL2"].ToString().Trim();
						sL_EmcTel3 = aReader["EMCTEL3"].ToString().Trim();
						sL_EmcInstAddr = aReader["EMCINSTADDRESS"].ToString().Trim();
						sL_EmcMduName = aReader["EMCMDUNAME"].ToString().Trim();
						sL_EmcBtName = aReader["EMCBTNAME"].ToString().Trim();
					
						sL_EmcPipelineName = aReader["EMCPIPELINENAME"].ToString().Trim();
						sL_EmcNodeNo = aReader["EMCNODENO"].ToString().Trim();
						sL_EmcPeriod = aReader["EMCPERIOD"].ToString().Trim();

						L_EmcDataRow = L_EmcDataTable.NewRow();
						L_EmcDataRow[0] = aCompName;
						L_EmcDataRow[1] = aCustId;
						L_EmcDataRow[2] = sL_EmcCustName + "<br>" + "<font color=MediumVioletRed>" + sL_EmcInstAddr + "<font>";
						L_EmcDataRow[3] = sL_EmcInstTime;
						L_EmcDataRow[4] = sL_EmcPrTime;
						L_EmcDataRow[5] = "一." + sL_EmcClassName1 + "<br>二." + sL_EmcClassName2 + "<br>三." + sL_EmcClassName3;
						L_EmcDataRow[6] = sL_EmcPeriod;
						L_EmcDataRow[7] = sL_EmcCustStatusName;
						L_EmcDataRow[8] = "一." + sL_EmcTel1 + "<br>二." + sL_EmcTel2 + "<br>三." + sL_EmcTel3;
						L_EmcDataRow[9] = sL_EmcMduName;
						L_EmcDataRow[10] = sL_EmcBtName;
						L_EmcDataRow[11] = sL_EmcPipelineName;
						L_EmcDataRow[12] = sL_EmcNodeNo;
						L_EmcDataTable.Rows.Add( L_EmcDataRow );

						aHasHeader = true;
					}
					
					if ( CommFun.IsEbtDataArea( aCompCode ) )
					{
					
						/* CM/CP 資料 */

						sL_EbtServiceType = aReader["EBTSERVICETYPE"].ToString().Trim();

						if ( sL_EbtServiceType != string.Empty )
							aHasEbtData = true;

						sL_EbtCustID = aReader["EBTCUSTID"].ToString().Trim();
						sL_EbtCustCName = aReader["EBTCUSTCNAME"].ToString().Trim();
						sL_EbtCustContactPhone = aReader["EBTCUSTCONTACTPHONE"].ToString().Trim();
						sL_EbtCustContactMobile = aReader["EBTCUSTCONTACTMOBILE"].ToString().Trim();
						sL_EbtContraceNo = aReader["EBTCONTRACTNO"].ToString().Trim();
						sL_EbtContractBDate = aReader["EBTCONTRACTBDATE"].ToString().Trim();
						sL_EbtContractEate = aReader["EBTCONTRACTEDATE"].ToString().Trim();
						sL_EbtContractStatusCode = aReader["EBTCONTRACTSTATUSCODE"].ToString().Trim();
						sL_EbtContractStatusDesc = aReader["EBTCONTRACTSTATUSDESC"].ToString().Trim();
						sL_EbtFeePeriodCode = aReader["EBTFEEPERIODCODE"].ToString().Trim();
						sL_EbtFeePeriodDesc = aReader["EBTFEEPERIODDESC"].ToString().Trim();
						sL_EbtCompOwnerName = aReader["EBTCOMPOWNERNAME"].ToString().Trim();
						sL_EbtContactPhone = aReader["EBTCONTACTPHONE"].ToString().Trim();                
						sL_EbtItContactName = aReader["EBTITCONTACTNAME"].ToString().Trim();
						sL_EbtItContactPhone = aReader["EBTITCONTACTPHONE"].ToString().Trim();
						sL_EbtItContactMobile = aReader["EBTITCONTACTMOBILE"].ToString().Trim();
						sL_EbtItEMail = aReader["EBTITEMAIL"].ToString().Trim();                    
						sL_EbtInstAddr = aReader["EBTINSTADDR"].ToString().Trim();
						sL_EbtCustAddr = aReader["EBTCUSTADDR"].ToString().Trim();
						sL_EbtBillAddr = aReader["EBTBILLADDR"].ToString().Trim();
						sL_EbtAgentName = aReader["EBTAGENTNAME"].ToString().Trim();
						sL_EbtAgentPhone = aReader["EBTAGENTPHONE"].ToString().Trim();					
						sL_EbtAgentAddress = aReader["EBTAGENTADDRESS"].ToString().Trim();
						aIfMoveToOtherSo = aReader["IFMOVETOOTHERSO"].ToString().Trim().ToUpper();

						aDataRow0 = ( aIfMoveToOtherSo != "Y" ) ? 
							string.Format( "<font color={0}>{1}</font>", GetServiceTypeColor(sL_EbtServiceType), sL_EbtServiceType ) :
							string.Format( "<font color={0}>{1}</font><br>(已移機)", GetServiceTypeColor(sL_EbtServiceType), sL_EbtServiceType );

						aEbtDataRow = aEbtDataTable.NewRow();
						aEbtDataRow[0] = aDataRow0; 
						aEbtDataRow[1] = "客編:" +  sL_EbtCustID  + "<br>" + "名稱:" + sL_EbtCustCName + "<br>" + "電話:" + sL_EbtCustContactPhone + "<br>" + "手機:" +  sL_EbtCustContactMobile;
						aEbtDataRow[2] = "編號:" + sL_EbtContraceNo + "<br>" + "生效日:" + sL_EbtContractBDate + "<br>" + "截止日:" + sL_EbtContractEate + "<br>" + "狀態" + "<b>:</b>" + sL_EbtContractStatusDesc + "(" + sL_EbtContractStatusCode + ")" + "<br>" + "繳費週期" + "<b>:</b>" + sL_EbtFeePeriodDesc + "(" + sL_EbtFeePeriodCode + ")";
						aEbtDataRow[3] = "負責人姓名:" + sL_EbtCompOwnerName + "<br>" + "電話:" +  sL_EbtContactPhone ;
						aEbtDataRow[4] = "姓名:" + sL_EbtItContactName + "<br>" + "電話:" +  sL_EbtItContactPhone + "<br>" + "手機:" + sL_EbtItContactMobile + "<br>" + "EMail:" + sL_EbtItEMail;
						aEbtDataRow[5] = "裝機地址:" + sL_EbtInstAddr + "<br>" + "戶籍地址:" + sL_EbtCustAddr + "<br>" + "帳單地址:" + sL_EbtBillAddr;					
						aEbtDataRow[6] = "姓名:" + sL_EbtAgentName + "<br>" + "電話:" + sL_EbtAgentPhone+ "<br>" + "地址:" + sL_EbtAgentAddress;

						aEbtDataTable.Rows.Add( aEbtDataRow );
					}
					else
					{
						ConnToDBEx( aCompCode );
						try
						{
							TempCommand.CommandText = string.Format( 
								"  SELECT DECODE( A.SERVICETYPE, 'C', 'CT',								  " +
								"                                'I', 'CM',								  " +
								"                                'P', 'CP' ) AS EBTSERVICETYPE,     " +
								"         A.DECLARANTNAME AS EBTCUSTCNAME,                          " +
								"         A.CONTTEL AS EBTCUSTCONTACTPHONE,                         " +
								"         A.CONTNO AS EBTCONTRACTNO,                                " +
								"         TO_CHAR( A.INSTDATE, 'YYYY/MM/DD' ) AS EBTCONTRACTBDATE,  " +
								"         A.CONTRACT_STATUS AS EBTCONTRACTSTATUSCODE,               " +
								"         A.CONTRACT_STATUS_DESC AS EBTCONTRACTSTATUSDESC,          " + 
								"         B.BPCODE AS EBTFEEPERIODCODE,									  " +
								"         B.BPNAME AS EBTFEEPERIODDESC,									  " +
								"         A.BOSS AS EBTCOMPOWNERNAME,                               " + 								
								"         A.CONTNAME2 AS EBTITCONTACTNAME,                          " +
								"         A.CONTTEL2 AS EBTITCONTACTPHONE,                          " +
								"         A.EMAIL AS EBTITEMAIL,                                    " +
								"         C.INSTADDRESS AS EBTINSTADDR,                             " +								
								"         A.DOMICILEADDRESS AS EBTCUSTADDR,                         " +
								"         C.MAILADDRESS AS EBTBILLADDR1,                            " +
								"         D.MAILADDRESS AS EBTBILLADDR2,                            " +
								"         A.AGENTNAME AS EBTAGENTNAME,                              " +
								"         A.AGENTTEL AS EBTAGENTPHONE,                              " +
								"         A.DOMICILEADDRESS2 AS EBTAGENTADDRESS                     " +
								"    FROM SO004 A, SO003 B, SO001 C, SO002A D                       " +
								"   WHERE A.CUSTID = B.CUSTID(+)                                    " +
								"     AND A.SERVICETYPE = B.SERVICETYPE(+)                          " + 
								"     AND A.CONTNO = B.CONTNO(+)                                    " +
								"     AND A.CUSTID = C.CUSTID(+)                                    " +
								"     AND B.CUSTID = D.CUSTID(+)                                    " +
								"     AND B.ACCOUNTNO = D.ACCOUNTNO(+)                              " +
								"     AND A.CUSTID = '{0}'                                          " +
								"     AND A.SERVICETYPE IN ( 'I' )                                  ", aCustId );

							aReader2 = TempCommand.ExecuteReader();

							while ( aReader2.Read() )
							{
								aHasEbtData = true;
								sL_EbtCustID = aCustId;
								sL_EbtCustContactMobile = string.Empty;
								sL_EbtContractEate = string.Empty;
								sL_EbtContactPhone = string.Empty;
								aIfMoveToOtherSo = string.Empty;

								sL_EbtCustID = string.Empty;
								sL_EbtCustCName = aReader2["EBTCUSTCNAME"].ToString().Trim();
								sL_EbtCustContactPhone = aReader2["EBTCUSTCONTACTPHONE"].ToString().Trim();
								sL_EbtContraceNo = aReader2["EBTCONTRACTNO"].ToString().Trim();
								sL_EbtContractBDate = aReader2["EBTCONTRACTBDATE"].ToString().Trim();								
								sL_EbtContractStatusCode = aReader2["EBTCONTRACTSTATUSCODE"].ToString().Trim();
								sL_EbtContractStatusDesc = aReader2["EBTCONTRACTSTATUSDESC"].ToString().Trim();
								sL_EbtFeePeriodCode = aReader2["EBTFEEPERIODCODE"].ToString().Trim();
								sL_EbtFeePeriodDesc = aReader2["EBTFEEPERIODDESC"].ToString().Trim();
								sL_EbtCompOwnerName = aReader2["EBTCOMPOWNERNAME"].ToString().Trim();
								
								sL_EbtItContactName = aReader2["EBTITCONTACTNAME"].ToString().Trim();
								sL_EbtItContactPhone = aReader2["EBTITCONTACTPHONE"].ToString().Trim();
								sL_EbtItEMail = aReader2["EBTITEMAIL"].ToString().Trim();
								sL_EbtInstAddr = aReader2["EBTINSTADDR"].ToString().Trim();
								sL_EbtCustAddr = aReader2["EBTCUSTADDR"].ToString().Trim();

								sL_EbtBillAddr = aReader2["EBTBILLADDR2"].ToString().Trim();
								if ( sL_EbtBillAddr == string.Empty )
									sL_EbtBillAddr = aReader2["EBTBILLADDR1"].ToString().Trim();						
								
								sL_EbtAgentName = aReader2["EBTAGENTNAME"].ToString().Trim();
								sL_EbtAgentPhone = aReader2["EBTAGENTPHONE"].ToString().Trim();					
								sL_EbtAgentAddress = aReader2["EBTAGENTADDRESS"].ToString().Trim();
								 

								aDataRow0 = string.Format( "<font color={0}>{1}</font>", 
									GetServiceTypeColor( aReader2["EBTSERVICETYPE"].ToString() ),
									aReader2["EBTSERVICETYPE"].ToString() );
									

								aEbtDataRow = aEbtDataTable.NewRow();
								aEbtDataRow[0] = aDataRow0; 
								aEbtDataRow[1] = "客編:" +  sL_EbtCustID  + "<br>" + "名稱:" + sL_EbtCustCName + "<br>" + "電話:" + sL_EbtCustContactPhone + "<br>" + "手機:" +  sL_EbtCustContactMobile;
								aEbtDataRow[2] = "編號:" + sL_EbtContraceNo + "<br>" + "生效日:" + sL_EbtContractBDate + "<br>" + "截止日:" + sL_EbtContractEate + "<br>" + "狀態" + "<b>:</b>" + sL_EbtContractStatusDesc + "(" + sL_EbtContractStatusCode + ")" + "<br>" + "繳費週期" + "<b>:</b>" + sL_EbtFeePeriodDesc + "(" + sL_EbtFeePeriodCode + ")";
								aEbtDataRow[3] = "負責人姓名:" + sL_EbtCompOwnerName + "<br>" + "電話:" +  sL_EbtContactPhone ;
								aEbtDataRow[4] = "姓名:" + sL_EbtItContactName + "<br>" + "電話:" +  sL_EbtItContactPhone + "<br>" + "手機:" + sL_EbtItContactMobile + "<br>" + "EMail:" + sL_EbtItEMail;
								aEbtDataRow[5] = "裝機地址:" + sL_EbtInstAddr + "<br>" + "戶籍地址:" + sL_EbtCustAddr + "<br>" + "帳單地址:" + sL_EbtBillAddr;					
								aEbtDataRow[6] = "姓名:" + sL_EbtAgentName + "<br>" + "電話:" + sL_EbtAgentPhone+ "<br>" + "地址:" + sL_EbtAgentAddress;

								aEbtDataTable.Rows.Add( aEbtDataRow );

							}

							if ( !aReader2.IsClosed ) aReader2.Close();

						}
						finally
						{
							CloseDbConnEx();
						}

					}
					
				}

				if ( !aReader.IsClosed ) aReader.Close();

				dbgEmc.DataSource = L_EmcDataTable;
				dbgEmc.DataBind();				

				if ( aHasEbtData )
				{
					lblEbtMessage.Text = "CM/CT相關資料";
					dbgEbt.DataSource = aEbtDataTable;
					dbgEbt.DataBind();				
				}
				else
				{
					lblEbtMessage.Text = "沒有CM/CT相關資料";
					dbgEbt.Visible = false;
				}
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
