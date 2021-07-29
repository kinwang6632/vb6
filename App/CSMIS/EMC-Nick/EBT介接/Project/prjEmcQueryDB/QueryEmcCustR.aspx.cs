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
	/// QueryEmcCustR ���K�n�y�z�C
	/// </summary>
	public class QueryEmcCustR : BasePage
	{
		
		private DataTable G_DataTable;
		private int aCurrentPageNumber = 1;		
		private int nG_FirstRec = 0, nG_LastRec = 0;
		private int nG_CurRec = 1;

		protected System.Web.UI.WebControls.Button btnFirstPage;
		protected System.Web.UI.WebControls.Button btnLastPage;
		protected System.Web.UI.WebControls.Label lblPageNum;

		private void ShowPageNum()
		{
			int aTotalRecCount = int.Parse( Session["TOTALRECCOUNT"].ToString() );
			int aRecCountPerPage = int.Parse( ConfigurationSettings.AppSettings["RecCountPerPage"].ToString() );
			int aTotalPageNum = ( aTotalRecCount / aRecCountPerPage );
			if ( aTotalRecCount % aRecCountPerPage > 0 ) ++aTotalPageNum;				
			
			lblPageNum.Text = string.Format( "����:{0}/{1}", aCurrentPageNumber, aTotalPageNum ); 				
			Session["TOTALPAGENUM"] = aTotalPageNum;
		}

		private void WriteQueryLog(string aUserID, string aIp, string aSoName, string aQueryCond, string aQueryKeyValue, 
			int aCount, string aNote)
		{
			if ( ConfigurationSettings.AppSettings[ "IfWriteLogOfQuery" ] == "Y" )
			{
				ConnToDB( "1" );
				try
				{
					oracleCommandForEbtUser.CommandText = string.Format(
						" INSERT INTO EMC_EBT005 ( USERID, FROMIP, QUERYTIME, SONAME, QUERYCOND, QUERYKEYVALUE, RESULTCOUNT, NOTE ) " +
						" VALUES ( '{0}', '{1}', {2}, '{3}', '{4}', '{5}', '{6}', '{7}' ) ", aUserID, aIp, CommFun.getOracleDateTimeStr(),
						aSoName, aQueryCond, aQueryKeyValue, aCount.ToString(), aNote );
					try
					{
						oracleCommandForEbtUser.ExecuteNonQuery();
					}
					catch {}
				}
				finally
				{
					CloseDbConn();
				}
			}
		}

		private void InitDataTable()
		{
			if ( G_DataTable == null ) G_DataTable = new DataTable();				
			if ( G_DataTable.Columns.Count == 0 ) SetDataColumn( G_DataTable );
		}

		private void GetData(string aQueryStr, string aCompName, string aViewName, string aCompCode)
		{

			System.Data.OracleClient.OracleDataReader aDataReader;
			
			DataRow  aDataRow;

			ConnToDB( aCompCode );
			try
			{
				oracleCommandForEbtUser.CommandText = aQueryStr;
				aDataReader = oracleCommandForEbtUser.ExecuteReader();
			
				InitDataTable();

				while ( aDataReader.Read() )
				{
					if ((nG_CurRec>=nG_FirstRec) && (nG_CurRec<=nG_LastRec))
					{
						aDataRow = G_DataTable.NewRow();
						aDataRow[0] = aCompName;
						aDataRow[1] = aDataReader["EMCCOMPCODECUSTID"].ToString().Trim();
						aDataRow[2] = string.Format(
							"<a onclick=javascript:CustIdClick({0},{1}) class=\"Link\">" + 
							"<font color=MediumVioletRed>{2}</font><br></a>",						          
							aDataReader["EMCCUSTID"].ToString().Trim(), 
							aDataReader["EMCCOMPCODE"].ToString().Trim(),  
							aDataReader["EMCCUSTID"].ToString().Trim() );					

						aDataRow[3] = 
							aDataReader["EMCCUSTNAME"].ToString().Trim() + "<br>" + "<font color=MediumVioletRed>" +
							aDataReader["EMCINSTADDRESS"].ToString().Trim() + "</font>";

						aDataRow[4] = aDataReader["EMCINSTTIME"].ToString().Trim();
						aDataRow[5] = aDataReader["EMCPRTIME"].ToString().Trim();

						aDataRow[6] = 
							"�@." + aDataReader["EMCCLASSNAME1"].ToString().Trim() + "<br>" + 
							"�G." + aDataReader["EMCCLASSNAME2"].ToString().Trim() + "<br>" + 
							"�T." + aDataReader["EMCCLASSNAME3"].ToString().Trim();

						aDataRow[7] = aDataReader["EMCPERIOD"].ToString().Trim();

						if ( aDataReader["EMCCUSTSTATUSCODE"].ToString() == "1" )
							aDataRow[8] = string.Format( "<font color=RED>{0}</font>", aDataReader["EMCCUSTSTATUSNAME"].ToString().Trim() );
						else
							aDataRow[8] = aDataReader["EMCCUSTSTATUSNAME"].ToString().Trim();

						aDataRow[9] = 
							"�@." + aDataReader["EMCWIPNAME1"].ToString().Trim() + "<br>" +
							"�G." + aDataReader["EMCWIPNAME2"].ToString().Trim() + "<br>" +
							"�T." + aDataReader["EMCWIPNAME3"].ToString().Trim();					

						aDataRow[10] = 
							"�@." + aDataReader["EMCTEL1"].ToString().Trim() + "<br>" + 
							"�G." + aDataReader["EMCTEL2"].ToString().Trim() + "<br>" + 
							"�T." + aDataReader["EMCTEL3"].ToString().Trim();
                               
						aDataRow[11] = 
							GetEbtServiceTypeCount( ConfigurationSettings.AppSettings["CmKeyWord"],						
							aDataReader["EMCCUSTID"].ToString().Trim(), aViewName, 
							aDataReader["EMCCOMPCODE"].ToString().Trim() );

						aDataRow[12] = 
							GetEbtServiceTypeCount( ConfigurationSettings.AppSettings["CtKeyWord"],						
							aDataReader["EMCCUSTID"].ToString().Trim(), aViewName,
							aDataReader["EMCCOMPCODE"].ToString().Trim() );

						aDataRow[13] = aDataReader["EMCMDUNAME"].ToString().Trim();
						aDataRow[14] = aDataReader["EMCBTNAME"].ToString().Trim();
						aDataRow[15] = aDataReader["EMCPIPELINENAME"].ToString().Trim();
						aDataRow[16] = aDataReader["EMCNODENO"].ToString().Trim();

						G_DataTable.Rows.Add( aDataRow );
					
					}

					//nG_CurRec ++;

					if ( ++nG_CurRec > nG_LastRec ) break;

				}

				if ( !aDataReader.IsClosed ) aDataReader.Close();
			}
			finally
			{
				CloseDbConn();
			}
		}


		private string GetEbtServiceTypeCount(string aServiceType, string aCustID, string aViewName, 
			string aCompCode)
		{
			string aResult = "0";

			if ( CommFun.IsEbtDataArea( aCompCode ) )
			{
				TempCommand.CommandText = string.Format(
					" SELECT COUNT(*) RECCOUNT       " + 
					"   FROM {0}                     " +
					"  WHERE EMCCUSTID = '{1}'       " +
					"    AND EBTSERVICETYPE = '{2}'  " +
					"  GROUP BY EBTSERVICETYPE       ", aViewName, aCustID, aServiceType );
			}
			else
			{
				if ( aServiceType == "CM" )
				{
					TempCommand.CommandText = string.Format(
						" SELECT COUNT(*) RECCOUNT            " + 
						"   FROM SO004                        " +
						"  WHERE CUSTID = '{0}'               " +
						"    AND SERVICETYPE IN ( 'I' )       ", aCustID  );

				}
				else
				{
					aResult = "0";
					return aResult;
				}
				
			}

			ConnToDBEx( aCompCode );
			try
			{
				try
				{
					aResult = TempCommand.ExecuteScalar().ToString();
				}
				catch {}
			}
			finally
			{
				CloseDbConnEx();
			}
			return aResult;
		}


		private string GetWhereConditon(int nI_Key, String sI_KeyValue,String sI_KeyValue2, String sI_KeyValue3, out String sI_QueryCond)
		{
			string sL_Result = ""; 
			
			sI_QueryCond = "";

			switch (nI_Key)
			{
					//�H �Ȥ�W�� �d��
				case 1:
					sL_Result = " EMCCUSTNAME=" + "'" + sI_KeyValue + "' ";
					sI_QueryCond = "�Ȥ�W��=" + sI_KeyValue;
					break;
					//�H �Ȥ�q�� �d��
				case 2:
					sL_Result = " EMCTEL1=" + "'" + sI_KeyValue + "' or EMCTEL2='" + sI_KeyValue + "' or EMCTEL3='" + sI_KeyValue + "'";
					sI_QueryCond = "�Ȥ�q��=" + sI_KeyValue;
					break;
					//�H �a�} �d��(�b�ҽk���)
				case 3:
					if ((sI_KeyValue2=="") && (sI_KeyValue3==""))
					{
						sL_Result = " EMCINSTADDRESS like '" + sI_KeyValue + "%' ";
						sI_QueryCond = "�a�}�H���r��}�Y��=>" + sI_KeyValue;
					}
					else if ((sI_KeyValue2!="") && (sI_KeyValue3==""))
					{
						sL_Result = " EMCINSTADDRESS like '" + sI_KeyValue + "%' or EMCINSTADDRESS like '" + sI_KeyValue2 + "%' ";
						sI_QueryCond = "�a�}�H����Ӧr��}�Y��=>"  + sI_KeyValue + "," + "<font color=green>"+ sI_KeyValue2 + "</font>";
					}
					else if ((sI_KeyValue2=="") && (sI_KeyValue3!=""))
					{
						sL_Result = " EMCINSTADDRESS like '" + sI_KeyValue + "%' or EMCINSTADDRESS like '" + sI_KeyValue3 + "%' ";
						sI_QueryCond = "�a�}�H����Ӧr��}�Y��=>" + sI_KeyValue + "," + "<font color=green>" + sI_KeyValue3 + "</font>";
					}
					else if ((sI_KeyValue2!="") && (sI_KeyValue3!=""))
					{
						sL_Result = " EMCINSTADDRESS like '" + sI_KeyValue + "%' or EMCINSTADDRESS like '" + sI_KeyValue2 + "%' or EMCINSTADDRESS like '" + sI_KeyValue3 + "%' ";
						sI_QueryCond = "�a�}�H���T�Ӧr��}�Y��=>" + sI_KeyValue + "," + "<font color=green>" + sI_KeyValue2 + "</font>" + "," + "<font color=Teal>"+ sI_KeyValue3 + "</font>";
					}				
					break;
					//�H EMC�Ƚs �d��
				case 4:
					sL_Result = " EMCCUSTID=" + "'" + sI_KeyValue + "' ";
					sI_QueryCond = "CATV�Ƚs=" + sI_KeyValue;
					break;
					//�H EBT�Ƚs �d��
				case 5:
					sL_Result = " EBTCUSTID=" + "'" + sI_KeyValue + "' ";
					sI_QueryCond = "CM/CT�Ƚs=" + sI_KeyValue;
					break;
					//�H EBT�X���s�� �d��
				case 6:
					sL_Result = " EBTCONTRACTNO=" + "'" + sI_KeyValue + "' ";
					sI_QueryCond = "CM/CT�X���s��=" + sI_KeyValue;
					break;
					//�H �a�} �d��(���ҽk���)
				case 7:
					sL_Result = " EMCINSTADDRESS like '%" + sI_KeyValue + "%' ";
					sI_QueryCond = "�a�}�]�t���r���=>" + sI_KeyValue;
					break;
			}

			return sL_Result;
		}



		private string GetWhereConditonEx(int aKey, string aKeyValue, string aKeyValue2, string aKeyValue3, out string aQueryCond)
		{
			string aResult = string.Empty;
			
			aQueryCond = string.Empty;

			switch ( aKey )
			{
					//�H �Ȥ�W�� �d��
				case 1:
					aResult = string.Format( " A.CUSTNAME='{0}' ", aKeyValue );
					aQueryCond = string.Format( "�Ȥ�W��={0}", aKeyValue );
					break;
					//�H �Ȥ�q�� �d��
				case 2:
					aResult = string.Format( " ( A.TEL1='{0}' OR A.TEL2='{1}' OR A.TEL3='{2}' ) ", aKeyValue, aKeyValue, aKeyValue );					
					aQueryCond = string.Format( "�Ȥ�q��=", aKeyValue );
					break;
					//�H �a�} �d��(�b�ҽk���)
				case 3:
					if (( aKeyValue2 == string.Empty ) && ( aKeyValue3 == string.Empty ) )
					{
						aResult = string.Format( " A.INSTADDRESS LIKE '{0}%' ", aKeyValue );
						aQueryCond = string.Format( "�a�}�H���r��}�Y��=>", aKeyValue );
					}
					else if ( ( aKeyValue2 != string.Empty ) && ( aKeyValue3 == string.Empty ) )
					{
						aResult = string.Format( " ( A.INSTADDRESS LIKE '{0}%' OR A.INSTADDRESS LIKE '{1}%' ) ", aKeyValue, aKeyValue2 );
						aQueryCond = string.Format( "�a�}�H����Ӧr��}�Y��=>{0},<font color=green>{1}</font>", aKeyValue, aKeyValue2 );
					}
					else if ( ( aKeyValue2 == string.Empty ) && ( aKeyValue3 != string.Empty ) )
					{
						aResult = string.Format( " ( A.INSTADDRESS LIKE '{0}%' OR A.INSTADDRESS LIKE '{1}%' ) ", aKeyValue, aKeyValue3 );
						aQueryCond = string.Format( "�a�}�H����Ӧr��}�Y��=>{0},<font color=green>{1}</font>", aKeyValue, aKeyValue3 );
					}
					else if ( ( aKeyValue2 != string.Empty ) && ( aKeyValue3 != string.Empty ) )
					{
						aResult = string.Format( " ( A.INSTADDRESS LIKE '{0}%' OR A.INSTADDRESS LIKE '{1}%' OR A.INSTADDRESS LIKE '{2}%' ) ", aKeyValue, aKeyValue2, aKeyValue3 );
						aQueryCond = string.Format( "�a�}�H���T�Ӧr��}�Y��=>{0},<font color=green>{1}</font>,<font color=Teal>{2}</font>", aKeyValue, aKeyValue2, aKeyValue3 );
					}				
					break;
					//�H EMC�Ƚs �d��
				case 4:
					aResult = string.Format( " A.CUSTID='{0}' ", aKeyValue );
					aQueryCond =  string.Format( "CATV�Ƚs={0}", aKeyValue );
					break;
					//�H EBT�Ƚs �d��
				case 5:
					aResult = string.Format( " C.EBTCUSTID='{0}' ", aKeyValue );
					aQueryCond = string.Format( "CM/CT�Ƚs={0}", aKeyValue );
					break;
					//�H EBT�X���s�� �d��
				case 6:
					aResult = string.Format( " E.CONTNO='{0}' ", aKeyValue );
					aQueryCond = string.Format( "CM/CT�X���s��={0}", aKeyValue );
					break;
					//�H �a�} �d��(���ҽk���)
				case 7:
					aResult = string.Format( " A.INSTADDRESS LIKE '%{0}%' ", aKeyValue );
					aQueryCond = string.Format( "�a�}�]�t���r���=>{0}", aKeyValue );
					break;
			}

			return aResult;
		}


		private void SetDataColumn(DataTable ADataTable)
		{
			  
			ADataTable.Columns.Add( new DataColumn( "�t�Υx",typeof(string)));			  
			ADataTable.Columns.Add( new DataColumn( "<font color=yellow>�t�Υx�Ƚs</font>", typeof( string ) ) );
			ADataTable.Columns.Add( new DataColumn( "CATV�Ƚs", typeof( string ) ) );
			ADataTable.Columns.Add( new DataColumn( "�Ȥ�W��<br>�˾��a�}",typeof( string ) ) );
			ADataTable.Columns.Add( new DataColumn( "�˾���", typeof( string ) ) );
			ADataTable.Columns.Add( new DataColumn( "�����", typeof( string ) ) );
			ADataTable.Columns.Add( new DataColumn( "�Ȥ����O", typeof( string ) ) );
			ADataTable.Columns.Add( new DataColumn( "ú�O�g��(��)", typeof( string ) ) );
			ADataTable.Columns.Add( new DataColumn( "�Ȥ᪬�A", typeof( string ) ) );           
			ADataTable.Columns.Add( new DataColumn( "���u���A", typeof( string ) ) );
			ADataTable.Columns.Add( new DataColumn( "�q��", typeof( string ) ) );				  
			ADataTable.Columns.Add( new DataColumn( "CM", typeof( string ) ) );
			ADataTable.Columns.Add( new DataColumn( "CT", typeof( string ) ) );
			ADataTable.Columns.Add( new DataColumn( "�j�ӦW��", typeof( string ) ) );
			ADataTable.Columns.Add( new DataColumn( "�ؿv���A", typeof( string ) ) );
			ADataTable.Columns.Add( new DataColumn( "�޽u���O", typeof( string ) ) );
			ADataTable.Columns.Add( new DataColumn( "Fiber Node",typeof( string ) ) );
		}

		
		protected System.Web.UI.WebControls.DataGrid dbgComp1;
		protected System.Web.UI.WebControls.Label lblQueryResult;
		protected System.Web.UI.WebControls.Label lblQueryCond;
		protected System.Web.UI.WebControls.Button btnPreviousPage;
		protected System.Web.UI.WebControls.Button btnNextPage;
		
		public QueryEmcCust SourcePage;

		private void NavigateToPage(object sender, System.EventArgs e)
		{
			try
			{				
				aCurrentPageNumber = System.Convert.ToInt32( Session["CURPAGENUM"].ToString() );
				int aTotalPageNumber = System.Convert.ToInt32( Session["TOTALPAGENUM"].ToString() );
		  	
				string sL_PageInfo = ((Button)sender).CommandName;
		  	
				switch ( sL_PageInfo )
				{
					case "�W�@��":
						aCurrentPageNumber -= 1;
						break;
					case "�U�@��":
						aCurrentPageNumber += 1;				
						break;
					case "�Ĥ@��":
						aCurrentPageNumber = 1;
						break;
					case "�̫�@��":
						aCurrentPageNumber = aTotalPageNumber;
						break;
				}
				
				if ( aCurrentPageNumber <= 0 ) aCurrentPageNumber = 1;
  					
				if ( aTotalPageNumber < aCurrentPageNumber ) return;
  				
				Session["CURPAGENUM"] = aCurrentPageNumber;
  			
				computeRecPos();

				if ( Session["COMPCOUNT"] != null )
				{
					int aCompCount = System.Convert.ToInt32(Session["COMPCOUNT"].ToString());
	  			
					string aQueryStr = string.Empty;
					string aCompName = string.Empty; 
					string aViewName = string.Empty;
					string aCompCode = string.Empty;

					for ( int aIndex = 0; aIndex < aCompCount; aIndex++ )
					{
						aQueryStr = Session["SQL" + aIndex.ToString()].ToString();
						aCompName = Session["COMPNAME" + aIndex.ToString()].ToString();
						aCompCode = Session["COMPCODE" + aIndex.ToString()].ToString();
						aViewName = Session["VIEWNAME" + aIndex.ToString()].ToString();

						if ( ( aQueryStr != string.Empty ) && 
							( aCompName != string.Empty ) )
							//( aViewName != string.Empty ) )
						{
							GetData( aQueryStr, aCompName, aViewName, aCompCode );
						}
					}
  				
					dbgComp1.DataSource = G_DataTable;
					dbgComp1.DataBind();
					ShowPageNum();
				}
				
			}
			catch { }
		}

		private void computeRecPos()
		{
			int nL_RecCountPerPage = System.Convert.ToInt32( ConfigurationSettings.AppSettings["RecCountPerPage"] );
			nG_FirstRec = ( aCurrentPageNumber - 1 )* nL_RecCountPerPage + 1;
			nG_LastRec = ( aCurrentPageNumber ) * nL_RecCountPerPage ;		
		}

		private void Page_Load(object sender, System.EventArgs e)
		{
			// �b�o�̩�m�ϥΪ̵{���X�H��l�ƺ���
			CommFun aCommFun = new CommFun();
			aCommFun.RedirectToLogin( this );			
			
			this.btnPreviousPage.Click += new System.EventHandler(this.NavigateToPage);
			this.btnNextPage.Click += new System.EventHandler(this.NavigateToPage);
			this.btnFirstPage.Click += new System.EventHandler(this.NavigateToPage);
			this.btnLastPage.Click += new System.EventHandler(this.NavigateToPage);

			if (!IsPostBack)
			{

				int aCurrentQueryCount = 0, 
					aPos = 0, 				    
					aTotalQueryCount = 0;				    

				string aDisplayMsg = "", 		       
					aQueryCondition = "", 
					sL_QueryCondForLog = "",
					aSQL = "", 
					aWhere = "", 
					aKeyValue = "", 
					aKeyValue2 = "", 
					aKeyValue3="";

				string sL_Addr1 = "", 
					sL_Addr2 = "", 
					sL_Addr3 = "",
					sL_Addr4 = "",
					sL_Addr41 = "", 
					sL_Addr5 = "",
					sL_Addr51 = "", 
					sL_Addr6 = "", 
					sL_Addr61 = "", 
					sL_Addr7 = "",
					sL_Addr71 = "",  
					sL_Addr8 = "", 
					sL_Addr9 = "", 
					sL_Addr10 = "";

				string sL_QueryKey1="", sL_QueryKey2="", sL_QueryKey3="", sL_QueryKey4="", sL_QueryKey5="", sL_QueryKey6="";
				string sL_QueryKey7="", sL_QueryKey8="", sL_QueryKey10="", sL_QueryKey11="";
				string sL_QueryCond1 = "", sL_QueryCond2="";

				
				HttpCookie L_MyCookie;

				SourcePage = ( QueryEmcCust )Context.Handler;
				 
				string aNotes = "";				
				string sL_UserID = Session["USERID"].ToString();
				string sL_Ip = Request.UserHostAddress;

				//down, write cookie...
				sL_QueryKey1= SourcePage.edtKeyValueRef.Text.Trim();
				sL_QueryKey2= SourcePage.edtAddrField1Ref.Text.Trim();
				sL_QueryKey3= SourcePage.edtAddrField2Ref.Text.Trim();
				sL_QueryKey4= SourcePage.edtAddrField3Ref.Text.Trim();
				sL_QueryKey5= SourcePage.edtAddrField4Ref.Text.Trim();
				sL_QueryKey6= SourcePage.edtAddrField5Ref.Text.Trim();
				sL_QueryKey7= SourcePage.edtAddrField6Ref.Text.Trim();
				sL_QueryKey8= SourcePage.edtAddrField7Ref.Text.Trim();

				sL_QueryKey10= SourcePage.edtAddrField8Ref.Text.Trim();
				sL_QueryKey11= SourcePage.edtAddrField9Ref.Text.Trim();
			
			
				sL_QueryCond1= SourcePage.rdgQueryCondRef.SelectedIndex.ToString();

				for( int AIndex = 0; AIndex < SourcePage.chbCompListRef.Items.Count; AIndex++ )
				{
					if ( SourcePage.chbCompListRef.Items[AIndex].Selected ) 
						sL_QueryCond2 += "1";
					else
						sL_QueryCond2 += "0";
				}		 

				L_MyCookie = new HttpCookie("USERINFO");
				L_MyCookie.Values.Add("QK1",sL_QueryKey1);
				L_MyCookie.Values.Add("QK2",sL_QueryKey2);
				L_MyCookie.Values.Add("QK3",sL_QueryKey3);
				L_MyCookie.Values.Add("QK4",sL_QueryKey4);
				L_MyCookie.Values.Add("QK5",sL_QueryKey5);
				L_MyCookie.Values.Add("QK6",sL_QueryKey6);
				L_MyCookie.Values.Add("QK7",sL_QueryKey7);
				L_MyCookie.Values.Add("QK8",sL_QueryKey8);
				
				L_MyCookie.Values.Add("QK10",sL_QueryKey10);
				L_MyCookie.Values.Add("QK11",sL_QueryKey11);
				L_MyCookie.Values.Add("QC1",sL_QueryCond1);
				L_MyCookie.Values.Add("QC2",sL_QueryCond2);
				Response.AppendCookie(L_MyCookie);
				//up, write cookie...
	
				sL_QueryCondForLog = SourcePage.rdgQueryCondRef.SelectedItem.Text;
				aQueryCondition = SourcePage.rdgQueryCondRef.SelectedItem.Text;
				if (aQueryCondition=="�˾��a�}")
				{
					sL_Addr1 = SourcePage.edtAddrField1Ref.Text.Trim();

					sL_Addr2 = SourcePage.edtAddrField2Ref.Text.Trim();
					if (sL_Addr2!="")
						sL_Addr2 = sL_Addr2 + "�F";

					sL_Addr3 = SourcePage.edtAddrField3Ref.Text.Trim();
					if (sL_Addr3!="")
						sL_Addr3 = sL_Addr3 + "�q";

					sL_Addr4 = SourcePage.edtAddrField4Ref.Text.Trim();
					sL_Addr41 = SourcePage.edtAddrField41Ref.Text.Trim();
					sL_Addr5 = SourcePage.edtAddrField5Ref.Text.Trim();
					sL_Addr51 = SourcePage.edtAddrField51Ref.Text.Trim();
					sL_Addr6 = SourcePage.edtAddrField6Ref.Text.Trim();
					sL_Addr61 = SourcePage.edtAddrField61Ref.Text.Trim();

					sL_Addr7 = SourcePage.edtAddrField7Ref.Text.Trim();
					sL_Addr71 = SourcePage.edtAddrField71Ref.Text.Trim();
					

					sL_Addr8 = SourcePage.edtAddrField8Ref.Text.Trim(); //�m��
					sL_Addr9 = SourcePage.edtAddrField9Ref.Text.Trim(); //����
					sL_Addr10 = SourcePage.edtAddrField10Ref.Text.Trim(); //��

					aKeyValue = sL_Addr1 + sL_Addr2 + sL_Addr3 + sL_Addr41 + sL_Addr51 + sL_Addr61 + sL_Addr71;

					if ((sL_Addr8!="")&&(sL_Addr9==""))
						aKeyValue2 = sL_Addr8 + aKeyValue;
					else if ((sL_Addr8=="")&&(sL_Addr9!=""))
						aKeyValue2 = sL_Addr9 + aKeyValue;
					else if ((sL_Addr8!="") && (sL_Addr9!=""))
					{
						aKeyValue3 = sL_Addr9 + sL_Addr8 + aKeyValue;
						aKeyValue2 = sL_Addr8 + aKeyValue;
					}

					if (sL_Addr10!="")
					{
						if (aKeyValue!="")
							aKeyValue = aKeyValue + sL_Addr10;
						if (aKeyValue2!="")
							aKeyValue2 = aKeyValue2 + sL_Addr10;
						if (aKeyValue3!="")
							aKeyValue3 = aKeyValue3 + sL_Addr10;
					}					
				}
				else
					aKeyValue = SourcePage.edtKeyValueRef.Text.Trim();

				InitDataTable();
				aCurrentPageNumber = 1;
				Session["CURPAGENUM"] = aCurrentPageNumber;
				computeRecPos();

				string aViewName = string.Empty;
				string aCompName = string.Empty;
				string aServiceType = string.Empty;
				byte aCompCode = 0;
				bool aShowData = false;
				int aArrayPos = 0;

				for( int aIndex = 0; aIndex < SourcePage.chbCompListRef.Items.Count; aIndex++ )
				{
					if ( !SourcePage.chbCompListRef.Items[aIndex].Selected ) continue;
					
					aPos++;						
					aCompCode = Byte.Parse( SourcePage.chbCompListRef.Items[aIndex].Value );
					aCompName = CommFun.GetCompChineseName( aCompCode );
					aViewName = string.Empty;

					if ( CommFun.IsEbtDataArea( aCompCode.ToString() ) )
					{
						aViewName = CommFun.GetQueryViewName( aCompCode, 1 );
						aWhere = " WHERE " + GetWhereConditon( int.Parse( SourcePage.rdgQueryCondRef.SelectedValue ),  
							aKeyValue, aKeyValue2, aKeyValue3, out aQueryCondition );
						aSQL = string.Format( "SELECT COUNT( DISTINCT( EMCCUSTID ) ) RECCOUNT FROM {0} {1} ",
							aViewName, aWhere );
					}
					else
					{
						aWhere = " AND " + GetWhereConditonEx( int.Parse( SourcePage.rdgQueryCondRef.SelectedValue ),  
							aKeyValue, aKeyValue2, aKeyValue3, out aQueryCondition );

						aServiceType = GetServiceTypeByCondition( int.Parse( 
							SourcePage.rdgQueryCondRef.SelectedValue ) );

						aSQL = string.Format( 
							" SELECT COUNT( DISTINCT( A.CUSTID ) ) RECCOUNT            " +
							"   FROM SO001 A , SO002 B, SO152 C, SO014 D, SO003 E      " + 
							"  WHERE A.CUSTID = B.CUSTID                               " + 
							"    AND A.COMPCODE = B.COMPCODE                           " + 
							"    AND B.SERVICETYPE = 'C'                               " + 
							"    AND A.COMPCODE = D.COMPCODE                           " + 
							"    AND A.INSTADDRNO = D.ADDRNO                           " + 
							" 	  AND A.CUSTID = E.CUSTID(+)                            " +
							"    AND E.SERVICETYPE(+) = {0}									  " +
							"    AND A.COMPCODE = C.EMCCOMPCODE(+)                     " + 
							"    AND A.CUSTID = C.EMCCUSTID(+)                         ", aServiceType ) + aWhere;
					}

					aCurrentQueryCount = GetQueryRecordCount( aSQL, aCompCode.ToString() );
					aTotalQueryCount += aCurrentQueryCount;

					if ( aCurrentQueryCount <= 0 )
					{
						/* �d�L���� */
						aDisplayMsg = ( aDisplayMsg == string.Empty ) ? 
							string.Format( "{0}: 0 ��", aCompName ) :
							string.Format( "{0}; {1}: 0 ��", aDisplayMsg, aCompName );
					}						
					else if ( aCurrentQueryCount > int.Parse( ConfigurationSettings.AppSettings["MaxQueryRecordCount"] ) )								
					{
						/* ���ƶW�L���� */
						aNotes = "�Ҭd�X����Ƶ��ƤӦh,�t�Τ���ܸ��";
						aDisplayMsg = ( aDisplayMsg == string.Empty ) ? 
							string.Format( "{0}�Ҭd�X����Ƶ��ƤӦh,�G�L�k��ܸ��. �Э��s�d��.", aCompName ) :
							string.Format( "{0}; {1}�Ҭd�X����Ƶ��ƤӦh,�G�L�k��ܸ��. �Э��s�d��.", aDisplayMsg, aCompName );
					}
					else
					{
						aShowData = true;
						aDisplayMsg = ( aDisplayMsg == string.Empty ) ? 
							string.Format( "{0}:{1}��", aCompName, aCurrentQueryCount.ToString() ) : 
							string.Format( "{0}; {1}:{2}��", aDisplayMsg, aCompName, aCurrentQueryCount.ToString() );

						if ( CommFun.IsEbtDataArea( aCompCode.ToString() ) )
						{
							/* EBT ��ƥ洫�� */
							aSQL = string.Format(
								" SELECT DISTINCT CONCAT(TRIM(TO_CHAR(EMCCOMPCODE,'000')),        " +
								"        TRIM(TO_CHAR(EMCCUSTID,'00000000'))) EMCCOMPCODECUSTID,  " +
								"        EMCCOMPCODE,                                             " +
								"        EMCCUSTID,                                               " +
								"        EMCCUSTNAME,                                             " +
								"        EMCPERIOD,                                               " +
								"        EMCCUSTSTATUSCODE,                                       " +
								"        EMCCUSTSTATUSNAME,                                       " +
								"        EMCWIPNAME1,                                             " +
								"        EMCWIPNAME2,                                             " +
								"        EMCWIPNAME3,                                             " +
								"        TO_CHAR( EMCPRTIME,'YYYY/MM/DD' ) EMCPRTIME,             " +
								"        EMCTEL1,                                                 " +
								"        EMCTEL2,                                                 " +
								"        EMCTEL3,                                                 " +
								"        EMCINSTADDRESS,                                          " +
								"        EMCCLASSNAME1,                                           " +
								"        EMCCLASSNAME2,                                           " +
								"        EMCCLASSNAME3,                                           " +
								"        TO_CHAR( EMCINSTTIME, 'YYYY/MM/DD' ) EMCINSTTIME,        " +
								"        EMCMDUNAME,                                              " +
								"        EMCBTNAME,                                               " +
								"        EMCPIPELINENAME,                                         " +
								"        EMCNODENO                                                " +
								"   FROM {0} {1}                                                  " +
								"  ORDER BY EMCCUSTSTATUSCODE                                     ",
								aViewName, aWhere );

							GetData( aSQL, aCompName, aViewName, aCompCode.ToString() );

						}
						else
						{
							/* ������ư� */
							aSQL = string.Format(
								" SELECT DISTINCT CONCAT(TRIM(TO_CHAR(A.COMPCODE,'000')),         " +
								"        TRIM(TO_CHAR(A.CUSTID,'00000000'))) EMCCOMPCODECUSTID,   " +
								"        A.COMPCODE AS EMCCOMPCODE,                               " +
								"        A.CUSTID AS EMCCUSTID,                                   " +
								"        A.CUSTNAME AS EMCCUSTNAME,                               " +
								"        E.PERIOD AS EMCPERIOD,                                   " +
								"        B.CUSTSTATUSCODE AS EMCCUSTSTATUSCODE,                   " +
								"        B.CUSTSTATUSNAME AS EMCCUSTSTATUSNAME,                   " +
								"        B.WIPNAME1 AS EMCWIPNAME1,                               " +
								"        B.WIPNAME2 AS EMCWIPNAME2,                               " +
								"        B.WIPNAME3 AS EMCWIPNAME3,                               " +
								"        TO_CHAR( B.PRTIME,'YYYY/MM/DD' ) AS EMCPRTIME,           " +
								"        A.TEL1 AS EMCTEL1,                                       " +
								"        A.TEL2 AS EMCTEL2,                                       " +
								"        A.TEL3 AS EMCTEL3,                                       " +
								"        A.INSTADDRESS AS EMCINSTADDRESS,                         " +
								"        A.CLASSNAME1 AS EMCCLASSNAME1,                           " +
								"        A.CLASSNAME2 AS EMCCLASSNAME2,                           " +
								"        A.CLASSNAME3 AS EMCCLASSNAME3,                           " +
								"        TO_CHAR( B.INSTTIME , 'YYYY/MM/DD' ) AS EMCINSTTIME,     " +
								"        D.MDUNAME AS EMCMDUNAME,                                 " +
								"        D.BTNAME AS EMCBTNAME,                                   " +
								"        D.PIPELINENAME AS EMCPIPELINENAME,                       " +
								"        D.NODENO AS EMCNODENO                                    " +
								"   FROM SO001 A , SO002 B, SO152 C , SO014 D, SO003 E            " + 
							  "  WHERE A.CUSTID = B.CUSTID                                       " + 
							  "    AND A.COMPCODE = B.COMPCODE                                   " + 
							  "    AND B.SERVICETYPE = 'C'                                       " + 
							  "    AND A.COMPCODE = D.COMPCODE                                   " + 
							  "    AND A.INSTADDRNO = D.ADDRNO                                   " + 
							  "    AND A.CUSTID = E.CUSTID(+)                                    " + 
							  "    AND E.SERVICETYPE(+) = {0}                                    " +  
							  "    AND A.COMPCODE = C.EMCCOMPCODE(+)                             " + 
							  "    AND A.CUSTID = C.EMCCUSTID(+)                                 ", aServiceType );

							aSQL = ( aSQL + aWhere + "  ORDER BY B.CUSTSTATUSCODE " );

							GetData( aSQL, aCompName, aViewName, aCompCode.ToString() );
  
						}

						Session["SQL" + aArrayPos.ToString()] = aSQL;
						Session["COMPNAME" + aArrayPos.ToString()] = aCompName;
						Session["VIEWNAME" + aArrayPos.ToString()] = aViewName;
						Session["COMPCODE" + aArrayPos.ToString()] = aCompCode.ToString();
						aArrayPos++;

					}

						
//						aViewName = CommFun.GetQueryViewName( aCompCode, 1 );
//
//						if ( aViewName != string.Empty )
//						{
//							/* EBT ��ƥ洫�� */
//
//							aWhere = " WHERE " + getWhereCond( int.Parse(
//								SourcePage.rdgQueryCondRef.SelectedValue ),  aKeyValue, aKeyValue2, aKeyValue3, 
//								out aQueryCondition );
//
//							aSQL = string.Format( "SELECT COUNT( DISTINCT( EMCCUSTID ) ) RECCOUNT FROM {0} {1} ",
//								aViewName, aWhere );						
//
//							ConnToDB( aCompCode.ToString() );
//							try
//							{
//								oracleCommandForEbtUser.CommandText = aSQL;
//								aReader = oracleCommandForEbtUser.ExecuteReader();						
//								aReader.Read();
//								aCurrentQueryCount = System.Convert.ToInt32(aReader["RECCOUNT"].ToString());
//								aTotalQueryCount += aCurrentQueryCount;
//								aReader.Close();
//							}
//							finally
//							{
//								CloseDbConn();
//							}						
//							
//							if ( aCurrentQueryCount <= 0 )
//							{
//								if ( aDisplayMsg == string.Empty )
//									aDisplayMsg = string.Format( "{0}: 0 ��", aCompName );
//								else
//									aDisplayMsg = string.Format( "{0}; {1}: 0 ��", aDisplayMsg, aCompName );
//							}						
//							else if ( aCurrentQueryCount > int.Parse( ConfigurationSettings.AppSettings["MaxQueryRecordCount"] ) )								
//							{
//								sL_Note = "�Ҭd�X����Ƶ��ƤӦh,�t�Τ���ܸ��";
//								if ( aDisplayMsg == string.Empty )
//									aDisplayMsg = string.Format( "{0}�Ҭd�X����Ƶ��ƤӦh,�G�L�k��ܸ��. �Э��s�d��.", aCompName );
//								else
//									aDisplayMsg = string.Format( "{0}; {1}�Ҭd�X����Ƶ��ƤӦh,�G�L�k��ܸ��. �Э��s�d��.", aDisplayMsg, aCompName );
//							}
//							else
//							{
//								bL_ShowDataGrid = true;
//								if (aDisplayMsg=="")
//									aDisplayMsg = aCompName + ": " + aCurrentQueryCount.ToString() +  " ��";
//								else
//									aDisplayMsg = aDisplayMsg + "; " + aCompName + ": " + aCurrentQueryCount.ToString() +  " ��";
//
//								aSQL = 
//									" SELECT DISTINCT CONCAT(TRIM(TO_CHAR(EMCCOMPCODE,'000')),        " +
//									"        TRIM(TO_CHAR(EMCCUSTID,'00000000'))) EMCCOMPCODECUSTID,  " +
//									"        EMCCOMPCODE,                                             " +
//									"        EMCCUSTID,                                               " +
//									"        EMCCUSTNAME,                                             " +
//									"        EMCPERIOD,                                               " +
//									"        EMCCUSTSTATUSCODE,                                       " +
//									"        EMCCUSTSTATUSNAME,                                       " +
//									"        EMCWIPNAME1,                                             " +
//									"        EMCWIPNAME2,                                             " +
//									"        EMCWIPNAME3,                                             " +
//									"        TO_CHAR( EMCPRTIME,'YYYY/MM/DD' ) EMCPRTIME,             " +
//									"        EMCTEL1,                                                 " +
//									"        EMCTEL2,                                                 " +
//									"        EMCTEL3,                                                 " +
//									"        EMCINSTADDRESS,                                          " +
//									"        EMCCLASSNAME1,                                           " +
//									"        EMCCLASSNAME2,                                           " +
//									"        EMCCLASSNAME3,                                           " +
//									"        TO_CHAR( EMCINSTTIME, 'YYYY/MM/DD' ) EMCINSTTIME,        " +
//									"        EMCMDUNAME,                                              " +
//									"        EMCBTNAME,                                               " +
//									"        EMCPIPELINENAME,                                         " +
//									"        EMCNODENO                                                " +
//									"   FROM  " + aViewName + aWhere + 
//									"  ORDER BY EMCCUSTSTATUSCODE                                     ";
//
//
//								Session["SQL" + nL_ArrayPos.ToString()] = aSQL;
//								Session["COMPNAME" + nL_ArrayPos.ToString()] = aCompName;
//								Session["VIEWNAME" + nL_ArrayPos.ToString()] = aViewName;
//								Session["COMPCODE" + nL_ArrayPos.ToString()] = aCompCode.ToString();
//								nL_ArrayPos++;
//
//								getData(aSQL, aCompName, aViewName, aCompCode.ToString() );
//
//							}
//								
//						}
//						else 
//						{
//							/* ������ư� */
//
//							aWhere = " WHERE " + getWhereCond( int.Parse(
//								SourcePage.rdgQueryCondRef.SelectedValue ),  aKeyValue, aKeyValue2, aKeyValue3, 
//								out aQueryCondition );
//
//							aSQL = 
//								" SELECT COUNT( DISTINCT( A.CUSTID ) ) RECCOUNT FROM       " +
//							   "   FROM SO001 A , SO002 B, SO152 C , SO014 D, SO003 E     " + 
//								"  WHERE A.CUSTID = B.CUSTID                               " + 
//								"    AND A.COMPCODE = B.COMPCODE                           " + 
//								"    AND B.SERVICETYPE = 'C'                               " + 
//								"    AND A.COMPCODE = D.COMPCODE                           " + 
//								"    AND A.INSTADDRNO = D.ADDRNO                           " + 
//								" 	  AND A.CUSTID = E.CUSTID(+)                            " + 
//								" 	  AND A.COMPCODE = C.EMCCOMPCODE(+)                     " + 
//								"    AND A.CUSTID = C.EMCCUSTID(+)                         ";
//
//							ConnToDB( aCompCode.ToString() );
//							try
//							{
//								oracleCommandForEbtUser.CommandText = aSQL;
//								aReader = oracleCommandForEbtUser.ExecuteReader();						
//								aReader.Read();
//								aCurrentQueryCount = System.Convert.ToInt32(aReader["RECCOUNT"].ToString());
//								aTotalQueryCount = aTotalQueryCount + aCurrentQueryCount;
//								aReader.Close();
//							}
//							finally
//							{
//								CloseDbConn();
//							}						
//							
//
//
//						}


						string aLogKeyValue = string.Empty;

						aLogKeyValue = ( aKeyValue3.Length >= aKeyValue2.Length ) ? aKeyValue3 : aKeyValue2;
						aLogKeyValue = ( aKeyValue.Length > aLogKeyValue.Length ) ? aKeyValue : aLogKeyValue;

						WriteQueryLog( sL_UserID, sL_Ip, aCompName, sL_QueryCondForLog, aLogKeyValue,
							aCurrentQueryCount, aNotes );
					
				}

				lblQueryCond.Text = "�d�߱���: " + aQueryCondition;				
				Session["COMPCOUNT"] = aArrayPos.ToString();
				Session["TOTALRECCOUNT"] = aTotalQueryCount;
				ShowPageNum();
				lblQueryResult.Text = aDisplayMsg;				
				DoSetDataGridVisible( aShowData );

			}
		}


		private int GetQueryRecordCount(string aSQL, string aCompCode)
		{
			int aResult = 0;
			if ( CommFun.IsEbtDataArea( aCompCode ) )
			{
				ConnToDB( aCompCode );
				try
				{
					try
					{
						oracleCommandForEbtUser.CommandText = aSQL;
						aResult = int.Parse( oracleCommandForEbtUser.ExecuteOracleScalar().ToString() );
					}
					catch	{}
				}
				finally
				{
					CloseDbConn();
				}
			}
			else
			{
				ConnToDBEx( aCompCode );
				try
				{
					try
					{
						TempCommand.CommandText = aSQL;
						aResult = int.Parse( TempCommand.ExecuteOracleScalar().ToString() );
					}
					catch	{}
				}
				finally
				{
					CloseDbConnEx();
				}
			}
			return aResult;
		}


		private void DoSetDataGridVisible(bool aShow)
		{
			dbgComp1.Visible = aShow;
			btnFirstPage.Visible = aShow;
			btnPreviousPage.Visible = aShow;
			btnNextPage.Visible = aShow;
			btnLastPage.Visible = aShow;
			lblPageNum.Visible = aShow;

			if ( aShow )
			{
				dbgComp1.DataSource = G_DataTable;
				dbgComp1.DataBind();
			}

		}


		private string GetServiceTypeByCondition(int aCodition)
		{
			string aResult = " 'C' ";
			if ( aCodition == 6 )
			{
				aResult = " 'I' ";
			}
			return aResult;
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
