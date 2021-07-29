// Connection Pool    <~~ OK
// 讀取 MDB 設定 <~~ OK
// MDB 結構訂定  <~~ OK

//還回 Connection Pool, Connection 是否要 Dispose <~~ 不用
//測試 InitialSetup 是否有機會執行第二次 <~~ 不會
//MDB 值更改 是否應該要重啟 <~~ 因為 InitialSetup 不會執行第二次...所以要重啟
// Stress Test

using System;
using System.Web;
using System.Web.Services;
using System.Web.Services.Protocols;
using System.Collections;
using System.Collections.Generic;
using Oracle.DataAccess.Client;
using System.Data;
using System.Data.OleDb;

#region ParmElement、ParmStruct、RetStruct
public class ParmElement
{

    private string parmNameField;

    private string parmValueField;

    /// <remarks/>
    public string ParmName
    {
        get
        {
            return this.parmNameField;
        }
        set
        {
            this.parmNameField = value;
        }
    }

    /// <remarks/>
    public string ParmValue
    {
        get
        {
            return this.parmValueField;
        }
        set
        {
            this.parmValueField = value;
        }
    }
}

//public class ParmStruct
//{

//    private ParmElement[] parmAryField;

//    /// <remarks/>
//    public ParmElement[] ParmAry
//    {
//        get
//        {
//            return this.parmAryField;
//        }
//        set
//        {
//            this.parmAryField = value;
//        }
//    }
//}

public class RetStruct
{

    private int retCodeField;

    private string retMsgField;

    private ParmElement[] parmAryField;

    /// <remarks/>
    [System.Xml.Serialization.SoapElementAttribute(DataType = "integer")]
    public int RetCode
    {
        get
        {
            return this.retCodeField;
        }
        set
        {
            this.retCodeField = value;
        }
    }

    /// <remarks/>
    public string RetMsg
    {
        get
        {
            return this.retMsgField;
        }
        set
        {
            this.retMsgField = value;
        }
    }

    /// <remarks/>
    public ParmElement[] ParmAry
    {
        get
        {
            return this.parmAryField;
        }
        set
        {
            this.parmAryField = value;
        }
    }
}
#endregion

[WebService(Namespace = "CRM")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
public class CRM : System.Web.Services.WebService
{
    #region Declarations
    private static object SetupLockObj = new object();
    private static string SetupConnectionString = null;
    private static Hashtable hashOracleConnectionString = null;
    #endregion

    #region LoadSetup
    private bool InitialSetup()
    {
        lock (CRM.SetupLockObj)
        {
            if (CRM.SetupConnectionString != null) return true; //表示已經初始化

            CRM.SetupConnectionString =
                string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Persist Security Info=False;User ID={1};Jet OLEDB:Database Password={2}",
                    new object[] { Server.MapPath("~/App_Data/Setup.mdb"), "Admin", "cyc84177282" });

            OleDbConnection Conn = null;
            OleDbDataAdapter DA = null;
            DataSet DS = null;
            DataView DV = null;
            int MinPoolSize = -1;
            int MaxPoolSize = -1;
            int ConnectionLifetime = -1;
            string CompanyCode = null;
            string CompanyName = null;
            string ServiceName = null;
            string Account = null;
            string Password = null;

            hashOracleConnectionString = new Hashtable();
            string Template_OracleConnString =
                "Data Source={0};Persist Security Info=True;User ID={1};Password={2};Min Pool Size={3};Max Pool Size={4};Connection Lifetime={5}";

            try
            {
                Conn = new OleDbConnection(CRM.SetupConnectionString);
                Conn.Open();

                DA = new OleDbDataAdapter("SELECT * FROM [COMMONSETUP]", Conn);
                DS = new System.Data.DataSet();
                DA.FillSchema(DS, System.Data.SchemaType.Mapped);
                DA.Fill(DS, "COMMONSETUP");

                DV = DS.Tables["COMMONSETUP"].DefaultView;
                if (DV.Count != 0)
                {
                    MinPoolSize = System.Convert.ToInt32(DV[0]["MinPoolSize"]);
                    MaxPoolSize = System.Convert.ToInt32(DV[0]["MaxPoolSize"]);
                    ConnectionLifetime = System.Convert.ToInt32(DV[0]["ConnectionLifetime"]);
                }
                else if (DV.Count == 0) return false;
                
                DA.SelectCommand.CommandText = "SELECT * FROM [CONNECTION]";
                DS = new System.Data.DataSet();
                DA.FillSchema(DS, System.Data.SchemaType.Mapped);
                DA.Fill(DS, "CONNECTION");

                DV = DS.Tables["CONNECTION"].DefaultView;

                for (int i = 0; i < DV.Count; i++)
                {
                    CompanyCode = System.Convert.ToString(DV[i]["COMPANYCODE"]);
                    CompanyName = System.Convert.ToString(DV[i]["COMPANYNAME"]);
                    ServiceName = System.Convert.ToString(DV[i]["SERVICENAME"]);
                    Account = System.Convert.ToString(DV[i]["ACCOUNT"]);
                    Password = System.Convert.ToString(DV[i]["PASSWORD"]);

                    CRM.hashOracleConnectionString[CompanyCode] = string.Format(Template_OracleConnString,
                        new object[] { ServiceName, Account, Password, MinPoolSize, MaxPoolSize, ConnectionLifetime });
                }

                if (DV.Count == 0) return false;
                else return true;
            }
            catch //(System.Exception ex)
            {
                return false;
            }
            finally
            {
                try
                {
                    Conn.Close();
                }
                catch { }
            }
        }
    }
    #endregion

    //(參數前有*為必要參數)
    //	QueryCrmCustInfo: 取得CRM客戶資訊
    //	使用時機: 欲取得CRM系統中特定客戶資訊
    //	呼叫參數: *CrmId, *CrmCustId
    //	回傳參數: *CrmId, *CrmCustId, *CrmCustName, *CrmCustTel, *CrmCustAddress
    [WebMethod]
    public RetStruct QueryCrmCustInfo(ParmElement[] ParmAry)
    {
        ////Changed
        RetStruct RetObj = new RetStruct();

        ////Changed
        if (CRM.SetupConnectionString == null)
        {
            if (!InitialSetup())
            {
                RetObj.RetCode = 176; //0xB0
                RetObj.RetMsg = "初始環境失敗";
                return RetObj;
            }
        }

        Dictionary<string, string> DnryParms = null;
        if (ParmAry != null)
        {
            DnryParms = new Dictionary<string, string>();

            for (int i = 0; i < ParmAry.Length; i++)
            {
                DnryParms[ParmAry[i].ParmName] = ParmAry[i].ParmValue;
            }
        }

        ParmElement ParmUnit;
        ArrayList arylstParmElement = new ArrayList(); //Changed
        
        if (DnryParms == null)
        {
            //Changed
            RetObj.RetCode = 129; //0x81
            RetObj.RetMsg = "沒有任何參數";
            return RetObj;
        }
        else
        {
            //check necessary parameters
            String RetMsg = "缺少必要參數 ";
            if (DnryParms["CrmId"] == null)
            {
                RetMsg += "CrmId, ";
            }

            if (DnryParms["CrmCustId"] == null)
            {
                RetMsg += "CrmCustId, ";
            }

            if (!RetMsg.Equals("缺少必要參數 "))
            {
                RetObj.RetCode = 129; //0x81
                RetObj.RetMsg = RetMsg.Substring(0, RetMsg.Length - 2);
                return RetObj;
            }

            //Get Connection from Connection Pool
            string ConnectionString = (string)hashOracleConnectionString[DnryParms["CrmId"]];

            if (ConnectionString == null)
            {
                RetObj.RetCode = 130; //0x82
                RetObj.RetMsg = "參數值 CrmId 不合法";
                return RetObj;
            }

            DB_Oracle Conn = new DB_Oracle(ConnectionString); //Get Connection from Connection Pool

            //Test Connection
            if (!Conn.TestConnection())
            {
                RetObj.RetCode = 160; //0xA0
                RetObj.RetMsg = "資料庫連結失敗";
                return RetObj;
            }

            //Do Select
            DataTable DT = Conn.ExecuteQuery("select CUSTNAME,TEL1,INSTADDRESS from so001 where custid = " + DnryParms["CrmCustId"]);
            if (DT == null)
            {
                RetObj.RetCode = 161; //0xA1
                RetObj.RetMsg = "資料庫查詢指令失敗";
                return RetObj;
            }

            //Save Result
            DataView DV = DT.DefaultView;
            
            if (DV.Count != 0)
            {
                ParmUnit = new ParmElement();
                ParmUnit.ParmName = "CrmId";
                ParmUnit.ParmValue = DnryParms["CrmId"];
                arylstParmElement.Add(ParmUnit);

                ParmUnit = new ParmElement();
                ParmUnit.ParmName = "CrmCustId";
                ParmUnit.ParmValue = DnryParms["CrmCustId"];
                arylstParmElement.Add(ParmUnit);

                ParmUnit = new ParmElement();
                ParmUnit.ParmName = "CrmCustName";
                ParmUnit.ParmValue = System.Convert.ToString(DV[0]["CUSTNAME"]);
                arylstParmElement.Add(ParmUnit);

                ParmUnit = new ParmElement();
                ParmUnit.ParmName = "CrmCustTel";
                ParmUnit.ParmValue = System.Convert.ToString(DV[0]["TEL1"]);
                arylstParmElement.Add(ParmUnit);

                ParmUnit = new ParmElement();
                ParmUnit.ParmName = "CrmCustAddress";
                ParmUnit.ParmValue = System.Convert.ToString(DV[0]["INSTADDRESS"]);
                arylstParmElement.Add(ParmUnit);
            }
            else
            {
                RetObj.RetCode = 176; //0xB0
                RetObj.RetMsg = "CRM_GetCrmCust error";
                return RetObj;
            }

            //Apply Result to RetObj
            RetObj.ParmAry = (ParmElement[])arylstParmElement.ToArray(typeof(ParmElement)); //Changed
            RetObj.RetCode = 0;
            RetObj.RetMsg = "OK";
        }

        return RetObj;
    }

    //	QueryServiceByRelationNo: 依CRM關聯單資訊取得服務內容
    //	使用時機: SCM提供的裝機界面接受第一線工程人員輸入相關資訊時, 將資訊透過此功能詢問CRM確認是否為正確資訊, 並提供裝機服務或換機服務資訊給SCM.
    //	呼叫參數: *ScmServiceType, *CrmId, *CrmCustId, *CrmRelationNo, *CrmOperatorId.
    //	回傳參數: 
    //	若為建立新服務(裝機): *CrmId, *CrmCustId, *CrmCustName, *CrmCustTel, *CrmCustAddress, UserIdNo, UserName, UserBirthday, *SchemeCode1, SchemeCode[2..9]
    //	若為換機服務(維修/換機): *ScmAccountId, SchemeCode[1..9]
    [WebMethod]
    public RetStruct QueryServiceByRelationNo(ParmElement[] ParmAry)
    {
        ////Changed
        RetStruct RetObj = new RetStruct();

        ////Changed
        if (CRM.SetupConnectionString == null)
        {
            if (!InitialSetup())
            {
                RetObj.RetCode = 176; //0xB0
                RetObj.RetMsg = "初始環境失敗";
                return RetObj;
            }
        }

        Dictionary<string, string> DnryParms = null;
        if (ParmAry != null)
        {
            DnryParms = new Dictionary<string, string>();

            for (int i = 0; i < ParmAry.Length; i++)
            {
                DnryParms[ParmAry[i].ParmName] = ParmAry[i].ParmValue;
            }
        }

        ParmElement ParmUnit;
        ArrayList arylstParmElement = new ArrayList(); //Changed

        if (DnryParms == null)
        {
            //Changed
            RetObj.RetCode = 129; //0x81
            RetObj.RetMsg = "沒有任何參數";
            return RetObj;
        }
        else
        {
            //check necessary parameters
            String RetMsg = "缺少必要參數 ";
            if (DnryParms["CrmId"] == null)
            {
                RetMsg += "CrmId, ";
            }

            if (DnryParms["CrmCustId"] == null)
            {
                RetMsg += "CrmCustId, ";
            }

            if (DnryParms["ScmServiceType"] == null)
            {
                RetMsg += "ScmServiceType, ";
            }

            if (DnryParms["CrmRelationNo"] == null)
            {
                RetMsg += "CrmRelationNo, ";
            }

            if (DnryParms["CrmOperatorId"] == null)
            {
                RetMsg += "CrmOperatorId, ";
            }            

            if (!RetMsg.Equals("缺少必要參數 "))
            {
                RetObj.RetCode = 129; //0x81
                RetObj.RetMsg = RetMsg.Substring(0, RetMsg.Length - 2);
                return RetObj;
            }

            //Get Connection from Connection Pool
            string ConnectionString = (string)hashOracleConnectionString[DnryParms["CrmId"]];

            if (ConnectionString == null)
            {
                RetObj.RetCode = 130; //0x82
                RetObj.RetMsg = "參數值 CrmId 不合法";
                return RetObj;
            }

            DB_Oracle Conn = new DB_Oracle(ConnectionString); //Get Connection from Connection Pool

            //Test Connection
            if (!Conn.TestConnection())
            {
                RetObj.RetCode = 160; //0xA0
                RetObj.RetMsg = "資料庫連結失敗";
                return RetObj;
            }

            //判斷 CrmOperatorId 是否存在
            string SQL =
                string.Format("Select Count(*) CNT from CM003 Where EMPNO = '{0}'",
                              DnryParms["CrmOperatorId"]);
            DataTable DT = Conn.ExecuteQuery(SQL);
            if (DT == null)
            {
                RetObj.RetCode = 161; //0xA1
                RetObj.RetMsg = "資料庫查詢指令失敗 - 判斷 CrmOperatorId 是否存在";
                return RetObj;
            }

            DataView DV = DT.DefaultView;
            if (System.Convert.ToInt32(DV[0]["CNT"]) == 0)
            {
                RetObj.RetCode = 130; //0x82
                RetObj.RetMsg = "參數值不合法";
                return RetObj;
            }

            //裝機? 維修?
             SQL =
                string.Format("Select '裝機' Type from SO007 Where SNO = '{0}' union Select '維修' Type from SO008 Where SNO = '{0}'", 
                              DnryParms["CrmRelationNo"]);
            DT = Conn.ExecuteQuery(SQL);
            if (DT == null)
            {
                RetObj.RetCode = 161; //0xA1
                RetObj.RetMsg = "裝機? 維修? 指令失敗 " + Conn.LastExMessage;
                return RetObj;
            }

            DV = DT.DefaultView;
            if (DV.Count == 0)
            {
                RetObj.RetCode = 130; //0x82
                RetObj.RetMsg = "CrmRelationNo 參數值不合法";
                return RetObj;
            }

            //儲存 TYPE
            string Type = System.Convert.ToString(DV[0]["TYPE"]);

            //到此 工單已存在
            //判斷工單是否已完工
            if (Type.Equals("裝機"))
            {
                SQL =
                   string.Format("select Count(*) CNT from So007 where SNO = '{0}' and FinTime is null and CallOKTime is null",
                                 DnryParms["CrmRelationNo"]);
                DT = Conn.ExecuteQuery(SQL);
                if (DT == null)
                {
                    RetObj.RetCode = 161; //0xA1
                    RetObj.RetMsg = "資料庫查詢指令失敗 - 判斷裝機工單是否已完工";
                    return RetObj;
                }

                DV = DT.DefaultView;
                if (System.Convert.ToInt32(DV[0]["CNT"]) == 0)
                {
                    RetObj.RetCode = 130; //0x82
                    RetObj.RetMsg = "裝機工單已完工";
                    return RetObj;
                }
            }
            else
            {
                SQL =
                   string.Format("select Count(*) CNT from So008 where SNO = '{0}' and FinTime is null and CallOKTime is null",
                                 DnryParms["CrmRelationNo"]);
                DT = Conn.ExecuteQuery(SQL);
                if (DT == null)
                {
                    RetObj.RetCode = 161; //0xA1
                    RetObj.RetMsg = "資料庫查詢指令失敗 - 判斷維修工單是否已完工";
                    return RetObj;
                }

                DV = DT.DefaultView;
                if (System.Convert.ToInt32(DV[0]["CNT"]) == 0)
                {
                    RetObj.RetCode = 130; //0x82
                    RetObj.RetMsg = "維修工單已完工";
                    return RetObj;
                }
            }

            //裝機
            if (Type.Equals("裝機"))
            {
                ParmUnit = new ParmElement();
                ParmUnit.ParmName = "CrmId";
                ParmUnit.ParmValue = DnryParms["CrmId"];
                arylstParmElement.Add(ParmUnit);

                //抓取【服務方案代碼】及【路由廠商代碼】及【客戶資料】
                SQL =
                    "  select c.custid, c.custname, c.tel1, c.instaddress, " +
                    "         b.cmbaudno as SchemaCode1,                   " +
                    "         b.vendorcode as SchemaCode2                  " +
                    "    from so007 a, so004 b, so001 c                    " +
                    "   where a.sno = b.sno                                " +
                    "     and b.custid = c.custid                          " +
                    "     and a.sno = '" + DnryParms["CrmRelationNo"] + "'";

                DT = Conn.ExecuteQuery(SQL);
                if (DT == null)
                {
                    RetObj.RetCode = 161; //0xA1
                    RetObj.RetMsg = "裝機 抓取【服務方案代碼】及【路由廠商代碼】及【客戶資料】指令失敗 " + Conn.LastExMessage;
                    return RetObj;
                }

                DV = DT.DefaultView;
                if (DV.Count != 0)
                {
                    ParmUnit = new ParmElement();
                    ParmUnit.ParmName = "CrmCustId"; //客編
                    ParmUnit.ParmValue = System.Convert.ToString(DV[0]["CUSTID"]);
                    arylstParmElement.Add(ParmUnit);

                    ParmUnit = new ParmElement();
                    ParmUnit.ParmName = "CrmCustName"; //CRM系統客戶名稱
                    ParmUnit.ParmValue = System.Convert.ToString(DV[0]["CUSTNAME"]);
                    arylstParmElement.Add(ParmUnit);

                    ParmUnit = new ParmElement();
                    ParmUnit.ParmName = "CrmCustTel"; //電話
                    ParmUnit.ParmValue = System.Convert.ToString(DV[0]["TEL1"]);
                    arylstParmElement.Add(ParmUnit);

                    ParmUnit = new ParmElement();
                    ParmUnit.ParmName = "CrmCustAddress"; //裝機地址
                    ParmUnit.ParmValue = System.Convert.ToString(DV[0]["INSTADDRESS"]);
                    arylstParmElement.Add(ParmUnit);

                    ParmUnit = new ParmElement();
                    ParmUnit.ParmName = "SchemeCode1"; //服務方案類別代碼
                    ParmUnit.ParmValue = System.Convert.ToString(DV[0]["SCHEMACODE1"]);
                    arylstParmElement.Add(ParmUnit);

                    if ( !string.IsNullOrEmpty( Convert.ToString( DV[0]["SCHEMACODE2"] ).Trim() ) )
                    {
                        ParmUnit = new ParmElement();
                        ParmUnit.ParmName = "SchemeCode2"; //路由代碼
                        ParmUnit.ParmValue = System.Convert.ToString(DV[0]["SCHEMACODE2"]);
                        arylstParmElement.Add(ParmUnit);
                    }
                }
                else
                {
                    RetObj.RetCode = 176; //0xB0
                    RetObj.RetMsg = "裝機 抓取【服務方案代碼】及【路由廠商代碼】及【客戶資料】查無資料";
                    return RetObj;
                }

                //抓取【申請人】資訊, 此非必回傳資料
                SQL =
                    " select decode( c.refno, 1, b.id, decode( d.refno, 1, e.id2, null ) ) as UserIdNo," +
                    "        decode( c.refno, 1, b.declarantname , decode( d.refno, 1, e.declarantname, null ) )as UserName," +
                    "        decode( c.refno, 1, b.Birthday , decode( d.refno, 1, e.Birthday, null ) ) as UserBirthday " +
                    "   from so007 a, so004 b, so004 e, cd070 c , cd070 d  " +
                    "  where a.sno = b.sno                                 " +
                    "    and a.sno = e.sno                                 " +
                    "    and b.idkindcode1 = c.codeno(+)                   " +
                    "    and e.idkindcode = d.codeno(+)                    " +
                    "    and a.sno = '" + DnryParms["CrmRelationNo"] + "'";

                DT = Conn.ExecuteQuery(SQL);
                if (DT == null)
                {
                    RetObj.RetCode = 161; //0xA1
                    RetObj.RetMsg = "裝機 抓取【申請人】資訊 指令失敗 " + Conn.LastExMessage;
                    return RetObj;
                }

                DV = DT.DefaultView;
                if ( DV.Count != 0 )
                {
                    if ( !string.IsNullOrEmpty( Convert.ToString( DV[0]["USERIDNO"] ) ) )
                    {
                        ParmUnit = new ParmElement();
                        ParmUnit.ParmName = "UserIdNo"; //申請人生份證號
                        ParmUnit.ParmValue = System.Convert.ToString(DV[0]["USERIDNO"]);
                        arylstParmElement.Add( ParmUnit );
                    }

                    if ( !string.IsNullOrEmpty( Convert.ToString( DV[0]["USERNAME"] ) ) )
                    {
                        ParmUnit = new ParmElement();
                        ParmUnit.ParmName = "UserName"; //申請人姓名
                        ParmUnit.ParmValue = System.Convert.ToString(DV[0]["USERNAME"]);
                        arylstParmElement.Add( ParmUnit );
                    }

                    if ( !string.IsNullOrEmpty( Convert.ToString( DV[0]["USERBIRTHDAY"] ) ) )
                    {
                        ParmUnit = new ParmElement();
                        ParmUnit.ParmName = "UserBirthday"; //申請人生日
                        ParmUnit.ParmValue = System.Convert.ToDateTime(DV[0]["USERBIRTHDAY"]).ToString("yyyy-MM-dd");
                        arylstParmElement.Add( ParmUnit );
                    }
                }
            }

            //維修
            if (Type.Equals("維修"))
            {
                ParmUnit = new ParmElement();
                ParmUnit.ParmName = "CrmId";
                ParmUnit.ParmValue = DnryParms["CrmId"];
                arylstParmElement.Add(ParmUnit);

                //抓取【CM識別帳號】【方案速率代碼】【路營由代碼】資訊
                SQL =
                    " select b.dialaccount as ScmAccountId,   " +
                    "        b.cmbaudno as SchemaCode1,       " +
                    "        b.vendorcode as SchemaCode2      " +
                    "   from so008 a, so004 b                 " +
                    "  where a.sno = b.prsno                  " +
                    "    and a.sno = '" + DnryParms["CrmRelationNo"] + "'";

                DT = Conn.ExecuteQuery(SQL);
                if (DT == null)
                {
                    RetObj.RetCode = 161; //0xA1
                    RetObj.RetMsg = "維修 【CM識別帳號】相關資訊 指令失敗 " + Conn.LastExMessage;
                    return RetObj;
                }

                DV = DT.DefaultView;
                if ( DV.Count != 0 )
                {

                    ParmUnit = new ParmElement();
                    ParmUnit.ParmName = "ScmAccountId"; //CM識別帳號
                    ParmUnit.ParmValue = System.Convert.ToString( DV[0]["SCMACCOUNTID"] );
                    arylstParmElement.Add( ParmUnit );

                    if ( !string.IsNullOrEmpty( Convert.ToString( DV[0]["SCHEMACODE1"] ) ) )
                    {
                        ParmUnit = new ParmElement();
                        ParmUnit.ParmName = "SchemeCode1"; //服務方案類別代碼
                        ParmUnit.ParmValue = System.Convert.ToString( DV[0]["SCHEMACODE1"] );
                        arylstParmElement.Add( ParmUnit );
                    }

                    if ( !string.IsNullOrEmpty( Convert.ToString( DV[0]["SCHEMACODE2"] ) ) )
                    {
                        ParmUnit = new ParmElement();
                        ParmUnit.ParmName = "SchemeCode2"; //路由代碼
                        ParmUnit.ParmValue = System.Convert.ToString(DV[0]["SCHEMACODE2"]);
                        arylstParmElement.Add( ParmUnit );
                    }
                }
                else
                {
                    RetObj.RetCode = 176; //0xB0
                    RetObj.RetMsg = "維修 抓取【CM識別帳號】相關資訊 查無資料";
                    return RetObj;
                }
            }
            
            //Apply Result to RetObj
            RetObj.ParmAry = (ParmElement[])arylstParmElement.ToArray(typeof(ParmElement)); //Changed
            RetObj.RetCode = 0;
            RetObj.RetMsg = "OK";
        }

        return RetObj;
    }

    //	ActiveNewAccountByRelationNo: 回報啟用新服務帳號相關資訊
    //	使用時機: SCM供裝界面完成新用戶帳號建立開通後, 以此功能告知CRM開通相關資訊.
    //	呼叫參數: *ScmServiceType, *CrmId, *CrmCustId, *CrmRelationNo, *CrmOperatorId. *ScmAccountId, *ScmAccountPassword, *DeviceSNo1, *SchemeCode1, SchemeCode[2..9].
    //	回傳參數: 無.
    [WebMethod]
    public RetStruct ActiveNewAccountByRelationNo(ParmElement[] ParmAry)
    {
        ////Changed
        RetStruct RetObj = new RetStruct();

        ////Changed
        if (CRM.SetupConnectionString == null)
        {
            if (!InitialSetup())
            {
                RetObj.RetCode = 176; //0xB0
                RetObj.RetMsg = "初始環境失敗";
                return RetObj;
            }
        }

        Dictionary<string,string> DnryParms = null;
        if (ParmAry != null)
        {
            DnryParms = new Dictionary<string, string>();

            for (int i = 0; i < ParmAry.Length; i++)
            {
                DnryParms[ParmAry[i].ParmName] = ParmAry[i].ParmValue;
            }
        }

        if (DnryParms == null)
        {
            //Changed
            RetObj.RetCode = 129; //0x81
            RetObj.RetMsg = "沒有任何參數";
            return RetObj;
        }
        else
        {
            //check necessary parameters
            String RetMsg = "缺少必要參數 ";
            if (DnryParms["CrmId"] == null)
            {
                RetMsg += "CrmId, ";
            }

            if (DnryParms["CrmCustId"] == null)
            {
                RetMsg += "CrmCustId, ";
            }

            if (DnryParms["ScmServiceType"] == null)
            {
                RetMsg += "ScmServiceType, ";
            }

            if (DnryParms["CrmRelationNo"] == null)
            {
                RetMsg += "CrmRelationNo, ";
            }

            if (DnryParms["CrmOperatorId"] == null)
            {
                RetMsg += "CrmOperatorId, ";
            }

            if (DnryParms["ScmAccountId"] == null)
            {
                RetMsg += "ScmAccountId, ";
            }

            if (DnryParms["ScmAccountPassword"] == null)
            {
                RetMsg += "ScmAccountPassword, ";
            }

            if (DnryParms["DeviceSNo1"] == null)
            {
                RetMsg += "DeviceSNo1, ";
            }

            if (!RetMsg.Equals("缺少必要參數 "))
            {
                RetObj.RetCode = 129; //0x81
                RetObj.RetMsg = RetMsg.Substring(0, RetMsg.Length - 2);
                return RetObj;
            }

            //Get Connection from Connection Pool
            string ConnectionString = (string)hashOracleConnectionString[DnryParms["CrmId"]];

            if (ConnectionString == null)
            {
                RetObj.RetCode = 130; //0x82
                RetObj.RetMsg = "參數值 CrmId 不合法";
                return RetObj;
            }

            DB_Oracle Conn = new DB_Oracle(ConnectionString); //Get Connection from Connection Pool

            //Test Connection
            if (!Conn.TestConnection())
            {
                RetObj.RetCode = 160; //0xA0
                RetObj.RetMsg = "資料庫連結失敗";
                return RetObj;
            }

            //判斷 CrmOperatorId 是否存在
            string SQL =
                string.Format("Select Count(*) CNT from CM003 Where EMPNO = '{0}'",
                              DnryParms["CrmOperatorId"]);
            DataTable DT = Conn.ExecuteQuery(SQL);
            if (DT == null)
            {
                RetObj.RetCode = 161; //0xA1
                RetObj.RetMsg = "資料庫查詢指令失敗 - 判斷 CrmOperatorId 是否存在";
                return RetObj;
            }

            DataView DV = DT.DefaultView;
            if (System.Convert.ToInt32(DV[0]["CNT"]) == 0)
            {
                RetObj.RetCode = 130; //0x82
                RetObj.RetMsg = "參數值不合法";
                return RetObj;
            }

            //檢查 SO004 設備檔
            SQL = "SELECT Count(*) CNT" +
                  "  FROM SO004" +
                  " WHERE SNO = '{0}'";
            SQL = string.Format(SQL, DnryParms["CrmRelationNo"]);

            DT = Conn.ExecuteQuery(SQL);
            if (DT == null)
            {
                RetObj.RetCode = 161; //0xA1
                RetObj.RetMsg = "資料庫查詢指令失敗：檢查 SO004 設備檔";
                return RetObj;
            }
            else
            {
                DV = DT.DefaultView;
                if (System.Convert.ToInt32(DV[0]["CNT"]) != 1)
                {
                    RetObj.RetCode = 130; //0x82
                    RetObj.RetMsg = "設備檔資料筆數錯誤";
                    return RetObj;
                }
            }

            //檢查裝機工單
            SQL = "SELECT Count(*) CNT" +
                  "  FROM SO007" +
                  " WHERE SNO = '{0}'";
            SQL = string.Format(SQL, DnryParms["CrmRelationNo"]);

            DT = Conn.ExecuteQuery(SQL);
            if (DT == null)
            {
                RetObj.RetCode = 161; //0xA1
                RetObj.RetMsg = "資料庫查詢指令失敗：檢查 裝機工單";
                return RetObj;
            }
            else
            {
                DV = DT.DefaultView;
                if (System.Convert.ToInt32(DV[0]["CNT"]) != 1)
                {
                    RetObj.RetCode = 130; //0x82
                    RetObj.RetMsg = "裝機工單筆數錯誤";
                    return RetObj;
                }
            }

            //1.先取工程人員姓名
            string EmpName = string.Empty;
            
            SQL = "Select EmpName from CM003 where EmpNo = '{0}'";
            SQL = string.Format(SQL, DnryParms["CrmOperatorId"].Replace("'", "''"));

            DT = Conn.ExecuteQuery(SQL);
            if (DT == null)
            {
                RetObj.RetCode = 161; //0xA1
                RetObj.RetMsg = "資料庫查詢指令失敗：取得工程人員姓名失敗";
                return RetObj;
            }
            else
            {
                DV = DT.DefaultView;
                if (DV.Count != 0) EmpName = System.Convert.ToString(DV[0]["EmpName"]);
            }

            //2.取設備型號
            string ModelCode = string.Empty;
            string ModelName = string.Empty;
            SQL = "SELECT b.codeno AS modelcode, b.description AS modelname" +
                  "  FROM so306 a, cd043 b       " +
                  " WHERE a.modelcode = b.codeno " +
                  "   AND a.compcode = {0}       " +
                  "   AND a.hfcmac = '{1}'       ";
            SQL = string.Format(SQL, DnryParms["CrmId"], DnryParms["DeviceSNo1"]);

            DT = Conn.ExecuteQuery(SQL);
            if (DT == null)
            {
                RetObj.RetCode = 161; //0xA1
                RetObj.RetMsg = "資料庫查詢指令失敗：取得設備型號失敗";
                return RetObj;
            }
            else
            {
                DV = DT.DefaultView;
                if (DV.Count != 0)
                {
                    ModelCode = System.Convert.ToString(DV[0]["ModelCode"]);
                    ModelName = System.Convert.ToString(DV[0]["ModelName"]);
                }
            }


            //3.若是有傳路由代碼, 則取路由名稱
            string aVendorCode = string.Empty;
            string aVendorName = string.Empty;
            if (DnryParms["SchemeCode2"] != null)
            {
                SQL = string.Format(
                    " SELECT a.VendorCode, a.VendorName " +
                    "   FROM CM006 a                    " +
                    "  WHERE a.VendorCode = '{0}'       ", DnryParms["SchemeCode2"]);

                DT = Conn.ExecuteQuery(SQL);

                if (DT == null)
                {
                    RetObj.RetCode = 161; //0xA1
                    RetObj.RetMsg = "資料庫查詢指令失敗：取得路由名稱失敗";
                    return RetObj;
                }
                else
                {
                    DV = DT.DefaultView;
                    if (DV.Count != 0)
                    {
                        aVendorCode = System.Convert.ToString(DV[0]["VendorCode"]);
                        aVendorName = System.Convert.ToString(DV[0]["VendorName"]);
                    }
                }
            }


            //4.更新 SO004 設備檔
            ArrayList SQLs = new ArrayList();
            SQL = " UPDATE SO004                  " +
                  "    SET CMOpenDate = sysdate,  " +
                  "        FaciStatusCode = 1,    " +
                  "        InstEn1 = '{0}',       " +
                  "        InstName1 = '{1}',     " +
                  "        FaciSNo = '{2}',       " +
                  "        ModelCode = '{3}',     " +
                  "        ModelName = '{4}',     " +
                  "        DialAccount = '{5}',   " +
                  "        DialPassword = '{6}',  " +
                  "        VendorCode = nvl( '{7}', VendorCode ),   " +
                  "        VendorName = nvl( '{8}', VendorName )    " +
                  "  WHERE SNo = '{9}'            ";

            SQL = string.Format( SQL,
                    DnryParms["CrmOperatorId"], 
                    EmpName, 
                    DnryParms["DeviceSNo1"], 
                    ModelCode, 
                    ModelName, 
                    DnryParms["ScmAccountId"], 
                    DnryParms["ScmAccountPassword"], 
                    aVendorCode,
                    aVendorName,
                    DnryParms["CrmRelationNo"] );

            SQLs.Add(SQL);

            //5.更新裝機工單
            SQL = "UPDATE so007" +
                  "   SET WorkerEn1 = '{0}'," +
                  "       WorkerName1 = '{1}'," +
                  "       CallOkTime = sysdate" +
                  " WHERE SNo = '{2}'";
            SQL = string.Format(SQL, DnryParms["CrmOperatorId"], EmpName, DnryParms["CrmRelationNo"]);
            SQLs.Add(SQL);

            //更新 SO004 設備檔 及 更新裝機工單
            if (!Conn.ExecuteNonQuery(SQLs))
            {
                RetObj.RetCode = 162; //0xA2

                if (Conn.LastExecIndex == 0)
                    RetObj.RetMsg = "更新 SO004 設備檔錯誤; " + Conn.LastExMessage;
                else if (Conn.LastExecIndex == 1)
                    RetObj.RetMsg = "更新裝機工單錯誤; " + Conn.LastExMessage;
                else
                    RetObj.RetMsg = Conn.LastExMessage;

                return RetObj;
            }

            //Apply Result to RetObj
            RetObj.RetCode = 0;
            RetObj.RetMsg = "OK";
        }

        return RetObj;
    }

    //	ChangeDeviceByRelationNo: 回報服務帳號變更配接設備相關資訊
    //	使用時機: SCM供裝界面完成用戶換機設定後, 以此功能告知CRM換裝設備資訊.
    //	呼叫參數: *ScmServiceType, *CrmId, *CrmCustId, *CrmRelationNo, *CrmOperatorId, *ScmAccountId, *DeviceSNo1
    //	回傳參數: 無
    [WebMethod]
    public RetStruct ChangeDeviceByRelationNo(ParmElement[] ParmAry)
    {
        ////Changed
        RetStruct RetObj = new RetStruct();

        ////Changed
        if (CRM.SetupConnectionString == null)
        {
            if (!InitialSetup())
            {
                RetObj.RetCode = 176; //0xB0
                RetObj.RetMsg = "初始環境失敗";
                return RetObj;
            }
        }

        Dictionary<string, string> DnryParms = null;
        if (ParmAry != null)
        {
            DnryParms = new Dictionary<string, string>();

            for (int i = 0; i < ParmAry.Length; i++)
            {
                DnryParms[ParmAry[i].ParmName] = ParmAry[i].ParmValue;
            }
        }

        if (DnryParms == null)
        {
            //Changed
            RetObj.RetCode = 129; //0x81
            RetObj.RetMsg = "沒有任何參數";
            return RetObj;
        }
        else
        {
            //check necessary parameters
            String RetMsg = "缺少必要參數 ";
            if (DnryParms["CrmId"] == null)
            {
                RetMsg += "CrmId, ";
            }

            if (DnryParms["CrmCustId"] == null)
            {
                RetMsg += "CrmCustId, ";
            }

            if (DnryParms["ScmServiceType"] == null)
            {
                RetMsg += "ScmServiceType, ";
            }

            if (DnryParms["CrmRelationNo"] == null)
            {
                RetMsg += "CrmRelationNo, ";
            }

            if (DnryParms["CrmOperatorId"] == null)
            {
                RetMsg += "CrmOperatorId, ";
            }

            if (DnryParms["ScmAccountId"] == null)
            {
                RetMsg += "ScmAccountId, ";
            }

            if (DnryParms["DeviceSNo1"] == null)
            {
                RetMsg += "DeviceSNo1, ";
            }

            if (!RetMsg.Equals("缺少必要參數 "))
            {
                RetObj.RetCode = 129; //0x81
                RetObj.RetMsg = RetMsg.Substring(0, RetMsg.Length - 2);
                return RetObj;
            }

            //Get Connection from Connection Pool
            string ConnectionString = (string)hashOracleConnectionString[(string)DnryParms["CrmId"]];

            if (ConnectionString == null)
            {
                RetObj.RetCode = 130; //0x82
                RetObj.RetMsg = "參數值 CrmId 不合法";
                return RetObj;
            }

            DB_Oracle Conn = new DB_Oracle(ConnectionString); //Get Connection from Connection Pool

            //Test Connection
            if (!Conn.TestConnection())
            {
                RetObj.RetCode = 160; //0xA0
                RetObj.RetMsg = "資料庫連結失敗";
                return RetObj;
            }

            //判斷 CrmOperatorId 是否存在
            string SQL =
                string.Format("Select Count(*) CNT from CM003 Where EMPNO = '{0}'",
                              DnryParms["CrmOperatorId"]);
            DataTable DT = Conn.ExecuteQuery(SQL);
            if (DT == null)
            {
                RetObj.RetCode = 161; //0xA1
                RetObj.RetMsg = "資料庫查詢指令失敗 - 判斷 CrmOperatorId 是否存在";
                return RetObj;
            }

            DataView DV = DT.DefaultView;
            if (System.Convert.ToInt32(DV[0]["CNT"]) == 0)
            {
                RetObj.RetCode = 130; //0x82
                RetObj.RetMsg = "參數值不合法";
                return RetObj;
            }

            //檢查 SO004 設備檔
            SQL = "SELECT Count(*) CNT" +
                  "  FROM SO004" +
                  " WHERE SNO = '{0}'";
            SQL = string.Format(SQL, DnryParms["CrmRelationNo"]);

            DT = Conn.ExecuteQuery(SQL);
            if (DT == null)
            {
                RetObj.RetCode = 161; //0xA1
                RetObj.RetMsg = "資料庫查詢指令失敗：檢查 SO004 設備檔";
                return RetObj;
            }
            else
            {
                DV = DT.DefaultView;
                if (System.Convert.ToInt32(DV[0]["CNT"]) != 1)
                {
                    RetObj.RetCode = 130; //0x82
                    RetObj.RetMsg = "新設備資料筆數錯誤";
                    return RetObj;
                }
            }

            SQL = "SELECT Count(*) CNT" +
                  "  FROM SO004" +
                  " WHERE PRSNO = '{0}'";
            SQL = string.Format(SQL, DnryParms["CrmRelationNo"]);

            DT = Conn.ExecuteQuery(SQL);
            if (DT == null)
            {
                RetObj.RetCode = 161; //0xA1
                RetObj.RetMsg = "資料庫查詢指令失敗：檢查 SO004 設備檔";
                return RetObj;
            }
            else
            {
                DV = DT.DefaultView;
                if (System.Convert.ToInt32(DV[0]["CNT"]) != 1)
                {
                    RetObj.RetCode = 130; //0x82
                    RetObj.RetMsg = "舊設備資料筆數錯誤";
                    return RetObj;
                }
            }

            //檢查維修工單
            SQL = "SELECT Count(*) CNT" +
                  "  FROM SO008" +
                  " WHERE SNO = '{0}'";
            SQL = string.Format(SQL, DnryParms["CrmRelationNo"]);

            DT = Conn.ExecuteQuery(SQL);
            if (DT == null)
            {
                RetObj.RetCode = 161; //0xA1
                RetObj.RetMsg = "資料庫查詢指令失敗：檢查 維修工單";
                return RetObj;
            }
            else
            {
                DV = DT.DefaultView;
                if (System.Convert.ToInt32(DV[0]["CNT"]) != 1)
                {
                    RetObj.RetCode = 130; //0x82
                    RetObj.RetMsg = "維修工單筆數錯誤";
                    return RetObj;
                }
            }

            //1.取工程人員姓名
            string EmpName = string.Empty;

            SQL = "Select EmpName from CM003 where EmpNo = '{0}'";
            SQL = string.Format(SQL, DnryParms["CrmOperatorId"].Replace("'", "''"));

            DT = Conn.ExecuteQuery(SQL);
            if (DT == null)
            {
                RetObj.RetCode = 161; //0xA1
                RetObj.RetMsg = "資料庫查詢指令失敗：取得工程人員姓名失敗";
                return RetObj;
            }
            else
            {
                DV = DT.DefaultView;
                if (DV.Count != 0) EmpName = System.Convert.ToString(DV[0]["EmpName"]);
            }

            //2.取設備型號
            string ModelCode = string.Empty;
            string ModelName = string.Empty;
            SQL = "SELECT b.codeno AS modelcode, b.description AS modelname" +
                  "  FROM so306 a, cd043 b       " +
                  " WHERE a.modelcode = b.codeno " +
                  "   AND a.compcode = {0}       " +
                  "   AND a.hfcmac = '{1}'       ";
            SQL = string.Format(SQL, DnryParms["CrmId"], DnryParms["DeviceSNo1"]);

            DT = Conn.ExecuteQuery(SQL);
            if (DT == null)
            {
                RetObj.RetCode = 161; //0xA1
                RetObj.RetMsg = "資料庫查詢指令失敗：取得設備型號失敗";
                return RetObj;
            }
            else
            {
                DV = DT.DefaultView;
                if (DV.Count != 0)
                {
                    ModelCode = System.Convert.ToString(DV[0]["ModelCode"]);
                    ModelName = System.Convert.ToString(DV[0]["ModelName"]);
                }
            }

            //3.取舊設備的 AccountPassword
            string DialPassword = string.Empty;
            SQL = "Select DialPassword As OldPwd from SO004 Where PRSNO = '{0}'";
            SQL = string.Format(SQL, DnryParms["CrmRelationNo"]);
            
            DT = Conn.ExecuteQuery(SQL);
            if (DT == null)
            {
                RetObj.RetCode = 161; //0xA1
                RetObj.RetMsg = "資料庫查詢指令失敗：取舊設備的 AccountPassword 失敗";
                return RetObj;
            }
            else
            {
                DV = DT.DefaultView;
                if (DV.Count != 0) DialPassword = System.Convert.ToString(DV[0]["DialPassword"]);
            }


            //5.若是有傳路由代碼, 則取路由名稱
            string aVendorCode = string.Empty;
            string aVendorName = string.Empty;
            if (DnryParms["SchemaCode2"] != null)
            {
                SQL = string.Format(
                    " SELECT a.VendorCode, a.VendorName " +
                    "   FROM CM006 a                    " +
                    "  WHERE a.VendorCode = '{0}'       ", DnryParms["SchemdCode2"] );

                DT = Conn.ExecuteQuery( SQL );

                if (DT == null)
                {
                    RetObj.RetCode = 161; //0xA1
                    RetObj.RetMsg = "資料庫查詢指令失敗：取得路由名稱失敗";
                    return RetObj;
                }
                else
                {
                    DV = DT.DefaultView;
                    if (DV.Count != 0)
                    {
                        aVendorCode = System.Convert.ToString( DV[0]["VendorCode"] );
                        aVendorName = System.Convert.ToString( DV[0]["VendorName"] );
                    }
                }
            }


            //6.更新 SO004 設備檔 ( 新設備 )
            ArrayList SQLs = new ArrayList();

            SQL = " UPDATE SO004                  " +
                  "    SET CMOpenDate = sysdate,  " +
                  "        FaciStatusCode = 1,    " +
                  "        InstEn1 = '{0}',       " +
                  "        InstName1 = '{1}',     " +
                  "        FaciSNo = '{2}',       " +
                  "        ModelCode = '{3}',     " +
                  "        ModelName = '{4}',     " +
                  "        DialAccount = '{5}',   " +
                  "        DialPassword = '{6}',  " +
                  "        VendorCode = nvl( '{7}', VendorCode ),   " +
                  "        VendorName = nvl( '{8}', VendorName )    " +
                  "  WHERE SNo = '{9}'            ";

            SQL = string.Format( SQL,
                    DnryParms["CrmOperatorId"],
                    EmpName,
                    DnryParms["DeviceSNo1"],
                    ModelCode,
                    ModelName,
                    DnryParms["ScmAccountId"],
                    DialPassword,
                    aVendorCode,
                    aVendorName,
                    DnryParms["CrmRelationNo"] );

            SQLs.Add( SQL );

            //7.更新 SO004 設備檔 ( 舊設備 )
            SQL = " UPDATE SO004                   " +
                  "    SET CMCloseDate = sysdate,  " +
                  "        FaciStatusCode = 2,     " +
                  "        PrEn1 = '{1}',          " +
                  "        PrName1 = '{2}'         " +
                  "  WHERE PRSNO = '{3}'           ";
            SQL = string.Format( SQL, DnryParms["CrmOperatorId"], EmpName, DnryParms["CrmRelationNo"] );
            SQLs.Add(SQL);


            //8.更新維修工單
            SQL = "UPDATE SO008                    " +
                  "   SET WorkerEn1 = '{0}',       " +
                  "       WorkerName1 = '{1}',     " +
                  "       CallOkTime = sysdate     " +
                  " WHERE SNo = '{2}'              ";
            SQL = string.Format( SQL, DnryParms["CrmOperatorId"], EmpName, DnryParms["CrmRelationNo"] );
            SQLs.Add( SQL );

            //更新 SO004 設備檔 ( 新設備 舊設備) 及 更新維修工單
            if (!Conn.ExecuteNonQuery(SQLs))
            {
                RetObj.RetCode = 162; //0xA2

                if (Conn.LastExecIndex == 0)
                    RetObj.RetMsg = "更新 SO004 設備檔 ( 新設備 ) 錯誤; " + Conn.LastExMessage;
                else if (Conn.LastExecIndex == 1)
                    RetObj.RetMsg = "更新 SO004 設備檔 ( 舊設備 ) 錯誤; " + Conn.LastExMessage;
                else if (Conn.LastExecIndex == 2)
                    RetObj.RetMsg = "更新維修工單錯誤; " + Conn.LastExMessage;
                else
                    RetObj.RetMsg = Conn.LastExMessage;

                return RetObj;
            }

            //Apply Result to RetObj
            RetObj.RetCode = 0;
            RetObj.RetMsg = "OK";
        }

        return RetObj;
    }

    #region DB_Oracle
    class DB_Oracle : IDisposable
    {
        #region Declarations
        private bool disposed = false;
        //private SystemMessageBox SysMessageBox;
        private OracleConnection Conn;
        public ConnectionStates ConnectionState = ConnectionStates.Unknow;

        //private ArrayList observers;
        #endregion

        #region Enum
        public enum ConnectionStates
        {
            Unknow = 0,
            OK = 1,
            Error = 2,
        }
        #endregion

        #region Properties
        string _CompanyCode;
        public string CompanyCode
        {
            get
            {
                return _CompanyCode;
            }

            set
            {
                if (Convert.IsDBNull(value))
                {
                    _CompanyCode = "-1";
                }
                else _CompanyCode = value;
            }
        }

        //string _CompanyName;
        //public string CompanyName
        //{
        //    get
        //    {
        //        return _CompanyName;
        //    }

        //    set
        //    {
        //        if (Convert.IsDBNull(value))
        //        {
        //            _CompanyName = string.Empty;
        //        }
        //        else _CompanyName = value;
        //    }
        //}

        //int _WaitInProcessCount;
        //public int WaitInProcessCount
        //{
        //    get
        //    {
        //        return _WaitInProcessCount;
        //    }

        //    set
        //    {
        //        _WaitInProcessCount = value;
        //    }
        //}

        private string _ExMessage = string.Empty;
        public string LastExMessage
        {
            get { return _ExMessage; }
        }

        private int _LastExecIndex = -1;
        public int LastExecIndex
        {
            get { return _LastExecIndex; }
        }
        #endregion

        #region Constructor / Destructor
        //public DB_Oracle(SystemMessageBox SysMessageBox, string ServiceName, string UserID, string Password)
        public DB_Oracle(string ServiceName, string UserID, string Password)
        {
            //this.SysMessageBox = SysMessageBox;
            string strConn = string.Format("Data Source={0};Persist Security Info=True;User ID={1};Password={2}", new object[] { ServiceName, UserID, Password });
            Conn = new OracleConnection(strConn);
            //observers = new ArrayList();
        }

        public DB_Oracle(string ConnectionString)
        {
            //this.SysMessageBox = SysMessageBox;
            Conn = new OracleConnection(ConnectionString);
            //observers = new ArrayList();
        }

        ~DB_Oracle()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                if (disposing)
                {
                    try
                    {
                        ////if (Conn != null) Conn.Dispose();
                        Conn = null;

                        //for (int i = 0; i < observers.Count; i++)
                        //{
                        //    ((DBObserver)observers[i]).Dispose();
                        //}
                        //observers.Clear();
                        //observers = null;
                    }
                    catch (Exception) { }
                }

            }
            disposed = true;
        }
        #endregion

        #region Test Connection
        internal bool TestConnection()
        {
            lock (this.Conn)
            {
                if (Conn == null)
                {
                    this.ConnectionState = ConnectionStates.Error;
                    //this.notifyObservers();
                    return false;
                }

                try
                {
                    if (Conn.State == System.Data.ConnectionState.Closed) Conn.Open();
                    Conn.Close();
                    this.ConnectionState = ConnectionStates.OK;
                    _ExMessage = "";
                    //this.notifyObservers();
                    return true;
                }
                catch (Exception ex)
                {
                    this.ConnectionState = ConnectionStates.Error;
                    //SysMessageBox.showError(ex, false); //重試連線不寫 Log
                    _ExMessage = ex.Message;
                    //this.notifyObservers();
                    return false;
                }
            }
        }
        #endregion

        #region ExecuteNonQuery
        public bool ExecuteNonQuery(ArrayList arylstSQL)
        {
            OracleCommand cmd = null;
            OracleTransaction transaction = null;
            int iExecIndex = -1;

            try
            {
                if (Conn.State == System.Data.ConnectionState.Closed) Conn.Open();
                cmd = Conn.CreateCommand();
                transaction = Conn.BeginTransaction(IsolationLevel.ReadCommitted);

                for (int i = 0; i < arylstSQL.Count; i++)
                {
                    iExecIndex = i;
                    cmd.CommandText = (string)arylstSQL[i];
                    if (cmd.ExecuteNonQuery() == 0)
                        throw new Exception("找不到資料可以更新。");
                }

                transaction.Commit();
                this.ConnectionState = ConnectionStates.OK;
                //this.notifyObservers();
                _ExMessage = "";
                return true;
            }
            catch (System.Exception ex)
            {
                try
                {
                    if (transaction != null) transaction.Rollback();
                }
                catch //(System.Exception ex2)
                {
                    //SysMessageBox.showError(ex2, "transaction.Rollback() Error" + " " + ex.Message + System.Environment.NewLine + ex.StackTrace, true);
                }

                //if (iExecIndex == -1)
                //    SysMessageBox.showError(ex, true);
                //else
                //    SysMessageBox.showError(ex, "執行錯誤的 SQL 語法：" + ((string)arylstSQL[iExecIndex]).Replace("'", "''"), true);

                _ExMessage = ex.Message;
                _LastExecIndex = iExecIndex;
                this.ConnectionState = ConnectionStates.Error;
                //this.notifyObservers();
                return false;
            }
            finally
            {
                try
                {
                    Conn.Close();
                }
                catch //(System.Exception ex2)
                {
                    //SysMessageBox.showError(ex2, "Conn.Close() Error" + " " + ex2.Message + System.Environment.NewLine + ex2.StackTrace, true);
                }
            }
        }

        public bool ExecuteNonQuery(string strSQL)
        {
            ArrayList arylstSQL = new ArrayList();
            arylstSQL.Add(strSQL);
            return ExecuteNonQuery(arylstSQL);
        }
        #endregion

        #region ExecuteQuery
        public System.Data.DataTable ExecuteQuery(string strSQL)
        {
            OracleDataAdapter DA = null;
            System.Data.DataSet DS = null;

            try
            {
                if (Conn.State == System.Data.ConnectionState.Closed) Conn.Open();

                DA = new OracleDataAdapter(strSQL, Conn);
                DS = new System.Data.DataSet();
                DA.FillSchema(DS, System.Data.SchemaType.Mapped);
                DA.Fill(DS);

                _ExMessage = "";
                this.ConnectionState = ConnectionStates.OK;
                //this.notifyObservers();
                return DS.Tables[0];
            }
            catch (System.Exception ex)
            {
                //SysMessageBox.showError(ex, true);
                _ExMessage = ex.Message;
                this.ConnectionState = ConnectionStates.Error;
                //this.notifyObservers();
                return null;
            }
            finally
            {
                try
                {
                    Conn.Close();
                }
                catch //(System.Exception ex2)
                {
                    //SysMessageBox.showError(ex2, "Conn.Close() Error" + " " + ex2.Message + System.Environment.NewLine + ex2.StackTrace, true);
                }
            }
        }
        #endregion

        //#region Functions of Observer Pattern
        //public void addObserver(DBObserver observer)
        //{
        //    observers.Add(observer);
        //}

        //public void deleteObserver(DBObserver observer)
        //{
        //    observers.Remove(observer);
        //}

        //public void notifyObservers()
        //{
        //    IEnumerator it = observers.GetEnumerator();
        //    while (it.MoveNext())
        //    {
        //        DBObserver o = (DBObserver)it.Current;
        //        o.update(this);
        //    }
        //}
        //#endregion
    }
    #endregion
}