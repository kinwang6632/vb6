unit frmLoadDataU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ExtCtrls, ComCtrls, Grids, DBGrids, DB,ADODB,
  fraYMDU;

type
  TSGSCD024Data = class(TObject)
    sChCode : String;
    sChName : String;
    sProductCode : String;
  end;


  TfrmLoadData = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    medICCardPath: TEdit;
    cobCompCode: TComboBox;
    lblComp: TLabel;
    lblICCardPath: TLabel;
    OpenDialog1: TOpenDialog;
    btnLoad: TBitBtn;
    btnExit: TBitBtn;
    btnICCardPath: TButton;
    stbStatus: TStatusBar;
    lblChannelPath: TLabel;
    medChannelPath: TEdit;
    btnChannelPath: TButton;
    btnXml: TBitBtn;
    lblCAProductPath: TLabel;
    medCAProductPath: TEdit;
    btnCAProductPath: TButton;
    lblProdPurchasePath: TLabel;
    medProdPurchasePath: TEdit;
    btnProdPurchasePath: TButton;
    lblEntitlementPath: TLabel;
    medEntitlementPath: TEdit;
    btnEntitlementPath: TButton;
    chkICCard: TCheckBox;
    chkCAProduct: TCheckBox;
    chkProdPurchase: TCheckBox;
    chkEntitlement: TCheckBox;
    memErrorLog: TMemo;
    btnTransfer: TBitBtn;
    lblTransferPath: TLabel;
    medTransferPath: TEdit;
    btnTransferPath: TButton;
    btnTransferSMS: TBitBtn;
    btnTransferSMSSO005: TBitBtn;
    lblLogAndXmlPath: TLabel;
    fraYMD3: TfraYMD;
    Label1: TLabel;
    fraYMD4: TfraYMD;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnICCardPathClick(Sender: TObject);
    procedure btnLoadClick(Sender: TObject);
    procedure btnChannelPathClick(Sender: TObject);
    procedure btnXmlClick(Sender: TObject);
    procedure btnCAProductPathClick(Sender: TObject);
    procedure btnProdPurchasePathClick(Sender: TObject);
    procedure btnEntitlementPathClick(Sender: TObject);
    procedure btnTransferClick(Sender: TObject);
    procedure btnTransferPathClick(Sender: TObject);
    procedure btnTransferSMSClick(Sender: TObject);
    procedure btnTransferSMSSO005Click(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure fraYMD3Exit(Sender: TObject);
  private
    { Private declarations }
    procedure setLanguageString;
    function isDataOK : Boolean;
    function isDataOK2 : Boolean;
    function isDataOK3 : Boolean;
    function isDataOK4 : Boolean;
    function getCobCompCode : String;
    function checkIsOkDataFile(sI_ModeType : String;nI_ParamCounts,nI_ChannelParamCounts : Integer) : Boolean;
    function handleICCardQuery(I_AllDataTxt : TStringList;sI_CompCode : String;dI_CurrTime : TTime) : String;
    function handleICCardTransTxt(I_AllDataTxt : TStringList;sI_CompCode,sI_ModeType : String;dI_CurrTime : TTime) : String;

    function handleSMSSO004TransTxt(sI_CompCode,sI_ModeType : String;dI_CurrTime : TTime) : String;
    function handleSMSSO005TransTxt(sI_CompCode,sI_ModeType : String;dI_CurrTime : TTime) : String;

    function handleCAProductQuery(I_AllDataTxt,L_ChannelDataTxt : TStringList;sI_CompCode : String;dI_CurrTime : TTime) : String;
    function handleCAProductTransTxt(I_AllDataTxt,L_ChannelDataTxt : TStringList;sI_CompCode,sI_ModeType : String;dI_CurrTime : TTime) : String;

    function handleProdPurchaseQuery(I_AllDataTxt : TStringList;sI_CompCode : String;dI_CurrTime : TTime) : String;
    function handleProdPurchaseTransTxt(I_AllDataTxt : TStringList;sI_CompCode,sI_ModeType : String;dI_CurrTime : TTime) : String;

    function handleEntitlementQuery(I_AllDataTxt : TStringList;sI_CompCode : String;dI_CurrTime : TTime) : String;
    function handleEntitlementTransTxt(I_AllDataTxt : TStringList;sI_CompCode,sI_ModeType : String;dI_CurrTime : TTime) : String;


    function handleLoadDataLog(sI_CompCode,sI_ModeType,sI_StartDate,sI_StopDate : String) : String;
    function handleLoadDataXml(sI_CompCode,sI_ModeType,sI_StartDate,sI_StopDate : String) : String;
    procedure showLogMsg(sI_LogMsg : String);
    function getSGSCD024Data(sI_CompCode : String) : String;

    //L_Obj : TSGSParamIniData;



  public
    { Public declarations }
    G_SGSCD024DataList : TStringList;
    function getTransferPath(sI_ModeType : String;nI_FileNum : Integer) : String;
    function getSMSChCode(sI_ProductCode : String;var sI_ChCode,sI_ChName : String) : String;
  end;

var
  frmLoadData: TfrmLoadData;

implementation

uses dtmMainU, Ustru, UdateTimeu, ConstU, frmMainU;

{$R *.dfm}

procedure TfrmLoadData.FormShow(Sender: TObject);
var
  ii : Integer;
begin
  cobCompCode.Clear;
  for ii:=0 to dtmMain.G_CompCodeAndNameStrList.Count -1 do
    cobCompCode.Items.Add(dtmMain.G_CompCodeAndNameStrList.Strings[ii]);

  cobCompCode.ItemIndex := 0;

  medICCardPath.Text := dtmMain.sG_ExePath + '\TestData\sms_iccard.txt';
  medCAProductPath.Text := dtmMain.sG_ExePath + '\TestData\sms_caproduct.txt';
  medChannelPath.Text := dtmMain.sG_ExePath + '\TestData\sms_channel.txt';
  medProdPurchasePath.Text := dtmMain.sG_ExePath + '\TestData\sms_prodpurchase.txt';
  medEntitlementPath.Text := dtmMain.sG_ExePath + '\TestData\sms_entitlement.txt';
  //medTransferPath.Text := dtmMain.sG_ExePath + '\TransferData';
end;

procedure TfrmLoadData.setLanguageString;
begin
    Caption := dtmMain.getLanguageSettingInfo('frmLoadData_Caption');
    Panel1.Caption := dtmMain.getLanguageSettingInfo('frmLoadData_Panel1_Caption');
    lblComp.Caption := dtmMain.getLanguageSettingInfo('frmLoadData_lblComp_Caption');
    lblICCardPath.Caption := dtmMain.getLanguageSettingInfo('frmLoadData_lblICCardPath_Caption');
    lblCAProductPath.Caption := dtmMain.getLanguageSettingInfo('frmLoadData_lblCAProductPath_Caption');
    lblChannelPath.Caption := dtmMain.getLanguageSettingInfo('frmLoadData_lblChannelPath_Caption');
    lblProdPurchasePath.Caption := dtmMain.getLanguageSettingInfo('frmLoadData_lblProdPurchasePath_Caption');
    lblEntitlementPath.Caption := dtmMain.getLanguageSettingInfo('frmLoadData_lblEntitlementPath_Caption');



    btnLoad.Caption := dtmMain.getLanguageSettingInfo('frmLoadData_btnLoad_Caption');
    lblChannelPath.Caption := dtmMain.getLanguageSettingInfo('frmLoadData_lblChannelPath_Caption');
    btnXml.Caption := dtmMain.getLanguageSettingInfo('frmLoadData_btnXml_Caption');
    btnTransfer.Caption := dtmMain.getLanguageSettingInfo('frmLoadData_btnTransfer_Caption');
    lblTransferPath.Caption := dtmMain.getLanguageSettingInfo('frmLoadData_lblTransferPath_Caption');
    btnTransferSMS.Caption := dtmMain.getLanguageSettingInfo('frmLoadData_btnTransferSMS_Caption');
    btnTransferSMSSO005.Caption := dtmMain.getLanguageSettingInfo('frmLoadData_btnTransferSMSSO005_Caption');
    lblLogAndXmlPath.Caption := dtmMain.getLanguageSettingInfo('frmLoadData_lblLogAndXmlPath_Caption');

    {
    cobQueryType.Items.Add(dtmMain.getLanguageSettingInfo('frmLoadData_rgpQueryType_Item1_Caption'));
    cobQueryType.Items.Add(dtmMain.getLanguageSettingInfo('frmLoadData_rgpQueryType_Item2_Caption'));
    cobQueryType.Items.Add(dtmMain.getLanguageSettingInfo('frmLoadData_rgpQueryType_Item3_Caption'));
    cobQueryType.Items.Add(dtmMain.getLanguageSettingInfo('frmLoadData_rgpQueryType_Item4_Caption'));
    }
    chkICCard.Caption := dtmMain.getLanguageSettingInfo('frmLoadData_rgpQueryType_Item1_Caption');
    chkCAProduct.Caption := dtmMain.getLanguageSettingInfo('frmLoadData_rgpQueryType_Item2_Caption');
    chkProdPurchase.Caption := dtmMain.getLanguageSettingInfo('frmLoadData_rgpQueryType_Item3_Caption');
    chkEntitlement.Caption := dtmMain.getLanguageSettingInfo('frmLoadData_rgpQueryType_Item4_Caption');

end;

procedure TfrmLoadData.FormCreate(Sender: TObject);
begin
    setLanguageString;
    G_SGSCD024DataList := TStringList.Create;
end;

procedure TfrmLoadData.SpeedButton1Click(Sender: TObject);
var
    sL_Path : String;
begin
    if OpenDialog1.Execute then
    begin
      sL_Path := OpenDialog1.FileName;
      medICCardPath.Text := sL_Path;
    end;
end;

procedure TfrmLoadData.btnExitClick(Sender: TObject);
begin
    Close;
end;

procedure TfrmLoadData.btnICCardPathClick(Sender: TObject);
var
    sL_Path : String;
begin
    if OpenDialog1.Execute then
    begin
      sL_Path := OpenDialog1.FileName;
      medICCardPath.Text := sL_Path;
    end;
end;

function TfrmLoadData.isDataOK: Boolean;
var
    sL_CompCode,sL_QueryType,sL_ICCardPath,sL_ChannelPath : String;
    sL_CAProductPath,sL_ProdPurchasePath,sL_EntitlementPath,sL_TransferPath : String;
begin
  IsDataOk := true;

  if (not chkICCard.Checked) and (not chkCAProduct.Checked) and
     (not chkProdPurchase.Checked) and (not chkEntitlement.Checked) then
  begin
      MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_2_Content'),mtError, [mbOK],0);
      chkICCard.SetFocus;
      IsDataOk := false;
      exit;
  end;

  sL_CompCode := cobCompCode.Text;
  if sL_CompCode = '' then
  begin
      MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_1_Content'),mtError, [mbOK],0);
      cobCompCode.SetFocus;
      IsDataOk := false;
      exit;
  end;

{
  sL_QueryType := cobQueryType.Text;
  if sL_QueryType = '' then
  begin
      MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_2_Content'),mtError, [mbOK],0);
      cobQueryType.SetFocus;
      IsDataOk := false;
      exit;
  end;
}

  if chkICCard.Checked then
  begin
    sL_ICCardPath := Trim(medICCardPath.Text);
    if sL_ICCardPath = '' then
    begin
        MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_3_Content'),mtError, [mbOK],0);
        btnICCardPath.SetFocus;
        IsDataOk := false;
        exit;
    end;
  end;


  if chkCAProduct.Checked then
  begin
    sL_CAProductPath := Trim(medCAProductPath.Text);
    if sL_CAProductPath = '' then
    begin
        MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_3_Content'),mtError, [mbOK],0);
        btnCAProductPath.SetFocus;
        IsDataOk := false;
        exit;
    end;


    sL_ChannelPath := Trim(medChannelPath.Text);
    if sL_ChannelPath = '' then
    begin
        MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_3_Content'),mtError, [mbOK],0);
        btnChannelPath.SetFocus;
        IsDataOk := false;
        exit;
    end;
  end;


  if chkProdPurchase.Checked then
  begin
    sL_ProdPurchasePath := Trim(medProdPurchasePath.Text);
    if sL_ProdPurchasePath = '' then
    begin
        MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_3_Content'),mtError, [mbOK],0);
        btnProdPurchasePath.SetFocus;
        IsDataOk := false;
        exit;
    end;
  end;


  if chkEntitlement.Checked then
  begin
    sL_EntitlementPath := Trim(medEntitlementPath.Text);
    if sL_EntitlementPath = '' then
    begin
        MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_3_Content'),mtError, [mbOK],0);
        btnEntitlementPath.SetFocus;
        IsDataOk := false;
        exit;
    end;
  end;



    sL_TransferPath := Trim(medTransferPath.Text);
    if sL_TransferPath = '' then
    begin
        MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_3_Content'),mtError, [mbOK],0);
        btnTransferPath.SetFocus;
        IsDataOk := false;
        exit;
    end;


end;

procedure TfrmLoadData.btnLoadClick(Sender: TObject);
var
    sL_CompCode,sL_QueryType,sL_ModeType,sL_ReadStr : String;
    L_AllDataTxt,L_Master : TStringList;
    L_ChannelDataTxt,L_ChannelMaster : TStringList;
    ii,nL_ParamCounts,nL_ChannelParamCounts,nL_AreaDataNo : Integer;
    bL_IsOkDataFile : Boolean;
    dL_CurrTime : TTime;
    sL_ICCardPath,sL_ChannelPath : String;
    sL_CAProductPath,sL_ProdPurchasePath,sL_EntitlementPath : String;
begin
    if isDataOK then
    begin
      if MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_7_Content'), mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        stbStatus.Panels[0].Text := '';
        stbStatus.Panels[1].Text := '';
        stbStatus.Panels[2].Text := '';

        dL_CurrTime := now;
        bL_IsOkDataFile := true;
        sL_CompCode := getCobCompCode;


        sL_ICCardPath := Trim(medICCardPath.Text);
        sL_CAProductPath := Trim(medCAProductPath.Text);
        sL_ChannelPath := Trim(medChannelPath.Text);
        sL_ProdPurchasePath := Trim(medProdPurchasePath.Text);
        sL_EntitlementPath := Trim(medEntitlementPath.Text);
        nL_AreaDataNo := dtmMain.getAreaDataNo(sL_CompCode);


  //*************************************
  //1:IC卡資訊
        if chkICCard.Checked then
        begin
          showLogMsg('BLANK');
          showLogMsg('[ICCardData] Load File...' + medICCardPath.Text);

          if FileExists(sL_ICCardPath) then
          begin
            sL_ModeType := '1';
            L_AllDataTxt := TStringList.Create;
            L_Master := TStringList.Create;
            L_AllDataTxt.LoadFromFile(sL_ICCardPath);//將路徑下的文字檔轉到TStringList內

            //檢查該查詢種類的資料檔參數個數是否正確(只檢查第一筆資料)
            sL_ReadStr := L_AllDataTxt.Strings[0];
            L_Master := TUstr.ParseStrings(sL_ReadStr, TAB_STRING);
            nL_ParamCounts := L_Master.count;
            showLogMsg('[ICCardData] Check :  ' + medICCardPath.Text);
            bL_IsOkDataFile := checkIsOkDataFile(sL_ModeType,nL_ParamCounts,nL_ChannelParamCounts);

            if bL_IsOkDataFile then
            begin
              try
                //將程式包在 Transaction 交易中
                dtmMain.G_DbConnAry[nL_AreaDataNo].BeginTrans;

                showLogMsg('[ICCardData] Exec Import...');
                handleICCardQuery(L_AllDataTxt,sL_CompCode,dL_CurrTime);

                //若執行中沒有問題,則Commit
                dtmMain.G_DbConnAry[nL_AreaDataNo].CommitTrans;
              except
                //若執行中有問題,則Rollback
                if dtmMain.G_DbConnAry[nL_AreaDataNo].InTransaction then
                  dtmMain.G_DbConnAry[nL_AreaDataNo].RollbackTrans;
              end;
            end
            else
              showLogMsg('[ICCardData] Error File : ' + medICCardPath.Text);

            L_AllDataTxt.Free;
            L_Master.Free;
          end
          else
            showLogMsg('[ICCardData] File not Exist: ' + medICCardPath.Text);
        end;



  //*************************************
        //2:產品定義資訊
        if chkCAProduct.Checked then
        begin
          showLogMsg('BLANK');
          showLogMsg('[CAProductData] Load File...' + medCAProductPath.Text);

          if FileExists(sL_CAProductPath) then
          begin
            if FileExists(sL_ChannelPath) then
            begin
              sL_ModeType := '2';
              L_AllDataTxt := TStringList.Create;
              L_Master := TStringList.Create;
              L_AllDataTxt.LoadFromFile(sL_CAProductPath);//將路徑下的文字檔轉到TStringList內


              //檢查該查詢種類的資料檔參數個數是否正確(只檢查第一筆資料)
              sL_ReadStr := L_AllDataTxt.Strings[0];
              L_Master := TUstr.ParseStrings(sL_ReadStr, TAB_STRING);
              nL_ParamCounts := L_Master.count;


              L_ChannelDataTxt := TStringList.Create;
              L_ChannelMaster := TStringList.Create;
              L_ChannelDataTxt.LoadFromFile(sL_ChannelPath);//將路徑下的文字檔轉到TStringList內
              L_ChannelMaster := TUstr.ParseStrings(L_ChannelDataTxt.Strings[0], TAB_STRING);
              nL_ChannelParamCounts := L_ChannelMaster.count;


              showLogMsg('[CAProductData] Check :  ' + medCAProductPath.Text);
              bL_IsOkDataFile := checkIsOkDataFile(sL_ModeType,nL_ParamCounts,nL_ChannelParamCounts);

              if bL_IsOkDataFile then
              begin
                try
                  //將程式包在 Transaction 交易中
                  dtmMain.G_DbConnAry[nL_AreaDataNo].BeginTrans;

                  showLogMsg('[CAProductData] Exec Import...');
                  handleCAProductQuery(L_AllDataTxt,L_ChannelDataTxt,sL_CompCode,dL_CurrTime);

                  //若執行中沒有問題,則Commit
                  dtmMain.G_DbConnAry[nL_AreaDataNo].CommitTrans;
                except
                  //若執行中有問題,則Rollback
                  if dtmMain.G_DbConnAry[nL_AreaDataNo].InTransaction then
                    dtmMain.G_DbConnAry[nL_AreaDataNo].RollbackTrans;
                end;
              end
              else
                showLogMsg('[CAProductData] Error File : ' + medCAProductPath.Text);

              L_AllDataTxt.Free;
              L_Master.Free;
              L_ChannelDataTxt.Free;
              L_ChannelMaster.Free;
            end
            else
              showLogMsg('[CAProductData] File not Exist: ' + medChannelPath.Text);
          end
          else
            showLogMsg('[CAProductData] File not Exist: ' + medCAProductPath.Text);
        end;


  //*************************************
        //3:產品定購資訊
        if chkProdPurchase.Checked then
        begin
          showLogMsg('BLANK');
          showLogMsg('[ProdPurchaseData] Load File...' + medProdPurchasePath.Text);

          if FileExists(sL_ProdPurchasePath) then
          begin
            sL_ModeType := '3';
            L_AllDataTxt := TStringList.Create;
            L_Master := TStringList.Create;
            L_AllDataTxt.LoadFromFile(sL_ProdPurchasePath);//將路徑下的文字檔轉到TStringList內

            //檢查該查詢種類的資料檔參數個數是否正確(只檢查第一筆資料)
            sL_ReadStr := L_AllDataTxt.Strings[0];
            L_Master := TUstr.ParseStrings(sL_ReadStr, TAB_STRING);
            nL_ParamCounts := L_Master.count;
            showLogMsg('[ProdPurchaseData] Check :  ' + medProdPurchasePath.Text);
            bL_IsOkDataFile := checkIsOkDataFile(sL_ModeType,nL_ParamCounts,nL_ChannelParamCounts);

            if bL_IsOkDataFile then
            begin
              try
                //將程式包在 Transaction 交易中
                dtmMain.G_DbConnAry[nL_AreaDataNo].BeginTrans;

                showLogMsg('[ProdPurchaseData] Exec Import...');
                handleProdPurchaseQuery(L_AllDataTxt,sL_CompCode,dL_CurrTime);

                //若執行中沒有問題,則Commit
                dtmMain.G_DbConnAry[nL_AreaDataNo].CommitTrans;
              except
                //若執行中有問題,則Rollback
                if dtmMain.G_DbConnAry[nL_AreaDataNo].InTransaction then
                  dtmMain.G_DbConnAry[nL_AreaDataNo].RollbackTrans;
              end;
            end
            else
              showLogMsg('[ProdPurchaseData] Error File : ' + medProdPurchasePath.Text);

            L_AllDataTxt.Free;
            L_Master.Free;
          end
          else
            showLogMsg('[ProdPurchaseData] File not Exist: ' + medProdPurchasePath.Text);
        end;



  //*************************************
        //4:授權資訊
        if chkEntitlement.Checked then
        begin
          showLogMsg('BLANK');
          showLogMsg('[EntitlementData] Load File...' + medEntitlementPath.Text);

          if FileExists(sL_EntitlementPath) then
          begin
            sL_ModeType := '4';
            L_AllDataTxt := TStringList.Create;
            L_Master := TStringList.Create;
            L_AllDataTxt.LoadFromFile(sL_EntitlementPath);//將路徑下的文字檔轉到TStringList內

            //檢查該查詢種類的資料檔參數個數是否正確(只檢查第一筆資料)
            sL_ReadStr := L_AllDataTxt.Strings[0];
            L_Master := TUstr.ParseStrings(sL_ReadStr, TAB_STRING);
            nL_ParamCounts := L_Master.count;
            showLogMsg('[EntitlementData] Check :  ' + medEntitlementPath.Text);
            bL_IsOkDataFile := checkIsOkDataFile(sL_ModeType,nL_ParamCounts,nL_ChannelParamCounts);

            if bL_IsOkDataFile then
            begin
              try
                //將程式包在 Transaction 交易中
                dtmMain.G_DbConnAry[nL_AreaDataNo].BeginTrans;

                showLogMsg('[EntitlementData] Exec Import...');
                handleEntitlementQuery(L_AllDataTxt,sL_CompCode,dL_CurrTime);

                //若執行中沒有問題,則Commit
                dtmMain.G_DbConnAry[nL_AreaDataNo].CommitTrans;
              except
                //若執行中有問題,則Rollback
                if dtmMain.G_DbConnAry[nL_AreaDataNo].InTransaction then
                  dtmMain.G_DbConnAry[nL_AreaDataNo].RollbackTrans;
              end;

            end
            else
              showLogMsg('[EntitlementData] Error File : ' + medEntitlementPath.Text);

            L_AllDataTxt.Free;
            L_Master.Free;
          end
          else
            showLogMsg('[EntitlementData] File not Exist: ' + medEntitlementPath.Text);
        end;


  {
        //1:IC卡資訊2:產品定義資訊3:產品定購資訊4:授權資訊
        if cobQueryType.ItemIndex = 0 then
          sL_ModeType := '1'
        else if cobQueryType.ItemIndex = 1 then
        begin
          sL_ModeType := '2';

          L_ChannelDataTxt := TStringList.Create;
          L_ChannelMaster := TStringList.Create;
          L_ChannelDataTxt.LoadFromFile(sL_ChannelPath);//將路徑下的文字檔轉到TStringList內

          L_ChannelMaster := TUstr.ParseStrings(L_ChannelDataTxt.Strings[0], TAB_STRING);
          nL_ChannelParamCounts := L_ChannelMaster.count;
        end
        else if cobQueryType.ItemIndex = 2 then
          sL_ModeType := '3'
        else if cobQueryType.ItemIndex = 3 then
          sL_ModeType := '4';

        L_AllDataTxt := TStringList.Create;
        L_Master := TStringList.Create;
        L_AllDataTxt.LoadFromFile(sL_DataPath);//將路徑下的文字檔轉到TStringList內


        //檢查該查詢種類的資料檔參數個數是否正確(只檢查第一筆資料)
        sL_ReadStr := L_AllDataTxt.Strings[0];
        L_Master := TUstr.ParseStrings(sL_ReadStr, TAB_STRING);
        nL_ParamCounts := L_Master.count;
        bL_IsOkDataFile := checkIsOkDataFile(sL_ModeType,nL_ParamCounts,nL_ChannelParamCounts);


        if bL_IsOkDataFile then
        begin

          if sL_ModeType = '1' then
            handleICCardQuery(L_AllDataTxt,sL_CompCode,dL_CurrTime)

          else if sL_ModeType = '2' then
            handleCAProductQuery(L_AllDataTxt,L_ChannelDataTxt,sL_CompCode,dL_CurrTime)

          else if sL_ModeType = '3' then
            handleProdPurchaseQuery(L_AllDataTxt,sL_CompCode,dL_CurrTime)

          else if sL_ModeType = '4' then
            handleEntitlementQuery(L_AllDataTxt,sL_CompCode,dL_CurrTime);

        end;
        L_AllDataTxt.Free;
        L_Master.Free;
  }
        //showmessage('OK');
      end;
    end;

end;

function TfrmLoadData.checkIsOkDataFile(sI_ModeType: String;
  nI_ParamCounts,nI_ChannelParamCounts : Integer): Boolean;
begin
    //檢查該查詢種類的資料檔參數個數是否正確
    {
    sms_iccard.txt       ---->>> 7 個參數
    sms_caproduct.txt    ---->>> 4 個參數
    sms_prodpurchase.txt ---->>> 9 個參數
    sms_entitlement.txt  ---->>> 6 個參數
    sms_channel.txt      ---->>> 3 個參數
    }

    Result := true;

    if sI_ModeType =  '1' then
    begin
      if nI_ParamCounts <> 7 then
      begin
        //MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_4_Content'),mtError, [mbOK],0);
        Result := false;
      end;
    end
    else if sI_ModeType =  '2' then
    begin
      if nI_ParamCounts <> 4 then
      begin
        //MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_4_Content'),mtError, [mbOK],0);
        Result := false;
      end
      else
      begin
        if nI_ChannelParamCounts <> 3 then
        begin
          //MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_6_Content'),mtError, [mbOK],0);
          Result := false;
        end;
      end;
    end
    else if sI_ModeType =  '3' then
    begin
      if nI_ParamCounts <> 9 then
      begin
        //MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_4_Content'),mtError, [mbOK],0);
        Result := false;
      end;
    end
    else if sI_ModeType =  '4' then
    begin
      if nI_ParamCounts <> 6 then
      begin
        //MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_4_Content'),mtError, [mbOK],0);
        Result := false;
      end;
    end;

end;

function TfrmLoadData.handleICCardQuery(I_AllDataTxt: TStringList;
  sI_CompCode: String;dI_CurrTime : TTime): String;
var
    nL_CurLineNo,ii : Integer;
    L_Master : TStringList;
    sL_HandleEDateTime,sL_RealDateTime,sL_ExecDateTime,sL_Status : String;
    sL_ICCardNo,sL_ICCardID,sL_CasID,sL_Feature,sL_Version,sL_ReadStr : String;
    dL_DateTime2,dL_DiffDateTime : TTime;
    sL_AllCounts,sL_SQL,sL_DiffDateTime  : String;
begin
    L_Master := TStringList.Create;
    sL_AllCounts := IntTOsTR(I_AllDataTxt.Count);
    sL_ExecDateTime := TUdateTime.GetPureDateTimeStr(NOW,'/',':');

    //先清空原來的資料
    sL_SQL := 'DELETE SO451';
    dtmMain.runSQL(IUD_MODE,sI_CompCode,sL_SQL);

    for nL_CurLineNo:=0 to I_AllDataTxt.Count-1 do
    begin

      sL_ReadStr := I_AllDataTxt.Strings[nL_CurLineNo];
      L_Master := TUstr.ParseStrings(sL_ReadStr, TAB_STRING);

      //文字檔參數順序(1)對IC卡進行操作的時間(2)New為創建IC卡,Delete為刪除(3)IC卡號
      //              (4)IC卡內部卡號(5)CASID(6)IC卡特性(7)版本
      sL_RealDateTime := Trim(TUstr.replaceStr(L_Master.Strings[0],'-','/'));

      sL_Status := Trim(L_Master.Strings[1]);
      if Uppercase(sL_Status) = 'NEW' then
         sL_Status := '0'
      else if Uppercase(sL_Status) = 'DELETE' then
         sL_Status := '1'
      else
         sL_Status := '';

      sL_ICCardNo := Trim(L_Master.Strings[2]);
      sL_ICCardID := Trim(L_Master.Strings[3]);
      sL_CasID := Trim(L_Master.Strings[4]);
      sL_Feature := Trim(L_Master.Strings[5]);
      sL_Version := Trim(L_Master.Strings[6]);

      sL_HandleEDateTime := Copy(sL_RealDateTime,0,10) + ' 23:59:59';


      sL_SQL := 'INSERT INTO SO451(CompCode,HandleEDateTime,RealDateTime,ExecDateTime,' +
                'Status,ICCardNo,ICCardID,CasID,Feature,Version) VALUES(' + sI_CompCode + ',' +
                TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_HandleEDateTime)) + ',' +
                TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_RealDateTime)) + ',' +
                TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_ExecDateTime)) + ',' +
                sL_Status + ',''' + sL_ICCardNo + ''',''' +  sL_ICCardID + ''',''' +
                sL_CasID + ''',''' + sL_Feature + ''',''' + sL_Version +''')';

      dtmMain.runSQL(IUD_MODE,sI_CompCode,sL_SQL);


      if (nL_CurLineNo Mod SHOW_UI_LINE_COUNTS)=0 then
      begin
        stbStatus.Panels[1].Text := IntToStr((nL_CurLineNo+1)) + '/' + sL_AllCounts;
        stbStatus.Refresh;
      end;
    end;

    //顯示剩餘筆數
    stbStatus.Panels[1].Text := IntToStr((nL_CurLineNo)) + '/' + sL_AllCounts;
    stbStatus.Refresh;

    dL_DateTime2 := now;
    dL_DiffDateTime := dL_DateTime2 - dI_CurrTime;
    sL_DiffDateTime := TimeToStr(dL_DiffDateTime);
    L_Master.Free;

    showLogMsg('[ICCardData] Exec Import OK-[Total Time:' + sL_DiffDateTime + ']');
end;

function TfrmLoadData.handleCAProductQuery(I_AllDataTxt,L_ChannelDataTxt: TStringList;
  sI_CompCode: String;dI_CurrTime : TTime): String;
var
    nL_CurLineNo,ii : Integer;
    L_Master,L_ChannelMaster : TStringList;
    sL_HandleEDateTime,sL_RealDateTime,sL_ExecDateTime : String;
    dL_DateTime2,dL_DiffDateTime : TTime;
    sL_CasID,sL_ReadStr,sL_AllCounts,sL_SQL,sL_DiffDateTime  : String;
    sL_ProductCode,sL_ProductName,sL_Provider,sL_ChannelName : String;
begin
    sL_ExecDateTime := TUdateTime.GetPureDateTimeStr(NOW,'/',':');

    //先將產品頻道對應存至CDS
    for ii:=0 to L_ChannelDataTxt.Count-1 do
    begin
      //文字檔參數順序(1)產品代碼(2)節目平臺名稱(3)原始頻道名稱
      L_ChannelMaster := TStringList.Create;
      L_ChannelMaster := TUstr.ParseStrings(L_ChannelDataTxt.Strings[ii], TAB_STRING);
      sL_ProductCode := L_ChannelMaster.Strings[0];
      sL_Provider := L_ChannelMaster.Strings[1];
      sL_ChannelName := L_ChannelMaster.Strings[2];

      dtmMain.cdsChannelID.Append;
      dtmMain.cdsChannelID.FieldByName('ProductCode').AsString := sL_ProductCode;
      dtmMain.cdsChannelID.FieldByName('Provider').AsString := sL_Provider;
      dtmMain.cdsChannelID.FieldByName('ChannelName').AsString := sL_ChannelName;
      dtmMain.cdsChannelID.Post;
      L_ChannelMaster.Free;
    end;

    L_Master := TStringList.Create;
    sL_AllCounts := IntTOsTR(I_AllDataTxt.Count);


    //先清空原來的資料
    sL_SQL := 'DELETE SO453';
    dtmMain.runSQL(IUD_MODE,sI_CompCode,sL_SQL);

    for nL_CurLineNo:=0 to I_AllDataTxt.Count-1 do
    begin

      sL_ReadStr := I_AllDataTxt.Strings[nL_CurLineNo];
      L_Master := TUstr.ParseStrings(sL_ReadStr, TAB_STRING);

      //文字檔參數順序(1)增加產品操作時間(2)產品代碼(3)產品名稱(4)CASID
      sL_RealDateTime := Trim(TUstr.replaceStr(L_Master.Strings[0],'-','/'));

      sL_ProductCode := Trim(L_Master.Strings[1]);
      sL_ProductName := Trim(L_Master.Strings[2]);
      sL_CasID := Trim(L_Master.Strings[3]);

      sL_Provider := '';
      sL_ChannelName := '';
      if dtmMain.cdsChannelID.Locate('ProductCode',VarArrayOf([sL_ProductCode]) , []) then
      begin
        sL_Provider := dtmMain.cdsChannelID.fieldbyname('Provider').asstring;
        sL_ChannelName := dtmMain.cdsChannelID.fieldbyname('ChannelName').asstring;
      end;

      sL_HandleEDateTime := Copy(sL_RealDateTime,0,10) + ' 23:59:59';


      sL_SQL := 'INSERT INTO SO453(CompCode,HandleEDateTime,RealDateTime,ExecDateTime,' +
                'ProductCode,ProductName,CasID,Provider,ChannelName) VALUES(' + sI_CompCode + ',' +
                TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_HandleEDateTime)) + ',' +
                TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_RealDateTime)) + ',' +
                TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_ExecDateTime)) + ',''' +
                sL_ProductCode + ''',''' +  sL_ProductName + ''',''' +
                sL_CasID + ''',''' + sL_Provider + ''',''' + sL_ChannelName +''')';

      dtmMain.runSQL(IUD_MODE,sI_CompCode,sL_SQL);

      if (nL_CurLineNo Mod SHOW_UI_LINE_COUNTS)=0 then
      begin
        stbStatus.Panels[1].Text := IntToStr((nL_CurLineNo+1)) + '/' + sL_AllCounts;
        stbStatus.Refresh;
      end;
    end;

    //顯示剩餘筆數
    stbStatus.Panels[1].Text := IntToStr((nL_CurLineNo)) + '/' + sL_AllCounts;
    stbStatus.Refresh;

    dL_DateTime2 := now;
    dL_DiffDateTime := dL_DateTime2 - dI_CurrTime;
    sL_DiffDateTime := TimeToStr(dL_DiffDateTime);
    L_Master.Free;

    showLogMsg('[CAProductData] Exec Import OK-[Total Time:' + sL_DiffDateTime + ']');
end;

function TfrmLoadData.handleProdPurchaseQuery(I_AllDataTxt: TStringList;
  sI_CompCode: String;dI_CurrTime : TTime): String;
var
    nL_CurLineNo,ii,nL_HandleCounts : Integer;
    L_Master,L_ProductList : TStringList;
    sL_HandleEDateTime,sL_RealDateTime,sL_ExecDateTime,sL_Status : String;
    sL_ICCardNo,sL_ICCardID,sL_CasID,sL_ProductCode,sL_ProductCodeList,sL_ReadStr : String;
    dL_DateTime2,dL_DiffDateTime : TTime;
    sL_AllCounts,sL_SQL,sL_DiffDateTime,sL_PurchaseID,sL_StartTime,sL_EndTime  : String;
begin
    nL_HandleCounts := 0;
    L_Master := TStringList.Create;
    L_ProductList := TStringList.Create;
    sL_AllCounts := IntTOsTR(I_AllDataTxt.Count);
    sL_ExecDateTime := TUdateTime.GetPureDateTimeStr(NOW,'/',':');

    //先清空原來的資料
    sL_SQL := 'DELETE SO455';
    dtmMain.runSQL(IUD_MODE,sI_CompCode,sL_SQL);

    for nL_CurLineNo:=0 to I_AllDataTxt.Count-1 do
    begin

      sL_ReadStr := I_AllDataTxt.Strings[nL_CurLineNo];
      L_Master := TUstr.ParseStrings(sL_ReadStr, TAB_STRING);

      //文字檔參數順序(1)產品購買操作時間(2)MODIFY新加/修改定購,Delete刪除定購(3)IC卡號
      //              (4)IC卡內部卡號(5)CASID(6)購買產品代碼組(7)定購ID
      //              (8)生效時間(9)失效時間
      sL_RealDateTime := Trim(TUstr.replaceStr(L_Master.Strings[0],'-','/'));

      sL_Status := Trim(L_Master.Strings[1]);
      if (Uppercase(sL_Status) = 'MODIFY') OR (Uppercase(sL_Status) = 'CREATE') then
         sL_Status := '0'
      else if Uppercase(sL_Status) = 'DELETE' then
         sL_Status := '1'
      else
         sL_Status := '';

      sL_ICCardNo := Trim(L_Master.Strings[2]);
      sL_ICCardID := Trim(L_Master.Strings[3]);
      sL_CasID := Trim(L_Master.Strings[4]);
      sL_ProductCodeList := Trim(L_Master.Strings[5]);
      sL_PurchaseID := Trim(L_Master.Strings[6]);
      sL_StartTime := Trim(TUstr.replaceStr(L_Master.Strings[7],'-','/'));
      sL_EndTime := Trim(TUstr.replaceStr(L_Master.Strings[8],'-','/'));

      sL_HandleEDateTime := Copy(sL_RealDateTime,0,10) + ' 23:59:59';


      L_ProductList := TUstr.ParseStrings(sL_ProductCodeList,',');
      for ii:=0 to L_ProductList.Count-1 do
      begin
        nL_HandleCounts := nL_HandleCounts + 1;
        sL_ProductCode := L_ProductList.Strings[ii];

        sL_SQL := 'INSERT INTO SO455(CompCode,HandleEDateTime,RealDateTime,ExecDateTime,' +
                  'Status,PurchaseID,ICCardNo,ICCardID,CasID,ProductCode,StartTime,EndTime) VALUES(' + sI_CompCode + ',' +
                  TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_HandleEDateTime)) + ',' +
                  TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_RealDateTime)) + ',' +
                  TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_ExecDateTime)) + ',' +
                  sL_Status + ',''' + sL_PurchaseID + ''',''' + sL_ICCardNo + ''',''' +  sL_ICCardID + ''',''' +
                  sL_CasID + ''',''' + sL_ProductCode + ''',' +
                  TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_StartTime)) + ',' +
                  TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_EndTime)) + ')';

        dtmMain.runSQL(IUD_MODE,sI_CompCode,sL_SQL);


        if (nL_HandleCounts Mod SHOW_UI_COUNTS)=0 then
        begin
          stbStatus.Panels[0].Text := IntToStr((nL_HandleCounts));
          stbStatus.Refresh;
        end;
      end;

      if (nL_CurLineNo Mod SHOW_UI_LINE_COUNTS)=0 then
      begin
        stbStatus.Panels[1].Text := IntToStr((nL_CurLineNo+1)) + '/' + sL_AllCounts;
        stbStatus.Refresh;
      end;
    end;


    //顯示剩餘筆數
    stbStatus.Panels[0].Text := IntToStr((nL_HandleCounts));
    stbStatus.Refresh;
    stbStatus.Panels[1].Text := IntToStr((nL_CurLineNo)) + '/' + sL_AllCounts;
    stbStatus.Refresh;


    dL_DateTime2 := now;
    dL_DiffDateTime := dL_DateTime2 - dI_CurrTime;
    sL_DiffDateTime := TimeToStr(dL_DiffDateTime);
    L_Master.Free;
    L_ProductList.free;

    showLogMsg('[ProdPurchaseData] Exec Import OK-[Total Time:' + sL_DiffDateTime + ']');
end;

function TfrmLoadData.handleEntitlementQuery(I_AllDataTxt: TStringList;
  sI_CompCode: String;dI_CurrTime : TTime): String;
var
    nL_CurLineNo,ii,nL_HandleCounts : Integer;
    L_Master,L_ProductList : TStringList;
    sL_HandleEDateTime,sL_RealDateTime,sL_ExecDateTime,sL_Status : String;
    sL_ICCardNo,sL_ICCardID,sL_CasID,sL_ProductCode,sL_ProductCodeList,sL_ReadStr : String;
    dL_DateTime2,dL_DiffDateTime : TTime;
    sL_AllCounts,sL_SQL,sL_DiffDateTime  : String;
begin
    nL_HandleCounts := 0;
    L_Master := TStringList.Create;
    L_ProductList := TStringList.Create;
    sL_AllCounts := IntTOsTR(I_AllDataTxt.Count);
    sL_ExecDateTime := TUdateTime.GetPureDateTimeStr(NOW,'/',':');

    //先清空原來的資料
    sL_SQL := 'DELETE SO457';
    dtmMain.runSQL(IUD_MODE,sI_CompCode,sL_SQL);

    for nL_CurLineNo:=0 to I_AllDataTxt.Count-1 do
    begin
      sL_ReadStr := I_AllDataTxt.Strings[nL_CurLineNo];
      L_Master := TUstr.ParseStrings(sL_ReadStr, TAB_STRING);

      //文字檔參數順序(1)授權操作的時間(2)Create為授權,Delete為反授權(3)IC卡號
      //              (4)IC卡內部卡號(5)CASID(6)授權產品代碼組
      sL_RealDateTime := Trim(TUstr.replaceStr(L_Master.Strings[0],'-','/'));

      sL_Status := Trim(L_Master.Strings[1]);
      if Uppercase(sL_Status) = 'CREATE' then
         sL_Status := '0'
      else if Uppercase(sL_Status) = 'DELETE' then
         sL_Status := '1'
      else
         sL_Status := '';

      sL_ICCardNo := Trim(L_Master.Strings[2]);
      sL_ICCardID := Trim(L_Master.Strings[3]);
      sL_CasID := Trim(L_Master.Strings[4]);
      sL_ProductCodeList := Trim(L_Master.Strings[5]);

      sL_HandleEDateTime := Copy(sL_RealDateTime,0,10) + ' 23:59:59';


      L_ProductList := TUstr.ParseStrings(sL_ProductCodeList,',');
      for ii:=0 to L_ProductList.Count-1 do
      begin
        nL_HandleCounts := nL_HandleCounts + 1;
        sL_ProductCode := L_ProductList.Strings[ii];

        sL_SQL := 'INSERT INTO SO457(CompCode,HandleEDateTime,RealDateTime,ExecDateTime,' +
                  'Status,ICCardNo,ICCardID,CasID,ProductCode) VALUES(' + sI_CompCode + ',' +
                  TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_HandleEDateTime)) + ',' +
                  TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_RealDateTime)) + ',' +
                  TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_ExecDateTime)) + ',' +
                  sL_Status + ',''' + sL_ICCardNo + ''',''' +  sL_ICCardID + ''',''' +
                  sL_CasID + ''',''' + sL_ProductCode + ''')';

        dtmMain.runSQL(IUD_MODE,sI_CompCode,sL_SQL);

        if (nL_HandleCounts Mod SHOW_UI_COUNTS)=0 then
        begin
          stbStatus.Panels[0].Text := IntToStr((nL_HandleCounts));
          stbStatus.Refresh;
        end;

      end;

      if (nL_CurLineNo Mod SHOW_UI_LINE_COUNTS)=0 then
      begin
        stbStatus.Panels[1].Text := IntToStr((nL_CurLineNo+1)) + '/' + sL_AllCounts;
        stbStatus.Refresh;
      end;
    end;


    //顯示剩餘筆數
    stbStatus.Panels[0].Text := IntToStr((nL_HandleCounts));
    stbStatus.Refresh;
    stbStatus.Panels[1].Text := IntToStr((nL_CurLineNo)) + '/' + sL_AllCounts;
    stbStatus.Refresh;

    dL_DateTime2 := now;
    dL_DiffDateTime := dL_DateTime2 - dI_CurrTime;
    sL_DiffDateTime := TimeToStr(dL_DiffDateTime);
    L_Master.Free;
    L_ProductList.free;

    showLogMsg('[EntitlementData] Exec Import OK-[Total Time:' + sL_DiffDateTime + ']');
end;

function TfrmLoadData.getCobCompCode: String;
var
    L_CompCodeStrList : TStringList;
    sL_CompCodeAndName : String;
begin
    //取得所屬公司
    sL_CompCodeAndName := cobCompCode.Text;

    if sL_CompCodeAndName <> '' then
    begin
      L_CompCodeStrList := TStringList.Create;
      L_CompCodeStrList := TUstr.ParseStrings(cobCompCode.Text,'-');
      Result := L_CompCodeStrList.Strings[0];
    end
    else
      Result := '';

    L_CompCodeStrList.Free;
end;

procedure TfrmLoadData.btnChannelPathClick(Sender: TObject);
var
    sL_Path : String;
begin
    if OpenDialog1.Execute then
    begin
      sL_Path := OpenDialog1.FileName;
      medChannelPath.Text := sL_Path;
    end;
end;

procedure TfrmLoadData.btnXmlClick(Sender: TObject);
var
    sL_CompCode,sL_QueryType,sL_ModeType,sL_DiffDateTime : String;
    sL_StartDate,sL_StopDate : String;
    dL_CurrTime,dL_DateTime2,dL_DiffDateTime : TTime;
begin
    //產生Log及Xml
    if isDataOK2 then
    begin
      stbStatus.Panels[0].Text := '';
      stbStatus.Panels[1].Text := '';
      stbStatus.Panels[2].Text := '';

      sL_CompCode := getCobCompCode;
      sL_StartDate := Trim(fraYMD3.getYMD) + ' 00:00:00';
      sL_StopDate := Trim(fraYMD4.getYMD) + ' 23:59:59';

//***********************************
      //1:IC卡資訊
      if chkICCard.Checked then
      begin
        dL_CurrTime := now;
        showLogMsg('BLANK');

        sL_ModeType := '1';

        //補填 LoadData 的 Log
        showLogMsg('[ICCardData] Log Data...');
        handleLoadDataLog(sL_CompCode,sL_ModeType,sL_StartDate,sL_StopDate);


        //產生 handleLoadData 的 XML
        showLogMsg('[ICCardData] Produce Xml...');
        handleLoadDataXml(sL_CompCode,sL_ModeType,sL_StartDate,sL_StopDate);


        dL_DateTime2 := now;
        dL_DiffDateTime := dL_DateTime2 - dL_CurrTime;
        sL_DiffDateTime := TimeToStr(dL_DiffDateTime);

        showLogMsg('[ICCardData] OK-[Total Time:' + sL_DiffDateTime + ']');
      end;


//***********************************
      //2:產品定義資訊
      if chkCAProduct.Checked then
      begin
        dL_CurrTime := now;
        showLogMsg('BLANK');

        sL_ModeType := '2';

        //補填 LoadData 的 Log
        showLogMsg('[CAProductData] Log Data...');
        dtmMain.runTestDataInsertSO453SP(sL_CompCode);

        handleLoadDataLog(sL_CompCode,sL_ModeType,sL_StartDate,sL_StopDate);


        //產生 handleLoadData 的 XML
        showLogMsg('[CAProductData] Produce Xml...');
        handleLoadDataXml(sL_CompCode,sL_ModeType,sL_StartDate,sL_StopDate);

        dL_DateTime2 := now;
        dL_DiffDateTime := dL_DateTime2 - dL_CurrTime;
        sL_DiffDateTime := TimeToStr(dL_DiffDateTime);

        showLogMsg('[CAProductData] OK-[Total Time:' + sL_DiffDateTime + ']');
      end;


//***********************************
      //3:產品定購資訊
      if chkProdPurchase.Checked then
      begin
        dL_CurrTime := now;
        showLogMsg('BLANK');

        sL_ModeType := '3';

        //補填 LoadData 的 Log
        showLogMsg('[ProdPurchaseData] Log Data...');
        handleLoadDataLog(sL_CompCode,sL_ModeType,sL_StartDate,sL_StopDate);


        //產生 handleLoadData 的 XML
        showLogMsg('[ProdPurchaseData] Produce Xml...');
        handleLoadDataXml(sL_CompCode,sL_ModeType,sL_StartDate,sL_StopDate);

        dL_DateTime2 := now;
        dL_DiffDateTime := dL_DateTime2 - dL_CurrTime;
        sL_DiffDateTime := TimeToStr(dL_DiffDateTime);

        showLogMsg('[ProdPurchaseData] OK-[Total Time:' + sL_DiffDateTime + ']');
      end;


//***********************************
      //4:授權資訊
      if chkEntitlement.Checked then
      begin
        dL_CurrTime := now;
        showLogMsg('BLANK');

        sL_ModeType := '4';

        //補填 LoadData 的 Log
        showLogMsg('[EntitlementData] Log Data...');
        handleLoadDataLog(sL_CompCode,sL_ModeType,sL_StartDate,sL_StopDate);


        //產生 handleLoadData 的 XML
        showLogMsg('[EntitlementData] Produce Xml...');
        handleLoadDataXml(sL_CompCode,sL_ModeType,sL_StartDate,sL_StopDate);

        dL_DateTime2 := now;
        dL_DiffDateTime := dL_DateTime2 - dL_CurrTime;
        sL_DiffDateTime := TimeToStr(dL_DiffDateTime);

        showLogMsg('[EntitlementData] OK-[Total Time:' + sL_DiffDateTime + ']');
      end;
    end;
end;

function TfrmLoadData.isDataOK2: Boolean;
var
    sL_CompCode,sL_QueryType,sL_StartDate,sL_StopDate : String;
begin
  IsDataOk2 := true;

  if (not chkICCard.Checked) and (not chkCAProduct.Checked) and
     (not chkProdPurchase.Checked) and (not chkEntitlement.Checked) then
  begin
      MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_2_Content'),mtError, [mbOK],0);
      chkICCard.SetFocus;
      IsDataOk2 := false;
      exit;
  end;

  sL_CompCode := cobCompCode.Text;
  if sL_CompCode = '' then
  begin
      MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_1_Content'),mtError, [mbOK],0);
      cobCompCode.SetFocus;
      IsDataOk2 := false;
      exit;
  end;


  sL_StartDate := Trim(TUstr.replaceStr(Trim(fraYMD3.getYMD),'/',''));
  sL_StopDate := Trim(TUstr.replaceStr(Trim(fraYMD4.getYMD),'/',''));


  if sL_StartDate = '' then
  begin
    MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_8_Content'),mtError, [mbOK],0);
    fraYMD3.setYMD('');
    fraYMD3.mseYMD.SetFocus;
    IsDataOk2 := false;
    exit;
  end;

  if sL_StopDate = '' then
  begin
    MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_8_Content'),mtError, [mbOK],0);
    fraYMD4.setYMD('');
    fraYMD4.mseYMD.SetFocus;
    IsDataOk2 := false;
    exit;
  end;



  if StrToInt(sL_StartDate) > StrToInt(sL_StopDate) then
  begin
    MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_9_Content'),mtError, [mbOK],0);
    fraYMD4.setYMD('');
    fraYMD4.mseYMD.SetFocus;
    IsDataOk2 := false;
    exit;
  end;

  IsDataOk2 := true;
end;

function TfrmLoadData.handleLoadDataLog(sI_CompCode,
  sI_ModeType,sI_StartDate,sI_StopDate: String): String;
var
    sL_SQL,sL_Table,sL_HandleEDateTime : String;
    sL_FirstFlag,sL_RetCode,sL_Msg,sL_FirstDate,sL_TempFirstDate : String;
    sL_TotalICCardNum,sL_TotalSubNum,sL_TotalProductNum,sL_TotalChannelNum : String;
    L_AdoQry : TADOQuery;
    nL_TotalHandleCounts,nL_HandleCounts,nL_LogCounts,nL_TotalLogCounts : Integer;
    nL_TotalICCardNum,nL_TotalSubNum,nL_TotalProductNum,nL_TotalChannelNum : Integer;
begin
    //1:IC卡資訊2:產品定義資訊3:產品定購資訊4:授權資訊
    if sI_ModeType = '1' then
      sL_Table := 'SO451'
    else if sI_ModeType = '2' then
      sL_Table := 'SO453'
    else if sI_ModeType = '3' then
      sL_Table := 'SO455'
    else if sI_ModeType = '4' then
      sL_Table := 'SO457';

    sL_SQL := 'SELECT HandleEDateTime,Count(HandleEDateTime) HandleCounts from ' + sL_Table +
              ' Group by HandleEDateTime Having HandleEDateTime between ' +
              TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_StartDate)) + ' and ' +
              TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_StopDate)) ;


    L_AdoQry := dtmMain.runSQL(SELECT_MODE,sI_CompCode,sL_SQL);
    dtmMain.adoQryLoadExecCounts.Clone(L_AdoQry,ltUnspecified);
    sL_FirstDate := dtmMain.getFirstDate(sI_CompCode);


    if sI_ModeType = '1' then
    begin
      //取得前一天 ICCard 的總數量
      nL_TotalHandleCounts := dtmMain.getYesterdayTotalICCardNum(sL_Table,sI_CompCode,sI_StartDate);
    end;

    nL_LogCounts := 0;
    nL_TotalLogCounts := dtmMain.adoQryLoadExecCounts.RecordCount;
    with dtmMain.adoQryLoadExecCounts do
    begin
      First;
      while not Eof do
      begin
        sL_HandleEDateTime := TUstr.replaceStr(dtmMain.adoQryLoadExecCounts.FieldByName('HandleEDateTime').AsString,'-','/');
        nL_HandleCounts := dtmMain.adoQryLoadExecCounts.FieldByName('HandleCounts').AsInteger;
        nL_TotalHandleCounts := nL_TotalHandleCounts + nL_HandleCounts;

        sL_TempFirstDate := Copy(sL_HandleEDateTime,0,10);
        if sL_TempFirstDate = sL_FirstDate then
          sL_FirstFlag := '1'
        else
          sL_FirstFlag := '0';

        //1:IC卡資訊2:產品定義資訊3:產品定購資訊4:授權資訊
        if sI_ModeType = '1' then
        begin
          dtmMain.getICCardNumAndSubNum(true,sI_CompCode,sL_HandleEDateTime,nL_TotalICCardNum,nL_TotalSubNum);

          sL_TotalICCardNum := IntToStr(nL_TotalHandleCounts);  //以測評資料為資料
          sL_TotalSubNum := IntToStr(nL_TotalSubNum);   //以實際SO004當前整個SMS內用戶總數
          sL_TotalProductNum := '';
          sL_TotalChannelNum := '';
        end
        else if sI_ModeType = '2' then
        begin
          sL_TotalICCardNum := '';
          sL_TotalSubNum := '';
          dtmMain.getProductNumAndChannelNum(sI_CompCode,sL_HandleEDateTime,nL_TotalProductNum,nL_TotalChannelNum);
          sL_TotalProductNum := IntToStr(nL_TotalProductNum);
          sL_TotalChannelNum := IntToStr(nL_TotalChannelNum);
        end
        else
        begin
          sL_TotalICCardNum := '';
          sL_TotalSubNum := '';
          sL_TotalProductNum := '';
          sL_TotalChannelNum := '';
        end;


        sL_SQL := 'DELETE SO459 where HandleEDateTime=' +
                  TUstr.getOracleSQLDateTimeStr(StrToDateTime(sL_HandleEDateTime)) +
                  ' AND ModeType=' + sI_ModeType;
        dtmMain.runSQL(IUD_MODE,sI_CompCode,sL_SQL);

        //SO459存原Oracle的錯誤訊息
        dtmMain.logToSO459(sI_CompCode,sI_ModeType,sL_FirstFlag,sL_HandleEDateTime,
                           sL_RetCode,sL_Msg,sL_TotalICCardNum,sL_TotalSubNum,
                           sL_TotalProductNum,sL_TotalChannelNum);


        nL_LogCounts := nL_LogCounts + 1;
        stbStatus.Panels[1].Text := IntToStr(nL_LogCounts) + '/' + IntToStr(nL_TotalLogCounts);
        stbStatus.Refresh;                           

        Next;
      end;
    end;
end;

function TfrmLoadData.handleLoadDataXml(sI_CompCode,
  sI_ModeType,sI_StartDate,sI_StopDate : String) : String;
var
    sL_SQL,sL_Table,sL_HandleEDateTime,sL_FirstFlag,sL_RetCode,sL_Msg : String;
    sL_Version,sL_SrcCode,sL_SrcType,sL_DstCode,sL_DstType,sL_CasID : String;
    sL_TotalICCardNum,sL_TotalSubNum,sL_TotalProductNum,sL_TotalChannelNum : String;
    sL_DstMsgID,sL_InstQuery,sL_QueryDateTime,sL_QueryType,sL_CmdSeq,sL_SrcMsgID : String;
    sL_FirstDate : String;
    L_AdoQry : TADOQuery;
    nL_TotalHandleCounts,nL_HandleCounts,nL_Ndx,nL_TotalExecCounts,nL_XmlCounts,nL_TotalXmlCounts : Integer;
    nL_TotalICCardNum,nL_TotalSubNum,nL_TotalProductNum,nL_TotalChannelNum : Integer;
begin

    //1:IC卡資訊2:產品定義資訊3:產品定購資訊4:授權資訊
    if sI_ModeType = '1' then
      sL_Table := 'SO451'
    else if sI_ModeType = '2' then
      sL_Table := 'SO453'
    else if sI_ModeType = '3' then
      sL_Table := 'SO455'
    else if sI_ModeType = '4' then
      sL_Table := 'SO457';
//*************************************




    //取出相關參數
    with dtmMain do
    begin
      nL_Ndx := G_SGSParamIniStrList.IndexOf(sI_CompCode);

      //nDataAreaNo := (G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).nDataAreaNo;
      //sL_CompName := (G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sCompName;
      //sL_DataPath1 := (G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sDataPath1;
      //sL_DataPath2 := (G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sDataPath2;

      //sL_AdministratorMail1 := (G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sAdministratorMail1;
      //sL_AdministratorMail2 := (G_SGSParamIniStrList.Objects[ii] as TSGSParamIniData).sAdministratorMail2;
      sL_Version := (G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sVersion;
      sL_SrcCode := (G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sSrcCode;
      sL_SrcType := (G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sSrcType;

      sL_DstCode := (G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sDstCode;
      sL_DstType := (G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sDstType;
      sL_CasID := (G_SGSParamIniStrList.Objects[nL_Ndx] as TSGSParamIniData).sCasID;



      sL_DstMsgID := '';     //執行NightRun時為空值
      sL_InstQuery := '0';   //執行NightRun時固定為歷史查詢0:歷史查詢  1:即時查詢
      //sL_HandleEDateTime := sL_ActionDate + ' 23:59:59';
      sL_QueryDateTime := '';

      nL_TotalExecCounts := 100;//此為測試資料XML,故隨便設大於0即可
    end;
//*************************************

    sL_SQL := 'SELECT HandleEDateTime,Count(HandleEDateTime) HandleCounts from ' + sL_Table +
              ' Group by HandleEDateTime Having HandleEDateTime between ' +
              TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_StartDate)) + ' and ' +
              TUstr.getOracleSQLDateTimeStr(StrToDateTime(sI_StopDate)) ;
              
    L_AdoQry := dtmMain.runSQL(SELECT_MODE,sI_CompCode,sL_SQL);
    dtmMain.adoQryLoadExecCounts.Clone(L_AdoQry,ltUnspecified);
    nL_TotalXmlCounts := dtmMain.adoQryLoadExecCounts.RecordCount;

    if sI_ModeType = '1' then
    begin
      //取得前一天 ICCard 的總數量
      nL_TotalHandleCounts := dtmMain.getYesterdayTotalICCardNum(sL_Table,sI_CompCode,sI_StartDate);
    end;    

    nL_XmlCounts := 0;
    with dtmMain.adoQryLoadExecCounts do
    begin
      First;            
      while not Eof do
      begin
        sL_HandleEDateTime := TUstr.replaceStr(dtmMain.adoQryLoadExecCounts.FieldByName('HandleEDateTime').AsString,'-','/');
        nL_HandleCounts := dtmMain.adoQryLoadExecCounts.FieldByName('HandleCounts').AsInteger;
        nL_TotalHandleCounts := nL_TotalHandleCounts + nL_HandleCounts;

        //1:IC卡資訊2:產品定義資訊3:產品定購資訊4:授權資訊
        if sI_ModeType = '1' then
        begin
          //sL_TotalICCardNum := IntToStr(nL_TotalHandleCounts);
          //sL_TotalSubNum := IntToStr(nL_TotalHandleCounts);

          dtmMain.getICCardNumAndSubNum(true,sI_CompCode,sL_HandleEDateTime,nL_TotalICCardNum,nL_TotalSubNum);

          sL_TotalICCardNum := IntToStr(nL_TotalHandleCounts);  //以測評資料為資料
          sL_TotalSubNum := IntToStr(nL_TotalSubNum);   //以實際SO004當前整個SMS內用戶總數
          sL_TotalProductNum := '';
          sL_TotalChannelNum := '';

          sL_QueryType := IC_CARD_Query;
        end
        else if sI_ModeType = '2' then
        begin
          sL_TotalICCardNum := '';
          sL_TotalSubNum := '';
          //sL_TotalProductNum := IntToStr(nL_TotalHandleCounts);
          //sL_TotalChannelNum := IntToStr(nL_TotalHandleCounts);
          dtmMain.getProductNumAndChannelNum(sI_CompCode,sL_HandleEDateTime,nL_TotalProductNum,nL_TotalChannelNum);
          sL_TotalProductNum := IntToStr(nL_TotalProductNum);
          sL_TotalChannelNum := IntToStr(nL_TotalChannelNum);

          sL_QueryType := CA_PRODUCT_QUERY;
        end
        else if sI_ModeType = '3' then
        begin
          sL_TotalICCardNum := '';
          sL_TotalSubNum := '';
          sL_TotalProductNum := '';
          sL_TotalChannelNum := '';

          sL_QueryType := PROD_PURCHASE_QUERY;
        end
        else if sI_ModeType = '4' then
        begin
          sL_TotalICCardNum := '';
          sL_TotalSubNum := '';
          sL_TotalProductNum := '';
          sL_TotalChannelNum := '';

          sL_QueryType := ENTITLEMENT_QUERY; 
        end;


        frmMain.returnQueryXML(false, sL_QueryType,sL_Table,sL_CmdSeq,sI_CompCode,sL_Version,sL_QueryDateTime,
                               sL_SrcCode,sL_SrcType,sL_SrcMsgID,
                               sL_DstCode,sL_DstType,sL_DstMsgID,
                               sL_CasID,sL_InstQuery,sL_FirstDate,sL_HandleEDateTime,
                               sL_RetCode,sL_Msg,sL_TotalICCardNum,sL_TotalSubNum,
                               sL_TotalProductNum,sL_TotalChannelNum,nL_TotalExecCounts);


        nL_XmlCounts := nL_XmlCounts + 1;
        stbStatus.Panels[1].Text := IntToStr(nL_XmlCounts) + '/' + IntToStr(nL_TotalXmlCounts);
        stbStatus.Refresh;

        Next;
      end;
    end;
end;

procedure TfrmLoadData.btnCAProductPathClick(Sender: TObject);
var
    sL_Path : String;
begin
    if OpenDialog1.Execute then
    begin
      sL_Path := OpenDialog1.FileName;
      medCAProductPath.Text := sL_Path;
    end;
end;

procedure TfrmLoadData.btnProdPurchasePathClick(Sender: TObject);
var
    sL_Path : String;
begin
    if OpenDialog1.Execute then
    begin
      sL_Path := OpenDialog1.FileName;
      medProdPurchasePath.Text := sL_Path;
    end;
end;

procedure TfrmLoadData.btnEntitlementPathClick(Sender: TObject);
var
    sL_Path : String;
begin
    if OpenDialog1.Execute then
    begin
      sL_Path := OpenDialog1.FileName;
      medEntitlementPath.Text := sL_Path;
    end;
end;

procedure TfrmLoadData.showLogMsg(sI_LogMsg: String);
begin
    if UpperCase(sI_LogMsg) = 'BLANK' then
      memErrorLog.Lines.Add('')
    else
      memErrorLog.Lines.Add('[' + DateTimeToStr(now) + ']--' + sI_LogMsg);
end;

procedure TfrmLoadData.btnTransferClick(Sender: TObject);
var
    sL_CompCode,sL_QueryType,sL_ModeType,sL_ReadStr : String;
    L_AllDataTxt,L_Master : TStringList;
    L_ChannelDataTxt,L_ChannelMaster : TStringList;
    ii,nL_ParamCounts,nL_ChannelParamCounts,nL_AreaDataNo,nL_CurLineNo : Integer;
    bL_IsOkDataFile : Boolean;
    dL_CurrTime : TTime;
    sL_ICCardPath,sL_ChannelPath : String;
    sL_CAProductPath,sL_ProdPurchasePath,sL_EntitlementPath : String;
    F_SourcePath : TextFile;
begin
    if isDataOK then
    begin
      if MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_7_Content'), mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        stbStatus.Panels[0].Text := '';
        stbStatus.Panels[1].Text := '';
        stbStatus.Panels[2].Text := '';

        dL_CurrTime := now;
        bL_IsOkDataFile := true;
        sL_CompCode := getCobCompCode;


        sL_ICCardPath := Trim(medICCardPath.Text);
        sL_CAProductPath := Trim(medCAProductPath.Text);
        sL_ChannelPath := Trim(medChannelPath.Text);
        sL_ProdPurchasePath := Trim(medProdPurchasePath.Text);
        sL_EntitlementPath := Trim(medEntitlementPath.Text);
        nL_AreaDataNo := dtmMain.getAreaDataNo(sL_CompCode);

  //****************************************************
  //****************************************************
  //1:IC卡資訊
        if chkICCard.Checked then
        begin
          showLogMsg('BLANK');
          showLogMsg('[ICCardData] Open File...' + medICCardPath.Text);

          if FileExists(sL_ICCardPath) then
          begin
            sL_ModeType := '1';
            L_AllDataTxt := TStringList.Create;
            L_Master := TStringList.Create;
            //L_AllDataTxt.LoadFromFile(sL_ICCardPath);//將路徑下的文字檔轉到TStringList內

            //檢查該查詢種類的資料檔參數個數是否正確(只檢查第一筆資料)
            AssignFile(F_SourcePath,Trim(medICCardPath.Text));
            Reset(F_SourcePath);
            Readln(F_SourcePath,sL_ReadStr);
            CloseFile(F_SourcePath);
            L_Master := TUstr.ParseStrings(sL_ReadStr, TAB_STRING);
            nL_ParamCounts := L_Master.count;
            showLogMsg('[ICCardData] Check :  ' + medICCardPath.Text);
            bL_IsOkDataFile := checkIsOkDataFile(sL_ModeType,nL_ParamCounts,nL_ChannelParamCounts);

            if bL_IsOkDataFile then
            begin
              showLogMsg('[ICCardData] Transfer Txt...');
              handleICCardTransTxt(L_AllDataTxt,sL_CompCode,sL_ModeType,dL_CurrTime);
            end
            else
              showLogMsg('[ICCardData] Error File : ' + medICCardPath.Text);

            L_AllDataTxt.Free;
            L_Master.Free;
          end
          else
            showLogMsg('[ICCardData] File not Exist: ' + medICCardPath.Text);
        end;


  //****************************************************
  //****************************************************
        //2:產品定義資訊
        if chkCAProduct.Checked then
        begin
          showLogMsg('BLANK');
          showLogMsg('[CAProductData] Open File...' + medCAProductPath.Text);

          if FileExists(sL_CAProductPath) then
          begin
            if FileExists(sL_ChannelPath) then
            begin
              sL_ModeType := '2';
              L_AllDataTxt := TStringList.Create;
              L_Master := TStringList.Create;
              L_AllDataTxt.LoadFromFile(sL_CAProductPath);//將路徑下的文字檔轉到TStringList內


              //檢查該查詢種類的資料檔參數個數是否正確(只檢查第一筆資料)
              sL_ReadStr := L_AllDataTxt.Strings[0];
              L_Master := TUstr.ParseStrings(sL_ReadStr, TAB_STRING);
              nL_ParamCounts := L_Master.count;


              L_ChannelDataTxt := TStringList.Create;
              L_ChannelMaster := TStringList.Create;
              L_ChannelDataTxt.LoadFromFile(sL_ChannelPath);//將路徑下的文字檔轉到TStringList內
              L_ChannelMaster := TUstr.ParseStrings(L_ChannelDataTxt.Strings[0], TAB_STRING);
              nL_ChannelParamCounts := L_ChannelMaster.count;


              showLogMsg('[CAProductData] Check :  ' + medCAProductPath.Text);
              bL_IsOkDataFile := checkIsOkDataFile(sL_ModeType,nL_ParamCounts,nL_ChannelParamCounts);

              if bL_IsOkDataFile then
              begin
                showLogMsg('[CAProductData] Transfer Txt...');
                handleCAProductTransTxt(L_AllDataTxt,L_ChannelDataTxt,sL_CompCode,sL_ModeType,dL_CurrTime);
              end
              else
                showLogMsg('[CAProductData] Error File : ' + medCAProductPath.Text);

              L_AllDataTxt.Free;
              L_Master.Free;
              L_ChannelDataTxt.Free;
              L_ChannelMaster.Free;
            end
            else
              showLogMsg('[CAProductData] File not Exist: ' + medChannelPath.Text);
          end
          else
            showLogMsg('[CAProductData] File not Exist: ' + medCAProductPath.Text);
        end;


  //****************************************************
  //****************************************************
        //3:產品定購資訊
        if chkProdPurchase.Checked then
        begin
          showLogMsg('BLANK');
          showLogMsg('[ProdPurchaseData] Open File...' + medProdPurchasePath.Text);

          if FileExists(sL_ProdPurchasePath) then
          begin
            sL_ModeType := '3';
            L_AllDataTxt := TStringList.Create;
            L_Master := TStringList.Create;
            //L_AllDataTxt.LoadFromFile(sL_ProdPurchasePath);//將路徑下的文字檔轉到TStringList內

            //檢查該查詢種類的資料檔參數個數是否正確(只檢查第一筆資料)
            AssignFile(F_SourcePath,Trim(medProdPurchasePath.Text));
            Reset(F_SourcePath);
            Readln(F_SourcePath,sL_ReadStr);
            CloseFile(F_SourcePath);
            L_Master := TUstr.ParseStrings(sL_ReadStr, TAB_STRING);
            nL_ParamCounts := L_Master.count;
            showLogMsg('[ProdPurchaseData] Check :  ' + medProdPurchasePath.Text);
            bL_IsOkDataFile := checkIsOkDataFile(sL_ModeType,nL_ParamCounts,nL_ChannelParamCounts);

            if bL_IsOkDataFile then
            begin
              showLogMsg('[ProdPurchaseData] Transfer Txt...');
              handleProdPurchaseTransTxt(L_AllDataTxt,sL_CompCode,sL_ModeType,dL_CurrTime);
            end
            else
              showLogMsg('[ProdPurchaseData] Error File : ' + medProdPurchasePath.Text);

            L_AllDataTxt.Free;
            L_Master.Free;
          end
          else
            showLogMsg('[ProdPurchaseData] File not Exist: ' + medProdPurchasePath.Text);
        end;



  //****************************************************
  //****************************************************
        //4:授權資訊
        if chkEntitlement.Checked then
        begin
          showLogMsg('BLANK');
          showLogMsg('[EntitlementData] Open File...' + medEntitlementPath.Text);

          if FileExists(sL_EntitlementPath) then
          begin
            sL_ModeType := '4';
            L_AllDataTxt := TStringList.Create;
            L_Master := TStringList.Create;
            //L_AllDataTxt.LoadFromFile(sL_EntitlementPath);//將路徑下的文字檔轉到TStringList內


            //檢查該查詢種類的資料檔參數個數是否正確(只檢查第一筆資料)
            AssignFile(F_SourcePath,Trim(medEntitlementPath.Text));
            Reset(F_SourcePath);
            Readln(F_SourcePath,sL_ReadStr);
            CloseFile(F_SourcePath);
            L_Master := TUstr.ParseStrings(sL_ReadStr, TAB_STRING);
            nL_ParamCounts := L_Master.count;
            showLogMsg('[EntitlementData] Check :  ' + medEntitlementPath.Text);
            bL_IsOkDataFile := checkIsOkDataFile(sL_ModeType,nL_ParamCounts,nL_ChannelParamCounts);

            if bL_IsOkDataFile then
            begin
              showLogMsg('[EntitlementData] Transfer Txt...');
              handleEntitlementTransTxt(L_AllDataTxt,sL_CompCode,sL_ModeType,dL_CurrTime);
            end
            else
              showLogMsg('[EntitlementData] Error File : ' + medEntitlementPath.Text);

            L_AllDataTxt.Free;
            L_Master.Free;
          end
          else
            showLogMsg('[EntitlementData] File not Exist: ' + medEntitlementPath.Text);
        end;
      end;
    end;
end;

function TfrmLoadData.handleICCardTransTxt(I_AllDataTxt: TStringList;
  sI_CompCode,sI_ModeType: String; dI_CurrTime: TTime): String;
var
    nL_CurLineNo,ii,nL_FileNum : Integer;
    L_Master : TStringList;
    sL_HandleEDateTime,sL_RealDateTime,sL_ExecDateTime,sL_Status : String;
    sL_ICCardNo,sL_ICCardID,sL_CasID,sL_Feature,sL_Version,sL_ReadStr : String;
    dL_DateTime2,dL_DiffDateTime : TTime;
    sL_AllCounts,sL_SQL,sL_DiffDateTime,sL_RealTxt,sL_TransferPath  : String;
    F_FilePath,F_SourcePath : TextFile;
begin
//jackal
    nL_FileNum := 1;
    //讀進資料
    AssignFile(F_SourcePath,Trim(medICCardPath.Text));
    Reset(F_SourcePath);

    sL_TransferPath := getTransferPath(sI_ModeType,nL_FileNum);
    AssignFile(F_FilePath,sL_TransferPath);
    Append(F_FilePath);

    L_Master := TStringList.Create;
    //sL_AllCounts := IntTOsTR(I_AllDataTxt.Count);
    sL_ExecDateTime := TUdateTime.GetPureDateTimeStr(NOW,'','');

    //for nL_CurLineNo:=0 to I_AllDataTxt.Count-1 do
    //begin
    nL_CurLineNo := 0;
    while not Eof(F_SourcePath) do
    begin
      nL_CurLineNo := nL_CurLineNo + 1;
      //sL_ReadStr := I_AllDataTxt.Strings[nL_CurLineNo];
      Readln(F_SourcePath,sL_ReadStr);
      L_Master := TUstr.ParseStrings(sL_ReadStr, TAB_STRING);

      //文字檔參數順序(1)對IC卡進行操作的時間(2)New為創建IC卡,Delete為刪除(3)IC卡號
      //              (4)IC卡內部卡號(5)CASID(6)IC卡特性(7)版本
      sL_RealDateTime := Trim(TUstr.replaceStr(TUstr.replaceStr(L_Master.Strings[0],'-',''),':',''));

      sL_Status := Trim(L_Master.Strings[1]);
      if Uppercase(sL_Status) = 'NEW' then
         sL_Status := '0'
      else if Uppercase(sL_Status) = 'DELETE' then
         sL_Status := '1'
      else
         sL_Status := '';

      sL_ICCardNo := Trim(L_Master.Strings[2]);
      sL_ICCardID := Trim(L_Master.Strings[3]);
      sL_CasID := Trim(L_Master.Strings[4]);
      sL_Feature := Trim(L_Master.Strings[5]);
      sL_Version := Trim(L_Master.Strings[6]);

      sL_HandleEDateTime := Copy(sL_RealDateTime,0,8) + ' 235959';

      sL_RealTxt := sI_CompCode + TAB_STRING + sL_HandleEDateTime + TAB_STRING +
                    sL_RealDateTime + TAB_STRING + sL_ExecDateTime + TAB_STRING +
                    sL_Status + TAB_STRING + sL_ICCardNo + TAB_STRING +
                    sL_ICCardID + TAB_STRING + sL_CasID + TAB_STRING +
                    sL_Feature + TAB_STRING + sL_Version;
      Writeln(F_FilePath,sL_RealTxt);


      if (nL_CurLineNo Mod SHOW_UI_LINE_COUNTS)=0 then
      begin
        Flush(F_FilePath);
        stbStatus.Panels[1].Text := IntToStr((nL_CurLineNo));
        stbStatus.Refresh;
      end;
    end;

    //顯示剩餘筆數
    stbStatus.Panels[1].Text := IntToStr((nL_CurLineNo));
    stbStatus.Refresh;

    CloseFile(F_FilePath);
    CloseFile(F_SourcePath);
    L_Master.Free;

    dL_DateTime2 := now;
    dL_DiffDateTime := dL_DateTime2 - dI_CurrTime;
    sL_DiffDateTime := TimeToStr(dL_DiffDateTime);

    showLogMsg('[ICCardData] Txt Transfer OK-[Total Time:' + sL_DiffDateTime + ']');
end;

procedure TfrmLoadData.btnTransferPathClick(Sender: TObject);
var
    sL_Path : String;
begin
    if OpenDialog1.Execute then
    begin
      sL_Path := ExtractFileDir(OpenDialog1.FileName);
      medTransferPath.Text := sL_Path;
    end;
end;

function TfrmLoadData.handleCAProductTransTxt(I_AllDataTxt,L_ChannelDataTxt : TStringList;sI_CompCode,sI_ModeType : String;dI_CurrTime : TTime) : String;
var
    nL_CurLineNo,ii,nL_FileNum : Integer;
    L_Master,L_ChannelMaster : TStringList;
    sL_HandleEDateTime,sL_RealDateTime,sL_ExecDateTime : String;
    dL_DateTime2,dL_DiffDateTime : TTime;
    sL_CasID,sL_ReadStr,sL_AllCounts,sL_SQL,sL_DiffDateTime,sL_TransferPath  : String;
    sL_ProductCode,sL_ProductName,sL_Provider,sL_ChannelName,sL_RealTxt : String;
    F_FilePath : TextFile;
begin
    nL_FileNum := 1;
    sL_TransferPath := getTransferPath(sI_ModeType,nL_FileNum);
    AssignFile(F_FilePath,sL_TransferPath);
    Append(F_FilePath);

    sL_ExecDateTime := TUdateTime.GetPureDateTimeStr(NOW,'','');

    //先將產品頻道對應存至CDS
    for ii:=0 to L_ChannelDataTxt.Count-1 do
    begin
      //文字檔參數順序(1)產品代碼(2)節目平臺名稱(3)原始頻道名稱
      L_ChannelMaster := TStringList.Create;
      L_ChannelMaster := TUstr.ParseStrings(L_ChannelDataTxt.Strings[ii], TAB_STRING);
      sL_ProductCode := L_ChannelMaster.Strings[0];
      sL_Provider := L_ChannelMaster.Strings[1];
      sL_ChannelName := L_ChannelMaster.Strings[2];

      dtmMain.cdsChannelID.Append;
      dtmMain.cdsChannelID.FieldByName('ProductCode').AsString := sL_ProductCode;
      dtmMain.cdsChannelID.FieldByName('Provider').AsString := sL_Provider;
      dtmMain.cdsChannelID.FieldByName('ChannelName').AsString := sL_ChannelName;
      dtmMain.cdsChannelID.Post;
      L_ChannelMaster.Free;
    end;

    L_Master := TStringList.Create;
    //sL_AllCounts := IntTOsTR(I_AllDataTxt.Count);

    for nL_CurLineNo:=0 to I_AllDataTxt.Count-1 do
    begin

      sL_ReadStr := I_AllDataTxt.Strings[nL_CurLineNo];
      L_Master := TUstr.ParseStrings(sL_ReadStr, TAB_STRING);

      //文字檔參數順序(1)增加產品操作時間(2)產品代碼(3)產品名稱(4)CASID
      sL_RealDateTime := Trim(TUstr.replaceStr(TUstr.replaceStr(L_Master.Strings[0],'-',''),':',''));

      sL_ProductCode := Trim(L_Master.Strings[1]);
      sL_ProductName := Trim(L_Master.Strings[2]);
      sL_CasID := Trim(L_Master.Strings[3]);

      sL_Provider := '';
      sL_ChannelName := '';
      if dtmMain.cdsChannelID.Locate('ProductCode',VarArrayOf([sL_ProductCode]) , []) then
      begin
        sL_Provider := dtmMain.cdsChannelID.fieldbyname('Provider').asstring;
        sL_ChannelName := dtmMain.cdsChannelID.fieldbyname('ChannelName').asstring;
      end;

      sL_HandleEDateTime := Copy(sL_RealDateTime,0,8) + ' 235959';


      sL_RealTxt := sI_CompCode + TAB_STRING + sL_HandleEDateTime + TAB_STRING +
                    sL_RealDateTime + TAB_STRING + sL_ExecDateTime + TAB_STRING +
                    sL_ProductCode + TAB_STRING + sL_ProductName + TAB_STRING +
                    sL_CasID + TAB_STRING + sL_Provider + TAB_STRING +
                    sL_ChannelName;
      Writeln(F_FilePath,sL_RealTxt);
      
      if (nL_CurLineNo Mod SHOW_UI_LINE_COUNTS)=0 then
      begin
        stbStatus.Panels[1].Text := IntToStr((nL_CurLineNo));
        stbStatus.Refresh;
      end;


      if (nL_CurLineNo>0) and ((nL_CurLineNo Mod MAX_STRING_LIST_COUNTS)=0) then
        Flush(F_FilePath);
    end;

    //顯示剩餘筆數
    stbStatus.Panels[1].Text := IntToStr((nL_CurLineNo));
    stbStatus.Refresh;

    CloseFile(F_FilePath);
    L_Master.Free;

    dL_DateTime2 := now;
    dL_DiffDateTime := dL_DateTime2 - dI_CurrTime;
    sL_DiffDateTime := TimeToStr(dL_DiffDateTime);


    showLogMsg('[CAProductData] Txt Transfer OK-[Total Time:' + sL_DiffDateTime + ']');
end;

function TfrmLoadData.handleProdPurchaseTransTxt(I_AllDataTxt: TStringList;
  sI_CompCode,sI_ModeType: String; dI_CurrTime: TTime): String;
var
    nL_CurLineNo,ii,nL_HandleCounts,nL_FileNum : Integer;
    L_Master,L_ProductList,L_YYYYMMDD,L_YYYYMMDDHHMMSS : TStringList;
    sL_HandleEDateTime,sL_RealDateTime,sL_ExecDateTime,sL_Status,sL_RealTxt : String;
    sL_ICCardNo,sL_ICCardID,sL_CasID,sL_ProductCode,sL_ProductCodeList,sL_ReadStr : String;
    dL_DateTime2,dL_DiffDateTime : TTime;
    sL_AllCounts,sL_SQL,sL_DiffDateTime,sL_PurchaseID,sL_StartTime,sL_EndTime,sL_TransferPath  : String;
    F_FilePath,F_SourcePath : TextFile;
    sL_YYYY,sL_MM,sL_DD : String;
begin
    nL_FileNum := 1;
    //讀進資料
    AssignFile(F_SourcePath,Trim(medProdPurchasePath.Text));
    Reset(F_SourcePath);

    sL_TransferPath := getTransferPath(sI_ModeType,nL_FileNum);
    AssignFile(F_FilePath,sL_TransferPath);
    Append(F_FilePath);

    nL_HandleCounts := 0;
    L_Master := TStringList.Create;
    L_ProductList := TStringList.Create;
    //sL_AllCounts := IntTOsTR(I_AllDataTxt.Count);
    sL_ExecDateTime := TUdateTime.GetPureDateTimeStr(NOW,'','');

    //for nL_CurLineNo:=0 to I_AllDataTxt.Count-1 do
    //begin
    nL_CurLineNo := 0;
    while not Eof(F_SourcePath) do
    begin
      nL_CurLineNo := nL_CurLineNo + 1;
      //sL_ReadStr := I_AllDataTxt.Strings[nL_CurLineNo];
      Readln(F_SourcePath,sL_ReadStr);
      L_Master := TUstr.ParseStrings(sL_ReadStr, TAB_STRING);

      //文字檔參數順序(1)產品購買操作時間(2)MODIFY新加/修改定購,Delete刪除定購(3)IC卡號
      //              (4)IC卡內部卡號(5)CASID(6)購買產品代碼組(7)定購ID
      //              (8)生效時間(9)失效時間
      sL_RealDateTime := Trim(TUstr.replaceStr(TUstr.replaceStr(L_Master.Strings[0],'-',''),':',''));

      sL_Status := Trim(L_Master.Strings[1]);
      if (Uppercase(sL_Status) = 'MODIFY') OR (Uppercase(sL_Status) = 'CREATE') then
         sL_Status := '0'
      else if Uppercase(sL_Status) = 'DELETE' then
         sL_Status := '1'
      else
         sL_Status := '';

      sL_ICCardNo := Trim(L_Master.Strings[2]);
      sL_ICCardID := Trim(L_Master.Strings[3]);
      sL_CasID := Trim(L_Master.Strings[4]);
      sL_ProductCodeList := Trim(L_Master.Strings[5]);
      sL_PurchaseID := Trim(L_Master.Strings[6]);
      sL_StartTime := Trim(TUstr.replaceStr(TUstr.replaceStr(L_Master.Strings[7],'-',''),':',''));

      //**********************
      L_YYYYMMDD := TStringList.Create;
      L_YYYYMMDDHHMMSS := TStringList.Create;
      L_YYYYMMDDHHMMSS := TUstr.ParseStrings(L_Master.Strings[8],' ');
      L_YYYYMMDD := TUstr.ParseStrings(L_YYYYMMDDHHMMSS.Strings[0],'-');

      sL_YYYY := L_YYYYMMDD.Strings[0];
      sL_MM := Format('%.2d',[StrToInt(L_YYYYMMDD.Strings[1])]);
      sL_DD := Format('%.2d',[StrToInt(L_YYYYMMDD.Strings[2])]);
      sL_EndTime := sL_YYYY + sL_MM + sL_DD + ' ' + TUstr.replaceStr(L_YYYYMMDDHHMMSS.Strings[1],':','');

      L_YYYYMMDD.Free;
      L_YYYYMMDDHHMMSS.Free;
      //**********************


      //sL_EndTime := Trim(TUstr.replaceStr(TUstr.replaceStr(L_Master.Strings[8],'-',''),':',''));

      sL_HandleEDateTime := Copy(sL_RealDateTime,0,8) + ' 235959';


      L_ProductList := TUstr.ParseStrings(sL_ProductCodeList,',');
      for ii:=0 to L_ProductList.Count-1 do
      begin
        nL_HandleCounts := nL_HandleCounts + 1;
        sL_ProductCode := L_ProductList.Strings[ii];

        sL_RealTxt := sI_CompCode + TAB_STRING + sL_HandleEDateTime + TAB_STRING +
                      sL_RealDateTime + TAB_STRING + sL_ExecDateTime + TAB_STRING +
                      sL_Status + TAB_STRING + sL_PurchaseID + TAB_STRING +
                      sL_ICCardNo + TAB_STRING + sL_ICCardID + TAB_STRING +
                      sL_CasID + TAB_STRING + sL_ProductCode + TAB_STRING +
                      sL_StartTime + TAB_STRING + sL_EndTime;
        Writeln(F_FilePath,sL_RealTxt);

        if (nL_HandleCounts Mod SHOW_UI_COUNTS)=0 then
        begin
          Flush(F_FilePath);
          stbStatus.Panels[0].Text := IntToStr((nL_HandleCounts));
          stbStatus.Refresh;
        end;

        {
        if (nL_HandleCounts>0) and ((nL_HandleCounts Mod MAX_STRING_LIST_COUNTS)=0) then
        begin
          Flush(F_FilePath);
          CloseFile(F_FilePath);
          nL_FileNum := nL_FileNum + 1;
          sL_TransferPath := getTransferPath(sI_ModeType,nL_FileNum);
          AssignFile(F_FilePath,sL_TransferPath);
          Append(F_FilePath);
        end;
        }
      end;
      L_ProductList.Free;
      L_ProductList := TStringList.Create;

      if (nL_CurLineNo Mod SHOW_UI_LINE_COUNTS)=0 then
      begin
        stbStatus.Panels[1].Text := IntToStr((nL_CurLineNo));
        stbStatus.Refresh;
      end;
    end;

    //顯示剩餘筆數
    stbStatus.Panels[0].Text := IntToStr((nL_HandleCounts));
    stbStatus.Refresh;
    stbStatus.Panels[1].Text := IntToStr((nL_CurLineNo));
    stbStatus.Refresh;

    CloseFile(F_FilePath);
    CloseFile(F_SourcePath);
    L_Master.Free;
    L_ProductList.free;

    dL_DateTime2 := now;
    dL_DiffDateTime := dL_DateTime2 - dI_CurrTime;
    sL_DiffDateTime := TimeToStr(dL_DiffDateTime);

    showLogMsg('[ProdPurchaseData] Txt Transfer OK-[Total Time:' + sL_DiffDateTime + ']');
end;

function TfrmLoadData.handleEntitlementTransTxt(I_AllDataTxt: TStringList;
  sI_CompCode,sI_ModeType: String; dI_CurrTime: TTime): String;
var
    nL_CurLineNo,ii,nL_HandleCounts,nL_FileNum : Integer;
    L_Master,L_ProductList : TStringList;
    sL_HandleEDateTime,sL_RealDateTime,sL_ExecDateTime,sL_Status : String;
    sL_ICCardNo,sL_ICCardID,sL_CasID,sL_ProductCode,sL_ProductCodeList,sL_ReadStr : String;
    dL_DateTime2,dL_DiffDateTime : TTime;
    sL_AllCounts,sL_SQL,sL_DiffDateTime,sL_RealTxt,sL_TransferPath  : String;
    F_FilePath,F_SourcePath : TextFile;
begin
    nL_FileNum := 1;
    //讀進資料
    AssignFile(F_SourcePath,Trim(medEntitlementPath.Text));
    Reset(F_SourcePath);

    sL_TransferPath := getTransferPath(sI_ModeType,nL_FileNum);
    AssignFile(F_FilePath,sL_TransferPath);
    Append(F_FilePath);

    nL_HandleCounts := 0;
    L_Master := TStringList.Create;
    L_ProductList := TStringList.Create;
    //sL_AllCounts := IntTOsTR(I_AllDataTxt.Count);
    sL_ExecDateTime := TUdateTime.GetPureDateTimeStr(NOW,'','');


    //for nL_CurLineNo:=0 to I_AllDataTxt.Count-1 do
    //begin
    nL_CurLineNo := 0;
    while not Eof(F_SourcePath) do
    begin
      nL_CurLineNo := nL_CurLineNo + 1;
      //sL_ReadStr := I_AllDataTxt.Strings[nL_CurLineNo];
      Readln(F_SourcePath,sL_ReadStr);
      L_Master := TUstr.ParseStrings(sL_ReadStr, TAB_STRING);

      //文字檔參數順序(1)授權操作的時間(2)Create為授權,Delete為反授權(3)IC卡號
      //              (4)IC卡內部卡號(5)CASID(6)授權產品代碼組
      sL_RealDateTime := Trim(TUstr.replaceStr(TUstr.replaceStr(L_Master.Strings[0],'-',''),':',''));

      sL_Status := Trim(L_Master.Strings[1]);
      if Uppercase(sL_Status) = 'CREATE' then
         sL_Status := '0'
      else if Uppercase(sL_Status) = 'DELETE' then
         sL_Status := '1'
      else
         sL_Status := '';

      sL_ICCardNo := Trim(L_Master.Strings[2]);
      sL_ICCardID := Trim(L_Master.Strings[3]);
      sL_CasID := Trim(L_Master.Strings[4]);
      sL_ProductCodeList := Trim(L_Master.Strings[5]);

      sL_HandleEDateTime := Copy(sL_RealDateTime,0,8) + ' 235959';


      L_ProductList := TUstr.ParseStrings(sL_ProductCodeList,',');
      for ii:=0 to L_ProductList.Count-1 do
      begin
        nL_HandleCounts := nL_HandleCounts + 1;
        sL_ProductCode := L_ProductList.Strings[ii];


        sL_RealTxt := sI_CompCode + TAB_STRING + sL_HandleEDateTime + TAB_STRING +
                      sL_RealDateTime + TAB_STRING + sL_ExecDateTime + TAB_STRING +
                      sL_Status + TAB_STRING + sL_ICCardNo + TAB_STRING +
                      sL_ICCardID + TAB_STRING + sL_CasID + TAB_STRING +
                      sL_ProductCode;
        Writeln(F_FilePath,sL_RealTxt);



        if (nL_HandleCounts Mod SHOW_UI_COUNTS)=0 then
        begin
          Flush(F_FilePath);
          stbStatus.Panels[0].Text := IntToStr((nL_HandleCounts));
          stbStatus.Refresh;
        end;

        {
        if (nL_HandleCounts>0) and ((nL_HandleCounts Mod MAX_STRING_LIST_COUNTS)=0) then
        begin
          CloseFile(F_FilePath);
          nL_FileNum := nL_FileNum + 1;
          sL_TransferPath := getTransferPath(sI_ModeType,nL_FileNum);
          AssignFile(F_FilePath,sL_TransferPath);
          Append(F_FilePath);
        end;
        }
      end;
      L_ProductList.Free;
      L_ProductList := TStringList.Create;


      if (nL_CurLineNo Mod SHOW_UI_LINE_COUNTS)=0 then
      begin
        //stbStatus.Panels[1].Text := IntToStr((nL_CurLineNo)) + '/' + sL_AllCounts;
        stbStatus.Panels[1].Text := IntToStr((nL_CurLineNo));
        stbStatus.Refresh;
      end;
    end;


    //顯示剩餘筆數
    stbStatus.Panels[0].Text := IntToStr((nL_HandleCounts));
    stbStatus.Refresh;
    stbStatus.Panels[1].Text := IntToStr((nL_CurLineNo));
    stbStatus.Refresh;


    CloseFile(F_FilePath);
    CloseFile(F_SourcePath);
    L_Master.Free;
    L_ProductList.free;

    dL_DateTime2 := now;
    dL_DiffDateTime := dL_DateTime2 - dI_CurrTime;
    sL_DiffDateTime := TimeToStr(dL_DiffDateTime);


    showLogMsg('[EntitlementData] Txt Transfer OK-[Total Time:' + sL_DiffDateTime + ']');
end;

function TfrmLoadData.getTransferPath(sI_ModeType: String;nI_FileNum : Integer): String;
var
    sL_TransferPath,sL_FileNum : String;
    nL_FileHandel : Integer;
begin
  sL_FileNum := IntToStr(nI_FileNum);

  if sI_ModeType = '1' then
    sL_TransferPath := medTransferPath.Text + '\Transfer_SMS_ICCard.txt'
  else if sI_ModeType = '2' then
    sL_TransferPath := medTransferPath.Text + '\Transfer_SMS_CAProduct.txt'
  else if sI_ModeType = '3' then
    sL_TransferPath := medTransferPath.Text + '\Transfer_SMS_ProdPurchase.txt'
    //sL_TransferPath := medTransferPath.Text + '\Transfer_SMS_ProdPurchase_' + sL_FileNum + '.txt'
  else if sI_ModeType = '4' then
    sL_TransferPath := medTransferPath.Text + '\Transfer_SMS_Entitlement.txt'
  else if sI_ModeType = 'SO004' then
    sL_TransferPath := medTransferPath.Text + '\Transfer_SMS_SO004.txt'
  else if sI_ModeType = 'SOAC0201A' then
    sL_TransferPath := medTransferPath.Text + '\Transfer_SMS_SOAC0201A.txt'
  else if sI_ModeType = 'SOAC0201B' then
    sL_TransferPath := medTransferPath.Text + '\Transfer_SMS_SOAC0201B.txt'
  else if sI_ModeType = 'SO005' then
    sL_TransferPath := medTransferPath.Text + '\Transfer_SMS_SO005.txt'
  else if sI_ModeType = 'SOAC0202' then
    sL_TransferPath := medTransferPath.Text + '\Transfer_SMS_SOAC0202.txt';


  if not FileExists(sL_TransferPath) then
    nL_FileHandel := FileCreate(sL_TransferPath)
  else
  begin
    DeleteFile(sL_TransferPath);
    nL_FileHandel := FileCreate(sL_TransferPath);
  end;

  if (nL_FileHandel>=0) then
    FileClose(nL_FileHandel);
  Result := sL_TransferPath;
end;

function TfrmLoadData.isDataOK3: Boolean;
var
    sL_CompCode,sL_ICCardPath,sL_TransferPath : String;
begin
    IsDataOk3 := true;
    sL_CompCode := cobCompCode.Text;
    if sL_CompCode = '' then
    begin
        MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_1_Content'),mtError, [mbOK],0);
        cobCompCode.SetFocus;
        IsDataOk3 := false;
        exit;
    end;




    sL_ICCardPath := Trim(medICCardPath.Text);
    if sL_ICCardPath = '' then
    begin
        MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_3_Content'),mtError, [mbOK],0);
        btnICCardPath.SetFocus;
        IsDataOk3 := false;
        exit;
    end;




    sL_TransferPath := Trim(medTransferPath.Text);
    if sL_TransferPath = '' then
    begin
        MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_3_Content'),mtError, [mbOK],0);
        btnTransferPath.SetFocus;
        IsDataOk3 := false;
        exit;
    end;


end;

procedure TfrmLoadData.btnTransferSMSClick(Sender: TObject);
var
    dL_CurrTime : TTime;
    bL_IsOkDataFile : Boolean;
    sL_CompCode,sL_ICCardPath,sL_ModeType,sL_ReadStr : String;
    L_AllDataTxt,L_Master : TStringList;
    F_SourcePath : TextFile;
    nL_ParamCounts,nL_ChannelParamCounts : Integer;
begin
    if isDataOK3 then
    begin
      if MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_10_Content'), mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        stbStatus.Panels[0].Text := '';
        stbStatus.Panels[1].Text := '';
        stbStatus.Panels[2].Text := '';

        dL_CurrTime := now;
        bL_IsOkDataFile := true;
        sL_CompCode := getCobCompCode;
        sL_ICCardPath := Trim(medICCardPath.Text);

        showLogMsg('BLANK');
        showLogMsg('[SMS_SO004Data] Open File...' + medICCardPath.Text);

        if FileExists(sL_ICCardPath) then
        begin
          sL_ModeType := '1';
          L_AllDataTxt := TStringList.Create;
          L_Master := TStringList.Create;

          //檢查該查詢種類的資料檔參數個數是否正確(只檢查第一筆資料)
          AssignFile(F_SourcePath,Trim(medICCardPath.Text));
          Reset(F_SourcePath);
          Readln(F_SourcePath,sL_ReadStr);
          CloseFile(F_SourcePath);
          L_Master := TUstr.ParseStrings(sL_ReadStr, TAB_STRING);
          nL_ParamCounts := L_Master.count;
          showLogMsg('[SMS_SO004Data] Check :  ' + medICCardPath.Text);
          bL_IsOkDataFile := checkIsOkDataFile(sL_ModeType,nL_ParamCounts,nL_ChannelParamCounts);

          if bL_IsOkDataFile then
          begin
            showLogMsg('[SMS_SO004Data] Transfer Txt...');
            handleSMSSO004TransTxt(sL_CompCode,sL_ModeType,dL_CurrTime);
          end
          else
            showLogMsg('[SMS_SO004Data] Error File : ' + medICCardPath.Text);

          L_AllDataTxt.Free;
          L_Master.Free;
        end
        else
          showLogMsg('[SMS_SO004Data] File not Exist: ' + medICCardPath.Text);
      end;
    end;
end;

function TfrmLoadData.handleSMSSO004TransTxt(sI_CompCode,sI_ModeType : String;dI_CurrTime : TTime) : String;
var
    L_Master,L_CustStrLis : TStringList;
    nL_FileNum,nL_CurLineNo,nL_SeqNo,nL_CustNo : Integer;
    F_SourcePath,F_SO004File,F_SOAC0201AFile,F_SOAC0201BFile : TextFile;
    sL_TransferPath,sL_ExecDateTime,sL_ReadStr,sL_InstDate,sL_ICCardNo : String;
    sL_InitPlaceNo,sL_ServiceType,sL_UpdEn,sL_BatchNo1,sL_BatchNo2,sL_FaciStatusCode : String;
    sL_MaterialNo1,sL_MaterialNo2,sL_SeqNo,sL_ErrMsg,sL_CustID,sL_STBSeqNo : String;
    sL_STBFaciCode,sL_STBFaciName,sL_ICCFaciCode,sL_ICCFaciName,sL_BuyCode,sL_BuyName : String;
    L_SO002AdoQry : TADOQuery;
    dL_DateTime2,dL_DiffDateTime : TTime;

    sL_STBSO004Str,sL_ICCSO004Str,sL_SOAC0201AStr,sL_SOAC0201BStr,sL_DiffDateTime : String;
begin
//讀進ICCard測評測試資料
    AssignFile(F_SourcePath,Trim(medICCardPath.Text));
    Reset(F_SourcePath);

//開啟要寫入的三個文字檔
    sL_TransferPath := getTransferPath('SO004',nL_FileNum);
    AssignFile(F_SO004File,sL_TransferPath);
    Append(F_SO004File);

    sL_TransferPath := getTransferPath('SOAC0201A',nL_FileNum);
    AssignFile(F_SOAC0201AFile,sL_TransferPath);
    Append(F_SOAC0201AFile);

    sL_TransferPath := getTransferPath('SOAC0201B',nL_FileNum);
    AssignFile(F_SOAC0201BFile,sL_TransferPath);
    Append(F_SOAC0201BFile);

//整理共用參數資料
    L_Master := TStringList.Create;
    L_CustStrLis := TStringList.Create;
    sL_ExecDateTime := TUdateTime.GetPureDateTimeStr(dI_CurrTime,'','');
    sL_InitPlaceNo := '1';		    // location code for the STB/ICC
    sL_ServiceType := 'C';		// Service type code
    sL_UpdEn := 'PIERREJACKAL';// Update EN name
    sL_BatchNo1 := '1';		    //Batch number for ICC records
    sL_BatchNo2 := '2';		    //Batch number for STB records
    sL_MaterialNo1 := 'A123';    //Material number for ICC records
    sL_MaterialNo2 := 'A456';    //Material number for STB records
    sL_FaciStatusCode := '1';
    nL_SeqNo := 1000000;

    //取得CC&B基本資料
    sL_ErrMsg := dtmMain.getSMSBasicData(sI_CompCode,sL_STBFaciCode,sL_STBFaciName,sL_ICCFaciCode,sL_ICCFaciName,sL_BuyCode,sL_BuyName,L_CustStrLis);


    if sL_ErrMsg = '' then
    begin
      nL_CurLineNo := 0;
      nL_CustNo := 0;

      while not Eof(F_SourcePath) do
      begin
        //取得 SO002 有效戶
        //每個 SO002 中服務類別為'C'的有效戶各新增50筆ICC與50筆STB資料
        if (nL_CurLineNo Mod 50)=0 then
        begin
          if nL_CurLineNo <> 0 then
            nL_CustNo := nL_CustNo + 1;
          sL_CustID := L_CustStrLis.Strings[nL_CustNo];
        end;

        nL_CurLineNo := nL_CurLineNo + 1; //資料筆數控制,50筆資料換一個CustID


        Readln(F_SourcePath,sL_ReadStr);
        L_Master := TUstr.ParseStrings(sL_ReadStr, TAB_STRING);

        //文字檔參數順序(1)對IC卡進行操作的時間(2)New為創建IC卡,Delete為刪除(3)IC卡號
        //              (4)IC卡內部卡號(5)CASID(6)IC卡特性(7)版本
        sL_InstDate := Trim(Copy(TUstr.replaceStr(L_Master.Strings[0],'-',''),1,8));


        sL_ICCardNo := Trim(L_Master.Strings[2]);
        sL_STBSeqNo := Copy(sL_ICCardNo,5,12);

        //寫入SO004 STB
        nL_SeqNo := nL_SeqNo + 1;         //STB設備流水號
        sL_SeqNo := sL_InstDate + Copy(IntToStr(nL_SeqNo),1,7);
        sL_STBSO004Str := sL_CustID + TAB_STRING + sL_STBFaciCode + TAB_STRING +
                          sL_STBFaciName + TAB_STRING + sL_STBSeqNo + TAB_STRING +
                          sL_BuyCode + TAB_STRING + sL_BuyName + TAB_STRING +
                          sL_InstDate + TAB_STRING + sL_ICCardNo + TAB_STRING +
                          sL_ServiceType + TAB_STRING + sI_CompCode + TAB_STRING +
                          sL_InitPlaceNo + TAB_STRING + sL_SeqNo + TAB_STRING +
                          sL_UpdEn + TAB_STRING + sL_FaciStatusCode;
         Writeln(F_SO004File,sL_STBSO004Str);


        //寫入SO004 ICC
        nL_SeqNo := nL_SeqNo + 1;         //ICC設備流水號
        sL_SeqNo := sL_InstDate + Copy(IntToStr(nL_SeqNo),1,7);
        sL_ICCSO004Str := sL_CustID + TAB_STRING + sL_ICCFaciCode + TAB_STRING +
                          sL_ICCFaciName + TAB_STRING + sL_ICCardNo + TAB_STRING +
                          sL_BuyCode + TAB_STRING + sL_BuyName + TAB_STRING +
                          sL_InstDate + TAB_STRING + '' + TAB_STRING +
                          sL_ServiceType + TAB_STRING + sI_CompCode + TAB_STRING +
                          sL_InitPlaceNo + TAB_STRING + sL_SeqNo + TAB_STRING +
                          sL_UpdEn + TAB_STRING + sL_FaciStatusCode;
         Writeln(F_SO004File,sL_ICCSO004Str);


        //寫入SOAC0201A
        sL_SOAC0201AStr := sI_CompCode + TAB_STRING + sL_BatchNo2 + TAB_STRING +
                           sL_ExecDateTime + TAB_STRING + sL_MaterialNo2 + TAB_STRING +
                           sL_STBSeqNo + TAB_STRING + 'STB' + TAB_STRING +
                           '0';
        Writeln(F_SOAC0201AFile,sL_SOAC0201AStr);


        //寫入SOAC0201B
        sL_SOAC0201BStr := sI_CompCode + TAB_STRING + sL_BatchNo1 + TAB_STRING +
                           sL_ExecDateTime + TAB_STRING + sL_MaterialNo1 + TAB_STRING +
                           sL_ICCardNo + TAB_STRING + 'ICC' + TAB_STRING +
                           '0';
        Writeln(F_SOAC0201BFile,sL_SOAC0201BStr);



        if (nL_CurLineNo Mod SHOW_UI_LINE_COUNTS)=0 then
        begin
          Flush(F_SO004File);
          Flush(F_SOAC0201AFile);
          Flush(F_SOAC0201BFile);
          stbStatus.Panels[1].Text := IntToStr((nL_CurLineNo));
          stbStatus.Refresh;
        end;
      end;
    end
    else
      showLogMsg('[SMS_SO004Data] Transfer Error :' + sL_ErrMsg);


//關閉結束
    CloseFile(F_SourcePath);
    CloseFile(F_SO004File);
    CloseFile(F_SOAC0201AFile);
    CloseFile(F_SOAC0201BFile);
    L_Master.Free;
    L_CustStrLis.Free;

    dL_DateTime2 := now;
    dL_DiffDateTime := dL_DateTime2 - dI_CurrTime;
    sL_DiffDateTime := TimeToStr(dL_DiffDateTime);

    showLogMsg('[SMS_SO004Data] Txt Transfer OK-[Total Time:' + sL_DiffDateTime + ']');
end;

procedure TfrmLoadData.btnTransferSMSSO005Click(Sender: TObject);
var
    dL_CurrTime : TTime;
    bL_IsOkDataFile : Boolean;
    sL_CompCode,sL_EntitlementPath,sL_ModeType,sL_ReadStr : String;
    L_AllDataTxt,L_Master : TStringList;
    F_SourcePath : TextFile;
    nL_ParamCounts,nL_ChannelParamCounts : Integer;
begin
    if isDataOK4 then
    begin
      if MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_11_Content'), mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
        stbStatus.Panels[0].Text := '';
        stbStatus.Panels[1].Text := '';
        stbStatus.Panels[2].Text := '';

        dL_CurrTime := now;
        bL_IsOkDataFile := true;
        sL_CompCode := getCobCompCode;
        sL_EntitlementPath := Trim(medEntitlementPath.Text);

        showLogMsg('BLANK');
        showLogMsg('[SMS_SO005Data] Open File...' + medEntitlementPath.Text);

        if FileExists(sL_EntitlementPath) then
        begin
          sL_ModeType := '4';
          L_AllDataTxt := TStringList.Create;
          L_Master := TStringList.Create;

          //檢查該查詢種類的資料檔參數個數是否正確(只檢查第一筆資料)
          AssignFile(F_SourcePath,sL_EntitlementPath);
          Reset(F_SourcePath);
          Readln(F_SourcePath,sL_ReadStr);
          CloseFile(F_SourcePath);
          L_Master := TUstr.ParseStrings(sL_ReadStr, TAB_STRING);
          nL_ParamCounts := L_Master.count;
          showLogMsg('[SMS_SO005Data] Check :  ' + sL_EntitlementPath);
          bL_IsOkDataFile := checkIsOkDataFile(sL_ModeType,nL_ParamCounts,nL_ChannelParamCounts);

          if bL_IsOkDataFile then
          begin
            showLogMsg('[SMS_SO005Data] Transfer Txt...');
            handleSMSSO005TransTxt(sL_CompCode,sL_ModeType,dL_CurrTime);
          end
          else
            showLogMsg('[SMS_SO005Data] Error File : ' + sL_EntitlementPath);

          L_AllDataTxt.Free;
          L_Master.Free;
        end
        else
          showLogMsg('[SMS_SO005Data] File not Exist: ' + sL_EntitlementPath);
      end;
    end;
end;

function TfrmLoadData.isDataOK4: Boolean;
var
    sL_CompCode,sL_EntitlementPath,sL_TransferPath : String;
begin
    IsDataOk4 := true;
    sL_CompCode := cobCompCode.Text;
    if sL_CompCode = '' then
    begin
        MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_1_Content'),mtError, [mbOK],0);
        cobCompCode.SetFocus;
        IsDataOk4 := false;
        exit;
    end;




    sL_EntitlementPath := Trim(medEntitlementPath.Text);
    if sL_EntitlementPath = '' then
    begin
        MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_3_Content'),mtError, [mbOK],0);
        btnEntitlementPath.SetFocus;
        IsDataOk4 := false;
        exit;
    end;




    sL_TransferPath := Trim(medTransferPath.Text);
    if sL_TransferPath = '' then
    begin
        MessageDlg(dtmMain.getLanguageSettingInfo('frmLoadData_Msg_3_Content'),mtError, [mbOK],0);
        btnTransferPath.SetFocus;
        IsDataOk4 := false;
        exit;
    end;
end;

function TfrmLoadData.handleSMSSO005TransTxt(sI_CompCode,
  sI_ModeType: String; dI_CurrTime: TTime): String;
var
    F_SourcePath,F_SO005File,F_SOAC0202File : TextFile;
    sL_EntitlementPath,sL_TransferPath,sL_DiffDateTime,sL_ExecDateTime : String;
    nL_FileNum,nL_CurLineNo,ii,nL_HandleCounts : Integer;
    dL_DateTime2,dL_DiffDateTime : TTime;
    sL_Status,sL_DueDate,sL_SetEn,sL_SetName,sL_ServiceType,sL_SMSModeType,sL_ProductCode : String;
    sL_UpdEn,sL_AuthorStopDate,sL_ReadStr,sL_RealDateTime,sL_ICCardNo,sL_ProductCodeList,sL_ErrMsg : String;
    sL_ChCode,sL_ChName,sL_STBSeqNo,sL_CustID ,sL_AuthorStartDate: String;
    L_Master,L_ProductList : TStringList;
    sL_SO005Str,sL_SOAC0202Str,sL_UpdTime : String;
begin
    sL_EntitlementPath := Trim(medEntitlementPath.Text);
//讀進ICCard測評測試資料
    AssignFile(F_SourcePath,sL_EntitlementPath);
    Reset(F_SourcePath);

//開啟要寫入的二個文字檔
    sL_TransferPath := getTransferPath('SO005',nL_FileNum);
    AssignFile(F_SO005File,sL_TransferPath);
    Append(F_SO005File);

    sL_TransferPath := getTransferPath('SOAC0202',nL_FileNum);
    AssignFile(F_SOAC0202File,sL_TransferPath);
    Append(F_SOAC0202File);

//整理共用參數資料
    L_Master := TStringList.Create;
    L_ProductList := TStringList.Create;
    sL_ExecDateTime := TUdateTime.GetPureDateTimeStr(dI_CurrTime,'','');

    sL_Status := '1';
    sL_DueDate := '20101231';
    sL_SetEn := 'JACKAL';
    sL_SetName := 'PIERREJACKAL';
    sL_ServiceType := 'C';


    sL_SMSModeType := 'B1';
    sL_UpdEn := 'PIERREJACKAL';
    sL_AuthorStopDate := '20101231';

    sL_ErrMsg := getSGSCD024Data(sI_CompCode);


    if sL_ErrMsg = '' then
    begin
      nL_CurLineNo := 0;
      nL_HandleCounts := 0;
      while not Eof(F_SourcePath) do
      begin
        nL_CurLineNo := nL_CurLineNo + 1; //資料筆數控制,50筆資料換一個CustID
        Readln(F_SourcePath,sL_ReadStr);
        L_Master := TUstr.ParseStrings(sL_ReadStr, TAB_STRING);

        //文字檔參數順序(1)授權操作的時間(2)Create為授權,Delete為反授權(3)IC卡號
        //              (4)IC卡內部卡號(5)CASID(6)授權產品代碼組
        sL_RealDateTime := Trim(TUstr.replaceStr(TUstr.replaceStr(L_Master.Strings[0],'-',''),':',''));
        sL_AuthorStartDate := Copy(sL_RealDateTime,1,8);
        sL_UpdTime := Trim(TUstr.replaceStr(L_Master.Strings[0],'-','/'));



        sL_ICCardNo := Trim(L_Master.Strings[2]);
        sL_STBSeqNo := Copy(sL_ICCardNo,5,12);

        sL_ProductCodeList := Trim(L_Master.Strings[5]);
        sL_ErrMsg := dtmMain.getSMSSO005TransCustID(sI_CompCode,sL_ICCardNo,sL_CustID);

        if sL_ErrMsg = '' then
        begin
          L_ProductList := TUstr.ParseStrings(sL_ProductCodeList,',');
          for ii:=0 to L_ProductList.Count-1 do
          begin
            nL_HandleCounts := nL_HandleCounts + 1;
            sL_ProductCode := L_ProductList.Strings[ii];
            sL_ProductCode := sL_ProductCode;
            getSMSChCode(sL_ProductCode,sL_ChCode,sL_ChName);

            //寫入SO005
            sL_SO005Str := sL_CustID + TAB_STRING + sL_ChCode + TAB_STRING +
                           sL_Status + TAB_STRING + sL_RealDateTime + TAB_STRING +
                           sL_STBSeqNo + TAB_STRING + sL_DueDate + TAB_STRING +
                           sL_SetEn + TAB_STRING + sL_SetName + TAB_STRING +
                           sL_ServiceType + TAB_STRING + sI_CompCode + TAB_STRING +
                           sL_ICCardNo;
            Writeln(F_SO005File,sL_SO005Str);


            //寫入SOAC0202
            sL_SOAC0202Str := sI_CompCode + TAB_STRING + sL_CustId + TAB_STRING +
                           sL_STBSeqNo + TAB_STRING + sL_ICCardNo + TAB_STRING +
                           sL_SMSModeType + TAB_STRING + sL_ChCode + TAB_STRING +
                           sL_ChName + TAB_STRING + sL_UpdTime + TAB_STRING +
                           sL_UpdEn + TAB_STRING + sL_AuthorStartDate + TAB_STRING +
                           sL_AuthorStopDate;
            Writeln(F_SOAC0202File,sL_SOAC0202Str);


            if (nL_HandleCounts Mod SHOW_UI_COUNTS)=0 then
            begin
              Flush(F_SO005File);
              Flush(F_SOAC0202File);
              stbStatus.Panels[0].Text := IntToStr((nL_HandleCounts));
              stbStatus.Refresh;
            end;
          end;


          if (nL_CurLineNo Mod SHOW_UI_LINE_COUNTS)=0 then
          begin
            Flush(F_SO005File);
            Flush(F_SOAC0202File);
            stbStatus.Panels[1].Text := IntToStr((nL_CurLineNo));
            stbStatus.Refresh;
          end;
        end;

        L_ProductList.Free;
        L_ProductList := TStringList.Create;
      end;
    end
    else
      showLogMsg('[SMS_SO005Data] Transfer Error :' + sL_ErrMsg);

//關閉結束
    CloseFile(F_SourcePath);
    CloseFile(F_SO005File);
    CloseFile(F_SOAC0202File);
    L_Master.Free;
    L_ProductList.Free;

    dL_DateTime2 := now;
    dL_DiffDateTime := dL_DateTime2 - dI_CurrTime;
    sL_DiffDateTime := TimeToStr(dL_DiffDateTime);

    showLogMsg('[SMS_SO005Data] Txt Transfer OK-[Total Time:' + sL_DiffDateTime + ']');
end;

function TfrmLoadData.getSGSCD024Data(sI_CompCode : String) : String;
var
    L_Obj : TSGSCD024Data;
    sL_SQL,sL_CompName,sL_Msg,sL_ProductCode : String;
    L_AdoQry : TADOQuery;
begin
    sL_SQL := 'SELECT CodeNo AS ChCode,Description AS ChName,ChannelID AS ProductCode ' +
              'FROM CD024 ORDER BY ChannelID';

    L_AdoQry := dtmMain.runSQL(SELECT_MODE,sI_CompCode,sL_SQL);

    if L_AdoQry=nil then
    begin
      sL_CompName := dtmMain.getCompName(sI_CompCode);
      sL_Msg := dtmMain.getLanguageSettingInfo('Return_Msg_100_Content') + '=<' + sL_CompName + '>'; //資料庫不連線, 公司名稱
      Result := sL_Msg;
      Exit;
    end
    else
    begin

      with L_AdoQry do
      begin
        First;
        while not Eof do
        begin
          L_Obj := TSGSCD024Data.Create;

          L_Obj.sChCode := FieldByName('ChCode').AsString;
          L_Obj.sChName := FieldByName('ChName').AsString;

          sL_ProductCode := FieldByName('ProductCode').AsString;
          L_Obj.sProductCode := sL_ProductCode;

          G_SGSCD024DataList.AddObject(sL_ProductCode, L_Obj);

          Next;
        end;
      end;
    end;
end;

procedure TfrmLoadData.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
    G_SGSCD024DataList.Free;
end;

function TfrmLoadData.getSMSChCode(sI_ProductCode: String; var sI_ChCode,
  sI_ChName: String): String;
var
    nL_Ndx : Integer;

begin
    nL_Ndx := G_SGSCD024DataList.IndexOf(sI_ProductCode);
    if nL_Ndx<>-1 then
    begin
      sI_ChCode := (G_SGSCD024DataList.Objects[nL_Ndx] as TSGSCD024Data).sChCode;
      sI_ChName := (G_SGSCD024DataList.Objects[nL_Ndx] as TSGSCD024Data).sChName;

    end;



end;

procedure TfrmLoadData.fraYMD3Exit(Sender: TObject);
begin
    fraYMD4.setYMD(fraYMD3.getYMD);
end;

end.
