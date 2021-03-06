unit frmInvB06_1U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, cxLookAndFeelPainters, dxSkinsCore, dxSkinBlack,
  dxSkinBlue, dxSkinCaramel, dxSkinCoffee, dxSkinsDefaultPainters,
  cxMaskEdit, cxDropDownEdit, cxCalendar, cxTextEdit, cxControls,
  cxContainer, cxEdit, cxGroupBox, cxRadioGroup, StdCtrls, cxButtons,
  Buttons, ExtCtrls, cxStyles, dxSkinscxPCPainter, cxCustomData,
  cxGraphics, cxFilter, cxData, cxDataStorage, DB, cxDBData, cxGridLevel,
  cxClasses, cxGridCustomView, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, cxGrid, cxGridBandedTableView, cxGridDBBandedTableView,
  cxCurrencyEdit, cxLabel,ADODB, DBClient,dtmReportModule;

type
  TfrmInvB06_1 = class(TForm)
    pnlFun1: TPanel;
    Bevel1: TBevel;
    btnQuery: TcxButton;
    btnExit: TcxButton;
    btnCreate: TcxButton;
    BtnCheck: TcxButton;
    rgpQueryCond: TcxRadioGroup;
    mskInvUseYear: TcxMaskEdit;
    Panel1: TPanel;
    gvGrid: TcxGrid;
    gvView: TcxGridDBBandedTableView;
    gvYearMonth: TcxGridDBBandedColumn;
    gvExecuteDate: TcxGridDBBandedColumn;
    gvSPECIALPRIZE1: TcxGridDBBandedColumn;
    gvFIRSTPRIZE2: TcxGridDBBandedColumn;
    gvSPECIALPRIZE3: TcxGridDBBandedColumn;
    gvEXTRAPRIZE1: TcxGridDBBandedColumn;
    gvFIRSTPRIZE1: TcxGridDBBandedColumn;
    gvSPECIALPRIZE2: TcxGridDBBandedColumn;
    gvViewUPTTIME: TcxGridDBBandedColumn;
    gvViewUPTEN: TcxGridDBBandedColumn;
    gvEXTRAPRIZE2: TcxGridDBBandedColumn;
    gvLevel1: TcxGridLevel;
    cmbInvUseMonth: TcxComboBox;
    cxLabel1: TcxLabel;
    gvFIRSTPRIZE3: TcxGridDBBandedColumn;
    adoINV035: TADOQuery;
    dsInv035: TDataSource;
    cdsTmp: TClientDataSet;
    btnPrint: TcxButton;
    adoCheck: TADOQuery;
    pnlStatus: TPanel;
    btnVerify: TcxButton;
    gvVerifyEn: TcxGridDBBandedColumn;
    gvVerifyTime: TcxGridDBBandedColumn;
    gvSPECIALPRIZE4: TcxGridDBBandedColumn;
    gvSPECIALPRIZE5: TcxGridDBBandedColumn;
    gvFIRSTPRIZE4: TcxGridDBBandedColumn;
    gvFIRSTPRIZE5: TcxGridDBBandedColumn;
    gvEXTRAPRIZE3: TcxGridDBBandedColumn;
    gvEXTRAPRIZE4: TcxGridDBBandedColumn;
    gvEXTRAPRIZE5: TcxGridDBBandedColumn;
    rgpRunCompWhere: TcxRadioGroup;
    gvUNUSUALPRIZE1: TcxGridDBBandedColumn;
    gvUNUSUALPRIZE2: TcxGridDBBandedColumn;
    gvUNUSUALPRIZE3: TcxGridDBBandedColumn;
    gvUNUSUALPRIZE4: TcxGridDBBandedColumn;
    gvViewColumn1: TcxGridDBBandedColumn;
    cxButton1: TcxButton;
    procedure btnExitClick(Sender: TObject);
    procedure btnQueryClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure gvViewDblClick(Sender: TObject);
    procedure btnCreateClick(Sender: TObject);

    procedure BtnCheckClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure btnPrintClick(Sender: TObject);
    procedure cmbInvUseMonthPropertiesEditValueChanged(Sender: TObject);
    procedure btnVerifyClick(Sender: TObject);
    procedure cxButton1Click(Sender: TObject);

  private
    { Private declarations }
    FMasterDataSet: TADOQuery;
    FChkInvLst : TStringList;
    FUseCom : Boolean;
    procedure ClearINV007Prize(aWhere : String;blnUseAllComp : Boolean);
    procedure UpdINV007Prize;
    procedure UpdINV035Data;
    procedure PrepareInvTmpDataSet;
    procedure AddPrizeInv(aINVId:String;aPrizeType:String);
    function RetPrizeType(aInvId : String):string;
    
    procedure UnPrepareDataSet;
    procedure PrepareDataSet;

    procedure PrepareINV035DataSet;
    procedure PrintToWin(aWhere : String;aUseAllComp : Boolean);
    function InvDataFillToDataSet(aWhere:String; aUseAllComp : Boolean) : Boolean;
    procedure AfterPrintReport(aUseAllComp : Boolean);
    function ValidateCanResult(strWhere : String):Boolean;

    function CheckPrize:Boolean;
    procedure PrintErrPrize;
    function ChkPrizeNum(aNum : String; aType : Integer) : Boolean;
    procedure TranPrizeTxtLst(aInv007Where : String;aUseAllComp : Boolean;
      var aUseGiveCnt,aNoUseGiveCnt : Integer);
  public
    { Public declarations }
  end;

var
  frmInvB06_1: TfrmInvB06_1;

implementation
  uses dtmMainU,dtmMainJU,dtmSOU,dtmMainHU,
  frmInvB06U,cbUtilis,frmInvB06_2U,frmInvB06_3U;
{$R *.dfm}

procedure TfrmInvB06_1.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmInvB06_1.btnQueryClick(Sender: TObject);
  var
    aWhere,aSQL : String;

begin
  aWhere := EmptyStr;
  aWhere := Format('YEARMONTH = ''%s''',
              [  mskInvUseYear.Text + cmbInvUseMonth.Text ]);
  Screen.Cursor := crSQLWait;
  try
    btnExit.Enabled := False;
    btnCreate.Enabled := False;
    adoINV035.Close;
    adoINV035.SQL.Text := 'SELECT * FROM INV035 WHERE ' + aWhere;
    adoINV035.Open;
    if adoINV035.RecordCount = 0 then
    begin
      MessageDlg( '??????????????',mtInformation,[mbOK],0);
    end;
  finally
    btnExit.Enabled := True;
    btnCreate.Enabled := True;
    Screen.Cursor := crDefault;
  end;
      
end;

procedure TfrmInvB06_1.FormShow(Sender: TObject);
 var aYear,aMonth : String;
    aQuery,aInsert,aDelete,aUpdate,aExecute,aVerify : String;
begin
  aYear := FormatDateTime('yyyy',Now);
  aMonth := FormatDateTime('MM',Now);
  if StrToInt(aMonth) > 2 then
  begin
    if ( StrToInt( aMonth ) mod 2 ) = 0 then
    begin
      aMonth := IntToStr( StrToInt( aMonth ) - 3 )
    end else
      aMonth := IntToStr( StrToInt( aMonth ) - 2 );
    if Length( aMonth ) = 1 then
      aMonth := '0' + aMonth;
  end else
  begin
    aMonth := '11';
    aYear := IntToStr( StrToInt( aYear ) - 1 );
  end;
  mskInvUseYear.Text := aYear;
  cmbInvUseMonth.Text := aMonth;

  FMasterDataSet := dtmMainH.adoEinv;
  //#5760 ?W?[?f?????s?????v?????? By Kin 2010/09/08
  dtmMain.GetCompetence('B06',aQuery,aInsert,aDelete,aUpdate,aExecute,aVerify );
  btnVerify.Enabled := aVerify = 'Y';

end;

procedure TfrmInvB06_1.FormCreate(Sender: TObject);
 var
   aSid,aUserId,aPassword : String;
begin
  FUseCom := False;
  FUseCom := PSMSDb(dtmMain.GetSoInfo( dtmMain.GetConnSeq )).USEPrizeCom;
  if FUseCom then
  begin
    aSid := PSMSDb(dtmMain.GetSoInfo( dtmMain.GetConnSeq )).EInvDb;
    aUserId := PSMSDb(dtmMain.GetSoInfo( dtmMain.GetConnSeq )).EInvUser;
    aPassword := PSMSDb(dtmMain.GetSoInfo( dtmMain.GetConnSeq )).EInvPassword;
    dtmMain.ConnectToEINV( aSid, aUserId , aPassword);
  end else
  begin
    dtmMain.EInvConnection.ConnectionString := dtmMain.InvConnection.ConnectionString;
  end;
  {
  if dtmMain.GetUseCOM then
  begin

    dtmMain.ConnectToEINV( PRIZECOMDB(dtmMain.ComDbList[0] ) );
  end else
  begin
    dtmMain.EInvConnection.ConnectionString := dtmMain.InvConnection.ConnectionString;
  end;
  }
  FChkInvLst := TStringList.Create;
end;

procedure TfrmInvB06_1.gvViewDblClick(Sender: TObject);
  var
    aForm : TfrmInvB06;
    aYearMonth : String;
begin
  Screen.Cursor := crSQLWait;
  try
    if not adoINV035.Active then Exit;
    if adoINV035.IsEmpty then Exit;
    aYearMonth := adoINV035.FieldByName( 'YEARMONTH' ).AsString;
    aForm := TfrmInvB06.Create( aYearMonth );
    try
      aForm.ShowModal;
    finally
        aForm.Free;
        btnQueryClick( btnQuery );
    end;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TfrmInvB06_1.btnCreateClick(Sender: TObject);
  var
    aYearMonth : String;
    aForm : TfrmInvB06;
begin
  aYearMonth := EmptyStr;
  aForm := TfrmInvB06.Create( aYearMonth );
  try
    aForm.ShowModal;
  finally
    aForm.Free;
  end;
end;

function TfrmInvB06_1.ValidateCanResult(strWhere : String): Boolean;
  var
    aWhere : string;
    i :Integer;
begin
  FChkInvLst.Clear;
  aWhere := Format('YEARMONTH = ''%s''',[mskInvUseYear.Text + cmbInvUseMonth.Text]);
  if strWhere <> '' then
    aWhere := aWhere + ' AND ' + strWhere;
  Result := dtmMainH.getInv035Data(aWhere) =1 ;
  for i := 0 to dtmMainH.adoEinv.RecordCount -1 do
  begin

    FChkInvLst.Add( dtmMainH.adoEinv.FieldByName( 'SPECIALPRIZE1' ).AsString);
    FChkInvLst.Add( dtmMainH.adoEinv.FieldByName( 'SPECIALPRIZE2' ).AsString);
    FChkInvLst.Add( dtmMainH.adoEinv.FieldByName( 'SPECIALPRIZE3' ).AsString);
    FChkInvLst.Add( dtmMainH.adoEinv.FieldByName( 'SPECIALPRIZE4' ).AsString);
    FChkInvLst.Add( dtmMainH.adoEinv.FieldByName( 'SPECIALPRIZE5' ).AsString);
    FChkInvLst.Add( dtmMainH.adoEinv.FieldByName( 'FIRSTPRIZE1' ).AsString);
    FChkInvLst.Add( dtmMainH.adoEinv.FieldByName( 'FIRSTPRIZE2' ).AsString);
    FChkInvLst.Add( dtmMainH.adoEinv.FieldByName( 'FIRSTPRIZE3' ).AsString);
    FChkInvLst.Add( dtmMainH.adoEinv.FieldByName( 'FIRSTPRIZE4' ).AsString);
    FChkInvLst.Add( dtmMainH.adoEinv.FieldByName( 'FIRSTPRIZE5' ).AsString);
    FChkInvLst.Add( _Right( Lpad( dtmMainH.adoEinv.FieldByName( 'EXTRAPRIZE1' ).AsString,8,'X'),8 ));
    FChkInvLst.Add( _Right( Lpad( dtmMainH.adoEinv.FieldByName( 'EXTRAPRIZE2' ).AsString,8,'X'),8 ));
    FChkInvLst.Add( _Right( Lpad( dtmMainH.adoEinv.FieldByName( 'EXTRAPRIZE3' ).AsString,8,'X'),8 ));
    FChkInvLst.Add( _Right( Lpad( dtmMainH.adoEinv.FieldByName( 'EXTRAPRIZE4' ).AsString,8,'X'),8 ));
    FChkInvLst.Add( _Right( Lpad( dtmMainH.adoEinv.FieldByName( 'EXTRAPRIZE5' ).AsString,8,'X'),8 ));
    //#5966 ?W?[?S?O?????X By Kin 2011/03/31
    FChkInvLst.Add(dtmMainH.adoEinv.FieldByName( 'UNUSUALPRIZE1' ).AsString );
    FChkInvLst.Add(dtmMainH.adoEinv.FieldByName( 'UNUSUALPRIZE2' ).AsString );
    FChkInvLst.Add(dtmMainH.adoEinv.FieldByName( 'UNUSUALPRIZE3' ).AsString );
    FChkInvLst.Add(dtmMainH.adoEinv.FieldByName( 'UNUSUALPRIZE4' ).AsString );
    FChkInvLst.Add(dtmMainH.adoEinv.FieldByName( 'UNUSUALPRIZE5' ).AsString );
    
  end;
  dtmMainH.adoEinv.Close;
end;

procedure TfrmInvB06_1.BtnCheckClick(Sender: TObject);
var
  aInv007Where, aDate,aInvId,aWhere : string;
  aExecCount:Integer;
  aUseAllComp : Boolean;
  aMsg,aPrtMsg  : string;
  aUseGiveCnt,aNoUseGiveCnt:Integer;

begin


  Screen.Cursor := crSQLWait;
  aUseGiveCnt := 0;
  aNoUseGiveCnt := 0;
  pnlStatus.Caption := '??????....';
  aUseAllComp := True;
  if rgpRunCompWhere.ItemIndex = 0 then
  begin
    aUseAllComp := False;
  end;
  try
    aDate := mskInvUseYear.Text + cmbInvUseMonth.Text;
    PrepareInvTmpDataSet;
    if not ValidateCanResult('') then
    begin
      MessageDlg( '?o???????????|?????I?L?k?????????C',mtInformation,[mbOK],0);
      Exit;
    end;
    if not ValidateCanResult( 'VerifyEn Is not Null And VerifyTime is Not Null' ) then
    begin
      MessageDlg( '???????X???g?L?f???A???i?????I',mtInformation,[mbOK],0);
      Exit;
    end;
    {#5945 ???????A?????d?F}
    {
    if not CheckPrize then
    begin
      MessageDlg( '?P?U?a???????~?I?L?k?????????C',mtInformation,[mbOK],0);
      PrintErrPrize;
      Exit;
    end;
    }
    
    aInv007Where := Format(' AND INVDATE >= TO_DATE(''%s'',''yyyymmdd'')         ' +
                            ' AND INVDATE <= TO_DATE(''%s'',''yyyymmdd'')        ' ,
                            [dtmMainJ.GetFulDate(0,aDate),dtmMainJ.GetFulDate(1,aDate)]);

    //#5966 ?W?[?S?O?????X By Kin 2011/03/31
    aMsg :='?S?O?????X : ' + FChkInvLst.Strings[15] + ' , ' + FChkInvLst.Strings[16] + ' , ' +
            FChkInvLst.Strings[17] + ' , ' + FChkInvLst.Strings[18] + ' , ' +
            FChkInvLst.Strings[19] + #13#10+#13#10;

    aMsg := aMsg + '?S?????X : ' + FChkInvLst.Strings[0] + ' , ' + FChkInvLst.Strings[1] + ' , ' +
          FChkInvLst.Strings[2] + ' , ' + FChkInvLst.Strings[3] + ' , ' +
          FChkInvLst.Strings[4] + #13#10 +#13#10 + '?Y?????X : ' + FChkInvLst.Strings[5] + ' , ' +
          FChkInvLst.Strings[6] + ' , ' + FChkInvLst.Strings[7] + ' , ' +
          FChkInvLst.Strings[8] + ' , ' + FChkInvLst.Strings[9]  + #13#10 +#13#10+
          '?[?}???X : ' + TrimString(FChkInvLst.Strings[10],'X') + ' , ' +
          TrimString(FChkInvLst.Strings[11],'X') + ' , ' +
          TrimString(FChkInvLst.Strings[12],'X') + ' , ' +
          TrimString(FChkInvLst.Strings[13],'X') + ' , ' +
          TrimString(FChkInvLst.Strings[14],'X') +#13#10+ #13#10;
    if rgpRunCompWhere.ItemIndex =0 then
    begin
      aMsg := aMsg + '???q?O???? : [ ???O???q ]';
    end else
    begin
      aMsg := aMsg + '???q?O???? : [ ???????q ]';
    end;
    aMsg := aMsg + '    ?O?_?T?w?????H';

    if MessageDlg( aMsg,mtConfirmation,[mbYes,mbNo],0)= mrNo then
    begin
      Exit;
    end;

    //?Y??
    aWhere := Format( ' AND ( SUBSTR(INVID,8,3) = ''%s''                              ' +
            ' OR SUBSTR(INVID,8,3) = ''%s'' OR SUBSTR(INVID,8,3) = ''%s''             ' +
            ' OR SUBSTR(INVID,8,3) = ''%s'' OR SUBSTR(INVID,8,3) = ''%s''             ' ,
            [ Copy(FChkInvLst.Strings[5],6,3),
             Copy(FChkInvLst.Strings[6],6,3),
             Copy(FChkInvLst.Strings[7],6,3),
             Copy(FChkInvLst.Strings[8],6,3),
             Copy(FChkInvLst.Strings[9],6,3) ] );

    //?[?}???X
    if FChkInvLst.Strings[10] <> '' then
    begin
//      aPos := Length(FChkInvLst.Strings[10]);
      aWhere := aWhere + Format( ' OR SUBSTR(INVID,8,3) = ''%s''' ,
                [ Copy(FChkInvLst.Strings[10],6,3) ] );
    end;
    if FChkInvLst.Strings[11] <> '' then
    begin
      aWhere := aWhere + Format( ' OR SUBSTR(INVID,8,3) = ''%s''' ,
              [ Copy(FChkInvLst.Strings[11],6,3) ] );
    end;
    if FChkInvLst.Strings[12] <> '' then
    begin
      aWhere := aWhere + Format( ' OR SUBSTR(INVID,8,3) = ''%s''' ,
              [ Copy(FChkInvLst.Strings[12],6,3) ] );
    end;
    if FChkInvLst.Strings[13] <> '' then
    begin
      aWhere := aWhere + Format( ' OR SUBSTR(INVID,8,3) = ''%s''' ,
              [ Copy(FChkInvLst.Strings[13],6,3) ] );
    end;
    if FChkInvLst.Strings[14] <> '' then
    begin
      aWhere := aWhere + Format( ' OR SUBSTR(INVID,8,3) = ''%s''' ,
              [ Copy(FChkInvLst.Strings[14],6,3) ] );
    end;


    aWhere := aWhere + ')';
    Application.ProcessMessages;
    if dtmMainH.getInv007Data( aInv007Where + aWhere,aUseAllComp ) > 0 then
    begin
      while not dtmMainH.adoInv007.Eof do
      begin
        aInvId := dtmMainH.adoInv007.FieldByName( 'INVID' ).AsString;
        AddPrizeInv( aInvId,RetPrizeType(Copy(aInvId,3,8)));
        dtmMainH.adoInv007.Next;
      end;
    end;

     //?S??
    aWhere := Format( ' AND ( SUBSTR(INVID,3,8) = ''%s''                              ' +
            ' OR SUBSTR(INVID,3,8) = ''%s'' OR SUBSTR(INVID,3,8) = ''%s''             ' +
            ' OR SUBSTR(INVID,3,8) = ''%s'' OR SUBSTR(INVID,3,8) = ''%s'' )           ' ,
            [ FChkInvLst.Strings[0],FChkInvLst.Strings[1],FChkInvLst.Strings[2],
              FChkInvLst.Strings[3],FChkInvLst.Strings[4] ]);
    Application.ProcessMessages;
    if dtmMainH.getInv007Data( aInv007Where + aWhere,aUseAllComp ) > 0 then
    begin
      while not dtmMainH.adoInv007.Eof do
      begin
        AddPrizeInv( dtmMainH.adoInv007.FieldByName( 'INVID' ).AsString,'0' );
        dtmMainH.adoInv007.Next;
      end;

    end;

    //#5966 ?W?[?S?O?????X By Kin 2011/03/31
    aWhere := Format( ' AND ( SUBSTR(INVID,3,8) = ''%s''                              ' +
            ' OR SUBSTR(INVID,3,8) = ''%s'' OR SUBSTR(INVID,3,8) = ''%s''             ' +
            ' OR SUBSTR(INVID,3,8) = ''%s'' OR SUBSTR(INVID,3,8) = ''%s'' )           ' ,
            [ FChkInvLst.Strings[15],FChkInvLst.Strings[16],FChkInvLst.Strings[17],
              FChkInvLst.Strings[18],FChkInvLst.Strings[19] ]);

    Application.ProcessMessages;
    if dtmMainH.getInv007Data( aInv007Where + aWhere,aUseAllComp ) > 0 then
    begin
      while not dtmMainH.adoInv007.Eof do
      begin
        AddPrizeInv( dtmMainH.adoInv007.FieldByName( 'INVID' ).AsString,'A' );
        dtmMainH.adoInv007.Next;
      end;

    end;

    if cdsTmp.IsEmpty then
    begin
      UpdINV035Data;
      MessageDlg( '?L?????????o??',mtInformation,[mbOK],0 );
      Exit;
    end;
    try
      dtmMain.EInvConnection.BeginTrans;
      dtmMain.InvConnection.BeginTrans;
      UpdINV035Data;
      ClearINV007Prize( aInv007Where,aUseAllComp  );
      UpdINV007Prize;
      dtmMain.EInvConnection.CommitTrans;
      dtmMain.InvConnection.CommitTrans;
    except
      on E: Exception do
      begin
        dtmMain.EInvConnection.RollbackTrans;
        dtmMain.InvConnection.RollbackTrans;
        MessageDlg( '?^???o?????O???~?I,???]?G' + E.Message,mtError,[mbOK],0 );
      end;
    end;
    //#5874 ?C?L???????? By Kin 2010/12/13
    TranPrizeTxtLst( aInv007Where + ' AND PRIZETYPE IS NOT NULL ',
            aUseAllComp ,aUseGiveCnt,aNoUseGiveCnt);

  {
    if aUseAllComp then
    begin
      TranPrizeTxtLst( aInv007Where,aUseAllComp );
    end else
    begin
      TranPrizeTxtLst(  Format( aInv007Where + ' AND INV007.COMPID =''%s''',[dtmMain.getCompID]),aUseAllComp);
    end;
   }

    aPrtMsg :=  Format('???C?L???????o???@?G[ %d ] ?i '+#13#10  + #13#10 +
                      '?????G[ %d ] ?i' + #13#10 + #13#10 +
                      '?D?????G[ %d ] ?i',
                      [aUseGiveCnt + aNoUseGiveCnt,aUseGiveCnt,aNoUseGiveCnt] );
    if (aUseGiveCnt + aNoUseGiveCnt) > 0 then
    begin
      if MessageDlg( aPrtMsg,mtConfirmation,[mbYes,mbNo],0) = mrYes then
      begin
        Application.ProcessMessages;
        PrintToWin(  aInv007Where + ' AND PRIZETYPE IS NOT NULL ',aUseAllComp );
      end;
    end else
    begin
      MessageDlg('?L?????????i?C?L?I',mtInformation,[mbOK],0);
    end;
    MessageDlg( '?????????I',mtInformation,[mbOK],0);
  finally
    pnlStatus.Caption := EmptyStr;
    Screen.Cursor := crDefault;
  end;
end;


procedure TfrmInvB06_1.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin
  FChkInvLst.Clear;
  FChkInvLst.Free;
  FChkInvLst := nil;
  dtmMain.EInvConnection.Connected := False;
  dtmMain.EInvConnection.Close;
end;

procedure TfrmInvB06_1.PrepareInvTmpDataSet;
begin
  if ( cdsTmp.FieldDefs.Count <= 0 ) then
  begin
    cdsTmp.FieldDefs.Add( 'INVID',ftString,10 );
    cdsTmp.FieldDefs.Add( 'PRIZETYPE',ftString,10 );
    cdsTmp.CreateDataSet;
  end;
  cdsTmp.EmptyDataSet;
end;

procedure TfrmInvB06_1.AddPrizeInv(aINVId: String; aPrizeType: string);
begin
  if aPrizeType = EmptyStr then  Exit;
  cdsTmp.Append;
  cdsTmp.FieldByName( 'INVID' ).AsString := aINVId;
  cdsTmp.FieldByName( 'PRIZETYPE' ).AsString := aPrizeType;
  cdsTmp.Post;
end;

function TfrmInvB06_1.RetPrizeType(aInvId : String): String;
  var i:Integer;
    aInvTmp,aChkTmp : string;
begin
  Result := '-1';
  if ( aInvId = FChkInvLst.Strings[ 5 ]) or ( aInvId = FChkInvLst.Strings[ 6 ])  or
      ( aInvId = FChkInvLst.Strings[ 7 ]) or ( aInvId = FChkInvLst.Strings[ 8 ]) or
      ( aInvId = FChkInvLst.Strings[ 9 ]) or ( aInvId = FChkInvLst.Strings[10]) or
      ( aInvId = FChkInvLst.Strings[11]) or ( aInvId = FChkInvLst.Strings[12] ) or
      ( aInvId = FChkInvLst.Strings[13] ) or ( aInvId = FChkInvLst.Strings[14] ) then
  begin
    Result := '1';
    Exit;
  end;
  if (Copy(aInvId,2,7) = Copy(FChkInvLst.Strings[ 5 ],2,7)) or
    (Copy(aInvId,2,7) = Copy(FChkInvLst.Strings[ 6 ],2,7)) or
    (Copy(aInvId,2,7) = Copy(FChkInvLst.Strings[ 7 ],2,7)) or
    (Copy(aInvId,2,7) = Copy(FChkInvLst.Strings[ 8 ],2,7)) or
    (Copy(aInvId,2,7) = Copy(FChkInvLst.Strings[ 9 ],2,7)) or
    (Copy(aInvId,2,7) = Copy(FChkInvLst.Strings[ 10 ],2,7)) or
    (Copy(aInvId,2,7) = Copy(FChkInvLst.Strings[ 11 ],2,7)) or
    (Copy(aInvId,2,7) = Copy(FChkInvLst.Strings[ 12 ],2,7)) or
    (Copy(aInvId,2,7) = Copy(FChkInvLst.Strings[ 13 ],2,7)) or
    (Copy(aInvId,2,7) = Copy(FChkInvLst.Strings[ 14 ],2,7))  then
  begin
    Result := '2';
    Exit;
  end;
  if (Copy(aInvId,3,6) = Copy(FChkInvLst.Strings[ 5 ],3,6)) or
    (Copy(aInvId,3,6) = Copy(FChkInvLst.Strings[ 6 ],3,6)) or
    (Copy(aInvId,3,6) = Copy(FChkInvLst.Strings[ 7 ],3,6)) or
    (Copy(aInvId,3,6) = Copy(FChkInvLst.Strings[ 8 ],3,6)) or
    (Copy(aInvId,3,6) = Copy(FChkInvLst.Strings[ 9 ],3,6)) or
    (Copy(aInvId,3,6) = Copy(FChkInvLst.Strings[ 10 ],3,6)) or
    (Copy(aInvId,3,6) = Copy(FChkInvLst.Strings[ 11 ],3,6)) or
    (Copy(aInvId,3,6) = Copy(FChkInvLst.Strings[ 12 ],3,6)) or
    (Copy(aInvId,3,6) = Copy(FChkInvLst.Strings[ 13 ],3,6)) or
    (Copy(aInvId,3,6) = Copy(FChkInvLst.Strings[ 14 ],3,6))  then
  begin
    Result := '3';
    Exit;
  end;
  if (Copy(aInvId,4,5) = Copy(FChkInvLst.Strings[ 5 ],4,5)) or
    (Copy(aInvId,4,5) = Copy(FChkInvLst.Strings[ 6 ],4,5)) or
    (Copy(aInvId,4,5) = Copy(FChkInvLst.Strings[ 7 ],4,5)) or
    (Copy(aInvId,4,5) = Copy(FChkInvLst.Strings[ 8 ],4,5)) or
    (Copy(aInvId,4,5) = Copy(FChkInvLst.Strings[ 9 ],4,5)) or
    (Copy(aInvId,4,5) = Copy(FChkInvLst.Strings[ 10 ],4,5)) or
    (Copy(aInvId,4,5) = Copy(FChkInvLst.Strings[ 11 ],4,5)) or
    (Copy(aInvId,4,5) = Copy(FChkInvLst.Strings[ 12 ],4,5)) or
    (Copy(aInvId,4,5) = Copy(FChkInvLst.Strings[ 13 ],4,5)) or
    (Copy(aInvId,4,5) = Copy(FChkInvLst.Strings[ 14 ],4,5)) then
  begin
    Result := '4';
    Exit;
  end;
  if (Copy(aInvId,5,4) = Copy(FChkInvLst.Strings[ 5 ],5,4)) or
    (Copy(aInvId,5,4) = Copy(FChkInvLst.Strings[ 6 ],5,4)) or
    (Copy(aInvId,5,4) = Copy(FChkInvLst.Strings[ 7 ],5,4)) or
    (Copy(aInvId,5,4) = Copy(FChkInvLst.Strings[ 8 ],5,4)) or
    (Copy(aInvId,5,4) = Copy(FChkInvLst.Strings[ 9 ],5,4)) or
    (Copy(aInvId,5,4) = Copy(FChkInvLst.Strings[ 10 ],5,4)) or
    (Copy(aInvId,5,4) = Copy(FChkInvLst.Strings[ 11 ],5,4)) or
    (Copy(aInvId,5,4) = Copy(FChkInvLst.Strings[ 12 ],5,4)) or
    (Copy(aInvId,5,4) = Copy(FChkInvLst.Strings[ 13 ],5,4)) or
    (Copy(aInvId,5,4) = Copy(FChkInvLst.Strings[ 14 ],5,4)) then
  begin
    Result := '5';
    Exit;
  end;
  if (Copy(aInvId,6,3) = Copy(FChkInvLst.Strings[ 5 ],6,3)) or
    (Copy(aInvId,6,3) = Copy(FChkInvLst.Strings[ 6 ],6,3)) or
    (Copy(aInvId,6,3) = Copy(FChkInvLst.Strings[ 7 ],6,3)) or
    (Copy(aInvId,6,3) = Copy(FChkInvLst.Strings[ 8 ],6,3)) or
    (Copy(aInvId,6,3) = Copy(FChkInvLst.Strings[ 9 ],6,3)) or
    (Copy(aInvId,6,3) = Copy(FChkInvLst.Strings[ 10 ],6,3)) or
    (Copy(aInvId,6,3) = Copy(FChkInvLst.Strings[ 11 ],6,3)) or
    (Copy(aInvId,6,3) = Copy(FChkInvLst.Strings[ 12 ],6,3)) or
    (Copy(aInvId,6,3) = Copy(FChkInvLst.Strings[ 13 ],6,3)) or
    (Copy(aInvId,6,3) = Copy(FChkInvLst.Strings[ 14 ],6,3)) then
  begin
    Result := '6';
    Exit;
  end;


end;

procedure TfrmInvB06_1.ClearINV007Prize(aWhere : String;blnUseAllComp : Boolean);
var
  aSQL : String;
begin
  aSQL := format('UPDATE INV007 SET PRIZETYPE = NULL                  ' +
        '   WHERE IDENTIFYID1 = ''%s''                                ' +
        '     AND IDENTIFYID2 = %s                                    ' +
        '     AND PRIZETYPE IS NOT NULL                               ' +
        '     %s                                                      ' ,
    [IDENTIFYID1, IDENTIFYID2,  aWhere] );
  if (not blnUseAllComp ) then
  begin
    aSQL := Format( aSQL + '     AND COMPID = ''%s'' ',
          [dtmMain.getCompID] );
  end;
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := aSQL;
  dtmMain.adoComm.ExecSQL;
end;

procedure TfrmInvB06_1.UpdINV007Prize;
var
   aSQL : String;
begin
    if not cdsTmp.IsEmpty then
    begin
      cdsTmp.First;
      while not cdsTmp.Eof do
      begin
        aSQL := EmptyStr;
        aSQL := Format('UPDATE INV007 SET PRIZETYPE = ''%s''                ' +
              '   WHERE IDENTIFYID1 = ''%s''                                ' +
              '     AND IDENTIFYID2 = %s                                    ' +
              '     AND INVID= ''%s''                                       ' ,
        [cdsTmp.FieldByName( 'PRIZETYPE' ).AsString ,
        IDENTIFYID1, IDENTIFYID2,
        cdsTmp.FieldByName('INVID').AsString] );
        dtmMain.adoComm.Close;
        dtmMain.adoComm.SQL.Text := aSQL;
        dtmMain.adoComm.ExecSQL;
        cdsTmp.Next;
      end;
    end;

end;

procedure TfrmInvB06_1.UpdINV035Data;
var
  aSQL : String;
begin
  aSQL := format('UPDATE INV035 SET EXECUTEDATE = SYSDATE,     ' +
            'UPTTIME = SYSDATE , UPTEN = ''%s''                ' +
            'WHERE YEARMONTH = ''%s''                          ' ,
            [ dtmMain.getLoginUser,
            mskInvUseYear.Text + cmbInvUseMonth.Text ] );
  dtmMain.adoEComm.Close;
  dtmMain.adoEComm.SQL.Text := aSQL;
  dtmMain.adoEComm.ExecSQL;
end;

procedure TfrmInvB06_1.UnPrepareDataSet;
begin
  if not VarIsNull( dtmReport.cdsTempory.Data ) then
    dtmReport.cdsTempory.Data := Null;
  dtmReport.cdsTempory.FieldDefs.Clear;
end;

procedure TfrmInvB06_1.PrepareDataSet;
var
  aAutoCreateNum, aIndex: Integer;
begin
  dtmReport.cdsTempory.FieldDefs.Clear;
  dtmReport.cdsTempory.FieldDefs.Add( 'CHECKNO', ftString, 2 );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVID', ftString, 10 );
  dtmReport.cdsTempory.FieldDefs.Add( 'BUSINESSID', ftString, 8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVDATE', ftDateTime );
  dtmReport.cdsTempory.FieldDefs.Add( 'CUSTID', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'CUSTSNAME', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVTITLE', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'ZIPCODE', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'MAILADDR', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVADDR', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'INSTADDR', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVFORMAT', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'TAXTYPE', ftString, 100 );
  dtmReport.cdsTempory.FieldDefs.Add( 'SALEAMOUNT', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'TAXAMOUNT', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVAMOUNT', ftInteger );
  {}
  dtmReport.cdsTempory.FieldDefs.Add( 'MAININVID', ftString, 10 );
  dtmReport.cdsTempory.FieldDefs.Add( 'MAINSALEAMOUNT', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'MAINTAXAMOUNT', ftInteger );
  dtmReport.cdsTempory.FieldDefs.Add( 'MAININVAMOUNT', ftInteger );
  {}
  aAutoCreateNum := dtmMain.GetAutoCreateNum;
  if aAutoCreateNum <= 0 then aAutoCreateNum := 10;
  {}
  for aIndex := 1 to aAutoCreateNum do
  begin
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'BILLID%d', [aIndex] ), ftString, 15 );
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'DESCRIPTION%d', [aIndex] ), ftString, 200 );
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'STARTDATE%d', [aIndex] ), ftDateTime );
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'ENDDATE%d', [aIndex] ), ftDateTime );
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'QUANTITY%d', [aIndex] ), ftInteger );
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'UNITPRICE%d', [aIndex] ), ftFloat );  //#4370 ?????B?I?? By Kin 2009/02/19
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'SALEAMOUNT%d', [aIndex] ), ftInteger );
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'TOTALAMOUNT%d', [aIndex] ), ftInteger );
    dtmReport.cdsTempory.FieldDefs.Add( Format( 'FACISNO%d', [aIndex] ), ftString, 20 );
  end;
  {}
  dtmReport.cdsTempory.FieldDefs.Add( 'MEMO1', ftString, 60 );
  dtmReport.cdsTempory.FieldDefs.Add( 'ACCOUNTNO', ftString, 1000 );
  dtmReport.cdsTempory.FieldDefs.Add( 'HOWTOCREATE', ftString, 1 );
  dtmReport.cdsTempory.FieldDefs.Add( 'INVUSEDESC', ftString, 20 );
  {}
  dtmReport.cdsTempory.CreateDataSet;  
end;

procedure TfrmInvB06_1.PrintToWin(aWhere: String;aUseAllComp : Boolean);
var
  aPath, aPrintTitleSet: String ;
  aPrintTitle, aPrintMailAddr, aPrintMemo: String;
  aExecCount :Integer;
begin
  {
  dtmMainJ.getPrintInvoiceTempData(EmptyStr,EmptyStr,EmptyStr,
                        EmptyStr,'2','1',EmptyStr,aWhere,EmptyStr,aUseAllComp);
  }
  dtmMainJ.adoQryCommon.Close;
  dtmMainJ.adoQryCommon.SQL.Text := Format(
    ' SELECT COUNT(1) AS COUNTS FROM INV023 A  ' +
    '    WHERE A.IDENTIFYID1 = ''%s''          ' +
    '      AND A.IDENTIFYID2 = ''%s''          ' +
    '      AND A.OWNER = ''%s''                ',
    [IDENTIFYID1, IDENTIFYID2,  dtmMain.getLoginUserName] );
  if ( not aUseAllComp ) then
  begin
    dtmMainJ.adoQryCommon.SQL.Text := Format(dtmMainJ.adoQryCommon.SQL.Text +
            ' AND COMPID = ''%s'' ',[dtmMain.getCompID]);
  end;

  dtmMainJ.adoQryCommon.Open;
  aExecCount := dtmMainJ.adoQryCommon.FieldByName( 'COUNTS' ).AsInteger;
  dtmMainJ.adoQryCommon.Close;

  if ( aExecCount <= 0 ) then
  begin
    WarningMsg( '?L?????????o???????i???M?L?C' );
    Exit;
  end;

  if not ConfirmMsg( '???T?{?z???L???????X???i?j?p?]?w?O?_???T?C' ) then Exit;

  UnPrepareDataSet;
  PrepareDataSet;

  if not InvDataFillToDataSet(aWhere,aUseAllComp) then
  begin
    WarningMsg( '?S???o???????i???C?L!' );
    Exit;
  end;
  aPrintTitleSet := dtmMainJ.getIsPrintTitle( dtmMain.getCompID );
  if aPrintTitleSet = EmptyStr then aPrintTitle := 'Y';
  dtmReport.AddA05DataField( dtmMain.GetAutoCreateNum );
  aPath := IncludeTrailingPathDelimiter( ExtractFilePath(
      Application.ExeName ) ) + IncludeTrailingPathDelimiter(
      REPORT_FOLDER ) + 'FrptInvA05_1.fr3';
  dtmReport.frxMainReport.LoadFromFile( aPath );
  dtmReport.frxMainReport.Variables.Variables['?C?L?o?????Y'] :=
    QuotedStr( aPrintTitleSet );
  dtmReport.frxMainReport.ShowReport;
  AfterPrintReport(aUseAllComp);
  dtmReport.cdsTempory.EmptyDataSet;
end;

function TfrmInvB06_1.InvDataFillToDataSet(aWhere: String;aUseAllComp : Boolean) : Boolean;
var
  aSQL, aLastInvId, aDaySt, aDayEd, aDescription, aFacisNo,aSQLGive,aSQLCompWhere : String;
  aCurrentFieldNumber: Integer;
begin
  if aUseAllComp then
  begin
    aSQLCompWhere := 'SELECT INVID FROM INV023';
  end else
  begin
    aSQLCompWhere := Format( 'SELECT INVID FROM INV023 WHERE COMPID = ''%s''',[dtmMain.getCompID]);
  end;
  aSQLGive := 'SELECT INVID,Nvl(GIVEUNITID,9999) GIVEUNITID,GIVEUNITDESC,                    ' +
            ' 0 USEGIVE FROM INV007  ,INV028                                                 ' +
            ' WHERE   PRIZETYPE IS NOT NULL                                                  ' +
            ' AND INV007.COMPID = INV028.COMPID                                              ' +
            ' AND INVUSEID = ITEMID AND INV028.REFNO = 1                                     ' +
            aWhere                                                                             +
            ' UNION ALL                                                                      ' +
            ' SELECT INVID,9999 GIVEUNITID,NULL GIVEUNITDESC,1 USEGIVE FROM INV007,INV028    ' +
            ' WHERE PRIZETYPE IS NOT NULL                                                    ' +
            aWhere                                                                             +
            ' AND NVL(INVUSEID,0) = ITEMID(+) AND NVL(INV028.REFNO,0) <> 1 ';
  aSQLGive := 'SELECT DISTINCT * FROM ( ' + aSQLGive + ' )';





  aSQL := Format(
      ' SELECT A.CHECKNO, A.INVID, A.BUSINESSID, A.INVDATE, A.CUSTID, ' +
      '        A.CUSTSNAME, A.INVFORMAT, A.INVTITLE, A.ZIPCODE,       ' +
      '        A.MAILADDR, A.TAXTYPE, A.SALEAMOUNT, A.TAXAMOUNT,      ' +
      '        A.INVAMOUNT, A.MEMO1, B.DESCRIPTION, B.STARTDATE,      ' +
      '        B.ENDDATE, B.QUANTITY, B.UNITPRICE, B.TOTALAMOUNT,     ' +
      '        B.SEQ, A.MAININVID, A.INVADDR,                         ' +
      '        A.MAINSALEAMOUNT, A.MAINTAXAMOUNT, A.MAININVAMOUNT,    ' +
      '        B.SALEAMOUNT AS SALEAMOUNT2, B.BILLID, A.INSTADDR,     ' +
      '        A.HOWTOCREATE, A.INVUSEDESC, C.IFPRINTTITLE,           ' +
      '        D.USEGIVE,A.ISOBSOLETE, A.COMPID                       ' +
      '   FROM INV023 A, INV008 B, INV002 C, (' + aSQLGive + ') D     ' +
      '  WHERE A.IDENTIFYID1 = B.IDENTIFYID1                          ' +
      '    AND A.IDENTIFYID2 = B.IDENTIFYID2                          ' +
      '    AND A.INVID = B.INVID                                      ' +
      '    AND A.IDENTIFYID1 = C.IDENTIFYID1(+)                       ' +
      '    AND A.IDENTIFYID2 = C.IDENTIFYID2(+)                       ' +
      '    AND A.COMPID = C.COMPID(+)                                 ' +
      '    AND A.CUSTID = C.CUSTID(+)                                 ' +
      '    AND A.IDENTIFYID1 = ''%s''                                 ' +
      '    AND A.IDENTIFYID2 = ''%s''                                 ' +
      '    AND A.INVID = D.INVID                                      ' +
      '    AND A.OWNER = ''%s''                                       ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getLoginUserName ] );

  if ( not aUseAllComp) then
  begin
    aSQL := Format( aSQL + ' AND A.COMPID = ''%s'' ',[dtmMain.getCompID]);
  end;


  aSQL := ( aSQL + ' ORDER BY D.USEGIVE,D.GIVEUNITID,A.INVID,B.SEQ ');

  aSQL := 'SELECT * FROM (' + aSQL + ') ' +
      ' WHERE BUSINESSID IS NULL AND ISOBSOLETE <> ''Y'' ' +
      ' AND   TAXTYPE <> 2 ';
//  aSQL := ( aSQL + ' ORDER BY A.INVID, B.SEQ ' );
  dtmMain.adoComm.Close;
  dtmMain.adoComm.SQL.Text := aSQL;
  dtmMain.adoComm.Open;
  Result := not dtmMain.adoComm.IsEmpty;
  if not Result then
  begin
    dtmMain.adoComm.Close;
    Exit;
  end;
  aLastInvId := EmptyStr;
  aCurrentFieldNumber := 0;
  dtmMain.adoComm.First;
  while not dtmMain.adoComm.Eof do
  begin
    if ( dtmMain.adoComm.FieldByName( 'InvAmount' ).AsFloat =
             dtmMainH.getInv014Amt(dtmMain.adoComm.FieldByName( 'INVID' ).AsString,
                                  dtmMain.adoComm.FieldByName( 'COMPID' ).AsString )) then
    begin

    end else
    begin
      if ( aLastInvId <> dtmMain.adoComm.FieldByName( 'INVID' ).AsString ) then
      begin
        aLastInvId := dtmMain.adoComm.FieldByName( 'INVID' ).AsString;
        aCurrentFieldNumber := 1;
        dtmReport.cdsTempory.Append;
      end else
      begin
        Inc( aCurrentFieldNumber );
        dtmReport.cdsTempory.Edit;
      end;
      if ( aCurrentFieldNumber <= 1 ) then
      begin
        dtmReport.cdsTempory.FieldByName( 'CHECKNO' ).AsString :=
          dtmMain.adoComm.FieldByName( 'CHECKNO' ).AsString;
        {}
        dtmReport.cdsTempory.FieldByName( 'INVID' ).AsString :=
          dtmMain.adoComm.FieldByName( 'INVID' ).AsString;
        {}
        dtmReport.cdsTempory.FieldByName( 'BUSINESSID' ).AsString :=
          dtmMain.adoComm.FieldByName( 'BUSINESSID' ).AsString;
        {}
        dtmReport.cdsTempory.FieldByName( 'INVDATE' ).Value :=
          dtmMain.adoComm.FieldByName( 'INVDATE' ).Value;
        {}
        dtmReport.cdsTempory.FieldByName( 'CUSTID' ).AsString :=
          dtmMain.adoComm.FieldByName( 'CUSTID' ).AsString;
        {}
        dtmReport.cdsTempory.FieldByName( 'CUSTSNAME' ).AsString :=
          dtmMain.adoComm.FieldByName( 'CUSTSNAME' ).AsString;
        {}
        { ???????@?s???O?_????, ?????J INVTITLE }
        if dtmMain.GetIfPrintTitle then
          dtmReport.cdsTempory.FieldByName( 'INVTITLE' ).AsString :=
            dtmMain.adoComm.FieldByName( 'INVTITLE' ).AsString
        else begin
          { ???@?s??????, ?~???J INVTITLE }
          dtmReport.cdsTempory.FieldByName( 'INVTITLE' ).AsString := EmptyStr;
          if ( dtmMain.adoComm.FieldByName( 'BUSINESSID' ).AsString <> EmptyStr ) then
            dtmReport.cdsTempory.FieldByName( 'INVTITLE' ).AsString :=
              dtmMain.adoComm.FieldByName( 'INVTITLE' ).AsString
          { ?Y???s?S??, ???O INV002 ?? IfPrintTitle = Y ???M?n?L }
          else if ( dtmMain.adoComm.FieldByName( 'IFPRINTTITLE' ).AsString = 'Y' ) then
            dtmReport.cdsTempory.FieldByName( 'INVTITLE' ).AsString :=
              dtmMain.adoComm.FieldByName( 'INVTITLE' ).AsString;
        end;
        {}
        dtmReport.cdsTempory.FieldByName( 'ZIPCODE' ).AsString :=
          dtmMain.adoComm.FieldByName( 'ZIPCODE' ).AsString;
        // ?l?H?a?}
        dtmReport.cdsTempory.FieldByName( 'MAILADDR' ).AsString :=
          dtmMain.adoComm.FieldByName( 'MAILADDR' ).AsString;
        // ?o???a?}
        dtmReport.cdsTempory.FieldByName( 'INVADDR' ).AsString :=
          dtmMain.adoComm.FieldByName( 'INVADDR' ).AsString;
        // ?????a?}
        dtmReport.cdsTempory.FieldByName( 'INSTADDR' ).AsString :=
          dtmMain.adoComm.FieldByName( 'INSTADDR' ).AsString;
        {}
        dtmReport.cdsTempory.FieldByName( 'INVFORMAT' ).AsString :=
          dtmMain.adoComm.FieldByName( 'INVFORMAT' ).AsString;
        {}
        dtmReport.cdsTempory.FieldByName( 'TAXTYPE' ).AsString :=
          dtmMain.adoComm.FieldByName( 'TAXTYPE' ).AsString;
        {}
        dtmReport.cdsTempory.FieldByName( 'SALEAMOUNT' ).AsInteger :=
          dtmMain.adoComm.FieldByName( 'SALEAMOUNT' ).AsInteger;
        {}
        dtmReport.cdsTempory.FieldByName( 'TAXAMOUNT' ).AsInteger :=
          dtmMain.adoComm.FieldByName( 'TAXAMOUNT' ).AsInteger;
        {}
        dtmReport.cdsTempory.FieldByName( 'INVAMOUNT' ).AsInteger :=
          dtmMain.adoComm.FieldByName( 'INVAMOUNT' ).AsInteger;

        dtmReport.cdsTempory.FieldByName( 'MAININVID' ).AsString :=
          dtmMain.adoComm.FieldByName( 'MAININVID' ).AsString;
        {}
        dtmReport.cdsTempory.FieldByName( 'MAINSALEAMOUNT' ).AsInteger :=
          dtmMain.adoComm.FieldByName( 'MAINSALEAMOUNT' ).AsInteger;
        {}
        dtmReport.cdsTempory.FieldByName( 'MAINTAXAMOUNT' ).AsInteger :=
          dtmMain.adoComm.FieldByName( 'MAINTAXAMOUNT' ).AsInteger;
        {}
        dtmReport.cdsTempory.FieldByName( 'MAININVAMOUNT' ).AsInteger :=
          dtmMain.adoComm.FieldByName( 'MAININVAMOUNT' ).AsInteger;
        {}
        dtmReport.cdsTempory.FieldByName( 'MEMO1' ).AsString :=
          dtmMain.adoComm.FieldByName( 'MEMO1' ).AsString;
        {}
        dtmReport.cdsTempory.FieldByName( 'ACCOUNTNO' ).AsString := dtmMainJ.GetAccountNo(
          dtmReport.cdsTempory.FieldByName( 'INVID' ).AsString,
          EmptyStr, EmptyStr, EmptyStr );
        {}
        dtmReport.cdsTempory.FieldByName( 'HOWTOCREATE' ).AsString :=
          dtmMain.adoComm.FieldByName( 'HOWTOCREATE' ).AsString;
        {}
        dtmReport.cdsTempory.FieldByName( 'INVUSEDESC' ).AsString :=
          dtmMain.adoComm.FieldByName( 'INVUSEDESC' ).AsString;
      end;
      {}
      dtmReport.cdsTempory.FieldByName( Format( 'BILLID%d', [aCurrentFieldNumber] ) ).Value :=
        dtmMain.adoComm.FieldByName( 'BILLID' ).Value;
      {}
      aDescription := dtmMain.adoComm.FieldByName( 'DESCRIPTION' ).AsString;
      aDaySt := EmptyStr;
      aDayEd := EmptyStr;
      if not VarIsNull( dtmMain.adoComm.FieldByName( 'STARTDATE' ).Value ) then
        aDaySt := FormatDateTime( 'eee/mm/dd', dtmMain.adoComm.FieldByName( 'STARTDATE' ).AsDateTime );
      if not VarIsNull( dtmMain.adoComm.FieldByName( 'ENDDATE' ).Value ) then
        aDayEd := FormatDateTime( 'eee/mm/dd', dtmMain.adoComm.FieldByName( 'ENDDATE' ).AsDateTime );
      aDescription := Format( aDescription + '( %s - %s )', [aDaySt, aDayEd] );
      {}
      dtmReport.cdsTempory.FieldByName( Format( 'DESCRIPTION%d', [aCurrentFieldNumber] ) ).AsString :=
        dtmMain.adoComm.FieldByName( 'DESCRIPTION' ).AsString;
      {}
      dtmReport.cdsTempory.FieldByName( Format( 'STARTDATE%d', [aCurrentFieldNumber] ) ).Value :=
        dtmMain.adoComm.FieldByName( 'STARTDATE' ).Value;
      {}
      dtmReport.cdsTempory.FieldByName( Format( 'ENDDATE%d', [aCurrentFieldNumber] ) ).Value :=
        dtmMain.adoComm.FieldByName( 'ENDDATE' ).Value;
      {}
      dtmReport.cdsTempory.FieldByName( Format( 'QUANTITY%d', [aCurrentFieldNumber] ) ).Value :=
        dtmMain.adoComm.FieldByName( 'QUANTITY' ).AsInteger;
      {}
      dtmReport.cdsTempory.FieldByName( Format( 'UNITPRICE%d', [aCurrentFieldNumber] ) ).Value :=
        dtmMain.adoComm.FieldByName( 'UNITPRICE' ).Value;
      {}
      dtmReport.cdsTempory.FieldByName( Format( 'SALEAMOUNT%d', [aCurrentFieldNumber] ) ).Value :=
        dtmMain.adoComm.FieldByName( 'SALEAMOUNT2' ).Value;
      {}
      dtmReport.cdsTempory.FieldByName( Format( 'TOTALAMOUNT%d', [aCurrentFieldNumber] ) ).Value :=
        dtmMain.adoComm.FieldByName( 'TOTALAMOUNT' ).Value;
      {}
      aFacisNo := dtmMainJ.GetFacisNo( dtmMain.adoComm.FieldByName( 'INVID' ).AsString,
          dtmMain.adoComm.FieldByName( 'SEQ' ).AsString );
      if ( Pos( ',', aFacisNo ) > 0 ) then aFacisNo := Copy( aFacisNo, 1, Pos( ',', aFacisNo ) - 1 );
      aFacisNo := dtmMainU.FacisNoAddMask( aFacisNo );
      dtmReport.cdsTempory.FieldByName( Format( 'FACISNO%d', [aCurrentFieldNumber] ) ).Value := aFacisNo;
      {}
      dtmReport.cdsTempory.Post;
    end;





    dtmMain.adoComm.Next;
  end;

  dtmMain.adoComm.Close;
end;

procedure TfrmInvB06_1.AfterPrintReport(aUseAllComp : Boolean);
begin
  dtmMainJ.AfterPrintOrTrans( 1, nil,aUseAllComp );
end;

procedure TfrmInvB06_1.PrepareINV035DataSet;
begin
  //#5966 ?W?[?S?O?????X By Kin 2011/03/31
  dtmReport.cdsTempory.FieldDefs.Clear;
  dtmReport.cdsTempory.FieldDefs.Add( 'COMPID', ftString, 6 );
  dtmReport.cdsTempory.FieldDefs.Add( 'COMPNAME', ftString, 50 );
  dtmReport.cdsTempory.FieldDefs.Add( 'YEARMONTH', ftString,6 );
  dtmReport.cdsTempory.FieldDefs.Add( 'YEAR', ftString,4 );
  dtmReport.cdsTempory.FieldDefs.Add( 'MONTH', ftString,2 );
  dtmReport.cdsTempory.FieldDefs.Add( 'STARTMONTH', ftString,2);
  dtmReport.cdsTempory.FieldDefs.Add( 'ENDMONTH', ftString,2);
  dtmReport.cdsTempory.FieldDefs.Add( 'EXECUTEDATE',ftDateTime);
  dtmReport.cdsTempory.FieldDefs.Add( 'UPTTIME', ftDateTime);
  dtmReport.cdsTempory.FieldDefs.Add( 'UPTEN',ftString, 20 );
  dtmReport.cdsTempory.FieldDefs.Add( 'SPECIALPRIZE1', ftString,8);
  dtmReport.cdsTempory.FieldDefs.Add( 'SPECIALPRIZE2', ftString,8);
  dtmReport.cdsTempory.FieldDefs.Add( 'SPECIALPRIZE3', ftString,8);
  dtmReport.cdsTempory.FieldDefs.Add( 'SPECIALPRIZE4', ftString,8);
  dtmReport.cdsTempory.FieldDefs.Add( 'SPECIALPRIZE5', ftString,8);
  dtmReport.cdsTempory.FieldDefs.Add( 'FIRSTPRIZE1', ftString,8);
  dtmReport.cdsTempory.FieldDefs.Add( 'FIRSTPRIZE2', ftString,8);
  dtmReport.cdsTempory.FieldDefs.Add( 'FIRSTPRIZE3', ftString,8);
  dtmReport.cdsTempory.FieldDefs.Add( 'FIRSTPRIZE4', ftString,8);
  dtmReport.cdsTempory.FieldDefs.Add( 'FIRSTPRIZE5', ftString,8);
  dtmReport.cdsTempory.FieldDefs.Add( 'EXTRAPRIZE1', ftString,8);
  dtmReport.cdsTempory.FieldDefs.Add( 'EXTRAPRIZE2', ftString,8);
  dtmReport.cdsTempory.FieldDefs.Add( 'EXTRAPRIZE3', ftString,8);
  dtmReport.cdsTempory.FieldDefs.Add( 'EXTRAPRIZE4', ftString,8);
  dtmReport.cdsTempory.FieldDefs.Add( 'EXTRAPRIZE5', ftString,8);
  dtmReport.cdsTempory.FieldDefs.Add( 'UNUSUALPRIZE1',ftString,8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'UNUSUALPRIZE2',ftString,8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'UNUSUALPRIZE3',ftString,8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'UNUSUALPRIZE4',ftString,8 );
  dtmReport.cdsTempory.FieldDefs.Add( 'UNUSUALPRIZE5',ftString,8 );
  dtmReport.cdsTempory.CreateDataSet;
end;

procedure TfrmInvB06_1.btnPrintClick(Sender: TObject);
var
  aForm : TfrmInvB06_2;
begin
  try
    aForm := TfrmInvB06_2.Create( Application );
    aForm.YearMonth := mskInvUseYear.Text + cmbInvUseMonth.Text;
    if rgpRunCompWhere.ItemIndex = 0 then
      aForm.CompWhere := 0
    else
      aForm.CompWhere := 1;
    aForm.ShowModal;
  finally
    aForm.Free;
  end;

end;

function TfrmInvB06_1.CheckPrize: Boolean;
var
  aIndex:Integer;
  aSQL,aSid,aPassword,aUserId : String;
  aCompId,aCompName,aEndMonth : String;
  aConnSeq:Integer;
  aRet : Boolean;
  aObj : PSMSDb;
begin
  Result := True;
  aRet := True;
  aSQL :=Format('SELECT * FROM INV035 WHERE YEARMONTH = ''%s''  ',
        [mskInvUseYear.Text + cmbInvUseMonth.Text]);
  try
    if not FUseCom then
    begin
      if dtmMain.PRIZEDbList.Count > 0 then
      begin
        UnPrepareDataSet;
        PrepareINV035DataSet;

        for aIndex := 0 to dtmMain.PRIZEDbList.Count - 1 do
        begin
          aConnSeq := PPRIZEDb( dtmMain.PRIZEDbList[aIndex] ).ConnSeq;
          aSid := PPRIZEDb( dtmMain.PRIZEDbList[aIndex] ).DBSid;
          aPassword := PPRIZEDb( dtmMain.PRIZEDbList[aIndex] ).Password;
          aUserId := PPRIZEDb( dtmMain.PRIZEDbList[aIndex] ).UserId;
          aCompId := IntToStr( aConnSeq );
          aCompName := PPRIZEDb( dtmMain.PRIZEDbList[aIndex] ).CompTitle;
          if aCompName = EmptyStr then
          begin
            aObj := dtmMain.GetSoInfo(aConnSeq);
            if Assigned(aObj) then
            begin
              aCompName := aObj.Description;
              //aCompName := dtmMain.GetSoInfo(aConnSeq).Description;
              aObj := nil;

            end;
          end;
          aEndMonth := IntToStr(StrToInt( cmbInvUseMonth.Text ) + 1);
          dtmMain.ConnectToEINVComm( aSid,aUserId,aPassword );
          adoCheck.Close;
          adoCheck.SQL.Text := aSQL;
          adoCheck.Open;
          dtmReport.cdsTempory.Append;
          dtmReport.cdsTempory.FieldByName( 'COMPID' ).AsString := aCompId;
          dtmReport.cdsTempory.FieldByName( 'COMPNAME' ).AsString := aCompName;
          dtmReport.cdsTempory.FieldByName( 'YEARMONTH' ).AsString :=
            mskInvUseYear.Text + cmbInvUseMonth.Text;
          dtmReport.cdsTempory.FieldByName( 'YEAR' ).AsString := mskInvUseYear.Text;
          dtmReport.cdsTempory.FieldByName( 'MONTH' ).AsString := cmbInvUseMonth.Text;
          dtmReport.cdsTempory.FieldByName( 'STARTMONTH' ).AsString := cmbInvUseMonth.Text;
          dtmReport.cdsTempory.FieldByName( 'ENDMONTH' ).AsString := aEndMonth;
          if not adoCheck.IsEmpty then
          begin
            dtmReport.cdsTempory.FieldByName( 'EXECUTEDATE' ).AsString :=
              adoCheck.FieldByName( 'EXECUTEDATE' ).AsString;
            dtmReport.cdsTempory.FieldByName( 'UPTTIME' ).AsString :=
              adoCheck.FieldByName( 'UPTTIME' ).AsString;
            dtmReport.cdsTempory.FieldByName( 'UPTEN' ).AsString :=
              adoCheck.FieldByName( 'UPTEN' ).AsString;
            dtmReport.cdsTempory.FieldByName( 'SPECIALPRIZE1' ).AsString :=
              adoCheck.FieldByName( 'SPECIALPRIZE1' ).AsString;
            dtmReport.cdsTempory.FieldByName( 'SPECIALPRIZE2' ).AsString :=
              adoCheck.FieldByName( 'SPECIALPRIZE2' ).AsString;
            dtmReport.cdsTempory.FieldByName( 'SPECIALPRIZE3' ).AsString :=
              adoCheck.FieldByName( 'SPECIALPRIZE3' ).AsString;
            dtmReport.cdsTempory.FieldByName( 'SPECIALPRIZE4' ).AsString :=
              adoCheck.FieldByName( 'SPECIALPRIZE4' ).AsString;
            dtmReport.cdsTempory.FieldByName( 'SPECIALPRIZE5' ).AsString :=
              adoCheck.FieldByName( 'SPECIALPRIZE5' ).AsString;
            dtmReport.cdsTempory.FieldByName( 'FIRSTPRIZE1' ).AsString :=
              adoCheck.FieldByName( 'FIRSTPRIZE1' ).AsString;
            dtmReport.cdsTempory.FieldByName( 'FIRSTPRIZE2' ).AsString :=
              adoCheck.FieldByName( 'FIRSTPRIZE2' ).AsString;
            dtmReport.cdsTempory.FieldByName( 'FIRSTPRIZE3' ).AsString :=
              adoCheck.FieldByName( 'FIRSTPRIZE3' ).AsString;
            dtmReport.cdsTempory.FieldByName( 'FIRSTPRIZE4' ).AsString :=
              adoCheck.FieldByName( 'FIRSTPRIZE4' ).AsString;
            dtmReport.cdsTempory.FieldByName( 'FIRSTPRIZE5' ).AsString :=
              adoCheck.FieldByName( 'FIRSTPRIZE5' ).AsString;
            dtmReport.cdsTempory.FieldByName( 'EXTRAPRIZE1' ).AsString :=
              adoCheck.FieldByName( 'EXTRAPRIZE1' ).AsString;
            dtmReport.cdsTempory.FieldByName( 'EXTRAPRIZE2' ).AsString :=
              adoCheck.FieldByName( 'EXTRAPRIZE2' ).AsString;
            dtmReport.cdsTempory.FieldByName( 'EXTRAPRIZE3' ).AsString :=
              adoCheck.FieldByName( 'EXTRAPRIZE3' ).AsString;
            dtmReport.cdsTempory.FieldByName( 'EXTRAPRIZE4' ).AsString :=
              adoCheck.FieldByName( 'EXTRAPRIZE4' ).AsString;
            dtmReport.cdsTempory.FieldByName( 'EXTRAPRIZE5' ).AsString :=
              adoCheck.FieldByName( 'EXTRAPRIZE5' ).AsString;
            dtmReport.cdsTempory.FieldByName( 'UNUSUALPRIZE1' ).AsString :=
              adoCheck.FieldByName( 'UNUSUALPRIZE1' ).AsString;
            dtmReport.cdsTempory.FieldByName( 'UNUSUALPRIZE2' ).AsString :=
              adoCheck.FieldByName( 'UNUSUALPRIZE2' ).AsString;
            dtmReport.cdsTempory.FieldByName( 'UNUSUALPRIZE3' ).AsString :=
              adoCheck.FieldByName( 'UNUSUALPRIZE3' ).AsString;
            dtmReport.cdsTempory.FieldByName( 'UNUSUALPRIZE4' ).AsString :=
              adoCheck.FieldByName( 'UNUSUALPRIZE4' ).AsString;
            dtmReport.cdsTempory.FieldByName( 'UNUSUALPRIZE5' ).AsString :=
              adoCheck.FieldByName( 'UNUSUALPRIZE5' ).AsString;
            if aRet then
            begin
              aRet := ChkPrizeNum( adoCheck.FieldByName( 'SPECIALPRIZE1' ).AsString,0);
              if aRet then
                aRet := ChkPrizeNum( adoCheck.FieldByName( 'SPECIALPRIZE2' ).AsString,0);
              if aRet then
                aRet := ChkPrizeNum( adoCheck.FieldByName( 'SPECIALPRIZE3' ).AsString,0);
              if aRet then
                aRet := ChkPrizeNum( adoCheck.FieldByName( 'SPECIALPRIZE4' ).AsString,0);
              if aRet then
                aRet := ChkPrizeNum( adoCheck.FieldByName( 'SPECIALPRIZE5' ).AsString,0);
              if aRet then
              begin
                aRet := ChkPrizeNum( adoCheck.FieldByName( 'FIRSTPRIZE1' ).AsString,1);
                if aRet then
                  aRet := ChkPrizeNum( adoCheck.FieldByName( 'FIRSTPRIZE2' ).AsString,1);
                if aRet then
                  aRet := ChkPrizeNum( adoCheck.FieldByName( 'FIRSTPRIZE3' ).AsString,1);
                if aRet then
                  aRet := ChkPrizeNum( adoCheck.FieldByName( 'FIRSTPRIZE4' ).AsString,1);
                if aRet then
                  aRet := ChkPrizeNum( adoCheck.FieldByName( 'FIRSTPRIZE5' ).AsString,1);
              end;
              if aRet then
              begin
                aRet := ChkPrizeNum( adoCheck.FieldByName( 'EXTRAPRIZE1' ).AsString,2);
                if aRet then
                  aRet := ChkPrizeNum( adoCheck.FieldByName( 'EXTRAPRIZE2' ).AsString,2);
                if aRet then
                  aRet := ChkPrizeNum( adoCheck.FieldByName( 'EXTRAPRIZE3' ).AsString,2);
                if aRet then
                  aRet := ChkPrizeNum( adoCheck.FieldByName( 'EXTRAPRIZE4' ).AsString,2);
                if aRet then
                  aRet := ChkPrizeNum( adoCheck.FieldByName( 'EXTRAPRIZE5' ).AsString,2);
              end;
              //#5966 ?W?[?S?O?????X By Kin 2011/03/31
              if aRet then
              begin
                aRet := ChkPrizeNum( adoCheck.FieldByName( 'UNUSUALPRIZE1' ).AsString,3);
                if aRet then
                  aRet := ChkPrizeNum( adoCheck.FieldByName( 'UNUSUALPRIZE2' ).AsString,3);
                if aRet then
                  aRet := ChkPrizeNum( adoCheck.FieldByName( 'UNUSUALPRIZE3' ).AsString,3);
                if aRet then
                  aRet := ChkPrizeNum( adoCheck.FieldByName( 'UNUSUALPRIZE4' ).AsString,3);
                if aRet then
                  aRet := ChkPrizeNum( adoCheck.FieldByName( 'UNUSUALPRIZE5' ).AsString,3);
              end;
            end;
          end else
          begin
            aRet := False;
          end;
          dtmReport.cdsTempory.Post;
        end;
      end;

    end;
    Result := aRet;
  finally
    dtmMain.EInvConnectComm.Connected := False;
    dtmMain.EInvConnectComm.Close;
  end;



end;

procedure TfrmInvB06_1.PrintErrPrize;
var
  aPath : String;
begin
  aPath := IncludeTrailingPathDelimiter( ExtractFilePath(
      Application.ExeName ) ) + IncludeTrailingPathDelimiter(
      REPORT_FOLDER ) + 'FrptInvB06_1.fr3';
  dtmReport.frxMainReport.LoadFromFile( aPath );
  dtmReport.frxMainReport.Variables.Variables['???????X?????Y'] :=
    QuotedStr( '?U?a???????X??');
  dtmReport.frxMainReport.ShowReport;
end;

function TfrmInvB06_1.ChkPrizeNum(aNum: String;aType : Integer): Boolean;
begin
  Result := False;
  if aType = 0 then
  begin
    if aNum = FChkInvLst.Strings[0] then
    begin
      Result := True;
      Exit;
    end;
    if aNum = FChkInvLst.Strings[1] then
    begin
      Result := True;
      Exit;
    end;
    if aNum = FChkInvLst.Strings[2] then
    begin
      Result := True;
      Exit;
    end;
    if aNum = FChkInvLst.Strings[3] then
    begin
      Result := True;
      Exit;
    end;
    if aNum = FChkInvLst.Strings[4] then
    begin
      Result := True;
      Exit;
    end;

  end;
  if aType = 1 then
  begin
    if aNum = FChkInvLst.Strings[5] then
    begin
      Result := True;
      Exit;
    end;
    if aNum = FChkInvLst.Strings[6] then
    begin
      Result := True;
      Exit;
    end;
    if aNum = FChkInvLst.Strings[7] then
    begin
      Result := True;
      Exit;
    end;
    if aNum = FChkInvLst.Strings[8] then
    begin
      Result := True;
      Exit;
    end;
    if aNum = FChkInvLst.Strings[9] then
    begin
      Result := True;
      Exit;
    end;
  end;
  if aType = 2 then
  begin
    if aNum = StringReplace(FChkInvLst.Strings[10],'X','',[rfReplaceAll]) then
    begin
      Result := True;
      Exit;
    end;
    if aNum = StringReplace(FChkInvLst.Strings[11],'X','',[rfReplaceAll]) then
    begin
      Result := True;
      Exit;
    end;
    if aNum = StringReplace(FChkInvLst.Strings[12],'X','',[rfReplaceAll]) then
    begin
      Result := True;
      Exit;
    end;
    if aNum = StringReplace(FChkInvLst.Strings[13],'X','',[rfReplaceAll]) then
    begin
      Result := True;
      Exit;
    end;
    if aNum = StringReplace(FChkInvLst.Strings[14],'X','',[rfReplaceAll]) then
    begin
      Result := True;
      Exit;
    end;
  end;

  if aType = 3 then
  begin
    if aNum = StringReplace(FChkInvLst.Strings[15],'X','',[rfReplaceAll]) then
    begin
      Result := True;
      Exit;
    end;
    if aNum = StringReplace(FChkInvLst.Strings[16],'X','',[rfReplaceAll]) then
    begin
      Result := True;
      Exit;
    end;
    if aNum = StringReplace(FChkInvLst.Strings[17],'X','',[rfReplaceAll]) then
    begin
      Result := True;
      Exit;
    end;
    if aNum = StringReplace(FChkInvLst.Strings[18],'X','',[rfReplaceAll]) then
    begin
      Result := True;
      Exit;
    end;
    if aNum = StringReplace(FChkInvLst.Strings[19],'X','',[rfReplaceAll]) then
    begin
      Result := True;
      Exit;
    end;
  end;


end;

procedure TfrmInvB06_1.cmbInvUseMonthPropertiesEditValueChanged(
  Sender: TObject);
var
  aMonth : String;
begin
  aMonth := cmbInvUseMonth.Text;
  if aMonth = '' then Exit;
  if StrToInt( aMonth ) >= 2 then
  begin
    if ( StrToInt( aMonth ) mod 2 ) = 0 then
    begin
      aMonth := IntToStr( StrToInt( aMonth ) - 1 );
      if Length( aMonth ) = 1 then
        aMonth := '0' + aMonth;
    end;
  end;
  cmbInvUseMonth.Text := aMonth;
end;

procedure TfrmInvB06_1.btnVerifyClick(Sender: TObject);
begin
  if adoINV035.IsEmpty then
  begin
    MessageDlg('?????d??!',mtInformation,[mbOk],0);
    Exit;
  end;
  if ( adoINV035.FieldByName( 'VerifyEn' ).AsString <> '' ) and
    ( adoINV035.FieldByName( 'VerifyTime' ).AsString <> '' ) then
  begin
    MessageDlg('?????????w?f???L?I',mtInformation,[mbOK],0);
    Exit;
  end;
  if not ValidateCanResult('') then
  begin
    MessageDlg( '?o???????????|?????I?L?k?????????C',mtInformation,[mbOK],0);
    Exit;
  end;

  {#5945 ?????d?U?a???T?????o?????? By Kin 2011/03/15}
  if not CheckPrize then
  begin
    MessageDlg( '?P?U?a???????~?I?L?k?????????C',mtInformation,[mbOK],0);
    PrintErrPrize;
    Exit;
  end;
  adoINV035.Edit;
  adoINV035.FieldByName( 'VERIFYEN' ).AsString := dtmMain.getLoginUserName;
  adoINV035.FieldByName( 'VERIFYTIME' ).AsDateTime := now;
  adoINV035.Post;
  MessageDlg('?f???????I',mtInformation,[mbOk],0);
end;

procedure TfrmInvB06_1.TranPrizeTxtLst(aInv007Where : string;aUseAllComp : Boolean;
  var aUseGiveCnt,aNoUseGiveCnt : Integer);
var sL_SQL,aPath,aSQL,aTitle,aFileName  : String;
    L_TxtList1: TStringList;
    aYear:Word;
    aSQLWhere : String;
begin
  {???X???????|}
  try
    dtmMainJ.getPrintInvoiceTempData(EmptyStr,EmptyStr,EmptyStr,
                        EmptyStr,'2','1',EmptyStr,aInv007Where,
                        EmptyStr,aUseAllComp,False,False);
    L_TxtList1 := TStringList.Create;
//    DateTimeToStr()
//    FormatDateTime()
    //#5966 ????csv???? By Kin 2011/04/08
    aFileName := '?????C?L?M??_'+ mskInvUseYear.Text + cmbInvUseMonth.Text +
                '_' + FormatDateTime('yyyymmdd', Now) + '_' +
                FormatDateTime('hhmmss',Now)+'.CSV';
    sL_SQL := 'SELECT INVPARAM FROM INV003 WHERE COMPID = ' +
        QuotedStr( dtmMain.getCompID );
    aPath := dtmMainJ.GetXMLAttribute( sL_SQL, 'MediaFilePath' );
    aPath := IncludeTrailingPathDelimiter( aPath );



    aSQLWhere := Format(
      ' SELECT Distinct A.INVID                                       ' +
      '   FROM INV023 A, INV008 B, INV002 C                           ' +
      '  WHERE A.IDENTIFYID1 = B.IDENTIFYID1                          ' +
      '    AND A.IDENTIFYID2 = B.IDENTIFYID2                          ' +
      '    AND A.INVID = B.INVID                                      ' +
      '    AND A.IDENTIFYID1 = C.IDENTIFYID1(+)                       ' +
      '    AND A.IDENTIFYID2 = C.IDENTIFYID2(+)                       ' +
      '    AND A.COMPID = C.COMPID(+)                                 ' +
      '    AND A.CUSTID = C.CUSTID(+)                                 ' +
      '    AND A.IDENTIFYID1 = ''%s''                                 ' +
      '    AND A.IDENTIFYID2 = ''%s''                                 ' +
      '    AND A.OWNER = ''%s''                                       ',
      [IDENTIFYID1, IDENTIFYID2, dtmMain.getLoginUserName ] );

    if ( not aUseAllComp) then
    begin
      aSQLWhere := Format( aSQLWhere + ' AND A.COMPID = ''%s'' ',[dtmMain.getCompID]);
    end;



    aSQL := ' SELECT INVID,Nvl(GIVEUNITID,9999) GIVEUNITID,GIVEUNITDESC,                     ' +
            ' 0 USEGIVE,INV007.PRIZETYPE,INV007.BUSINESSID,                                  ' +
            ' INV007.ISOBSOLETE,INV007.TAXTYPE,INV007.INVAMOUNT,INV007.COMPID                ' +
            ' FROM INV007 ,INV028                                                            ' +
            ' WHERE   INV007.PRIZETYPE IS NOT NULL ' + aInv007Where                            +
            ' AND INV007.COMPID = INV028.COMPID                                              ' +
            ' AND INVUSEID = ITEMID AND INV028.REFNO = 1                                     ' +
            ' UNION ALL                                                                      ' +
            ' SELECT INVID,9999 GIVEUNITID,NULL GIVEUNITDESC,1                               ' +
            ' USEGIVE,INV007.PRIZETYPE,INV007.BUSINESSID,                                    ' +
            ' INV007.ISOBSOLETE,INV007.TAXTYPE,INV007.INVAMOUNT,INV007.COMPID                ' +
            ' FROM INV007,INV028                                                             ' +
            ' WHERE PRIZETYPE IS NOT NULL ' + aInv007Where                                     +
            ' AND NVL(INVUSEID,0) = ITEMID(+) AND NVL(INV028.REFNO,0) <> 1                   ' +
            ' AND INV007.TAXTYPE <> 2 ';                                                      



    aSQL := 'SELECT DISTINCT * FROM ( ' + aSQL + ' )                                ' +
            ' WHERE INVID IN ( ' + aSQLWhere + ') ORDER BY USEGIVE,GIVEUNITID,INVID ';
    dtmMainH.adoComm.Close;
    dtmMainH.adoComm.SQL.Text := aSQL;
    dtmMainH.adoComm.Open;
    if dtmMainH.adoComm.RecordCount > 0 then dtmMainH.adoComm.First;
    aTitle := EmptyStr;
    while not dtmMainH.adoComm.Eof do
    begin
      {#6390 ?@?o?B?????~?H?B???s?|?v?B?w?????????A?????i???????????C?L?M?? By Kin 2012/12/14}
      if  ( dtmMainH.adoComm.FieldByName( 'BUSINESSID' ).AsString <> '' ) or
          ( dtmMainH.adoComm.FieldByName( 'ISOBSOLETE' ).AsString = 'Y' ) or
          ( dtmMainH.adoComm.FieldByName( 'TaxType' ).AsString = '?s?|?v' ) or
          ( dtmMainH.adoComm.FieldByName( 'TaxType' ).AsString = '2' ) or
          ( dtmMainH.adoComm.FieldByName( 'InvAmount' ).AsFloat =
             dtmMainH.getInv014Amt(dtmMainH.adoComm.FieldByName( 'INVID' ).AsString,
                                  dtmMainH.adoComm.FieldByName( 'COMPID' ).AsString )) then
      begin

      end else
      begin
        if dtmMainH.adoComm.FieldByName('USEGIVE').AsInteger = 0 then
        begin
            if ( aTitle = EmptyStr ) and ( dtmMain.adoComm.Bof ) then
            begin
              aTitle := '[????]';
              L_TxtList1.Add(aTitle);
              aTitle := EmptyStr;
            end;
            // #5966 ?O???????????A?C?L?????n?i?? By Kin 2011/03/31
            Inc(aUseGiveCnt);
            //#5966 ?W?[?????O By Kin 2011/04/08
            L_TxtList1.Add(dtmMainH.adoComm.FieldByName('INVID').AsString + ',' +
                        dtmMainH.adoComm.FieldByName( 'PRIZETYPE' ).AsString + ',' +
                        dtmMainH.adoComm.FieldByName('GIVEUNITDESC').AsString);

        end else
        begin
          if aTitle = EmptyStr then
          begin
            aTitle := '[?D????]';
            L_TxtList1.Add(aTitle);
          end;
          // #5966 ?O???????????A?C?L?????n?i?? By Kin 2011/03/31
          Inc(aNoUseGiveCnt);
          //#5966 ?W?[?????O By Kin 2011/04/08
          L_TxtList1.Add(dtmMainH.adoComm.FieldByName('INVID').AsString + ',' +
                          dtmMainH.adoComm.FieldByName( 'PRIZETYPE' ).AsString);
        end;
      end;
      dtmMainH.adoComm.Next;
    end;
    if L_TxtList1.Count > 0 then
    begin
      L_TxtList1.SaveToFile(aPath + aFileName);
    end;
  finally
    L_TxtList1.Free;
  end;


end;

procedure TfrmInvB06_1.cxButton1Click(Sender: TObject);
var
  aForm : TfrmInvB06_3;
begin
   try
    aForm := TfrmInvB06_3.Create( Application );
    aForm.ShowModal;
  finally
    aForm.Free;
  end;
end;

end.
