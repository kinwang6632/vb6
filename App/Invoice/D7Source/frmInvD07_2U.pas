unit frmInvD07_2U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, cxStyles, dxSkinsCore, dxSkinBlack, dxSkinBlue,
  dxSkinCaramel, dxSkinCoffee, dxSkinsDefaultPainters, dxSkinscxPCPainter,
  cxCustomData, cxGraphics, cxFilter, cxData, cxDataStorage, cxEdit, DB,
  cxDBData, cxImageComboBox, cxCurrencyEdit, cxGridLevel,
  cxGridCustomTableView, cxGridTableView, cxGridDBTableView, cxClasses,
  cxControls, cxGridCustomView, cxGrid,ADODB,DBClient, cxTextEdit,
  cxCheckBox, cxMemo, Menus, cxLookAndFeelPainters, StdCtrls, cxButtons;

type
  TvarAry = array of String;
  TfrmInvD07_2 = class(TForm)
    pnl1: TPanel;
    dsCSV: TClientDataSet;
    gvCSV: TcxGridDBTableView;
    glCSV: TcxGridLevel;
    grdCSV: TcxGrid;
    dsView: TDataSource;
    gvCSVBussiness: TcxGridDBColumn;
    gvCSVInvSortID: TcxGridDBColumn;
    gvCSVInvType: TcxGridDBColumn;
    gvCSVInvPeriod: TcxGridDBColumn;
    gvCSVInvHead: TcxGridDBColumn;
    gvCSVInvStart: TcxGridDBColumn;
    gvCSVInvEnd: TcxGridDBColumn;
    gvCSVInvUploadFlag: TcxGridDBColumn;
    gvCSVNotes: TcxGridDBColumn;
    gvCSVErrMsg: TcxGridDBColumn;
    btnOK: TcxButton;
    btnExit: TcxButton;
    gvCSVInvIsErr: TcxGridDBColumn;
    procedure FormCreate(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnOKClick(Sender: TObject);
  private
    { Private declarations }
    FAdoQry : TADOQuery;
    FHaveErrs : Boolean;
    procedure PrepareDataSet;
    procedure FillData;
    function Split(const aSplitStr: String;const aSeparatedWord: String) :TvarAry;overload;
  public
    { Public declarations }
    constructor Create(aDataSet: TADOQuery); reintroduce;
  end;


var
  frmInvD07_2: TfrmInvD07_2;

implementation
uses cbUtilis,  Ustru,dtmMainJU,dtmMainU;
{$R *.dfm}

{ TfrmInvD07_2 }

constructor TfrmInvD07_2.Create(aDataSet: TADOQuery);
begin
  inherited Create( Application );
  FAdoQry := aDataSet;
  FHaveErrs := False;
end;

procedure TfrmInvD07_2.PrepareDataSet;
begin
   if ( dsCSV.FieldDefs.Count <= 0 ) then
  begin

    dsCSV.FieldDefs.Add( 'INVISERR', ftInteger );
    dsCSV.FieldDefs.Add( 'BUSSINESSID', ftString, 100 );
    dsCSV.FieldDefs.Add( 'INVSORTID', ftString, 100 );
    dsCSV.FieldDefs.Add( 'INVTYPE', ftString, 100 );
    dsCSV.FieldDefs.Add( 'INVPERIOD', ftString, 100 );
    dsCSV.FieldDefs.Add( 'INVHEAD',ftString,20 );
    dsCSV.FieldDefs.Add( 'INVSTART', ftString,50 );
    dsCSV.FieldDefs.Add( 'INVEND', ftString,50 );
    dsCSV.FieldDefs.Add( 'INVUPLOADFLAG', ftInteger );
    dsCSV.FieldDefs.Add( 'NOTES', ftString, 200 );
    dsCSV.FieldDefs.Add( 'ERRMSG', ftString, 200 );
    dsCSV.CreateDataSet;
  end;
  dsCSV.EmptyDataSet;

end;

procedure TfrmInvD07_2.FormCreate(Sender: TObject);
begin

  dtmMain.SetWaitingCursor;
  try
    PrepareDataSet;
    FillData;
    btnOK.Enabled := not FHaveErrs;
  finally
    dtmMain.SetDefaultCursor;
  end;


end;

procedure TfrmInvD07_2.FillData;
 var
   aAryPeriod: TvarAry;
   aMod : Integer;

begin
  FAdoQry.First;
  try
    while not FAdoQry.eof do
    begin
      try
         dsCSV.Append;
        dsCSV.FieldByName('INVISERR').AsInteger := 0;
        if FAdoQry.Fields[0].AsString <> EmptyStr then
          dsCSV.FieldByName('BUSSINESSID').AsString := FAdoQry.Fields[0].AsString;
        if FAdoQry.Fields[1].AsString <> EmptyStr then
          dsCSV.FieldByName('INVSORTID').AsString := FAdoQry.Fields[1].AsString;
        if FAdoQry.Fields[2].AsString <> EmptyStr then
          dsCSV.FieldByName('INVTYPE').AsString := FAdoQry.Fields[2].AsString;
        if FAdoQry.Fields[3].AsString <> EmptyStr then
          dsCSV.FieldByName('INVPERIOD').AsString := FAdoQry.Fields[3].AsString;
        if FAdoQry.Fields[4].AsString <> EmptyStr then
          dsCSV.FieldByName('INVHEAD').AsString := FAdoQry.Fields[4].AsString;
        if FAdoQry.Fields[5].AsString <> EmptyStr then
          dsCSV.FieldByName('INVSTART').AsString := FAdoQry.Fields[5].AsString;
        if FAdoQry.Fields[6].AsString <> EmptyStr then
          dsCSV.FieldByName('INVEND').AsString := FAdoQry.Fields[6].AsString;
        if FAdoQry.Fields[7].AsString <> EmptyStr then
          dsCSV.FieldByName('INVUPLOADFLAG').AsString := FAdoQry.Fields[7].AsString;
        if FAdoQry.Fields[8].AsString <> EmptyStr then
          dsCSV.FieldByName('NOTES').AsString := FAdoQry.Fields[8].AsString;
        dsCSV.Post;
        dsCSV.Edit;
        if dsCSV.FieldByName('INVPERIOD').AsString = '' then
        begin
          dsCSV.FieldByName('ERRMSG').AsString := '發票期別無值！';
          dsCSV.FieldByName('INVISERR').AsInteger := 1;
          FHaveErrs := True;
          FAdoQry.Next;
          Continue;
        end;
        aAryPeriod := Split(dsCSV.FieldByName('INVPERIOD').AsString,'~');
        if high(aAryPeriod) <> 1 then
        begin
          dsCSV.FieldByName('ERRMSG').AsString := '發票期別格式不正確！';
          dsCSV.FieldByName('INVISERR').AsInteger := 1;
          FHaveErrs := True;
          FAdoQry.Next;
          Continue;
        end;
        if ( StrToFloat(StringReplace(aAryPeriod[1],'/','',[rfReplaceAll])) <=
                StrToFloat(StringReplace(aAryPeriod[0],'/','',[rfReplaceAll]))) then
        begin
          dsCSV.FieldByName('ERRMSG').AsString := '發票期別有誤！';
          dsCSV.FieldByName('INVISERR').AsInteger := 1;
          FHaveErrs := True;
          FAdoQry.Next;
          Continue;
        end;

        if dsCSV.FieldByName('INVHEAD').AsString = '' then
        begin
          dsCSV.FieldByName('ERRMSG').AsString := '發票字軌需有值！';
          dsCSV.FieldByName('INVISERR').AsInteger := 1;
          FHaveErrs := True;
          FAdoQry.Next;
          Continue;
        end;


        if dsCSV.FieldByName('INVSTART').AsString ='' then
        begin
          dsCSV.FieldByName('ERRMSG').AsString := '發票開始編號需有值！';
          dsCSV.FieldByName('INVISERR').AsInteger := 1;
          FHaveErrs := True;
          FAdoQry.Next;
          Continue;
        end;


        if dsCSV.FieldByName('INVEND').AsString ='' then
        begin
          dsCSV.FieldByName('ERRMSG').AsString := '發票結束編號需有值！';
          dsCSV.FieldByName('INVISERR').AsInteger := 1;
          FHaveErrs := True;
          FAdoQry.Next;
          Continue;
        end;
        aMod := ( dsCSV.FieldByName('INVEND').AsInteger - dsCSV.FieldByName('INVSTART').AsInteger ) +1 ;
        aMod := aMod mod 50;
        if (aMod <> 0 ) then
        begin
          dsCSV.FieldByName('ERRMSG').AsString := '發票起迄總張數必須為50的倍數！';
          dsCSV.FieldByName('INVISERR').AsInteger := 1;
          FHaveErrs := True;
          FAdoQry.Next;
          Continue;
        end;


        if dsCSV.FieldByName('INVUPLOADFLAG').AsString ='' then
        begin
          dsCSV.FieldByName('ERRMSG').AsString := '上傳註記必須有值！';
          dsCSV.FieldByName('INVISERR').AsInteger := 1;
          FHaveErrs := True;
          FAdoQry.Next;
          Continue;
        end;
         FAdoQry.Next;
      except
        on E: Exception do
        begin
          dsCSV.FieldByName('ERRMSG').AsString := E.Message;
          FHaveErrs := True;
          FAdoQry.Next;
        end;
      end;

    end;
  except

    on E: Exception do
    begin
      ErrorMsg( Format( '訊息:%s', [E.Message] ) );
      Exit;
    end;
  end;


end;

function TfrmInvD07_2.Split(const aSplitStr,
  aSeparatedWord: String): TvarAry;
  var
  aStr: string;
  aSource: String;
  aIndex : Integer;
  aArry : TvarAry;
begin
  if Length(aSplitStr) = 0 then
    begin
      Result := nil;
    end else
    begin
      aStr := EmptyStr;
      aSource := aSplitStr;
      if Pos(aSeparatedWord,aSource) > 0 then
      begin
        for aIndex := 1 to Length( aSource ) do
        begin
          if ( aSource[aIndex] <> aSeparatedWord ) then
          begin
            aStr := ( aStr + aSource[aIndex] );
          end
          else begin
            SetLength( aArry, Length( aArry ) + 1 );
            aArry[ Length( aArry ) - 1 ] := aStr;
            aStr := '';
          end;
        end;
        if ( aStr <> '' ) then
        begin
          SetLength( aArry, Length( aArry ) + 1 );
          aArry[ Length( aArry ) - 1 ] := aStr;
          aStr := '';
        end;
      end else
      begin
        SetLength(aArry,1);
        aArry[ Length( aArry ) - 1 ] := aSource;
      end;
    end;
    Result := aArry;
end;

procedure TfrmInvD07_2.btnExitClick(Sender: TObject);
begin
  Close;
end;

procedure TfrmInvD07_2.btnOKClick(Sender: TObject);
var
    sL_YearMonth,sL_Prefix,sL_StartNum,sL_EndNum,sL_CurNum : String;
    sL_LastInvDate,sL_Memo,sL_InvFormat,sL_Useful,sL_SQL : String;
    bL_HavePKValue : Boolean;
    aCheckCurNum: String;
    sL_UploadFlag,sL_UseKind: Integer;
    aAryPeriod: TvarAry;
    aErrMsg : string;
begin
  dtmMain.SetWaitingCursor;
  try
    dsCSV.First;
    aErrMsg := '';
    while not dsCSV.eof do
    begin
      try
        aAryPeriod := Split(dsCSV.FieldByName('INVPERIOD').AsString,'~');
        aAryPeriod[0] := StringReplace(aAryPeriod[0],' ','',[rfReplaceAll]);
        aAryPeriod[1] := StringReplace(aAryPeriod[1],' ','',[rfReplaceAll]);
        aAryPeriod[0] := IntToStr( StrToInt(copy(aAryPeriod[0],1,3))+ 1911) + '/' + copy(aAryPeriod[0],5,2);
        aAryPeriod[1] := IntToStr( StrToInt(copy(aAryPeriod[1],1,3))+ 1911 ) +'/' + copy(aAryPeriod[1],5,2);
        sL_YearMonth := TUstr.replaceStr(Trim(aAryPeriod[0]),'/','');
        sL_YearMonth := dtmMainJ.changePrefixYM(sL_YearMonth);

        sL_Prefix :=  UpperCase(Trim(dsCSV.FieldByName('INVHEAD').AsString) );
        sL_StartNum := Lpad( Trim(dsCSV.FieldByName('INVSTART').AsString), 8, '0' );
        sL_EndNum := Lpad( Trim(dsCSV.FieldByName('INVEND').AsString), 8, '0' );
        sL_CurNum := Lpad( Trim(dsCSV.FieldByName('INVSTART').AsString), 8, '0' );
        sL_LastInvDate :=  aAryPeriod[0] + '/01';
        sL_Memo := dsCSV.FieldByName('NOTES').AsString;
        sL_InvFormat := '1';
        sL_Useful := 'Y';
        sL_UploadFlag := dsCSV.FieldByName('INVUPLOADFLAG').AsInteger;
        sL_UseKind := 0;
        sL_SQL := 'select * from inv099 where IdentifyId1=''' + IDENTIFYID1 +
                  ''' and IdentifyId2= ''' + IDENTIFYID2 +
                  ''' and YearMonth=''' + sL_YearMonth +
                  ''' and Prefix=''' + sL_Prefix +
                  ''' and StartNum=' + sL_StartNum;
        bL_HavePKValue := dtmMainJ.checkPK(sL_SQL);
        if bL_HavePKValue then
        begin
          if aErrMsg <> '' then aErrMsg := aErrMsg + #13#10;
          aErrMsg := aErrMsg + Format('違反唯一值條件 : %s',
               [UpperCase(Trim(dsCSV.FieldByName('INVHEAD').AsString)) +
                UpperCase(Trim(dsCSV.FieldByName('INVSTART').AsString))]);

          dsCSV.Next;
          Continue;
        end;
        dtmMain.adoComm.Close;

        dtmMain.adoComm.SQL.Text := Format(
            ' INSERT INTO INV099 ( IDENTIFYID1, IDENTIFYID2,                     ' +
            '    COMPID, YEARMONTH, INVFORMAT, PREFIX, STARTNUM,                 ' +
            '    ENDNUM, CURNUM,  LASTINVDATE, USEFUL,                           ' +
            '    MEMO,UPTEN,UPLOADFLAG,USEKIND,UPTTIME )                         ' +
            '  VALUES ( ''%s'', ''%s'',                                          ' +
            '          ''%s'', ''%s'', ''%s'', ''%s'',''%s'',                    ' +
            '           ''%s'', ''%s'',to_date( ''%s'', ''YYYY/MM/DD'' ),        ' +
            '           ''%s'',                                                  ' +
            '           ''%s'', ''%s'', %d, %d, SYSDATE )                        ' ,
            [IDENTIFYID1, IDENTIFYID2,
             dtmMain.getCompID, sL_YearMonth,sL_InvFormat,sL_Prefix,sL_StartNum,
             sL_EndNum,sL_CurNum, sL_LastInvDate,
             sL_Useful,
             sL_Memo,dtmMain.getLoginUser,sL_UploadFlag,sL_UseKind ] );
        dtmMain.adoComm.ExecSQL;
        dsCSV.Next;
      except
        on E: Exception do
        begin
          if aErrMsg <> '' then aErrMsg := aErrMsg + #13#10;
          aErrMsg := aErrMsg + Format('%s : %s',
               [UpperCase(Trim(dsCSV.FieldByName('INVHEAD').AsString)) +
                UpperCase(Trim(dsCSV.FieldByName('INVSTART').AsString)),
                E.Message]);

          dsCSV.Next;
        end;
      end;

    end;

    if aErrMsg = '' then
    begin
      MessageDlg('匯入作業完成！',mtInformation,[mbOK],0);
    end else
    begin
      MessageDlg(Format('匯入作業完成！'+#13#10 + '%s',[aErrMsg]),
              mtInformation,[mbOK],0);
    end;
    dtmMain.SetDefaultCursor;
    except
      on E: Exception do
      begin
        dtmMain.SetDefaultCursor;
        ErrorMsg( Format( '訊息:%s', [E.Message] ) );
        Exit;
      end;
   end;

end;

end.
