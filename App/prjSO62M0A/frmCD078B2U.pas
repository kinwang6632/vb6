unit frmCD078B2U;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ActnList, ExtCtrls, DB, DBClient, ADODB, Menus,
  frmCD078B0U, frmCD078B1U, cbDBController, cbStyleController,
  cxControls, cxContainer, cxEdit, cxTextEdit, cxMaskEdit, cxDropDownEdit,
  cxImageComboBox, cxDBLookupComboBox,
  cbDataLookup, cxGraphics, dxSkinsCore, dxSkinsDefaultPainters,
  dxSkinscxPCPainter, cxLookAndFeelPainters, cxButtons, dxSkinBlack,
  dxSkinBlue, dxSkinCaramel, dxSkinCoffee, cxCheckBox;

type
  TfrmCD078B2 = class(TForm)
    Panel1: TPanel;
    Panel2: TPanel;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    edtFaciItem: TcxTextEdit;
    cmbFaciType: TcxComboBox;
    cmbFixIpCount: TcxImageComboBox;
    cmbDynIpCount: TcxImageComboBox;
    edtDML: TcxTextEdit;
    Bevel1: TBevel;
    btnSave: TButton;
    Button1: TButton;
    FaciCode: TDataLookup;
    ModelCode: TDataLookup;
    BuyCode: TDataLookup;
    ActionList1: TActionList;
    actSave: TAction;
    actCancel: TAction;
    ServiceType: TDataLookup;
    CMBaudNo: TDataLookup;
    btnInstCodeStr: TButton;
    cxButton1: TcxButton;
    edtInstCodeStr: TcxTextEdit;
    chkFixIP: TcxCheckBox;
    chkDynIP: TcxCheckBox;
    Label10: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure actSaveExecute(Sender: TObject);
    procedure actCancelExecute(Sender: TObject);
    procedure FaciCodeCodeNamePropertiesInitPopup(Sender: TObject);
    procedure ServiceTypeCodeNamePropertiesInitPopup(Sender: TObject);
    procedure ModelCodeCodeNamePropertiesInitPopup(Sender: TObject);
    procedure BuyCodeCodeNamePropertiesInitPopup(Sender: TObject);
    procedure CMBaudNoCodeNamePropertiesInitPopup(Sender: TObject);
    procedure cmbFaciTypePropertiesChange(Sender: TObject);
    procedure ServiceTypeCodeNoPropertiesChange(Sender: TObject);
    procedure FaciCodeCodeNoPropertiesChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FaciCodeCodeNamePropertiesChange(Sender: TObject);
    procedure ServiceTypeCodeNamePropertiesChange(Sender: TObject);
    procedure btnInstCodeStrClick(Sender: TObject);
    procedure cxButton1Click(Sender: TObject);
  private
    { Private declarations }
    FMode: TDML;
    FKeyValue: String;
    FDataFrom: String;
    FDataSet: TClientDataSet;
    FReader: TADOQuery;
    FMaxFaciItem: String;
    FQueryParam: TQueryParam;
    FInstCodeDataSet: TClientDataSet;
    procedure LoacateItem(const aValue: String; aControl: TcxImageComboBox);
    procedure LoadFaciCodeRelationItem;
    procedure ClearEditor;
    procedure DataToEditor;
    procedure LoadCodeData;
    procedure DMLModeChange(aMode: TDML);
    procedure ButtonStateChange;
    procedure EditorStateChange;
    procedure ChangeIpCountState;
    function DataSetStateChange(const aMode: TDML): Boolean;
    function DataValidate: Boolean;
    function EditorToData: Boolean;
    function VdMustInputIpCount: Boolean;
  public
    { Public declarations }
    constructor Create(aMode: TDML; aKeyValue, aDataFrom: String;
      aDataSet: TClientDataSet); reintroduce;
    property InstCodeDataSet: TClientDataSet read FInstCodeDataSet write FInstCodeDataSet;
  end;

var
  frmCD078B2: TfrmCD078B2;

implementation

uses cbUtilis, frmMultiSelectU;

{$R *.dfm}

{ TfrmCD078B2 }

{ ---------------------------------------------------------------------------- }

constructor TfrmCD078B2.Create(aMode: TDML; aKeyValue, aDataFrom: String;
  aDataSet: TClientDataSet);
begin
  inherited Create( Application );
  FMode := aMode;
  FKeyValue := aKeyValue;
  FDataFrom := aDataFrom;
  FDataSet := aDataSet;
  FReader := DBController.CodeReader;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.FormCreate(Sender: TObject);
begin
  ServiceType.Initializa;
  FaciCode.Initializa;
  ModelCode.Initializa;
  BuyCode.Initializa;
  CMBaudNo.Initializa;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.FormShow(Sender: TObject);
begin
  Screen.Cursor := crSQLWait;
  try
    LoadCodeData;
    DMLModeChange( FMode );
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if ( FMode <> dmCancel ) then
    DMLModeChange( dmCancel );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.FormDestroy(Sender: TObject);
begin
  FReader := nil;
  FDataSet := nil;
  FaciCode.Finaliza;
  ServiceType.Finaliza;
  ModelCode.Finaliza;
  BuyCode.Finaliza;
  CMBaudNo.Finaliza;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.LoacateItem(const aValue: String; aControl: TcxImageComboBox);
var
  aIndex: Integer;
  aItem: TcxImageComboBoxItem;
begin
  for aIndex := 0 to aControl.Properties.Items.Count - 1 do
  begin
    aItem := aControl.Properties.Items[aIndex];
    if ( VarToStrDef( aItem.Value, EmptyStr ) = aValue ) then
    begin
      aControl.ItemIndex := aIndex;
      Break;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.LoadFaciCodeRelationItem;
begin
  FReader.Close;
  Screen.Cursor := crSQLWait;
  try
    FReader.SQL.Text := Format(
      ' SELECT B.FACIMODELNO,             ' +
      '        A.SERVICETYPE,             ' +
      '        A.DEFBUYCODE               ' +
      '   FROM %s.CD022 A,                ' +
      '       ( SELECT DISTINCT           ' +
      '                FACICODE,          ' +
      '                FACIMODELNO        ' +
      '           FROM %s.CD022A          ' +
      '          WHERE FACICODE = ''%s''  ' +
      '            AND ROWNUM <= 1 ) B    ' +
      '  WHERE A.CODENO = B.FACICODE(+)   ' +
      '    AND A.CODENO = ''%s''          ',
      [DBController.LoginInfo.DbAccount, DBController.LoginInfo.DbAccount,
       FaciCode.Value, FaciCode.Value] );
    FReader.Open;
    if ( FReader.FieldByName( 'FACIMODELNO' ).AsString <> EmptyStr ) then
      ModelCode.Value := FReader.FieldByName( 'FACIMODELNO' ).AsString;
    if ( FReader.FieldByName( 'SERVICETYPE' ).AsString <> EmptyStr ) and
       ( FReader.FieldByName( 'SERVICETYPE' ).AsString <> ServiceType.Value ) then
      ServiceType.Value := FReader.FieldByName( 'SERVICETYPE' ).AsString;
    if ( FReader.FieldByName( 'DEFBUYCODE' ).AsString <> EmptyStr ) then
      BuyCode.Value := FReader.FieldByName( 'DEFBUYCODE' ).AsString;
    FReader.Close;
  finally
    Screen.Cursor := crDefault;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.ClearEditor;
begin
  edtFaciItem.Text := EmptyStr;
  FaciCode.Clear;
  ServiceType.Clear;
  ModelCode.Clear;
  BuyCode.Clear;
  CMBaudNo.Clear;
  {}
  cmbFaciType.ItemIndex := 0;
  cmbFixIpCount.ItemIndex := 0;
  cmbDynIpCount.ItemIndex := 0;

  {5859 ?T?wIP?w?]False,?B???w?] True By Kin 2011/04/27}
  chkFixIP.Checked := False;
  chkDynIP.Checked := True;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.DataToEditor;
begin
  edtFaciItem.Text := FDataSet.FieldByName( 'FACIITEM' ).AsString;
  if ( FDataSet.FieldByName( 'FACITYPE' ).AsInteger in
    [0..cmbFaciType.Properties.Items.Count - 1 ] ) then
    cmbFaciType.ItemIndex := FDataSet.FieldByName( 'FACITYPE' ).AsInteger;
  ServiceType.Value := FDataSet.FieldByName( 'SERVICETYPE' ).AsString;
  FaciCode.Value := FDataSet.FieldByName( 'FACICODE' ).AsString;
  ModelCode.Value := FDataSet.FieldByName( 'MODELCODE' ).AsString;
  BuyCode.Value := FDataSet.FieldByName( 'BUYCODE' ).AsString;
  CMBaudNo.Value := FDataSet.FieldByName( 'CMBAUDNO' ).AsString;
  FQueryParam.Param1 := FDataSet.FieldByName( 'INSTCODESTR' ).AsString;
  FQueryParam.Param2 := '3,1,2';
  edtInstCodeStr.Text := DBController.GetInstName( FQueryParam.Param1 );
  LoacateItem( FDataSet.FieldByName( 'FIXIPCOUNT' ).AsString, cmbFixIpCount );
  LoacateItem( FDataSet.FieldByName( 'DYNIPCOUNT' ).AsString, cmbDynIpCount);
  {#5859 ?W?[CBPChangeFixIP?BCBPChangeDynIP By Kin 2011/04/27}
  chkFixIP.Checked := ( FDataSet.FieldByName( 'CBPCHANGEFIXIP' ).AsString = '1' );
  chkDynIP.Checked := ( FDataSet.FieldByName( 'CBPCHANGEDYNIP' ).AsString = '1' );
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B2.EditorToData: Boolean;
begin
  Result := False;
  try
    if ( FMode in [dmInsert] ) then
    begin
      FDataSet.FieldByName( 'BPCODE' ).AsString := FKeyValue;
      FDataSet.FieldByName( 'NEWBPCODE' ).AsString := FDataFrom;
      FDataSet.FieldByName( 'FACIITEM' ).AsString := edtFaciItem.Text;
    end;
    FDataSet.FieldByName( 'FACITYPE' ).AsString := IntToStr( cmbFaciType.ItemIndex );
    FDataSet.FieldByName( 'FACICODE' ).AsString := FaciCode.Value;
    FDataSet.FieldByName( 'FACINAME' ).AsString := FaciCode.ValueName;
    FDataSet.FieldByName( 'REFNO' ).AsString := FaciCode.RefNo;
    FDataSet.FieldByName( 'SERVICETYPE' ).AsString := ServiceType.Value;
    FDataSet.FieldByName( 'MODELCODE' ).AsString :=  ModelCode.Value;
    FDataSet.FieldByName( 'MODELNAME' ).AsString := ModelCode.ValueName;
    FDataSet.FieldByName( 'BUYCODE' ).AsString := BuyCode.Value;
    FDataSet.FieldByName( 'BUYNAME' ).AsString := BuyCode.ValueName;
    FDataSet.FieldByName( 'CMBAUDNO' ).AsString := CMBaudNo.Value;
    FDataSet.FieldByName( 'CMBAUDRATE' ).AsString := CMBaudNo.ValueName;
    FDataSet.FieldByName( 'FIXIPCOUNT' ).AsString := VarToStrDef(
      cmbFixIpCount.Properties.Items[cmbFixIpCount.ItemIndex].Value, '0' );
    FDataSet.FieldByName( 'DYNIPCOUNT' ).AsString := VarToStrDef(
      cmbDynIpCount.Properties.Items[cmbDynIpCount.ItemIndex].Value, '0' );
    FDataSet.FieldByName( 'INSTCODESTR' ).AsString := TrimChar( FQueryParam.Param1, [''''] );

    {#5859 ?W?[CBPCHANGEFIXIP?BCBPCHANGEDYNIP By Kin 2011/04/27}
    FDataSet.FieldByName( 'CBPCHANGEFIXIP' ).AsString := '0';
    FDataSet.FieldByName( 'CBPCHANGEDYNIP' ).AsString := '0';
    if chkFixIP.Checked then
      FDataSet.FieldByName( 'CBPCHANGEFIXIP' ).AsString := '1';
    if chkDynIP.Checked then
      FDataSet.FieldByName( 'CBPCHANGEDYNIP' ).AsString := '1';
  except
    on E: Exception do
    begin
      ErrorMsg( Format( '?s??????,???]:%s', [E.Message] ) );
      Exit;
    end;
  end;
  Result := True;
end;


{ ---------------------------------------------------------------------------- }

function TfrmCD078B2.VdMustInputIpCount: Boolean;
begin
  //#6352 ?W?[??????7?B8 By Kin 2012/11/07
  Result :=
    ( ServiceType.Value = 'I' ) and
    ( ( FaciCode.RefNo = '2' ) or ( FaciCode.RefNo = '5' )
    or ( FaciCode.RefNo = '7' ) or ( FaciCode.RefNo = '8' )  );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.DMLModeChange(aMode: TDML);
begin
  if not DataSetStateChange( aMode ) then Exit;
  if ( aMode = dmSave  ) and ( FMode = dmInsert ) then
  begin
    DMLModeChange( dmInsert );
    Exit;
  end
  else  if ( aMode in [dmSave, dmCancel] ) then
    FMode := dmBrowser;
  ButtonStateChange;
  EditorStateChange;
  edtDML.Text := TDMLString[FMode];
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.ButtonStateChange;
begin
  if ( FMode = dmInsert ) then
  begin
    actSave.Enabled := True;
  end else
  if ( FMode = dmUpdate ) then
  begin
    actSave.Enabled := True;
  end else
  begin
    actSave.Enabled := False;
    actCancel.Caption := '????(&C)';
    actCancel.ShortCut := TextToShortCut( 'Alt+C' );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.EditorStateChange;
begin
  case FMode of
    dmInsert:
      begin
        ClearEditor;
        edtFaciItem.Text := FMaxFaciItem;
        edtFaciItem.Enabled := False;
        if ( cmbFaciType.CanFocusEx ) then cmbFaciType.SetFocus;
        ChangeIpCountState;
      end;
    dmUpdate:
      begin
        edtFaciItem.Enabled := False;
        cmbFaciType.Enabled := True;
        FaciCode.Enabled := True;
        ServiceType.Enabled := True;
        ModelCode.Enabled := True;
        BuyCode.Enabled := True;
        {}
        ChangeIpCountState;
        {}
        if ( cmbFaciType.CanFocusEx ) then cmbFaciType.SetFocus;
      end;
    dmBrowser:
      begin
        edtFaciItem.Enabled := False;
        cmbFaciType.Enabled := False;
        FaciCode.Enabled := False;
        ServiceType.Enabled := False;
        ModelCode.Enabled := False;
        BuyCode.Enabled := False;
        ChangeIpCountState;
        if ( cmbFaciType.CanFocusEx ) then cmbFaciType.SetFocus;
      end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.ChangeIpCountState;
begin
  CMBaudNo.Enabled := False;
  cmbFixIpCount.Enabled := False;
  cmbDynIpCount.Enabled := False;
  chkFixIP.Enabled := False;
  chkDynIP.Enabled := False;
  Label8.Font.Color := clWindowText;
  Label9.Font.Color := clWindowText;
  if ( VdMustInputIpCount ) then
  begin
    CMBaudNo.Enabled := True;
    cmbFixIpCount.Enabled := True;
    cmbDynIpCount.Enabled := True;
    chkFixIP.Enabled := True;
    chkDynIP.Enabled := True;
    Label8.Font.Color := clRed;
    Label9.Font.Color := clRed;
  end;
  if not CMBaudNo.Enabled then CMBaudNo.Clear;
  if not cmbFixIpCount.Enabled then cmbFixIpCount.ItemIndex := 0;
  if not cmbDynIpCount.Enabled then cmbDynIpCount.ItemIndex := 0;


end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B2.DataSetStateChange(const aMode: TDML): Boolean;
begin
  Result := True;
  case aMode of
    dmInsert:
      begin
        FMaxFaciItem := DBController.GetFaciItemSeqNo( FDataSet );
        FDataSet.Append;
      end;
    dmUpdate:
      begin
        ClearEditor;
        DataToEditor;
        FDataSet.Edit;
      end;
    dmCancel:
      begin
        FDataSet.Cancel;
      end;
    dmBrowser:
      begin
        ClearEditor;
        DataToEditor;
      end;
    dmSave:
      begin
        Result := DataValidate;
        if Result then Result := EditorToData;
        if Result then
        begin
          try
            FDataSet.Post;
          except
            on E: Exception do ErrorMsg( Format(
              '?s?????~, ???]:%s?C', [E.Message] ) );
          end;
        end;
      end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.actSaveExecute(Sender: TObject);
begin
  DMLModeChange( dmSave );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.actCancelExecute(Sender: TObject);
begin
  DMLModeChange( dmCancel );
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

function TfrmCD078B2.DataValidate: Boolean;
var
  aErrMsg: String;
  aControl: TWinControl;
begin
  aErrMsg := EmptyStr;
  aControl := nil;
  try
    if ( cmbFaciType.ItemIndex <= 0 ) then
    begin
      aErrMsg := '?????J?]???????C';
      aControl := cmbFaciType;
      Exit;
    end;
    {}
    if ( FaciCode.Value = EmptyStr ) then
    begin
      aErrMsg := '?????J?]???????C';
      aControl := FaciCode.CodeNo;
      Exit;
    end;
    {}
    if ( ServiceType.Value = EmptyStr ) then
    begin
      aErrMsg := '?????J?A?????O?C';
      aControl := ServiceType.CodeNo;
      Exit;
    end;
    {}
    if ( BuyCode.Value = EmptyStr ) then
    begin
      aErrMsg := '?????J?R???????C';
      aControl := BuyCode.CodeNo;
      Exit;
    end;
    {}
    if ( cmbDynIpCount.Enabled ) and ( cmbDynIpCount.ItemIndex <= 0 ) then
    begin
      aErrMsg := '???????B??IP???C';
      aControl := cmbDynIpCount;
      Exit;
    end;
    {}
    if ( cmbFixIpCount.Enabled ) and ( cmbFixIpCount.ItemIndex <= 0 ) then
    begin
      aErrMsg := '???????T?wIP???C';
      aControl := cmbFixIpCount;
      Exit;
    end;
    {}
    if ( FQueryParam.Param1 = EmptyStr ) then
    begin
      aErrMsg := '?????J?i???w???u???O?j?C';
      aControl := btnInstCodeStr;
      Exit;
    end;
    {}
    if not DBController.VdInstCodeExitisInCD078B( FQueryParam.Param1,
      FInstCodeDataSet ) then
    begin
      aErrMsg := '?i???w???u???O?j???????b?i???z?????u?????j???????s?b, ?????s?]?w?C';
      aControl := btnInstCodeStr;
      Exit;
    end;
  finally
    Result := ( aErrMsg = EmptyStr );
    if not Result then
    begin
      WarningMsg( aErrMsg );
      if Assigned( aControl ) then
        if aControl.CanFocus then aControl.SetFocus;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.FaciCodeCodeNamePropertiesInitPopup(Sender: TObject);
var
  aString : string;
begin
  { ?]?????? }
  if ( cmbFaciType.ItemIndex = 0 ) then
  begin
    FaciCode.SQL.Text := Format(
      ' SELECT A.CODENO,       ' +
      '        A.DESCRIPTION,  ' +
      '        A.REFNO         ' +
      '   FROM %s.CD022 A      ' +
      '  WHERE 1 = 2           ',
      [DBController.LoginInfo.DbAccount] );
    FaciCode.CodeNamePropertiesInitPopup( Sender );
  end  else
  begin
    { ?A???O??C(CATV)?h???????????X?? 0,1,3,4
      ?A???O??D(DTV)?h???????????X?? 0,3,4
      ?A???O??I(CM)?h???????????X?? 0,2,5
      ?A???O??P(CP)?h???????????X?? 6 }
    if ServiceType.Value = 'C' then
      aString := ' AND A.REFNO IN (0, 1, 3 ,4, 10) ' //#6253 ?W?[??????10 By Kin 2012/04/18
    else if ServiceType.Value ='D' then
      {?W?[??????11 For Jacky By Kin 2013/07/11}
      aString := ' AND A.REFNO IN (0, 3 ,4 ,9, 11) '     //#5139 ??????OK,?n?W?[??????9 By Kin 2009/07/02
    else if ServiceType.Value ='I' then
      aString := ' AND A.REFNO IN (0, 2 ,5, 7, 8) '  // #4207 ?W?[??????7?B8 By Kin 2009/02/03
    else if ServiceType.Value ='P' then
      aString := ' AND A.REFNO IN (6) '
    else
      aString := ' AND A.REFNO IN ('''') ';
    {}
    //#5926 SERVICETYPE = NULL ?]?nSHOW?X???? By Kin 2011/04/13
    FaciCode.SQL.Text := Format(
    ' SELECT A.CODENO,                                          ' +
    '        A.DESCRIPTION,                                     ' +
    '        A.REFNO                                            ' +
    '   FROM %s.CD022 A                                         ' +
    '  WHERE A.STOPFLAG = 0                                     ' +
    aString                                                       +
    '    AND ( A.SERVICETYPE=''%s'' Or A.SERVICETYPE IS NULL )  ' +
    ' ORDER BY A.CODENO                                         ',
    [DBController.LoginInfo.DbAccount, ServiceType.Value] );
  end;
  FaciCode.CodeNamePropertiesInitPopup( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.ServiceTypeCodeNamePropertiesInitPopup(
  Sender: TObject);
begin
  ServiceType.SQL.Text := Format(
    ' SELECT CODENO, DESCRIPTION    ' +
    '  FROM %s.CD046                ' ,
    [DBController.LoginInfo.DbAccount] );
  ServiceType.CodeNamePropertiesInitPopup( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.ModelCodeCodeNamePropertiesInitPopup(
  Sender: TObject);
begin
  { ?]?????? }
  ModelCode.SQL.Text := Format(
    ' SELECT A.CODENO, A.DESCRIPTION               ' +
    '   FROM %s.CD043 A                            ' +
    '  WHERE A.STOPFLAG = 0                        ' +
    '    AND A.SERVICETYPE = ''%s''                ' +
    ' ORDER BY A.CODENO                            ',
    [DBController.LoginInfo.DbAccount, ServiceType.Value] );
  ModelCode.CodeNamePropertiesInitPopup( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.BuyCodeCodeNamePropertiesInitPopup(Sender: TObject);
begin
  { ?R?????? }
  BuyCode.SQL.Text := Format(
    ' SELECT A.CODENO, A.DESCRIPTION ' +
    '   FROM %s.CD034 A              ' +
    '  WHERE A.STOPFLAG = 0          ' +
    ' ORDER BY A.CODENO              ',
    [DBController.LoginInfo.DbAccount] );
  BuyCode.CodeNamePropertiesInitPopup( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.CMBaudNoCodeNamePropertiesInitPopup(Sender: TObject);
begin
  { CM?t?v }
  CMBaudNo.SQL.Text := Format(
    ' SELECT A.CODENO, A.DESCRIPTION ' +
    '   FROM %s.CD064 A              ' +
    '  WHERE A.SERVICETYPE = ''%s''  ' +
    '    AND A.STOPFLAG = 0          ' +
    ' ORDER BY A.CODENO              ',
    [DBController.LoginInfo.DbAccount, ServiceType.Value] );
  CMBaudNo.CodeNamePropertiesInitPopup( Sender );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.LoadCodeData;
var
  aItem: TcxImageComboBoxItem;
begin
  { ?T?wIP }
  cmbFixIpCount.Properties.Items.Clear;
  cmbFixIpCount.Properties.Items.Add;
  FReader.Close;
  FReader.SQL.Text := Format(
    ' SELECT DISTINCT NVL( A.IPQUANTITY, 0 ) AS IPQUANTITY ' +
    '   FROM %s.CD074 A                                    ' +
    '  WHERE A.COMPCODE = ''%s''                           ' +
    '    AND A.IPAPPLY IN ( 1, 3 )                         ' +
    '    AND A.STOPFLAG=0                                  ' +
    ' ORDER BY NVL( A.IPQUANTITY, 0 )                      ',
    [DBController.LoginInfo.DbAccount,  DBController.LoginInfo.CompCode] );
  FReader.Open;
  FReader.First;
  while not FReader.Eof do
  begin
    aItem := cmbFixIpCount.Properties.Items.Add;
    aItem.Value := FReader.FieldByName( 'IPQUANTITY' ).AsString;
    aItem.Description := Format( '%s', [FReader.FieldByName( 'IPQUANTITY' ).AsString] );
    FReader.Next;
  end;
  FReader.Close;
  { ?B??IP }
  cmbDynIpCount.Properties.Items.Clear;
  cmbDynIpCount.Properties.Items.Add;
  FReader.Close;
  FReader.SQL.Text := Format(
    ' SELECT DISTINCT NVL( A.IPQUANTITY, 0 ) AS IPQUANTITY ' +
    '   FROM %s.CD074 A                                    ' +
    '  WHERE A.COMPCODE = ''%s''                           ' +
    '    AND A.IPAPPLY IN ( 2, 3 )                         ' +
    '    AND A.STOPFLAG=0                                  ' +
    ' ORDER BY NVL( A.IPQUANTITY, 0 )                      ',
    [DBController.LoginInfo.DbAccount, DBController.LoginInfo.CompCode] );
  FReader.Open;
  FReader.First;
  while not FReader.Eof do
  begin
    aItem := cmbDynIpCount.Properties.Items.Add;
    aItem.Value := FReader.FieldByName( 'IPQUANTITY' ).AsString;
    aItem.Description := Format( '%s', [FReader.FieldByName( 'IPQUANTITY' ).AsString] );
    FReader.Next;
  end;
  FReader.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.cmbFaciTypePropertiesChange(Sender: TObject);
var
  aEnabled: Boolean;
begin
  FaciCode.Clear;
  cmbFixIpCount.ItemIndex := 0;
  cmbDynIpCount.ItemIndex := 0;
  BuyCode.Clear;
  CMBaudNo.Clear;
  {}
  aEnabled := ( ServiceType.Value = 'I' ) or ( ServiceType.Value = 'P' );
  CMBaudNo.Enabled := aEnabled;
  cmbFixIpCount.Enabled := aEnabled;
  cmbDynIpCount.Enabled := aEnabled;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.FaciCodeCodeNoPropertiesChange(Sender: TObject);
begin
  FaciCode.CodeNoPropertiesChange( Sender );
  ModelCode.Clear;
  ChangeIpCountState;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.FaciCodeCodeNamePropertiesChange(Sender: TObject);
begin
  FaciCode.CodeNamePropertiesChange( Sender );
  LoadFaciCodeRelationItem;
  ChangeIpCountState;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.ServiceTypeCodeNoPropertiesChange(Sender: TObject);
begin
  ServiceType.CodeNoPropertiesChange( Sender );
  FaciCode.Clear;
  ModelCode.Clear;
  BuyCode.Clear;
  ChangeIpCountState;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.ServiceTypeCodeNamePropertiesChange(Sender: TObject);
begin
  ServiceType.CodeNamePropertiesChange( Sender );
  ServiceTypeCodeNoPropertiesChange( Sender );
  {
  CMBaudNo.Enabled := ( ServiceType.Value = 'I' ) or ( ServiceType.Value = 'P' );
  if ( not CMBaudNo.Enabled ) then
  begin
    CMBaudNo.Clear;
    cmbFixIpCount.ItemIndex := 0;
    cmbDynIpCount.ItemIndex := 0;
  end;
  cmbFixIpCount.Enabled := CMBaudNo.Enabled;
  cmbDynIpCount.Enabled := CMBaudNo.Enabled;
  }
  ChangeIpCountState;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.btnInstCodeStrClick(Sender: TObject);
begin
  if ( FMode in [dmInsert, dmUpdate] ) then
  begin
    frmMultiSelect := TfrmMultiSelect.Create( Application );
    try
      frmMultiSelect.Connection := DBController.DataConnection;
      frmMultiSelect.KeyFields := 'INSTCODE';
      frmMultiSelect.KeyValues := FQueryParam.Param1;
      frmMultiSelect.DisplayFields := 'INSTCODE,?????N?X,INSTNAME,???????O?W??';
      frmMultiSelect.ResultFields := 'INSTNAME';
      
      FInstCodeDataSet.Filtered := False;
      FInstCodeDataSet.Filter := Format( 'SERVICETYPE=''%s''', [ServiceType.Value] );
      FInstCodeDataSet.Filtered := True;

      try
        frmMultiSelect.DataSet := FInstCodeDataSet;
        if frmMultiSelect.ShowModal = mrOk then
        begin
          if ( FMode in [dmInsert, dmUpdate] ) then
          begin
            FQueryParam.Param1 := frmMultiSelect.SelectedValue;
            edtInstCodeStr.Text := frmMultiSelect.SelectedDisplay;
          end;
        end;
      finally
        FInstCodeDataSet.Filtered := False;
        FInstCodeDataSet.Filter := EmptyStr;
      end;
    finally
      frmMultiSelect.Free;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmCD078B2.cxButton1Click(Sender: TObject);
begin
  if ( FMode in [dmInsert, dmUpdate] ) then
  begin
    FQueryParam.Param1 := EmptyStr;
    edtInstCodeStr.Text := EmptyStr;
  end;
end;

{ ---------------------------------------------------------------------------- }

end.
