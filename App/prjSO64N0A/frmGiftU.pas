unit frmGiftU;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ActnList, ExtCtrls, StdCtrls, Buttons, DB, DBClient, ADODB;

type
  TfrmGift = class(TForm)
    ActionList1: TActionList;
    actSave: TAction;
    actCancel: TAction;
    Panel1: TPanel;
    Panel2: TPanel;
    Bevel1: TBevel;
    btnSave: TButton;
    btnCancel: TButton;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Panel3: TPanel;
    EDT_GiftOrderPrice1: TEdit;
    Label8: TLabel;
    EDT_GiftOrderPriceMax1: TEdit;
    EDT_GiftType1: TComboBox;
    EDT_GiftKind1: TComboBox;
    Button1: TButton;
    X_ArticleNoStr1: TBitBtn;
    EDT_ArticleNoStr1: TEdit;
    EDT_ArticleNo1: TListBox;
    Label11: TLabel;
    Panel4: TPanel;
    Label12: TLabel;
    EDT_GiftOrderPrice2: TEdit;
    EDT_GiftOrderPriceMax2: TEdit;
    EDT_GiftType2: TComboBox;
    EDT_GiftKind2: TComboBox;
    Button2: TButton;
    X_ArticleNoStr2: TBitBtn;
    EDT_ArticleNoStr2: TEdit;
    EDT_ArticleNo2: TListBox;
    Panel5: TPanel;
    Label13: TLabel;
    EDT_GiftOrderPrice3: TEdit;
    EDT_GiftOrderPriceMax3: TEdit;
    EDT_GiftType3: TComboBox;
    EDT_GiftKind3: TComboBox;
    Button3: TButton;
    X_ArticleNoStr3: TBitBtn;
    EDT_ArticleNoStr3: TEdit;
    EDT_ArticleNo3: TListBox;
    Panel6: TPanel;
    Label14: TLabel;
    EDT_GiftOrderPrice4: TEdit;
    EDT_GiftOrderPriceMax4: TEdit;
    EDT_GiftType4: TComboBox;
    EDT_GiftKind4: TComboBox;
    Button4: TButton;
    X_ArticleNoStr4: TBitBtn;
    EDT_ArticleNoStr4: TEdit;
    EDT_ArticleNo4: TListBox;
    Panel7: TPanel;
    Label15: TLabel;
    EDT_GiftOrderPrice5: TEdit;
    EDT_GiftOrderPriceMax5: TEdit;
    EDT_GiftType5: TComboBox;
    EDT_GiftKind5: TComboBox;
    Button5: TButton;
    X_ArticleNoStr5: TBitBtn;
    EDT_ArticleNoStr5: TEdit;
    EDT_ArticleNo5: TListBox;
    Panel8: TPanel;
    Label16: TLabel;
    EDT_GiftOrderPrice6: TEdit;
    EDT_GiftOrderPriceMax6: TEdit;
    EDT_GiftType6: TComboBox;
    EDT_GiftKind6: TComboBox;
    Button6: TButton;
    X_ArticleNoStr6: TBitBtn;
    EDT_ArticleNoStr6: TEdit;
    EDT_ArticleNo6: TListBox;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure EDT_GiftOrderPrice1Change(Sender: TObject);
    procedure actSaveExecute(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure X_ArticleNoStr1Click(Sender: TObject);
    procedure actCancelExecute(Sender: TObject);
    procedure EDT_GiftOrderPrice1Exit(Sender: TObject);
    procedure EDT_GiftKind1Change(Sender: TObject);
    procedure EDT_GiftOrderPrice1KeyPress(Sender: TObject; var Key: Char);
    procedure EDT_GiftOrderPriceMax1Exit(Sender: TObject);
  private
    { Private declarations }
    FCD042DataSet: TClientDataSet;
    FDataReader: TADOQuery;
    FGifted: Boolean;
    FServcieTypes: String;
    FActStopDate: String;
    FActStartDate: String;
    procedure InitGiftTypeControl(const aControl: TComboBox; const aAddItem: Boolean);
    procedure InitGiftKindControl(const aControl: TComboBox; const aAddItem: Boolean);
    procedure SetEnableGiftControl(aIndex: Integer; const aEnabled: Boolean);
    procedure ShowArticleNo(const aEdit:TEdit; const aListBox:TListBox; const aGiftKind: Integer);
    procedure ClearArticleNoStr(const aEdit:TEdit; const aListBox:TListBox);
    procedure FollowPrice(const aEdit: TEdit; const aValue: String);
    procedure ShowMultiForm(const aIndex: Integer);
    procedure DataToEditor;
    procedure EditorToData;
    {}
    function CheckData: Boolean;
    function CheckPriceMinMax(const aValue: Integer): Boolean;
    function GetGiftTypeStr(aControl: TComboBox): String;
    function GetGiftKindStr(aControl: TComboBox): String;
    function GetArticleNoStr(aControl: TListBox): String;
  public
    { Public declarations }
    property CD042DataSet: TClientDataSet read FCD042DataSet write FCD042DataSet;
    property DataReader: TADOQuery read FDataReader write FDataReader;
    property Gifted: Boolean read FGifted write FGifted;
    property ServcieTypes: String read FServcieTypes write FServcieTypes;
    property ActStartDate: String read FActStartDate write FActStartDate;
    property ActStopDate: String read FActStopDate write FActStopDate;
  end;

var
  frmGift: TfrmGift;

implementation

uses cbUtilis, frmcd042, frmMultiSelectU;

{$R *.dfm}

{ ---------------------------------------------------------------------------- }

procedure TfrmGift.FormCreate(Sender: TObject);
var
  aIndex: Integer;
begin
  for aIndex := 1 to 6  do
    SetEnableGiftControl( aIndex, False );
  FGifted := False;  
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmGift.FormShow(Sender: TObject);
begin
  DataToEditor;
  if EDT_GiftOrderPrice1.CanFocus then EDT_GiftOrderPrice1.SetFocus;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmGift.FormDestroy(Sender: TObject);
begin
  {...}
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmGift.InitGiftTypeControl(const aControl: TComboBox; const aAddItem: Boolean);
begin
  if Assigned( aControl ) then
  begin
    if ( aAddItem ) and ( aControl.Items.Count <= 0 ) then
    begin
      aControl.Items.Add( '0.全部贈送' );
      aControl.Items.Add( '1.多選1' );
      aControl.Items.Add( '2.多選2' );
      aControl.Items.Add( '3.多選3' );
      aControl.Items.Add( '4.多選4' );
      aControl.Items.Add( '5.多選5' );
      aControl.Items.Add( '6.多選6' );
      aControl.Items.Add( '7.多選7' );
      aControl.Items.Add( '8.多選8' );
      aControl.Items.Add( '9.多選9' );
      aControl.ItemIndex := 0;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmGift.InitGiftKindControl(const aControl: TComboBox; const aAddItem: Boolean);
begin
  if Assigned( aControl ) then
  begin
    if ( aAddItem ) and ( aControl.Items.Count <= 0 ) then
    begin
      aControl.Items.Add( '1.贈品' );
      aControl.Items.Add( '2.付費頻道' );
      aControl.ItemIndex := 0;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmGift.SetEnableGiftControl(aIndex: Integer; const aEnabled: Boolean);
begin
  case aIndex of
    1:
      begin
        EDT_GiftType1.Enabled := aEnabled;
        EDT_GiftKind1.Enabled := aEnabled;
        Button1.Enabled := aEnabled;
        X_ArticleNoStr1.Enabled := aEnabled;
        EDT_ArticleNoStr1.Enabled := aEnabled;
        InitGiftTypeControl( EDT_GiftType1, aEnabled );
        InitGiftKindControl( EDT_GiftKind1, aEnabled );
        if not aEnabled then
        begin
          EDT_GiftType1.Clear;
          EDT_GiftKind1.Clear;
        end;
      end;
    2:
      begin
        EDT_GiftType2.Enabled := aEnabled;
        EDT_GiftKind2.Enabled := aEnabled;
        Button2.Enabled := aEnabled;
        X_ArticleNoStr2.Enabled := aEnabled;
        EDT_ArticleNoStr2.Enabled := aEnabled;
        InitGiftTypeControl( EDT_GiftType2, aEnabled );
        InitGiftKindControl( EDT_GiftKind2, aEnabled );
        if not aEnabled then
        begin
          EDT_GiftType2.Clear;
          EDT_GiftKind2.Clear;
        end;
      end;
    3:
      begin
        EDT_GiftType3.Enabled := aEnabled;
        EDT_GiftKind3.Enabled := aEnabled;
        Button3.Enabled := aEnabled;
        X_ArticleNoStr3.Enabled := aEnabled;
        EDT_ArticleNoStr3.Enabled := aEnabled;
        InitGiftTypeControl( EDT_GiftType3, aEnabled );
        InitGiftKindControl( EDT_GiftKind3, aEnabled );
        if not aEnabled then
        begin
          EDT_GiftType3.Clear;
          EDT_GiftKind3.Clear;
        end;
      end;
    4:
      begin
        EDT_GiftType4.Enabled := aEnabled;
        EDT_GiftKind4.Enabled := aEnabled;
        Button4.Enabled := aEnabled;
        X_ArticleNoStr4.Enabled := aEnabled;
        EDT_ArticleNoStr4.Enabled := aEnabled;
        InitGiftTypeControl( EDT_GiftType4, aEnabled );
        InitGiftKindControl( EDT_GiftKind4, aEnabled );
        if not aEnabled then
        begin
          EDT_GiftType4.Clear;
          EDT_GiftKind4.Clear;
        end;
      end;
    5:
      begin
        EDT_GiftType5.Enabled := aEnabled;
        EDT_GiftKind5.Enabled := aEnabled;
        Button5.Enabled := aEnabled;
        X_ArticleNoStr5.Enabled := aEnabled;
        EDT_ArticleNoStr5.Enabled := aEnabled;
        InitGiftTypeControl( EDT_GiftType5, aEnabled );
        InitGiftKindControl( EDT_GiftKind5, aEnabled );
        if not aEnabled then
        begin
          EDT_GiftType5.Clear;
          EDT_GiftKind5.Clear;
        end;
      end;
    6:
      begin
        EDT_GiftType6.Enabled := aEnabled;
        EDT_GiftKind6.Enabled := aEnabled;
        Button6.Enabled := aEnabled;
        X_ArticleNoStr6.Enabled := aEnabled;
        EDT_ArticleNoStr6.Enabled := aEnabled;
        InitGiftTypeControl( EDT_GiftType6, aEnabled );
        InitGiftKindControl( EDT_GiftKind6, aEnabled );
        if not aEnabled then
        begin
          EDT_GiftType6.Clear;
          EDT_GiftKind6.Clear;
        end;
      end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmGift.ShowArticleNo(const aEdit: TEdit; const aListBox: TListBox;
  const aGiftKind: Integer);
var
  aWhereText, aSql: String;
  aIndex: Integer;
begin
  aEdit.Text := EmptyStr;
  aWhereText := EmptyStr;
  for aIndex := 0 to aListBox.Items.Count-1 do
  begin
    if aIndex = 0 then
      aWhereText := aWhereText + ' WHERE ('
    else
      aWhereText := aWhereText + 'OR ';
    aWhereText := aWhereText + '( CODENO = ' + aListBox.Items[aIndex] + ') ';
  end;
  if ( aWhereText <> EmptyStr ) then
  begin
    if ( aGiftKind = 1 ) then
    begin
      aSql := Format( ' SELECT DESCRIPTION FROM %s.CD059 %s ) And ( COMPCODE = %s ) ',
        [Form1.sG_DbUserID, aWhereText, Form1.sG_CompID] );
    end else
    begin
      aSql := Format( ' SELECT DESCRIPTION FROM %s.CD024 %s ) And ( COMPCODE = %s ) ',
        [Form1.sG_DbUserID, aWhereText, Form1.sG_CompID] );
    end;
    FDataReader.Close;
    FDataReader.SQL.Text := aSql;
    FDataReader.Prepared := True;
    FDataReader.Open;
    while not FDataReader.Eof do
    begin
      aEdit.Text := ( aEdit.Text + FDataReader.FieldByName('Description').AsString );
      FDataReader.Next;
      if not FDataReader.Eof then aEdit.Text := ( aEdit.Text + ',' );
    end;
  end;
end;

{ ----------------------------------------------------------------------------- }

procedure TfrmGift.ShowMultiForm(const aIndex: Integer);
var
  aEdit: TEdit;
  aCombox: TListBox;
  aSql, aSelectedValue, aTmp: String;
  aGiftKind: Integer;

  { ---------------------------------------------------- }

  function GetValues(const aControl: TListBox): String;
  var
    aIndex: Integer;
  begin
    Result := EmptyStr;
    for aIndex := 0 to aControl.Items.Count - 1 do
      Result := ( Result + aControl.Items[aIndex] + ',' );
    if IsDelimiter( ',', Result, Length( Result ) ) then
      Delete( Result, Length( Result ), 1 );
  end;

  { ---------------------------------------------------- }

begin
  case aIndex of
    1:
      begin
        aEdit := EDT_ArticleNoStr1;
        aCombox := EDT_ArticleNo1;
        aGiftKind := ( EDT_GiftKind1.ItemIndex + 1 );
      end;
    2:
      begin
        aEdit := EDT_ArticleNoStr2;
        aCombox := EDT_ArticleNo2;
        aGiftKind := ( EDT_GiftKind2.ItemIndex + 1 );
      end;
    3:
      begin
        aEdit := EDT_ArticleNoStr3;
        aCombox := EDT_ArticleNo3;
        aGiftKind := ( EDT_GiftKind3.ItemIndex + 1 );
      end;
    4:
      begin
        aEdit := EDT_ArticleNoStr4;
        aCombox := EDT_ArticleNo4;
        aGiftKind := ( EDT_GiftKind4.ItemIndex + 1 );
      end;
    5:
      begin
        aEdit := EDT_ArticleNoStr5;
        aCombox := EDT_ArticleNo5;
        aGiftKind := ( EDT_GiftKind5.ItemIndex + 1 );
      end;
    6:
      begin
        aEdit := EDT_ArticleNoStr6;
        aCombox := EDT_ArticleNo6;
        aGiftKind := ( EDT_GiftKind6.ItemIndex + 1 );
      end;
  else
    aGiftKind := -1;
    aEdit := nil;
    aCombox := nil;
  end;
  aSelectedValue := GetValues( aCombox );
  if aGiftKind = 1 then
  begin
    aSql := Format(
      'SELECT CodeNo, Description FROM %s.CD059 WHERE StopFlag = 0 AND Quantity > 0 ',
      [Form1.sG_DbUserID] );
    { 服務別 }
    if ( FServcieTypes <> EmptyStr ) then
    begin
      aTmp := FServcieTypes;
      aTmp := ExtractValue( aTmp );
      aSql := ( aSql + Format(
        ' AND NVL( ServiceType, %0:s ) IN ( %1:s )', [aTmp, FServcieTypes] ) );
    end;
    { 贈品到期日 }
    aSql := ( aSql + Format(
       ' AND (          ' +
       '   (            ' +
       '     NVL( StartDate, TO_DATE( ''%0:s'', ''YYYY/MM/DD'' ) ) BETWEEN TO_DATE( ''%0:s'', ''YYYY/MM/DD'' ) AND TO_DATE( ''%1:s'', ''YYYY/MM/DD''  ) OR ' +
       '     NVL( StopDate,  TO_DATE( ''%0:s'', ''YYYY/MM/DD'' ) ) BETWEEN TO_DATE( ''%0:s'', ''YYYY/MM/DD'' ) AND TO_DATE( ''%1:s'', ''YYYY/MM/DD''  )    ' +
       '   )            ' +
       '   OR           ' +
       '   (            ' +
       '     NVL( StartDate, TO_DATE( ''%0:s'', ''YYYY/MM/DD'' ) ) <= TO_DATE( ''%1:s'', ''YYYY/MM/DD'' ) AND ' +
       '     NVL( StopDate,  TO_DATE( ''%0:s'', ''YYYY/MM/DD'' ) ) >= TO_DATE( ''%1:s'', ''YYYY/MM/DD'' )    ' +
       '   ) ' +
       ' )  AND NVL( StopDate, SYSDATE ) >= SYSDATE ', [FActStartDate, FActStopDate] ) );
  end else
  begin
    aSql := Format( 'SELECT CODENO, Description From %s.CD024 Where StopFlag = 0',
      [Form1.sG_DbUserID] );
  end;
  frmMultiSelect:=TfrmMultiSelect.Create(Application);
  try
    frmMultiSelect.Connection := Form1.ADOConnection1;
    frmMultiSelect.KeyFields:='CODENO';
    frmMultiSelect.KeyValues:= aSelectedValue;
    frmMultiSelect.DisplayFields:='CODENO,產品編號,Description,產品名稱';
    frmMultiSelect.ResultFields:='Description';
    frmMultiSelect.SQL.Text := aSql;
    frmMultiSelect.QuotedString:='';
    if frmMultiSelect.ShowModal=mrOk then
    begin
      aSelectedValue := frmMultiSelect.SelectedValue;
      aCombox.Clear;
      repeat
        aCombox.Items.Add( ExtractValue( aSelectedValue ) );
      until ( aSelectedValue = EmptyStr );
      if Assigned( aEdit ) then
        aEdit.Text := frmMultiSelect.SelectedDisplay;
    end;
  finally
    frmMultiSelect.Free;
  end;
end;

{ ----------------------------------------------------------------------------- }

procedure TfrmGift.ClearArticleNoStr(const aEdit: TEdit; const aListBox: TListBox);
begin
  if Assigned( aEdit ) then aEdit.Text := EmptyStr;
  if Assigned( aListBox ) then aListBox.Items.Clear;
end;

{ ----------------------------------------------------------------------------- }

procedure TfrmGift.FollowPrice(const aEdit: TEdit; const aValue: String);
begin
  if Assigned( aEdit ) then
  begin
    if ( aEdit.Text = EmptyStr ) and ( aValue <> EmptyStr ) then
      aEdit.Text := aValue;
  end;
end;

{ ----------------------------------------------------------------------------- }

procedure TfrmGift.DataToEditor;
var
  DataString, String1: String;
begin

  EDT_GiftKind1.OnChange := nil;
  EDT_GiftKind2.OnChange := nil;
  EDT_GiftKind3.OnChange := nil;
  EDT_GiftKind4.OnChange := nil;
  EDT_GiftKind5.OnChange := nil;
  EDT_GiftKind6.OnChange := nil;

  { Gift1 }
  EDT_GiftOrderPrice1.Text := FCD042DataSet.FieldByName('GiftOrderPrice1').AsString;
  EDT_GiftOrderPriceMax1.Text := Nvl(
    FCD042DataSet.FieldByName('GiftOrderPriceMax1').AsString,
    FCD042DataSet.FieldByName('GiftOrderPrice1').AsString );
    
  EDT_GiftType1.ItemIndex:=-1;
  EDT_GiftKind1.ItemIndex:=-1;

  if FCD042DataSet.FieldByName('GiftType1').AsString <> '' then
    EDT_GiftType1.ItemIndex:=FCD042DataSet.FieldByName('GiftType1').AsInteger;
  if FCD042DataSet.FieldByName('GiftKind1').AsString <> '' then
    EDT_GiftKind1.ItemIndex:= FCD042DataSet.FieldByName('GiftKind1').AsInteger-1;

  DataString := FCD042DataSet.FieldByName('ArticleNoStr1').AsString;
  while DataString<>'' do
  begin
    String1 := ExtractValue(DataString);
    EDT_ArticleNo1.Items.Add(String1);
  end;
  ShowArticleNo( EDT_ArticleNoStr1, EDT_ArticleNo1,
    FCD042DataSet.FieldByName('GiftKind1').AsInteger );

  { Gift2 }
    EDT_GiftOrderPrice2.Text := FCD042DataSet.FieldByName('GiftOrderPrice2').AsString;
  EDT_GiftOrderPriceMax2.Text := Nvl(
    FCD042DataSet.FieldByName('GiftOrderPriceMax2').AsString,
   FCD042DataSet.FieldByName('GiftOrderPrice2').AsString );

  EDT_GiftType2.ItemIndex:=-1;
  EDT_GiftKind2.ItemIndex:=-1;

  if FCD042DataSet.FieldByName('GiftType2').AsString <> '' then
    EDT_GiftType2.ItemIndex:=FCD042DataSet.FieldByName('GiftType2').AsInteger;
  if FCD042DataSet.FieldByName('GiftKind2').AsString <> '' then
    EDT_GiftKind2.ItemIndex:=FCD042DataSet.FieldByName('GiftKind2').AsInteger-1;

  DataString := FCD042DataSet.FieldByName('ArticleNoStr2').AsString;
  while DataString<>'' do
  begin
    String1 := ExtractValue(DataString);
    EDT_ArticleNo2.Items.Add(String1);
  end;
  ShowArticleNo( EDT_ArticleNoStr2, EDT_ArticleNo2,
    FCD042DataSet.FieldByName('GiftKind2').AsInteger );

  { Gift3 }
  EDT_GiftOrderPrice3.Text := FCD042DataSet.FieldByName('GiftOrderPrice3').AsString;
  EDT_GiftOrderPriceMax3.Text := Nvl(
    FCD042DataSet.FieldByName('GiftOrderPriceMax3').AsString,
    FCD042DataSet.FieldByName('GiftOrderPrice3').AsString );
  EDT_GiftType3.ItemIndex:= -1;
  EDT_GiftKind3.ItemIndex:=-1;

  if FCD042DataSet.FieldByName('GiftType3').AsString <> '' then
    EDT_GiftType3.ItemIndex := FCD042DataSet.FieldByName('GiftType3').AsInteger;
  if FCD042DataSet.FieldByName('GiftKind3').AsString <> '' then
    EDT_GiftKind3.ItemIndex:=FCD042DataSet.FieldByName('GiftKind3').AsInteger-1;

  DataString := FCD042DataSet.FieldByName('ArticleNoStr3').AsString;
  while DataString<>'' do
  begin
    String1 := ExtractValue(DataString);
    EDT_ArticleNo3.Items.Add(String1);
  end;
  ShowArticleNo( EDT_ArticleNoStr3, EDT_ArticleNo3,
    FCD042DataSet.FieldByName('GiftKind2').AsInteger );

  { Gift4 }
  EDT_GiftOrderPrice4.Text := FCD042DataSet.FieldByName('GiftOrderPrice4').AsString;
  EDT_GiftOrderPriceMax4.Text := Nvl(
    FCD042DataSet.FieldByName('GiftOrderPriceMax4').AsString,
    FCD042DataSet.FieldByName('GiftOrderPrice4').AsString );

  EDT_GiftType4.ItemIndex:=-1;
  EDT_GiftKind4.ItemIndex:=-1;

  if FCD042DataSet.FieldByName('GiftType4').AsString <> '' then
    EDT_GiftType4.ItemIndex := FCD042DataSet.FieldByName('GiftType4').AsInteger;
  if FCD042DataSet.FieldByName('GiftKind4').AsString <> '' then
    EDT_GiftKind4.ItemIndex:= FCD042DataSet.FieldByName('GiftKind4').AsInteger-1;

  DataString := FCD042DataSet.FieldByName('ArticleNoStr4').AsString;
  while DataString<>'' do
  begin
    String1 := ExtractValue(DataString);
    EDT_ArticleNo4.Items.Add(String1);
  end;
  ShowArticleNo( EDT_ArticleNoStr4, EDT_ArticleNo4,
    FCD042DataSet.FieldByName('GiftKind4').AsInteger );

  { Gift5 }
  EDT_GiftOrderPrice5.Text := FCD042DataSet.FieldByName('GiftOrderPrice5').AsString;
  EDT_GiftOrderPriceMax5.Text := Nvl(
    FCD042DataSet.FieldByName('GiftOrderPriceMax5').AsString,
    FCD042DataSet.FieldByName('GiftOrderPrice5').AsString );

  EDT_GiftType5.ItemIndex:=-1;
  EDT_GiftKind5.ItemIndex:=-1;
  
  if FCD042DataSet.FieldByName('GiftType5').AsString <> '' then
    EDT_GiftType5.ItemIndex := FCD042DataSet.FieldByName('GiftType5').AsInteger;
  if FCD042DataSet.FieldByName('GiftKind5').AsString <> '' then
    EDT_GiftKind5.ItemIndex := FCD042DataSet.FieldByName('GiftKind5').AsInteger-1;

  DataString := FCD042DataSet.FieldByName('ArticleNoStr5').AsString;
  while DataString<>'' do
  begin
    String1 := ExtractValue(DataString);
    EDT_ArticleNo5.Items.Add(String1);
  end;
  ShowArticleNo( EDT_ArticleNoStr5, EDT_ArticleNo5,
    FCD042DataSet.FieldByName('GiftKind5').AsInteger );

  { Gift6 }
  EDT_GiftOrderPrice6.Text := FCD042DataSet.FieldByName('GiftOrderPrice6').AsString;
  EDT_GiftOrderPriceMax6.Text := Nvl(
    FCD042DataSet.FieldByName('GiftOrderPriceMax6').AsString,
    FCD042DataSet.FieldByName('GiftOrderPrice6').AsString );

  EDT_GiftType6.ItemIndex:=-1;
  EDT_GiftKind6.ItemIndex:=-1;

  if FCD042DataSet.FieldByName('GiftType6').AsString <> '' then
    EDT_GiftType6.ItemIndex := FCD042DataSet.FieldByName('GiftType6').AsInteger;
  if FCD042DataSet.FieldByName('GiftKind6').AsString <> '' then
    EDT_GiftKind6.ItemIndex := FCD042DataSet.FieldByName('GiftKind6').AsInteger-1;

  DataString := FCD042DataSet.FieldByName('ArticleNoStr6').AsString;
  while DataString<>'' do
  begin
    String1 := ExtractValue(DataString);
    EDT_ArticleNo6.Items.Add(String1);
  end;
  ShowArticleNo( EDT_ArticleNoStr6, EDT_ArticleNo6,
    FCD042DataSet.FieldByName('GiftKind6').AsInteger );

  EDT_GiftKind1.OnChange := EDT_GiftKind1Change;
  EDT_GiftKind2.OnChange := EDT_GiftKind1Change;
  EDT_GiftKind3.OnChange := EDT_GiftKind1Change;
  EDT_GiftKind4.OnChange := EDT_GiftKind1Change;
  EDT_GiftKind5.OnChange := EDT_GiftKind1Change;
  EDT_GiftKind6.OnChange := EDT_GiftKind1Change;

  FGifted :=
    ( EDT_GiftOrderPrice1.Text <> EmptyStr ) or
    ( EDT_GiftOrderPrice2.Text <> EmptyStr ) or
    ( EDT_GiftOrderPrice3.Text <> EmptyStr ) or
    ( EDT_GiftOrderPrice4.Text <> EmptyStr ) or
    ( EDT_GiftOrderPrice5.Text <> EmptyStr ) or
    ( EDT_GiftOrderPrice6.Text <> EmptyStr );
    
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmGift.EditorToData;
begin
  FCD042DataSet.FieldByName( 'GiftOrderPrice1' ).AsString := EDT_GiftOrderPrice1.Text;
  FCD042DataSet.FieldByName( 'GiftOrderPriceMax1' ).AsString := EDT_GiftOrderPriceMax1.Text;
  FCD042DataSet.FieldByName( 'GiftType1' ).AsString := GetGiftTypeStr( EDT_GiftType1 );
  FCD042DataSet.FieldByName( 'GiftKind1' ).AsString := GetGiftKindStr( EDT_GiftKind1 );
  FCD042DataSet.FieldByName( 'ArticleNoStr1').AsString := GetArticleNoStr( EDT_ArticleNo1 );
  {}
  FCD042DataSet.FieldByName( 'GiftOrderPrice2' ).AsString := EDT_GiftOrderPrice2.Text;
  FCD042DataSet.FieldByName( 'GiftOrderPriceMax2' ).AsString := EDT_GiftOrderPriceMax2.Text;
  FCD042DataSet.FieldByName( 'GiftType2' ).AsString := GetGiftTypeStr( EDT_GiftType2 );
  FCD042DataSet.FieldByName( 'GiftKind2' ).AsString := GetGiftKindStr( EDT_GiftKind2 );
  FCD042DataSet.FieldByName( 'ArticleNoStr2').AsString := GetArticleNoStr( EDT_ArticleNo2 );
  {}
  FCD042DataSet.FieldByName( 'GiftOrderPrice3' ).AsString := EDT_GiftOrderPrice3.Text;
  FCD042DataSet.FieldByName( 'GiftOrderPriceMax3' ).AsString := EDT_GiftOrderPriceMax3.Text;
  FCD042DataSet.FieldByName( 'GiftType3' ).AsString := GetGiftTypeStr( EDT_GiftType3 );
  FCD042DataSet.FieldByName( 'GiftKind3' ).AsString := GetGiftKindStr( EDT_GiftKind3 );
  FCD042DataSet.FieldByName( 'ArticleNoStr3').AsString := GetArticleNoStr( EDT_ArticleNo3 );
  {}
  FCD042DataSet.FieldByName( 'GiftOrderPrice4' ).AsString := EDT_GiftOrderPrice4.Text;
  FCD042DataSet.FieldByName( 'GiftOrderPriceMax4' ).AsString := EDT_GiftOrderPriceMax4.Text;
  FCD042DataSet.FieldByName( 'GiftType4' ).AsString := GetGiftTypeStr( EDT_GiftType4 );
  FCD042DataSet.FieldByName( 'GiftKind4' ).AsString := GetGiftKindStr( EDT_GiftKind4 );
  FCD042DataSet.FieldByName( 'ArticleNoStr4').AsString := GetArticleNoStr( EDT_ArticleNo4 );
  {}
  FCD042DataSet.FieldByName( 'GiftOrderPrice5' ).AsString := EDT_GiftOrderPrice5.Text;
  FCD042DataSet.FieldByName( 'GiftOrderPriceMax5' ).AsString := EDT_GiftOrderPriceMax5.Text;
  FCD042DataSet.FieldByName( 'GiftType5' ).AsString := GetGiftTypeStr( EDT_GiftType5 );
  FCD042DataSet.FieldByName( 'GiftKind5' ).AsString := GetGiftKindStr( EDT_GiftKind5 );
  FCD042DataSet.FieldByName( 'ArticleNoStr5').AsString := GetArticleNoStr( EDT_ArticleNo5 );
  {}
  FCD042DataSet.FieldByName( 'GiftOrderPrice6' ).AsString := EDT_GiftOrderPrice6.Text;
  FCD042DataSet.FieldByName( 'GiftOrderPriceMax6' ).AsString := EDT_GiftOrderPriceMax6.Text;
  FCD042DataSet.FieldByName( 'GiftType6' ).AsString := GetGiftTypeStr( EDT_GiftType6 );
  FCD042DataSet.FieldByName( 'GiftKind6' ).AsString := GetGiftKindStr( EDT_GiftKind6 );
  FCD042DataSet.FieldByName( 'ArticleNoStr6').AsString := GetArticleNoStr( EDT_ArticleNo6 );
  {}
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmGift.EDT_GiftOrderPrice1Change(Sender: TObject);
var
  aIndex: Integer;
  aEnabled: Boolean;
begin
  aIndex := -1;
  aEnabled := False;
  if ( Sender = EDT_GiftOrderPrice1 ) or
     ( Sender = EDT_GiftOrderPriceMax1 ) then
  begin
    aIndex := 1;
    aEnabled :=
      ( EDT_GiftOrderPrice1.Text <> EmptyStr ) or
      ( EDT_GiftOrderPriceMax1.Text <> EmptyStr );
  end else
  if ( Sender = EDT_GiftOrderPrice2 ) or
     ( Sender = EDT_GiftOrderPriceMax2 ) then
  begin
    aIndex := 2;
    aEnabled :=
      ( EDT_GiftOrderPrice2.Text <> EmptyStr ) or
      ( EDT_GiftOrderPriceMax2.Text <> EmptyStr );
  end else
  if ( Sender = EDT_GiftOrderPrice3 ) or
     ( Sender = EDT_GiftOrderPriceMax3 ) then
  begin
    aIndex := 3;
    aEnabled :=
     ( EDT_GiftOrderPrice3.Text <> EmptyStr ) or
      ( EDT_GiftOrderPriceMax3.Text <> EmptyStr );
  end else
  if ( Sender = EDT_GiftOrderPrice4 ) or
     ( Sender = EDT_GiftOrderPriceMax4 ) then
  begin
    aIndex := 4;
    aEnabled :=
      ( EDT_GiftOrderPrice4.Text <> EmptyStr ) or
      ( EDT_GiftOrderPriceMax4.Text <> EmptyStr );
  end else
  if ( Sender = EDT_GiftOrderPrice5 ) or
     ( Sender = EDT_GiftOrderPriceMax5 ) then
  begin
    aIndex := 5;
    aEnabled :=
      ( EDT_GiftOrderPrice5.Text <> EmptyStr ) or
      ( EDT_GiftOrderPriceMax5.Text <> EmptyStr );
  end else
  if ( Sender = EDT_GiftOrderPrice6 ) or
     ( Sender = EDT_GiftOrderPriceMax6 ) then
  begin
    aIndex := 6;
    aEnabled :=
      ( EDT_GiftOrderPrice6.Text <> EmptyStr ) or
      ( EDT_GiftOrderPriceMax6.Text <> EmptyStr );
  end;
  {}
  if aIndex in [1..6] then
    SetEnableGiftControl( aIndex, aEnabled );
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmGift.EDT_GiftOrderPrice1Exit(Sender: TObject);
begin
  if not CheckPriceMinMax( StrToIntDef( TEdit( Sender ).Text, 0 ) ) then
  Abort;
  if ( Sender = EDT_GiftOrderPrice1 ) or
     ( Sender = EDT_GiftOrderPriceMax1 ) then
  begin
    FollowPrice( EDT_GiftOrderPriceMax1,  EDT_GiftOrderPrice1.Text );
  end else
  if ( Sender = EDT_GiftOrderPrice2 ) or
     ( Sender = EDT_GiftOrderPriceMax2 ) then
  begin
    FollowPrice( EDT_GiftOrderPriceMax2,  EDT_GiftOrderPrice2.Text );
  end else
  if ( Sender = EDT_GiftOrderPrice3 ) or
     ( Sender = EDT_GiftOrderPriceMax3 ) then
  begin
    FollowPrice( EDT_GiftOrderPriceMax3,  EDT_GiftOrderPrice3.Text );
  end else
  if ( Sender = EDT_GiftOrderPrice4 ) or
     ( Sender = EDT_GiftOrderPriceMax4 ) then
  begin
    FollowPrice( EDT_GiftOrderPriceMax4,  EDT_GiftOrderPrice4.Text );
  end else
  if ( Sender = EDT_GiftOrderPrice5 ) or
     ( Sender = EDT_GiftOrderPriceMax5 ) then
  begin
    FollowPrice( EDT_GiftOrderPriceMax5,  EDT_GiftOrderPrice5.Text );
  end else
  if ( Sender = EDT_GiftOrderPrice6 ) or
     ( Sender = EDT_GiftOrderPriceMax6 ) then
  begin
    FollowPrice( EDT_GiftOrderPriceMax6,  EDT_GiftOrderPrice6.Text );
  end;
  {}
end;

{ ---------------------------------------------------------------------------- }

function TfrmGift.CheckData: Boolean;
var
  aIndex, aLastIdx: Integer;
  aAllEmpty: Boolean;
  aStVal, aEdVal, aLastVal: String;
  aGiftCount, aArticleNoCount: Integer;
  aEdit: TEdit;
  aComBox:  TComboBox;
begin
  Result := False;
  aAllEmpty :=
    ( EDT_GiftOrderPrice1.Text = EmptyStr ) and
    ( EDT_GiftOrderPriceMax1.Text = EmptyStr ) and
    ( EDT_GiftOrderPrice2.Text = EmptyStr ) and
    ( EDT_GiftOrderPriceMax2.Text = EmptyStr ) and
    ( EDT_GiftOrderPrice3.Text = EmptyStr ) and
    ( EDT_GiftOrderPriceMax3.Text = EmptyStr ) and
    ( EDT_GiftOrderPrice4.Text = EmptyStr ) and
    ( EDT_GiftOrderPriceMax4.Text = EmptyStr ) and
    ( EDT_GiftOrderPrice5.Text = EmptyStr ) and
    ( EDT_GiftOrderPriceMax5.Text = EmptyStr ) and
    ( EDT_GiftOrderPrice6.Text = EmptyStr ) and
    ( EDT_GiftOrderPriceMax6.Text = EmptyStr );
  if aAllEmpty then
  begin
    WarningMsg( '已指定贈品, 請設定贈品內容(至少指定一項)。' );
    Exit;
  end;
  {}
  for aIndex := 1 to 6 do
  begin
    aStVal := TEdit( Self.FindComponent( Format( 'EDT_GiftOrderPrice%d', [aIndex] ) ) ).Text;
    aEdVal := TEdit( Self.FindComponent( Format( 'EDT_GiftOrderPriceMax%d', [aIndex] ) ) ).Text;
    if ( ( aStVal <> EmptyStr ) and ( aEdVal = EmptyStr ) ) or
       ( ( aStVal = EmptyStr ) and ( aEdVal <> EmptyStr ) ) then
    begin
      WarningMsg( '贈品設定, 訂購金額【起】【迄】必須輸入。' );
      aEdit := TEdit( Self.FindComponent( Format( 'EDT_GiftOrderPrice%d', [aIndex] ) ) );
      if TEdit( Self.FindComponent( Format( 'EDT_GiftOrderPriceMax%d', [aIndex] ) ) ).Text = EmptyStr then
        aEdit := TEdit( Self.FindComponent( Format( 'EDT_GiftOrderPriceMax%d', [aIndex] ) ) );
      if aEdit.CanFocus then aEdit.SetFocus;
      Exit;
    end;
  end;
  {}
  aLastIdx := MaxInt;
  for aIndex := 1 to 6 do
  begin
    aStVal := TEdit( Self.FindComponent( Format( 'EDT_GiftOrderPrice%d', [aIndex] ) ) ).Text;
    aEdVal := TEdit( Self.FindComponent( Format( 'EDT_GiftOrderPriceMax%d', [aIndex] ) ) ).Text;
    if ( aStVal = EmptyStr ) then
      aLastIdx := aIndex
    else  begin
      if ( aIndex > aLastIdx ) then
      begin
        WarningMsg( '贈品設定, 必須依照順序一～六輸入, 不可跳階。' );
        Exit;
      end;
    end;
  end;
  {}
  for aIndex := 1 to 6 do
  begin
    aStVal := TEdit( Self.FindComponent( Format( 'EDT_GiftOrderPrice%d', [aIndex] ) ) ).Text;
    aEdVal := TEdit( Self.FindComponent( Format( 'EDT_GiftOrderPriceMax%d', [aIndex] ) ) ).Text;
    if ( aStVal <> EmptyStr ) then
    begin
      if ( StrToFloatDef( aEdVal, 0 ) ) < ( StrToFloatDef( aStVal, 0 ) ) then
      begin
        WarningMsg( Format( '贈品設定, 贈品%d 訂購金額【迄】必須大於或等於【起】。', [aIndex] ) );
        aEdit := TEdit( Self.FindComponent( Format( 'EDT_GiftOrderPriceMax%d', [aIndex] ) ) );
        if aEdit.CanFocus then aEdit.SetFocus;
        Exit;
      end;
      if ( aIndex > 1 ) then
      begin
        if ( StrToFloatDef( aLastVal, 0 ) + 1 ) <> ( StrToFloatDef( aStVal, 0 ) ) then
        begin
          WarningMsg( Format( '贈品設定, 贈品%d 訂購金額一～六必須按照大小排列, 起迄金額不可重疊。', [aIndex] ) );
          aEdit := TEdit( Self.FindComponent( Format( 'EDT_GiftOrderPrice%d', [aIndex] ) ) );
          if aEdit.CanFocus then aEdit.SetFocus;
          Exit;
        end;
      end;
      aLastVal := aEdVal;
    end;
  end;
  {}
  for aIndex := 1 to 6 do
  begin
    aStVal := TEdit( Self.FindComponent( Format( 'EDT_GiftOrderPrice%d', [aIndex] ) ) ).Text;
    aEdVal := TEdit( Self.FindComponent( Format( 'EDT_GiftOrderPriceMax%d', [aIndex] ) ) ).Text;
    if ( aStVal <> EmptyStr ) then
    begin
      if TComboBox( Self.FindComponent( Format( 'EDT_GiftType%d', [aIndex] ) ) ).Text = EmptyStr then
      begin
        WarningMsg( Format( '贈品設定, 贈品%d 已設定訂購金額，但沒有設定【贈送方式】。', [aIndex] ) );
        aComBox := TComboBox( Self.FindComponent( Format( 'EDT_GiftType%d', [aIndex] ) ) );
        if aComBox.CanFocus then aComBox.SetFocus;
        Exit;
      end;
      if TComboBox( Self.FindComponent( Format( 'EDT_GiftKind%d', [aIndex] ) ) ).Text = EmptyStr then
      begin
        WarningMsg( Format( '贈品設定, 贈品%d 已設定訂購金額，但沒有設定【贈送項目】。', [aIndex] ) );
        aComBox := TComboBox( Self.FindComponent( Format( 'EDT_GiftKind%d', [aIndex] ) ) );
        if aComBox.CanFocus then aComBox.SetFocus;
        Exit;
      end;
      aArticleNoCount := TListBox( Self.FindComponent( Format( 'EDT_ArticleNo%d', [aIndex] ) ) ).Items.Count;
      aGiftCount := GiftCount( TComboBox( Self.FindComponent( Format( 'EDT_GiftType%d', [aIndex] ) ) ).ItemIndex );
      if ( ( aGiftCount in [1..9] ) and ( aGiftCount > aArticleNoCount ) ) then
      begin
        WarningMsg( Format( '贈品設定, 贈品%d 至少須選定%d組。', [aIndex, aGiftCount] ) );
        Exit;
      end else
      if ( aArticleNoCount = 0 ) then
      begin
        WarningMsg( Format( '贈品設定, 贈品%d 至少須選定1組。', [aIndex] ) );
        Exit;
      end;
    end;
  end;
  {}
  Result := True;
end;

{ ---------------------------------------------------------------------------- }

function TfrmGift.CheckPriceMinMax(const aValue: Integer): Boolean;
begin
  Result := ( aValue >=0 ) or ( aValue <= 99999999 );
end;

{ ---------------------------------------------------------------------------- }

function TfrmGift.GetGiftKindStr(aControl: TComboBox): String;
begin
  Result := EmptyStr;
  if Assigned( aControl ) then
  begin
    if ( aControl.ItemIndex >= 0  ) then
     Result := IntToStr( aControl.ItemIndex + 1 );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmGift.GetGiftTypeStr(aControl: TComboBox): String;
begin
  Result := EmptyStr;
  if Assigned( aControl ) then
  begin
    if ( aControl.ItemIndex >= 0  ) then
     Result := IntToStr( aControl.ItemIndex );
  end;
end;

{ ---------------------------------------------------------------------------- }

function TfrmGift.GetArticleNoStr(aControl: TListBox): String;
var
  aIndex: Integer;
begin
  Result := EmptyStr;
  if Assigned( aControl ) then
  begin
    for aIndex := 0 to aControl.Items.Count - 1 do
      Result := ( Result + aControl.Items[aIndex] + ',' );
    if IsDelimiter( ',', Result, Length( Result ) ) then
      Delete( Result, Length( Result ), 1 );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmGift.actSaveExecute(Sender: TObject);
begin
  if not CheckData then Exit;
  EditorToData;
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmGift.actCancelExecute(Sender: TObject);
begin
  Self.Close;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmGift.Button1Click(Sender: TObject);
begin
  if ( Sender = Button1 ) then
    ShowMultiForm( 1 )
  else
  if ( Sender = Button2 ) then
    ShowMultiForm( 2 )
  else
  if ( Sender = Button3 ) then
    ShowMultiForm( 3 )
  else
  if ( Sender = Button4 ) then
    ShowMultiForm( 4 )
  else
  if ( Sender = Button5 ) then
    ShowMultiForm( 5 )
  else
  if ( Sender = Button6 ) then
    ShowMultiForm( 6 )
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmGift.X_ArticleNoStr1Click(Sender: TObject);
begin
  if ( Sender = X_ArticleNoStr1 ) then
    ClearArticleNoStr( EDT_ArticleNoStr1, EDT_ArticleNo1 )
  else
  if ( Sender = X_ArticleNoStr2 ) then
    ClearArticleNoStr( EDT_ArticleNoStr2, EDT_ArticleNo2 )
  else
  if ( Sender = X_ArticleNoStr3 ) then
    ClearArticleNoStr( EDT_ArticleNoStr3, EDT_ArticleNo3 )
  else
  if ( Sender = X_ArticleNoStr4 ) then
    ClearArticleNoStr( EDT_ArticleNoStr4, EDT_ArticleNo4 )
  else
  if ( Sender = X_ArticleNoStr5 ) then
    ClearArticleNoStr( EDT_ArticleNoStr5, EDT_ArticleNo5 )
  else
  if ( Sender = X_ArticleNoStr6 ) then
    ClearArticleNoStr( EDT_ArticleNoStr6, EDT_ArticleNo6 )
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmGift.EDT_GiftKind1Change(Sender: TObject);
begin
  if ( Sender = EDT_GiftKind1 ) then
    ClearArticleNoStr( EDT_ArticleNoStr1, EDT_ArticleNo1 )
  else
  if ( Sender = EDT_GiftKind2 ) then
    ClearArticleNoStr( EDT_ArticleNoStr2, EDT_ArticleNo2 )
  else
  if ( Sender = EDT_GiftKind3 ) then
    ClearArticleNoStr( EDT_ArticleNoStr3, EDT_ArticleNo3 )
  else
  if ( Sender = EDT_GiftKind4 ) then
    ClearArticleNoStr( EDT_ArticleNoStr4, EDT_ArticleNo4 )
  else
  if ( Sender = EDT_GiftKind5 ) then
    ClearArticleNoStr( EDT_ArticleNoStr5, EDT_ArticleNo5 )
  else
  if ( Sender = EDT_GiftKind6 ) then
    ClearArticleNoStr( EDT_ArticleNoStr6, EDT_ArticleNo6 )
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmGift.EDT_GiftOrderPrice1KeyPress(Sender: TObject; var Key: Char);
begin
  if ( Ord( Key ) in [48..57] ) or ( Ord( Key ) = VK_BACK ) then Exit;
  Key := #0;
end;

{ ---------------------------------------------------------------------------- }

procedure TfrmGift.EDT_GiftOrderPriceMax1Exit(Sender: TObject);
begin
  if not CheckPriceMinMax( StrToIntDef( TEdit( Sender ).Text, 0 ) ) then
  Abort;
end;

{ ---------------------------------------------------------------------------- }

end.
