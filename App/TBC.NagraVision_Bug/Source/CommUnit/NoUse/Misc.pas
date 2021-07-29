unit Misc;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Buttons, ComCtrls, StdCtrls, ExtCtrls, DB, DBClient,
  { Developer Express http://www.devexpress.com }
  { Common Unit}
  cxClasses, cxStyles, cxCustomData, cxGraphics, cxFilter, cxData, cxDBData,
  cxControls, cxContainer, cxLookAndFeels, cxDataStorage, cxFormats,
  { NavBar }
  dxNavBarCollns, dxNavBarBase, dxNavBar,
  { ExpressGrid }
  cxGridLevel, cxGridCustomView, cxGridCustomTableView, cxGridTableView,
  cxGridDBTableView, cxGridBandedTableView, cxGridDBBandedTableView, cxGrid,
  { Editor }
  cxEdit, cxTextEdit, cxMaskEdit, cxMemo, cxCalendar, cxButtonEdit, cxCheckBox,
  cxDropDownEdit, cxImageComboBox, cxSpinEdit, cxCalc, cxHyperLinkEdit,
  cxCurrencyEdit, cxImage, cxRadioGroup, cxPC,
  cxLookupEdit, cxDBLookupEdit, cxDBLookupComboBox;


function GetItemText(AControl: TControl; APartValue: String): String;
function GetItemValue(AItemText: String): String;
function GetControltValue(AControl: TControl): String;
function GetFormActiveControl(AForm: TForm): TWinControl;
function IsReadonlyControl(AObject: TObject): Boolean;
function CheckNotNull(AForm: TForm): Boolean;
function TrimDateMask(AValue: String): String;
function TrimCurrencyMask(AValue: String): String;
function CalcUploadFtpFullPath(ARoot: String; const AID: String; const ALen: Integer): String;
function CalcUploadFtpOuterPath(const AID: String; const ALen: Integer): String;
function CalcUploadFtpInnerPath(const AID: String; const ALen: Integer): String;
procedure ValueToControl(AValue: Variant; AControl: TControl);
procedure ControlSetReadOnly(const AReadOnly: Boolean; AControl: TControl);
procedure EditorSetFocused(Sender: TObject);
procedure ClearEditor(AList: TStringList);
procedure DataSetToEdit(ACDS: TClientDataSet; AList: TStringList);
procedure PopupDateWindow(Sender: TcxCustomTextEdit);
procedure ResetGridSort(AGridView: TcxCustomGridView);


implementation


uses ApConst, ApUtilis, DateWindow;

const AItemDot = '.';


{ ---------------------------------------------------------------------------- }

procedure InternalTypeCastClear(AObject: TObject);
begin
  if AObject is TLabel then
    TLabel( AObject ).Caption := ''
  else if AObject is TcxTextEdit then
    TcxTextEdit( AObject ).Clear
  else if AObject is TcxMaskEdit then
    TcxMaskEdit( AObject ).Clear
  else if AObject is TcxMemo then
    TcxMemo( AObject ).Clear
  else if AObject is TcxButtonEdit then
    TcxButtonEdit( AObject ).Clear
  else if AObject is TcxCheckBox then
    TcxCheckBox( AObject ).Checked := False
  else if AObject is TcxComboBox then
    TcxComboBox( AObject ).ItemIndex := -1
  else if AObject is TcxImageComboBox then
    TcxImageComboBox( AObject ).ItemIndex := -1
  else if AObject is TcxSpinEdit then
    TcxSpinEdit( AObject ).Clear
  else if AObject is TcxCalcEdit then
    TcxCalcEdit( AObject ).Clear
  else if AObject is TcxHyperLinkEdit then
    TcxHyperLinkEdit( AObject ).Clear
  else if AObject is TcxCurrencyEdit then
    TcxCurrencyEdit( AObject ).Clear
  else if AObject is TcxImage then
     TcxImage( AObject ).Clear
  else if AObject is TRadioButton then
    TRadioButton( AObject ).Checked := False
  else if AObject is TRadioGroup then
    TRadioGroup( AObject ).ItemIndex := -1
  else if AObject is TcxPopupEdit then
    TcxPopupEdit( AObject ).Text := '';
end;

{ ---------------------------------------------------------------------------- }

procedure InternalDataSetValueToControl(AValue: Variant; AControl: TControl);
var
  AOldReadOnly: Boolean;
begin
  if not Assigned( AControl ) then Exit;
  AValue := VarToStr( AValue );
  if AControl is TLabel then
    TLabel( AControl ).Caption := AValue
  else if AControl is TcxTextEdit then
  begin
    if TcxTextEdit( AControl ).Properties.LookupItems.Count > 0 then
      TcxTextEdit( AControl ).Text := GetItemText( AControl, AValue )
    else
      TcxTextEdit( AControl ).Text := AValue;
  end
  else if AControl is TcxButtonEdit then
  begin
    if TcxButtonEdit( AControl ).Properties.LookupItems.Count > 0 then
      TcxButtonEdit( AControl ).Text := GetItemText( AControl, AValue )
    else
      TcxButtonEdit( AControl ).Text := AValue
  end
  else if AControl is TcxPopupEdit then
    TcxPopupEdit( AControl ).Text := AValue
  else if AControl is TcxMaskEdit then
    TcxMaskEdit( AControl ).Text := AValue
  else if AControl is TcxComboBox then
    TcxComboBox( AControl ).Text := GetItemText( AControl, AValue )
  else if AControl is TcxImageComboBox then
  begin
    AOldReadOnly := TcxImageComboBox( AControl ).Properties.ReadOnly;
    TcxImageComboBox( AControl ).Properties.ReadOnly := False;
    TcxImageComboBox( AControl ).ItemIndex :=
      StrToInt( GetItemText( AControl, AValue ) );
    TcxImageComboBox( AControl ).Properties.ReadOnly := AOldReadOnly;
  end
  else if AControl is TcxMemo then
    TcxMemo( AControl ).Text := AValue
  else if AControl is TcxCheckBox then
    TcxCheckBox( AControl ).EditValue := AValue
  else if AControl is TcxSpinEdit then
    TcxSpinEdit( AControl ).EditValue := Nvl( AValue, 0 )
  else if AControl is TcxCurrencyEdit then
    TcxCurrencyEdit( AControl ).EditValue := AValue
  else if AControl is TcxHyperLinkEdit then
    TcxHyperLinkEdit( AControl ).Text := AValue
  else if AControl is TcxLookupComboBox then
    TcxLookupComboBox( AControl ).Text := AValue;
end;

{ ---------------------------------------------------------------------------- }

function GetItemText(AControl: TControl; APartValue: String): String;
var
  AcxComboBox: TcxComboBox;
  AcxTextEdit: TcxTextEdit;
  AcxButtonEdit: TcxButtonEdit;
  AcxImageComboBox: TcxImageComboBox;
  AItemText: String;
  AIndex: Integer;
begin
  Result := APartValue;
  if AControl is TcxComboBox then
  begin
    AcxComboBox := TcxComboBox( AControl );
    for AIndex := 0 to AcxComboBox.Properties.Items.Count - 1 do
    begin
      AItemText := AcxComboBox.Properties.Items.Strings[AIndex];
      if ( AItemText = '' ) or ( Pos( AItemDot, AItemText ) = 0 ) then
        Continue;
      if APartValue =
        Copy( AItemText, 1, Pos( AItemDot, AItemText ) - 1 ) then
      begin
        Result := AItemText;
        Break;
      end;
    end;
    if AcxComboBox.Properties.Items.IndexOf( Result ) < 0 then Result := '';
  end
  else if AControl is TcxTextEdit then
  begin
    AcxTextEdit := TcxTextEdit( AControl );
    for AIndex := 0 to AcxTextEdit.Properties.LookupItems.Count - 1 do
    begin
      AItemText := AcxTextEdit.Properties.LookupItems.Strings[AIndex];
      if ( AItemText = '' ) or ( Pos( AItemDot, AItemText ) = 0 ) then
        Continue;
      if APartValue =
        Copy( AItemText, 1, Pos( AItemDot, AItemText ) - 1 ) then
      begin
        Result := AItemText;
        Break;
      end;
    end;
  end
  else if AControl is TcxButtonEdit then
  begin
    AcxButtonEdit := TcxButtonEdit( AControl );
    for AIndex := 0 to AcxButtonEdit.Properties.LookupItems.Count - 1 do
    begin
      AItemText := AcxButtonEdit.Properties.LookupItems.Strings[AIndex];
      if ( AItemText = '' ) or ( Pos( AItemDot, AItemText ) = 0 ) then
        Continue;
      if APartValue =
        Copy( AItemText, 1, Pos( AItemDot, AItemText ) - 1 ) then
      begin
        Result := AItemText;
        Break;
      end;
    end;
  end
  else if AControl is TcxImageComboBox then
  begin
    Result := '-1';
    AcxImageComboBox := TcxImageComboBox( AControl );
    for AIndex := 0 to AcxImageComboBox.Properties.Items.Count - 1 do
    begin
      if APartValue = AcxImageComboBox.Properties.Items[AIndex].Value then
      begin
        Result := IntToStr( AIndex );
        Break;
      end;
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

function GetItemValue(AItemText: String): String;
var
  AIndex: Integer;
begin
  AIndex := AnsiPos( AItemDot, AItemText );
  Result := AItemText;
  if AIndex > 1 then Result := Copy( AItemText, 1, AIndex - 1 );
end;

{ ---------------------------------------------------------------------------- }

function GetControltValue(AControl: TControl): String;
begin
  Result := '';
  if AControl is TcxTextEdit then
  begin
    if TcxTextEdit( AControl ).Properties.LookupItems.Count > 0 then
      Result := GetItemValue( TcxTextEdit( AControl ).Text )
    else
      Result := Trim( TcxTextEdit( AControl ).Text )
  end
  else if AControl is TcxButtonEdit then
  begin
    if TcxButtonEdit( AControl ).Properties.EditMask <> '' then
      Result := TrimDateMask( TcxButtonEdit( AControl ).Text )
    else if TcxButtonEdit( AControl ).Properties.LookupItems.Count > 0 then
      Result := GetItemValue( TcxButtonEdit( AControl ).Text )
    else
      Result := TcxButtonEdit( AControl ).Text;
  end
  else if AControl is TcxMaskEdit then
    Result := Trim( TcxMaskEdit( AControl ).Text )
  else if AControl is TcxCheckBox then
    Result := TcxCheckBox( AControl ).EditValue
  else if AControl is TcxMemo then
    Result := StringReplace( TcxMemo( AControl ).Lines.Text, #13#10, #8, [rfReplaceAll] )
  else if AControl is TcxComboBox then
    Result := GetItemValue( TcxComboBox( AControl ).Text )
  else if AControl is TcxSpinEdit then
    Result := VarToStr( TcxSpinEdit( AControl ).Value )
  else if AControl is TcxCurrencyEdit then
    Result := TrimCurrencyMask( TcxCurrencyEdit( AControl ).Text )
  else if AControl is TcxHyperLinkEdit then
    Result := TcxHyperLinkEdit( AControl ).Text
  else if AControl is TcxLookupComboBox then
    Result := TcxLookupComboBox( AControl ).Text;

end;

{ ---------------------------------------------------------------------------- }

function GetFormActiveControl(AForm: TForm): TWinControl;
var
  AIndex: Integer;
  AComp: TComponent;
begin
  Result := nil;
  if not Assigned( AForm ) then Exit;
  for AIndex := 0 to AForm.ComponentCount - 1 do
  begin
    AComp := AForm.Components[AIndex];
    if not ( AComp is TWinControl ) then Continue;
    if ( TWinControl( AComp ).Focused ) then
    begin
      Result := TWinControl( AComp );
      Break;
    end;  
  end;
end;

{ ---------------------------------------------------------------------------- }

function IsReadonlyControl(AObject: TObject): Boolean;
begin
  Result := False;
end;

{ ---------------------------------------------------------------------------- }

function CheckNotNull(AForm: TForm): Boolean;
type
  TValidateType = ( vdString, vdNumeric, vdBoolean );
var
  AIndex: Integer;
  AStringList: TStringList;
  ALabel: TLabel;
  AWinControl: TWinControl;
  AValue, AText: String;

    { ---------------------------------------------------------- }

    function GetVaidateType: TValidateType;
    begin
      Result := vdString;
      if ( AWinControl is TcxSpinEdit ) or
         ( AWinControl is TcxCurrencyEdit ) then
        Result := vdNumeric
      else if ( AWinControl is TcxCheckBox ) then
        Result := vdBoolean;
    end;

    { -----------------------------------------------------------}

    procedure LocateTabSheet;
    var
       AParent: TWinControl;
       AEvent: TNotifyEvent;
    begin
      AParent := AWinControl.Parent;
      while Assigned( AParent ) do
      begin
        if AParent is TcxTabSheet then
        begin
          AEvent := TcxTabSheet( AParent ).PageControl.OnChange;
          TcxTabSheet( AParent ).PageControl.OnChange := nil;
          try
            TcxTabSheet( AParent ).PageControl.ActivePage :=
              TcxTabSheet( AParent );
          finally
            TcxTabSheet( AParent ).PageControl.OnChange := AEvent;
          end;  
          Break;
        end;
        AParent := AParent.Parent;
      end;
    end;

    { -----------------------------------------------------------}

begin
  AStringList := TStringList.Create;
  try
    for AIndex := 0 to AForm.ComponentCount - 1 do
    begin
      if not ( AForm.Components[AIndex] is TLabel ) then Continue;
      ALabel := TLabel( AForm.Components[AIndex] );
      if ( ALabel.Visible ) and ( ALabel.Enabled ) and
         ( ALabel.Font.Color = clMaroon ) and ( ALabel.FocusControl <> nil ) then
      begin
        AWinControl := ALabel.FocusControl;
        if AWinControl is TcxTextEdit then
          AValue := Trim( TcxTextEdit( AWinControl ).Text )
        else if AWinControl is TcxButtonEdit then
          AValue := TrimDateMask( Trim( TcxButtonEdit( AWinControl ).Text ) )
        else if AWinControl is TcxMaskEdit then
          AValue := TrimDateMask( Trim( TcxMaskEdit( AWinControl ).Text ) )
        else if AWinControl is TcxComboBox then
          AValue := Trim( TcxComboBox( AWinControl ).Text )
        else if AWinControl is TcxCheckBox then
          AValue := Trim( VarToStr( TcxCheckBox( AWinControl ).EditValue ) )
        else if AWinControl is TcxMemo then
          AValue := Trim( TcxMemo( AWinControl ).Lines.Text )
        else if AWinControl is TcxSpinEdit then
          AValue := Trim( TcxSpinEdit( AWinControl ).EditingText )
        else if AWinControl is TcxCurrencyEdit then
          AValue := TrimCurrencyMask( TcxCurrencyEdit( AWinControl ).EditingText )
        else if AWinControl is TcxHyperLinkEdit then
          AValue := Trim( TcxHyperLinkEdit( AWinControl ).Text );
        AText := '';
        case GetVaidateType of
          vdString:
             begin
               if AValue = '' then
                 AText := Copy( ALabel.Caption, 1,
                   Length( ALabel.Caption ) - 1 ) + ', 必須輸入!';
             end;
          vdNumeric:
             begin
               if ( AValue = '' ) or ( AValue = '0' ) then
                 AText := Copy( ALabel.Caption, 1,
                   Length( ALabel.Caption ) - 1 ) + ', 輸入必須大於零!'
             end;
          vdBoolean:
             begin
               if ( AValue = '' ) or ( AValue = 'N' ) then
                 AText := Copy( ALabel.Caption, 1,
                   Length( ALabel.Caption ) - 1 ) + ', 必須選取!'
             end;
        end;
        if AText <> '' then AStringList.AddObject( AText, AWinControl );
      end;
    end;
    Result := ( AStringList.Count = 0 );
    if not Result then
    begin
      for AIndex := 0 to AStringList.Count - 1 do
      begin
        AWinControl := TWinControl( AStringList.Objects[AIndex] );
        LocateTabSheet;
        if ( not AWinControl.Focused ) and AWinControl.CanFocus then
        begin
          AWinControl.SetFocus;
          Break;
        end;
      end;
      WarningMsg( AStringList.Text );
    end;
  finally
    AStringList.Free;
  end;
end;

{ ---------------------------------------------------------------------------- }

function TrimDateMask(AValue: String): String;
const
  ADateMask: array [1..2] of String = ( '_', '/' );
var
  AIndex: Integer;
begin
  if AValue <> '' then
  begin
    for AIndex := Low( ADateMask ) to High( ADateMask ) do
      AValue := TrimString( AValue, ADateMask[AIndex] );
  end;
  Result := AValue;
end;

{ ---------------------------------------------------------------------------- }

function TrimCurrencyMask(AValue: String): String;
const
  ACurrencyMask: array [1..5] of String = ( 'NT', 'nt', '$', ',', ' ' );
var
  AIndex: Integer;
begin
  if AValue <> '' then
  begin
    for AIndex := Low( ACurrencyMask ) to High( ACurrencyMask ) do
      AValue := TrimString( AValue, ACurrencyMask[AIndex] );
  end;
  Result := AValue;
end;

{ ---------------------------------------------------------------------------- }

procedure ValueToControl(AValue: Variant; AControl: TControl);
var
  AOldReadOnly: Boolean;
begin
  if not Assigned( AControl ) then Exit;
  AValue := VarToStr( AValue );
  if AControl is TLabel then
    TLabel( AControl ).Caption := AValue
  else if AControl is TcxTextEdit then
  begin
    if TcxTextEdit( AControl ).Properties.LookupItems.Count > 0 then
      TcxTextEdit( AControl ).Text := GetItemText( AControl, AValue )
    else
      TcxTextEdit( AControl ).Text := AValue;
  end
  else if AControl is TcxButtonEdit then
  begin
    if TcxButtonEdit( AControl ).Properties.LookupItems.Count > 0 then
      TcxButtonEdit( AControl ).Text := GetItemText( AControl, AValue )
    else
      TcxButtonEdit( AControl ).Text := AValue
  end
  else if AControl is TcxMaskEdit then
    TcxMaskEdit( AControl ).Text := AValue
  else if AControl is TcxComboBox then
    TcxComboBox( AControl ).Text := GetItemText( AControl, AValue )
  else if AControl is TcxImageComboBox then
  begin
    AOldReadOnly := TcxImageComboBox( AControl ).Properties.ReadOnly;
    TcxImageComboBox( AControl ).Properties.ReadOnly := False;
    TcxImageComboBox( AControl ).ItemIndex :=
      StrToInt( GetItemText( AControl, AValue ) );
    TcxImageComboBox( AControl ).Properties.ReadOnly := AOldReadOnly;  
  end
  else if AControl is TcxMemo then
    TcxMemo( AControl ).Text := AValue
  else if AControl is TcxCheckBox then
    TcxCheckBox( AControl ).EditValue := AValue
  else if AControl is TcxSpinEdit then
    TcxSpinEdit( AControl ).EditValue := Nvl( AValue, 0 )
  else if AControl is TcxCurrencyEdit then
    TcxCurrencyEdit( AControl ).EditValue := AValue
  else if AControl is TcxHyperLinkEdit then
    TcxHyperLinkEdit( AControl ).Text := AValue;
end;

{ ---------------------------------------------------------------------------- }

procedure ControlSetReadOnly(const AReadOnly: Boolean; AControl: TControl);
begin
  if not Assigned( AControl ) then Exit;
  if AControl is TcxTextEdit then
  begin
    TcxTextEdit( AControl ).Style.Color := EDIT_COLOR;
    TcxTextEdit( AControl ).Properties.ReadOnly := AReadOnly;
    if AReadOnly then
      TcxTextEdit( AControl ).Style.Color := NONE_COLOR;
  end
  else if AControl is TcxButtonEdit then
  begin
    TcxButtonEdit( AControl ).Style.Color := EDIT_COLOR;
    TcxButtonEdit( AControl ).Properties.ReadOnly := AReadOnly;
    if AReadOnly then
      TcxButtonEdit( AControl ).Style.Color := NONE_COLOR;
  end
  else if AControl is TcxMaskEdit then
  begin
    TcxMaskEdit( AControl ).Style.Color := EDIT_COLOR;
    TcxMaskEdit( AControl ).Properties.ReadOnly := AReadOnly;
    if AReadOnly then
      TcxMaskEdit( AControl ).Style.Color := NONE_COLOR;
  end
  else if AControl is TcxComboBox then
  begin
    TcxComboBox( AControl ).Style.Color := EDIT_COLOR;
    TcxComboBox( AControl ).Properties.ReadOnly := AReadOnly;
    if AReadOnly then
      TcxComboBox( AControl ).Style.Color := NONE_COLOR;
  end
  else if AControl is TcxImageComboBox then
  begin
    TcxImageComboBox( AControl ).Style.Color := EDIT_COLOR;
    TcxImageComboBox( AControl ).Properties.ReadOnly := AReadOnly;
    if AReadOnly then
      TcxImageComboBox( AControl ).Style.Color := NONE_COLOR;
  end
  else if AControl is TcxMemo then
  begin
    TcxMemo( AControl ).Style.Color := EDIT_COLOR;
    TcxMemo( AControl ).Properties.ReadOnly := AReadOnly;
    if AReadOnly then
      TcxMemo( AControl ).Style.Color := NONE_COLOR;
  end  
  else if AControl is TcxCheckBox then
  begin
    TcxCheckBox( AControl ).Style.Font.Color := clWindowText;
    TcxCheckBox( AControl ).Properties.ReadOnly := AReadOnly;
    if AReadOnly then
      TcxMemo( AControl ).Style.Font.Color := clGrayText;
  end  
  else if AControl is TcxSpinEdit then
  begin
    TcxSpinEdit( AControl ).Style.Color := EDIT_COLOR;
    TcxSpinEdit( AControl ).Properties.ReadOnly := AReadOnly;
    if AReadOnly then
      TcxSpinEdit( AControl ).Style.Color := NONE_COLOR;
  end
  else if AControl is TcxCurrencyEdit then
  begin
    TcxCurrencyEdit( AControl ).Style.Color := EDIT_COLOR;
    TcxCurrencyEdit( AControl ).Properties.ReadOnly := AReadOnly;
    if AReadOnly then
      TcxCurrencyEdit( AControl ).Style.Color := NONE_COLOR;
  end
  else if AControl is TcxHyperLinkEdit then
  begin
    TcxHyperLinkEdit( AControl ).Style.Color := EDIT_COLOR;
    TcxHyperLinkEdit( AControl ).Properties.ReadOnly := AReadOnly;
    if AReadOnly then
      TcxHyperLinkEdit( AControl ).Style.Color := NONE_COLOR;
  end  
end;

{ ---------------------------------------------------------------------------- }

procedure EditorSetFocused(Sender: TObject);
begin
  if Sender is TcxCustomEdit then
  begin
    if TcxPublicProperty( Sender ).CanFocusEx then
      TcxPublicProperty( Sender ).SetFocus;
    TcxPublicProperty( Sender ).SelectAll;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure ClearEditor(AList: TStringList);
var
  AIndex: Integer;
begin
  for AIndex := 0 to AList.Count - 1 do
  begin
    if Assigned( AList.Objects[AIndex] ) then
    begin
      if not IsReadonlyControl( AList.Objects[AIndex] ) then
        InternalTypeCastClear( AList.Objects[AIndex] );
    end;
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure DataSetToEdit(ACDS: TClientDataSet; AList: TStringList);
var
  AIndex, AObjIdx: Integer;
  AValue: Variant;
  AControl: TControl;
begin
  for AIndex := 0 to ACDS.FieldCount - 1 do
  begin
    AControl := nil;
    AValue := ACDS.Fields[AIndex].Value;
    AObjIdx := AList.IndexOf( ACDS.Fields[AIndex].FieldName );
    if AObjIdx in [0..AList.Count - 1] then
      AControl := TControl( AList.Objects[AObjIdx] );
    if Assigned( AControl ) then
      InternalDataSetValueToControl( AValue, AControl );
  end;
end;

{ ---------------------------------------------------------------------------- }

procedure PopupDateWindow(Sender: TcxCustomTextEdit);
begin
  frmDateWindow := TfrmDateWindow.Create( nil );
  frmDateWindow.ReturnControl := Sender;
  frmDateWindow.Show;
end;

{ ---------------------------------------------------------------------------- }

procedure ResetGridSort(AGridView: TcxCustomGridView);
var
  ADBTableView: TcxGridDBTableView;
  ADBBandView: TcxGridDBBandedTableView;

  { ------------------------------------------------------------- }

  procedure ResetColumnSort;
  var
    AIndex: Integer;
  begin
    for AIndex := 0 to ADBTableView.ColumnCount - 1 do
      ADBTableView.Columns[AIndex].SortOrder := soNone;
  end;

  { ------------------------------------------------------------- }

  procedure ResetColumnSort2;
  var
    AIndex: Integer;
  begin
    for AIndex := 0 to ADBBandView.ColumnCount - 1 do
      ADBBandView.Columns[AIndex].SortOrder := soNone;
  end;

  { ------------------------------------------------------------- }

begin
  ADBTableView := nil;
  ADBBandView := nil;
  if AGridView is TcxGridDBTableView then
    ADBTableView := TcxGridDBTableView( AGridView )
  else if AGridView is TcxGridDBBandedTableView then
    ADBBandView :=  TcxGridDBBandedTableView( AGridView );
  if Assigned( ADBTableView ) then
    ResetColumnSort
  else if Assigned( ADBBandView ) then
    ResetColumnSort2;
end;

{ ---------------------------------------------------------------------------- }

{ const ALen: Integer,  商品長度請傳 6, 行銷案請傳 8  }

function CalcUploadFtpFullPath(ARoot: String; const AID: String;
  const ALen: Integer): String;
var
  AOuterPath, AInnerPath: String;
begin
  if ( AID = '' ) or ( ARoot = '' ) then Exit;
  if not IsDelimiter( '/', ARoot, Length( ARoot ) ) then
    ARoot := ( ARoot + '/' );
  AOuterPath := CalcUploadFtpOuterPath( AID, ALen );
  if not IsDelimiter( '/', AOuterPath, Length( AOuterPath ) ) then
    AOuterPath := ( AOuterPath + '/' );
  AInnerPath := Lpad( AID, ALen, '0' );
  Result := ( ARoot + AOuterPath + AInnerPath );
end;

{ ---------------------------------------------------------------------------- }


function CalcUploadFtpInnerPath(const AID: String; const ALen: Integer): String;
begin
  if AID = '' then raise Exception.Create( '編號錯誤。' );
  Result := Lpad( AID, ALen, '0' );
end;

{ ---------------------------------------------------------------------------- }

function CalcUploadFtpOuterPath(const AID: String; const ALen: Integer): String;
begin
  if AID = '' then raise Exception.Create( '編號錯誤。' );
  Result := Lpad( IntToStr( StrToInt( AID ) div 100 ), ALen - 2 , '0' );
end;

{ ---------------------------------------------------------------------------- }

end.
