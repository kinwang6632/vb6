unit cbControls;

interface

uses
  Windows, Messages, SysUtils, Types, Classes, Graphics, Forms, Controls, Dialogs, 
{ Developer Express: }
  cxControls, cxClasses, cxGraphics, cxStyles, cxLookAndFeels,
  dxOffice11, cxGeometry,
  dxRibbon, dxRibbonForm, dxRibbonGallery, dxRibbonGalleryFilterEd,
  dxSkinsCore, dxSkinsDefaultPainters, dxSkinsdxRibbonPainter,
  dxSkinsdxBarPainter,
  dxBar, dxBarExtItems;

const
  SchemeColorCount = 10;

type

  TColorMap = array [0..SchemeColorCount-1] of TColor;

  TColorPickerController = class
  private
    FColor: TColor;
    FColorGlyphSize: Integer;
    FColorDialog: TColorDialog;
    FColorDropDownGallery: TdxRibbonDropDownGallery;
    FColorItem: TdxRibbonGalleryItem;
    FColorMapItem: TdxRibbonGalleryItem;
    FThemeColorsGroup: TdxRibbonGalleryGroup;
    FAccentColorsGroup: TdxRibbonGalleryGroup;
    FStandardColorsGroup: TdxRibbonGalleryGroup;
    FCustomColorsGroup: TdxRibbonGalleryGroup;
    FMoreColorsButton: TdxBarButton;
    FNoColorButton: TdxBarButton;
    FOnColorChanged: TNotifyEvent;
    function GetBarManager: TdxBarManager;
    function GetRibbon: TdxCustomRibbon;
    procedure SetColor(Value: TColor);
    procedure ColorItemClick(Sender: TdxRibbonGalleryItem; AItem: TdxRibbonGalleryGroupItem);
    procedure ColorMapItemClick(Sender: TdxRibbonGalleryItem; AItem: TdxRibbonGalleryGroupItem);
    procedure NoColorButtonClick(Sender: TObject);
    procedure MoreColorsClick(Sender: TObject);
    property BarManager: TdxBarManager read GetBarManager;
    property Ribbon: TdxCustomRibbon read GetRibbon;
  protected
    function AddColorItem(AGalleryGroup: TdxRibbonGalleryGroup; AColor: TColor): TdxRibbonGalleryGroupItem;
    function CreateColorBitmap(AColor: TColor; AGlyphSize: Integer = 0): TcxBitmap;
    procedure CreateColorRow(AGalleryGroup: TdxRibbonGalleryGroup; AColorMap: TColorMap);
    procedure BuildThemeColorGallery;
    procedure BuildStandardColorGallery;
    procedure BuildColorSchemeGallery;
    procedure ColorChanged;
    procedure ColorMapChanged;
  public
    constructor Create(AColorItem, AColorMapItem: TdxRibbonGalleryItem; AColorDropDownGallery: TdxRibbonDropDownGallery);
    destructor Destroy; override;
    property Color: TColor read FColor;
    property OnColorChanged: TNotifyEvent read FOnColorChanged write FOnColorChanged;
  end;

implementation

type
  TColorMapInfo = record
    Name: string;
    Map: TColorMap;
  end;
  TAccent = (aLight80, aLight60, aLight50, aLight40, aLight35, aLight25, aLight15, aLight5, aDark10, aDark25, aDark50, aDark75, aDark90);

const

  AStandardColorMap: TColorMap =
    ($0000C0, $0000FF, $00C0FF, $00FFFF, $50D092, $50B000, $F0B000, $C07000, $602000, $A03070);
  AColorMaps: array [0..5] of TColorMapInfo =(
    (Name: '預設'; Map: (clWindow, clWindowText, $D2B48C, $00008B, $0000FF, $FF0000, $556B2F, $800080, clAqua, $FFA500)),
    (Name: '色彩1'; Map: (clWindow, clWindowText, $7D491F, $E1ECEE, $BD814F, $4D50C0, $59BB9B, $A26480, $C6AC4B, $4696F7)),
    (Name: '色彩2'; Map: (clWindow, clWindowText, $6D6769, $D1C2C9, $66B9CE, $84B09C, $C9B16B, $CF8565, $C96B7E, $BB79A3)),
    (Name: '色彩3'; Map: (clWindow, clWindowText, $323232, $D1DEE3, $097FF0, $36299F, $7C581B, $42854E, $784860, $5998C1)),
    (Name: '色彩4'; Map: (clWindow, clWindowText, $866B64, $D7D1C5, $4963D1, $00B4CC, $AEAD8C, $707B8C, $8CB08F, $4990D1)),
    (Name: '色彩5'; Map: (clWindow, clWindowText, $464646, $FAF5DE, $BFA22D, $281FDA, $1B64EB, $9D6339, $784B47, $4A3C7D))
  );


{ TColorPickerController }

constructor TColorPickerController.Create(AColorItem,
  AColorMapItem: TdxRibbonGalleryItem; AColorDropDownGallery: TdxRibbonDropDownGallery);

  procedure InitColorItem;
  begin
    FColorItem.GalleryOptions.ColumnCount := SchemeColorCount;
    FColorItem.GalleryOptions.SpaceBetweenGroups := 4;
    FColorItem.GalleryOptions.ItemTextKind := itkNone;
    FColorItem.OnGroupItemClick := ColorItemClick;

    FThemeColorsGroup := FColorItem.GalleryGroups.Add;
    FThemeColorsGroup.Header.Caption := '色彩';
    FThemeColorsGroup.Header.Visible := True;
    FAccentColorsGroup := FColorItem.GalleryGroups.Add;
    FStandardColorsGroup := FColorItem.GalleryGroups.Add;
    FStandardColorsGroup.Header.Caption := '標準';
    FStandardColorsGroup.Header.Visible := True;
    FCustomColorsGroup := FColorItem.GalleryGroups.Add;
    FCustomColorsGroup.Header.Caption := '自訂';
  end;

  procedure InitColorMapItem;
  begin
    FColorMapItem.GalleryOptions.ColumnCount := 1;
    FColorMapItem.GalleryOptions.SpaceBetweenItemsAndBorder := 0;
    FColorMapItem.GalleryOptions.ItemTextKind := itkCaption;
    FColorMapItem.GalleryGroups.Add;
    FColorMapItem.OnGroupItemClick := ColorMapItemClick;
  end;

  procedure InitDropDownGallery;
  var
    ANoColorGlyph: TBitmap;
  begin
    FNoColorButton := TdxBarButton(FColorDropDownGallery.ItemLinks.AddButton.Item);
    FNoColorButton.ButtonStyle := bsChecked;
    FNoColorButton.Caption := '無';
    FNoColorButton.OnClick := NoColorButtonClick;
    ANoColorGlyph := CreateColorBitmap(clNone, 16);
    FNoColorButton.Glyph := ANoColorGlyph;
    ANoColorGlyph.Free;
    FMoreColorsButton := TdxBarButton(FColorDropDownGallery.ItemLinks.AddButton.Item);
    FMoreColorsButton.Caption := '自訂...';
    FMoreColorsButton.OnClick := MoreColorsClick;
  end;

  procedure PopulateGalleries;
  begin
    BuildColorSchemeGallery;
    BuildStandardColorGallery;
  end;

  procedure SelectDefaultColor;
  begin
    FNoColorButton.Click;
  end;

begin
  inherited Create;
  FColorItem := AColorItem;
  FColorMapItem := AColorMapItem;
  FColorDropDownGallery := AColorDropDownGallery;
  FColorGlyphSize := cxTextHeight(Ribbon.Fonts.Group);
  FColorDialog := TColorDialog.Create(nil);

  InitColorMapItem;
  InitColorItem;
  InitDropDownGallery;
  PopulateGalleries;
  SelectDefaultColor;
end;

destructor TColorPickerController.Destroy;
begin
  FreeAndNil(FColorDialog);
  inherited;
end;

function TColorPickerController.AddColorItem(AGalleryGroup: TdxRibbonGalleryGroup; AColor: TColor): TdxRibbonGalleryGroupItem;
var
  ABitmap: TcxBitmap;
  AColorName: string;
begin
  Result := AGalleryGroup.Items.Add;

  ABitmap := CreateColorBitmap(AColor);
  try
    Result.Glyph := ABitmap;
    if cxNameByColor(AColor, AColorName) then
      Result.Caption := AColorName
    else
      Result.Caption := '$' + IntToHex(AColor, 6);
    Result.Tag := AColor;
  finally
    ABitmap.Free;
  end;
end;

function TColorPickerController.CreateColorBitmap(AColor: TColor; AGlyphSize: Integer): TcxBitmap;
begin
  if AGlyphSize = 0 then
    AGlyphSize := FColorGlyphSize;
  Result := TcxBitmap.CreateSize(AGlyphSize, AGlyphSize);
  FillRectByColor(Result.Canvas.Handle, Result.ClientRect, AColor);
  FrameRectByColor(Result.Canvas.Handle, Result.ClientRect, clGray);
  if AColor = clNone then
    Result.RecoverAlphaChannel(0)
  else
    Result.TransformBitmap(btmSetOpaque);
end;

procedure TColorPickerController.CreateColorRow(AGalleryGroup: TdxRibbonGalleryGroup; AColorMap: TColorMap);
var
  I: Integer;
begin
  for I := Low(AColorMap) to High(AColorMap) do
    AddColorItem(AGalleryGroup, AColorMap[I]);
end;

procedure TColorPickerController.BuildThemeColorGallery;

const
  AnAccentCount = 5;

  function GetBrightness(ARGBColor: DWORD): Integer;
  begin
    Result := (GetBValue(ARGBColor) + GetGValue(ARGBColor) + GetRValue(ARGBColor)) div 3;
  end;

  procedure GetAccentColorScheme(AColorMap: TColorMap; var AnAccentColorScheme: array of TColorMap);

    procedure CreateAccent(AnAccents: array of TAccent; AMapIndex: Integer);
    var
      I: Integer;
      AColor: TColor;
    begin
      for I := Low(AnAccents) to High(AnAccents) do
      begin
        case AnAccents[I] of
          aLight80: AColor := Light(AColorMap[AMapIndex], 80);
          aLight60: AColor := Light(AColorMap[AMapIndex], 60);
          aLight50: AColor := Light(AColorMap[AMapIndex], 50);
          aLight40: AColor := Light(AColorMap[AMapIndex], 40);
          aLight35: AColor := Light(AColorMap[AMapIndex], 35);
          aLight25: AColor := Light(AColorMap[AMapIndex], 25);
          aLight15: AColor := Light(AColorMap[AMapIndex], 15);
          aLight5: AColor := Light(AColorMap[AMapIndex], 5);
          aDark10: AColor := Dark(AColorMap[AMapIndex], 90);
          aDark25: AColor := Dark(AColorMap[AMapIndex], 75);
          aDark50: AColor := Dark(AColorMap[AMapIndex], 50);
          aDark75: AColor := Dark(AColorMap[AMapIndex], 25);
        else {aDark90}
          AColor := Dark(AColorMap[I], 10);
        end;
        AnAccentColorScheme[I][AMapIndex] := AColor;
      end;
    end;

  var
    I: Integer;
  begin
    for I := Low(AColorMap) to High(AColorMap) do
      if GetBrightness(ColorToRGB(AColorMap[I])) < 20 then
        CreateAccent([aLight50, aLight35, aLight25, aLight15, aLight5], I)
      else
        if GetBrightness(ColorToRGB(AColorMap[I])) < 230 then
          CreateAccent([aLight80, aLight60, aLight60, aDark25, aDark50], I)
        else
          CreateAccent([aDark10, aDark25, aDark50, aDark75, aDark90], I)
  end;

var
  I: Integer;
  AColorMap: TColorMap;
  AnAccentColorScheme: array [0..AnAccentCount-1] of TColorMap;
begin
  BarManager.BeginUpdate;
  try
    FThemeColorsGroup.Items.Clear;
    AColorMap := AColorMaps[FColorMapItem.SelectedGroupItem.Index].Map;
    CreateColorRow(FThemeColorsGroup, AColorMap);

    FAccentColorsGroup.Items.Clear;
    GetAccentColorScheme(AColorMap, AnAccentColorScheme);
    for I := Low(AnAccentColorScheme) to High(AnAccentColorScheme) do
      CreateColorRow(FAccentColorsGroup, AnAccentColorScheme[I]);
  finally
    BarManager.EndUpdate;
  end;
end;

procedure TColorPickerController.BuildStandardColorGallery;
begin
  BarManager.BeginUpdate;
  try
    FStandardColorsGroup.Items.Clear;
    CreateColorRow(FStandardColorsGroup, AStandardColorMap);
  finally
    BarManager.EndUpdate;
  end;
end;

procedure TColorPickerController.BuildColorSchemeGallery;
const
  ASystemColorCount = 2;
  AGlyphOffset = 1;
var
  I, J: Integer;
  ABitmap, AColorBitmap: TcxBitmap;
  ARect: TRect;
  AGroupItem: TdxRibbonGalleryGroupItem;
  AThemeColorCount: Integer;
begin
  BarManager.BeginUpdate;
  try
    AThemeColorCount := SchemeColorCount - ASystemColorCount;
    ABitmap := TcxBitmap.CreateSize(FColorGlyphSize * AThemeColorCount + (AThemeColorCount - 1) * AGlyphOffset, FColorGlyphSize);
    try
      for I := High(AColorMaps) downto Low(AColorMaps) do
      begin
        AGroupItem := FColorMapItem.GalleryGroups[0].Items.Insert(0);
        for J := Low(AColorMaps[I].Map) + ASystemColorCount to High(AColorMaps[I].Map) do
        begin
          AColorBitmap := CreateColorBitmap(AColorMaps[I].Map[J]);
          try
            ARect := cxRectOffset(AColorBitmap.ClientRect, (AColorBitmap.Width + AGlyphOffset) * (J - ASystemColorCount), 0);
            ABitmap.CopyBitmap(AColorBitmap, ARect, cxNullPoint);
          finally
            AColorBitmap.Free;
          end;
        end;
        AGroupItem.Glyph := ABitmap;
        AGroupItem.Caption := AColorMaps[I].Name;
      end;
      AGroupItem.Selected := True;
    finally
      ABitmap.Free;
    end;
  finally
    BarManager.EndUpdate;
  end;
end;

procedure TColorPickerController.ColorChanged;
var
  AGlyph: TcxBitmap;
begin
  AGlyph := CreateColorBitmap(Color, Round(16 * Screen.PixelsPerInch / 96));
  try
    FColorItem.Glyph := AGlyph;
  finally
    AGlyph.Free;
  end;

  if Assigned(OnColorChanged) then
    OnColorChanged(Self);
end;

procedure TColorPickerController.ColorMapChanged;

  procedure FillGlyph(AGlyph: TcxBitmap);
  var
    ARect: TRect;
    ADC: HDC;
  begin
    ARect := Rect(0, 0, AGlyph.Width div 2, AGlyph.Height div 2);
    ADC := AGlyph.Canvas.Handle;
    FillRectByColor(ADC, ARect, AColorMaps[FColorMapItem.SelectedGroupItem.Index].Map[2]);
    FillRectByColor(ADC, cxRectOffset(ARect, cxRectWidth(ARect), 0), AColorMaps[FColorMapItem.SelectedGroupItem.Index].Map[3]);
    FillRectByColor(ADC, cxRectOffset(ARect, 0, cxRectHeight(ARect)), AColorMaps[FColorMapItem.SelectedGroupItem.Index].Map[4]);
    FillRectByColor(ADC, cxRectOffset(ARect, cxRectWidth(ARect), cxRectHeight(ARect)), AColorMaps[FColorMapItem.SelectedGroupItem.Index].Map[5]);
    FrameRectByColor(ADC, AGlyph.ClientRect, clGray);
    AGlyph.TransformBitmap(btmSetOpaque);
  end;

var
  AGlyph: TcxBitmap;
begin
  BarManager.BeginUpdate;
  try
    AGlyph := TcxBitmap.CreateSize(16, 16);
    FillGlyph(AGlyph);
    FColorMapItem.Glyph := AGlyph;
    AGlyph.SetSize(32, 32);
    FillGlyph(AGlyph);
    FColorMapItem.LargeGlyph := AGlyph;
    AGlyph.Free;
  finally
    BarManager.EndUpdate(False);
  end
end;

function TColorPickerController.GetBarManager: TdxBarManager;
begin
  Result := FColorItem.BarManager;
end;

function TColorPickerController.GetRibbon: TdxCustomRibbon;
begin
  Result := FColorDropDownGallery.Ribbon;
end;

procedure TColorPickerController.SetColor(Value: TColor);
begin
  if FColor <> Value then
  begin
    FColor := Value;
    ColorChanged;
  end;
end;

procedure TColorPickerController.ColorItemClick(Sender: TdxRibbonGalleryItem; AItem: TdxRibbonGalleryGroupItem);
begin
  FNoColorButton.Down := False;
  if FColorItem.SelectedGroupItem <> nil then
    SetColor(FColorItem.SelectedGroupItem.Tag);
end;

procedure TColorPickerController.ColorMapItemClick(Sender: TdxRibbonGalleryItem; AItem: TdxRibbonGalleryGroupItem);
begin
  BuildThemeColorGallery;
  ColorMapChanged;
end;

procedure TColorPickerController.NoColorButtonClick(Sender: TObject);
begin
  if FColorItem.SelectedGroupItem <> nil then
    FColorItem.SelectedGroupItem.Selected := False;
  SetColor(clNone);
end;

procedure TColorPickerController.MoreColorsClick(Sender: TObject);
begin
  FColorDialog.Color := Color;
  if FColorDialog.Execute then
  begin
    FCustomColorsGroup.Header.Visible := True;
    AddColorItem(FCustomColorsGroup, FColorDialog.Color).Selected := True;
  end;
end;

end.
