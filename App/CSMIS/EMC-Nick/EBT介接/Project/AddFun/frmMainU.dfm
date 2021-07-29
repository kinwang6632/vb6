object frmMain: TfrmMain
  Left = 358
  Top = 136
  BorderIcons = []
  BorderStyle = bsSingle
  Caption = 'CSIS'
  ClientHeight = 216
  ClientWidth = 109
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  Scaled = False
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object CoolBar1: TCoolBar
    Left = 0
    Top = 0
    Width = 145
    Height = 216
    Align = alLeft
    Bands = <
      item
        Break = False
        Color = clInfoBk
        Control = ToolBar6
        ImageIndex = -1
        MinHeight = 105
        ParentColor = False
        Text = '&1. '#20351#29992#32773#31649#29702
        Width = 133
      end
      item
        Break = False
        Color = clAqua
        Control = ToolBar5
        ImageIndex = -1
        MinHeight = 62
        ParentColor = False
        Width = 77
      end>
    Color = clGradientInactiveCaption
    ParentColor = False
    Vertical = True
    object ToolBar5: TToolBar
      Left = 21
      Top = 144
      Width = 62
      Height = 23
      AutoSize = True
      ButtonHeight = 21
      ButtonWidth = 55
      Caption = 'ToolBar5'
      Color = clAqua
      Flat = True
      ParentColor = False
      ParentShowHint = False
      ShowCaptions = True
      ShowHint = True
      TabOrder = 0
      object ToolButton36: TToolButton
        Left = 0
        Top = 0
        AutoSize = True
        Caption = #32080#26463#31243#24335
        ImageIndex = 0
        OnClick = ToolButton36Click
      end
    end
    object ToolBar6: TToolBar
      Left = 0
      Top = 26
      Width = 105
      Height = 103
      ButtonHeight = 21
      ButtonWidth = 110
      Caption = 'ToolBar6'
      Color = clInfoBk
      Flat = True
      ParentColor = False
      ShowCaptions = True
      TabOrder = 1
      object ToolButton38: TToolButton
        Left = 0
        Top = 0
        Caption = '&A. '#20351#29992#32773#36039#26009#32173#35703
        ImageIndex = 1
        Wrap = True
        OnClick = ToolButton38Click
      end
      object ToolButton39: TToolButton
        Left = 0
        Top = 21
        Caption = '&B. '#32676#32068#36039#26009#32173#35703
        ImageIndex = 2
        Wrap = True
        OnClick = ToolButton39Click
      end
      object ToolButton1: TToolButton
        Left = 0
        Top = 42
        Caption = '&C.'#20844#21496#21029#36039#26009#32173#35703
        ImageIndex = 3
        Wrap = True
        OnClick = ToolButton1Click
      end
      object ToolButton2: TToolButton
        Left = 0
        Top = 63
        Caption = '&D.'#37096#38272#21029#36039#26009#32173#35703
        ImageIndex = 4
        OnClick = ToolButton2Click
      end
    end
  end
end
