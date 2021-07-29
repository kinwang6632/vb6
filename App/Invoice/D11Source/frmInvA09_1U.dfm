object frmInvA09_1: TfrmInvA09_1
  Left = 267
  Top = 193
  Width = 577
  Height = 447
  Caption = 'frmInvA09_1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 569
    Height = 32
    Align = alTop
    Color = clSkyBlue
    TabOrder = 0
    object Label6: TLabel
      Left = 10
      Top = 8
      Width = 136
      Height = 16
      Caption = #30332#31080#36039#26009#27298#26680#32080#26524
      Font.Charset = ANSI_CHARSET
      Font.Color = clBlue
      Font.Height = -16
      Font.Name = #26032#32048#26126#39636
      Font.Style = [fsBold]
      ParentFont = False
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 32
    Width = 569
    Height = 388
    Align = alClient
    Caption = 'Panel2'
    TabOrder = 1
    object Label11: TLabel
      Left = 11
      Top = 358
      Width = 91
      Height = 13
      Alignment = taRightJustify
      Caption = 'Excel'#23384#25918#36335#24465#65306
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clTeal
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      Transparent = True
    end
    object PageControl1: TPageControl
      Left = 1
      Top = 1
      Width = 568
      Height = 344
      ActivePage = tbsChkNumber
      BiDiMode = bdLeftToRight
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlack
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      MultiLine = True
      ParentBiDiMode = False
      ParentFont = False
      TabOrder = 0
      TabPosition = tpBottom
      object tbsChkNumber: TTabSheet
        Caption = #32113#32232#27298#26680
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ParentFont = False
        object DBGrid1: TDBGrid
          Left = 8
          Top = 8
          Width = 513
          Height = 145
          DataSource = dsrCheckBusinessID
          TabOrder = 0
          TitleFont.Charset = ANSI_CHARSET
          TitleFont.Color = clWindowText
          TitleFont.Height = -13
          TitleFont.Name = #26032#32048#26126#39636
          TitleFont.Style = []
        end
        object btnToExcel_1: TBitBtn
          Left = 525
          Top = 128
          Width = 28
          Height = 25
          TabOrder = 1
          OnClick = btnToExcel_1Click
          Glyph.Data = {
            76010000424D7601000000000000760000002800000020000000100000000100
            04000000000000010000130B0000130B00001000000000000000000000000000
            800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
            FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333303
            333333333333337FF3333333333333903333333333333377FF33333333333399
            03333FFFFFFFFF777FF3000000999999903377777777777777FF0FFFF0999999
            99037F3337777777777F0FFFF099999999907F3FF777777777770F00F0999999
            99037F773777777777730FFFF099999990337F3FF777777777330F00FFFFF099
            03337F773333377773330FFFFFFFF09033337F3FF3FFF77733330F00F0000003
            33337F773777777333330FFFF0FF033333337F3FF7F3733333330F08F0F03333
            33337F7737F7333333330FFFF003333333337FFFF77333333333000000333333
            3333777777333333333333333333333333333333333333333333}
          NumGlyphs = 2
        end
        object DBGrid4: TDBGrid
          Left = 8
          Top = 160
          Width = 513
          Height = 145
          DataSource = dsrCheckJumpInvID
          TabOrder = 2
          TitleFont.Charset = ANSI_CHARSET
          TitleFont.Color = clWindowText
          TitleFont.Height = -13
          TitleFont.Name = #26032#32048#26126#39636
          TitleFont.Style = []
        end
        object btnToExcel_2: TBitBtn
          Left = 524
          Top = 280
          Width = 28
          Height = 25
          TabOrder = 3
          OnClick = btnToExcel_2Click
          Glyph.Data = {
            76010000424D7601000000000000760000002800000020000000100000000100
            04000000000000010000130B0000130B00001000000000000000000000000000
            800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
            FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333303
            333333333333337FF3333333333333903333333333333377FF33333333333399
            03333FFFFFFFFF777FF3000000999999903377777777777777FF0FFFF0999999
            99037F3337777777777F0FFFF099999999907F3FF777777777770F00F0999999
            99037F773777777777730FFFF099999990337F3FF777777777330F00FFFFF099
            03337F773333377773330FFFFFFFF09033337F3FF3FFF77733330F00F0000003
            33337F773777777333330FFFF0FF033333337F3FF7F3733333330F08F0F03333
            33337F7737F7333333330FFFF003333333337FFFF77333333333000000333333
            3333777777333333333333333333333333333333333333333333}
          NumGlyphs = 2
        end
      end
      object tbsChkAmount: TTabSheet
        Caption = #20027#21103#27284#37329#38989#27298#26597
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ImageIndex = 1
        ParentFont = False
        object btnToExcel_3: TBitBtn
          Left = 524
          Top = 280
          Width = 28
          Height = 25
          TabOrder = 0
          OnClick = btnToExcel_3Click
          Glyph.Data = {
            76010000424D7601000000000000760000002800000020000000100000000100
            04000000000000010000130B0000130B00001000000000000000000000000000
            800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
            FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333303
            333333333333337FF3333333333333903333333333333377FF33333333333399
            03333FFFFFFFFF777FF3000000999999903377777777777777FF0FFFF0999999
            99037F3337777777777F0FFFF099999999907F3FF777777777770F00F0999999
            99037F773777777777730FFFF099999990337F3FF777777777330F00FFFFF099
            03337F773333377773330FFFFFFFF09033337F3FF3FFF77733330F00F0000003
            33337F773777777333330FFFF0FF033333337F3FF7F3733333330F08F0F03333
            33337F7737F7333333330FFFF003333333337FFFF77333333333000000333333
            3333777777333333333333333333333333333333333333333333}
          NumGlyphs = 2
        end
        object DBGrid2: TDBGrid
          Left = 8
          Top = 8
          Width = 513
          Height = 297
          DataSource = dsrCheckAmount
          TabOrder = 1
          TitleFont.Charset = ANSI_CHARSET
          TitleFont.Color = clWindowText
          TitleFont.Height = -13
          TitleFont.Name = #26032#32048#26126#39636
          TitleFont.Style = []
        end
      end
      object tbsChkMisData: TTabSheet
        Caption = #30332#31080#38283#31435#33287#23458#26381#36039#26009#27604#23565
        Font.Charset = ANSI_CHARSET
        Font.Color = clBlack
        Font.Height = -13
        Font.Name = #26032#32048#26126#39636
        Font.Style = []
        ImageIndex = 2
        ParentFont = False
        object DBGrid3: TDBGrid
          Left = 8
          Top = 8
          Width = 513
          Height = 297
          DataSource = dsrCheckMisData
          TabOrder = 0
          TitleFont.Charset = ANSI_CHARSET
          TitleFont.Color = clBlack
          TitleFont.Height = -13
          TitleFont.Name = #26032#32048#26126#39636
          TitleFont.Style = []
        end
        object btnToExcel_4: TBitBtn
          Left = 524
          Top = 280
          Width = 28
          Height = 25
          TabOrder = 1
          OnClick = btnToExcel_4Click
          Glyph.Data = {
            76010000424D7601000000000000760000002800000020000000100000000100
            04000000000000010000130B0000130B00001000000000000000000000000000
            800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
            FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333303
            333333333333337FF3333333333333903333333333333377FF33333333333399
            03333FFFFFFFFF777FF3000000999999903377777777777777FF0FFFF0999999
            99037F3337777777777F0FFFF099999999907F3FF777777777770F00F0999999
            99037F773777777777730FFFF099999990337F3FF777777777330F00FFFFF099
            03337F773333377773330FFFFFFFF09033337F3FF3FFF77733330F00F0000003
            33337F773777777333330FFFF0FF033333337F3FF7F3733333330F08F0F03333
            33337F7737F7333333330FFFF003333333337FFFF77333333333000000333333
            3333777777333333333333333333333333333333333333333333}
          NumGlyphs = 2
        end
      end
    end
    object btnExit: TBitBtn
      Left = 492
      Top = 356
      Width = 71
      Height = 26
      Caption = #32080#26463
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      TabOrder = 3
      OnClick = btnExitClick
      Glyph.Data = {
        76010000424D7601000000000000760000002800000020000000100000000100
        04000000000000010000130B0000130B00001000000000000000000000000000
        800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
        FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333333333
        333333333333333333333333333333333333333FFF33FF333FFF339993370733
        999333777FF37FF377733339993000399933333777F777F77733333399970799
        93333333777F7377733333333999399933333333377737773333333333990993
        3333333333737F73333333333331013333333333333777FF3333333333910193
        333333333337773FF3333333399000993333333337377737FF33333399900099
        93333333773777377FF333399930003999333337773777F777FF339993370733
        9993337773337333777333333333333333333333333333333333333333333333
        3333333333333333333333333333333333333333333333333333}
      NumGlyphs = 2
    end
    object btnTransferPath: TButton
      Left = 398
      Top = 357
      Width = 22
      Height = 22
      Caption = '...'
      TabOrder = 2
      OnClick = btnTransferPathClick
    end
    object edtExcelPath: TEdit
      Left = 113
      Top = 358
      Width = 281
      Height = 21
      Font.Charset = CHINESEBIG5_CHARSET
      Font.Color = clBlack
      Font.Height = -13
      Font.Name = #26032#32048#26126#39636
      Font.Style = []
      ParentFont = False
      ReadOnly = True
      TabOrder = 1
      Text = 'C:'
    end
  end
  object dsrCheckBusinessID: TDataSource
    DataSet = dtmMainJ.cdsCheckBusinessID
    Left = 16
    Top = 328
  end
  object dsrCheckJumpInvID: TDataSource
    DataSet = dtmMainJ.cdsCheckJumpInvID
    Left = 40
    Top = 328
  end
  object dsrCheckAmount: TDataSource
    DataSet = dtmMainJ.cdsCheckAmount
    Left = 104
    Top = 328
  end
  object dsrCheckMisData: TDataSource
    DataSet = dtmMainJ.cdsCheckMisData
    Left = 200
    Top = 328
  end
end
