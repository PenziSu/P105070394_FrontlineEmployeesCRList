object Form1: TForm1
  Left = 192
  Top = 107
  Width = 467
  Height = 340
  Caption = '�NDBGrid���Excel'
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = '�ө���'
  Font.Style = []
  OldCreateOrder = False
  Scaled = False
  PixelsPerInch = 96
  TextHeight = 12
  object Label1: TLabel
    Left = 12
    Top = 256
    Width = 114
    Height = 12
    Caption = '��X��Excel�ɮצW��'
  end
  object LabelDelphiKTop: TLabel
    Left = 0
    Top = 0
    Width = 459
    Height = 16
    Cursor = crHandPoint
    Hint = '�����s����Delphi K.Top����'
    Align = alTop
    Alignment = taCenter
    Caption = '�{���ӷ� http://delphi.ktop.com.tw'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    ParentShowHint = False
    ShowHint = True
    OnClick = LabelDelphiKTopClick
  end
  object DBGrid1: TDBGrid
    Left = 12
    Top = 24
    Width = 429
    Height = 217
    DataSource = DataSource1
    TabOrder = 0
    TitleFont.Charset = ANSI_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -12
    TitleFont.Name = '�ө���'
    TitleFont.Style = []
  end
  object Button1: TButton
    Left = 132
    Top = 284
    Width = 145
    Height = 25
    Caption = '�NDBGrid���Excel'
    TabOrder = 1
    OnClick = Button1Click
  end
  object Edit1: TEdit
    Left = 132
    Top = 252
    Width = 309
    Height = 20
    TabOrder = 2
    Text = 'c:\test.xls'
  end
  object Table1: TTable
    Active = True
    DatabaseName = 'DBDEMOS'
    TableName = 'country.DB'
    Left = 136
    Top = 80
  end
  object DataSource1: TDataSource
    DataSet = Table1
    Left = 96
    Top = 80
  end
end
