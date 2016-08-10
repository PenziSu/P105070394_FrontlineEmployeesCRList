object Form1: TForm1
  Left = 192
  Top = 107
  Width = 700
  Height = 500
  Caption = 'GridExport Demo'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 85
    Top = 185
    Width = 86
    Height = 13
    Caption = #25353' Title '#21487#20197#25490#24207
  end
  object Label2: TLabel
    Left = 85
    Top = 5
    Width = 155
    Height = 13
    Caption = 'Design time '#26377#35373#23450' PopupMenu'
  end
  object DBGrid1: TDBGrid
    Left = 85
    Top = 20
    Width = 600
    Height = 150
    DataSource = DataSource1
    PopupMenu = PopupMenu1
    TabOrder = 0
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
  end
  object StringGrid1: TStringGrid
    Left = 85
    Top = 360
    Width = 600
    Height = 106
    DefaultColWidth = 10
    DefaultRowHeight = 16
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing, goEditing]
    TabOrder = 1
  end
  object Button1: TButton
    Left = 5
    Top = 20
    Width = 75
    Height = 25
    Caption = 'BDE Table'
    TabOrder = 2
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 5
    Top = 360
    Width = 75
    Height = 25
    Caption = 'Fill String'
    TabOrder = 3
    OnClick = Button2Click
  end
  object Button3: TButton
    Left = 5
    Top = 390
    Width = 75
    Height = 25
    Caption = 'Resize'
    TabOrder = 4
    OnClick = Button3Click
  end
  object DBGrid2: TDBGrid
    Left = 85
    Top = 200
    Width = 600
    Height = 150
    DataSource = DataSource2
    TabOrder = 5
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
  end
  object Button4: TButton
    Left = 5
    Top = 200
    Width = 75
    Height = 25
    Caption = 'ADO Query'
    TabOrder = 6
    OnClick = Button4Click
  end
  object Table1: TTable
    DatabaseName = 'DBDEMOS'
    TableName = 'country.db'
    Left = 205
    Top = 50
  end
  object DataSource1: TDataSource
    DataSet = Table1
    Left = 205
    Top = 100
  end
  object GridExport1: TGridExport
    Left = 455
  end
  object PopupMenu1: TPopupMenu
    Left = 410
    Top = 80
    object N1: TMenuItem
      Caption = #23531#22312#31243#24335#35041
    end
  end
  object ADOConnection1: TADOConnection
    ConnectionString = 
      'FILE NAME=C:\Program Files\Common Files\System\Ole DB\Data Links' +
      '\DBDEMOS.udl'
    KeepConnection = False
    LoginPrompt = False
    Mode = cmShareDenyNone
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 205
    Top = 215
  end
  object ADOQuery1: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    Prepared = True
    Left = 205
    Top = 260
  end
  object DataSource2: TDataSource
    DataSet = ADOQuery1
    Left = 205
    Top = 305
  end
end
