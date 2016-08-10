unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, DB, DBGrids, DBTables, StrUtils, GridExport,
  Menus, ADODB;

type
  TForm1 = class(TForm)
    Table1: TTable;
    DBGrid1: TDBGrid;
    DataSource1: TDataSource;
    StringGrid1: TStringGrid;
    Button1: TButton;
    GridExport1: TGridExport;
    Button2: TButton;
    Button3: TButton;
    PopupMenu1: TPopupMenu;
    N1: TMenuItem;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    DataSource2: TDataSource;
    DBGrid2: TDBGrid;
    Button4: TButton;
    Label1: TLabel;
    Label2: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
begin
  GridExport1.Initial(Self, True);
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  GridExport1.FreeAll;
end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  Table1.Open;
end;

procedure TForm1.Button2Click(Sender: TObject);
var
  C, R: integer;
begin
  with StringGrid1 do
    for C := 1 to ColCount - 1 do
      for R := 1 to RowCount - 1 do
        Cells[C, R] := DupeString('Ya', Random(10) + 1);
end;

procedure TForm1.Button3Click(Sender: TObject);
begin
  AutoSizeStringGridColumn(StringGrid1, Random(StringGrid1.ColCount));
end;

procedure TForm1.Button4Click(Sender: TObject);
begin
  ADOQuery1.SQL.Text := 'SELECT * FROM COUNTRY';
  ADOQuery1.Open;
end;

end.
