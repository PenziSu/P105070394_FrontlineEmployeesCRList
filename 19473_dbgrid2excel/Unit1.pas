unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Grids, DBGrids, Db, DBTables,comobj,shellapi;

type
  TForm1 = class(TForm)
    Table1: TTable;
    DataSource1: TDataSource;
    DBGrid1: TDBGrid;
    Button1: TButton;
    Edit1: TEdit;
    Label1: TLabel;
    LabelDelphiKTop: TLabel;
    procedure Button1Click(Sender: TObject);
    procedure LabelDelphiKTopClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}
procedure DBGrid2Excel(DBGrid:TDBGrid;ExcelFileName:string);
var  MyExcel: Variant;
     x,y:integer;
begin
  deletefile(ExcelFileName);
  MyExcel := CreateOleOBject('Excel.Application');
  // 這一段為會員suda幫忙修正的
  MyExcel.WorkBooks.Add;
  MyExcel.Visible := True;
  MyExcel.WorkBooks[1].Saveas(ExcelFileName);
  dbgrid.DataSource.DataSet.First;

  y:=1;
  for x:=1 to dbgrid.FieldCount do
  begin
    MyExcel.WorkBooks[1].WorkSheets[1].Cells[y,x] := dbgrid.Fields[x-1].DisplayName;
    // 將該欄設為標選
    MyExcel.WorkBooks[1].WorkSheets[1].Cells[y,x].Select;
    // 將標題欄位變粗體字
    MyExcel.Selection.Font.Bold := true;
    // 設定欄位寬度
    MyExcel.WorkBooks[1].WorkSheets[1].Columns[x].ColumnWidth := dbgrid.Fields[x-1].DisplayWidth;
  end;
  inc(y);
  while not dbgrid.DataSource.DataSet.eof do
  begin
    for x:=1 to dbgrid.FieldCount do
    begin
       MyExcel.WorkBooks[1].WorkSheets[1].Cells[y,x] := dbgrid.Fields[x-1].AsString;
    end;
    inc(y);
    dbgrid.DataSource.DataSet.next;
  end;
end;
procedure TForm1.Button1Click(Sender: TObject);
begin
   DBGrid2Excel(dbgrid1,edit1.text);
end;
procedure TForm1.LabelDelphiKTopClick(Sender: TObject);
begin
   ShellExecute(application.handle,pchar('OPEN'),pchar('http://delphi.ktop.com.tw'),nil,nil,0);
end;

end.
