unit GridExport;
// Written by Silence Wang ( silence@cmo.com.tw )
// Introduced: some day in 2003
// Last Modified: 2005/01/06
//
// 因為工作上使用到的, 幫部門寫的小程式經常要輸出 Excel
// 所以寫個元件來簡化
//
// 功能：
//  1. DBGrid 或 StringGrid 上按下右鍵後可以選擇 [輸出到 Excel]
//  2. 自動生成 PopupMenu / MenuItem, 並整合入自訂的 PopupMenu, 不必額外煩惱操作介面
//  3. DBGrid 可使用按下 Title 後做 Sort 的功能 (但 DBGrid 的 DataSet 要是 SQL Query, 且不能太複雜)
//  4. 在執行時期若要關閉, 不允許 user 轉出資料, 只要將特定 Grid 的 OnMouseUp := nil 即可
//     要再打開, 只要重新 Initial 即可
//
// 使用：
//  0. 將元件拉放到 Form 上
//  1. 在 FormCreate 中加入
//     GridExport.Initial(Self);
//     即可自動使用, 額外有預設為 False 的參數 aTitleSort: boolean
//  2. 在 FormClose 中加入
//     GridExport.FreeAll;
//     確保 [自有物件] 及 Excel 正確 Free
//  3. 程式執行時, 於 DBGrid / StringGrid 上按右鍵跳出 PopupMenu,
//     其中 MenuItem 會自動跟你指定給 Grid 的 PopupMenu 結合
//  4. 執行 Export to Excel, 從 DBGrid / StringGrid 將資料匯出到 Excel
//
//  附帶函式四個
//  A. 自動設定 String Grid 的第 aCol 個 column width
//  B. 應用 A 來做多個 column 的自動調整
//  C. "自訂 DBGrid 的文字外觀時, 放在 OnDrawColumnCell 的最後來使用"
//  D. "自訂 StringGrid 的文字外觀時, 放在 OnDrawCell 的最後來使用"

interface

uses
  Classes, Windows, Menus, Forms, SysUtils, Controls, Dialogs, Graphics,
  Variants, DB, ADODB, Grids, DBGrids, Excel2000;

type
  TGridExport = class(TComponent)
    ParentForm: TForm;
    procedure Execute(Sender: TObject);
    procedure DBGridToExcel(aEXCEL: TExcelApplication; aGrid: TDBGrid; TrimSpace: Boolean = False);
    procedure StringGridToExcel(aEXCEL: TExcelApplication; aGrid: TStringGrid; TrimSpace: Boolean = False);
    procedure GridMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure GridTitleClick(Column: TColumn);
    function GetColA2Z(Col: Integer):string;
  private
    { Private declarations }
  protected
    { Protected declarations }
  public
    { Public declarations }
    procedure Initial(Owner: TComponent; aTitleSort: boolean = False);
    procedure FreeAll;
  published
    { Published declarations }
  end;

procedure Register;

procedure AutoSizeStringGridColumn(aStringGrid: TStringGrid; aCol: Integer; aMin: Integer = 0; aMax: Integer = 0);
procedure SetStringGridWidth(aStringGrid: TStringGrid; aCols: array of integer);
procedure DrawColumn(var aRect: TRect; var aCanvas: TCanvas; var aColumn: TColumn);
procedure DrawStrColumn(var aRect: TRect; var aCanvas: TCanvas; var aText: string);

implementation

// 附帶函式 A
// 設定 String Grid 的第 aCol 個 column width
// 這個函式針對指定的 aCol 內的字串長度, 按照 aMin 及 aMax 的範圍來做自動設定
// aMin 及 aMax 預設值為 0 則不做限制
// 用法: AutoSizeStringGridColumn(mySGrid, 2);
//       AutoSizeStringGridColumn(mySGrid, 3, 20, 90);
procedure AutoSizeStringGridColumn(aStringGrid: TStringGrid; aCol: Integer; aMin: Integer = 0; aMax: Integer = 0);
var
  iRow, iMax: Integer;
begin
  iMax := 0;
  for iRow := 0 to (aStringGrid.RowCount - 1) do
  begin
    if aStringGrid.Canvas.TextWidth(aStringGrid.Cells[aCol, iRow]) > iMax then
      iMax := aStringGrid.Canvas.TextWidth(aStringGrid.Cells[aCol, iRow]);
    if (aMax > 0) and (iMax > aMax) then
    begin
      iMax := aMax;
      Break;
    end;
  end;
  if (aMin > 0) and (iMax < aMin) then
    iMax := aMin;
  aStringGrid.ColWidths[aCol] := iMax + aStringGrid.GridLineWidth + 4;
end;

// 附帶函式 B
// 應用 AutoSizeStringGridColumn, 只要指定哪些 col 需要調整大小即可, 但不使用 Min / Max 功能
// (不想寫得太複雜, 要每個 column 用不同的 Min/Max, 就單獨使用 AutoSizeStringGridColumn 就好了)
// 用法: SetStringGridWidth(mySGrid1, []);
//       SetStringGridWidth(mySGrid2, colArray);
//       SetStringGridWidth(mySGrid3, [0, 2], 0, 30);
procedure SetStringGridWidth(aStringGrid: TStringGrid; aCols: array of integer);
var
  iCol, I: integer;
begin
  // 計算 StrGrid 各欄的寬度
  if High(aCols) = -1 then
    // 沒指定就全部調整
    for iCol := 0 to aStringGrid.ColCount - 1 do
      AutoSizeStringGridColumn(aStringGrid, iCol, 0, 0)
  else
    // 有指定就逐一調整, 超出範圍的不理會
    for I := Low(aCols) to High(aCols) do
      if (aCols[I] >= 0) and
         (aCols[I] < aStringGrid.ColCount) then
        AutoSizeStringGridColumn(aStringGrid, aCols[I], 0, 0);
end;

// 附帶函式 C
// 這個函式是給 "自訂 DBGrid 的文字外觀時, 放在 OnDrawColumnCell 的最後來使用"
// 因為自訂 DBGrid 的 ColumnCell 的文字外觀時, 要自己控制 Text 的 Alignment;
// 用法:
//  OnDrawColumnCell 內
//    myCanvas := (Sender as TDBGrid).Canvas;
//    myRect := Rect;
//    ... 設定 Text Font 外觀
//    DrawColumn(myRect, myCanvas, Column);
//  end;
procedure DrawColumn(var aRect: TRect; var aCanvas: TCanvas; var aColumn: TColumn);
var
  iX, iY: Integer;
begin
  case aColumn.Alignment of
    taRightJustify: begin
      iX := aRect.BottomRight.X - 3 - aCanvas.TextWidth(aColumn.Field.DisplayText);
      iY := aRect.BottomRight.Y - 2 - aCanvas.TextHeight(aColumn.Field.DisplayText);
    end;
    taCenter: begin
      iX := aRect.Left + (((aRect.Right - aRect.Left) - aCanvas.TextWidth(aColumn.Field.DisplayText)) div 2);
      iY := aRect.Top + (((aRect.Bottom - aRect.Top)- aCanvas.TextHeight(aColumn.Field.DisplayText)) div 2);
    end;
    taLeftJustify: begin
      iX := aRect.Left + 2;
      iY := aRect.Top + 2;
    end;
  end;
  aCanvas.TextRect(aRect, iX, iY, aColumn.Field.DisplayText);
end;

// 附帶函式 D
// 這個函式是給 "自訂 StringGrid 的文字外觀時, 放在 OnDrawCell 的最後來使用"
// 其實只有一行而已 ++" , 還可以擴充啦, 例如靠左靠右靠中之類的
// 用法:
//  OnDrawCell 內
//    myRect := Rect;
//    myCanvas := Canvas;
//    myText := Cells[ACol, ARow];
//    ... 設定 Text Font 外觀
//    DrawStrColumn(myRect, myCanvas, myText);
//  end;
procedure DrawStrColumn(var aRect: TRect; var aCanvas: TCanvas; var aText: string);
begin
  aCanvas.TextRect(aRect, aRect.Left + 2, aRect.Top + 2, aText);
end;


{ TGridExport }

// 一開始的 "啟動"
// 設定 Owner 底下的 Grid ( DBGrid / StringGrid )
procedure TGridExport.Initial(Owner: TComponent; aTitleSort: boolean = False);
var
  I: integer;
begin
  ParentForm := (Owner as TForm);
  for I := 0 to (Owner as TComponent).ComponentCount - 1 do
  begin
    if (Owner as TComponent).Components[I].ClassName = 'TDBGrid' then
    begin
      ((Owner as TComponent).Components[I] as TDBGrid).OnMouseUp := GridMouseUp;
      if aTitleSort then
        ((Owner as TComponent).Components[I] as TDBGrid).OnTitleClick := GridTitleClick;
    end;
    if (Owner as TComponent).Components[I].ClassName = 'TStringGrid' then
      ((Owner as TComponent).Components[I] as TStringGrid).OnMouseUp := GridMouseUp;
  end;
end;

// Release 專用的 TPopupMenu : pmGRID
// 並關閉與 Excel 的連結 ==重要==
// 程式結束時一定要執行, 可以避免 Excel 太多 Instance 及後續開啟 Excel 時產生視窗錯誤
// (就是 "只有出現 Menu 及工具列, 而 Sheet 跑不出來" 的那種錯誤)
procedure TGridExport.FreeAll;
begin
  try
    TPopupMenu(ParentForm.FindComponent('pmGRID')).Free;
    TExcelApplication(ParentForm.FindComponent('ExcelApp')).Free;
  except
    TPopupMenu(ParentForm.FindComponent('pmGRID')).Free;
    TExcelApplication(ParentForm.FindComponent('ExcelApp')).Free;
  end;
end;

// 欄位超過 Z 怎麼辦? 這裡幫你算啦～
function TGridExport.GetColA2Z(Col: Integer):string;
begin
  if (Col mod 26 = 0) then
    Result := 'Z'
  else
    Result := Chr((Col mod 26)+64);
  // 遞迴, 就算你 ZZZ 我也幫你算
  if Col > 26 then
    Result := GetColA2Z(Col div 26) + Result;
end;

// 主要 DBGridToExcel
// TrimSpace? 用來去掉尾端的 "空白" 用的, 有些欄位不是用 Varchar 的 type, 尾端會有 "空白"
// 幹麼還要分, 一律 Trim 掉不是比較乾脆? .... 有人用著就留著吧
procedure TGridExport.DBGridToExcel(aEXCEL: TExcelApplication; aGrid: TDBGrid; TrimSpace: Boolean = False);
var
  iCol, iRow: integer;
  Cell: string;
  myBookMark: TBookMark;
begin
  // 輸出 [Title]
  // 不要問我為什麼用 Range.Value ... 因為我忘了 :p
  iRow := 1;
  for iCol := 0 to aGrid.Columns.Count-1 do
  begin
    Cell := GetColA2Z(iCol+1) + IntToStr(iRow);
    aEXCEL.Cells.Range[Cell, Cell].Value := aGrid.Columns.Items[iCol].Title.Caption;
  end;

  myBookMark := aGrid.DataSource.DataSet.GetBookmark;
  aGrid.DataSource.DataSet.First;
  // 避免畫面死機及誤按, 都給他 Disable
  aGrid.DataSource.DataSet.DisableControls;
  aGrid.Enabled := False;

  // 輸出 [欄位內容]
  iRow := 2;
  while not aGrid.DataSource.DataSet.Eof do
  begin
    if TrimSpace then
      for iCol := 0 to aGrid.Columns.Count-1 do
      begin
        Cell := GetColA2Z(iCol+1) + IntToStr(iRow);
        aEXCEL.Cells.Range[Cell, Cell].Value := Trim(aGrid.Fields[iCol].AsString);
      end
    else
      for iCol := 0 to aGrid.Columns.Count-1 do
      begin
        Cell := GetColA2Z(iCol+1) + IntToStr(iRow);
        aEXCEL.Cells.Range[Cell, Cell].Value := aGrid.Fields[iCol].AsString;
      end;
    iRow := iRow + 1;
    aGrid.DataSource.DataSet.Next;
  end;
  aGrid.DataSource.DataSet.EnableControls;
  aGrid.DataSource.DataSet.GotoBookmark(myBookMark);
  aGrid.Enabled := True;
end;

// 主要 StringGridToExcel
// 最簡單了...
procedure TGridExport.StringGridToExcel(aEXCEL: TExcelApplication; aGrid: TStringGrid; TrimSpace: Boolean = False);
var
  iCol, iRow: integer;
  Cell: string;
begin
  if TrimSpace then
    for iRow := 0 to aGrid.RowCount - 1 do
      for iCol := 0 to aGrid.ColCount - 1 do
      begin
        Cell := GetColA2Z(iCol+1) + IntToStr(iRow+1);
        aEXCEL.Cells.Range[Cell, Cell].Value := Trim(aGrid.Cells[iCol, iRow]);
      end
  else
    for iRow := 0 to aGrid.RowCount - 1 do
      for iCol := 0 to aGrid.ColCount - 1 do
      begin
        Cell := GetColA2Z(iCol+1) + IntToStr(iRow+1);
        aEXCEL.Cells.Range[Cell, Cell].Value := aGrid.Cells[iCol, iRow];
      end;
end;

// 主角, assign 給 MenuItem 的 OnClick
// 執行 Export Grid 資料到 Excel 的動作
procedure TGridExport.Execute(Sender: TObject);
var
  ExcelApp: TExcelApplication;
  Template, NewTemplate, ItemIndex: OleVariant;
  myGrid: TComponent;
  TrimSpace: boolean;
begin
  // 啟動 Excel, 並開啟新的 WorkBook
  ExcelApp := TExcelApplication.Create(ParentForm);
  Template := EmptyParam;
  NewTemplate := True;
  ItemIndex := 0;
  try
    ExcelApp.Connect;
  except
    MessageDlg('Excel may not be installed !!', mtError, [mbOk], 0);
    Abort;
  end;
  ExcelApp.Workbooks.Add(Template, ItemIndex);

  if Pos('TRIM', UpperCase((Sender as TMenuItem).Name)) > 0 then
    TrimSpace := True
  else
    TrimSpace := False;
  // 從 Sender 的 Hint 輾轉得到是從哪個 Grid 來的
  // 因為不適用正常的 Popup 方式開啟 PopupMenu   ( dbgrid.PopupMenu assign )
  // 所以不能用 myGrid := ((Sender as TMenuItem).GetParentMenu as TPopupMenu).PopupComponent;
  // 來取得是從哪一個 Grid 來的 PopupMenu
  myGrid := ParentForm.FindComponent((Sender as TMenuItem).Hint);
  if myGrid.ClassName = 'TDBGrid' then
    DBGridToExcel(ExcelApp, (myGrid as TDBGrid), TrimSpace);
  if myGrid.ClassName = 'TStringGrid' then
    StringGridToExcel(ExcelApp, (myGrid as TStringGrid), TrimSpace);

  ExcelApp.Cells.EntireColumn.AutoFit;
  // 要先 Show 出 Excel, 或是跑完資料再 Show 在這裡決定
  ExcelApp.Visible[ItemIndex] := True;
  // 保險起見多關幾次, 有時程式執行中若自己中斷會... 所以 FreeAll 裡會再 Free 一次
  ExcelApp.Free;
end;

// Grid 上按右鍵出現 PopupMenu, 注意: 此 PopupMenu 及 MenuItem 是整的 form 共用的
procedure TGridExport.GridMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  myPopupMenu: TPopupMenu;
  myMenuItem, myMenuItem_Trim: TMenuItem;
  pMouse : TPoint;
begin
  if Button = mbRight then
  begin
    pMouse.X := x;
    pMouse.Y := y;
    pMouse := (Sender as TWinControl).ClientToScreen(pMouse);

    // 應用 DBGrid / StringGrid 本身已有設定的 PopMenu, 將延伸的 MenuItem 加入其中
    if Sender is TDBGrid then
      myPopupMenu := (Sender as TDBGrid).PopupMenu;
    if Sender is TStringGrid then
      myPopupMenu := (Sender as TStringGrid).PopupMenu;
    if myPopupMenu = nil then
      myPopupMenu := (ParentForm.FindComponent('pmGRID') as TPopupMenu);
    // 都找不到 PopupMenu 就自己生
    if myPopupMenu = nil then
    begin
      myPopupMenu := TPopupMenu.Create(ParentForm);
      myPopupMenu.Name := 'pmGRID';
    end;
    // 找不到 TMenuItem : miEXCEL 就自己生
    myMenuItem := (myPopupMenu.FindComponent('miEXCEL') as TMenuItem);
    if myMenuItem = nil then
    begin
      myMenuItem := TMenuItem.Create(myPopupMenu);
      myMenuItem.Name := 'miEXCEL';
      myMenuItem.Caption := '轉EXCEL';
      myMenuItem.AutoHotkeys := maManual;
      // 重點： Excute
      myMenuItem.OnClick := Execute;
      myMenuItem.Hint := (Sender as TWinControl).Name;
      myPopupMenu.Items.Add(myMenuItem);
    end
    else
      myMenuItem.Hint := (Sender as TWinControl).Name;
    // 找不到 TMenuItem : miEXCEL_Trim 就自己生
    myMenuItem_Trim := (myPopupMenu.FindComponent('miEXCEL_Trim') as TMenuItem);
    if myMenuItem_Trim = nil then
    begin
      myMenuItem_Trim := TMenuItem.Create(myPopupMenu);
      myMenuItem_Trim.Name := 'miEXCEL_Trim';
      myMenuItem_Trim.Caption := '轉EXCEL-清除空白';
      myMenuItem_Trim.AutoHotkeys := maManual;
      // 重點： Excute
      myMenuItem_Trim.OnClick := Execute;
      myMenuItem_Trim.Hint := (Sender as TWinControl).Name;
      myPopupMenu.Items.Add(myMenuItem_Trim);
    end
    else
      myMenuItem_Trim.Hint := (Sender as TWinControl).Name;

    // 上面的 myMenuItem.Hint 及 myMenuItem_Trim.Hint 是很重要的
    // 用來識別 "是在哪一個 Grid 上按右鍵", 要不然從 Execute 中的 Sender 得到的是 MenuItem
    // 而又無法取得 PopupComponent 的值, 所以...
    myPopupMenu.Popup(pMouse.X, pMouse.Y);
  end;
end;

// 按下 DBGrid Title 時, 自動做 Sorting 排序
// 限制: 1. 要已 Open,
//       2. 要使用 TADOQuery, (有需要請自改)
//       3. 按下非 fkData 的欄位不會動作
//
procedure TGridExport.GridTitleClick(Column: TColumn);
var
  I: integer;
  Found: boolean;
  mySTR, myColumn: string;
  myBookMark: TBookMark;
  myADOQuery: TADOQuery;
begin
  if (not (Column.Grid as TDBGrid).DataSource.DataSet.Active) or
     (not (Column.Grid as TDBGrid).DataSource.DataSet.ClassNameIs('TADOQuery')) or
     (Column.Field.FieldKind <> fkData) then
    Exit;
  // 幹麼不省掉這兩個變數 myADOQuery 和 myColumn? 其實之前舊版是沒有的
  // 可是在寫 Demo 時, 不拿 myADOQuery 來取代 ((Column.Grid as TDBGrid).DataSource.DataSet as TADOQuery)
  // 到後面 Open 時會 Access Violation
  // Column.FieldName 也會莫名其妙變成 "空白值", 不知道為什麼....
  myADOQuery := ((Column.Grid as TDBGrid).DataSource.DataSet as TADOQuery);
  myColumn := Column.FieldName;
  Found := False;
  for I := myADOQuery.SQL.Count - 1 downto 0 do
  begin
    mySTR := myADOQuery.SQL.Strings[I];
    if Pos('ORDER BY', UpperCase(mySTR)) > 0 then
    begin
      Found := True;
      myADOQuery.SQL.Delete(I);
      if Pos(UpperCase(myColumn), UpperCase(mySTR)) > 0 then
        if Pos('DESC', UpperCase(mySTR)) > 0 then
          mySTR := 'ORDER BY ' + myColumn
        else
          mySTR := 'ORDER BY ' + myColumn + ' DESC'
      else
        mySTR := 'ORDER BY ' + myColumn;
      Break;
    end;
  end;

  if not Found then
    mySTR := 'ORDER BY ' + Column.FieldName;

  myBookMark := myADOQuery.GetBookmark;
  myADOQuery.SQL.Add(Chr(13) + mySTR);
  myADOQuery.Close;
  myADOQuery.Open;
  myADOQuery.GotoBookmark(myBookMark);
end;

procedure Register;
begin
  RegisterComponents('Samples', [TGridExport]);
end;

end.
