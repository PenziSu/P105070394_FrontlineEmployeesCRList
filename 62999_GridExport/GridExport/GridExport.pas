unit GridExport;
// Written by Silence Wang ( silence@cmo.com.tw )
// Introduced: some day in 2003
// Last Modified: 2005/01/06
//
// �]���u�@�W�ϥΨ쪺, �������g���p�{���g�`�n��X Excel
// �ҥH�g�Ӥ����²��
//
// �\��G
//  1. DBGrid �� StringGrid �W���U�k���i�H��� [��X�� Excel]
//  2. �۰ʥͦ� PopupMenu / MenuItem, �þ�X�J�ۭq�� PopupMenu, �����B�~�дo�ާ@����
//  3. DBGrid �i�ϥΫ��U Title �ᰵ Sort ���\�� (�� DBGrid �� DataSet �n�O SQL Query, �B����ӽ���)
//  4. �b����ɴ��Y�n����, �����\ user ��X���, �u�n�N�S�w Grid �� OnMouseUp := nil �Y�i
//     �n�A���}, �u�n���s Initial �Y�i
//
// �ϥΡG
//  0. �N����ԩ�� Form �W
//  1. �b FormCreate ���[�J
//     GridExport.Initial(Self);
//     �Y�i�۰ʨϥ�, �B�~���w�]�� False ���Ѽ� aTitleSort: boolean
//  2. �b FormClose ���[�J
//     GridExport.FreeAll;
//     �T�O [�ۦ�����] �� Excel ���T Free
//  3. �{�������, �� DBGrid / StringGrid �W���k����X PopupMenu,
//     �䤤 MenuItem �|�۰ʸ�A���w�� Grid �� PopupMenu ���X
//  4. ���� Export to Excel, �q DBGrid / StringGrid �N��ƶץX�� Excel
//
//  ���a�禡�|��
//  A. �۰ʳ]�w String Grid ���� aCol �� column width
//  B. ���� A �Ӱ��h�� column ���۰ʽվ�
//  C. "�ۭq DBGrid ����r�~�[��, ��b OnDrawColumnCell ���̫�Өϥ�"
//  D. "�ۭq StringGrid ����r�~�[��, ��b OnDrawCell ���̫�Өϥ�"

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

// ���a�禡 A
// �]�w String Grid ���� aCol �� column width
// �o�Ө禡�w����w�� aCol �����r�����, ���� aMin �� aMax ���d��Ӱ��۰ʳ]�w
// aMin �� aMax �w�]�Ȭ� 0 �h��������
// �Ϊk: AutoSizeStringGridColumn(mySGrid, 2);
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

// ���a�禡 B
// ���� AutoSizeStringGridColumn, �u�n���w���� col �ݭn�վ�j�p�Y�i, �����ϥ� Min / Max �\��
// (���Q�g�o�ӽ���, �n�C�� column �Τ��P�� Min/Max, �N��W�ϥ� AutoSizeStringGridColumn �N�n�F)
// �Ϊk: SetStringGridWidth(mySGrid1, []);
//       SetStringGridWidth(mySGrid2, colArray);
//       SetStringGridWidth(mySGrid3, [0, 2], 0, 30);
procedure SetStringGridWidth(aStringGrid: TStringGrid; aCols: array of integer);
var
  iCol, I: integer;
begin
  // �p�� StrGrid �U�檺�e��
  if High(aCols) = -1 then
    // �S���w�N�����վ�
    for iCol := 0 to aStringGrid.ColCount - 1 do
      AutoSizeStringGridColumn(aStringGrid, iCol, 0, 0)
  else
    // �����w�N�v�@�վ�, �W�X�d�򪺤��z�|
    for I := Low(aCols) to High(aCols) do
      if (aCols[I] >= 0) and
         (aCols[I] < aStringGrid.ColCount) then
        AutoSizeStringGridColumn(aStringGrid, aCols[I], 0, 0);
end;

// ���a�禡 C
// �o�Ө禡�O�� "�ۭq DBGrid ����r�~�[��, ��b OnDrawColumnCell ���̫�Өϥ�"
// �]���ۭq DBGrid �� ColumnCell ����r�~�[��, �n�ۤv���� Text �� Alignment;
// �Ϊk:
//  OnDrawColumnCell ��
//    myCanvas := (Sender as TDBGrid).Canvas;
//    myRect := Rect;
//    ... �]�w Text Font �~�[
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

// ���a�禡 D
// �o�Ө禡�O�� "�ۭq StringGrid ����r�~�[��, ��b OnDrawCell ���̫�Өϥ�"
// ���u���@��Ӥw ++" , �٥i�H�X�R��, �Ҧp�a���a�k�a��������
// �Ϊk:
//  OnDrawCell ��
//    myRect := Rect;
//    myCanvas := Canvas;
//    myText := Cells[ACol, ARow];
//    ... �]�w Text Font �~�[
//    DrawStrColumn(myRect, myCanvas, myText);
//  end;
procedure DrawStrColumn(var aRect: TRect; var aCanvas: TCanvas; var aText: string);
begin
  aCanvas.TextRect(aRect, aRect.Left + 2, aRect.Top + 2, aText);
end;


{ TGridExport }

// �@�}�l�� "�Ұ�"
// �]�w Owner ���U�� Grid ( DBGrid / StringGrid )
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

// Release �M�Ϊ� TPopupMenu : pmGRID
// �������P Excel ���s�� ==���n==
// �{�������ɤ@�w�n����, �i�H�קK Excel �Ӧh Instance �Ϋ���}�� Excel �ɲ��͵������~
// (�N�O "�u���X�{ Menu �Τu��C, �� Sheet �]���X��" �����ؿ��~)
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

// ���W�L Z ����? �o�����A��ա�
function TGridExport.GetColA2Z(Col: Integer):string;
begin
  if (Col mod 26 = 0) then
    Result := 'Z'
  else
    Result := Chr((Col mod 26)+64);
  // ���j, �N��A ZZZ �ڤ]���A��
  if Col > 26 then
    Result := GetColA2Z(Col div 26) + Result;
end;

// �D�n DBGridToExcel
// TrimSpace? �Ψӥh�����ݪ� "�ť�" �Ϊ�, ������줣�O�� Varchar �� type, ���ݷ|�� "�ť�"
// �F���٭n��, �@�� Trim �����O�������? .... ���H�ε۴N�d�ۧa
procedure TGridExport.DBGridToExcel(aEXCEL: TExcelApplication; aGrid: TDBGrid; TrimSpace: Boolean = False);
var
  iCol, iRow: integer;
  Cell: string;
  myBookMark: TBookMark;
begin
  // ��X [Title]
  // ���n�ݧڬ������ Range.Value ... �]���ڧѤF :p
  iRow := 1;
  for iCol := 0 to aGrid.Columns.Count-1 do
  begin
    Cell := GetColA2Z(iCol+1) + IntToStr(iRow);
    aEXCEL.Cells.Range[Cell, Cell].Value := aGrid.Columns.Items[iCol].Title.Caption;
  end;

  myBookMark := aGrid.DataSource.DataSet.GetBookmark;
  aGrid.DataSource.DataSet.First;
  // �קK�e�������λ~��, �����L Disable
  aGrid.DataSource.DataSet.DisableControls;
  aGrid.Enabled := False;

  // ��X [��줺�e]
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

// �D�n StringGridToExcel
// ��²��F...
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

// �D��, assign �� MenuItem �� OnClick
// ���� Export Grid ��ƨ� Excel ���ʧ@
procedure TGridExport.Execute(Sender: TObject);
var
  ExcelApp: TExcelApplication;
  Template, NewTemplate, ItemIndex: OleVariant;
  myGrid: TComponent;
  TrimSpace: boolean;
begin
  // �Ұ� Excel, �ö}�ҷs�� WorkBook
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
  // �q Sender �� Hint ����o��O�q���� Grid �Ӫ�
  // �]�����A�Υ��`�� Popup �覡�}�� PopupMenu   ( dbgrid.PopupMenu assign )
  // �ҥH����� myGrid := ((Sender as TMenuItem).GetParentMenu as TPopupMenu).PopupComponent;
  // �Ө��o�O�q���@�� Grid �Ӫ� PopupMenu
  myGrid := ParentForm.FindComponent((Sender as TMenuItem).Hint);
  if myGrid.ClassName = 'TDBGrid' then
    DBGridToExcel(ExcelApp, (myGrid as TDBGrid), TrimSpace);
  if myGrid.ClassName = 'TStringGrid' then
    StringGridToExcel(ExcelApp, (myGrid as TStringGrid), TrimSpace);

  ExcelApp.Cells.EntireColumn.AutoFit;
  // �n�� Show �X Excel, �άO�]����ƦA Show �b�o�̨M�w
  ExcelApp.Visible[ItemIndex] := True;
  // �O�I�_���h���X��, ���ɵ{�����椤�Y�ۤv���_�|... �ҥH FreeAll �̷|�A Free �@��
  ExcelApp.Free;
end;

// Grid �W���k��X�{ PopupMenu, �`�N: �� PopupMenu �� MenuItem �O�㪺 form �@�Ϊ�
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

    // ���� DBGrid / StringGrid �����w���]�w�� PopMenu, �N������ MenuItem �[�J�䤤
    if Sender is TDBGrid then
      myPopupMenu := (Sender as TDBGrid).PopupMenu;
    if Sender is TStringGrid then
      myPopupMenu := (Sender as TStringGrid).PopupMenu;
    if myPopupMenu = nil then
      myPopupMenu := (ParentForm.FindComponent('pmGRID') as TPopupMenu);
    // ���䤣�� PopupMenu �N�ۤv��
    if myPopupMenu = nil then
    begin
      myPopupMenu := TPopupMenu.Create(ParentForm);
      myPopupMenu.Name := 'pmGRID';
    end;
    // �䤣�� TMenuItem : miEXCEL �N�ۤv��
    myMenuItem := (myPopupMenu.FindComponent('miEXCEL') as TMenuItem);
    if myMenuItem = nil then
    begin
      myMenuItem := TMenuItem.Create(myPopupMenu);
      myMenuItem.Name := 'miEXCEL';
      myMenuItem.Caption := '��EXCEL';
      myMenuItem.AutoHotkeys := maManual;
      // ���I�G Excute
      myMenuItem.OnClick := Execute;
      myMenuItem.Hint := (Sender as TWinControl).Name;
      myPopupMenu.Items.Add(myMenuItem);
    end
    else
      myMenuItem.Hint := (Sender as TWinControl).Name;
    // �䤣�� TMenuItem : miEXCEL_Trim �N�ۤv��
    myMenuItem_Trim := (myPopupMenu.FindComponent('miEXCEL_Trim') as TMenuItem);
    if myMenuItem_Trim = nil then
    begin
      myMenuItem_Trim := TMenuItem.Create(myPopupMenu);
      myMenuItem_Trim.Name := 'miEXCEL_Trim';
      myMenuItem_Trim.Caption := '��EXCEL-�M���ť�';
      myMenuItem_Trim.AutoHotkeys := maManual;
      // ���I�G Excute
      myMenuItem_Trim.OnClick := Execute;
      myMenuItem_Trim.Hint := (Sender as TWinControl).Name;
      myPopupMenu.Items.Add(myMenuItem_Trim);
    end
    else
      myMenuItem_Trim.Hint := (Sender as TWinControl).Name;

    // �W���� myMenuItem.Hint �� myMenuItem_Trim.Hint �O�ܭ��n��
    // �Ψ��ѧO "�O�b���@�� Grid �W���k��", �n���M�q Execute ���� Sender �o�쪺�O MenuItem
    // �ӤS�L�k���o PopupComponent ����, �ҥH...
    myPopupMenu.Popup(pMouse.X, pMouse.Y);
  end;
end;

// ���U DBGrid Title ��, �۰ʰ� Sorting �Ƨ�
// ����: 1. �n�w Open,
//       2. �n�ϥ� TADOQuery, (���ݭn�Цۧ�)
//       3. ���U�D fkData ����줣�|�ʧ@
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
  // �F�򤣬ٱ��o����ܼ� myADOQuery �M myColumn? ��ꤧ�e�ª��O�S����
  // �i�O�b�g Demo ��, ���� myADOQuery �Ө��N ((Column.Grid as TDBGrid).DataSource.DataSet as TADOQuery)
  // ��᭱ Open �ɷ| Access Violation
  // Column.FieldName �]�|���W�䧮�ܦ� "�ťխ�", �����D������....
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
