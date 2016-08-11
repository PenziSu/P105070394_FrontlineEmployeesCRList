unit Main_process;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls,Grids, DBGrids, RzDBGrid, ComObj,
  iworklst1_dm_u,
  iworklst_ndm_u,
  session_dm_u,
  shareutils,
  isysctl_ndm_u,
  ichart_dm_u,
  ixryasg_dm_u,
  ixryspc_dm_u,
  ixryct_dm_u,
  iempdept_dm_u,
  iemploye_ndm_u,
  ixryreg2_dm_u, Db, DBClient, ADODB, RzButton;

type
  TDBGrid = class(DBGrids.TDBGrid)
  public
    function DoMouseWheel(Shift: TShiftState; WheelDelta: Integer; MousePos: TPoint):Boolean; override;
  end;

type
  TForm1 = class(TForm)
    btnStart: TButton;
    dds1: TDataSource;
    cds1: TClientDataSet;
    con1: TADOConnection;
    qry1: TADOQuery;
    btnExportXls: TRzButton;
    dbgrd1: TDBGrid;
    lb4: TLabel;
    edFilepath: TEdit;
    lb6: TLabel;
    lb7: TLabel;
    lb8: TLabel;
    lb9: TLabel;
    lb10: TLabel;
    cbBeginYear: TComboBox;
    cbBeginMonth: TComboBox;
    cbEndYear: TComboBox;
    cbEndMonth: TComboBox;
    grp1: TGroupBox;
    lb11: TLabel;
    lb12: TLabel;
    lb13: TLabel;
    lb2: TLabel;
    lbExamed: TLabel;
    lbunexam: TLabel;
    lb1: TLabel;
    lb3: TLabel;
    lb5: TLabel;
    procedure btnStartClick(Sender: TObject);
    procedure OPEN_FILE;
    procedure CLOSE_FILE;
    function IDNo_To_Department(ID:string):string;
    function IDNo_To_EmployeeID(ID:string):string;
    function IDNo_To_ExtNumber(ID:string):string;
    procedure FormCreate(Sender: TObject);
    procedure btnExportXlsClick(Sender: TObject);
    procedure DBGrid2Excel(DBGrid:TDBGrid;ExcelFileName:string);
    procedure cbBeginYearChange(Sender: TObject);
    procedure cbEndYearChange(Sender: TObject);
  private
    { Private declarations }
    function ReadIXRYASG2(DateFlag:string):Boolean;
    function WriteDataIntoDataset(IDNO,StuffName:string;Stuff_PID,ExamDate,OrderDate:Integer;AID,ImageFlag:string):Boolean;
    function GetIWORKLSTDataByAccessionNumber(AID:string):Boolean;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  IWL  : IWORKLST_REC;
  IWL1 : IWORKLST1_REC;
  IXA  : ixryasg_rec;
  IXSP : IXRYSPC_REC;
  IXCT : IXRYCT_REC;
  IXR2 : IXRYREG2_REC;
  IEDP : IEMPDEPT_REC;
  IEM  : IEMPLOYE_REC;
  sn           : Integer;
  exam_count   : Integer;
  unexam_count : Integer;

implementation


{$R *.DFM}

procedure TForm1.OPEN_FILE;
begin
  INIT_IWORKLST(IWL);
  IWL.FD  := OPENFILE(IWL.DRIVE, IWL.FNAME, 'INOUT');

  INIT_IWORKLST1(IWL1);
  IWL1.FD := OPENFILE(IWL1.DRIVE,IWL1.FNAME,'INOUT');

  INIT_IXRYCT(IXCT);
  IXCT.fd := openfile(IXCT.DRIVE,IXCT.FNAME,'INOUT');

  INIT_Ixryasg(IXA);
  IXA.FD  := OPENFILE(Ixa.DRIVE, Ixa.FNAME, 'INOUT');

  INIT_IXRYSPC(IXSP);
  IXSP.FD := OPENFILE(IXSP.DRIVE,IXSP.FNAME,'INOUT');

  INIT_IXRYREG2(IXR2);
  IXR2.FD := OPENFILE(IXR2.DRIVE,IXR2.FNAME,'INPUT');

  INIT_IEMPDEPT(IEDP);
  IEDP.FD := OPENFILE(IEDP.DRIVE,IEDP.FNAME,'INPUT');

  INIT_IEMPLOYE(IEM);
  IEM.FD := OPENFILE(IEM.DRIVE,IEM.FNAME,'INPUT');
end;

procedure TForm1.CLOSE_FILE;
begin
  CLOSFILE(IWL.DRIVE,IWL.FD);
  CLOSFILE(IWL1.DRIVE,IWL1.FD);
  CLOSFILE(IXA.drive,ixa.FD);
  CLOSFILE(Ixct.DRIVE,Ixct.FD);
  CLOSFILE(IXSP.drive,IXSP.FD);
  CLOSFILE(IXR2.drive,IXR2.FD);
  CLOSFILE(IEDP.drive,IEDP.FD);
  CLOSFILE(IEM.drive,IEM.FD);
end;

function TDBGrid.DoMouseWheel(Shift: TShiftState; WheelDelta: Integer; MousePos: TPoint):Boolean;
begin
  if WheelDelta < 0 then DataSource.DataSet.MoveBy(3);
  if WheelDelta > 0 Then DataSource.DataSet.MoveBy(-3);
end;


procedure TForm1.btnStartClick(Sender: TObject);
var
  thisMonth,DateFlag,StartDate,EndDate : string;
  x,y : integer;
begin
  StartDate := cbBeginYear.text+cbBeginMonth.text;
  DateFlag  := cbBeginYear.text+cbBeginMonth.text;
  EndDate   := cbEndYear.text+cbEndMonth.text;
  thisMonth := formatdatetime('eeemm',now);
  btnStart.Enabled := False;
  unexam_count := 0;
  exam_count := 0;
  sn := 0;
  while EndDate < StartDate do
  begin
    ShowMessage('日期範圍從: '+Copy(StartDate,1,3)+'年'+Copy(StartDate,4,2)+'月 到 '+Copy(EndDate,1,3)+'年'+Copy(EndDate,4,2)+'月');
    ShowMessage('你確定？');
    btnStart.Enabled := True;
    Exit;
  end;

  OPEN_FILE;
  if cds1.Active = False then cds1.CreateDataSet;
  cds1.EmptyDataSet;
  try
    for x := StrToInt(Copy(StartDate,1,3)) to StrToInt(Copy(EndDate,1,3)) do
    begin
      if (StrToInt(Copy(StartDate,1,3)) = x) and (x = StrToInt(Copy(EndDate,1,3))) then
        begin
          {從開始月份加到結束月}
          for y := StrToInt(Copy(StartDate,4,2)) to StrToInt(Copy(EndDate,4,2)) do
          begin
            //ShowMessage(IntToStr(x)+format('%.2d',[y]));
            ReadIXRYASG2(IntToStr(x)+format('%.2d',[y]));
          end;
        end
      else if (StrToInt(Copy(StartDate,1,3)) = x) and (x < StrToInt(Copy(EndDate,1,3)))then
        begin
          {從開始月份加到12月}
          for y := StrToInt(Copy(StartDate,4,2)) to 12 do
          begin
            //ShowMessage(IntToStr(x)+format('%.2d',[y]));
            ReadIXRYASG2(IntToStr(x)+format('%.2d',[y]));
          end;
        end
      else if (StrToInt(Copy(StartDate,1,3)) < x) and (x < StrToInt(Copy(EndDate,1,3)))then
        begin
          {從1月份加到12月}
          for y := 1 to 12 do
          begin
            //ShowMessage(IntToStr(x)+format('%.2d',[y]));
            ReadIXRYASG2(IntToStr(x)+format('%.2d',[y]));
          end;
        end
      else if (StrToInt(Copy(StartDate,1,3)) < x) and (x = StrToInt(Copy(EndDate,1,3)))then
        begin
          {從1月份加到結束月}
          for y := 1 to StrToInt(Copy(EndDate,4,2)) do
          begin
            //ShowMessage(IntToStr(x)+format('%.2d',[y]));
            ReadIXRYASG2(IntToStr(x)+format('%.2d',[y]));
          end;
        end
      else
        begin
          ShowMessage('What the FxxxK: '+IntToStr(x));
        end;
    end;
  finally
    cds1.First;
    btnStart.Enabled := True;
    con1.Connected   := False;
    lb2.Caption      := IntToStr(SN);
    lbExamed.Caption := IntToStr(exam_count);
    lbunexam.Caption := IntToStr(unexam_count);
    if cds1.Active = True then
    begin
      btnExportXls.Enabled := True;
    end;
    CLOSE_FILE;
  end;
end;

function TForm1.ReadIXRYASG2(DateFlag:string):Boolean;
var
  BeginDate,EndDate : string;
begin
  //ZERO_IXRYREG2(IXR2);
  con1.LoginPrompt := False;
  con1.Connected   := True;
  BeginDate := DateFlag +'01';
  EndDate   := DateFlag +'31';
  try
    if DateFlag = FormatDateTime('eeemm',now) then
    begin
      (*如果傳入的DateFlag與程式執行月份相同直接讀檔*)
      IXR2.DRIVE := 'MAS';
      IXR2.FNAME := 'IXRYREG2';
      IXR2.FD := OPENFILE(IXR2.DRIVE,IXR2.FNAME,'INOUT');
      IXR2.CODE := 'h32001';
      IXR2.DATE := StrToInt(BeginDate);
      SETKEYNO(ixr2.drive,IXR2.FD,2);
      READ_IXRYREG2(IXR2,27);
      while (IXR2.ERR = 0) and (IXR2.DATE < StrToInt(EndDate)) do
      begin
        //showmessage(IXR2.ACCESS_NO);
        GetIWORKLSTDataByAccessionNumber(IXR2.ACCESS_NO);
        READ_IXRYREG2(IXR2,2);
      end;
    end else
    begin
      (*如果傳入的DateFlag與程式執行月份不同就讀月檔*)
      IXR2.DRIVE := 'LAST';
      IXR2.FNAME := 'MXRY2;MXRY2' + Copy(DateFlag,3,3);
      IXR2.FD := OPENFILE(IXR2.DRIVE,IXR2.FNAME,'INOUT');
      IXR2.CODE := 'h32001';
      IXR2.DATE := StrToInt(BeginDate);
      SETKEYNO(ixr2.drive,IXR2.FD,2);
      READ_IXRYREG2(IXR2,27);
      while (IXR2.ERR = 0) and (IXR2.DATE < StrToInt(EndDate)) do
      begin
        //showmessage(IXR2.ACCESS_NO);
        GetIWORKLSTDataByAccessionNumber(IXR2.ACCESS_NO);
        READ_IXRYREG2(IXR2,2);
      end;
    end;
  except
    on E : Exception do ShowMessage(E.Message);
  end;
end;

function TForm1.GetIWORKLSTDataByAccessionNumber(AID:string):Boolean;
begin
  Result := True;
  IWL.ACCESS_NO := AID;
  READ_IWORKLST_COMP15(IWL);
  if (IWL.ERR = 0) and (iwl.ORDER_DR = 'i80') and (IXR2.CODE = 'h32001') then
  begin
    (*計算已檢查與未檢查人數*)
    if (IWL.MPPS_DATE = 0)
    then Inc(unexam_count)
    else Inc(exam_count);
    (*將資料丟到DataSet*)
    WriteDataIntoDataset(IWL.ID_NO, IWL.PT_NAME, IWL.CHART_NO, IWL.MPPS_DATE, IWL.ORDER_DATE, IWL.ACCESS_NO, IWL.HAVE_IMAGE);
    Inc(sn);
  end;
end;

function TForm1.WriteDataIntoDataset(IDNO,StuffName:string;Stuff_PID,ExamDate,OrderDate:Integer;AID,ImageFlag:string):Boolean;
begin
  Result := true;
  cds1.Append;
  cds1.FieldByName('單位').value     := IDNo_To_Department(IDNO);
  cds1.FieldByName('員工編號').value := IDNo_To_EmployeeID(IDNO);
  cds1.FieldByName('員工姓名').value := StuffName;
  cds1.FieldByName('分機號碼').value := IDNo_To_ExtNumber(IDNO);
  cds1.FieldByName('病歷號').value   := IntToStr(Stuff_PID);
  cds1.FieldByName('檢查日期').value := Format('%u',[ExamDate]);
  cds1.FieldByName('開單日期').value := Format('%u',[OrderDate]);
  cds1.FieldByName('檢查單號').value := AID;
  cds1.FieldByName('影像註記').value := ImageFlag;
  cds1.Post;
end;


function TForm1.IDNo_To_Department(ID:string):string;
begin
  IEM.ID_NO := ID;
  SETKEYNO(IEM.DRIVE,IEM.FD,3);
  READ_IEMPLOYE(IEM,35);
  if (IEM.ERR = 0) and (IEM.ID_NO = ID) then
  begin
    IEDP.DEPT_NO := IEM.DEPT_NO;
    SETKEYNO(IEDP.DRIVE,IEDP.FD,1);
    READ_IEMPDEPT(IEDP,15);
    READ_IEMPDEPT_COMP15(IEDP);
    if (IEDP.ERR = 0) and (IEDP.DEPT_NO = IEM.DEPT_NO) then
    begin
      Result := IEDP.NAME;
    end;
  end else
  begin
    Result := '部門不明';
  end;
end;

function TForm1.IDNo_To_EmployeeID(ID:string):string;
begin
  IEM.ID_NO := ID;
  SETKEYNO(IEM.DRIVE,IEM.FD,3);
  READ_IEMPLOYE(IEM,35);
  if (IEM.ERR = 0) and (IEM.ID_NO = ID) then
  begin
    result := IEM.Employee_no;
  end else
  begin
    Result := '員編不明';
  end;
end;

function TForm1.IDNo_To_ExtNumber(ID:string):string;
begin
  qry1.Close;
  qry1.Active := False;
  qry1.sql.Clear;
  qry1.SQL.Text := 'select 分機 from 部門員工職稱_含留職停薪 where 身份證字號=:EmployeeID';
  qry1.Parameters.ParamByName('EmployeeID').Value := ID;
  qry1.Active := True;
  Result := qry1.FieldByName('分機').AsString;
  qry1.Close;
end;

procedure TForm1.FormCreate(Sender: TObject);
var
  Months,Years            : TStringList;
  MonthsCount,YearsCount  : integer;
begin
  try
    try
      Years := TStringList.Create;
      for YearsCount := 101 to StrToInt(formatdatetime('eee',now)) do
      begin
        //ShowMessage(IntToStr(MonthsCount));
        Years.Add(Format('%.3d',[YearsCount]));
      end;

      Months := TStringList.Create;
      for MonthsCount := 1 to StrToInt(formatdatetime('mm',now)) do
      begin
        //ShowMessage(IntToStr(MonthsCount));
        Months.Add(Format('%.2d',[MonthsCount]));
      end;
      cbBeginYear.Items  := Years;
      cbEndYear.Items    := Years;
      cbBeginMonth.Items := Months;
      cbEndMonth.Items   := Months;
      cbBeginYear.Text   := formatdatetime('eee',now);
      cbBeginMonth.text  := formatdatetime('mm',now);
      cbEndYear.Text     := formatdatetime('eee',now);
      cbEndMonth.Text    := formatdatetime('mm',now);
      if cds1.Active = False then
      begin
        btnExportXls.Enabled := False;
      end;
    except
      On e : Exception
      Do ShowMessage(e.Message);
    end;
  finally
    lb2.Caption      := '';
    lbExamed.Caption := '';
    lbunexam.caption := '';
  end;
end;

procedure TForm1.cbBeginYearChange(Sender: TObject);
var
  Months : TStringList;
  MonthsCount : Integer;
begin
  (*如果選擇的是當年，月份就要重新計算到當月*)
  (*如果不是當年，月份就從一月到十二月*)
  Months := TStringList.Create;
  try
    if cbBeginYear.Text = FormatDateTime('eee',now) then
      begin
        for MonthsCount := 1 to StrToInt(formatdatetime('mm',now)) do
        begin
          //ShowMessage(IntToStr(MonthsCount));
          Months.Add(Format('%.2d',[MonthsCount]));
        end;
      end
    else if (cbBeginYear.Text < FormatDateTime('eee',now)) then
      begin
        for MonthsCount := 1 to 12 do
        begin
          //ShowMessage(IntToStr(MonthsCount));
          Months.Add(Format('%.2d',[MonthsCount]));
        end;
      end
    else if (FormatDateTime('eee',now) < cbBeginYear.Text) then
      begin
        ShowMessage('不可能啦，你有事嗎?');
      end;
    cbBeginMonth.Items := Months;
    Months.Free;
  except
    On e : Exception
    Do ShowMessage(e.Message);
  end;
end;

procedure TForm1.cbEndYearChange(Sender: TObject);
var
  Months : TStringList;
  MonthsCount : Integer;
begin
  (*如果選擇的是當年，月份就要重新計算到當月*)
  (*如果不是當年，月份就從一月到十二月*)
  Months := TStringList.Create;
  try
    if (cbEndYear.Text = FormatDateTime('eee',now)) then
      begin
        for MonthsCount := 1 to StrToInt(formatdatetime('mm',now)) do
        begin
          //ShowMessage(IntToStr(MonthsCount));
          Months.Add(Format('%.2d',[MonthsCount]));
        end;
      end
    else if (cbEndYear.Text < FormatDateTime('eee',now)) then
      begin
        for MonthsCount := 1 to 12 do
        begin
          //ShowMessage(IntToStr(MonthsCount));
          Months.Add(Format('%.2d',[MonthsCount]));
        end;
      end
    else if (FormatDateTime('eee',now) < cbEndYear.Text) then
      begin
        ShowMessage('不可能啦，你有事嗎?');
      end;
    cbEndMonth.Items := Months;
    Months.Free;
  except
    On e : Exception
    Do ShowMessage(e.Message);
  end;
end;

procedure TForm1.btnExportXlsClick(Sender: TObject);
begin
  DBGrid2Excel(dbgrd1,edFilepath.text);
end;

(*程式來源 http://delphi.ktop.com.tw*)
procedure TForm1.DBGrid2Excel(DBGrid:TDBGrid;ExcelFileName:string);
var  MyExcel: Variant;
     x,y:integer;
begin
  deletefile(ExcelFileName);
  MyExcel := CreateOleOBject('Excel.Application');
  // 這一段為會員suda幫忙修正的
  MyExcel.WorkBooks.Add;
  MyExcel.Visible := True;
  //MyExcel.WorkBooks[1].Saveas(ExcelFileName);
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
  MyExcel.WorkBooks[1].Saveas(ExcelFileName);
end;

end.
