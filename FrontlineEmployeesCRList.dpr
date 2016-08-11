program FrontlineEmployeesCRList;

uses
  Forms,
  Main_process in 'Main_process.pas' {MainForm},
  iworklst1_dm_u in '..\..\dm\iworklst1_dm_u.pas' {IWORKLST1_dm},
  iworklst_ndm_u in '..\..\dm\iworklst_ndm_u.pas',
  session_dm_u in '..\..\dm\session_dm_u.pas' {session_dm},
  ShareUtils in '..\..\public\shareutils.pas',
  isysctl_ndm_u in '..\..\dm\isysctl_ndm_u.pas',
  ichart_ndm_u in '..\..\dm\ichart_ndm_u.pas',
  ichart_dm_u in '..\..\dm\ichart_dm_u.pas' {ICHART_dm},
  idisicd_dm_u in '..\..\dm\idisicd_dm_u.pas' {IDISICD_dm},
  ipatopd_dm_u in '..\..\dm\ipatopd_dm_u.pas' {IPATOPD_dm},
  ireg_dm_u in '..\..\dm\ireg_dm_u.pas' {IREG_dm},
  icanasoph_dm_u in '..\..\dm\icanasoph_dm_u.pas' {ICANASOPH_dm},
  icaliver_dm_u in '..\..\dm\icaliver_dm_u.pas' {ICALIVER_dm},
  icalung_dm_u in '..\..\dm\icalung_dm_u.pas' {ICALUNG_dm},
  icabreast_dm_u in '..\..\dm\icabreast_dm_u.pas' {ICABREAST_dm},
  icaoralcav_dm_u in '..\..\dm\icaoralcav_dm_u.pas' {ICAORALCAV_dm},
  icacolore_dm_u in '..\..\dm\icacolore_dm_u.pas' {ICACOLORE_dm},
  icacervica_dm_u in '..\..\dm\icacervica_dm_u.pas' {ICACERVICA_dm},
  icagastric_dm_u in '..\..\dm\icagastric_dm_u.pas' {ICAGASTRIC_dm},
  icaesoph_dm_u in '..\..\dm\icaesoph_dm_u.pas',
  icabladder_dm_u in '..\..\dm\icabladder_dm_u.pas',
  icaovarian_dm_u in '..\..\dm\icaovarian_dm_u.pas' {ICAOVARIAN_dm},
  icauterus_dm_u in '..\..\dm\icauterus_dm_u.pas',
  icaprostat_dm_u in '..\..\dm\icaprostat_dm_u.pas',
  icalymph_dm_u in '..\..\dm\icalymph_dm_u.pas',
  icalntreat_dm_u in '..\..\dm\icalntreat_dm_u.pas' {ICALNTREAT_dm},
  ixryasg_dm_u in '..\..\dm\ixryasg_dm_u.pas' {IXRYASG_dm},
  ixryspc_dm_u in '..\..\dm\ixryspc_dm_u.pas' {IXRYSPC_dm},
  ixryct_dm_u in '..\..\dm\ixryct_dm_u.pas' {IXRYCT_dm},
  ixryreg2_dm_u in '..\..\dm\ixryreg2_dm_u.pas' {IXRYREG2_dm},
  iempdept_dm_u in '..\..\dm\iempdept_dm_u.pas' {IEMPDEPT_dm},
  iemploye_ndm_u in '..\..\dm\iemploye_ndm_u.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.Title := '一線員工X光檢查清單';
  Application.CreateForm(TMainForm, MainForm);
  Application.CreateForm(TIWORKLST1_dm, IWORKLST1_dm);
  Application.CreateForm(Tsession_dm, session_dm);
  Application.CreateForm(TICHART_dm, ICHART_dm);
  Application.CreateForm(TIDISICD_dm, IDISICD_dm);
  Application.CreateForm(TIPATOPD_dm, IPATOPD_dm);
  Application.CreateForm(TIREG_dm, IREG_dm);
  Application.CreateForm(TICANASOPH_dm, ICANASOPH_dm);
  Application.CreateForm(TICALIVER_dm, ICALIVER_dm);
  Application.CreateForm(TICALUNG_dm, ICALUNG_dm);
  Application.CreateForm(TICABREAST_dm, ICABREAST_dm);
  Application.CreateForm(TICAORALCAV_dm, ICAORALCAV_dm);
  Application.CreateForm(TICACOLORE_dm, ICACOLORE_dm);
  Application.CreateForm(TICACERVICA_dm, ICACERVICA_dm);
  Application.CreateForm(TICAGASTRIC_dm, ICAGASTRIC_dm);
  Application.CreateForm(TICAOVARIAN_dm, ICAOVARIAN_dm);
  Application.CreateForm(TICALNTREAT_dm, ICALNTREAT_dm);
  Application.CreateForm(TIXRYASG_dm, IXRYASG_dm);
  Application.CreateForm(TIXRYSPC_dm, IXRYSPC_dm);
  Application.CreateForm(TIXRYCT_dm, IXRYCT_dm);
  Application.CreateForm(TIXRYREG2_dm, IXRYREG2_dm);
  Application.CreateForm(TIEMPDEPT_dm, IEMPDEPT_dm);
  Application.Run;
end.
