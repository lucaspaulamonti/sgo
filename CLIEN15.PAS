unit Clien15;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ModRel, QRExport, Menus, AtualPrintDialog, QRPDFFilt, StdCtrls,
  ExtCtrls, ComCtrls, Buttons, ToolWin, QRCtrls, QuickRpt, DBCtrls, Mask,
  CheckLst, NumEdit, DB, DBClient, SmartPointerClass, pFIBDataSet;

type
  TfClien15 = class(TfModRel)
    LBDtInicial: TLabel;
    MEDtInicial: TMaskEdit;
    DTPDtInicial: TDateTimePicker;
    LBDtFinal: TLabel;
    MEDtFinal: TMaskEdit;
    DTPDtFinal: TDateTimePicker;
    LBUsuario: TLabel;
    ComboUsuario: TDBLookupComboBox;
    PNLTop: TPanel;
    PNLMiddle: TPanel;
    ComboProduto: TDBLookupComboBox;
    EDProduto: TEdit;
    LBProduto: TLabel;
    MEProduto: TMaskEdit;
    ComboServico: TDBLookupComboBox;
    EDServico: TEdit;
    LBServico: TLabel;
    MEServico: TMaskEdit;
    LBGrupo: TLabel;
    PNLBottom: TPanel;
    BTAjuda: TBitBtn;
    CLBItens: TCheckListBox;
    LBItens: TLabel;
    NEUsuario: TNumberEdit;
    CBSelecionarTodos: TCheckBox;
    QRDetailBand: TQRBand;
    QLUsuarioDetail: TQRLabel;
    QLAcaoDetail: TQRLabel;
    QLLocalDetail: TQRLabel;
    QLDataDetail: TQRLabel;
    QLRegistroDetail: TQRMemo;
    ComboCliente: TDBLookupComboBox;
    EDCliente: TEdit;
    MECliente: TMaskEdit;
    LBCliente: TLabel;
    QRBand1: TQRBand;
    ChildBand2: TQRChildBand;
    QLUsuarioHeader: TQRLabel;
    QLAcaoHeader: TQRLabel;
    QLLocalHeader: TQRLabel;
    QLDataHeader: TQRLabel;
    QLRegistroHeader: TQRLabel;
    QLFiltros: TQRMemo;
    QLFiltros2: TQRMemo;
    BTExportar: TButton;
    ComboGrupo: TDBLookupComboBox;
    NEGrupo: TNumberEdit;

    procedure MEDtInicialExit(Sender: TObject);
    procedure DTPDtInicialChange(Sender: TObject);
    procedure MEProdutoExit(Sender: TObject);
    procedure MEServicoExit(Sender: TObject);
    procedure NEUsuarioExit(Sender: TObject);
    procedure MEClienteExit(Sender: TObject);
    procedure NEGrupoExit(Sender: TObject);
    procedure ComboProdutoCloseUp(Sender: TObject);
    procedure ComboServicoCloseUp(Sender: TObject);
    procedure ComboUsuarioCloseUp(Sender: TObject);
    procedure ComboClienteCloseUp(Sender: TObject);
    procedure ComboGrupoCloseUp(Sender: TObject);
    procedure EDProdutoKeyPress(Sender: TObject; var Key: Char);
    procedure EDServicoKeyPress(Sender: TObject; var Key: Char);
    procedure EDClienteKeyPress(Sender: TObject; var Key: Char);
    procedure CBSelecionarTodosClick(Sender: TObject);

    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);

    procedure BTVideoClick(Sender: TObject);

  private
    // Armazena o estado das queries.
    LUsuario, LCliente, LProduto, LServico, LGrupo: Boolean;
    // Adiciona as localizações no Checklistbox.
    procedure AddOnCheckListBox;
    // Faz as validações iniciais para garantir que os filtros estão corretos.
    procedure InitialValidations;
    // Abre/Fecha todas as consultas ao DtGeral que são utilizadas.
    procedure ToogleQueryActive;
    procedure ToogleQueryInactive;
    // Adiciona a data atual no componente TDateTimePicker.
    procedure GetCurrentDate;
    // Monta o WHERE da Query Principal do relatório.
    function Data: String;
    function Usuario: String;
    function Cliente: String;
    function Produto: String;
    function Servico: String;
    function Grupo: String;
    function Local: String;

  public

  end;

var
  fClien15: TfClien15;

implementation

uses
  DtGeral, DtMod, ExportaRel;

{$R *.dfm}

procedure TfClien15.AddOnCheckListBox;
var
  fQry: TSmartPointer<TpFIBDataSet>;
begin
  fQry.Value.Database := fDtMod.Dados;
  fQry.Value.Transaction := fDtMod.TransDados;

  with fQry.Value.SelectSQL do
  begin
    Clear;
    Add('SELECT ALIASTABELA FROM LOGTABELA');
  end;
  fQry.Value.Open;

  while not fQry.Value.EOF do
  begin
    CLBItens.Items.Add(fQry.Value.FieldByName('ALIASTABELA').AsString);
  end;
end;

procedure TfClien15.BTVideoClick(Sender: TObject);
begin
  // Faz as validações iniciais para garantir que os filtros estão corretos.
  InitialValidations;

  // Monta a Query Principal do Relatório.
  fDtMod.QryRel.Close;
  with fDtMod.QryRel.SelectSQL do
  begin
    Clear;
    Add('SELECT');
    Add('     LTPID, LTPTABELA, LTPUSUCODIGO, USUNOME, LTPDATAHORA, LTPLOCAL, LTPACAO,');
    Add('     LTPCAMPO1, LTPANTVALOR1, LTPNOVOVALOR1, LTPCAMPO2, LTPANTVALOR2, LTPNOVOVALOR2,');
    Add('     LTPCAMPO3, LTPANTVALOR3, LTPNOVOVALOR3, LTPCAMPO4, LTPANTVALOR4, LTPNOVOVALOR4,');
    Add('     LTPCAMPO5, LTPANTVALOR5, LTPNOVOVALOR5, LTPCAMPO6, LTPANTVALOR6, LTPNOVOVALOR6,');
    Add('     LTPCAMPO7, LTPANTVALOR7, LTPNOVOVALOR7, LTPCAMPO8, LTPANTVALOR8, LTPNOVOVALOR8,');
    Add('     LTPCAMPO9, LTPANTVALOR9, LTPNOVOVALOR9, LTPCAMPO10, LTPANTVALOR10, LTPNOVOVALOR10,');
    Add('     LTPCAMPO11, LTPANTVALOR11, LTPNOVOVALOR11, LTPCAMPO12, LTPANTVALOR12, LTPNOVOVALOR12,');
    Add('     LTPCAMPO13, LTPANTVALOR13, LTPNOVOVALOR13');
    Add('FROM');
    Add('    LOGTABELAPRECO');
    Add('    LEFT JOIN USUARIO ON LOGTABELAPRECO.LTPUSUCODIGO = USUARIO.USUCODIGO');
    Add('WHERE');
    Add(Data); Add(Usuario); Add(Local); Add(Cliente); Add(Produto); Add(Servico); Add(Grupo);
    Add('ORDER BY');
    Add('    LTPID');
  end;
  fDtMod.QryRel.Prepare;
  fDtMod.QryRel.ParamByName('DATAINICIAL').AsDateTime := StrToDateTime(MEDtInicial.Text + ' 00:00:00');
  fDtMod.QryRel.ParamByName('DATAFINAL').AsDateTime := StrToDateTime(MEDtFinal.Text + ' 23:59:59');
  fDtMod.QryRel.Open;
end;

procedure TfClien15.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  inherited;
  // Abre/Fecha todas as consultas ao DtGeral que são utilizadas.
  ToogleQueryInactive;
end;

procedure TfClien15.FormCreate(Sender: TObject);
begin
  inherited;
  // Abre/Fecha todas as consultas ao DtGeral que são utilizadas.
  ToogleQueryActive;
  // Adiciona a data atual no componente TDateTimePicker.
  GetCurrentDate;
  // Adiciona as localizações no Checklistbox.
  AddOnCheckListBox;
end;

procedure TfClien15.CBSelecionarTodosClick(Sender: TObject);
begin
  inherited;
  if CBSelecionarTodos.Checked then
    CLBItens.CheckAll(CbChecked)
  else
    CLBItens.CheckAll(CbUnchecked);
end;

function TfClien15.Cliente: String;
begin
  if (MECliente.Text <> '') and (MECliente.Text <> '0') then
    Result :=
    'AND(' + #13#10 +
    '        ((LTPCAMPO1 CONTAINING ''CLICODIGO'') AND ((LTPANTVALOR1 = ' + QuotedStr(MECliente.Text) + ') OR (LTPNOVOVALOR1 = ' + QuotedStr(MECliente.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO2 CONTAINING ''CLICODIGO'') AND ((LTPANTVALOR2 = ' + QuotedStr(MECliente.Text) + ') OR (LTPNOVOVALOR2 = ' + QuotedStr(MECliente.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO3 CONTAINING ''CLICODIGO'') AND ((LTPANTVALOR3 = ' + QuotedStr(MECliente.Text) + ') OR (LTPNOVOVALOR3 = ' + QuotedStr(MECliente.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO4 CONTAINING ''CLICODIGO'') AND ((LTPANTVALOR4 = ' + QuotedStr(MECliente.Text) + ') OR (LTPNOVOVALOR4 = ' + QuotedStr(MECliente.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO5 CONTAINING ''CLICODIGO'') AND ((LTPANTVALOR5 = ' + QuotedStr(MECliente.Text) + ') OR (LTPNOVOVALOR5 = ' + QuotedStr(MECliente.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO6 CONTAINING ''CLICODIGO'') AND ((LTPANTVALOR6 = ' + QuotedStr(MECliente.Text) + ') OR (LTPNOVOVALOR6 = ' + QuotedStr(MECliente.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO7 CONTAINING ''CLICODIGO'') AND ((LTPANTVALOR7 = ' + QuotedStr(MECliente.Text) + ') OR (LTPNOVOVALOR7 = ' + QuotedStr(MECliente.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO8 CONTAINING ''CLICODIGO'') AND ((LTPANTVALOR8 = ' + QuotedStr(MECliente.Text) + ') OR (LTPNOVOVALOR8 = ' + QuotedStr(MECliente.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO9 CONTAINING ''CLICODIGO'') AND ((LTPANTVALOR9 = ' + QuotedStr(MECliente.Text) + ') OR (LTPNOVOVALOR9 = ' + QuotedStr(MECliente.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO10 CONTAINING ''CLICODIGO'') AND ((LTPANTVALOR10 = ' + QuotedStr(MECliente.Text) + ') OR (LTPNOVOVALOR10 = ' + QuotedStr(MECliente.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO11 CONTAINING ''CLICODIGO'') AND ((LTPANTVALOR11 = ' + QuotedStr(MECliente.Text) + ') OR (LTPNOVOVALOR11 = ' + QuotedStr(MECliente.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO12 CONTAINING ''CLICODIGO'') AND ((LTPANTVALOR12 = ' + QuotedStr(MECliente.Text) + ') OR (LTPNOVOVALOR12 = ' + QuotedStr(MECliente.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO13 CONTAINING ''CLICODIGO'') AND ((LTPANTVALOR13 = ' + QuotedStr(MECliente.Text) + ') OR (LTPNOVOVALOR13 = ' + QuotedStr(MECliente.Text) + ')))' + #13#10 +
    ')'+ #13#10;
end;

procedure TfClien15.ComboClienteCloseUp(Sender: TObject);
begin
  inherited;
  if ComboCliente.Text <> '' then
    MECliente.Text := ComboCliente.KeyValue;

  EDCliente.Text := ComboCliente.Text;
  EDCliente.SetFocus;
  ComboCliente.Visible := False;

  fDtGeral.QryCliNome.Close;
end;

procedure TfClien15.ComboGrupoCloseUp(Sender: TObject);
begin
  inherited;

  if ComboGrupo.Text <> '' then
    NEGrupo.Text := ComboGrupo.KeyValue;
end;

procedure TfClien15.ComboProdutoCloseUp(Sender: TObject);
begin
  inherited;

  if ComboProduto.Text <> '' then
    MEProduto.Text := ComboProduto.KeyValue;

  EDProduto.Text := ComboProduto.Text;
  EDProduto.SetFocus;
  ComboProduto.Visible := False;

  fDtGeral.QryProDesc.Close;
end;

procedure TfClien15.ComboServicoCloseUp(Sender: TObject);
begin
  inherited;

  if ComboServico.Text <> '' then
    MEServico.Text := ComboServico.KeyValue;

  EDServico.Text := ComboServico.Text;
  EDServico.SetFocus;
  ComboServico.Visible := False;

  fDtGeral.QrySerDesc.Close;
end;

procedure TfClien15.ComboUsuarioCloseUp(Sender: TObject);
begin
  inherited;

  if ComboUsuario.Text <> '' then
    NEUsuario.Text := ComboUsuario.KeyValue;
end;

function TfClien15.Data: String;
begin
  Result := '    LTPDATAHORA BETWEEN :DATAINICIAL AND :DATAFINAL';
end;

procedure TfClien15.DTPDtInicialChange(Sender: TObject);
begin
  inherited;

  if Sender = DTPDtInicial then
    MEDtInicial.Text := DateToStr(DTPDtInicial.Date);
  if Sender = DTPDtFinal then
    MEDtFinal.Text := DateToStr(DTPDtFinal.Date);
end;

procedure TfClien15.EDClienteKeyPress(Sender: TObject; var Key: Char);
var
  N: Integer;
begin
  inherited;

  if (Key = #13) and (EDCliente.Text <> '') then
  begin
    fDtGeral.QryCliNome.Close;
    fDtGeral.QryCliNome.ParamByName('CLINOME').AsString := EDCliente.Text + '%';
    fDtGeral.QryCliNome.Open;

    { Conta os registros para abrir as linhas do combo.}
    for N := 1 to 15 do fDtGeral.QryCliNome.Next;

    fDtGeral.QryCliNome.First;

    ComboCliente.Visible := True;
    ComboCliente.SetFocus;
    ComboCliente.DropDown;
  end;
end;

procedure TfClien15.EDProdutoKeyPress(Sender: TObject; var Key: Char);
var
  N: Integer;
begin
  inherited;

  if (Key = #13) and (EDProduto.Text <> '') then
  begin
    fDtGeral.QryProDescR.Close;
    fDtGeral.QryProDescR.ParamByName('ProDescricao').AsString := EDProduto.Text + '%';
    fDtGeral.QryProDescR.Open;

    { Conta os registros para abrir as linhas do combo.}
    for N := 1 to 15 do fDtGeral.QryProDescR.Next;

    fDtGeral.QryProDescR.First;

    ComboProduto.Visible := True;
    ComboProduto.SetFocus;
    ComboProduto.DropDown;
  end;
end;

procedure TfClien15.EDServicoKeyPress(Sender: TObject; var Key: Char);
var
  N: Integer;
begin
  inherited;

  if (Key = #13) and (EDServico.Text <> '') then
  begin
    fDtGeral.QrySerDesc.Close;
    fDtGeral.QrySerDesc.ParamByName('SERDESCRICAO').AsString := EDServico.Text + '%';
    fDtGeral.QrySerDesc.Open;

    { Conta os registros para abrir as linhas do combo.}
    for N := 1 to 15 do fDtGeral.QrySerDesc.Next;

    fDtGeral.QrySerDesc.First;

    ComboServico.Visible := True;
    ComboServico.SetFocus;
    ComboServico.DropDown;
  end;
end;

procedure TfClien15.GetCurrentDate;
begin
  MEDtInicial.Text := FormatDateTime('dd/mm/yyyy', Date);
  MEDtFinal.Text := FormatDateTime('dd/mm/yyyy', Date);
  DTPDtInicial.Date := StrToDate(FormatDateTime('dd/mm/yyyy', Date));
  DTPDtFinal.Date := StrToDate(FormatDateTime('dd/mm/yyyy', Date));
end;

function TfClien15.Grupo: String;
begin
  if (NEGrupo.Text <> '') and (NEGrupo.Text <> '0') then
    Result :=
    'AND(' + #13#10 +
    '        ((LTPCAMPO1 CONTAINING ''GR1CODIGO'') AND ((LTPANTVALOR1 = ' + QuotedStr(NEGrupo.Text) + ') OR (LTPNOVOVALOR1 = ' + QuotedStr(NEGrupo.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO2 CONTAINING ''GR1CODIGO'') AND ((LTPANTVALOR2 = ' + QuotedStr(NEGrupo.Text) + ') OR (LTPNOVOVALOR2 = ' + QuotedStr(NEGrupo.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO3 CONTAINING ''GR1CODIGO'') AND ((LTPANTVALOR3 = ' + QuotedStr(NEGrupo.Text) + ') OR (LTPNOVOVALOR3 = ' + QuotedStr(NEGrupo.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO4 CONTAINING ''GR1CODIGO'') AND ((LTPANTVALOR4 = ' + QuotedStr(NEGrupo.Text) + ') OR (LTPNOVOVALOR4 = ' + QuotedStr(NEGrupo.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO5 CONTAINING ''GR1CODIGO'') AND ((LTPANTVALOR5 = ' + QuotedStr(NEGrupo.Text) + ') OR (LTPNOVOVALOR5 = ' + QuotedStr(NEGrupo.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO6 CONTAINING ''GR1CODIGO'') AND ((LTPANTVALOR6 = ' + QuotedStr(NEGrupo.Text) + ') OR (LTPNOVOVALOR6 = ' + QuotedStr(NEGrupo.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO7 CONTAINING ''GR1CODIGO'') AND ((LTPANTVALOR7 = ' + QuotedStr(NEGrupo.Text) + ') OR (LTPNOVOVALOR7 = ' + QuotedStr(NEGrupo.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO8 CONTAINING ''GR1CODIGO'') AND ((LTPANTVALOR8 = ' + QuotedStr(NEGrupo.Text) + ') OR (LTPNOVOVALOR8 = ' + QuotedStr(NEGrupo.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO9 CONTAINING ''GR1CODIGO'') AND ((LTPANTVALOR9 = ' + QuotedStr(NEGrupo.Text) + ') OR (LTPNOVOVALOR9 = ' + QuotedStr(NEGrupo.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO10 CONTAINING ''GR1CODIGO'') AND ((LTPANTVALOR10 = ' + QuotedStr(NEGrupo.Text) + ') OR (LTPNOVOVALOR10 = ' + QuotedStr(NEGrupo.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO11 CONTAINING ''GR1CODIGO'') AND ((LTPANTVALOR11 = ' + QuotedStr(NEGrupo.Text) + ') OR (LTPNOVOVALOR11 = ' + QuotedStr(NEGrupo.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO12 CONTAINING ''GR1CODIGO'') AND ((LTPANTVALOR12 = ' + QuotedStr(NEGrupo.Text) + ') OR (LTPNOVOVALOR12 = ' + QuotedStr(NEGrupo.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO13 CONTAINING ''GR1CODIGO'') AND ((LTPANTVALOR13 = ' + QuotedStr(NEGrupo.Text) + ') OR (LTPNOVOVALOR13 = ' + QuotedStr(NEGrupo.Text) + ')))' + #13#10 +
    ')'+ #13#10;
end;

procedure TfClien15.InitialValidations;
var
  I: Integer;
  LIsChecked: Boolean;
begin
  // A Data Inicial não pode ser vazia.
  if (MEDtInicial.Text = '  /  /    ') then
    raise Exception.Create('Informe uma Data Inicial para o relatório.');

  // A Data Final não pode ser vazia.
  if (MEDtFinal.Text = '  /  /    ') then
    raise Exception.Create('Informe uma Data Final para o relatório.');

  // A Data Inicial tem de ter um formato válido.
  try StrToDate(MEDtInicial.Text)
  except raise Exception.Create('O formato de Data Inicial informado é inválido.');
  end;

  // A Data Final tem de ter um formato válido.
  try StrToDate(MEDtFinal.Text)
  except raise Exception.Create('O formato de Data Final informado é inválido.');
  end;

  // A Data Inicial não pode ser maior que a Data Final.
  if StrToDate(MEDtInicial.Text) > StrToDate(MEDtFinal.Text) then
    raise Exception.Create('A Data Inicial não pode ser maior que a Data Final.');

  // Ao menos 1 item da TCheckListBox deve estar marcado.
  LIsChecked := False;
  for I := 0 to CLBItens.Items.Count - 1 do
    if CLBItens.Checked[I] then
    begin
      LIsChecked := True;
      Break;
    end;

  if not LIsChecked then
    raise Exception.Create('Selecione ao menos uma localização para o relatório.');
end;

function TfClien15.Local: String;
begin

end;

procedure TfClien15.MEClienteExit(Sender: TObject);
begin
  inherited;

  fDtGeral.QryCli.Close;
  fDtGeral.QryCli.ParamByName('CLICODIGO').AsString := MECliente.Text;
  fDtGeral.QryCli.Open;

  EDCliente.Text := fDtGeral.QryCli.FieldByName('CLINOMEFANT').AsString;

  if (Trim(Copy(MECliente.Text, 1, 1)) <> '') and (EDCliente.Text = '') then
  begin
    MessageDlg('Cliente não Cadastrado.', MtWarning, [MbOk], 0);
    MECliente.SetFocus;
  end;
end;

procedure TfClien15.MEDtInicialExit(Sender: TObject);
begin
  inherited;

  try
    if (Sender = MEDtInicial) and (MEDtInicial.Text <> '  /  /    ') then
      DTPDtInicial.Date := StrToDate(TMaskEdit(Sender).Text);
    if (Sender = MEDtFinal) and (MEDtFinal.Text <> '  /  /    ') then
      DTPDtFinal.Date := StrToDate(TMaskEdit(Sender).Text);
  except
    TMaskEdit(Sender).SetFocus;
    raise Exception.Create('O formato de data informado é inválido.');
  end;
end;

procedure TfClien15.MEProdutoExit(Sender: TObject);
begin
  inherited;

  fDtGeral.QryPro.Close;
  fDtGeral.QryPro.ParamByName('PROCODIGO').AsString := MEProduto.Text;
  fDtGeral.QryPro.Open;

  EDProduto.Text := fDtGeral.QryPro.FieldByName('PRODESCRICAO').AsString;

  if (Trim(Copy(MEProduto.Text, 1, 1)) <> '') and (EDProduto.Text = '') then
  begin
    MessageDlg('Produto não Cadastrado.', MtWarning, [MbOk], 0);
    MEProduto.SetFocus;
  end;
end;

procedure TfClien15.MEServicoExit(Sender: TObject);
begin
  inherited;

  fDtGeral.QrySer.Close;
  fDtGeral.QrySer.ParamByName('SERCODIGO').AsString := MEServico.Text;
  fDtGeral.QrySer.Open;

  EDServico.Text := fDtGeral.QrySer.FieldByName('SERDESCRICAO').AsString;

  if (Trim(Copy(MEServico.Text, 1, 1)) <> '') and (EDServico.Text = '') then
  begin
    MessageDlg('Serviço não Cadastrado.', MtWarning, [MbOk], 0);
    MEServico.SetFocus;
  end;
end;

procedure TfClien15.NEGrupoExit(Sender: TObject);
begin
  inherited;

  ComboGrupo.KeyValue := NEGrupo.Value;
  if (NEGrupo.Value <> 0) and (ComboGrupo.Text = '') then
  begin
    MessageDlg('Grupo não Cadastrado.', MtWarning, [MbOk], 0);
    NEGrupo.SetFocus;
  end;
end;

procedure TfClien15.NEUsuarioExit(Sender: TObject);
begin
  inherited;

  ComboUsuario.KeyValue := NEUsuario.Value;
  if (NEUsuario.Value <> 0) and (ComboUsuario.Text = '') then
  begin
    MessageDlg('Usuário não Cadastrado.', MtWarning, [MbOk], 0);
    NeUsuario.SetFocus;
  end;
end;

function TfClien15.Produto: String;
begin
  if (MEProduto.Text <> '') and (MEProduto.Text <> '0') then
    Result :=
    'AND(' + #13#10 +
    '        ((LTPCAMPO1 CONTAINING ''PROCODIGO'') AND ((LTPANTVALOR1 = ' + QuotedStr(MEProduto.Text) + ') OR (LTPNOVOVALOR1 = ' + QuotedStr(MEProduto.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO2 CONTAINING ''PROCODIGO'') AND ((LTPANTVALOR2 = ' + QuotedStr(MEProduto.Text) + ') OR (LTPNOVOVALOR2 = ' + QuotedStr(MEProduto.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO3 CONTAINING ''PROCODIGO'') AND ((LTPANTVALOR3 = ' + QuotedStr(MEProduto.Text) + ') OR (LTPNOVOVALOR3 = ' + QuotedStr(MEProduto.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO4 CONTAINING ''PROCODIGO'') AND ((LTPANTVALOR4 = ' + QuotedStr(MEProduto.Text) + ') OR (LTPNOVOVALOR4 = ' + QuotedStr(MEProduto.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO5 CONTAINING ''PROCODIGO'') AND ((LTPANTVALOR5 = ' + QuotedStr(MEProduto.Text) + ') OR (LTPNOVOVALOR5 = ' + QuotedStr(MEProduto.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO6 CONTAINING ''PROCODIGO'') AND ((LTPANTVALOR6 = ' + QuotedStr(MEProduto.Text) + ') OR (LTPNOVOVALOR6 = ' + QuotedStr(MEProduto.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO7 CONTAINING ''PROCODIGO'') AND ((LTPANTVALOR7 = ' + QuotedStr(MEProduto.Text) + ') OR (LTPNOVOVALOR7 = ' + QuotedStr(MEProduto.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO8 CONTAINING ''PROCODIGO'') AND ((LTPANTVALOR8 = ' + QuotedStr(MEProduto.Text) + ') OR (LTPNOVOVALOR8 = ' + QuotedStr(MEProduto.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO9 CONTAINING ''PROCODIGO'') AND ((LTPANTVALOR9 = ' + QuotedStr(MEProduto.Text) + ') OR (LTPNOVOVALOR9 = ' + QuotedStr(MEProduto.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO10 CONTAINING ''PROCODIGO'') AND ((LTPANTVALOR10 = ' + QuotedStr(MEProduto.Text) + ') OR (LTPNOVOVALOR10 = ' + QuotedStr(MEProduto.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO11 CONTAINING ''PROCODIGO'') AND ((LTPANTVALOR11 = ' + QuotedStr(MEProduto.Text) + ') OR (LTPNOVOVALOR11 = ' + QuotedStr(MEProduto.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO12 CONTAINING ''PROCODIGO'') AND ((LTPANTVALOR12 = ' + QuotedStr(MEProduto.Text) + ') OR (LTPNOVOVALOR12 = ' + QuotedStr(MEProduto.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO13 CONTAINING ''PROCODIGO'') AND ((LTPANTVALOR13 = ' + QuotedStr(MEProduto.Text) + ') OR (LTPNOVOVALOR13 = ' + QuotedStr(MEProduto.Text) + ')))' + #13#10 +
    ')'+ #13#10;
end;

function TfClien15.Servico: String;
begin
  if (MEServico.Text <> '') and (MEServico.Text <> '0') then
    Result :=
    'AND(' + #13#10 +
    '        ((LTPCAMPO1 CONTAINING ''SERCODIGO'') AND ((LTPANTVALOR1 = ' + QuotedStr(MEServico.Text) + ') OR (LTPNOVOVALOR1 = ' + QuotedStr(MEServico.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO2 CONTAINING ''SERCODIGO'') AND ((LTPANTVALOR2 = ' + QuotedStr(MEServico.Text) + ') OR (LTPNOVOVALOR2 = ' + QuotedStr(MEServico.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO3 CONTAINING ''SERCODIGO'') AND ((LTPANTVALOR3 = ' + QuotedStr(MEServico.Text) + ') OR (LTPNOVOVALOR3 = ' + QuotedStr(MEServico.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO4 CONTAINING ''SERCODIGO'') AND ((LTPANTVALOR4 = ' + QuotedStr(MEServico.Text) + ') OR (LTPNOVOVALOR4 = ' + QuotedStr(MEServico.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO5 CONTAINING ''SERCODIGO'') AND ((LTPANTVALOR5 = ' + QuotedStr(MEServico.Text) + ') OR (LTPNOVOVALOR5 = ' + QuotedStr(MEServico.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO6 CONTAINING ''SERCODIGO'') AND ((LTPANTVALOR6 = ' + QuotedStr(MEServico.Text) + ') OR (LTPNOVOVALOR6 = ' + QuotedStr(MEServico.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO7 CONTAINING ''SERCODIGO'') AND ((LTPANTVALOR7 = ' + QuotedStr(MEServico.Text) + ') OR (LTPNOVOVALOR7 = ' + QuotedStr(MEServico.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO8 CONTAINING ''SERCODIGO'') AND ((LTPANTVALOR8 = ' + QuotedStr(MEServico.Text) + ') OR (LTPNOVOVALOR8 = ' + QuotedStr(MEServico.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO9 CONTAINING ''SERCODIGO'') AND ((LTPANTVALOR9 = ' + QuotedStr(MEServico.Text) + ') OR (LTPNOVOVALOR9 = ' + QuotedStr(MEServico.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO10 CONTAINING ''SERCODIGO'') AND ((LTPANTVALOR10 = ' + QuotedStr(MEServico.Text) + ') OR (LTPNOVOVALOR10 = ' + QuotedStr(MEServico.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO11 CONTAINING ''SERCODIGO'') AND ((LTPANTVALOR11 = ' + QuotedStr(MEServico.Text) + ') OR (LTPNOVOVALOR11 = ' + QuotedStr(MEServico.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO12 CONTAINING ''SERCODIGO'') AND ((LTPANTVALOR12 = ' + QuotedStr(MEServico.Text) + ') OR (LTPNOVOVALOR12 = ' + QuotedStr(MEServico.Text) + '))) OR' + #13#10 +
    '        ((LTPCAMPO13 CONTAINING ''SERCODIGO'') AND ((LTPANTVALOR13 = ' + QuotedStr(MEServico.Text) + ') OR (LTPNOVOVALOR13 = ' + QuotedStr(MEServico.Text) + ')))' + #13#10 +
    ')'+ #13#10;
end;

procedure TfClien15.ToogleQueryActive;
begin
  LProduto := fDtGeral.QryPro.Active; if not LProduto then fDtGeral.QryPro.Open;
  LServico := fDtGeral.QrySer.Active; if not LServico then fDtGeral.QrySer.Open;
  LUsuario := fDtGeral.QryUsu.Active; if not LUsuario then fDtGeral.QryUsu.Open;
  LCliente := fDtGeral.QryCli.Active; if not LCliente then fDtGeral.QryCli.Open;
  LGrupo := fDtMod.Grupo1.Active; if not LGrupo then fDtMod.Grupo1.Open;
end;

procedure TfClien15.ToogleQueryInactive;
begin
  if LProduto then fDtGeral.QryPro.Close;
  if LServico then fDtGeral.QrySer.Close;
  if LUsuario then fDtGeral.QryUsu.Close;
  if LCliente then fDtGeral.QryCli.Close;
  if LGrupo then fDtMod.Grupo1.Close;
end;

function TfClien15.Usuario: String;
begin
  if (NEUsuario.Text <> '0') and (NEUsuario.Text <> '') then
    Result := '    AND LTPUSUCODIGO = :USUCODIGO';
end;

end.
