unit Relatorio7;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, Grids, DBGrids, StdCtrls, Mask, Buttons;

type
  TfrmRelatorio7 = class(TForm)
    Label1: TLabel;
    btnExportar: TSpeedButton;
    btnImprimir: TSpeedButton;
    lblTotal: TLabel;
    lblRTotal: TLabel;
    GroupBox1: TGroupBox;
    btnGerar: TSpeedButton;
    GroupBox2: TGroupBox;
    Label2: TLabel;
    dbgDados: TDBGrid;
    dsDados: TDataSource;
    qryDados: TADOQuery;
    SaveDialog1: TSaveDialog;
    qryExec: TADOQuery;
    Med_DataFinal: TMaskEdit;
    Med_DataInicio: TMaskEdit;
    lblDia: TLabel;
    edtDia: TEdit;
    Med_DataBase: TMaskEdit;
    chkDtConciliacao: TCheckBox;
    EdtPlano: TEdit;
    ChkPlano: TCheckBox;
    procedure Med_DataInicioExit(Sender: TObject);
    procedure Med_DataFinalExit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure btnGerarClick(Sender: TObject);
    procedure btnExportarClick(Sender: TObject);
    procedure edtDiaKeyPress(Sender: TObject; var Key: Char);
    procedure chkDtConciliacaoClick(Sender: TObject);
    procedure Med_DataBaseExit(Sender: TObject);
    procedure qryDadosAfterOpen(DataSet: TDataSet);
    procedure ChkPlanoClick(Sender: TObject);
  private
    { Private declarations }
    procedure DimensionarGrid(dbg: TDbGrid; var formulario);
  public
    { Public declarations }
  end;

var
  frmRelatorio7: TfrmRelatorio7;
  parametro: String;
implementation

{$R *.dfm}
uses Principal, DateUtils;
procedure TfrmRelatorio7.Med_DataInicioExit(Sender: TObject);
var d_Data : TDateTime;
begin
  if Med_DataInicio.Text <> '  /  /    ' then
   begin
    try
      d_Data := StrToDateTime(Med_DataInicio.Text);
      Med_DataInicio.Text := FormatDateTime('dd/mm/yyyy',d_Data);
    except
      ShowMessage('Data inv�lida');
      Med_DataInicio.Text := '  /  /    ';
      Med_DataInicio.SetFocus;
    end;
  end;
end;

procedure TfrmRelatorio7.Med_DataFinalExit(Sender: TObject);
var d_Data : TDateTime;
begin
  if Med_DataFinal.Text <> '  /  /    ' then
   begin
    try
      d_Data := StrToDateTime(Med_DataFinal.Text);
      Med_DataFinal.Text := FormatDateTime('dd/mm/yyyy',d_Data);
    except
      ShowMessage('Data inv�lida');
      Med_DataFinal.Text := '  /  /    ';
      Med_DataFinal.SetFocus;
    end;
  end;
end;

procedure TfrmRelatorio7.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
qryDados.Close;
end;

procedure TfrmRelatorio7.FormShow(Sender: TObject);
begin
  if frmPrincipal.lblExporta.Caption = 'T' then //Habilita bot�o Exportar
  begin
  btnExportar.Enabled := True;
  end else
  begin
  btnExportar.Enabled := False;
  end;

  with qryExec do
  begin
  Close;
  SQL.Clear;
  SQL.Add('ALTER SESSION SET NLS_DATE_FORMAT = '+#39+'DD/MM/YYYY HH24:MI:SS'+#39);
  ExecSQL;
  end;

  if (frmPrincipal.lblCodRelatorio.Caption = '80') or (frmPrincipal.lblCodRelatorio.Caption = '81') then
  begin
    lblDia.Visible := True;
    edtDia.Visible := True;
    lblDia.Caption := 'Dias em aberto';
    GroupBox2.Caption := 'Data Anterior e Data Atual';
    Label2.Caption := 'e';
    chkDtConciliacao.Visible := False;
    Med_DataBase.Visible := False;
    Med_DataFinal.Text := '  /  /    ';
    Med_DataInicio.Text := '  /  /    ';
    edtDia.Text := '';
    Abort;
  end;

  if (frmPrincipal.lblCodRelatorio.Caption = '91') or (frmPrincipal.lblCodRelatorio.Caption = '93') then
  begin
    lblDia.Visible := False;
    edtDia.Visible := False;
    GroupBox2.Caption := 'Per�odo de datas';
    Label2.Caption := 'at�';
    chkDtConciliacao.Visible := False;
    Med_DataBase.Visible := False;
    Med_DataFinal.Text := '  /  /    ';
    Med_DataInicio.Text := '  /  /    ';
    EdtPlano.Visible := True;
    ChkPlano.Visible := True;
    ChkPlano.Checked := False;
    EdtPlano.Clear;
    EdtPlano.Color := cl3DLight;
    EdtPlano.Enabled := False;
    Abort;
  end;

  if (frmPrincipal.lblCodRelatorio.Caption = '48') or (frmPrincipal.lblCodRelatorio.Caption = '70') then
  begin
    lblDia.Visible := False;
    edtDia.Visible := False;
    GroupBox2.Caption := 'Per�odo de inclus�o';
    Label2.Caption := 'at�';
    chkDtConciliacao.Visible := True;
    chkDtConciliacao.Checked := False;
    Med_DataBase.Visible := True;
    Med_DataFinal.Text := '  /  /    ';
    Med_DataInicio.Text := '  /  /    ';
    Med_DataBase.Text := '  /  /    ';
    Abort;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '85' then
  begin
    lblDia.Visible := True;
    edtDia.Visible := True;
    lblDia.Caption := 'Diferen�a de dias';
    GroupBox2.Caption := 'Per�odo de inclus�o';
    Label2.Caption := 'at�';
    chkDtConciliacao.Visible := False;
    chkDtConciliacao.Checked := False;
    Med_DataBase.Visible := False;
    Med_DataFinal.Text := '  /  /    ';
    Med_DataInicio.Text := '  /  /    ';
    Med_DataBase.Text := '  /  /    ';
    Abort;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '128' then
  begin
    lblDia.Visible := True;
    edtDia.Visible := True;
    lblDia.Caption := 'Forma de pagamento';
    GroupBox2.Caption := 'Per�odo de cadastro';
    Label2.Caption := 'at�';
    chkDtConciliacao.Visible := False;
    Med_DataBase.Visible := False;
    Med_DataFinal.Text := '  /  /    ';
    Med_DataInicio.Text := '  /  /    ';
    edtDia.Text := '';
    Abort;
  end;

  lblDia.Visible := False;
  edtDia.Visible := False;
  GroupBox2.Caption := 'Per�odo de datas';
  Label2.Caption := 'at�';
  chkDtConciliacao.Visible := False;
  Med_DataBase.Visible := False;
  Med_DataFinal.Text := '  /  /    ';
  Med_DataInicio.Text := '  /  /    ';
  EdtPlano.Visible := False;
  ChkPlano.Visible := False;
  EdtPlano.Clear;
  ChkPlano.Checked := False;
  EdtPlano.Color := cl3DLight;
  EdtPlano.Enabled := False;

  if frmPrincipal.lblCodRelatorio.Caption = '106' then
  begin
    GroupBox2.Caption := 'In�cio de Vig�ncia';
  end;
end;

procedure TfrmRelatorio7.btnGerarClick(Sender: TObject);
begin
  Med_DataInicio.OnExit(Sender);
  Med_DataFinal.OnExit(Sender);

  if Med_DataBase.text <> '  /  /    ' then
  Med_DataBase.OnExit(Sender);

  parametro := '';

  if (frmPrincipal.lblCodRelatorio.Caption = '80') or (frmPrincipal.lblCodRelatorio.Caption = '81') then
  begin
    if (Med_DataInicio.Text <> '  /  /    ') and (Med_DataFinal.Text <> '  /  /    ') and (edtDia.Text <> '') then
    begin
      parametro := 'DtInicio: '+Med_DataInicio.Text + ' DtFinal: ' + Med_DataFinal.Text + ' Dias: '+edtDia.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&DATA1',Med_DataInicio.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&DATA2',Med_DataFinal.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&DIAS_ATRASO',edtDia.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;

      lblRTotal.Caption := IntToStr(qryDados.RecordCount);

      with qryExec do // INSERIR LOG DE VISUALIZA��O
      begin
      Close;
      SQL.Clear;
      SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+', '+#39+parametro+#39+', SYSDATE)');
      ExecSQL;
      end;
      Abort;
    end else
    begin
      ShowMessage('Favor informar todos os par�metros.');
    end;
    Abort;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '48' then
  begin
    if (Med_DataInicio.Text <> '  /  /    ') and (Med_DataFinal.Text <> '  /  /    ') then
    begin
      parametro := 'DtInicio: '+Med_DataInicio.Text + ' DtFinal: ' + Med_DataFinal.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&DATA1',Med_DataInicio.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&DATA2',Med_DataFinal.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco

      if chkDtConciliacao.Checked = True then
      begin
        SQL.Add('AND   B.DT_CONCIL = '+#39+Med_DataBase.Text+#39);
        parametro := parametro + 'Concilia��o: '+ Med_DataBase.Text;
      end;

      SQL.Add('GROUP BY S.DT_INCL, S.COD_SEG, S.NOME, DECODE(S.TIPO,'+#39+'E'+#39+','+#39+'TITULAR'+#39+','+#39+'D'+#39+','+#39+'DEPENDENTE'+#39+'), S.MAT, S.COD_FAM, E.DT_CAD, S.COD_SET || '+#39+' - '+#39+' || E.DESCRICAO, S.COD_SETSET || '+#39+' - '+#39+' || U.NOME, C.COD_VEND || '+#39+' - '+#39+' || C.NOME, CAR.DT_INICIO, E.COD_SET) P,');
      SQL.Add('(SELECT V.COD_SEG, C.COD_VEND || '+#39+' - '+#39+' || C.NOME SUBCORRETOR');
      SQL.Add('FROM IM_VENDA V, IM_SGVEA L, IM_VEND C');
      SQL.Add('WHERE V.COD_VENDA = L.COD_VENDA AND L.COD_VEND  = C.COD_VEND AND C.TIPO      = '+#39+'S'+#39+') P2,');
      SQL.Add('(SELECT V.COD_SEG, C.COD_VEND || '+#39+' - '+#39+' || C.NOME DIRETOR');
      SQL.Add('FROM IM_VENDA V, IM_SGVEA L, IM_VEND C');
      SQL.Add('WHERE V.COD_VENDA = L.COD_VENDA AND L.COD_VEND  = C.COD_VEND AND C.TIPO      = '+#39+'D'+#39+') P3,');
      SQL.Add('(SELECT V.COD_SEG, C.COD_VEND || '+#39+' - '+#39+' || C.NOME PRESIDENTE');
      SQL.Add('FROM IM_VENDA V, IM_SGVEA L, IM_VEND C');
      SQL.Add('WHERE V.COD_VENDA = L.COD_VENDA AND L.COD_VEND  = C.COD_VEND AND C.TIPO      = '+#39+'P'+#39+') P4,');
      SQL.Add('(SELECT E.COD_SET, MIN(CAR.DT_INICIO) DT_INICIO');
      SQL.Add('FROM IM_SGCAR CAR, IM_SEG S, IM_SETOR E');
      SQL.Add('WHERE S.COD_SEG = CAR.COD_SEG AND   S.COD_SET = E.COD_SET AND   CAR.COD_CAR = 12 GROUP BY E.COD_SET) EMP');
      SQL.Add('WHERE P.COD_SEG = P2.COD_SEG (+) AND P.COD_SEG = P3.COD_SEG (+) AND P.COD_SEG = P4.COD_SEG (+) AND P.DT_INICIO = EMP.DT_INICIO AND P.COD_SET   = EMP.COD_SET');
      SQL.Add('ORDER BY 7 DESC, 8, 2');
      Open;
      end;

      lblRTotal.Caption := IntToStr(qryDados.RecordCount);

      with qryExec do // INSERIR LOG DE VISUALIZA��O
      begin
      Close;
      SQL.Clear;
      SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+', '+#39+parametro+#39+', SYSDATE)');
      ExecSQL;
      end;
      Abort;
    end else
    begin
      ShowMessage('Favor informar todos os par�metros.');
    end;
    Abort;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '85' then
  begin
    if (Med_DataInicio.Text <> '  /  /    ') and (Med_DataFinal.Text <> '  /  /    ') and (edtDia.Text <> '') then
    begin
      parametro := 'DtInicio: '+Med_DataInicio.Text + ' DtFinal: ' + Med_DataFinal.Text + ' Dias: '+edtDia.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&DATA1',Med_DataInicio.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&DATA2',Med_DataFinal.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&DIA',edtDia.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;

      lblRTotal.Caption := IntToStr(qryDados.RecordCount);

      with qryExec do // INSERIR LOG DE VISUALIZA��O
      begin
      Close;
      SQL.Clear;
      SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+', '+#39+parametro+#39+', SYSDATE)');
      ExecSQL;
      end;
      Abort;
    end else
    begin
      ShowMessage('Favor informar todos os par�metros.');
    end;
    Abort;
  end;

  if (frmPrincipal.lblCodRelatorio.Caption = '91') or (frmPrincipal.lblCodRelatorio.Caption = '93') then
  begin
    if (Med_DataInicio.Text <> '  /  /    ') and (Med_DataFinal.Text <> '  /  /    ') then
    begin
      parametro := 'DtInicio: '+Med_DataInicio.Text + ' DtFinal: ' + Med_DataFinal.Text + ' planos: '+EdtPlano.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&DATA1',Med_DataInicio.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&DATA2',Med_DataFinal.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco

      if (ChkPlano.Checked = True) and (EdtPlano.Text <> '') then
      SQL.Text := StringReplace(SQL.Text,'&PLANO',EdtPlano.Text,[rfReplaceAll]) //Alterar parametro dentro do SQL puxado do banco
      else
      SQL.Text := StringReplace(SQL.Text,'&PLANO','999',[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco

      Open;
      end;

      lblRTotal.Caption := IntToStr(qryDados.RecordCount);

      with qryExec do // INSERIR LOG DE VISUALIZA��O
      begin
      Close;
      SQL.Clear;
      SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+', '+#39+parametro+#39+', SYSDATE)');
      ExecSQL;
      end;
      Abort;
    end else
    begin
      ShowMessage('Favor informar todos os par�metros.');
    end;
    Abort;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '128' then
  begin
    if (Med_DataInicio.Text <> '  /  /    ') and (Med_DataFinal.Text <> '  /  /    ') and (edtDia.Text <> '') then
    begin
      parametro := 'DtInicio: '+Med_DataInicio.Text + ' DtFinal: ' + Med_DataFinal.Text + ' Forma Pag: '+edtDia.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&DATA1',Med_DataInicio.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&DATA2',Med_DataFinal.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      SQL.Text := StringReplace(SQL.Text,'&FORMA',edtDia.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;

      lblRTotal.Caption := IntToStr(qryDados.RecordCount);

      with qryExec do // INSERIR LOG DE VISUALIZA��O
      begin
      Close;
      SQL.Clear;
      SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+', '+#39+parametro+#39+', SYSDATE)');
      ExecSQL;
      end;
      Abort;
    end else
    begin
      ShowMessage('Favor informar todos os par�metros.');
    end;
    Abort;
  end;

  if (Med_DataInicio.Text <> '  /  /    ') and (Med_DataFinal.Text <> '  /  /    ') then
  begin
    parametro := 'DtInicio: '+Med_DataInicio.Text + ' DtFinal: ' + Med_DataFinal.Text;
    with qryDados do
    begin
    Close;
    SQL.Clear;
    SQL.Add(frmPrincipal.memSQL.Text);
    SQL.Text := StringReplace(SQL.Text,'&DATA1',Med_DataInicio.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
    SQL.Text := StringReplace(SQL.Text,'&DATA2',Med_DataFinal.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
    Open;
    end;

    lblRTotal.Caption := IntToStr(qryDados.RecordCount);

    with qryExec do // INSERIR LOG DE VISUALIZA��O
    begin
    Close;
    SQL.Clear;
    SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+', '+#39+parametro+#39+', SYSDATE)');
    ExecSQL;
    end;
  end else
  begin
    ShowMessage('Favor informar o per�odo.');
    Abort;
  end;

end;

procedure TfrmRelatorio7.btnExportarClick(Sender: TObject);
var
F: TextFile;
P1, P2, P3, P4, P5, P6, P7, P8, P9, P10, P11, P12, P13, P14, P15, P16, P17, P18, P19, P20, P21, P22, P23, P24,
P25, P26, P27, P28, P29, P30, P31, P32, P33, P34, P35, P36, P37, P38, cLinha  : String;
begin
 if qryDados.RecordCount > 0 then
 begin
  if frmPrincipal.lblCodRelatorio.Caption = '14' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'NOME' +';'+ 'TIPO' +';'+ 'CPF' +';'+ 'DESCRICAO' +';'+ 'COD_ADTV' +';'+ 'ODONTO' +';'+ 'DT_INICIO' +';'+ 'DT_CANC' +';'+ 'COD_PLANO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P2  := qryDados.FIELDBYNAME('NOME').AsString;
          P3  := qryDados.FIELDBYNAME('TIPO').AsString;
          P4  := qryDados.FIELDBYNAME('CPF').AsString;
          P5  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P6  := qryDados.FIELDBYNAME('COD_ADTV').AsString;
          P7  := qryDados.FIELDBYNAME('ODONTO').AsString;
          P8  := qryDados.FIELDBYNAME('DT_INICIO').AsString;
          P9  := qryDados.FIELDBYNAME('DT_CANC').AsString;
          P10 := qryDados.FIELDBYNAME('COD_PLANO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';          

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '20' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'COD_FORMA' +';'+ 'NOME' +';'+ 'SEXO' +';'+ 'IDADE' +';'+ 'COD_PLANO' +';'+ 'ADESAO' +';'+ 'VIGENCIA'+';'+ 'TIPO' +';'+ 'CANCELADO' +';'+ 'VENDEDOR' +';'+ 'CPF' +';'+ 'MAT' +';'+ 'TIPO_VEND');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P2  := qryDados.FIELDBYNAME('COD_FORMA').AsString;
          P3  := qryDados.FIELDBYNAME('NOME').AsString;
          P4  := qryDados.FIELDBYNAME('SEXO').AsString;
          P5  := qryDados.FIELDBYNAME('IDADE').AsString;
          P6  := qryDados.FIELDBYNAME('COD_PLANO').AsString;
          P7  := qryDados.FIELDBYNAME('ADESAO').AsString;
          P8  := qryDados.FIELDBYNAME('VIGENCIA').AsString;
          P9  := qryDados.FIELDBYNAME('TIPO').AsString;
          P10 := qryDados.FIELDBYNAME('CANCELADO').AsString;
          P11 := qryDados.FIELDBYNAME('VENDEDOR').AsString;
          P12 := qryDados.FIELDBYNAME('CPF').AsString;
          P13 := qryDados.FIELDBYNAME('MAT').AsString;
          P14 := qryDados.FIELDBYNAME('TIPO_VEND').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';
          cLinha := cLinha + P12 +';';
          cLinha := cLinha + P13 +';';
          cLinha := cLinha + P14 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '22' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'PLANO' +';'+ 'COD_SEG' +';'+ 'NOME' +';'+ 'LOGIN' +';'+ 'DATA_ADESAO' +';'+ 'DT_INICIO' +';'+ 'CONTRATO' +';'+ 'DESCRICAO' +';'+ 'TIPO' +';'+ 'FORMA_PAG' +';'+ 'VENDEDOR' +';'+ 'SUPERVISOR' +';'+ 'MENSALIDADE'  +';'+ 'DATA_ENTRADA');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('PLANO').AsString;
          P2  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P3  := qryDados.FIELDBYNAME('NOME').AsString;
          P4  := qryDados.FIELDBYNAME('LOGIN').AsString;
          P5  := qryDados.FIELDBYNAME('DATA_ADESAO').AsString;
          P6  := qryDados.FIELDBYNAME('DT_INICIO').AsString;
          P7  := qryDados.FIELDBYNAME('CONTRATO').AsString;
          P8  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P9  := qryDados.FIELDBYNAME('TIPO').AsString;
          P10 := qryDados.FIELDBYNAME('FORMA_PAG').AsString;
          P11 := qryDados.FIELDBYNAME('VENDEDOR').AsString;
          P12 := qryDados.FIELDBYNAME('SUPERVISOR').AsString;
          P13 := qryDados.FIELDBYNAME('MENSALIDADE').AsString;
          P14 := qryDados.FIELDBYNAME('DATA_ENTRADA').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';
          cLinha := cLinha + P12 +';';
          cLinha := cLinha + P13 +';';
          cLinha := cLinha + P14 +';';          

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '24' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'BASE' +';'+ 'COD_SET' +';'+ 'DESCRICAO' +';'+ 'COD_SEG' +';'+ 'NOME' +';'+ 'CPF' +';'+ 'DT_CANC' +';'+ 'SITUACAO' +';'+ 'TIPO' +';'+ 'COD_PRODUTO' +';'+ 'NOME_PRODUTO' +';'+ 'DT_INCLUSAO' +';'+ 'DATA_ENTRADA' +';'+ 'COD_FORMA' +';'+ 'FORMA_PAG' +';'+ 'COD_VEND' +';'+ 'VENDEDOR' +';'+ 'MAT' +';'+ 'VALOR');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1   := qryDados.FIELDBYNAME('BASE').AsString;
          P2   := qryDados.FIELDBYNAME('COD_SET').AsString;
          P3   := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P4   := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P5   := qryDados.FIELDBYNAME('NOME').AsString;
          P6   := qryDados.FIELDBYNAME('CPF').AsString;
          P7   := qryDados.FIELDBYNAME('DT_CANC').AsString;
          P8   := qryDados.FIELDBYNAME('SITUACAO').AsString;
          P9   := qryDados.FIELDBYNAME('TIPO').AsString;
          P10  := qryDados.FIELDBYNAME('COD_PRODUTO').AsString;
          P11  := qryDados.FIELDBYNAME('NOME_PRODUTO').AsString;
          P12  := qryDados.FIELDBYNAME('DT_INCLUSAO').AsString;
          P13  := qryDados.FIELDBYNAME('DATA_ENTRADA').AsString;
          P14  := qryDados.FIELDBYNAME('COD_FORMA').AsString;
          P15  := qryDados.FIELDBYNAME('FORMA_PAG').AsString;
          P16  := qryDados.FIELDBYNAME('COD_VEND').AsString;
          P17  := qryDados.FIELDBYNAME('VENDEDOR').AsString;
          P18  := qryDados.FIELDBYNAME('MAT').AsString;
          P19  := qryDados.FIELDBYNAME('VALOR').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';
          cLinha := cLinha + P12 +';';
          cLinha := cLinha + P13 +';';
          cLinha := cLinha + P14 +';';
          cLinha := cLinha + P15 +';';
          cLinha := cLinha + P16 +';';
          cLinha := cLinha + P17 +';';
          cLinha := cLinha + P18 +';';
          cLinha := cLinha + P19 +';';          

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '26' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'NOME' +';'+ 'RES_RUA' +';'+ 'RES_NUMERO' +';'+ 'RES_COMPLEMENTO' +';'+ 'BAIRRO' +';'+ 'CIDADE' +';'+ 'ESTADO' +';'+ 'CEP');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P2  := qryDados.FIELDBYNAME('NOME').AsString;
          P3  := qryDados.FIELDBYNAME('RES_RUA').AsString;
          P4  := qryDados.FIELDBYNAME('RES_NUMERO').AsString;
          P5  := qryDados.FIELDBYNAME('RES_COMPLEMENTO').AsString;
          P6  := qryDados.FIELDBYNAME('BAIRRO').AsString;
          P7  := qryDados.FIELDBYNAME('CIDADE').AsString;
          P8  := qryDados.FIELDBYNAME('ESTADO').AsString;
          P9  := qryDados.FIELDBYNAME('RES_CEP').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '27' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'GUIA' +';'+ 'SITUACAO' +';'+ 'EMISSAO' +';'+ 'LOGIN' +';'+ 'TIPO' +';'+ 'COD_SEG' +';'+ 'VALOR' +';'+ 'GLOSA' +';'+ 'PAGO' +';'+ 'DATA_ALTERACAO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('GUIA').AsString;
          P2  := qryDados.FIELDBYNAME('SITUACAO').AsString;
          P3  := qryDados.FIELDBYNAME('DT_EMISSAO').AsString;
          P4  := qryDados.FIELDBYNAME('LOGIN').AsString;
          P5  := qryDados.FIELDBYNAME('TIPO').AsString;
          P6  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P7  := qryDados.FIELDBYNAME('VALOR').AsString;
          P8  := qryDados.FIELDBYNAME('GLOSA').AsString;
          P9  := qryDados.FIELDBYNAME('PAGO').AsString;
          P10 := qryDados.FIELDBYNAME('DATA_ALTERACAO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '30' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'LOGIN' +';'+ 'DT_LIB_SENHA' +';'+ 'SENHA_LIB' +';'+ 'COD_AMB' +';'+ 'DESCRICAO' +';'+ 'SITUACAO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('LOGIN').AsString;
          P2  := qryDados.FIELDBYNAME('DT_LIB_SENHA').AsString;
          P3  := qryDados.FIELDBYNAME('SENHA_LIB').AsString;
          P4  := qryDados.FIELDBYNAME('COD_AMB').AsString;
          P5  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P6  := qryDados.FIELDBYNAME('SITUACAO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '31' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'GUIA' +';'+ 'CODIGO' +';'+ 'COD_AMB' +';'+ 'COD_SEG' +';'+ 'NOME' +';'+ 'GERACAO' +';'+ 'COD_MED_SOL' +';'+ 'MEDICO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('GUIA').AsString;
          P2  := qryDados.FIELDBYNAME('CODIGO').AsString;
          P3  := qryDados.FIELDBYNAME('COD_AMB').AsString;
          P4  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P5  := qryDados.FIELDBYNAME('NOME').AsString;
          P6  := qryDados.FIELDBYNAME('GERACAO').AsString;
          P7  := qryDados.FIELDBYNAME('COD_MED_SOL').AsString;
          P8  := qryDados.FIELDBYNAME('MEDICO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '41' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SET' +';'+ 'DESCRICAO' +';'+ 'COD_SEG' +';'+ 'NOME' +';'+ 'TIPO' +';'+ 'TITULAR' +';'+ 'COD_VEND' +';'+ 'NOME_VEND' +';'+ 'COD_SUP' +';'+ 'NOME_SUP' +';'+ 'COD_DIR' +';'+ 'NOME_DIR');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SET').AsString;
          P2  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P3  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P4  := qryDados.FIELDBYNAME('NOME').AsString;
          P5  := qryDados.FIELDBYNAME('TIPO').AsString;
          P6  := qryDados.FIELDBYNAME('TITULAR').AsString;
          P7  := qryDados.FIELDBYNAME('COD_VEND').AsString;
          P8  := qryDados.FIELDBYNAME('NOME_VEND').AsString;
          P9  := qryDados.FIELDBYNAME('COD_SUP').AsString;
          P10 := qryDados.FIELDBYNAME('NOME_SUP').AsString;
          P11 := qryDados.FIELDBYNAME('COD_DIR').AsString;
          P12 := qryDados.FIELDBYNAME('NOME_DIR').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';
          cLinha := cLinha + P12 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '43' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'ADESAO_EMPRESA' +';'+ 'ADESAO_BENEFICIARIO' +';'+ 'VIGENCIA_BENEFICIARIO' +';'+ 'SITUACAO' +';'+ 'MAT' +';'+ 'TIPO' +';'+ 'COD_SEG' +';'+ 'NOME' +';'+ 'COD_SET' +';'+ 'DESCRICAO' +';'+ 'COD_FORMA' +';'+ 'FORMA' +';'+ 'COD_VENDEDOR' +';'+ 'VENDEDOR' +';'+ 'COD_SUPERVISOR' +';'+ 'SUPERVISOR' +';'+ 'COD_DIRETOR' +';'+ 'DIRETOR' +';'+ 'P_MENSALIDADE' +';'+ 'P_MENSALIDADE2' +';'+ 'P_MENSALIDADE3' +';'+ 'TIPO_CONTRATO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('ADESAO_EMPRESA').AsString;
          P2  := qryDados.FIELDBYNAME('ADESAO_BENEFICIARIO').AsString;
          P3  := qryDados.FIELDBYNAME('VIGENCIA_BENEFICIARIO').AsString;
          P4  := qryDados.FIELDBYNAME('SITUACAO').AsString;
          P5  := qryDados.FIELDBYNAME('MAT').AsString;
          P6  := qryDados.FIELDBYNAME('TIPO').AsString;
          P7  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P8  := qryDados.FIELDBYNAME('NOME').AsString;
          P9  := qryDados.FIELDBYNAME('COD_SET').AsString;
          P10 := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P11 := qryDados.FIELDBYNAME('COD_FORMA').AsString;
          P12 := qryDados.FIELDBYNAME('FORMA').AsString;
          P13 := qryDados.FIELDBYNAME('COD_VENDEDOR').AsString;
          P14 := qryDados.FIELDBYNAME('VENDEDOR').AsString;
          P15 := qryDados.FIELDBYNAME('COD_SUPERVISOR').AsString;
          P16 := qryDados.FIELDBYNAME('SUPERVISOR').AsString;
          P17 := qryDados.FIELDBYNAME('COD_DIRETOR').AsString;
          P18 := qryDados.FIELDBYNAME('DIRETOR').AsString;
          P19 := qryDados.FIELDBYNAME('P_MENSALIDADE').AsString;
          P20 := qryDados.FIELDBYNAME('P_MENSALIDADE2').AsString;
          P21 := qryDados.FIELDBYNAME('P_MENSALIDADE3').AsString;
          P22 := qryDados.FIELDBYNAME('TIPO_CONTRATO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';
          cLinha := cLinha + P12 +';';
          cLinha := cLinha + P13 +';';
          cLinha := cLinha + P14 +';';
          cLinha := cLinha + P15 +';';
          cLinha := cLinha + P16 +';';
          cLinha := cLinha + P17 +';';
          cLinha := cLinha + P18 +';';
          cLinha := cLinha + P19 +';';
          cLinha := cLinha + P20 +';';
          cLinha := cLinha + P21 +';';
          cLinha := cLinha + P22 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '55' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'NOME' +';'+ 'TIPO' +';'+ 'CPF' +';'+ 'DEESCRICAO' +';'+ 'COD_ADTV' +';'+ 'ODONTO' +';'+ 'DT_INICIO' +';'+ 'DT_CANC' +';'+ 'DATA_ENTRADA' +';'+ 'COD_FORMA' +';'+ 'FORMA_PAG' +';'+ 'COD_VEND' +';'+ 'VENDEDOR');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P2  := qryDados.FIELDBYNAME('NOME').AsString;
          P3  := qryDados.FIELDBYNAME('TIPO').AsString;
          P4  := qryDados.FIELDBYNAME('CPF').AsString;
          P5  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P6  := qryDados.FIELDBYNAME('COD_ADTV').AsString;
          P7  := qryDados.FIELDBYNAME('ODONTO').AsString;
          P8  := qryDados.FIELDBYNAME('DT_INICIO').AsString;
          P9  := qryDados.FIELDBYNAME('DT_CANC').AsString;
          P10 := qryDados.FIELDBYNAME('DATA_ENTRADA').AsString;
          P11 := qryDados.FIELDBYNAME('COD_FORMA').AsString;
          P12 := qryDados.FIELDBYNAME('FORMA_PAG').AsString;
          P13 := qryDados.FIELDBYNAME('COD_VEND').AsString;
          P14 := qryDados.FIELDBYNAME('VENDEDOR').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';
          cLinha := cLinha + P12 +';';
          cLinha := cLinha + P13 +';';
          cLinha := cLinha + P14 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '63' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'NOME' +';'+ 'TIPO' +';'+ 'PLANO' +';'+ 'VALOR' +';'+ 'COD_BOL' +';'+ 'DT_EMISSAO' +';'+ 'DIA_VENC' +';'+ 'DT_CONCIL' +';'+ 'EMPRESA' +';'+ 'UNIDADE');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P2  := qryDados.FIELDBYNAME('NOME').AsString;
          P3  := qryDados.FIELDBYNAME('TIPO').AsString;
          P4  := qryDados.FIELDBYNAME('PLANO').AsString;
          P5  := qryDados.FIELDBYNAME('VALOR').AsString;
          P6  := qryDados.FIELDBYNAME('COD_BOL').AsString;
          P7  := qryDados.FIELDBYNAME('DT_EMISSAO').AsString;
          P8  := qryDados.FIELDBYNAME('DIA_VENC').AsString;
          P9  := qryDados.FIELDBYNAME('DT_CONCIL').AsString;
          P10 := qryDados.FIELDBYNAME('EMPRESA').AsString;
          P11 := qryDados.FIELDBYNAME('UNIDADE').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '64' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'USUARIO' +';'+ 'DIA' +';'+ 'COD_ENT' +';'+ 'FANTASIA' +';'+ 'QTD' +';'+ 'V_TOT' +';'+ 'V_PAG');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('USU').AsString;
          P2  := qryDados.FIELDBYNAME('DIA').AsString;
          P3  := qryDados.FIELDBYNAME('COD_ENT').AsString;
          P4  := qryDados.FIELDBYNAME('FANTASIA').AsString;
          P5  := qryDados.FIELDBYNAME('QTD').AsString;
          P6  := qryDados.FIELDBYNAME('V_TOT').AsString;
          P7  := qryDados.FIELDBYNAME('V_PAG').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if (frmPrincipal.lblCodRelatorio.Caption = '72') or (frmPrincipal.lblCodRelatorio.Caption = '73') then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_VEND' +';'+ 'VENDEDOR' +';'+ 'COD_SUP' +';'+ 'SUPERVISOR' +';'+ 'VIDAS');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_VEND').AsString;
          P2  := qryDados.FIELDBYNAME('VENDEDOR').AsString;
          P3  := qryDados.FIELDBYNAME('COD_SUP').AsString;
          P4  := qryDados.FIELDBYNAME('SUPERVISOR').AsString;
          P5  := qryDados.FIELDBYNAME('VIDAS').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '84' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_QPS' +';'+ 'COD_SEG' +';'+ 'CONTRATO' +';'+ 'NOME' +';'+ 'IDADE' +';'+ 'CONTRATOS' +';'+ 'TIPO' +';'+ 'DT_CANC' +';'+ 'MOTIVO' +';'+ 'VENDEDOR' +';'+ 'TELEFONE' +';'+ 'TELEFONE2' +';'+ 'CELULAR');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_QPS').AsString;
          P2  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P3  := qryDados.FIELDBYNAME('CONTRATO').AsString;
          P4  := qryDados.FIELDBYNAME('NOME').AsString;
          P5  := qryDados.FIELDBYNAME('IDADE').AsString;
          P6  := qryDados.FIELDBYNAME('CONTRATOS').AsString;
          P7  := qryDados.FIELDBYNAME('TIPO').AsString;
          P8  := qryDados.FIELDBYNAME('DT_CANC').AsString;
          P9  := qryDados.FIELDBYNAME('MOTIVO').AsString;
          P10  := qryDados.FIELDBYNAME('VENDEDOR').AsString;
          P11 := qryDados.FIELDBYNAME('TELEFONE').AsString;
          P12 := qryDados.FIELDBYNAME('TELEFONE2').AsString;
          P13 := qryDados.FIELDBYNAME('CELULAR').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';
          cLinha := cLinha + P12 +';';
          cLinha := cLinha + P13 +';';          

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  if (frmPrincipal.lblCodRelatorio.Caption = '80') or (frmPrincipal.lblCodRelatorio.Caption = '81') then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'CODIGO' +';'+ 'NOME' +';'+ 'CPF' +';'+ 'TIPO' +';'+ 'ADITIVO' +';'+ 'SITUACAO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('CODIGO').AsString;
          P2  := qryDados.FIELDBYNAME('NOME').AsString;
          P3  := qryDados.FIELDBYNAME('CPF').AsString;
          P4  := qryDados.FIELDBYNAME('TIPO').AsString;
          P5  := qryDados.FIELDBYNAME('NOME_ADT').AsString;
          P6  := qryDados.FIELDBYNAME('SITUACAO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '48' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'ADESAO' +';'+ 'COD_SEG' +';'+ 'NOME' +';'+ 'TIPO' +';'+ 'CONTRATO' +';'+ 'COD_TITULAR' +';'+ 'ADESAO_EMPRESA' +';'+ 'EMPRESA' +';'+ 'UNIDADE' +';'+ 'CORRETOR' +';'+ 'DT_INICIO' +';'+ 'COD_SET' +';'+ 'SUBCORRETOR' +';'+ 'DIRETOR' +';'+ 'PRESIDENTE');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('ADESAO').AsString;
          P2  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P3  := qryDados.FIELDBYNAME('NOME').AsString;
          P4  := qryDados.FIELDBYNAME('TIPO').AsString;
          P5  := qryDados.FIELDBYNAME('CONTRATO').AsString;
          P6  := qryDados.FIELDBYNAME('COD_TITULAR').AsString;
          P7  := qryDados.FIELDBYNAME('ADESAO_EMPRESA').AsString;
          P8  := qryDados.FIELDBYNAME('EMPRESA').AsString;
          P9  := qryDados.FIELDBYNAME('UNIDADE').AsString;
          P10 := qryDados.FIELDBYNAME('CORRETOR').AsString;
          P11 := qryDados.FIELDBYNAME('DT_INICIO').AsString;
          P12 := qryDados.FIELDBYNAME('COD_SET').AsString;
          P13 := qryDados.FIELDBYNAME('SUBCORRETOR').AsString;
          P14 := qryDados.FIELDBYNAME('DIRETOR').AsString;
          P15 := qryDados.FIELDBYNAME('PRESIDENTE').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';
          cLinha := cLinha + P12 +';';
          cLinha := cLinha + P13 +';';
          cLinha := cLinha + P14 +';';
          cLinha := cLinha + P15 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '70' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'NOME' +';'+ 'DT_INCL' +';'+ 'CORRETOR' +';'+ 'CPF_CORRETOR' +';'+ 'SUPERVISOR' +';'+ 'CPF_SUPERVISOR' +';'+ 'TELEFONE_SUB' +';'+ 'GRUPO_CARENCIA' +';'+ 'VEND_PLATAFORMA');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P2  := qryDados.FIELDBYNAME('NOME').AsString;
          P3  := qryDados.FIELDBYNAME('DT_INCL').AsString;
          P4  := qryDados.FIELDBYNAME('CORRETOR').AsString;
          P5  := qryDados.FIELDBYNAME('CPF_CORRETOR').AsString;
          P6  := qryDados.FIELDBYNAME('SUPERVISOR').AsString;
          P7  := qryDados.FIELDBYNAME('CPF_SUPERVISOR').AsString;
          P8  := qryDados.FIELDBYNAME('TELEFONE').AsString;
          P9  := qryDados.FIELDBYNAME('GRUPO_CARENCIA').AsString;
          P10 := qryDados.FIELDBYNAME('CORRETOR_PLATAFORMA').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '85' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'CODIGO' +';'+ 'NOME' +';'+ 'CPF' +';'+ 'TIPO' +';'+ 'MATRICULA' +';'+ 'SITUACAO' +';'+ 'ADESAO' +';'+ '----|' +';'+ 'DIAS' +';'+ '|----' +';'+ 'CODIGO_CANC' +';'+ 'NOME_CANC' +';'+ 'CPF_CANC' +';'+ 'TIPO_CANC' +';'+ 'MATRICULA_CANC' +';'+ 'SITUACAO_CANC' +';'+ 'ADESAO_CANC' +';'+ 'DT_CANC' +';'+ 'MOTIVO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('CODIGO').AsString;
          P2  := qryDados.FIELDBYNAME('NOME').AsString;
          P3  := qryDados.FIELDBYNAME('CPF').AsString;
          P4  := qryDados.FIELDBYNAME('TIPO').AsString;
          P5  := qryDados.FIELDBYNAME('MATRICULA').AsString;
          P6  := qryDados.FIELDBYNAME('SITUACAO').AsString;
          P7  := qryDados.FIELDBYNAME('ADESAO').AsString;
          P8  := qryDados.FIELDBYNAME('DIAS').AsString;
          P9  := qryDados.FIELDBYNAME('CODIGO_CANC').AsString;
          P10 := qryDados.FIELDBYNAME('NOME_CANC').AsString;
          P11 := qryDados.FIELDBYNAME('CPF_CANC').AsString;
          P12 := qryDados.FIELDBYNAME('TIPO_CANC').AsString;
          P13 := qryDados.FIELDBYNAME('MATRICULA_CANC').AsString;
          P14 := qryDados.FIELDBYNAME('SITUACAO_CANC').AsString;
          P15 := qryDados.FIELDBYNAME('ADESAO_CANC').AsString;
          P16 := qryDados.FIELDBYNAME('DT_CANC').AsString;
          P17 := qryDados.FIELDBYNAME('DESCRICAO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + qryDados.FIELDBYNAME(' --|').AsString +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + qryDados.FIELDBYNAME('|-- ').AsString +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';
          cLinha := cLinha + P12 +';';
          cLinha := cLinha + P13 +';';
          cLinha := cLinha + P14 +';';
          cLinha := cLinha + P15 +';';
          cLinha := cLinha + P16 +';';
          cLinha := cLinha + P17 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '86' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'DESCR_SERV' +';'+ 'COD_SERV' +';'+ 'GUIA' +';'+ 'DATA_LIB' +';'+ 'DATA_REALIZACAO' +';'+ 'PRESTADOR' +';'+ 'COD_AMB' +';'+ 'GLOSA');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('DESCR_SERV').AsString;
          P2  := qryDados.FIELDBYNAME('COD_SERV').AsString;
          P3  := qryDados.FIELDBYNAME('GUIA').AsString;
          P4  := qryDados.FIELDBYNAME('DATA_LIB').AsString;
          P5  := qryDados.FIELDBYNAME('DATA_REALIZACAO').AsString;
          P6  := qryDados.FIELDBYNAME('PRESTADOR').AsString;
          P7  := qryDados.FIELDBYNAME('COD_AMB').AsString;
          P8  := qryDados.FIELDBYNAME('GLOSA').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '88' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'BENEFICIARIO' +';'+ 'ADESAO' +';'+ 'SITUACAO' +';'+ 'SENHA_LIB' +';'+ 'PRESTADOR' +';'+ 'COD_AMB' +';'+ 'DESCRICAO' +';'+ 'DT_LIB' +';'+ 'SERV_SITUACAO' +';'+ 'CONCILIADO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('BENEFICIARIO').AsString;
          P2  := qryDados.FIELDBYNAME('ADESAO').AsString;
          P3  := qryDados.FIELDBYNAME('SITUACAO').AsString;
          P4  := qryDados.FIELDBYNAME('SENHA_LIB').AsString;
          P5  := qryDados.FIELDBYNAME('PRESTADOR').AsString;
          P6  := qryDados.FIELDBYNAME('COD_AMB').AsString;
          P7  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P8  := qryDados.FIELDBYNAME('DT_LIB').AsString;
          P9  := qryDados.FIELDBYNAME('SERV_SITUACAO').AsString;
          P10  := qryDados.FIELDBYNAME('CONCILIADO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';          

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '91' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_VEND' +';'+ 'VENDEDOR' +';'+ 'COD_SUP' +';'+ 'SUPERVISOR' +';'+ 'DT_CAD' +';'+ 'VIDAS_TOTAL' +';'+ 'ATIVAS' +';'+ 'CANCELADAS' +';'+ 'MEDIA_DIAS_ATIVO' +';'+ 'INADIMPLENTES' +';'+ 'ATIVOS_ADIMPLENTES' +';'+ 'RECEITA' +';'+ 'UTILIZACAO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_VEND').AsString;
          P2  := qryDados.FIELDBYNAME('VENDEDOR').AsString;
          P3  := qryDados.FIELDBYNAME('COD_SUP').AsString;
          P4  := qryDados.FIELDBYNAME('SUPERVISOR').AsString;
          P5  := qryDados.FIELDBYNAME('DT_CAD').AsString;
          P6  := qryDados.FIELDBYNAME('VIDAS_TOTAL').AsString;
          P7  := qryDados.FIELDBYNAME('ATIVAS').AsString;
          P8  := qryDados.FIELDBYNAME('CANCELADAS').AsString;
          P9  := qryDados.FIELDBYNAME('MEDIA_DIAS_ATIVO').AsString;
          P10 := qryDados.FIELDBYNAME('INADIMPLENTES').AsString;
          P11 := qryDados.FIELDBYNAME('ATIVOS_ADIMPLENTES').AsString;
          P12 := qryDados.FIELDBYNAME('RECEITA').AsString;
          P13 := qryDados.FIELDBYNAME('UTILIZACAO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';   
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';
          cLinha := cLinha + P12 +';';
          cLinha := cLinha + P13 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '93' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_VEND' +';'+ 'VENDEDOR' +';'+ 'DT_CAD' +';'+ 'VIDAS_TOTAL' +';'+ 'ATIVAS' +';'+ 'CANCELADAS' +';'+ 'MEDIA_DIAS_ATIVO' +';'+ 'INADIMPLENTES' +';'+ 'ATIVOS_ADIMPLENTES' +';'+ 'RECEITA' +';'+ 'UTILIZACAO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_VEND').AsString;
          P2  := qryDados.FIELDBYNAME('VENDEDOR').AsString;
          P3  := qryDados.FIELDBYNAME('DT_CAD').AsString;
          P4  := qryDados.FIELDBYNAME('VIDAS_TOTAL').AsString;
          P5  := qryDados.FIELDBYNAME('ATIVAS').AsString;
          P6  := qryDados.FIELDBYNAME('CANCELADAS').AsString;
          P7  := qryDados.FIELDBYNAME('MEDIA_DIAS_ATIVO').AsString;
          P8  := qryDados.FIELDBYNAME('INADIMPLENTES').AsString;
          P9  := qryDados.FIELDBYNAME('ATIVOS_ADIMPLENTES').AsString;
          P10 := qryDados.FIELDBYNAME('RECEITA').AsString;
          P11 := qryDados.FIELDBYNAME('UTILIZACAO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if (frmPrincipal.lblCodRelatorio.Caption = '92') or (frmPrincipal.lblCodRelatorio.Caption = '94') then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'SUPERVISOR' +';'+ 'DT_CAD_SUP' +';'+ 'VENDEDOR' +';'+ 'DT_CAD_VEND' +';'+ 'VIDAS_TOTAL' +';'+ 'ATIVAS' +';'+ 'CANCELADAS' +';'+ 'MEDIA_DIAS_ATIVO' +';'+ 'INADIMPLENTES');      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('SUPERVISOR').AsString;
          P2  := qryDados.FIELDBYNAME('DT_CAD_SUP').AsString;
          P3  := qryDados.FIELDBYNAME('VENDEDOR').AsString;
          P4  := qryDados.FIELDBYNAME('DT_CAD_VEND').AsString;
          P5  := qryDados.FIELDBYNAME('VIDAS_TOTAL').AsString;
          P6  := qryDados.FIELDBYNAME('ATIVAS').AsString;
          P7  := qryDados.FIELDBYNAME('CANCELADAS').AsString;
          P8  := qryDados.FIELDBYNAME('MEDIA_DIAS_ATIVO').AsString;
          P9  := qryDados.FIELDBYNAME('INADIMPLENTES').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '95' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'COD_BOL' +';'+ 'NOME' +';'+ 'VALOR' +';'+ 'DT_VENC' +';'+ 'EMAIL' +';'+ 'PLANO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P2  := qryDados.FIELDBYNAME('COD_BOL').AsString;
          P3  := qryDados.FIELDBYNAME('NOME').AsString;
          P4  := qryDados.FIELDBYNAME('VALOR').AsString;
          P5  := qryDados.FIELDBYNAME('DT_VENC').AsString;
          P6  := qryDados.FIELDBYNAME('EMAIL').AsString;
          P7  := qryDados.FIELDBYNAME('PLANO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '98' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'DEPENDENTE' +';'+ 'INCLUSAO_DEPENDENTE' +';'+ 'TITULAR' +';'+ 'INCLUSAO_TITULAR' +';'+ 'DT_ANT' +';'+ 'DIA_VENC' +';'+ 'COD_FORMA');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('DEPENDENTE').AsString;
          P2  := qryDados.FIELDBYNAME('DT_CADASTRO_DEP').AsString;
          P3  := qryDados.FIELDBYNAME('TITULAR').AsString;
          P4  := qryDados.FIELDBYNAME('DT_CADASTRO_TIT').AsString;
          P5  := qryDados.FIELDBYNAME('DT_ANT').AsString;
          P6  := qryDados.FIELDBYNAME('DIA_VENC').AsString;
          P7  := qryDados.FIELDBYNAME('COD_FORMA').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '99' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'CHAVE' +';'+ 'GUIA' +';'+ 'DT_REGISTRO' +';'+ 'TIPO' +';'+ 'STATUS');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('ID').AsString;
          P2  := qryDados.FIELDBYNAME('GUIA').AsString;
          P3  := qryDados.FIELDBYNAME('DT_REGISTRO').AsString;
          P4  := qryDados.FIELDBYNAME('TIPO').AsString;
          P5  := qryDados.FIELDBYNAME('STATUS').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '102' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_AMB' +';'+ 'DESCRICAO' +';'+ 'QUANTIDADE' +';'+ 'MOTIVO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_AMB').AsString;
          P2  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P3  := qryDados.FIELDBYNAME('QUANTIDADE').AsString;
          P4  := qryDados.FIELDBYNAME('MOTIVO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '103' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'DT_INCL' +';'+ 'CPF' +';'+ 'IDADE' +';'+ 'MOTIVO' +';'+ 'MESES_ATIVO' +';'+ 'QTD_BOLETO' +';'+ 'RECEITA' +';'+ 'UTILIZACAO' +';'+ 'CONTRATOS' +';'+ 'TELEFONE' +';'+ 'TELEFONE2' +';'+ 'CELULAR' +';'+ 'LOGIN_CANCELAMENTO' +';'+ 'DT_CANC');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P2  := qryDados.FIELDBYNAME('DT_INCL').AsString;
          P3  := qryDados.FIELDBYNAME('CPF').AsString;
          P4  := qryDados.FIELDBYNAME('IDADE').AsString;
          P5  := qryDados.FIELDBYNAME('MOTIVO').AsString;
          P6  := qryDados.FIELDBYNAME('MESES_ATIVO').AsString;
          P7  := qryDados.FIELDBYNAME('QTD_BOLETO').AsString;
          P8  := qryDados.FIELDBYNAME('RECEITA').AsString;
          P9  := qryDados.FIELDBYNAME('UTILIZACAO').AsString;
          P10 := qryDados.FIELDBYNAME('CONTRATOS').AsString;
          P11 := qryDados.FIELDBYNAME('TELEFONE').AsString;
          P12 := qryDados.FIELDBYNAME('TELEFONE2').AsString;
          P13 := qryDados.FIELDBYNAME('CELULAR').AsString;
          P14 := qryDados.FIELDBYNAME('LOGIN_CANCELAMENTO').AsString;
          P15 := qryDados.FIELDBYNAME('DT_CANC').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';
          cLinha := cLinha + P12 +';';
          cLinha := cLinha + P13 +';';
          cLinha := cLinha + P14 +';';
          cLinha := cLinha + P15 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '106' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'NOME' +';'+ 'TIPO' +';'+ 'DT_INICIO' +';'+ 'DT_VENC' +';'+ 'VALOR' +';'+ 'DT_CONCIL' +';'+ 'REVENDA');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P2  := qryDados.FIELDBYNAME('NOME').AsString;
          P3  := qryDados.FIELDBYNAME('TIPO').AsString;
          P4  := qryDados.FIELDBYNAME('DT_INICIO').AsString;
          P5  := qryDados.FIELDBYNAME('DT_VENC').AsString;
          P6  := qryDados.FIELDBYNAME('VALOR').AsString;
          P7  := qryDados.FIELDBYNAME('DT_CONCIL').AsString;
          P8  := qryDados.FIELDBYNAME('REVENDA').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
   end;

   if frmPrincipal.lblCodRelatorio.Caption = '107' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'CODIGO' +';'+ 'PRESTADOR' +';'+ 'QTD_GUIAS');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('CODIGO').AsString;
          P2  := qryDados.FIELDBYNAME('PRESTADOR').AsString;
          P3  := qryDados.FIELDBYNAME('QTD_GUIAS').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
   end;

  if frmPrincipal.lblCodRelatorio.Caption = '117' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'TIPO_REGISTRO' +';'+ 'COD_PLANO' +';'+ 'CODIGO_BENEFICIARIO' +';'+ 'NOME' +';'+ 'CPF' +';'+ 'MAE' +';'+ 'DT_NASC' +';'+ 'SEXO' +';'+ 'EST_CIVIL' +';'+ 'LOGRADOURO' +';'+ 'NUMERO_LOGRADOURO' +';'+ 'COMPLEMENTO' +';'+ 'BAIRRO' +';'+ 'CIDADE' +';'+ 'ESTADO' +';'+ 'CEP' +';'+ 'TIPO_MOVIMENTACAO' +';'+ 'DATA_OPERACAO' +';'+ 'DT_VIGENCIA' +';'+ 'MOTIV_CANC' +';'+ 'LOCACAO' +';'+ 'PARENT' +';'+ 'CPF_TITULAR' +';'+ 'FUNCIONAL_MATRICULA' +';'+ 'MAT');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('TIPO_REGISTRO').AsString;
          P2  := qryDados.FIELDBYNAME('COD_PLANO').AsString;
          P3  := qryDados.FIELDBYNAME('CODIGO_BENEFICIARIO').AsString;
          P4  := qryDados.FIELDBYNAME('NOME').AsString;
          P5  := qryDados.FIELDBYNAME('CPF').AsString;
          P6  := qryDados.FIELDBYNAME('MAE').AsString;
          P7  := qryDados.FIELDBYNAME('DT_NASC').AsString;
          P8  := qryDados.FIELDBYNAME('SEXO').AsString;
          P9  := qryDados.FIELDBYNAME('EST_CIVIL').AsString;
          P10  := qryDados.FIELDBYNAME('LOGRADOURO').AsString;
          P11  := qryDados.FIELDBYNAME('NUMERO_LOGRADOURO').AsString;
          P12  := qryDados.FIELDBYNAME('COMPLEMENTO').AsString;
          P13  := qryDados.FIELDBYNAME('BAIRRO').AsString;
          P14  := qryDados.FIELDBYNAME('CIDADE').AsString;
          P15  := qryDados.FIELDBYNAME('ESTADO').AsString;
          P16  := qryDados.FIELDBYNAME('CEP').AsString;
          P17  := qryDados.FIELDBYNAME('TIPO_MOVIMENTACAO').AsString;
          P18  := qryDados.FIELDBYNAME('DATA_OPERACAO').AsString;
          P19  := qryDados.FIELDBYNAME('DT_VIGENCIA').AsString;
          P20  := qryDados.FIELDBYNAME('MOTIV_CANC').AsString;
          P21  := qryDados.FIELDBYNAME('LOCACAO').AsString;
          P22  := qryDados.FIELDBYNAME('PARENT').AsString;
          P23  := qryDados.FIELDBYNAME('CPF_TITULAR').AsString;
          P24  := qryDados.FIELDBYNAME('FUNCIONAL_MATRICULA').AsString;
          P25  := qryDados.FIELDBYNAME('MAT').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';
          cLinha := cLinha + P12 +';';
          cLinha := cLinha + P13 +';';
          cLinha := cLinha + P14 +';';
          cLinha := cLinha + P15 +';';
          cLinha := cLinha + P16 +';';
          cLinha := cLinha + P17 +';';
          cLinha := cLinha + P18 +';';
          cLinha := cLinha + P19 +';';
          cLinha := cLinha + P20 +';';
          cLinha := cLinha + P21 +';';
          cLinha := cLinha + P22 +';';
          cLinha := cLinha + P23 +';';
          cLinha := cLinha + P24 +';';
          cLinha := cLinha + P25 +';';
          cLinha := cLinha + P26 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if (frmPrincipal.lblCodRelatorio.Caption = '118') or (frmPrincipal.lblCodRelatorio.Caption = '119') then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'DT_INCL' +';'+ 'COD_SEG' +';'+ 'NOME' +';'+ 'TIPO' +';'+ 'DT_TROCA' +';'+ 'CORRETOR' +';'+ 'CPF_CORRETOR' +';'+ 'SUPERVISOR' +';'+ 'CPF_SUPERVISOR' +';'+ 'CORRETOR_PLATAFORMA' +';'+ 'COD_PLANO_ANTIGO' +';'+ 'PLANO_ANTIGO' +';'+ 'COD_PLANO_ATUAL' +';'+ 'PLANO_ATUAL' +';'+ 'VALOR_PLANO_ANTIGO' +';'+ 'VALOR_PLANO_ATUAL');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('DT_INCL').AsString;
          P2  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P3  := qryDados.FIELDBYNAME('NOME').AsString;
          P4  := qryDados.FIELDBYNAME('TIPO').AsString;
          P5  := qryDados.FIELDBYNAME('DT_TROCA').AsString;
          P6  := qryDados.FIELDBYNAME('CORRETOR').AsString;
          P7  := qryDados.FIELDBYNAME('CPF_CORRETOR').AsString;
          P8  := qryDados.FIELDBYNAME('SUPERVISOR').AsString;
          P9  := qryDados.FIELDBYNAME('CPF_SUPERVISOR').AsString;
          P10  := qryDados.FIELDBYNAME('CORRETOR_PLATAFORMA').AsString;
          P11  := qryDados.FIELDBYNAME('COD_PLANO_ANTIGO').AsString;
          P12  := qryDados.FIELDBYNAME('PLANO_ANTIGO').AsString;
          P13  := qryDados.FIELDBYNAME('COD_PLANO_ATUAL').AsString;
          P14  := qryDados.FIELDBYNAME('PLANO_ATUAL').AsString;
          P15  := qryDados.FIELDBYNAME('VALOR_PLANO_ANTIGO').AsString;
          P16  := qryDados.FIELDBYNAME('VALOR_PLANO_ATUAL').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';
          cLinha := cLinha + P12 +';';
          cLinha := cLinha + P13 +';';
          cLinha := cLinha + P14 +';';
          cLinha := cLinha + P15 +';';
          cLinha := cLinha + P16 +';';          

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
   end;

   if frmPrincipal.lblCodRelatorio.Caption = '124' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'DT_CAD' +';'+ 'COD_EMPRESA' +';'+ 'RAZ_SOC' +';'+ 'FANTASIA' +';'+ 'SITUACAO' +';'+ 'COD_UNIDADE' +';'+ 'UNIDADE' +';'+ 'VIDAS' +';'+ 'COD_PLANO'  +';'+ 'PLANO'  +';'+ 'COD_TAB_COMISSAO' +';'+ 'TAB_COMISSAO' +';'+ 'COD_CORRETOR' +';'+ 'CORRETOR' +';'+ 'COD_SUPERVISOR' +';'+ 'SUPERVISOR' +';'+ 'UTILIZACAO' +';'+ 'RECEITA' +';'+ 'COMISSAO' +';'+ 'INADIMPLENCIA' +';'+ 'GRUPO_ATUAL');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('DT_CAD').AsString;
          P2  := qryDados.FIELDBYNAME('COD_EMPRESA').AsString;
          P3  := qryDados.FIELDBYNAME('RAZ_SOC').AsString;
          P4  := qryDados.FIELDBYNAME('FANTASIA').AsString;
          P5  := qryDados.FIELDBYNAME('SITUACAO').AsString;
          P6  := qryDados.FIELDBYNAME('COD_UNIDADE').AsString;
          P7  := qryDados.FIELDBYNAME('UNIDADE').AsString;
          P8  := qryDados.FIELDBYNAME('VIDAS').AsString;
          P9  := qryDados.FIELDBYNAME('COD_PLANO').AsString;
          P10 := qryDados.FIELDBYNAME('PLANO').AsString;
          P11 := qryDados.FIELDBYNAME('COD_TAB_COMISSAO').AsString;
          P12 := qryDados.FIELDBYNAME('TAB_COMISSAO').AsString;
          P13 := qryDados.FIELDBYNAME('COD_CORRETOR').AsString;
          P14 := qryDados.FIELDBYNAME('CORRETOR').AsString;
          P15 := qryDados.FIELDBYNAME('COD_SUPERVISOR').AsString;
          P16 := qryDados.FIELDBYNAME('SUPERVISOR').AsString;
          P17 := qryDados.FIELDBYNAME('UTILIZACAO').AsString;
          P18 := qryDados.FIELDBYNAME('RECEITA').AsString;
          P19 := qryDados.FIELDBYNAME('COMISSAO').AsString;
          P20 := qryDados.FIELDBYNAME('INADIMPLENCIA').AsString;
          P21 := qryDados.FIELDBYNAME('GRUPO_ATUAL').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';
          cLinha := cLinha + P12 +';';
          cLinha := cLinha + P13 +';';
          cLinha := cLinha + P14 +';';
          cLinha := cLinha + P15 +';';
          cLinha := cLinha + P16 +';';
          cLinha := cLinha + P17 +';';
          cLinha := cLinha + P18 +';';
          cLinha := cLinha + P19 +';';
          cLinha := cLinha + P20 +';';
          cLinha := cLinha + P21 +';';          

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '125' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'MES_ANO' +';'+ 'COD_PLANO' +';'+ 'NOME_PLANO' +';'+ 'UTILIZACAO' +';'+ 'GLOSA' +';'+ 'RECEITA');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('MES_ANO').AsString;
          P2  := qryDados.FIELDBYNAME('COD_PLANO').AsString;
          P3  := qryDados.FIELDBYNAME('NOME_PLANO').AsString;
          P4  := qryDados.FIELDBYNAME('UTILIZACAO').AsString;
          P5  := qryDados.FIELDBYNAME('GLOSA').AsString;
          P6  := qryDados.FIELDBYNAME('RECEITA').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '126' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'CONTRATO' +';'+ 'COD_SEG' +';'+ 'TIPO' +';'+ 'NOME' +';'+ 'COD_FORMA' +';'+ 'FORMA_PAG' +';'+ 'COD_GRP' +';'+ 'GRUPO_CAR' +';'+ 'COD_PLANO' +';'+ 'PLANO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('CONTRATO').AsString;
          P2  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P3  := qryDados.FIELDBYNAME('TIPO').AsString;
          P4  := qryDados.FIELDBYNAME('NOME').AsString;
          P5  := qryDados.FIELDBYNAME('COD_FORMA').AsString;
          P6  := qryDados.FIELDBYNAME('FORMA_PAG').AsString;
          P7  := qryDados.FIELDBYNAME('COD_GRP').AsString;
          P8  := qryDados.FIELDBYNAME('GRUPO_CAR').AsString;
          P9  := qryDados.FIELDBYNAME('COD_PLANO').AsString;
          P10 := qryDados.FIELDBYNAME('PLANO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10+';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

   if frmPrincipal.lblCodRelatorio.Caption = '127' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'DT_INCL' +';'+ 'COD_SEG' +';'+ 'NOME' +';'+ 'TIPO' +';'+ 'CORRETOR' +';'+ 'CPF_CORRETOR' +';'+ 'SUPERVISOR' +';'+ 'CPF_SUPERVISOR' +';'+ 'CORRETOR_PLATAFORMA' +';'+ 'COD_PLANO' +';'+ 'PLANO' +';'+ 'VALOR_PLANO' +';'+ 'RECEITA' +';'+ 'UTILIZACAO' +';'+ 'IDADE');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('DT_INCL').AsString;
          P2  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P3  := qryDados.FIELDBYNAME('NOME').AsString;
          P4  := qryDados.FIELDBYNAME('TIPO').AsString;
          P5  := qryDados.FIELDBYNAME('CORRETOR').AsString;
          P6  := qryDados.FIELDBYNAME('CPF_CORRETOR').AsString;
          P7  := qryDados.FIELDBYNAME('SUPERVISOR').AsString;
          P8  := qryDados.FIELDBYNAME('CPF_SUPERVISOR').AsString;
          P9  := qryDados.FIELDBYNAME('CORRETOR_PLATAFORMA').AsString;
          P10  := qryDados.FIELDBYNAME('COD_PLANO').AsString;
          P11  := qryDados.FIELDBYNAME('PLANO').AsString;
          P12  := qryDados.FIELDBYNAME('VALOR_PLANO').AsString;
          P13  := qryDados.FIELDBYNAME('RECEITA').AsString;
          P14  := qryDados.FIELDBYNAME('UTILIZACAO').AsString;
          P15  := qryDados.FIELDBYNAME('IDADE').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';
          cLinha := cLinha + P12 +';';
          cLinha := cLinha + P13 +';';
          cLinha := cLinha + P14 +';';
          cLinha := cLinha + P15 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
   end;

  if frmPrincipal.lblCodRelatorio.Caption = '128' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'NOME' +';'+ 'CPF' +';'+ 'COD_FORMA' +';'+ 'DESCRICAO' +';'+ 'DATA_ENTRADA' +';'+ 'ADESAO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P2  := qryDados.FIELDBYNAME('NOME').AsString;
          P3  := qryDados.FIELDBYNAME('CPF').AsString;
          P4  := qryDados.FIELDBYNAME('COD_FORMA').AsString;
          P5  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P6  := qryDados.FIELDBYNAME('DATA_ENTRADA').AsString;
          P7  := qryDados.FIELDBYNAME('ADESAO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '146' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'SEXO' +';'+ 'REGIAO' +';'+ 'RECEITA' +';'+ 'SINISTRO' +';'+ 'MESANO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P2  := qryDados.FIELDBYNAME('SEXO').AsString;
          P3  := qryDados.FIELDBYNAME('REGIAO').AsString;
          P4  := qryDados.FIELDBYNAME('RECEITA').AsString;
          P5  := qryDados.FIELDBYNAME('SINISTRO').AsString;
          P6  := qryDados.FIELDBYNAME('MESANO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;


  if frmPrincipal.lblCodRelatorio.Caption = '147' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_PLANO' +';'+ 'PLANO' +';'+ 'SEGMENTO' +';'+ 'IDADE' +';'+ 'SEXO' +';'+ 'REGIAO' +';'+ 'RECEITA' +';'+ 'SINISTRO' +';'+ 'MESANO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_PLANO').AsString;
          P2  := qryDados.FIELDBYNAME('PLANO').AsString;
          P3  := qryDados.FIELDBYNAME('TIPO_PLANO').AsString;
          P4  := qryDados.FIELDBYNAME('IDADE').AsString;
          P5  := qryDados.FIELDBYNAME('SEXO').AsString;
          P6  := qryDados.FIELDBYNAME('REGIAO').AsString;
          P7  := qryDados.FIELDBYNAME('RECEITA').AsString;
          P8  := qryDados.FIELDBYNAME('SINISTRO').AsString;
          P9  := qryDados.FIELDBYNAME('MESANO').AsString;

          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '148' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'NUM_GUIA' +';'+ 'SENHA' +';'+ 'COD_SERV' +';'+ 'ENTIDADE' +';'+ 'DT_LIB' +';'+ 'DT_REALIZACAO' +';'+ 'DT_INTERNACAO' +';'+ 'DATA_ALTA' +';'+ 'POSSUI_REGULACAO' +';'+ 'COD_AMB' +';'+ 'DESCRICAO' +';'+ 'VALOR_TOTAL' +';'+ 'SITUACAO' +';'+ 'LOGIN_LIBERACAO' +';'+ 'LOGIN_AUTORIZACAO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('NUM_GUIA').AsString;
          P2  := qryDados.FIELDBYNAME('SENHA').AsString;
          P3  := qryDados.FIELDBYNAME('COD_SERV').AsString;
          P4  := qryDados.FIELDBYNAME('ENTIDADE').AsString;
          P5  := qryDados.FIELDBYNAME('DT_LIB').AsString;
          P6  := qryDados.FIELDBYNAME('DT_REALIZACAO').AsString;
          P7  := qryDados.FIELDBYNAME('DATA_INTERNACAO').AsString;
          P8  := qryDados.FIELDBYNAME('DATA_ALTA').AsString;
          P9  := qryDados.FIELDBYNAME('POSSUI_REGULACAO').AsString;
          P10  := qryDados.FIELDBYNAME('COD_AMB').AsString;
          P11  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P12  := qryDados.FIELDBYNAME('VALOR_TOTAL').AsString;
          P13  := qryDados.FIELDBYNAME('SITUACAO').AsString;
          P14  := qryDados.FIELDBYNAME('LOGIN_LIBERACAO').AsString;
          P15  := qryDados.FIELDBYNAME('LOGIN_AUTORIZACAO').AsString;


          cLinha := '';
          cLinha := cLinha + P1 +';';
          cLinha := cLinha + P2 +';';
          cLinha := cLinha + P3 +';';
          cLinha := cLinha + P4 +';';
          cLinha := cLinha + P5 +';';
          cLinha := cLinha + P6 +';';
          cLinha := cLinha + P7 +';';
          cLinha := cLinha + P8 +';';
          cLinha := cLinha + P9 +';';
          cLinha := cLinha + P10 +';';
          cLinha := cLinha + P11 +';';
          cLinha := cLinha + P12 +';';
          cLinha := cLinha + P13 +';';
          cLinha := cLinha + P14 +';';
          cLinha := cLinha + P15 +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  with qryExec do // INSERIR LOG DE EXPORTA��O
  begin
  Close;
  SQL.Clear;
  SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'E'+#39+', '+#39+parametro+#39+', SYSDATE)');
  ExecSQL;
  end;
 end;
end;

procedure TfrmRelatorio7.DimensionarGrid(dbg: TDbGrid; var formulario);
   type
      TArray = Array of integer;
   procedure AjustarColumns(Swidth,TSize:integer;Asize:TArray);
     var
       idx:integer;
   begin
     if Tsize = 0 then
        begin
           Tsize:=dbg.Columns.Count;
             for idx:=0 to dbg.Columns.Count-1 do
               dbg.Columns[Idx].Width:= (dbg.Canvas.TextWidth('AAAAAA')) div Tsize
        end
     else
      for idx:=0 to dbg.Columns.Count-1 do
        dbg.Columns[Idx].Width := dbg.Columns[Idx].Width + (Swidth*Asize[idx] div Tsize);
    end;
var
   Idx,Twidth,Tsize,Swidth: Integer;
   AWidth:TArray;
   Asize:TArray;
   NomeColuna:String;
begin
   SetLength(AWidth,dbg.Columns.Count);
   SetLength(ASize,dbg.Columns.Count);
   TWidth:=0;
   TSize:=0;
     for Idx := 0 to dbg.Columns.Count - 1  do
        begin
          NomeColuna := Dbg.Columns[Idx].Title.Caption;
          dbg.Columns[Idx].Width := dbg.Canvas.TextWidth(Dbg.Columns[Idx].Title.Caption+'A');
          AWidth[idx]:=dbg.Columns[Idx].Width;
          TWidth:= TWidth + AWidth[idx];
          Asize[idx]:= dbg.Columns[idx].Field.Size;
          Tsize:= Tsize+Asize[idx];
       end;

if dgColLines in dbg.Options then
     TWidth:= TWidth+ Dbg.Columns.Count;

//adiciona a largura da coluna indicada do cursor
if dgIndicator in Dbg.Options then
    TWidth:=TWidth+IndicatorWidth;

Swidth:= dbg.ClientWidth - TWidth;
AjustarColumns(Swidth,TSize,Asize);
end;

procedure TfrmRelatorio7.edtDiaKeyPress(Sender: TObject; var Key: Char);
begin
    if not (key in ['0'..'9',Chr(8), Chr(3), Chr(22), Chr(24), Chr(44)]) then
      key := #0;
end;

procedure TfrmRelatorio7.chkDtConciliacaoClick(Sender: TObject);
begin
  if chkDtConciliacao.Checked = True then
    Med_DataBase.Enabled := True
  else
    Med_DataBase.Enabled := False;
end;

procedure TfrmRelatorio7.Med_DataBaseExit(Sender: TObject);
var d_Data : TDateTime;
begin
  try
    d_Data := StrToDateTime(Med_DataBase.Text);
    Med_DataBase.Text := FormatDateTime('dd/mm/yyyy',d_Data);
  except
    ShowMessage('Data inv�lida');
    Med_DataBase.Text := '  /  /    ';
    Med_DataBase.SetFocus;
  end;
end;

procedure TfrmRelatorio7.qryDadosAfterOpen(DataSet: TDataSet);
begin
frmPrincipal.DimensionarGrid(dbgDados,Self);

  if frmPrincipal.lblCodRelatorio.Caption = '22' then
  begin
    TCurrencyField(qryDados.FieldByName('COD_SEG')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DATA_ADESAO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_INICIO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('MENSALIDADE')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('DATA_ENTRADA')).Alignment := taCenter;    
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '24' then
  begin
    TCurrencyField(qryDados.FieldByName('COD_SEG')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_CANC')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_PRODUTO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_INCLUSAO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DATA_ENTRADA')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_FORMA')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VALOR')).DisplayFormat := 'R$ #,##0.00';
  end;

  if (frmPrincipal.lblCodRelatorio.Caption = '91') or (frmPrincipal.lblCodRelatorio.Caption = '92') or
     (frmPrincipal.lblCodRelatorio.Caption = '93') or (frmPrincipal.lblCodRelatorio.Caption = '94') then
  begin
    TCurrencyField(qryDados.FieldByName('VIDAS_TOTAL')).DisplayFormat := '#,##0';
    TCurrencyField(qryDados.FieldByName('VIDAS_TOTAL')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('ATIVAS')).DisplayFormat := '#,##0';
    TCurrencyField(qryDados.FieldByName('ATIVAS')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('CANCELADAS')).DisplayFormat := '#,##0';
    TCurrencyField(qryDados.FieldByName('CANCELADAS')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('MEDIA_DIAS_ATIVO')).DisplayFormat := '#,##0';
    TCurrencyField(qryDados.FieldByName('MEDIA_DIAS_ATIVO')).Alignment := taCenter;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '95' then
  begin
    TCurrencyField(qryDados.FieldByName('COD_SEG')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_BOL')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VALOR')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('VALOR')).Alignment := taRightJustify;
    TCurrencyField(qryDados.FieldByName('DT_VENC')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('PLANO')).Alignment := taCenter;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '98' then
  begin
    TCurrencyField(qryDados.FieldByName('DEPENDENTE')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_CADASTRO_DEP')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('TITULAR')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_CADASTRO_TIT')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_ANT')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DIA_VENC')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_FORMA')).Alignment := taCenter;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '99' then
  begin
    TCurrencyField(qryDados.FieldByName('ID')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('GUIA')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_REGISTRO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('STATUS')).Alignment := taCenter;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '103' then
  begin
    TCurrencyField(qryDados.FieldByName('COD_SEG')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_INCL')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('CPF')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('IDADE')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('MOTIVO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('MESES_ATIVO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('QTD_BOLETO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('RECEITA')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('UTILIZACAO')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('CONTRATOS')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('TELEFONE')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('TELEFONE2')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('CELULAR')).Alignment := taCenter;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '106' then
  begin
    TCurrencyField(qryDados.FieldByName('COD_SEG')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('NOME')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('TIPO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_INICIO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_VENC')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VALOR')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('VALOR')).Alignment := taRightJustify;
    TCurrencyField(qryDados.FieldByName('DT_CONCIL')).Alignment := taCenter;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '107' then
  begin
    TCurrencyField(qryDados.FieldByName('CODIGO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('QTD_GUIAS')).DisplayFormat := '#,##0';
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '122' then
  begin
    TCurrencyField(qryDados.FieldByName('CODIGO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('QTD_GUIAS')).DisplayFormat := '#,##0';
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '124' then
  begin
    TCurrencyField(qryDados.FieldByName('DT_CAD')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_EMPRESA')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('SITUACAO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_UNIDADE')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VIDAS')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_PLANO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_TAB_COMISSAO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_CORRETOR')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_SUPERVISOR')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('UTILIZACAO')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('RECEITA')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('COMISSAO')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('INADIMPLENCIA')).DisplayFormat := 'R$ #,##0.00';    
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '125' then
  begin
    TCurrencyField(qryDados.FieldByName('MES_ANO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_PLANO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('UTILIZACAO')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('GLOSA')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('RECEITA')).DisplayFormat := 'R$ #,##0.00';
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '126' then
  begin
    TCurrencyField(qryDados.FieldByName('CONTRATO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_SEG')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('TIPO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_FORMA')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_GRP')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_PLANO')).Alignment := taCenter;                
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '127' then
  begin
    TCurrencyField(qryDados.FieldByName('DT_INCL')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_SEG')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('TIPO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('CPF_CORRETOR')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('CPF_SUPERVISOR')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_PLANO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VALOR_PLANO')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('RECEITA')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('UTILIZACAO')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('IDADE')).Alignment := taCenter;    
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '128' then
  begin
    TCurrencyField(qryDados.FieldByName('COD_SEG')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_FORMA')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DATA_ENTRADA')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('ADESAO')).Alignment := taCenter;
  end;

end;

procedure TfrmRelatorio7.ChkPlanoClick(Sender: TObject);
begin
  if ChkPlano.Checked = True then
  begin
    EdtPlano.Enabled := True;
    EdtPlano.Color   := clWindow; //cl3DLight
  end else
  begin
    EdtPlano.Enabled := False;
    EdtPlano.Color   := cl3DLight;
  end;


end;

end.
