unit Relatorio4;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, Grids, DBGrids, StdCtrls, Buttons, ComCtrls, FuncDulio, ExtCtrls, Menus, QRPrntr;

type
  TfrmRelatorio4 = class(TForm)
    GroupBox1: TGroupBox;
    lblDataBase: TLabel;
    btnGerar: TSpeedButton;
    edtProcedimento: TEdit;
    Label1: TLabel;
    btnExportar: TSpeedButton;
    btnImprimir: TSpeedButton;
    lblTotal: TLabel;
    lblRTotal: TLabel;
    dbgDados: TDBGrid;
    dsDados: TDataSource;
    qryDados: TADOQuery;
    SaveDialog1: TSaveDialog;
    qryExec: TADOQuery;
    lblParam: TLabel;
    edtNome: TEdit;
    edtCodigo: TEdit;
    lblCodigo: TLabel;
    lblLegenda: TLabel;
    procedure edtProcedimentoKeyPress(Sender: TObject; var Key: Char);
    procedure edtProcedimentoExit(Sender: TObject);
    procedure btnGerarClick(Sender: TObject);
    procedure btnExportarClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure btnImprimirClick(Sender: TObject);
    procedure qryDadosAfterOpen(DataSet: TDataSet);
    procedure edtCodigoKeyPress(Sender: TObject; var Key: Char);
    procedure edtCodigoExit(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    vPaperSize : TQRPaperSize;
    vLength : Extended;
  end;

var
  frmRelatorio4: TfrmRelatorio4;
  parametro: String;
  valor, glosa : Variant;
implementation

uses MaskUtils, Math, Principal, Relat60, DataModule, Relat59, Relat100, Aguarde,
  Relat116, QuickRpt;

{$R *.dfm}
procedure TfrmRelatorio4.edtProcedimentoKeyPress(Sender: TObject;
  var Key: Char);
begin
    if not (key in ['0'..'9',Chr(8), Chr(3), Chr(22), Chr(24), Chr(44)]) then
      key := #0;
end;

procedure TfrmRelatorio4.edtProcedimentoExit(Sender: TObject);
begin
{  if frmPrincipal.lblCodRelatorio.Caption = '10' then
  begin
    if Length(edtProcedimento.Text) = 8 then
    begin
    edtProcedimento.Text := FormatMaskText('9.99.99.99-9;0;_',edtProcedimento.Text);
    end else
    begin
    ShowMessage('C�digo TUSS inv�lido');
    end;
  end;

  if (frmPrincipal.lblCodRelatorio.Caption = '54') or (frmPrincipal.lblCodRelatorio.Caption = '97') then
  begin
    if Length(edtProcedimento.Text) = 11 then
    begin
    edtProcedimento.Text := FormatMaskText('999.999.999-99;0;_',edtProcedimento.Text);
    end else
    begin
    ShowMessage('CPF inv�lido');
    end;
  end; }
end;

procedure TfrmRelatorio4.btnGerarClick(Sender: TObject);
var p : String;
begin
  parametro := '';
  p := '';

  if frmPrincipal.lblCodRelatorio.Caption = '10' then
  begin
    p := StringReplace(StringReplace(edtProcedimento.Text,'-','',[rfReplaceAll]),'.','',[rfReplaceAll]);

    if Length(p) = 8 then
    begin
      parametro := 'C�digo TUSS: '+p;

      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&codigo_amb',p,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;
      lblRTotal.Caption := IntToStr(qryDados.RecordCount);

      with qryExec do // INSERIR LOG DE VISUALIZA��O
      begin
      Close;
      SQL.Clear;
      SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+','+#39+parametro+#39+', SYSDATE)');
      ExecSQL;
      end;
    end else
    begin
    ShowMessage('C�digo TUSS inv�lido');
    end;
  end;

  if (frmPrincipal.lblCodRelatorio.Caption = '54') or (frmPrincipal.lblCodRelatorio.Caption = '97')then
  begin
    p := StringReplace(StringReplace(edtProcedimento.Text,'-','',[rfReplaceAll]),'.','',[rfReplaceAll]);
    if Length(p) = 11 then
    begin
      parametro := 'CPF: '+p;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&CPF',p,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;
      lblRTotal.Caption := IntToStr(qryDados.RecordCount);

      with qryExec do // INSERIR LOG DE VISUALIZA��O
      begin
      Close;
      SQL.Clear;
      SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+','+#39+parametro+#39+', SYSDATE)');
      ExecSQL;
      end;
    end else
    begin
    ShowMessage('CPF inv�lido.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '59' then
  begin
    if edtProcedimento.Text <> '' then
    begin
      parametro := 'Fatura: '+edtProcedimento.Text;
      with DM.qryRelat59 do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&FATURA',edtProcedimento.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;

      with qryExec do  // Para Quantidade
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL2.Text);
      SQL.Text := StringReplace(SQL.Text,'&FATURA',edtProcedimento.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;

      lblRTotal.Caption := IntToStr(qryExec.RecordCount);

      valor := 0;
      glosa := 0;
      DM.qryRelat59.First;
      while not DM.qryRelat59.Eof do
      begin
        valor := valor + DM.qryRelat59.FieldByName('VALOR').AsVariant;
        glosa := glosa + DM.qryRelat59.FieldByName('VALOR_GLOSADO').AsVariant;
        DM.qryRelat59.Next;
      end;

      with qryExec do // INSERIR LOG DE VISUALIZA��O
      begin
      Close;
      SQL.Clear;
      SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+','+#39+parametro+#39+', SYSDATE)');
      ExecSQL;
      end;
      frmPrincipal.DimensionarGrid(dbgDados,Self);
    end else
    begin
    ShowMessage('Favor inforar o n�mero da fatura.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '60' then
  begin
    if edtProcedimento.Text <> '' then
    begin
      parametro := 'Fatura: '+edtProcedimento.Text;
      with DM.qryRelat60 do  // Para impressao de relatorio
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&FATURA',edtProcedimento.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;

      with qryExec do  // Para Quantidade
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL2.Text);
      SQL.Text := StringReplace(SQL.Text,'&FATURA',edtProcedimento.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;

      lblRTotal.Caption := IntToStr(qryExec.RecordCount);

      valor := 0;
      glosa := 0;
      DM.qryRelat60.First;
      while not DM.qryRelat60.Eof do
      begin
        valor := valor + DM.qryRelat60.FieldByName('VALOR').AsVariant;
        glosa := glosa + DM.qryRelat60.FieldByName('VALOR_GLOSADO').AsVariant;
        DM.qryRelat60.Next;
      end;

      with qryExec do // INSERIR LOG DE VISUALIZA��O
      begin
      Close;
      SQL.Clear;
      SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+','+#39+parametro+#39+', SYSDATE)');
      ExecSQL;
      end;
    end else
    begin
    ShowMessage('Favor inforar o n�mero da fatura.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '61' then
  begin
    if edtProcedimento.Text <> '' then
    begin
      parametro := 'Dia: '+edtProcedimento.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&DIA',edtProcedimento.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;
      lblRTotal.Caption := IntToStr(qryDados.RecordCount);

      with qryExec do // INSERIR LOG DE VISUALIZA��O
      begin
      Close;
      SQL.Clear;
      SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+','+#39+parametro+#39+', SYSDATE)');
      ExecSQL;
      end;
    end else
    begin
    ShowMessage('Favor informar n�mero de dias de atraso..');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '96' then
  begin
    if edtProcedimento.Text <> '' then
    begin
      parametro := 'Cod Geracao: '+edtProcedimento.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&CODIGO',edtProcedimento.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco
      Open;
      end;
      lblRTotal.Caption := IntToStr(qryDados.RecordCount);

      with qryExec do // INSERIR LOG DE VISUALIZA��O
      begin
      Close;
      SQL.Clear;
      SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+','+#39+parametro+#39+', SYSDATE)');
      ExecSQL;
      end;
    end else
    begin
    ShowMessage('Favor informar o c�digo de gera��o.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '100' then
  begin
    if edtProcedimento.Text <> '' then
    begin
      parametro := 'Boleto: '+edtProcedimento.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&CODIGO',edtProcedimento.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco

      if lblCodigo.Caption <> 'X' then
      begin
        SQL.Add('AND S.COD_SEG = '+#39+lblCodigo.Caption+#39);
        parametro := parametro + ' Benefici�rio: ' + lblCodigo.Caption;
      end;

      SQL.Add(' ORDER BY S.COD_FAM, DECODE(S.TIPO,'+#39+'E'+#39+','+#39+'TIT.'+#39+','+#39+'D'+#39+','+#39+'DEP.'+#39+') DESC, D.TIPO');

      Open;
      end;
      lblRTotal.Caption := IntToStr(qryDados.RecordCount);

      with qryExec do // INSERIR LOG DE VISUALIZA��O
      begin
      Close;
      SQL.Clear;
      SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+','+#39+parametro+#39+', SYSDATE)');
      ExecSQL;
      end;

      lblParam.Caption := edtProcedimento.Text;
    end else
    begin
    ShowMessage('Favor informar n�mero do boleto.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '116' then
  begin
    edtCodigoExit(Sender);
    if edtProcedimento.Text <> '' then
    begin
      parametro := 'Boleto: '+edtProcedimento.Text;
      with qryDados do
      begin
      Close;
      SQL.Clear;
      SQL.Add(frmPrincipal.memSQL.Text);
      SQL.Text := StringReplace(SQL.Text,'&CODIGO',edtProcedimento.Text,[rfReplaceAll]); //Alterar parametro dentro do SQL puxado do banco

      if lblCodigo.Caption <> 'X' then
      begin
        SQL.Add('and s.cod_fam = '+#39+lblCodigo.Caption+#39);
        parametro := parametro + ' Benefici�rio: ' + lblCodigo.Caption;
      end;
      SQL.Add(' GROUP BY S.COD_FAM, B.COD_BOL, DECODE(S.TIPO,'+#39+'E'+#39+','+#39+'TIT.'+#39+','+#39+'D'+#39+','+#39+'DEP.'+#39+'), S.COD_SEG, S.NOME');
      SQL.Add(' ORDER BY S.COD_FAM, DECODE(S.TIPO,'+#39+'E'+#39+','+#39+'TIT.'+#39+','+#39+'D'+#39+','+#39+'DEP.'+#39+') DESC');

      Open;
      end;
      lblRTotal.Caption := IntToStr(qryDados.RecordCount);

      with qryExec do // INSERIR LOG DE VISUALIZA��O
      begin
      Close;
      SQL.Clear;
      SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+','+#39+parametro+#39+', SYSDATE)');
      ExecSQL;
      end;

      lblParam.Caption := edtProcedimento.Text;
    end else
    begin
    ShowMessage('Favor informar n�mero do boleto.');
    end;
  end;

end;

procedure TfrmRelatorio4.btnExportarClick(Sender: TObject);
var
F: TextFile;
P1, P2, P3, P4, P5, P6, P7, P8, P9, P10, P11, P12, P13, P14, cLinha : String;
begin

with qryExec do // INSERIR LOG DE EXPORTA��O
begin
Close;
SQL.Clear;
SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'E'+#39+', '+#39+parametro+#39+', SYSDATE)');
ExecSQL;
end;

if frmPrincipal.lblCodRelatorio.Caption = '59' then
begin
  if DM.qryRelat59.RecordCount > 0 then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'NU_PROTOCOLO' +';'+ 'SENHA_LIB' +';'+ 'DATA' +';'+ 'NOME' +';'+ 'COD_AMB' +';'+ 'VALOR' +';'+ 'VALOR_GLOSADO' +';'+ 'COD_TISS_GLOSA' +';'+ 'DESCRICAO' +';'+ 'FATURA');
      DM.qryRelat59.First;
      while not DM.qryRelat59.Eof do
        begin
          P1  := DM.qryRelat59.FIELDBYNAME('NU_PROTOCOLO').AsString;
          P2  := DM.qryRelat59.FIELDBYNAME('SENHA_LIB').AsString;
          P3  := DM.qryRelat59.FIELDBYNAME('DATA').AsString;
          P4  := DM.qryRelat59.FIELDBYNAME('NOME').AsString;
          P5  := DM.qryRelat59.FIELDBYNAME('COD_AMB').AsString;
          P6  := DM.qryRelat59.FIELDBYNAME('VALOR').AsString;
          P7  := DM.qryRelat59.FIELDBYNAME('VALOR_GLOSADO').AsString;
          P8  := DM.qryRelat59.FIELDBYNAME('COD_TISS_GLOSA').AsString;
          P9  := DM.qryRelat59.FIELDBYNAME('DESCRICAO').AsString;
          P10 := DM.qryRelat59.FIELDBYNAME('FATURA').AsString;

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
          DM.qryRelat59.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;
  Abort;
end;

if frmPrincipal.lblCodRelatorio.Caption = '60' then
begin
  if DM.qryRelat60.RecordCount > 0 then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'SENHA_LIB' +';'+ 'DATA' +';'+ 'NOME' +';'+ 'COD_AMB' +';'+ 'VALOR' +';'+ 'VALOR_GLOSADO' +';'+ 'COD_TISS_GLOSA' +';'+ 'DESCRICAO' +';'+ 'FATURA');
      DM.qryRelat60.First;
      while not DM.qryRelat60.Eof do
        begin
          P1 := DM.qryRelat60.FIELDBYNAME('SENHA_LIB').AsString;
          P2 := DM.qryRelat60.FIELDBYNAME('DATA').AsString;
          P3 := DM.qryRelat60.FIELDBYNAME('NOME').AsString;
          P4 := DM.qryRelat60.FIELDBYNAME('COD_AMB').AsString;
          P5 := DM.qryRelat60.FIELDBYNAME('VALOR').AsString;
          P6 := DM.qryRelat60.FIELDBYNAME('VALOR_GLOSADO').AsString;
          P7 := DM.qryRelat60.FIELDBYNAME('COD_TISS_GLOSA').AsString;
          P8 := DM.qryRelat60.FIELDBYNAME('DESCRICAO').AsString;
          P9 := DM.qryRelat60.FIELDBYNAME('FATURA').AsString;

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
          DM.qryRelat60.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;
  Abort;
end;

 if qryDados.RecordCount > 0 then
 begin

   if frmPrincipal.lblCodRelatorio.Caption = '10' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'CODIGO' +';'+ 'FANTASIA' +';'+ 'COD_SERV' +';'+ 'DESCRICAO' +';'+ 'DT_INICIO' +';'+ 'VALOR');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1 := qryDados.FIELDBYNAME('CODIGO').AsString;
          P2 := qryDados.FIELDBYNAME('FANTASIA').AsString;
          P3 := qryDados.FIELDBYNAME('COD_SERV').AsString;
          P4 := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P5 := qryDados.FIELDBYNAME('DT_INICIO').AsString;
          P6 := qryDados.FIELDBYNAME('VALOR').AsString;

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

   if (frmPrincipal.lblCodRelatorio.Caption = '54') or (frmPrincipal.lblCodRelatorio.Caption = '96') then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'NOME' +';'+ 'CPF' +';'+ 'TIPO' +';'+ 'COD_VEND' +';'+ 'CORRETOR' +';'+ 'DT_INCL' +';'+ 'DT_CANC' +';'+ 'PLANO' +';'+ 'DESCRICAO' +';'+ 'QTD' +';'+ 'RECEITA' +';'+ 'DESPESA' +';'+ 'UTILIZACAO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P2  := qryDados.FIELDBYNAME('NOME').AsString;
          P3  := qryDados.FIELDBYNAME('CPF').AsString;
          P4  := qryDados.FIELDBYNAME('TIPO').AsString;
          P5  := qryDados.FIELDBYNAME('COD_VEND').AsString;
          P6  := qryDados.FIELDBYNAME('CORRETOR').AsString;
          P7  := qryDados.FIELDBYNAME('DT_INCL').AsString;
          P8  := qryDados.FIELDBYNAME('DT_CANC').AsString;
          P9  := qryDados.FIELDBYNAME('PLANO').AsString;
          P10 := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P11 := qryDados.FIELDBYNAME('QTD').AsString;
          P12 := qryDados.FIELDBYNAME('RECEITA').AsString;
          P13 := qryDados.FIELDBYNAME('DESPESA').AsString;
          P14 := qryDados.FIELDBYNAME('UTILIZACAO').AsString;

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

  if frmPrincipal.lblCodRelatorio.Caption = '61' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SET' +';'+ 'EMPRESA' +';'+ 'FANTASIA' +';'+ 'DT_VENC' +';'+ 'VALOR' +';'+ 'COD_BOL');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1  := qryDados.FIELDBYNAME('COD_SET').AsString;
          P2  := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P3  := qryDados.FIELDBYNAME('SIGLA').AsString;
          P4  := qryDados.FIELDBYNAME('DT_VENC').AsString;
          P5  := qryDados.FIELDBYNAME('VALOR').AsString;
          P6  := qryDados.FIELDBYNAME('COD_BOL').AsString;

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

  if frmPrincipal.lblCodRelatorio.Caption = '96' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha {+ 'COD_SEG' +';'}+ 'COD_BOL' +';'+ 'NOME' +';'+ 'VALOR' +';'+ 'DT_VENC' +';'+ 'EMAIL' +';'+ 'PLANO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          //P1  := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P2  := qryDados.FIELDBYNAME('COD_BOL').AsString;
          P3  := qryDados.FIELDBYNAME('NOME').AsString;
          P4  := qryDados.FIELDBYNAME('VALOR').AsString;
          P5  := qryDados.FIELDBYNAME('DT_VENC').AsString;
          P6  := qryDados.FIELDBYNAME('EMAIL').AsString;
          P7  := qryDados.FIELDBYNAME('PLANO').AsString;

          cLinha := '';
          //cLinha := cLinha + P1 +';';
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

   if frmPrincipal.lblCodRelatorio.Caption = '100' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_FAM' +';'+ 'COD_BOL' +';'+ 'TIPO' +';'+ 'COD_SEG' +';'+ 'NOME' +';'+ 'DESCRICAO' +';'+ 'VALOR');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1 := qryDados.FIELDBYNAME('COD_FAM').AsString;
          P2 := qryDados.FIELDBYNAME('COD_BOL').AsString;
          P3 := qryDados.FIELDBYNAME('TIPO').AsString;
          P4 := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P5 := qryDados.FIELDBYNAME('NOME').AsString;
          P6 := qryDados.FIELDBYNAME('DESCRICAO').AsString;
          P7 := qryDados.FIELDBYNAME('VALOR').AsString;

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

   if frmPrincipal.lblCodRelatorio.Caption = '116' then
   begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_FAM' +';'+ 'COD_BOL' +';'+ 'TIPO' +';'+ 'COD_SEG' +';'+ 'NOME' +';'+ 'VALOR');
      qryDados.First;
      while not qryDados.Eof do
        begin
          P1 := qryDados.FIELDBYNAME('COD_FAM').AsString;
          P2 := qryDados.FIELDBYNAME('COD_BOL').AsString;
          P3 := qryDados.FIELDBYNAME('TIPO').AsString;
          P4 := qryDados.FIELDBYNAME('COD_SEG').AsString;
          P5 := qryDados.FIELDBYNAME('NOME').AsString;
          P6 := qryDados.FIELDBYNAME('VALOR').AsString;

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

 end;
end;

procedure TfrmRelatorio4.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
qryDados.Close;
end;

procedure TfrmRelatorio4.FormShow(Sender: TObject);
begin
  dsDados.DataSet := qryDados;
  edtCodigo.Visible := False;
  edtNome.Visible := False;
  lblLegenda.Visible := False;
  lblCodigo.Caption := 'X';

  if frmPrincipal.lblExporta.Caption = 'T' then //Habilita bot�o Exportar
  begin
  btnExportar.Enabled := True;
  end else
  begin
  btnExportar.Enabled := False;
  end;

  if frmPrincipal.lblImprime.Caption = 'T' then //Habilita bot�o Imprimir
  begin
  btnImprimir.Enabled := True;
  btnImprimir.Visible := True;
  end else
  begin
  btnImprimir.Enabled := False;
  btnImprimir.Visible := False;
  end;

  with qryExec do
  begin
  Close;
  SQL.Clear;
  SQL.Add('ALTER SESSION SET NLS_DATE_FORMAT = '+#39+'DD/MM/YYYY HH24:MI:SS'+#39);
  ExecSQL;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '10' then
  lblDataBase.Caption := 'C�digo TUSS: ';

  if (frmPrincipal.lblCodRelatorio.Caption = '54') or (frmPrincipal.lblCodRelatorio.Caption = '97') then
  lblDataBase.Caption := 'CPF: ';

  if (frmPrincipal.lblCodRelatorio.Caption = '59') or (frmPrincipal.lblCodRelatorio.Caption = '60') then
  lblDataBase.Caption := 'FATURA: ';

  if frmPrincipal.lblCodRelatorio.Caption = '61' then
  lblDataBase.Caption := 'Dias: ';

  if frmPrincipal.lblCodRelatorio.Caption = '96' then
  lblDataBase.Caption := 'C�d. Gera��o: ';

  if (frmPrincipal.lblCodRelatorio.Caption = '100') or (frmPrincipal.lblCodRelatorio.Caption = '116') then
  begin
    lblDataBase.Caption := 'C�d. Boleto: ';
    edtCodigo.Visible := True;
    edtNome.Visible := True;
    lblLegenda.Visible := True;

    if frmPrincipal.lblCodRelatorio.Caption = '100' then
      lblLegenda.Caption := 'C�digo Benefici�rio'
    else
      lblLegenda.Caption := 'C�digo Titular';
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '59' then
  dsDados.DataSet := DM.qryRelat59;

  if frmPrincipal.lblCodRelatorio.Caption = '60' then
  dsDados.DataSet := DM.qryRelat60;

  edtProcedimento.Text := '';
  lblRTotal.Caption := '0';

  DM.qryRelat60.Close;
  DM.qryRelat59.Close;
  qryDados.Close;

  lblParam.Caption := '';
end;

procedure TfrmRelatorio4.btnImprimirClick(Sender: TObject);
var Cod_Prest, Prestador, MesPag : String;
Geral, Titular, Dependente: Real;
JanelasCongeladas: Pointer;
begin
  Cod_Prest := '';
  Prestador := '';
  MesPag    := '';
  with qryExec do // INSERIR LOG DE IMPRESSAO
  begin
  Close;
  SQL.Clear;
  SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, PARAMETRO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'I'+#39+', '+#39+parametro+#39+', SYSDATE)');
  ExecSQL;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '59' then
  begin
    with qryExec do
    begin
      Close;
      SQL.Clear;
      SQL.Add('SELECT V.COD_ENT, V.NOME, F.MES_PGTO');
      SQL.Add('FROM IM_FAT F, TI.VW_PRESTADORES V');
      SQL.Add('WHERE F.TIPO_ENT = V.TIPO_ENT');
      SQL.Add('AND   F.COD_ENT  = V.COD_ENT');
      SQL.Add('AND   F.CHAVE    = '+DM.qryRelat59.FieldByName('FATURA').AsVariant);
      SQL.Add('GROUP BY V.COD_ENT, V.NOME, F.MES_PGTO');
      Open;
    end;

    Prestador := qryExec.FieldByName('NOME').AsString;
    MesPag    := qryExec.FieldByName('MES_PGTO').AsString;
    Cod_Prest := qryExec.FieldByName('COD_ENT').AsString;
    qrpRelat59.qrlblTitulo.Caption := 'Relat�rio de Glosa - Pagamento ' + MesPag + ' - ' + Cod_Prest +' '+ Prestador;
    qrpRelat59.qrlblQuantidade.Caption := 'Quantidade: ' + lblRTotal.Caption;
    qrpRelat59.qrlblValor.Caption := 'Valor Total: ' + 'R$ ' + FormatFloat('#,0.00',valor);
    qrpRelat59.qrlblGlosa.Caption := 'Glosa Total: ' + 'R$ ' + FormatFloat('#,0.00',glosa);
    vPaperSize := qrpRelat59.Page.PaperSize;
    vLength    := qrpRelat59.Page.Length;
    qrpRelat59.Prepare;
    qrpRelat59.qrlblPagina.Caption := IntToStr(qrpRelat59.QRPrinter.PageCount);
    qrpRelat59.PreviewModal;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '60' then
  begin
    with qryExec do
    begin
      Close;
      SQL.Clear;
      SQL.Add('SELECT V.NOME, F.MES_PGTO');
      SQL.Add('FROM IM_FAT F, TI.VW_PRESTADORES V');
      SQL.Add('WHERE F.TIPO_ENT = V.TIPO_ENT');
      SQL.Add('AND   F.COD_ENT  = V.COD_ENT');
      SQL.Add('AND   F.CHAVE    = '+DM.qryRelat60.FieldByName('FATURA').AsVariant);
      SQL.Add('GROUP BY V.NOME, F.MES_PGTO');
      Open;
    end;

    Prestador := qryExec.FieldByName('NOME').AsString;
    MesPag    := qryExec.FieldByName('MES_PGTO').AsString;

    qrpRelat60.qrlblTitulo.Caption := 'Relat�rio de Glosa - Pagamento ' + MesPag + ' - ' + Prestador;
    qrpRelat60.qrlblQuantidade.Caption := 'Quantidade: ' + lblRTotal.Caption;
    qrpRelat60.qrlblValor.Caption := 'Valor Total: ' + 'R$ ' + FormatFloat('#,0.00',valor);
    qrpRelat60.qrlblGlosa.Caption := 'Glosa Total: ' + 'R$ ' + FormatFloat('#,0.00',glosa);
    qrpRelat60.Prepare;
    qrpRelat60.qrlblPagina.Caption := IntToStr(qrpRelat60.QRPrinter.PageCount);
    qrpRelat60.PreviewModal;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '100' then
  begin
    qrpRelat100.qrlblTitulo.Caption := 'Extrato do Boleto (' + lblParam.Caption + ')';

    with DM.qryRelat100Tit do
    begin
      Close;
      SQL.Clear;
      SQL.Add('SELECT '+#39+' TITULAR: '+#39+' || T.COD_SEG || '+#39+' - '+#39+' || T.NOME TITULAR, TO_CHAR(B.COD_BOL) BOLETO, T.COD_SEG COD_TITULAR');

      if lblCodigo.Caption <> 'X' then
        SQL.Add('      ,S.COD_SEG||'+#39+'%'+#39+' COD_SEGX')
      else
        SQL.Add('      ,'+#39+'%'+#39+' COD_SEGX');

      SQL.Add('  FROM IM_BOLAB B');
      SQL.Add('  JOIN IM_DBOL  D ON B.COD_BOL = D.COD_BOL');
      SQL.Add('  JOIN IM_SEG   S ON D.COD_SEG = S.COD_SEG');
      SQL.Add('  JOIN IM_SEG   T ON S.COD_FAM = T.COD_SEG');
      SQL.Add('  JOIN IM_VTCN1 V');
      SQL.Add('    ON SUBSTR(D.CODIGO,INSTR(D.CODIGO,'+#39+';'+#39+',1,2)+1,INSTR(SUBSTR(D.CODIGO,INSTR(D.CODIGO,'+#39+';'+#39+',1,2)+1),'+#39+';'+#39+')-1) = V.COD_SERV');
      SQL.Add('   AND SUBSTR(D.CODIGO,1,INSTR(D.CODIGO,'+#39+';'+#39+',1)-1) = V.TIPO_ENT');
      SQL.Add('   AND SUBSTR(D.CODIGO,INSTR(D.CODIGO,'+#39+';'+#39+',1)+1,INSTR(D.CODIGO,'+#39+';'+#39+',1,2)-3) = V.TIPO_SERV');
      SQL.Add(' WHERE B.COD_BOL = '+lblParam.Caption);

      if lblCodigo.Caption <> 'X' then
        SQL.Add('AND S.COD_SEG = '+#39+lblCodigo.Caption+#39);

      SQL.Add('   AND T.TIPO = '+#39+'E'+#39);
      SQL.Add(' GROUP BY '+#39+' TITULAR: '+#39+' || T.COD_SEG || '+#39+' - '+#39+' || T.NOME, TO_CHAR(B.COD_BOL) , T.COD_SEG');

      if lblCodigo.Caption <> 'X' then
        SQL.Add('      ,S.COD_SEG||'+#39+'%'+#39)
      else
        SQL.Add('      ,'+#39+'%'+#39);
      Open;
    end;

    DM.qryRelat100Detalhe.Open;

    with qryExec do
    begin
      Close;
      SQL.Clear;
      SQL.Add('SELECT COD_BOL, SUM(VLR_TITULAR) VLR_TITULAR, SUM(VLR_DEPENDENTE) VLR_DEPENDENTE, SUM(VLR_TITULAR) + SUM(VLR_DEPENDENTE) VLR_GERAL');
      SQL.Add('  FROM (SELECT D.COD_BOL, SUM(VALOR_COPART_VTCN1(V.FMOD_VLR,V.FATOR_SEG,V.VALOR_1,V.VALOR_INSSADM_1,V.VALOR_GLOSADO,V.USA_PRC_DIF_HON,V.FATOR_SEG_HON,V.VALOR1_HON,V.VALOR_GLOSA_HON');
      SQL.Add('  ,V.FATOR_SEG_OUTROS,V.VALOR1_OUTROS,V.VALOR_GLOSA_OUTROS)) VLR_TITULAR, 0 VLR_DEPENDENTE');
      SQL.Add('          FROM IM_DBOL D');
      SQL.Add('          JOIN IM_SEG  S ON D.COD_SEG = S.COD_SEG');
      SQL.Add('  JOIN IM_VTCN1 V');
      SQL.Add('    ON SUBSTR(D.CODIGO,INSTR(D.CODIGO,'+#39+';'+#39+',1,2)+1,INSTR(SUBSTR(D.CODIGO,INSTR(D.CODIGO,'+#39+';'+#39+',1,2)+1),'+#39+';'+#39+')-1) = V.COD_SERV');
      SQL.Add('   AND SUBSTR(D.CODIGO,1,INSTR(D.CODIGO,'+#39+';'+#39+',1)-1) = V.TIPO_ENT');
      SQL.Add('   AND SUBSTR(D.CODIGO,INSTR(D.CODIGO,'+#39+';'+#39+',1)+1,INSTR(D.CODIGO,'+#39+';'+#39+',1,2)-3) = V.TIPO_SERV');
      SQL.Add('         WHERE D.COD_BOL = '+lblParam.Caption);
      SQL.Add('           AND S.TIPO = '+#39+'E'+#39);

        if lblCodigo.Caption <> 'X' then
          SQL.Add('           AND S.COD_SEG = '+#39+lblCodigo.Caption+#39);

      SQL.Add('         GROUP BY D.COD_BOL');
      SQL.Add('         UNION ALL');
      SQL.Add('        SELECT D.COD_BOL, 0 VLR_TITULAR, SUM(VALOR_COPART_VTCN1(V.FMOD_VLR,V.FATOR_SEG,V.VALOR_1,V.VALOR_INSSADM_1,V.VALOR_GLOSADO,V.USA_PRC_DIF_HON,V.FATOR_SEG_HON,V.VALOR1_HON');
      SQL.Add('      ,V.VALOR_GLOSA_HON,V.FATOR_SEG_OUTROS,V.VALOR1_OUTROS,V.VALOR_GLOSA_OUTROS)) VLR_DEPENDENTE');
      SQL.Add('          FROM IM_DBOL D');
      SQL.Add('          JOIN IM_SEG  S ON D.COD_SEG = S.COD_SEG');
      SQL.Add('  JOIN IM_VTCN1 V');
      SQL.Add('    ON SUBSTR(D.CODIGO,INSTR(D.CODIGO,'+#39+';'+#39+',1,2)+1,INSTR(SUBSTR(D.CODIGO,INSTR(D.CODIGO,'+#39+';'+#39+',1,2)+1),'+#39+';'+#39+')-1) = V.COD_SERV');
      SQL.Add('   AND SUBSTR(D.CODIGO,1,INSTR(D.CODIGO,'+#39+';'+#39+',1)-1) = V.TIPO_ENT');
      SQL.Add('   AND SUBSTR(D.CODIGO,INSTR(D.CODIGO,'+#39+';'+#39+',1)+1,INSTR(D.CODIGO,'+#39+';'+#39+',1,2)-3) = V.TIPO_SERV');
      SQL.Add('         WHERE D.COD_BOL = '+lblParam.Caption);
      SQL.Add('           AND S.TIPO = '+#39+'D'+#39);

        if lblCodigo.Caption <> 'X' then
          SQL.Add('           AND S.COD_SEG = '+#39+lblCodigo.Caption+#39);

      SQL.Add('         GROUP BY D.COD_BOL)');
      SQL.Add('  GROUP BY COD_BOL');
      Open;
    end;

    Geral      := qryExec.FieldByName('VLR_GERAL').AsFloat;
    Titular    := qryExec.FieldByName('VLR_TITULAR').AsFloat;
    Dependente := qryExec.FieldByName('VLR_DEPENDENTE').AsFloat;

    qrpRelat100.qrlblTotalGeral.Caption := 'Valor Total (Geral): ' + 'R$ ' + FormatFloat('#,0.00',Geral);
    qrpRelat100.qrlblTotalTitular.Caption := 'Valor Total (Titular): ' + 'R$ ' + FormatFloat('#,0.00',Titular);
    qrpRelat100.qrlblTotalDependente.Caption := 'Valor Total (Dependente): ' + 'R$ ' + FormatFloat('#,0.00',Dependente);

    JanelasCongeladas := DisableTaskWindows(frmAguarde.Handle); // Desabilita as janelas, exceto a janela informada

    try
      frmAguarde.Show;
      qrpRelat100.Prepare;
    finally
      EnableTaskWindows(JanelasCongeladas); // Reabibita Janelas Congeladas
      frmAguarde.Close;
    end;
    
    qrpRelat100.qrlblPagina.Caption := IntToStr(qrpRelat100.QRPrinter.PageCount);
    qrpRelat100.PreviewModal;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '116' then
  begin

    qrpRelat116.qrlblTitulo.Caption := 'Extrato do Boleto (' + lblParam.Caption + ')';

    with DM.qryRelat116Tit do
    begin
      Close;
      SQL.Clear;
      SQL.Add('SELECT '+#39+' TITULAR: '+#39+' || T.COD_SEG || '+#39+' - '+#39+' || T.NOME TITULAR, TO_CHAR(B.COD_BOL) BOLETO, T.COD_SEG COD_TITULAR');
      SQL.Add('  FROM IM_BOLAB B');
      SQL.Add('  JOIN IM_DBOL  D ON B.COD_BOL = D.COD_BOL');
      SQL.Add('  JOIN IM_SEG   S ON D.COD_SEG = S.COD_SEG');
      SQL.Add('  JOIN IM_SEG   T ON S.COD_FAM = T.COD_SEG');
      SQL.Add('  JOIN IM_VTCN1 V');
      SQL.Add('    ON SUBSTR(D.CODIGO,INSTR(D.CODIGO,'+#39+';'+#39+',1,2)+1,INSTR(SUBSTR(D.CODIGO,INSTR(D.CODIGO,'+#39+';'+#39+',1,2)+1),'+#39+';'+#39+')-1) = V.COD_SERV');
      SQL.Add('   AND SUBSTR(D.CODIGO,1,INSTR(D.CODIGO,'+#39+';'+#39+',1)-1) = V.TIPO_ENT');
      SQL.Add('   AND SUBSTR(D.CODIGO,INSTR(D.CODIGO,'+#39+';'+#39+',1)+1,INSTR(D.CODIGO,'+#39+';'+#39+',1,2)-3) = V.TIPO_SERV');

      SQL.Add(' WHERE B.COD_BOL = '+lblParam.Caption);

      if lblCodigo.Caption <> 'X' then
        SQL.Add('AND T.COD_SEG = '+#39+lblCodigo.Caption+#39);

      SQL.Add('   AND T.TIPO = '+#39+'E'+#39);
      SQL.Add(' GROUP BY '+#39+' TITULAR: '+#39+' || T.COD_SEG || '+#39+' - '+#39+' || T.NOME, TO_CHAR(B.COD_BOL) , T.COD_SEG');
      Open;
    end;

    DM.qryRelat116Detalhe.Open;

    with qryExec do
    begin
      Close;
      SQL.Clear;
      SQL.Add('SELECT COD_BOL, SUM(VLR_TITULAR) VLR_TITULAR, SUM(VLR_DEPENDENTE) VLR_DEPENDENTE, SUM(VLR_TITULAR) + SUM(VLR_DEPENDENTE) VLR_GERAL');
      SQL.Add('  FROM (SELECT D.COD_BOL, SUM(D.VALOR) VLR_TITULAR, 0 VLR_DEPENDENTE');
      SQL.Add('          FROM IM_DBOL D');
      SQL.Add('          JOIN IM_SEG  S ON D.COD_SEG = S.COD_SEG');
      SQL.Add('  JOIN IM_VTCN1 V');
      SQL.Add('    ON SUBSTR(D.CODIGO,INSTR(D.CODIGO,'+#39+';'+#39+',1,2)+1,INSTR(SUBSTR(D.CODIGO,INSTR(D.CODIGO,'+#39+';'+#39+',1,2)+1),'+#39+';'+#39+')-1) = V.COD_SERV');
      SQL.Add('   AND SUBSTR(D.CODIGO,1,INSTR(D.CODIGO,'+#39+';'+#39+',1)-1) = V.TIPO_ENT');
      SQL.Add('   AND SUBSTR(D.CODIGO,INSTR(D.CODIGO,'+#39+';'+#39+',1)+1,INSTR(D.CODIGO,'+#39+';'+#39+',1,2)-3) = V.TIPO_SERV');
      SQL.Add('         WHERE D.COD_BOL = '+lblParam.Caption);

        if lblCodigo.Caption <> 'X' then
          SQL.Add('           AND S.COD_FAM = '+#39+lblCodigo.Caption+#39);

      SQL.Add('           AND S.TIPO = '+#39+'E'+#39);
      SQL.Add('         GROUP BY D.COD_BOL');
      SQL.Add('         UNION ALL');
      SQL.Add('        SELECT D.COD_BOL, 0 VLR_TITULAR, SUM(D.VALOR) VLR_DEPENDENTE');
      SQL.Add('          FROM IM_DBOL D');
      SQL.Add('          JOIN IM_SEG  S ON D.COD_SEG = S.COD_SEG');
      SQL.Add('  JOIN IM_VTCN1 V');
      SQL.Add('    ON SUBSTR(D.CODIGO,INSTR(D.CODIGO,'+#39+';'+#39+',1,2)+1,INSTR(SUBSTR(D.CODIGO,INSTR(D.CODIGO,'+#39+';'+#39+',1,2)+1),'+#39+';'+#39+')-1) = V.COD_SERV');
      SQL.Add('   AND SUBSTR(D.CODIGO,1,INSTR(D.CODIGO,'+#39+';'+#39+',1)-1) = V.TIPO_ENT');
      SQL.Add('   AND SUBSTR(D.CODIGO,INSTR(D.CODIGO,'+#39+';'+#39+',1)+1,INSTR(D.CODIGO,'+#39+';'+#39+',1,2)-3) = V.TIPO_SERV');
      SQL.Add('         WHERE D.COD_BOL = '+lblParam.Caption);

        if lblCodigo.Caption <> 'X' then
          SQL.Add('           AND S.COD_FAM = '+#39+lblCodigo.Caption+#39);

      SQL.Add('           AND S.TIPO = '+#39+'D'+#39);
      SQL.Add('         GROUP BY D.COD_BOL)');
      SQL.Add('  GROUP BY COD_BOL');
      Open;
    end;

    Geral      := qryExec.FieldByName('VLR_GERAL').AsFloat;
    Titular    := qryExec.FieldByName('VLR_TITULAR').AsFloat;
    Dependente := qryExec.FieldByName('VLR_DEPENDENTE').AsFloat;

    qrpRelat116.qrlblTotalGeral.Caption := 'Valor Total (Geral): ' + 'R$ ' + FormatFloat('#,0.00',Geral);
    qrpRelat116.qrlblTotalTitular.Caption := 'Valor Total (Titular): ' + 'R$ ' + FormatFloat('#,0.00',Titular);
    qrpRelat116.qrlblTotalDependente.Caption := 'Valor Total (Dependente): ' + 'R$ ' + FormatFloat('#,0.00',Dependente);

    JanelasCongeladas := DisableTaskWindows(frmAguarde.Handle); // Desabilita as janelas, exceto a janela informada

    try
      frmAguarde.Show;
      qrpRelat116.Prepare;
    finally
      EnableTaskWindows(JanelasCongeladas); // Reabibita Janelas Congeladas
      frmAguarde.Close;
    end;
    
    qrpRelat116.qrlblPagina.Caption := IntToStr(qrpRelat116.QRPrinter.PageCount);
    qrpRelat116.PreviewModal;
  end;

end;

procedure TfrmRelatorio4.qryDadosAfterOpen(DataSet: TDataSet);
begin
frmPrincipal.DimensionarGrid(dbgDados,Self);

  if (frmPrincipal.lblCodRelatorio.Caption = '54') or (frmPrincipal.lblCodRelatorio.Caption = '97') then
  begin
    TCurrencyField(qryDados.FieldByName('RECEITA')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('DESPESA')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('COD_SEG')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_VEND')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_INCL')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('DT_CANC')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('QTD')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('UTILIZACAO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('RECEITA')).Alignment := taRightJustify;
    TCurrencyField(qryDados.FieldByName('DESPESA')).Alignment := taRightJustify;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '100' then
  begin
    TCurrencyField(qryDados.FieldByName('COD_FAM')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_BOL')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('TIPO')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('COD_SEG')).Alignment := taCenter;
    TCurrencyField(qryDados.FieldByName('VALOR')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryDados.FieldByName('REALIZACAO')).Alignment := taCenter;    
  end;

end;

procedure TfrmRelatorio4.edtCodigoKeyPress(Sender: TObject; var Key: Char);
begin
    if not (key in ['0'..'9',Chr(8), Chr(3), Chr(22), Chr(24), Chr(44), Chr(45)]) then
      key := #0;
end;

procedure TfrmRelatorio4.edtCodigoExit(Sender: TObject);
begin
  if (frmPrincipal.lblCodRelatorio.Caption = '100') or (frmPrincipal.lblCodRelatorio.Caption = '116') then
  begin

    if edtCodigo.Text = '' then
    begin
      lblCodigo.Caption := 'X';
      edtNome.Clear;
      edtCodigo.Clear;
      lblCodigo.Caption := 'X';
    end else
    begin
      with qryExec do
      begin
        Close;
        SQL.Clear;
        SQL.Add('select s.cod_seg, s.nome');
        SQL.Add('  from im_seg s');
        SQL.Add(' where s.cod_seg = '+#39+edtCodigo.Text+#39);

        if frmPrincipal.lblCodRelatorio.Caption = '116' then
          SQL.Add('   and s.tipo = '+#39+'E'+#39);

        Open;
      end;

      if qryExec.RecordCount > 0 then
      begin
        edtNome.Text := qryExec.FieldByName('NOME').AsString;
        lblCodigo.Caption := qryExec.FieldByName('COD_SEG').AsString;
      end else
      begin
        ShowMessage('Titular n�o encontrado');
        edtNome.Clear;
        edtCodigo.Clear;
        lblCodigo.Caption := 'X';
      end;
    end;
  end;
end;

end.
