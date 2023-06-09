unit Relatorio8;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, ADODB, Grids, DBGrids, Buttons, StdCtrls;

type
  TfrmRelatorio8 = class(TForm)
    Label1: TLabel;
    btnExportar: TSpeedButton;
    lblTotal: TLabel;
    lblRTotal: TLabel;
    btnGerar: TSpeedButton;
    dbgDados: TDBGrid;
    dsDados: TDataSource;
    qryDados: TADOQuery;
    SaveDialog1: TSaveDialog;
    qryExec: TADOQuery;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure btnGerarClick(Sender: TObject);
    procedure btnExportarClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmRelatorio8: TfrmRelatorio8;

implementation

{$R *.dfm}
uses Principal, DateUtils, DataModule;
procedure TfrmRelatorio8.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
qryDados.Close;
end;

procedure TfrmRelatorio8.FormShow(Sender: TObject);
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
end;

procedure TfrmRelatorio8.btnGerarClick(Sender: TObject);
begin
  with qryDados do
  begin
  Close;
  SQL.Clear;
  SQL.Add(frmPrincipal.memSQL.Text);
  Open;
  end;
  lblRTotal.Caption := IntToStr(qryDados.RecordCount);

  with qryExec do // INSERIR LOG DE VISUALIZA��O
  begin
  Close;
  SQL.Clear;
  SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'V'+#39+', SYSDATE)');
  ExecSQL;
  end;
end;

procedure TfrmRelatorio8.btnExportarClick(Sender: TObject);
var
F: TextFile;
Cod_Seg, Nome, Cod_Outro2, Tipo, Plano, Cpf, cLinha : String;
begin

 if qryDados.RecordCount > 0 then
 begin
  if frmPrincipal.lblCodRelatorio.Caption = '16' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'NOME');
      qryDados.First;
      while not qryDados.Eof do
        begin
          Cod_Seg               := qryDados.FIELDBYNAME('COD_SEG').AsString;
          Nome                  := qryDados.FIELDBYNAME('NOME').AsString;

          cLinha := '';
          cLinha := cLinha + Cod_seg +';';
          cLinha := cLinha + Nome +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '17' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'COD_OUTRO2' +';'+ 'NOME' +';'+ 'TIPO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          Cod_Seg               := qryDados.FIELDBYNAME('COD_SEG').AsString;
          Cod_Outro2            := qryDados.FIELDBYNAME('COD_OUTRO2').AsString;
          Nome                  := qryDados.FIELDBYNAME('NOME').AsString;
          Tipo                  := qryDados.FIELDBYNAME('TIPO').AsString;

          cLinha := '';
          cLinha := cLinha + Cod_seg +';';
          cLinha := cLinha + Cod_Outro2 +';';
          cLinha := cLinha + Nome +';';
          cLinha := cLinha + Tipo +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '66' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'COD_SEG' +';'+ 'COD_OUTRO2' +';'+ 'NOME' +';'+ 'TIPO');
      qryDados.First;
      while not qryDados.Eof do
        begin
          Cod_Seg               := qryDados.FIELDBYNAME('COD_SEG').AsString;
          Cod_Outro2            := qryDados.FIELDBYNAME('COD_OUTRO2').AsString;
          Nome                  := qryDados.FIELDBYNAME('NOME').AsString;
          Tipo                  := qryDados.FIELDBYNAME('TIPO').AsString;

          cLinha := '';
          cLinha := cLinha + Cod_seg +';';
          cLinha := cLinha + Cod_Outro2 +';';
          cLinha := cLinha + Nome +';';
          cLinha := cLinha + Tipo +';';

          Writeln(F,cLinha);
          qryDados.Next;
        end;
      CloseFile (F);
      ShowMessage('Arquivo Gerado com sucesso.');
    end;
  end;

  if frmPrincipal.lblCodRelatorio.Caption = '67' then
  begin
    if SaveDialog1.Execute then
    begin
      AssignFile(F, SaveDialog1.FileName);
      Rewrite (F);
      cLinha := '';
      Writeln (F,cLinha + 'PLANO' +';'+ 'CODIGO' +';'+ 'NOME' +';'+ 'CPF');
      qryDados.First;
      while not qryDados.Eof do
        begin
          Plano               := qryDados.FIELDBYNAME('PLANO').AsString;
          Cod_Seg             := qryDados.FIELDBYNAME('CODIGO').AsString;
          Nome                := qryDados.FIELDBYNAME('NOME').AsString;
          Cpf                 := qryDados.FIELDBYNAME('CPF').AsString;

          cLinha := '';
          cLinha := cLinha + Plano +';';
          cLinha := cLinha + Cod_Seg +';';
          cLinha := cLinha + Nome +';';
          cLinha := cLinha + Cpf +';';

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
  SQL.Add('INSERT INTO TI.RELATORIOS_LOG (ID_RELATORIO, USUARIO, ACAO, DATA) VALUES ('+frmPrincipal.lblCodRelatorio.Caption+','+#39+frmPrincipal.lbl_Usuario.Caption+#39+', '+#39+'E'+#39+', SYSDATE)');
  ExecSQL;
  end;

 end;
end;

end.
