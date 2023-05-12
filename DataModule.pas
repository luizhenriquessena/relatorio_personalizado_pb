unit DataModule;

interface

uses
  SysUtils, Classes, DB, ADODB, Dialogs, Registry, Windows;

type
  TDM = class(TDataModule)
    OpenDialog1: TOpenDialog;
    conTI: TADOConnection;
    SaveDialog1: TSaveDialog;
    conLocaWeb1: TADOConnection;
    qryRelat60: TADOQuery;
    qryRelat59: TADOQuery;
    qryRelat60NUM_DOC: TWideStringField;
    qryRelat60SENHA_LIB: TWideStringField;
    qryRelat60DATA: TWideStringField;
    qryRelat60NOME: TWideStringField;
    qryRelat60COD_AMB: TWideStringField;
    qryRelat60COD_SERV: TBCDField;
    qryRelat60VALOR: TBCDField;
    qryRelat60VALOR_GLOSADO: TBCDField;
    qryRelat60OBS_GLOSA: TWideStringField;
    qryRelat60COD_TISS_GLOSA: TIntegerField;
    qryRelat60DESCRICAO: TWideStringField;
    qryRelat60FATURA: TWideStringField;
    qryRelat59LOTE: TWideStringField;
    qryRelat59NU_PROTOCOLO: TBCDField;
    qryRelat59NUM_DOC: TWideStringField;
    qryRelat59SENHA_LIB: TWideStringField;
    qryRelat59DATA: TWideStringField;
    qryRelat59NOME: TWideStringField;
    qryRelat59COD_AMB: TWideStringField;
    qryRelat59COD_SERV: TBCDField;
    qryRelat59VALOR: TBCDField;
    qryRelat59VALOR_GLOSADO: TBCDField;
    qryRelat59OBS_GLOSA: TWideStringField;
    qryRelat59COD_TISS_GLOSA: TIntegerField;
    qryRelat59DESCRICAO: TWideStringField;
    qryRelat59FATURA: TWideStringField;
    qryRelat65: TADOQuery;
    qryRelat65ANO_MES: TWideStringField;
    qryRelat65VIDAS: TBCDField;
    qryRelat65RECEITA: TBCDField;
    qryRelat100Tit: TADOQuery;
    qryRelat100Detalhe: TADOQuery;
    qryRelat100TitTITULAR: TWideStringField;
    qryRelat100DetalheCOD_FAM: TWideStringField;
    qryRelat100DetalheCOD_BOL: TBCDField;
    qryRelat100DetalheTIPO: TWideStringField;
    qryRelat100DetalheCOD_SEG: TWideStringField;
    qryRelat100DetalheNOME: TWideStringField;
    qryRelat100DetalheDESCRICAO: TWideStringField;
    qryRelat100DetalheVALOR: TBCDField;
    dsRelat100: TDataSource;
    qryRelat100TitCOD_TITULAR: TWideStringField;
    qryRelat100TitBOLETO: TWideStringField;
    qryRelat100DetalheREALIZACAO: TDateTimeField;
    qryRelat116Detalhe: TADOQuery;
    WideStringField4: TWideStringField;
    BCDField1: TBCDField;
    WideStringField5: TWideStringField;
    WideStringField6: TWideStringField;
    WideStringField7: TWideStringField;
    BCDField2: TBCDField;
    dsRelat116: TDataSource;
    qryRelat116Tit: TADOQuery;
    WideStringField1: TWideStringField;
    WideStringField2: TWideStringField;
    WideStringField3: TWideStringField;
    qryRelat100TitCOD_SEGX: TWideStringField;
    qryRelat65SINISTRALIDADE: TBCDField;
    qryRelat65COMISSAO: TBCDField;
    qryRelat100DetalhePRESTADOR: TWideStringField;
    qryRelat100DetalheVALOR_C: TWideStringField;
    procedure qryRelat59AfterOpen(DataSet: TDataSet);
    procedure qryRelat65AfterOpen(DataSet: TDataSet);
  private
    { Private declarations }
  public
    { Public declarations }
    function RetirarZerosEsquerda(text: String): String;
    function StrToFloatDulio(vlr: String): Currency;
    function NumeroExtenso(const Valor: double; Moeda: Boolean = False): string;
    function NomeComputador: String;
  end;

var
  DM: TDM;

implementation

Uses Relatorio13, Principal, Relatorio4;
{$R *.dfm}

{ TDM }
function TDM.NomeComputador: String;
var   I: DWord;
begin
  I := MAX_COMPUTERNAME_LENGTH + 1;
  SetLength(Result, I);
  Windows.GetComputerName(PChar(Result), I);
  Result := string(PChar(Result));
end;

function TDM.NumeroExtenso(const Valor: double; Moeda: Boolean): string;
const
  Centenas: array[1..9] of string[12] = ('CEM', 'DUZENTOS', 'TREZENTOS',
    'QUATROCENTOS', 'QUINHENTOS', 'SEISCENTOS', 'SETECENTOS',
    'OITOCENTOS', 'NOVECENTOS');
  Dezenas: array[2..9] of string[10] = ('VINTE', 'TRINTA', 'QUARENTA',
    'CINQUENTA', 'SESSENTA', 'SETENTA', 'OITENTA', 'NOVENTA');
  Dez: array[0..9] of string[10] = ('DEZ', 'ONZE', 'DOZE', 'TREZE', 'QUATORZE',
    'QUINZE', 'DEZESSEIS', 'DEZESSETE', 'DEZOITO', 'DEZENOVE');
  Unidades: array[1..9] of string[10] = ('UM', 'DOIS', 'TRES', 'QUATRO', 'CINCO',
    'SEIS', 'SETE', 'OITO', 'NOVE');
  Zero = 'ZERO';

  function Ext3(Parte: string): string;
  var
    Base: string;
    digito: integer;
  begin
    Base := '';
    digito := StrToInt(Parte[1]);
    if digito = 0 then
      Base := ''
    else
      Base := Centenas[digito];
    if (digito = 1) and (Parte > '100') then
      Base := 'CENTO';
    Digito := StrToInt(Parte[2]);
    if digito = 1 then
    begin
      Digito := StrToInt(Parte[3]);
      if Base <> '' then
        Base := Base + ' E ';
      Base := Base + Dez[Digito];
    end
    else
    begin
      if (Base <> '') then
        Base := Base + ' E ';
      if Digito > 1 then
        Base := Base + Dezenas[digito];
      Digito := StrToInt(Parte[3]);
      if Digito > 0 then
      begin
        if Base <> '' then
          Base := Base + ' E ';
        Base := Base + Unidades[Digito];
      end;
    end;
    Result := Base;
  end;

var
  ComoTexto: string;
  Parte: string;
  MoedaSingular: String;
  MoedaPlural: String;
  CentSingular: String;
  CentPlural: String;

begin
  if Moeda then
  begin
    MoedaSingular := 'REAL';
    MoedaPlural := 'REAIS';
    CentSingular := 'CENTAVO';
    CentPlural := 'CENTAVOS';
  end
  else
  begin
    MoedaSingular := '';
    MoedaPlural := '';
    CentSingular := '';
    CentPlural := '';
  end;

  Result := '';
  ComoTexto := FloatToStrF(Abs(Valor), ffFixed, 18, 2);
  // Acrescenta zeros a esquerda ate 12 digitos
  while length(ComoTexto) < 15  do
    ComoTexto := '0'+ ComoTexto;
  // Calcula os bilhões
  Parte := Ext3(copy(ComoTexto, 1, 3));
  if StrToInt(copy(ComoTexto, 1, 3)) = 1 then
    Parte := Parte + ' BILHAO'
  else
    if Parte <> '' then
      Parte := Parte + ' BILHOES';
  Result := Parte;
  // Calcula os milhões
  Parte := Ext3(copy(ComoTexto, 4, 3));
  if Parte <> '' then
  begin
    if Result <> '' then
      Result := Result + ', ';
    if StrToInt(copy(ComoTexto, 4, 3)) = 1 then
      Parte := Parte + ' MILHAO'
    else
      Parte := Parte + ' MILHOES';
    Result := Result + Parte;
  end;
  // Calcula os milhares
  Parte := Ext3(copy(ComoTexto, 7, 3));
  if Parte <> '' then
  begin
    if Result <> '' then
      Result := Result + ', ';
    Parte := Parte + ' MIL';
    Result := Result + Parte;
  end;
  // Calcula as unidades
  Parte := Ext3(copy(ComoTexto, 10, 3));
  if Parte <> '' then
  begin
    if Result <> '' then
      if Frac(Valor) = 0 then
        Result := Result + ' E '
      else
        Result := Result + ', ';
    Result := Result + Parte;
  end;
  // Acrescenta o texto da moeda
  if Int(Valor) = 1 then
    Parte := ' ' + MoedaSingular
  else
    Parte := ' ' + MoedaPlural;
  if copy(ComoTexto, 7, 6) = '000000' then
    Parte := 'DE ' + MoedaPlural;
  Result := Result + Parte;
  // Se o valor for zero, limpa o resultado
  if int(Valor) = 0 then
    Result := '';
  //Calcula os centavos
  Parte := Ext3('0' + copy(ComoTexto, 14, 2));
  if Parte <> '' then
  begin
    if Result <> '' then
      Result := Result + ' E ';
    if Parte = Unidades[1] then
      Parte := Parte + ' ' + CentSingular
    else
      Parte := Parte + ' ' + CentPlural;
    Result := Result + Parte;
  end;
  // Se o valor for zero, assume a constante ZERO
  if Valor = 0 then
    Result := Zero;
end;

function TDM.RetirarZerosEsquerda(text: String): String;
var i, t: Integer;
begin
  t := length(text);
  for i:=1 to t do
  begin
    if not(text[i] = '0') then
      Break;
  end;
  Result := copy(text, i, t);
end;

function TDM.StrToFloatDulio(vlr: String): Currency;
var aux: String;   i: Integer;
begin
  aux:= '';
  for i:=1 to Length(vlr) do
    if vlr[i]<>'.' then
      aux:= aux+vlr[i];
  Result:= StrToFloat(aux);
end;

procedure TDM.qryRelat59AfterOpen(DataSet: TDataSet);
begin
  if frmPrincipal.lblCodRelatorio.Caption = '59' then
  begin
    TCurrencyField(qryRelat59.FieldByName('LOTE')).Alignment := taCenter;
    TCurrencyField(qryRelat59.FieldByName('NU_PROTOCOLO')).Alignment := taCenter;
    TCurrencyField(qryRelat59.FieldByName('NUM_DOC')).Alignment := taCenter;
    TCurrencyField(qryRelat59.FieldByName('SENHA_LIB')).Alignment := taCenter;
    TCurrencyField(qryRelat59.FieldByName('DATA')).Alignment := taCenter;
    TCurrencyField(qryRelat59.FieldByName('COD_AMB')).Alignment := taCenter;
    TCurrencyField(qryRelat59.FieldByName('COD_SERV')).Alignment := taCenter;
    TCurrencyField(qryRelat59.FieldByName('VALOR')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryRelat59.FieldByName('VALOR_GLOSADO')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryRelat59.FieldByName('COD_TISS_GLOSA')).Alignment := taCenter;
    TCurrencyField(qryRelat59.FieldByName('FATURA')).Alignment := taCenter;
  end;
end;

procedure TDM.qryRelat65AfterOpen(DataSet: TDataSet);
begin
  if frmPrincipal.lblCodRelatorio.Caption = '65' then
  begin
    TCurrencyField(qryRelat65.FieldByName('ANO_MES')).Alignment := taCenter;
    TCurrencyField(qryRelat65.FieldByName('VIDAS')).DisplayFormat := '#,##0';
    TCurrencyField(qryRelat65.FieldByName('SINISTRALIDADE')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryRelat65.FieldByName('COMISSAO')).DisplayFormat := 'R$ #,##0.00';
    TCurrencyField(qryRelat65.FieldByName('RECEITA')).DisplayFormat := 'R$ #,##0.00';
  end;
end;

end.
