unit Relat100;

interface

uses Windows, SysUtils, Messages, Classes, Graphics, Controls,
  StdCtrls, ExtCtrls, Forms, QuickRpt, QRCtrls;

type
  TqrpRelat100 = class(TQuickRep)
    QRBand1: TQRBand;
    qrlblTitulo: TQRLabel;
    QRBand3: TQRBand;
    QRShape2: TQRShape;
    QRShape3: TQRShape;
    QRDBText1: TQRDBText;
    QRLabel6: TQRLabel;
    QRLabel7: TQRLabel;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    QRLabel10: TQRLabel;
    QRSubDetail1: TQRSubDetail;
    QRDBText2: TQRDBText;
    QRShape1: TQRShape;
    QRDBText3: TQRDBText;
    QRDBText4: TQRDBText;
    QRDBText5: TQRDBText;
    QRDBText6: TQRDBText;
    qrlblPagina: TQRLabel;
    QRSysData1: TQRSysData;
    QRLabel22: TQRLabel;
    QRBand2: TQRBand;
    QRShape8: TQRShape;
    QRLabel14: TQRLabel;
    QRLabel18: TQRLabel;
    QRLabel20: TQRLabel;
    QRBand4: TQRBand;
    qrlblTotalTitular: TQRLabel;
    qrlblTotalDependente: TQRLabel;
    qrlblTotalGeral: TQRLabel;
    QRShape4: TQRShape;
    QRLabel1: TQRLabel;
    QRDBText7: TQRDBText;
    QRLabel2: TQRLabel;
    QRDBText8: TQRDBText;
  private

  public

  end;

var
  qrpRelat100: TqrpRelat100;

implementation

{$R *.DFM}
uses DataModule;
end.
