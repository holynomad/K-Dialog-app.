unit DlgPrint;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, 
  Qrctrls, QuickRpt, ExtCtrls, Variants, Vcl.Imaging.jpeg;

type
  TFtpPrint = class(TForm)
    qr_KDial: TQuickRep;
    DetailBand1: TQRBand;
    qrimg_KDial: TQRImage;
    qrlb_DocTitle: TQRLabel;
    qrlb_HqAmt: TQRLabel;
    qrlb_Period: TQRLabel;
    qrlb_GrAmt: TQRLabel;
    qrlb_DocName: TQRLabel;
    qrlb_BarPatno: TQRLabel;
    qrlb_AaAmt: TQRLabel;
    qrlb_TotAmt1: TQRLabel;
    qrlb_Company: TQRLabel;
    qrlb_AsAmt: TQRLabel;
    qrlb_MedDate_Day: TQRLabel;
    qrlb_MedDate_Hour: TQRLabel;
    qrlb_MedDate_Min: TQRLabel;
    qrlb_TotAmt2: TQRLabel;
    qrlb_ImgCode: TQRLabel;
    qrlb_DutyUser: TQRLabel;
    qrlb_Manager: TQRLabel;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure qr_KDialBeforePrint(Sender: TCustomQuickRep;
      var PrintReport: Boolean);
  private
    { Private declarations }

    procedure GetFileImage;

  public
    { Public declarations }
  end;

var
  FtpPrint: TFtpPrint;


implementation

uses
  CComFunc,
   TuxCom,
   VarCom,
   MainDialog;

{$R *.DFM}

procedure TFtpPrint.FormClose(Sender: TObject; var Action: TCloseAction);
begin
   Action := caFree;
end;

procedure TFtpPrint.FormDestroy(Sender: TObject);
begin
   FtpPrint := nil;
end;


procedure TFtpPrint.GetFileImage;
var
   ServerFileName, ClientFileName: String;
   FTP_SVRIP, FTP_USERID, FTP_PASSWD, FTP_HOSTNAME, FTP_DIR : String;
begin
   //-------------------------------------------------------
   // ���� �� Ŭ���̾�Ʈ Path
   //-------------------------------------------------------
   ServerFileName := 'KDIAL_' + qrlb_ImgCode.Caption + '.JPG';

   // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
   ClientFileName := G_HOMEDIR + 'TEMP\SPOOL\'+ ServerFileName;

   // test
   //showmessage(ClientFileName);

   //-------------------------------------------------------
   // ���� ���� Ȯ��
   //-------------------------------------------------------
   if Not GetBinUploadInfo(FTP_SVRIP,
                           FTP_USERID,
                           FTP_PASSWD,
                           FTP_HOSTNAME,
                           FTP_DIR) then
   begin
      ShowMessage('�ٿ�ε尡 �����Ͽ� ������ �� �����ϴ�.');
      Exit;
   end;

   //-------------------------------------------------------
   // FTP ���� ����
   //-------------------------------------------------------
   {
   FTP_USERID := '';
   FTP_PASSWD := '';
   }
   FTP_DIR    := '/ftpspool/KDIALIMG/';

   //-------------------------------------------------------
   // FTP ����
   //-------------------------------------------------------
   if Not GetBINFTP(FTP_SVRIP,
                    FTP_USERID,
                    FTP_PASSWD,
                    FTP_DIR + ServerFileName,
                    ClientFileName,
                    False) then
   begin
      Showmessage('�̹��� �ٿ�ε忡 �����߽��ϴ�. �ٽ��ѹ� �õ��� �ֽʽÿ�');
      Exit;
   end;
end;

procedure TFtpPrint.qr_KDialBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
var
   isPicture : String;
begin

   //-------------------------------------------------------
   // �̹��� �ڵ�
   //-------------------------------------------------------
   qrlb_ImgCode.Top      := 999;
   qrlb_ImgCode.Left     := 999;

   {

   //-------------------------------------------------------
   // ȯ�� �⺻���� ��ȸ
   //-------------------------------------------------------
   GetBasPatNo(qrlb_Patno.Caption,
               Patname,
               Resno1,
               Resno2,
               Birtdate,
               Sex,
               Telno1,
               Telno2,
               Telno3,
               Email,
               Zipcd,
               Address,
               Bldtype,
               Scanyn,
               Resno3,
               Zpaddress,
               Address2);

   //-------------------------------------------------------
   // ȯ�� ������� ǥ��
   //-------------------------------------------------------
   qrlb_Birtdate.Caption := '(' + CopyByte(Birtdate, 1, 4) + '.' +
                                  CopyByte(Birtdate, 5, 2) + '.' +
                                  CopyByte(Birtdate, 7, 2) + ')';

   //-------------------------------------------------------
   // ȯ�� ������� ��ġ
   //-------------------------------------------------------
   qrlb_Birtdate.Left := qrlb_PatName.Left + qrlb_PatName.Width + 1;


   //-------------------------------------------------------
   // ������ ������ ���� �ֱ� �������� ��ȸ
   //-------------------------------------------------------
   if GetMinMeddate('W',
                    qrlb_PatNo.Caption,                                                         // ȯ�ڹ�ȣ
                    InvertSDate(qrlb_MedDate.Caption),                                          // �� �������� (yyyymmdd)
                    CopyByte(Trim(qrlb_DeptCd.Caption), 1, PosByte('0', Trim(qrlb_DeptCd.Caption))-1),  // �μ��ڵ� (MedDept)
                    sMinMedDate,
                    sMeddrId,
                    sMeddrNm) > 0 Then
   begin
      qrlb_MedDate_Month.Caption := CopyByte(sMinMedDate, 6,  2);
      qrlb_MedDate_Day.Caption   := CopyByte(sMinMedDate, 9,  2);
      qrlb_MedDate_Hour.Caption  := CopyByte(sMinMedDate, 12, 2);
      qrlb_MedDate_Min.Caption   := CopyByte(sMinMedDate, 15, 2);
   end
   else
   begin
      qrlb_MedDate_Month.Caption := '';
      qrlb_MedDate_Day.Caption   := '';
      qrlb_MedDate_Hour.Caption  := '';
      qrlb_MedDate_Min.Caption   := '';
   end;

   }


   //-------------------------------------------------------
   // ��� FTP �̹��� �ٿ�
   //-------------------------------------------------------
   GetFileImage;

   //-------------------------------------------------------
   // �̹��� Path ����
   //    - C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
   //-------------------------------------------------------
   isPicture := G_HOMEDIR + 'TEMP\SPOOL\KDIAL_' + qrlb_ImgCode.Caption + '.JPG';

   //-------------------------------------------------------
   // �̹��� Loading
   //-------------------------------------------------------
   qrimg_KDial.Picture.LoadFromFile(isPicture);

end;

end.
