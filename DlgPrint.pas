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
   // 서버 및 클라이언트 Path
   //-------------------------------------------------------
   ServerFileName := 'KDIAL_' + qrlb_ImgCode.Caption + '.JPG';

   // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
   ClientFileName := G_HOMEDIR + 'TEMP\SPOOL\'+ ServerFileName;

   // test
   //showmessage(ClientFileName);

   //-------------------------------------------------------
   // 계정 정보 확인
   //-------------------------------------------------------
   if Not GetBinUploadInfo(FTP_SVRIP,
                           FTP_USERID,
                           FTP_PASSWD,
                           FTP_HOSTNAME,
                           FTP_DIR) then
   begin
      ShowMessage('다운로드가 실패하여 실행할 수 없습니다.');
      Exit;
   end;

   //-------------------------------------------------------
   // FTP 계정 정보
   //-------------------------------------------------------
   {
   FTP_USERID := '';
   FTP_PASSWD := '';
   }
   FTP_DIR    := '/ftpspool/KDIALIMG/';

   //-------------------------------------------------------
   // FTP 접속
   //-------------------------------------------------------
   if Not GetBINFTP(FTP_SVRIP,
                    FTP_USERID,
                    FTP_PASSWD,
                    FTP_DIR + ServerFileName,
                    ClientFileName,
                    False) then
   begin
      Showmessage('이미지 다운로드에 실패했습니다. 다시한번 시도해 주십시오');
      Exit;
   end;
end;

procedure TFtpPrint.qr_KDialBeforePrint(Sender: TCustomQuickRep;
  var PrintReport: Boolean);
var
   isPicture : String;
begin

   //-------------------------------------------------------
   // 이미지 코드
   //-------------------------------------------------------
   qrlb_ImgCode.Top      := 999;
   qrlb_ImgCode.Left     := 999;

   {

   //-------------------------------------------------------
   // 환자 기본정보 조회
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
   // 환자 생년월일 표기
   //-------------------------------------------------------
   qrlb_Birtdate.Caption := '(' + CopyByte(Birtdate, 1, 4) + '.' +
                                  CopyByte(Birtdate, 5, 2) + '.' +
                                  CopyByte(Birtdate, 7, 2) + ')';

   //-------------------------------------------------------
   // 환자 생년월일 위치
   //-------------------------------------------------------
   qrlb_Birtdate.Left := qrlb_PatName.Left + qrlb_PatName.Width + 1;


   //-------------------------------------------------------
   // 진료일 이후의 가장 최근 진료일자 조회
   //-------------------------------------------------------
   if GetMinMeddate('W',
                    qrlb_PatNo.Caption,                                                         // 환자번호
                    InvertSDate(qrlb_MedDate.Caption),                                          // 현 진료일자 (yyyymmdd)
                    CopyByte(Trim(qrlb_DeptCd.Caption), 1, PosByte('0', Trim(qrlb_DeptCd.Caption))-1),  // 부서코드 (MedDept)
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
   // 출력 FTP 이미지 다운
   //-------------------------------------------------------
   GetFileImage;

   //-------------------------------------------------------
   // 이미지 Path 세팅
   //    - C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
   //-------------------------------------------------------
   isPicture := G_HOMEDIR + 'TEMP\SPOOL\KDIAL_' + qrlb_ImgCode.Caption + '.JPG';

   //-------------------------------------------------------
   // 이미지 Loading
   //-------------------------------------------------------
   qrimg_KDial.Picture.LoadFromFile(isPicture);

end;

end.
