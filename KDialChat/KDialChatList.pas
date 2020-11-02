{===============================================================================
   Program ID    : KDialChatList.pas
   Program Name  : 다이얼로그 Chat 리스트 관리(BPL)
   Program Desc. : Chat 대상 리스트(User) 관리

   Author        : Lee, Se-Ha
   Date          : 2015.03.27
===============================================================================}
unit KDialChatList;


interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, 
  Grids, BaseGrid, AdvGrid, StdCtrls, ExtCtrls, ScktComp, TFlatSpeedButtonUnit,
  TFlatEditUnit, TFlatGaugeUnit, Variants, AdvObj;

type
  TKDialChat = class(TForm)
    lb_Status: TLabel;
    asg_ChatList: TAdvStringGrid;
    tm_Chat: TTimer;
    ServerSocket: TServerSocket;
    fsbt_Network: TFlatSpeedButton;
    fsbt_DialBook: TFlatSpeedButton;
    fsbt_Search: TFlatSpeedButton;
    fed_Scan: TFlatEdit;
    fgau_Progress: TFlatGauge;
    fed_Statusbar: TFlatEdit;
    procedure FormShow(Sender: TObject);
    procedure tm_ChatTimer(Sender: TObject);
    procedure asg_ChatListDblClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure fsbt_NetworkClick(Sender: TObject);
    procedure fsbt_DialBookClick(Sender: TObject);
    procedure fsbt_SearchClick(Sender: TObject);
    procedure fed_ScanKeyPress(Sender: TObject; var Key: Char);
    procedure fed_ScanClick(Sender: TObject);
    procedure asg_ChatListClick(Sender: TObject);
  private
    { Private declarations }
    FsUserNm : String;
    FsSearchGubun : String;

    iUserRowId : Integer;

   //----------------------------------------------------------------
   // String 내의 특정 문자열 삭제
   //    - 소스출처 : MComFunc.pas
   //----------------------------------------------------------------
   function DeleteStr(OrigStr, DelStr : String) : String;


   //----------------------------------------------------------------
   // 한영키를 한글키로 변환
   //    - 소스출처 : MComFunc.pas
   //----------------------------------------------------------------
   procedure HanKeyChg(Handle1:THandle);


   //----------------------------------------------------------------
   // 병원별 심사팀 담당자 Info 조회
   //    - 2015.06.05 LSH
   //----------------------------------------------------------------
   function GetChaInfo(in_ChaId, in_LocateNm : String) : String;


  public
    { Public declarations }

  published
      property prop_UserNm  : String read FsUserNm     write FsUserNm;


      //----------------------------------------------------------------
      // Chat 리스트 조회 @ 2015.03.27 LSH
      //----------------------------------------------------------------
      procedure SelChatList(Sender : TObject; in_Gubun : String);

  end;


const
   // 다이얼 Chat Columns
   C_CL_PHOTO     = 0;
   C_CL_USERNM    = 1;
   C_CL_LOCATE    = 2;
   C_CL_USERID    = 3;
   C_CL_DUTYPART  = 4;
   C_CL_MOBILE    = 5;
   C_CL_EMAIL     = 6;
   C_CL_USERIP    = 7;
   C_CL_PHOTOFILE = 8;
   C_CL_HIDEFILE  = 9;
   C_CL_DUTYRMK   = 10;
   C_CL_CALLNO    = 11;
   C_CL_USERORGNM = 12;

   // MIS FTP 서버 주소
   C_MIS_FTP_IP   = '10.1.2.17';

   // K-Dial FTP 서버 주소
   C_KDIAL_FTP_IP = '10.1.3.77';


var
  KDialChat: TKDialChat;



implementation

uses
   VarCom,
   CComfunc,
   StringLib,
   CMsg,
   MsgCom,
   Uencrypt,
   TuxMsg,
   SysMsg,
   TuxCom,
   HisUtil,
   TpSvc,   // TpCall Agent 사용관련
//   Excel97,
   Imm,     // HIMC 사용관련추가 @ 2013.08.22 LSH
   ShellApi,// ShellExec 사용관련 추가 @ 2013.08.22 LSH
   Winsock, // Windows Sockets API Unit
   ComObj,
   Qrctrls,
   QuickRpt,
   IdFTPCommon;   // 델파이 XE 컨버젼관련 Id_FTP --> IndyFTP 변환 @ 2014.04.07 LSH


{$R *.DFM}



procedure TKDialChat.FormShow(Sender: TObject);
var
   TpGetSearch : TTpSvc;
   iRowCnt, i  : Integer;
   ServerFileName, ClientFileName: String;
   FTP_SVRIP, FTP_USERID, FTP_PASSWD, FTP_HOSTNAME, FTP_DIR : String;
   sLocalFile, sRemoteFile : String;
   DownFile     : Variant;
   DownCnt      : Integer;
begin
   //-----------------------------------------------------------------
   // 1. 조회
   //-----------------------------------------------------------------
   FsSearchGubun := 'MYBOOK'; // Network에서 Mybook으로 변경 @ 2019.12.03 LSH


   // 아래 조회부분을 캡슐화
   SelChatList(Sender, FsSearchGubun);


   {
   //-----------------------------------------------------------------
   // 2-1. Load Data by TpSvc
   //-----------------------------------------------------------------
   TpGetSearch := TTpSvc.Create;
   TpGetSearch.Init(Self);


   Screen.Cursor := crHourglass;


   try
      TpGetSearch.CountField  := 'S_CODE1';
      TpGetSearch.ShowMsgFlag := False;

      if TpGetSearch.GetSvc('MD_KUMCM_S1',
                          ['S_TYPE1'  , '4'
                         , 'S_TYPE2'  , ''
                         , 'S_TYPE3'  , ''
                         , 'S_TYPE4'  , ''
                         , 'S_TYPE5'  , ''
                         , 'S_TYPE6'  , ''
                         , 'S_TYPE7'  , ''
                         , 'S_TYPE8'  , ''
                         , 'S_TYPE9'  , ''
                         , 'S_TYPE10' , ''
                          ],
                          [
                           'S_CODE1'  , 'sLocate'
                         , 'S_CODE2'  , 'sUserId'
                         , 'S_CODE3'  , 'sUserNm'
                         , 'S_CODE4'  , 'sDutyPart'
                         , 'S_CODE5'  , 'sDutySpec'
                         , 'S_CODE6'  , 'sDutyRmk'
                         , 'S_CODE7'  , 'sMobile'
                         , 'S_CODE8'  , 'sEmail'
                         , 'S_CODE9'  , 'sUserIp'
                         , 'S_CODE10' , 'sDutyUser'
                         , 'S_CODE11' , 'sDutyPtnr'
                         , 'S_CODE12' , 'sCallNo'
                         , 'S_CODE13' , 'sDelDate'
                         , 'S_CODE14' , 'sDeptCd'
                         , 'S_CODE15' , 'sFlag'
                         , 'S_CODE16' , 'sEditIp'
                         , 'S_CODE17' , 'sEditDate'
                         , 'S_CODE18' , 'sDocList'
                         , 'S_CODE19' , 'sDocYear'
                         , 'S_CODE20' , 'sDocSeq'
                         , 'S_CODE21' , 'sDocTitle'
                         , 'S_CODE22' , 'sRegDate'
                         , 'S_CODE23' , 'sRegUser'
                         , 'S_CODE24' , 'sRelDept'
                         , 'S_CODE25' , 'sDocRmk'
                         , 'S_CODE26' , 'sDocLoc'
                         , 'S_STRING1', 'sBoardSeq'
                         , 'S_STRING2', 'sCateUp'
                         , 'S_STRING3', 'sCateDown'
                         , 'S_STRING4', 'sContext'
                         , 'S_STRING5', 'sAttachNm'
                         , 'S_STRING6', 'sHideFile'
                         , 'S_STRING7', 'sServerIp'
                         , 'S_STRING8', 'sHeadTail'
                         , 'S_STRING9', 'sHeadSeq'
                         , 'S_STRING10','sTailSeq'
                         , 'S_STRING11','sReplyCnt'
                         , 'S_STRING12','sLikeCnt'
                         , 'S_STRING14','sNickNm'
                         , 'S_STRING15','sLogYn'
                         , 'S_STRING16','sDeptNm'
                         , 'S_STRING17','sDeptSpec'
                         , 'S_STRING18','sDataBase'
                         , 'S_STRING19','sLinkCnt'
                         , 'S_STRING20','sLinkSeq'
                         , 'S_STRING21','sReqNo'
                         , 'S_STRING22','sClienSrc'
                         , 'S_STRING23','sServeSrc'
                         , 'S_STRING24','sCofmUser'
                         , 'S_STRING25','sReleasDt'
                         , 'S_STRING26','sTestDate'
                         , 'S_STRING27','sClient'
                         , 'S_STRING28','sServer'
                         , 'S_STRING29','sRelesUsr'
                          ]) then

         if TpGetSearch.RowCount < 0 then
         begin
            asg_ChatList.RowCount   := 2;

            ShowMessage(GetTxMsg);

            Exit;
         end
         else if TpGetSearch.RowCount = 0 then
         begin
            lb_UserCnt.Caption  := '▶ 조회내역이 존재하지 않습니다.';

            Exit;
         end;


      with asg_ChatList do
      begin
         // Maximum value of progress status
         //fpb_DataLoading.Max := TpGetSearch.RowCount * (asg_Master.ColCount);

         iRowCnt  := TpGetSearch.RowCount;
         RowCount := iRowCnt + FixedRows + 1;


         for i := 0 to iRowCnt - 1 do
         begin
            Cells[C_CL_PHOTO,    i+FixedRows] := '';
            Cells[C_CL_USERNM,   i+FixedRows] := TpGetSearch.GetOutputDataS('sUserNm'    , i) + #13#10 + TpGetSearch.GetOutputDataS('sLocate'    , i) + TpGetSearch.GetOutputDataS('sDutyPart'  , i) + #13#10 + TpGetSearch.GetOutputDataS('sDutyRmk'   , i);
            Cells[C_CL_LOCATE,   i+FixedRows] := TpGetSearch.GetOutputDataS('sLocate'    , i);
            Cells[C_CL_USERID,   i+FixedRows] := TpGetSearch.GetOutputDataS('sUserId'    , i);
            Cells[C_CL_DUTYPART, i+FixedRows] := TpGetSearch.GetOutputDataS('sDutyPart'  , i);
            Cells[C_CL_DUTYRMK,  i+FixedRows] := TpGetSearch.GetOutputDataS('sDutyRmk'   , i);
            Cells[C_CL_MOBILE,   i+FixedRows] := TpGetSearch.GetOutputDataS('sMobile'    , i);
            Cells[C_CL_EMAIL,    i+FixedRows] := TpGetSearch.GetOutputDataS('sEmail'     , i);
            Cells[C_CL_USERIP,   i+FixedRows] := TpGetSearch.GetOutputDataS('sUserIp'    , i);
            Cells[C_CL_PHOTOFILE,i+FixedRows] := DeleteStr(TpGetSearch.GetOutputDataS('sAttachNm'  , i), '/media/cq/photo/');
            Cells[C_CL_HIDEFILE, i+FixedRows] := TpGetSearch.GetOutputDataS('sHideFile'  , i);

            AutoSizeRow(i + FixedRows);



            //------------------------------------------------------------
            // 접속하려는 Server 사용가능 여부 (현재 Session On/Off) Check
            //------------------------------------------------------------
            if (TpGetSearch.GetOutputDataS('sLogYn', i) = 'Y') then
               Colors[C_CL_USERNM, i+FixedRows] := $0000CDA3
            else
               Colors[C_CL_USERNM, i+FixedRows] := clWhite;





            if (Cells[C_CL_PHOTOFILE, i+FixedRows] <> '') then
            begin
               // MIS 기본 이미지 정보인 경우.
               if Cells[C_CL_HIDEFILE, i+FixedRows] = '' then
               begin

                  // 담당자 프로필 이미지 파일명
                  ServerFileName := Cells[C_CL_PHOTOFILE, i+FixedRows];


                  // Local 파일이름 Set
                  ClientFileName := 'C:\Dev\TEMP\' + ServerFileName;



                  // Local에 해당 Image 존재유무 체크
                  if Not FileExists(ClientFileName) then
                  begin
                     // FTP 서버정보 가져오기
                     if Not GetBinUploadInfo(FTP_SVRIP,
                                             FTP_USERID,
                                             FTP_PASSWD,
                                             FTP_HOSTNAME,
                                             FTP_DIR) then
                     begin
                        ShowMessage('다운로드가 실패하여 실행할 수 없습니다.');
                        TUXFTP := nil;
                        Exit;
                     end;


                     // FTP 서버 IP Set
                     FTP_SVRIP        := C_MIS_FTP_IP;


                     // FTP 계정정보 Set
                     FTP_USERID := 'kumis';
                     FTP_PASSWD := 'kumis';
                     FTP_DIR    := '/kumis/app/mis/media/cq/photo/';



                     // Image 다운로드
                     if Not GetBINFTP(FTP_SVRIP,
                                      FTP_USERID,
                                      FTP_PASSWD,
                                      FTP_DIR + ServerFileName,
                                      ClientFileName,
                                      False) then
                     begin
                        Showmessage('사진 이미지 다운로드에 실패했습니다. 다시한번 시도해 주십시오');

                        TUXFTP := nil;

                        Exit;
                     end;



                     // 해당 Image Variant로 받음.
                     AppendVariant(DownFile, ClientFileName);



                     // 다운로드 횟수 체크
                     DownCnt := DownCnt + 1;

                  end;


                  // 수신한 Image 파일을 Grid에 표기
                  CreatePicture(C_CL_PHOTO,
                                i+FixedRows,
                                False,
                                ShrinkWithAspectRatio,
                                0,
                                haLeft,
                                vaTop).LoadFromFile(ClientFileName);

               end
               else
               // 다이얼로그 자체 이미지 등록정보 여부 확인
               begin

                  // 파일 업/다운로드를 위한 정보 조회
                  if not GetBinUploadInfo(FTP_SVRIP,
                                          FTP_USERID,
                                          FTP_PASSWD,
                                          FTP_HOSTNAME,
                                          FTP_DIR) then
                  begin
                     MessageDlg('파일 다운을 위한 서버정보 조회중, 오류가 발생했습니다.', Dialogs.mtError, [Dialogs.mbOK], 0);
                     exit;
                  end;


                  // 실제저장된 서버 IP
                  //sServerIp := C_KDIAL_FTP_IP;


                  // 실제 서버에 저장되어 있는 파일명 지정
                  if PosByte('/ftpspool/KDIALFILE/', Cells[C_CL_HIDEFILE, i+FixedRows]) > 0 then
                     sRemoteFile := Cells[C_CL_HIDEFILE, i+FixedRows]
                  else
                     sRemoteFile := '/ftpspool/KDIALFILE/' + Cells[C_CL_HIDEFILE, i+FixedRows];


                  sLocalFile  := 'C:\Dev\TEMP\SPOOL\' + Cells[C_CL_PHOTOFILE, i+FixedRows];



                  if (GetBINFTP(FTP_SVRIP, FTP_USERID, FTP_PASSWD, sRemoteFile, sLocalFile, False)) then
                  begin
                     //	정상적인 FTP 다운로드
                  end;

                  // 이미지만 Preview 설정
                  if (PosByte('.bmp', sLocalFile) > 0) or
                     (PosByte('.BMP', sLocalFile) > 0) or
                     (PosByte('.jpg', sLocalFile) > 0) or
                     (PosByte('.JPG', sLocalFile) > 0) or
                     (PosByte('.png', sLocalFile) > 0) or
                     (PosByte('.PNG', sLocalFile) > 0) or
                     (PosByte('.gif', sLocalFile) > 0) or
                     (PosByte('.GIF', sLocalFile) > 0) then

                     // 수신한 Image 파일을 Grid에 표기
                     CreatePicture(C_CL_PHOTO,
                                   i+FixedRows,
                                   False,
                                   ShrinkWithAspectRatio,
                                   0,
                                   haLeft,
                                   vaTop).LoadFromFile(sLocalFile);
               end;
            end
            else
            begin
               // 담당자 프로필 이미지 파일명
               if Cells[C_CL_USERID, i+FixedRows] = 'XXXXX' then
               begin
                  // 프로필 이미지 Remove
                  RemovePicture(C_CL_PHOTO, i+FixedRows);

                  // 이미지 등록 버튼 생성
                  //AddButton(C_CL_PHOTO, i+FixedRows, asg_ChatList.ColWidths[C_CL_PHOTO], asg_ChatList.RowHeights[i+FixedRows], '등록', haBeforeText, vaCenter);

                  Continue;
               end
               else
                  // 담당자 HRM 프로필 이미지 파일명
                  ServerFileName := Cells[C_CL_USERID, i+FixedRows] + '.jpg';

                  // Local 파일이름 Set
                  ClientFileName := 'C:\Dev\TEMP\' + ServerFileName;


                  // Local에 해당 Image 존재유무 체크
                  if Not FileExists(ClientFileName) then
                  begin
                     // FTP 서버정보 가져오기
                     if Not GetBinUploadInfo(FTP_SVRIP,
                                             FTP_USERID,
                                             FTP_PASSWD,
                                             FTP_HOSTNAME,
                                             FTP_DIR) then
                     begin
                        ShowMessage('다운로드가 실패하여 실행할 수 없습니다.');
                        TUXFTP := nil;
                        Exit;
                     end;


                     // FTP 서버 IP Set
                     FTP_SVRIP        := C_MIS_FTP_IP;

                     // FTP 계정정보 Set
                     FTP_USERID := 'kuhrm';
                     FTP_PASSWD := 'kuhrm';
                     FTP_DIR    := '/kuhrm/app/hrm/photo/';


                     // HRM Image 다운로드
                     if Not GetBINFTP(FTP_SVRIP,
                                      FTP_USERID,
                                      FTP_PASSWD,
                                      FTP_DIR + ServerFileName,
                                      ClientFileName,
                                      False) then
                     begin
                        Showmessage('이미지 다운로드에 실패했습니다. 다시한번 시도해 주십시오');

                        TUXFTP := nil;

                        Exit;
                     end;


                     // 해당 Image Variant로 받음.
                     AppendVariant(DownFile, ClientFileName);


                     // 다운로드 횟수 체크
                     DownCnt := DownCnt + 1;
                  end;


                  // 수신한 Image 파일을 Grid에 표기
                  CreatePicture(C_CL_PHOTO,
                                i+FixedRows,
                                False,
                                ShrinkWithAspectRatio,
                                0,
                                haLeft,
                                vaTop).LoadFromFile(ClientFileName);
            end;
         end;
      end;


      // RowCount 정리
      asg_ChatList.RowCount := asg_ChatList.RowCount - 1;

      // Comments
      lb_UserCnt.Caption := '▶ ' + IntToStr(iRowCnt) + '명 사용중.';


   finally
      FreeAndNil(TpGetSearch);
      Screen.Cursor := crDefault;
   end;
   }

end;




//----------------------------------------------------------------
// String 내의 특정 문자열 삭제
//    - 소스출처 : MComFunc.pas
//----------------------------------------------------------------
function TKDialChat.DeleteStr(OrigStr,DelStr : String) : String;
begin
   while PosByte(DelStr, OrigStr) > 0 do
      System.Delete(OrigStr, PosByte(DelStr, OrigStr), LengthByte(DelStr));

   Result := OrigStr;
end;


procedure TKDialChat.SelChatList(Sender : TObject; in_Gubun : String);
var
   TpGetSearch : TTpSvc;
   iRowCnt, i  : Integer;
   ServerFileName, ClientFileName: String;
   FTP_SVRIP, FTP_USERID, FTP_PASSWD, FTP_HOSTNAME, FTP_DIR : String;
   sLocalFile, sRemoteFile : String;
   DownFile     : Variant;
   DownCnt      : Integer;
   sSearchType,
   sType2,
   sType3,
   sType4,
   sType5       : String;
   ConvMaxVal   : Integer;
   tmp_ChaCallNo: String;
begin
   //-----------------------------------------------------------------
   // 2. 조회
   //-----------------------------------------------------------------
   Screen.Cursor := crHourGlass;


   fgau_Progress.Top       := 543;
   fgau_Progress.Left      := 0;
   fgau_Progress.BringToFront;
   fed_StatusBar.SendToBack;

   fgau_Progress.Progress := 0;


   asg_ChatList.RowCount   := 1;
   asg_ChatList.ClearRows(0, asg_ChatList.RowCount);


   // 나의 RowId 초기화
   iUserRowId := 0;


   // 조회 구분별 Flag 분기 @ 2015.04.09 LSH
   if (in_Gubun = 'NETWORK') then
   begin
      sSearchType := '4';
      sType2      := GetIp;
   end
   else if (in_Gubun = 'MYBOOK') then
   begin
      sSearchType := '40';
      sType2      := GetIp;
   end
   else if (in_Gubun = 'SEARCH') then
   begin
      sSearchType := '41';
      sType2      := Trim(fed_Scan.Text);

      // 접속자 IP 식별후, 근무처별 order by 적용 @ 2014.07.18 LSH
      if PosByte('10.1.', GetIp) > 0 then
         sType3 := '안암'
      else if PosByte('10.2.', GetIp) > 0 then
         sType3 := '구로'
      else if PosByte('10.3.', GetIp) > 0 then
         sType3 := '안산';
   end;


   //-----------------------------------------------------------------
   // 2-1. Load Data by TpSvc
   //-----------------------------------------------------------------
   TpGetSearch := TTpSvc.Create;
   TpGetSearch.Init(Self);





   try
      TpGetSearch.CountField  := 'S_CODE1';
      TpGetSearch.ShowMsgFlag := False;

      if TpGetSearch.GetSvc('MD_KUMCM_S1',
                          ['S_TYPE1'  , sSearchType
                         , 'S_TYPE2'  , sType2
                         , 'S_TYPE3'  , sType3
                         , 'S_TYPE4'  , sType4
                         , 'S_TYPE5'  , sType5
                         , 'S_TYPE6'  , ''
                         , 'S_TYPE7'  , ''
                         , 'S_TYPE8'  , ''
                         , 'S_TYPE9'  , ''
                         , 'S_TYPE10' , ''
                          ],
                          [
                           'S_CODE1'  , 'sLocate'
                         , 'S_CODE2'  , 'sUserId'
                         , 'S_CODE3'  , 'sUserNm'
                         , 'S_CODE4'  , 'sDutyPart'
                         , 'S_CODE5'  , 'sDutySpec'
                         , 'S_CODE6'  , 'sDutyRmk'
                         , 'S_CODE7'  , 'sMobile'
                         , 'S_CODE8'  , 'sEmail'
                         , 'S_CODE9'  , 'sUserIp'
                         , 'S_CODE10' , 'sDutyUser'
                         , 'S_CODE11' , 'sDutyPtnr'
                         , 'S_CODE12' , 'sCallNo'
                         , 'S_CODE13' , 'sDelDate'
                         , 'S_CODE14' , 'sDeptCd'
                         , 'S_CODE15' , 'sFlag'
                         , 'S_CODE16' , 'sEditIp'
                         , 'S_CODE17' , 'sEditDate'
                         , 'S_CODE18' , 'sDocList'
                         , 'S_CODE19' , 'sDocYear'
                         , 'S_CODE20' , 'sDocSeq'
                         , 'S_CODE21' , 'sDocTitle'
                         , 'S_CODE22' , 'sRegDate'
                         , 'S_CODE23' , 'sRegUser'
                         , 'S_CODE24' , 'sRelDept'
                         , 'S_CODE25' , 'sDocRmk'
                         , 'S_CODE26' , 'sDocLoc'
                         , 'S_STRING1', 'sBoardSeq'
                         , 'S_STRING2', 'sCateUp'
                         , 'S_STRING3', 'sCateDown'
                         , 'S_STRING4', 'sContext'
                         , 'S_STRING5', 'sAttachNm'
                         , 'S_STRING6', 'sHideFile'
                         , 'S_STRING7', 'sServerIp'
                         , 'S_STRING8', 'sHeadTail'
                         , 'S_STRING9', 'sHeadSeq'
                         , 'S_STRING10','sTailSeq'
                         , 'S_STRING11','sReplyCnt'
                         , 'S_STRING12','sLikeCnt'
                         , 'S_STRING14','sNickNm'
                         , 'S_STRING15','sLogYn'
                         , 'S_STRING16','sDeptNm'
                         , 'S_STRING17','sDeptSpec'
                         , 'S_STRING18','sDataBase'
                         , 'S_STRING19','sLinkCnt'
                         , 'S_STRING20','sLinkSeq'
                         , 'S_STRING21','sReqNo'
                         , 'S_STRING22','sClienSrc'
                         , 'S_STRING23','sServeSrc'
                         , 'S_STRING24','sCofmUser'
                         , 'S_STRING25','sReleasDt'
                         , 'S_STRING26','sTestDate'
                         , 'S_STRING27','sClient'
                         , 'S_STRING28','sServer'
                         , 'S_STRING29','sRelesUsr'
                          ]) then

         if TpGetSearch.RowCount < 0 then
         begin
            asg_ChatList.RowCount   := 2;

            ShowMessage(GetTxMsg);

            Exit;
         end
         else if TpGetSearch.RowCount = 0 then
         begin
            fed_Statusbar.Text := '▶ 조회건수없음';

            Exit;
         end;




      with asg_ChatList do
      begin
         // Maximum value of progress status
         //fpb_DataLoading.Max := TpGetSearch.RowCount * (asg_Master.ColCount);
         iRowCnt  := TpGetSearch.RowCount;
         RowCount := iRowCnt + FixedRows + 1;

         //ConvMaxVal := iRowCnt div 100;

         {
         if fgau_Progress.MaxValue < iRowCnt * ColCount then
            fgau_Progress.MaxValue := iRowCnt * ColCount;
         }

         //fgau_Progress.MaxValue := iRowCnt;


         for i := 0 to iRowCnt - 1 do
         begin


            Cells[C_CL_PHOTO,    i+FixedRows] := '';


            if (in_Gubun = 'NETWORK') then
               Cells[C_CL_USERNM,   i+FixedRows] := TpGetSearch.GetOutputDataS('sUserNm'    , i) + #13#10 + TpGetSearch.GetOutputDataS('sLocate'    , i) + TpGetSearch.GetOutputDataS('sDutyPart'  , i) + #13#10 + TpGetSearch.GetOutputDataS('sDutyRmk'   , i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo'   , i)
            else if (in_Gubun = 'MYBOOK') then
            begin
               if TpGetSearch.GetOutputDataS('sDutyRmk'   , i) <> '' then
                  Cells[C_CL_USERNM,   i+FixedRows] := TpGetSearch.GetOutputDataS('sUserNm'    , i) + '(' + TpGetSearch.GetOutputDataS('sDutyRmk'   , i) + ')' + #13#10 + TpGetSearch.GetOutputDataS('sLocate'    , i) + #13#10 + TpGetSearch.GetOutputDataS('sDutyPart'  , i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo'   , i) + #13#10 + TpGetSearch.GetOutputDataS('sMobile'   , i)
               else
                  Cells[C_CL_USERNM,   i+FixedRows] := TpGetSearch.GetOutputDataS('sUserNm'    , i) + #13#10 + TpGetSearch.GetOutputDataS('sLocate'    , i) + #13#10 + TpGetSearch.GetOutputDataS('sDutyPart'  , i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo'   , i) + #13#10 + TpGetSearch.GetOutputDataS('sMobile'   , i)
            end
            else if (in_Gubun = 'SEARCH') then
            begin
               // My 다이얼북 D/B
               if PosByte('Book', TpGetSearch.GetOutputDataS('sDataBase', i)) > 0 then
                  Cells[C_CL_USERNM,   i+FixedRows] := '<' + TpGetSearch.GetOutputDataS('sDataBase'     , i) + '>' + #13#10 +
                                                       TpGetSearch.GetOutputDataS('sDutyUser'  , i) + #13#10 +
                                                       TpGetSearch.GetOutputDataS('sLocate'    , i) + ' ' +
                                                       TpGetSearch.GetOutputDataS('sDeptNm'    , i) + #13#10 +
                                                       TpGetSearch.GetOutputDataS('sDeptSpec'  , i) + #13#10 +
                                                       TpGetSearch.GetOutputDataS('sCallNo'    , i)
               else
               begin
                  // 심사팀 담당자 Keyword Searching시, AICHAMST.TELNO1 참조 @ 2015.06.05 LSH
                  if PosByte('심사', TpGetSearch.GetOutputDataS('sDeptNm', i)) > 0 then
                  begin
                     // 심사팀 정보 (AICHAMST) 체크 함수
                     tmp_ChaCallNo :=  GetChaInfo(TpGetSearch.GetOutputDataS('sUserId', i),
                                                  TpGetSearch.GetOutputDataS('sLocate', i));


                     // 검색한 연락처에 심사자 전화번호 있으면, 중복표기 Skip.
                     if PosByte(CopyByte(tmp_ChaCallNo, LengthByte(tmp_ChaCallNo) - 3, 4), TpGetSearch.GetOutputDataS('sCallNo', i)) > 0 then
                        Cells[C_CL_USERNM,   i+FixedRows] := TpGetSearch.GetOutputDataS('sDutyUser'  , i) + #13#10 +
                                                             TpGetSearch.GetOutputDataS('sLocate'    , i) + ' ' +
                                                             TpGetSearch.GetOutputDataS('sDeptNm'    , i) + #13#10 +
                                                             TpGetSearch.GetOutputDataS('sDeptSpec'  , i) + #13#10 +
                                                             TpGetSearch.GetOutputDataS('sCallNo'    , i)
                     else
                        Cells[C_CL_USERNM,   i+FixedRows] := TpGetSearch.GetOutputDataS('sDutyUser'  , i) + #13#10 +
                                                             TpGetSearch.GetOutputDataS('sLocate'    , i) + ' ' +
                                                             TpGetSearch.GetOutputDataS('sDeptNm'    , i) + #13#10 +
                                                             TpGetSearch.GetOutputDataS('sDeptSpec'  , i) + #13#10 +
                                                             tmp_ChaCallNo + #13#10 +
                                                             TpGetSearch.GetOutputDataS('sCallNo'    , i)

                  end
                  else
                  begin
                     Cells[C_CL_USERNM,   i+FixedRows] := TpGetSearch.GetOutputDataS('sDutyUser'  , i) + #13#10 +
                                                          TpGetSearch.GetOutputDataS('sLocate'    , i) + ' ' +
                                                          TpGetSearch.GetOutputDataS('sDeptNm'    , i) + #13#10 +
                                                          TpGetSearch.GetOutputDataS('sDeptSpec'  , i) + #13#10 +
                                                          TpGetSearch.GetOutputDataS('sCallNo'    , i);

                  end;
               end
            end;


            if (in_Gubun <> 'SEARCH') then
            begin
               Cells[C_CL_LOCATE,   i+FixedRows] := TpGetSearch.GetOutputDataS('sLocate'    , i);
               Cells[C_CL_USERID,   i+FixedRows] := TpGetSearch.GetOutputDataS('sUserId'    , i);
               Cells[C_CL_DUTYPART, i+FixedRows] := TpGetSearch.GetOutputDataS('sDutyPart'  , i);
               Cells[C_CL_DUTYRMK,  i+FixedRows] := TpGetSearch.GetOutputDataS('sDutyRmk'   , i);
               Cells[C_CL_MOBILE,   i+FixedRows] := TpGetSearch.GetOutputDataS('sMobile'    , i);
               Cells[C_CL_EMAIL,    i+FixedRows] := TpGetSearch.GetOutputDataS('sEmail'     , i);
               Cells[C_CL_USERIP,   i+FixedRows] := TpGetSearch.GetOutputDataS('sUserIp'    , i);
               Cells[C_CL_PHOTOFILE,i+FixedRows] := DeleteStr(TpGetSearch.GetOutputDataS('sAttachNm'  , i), '/media/cq/photo/');
               Cells[C_CL_HIDEFILE, i+FixedRows] := TpGetSearch.GetOutputDataS('sHideFile'  , i);
               Cells[C_CL_USERORGNM,i+FixedRows] := TpGetSearch.GetOutputDataS('sUserNm'    , i);
            end
            else
            begin
               try
                  Cells[C_CL_LOCATE,   i+FixedRows] := TpGetSearch.GetOutputDataS('sLocate'    , i);
                  Cells[C_CL_USERID,   i+FixedRows] := TpGetSearch.GetOutputDataS('sUserId'    , i);
                  Cells[C_CL_DUTYPART, i+FixedRows] := TpGetSearch.GetOutputDataS('sDeptNm'    , i);
                  Cells[C_CL_DUTYRMK,  i+FixedRows] := TpGetSearch.GetOutputDataS('sDeptSpec'  , i);
                  Cells[C_CL_MOBILE,   i+FixedRows] := TpGetSearch.GetOutputDataS('sCallNo'    , i);
                  Cells[C_CL_EMAIL,    i+FixedRows] := TpGetSearch.GetOutputDataS('sEmail'     , i);
                  Cells[C_CL_USERIP,   i+FixedRows] := TpGetSearch.GetOutputDataS('sUserIp'    , i);
                  Cells[C_CL_PHOTOFILE,i+FixedRows] := DeleteStr(TpGetSearch.GetOutputDataS('sAttachNm'  , i), '/media/cq/photo/');
                  Cells[C_CL_HIDEFILE, i+FixedRows] := TpGetSearch.GetOutputDataS('sHideFile'  , i);
                  Cells[C_CL_USERORGNM,i+FixedRows] := TpGetSearch.GetOutputDataS('sDutyUser'  , i);

                  // 프로필 이미지 Memory 초기화 (Remove) @ 2015.06.03 LSH
                  // JPEG ERROR #41 방지위함
                  RemovePicture(C_CL_PHOTO, i+FixedRows);


               except
                  on E: Exception do
                     ShowMessage(E.message);
               end;
            end;



            try
               AutoSizeRow(i + FixedRows);

            except
               on E: Exception do
                  ShowMessage(E.message);
            end;


            // 현재 접속한 User의 Chat-List 위치를 기억해둔다.
            if Cells[C_CL_USERORGNM,i+FixedRows] = FsUserNm then
               iUserRowId := i+FixedRows;


            //------------------------------------------------------------
            // 접속하려는 Server 사용가능 여부 (현재 Session On/Off) Check
            //------------------------------------------------------------
            if (TpGetSearch.GetOutputDataS('sLogYn', i) = 'Y') then
            begin
               Colors[C_CL_USERNM, i+FixedRows] := $0000CDA3;
               Colors[C_CL_PHOTO, i+FixedRows]  := $0000CDA3;
            end
            else
            begin
               Colors[C_CL_USERNM, i+FixedRows] := clWhite;
               Colors[C_CL_PHOTO, i+FixedRows]  := clWhite;
            end;


            // Progress Gauge ++
            //fgau_Progress.Progress := ((i+FixedRows) div iRowCnt) * 100;
            fgau_Progress.Progress  := (i+FixedRows) * ColCount;
            fed_Statusbar.Text      := '▶ 검색중(' + IntToStr(fgau_Progress.Progress) + '%)..';


            // 사진 표기
            if (Cells[C_CL_PHOTOFILE, i+FixedRows] <> '') then
            begin
               // MIS 기본 이미지 정보인 경우.
               if Cells[C_CL_HIDEFILE, i+FixedRows] = '' then
               begin

                  // 담당자 프로필 이미지 파일명
                  ServerFileName := Cells[C_CL_PHOTOFILE, i+FixedRows];


                  // 사진 이미지가 없는경우 Skip (JPEG ERROR #41 방지) @ 2015.06.03 LSH
                  if AnsiUpperCase(ServerFileName) = '.JPG' then
                     Continue;

                  // Local 파일이름 Set
                  // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
                  ClientFileName := G_HOMEDIR + 'TEMP\SPOOL\' + ServerFileName;



                  // Local에 해당 Image 존재유무 체크
                  // Local file의 Size 체크 추가 (Local에 해당 사번.jpg가 0 size로 존재하면 JPEG Error #41 떨어짐) @ 2015.06.03 LSH
                  if (Not FileExists(ClientFileName)) or
                     (CMsg.GetFileSize(ClientFileName) =  0) then
                  begin
                     // FTP 서버정보 가져오기
                     if Not GetBinUploadInfo(FTP_SVRIP,
                                             FTP_USERID,
                                             FTP_PASSWD,
                                             FTP_HOSTNAME,
                                             FTP_DIR) then
                     begin
                        ShowMessage('다운로드가 실패하여 실행할 수 없습니다.');
                        TUXFTP := nil;
                        Exit;
                     end;


                     // FTP 서버 IP Set
                     FTP_SVRIP        := C_MIS_FTP_IP;


                     // FTP 계정정보 Set
                     FTP_USERID := 'kumis';
                     FTP_PASSWD := 'kumis';
                     FTP_DIR    := '/kumis/app/mis/media/cq/photo/';



                     // Image 다운로드
                     if Not GetBINFTP(FTP_SVRIP,
                                      FTP_USERID,
                                      FTP_PASSWD,
                                      FTP_DIR + ServerFileName,
                                      ClientFileName,
                                      False) then
                     begin
                        Showmessage('MIS 이미지 다운로드에 실패했습니다. 다시한번 시도해 주십시오');

                        TUXFTP := nil;

                        Exit;
                     end;



                     // 해당 Image Variant로 받음.
                     AppendVariant(DownFile, ClientFileName);



                     // 다운로드 횟수 체크
                     DownCnt := DownCnt + 1;

                  end;

                  try
                     // 수신한 Image 파일을 Grid에 표기
                     CreatePicture(C_CL_PHOTO,
                                   i+FixedRows,
                                   False,
                                   ShrinkWithAspectRatio,
                                   0,
                                   haLeft,
                                   vaTop).LoadFromFile(ClientFileName);

                  except
                     // 프로필 이미지 Remove
                     RemovePicture(C_CL_PHOTO, i+FixedRows);

                     Continue;
                     {
                     on E: Exception do
                        ShowMessage(E.message);
                     }
                  end;

               end
               else
               // 다이얼로그 자체 이미지 등록정보 여부 확인
               begin

                  // 파일 업/다운로드를 위한 정보 조회
                  if not GetBinUploadInfo(FTP_SVRIP,
                                          FTP_USERID,
                                          FTP_PASSWD,
                                          FTP_HOSTNAME,
                                          FTP_DIR) then
                  begin
                     MessageDlg('파일 다운을 위한 서버정보 조회중, 오류가 발생했습니다.', Dialogs.mtError, [Dialogs.mbOK], 0);
                     exit;
                  end;


                  // 실제저장된 서버 IP
                  //sServerIp := C_KDIAL_FTP_IP;


                  // 실제 서버에 저장되어 있는 파일명 지정
                  if PosByte('/ftpspool/KDIALFILE/', Cells[C_CL_HIDEFILE, i+FixedRows]) > 0 then
                     sRemoteFile := Cells[C_CL_HIDEFILE, i+FixedRows]
                  else
                     sRemoteFile := '/ftpspool/KDIALFILE/' + Cells[C_CL_HIDEFILE, i+FixedRows];

                  // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
                  sLocalFile  := G_HOMEDIR + 'TEMP\SPOOL\' + Cells[C_CL_PHOTOFILE, i+FixedRows];


                  // 사진 이미지가 없는경우 Skip (JPEG ERROR #41 방지) @ 2015.06.03 LSH
                  if PosByte('SPOOL\.JPG', AnsiUpperCase(sLocalFile)) > 0 then
                     Continue;


                  if (GetBINFTP(FTP_SVRIP, FTP_USERID, FTP_PASSWD, sRemoteFile, sLocalFile, False)) then
                  begin
                     //	정상적인 FTP 다운로드
                  end;

                  // 이미지만 Preview 설정
                  if (PosByte('.bmp', sLocalFile) > 0) or
                     (PosByte('.BMP', sLocalFile) > 0) or
                     (PosByte('.jpg', sLocalFile) > 0) or
                     (PosByte('.JPG', sLocalFile) > 0) or
                     (PosByte('.png', sLocalFile) > 0) or
                     (PosByte('.PNG', sLocalFile) > 0) or
                     (PosByte('.gif', sLocalFile) > 0) or
                     (PosByte('.GIF', sLocalFile) > 0) then

                     try
                        // 수신한 Image 파일을 Grid에 표기
                        CreatePicture(C_CL_PHOTO,
                                      i+FixedRows,
                                      False,
                                      ShrinkWithAspectRatio,
                                      0,
                                      haLeft,
                                      vaTop).LoadFromFile(sLocalFile);

                     except
                        // 프로필 이미지 Remove
                        RemovePicture(C_CL_PHOTO, i+FixedRows);

                        Continue;
                     end;
               end;
            end
            else
            begin
               // 담당자 프로필 이미지 파일명
               if Cells[C_CL_USERID, i+FixedRows] = 'XXXXX' then
               begin
                  // 프로필 이미지 Remove
                  RemovePicture(C_CL_PHOTO, i+FixedRows);

                  // 이미지 등록 버튼 생성
                  //AddButton(C_CL_PHOTO, i+FixedRows, asg_ChatList.ColWidths[C_CL_PHOTO], asg_ChatList.RowHeights[i+FixedRows], '등록', haBeforeText, vaCenter);

                  Continue;
               end
               else
                  // 담당자 HRM 프로필 이미지 파일명
                  ServerFileName := Cells[C_CL_USERID, i+FixedRows] + '.jpg';


                  // 사진 이미지가 없는경우 Skip (JPEG ERROR #41 방지) @ 2015.06.03 LSH
                  if AnsiUpperCase(ServerFileName) = '.JPG' then
                     Continue;


                  // Local 파일이름 Set
                  // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
                  ClientFileName := G_HOMEDIR + 'TEMP\SPOOL\' + ServerFileName;


                  // Local에 해당 Image 존재유무 체크
                  // Local file의 Size 체크 추가 (Local에 해당 사번.jpg가 0 size로 존재하면 JPEG Error #41 떨어짐) @ 2015.06.03 LSH
                  if (Not FileExists(ClientFileName)) or
                     (CMsg.GetFileSize(ClientFileName) =  0) then
                  begin
                     // FTP 서버정보 가져오기
                     if Not GetBinUploadInfo(FTP_SVRIP,
                                             FTP_USERID,
                                             FTP_PASSWD,
                                             FTP_HOSTNAME,
                                             FTP_DIR) then
                     begin
                        ShowMessage('다운로드가 실패하여 실행할 수 없습니다.');
                        TUXFTP := nil;
                        Exit;
                     end;


                     // FTP 서버 IP Set
                     FTP_SVRIP        := C_MIS_FTP_IP;

                     // FTP 계정정보 Set
                     FTP_USERID := 'kuhrm';
                     FTP_PASSWD := 'kuhrm';
                     FTP_DIR    := '/kuhrm/app/hrm/photo/';


                     // HRM Image 다운로드
                     try
                        if Not GetBINFTP(FTP_SVRIP,
                                         FTP_USERID,
                                         FTP_PASSWD,
                                         FTP_DIR + ServerFileName,
                                         ClientFileName,
                                         False) then
                        begin

                           //Showmessage('이미지 다운로드에 실패했습니다. 다시한번 시도해 주십시오');

                           TUXFTP := nil;

                           {
                           Exit;
                           }

                           Continue;
                        end;
                     except
                        // 프로필 이미지 Remove
                        RemovePicture(C_CL_PHOTO, i+FixedRows);

                        Continue;
                     end;


                     // 해당 Image Variant로 받음.
                     AppendVariant(DownFile, ClientFileName);


                     // 다운로드 횟수 체크
                     DownCnt := DownCnt + 1;
                  end;

                  try
                     // 수신한 Image 파일을 Grid에 표기
                     CreatePicture(C_CL_PHOTO,
                                   i+FixedRows,
                                   False,
                                   ShrinkWithAspectRatio,
                                   0,
                                   haLeft,
                                   vaTop).LoadFromFile(ClientFileName);

                  except
                     // 프로필 이미지 Remove
                     RemovePicture(C_CL_PHOTO, i+FixedRows);

                     Continue;
                  end;
            end;
         end;




         Screen.Cursor := crHourglass;

      end;




      // RowCount 정리
      asg_ChatList.RowCount := asg_ChatList.RowCount - 1;

      // Comments
      //lb_UserCnt.Caption := '▶ ' + IntToStr(iRowCnt) + '명 사용중.';



   finally
      asg_ChatList.Row  := iUserRowId;

      // Progress Gauge Completed
      fgau_Progress.Progress := fgau_Progress.MaxValue;
      fed_Statusbar.Text      := '▶ 검색중(' + IntToStr(fgau_Progress.Progress) + '%)..';
      fgau_Progress.Repaint;

      fed_Statusbar.Top    := 542;
      fed_Statusbar.Left   := 1;
      fed_Statusbar.BringToFront;
      fgau_Progress.SendToBack;

      fed_Statusbar.Text   := '▶ ' + IntToStr(iRowCnt) + '건 검색완료';

      FreeAndNil(TpGetSearch);

      Screen.Cursor := crDefault;
   end;


end;


procedure TKDialChat.tm_ChatTimer(Sender: TObject);
begin
   lb_Status.Visible := True;

   // Chat 리스트 Refresh (60초마다)
   SelChatList(Sender, FsSearchGubun);
end;



procedure TKDialChat.asg_ChatListDblClick(Sender: TObject);
var
   FForm     : TForm;
   tmp_MyPhoto,
   tmp_ChatPhoto : String;
begin
   {
   if not IsLogonUser then
   begin
      MessageBox(self.Handle,
                 PChar('[비상연락망] - [담당자정보]에 담당자 IP 등록후 사용하실 수 있습니다.'),
                 PChar(Self.Caption + ' : 등록되지 않은 User 제한 알림'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;
   }


   // 로그 Update
   {  -- 이유모를 Access Violation 오류로 MainDlg.pas에서 로그 update 하도록 변경 @ 2015.03.27 LSH
   UpdateLog('CBLOG',
             asg_Master.Cells[C_M_USERIP, asg_Master.Row],
             FsUserIp,
             'R',
             '',
             '',
             varResult
             );
   }


   //-------------------------------------------------------
   // 미접속 User 채팅 제한 Msg 팝업
   //-------------------------------------------------------
   if asg_ChatList.Colors[C_CL_USERNM, asg_ChatList.Row] <> $0000CDA3 then
   begin
      MessageBox(self.Handle,
                 '현재 KUMC 다이얼로그에 미접속 중입니다.',
                 '다이얼 Chat 미접속 User 대화 제한 알림',
                 MB_ICONERROR + MB_OK);


      Exit;
   end;


   //-------------------------------------------------------
   // 다이얼 Chat - Client (BPL) Form 호출
   //-------------------------------------------------------
   if FindForm('CHATCLIENT') = nil then
      FForm := BplFormCreate('CHATCLIENT', True);


   if Trim(asg_ChatList.Cells[C_CL_PHOTOFILE, asg_ChatList.Row]) = '' then
   begin
      tmp_ChatPhoto := asg_ChatList.Cells[C_CL_USERID, asg_ChatList.Row] + '.jpg';
   end
   else
      tmp_ChatPhoto  := asg_ChatList.Cells[C_CL_PHOTOFILE, asg_ChatList.Row];


   if Trim(asg_ChatList.Cells[C_CL_PHOTOFILE, iUserRowId]) = '' then
   begin
      tmp_MyPhoto := asg_ChatList.Cells[C_CL_USERID, iUserRowId] + '.jpg';
   end
   else
      tmp_MyPhoto := asg_ChatList.Cells[C_CL_PHOTOFILE, iUserRowId];


   // Chat-Client (BPL) 생성
   if FForm <> nil then
   begin
      SetBplStrProp(FForm, 'prop_EditIp',       GetIp);
      SetBplStrProp(FForm, 'prop_EditNick',     FsUserNm);
      SetBplStrProp(FForm, 'prop_ChatIp',       asg_ChatList.Cells[C_CL_USERIP,     asg_ChatList.Row]);
      SetBplStrProp(FForm, 'prop_ChatUserNm',   asg_ChatList.Cells[C_CL_USERORGNM,  asg_ChatList.Row]);
      SetBplStrProp(FForm, 'prop_ChatPhoto' ,   G_HOMEDIR + 'TEMP\SPOOL\' + tmp_ChatPhoto);
      SetBplStrProp(FForm, 'prop_MyPhoto'   ,   G_HOMEDIR + 'TEMP\SPOOL\' + tmp_MyPhoto);

      FForm.Show;
   end;

end;

procedure TKDialChat.FormClose(Sender: TObject; var Action: TCloseAction);
begin
   KDialChat := nil;
   Action    := caFree;
end;

procedure TKDialChat.fsbt_NetworkClick(Sender: TObject);
begin
   tm_Chat.Enabled   := True;

   fed_Statusbar.Text      := '▶ 검색중..';
   fgau_Progress.Progress  := 0;

   fed_Scan.Visible  := False;

   FsSearchGubun := 'NETWORK';

   // Chat 비상연락망 리스트 Refresh
   SelChatList(Sender, FsSearchGubun);
end;

procedure TKDialChat.fsbt_DialBookClick(Sender: TObject);
begin
   tm_Chat.Enabled   := False;

   fed_Statusbar.Text      := '▶ 검색중..';
   fgau_Progress.Progress  := 0;

   fed_Scan.Visible  := False;

   FsSearchGubun := 'MYBOOK';

   // My 다이얼북 리스트 Refresh
   SelChatList(Sender, FsSearchGubun);
end;

procedure TKDialChat.fsbt_SearchClick(Sender: TObject);
begin
   tm_Chat.Enabled         := False;

   fed_Statusbar.Text      := '▶ 검색대기..';
   fgau_Progress.Progress  := 0;

   asg_ChatList.ClearRows(0, asg_ChatList.RowCount);
   asg_ChatList.RowCount   := 0;

   // 입력 Edit-Box 초기화
   fed_Scan.Top      := 30;
   fed_Scan.Left     := 4;
   fed_Scan.Visible  := True;

   fed_Scan.Clear;

   HanKeyChg(fed_Scan.Handle);

   fed_Scan.BringToFront;

   fed_Scan.SetFocus;
end;

procedure TKDialChat.fed_ScanKeyPress(Sender: TObject; var Key: Char);
type
   TProcedureType = Procedure(Sender: TObject) of Object;
var
   Procedure1  : TProcedureType;
   PProcedure1 : PMethod;
begin
   if (fed_Scan.Text <> '') and (Key = #13) then
   begin
      fed_Statusbar.Text      := '▶ 검색중..';

      FsSearchGubun := 'SEARCH';

      // 퀵 서치 Refresh
      SelChatList(Sender, FsSearchGubun);

      //---------------------------------------------------------
      // Main 화면 Chat - Search log 업데이트 진행
      //---------------------------------------------------------
      if FindForm('MAINDLG') <> nil then
      begin
         PProcedure1 := @TMethod(Procedure1);

         (GetComp('MAINDLG', 'pn_ChatSearchKey') as TPanel).Caption := Trim(fed_Scan.Text);

         if FindFormMethod('MAINDLG', 'SetUpdChatSearch', PProcedure1) = True then
         begin
            Procedure1(Sender);
         end;
      end;
   end;
end;

procedure TKDialChat.fed_ScanClick(Sender: TObject);
begin
   HanKeyChg(fed_Scan.Handle);

   fed_Scan.Clear;
end;

procedure TKDialChat.asg_ChatListClick(Sender: TObject);
begin
   lb_Status.Visible := False;
end;


//------------------------------------------------------------------------------
// 한영키를 한글키로 변환
//       - 소스출처 : MComFunc.pas
//
// Author : Lee, Se-Ha
// Date   : 2013.08.22
//------------------------------------------------------------------------------
procedure TKDialChat.HanKeyChg(Handle1:THandle);
var
   tMode : HIMC;
begin
   tMode := ImmGetContext(Handle1);
   ImmSetConversionStatus(tMode, IME_CMODE_HANGUL,
                                 IME_CMODE_HANGUL);
end;



//------------------------------------------------------------------------------
// 병원별 심사팀 담당자 Info 조회
//
// Author : Lee, Se-Ha
// Date   : 2015.06.05
//------------------------------------------------------------------------------
function TKDialChat.GetChaInfo(in_ChaId, in_LocateNm : String) : String;
var
   TpGetCha : TTpSvc;
   iRowCnt, i  : Integer;
   in_Flag  : String;
begin

   // Init.
   Result := '';


   // 병원별 D/B 링크 Dynamic SQL로 쓰지 않고 서비스로 분기 적용
   if (in_LocateNm = '안암') then
      in_Flag := '42'
   else if (in_LocateNm = '구로') then
      in_Flag := '43'
   else if (in_LocateNm = '안산') then
      in_Flag := '44';

   //-----------------------------------------------------------------
   // 1. Load Data by TpSvc
   //-----------------------------------------------------------------
   TpGetCha := TTpSvc.Create;
   TpGetCha.Init(nil);


   Screen.Cursor := crHourglass;


   try
      TpGetCha.CountField  := 'S_CODE12';
      TpGetCha.ShowMsgFlag := False;

      if TpGetCha.GetSvc('MD_KUMCM_S1',
                          ['S_TYPE1'  , in_Flag
                         , 'S_TYPE2'  , in_ChaId
                         , 'S_TYPE3'  , '29991231'
                         , 'S_TYPE4'  , ''
                         , 'S_TYPE5'  , ''
                         , 'S_TYPE6'  , ''
                         , 'S_TYPE7'  , ''
                         , 'S_TYPE8'  , ''
                         , 'S_TYPE9'  , ''
                         , 'S_TYPE10' , ''
                          ],
                          [
                           'S_CODE1'  , 'sLocate'
                         , 'S_CODE2'  , 'sUserId'
                         , 'S_CODE3'  , 'sUserNm'
                         , 'S_CODE4'  , 'sDutyPart'
                         , 'S_CODE5'  , 'sDutySpec'
                         , 'S_CODE6'  , 'sDutyRmk'
                         , 'S_CODE7'  , 'sMobile'
                         , 'S_CODE8'  , 'sEmail'
                         , 'S_CODE9'  , 'sUserIp'
                         , 'S_CODE10' , 'sDutyUser'
                         , 'S_CODE11' , 'sDutyPtnr'
                         , 'S_CODE12' , 'sCallNo'
                         , 'S_CODE13' , 'sDelDate'
                          ]) then

         if TpGetCha.RowCount < 0 then
         begin
            ShowMessage(GetTxMsg);

            Exit;
         end
         else if TpGetCha.RowCount = 0 then
         begin
            Exit;
         end;


      // 유효한 연락처 Return
      Result := TpGetCha.GetOutputDataS('sCallNo', 0)


   finally
      FreeAndNil(TpGetCha);

      Screen.Cursor := crDefault;
   end;


end;



initialization
   RegisterClass(TKDialChat);

finalization
   UnRegisterClass(TKDialChat);

end.
