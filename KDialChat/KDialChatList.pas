{===============================================================================
   Program ID    : KDialChatList.pas
   Program Name  : ���̾�α� Chat ����Ʈ ����(BPL)
   Program Desc. : Chat ��� ����Ʈ(User) ����

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
   // String ���� Ư�� ���ڿ� ����
   //    - �ҽ���ó : MComFunc.pas
   //----------------------------------------------------------------
   function DeleteStr(OrigStr, DelStr : String) : String;


   //----------------------------------------------------------------
   // �ѿ�Ű�� �ѱ�Ű�� ��ȯ
   //    - �ҽ���ó : MComFunc.pas
   //----------------------------------------------------------------
   procedure HanKeyChg(Handle1:THandle);


   //----------------------------------------------------------------
   // ������ �ɻ��� ����� Info ��ȸ
   //    - 2015.06.05 LSH
   //----------------------------------------------------------------
   function GetChaInfo(in_ChaId, in_LocateNm : String) : String;


  public
    { Public declarations }

  published
      property prop_UserNm  : String read FsUserNm     write FsUserNm;


      //----------------------------------------------------------------
      // Chat ����Ʈ ��ȸ @ 2015.03.27 LSH
      //----------------------------------------------------------------
      procedure SelChatList(Sender : TObject; in_Gubun : String);

  end;


const
   // ���̾� Chat Columns
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

   // MIS FTP ���� �ּ�
   C_MIS_FTP_IP   = '10.1.2.17';

   // K-Dial FTP ���� �ּ�
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
   TpSvc,   // TpCall Agent ������
//   Excel97,
   Imm,     // HIMC �������߰� @ 2013.08.22 LSH
   ShellApi,// ShellExec ������ �߰� @ 2013.08.22 LSH
   Winsock, // Windows Sockets API Unit
   ComObj,
   Qrctrls,
   QuickRpt,
   IdFTPCommon;   // ������ XE ���������� Id_FTP --> IndyFTP ��ȯ @ 2014.04.07 LSH


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
   // 1. ��ȸ
   //-----------------------------------------------------------------
   FsSearchGubun := 'MYBOOK'; // Network���� Mybook���� ���� @ 2019.12.03 LSH


   // �Ʒ� ��ȸ�κ��� ĸ��ȭ
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
            lb_UserCnt.Caption  := '�� ��ȸ������ �������� �ʽ��ϴ�.';

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
            // �����Ϸ��� Server ��밡�� ���� (���� Session On/Off) Check
            //------------------------------------------------------------
            if (TpGetSearch.GetOutputDataS('sLogYn', i) = 'Y') then
               Colors[C_CL_USERNM, i+FixedRows] := $0000CDA3
            else
               Colors[C_CL_USERNM, i+FixedRows] := clWhite;





            if (Cells[C_CL_PHOTOFILE, i+FixedRows] <> '') then
            begin
               // MIS �⺻ �̹��� ������ ���.
               if Cells[C_CL_HIDEFILE, i+FixedRows] = '' then
               begin

                  // ����� ������ �̹��� ���ϸ�
                  ServerFileName := Cells[C_CL_PHOTOFILE, i+FixedRows];


                  // Local �����̸� Set
                  ClientFileName := 'C:\Dev\TEMP\' + ServerFileName;



                  // Local�� �ش� Image �������� üũ
                  if Not FileExists(ClientFileName) then
                  begin
                     // FTP �������� ��������
                     if Not GetBinUploadInfo(FTP_SVRIP,
                                             FTP_USERID,
                                             FTP_PASSWD,
                                             FTP_HOSTNAME,
                                             FTP_DIR) then
                     begin
                        ShowMessage('�ٿ�ε尡 �����Ͽ� ������ �� �����ϴ�.');
                        TUXFTP := nil;
                        Exit;
                     end;


                     // FTP ���� IP Set
                     FTP_SVRIP        := C_MIS_FTP_IP;


                     // FTP �������� Set
                     FTP_USERID := 'kumis';
                     FTP_PASSWD := 'kumis';
                     FTP_DIR    := '/kumis/app/mis/media/cq/photo/';



                     // Image �ٿ�ε�
                     if Not GetBINFTP(FTP_SVRIP,
                                      FTP_USERID,
                                      FTP_PASSWD,
                                      FTP_DIR + ServerFileName,
                                      ClientFileName,
                                      False) then
                     begin
                        Showmessage('���� �̹��� �ٿ�ε忡 �����߽��ϴ�. �ٽ��ѹ� �õ��� �ֽʽÿ�');

                        TUXFTP := nil;

                        Exit;
                     end;



                     // �ش� Image Variant�� ����.
                     AppendVariant(DownFile, ClientFileName);



                     // �ٿ�ε� Ƚ�� üũ
                     DownCnt := DownCnt + 1;

                  end;


                  // ������ Image ������ Grid�� ǥ��
                  CreatePicture(C_CL_PHOTO,
                                i+FixedRows,
                                False,
                                ShrinkWithAspectRatio,
                                0,
                                haLeft,
                                vaTop).LoadFromFile(ClientFileName);

               end
               else
               // ���̾�α� ��ü �̹��� ������� ���� Ȯ��
               begin

                  // ���� ��/�ٿ�ε带 ���� ���� ��ȸ
                  if not GetBinUploadInfo(FTP_SVRIP,
                                          FTP_USERID,
                                          FTP_PASSWD,
                                          FTP_HOSTNAME,
                                          FTP_DIR) then
                  begin
                     MessageDlg('���� �ٿ��� ���� �������� ��ȸ��, ������ �߻��߽��ϴ�.', Dialogs.mtError, [Dialogs.mbOK], 0);
                     exit;
                  end;


                  // ��������� ���� IP
                  //sServerIp := C_KDIAL_FTP_IP;


                  // ���� ������ ����Ǿ� �ִ� ���ϸ� ����
                  if PosByte('/ftpspool/KDIALFILE/', Cells[C_CL_HIDEFILE, i+FixedRows]) > 0 then
                     sRemoteFile := Cells[C_CL_HIDEFILE, i+FixedRows]
                  else
                     sRemoteFile := '/ftpspool/KDIALFILE/' + Cells[C_CL_HIDEFILE, i+FixedRows];


                  sLocalFile  := 'C:\Dev\TEMP\SPOOL\' + Cells[C_CL_PHOTOFILE, i+FixedRows];



                  if (GetBINFTP(FTP_SVRIP, FTP_USERID, FTP_PASSWD, sRemoteFile, sLocalFile, False)) then
                  begin
                     //	�������� FTP �ٿ�ε�
                  end;

                  // �̹����� Preview ����
                  if (PosByte('.bmp', sLocalFile) > 0) or
                     (PosByte('.BMP', sLocalFile) > 0) or
                     (PosByte('.jpg', sLocalFile) > 0) or
                     (PosByte('.JPG', sLocalFile) > 0) or
                     (PosByte('.png', sLocalFile) > 0) or
                     (PosByte('.PNG', sLocalFile) > 0) or
                     (PosByte('.gif', sLocalFile) > 0) or
                     (PosByte('.GIF', sLocalFile) > 0) then

                     // ������ Image ������ Grid�� ǥ��
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
               // ����� ������ �̹��� ���ϸ�
               if Cells[C_CL_USERID, i+FixedRows] = 'XXXXX' then
               begin
                  // ������ �̹��� Remove
                  RemovePicture(C_CL_PHOTO, i+FixedRows);

                  // �̹��� ��� ��ư ����
                  //AddButton(C_CL_PHOTO, i+FixedRows, asg_ChatList.ColWidths[C_CL_PHOTO], asg_ChatList.RowHeights[i+FixedRows], '���', haBeforeText, vaCenter);

                  Continue;
               end
               else
                  // ����� HRM ������ �̹��� ���ϸ�
                  ServerFileName := Cells[C_CL_USERID, i+FixedRows] + '.jpg';

                  // Local �����̸� Set
                  ClientFileName := 'C:\Dev\TEMP\' + ServerFileName;


                  // Local�� �ش� Image �������� üũ
                  if Not FileExists(ClientFileName) then
                  begin
                     // FTP �������� ��������
                     if Not GetBinUploadInfo(FTP_SVRIP,
                                             FTP_USERID,
                                             FTP_PASSWD,
                                             FTP_HOSTNAME,
                                             FTP_DIR) then
                     begin
                        ShowMessage('�ٿ�ε尡 �����Ͽ� ������ �� �����ϴ�.');
                        TUXFTP := nil;
                        Exit;
                     end;


                     // FTP ���� IP Set
                     FTP_SVRIP        := C_MIS_FTP_IP;

                     // FTP �������� Set
                     FTP_USERID := 'kuhrm';
                     FTP_PASSWD := 'kuhrm';
                     FTP_DIR    := '/kuhrm/app/hrm/photo/';


                     // HRM Image �ٿ�ε�
                     if Not GetBINFTP(FTP_SVRIP,
                                      FTP_USERID,
                                      FTP_PASSWD,
                                      FTP_DIR + ServerFileName,
                                      ClientFileName,
                                      False) then
                     begin
                        Showmessage('�̹��� �ٿ�ε忡 �����߽��ϴ�. �ٽ��ѹ� �õ��� �ֽʽÿ�');

                        TUXFTP := nil;

                        Exit;
                     end;


                     // �ش� Image Variant�� ����.
                     AppendVariant(DownFile, ClientFileName);


                     // �ٿ�ε� Ƚ�� üũ
                     DownCnt := DownCnt + 1;
                  end;


                  // ������ Image ������ Grid�� ǥ��
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


      // RowCount ����
      asg_ChatList.RowCount := asg_ChatList.RowCount - 1;

      // Comments
      lb_UserCnt.Caption := '�� ' + IntToStr(iRowCnt) + '�� �����.';


   finally
      FreeAndNil(TpGetSearch);
      Screen.Cursor := crDefault;
   end;
   }

end;




//----------------------------------------------------------------
// String ���� Ư�� ���ڿ� ����
//    - �ҽ���ó : MComFunc.pas
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
   // 2. ��ȸ
   //-----------------------------------------------------------------
   Screen.Cursor := crHourGlass;


   fgau_Progress.Top       := 543;
   fgau_Progress.Left      := 0;
   fgau_Progress.BringToFront;
   fed_StatusBar.SendToBack;

   fgau_Progress.Progress := 0;


   asg_ChatList.RowCount   := 1;
   asg_ChatList.ClearRows(0, asg_ChatList.RowCount);


   // ���� RowId �ʱ�ȭ
   iUserRowId := 0;


   // ��ȸ ���к� Flag �б� @ 2015.04.09 LSH
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

      // ������ IP �ĺ���, �ٹ�ó�� order by ���� @ 2014.07.18 LSH
      if PosByte('10.1.', GetIp) > 0 then
         sType3 := '�Ⱦ�'
      else if PosByte('10.2.', GetIp) > 0 then
         sType3 := '����'
      else if PosByte('10.3.', GetIp) > 0 then
         sType3 := '�Ȼ�';
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
            fed_Statusbar.Text := '�� ��ȸ�Ǽ�����';

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
               // My ���̾�� D/B
               if PosByte('Book', TpGetSearch.GetOutputDataS('sDataBase', i)) > 0 then
                  Cells[C_CL_USERNM,   i+FixedRows] := '<' + TpGetSearch.GetOutputDataS('sDataBase'     , i) + '>' + #13#10 +
                                                       TpGetSearch.GetOutputDataS('sDutyUser'  , i) + #13#10 +
                                                       TpGetSearch.GetOutputDataS('sLocate'    , i) + ' ' +
                                                       TpGetSearch.GetOutputDataS('sDeptNm'    , i) + #13#10 +
                                                       TpGetSearch.GetOutputDataS('sDeptSpec'  , i) + #13#10 +
                                                       TpGetSearch.GetOutputDataS('sCallNo'    , i)
               else
               begin
                  // �ɻ��� ����� Keyword Searching��, AICHAMST.TELNO1 ���� @ 2015.06.05 LSH
                  if PosByte('�ɻ�', TpGetSearch.GetOutputDataS('sDeptNm', i)) > 0 then
                  begin
                     // �ɻ��� ���� (AICHAMST) üũ �Լ�
                     tmp_ChaCallNo :=  GetChaInfo(TpGetSearch.GetOutputDataS('sUserId', i),
                                                  TpGetSearch.GetOutputDataS('sLocate', i));


                     // �˻��� ����ó�� �ɻ��� ��ȭ��ȣ ������, �ߺ�ǥ�� Skip.
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

                  // ������ �̹��� Memory �ʱ�ȭ (Remove) @ 2015.06.03 LSH
                  // JPEG ERROR #41 ��������
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


            // ���� ������ User�� Chat-List ��ġ�� ����صд�.
            if Cells[C_CL_USERORGNM,i+FixedRows] = FsUserNm then
               iUserRowId := i+FixedRows;


            //------------------------------------------------------------
            // �����Ϸ��� Server ��밡�� ���� (���� Session On/Off) Check
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
            fed_Statusbar.Text      := '�� �˻���(' + IntToStr(fgau_Progress.Progress) + '%)..';


            // ���� ǥ��
            if (Cells[C_CL_PHOTOFILE, i+FixedRows] <> '') then
            begin
               // MIS �⺻ �̹��� ������ ���.
               if Cells[C_CL_HIDEFILE, i+FixedRows] = '' then
               begin

                  // ����� ������ �̹��� ���ϸ�
                  ServerFileName := Cells[C_CL_PHOTOFILE, i+FixedRows];


                  // ���� �̹����� ���°�� Skip (JPEG ERROR #41 ����) @ 2015.06.03 LSH
                  if AnsiUpperCase(ServerFileName) = '.JPG' then
                     Continue;

                  // Local �����̸� Set
                  // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
                  ClientFileName := G_HOMEDIR + 'TEMP\SPOOL\' + ServerFileName;



                  // Local�� �ش� Image �������� üũ
                  // Local file�� Size üũ �߰� (Local�� �ش� ���.jpg�� 0 size�� �����ϸ� JPEG Error #41 ������) @ 2015.06.03 LSH
                  if (Not FileExists(ClientFileName)) or
                     (CMsg.GetFileSize(ClientFileName) =  0) then
                  begin
                     // FTP �������� ��������
                     if Not GetBinUploadInfo(FTP_SVRIP,
                                             FTP_USERID,
                                             FTP_PASSWD,
                                             FTP_HOSTNAME,
                                             FTP_DIR) then
                     begin
                        ShowMessage('�ٿ�ε尡 �����Ͽ� ������ �� �����ϴ�.');
                        TUXFTP := nil;
                        Exit;
                     end;


                     // FTP ���� IP Set
                     FTP_SVRIP        := C_MIS_FTP_IP;


                     // FTP �������� Set
                     FTP_USERID := 'kumis';
                     FTP_PASSWD := 'kumis';
                     FTP_DIR    := '/kumis/app/mis/media/cq/photo/';



                     // Image �ٿ�ε�
                     if Not GetBINFTP(FTP_SVRIP,
                                      FTP_USERID,
                                      FTP_PASSWD,
                                      FTP_DIR + ServerFileName,
                                      ClientFileName,
                                      False) then
                     begin
                        Showmessage('MIS �̹��� �ٿ�ε忡 �����߽��ϴ�. �ٽ��ѹ� �õ��� �ֽʽÿ�');

                        TUXFTP := nil;

                        Exit;
                     end;



                     // �ش� Image Variant�� ����.
                     AppendVariant(DownFile, ClientFileName);



                     // �ٿ�ε� Ƚ�� üũ
                     DownCnt := DownCnt + 1;

                  end;

                  try
                     // ������ Image ������ Grid�� ǥ��
                     CreatePicture(C_CL_PHOTO,
                                   i+FixedRows,
                                   False,
                                   ShrinkWithAspectRatio,
                                   0,
                                   haLeft,
                                   vaTop).LoadFromFile(ClientFileName);

                  except
                     // ������ �̹��� Remove
                     RemovePicture(C_CL_PHOTO, i+FixedRows);

                     Continue;
                     {
                     on E: Exception do
                        ShowMessage(E.message);
                     }
                  end;

               end
               else
               // ���̾�α� ��ü �̹��� ������� ���� Ȯ��
               begin

                  // ���� ��/�ٿ�ε带 ���� ���� ��ȸ
                  if not GetBinUploadInfo(FTP_SVRIP,
                                          FTP_USERID,
                                          FTP_PASSWD,
                                          FTP_HOSTNAME,
                                          FTP_DIR) then
                  begin
                     MessageDlg('���� �ٿ��� ���� �������� ��ȸ��, ������ �߻��߽��ϴ�.', Dialogs.mtError, [Dialogs.mbOK], 0);
                     exit;
                  end;


                  // ��������� ���� IP
                  //sServerIp := C_KDIAL_FTP_IP;


                  // ���� ������ ����Ǿ� �ִ� ���ϸ� ����
                  if PosByte('/ftpspool/KDIALFILE/', Cells[C_CL_HIDEFILE, i+FixedRows]) > 0 then
                     sRemoteFile := Cells[C_CL_HIDEFILE, i+FixedRows]
                  else
                     sRemoteFile := '/ftpspool/KDIALFILE/' + Cells[C_CL_HIDEFILE, i+FixedRows];

                  // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
                  sLocalFile  := G_HOMEDIR + 'TEMP\SPOOL\' + Cells[C_CL_PHOTOFILE, i+FixedRows];


                  // ���� �̹����� ���°�� Skip (JPEG ERROR #41 ����) @ 2015.06.03 LSH
                  if PosByte('SPOOL\.JPG', AnsiUpperCase(sLocalFile)) > 0 then
                     Continue;


                  if (GetBINFTP(FTP_SVRIP, FTP_USERID, FTP_PASSWD, sRemoteFile, sLocalFile, False)) then
                  begin
                     //	�������� FTP �ٿ�ε�
                  end;

                  // �̹����� Preview ����
                  if (PosByte('.bmp', sLocalFile) > 0) or
                     (PosByte('.BMP', sLocalFile) > 0) or
                     (PosByte('.jpg', sLocalFile) > 0) or
                     (PosByte('.JPG', sLocalFile) > 0) or
                     (PosByte('.png', sLocalFile) > 0) or
                     (PosByte('.PNG', sLocalFile) > 0) or
                     (PosByte('.gif', sLocalFile) > 0) or
                     (PosByte('.GIF', sLocalFile) > 0) then

                     try
                        // ������ Image ������ Grid�� ǥ��
                        CreatePicture(C_CL_PHOTO,
                                      i+FixedRows,
                                      False,
                                      ShrinkWithAspectRatio,
                                      0,
                                      haLeft,
                                      vaTop).LoadFromFile(sLocalFile);

                     except
                        // ������ �̹��� Remove
                        RemovePicture(C_CL_PHOTO, i+FixedRows);

                        Continue;
                     end;
               end;
            end
            else
            begin
               // ����� ������ �̹��� ���ϸ�
               if Cells[C_CL_USERID, i+FixedRows] = 'XXXXX' then
               begin
                  // ������ �̹��� Remove
                  RemovePicture(C_CL_PHOTO, i+FixedRows);

                  // �̹��� ��� ��ư ����
                  //AddButton(C_CL_PHOTO, i+FixedRows, asg_ChatList.ColWidths[C_CL_PHOTO], asg_ChatList.RowHeights[i+FixedRows], '���', haBeforeText, vaCenter);

                  Continue;
               end
               else
                  // ����� HRM ������ �̹��� ���ϸ�
                  ServerFileName := Cells[C_CL_USERID, i+FixedRows] + '.jpg';


                  // ���� �̹����� ���°�� Skip (JPEG ERROR #41 ����) @ 2015.06.03 LSH
                  if AnsiUpperCase(ServerFileName) = '.JPG' then
                     Continue;


                  // Local �����̸� Set
                  // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
                  ClientFileName := G_HOMEDIR + 'TEMP\SPOOL\' + ServerFileName;


                  // Local�� �ش� Image �������� üũ
                  // Local file�� Size üũ �߰� (Local�� �ش� ���.jpg�� 0 size�� �����ϸ� JPEG Error #41 ������) @ 2015.06.03 LSH
                  if (Not FileExists(ClientFileName)) or
                     (CMsg.GetFileSize(ClientFileName) =  0) then
                  begin
                     // FTP �������� ��������
                     if Not GetBinUploadInfo(FTP_SVRIP,
                                             FTP_USERID,
                                             FTP_PASSWD,
                                             FTP_HOSTNAME,
                                             FTP_DIR) then
                     begin
                        ShowMessage('�ٿ�ε尡 �����Ͽ� ������ �� �����ϴ�.');
                        TUXFTP := nil;
                        Exit;
                     end;


                     // FTP ���� IP Set
                     FTP_SVRIP        := C_MIS_FTP_IP;

                     // FTP �������� Set
                     FTP_USERID := 'kuhrm';
                     FTP_PASSWD := 'kuhrm';
                     FTP_DIR    := '/kuhrm/app/hrm/photo/';


                     // HRM Image �ٿ�ε�
                     try
                        if Not GetBINFTP(FTP_SVRIP,
                                         FTP_USERID,
                                         FTP_PASSWD,
                                         FTP_DIR + ServerFileName,
                                         ClientFileName,
                                         False) then
                        begin

                           //Showmessage('�̹��� �ٿ�ε忡 �����߽��ϴ�. �ٽ��ѹ� �õ��� �ֽʽÿ�');

                           TUXFTP := nil;

                           {
                           Exit;
                           }

                           Continue;
                        end;
                     except
                        // ������ �̹��� Remove
                        RemovePicture(C_CL_PHOTO, i+FixedRows);

                        Continue;
                     end;


                     // �ش� Image Variant�� ����.
                     AppendVariant(DownFile, ClientFileName);


                     // �ٿ�ε� Ƚ�� üũ
                     DownCnt := DownCnt + 1;
                  end;

                  try
                     // ������ Image ������ Grid�� ǥ��
                     CreatePicture(C_CL_PHOTO,
                                   i+FixedRows,
                                   False,
                                   ShrinkWithAspectRatio,
                                   0,
                                   haLeft,
                                   vaTop).LoadFromFile(ClientFileName);

                  except
                     // ������ �̹��� Remove
                     RemovePicture(C_CL_PHOTO, i+FixedRows);

                     Continue;
                  end;
            end;
         end;




         Screen.Cursor := crHourglass;

      end;




      // RowCount ����
      asg_ChatList.RowCount := asg_ChatList.RowCount - 1;

      // Comments
      //lb_UserCnt.Caption := '�� ' + IntToStr(iRowCnt) + '�� �����.';



   finally
      asg_ChatList.Row  := iUserRowId;

      // Progress Gauge Completed
      fgau_Progress.Progress := fgau_Progress.MaxValue;
      fed_Statusbar.Text      := '�� �˻���(' + IntToStr(fgau_Progress.Progress) + '%)..';
      fgau_Progress.Repaint;

      fed_Statusbar.Top    := 542;
      fed_Statusbar.Left   := 1;
      fed_Statusbar.BringToFront;
      fgau_Progress.SendToBack;

      fed_Statusbar.Text   := '�� ' + IntToStr(iRowCnt) + '�� �˻��Ϸ�';

      FreeAndNil(TpGetSearch);

      Screen.Cursor := crDefault;
   end;


end;


procedure TKDialChat.tm_ChatTimer(Sender: TObject);
begin
   lb_Status.Visible := True;

   // Chat ����Ʈ Refresh (60�ʸ���)
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
                 PChar('[��󿬶���] - [���������]�� ����� IP ����� ����Ͻ� �� �ֽ��ϴ�.'),
                 PChar(Self.Caption + ' : ��ϵ��� ���� User ���� �˸�'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;
   }


   // �α� Update
   {  -- ������ Access Violation ������ MainDlg.pas���� �α� update �ϵ��� ���� @ 2015.03.27 LSH
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
   // ������ User ä�� ���� Msg �˾�
   //-------------------------------------------------------
   if asg_ChatList.Colors[C_CL_USERNM, asg_ChatList.Row] <> $0000CDA3 then
   begin
      MessageBox(self.Handle,
                 '���� KUMC ���̾�α׿� ������ ���Դϴ�.',
                 '���̾� Chat ������ User ��ȭ ���� �˸�',
                 MB_ICONERROR + MB_OK);


      Exit;
   end;


   //-------------------------------------------------------
   // ���̾� Chat - Client (BPL) Form ȣ��
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


   // Chat-Client (BPL) ����
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

   fed_Statusbar.Text      := '�� �˻���..';
   fgau_Progress.Progress  := 0;

   fed_Scan.Visible  := False;

   FsSearchGubun := 'NETWORK';

   // Chat ��󿬶��� ����Ʈ Refresh
   SelChatList(Sender, FsSearchGubun);
end;

procedure TKDialChat.fsbt_DialBookClick(Sender: TObject);
begin
   tm_Chat.Enabled   := False;

   fed_Statusbar.Text      := '�� �˻���..';
   fgau_Progress.Progress  := 0;

   fed_Scan.Visible  := False;

   FsSearchGubun := 'MYBOOK';

   // My ���̾�� ����Ʈ Refresh
   SelChatList(Sender, FsSearchGubun);
end;

procedure TKDialChat.fsbt_SearchClick(Sender: TObject);
begin
   tm_Chat.Enabled         := False;

   fed_Statusbar.Text      := '�� �˻����..';
   fgau_Progress.Progress  := 0;

   asg_ChatList.ClearRows(0, asg_ChatList.RowCount);
   asg_ChatList.RowCount   := 0;

   // �Է� Edit-Box �ʱ�ȭ
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
      fed_Statusbar.Text      := '�� �˻���..';

      FsSearchGubun := 'SEARCH';

      // �� ��ġ Refresh
      SelChatList(Sender, FsSearchGubun);

      //---------------------------------------------------------
      // Main ȭ�� Chat - Search log ������Ʈ ����
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
// �ѿ�Ű�� �ѱ�Ű�� ��ȯ
//       - �ҽ���ó : MComFunc.pas
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
// ������ �ɻ��� ����� Info ��ȸ
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


   // ������ D/B ��ũ Dynamic SQL�� ���� �ʰ� ���񽺷� �б� ����
   if (in_LocateNm = '�Ⱦ�') then
      in_Flag := '42'
   else if (in_LocateNm = '����') then
      in_Flag := '43'
   else if (in_LocateNm = '�Ȼ�') then
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


      // ��ȿ�� ����ó Return
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