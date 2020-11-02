{===============================================================================
   Program ID    : DIALOG_XE7.exe [D5 >> XE7]
   Program Name  : 고대의료원 정보전산팀 비상연락망 및 업무관리
   Program Desc. : 1. 담당자정보 업무프로필 업데이트를 통한 비상연락망 관리
                   2. 발송/협조 등 대내외 문서 관리 (검색/등록/수정/삭제)
                   3. 기타 커뮤니케이션 업무

   Author        : Lee, Se-Ha
   Date          : 2013.08.08
===============================================================================}
unit MainDialog;

{$IFNDEF FREE_PARKING}
{$DEFINE FREE_PARKING}

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, StdCtrls, Grids, BaseGrid, AdvGrid, ComCtrls, Buttons, AdvPanel,
  TFlatButtonUnit, TFlatSpeedButtonUnit, TFlatTabControlUnit,
  TFlatAnimationUnit, TFlatPanelUnit, TFlatEditUnit, TFlatMemoUnit,
  TFlatComboBoxUnit, DlgAbout, DlgUser,
  ExtDlgs, ScktComp, Menus, TFlatListBoxUnit, TFlatCheckBoxUnit, Mask,
  TFlatMaskEditUnit, DlgPrint, TFlatProgressBarUnit, TFlatAnimWndUnit, 
  IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdFTP, HTMLPopup, 
  Variants, VclTee.TeEngine, VclTee.TeeProcs, VclTee.Chart, AdvObj,
  IdExplicitTLSClientServerBase;
   //, RXCtrls;


const
   // 주요업무공유 Columns
   C_WC_LOCATE    = 0;
   C_WC_DUTYPART  = 1;
   C_WC_CONTEXT   = 2;
   C_WC_REGDATE   = 3;
   C_WC_DUTYUSER  = 4;
   C_WC_CONFIRM   = 5;

   // D/B 디테일(상세)
   C_DBD_SECCD    = 0;
   C_DBD_SECCDNM  = 1;
   C_DBD_3RDCD    = 2;
   C_DBD_3RDCDNM  = 3;
   C_DBD_WORKINFO = 4;

   // D/B 마스터
   C_DBM_FSTCD    = 0;
   C_DBM_FSTCDNM  = 1;
   C_DBM_WORKINFO = 2;

   // 근무 및 휴가 현황 Columns
   C_DT_DUTYDATE  = 0;
   C_DT_YOIL      = 1;
   C_DT_DUTYFLAG  = 2;
   C_DT_DUTYUSER  = 3;
   C_DT_REMARK    = 4;
   C_DT_SEQNO     = 5;
   C_DT_ORGDTYUSR = 6;
   C_DT_ORGREMARK = 7;

   // 개발관련문서(릴리즈대장) Columns
   C_RL_REGDATE   = 0;
   C_RL_DUTYSPEC  = 1;
   C_RL_CONTEXT   = 2;
   C_RL_DUTYUSER  = 3;
   C_RL_COFMUSER  = 4;
   C_RL_RELEASDT  = 5;
   C_RL_TESTDATE  = 6;
   C_RL_REQUSER   = 7;
   C_RL_SRREQNO   = 8;
   C_RL_CLIENSRC  = 9;
   C_RL_CLIENT    = 10;
   C_RL_SERVESRC  = 11;
   C_RL_SERVER    = 12;
   C_RL_REMARK    = 13;
   C_RL_RELEASUSER= 14;
   C_RL_REQDEPT   = 15;


   // 문서관리 ID 상세정보 Rows
   R_DS_USERID    = 0;
   R_DS_USERNM    = 1;
   R_DS_DEPTCD    = 2;
   R_DS_DEPTNM    = 3;
   R_DS_SOCNO     = 4;
   R_DS_STARTDT   = 5;
   R_DS_ENDDT     = 6;
   R_DS_MOBILE    = 7;
   R_DS_GRPID     = 8;
   R_DS_USELVL    = 9;
   R_DS_BUTTON    = 10;

   // 담당자 프로필 Rows
   R_PR_USERNM    = 0;
   R_PR_DUTYPART  = 1;
   R_PR_REMARK    = 2;
   R_PR_CALLNO    = 3;
   R_PR_DUTYSCH   = 4;
   R_PR_PHOTO     = 5;
   R_PR_USERID    = 6;

   // MIS FTP 서버 주소
   C_MIS_FTP_IP   = '원내 MIS FTP 주소';

   // K-Dial FTP 서버 주소 --> ISMS 인증 후속조치 관련 개발계 IP 변경 @ 2017.10.31 LSH
   C_KDIAL_FTP_IP = '자체 FTP 주소';


   // 게시판(커뮤니티) Page MaxRow 및 Block 단위
   C_PAGE_MAXROW_UNIT = 20;
   C_PAGE_BLOCK_UNIT  = 10;


   // 리포트 Columns
   C_WK_DATE      = 0;
   C_WK_GUBUN     = 1;
   C_WK_CONTEXT   = 2;
   C_WK_STEP      = 3;


   // 업무일지(KUMC) Columns
   C_WR_LOCATE    = 0;
   C_WR_DUTYUSER  = 1;
   C_WR_DUTYPART  = 2;
   C_WR_REGDATE   = 3;
   C_WR_CONTEXT   = 4;
   C_WR_CONFIRM   = 5;


   // 업무(S/R) 상세 리스트 Columns
   C_AL_TITLE     = 0;
   C_AL_CONTEXT   = 1;
   C_AL_ADDINFO   = 2;


   // 업무(S/R) 분석 Columns
   C_AN_REQFLAG   = 0;
   C_AN_SRTITLE   = 1;
   C_AN_REQDEPT   = 2;
   C_AN_REQUSER   = 3;
   C_AN_REQDATE   = 4;
   C_AN_DUTYUSER  = 5;
   C_AN_PROCESS   = 6;
   C_AN_FILECNT   = 7;
   C_AN_REQNO     = 8;
   C_AN_REQOBJT   = 9;
   C_AN_REQRMK    = 10;
   C_AN_ANSRMK    = 11;
   C_AN_REQISSU   = 12;
   C_AN_ATTACHNM  = 13;
   C_AN_USERID    = 14;
   C_AN_REQDEPTNM = 15;
   C_AN_REQTELNO  = 16;
   C_AN_REQSPEC   = 17;


   // 링크된 다이얼 Book 리스트 Coloumns
   C_LL_USERNM    = 0;
   C_LL_CALLNO    = 1;
   C_LL_DUTYUSER  = 2;
   C_LL_MOBILE    = 3;
   C_LL_LINKDB    = 4;
   C_LL_LINKSEQ   = 5;
   C_LL_DUTYRMK   = 6;
   C_LL_DTYUSRID  = 7;


   // My 다이얼 Book Columns
   C_MD_LOCATE    = 0;
   C_MD_DEPTNM    = 1;
   C_MD_DEPTSPEC  = 2;
   C_MD_CALLNO    = 3;
   C_MD_DUTYUSER  = 4;
   C_MD_MOBILE    = 5;
   C_MD_REMARK    = 6;
   C_MD_LINKDB    = 7;
   C_MD_LINKSEQ   = 8;
   C_MD_DTYUSRID  = 9;


   // 다이얼 Book 검색리스트 Columns
   C_DL_LOCATE    = 0;
   C_DL_DEPTNM    = 1;
   C_DL_DEPTSPEC  = 2;
   C_DL_CALLNO    = 3;
   C_DL_DUTYUSER  = 4;
   C_DL_LINKDB    = 5;
   C_DL_LINKCNT   = 6;
   C_DL_LINKSEQ   = 7;
   C_DL_DTYUSRID  = 8;


   // 다이얼 Book 최근검색 Rows
   C_DS_KEYWORD = 0;


   // flat Tab-Contrl Active Tab 인덱스
   AT_DIALOG    = 0;

   { -- 비상연락망 Tab 통합에 따른 주석 @ 2013.12.30 LSH
   AT_DETAIL   = 1;
   AT_MASTER   = 2;
   }

   AT_DIALBOOK = 1;
   AT_BOARD    = 2;
   AT_DOC      = 3;
   AT_ANALYSIS = 4;
   AT_WORKCONN = 5;


   // 식단 메뉴 Columns
   C_MN_MON       = 0;
   C_MN_TUE       = 1;
   C_MN_WED       = 2;
   C_MN_THU       = 3;
   C_MN_FRI       = 4;
   C_MN_SAT       = 5;
   C_MN_SUN       = 6;


   // Chat Send Columns
   C_SD_TEXT      = 0;
   C_SD_BUTTON    = 1;


   // Chat 박스 Columns
   C_CH_MYCOL     = 0;
   C_CH_TEXT      = 1;
   C_CH_YOURCOL   = 2;
   C_CH_HIDDEN    = 3;


   // 커뮤니티(게시판) Columns
   C_B_BOARDSEQ   = 0;
   C_B_CATEUP     = 1;
   C_B_TITLE      = 2;
   C_B_REGDATE    = 3;
   C_B_REGUSER    = 4;
   C_B_LIKE       = 5;
   C_B_ATTACH     = 6;
   C_B_HEADSEQ    = 7;
   C_B_TAILSEQ    = 8;
   C_B_CLOSEYN    = 9;
   C_B_CONTEXT    = 10;
   C_B_USERIP     = 11;
   C_B_REPLY      = 12;
   C_B_HEADTAIL   = 13;
   C_B_HIDEFILE   = 14;
   C_B_SERVERIP   = 15;


   // 문서 등록/수정 Columns
   C_RD_GUBUN     = 0;
   C_RD_DOCLIST   = 1;
   C_RD_DOCYEAR   = 2;
   C_RD_DOCSEQ    = 3;
   C_RD_DOCTITLE  = 4;
   C_RD_REGDATE   = 5;
   C_RD_REGUSER   = 6;
   C_RD_RELDEPT   = 7;
   C_RD_DOCRMK    = 8;
   C_RD_DOCLOC    = 9;
   C_RD_BUTTON    = 10;


   // 문서 조회 Columns
   C_SD_DOCLIST   = 0;
   C_SD_DOCSEQ    = 1;
   C_SD_DOCTITLE  = 2;
   C_SD_REGDATE   = 3;
   C_SD_REGUSER   = 4;
   C_SD_RELDEPT   = 5;
   C_SD_DOCRMK    = 6;
   C_SD_DOCLOC    = 7;
   C_SD_HQREMARK  = 8;
   C_SD_AAREMARK  = 9;
   C_SD_GRREMARK  = 10;
   C_SD_ASREMARK  = 11;
   C_SD_TOTRMK    = 12;


   // 비상연락망 Columns
   C_NW_LOCATE    = 0;
   C_NW_DUTYPART  = 1;
   C_NW_DUTYSPEC  = 2;
   C_NW_USERNM    = 3;
   C_NW_CALLNO    = 4;
   C_NW_MOBILE    = 5;
   C_NW_DUTYRMK   = 6;
   C_NW_EMAIL     = 7;
   C_NW_USERIP    = 8;
   C_NW_PHOTOFILE = 9;
   C_NW_HIDEFILE  = 10;
   C_NW_USERID    = 11;

   // 최근 업데이트 Columns
   C_NU_EDITDATE = 0;
   C_NU_EDITTEXT = 1;
   C_NU_EDITIP   = 2;


   // 업무프로필 Columns
   C_D_LOCATE   = 0;
   C_D_DUTYPART = 1;
   C_D_DUTYSPEC = 2;
   C_D_DUTYRMK  = 3;
   C_D_DUTYUSER = 4;
   C_D_CALLNO   = 5;
   C_D_DUTYPTNR = 6;
   C_D_DELDATE  = 7;
   C_D_BUTTON   = 8;
   C_D_USERADD  = 9;


   // 담당자정보 Columns
   C_M_LOCATE   = 0;
   C_M_USERID   = 1;
   C_M_USERNM   = 2;
   C_M_DUTYPART = 3;
   C_M_MOBILE   = 4;
   C_M_EMAIL    = 5;
   C_M_USERIP   = 6;
   C_M_DELDATE  = 7;
   C_M_BUTTON   = 8;
   C_M_USERADD  = 9;
   C_M_NICKNM   = 10;

   // 빅데이터 Lab. Columns @ 2015.04.02 LSH
   C_BD_YEAR		 = 0;
   C_BD_SEQNO      = 1;
   C_BD_SHOWDT     = 2;
   C_BD_NO1        = 3;
   C_BD_NO2        = 4;
   C_BD_NO3        = 5;
   C_BD_NO4        = 6;
   C_BD_NO5        = 7;
   C_BD_NO6        = 8;
   C_BD_NOBONUS    = 9;
   C_BD_1STCNT     = 10;
   C_BD_1STAMT     = 11;
   C_BD_2NDCNT     = 12;
   C_BD_2NDAMT     = 13;
   C_BD_3THCNT     = 14;
   C_BD_3THAMT     = 15;
   C_BD_4THCNT     = 16;
   C_BD_4THAMT     = 17;
   C_BD_5THCNT     = 18;
   C_BD_5THAMT     = 19;




type
  TMainDlg = class(TForm)
    pc_Dialog: TPageControl;
    ts_Dialog: TTabSheet;
    GroupBox2_: TGroupBox;
    ts_Detail: TTabSheet;
    ts_Master: TTabSheet;
    asg_Master_: TAdvStringGrid;
    TabSheet4: TTabSheet;
    AdvStringGrid3: TAdvStringGrid;
    Button1: TButton;
    Button2: TButton;
    RadioGroup2: TRadioGroup;
    asg_NwUpdate_: TAdvStringGrid;
    bbt_MasterPrint: TBitBtn;
    PrintDialog1: TPrintDialog;
    asg_Detail_: TAdvStringGrid;
    bbt_DetailPrint: TBitBtn;
    bbt_MasterRefresh: TBitBtn;
    bbt_DetailRefresh: TBitBtn;
    asg_Network_: TAdvStringGrid;
    bbt_NetworkRefresh: TBitBtn;
    bbt_NetworkPrint: TBitBtn;
    BitBtn1: TBitBtn;
    Advpn_Log: TAdvPanel;
    Memo1: TMemo;
    ts_Doc: TTabSheet;
    asg_RegDoc_: TAdvStringGrid;
    asg_SelDoc_: TAdvStringGrid;
    lb_Detail_: TLabel;
    lb_Master_: TLabel;
    lb_RegDoc_: TLabel;
    lb_SelDoc_: TLabel;
    Label1: TLabel;
    ts_Board: TTabSheet;
    asg_Board_: TAdvStringGrid;
    ftc_Dialog: TFlatTabControl;
    fpn_Dialog: TFlatPanel;
    BitBtn4: TBitBtn;
    lb_Network: TLabel;
    fpn_Detail: TFlatPanel;
    lb_Detail: TLabel;
    fpn_Master: TFlatPanel;
    lb_Master: TLabel;
    fpn_Doc: TFlatPanel;
    fpn_Board: TFlatPanel;
    asg_Board: TAdvStringGrid;
    fsbt_NetworkRefresh: TFlatSpeedButton;
    fsbt_NetworkPrint: TFlatSpeedButton;
    FlatSpeedButton3: TFlatSpeedButton;
    FlatSpeedButton4: TFlatSpeedButton;
    FlatSpeedButton5: TFlatSpeedButton;
    FlatSpeedButton6: TFlatSpeedButton;
    fsbt_Version: TFlatSpeedButton;
    fsbt_Write: TFlatSpeedButton;
    fpn_Write: TFlatPanel;
    fcb_CateUp: TFlatComboBox;
    fmm_Text: TFlatMemo;
    fed_Gubun: TFlatEdit;
    FlatEdit1: TFlatEdit;
    FlatEdit2: TFlatEdit;
    fed_Title: TFlatEdit;
    FlatEdit4: TFlatEdit;
    fed_Attached: TFlatEdit;
    FlatEdit6: TFlatEdit;
    fed_Writer: TFlatEdit;
    fsbt_WriteCancel: TFlatSpeedButton;
    fsbt_WriteReg: TFlatSpeedButton;
    fed_CateDownNm: TFlatEdit;
    fcb_CateDown: TFlatComboBox;
    FlatEdit5: TFlatEdit;
    fed_BoardSeq: TFlatEdit;
    fsbt_Reply: TFlatSpeedButton;
    fsbt_Delete: TFlatSpeedButton;
    lb_HeadTail: TLabel;
    lb_HeadSeq: TLabel;
    lb_TailSeq: TLabel;
    fsbt_BoardRefresh: TFlatSpeedButton;
    lb_Board: TLabel;
    fcb_Page: TFlatComboBox;
    fsbt_BackWard: TFlatSpeedButton;
    fsbt_ForWard: TFlatSpeedButton;
    FlatSpeedButton10: TFlatSpeedButton;
    FlatSpeedButton11: TFlatSpeedButton;
    lb_Write: TLabel;
    fsbt_Search: TFlatSpeedButton;
    fsbt_WriteFind: TFlatSpeedButton;
    fed_AlarmFro: TFlatEdit;
    fmed_AlarmFro: TFlatMaskEdit;
    fed_AlarmTo: TFlatEdit;
    fmed_AlarmTo: TFlatMaskEdit;
    Label3: TLabel;
    fsbt_FileOpen: TFlatSpeedButton;
    fsbt_FileClear: TFlatSpeedButton;
    od_File: TOpenDialog;
    Id_FTP: TIdFTP;
    apn_LikeList: TAdvPanel;
    flbx_Like: TFlatListBox;
    ServerSocket: TServerSocket;
    apn_ChatBox: TAdvPanel;
    Memo2: TMemo;
    EditSay: TEdit;
    ButtonSend: TButton;
    asg_Chat: TAdvStringGrid;
    asg_ChatSend: TAdvStringGrid;
    Image2: TImage;
    pm_Chat: TPopupMenu;
    mi_Emoti: TMenuItem;
    mi_FileSend: TMenuItem;
    tm_Master: TTimer;
    OpenPictureDialog1: TOpenPictureDialog;
    tm_TxInit: TTimer;
    apn_Menu: TAdvPanel;
    asg_BigData: TAdvStringGrid;
    asg_MenuBar: TAdvStringGrid;
    fsbt_Menu: TFlatSpeedButton;
    pm_Menu: TPopupMenu;
    mi_MenuOpen: TMenuItem;
    opendlg_File: TOpenDialog;
    fcb_MenuLoc: TFlatComboBox;
    FlatSpeedButton8: TFlatSpeedButton;
    FlatSpeedButton9: TFlatSpeedButton;
    fpn_DialBook: TFlatPanel;
    lb_DialScan: TLabel;
    Label7: TLabel;
    asg_DialScan: TAdvStringGrid;
    asg_DialList: TAdvStringGrid;
    asg_MyDial: TAdvStringGrid;
    lb_MyDial: TLabel;
    fed_Scan: TFlatEdit;
    fcb_Scan: TFlatComboBox;
    FlatSpeedButton12: TFlatSpeedButton;
    FlatSpeedButton13: TFlatSpeedButton;
    FlatSpeedButton14: TFlatSpeedButton;
    Label6: TLabel;
    apn_LinkList: TAdvPanel;
    asg_LinkList: TAdvStringGrid;
    fpn_Analysis: TFlatPanel;
    lb_Analysis: TLabel;
    Label9: TLabel;
    AdvStringGrid1: TAdvStringGrid;
    asg_Analysis: TAdvStringGrid;
    fed_Analysis: TFlatEdit;
    fcb_Analysis: TFlatComboBox;
    fmed_AnalFrom: TFlatMaskEdit;
    fmed_AnalTo: TFlatMaskEdit;
    fed_DateTitle: TFlatEdit;
    Label8: TLabel;
    fsbt_WorkRpt: TFlatSpeedButton;
    asg_WorkRpt: TAdvStringGrid;
    fsbt_Analysis: TFlatSpeedButton;
    fcb_WorkArea: TFlatComboBox;
    fcb_WorkGubun: TFlatComboBox;
    fsbt_Weekly: TFlatSpeedButton;
    fsbt_Print: TFlatSpeedButton;
    fsbt_DialMap: TFlatSpeedButton;
    apn_DialMap: TAdvPanel;
    asg_DialMap: TAdvStringGrid;
    fcb_ReScan: TFlatCheckBox;
    fsbt_AddMyDial: TFlatSpeedButton;
    apn_AddMyDial: TAdvPanel;
    asg_AddMyDial: TAdvStringGrid;
    apn_Congrats: TAdvPanel;
    fpn_Title: TFlatPanel;
    FlatSpeedButton15: TFlatSpeedButton;
    fpn_Photo: TFlatPanel;
    fani_Photo: TFlatAnimation;
    fcb_Today: TFlatCheckBox;
    apn_Network: TAdvPanel;
    GroupBox2: TGroupBox;
    asg_NwUpdate: TAdvStringGrid;
    asg_Network: TAdvStringGrid;
    fsbt_Detail: TFlatSpeedButton;
    fsbt_Master: TFlatSpeedButton;
    fsbt_Network: TFlatSpeedButton;
    apn_Detail: TAdvPanel;
    asg_Detail: TAdvStringGrid;
    apn_Master: TAdvPanel;
    asg_Master: TAdvStringGrid;
    apn_ComDoc: TAdvPanel;
    asg_RegDoc: TAdvStringGrid;
    lb_SelDoc: TLabel;
    asg_SelDoc: TAdvStringGrid;
    fsbt_ComDoc: TFlatSpeedButton;
    fsbt_Release: TFlatSpeedButton;
    lb_RegDoc: TLabel;
    apn_Release: TAdvPanel;
    asg_Release: TAdvStringGrid;
    fsbt_Upload: TFlatSpeedButton;
    fsbt_Insert: TFlatSpeedButton;
    FlatEdit3: TFlatEdit;
    fmed_RegFrDt: TFlatMaskEdit;
    fmed_RegToDt: TFlatMaskEdit;
    Label2: TLabel;
    fcb_DutySpec: TFlatComboBox;
    fed_Release: TFlatEdit;
    fcb_ReScanRelease: TFlatCheckBox;
    Label10: TLabel;
    Shape2: TShape;
    Label11: TLabel;
    Shape1: TShape;
    Shape3: TShape;
    Label12: TLabel;
    Shape4: TShape;
    Label13: TLabel;
    asg_Congrats: TAdvStringGrid;
    fsbt_Duty: TFlatSpeedButton;
    apn_Duty: TAdvPanel;
    Label14: TLabel;
    asg_Duty: TAdvStringGrid;
    Panel1: TPanel;
    Label19: TLabel;
    ProgressBar1: TProgressBar;
    FlatEdit7: TFlatEdit;
    fmed_DutyFrDt: TFlatMaskEdit;
    fmed_DutyToDt: TFlatMaskEdit;
    FlatComboBox1: TFlatComboBox;
    fed_Duty: TFlatEdit;
    fcb_DutyList: TFlatCheckBox;
    apn_DutyList: TAdvPanel;
    asg_DutyList: TAdvStringGrid;
    FlatPanel1: TFlatPanel;
    Label15: TLabel;
    fsbt_UpdDuty: TFlatSpeedButton;
    apn_DocSpec: TAdvPanel;
    asg_DocSpec: TAdvStringGrid;
    apn_AnalList: TAdvPanel;
    asg_AnalList: TAdvStringGrid;
    apn_Weekly: TAdvPanel;
    asg_WeeklyRpt: TAdvStringGrid;
    fsbt_DBMaster: TFlatSpeedButton;
    apn_DBMaster: TAdvPanel;
    asg_DBMaster: TAdvStringGrid;
    fcb_DutyPart: TFlatComboBox;
    asg_DBDetail: TAdvStringGrid;
    pm_Work: TPopupMenu;
    mi_InsWorkRpt: TMenuItem;
    apn_MultiIns: TAdvPanel;
    asg_MultiIns: TAdvStringGrid;
    pm_Doc: TPopupMenu;
    mi_SystemPrint: TMenuItem;
    pm_Master: TPopupMenu;
    mi_NickNm: TMenuItem;
    mi_Chat: TMenuItem;
    mi_SMS: TMenuItem;
    apn_Profile: TAdvPanel;
    fsbt_Profile2Grid: TFlatSpeedButton;
    Label4: TLabel;
    lb_UserNm: TLabel;
    Label5: TLabel;
    fed_NickNm: TFlatEdit;
    apn_SMS: TAdvPanel;
    asg_RecvList: TAdvStringGrid;
    fm_SMS: TFlatMemo;
    fsbt_SendMsg: TFlatSpeedButton;
    mi_Excel: TMenuItem;
    mi_GridPrint: TMenuItem;
    apn_UserProfile: TAdvPanel;
    asg_UserProfile: TAdvStringGrid;
    HTMLPopup: THTMLPopup;
    tm_DialPop: TTimer;
    fpn_WorkConn: TFlatPanel;
    Label16: TLabel;
    Label17: TLabel;
    fsbt_ConnWrite: TFlatSpeedButton;
    lb_WorkConn: TLabel;
    FlatEdit8: TFlatEdit;
    fmed_ConnFrom: TFlatMaskEdit;
    fmed_ConnTo: TFlatMaskEdit;
    FlatEdit9: TFlatEdit;
    asg_WorkConn: TAdvStringGrid;
    FlatComboBox3: TFlatComboBox;
    FlatCheckBox1: TFlatCheckBox;
    mi_Delete: TMenuItem;
    mi_Modify: TMenuItem;
    mi_SystemAllPrint: TMenuItem;
    mi_CRADelUp: TMenuItem;
    mi_CopyDocInfo: TMenuItem;
    pn_ChatUserIp: TPanel;
    fsbt_DialChat: TFlatSpeedButton;
    fsbt_LotUpload: TFlatSpeedButton;
    fmed_LotFrDt: TFlatMaskEdit;
    Label18: TLabel;
    fmed_LotToDt: TFlatMaskEdit;
    pn_Loading: TPanel;
    lb_Message: TLabel;
    pb_DataLoading: TProgressBar;
    fsbt_LotRefresh: TFlatSpeedButton;
    lb_InformMessage: TLabel;
    asg_NotiImg: TAdvStringGrid;
    asg_Like: TAdvStringGrid;
    pn_ChatSearchKey: TPanel;
    mi_RlzScript: TMenuItem;
    apn_Memo: TAdvPanel;
    mm_Log: TMemo;
    asg_Memo: TAdvStringGrid;
    fcb_DutyMake: TFlatCheckBox;
    fsbt_DutyInsert: TFlatSpeedButton;
    fpn_DutyRpt: TFlatPanel;
    fsbt_DutyRptCancel: TFlatSpeedButton;
    fsbt_DutyRptInsert: TFlatSpeedButton;
    Label20: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    lb_DutyRpt: TLabel;
    FlatSpeedButton7: TFlatSpeedButton;
    Label24: TLabel;
    FlatSpeedButton16: TFlatSpeedButton;
    FlatSpeedButton17: TFlatSpeedButton;
    FlatComboBox2: TFlatComboBox;
    FlatMemo1: TFlatMemo;
    FlatEdit10: TFlatEdit;
    FlatEdit11: TFlatEdit;
    fed_DutyDate: TFlatEdit;
    FlatEdit14: TFlatEdit;
    FlatEdit15: TFlatEdit;
    FlatEdit18: TFlatEdit;
    FlatComboBox4: TFlatComboBox;
    FlatEdit19: TFlatEdit;
    FlatEdit20: TFlatEdit;
    FlatEdit23: TFlatEdit;
    FlatEdit24: TFlatEdit;
    FlatEdit25: TFlatEdit;
    FlatEdit22: TFlatEdit;
    FlatEdit26: TFlatEdit;
    FlatEdit16: TFlatEdit;
    fed_DutyUser: TFlatEdit;
    FlatEdit27: TFlatEdit;
    FlatEdit28: TFlatEdit;
    FlatEdit13: TFlatEdit;
    FlatComboBox5: TFlatComboBox;
    AdvStringGrid2: TAdvStringGrid;
    mi_BPL2Delphi: TMenuItem;
    pm_SmsAutoMsg: TPopupMenu;
    mi_Sms_Template: TMenuItem;
    mi_Nw2Excel: TMenuItem;
    mi_DelRelease: TMenuItem;
    procedure FormShow(Sender: TObject);
    procedure asg_Master_ButtonClick(Sender: TObject; ACol, ARow: Integer);
    procedure asg_Master_GetEditorType(Sender: TObject; ACol, ARow: Integer;
      var AEditor: TEditorType);
    procedure pc_DialogChange(Sender: TObject);
    procedure asg_Master_DblClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure asg_Master_CanEditCell(Sender: TObject; ARow, ACol: Integer;
      var CanEdit: Boolean);
    procedure asg_Master_ClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure bbt_MasterPrintClick(Sender: TObject);
    procedure asg_Master_PrintPage(Sender: TObject; Canvas: TCanvas; PageNr,
      PageXSize, PageYSize: Integer);
    procedure asg_Master_PrintStart(Sender: TObject; NrOfPages: Integer;
      var FromPage, ToPage: Integer);
    procedure asg_Detail_ButtonClick(Sender: TObject; ACol, ARow: Integer);
    procedure asg_Detail_CanEditCell(Sender: TObject; ARow, ACol: Integer;
      var CanEdit: Boolean);
    procedure asg_Detail_DblClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure asg_Detail_GetEditorType(Sender: TObject; ACol, ARow: Integer;
      var AEditor: TEditorType);
    procedure asg_Detail_ClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure asg_Detail_EditingDone(Sender: TObject);
    procedure bbt_DetailPrintClick(Sender: TObject);
    procedure asg_Detail_PrintPage(Sender: TObject; Canvas: TCanvas; PageNr,
      PageXSize, PageYSize: Integer);
    procedure asg_Detail_PrintStart(Sender: TObject; NrOfPages: Integer;
      var FromPage, ToPage: Integer);
    procedure bbt_MasterRefreshClick(Sender: TObject);
    procedure bbt_DetailRefreshClick(Sender: TObject);
    procedure bbt_NetworkRefreshClick(Sender: TObject);
    procedure bbt_NetworkPrintClick(Sender: TObject);
    procedure asg_Network_PrintPage(Sender: TObject; Canvas: TCanvas;
      PageNr, PageXSize, PageYSize: Integer);
    procedure asg_Network_PrintStart(Sender: TObject; NrOfPages: Integer;
      var FromPage, ToPage: Integer);
    procedure BitBtn1Click(Sender: TObject);
    procedure asg_NwUpdate_GetAlignment(Sender: TObject; ARow,
      ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asg_Network_GetAlignment(Sender: TObject; ARow, ACol: Integer;
      var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asg_Detail_GetAlignment(Sender: TObject; ARow, ACol: Integer;
      var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asg_Master_GetAlignment(Sender: TObject; ARow, ACol: Integer;
      var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asg_RegDoc_GetEditorType(Sender: TObject; ACol, ARow: Integer;
      var AEditor: TEditorType);
    procedure asg_RegDoc_EditingDone(Sender: TObject);
    procedure asg_RegDoc_ButtonClick(Sender: TObject; ACol, ARow: Integer);
    procedure asg_RegDoc_DblClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure FormDestroy(Sender: TObject);
    procedure asg_RegDoc_ClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure asg_Board_DblClick(Sender: TObject);
    procedure ftc_DialogTabChanged(Sender: TObject);
    procedure fsbt_WriteClick(Sender: TObject);
    procedure fsbt_WriteCancelClick(Sender: TObject);
    procedure fsbt_WriteRegClick(Sender: TObject);
    procedure fsbt_ReplyClick(Sender: TObject);
    procedure fsbt_DeleteClick(Sender: TObject);
    procedure asg_BoardClick(Sender: TObject);
    procedure asg_BoardGetAlignment(Sender: TObject; ARow, ACol: Integer;
      var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asg_BoardButtonClick(Sender: TObject; ACol, ARow: Integer);
    procedure fsbt_BoardRefreshClick(Sender: TObject);
    procedure fcb_CateUpExit(Sender: TObject);
    procedure fcb_PageChange(Sender: TObject);
    procedure fsbt_ForWardClick(Sender: TObject);
    procedure fsbt_BackWardClick(Sender: TObject);
    procedure asg_NetworkGetCellPrintColor(Sender: TObject; ARow,
      ACol: Integer; AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
    procedure asg_DetailGetCellPrintColor(Sender: TObject; ARow,
      ACol: Integer; AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
    procedure asg_MasterGetCellPrintColor(Sender: TObject; ARow,
      ACol: Integer; AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
    procedure fsbt_VersionClick(Sender: TObject);
    procedure fsbt_SearchClick(Sender: TObject);
    procedure fsbt_WriteFindClick(Sender: TObject);
    procedure fmm_TextExit(Sender: TObject);
    procedure fed_WriterKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure fcb_CateDownExit(Sender: TObject);
    procedure asg_BoardMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure fsbt_FileOpenClick(Sender: TObject);
    procedure fsbt_FileClearClick(Sender: TObject);
    procedure asg_MasterEditingDone(Sender: TObject);
    procedure mi_NickNmClick(Sender: TObject);
    procedure fsbt_Profile2GridClick(Sender: TObject);
    procedure asg_BoardClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure mi_ChatClick(Sender: TObject);
    {procedure tm_ChatTimer(Sender: TObject);}
    procedure ServerSocketClientConnect(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ServerSocketClientDisconnect(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ServerSocketClientRead(Sender: TObject;
      Socket: TCustomWinSocket);
    procedure ButtonSendClick(Sender: TObject);
    procedure EditSayKeyPress(Sender: TObject; var Key: Char);
    procedure asg_ChatSendButtonClick(Sender: TObject; ACol,
      ARow: Integer);
    procedure asg_ChatSendKeyPress(Sender: TObject; var Key: Char);
    procedure asg_RegDocKeyPress(Sender: TObject; var Key: Char);
    procedure pm_MasterPopup(Sender: TObject);

    procedure ServerSocketClientError(Sender: TObject;
      Socket: TCustomWinSocket; ErrorEvent: TErrorEvent;
      var ErrorCode: Integer);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure asg_ChatButtonClick(Sender: TObject; ACol, ARow: Integer);
    procedure mi_FileSendClick(Sender: TObject);
    procedure fmm_TextKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure tm_MasterTimer(Sender: TObject);
    procedure mi_EmotiClick(Sender: TObject);
    procedure tm_TxInitTimer(Sender: TObject);
    procedure fsbt_MenuClick(Sender: TObject);
    procedure mi_MenuOpenClick(Sender: TObject);
    procedure apn_MenuClose(Sender: TObject);
    procedure asg_MenuBarButtonClick(Sender: TObject; ACol, ARow: Integer);
    procedure FlatSpeedButton8Click(Sender: TObject);
    procedure FlatSpeedButton9Click(Sender: TObject);
    procedure asg_DialScanKeyPress(Sender: TObject; var Key: Char);
    procedure asg_DialScanKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure asg_DialScanClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure fed_ScanKeyPress(Sender: TObject; var Key: Char);
    procedure fed_ScanClick(Sender: TObject);
    procedure fed_ScanChange(Sender: TObject);
    procedure FlatSpeedButton12Click(Sender: TObject);
    procedure FlatSpeedButton13Click(Sender: TObject);
    procedure FlatSpeedButton14Click(Sender: TObject);
    procedure asg_DialListGetAlignment(Sender: TObject; ARow,
      ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asg_MyDialGetAlignment(Sender: TObject; ARow, ACol: Integer;
      var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asg_MyDialEditingDone(Sender: TObject);
    procedure asg_MyDialButtonClick(Sender: TObject; ACol, ARow: Integer);
    procedure asg_MyDialCanEditCell(Sender: TObject; ARow, ACol: Integer;
      var CanEdit: Boolean);
    procedure asg_MyDialKeyPress(Sender: TObject; var Key: Char);
    procedure asg_DialListClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure asg_LinkListGetAlignment(Sender: TObject; ARow,
      ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asg_DialListDblClick(Sender: TObject);
    procedure asg_DialScanGetAlignment(Sender: TObject; ARow,
      ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asg_DialScanDblClick(Sender: TObject);
    procedure fed_AnalysisClick(Sender: TObject);
    procedure fed_AnalysisKeyPress(Sender: TObject; var Key: Char);
    procedure asg_AnalysisGetAlignment(Sender: TObject; ARow,
      ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asg_AnalysisDblClick(Sender: TObject);
    procedure asg_AnalysisClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure asg_AnalListGetAlignment(Sender: TObject; ARow,
      ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asg_AnalListDblClick(Sender: TObject);
    procedure asg_AnalListKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure asg_AnalListMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure fsbt_WorkRptClick(Sender: TObject);
    procedure fsbt_AnalysisClick(Sender: TObject);
    procedure fmed_AnalFromChange(Sender: TObject);
    procedure fmed_AnalToChange(Sender: TObject);
    procedure asg_WorkRptGetAlignment(Sender: TObject; ARow, ACol: Integer;
      var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asg_WorkRptKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure fcb_WorkAreaChange(Sender: TObject);
    procedure fcb_WorkGubunChange(Sender: TObject);
    procedure fsbt_WeeklyClick(Sender: TObject);
    procedure asg_WeeklyRptKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure asg_WeeklyRptGetAlignment(Sender: TObject; ARow,
      ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asg_SelDocGetAlignment(Sender: TObject; ARow, ACol: Integer;
      var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asg_SelDocKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure asg_WorkRptPrintPage(Sender: TObject; Canvas: TCanvas;
      PageNr, PageXSize, PageYSize: Integer);
    procedure asg_WorkRptPrintStart(Sender: TObject; NrOfPages: Integer;
      var FromPage, ToPage: Integer);
    procedure fsbt_PrintClick(Sender: TObject);
    procedure asg_WeeklyRptPrintPage(Sender: TObject; Canvas: TCanvas;
      PageNr, PageXSize, PageYSize: Integer);
    procedure asg_WeeklyRptPrintStart(Sender: TObject; NrOfPages: Integer;
      var FromPage, ToPage: Integer);
    procedure fsbt_DialMapClick(Sender: TObject);
    procedure asg_DialMapGetAlignment(Sender: TObject; ARow, ACol: Integer;
      var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asg_AnalysisKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure asg_MasterMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure fsbt_AddMyDialClick(Sender: TObject);
    procedure asg_AddMyDialGetAlignment(Sender: TObject; ARow,
      ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asg_AddMyDialKeyPress(Sender: TObject; var Key: Char);
    procedure asg_AddMyDialEditingDone(Sender: TObject);
    procedure asg_AddMyDialButtonClick(Sender: TObject; ACol,
      ARow: Integer);
    procedure asg_MyDialClick(Sender: TObject);
    procedure asg_NetworkMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure asg_UserProfileGetAlignment(Sender: TObject; ARow,
      ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asg_DocSpecGetAlignment(Sender: TObject; ARow, ACol: Integer;
      var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asg_DocSpecKeyPress(Sender: TObject; var Key: Char);
    procedure asg_DocSpecEditingDone(Sender: TObject);
    procedure asg_DocSpecButtonClick(Sender: TObject; ACol, ARow: Integer);
    procedure FlatSpeedButton15Click(Sender: TObject);
    procedure fsbt_DetailClick(Sender: TObject);
    procedure fsbt_NetworkClick(Sender: TObject);
    procedure fsbt_MasterClick(Sender: TObject);
    procedure fsbt_ComDocClick(Sender: TObject);
    procedure fsbt_ReleaseClick(Sender: TObject);
    procedure fsbt_UploadClick(Sender: TObject);
    procedure fsbt_InsertClick(Sender: TObject);
    procedure asg_ReleaseGetAlignment(Sender: TObject; ARow, ACol: Integer;
      var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure fed_ReleaseClick(Sender: TObject);
    procedure fed_ReleaseKeyPress(Sender: TObject; var Key: Char);
    procedure fcb_DutySpecChange(Sender: TObject);
    procedure asg_ReleaseKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure asg_CongratsGetAlignment(Sender: TObject; ARow,
      ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asg_CongratsButtonClick(Sender: TObject; ACol,
      ARow: Integer);
    procedure fsbt_DutyClick(Sender: TObject);
    procedure fmed_DutyToDtChange(Sender: TObject);
    procedure asg_DutyGetAlignment(Sender: TObject; ARow, ACol: Integer;
      var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure fcb_DutyListClick(Sender: TObject);
    procedure asg_DutyListGetAlignment(Sender: TObject; ARow,
      ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asg_DutyListCanEditCell(Sender: TObject; ARow, ACol: Integer;
      var CanEdit: Boolean);
    procedure fsbt_UpdDutyClick(Sender: TObject);
    procedure asg_ReleaseDblClickCell(Sender: TObject; ARow,
      ACol: Integer);
    procedure asg_WeeklyRptDblClickCell(Sender: TObject; ARow,
      ACol: Integer);
    procedure fmed_DutyFrDtChange(Sender: TObject);
    procedure fed_DutyClick(Sender: TObject);
    procedure fed_DutyKeyPress(Sender: TObject; var Key: Char);
    procedure asg_BoardDblClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure asg_UserProfileButtonClick(Sender: TObject; ACol,
      ARow: Integer);
    procedure fsbt_DBMasterClick(Sender: TObject);
    procedure asg_DBMasterClick(Sender: TObject);
    procedure fcb_DutyPartChange(Sender: TObject);
    procedure asg_DBMasterKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure asg_DBDetailKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure asg_UserProfileDblClickCell(Sender: TObject; ARow,
      ACol: Integer);
    procedure mi_InsWorkRptClick(Sender: TObject);
    procedure asg_MultiInsGetEditorType(Sender: TObject; ACol,
      ARow: Integer; var AEditor: TEditorType);
    procedure asg_MultiInsEditingDone(Sender: TObject);
    procedure asg_MultiInsCanEditCell(Sender: TObject; ARow, ACol: Integer;
      var CanEdit: Boolean);
    procedure asg_MultiInsButtonClick(Sender: TObject; ACol,
      ARow: Integer);
    procedure asg_MultiInsGetAlignment(Sender: TObject; ARow,
      ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure mi_SystemPrintClick(Sender: TObject);
    procedure pm_DocPopup(Sender: TObject);
    procedure mi_SMSClick(Sender: TObject);
    procedure asg_NetworkMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure asg_NetworkMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure fm_SMSClick(Sender: TObject);
    procedure fsbt_SendMsgClick(Sender: TObject);
    procedure asg_RecvListGetAlignment(Sender: TObject; ARow,
      ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure mi_ExcelClick(Sender: TObject);
    procedure pm_WorkPopup(Sender: TObject);
    procedure mi_GridPrintClick(Sender: TObject);
    procedure asg_DialMapPrintPage(Sender: TObject; Canvas: TCanvas;
      PageNr, PageXSize, PageYSize: Integer);
    procedure asg_DialMapPrintStart(Sender: TObject; NrOfPages: Integer;
      var FromPage, ToPage: Integer);
    procedure asg_DialMapPrintSetColumnWidth(Sender: TObject;
      ACol: Integer; var Width: Integer);
    procedure asg_DialMapPrintSetRowHeight(Sender: TObject; ARow: Integer;
      var Height: Integer);
    procedure asg_DialListMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure asg_DialListMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure asg_MyDialMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure asg_MyDialMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure FormPaint(Sender: TObject);
    procedure asg_DialListMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure asg_AnalysisMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure FormCreate(Sender: TObject);
    procedure HTMLPopupAnchorClick(Sender: TObject; Anchor: String);
    procedure tm_DialPopTimer(Sender: TObject);
    procedure asg_ReleaseClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure asg_SelDocDblClick(Sender: TObject);
    procedure asg_SelDocClick(Sender: TObject);
    procedure asg_MultiInsClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure asg_MultiInsTopLeftChanged(Sender: TObject);
    procedure asg_MultiInsDblClick(Sender: TObject);
    procedure fsbt_ConnWriteClick(Sender: TObject);
    procedure asg_WorkConnGetAlignment(Sender: TObject; ARow,
      ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure fmed_ConnFromChange(Sender: TObject);
    procedure fmed_ConnToChange(Sender: TObject);
    procedure mi_DeleteClick(Sender: TObject);
    procedure mi_ModifyClick(Sender: TObject);
    procedure mi_SystemAllPrintClick(Sender: TObject);
    procedure mi_CRADelUpClick(Sender: TObject);
    procedure mi_CopyDocInfoClick(Sender: TObject);
    procedure asg_BoardKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure fed_TitleKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure fsbt_DialChatClick(Sender: TObject);
    procedure fsbt_LotRefreshClick(Sender: TObject);
    procedure asg_BigDataGetAlignment(Sender: TObject; ARow, ACol: Integer;
      var HAlign: TAlignment; var VAlign: TVAlignment);
    procedure asg_LikeMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure mi_RlzScriptClick(Sender: TObject);
    procedure asg_MemoDblClick(Sender: TObject);
    procedure asg_MemoClickCell(Sender: TObject; ARow, ACol: Integer);
    procedure fcb_DutyMakeClick(Sender: TObject);
    procedure fsbt_DutyInsertClick(Sender: TObject);
    procedure asg_DutyCanEditCell(Sender: TObject; ARow, ACol: Integer;
      var CanEdit: Boolean);
    procedure asg_DutyKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure fsbt_DutyRptCancelClick(Sender: TObject);
    procedure asg_DutyDblClick(Sender: TObject);
    procedure mi_BPL2DelphiClick(Sender: TObject);
    procedure fmed_RegFrDtKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure fmed_RegToDtKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure fmed_AnalFromKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure fmed_AnalToKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure mi_Sms_TemplateClick(Sender: TObject);
    procedure mi_DelReleaseClick(Sender: TObject);

  private
    { Private declarations }
    FsUserIp, FsUserNm, FsNickNm, FsUserPart, FsUserSpec, FsUserMobile, FsUserCallNo, FsPopMsg, FsVersion, FsMngrNm : String;
    FsChatMyPhoto : String;
    iNowBoardCnt : Integer;
    iChatRowCnt  : Integer;
    IsLogonUser  : Boolean;
//    isChatOn     : Boolean;
    Reciving     : Boolean;
    DataSize     : integer;
    Data         : TMemoryStream;
    IsFileSend   : Boolean;
    IsEmotiSend  : Boolean;
    DownFile     : Variant;
    DownCnt      : Integer;
    iSelRow1, iSelRow2 : Integer;
    iUserRowId   : Integer;


   //----------------------------------------------------------------
   // Variant Array 마지막 자리에 빈칸 하나 삽입
   //    - 소스출처 : CComFunc.pas
   //----------------------------------------------------------------
   function AppendVariant(var in_var : Variant; in_str : String ) : Integer;


   //----------------------------------------------------------------
   // String 내의 특정 문자열 삭제
   //    - 소스출처 : MComFunc.pas
   //----------------------------------------------------------------
   function DeleteStr(OrigStr, DelStr : String) : String;



   //----------------------------------------------------------------
   // TAdvStringGrid 프린트 속성 정의
   //    - 2013.08.08 LSH
   //----------------------------------------------------------------
   procedure SetPrintOptions;


   //---------------------------------------------------------------------------
   // [조회] Welcome 메세지 정보 가져오기
   //    - 2013.08.08 LSH
   //---------------------------------------------------------------------------
   function GetMsgInfo(in_ComCd2, in_ComCd3 : String) : String;


   //---------------------------------------------------------------------------
   // [조회] 공지사항 팝업메세지 정보 가져오기
   //    - 2013.08.22 LSH
   //---------------------------------------------------------------------------
   function GetAlarmPopup(in_Gubun, in_CheckDate : String) : String;


   //---------------------------------------------------------------------------
   // [조회] 각 문서별 Max Seqno 정보
   //    - 2013.08.13 LSH
   //---------------------------------------------------------------------------
   function GetMaxDocSeq(in_DocList, in_DocYear : String) : String;


   //---------------------------------------------------------------------------
   // [조회] 사용자 IP 별 정보
   //    - 2013.08.13 LSH
   //---------------------------------------------------------------------------
   function CheckUserInfo(in_UserIp : String; var varUserNm,
                                                  varNickNm,
                                                  varUserPart,
                                                  varUserSpec,
                                                  varUserMobile,
                                                  varUserCallNo,
                                                  varMngrNm : String) : Boolean;


   //---------------------------------------------------------------------------
   // [조회] 요일 정보
   //    - 소스출처 : MComFunc.pas
   //---------------------------------------------------------------------------
   function GetDayofWeek(Adate : TDatetime; Type1 : String): String;


   //---------------------------------------------------------------------------
   // [업데이트] 게시판 (커뮤니티) Update
   //    - 2013.08.19 LSH
   //---------------------------------------------------------------------------
   procedure UpdateBoard(in_Type,     in_BoardSeq, in_UserIp,    in_RegDate, in_RegUser,
                         in_CateUp,   in_CateDown, in_Title,     in_Context, in_AttachNm,
                         in_HideFile, in_ServerIp, in_HeadTail,  in_HeadSeq, in_TailSeq,
                         in_AlarmFro, in_AlarmTo,  in_DelDate,   in_EditIp : String);


   //---------------------------------------------------------------------------
   // [조회] 게시판 (커뮤니티) 댓글 상세조회
   //    - 2013.08.20 LSH
   //---------------------------------------------------------------------------
   procedure SetReplyList(in_HeaderRow : Integer;
                          in_BoardSeq,
                          in_UserIp,
                          in_RegDate : String);


   //---------------------------------------------------------------------------
   // [업데이트] Log 업데이트 (게시글 추천 등)
   //    - 2013.08.20 LSH
   //---------------------------------------------------------------------------
   procedure UpdateLog(in_Gubun,
                       in_BoardSeq,
                       in_UserIp,
                       in_LogFlag,
                       in_Item1,
                       in_Item2  : String;
                       var varResult : String);


   //---------------------------------------------------------------------------
   // 한영키를 한글키로 변환
   //    - 소스출처 : MComFunc.pas
   //---------------------------------------------------------------------------
   procedure HanKeyChg(Handle1:THandle);



   //---------------------------------------------------------------------------
   // [조회] 게시판 (커뮤니티) 추천 List 상세조회
   //    - 2013.08.27 LSH
   //---------------------------------------------------------------------------
   procedure GetLikeList(in_HeaderRow : Integer;
                         in_BoardSeq,
                         in_UserIp,
                         in_RegDate : String);




   // Client 메세지 전송
   function SendToAllClients( s : String) : boolean;



   //---------------------------------------------------------------------------
   // FTP 업로드
   //    - 소스출처 : MDN191F2.pas
   //---------------------------------------------------------------------------
   function FileUpLoad(TargetFile, DelFile: String; var S_IP: String): Boolean;



   //---------------------------------------------------------------------------
   // [함수] TCP Port 오픈여부 Check
   //    - 2013.08.30 LSH
   //    - 소스출처 : Google
   //---------------------------------------------------------------------------
   function PortTCPIsOpen(dwPort : Word; ipAddressStr:string) : boolean;


   //---------------------------------------------------------------------------
   // Data Loading Bar Controller
   //    - 2013.09.06 LSH
   //---------------------------------------------------------------------------
   procedure SetLoadingBar(AsStatus : String);


   //----------------------------------------------------------------------
   // 엑셀 파일 Grid에 불러오기
   //    - 소스출처 : dkChoi
   //    - 2013.09.13 LSH
   //----------------------------------------------------------------------
   procedure LoadExcelFile(in_FileName : String);



   //---------------------------------------------------------------------------
   // [조회] 각 근무처별 직원식 정보
   //    - 2013.09.16 LSH
   //---------------------------------------------------------------------------
   function GetDietInfo(in_Gubun, in_DayGubun : String) : String;


   //---------------------------------------------------------------------------
   // [업데이트] 나의 다이얼 Book Update
   //    - 2013.09.26 LSH
   //---------------------------------------------------------------------------
   procedure UpdateMyDial(in_Type,     in_Locate,     in_UserIp,  in_DeptNm,  in_DeptSpec,
                          in_CallNo,   in_DutyUser,   in_Mobile,  in_DutyRmk, in_DataBase,
                          in_LinkSeq,  in_DelDate,    in_DtyUsrId : String);



   //---------------------------------------------------------------------------
   // [조회] 다이얼 Book 링크 List 상세조회
   //    - 2013.09.27 LSH
   //---------------------------------------------------------------------------
   procedure GetLinkList(in_Gubun,
                         in_DialLocate,
                         in_DialLinkDb,
                         in_DialLinkSeq,
                         in_UserIp : String);

   //---------------------------------------------------------------------------
   // User별 맞춤형 알람 Check
   //    - 2013.12.06 LSH
   //---------------------------------------------------------------------------
   function CheckUserAlarm(in_Gubun,
                           in_UserNm,
                           in_UserIp,
                           in_Option : String) : Boolean;

   //---------------------------------------------------------------------------
   // Panel 상태 변경하기
   //    - 2014.01.29 LSH
   //---------------------------------------------------------------------------
   procedure SetPanelStatus(in_Gubun,
                            in_Status : String);

   //---------------------------------------------------------------------------
   // [조회] FlatComboBox 공통 Code 조회
   //    - 2014.02.03 LSH
   //---------------------------------------------------------------------------
   procedure GetComboBoxList(in_Type1,
                             in_Type2 : String;
                             NmCombo  : TFlatComboBox);

   //---------------------------------------------------------------------------
   // [등록] 구분별 첨부 이미지 등록
   //    - 2014.02.06 LSH
   //---------------------------------------------------------------------------
   procedure UpdateImage(in_Gubun,
                         in_FileName,
                         in_HideFile : String;
                         in_NowRow : Integer);

   //---------------------------------------------------------------------------
   // [엑셀] Grid - Excel 변환
   //    - 2014.04.07 LSH
   //---------------------------------------------------------------------------
   procedure ExcelSave(ExcelGrid : TAdvStringGrid; Title : String; Disp : Integer);


   //---------------------------------------------------------------------------
   // 다이얼 Pop 메세지 실시간 Check
   //    - 2014.07.01 LSH
   //---------------------------------------------------------------------------
   function CheckDialPop(in_Gubun : String) : String;



  public
    //--------------------------------------------------------------------------
    // [조회] 동일 IP 다중 User 선택 정보 가져오기
    //    - 2013.08.29 LSH
    //--------------------------------------------------------------------------
    procedure GetMultiUser(in_UserNm,
                           in_NickNm,
                           in_UserPart,
                           in_UserSpec,
                           in_Mobile,
                           in_CallNo,
                           in_MngrNm : String);

    //--------------------------------------------------------------------------
    // String 내의 특정 문자 삭제
    //    - 소스출처 : MComFunc.pas
    //--------------------------------------------------------------------------
    function DelChar( const Str : String ; DelC : Char ) : String;


    //--------------------------------------------------------------------------
    // [조회] Grid 조건별 상세 procedure
    //    - 2013.08.08 LSH
    //--------------------------------------------------------------------------
    procedure SelGridInfo(in_SearchFlag : String);


    //--------------------------------------------------------------------------
    // 개발관련 문서(릴리즈대장) 엑셀 업로드 일괄등록
    //    - 2014.01.02 LSH
    //--------------------------------------------------------------------------
    procedure InsReleaseList;

    //--------------------------------------------------------------------------
    // 빅데이터 Lab. 엑셀 업로드 일괄등록
    //    - 2015.04.02 LSH
    //--------------------------------------------------------------------------
    procedure InsBigDataList;

    //--------------------------------------------------------------------------
    // Grid 컬럼 정보 변경
    //    - 2014.02.03 LSH
    //--------------------------------------------------------------------------
    procedure SetGridCol(in_Flag : String);

    //--------------------------------------------------------------------------
    // [통합출력] FTP 이미지 통합 출력 함수
    //--------------------------------------------------------------------------
    procedure KDialPrint(in_ImgCode  : String;
                         in_GridName : TAdvStringGrid);

    //--------------------------------------------------------------------------
    // [다이얼Pop] Pop-up 메세지 Anchor 이벤트
    //     - 2014.07.01 LSH
    //--------------------------------------------------------------------------
    procedure AnchorClick(sender: TObject; Anchor: string);

    //--------------------------------------------------------------------------
    // [다이얼Pop] Pop-up 메세지 팝업 이벤트
    //     - 2014.07.01 LSH
    //--------------------------------------------------------------------------
    function CreatePopup(in_Gubun : String) : Boolean;


    //--------------------------------------------------------------------------
    // 문자열 처음부터 KeyStr이 나오기 직전까지를 잘라냄 @ 2015.03.31 LSH
    //      - 소스출처 : MComfunc.pas
    //--------------------------------------------------------------------------
    function FindStr(Fstr : String; KeyStr : String) : String;

    //--------------------------------------------------------------------------
    // String내의 특정 문자열을 다른 문자열로 대체 @ 2015.04.02 LSH
    //      - 소스출처 : MComfunc.pas
    //      - StringLib.pas내 ReplaceStr과 충돌나서 별도로 추가함.
    //--------------------------------------------------------------------------
    function ReplaceChar(TextStr, OrigStr, ChgStr : String) : String;

    //--------------------------------------------------------------------------
    // [근무 및 근태] 근무-근태-당직일지 변경사항 Update
    //      - 2015.06.12 LSH
    //--------------------------------------------------------------------------
    procedure UpdateDuty(in_UpdFlag : String);

    //--------------------------------------------------------------------------
    // [SMS] SMS 템플릿 목록 D/B 조회
    //      - 2017.07.11 LSH
    //--------------------------------------------------------------------------
    procedure GetSMSTemplate;

    //--------------------------------------------------------------------------
    // 개발관련 문서(릴리즈대장) 이력 삭제 (Deldate Updated)
    //    - 2018.06.12 LSH
    //--------------------------------------------------------------------------
    procedure UpdReleaseDeldt;

  published
      // Chat - Client에서 Log 업데이트 위한 Call-back 함수 @ 2015.03.27 LSH
      procedure SetUpdLogOut(Sender : TObject);
      procedure SetUpdChatLog(Sender : TObject);
      procedure SetUpdChatSearch(Sender : TObject);

      // 사용자 프로필 이미지 동적 FTP 다운로드 @ 2015.03.31 LSH
      function GetFTPImage(in_PhotoFile, in_HiddenFile, in_UserId : String) : String;

  end;

var
  MainDlg: TMainDlg;




implementation

{$R *.DFM}


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
   //AsPrev,  // AdvStrGrid 출력 미리보기 모듈
   TpSvc,   // TpCall Agent 사용관련
   Excel2000,
   Imm,     // HIMC 사용관련추가 @ 2013.08.22 LSH
   ShellApi,// ShellExec 사용관련 추가 @ 2013.08.22 LSH
   Winsock, // Windows Sockets API Unit
   ComObj,
   Qrctrls,
   QuickRpt,
   IdFTPCommon   // 델파이 XE 컨버젼관련 Id_FTP --> IndyFTP 변환 @ 2014.04.07 LSH
   ;




//------------------------------------------------------------------------------
// [FormShow] TForm onShow Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2013.08.07
//------------------------------------------------------------------------------
procedure TMainDlg.FormShow(Sender: TObject);
var
   sUserJikGup, sWcMsg, sVerMsg, sInformMsg, sDietMsg, sTotalMsg, sNewHospital : string;
   varResult{, sTargetFile, sTargetTitle }: String;
   //SetServer : TServer;
//   sLocalFile, sRemoteFile : String;
//   sServerIp, sFtpUserID, sFtpPasswd, sFtpHostName, sFtpDir : String;
//   FForm     : TForm;
begin

   sNewHospital := '';

   //----------------------------------------------------------
   // 개발계 (AAOCSDEV) Param. Set.
   //----------------------------------------------------------
   begin
      sNewHospital := 'D0';
//      G_SVRID      := '01';
   end;


   //----------------------------------------------------------
   // 개발계 (또는 운영계, 추후 확장성 고려) 접속
   //----------------------------------------------------------
   if sNewHospital <> '' then
   begin
      if not (TxInit(TokenStr(String(CONST_ENV_FILENAME), '.', 1) + '_'+ sNewHospital +'.ENV', G_SVRID)) then
      begin
         txTerm;
         Close;
      end;
   end
   else
   begin
      if not (TxInit(String(CONST_ENV_FILENAME), G_SVRID)) then
      begin
         txTerm;
         Close;
      end;
   end;



   //---------------------------------------------------------
   // Check TxInit
   //---------------------------------------------------------
   {
   if not txInit(G_ENV_FILENAME, G_SVRID) then
      Application.Terminate;
   }


   //---------------------------------------------------------
   // Get User IP
   //---------------------------------------------------------
   FsUserIp := String(GetIp);




   // test : 다중사용자 IP 선택
   {
   if (FsUserIp = '개발자IP') then
      FsUserIp := '테스트IP';
   }







   //---------------------------------------------------------
   // [최근업데이트] TabSheet Frame Set.
   //---------------------------------------------------------
   with asg_NwUpdate do
   begin
      ColWidths[C_NU_EDITDATE]   := 150;
      ColWidths[C_NU_EDITTEXT]   := 350;
      ColWidths[C_NU_EDITIP]     := 110;
   end;


   //---------------------------------------------------------
   // [비상연락망] TabSheet Frame Set.
   //---------------------------------------------------------
   with asg_NetWork do
   begin
      ColWidths[C_NW_LOCATE]   := 45;
      ColWidths[C_NW_DUTYPART] := 60;
      ColWidths[C_NW_DUTYSPEC] := 85;
      ColWidths[C_NW_USERNM]   := 50;
      ColWidths[C_NW_CALLNO]   := 45;
      ColWidths[C_NW_MOBILE]   := 85;
      ColWidths[C_NW_DUTYRMK]  := 125;
      ColWidths[C_NW_EMAIL]    := 125;
//      ColWidths[C_NW_USERIP]    := 0;
//      ColWidths[C_NW_PHOTOFILE] := 0;
   end;



   //---------------------------------------------------------
   // [업무프로필] TabSheet Frame Set.
   //---------------------------------------------------------
   with asg_Detail do
   begin
      ColWidths[C_D_LOCATE]   := 45;
      ColWidths[C_D_DUTYPART] := 60;
      ColWidths[C_D_DUTYSPEC] := 80;
      ColWidths[C_D_DUTYRMK]  := 125;
      ColWidths[C_D_DUTYUSER] := 50;
      ColWidths[C_D_CALLNO]   := 45;
      ColWidths[C_D_DUTYPTNR] := 55;
      ColWidths[C_D_DELDATE]  := 65;
      ColWidths[C_D_BUTTON]   := 55;
      ColWidths[C_D_USERADD]  := 40;


      AddComment(C_D_DUTYRMK,
                 0,
                 '구체적으로 담당하는 업무의 상세내역을,'#13'입력하실 수 있습니다.');

      AddComment(C_D_DUTYPTNR,
                 0,
                 '업무구분이 [개발파트]인 경우 필수입력 항목으로,'#13'파트내 P/L 또는 의료원 또는 파트너쉽(협력업체) 여부를 표기합니다.');

      AddComment(C_D_DELDATE,
                 0,
                 '비상연락망에서 선택한 담당자의 업무프로필 정보를 삭제하기 원하시면,'#13'해당 종료일자를 더블클릭으로 넣으신후 [Update] 해주세요.'#13'(더블클릭시 종료일자 Toggle)');

      AddComment(C_D_USERADD,
                 0,
                 '신규 업무프로필을 추가하기 원하시면 [+]를 눌러주세요.');
   end;



   //---------------------------------------------------------
   // [담당자정보] TabSheet Frame Set.
   //---------------------------------------------------------
   with asg_Master do
   begin
      ColWidths[C_M_LOCATE]   := 45;
      ColWidths[C_M_USERID]   := 40;
      ColWidths[C_M_USERNM]   := 55;
      ColWidths[C_M_DUTYPART] := 55;
      ColWidths[C_M_MOBILE]   := 80;
      ColWidths[C_M_EMAIL]    := 120;
      ColWidths[C_M_USERIP]   := 65;
      ColWidths[C_M_DELDATE]  := 65;
      ColWidths[C_M_BUTTON]   := 55;
      ColWidths[C_M_USERADD]  := 40;
      ColWidths[C_M_NICKNM]   := 0;

      AddComment(C_M_DELDATE,
                 0,
                 '비상연락망에서 선택한 담당자정보를 삭제하기 원하시면,'#13'해당 종료일자를 더블클릭으로 넣으신후 [Update] 해주세요.'#13'(더블클릭시 종료일자 Toggle)');

      AddComment(C_M_USERADD,
                 0,
                 '신규 담당자를 추가하기 원하시면 [+]를 눌러주세요.');

      AddComment(C_M_USERIP,
                 0,
                 '선택한 Cell을 더블클릭하시면, 현위치의 IP 정보를 자동조회 합니다.');
   end;




   //---------------------------------------------------------
   // [발송/협조문] 검색 및 등록 TabSheet Frame Set.
   //---------------------------------------------------------
   with asg_RegDoc do
   begin
      AddComment(C_RD_DOCYEAR,
                 0,
                 '회계년도 (당년 3월~ 익년 2월)를 [yyyy] 4자리로 입력해주세요.');

      AddComment(C_RD_DOCSEQ,
                 0,
                 '[등록]시 더블클릭 하시면, 문서번호가 자동으로 채번됩니다.'#13'검색/수정/삭제시 직접 문서번호 3자리 입력해주세요.');

      AddComment(C_RD_DOCLOC,
                 0,
                 '실제 오프라인에 보관중(또는 보관예정)인 문서철의 이름을 적어주세요.');
   end;


   //---------------------------------------------------------
   // [발송/협조문] 조회결과 TabSheet Frame Set.
   //---------------------------------------------------------
   with asg_SelDoc do
   begin
      ColWidths[C_SD_HQREMARK]   := 0;
      ColWidths[C_SD_AAREMARK]   := 0;
      ColWidths[C_SD_GRREMARK]   := 0;
      ColWidths[C_SD_ASREMARK]   := 0;
      ColWidths[C_SD_TOTRMK]     := 0;
   end;




   //---------------------------------------------------------
   // [커뮤니티] TabSheet Frame Set.
   //---------------------------------------------------------
   with asg_Board do
   begin
      ColWidths[C_B_HEADSEQ]    := 0;
      ColWidths[C_B_TAILSEQ]    := 0;
      ColWidths[C_B_CLOSEYN]    := 0;
      ColWidths[C_B_CONTEXT]    := 0;
      ColWidths[C_B_USERIP]     := 0;
      ColWidths[C_B_REPLY]      := 0;
      ColWidths[C_B_HEADTAIL]   := 0;
      ColWidths[C_B_HIDEFILE]   := 0;
      ColWidths[C_B_SERVERIP]   := 0;

      AddComment(C_B_LIKE,
                 0,
                 '게시글 당 1회씩 추천이 가능하며,'#13'추천횟수가 5회 이상인 경우 Best 게시글로 7일간 Page 상단에 피드됩니다.');
   end;



   //---------------------------------------------------------
   // [Chat Box] Frame Set.
   //---------------------------------------------------------
   with asg_Chat do
   begin
      ColWidths[C_CH_MYCOL]   := 40;
      ColWidths[C_CH_TEXT]    := 285;
      ColWidths[C_CH_YOURCOL] := 40;
   end;



   //---------------------------------------------------------
   // [Chat 전송] Frame Set.
   //---------------------------------------------------------
   with asg_ChatSend do
   begin
      ColWidths[C_SD_TEXT]   := 318;
//      ColWidths[C_SD_BUTTON  := 227;

      AddButton(C_SD_BUTTON, 0, ColWidths[C_SD_BUTTON]-5, 20, '전송', haBeforeText, vaCenter);     // 전송 Button
   end;


   //---------------------------------------------------------
   // [다이얼 Book 검색] Frame Set.
   //---------------------------------------------------------
   with asg_DialList do
   begin
      ColWidths[C_DL_LINKSEQ]   := 0;
      ColWidths[C_DL_DTYUSRID]  := 0;
   end;


   //---------------------------------------------------------
   // [나의 다이얼 Book] Frame Set.
   //---------------------------------------------------------
   with asg_MyDial do
   begin
      ColWidths[C_MD_LINKDB]    := 0;
      ColWidths[C_MD_LINKSEQ]   := 0;
   end;


   //---------------------------------------------------------
   // [나의 다이얼 Book] Frame Set.
   //---------------------------------------------------------
   with asg_LinkList do
   begin
      ColWidths[C_LL_LINKDB]    := 0;
      ColWidths[C_LL_LINKSEQ]   := 0;
   end;


   //---------------------------------------------------------
   // [업무(S/R) 분석] Frame Set.
   //---------------------------------------------------------
   with asg_Analysis do
   begin
      ColWidths[C_AN_REQNO]     := 0;
      ColWidths[C_AN_REQOBJT]   := 0;
      ColWidths[C_AN_REQRMK]    := 0;
      ColWidths[C_AN_ANSRMK]    := 0;
      ColWidths[C_AN_REQISSU]   := 0;
      ColWidths[C_AN_ATTACHNM]  := 0;
      ColWidths[C_AN_USERID]    := 0;
      ColWidths[C_AN_REQDEPTNM] := 0;
      ColWidths[C_AN_REQTELNO]  := 0;
      ColWidths[C_AN_REQSPEC]   := 0;
   end;


   //---------------------------------------------------------
   // [업무(S/R) 분석 상세] Frame Set.
   //---------------------------------------------------------
   with asg_AnalList do
   begin
      ColWidths[C_AL_ADDINFO]   := 0;
   end;


   //---------------------------------------------------------
   // [업무일지(KUMC)] Frame Set.
   //---------------------------------------------------------
   with asg_WorkRpt do
   begin
      ColWidths[C_WR_CONFIRM]   := 0;
   end;


   //---------------------------------------------------------
   // [비상연락망] 담당자 프로필 Set.
   //---------------------------------------------------------
   with asg_UserProfile do
   begin
      RowHeights[R_PR_PHOTO] := 132;
      RowHeights[R_PR_USERID]:= 0;
   end;

   //---------------------------------------------------------
   // [주요업무공유] Frame Set.
   //---------------------------------------------------------
   {  -- 미사용으로 주석 @ 2017.10.31 LSH
   with asg_WorkConn do
   begin
      ColWidths[C_WC_CONFIRM]   := 0;

      fmed_ConnFrom.Text   := FormatDateTime('yyyy-mm-dd', DATE - 30);   // 약 1개월 전 이력부터 검색
      fmed_ConnTo.Text     := FormatDateTime('yyyy-mm-dd', DATE);
   end;
   }



   //---------------------------------------------------------
   // K-Dialog 사용자 User 정보 조회
   //---------------------------------------------------------
   IsLogonUser := CheckUserInfo(FsUserIp, FsUserNm, FsNickNm, FsUserPart, FsUserSpec, FsUserMobile, FsUserCallNo, FsMngrNm);





   // 사용자 직급 확인
   if (PosByte('실장',   FsUserSpec) > 0) or
      (PosByte('팀장',   FsUserSpec) > 0) or
      (PosByte('파트장', FsUserSpec) > 0) then
      sUserJikGup := FsUserSpec  + '님'
   else
      sUserJikGup := '선생님';


   // Welcome 메세지 정보 조회
   sWcMsg := GetMsgInfo('WELCOME',
                        GetDayofWeek(Date, 'HS'));


   // Version 체크 정보 조회
   sVerMsg := GetMsgInfo('INFORM',
                         'VERSION');


   // Message 정리 #1
   if (sVerMsg <> '') then
      sTotalMsg := sWcMsg + #13#10 + #13#10 + #13#10 + sVerMsg
   else
      sTotalMsg := sWcMsg;

   // 직원식 메뉴 정보 조회
   sDietMsg := GetDietInfo(FormatDateTime('dd', Date),
                           GetDayofWeek(Date, 'HS'));


   // Message 정리 #2
   if (sDietMsg <> '') then
      sTotalMsg := sTotalMsg + #13#10 + #13#10 + #13#10 + sDietMsg
   else
      sTotalMsg := sTotalMsg;



   // Alarm 팝업메세지 정보 조회
   sInformMsg := GetAlarmPopup('POPUP',
                               FormatDateTime('yyyy-mm-dd', DATE));


   // Message 정리 #3
   if (sInformMsg <> '') then
      sTotalMsg := sTotalMsg + #13#10 + #13#10 + #13#10 + sInformMsg
   else
      sTotalMsg := sTotalMsg;


   //-----------------------------------------------------
   // 다이얼로그 Tab Skin Color
   //-----------------------------------------------------
   ftc_Dialog.Color     := clWhite;
   fpn_DialBook.Color   := clWhite;
   fpn_Analysis.Color   := clWhite;
   apn_Network.Color    := clWhite;
   apn_Detail.Color     := clWhite;
   apn_Master.Color     := clWhite;
   apn_Network.Color    := clWhite;


   //-----------------------------------------------------
   // Welcome 메세지 처리
   //-----------------------------------------------------
   if (PosByte('개발자IP', FsUserIp) > 0) then
   begin
      MessageBox(self.Handle,
                 PChar('환영합니다, ' + FsUserNm + ' 개발자님.' + #13#10 + #13#10 + sTotalMsg),
                 '★★ KUMC 다이얼로그 ★★',
                 MB_OK + MB_ICONINFORMATION);

      ftc_Dialog.ActiveTab := AT_DOC;

      fsbt_releaseClick(Sender);
      //fsbt_WorkRptClick(Sender);
      //fsbt_NetworkClick(sender);
      //fsbt_ComDocClick(Sender);
      //fsbt_DBMasterClick(Sender);
      //fsbt_DutyClick(Sender);

      //---------------------------------------------------
      // [관리자] 주간 식단정보(-> Big-Data lab) 입력 기능 On
      //       - 식단안내 기능 전체 오픈 @ 2017.10.26 LSH
      //---------------------------------------------------
      //fsbt_Menu.Visible := True;


   end
   else if (PosByte('실장',   FsUserSpec) > 0) or
           (PosByte('팀장',   FsUserSpec) > 0) or
           (PosByte('파트장', FsUserSpec) > 0) then
   begin
      MessageBox(self.Handle,
                 PChar('환영합니다, ' + FsUserNm + ' ' +  sUserJikGup + '.' + #13#10 + #13#10 + sTotalMsg),
                 'Welcome to KUMC 다이얼로그',
                 MB_OK + MB_ICONINFORMATION);

      ftc_Dialog.ActiveTab := AT_BOARD;
   end
   else if (FsUserNm <> '') and (FsUserPart <> '') then
   begin
      MessageBox(self.Handle,
                 PChar('환영합니다, ' + FsUserNm + ' ' +  sUserJikGup + '.' + #13#10 + #13#10 + sTotalMsg),
                 'Welcome to KUMC 다이얼로그',
                 MB_OK + MB_ICONINFORMATION);


      // 개발파트는 [업무분석], 그 외 파트는 [커뮤니티] 첫 Page.
      // 개발파트도 [게시판]부터 ~ ^^ @ 2014.06.26 LSH
      if PosByte('개발', FsUserPart) > 0 then
      begin
         ftc_Dialog.ActiveTab := AT_BOARD;
      end
      else
         ftc_Dialog.ActiveTab := AT_BOARD;

   end
   else if (FsUserNm <> '') and (FsUserPart = '') then
   begin
      MessageBox(self.Handle,
                 PChar('환영합니다, ' + FsUserNm + ' ' +  sUserJikGup + '.' + #13#10 + #13#10 + sTotalMsg + #13#10 + #13#10 + '[업무프로필]을 업데이트 해주셔야 비상연락망에 자동표기(정렬) 됩니다.' + #13#10 + #13#10 + '작성 Page로 이동합니다.'),
                 'Welcome to KUMC 다이얼로그',
                 MB_OK + MB_ICONINFORMATION);


      ftc_Dialog.ActiveTab := AT_DIALOG;

      fsbt_DetailClick(Sender);

   end
   else
   begin
      MessageBox(self.Handle,
                 PChar('환영합니다!' + #13#10 + #13#10 + sTotalMsg + #13#10 + #13#10 + '비상연락망(또는 사용자 IP) 미등록 중이시네요.' + #13#10 + '[담당자정보] 입력후, [업무프로필] 업데이트 부탁드립니다.'),
                 'Welcome to KUMC 다이얼로그',
                 MB_OK + MB_ICONINFORMATION);


      ftc_Dialog.ActiveTab := AT_DIALOG;

      fsbt_MasterClick(Sender);
   end;



//-----------------------------------------------------
   // 서버소켓 Port 정보 및 Enabled
   //-----------------------------------------------------
   ServerSocket.Port := StrToInt(DelChar(CopyByte(FsUserIp, LengthByte(FsUserIp)-3, 4), '.'));


   try
      ServerSocket.Active := True;

   except
       MessageBox(self.Handle,
                  PChar('접속하신 IP/Port 정보가 이미 사용중입니다.' + #13#10 + #13#10 + '동일한 프로그램을 먼저 실행중이 아닌지 확인해 주세요.'),
                  PChar(Self.Caption + ' : 프로그램 중복실행 오류'),
                  MB_OK + MB_ICONWARNING);
   end;



   //-----------------------------------------------------
   // Caption 사용자 정보 업데이트
   //-----------------------------------------------------
   Self.Caption := Self.Caption + ' :: ' + FsUserNm + ' ' + sUserJikGup + ', 환영합니다!';



   //-----------------------------------------------------
   // Chat 박스 Location Disabled
   //-----------------------------------------------------
   apn_ChatBox.Top   := -1;
   apn_ChatBox.Left  := 999;//460;   --> 숨김처리.
   apn_ChatBox.Caption.Color     := $000066C9;
   apn_ChatBox.Caption.ColorTo   := clMaroon;



   //-----------------------------------------------------
   // 파일 및 이모티콘 전송모드 Off
   //-----------------------------------------------------
   IsEmotiSend := False;
   IsFileSend  := False;



   //-----------------------------------------------------
   // Tmax 개발계 자동 복구 Timer On
   //-----------------------------------------------------
   tm_TxInit.Enabled := True;





   //-----------------------------------------------------
   // 등록된 User 다이얼Pop 메세지 Timer On
   //-----------------------------------------------------
   if (IsLogonUser) then
   begin
      tm_DialPop.Enabled := True;

      //-----------------------------------------------------
      // 다이얼 Pop 로그인 메세지 Check
      //-----------------------------------------------------
      CreatePopUp(CheckDialPop('REALTIME'));
   end
   else
      tm_DialPop.Enabled := False;


   // Test시 아래 주석 후 디버깅할 것..
   // 로그 Update
   UpdateLog('START',
             CopyByte(self.Caption, PosByte('[', self.Caption) + 1, 6),
             FsUserIp,
             'R',
             '',
             FsUserNm,
             varResult
             );

   //-----------------------------------------------------
   // User별 맞춤형 알람 서비스 [생일]
   //-----------------------------------------------------
   if CheckUserAlarm('BORN',
                     FsUserNm,
                     FsUserIp,
                     'ALARM') then
   begin
      apn_Congrats.Top     := 0;
      apn_Congrats.Left    := 0;
      apn_Congrats.Collaps := True;
      apn_Congrats.Visible := True;
      apn_Congrats.Collaps := False;

      Exit;
   end
   else
   begin
      apn_Congrats.Top     := 999;
      apn_Congrats.Left    := 999;
      apn_Congrats.Visible := False;
   end;



   //--------------------------------------------------------------
   // User별 맞춤형 알람 서비스 [메인공지] @ 2015.04.08 LSH
   //       - 이미지, 타이틀, 내용 D/B 동적관리
   //       - COMCD1 = KDIAL, COMCD2 = NOTI, COMCD3 = ALARM
   //--------------------------------------------------------------
   if CheckUserAlarm('NOTI',
                     FsUserNm,
                     FsUserIp,
                     'ALARM') then
   begin
      apn_Congrats.Top     := 0;
      apn_Congrats.Left    := 0;
      apn_Congrats.Collaps := True;
      apn_Congrats.Visible := True;
      apn_Congrats.Collaps := False;

      Exit;
   end
   else
   begin
      apn_Congrats.Top     := 999;
      apn_Congrats.Left    := 999;
      apn_Congrats.Visible := False;
   end;



   //------------------------------------------------------
   // 최신 Ver. 바탕화면에 자동 다운로드(업데이트)
   //          - 형상관리 및 version 충돌 문제로 주석 @ 2017.12.21 LSH
   //------------------------------------------------------
   {
   if (IsLogonUser) then
   begin
      sTargetTitle   := FsVersion + '.exe';

      // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
      sTargetFile    := G_HOMEDIR + 'Exe\' + sTargetTitle;


      try
         Screen.Cursor := crHourGlass;


         // 최신 파일이 존재할 경우, skip
         if FileExists(sTargetFile) and
            // Local file의 Size 체크 추가 (Local에 해당 사번.jpg가 0 size로 존재하면 JPEG Error #41 떨어짐) @ 2015.06.03 LSH
            (CMsg.GetFileSize(sTargetFile) <>  0) then
         begin
            // 기존 최신파일이 존재하면 O.K.
         end
         else
         begin
            //------------------------------------------------------------------
            // Version 체크후, 구 ver. 이면 최신 ver. 다운로드 진행
            //       - ver. 체크시, 'v' 문자 빼고 숫자만 Check 수정 @ 2015.03.05 LSH
            //------------------------------------------------------------------
            if (CopyByte(self.Caption, PosByte('[', self.Caption) + 1, 6) < 'ver' + DelChar(FsVersion, 'v')) then
            begin
               //  -- 병원별 FTP 계정말고, 개발계FTP 계정 접속위해 주석처리 @ 2015.03.04 LSH
//               // 파일 업/다운로드를 위한 정보 조회
//               if not GetBinUploadInfo(sServerIp,
//                                       sFtpUserID,
//                                       sFtpPasswd,
//                                       sFtpHostName,
//                                       sFtpDir) then
//               begin
//                  MessageDlg('파일 저장을 위한 담당자 정보 조회중, 오류가 발생했습니다.',
//                             Dialogs.mtError,
//                             [Dialogs.mbOK],
//                             0);
//                  Exit;
//               end;



               sServerIp   := C_KDIAL_FTP_IP;
               sFtpUserId  := 'bingo';
               sFtpPasswd  := 'bingo123';



               // 실제 서버에 저장되어 있는 파일명 지정
               if PosByte('/ftpspool/',sTargetTitle) > 0 then
                  sRemoteFile := sTargetTitle
               else
                  sRemoteFile := '/ftpspool/' + sTargetTitle;


               sLocalFile  := sTargetFile;

               try
                  if not (GetBINFTP(sServerIp,
                                    sFtpUserID,
                                    sFtpPasswd,
                                    sRemoteFile,
                                    sLocalFile,
                                    False)) then
                  begin
                     //	정상적인 FTP 파일 다운이 안된경우, Msg.
                     MessageBox(self.Handle,
                                PChar('KUMC 다이얼로그 최신Ver. FTP 다운로드 실행중 오류가 발생하였습니다.' + #13#10 + #13#10 +
                                      '프로그램 종료 후 다시 실행해 주시기 바랍니다.'),
                                'KUMC 다이얼로그 최신Ver. FTP 다운로드 오류',
                                MB_OK + MB_ICONERROR);

                     Exit;
                  end;
               except
                  on E : Exception do
                     ShowMessage(E.Message);
               end;


               try
                  if (PosByte('.exe', sLocalFile) > 0) or
                     (PosByte('.EXE', sLocalFile) > 0) then
                  begin
                     MessageBox(self.Handle,
                                PChar('다이얼로그 [' + sTargetTitle + '] 최신버젼이 [바탕화면]에 다운로드 되었습니다.' + #13#10 + #13#10 +
                                      '개선된 기능을 업무에 활용해 보세요.' ),
                                'KUMC 다이얼로그 최신Ver. 자동 다운로드 완료',
                                MB_OK + MB_ICONINFORMATION);
                  end;


               except
                  MessageBox(self.Handle,
                             PChar('KUMC 다이얼로그 최신Ver. 업데이트 실행중 오류가 발생하였습니다.' + #13#10 + #13#10 +
                                   '프로그램 종료 후 다시 실행해 주시기 바랍니다.'),
                             'KUMC 다이얼로그 최신Ver. 자동 업데이트 오류',
                             MB_OK + MB_ICONERROR);

                  Exit;
               end;
            end;
         end;


      finally
         Screen.Cursor := crDefault;
      end;
   end;
   }




   //-----------------------------------------------------
   // 등록된 User 다이얼 Chat 리스트 팝업 @ 2015.03.30 LSH
   //-----------------------------------------------------
   {  -- Chatting 클라이언트 로딩속도 개선작업후 오픈 하자!!
   if (IsLogonUser) then
   begin
      if FindForm('KDIALCHAT') = nil then
      begin
         FForm := BplFormCreate('KDIALCHAT', True);

         SetBplStrProp(FForm, 'prop_UserNm', FsUserNm);

         FForm.Height := (self.Height);
         FForm.Top    := (screen.Height - FForm.Height) - 48;
         FForm.Left   := 8;

         try
            FForm.Show;

         except
            on e : Exception do
               showmessage('Show Form Error(KDIALCHAT) : ' + e.Message);
         end;

         self.SetFocus;

      end;
   end;
   }






   //-------------------------------------------------
   // Login 계정 확인 로직 추후 확장성 고려 남겨둠
   //-------------------------------------------------
   {
   Screen.Cursor := crHourGlass;



   ccusermt := Hccusermt.Create;

   row := ccusermt.Select(ed_userid.Text, s_locate);

   Screen.Cursor := crDefault;

   if row = -1 then
   begin
      ShowErrMsg;                // Tmax Error process

      ccusermt.Free;

      nTry := nTry + 1;

      if nTry > 3 then
      begin
         ModalResult := mrNo
      end
      else
      begin
         ed_passwd.SetFocus;
         ed_passwd.SelectAll;

         Exit;
      end;
   end
   else
   begin
      //if (ed_passwd.Text = ccusermt.sPassword[0]) then
      if (tmpString = ccusermt.sPassword[0]) then
      begin
         G_USERID   := ccusermt.sUserid[0];
         G_USERNM   := ccusermt.sUsername[0];
         G_DEPTCD   := ccusermt.sDeptcd[0];
         G_LOCATE   := ccusermt.sLocate[0];
         G_DEPTNM   := ccusermt.sDeptnm[0];

         // Test by LSH
         if (sNewHospital <> '') then
            G_LOCATENM := '고대의료원 ' + combx_locate.Text + '(개발)'
         else
            G_LOCATENM := '고대의료원 ' + combx_locate.Text;


         G_USERIP := GetIp;

         iUserChk := mrOk;

         ccusermt.Free;
      end
      else
      begin
         ccusermt.Free;

         MakeMsg(CF_D620, D_PASSWORD, NF290); //비밀번호가 일치하지않습니다.

         nTry := nTry + 1;

         if nTry > 3 then
         begin
            iUserChk := mrNo
         end
         else
         begin
            ed_passwd.SetFocus;
            ed_passwd.SelectAll;
           \
            Exit;
         end;
      end;
   end;
   }

end;




//------------------------------------------------------------------------------
// [담당자정보수정] TAdvStringGrid onButtonClick Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2013.08.08
//------------------------------------------------------------------------------
procedure TMainDlg.asg_Master_ButtonClick(Sender: TObject; ACol,
  ARow: Integer);
var
   TpSetMast   : TTpSvc;
   i, iCnt     : Integer;
   vType1,
   vLocate,
   vUserId,
   vUserNm,
   vDeptCd,
   vDutyFlag,
   vMobile,
   vEmail,
   vUserIp,
   vDelDate,
   vNickNm,
   vEditIp : Variant;
   sType   : String;
begin
   // 개인별 정보변경 Update
   if (ACol = C_M_BUTTON) then
   begin


      sType := '1';
      iCnt  := 0;


      //-------------------------------------------------------------------
      // 2. Create Variables
      //-------------------------------------------------------------------
      with asg_Master do
      begin

         if (Cells[C_M_LOCATE, ARow] = '') or
            (Cells[C_M_USERNM, ARow] = '') or
            (Cells[C_M_DUTYPART, ARow] = '') then
         begin
            MessageBox(self.Handle,
                       PChar('[근무처/이름/업무구분]은 필수입력 항목입니다.'),
                       PChar(Self.Caption + ' : 담당자정보 필수입력 항목 알림'),
                       MB_OK + MB_ICONWARNING);

            Exit;
         end;


         if ((PosByte('운영', Cells[C_M_DUTYPART, ARow]) > 0) or
             (PosByte('개발', Cells[C_M_DUTYPART, ARow]) > 0) or
             (PosByte('기획', Cells[C_M_DUTYPART, ARow]) > 0)) and
            ({(Cells[C_M_USERID, ARow] = '') or}   // 개발파트 협력업체 사번 없는 경우 Pass @ 2013.08.28 LSH
             (Cells[C_M_MOBILE, ARow] = '') or
             (Cells[C_M_USERIP, ARow] = '')) then
         begin
            MessageBox(self.Handle,
                       PChar('운영/개발/기획파트의 경우 원활한 업무소통을 위해 [휴대폰/사용IP] 필수입력을 부탁드립니다.'),
                       PChar(Self.Caption + ' : 담당자정보 필수입력 항목 알림'),
                       MB_OK + MB_ICONWARNING);

            Exit;
         end;


         for i := ARow to ARow do
         begin
            //------------------------------------------------------------------
            // 2-1. Append Variables
            //------------------------------------------------------------------
            AppendVariant(vType1   ,   sType                  );
            AppendVariant(vLocate  ,   Cells[C_M_LOCATE,    i]);
            AppendVariant(vUserId  ,   Cells[C_M_USERID,    i]);
            AppendVariant(vUserNm  ,   Cells[C_M_USERNM,    i]);
            AppendVariant(vDutyFlag,   Cells[C_M_DUTYPART,  i]);
            AppendVariant(vMobile  ,   Cells[C_M_MOBILE,    i]);
            AppendVariant(vEmail   ,   Cells[C_M_EMAIL,     i]);
            AppendVariant(vUserIp  ,   Cells[C_M_USERIP,    i]);
            AppendVariant(vDelDate ,   Cells[C_M_DELDATE,   i]);
            AppendVariant(vEditIp  ,   FsUserIp               );
            AppendVariant(vDeptCd  ,   'ITT'                  );
            AppendVariant(vNickNm  ,   Cells[C_M_NICKNM,    i]);

            Inc(iCnt);
         end;
      end;



      if iCnt <= 0 then
         Exit;

      //-------------------------------------------------------------------
      // 3. Insert by TpSvc
      //-------------------------------------------------------------------
      TpSetMast := TTpSvc.Create;
      TpSetMast.Init(Self);


      Screen.Cursor := crHourGlass;


      try
         if TpSetMast.PutSvc('MD_KUMCM_M1',
                             [
                              'S_TYPE1'   , vType1
                            , 'S_STRING1' , vLocate
                            , 'S_STRING2' , vUserId
                            , 'S_STRING3' , vUserNm
                            , 'S_STRING4' , vDutyFlag
                            , 'S_STRING5' , vMobile
                            , 'S_STRING6' , vEmail
                            , 'S_STRING7' , vUserIp
                            , 'S_STRING8' , vEditIp
                            , 'S_STRING9' , vDeptCd
                            , 'S_STRING10', vDelDate
                            , 'S_STRING41', vNickNm
                            ] ) then
         begin
            MessageBox(self.Handle,
                       PChar(asg_Master.Cells[C_M_USERNM, asg_Master.Row] + ' 님의 정보가 정상적으로 [업데이트]되었습니다.'),
                       '[KUMC 다이얼로그] 담당자정보 업데이트 알림 ',
                       MB_OK + MB_ICONINFORMATION);

            // 사용자 Logon 정보 Update
            IsLogonUser := CheckUserInfo(FsUserIp, FsUserNm, FsNickNm, FsUserPart, FsUserSpec, FsUserMobile, FsUserCallNo, FsMngrNm);


            // Refresh
            SelGridInfo('MASTER');
         end
         else
         begin
            ShowMessageM(GetTxMsg);
         end;

      finally
         FreeAndNil(TpSetMast);
         Screen.Cursor  := crDefault;
      end;
   end;




   if (ACol = C_M_USERADD) then
   begin
      asg_Master.InsertRows(ARow + 1, 1);
      asg_Master.AddButton(C_M_USERADD, ARow + 1, asg_Master.ColWidths[C_M_USERADD]-5, 20, '+', haBeforeText, vaCenter);                // 사용자 Add
   end;

end;




//------------------------------------------------------------------------------
// [함수] Variant 속성 생성 함수
//       - 소스출처 : CComFunc.pas
//
// Author : Lee, Se-Ha
// Date   : 2013.08.08
//------------------------------------------------------------------------------
function TMainDlg.AppendVariant(var in_var : Variant; in_str : String ) : Integer;
var
   cnt     : Integer;
   tmp_str : String;
begin
   if ( VarIsArray(in_var) ) then
   begin
      cnt := VarArrayHighBound(in_var,1);

      VarArrayRedim(in_var,cnt + 1 );

      in_var[cnt+1] := in_str;
      result := cnt+1;
   end
   else
   begin
      if ( VarIsEmpty(in_var) ) or
         ( VarIsNull(in_var) ) then
      begin
         in_var    := VarArrayCreate([0,0],varVariant);
         in_var[0] := in_str;
         result    := 0;
      end
      else
      begin
         tmp_str   := VarToStr(in_var);
         in_var    := VarArrayCreate([0,1],varVariant);
         in_var[0] := tmp_str;
         in_var[1] := in_str;
         result    := 1;
      end;
   end;
end;




procedure TMainDlg.asg_Master_GetEditorType(Sender: TObject; ACol,
  ARow: Integer; var AEditor: TEditorType);
begin
   with asg_Master do
   begin
      if (ARow > 0) and (ACol = C_M_LOCATE) then
      begin
         AEditor := edComboList;

         ClearComboString;

         AddComboString('의료원');
         AddComboString('안암');
         AddComboString('구로');
         AddComboString('안산');
      end;

      if (ARow > 0) and (PosByte('의료원', Cells[C_M_LOCATE, ARow]) > 0) and
                        (ACol = C_M_DUTYPART) then
      begin
         AEditor := edComboList;

         ClearComboString;

         AddComboString('관리자');
         AddComboString('홈페이지');
      end;


      if (ARow > 0) and (PosByte('안암', Cells[C_M_LOCATE, ARow]) > 0) and
                        (ACol = C_M_DUTYPART) then
      begin
         AEditor := edComboList;

         ClearComboString;

         AddComboString('운영파트');
         AddComboString('개발파트');
         AddComboString('기획파트');
         AddComboString('OP');
         AddComboString('PC/AS');
         AddComboString('PACS');
      end;


      if (ARow > 0) and ((PosByte('구로', Cells[C_M_LOCATE, ARow]) > 0) or
                         (PosByte('안산', Cells[C_M_LOCATE, ARow]) > 0)) and
                        (ACol = C_M_DUTYPART) then
      begin
         AEditor := edComboList;

         ClearComboString;

         AddComboString('운영파트');
         AddComboString('OP');
         AddComboString('PC/AS');
         AddComboString('PACS');
      end;

      if (ACol = C_M_DELDATE) then
      begin
         AEditor := edDateEdit;
      end;
   end;
end;



procedure TMainDlg.pc_DialogChange(Sender: TObject);
begin

{
   if (pc_Dialog.ActivePage = ts_Dialog) then
   begin
      groupBox2.Caption := '최근 연락망 업데이트 (' + FormatDateTime('yyyy-mm-dd hh:nn', Now) + ' 기준)';

      SelGridInfo('UPDATE');
      SelGridInfo('DIALOG');

      ts_Dialog.Highlighted := True;
      ts_Detail.Highlighted := False;
      ts_Master.Highlighted := False;
      ts_Doc.Highlighted    := False;
   end
   else if (pc_Dialog.ActivePage = ts_Detail) then
   begin
      SelGridInfo('DETAIL');

      ts_Dialog.Highlighted := False;
      ts_Detail.Highlighted := True;
      ts_Master.Highlighted := False;
      ts_Doc.Highlighted    := False;
   end
   else if (pc_Dialog.ActivePage = ts_Master) then
   begin
      SelGridInfo('MASTER');

      ts_Dialog.Highlighted := False;
      ts_Detail.Highlighted := False;
      ts_Master.Highlighted := True;
      ts_Doc.Highlighted    := False;

   end
   else if (pc_Dialog.ActivePage = ts_Doc) then
   begin
      lb_RegDoc.Caption := '▶ 부서공통문서(발송/협조 등)를 검색 또는 등록(수정)하실 수 있습니다.';

      SelGridInfo('DOC');

      ts_Dialog.Highlighted := False;
      ts_Detail.Highlighted := False;
      ts_Master.Highlighted := False;
      ts_Doc.Highlighted    := True;

   end;
   }

end;



//------------------------------------------------------------------------------
// [조회] Grid 조건별 상세 procedure
//
// Author : Lee, Se-Ha
// Date   : 2013.08.08
//------------------------------------------------------------------------------
procedure TMainDlg.SelGridInfo(in_SearchFlag : String);
var
   iRowCnt     : Integer;
   i, j, k     : Integer;
   TpGetSearch : TTpSvc;
   sType1, sType2, sType3, sType4, sType5, sType6, sType7, sType8, sType9, sType10 : String;
   iTargetServerPort  : Integer;
   sTargetServerIp    : String;
   tmpEditTxt         : String;
   iBaseRow, iRowSpan : Integer;
   jBaseRow, jRowSpan : Integer;
   kBaseRow, kRowSpan : Integer;
   sPrevUptitle       : String;
   sPrevMidtitle      : String;
   sPrevDowntitle     : String;
   iLstRowCnt         : Integer;
   ePageCnt           : Real;
   //GetChatUser        : TSelChatUser;
   tmpWkLine1, tmpWkLine2, tmpWkLine3 : String;
   tmpWkProc1, tmpWkProc2, tmpWkProc3, tmpWkProcAll : String;
   FForm : TForm;
   vSvcName : String;
   tmpFlag, tmpLocate, tmpDBLink, tmpRefTable, tmpFstColNm, tmpFstColValue : String;
begin

   vSvcName := 'MD_KUMCM_S1';

   //-----------------------------------------------------------------
   // 1. Grid 초기화
   //-----------------------------------------------------------------
   if (in_SearchFlag = 'DIALOG') then
   begin
      asg_Network.ClearRows(1, asg_Network.RowCount);
      asg_Network.RowCount := 2;

      sType1 := '4';
      sType2 := '';
      sType3 := '';
   end
   else if (in_SearchFlag = 'UPDATE') then
   begin
      asg_NwUpdate.ClearRows(1, asg_NwUpdate.RowCount);
      asg_NwUpdate.RowCount := 2;

      sType1 := '3';
      sType2 := '';
      sType3 := '';
   end
   else if (in_SearchFlag = 'DETAIL') then
   begin
      asg_Detail.ClearRows(1, asg_Detail.RowCount);
      asg_Detail.RowCount := 2;
      asg_Detail.Cells[C_D_DELDATE, 1] := '';
      asg_Detail.AddButton(C_D_USERADD, 1, asg_Detail.ColWidths[C_D_USERADD]-5, 20, '+', haBeforeText, vaCenter);                // 사용자 Add

      sType1 := '2';
      sType2 := '';
      sType3 := '';
   end
   else if (in_SearchFlag = 'MASTER') then
   begin
      asg_Master.ClearRows(1, asg_Master.RowCount);
      asg_Master.RowCount := 2;
      asg_Master.Cells[C_M_DELDATE, 1] := '';
      asg_Master.AddButton(C_M_USERADD, 1, asg_Master.ColWidths[C_M_USERADD]-5, 20, '+', haBeforeText, vaCenter);                // 사용자 Add

      sType1 := '1';
      sType2 := '';
      sType3 := '';
   end
   else if (in_SearchFlag = 'DOC') then
   begin
      asg_RegDoc.ClearRows(1, asg_RegDoc.RowCount);
      asg_RegDoc.RowCount := 2;

      asg_SelDoc.ClearRows(1, asg_SelDoc.RowCount);
      asg_SelDoc.RowCount := 2;


      sType1 := '5';

      if (FormatDateTime('mm', Date) = '01') or
         (FormatDateTime('mm', Date) = '02') then
         sType2 := IntToStr(StrToInt(FormatDateTime('yyyy', Date)) - 1)
      else
         sType2 := FormatDateTime('yyyy', Date);

      // 접속자 IP 식별후, 근무처 Assign
      if PosByte('안암도메인', FsUserIp) > 0 then
         sType3 := '안암'
      else if PosByte('구로도메인', FsUserIp) > 0 then
         sType3 := '구로'
      else if PosByte('안산도메인', FsUserIp) > 0 then
         sType3 := '안산';
   end
   else if (in_SearchFlag = 'DOCREFRESH') then
   begin
      asg_SelDoc.ClearRows(1, asg_SelDoc.RowCount);
      asg_SelDoc.RowCount := 2;


      sType1 := '5';

      {
      if (FormatDateTime('mm', Date) = '01') or
         (FormatDateTime('mm', Date) = '02') then
         sType2 := IntToStr(StrToInt(asg_RegDoc.Cells[C_RD_DOCYEAR, asg_RegDoc.Row]) - 1)
      else
      }

      sType2 := asg_RegDoc.Cells[C_RD_DOCYEAR, asg_RegDoc.Row];

      // 접속자 IP 식별후, 근무처 Assign
      if PosByte('안암도메인', FsUserIp) > 0 then
         sType3 := '안암'
      else if PosByte('구로도메인', FsUserIp) > 0 then
         sType3 := '구로'
      else if PosByte('안산도메인', FsUserIp) > 0 then
         sType3 := '안산';

      sType4 := asg_RegDoc.Cells[C_RD_DOCLIST, asg_RegDoc.Row];
   end
   else if (in_SearchFlag = 'BOARD') then
   begin
      asg_Board.ClearRows(1, asg_Board.RowCount);
      asg_Board.RowCount := 2;

      sType1 := '6';
      sType2 := Trim(fcb_Page.Text);

      // 마지막 Page 잔여 Row 가져오도록 End-Rownum을 Set.
      if (fcb_Page.Items.Count = StrToInt(fcb_Page.Text)) then
         sType3 := IntToStr((fcb_Page.Items.Count - 1) * C_PAGE_MAXROW_UNIT)
      else
         sType3 := '';

   end
   else if (in_SearchFlag = 'BOARDPAGE') then
   begin
      sType1 := '11';
      sType2 := '';
   end
   else if (in_SearchFlag = 'BOARDFIND') then
   begin
      asg_Board.ClearRows(1, asg_Board.RowCount);
      asg_Board.RowCount := 2;

      sType1 := '12';
      sType2 := Trim(fcb_CateUp.Text);
      sType3 := Trim(fcb_CateDown.Text);
      sType4 := Trim(fed_Writer.Text);
      sType5 := Trim(fed_Title.Text);
      sType6 := Trim(fmm_Text.Lines.Text);
   end
   else if (in_SearchFlag = 'CHATUSER') then
   begin
      {  -- Chat XE7 개선전까지 임시 주석 @ 2017.10.31 LSH
      SelChatUser.asg_ChatList.ClearRows(1, SelChatUser.asg_ChatList.RowCount);
      SelChatUser.asg_ChatList.RowCount := 2;

      sType1 := '1';
      sType2 := '';
      sType3 := '';
      }
   end
   else if (in_SearchFlag = 'DIALLIST') then
   begin
      asg_DialList.ClearRows(1, asg_DialList.RowCount);
      asg_DialList.RowCount := 2;


      if (fcb_Scan.Text = '통합검색') then
      begin
         sType1 := '16';
         sType2 := Trim(fed_Scan.Text);

         // 접속자 IP 식별후, 근무처별 order by 적용 @ 2014.07.18 LSH
         if PosByte('안암도메인', FsUserIp) > 0 then
            sType3 := '안암'
         else if PosByte('구로도메인', FsUserIp) > 0 then
            sType3 := '구로'
         else if PosByte('안산도메인', FsUserIp) > 0 then
            sType3 := '안산';

         sType4 := '';
      end
      else if (fcb_Scan.Text = '부서명') then
      begin
         sType1 := '17';
         sType2 := Trim(fed_Scan.Text);
         sType3 := '';
         sType4 := '';
      end
      else if (fcb_Scan.Text = '부서상세') then
      begin
         sType1 := '17';
         sType2 := '';
         sType3 := Trim(fed_Scan.Text);
         sType4 := '';
      end
      else if (fcb_Scan.Text = '연락처') then
      begin
         sType1 := '17';
         sType2 := '';
         sType3 := '';
         sType4 := Trim(fed_Scan.Text);
      end;
   end
   else if (in_SearchFlag = 'MYDIAL') then
   begin
      asg_MyDial.ClearRows(1, asg_MyDial.RowCount);
      asg_MyDial.RowCount := 2;

      sType1 := '18';
      sType2 := FsUserIp;

   end
   else if (in_SearchFlag = 'DIALSCAN') then
   begin
      asg_DialScan.ClearRows(1, asg_DialScan.RowCount);
      asg_DialScan.RowCount := 2;

      sType1 := '20';
      sType2 := 'DIALBOOK';
      sType3 := 'KEYWORD';
      sType4 := FsUserIp;
      sType5 := Trim(fed_Scan.Text);

   end
   else if (in_SearchFlag = 'ANALYSIS') then
   begin
      asg_Analysis.ClearRows(1, asg_Analysis.RowCount);
      asg_Analysis.RowCount := 2;

      if (fcb_Analysis.Text = '통합검색') then
      begin
         sType1 := '21';
         sType2 := '';
         sType3 := Trim(fmed_AnalFrom.Text);
         sType4 := Trim(fmed_AnalTo.Text);
         sType5 := TokenStr(fed_Analysis.Text, ' ', 1);     // 이중 키워드 검색중 1st 키워드 @ 2015.05.20 LSH
         sType6 := TokenStr(fed_Analysis.Text, ' ', 2);     // 이중 키워드 검색중 2nd 키워드 @ 2015.05.20 LSH
         sType7 := '';
      end
      else if (fcb_Analysis.Text = 'S/R제목') then
      begin
         sType1 := '22';
         sType2 := Trim(fed_Analysis.Text);
         sType3 := '';
         sType4 := '';
         sType5 := '';
         sType6 := '';
         sType7 := '';
      end
      else if (fcb_Analysis.Text = '부서') then
      begin
         sType1 := '22';
         sType2 := '';
         sType3 := Trim(fed_Analysis.Text);
         sType4 := '';
         sType5 := '';
         sType6 := '';
         sType7 := '';
      end
      else if (fcb_Analysis.Text = '요청자') then
      begin
         sType1 := '22';
         sType2 := '';
         sType3 := '';
         sType4 := Trim(fed_Analysis.Text);
         sType5 := '';
         sType6 := '';
         sType7 := '';
      end
      else if (fcb_Analysis.Text = '담당자') then
      begin
         sType1 := '22';
         sType2 := '';
         sType3 := '';
         sType4 := '';
         sType5 := Trim(fed_Analysis.Text);
         sType6 := '';
         sType7 := '';
      end
      else if (fcb_Analysis.Text = '처리내역') then
      begin
         sType1 := '22';
         sType2 := '';
         sType3 := '';
         sType4 := '';
         sType5 := '';
         sType6 := Trim(fed_Analysis.Text);
         sType7 := '';
      end;

   end
   else if (in_SearchFlag = 'WORKRPT') then
   begin
      asg_WorkRpt.ClearRows(1, asg_WorkRpt.RowCount);
      asg_WorkRpt.RowCount := 2;

      sType1 := '24';

      if (fcb_WorkArea.Text = '') then
         sType2 := ''
      else if (fcb_WorkArea.Text = '안암') then
         sType2 := 'AA'
      else if (fcb_WorkArea.Text = '구로') then
         sType2 := 'GR'
      else if (fcb_WorkArea.Text = '안산') then
         sType2 := 'AS';

      sType3 := Trim(fmed_AnalFrom.Text);
      sType4 := Trim(fmed_AnalTo.Text);

      {
      if (fcb_WorkGubun.Text = '') then
         sType5 := ''
      else if (fcb_WorkGubun.Text = '운영') then
         sType5 := '운영파트'
      else if (fcb_WorkGubun.Text = '개발') then
         sType5 := '개발파트'
      else if (fcb_WorkGubun.Text = '기획') then
         sType5 := '기획파트'
      else if (fcb_WorkGubun.Text = '본인') then
         sType5 := FsUserNm;
      }

      sType5 := Trim(fed_Analysis.Text);


   end
   else if (in_SearchFlag = 'WEEKLYRPT') then
   begin
      asg_WeeklyRpt.ClearRows(1, asg_WeeklyRpt.RowCount);
      asg_WeeklyRpt.RowCount := 2;


      // Title
      apn_Weekly.Caption.Text := Trim(fed_Analysis.Text) + '의 업무리포트';

      sType1 := '25';

      // 접속자 IP 식별후, 근무처 Assign
      if PosByte('안암', Trim(fcb_WorkArea.Text)) > 0 then
      begin
         sType2 := 'AA';
      end
      else if PosByte('구로', Trim(fcb_WorkArea.Text)) > 0 then
      begin
         sType2 := 'GR';
      end
      else if PosByte('안산', Trim(fcb_WorkArea.Text)) > 0 then
      begin
         sType2 := 'AS';
      end;


      sType3 := Trim(fmed_AnalFrom.Text);
      sType4 := Trim(fmed_AnalTo.Text);
      sType5 := Trim(fed_Analysis.Text);


   end
   else if (in_SearchFlag = 'DIALMAP') then
   begin
      asg_DialMap.ClearRows(0, asg_DialMap.RowCount);
      asg_DialMap.RowCount := 8;

      sType1 := '4';
      sType2 := '';
      sType3 := '';
   end
   else if (in_SearchFlag = 'RELEASE') then
   begin
      asg_Release.ClearRows(1, asg_Release.RowCount);
      asg_Release.RowCount := 2;

      sType1 := '22';
      sType2 := FsUserNm;
      sType3 := Trim(fmed_RegFrDt.Text);
      sType4 := Trim(fmed_RegToDt.Text);
   end
   else if (in_SearchFlag = 'RELEASESCAN') then
   begin
      asg_Release.ClearRows(1, asg_Release.RowCount);
      asg_Release.RowCount := 2;

      sType1 := '28';
      sType2 := Trim(fmed_RegFrDt.Text);
      sType3 := Trim(fmed_RegToDt.Text);
      sType4 := Trim(fcb_DutySpec.Text);
      sType5 := Trim(fed_Release.Text);
   end
   // 근무-근태 표 생성
   else if (in_SearchFlag = 'DUTYMAKE') then
   begin
      asg_Duty.ClearRows(1, asg_Duty.RowCount);
      asg_Duty.RowCount := 2;

      sType1 := '29';
      sType2 := Trim(fmed_DutyFrDt.Text);
      sType3 := Trim(fmed_DutyToDt.Text);
   end
   else if (in_SearchFlag = 'DUTYLIST') then
   begin
      asg_DutyList.ClearRows(1, asg_DutyList.RowCount);
      asg_DutyList.RowCount := 2;

      sType1 := '30';
   end
   else if (in_SearchFlag = 'DBCATEGORY') then
   begin
      asg_DBMaster.Options    := [goFixedVertLine,goFixedHorzLine,goVertLine,goHorzLine,goColSizing,goRowSelect];
      asg_DBDetail.Options    := [goFixedVertLine,goFixedHorzLine,goVertLine,goHorzLine,goColSizing,goRowSelect,goRangeSelect];

      asg_DBMaster.ClearRows(1, asg_DBMaster.RowCount);
      asg_DBMaster.RowCount := 2;

      asg_DBDetail.ClearRows(1, asg_DBDetail.RowCount);
      asg_DBDetail.RowCount := 2;

      if (fcb_WorkArea.Text = '안암') then
         tmpLocate := '@AAOCS'
      else if (fcb_WorkArea.Text = '구로') then
         tmpLocate := '@GROCS'
      else if (fcb_WorkArea.Text = '안산') then
         tmpLocate := '@ASOCS';

      if (fcb_DutyPart.Text = '원무') then
         tmpDBLink := tmpLocate + '_ORAA1_REAL'
      else if (fcb_DutyPart.Text = '진료') then
         tmpDBLink := tmpLocate + '_ORAM1_REAL'
      else if (fcb_DutyPart.Text = '진료지원') then
         tmpDBLink := tmpLocate + '_ORAS1_REAL'
      else if (fcb_DutyPart.Text = '공통관리') then
         tmpDBLink := tmpLocate + '_ORAC1_REAL'
      else if (fcb_DutyPart.Text = 'e-Refer') then
         tmpDBLink := tmpLocate + '_ORAE1'
      else if (fcb_DutyPart.Text = 'HRM') then
         tmpDBLink := '@AAINT_HRMINF'
      else if (fcb_DutyPart.Text = '일반관리') then
         tmpDBLink := '@AAINT_ORAG1';

      if (fcb_WorkGubun.Text = '공통코드') then
      begin
         if (fcb_DutyPart.Text = '원무') then
         begin
            tmpFlag        := 'A';
            tmpRefTable    := 'CCCOMCDT' + tmpDBLink;
            tmpFstColNm    := 'LARGCD';
            tmpFstColValue := '0000';
         end
         else if (fcb_DutyPart.Text = '진료') then
         begin
            tmpFlag        := 'M';
            tmpRefTable    := 'MDMCOMCT' + tmpDBLink;
            tmpFstColNm    := 'COMCD1';
            tmpFstColValue := '000';
         end
         else if (fcb_DutyPart.Text = '진료지원') then
         begin
            tmpFlag        := 'S';
            tmpRefTable    := 'SDCOMCDT' + tmpDBLink;
            tmpFstColNm    := 'LARGCD';
            tmpFstColValue := 'SD';
         end
         else if (fcb_DutyPart.Text = '일반관리') then
         begin
            tmpFlag        := 'G';
            tmpRefTable    := 'GPCOMCDT' + tmpDBLink;
            tmpFstColNm    := 'LARGCD';
            tmpFstColValue := 'GP00';
         end;

         sType1 := '31';

      end
      else if (fcb_WorkGubun.Text = 'Table') then
      begin
         if (fcb_DutyPart.Text = '원무') then
            tmpFlag        := 'ORAA1'
         else if (fcb_DutyPart.Text = '진료') then
            tmpFlag        := 'ORAM1'
         else if (fcb_DutyPart.Text = '진료지원') then
            tmpFlag        := 'ORAS1'
         else if (fcb_DutyPart.Text = '공통관리') then
            tmpFlag        := 'ORAC1'
         else if (fcb_DutyPart.Text = 'e-Refer') then
            tmpFlag        := 'ORAE1'
         else if (fcb_DutyPart.Text = 'HRM') then
            tmpFlag        := 'HRMINF'
         else if (fcb_DutyPart.Text = '일반관리') then
            tmpFlag        := 'ORAG1';


         tmpRefTable    := 'ALL_OBJECTS' + tmpDBLink + ' a';

         if (fcb_DutyPart.Text = 'HRM') or
            (fcb_DutyPart.Text = '일반관리') then
            tmpFstColNm    := 'USER_TAB_COMMENTS' + tmpDBLink + ' b'
         else
            tmpFstColNm    := 'USER_TAB_COMMENTS' + tmpLocate + '_' + tmpFlag + ' b';

         tmpFstColValue := AnsiUpperCase(fcb_WorkGubun.Text);

         sType1 := '33';
         sType6 := 'TABLE_NAME';

      end
      else if (fcb_WorkGubun.Text = 'View') then
      begin
         if (fcb_DutyPart.Text = '원무') then
            tmpFlag        := 'ORAA1'
         else if (fcb_DutyPart.Text = '진료') then
            tmpFlag        := 'ORAM1'
         else if (fcb_DutyPart.Text = '진료지원') then
            tmpFlag        := 'ORAS1'
         else if (fcb_DutyPart.Text = '공통관리') then
            tmpFlag        := 'ORAC1'
         else if (fcb_DutyPart.Text = 'e-Refer') then
            tmpFlag        := 'ORAE1'
         else if (fcb_DutyPart.Text = 'HRM') then
            tmpFlag        := 'HRMINF'
         else if (fcb_DutyPart.Text = '일반관리') then
            tmpFlag        := 'ORAG1';


         tmpRefTable    := 'ALL_OBJECTS' + tmpDBLink + ' a';

         if (fcb_DutyPart.Text = 'HRM') or
            (fcb_DutyPart.Text = '일반관리') then
            tmpFstColNm    := 'USER_VIEWS' + tmpDBLink + ' b'
         else
            tmpFstColNm    := 'USER_VIEWS' + tmpLocate + '_' + tmpFlag + ' b';

         tmpFstColValue := AnsiUpperCase(fcb_WorkGubun.Text);

         sType1 := '33';
         sType6 := 'VIEW_NAME';

      end
      else if (fcb_WorkGubun.Text = 'Procedure') or
              (fcb_WorkGubun.Text = 'Function')  or
              (fcb_WorkGubun.Text = 'Trigger')   or
              (fcb_WorkGubun.Text = 'Package')   then
      begin
         if (fcb_DutyPart.Text = '원무') then
            tmpFlag        := 'ORAA1'
         else if (fcb_DutyPart.Text = '진료') then
            tmpFlag        := 'ORAM1'
         else if (fcb_DutyPart.Text = '진료지원') then
            tmpFlag        := 'ORAS1'
         else if (fcb_DutyPart.Text = '공통관리') then
            tmpFlag        := 'ORAC1'
         else if (fcb_DutyPart.Text = 'e-Refer') then
            tmpFlag        := 'ORAE1'
         else if (fcb_DutyPart.Text = 'HRM') then
            tmpFlag        := 'HRMINF'
         else if (fcb_DutyPart.Text = '일반관리') then
            tmpFlag        := 'ORAG1';


         tmpRefTable    := 'ALL_OBJECTS' + tmpDBLink + ' a';

         if (fcb_DutyPart.Text = 'HRM') or
            (fcb_DutyPart.Text = '일반관리')then
            tmpFstColNm    := 'USER_PROCEDURES' + tmpDBLink + ' b'
         else
            tmpFstColNm    := 'USER_PROCEDURES' + tmpLocate + '_' + tmpFlag + ' b';

         tmpFstColValue := AnsiUpperCase(fcb_WorkGubun.Text);

         sType1 := '34';
         sType6 := 'OBJECT_NAME';
      end
      else if (fcb_WorkGubun.Text = 'Slip코드') then
      begin
         tmpFlag        := 'SLIP';
         tmpRefTable    := 'MDSLIPTT' + tmpDBLink;
         tmpFstColNm    := 'SLIP1CD';
         tmpFstColValue := '*';

         sType1 := '31';

      end;


      sType2 := tmpFlag;
      sType3 := tmpRefTable;
      sType4 := tmpFstColNm;
      sType5 := tmpFstColValue;

      // Grid 컬럼 정보 변경
      SetGridCol(AnsiUpperCase(fcb_WorkGubun.Text));

   end
   else if (in_SearchFlag = 'DBCATEDET') then
   begin
      asg_DBDetail.ClearRows(1, asg_DBDetail.RowCount);
      asg_DBDetail.RowCount := 2;

      if (fcb_WorkArea.Text = '안암') then
         tmpLocate := '@AAOCS'
      else if (fcb_WorkArea.Text = '구로') then
         tmpLocate := '@GROCS'
      else if (fcb_WorkArea.Text = '안산') then
         tmpLocate := '@ASOCS';

      if (fcb_DutyPart.Text = '원무') then
         tmpDBLink := tmpLocate + '_ORAA1_REAL'
      else if (fcb_DutyPart.Text = '진료') then
         tmpDBLink := tmpLocate + '_ORAM1_REAL'
      else if (fcb_DutyPart.Text = '진료지원') then
         tmpDBLink := tmpLocate + '_ORAS1_REAL'
      else if (fcb_DutyPart.Text = '공통관리') then
         tmpDBLink := tmpLocate + '_ORAC1_REAL'
      else if (fcb_DutyPart.Text = 'e-Refer') then
         tmpDBLink := tmpLocate + '_ORAE1'
      else if (fcb_DutyPart.Text = '일반관리') then
         tmpDBLink := '@AAINT_ORAG1'
      else if (fcb_DutyPart.Text = 'HRM') then
         tmpDBLink := '@AAINT_HRMINF';

      if (fcb_WorkGubun.Text = '공통코드') then
      begin
         if (fcb_DutyPart.Text = '원무') then
         begin
            tmpFlag        := 'A';
            tmpRefTable    := 'CCCOMCDT' + tmpDBLink;
            tmpFstColNm    := 'LARGCD';
            tmpFstColValue := Trim(asg_DBMaster.Cells[0, asg_DBMaster.Row]);
         end
         else if (fcb_DutyPart.Text = '진료') then
         begin
            tmpFlag        := 'M';
            tmpRefTable    := 'MDMCOMCT' + tmpDBLink;
            tmpFstColNm    := 'COMCD1';
            tmpFstColValue := Trim(asg_DBMaster.Cells[0, asg_DBMaster.Row]);
         end
         else if (fcb_DutyPart.Text = '진료지원') then
         begin
            tmpFlag        := 'S';
            tmpRefTable    := 'SDCOMCDT' + tmpDBLink;
            tmpFstColNm    := 'LARGCD';
            tmpFstColValue := Trim(asg_DBMaster.Cells[0, asg_DBMaster.Row]);
         end
         else if (fcb_DutyPart.Text = '일반관리') then
         begin
            tmpFlag        := 'G';
            tmpRefTable    := 'GPCOMCDT' + tmpDBLink;
            tmpFstColNm    := 'LARGCD';
            tmpFstColValue := Trim(asg_DBMaster.Cells[0, asg_DBMaster.Row]);
         end;
      end
      else if (fcb_WorkGubun.Text = 'Table') then
      begin
         tmpFlag        := AnsiUpperCase(fcb_WorkGubun.Text);

         if (fcb_DutyPart.Text = 'HRM') then
            tmpRefTable    := CopyByte(tmpDBLink, 1, 13)
         else
            tmpRefTable    := CopyByte(tmpDBLink, 1, 12);

         tmpFstColValue := Trim(asg_DBMaster.Cells[0, asg_DBMaster.Row]) ;
      end
      else if (fcb_WorkGubun.Text = 'Procedure') or
              (fcb_WorkGubun.Text = 'Function')  or
              (fcb_WorkGubun.Text = 'Package')   then
      begin
         tmpFlag        := AnsiUpperCase(fcb_WorkGubun.Text);
         tmpRefTable    := 'USER_SOURCE' + tmpDBLink;
         tmpFstColValue := Trim(asg_DBMaster.Cells[0, asg_DBMaster.Row]) ;
      end
      else if (fcb_WorkGubun.Text = 'Trigger') then
      begin
         tmpFlag        := AnsiUpperCase(fcb_WorkGubun.Text);

         if (fcb_DutyPart.Text = 'HRM') then
            tmpRefTable    := 'USER_TRIGGERS' + CopyByte(tmpDBLink, 1, 13)
         else
            tmpRefTable    := 'USER_TRIGGERS' + CopyByte(tmpDBLink, 1, 12);

         tmpFstColValue := Trim(asg_DBMaster.Cells[0, asg_DBMaster.Row]) ;
      end
      else if (fcb_WorkGubun.Text = 'View') then
      begin
         tmpFlag        := AnsiUpperCase(fcb_WorkGubun.Text);

         if (fcb_DutyPart.Text = 'HRM') then
            tmpRefTable    := 'USER_VIEWS' + CopyByte(tmpDBLink, 1, 13)
         else
            tmpRefTable    := 'USER_VIEWS' + CopyByte(tmpDBLink, 1, 12);

         tmpFstColValue := Trim(asg_DBMaster.Cells[0, asg_DBMaster.Row]) ;
      end
      else if (fcb_WorkGubun.Text = 'Slip코드') then
      begin
         tmpFlag        := 'SLIP';
         tmpRefTable    := 'MDSLIPTT' + tmpDBLink;
         tmpFstColNm    := 'SLIP1CD';
         tmpFstColValue := Trim(asg_DBMaster.Cells[0, asg_DBMaster.Row]);
      end;


      sType1 := '32';
      sType2 := tmpFlag;
      sType3 := tmpRefTable;
      sType4 := tmpFstColValue;
      sType5 := '';

      // Grid 컬럼 정보 변경
      SetGridCol(AnsiUpperCase(fcb_WorkGubun.Text));

   end
   else if (in_SearchFlag = 'DBSEARCH') then
   begin
      asg_DBMaster.Options    := [goFixedVertLine,goFixedHorzLine,goVertLine,goHorzLine,goColSizing,goRowSelect];
      asg_DBDetail.Options    := [goFixedVertLine,goFixedHorzLine,goVertLine,goHorzLine,goColSizing,goRangeSelect];

      asg_DBMaster.ClearRows(1, asg_DBMaster.RowCount);
      asg_DBMaster.RowCount := 2;

      asg_DBDetail.ClearRows(1, asg_DBDetail.RowCount);
      asg_DBDetail.RowCount := 2;

      if (fcb_WorkArea.Text = '안암') then
         tmpLocate := '@AAOCS'
      else if (fcb_WorkArea.Text = '구로') then
         tmpLocate := '@GROCS'
      else if (fcb_WorkArea.Text = '안산') then
         tmpLocate := '@ASOCS';

      if (fcb_DutyPart.Text = '원무') then
         tmpDBLink := tmpLocate + '_ORAA1_REAL'
      else if (fcb_DutyPart.Text = '진료') then
         tmpDBLink := tmpLocate + '_ORAM1_REAL'
      else if (fcb_DutyPart.Text = '진료지원') then
         tmpDBLink := tmpLocate + '_ORAS1_REAL'
      else if (fcb_DutyPart.Text = '공통관리') then
         tmpDBLink := tmpLocate + '_ORAC1_REAL'
      else if (fcb_DutyPart.Text = 'e-Refer') then
         tmpDBLink := tmpLocate + '_ORAE1'
      else if (fcb_DutyPart.Text = 'HRM') then
         tmpDBLink := '@AAINT_HRMINF'
      else if (fcb_DutyPart.Text = '일반관리') then
         tmpDBLink := '@AAINT_ORAG1';




      if (fcb_DutyPart.Text = '원무') then
         tmpFlag        := 'ORAA1'
      else if (fcb_DutyPart.Text = '진료') then
         tmpFlag        := 'ORAM1'
      else if (fcb_DutyPart.Text = '진료지원') then
         tmpFlag        := 'ORAS1'
      else if (fcb_DutyPart.Text = '공통관리') then
         tmpFlag        := 'ORAC1'
      else if (fcb_DutyPart.Text = 'e-Refer') then
         tmpFlag        := 'ORAE1'
      else if (fcb_DutyPart.Text = 'HRM') then
         tmpFlag        := 'HRMINF'
      else if (fcb_DutyPart.Text = '일반관리') then
         tmpFlag        := 'ORAG1';


      tmpRefTable    := 'all_objects' + tmpDBLink + ' a';


      if (fcb_DutyPart.Text = 'HRM') or
         (fcb_DutyPart.Text = '일반관리') then
      begin
         tmpFstColNm    := 'USER_TAB_COLUMNS' + tmpDBLink + ' b';
         tmpFstColValue := 'USER_COL_COMMENTS'+ tmpDBLink + ' c';
         sType6         := 'USER_TAB_COMMENTS'+ tmpDBLink + ' d';
      end
      else
      begin
         tmpFstColNm    := 'USER_TAB_COLUMNS'  + tmpLocate + '_' + tmpFlag + ' b';
         tmpFstColValue := 'USER_COL_COMMENTS' + tmpLocate + '_' + tmpFlag + ' c';
         sType6         := 'USER_TAB_COMMENTS' + tmpLocate + '_' + tmpFlag + ' d';
      end;


      sType1 := '35';
      sType2 := tmpFlag;
      sType3 := tmpRefTable;
      sType4 := tmpFstColNm;
      sType5 := tmpFstColValue;
      sType7 := 'USER_PROCEDURES'  + tmpDBLink;
      sType8 := AnsiUpperCase(Trim(fed_Analysis.Text));
      sType9 := 'USER_SOURCE'      + tmpDBLink;


      // Grid 컬럼 정보 변경
      SetGridCol(in_SearchFlag);

   end
   else if (in_SearchFlag = 'SRHISTORY') then
   begin
      asg_WeeklyRpt.ClearRows(1, asg_WeeklyRpt.RowCount);
      asg_WeeklyRpt.RowCount := 2;

      // Title
      apn_Weekly.Caption.Text := Trim(fed_Analysis.Text) + '의 S/R 요청현황';

      sType1 := '36';
      sType2 := '';
      sType3 := '';
      sType4 := '';
      sType5 := Trim(fed_Analysis.Text);

   end
   else if (in_SearchFlag = 'WORKCONN') then
   begin
      asg_WorkConn.ClearRows(1, asg_WorkConn.RowCount);
      asg_WorkConn.RowCount := 2;


      sType1 := '38';
      sType2 := Trim(fmed_ConnFrom.Text);
      sType3 := Trim(fmed_ConnTo.Text);
      sType4 := Trim(fcb_WorkArea.Text);
      //sType5 := Trim(fed_Analysis.Text);
   end
   else if (in_SearchFlag = 'BIGDATA') then
   begin
      asg_BigData.ClearRows(1, asg_BigData.RowCount);
      asg_BigData.RowCount := 2;

      sType1 := '39';
      sType2 := '';
      sType3 := Trim(fmed_LotFrDt.Text);
      sType4 := Trim(fmed_LotToDt.Text);
   end
   // 당직근무-근태(확정) 이력 조회
   else if (in_SearchFlag = 'DUTYSEARCH') then
   begin
      asg_Duty.ClearRows(1, asg_Duty.RowCount);
      asg_Duty.RowCount := 2;

      sType1 := '45';
      sType2 := Trim(fmed_DutyFrDt.Text);
      sType3 := Trim(fmed_DutyToDt.Text);
   end;




   //-----------------------------------------------------------------
   // Dynamic Column 관리 위한 Index 초기화
   //-----------------------------------------------------------------
   iLstRowCnt := 0;





   //-----------------------------------------------------------------
   // 2. 조회
   //-----------------------------------------------------------------
   Screen.Cursor := crHourGlass;





   //-----------------------------------------------------------------
   // 2-1. Load Data by TpSvc
   //-----------------------------------------------------------------
   TpGetSearch := TTpSvc.Create;
   TpGetSearch.Init(Self);



   try
      TpGetSearch.CountField  := 'S_CODE1';
      TpGetSearch.ShowMsgFlag := False;

      if TpGetSearch.GetSvc(vSvcName,
                          ['S_TYPE1'  , sType1
                         , 'S_TYPE2'  , sType2
                         , 'S_TYPE3'  , sType3
                         , 'S_TYPE4'  , sType4
                         , 'S_TYPE5'  , sType5
                         , 'S_TYPE6'  , sType6
                         , 'S_TYPE7'  , sType7
                         , 'S_TYPE8'  , sType8
                         , 'S_TYPE9'  , sType9
                         , 'S_TYPE10' , sType10
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
            asg_NwUpdate.RowCount   := 2;
            asg_NetWork.RowCount    := 2;
            asg_Master.RowCount     := 2;
            asg_Detail.RowCount     := 2;
            asg_SelDoc.RowCount     := 2;
            asg_Board.RowCount      := 2;
            asg_DialList.RowCount   := 2;
            asg_MyDial.RowCount     := 2;
            asg_Release.RowCount    := 2;

            ShowMessage(GetTxMsg);

            Exit;
         end
         else if TpGetSearch.RowCount = 0 then
         begin
            lb_NetWork.Caption   := '▶ 조회내역이 존재하지 않습니다.';
            lb_SelDoc.Caption    := '▶ 조회내역이 존재하지 않습니다.';
            lb_Board.Caption     := '▶ 조회내역이 존재하지 않습니다.';
            lb_DialScan.Caption  := '▶ 조회내역이 존재하지 않습니다.';
            lb_MyDial.Caption    := '';
            lb_Analysis.Caption  := '▶ 조회내역이 존재하지 않습니다.';
            //lb_WorkConn.Caption  := '▶ 조회내역이 존재하지 않습니다.';


            if (in_SearchFlag = 'RELEASE') or
               (in_SearchFlag = 'RELEASESCAN') then
               lb_RegDoc.Caption    := '▶ 조회내역이 존재하지 않습니다.';

            Exit;
         end;




      //-----------------------------------------------------------------
      // 4. Display Data
      //-----------------------------------------------------------------
      if (in_SearchFlag = 'DIALOG') then
      begin
         with asg_NetWork do
         begin
            iRowCnt  := TpGetSearch.RowCount;


            // Merge 변수
            iBaseRow       := 0;    // 대분류 Merge Start Index
            iRowSpan       := 1;    // 대분류 Span  Index
            jBaseRow       := 0;    // 중분류 Merge Start Index
            jRowSpan       := 1;    // 중분류 Span  Index
            kBaseRow       := 0;    // 소분류 Merge Start Index
            kRowSpan       := 1;    // 소분류 Span  Index

            sPrevUptitle    := '';   // 대분류 이전 Title
            sPrevMidtitle   := '';   // 중분류 이전 Title
            sPrevDowntitle  := '';   // 소분류 이전 Title




            for i := 0 to iRowCnt - 1 do
            begin
               Cells[C_NW_LOCATE,   RowCount - 1] := TpGetSearch.GetOutputDataS('sLocate'    , i);
               Cells[C_NW_DUTYPART, RowCount - 1] := TpGetSearch.GetOutputDataS('sDutyPart'  , i);
               Cells[C_NW_DUTYSPEC, RowCount - 1] := TpGetSearch.GetOutputDataS('sDutySpec'  , i);
               Cells[C_NW_USERNM,   RowCount - 1] := TpGetSearch.GetOutputDataS('sUserNm'    , i);
               Cells[C_NW_CALLNO,   RowCount - 1] := TpGetSearch.GetOutputDataS('sCallNo'    , i);
               Cells[C_NW_MOBILE,   RowCount - 1] := TpGetSearch.GetOutputDataS('sMobile'    , i);
               Cells[C_NW_DUTYRMK,  RowCount - 1] := TpGetSearch.GetOutputDataS('sDutyRmk'   , i);
               Cells[C_NW_EMAIL,    RowCount - 1] := TpGetSearch.GetOutputDataS('sEmail'     , i);
               Cells[C_NW_USERIP,   RowCount - 1] := TpGetSearch.GetOutputDataS('sUserIp'    , i);
               Cells[C_NW_PHOTOFILE,RowCount - 1] := DeleteStr(TpGetSearch.GetOutputDataS('sAttachNm'  , i), '/media/cq/photo/');
               Cells[C_NW_HIDEFILE, RowCount - 1] := TpGetSearch.GetOutputDataS('sHideFile'  , i);
               Cells[C_NW_USERID,   RowCount - 1] := TpGetSearch.GetOutputDataS('sUserId'    , i);     // 사용자 ID 추가 @ 2014.06.26 LSH

               //---------------------------------------------------------------
               // 접속하려는 Server 사용가능 여부 (현재 Session On/Off) Check
               //---------------------------------------------------------------
               if (TpGetSearch.GetOutputDataS('sLogYn', i) = 'Y') then
                  FontColors[C_NW_USERNM, RowCount - 1] := $0000CDA3
               else
                  FontColors[C_NW_USERNM, RowCount - 1] := clBlack;


               // 현재 접속한 User의 Chat-List 위치를 기억해둔다 @ 2015.03.30 LSH
               if TpGetSearch.GetOutputDataS('sUserNm', i) = FsUserNm then
                  iUserRowId := RowCount - 1;

               // 해당 화면 1-Cycle 코드의 마지막 Index를 최종 RowCount 로 사용
               if (i = iRowCnt) then
                  iLstRowCnt := i+1
               else
                  iLstRowCnt := iRowCnt;



               // 대분류 항목 MergeCell 체크
               if TpGetSearch.GetOutputDataS('sLocate', i) <> sPrevUptitle then
               begin
                  // test
                  //Memo1.Lines.Add('--- Test Logging Start ---');


                  // 이전 대분류와 다른 경우
                  if i > 0 then
                  begin
                     MergeCells(0, iBaseRow, 1, iRowSpan);
                     iRowSpan := 1;
                  end;

                  // ★ Key Point
                  iBaseRow := RowCount - 1;


                  // 중분류 항목 MergeCell 체크
                  if TpGetSearch.GetOutputDataS('sDutyPart', i) <> sPrevMidTitle then
                  begin
                     // 이전 중분류와 다른 경우
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;


                     // ★ Key Point
                     jBaseRow := RowCount - 1;



                     // 소분류 항목 MergeCell 체크
                     if TpGetSearch.GetOutputDataS('sDutySpec', i) <> sPrevDownTitle then
                     begin
                        // 이전 소분류와 다른 경우
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // ★ Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // 이전 소분류와 다른 경우
                        Inc(kRowSpan);


                  end
                  else
                  begin
                     // 이전 중분류와 같은 경우
                     Inc(jRowSpan);
                  end;
               end
               else
               begin

                  // 중분류 항목 MergeCell 체크
                  if TpGetSearch.GetOutputDataS('sDutyPart', i) <> sPrevMidTitle then
                  begin
                     // 이전 중분류와 다른 경우
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;

                     // ★ Key Point
                     jBaseRow := RowCount - 1;


                     // 소분류 항목 MergeCell 체크
                     if TpGetSearch.GetOutputDataS('sDutySpec', i) <> sPrevDownTitle then
                     begin
                        // 이전 소분류와 같은 경우
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // ★ Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // 이전 소분류와 다른 경우
                        Inc(kRowSpan);
                  end
                  else
                  begin
                     // 중분류 span Index ++
                     Inc(jRowSpan);


                     // 소분류 항목 MergeCell 체크
                     if TpGetSearch.GetOutputDataS('sDutySpec', i) <> sPrevDownTitle then
                     begin
                        // 이전 소분류와 다른 경우
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // ★ Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // 이전 소분류와 동일한 경우
                        Inc(kRowSpan);

                  end;


                  // 대분류 Span index ++
                  Inc(iRowSpan);

               end;


               // Set previous value
               sPrevUptitle   := TpGetSearch.GetOutputDataS('sLocate'  , i);
               sPrevmidtitle  := TpGetSearch.GetOutputDataS('sDutyPart', i);
               sPrevDowntitle := TpGetSearch.GetOutputDataS('sDutySpec', i);


               // RowCount 다음줄로 넘김.
               RowCount := RowCount + 1;

            end;


            // ★ 최종 Code List-up후 마지막 Row Merge 처리
            asg_NetWork.MergeCells(0, iBaseRow, 1, iRowSpan);
            asg_NetWork.MergeCells(1, jBaseRow, 1, jRowSpan);
            asg_NetWork.MergeCells(2, kBaseRow, 1, kRowSpan);
            asg_NetWork.RowCount := iLstRowCnt + 1;

            // Comments
            lb_Network.Caption := '▶ [담당자정보]와 [업무프로필]을 업데이트 하시면 비상연락망에 자동조회 됩니다.';


         end;
      end
      else if (in_SearchFlag = 'UPDATE') then
      begin
         with asg_NwUpdate do
         begin
            iRowCnt  := TpGetSearch.RowCount;
            RowCount := iRowCnt + FixedRows + 1;


            for i := 0 to iRowCnt - 1 do
            begin
               for j := 0 to ColCount - 1 do
               begin
                  if (TpGetSearch.GetOutputDataS('sDelDate'  , i) = '') then
                     tmpEditTxt :=  TpGetSearch.GetOutputDataS('sLocate',   i) + ' '   +
                                    TpGetSearch.GetOutputDataS('sDutyPart', i) + ' '   +
                                    TpGetSearch.GetOutputDataS('sUserNm',   i) + '의 "' +
                                    TpGetSearch.GetOutputDataS('sFlag'  ,   i) + '" 업데이트됨'
                  else
                     tmpEditTxt :=  TpGetSearch.GetOutputDataS('sLocate',   i) + ' '   +
                                    TpGetSearch.GetOutputDataS('sDutyPart', i) + ' '   +
                                    TpGetSearch.GetOutputDataS('sUserNm',   i) + '의 "' +
                                    TpGetSearch.GetOutputDataS('sFlag'  ,   i) + '" 종료(삭제)됨';

                  Cells[C_NU_EDITDATE, i+FixedRows] := TpGetSearch.GetOutputDataS('sEditDate', i);
                  Cells[C_NU_EDITTEXT, i+FixedRows] := tmpEditTxt;
                  Cells[C_NU_EDITIP,   i+FixedRows] := TpGetSearch.GetOutputDataS('sEditIp'  , i);
               end;
            end;
         end;

         // RowCount 정리
         asg_NwUpdate.RowCount := asg_NwUpdate.RowCount - 1;

      end
      else if (in_SearchFlag = 'DETAIL') then
      begin
         with asg_Detail do
         begin
            iRowCnt  := TpGetSearch.RowCount;
            RowCount := iRowCnt + FixedRows + 1;


            for i := 0 to iRowCnt - 1 do
            begin
               for j := 0 to ColCount - 1 do
               begin
                  Cells[C_D_LOCATE,  i+FixedRows] := TpGetSearch.GetOutputDataS('sLocate'    , i);
                  Cells[C_D_DUTYPART,i+FixedRows] := TpGetSearch.GetOutputDataS('sDutyPart'  , i);
                  Cells[C_D_DUTYSPEC,i+FixedRows] := TpGetSearch.GetOutputDataS('sDutySpec'  , i);
                  Cells[C_D_DUTYRMK, i+FixedRows] := TpGetSearch.GetOutputDataS('sDutyRmk'   , i);
                  Cells[C_D_DUTYUSER,i+FixedRows] := TpGetSearch.GetOutputDataS('sDutyUser'  , i);
                  Cells[C_D_CALLNO,  i+FixedRows] := TpGetSearch.GetOutputDataS('sCallNo'    , i);
                  Cells[C_D_DUTYPTNR,i+FixedRows] := TpGetSearch.GetOutputDataS('sDutyPtnr'  , i);

                  AddButton(C_D_USERADD, i+FixedRows, ColWidths[C_D_USERADD]-5, 20, '+', haBeforeText, vaCenter);                // 사용자 Add
               end;
            end;
         end;


         // RowCount 정리
         asg_Detail.RowCount := asg_Detail.RowCount - 1;

         // Comments
         lb_Network.Caption := '▶ ' + IntToStr(iRowCnt) + '건의 업무프로필을 조회하였습니다.';


      end
      else if (in_SearchFlag = 'MASTER') then
      begin
          with asg_Master do
          begin
            // Maximum value of progress status
            //fpb_DataLoading.Max := TpGetSearch.RowCount * (asg_Master.ColCount);

            iRowCnt  := TpGetSearch.RowCount;
            RowCount := iRowCnt + FixedRows + 1;


            for i := 0 to iRowCnt - 1 do
            begin
               for j := 0 to ColCount - 1 do
               begin
                  Cells[C_M_LOCATE,  i+FixedRows] := TpGetSearch.GetOutputDataS('sLocate'    , i);
                  Cells[C_M_USERID,  i+FixedRows] := TpGetSearch.GetOutputDataS('sUserId'    , i);
                  Cells[C_M_USERNM,  i+FixedRows] := TpGetSearch.GetOutputDataS('sUserNm'    , i);
                  Cells[C_M_DUTYPART,i+FixedRows] := TpGetSearch.GetOutputDataS('sDutyPart'  , i);
                  Cells[C_M_MOBILE,  i+FixedRows] := TpGetSearch.GetOutputDataS('sMobile'    , i);
                  Cells[C_M_EMAIL,   i+FixedRows] := TpGetSearch.GetOutputDataS('sEmail'     , i);
                  Cells[C_M_USERIP,  i+FixedRows] := TpGetSearch.GetOutputDataS('sUserIp'    , i);
                  Cells[C_M_NICKNM,  i+FixedRows] := TpGetSearch.GetOutputDataS('sNickNm'    , i);


                  AddButton(C_M_USERADD, i+FixedRows, ColWidths[C_M_USERADD]-5, 20, '+', haBeforeText, vaCenter);                // 사용자 Add



                  //------------------------------------------------------------
                  // 접속하려는 Server 사용가능 여부 (현재 Session On/Off) Check
                  //------------------------------------------------------------
                  if (TpGetSearch.GetOutputDataS('sLogYn', i) = 'Y') then
                     Colors[C_M_USERNM, i+FixedRows] := $0000CDA3
                  else
                     Colors[C_M_USERNM, i+FixedRows] := clWhite;

               end;
            end;
         end;


         // RowCount 정리
         asg_Master.RowCount := asg_Master.RowCount - 1;

         // Comments
         lb_Network.Caption := '▶ ' + IntToStr(iRowCnt) + '건의 담당자정보를 조회하였습니다.';


      end
      else if (in_SearchFlag = 'DOC') or
              (in_SearchFlag = 'DOCREFRESH') then
      begin
         with asg_SelDoc do
         begin
            iRowCnt  := TpGetSearch.RowCount;
            //RowCount := iRowCnt + FixedRows + 1;


            // Merge 변수
            iBaseRow       := 0;    // 대분류 Merge Start Index
            iRowSpan       := 1;    // 대분류 Span  Index
            jBaseRow       := 0;    // 중분류 Merge Start Index
            jRowSpan       := 1;    // 중분류 Span  Index
            kBaseRow       := 0;    // 소분류 Merge Start Index
            kRowSpan       := 1;    // 소분류 Span  Index

            sPrevUptitle    := '';   // 대분류 이전 Title
            sPrevMidtitle   := '';   // 중분류 이전 Title
            sPrevDowntitle  := '';   // 소분류 이전 Title


            for i := 0 to iRowCnt - 1 do
            begin


               Cells[C_SD_DOCLIST, RowCount - 1] := TpGetSearch.GetOutputDataS('sDocList',  i);
               Cells[C_SD_DOCSEQ , RowCount - 1] := TpGetSearch.GetOutputDataS('sDocSeq',   i);
               Cells[C_SD_DOCTITLE,RowCount - 1] := TpGetSearch.GetOutputDataS('sDocTitle', i);
               Cells[C_SD_REGDATE, RowCount - 1] := TpGetSearch.GetOutputDataS('sRegDate',  i);
               Cells[C_SD_REGUSER, RowCount - 1] := TpGetSearch.GetOutputDataS('sRegUser',  i);
               Cells[C_SD_RELDEPT, RowCount - 1] := TpGetSearch.GetOutputDataS('sRelDept',  i);
               Cells[C_SD_DOCRMK,  RowCount - 1] := TpGetSearch.GetOutputDataS('sDocRmk',   i);
               Cells[C_SD_DOCLOC,  RowCount - 1] := TpGetSearch.GetOutputDataS('sDocLoc',   i);
               Cells[C_SD_HQREMARK,RowCount - 1] := TpGetSearch.GetOutputDataS('sAttachNm', i);
               Cells[C_SD_AAREMARK,RowCount - 1] := TpGetSearch.GetOutputDataS('sHideFile', i);
               Cells[C_SD_GRREMARK,RowCount - 1] := TpGetSearch.GetOutputDataS('sDeptNm',   i);
               Cells[C_SD_ASREMARK,RowCount - 1] := TpGetSearch.GetOutputDataS('sDeptSpec', i);
               Cells[C_SD_TOTRMK,  RowCount - 1] := TpGetSearch.GetOutputDataS('sDutyRmk',  i);

               // 계약만료된 항목은 FontColor = clRed 업데이트 @ 2014.06.02 LSH
               if (Cells[C_SD_DOCLIST, RowCount - 1] = '계약처') then
               begin
                  if LengthByte(Trim(CopyByte(Cells[C_SD_DOCRMK,  RowCount - 1], 19, 2))) > 1 then
                  begin
                     if ((CopyByte(Cells[C_SD_DOCRMK,  RowCount - 1], 14, 4) +
                          CopyByte(Cells[C_SD_DOCRMK,  RowCount - 1], 19, 2) +
                          CopyByte(Cells[C_SD_DOCRMK,  RowCount - 1], 23, 4)) < FormatDateTime('yyyymmdd', Date)) then
                     begin
                        FontColors[C_SD_DOCTITLE,  RowCount - 1] := clRed;
                        FontColors[C_SD_DOCRMK,    RowCount - 1] := clRed;
                     end;
                  end
                  else
                  begin
                     if ((CopyByte(Cells[C_SD_DOCRMK,  RowCount - 1], 14, 4) + '0' +
                          Trim(CopyByte(Cells[C_SD_DOCRMK,  RowCount - 1], 19, 2)) +
                          CopyByte(Cells[C_SD_DOCRMK,  RowCount - 1], 23, 4)) < FormatDateTime('yyyymmdd', Date)) then
                     begin
                        FontColors[C_SD_DOCTITLE,  RowCount - 1] := clRed;
                        FontColors[C_SD_DOCRMK,    RowCount - 1] := clRed;
                     end;
                  end;
               end;


               AutoSizeRow(RowCount - 1);


               // 해당 화면 1-Cycle 코드의 마지막 Index를 최종 RowCount 로 사용
               if (i = iRowCnt) then
                  iLstRowCnt := i+1
               else
                  iLstRowCnt := iRowCnt;




               // 대분류 항목 MergeCell 체크
               if TpGetSearch.GetOutputDataS('sDocList', i) <> sPrevUptitle then
               begin
                  // test
                  //Memo1.Lines.Add('--- Test Logging Start ---');


                  // 이전 대분류와 다른 경우
                  if i > 0 then
                  begin
                     MergeCells(0, iBaseRow, 1, iRowSpan);
                     iRowSpan := 1;
                  end;

                  // ★ Key Point
                  iBaseRow := RowCount - 1;


                  // 중분류 항목 MergeCell 체크

                     // 이전 중분류와 다른 경우
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;


                     // ★ Key Point
                     jBaseRow := RowCount - 1;



                     // 소분류 항목 MergeCell 체크
                     if TpGetSearch.GetOutputDataS('sDocTitle', i) <> sPrevDownTitle then
                     begin
                        // 이전 소분류와 다른 경우
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // ★ Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // 이전 소분류와 다른 경우
                        Inc(kRowSpan);

               end
               else
               begin

                  // 중분류 항목 MergeCell 체크
                  if TpGetSearch.GetOutputDataS('sDocSeq', i) <> sPrevMidTitle then
                  begin
                     // 이전 중분류와 다른 경우
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;

                     // ★ Key Point
                     jBaseRow := RowCount - 1;


                     // 소분류 항목 MergeCell 체크
                     if TpGetSearch.GetOutputDataS('sDocTitle', i) <> sPrevDownTitle then
                     begin
                        // 이전 소분류와 같은 경우
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // ★ Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // 이전 소분류와 다른 경우
                        Inc(kRowSpan);
                  end
                  else
                  begin
                     // 중분류 span Index ++
                     Inc(jRowSpan);


                     // 소분류 항목 MergeCell 체크
                     if TpGetSearch.GetOutputDataS('sDocTitle', i) <> sPrevDownTitle then
                     begin
                        // 이전 소분류와 다른 경우
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // ★ Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // 이전 소분류와 동일한 경우
                        Inc(kRowSpan);

                  end;


                  // 대분류 Span index ++
                  Inc(iRowSpan);

               end;


               // Set previous value
               sPrevUptitle   := TpGetSearch.GetOutputDataS('sDocList',   i);
               sPrevmidtitle  := TpGetSearch.GetOutputDataS('sDocSeq',   i);
               sPrevDowntitle := TpGetSearch.GetOutputDataS('sDocTitle', i);


               // RowCount 다음줄로 넘김.
               RowCount := RowCount + 1;


            end;


            // ★ 최종 Code List-up후 마지막 Row Merge 처리
            asg_SelDoc.MergeCells(0, iBaseRow, 1, iRowSpan);
            asg_SelDoc.MergeCells(1, jBaseRow, 1, jRowSpan);
            asg_SelDoc.MergeCells(2, kBaseRow, 1, kRowSpan);
            asg_SelDoc.RowCount := iLstRowCnt + 1;


            {
               Cells[C_SD_DOCLIST, i+FixedRows] := TpGetSearch.GetOutputDataS('sDocList',  i);
               Cells[C_SD_DOCSEQ , i+FixedRows] := TpGetSearch.GetOutputDataS('sDocSeq',   i);
               Cells[C_SD_DOCTITLE,i+FixedRows] := TpGetSearch.GetOutputDataS('sDocTitle', i);
               Cells[C_SD_REGDATE, i+FixedRows] := TpGetSearch.GetOutputDataS('sRegDate',  i);
               Cells[C_SD_REGUSER, i+FixedRows] := TpGetSearch.GetOutputDataS('sRegUser',  i);
               Cells[C_SD_RELDEPT, i+FixedRows] := TpGetSearch.GetOutputDataS('sRelDept',  i);
               Cells[C_SD_DOCRMK,  i+FixedRows] := TpGetSearch.GetOutputDataS('sDocRmk',   i);
               Cells[C_SD_DOCLOC,  i+FixedRows] := TpGetSearch.GetOutputDataS('sDocLoc',   i);
               Cells[C_SD_HQREMARK,i+FixedRows] := TpGetSearch.GetOutputDataS('sAttachNm', i);
               Cells[C_SD_AAREMARK,i+FixedRows] := TpGetSearch.GetOutputDataS('sHideFile', i);
               Cells[C_SD_GRREMARK,i+FixedRows] := TpGetSearch.GetOutputDataS('sDeptNm',   i);
               Cells[C_SD_ASREMARK,i+FixedRows] := TpGetSearch.GetOutputDataS('sDeptSpec', i);
               Cells[C_SD_TOTRMK,  i+FixedRows] := TpGetSearch.GetOutputDataS('sDutyRmk',  i);

               AutoSizeRow(i+FixedRows);
            end;

         // RowCount 정리
         asg_SelDoc.RowCount := asg_SelDoc.RowCount - 1;
         }

         end;


         // Comments
         if (asg_RegDoc.Cells[C_RD_DOCYEAR, asg_RegDoc.Row] = '') or
            (asg_RegDoc.Cells[C_RD_DOCLIST, asg_RegDoc.Row] = '') then
            lb_SelDoc.Caption := '▶ ' + IntToStr(iRowCnt) + '건의 현 회계년도 전체 내역을 검색하였습니다.'
         else
            lb_SelDoc.Caption := '▶ ' + IntToStr(iRowCnt) + '건의 [' + asg_RegDoc.Cells[C_RD_DOCYEAR, asg_RegDoc.Row] + ' ' +
                                                                   asg_RegDoc.Cells[C_RD_DOCLIST, asg_RegDoc.Row] + '] ' +
                                                             '내역을 검색하였습니다.';

      end
      else if (in_SearchFlag = 'BOARD') or
              (in_SearchFlag = 'BOARDFIND') then
      begin
         with asg_Board do
         begin
            iRowCnt  := TpGetSearch.RowCount;
            RowCount := iRowCnt + FixedRows + 1;


            for i := 0 to iRowCnt - 1 do
            begin
               for j := 0 to ColCount - 1 do
               begin
                  Cells[C_B_BOARDSEQ,  i+FixedRows] := TpGetSearch.GetOutputDataS('sBoardSeq', i);
                  Cells[C_B_CATEUP,    i+FixedRows] := TpGetSearch.GetOutputDataS('sCateUp'  , i);
                  Cells[C_B_TITLE,     i+FixedRows] := TpGetSearch.GetOutputDataS('sDocTitle', i);
                  Cells[C_B_REGDATE,   i+FixedRows] := TpGetSearch.GetOutputDataS('sRegDate' , i);
                  Cells[C_B_REGUSER,   i+FixedRows] := TpGetSearch.GetOutputDataS('sRegUser' , i);
                  Cells[C_B_LIKE,      i+FixedRows] := TpGetSearch.GetOutputDataS('sLikeCnt' , i);
                  Cells[C_B_ATTACH,    i+FixedRows] := TpGetSearch.GetOutputDataS('sAttachNm', i);
                  Cells[C_B_HEADSEQ,   i+FixedRows] := TpGetSearch.GetOutputDataS('sHeadSeq' , i);
                  Cells[C_B_TAILSEQ,   i+FixedRows] := TpGetSearch.GetOutputDataS('sTailSeq' , i);
                  Cells[C_B_CONTEXT,   i+FixedRows] := TpGetSearch.GetOutputDataS('sContext' , i);
                  Cells[C_B_USERIP ,   i+FixedRows] := TpGetSearch.GetOutputDataS('sUserIp'  , i);
                  Cells[C_B_REPLY,     i+FixedRows] := TpGetSearch.GetOutputDataS('sReplyCnt', i);
                  Cells[C_B_HEADTAIL,  i+FixedRows] := TpGetSearch.GetOutputDataS('sHeadTail', i);
                  Cells[C_B_HIDEFILE,  i+FixedRows] := TpGetSearch.GetOutputDataS('sHideFile', i);
                  Cells[C_B_SERVERIP,  i+FixedRows] := TpGetSearch.GetOutputDataS('sServerIp', i);

                  //------------------------------------------------------------
                  // Reply 카운트 표기
                  //------------------------------------------------------------
                  if TpGetSearch.GetOutputDataS('sReplyCnt', i) > '0' then
                     Cells[C_B_TITLE,  i+FixedRows] := TpGetSearch.GetOutputDataS('sDocTitle', i) + '  [' + TpGetSearch.GetOutputDataS('sReplyCnt', i) + ']'
                  else
                     Cells[C_B_TITLE,  i+FixedRows] := TpGetSearch.GetOutputDataS('sDocTitle', i);


                  //------------------------------------------------------------
                  // 첨부 파일 표기
                  //------------------------------------------------------------
                  if TpGetSearch.GetOutputDataS('sAttachNm', i) <> '' then
                     AddButton(C_B_ATTACH,  i+FixedRows, ColWidths[C_B_ATTACH]-5,  20, 'Λ'{'⊥','∨'}, haBeforeText, vaCenter);


                  //------------------------------------------------------------
                  // 제일 마지막 Cell Index 세팅 (댓글 보기 기능위해 Set)
                  //------------------------------------------------------------
                  if (i = iRowCnt - 1) then
                  begin
                     Cells[C_B_CLOSEYN,   i+FixedRows] := 'E';
                     iNowBoardCnt := i + 1;
                  end
                  else
                     Cells[C_B_CLOSEYN,   i+FixedRows] := 'C';


                  //------------------------------------------------------------
                  // 공지사항 Coloring
                  //------------------------------------------------------------
                  if (Cells[C_B_HEADTAIL,  i+FixedRows] = 'A') then
                     for k := C_B_BOARDSEQ to C_B_ATTACH do
                        asg_Board.Colors[j, i+FixedRows]   := $00BDABFF;


                  //------------------------------------------------------------
                  // Best 추천글 Coloring
                  //------------------------------------------------------------
                  if PosByte('[★BEST★]', Cells[C_B_TITLE,  i+FixedRows]) > 0 then
                     for k := C_B_BOARDSEQ to C_B_ATTACH do
                        asg_Board.Colors[j, i+FixedRows]   := $00D5D7AF;

               end;
            end;
         end;


         // RowCount 정리
         asg_Board.RowCount := asg_Board.RowCount - 1;

         // Comments
         lb_Board.Caption := '▶ ' + IntToStr(iRowCnt) + '건의 게시글을 조회하였습니다.';


      end
      else if (in_SearchFlag = 'BOARDPAGE') then
      begin
         iRowCnt  := TpGetSearch.RowCount;
         ePageCnt := iRowCnt / C_PAGE_MAXROW_UNIT;


         if (iRowCnt mod C_PAGE_MAXROW_UNIT) > 0 then
            ePageCnt := ePageCnt + 1;

         fcb_Page.Items.Clear;

         for i := 1 to Trunc(ePageCnt) do
            fcb_Page.Items.Add(IntToStr(i));

         fcb_Page.Text := '1';

         // 앞/뒤 게시글 Paging 검색 버튼
         if (fcb_Page.Items.Count > 1) then
         begin
            fsbt_ForWard.Visible  := False;
            fsbt_BackWard.Visible := True;
         end
         else
         begin
            fsbt_ForWard.Visible  := False;
            fsbt_BackWard.Visible := False;
         end;
      end
      // --> ChatUser XE7 개선전까지 사용안함..
      else if (in_SearchFlag = 'CHATUSER') then
      begin
         {  -- XE7 개선전까지 임시주석 @ 2017.10.31 LSH
         with SelChatUser.asg_ChatList do
         begin
            iRowCnt  := TpGetSearch.RowCount;
            RowCount := iRowCnt + FixedRows + 1;


            for i := 0 to iRowCnt - 1 do
            begin
               for j := 0 to ColCount - 1 do
               begin
                  Cells[0,  i+FixedRows] := TpGetSearch.GetOutputDataS('sUserNm'    , i);
                  Cells[1,  i+FixedRows] := TpGetSearch.GetOutputDataS('sUserIp'    , i);


                  // 접속 서버의 IP 및 Port 정보
                  if (Cells[1, i+FixedRows] <> '') then
                  begin
                     sTargetServerIp    := Cells[1, i+FixedRows];
                     iTargetServerPort  := StrToInt(DelChar(CopyByte(Cells[1, i+FixedRows], LengthByte(Cells[1, i+FixedRows])-3, 4), '.'));


                     //---------------------------------------------------------
                     // 접속하려는 Server Port 사용가능 여부 Check
                     //---------------------------------------------------------
                     if (PortTCPIsOpen(iTargetServerPort, sTargetServerIp)) then
                        Colors[0, i+FixedRows] := clGreen
                     else
                        Colors[0, i+FixedRows] := clSilver;

                  end
                  else
                     Colors[0, i+FixedRows] := clSilver;

               end;
            end;

            // RowCount 정리
            RowCount := RowCount - 1;

         end;
         }
      end
      else if (in_SearchFlag = 'DIALLIST') then
      begin
          with asg_DialList do
          begin
            iRowCnt  := TpGetSearch.RowCount;
            RowCount := iRowCnt + FixedRows + 1;


            for i := 0 to iRowCnt - 1 do
            begin
               for j := 0 to ColCount - 1 do
               begin
                  // 이중 키워드 매칭 적용 @ 2015.06.03 LSH
                  if (PosByte(AnsiUpperCase(TokenStr(fed_Scan.Text + ' ', ' ', 1)), AnsiUpperCase(TpGetSearch.GetOutputDataS('sLocate', i))) > 0) or
                     (PosByte(AnsiUpperCase(TokenStr(fed_Scan.Text + ' ', ' ', 2)), AnsiUpperCase(TpGetSearch.GetOutputDataS('sLocate', i))) > 0) then
                  begin
                     Cells[C_DL_LOCATE,    i+FixedRows]                 := TpGetSearch.GetOutputDataS('sLocate', i);
                     asg_DialList.FontStyles[C_DL_LOCATE, i+FixedRows]  := [fsBold];
                  end
                  else
                  begin
                     Cells[C_DL_LOCATE,    i+FixedRows]                 := TpGetSearch.GetOutputDataS('sLocate', i);
                     asg_DialList.FontStyles[C_DL_LOCATE, i+FixedRows]  := [];
                  end;

                  if (PosByte(AnsiUpperCase(TokenStr(fed_Scan.Text + ' ', ' ', 1)), AnsiUpperCase(TpGetSearch.GetOutputDataS('sDeptNm', i))) > 0) or
                     (PosByte(AnsiUpperCase(TokenStr(fed_Scan.Text + ' ', ' ', 2)), AnsiUpperCase(TpGetSearch.GetOutputDataS('sDeptNm', i))) > 0) then
                  begin
                     Cells[C_DL_DEPTNM,    i+FixedRows]                 := TpGetSearch.GetOutputDataS('sDeptNm', i);
                     asg_DialList.FontStyles[C_DL_DEPTNM, i+FixedRows]  := [fsBold];
                  end
                  else
                  begin
                     Cells[C_DL_DEPTNM,    i+FixedRows]                 := TpGetSearch.GetOutputDataS('sDeptNm', i);
                     asg_DialList.FontStyles[C_DL_DEPTNM, i+FixedRows]  := [];
                  end;

                  if (PosByte(AnsiUpperCase(TokenStr(fed_Scan.Text + ' ', ' ', 1)), AnsiUpperCase(TpGetSearch.GetOutputDataS('sDeptSpec', i))) > 0) or
                     (PosByte(AnsiUpperCase(TokenStr(fed_Scan.Text + ' ', ' ', 2)), AnsiUpperCase(TpGetSearch.GetOutputDataS('sDeptSpec', i))) > 0) then
                  begin
                     Cells[C_DL_DEPTSPEC,  i+FixedRows]                    := TpGetSearch.GetOutputDataS('sDeptSpec', i);
                     asg_DialList.FontStyles[C_DL_DEPTSPEC, i+FixedRows]   := [fsBold];
                  end
                  else
                  begin
                     Cells[C_DL_DEPTSPEC,  i+FixedRows]                    := TpGetSearch.GetOutputDataS('sDeptSpec', i);
                     asg_DialList.FontStyles[C_DL_DEPTSPEC, i+FixedRows]   := [];
                  end;

                  if (PosByte(AnsiUpperCase(TokenStr(fed_Scan.Text + ' ', ' ', 1)), AnsiUpperCase(TpGetSearch.GetOutputDataS('sCallNo', i))) > 0) or
                     (PosByte(AnsiUpperCase(TokenStr(fed_Scan.Text + ' ', ' ', 2)), AnsiUpperCase(TpGetSearch.GetOutputDataS('sCallNo', i))) > 0) then
                  begin
                     Cells[C_DL_CALLNO,    i+FixedRows]                 := TpGetSearch.GetOutputDataS('sCallNo', i);
                     asg_DialList.FontStyles[C_DL_CALLNO, i+FixedRows]  := [fsBold];
                  end
                  else
                  begin
                     Cells[C_DL_CALLNO,    i+FixedRows]                 := TpGetSearch.GetOutputDataS('sCallNo', i);
                     asg_DialList.FontStyles[C_DL_CALLNO, i+FixedRows]  := [];
                  end;

                  if (PosByte(AnsiUpperCase(TokenStr(fed_Scan.Text + ' ', ' ', 1)), AnsiUpperCase(TpGetSearch.GetOutputDataS('sDutyUser', i))) > 0) or
                     (PosByte(AnsiUpperCase(TokenStr(fed_Scan.Text + ' ', ' ', 2)), AnsiUpperCase(TpGetSearch.GetOutputDataS('sDutyUser', i))) > 0) then
                  begin
                     Cells[C_DL_DUTYUSER,  i+FixedRows]                    := TpGetSearch.GetOutputDataS('sDutyUser'  , i);
                     asg_DialList.FontStyles[C_DL_DUTYUSER, i+FixedRows]   := [fsBold];
                  end
                  else
                  begin
                     Cells[C_DL_DUTYUSER,  i+FixedRows]                    := TpGetSearch.GetOutputDataS('sDutyUser', i);
                     asg_DialList.FontStyles[C_DL_DUTYUSER, i+FixedRows]   := [];
                  end;

                  if (PosByte(AnsiUpperCase(TokenStr(fed_Scan.Text + ' ', ' ', 1)), AnsiUpperCase(TpGetSearch.GetOutputDataS('sDataBase', i))) > 0) or
                     (PosByte(AnsiUpperCase(TokenStr(fed_Scan.Text + ' ', ' ', 2)), AnsiUpperCase(TpGetSearch.GetOutputDataS('sDataBase', i))) > 0) then
                  begin
                     Cells[C_DL_LINKDB,  i+FixedRows]                    := TpGetSearch.GetOutputDataS('sDataBase', i);
                     asg_DialList.FontStyles[C_DL_LINKDB, i+FixedRows]   := [fsBold];
                  end
                  else
                  begin
                     Cells[C_DL_LINKDB,  i+FixedRows]                    := TpGetSearch.GetOutputDataS('sDataBase', i);
                     asg_DialList.FontStyles[C_DL_LINKDB, i+FixedRows]   := [];
                  end;

                  if (PosByte(AnsiUpperCase(TokenStr(fed_Scan.Text + ' ', ' ', 1)), AnsiUpperCase(TpGetSearch.GetOutputDataS('sLinkCnt', i))) > 0) or
                     (PosByte(AnsiUpperCase(TokenStr(fed_Scan.Text + ' ', ' ', 2)), AnsiUpperCase(TpGetSearch.GetOutputDataS('sLinkCnt', i))) > 0) then
                  begin
                     Cells[C_DL_LINKCNT,     i+FixedRows]                 := TpGetSearch.GetOutputDataS('sLinkCnt'   , i);
                     asg_DialList.FontStyles[C_DL_LINKCNT, i+FixedRows]   := [fsBold];
                  end
                  else
                  begin
                     Cells[C_DL_LINKCNT,     i+FixedRows]                 := TpGetSearch.GetOutputDataS('sLinkCnt'   , i);
                     asg_DialList.FontStyles[C_DL_LINKCNT, i+FixedRows]   := [];
                  end;


                  Cells[C_DL_LINKSEQ,     i+FixedRows]   := TpGetSearch.GetOutputDataS('sLinkSeq', i);
                  Cells[C_DL_DTYUSRID,    i+FixedRows]   := TpGetSearch.GetOutputDataS('sUserId',  i);   // 사용자 ID 추가 @ 2014.06.26 LSH


               end;
            end;


            // RowCount 정리
            RowCount := RowCount - 1;


            // Comments
            lb_DialScan.Caption  := '▶ ' + IntToStr(iRowCnt) + '건의 연락처 정보를 조회하였습니다.';


         end;
      end
      else if (in_SearchFlag = 'MYDIAL') then
      begin
          with asg_MyDial do
          begin
            iRowCnt  := TpGetSearch.RowCount;
            //RowCount := iRowCnt + FixedRows + 1;


            // Merge 변수
            iBaseRow       := 0;    // 대분류 Merge Start Index
            iRowSpan       := 1;    // 대분류 Span  Index
            jBaseRow       := 0;    // 중분류 Merge Start Index
            jRowSpan       := 1;    // 중분류 Span  Index
            kBaseRow       := 0;    // 소분류 Merge Start Index
            kRowSpan       := 1;    // 소분류 Span  Index

            sPrevUptitle    := '';   // 대분류 이전 Title
            sPrevMidtitle   := '';   // 중분류 이전 Title
            sPrevDowntitle  := '';   // 소분류 이전 Title


            for i := 0 to iRowCnt - 1 do
            begin
               Cells[C_MD_LOCATE,    RowCount - 1]  := TpGetSearch.GetOutputDataS('sLocate',     i);
               Cells[C_MD_DEPTNM,    RowCount - 1]  := TpGetSearch.GetOutputDataS('sDeptNm',     i);
               Cells[C_MD_DEPTSPEC,  RowCount - 1]  := TpGetSearch.GetOutputDataS('sDeptSpec',   i);
               Cells[C_MD_CALLNO,    RowCount - 1]  := TpGetSearch.GetOutputDataS('sCallNo',     i);
               Cells[C_MD_DUTYUSER,  RowCount - 1]  := TpGetSearch.GetOutputDataS('sDutyUser',   i);
               Cells[C_MD_MOBILE,    RowCount - 1]  := TpGetSearch.GetOutputDataS('sMobile'  ,   i);
               Cells[C_MD_REMARK,    RowCount - 1]  := TpGetSearch.GetOutputDataS('sDutyRmk',    i);
               Cells[C_MD_LINKDB,    RowCount - 1]  := TpGetSearch.GetOutputDataS('sDataBase',   i);
               Cells[C_MD_LINKSEQ,   RowCount - 1]  := TpGetSearch.GetOutputDataS('sLinkSeq',    i);
               Cells[C_MD_DTYUSRID,  RowCount - 1]  := TpGetSearch.GetOutputDataS('sUserId',     i);  // 등록된 UserId 정보 @ 2015.04.13 LSH



               //AutoSizeRow(RowCount - 1);



               // 해당 화면 1-Cycle 코드의 마지막 Index를 최종 RowCount 로 사용
               if (i = iRowCnt) then
                  iLstRowCnt := i+1
               else
                  iLstRowCnt := iRowCnt;




               // 대분류 항목 MergeCell 체크
               if TpGetSearch.GetOutputDataS('sLocate', i) <> sPrevUptitle then
               begin
                  // test
                  //Memo1.Lines.Add('--- Test Logging Start ---');


                  // 이전 대분류와 다른 경우
                  if i > 0 then
                  begin
                     MergeCells(0, iBaseRow, 1, iRowSpan);
                     iRowSpan := 1;
                  end;

                  // ★ Key Point
                  iBaseRow := RowCount - 1;


                  // 중분류 항목 MergeCell 체크

                  { -- 대분류가 상위와 다른경우, 중분류가 같더라도, 구분 Grid line을 그리기 위해 주석처리 @ 2013.10.14 LSH
                  if TpGetSearch.GetOutputDataS('sDutyRmk', i) <> sPrevMidTitle then
                  begin
                  }
                     // 이전 중분류와 다른 경우
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;


                     // ★ Key Point
                     jBaseRow := RowCount - 1;



                     // 소분류 항목 MergeCell 체크
                     if TpGetSearch.GetOutputDataS('sDeptSpec', i) <> sPrevDownTitle then
                     begin
                        // 이전 소분류와 다른 경우
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // ★ Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // 이전 소분류와 다른 경우
                        Inc(kRowSpan);

                  { -- 대분류가 상위와 다른경우, 중분류가 같더라도, 구분 Grid line을 그리기 위해 주석처리 @ 2013.10.14 LSH
                  end
                  else
                  begin
                     // 이전 중분류와 같은 경우
                     Inc(jRowSpan);
                  end;
                  }
               end
               else
               begin

                  // 중분류 항목 MergeCell 체크
                  if TpGetSearch.GetOutputDataS('sDeptNm', i) <> sPrevMidTitle then
                  begin
                     // 이전 중분류와 다른 경우
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;

                     // ★ Key Point
                     jBaseRow := RowCount - 1;


                     // 소분류 항목 MergeCell 체크
                     if TpGetSearch.GetOutputDataS('sDeptSpec', i) <> sPrevDownTitle then
                     begin
                        // 이전 소분류와 같은 경우
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // ★ Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // 이전 소분류와 다른 경우
                        Inc(kRowSpan);
                  end
                  else
                  begin
                     // 중분류 span Index ++
                     Inc(jRowSpan);


                     // 소분류 항목 MergeCell 체크
                     if TpGetSearch.GetOutputDataS('sDeptSpec', i) <> sPrevDownTitle then
                     begin
                        // 이전 소분류와 다른 경우
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // ★ Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // 이전 소분류와 동일한 경우
                        Inc(kRowSpan);

                  end;


                  // 대분류 Span index ++
                  Inc(iRowSpan);

               end;


               // Set previous value
               sPrevUptitle   := TpGetSearch.GetOutputDataS('sLocate',   i);
               sPrevmidtitle  := TpGetSearch.GetOutputDataS('sDeptNm',   i);
               sPrevDowntitle := TpGetSearch.GetOutputDataS('sDeptSpec', i);


               // RowCount 다음줄로 넘김.
               RowCount := RowCount + 1;


            end;


            // ★ 최종 Code List-up후 마지막 Row Merge 처리
            asg_MyDial.MergeCells(0, iBaseRow, 1, iRowSpan);
            asg_MyDial.MergeCells(1, jBaseRow, 1, jRowSpan);
            asg_MyDial.MergeCells(2, kBaseRow, 1, kRowSpan);
            asg_MyDial.RowCount := iLstRowCnt + 1;






            {
               Cells[C_MD_LOCATE,    i+FixedRows]  := TpGetSearch.GetOutputDataS('sLocate',     i);
               Cells[C_MD_DEPTNM,    i+FixedRows]  := TpGetSearch.GetOutputDataS('sDeptNm',     i);
               Cells[C_MD_DEPTSPEC,  i+FixedRows]  := TpGetSearch.GetOutputDataS('sDeptSpec',   i);
               Cells[C_MD_CALLNO,    i+FixedRows]  := TpGetSearch.GetOutputDataS('sCallNo',     i);
               Cells[C_MD_DUTYUSER,  i+FixedRows]  := TpGetSearch.GetOutputDataS('sDutyUser',   i);
               Cells[C_MD_MOBILE,    i+FixedRows]  := TpGetSearch.GetOutputDataS('sMobile'  ,   i);
               Cells[C_MD_REMARK,    i+FixedRows]  := TpGetSearch.GetOutputDataS('sDutyRmk',    i);
               Cells[C_MD_LINKDB,    i+FixedRows]  := TpGetSearch.GetOutputDataS('sDataBase',   i);
               Cells[C_MD_LINKSEQ,   i+FixedRows]  := TpGetSearch.GetOutputDataS('sLinkSeq',    i);
            end;


            // RowCount 정리
            RowCount := RowCount - 1;
            }


            // Comments
            lb_MyDial.Caption  := IntToStr(iRowCnt) + '건의 연락처 정보를 조회하였습니다.';


         end;
      end
      else if (in_SearchFlag = 'DIALSCAN') then
      begin
         with asg_DialScan do
         begin
            iRowCnt  := TpGetSearch.RowCount;
            RowCount := iRowCnt + FixedRows + 1;


            for i := 0 to iRowCnt - 1 do
            begin
               for j := 0 to ColCount - 1 do
               begin
                  Cells[C_DS_KEYWORD,    i+FixedRows]  := TpGetSearch.GetOutputDataS('sContext', i);
               end;
            end;


            // RowCount 정리
            RowCount := RowCount - 1;

         end;
      end
      else if (in_SearchFlag = 'ANALYSIS') then
      begin
         with asg_Analysis do
         begin
            iRowCnt  := TpGetSearch.RowCount;
            RowCount := iRowCnt + FixedRows + 1;


            for i := 0 to iRowCnt - 1 do
            begin

               Cells[C_AN_REQFLAG,    i+FixedRows]  := TpGetSearch.GetOutputDataS('sDeptSpec',    i);

               if (TpGetSearch.GetOutputDataS('sHeadTail',   i) = 'A') then
                  Cells[C_AN_SRTITLE,    i+FixedRows]  := '[원무]' + TpGetSearch.GetOutputDataS('sDocTitle',   i)
               else if (TpGetSearch.GetOutputDataS('sHeadTail',   i) = 'M') then
                  Cells[C_AN_SRTITLE,    i+FixedRows]  := '[진료]' + TpGetSearch.GetOutputDataS('sDocTitle',   i)
               else if (TpGetSearch.GetOutputDataS('sHeadTail',   i) = 'S') then
                  Cells[C_AN_SRTITLE,    i+FixedRows]  := '[진지]' + TpGetSearch.GetOutputDataS('sDocTitle',   i)
               else if (TpGetSearch.GetOutputDataS('sHeadTail',   i) = 'C') then
                  Cells[C_AN_SRTITLE,    i+FixedRows]  := '[공통]' + TpGetSearch.GetOutputDataS('sDocTitle',   i)
               else if (TpGetSearch.GetOutputDataS('sHeadTail',   i) = 'G') then
                  Cells[C_AN_SRTITLE,    i+FixedRows]  := '[일반]' + TpGetSearch.GetOutputDataS('sDocTitle',   i)
               else if (TpGetSearch.GetOutputDataS('sHeadTail',   i) = 'H') then
                  Cells[C_AN_SRTITLE,    i+FixedRows]  := '[HRM]'  + TpGetSearch.GetOutputDataS('sDocTitle',   i)
               else if (TpGetSearch.GetOutputDataS('sHeadTail',   i) = 'E') then
                  Cells[C_AN_SRTITLE,    i+FixedRows]  := '[e-Biz]'+ TpGetSearch.GetOutputDataS('sDocTitle',   i);



               Cells[C_AN_REQDEPT,    i+FixedRows]  := TpGetSearch.GetOutputDataS('sRelDept',    i);
               Cells[C_AN_REQUSER,    i+FixedRows]  := TpGetSearch.GetOutputDataS('sUserNm',     i);
               Cells[C_AN_REQDATE,    i+FixedRows]  := TpGetSearch.GetOutputDataS('sRegDate',    i);
               Cells[C_AN_DUTYUSER,   i+FixedRows]  := TpGetSearch.GetOutputDataS('sDutyUser',   i);
               Cells[C_AN_PROCESS,    i+FixedRows]  := TpGetSearch.GetOutputDataS('sDeptNm',     i);


               if (TpGetSearch.GetOutputDataS('sAttachNm',   i) <> '') then
                  Cells[C_AN_FILECNT,  i+FixedRows]  := 'O'
               else
                  Cells[C_AN_FILECNT,  i+FixedRows]  := '';

               {
               Cells[C_AN_REQOBJT,    i+FixedRows]  := TpGetSearch.GetOutputDataS('sReqObjt',    i);
               Cells[C_AN_REQRMK,     i+FixedRows]  := TpGetSearch.GetOutputDataS('sReqRmk',     i);
               Cells[C_AN_ANSRMK,     i+FixedRows]  := TpGetSearch.GetOutputDataS('sAnsRmk',     i);
               Cells[C_AN_REQISSU,    i+FixedRows]  := TpGetSearch.GetOutputDataS('sReqIssu',    i);
               }


               // Hidden info
               Cells[C_AN_REQNO,      i+FixedRows]  := TpGetSearch.GetOutputDataS('sReqNo',      i);
               Cells[C_AN_ATTACHNM,   i+FixedRows]  := TpGetSearch.GetOutputDataS('sAttachNm',   i);
               Cells[C_AN_USERID,     i+FixedRows]  := TpGetSearch.GetOutputDataS('sUserID',     i);
               Cells[C_AN_REQDEPTNM,  i+FixedRows]  := TpGetSearch.GetOutputDataS('sDutyRmk',    i);
               Cells[C_AN_REQTELNO,   i+FixedRows]  := TpGetSearch.GetOutputDataS('sDutySpec',   i);
               Cells[C_AN_REQSPEC,    i+FixedRows]  := TpGetSearch.GetOutputDataS('sDocRmk',     i);


               if (PosByte('고객승인', Cells[C_AN_PROCESS,    i+FixedRows]) > 0) or
                  (PosByte('요청기각', Cells[C_AN_PROCESS,    i+FixedRows]) > 0) then
               begin
                  FontColors[C_AN_REQFLAG,   i+FixedRows] := clGray;
                  FontColors[C_AN_REQNO,     i+FixedRows] := clGray;
                  FontColors[C_AN_SRTITLE,   i+FixedRows] := clGray;
                  FontColors[C_AN_REQDEPT,   i+FixedRows] := clGray;
                  FontColors[C_AN_REQUSER,   i+FixedRows] := clGray;
                  FontColors[C_AN_REQDATE,   i+FixedRows] := clGray;
                  FontColors[C_AN_DUTYUSER,  i+FixedRows] := clGray;
                  FontColors[C_AN_PROCESS,   i+FixedRows] := clGray;
                  FontColors[C_AN_FILECNT,   i+FixedRows] := clGray;
               end
               else if (PosByte('요청완료',   Cells[C_AN_PROCESS,    i+FixedRows]) > 0) or
                       (PosByte('요청접수',   Cells[C_AN_PROCESS,    i+FixedRows]) > 0) or
                       (PosByte('요청서접수', Cells[C_AN_PROCESS,    i+FixedRows]) > 0) then
                  FontColors[C_AN_PROCESS,  i+FixedRows] := clRed
               else if (PosByte('계획',   Cells[C_AN_PROCESS,    i+FixedRows]) > 0) then
                  FontColors[C_AN_PROCESS,  i+FixedRows] := clBlue
               else if (PosByte('결과',   Cells[C_AN_PROCESS,    i+FixedRows]) > 0) then
                  FontColors[C_AN_PROCESS,  i+FixedRows] := clPurple
               else if (PosByte('입력중',   Cells[C_AN_PROCESS,    i+FixedRows]) > 0) then
                  FontColors[C_AN_PROCESS,  i+FixedRows] := clGreen;

            end;


            // RowCount 정리
            RowCount := RowCount - 1;


            // Comments
            if (iRowCnt = -1) then
               lb_Analysis.Caption  := '▶ ' + GetTxMsg + '.'     // 검색건수 MAXROWCNT 초과 메세지 알림 @ 2016.11.18 LSH
            else
               lb_Analysis.Caption  := '▶ ' + IntToStr(iRowCnt) + '건의 S/R 내역을 조회하였습니다.';

         end;
      end
      else if (in_SearchFlag = 'WORKRPT') then
      begin
         with asg_WorkRpt do
         begin
            iRowCnt  := TpGetSearch.RowCount;
            //RowCount := iRowCnt + FixedRows + 1;


            // Merge 변수
            iBaseRow       := 0;    // 대분류 Merge Start Index
            iRowSpan       := 1;    // 대분류 Span  Index
            jBaseRow       := 0;    // 중분류 Merge Start Index
            jRowSpan       := 1;    // 중분류 Span  Index
            kBaseRow       := 0;    // 소분류 Merge Start Index
            kRowSpan       := 1;    // 소분류 Span  Index

            sPrevUptitle    := '';   // 대분류 이전 Title
            sPrevMidtitle   := '';   // 중분류 이전 Title
            sPrevDowntitle  := '';   // 소분류 이전 Title



            for i := 0 to iRowCnt - 1 do
            begin

               Cells[C_WR_LOCATE,   RowCount - 1] := TpGetSearch.GetOutputDataS('sLocate'    , i);
               Cells[C_WR_DUTYUSER, RowCount - 1] := StringReplace(TpGetSearch.GetOutputDataS('sDutyPart'  , i), #13#10, ' ', [rfReplaceAll]);
               Cells[C_WR_DUTYPART, RowCount - 1] := TpGetSearch.GetOutputDataS('sDutyUser'  , i);
               Cells[C_WR_CONTEXT,  RowCount - 1] := Trim(TpGetSearch.GetOutputDataS('sContext'   , i));


               // 작성일 아래 공휴일 및 주말표기 및 Coloring
               if (Trim(TpGetSearch.GetOutputDataS('sFlag'   , i)) <> 'WEEKEND') and
                  (Trim(TpGetSearch.GetOutputDataS('sFlag'   , i)) <> '') then
               begin
                  Cells[C_WR_REGDATE,  RowCount - 1] := StringReplace(TpGetSearch.GetOutputDataS('sRegDate'   , i) + #13#10 +
                                                                      '(' + GetDayofWeek(StrToDate(TpGetSearch.GetOutputDataS('sRegDate'   , i)), 'HS') + ')' + #13#10 +
                                                                      TpGetSearch.GetOutputDataS('sFlag'      , i), #13#10, ' ', [rfReplaceAll]);
                  Colors[C_WR_REGDATE, RowCount - 1] := $00AF9BE7;
               end
               else if (Trim(TpGetSearch.GetOutputDataS('sFlag'   , i)) = 'WEEKEND') then
               begin
                  Cells[C_WR_REGDATE,  RowCount - 1] := StringReplace(TpGetSearch.GetOutputDataS('sRegDate'   , i) + #13#10 +
                                                                      '(' + GetDayofWeek(StrToDate(TpGetSearch.GetOutputDataS('sRegDate'   , i)), 'HS') + ')', #13#10, ' ', [rfReplaceAll]);
                  Colors[C_WR_REGDATE, RowCount - 1] := $00AFCFAF;
               end
               else if (Trim(TpGetSearch.GetOutputDataS('sFlag'   , i)) = '') then
               begin
                  Cells[C_WR_REGDATE,  RowCount - 1] := StringReplace(TpGetSearch.GetOutputDataS('sRegDate'   , i) + #13#10 +
                                                                     '(' + GetDayofWeek(StrToDate(TpGetSearch.GetOutputDataS('sRegDate'   , i)), 'HS') + ')', #13#10, ' ', [rfReplaceAll]);
                  Colors[C_WR_REGDATE, RowCount - 1] := clWhite;
               end;



               AutoSizeRow(RowCount - 1);



               // 해당 화면 1-Cycle 코드의 마지막 Index를 최종 RowCount 로 사용
               if (i = iRowCnt) then
                  iLstRowCnt := i+1
               else
                  iLstRowCnt := iRowCnt;




               // 대분류 항목 MergeCell 체크
               if TpGetSearch.GetOutputDataS('sLocate', i) <> sPrevUptitle then
               begin
                  // test
                  //Memo1.Lines.Add('--- Test Logging Start ---');


                  // 이전 대분류와 다른 경우
                  if i > 0 then
                  begin
                     MergeCells(0, iBaseRow, 1, iRowSpan);
                     iRowSpan := 1;
                  end;

                  // ★ Key Point
                  iBaseRow := RowCount - 1;


                  // 중분류 항목 MergeCell 체크

                  { -- 대분류가 상위와 다른경우, 중분류가 같더라도, 구분 Grid line을 그리기 위해 주석처리 @ 2013.10.14 LSH
                  if TpGetSearch.GetOutputDataS('sDutyRmk', i) <> sPrevMidTitle then
                  begin
                  }
                     // 이전 중분류와 다른 경우
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;


                     // ★ Key Point
                     jBaseRow := RowCount - 1;



                     // 소분류 항목 MergeCell 체크
                     if TpGetSearch.GetOutputDataS('sDutyUser', i) <> sPrevDownTitle then
                     begin
                        // 이전 소분류와 다른 경우
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // ★ Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // 이전 소분류와 다른 경우
                        Inc(kRowSpan);

                  { -- 대분류가 상위와 다른경우, 중분류가 같더라도, 구분 Grid line을 그리기 위해 주석처리 @ 2013.10.14 LSH
                  end
                  else
                  begin
                     // 이전 중분류와 같은 경우
                     Inc(jRowSpan);
                  end;
                  }
               end
               else
               begin

                  // 중분류 항목 MergeCell 체크
                  if TpGetSearch.GetOutputDataS('sDutyPart', i) <> sPrevMidTitle then
                  begin
                     // 이전 중분류와 다른 경우
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;

                     // ★ Key Point
                     jBaseRow := RowCount - 1;


                     // 소분류 항목 MergeCell 체크
                     if TpGetSearch.GetOutputDataS('sDutyUser', i) <> sPrevDownTitle then
                     begin
                        // 이전 소분류와 같은 경우
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // ★ Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // 이전 소분류와 다른 경우
                        Inc(kRowSpan);
                  end
                  else
                  begin
                     // 중분류 span Index ++
                     Inc(jRowSpan);


                     // 소분류 항목 MergeCell 체크
                     if TpGetSearch.GetOutputDataS('sDutyUser', i) <> sPrevDownTitle then
                     begin
                        // 이전 소분류와 다른 경우
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // ★ Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // 이전 소분류와 동일한 경우
                        Inc(kRowSpan);

                  end;


                  // 대분류 Span index ++
                  Inc(iRowSpan);

               end;


               // Set previous value
               sPrevUptitle   := TpGetSearch.GetOutputDataS('sLocate', i);
               sPrevmidtitle  := TpGetSearch.GetOutputDataS('sDutyPart', i);
               sPrevDowntitle := TpGetSearch.GetOutputDataS('sDutyUser', i);


               // RowCount 다음줄로 넘김.
               RowCount := RowCount + 1;


            end;


            // ★ 최종 Code List-up후 마지막 Row Merge 처리
            asg_WorkRpt.MergeCells(0, iBaseRow, 1, iRowSpan);
            asg_WorkRpt.MergeCells(1, jBaseRow, 1, jRowSpan);
            asg_WorkRpt.MergeCells(2, kBaseRow, 1, kRowSpan);
            asg_WorkRpt.RowCount := iLstRowCnt + 1;


            //------------------------------------------------------------------
            // 로그인 User의 Row에 화면 고정위한 로직 추가 @ 2016.09.02 LSH
            //------------------------------------------------------------------
            for i := 0 to iRowCnt - 1 do
            begin
               if Cells[C_WR_DUTYPART, i] = FsUserNm then
               begin
                  Row := i;

                  Break;
               end;
            end;
         end;




            {
               Cells[C_WR_LOCATE,   i + FixedRows] := TpGetSearch.GetOutputDataS('sLocate'    , i);
               Cells[C_WR_DUTYUSER, i + FixedRows] := TpGetSearch.GetOutputDataS('sDutyUser'  , i);
               Cells[C_WR_DUTYPART, i + FixedRows] := TpGetSearch.GetOutputDataS('sDutyPart'  , i);
               Cells[C_WR_CONTEXT,  i + FixedRows] := TpGetSearch.GetOutputDataS('sContext'   , i);


               // 작성일 아래 공휴일 및 주말표기 및 Coloring
               if (Trim(TpGetSearch.GetOutputDataS('sFlag'   , i)) <> 'WEEKEND') and
                  (Trim(TpGetSearch.GetOutputDataS('sFlag'   , i)) <> '') then
               begin
                  Cells[C_WR_REGDATE,  i + FixedRows] := TpGetSearch.GetOutputDataS('sRegDate'   , i) + #13#10 +
                                                         '(' + GetDayofWeek(StrToDate(TpGetSearch.GetOutputDataS('sRegDate'   , i)), 'HS') + ')' + #13#10 +
                                                         TpGetSearch.GetOutputDataS('sFlag'      , i);
                  Colors[C_WR_REGDATE, i + FixedRows] := $00AF9BE7;
               end
               else if (Trim(TpGetSearch.GetOutputDataS('sFlag'   , i)) = 'WEEKEND') then
               begin
                  Cells[C_WR_REGDATE,  i + FixedRows] := TpGetSearch.GetOutputDataS('sRegDate'   , i) + #13#10 +
                                                         '(' + GetDayofWeek(StrToDate(TpGetSearch.GetOutputDataS('sRegDate'   , i)), 'HS') + ')';
                  Colors[C_WR_REGDATE, i + FixedRows] := $00AFCFAF;
               end
               else if (Trim(TpGetSearch.GetOutputDataS('sFlag'   , i)) = '') then
               begin
                  Cells[C_WR_REGDATE,  i + FixedRows] := TpGetSearch.GetOutputDataS('sRegDate'   , i) + #13#10 +
                                                         '(' + GetDayofWeek(StrToDate(TpGetSearch.GetOutputDataS('sRegDate'   , i)), 'HS') + ')';
                  Colors[C_WR_REGDATE, i + FixedRows] := clWhite;
               end;



               AutoSizeRow(i+FixedRows);

            end;
         end;


         // RowCount 정리
         asg_WorkRpt.RowCount := asg_WorkRpt.RowCount - 1;
         }


         // Comments
         if (iRowCnt = -1) then
            lb_Analysis.Caption := '▶ ' + GetTxMsg + '.'      // 검색건수 MAXROWCNT 초과 알림 @ 2016.11.18 LSH
         else
            lb_Analysis.Caption := '▶ ' + IntToStr(iRowCnt) + '건의 업무일지를 조회하였습니다.';

      end
      else if (in_SearchFlag = 'WEEKLYRPT') then
      begin
         with asg_WeeklyRpt do
         begin
            iRowCnt  := TpGetSearch.RowCount;

            //RowCount := iRowCnt + FixedRows + 1;

            // Merge 변수
            iBaseRow       := 0;    // 대분류 Merge Start Index
            iRowSpan       := 1;    // 대분류 Span  Index
            jBaseRow       := 0;    // 중분류 Merge Start Index
            jRowSpan       := 1;    // 중분류 Span  Index
            kBaseRow       := 0;    // 소분류 Merge Start Index
            kRowSpan       := 1;    // 소분류 Span  Index

            sPrevUptitle    := '';   // 대분류 이전 Title
            sPrevMidtitle   := '';   // 중분류 이전 Title
            sPrevDowntitle  := '';   // 소분류 이전 Title


            for i := 0 to iRowCnt - 1 do
            begin
               Cells[C_WK_DATE,     RowCount - 1] := TpGetSearch.GetOutputDataS('sRegDate'    , i) +
                                                      '(' + GetDayofWeek(StrToDate(TpGetSearch.GetOutputDataS('sRegDate'   , i)), 'HS') + ')';
               Cells[C_WK_GUBUN,    RowCount - 1] := StringReplace(Trim(TpGetSearch.GetOutputDataS('sDutyRmk'    , i)), #13#10, ' ', [rfReplaceAll]);
               Cells[C_WK_CONTEXT,  RowCount - 1] := Trim(TpGetSearch.GetOutputDataS('sContext'    , i));

               // 임시변수 초기화
               tmpWkLine1   := '';
               tmpWkLine2   := '';
               tmpWkLine3   := '';
               tmpWkProc1   := '';
               tmpWkProc2   := '';
               tmpWkProc3   := '';
               tmpWkProcAll := '';

               //---------------------------------------------------------------
               // 업무일지 [진행단계] 표기
               //    - 해당 Cell의 Text를 CRLF 단위로 Parsing 하여
               //      검토/진행/완료의 3단계로 진행상황을 표기
               //---------------------------------------------------------------
               if (Cells[C_WK_GUBUN,    RowCount - 1] = '업무일지') then
               begin
                  // CRLF 단위로 Trim 수행
                  tmpWkLine1 := CopyByte(Trim(Cells[C_WK_CONTEXT,  RowCount - 1]), 1, PosByte(#13#10, Trim(Cells[C_WK_CONTEXT,  RowCount - 1])));
                  tmpWkLine2 := CopyByte(Trim(Cells[C_WK_CONTEXT,  RowCount - 1]), PosByte(#13#10, Trim(Cells[C_WK_CONTEXT,  RowCount - 1])) + 1, LengthByte(Trim(Cells[C_WK_CONTEXT,  RowCount - 1])));
                  tmpWkLine3 := CopyByte(Trim(tmpWkLine2), PosByte(#13#10, Trim(tmpWkLine2)), LengthByte(Trim(tmpWkLine2)));

                  // 다시 Line by Line으로 Trim 수행
                  tmpWkLine2 := CopyByte(Trim(tmpWkLine2), 1, PosByte(#13#10, Trim(tmpWkLine2)) - 1);
                  tmpWkLine3 := CopyByte(Trim(tmpWkLine3), 1, LengthByte(Trim(tmpWkLine3)));


                  if PosByte('진행', Trim(tmpWkLine1)) > 0 then
                     tmpWkProc1 := '진행'
                  else if PosByte('검토', Trim(tmpWkLine1)) > 0 then
                     tmpWkProc1 := '검토'
                  else if Trim(tmpWkLine1) <> '' then
                     tmpWkProc1 := '완료'
                  else
                     tmpWkProc1 := '';

                  if PosByte('진행', Trim(tmpWkLine2)) > 0 then
                     tmpWkProc2 := '진행'
                  else if PosByte('검토', Trim(tmpWkLine2)) > 0 then
                     tmpWkProc2 := '검토'
                  else if Trim(tmpWkLine2) <> '' then
                     tmpWkProc2 := '완료'
                  else
                     tmpWkProc2 := '';

                  if PosByte('진행', Trim(tmpWkLine3)) > 0 then
                     tmpWkProc3 := '진행'
                  else if PosByte('검토', Trim(tmpWkLine3)) > 0 then
                     tmpWkProc3 := '검토'
                  else if Trim(tmpWkLine3) <> '' then
                     tmpWkProc3 := '완료'
                  else
                     tmpWkProc3 := '';


                  // 진행단계 취합
                  if (tmpWkProc1 <> '') then
                     tmpWkProcAll := tmpWkProc1;

                  if (tmpWkProc2 <> '') then
                     tmpWkProcAll := tmpWkProcAll + #13#10 + tmpWkProc2
                  else if (tmpWkProc2 = '') then
                     tmpWkProcAll := tmpWkProcAll;

                  if (tmpWkProc3 <> '') then
                     tmpWkProcAll := tmpWkProcAll + #13#10 + tmpWkProc3
                  else if (tmpWkProc3 = '') then
                     tmpWkProcAll := tmpWkProcAll;



                  Cells[C_WK_STEP,     RowCount - 1] :=  tmpWkProcAll;


               end
               else
                  Cells[C_WK_STEP,     RowCount - 1] := TpGetSearch.GetOutputDataS('sDutySpec'   , i);


               // 진행단계  Coloring
               if PosByte('검토', Cells[C_WK_STEP,        RowCount - 1]) > 0 then
               begin
                  Colors[C_WK_CONTEXT,    RowCount - 1] := $0074D7AF;
                  Colors[C_WK_STEP,       RowCount - 1] := $0074D7AF;

                  AddComment(C_WK_STEP,
                             RowCount - 1,
                             '업무일지에 [검토]라는 단어가 포함된 경우,' + #13#10 +
                             '또는 S/R내역 요청 특이사항에 담당예정자가 적힌 경우에 녹색으로 표기됩니다.');
               end
               else if PosByte('진행', Cells[C_WK_STEP,   RowCount - 1]) > 0 then
               begin
                  Colors[C_WK_CONTEXT,    RowCount - 1] := $0074A9FF;
                  Colors[C_WK_STEP,       RowCount - 1] := $0074A9FF;

                  AddComment(C_WK_STEP,
                             RowCount - 1,
                             '업무일지상 [진행]이라는 단어가 포함된 경우,' + #13#10 +
                             '또는 S/R내역 담당자 지정후 결과승인요청 상태 이전인 경우에 붉은색으로 표기됩니다.');
               end
               else if (PosByte('발송', Cells[C_WK_STEP,   RowCount - 1]) > 0) or
                       (PosByte('협조', Cells[C_WK_STEP,   RowCount - 1]) > 0) or
                       (PosByte('회보', Cells[C_WK_STEP,   RowCount - 1]) > 0) or
                       (PosByte('계약처', Cells[C_WK_STEP, RowCount - 1]) > 0) or
                       (PosByte('OCS',  Cells[C_WK_STEP,   RowCount - 1]) > 0) or
                       (PosByte('CRA/CRC',Cells[C_WK_STEP, RowCount - 1]) > 0) then
               begin
                  Colors[C_WK_CONTEXT,    RowCount - 1] := $00C0D5B5;
                  Colors[C_WK_STEP,       RowCount - 1] := $00C0D5B5;

                  AddComment(C_WK_STEP,
                             RowCount - 1,
                             '문서관리 Tab에서 괄호안 번호를 참조해서 원본문서 보관위치를 확인하실 수 있습니다.');
               end
               else if ((PosByte('완료', Cells[C_WK_STEP,   RowCount - 1]) > 0) and
                        (PosByte('진행', Cells[C_WK_STEP,   RowCount - 1]) = 0) and
                        (PosByte('검토', Cells[C_WK_STEP,   RowCount - 1]) = 0)) or
                        (PosByte('공지', Cells[C_WK_STEP,   RowCount - 1]) > 0) or
                        (PosByte('등록', Cells[C_WK_STEP,   RowCount - 1]) > 0)
                        then
               begin
                  Colors[C_WK_CONTEXT,    RowCount - 1] := $00D5CDC7;
                  Colors[C_WK_STEP,       RowCount - 1] := $00D5CDC7;
               end
               else
               begin
                  Colors[C_WK_CONTEXT,    RowCount - 1] := clWhite;
                  Colors[C_WK_STEP,       RowCount - 1] := clWhite;
               end;


               AutoSizeRow(RowCount - 1);


               // 해당 화면 1-Cycle 코드의 마지막 Index를 최종 RowCount 로 사용
               if (i = iRowCnt) then
                  iLstRowCnt := i+1
               else
                  iLstRowCnt := iRowCnt;




               // 대분류 항목 MergeCell 체크
               if TpGetSearch.GetOutputDataS('sRegDate', i) <> sPrevUptitle then
               begin
                  // test
                  //Memo1.Lines.Add('--- Test Logging Start ---');


                  // 이전 대분류와 다른 경우
                  if i > 0 then
                  begin
                     MergeCells(0, iBaseRow, 1, iRowSpan);
                     iRowSpan := 1;
                  end;

                  // ★ Key Point
                  iBaseRow := RowCount - 1;


                  // 중분류 항목 MergeCell 체크

                  { -- 대분류가 상위와 다른경우, 중분류가 같더라도, 구분 Grid line을 그리기 위해 주석처리 @ 2013.10.14 LSH
                  if TpGetSearch.GetOutputDataS('sDutyRmk', i) <> sPrevMidTitle then
                  begin
                  }
                     // 이전 중분류와 다른 경우
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;


                     // ★ Key Point
                     jBaseRow := RowCount - 1;



                     // 소분류 항목 MergeCell 체크
                     if TpGetSearch.GetOutputDataS('sContext', i) <> sPrevDownTitle then
                     begin
                        // 이전 소분류와 다른 경우
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // ★ Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // 이전 소분류와 다른 경우
                        Inc(kRowSpan);

                  { -- 대분류가 상위와 다른경우, 중분류가 같더라도, 구분 Grid line을 그리기 위해 주석처리 @ 2013.10.14 LSH
                  end
                  else
                  begin
                     // 이전 중분류와 같은 경우
                     Inc(jRowSpan);
                  end;
                  }
               end
               else
               begin

                  // 중분류 항목 MergeCell 체크
                  if TpGetSearch.GetOutputDataS('sDutyRmk', i) <> sPrevMidTitle then
                  begin
                     // 이전 중분류와 다른 경우
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;

                     // ★ Key Point
                     jBaseRow := RowCount - 1;


                     // 소분류 항목 MergeCell 체크
                     if TpGetSearch.GetOutputDataS('sContext', i) <> sPrevDownTitle then
                     begin
                        // 이전 소분류와 같은 경우
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // ★ Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // 이전 소분류와 다른 경우
                        Inc(kRowSpan);
                  end
                  else
                  begin
                     // 중분류 span Index ++
                     Inc(jRowSpan);


                     // 소분류 항목 MergeCell 체크
                     if TpGetSearch.GetOutputDataS('sContext', i) <> sPrevDownTitle then
                     begin
                        // 이전 소분류와 다른 경우
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // ★ Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // 이전 소분류와 동일한 경우
                        Inc(kRowSpan);

                  end;


                  // 대분류 Span index ++
                  Inc(iRowSpan);

               end;


               // Set previous value
               sPrevUptitle   := TpGetSearch.GetOutputDataS('sRegDate', i);
               sPrevmidtitle  := TpGetSearch.GetOutputDataS('sDutyRmk', i);
               sPrevDowntitle := TpGetSearch.GetOutputDataS('sContext', i);


               // RowCount 다음줄로 넘김.
               RowCount := RowCount + 1;


            end;


            // ★ 최종 Code List-up후 마지막 Row Merge 처리
            asg_WeeklyRpt.MergeCells(0, iBaseRow, 1, iRowSpan);
            asg_WeeklyRpt.MergeCells(1, jBaseRow, 1, jRowSpan);
            asg_WeeklyRpt.RowCount := iLstRowCnt + 1;




            { -- Cell-Merge 적용전 Logic
               Cells[C_WK_DATE,     i + FixedRows] := TpGetSearch.GetOutputDataS('sRegDate'    , i) +
                                                      '(' + GetDayofWeek(StrToDate(TpGetSearch.GetOutputDataS('sRegDate'   , i)), 'HS') + ')';
               Cells[C_WK_GUBUN,    i + FixedRows] := TpGetSearch.GetOutputDataS('sDutyRmk'    , i);
               Cells[C_WK_CONTEXT,  i + FixedRows] := TpGetSearch.GetOutputDataS('sContext'    , i);
               Cells[C_WK_STEP,     i + FixedRows] := TpGetSearch.GetOutputDataS('sDutySpec'   , i);



               // 표기 및 Coloring
               if Cells[C_WK_STEP,        i + FixedRows] = '검토' then
               begin
                  Colors[C_WK_STEP,       i + FixedRows] := $00AF9BE7;
               end
               else if Cells[C_WK_STEP,   i + FixedRows] = '진행' then
               begin
                  Colors[C_WR_REGDATE,    i + FixedRows] := $00AFCFAF;
               end
               else if Cells[C_WK_STEP,   i + FixedRows] = '완료' then
               begin
                  Colors[C_WR_REGDATE,   2 i + FixedRows] := clSilver;
               end
               else
                  Colors[C_WR_REGDATE,    i + FixedRows] := clWhite;

               AutoSizeRow(i+FixedRows);


            end;
            }

         end;


         // RowCount 정리
         //asg_WeeklyRpt.RowCount := asg_WeeklyRpt.RowCount - 1;


         // Comments
         lb_Analysis.Caption := '▶ [' + Trim(fed_Analysis.Text) + '] 의 리포팅 ' + IntToStr(iRowCnt) + '건을 작성하였습니다.';

      end
      else if (in_SearchFlag = 'DIALMAP') then
      begin
         iRowCnt  := TpGetSearch.RowCount;

            with asg_DialMap do
            begin

               for i := 0 to iRowCnt - 1 do
               begin
                  // 팀장
                  if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5141') then
                  begin
                     Cells[7, 0]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[7, 0]   := $0093D7AB;
                  end
                  // 협력업체 PM
                  {
                  else if (TpGetSearch.GetOutputDataS('sDutyPart'  , i) = '관리자') and
                          (TpGetSearch.GetOutputDataS('sDutySpec'  , i) = '팀장')   and
                          (TpGetSearch.GetOutputDataS('sCallNo'    , i) <> '5141')  then
                  begin
                     Cells[1, 0]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[1, 0]   := $0093D7AB;
                  end
                  }
                  // 운영파트장
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5142') then
                  begin
                     Cells[5, 0]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[5, 0]   := $0093D7AB;
                  end
                  // 기획파트장
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5861') then
                  begin
                     Cells[3, 0]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[3, 0]   := $0093D7AB;
                  end
                  // 개발파트장 @ 2016.07.26 LSH
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5144') then
                  begin
                     Cells[1, 0]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[1, 0]   := $0093D7AB;
                  end
                  {
                  // 진료지원파트(1)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6214') then
                  begin
                     Cells[0, 1]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[0, 1]   := $00DDD7E9;
                  end
                  }
                  // 진료지원파트(1) --> 원무(1)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6225') then
                  begin
                     Cells[0, 1]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[0, 1]   := $00DDD7E9;
                  end
                  // 진료지원파트(2) --> 원무(2)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6213') then
                  begin
                     Cells[1, 1]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[1, 1]   := $00DDD7E9;
                  end
                  // 진료지원파트(3) --> 원무(3)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6228') then
                  begin
                     Cells[0, 2]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[0, 2]   := $00DDD7E9;
                  end
                  // 진료지원파트(4) --> 원무(4)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6230') then
                  begin
                     Cells[0, 3]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[0, 3]   := $00DDD7E9;
                  end
                  // 진료지원파트(5) --> 원무(5)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6227') then
                  begin
                     Cells[1, 2]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[1, 2]   := $00DDD7E9;
                  end
                  // 진료지원파트(6) --> 원무(6)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5037') then
                  begin
                     Cells[1, 3]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[1, 3]   := $00DDD7E9;
                  end
                  // 진료지원파트(7)
                  // 자리배치 변경으로 추가 @ 2015.06.05 LSH
                  // --> 임시자리로 변경 @ 2016.07.26 LSH
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = 'XXXX') then
                  begin
                     Cells[2, 3]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[2, 3]   := $00DDD7E9;
                  end
                  // 진료파트(1)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5564') then
                  begin
                     Cells[0, 4]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[0, 4]   := $00FFD7A5;
                  end
                  // 진료파트(2)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6219') then
                  begin
                     Cells[1, 4]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[1, 4]   := $00FFD7A5;
                  end
                  // 진료파트(3)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6224') then
                  begin
                     Cells[0, 5]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[0, 5]   := $00FFD7A5;
                  end
                  // 진료파트(4)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5143') then
                  begin
                     Cells[1, 5]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[1, 5]   := $00FFD7A5;
                  end
                  // 진료파트(5)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6218') then
                  begin
                     Cells[2, 4]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[2, 4]   := $00FFD7A5;
                  end
                  // 진료파트(6)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6217') then
                  begin
                     Cells[3, 4]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[3, 4]   := $00FFD7A5;
                  end
                  // 진료파트(7)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6223') then
                  begin
                     Cells[2, 5]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[2, 5]   := $00FFD7A5;
                  end
                  // 진료파트(8)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6216') then
                  begin
                     Cells[3, 5]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[3, 5]   := $00FFD7A5;
                  end
                  // 원무파트(1) --> 진지(1)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6222') then
                  begin
                     Cells[0, 6]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[0, 6]   := $009DA1FF;
                  end
                  // 원무파트(2) --> 진지(2)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6221') then
                  begin
                     Cells[1, 6]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[1, 6]   := $009DA1FF;
                  end
                  // 원무파트(3) --> 진지(3)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6211') then
                  begin
                     Cells[1, 7]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[1, 7]   := $009DA1FF;
                  end
                  // 원무파트(4) --> 진지(4)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6214') then
                  begin
                     Cells[2, 6]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[2, 6]   := $009DA1FF;
                  end
                  // 원무파트(5) --> 진지(5)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6212') then
                  begin
                     Cells[2, 7]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[2, 7]   := $009DA1FF;
                  end
                  // 원무파트(6) --> 진지(6)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6220') then
                  begin
                     Cells[3, 6]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[3, 6]   := $009DA1FF;
                  end
                  // 원무파트(7) --> 진지(7)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6215') then
                  begin
                     Cells[3, 7]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[3, 7]   := $009DA1FF;
                  end
                  // 운영파트(1)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5565') then
                  begin
                     Cells[6, 1]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[6, 1]   := $009DB1C1;
                  end
                  // 운영파트(2)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5146') then
                  begin
                     Cells[7, 1]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[7, 1]   := $009DB1C1;
                  end
                  // 운영파트(3)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5563') then
                  begin
                     Cells[7, 2]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[7, 2]   := $009DB1C1;
                  end
                  // 운영파트(4) - 시스템 담당자 추가 @ 2014.07.17 LSH
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5145') then
                  begin
                     Cells[6, 2]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[6, 2]   := $009DB1C1;
                  end
                  // 기획파트(1)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5863') then
                  begin
                     Cells[3, 1]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[3, 1]   := $00C5B1C1;
                  end
                  // 기획파트(2)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5862') then
                  begin
                     Cells[4, 1]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[4, 1]   := $00C5B1C1;
                  end
                  // 일반관리파트(1)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6245') then
                  begin
                     Cells[7, 4]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[7, 4]   := $0085DFF1;
                  end
                  // 일반관리파트(2)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6234') then
                  begin
                     Cells[6, 5]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[6, 5]   := $0085DFF1;
                  end
                  // 일반관리파트(3)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6233') then
                  begin
                     Cells[7, 5]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[7, 5]   := $0085DFF1;
                  end
                  // 일반관리파트(4)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6232') then
                  begin
                     Cells[6, 4]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[6, 4]   := $0085DFF1;
                  end
                  // OP 선생님
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5880') then
                  begin
                     Cells[3, 3]    := TpGetSearch.GetOutputDataS('sDutyPart', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[3, 3]   := $009DB1C1;
                  end;






            end;
         end;
      end
      else if (in_SearchFlag = 'RELEASE') or
              (in_SearchFlag = 'RELEASESCAN') then
      begin
          with asg_Release do
          begin
            iRowCnt  := TpGetSearch.RowCount;
            RowCount := iRowCnt + FixedRows + 1;


            for i := 0 to iRowCnt - 1 do
            begin

               Cells[C_RL_REGDATE,    i+FixedRows]  := TpGetSearch.GetOutputDataS('sRegDate',    i);
               Cells[C_RL_DUTYSPEC,   i+FixedRows]  := TpGetSearch.GetOutputDataS('sDutySpec',   i);
               Cells[C_RL_CONTEXT,    i+FixedRows]  := TpGetSearch.GetOutputDataS('sContext',    i);
               Cells[C_RL_SRREQNO,    i+FixedRows]  := TpGetSearch.GetOutputDataS('sDataBase',   i);
               Cells[C_RL_REQUSER,    i+FixedRows]  := TpGetSearch.GetOutputDataS('sRegUser',    i);
               Cells[C_RL_DUTYUSER,   i+FixedRows]  := TpGetSearch.GetOutputDataS('sDutyUser',   i);
               Cells[C_RL_RELEASDT,   i+FixedRows]  := TpGetSearch.GetOutputDataS('sReleasDt',   i);
               Cells[C_RL_COFMUSER,   i+FixedRows]  := TpGetSearch.GetOutputDataS('sCofmUser',   i);
               Cells[C_RL_TESTDATE,   i+FixedRows]  := TpGetSearch.GetOutputDataS('sTestDate',   i);
               Cells[C_RL_CLIENSRC,   i+FixedRows]  := TpGetSearch.GetOutputDataS('sClienSrc',   i);
               Cells[C_RL_CLIENT,     i+FixedRows]  := TpGetSearch.GetOutputDataS('sClient',     i);
               Cells[C_RL_SERVESRC,   i+FixedRows]  := TpGetSearch.GetOutputDataS('sServeSrc',   i);
               Cells[C_RL_SERVER,     i+FixedRows]  := TpGetSearch.GetOutputDataS('sServer',     i);
               Cells[C_RL_RELEASUSER, i+FixedRows]  := TpGetSearch.GetOutputDataS('sRelesUsr',   i);
               Cells[C_RL_REMARK,     i+FixedRows]  := TpGetSearch.GetOutputDataS('sDutyRmk',    i);

               AutoSizeRow(i+FixedRows);



               {
               // Hidden info
               Cells[C_AN_REQNO,      i+FixedRows]  := TpGetSearch.GetOutputDataS('sReqNo',      i);
               Cells[C_AN_ATTACHNM,   i+FixedRows]  := TpGetSearch.GetOutputDataS('sAttachNm',   i);
               }


               if (Cells[C_RL_COFMUSER,    i+FixedRows] <> '') and
                  (Cells[C_RL_TESTDATE,    i+FixedRows] <> '') then
               begin
                  FontColors[C_RL_REGDATE,   i+FixedRows] := clGray;
                  FontColors[C_RL_DUTYSPEC,  i+FixedRows] := clGray;
                  FontColors[C_RL_CONTEXT,   i+FixedRows] := clGray;
                  FontColors[C_RL_COFMUSER,  i+FixedRows] := clGray;
                  FontColors[C_RL_SRREQNO,   i+FixedRows] := clGray;
                  FontColors[C_RL_RELEASDT,  i+FixedRows] := clGray;
                  FontColors[C_RL_TESTDATE,  i+FixedRows] := clGray;
                  FontColors[C_RL_DUTYUSER,  i+FixedRows] := clGray;
                  FontColors[C_RL_REQUSER,   i+FixedRows] := clGray;
               end
               else if (Cells[C_RL_COFMUSER,    i+FixedRows] = '') then
               begin
                  FontColors[C_RL_REGDATE,  i+FixedRows] := clRed;
                  FontColors[C_RL_DUTYSPEC, i+FixedRows] := clRed;
                  FontColors[C_RL_CONTEXT,  i+FixedRows] := clRed;
               end
               else if (Cells[C_RL_COFMUSER,    i+FixedRows] <> '') and
                       (Cells[C_RL_RELEASDT,    i+FixedRows] =  '') and
                       (Cells[C_RL_TESTDATE,    i+FixedRows] =  '') then
               begin
                  FontColors[C_RL_REGDATE,  i+FixedRows] := clBlue;
                  FontColors[C_RL_DUTYSPEC, i+FixedRows] := clBlue;
                  FontColors[C_RL_CONTEXT,  i+FixedRows] := clBlue;
                  FontColors[C_RL_COFMUSER, i+FixedRows] := clBlue;
                  FontColors[C_RL_DUTYUSER, i+FixedRows] := clBlue;
                  FontColors[C_RL_REQUSER,  i+FixedRows] := clBlue;
               end
               else if (Cells[C_RL_COFMUSER,    i+FixedRows] <> '') and
                       (Cells[C_RL_RELEASDT,    i+FixedRows] <> '') and
                       (Cells[C_RL_TESTDATE,    i+FixedRows] =  '') then
               begin
                  FontColors[C_RL_REGDATE,  i+FixedRows] := clPurple;
                  FontColors[C_RL_DUTYSPEC, i+FixedRows] := clPurple;
                  FontColors[C_RL_CONTEXT,  i+FixedRows] := clPurple;
                  FontColors[C_RL_COFMUSER, i+FixedRows] := clPurple;
                  FontColors[C_RL_RELEASDT, i+FixedRows] := clPurple;
                  FontColors[C_RL_DUTYUSER, i+FixedRows] := clPurple;
                  FontColors[C_RL_REQUSER,  i+FixedRows] := clPurple;
               end;

               // 특기사항 Coloring
               if (Cells[C_RL_REMARK,    i+FixedRows] <> '') then
                  FontColors[C_RL_REMARK,  i+FixedRows] := $0089A1FF;

            end;


            // RowCount 정리
            RowCount := RowCount - 1;


            // Comments
            if (iRowCnt = -1) then
               lb_RegDoc.Caption  := '▶ '  + GetTxMsg + '.'      // 검색자료 MAXROWCNT 초과 메세지 표기 @ 2016.11.18 LSH
            else
               lb_RegDoc.Caption  := '▶ ' + IntToStr(iRowCnt) + '건의 릴리즈 내역을 조회하였습니다.';

         end;
      end
      else if (in_SearchFlag = 'DUTYMAKE') or
              (in_SearchFlag = 'DUTYSEARCH') then
      begin
          with asg_Duty do
          begin
            iRowCnt  := TpGetSearch.RowCount;
            //RowCount := iRowCnt + FixedRows + 1;

            RowCount := 2;

            for i := 0 to iRowCnt - 1 do
            begin

               //if (TpGetSearch.GetOutputDataS('sRegDate', i) >= fmed_DutyFrDt.Text) and
               //   (TpGetSearch.GetOutputDataS('sRegDate', i) <= fmed_DutyToDt.Text) then
               begin

                  Cells[C_DT_DUTYDATE,  RowCount - 1{i+FixedRows}] := TpGetSearch.GetOutputDataS('sRegDate'   , i);
                  Cells[C_DT_YOIL,      RowCount - 1{i+FixedRows}] := GetDayofWeek(StrToDate(TpGetSearch.GetOutputDataS('sRegDate'   , i)), 'HS');


                  // 근무-근태(확정) 조회 vs. 근무표 생성 분기 @ 2015.06.12 LSH
                  if (in_SearchFlag = 'DUTYMAKE') then
                  begin
                     // 주말(법정공휴일 제외)
                     if TpGetSearch.GetOutputDataS('sFlag', i) = 'WEEK' then
                     begin
                        Cells[C_DT_DUTYFLAG,  RowCount - 1] := 'W'
                     end
                     // 평일(야간)
                     else if TpGetSearch.GetOutputDataS('sFlag', i) = '' then
                     begin
                        Cells[C_DT_DUTYFLAG,  RowCount - 1] := 'P'
                     end
                     // 법정공휴일(주말포함)
                     else
                     begin
                        Cells[C_DT_DUTYFLAG,  RowCount - 1] := TpGetSearch.GetOutputDataS('sFlag', i);
                     end;
                  end
                  else if (in_SearchFlag = 'DUTYSEARCH') then
                  begin
                     if TpGetSearch.GetOutputDataS('sFlag', i) = 'P' then
                        Cells[C_DT_DUTYFLAG,  RowCount - 1] := '평일(야간)'
                     else if TpGetSearch.GetOutputDataS('sFlag', i) = 'W' then
                        Cells[C_DT_DUTYFLAG,  RowCount - 1] := '휴일(주말)'
                     else
                        Cells[C_DT_DUTYFLAG,  RowCount - 1{i+FixedRows}] := '휴일(' + TpGetSearch.GetOutputDataS('sFlag', i) + ')'
                  end;


                  Cells[C_DT_DUTYUSER,  RowCount - 1{i+FixedRows}] := TpGetSearch.GetOutputDataS('sDutyUser'  , i);
                  Cells[C_DT_REMARK,    RowCount - 1{i+FixedRows}] := TpGetSearch.GetOutputDataS('sDutyRmk'   , i);
                  Cells[C_DT_SEQNO ,    RowCount - 1{i+FixedRows}] := TpGetSearch.GetOutputDataS('sDocSeq'    , i);     // 당직자 변경위한 Seqno @ 2015.06.12 LSH
                  Cells[C_DT_ORGDTYUSR, RowCount - 1{i+FixedRows}] := TpGetSearch.GetOutputDataS('sDutyUser'  , i);     // 당직자 변경전 Orig. 근무자 @ 2015.06.12 LSH
                  Cells[C_DT_ORGREMARK, RowCount - 1{i+FixedRows}] := TpGetSearch.GetOutputDataS('sDutyRmk'   , i);     // 근태 특이사항 변경전 Orig. 특이사항 @ 2015.06.12 LSH


                  //AddButton(C_M_USERADD, i+FixedRows, ColWidths[C_M_USERADD]-5, 20, '+', haBeforeText, vaCenter);                // 사용자 Add


                  //------------------------------------------------------------
                  // 휴일근무 Coloring
                  //------------------------------------------------------------
                  if (TpGetSearch.GetOutputDataS('sFlag', i) <> 'P') then
                  begin
                     FontColors[C_DT_YOIL,      RowCount - 1{i+FixedRows}] := clRed;
                     FontColors[C_DT_DUTYFLAG,  RowCount - 1{i+FixedRows}] := clRed;
                     FontColors[C_DT_DUTYDATE,  RowCount - 1{i+FixedRows}] := clRed;
                  end;


                  // 근무-근태(확정) 조회시, Coloring @ 2015.06.12 LSH
                  if (in_SearchFlag = 'DUTYSEARCH') then
                  begin

                     //------------------------------------------------------------
                     // 오늘일자 Coloring
                     //------------------------------------------------------------
                     if (TpGetSearch.GetOutputDataS('sRegDate', i) = FormatDateTime('yyyy-mm-dd', Date)) then
                     begin
                        Colors[C_DT_DUTYDATE,      RowCount - 1{i+FixedRows}] := $00ABADD9;
                        Colors[C_DT_YOIL,          RowCount - 1{i+FixedRows}] := $00ABADD9;
                        Colors[C_DT_DUTYFLAG,      RowCount - 1{i+FixedRows}] := $00ABADD9;
                        Colors[C_DT_DUTYUSER,      RowCount - 1{i+FixedRows}] := $00ABADD9;
                        Colors[C_DT_REMARK,        RowCount - 1{i+FixedRows}] := $00ABADD9;

                        FontStyles[C_DT_DUTYDATE,      RowCount - 1{i+FixedRows}] := [fsBold];
                        FontStyles[C_DT_YOIL,          RowCount - 1{i+FixedRows}] := [fsBold];
                        FontStyles[C_DT_DUTYFLAG,      RowCount - 1{i+FixedRows}] := [fsBold];
                        FontStyles[C_DT_DUTYUSER,      RowCount - 1{i+FixedRows}] := [fsBold];
                        FontStyles[C_DT_REMARK,        RowCount - 1{i+FixedRows}] := [fsBold];
                     end;

                     //------------------------------------------------------------
                     // 본인근무 Coloring
                     //------------------------------------------------------------
                     if (TpGetSearch.GetOutputDataS('sDutyUser', i) = FsUserNm) then
                     begin
                        Colors[C_DT_DUTYUSER,      RowCount - 1{i+FixedRows}] := $0076D1DB;
                     end;

                  end;

                  RowCount := RowCount + 1;

               end;

            end;
         end;


         // RowCount 정리
         asg_Duty.RowCount := asg_Duty.RowCount - 1;


         // Comments
         if (in_SearchFlag = 'DUTYSEARCH') then
            lb_RegDoc.Caption := '▶ 선택기간내 근무 및 근태현황을 조회하였습니다.'
         else if (in_SearchFlag = 'DUTYMAKE') then
            lb_RegDoc.Caption := '▶ 선택기간내 (확정등록전) 근무표를 조회하였습니다.';


      end
      else if (in_SearchFlag = 'DUTYLIST') then
      begin
          with asg_DutyList do
          begin
            iRowCnt  := TpGetSearch.RowCount;
            RowCount := iRowCnt + FixedRows + 1;


            for i := 0 to iRowCnt - 1 do
            begin
               for j := 0 to ColCount - 1 do
               begin
                  Cells[0,  i+FixedRows] := TpGetSearch.GetOutputDataS('sLinkSeq'   , i);
                  Cells[1,  i+FixedRows] := TpGetSearch.GetOutputDataS('sDutyUser'  , i);
                  Cells[2,  i+FixedRows] := TpGetSearch.GetOutputDataS('sRegDate'   , i);    // (근무기준 변경위한) 평일근무 시작일 @ 2015.06.12 LSH
                  Cells[3,  i+FixedRows] := TpGetSearch.GetOutputDataS('sTestDate'  , i);    // (근무기준 변경위한) 휴일근무 시작일 @ 2015.06.12 LSH
               end;
            end;
         end;


         // RowCount 정리
         asg_DutyList.RowCount := asg_DutyList.RowCount - 1;

         // Comments
         lb_RegDoc.Caption := '▶ 부서 당직근무표 기준 및 순번을 조회하였습니다.';
      end
      else if (in_SearchFlag = 'DBCATEGORY') then
      begin
          with asg_DBMaster do
          begin
            iRowCnt  := TpGetSearch.RowCount;
            RowCount := iRowCnt + FixedRows + 1;


            for i := 0 to iRowCnt - 1 do
            begin
               Cells[0,  i+FixedRows] := TpGetSearch.GetOutputDataS('sLinkSeq'  , i);
               Cells[1,  i+FixedRows] := TpGetSearch.GetOutputDataS('sDutyRmk'  , i);
               Cells[2,  i+FixedRows] := TpGetSearch.GetOutputDataS('sDutyUser' , i);
            end;
         end;


         // RowCount 정리
         asg_DBMaster.RowCount := asg_DBMaster.RowCount - 1;

         // Comments
         lb_Analysis.Caption := '▶ ' + fcb_WorkArea.Text + ' ' + fcb_DutyPart.Text + '파트 ' + fcb_WorkGubun.Text + '를 조회하였습니다.';
      end
      else if (in_SearchFlag = 'DBCATEDET') then
      begin
          with asg_DBDetail do
          begin
            iRowCnt  := TpGetSearch.RowCount;
            RowCount := iRowCnt + FixedRows + 1;


            for i := 0 to iRowCnt - 1 do
            begin
               Cells[0,  i+FixedRows] := TpGetSearch.GetOutputDataS('sLinkSeq'  , i);
               Cells[1,  i+FixedRows] := TpGetSearch.GetOutputDataS('sDutyRmk'  , i) + TpGetSearch.GetOutputDataS('sClient'  , i);
               Cells[2,  i+FixedRows] := TpGetSearch.GetOutputDataS('sDocSeq'   , i);
               Cells[3,  i+FixedRows] := TpGetSearch.GetOutputDataS('sDocRmk'   , i) + TpGetSearch.GetOutputDataS('sServer'  , i);;
               Cells[4,  i+FixedRows] := TpGetSearch.GetOutputDataS('sDutyUser' , i);

               {
               if (AnsiUpperCase(fcb_WorkGubun.Text) = 'TRIGGER') then
               begin
                  AutoSizeRow(i+FixedRows);
               end;
               }
            end;
         end;


         // RowCount 정리
         asg_DBDetail.RowCount := asg_DBDetail.RowCount - 1;

         // Comments
         lb_Analysis.Caption := '▶ ' + fcb_WorkArea.Text + ' ' + fcb_DutyPart.Text + '파트 ' + fcb_WorkGubun.Text + ' [' + asg_DBMaster.Cells[0, asg_DBMaster.Row] + '] 조회하였습니다.';
      end
      else if (in_SearchFlag = 'DBSEARCH') then
      begin
          with asg_DBDetail do
          begin
            iRowCnt  := TpGetSearch.RowCount;
            RowCount := iRowCnt + FixedRows + 1;


            for i := 0 to iRowCnt - 1 do
            begin
               for j := 0 to ColCount - 1 do
               begin
                  // 이중 키워드 매칭 적용 @ 2015.06.03 LSH
                  if (PosByte(AnsiUpperCase(TokenStr(fed_Analysis.Text + ' ', ' ', 1)), AnsiUpperCase(TpGetSearch.GetOutputDataS('sLinkSeq', i))) > 0) or
                     (PosByte(AnsiUpperCase(TokenStr(fed_Analysis.Text + ' ', ' ', 2)), AnsiUpperCase(TpGetSearch.GetOutputDataS('sLinkSeq', i))) > 0) then
                  begin
                     Cells[0,    i+FixedRows]    := TpGetSearch.GetOutputDataS('sLinkSeq', i);
                     FontStyles[0, i+FixedRows]  := [fsBold];
                     FontColors[0, i+FixedRows]  := clMaroon;
                  end
                  else
                  begin
                     Cells[0,    i+FixedRows]    := TpGetSearch.GetOutputDataS('sLinkSeq', i);
                     FontStyles[0, i+FixedRows]  := [];
                  end;

                  if (PosByte(AnsiUpperCase(TokenStr(fed_Analysis.Text + ' ', ' ', 1)), AnsiUpperCase(TpGetSearch.GetOutputDataS('sDutyRmk', i))) > 0) or
                     (PosByte(AnsiUpperCase(TokenStr(fed_Analysis.Text + ' ', ' ', 2)), AnsiUpperCase(TpGetSearch.GetOutputDataS('sDutyRmk', i))) > 0) then
                  begin
                     Cells[1,    i+FixedRows]    := TpGetSearch.GetOutputDataS('sDutyRmk', i);
                     FontStyles[1, i+FixedRows]  := [fsBold];
                     FontColors[1, i+FixedRows]  := clMaroon;
                  end
                  else
                  begin
                     Cells[1,    i+FixedRows]    := TpGetSearch.GetOutputDataS('sDutyRmk', i);
                     FontStyles[1, i+FixedRows]  := [];
                  end;

                  if (PosByte(AnsiUpperCase(TokenStr(fed_Analysis.Text + ' ', ' ', 1)), AnsiUpperCase(TpGetSearch.GetOutputDataS('sClient', i))) > 0) or
                     (PosByte(AnsiUpperCase(TokenStr(fed_Analysis.Text + ' ', ' ', 2)), AnsiUpperCase(TpGetSearch.GetOutputDataS('sClient', i))) > 0)then
                  begin
                     Cells[2,    i+FixedRows]    := TpGetSearch.GetOutputDataS('sClient', i);
                     FontStyles[2, i+FixedRows]  := [fsBold];
                     FontColors[2, i+FixedRows]  := clMaroon;
                  end
                  else
                  begin
                     Cells[2,    i+FixedRows]    := TpGetSearch.GetOutputDataS('sClient', i);
                     FontStyles[2, i+FixedRows]  := [];
                  end;

                  if (PosByte(AnsiUpperCase(TokenStr(fed_Analysis.Text + ' ', ' ', 1)), AnsiUpperCase(TpGetSearch.GetOutputDataS('sDocRmk', i))) > 0) or
                     (PosByte(AnsiUpperCase(TokenStr(fed_Analysis.Text + ' ', ' ', 2)), AnsiUpperCase(TpGetSearch.GetOutputDataS('sDocRmk', i))) > 0) then
                  begin
                     Cells[3,    i+FixedRows]    := TpGetSearch.GetOutputDataS('sDocRmk', i);
                     FontStyles[3, i+FixedRows]  := [fsBold];
                     FontColors[3, i+FixedRows]  := clMaroon;
                  end
                  else
                  begin
                     Cells[3,    i+FixedRows]    := TpGetSearch.GetOutputDataS('sDocRmk', i);
                     FontStyles[3, i+FixedRows]  := [];
                  end;

                  if (PosByte(AnsiUpperCase(TokenStr(fed_Analysis.Text + ' ', ' ', 1)), AnsiUpperCase(TpGetSearch.GetOutputDataS('sServer', i))) > 0) or
                     (PosByte(AnsiUpperCase(TokenStr(fed_Analysis.Text + ' ', ' ', 2)), AnsiUpperCase(TpGetSearch.GetOutputDataS('sServer', i))) > 0) then
                  begin
                     Cells[4,    i+FixedRows]    := TpGetSearch.GetOutputDataS('sServer', i);
                     FontStyles[4, i+FixedRows]  := [fsBold];
                     FontColors[4, i+FixedRows]  := clMaroon;
                  end
                  else
                  begin
                     Cells[4,    i+FixedRows]    := TpGetSearch.GetOutputDataS('sServer', i);
                     FontStyles[4, i+FixedRows]  := [];
                  end;

                  {
                  Cells[0,  i+FixedRows] := TpGetSearch.GetOutputDataS('sLinkSeq'  , i);
                  Cells[1,  i+FixedRows] := TpGetSearch.GetOutputDataS('sDutyRmk'  , i);
                  Cells[2,  i+FixedRows] := TpGetSearch.GetOutputDataS('sClient'   , i);
                  Cells[3,  i+FixedRows] := TpGetSearch.GetOutputDataS('sDocRmk'   , i);
                  Cells[4,  i+FixedRows] := TpGetSearch.GetOutputDataS('sServer'   , i);
                  }
               end;
            end;
         end;


         // RowCount 정리
         asg_DBDetail.RowCount := asg_DBDetail.RowCount - 1;

         // Comments
         lb_Analysis.Caption := '▶ ' + fcb_WorkArea.Text + ' ' + fcb_DutyPart.Text + ' 키워드 검색 [' + Trim(fed_Analysis.Text) + '] ' + IntToStr(iRowCnt) + '건을 조회하였습니다.';
      end
      else if (in_SearchFlag = 'SRHISTORY') then
      begin
         with asg_WeeklyRpt do
         begin
            iRowCnt  := TpGetSearch.RowCount;

            //RowCount := iRowCnt + FixedRows + 1;

            // Merge 변수
            iBaseRow       := 0;    // 대분류 Merge Start Index
            iRowSpan       := 1;    // 대분류 Span  Index
            jBaseRow       := 0;    // 중분류 Merge Start Index
            jRowSpan       := 1;    // 중분류 Span  Index
            kBaseRow       := 0;    // 소분류 Merge Start Index
            kRowSpan       := 1;    // 소분류 Span  Index

            sPrevUptitle    := '';   // 대분류 이전 Title
            sPrevMidtitle   := '';   // 중분류 이전 Title
            sPrevDowntitle  := '';   // 소분류 이전 Title


            for i := 0 to iRowCnt - 1 do
            begin
               Cells[C_WK_DATE,     RowCount - 1] := TpGetSearch.GetOutputDataS('sRegDate'    , i);
               Cells[C_WK_GUBUN,    RowCount - 1] := TpGetSearch.GetOutputDataS('sDutyRmk'    , i);
               Cells[C_WK_CONTEXT,  RowCount - 1] := Trim(TpGetSearch.GetOutputDataS('sContext'    , i));

               Cells[C_WK_STEP,     RowCount - 1] := TpGetSearch.GetOutputDataS('sDutySpec'   , i);


               // 진행단계  Coloring
               if PosByte('검토', Cells[C_WK_STEP,        RowCount - 1]) > 0 then
               begin
                  Colors[C_WK_CONTEXT,    RowCount - 1] := $0074D7AF;
                  Colors[C_WK_STEP,       RowCount - 1] := $0074D7AF;
                  AddComment(C_WK_STEP,
                             RowCount - 1,
                             '업무일지에 [검토]라는 단어가 포함된 경우,' + #13#10 +
                             '또는 S/R내역 요청 특이사항에 담당예정자가 적힌 경우에 녹색으로 표기됩니다.');
               end
               else if PosByte('진행', Cells[C_WK_STEP,   RowCount - 1]) > 0 then
               begin
                  Colors[C_WK_CONTEXT,    RowCount - 1] := $0074A9FF;
                  Colors[C_WK_STEP,       RowCount - 1] := $0074A9FF;
                  AddComment(C_WK_STEP,
                             RowCount - 1,
                             '업무일지상 [진행]이라는 단어가 포함된 경우,' + #13#10 +
                             '또는 S/R내역 담당자 지정후 결과승인요청 상태 이전인 경우에 붉은색으로 표기됩니다.');
               end
               else if (PosByte('완료', Cells[C_WK_STEP,   RowCount - 1]) > 0) and
                       (PosByte('진행', Cells[C_WK_STEP,   RowCount - 1]) = 0) and
                       (PosByte('검토', Cells[C_WK_STEP,   RowCount - 1]) = 0) then
               begin
                  Colors[C_WK_CONTEXT,    RowCount - 1] := $00D5CDC7;
                  Colors[C_WK_STEP,       RowCount - 1] := $00D5CDC7;
               end
               else
               begin
                  Colors[C_WK_CONTEXT,    RowCount - 1] := clWhite;
                  Colors[C_WK_STEP,       RowCount - 1] := clWhite;
               end;


               AutoSizeRow(RowCount - 1);


               // 해당 화면 1-Cycle 코드의 마지막 Index를 최종 RowCount 로 사용
               if (i = iRowCnt) then
                  iLstRowCnt := i+1
               else
                  iLstRowCnt := iRowCnt;




               // 대분류 항목 MergeCell 체크
               if TpGetSearch.GetOutputDataS('sRegDate', i) <> sPrevUptitle then
               begin
                  // test
                  //Memo1.Lines.Add('--- Test Logging Start ---');


                  // 이전 대분류와 다른 경우
                  if i > 0 then
                  begin
                     MergeCells(0, iBaseRow, 1, iRowSpan);
                     iRowSpan := 1;
                  end;

                  // ★ Key Point
                  iBaseRow := RowCount - 1;


                  // 중분류 항목 MergeCell 체크

                  { -- 대분류가 상위와 다른경우, 중분류가 같더라도, 구분 Grid line을 그리기 위해 주석처리 @ 2013.10.14 LSH
                  if TpGetSearch.GetOutputDataS('sDutyRmk', i) <> sPrevMidTitle then
                  begin
                  }
                     // 이전 중분류와 다른 경우
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;


                     // ★ Key Point
                     jBaseRow := RowCount - 1;



                     // 소분류 항목 MergeCell 체크
                     if TpGetSearch.GetOutputDataS('sContext', i) <> sPrevDownTitle then
                     begin
                        // 이전 소분류와 다른 경우
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // ★ Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // 이전 소분류와 다른 경우
                        Inc(kRowSpan);

                  { -- 대분류가 상위와 다른경우, 중분류가 같더라도, 구분 Grid line을 그리기 위해 주석처리 @ 2013.10.14 LSH
                  end
                  else
                  begin
                     // 이전 중분류와 같은 경우
                     Inc(jRowSpan);
                  end;
                  }
               end
               else
               begin

                  // 중분류 항목 MergeCell 체크
                  if TpGetSearch.GetOutputDataS('sDutyRmk', i) <> sPrevMidTitle then
                  begin
                     // 이전 중분류와 다른 경우
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;

                     // ★ Key Point
                     jBaseRow := RowCount - 1;


                     // 소분류 항목 MergeCell 체크
                     if TpGetSearch.GetOutputDataS('sContext', i) <> sPrevDownTitle then
                     begin
                        // 이전 소분류와 같은 경우
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // ★ Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // 이전 소분류와 다른 경우
                        Inc(kRowSpan);
                  end
                  else
                  begin
                     // 중분류 span Index ++
                     Inc(jRowSpan);


                     // 소분류 항목 MergeCell 체크
                     if TpGetSearch.GetOutputDataS('sContext', i) <> sPrevDownTitle then
                     begin
                        // 이전 소분류와 다른 경우
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // ★ Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // 이전 소분류와 동일한 경우
                        Inc(kRowSpan);

                  end;


                  // 대분류 Span index ++
                  Inc(iRowSpan);

               end;


               // Set previous value
               sPrevUptitle   := TpGetSearch.GetOutputDataS('sRegDate', i);
               sPrevmidtitle  := TpGetSearch.GetOutputDataS('sDutyRmk', i);
               sPrevDowntitle := TpGetSearch.GetOutputDataS('sContext', i);


               // RowCount 다음줄로 넘김.
               RowCount := RowCount + 1;


            end;


            // ★ 최종 Code List-up후 마지막 Row Merge 처리
            asg_WeeklyRpt.MergeCells(0, iBaseRow, 1, iRowSpan);
            asg_WeeklyRpt.MergeCells(1, jBaseRow, 1, jRowSpan);
            asg_WeeklyRpt.RowCount := iLstRowCnt + 1;
         end;

         // Comments
         //lb_.Caption := '▶ ' + Trim(fed_Analysis.Text) + '의 리포팅 ' + IntToStr(iRowCnt) + '건을 작성하였습니다.';

      end

      else if (in_SearchFlag = 'WORKCONN') then
      begin
         with asg_WorkConn do
         begin
            iRowCnt  := TpGetSearch.RowCount;
            //RowCount := iRowCnt + FixedRows + 1;


            // Merge 변수
            iBaseRow       := 0;    // 대분류 Merge Start Index
            iRowSpan       := 1;    // 대분류 Span  Index
            jBaseRow       := 0;    // 중분류 Merge Start Index
            jRowSpan       := 1;    // 중분류 Span  Index
            kBaseRow       := 0;    // 소분류 Merge Start Index
            kRowSpan       := 1;    // 소분류 Span  Index

            sPrevUptitle    := '';   // 대분류 이전 Title
            sPrevMidtitle   := '';   // 중분류 이전 Title
            sPrevDowntitle  := '';   // 소분류 이전 Title



            for i := 0 to iRowCnt - 1 do
            begin

               Cells[C_WC_LOCATE,   RowCount - 1] := TpGetSearch.GetOutputDataS('sLocate'    , i);

               if PosByte('개발', TpGetSearch.GetOutputDataS('sDutyPart'  , i)) > 0 then
                  Cells[C_WC_DUTYPART, RowCount - 1] := TpGetSearch.GetOutputDataS('sDutyPart'  , i) + #13#10 + '(' + TpGetSearch.GetOutputDataS('sDutySpec'  , i) + ')'
               else
                  Cells[C_WC_DUTYPART, RowCount - 1] := TpGetSearch.GetOutputDataS('sDutyPart'  , i);

               Cells[C_WC_CONTEXT,  RowCount - 1] := Trim(TpGetSearch.GetOutputDataS('sContext' , i));
               Cells[C_WC_DUTYUSER, RowCount - 1] := TpGetSearch.GetOutputDataS('sDutyUser'  , i);
               Cells[C_WC_REGDATE,  RowCount - 1] := TpGetSearch.GetOutputDataS('sRegDate'   , i);


               AutoSizeRow(RowCount - 1);



               // 해당 화면 1-Cycle 코드의 마지막 Index를 최종 RowCount 로 사용
               if (i = iRowCnt) then
                  iLstRowCnt := i+1
               else
                  iLstRowCnt := iRowCnt;




               // 대분류 항목 MergeCell 체크
               if TpGetSearch.GetOutputDataS('sLocate', i) <> sPrevUptitle then
               begin
                  // test
                  //Memo1.Lines.Add('--- Test Logging Start ---');


                  // 이전 대분류와 다른 경우
                  if i > 0 then
                  begin
                     MergeCells(0, iBaseRow, 1, iRowSpan);
                     iRowSpan := 1;
                  end;

                  // ★ Key Point
                  iBaseRow := RowCount - 1;


                  // 중분류 항목 MergeCell 체크

                  { -- 대분류가 상위와 다른경우, 중분류가 같더라도, 구분 Grid line을 그리기 위해 주석처리 @ 2013.10.14 LSH
                  if TpGetSearch.GetOutputDataS('sDutyRmk', i) <> sPrevMidTitle then
                  begin
                  }
                     // 이전 중분류와 다른 경우
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;


                     // ★ Key Point
                     jBaseRow := RowCount - 1;



                     // 소분류 항목 MergeCell 체크
                     if TpGetSearch.GetOutputDataS('sDutyUser', i) <> sPrevDownTitle then
                     begin
                        // 이전 소분류와 다른 경우
                        if i > 0 then
                        begin
                           MergeCells(4, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // ★ Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // 이전 소분류와 다른 경우
                        Inc(kRowSpan);

                  { -- 대분류가 상위와 다른경우, 중분류가 같더라도, 구분 Grid line을 그리기 위해 주석처리 @ 2013.10.14 LSH
                  end
                  else
                  begin
                     // 이전 중분류와 같은 경우
                     Inc(jRowSpan);
                  end;
                  }
               end
               else
               begin

                  // 중분류 항목 MergeCell 체크
                  if TpGetSearch.GetOutputDataS('sDutySpec', i) <> sPrevMidTitle then
                  begin
                     // 이전 중분류와 다른 경우
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;

                     // ★ Key Point
                     jBaseRow := RowCount - 1;


                     // 소분류 항목 MergeCell 체크
                     if TpGetSearch.GetOutputDataS('sDutyUser', i) <> sPrevDownTitle then
                     begin
                        // 이전 소분류와 같은 경우
                        if i > 0 then
                        begin
                           MergeCells(4, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // ★ Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // 이전 소분류와 다른 경우
                        Inc(kRowSpan);
                  end
                  else
                  begin
                     // 중분류 span Index ++
                     Inc(jRowSpan);


                     // 소분류 항목 MergeCell 체크
                     if TpGetSearch.GetOutputDataS('sDutyUser', i) <> sPrevDownTitle then
                     begin
                        // 이전 소분류와 다른 경우
                        if i > 0 then
                        begin
                           MergeCells(4, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // ★ Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // 이전 소분류와 동일한 경우
                        Inc(kRowSpan);

                  end;


                  // 대분류 Span index ++
                  Inc(iRowSpan);

               end;


               // Set previous value
               sPrevUptitle   := TpGetSearch.GetOutputDataS('sLocate', i);
               sPrevmidtitle  := TpGetSearch.GetOutputDataS('sDutySpec', i);
               sPrevDowntitle := TpGetSearch.GetOutputDataS('sDutyUser', i);


               // RowCount 다음줄로 넘김.
               RowCount := RowCount + 1;


            end;


            // ★ 최종 Code List-up후 마지막 Row Merge 처리
            asg_WorkConn.MergeCells(0, iBaseRow, 1, iRowSpan);
            asg_WorkConn.MergeCells(1, jBaseRow, 1, jRowSpan);
            asg_WorkConn.MergeCells(5, kBaseRow, 1, kRowSpan);
            asg_WorkConn.RowCount := iLstRowCnt + 1;

         end;

         // Comments
         lb_WorkConn.Caption := '▶ ' + IntToStr(iRowCnt) + '건의 주요업무공유 내역을 조회하였습니다.';

      end
      else if (in_SearchFlag = 'BIGDATA') then
      begin
          with asg_BigData do
          begin
            iRowCnt  := TpGetSearch.RowCount;
            RowCount := iRowCnt + FixedRows + 1;


            for i := 0 to iRowCnt - 1 do
            begin
               for j := 0 to ColCount - 1 do
               begin
                  Cells[C_BD_YEAR	,  i+FixedRows]  := TpGetSearch.GetOutputDataS('sDutyPart',    i);
                  Cells[C_BD_SEQNO  ,  i+FixedRows]  := TpGetSearch.GetOutputDataS('sDutySpec',    i);
                  Cells[C_BD_SHOWDT ,  i+FixedRows]  := TpGetSearch.GetOutputDataS('sRegDate',     i);
                  Cells[C_BD_NO1    ,  i+FixedRows]  := TpGetSearch.GetOutputDataS('sRelDept',     i);
                  Cells[C_BD_NO2    ,  i+FixedRows]  := TpGetSearch.GetOutputDataS('sRelesUsr',    i);
                  Cells[C_BD_NO3    ,  i+FixedRows]  := TpGetSearch.GetOutputDataS('sCateUp',      i);
                  Cells[C_BD_NO4    ,  i+FixedRows]  := TpGetSearch.GetOutputDataS('sCateDown',    i);
                  Cells[C_BD_NO5    ,  i+FixedRows]  := TpGetSearch.GetOutputDataS('sAttachNm',    i);
                  Cells[C_BD_NO6    ,  i+FixedRows]  := TpGetSearch.GetOutputDataS('sHideFile',    i);
                  Cells[C_BD_NOBONUS,  i+FixedRows]  := TpGetSearch.GetOutputDataS('sServerIp',    i);
                  Cells[C_BD_1STCNT ,  i+FixedRows]  := TpGetSearch.GetOutputDataS('sCofmUser',    i);
                  Cells[C_BD_1STAMT ,  i+FixedRows]  := TpGetSearch.GetOutputDataS('sContext',     i);
                  Cells[C_BD_2NDCNT ,  i+FixedRows]  := TpGetSearch.GetOutputDataS('sReleasDt',    i);
                  Cells[C_BD_2NDAMT ,  i+FixedRows]  := TpGetSearch.GetOutputDataS('sClient',      i);
                  Cells[C_BD_3THCNT ,  i+FixedRows]  := TpGetSearch.GetOutputDataS('sTestDate',    i);
                  Cells[C_BD_3THAMT ,  i+FixedRows]  := TpGetSearch.GetOutputDataS('sServer',      i);
                  Cells[C_BD_4THCNT ,  i+FixedRows]  := TpGetSearch.GetOutputDataS('sDataBase',    i);
                  Cells[C_BD_4THAMT ,  i+FixedRows]  := TpGetSearch.GetOutputDataS('sClienSrc',    i);
                  Cells[C_BD_5THCNT ,  i+FixedRows]  := TpGetSearch.GetOutputDataS('sDutyRmk',     i);
                  Cells[C_BD_5THAMT ,  i+FixedRows]  := TpGetSearch.GetOutputDataS('sServeSrc',    i);


                  //AutoSizeRow(i+FixedRows);

                  {
                  // Hidden info
                  Cells[C_AN_REQNO,      i+FixedRows]  := TpGetSearch.GetOutputDataS('sReqNo',      i);
                  Cells[C_AN_ATTACHNM,   i+FixedRows]  := TpGetSearch.GetOutputDataS('sAttachNm',   i);
                  }

                  // No1. Range 구간별 Coloring
                  if (StrToInt(Cells[C_BD_NO1,    i+FixedRows]) >= 1)  and
                     (StrToInt(Cells[C_BD_NO1,    i+FixedRows]) <= 10) then
                  begin
                     Colors[C_BD_NO1,    i+FixedRows] := clYellow;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NO1,    i+FixedRows]) >= 11)  and
                     (StrToInt(Cells[C_BD_NO1,    i+FixedRows]) <= 20) then
                  begin
                     Colors[C_BD_NO1,    i+FixedRows] := $00E09000;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NO1,    i+FixedRows]) >= 21)  and
                     (StrToInt(Cells[C_BD_NO1,    i+FixedRows]) <= 30) then
                  begin
                     Colors[C_BD_NO1,    i+FixedRows] := clRed;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NO1,    i+FixedRows]) >= 31)  and
                     (StrToInt(Cells[C_BD_NO1,    i+FixedRows]) <= 40) then
                  begin
                     Colors[C_BD_NO1,    i+FixedRows] := clGray;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NO1,    i+FixedRows]) >= 41)  and
                     (StrToInt(Cells[C_BD_NO1,    i+FixedRows]) <= 45) then
                  begin
                     Colors[C_BD_NO1,    i+FixedRows] := clGreen;
                  end;


                  // No2. Range 구간별 Coloring
                  if (StrToInt(Cells[C_BD_NO2,    i+FixedRows]) >= 1)  and
                     (StrToInt(Cells[C_BD_NO2,    i+FixedRows]) <= 10) then
                  begin
                     Colors[C_BD_NO2,    i+FixedRows] := clYellow;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NO2,    i+FixedRows]) >= 11)  and
                     (StrToInt(Cells[C_BD_NO2,    i+FixedRows]) <= 20) then
                  begin
                     Colors[C_BD_NO2,    i+FixedRows] := $00E09000;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NO2,    i+FixedRows]) >= 21)  and
                     (StrToInt(Cells[C_BD_NO2,    i+FixedRows]) <= 30) then
                  begin
                     Colors[C_BD_NO2,    i+FixedRows] := clRed;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NO2,    i+FixedRows]) >= 31)  and
                     (StrToInt(Cells[C_BD_NO2,    i+FixedRows]) <= 40) then
                  begin
                     Colors[C_BD_NO2,    i+FixedRows] := clGray;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NO2,    i+FixedRows]) >= 41)  and
                     (StrToInt(Cells[C_BD_NO2,    i+FixedRows]) <= 45) then
                  begin
                     Colors[C_BD_NO2,    i+FixedRows] := clGreen;
                  end;


                  // No3. Range 구간별 Coloring
                  if (StrToInt(Cells[C_BD_NO3,    i+FixedRows]) >= 1)  and
                     (StrToInt(Cells[C_BD_NO3,    i+FixedRows]) <= 10) then
                  begin
                     Colors[C_BD_NO3,    i+FixedRows] := clYellow;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NO3,    i+FixedRows]) >= 11)  and
                     (StrToInt(Cells[C_BD_NO3,    i+FixedRows]) <= 20) then
                  begin
                     Colors[C_BD_NO3,    i+FixedRows] := $00E09000;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NO3,    i+FixedRows]) >= 21)  and
                     (StrToInt(Cells[C_BD_NO3,    i+FixedRows]) <= 30) then
                  begin
                     Colors[C_BD_NO3,    i+FixedRows] := clRed;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NO3,    i+FixedRows]) >= 31)  and
                     (StrToInt(Cells[C_BD_NO3,    i+FixedRows]) <= 40) then
                  begin
                     Colors[C_BD_NO3,    i+FixedRows] := clGray;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NO3,    i+FixedRows]) >= 41)  and
                     (StrToInt(Cells[C_BD_NO3,    i+FixedRows]) <= 45) then
                  begin
                     Colors[C_BD_NO3,    i+FixedRows] := clGreen;
                  end;


                  // No4. Range 구간별 Coloring
                  if (StrToInt(Cells[C_BD_NO4,    i+FixedRows]) >= 1)  and
                     (StrToInt(Cells[C_BD_NO4,    i+FixedRows]) <= 10) then
                  begin
                     Colors[C_BD_NO4,    i+FixedRows] := clYellow;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NO4,    i+FixedRows]) >= 11)  and
                     (StrToInt(Cells[C_BD_NO4,    i+FixedRows]) <= 20) then
                  begin
                     Colors[C_BD_NO4,    i+FixedRows] := $00E09000;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NO4,    i+FixedRows]) >= 21)  and
                     (StrToInt(Cells[C_BD_NO4,    i+FixedRows]) <= 30) then
                  begin
                     Colors[C_BD_NO4,    i+FixedRows] := clRed;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NO4,    i+FixedRows]) >= 31)  and
                     (StrToInt(Cells[C_BD_NO4,    i+FixedRows]) <= 40) then
                  begin
                     Colors[C_BD_NO4,    i+FixedRows] := clGray;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NO4,    i+FixedRows]) >= 41)  and
                     (StrToInt(Cells[C_BD_NO4,    i+FixedRows]) <= 45) then
                  begin
                     Colors[C_BD_NO4,    i+FixedRows] := clGreen;
                  end;


                  // No5. Range 구간별 Coloring
                  if (StrToInt(Cells[C_BD_NO5,    i+FixedRows]) >= 1)  and
                     (StrToInt(Cells[C_BD_NO5,    i+FixedRows]) <= 10) then
                  begin
                     Colors[C_BD_NO5,    i+FixedRows] := clYellow;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NO5,    i+FixedRows]) >= 11)  and
                     (StrToInt(Cells[C_BD_NO5,    i+FixedRows]) <= 20) then
                  begin
                     Colors[C_BD_NO5,    i+FixedRows] := $00E09000;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NO5,    i+FixedRows]) >= 21)  and
                     (StrToInt(Cells[C_BD_NO5,    i+FixedRows]) <= 30) then
                  begin
                     Colors[C_BD_NO5,    i+FixedRows] := clRed;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NO5,    i+FixedRows]) >= 31)  and
                     (StrToInt(Cells[C_BD_NO5,    i+FixedRows]) <= 40) then
                  begin
                     Colors[C_BD_NO5,    i+FixedRows] := clGray;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NO5,    i+FixedRows]) >= 41)  and
                     (StrToInt(Cells[C_BD_NO5,    i+FixedRows]) <= 45) then
                  begin
                     Colors[C_BD_NO5,    i+FixedRows] := clGreen;
                  end;


                  // No6. Range 구간별 Coloring
                  if (StrToInt(Cells[C_BD_NO6,    i+FixedRows]) >= 1)  and
                     (StrToInt(Cells[C_BD_NO6,    i+FixedRows]) <= 10) then
                  begin
                     Colors[C_BD_NO6,    i+FixedRows] := clYellow;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NO6,    i+FixedRows]) >= 11)  and
                     (StrToInt(Cells[C_BD_NO6,    i+FixedRows]) <= 20) then
                  begin
                     Colors[C_BD_NO6,    i+FixedRows] := $00E09000;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NO6,    i+FixedRows]) >= 21)  and
                     (StrToInt(Cells[C_BD_NO6,    i+FixedRows]) <= 30) then
                  begin
                     Colors[C_BD_NO6,    i+FixedRows] := clRed;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NO6,    i+FixedRows]) >= 31)  and
                     (StrToInt(Cells[C_BD_NO6,    i+FixedRows]) <= 40) then
                  begin
                     Colors[C_BD_NO6,    i+FixedRows] := clGray;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NO6,    i+FixedRows]) >= 41)  and
                     (StrToInt(Cells[C_BD_NO6,    i+FixedRows]) <= 45) then
                  begin
                     Colors[C_BD_NO6,    i+FixedRows] := clGreen;
                  end;

                  // No (보너스). Range 구간별 Coloring
                  if (StrToInt(Cells[C_BD_NOBONUS,    i+FixedRows]) >= 1)  and
                     (StrToInt(Cells[C_BD_NOBONUS,    i+FixedRows]) <= 10) then
                  begin
                     Colors[C_BD_NOBONUS,    i+FixedRows] := clYellow;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NOBONUS,    i+FixedRows]) >= 11)  and
                     (StrToInt(Cells[C_BD_NOBONUS,    i+FixedRows]) <= 20) then
                  begin
                     Colors[C_BD_NOBONUS,    i+FixedRows] := $00E09000;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NOBONUS,    i+FixedRows]) >= 21)  and
                     (StrToInt(Cells[C_BD_NOBONUS,    i+FixedRows]) <= 30) then
                  begin
                     Colors[C_BD_NOBONUS,    i+FixedRows] := clRed;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NOBONUS,    i+FixedRows]) >= 31)  and
                     (StrToInt(Cells[C_BD_NOBONUS,    i+FixedRows]) <= 40) then
                  begin
                     Colors[C_BD_NOBONUS,    i+FixedRows] := clGray;
                  end
                  else
                  if (StrToInt(Cells[C_BD_NOBONUS,    i+FixedRows]) >= 41)  and
                     (StrToInt(Cells[C_BD_NOBONUS,    i+FixedRows]) <= 45) then
                  begin
                     Colors[C_BD_NOBONUS,    i+FixedRows] := clGreen;
                  end;
               end;
            end;


            // RowCount 정리
            RowCount := RowCount - 1;


            // Comments
            //lb_RegDoc.Caption  := '▶ ' + IntToStr(iRowCnt) + '건의 릴리즈 내역을 조회하였습니다.';

         end;
      end;

   finally
      FreeAndNil(TpGetSearch);
      Screen.Cursor := crDefault;
   end;
end;




procedure TMainDlg.asg_Master_DblClickCell(Sender: TObject; ARow,
  ACol: Integer);
begin
   with asg_Master do
   begin
      // 종료일자 Toggle 기능
      if (ACol = C_M_DELDATE) then
         if (Cells[ACol, ARow] <> '') then
            Cells[ACol, ARow] := ''
         else
            Cells[ACol, ARow] := FormatDateTime('yyyy-mm-dd', Date);

      // 현재 사용자 IP 정보 조회
      if (ACol = C_M_USERIP) then
         Cells[ACol, ARow] := FsUserIp;

   end;
end;



procedure TMainDlg.asg_Master_CanEditCell(Sender: TObject; ARow,
  ACol: Integer; var CanEdit: Boolean);
begin
   // 종료일자, Update, 사용자 추가는 Edit 제한
   CanEdit := (ACol <> C_M_DELDATE) and
              (ACol <> C_M_BUTTON)  and
              (ACol <> C_M_USERADD);
end;



procedure TMainDlg.asg_Master_ClickCell(Sender: TObject; ARow,
  ACol: Integer);
begin
   with asg_Master do
   begin
      if (ACol = C_M_USERIP) then
         lb_Network.Caption := '▶ 담당자 IP를 등록하시면, 원활한 업무소통(예: 훈민메모,챗박스 등)에 도움을 받으실 수 있습니다.'
      else if (ACol = C_M_USERID) then
         lb_Network.Caption := '▶ 사번이 없으신 분은 입력 안하셔도 괜찮습니다.'
      else if (ACol = C_M_DELDATE) then
         lb_Network.Caption := '▶ 선택한 담당자정보를 삭제하기 원하시면, 종료일자를 [더블클릭]하여 업데이트 해주세요.'
      else
         lb_Network.Caption := '';
   end;
end;



//------------------------------------------------------------------------------
// [출력] 시용자 정보 List 출력
//
// Author : Lee, Se-Ha
// Date   : 2013.08.08
//------------------------------------------------------------------------------
procedure TMainDlg.bbt_MasterPrintClick(Sender: TObject);
var
   varResult : String;
begin
   if not IsLogonUser then
   begin
      MessageBox(self.Handle,
                 PChar('[비상연락망] - [담당자정보]에 담당자 IP 등록후 사용하실 수 있습니다.'),
                 PChar(Self.Caption + ' : 등록되지 않은 User 제한 알림'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;

   //-------------------------------------------------------------------
   // 2. Set Print Options
   //-------------------------------------------------------------------
   SetPrintOptions;


   asg_Master.HideColumn(C_M_DELDATE);
   asg_Master.HideColumn(C_M_BUTTON);
   asg_Master.HideColumn(C_M_USERADD);
   asg_Master.HideColumn(C_M_NICKNM);


   //-------------------------------------------------------------------
   // 3. Print
   //-------------------------------------------------------------------
   asg_Master.Print;


   asg_Master.UnHideColumn(C_M_DELDATE);
   asg_Master.UnHideColumn(C_M_BUTTON);
   asg_Master.UnHideColumn(C_M_USERADD);
   asg_Master.UnHideColumn(C_M_NICKNM);


   // 로그 Update
   UpdateLog('MASTER',
             'PRINT',
             FsUserIp,
             'P',
             FsUserSpec,
             FsUserNm,
             varResult
             );


   // Comments
   lb_Network.Caption := '▶ 담당자정보 출력이 완료되었습니다.';
end;



//------------------------------------------------------------------------------
// [프린트옵션] Set Print Options Handler
//
// Author : Lee, Se-Ha
// Date   : 2013.08.08
//------------------------------------------------------------------------------
procedure TMainDlg.SetPrintOptions;
begin
   with asg_Master do
   begin
      //-----------------------------------------------------------------
      // 1-1. Set Default Print Options
      //-----------------------------------------------------------------
      PrintSettings.Borders      := AdvGrid.pbSingle;
      Printsettings.FitToPage    := fpShrink;
      Printsettings.NoAutoSize   := True;
      Printsettings.Centered     := True;
      printsettings.Headersize   := 150;
   end;

   with asg_Detail do
   begin
      //-----------------------------------------------------------------
      // 2-1. Set Default Print Options
      //-----------------------------------------------------------------
      PrintSettings.Borders      := AdvGrid.pbSingle;
      Printsettings.FitToPage    := fpShrink;
      Printsettings.NoAutoSize   := True;
      Printsettings.Centered     := True;
      printsettings.Headersize   := 150;
   end;

   with asg_Network do
   begin
      //-----------------------------------------------------------------
      // 3-1. Set Default Print Options
      //-----------------------------------------------------------------
      PrintSettings.Borders      := AdvGrid.pbSingle;
      Printsettings.FitToPage    := fpShrink;
      Printsettings.NoAutoSize   := True;
      Printsettings.Centered     := True;
      printsettings.Headersize   := 230;
   end;

   with asg_WorkRpt do
   begin
      //-----------------------------------------------------------------
      // 4-1. Set Default Print Options
      //-----------------------------------------------------------------
      PrintSettings.Borders      := AdvGrid.pbSingle;
      Printsettings.FitToPage    := fpAlways;
      Printsettings.NoAutoSize   := True;
      Printsettings.Centered     := True;
      printsettings.Headersize   := 150;
   end;

   with asg_WeeklyRpt do
   begin
      //-----------------------------------------------------------------
      // 5-1. Set Default Print Options
      //-----------------------------------------------------------------
      PrintSettings.Borders      := AdvGrid.pbSingle;
      Printsettings.FitToPage    := fpAlways;
      Printsettings.NoAutoSize   := True;
      Printsettings.Centered     := True;
      printsettings.Headersize   := 150;
   end;

   with asg_DialMap do
   begin
      //-----------------------------------------------------------------
      // 5-1. Set Default Print Options
      //-----------------------------------------------------------------
      PrintSettings.Borders      := AdvGrid.pbSingle;
      Printsettings.FitToPage    := fpCustom;
      Printsettings.NoAutoSize   := True;
      Printsettings.Centered     := True;
      printsettings.Headersize   := 300;
   end;
end;




//------------------------------------------------------------------------------
// AdvStringGrid onPrintPage Event Handler
//
// Date   : 2013.08.09
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_Master_PrintPage(Sender: TObject; Canvas: TCanvas;
  PageNr, PageXSize, PageYSize: Integer);
var
   savefont : tfont;
   ts,tw    : integer;
   RecName  : String;
const
   myowntitle : string = ' ';
begin
   if asg_Master.PrintColStart <> 0 then exit;


   // Report Name 정의
   RecName := '정보전산팀 담당자 정보';


   // 의료원명 Display
   with canvas do
   begin
      savefont := TFont.Create;
      savefont.assign(font);

      font.name   := 'Arial';
      font.Size   := 8;
      font.Style  := [];
      font.height := asg_Master.mapfontheight(8);
      font.Color  := clBlack;

      ts := asg_Master.Printcoloffset[0];
      tw := asg_Master.Printpagewidth;
      ts := ts+ ((tw-textwidth('고려대학교의료원')) shr 1);

      Textout(ts, -15, '고려대학교의료원');

      font.assign(savefont);

      savefont.free;
   end;


   // 기록지명 Display
   with canvas do
   begin
      savefont := TFont.Create;
      savefont.assign(font);

      font.name   := 'Arial';
      font.Size   := 15;
      font.Style  := [fsBold];
      font.height := asg_Master.mapfontheight(15);
      font.Color  := clBlack;

      ts := asg_Master.Printcoloffset[0];
      tw := asg_Master.Printpagewidth;
      ts := ts+ ((tw-textwidth(RecName)) shr 1);

      Textout(ts, -50, RecName);

      font.assign(savefont);

      savefont.free;
   end;
end;



procedure TMainDlg.asg_Master_PrintStart(Sender: TObject;
  NrOfPages: Integer; var FromPage, ToPage: Integer);
begin
   Printdialog1.FromPage   := Frompage;
   Printdialog1.ToPage     := ToPage;
   Printdialog1.Maxpage    := ToPage;
   Printdialog1.minpage    := 1;

   if Printdialog1.execute then
   begin
      Frompage := Printdialog1.FromPage;
      Topage   := Printdialog1.ToPage;
   end
   else
   begin
      Frompage := 0;
      Topage   := 0;
   end;
end;




procedure TMainDlg.asg_Detail_ButtonClick(Sender: TObject; ACol,
  ARow: Integer);
var
   TpSetDetl   : TTpSvc;
   i, iCnt     : Integer;
   vType1,
   vLocate,
   vDutyPart,
   vDutySpec,
   vDutyRmk ,
   vDutyUser,
   vDutyPtnr,
   vCallNo  ,
   vDeptCd,
   vDelDate,
   vEditIp : Variant;
   sType   : String;
begin

   // 개인별 정보변경 Update
   if (ACol = C_D_BUTTON) then
   begin

      //
      sType := '2';
      iCnt  := 0;


      //-------------------------------------------------------------------
      // 2. Create Variables
      //-------------------------------------------------------------------
      with asg_Detail do
      begin

         if (Cells[C_D_LOCATE,   ARow] = '') or
            (Cells[C_D_DUTYUSER, ARow] = '') or
            (Cells[C_D_DUTYPART, ARow] = '') or
            (Cells[C_D_CALLNO,   ARow] = '') then
         begin
            MessageBox(self.Handle,
                       PChar('[근무처/업무구분/담당자/연락처]는 필수입력 항목입니다.'),
                       PChar(Self.Caption + ' : 업무프로필 필수입력 항목 알림'),
                       MB_OK + MB_ICONWARNING);

            Exit;
         end;


         if ((PosByte('운영', Cells[C_D_DUTYPART, ARow]) > 0) or
             (PosByte('기획', Cells[C_D_DUTYPART, ARow]) > 0)) and
            (Cells[C_D_DUTYSPEC, ARow] = '') then
         begin
            MessageBox(self.Handle,
                       PChar('[업무상세]를 입력해 주시기 바랍니다.'),
                       PChar(Self.Caption + ' : 업무프로필 필수입력 항목 알림'),
                       MB_OK + MB_ICONWARNING);

            Exit;
         end;

         if (PosByte('개발', Cells[C_D_DUTYPART, ARow]) > 0) and
            ((Cells[C_D_DUTYSPEC, ARow] = '') or
             (Cells[C_D_DUTYRMK,  ARow] = '') or
             (Cells[C_D_DUTYPTNR, ARow] = '')) then
         begin
            MessageBox(self.Handle,
                       PChar('개발파트의 경우 [업무상세/Remark/파트너쉽]은 필수입력 항목입니다.'),
                       PChar(Self.Caption + ' : 업무프로필 필수입력 항목 알림'),
                       MB_OK + MB_ICONWARNING);

            Exit;
         end;


         for i := ARow to ARow do
         begin
            //------------------------------------------------------------------
            // 2-1. Append Variables
            //------------------------------------------------------------------
            AppendVariant(vType1   ,   sType                  );
            AppendVariant(vLocate  ,   Cells[C_D_LOCATE,    i]);
            AppendVariant(vDutyPart,   Cells[C_D_DUTYPART,  i]);
            AppendVariant(vDutySpec,   Cells[C_D_DUTYSPEC,  i]);
            AppendVariant(vDutyRmk ,   Cells[C_D_DUTYRMK,   i]);
            AppendVariant(vDutyUser,   Cells[C_D_DUTYUSER,  i]);
            AppendVariant(vCallNo  ,   Cells[C_D_CALLNO,    i]);
            AppendVariant(vDutyPtnr,   Cells[C_D_DUTYPTNR,  i]);
            AppendVariant(vDelDate ,   Cells[C_D_DELDATE,   i]);
            AppendVariant(vEditIp  ,   FsUserIp               );
            AppendVariant(vDeptCd  ,   'ITT'                  );

            Inc(iCnt);
         end;
      end;



      if iCnt <= 0 then
         Exit;

      //-------------------------------------------------------------------
      // 3. Insert by TpSvc
      //-------------------------------------------------------------------
      TpSetDetl := TTpSvc.Create;
      TpSetDetl.Init(Self);


      Screen.Cursor := crHourGlass;


      try
         if TpSetDetl.PutSvc('MD_KUMCM_M1',
                             [
                              'S_TYPE1'   , vType1
                            , 'S_STRING1' , vLocate
                            , 'S_STRING4' , vDutyPart
                            , 'S_STRING8' , vEditIp
                            , 'S_STRING9' , vDeptCd
                            , 'S_STRING10', vDelDate
                            , 'S_STRING11', vDutySpec
                            , 'S_STRING12', vDutyRmk
                            , 'S_STRING13', vDutyUser
                            , 'S_STRING14', vDutyPtnr
                            , 'S_STRING15', vCallNo
                            ] ) then
         begin
            MessageBox(self.Handle,
                       PChar(asg_Detail.Cells[C_D_DUTYPART, asg_Detail.Row] + '('  +
                             asg_Detail.Cells[C_D_DUTYSPEC, asg_Detail.Row] + ') ' +
                             asg_Detail.Cells[C_D_DUTYUSER, asg_Detail.Row] +' 님의 정보가 정상적으로 [업데이트]되었습니다.'),
                       '[KUMC 다이얼로그] 업무프로필 업데이트 알림 ',
                       MB_OK + MB_ICONINFORMATION);


            // Refresh
            SelGridInfo('DETAIL');
         end
         else
         begin
            ShowMessageM(GetTxMsg);
         end;

      finally
         FreeAndNil(TpSetDetl);
         Screen.Cursor  := crDefault;
      end;
   end;




   if (ACol = C_D_USERADD) then
   begin
      asg_Detail.InsertRows(ARow + 1, 1);
      asg_Detail.AddButton(C_D_USERADD, ARow + 1, asg_Detail.ColWidths[C_D_USERADD]-5, 20, '+', haBeforeText, vaCenter);                // 사용자 Add
   end;

end;




procedure TMainDlg.asg_Detail_CanEditCell(Sender: TObject; ARow,
  ACol: Integer; var CanEdit: Boolean);
begin
   // 종료일자, Update, 사용자 추가는 Edit 제한
   CanEdit := (ACol <> C_D_DELDATE) and
              (ACol <> C_D_BUTTON)  and
              (ACol <> C_D_USERADD);
end;



procedure TMainDlg.asg_Detail_DblClickCell(Sender: TObject; ARow,
  ACol: Integer);
begin
   with asg_Detail do
   begin
      // 종료일자 Toggle 기능
      if (ACol = C_D_DELDATE) then
         if (Cells[ACol, ARow] <> '') then
            Cells[ACol, ARow] := ''
         else
            Cells[ACol, ARow] := FormatDateTime('yyyy-mm-dd', Date);

   end;
end;





procedure TMainDlg.asg_Detail_GetEditorType(Sender: TObject; ACol,
  ARow: Integer; var AEditor: TEditorType);
begin
   with asg_Detail do
   begin
      if (ARow > 0) and (ACol = C_D_LOCATE) then
      begin
         AEditor := edComboList;

         ClearComboString;

         AddComboString('의료원');
         AddComboString('안암');
         AddComboString('구로');
         AddComboString('안산');
      end;

      if (ARow > 0) and (PosByte('의료원', Cells[C_D_LOCATE, ARow]) > 0) and
                        (ACol = C_D_DUTYPART) then
      begin
         AEditor := edComboList;

         ClearComboString;

         AddComboString('관리자');
         AddComboString('홈페이지');
      end;


      if (ARow > 0) and (PosByte('안암', Cells[C_D_LOCATE, ARow]) > 0) and
                        (ACol = C_D_DUTYPART) then
      begin
         AEditor := edComboList;

         ClearComboString;

         AddComboString('운영파트');
         AddComboString('개발파트');
         AddComboString('기획파트');
         AddComboString('OP');
         AddComboString('PC/AS');
         AddComboString('PACS');
      end;


      if (ARow > 0) and ((PosByte('구로', Cells[C_D_LOCATE, ARow]) > 0) or
                         (PosByte('안산', Cells[C_D_LOCATE, ARow]) > 0)) and
                        (ACol = C_D_DUTYPART) then
      begin
         AEditor := edComboList;

         ClearComboString;

         AddComboString('운영파트');
         AddComboString('OP');
         AddComboString('PC/AS');
         AddComboString('PACS');
      end;

      if (ARow > 0) and (PosByte('운영', Cells[C_D_DUTYPART, ARow]) > 0)
                    and (ACol = C_D_DUTYSPEC) then
      begin
         AEditor := edComboList;

         ClearComboString;

         AddComboString('파트장');
         AddComboString('PC관리');
         AddComboString('네트워크');
         AddComboString('시스템');
         AddComboString('D/B');
      end;

      if (ARow > 0) and (PosByte('개발', Cells[C_D_DUTYPART, ARow]) > 0)
                    and (ACol = C_D_DUTYSPEC) then
      begin
         AEditor := edComboList;

         ClearComboString;

         AddComboString('파트장');  // 개발파트장 추가 @ 2016.07.26 LSH
         AddComboString('원무');
         AddComboString('진료');
         AddComboString('진료지원');
         AddComboString('일반관리');
         AddComboString('EMR');
         AddComboString('기타');
      end;

      if (ARow > 0) and (PosByte('개발',  Cells[C_D_DUTYPART, ARow]) > 0)
                    and (ACol = C_D_DUTYPTNR) then
      begin
         AEditor := edComboList;

         ClearComboString;

         AddComboString('P/L');
         AddComboString('의료원');
         AddComboString('파트너쉽');
      end;



      if (ACol = C_D_DELDATE) then
      begin
         AEditor := edDateEdit;
      end;
   end;
end;



procedure TMainDlg.asg_Detail_ClickCell(Sender: TObject; ARow,
  ACol: Integer);
begin

   with asg_Detail do
   begin
      if (ACol = C_D_DUTYRMK) then
         lb_Network.Caption := '▶ 담당하고 있는 업무상세내역(예: 개발파트 - 청구/외래,병동간호/약국 등)을 적어주세요.'
      else if (ACol = C_D_DUTYPTNR) then
         lb_Network.Caption := '▶ 업무구분이 [개발파트] 필수입력항목으로, P/L-의료원-파트너쉽(협력업체)을 표기합니다.'
      else if (ACol = C_D_DELDATE) then
         lb_Network.Caption := '▶ 선택한 업무프로필을 삭제하기 원하시면, 종료일자를 [더블클릭]하여 업데이트 해주세요.'
      else
         lb_Network.Caption := '';
   end;

end;



procedure TMainDlg.asg_Detail_EditingDone(Sender: TObject);
begin
   with asg_Detail do
   begin
      if (Cells[C_D_LOCATE,   Row] <> '') and
         (Cells[C_D_DUTYUSER, Row] <> '') and
         (Cells[C_D_DUTYPART, Row] <> '') and
         (Cells[C_D_CALLNO,   Row] <> '') then
      begin
         AddButton(C_D_BUTTON,  Row, ColWidths[C_D_BUTTON]-5,  20, 'Update', haBeforeText, vaCenter);           // Update
      end;
   end;
end;




procedure TMainDlg.bbt_DetailPrintClick(Sender: TObject);
var
   varResult : String;
begin
   if not IsLogonUser then
   begin
      MessageBox(self.Handle,
                 PChar('[비상연락망] - [담당자정보]에 담당자 IP 등록후 사용하실 수 있습니다.'),
                 PChar(Self.Caption + ' : 등록되지 않은 User 제한 알림'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;

   //-------------------------------------------------------------------
   // 2. Set Print Options
   //-------------------------------------------------------------------
   SetPrintOptions;


   asg_Detail.HideColumn(C_D_DELDATE);
   asg_Detail.HideColumn(C_D_BUTTON);
   asg_Detail.HideColumn(C_D_USERADD);


   //-------------------------------------------------------------------
   // 3. Print
   //-------------------------------------------------------------------
   asg_Detail.Print;


   asg_Detail.UnHideColumn(C_D_DELDATE);
   asg_Detail.UnHideColumn(C_D_BUTTON);
   asg_Detail.UnHideColumn(C_D_USERADD);


   // 로그 Update
   UpdateLog('DETAIL',
             'PRINT',
             FsUserIp,
             'P',
             FsUserSpec,
             FsUserNm,
             varResult
             );


   // Comments
   lb_Network.Caption := '▶ 업무프로필 출력이 완료되었습니다.';

end;




//------------------------------------------------------------------------------
// AdvStringGrid onPrintPage Event Handler
//
// Date   : 2013.08.09
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_Detail_PrintPage(Sender: TObject; Canvas: TCanvas;
  PageNr, PageXSize, PageYSize: Integer);
var
   savefont : tfont;
   ts,tw    : integer;
   RecName  : String;
const
   myowntitle : string = ' ';
begin
   if asg_Master.PrintColStart <> 0 then exit;


   // Report Name 정의
   RecName := '정보전산팀 업무 프로필';


   // 의료원명 Display
   with canvas do
   begin
      savefont := TFont.Create;
      savefont.assign(font);

      font.name   := 'Arial';
      font.Size   := 8;
      font.Style  := [];
      font.height := asg_Detail.mapfontheight(8);
      font.Color  := clBlack;

      ts := asg_Detail.Printcoloffset[0];
      tw := asg_Detail.Printpagewidth;
      ts := ts+ ((tw-textwidth('고려대학교의료원')) shr 1);

      Textout(ts, -15, '고려대학교의료원');

      font.assign(savefont);

      savefont.free;
   end;


   // 기록지명 Display
   with canvas do
   begin
      savefont := TFont.Create;
      savefont.assign(font);

      font.name   := 'Arial';
      font.Size   := 15;
      font.Style  := [fsBold];
      font.height := asg_Detail.mapfontheight(15);
      font.Color  := clBlack;

      ts := asg_Detail.Printcoloffset[0];
      tw := asg_Detail.Printpagewidth;
      ts := ts+ ((tw-textwidth(RecName)) shr 1);

      Textout(ts, -50, RecName);

      font.assign(savefont);

      savefont.free;
   end;
end;


//------------------------------------------------------------------------------
// AdvStringGrid onPrintStart Event Handler
//
// Date   : 2013.08.09
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_Detail_PrintStart(Sender: TObject;
  NrOfPages: Integer; var FromPage, ToPage: Integer);
begin
   Printdialog1.FromPage   := Frompage;
   Printdialog1.ToPage     := ToPage;
   Printdialog1.Maxpage    := ToPage;
   Printdialog1.minpage    := 1;

   if Printdialog1.execute then
   begin
      Frompage := Printdialog1.FromPage;
      Topage   := Printdialog1.ToPage;
   end
   else
   begin
      Frompage := 0;
      Topage   := 0;
   end;
end;



procedure TMainDlg.bbt_MasterRefreshClick(Sender: TObject);
begin
   //------------------------------------------------------------------
   //  Data Loading bar 보이기
   //------------------------------------------------------------------
   //SetLoadingBar('ON');


   SelGridInfo('MASTER');


   //------------------------------------------------------------------
   //  Data Loading bar 감추기
   //------------------------------------------------------------------
   //SetLoadingBar('OFF');
end;



procedure TMainDlg.bbt_DetailRefreshClick(Sender: TObject);
begin
   SelGridInfo('DETAIL');
end;



procedure TMainDlg.bbt_NetworkRefreshClick(Sender: TObject);
begin
   if (apn_Network.Visible) then
   begin
      SelGridInfo('UPDATE');
      SelGridInfo('DIALOG');
   end
   else if (apn_Detail.Visible) then
   begin
      SelGridInfo('DETAIL');
   end
   else if (apn_Master.Visible) then
   begin
      SelGridInfo('MASTER');
   end;
end;


//------------------------------------------------------------------------------
// [출력] Button onClick Event Handler
//
// Date   : 2013.08.12
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.bbt_NetworkPrintClick(Sender: TObject);
var
   varResult : String;
begin
   if not IsLogonUser then
   begin
      MessageBox(self.Handle,
                 PChar('[비상연락망] - [담당자정보]에 담당자 IP 등록후 사용하실 수 있습니다.'),
                 PChar(Self.Caption + ' : 등록되지 않은 User 제한 알림'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;


   //-------------------------------------------------------------------
   // 2. Set Print Options
   //-------------------------------------------------------------------
   SetPrintOptions;


   if (apn_Network.Visible) then
   begin
      asg_NetWork.ColWidths[C_NW_DUTYRMK]   := 200;

      //-------------------------------------------------------------------
      // 3. Print
      //-------------------------------------------------------------------
      asg_NetWork.Print;



      // 로그 Update
      UpdateLog('NETWORK',
                'PRINT',
                FsUserIp,
                'P',
                FsUserSpec,
                FsUserNm,
                varResult
                );


      asg_NetWork.ColWidths[C_NW_DUTYRMK]   := 125;

      // Comments
      lb_Network.Caption := '▶ 비상연락망 출력이 완료되었습니다.';

   end
   else if (apn_Detail.Visible) then
   begin
      asg_Detail.HideColumn(C_D_DELDATE);
      asg_Detail.HideColumn(C_D_BUTTON);
      asg_Detail.HideColumn(C_D_USERADD);


      //-------------------------------------------------------------------
      // 3. Print
      //-------------------------------------------------------------------
      asg_Detail.Print;


      asg_Detail.UnHideColumn(C_D_DELDATE);
      asg_Detail.UnHideColumn(C_D_BUTTON);
      asg_Detail.UnHideColumn(C_D_USERADD);


      // 로그 Update
      UpdateLog('DETAIL',
                'PRINT',
                FsUserIp,
                'P',
                FsUserSpec,
                FsUserNm,
                varResult
                );


      // Comments
      lb_Network.Caption := '▶ 업무프로필 출력이 완료되었습니다.';

   end
   else  if (apn_Master.Visible) then
   begin
      asg_Master.HideColumn(C_M_DELDATE);
      asg_Master.HideColumn(C_M_BUTTON);
      asg_Master.HideColumn(C_M_USERADD);
      asg_Master.HideColumn(C_M_NICKNM);


      //-------------------------------------------------------------------
      // 3. Print
      //-------------------------------------------------------------------
      asg_Master.Print;


      asg_Master.UnHideColumn(C_M_DELDATE);
      asg_Master.UnHideColumn(C_M_BUTTON);
      asg_Master.UnHideColumn(C_M_USERADD);
      asg_Master.UnHideColumn(C_M_NICKNM);


      // 로그 Update
      UpdateLog('MASTER',
                'PRINT',
                FsUserIp,
                'P',
                FsUserSpec,
                FsUserNm,
                varResult
                );


      // Comments
      lb_Network.Caption := '▶ 담당자정보 출력이 완료되었습니다.';

   end;

end;



//------------------------------------------------------------------------------
// AdvStringGrid onPrintPage Event Handler
//
// Date   : 2013.08.09
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_Network_PrintPage(Sender: TObject; Canvas: TCanvas;
  PageNr, PageXSize, PageYSize: Integer);
var
   savefont : tfont;
   ts,tw    : integer;
   RecName, addInfo  : String;
const
   myowntitle : string = ' ';
begin
   if asg_Network.PrintColStart <> 0 then exit;


   // Report Name 정의
   RecName := '정보전산실 비상연락망';
   addInfo := ' Fax no. [안암] 02-920-5672 [구로] 02-2626-2239 [안산] 031-412-5999';


   // 의료원명 Display
   with canvas do
   begin
      savefont := TFont.Create;
      savefont.assign(font);

      font.name   := 'Arial';
      font.Size   := 8;
      font.Style  := [];
      font.height := asg_Network.mapfontheight(8);
      font.Color  := clBlack;

      ts := asg_Network.Printcoloffset[0];
      tw := asg_Network.Printpagewidth;
      ts := ts+ ((tw-textwidth('고려대학교의료원')) shr 1);

      Textout(ts, -15, '고려대학교의료원');

      font.assign(savefont);

      savefont.free;
   end;


   // 기록지명 Display
   with canvas do
   begin
      savefont := TFont.Create;
      savefont.assign(font);

      font.name   := 'Arial';
      font.Size   := 15;
      font.Style  := [fsBold];
      font.height := asg_Network.mapfontheight(15);
      font.Color  := clBlack;

      ts := asg_Network.Printcoloffset[0];
      tw := asg_Network.Printpagewidth;
      ts := ts+ ((tw-textwidth(RecName)) shr 1);

      Textout(ts, -50, RecName);

      font.assign(savefont);

      savefont.free;
   end;

   // 최종수정 ver. Display
   with canvas do
   begin
      savefont := TFont.Create;
      savefont.assign(font);

      font.name   := 'Arial';
      font.Size   := 9;
      font.Style  := [];
      font.height := asg_Network.mapfontheight(9);
      font.Color  := clBlack;

      ts := asg_Network.Printcoloffset[0];
      tw := asg_Network.Printpagewidth;
      ts := ts+ ((tw-textwidth('[' + asg_NwUpdate.Cells[C_NU_EDITDATE, 1] + ' Updated]')) shr 1);

      Textout(ts, -115, '[' + asg_NwUpdate.Cells[C_NU_EDITDATE, 1] + ' Updated]');

      font.assign(savefont);

      savefont.free;
   end;


   // 병원별 Fax. Info Display
   with canvas do
   begin
      savefont := TFont.Create;
      savefont.assign(font);

      font.name   := 'Arial';
      font.Size   := 9;
      font.Style  := [];
      font.height := asg_Network.mapfontheight(9);
      font.Color  := clBlack;

      ts := asg_Network.Printcoloffset[0];
      tw := asg_Network.Printpagewidth;
      ts := ts+ ((tw-textwidth(addInfo)) shr 1);

      Textout(ts + 480, -180, addInfo);

      font.assign(savefont);

      savefont.free;
   end;

end;


//------------------------------------------------------------------------------
// AdvStringGrid onPrintStart Event Handler
//
// Date   : 2013.08.09
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_Network_PrintStart(Sender: TObject;
  NrOfPages: Integer; var FromPage, ToPage: Integer);
begin
   Printdialog1.FromPage   := Frompage;
   Printdialog1.ToPage     := ToPage;
   Printdialog1.Maxpage    := ToPage;
   Printdialog1.minpage    := 1;

   if Printdialog1.execute then
   begin
      Frompage := Printdialog1.FromPage;
      Topage   := Printdialog1.ToPage;
   end
   else
   begin
      Frompage := 0;
      Topage   := 0;
   end;
end;


//------------------------------------------------------------------------------
// Form onClick Event Handler - Test.
//
// Date   : 2013.08.09
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.BitBtn1Click(Sender: TObject);
//var
//   FForm : TForm;
begin

   // 공지사항 BPL 테스트
   //try
      {
      FForm := BplFormCreate('INFORM');

      FForm.Position := poMainFormCenter;

      FForm.Show;
      }

   //-----------------------------------------------------
   // 등록된 User 다이얼 Chat 리스트 팝업 @ 2015.03.30 LSH
   //-----------------------------------------------------
   {  -- Chat XE7 개선전까지 주석 @ 2017.10.31 LSH
   if (IsLogonUser) then
   begin
      FForm := BplFormCreate('KDIALCHAT', True);


      if FForm <> nil then
      begin
         SetBplStrProp(FForm, 'prop_UserNm', FsUserNm);

         FForm.Height := (self.Height);
         FForm.Top    := (screen.Height - FForm.Height) - 48;
         FForm.Left   := 8;

         try
            FForm.Show;

         except
            on e : Exception do
            showmessage('Show Form Error(KDIALCHAT) : ' + e.Message);
         end;

         self.SetFocus;

      end;
   end;
   }

  // except
  //    FForm.Free;
  // end;
end;





//------------------------------------------------------------------------------
// [조회] 각 Grid 별 메세지 정보
//    - 진료 공통코드(MDMCOMCT.COMCD1 = 'KDIAL',
//                             COMCD2 : 각 화면 구분자
//                             COMCD3 : 메세지 종류 구분자
//
// Author : Lee, Se-Ha
// Date   : 2013.08.12
//------------------------------------------------------------------------------
function TMainDlg.GetMsgInfo(in_ComCd2, in_ComCd3 : String) : String;
var
//   i : integer;
   TpGetColumn : TTpSvc;
begin
   // 리턴값 초기화
   GetMsgInfo := '';



   Screen.Cursor := crHourglass;



   //-----------------------------------------------------------------
   // 2-1. Load Data by TpSvc
   //-----------------------------------------------------------------
   TpGetColumn := TTpSvc.Create;
   TpGetColumn.Init(Self);


   Screen.Cursor := crHourglass;


   try
      TpGetColumn.CountField  := 'S_STRING1';
      TpGetColumn.ShowMsgFlag := False;

      if TpGetColumn.GetSvc('MD_MCOMC_L1',
                          ['S_TYPE1'  , 'I'
                         , 'S_CODE1'  , 'KDIAL'
                         , 'S_CODE2'  , in_ComCd2
                         , 'S_CODE3'  , in_ComCd3
                          ],
                          [
                           'L_CNT1'    , 'iCitem06'
                         , 'L_CNT2'    , 'iCitem07'
                         , 'L_CNT3'    , 'iCitem08'
                         , 'L_CNT4'    , 'iCitem09'
                         , 'L_CNT5'    , 'iCitem10'
                         , 'L_SEQNO1'  , 'iDispseq'
                         , 'S_STRING1' , 'sComcd1'
                         , 'S_STRING2' , 'sComcd2'
                         , 'S_STRING3' , 'sComcd3'
                         , 'S_STRING4' , 'sComcdnm1'
                         , 'S_STRING5' , 'sComcdnm2'
                         , 'S_STRING6' , 'sComcdnm3'
                         , 'S_STRING7' , 'sCitem01'
                         , 'S_STRING8' , 'sCitem02'
                         , 'S_STRING9' , 'sCitem03'
                         , 'S_STRING10', 'sCitem04'
                         , 'S_STRING11', 'sCitem05'
                         , 'S_STRING12', 'sCitem11'
                         , 'S_STRING13', 'sCitem12'
                         , 'S_STRING14', 'sDeldate'
                          ]) then

         if TpGetColumn.RowCount < 0 then
         begin
            ShowMessage(GetTxMsg);
            Exit;
         end
         else if TpGetColumn.RowCount = 0 then
         begin
            Exit;
         end;

         //-------------------------------------------------------------
         // 특정 기간 공지성 Inform인 경우, 공지기간 유효성 Check
         //-------------------------------------------------------------
         if (TpGetColumn.GetOutputDataS('sCitem03', 0) <> '') and
            (TpGetColumn.GetOutputDataS('sCitem04', 0) <> '') then
         begin

            if (FormatDateTime('yyyy-mm-dd hh:nn', Now) >= TpGetColumn.GetOutputDataS('sCitem03', 0)) and
               (FormatDateTime('yyyy-mm-dd hh:nn', Now) <= TpGetColumn.GetOutputDataS('sCitem04', 0)) then
               Result := TpGetColumn.GetOutputDataS('sComcdnm3', 0)
            else
               Result := '';
         end
         else
         begin
            //-------------------------------------------------------------
            // 특정 기간 공지성 Inform 이 아닌경우, Version 체크
            //-------------------------------------------------------------
            if (in_ComCd3 = 'VERSION') then
            begin
               // 프로그램 최신 Version 정보
               FsVersion := 'v' + TpGetColumn.GetOutputDataS('sComcdnm3', 0);

               //-------------------------------------------------------------
               // Version 체크후, 구 ver. 이면 공지 추가
               //-------------------------------------------------------------
               if (CopyByte(self.Caption, PosByte('[', self.Caption) + 1, 6) < 'ver' + TpGetColumn.GetOutputDataS('sComcdnm3', 0)) then
                  Result := '※ [' + TpGetColumn.GetOutputDataS('sComcdnm3', 0) + '] 다운로드 안내: [커뮤니티] -> 최신 Version을 사용해보세요.'
               else
                  Result := '';
            end
            else
               Result := TpGetColumn.GetOutputDataS('sComcdnm3', 0);

         end


   finally
      FreeAndNil(TpGetColumn);
      screen.Cursor  :=  crDefault;
   end;
end;




//------------------------------------------------------------------------------
// [조회] 각 근무처별 직원식 정보
//       - 최초 개발은 직원식단 Excel 업로드로 갔으나, 직원식 Table 조회로 전환.
//
// Author : Lee, Se-Ha
// Date   : 2013.09.16
//------------------------------------------------------------------------------
function TMainDlg.GetDietInfo(in_Gubun, in_DayGubun : String) : String;
var
   TpGetMenu : TTpSvc;
   sType1, sType2, sType3, sType4, sType5, sLocateNm, sDietTime : String;
begin
   // 리턴값 초기화
   GetDietInfo := '';



   Screen.Cursor := crHourglass;


   //----------------------------------------------------
   // 병원별 끼니별 직원식 조회 Var.
   //----------------------------------------------------
   sType1 := '15';
   sType2 := 'DIET';

   // 접속자 IP 식별후, 근무처 Assign
   if PosByte('안암도메인', FsUserIp) > 0 then
   begin
      sType3 := '안암식단테이블리모트DB';
      sType4 := '안암식단테이블리모트DB';
      sLocateNm := '안암';
   end
   else if PosByte('구로도메인', FsUserIp) > 0 then
   begin
      sType3 := '구로식단테이블리모트DB';
      sType4 := '구로식단테이블리모트DB';
      sLocateNm := '구로';
   end
   else if PosByte('안산도메인', FsUserIp) > 0 then
   begin
      sType3 := '안산식단테이블리모트DB';
      sType4 := '안산식단테이블리모트DB';
      sLocateNm := '안산';
   end
   // 원내 IP가 아닌경우(예: 노트북, 의생명 연구동 등), 종료
   else
   begin
      Exit;
   end;

   // 끼니 Flag
   if (FormatDateTime('hh:nn', Now) > '00:00') and (FormatDateTime('hh:nn', Now) <= '08:30') then
   begin
      sType5      := '1';
      sDietTime   := '아침';
   end
   else if (FormatDateTime('hh:nn', Now) > '08:30') and (FormatDateTime('hh:nn', Now) <= '14:00') then
   begin
      sType5      := '2';
      sDietTime   := '점심';
   end
   else if (FormatDateTime('hh:nn', Now) > '14:00') and (FormatDateTime('hh:nn', Now) <= '19:00') then
   begin
      sType5      := '3';
      sDietTime   := '저녁';
   end;



   //-----------------------------------------------------------------
   // 2-1. Load Data by TpSvc
   //-----------------------------------------------------------------
   TpGetMenu := TTpSvc.Create;
   TpGetMenu.Init(Self);


   Screen.Cursor := crHourglass;


   try
      TpGetMenu.CountField  := 'S_CODE6';
      TpGetMenu.ShowMsgFlag := False;

      if TpGetMenu.GetSvc('MD_KUMCM_S1',
                          ['S_TYPE1'  , sType1
                         , 'S_TYPE2'  , sType2
                         , 'S_TYPE3'  , sType3
                         , 'S_TYPE4'  , sType4
                         , 'S_TYPE5'  , sType5
                          ],
                          [
                           'S_CODE6'  , 'sDutyRmk'       // 끼니별 식이정보 (묶음)
                         { -- 병원별 끼니별 직원식 Table 자동조회 관련 주석 @ 2014.08.26 LSH
                         , 'S_STRING5', 'sAttachNm'      // 점심
                         , 'S_STRING6', 'sHideFile'      // 저녁
                         }
                         ]) then

         if TpGetMenu.RowCount < 0 then
         begin
            ShowMessage(GetTxMsg);
            Exit;
         end
         else if TpGetMenu.RowCount = 0 then
         begin
            Exit;
         end;



         //-------------------------------------------------------------
         // 현재 시간대별 맞춤 끼니 정보 Return
         //-------------------------------------------------------------
         { -- 병원별 끼니별 직원식 Table 자동조회 관련 주석 @ 2014.08.26 LSH
         if (FormatDateTime('hh:nn', Now) > '00:00') and (FormatDateTime('hh:nn', Now) <= '08:30') then
            Result := '▶ [' +  sType2 + '] 직원식 (조식) 메뉴 정보 ◀' + #13#10 + #13#10 + TpGetMenu.GetOutputDataS('sDutyRmk', 0)
         else if (FormatDateTime('hh:nn', Now) > '08:30') and (FormatDateTime('hh:nn', Now) <= '14:00') then
            Result := '▶ [' +  sType2 + '] 직원식 (중식) 메뉴 정보 ◀' + #13#10 + #13#10 + TpGetMenu.GetOutputDataS('sAttachNm', 0)
         else if (FormatDateTime('hh:nn', Now) > '14:00') and (FormatDateTime('hh:nn', Now) <= '19:00') then
            Result := '▶ [' +  sType2 + '] 직원식 (석식) 메뉴 정보 ◀' + #13#10 + #13#10 + TpGetMenu.GetOutputDataS('sHideFile', 0)
         else
            Result := '';
         }


         Result := '▶ [' +  sLocateNm + '] 직원식 (' + sDietTime + ') 메뉴 정보 ◀' + #13#10 + #13#10 + TpGetMenu.GetOutputDataS('sDutyRmk', 0);

   finally
      FreeAndNil(TpGetMenu);
      screen.Cursor  :=  crDefault;
   end;
end;





//------------------------------------------------------------------------------
// [Cell Align] TAdvStringGrid onGetAlignment Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2013.08.12
//------------------------------------------------------------------------------
procedure TMainDlg.asg_NwUpdate_GetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   // Horizontal Alignment
   HAlign := taCenter;

   // Vertical Alignment
   VAlign := vtaCenter;
end;




//------------------------------------------------------------------------------
// [Cell Align] TAdvStringGrid onGetAlignment Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2013.08.12
//------------------------------------------------------------------------------
procedure TMainDlg.asg_Network_GetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   // Horizontal Alignment
   if (ARow = 0) or ((ARow > 0) and ((ACol = C_NW_LOCATE) or
                                     (ACol = C_NW_CALLNO) or
                                     (ACol = C_NW_USERNM) or
                                     (ACol = C_NW_MOBILE))) then
      HAlign := taCenter;

   // Vertical Alignment
   VAlign := vtaCenter;
end;




//------------------------------------------------------------------------------
// AdvStringGrid onGetAlignment Event Handler
//
// Date   : 2013.08.09
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_Detail_GetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   // Horizontal Alignment
   if (ARow = 0) or ((ARow > 0) and ((ACol = C_D_LOCATE) or
                                     (ACol = C_D_DUTYUSER) or
                                     (ACol = C_D_CALLNO))
                                     ) then
      HAlign := taCenter;

   // Vertical Alignment
   VAlign := vtaCenter;
end;


//------------------------------------------------------------------------------
// AdvStringGrid onGetAlignment Event Handler
//
// Date   : 2013.08.09
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_Master_GetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   // Horizontal Alignment
   if (ARow = 0) or ((ARow > 0) and ((ACol = C_M_LOCATE) or
                                     (ACol = C_M_USERID) or
                                     (ACol = C_M_USERNM) or
                                     (ACol = C_M_MOBILE) or
                                     (ACol = C_M_USERIP))
                                     ) then
      HAlign := taCenter;

   // Vertical Alignment
   VAlign := vtaCenter;
end;



//------------------------------------------------------------------------------
// AdvStringGrid onGetEditorType Event Handler
//
// Date   : 2013.08.09
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_RegDoc_GetEditorType(Sender: TObject; ACol,
  ARow: Integer; var AEditor: TEditorType);
begin
   with asg_RegDoc do
   begin
      if (ARow > 0) and (ACol = C_RD_GUBUN) then
      begin
         AEditor := edComboList;

         ClearComboString;

         AddComboString('검색');
         AddComboString('등록');
         AddComboString('수정');
         AddComboString('삭제');
      end;

      if (ARow > 0) and (ACol = C_RD_DOCLIST) then
      begin
         AEditor := edComboList;

         ClearComboString;

         AddComboString('발송');
         AddComboString('협조');
         AddComboString('회보');
         AddComboString('계약처');
         AddComboString('OCS권한');
         AddComboString('CRA/CRC');
      end;

      if (ARow > 0) and (ACol = C_RD_REGDATE) then
         AEditor := edDateEdit;


   end;
end;


//------------------------------------------------------------------------------
// AdvStringGrid onEditingDone Event Handler
//
// Date   : 2013.08.12
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_RegDoc_EditingDone(Sender: TObject);
var
   tmpGubun : String;
begin
   with asg_RegDoc do
   begin
      // 버튼 종류 분기
      if (
            (Col = C_RD_GUBUN) or
            (Col = C_RD_DOCLIST)
         ) and (Cells[C_RD_GUBUN, Row] <> '검색') then
      begin
         tmpGubun := Cells[C_RD_GUBUN, Row];

         // 신규등록시, 가장 최근 해당 Category 이력 자동조회 @ 2015.05.11 LSH
         if (Cells[C_RD_GUBUN, Row] = '등록') then
         begin
            fsbt_ComDoc.Tag          := 9999;
            Cells[C_RD_GUBUN,   Row] := '검색';
            Cells[C_RD_DOCYEAR, Row] := FormatDateTime('yyyy', Date);

            AddButton(C_RD_BUTTON,  Row, ColWidths[C_RD_BUTTON]-5,  20, 'Search', haBeforeText, vaCenter);           // Search

            asg_RegDoc_ButtonClick(Sender, C_RD_BUTTON, Row);

            fsbt_ComDoc.Tag          := 0;
         end;

         Cells[C_RD_GUBUN,    Row] := tmpGubun;

         AddButton(C_RD_BUTTON,  Row, ColWidths[C_RD_BUTTON]-5,  20, 'Update', haBeforeText, vaCenter);           // Update

         // 상세하위 항목 초기화
         Cells[C_RD_DOCSEQ,   Row] := '';
         Cells[C_RD_DOCTITLE, Row] := '';
         Cells[C_RD_REGDATE,  Row] := '';
         Cells[C_RD_REGUSER,  Row] := '';
         Cells[C_RD_RELDEPT,  Row] := '';
         Cells[C_RD_DOCRMK,   Row] := '';
         Cells[C_RD_DOCLOC,   Row] := '';

      end
      else if (
                  (Col = C_RD_GUBUN) or
                  (Col = C_RD_DOCLIST)
               ) and (Cells[C_RD_GUBUN, Row] = '검색') then
      begin
         AddButton(C_RD_BUTTON,  Row, ColWidths[C_RD_BUTTON]-5,  20, 'Search', haBeforeText, vaCenter);           // Search

         // 상세하위 항목 초기화
         Cells[C_RD_DOCSEQ,   Row] := '';
         Cells[C_RD_DOCTITLE, Row] := '';
         Cells[C_RD_REGDATE,  Row] := '';
         Cells[C_RD_REGUSER,  Row] := '';
         Cells[C_RD_RELDEPT,  Row] := '';
         Cells[C_RD_DOCRMK,   Row] := '';
         Cells[C_RD_DOCLOC,   Row] := '';

         {  -- 원하는 순서대로 작동하지 않아서 일단 보류 ... @ 2015.05.12 LSH
         // 검색시 년도 정보 있으면, 자동 검색 Start @ 2015.05.11 LSH
         if (Col = C_RD_DOCYEAR) and (Cells[C_RD_DOCYEAR, Row] <> '') then
            asg_RegDoc_ButtonClick(Sender, C_RD_BUTTON, Row);
         }
      end;


   end;

end;


//------------------------------------------------------------------------------
// [문서관리 Update] TAdvStringGrid onButtonClick Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2013.08.12
//------------------------------------------------------------------------------
procedure TMainDlg.asg_RegDoc_ButtonClick(Sender: TObject; ACol,
  ARow: Integer);
var
   TpActDoc   : TTpSvc;
   TpGetDoc   : TTpSvc;
   i, iCnt    : Integer;
   vType1   ,
   vLocate  ,
   vDocList ,
   vDocYear ,
   vDocSeq  ,
   vDocTitle,
   vRegDate ,
   vRegUser ,
   vRelDept ,
   vDocRmk  ,
   vDocLoc  ,
   vDelDate ,
   vEditIp,
   vDeptCd,
   vSocNo,
   vMobile,
   vGrpId,
   vUseLvl,
   vTotRmk,
   vAARemark,
   vGRRemark,
   vASRemark,
   vHQRemark   : Variant;
   sType, sLocateNm, sDelDate : String;
   varResult : String;
begin
   if not IsLogonUser then
   begin
      MessageBox(self.Handle,
                 PChar('[비상연락망] - [담당자정보]에 담당자 IP 등록후 사용하실 수 있습니다.'),
                 PChar(Self.Caption + ' : 등록되지 않은 User 제한 알림'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;




   // 개인별 정보변경 Update
   if (ACol = C_RD_BUTTON) then
   begin

      // 요청구분별 Input variables Assign
      if (asg_RegDoc.Cells[C_RD_GUBUN, ARow] = '검색') then
      begin
         sType    := '6';
         sDelDate := '';
      end
      else if ((asg_RegDoc.Cells[C_RD_GUBUN, ARow] = '등록') or
               (asg_RegDoc.Cells[C_RD_GUBUN, ARow] = '수정')) then
      begin
         sType    := '3';
         sDelDate := '';
      end
      else if  (asg_RegDoc.Cells[C_RD_GUBUN, ARow] = '삭제') then
      begin
         sType    := '3';
         sDelDate := FormatDateTime('yyyy-mm-dd', Date);
      end;


      // 접속자 IP 식별후, 근무처 Assign
      if PosByte('안산도메인', FsUserIp) > 0 then
         sLocateNm := '안암'
      else if PosByte('구로도메인', FsUserIp) > 0 then
         sLocateNm := '구로'
      else if PosByte('안산도메인', FsUserIp) > 0 then
         sLocateNm := '안산';



      iCnt  := 0;


      //-------------------------------------------------------------------
      // 2. Create Variables
      //-------------------------------------------------------------------
      if (asg_RegDoc.Cells[C_RD_GUBUN, ARow] <> '검색') then
      begin
         with asg_RegDoc do
         begin

            if (Cells[C_RD_GUBUN, ARow] = '등록') and
               ((Cells[C_RD_DOCLIST, ARow] = '') or
                (Cells[C_RD_DOCYEAR, ARow] = '') or
                (Cells[C_RD_DOCSEQ,  ARow] = '') or
                (Cells[C_RD_DOCTITLE,ARow] = '') or
                (Cells[C_RD_REGDATE, ARow] = '') or
                (Cells[C_RD_REGUSER, ARow] = '')) then
            begin
               MessageBox(self.Handle,
                          PChar('문서 등록시 [종류/년도/번호/제목/등록일/등록자]는 필수입력 항목입니다.'),
                          PChar(Self.Caption + ' : 문서필수입력 항목 알림'),
                          MB_OK + MB_ICONWARNING);

               Exit;
            end;

            if ((Cells[C_RD_GUBUN, ARow] = '수정') or
                (Cells[C_RD_GUBUN, ARow] = '삭제')) and
               ((Cells[C_RD_DOCLIST, ARow] = '') or
                (Cells[C_RD_DOCYEAR, ARow] = '') or
                (Cells[C_RD_DOCSEQ,  ARow] = '')) then
            begin
               MessageBox(self.Handle,
                          PChar('문서 수정/삭제시 [종류/년도/번호]는 필수입력 항목입니다.'),
                          PChar(Self.Caption + ' : 문서필수체크 항목 알림'),
                          MB_OK + MB_ICONWARNING);

               Exit;
            end;



            // 계약처 등록(수정) 분기 적용 @ 2014.07.18 LSH
            if Cells[C_RD_DOCLIST, ARow] = '계약처' then
            begin
               for i := ARow to ARow do
               begin
                  //------------------------------------------------------------------
                  // 2-1. Append Variables
                  //------------------------------------------------------------------
                  AppendVariant(vType1   ,   sType                    );
                  AppendVariant(vLocate  ,   sLocateNm                );
                  AppendVariant(vDocList ,   Cells[C_RD_DOCLIST,    i]);
                  AppendVariant(vDocYear ,   Cells[C_RD_DOCYEAR,    i]);
                  AppendVariant(vDocSeq  ,   Cells[C_RD_DOCSEQ,     i]);
                  AppendVariant(vDocTitle,   Cells[C_RD_DOCTITLE,   i]);
                  AppendVariant(vRegDate ,   Cells[C_RD_REGDATE,    i]);
                  AppendVariant(vRegUser ,   Cells[C_RD_REGUSER,    i]);
                  AppendVariant(vRelDept ,   Cells[C_RD_RELDEPT,    i]);
                  AppendVariant(vDocRmk  ,   Cells[C_RD_DOCRMK,     i]);
                  AppendVariant(vDocLoc  ,   Cells[C_RD_DOCLOC,     i]);
                  AppendVariant(vEditIp  ,   FsUserIp                 );
                  AppendVariant(vDelDate ,   sDelDate                 );
                  AppendVariant(vTotRmk  ,   asg_DocSpec.Cells[1, R_DS_SOCNO]);
                  AppendVariant(vHQRemark,   asg_DocSpec.Cells[1, R_DS_STARTDT]);
                  AppendVariant(vAARemark,   asg_DocSpec.Cells[1, R_DS_ENDDT]);
                  AppendVariant(vGRRemark,   asg_DocSpec.Cells[1, R_DS_MOBILE]);
                  AppendVariant(vASRemark,   asg_DocSpec.Cells[1, R_DS_GRPID]);

                  Inc(iCnt);
               end;
            end
            else
            begin
               for i := ARow to ARow do
               begin
                  //------------------------------------------------------------------
                  // 2-1. Append Variables
                  //------------------------------------------------------------------
                  AppendVariant(vType1   ,   sType                    );
                  AppendVariant(vLocate  ,   sLocateNm                );
                  AppendVariant(vDocList ,   Cells[C_RD_DOCLIST,    i]);
                  AppendVariant(vDocYear ,   Cells[C_RD_DOCYEAR,    i]);
                  AppendVariant(vDocSeq  ,   Cells[C_RD_DOCSEQ,     i]);
                  AppendVariant(vDocTitle,   Cells[C_RD_DOCTITLE,   i]);
                  AppendVariant(vRegDate ,   Cells[C_RD_REGDATE,    i]);
                  AppendVariant(vRegUser ,   Cells[C_RD_REGUSER,    i]);
                  AppendVariant(vRelDept ,   Cells[C_RD_RELDEPT,    i]);
                  AppendVariant(vDocRmk  ,   Cells[C_RD_DOCRMK,     i]);
                  AppendVariant(vDocLoc  ,   Cells[C_RD_DOCLOC,     i]);
                  AppendVariant(vEditIp  ,   FsUserIp                 );
                  AppendVariant(vDeptCd  ,   asg_DocSpec.Cells[1, R_DS_DEPTCD]);
                  AppendVariant(vSocNo   ,   asg_DocSpec.Cells[1, R_DS_SOCNO] );
                  AppendVariant(vRegDate ,   asg_DocSpec.Cells[1, R_DS_STARTDT]);
                  AppendVariant(vDelDate ,   sDelDate {asg_DocSpec.Cells[1, R_DS_ENDDT]} );
                  AppendVariant(vMobile  ,   asg_DocSpec.Cells[1, R_DS_MOBILE]);
                  AppendVariant(vGrpId   ,   asg_DocSpec.Cells[1, R_DS_GRPID] );
                  AppendVariant(vUseLvl  ,   asg_DocSpec.Cells[1, R_DS_USELVL]);

                  Inc(iCnt);
               end;
            end;
         end;



         if iCnt <= 0 then
            Exit;

         //-------------------------------------------------------------------
         // 3. Insert by TpSvc
         //-------------------------------------------------------------------
         TpActDoc := TTpSvc.Create;
         TpActDoc.Init(Self);


         Screen.Cursor := crHourGlass;


         try
            if (asg_RegDoc.Cells[C_RD_DOCLIST, ARow] = '계약처') then
            begin
               if TpActDoc.PutSvc('MD_KUMCM_M1',
                                   [
                                    'S_TYPE1'   , vType1
                                  , 'S_STRING1' , vLocate
                                  , 'S_STRING8' , vEditIp
                                  , 'S_STRING10', vDelDate
                                  , 'S_STRING16', vDocList
                                  , 'S_STRING17', vDocYear
                                  , 'S_STRING18', vDocSeq
                                  , 'S_STRING19', vDocTitle
                                  , 'S_STRING20', vRegDate
                                  , 'S_STRING21', vRegUser
                                  , 'S_STRING24', vDocLoc
                                  , 'S_STRING28', vRelDept
                                  , 'S_STRING29', vTotRmk
                                  , 'S_STRING30', vHQRemark
                                  , 'S_STRING37', vAARemark
                                  , 'S_STRING38', vGRRemark
                                  , 'S_STRING46', vASRemark
                                  , 'S_STRING47', vDocRmk
                                  ] ) then
               begin
                  MessageBox(self.Handle,
                             PChar('[' + asg_RegDoc.Cells[C_RD_DOCLIST, asg_RegDoc.Row] + '] ' +
                                   asg_RegDoc.Cells[C_RD_DOCYEAR, asg_RegDoc.Row] + '-' +
                                   asg_RegDoc.Cells[C_RD_DOCSEQ,  asg_RegDoc.Row] +
                                   ' 문서가 정상적으로 [' + asg_RegDoc.Cells[C_RD_GUBUN,  asg_RegDoc.Row] + '] 되었습니다.'),
                             '[KUMC 다이얼로그] 문서관리 업데이트 알림 ',
                             MB_OK + MB_ICONINFORMATION);

                  // Refresh
                  SelGridInfo('DOCREFRESH');
               end
               else
               begin
                  ShowMessageM(GetTxMsg);
               end;
            end
            else
            begin
               if TpActDoc.PutSvc('MD_KUMCM_M1',
                                   [
                                    'S_TYPE1'   , vType1
                                  , 'S_STRING1' , vLocate
                                  , 'S_STRING5' , vMobile
                                  , 'S_STRING8' , vEditIp
                                  , 'S_STRING9' , vDeptCd
                                  , 'S_STRING10', vDelDate
                                  , 'S_STRING16', vDocList
                                  , 'S_STRING17', vDocYear
                                  , 'S_STRING18', vDocSeq
                                  , 'S_STRING19', vDocTitle
                                  , 'S_STRING20', vRegDate
                                  , 'S_STRING21', vRegUser
                                  , 'S_STRING22', vRelDept
                                  , 'S_STRING23', vDocRmk
                                  , 'S_STRING24', vDocLoc
                                  , 'S_STRING32', vUseLvl
                                  , 'S_STRING33', vGrpId
                                  , 'S_STRING44', vSocNo
                                  ] ) then
               begin
                  MessageBox(self.Handle,
                             PChar('[' + asg_RegDoc.Cells[C_RD_DOCLIST, asg_RegDoc.Row] + '] ' +
                                   asg_RegDoc.Cells[C_RD_DOCYEAR, asg_RegDoc.Row] + '-' +
                                   asg_RegDoc.Cells[C_RD_DOCSEQ,  asg_RegDoc.Row] +
                                   ' 문서가 정상적으로 [' + asg_RegDoc.Cells[C_RD_GUBUN,  asg_RegDoc.Row] + '] 되었습니다.'),
                             '[KUMC 다이얼로그] 문서관리 업데이트 알림 ',
                             MB_OK + MB_ICONINFORMATION);

                  // Refresh
                  SelGridInfo('DOCREFRESH');
               end
               else
               begin
                  ShowMessageM(GetTxMsg);
               end;
            end;


         finally
            FreeAndNil(TpActDoc);
            Screen.Cursor  := crDefault;
         end;
      end
      else if (asg_RegDoc.Cells[C_RD_GUBUN, ARow] = '검색') then
      begin

         // 등록 또는 수정시 해당 Category 임시조회 Tag (9999)시에는 Log 기록 제외 @ 2015.05.11 LSH
         if (fsbt_ComDoc.Tag <> 9999) and
            (
               (asg_RegDoc.Cells[C_RD_DOCLIST, ARow] = '') or
               (asg_RegDoc.Cells[C_RD_DOCYEAR, ARow] = '')
             ) then
         begin
            MessageBox(self.Handle,
                       PChar('문서 검색시 [종류/년도]는 필수입력 항목입니다.'),
                       PChar(Self.Caption + ' : 문서검색 필수입력 항목 알림'),
                       MB_OK + MB_ICONWARNING);

            Exit;
         end;



         asg_SelDoc.ClearRows(1, asg_SelDoc.RowCount);
         asg_SelDoc.RowCount := 2;



         //-----------------------------------------------------------------
         // 1. 입력 변수 Assign
         //-----------------------------------------------------------------
         vType1      := '8';
         vLocate     := sLocateNm;
         vDocList    := Trim(asg_RegDoc.Cells[C_RD_DOCLIST,  ARow]);
         vDocYear    := Trim(asg_RegDoc.Cells[C_RD_DOCYEAR,  ARow]);
         vDocTitle   := Trim(asg_RegDoc.Cells[C_RD_DOCTITLE, ARow]);
         vRegUser    := Trim(asg_RegDoc.Cells[C_RD_REGUSER,  ARow]);
         vRelDept    := Trim(asg_RegDoc.Cells[C_RD_RELDEPT,  ARow]);



         //-----------------------------------------------------------------
         // 2-1. Load Data by TpSvc
         //-----------------------------------------------------------------
         TpGetDoc := TTpSvc.Create;
         TpGetDoc.Init(Self);


         Screen.Cursor := crHourglass;


         try
            TpGetDoc.CountField  := 'S_CODE18';
            TpGetDoc.ShowMsgFlag := False;

            if TpGetDoc.GetSvc(  'MD_KUMCM_S1',
                                ['S_TYPE1'  , vType1
                               , 'S_TYPE2'  , vLocate
                               , 'S_TYPE3'  , vDocList
                               , 'S_TYPE4'  , vDocYear
                               , 'S_TYPE5'  , vDocTitle
                               , 'S_TYPE6'  , vRegUser
                               , 'S_TYPE7'  , vRelDept
                                ],
                                [
                                 'S_CODE6'  , 'sDutyRmk'
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
                               , 'S_STRING5', 'sAttachNm'
                               , 'S_STRING6', 'sHideFile'
                               , 'S_STRING16','sDeptNm'
                               , 'S_STRING17','sDeptSpec'

                                ]) then

               if TpGetDoc.RowCount < 0 then
               begin
                  ShowMessage(GetTxMsg);
                  Exit;
               end
               else if TpGetDoc.RowCount = 0 then
               begin
                  lb_SelDoc.Caption := '▶ 해당 회계년도에 대한 검색자료가 없습니다.';
                  Exit;
               end;



            with asg_SelDoc do
            begin
               iCnt  := TpGetDoc.RowCount;
               RowCount := iCnt + FixedRows + 1;


               for i := 0 to iCnt - 1 do
               begin
                  Cells[C_SD_DOCLIST, i+FixedRows] := TpGetDoc.GetOutputDataS('sDocList',  i);
                  Cells[C_SD_DOCSEQ , i+FixedRows] := TpGetDoc.GetOutputDataS('sDocSeq',   i);
                  Cells[C_SD_DOCTITLE,i+FixedRows] := TpGetDoc.GetOutputDataS('sDocTitle', i);
                  Cells[C_SD_REGDATE, i+FixedRows] := TpGetDoc.GetOutputDataS('sRegDate',  i);
                  Cells[C_SD_REGUSER, i+FixedRows] := TpGetDoc.GetOutputDataS('sRegUser',  i);
                  Cells[C_SD_RELDEPT, i+FixedRows] := TpGetDoc.GetOutputDataS('sRelDept',  i);
                  Cells[C_SD_DOCRMK,  i+FixedRows] := TpGetDoc.GetOutputDataS('sDocRmk',   i);
                  Cells[C_SD_DOCLOC,  i+FixedRows] := TpGetDoc.GetOutputDataS('sDocLoc',   i);
                  Cells[C_SD_HQREMARK,i+FixedRows] := TpGetDoc.GetOutputDataS('sAttachNm', i);
                  Cells[C_SD_AAREMARK,i+FixedRows] := TpGetDoc.GetOutputDataS('sHideFile', i);
                  Cells[C_SD_GRREMARK,i+FixedRows] := TpGetDoc.GetOutputDataS('sDeptNm',   i);
                  Cells[C_SD_ASREMARK,i+FixedRows] := TpGetDoc.GetOutputDataS('sDeptSpec', i);
                  Cells[C_SD_TOTRMK,  i+FixedRows] := TpGetDoc.GetOutputDataS('sDutyRmk',  i);

                  // 계약만료된 항목은 FontColor = clRed 업데이트 @ 2014.06.02 LSH
                  if (Cells[C_SD_DOCLIST, i+FixedRows] = '계약처') then
                  begin
                     if LengthByte(Trim(CopyByte(Cells[C_SD_DOCRMK,  i+FixedRows], 19, 2))) > 1 then
                     begin
                        if ((CopyByte(Cells[C_SD_DOCRMK,  i+FixedRows], 14, 4) +
                             CopyByte(Cells[C_SD_DOCRMK,  i+FixedRows], 19, 2) +
                             CopyByte(Cells[C_SD_DOCRMK,  i+FixedRows], 23, 4)) < FormatDateTime('yyyymmdd', Date)) then
                        begin
                           FontColors[C_SD_DOCTITLE,  i+FixedRows] := clRed;
                           FontColors[C_SD_DOCRMK,    i+FixedRows] := clRed;
                        end;
                     end
                     else
                     begin
                        if ((CopyByte(Cells[C_SD_DOCRMK,  i+FixedRows], 14, 4) + '0' +
                             Trim(CopyByte(Cells[C_SD_DOCRMK,  i+FixedRows], 19, 2)) +
                             CopyByte(Cells[C_SD_DOCRMK,  i+FixedRows], 23, 4)) < FormatDateTime('yyyymmdd', Date)) then
                        begin
                           FontColors[C_SD_DOCTITLE,  i+FixedRows] := clRed;
                           FontColors[C_SD_DOCRMK,    i+FixedRows] := clRed;
                        end;
                     end;
                  end;

                  AutoSizeRow(i+FixedRows);
               end;
            end;

            // RowCount 정리
            asg_SelDoc.RowCount := asg_SelDoc.RowCount - 1;

            // Comments
            lb_SelDoc.Caption := '▶ ' + IntToStr(iCnt) + '건의 ' + '[' + asg_RegDoc.Cells[C_RD_DOCYEAR, asg_RegDoc.Row] + ' ' +
                                                                          asg_RegDoc.Cells[C_RD_DOCLIST, asg_RegDoc.Row] + '] ' +
                                                                    '내역을 검색하였습니다.';

         finally
            FreeAndNil(TpGetDoc);
            Screen.Cursor := crDefault;
         end;


         // 등록 또는 수정시 해당 Category 임시조회 Tag (9999)시에는 Log 기록 제외 @ 2015.05.11 LSH
         if fsbt_ComDoc.Tag <> 9999 then
         begin
            // 로그 Update
            UpdateLog(asg_RegDoc.Cells[C_RD_DOCLIST,  asg_RegDoc.Row],
                      asg_RegDoc.Cells[C_RD_DOCYEAR,  asg_RegDoc.Row] + asg_RegDoc.Cells[C_RD_DOCSEQ,   asg_RegDoc.Row],
                      FsUserIp,
                      'S',
                      asg_RegDoc.Cells[C_RD_DOCTITLE, asg_RegDoc.Row],
                      asg_RegDoc.Cells[C_RD_RELDEPT,  asg_RegDoc.Row],
                      varResult
                      );
         end;
      end;
   end;
end;




//------------------------------------------------------------------------------
// AdvStringGrid onDblClick Event Handler
//
// Date   : 2013.08.12
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_RegDoc_DblClickCell(Sender: TObject; ARow,
  ACol: Integer);
begin

end;





//------------------------------------------------------------------------------
// [조회] 각 문서별 Max Seqno 정보
//
// Author : Lee, Se-Ha
// Date   : 2013.08.13
//------------------------------------------------------------------------------
function TMainDlg.GetMaxDocSeq(in_DocList, in_DocYear : String) : String;
var
   TpGetDocSeq : TTpSvc;
   sType1, sDocList, sDocYear, sLocate : String;
begin

   //-----------------------------------------------------------------
   // 1. 입력 변수 Assign
   //-----------------------------------------------------------------
   sType1   := '7';
   sDocList := in_DocList;
   sDocYear := in_DocYear;

   // CRA/CRC 인사 및 프로그램 자동입력 개발완료이나, 운영 D/B Transaction 에러(ora-24777) 안되어 일단 주석처리 @ 2013.12.06 LSH
   // non-XA 서버 이관작업후, 하기 프로세스 오픈 @ 2014.01.10 LSH
   // ID 상세정보 입력 Panel on시, max ID 채번위해 조회조건 분기
   if (
         (PosByte('CRA', sDocList) > 0) or
         (PosByte('CRC', sDocList) > 0)
      ) then

      sType1 := '26';



   // 접속자 IP 식별후, 근무처 Assign
   if PosByte('안암도메인', FsUserIp) > 0 then
      sLocate := '안암'
   else if PosByte('구로도메인', FsUserIp) > 0 then
      sLocate := '구로'
   else if PosByte('안산도메인', FsUserIp) > 0 then
      sLocate := '안산';




   //-----------------------------------------------------------------
   // 2. 조회
   //-----------------------------------------------------------------
   Screen.Cursor := crHourGlass;





   //-----------------------------------------------------------------
   // 2-1. Load Data by TpSvc
   //-----------------------------------------------------------------
   TpGetDocSeq := TTpSvc.Create;
   TpGetDocSeq.Init(Self);




   Screen.Cursor := crHourglass;





   try
      TpGetDocSeq.CountField  := 'S_CODE20';
      TpGetDocSeq.ShowMsgFlag := False;


      if TpGetDocSeq.GetSvc('MD_KUMCM_S1',
                          ['S_TYPE1'  , sType1
                         , 'S_TYPE2'  , sDocList
                         , 'S_TYPE3'  , sDocYear
                         , 'S_TYPE4'  , sLocate
                          ],
                          [
                           'S_CODE20' , 'sDocSeq'
                         , 'S_CODE25' , 'sDocRmk'
                          ]) then

         if TpGetDocSeq.RowCount < 0 then
         begin
            ShowMessage(GetTxMsg);
            Exit;
         end
         else if TpGetDocSeq.RowCount = 0 then
         begin
            Exit;
         end;

         //----------------------------------------------------
         // ID 상세정보 입력 Panel on시, max ID 채번값 assign
         //----------------------------------------------------
         if (sType1 = '26') then
         begin
            asg_RegDoc.Cells[C_RD_DOCSEQ, asg_RegDoc.Row] := TpGetDocSeq.GetOutputDataS('sDocSeq', 0);
            Result := TpGetDocSeq.GetOutputDataS('sDocRmk', 0);
         end
         else if (sType1 = '7') then
            //----------------------------------------------------
            // 최신 Max 문서번호 Return
            //----------------------------------------------------
            Result := TpGetDocSeq.GetOutputDataS('sDocSeq', 0);





   finally
      FreeAndNil(TpGetDocSeq);
      Screen.Cursor := crDefault;
   end;
end;



//------------------------------------------------------------------------------
// [조회] IP 별 사용자 상세 정보
//
// Author : Lee, Se-Ha
// Date   : 2013.08.13
//------------------------------------------------------------------------------
function TMainDlg.CheckUserInfo(in_UserIp : String; var varUserNm,
                                                        varNickNm,
                                                        varUserPart,
                                                        varUserSpec,
                                                        varUserMobile,
                                                        varUserCallNo,
                                                        varMngrNm : String) : Boolean;
var
   TpGetUser : TTpSvc;
   GetUser   : TSelUser;
   i         : Integer;
begin
   Result := False;


   varUserNm      := '';
   varNickNm      := '';
   varUserPart    := '';
   varUserSpec    := '';
   varUserMobile  := '';
   varUserCallNo  := '';
   varMngrNm      := '';   // 문서 관리자 (팀장) @ 2014.07.18 LSH



   Screen.Cursor := crHourglass;


   //-----------------------------------------------------------------
   // 2-1. Load Data by TpSvc
   //-----------------------------------------------------------------
   TpGetUser := TTpSvc.Create;
   TpGetUser.Init(nil);



   Screen.Cursor := crHourglass;



   try
      TpGetUser.CountField  := 'S_CODE3';
      TpGetUser.ShowMsgFlag := False;


      if TpGetUser.GetSvc('MD_KUMCM_S1',
                          ['S_TYPE1', '9'
                         , 'S_TYPE2', in_UserIp
                          ],
                          [
                          'S_CODE3',   'sUserNm'
                         ,'S_CODE4',   'sDutyPart'
                         ,'S_CODE5',   'sDutySpec'
                         ,'S_CODE7',   'sMobile'
                         ,'S_CODE10' , 'sDutyUser'
                         ,'S_CODE12',  'sCallNo'
                         ,'S_STRING14','sNickNm'
                          ]) then
      begin
         if TpGetUser.RowCount = 0 then
         begin
            Exit;
         end;
      end
      else
      begin

         if TpGetUser.RowCount < 0 then
         begin
            ShowMessage(GetTxMsg);
            Exit;
         end
      end;


      //------------------------------------------------------------
      // 단일 IP + 다중사용자 인 경우, 사용자 선택화면 Call
      //------------------------------------------------------------
      if TpGetUser.RowCount > 1 then
      begin

         GetUser := TSelUser.Create(self);

         try
            // 다중 User 정보 표기
            GetUser.lb_UserCnt.Caption := IntToStr(TpGetUser.RowCount) + '명의 사용자가 존재합니다.';


            for i := 0 to TpGetUser.RowCount - 1 do
            begin
               GetUser.asg_UserList.Cells[0, i+1] := TpGetUser.GetOutputDataS('sUserNm',   i);
               GetUser.asg_UserList.Cells[1, i+1] := FsUserIp;
               GetUser.asg_UserList.Cells[2, i+1] := TpGetUser.GetOutputDataS('sNickNm',   i);
               GetUser.asg_UserList.Cells[3, i+1] := TpGetUser.GetOutputDataS('sDutyPart', i);
               GetUser.asg_UserList.Cells[4, i+1] := TpGetUser.GetOutputDataS('sDutySpec', i);
               GetUser.asg_UserList.Cells[5, i+1] := TpGetUser.GetOutputDataS('sMobile',   i);
               GetUser.asg_UserList.Cells[6, i+1] := TpGetUser.GetOutputDataS('sCallNo',   i);
               GetUser.asg_UserList.Cells[7, i+1] := TpGetUser.GetOutputDataS('sDutyUser', i);
            end;


            GetUser.ShowModal;


         finally
            GetUser.Free;
         end;
      end
      else
      begin
         varUserNm      := TpGetUser.GetOutputDataS('sUserNm',    0);
         varNickNm      := TpGetUser.GetOutputDataS('sNickNm',    0);
         varUserPart    := TpGetUser.GetOutputDataS('sDutyPart',  0);
         varUserSpec    := TpGetUser.GetOutputDataS('sDutySpec',  0);
         varUserMobile  := TpGetUser.GetOutputDataS('sMobile',    0);
         varUserCallNo  := TpGetUser.GetOutputDataS('sCallNo',    0);
         varMngrNm      := TpGetUser.GetOutputDataS('sDutyUser',  0);   // 문서 관리자 (팀장) @ 2014.07.18 LSH
      end;


      Result := True;


   finally
      FreeAndNil(TpGetUser);
      screen.Cursor  :=  crDefault;
   end;
end;



//------------------------------------------------------------------------------
// [조회] 요일 정보 가져오는 함수
//       - 소스출처 : MComFunc.pas
//
// Author : Lee, Se-Ha
// Date   : 2013.08.14
//------------------------------------------------------------------------------
function TMainDlg.GetDayofWeek(Adate : TDatetime; Type1 : String): String;
var
  days: array[1..7] of string;
begin
   if Type1 = 'EF' then
   begin
      days[1] := 'Sunday';
      days[2] := 'Monday';
      days[3] := 'Tuesday';
      days[4] := 'Wednesday';
      days[5] := 'Thursday';
      days[6] := 'Friday';
      days[7] := 'Saturday';
   end
   else if Type1 = 'ES' then
   begin
      days[1] := 'SUN';
      days[2] := 'MON';
      days[3] := 'TUE';
      days[4] := 'WED';
      days[5] := 'THU';
      days[6] := 'FRI';
      days[7] := 'SAT';
   end
   else if Type1 = 'HF' then
   begin
      days[1] := '일요일';
      days[2] := '월요일';
      days[3] := '화요일';
      days[4] := '수요일';
      days[5] := '목요일';
      days[6] := '금요일';
      days[7] := '토요일';
   end
   else if Type1 = 'HS' then
   begin
      days[1] := '일';
      days[2] := '월';
      days[3] := '화';
      days[4] := '수';
      days[5] := '목';
      days[6] := '금';
      days[7] := '토';
   end;

   Result := days[DayOfWeek(ADate)];
end;



//------------------------------------------------------------------------------
// Form onDestroy Event Handler
//
// Date   : 2013.08.08
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.FormDestroy(Sender: TObject);
begin
   ServerSocket.Active := False;
   MainDlg := Nil;
end;


//------------------------------------------------------------------------------
// AdvStringGrid onClickCell Event Handler
//
// Date   : 2013.08.12
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_RegDoc_ClickCell(Sender: TObject; ARow,
  ACol: Integer);
begin
   with asg_RegDoc do
   begin
      if (ARow > 0) then
      begin
         if (ACol = C_RD_DOCYEAR) then
         begin
            if (FormatDateTime('mm', Date) = '01') or
               (FormatDateTime('mm', Date) = '02') then
               Cells[C_RD_DOCYEAR, ARow] := IntToStr(StrToInt(FormatDateTime('yyyy', Date)) - 1)
            else
               Cells[C_RD_DOCYEAR, ARow] := FormatDateTime('yyyy', Date);

            lb_RegDoc.Caption := '▶ 2010년 문서부터 D/B로 관리중입니다. 4자리 회계년도를 입력해주세요.';
         end
         else if (ACol = C_RD_GUBUN) then
            lb_RegDoc.Caption := '▶ 부서공통문서(발송/협조 등)를 검색 또는 등록(수정)하실 수 있습니다.'
         else if (ACol = C_RD_DOCSEQ) then
            lb_RegDoc.Caption := '▶ [클릭]하시면 자동으로 문서번호를 채번합니다. (수정/삭제시 문서번호가 필요)'
         else if (ACol = C_RD_DOCTITLE) and (Cells[C_RD_GUBUN, ARow] = '검색') then
            lb_RegDoc.Caption := '▶ 검색시 문서에 포함된 단어(예: 네트워크, 협조 등)로 빠른검색이 가능합니다.'
         else if (ACol = C_RD_REGUSER) then
         begin
            if Cells[C_RD_GUBUN, ARow] = '등록' then
               Cells[C_RD_REGUSER, ARow] := FsUserNm;

            lb_RegDoc.Caption := '▶ 발송의 경우 문서 기안자, 협조의 경우 접수자(또는 업무관련자)를 적어주세요.'
         end
         else if (ACol = C_RD_RELDEPT) then
            lb_RegDoc.Caption := '▶ 해당 문서를 받는 부서명(또는 유관부서)을, 또는 기안한 부서명을 적어주세요.'
         else if (ACol = C_RD_DOCRMK) then
            lb_RegDoc.Caption := '▶ 해당 문서에 관련된 Tag 또는 근무처/담당자/연락처 등 특이사항을 적어주세요.'
         else if (ACol = C_RD_DOCLOC) then
            lb_RegDoc.Caption := '▶ 마지막 등록문서의 보관위치를 default로 적어주세요.'
         else
            lb_RegDoc.Caption := '';




         // 신규 [등록]의 경우, 문서번호 최신 Seqno 가져오기
         if (ACol = C_RD_DOCSEQ) and
            (Cells[C_RD_GUBUN, ARow] = '등록') then
         begin
            Cells[ACol, ARow] := GetMaxDocSeq(Cells[C_RD_DOCLIST, ARow],
                                              Cells[C_RD_DOCYEAR, ARow]);

         end;


         // ora-24777 : use of non-migratable data base link not allowed ] 오류 (XA와 DB link 동시사용 안됨)로
         // 개발완료후, 해당 서비스(md_kumcm_m1) non-XA 서버 이관작업후, 오픈 @ 2014.01.10 LSH
         if (ACol = C_RD_DOCSEQ) or
            (ACol = C_RD_DOCYEAR) then
         begin
            // 임상연구자(CRA/CRC) ID 등록(또는 수정)시, 상세정보 입력 Panel Open
            if (
                  (Cells[C_RD_GUBUN, ARow] = '등록') or
                  (Cells[C_RD_GUBUN, ARow] = '수정')
               ) and
               (
                  (PosByte('CRA', Cells[C_RD_DOCLIST, ARow]) > 0) or
                  (PosByte('CRC', Cells[C_RD_DOCLIST, ARow]) > 0)
               ) then
            begin
               asg_DocSpec.Cells[0, R_DS_USERID]   := 'ID';
               asg_DocSpec.Cells[0, R_DS_USERNM]   := '성명';
               asg_DocSpec.Cells[0, R_DS_DEPTCD]   := '소속코드';
               asg_DocSpec.Cells[0, R_DS_DEPTNM]   := '소속명';
               asg_DocSpec.Cells[0, R_DS_SOCNO]    := '주민등록번호';
               asg_DocSpec.Cells[0, R_DS_STARTDT]  := '시작일자';
               asg_DocSpec.Cells[0, R_DS_ENDDT]    := '종료일자';
               asg_DocSpec.Cells[0, R_DS_MOBILE]   := '휴대폰';
               asg_DocSpec.Cells[0, R_DS_GRPID]    := '프로그램ID';
               asg_DocSpec.Cells[0, R_DS_USELVL]   := '프로그램레벨';


               asg_DocSpec.Cells[1, R_DS_USERID]   := GetMaxDocSeq(Cells[C_RD_DOCLIST, ARow],
                                                                   Cells[C_RD_DOCYEAR, ARow]);
               asg_DocSpec.Cells[1, R_DS_USERNM]   := '';
               asg_DocSpec.Cells[1, R_DS_DEPTCD]   := 'CTC';
               asg_DocSpec.Cells[1, R_DS_DEPTNM]   := '';
               asg_DocSpec.Cells[1, R_DS_SOCNO]    := '';
               asg_DocSpec.Cells[1, R_DS_STARTDT]  := FormatDateTime('yyyy-mm-dd', Date);
               asg_DocSpec.Cells[1, R_DS_ENDDT]    := '';
               asg_DocSpec.Cells[1, R_DS_MOBILE]   := '';
               asg_DocSpec.Cells[1, R_DS_GRPID]    := 'MD1';
               asg_DocSpec.Cells[1, R_DS_USELVL]   := '3';

               apn_DocSpec.Top     := 155;
               apn_DocSpec.Left    := 140;
               apn_DocSpec.Collaps := True;
               apn_DocSpec.Visible := True;
               apn_DocSpec.Collaps := False;

               apn_DocSpec.Caption.Text := Cells[C_RD_DOCLIST, ARow] + ' ID 상세정보 등록(수정)';
            end
            // 계약처 상세정보 입력 Panel Open
            else if
                     (ACol = C_RD_DOCSEQ) and
                     (
                        (Cells[C_RD_GUBUN, ARow] = '등록') or
                        (Cells[C_RD_GUBUN, ARow] = '수정')
                     )  and
                     (
                        (PosByte('계약처', Cells[C_RD_DOCLIST, ARow]) > 0)
                     ) then
            begin
               asg_DocSpec.Cells[0, R_DS_USERID]   := '계약명';
               asg_DocSpec.Cells[0, R_DS_USERNM]   := '계약업체명';
               asg_DocSpec.Cells[0, R_DS_DEPTCD]   := '계약기간';
               asg_DocSpec.Cells[0, R_DS_DEPTNM]   := '';
               asg_DocSpec.Cells[0, R_DS_SOCNO]    := '분담액(총액)';
               asg_DocSpec.Cells[0, R_DS_STARTDT]  := '분담액(HQ)';
               asg_DocSpec.Cells[0, R_DS_ENDDT]    := '분담액(AA)';
               asg_DocSpec.Cells[0, R_DS_MOBILE]   := '분담액(GR)';
               asg_DocSpec.Cells[0, R_DS_GRPID]    := '분담액(AS)';
               asg_DocSpec.Cells[0, R_DS_USELVL]   := '문서위치';


               asg_DocSpec.Cells[1, R_DS_USERID]   := '예) 서버보안 S/W 유지보수 지급';
               asg_DocSpec.Cells[1, R_DS_USERNM]   := '예) (주)시큐브';
               asg_DocSpec.Cells[1, R_DS_DEPTCD]   := '예) 2014. 3. 1 ~ 2015. 2. 28 (12개월)';
               asg_DocSpec.Cells[1, R_DS_DEPTNM]   := '';
               asg_DocSpec.Cells[1, R_DS_SOCNO]    := '예) 375,000원';
               asg_DocSpec.Cells[1, R_DS_STARTDT]  := '예) 0원';
               asg_DocSpec.Cells[1, R_DS_ENDDT]    := '예) 146,300원';
               asg_DocSpec.Cells[1, R_DS_MOBILE]   := '예) 142,500원';
               asg_DocSpec.Cells[1, R_DS_GRPID]    := '예) 86,200원';
               asg_DocSpec.Cells[1, R_DS_USELVL]   := '예) PDF, 계약처';

               apn_DocSpec.Top     := 155;
               apn_DocSpec.Left    := 140;
               apn_DocSpec.Collaps := True;
               apn_DocSpec.Visible := True;
               apn_DocSpec.Collaps := False;

               apn_DocSpec.Caption.Text := Cells[C_RD_DOCLIST, ARow] + ' 유지보수 상세정보 등록(수정)';
            end;
         end;
      end;
   end;
end;


//------------------------------------------------------------------------------
// AdvStringGrid onDblClick Event Handler
//
// Date   : 2013.08.12
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_Board_DblClick(Sender: TObject);
var
   i, j : Integer;
   varResult : String;
   sServerIp, sRemoteFile, sLocalFile, sFtpUserID, sFtpPasswd, sFtpHostName, sFtpDir, varUpResult : String;
begin

   if (IsLogonUser) then
   begin
      if (asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row]   <> '┗') and
         (asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row+1] <> '┗') and
         (asg_Board.Cells[C_B_CLOSEYN,  asg_Board.Row] = 'C') then
      begin
         asg_Board.InsertRows(asg_Board.Row + 1, 1 + StrToInt(asg_Board.Cells[C_B_REPLY, asg_Board.Row]));
         asg_Board.Cells[C_B_CLOSEYN, asg_Board.Row]     := 'O';


         // 첨부파일 유무에 따른, 상세 Row Merege 조건 분기
         if (asg_Board.Cells[C_B_HIDEFILE, asg_Board.Row] = '') then
            asg_Board.MergeCells(C_B_CATEUP,  asg_Board.Row + 1, 4, 1)
         else
         begin
            asg_Board.MergeCells(C_B_CATEUP,  asg_Board.Row + 1, 2, 1);
            asg_Board.MergeCells(C_B_REGDATE, asg_Board.Row + 1, 2, 1);
         end;

         // 첨부 이미지 FTP 파일내역이 존재하면, 이미지 Download
         if (asg_Board.Cells[C_B_HIDEFILE, asg_Board.Row] <> '') then
         begin
            // 파일 업/다운로드를 위한 정보 조회
            if not GetBinUploadInfo(sServerIp, sFtpUserID, sFtpPasswd, sFtpHostName, sFtpDir) then
            begin
               MessageDlg('파일 저장을 위한 담당자 정보 조회중, 오류가 발생했습니다.', Dialogs.mtError, [Dialogs.mbOK], 0);
               exit;
            end;


            // 실제저장된 서버 IP
            sServerIp := C_KDIAL_FTP_IP;


            // 실제 서버에 저장되어 있는 파일명 지정
            if PosByte('/ftpspool/KDIALFILE/',asg_Board.Cells[C_B_HIDEFILE, asg_Board.Row]) > 0 then
               sRemoteFile := asg_Board.Cells[C_B_HIDEFILE, asg_Board.Row]
            else
               sRemoteFile := '/ftpspool/KDIALFILE/' + asg_Board.Cells[C_B_HIDEFILE, asg_Board.Row];

            // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
            sLocalFile  := G_HOMEDIR + 'TEMP\SPOOL\' + asg_Board.Cells[C_B_ATTACH, asg_Board.Row];



            if (GetBINFTP(sServerIp,
                          sFtpUserID,
                          sFtpPasswd,
                          sRemoteFile,
                          sLocalFile,
                          False)) then
            begin
               //	정상적인 FTP 다운로드
            end;



            // Local에 해당 Image 존재유무 체크 (Local file의 Size 체크 추가 : Local에 해당 사번.jpg가 0 size로 존재하면 JPEG Error #41/#42 떨어짐) @ 2019.12.03 LSH
            if (FileExists(sLocalFile)) and
               (CMsg.GetFileSize(sLocalFile) > 0) then
            begin
               // 이미지만 Preview 설정
               if (PosByte('.bmp', sLocalFile) > 0) or
                  (PosByte('.BMP', sLocalFile) > 0) or
                  (PosByte('.jpg', sLocalFile) > 0) or
                  (PosByte('.JPG', sLocalFile) > 0) or
                  {  -- TMS Grid에서 png 타입 미인식 오류로 주석 @ 2017.11.01 LSH
                  (PosByte('.png', sLocalFile) > 0) or
                  (PosByte('.PNG', sLocalFile) > 0) or
                  }
                  (PosByte('.gif', sLocalFile) > 0) or
                  (PosByte('.GIF', sLocalFile) > 0) then
               begin
                  // 수신한 Image 파일을 Grid에 표기
                  asg_Board.CreatePicture(C_B_REGDATE,
                                          asg_Board.Row + 1,
                                          False,
                                          ShrinkWithAspectRatio,
                                          0,
                                          haLeft,
                                          vaTop).LoadFromFile(sLocalFile);
               end
               else
                  asg_Board.Cells[C_B_REGDATE, asg_Board.Row + 1] := asg_Board.Cells[C_B_ATTACH, asg_Board.Row];
            end;
         end
         else
         begin
            // 프로필 이미지 Remove
            asg_Board.RemovePicture(C_B_REGDATE, asg_Board.Row + 1);
         end;




         asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row + 1]   := '┗';
         asg_Board.Cells[C_B_TITLE,    asg_Board.Row + 1]   := asg_Board.Cells[C_B_CONTEXT, asg_Board.Row];
         asg_Board.AddButton(C_B_LIKE, asg_Board.Row + 1, asg_Board.ColWidths[C_B_LIKE]-5,  20, 'ㅿ', haBeforeText, vaCenter);
         asg_Board.AutoSizeRow(asg_Board.Row + 1);



         if (asg_Board.Cells[C_B_REPLY, asg_Board.Row] > '0') then
            SetReplyList(asg_Board.Row,
                         asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row],
                         asg_Board.Cells[C_B_USERIP,   asg_Board.Row],
                         asg_Board.Cells[C_B_REGDATE,  asg_Board.Row]);



         Inc(iNowBoardCnt);



         if asg_Board.Cells[C_B_REPLY, asg_Board.Row] > '0' then
         begin
            for i := 0 to StrToInt(asg_Board.Cells[C_B_REPLY, asg_Board.Row]) - 1 do
               for j := C_B_BOARDSEQ to C_B_REPLY do
               begin
                  asg_Board.Colors[j, asg_Board.Row+1]   := $007CE1FF;
                  asg_Board.Colors[j, asg_Board.Row+2+i] := $0083C1E7;
               end;
         end
         else
         begin
            for j := C_B_BOARDSEQ to C_B_REPLY do
               asg_Board.Colors[j, asg_Board.Row+1]   := $007CE1FF;
         end;

         // 로그 Update
         UpdateLog('BOARD',
                   asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row],
                   asg_Board.Cells[C_B_USERIP,   asg_Board.Row],
                   'S',
                   '',
                   '',
                   varResult
                   );


      end
      else if (asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row]   <> '┗') and
              (asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row+1] = '┗') and
              (asg_Board.Cells[C_B_CLOSEYN,  asg_Board.Row] = 'O') then
      begin
         asg_Board.RemoveRows(asg_Board.Row + 1, 1 + StrToInt(asg_Board.Cells[C_B_REPLY, asg_Board.Row]));
         asg_Board.Cells[C_B_CLOSEYN, asg_Board.Row] := 'C';

         Dec(iNowBoardCnt);

      end
      else if (asg_Board.Cells[C_B_CLOSEYN, asg_Board.Row] = 'E') or
              (asg_Board.Cells[C_B_CLOSEYN, asg_Board.Row] = 'EC')then
      begin
         asg_Board.InsertRows(asg_Board.Row + 1, 1 + StrToInt(asg_Board.Cells[C_B_REPLY, asg_Board.Row]));
         asg_Board.Cells[C_B_CLOSEYN,  asg_Board.Row]       := 'EO';

         asg_Board.MergeCells(C_B_CATEUP, asg_Board.Row + 1, 4, 1);
         asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row + 1]   := '┗';
         asg_Board.Cells[C_B_TITLE,    asg_Board.Row + 1]   := asg_Board.Cells[C_B_CONTEXT, asg_Board.Row];
         asg_Board.AddButton(C_B_LIKE, asg_Board.Row + 1, asg_Board.ColWidths[C_B_LIKE]-5,  20, 'ㅿ', haBeforeText, vaCenter);
         asg_Board.AutoSizeRow(asg_Board.Row + 1);


         if (asg_Board.Cells[C_B_REPLY, asg_Board.Row] > '0') then
            SetReplyList(asg_Board.Row,
                         asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row],
                         asg_Board.Cells[C_B_USERIP,   asg_Board.Row],
                         asg_Board.Cells[C_B_REGDATE,  asg_Board.Row]);



         Inc(iNowBoardCnt);



         if asg_Board.Cells[C_B_REPLY, asg_Board.Row] > '0' then
         begin
            for i := 0 to StrToInt(asg_Board.Cells[C_B_REPLY, asg_Board.Row]) - 1 do
               for j := C_B_BOARDSEQ to C_B_REPLY do
               begin
                  asg_Board.Colors[j, asg_Board.Row+1]   := $007CE1FF;
                  asg_Board.Colors[j, asg_Board.Row+2+i] := $0083C1E7;
               end;
         end
         else
         begin
            for j := C_B_BOARDSEQ to C_B_REPLY do
               asg_Board.Colors[j, asg_Board.Row+1]   := $007CE1FF;
         end;


         // 로그 Update
         UpdateLog('BOARD',
                   asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row],
                   asg_Board.Cells[C_B_USERIP,   asg_Board.Row],
                   'S',
                   '',
                   '',
                   varResult
                   );

      end
      else if (asg_Board.Cells[C_B_CLOSEYN, asg_Board.Row] = 'EO') then
      begin
         asg_Board.RemoveRows(asg_Board.Row + 1, 1 + StrToInt(asg_Board.Cells[C_B_REPLY, asg_Board.Row]));
         asg_Board.Cells[C_B_CLOSEYN, asg_Board.Row] := 'EC';

         Dec(iNowBoardCnt);

      end;




      // [Best] 글 게시기간 Update
      if (PosByte('[★BEST★]', asg_Board.Cells[C_B_TITLE, asg_Board.Row]) > 0) then
      begin
         // 게시판 (커뮤니티) Best 게시기간(최초클릭 ~ 7일간) Update.
         UpdateBoard('6',
                     asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row],
                     FsUserIp,
                     asg_Board.Cells[C_B_REGDATE,  asg_Board.Row],
                     '',
                     '',
                     '',
                     '',
                     '',
                     '',
                     '',
                     '',
                     '',
                     asg_Board.Cells[C_B_HEADSEQ, asg_Board.Row],
                     '',
                     FormatDateTime('yyyy-mm-dd', DATE),
                     FormatDateTime('yyyy-mm-dd', DATE + 7),
                     '',
                     '');

      end;
   end;
end;



//------------------------------------------------------------------------------
// FlatTabControl onTabChanged Event Handler
//
// Date   : 2013.08.13
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.ftc_DialogTabChanged(Sender: TObject);
var
   FForm : TForm;
begin
{
   if apn_AnalList.Visible then
   begin
      apn_AnalList.Collaps := True;
      apn_AnalList.Visible := False;
   end;
}
   //-----------------------------------------------------
   // 공통 Panel Off
   //-----------------------------------------------------
   SetPanelStatus('ALL', 'OFF');



   case ftc_Dialog.ActiveTab of

      AT_DIALOG :
         begin
            fpn_Dialog.Top       := 23;
            fpn_Dialog.Left      := 7;
            fpn_Dialog.BringToFront;


            fsbt_NetworkClick(Sender);

            {
            apn_UserProfile.Visible := False;

            apn_Master.Left         := 999;
            apn_Master.Top          := 999;
            apn_Detail.Left         := 999;
            apn_Detail.Top          := 999;


            if (apn_Detail.Visible) then
            begin
               apn_Detail.Collaps := True;
               apn_Detail.Visible := False;
            end;

            if (apn_Master.Visible) then
            begin
               apn_Master.Collaps := True;
               apn_Master.Visible := False;
            end;


            apn_Network.Top      := 33; //30;
            apn_Network.Left     := 10; //7;
            apn_Network.Visible  := True;
            //apn_Network.BringToFront;

            SelGridInfo('UPDATE');
            SelGridInfo('DIALOG');

            tm_Master.Enabled := False;
            }


            //-----------------------------------------------------
            // 등록된 User 다이얼 Chat 리스트 팝업 @ 2015.03.30 LSH
            //-----------------------------------------------------
            if (IsLogonUser) then
            begin
               if (FindForm('KDIALCHAT') = nil) then
               begin
                  FForm := BplFormCreate('KDIALCHAT', True);

                  try
                     Screen.Cursor := crHourGlass;

                     SetBplStrProp(FForm, 'prop_UserNm', FsUserNm);

                     FForm.Height := (self.Height);
                     FForm.Top    := (screen.Height - FForm.Height) - 48;
                     FForm.Left   := 8;

                     try
                        FForm.Show;

                     except
                        on e : Exception do
                           showmessage('Show Form Error(KDIALCHAT) : ' + e.Message);
                     end;

                  finally
                     Screen.Cursor := crDefault;
                  end;

                  self.SetFocus;

               end;
            end;

         end;


      { -- 비상연락망 Tab 통합에 따른 주석 @ 2013.12.30 LSH
      AT_DETAIL :
         begin
            fpn_Detail.Top  := 23;
            fpn_Detail.Left := 7;
            fpn_Detail.BringToFront;

            SelGridInfo('DETAIL');

            tm_Master.Enabled := False;
         end;

      AT_MASTER :
         begin
            fpn_Master.Top  := 23;
            fpn_Master.Left := 7;
            fpn_Master.BringToFront;

            SelGridInfo('MASTER');


            tm_Master.Enabled := True;
         end;
      }

      AT_DOC :
         begin
            fpn_Doc.Top  := 23;
            fpn_Doc.Left := 7;
            fpn_Doc.BringToFront;

            // 최초 [문서관리] 클릭시에만 Release 대장 조회하고(Tag = 0)
            // 이후 부터는 마지막으로 활성화 했던 Panel을 띄우도록 변경 @ 2015.06.15 LSH
            if (fsbt_Release.Tag = 0) then
               fsbt_ReleaseClick(Sender);
         end;

      AT_BOARD :
         begin
            fpn_Board.Top  := 23;
            fpn_Board.Left := 7;
            fpn_Board.BringToFront;

            fpn_Write.Visible := False;
            fpn_Write.SendToBack;

            // 게시판 Page Block 단위 조회
            SelGridInfo('BOARDPAGE');

            SelGridInfo('BOARD');

            tm_Master.Enabled := False;
         end;

      AT_DIALBOOK :
         begin
            fpn_DialBook.Top  := 23;
            fpn_DialBook.Left := 7;
            fpn_DialBook.BringToFront;

            asg_DialScan.Visible  := False;
            apn_LinkList.Visible  := False;
            apn_DialMap.Visible   := False;
            apn_AddMyDial.Visible := False;

            asg_DialScan.ClearRows(1, asg_DialScan.RowCount);
            asg_DialScan.RowCount := 2;

            asg_DialList.ClearRows(1, asg_DialList.RowCount);
            asg_DialList.RowCount := 2;

            // [안암] 접속자 IP 식별후, 근무처 다이얼 Map 정보표기 @ 2013.11.05 LSH
            if PosByte('안암도메인', FsUserIp) > 0 then
               fsbt_DialMap.Visible := True
            else
               fsbt_DialMap.Visible := False;

            // 나의 다이얼 Book 조회
            SelGridInfo('MYDIAL');

            tm_Master.Enabled := False;

            fcb_Scan.Text := '통합검색';
            fed_Scan.Text := '';
            fed_Scan.CanFocus;
            fed_Scan.SetFocus;

            lb_DialScan.Caption := '▶ [전화번호부], [다이얼로그] 및 [S/R요청내역]을 연동한 통합검색을 제공합니다 ^^b';

            //-----------------------------------------------------
            // 등록된 User 다이얼 Chat 리스트 팝업 @ 2015.03.30 LSH
            //-----------------------------------------------------
            if (IsLogonUser) then
            begin
               FForm := FindForm('KDIALCHAT');

               if (FForm = nil) then
               begin
                  FForm := BplFormCreate('KDIALCHAT', True);

                  try
                     Screen.Cursor := crHourGlass;

                     SetBplStrProp(FForm, 'prop_UserNm', FsUserNm);

                     FForm.Height := (self.Height);
                     FForm.Top    := (screen.Height - FForm.Height) - 48;
                     FForm.Left   := 8;

                     try
                        FForm.Show;

                     except
                        on e : Exception do
                           showmessage('Show Form Error(KDIALCHAT) : ' + e.Message);
                     end;
                  finally
                     Screen.Cursor := crDefault;
                  end;

                  self.SetFocus;

               end;
            end;

         end;


      AT_ANALYSIS :
         begin
            fpn_Analysis.Top  := 23;
            fpn_Analysis.Left := 7;
            fpn_Analysis.BringToFront;


            // D/B 마스터 사용권한은 [운영/개발파트]로 제한
            if (PosByte('운영', FsUserPart) > 0) or
               (PosByte('개발', FsUserPart) > 0) then
               fsbt_DBMaster.Visible := True
            else
               fsbt_DBMaster.Visible := False;


            {
            asg_WorkRpt.Top      := 999;
            asg_WorkRpt.Left     := 999;
            asg_WorkRpt.Visible  := False;


            fmed_AnalFrom.Text   := FormatDateTime('yyyy-mm-dd', DATE - 1095);   // 약 3년치 검색
            fmed_AnalTo.Text     := FormatDateTime('yyyy-mm-dd', DATE);


            asg_Analysis.Top     := 62;
            asg_Analysis.Left    := 6;
            asg_Analysis.Visible := True;

            fed_Analysis.Visible := True;

            fcb_WorkArea.Top     := 999;
            fcb_WorkArea.Left    := 999;
            fcb_WorkArea.Visible := False;

            fcb_WorkGubun.Top     := 999;
            fcb_WorkGubun.Left    := 999;
            fcb_WorkGubun.Visible := False;

            asg_Analysis.ClearRows(1, asg_Analysis.RowCount);
            asg_Analysis.RowCount := 2;

            apn_AnalList.Visible := False;

            lb_Analysis.Caption := '▶ S/R 요청내역 분석을 위한 효율적인 검색을 제공합니다.';
            }

            fcb_Analysis.Text := '통합검색';
            fed_Analysis.Text := '';
            fed_Analysis.CanFocus;
            fed_Analysis.SetFocus;


            //fsbt_AnalysisClick(Sender);
            fsbt_WorkRptClick(Sender);


            tm_Master.Enabled := False;



         end;

      //------------------------------------------------------------------------
      // KKP 팀장님 order로 병원별 업무공유 등록/조회 개설 @ 2014.08.08 LSH
      //------------------------------------------------------------------------
      {  -- 사용률 저하로 제거 @ 2017.10.31 LSH
      AT_WORKCONN :
         begin
            fpn_WorkConn.Top  := 23;
            fpn_WorkConn.Left := 7;
            fpn_WorkConn.BringToFront;


//            fcb_Analysis.Text := '통합검색';
//            fed_Analysis.Text := '';
//            fed_Analysis.CanFocus;
//            fed_Analysis.SetFocus;
//
//
//            //fsbt_AnalysisClick(Sender);
//            fsbt_WorkRptClick(Sender);



            // 기간별 주요업무공유 조회
            SelGridInfo('WORKCONN');


            tm_Master.Enabled := False;

         end;
      }
   end;
end;



//------------------------------------------------------------------------------
// FlatSpeedButton onClick Event Handler
//
// Date   : 2013.08.13
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.fsbt_WriteClick(Sender: TObject);
begin
   if not IsLogonUser then
   begin
      MessageBox(self.Handle,
                 PChar('[비상연락망] - [담당자정보]에 담당자 IP 등록후 사용하실 수 있습니다.'),
                 PChar(Self.Caption + ' : 등록되지 않은 User 제한 알림'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;

   fpn_Write.Top     := 0;
   fpn_Write.Left    := 0;
   fpn_Write.Visible := True;
   fpn_Write.BringToFront;

   fed_BoardSeq.Text    := '';

   lb_HeadTail.Caption  := 'H';
   lb_HeadSeq.Caption   := '';
   lb_TailSeq.Caption   := '1';

   fcb_CateUp.Enabled   := True;
   fed_Writer.Enabled   := True;
   fmm_Text.Enabled     := True;
   fed_Title.Enabled    := True;
   fed_Title.Text       := '';
   fed_Title.Hint       := '';
   fmm_Text.Hint        := '';

   if (PosByte('.', FsNickNm) = 0) and
      (FsNickNm <> '') then
      fed_Writer.Text      := FsNickNm
   else
      fed_Writer.Text      := FsUserNm;

   fmm_Text.Lines.Text  := '';
   fed_Attached.Text    := '';
   fcb_CateUp.Text      := '';
   fcb_CateUp.SetFocus;

   fed_CateDownNm.Visible  := False;
   fcb_CateDown.Visible    := False;
   fed_AlarmFro.Visible    := False;
   fed_AlarmTo.Visible     := False;
   fmed_AlarmFro.Visible   := False;
   fmed_AlarmTo.Visible    := False;

   fsbt_FileOpen.Enabled  := True;
   fsbt_FileClear.Enabled := True;

   fsbt_WriteReg.Visible  := True;
   fsbt_WriteFind.Visible := False;
   fsbt_WriteReg.Top      := 480;
   fsbt_WriteReg.Left     := 308;

   fed_Title.Width    := 510;
   fmm_Text.Width     := 510;
   fmm_Text.MaxLength := 500;
   fmm_Text.Hint      := '신규 작성글은 한글기준 15줄(500자 정도)이내로 제한합니다.';

   lb_Write.Caption    := '▶ 신규 글 작성 모드입니다.';
   lb_Write.Font.Color := $001079AF;
   lb_Write.Font.Style := [fsBold];
end;


//------------------------------------------------------------------------------
// FlatSpeedButton onClick Event Handler
//
// Date   : 2013.08.14
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.fsbt_WriteCancelClick(Sender: TObject);
begin
   fpn_Write.Top     := 999;
   fpn_Write.Left    := 999;
   fpn_Write.Visible := False;
end;



//------------------------------------------------------------------------------
// FlatSpeedButton onClick Event Handler
//
// Date   : 2013.08.14
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.fsbt_WriteRegClick(Sender: TObject);
var
   sType, sTitle, sHeadTail, sHeadSeq, sTailSeq, sFileName, sHideFile, sServerIp,
   sAlarmFro, sAlarmTo, sDelDate : String;
begin
   //-------------------------------------------------------------------
   // 2. Create Variables
   //-------------------------------------------------------------------
   if (lb_HeadTail.Caption = 'H') and
      (fcb_CateUp.Text = '') then
   begin
      MessageBox(self.Handle,
                 PChar('글 신규 등록시 [분류]항목은 필수입력 입니다.'),
                 PChar(Self.Caption + ' : 커뮤니티 필수입력 항목 알림'),
                 MB_OK + MB_ICONWARNING);

      fcb_CateUp.SetFocus;

      Exit;
   end;

   {  -- 제한없이 공지사항 관리할 수 있게 하기 위해 주석 @ 2018.03.20 LSH
   if (lb_HeadTail.Caption = 'H') and
      (PosByte('공지', fcb_CateUp.Text) > 0) and
      (FsUserIp <> '개발자IP') then
   begin
      MessageBox(self.Handle,
                 PChar('[공지사항] 등록은 관리자만 가능합니다. 양해 부탁드려요 ^^'),
                 PChar(Self.Caption + ' : 커뮤니티 공지사항 등록제한 알림'),
                 MB_OK + MB_ICONWARNING);

      fcb_CateUp.SetFocus;

      Exit;
   end;
   }

   if (lb_HeadTail.Caption = 'H') and
      (PosByte('공지', fcb_CateUp.Text) > 0) and
      ((fmed_AlarmFro.Text = '    -  -  ') or (fmed_AlarmTo.Text = '    -  -  '))  then
   begin
      MessageBox(self.Handle,
                 PChar('[공지사항] 등록시, 공지기간 (From~To) 필수입력 입니다.'),
                 PChar(Self.Caption + ' : 커뮤니티 공지사항 기간등록 알림'),
                 MB_OK + MB_ICONWARNING);

      fmed_AlarmFro.SetFocus;

      Exit;
   end;

   if (lb_HeadTail.Caption = 'H') and
      (PosByte('공지', fcb_CateUp.Text) > 0) and
      (fmed_AlarmFro.Text  > fmed_AlarmTo.Text)  then
   begin
      MessageBox(self.Handle,
                 PChar('[공지사항] 등록시, 공지기간 To가 From보다 앞설 수 없습니다.'),
                 PChar(Self.Caption + ' : 커뮤니티 공지사항 기간등록 알림'),
                 MB_OK + MB_ICONWARNING);

      fmed_AlarmTo.SetFocus;

      Exit;
   end;


   if (fed_Writer.Text = '') then
   begin
      MessageBox(self.Handle,
                 PChar('작성자(닉네임)을 입력해 주시기 바랍니다.'),
                 PChar(Self.Caption + ' : 커뮤니티 필수입력 항목 알림'),
                 MB_OK + MB_ICONWARNING);

      fed_Writer.SetFocus;

      Exit;
   end;

   if (fed_Title.Text = '') then
   begin
      MessageBox(self.Handle,
                 PChar('[글 제목]을 입력해 주시기 바랍니다.'),
                 PChar(Self.Caption + ' : 커뮤니티 필수입력 항목 알림'),
                 MB_OK + MB_ICONWARNING);

      fed_Title.SetFocus;

      Exit;
   end;

   if (fmm_Text.Lines.Text = '') then
   begin
      MessageBox(self.Handle,
                 PChar('[내용]에 입력하신 내용이 없습니다.' + #13#10 + '아무 내용이라도 하나 적어주세요 ^^'),
                 PChar(Self.Caption + ' : 커뮤니티 필수입력 항목 알림'),
                 MB_OK + MB_ICONWARNING);

      fmm_Text.SetFocus;

      Exit;
   end;

   if (lb_HeadTail.Caption = 'H') then
   begin
      if (fmm_Text.Lines.Count > 15) then
      begin
         MessageBox(self.Handle,
                    PChar('신규 글작성은 15줄 이내로 제한합니다.'),
                    PChar(Self.Caption + ' : 커뮤니티 글 작성 제한 알림'),
                    MB_OK + MB_ICONWARNING);

         fmm_Text.SetFocus;

         Exit;
      end;
   end
   else if (lb_HeadTail.Caption = 'T') then
   begin
      if (fmm_Text.Lines.Count > 5) then
      begin
         MessageBox(self.Handle,
                    PChar('댓글작성은 스크롤의 압박을 방지하기위해 5줄 이내로 제한합니다.'),
                    PChar(Self.Caption + ' : 커뮤니티 글 작성 제한 알림'),
                    MB_OK + MB_ICONWARNING);

         fmm_Text.SetFocus;

         Exit;
      end;
   end;


   // 공지사항 Flag
   if (PosByte('공지', fcb_CateUp.Text) > 0) and
      (lb_HeadTail.Caption = 'H') then
   begin
      sHeadTail   := 'A';
      sAlarmFro   := fmed_AlarmFro.Text;
      sAlarmTo    := fmed_AlarmTo.Text;
   end
   else
   begin
      sHeadTail   := lb_HeadTail.Caption;
      sAlarmFro   := '';
      sAlarmTo    := '';
   end;


   sType       := '4';
   sTitle      := fed_Title.Text;
   sHeadSeq    := lb_HeadSeq.Caption;
   sTailSeq    := lb_TailSeq.Caption;
   sFileName   := '';
   sHideFile   := '';
   sServerIp   := '';
   sDelDate    := '';


   // 첨부파일 Upload
   if (Trim(fed_Attached.Text) <> '') then
   begin

      sFileName := ExtractFileName(Trim(fed_Attached.Text));
      sHideFile := 'KDIALAPPEND' + Trim(fed_BoardSeq.Text) + FormatDateTime('YYYYMMDDHHNNSS', Now) + '.KDIALFILE';


      if Not FileUpLoad(sHideFile, asg_Board.Cells[C_B_ATTACH, asg_Board.Row], sServerIp) then
      begin
         Showmessage('첨부파일 UpLoad 중 에러가 발생했습니다.' + #13#10 + #13#10 +
                     '다시한번 시도해 주시기 바랍니다.');
         Exit;
      end;
   end;



   // 게시판 (커뮤니티) Update.
   UpdateBoard(sType,
               Trim(fed_BoardSeq.Text),
               FsUserIp,
               FormatDateTime('yyyy-mm-dd hh:nn', Now),
               fed_Writer.Text,
               fcb_CateUp.Text,
               fcb_CateDown.Text,
               sTitle,
               fmm_Text.Lines.Text,
               sFileName,
               sHideFile,
               sServerIp,
               sHeadTail,
               sHeadSeq,
               sTailSeq,
               sAlarmFro,
               sAlarmTo,
               sDelDate,
               FsUserIp);



   fsbt_WriteCancelClick(Sender);


end;



//------------------------------------------------------------------------------
// FlatSpeedButton onClick Event Handler
//
// Date   : 2013.08.14
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.fsbt_ReplyClick(Sender: TObject);
begin
   if not IsLogonUser then
   begin
      MessageBox(self.Handle,
                 PChar('[비상연락망] - [담당자정보]에 담당자 IP 등록후 사용하실 수 있습니다.'),
                 PChar(Self.Caption + ' : 등록되지 않은 User 제한 알림'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;

   fpn_Write.Top     := 0;
   fpn_Write.Left    := 0;
   fpn_Write.Visible := True;
   fpn_Write.BringToFront;


   fed_BoardSeq.Text    := '';
   fed_Title.Text       := '[댓글] ' + asg_Board.Cells[C_B_TITLE, asg_Board.Row];
   lb_HeadTail.Caption  := 'T';
   lb_HeadSeq.Caption   := asg_Board.Cells[C_B_HEADSEQ, asg_Board.Row];
   lb_TailSeq.Caption   := IntToStr(StrToInt(asg_Board.Cells[C_B_REPLY,  asg_Board.Row]) + 2);

   fcb_CateUp.Text      := asg_Board.Cells[C_B_CATEUP,  asg_Board.Row];
   fcb_CateUp.Enabled   := False;
   fed_Writer.Enabled   := True;
   fed_Title.Enabled    := False;
   fmm_Text.Enabled     := True;

   if (PosByte('.', FsNickNm) = 0) and
      (FsNickNm <> '') then
      fed_Writer.Text      := FsNickNm
   else
      fed_Writer.Text      := FsUserNm;

   fmm_Text.Lines.Text  := '';
   fed_Attached.Text    := '';

   fed_CateDownNm.Visible := False;
   fcb_CateDown.Visible   := False;

   fsbt_FileOpen.Enabled  := True;
   fsbt_FileClear.Enabled := True;

   fmm_Text.SetFocus;

   fsbt_WriteReg.Visible  := True;
   fsbt_WriteFind.Visible := False;
   fsbt_WriteReg.Top      := 480;
   fsbt_WriteReg.Left     := 308;


   fed_Title.Width    := 350;
   fmm_Text.Width     := 350;
   fmm_Text.MaxLength := 400;
   fmm_Text.Hint      := '스크롤의 압박을 피하기위해 댓글은 4~5줄(한글기준 200자)이내로 제한합니다.';

   lb_Write.Caption    := '▶ 댓글 작성 모드입니다.';
   lb_Write.Font.Color := $004ABDAF;
   lb_Write.Font.Style := [fsBold];
end;



//------------------------------------------------------------------------------
// FlatSpeedButton onClick Event Handler
//
// Date   : 2013.08.16
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.fsbt_DeleteClick(Sender: TObject);
begin
   if not IsLogonUser then
   begin
      MessageBox(self.Handle,
                 PChar('[비상연락망] - [담당자정보]에 담당자 IP 등록후 사용하실 수 있습니다.'),
                 PChar(Self.Caption + ' : 등록되지 않은 User 제한 알림'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;

   fed_BoardSeq.Text    := asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row];
   fed_Title.Text       := '';
   lb_HeadTail.Caption  := '';
   lb_HeadSeq.Caption   := '';
   lb_TailSeq.Caption   := '';


   // 작성자 본인 IP 확인
   if (FsUserIp <> asg_Board.Cells[C_B_USERIP, asg_Board.Row]) then
   begin
      MessageBox(self.Handle,
                 PChar('작성자 본인의 게시글만 삭제가 가능합니다.'),
                 '[KUMC 다이얼로그] 커뮤니티 본인 확인 알림 ',
                 MB_OK + MB_ICONINFORMATION);

      Exit;
   end;


   // 삭제 Confirm Message.
   if Application.MessageBox('선택한 게시글을 [삭제]하시겠습니까?',
                             PChar(self.Caption + ' : 커뮤니티 삭제 알림'), MB_OKCANCEL) <> IDOK then
      Exit;


   if (asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row] <> '┗') and
      (asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row] <> '')   then
   begin
      // 게시판 (커뮤니티) 원본글 삭제 Update.
      UpdateBoard('4',
                  asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row],
                  FsUserIp,
                  asg_Board.Cells[C_B_REGDATE, asg_Board.Row],
                  '',
                  '',
                  '',
                  '',
                  '',
                  '',
                  '',
                  '',
                  '',
                  '',
                  '',
                  '',
                  '',
                  FormatDateTime('yyyy-mm-dd', Date),
                  FsUserIp);
   end
   else
   begin
      // 게시판 (커뮤니티) 댓글 삭제 Update.
      UpdateBoard('4',
                  asg_Board.Cells[C_B_CONTEXT, asg_Board.Row],
                  FsUserIp,
                  asg_Board.Cells[C_B_REGDATE, asg_Board.Row],
                  '',
                  '',
                  '',
                  '',
                  '',
                  '',
                  '',
                  '',
                  '',
                  '',
                  '',
                  '',
                  '',
                  FormatDateTime('yyyy-mm-dd', Date),
                  FsUserIp);

   end;
end;



//------------------------------------------------------------------------------
// [조회] 게시판 (커뮤니티) Update Procedure
//       - 신규/댓글/삭제 Upd.
//
// Author : Lee, Se-Ha
// Date   : 2013.08.19
//------------------------------------------------------------------------------
procedure TMainDlg.UpdateBoard(in_Type,
                               in_BoardSeq,
                               in_UserIp,
                               in_RegDate,
                               in_RegUser,
                               in_CateUp,
                               in_CateDown,
                               in_Title,
                               in_Context,
                               in_AttachNm,
                               in_HideFile,
                               in_ServerIp,
                               in_HeadTail,
                               in_HeadSeq,
                               in_TailSeq,
                               in_AlarmFro,
                               in_AlarmTo,
                               in_DelDate,
                               in_EditIp : String);
var
   vType1   ,
   vBoardSeq,
   vUserIp  ,
   vRegDate ,
   vRegUser ,
   vCateUp  ,
   vCateDown,
   vTitle   ,
   vContext ,
   vAttachNm,
   vHideFile,
   vServerIp,
   vHeadTail,
   vHeadSeq,
   vTailSeq,
   vAlarmFro,
   vAlarmTo,
   vDelDate ,
   vEditIp : Variant;
   TpPutBoard : TTpSvc;
begin


   //------------------------------------------------------------------
   // 2-1. Append Variables
   //------------------------------------------------------------------
   AppendVariant(vType1   ,  in_Type);
   AppendVariant(vBoardSeq,  in_BoardSeq);
   AppendVariant(vUserIp  ,  in_UserIp);
   AppendVariant(vRegDate ,  in_RegDate);
   AppendVariant(vRegUser ,  in_RegUser);
   AppendVariant(vCateUp  ,  in_CateUp);
   AppendVariant(vCateDown,  in_CateDown);
   AppendVariant(vTitle   ,  in_Title);
   AppendVariant(vContext ,  in_Context);
   AppendVariant(vAttachNm,  in_AttachNm);
   AppendVariant(vHideFile,  in_HideFile);
   AppendVariant(vServerIp,  in_ServerIp);
   AppendVariant(vHeadTail,  in_HeadTail);
   AppendVariant(vHeadSeq ,  in_HeadSeq);
   AppendVariant(vTailSeq ,  in_TailSeq);
   AppendVariant(vAlarmFro,  in_AlarmFro);
   AppendVariant(vAlarmTo ,  in_AlarmTo);
   AppendVariant(vDelDate ,  in_DelDate);
   AppendVariant(vEditIp  ,  in_EditIp);




   //-------------------------------------------------------------------
   // 3. Insert by TpSvc
   //-------------------------------------------------------------------
   TpPutBoard := TTpSvc.Create;
   TpPutBoard.Init(Self);



   Screen.Cursor := crHourGlass;



   try
      if TpPutBoard.PutSvc('MD_KUMCM_M1',
                          [
                            'S_TYPE1'   , in_Type
                         ,  'S_STRING25', in_BoardSeq
                         ,  'S_STRING7' , in_UserIp
                         ,  'S_STRING20', in_RegDate
                         ,  'S_STRING21', in_RegUser
                         ,  'S_STRING26', in_CateUp
                         ,  'S_STRING27', in_CateDown
                         ,  'S_STRING19', in_Title
                         ,  'S_STRING28', in_Context
                         ,  'S_STRING29', in_AttachNm
                         ,  'S_STRING30', in_HideFile
                         ,  'S_STRING31', in_ServerIp
                         ,  'S_STRING32', in_HeadTail
                         ,  'S_STRING33', in_HeadSeq
                         ,  'S_STRING34', in_TailSeq
                         ,  'S_STRING39', in_AlarmFro
                         ,  'S_STRING40', in_AlarmTo
                         ,  'S_STRING10', in_DelDate
                         ,  'S_STRING8' , in_EditIp
                         ] ) then
      begin

         if (in_Type <> '6') then
         begin
            if (in_DelDate = '') then
               MessageBox(self.Handle,
                          PChar('[' + in_Title + '] ' + ' 게시글이 정상적으로 [업데이트] 되었습니다.'),
                          '[KUMC 다이얼로그] 커뮤니티 업데이트 알림 ',
                          MB_OK + MB_ICONINFORMATION)
            else
               MessageBox(self.Handle,
                          PChar('게시글이 정상적으로 [삭제] 되었습니다.'),
                          '[KUMC 다이얼로그] 커뮤니티 업데이트 알림 ',
                          MB_OK + MB_ICONINFORMATION);

            // Refresh
            SelGridInfo('BOARDPAGE');
            SelGridInfo('BOARD');
         end;
      end
      else
      begin
         ShowMessageM(GetTxMsg);
      end;



   finally
      FreeAndNil(TpPutBoard);
      Screen.Cursor  := crDefault;
   end;
end;





//------------------------------------------------------------------------------
// [조회] 게시판 (커뮤니티) 댓글 상세조회
//
// Author : Lee, Se-Ha
// Date   : 2013.08.20
//------------------------------------------------------------------------------
procedure TMainDlg.SetReplyList(in_HeaderRow : Integer;
                                in_BoardSeq,
                                in_UserIp,
                                in_RegDate : String);
var
   TpGetReply  : TTpSvc;
   sType       : String;
   i, iRowCnt  : Integer;
begin


   sType := '10';


   //-----------------------------------------------------------------
   // 2-1. Load Data by TpSvc
   //-----------------------------------------------------------------
   TpGetReply := TTpSvc.Create;
   TpGetReply.Init(Self);


   Screen.Cursor := crHourglass;


   try
      TpGetReply.CountField  := 'S_CODE21';
      TpGetReply.ShowMsgFlag := False;

      if TpGetReply.GetSvc('MD_KUMCM_S1',
                          ['S_TYPE1'  , sType
                         , 'S_TYPE2'  , in_BoardSeq
                         , 'S_TYPE3'  , in_UserIp
                         , 'S_TYPE4'  , in_RegDate
                          ],
                          [
                           'S_CODE9'  , 'sUserIp'
                         , 'S_CODE13' , 'sDelDate'
                         , 'S_CODE16' , 'sEditIp'
                         , 'S_CODE17' , 'sEditDate'
                         , 'S_CODE21' , 'sDocTitle'
                         , 'S_CODE22' , 'sRegDate'
                         , 'S_CODE23' , 'sRegUser'
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
                         , 'S_STRING12','sLikeCnt'
                          ]) then

         if TpGetReply.RowCount < 0 then
         begin
            ShowMessage(GetTxMsg);
            Exit;
         end
         else if TpGetReply.RowCount = 0 then
         begin
            //stb_Message.Panels[0].Text := '해당기간내 검색된 자료가 없습니다.';
            Exit;
         end;



      with asg_Board do
      begin
         iRowCnt  := TpGetReply.RowCount;

         for i := 0 to iRowCnt - 1 do
         begin
            Cells[C_B_BOARDSEQ,  in_HeaderRow + 2 + i] := '';
            Cells[C_B_CATEUP,    in_HeaderRow + 2 + i] := '┗';
            Cells[C_B_TITLE,     in_HeaderRow + 2 + i] := TpGetReply.GetOutputDataS('sContext' , i);
            Cells[C_B_REGDATE,   in_HeaderRow + 2 + i] := TpGetReply.GetOutputDataS('sRegDate' , i);
            Cells[C_B_REGUSER,   in_HeaderRow + 2 + i] := TpGetReply.GetOutputDataS('sRegUser' , i);
            Cells[C_B_LIKE,      in_HeaderRow + 2 + i] := TpGetReply.GetOutputDataS('sLikeCnt' , i);
            Cells[C_B_ATTACH,    in_HeaderRow + 2 + i] := TpGetReply.GetOutputDataS('sAttachNm', i);
            Cells[C_B_HEADSEQ,   in_HeaderRow + 2 + i] := TpGetReply.GetOutputDataS('sHeadSeq' , i);
            Cells[C_B_TAILSEQ,   in_HeaderRow + 2 + i] := TpGetReply.GetOutputDataS('sTailSeq' , i);
            Cells[C_B_CONTEXT,   in_HeaderRow + 2 + i] := TpGetReply.GetOutputDataS('sBoardSeq', i);
            Cells[C_B_USERIP ,   in_HeaderRow + 2 + i] := TpGetReply.GetOutputDataS('sUserIp'  , i);
            Cells[C_B_REPLY,     in_HeaderRow + 2 + i] := TpGetReply.GetOutputDataS('sReplyCnt', i);
            Cells[C_B_HIDEFILE,  in_HeaderRow + 2 + i] := TpGetReply.GetOutputDataS('sHideFile', i);
            Cells[C_B_SERVERIP,  in_HeaderRow + 2 + i] := TpGetReply.GetOutputDataS('sServerIp', i);



            //------------------------------------------------------------
            // 추천 버튼 표기
            //------------------------------------------------------------
            //AddButton(C_B_ATTACH, in_HeaderRow + 2 + i, ColWidths[C_B_ATTACH]-5,  20, 'ㅿ', haBeforeText, vaCenter);
            AddButton(C_B_LIKE, in_HeaderRow + 2 + i, ColWidths[C_B_LIKE]-5,  20, 'ㅿ', haCenter, vaUnderText);


            //------------------------------------------------------------
            // 첨부 파일 표기
            //------------------------------------------------------------
            if TpGetReply.GetOutputDataS('sAttachNm', i) <> '' then
               AddButton(C_B_ATTACH,  in_HeaderRow + 2 + i, ColWidths[C_B_ATTACH]-5,  20, 'Λ', haBeforeText, vaCenter);


            // 첨부파일명 길이만큼, AutoSizing 되는 문제를 제거하기 위해 주석처리 --> 추천횟수 + 버튼표기 위해 다시 주석 해제
            AutoSizeRow(in_HeaderRow + 2 + i);
         end;
      end;




   finally
      FreeAndNil(TpGetReply);
      Screen.Cursor := crDefault;
   end;

end;




//------------------------------------------------------------------------------
// [Update] 게시판 Log 관리
//
// Author : Lee, Se-Ha
// Date   : 2013.08.20
//------------------------------------------------------------------------------
procedure TMainDlg.UpdateLog(in_Gubun,
                             in_BoardSeq,
                             in_UserIp,
                             in_LogFlag,
                             in_Item1,
                             in_Item2  : String;
                             var varResult : String);
var
   TpUpdLog : TTpSvc;
   vType1   ,
   vBoardNm ,
   vBoardSeq,
   vUserIp  ,
   vLogFlag ,
   vItem1   ,
   vItem2   ,
   vEditIp  : Variant;
begin
   varResult := 'N';


   //------------------------------------------------------------------
   // 2-1. Append Variables
   //------------------------------------------------------------------
   AppendVariant(vType1   ,  '5');
   AppendVariant(vBoardNm ,  in_Gubun);
   AppendVariant(vBoardSeq,  in_BoardSeq);
   AppendVariant(vUserIp  ,  in_UserIp);//FsUserIp);
   AppendVariant(vLogFlag ,  in_LogFlag);
   AppendVariant(vItem1   ,  in_Item1);
   AppendVariant(vItem2   ,  in_Item2);
   AppendVariant(vEditIp  ,  FsUserIp);



   //-------------------------------------------------------------------
   // 3. Insert by TpSvc
   //-------------------------------------------------------------------
   TpUpdLog := TTpSvc.Create;
   TpUpdLog.Init(Self);



   Screen.Cursor := crHourGlass;



   try
      if TpUpdLog.PutSvc('MD_KUMCM_M1',
                          [
                            'S_TYPE1'   , vType1
                         ,  'S_STRING35', vBoardNm
                         ,  'S_STRING25', vBoardSeq
                         ,  'S_STRING7' , vUserIp
                         ,  'S_STRING36', vLogFlag
                         ,  'S_STRING37', vItem1
                         ,  'S_STRING38', vItem2
                         ,  'S_STRING8' , vEditIp
                         ] ) then
      begin
         if (in_LogFlag = 'L') then
            MessageBox(self.Handle,
                       PChar('선택한 게시글이 정상적으로 [추천] 되었습니다.'),
                       '[KUMC 다이얼로그] 커뮤니티 업데이트 알림 ',
                       MB_OK + MB_ICONINFORMATION);

         varResult := 'Y';
      end
      else
      begin
         // SVC 단에서 GetTxMsg 처리중. 별도 처리 필요 없음.
         varResult := 'N';
      end;


   finally
      FreeAndNil(TpUpdLog);
      Screen.Cursor  := crDefault;
   end;

end;


//------------------------------------------------------------------------------
// AdvStringGrid onClick Event Handler
//
// Date   : 2013.08.16
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_BoardClick(Sender: TObject);
begin
   if (asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row] = '') or
      (asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row] = '┗') then
      fsbt_Reply.Enabled := False
   else
      fsbt_Reply.Enabled := True;


   if (asg_Board.Cells[C_B_USERIP, asg_Board.Row] = '') then
      fsbt_Delete.Enabled := False
   else
      fsbt_Delete.Enabled := True;

   if (asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row] <> '') and
      (asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row] <> '┗') then
      lb_Board.Caption := '▶ No.' + asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row] + ' ' + asg_Board.Cells[C_B_TITLE, asg_Board.Row];

end;



//------------------------------------------------------------------------------
// AdvStringGrid onGetAlignment Event Handler
//
// Date   : 2013.08.19
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_BoardGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   // Horizontal Alignment
   if (ARow = 0) or (ACol = C_B_BOARDSEQ)
                 or ((ARow > 0) and (asg_Board.Cells[C_B_USERIP, ARow] <> '')
                                and
                                    (
                                     (ACol = C_B_CATEUP)  or
                                     (ACol = C_B_REGDATE) or
                                     (ACol = C_B_REGUSER) or
                                     ((ACol = C_B_LIKE) {and (asg_Board.Cells[C_B_BOARDSEQ, ARow] <> '')})
                                     //(ACol = C_B_ATTACH)
                                     )) then
      HAlign := taCenter;

   // Vertical Alignment
   VAlign := vtaCenter;
end;



//------------------------------------------------------------------------------
// AdvStringGrid onButtonClick Event Handler
//
// Date   : 2013.08.19
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_BoardButtonClick(Sender: TObject; ACol,
  ARow: Integer);
var
   sLocalFile, sRemoteFile : String;
   sServerIp, sFtpUserID, sFtpPasswd, sFtpHostName, sFtpDir, varUpResult : String;
begin
   if not IsLogonUser then
   begin
      MessageBox(self.Handle,
                 PChar('[비상연락망] - [담당자정보]에 담당자 IP 등록후 사용하실 수 있습니다.'),
                 PChar(Self.Caption + ' : 등록되지 않은 User 제한 알림'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;


   with asg_Board do
   begin
      if (ACol = C_B_LIKE) then
         // or ((ACol = C_B_ATTACH) and (Cells[C_B_BOARDSEQ, ARow] = '')) then
      begin
         if Cells[C_B_BOARDSEQ, ARow] = '┗' then
            UpdateLog('BOARD',
                      Cells[C_B_BOARDSEQ, ARow - 1],
                      FsUserIp{Cells[C_B_USERIP,   ARow - 1]},
                      'L',
                      '',
                      FsUserNm,
                      varUpResult
                      )
         else if Cells[C_B_CATEUP, ARow] = '┗' then
            UpdateLog('BOARD',
                      Cells[C_B_CONTEXT, ARow],
                      FsUserIp{Cells[C_B_USERIP,  ARow]},
                      'L',
                      '',
                      FsUserNm,
                      varUpResult);

      end;



      //---------------------------------------------
      // 첨부파일 Click시 Event 처리
      //---------------------------------------------
      if ((ACol = C_B_ATTACH) and (Cells[C_B_HIDEFILE, ARow] <> '')) then
      begin
         // 파일 업/다운로드를 위한 정보 조회
         if not GetBinUploadInfo(sServerIp, sFtpUserID, sFtpPasswd, sFtpHostName, sFtpDir) then
         begin
            MessageDlg('파일 저장을 위한 담당자 정보 조회중, 오류가 발생했습니다.', Dialogs.mtError, [Dialogs.mbOK], 0);
            exit;
         end;


         sServerIp := Cells[C_B_SERVERIP, ARow]; //lb_ServerIp.Caption;    // 실제저장된 서버 IP


         // 실제 서버에 저장되어 있는 파일명 지정
         if PosByte('/ftpspool/KDIALFILE/',Cells[C_B_HIDEFILE, ARow]) > 0 then
            sRemoteFile := Cells[C_B_HIDEFILE, ARow]
         else
            sRemoteFile := '/ftpspool/KDIALFILE/' + Cells[C_B_HIDEFILE, ARow];

         // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
         sLocalFile  := G_HOMEDIR + 'TEMP\SPOOL\' + Cells[C_B_ATTACH, ARow]; //lb_Filename.Caption;



         if (GetBINFTP(sServerIp, sFtpUserID, sFtpPasswd, sRemoteFile, sLocalFile, False)) then
         begin
            //		ShowMessage('정상적으로 저장되었습니다.');
         end;



         Screen.Cursor := crDefault;



         try
            if (PosByte('.exe', sLocalFile) > 0) or
               (PosByte('.zip', sLocalFile) > 0) or
               (PosByte('.rar', sLocalFile) > 0) then
            begin
               MessageBox(self.Handle,
                          PChar('첨부파일(' + sLocalFile + ') 다운로드가 완료되었습니다.' + #13#10 + #13#10 +
                                '※ 임시 다운로드 폴더 --> C:\KUMC(_DEV)\TEMP\SPOOL\'),
                          '첨부파일 다운로드 실행 완료',
                          MB_OK + MB_ICONINFORMATION);
            end
            else
            begin
               ShellExecute(HANDLE, 'open',
                            PCHAR(Cells[C_B_ATTACH, ARow]),
                            PCHAR(''),
                            // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
                            PCHAR(G_HOMEDIR + 'TEMP\SPOOL\'),
                            SW_SHOWNORMAL);
            end;

         except
            MessageBox(self.Handle,
                       PChar('해당 프로그램 실행중 오류가 발생하였습니다.' + #13#10 + #13#10 +
                             '프로그램 종료 후 다시 실행해 주시기 바랍니다.'),
                       '첨부파일 다운로드 실행 오류',
                       MB_OK + MB_ICONERROR);

            Exit;
         end;
      end;


      // Refresh
      if Trim(varUpResult) = 'Y' then
         SelGridInfo('BOARD');



   end;
end;


//------------------------------------------------------------------------------
// FlatSpeedButton onClick Event Handler
//
// Date   : 2013.08.20
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.fsbt_BoardRefreshClick(Sender: TObject);
begin
   SelGridInfo('BOARD');
end;


//------------------------------------------------------------------------------
// FlatComboBox onExit Event Handler
//
// Date   : 2013.08.21
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.fcb_CateUpExit(Sender: TObject);
begin
   // Title Init.
   fed_Title.Clear;


   if (PosByte('업무', fcb_CateUp.Text) > 0) or
      (PosByte('문의', fcb_CateUp.Text) > 0) then
   begin
      fed_CateDownNm.Visible := True;
      fcb_CateDown.Visible   := True;

      fcb_CateDown.SetFocus;
   end
   else
   begin
      fed_CateDownNm.Visible := False;
      fcb_CateDown.Visible   := False;
      fcb_CateDown.ItemIndex := -1;
   end;


   if (PosByte('공지', fcb_CateUp.Text) > 0) then
      {and (FsUserIp = '개발자IP') then}          // 공지사항 관리 누구나 할 수 있도록 주석 @ 2018.03.20 LSH
   begin
      fed_AlarmFro.Visible   := True;
      fed_AlarmTo.Visible    := True;
      fmed_AlarmFro.Visible  := True;
      fmed_AlarmTo.Visible   := True;

      fmed_AlarmFro.SetFocus;
   end
   else
   begin
      fed_AlarmFro.Visible   := False;
      fed_AlarmTo.Visible    := False;
      fmed_AlarmFro.Visible  := False;
      fmed_AlarmTo.Visible   := False;
      fmed_AlarmFro.Clear;
      fmed_AlarmTo.Clear;
   end;
end;



//------------------------------------------------------------------------------
// FlatComboBox onChange Event Handler
//
// Date   : 2013.08.21
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.fcb_PageChange(Sender: TObject);
begin

   if (StrToInt(fcb_Page.Text) <= fcb_Page.Items.Count) and
      (StrToInt(fcb_Page.Text) = 1) then
   begin
      fsbt_ForWard.Visible    := False;
      fsbt_BackWard.Visible   := True;
   end
   else if (StrToInt(fcb_Page.Text) <= fcb_Page.Items.Count) and
           (StrToInt(fcb_Page.Text) > 1) and
           (StrToInt(fcb_Page.Text) < fcb_Page.Items.Count) then
   begin
      fsbt_ForWard.Visible    := True;
      fsbt_BackWard.Visible   := True;
   end
   else if (StrToInt(fcb_Page.Text) <= fcb_Page.Items.Count) and
           (StrToInt(fcb_Page.Text) > 1) and
           (StrToInt(fcb_Page.Text) = fcb_Page.Items.Count) then
   begin
      fsbt_ForWard.Visible    := True;
      fsbt_BackWard.Visible   := False;
   end
   else if (StrToInt(fcb_Page.Text) < fcb_Page.Items.Count) and
           (StrToInt(fcb_Page.Text) < 1) then
   begin
      fsbt_ForWard.Visible    := True;
      fsbt_BackWard.Visible   := False;
   end
   else
   begin
      fsbt_ForWard.Visible    := False;
      fsbt_BackWard.Visible   := False;
   end;


   SelGridInfo('BOARD');

end;


//------------------------------------------------------------------------------
// FlatSpeedButton onClick Event Handler
//
// Date   : 2013.08.21
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.fsbt_ForWardClick(Sender: TObject);
var
   tmpPageIdx : Integer;
begin
   tmpPageIdx := StrToInt(fcb_Page.Text) - 1;
   fcb_Page.Text := IntToStr(tmpPageIdx);

   fcb_PageChange(Sender);
end;


//------------------------------------------------------------------------------
// FlatSpeedButton onClick Event Handler
//
// Date   : 2013.08.21
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.fsbt_BackWardClick(Sender: TObject);
var
   tmpPageIdx : Integer;
begin
   tmpPageIdx := StrToInt(fcb_Page.Text) + 1;
   fcb_Page.Text := IntToStr(tmpPageIdx);

   fcb_PageChange(Sender);
end;


//------------------------------------------------------------------------------
// AdvStringGrid onGetCellPrintColor Event Handler
//
// Date   : 2013.08.21
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_NetworkGetCellPrintColor(Sender: TObject; ARow,
  ACol: Integer; AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
var
   i : Integer;
begin
   if (ARow >= 0) and
      (ARow < asg_Network.Rowcount - asg_Network.Fixedfooters) then
   begin
      if (ARow = 0) then
         for i := 0 to asg_Network.ColCount- 1 do
         begin
            if (ACol = i) then
            begin
               ABrush.Color   := clSilver;
               AFont.Color    := clBlack;
            end;
         end;


      if (ARow > 0) then
         for i := 0 to asg_Network.FixedCols - 1 do
         begin
            if (ACol = i) then
            begin
               ABrush.Color   := clSilver;
               AFont.Color    := clBlack;
            end;
         end;
   end;

   if ARow = 0 then
   begin
      AFont.Style := [fsBold];
   end;

end;

//------------------------------------------------------------------------------
// AdvStringGrid onGetCellPrintColor Event Handler
//
// Date   : 2013.08.21
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_DetailGetCellPrintColor(Sender: TObject; ARow,
  ACol: Integer; AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
var
   i : Integer;
begin
   if (ARow >= 0) and
      (ARow < asg_Detail.Rowcount - asg_Detail.Fixedfooters) then
   begin
      if (ARow = 0) then
         for i := 0 to asg_Detail.ColCount- 1 do
         begin
            if (ACol = i) then
            begin
               ABrush.Color   := clSilver;
               AFont.Color    := clBlack;
            end;
         end;
   end;

   if ARow = 0 then
   begin
      AFont.Style := [fsBold];
   end;

end;


//------------------------------------------------------------------------------
// AdvStringGrid onGetCellPrintColor Event Handler
//
// Date   : 2013.08.22
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_MasterGetCellPrintColor(Sender: TObject; ARow,
  ACol: Integer; AState: TGridDrawState; ABrush: TBrush; AFont: TFont);
var
   i : Integer;
begin
   if (ARow >= 0) and
      (ARow < asg_Master.Rowcount - asg_Master.Fixedfooters) then
   begin
      if (ARow = 0) then
         for i := 0 to asg_Master.ColCount- 1 do
         begin
            if (ACol = i) then
            begin
               ABrush.Color   := clSilver;
               AFont.Color    := clBlack;
            end;
         end;
   end;

   if ARow = 0 then
   begin
      AFont.Style := [fsBold];
   end;

end;


//------------------------------------------------------------------------------
// FlatSpeedButton onClick Event Handler
//
// Date   : 2013.08.23
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.fsbt_VersionClick(Sender: TObject);
var
  About: TAbout;
begin
   if (IsLogonUser) then
   begin
      About := TAbout.Create(self);

      try
         // 프로그램 version 확인.
         About.lb_Version.Caption := Self.Caption;
         About.ShowModal;

      finally
         About.Free;
      end;
   end;
end;



//------------------------------------------------------------------------------
// FlatSpeedButton onClick Event Handler
//
// Date   : 2013.08.23
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.fsbt_SearchClick(Sender: TObject);
begin
   if not IsLogonUser then
   begin
      MessageBox(self.Handle,
                 PChar('[비상연락망] - [담당자정보]에 담당자 IP 등록후 사용하실 수 있습니다.'),
                 PChar(Self.Caption + ' : 등록되지 않은 User 제한 알림'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;


   fpn_Write.Top     := 0;
   fpn_Write.Left    := 0;
   fpn_Write.Visible := True;
   fpn_Write.BringToFront;

   fed_BoardSeq.Text    := '';

   lb_HeadTail.Caption  := '';
   lb_HeadSeq.Caption   := '';
   lb_TailSeq.Caption   := '';

   fcb_CateUp.Text      := '';
   fcb_CateUp.Enabled   := True;
   fed_Writer.Enabled   := True;
   fed_Title.Enabled    := True;
   fmm_Text.Enabled     := True;

   fed_Title.Text       := '';
   fed_Title.Hint       := '★ [제목] 이중키워드 검색 예시)' + #13#10 + #13#10 + '산정특례 (한칸 띄우고) 중증 (엔터)';
   fed_Writer.Text      := '';
   fmm_Text.Lines.Text  := '';
   fmm_Text.Hint        := '★ [내용] 이중키워드 검색 예시)' + #13#10 + #13#10 + '구로 (한칸 띄우고) 동의서 (입력후 Ctrl + 엔터시 바로검색 단축키)';
   fed_Attached.Text    := '';
   //fcb_CateUp.SetFocus;

   fcb_CateDown.Text      := '';
   fed_CateDownNm.Visible := False;
   fcb_CateDown.Visible   := False;

   fsbt_FileOpen.Enabled  := False;
   fsbt_FileClear.Enabled := False;

   fsbt_WriteReg.Visible  := False;
   fsbt_WriteFind.Visible := True;
   fsbt_WriteFind.Top     := 480;
   fsbt_WriteFind.Left    := 308;

   lb_Write.Caption    := '▶ 글 검색 모드입니다.';
   lb_Write.Font.Color := $00BD5E20;
   lb_Write.Font.Style := [fsBold];

   fed_Title.SetFocus;
end;


//------------------------------------------------------------------------------
// FlatSpeedButton onClick Event Handler
//
// Date   : 2013.08.23
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.fsbt_WriteFindClick(Sender: TObject);
var
   varResult : String;
   tmpCateUp : String;
begin

   SelGridInfo('BOARDFIND');


   if (fcb_CateUp.Text = '') then
      tmpCateUp := 'NULL'
   else
      tmpCateUp := fcb_CateUp.Text;


   // 로그 Update
   UpdateLog('BOARD',
             tmpCateUp,
             FsUserIp,
             'F',
             fed_Title.Text,
             fmm_Text.Lines.Text,
             varResult
             );


   fsbt_WriteCancelClick(Sender);

   // 최종 Focus는 Grid에 고정 (바로 단축키 Ctrl+F 사용가능 하도록) @ 2016.11.18 LSH
   asg_Board.SetFocus;
end;

procedure TMainDlg.fmm_TextExit(Sender: TObject);
begin
   if (lb_HeadTail.Caption = 'H') then
   begin
      if (fmm_Text.Lines.Count > 15) then
      begin
         MessageBox(self.Handle,
                    PChar('신규 글작성은 15줄 이내로 제한합니다.'),
                    PChar(Self.Caption + ' : 커뮤니티 글 작성 제한 알림'),
                    MB_OK + MB_ICONWARNING);

         fmm_Text.SetFocus;
      end;
   end
   else if (lb_HeadTail.Caption = 'T') then
   begin
      if (fmm_Text.Lines.Count > 5) then
      begin
         MessageBox(self.Handle,
                    PChar('댓글작성은 스크롤의 압박을 방지하기위해 5줄 이내로 제한합니다.'),
                    PChar(Self.Caption + ' : 커뮤니티 댓글 작성 제한 알림'),
                    MB_OK + MB_ICONWARNING);

         fmm_Text.SetFocus;
      end;
   end;

   // Event 발생시, 갑작스런 [응답없음] 문제해결위한 처리
   Application.ProcessMessages;

end;



//------------------------------------------------------------------------------
// 한영키를 한글키로 변환
//       - 소스출처 : MComFunc.pas
//
// Author : Lee, Se-Ha
// Date   : 2013.08.22
//------------------------------------------------------------------------------
procedure TMainDlg.HanKeyChg(Handle1:THandle);
var
   tMode : HIMC;
begin
   tMode := ImmGetContext(Handle1);
   ImmSetConversionStatus(tMode, IME_CMODE_HANGUL,
                                 IME_CMODE_HANGUL);
end;

procedure TMainDlg.fed_WriterKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   HanKeyChg(fed_Writer.Handle);

   if fsbt_WriteFind.Visible then
   begin
      // 글 검색 모드에서, [작성자] 란에 키워드 입력후 <Enter> 누르면 자동 검색 실시 @ 2016.11.18 LSH
      if Key = VK_RETURN then
         fsbt_WriteFindClick(Sender);
   end;
end;




//------------------------------------------------------------------------------
// [조회] 알람 팝업 메세지 정보
//       - 게시판 유효기간내 [공지사항] 조회
//
// Author : Lee, Se-Ha
// Date   : 2013.08.22
//------------------------------------------------------------------------------
function TMainDlg.GetAlarmPopup(in_Gubun, in_CheckDate : String) : String;
var
   TpGetAlarm : TTpSvc;
   sType1, tmpInfo : String;
   i : Integer;
begin


   //-----------------------------------------------------------------
   // 2. 조회
   //-----------------------------------------------------------------
   Screen.Cursor := crHourGlass;



   //
   sType1   := '13';
   tmpInfo  := '';



   Result   := tmpInfo;



   //-----------------------------------------------------------------
   // 2-1. Load Data by TpSvc
   //-----------------------------------------------------------------
   TpGetAlarm := TTpSvc.Create;
   TpGetAlarm.Init(Self);




   Screen.Cursor := crHourglass;





   try
      TpGetAlarm.CountField  := 'S_CODE9';
      TpGetAlarm.ShowMsgFlag := False;


      if TpGetAlarm.GetSvc('MD_KUMCM_S1',
                          ['S_TYPE1'  , sType1
                          ],
                          [
                           'S_CODE9'  , 'sUserIp'
                         , 'S_CODE21' , 'sDocTitle'
                         , 'S_CODE22' , 'sRegDate'
                         , 'S_STRING1', 'sBoardSeq'
                         , 'S_STRING2', 'sCateUp'
                         , 'S_STRING4', 'sContext'
                          ]) then

         if TpGetAlarm.RowCount < 0 then
         begin
            ShowMessage(GetTxMsg);
            Exit;
         end
         else if TpGetAlarm.RowCount = 0 then
         begin
            Exit;
         end;



         //----------------------------------------------------
         // 알람 정보  Return
         //----------------------------------------------------
         tmpInfo := '▶ ' + IntToStr(TpGetAlarm.RowCount) + '건의 알림 ◀' + #13#10;


         for i := 0 to TpGetAlarm.RowCount - 1 do
            tmpInfo := tmpInfo + #13#10 + '[' + TpGetAlarm.GetOutputDataS('sCateUp', i) + '] ' //+ TpGetAlarm.GetOutputDataS('sBoardSeq', i) + ' '
                                                                                                  + TpGetAlarm.GetOutputDataS('sDocTitle', i);


         Result := tmpInfo;



   finally
      FreeAndNil(TpGetAlarm);
      Screen.Cursor := crDefault;
   end;
end;



//------------------------------------------------------------------------------
// FlatComboBox onExit Event Handler
//
// Date   : 2013.08.23
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.fcb_CateDownExit(Sender: TObject);
begin
   fed_Title.Text := '[' + fcb_CateDown.Text + ']' + fed_Title.Text;
end;


//------------------------------------------------------------------------------
// AdvStringGrid onMouseMove Event Handler
//
// Date   : 2013.08.23
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_BoardMouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
var
   NowCol, NowRow : Integer;
begin
   with asg_Board do
   begin
      // 현재 Cell 좌표 가져오기
      MouseToCell(X, Y, NowCol, NowRow);


      if (NowRow > 0) and
         (
            (
               (NowCol = C_B_TITLE) and
               (Cells[C_B_HEADTAIL, NowRow] = 'H')       // 게시판 제목 @ 2016.11.18 LSH
            ) or
            (NowCol = C_B_ATTACH)                        // 첨부 파일 명칭만 @ 2016.11.18 LSH
         ) then
      begin
         ShowHint := True;
      end
      else
         ShowHint := False;
   end;
end;



//------------------------------------------------------------------------------
// FlatSpeedButton onClick Event Handler
//
// Date   : 2013.08.27
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.fsbt_FileOpenClick(Sender: TObject);
begin
   if od_File.Execute then
      fed_Attached.Text := od_File.FileName;
end;



//------------------------------------------------------------------------------
// FlatSpeedButton onClick Event Handler
//
// Date   : 2013.08.27
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.fsbt_FileClearClick(Sender: TObject);
begin
   fed_Attached.Clear;
end;



//------------------------------------------------------------------------------
// [FTP] File Upload function Event Handler
//
// Date   : 2013.08.23
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
function TMainDlg.FileUpLoad(TargetFile, DelFile: String; var S_IP: String):Boolean;
var
   sServerIp, sFtpUserId, sFtpPasswd, sFtpHostName, sFtpDir: String;
begin

   Result := True;


   // 파일 업로드를 위한 정보 조회
   if not GetBinUploadInfo(sServerIp, sFtpUserID, sFtpPasswd, sFtpHostName, sFtpDir) then
   begin
      ShowMessage('파일 저장을 위한 서버정보 조회중, 오류가 발생했습니다.');
      Result := False;
      exit;
   end;

   try
      if Id_FTP.Connected then
         Id_FTP.Disconnect;
   except
      Showmessage('초기설정 에러');
      Result := False;
      Exit;
   end;


   Id_FTP.Host       := sServerIp;
   Id_FTP.UserName   := sFtpUserId;
   Id_FTP.Password   := sFtpPasswd;


   try
      Id_FTP.Connect;

   except
      Showmessage('FTP 연결에러');
      Result := False;
      Exit;
   end;


   try
      // Chat 박스 파일 전송모드이면, Path 변경
      if (IsFileSend) then
         Id_FTP.ChangeDir('/ftpspool/CHATFILE')
      else
         Id_FTP.ChangeDir('/ftpspool/KDIALFILE')

   except
      Showmessage('ChangeDir 에러');
      Result := False;
      Exit;
   end;


   try
      Id_FTP.TransferType := ftBinary;

   except
      Showmessage('Mode 변경 에러');
      Result := False;
      Exit;
   end;


   {  -- 첨부내용 수정기능 미지원.. 주석처리 @ 2013.08.28 LSH
   if DelFile <> '' then
   begin
      if Trim(fed_Attached.Text) <> DelFile then
      begin
         try
            Id_FTP.Delete(DelFile);

         except
            Showmessage('파일삭제 에러');
            Result := False;
            Exit;
         end;
      end;
   end;
   }


   //--------------------------------------------------------
   // Chat 박스 파일 전송모드 분기
   //--------------------------------------------------------
   if (IsFileSend) then
   begin

      if Trim(asg_ChatSend.Cells[C_SD_TEXT, 0]) <> '' then
      begin
         try
            Id_FTP.Put(Trim(asg_ChatSend.Cells[C_SD_TEXT, 0]), TargetFile);

         except
            Showmessage('파일전송에러');
            Result := False;
            Exit;
         end;

         try
            Id_FTP.SendCmd('SITE CHMOD 644 ' + TargetFile)

         except
            Showmessage('SITE CHMOD Error');
            Result := False;
            Exit;
         end;
      end;
   end
   else
   begin
      // 커뮤니티 [첨부]
      if Trim(fed_Attached.Text) <> '' then
      begin
         try
            Id_FTP.Put(Trim(fed_Attached.Text), TargetFile);

         except
            Showmessage('[커뮤니티] 파일전송에러');
            Result := False;
            Exit;
         end;

         try
            Id_FTP.SendCmd('SITE CHMOD 644 ' + TargetFile)

         except
            Showmessage('SITE CHMOD Error');
            Result := False;
            Exit;
         end;
      end
      else
      // 비상연락망 프로필 이미지 [첨부]
      begin
         try
            Id_FTP.Put(DelFile, TargetFile);

         except
            Showmessage('[기타] 파일전송에러');
            Result := False;
            Exit;
         end;

         try
            Id_FTP.SendCmd('SITE CHMOD 644 ' + TargetFile)

         except
            Showmessage('SITE CHMOD Error');
            Result := False;
            Exit;
         end;
      end;


   end;



   try
      Id_FTP.Disconnect;

   except
      Showmessage('객체해제 에러');
      Result := False;
      Exit;
   end;


   S_IP := sServerIp;


end;



//------------------------------------------------------------------------------
// AdvStringGrid onEditingDone Event Handler
//
// Date   : 2013.08.27
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_MasterEditingDone(Sender: TObject);
begin
   with asg_Master do
   begin
      if (Cells[C_M_LOCATE,   Row] <> '') and
         (Cells[C_M_USERNM,   Row] <> '') and
         (Cells[C_M_DUTYPART, Row] <> '') then
      begin
         AddButton(C_M_BUTTON,  Row, ColWidths[C_M_BUTTON]-5,  20, 'Update', haBeforeText, vaCenter);           // Update
      end;
   end;
end;


//------------------------------------------------------------------------------
// MenuItem onClick Event Handler
//
// Date   : 2013.08.27
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.mi_NickNmClick(Sender: TObject);
begin
   apn_Profile.Visible := True;

   lb_UserNm.Caption := asg_Master.Cells[C_M_USERNM, asg_Master.Row];
   fed_NickNm.Text   := asg_Master.Cells[C_M_NICKNM, asg_Master.Row];
end;




//------------------------------------------------------------------------------
// FlatSpeedButton onClick Event Handler
//
// Date   : 2013.08.27
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.fsbt_Profile2GridClick(Sender: TObject);
begin
   asg_Master.Cells[C_M_NICKNM, asg_Master.Row] := Trim(fed_NickNm.Text);

   apn_Profile.Visible := False;
end;




//------------------------------------------------------------------------------
// [조회] 게시판 (커뮤니티) 추천 List 상세조회
//
// Author : Lee, Se-Ha
// Date   : 2013.08.27
//------------------------------------------------------------------------------
procedure TMainDlg.GetLikeList(in_HeaderRow : Integer;
                               in_BoardSeq,
                               in_UserIp,
                               in_RegDate : String);
var
   TpGetLike  : TTpSvc;
   sType, tmp_UsrPhotoInfo  : String;
   i, iRowCnt  : Integer;
begin


   sType := '14';



   //-----------------------------------------------------------------
   // 2-1. Load Data by TpSvc
   //-----------------------------------------------------------------
   TpGetLike := TTpSvc.Create;
   TpGetLike.Init(Self);


   Screen.Cursor := crHourglass;


   try
      TpGetLike.CountField  := 'S_CODE9';
      TpGetLike.ShowMsgFlag := False;

      if TpGetLike.GetSvc('MD_KUMCM_S1',
                          ['S_TYPE1'  , sType
                         , 'S_TYPE2'  , in_BoardSeq
                         , 'S_TYPE3'  , in_UserIp
                         , 'S_TYPE4'  , in_RegDate
                          ],
                          [
                           'S_CODE2'   , 'sUserId'
                         , 'S_CODE9'   , 'sUserIp'
                         , 'S_STRING5' , 'sAttachNm'
                         , 'S_STRING6' , 'sHideFile'
                         , 'S_STRING14', 'sNickNm'
                          ]) then

         if TpGetLike.RowCount <= 0 then
         begin
            Exit;
         end;




      with flbx_Like do
      begin
         Items.Clear;

         iRowCnt  := TpGetLike.RowCount;


         //---------------------------------------------------------------------
         // 추천인 Photo - AdvStringGrid 연동으로 전환 @ 2015.04.09 LSH
         //---------------------------------------------------------------------
         asg_Like.RowCount := 2;
         asg_Like.ClearRows(1, asg_Like.RowCount);

         asg_Like.ColCount := iRowCnt;


         asg_Like.RowHeights[0] := 45;
         asg_Like.RowHeights[1] := 0;


         for i := 0 to iRowCnt - 1 do
         begin
            asg_Like.RemovePicture(i, 0);




            if TpGetLike.GetOutputDataS('sNickNm', i) <> '' then
               asg_Like.Cells[i, 1] := TpGetLike.GetOutputDataS('sNickNm', i)
            else
               asg_Like.Cells[i, 1] := TpGetLike.GetOutputDataS('sUserIp', i);



            // 추천 User 이미지 표기 적용
            tmp_UsrPhotoInfo :=  GetFTPImage(DeleteStr(TpGetLike.GetOutputDataS('sAttachNm', i), '/media/cq/photo/'),
                                             TpGetLike.GetOutputDataS('sHideFile', i),
                                             TpGetLike.GetOutputDataS('sUserId',   i));

            if (tmp_UsrPhotoInfo <> '') then
            begin
               // 수신한 Image 파일을 Grid에 표기
               asg_Like.CreatePicture(i,
                                      0,
                                      False,
                                      Stretch, // StretchWithAspectRatio, ShrinkWithAspectRatio
                                      0,
                                      haCenter,
                                      vaCenter).LoadFromFile(tmp_UsrPhotoInfo);
            end
            else
            begin
               if TpGetLike.GetOutputDataS('sNickNm', i) <> '' then
                  asg_Like.Cells[i, 0] := TpGetLike.GetOutputDataS('sNickNm', i)
               else
                  asg_Like.Cells[i, 0] := 'unknown user';
            end;
         end;


         {  -- 기존 Flat-ListBox String 추가부분 주석 @ 2015.04.09 LSH
         for i := 0 to iRowCnt - 1 do
         begin
            if TpGetLike.GetOutputDataS('sNickNm', i) = '' then
               Items.Add(TpGetLike.GetOutputDataS('sUserIp', i))
            else
               Items.Add(TpGetLike.GetOutputDataS('sNickNm', i));
         end;
         }

      end;


      // Comments
      apn_LikeList.Caption.Text := IntToStr(iRowCnt) + '명이 추천합니다';


      // Location & Display
      if (asg_Board.Height - asg_Board.CellRect(C_B_LIKE, asg_Board.Row).Top) > (apn_LikeList.Height + 10) then
      begin
         apn_LikeList.Top     := asg_Board.CellRect(C_B_LIKE, asg_Board.Row).Top + 66;
         apn_LikeList.Left    := 444;
      end
      else
      begin
         apn_LikeList.Top     := asg_Board.CellRect(C_B_LIKE, asg_Board.Row).Top - apn_LikeList.Height + 40;
         apn_LikeList.Left    := 444;
      end;

      apn_LikeList.Collaps := True;
      apn_LikeList.Visible := True;
      apn_LikeList.Collaps := False;


   finally
      FreeAndNil(TpGetLike);
      Screen.Cursor := crDefault;
   end;
end;





//------------------------------------------------------------------------------
// AdvStringGrid onClickCell Event Handler
//
// Date   : 2013.08.28
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_BoardClickCell(Sender: TObject; ARow,
  ACol: Integer);
begin
   if apn_LikeList.Visible then
   begin
      apn_LikeList.Collaps := True;
      apn_LikeList.Visible := False;
   end;

   with asg_Board do
   begin
      if (ACol = C_B_LIKE) then
      begin
         if (Cells[C_B_BOARDSEQ, ARow] <> '┗') and
            (Cells[C_B_BOARDSEQ, ARow] <> '') then

            GetLikeList(asg_Board.Row,
                        asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row],
                        asg_Board.Cells[C_B_USERIP,   asg_Board.Row],
                        asg_Board.Cells[C_B_REGDATE,  asg_Board.Row])

         else if Cells[C_B_CATEUP, ARow] = '┗' then

            GetLikeList(asg_Board.Row,
                        asg_Board.Cells[C_B_CONTEXT, asg_Board.Row],
                        asg_Board.Cells[C_B_USERIP,  asg_Board.Row],
                        asg_Board.Cells[C_B_REGDATE, asg_Board.Row]);

      end;
   end;
end;




//------------------------------------------------------------------------------
// [조회] 동일 IP 다중 User 선택 정보 가져오기
//       - DlgUser.pas에서 Callback 용으로 사용중
//
// Date   : 2013.08.29
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.GetMultiUser(in_UserNm,
                                in_NickNm,
                                in_UserPart,
                                in_UserSpec,
                                in_Mobile,
                                in_CallNo,
                                in_MngrNm : String);
begin
   FsUserNm       := in_UserNm;
   FsNickNm       := in_NickNm;
   FsUserPart     := in_UserPart;
   FsUserSpec     := in_UserSpec;
   FsUserMobile   := in_Mobile;
   FsUserCallNo   := in_CallNo;
   FsMngrNm       := in_MngrNm;   // 문서 관리자 (팀장) @ 2014.07.18 LSH
end;





//------------------------------------------------------------------------------
// [채팅] 채팅 Client 호출
//
// Date   : 2013.08.29
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.mi_ChatClick(Sender: TObject);
var
   //SetClient : TClient;
//   varResult : String;
   FForm     : TForm;
   tmp_ChatUserIp,
   tmp_ChatUserNm : String;

begin
   if not IsLogonUser then
   begin
      MessageBox(self.Handle,
                 PChar('[비상연락망] - [담당자정보]에 담당자 IP 등록후 사용하실 수 있습니다.'),
                 PChar(Self.Caption + ' : 등록되지 않은 User 제한 알림'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;


   if (apn_Master.Visible) then
   begin
      tmp_ChatUserIp := asg_Master.Cells[C_M_USERIP, asg_Master.Row];
      tmp_ChatUserNm := asg_Master.Cells[C_M_USERNM, asg_Master.Row]
   end
   else if (apn_Network.Visible) then
   begin
      tmp_ChatUserIp := asg_NetWork.Cells[C_NW_USERIP,  asg_NetWork.Row];
      tmp_ChatUserNm := asg_NetWork.Cells[C_NW_USERNM,  asg_NetWork.Row];
   end;


   {  -- Chat-Client에서 Log 업데이트 하므로 주석 @ 2015.03.30 LSH
   // 로그 Update
   UpdateLog('CBLOG',
             tmp_ChatUserIp,
             FsUserIp,
             'R',
             '',
             '',
             varResult
             );
   }


   {  -- Chat - Client (BPL) 적용에 따른 주석처리 @ 2015.03.30 LSH
   SetClient := TClient.Create(self);


   try
      // 접속대상 IP와 닉네임 설정
      SetClient.EditIp.Text   := asg_Master.Cells[C_M_USERIP, asg_Master.Row];
      SetClient.EditNick.Text := FsUserNm;

      // 접속 서버의 IP 및 Port 정보
      SetClient.ClientSocket.Host    := SetClient.EditIp.Text;
      SetClient.ClientSocket.Port    := StrToInt(DelChar(CopyByte(asg_Master.Cells[C_M_USERIP, asg_Master.Row], LengthByte(asg_Master.Cells[C_M_USERIP, asg_Master.Row])-3, 4), '.'));


      //---------------------------------------------------------
      // 접속하려는 Server Port 사용가능 여부 Check
      //---------------------------------------------------------
      if PortTCPIsOpen(SetClient.ClientSocket.Port, SetClient.ClientSocket.Host) then
      begin
         SetClient.ClientSocket.Active  := True;
         SetClient.ShowModal;
      end
      else
      begin
         MessageBox(self.Handle,
                    PChar('접속하려는 대상 Port가 닫혀있습니다.' + #13#10 + #13#10 + '접속 대상자의 프로그램이 실행중인지 확인이 필요합니다.'),
                    '챗(Chat) 박스 :: 접속 서버 Port 닫힘',
                    MB_OK + MB_ICONWARNING);

         SetClient.ClientSocket.Active  := False;
      end;



   finally
      SetClient.Free;
   end;
   }


   // 내 사진 (이미지) 경로 정보
   if Trim(asg_NetWork.Cells[C_NW_PHOTOFILE, iUserRowId]) = '' then
   begin
      FsChatMyPhoto := asg_NetWork.Cells[C_NW_USERID, iUserRowId] + '.jpg';
   end
   else
      // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
      FsChatMyPhoto  := G_HOMEDIR + 'TEMP\SPOOL\' + asg_Network.Cells[C_NW_PHOTOFILE, iUserRowId];



   //-------------------------------------------------------
   // 다이얼 Chat - Client (BPL) Form 호출
   //-------------------------------------------------------
   FForm := BplFormCreate('CHATCLIENT');

   try
      Screen.Cursor := crHourGlass;

      if FForm <> nil then
      begin
         SetBplStrProp(FForm, 'prop_EditIp',       FsUserIp);
         SetBplStrProp(FForm, 'prop_EditNick',     FsUserNm);
         SetBplStrProp(FForm, 'prop_ChatIp',       tmp_ChatUserIp);
         SetBplStrProp(FForm, 'prop_ChatUserNm',   tmp_ChatUserNm);
         // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
         SetBplStrProp(FForm, 'prop_ChatPhoto' ,   G_HOMEDIR + 'TEMP\SPOOL\' + asg_Network.Cells[C_NW_PHOTOFILE, asg_Network.Row]);
         SetBplStrProp(FForm, 'prop_MyPhoto'   ,   FsChatMyPhoto);

         FForm.Show;
      end;
   finally
      Screen.Cursor := crDefault;
   end;
end;



{
procedure TMainDlg.tm_ChatTimer(Sender: TObject);
var
   SetServer : TServer;
   GetClient : TClient;
begin
   //showmessage('tm_ChatTimer Call');
   //Memo2.Lines.Add('--- tm_ChatTimer Call ---');

   try

      GetClient := TClient.Create(Self);

      if (GetClient.ClientSocket.Host <> '') and
         (GetClient.ClientSocket.Active) then
      begin

         Memo2.Lines.Add('--- ClientSocket Host Call !! ---');

         try
            SetServer := TServer.Create(self);

            SetServer.ServerSocket.Active := True;
            SetServer.ShowModal;

         finally
            SetServer.Free;
         end;

         //showmessage('Client 호스트 요청중 --> 서버화면 call');

      end

      else
         //SetServer.ServerSocket.Active := False;
         // test
         Memo2.Lines.Add('--- ClientSocket Host is null ---');



   finally
      GetClient.Free;
   end;



end;
}



//------------------------------------------------------------------------------
// String 내의 특정 문자 삭제
//       - 소스 출처 : MComFunc.pas
//
// Date   : 2013.08.29
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
function TMainDlg.DelChar( const Str : String ; DelC : Char ) : String;
var
  I, L : Integer;
begin
  L := LengthByte(Str);

  result := '';

    for I := 1 to L  do begin
      if DelC <> '' then
      begin
        if Str[I] <> DelC then
          result := result + Str[I];
      end
      else
      begin
        if Str[I] in ['0'..'9'] then
          result := result + Str[I];
      end;
    end;
end;



//------------------------------------------------------------------------------
// [Chat] Client로 Message Echo 전송
//
// Date   : 2013.08.29
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
function TMainDlg.SendToAllClients( s : string) : boolean;
var
   i : integer;
begin
   if ServerSocket.Socket.ActiveConnections = 0 then
   begin
      MessageBox(self.Handle,
                 PChar('메세지를 보낼 수 있게 연결된 User가 현재 없습니다.'),
                 '챗(Chat) :: 현재 접속중인 User 없음',
                 MB_OK + MB_ICONWARNING);

      Result := False;

      Exit;
   end;

   for i := 0 to (ServerSocket.Socket.ActiveConnections - 1) do
   begin
      //------------------------------------------------------------
      // 현재 접속(대화)중인 Client에게만 Send 메세지 처리
      //------------------------------------------------------------
      if PosByte(ServerSocket.Socket.Connections[i].RemoteAddress, apn_ChatBox.Caption.Text) > 0 then
         ServerSocket.Socket.Connections[i].SendText(AnsiString(s));
   end;

   Result := True;

end;





//------------------------------------------------------------------------------
// [Chat] Server Socket - Client 연결 Event
//
// Date   : 2013.08.29
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.ServerSocketClientConnect(Sender: TObject;
  Socket: TCustomWinSocket);
//var
//   i : Integer;
begin


   SendToAllClients('<'+ Socket.RemoteHost + '(' + socket.RemoteAddress + ')' +' just connected.>');
   SendToAllClients('★[대화중]에 다른 User로부터 추가적인 대화요청이 올 수 있지만, 전송하는 메세지는 현재 [대화중]인 User에게만 전송됩니다★');


   if PosByte('대화중', apn_ChatBox.Caption.Text) = 0 then
      apn_ChatBox.Caption.Text := '챗(Chat) :: ' + Socket.RemoteHost + '(' + socket.RemoteAddress + ')' + ' 대화중';



   {
   asg_Chat.InsertRows(iChatRowCnt, 1);
   //asg_Chat.MergeCells(C_CH_MYCOL, iChatRowCnt, 3, 1);
   asg_Chat.Cells[C_CH_TEXT, iChatRowCnt] := '<' + Socket.RemoteHost  + ' just connected.>';
   asg_Chat.Row := iChatRowCnt;
   asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt] := taCenter;
   asg_Chat.AutoSizeRow(iChatRowCnt);

   Inc(iChatRowCnt);
   }

end;





//------------------------------------------------------------------------------
// [Chat] Server Socket - Client 연결끊김 Event
//
// Date   : 2013.08.29
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.ServerSocketClientDisconnect(Sender: TObject;
  Socket: TCustomWinSocket);
//var
//   i : Integer;
begin
   try
      SendToAllClients('<' + Socket.RemoteHost + '(' + socket.RemoteAddress + ')'+' just disconnected.>');


      if PosByte(Socket.RemoteAddress, apn_ChatBox.Caption.Text) > 0 then
      begin
         apn_ChatBox.Caption.Text      := '챗(Chat) :: ' + Socket.RemoteHost + '(' + socket.RemoteAddress + ')' + ' 대화종료';
         apn_ChatBox.Collaps           := True;
         apn_ChatBox.Top               := 540;
         apn_ChatBox.Left              := 304;
         apn_ChatBox.Caption.Color     := $000066C9;
         apn_ChatBox.Caption.ColorTo   := clMaroon;
      end;



      { ------------> onClientError 이벤트에서 서버소켓 Off->On 처리 진행함 !!
      try
         ServerSocket.Active := False;

      except
         on E : Exception do
         ShowMessage(E.ClassName+' error raised by ServerSocket.active off, with message : ' + E.Message);
      end;
      }
      {
      try
         ServerSocket.Port   := StrToInt(DelChar(CopyByte(FsUserIp, LengthByte(FsUserIp)-3, 4), '.'));
         ServerSocket.Active := True;
      except
         on E : Exception do
         ShowMessage(E.ClassName+' error raised by ServerSocket.active on, with message : ' + E.Message);
      end;
      }




   except
      on E : Exception do
      ShowMessage(E.ClassName+' error raised by ServerSocket, with message : ' + E.Message);
   end;



   {
   asg_Chat.InsertRows(iChatRowCnt, 1);
   //asg_Chat.MergeCells(C_CH_MYCOL, iChatRowCnt, 3, 1);
   asg_Chat.Cells[C_CH_TEXT, iChatRowCnt] := '<' + Socket.RemoteHost + ' just disconnected.>';
   asg_Chat.Row := iChatRowCnt;
   asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt] := taCenter;
   asg_Chat.AutoSizeRow(iChatRowCnt);

   Inc(iChatRowCnt);
   }

end;





//------------------------------------------------------------------------------
// [Chat] Server Socket - Client 메세지 Read Event
//
// Date   : 2013.08.29
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.ServerSocketClientRead(Sender: TObject;
  Socket: TCustomWinSocket);
var
   sl, sReceived, tmpOrgReceived : string;
//   Buffer: Array[0..10000] of byte;
//   R : TRect;
   hThreadID: Cardinal;
//   FocusWnd: THandle;
begin

   //-----------------------------------------------------------------
   // Window API : 창을 비활성화 --> 활성화 @ 2014.02.28 LSH
   //     - 델마당 참조
   //-----------------------------------------------------------------
   if IsWindowVisible(Self.Handle) then
   begin
      hThreadID:= GetWindowThreadProcessId((GetForegroundWindow), nil);
      AttachThreadInput(GetCurrentThreadID, hThreadID, True);
      SetForegroundWindow(Self.Handle);
      AttachThreadInput(GetCurrentThreadID, hThreadID, False);
   end
   else
   begin
      Application.Restore;
   end;


   apn_ChatBox.Top      := 296;
   apn_ChatBox.Left     := 304;
   apn_ChatBox.Collaps  := False;

   apn_ChatBox.Caption.Color     := $0089B354;
   apn_ChatBox.Caption.ColorTo   := $005ED7A9;


   sReceived := '';
   sReceived := String(Socket.ReceiveText);


   Inc(iChatRowCnt);
   asg_Chat.RowCount := iChatRowCnt;   // Grid Index out of Range 방지위해 반드시 필요 !


   //-----------------------------------------------------------
   // 상태 알람 Text 수신모드
   //-----------------------------------------------------------
   if (PosByte('just connected',     sReceived) > 0) or
      (PosByte('just disconnected',  sReceived) > 0) or
      (PosByte('[대화중]',           sReceived) > 0) then
   begin
      asg_Chat.InsertRows(iChatRowCnt, 1);
      asg_Chat.Cells[C_CH_TEXT,      iChatRowCnt] := sReceived;
      asg_Chat.Row                                := iChatRowCnt;
      asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt] := taCenter;
      asg_Chat.FontColors[C_CH_TEXT, iChatRowCnt] := $009589CC;
      asg_Chat.FontStyles[C_CH_TEXT, iChatRowCnt] := [fsBold];
      asg_Chat.AutoSizeRow(iChatRowCnt);
   end
   //-----------------------------------------------------------
   // 이모티콘(BMP 이미지) 수신모드
   //-----------------------------------------------------------
   else if (PosByte('[E]',  sReceived) > 0) then
   begin

      // 원본 수신내역 Temp 파일에 저장 (Client User 이름 추출위함)
      tmpOrgReceived := sReceived;

      DataSize := 0;
      sl       := '';


      if not Reciving then
      begin
         // Now we need to get the LengthByte of the data stream.
         SetLength(sl, StrLen(PChar(sReceived)) +1);           // +1 for the null terminator




         StrLCopy(@sl[1], PChar(sReceived), {PosByte('[', PChar(sReceived))-1)}LengthByte(sl)- 13);   // 이름태그(8) + &(1) + [E] (3) + 기존 Null (1)




         DataSize := StrToInt(sl);

         Data     := TMemoryStream.Create;
         Data.Size := 0;

         // DeleteByte the size information from the data.
         DeleteByte(sReceived, 1, LengthByte(sl));

         Reciving:= True;
      end;




      // Store the data to the file, until we've received all the data.
      try
         Data.Write(sReceived[1], LengthByte(sReceived));

         if Data.Size = DataSize then
         begin

            Data.Position:= 0;



            try
               // 1줄 추가
               asg_Chat.InsertRows(iChatRowCnt, 1);

            except
               on e : Exception do
               showmessage('Server Read Emoti ERROR(InsertRows) : ' + e.Message);
            end;



            try
               // 수신한 Data-Stream을 Grid에 표기
               asg_Chat.CreatePicture(C_CH_TEXT, iChatRowCnt, True, ShrinkWithAspectRatio, 0, haRight, vaTop).Bitmap.LoadFromStream(Data);

            except
               asg_Chat.Cells[C_CH_TEXT, iChatRowCnt] := '[이미지 수신오류]';
            end;

            // Data 송신자 표기
            //asg_Chat.Cells[C_CH_YOURCOL,   iChatRowCnt] := CopyByte(tmpOrgReceived, PosByte('&', tmpOrgReceived) + 1, 8);


            // Chat 상대 프로필 이미지 있으면, 사진 표기
            if FileExists(CopyByte(FindStr(tmpOrgReceived, '%'), PosByte('&', FindStr(tmpOrgReceived, '%')) + 2, LengthByte(FindStr(tmpOrgReceived, '%')) - PosByte('&', FindStr(tmpOrgReceived, '%')) - 2)) then
            begin
               // Chat 상대로부터 받은 나의 프로필 이미지 저장.
               FsChatMyPhoto  := CopyByte(tmpOrgReceived, PosByte('%', tmpOrgReceived) + 2, LengthByte(tmpOrgReceived) - PosByte('%', tmpOrgReceived) - 2);

               try
                  asg_Chat.CreatePicture(C_CH_YOURCOL,
                                         iChatRowCnt,
                                         False,
                                         ShrinkWithAspectRatio,
                                         0,
                                         haLeft,
                                         vaTop).LoadFromFile(CopyByte(FindStr(tmpOrgReceived, '%'), PosByte('&', FindStr(tmpOrgReceived, '%')) + 2, LengthByte(FindStr(tmpOrgReceived, '%')) - PosByte('&', FindStr(tmpOrgReceived, '%')) - 2));

               except
                  on e : Exception do
                     showmessage('Create Picture ERROR(ServerSocketClientRead) : ' + e.Message);
               end;

               asg_Chat.RowHeights[iChatRowCnt] := 40;

            end
            else
               asg_Chat.Cells[C_CH_YOURCOL,  iChatRowCnt] := CopyByte(FindStr(tmpOrgReceived, '%'), PosByte('&', FindStr(tmpOrgReceived, '%')) + 1, LengthByte(FindStr(tmpOrgReceived, '%')) - PosByte('%', FindStr(tmpOrgReceived, '%')) - 1);


            {  -- 기존 Text User 명 수정하는 부분 주석 @ 2015.03.31 LSH
            if FileExists(CopyByte(FindStr(tmpOrgReceived, '%'), PosByte('&', FindStr(tmpOrgReceived, '%')) + 2, LengthByte(FindStr(tmpOrgReceived, '%')) - PosByte('%', FindStr(tmpOrgReceived, '%')) - 2)) then
            begin

               FsChatMyPhoto  := CopyByte(FindStr(tmpOrgReceived, '%'), PosByte('&', FindStr(tmpOrgReceived, '%')) + 2, LengthByte(FindStr(tmpOrgReceived, '%')) - PosByte('%', FindStr(tmpOrgReceived, '%')) - 2);

               try
                  asg_Chat.CreatePicture(C_CH_YOURCOL,
                                         iChatRowCnt,
                                         False,
                                         ShrinkWithAspectRatio,
                                         0,
                                         haLeft,
                                         vaTop).LoadFromFile(CopyByte(FindStr(tmpOrgReceived, '%'), PosByte('&', FindStr(tmpOrgReceived, '%')) + 2, LengthByte(FindStr(tmpOrgReceived, '%')) - PosByte('%', FindStr(tmpOrgReceived, '%')) - 2));

               except
                  on e : Exception do
                     showmessage('Create Picture ERROR(ServerSocketClientRead) : ' + e.Message);
               end;

               asg_Chat.RowHeights[iChatRowCnt] := 40;

            end
            else
               asg_Chat.Cells[C_CH_YOURCOL,  iChatRowCnt] := CopyByte(FindStr(tmpOrgReceived, '%'), PosByte('&', FindStr(tmpOrgReceived, '%')) + 1, LengthByte(FindStr(tmpOrgReceived, '%')) - PosByte('%', FindStr(tmpOrgReceived, '%')));
            }


            // Scroll 조절
            asg_Chat.Row := iChatRowCnt;

            // 송신자 컬럼 Align
            asg_Chat.Alignments[C_CH_YOURCOL, iChatRowCnt] := taLeftJustify;

            // 해당 Row Auto-Sizing
            asg_Chat.AutoSizeRow(iChatRowCnt);

            // 메모리 해제
            Data.Free;


            Reciving:= False;

         end;

      except
         on e : Exception do
         begin
            showmessage('Server Read Data.Write ERROR : ' + e.Message);

            Data.Free;
         end;
      end;

   end
   //-----------------------------------------------------------
   // 파일 수신모드
   //-----------------------------------------------------------
   else if (PosByte('[A]',  sReceived) > 0) then
   begin
      // Client에 Msg Echo
      //SendToAllClients(sReceived);


      asg_Chat.InsertRows(iChatRowCnt, 1);

      asg_Chat.AddButton(C_CH_MYCOL,iChatRowCnt, asg_Chat.ColWidths[C_CH_MYCOL],  20, '다운', haBeforeText, vaCenter);

      asg_Chat.Cells[C_CH_TEXT,     iChatRowCnt] := CopyByte(sReceived, 1, PosByte('$', sReceived) - 1);  //CopyByte(sReceived, PosByte('$', sReceived) + 1, PosByte('[', sReceived) - PosByte('$', sReceived) -1); //

      //asg_Chat.Cells[C_CH_YOURCOL,  iChatRowCnt] := CopyByte(sReceived, PosByte('&', sReceived) + 1, 8);


      // Chat 상대 프로필 이미지 있으면, 사진 표기
      if FileExists(CopyByte(FindStr(sReceived, '%'), PosByte('&', FindStr(sReceived, '%')) + 2, LengthByte(FindStr(sReceived, '%')) - PosByte('&', FindStr(sReceived, '%')) - 2)) then
      begin
         // Chat 상대로부터 받은 나의 프로필 이미지 저장.
         FsChatMyPhoto  := CopyByte(sReceived, PosByte('%', sReceived) + 2, LengthByte(sReceived) - PosByte('%', sReceived) - 2);

         try
            asg_Chat.CreatePicture(C_CH_YOURCOL,
                                   iChatRowCnt,
                                   False,
                                   ShrinkWithAspectRatio,
                                   0,
                                   haLeft,
                                   vaTop).LoadFromFile(CopyByte(FindStr(sReceived, '%'), PosByte('&', FindStr(sReceived, '%')) + 2, LengthByte(FindStr(sReceived, '%')) - PosByte('&', FindStr(sReceived, '%')) - 2));

         except
            on e : Exception do
               showmessage('Create Picture ERROR(ServerSocketClientRead) : ' + e.Message);
         end;

         asg_Chat.RowHeights[iChatRowCnt] := 40;

      end
      else
      begin
         asg_Chat.Cells[C_CH_YOURCOL,  iChatRowCnt] := CopyByte(FindStr(sReceived, '%'), PosByte('&', FindStr(sReceived, '%')) + 1, LengthByte(FindStr(sReceived, '%')) - PosByte('%', FindStr(sReceived, '%')));
      end;


      {
      if FileExists(CopyByte(FindStr(sReceived, '%'), PosByte('&', FindStr(sReceived, '%')) + 2, LengthByte(FindStr(sReceived, '%')) - PosByte('%', FindStr(sReceived, '%')) - 2)) then
      begin

         FsChatMyPhoto  := CopyByte(FindStr(sReceived, '%'), PosByte('&', FindStr(sReceived, '%')) + 2, LengthByte(FindStr(sReceived, '%')) - PosByte('%', FindStr(sReceived, '%')) - 2);

         try
            asg_Chat.CreatePicture(C_CH_YOURCOL,
                                   iChatRowCnt,
                                   False,
                                   ShrinkWithAspectRatio,
                                   0,
                                   haLeft,
                                   vaTop).LoadFromFile(CopyByte(FindStr(sReceived, '%'), PosByte('&', FindStr(sReceived, '%')) + 2, LengthByte(FindStr(sReceived, '%')) - PosByte('%', FindStr(sReceived, '%')) - 2));

         except
            on e : Exception do
               showmessage('Create Picture ERROR(ServerSocketClientRead) : ' + e.Message);
         end;

         asg_Chat.RowHeights[iChatRowCnt] := 40;

      end
      else
         asg_Chat.Cells[C_CH_YOURCOL,  iChatRowCnt] := CopyByte(FindStr(sReceived, '%'), PosByte('&', FindStr(sReceived, '%')) + 1, LengthByte(FindStr(sReceived, '%')) - PosByte('%', FindStr(sReceived, '%')));
      }

      asg_Chat.Cells[C_CH_HIDDEN,   iChatRowCnt] := CopyByte(sReceived, PosByte('$', sReceived) + 1, PosByte('[', sReceived) - PosByte('$', sReceived) -1);

      asg_Chat.Row := iChatRowCnt;

      asg_Chat.Alignments[C_CH_YOURCOL, iChatRowCnt] := taLeftJustify;

      if LengthByte(asg_Chat.Cells[C_CH_TEXT,     iChatRowCnt]) < 36 then
         asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt] := taRightJustify
      else
         asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt] := taLeftJustify;

      //asg_Chat.AutoSizeRow(iChatRowCnt);


   end
   //-----------------------------------------------------------
   // 일반 Text 수신모드
   //-----------------------------------------------------------
   else
   begin
      // Client에 Msg Echo
      //SendToAllClients(sReceived);

      asg_Chat.InsertRows(iChatRowCnt, 1);


      //asg_Chat.Cells[C_CH_YOURCOL,  iChatRowCnt] := CopyByte(sReceived, PosByte('&', sReceived) + 1, 10);


      // test
      //showmessage('MainDlg에서 수신시 : ' + sReceived + #13#10#13#10 + CopyByte(FindStr(sReceived, '%'), PosByte('&', FindStr(sReceived, '%')) + 2, LengthByte(FindStr(sReceived, '%')) - PosByte('&', FindStr(sReceived, '%')) - 2) + #13#10#13#10 + CopyByte(sReceived, PosByte('%', sReceived) + 2, LengthByte(sReceived) - PosByte('%', sReceived) - 2));

      // Chat 상대 프로필 이미지 있으면, 사진 표기
      if FileExists(CopyByte(FindStr(sReceived, '%'), PosByte('&', FindStr(sReceived, '%')) + 2, LengthByte(FindStr(sReceived, '%')) - PosByte('&', FindStr(sReceived, '%')) - 2)) then
      begin
         // Chat 상대로부터 받은 나의 프로필 이미지 저장.
         FsChatMyPhoto  := CopyByte(sReceived, PosByte('%', sReceived) + 2, LengthByte(sReceived) - PosByte('%', sReceived) - 2);

         try
            asg_Chat.CreatePicture(C_CH_YOURCOL,
                                   iChatRowCnt,
                                   False,
                                   ShrinkWithAspectRatio,
                                   0,
                                   haLeft,
                                   vaTop).LoadFromFile(CopyByte(FindStr(sReceived, '%'), PosByte('&', FindStr(sReceived, '%')) + 2, LengthByte(FindStr(sReceived, '%')) - PosByte('&', FindStr(sReceived, '%')) - 2));

         except
            on e : Exception do
               showmessage('Create Picture ERROR(ServerSocketClientRead) : ' + e.Message);
         end;

         asg_Chat.RowHeights[iChatRowCnt] := 40;

      end
      else
      begin
         asg_Chat.Cells[C_CH_YOURCOL,  iChatRowCnt] := CopyByte(FindStr(sReceived, '%'), PosByte('&', FindStr(sReceived, '%')) + 1, LengthByte(FindStr(sReceived, '%')) - PosByte('%', FindStr(sReceived, '%')));
      end;


      asg_Chat.Cells[C_CH_TEXT,     iChatRowCnt] := CopyByte(sReceived, 1, PosByte('&', sReceived) - 1);
      asg_Chat.Row := iChatRowCnt;

      asg_Chat.Alignments[C_CH_YOURCOL, iChatRowCnt] := taLeftJustify;

      if LengthByte(asg_Chat.Cells[C_CH_TEXT,     iChatRowCnt]) < 36 then
         asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt] := taRightJustify
      else
         asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt] := taLeftJustify;

      asg_Chat.AutoSizeRow(iChatRowCnt);


   end;
end;




//------------------------------------------------------------------------------
// [Chat] Server Socket 메세지 전송 Event
//
// Date   : 2013.08.29
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.ButtonSendClick(Sender: TObject);
begin

   if FileExists(FsChatMyPhoto) then
      SendToAllClients(EditSay.Text + '&[' + FsChatMyPhoto + ']')
   else
      SendToAllClients(EditSay.Text + '&[' + FsUserNm + ']');

   asg_Chat.InsertRows(iChatRowCnt, 1);

   //asg_Chat.Cells[C_CH_MYCOL, iChatRowCnt] := '[' + FsUserNm + '] ';

   // 사용자 이미지 표기 @ 2015.03.30 LSH
   if FileExists(FsChatMyPhoto) then
   begin
      try
         // 대화상대 Image 파일을 Grid에 표기
         asg_Chat.CreatePicture(C_CH_MYCOL,
                                iChatRowCnt,
                                False,
                                ShrinkWithAspectRatio,
                                0,
                                haLeft,
                                vaTop).LoadFromFile(FsChatMyPhoto);
      except
         on e : Exception do
            showmessage('Create Picture ERROR(ButtonSendClick) : ' + e.Message);
      end;

      asg_Chat.RowHeights[iChatRowCnt] := 40;

   end
   else
   begin
      // test
      //showmessage('text 전송시 : ' + FsChatMyPhoto);

      asg_Chat.Cells[C_CH_MYCOL, iChatRowCnt]      := '[' + FsUserNm + ']';
   end;


   asg_Chat.AutoSizeRow(iChatRowCnt);

   asg_Chat.Cells[C_CH_TEXT,  iChatRowCnt] := EditSay.Text;
   asg_Chat.Row := iChatRowCnt;
   asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt] := taLeftJustify;


   Inc(iChatRowCnt);

   EditSay.Text := '';
end;



procedure TMainDlg.EditSayKeyPress(Sender: TObject; var Key: Char);
begin
   if (Key = #13) and (EditSay.Text <> '') then
      ButtonSendClick(Sender);
end;



//------------------------------------------------------------------------------
// [Chat] Server Socket 메세지 전송 onButtonClick Event Handler
//
// Date   : 2013.09.03
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_ChatSendButtonClick(Sender: TObject; ACol,
  ARow: Integer);
var
   sFileName, sHideFile, sServerIp : String;
   ms : TMemoryStream;
//   i : Integer;
begin
   if (ACol = C_SD_BUTTON) then
   begin


      Inc(iChatRowCnt);
      asg_Chat.RowCount := iChatRowCnt;   // Grid Index out of Range 방지위해 반드시 필요 !


      //----------------------------------------------------------
      // 이모티콘(Bitmap) 전송모드
      //----------------------------------------------------------
      if (IsEmotiSend) then
      begin

         ms := TMemoryStream.Create;

         try
            // 현재 불러온 이미지를 Memory-Stream으로 저장
            asg_ChatSend.GetPicture(C_SD_TEXT, 0).Bitmap.SaveToStream(ms);


            if ms.Size > 3000 then
            begin
               //
               MessageBox(self.Handle,
                          PChar('이모티콘용 Bitmap 이미지는 3KB 이하만 가능합니다.'),
                          '챗(Chat) :: 이모티콘(Bitmap) 전송 Size 제한',
                          MB_OK + MB_ICONERROR);

               // 메모리 해제
               ms.Free;

               // 입력창 이미지 제거
               asg_ChatSend.RemovePicture(C_SD_TEXT, 0);

               // 상태 Chagne
               IsEmotiSend := False;

               Exit;
            end;



            ms.Position := 0;


            // Data Stream 사이즈 정보 및 구분 Flag + 전송자 이름
            // Data Stream 사이즈 정보 + Chat 프로필 이미지(타이틀) 전송 @ 2015.03.31 LSH
            if FileExists(FsChatMyPhoto) then
               SendToAllClients(IntToStr(ms.Size) + '[E]' + '&[' + FsChatMyPhoto + ']' + #0)
            else
               SendToAllClients(IntToStr(ms.Size) + '[E]' + '&[' + FsUserNm + ']' + #0);



            // 현재 연결된 Client로 Data-Stream 전송
            ServerSocket.Socket.Connections[ServerSocket.Socket.ActiveConnections-1].SendStream(ms);



            try
               // 1줄 추가
               asg_Chat.InsertRows(iChatRowCnt, 1);

            except
               on e : Exception do
               showmessage('Server Send ERROR(InsertRows) : ' + e.Message);
            end;




            try
               // 전송창 Row에 전송하려는 이미지를 표기
               asg_Chat.CreatePicture(C_CH_TEXT, iChatRowCnt, True, ShrinkWithAspectRatio, 0, haleft, vaTop).Bitmap.LoadFromFile(OpenPictureDialog1.FileName);

            except
               asg_Chat.Cells[C_CH_TEXT, iChatRowCnt] := '[이미지 전송실패]';
            end;

            // 전송자 표기
            //asg_Chat.Cells[C_CH_MYCOL, iChatRowCnt]      := '[' + FsUserNm + '] ';
            if FileExists(FsChatMyPhoto) then
            begin
               try
                  // 대화상대 Image 파일을 Grid에 표기
                  asg_Chat.CreatePicture(C_CH_MYCOL,
                                         iChatRowCnt,
                                         False,
                                         ShrinkWithAspectRatio,
                                         0,
                                         haLeft,
                                         vaTop).LoadFromFile(FsChatMyPhoto);

               except
                  on e : Exception do
                  showmessage('Create Picture ERROR(asg_ChatSendButtonClick) : ' + e.Message);
               end;

               asg_Chat.RowHeights[iChatRowCnt] := 40;

            end
            else
               asg_Chat.Cells[C_CH_MYCOL, iChatRowCnt]      := '[' + FsUserNm + ']';

            asg_Chat.AutoSizeRow(iChatRowCnt);

            // Scroll 조절
            asg_Chat.Row   := iChatRowCnt;

            // 입력창 이미지 제거
            asg_ChatSend.RemovePicture(C_SD_TEXT, 0);

            // 상태 Chagne
            IsEmotiSend := False;


         except
            ms.Free;
         end;

      end
      //----------------------------------------------------------
      // File 전송모드
      //----------------------------------------------------------
      else if (IsFileSend) then
      begin

         sFileName   := '';
         sHideFile   := '';
         sServerIp   := '';


         // 첨부파일 Upload
         if (Trim(asg_ChatSend.Cells[C_SD_TEXT, 0]) <> '') then
         begin

            sFileName := ExtractFileName(Trim(asg_ChatSend.Cells[C_SD_TEXT, 0]));
            sHideFile := 'CHATAPPEND' + DelChar(CopyByte(asg_Master.Cells[C_M_USERIP, asg_Master.Row], LengthByte(asg_Master.Cells[C_M_USERIP, asg_Master.Row])-3, 4), '.') +
                          FormatDateTime('YYYYMMDDHHNNSS', Now) + '.KDIALFILE';


            if Not FileUpLoad(sHideFile, asg_ChatSend.Cells[C_SD_TEXT, 0], sServerIp) then
            begin
               Showmessage('첨부파일 UpLoad 중 에러가 발생했습니다.' + #13#10 + #13#10 +
                           '다시한번 시도해 주시기 바랍니다.');
               Exit;
            end;
         end;

         // 첨부파일 + Chat 프로필 이미지(타이틀) 전송
         if FileExists(FsChatMyPhoto) then
            SendToAllClients(sFileName + '$' + sHideFile + '[A]' + '&[' + FsChatMyPhoto + ']')
         else
            SendToAllClients(sFileName + '$' + sHideFile + '[A]' + '&[' + FsUserNm + ']');



         asg_Chat.InsertRows(iChatRowCnt, 1);

         //asg_Chat.Cells[C_CH_MYCOL,    iChatRowCnt]      := '[' + FsUserNm + ']';

         // 전송자 이미지 표기 @ 2015.03.30 LSH
         if FileExists(FsChatMyPhoto) then
         begin
            try
               // 대화상대 Image 파일을 Grid에 표기
               asg_Chat.CreatePicture(C_CH_MYCOL,
                                      iChatRowCnt,
                                      False,
                                      ShrinkWithAspectRatio,
                                      0,
                                      haLeft,
                                      vaTop).LoadFromFile(FsChatMyPhoto);

            except
               on e : Exception do
                  showmessage('Create Picture ERROR(asg_ChatSendButtonClick) : ' + e.Message);
            end;

            asg_Chat.RowHeights[iChatRowCnt] := 40;

         end
         else
            asg_Chat.Cells[C_CH_MYCOL, iChatRowCnt]      := '[' + FsUserNm + ']';

         asg_Chat.AutoSizeRow(iChatRowCnt);

         asg_Chat.Cells[C_CH_TEXT,     iChatRowCnt]      := sFileName;
         asg_Chat.Cells[C_CH_HIDDEN,   iChatRowCnt]      := sHideFile;
         asg_Chat.Cells[C_CH_YOURCOL,  iChatRowCnt]      := '<전송됨>';

         //asg_Chat.AddButton(C_CH_YOURCOL,  iChatRowCnt, asg_Chat.ColWidths[C_CH_YOURCOL]-20,  20, '다운', haBeforeText, vaCenter);

         asg_Chat.Row                                 := iChatRowCnt;
         asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt]  := taLeftJustify;

         asg_ChatSend.Cells[C_SD_TEXT, 0] := '';

         IsFileSend := False;

      end
      //----------------------------------------------------------
      // 일반 Text 전송모드
      //----------------------------------------------------------
      else
      begin
         // test
         //showmessage('MainDlg에서 Text 전송시 : ' + FsChatMyPhoto);

         if FileExists(FsChatMyPhoto) then
            SendToAllClients(asg_ChatSend.Cells[C_SD_TEXT, 0] + '&[' + FsChatMyPhoto + ']')
         else
            SendToAllClients(asg_ChatSend.Cells[C_SD_TEXT, 0] + '&[' + FsUserNm + ']');


         asg_Chat.InsertRows(iChatRowCnt, 1);

         //asg_Chat.Cells[C_CH_MYCOL, iChatRowCnt] := '[' + FsUserNm + '] ';

         // 전송자 이미지 표기 @ 2015.03.30 LSH
         if FileExists(FsChatMyPhoto) then
         begin
            try
               // 대화상대 Image 파일을 Grid에 표기
               asg_Chat.CreatePicture(C_CH_MYCOL,
                                      iChatRowCnt,
                                      False,
                                      ShrinkWithAspectRatio,
                                      0,
                                      haLeft,
                                      vaTop).LoadFromFile(FsChatMyPhoto);

            except
               on e : Exception do
                  showmessage('Create Picture ERROR(asg_ChatSendButtonClick) : ' + e.Message);
            end;

            asg_Chat.RowHeights[iChatRowCnt] := 40;

         end
         else
         begin
            // test
            //showmessage('일반 Text 전송모드 : ' + FsChatMyPhoto);

            asg_Chat.Cells[C_CH_MYCOL, iChatRowCnt]      := '[' + FsUserNm + ']';
         end;


         asg_Chat.AutoSizeRow(iChatRowCnt);

         asg_Chat.Cells[C_CH_TEXT,  iChatRowCnt] := asg_ChatSend.Cells[C_SD_TEXT, 0];
         asg_Chat.Row := iChatRowCnt;
         asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt] := taLeftJustify;




         asg_ChatSend.Cells[C_SD_TEXT, 0] := '';
      end;
   end;
end;



//------------------------------------------------------------------------------
// [Chat] Server Socket 메세지 전송 onKeyPress Event Handler
//
// Date   : 2013.09.03
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_ChatSendKeyPress(Sender: TObject; var Key: Char);
begin
   if (Key = #13) and (asg_ChatSend.Cells[C_SD_TEXT, 0] <> '') then
      asg_ChatSendButtonClick(Sender, C_SD_BUTTON, 0);
end;




procedure TMainDlg.asg_RegDocKeyPress(Sender: TObject; var Key: Char);
begin
   if (Key = #13) then
   begin
      if asg_RegDoc.Col = C_RD_DOCLOC then
         asg_RegDoc.Navigation.AdvanceOnEnter := False
      else
         asg_RegDoc.Navigation.AdvanceOnEnter := True;

      // 문서번호 Cell에서 'Enter' 입력시, max-seq 자동채번 @ 2015.05.13 LSH
      if asg_RegDoc.Col = C_RD_DOCSEQ then
         asg_RegDoc_ClickCell(Sender, asg_RegDoc.Row, asg_RegDoc.Col);

   end;
end;



procedure TMainDlg.pm_MasterPopup(Sender: TObject);
begin

   if (IsLogonUser) then
   begin
      //------------------------------------------------------------
      // 접속하려는 Server 사용가능 여부 (현재 Session On/Off) Check
      //------------------------------------------------------------
      if (apn_Master.Visible) then
      begin
         if (asg_Master.Colors[C_M_USERNM, asg_Master.Row] = clWhite) then
            mi_Chat.Visible := False
         else
            mi_Chat.Visible := True;
      end
      else if (apn_Network.Visible) then
      begin
         if (asg_Network.FontColors[C_NW_USERNM, asg_Network.Row] = clBlack) then
            mi_Chat.Visible := False
         else
            mi_Chat.Visible := True;
      end;

      //------------------------------------------------------------
      // 닉네임 입력/수정은 [담당자 프로필]에서만 가능
      //------------------------------------------------------------
      if (apn_Master.Visible) then
         mi_NickNm.Visible := True
      else
         mi_NickNm.Visible := False;

      //------------------------------------------------------------
      // SMS는 [비상연락망], [다이얼Book]에서만 사용시작 (추후 확대예정)
      //------------------------------------------------------------
      if (ftc_Dialog.ActiveTab = AT_DIALOG) or
         (ftc_Dialog.ActiveTab = AT_DIALBOOK) then
      begin
         if (
               (ftc_Dialog.ActiveTab = AT_DIALBOOK)
               and
               (
                  (
                     (CopyByte(asg_DialList.Cells[C_DL_CALLNO, asg_DialList.Row], 1, 3) = '010') and
                     (TokenStr(asg_DialList.Cells[C_DL_CALLNO, asg_DialList.Row], #13#10, 2) = '')
                  ) or
                  (
                     (CopyByte(asg_MyDial.Cells[C_MD_MOBILE, asg_MyDial.Row], 1, 3) = '010') and
                     (TokenStr(asg_MyDial.Cells[C_MD_MOBILE, asg_MyDial.Row], #13#10, 2) = '')
                  ) or
                  (
                     PosByte('010', asg_DialList.Cells[C_DL_CALLNO, asg_DialList.Row]) > 0
                  ) or
                  (
                     PosByte('010', asg_MyDial.Cells[C_MD_MOBILE, asg_MyDial.Row]) > 0
                  )
               )
            ) or
            (
               (ftc_Dialog.ActiveTab = AT_DIALOG) and
               (apn_Network.Visible)
            ) then
            mi_SMS.Visible    := True
         else
            mi_SMS.Visible    := False;
      end
      else
         mi_SMS.Visible    := False;


   end;
end;



//------------------------------------------------------------------------------
// [함수] TCP Port 오픈여부 Check
//    - 2013.08.30 LSH
//    - 소스출처 : Google
//------------------------------------------------------------------------------
function TMainDlg.PortTCPIsOpen(dwPort : Word; ipAddressStr:string) : boolean;
var
   client : sockaddr_in;                                                         // sockaddr_in is used by Windows Sockets to specify a local or remote endpoint address
   sock   : Integer;
begin
   client.sin_family       := AF_INET;
   client.sin_port         := htons(dwPort);                                     // htons converts a u_short from host to TCP/IP network byte order.
   client.sin_addr.s_addr  := inet_addr(PAnsiChar(AnsiString(ipAddressStr)));    // the inet_addr function converts a string containing an IPv4 dotted-decimal address into a proper address for the IN_ADDR structure.
   sock                    := socket(AF_INET, SOCK_STREAM, 0);                   // The socket function creates a socket
   Result                  := connect(sock,client,SizeOf(client)) = 0;           // establishes a connection to a specified socket.
end;







//------------------------------------------------------------------------------
// ServerSocket onClientError Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2013.09.02
//------------------------------------------------------------------------------
procedure TMainDlg.ServerSocketClientError(Sender: TObject;
  Socket: TCustomWinSocket; ErrorEvent: TErrorEvent;
  var ErrorCode: Integer);
begin

   if (ErrorCode = 10053) then
   begin
      // Asynch 소켓 에러시, 서버소켓 Off-> On
      ServerSocket.Active := False;
      ServerSocket.Active := True;

      {
      // ErrorCode = 0으로 두면, '응답없음'으로 강제 종료됨.. 버그 잡아야 함..
      //ErrorCode := 1;
      }

      // 에러메세지 출력 안함.
      ErrorCode := 0;


      {
      MessageBox(self.Handle,
                 PChar('ServerSocketError(' + FsUserIP + '): ' + Socket.RemoteHost + '(' + Socket.RemoteAddress + ') 로부터 잘못된 Connection Error 발생(소켓핸들값: ' + IntToStr(Socket.SocketHandle) + ')'),
                 '챗(Chat) :: 잘못된 Socket 통신 오류',
                 MB_OK + MB_ICONERROR);
      }

   end;


end;






//------------------------------------------------------------------------------
// [종료] Form onClose Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2013.09.02
//------------------------------------------------------------------------------
procedure TMainDlg.FormClose(Sender: TObject; var Action: TCloseAction);
var
   i : Integer;
   varResult : String;
begin



   // 로그 Update
   UpdateLog('START',
             CopyByte(self.Caption, PosByte('[', self.Caption) + 1, 6),
             FsUserIp,
             'O',
             FormatDateTime('yyyy-mm-dd hh:nn', Now),
             FsUserNm,
             varResult
             );


   {
   try
      if (ServerSocket.Socket.ActiveConnections > 0) then
      begin

         showmessage('ServerSocket.Socket.ActiveConnections = '  + inttostr(ServerSocket.Socket.ActiveConnections));

         for i := 0 to (ServerSocket.Socket.ActiveConnections - 1) do
         begin
            ServerSocket.Socket.Connections[i].Close;
         end;

      end;

      //ServerSocket.Socket.Disconnect(ServerSocket.Socket.SocketHandle);
      //ServerSocket.Socket.Close;
      //Action := caFree;
   except
      on E : Exception do
      ShowMessage(E.ClassName+' error raised by FormClose, with message : ' + E.Message);
   end;
   }

   // 타이머 종료 추가 @ 2014.06.03 LSH
   for i:= 0 to ComponentCount - 1 do
   begin
      if Components[i] is TTimer then
         TTimer(Components[i]).Enabled := False;
   end;

   Action := caFree;
end;



//------------------------------------------------------------------------------
// [Chat 첨부] AdvStringGrid onButtonClick Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2013.09.09
//------------------------------------------------------------------------------
procedure TMainDlg.asg_ChatButtonClick(Sender: TObject; ACol,
  ARow: Integer);
var
   sLocalFile, sRemoteFile : String;
   sServerIp,  sFtpUserID, sFtpPasswd, sFtpHostName, sFtpDir : String;
begin
   //---------------------------------------------
   // 첨부파일 Click시 Event 처리
   //---------------------------------------------
   with asg_Chat do
   begin
      if ((ACol = C_CH_MYCOL) and (Cells[C_CH_HIDDEN, ARow] <> '')) then
      begin
         // 파일 업/다운로드를 위한 정보 조회
         if not GetBinUploadInfo(sServerIp, sFtpUserID, sFtpPasswd, sFtpHostName, sFtpDir) then
         begin
            MessageDlg('파일 저장을 위한 담당자 정보 조회중, 오류가 발생했습니다.', Dialogs.mtError, [Dialogs.mbOK], 0);
            exit;
         end;



         //실제 서버에 저장되어 있는 파일명 지정
         if PosByte('/ftpspool/CHATFILE/',Cells[C_CH_HIDDEN, ARow]) > 0 then
            sRemoteFile := Cells[C_CH_HIDDEN, ARow]
         else
            sRemoteFile := '/ftpspool/CHATFILE/' + Cells[C_CH_HIDDEN, ARow];

         // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
         sLocalFile  := G_HOMEDIR + 'TEMP\SPOOL\' + Cells[C_CH_TEXT, ARow]; //lb_Filename.Caption;



         if (GetBINFTP(sServerIp, sFtpUserID, sFtpPasswd, sRemoteFile, sLocalFile, False)) then
         begin
            //	ShowMessage('정상적으로 저장되었습니다.');
         end;



         Screen.Cursor := crDefault;



         try

            if (PosByte('.exe', sLocalFile) > 0) or
               (PosByte('.zip', sLocalFile) > 0) or
               (PosByte('.rar', sLocalFile) > 0) then
            begin
               MessageBox(self.Handle,
                          PChar('첨부파일(' + sLocalFile + ') 다운로드가 완료되었습니다.' + #13#10 + #13#10 +
                                '※ 임시 다운로드 폴더 --> C:\KUMC(_DEV)\TEMP\SPOOL\'),
                          '첨부파일 다운로드 실행 완료',
                          MB_OK + MB_ICONINFORMATION);
            end
            else
            begin
               ShellExecute(HANDLE, 'open',
                            PCHAR(Cells[C_CH_TEXT, ARow]),
                            PCHAR(''),
                            // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
                            PCHAR(G_HOMEDIR + 'TEMP\SPOOL\'),
                            SW_SHOWNORMAL);
            end;

         except
            MessageBox(self.Handle,
                       PChar('해당 프로그램 다운로드중 오류가 발생하였습니다.' + #13#10 + #13#10 +
                             '프로그램 종료 후 다시 실행해 주시기 바랍니다.'),
                       '첨부파일 다운로드 중 오류',
                       MB_OK + MB_ICONERROR);

            Exit;
         end;
      end;
   end;
end;






//------------------------------------------------------------------------------
// [팝업메뉴] Chat 파일 전송
//
// Author : Lee, Se-Ha
// Date   : 2013.09.09
//------------------------------------------------------------------------------
procedure TMainDlg.mi_FileSendClick(Sender: TObject);
begin
   if od_File.Execute then
      asg_ChatSend.Cells[C_SD_TEXT, 0] := ExtractFileName(od_File.FileName);


   if asg_ChatSend.Cells[C_SD_TEXT, 0] <> '' then
   begin
      // 파일 전송모드 True
      IsFileSend := True;

      // 자동 전송
      asg_ChatSendButtonClick(Sender, C_SD_BUTTON, 0);
   end;

end;



//------------------------------------------------------------------------------
// [전체선택] FlatMemo onKeyUp Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2013.09.03
//------------------------------------------------------------------------------
procedure TMainDlg.fmm_TextKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   // 작성모드시, [Ctrl + A]로 전체 글 선택되도록.
   if(ssCtrl in Shift) and (Key = 65) then
	   fmm_Text.SelectAll;


   if fsbt_WriteFind.Visible then
   begin
      // 글 검색 모드에서, [Ctrl + Enter]로 검색 자동실행 @ 2016.11.18 LSH
      if(ssCtrl in Shift) and (Key = VK_RETURN) then
         fsbt_WriteFindClick(Sender);
   end;
end;




//------------------------------------------------------------------------------
// Data Loading Bar Controller
//
// Author : Lee, Se-Ha
// Date   : 2013.09.06
//------------------------------------------------------------------------------
procedure TMainDlg.SetLoadingBar(AsStatus : String);
begin
   //---------------------------------------------------------------
   // 1. Set Status
   //---------------------------------------------------------------
   if AsStatus = 'ON' then
   begin
      // 진행중 표시 보이기
      pn_Loading.Left    := 128;
      pn_Loading.Top     := 200;
      pn_Loading.Visible := True;

      pn_Loading.Repaint;

   end
   else if AsStatus = 'OFF' then
   begin
      // 진행중 표시 끄기

      pb_DataLoading.Position := 0;
      pn_Loading.Visible      := False;

   end;
end;



//------------------------------------------------------------------------------
// [타이머] 담당자 정보 자동 Refresh 위한 Timer
//
// Author : Lee, Se-Ha
// Date   : 2013.09.06
//------------------------------------------------------------------------------
procedure TMainDlg.tm_MasterTimer(Sender: TObject);
begin
   // Interval = 60,000 (1분) 마다 자동 Refresh
   bbt_NetworkRefreshClick(Sender);
end;




//------------------------------------------------------------------------------
// [팝업메뉴] Chat 이모티콘 전송
//
// Author : Lee, Se-Ha
// Date   : 2013.09.09
//------------------------------------------------------------------------------
procedure TMainDlg.mi_EmotiClick(Sender: TObject);
begin
   if OpenPictureDialog1.Execute then
   begin
      with asg_ChatSend do
      begin
         CreatePicture(C_SD_TEXT, 0, True, ShrinkWithAspectRatio, 0, haLeft, vaTop).LoadFromFile(OpenPictureDialog1.FileName);
      end;
   end;

   if asg_ChatSend.GetPicture(C_SD_TEXT, 0) <> nil then
   begin
      // 이모티콘 전송모드 True
      IsEmotiSend := True;

      // 자동 전송
      asg_ChatSendButtonClick(Sender, C_SD_BUTTON, 0);
   end;
end;




//------------------------------------------------------------------------------
// [타이머] Timer onTimer Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2013.09.12
//------------------------------------------------------------------------------
procedure TMainDlg.tm_TxInitTimer(Sender: TObject);
begin
   //--------------------------------------------------------------------
   // 90초 마다, 개발계 Tmax 연결 확인 및 복구
   //--------------------------------------------------------------------
   if not (TxInit(TokenStr(String(CONST_ENV_FILENAME), '.', 1) + '_D0.ENV', '01')) then
   begin
      MessageBox(self.Handle,
                 PChar('현재 Server와의 연결이 끊겼습니다.' + #13#10 + #13#10 + '90초마다 자동으로 연결을 복구하오니, 잠시만 기다려 주십시오.'),
                 PChar(Self.Caption + ' : TMAX 접속끊김 알림'),
                 MB_OK + MB_ICONWARNING);
   end;
end;




procedure TMainDlg.fsbt_MenuClick(Sender: TObject);
begin
   // 현 시점 식단 메뉴 안내 팝업 @ 2017.10.26 LSH
   MessageBox(self.Handle,
              PChar(GetDietInfo(FormatDateTime('dd', Date),
                                GetDayofWeek(Date, 'HS'))),
              '병원 밥 맛있게 드세요 :-)',
              MB_OK + MB_ICONINFORMATION);

   try
      // 관리자만 비밀 실험실(?)에 접속가능 @ 2017.10.26 LSH
      if (FsUserIp = '개발자IP') then
      begin
         // 기존 식이정보 업로드 Panel을 빅데이터 분석으로 전환 @ 2015.04.02 LSH
         apn_Menu.Top      := 43;
         apn_Menu.Left     := 0;
         apn_Menu.Collaps  := True;
         apn_Menu.Visible  := True;
         apn_Menu.Collaps  := False;

         fmed_LotFrDt.Text := FormatDateTime('yyyy-mm-dd', Date - 730);
         fmed_LotToDt.Text := FormatDateTime('yyyy-mm-dd', Date);


         // Refresh
         SelGridInfo('BIGDATA');

         {
         // 접속자 IP 식별후, 근무처 Assign
         if PosByte('안암도메인', FsUserIp) > 0 then
            fcb_MenuLoc.Text := '안암병원'
         else if PosByte('구로도메인', FsUserIp) > 0 then
            fcb_MenuLoc.Text := '구로병원'
         else if PosByte('안산도메인', FsUserIp) > 0 then
            fcb_MenuLoc.Text := '안산병원';

         }

         //asg_MenuBar.AddButton(1, 0, asg_MenuBar.ColWidths[1]-5, 20, '등록', haBeforeText, vaCenter);    // 등록 Button


         {
         asg_Menu.ClearRows(2, asg_Menu.RowCount);
         asg_Menu.RowCount := 2;
         }
      end;

   except
      //
   end;

end;



procedure TMainDlg.mi_MenuOpenClick(Sender: TObject);
var
   sFileExt : String;
begin
   //--------------------------------------------------------------------
   // 1. 파일 오픈(.CSV, .XLS) 및 Data Loading
   //--------------------------------------------------------------------
   if opendlg_File.Execute then
   begin
      if opendlg_File.FileName = '' then
        Exit;


      sFileExt  := AnsiUpperCase(ExtractFileExt(opendlg_File.FileName));


      // Excel 파일인 경우, 별도 Procedure (LoadExcelFile) Call
      if (sFileExt = '.XLSX') or (sFileExt = '.XLS') Then
      begin
         if opendlg_File.FileName <> '' then
         begin
            LoadExcelFile(opendlg_File.FileName);
         end;
      end
      else
      // CSV 는 아래 로직 진행.
      begin
         // 준비중..
      end;
   end;
end;



//------------------------------------------------------------------------------
// [파일] Excel Load Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2013.09.12
//------------------------------------------------------------------------------
procedure TMainDlg.LoadExcelFile(in_FileName : String);
var
   xlsfilename: String;
   oXL, oWK, oSheet: Variant;
   XLRows, XLCols : integer;
   i, j, k : integer;
   iCol_RegDate ,
   iCol_DutySpec,
   iCol_Context ,
   iCol_ClienSrc,
   iCol_Client  ,
   iCol_ServeSrc,
   iCol_Server  ,
   iCol_SrReqNo ,
   iCol_DutyUser,
   iCol_ReqUser ,
   iCol_CofmUser,
   iCol_TestDate,
   iCol_ReleasDt,
   iCol_Remark,
   iCol_ReleasUser, // 릴리즈 담당자 추가 @ 2014.06.19 LSH
   iRow  : Integer;

   // 빅데이터 Lab.관련 항목 추가 @ 2015.04.02 LSH
   iCol_Year,
   iCol_Seqno,
   iCol_ShowDate,
   iCol_1stCount,
   iCol_1stRcpamt,
   iCol_2ndCount,
   iCol_2ndRcpamt,
   iCol_3thCount,
   iCol_3thRcpamt,
   iCol_4thCount,
   iCol_4thRcpamt,
   iCol_5thCount,
   iCol_5thRcpamt,
   iCol_No1,
   iCol_No2,
   iCol_No3,
   iCol_No4,
   iCol_No5,
   iCol_No6,
   iCol_NoBonus : Integer;

   XLLastUsedRows : Integer;
begin

   iCol_RegDate      := 0;
   iCol_DutySpec     := 0;
   iCol_Context      := 0;
   iCol_ClienSrc     := 0;
   iCol_Client       := 0;
   iCol_ServeSrc     := 0;
   iCol_Server       := 0;
   iCol_SrReqNo      := 0;
   iCol_DutyUser     := 0;
   iCol_ReqUser      := 0;
   iCol_CofmUser     := 0;
   iCol_TestDate     := 0;
   iCol_ReleasDt     := 0;
   iCol_Remark       := 0;
   iCol_ReleasUser   := 0;
   iRow              := 0;

   // 빅데이터 Lab.관련 항목 추가 @ 2015.04.02 LSH
   iCol_Year         := 0;
   iCol_Seqno        := 0;
   iCol_ShowDate     := 0;
   iCol_1stCount     := 0;
   iCol_1stRcpamt    := 0;
   iCol_2ndCount     := 0;
   iCol_2ndRcpamt    := 0;
   iCol_3thCount     := 0;
   iCol_3thRcpamt    := 0;
   iCol_4thCount     := 0;
   iCol_4thRcpamt    := 0;
   iCol_5thCount     := 0;
   iCol_5thRcpamt    := 0;
   iCol_No1          := 0;
   iCol_No2          := 0;
   iCol_No3          := 0;
   iCol_No4          := 0;
   iCol_No5          := 0;
   iCol_No6          := 0;
   iCol_NoBonus      := 0;


   try

      xlsfilename := in_FileName;

      oXL    := CreateOleObject('Excel.Application');
      oXL.WorkBooks.Open(xlsfilename, 0, true);          // 읽기 전용으로 읽기

      oWK    := oXL.WorkBooks.Item[1];
      oSheet := oWK.ActiveSheet;                         // 선택된 쉬트 가져오기

      //XLRows := oXL.ActiveSheet.UsedRange.Rows.Count;    // 총 Rows 수 구하기.
      //XLRows := 100;

      XLCols := oXL.ActiveSheet.UsedRange.Columns.Count; // 총 Columns 수 구하기.

      // 실제 마지막으로 작성된 Row의 Index를 fetch @ 2017.10.25 LSH
      XLLastUsedRows := oXL.ActiveSheet.Cells.Find('*', SearchOrder:=xlByRows, LookIn:=xlValues, SearchDirection:=xlPrevious).Row;


      if (ftc_Dialog.ActiveTab = AT_BOARD) then
      begin
         {  -- 기존 병원별 식이정보 update 주석처리 @ 2015.04.02 LSH
         //--------------------------------------------------------------------
         // 2-1. Excel Data Loading
         //--------------------------------------------------------------------
         asg_Menu.BeginUpdate;



         // 병원별로 영양팀 Xls 포맷이 다름.
         if PosByte('안암', fcb_MenuLoc.Text) > 0 then
         begin

            for i := 3 to XLRows do
            begin

               for j := 2 to XLCols do
               begin
                  //---------------------------------------------------------------
                  // Excel sheet의 data를 Grid에 뿌려주기
                  //---------------------------------------------------------------


                  case i of
                     3 :
                        asg_Menu.Cells[j - 2, i - 3] := oXL.Cells[i, j];

                     4..8 :
                        begin
                           // 아침끼니
                           asg_Menu.Cells[j - 2,  i - 3] := oXL.Cells[i, j];
                           asg_Menu.Colors[j - 2, i - 3] := $00C5B189;
                        end;


                     10..15 :
                        begin
                           // 점심끼니
                           asg_Menu.CellS[j - 2,  i - 3] := oXL.Cells[i, j];
                           asg_Menu.Colors[j - 2, i - 3] := $0070CDBD;
                        end;


                     17..21 :
                        begin
                           // 저녁끼니
                           asg_Menu.CellS[j - 2,  i - 3] := oXL.Cells[i, j];
                           asg_Menu.Colors[j - 2, i - 3] := $0070AFF7;
                        end;
                  end;
               end;
            end;

            // RowCount 정리
            asg_Menu.RowCount := XLRows - 5;

         end
         else if PosByte('구로', fcb_MenuLoc.Text) > 0 then
         begin
            for i := 3 to XLRows do
            begin
               for j := 2 to XLCols do
               begin
                  //---------------------------------------------------------------
                  // Excel sheet의 data를 Grid에 뿌려주기
                  //---------------------------------------------------------------
                  case i of
                     3 :
                        asg_Menu.CellS[j - 2, i - 3] := oXL.Cells[i, j];

                     4..9 :
                        begin
                           // 아침끼니
                           asg_Menu.CellS[j - 2, i - 3]  := oXL.Cells[i, j];
                           asg_Menu.Colors[j - 2, i - 3] := $00C5B189;
                        end;


                     11..16 :
                        begin
                           // 점심끼니
                           asg_Menu.CellS[j - 2, i - 4]  := oXL.Cells[i, j];
                           asg_Menu.Colors[j - 2, i - 4] := $0070CDBD;
                        end;


                     18..23 :
                        begin
                           // 저녁끼니
                           asg_Menu.CellS[j - 2, i - 5]  := oXL.Cells[i, j];
                           asg_Menu.Colors[j - 2, i - 5] := $0070AFF7;
                        end;
                  end;
               end;
            end;


            // RowCount 정리
            asg_Menu.RowCount := XLRows - 11;


         end
         else if PosByte('안산', fcb_MenuLoc.Text) > 0 then
         begin
            for i := 3 to XLRows do
            begin
               for j := 2 to XLCols do
               begin
                  //---------------------------------------------------------------
                  // Excel sheet의 data를 Grid에 뿌려주기
                  //---------------------------------------------------------------
                  case i of
                     3 :
                        asg_Menu.CellS[j - 2, i - 3] := oXL.Cells[i, j];

                     5..8 :
                        begin
                           // 아침끼니
                           asg_Menu.CellS[j - 2, i - 4]  := oXL.Cells[i, j];
                           asg_Menu.Colors[j - 2, i - 4] := $00C5B189;
                        end;


                     13..16 :
                        begin
                           // 점심끼니
                           asg_Menu.CellS[j - 2, i - 8]  := oXL.Cells[i, j];
                           asg_Menu.Colors[j - 2, i - 8] := $0070CDBD;
                        end;


                     21..24 :
                        begin
                           // 저녁끼니
                           asg_Menu.CellS[j - 2, i - 12]  := oXL.Cells[i, j];
                           asg_Menu.Colors[j - 2, i - 12] := $0070AFF7;
                        end;
                  end;
               end;
            end;

            // RowCount 정리
            asg_Menu.RowCount := XLRows - 17;

         end;


         asg_Menu.EndUpdate;

         asg_Menu.Refresh;
         }



         //--------------------------------------------------------------------
         // 1. Data Loading bar 보이기
         //--------------------------------------------------------------------
         SetLoadingBar('ON');


         // Maximum value of progress status
         pb_DataLoading.Max := XLRows;


         //--------------------------------------------------------------------
         // 2-2. Excel Data Loading
         //--------------------------------------------------------------------
         //asg_BigData.BeginUpdate;


//         iRow := 0;


         for i := 1 to XLRows do
         begin

            //---------------------------------------------------------------------
            // 빅데이터 Excel sheet의 Title 항목의 Index 값을 가져오기
            //---------------------------------------------------------------------
            if i = 1 then
            begin
               for k := 1 to XLCols do
               begin
                  if Trim(oXL.Cells[i, k]) = '년도' then
                     iCol_Year      := k
                  else if Trim(oXL.Cells[i, k]) = '회차' then
                     iCol_Seqno     := k
                  else if Trim(oXL.Cells[i, k]) = '기준일' then
                     iCol_ShowDate  := k
                  else if Trim(oXL.Cells[i, k]) = '인원1' then
                     iCol_1stCount  := k
                  else if Trim(oXL.Cells[i, k]) = '금액1' then
                     iCol_1stRcpamt := k
                  else if Trim(oXL.Cells[i, k]) = '인원2' then
                     iCol_2ndCount  := k
                  else if Trim(oXL.Cells[i, k]) = '금액2' then
                     iCol_2ndRcpamt := k
                  else if Trim(oXL.Cells[i, k]) = '인원3' then
                     iCol_3thCount  := k
                  else if Trim(oXL.Cells[i, k]) = '금액3' then
                     iCol_3thRcpamt := k
                  else if Trim(oXL.Cells[i, k]) = '인원4' then
                     iCol_4thCount  := k
                  else if Trim(oXL.Cells[i, k]) = '금액4' then
                     iCol_4thRcpamt := k
                  else if Trim(oXL.Cells[i, k]) = '인원5' then
                     iCol_5thCount  := k
                  else if Trim(oXL.Cells[i, k]) = '금액5' then
                     iCol_5thRcpamt := k
                  else if Trim(oXL.Cells[i, k]) = '넘버1' then
                     iCol_No1 := k
                  else if Trim(oXL.Cells[i, k]) = '넘버2' then
                     iCol_No2 := k
                  else if Trim(oXL.Cells[i, k]) = '넘버3' then
                     iCol_No3 := k
                  else if Trim(oXL.Cells[i, k]) = '넘버4' then
                     iCol_No4 := k
                  else if Trim(oXL.Cells[i, k]) = '넘버5' then
                     iCol_No5 := k
                  else if Trim(oXL.Cells[i, k]) = '넘버6' then
                     iCol_No6 := k
                  else if Trim(oXL.Cells[i, k]) = '보너스' then
                     iCol_NoBonus := k;

               end;
            end;



            for j := 1 to XLCols do
            begin
               if i > 1 then
               begin
                  //---------------------------------------------------------------
                  // 빅데이터 Excel sheet의 Title별 data를 Grid에 뿌려주기
                  //---------------------------------------------------------------
                  begin
                     // 년차
                     if (j mod iCol_Year = 0) and (j div iCol_Year = 1) then
                     begin
                        asg_BigData.Cells[C_BD_YEAR, i - 1] := CopyByte(oXL.Cells[i, j], 1, 16);
                     end
                     // 회차
                     else if (j mod iCol_Seqno = 0) and (j div iCol_Seqno = 1) then
                        asg_BigData.Cells[C_BD_SEQNO,   i - 1] := oXL.Cells[i, j]
                     // 기준일
                     else if (j mod iCol_ShowDate = 0) and (j div iCol_ShowDate = 1) then
                        asg_BigData.Cells[C_BD_SHOWDT,i - 1] := ReplaceChar(CopyByte(oXL.Cells[i, j], 1, 10), '.', '-')
                     // 인원1
                     else if (j mod iCol_1stCount = 0) and (j div iCol_1stCount = 1) then
                        asg_BigData.Cells[C_BD_1STCNT,   i - 1] := oXL.Cells[i, j]
                     // 금액1
                     else if (j mod iCol_1stRcpamt = 0) and (j div iCol_1stRcpamt = 1) then
                        asg_BigData.Cells[C_BD_1STAMT,   i - 1] := oXL.Cells[i, j]
                     // 인원2
                     else if (j mod iCol_2ndCount = 0) and (j div iCol_2ndCount = 1) then
                        asg_BigData.Cells[C_BD_2NDCNT,   i - 1] := oXL.Cells[i, j]
                     // 금액2
                     else if (j mod iCol_2ndRcpamt = 0) and (j div iCol_2ndRcpamt = 1) then
                        asg_BigData.Cells[C_BD_2NDAMT,   i - 1] := oXL.Cells[i, j]
                     // 인원3
                     else if (j mod iCol_3thCount = 0) and (j div iCol_3thCount = 1) then
                        asg_BigData.Cells[C_BD_3THCNT,   i - 1] := oXL.Cells[i, j]
                     // 금액3
                     else if (j mod iCol_3thRcpamt = 0) and (j div iCol_3thRcpamt = 1) then
                        asg_BigData.Cells[C_BD_3THAMT,   i - 1] := oXL.Cells[i, j]
                     // 인원4
                     else if (j mod iCol_4thCount = 0) and (j div iCol_4thCount = 1) then
                        asg_BigData.Cells[C_BD_4THCNT,   i - 1] := oXL.Cells[i, j]
                     // 금액4
                     else if (j mod iCol_4thRcpamt = 0) and (j div iCol_4thRcpamt = 1) then
                        asg_BigData.Cells[C_BD_4THAMT,   i - 1] := oXL.Cells[i, j]
                     // 인원5
                     else if (j mod iCol_5thCount = 0) and (j div iCol_5thCount = 1) then
                        asg_BigData.Cells[C_BD_5THCNT,   i - 1] := oXL.Cells[i, j]
                     // 금액5
                     else if (j mod iCol_5thRcpamt = 0) and (j div iCol_5thRcpamt = 1) then
                        asg_BigData.Cells[C_BD_5THAMT,   i - 1] := oXL.Cells[i, j]
                     // 넘버1
                     else if (j mod iCol_No1 = 0) and (j div iCol_No1 = 1) then
                        asg_BigData.Cells[C_BD_NO1,   i - 1] := oXL.Cells[i, j]
                     // 넘버2
                     else if (j mod iCol_No2 = 0) and (j div iCol_No2 = 1) then
                        asg_BigData.Cells[C_BD_NO2,   i - 1] := oXL.Cells[i, j]
                     // 넘버3
                     else if (j mod iCol_No3 = 0) and (j div iCol_No3 = 1) then
                        asg_BigData.Cells[C_BD_NO3,   i - 1] := oXL.Cells[i, j]
                     // 넘버4
                     else if (j mod iCol_No4 = 0) and (j div iCol_No4 = 1) then
                        asg_BigData.Cells[C_BD_NO4,   i - 1] := oXL.Cells[i, j]
                     // 넘버5
                     else if (j mod iCol_No5 = 0) and (j div iCol_No5 = 1) then
                        asg_BigData.Cells[C_BD_NO5,   i - 1] := oXL.Cells[i, j]
                     // 넘버6
                     else if (j mod iCol_No6 = 0) and (j div iCol_No6 = 1) then
                        asg_BigData.Cells[C_BD_NO6,   i - 1] := oXL.Cells[i, j]
                     // 보너스
                     else if (j mod iCol_NoBonus = 0) and (j div iCol_NoBonus = 1) then
                        asg_BigData.Cells[C_BD_NOBONUS,   i - 1] := oXL.Cells[i, j];


                     // 해당 Thread가 CPU점유 하는 것을 방지위한 예외처리
                     Application.ProcessMessages;

                  end;
               end;
            end;

            // Update Loading Bar
            pb_DataLoading.StepIt;

         end;



         // RowCount 정리
         asg_BigData.RowCount := i - 1;

         //asg_BigData.EndUpdate;

         asg_BigData.Refresh;


         //--------------------------------------------------------------------
         // 3. Data Loading bar 숨기기
         //--------------------------------------------------------------------
         SetLoadingBar('OFF');


         // Comments
         //lb_RegDoc.Caption := '▶ ' + IntToStr(i) + '건의 엑셀업로드 내역을 D/B에 등록중입니다.';




         // 엑셀업로드 대상내역이 1건이상 존재시, D/B Insert 진행.
         if (i > 0) then
            InsBigDataList;

      end
      else if (ftc_Dialog.ActiveTab = AT_DOC) then
      begin
         //--------------------------------------------------------------------
         // 1. Data Loading bar 보이기
         //--------------------------------------------------------------------
         SetLoadingBar('ON');


         // Maximum value of progress status
         //pb_DataLoading.Max := XLRows;
         // 유효하게 100건 정도만 Upload 적용 @ 2017.10.25 LSH
         pb_DataLoading.Max := 100;


         //--------------------------------------------------------------------
         // 2-2. Excel Data Loading
         //--------------------------------------------------------------------
         //asg_Release.BeginUpdate;


//         iRow := 0;


         //for i := 1 to XLRows do
         // Title- Row만 먼저 변수값을 설정해 놓고, 유효한 Range만 아래에서 Searching & Upload 적용 @ 2017.10.25 LSH
         for i := 1 to 1 do
         begin
            //---------------------------------------------------------------------
            // 개발 릴리즈대장 Excel sheet의 Title 항목의 Index 값을 가져오기
            //---------------------------------------------------------------------
            if i = 1 then
            begin
               for k := 1 to XLCols do
               begin
                  if Trim(oXL.Cells[i, k]) = '작성일시' then
                     iCol_RegDate   := k
                  else if Trim(oXL.Cells[i, k]) = '분류' then
                     iCol_DutySpec  := k
                  else if Trim(oXL.Cells[i, k]) = '시스템' then
                     iCol_Context   := k
                  else if Trim(oXL.Cells[i, k]) = '클라이언트 소스' then
                     iCol_ClienSrc  := k
                  else if Trim(oXL.Cells[i, k]) = '클라이언트' then
                     iCol_Client    := k
                  else if Trim(oXL.Cells[i, k]) = '서버 소스' then
                     iCol_ServeSrc  := k
                  else if Trim(oXL.Cells[i, k]) = '서버' then
                     iCol_Server    := k
                  else if PosByte('관련 SR', Trim(oXL.Cells[i, k])) > 0 then
                     iCol_SrReqNo   := k
                  else if Trim(oXL.Cells[i, k]) = '담당자' then
                     iCol_DutyUser  := k
                  else if Trim(oXL.Cells[i, k]) = '요청자' then
                     iCol_ReqUser   := k
                  else if Trim(oXL.Cells[i, k]) = '관리자' then
                     iCol_CofmUser  := k
                  else if Trim(oXL.Cells[i, k]) = '테스트일시' then
                     iCol_TestDate  := k
                  else if Trim(oXL.Cells[i, k]) = '릴리즈 일시' then
                     iCol_ReleasDt  := k
                  else if Trim(oXL.Cells[i, k]) = '릴리즈' then
                     iCol_ReleasUser:= k
                  else if PosByte('작업경과', Trim(oXL.Cells[i, k])) > 0 then
                     iCol_Remark    := k;
               end;
            end;
         end;



         // Title-Row를 제외한 실제 유효한 마지막 Row-Index - 100 Line 정도만 업데이트 적용 @ 2017.10.25 LSH (업로드 속도개선 이슈)
         for i := XLLastUsedRows - 100 to XLLastUsedRows + 1 do
         begin
            for j := 1 to XLCols do
            begin
               if i > 1 then
               begin
                  //---------------------------------------------------------------
                  // 릴리즈대장 Excel sheet의 Title별 data를 Grid에 뿌려주기
                  //---------------------------------------------------------------
                  if (LengthByte(Trim(oXL.Cells[i, iCol_RegDate])) > 10)
                     {
                     and
                     (
                        (CopyByte(Trim(oXL.Cells[i, iCol_RegDate]), 1, 10) >= '2017-10-22') and
                        (CopyByte(Trim(oXL.Cells[i, iCol_RegDate]), 1, 10) <= '2017-10-25')
                     )
                     }
                  then
                  begin

                     // 작성일시
                     if (j mod iCol_RegDate = 0) and (j div iCol_RegDate = 1) then
                     begin
                       {
                        if (i > 1) and
                           (CopyByte(oXL.Cells[i, j], 1, 16) = Trim(asg_Release.Cells[C_RL_REGDATE  ,i-2])) then
                           asg_Release.Cells[C_RL_REGDATE, i - 1] := CopyByte(oXL.Cells[i, j], 1, 15) + IntToStr(StrToInt(CopyByte(oXL.Cells[i, j], 16, 1)) + 1)
                              //showmessage(CopyByte(Trim(Cells[C_RL_REGDATE  ,i]), 1, 16) + '--->' + CopyByte(Trim(Cells[C_RL_REGDATE  ,i]), 1, 15) + IntToStr(StrToInt(CopyByte(Trim(Cells[C_RL_REGDATE  ,i]), 16, 1)) + 1));
                           else
                           }

                        asg_Release.Cells[C_RL_REGDATE, i - 1] := CopyByte(oXL.Cells[i, j], 1, 16);

                        //asg_Release.Cells[C_RL_REGDATE, i - 1] := CopyByte(oXL.Cells[i, j], 1, 16)

                     end
                     // 분류
                     else if (j mod iCol_DutySpec = 0) and (j div iCol_DutySpec = 1) then
                        asg_Release.Cells[C_RL_DUTYSPEC,   i - 1] := oXL.Cells[i, j]
                     // 변경사항
                     else if (j mod iCol_Context = 0) and (j div iCol_Context = 1) then
                        asg_Release.Cells[C_RL_CONTEXT,   i - 1] := oXL.Cells[i, j]
                     // 관련 S/R
                     else if (j mod iCol_SrReqNo = 0) and (j div iCol_SrReqNo = 1) then
                        asg_Release.Cells[C_RL_SRREQNO,   i - 1] := oXL.Cells[i, j]
                     // 요청자
                     else if (j mod iCol_ReqUser = 0) and (j div iCol_ReqUser = 1) then
                        asg_Release.Cells[C_RL_REQUSER,   i - 1] := oXL.Cells[i, j]
                     // 담당자
                     else if (j mod iCol_DutyUser = 0) and (j div iCol_DutyUser = 1) then
                        asg_Release.Cells[C_RL_DUTYUSER,   i - 1] := oXL.Cells[i, j]
                     // 릴리즈일시
                     else if (j mod iCol_ReleasDt = 0) and (j div iCol_ReleasDt = 1) then
                        asg_Release.Cells[C_RL_RELEASDT,   i - 1] := CopyByte(oXL.Cells[i, j], 1, 16)
                     // 관리자
                     else if (j mod iCol_CofmUser = 0) and (j div iCol_CofmUser = 1) then
                        asg_Release.Cells[C_RL_COFMUSER,   i - 1] := oXL.Cells[i, j]
                     // 테스트일시
                     else if (j mod iCol_TestDate = 0) and (j div iCol_TestDate = 1) then
                        asg_Release.Cells[C_RL_TESTDATE,   i - 1] := CopyByte(oXL.Cells[i, j], 1, 16)
                     // Client Source
                     else if (j mod iCol_ClienSrc = 0) and (j div iCol_ClienSrc = 1) then
                        asg_Release.Cells[C_RL_CLIENSRC,   i - 1] := oXL.Cells[i, j]
                     // Client
                     else if (j mod iCol_Client = 0) and (j div iCol_Client = 1) then
                        asg_Release.Cells[C_RL_CLIENT,   i - 1] := oXL.Cells[i, j]
                     // Server Source
                     else if (j mod iCol_ServeSrc = 0) and (j div iCol_ServeSrc = 1) then
                        asg_Release.Cells[C_RL_SERVESRC,   i - 1] := oXL.Cells[i, j]
                     // Server
                     else if (j mod iCol_Server = 0) and (j div iCol_Server = 1) then
                        asg_Release.Cells[C_RL_SERVER,   i - 1] := oXL.Cells[i, j]
                     // 릴리즈 담당자
                     else if (j mod iCol_ReleasUser = 0) and (j div iCol_ReleasUser = 1) then
                        asg_Release.Cells[C_RL_RELEASUSER,   i - 1] := oXL.Cells[i, j]
                     // Remark
                     else if (j mod iCol_Remark = 0) and (j div iCol_Remark = 1) then
                        asg_Release.Cells[C_RL_REMARK,   i - 1] := oXL.Cells[i, j];


                     // 해당 Thread가 CPU점유 하는 것을 방지위한 예외처리
                     Application.ProcessMessages;

                  end;
               end;
            end;

            // Update Loading Bar
            pb_DataLoading.StepIt;

         end;



         // RowCount 정리
         asg_Release.RowCount := i - 2;

         //asg_Release.EndUpdate;

         asg_Release.Refresh;


         //--------------------------------------------------------------------
         // 3. Data Loading bar 숨기기
         //--------------------------------------------------------------------
         SetLoadingBar('OFF');


         // Comments
         lb_RegDoc.Caption := '▶ ' + IntToStr(i) + '건의 엑셀업로드 내역을 D/B에 등록중입니다.';


         // 엑셀업로드 대상내역이 1건이상 존재시, D/B Insert 진행.
         if (i > 0) then
            InsReleaseList;

      end;



      // Excel Close
      oXL.WorkBooks.Close;


      // 엑셀 업로드 완료시 주요 Timer On @ 2014.07.23 LSH
      tm_DialPop.Enabled := True;
      tm_TxInit.Enabled  := True;


   except
      // 엑셀 업로드 예외처리시 주요 Timer On @ 2014.07.23 LSH
      tm_DialPop.Enabled := True;
      tm_TxInit.Enabled  := True;

      //SetLoadingBar('OFF');
      oXL.WorkBooks.Close;
   end;
end;




procedure TMainDlg.apn_MenuClose(Sender: TObject);
begin
   apn_Menu.Visible := False;
end;



//------------------------------------------------------------------------------
// [식단등록] AdvStringGrid onButtonClick Event Handler
//    - 사용안함.
//
// Author : Lee, Se-Ha
// Date   : 2013.09.16
//------------------------------------------------------------------------------
procedure TMainDlg.asg_MenuBarButtonClick(Sender: TObject; ACol,
  ARow: Integer);
//var
//   TpSetMenu   : TTpSvc;
//   i, j,  iCnt     : Integer;
//   vType1,
//   vLocate,
//   vUserId,
//   vUserNm,
//   vDiet1,
//   vDiet2,
//   vDiet3,
//   vEditIp : Variant;
//   sType, sLocateNm,
//   sDietDay,
//   sDayGubun,
//   sDiet1   ,
//   sDiet2   ,
//   sDiet3   : String;

begin
{
   // 직원식 메뉴 Update
   if (ACol = 1) then
   begin


      sType := '7';

      iCnt  := 0;


      // 접속자 IP 식별후, 근무처 Assign
      if PosByte('안암도메인', FsUserIp) > 0 then
         sLocateNm := '안암'
      else if PosByte('구로도메인', FsUserIp) > 0 then
         sLocateNm := '구로'
      else if PosByte('안산도메인', FsUserIp) > 0 then
         sLocateNm := '안산';



      //-------------------------------------------------------------------
      // 2. Create Variables
      //-------------------------------------------------------------------
      with asg_Menu do
      begin
         for j := 0 to ColCount  -1 do
         begin

            sDiet1 := '';
            sDiet2 := '';
            sDiet3 := '';

            for i := 0 to RowCount - 1 do
            begin
               // 식단 일자
               sDietDay   := CopyByte(Cells[j, 0], 1, 2);

               // 식단 요일
               sDayGubun  := CopyByte(Cells[j, 0], 4, 2);

               // 아침
               if (i >= 1) and (i <= 6) then
               begin
                  if (i = 1) then
                     sDiet1 := Cells[j, i]
                  else if Cells[j, i] <> '' then
                     sDiet1 := sDiet1 + ',' + Cells[j, i];
               end;

               // 점심
               if (i >= 7) and (i <= 12) then
               begin
                  if (i = 7) then
                     sDiet2 := Cells[j, i]
                  else if Cells[j, i] <> '' then
                     sDiet2 := sDiet2 + ',' + Cells[j, i];
               end;

               // 저녁
               if (i >= 13) then
               begin
                  if (i = 13) then
                     sDiet3 := Cells[j, i]
                  else if Cells[j, i] <> '' then
                     sDiet3 := sDiet3 + ',' + Cells[j, i];
               end;
            end;


            //------------------------------------------------------------------
            // 2-1. Append Variables
            //------------------------------------------------------------------
            AppendVariant(vType1   ,   sType                  );
            AppendVariant(vLocate  ,   sLocateNm              );  // 소속
            AppendVariant(vUserID  ,   sDietDay               );  // 일자(dd)
            AppendVariant(vUserNm  ,   sDayGubun              );  // 요일
            AppendVariant(vDiet1   ,   sDiet1                 );  // 아침
            AppendVariant(vDiet2   ,   sDiet2                 );  // 점심
            AppendVariant(vDiet3   ,   sDiet3                 );  // 저녁
            AppendVariant(vEditIp  ,   FsUserIp               );


            Inc(iCnt);
         end;









      end;




      if iCnt <= 0 then
         Exit;



      //-------------------------------------------------------------------
      // 3. Insert by TpSvc
      //-------------------------------------------------------------------
      TpSetMenu := TTpSvc.Create;
      TpSetMenu.Init(Self);



      Screen.Cursor := crHourGlass;



      try
         if TpSetMenu.PutSvc('MD_KUMCM_M1',
                             [
                              'S_TYPE1'   , vType1
                            , 'S_STRING1' , vLocate
                            , 'S_STRING2' , vUserId
                            , 'S_STRING3' , vUserNm
                            , 'S_STRING12', vDiet1
                            , 'S_STRING37', vDiet2
                            , 'S_STRING38', vDiet3
                            , 'S_STRING8' , vEditIp
                            ] ) then
         begin
            MessageBox(self.Handle,
                       PChar('주간 식단정보가 정상적으로 [업데이트]되었습니다.'),
                       '[KUMC 다이얼로그] 주간 식단정보 업데이트 알림 ',
                       MB_OK + MB_ICONINFORMATION);

         end
         else
         begin
            ShowMessageM(GetTxMsg);
         end;

      finally
         FreeAndNil(TpSetMenu);
         Screen.Cursor  := crDefault;
      end;
   end;

}

end;



procedure TMainDlg.FlatSpeedButton8Click(Sender: TObject);
//var
//   sFileExt : String;
begin
{
   asg_Menu.ClearRows(1, asg_Menu.RowCount);
   asg_Menu.RowCount := 2;

   //--------------------------------------------------------------------
   // 1. 파일 오픈(.CSV, .XLS) 및 Data Loading
   //--------------------------------------------------------------------
   if opendlg_File.Execute then
   begin
      if opendlg_File.FileName = '' then
        Exit;


      sFileExt  := AnsiUpperCase(ExtractFileExt(opendlg_File.FileName));


      // Excel 파일인 경우, 별도 Procedure (LoadExcelFile) Call
      if (sFileExt = '.XLSX') or (sFileExt = '.XLS') Then
      begin
         if opendlg_File.FileName <> '' then
         begin
            LoadExcelFile(opendlg_File.FileName);
         end;
      end
      else
      // CSV 는 아래 로직 진행.
      begin
         // 준비중..
      end;
   end;
}

end;

procedure TMainDlg.FlatSpeedButton9Click(Sender: TObject);
//var
//   TpSetMenu   : TTpSvc;
//   i, j,  iCnt     : Integer;
//   vType1,
//   vLocate,
//   vUserId,
//   vUserNm,
//   vDiet1,
//   vDiet2,
//   vDiet3,
//   vEditIp : Variant;
//   sType, sLocateNm,
//   sDietDay,
//   sDayGubun,
//   sDiet1   ,
//   sDiet2   ,
//   sDiet3   : String;

begin
{
   // 직원식 메뉴 Update
   sType := '7';

   iCnt  := 0;


   // 접속자 IP 식별후, 근무처 Assign
   if PosByte('안암', fcb_MenuLoc.Text) > 0 then
      sLocateNm := '안암'
   else if PosByte('구로', fcb_MenuLoc.Text) > 0 then
      sLocateNm := '구로'
   else if PosByte('안산', fcb_MenuLoc.Text) > 0 then
      sLocateNm := '안산';



   //-------------------------------------------------------------------
   // 2. Create Variables
   //-------------------------------------------------------------------
   with asg_Menu do
   begin
      for j := 0 to ColCount  -1 do
      begin

         sDiet1 := '';
         sDiet2 := '';
         sDiet3 := '';

         // 병원별로 영양팀 xls 포맷이 다르다.. -_-;
         // 병원별 로직 분기
         if (sLocateNm = '안암') then
         begin
            for i := 0 to RowCount - 1 do
            begin
               // 식단 일자
               sDietDay   := CopyByte(Cells[j, 0], 1, 2);

               // 식단 요일
               sDayGubun  := CopyByte(Cells[j, 0], 4, 2);

               // 아침
               if (i >= 1) and (i <= 6) then
               begin
                  if (i = 1) then
                     sDiet1 := Cells[j, i]
                  else if Cells[j, i] <> '' then
                     sDiet1 := sDiet1 + ',' + Cells[j, i];
               end;

               // 점심
               if (i >= 7) and (i <= 12) then
               begin
                  if (i = 7) then
                     sDiet2 := Cells[j, i]
                  else if Cells[j, i] <> '' then
                     sDiet2 := sDiet2 + ',' + Cells[j, i];
               end;

               // 저녁
               if (i >= 13) then
               begin
                  if (i = 13) then
                     sDiet3 := Cells[j, i]
                  else if Cells[j, i] <> '' then
                     sDiet3 := sDiet3 + ',' + Cells[j, i];
               end;
            end;
         end
         else if (sLocateNm = '구로') then
         begin
            for i := 0 to RowCount - 1 do
            begin
               // 식단 일자
               sDietDay   := CopyByte(Cells[j, 0], 1, 2);

               // 식단 요일
               sDayGubun  := CopyByte(Cells[j, 0], 4, 2);

               // 아침
               if (i >= 1) and (i <= 6) then
               begin
                  if (i = 1) then
                     sDiet1 := Cells[j, i]
                  else if Cells[j, i] <> '' then
                     sDiet1 := sDiet1 + ',' + Cells[j, i];
               end;

               // 점심
               if (i >= 7) and (i <= 12) then
               begin
                  if (i = 7) then
                     sDiet2 := Cells[j, i]
                  else if Cells[j, i] <> '' then
                     sDiet2 := sDiet2 + ',' + Cells[j, i];
               end;

               // 저녁
               if (i >= 13) then
               begin
                  if (i = 13) then
                     sDiet3 := Cells[j, i]
                  else if Cells[j, i] <> '' then
                     sDiet3 := sDiet3 + ',' + Cells[j, i];
               end;
            end;
         end
         else if (sLocateNm = '안산') then
         begin
            for i := 0 to RowCount - 1 do
            begin
               // 식단 일자
               sDietDay   := CopyByte(Cells[j, 0], 1, 2);

               // 식단 요일
               sDayGubun  := CopyByte(Cells[j, 0], 4, 2);

               // 아침
               if (i >= 1) and (i <= 4) then
               begin
                  if (i = 1) then
                     sDiet1 := Cells[j, i]
                  else if Cells[j, i] <> '' then
                     sDiet1 := sDiet1 + ',' + Cells[j, i];
               end;

               // 점심
               if (i >= 5) and (i <= 8) then
               begin
                  if (i = 5) then
                     sDiet2 := Cells[j, i]
                  else if Cells[j, i] <> '' then
                     sDiet2 := sDiet2 + ',' + Cells[j, i];
               end;

               // 저녁
               if (i >= 9) then
               begin
                  if (i = 9) then
                     sDiet3 := Cells[j, i]
                  else if Cells[j, i] <> '' then
                     sDiet3 := sDiet3 + ',' + Cells[j, i];
               end;
            end;
         end;




         //------------------------------------------------------------------
         // 2-1. Append Variables
         //------------------------------------------------------------------
         AppendVariant(vType1   ,   sType                  );
         AppendVariant(vLocate  ,   sLocateNm              );  // 소속
         AppendVariant(vUserID  ,   sDietDay               );  // 일자(dd)
         AppendVariant(vUserNm  ,   sDayGubun              );  // 요일
         AppendVariant(vDiet1   ,   sDiet1                 );  // 아침
         AppendVariant(vDiet2   ,   sDiet2                 );  // 점심
         AppendVariant(vDiet3   ,   sDiet3                 );  // 저녁
         AppendVariant(vEditIp  ,   FsUserIp               );


         Inc(iCnt);
      end;

   end;




   if iCnt <= 0 then
      Exit;



   //-------------------------------------------------------------------
   // 3. Insert by TpSvc
   //-------------------------------------------------------------------
   TpSetMenu := TTpSvc.Create;
   TpSetMenu.Init(Self);



   Screen.Cursor := crHourGlass;



   try
      if TpSetMenu.PutSvc('MD_KUMCM_M1',
                          [
                           'S_TYPE1'   , vType1
                         , 'S_STRING1' , vLocate
                         , 'S_STRING2' , vUserId
                         , 'S_STRING3' , vUserNm
                         , 'S_STRING12', vDiet1
                         , 'S_STRING37', vDiet2
                         , 'S_STRING38', vDiet3
                         , 'S_STRING8' , vEditIp
                         ] ) then
      begin

         MessageBox(self.Handle,
                    PChar('[' + sLocateNm + '] 주간 식단정보가 정상적으로 [업데이트]되었습니다.'),
                    '[KUMC 다이얼로그] 주간 식단정보 업데이트 알림 ',
                    MB_OK + MB_ICONINFORMATION);


      end
      else
      begin
         ShowMessageM(GetTxMsg);
      end;

   finally
      FreeAndNil(TpSetMenu);
      Screen.Cursor  := crDefault;
   end;
}

end;

procedure TMainDlg.asg_DialScanKeyPress(Sender: TObject; var Key: Char);
begin

      if (asg_DialScan.Cells[C_DS_KEYWORD, asg_DialScan.Row] <> '') and (Key = #13) then
      begin
         asg_DialScanDblClick(Sender);
      end

end;

procedure TMainDlg.asg_DialScanKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
{
   if (Key = 13) then

      if (asg_DialScan.Cells[1, R_DS_SCAN] <> '') then
      begin
         //
         SelGridInfo('DIALLIST');


         asg_DialScan.SetFocus;
         asg_DialScan.Navigation.AdvanceOnEnter := False;
      end
      else
         asg_DialScan.Navigation.AdvanceOnEnter := True;
         }
end;

procedure TMainDlg.asg_DialScanClickCell(Sender: TObject; ARow,
  ACol: Integer);
begin
{
   if (ACol = 1) and (ARow = R_DS_SCAN) then
      asg_DialScan.Cells[1, R_DS_SCAN] := '';
}
end;

procedure TMainDlg.fed_ScanKeyPress(Sender: TObject; var Key: Char);
var
   varResult : String;
begin
   if (fed_Scan.Text <> '') and (Key = #13) then
   begin
      if apn_LinkList.Visible then
      begin
         apn_LinkList.Collaps := True;
         apn_LinkList.Visible := False;
      end;


      // Refresh
      SelGridInfo('DIALLIST');


      // 로그 Update
      UpdateLog('DIALBOOK',
                'KEYWORD',
                FsUserIp,
                'F',
                Trim(fed_Scan.Text),
                '',
                varResult
                );

      asg_DialScan.Visible := False;
   end;

end;

procedure TMainDlg.fed_ScanClick(Sender: TObject);
begin
   // 한글입력 초기화 적용 @ 2015.06.03 LSH
   HanKeyChg(fed_Scan.Handle);

   fed_Scan.Text := '';

   if (apn_AddMyDial.Visible) then
   begin
      apn_AddMyDial.Collaps := True;
      apn_AddMyDial.Visible := False;
   end;

   if (apn_DialMap.Visible) then
   begin
      apn_DialMap.Collaps := True;
      apn_DialMap.Visible := False;
   end;
end;

procedure TMainDlg.fed_ScanChange(Sender: TObject);
begin

   asg_DialScan.Visible := True;

   lb_DialScan.Caption := '▶ 다이얼Book 검색 입력대기중....';

   // 최근검색 Key 조회
   SelGridInfo('DIALSCAN');

end;

procedure TMainDlg.FlatSpeedButton12Click(Sender: TObject);
//var
//   tmpLinkSeq : String;
//   tmpUserNm  : String;
begin
   if not IsLogonUser then
   begin
      MessageBox(self.Handle,
                 PChar('[비상연락망] - [담당자정보]에 사용자 IP 등록후 사용하실 수 있습니다.'),
                 PChar(Self.Caption + ' : 등록되지 않은 User 제한 알림'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;





   with asg_DialList do
   begin

      if (Cells[C_DL_LOCATE, Row] = '') then
      begin
         MessageBox(self.Handle,
                    PChar('등록하실 연락처 정보가 선택되지 않았습니다.'),
                    PChar('나의 다이얼 Book 등록 오류 알림'),
                    MB_OK + MB_ICONWARNING);

         Exit;
      end;

   {
      if CopyByte(Cells[C_DL_LINKSEQ, Row], 1, 4) = '(02)' then
         tmpLinkSeq := CopyByte(Cells[C_DL_LINKSEQ, Row], 5, 9)
      else
         tmpLinkSeq := Trim(Cells[C_DL_LINKSEQ, Row]);
   }




      { --  개인 D/B 등록은, data의 품질을 저하시킬 수 있는 문제가 있음.
        --  이 문제가 개선되기 전까지는 기존 D/B (전화번호부/다이얼로그)만 등록되도록 함.
      if (Cells[C_DL_LINKDB, Row] = '전화번호부') or
         (Cells[C_DL_LINKDB, Row] = '다이얼로그') then
         tmpUserNm := FsUserNm
      else
         tmpUserNm := Trim(Cells[C_DL_LINKDB, Row]);
      }


      if (apn_LinkList.Visible) then
         // 현재 링크된 List에서 나의 다이얼 Book에 선택한 연락처 등록
         UpdateMyDial('8',
                      Cells[C_DL_LOCATE,     Row],
                      FsUserIp,
                      Cells[C_DL_DEPTNM,     Row],
                      Cells[C_DL_DEPTSPEC,   Row],
                      Cells[C_DL_CALLNO,     Row],
                      asg_LinkList.Cells[C_LL_DUTYUSER,   asg_LinkList.Row],
                      asg_LinkList.Cells[C_LL_MOBILE,     asg_LinkList.Row],
                      asg_LinkList.Cells[C_LL_DUTYRMK,    asg_LinkList.Row],
                      Cells[C_DL_LINKDB,     Row],
                      Cells[C_DL_LINKSEQ,    Row],
                      '',
                      Cells[C_DL_DTYUSRID,   Row]        // 등록된 UserId @ 2015.04.13 LSH
                      )


      else
         // 현재 검색 List에서 나의 다이얼 Book에 선택한 연락처 등록
         UpdateMyDial('8',
                      Cells[C_DL_LOCATE,     Row],
                      FsUserIp,
                      Cells[C_DL_DEPTNM,     Row],
                      Cells[C_DL_DEPTSPEC,   Row],
                      Cells[C_DL_CALLNO,     Row],
                      Cells[C_DL_DUTYUSER,   Row],
                      '',
                      '',
                      Cells[C_DL_LINKDB,     Row],
                      Cells[C_DL_LINKSEQ,    Row],
                      '',
                      Cells[C_DL_DTYUSRID,   Row]        // 등록된 UserId @ 2015.04.13 LSH
                      );
   end;


   if apn_LinkList.Visible then
   begin
      apn_LinkList.Collaps := True;
      apn_LinkList.Visible := False;
   end;


   // Refresh
   SelGridInfo('DIALLIST');



end;

procedure TMainDlg.FlatSpeedButton13Click(Sender: TObject);
begin

   if not IsLogonUser then
   begin
      MessageBox(self.Handle,
                 PChar('[비상연락망] - [담당자정보]에 사용자 IP 등록후 사용하실 수 있습니다.'),
                 PChar(Self.Caption + ' : 등록되지 않은 User 제한 알림'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;




   with asg_MyDial do
   begin

      if (Cells[C_MD_LOCATE, Row] = '') then
      begin
         MessageBox(self.Handle,
                    PChar('삭제하실 연락처 정보가 선택되지 않았습니다.'),
                    PChar('나의 다이얼 Book 삭제 오류 알림'),
                    MB_OK + MB_ICONWARNING);

         Exit;
      end;

      // 나의 다이얼 Book에 선택한 연락처 삭제
      UpdateMyDial('8',
                   Cells[C_MD_LOCATE,     Row],
                   FsUserIp,
                   '',
                   '',
                   '',
                   '',
                   '',
                   '',
                   Cells[C_MD_LINKDB,     Row],
                   Cells[C_MD_LINKSEQ,    Row],
                   FormatDateTime('yyyy-mm-dd', DATE),
                   ''
                   );
   end;


   if apn_LinkList.Visible then
   begin
      apn_LinkList.Collaps := True;
      apn_LinkList.Visible := False;
   end;


   // Refresh
   if (Trim(fed_Scan.Text) <> '') then
      SelGridInfo('DIALLIST');

{
   with asg_MyDial do
   begin
      if (RowCount = 2) then
         ClearRows(1, RowCount)
      else
         RemoveRows(Row, 1);
   end;
}
end;

procedure TMainDlg.FlatSpeedButton14Click(Sender: TObject);
begin
   //
   SelGridInfo('MYDIAL');
end;

procedure TMainDlg.asg_DialListGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   // Horizontal Alignment
   if (ARow = 0) or ((ARow > 0) and ((ACol = C_DL_LOCATE) or
                                     (ACol = C_DL_DUTYUSER) or
                                     (ACol = C_DL_LINKDB) or
                                     (ACol = C_DL_LINKCNT))
                                     ) then
      HAlign := taCenter;

   // Vertical Alignment
   VAlign := vtaCenter;
end;

procedure TMainDlg.asg_MyDialGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   // Horizontal Alignment
   if (ARow = 0) or ((ARow > 0) and ((ACol = C_MD_LOCATE) or
                                     (ACol = C_MD_DUTYUSER) or
                                     (ACol = C_MD_MOBILE))
                                     ) then
      HAlign := taCenter;

   // Vertical Alignment
   VAlign := vtaCenter;
end;



//---------------------------------------------------------------------------
// [업데이트] 나의 다이얼 Book Update
//
// Date     : 2013.09.26
// Author   : Lee, Se-Ha
//---------------------------------------------------------------------------
procedure TMainDlg.UpdateMyDial(in_Type,     in_Locate,     in_UserIp,  in_DeptNm,  in_DeptSpec,
                                in_CallNo,   in_DutyUser,   in_Mobile,  in_DutyRmk, in_DataBase,
                                in_LinkSeq,  in_DelDate,    in_DtyUsrId : String);
var
   vType1   ,
   vLocate  ,
   vUserIp  ,
   vDeptNm  ,
   vDeptSpec,
   vCallNo  ,
   vDutyUser,
   vMobile  ,
   vDutyRmk ,
   vDataBase,
   vLinkSeq ,
   vDtyUsrId,
   vDelDate  : Variant;
   TpPutDial : TTpSvc;
   tmp_DutyUser : String;
begin

   // User 이름 filetering @ 2015.04.15 LSH
   if PosByte('[', in_DutyUser) > 0 then
      tmp_DutyUser := CopyByte(in_DutyUser, 1, PosByte('[', in_DutyUser) - 2)
   else
      tmp_DutyUser := in_DutyUser;


   //------------------------------------------------------------------
   // 2-1. Append Variables
   //------------------------------------------------------------------
   AppendVariant(vType1   ,  in_Type);
   AppendVariant(vLocate  ,  in_Locate);
   AppendVariant(vUserIp  ,  in_UserIp);
   AppendVariant(vDeptNm  ,  in_DeptNm);
   AppendVariant(vDeptSpec,  in_DeptSpec);
   AppendVariant(vCallNo  ,  in_CallNo);
   AppendVariant(vDutyUser,  tmp_DutyUser);
   AppendVariant(vMobile  ,  in_Mobile);
   AppendVariant(vDutyRmk ,  in_DutyRmk);
   AppendVariant(vDataBase,  in_DataBase);
   AppendVariant(vLinkSeq ,  in_LinkSeq);
   AppendVariant(vDelDate ,  in_DelDate);
   AppendVariant(vDtyUsrId,  in_DtyUsrId);   // 등록된 UserId @ 2015.04.13 LSH



   //-------------------------------------------------------------------
   // 3. Insert by TpSvc
   //-------------------------------------------------------------------
   TpPutDial := TTpSvc.Create;
   TpPutDial.Init(Self);



   Screen.Cursor := crHourGlass;



   try
      if TpPutDial.PutSvc('MD_KUMCM_M1',
                          [
                            'S_TYPE1'   , vType1
                         ,  'S_STRING1' , vLocate
                         ,  'S_STRING2' , vDtyUsrId
                         ,  'S_STRING7' , vUserIp
                         ,  'S_STRING42', vDeptNm
                         ,  'S_STRING43', vDeptSpec
                         ,  'S_STRING15', vCallNo
                         ,  'S_STRING13', vDutyUser
                         ,  'S_STRING5' , vMobile
                         ,  'S_STRING12', vDutyRmk
                         ,  'S_STRING44', vDataBase
                         ,  'S_STRING45', vLinkSeq
                         ,  'S_STRING10', vDelDate
                         ] ) then
      begin

         if (in_Type = '8') then
         begin
            if (in_DelDate = '') then
               //lb_MyDial.Caption := '연락처 [등록] 완료'


               MessageBox(self.Handle,
                          PChar('선택하신 연락처가 나의 다이얼 Book에 [등록] 되었습니다.'),
                          '[KUMC 다이얼로그] 다이얼 Book 업데이트 알림 ',
                          MB_OK + MB_ICONINFORMATION)

            else
               //lb_MyDial.Caption := '연락처 [삭제] 완료';

               MessageBox(self.Handle,
                          PChar('선택하신 연락처가 나의 다이얼 Book에서 [삭제] 되었습니다.'),
                          '[KUMC 다이얼로그] 다이얼 Book 업데이트 알림 ',
                          MB_OK + MB_ICONINFORMATION);

         end;


         // Refresh
         SelGridInfo('MYDIAL');

      end
      else
      begin
         ShowMessageM(GetTxMsg);
      end;



   finally
      FreeAndNil(TpPutDial);
      Screen.Cursor  := crDefault;
   end;
end;





procedure TMainDlg.asg_MyDialEditingDone(Sender: TObject);
begin
   // D/B 출처에 따른 Update 버튼 유무 (각 ID별 D/B 관리 시행않는한, 큰 의미 없음)
   if (asg_MyDial.Cells[C_MD_LINKDB, asg_MyDial.Row] = '전화번호부') or
      (asg_MyDial.Cells[C_MD_LINKDB, asg_MyDial.Row] = '다이얼로그') or
      (asg_MyDial.Cells[C_MD_LINKDB, asg_MyDial.Row] = '홈페이지') or
      (asg_MyDial.Cells[C_MD_LINKDB, asg_MyDial.Row] = 'S/R (MIS)') or
      (PosByte(FsUserNm, asg_MyDial.Cells[C_MD_LINKDB, asg_MyDial.Row]) > 0) then

      asg_MyDial.AddButton(C_MD_REMARK,  asg_MyDial.Row, asg_MyDial.ColWidths[C_MD_REMARK]-5,  20, 'Update', haBeforeText, vaCenter);           // Update

end;



procedure TMainDlg.asg_MyDialButtonClick(Sender: TObject; ACol,
  ARow: Integer);

begin
   if (ACol = C_MD_REMARK) then
   begin
      with asg_MyDial do
      begin
         if (Cells[C_MD_LOCATE, Row] = '') then
         begin
            MessageBox(self.Handle,
                       PChar('등록하실 연락처 정보가 선택되지 않았습니다.'),
                       PChar('나의 다이얼 Book 등록 오류 알림'),
                       MB_OK + MB_ICONWARNING);

            Exit;
         end;



         // 나의 다이얼 Book에 선택한 연락처 변경사항 등록
         UpdateMyDial('8',
                      Cells[C_MD_LOCATE,     Row],
                      FsUserIp,
                      Cells[C_MD_DEPTNM,     Row],
                      Cells[C_MD_DEPTSPEC,   Row],
                      Cells[C_MD_CALLNO,     Row],
                      Cells[C_MD_DUTYUSER,   Row],
                      Cells[C_MD_MOBILE,     Row],
                      Cells[C_MD_REMARK,     Row],
                      Cells[C_MD_LINKDB,     Row],
                      Cells[C_MD_LINKSEQ,    Row],
                      '',
                      Cells[C_MD_DTYUSRID,   Row]
                      );
      end;


      if apn_LinkList.Visible then
      begin
         apn_LinkList.Collaps := True;
         apn_LinkList.Visible := False;
      end;

      // Refresh
      if (Trim(fed_Scan.Text) <> '') then
         SelGridInfo('DIALLIST');



   end;
end;

procedure TMainDlg.asg_MyDialCanEditCell(Sender: TObject; ARow,
  ACol: Integer; var CanEdit: Boolean);
begin

   // 소속, 부서명, 부서상세는 Edit 제한
   CanEdit := (ACol <> C_MD_LOCATE) and
              (ACol <> C_MD_DEPTNM) and
              (ACol <> C_MD_DEPTSPEC);

{
   if (ARow = asg_MyDial.Row) then
   begin

      // 소속, 부서명, 부서상세는 Edit 제한
      CanEdit := (ACol <> C_MD_LOCATE) and
                 (ACol <> C_MD_DEPTNM) and
                 (ACol <> C_MD_DEPTSPEC);

      // D/B 출처에 따른 수정 제한 (사실, 각 ID별 D/B 관리 시행하지 않는한 큰 의미는 없음)
      if (asg_MyDial.Cells[C_MD_LINKDB, asg_MyDial.Row] = '전화번호부') or
         (asg_MyDial.Cells[C_MD_LINKDB, asg_MyDial.Row] = '다이얼로그') or
         (asg_MyDial.Cells[C_MD_LINKDB, asg_MyDial.Row] = FsUserNm) then
         CanEdit := (ACol = C_MD_CALLNO) or
                    (ACol = C_MD_MOBILE) or
                    (ACol = C_MD_DUTYUSER) or
                    (ACol = C_MD_REMARK)
      else
         CanEdit := (ACol <> C_MD_LOCATE) and
                    (ACol <> C_MD_DEPTNM) and
                    (ACol <> C_MD_DEPTSPEC) and
                    (ACol <> C_MD_CALLNO) and
                    (ACol <> C_MD_MOBILE) and
                    (ACol <> C_MD_DUTYUSER) and
                    (ACol <> C_MD_REMARK);

   end;
}


end;



procedure TMainDlg.asg_MyDialKeyPress(Sender: TObject; var Key: Char);
begin
   if (Key = #13) then
   begin
      if asg_MyDial.Col = C_MD_MOBILE then
         asg_MyDial.Navigation.AdvanceOnEnter := False
      else
         asg_MyDial.Navigation.AdvanceOnEnter := True;
   end;
end;

procedure TMainDlg.asg_DialListClickCell(Sender: TObject; ARow,
  ACol: Integer);
begin
   if apn_LinkList.Visible then
   begin
      apn_LinkList.Collaps := True;
      apn_LinkList.Visible := False;
   end;

   with asg_DialList do
   begin
      if (ACol = C_DL_LINKCNT) then
      begin

         if (Cells[C_DL_LINKCNT, ARow] <> '0') then
            GetLinkList('DIALBOOK',
                        Cells[C_DL_LOCATE,   Row],
                        Cells[C_DL_LINKDB,   Row],
                        Cells[C_DL_LINKSEQ,  Row],
                        FsUserIp);



      end
      else
      begin
         apn_LinkList.Collaps := True;
         apn_LinkList.Visible := False;
      end;
   end;
end;





//------------------------------------------------------------------------------
// [조회] 다이얼 Book / 업무(S/R) 분석 링크 List 상세조회
//
// Author : Lee, Se-Ha
// Date   : 2013.09.27
//------------------------------------------------------------------------------
procedure TMainDlg.GetLinkList(  in_Gubun,
                                 in_DialLocate,
                                 in_DialLinkDb,
                                 in_DialLinkSeq,
                                 in_UserIp : String);
var
   TpGetLink      : TTpSvc;
   sType, sLinkDb : String;
   i, iRowCnt     : Integer;
begin


   // 검색 구분
   if (in_Gubun = 'DIALBOOK') then
   begin
      sType := '19';

      if (PosByte('Book', asg_DialList.Cells[C_DL_LINKDB, asg_DialList.Row]) > 0) and
         (asg_DialList.Cells[C_DL_LOCATE, asg_DialList.Row] <> '협력업체') then
         sLinkDb := '전화번호부'
      else
         sLinkDb := in_DialLinkDb;

      asg_LinkList.ClearRows(1, asg_LinkList.RowCount);
      asg_LinkList.RowCount := 2;
   end
   else if (in_Gubun = 'ANALYSIS') then
   begin
      sType := '23';

      asg_AnalList.ClearCols(1, asg_AnalList.ColCount);
      asg_AnalList.RowCount := 7;
   end;




   //-----------------------------------------------------------------
   // 2-1. Load Data by TpSvc
   //-----------------------------------------------------------------
   TpGetLink := TTpSvc.Create;
   TpGetLink.Init(Self);


   Screen.Cursor := crHourglass;


   try
      TpGetLink.CountField  := 'S_CODE6';
      TpGetLink.ShowMsgFlag := False;

      if TpGetLink.GetSvc('MD_KUMCM_S1',
                          ['S_TYPE1'  , sType
                         , 'S_TYPE3'  , sLinkDb
                         , 'S_TYPE4'  , in_DialLinkSeq
                         , 'S_TYPE5'  , in_DialLocate
                          ],
                          [
                           'S_CODE1'   , 'sLocate'
                         , 'S_CODE3'   , 'sUserNm'
                         , 'S_CODE6'   , 'sDutyRmk'
                         , 'S_CODE7'   , 'sMobile'
                         , 'S_CODE8'   , 'sEmail'
                         , 'S_CODE15'  , 'sFlag'
                         , 'S_CODE10'  , 'sDutyUser'
                         , 'S_CODE12'  , 'sCallNo'
                         , 'S_CODE21'  , 'sDocTitle'
                         , 'S_CODE22'  , 'sRegDate'
                         , 'S_CODE24'  , 'sRelDept'
                         , 'S_CODE25'  , 'sDocRmk'
                         , 'S_STRING4' , 'sContext'
                         , 'S_STRING5' , 'sAttachNm'
                         , 'S_STRING6' , 'sHideFile'
                         , 'S_STRING16', 'sDeptNm'
                         , 'S_STRING17', 'sDeptSpec'
                         , 'S_STRING18', 'sDataBase'
                         , 'S_STRING20', 'sLinkSeq'
                         , 'S_STRING21', 'sReqNo'
                          ]) then

         if TpGetLink.RowCount <= 0 then
         begin
            Exit;
         end;


      //---------------------------------------------------
      // 다이얼 Book : 링크된 리스트 조회
      //---------------------------------------------------
      if (in_Gubun = 'DIALBOOK') then
      begin

         with asg_LinkList do
         begin

            iRowCnt  := TpGetLink.RowCount;
            RowCount := iRowCnt + FixedRows + 1;

            for i := 0 to iRowCnt - 1 do
            begin
               Cells[C_LL_USERNM,   i + FixedRows] := TpGetLink.GetOutputDataS('sUserNm',     i);
               Cells[C_LL_CALLNO,   i + FixedRows] := TpGetLink.GetOutputDataS('sCallNo',     i);
               Cells[C_LL_DUTYUSER, i + FixedRows] := TpGetLink.GetOutputDataS('sDutyUser',   i);
               Cells[C_LL_MOBILE,   i + FixedRows] := TpGetLink.GetOutputDataS('sMobile',     i);
               Cells[C_LL_DUTYRMK,  i + FixedRows] := TpGetLink.GetOutputDataS('sDutyRmk',    i);
               Cells[C_LL_LINKDB,   i + FixedRows] := TpGetLink.GetOutputDataS('sDataBase',   i);
               Cells[C_LL_LINKSEQ,  i + FixedRows] := TpGetLink.GetOutputDataS('sLinkSeq',    i);
               Cells[C_LL_DTYUSRID, i + FixedRows] := TpGetLink.GetOutputDataS('sUserId',     i);  // 등록된 UserId @ 2015.04.13 LSH
            end;

            RowCount := RowCount - 1;

         end;



         // Comments
         apn_LinkList.Caption.Text := IntToStr(iRowCnt) + '명이 해당 연락처를 아래와 같이 [나의 다이얼 Book]에 등록/사용중 입니다.';


         // Location & Display
         if (asg_DialList.Height - asg_DialList.CellRect(C_DL_LINKCNT, asg_DialList.Row).Top) > (apn_LinkList.Height + 70) then
         begin
            apn_LinkList.Top        := asg_DialList.CellRect(C_DL_LINKCNT, asg_DialList.Row).Top + 22;
            apn_LinkList.Left       := 0;
         end
         else
         begin
            apn_LinkList.Top        := asg_DialList.CellRect(C_DL_LINKCNT, asg_DialList.Row).Top - apn_LinkList.Height - 66 ;
            apn_LinkList.Left       := 0;
         end;


         apn_LinkList.Collaps := True;
         apn_LinkList.Visible := True;
         apn_LinkList.Collaps := False;

      end
      //---------------------------------------------------
      // 업무(S/R) 분석 : 링크된 상세내역 조회
      //---------------------------------------------------
      else if (in_Gubun = 'ANALYSIS') then
      begin

         with asg_AnalList do
         begin

            iRowCnt  := TpGetLink.RowCount;
            RowCount := 7 + iRowCnt - 1;



            for i := 0 to iRowCnt - 1 do
            begin
               Cells[C_AL_CONTEXT,    0]     := TpGetLink.GetOutputDataS('sDutyRmk',     i);
               Cells[C_AL_CONTEXT,    1]     := TpGetLink.GetOutputDataS('sDocTitle',    i);
               Cells[C_AL_CONTEXT,    2]     := TpGetLink.GetOutputDataS('sContext',     i);
               Cells[C_AL_CONTEXT,    3]     := '[' + TpGetLink.GetOutputDataS('sDutyUser', i) + '] ' + #13#10 + TpGetLink.GetOutputDataS('sHideFile',    i);
               Cells[C_AL_CONTEXT,    4]     := TpGetLink.GetOutputDataS('sDeptNm',      i);
               Cells[C_AL_CONTEXT,    5]     := TpGetLink.GetOutputDataS('sEmail',       i);
               Cells[C_AL_CONTEXT,    6 + i] := CopyByte(TpGetLink.GetOutputDataS('sAttachNm',    i), 11, LengthByte(TpGetLink.GetOutputDataS('sAttachNm',    i)));

               AutoSizeRow(1);
               AutoSizeRow(2);
               AutoSizeRow(3);
               AutoSizeRow(4);
            end;
         end;



         // Comments
         apn_AnalList.Caption.Text :=  CopyByte(TpGetLink.GetOutputDataS('sReqNo', 0), 1, 2) + ' ' +
                                       TpGetLink.GetOutputDataS('sRelDept', 0) + ' ' +
                                       TpGetLink.GetOutputDataS('sUserNm',  0) + ' ' +
                                       TpGetLink.GetOutputDataS('sRegDate', 0) + ' 요청 "' +
                                       TpGetLink.GetOutputDataS('sDocRmk', 0) + '" 상세내역';


         // Location & Display
         {
         if (asg_AnalList.Height - asg_AnalList.CellRect(C_AN_FILECNT, asg_AnalList.Row).Top) > (apn_AnalList.Height + 30) then
         begin
            apn_AnalList.Top        := asg_AnalList.CellRect(C_AN_FILECNT, asg_AnalList.Row).Top + 22;
            apn_AnalList.Left       := 0;
         end
         else
         begin
            apn_AnalList.Top        := asg_AnalList.CellRect(C_AN_FILECNT, asg_AnalList.Row).Top + 22 ;
            apn_AnalList.Left       := 0;
         end;
         }

         apn_AnalList.Top     := 106;
         apn_AnalList.Left    := 13;
         apn_AnalList.Collaps := True;
         apn_AnalList.Visible := True;
         apn_AnalList.Collaps := False;
      end;



   finally
      FreeAndNil(TpGetLink);
      Screen.Cursor := crDefault;
   end;

end;



procedure TMainDlg.asg_LinkListGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   // Horizontal Alignment
   if (ARow = 0) or ((ARow > 0) and ((ACol = C_LL_USERNM) or
                                     (ACol = C_LL_CALLNO) or
                                     (ACol = C_LL_DUTYUSER) or
                                     (ACol = C_LL_MOBILE))
                                     ) then
      HAlign := taCenter;

   // Vertical Alignment
   VAlign := vtaCenter;
end;

procedure TMainDlg.asg_DialListDblClick(Sender: TObject);
begin
   if apn_LinkList.Visible then
   begin
      apn_LinkList.Collaps := True;
      apn_LinkList.Visible := False;
   end;


   with asg_DialList do
   begin
   {
      // S/R 연락처 검색은 S/R 상세내역 조회
      if PosByte('S/R', Cells[C_DL_LINKDB,   Row]) > 0 then
      begin
         // 서비스단에서 LINKSEQ에 CQSRREQT.REQNO만 물리면, 자동 조회 가능하나,
         // 최근 1년이내 S/R 요청별로 모두 조회가 되어서, 유효한 data가 중복으로 조회되는 부분보다는
         // 정말 unique한(distinct) 연락처만 보여주는 것이 좋겠다는 판단.
         GetLinkList('ANALYSIS',
                     Trim(Cells[C_DL_LINKSEQ, Row]),
                     '',
                     '',
                     FsUserIp);
      end
      // 기타 D/B의 경우, 현재 Link 연결된 상세내역 조회
      else
      }
      begin
         GetLinkList('DIALBOOK',
                     Cells[C_DL_LOCATE,   Row],
                     Cells[C_DL_LINKDB,   Row],
                     Cells[C_DL_LINKSEQ,  Row],
                     FsUserIp);
      end;
   end;
end;

procedure TMainDlg.asg_DialScanGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   HAlign := taCenter;
   VAlign := vtaCenter;
end;

procedure TMainDlg.asg_DialScanDblClick(Sender: TObject);
var
   varResult : String;
begin
   // 선택한 최근검색 Keyword assign
   fed_Scan.Text := asg_DialScan.Cells[C_DS_KEYWORD, asg_DialScan.Row];


   if apn_LinkList.Visible then
   begin
      apn_LinkList.Collaps := True;
      apn_LinkList.Visible := False;
   end;


   // Refresh
   SelGridInfo('DIALLIST');


   // 로그 Update
   UpdateLog('DIALBOOK',
             'KEYWORD',
             FsUserIp,
             'F',
             Trim(fed_Scan.Text),
             '',
             varResult
             );

   asg_DialScan.Visible := False;

end;

procedure TMainDlg.fed_AnalysisClick(Sender: TObject);
begin
   // 리포팅 검색창 Click시에는 현 Login User 이름 Default Set. @ 2016.12.08 LSH
   if (apn_Weekly.Visible) then
      fed_Analysis.Text := FsUserNm
   else
      fed_Analysis.Text := '';

   lb_Analysis.Caption := '';
end;

procedure TMainDlg.fed_AnalysisKeyPress(Sender: TObject; var Key: Char);
var
   varResult : String;
   ResPoint  : TPoint;
   FndParam  : TFindparams;
begin
   // S/R 내역 검색(D/B)
   if (fed_Analysis.Text <> '') and
      (not fcb_ReScan.Checked)  and
      (Key = #13) then
   begin

      if (asg_Analysis.Visible) then
      begin
         asg_Analysis.Options := [goFixedVertLine,goFixedHorzLine,goVertLine,goHorzLine,goColSizing,goRowSelect];

         if apn_AnalList.Visible then
         begin
            apn_AnalList.Collaps := True;
            apn_AnalList.Visible := False;
         end;


         // Refresh
         SelGridInfo('ANALYSIS');


         // 로그 Update
         UpdateLog('ANALYSIS',
                   'KEYWORD',
                   FsUserIp,
                   'F',
                   Trim(fed_Analysis.Text),
                   '',
                   varResult
                   );

      end
      else if (apn_Weekly.Visible) then
      begin
         apn_Weekly.Collaps := True;
         apn_Weekly.Visible := True;
         asg_WeeklyRpt.ClearRows(1, asg_WeeklyRpt.RowCount);
         asg_WeeklyRpt.RowCount := 1;
         apn_Weekly.Collaps := False;

         asg_WeeklyRpt.Cells[C_WK_GUBUN, 0] := '구분';
         //fcb_ReScan.Visible := False;

         // Refresh
         if IsLogonUser then
            SelGridInfo('WEEKLYRPT');
      end
      else if (asg_WorkRpt.Visible) then
      begin
         asg_WorkRpt.ClearRows(1, asg_WorkRpt.RowCount);
         asg_WorkRpt.RowCount := 1;


         // Refresh
         if IsLogonUser then
            SelGridInfo('WORKRPT');
      end
      else if (apn_DBMaster.Visible) then
      begin
         apn_DBMaster.Collaps := True;
         apn_DBMaster.Visible := True;
         asg_DBMaster.ClearRows(1, asg_DBMaster.RowCount);
         asg_DBMaster.RowCount := 1;

         asg_DBDetail.ClearRows(1, asg_DBDetail.RowCount);
         asg_DBDetail.RowCount := 1;

         apn_DBMaster.Collaps := False;

         // Refresh
         if IsLogonUser then
            SelGridInfo('DBSEARCH');

         // 로그 Update
         UpdateLog('DBMASTER',
                   'KEYWORD',
                   FsUserIp,
                   'F',
                   Trim(fed_Analysis.Text),
                   '',
                   varResult
                   );
      end
   end
   // 결과내 재검색 (Grid)
   else if (fed_Analysis.Text <> '') and
           (fcb_ReScan.Checked)      and
           (Key = #13) then
   begin

      if (asg_Analysis.Visible) then
      begin
         asg_Analysis.Options := [goFixedVertLine,goFixedHorzLine,goVertLine,goHorzLine,goColSizing];

         FndParam := [];

         FndParam := FndParam + {[fnMatchCase] +} [fnDirectionLeftRight];     // Character InSensitive 검색위해 fnMatchCase 주석 @ 2016.11.18 LSH

         ResPoint := asg_Analysis.FindFirst(Trim(fed_Analysis.Text), FndParam);

         if ResPoint.X >= 0 then
         begin
            asg_Analysis.Col   := ResPoint.X;
            asg_Analysis.Row   := ResPoint.Y;
         end
         else
            MessageBox(self.Handle,
                       PChar('찾으시는 항목이 화면 List에 존재하지 않습니다.'),
                       PChar('S/R 검색 : 결과내 재검색 결과 알림'),
                       MB_OK + MB_ICONWARNING)
                       ;

         asg_Analysis.SetFocus;
      end
      else if (apn_Weekly.Visible) then
      begin
         //asg_WorkRpt.Options := [goFixedVertLine,goFixedHorzLine,goVertLine,goHorzLine,goColSizing];

         FndParam := [];

         FndParam := FndParam + {[fnMatchCase] +} [fnDirectionLeftRight];     // Character InSensitive 검색위해 fnMatchCase 주석 @ 2016.11.18 LSH

         ResPoint := asg_WeeklyRpt.FindFirst(Trim(fed_Analysis.Text), FndParam);

         if ResPoint.X >= 0 then
         begin
            asg_WeeklyRpt.Col   := ResPoint.X;
            asg_WeeklyRpt.Row   := ResPoint.Y;
         end
         else
            MessageBox(self.Handle,
                       PChar('찾으시는 항목이 화면 List에 존재하지 않습니다.'),
                       PChar('S/R 검색 : 결과내 재검색 결과 알림'),
                       MB_OK + MB_ICONWARNING)
                       ;

         asg_WeeklyRpt.SetFocus;
      end
      else if (asg_WorkRpt.Visible) then
      begin
         FndParam := [];

         FndParam := FndParam + {[fnMatchCase] +} [fnDirectionLeftRight];     // Character InSensitive 검색위해 fnMatchCase 주석 @ 2016.11.18 LSH

         ResPoint := asg_WorkRpt.FindFirst(Trim(fed_Analysis.Text), FndParam);

         if ResPoint.X >= 0 then
         begin
            asg_WorkRpt.Col   := ResPoint.X;
            asg_WorkRpt.Row   := ResPoint.Y;
         end
         else
            MessageBox(self.Handle,
                       PChar('찾으시는 항목이 화면 List에 존재하지 않습니다.'),
                       PChar('S/R 검색 : 결과내 재검색 결과 알림'),
                       MB_OK + MB_ICONWARNING)
                       ;

         asg_WorkRpt.SetFocus;
      end
      else if (apn_DBMaster.Visible) then
      begin
         asg_DBMaster.Options := [goFixedVertLine,goFixedHorzLine,goVertLine,goHorzLine,goColSizing];

         FndParam := [];

         FndParam := FndParam + {[fnMatchCase] +} [fnDirectionLeftRight];     // Character InSensitive 검색위해 fnMatchCase 주석 @ 2016.11.18 LSH

         ResPoint := asg_DBMaster.FindFirst(Trim(fed_Analysis.Text), FndParam);

         if ResPoint.X >= 0 then
         begin
            asg_DBMaster.Col   := ResPoint.X;
            asg_DBMaster.Row   := ResPoint.Y;

            asg_DBMaster.SetFocus;
         end
         else
         begin
            asg_DBDetail.Options    := [goFixedVertLine,goFixedHorzLine,goVertLine,goHorzLine,goColSizing,goRangeSelect];

            FndParam := [];

            FndParam := FndParam + {[fnMatchCase] +} [fnDirectionLeftRight];  // Character InSensitive 검색위해 fnMatchCase 주석 @ 2016.11.18 LSH

            ResPoint := asg_DBDetail.FindFirst(Trim(fed_Analysis.Text), FndParam);

            if ResPoint.X >= 0 then
            begin
               asg_DBDetail.Col   := ResPoint.X;
               asg_DBDetail.Row   := ResPoint.Y;

               asg_DBDetail.SetFocus;
            end
            else
               MessageBox(self.Handle,
                          PChar('찾으시는 항목이 화면 List에 존재하지 않습니다.'),
                          PChar('D/B 검색 : 결과내 재검색 결과 알림'),
                          MB_OK + MB_ICONWARNING)
                          ;
         end;
      end;
   end;
end;


procedure TMainDlg.asg_AnalysisGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   // Horizontal Alignment
   if (ARow = 0) or ((ARow > 0) and ((ACol = C_AN_REQFLAG)  or
                                     (ACol = C_AN_REQDEPT)  or
                                     (ACol = C_AN_REQUSER)  or
                                     (ACol = C_AN_REQDATE)) or
                                     (ACol = C_AN_DUTYUSER) or
                                     (ACol = C_AN_PROCESS)  or
                                     (ACol = C_AN_FILECNT)
                                     ) then
      HAlign := taCenter;

   // Vertical Alignment
   VAlign := vtaCenter;
end;

procedure TMainDlg.asg_AnalysisDblClick(Sender: TObject);
begin
   if apn_AnalList.Visible then
   begin
      apn_AnalList.Collaps := True;
      apn_AnalList.Visible := False;
   end;

   with asg_Analysis do
   begin
      GetLinkList('ANALYSIS',
                  Cells[C_AN_REQNO,   Row],
                  Cells[C_AN_REQDATE, Row],
                  Cells[C_AN_REQDEPT, Row],
                  FsUserIp);
   end;

   // S/R 상세내역 Grid 포커싱 @ 2016.11.18 LSH
   asg_AnalList.SetFocus;
end;

procedure TMainDlg.asg_AnalysisClickCell(Sender: TObject; ARow,
  ACol: Integer);
begin
   lb_Analysis.Caption := '';

   if apn_AnalList.Visible then
   begin
      apn_AnalList.Collaps := True;
      apn_AnalList.Visible := False;
   end;
end;

procedure TMainDlg.asg_AnalListGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   // Horizontal Alignment
   if ACol = C_AL_TITLE then
      HAlign := taCenter;

   // Vertical Alignment
   VAlign := vtaCenter;
end;

procedure TMainDlg.asg_AnalListDblClick(Sender: TObject);
var
   ServerFileName, ClientFileName: String;
   FTP_SVRIP, FTP_USERID, FTP_PASSWD, FTP_HOSTNAME, FTP_DIR : String;
begin
   with asg_AnalList do
   begin
      // 첨부 File 이 아닌 부분만 Event 처리
      if (Row < 6) then
      begin
         if apn_AnalList.Visible then
         begin
            apn_AnalList.Collaps := True;
            apn_AnalList.Visible := False;
         end
      end
      else if (Row >= 6) then
      begin
         // S/R 첨부 파일명
         ServerFileName := Trim(Cells[Col, Row]);


         // Local 파일이름 Set
         // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
         ClientFileName := G_HOMEDIR + 'TEMP\SPOOL\' + ServerFileName;



         // Local에 해당 파일 존재유무 체크
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
            FTP_USERID := '';
            FTP_PASSWD := '';
            FTP_DIR    := '/kumis/app/mis/media/cq/';



            // S/R 첨부파일 다운로드
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



            // 해당 파일 Variant로 받음.
            AppendVariant(DownFile, ClientFileName);



            // 다운로드 횟수 체크
            DownCnt := DownCnt + 1;

         end;

         try
            if (PosByte('.exe', ClientFileName) > 0) or
               (PosByte('.zip', ClientFileName) > 0) or
               (PosByte('.rar', ClientFileName) > 0) then
            begin
               MessageBox(self.Handle,
                          PChar('첨부파일(' + ClientFileName + ') 다운로드가 완료되었습니다.' + #13#10 + #13#10 +
                                '※ 임시 다운로드 폴더 --> C:\KUMC(_DEV)\TEMP\SPOOL\'),
                          'S/R 첨부파일 다운로드 실행 완료',
                          MB_OK + MB_ICONINFORMATION);
            end
            else
            begin
               ShellExecute(HANDLE, 'open',
                            PCHAR(Trim(Cells[Col, Row])),
                            PCHAR(''),
                            // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
                            PCHAR(G_HOMEDIR + 'TEMP\SPOOL\'),
                            SW_SHOWNORMAL);
            end;

         except
            MessageBox(self.Handle,
                       PChar('해당 첨부파일 실행중 오류가 발생하였습니다.' + #13#10 + #13#10 +
                             '다시 실행해 주시기 바랍니다.'),
                       'S/R 첨부파일 다운로드 실행 오류',
                       MB_OK + MB_ICONERROR);

            Exit;
         end;
      end;
   end;

   // S/R 업무분석 조회 Panel - on시, 해당 Grid 포커싱 적용 @ 2016.11.18 LSH
   if (ftc_Dialog.ActiveTab = AT_ANALYSIS) then
   begin
      if apn_Weekly.Visible then
         asg_WeeklyRpt.SetFocus
      else
         asg_Analysis.SetFocus;
   end;

end;


procedure TMainDlg.asg_AnalListKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   case (Key) of

      //[Ctrl + C] : 선택한 Cell 클립보드로 복사
      Ord('C') : if (ssCtrl in Shift) then
                 begin
                    asg_AnalList.CopySelectionToClipboard;
                 end;

      //[Ctrl + A] : 해당 Grid Cell 전체내역 클립보드로 복사
      Ord('A') : if (ssCtrl in Shift) then
                 begin
                    asg_AnalList.CopyToClipBoard;
                 end;

      //[ESC] : 해당 S/R 상세내역 Panel-Off @ 2016.11.18 LSH
      Ord(VK_ESCAPE) :
                     begin
                        asg_AnalListDblClick(Sender);
                     end;
      //[Enter] : 첨부 File 실행
      Ord(VK_RETURN) :
                     begin
                        // 첨부 File Row-Index [Enter] 입력시 실행 @ 2016.11.18 LSH
                        if asg_AnalList.Row >= 6 then
                           asg_AnalListDblClick(Sender);
                     end;

   end;
end;

procedure TMainDlg.asg_AnalListMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
var
   NowCol, NowRow : Integer;
begin
   with asg_AnalList do
   begin
      // 현재 Cell 좌표 가져오기
      MouseToCell(X, Y, NowCol, NowRow);

      if (NowRow >= 1) and
         (NowCol in [1]) then
      begin
         ShowHint := False;
      end
      else
         ShowHint := True;
   end;

end;

procedure TMainDlg.fsbt_WorkRptClick(Sender: TObject);
begin
   // 업무일지 작성(TMemo 동적 컴포넌트) Panel Close @ 2014.07.31 LSH
   SetPanelStatus('WORKRPT', 'OFF');

   if apn_Weekly.Visible then
   begin
      apn_Weekly.Collaps := True;
      apn_Weekly.Visible := False;
   end;

   if apn_AnalList.Visible then
   begin
      apn_AnalList.Collaps := True;
      apn_AnalList.Visible := False;
   end;

   if apn_DBMaster.Visible then
   begin
      apn_DBMaster.Collaps := True;
      apn_DBMaster.Visible := False;
   end;

   fcb_WorkGubun.Top     := 999;
   fcb_WorkGubun.Left    := 999;
   fcb_WorkGubun.Visible := False;

   fcb_DutyPart.Top      := 999;
   fcb_DutyPart.Left     := 999;
   fcb_DutyPart.Visible  := False;

   fcb_WorkArea.Top     := 30;
   fcb_WorkArea.Left    := 205;
   fcb_WorkArea.Visible := True;
   fcb_WorkArea.Color   := $0087D5DB;

   fcb_ReScan.Top          := 12;
   fcb_ReScan.Left         := 391;
   fcb_ReScan.ColorFocused := $0087D5DB;
   fcb_ReScan.Checked      := False;

   asg_Analysis.Top     := 999;
   asg_Analysis.Left    := 999;
   asg_Analysis.Visible := False;

   fed_Analysis.Top        := 30;
   fed_Analysis.Left       := 263;
   fed_Analysis.Width      := 221;
   fed_Analysis.Visible    := True;
   fed_Analysis.ColorFlat  := $0087D5DB;

   fed_Analysis.Text       := '';         // 업무일지 검색기능 적용으로 null 초기화 @ 2016.09.02 LSH (fed_Analysis.Text       := FsUserNm;)
   fed_Analysis.Hint       := '업무일지에서 검색하려는 키워드(예: 서버, 개발, 자료집계..)를 적고 Enter를 눌러보세요.';

   asg_WorkRpt.Top      := 62;
   asg_WorkRpt.Left     := 6;
   asg_WorkRpt.Visible  := True;

   fsbt_Analysis.ColorFocused := $0087D5DB;
   fsbt_WorkRpt.ColorFocused  := $0087D5DB;
   fsbt_Weekly.ColorFocused   := $0087D5DB;
   fsbt_DBMaster.ColorFocused := $0087D5DB;

   fed_DateTitle.ColorFlat := $0087D5DB;
   fmed_AnalFrom.ColorFlat := $0087D5DB;
   fmed_AnalTo.ColorFlat   := $0087D5DB;

   fed_DateTitle.Text      := '기간';
   label8.Left             := 124;
   fmed_AnalFrom.Left      := 57;
   fmed_AnalTo.Left        := 135;

   lb_Analysis.Caption := '▶ 업무일지(KUMC) 근무처별/업무구분별/기간별 조회 제공.';

   fmed_AnalFrom.Text   := FormatDateTime('yyyy-mm-dd', DATE - 1);
   fmed_AnalTo.Text     := FormatDateTime('yyyy-mm-dd', DATE);

   fcb_WorkGubun.Top     := 999;
   fcb_WorkGubun.Left    := 999;
   fcb_WorkGubun.Visible := False;

   // 접속자 IP 식별후, 근무처 Assign
   if PosByte('안암도메인', FsUserIp) > 0 then
      fcb_WorkArea.Text := '안암'
   else if PosByte('구로도메인', FsUserIp) > 0 then
      fcb_WorkArea.Text := '구로'
   else if PosByte('안산도메인', FsUserIp) > 0 then
      fcb_WorkArea.Text := '안산';

   fsbt_Print.Visible := True;
   fsbt_Print.ColorFocused := $0087D5DB;


   // 업무일지 내역 조회
   if IsLogonUser then
      SelGridInfo('WORKRPT');

end;

procedure TMainDlg.fsbt_AnalysisClick(Sender: TObject);
var
   vKey : Char;
begin
   // 업무일지 작성(TMemo 동적 컴포넌트) Panel Close @ 2014.07.31 LSH
   SetPanelStatus('WORKRPT', 'OFF');

   if apn_Weekly.Visible then
   begin
      apn_Weekly.Collaps := True;
      apn_Weekly.Visible := False;
   end;

   if apn_AnalList.Visible then
   begin
      apn_AnalList.Collaps := True;
      apn_AnalList.Visible := False;
   end;

   if apn_DBMaster.Visible then
   begin
      apn_DBMaster.Collaps := True;
      apn_DBMaster.Visible := False;
   end;

   fcb_WorkGubun.Top     := 999;
   fcb_WorkGubun.Left    := 999;
   fcb_WorkGubun.Visible := False;

   fcb_WorkArea.Top     := 999;
   fcb_WorkArea.Left    := 999;
   fcb_WorkArea.Visible := False;

   fcb_DutyPart.Top      := 999;
   fcb_DutyPart.Left     := 999;
   fcb_DutyPart.Visible  := False;

   fcb_ReScan.Top          := 12;
   fcb_ReScan.Left         := 334;
   fcb_ReScan.ColorFocused := $00FBDBE9;
   fcb_ReScan.Checked      := False;
   asg_Analysis.Options    := [goFixedVertLine,goFixedHorzLine,goVertLine,goHorzLine,goColSizing,goRowSelect];

   asg_Analysis.Top     := 62;
   asg_Analysis.Left    := 6;
   asg_Analysis.Visible := True;

   asg_Analysis.ClearRows(1, asg_Analysis.RowCount);
   asg_Analysis.RowCount := 2;


   fed_Analysis.Top        := 30;
   fed_Analysis.Left       := 205;
   fed_Analysis.Width      := 221;
   fed_Analysis.ColorFlat  := $00FBDBE9;
   fed_Analysis.Text       := '';
   fed_Analysis.Visible    := True;
   fed_Analysis.Hint       := 'S/R 제목, 요청시 포함된 문구(코드, 처방명 등), 요청자, 환자번호, 첨부파일명 등 아무거나 검색해보세요.';


   asg_WorkRpt.Top      := 999;
   asg_WorkRpt.Left     := 999;
   asg_WorkRpt.Visible  := False;

   fsbt_Analysis.ColorFocused := $00FBD1DD;
   fsbt_WorkRpt.ColorFocused  := $00FBD1DD;
   fsbt_Weekly.ColorFocused   := $00FBD1DD;
   fsbt_DBMaster.ColorFocused := $00FBD1DD;

   fed_DateTitle.ColorFlat := $00FBD1DD;
   fmed_AnalFrom.ColorFlat := $00FBD1DD;
   fmed_AnalTo.ColorFlat   := $00FBD1DD;

   fed_DateTitle.Text      := '기간';
   label8.Left             := 124;
   fmed_AnalFrom.Left      := 57;
   fmed_AnalTo.Left        := 135;

   lb_Analysis.Caption := '▶ S/R 요청내역 분석을 위한 효율적인 검색을 제공합니다.';

   fmed_AnalFrom.Text   := FormatDateTime('yyyy-mm-dd', DATE - 730);   // 약 3년치 검색 --> 2년치 검색 전환 @ 2016.11.18 LSH
   fmed_AnalTo.Text     := FormatDateTime('yyyy-mm-dd', DATE);

   fcb_WorkGubun.Top     := 999;
   fcb_WorkGubun.Left    := 999;
   fcb_WorkGubun.Visible := False;

   fsbt_Print.Visible   := False;

   // S/R 접수내역 자동조회
   vKey := #13;
   fed_Analysis.Text := '요청서접수';
   fed_AnalysisKeyPress(Sender, vKey);


   // S/R 이력 조회 Grid 포커싱 @ 2016.11.18 LSH
   asg_Analysis.SetFocus;

end;

procedure TMainDlg.fmed_AnalFromChange(Sender: TObject);
begin

   if (IsLogonUser) then
   begin
      if (asg_WorkRpt.Visible) and
         (
            (
               (CopyByte(Trim(fmed_AnalFrom.Text), 1, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalFrom.Text), 2, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalFrom.Text), 3, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalFrom.Text), 4, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalFrom.Text), 6, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalFrom.Text), 7, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalFrom.Text), 9, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalFrom.Text), 10, 1) <> '')
            )
          )
          then
         SelGridInfo('WORKRPT')
      {  -- S/R 기간변경시 실시간 조회 제외 @ 2016.11.18 LSH
      else if (asg_Analysis.Visible) and
              (Trim(fed_Analysis.Text) <> '') and
         (
            (
               (CopyByte(Trim(fmed_AnalFrom.Text), 1, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalFrom.Text), 2, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalFrom.Text), 3, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalFrom.Text), 4, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalFrom.Text), 6, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalFrom.Text), 7, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalFrom.Text), 9, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalFrom.Text), 10, 1) <> '')
            )
          )
          then
         SelGridInfo('ANALYSIS')
      }
      else if (apn_Weekly.Visible) and
         (
            (
               (CopyByte(Trim(fmed_AnalFrom.Text), 1, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalFrom.Text), 2, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalFrom.Text), 3, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalFrom.Text), 4, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalFrom.Text), 6, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalFrom.Text), 7, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalFrom.Text), 9, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalFrom.Text), 10, 1) <> '')
            )
          )
          then
         SelGridInfo('WEEKLYRPT')
      {  -- 릴리즈 대장 기간변경시 시작일은 실시간 조회 제외 @ 2016.11.18 LSH
      else if (apn_Release.Visible) and
         (
            (
               (CopyByte(Trim(fmed_RegFrDt.Text), 1, 1)  <> '') and
               (CopyByte(Trim(fmed_RegFrDt.Text), 2, 1)  <> '') and
               (CopyByte(Trim(fmed_RegFrDt.Text), 3, 1)  <> '') and
               (CopyByte(Trim(fmed_RegFrDt.Text), 4, 1)  <> '') and
               (CopyByte(Trim(fmed_RegFrDt.Text), 6, 1)  <> '') and
               (CopyByte(Trim(fmed_RegFrDt.Text), 7, 1)  <> '') and
               (CopyByte(Trim(fmed_RegFrDt.Text), 9, 1)  <> '') and
               (CopyByte(Trim(fmed_RegFrDt.Text), 10, 1) <> '')
            )
          )
          then
         SelGridInfo('RELEASESCAN')
      }
         ;
   end;

end;

procedure TMainDlg.fmed_AnalToChange(Sender: TObject);
begin
   if (IsLogonUser) then
   begin
      if (asg_WorkRpt.Visible) and
         (
            (
               (CopyByte(Trim(fmed_AnalTo.Text), 1, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalTo.Text), 2, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalTo.Text), 3, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalTo.Text), 4, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalTo.Text), 6, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalTo.Text), 7, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalTo.Text), 9, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalTo.Text), 10, 1) <> '')
            )
          )
          then
         SelGridInfo('WORKRPT')
      {  -- S/R 기간변경시 실시간 조회 제외 @ 2016.11.18 LSH
      else if (asg_Analysis.Visible) and
              (Trim(fed_Analysis.Text) <> '') and
         (
            (
               (CopyByte(Trim(fmed_AnalTo.Text), 1, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalTo.Text), 2, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalTo.Text), 3, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalTo.Text), 4, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalTo.Text), 6, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalTo.Text), 7, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalTo.Text), 9, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalTo.Text), 10, 1) <> '')
            )
          )
          then
         SelGridInfo('ANALYSIS')
      }
      else if (apn_Weekly.Visible) and
         (
            (
               (CopyByte(Trim(fmed_AnalTo.Text), 1, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalTo.Text), 2, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalTo.Text), 3, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalTo.Text), 4, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalTo.Text), 6, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalTo.Text), 7, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalTo.Text), 9, 1)  <> '') and
               (CopyByte(Trim(fmed_AnalTo.Text), 10, 1) <> '')
            )
          )
          then
         SelGridInfo('WEEKLYRPT')
      {
      else if (apn_Release.Visible) and
         (
            (
               (CopyByte(Trim(fmed_RegToDt.Text), 1, 1)  <> '') and
               (CopyByte(Trim(fmed_RegToDt.Text), 2, 1)  <> '') and
               (CopyByte(Trim(fmed_RegToDt.Text), 3, 1)  <> '') and
               (CopyByte(Trim(fmed_RegToDt.Text), 4, 1)  <> '') and
               (CopyByte(Trim(fmed_RegToDt.Text), 6, 1)  <> '') and
               (CopyByte(Trim(fmed_RegToDt.Text), 7, 1)  <> '') and
               (CopyByte(Trim(fmed_RegToDt.Text), 9, 1)  <> '') and
               (CopyByte(Trim(fmed_RegToDt.Text), 10, 1) <> '')
            )
          )
          then
         SelGridInfo('RELEASESCAN')
     }
         ;

   end;

end;

procedure TMainDlg.asg_WorkRptGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   // Horizontal Alignment
   if (ARow = 0) or ((ARow > 0 ) and (ACol in [C_WR_LOCATE,
                                               C_WR_DUTYUSER,
                                               C_WR_DUTYPART,
                                               C_WR_REGDATE])
                      ) then
      HAlign := taCenter;

   // Vertical Alignment
   VAlign := vtaCenter;
end;

procedure TMainDlg.asg_WorkRptKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  Res : TPoint;
begin
   case (Key) of

      //[Ctrl + C] : 선택한 Cell 클립보드로 복사
      Ord('C') : if (ssCtrl in Shift) then
                 begin
                    asg_WorkRpt.CopySelectionToClipboard;
                 end;

      //[Ctrl + A] : 해당 Grid Cell 전체내역 클립보드로 복사
      Ord('A') : if (ssCtrl in Shift) then
                 begin
                    asg_WorkRpt.CopyToClipBoard;
                 end;

      //[F3 Key] : 다음 찾기
      Ord(VK_F3) :
                  begin
                     Res := asg_WorkRpt.FindNext;

                     if (Res.X >= 0) and
                        (Res.Y >= 0) then
                     begin
                        asg_WorkRpt.Col := Res.X;
                        asg_WorkRpt.Row := Res.Y;
                     end
                     else
                        MessageBox(self.Handle,
                                   PChar('찾으시는 항목이 화면 List에 존재하지 않습니다.'),
                                   PChar('업무일지 : 결과내 재검색 결과 알림'),
                                   MB_OK + MB_ICONWARNING)
                                   ;
                  end;

      //[Ctrl + F] : 검색 EditBox 포커싱 @ 2016.11.18 LSH
      Ord('F') :
                  if (ssCtrl in Shift) then
                  begin
                     fed_Analysis.SetFocus;
                  end;

   end;

end;

procedure TMainDlg.fcb_WorkAreaChange(Sender: TObject);
begin
   if (asg_WorkRpt.Visible) and
      (IsLogonUser) then
      SelGridInfo('WORKRPT')
   else if (apn_Weekly.Visible) and
           (IsLogonUser) then
   begin
      fed_Analysis.Text := '';
      fed_Analysis.SetFocus;
   end;

end;

procedure TMainDlg.fcb_WorkGubunChange(Sender: TObject);
begin
   if IsLogonUser then
      if (apn_DBMaster.Visible) then
      begin
         // D/B 검색 카테고리 변경전, [결과내 재검색] 체크 자동해제 @ 2016.11.18 LSH
         if fcb_ReScan.Checked then
            fcb_ReScan.Checked := False;


         SelGridInfo('DBCATEGORY');

         // D/B 검색 카테고리 변경후, [결과내 재검색] 체크 자동설정 @ 2016.11.18 LSH
         fcb_ReScan.Checked := True;

         // 검색 EditBox 포커싱 @ 2016.11.18 LSH
         fed_Analysis.SetFocus;

      end;
end;

procedure TMainDlg.fsbt_WeeklyClick(Sender: TObject);
begin
   // 업무일지 작성(TMemo 동적 컴포넌트) Panel Close @ 2014.07.31 LSH
   SetPanelStatus('WORKRPT', 'OFF');


   //------------------------------------------------------------------
   // 다이얼Book / S/R검색 프로필 : S/R 최근 1년간 History 조회
   //------------------------------------------------------------------
   if (ftc_Dialog.ActiveTab = AT_DIALBOOK) then
   begin
      //
      fed_Analysis.Text    := asg_UserProfile.Cells[1, R_PR_USERID];
      apn_Weekly.Top       := 54;
      apn_Weekly.Left      := 13;

      apn_UserProfile.Collaps := True;
      apn_UserProfile.Visible := False;

      lb_DialScan.Caption   := '▶ ' + Trim(asg_UserProfile.Cells[1, R_PR_USERNM]) + '님의 최근 1년간 S/R 요청내역을 조회하였습니다. [더블클릭시, 다이얼Book 복귀]';

      apn_Weekly.Collaps := True;
      apn_Weekly.Visible := True;
      asg_WeeklyRpt.ClearRows(1, asg_WeeklyRpt.RowCount);
      asg_WeeklyRpt.RowCount := 1;
      apn_Weekly.Collaps := False;

      asg_WeeklyRpt.Cells[C_WK_GUBUN, 0] := 'S/R담당';


      // 업무리포트 조회
      if (Trim(fed_Analysis.Text) <> '') and
         (IsLogonUser) then
         SelGridInfo('SRHISTORY');

   end
   else if (ftc_Dialog.ActiveTab = AT_ANALYSIS) and
           (apn_UserProfile.Visible) then
   begin
      //
      fed_Analysis.Text    := asg_UserProfile.Cells[1, R_PR_USERID];
      apn_Weekly.Top       := 53;
      apn_Weekly.Left      := 13;

      apn_UserProfile.Collaps := True;
      apn_UserProfile.Visible := False;

      lb_Analysis.Caption   := '▶ ' + Trim(asg_UserProfile.Cells[1, R_PR_USERNM]) + '님의 최근 1년간 S/R 요청내역을 조회하였습니다. [더블클릭시, S/R검색 복귀]';

      apn_Weekly.Collaps := True;
      apn_Weekly.Visible := True;
      asg_WeeklyRpt.ClearRows(1, asg_WeeklyRpt.RowCount);
      asg_WeeklyRpt.RowCount := 1;
      apn_Weekly.Collaps := False;

      asg_WeeklyRpt.Cells[C_WK_GUBUN, 0] := 'S/R담당';
      //fcb_ReScan.Visible := False;


      // 업무리포트 조회
      if (Trim(fed_Analysis.Text) <> '') and
         (IsLogonUser) then
         SelGridInfo('SRHISTORY');

   end
   else
   begin
      fcb_WorkGubun.Top     := 999;
      fcb_WorkGubun.Left    := 999;
      fcb_WorkGubun.Visible := False;

      fcb_DutyPart.Top      := 999;
      fcb_DutyPart.Left     := 999;
      fcb_DutyPart.Visible  := False;

      fcb_ReScan.Top          := 12;
      fcb_ReScan.Left         := 391;
      fcb_ReScan.ColorFocused := $008B8FE3;
      fcb_ReScan.Checked      := False;

      fed_Analysis.Top        := 30;
      fed_Analysis.Left       := 263;
      fed_Analysis.Width      := 221;
      fed_Analysis.Visible    := True;
      fed_Analysis.ColorFlat  := $008B8FE3;
      fed_Analysis.Hint       := '리포팅을 작성하려는 담당자 이름(예: 이세하) + 검색 키워드(예: MDO110F1)를 입력한 후(예: 이세하 MDO110F1) Enter를 눌러보세요.';

      fsbt_Analysis.ColorFocused := $008B8FE3;
      fsbt_WorkRpt.ColorFocused  := $008B8FE3;
      fsbt_Weekly.ColorFocused   := $008B8FE3;
      fsbt_DBMaster.ColorFocused := $008B8FE3;

      apn_UserProfile.Collaps := True;
      apn_UserProfile.Visible := False;

      fcb_WorkArea.Top     := 30;
      fcb_WorkArea.Left    := 205;
      fcb_WorkArea.Visible := True;
      fcb_WorkArea.BringToFront;
      fcb_WorkArea.Color   := $008B8FE3;

      fed_DateTitle.ColorFlat := $008B8FE3;
      fmed_AnalFrom.ColorFlat := $008B8FE3;
      fmed_AnalTo.ColorFlat   := $008B8FE3;


      fed_DateTitle.Text      := '기간';
      label8.Left             := 124;
      fmed_AnalFrom.Left      := 57;
      fmed_AnalTo.Left        := 135;

      lb_Analysis.Caption := '▶ S/R 내역과 업무일지(KUMC) 분석을 통한 리포트를 작성합니다.';

      fmed_AnalFrom.Text   := FormatDateTime('yyyy-mm-dd', DATE - 7);
      fmed_AnalTo.Text     := FormatDateTime('yyyy-mm-dd', DATE);

      if apn_AnalList.Visible then
      begin
         apn_AnalList.Collaps := True;
         apn_AnalList.Visible := False;
      end;

      if apn_DBMaster.Visible then
      begin
         apn_DBMaster.Collaps := True;
         apn_DBMaster.Visible := False;
      end;


      // 접속자 IP 식별후, 근무처 Assign
      if (ftc_Dialog.ActiveTab = AT_DIALOG) then
      begin
         if PosByte('안암', asg_UserProfile.Cells[1, R_PR_DUTYPART]) > 0 then
         begin
            fcb_WorkArea.Text := '안암';
         end
         else if PosByte('구로', asg_UserProfile.Cells[1, R_PR_DUTYPART]) > 0 then
         begin
            fcb_WorkArea.Text := '구로';
         end
         else if PosByte('안산', asg_UserProfile.Cells[1, R_PR_DUTYPART]) > 0 then
         begin
            fcb_WorkArea.Text := '안산';
         end;

         fed_Analysis.Text    := Trim(asg_UserProfile.Cells[1, R_PR_USERNM]);

         apn_Weekly.Top       := 55;
         apn_Weekly.Left      := 15;

         lb_Network.Caption   := '▶ ' + Trim(asg_UserProfile.Cells[1, R_PR_USERNM]) + '님의 최근 업무스케쥴을 조회하였습니다. [더블클릭시, 비상연락망으로 복귀]'

      end
      else if (ftc_Dialog.ActiveTab = AT_ANALYSIS) then
      begin
         if PosByte('안암도메인', FsUserIp) > 0 then
         begin
            fcb_WorkArea.Text := '안암';
         end
         else if PosByte('구로도메인', FsUserIp) > 0 then
         begin
            fcb_WorkArea.Text := '구로';
         end
         else if PosByte('안산도메인', FsUserIp) > 0 then
         begin
            fcb_WorkArea.Text := '안산';
         end;

         fed_Analysis.Text    := FsUserNm;

         apn_Weekly.Top        := 83;
         apn_Weekly.Left       := 15;

      end;


      asg_Analysis.Top      := 999;
      asg_Analysis.Left     := 999;
      asg_Analysis.Visible  := False;

      asg_WorkRpt.Top       := 999;
      asg_WorkRpt.Left      := 999;
      asg_WorkRpt.Visible   := False;

      apn_Weekly.Collaps := True;
      apn_Weekly.Visible := True;
      asg_WeeklyRpt.ClearRows(1, asg_WeeklyRpt.RowCount);
      asg_WeeklyRpt.RowCount := 1;
      apn_Weekly.Collaps := False;

      asg_WeeklyRpt.Cells[C_WK_GUBUN, 0] := '구분';

      fsbt_Print.Visible := True;
      fsbt_Print.ColorFocused := $008B8FE3;

      // 업무리포트 조회
      if (Trim(fed_Analysis.Text) <> '') and
         (IsLogonUser) then
         SelGridInfo('WEEKLYRPT');
   end;
end;

procedure TMainDlg.asg_WeeklyRptKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  Res : TPoint;
begin
   case (Key) of

      //[Ctrl + C] : 선택한 Cell 클립보드로 복사
      Ord('C') : if (ssCtrl in Shift) then
                 begin
                    asg_WeeklyRpt.CopySelectionToClipboard;
                 end;

      //[Ctrl + A] : 해당 Grid Cell 전체내역 클립보드로 복사
      Ord('A') : if (ssCtrl in Shift) then
                 begin
                    asg_WeeklyRpt.CopyToClipBoard;
                 end;

      //[Ctrl + F] : 검색 EditBox 포커싱 @ 2016.12.08 LSH
      Ord('F') :
                  if (ssCtrl in Shift) then
                  begin
                     fed_Analysis.SetFocus;
                  end;

      //[F3 Key] : 다음 찾기
      Ord(VK_F3) :
                  begin
                     Res := asg_WeeklyRpt.FindNext;

                     if (Res.X >= 0) and
                        (Res.Y >= 0) then
                     begin
                        asg_WeeklyRpt.Col := Res.X;
                        asg_WeeklyRpt.Row := Res.Y;
                     end
                     else
                        MessageBox(self.Handle,
                                   PChar('찾으시는 항목이 화면 List에 존재하지 않습니다.'),
                                   PChar('리포트작성 : 결과내 재검색 결과 알림'),
                                   MB_OK + MB_ICONWARNING)
                                   ;
                  end;

   end;
end;

procedure TMainDlg.asg_WeeklyRptGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin

   // Horizontal Alignment
   if (ARow = 0) or ((ARow > 0) and ((ACol = C_WK_DATE)  or
                                     (ACol = C_WK_GUBUN) or
                                     (ACol = C_WK_STEP))) then
      HAlign := taCenter;

   // Vertical Alignment
   VAlign := vtaCenter;
end;

procedure TMainDlg.asg_SelDocGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin

   // Horizontal Alignment
   if (ARow = 0) or ((ARow > 0) and ((ACol = C_SD_DOCLIST)  or
                                     (ACol = C_SD_DOCSEQ)   or
                                     (ACol = C_SD_REGDATE)  or
                                     (ACol = C_SD_REGUSER)  or
                                     (ACol = C_SD_RELDEPT)  or
                                     (ACol = C_SD_DOCLOC))
                                     ) then
      HAlign := taCenter;

   // Vertical Alignment
   VAlign := vtaCenter;

end;


procedure TMainDlg.asg_SelDocKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   case (Key) of

      //[Ctrl + C] : 선택한 Cell 클립보드로 복사
      Ord('C') : if (ssCtrl in Shift) then
                 begin
                    asg_SelDoc.CopySelectionToClipboard;
                 end;

      //[Ctrl + A] : 해당 Grid Cell 전체내역 클립보드로 복사
      Ord('A') : if (ssCtrl in Shift) then
                 begin
                    asg_SelDoc.CopyToClipBoard;
                 end;

   end;
end;



procedure TMainDlg.asg_WorkRptPrintPage(Sender: TObject; Canvas: TCanvas;
  PageNr, PageXSize, PageYSize: Integer);
var
   savefont : tfont;
   ts{,tw}    : integer;
   RecName  : String;
const
   myowntitle : string = ' ';
begin
   if asg_WorkRpt.PrintColStart <> 0 then exit;


   // Report Name 정의
   RecName := '정보전산팀 업무일지';



   // Title에 작성일자 + 요일 표기
   if (fmed_AnalTo.Text > fmed_AnalFrom.Text) then
      RecName := RecName + ' [' + fmed_AnalFrom.Text + ', '  + GetDayofWeek(StrToDate(fmed_AnalFrom.Text), 'HS') + ' ~ ' + fmed_AnalTo.Text + ', ' + GetDayofWeek(StrToDate(fmed_AnalTo.Text), 'HS') +']'
   else
      RecName := RecName + ' [' + fmed_AnalFrom.Text + ', '  + GetDayofWeek(StrToDate(fmed_AnalFrom.Text), 'HS') + ']';



   // 의료원명 Display
   with canvas do
   begin
      savefont := TFont.Create;
      savefont.assign(font);

      font.name   := 'Arial';
      font.Size   := 10;
      font.Style  := [];
      font.height := asg_WorkRpt.mapfontheight(10);
      font.Color  := clBlack;

      ts := asg_WorkRpt.Printcoloffset[0];
//      tw := asg_WorkRpt.Printpagewidth;
      //ts := ts+ ((tw - textwidth('고려대학교의료원')) shr 1);
      ts := ts +  10;

      Textout(ts, -15, '고려대학교의료원');

      font.assign(savefont);

      savefont.free;
   end;


   // 기록지명 Display
   with canvas do
   begin
      savefont := TFont.Create;
      savefont.assign(font);

      font.name   := 'Arial';
      font.Size   := 18;
      font.Style  := [fsBold];
      font.height := asg_WorkRpt.mapfontheight(18);
      font.Color  := clBlack;

      ts := asg_WorkRpt.Printcoloffset[0];
//      tw := asg_WorkRpt.Printpagewidth;
      //ts := ts+ ((tw-textwidth(RecName)) shr 1);
      ts := ts +  10;

      Textout(ts, -50, RecName);

      font.assign(savefont);

      savefont.free;
   end;

end;

procedure TMainDlg.asg_WorkRptPrintStart(Sender: TObject;
  NrOfPages: Integer; var FromPage, ToPage: Integer);
begin
   Printdialog1.FromPage   := Frompage;
   Printdialog1.ToPage     := ToPage;
   Printdialog1.Maxpage    := ToPage;
   Printdialog1.minpage    := 1;

   if Printdialog1.execute then
   begin
      Frompage := Printdialog1.FromPage;
      Topage   := Printdialog1.ToPage;
   end
   else
   begin
      Frompage := 0;
      Topage   := 0;
   end;
end;

procedure TMainDlg.fsbt_PrintClick(Sender: TObject);
var
   //PagePreview: TPagePreview;
   varResult  : String;
   tmp_AnalTo : String;
begin
   if not IsLogonUser then
   begin
      MessageBox(self.Handle,
                 PChar('[비상연락망] - [담당자정보]에 사용자 IP 등록후 사용하실 수 있습니다.'),
                 PChar(Self.Caption + ' : 등록되지 않은 User 제한 알림'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;


   //-------------------------------------------------------------------
   // 2. Set Print Options
   //-------------------------------------------------------------------
   if (asg_WorkRpt.Visible) then
   begin
      // 조회 종료일자 Temp 변수에 임시보관
      tmp_AnalTo := fmed_AnalTo.Text;

      // 조회시작 당일만 출력되도록 개선 @ 2014.07.17 LSH
      fmed_AnalTo.Text := fmed_AnalFrom.Text;


      // AdvStringGrid.PrintSettings에서
      // AutoSizeRows 조건 체크 해제해야, 아래 Hide Column시 여러장 출력 방지됨.
      asg_WorkRpt.InsertRows(asg_WorkRpt.RowCount, 4);
      asg_WorkRpt.Cells[0, asg_WorkRpt.RowCount - 4] := '비 고';
      asg_WorkRpt.Cells[0, asg_WorkRpt.RowCount - 3] := '결재';
      asg_WorkRpt.Cells[1, asg_WorkRpt.RowCount - 3] := '담 당';
      asg_WorkRpt.Cells[1, asg_WorkRpt.RowCount - 2] := '파트장';
      asg_WorkRpt.Cells[1, asg_WorkRpt.RowCount - 1] := '팀 장';

      // 비고 항목 Merge Cell 생성
      asg_WorkRpt.MergeCells(0, asg_WorkRpt.RowCount - 4, 3, 1);
      asg_WorkRpt.MergeCells(3, asg_WorkRpt.RowCount - 4, 2, 1);

      // 결재 항목 Merge Cell 생성
      asg_WorkRpt.MergeCells(0, asg_WorkRpt.RowCount - 3, 1, 3);

      // 담당/파트장/팀장 항목 Merge Cell 생성
      asg_WorkRpt.MergeCells(1, asg_WorkRpt.RowCount - 3, 2, 1);
      asg_WorkRpt.MergeCells(3, asg_WorkRpt.RowCount - 3, 2, 1);
      asg_WorkRpt.MergeCells(1, asg_WorkRpt.RowCount - 2, 2, 1);
      asg_WorkRpt.MergeCells(3, asg_WorkRpt.RowCount - 2, 2, 1);
      asg_WorkRpt.MergeCells(1, asg_WorkRpt.RowCount - 1, 2, 1);
      asg_WorkRpt.MergeCells(3, asg_WorkRpt.RowCount - 1, 2, 1);

      // 확인란 Visible
      asg_WorkRpt.ColWidths[C_WR_CONFIRM]   := 30;

      // 작성일자는 Invisible
      asg_WorkRpt.HideColumn(C_WR_REGDATE);




      SetPrintOptions;




      {   -- 미리보기 Text 용
      //-------------------------------------------------------------------
      // 2. Set Print Options
      //-------------------------------------------------------------------
      SetPrintOptions;

      //-------------------------------------------------------------------
      // 3. Create Preview Object
      //-------------------------------------------------------------------
      asg_WorkRpt.PreviewPage := 1;
      Pagepreview                := tpagepreview.CreatePreview(self, asg_WorkRpt);

      //-------------------------------------------------------------------
      // 4. Show Preview Modal
      //-------------------------------------------------------------------
      try
         PagePreview.ShowModal;

      finally
         PagePreview.Free;
      end;
      }


      //-------------------------------------------------------------------
      // 3. Print
      //-------------------------------------------------------------------
      asg_WorkRpt.Print;





      try
         asg_WorkRpt.RemoveRows(asg_WorkRpt.RowCount-4, 4);
         asg_WorkRpt.UnHideColumn(C_WR_REGDATE);
         asg_WorkRpt.ColWidths[C_WR_CONFIRM]   := 0;

      except
         // Grid Index Out of Range 에러 발생되는 부분.. 출력물 이상없고,
         // AdvStringGrid 버그로 추정되는데, 원인 분석이 잘 안됨.
      end;




      // 로그 Update
      UpdateLog('WORKRPT',
                'PRINT',
                FsUserIp,
                'P',
                FsUserSpec,
                FsUserNm,
                varResult
                );



      // Comments
      lb_Analysis.Caption := '▶ 업무일지 출력이 완료되었습니다.';

      // 조회종료일 원상복귀
      fmed_AnalTo.Text := tmp_AnalTo;

   end
   else if (apn_Weekly.Visible) then
   begin

      SetPrintOptions;




      {   -- 미리보기 Text 용
      //-------------------------------------------------------------------
      // 2. Set Print Options
      //-------------------------------------------------------------------
      SetPrintOptions;

      //-------------------------------------------------------------------
      // 3. Create Preview Object
      //-------------------------------------------------------------------
      asg_WorkRpt.PreviewPage := 1;
      Pagepreview                := tpagepreview.CreatePreview(self, asg_WorkRpt);

      //-------------------------------------------------------------------
      // 4. Show Preview Modal
      //-------------------------------------------------------------------
      try
         PagePreview.ShowModal;

      finally
         PagePreview.Free;
      end;
      }


      //-------------------------------------------------------------------
      // 3. Print
      //-------------------------------------------------------------------
      asg_WeeklyRpt.Print;







      // 로그 Update
      UpdateLog('WEEKRPT',
                'PRINT',
                FsUserIp,
                'P',
                FsUserSpec,
                FsUserNm,
                varResult
                );



      // Comments
      lb_Analysis.Caption := '▶ 업무리포트 출력이 완료되었습니다.';

   end;
end;



procedure TMainDlg.asg_WeeklyRptPrintPage(Sender: TObject; Canvas: TCanvas;
  PageNr, PageXSize, PageYSize: Integer);
var
   savefont : tfont;
   ts{,tw}    : integer;
   RecName  : String;
const
   myowntitle : string = ' ';
begin
   if asg_WeeklyRpt.PrintColStart <> 0 then exit;


   // Report Name 정의
   if (Trim(fed_Analysis.Text) <> '') then
      RecName := Trim(fed_Analysis.Text) + '의 업무리포트'
   else
      RecName := '업무리포트';



   // Title에 작성일자 + 요일 표기
   if (fmed_AnalTo.Text > fmed_AnalFrom.Text) then
      RecName := RecName + ' [' + fmed_AnalFrom.Text + ', '  + GetDayofWeek(StrToDate(fmed_AnalFrom.Text), 'HS') + ' ~ ' + fmed_AnalTo.Text + ', ' + GetDayofWeek(StrToDate(fmed_AnalTo.Text), 'HS') +']'
   else
      RecName := RecName + ' [' + fmed_AnalFrom.Text + ', '  + GetDayofWeek(StrToDate(fmed_AnalFrom.Text), 'HS') + ']';



   // 의료원명 Display
   with canvas do
   begin
      savefont := TFont.Create;
      savefont.assign(font);

      font.name   := 'Arial';
      font.Size   := 10;
      font.Style  := [];
      font.height := asg_WeeklyRpt.mapfontheight(10);
      font.Color  := clBlack;

      ts := asg_WeeklyRpt.Printcoloffset[0];
//      tw := asg_WeeklyRpt.Printpagewidth;
      //ts := ts+ ((tw - textwidth('고려대학교의료원')) shr 1);
      ts := ts +  10;

      Textout(ts, -15, '고려대학교의료원 정보전산팀');

      font.assign(savefont);

      savefont.free;
   end;


   // 기록지명 Display
   with canvas do
   begin
      savefont := TFont.Create;
      savefont.assign(font);

      font.name   := 'Arial';
      font.Size   := 18;
      font.Style  := [fsBold];
      font.height := asg_WeeklyRpt.mapfontheight(18);
      font.Color  := clBlack;

      ts := asg_WeeklyRpt.Printcoloffset[0];
//      tw := asg_WeeklyRpt.Printpagewidth;
      //ts := ts+ ((tw-textwidth(RecName)) shr 1);
      ts := ts +  10;

      Textout(ts, -50, RecName);

      font.assign(savefont);

      savefont.free;
   end;
end;



procedure TMainDlg.asg_WeeklyRptPrintStart(Sender: TObject;
  NrOfPages: Integer; var FromPage, ToPage: Integer);
begin
   Printdialog1.FromPage   := Frompage;
   Printdialog1.ToPage     := ToPage;
   Printdialog1.Maxpage    := ToPage;
   Printdialog1.minpage    := 1;

   if Printdialog1.execute then
   begin
      Frompage := Printdialog1.FromPage;
      Topage   := Printdialog1.ToPage;
   end
   else
   begin
      Frompage := 0;
      Topage   := 0;
   end;
end;



procedure TMainDlg.fsbt_DialMapClick(Sender: TObject);
begin
   apn_DialMap.Top     := 59;
   apn_DialMap.Left    := 120;

   apn_DialMap.Collaps := True;
   apn_DialMap.Visible := True;
   apn_DialMap.Collaps := False;

   // 나의 근무처 다이얼 MAP 조회
   SelGridInfo('DIALMAP')
end;




procedure TMainDlg.asg_DialMapGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   // Horizontal Alignment
   HAlign := taCenter;

   // Vertical Alignment
   VAlign := vtaCenter;
end;



procedure TMainDlg.asg_AnalysisKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  Res : TPoint;
begin

   case (Key) of

      //[F3 Key] : 다음 찾기
      Ord(VK_F3) :
                  begin
                     Res := asg_Analysis.FindNext;

                     if (Res.X >= 0) and
                        (Res.Y >= 0) then
                     begin
                        asg_Analysis.Col := Res.X;
                        asg_Analysis.Row := Res.Y;
                     end
                     else
                        MessageBox(self.Handle,
                                   PChar('찾으시는 항목이 화면 List에 존재하지 않습니다.'),
                                   PChar('S/R 검색 : 결과내 재검색 결과 알림'),
                                   MB_OK + MB_ICONWARNING)
                                   ;
                  end;

      //[Ctrl + F] : 검색 EditBox 포커싱 @ 2016.11.18 LSH
      Ord('F') :
                  if (ssCtrl in Shift) then
                  begin
                     fcb_ReScan.Checked := True;

                     fed_Analysis.SetFocus;
                  end;


      //[Enter] : 해당 S/R 상세내역 Panel-On @ 2016.11.18 LSH
      Ord(VK_RETURN) :
                     begin
                        asg_AnalysisDblClick(Sender);
                     end;


      {if (ssCtrl in Shift) then
                 begin
                    asg_WorkRpt.CopySelectionToClipboard;
                 end;

       }

   end;
end;




//------------------------------------------------------------------------------
// AdvStringGrid onMouseMove Event Handler
//
// Date   : 2013.11.18
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_MasterMouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
var
   NowCol, NowRow : Integer;
begin
   with asg_Master do
   begin
      // 현재 Cell 좌표 가져오기
      MouseToCell(X, Y, NowCol, NowRow);


      if (NowRow > 0) and
         (NowCol = 2) then
      begin
         ShowHint := True;
      end
      else
         ShowHint := False;
   end;
end;



procedure TMainDlg.fsbt_AddMyDialClick(Sender: TObject);
begin
   if not IsLogonUser then
   begin
      MessageBox(self.Handle,
                 PChar('[비상연락망] - [담당자정보]에 사용자 IP 등록후 사용하실 수 있습니다.'),
                 PChar(Self.Caption + ' : 등록되지 않은 User 제한 알림'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;

   with asg_AddMyDial do
   begin
      ClearCols(1, ColCount);
      RowCount  := 8;

      Cells[1, 0] := '협력업체';
      Cells[1, 1] := '';
      Cells[1, 2] := '';
      Cells[1, 3] := '';
      Cells[1, 4] := '';
      Cells[1, 5] := '';
      Cells[1, 6] := '';
      Cells[1, 7] := '필수항목 입력후 이 칸에 등록버튼이 생성됨.';

      Top       := 147;
      Left      := 143;
   end;


   apn_AddMyDial.Collaps   := True;
   apn_AddMyDial.Visible   := True;
   apn_AddMyDial.Collaps   := False;

end;



procedure TMainDlg.asg_AddMyDialGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   // Horizontal Alignment
   if (ACol = 0) then
      HAlign := taCenter;

   // Vertical Alignment
   VAlign := vtaCenter;
end;



procedure TMainDlg.asg_AddMyDialKeyPress(Sender: TObject; var Key: Char);
begin
   //--------------------------------------------------------
   // 엔터 입력시, 다음 Cell 진행여부 분기
   //--------------------------------------------------------
   if (Key = #13) then
   begin
      if asg_AddMyDial.Row = 6 then
         asg_AddMyDial.Navigation.AdvanceOnEnter := False
      else
         asg_AddMyDial.Navigation.AdvanceOnEnter := True;
   end;
end;



procedure TMainDlg.asg_AddMyDialEditingDone(Sender: TObject);
begin
   with asg_AddMyDial do
   begin
      if (Trim(Cells[Col, 0]) <> '') and
         (Trim(Cells[Col, 1]) <> '') and
         (Trim(Cells[Col, 2]) <> '') and
         (Trim(Cells[Col, 3]) <> '')
      then
      begin
         AddButton(Col,
                   7,
                   ColWidths[Col]-5,
                   20,
                   '나의 다이얼 Book 신규등록',
                   haBeforeText,
                   vaCenter);
      end
      else
         RemoveButton(Col, 7);
   end;
end;



procedure TMainDlg.asg_AddMyDialButtonClick(Sender: TObject; ACol,
  ARow: Integer);
var
   iChoice     : Integer;
   TpAddDial   : TTpSvc;
   j, iCnt  : Integer;
   vType1,
   vLocate,
   vUserIp,
   vDeptNm,
   vDeptSpec,
   vCallNo,
   vDutyUser,
   vMobile,
   vDutyRmk,
   vDataBase,
   vLinkSeq,
   vDelDate,
   vEditIp : Variant;
begin

   // 1-1. 필수입력 항목 확인
   with asg_AddMyDial do
   begin
      if (Trim(Cells[Col, 0]) = '') or
         (Trim(Cells[Col, 1]) = '') or
         (Trim(Cells[Col, 2]) = '') or
         (Trim(Cells[Col, 3]) = '')
      then
      begin
         MessageBox(Self.Handle,
                    PChar('[연락처] 까지는 필수입력 항목입니다.'),
                    '나의 다이얼 Book 신규 등록전 필수입력 확인',
                    MB_OK + MB_ICONWARNING);

         Exit;
      end;
   end;




   // 1-2. 최종 등록 여부 확인
   iChoice := MessageBox(Self.Handle,
                         PChar('작성하신 신규 다이얼을 등록하시겠습니까?'),
                         '나의 다이얼 Book 신규 등록 확인',
                         MB_YESNO + MB_ICONQUESTION);


   if iChoice = IDNO then Exit;


   iCnt := 0;


   //-------------------------------------------------------------------
   // 2. Create Variables
   //-------------------------------------------------------------------
   with asg_AddMyDial do
   begin
      for j := ACol to ACol do
      begin
         //------------------------------------------------------------------
         // Append Variables
         //------------------------------------------------------------------
         AppendVariant(vType1,      '9'        );
         AppendVariant(vLocate,     Cells[j, 0]);
         AppendVariant(vUserIp,     FsUserIp   );
         AppendVariant(vDeptNm,     Cells[j, 1]);
         AppendVariant(vDeptSpec,   Cells[j, 2]);
         AppendVariant(vCallNo,     Cells[j, 3]);
         AppendVariant(vDutyUser,   Cells[j, 4]);
         AppendVariant(vMobile,     Cells[j, 5]);
         AppendVariant(vDutyRmk,    Cells[j, 6]);
         AppendVariant(vDataBase,   FsUserNm + 'Book');
         AppendVariant(vLinkSeq,    Cells[j, 3]);
         AppendVariant(vDelDate,    ''         );
         AppendVariant(vEditIp,     FsUserIp   );

         Inc(iCnt);
      end;
   end;


   if iCnt <= 0 then
      Exit;

   //-------------------------------------------------------------------
   // 3. Insert by TpSvc
   //-------------------------------------------------------------------
   TpAddDial := TTpSvc.Create;
   TpAddDial.Init(Self);

   Screen.Cursor := crHourGlass;

   try
      if TpAddDial.PutSvc('MD_KUMCM_M1',
                          [
                            'S_TYPE1'   , vType1
                         ,  'S_STRING1' , vLocate
                         ,  'S_STRING7' , vUserIp
                         ,  'S_STRING42', vDeptNm
                         ,  'S_STRING43', vDeptSpec
                         ,  'S_STRING15', vCallNo
                         ,  'S_STRING13', vDutyUser
                         ,  'S_STRING5' , vMobile
                         ,  'S_STRING12', vDutyRmk
                         ,  'S_STRING44', vDataBase
                         ,  'S_STRING45', vLinkSeq
                         ,  'S_STRING10', vDelDate
                         ] ) then
      begin
         MessageBox(self.Handle,
                    '나의 다이얼 Book에 정상적으로 [등록]되었습니다.',
                    '나의 다이얼 Book 등록완료 알림',
                    MB_OK + MB_ICONINFORMATION);
      end
      else
      begin
         ShowMessage(GetTxMsg);
      end;



   finally
      FreeAndNil(TpAddDial);

      apn_AddMyDial.Collaps := True;
      apn_AddMyDial.Visible := False;

      Screen.Cursor  := crDefault;
   end;


   //---------------------------------------------------------
   // 4. 나의 다이얼 Book List 조회
   //---------------------------------------------------------
   SelGridInfo('MYDIAL');

end;




procedure TMainDlg.asg_MyDialClick(Sender: TObject);
begin
   if (apn_AddMyDial.Visible) then
   begin
      apn_AddMyDial.Collaps := True;
      apn_AddMyDial.Visible := False;
   end;

   if (apn_DialMap.Visible) then
   begin
      apn_DialMap.Collaps := True;
      apn_DialMap.Visible := False;
   end;
end;




procedure TMainDlg.asg_NetworkMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
var
   NowCol, NowRow : Integer;
   ServerFileName, ClientFileName: String;
   FTP_SVRIP, FTP_USERID, FTP_PASSWD, FTP_HOSTNAME, FTP_DIR : String;
   sRemoteFile, sLocalFile : String;
begin
   with asg_Network do
   begin
      // 현 Mouse 포인터 좌표 받아오기
      MouseToCell(X,
                  Y,
                  NowCol,
                  NowRow);


      // 담당자 프로필 조회
      if NowCol = C_NW_USERNM then
      begin

         if (NowRow > 0) and
            (IsLogonUser) then
         begin
            lb_Network.Caption := '▶ ' + Trim(asg_UserProfile.Cells[1, R_PR_USERNM]) + '님의 담당자 프로필 조회중.';

            if apn_UserProfile.Visible = False then
            begin
               if (Height - CellRect(C_NW_USERNM, NowRow).Top) > (apn_UserProfile.Height + 110) then
               begin
                  apn_UserProfile.Top  := CellRect(C_NW_USERNM, NowRow).Top;
                  apn_UserProfile.Left := 253;
               end
               else
               begin
                  apn_UserProfile.Top  := CellRect(C_NW_USERNM, NowRow).Top - apn_UserProfile.Height;
                  apn_UserProfile.Left := 253;
               end;



               // Location & Display
               {
               apn_UserProfile.Top        := CellRect(C_NW_USERNM, NowRow).Top + 77;
               apn_UserProfile.Left       := 250;
               }

               apn_UserProfile.Caption.Color := $0095B1C1;
               asg_UserProfile.FixedColor    := $0095B1C1;
               apn_UserProfile.Collaps       := True;
               apn_UserProfile.Visible       := True;
               apn_UserProfile.Collaps       := False;
            end;

            // 프로필 내역 assign
            asg_UserProfile.Cells[1, R_PR_USERNM]     := Cells[C_NW_USERNM,   NowRow];
            asg_UserProfile.Cells[1, R_PR_DUTYPART]   := Cells[C_NW_LOCATE,   NowRow]  + ' ' +
                                                         Cells[C_NW_DUTYPART, NowRow]  + ' ' +
                                                         Cells[C_NW_DUTYSPEC, NowRow];
            asg_UserProfile.Cells[1, R_PR_REMARK]     := Cells[C_NW_DUTYRMK,  NowRow];
            asg_UserProfile.Cells[1, R_PR_CALLNO]     := Cells[C_NW_CALLNO,   NowRow];

            asg_UserProfile.AddButton(1, R_PR_DUTYSCH, asg_UserProfile.ColWidths[0]+45, 20, '최근 업무 History 보기', haBeforeText, vaCenter);


            // 프로필 이미지 FTP 파일내역이 존재하면, 이미지 Download
            if (Cells[C_NW_PHOTOFILE, NowRow] <> '') then
            begin
               // MIS 기본 이미지 정보인 경우.
               if Cells[C_NW_HIDEFILE, NowRow] = '' then
               begin

                  // 담당자 프로필 이미지 파일명
                  ServerFileName := Cells[C_NW_PHOTOFILE, NowRow];


                  // Local 파일이름 Set
                  // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
                  ClientFileName := G_HOMEDIR + 'TEMP\SPOOL\' + ServerFileName;



                  // Local에 해당 Image 존재유무 체크
                  // Local file의 Size 체크 추가 (Local에 해당 사번.jpg가 0 size로 존재하면 JPEG Error #41 떨어짐) @ 2015.06.03 LSH
                  if (FileExists(ClientFileName)) and
                     (CMsg.GetFileSize(ClientFileName) >  0) then
                  begin
                     // FTP 서버정보 가져오기
                     if Not GetBinUploadInfo(FTP_SVRIP,
                                             FTP_USERID,
                                             FTP_PASSWD,
                                             FTP_HOSTNAME,
                                             FTP_DIR) then
                     begin
                        //ShowMessage('다운로드가 실패하여 실행할 수 없습니다.');
                        TUXFTP := nil;
                        Exit;
                     end;


                     // FTP 서버 IP Set
                     FTP_SVRIP        := C_MIS_FTP_IP;


                     // FTP 계정정보 Set
                     FTP_USERID := '';
                     FTP_PASSWD := '';
                     FTP_DIR    := '/kumis/app/mis/media/cq/photo/';



                     // Image 다운로드
                     if Not GetBINFTP(FTP_SVRIP,
                                      FTP_USERID,
                                      FTP_PASSWD,
                                      FTP_DIR + ServerFileName,
                                      ClientFileName,
                                      False) then
                     begin

                        //Showmessage('이미지 다운로드에 실패했습니다. 다시한번 시도해 주십시오');

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
                     asg_UserProfile.CreatePicture(1,
                                                   R_PR_PHOTO,
                                                   False,
                                                   ShrinkWithAspectRatio,
                                                   0,
                                                   haLeft,
                                                   vaTop).LoadFromFile(ClientFileName);
                  except
                     Exit;
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
                  if PosByte('/ftpspool/KDIALFILE/', Cells[C_NW_HIDEFILE, NowRow]) > 0 then
                     sRemoteFile := Cells[C_NW_HIDEFILE, NowRow]
                  else
                     sRemoteFile := '/ftpspool/KDIALFILE/' + Cells[C_NW_HIDEFILE, NowRow];

                  // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
                  sLocalFile  := G_HOMEDIR + 'TEMP\SPOOL\' + Cells[C_NW_PHOTOFILE, NowRow];



                  if (GetBINFTP(FTP_SVRIP, FTP_USERID, FTP_PASSWD, sRemoteFile, sLocalFile, False)) then
                  begin
                     //	정상적인 FTP 다운로드
                  end;

                  if (FileExists(sLocalFile)) and
                     (CMsg.GetFileSize(sLocalFile) > 0) then
                  begin
                     // 이미지만 Preview 설정
                     if (PosByte('.bmp', sLocalFile) > 0) or
                        (PosByte('.BMP', sLocalFile) > 0) or
                        (PosByte('.jpg', sLocalFile) > 0) or
                        (PosByte('.JPG', sLocalFile) > 0) or
                        {  -- TMS Grid에서 png 타입 미인식 오류로 주석 @ 2017.11.01 LSH
                        (PosByte('.png', sLocalFile) > 0) or
                        (PosByte('.PNG', sLocalFile) > 0) or
                        }
                        (PosByte('.gif', sLocalFile) > 0) or
                        (PosByte('.GIF', sLocalFile) > 0) then
                        // 수신한 Image 파일을 Grid에 표기
                        asg_UserProfile.CreatePicture(1,
                                                      R_PR_PHOTO,
                                                      False,
                                                      ShrinkWithAspectRatio,
                                                      0,
                                                      haLeft,
                                                      vaTop).LoadFromFile(sLocalFile)
                     else
                        asg_UserProfile.Cells[1, R_PR_PHOTO] := asg_Network.Cells[C_NW_PHOTOFILE, NowRow];
                  end;
               end;
            end
            else
            begin
               // 담당자 프로필 이미지 파일명
               if Cells[C_NW_USERID, NowRow] = 'XXXXX' then
               begin
                  // 프로필 이미지 Remove
                  asg_UserProfile.RemovePicture(1, R_PR_PHOTO);

                  // 이미지 등록 버튼 생성
                  asg_UserProfile.AddButton(1, R_PR_PHOTO, asg_UserProfile.ColWidths[0]+45, asg_UserProfile.RowHeights[R_PR_PHOTO]-5, '프로필 이미지 등록', haBeforeText, vaCenter);

                  Exit;
               end
               else
                  // 담당자 HRM 프로필 이미지 파일명
                  ServerFileName := Cells[C_NW_USERID, NowRow] + '.jpg';

                  // Local 파일이름 Set
                  // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
                  ClientFileName := G_HOMEDIR + 'TEMP\SPOOL\' + ServerFileName;


                  // Local에 해당 Image 존재유무 체크
                  // Local file의 Size 체크 추가 (Local에 해당 사번.jpg가 0 size로 존재하면 JPEG Error #41 떨어짐) @ 2015.06.03 LSH
                  if (FileExists(ClientFileName)) and
                     (CMsg.GetFileSize(ClientFileName) >  0) then
                  begin
                     // FTP 서버정보 가져오기
                     if Not GetBinUploadInfo(FTP_SVRIP,
                                             FTP_USERID,
                                             FTP_PASSWD,
                                             FTP_HOSTNAME,
                                             FTP_DIR) then
                     begin
                        //ShowMessage('다운로드가 실패하여 실행할 수 없습니다.');
                        TUXFTP := nil;
                        Exit;
                     end;


                     // FTP 서버 IP Set
                     FTP_SVRIP        := C_MIS_FTP_IP;

                     // FTP 계정정보 Set
                     FTP_USERID := '';
                     FTP_PASSWD := '';
                     FTP_DIR    := '/kuhrm/app/hrm/photo/';


                     // HRM Image 다운로드
                     if Not GetBINFTP(FTP_SVRIP,
                                      FTP_USERID,
                                      FTP_PASSWD,
                                      FTP_DIR + ServerFileName,
                                      ClientFileName,
                                      False) then
                     begin
                        //Showmessage('이미지 다운로드에 실패했습니다. 다시한번 시도해 주십시오');

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
                     asg_UserProfile.CreatePicture(1,
                                                   R_PR_PHOTO,
                                                   False,
                                                   ShrinkWithAspectRatio,
                                                   0,
                                                   haLeft,
                                                   vaTop).LoadFromFile(ClientFileName);
                  except
                     Exit;
                  end;
            end;
         end;
      end
      else
      begin
         lb_Network.Caption := '▶ [담당자정보]와 [업무프로필]을 업데이트 하시면 비상연락망에 자동조회 됩니다.';

         apn_UserProfile.Collaps := True;
         apn_UserProfile.Visible := False;
      end;
   end;

end;




//----------------------------------------------------------------
// String 내의 특정 문자열 삭제
//    - 소스출처 : MComFunc.pas
//----------------------------------------------------------------
function TMainDlg.DeleteStr(OrigStr,DelStr : String) : String;
begin
   while PosByte(DelStr, OrigStr) > 0 do
      System.Delete(OrigStr, PosByte(DelStr, OrigStr), LengthByte(DelStr));

   Result := OrigStr;
end;




procedure TMainDlg.asg_UserProfileGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   // Horizontal Alignment
   if (ACol = 0)
   {or ((ACol > 0) and ((ARow = R_PR_USERNM) or
                                     (ARow = R_PR_DUTYPART) or
                                     (ARow = R_PR_CALLNO) or
                                     (ARow = R_PR_DUTYSCH)
                                     ))
                                     }
   then
      HAlign := taCenter;

   // Vertical Alignment
   VAlign := vtaCenter;
end;

procedure TMainDlg.asg_DocSpecGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   // Horizontal Alignment
   if (ACol = 0) then
      HAlign := taCenter;

   // Vertical Alignment
   VAlign := vtaCenter;

end;

procedure TMainDlg.asg_DocSpecKeyPress(Sender: TObject; var Key: Char);
begin
   if (Key = #13) then
   begin
      if asg_DocSpec.Row = R_DS_USELVL then
         asg_DocSpec.Navigation.AdvanceOnEnter := False
      else
         asg_DocSpec.Navigation.AdvanceOnEnter := True;
   end;
end;

procedure TMainDlg.asg_DocSpecEditingDone(Sender: TObject);
begin
   if (PosByte('등록', apn_DocSpec.Caption.Text) > 0) or
      (PosByte('수정', apn_DocSpec.Caption.Text) > 0) then
   begin
      with asg_DocSpec do
      begin
         if (Trim(Cells[Col, R_DS_USERID]) <> '') and
            (Trim(Cells[Col, R_DS_USERNM]) <> '') and
            (Trim(Cells[Col, R_DS_DEPTCD]) <> '') and
            //(Trim(Cells[Col, R_DS_DEPTNM]) <> '') and
            (Trim(Cells[Col, R_DS_SOCNO])  <> '') and
            (Trim(Cells[Col, R_DS_STARTDT])<> '') and
            (Trim(Cells[Col, R_DS_MOBILE]) <> '') and
            (Trim(Cells[Col, R_DS_GRPID])  <> '') and
            (Trim(Cells[Col, R_DS_USELVL]) <> '')
         then
         begin
            AddButton(Col,
                      R_DS_BUTTON,
                      ColWidths[Col]-5,
                      20,
                      asg_RegDoc.Cells[C_RD_DOCLIST, asg_RegDoc.Row] + ' 권한 등록',
                      haBeforeText,
                      vaCenter);
         end
         else
            RemoveButton(Col, R_DS_BUTTON);
      end;
   end;
end;

procedure TMainDlg.asg_DocSpecButtonClick(Sender: TObject; ACol,
  ARow: Integer);
begin
   with asg_RegDoc do
   begin
      // CRA/CRC ID 관리
      if (PosByte('CRA', Cells[C_RD_DOCLIST, Row]) > 0) or
         (PosByte('CRC', Cells[C_RD_DOCLIST, Row]) > 0) then
      begin
         if (asg_DocSpec.Cells[1, R_DS_DEPTCD] = 'CTC') then
            Cells[C_RD_DOCTITLE, Row] := 'CRA OCS 사용권한 요청'
         else
            Cells[C_RD_DOCTITLE, Row] := 'CRC OCS 사용권한 요청';

         Cells[C_RD_REGDATE, Row] := FormatDateTime('yyyy-mm-dd', Date);
         Cells[C_RD_REGUSER, Row] := asg_DocSpec.Cells[1, R_DS_USERNM];
         Cells[C_RD_RELDEPT, Row] := asg_DocSpec.Cells[1, R_DS_DEPTNM];
         Cells[C_RD_DOCRMK,  Row] := asg_DocSpec.Cells[1, R_DS_USERID];

      end
      // 유지보수 계약처 관리
      else if (PosByte('계약', Cells[C_RD_DOCLIST, Row]) > 0) then
      begin
         Cells[C_RD_DOCTITLE, Row] := asg_DocSpec.Cells[1, R_DS_USERID];   // 계약명
         Cells[C_RD_REGDATE,  Row] := FormatDateTime('yyyy-mm-dd', Date);  // 등록일
         Cells[C_RD_REGUSER,  Row] := FsUserNm;                            // 등록자
         Cells[C_RD_RELDEPT,  Row] := asg_DocSpec.Cells[1, R_DS_USERNM];   // 업체명
         Cells[C_RD_DOCRMK,   Row] := asg_DocSpec.Cells[1, R_DS_DEPTCD];   // 계약기간
         Cells[C_RD_DOCLOC,   Row] := asg_DocSpec.Cells[1, R_DS_USELVL];   // 계약관련 문서위치(ex. PDF, 계약처..)
      end;
   end;

   apn_DocSpec.Collaps := True;
   apn_DocSpec.Visible := False;

end;


//-----------------------------------------------------
// User별 맞춤형 알람 Check
//-----------------------------------------------------
function TMainDlg.CheckUserAlarm(in_Gubun,
                                 in_UserNm,
                                 in_UserIp,
                                 in_Option : String) : Boolean;
var
   TpGetAlarm : TTpSvc;
   sType1, sType2, sType3, sType4, sType5, sType6 : String;
   i, iRowCnt : Integer;
   FTP_SVRIP, FTP_USERID, FTP_PASSWD, FTP_HOSTNAME, FTP_DIR, ClientFileName : String;
begin
   // Init. Return Values
   Result := False;


   //-----------------------------------------------------------------
   // 1. Set Params.
   //-----------------------------------------------------------------
   sType1 := '27';
   sType2 := in_Gubun;
   sType3 := in_UserNm;
   sType4 := in_UserIp;
   sType5 := in_Option;



   //-----------------------------------------------------------------
   // 2. 조회
   //-----------------------------------------------------------------
   Screen.Cursor := crHourGlass;





   //-----------------------------------------------------------------
   // 2-1. Load Data by TpSvc
   //-----------------------------------------------------------------
   TpGetAlarm := TTpSvc.Create;
   TpGetAlarm.Init(Self);


   Screen.Cursor := crHourglass;


   try
      TpGetAlarm.CountField  := 'S_CODE3';
      TpGetAlarm.ShowMsgFlag := False;

      if TpGetAlarm.GetSvc('MD_KUMCM_S1',
                          ['S_TYPE1'  , sType1
                         , 'S_TYPE2'  , sType2
                         , 'S_TYPE3'  , sType3
                         , 'S_TYPE4'  , sType4
                         , 'S_TYPE5'  , sType5
                         , 'S_TYPE6'  , sType6
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
                          ]) then

         if TpGetAlarm.RowCount < 0 then
         begin
            ShowMessage(GetTxMsg);
            Exit;
         end
         else if TpGetAlarm.RowCount = 0 then
         begin
            Exit;
         end;



      // 구분 Flag 별 Return Value 분기
      if (in_Gubun = 'BORN') then      // 생일 축하
      begin
         if (TpGetAlarm.GetOutputDataS('sLogYn', 0) = 'Y') then
            Result := False
         else
         begin
            asg_Congrats.Top     := 405;
            asg_Congrats.Left    := 136;
            asg_Congrats.Visible := True;

            asg_NotiImg.Top      := 999;
            asg_NotiImg.Left     := 999;
            asg_NotiImg.Visible  := False;

            lb_InformMessage.Top := 999;
            lb_InformMessage.Top := 999;
            lb_InformMessage.Visible := False;

            iRowCnt := TpGetAlarm.RowCount;
            asg_Congrats.RowCount := iRowCnt + asg_Congrats.FixedRows + 1;

            Result := True;

            fpn_Title.Caption := '생일 축하합니다 ♬';

            fpn_Photo.Visible := True;

            fani_Photo.BringToFront;

            for i := 0 to iRowCnt - 1 do
            begin
               asg_Congrats.Cells[0, i + asg_Congrats.FixedRows] := TpGetAlarm.GetOutputDataS('sLocate',  i) +
                                                                    TpGetAlarm.GetOutputDataS('sDutyPart',i) + ' ' +
                                                                    TpGetAlarm.GetOutputDataS('sUserNm',  i) + '님';

               asg_Congrats.AddButton(1, i + asg_Congrats.FixedRows, asg_Congrats.ColWidths[1] - 5, 20, '메세지 보내기', haBeforeText, vaCenter);

               asg_Congrats.Cells[2, i + asg_Congrats.FixedRows] := TpGetAlarm.GetOutputDataS('sMobile',  i);
            end;
         end;

         asg_Congrats.RowCount := asg_Congrats.RowCount - 1;

      end
      // 끼니별 병/식 정보
      else if (in_Gubun = 'Diet1Off') or
              (in_Gubun = 'Diet2Off') or
              (in_Gubun = 'Diet3Off') then
      begin
         // 해당 식이끼니의 알람Off-Log가 있으면(Y), True 리턴.
         if (TpGetAlarm.GetOutputDataS('sLogYn', 0) = 'Y') then
            Result := True
         else
            Result := False;
      end
      // 메인공지
      else if (in_Gubun = 'NOTI') then
      begin
         // 로그확인 또는 유효한 Noti. 공지 없으면, Exit.
         if (TpGetAlarm.GetOutputDataS('sLogYn', 0) = 'Y') or
            (TpGetAlarm.GetOutputDataS('sContext',  0) = '') then
            Result := False
         else
         begin
            // Top Align.
            if TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 1) = 'T' then
            begin
               asg_NotiImg.Top      := 105;
               asg_NotiImg.Left     := 136;
               asg_NotiImg.Height   := 281;
               asg_NotiImg.Width    := 409;
               asg_NotiImg.Visible  := False;
               asg_NotiImg.BringToFront;

               asg_NotiImg.DefaultRowHeight  := 282;

               asg_Congrats.Top     := 999;
               asg_Congrats.Left    := 999;
               asg_Congrats.Visible := False;

               lb_InformMessage.Top     := StrToInt(TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 6));
               lb_InformMessage.Left    := StrToInt(TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 7));
               lb_InformMessage.Caption := TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 8);
               lb_InformMessage.Visible := True;
               lb_InformMessage.BringToFront;
            end
            // Bottom Align.
            else if TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 1) = 'B' then
            begin
               asg_NotiImg.Top      := 233;
               asg_NotiImg.Left     := 136;
               asg_NotiImg.Height   := 281;
               asg_NotiImg.Width    := 409;
               asg_NotiImg.Visible  := True;
               asg_NotiImg.BringToFront;

               asg_NotiImg.DefaultRowHeight  := 282;

               asg_Congrats.Top     := 999;
               asg_Congrats.Left    := 999;
               asg_Congrats.Visible := False;

               lb_InformMessage.Top     := StrToInt(TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 6));
               lb_InformMessage.Left    := StrToInt(TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 7));
               lb_InformMessage.Caption := TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 8);
               lb_InformMessage.Visible := True;
            end
            // Left Align.
            else if TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 1) = 'L' then
            begin
               asg_NotiImg.Top      := 105;
               asg_NotiImg.Left     := 32;
               asg_NotiImg.Height   := 376;
               asg_NotiImg.Width    := 273;
               asg_NotiImg.Visible  := True;
               asg_NotiImg.BringToFront;

               asg_NotiImg.DefaultRowHeight  := 377;

               asg_Congrats.Top     := 999;
               asg_Congrats.Left    := 999;
               asg_Congrats.Visible := False;

               lb_InformMessage.Top     := StrToInt(TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 6));
               lb_InformMessage.Left    := StrToInt(TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 7));
               lb_InformMessage.Caption := TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 8);
               lb_InformMessage.Visible := True;
            end
            // Right Align.
            else if TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 1) = 'R' then
            begin
               asg_NotiImg.Top      := 105;
               asg_NotiImg.Left     := 368;
               asg_NotiImg.Height   := 376;
               asg_NotiImg.Width    := 273;
               asg_NotiImg.Visible  := True;
               asg_NotiImg.BringToFront;

               asg_NotiImg.DefaultRowHeight  := 377;

               asg_Congrats.Top     := 999;
               asg_Congrats.Left    := 999;
               asg_Congrats.Visible := False;

               lb_InformMessage.Top     := StrToInt(TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 6));
               lb_InformMessage.Left    := StrToInt(TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 7));
               lb_InformMessage.Caption := TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 8);
               lb_InformMessage.Visible := True;
            end;


            lb_InformMessage.Font.Name := TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 4);


            Result := True;


            // 공지 Title 있으면 표기, 없으면 null
            if TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 5) <> '' then
            begin
               fpn_Title.Visible    := True;

               fpn_Title.Caption    := TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 5);
               fpn_Title.Font.Name  := TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 4);
               fpn_Title.Font.Size  := StrToInt(TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 3));
               fpn_Title.Font.Style := [fsBold];
            end
            else
            begin
               fpn_Title.Visible    := False;
            end;


            fpn_Photo.Visible    := False;


            // Local 파일이름 Set
            // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
            ClientFileName := G_HOMEDIR + 'TEMP\SPOOL\' + TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 9);


            // Local에 해당 Image 존재유무 체크
            // Local file의 Size 체크 추가 (Local에 해당 사번.jpg가 0 size로 존재하면 JPEG Error #41 떨어짐) @ 2015.06.03 LSH
            if (FileExists(ClientFileName)) and
               (CMsg.GetFileSize(ClientFileName) >  0) then
            begin
               // FTP 서버정보 가져오기
               if Not GetBinUploadInfo(FTP_SVRIP,
                                       FTP_USERID,
                                       FTP_PASSWD,
                                       FTP_HOSTNAME,
                                       FTP_DIR) then
               begin
                  ShowMessage('FTP 계정정보 조회를 실패하여 실행할 수 없습니다.');
                  TUXFTP := nil;
                  Exit;
               end;


               // FTP 계정정보 Set
               FTP_USERID := '';
               FTP_PASSWD := '';
               FTP_DIR    := '/ftpspool/KDIALIMG/';


               // Image 다운로드
               if Not GetBINFTP(C_KDIAL_FTP_IP,
                                FTP_USERID,
                                FTP_PASSWD,
                                FTP_DIR + TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 9),
                                ClientFileName,
                                False) then
               begin
                  Showmessage('공지 이미지 다운로드에 실패했습니다. 다시한번 시도해 주십시오');

                  TUXFTP := nil;

                  Exit;
               end;
            end;

            try
               // D/B에서 가져온 값 FTP 서버에서 다운로드 및 표기
               asg_NotiImg.CreatePicture(0,
                                         0,
                                         False,
                                         ShrinkWithAspectRatio,
                                         0,
                                         haCenter,
                                         vaCenter).LoadFromFile(ClientFileName);
            except
               Exit;
            end;
         end;
      end;



   finally
      FreeAndNil(TpGetAlarm);
      Screen.Cursor := crDefault;
   end;
end;


procedure TMainDlg.FlatSpeedButton15Click(Sender: TObject);
var
   varResult : String;
begin
   if (fcb_Today.Checked = True) then
   begin
      // 로그인시 Noti. 로그업데이트 @ 2015.04.08 LSH
      if (lb_InformMessage.Visible) then
         UpdateLog('ALARM',
                   'NOTI',
                   FsUserIp,
                   'C',
                   FormatDateTime('yyyy-mm-dd', Date),
                   FsUserNm,
                   varResult
                   )
      else
         UpdateLog('ALARM',
                   'BORN',
                   FsUserIp,
                   'C',
                   FormatDateTime('yyyy-mm-dd', Date),
                   FsUserNm,
                   varResult
                   )
   end;

   apn_Congrats.Collaps := True;
   apn_Congrats.Visible := False;
end;

procedure TMainDlg.fsbt_DetailClick(Sender: TObject);
begin
   if (apn_Network.Visible) then
   begin
      apn_Network.Collaps := True;
      apn_Network.Visible := False;
   end;

   if (apn_Master.Visible) then
   begin
      apn_Master.Collaps := True;
      apn_Master.Visible := False;
   end;

   if (apn_Detail.Visible) then
   begin
      apn_Detail.Top     := 33;
      apn_Detail.Left    := 10;
      apn_Detail.Visible := True;
   end
   else
   begin
      apn_Detail.Top     := 33;
      apn_Detail.Left    := 10;
      apn_Detail.Collaps := True;
      apn_Detail.Visible := True;
      apn_Detail.Collaps := False;
   end;

   fsbt_NetworkRefresh.ColorFocused := $00ABA3DD;
   fsbt_Network.ColorFocused        := $00ABA3DD;
   fsbt_Detail.ColorFocused         := $00ABA3DD;
   fsbt_Master.ColorFocused         := $00ABA3DD;
   fsbt_NetworkPrint.ColorFocused   := $00ABA3DD;

   SelGridInfo('DETAIL');

   tm_Master.Enabled := False;

end;

procedure TMainDlg.fsbt_NetworkClick(Sender: TObject);
begin
   if (apn_Detail.Visible) then
   begin
      apn_Detail.Collaps := True;
      apn_Detail.Visible := False;
   end;

   if (apn_Master.Visible) then
   begin
      apn_Master.Collaps := True;
      apn_Master.Visible := False;
   end;

   apn_UserProfile.Visible := False;

   if (apn_Network.Visible) then
   begin
      apn_Network.Top     := 30;
      apn_Network.Left    := 7;
      apn_Network.Visible := True;
   end
   else
   begin
      apn_Network.Top     := 30;
      apn_Network.Left    := 7;
      apn_Network.Collaps := True;
      apn_Network.Visible := True;
      apn_Network.Collaps := False;
   end;

   fsbt_NetworkRefresh.ColorFocused := $008DB7DB;
   fsbt_Network.ColorFocused        := $008DB7DB;
   fsbt_Detail.ColorFocused         := $008DB7DB;
   fsbt_Master.ColorFocused         := $008DB7DB;
   fsbt_NetworkPrint.ColorFocused   := $008DB7DB;

   SelGridInfo('UPDATE');
   SelGridInfo('DIALOG');

   tm_Master.Enabled := False;
end;

procedure TMainDlg.fsbt_MasterClick(Sender: TObject);
begin
   if (apn_Network.Visible) then
   begin
      apn_Network.Collaps := True;
      apn_Network.Visible := False;
   end;

   if (apn_Detail.Visible) then
   begin
      apn_Detail.Collaps := True;
      apn_Detail.Visible := False;
   end;


   if (apn_Master.Visible) then
   begin
      apn_Master.Top     := 33;
      apn_Master.Left    := 10;
      apn_Master.Visible := True;
   end
   else
   begin
      apn_Master.Top     := 33;
      apn_Master.Left    := 10;
      apn_Master.Collaps := True;
      apn_Master.Visible := True;
      apn_Master.Collaps := False;
   end;

   fsbt_NetworkRefresh.ColorFocused := $00169B66;
   fsbt_Network.ColorFocused        := $00169B66;
   fsbt_Detail.ColorFocused         := $00169B66;
   fsbt_Master.ColorFocused         := $00169B66;
   fsbt_NetworkPrint.ColorFocused   := $00169B66;


   SelGridInfo('MASTER');

   tm_Master.Enabled := True;

end;

procedure TMainDlg.fsbt_ComDocClick(Sender: TObject);
begin

   // Comments
   lb_RegDoc.Caption := '▶ 부서공통문서(발송/협조 등)를 검색 또는 등록(수정)하실 수 있습니다.';

   if (apn_AnalList.Visible) then
   begin
      apn_AnalList.Collaps := True;
      apn_AnalList.Visible := False;
   end;

   if (apn_Release.Visible) then
   begin
      apn_Release.Collaps := True;
      apn_Release.Visible := False;
   end;

   if (apn_Duty.Visible) then
   begin
      apn_Duty.Collaps := True;
      apn_Duty.Visible := False;
   end;

   if (apn_ComDoc.Visible) then
   begin
      apn_ComDoc.Top     := 39;
      apn_ComDoc.Left    := 6;
      apn_ComDoc.Visible := True;
   end
   else
   begin
      apn_ComDoc.Top     := 39;
      apn_ComDoc.Left    := 6;
      apn_ComDoc.Collaps := True;
      apn_ComDoc.Visible := True;
      apn_ComDoc.Collaps := False;
   end;

   fsbt_Duty.ColorFocused     := $00E9D5B7;
   fsbt_ComDoc.ColorFocused   := $00E9D5B7;
   fsbt_Release.ColorFocused  := $00E9D5B7;

   // 릴리즈대장 업로드
   fsbt_Upload.Visible        := False;
   fsbt_Insert.Visible        := False;

   if IsLogonUser then
      SelGridInfo('DOC');

   tm_Master.Enabled := False;

end;

procedure TMainDlg.fsbt_ReleaseClick(Sender: TObject);
begin
   // Comments
   lb_RegDoc.Caption := '▶ 개발관련문서(릴리즈대장)를 검색 또는 등록(수정)하실 수 있습니다.';

   if (apn_AnalList.Visible) then
   begin
      apn_AnalList.Collaps := True;
      apn_AnalList.Visible := False;
   end;

   if (apn_ComDoc.Visible) then
   begin
      apn_ComDoc.Collaps := True;
      apn_ComDoc.Visible := False;
   end;

   if (apn_Duty.Visible) then
   begin
      apn_Duty.Collaps := True;
      apn_Duty.Visible := False;
   end;

   if (apn_Release.Visible) then
   begin
      apn_Release.Top     := 39;
      apn_Release.Left    := 6;
      apn_Release.Visible := True;
   end
   else
   begin
      apn_Release.Top     := 39;
      apn_Release.Left    := 6;
      apn_Release.Collaps := True;
      apn_Release.Visible := True;
      apn_Release.Collaps := False;
   end;

   fsbt_Duty.ColorFocused     := $00AFC3BD;
   fsbt_ComDoc.ColorFocused   := $00AFC3BD;
   fsbt_Release.ColorFocused  := $00AFC3BD;

   // 릴리즈대장 업로드 메뉴
   fsbt_Upload.Visible        := True;
   fsbt_Upload.ColorFocused   := $00AFC3BD;

   // 특정 관리자 모드 : 특정기간 릴리즈대장 선택적 업로드 기능
   if (PosByte('개발자IP', FsUserIp) > 0) then
   begin
      fsbt_Insert.ColorFocused   := $00AFC3BD;
      fsbt_Insert.Visible        := True;
   end;

   // Default
   fcb_DutySpec.Text := '진료';

   fmed_RegFrDt.Text := FormatDateTime('yyyy-mm-dd', Date - 3);
   fmed_RegToDt.Text := FormatDateTime('yyyy-mm-dd', Date    );   // Date - 1에서 Date (릴리즈 당일)까지로 변경 (H 샘경우, 릴리즈 당일 오전에 대장작성...) @ 2016.11.16 LSH

   fed_Release.Clear;


   if IsLogonUser then
   begin
      // 개발자는 소속 업무파트별 릴리즈 현황 Scan
      if PosByte('개발', FsUserPart) > 0 then
         SelGridInfo('RELEASE')
      else
         SelGridInfo('RELEASESCAN');
   end;

   tm_Master.Enabled := False;

   // 최초 [문서관리] Click시에만 릴리즈 대장 조회위한 Tagging 처리 @ 2015.06.15 LSH
   fsbt_Release.Tag := 1;

   // 검색 완료후 Grid 포커싱 @ 2016.11.18 LSH
   asg_Release.SetFocus;

end;

procedure TMainDlg.fsbt_UploadClick(Sender: TObject);
var
   sFileExt : String;
begin
   //--------------------------------------------------------------------
   // 1. 파일 오픈(.CSV, .XLS) 및 Data Loading
   //--------------------------------------------------------------------
   if opendlg_File.Execute then
   begin
      if opendlg_File.FileName = '' then
        Exit;


      sFileExt  := AnsiUpperCase(ExtractFileExt(opendlg_File.FileName));


      // Excel 파일인 경우, 별도 Procedure (LoadExcelFile) Call
      if (sFileExt = '.XLSX') or (sFileExt = '.XLS') Then
      begin
         if opendlg_File.FileName <> '' then
         begin
            // 엑셀 업로드시 주요 Timer Off (속도저하) @ 2014.07.23 LSH
            tm_DialPop.Enabled := False;
            tm_TxInit.Enabled  := False;

            LoadExcelFile(opendlg_File.FileName);
         end;
      end
      else
      // CSV 는 아래 로직 진행.
      begin
         // 필요하면 준비.
      end;
   end;
end;


//--------------------------------------------------------------------------
// 빅데이터 Lab. 엑셀 업로드 일괄등록
//
// Date   : 2015.04.02
// Author : Lee, Se-Ha
//--------------------------------------------------------------------------
procedure TMainDlg.InsBigDataList;
var
   TpInsLab  : TTpSvc;
   i, iCnt   : Integer;
   vType1,
   vRegDate  ,
   vDutyPart ,
   vDutySpec ,
   vDutyUser ,
   vContext  ,
   vClienSrc ,
   vClient	,
   vServeSrc ,
   vServer	,
   vReleasDt ,
   vTestDate ,
   vReqUser  ,
   vReqDept  ,
   vCofmUser ,
   vSrreqNo  ,
   vRemark   ,
   vReleasUser,        // 릴리즈 담당자 추가 @ 2014.06.19 LSH
   vDelDate  ,

   vCateUp   ,
   vCateDown ,
   vAttachNm ,
   vHideFile ,
   vServerIp ,
   vEditIp   : Variant;
   sType     : String;
begin

   //-------------------------------------------------------------
   // 1. 빅데이터 Raw - data 등록
   //-------------------------------------------------------------
   sType := '15';

   iCnt  := 0;


   //-------------------------------------------------------------------
   // 2. Create Variables
   //-------------------------------------------------------------------
   with asg_BigData do
   begin
      for i := 1 to RowCount - 1 do
      begin
         if (Cells[C_BD_YEAR  ,i] <> '') and
            (CopyByte(Trim(Cells[C_BD_SHOWDT  ,i]), 1, 10) >= fmed_LotFrDt.Text) and
            (CopyByte(Trim(Cells[C_BD_SHOWDT  ,i]), 1, 10) <= fmed_LotToDt.Text) then
         begin
            //------------------------------------------------------------------
            // 2-1. Append Variables
            //------------------------------------------------------------------
            AppendVariant(vType1    ,   sType                  );
            AppendVariant(vDutyPart ,   Cells[C_BD_YEAR, i]    );                   // 년차
            AppendVariant(vDutySpec ,   Cells[C_BD_SEQNO ,i]   );                   // 회차
            AppendVariant(vRegDate  ,   CopyByte(Trim(Cells[C_BD_SHOWDT  ,i]), 1, 10)); // 기준일
            AppendVariant(vCofmUser ,   Cells[C_BD_1STCNT,i]   );                   // 인원1
            AppendVariant(vContext  ,   Cells[C_BD_1STAMT,i]   );                   // 금액1
            AppendVariant(vReleasDt ,   Cells[C_BD_2NDCNT,i]   );                   // 인원2
            AppendVariant(vClient	,   Cells[C_BD_2NDAMT,i]   );                   // 금액2
            AppendVariant(vTestDate ,   Cells[C_BD_3THCNT,i]   );                   // 인원3
            AppendVariant(vServer	,   Cells[C_BD_3THAMT,i]   );                   // 금액3
            AppendVariant(vSrreqNo  ,   Cells[C_BD_4THCNT,i]   );                   // 인원4
            AppendVariant(vClienSrc ,   Cells[C_BD_4THAMT,i]   );                   // 금액4
            AppendVariant(vRemark   ,   Cells[C_BD_5THCNT,i]   );                   // 인원5
            AppendVariant(vServeSrc ,   Cells[C_BD_5THAMT,i]   );                   // 금액5
            AppendVariant(vReqDept  ,   Cells[C_BD_NO1,i]      );                   // 넘버1
            AppendVariant(vReleasUser,  Cells[C_BD_NO2,i]      );                   // 넘버2
            AppendVariant(vCateUp   ,   Cells[C_BD_NO3,i]      );                   // 넘버3
            AppendVariant(vCateDown ,   Cells[C_BD_NO4,i]      );                   // 넘버4
            AppendVariant(vAttachNm ,   Cells[C_BD_NO5,i]      );                   // 넘버5
            AppendVariant(vHideFile ,   Cells[C_BD_NO6,i]      );                   // 넘버6
            AppendVariant(vServerIp ,   Cells[C_BD_NOBONUS,i]  );                   // 넘버(보너스)
            AppendVariant(vDelDate  ,   ''                     );
            AppendVariant(vEditIp   ,   FsUserIp               );

            Inc(iCnt);
         end;
      end;
   end;

   if iCnt <= 0 then
      Exit;


   //-------------------------------------------------------------------
   // 3. Insert by TpSvc
   //-------------------------------------------------------------------
   TpInsLab := TTpSvc.Create;
   TpInsLab.Init(Self);


   Screen.Cursor := crHourGlass;


   try
      if TpInsLab.PutSvc('MD_KUMCM_M1',
                          [
                           'S_TYPE1'   , vType1
                         , 'S_STRING4' , vDutyPart
                         , 'S_STRING8' , vEditIp
                         , 'S_STRING10', vDelDate
                         , 'S_STRING11', vDutySpec
                         , 'S_STRING12', vRemark
                         , 'S_STRING13', vDutyUser
                         , 'S_STRING20', vRegDate
                         , 'S_STRING21', vReqUser
                         , 'S_STRING22', vReqDept
                         , 'S_STRING26', vCateUp
                         , 'S_STRING27', vCateDown
                         , 'S_STRING28', vContext
                         , 'S_STRING29', vAttachNm
                         , 'S_STRING30', vHideFile
                         , 'S_STRING31', vServerIp
                         , 'S_STRING37', vClient
                         , 'S_STRING38', vServer
                         , 'S_STRING44', vSrreqNo
                         , 'S_STRING46', vClienSrc
                         , 'S_STRING47', vServeSrc
                         , 'S_STRING48', vCofmUser
                         , 'S_STRING49', vReleasDt
                         , 'S_STRING50', vTestDate
                         , 'S_TYPE2'   , vReleasUser
                         ] ) then
      begin
         MessageBox(self.Handle,
                    PChar(IntToStr(iCnt) +'건의 빅데이터가 정상적으로 [업데이트] 되었습니다.'),
                    '[KUMC 다이얼로그] 빅데이터 Lab. 업데이트 알림 ',
                    MB_OK + MB_ICONINFORMATION);


         // Refresh
         SelGridInfo('BIGDATA');
      end
      else
      begin
         ShowMessageM(GetTxMsg);
      end;


   finally
      FreeAndNil(TpInsLab);
      Screen.Cursor  := crDefault;
   end;
end;


//------------------------------------------------------------------------------
// 개발관련 문서(릴리즈대장) 엑셀 업로드 일괄등록
//
// Author : Lee, Se-Ha
// Date   : 2014.
//------------------------------------------------------------------------------
procedure TMainDlg.InsReleaseList;
var
   TpInsRel  : TTpSvc;
   i, iCnt   : Integer;
   vType1,
   vRegDate  ,
   vDutyPart ,
   vDutySpec ,
   vDutyUser ,
   vContext  ,
   vClienSrc ,
   vClient	,
   vServeSrc ,
   vServer	,
   vReleasDt ,
   vTestDate ,
   vReqUser  ,
   vReqDept  ,
   vCofmUser ,
   vSrreqNo  ,
   vRemark   ,
   vReleasUser,        // 릴리즈 담당자 추가 @ 2014.06.19 LSH
   vDelDate  ,
//   vEditId   ,
   vEditIp   : Variant;
   sType     : String;
begin


   sType := '10';

   iCnt  := 0;


   //-------------------------------------------------------------------
   // 2. Create Variables
   //-------------------------------------------------------------------
   with asg_Release do
   begin
      {
      if (Cells[C_D_LOCATE,   ARow] = '') or
         (Cells[C_D_DUTYUSER, ARow] = '') or
         (Cells[C_D_DUTYPART, ARow] = '') or
         (Cells[C_D_CALLNO,   ARow] = '') then
      begin
         MessageBox(self.Handle,
                    PChar('[근무처/업무구분/담당자/연락처]는 필수입력 항목입니다.'),
                    PChar(Self.Caption + ' : 업무프로필 필수입력 항목 알림'),
                    MB_OK + MB_ICONWARNING);

         Exit;
      end;
      }



      for i := 1 to RowCount - 1 do
      begin
         if (LengthByte(Cells[C_RL_REGDATE  ,i]) > 10) and
            (CopyByte(Trim(Cells[C_RL_REGDATE  ,i]), 1, 10) >= fmed_RegFrDt.Text) and
            (CopyByte(Trim(Cells[C_RL_REGDATE  ,i]), 1, 10) <= fmed_RegToDt.Text) then
         begin
            //------------------------------------------------------------------
            // 2-1. Append Variables
            //------------------------------------------------------------------
            AppendVariant(vType1    ,   sType                  );
            AppendVariant(vRegDate  ,   CopyByte(Trim(Cells[C_RL_REGDATE  ,i]), 1, 16));
            AppendVariant(vDutyPart ,   Trim(fcb_DutySpec.Text));
            AppendVariant(vDutySpec ,   Cells[C_RL_DUTYSPEC ,i]);
            AppendVariant(vDutyUser ,   Cells[C_RL_DUTYUSER ,i]);
            AppendVariant(vContext  ,   Cells[C_RL_CONTEXT  ,i]);
            AppendVariant(vClienSrc ,   Cells[C_RL_CLIENSRC ,i]);
            AppendVariant(vClient	,   Cells[C_RL_CLIENT	,i]);
            AppendVariant(vServeSrc ,   Cells[C_RL_SERVESRC ,i]);
            AppendVariant(vServer	,   Cells[C_RL_SERVER	,i]);
            AppendVariant(vReleasDt ,   CopyByte(Trim(Cells[C_RL_RELEASDT ,i]), 1, 16));
            AppendVariant(vTestDate ,   CopyByte(Trim(Cells[C_RL_TESTDATE ,i]), 1, 16));
            AppendVariant(vReqUser  ,   Cells[C_RL_REQUSER  ,i]);
            AppendVariant(vReqDept  ,   Cells[C_RL_REQDEPT  ,i]);
            AppendVariant(vCofmUser ,   Cells[C_RL_COFMUSER ,i]);
            AppendVariant(vSrreqNo  ,   Cells[C_RL_SRREQNO  ,i]);
            AppendVariant(vReleasUser,  Cells[C_RL_RELEASUSER,i]);
            AppendVariant(vRemark   ,   Cells[C_RL_REMARK   ,i]);
            AppendVariant(vDelDate  ,   '');
            AppendVariant(vEditIp   ,   FsUserIp);

            Inc(iCnt);
         end;
      end;
   end;

   if iCnt <= 0 then
      Exit;


   //-------------------------------------------------------------------
   // 3. Insert by TpSvc
   //-------------------------------------------------------------------
   TpInsRel := TTpSvc.Create;
   TpInsRel.Init(Self);


   Screen.Cursor := crHourGlass;


   try
      if TpInsRel.PutSvc('MD_KUMCM_M1',
                          [
                           'S_TYPE1'   , vType1
                         , 'S_STRING4' , vDutyPart
                         , 'S_STRING8' , vEditIp
                         , 'S_STRING10', vDelDate
                         , 'S_STRING11', vDutySpec
                         , 'S_STRING12', vRemark
                         , 'S_STRING13', vDutyUser
                         , 'S_STRING20', vRegDate
                         , 'S_STRING21', vReqUser
                         , 'S_STRING22', vReqDept
                         , 'S_STRING28', vContext
                         , 'S_STRING37', vClient
                         , 'S_STRING38', vServer
                         , 'S_STRING44', vSrreqNo
                         , 'S_STRING46', vClienSrc
                         , 'S_STRING47', vServeSrc
                         , 'S_STRING48', vCofmUser
                         , 'S_STRING49', vReleasDt
                         , 'S_STRING50', vTestDate
                         , 'S_TYPE2'   , vReleasUser
                         ] ) then
      begin
         MessageBox(self.Handle,
                    PChar(IntToStr(iCnt) +'건의 릴리즈내역이 정상적으로 [업데이트] 되었습니다.'),
                    '[KUMC 다이얼로그] 개발문서(릴리즈대장) 업데이트 알림 ',
                    MB_OK + MB_ICONINFORMATION);


         // Refresh
         SelGridInfo('RELEASE');
      end
      else
      begin
         ShowMessageM(GetTxMsg);
      end;


   finally
      FreeAndNil(TpInsRel);
      Screen.Cursor  := crDefault;
   end;

end;



procedure TMainDlg.fsbt_InsertClick(Sender: TObject);
begin
   // Confirm Message
   if Application.MessageBox('선택한 기간의 릴리즈내역을 [재등록] 하시겠습니까?',
                             PChar('릴리즈내역 선택기간별 등록'), MB_OKCANCEL) <> IDOK then
      Exit;

   // 선택기간별 D/B 등록
   InsReleaseList;
end;



procedure TMainDlg.asg_ReleaseGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   // Horizontal Alignment
   if (ARow = 0) or ((ARow > 0) and ((ACol = C_RL_REGDATE) or
                                     (ACol = C_RL_SRREQNO) or
                                     (ACol = C_RL_REQUSER) or
                                     (ACol = C_RL_DUTYUSER) or
                                     (ACol = C_RL_RELEASDT) or
                                     (ACol = C_RL_COFMUSER) or
                                     (ACol = C_RL_TESTDATE)
                                     )
                     ) then
      HAlign := taCenter;

   // Vertical Alignment
   VAlign := vtaCenter;
end;


procedure TMainDlg.fed_ReleaseClick(Sender: TObject);
begin
   fed_Release.Text := '';
end;


procedure TMainDlg.fed_ReleaseKeyPress(Sender: TObject; var Key: Char);
var
   varResult : String;
   ResPoint  : TPoint;
   FndParam  : TFindparams;
begin
   // 릴리즈 내역 검색(D/B)
   if (fed_Release.Text <> '') and
      (not fcb_ReScanRelease.Checked)  and
      (Key = #13) then
   begin
      // Comments
      lb_RegDoc.Caption := '▶ "' + Trim(fed_Release.Text) + '"로 릴리즈 내역을 검색중 입니다....';


      // Refresh
      if IsLogonUser then
         SelGridInfo('RELEASESCAN');


      // 로그 Update
      UpdateLog('RELEASE',
                'KEYWORD',
                FsUserIp,
                'F',
                Trim(fed_Release.Text),
                '',
                varResult
                );
   end
   // 결과내 재검색 (Grid)
   else if (fed_Release.Text <> '') and
           (fcb_ReScanRelease.Checked)      and
           (Key = #13) then
   begin
      FndParam := [];

      FndParam := FndParam + {[fnMatchCase] +} [fnDirectionLeftRight];  // Character InSensitive 검색위해 fnMatchCase 주석 @ 2016.11.18 LSH

      ResPoint := asg_Release.FindFirst(Trim(fed_Release.Text), FndParam);

      if ResPoint.X >= 0 then
      begin
         asg_Release.Col   := ResPoint.X;
         asg_Release.Row   := ResPoint.Y;
      end
      else
         MessageBox(self.Handle,
                    PChar('찾으시는 항목이 화면 List에 존재하지 않습니다.'),
                    PChar('릴리즈내역 결과내 재검색 결과 알림'),
                    MB_OK + MB_ICONWARNING)
                    ;

      asg_Release.SetFocus;
   end;
end;

procedure TMainDlg.fcb_DutySpecChange(Sender: TObject);
begin
   // Comments
   lb_RegDoc.Caption := '▶ "' + Trim(fcb_DutySpec.Text) + '"OCS 릴리즈 내역을 검색중 입니다....';

   // 릴리즈 내역 조회
   if IsLogonUser then
      SelGridInfo('RELEASESCAN')
end;



procedure TMainDlg.asg_ReleaseKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  Res : TPoint;
begin

   case (Key) of

      //[F3 Key] : 다음 찾기
      Ord(VK_F3) :
                  begin
                     Res := asg_Release.FindNext;

                     if (Res.X >= 0) and
                        (Res.Y >= 0) then
                     begin
                        asg_Release.Col := Res.X;
                        asg_Release.Row := Res.Y;
                     end
                     else
                        MessageBox(self.Handle,
                                   PChar('찾으시는 항목이 화면 List에 존재하지 않습니다.'),
                                   PChar('릴리즈내역 결과내 재검색 결과 알림'),
                                   MB_OK + MB_ICONWARNING)
                                   ;
                  end;

      //[Ctrl + C] : 선택한 Cell 클립보드로 복사 @ 2016.11.18 LSH
      Ord('C') :
                  if (ssCtrl in Shift) then
                  begin
                     asg_Release.CopySelectionToClipboard;
                  end;

      //[Ctrl + A] : 해당 Grid Cell 전체내역 클립보드로 복사 @ 2016.11.18 LSH
      Ord('A') :
                  if (ssCtrl in Shift) then
                  begin
                     asg_Release.CopyToClipBoard;
                  end;

      //[Ctrl + F] : 검색 EditBox 포커싱 @ 2016.11.18 LSH
      Ord('F') :
                  if (ssCtrl in Shift) then
                  begin
                     // 결과내 재검색 True 적용
                     if not fcb_ReScanRelease.Checked then
                        fcb_ReScanRelease.Checked := True;

                     fed_Release.SetFocus;
                  end;
   end;
end;

procedure TMainDlg.asg_CongratsGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   // Horizontal Alignment
   HAlign := taCenter;

   // Vertical Alignment
   VAlign := vtaCenter;
end;

procedure TMainDlg.asg_CongratsButtonClick(Sender: TObject; ACol,
  ARow: Integer);
var
  {varResult, l_hh,l_mm, }sType, sReceiver, sSender, sRsvtime, sMessage : String;
  vType1   ,
  vSender  ,
  vReceiver,
  vContext ,
  vRegTime ,
  vUserIp  : Variant;
  TpSetSMS : TTpSvc;
begin

   //------------------------------------------------------------------
   // 1. 송, 수신자 및 메세지 설정.
   //------------------------------------------------------------------
   sType        := '13';
   sSender      := FsUserCallNo;
   sReceiver    := DelChar(asg_Congrats.Cells[2, asg_Congrats.Row], '-');
   sMessage     := '[' + FsUserNm + '님의 Happy-Birth SMS 도착알림] 생일을 진심으로 축하합니다 ^ ^/';
   sRsvTime     := '';




   //------------------------------------------------------------------
   // 2-1. Append Variables
   //------------------------------------------------------------------
   AppendVariant(vType1   ,   sType       );
   AppendVariant(vSender  ,   sSender     );
   AppendVariant(vReceiver,   sReceiver   );
   AppendVariant(vContext ,   sMessage    );
   AppendVariant(vRegTime ,   sRsvTime    );
   AppendVariant(vUserIp  ,   FsUserIp    );






   //-------------------------------------------------------------------
   // 3. Insert by TpSvc
   //-------------------------------------------------------------------
   TpSetSMS := TTpSvc.Create;
   TpSetSMS.Init(Self);


   Screen.Cursor := crHourGlass;


   try
      if TpSetSMS.PutSvc('MD_KUMCM_M1',
                         [
                           'S_TYPE1'   , vType1
                         , 'S_STRING7' , vUserIp
                         , 'S_STRING20', vRegTime
                         , 'S_STRING28', vContext
                         , 'S_STRING46', vReceiver
                         , 'S_STRING47', vSender
                         ] ) then
      begin
         // Send Messages
      end
      else
      begin
         ShowMessageM(GetTxMsg);
      end;


   finally
      FreeAndNil(TpSetSMS);

      Screen.Cursor  := crDefault;
   end;


   MessageBox(self.Handle,
              PChar('축하 SMS가 발송완료 되었습니다.'),
              '[KUMC 다이얼로그] 생일 축하 SMS발송 알림 ',
              MB_OK + MB_ICONINFORMATION);

end;

procedure TMainDlg.fsbt_DutyClick(Sender: TObject);
begin
   // Comments
   lb_RegDoc.Caption := '▶ 부서 근무/근태(휴가) 현황을 조회 및 등록하실 수 있습니다.';

   if (apn_AnalList.Visible) then
   begin
      apn_AnalList.Collaps := True;
      apn_AnalList.Visible := False;
   end;

   if (apn_Release.Visible) then
   begin
      apn_Release.Collaps := True;
      apn_Release.Visible := False;
   end;

   if (apn_ComDoc.Visible) then
   begin
      apn_ComDoc.Collaps := True;
      apn_ComDoc.Visible := False;
   end;

   if (apn_Duty.Visible) then
   begin
      apn_Duty.Top     := 39;
      apn_Duty.Left    := 6;
      apn_Duty.Visible := True;
   end
   else
   begin
      apn_Duty.Top     := 39;
      apn_Duty.Left    := 6;
      apn_Duty.Collaps := True;
      apn_Duty.Visible := True;
      apn_Duty.Collaps := False;
   end;

   fmed_DutyFrDt.Text  := FormatDateTime('yyyy-mm-dd', Date - 7);
   fmed_DutyToDt.Text  := FormatDateTime('yyyy-mm-dd', Date + 7);

   fsbt_Duty.ColorFocused     := $006CB37E;
   fsbt_ComDoc.ColorFocused   := $006CB37E;
   fsbt_Release.ColorFocused  := $006CB37E;

   // D/B 권한에 따른 근무표 생성 CheckBox 조건 @ 2015.06.12 LSH
   if GetMsgInfo('INFORM',
                 'DUTYLIST') = FsUserNm then
   begin
      fcb_DutyMake.Enabled    := True;
   end
   else
   begin
      fcb_DutyMake.Enabled    := False;
   end;

   // 근무표 확정/저장 Visible 조건 @ 2015.06.12 LSH
   fsbt_DutyInsert.Visible := fcb_DutyMake.Enabled;

   fsbt_Upload.Visible        := False;
   fsbt_Insert.Visible        := False;

   if IsLogonUser then
      SelGridInfo('DUTYSEARCH');

   tm_Master.Enabled := False;
end;

procedure TMainDlg.fmed_DutyToDtChange(Sender: TObject);
begin
   if IsLogonUser then
   begin
      if (apn_Duty.Visible) and
         (
            (
               (CopyByte(Trim(fmed_DutyToDt.Text), 1, 1)  <> '') and
               (CopyByte(Trim(fmed_DutyToDt.Text), 2, 1)  <> '') and
               (CopyByte(Trim(fmed_DutyToDt.Text), 3, 1)  <> '') and
               (CopyByte(Trim(fmed_DutyToDt.Text), 4, 1)  <> '') and
               (CopyByte(Trim(fmed_DutyToDt.Text), 6, 1)  <> '') and
               (CopyByte(Trim(fmed_DutyToDt.Text), 7, 1)  <> '') and
               (CopyByte(Trim(fmed_DutyToDt.Text), 9, 1)  <> '') and
               (CopyByte(Trim(fmed_DutyToDt.Text), 10, 1) <> '')
            )
          )
          then
      begin
         if (fcb_DutyMake.Checked = False) then
            SelGridInfo('DUTYSEARCH')
         else if (fcb_DutyMake.Checked = True) then
            SelGridInfo('DUTYMAKE');
      end;
   end;
end;

procedure TMainDlg.asg_DutyGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   // Horizontal Alignment
   if (ARow = 0) or ((ARow > 0) and ((ACol = C_DT_DUTYDATE) or
                                     (ACol = C_DT_YOIL) or
                                     (ACol = C_DT_DUTYFLAG)
                                     )
                     ) then
      HAlign := taCenter;

   // Vertical Alignment
   VAlign := vtaCenter;
end;

procedure TMainDlg.fcb_DutyListClick(Sender: TObject);
begin

   if (not apn_DutyList.Visible) then
   begin
      apn_DutyList.Top     := 26;
      apn_DutyList.Left    := 386;
   end;

   if not fcb_DutyList.Checked then
      apn_DutyList.Collaps := not fcb_DutyList.Checked;

   apn_DutyList.Visible := fcb_DutyList.Checked;

   if fcb_DutyList.Checked then
      apn_DutyList.Collaps := not fcb_DutyList.Checked;

   apn_DutyList.Caption.Text := '근무표 기준 보기';

   // 근무표 순번 가져오기
   if IsLogonUser then
      SelGridInfo('DUTYLIST');

end;

procedure TMainDlg.asg_DutyListGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   // Horizontal Alignment
   HAlign := taCenter;

   // Vertical Alignment
   VAlign := vtaCenter;
end;

procedure TMainDlg.asg_DutyListCanEditCell(Sender: TObject; ARow,
  ACol: Integer; var CanEdit: Boolean);
begin
   // 근무순번은 Edit 제한 --> 신규 개발파트 충원시, 순번도 수정해서 입력 가능하도록 하기 위해 주석처리
   // D/B 권한에 따른 CanEdit 조건 변경 @ 2015.06.12 LSH
   if GetMsgInfo('INFORM',
                 'DUTYLIST') = FsUserNm then
   begin
      // 순번(SEQNO) 제외한 나머지 수정 Enabled
      if (ACol = 0) and (asg_DutyList.Cells[ACol, ARow] = '') then
         CanEdit := True
      else if (ACol = 0) and (asg_DutyList.Cells[ACol, ARow] <> '') then
         CanEdit := False
      else
         CanEdit := ACol in [1,2,3];
   end
   else
      CanEdit := False;
end;

procedure TMainDlg.fsbt_UpdDutyClick(Sender: TObject);
begin
   if GetMsgInfo('INFORM',
                 'DUTYLIST') = FsUserNm then
   begin
      UpdateDuty('DUTYSEQ');
   end
   else
      MessageBox(self.Handle,
                 '당직근무 관리 권한이 없는 User 입니다.',
                 '당직근무 관리 권한 미대상 알림',
                 MB_OK + MB_ICONWARNING);
end;


procedure TMainDlg.asg_ReleaseDblClickCell(Sender: TObject; ARow,
  ACol: Integer);
begin
   if ACol = C_RL_SRREQNO then
   begin
      with asg_Release do
      begin
         GetLinkList('ANALYSIS',
                     TokenStr(Trim(Cells[C_RL_SRREQNO, Row]) + #13#10, #13#10, 1),  // S/R 번호 2건이상인 경우 제일 최종 S/R 번호만 잘라서 상세조회 @ 2016.11.18 LSH
                     '',
                     '',
                     FsUserIp);
      end;
   end;
end;

procedure TMainDlg.asg_WeeklyRptDblClickCell(Sender: TObject; ARow,
  ACol: Integer);
begin
   if ACol = C_WK_CONTEXT then
   begin
      with asg_WeeklyRpt do
      begin
         if (Cells[C_WK_CONTEXT, Row] <> '') and
            (
               (PosByte('S/R', Cells[C_WK_GUBUN, Row]) > 0) or
               (ftc_Dialog.ActiveTab = AT_DIALBOOK)     or  // [다이얼 Book] 담당자 프로필 : S/R 상세보기 모드
               (
                  (PosByte('S/R',    Cells[C_WK_GUBUN, Row]) = 0) and
                  (ftc_Dialog.ActiveTab = AT_ANALYSIS)        and
                  (asg_Analysis.Visible)
               ) or  // [요청자 프로필] : S/R요청현황 상세보기 모드
               (
                  (PosByte('릴리즈', Cells[C_WK_GUBUN, Row]) > 0) and
                  (PosByte('<', Cells[C_WK_CONTEXT, Row]) > 0)    and
                  (PosByte('>', Cells[C_WK_CONTEXT, Row]) > 0)
               )    // [주간리포트] > [릴리즈]내역도 S/R 상세 연동 @ 2016.12.08 LSH
            ) then
         begin
            GetLinkList('ANALYSIS',
                        CopyByte(Cells[C_WK_CONTEXT, Row], PosByte('<', Cells[C_WK_CONTEXT, Row]) + 1, LengthByte(Cells[C_WK_CONTEXT, Row]) - PosByte('<', Cells[C_WK_CONTEXT, Row]) - 1),
                        '',
                        '',
                        FsUserIp);

            // 검색조건 Clear
            fed_Analysis.Text := '';
         end
         else if (ftc_Dialog.ActiveTab = AT_DIALOG)   or
                 (ftc_Dialog.ActiveTab = AT_DIALBOOK) or
                 ((ftc_Dialog.ActiveTab = AT_ANALYSIS) and (asg_Analysis.Visible)) then
         begin
            apn_Weekly.Collaps := True;
            apn_Weekly.Visible := False;

            // 검색조건 Clear
            fed_Analysis.Text  := '';

            // 결과내 검색 visible
            fcb_ReScan.Visible := True;
         end;
      end;
   end
   else if (ftc_Dialog.ActiveTab = AT_DIALOG)   or
           (ftc_Dialog.ActiveTab = AT_DIALBOOK) or
           ((ftc_Dialog.ActiveTab = AT_ANALYSIS) and (asg_Analysis.Visible)) then
   begin
      apn_Weekly.Collaps := True;
      apn_Weekly.Visible := False;

      // 검색조건 Clear
      fed_Analysis.Text := '';

      // 결과내 검색 visible
      fcb_ReScan.Visible := True;
   end;

end;

procedure TMainDlg.fmed_DutyFrDtChange(Sender: TObject);
begin
   if IsLogonUser then
   begin
      if (apn_Duty.Visible) and
         (
            (
               (CopyByte(Trim(fmed_DutyFrDt.Text), 1, 1)  <> '') and
               (CopyByte(Trim(fmed_DutyFrDt.Text), 2, 1)  <> '') and
               (CopyByte(Trim(fmed_DutyFrDt.Text), 3, 1)  <> '') and
               (CopyByte(Trim(fmed_DutyFrDt.Text), 4, 1)  <> '') and
               (CopyByte(Trim(fmed_DutyFrDt.Text), 6, 1)  <> '') and
               (CopyByte(Trim(fmed_DutyFrDt.Text), 7, 1)  <> '') and
               (CopyByte(Trim(fmed_DutyFrDt.Text), 9, 1)  <> '') and
               (CopyByte(Trim(fmed_DutyFrDt.Text), 10, 1) <> '')
            )
          )
          then
      begin
         if (fcb_DutyMake.Checked = False) then
            SelGridInfo('DUTYSEARCH')
         else if (fcb_DutyMake.Checked = True) then
            SelGridInfo('DUTYMAKE');
      end;
   end;
end;

procedure TMainDlg.fed_DutyClick(Sender: TObject);
begin
   fed_Duty.Text := '';
end;

procedure TMainDlg.fed_DutyKeyPress(Sender: TObject; var Key: Char);
var
   varResult : String;
begin
   showmessage('근무 및 휴가 키워드검색 서비스 준비중.');

   // 로그 Update
   UpdateLog('DUTY',
             'KEYWORD',
             FsUserIp,
             'F',
             Trim(fed_Duty.Text),
             '',
             varResult
             );
end;

procedure TMainDlg.asg_BoardDblClickCell(Sender: TObject; ARow,
  ACol: Integer);
var
   sServerIp, sRemoteFile, sLocalFile, sFtpUserID, sFtpPasswd, sFtpHostName, sFtpDir{, varUpResult} : String;
begin
   if not IsLogonUser then
   begin
      MessageBox(self.Handle,
                 PChar('[비상연락망] - [담당자정보]에 담당자 IP 등록후 사용하실 수 있습니다.'),
                 PChar(Self.Caption + ' : 등록되지 않은 User 제한 알림'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;

   if ((ACol = C_B_REGDATE) or
       (ACol = C_B_REGUSER))  and
      (asg_Board.Cells[C_B_BOARDSEQ, ARow] = '┗') then
   begin

      // 첨부 이미지 FTP 파일내역이 존재하면, 이미지 Download
      if (asg_Board.Cells[C_B_HIDEFILE, ARow - 1] <> '') then
      begin
         // 파일 업/다운로드를 위한 정보 조회
         if not GetBinUploadInfo(sServerIp, sFtpUserID, sFtpPasswd, sFtpHostName, sFtpDir) then
         begin
            MessageDlg('파일 저장을 위한 담당자 정보 조회중, 오류가 발생했습니다.', Dialogs.mtError, [Dialogs.mbOK], 0);
            exit;
         end;


         // 실제저장된 서버 IP
         sServerIp := C_KDIAL_FTP_IP;


         // 실제 서버에 저장되어 있는 파일명 지정
         if PosByte('/ftpspool/KDIALFILE/',asg_Board.Cells[C_B_HIDEFILE, ARow - 1]) > 0 then
            sRemoteFile := asg_Board.Cells[C_B_HIDEFILE, ARow]
         else
            sRemoteFile := '/ftpspool/KDIALFILE/' + asg_Board.Cells[C_B_HIDEFILE, ARow - 1];

         // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
         sLocalFile  := G_HOMEDIR + 'TEMP\SPOOL\' + asg_Board.Cells[C_B_ATTACH, ARow - 1];



         if (GetBINFTP(sServerIp, sFtpUserID, sFtpPasswd, sRemoteFile, sLocalFile, False)) then
         begin
            //	정상적인 FTP 다운로드
         end;


         try
            if (PosByte('.exe', sLocalFile) > 0) or
               (PosByte('.zip', sLocalFile) > 0) or
               (PosByte('.rar', sLocalFile) > 0) then
            begin
               MessageBox(self.Handle,
                          PChar('첨부파일(' + sLocalFile + ') 다운로드가 완료되었습니다.' + #13#10 + #13#10 +
                                '※ 임시 다운로드 폴더 --> C:\KUMC(_DEV)\TEMP\SPOOL\'),
                          '첨부파일 다운로드 실행 완료',
                          MB_OK + MB_ICONINFORMATION);
            end
            else
            begin
               ShellExecute(HANDLE, 'open',
                            PCHAR(asg_Board.Cells[C_B_ATTACH, ARow - 1]),
                            PCHAR(''),
                            // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
                            PCHAR(G_HOMEDIR + 'TEMP\SPOOL\'),
                            SW_SHOWNORMAL);
            end;

         except
            MessageBox(self.Handle,
                       PChar('해당 첨부 파일 실행중 오류가 발생하였습니다.' + #13#10 + #13#10 +
                             '종료 후 다시 실행해 주시기 바랍니다.'),
                       '첨부파일 다운로드 실행 오류',
                       MB_OK + MB_ICONERROR);

            Exit;
         end;


         {
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
         asg_Board.CreatePicture(C_B_REGDATE,
                                 asg_Board.Row + 1,
                                 False,
                                 ShrinkWithAspectRatio,
                                 0,
                                 haLeft,
                                 vaTop).LoadFromFile(sLocalFile)
         else
            asg_Board.Cells[C_B_REGDATE, asg_Board.Row + 1] := asg_Board.Cells[C_B_ATTACH, asg_Board.Row];
         }
      end;
   end;
end;

procedure TMainDlg.asg_UserProfileButtonClick(Sender: TObject; ACol,
  ARow: Integer);
var
   sFileName, sHideFile, sServerIp : String;
begin
   if ARow = R_PR_DUTYSCH then
      // 업무 리포트(스케쥴) 조회
      fsbt_WeeklyClick(Sender)
   else if ARow = R_PR_PHOTO then
   begin
      // 유효성 체크
      if PosByte(asg_UserProfile.Cells[1, R_PR_USERNM] , asg_NetWork.Cells[C_NW_USERNM, asg_NetWork.Row]) = 0 then
      begin
         MessageBox(self.Handle,
                    PChar('현재 프로필 조회중인 담당자정보가 선택된 정보와 일치하지 않습니다.' + #13#10 + '담당자 이름을 재클릭후, 이미지 등록 진행해주세요.'),
                    '비상연락망 담당자 프로필 이미지 등록전 확인',
                    MB_OK + MB_ICONERROR);
         Exit;
      end;


      // 이미지 등록 버튼 삭제
      asg_UserProfile.RemoveButton(1, R_PR_PHOTO);


      if od_File.Execute then
      begin
         // 첨부파일 Upload
         if (Trim(ExtractFileName(od_File.FileName)) <> '') then
         begin

            sFileName := ExtractFileName(Trim(ExtractFileName(od_File.FileName)));
            sHideFile := 'KDIALAPPEND' + 'IMG' + FormatDateTime('YYYYMMDDHHNNSS', Now) + '.KDIALFILE';


            if Not FileUpLoad(sHideFile, sFileName, sServerIp) then
            begin
               Showmessage('첨부파일 UpLoad 중 에러가 발생했습니다.' + #13#10 + #13#10 +
                           '다시한번 시도해 주시기 바랍니다.');
               Exit;
            end
            else
            begin
            {
               MessageBox(self.Handle,
                          PChar('이미지가 정상적으로 등록되었습니다.' + #13#10 + '담당자 프로필 재조회시 등록된 이미지를 보실 수 있습니다.'),
                          '비상연락망 담당자 프로필 이미지 등록',
                          MB_OK + MB_ICONINFORMATION);
             }

               // 담당자 마스터에 이미지 정보 Update
               UpdateImage('PROFILE',
                           sFileName,
                           sHideFile,
                           asg_Network.Row);

            end;
         end;
      end;
   end;
end;



//------------------------------------------------------------------------------
// Panel 상태 변경하기
//    - 2014.01.29 LSH
//------------------------------------------------------------------------------
procedure TMainDlg.SetPanelStatus(in_Gubun,
                                  in_Status : String);
//var
//   k : Integer;
begin

   if (in_Status = 'OFF') then
   begin
      if (in_Gubun = 'ALL') then
      begin
         apn_Weekly.Visible      := False;
         apn_AnalList.Visible    := False;
         apn_Profile.Visible     := False;
         apn_UserProfile.Visible := False;
         apn_LikeList.Visible    := False;
         apn_Menu.Visible        := False;
         apn_LinkList.Visible    := False;
         apn_DialMap.Visible     := False;
         apn_AddMyDial.Visible   := False;
         apn_Congrats.Visible    := False;
         apn_DocSpec.Visible     := False;
         apn_MultiIns.Visible    := False;

         // 동적생성 컴포넌트 강제 Free @ 2014.07.31 LSH
         with asg_MultiIns do
         begin
            // [업무일지 작성내역] 셀의 TMemo는 삭제.
            if Objects[1,2] <> nil Then
            begin
               Objects[1,2].Free;
               Objects[1,2] := Nil;
            end;
         end;
      end
      else if (in_Gubun = 'WORKRPT') then
      begin
         apn_MultiIns.Collaps := True;
         apn_MultiIns.Visible := False;

         // 동적생성 컴포넌트 강제 Free @ 2014.07.31 LSH
         with asg_MultiIns do
         begin
            // [업무일지 작성내역] 셀의 TMemo는 삭제.
            if Objects[1,2] <> nil Then
            begin
               Objects[1,2].Free;
               Objects[1,2] := Nil;
            end;
         end;
      end
      else if (in_Gubun = 'WORKCONN') then
      begin
         apn_MultiIns.Collaps := True;
         apn_MultiIns.Visible := False;

         // 동적생성 컴포넌트 강제 Free @ 2014.07.31 LSH
         with asg_MultiIns do
         begin
            // [작성내역] 셀의 TMemo는 삭제.
            if Objects[1,2] <> nil Then
            begin
               Objects[1,2].Free;
               Objects[1,2] := Nil;
            end;
         end;
      end;
   end
   else if (in_Status = 'ON') then
   begin
      if (in_Gubun = 'WORKRPT') then
      begin
         with asg_MultiIns do
         begin
            FixedColor  := $0000B1C9;
            HintColor   := $0087D5DB;

            // Column Header Set
            Cells[0, 0] := '근무처(소속)';
            Cells[0, 1] := '작성일자';
            Cells[0, 2] := '업무내용';
            Cells[0, 8] := '';

            // 근무처 Set
            if PosByte('안암도메인', FsUserIp) > 0 then
            begin
               if PosByte('기획', FsUserPart) > 0 then
                  Cells[1, 0] := 'MIS'
               else
                  Cells[1, 0] := 'AA';
            end
            else if PosByte('구로도메인', FsUserIp) > 0 then
               Cells[1, 0] := 'GR'
            else if PosByte('안산도메인', FsUserIp) > 0 then
               Cells[1, 0] := 'AS';


            // 작성일자, 업무내용 Set
            // 현재 선택한 일자 정보 Set @ 2017.10.26 LSH
            Cells[1, 1] := CopyByte(asg_WorkRpt.Cells[C_WR_REGDATE, asg_WorkRpt.Row], 1, 10); //FormatDateTime('yyyy-mm-dd', Date);


            // 오늘일자 업무일지 기 작성내역 있으면, 업무일지 작성 Memo 컬럼으로 복사 @ 2014.07.31 LSH
            //if (PosByte(FormatDateTime('yyyy-mm-dd', Date), asg_WorkRpt.Cells[C_WR_REGDATE, asg_WorkRpt.Row]) > 0) then
            //begin
               // 현재 선택한 일자의 업무일지 내역 Set @ 2017.10.26 LSH
               Cells[1, 2] := asg_WorkRpt.Cells[C_WR_CONTEXT, asg_WorkRpt.Row];
            //end;

            AddComment(0,
                       1,
                       '업무일지 작성 Panel (또는 Fixed-Column) 더블클릭시, 작성화면을 닫으실 수 있습니다.');

            AddComment(0,
                       2,
                       '업무일지 작성후 아래 [등록]버튼을 누르면 저장됩니다.');

            // Merge Cells
            MergeCells(0, 2, 1, 6);
            MergeCells(1, 2, 1, 6);


            AddButton(1,
                      8,
                      ColWidths[1]-5,
                      20,
                      '업무일지 등록',
                      haBeforeText,
                      vaCenter);

         end;

         // Panel Display
         apn_MultiIns.Caption.Text  := '업무일지 작성';
         apn_MultiIns.Top           := 205;
         apn_MultiIns.Left          := 111;
         apn_MultiIns.Collaps       := True;
         apn_MultiIns.Visible       := True;
         apn_MultiIns.Collaps       := False;

      end
      else if (in_Gubun = 'WORKCONN') then
      begin
         with asg_MultiIns do
         begin
            FixedColor  := $007091C5;
            HintColor   := $0070AFC5;

            // Column Header Set
            Cells[0, 0] := '근무처(소속)';
            Cells[0, 1] := '등록일자';
            Cells[0, 2] := '주요업무';
            Cells[0, 8] := '';

            // 근무처 Set
            if PosByte('안암도메인', FsUserIp) > 0 then
            begin
               if FsUserNm = FsMngrNm then
                  Cells[1, 0] := '의료원'
               else
                  Cells[1, 0] := '안암';
            end
            else if PosByte('구로도메인', FsUserIp) > 0 then
               Cells[1, 0] := '구로'
            else if PosByte('안산도메인', FsUserIp) > 0 then
               Cells[1, 0] := '안산';


            // 등록일자, 주요업무 Set
            Cells[1, 1] := FormatDateTime('yyyy-mm-dd', Date);


            // 오늘일자 주요업무 작성내역 있으면, 주요업무 작성 Memo 컬럼으로 복사 @ 2014.07.31 LSH
            if (PosByte(FormatDateTime('yyyy-mm-dd', Date), asg_WorkConn.Cells[C_WC_REGDATE, asg_WorkConn.Row]) > 0) then
            begin
               Cells[1, 2] := asg_WorkConn.Cells[C_WC_CONTEXT, asg_WorkConn.Row];
            end;

            AddComment(0,
                       1,
                       '주요업무공유 작성 Panel 더블클릭시, 작성화면을 종료합니다.');

            AddComment(0,
                       2,
                       '주요업무공유 작성후 아래 등록버튼을 누르면 저장됩니다.');

            // Merge Cells
            MergeCells(0, 2, 1, 6);
            MergeCells(1, 2, 1, 6);


            AddButton(1,
                      8,
                      ColWidths[1]-5,
                      20,
                      '주요업무공유 등록',
                      haBeforeText,
                      vaCenter);

         end;

         // Panel Display
         apn_MultiIns.Caption.Text  := '주요업무공유 작성';
         apn_MultiIns.Top           := 205;
         apn_MultiIns.Left          := 111;
         apn_MultiIns.Collaps       := True;
         apn_MultiIns.Visible       := True;
         apn_MultiIns.Collaps       := False;

      end
      else if (in_Gubun = 'WORKCONNMOD') then
      begin
         with asg_MultiIns do
         begin
            FixedColor  := $007091C5;
            HintColor   := $0070AFC5;

            // Column Header Set
            Cells[0, 0] := '근무처(소속)';
            Cells[0, 1] := '등록일자';
            Cells[0, 2] := '주요업무';
            Cells[0, 8] := '';

            // 근무처 Set
            if PosByte('안암도메인', FsUserIp) > 0 then
            begin
               if FsUserNm = FsMngrNm then
                  Cells[1, 0] := '의료원'
               else
                  Cells[1, 0] := '안암';
            end
            else if PosByte('구로도메인', FsUserIp) > 0 then
               Cells[1, 0] := '구로'
            else if PosByte('안산도메인', FsUserIp) > 0 then
               Cells[1, 0] := '안산';


            // 현재 수정하려는 일자 작성내역 조회
            Cells[1, 1] := asg_WorkConn.Cells[C_WC_REGDATE, asg_WorkConn.Row];
            Cells[1, 2] := asg_WorkConn.Cells[C_WC_CONTEXT, asg_WorkConn.Row];


            AddComment(0,
                       1,
                       '주요업무공유 작성 Panel 더블클릭시, 작성화면을 종료합니다.');

            AddComment(0,
                       2,
                       '주요업무공유 작성후 아래 등록버튼을 누르면 저장됩니다.');

            // Merge Cells
            MergeCells(0, 2, 1, 6);
            MergeCells(1, 2, 1, 6);


            AddButton(1,
                      8,
                      ColWidths[1]-5,
                      20,
                      '주요업무공유 등록',
                      haBeforeText,
                      vaCenter);

         end;

         // Panel Display
         apn_MultiIns.Caption.Text  := '주요업무공유 수정';
         apn_MultiIns.Top           := 205;
         apn_MultiIns.Left          := 111;
         apn_MultiIns.Collaps       := True;
         apn_MultiIns.Visible       := True;
         apn_MultiIns.Collaps       := False;

      end;
   end;
end;

procedure TMainDlg.fsbt_DBMasterClick(Sender: TObject);
begin
   // 업무일지 작성(TMemo 동적 컴포넌트) Panel Close @ 2014.07.31 LSH
   SetPanelStatus('WORKRPT', 'OFF');

   if apn_Weekly.Visible then
   begin
      apn_Weekly.Collaps := True;
      apn_Weekly.Visible := False;
   end;

   if apn_AnalList.Visible then
   begin
      apn_AnalList.Collaps := True;
      apn_AnalList.Visible := False;
   end;

   fcb_ReScan.Top          := 31;
   fcb_ReScan.Left         := 370;
   fcb_ReScan.ColorFocused := $00D7BF8B;
   fcb_ReScan.Checked      := False;

   asg_Analysis.Top     := 999;
   asg_Analysis.Left    := 999;
   asg_Analysis.Visible := False;

   asg_WorkRpt.Top      := 999;
   asg_WorkRpt.Left     := 999;
   asg_WorkRpt.Visible  := False;

   fed_Analysis.Top        := 30;
   fed_Analysis.Left       := 265;
   fed_Analysis.Width      := 100;
   fed_Analysis.Visible    := True;
   fed_Analysis.ColorFlat  := $00D7BF8B;
   fed_Analysis.Text       := '';
   fed_Analysis.Hint       := 'D/B에서 조회하려는 키워드(예: Table명, Column명, proc./func./trigger 등)를 입력후  Enter를 눌러보세요.';


   fsbt_Analysis.ColorFocused := $00D7BF8B;
   fsbt_WorkRpt.ColorFocused  := $00D7BF8B;
   fsbt_Weekly.ColorFocused   := $00D7BF8B;
   fsbt_DBMaster.ColorFocused := $00D7BF8B;

   fed_DateTitle.ColorFlat := $00D7BF8B;
   fmed_AnalFrom.ColorFlat := $00D7BF8B;
   fmed_AnalTo.ColorFlat   := $00D7BF8B;

   fed_DateTitle.Text      := ' 검색';
   label8.Left             := 999;
   fmed_AnalFrom.Left      := 999;
   fmed_AnalTo.Left        := 999;

   asg_Analysis.Top      := 999;
   asg_Analysis.Left     := 999;
   asg_Analysis.Visible  := False;


   lb_Analysis.Caption := '▶ D/B 분석의 모든것, D/B 마스터!';

   fmed_AnalFrom.Text   := FormatDateTime('yyyy-mm-dd', DATE - 1);
   fmed_AnalTo.Text     := FormatDateTime('yyyy-mm-dd', DATE);

   fcb_WorkArea.Top     := 30;
   fcb_WorkArea.Left    := 57;
   fcb_WorkArea.Visible := True;
   fcb_WorkArea.Color   := $00D7BF8B;

   fcb_WorkGubun.Top     := 30;
   fcb_WorkGubun.Left    := 185;
   fcb_WorkGubun.Visible := True;
   fcb_WorkGubun.Color   := $00D7BF8B;

   fcb_DutyPart.Top      := 30;
   fcb_DutyPart.Left     := 112;
   fcb_DutyPart.Visible  := True;
   fcb_DutyPart.Color    := $00D7BF8B;

   fsbt_Print.Visible    := False;

   apn_DBMaster.Top     := 63;
   apn_DBMaster.Left    := 7;
   apn_DBMaster.Collaps := True;
   apn_DBMaster.Visible := True;

   asg_DBMaster.ClearRows(1, asg_DBMaster.RowCount);
   asg_DBMaster.RowCount := 1;

   asg_DBDetail.ClearRows(1, asg_DBDetail.RowCount);
   asg_DBDetail.RowCount := 1;

   apn_DBMaster.Collaps := False;


   // [병원구분] 진료공통 Code 조회
   GetComboBoxList('KDIAL',
                   'LOCATE',
                   fcb_WorkArea);

   // [검색구분] 진료공통 Code 조회
   GetComboBoxList('KDIAL',
                   'DBLIST',
                   fcb_WorkGubun);

   // [파트구분] 진료공통 Code 조회
   GetComboBoxList('KDIAL',
                   'DUTYPART',
                   fcb_DutyPart);

   // 접속자 IP 식별후, 병원구분 Assign
   if PosByte('안암도메인', FsUserIp) > 0 then
      fcb_WorkArea.Text := '안암'
   else if PosByte('구로도메인', FsUserIp) > 0 then
      fcb_WorkArea.Text := '구로'
   else if PosByte('안산도메인', FsUserIp) > 0 then
      fcb_WorkArea.Text := '안산';

   // [검색구분] 기본 조건값
   fcb_WorkGubun.Text := fcb_WorkGubun.Items.Strings[0];

   // [파트구분] 소속파트별 세팅, 기본 조건은 "진료"
   if PosByte('개발', FsUserPart) > 0 then
   begin
      if PosByte('P/L', FsUserSpec) > 0 then
         fcb_DutyPart.Text := CopyByte(FsUserSpec, 1, PosByte('(', FsUserSpec) - 1)
      else
         fcb_DutyPart.Text := FsUserSpec;
   end
   else
      fcb_DutyPart.Text := '진료';


   // D/B 마스터 카테고리 조회
   if IsLogonUser then
   begin
      SelGridInfo('DBCATEGORY');

      if not fcb_ReScan.Checked then
         fcb_ReScan.Checked := True;

      // 검색 EditBox 포커싱 @ 2016.11.18 LSH
      fed_Analysis.SetFocus;
   end;


end;



//---------------------------------------------------------------------------
// [조회] FlatComboBox 공통 Code 조회
//---------------------------------------------------------------------------
procedure TMainDlg.GetComboBoxList(in_Type1,
                                   in_Type2 : String;
                                   NmCombo  : TFlatComboBox);
var
   i : integer;
   TpGetComboList : TTpSvc;
begin


   Screen.Cursor := crHourglass;




   //-----------------------------------------------------------------
   // 1. Load Data by TpSvc
   //-----------------------------------------------------------------
   TpGetComboList := TTpSvc.Create;
   TpGetComboList.Init(Self);


   Screen.Cursor := crHourglass;


   try
      TpGetComboList.CountField  := 'S_STRING1';
      TpGetComboList.ShowMsgFlag := False;

      if TpGetComboList.GetSvc('MD_MCOMC_L1',
                              ['S_TYPE1'  , 'J'
                             , 'S_CODE1'  , in_Type1
                             , 'S_CODE2'  , in_Type2
                             , 'S_CODE3'  , ''
                              ],
                              [
                               'L_CNT1'    , 'iCitem06'
                             , 'L_CNT2'    , 'iCitem07'
                             , 'L_CNT3'    , 'iCitem08'
                             , 'L_CNT4'    , 'iCitem09'
                             , 'L_CNT5'    , 'iCitem10'
                             , 'L_SEQNO1'  , 'iDispseq'
                             , 'S_STRING1' , 'sComcd1'
                             , 'S_STRING2' , 'sComcd2'
                             , 'S_STRING3' , 'sComcd3'
                             , 'S_STRING4' , 'sComcdnm1'
                             , 'S_STRING5' , 'sComcdnm2'
                             , 'S_STRING6' , 'sComcdnm3'
                             , 'S_STRING7' , 'sCitem01'
                             , 'S_STRING8' , 'sCitem02'
                             , 'S_STRING9' , 'sCitem03'
                             , 'S_STRING10', 'sCitem04'
                             , 'S_STRING11', 'sCitem05'
                             , 'S_STRING12', 'sCitem11'
                             , 'S_STRING13', 'sCitem12'
                             , 'S_STRING14', 'sDeldate'
                             ]) then

         if TpGetComboList.RowCount < 0 then
         begin
            ShowMessage(GetTxMsg);
            Exit;
         end
         else if TpGetComboList.RowCount = 0 then
         begin
            Exit;
         end;


      //----------------------------------------------------
      // 2. ComboBox List 등록
      //----------------------------------------------------
      with NmCombo do
      begin
         Items.Clear;

         for i := 0 to TpGetComboList.RowCount - 1 do
            Items.Add(TpGetComboList.GetOutputDataS('sComcdnm3', i));
      end;



   finally
      FreeAndNil(TpGetComboList);
      screen.Cursor  :=  crDefault;
   end;
end;



procedure TMainDlg.asg_DBMasterClick(Sender: TObject);
begin
   // D/B 마스터 카테고리 상세 조회
   if IsLogonUser then
      SelGridInfo('DBCATEDET');
end;

procedure TMainDlg.fcb_DutyPartChange(Sender: TObject);
begin
   if IsLogonUser then
      if (apn_DBMaster.Visible) then
         SelGridInfo('DBCATEGORY');
end;



//---------------------------------------------------------------------------
// Grid 컬럼 정보 변경
//---------------------------------------------------------------------------
procedure TMainDlg.SetGridCol(in_Flag : String);
begin
   if (in_Flag = 'TABLE') then
   begin
      with asg_DBMaster do
      begin
         ColWidths[C_DBM_FSTCD]     := 60;
         ColWidths[C_DBM_FSTCDNM]   := 180;
         ColWidths[C_DBM_WORKINFO]  := 52;

         Cells[C_DBM_FSTCD,   0]    := 'Table';
         Cells[C_DBM_FSTCDNM, 0]    := 'Table명';
         Cells[C_DBM_WORKINFO,0]    := '생성일';

         Width := 313;
      end;

      with asg_DBDetail do
      begin
         ColWidths[C_DBD_SECCD]     := 80;
         ColWidths[C_DBD_SECCDNM]   := 150;
         ColWidths[C_DBD_3RDCD]     := 30;
         ColWidths[C_DBD_3RDCDNM]   := 30;
         ColWidths[C_DBD_WORKINFO]  := 30;

         RowHeights[1]    := 22;

         Cells[C_DBD_SECCD,   0]    := 'Column';
         Cells[C_DBD_SECCDNM, 0]    := 'Column명';
         Cells[C_DBD_3RDCD,   0]    := 'Type';
         Cells[C_DBD_3RDCDNM, 0]    := 'Size';
         Cells[C_DBD_WORKINFO,0]    := 'PK';

         Left  := 312;
         Width := 332;
      end;
   end
   else if (in_Flag = 'PROCEDURE') then
   begin
      with asg_DBMaster do
      begin
         ColWidths[C_DBM_FSTCD]     := 130;
         ColWidths[C_DBM_FSTCDNM]   := 0;
         ColWidths[C_DBM_WORKINFO]  := 50;

         Cells[C_DBM_FSTCD,   0]    := 'Proc.';
         Cells[C_DBM_FSTCDNM, 0]    := '';
         Cells[C_DBM_WORKINFO,0]    := '생성일';

         Width := 200;
      end;

      with asg_DBDetail do
      begin
         ColWidths[C_DBD_SECCD]     := 25;
         ColWidths[C_DBD_SECCDNM]   := 395;
         ColWidths[C_DBD_3RDCD]     := 0;
         ColWidths[C_DBD_3RDCDNM]   := 0;
         ColWidths[C_DBD_WORKINFO]  := 0;

         RowHeights[1]    := 22;

         Cells[C_DBD_SECCD,   0]    := 'Line';
         Cells[C_DBD_SECCDNM, 0]    := 'Source';
         Cells[C_DBD_3RDCD,   0]    := '';
         Cells[C_DBD_3RDCDNM, 0]    := '';
         Cells[C_DBD_WORKINFO,0]    := '';

         Left  := 202;
         Width := 440;
      end;
   end
   else if (in_Flag = 'FUNCTION') then
   begin
      with asg_DBMaster do
      begin
         ColWidths[C_DBM_FSTCD]     := 130;
         ColWidths[C_DBM_FSTCDNM]   := 0;
         ColWidths[C_DBM_WORKINFO]  := 50;

         Cells[C_DBM_FSTCD,   0]    := 'Func.';
         Cells[C_DBM_FSTCDNM, 0]    := '';
         Cells[C_DBM_WORKINFO,0]    := '생성일';

         Width := 200;
      end;

      with asg_DBDetail do
      begin
         ColWidths[C_DBD_SECCD]     := 25;
         ColWidths[C_DBD_SECCDNM]   := 395;
         ColWidths[C_DBD_3RDCD]     := 0;
         ColWidths[C_DBD_3RDCDNM]   := 0;
         ColWidths[C_DBD_WORKINFO]  := 0;

         RowHeights[1]    := 22;

         Cells[C_DBD_SECCD,   0]   := 'Line';
         Cells[C_DBD_SECCDNM, 0]   := 'Source';
         Cells[C_DBD_3RDCD,   0]   := '';
         Cells[C_DBD_3RDCDNM, 0]   := '';
         Cells[C_DBD_WORKINFO,0]   := '';

         Left  := 202;
         Width := 440;

      end;
   end
   else if (in_Flag = 'TRIGGER') then
   begin
      with asg_DBMaster do
      begin
         ColWidths[C_DBM_FSTCD]     := 130;
         ColWidths[C_DBM_FSTCDNM]   := 0;
         ColWidths[C_DBM_WORKINFO]  := 50;

         Cells[C_DBM_FSTCD,   0]    := 'Trig.';
         Cells[C_DBM_FSTCDNM, 0]    := '';
         Cells[C_DBM_WORKINFO,0]    := '생성일';

         Width := 200;
      end;

      with asg_DBDetail do
      begin
         ColWidths[C_DBD_SECCD]     := 0;
         ColWidths[C_DBD_SECCDNM]   := 80;
         ColWidths[C_DBD_3RDCD]     := 0;
         ColWidths[C_DBD_3RDCDNM]   := 340;
         ColWidths[C_DBD_WORKINFO]  := 0;

         RowHeights[1]    := 450;

         Cells[C_DBD_SECCD,   0]    := '';
         Cells[C_DBD_SECCDNM, 0]    := 'when Clause';
         Cells[C_DBD_3RDCD,   0]    := '';
         Cells[C_DBD_3RDCDNM, 0]    := 'Source';
         Cells[C_DBD_WORKINFO,0]    := '';

         Left  := 202;
         Width := 440;

      end;
   end
   else if (in_Flag = 'PACKAGE') then
   begin
      with asg_DBMaster do
      begin
         ColWidths[C_DBM_FSTCD]     := 55;
         ColWidths[C_DBM_FSTCDNM]   := 75;
         ColWidths[C_DBM_WORKINFO]  := 50;

         Cells[C_DBM_FSTCD,   0]    := 'Pkg.';
         Cells[C_DBM_FSTCDNM, 0]    := 'Pkg.Body';
         Cells[C_DBM_WORKINFO,0]    := '생성일';

         Width := 200;
      end;

      with asg_DBDetail do
      begin
         ColWidths[C_DBD_SECCD]     := 25;
         ColWidths[C_DBD_SECCDNM]   := 395;
         ColWidths[C_DBD_3RDCD]     := 0;
         ColWidths[C_DBD_3RDCDNM]   := 0;
         ColWidths[C_DBD_WORKINFO]  := 0;


         RowHeights[1]    := 22;

         Cells[C_DBD_SECCD,   0]    := 'Line';
         Cells[C_DBD_SECCDNM, 0]    := 'Source';
         Cells[C_DBD_3RDCD,   0]    := '';
         Cells[C_DBD_3RDCDNM, 0]    := '';
         Cells[C_DBD_WORKINFO,0]    := '';

         Left  := 202;
         Width := 440;

      end;
   end
   else if (in_Flag = '공통코드') then
   begin
      with asg_DBMaster do
      begin
         ColWidths[C_DBM_FSTCD]     := 55;
         ColWidths[C_DBM_FSTCDNM]   := 190;
         ColWidths[C_DBM_WORKINFO]  := 47;

         Cells[C_DBM_FSTCD,   0]    := '대분류';
         Cells[C_DBM_FSTCDNM, 0]    := '대분류명';
         Cells[C_DBM_WORKINFO,0]    := '작업자';

         Width := 313;

      end;

      with asg_DBDetail do
      begin
         ColWidths[C_DBD_SECCD]     := 53;
         ColWidths[C_DBD_SECCDNM]   := 85;
         ColWidths[C_DBD_3RDCD]     := 60;
         ColWidths[C_DBD_3RDCDNM]   := 65;
         ColWidths[C_DBD_WORKINFO]  := 50;

         RowHeights[1]    := 22;

         Cells[C_DBD_SECCD,   0]    := '중분류';
         Cells[C_DBD_SECCDNM, 0]    := '중분류명';
         Cells[C_DBD_3RDCD,   0]    := '소분류';
         Cells[C_DBD_3RDCDNM, 0]    := '소분류명';
         Cells[C_DBD_WORKINFO,0]    := '작업자';

         Left  := 312;
         Width := 332;

      end;
   end
   else if (in_Flag = 'SLIP코드') then
   begin
      with asg_DBMaster do
      begin
         ColWidths[C_DBM_FSTCD]     := 55;
         ColWidths[C_DBM_FSTCDNM]   := 190;
         ColWidths[C_DBM_WORKINFO]  := 47;

         Cells[C_DBM_FSTCD,   0]    := '그룹Slip';
         Cells[C_DBM_FSTCDNM, 0]    := '그룹명';
         Cells[C_DBM_WORKINFO,0]    := '작업자';

         Width := 313;

      end;

      with asg_DBDetail do
      begin
         ColWidths[C_DBD_SECCD]     := 53;
         ColWidths[C_DBD_SECCDNM]   := 140;
         ColWidths[C_DBD_3RDCD]     := 35;
         ColWidths[C_DBD_3RDCDNM]   := 35;
         ColWidths[C_DBD_WORKINFO]  := 50;

         RowHeights[1]    := 22;

         Cells[C_DBD_SECCD,   0]    := '상세Slip';
         Cells[C_DBD_SECCDNM, 0]    := '상세Slip명';
         Cells[C_DBD_3RDCD,   0]    := '부서';
         Cells[C_DBD_3RDCDNM, 0]    := 'Type';
         Cells[C_DBD_WORKINFO,0]    := '작업자';

         Left  := 312;
         Width := 332;

      end;
   end
   else if (in_Flag = 'VIEW') then
   begin
      with asg_DBMaster do
      begin
         ColWidths[C_DBM_FSTCD]     := 130;
         ColWidths[C_DBM_FSTCDNM]   := 0;
         ColWidths[C_DBM_WORKINFO]  := 50;

         Cells[C_DBM_FSTCD,   0]    := 'View';
         Cells[C_DBM_FSTCDNM, 0]    := '';
         Cells[C_DBM_WORKINFO,0]    := '생성일';

         Width := 200;
      end;

      with asg_DBDetail do
      begin
         ColWidths[C_DBD_SECCD]     := 0;
         ColWidths[C_DBD_SECCDNM]   := 0;
         ColWidths[C_DBD_3RDCD]     := 0;
         ColWidths[C_DBD_3RDCDNM]   := 420;
         ColWidths[C_DBD_WORKINFO]  := 0;

         RowHeights[1]    := 450;

         Cells[C_DBD_SECCD,   0]    := '';
         Cells[C_DBD_SECCDNM, 0]    := '';
         Cells[C_DBD_3RDCD,   0]    := '';
         Cells[C_DBD_3RDCDNM, 0]    := 'Source';
         Cells[C_DBD_WORKINFO,0]    := '';

         Left  := 202;
         Width := 440;

      end;
   end
   else if (in_Flag = 'DBSEARCH') then
   begin
      with asg_DBMaster do
      begin
         ColWidths[C_DBM_FSTCD]     := 20;
         ColWidths[C_DBM_FSTCDNM]   := 20;
         ColWidths[C_DBM_WORKINFO]  := 20;

         Cells[C_DBM_FSTCD,   0]    := '';
         Cells[C_DBM_FSTCDNM, 0]    := '';
         Cells[C_DBM_WORKINFO,0]    := '';

         Width := 0;
      end;

      with asg_DBDetail do
      begin

         ColWidths[C_DBD_SECCD]     := 80;
         ColWidths[C_DBD_SECCDNM]   := 150;
         ColWidths[C_DBD_3RDCD]     := 145;
         ColWidths[C_DBD_3RDCDNM]   := 70;
         ColWidths[C_DBD_WORKINFO]  := 175;

         RowHeights[1]    := 22;

         Cells[C_DBD_SECCD,   0]    := 'Type';
         Cells[C_DBD_SECCDNM, 0]    := '테이블명';
         Cells[C_DBD_3RDCD,   0]    := 'Comments';
         Cells[C_DBD_3RDCDNM, 0]    := '컬럼명';
         Cells[C_DBD_WORKINFO,0]    := 'Comments';

         Left  := 0;
         Width := 640;


      end;
   end;
end;

procedure TMainDlg.asg_DBMasterKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  Res : TPoint;
begin
   case (Key) of

      //[Ctrl + C] : 선택한 Cell 클립보드로 복사
      Ord('C') : if (ssCtrl in Shift) then
                 begin
                    asg_DBMaster.CopySelectionToClipboard;
                 end;

      //[Ctrl + A] : 해당 Grid Cell 전체내역 클립보드로 복사
      Ord('A') : if (ssCtrl in Shift) then
                 begin
                    asg_DBMaster.CopyToClipBoard;
                 end;

      //[F3 Key] : 다음 찾기
      Ord(VK_F3) :
                  begin
                     Res := asg_DBMaster.FindNext;

                     if (Res.X >= 0) and
                        (Res.Y >= 0) then
                     begin
                        asg_DBMaster.Col := Res.X;
                        asg_DBMaster.Row := Res.Y;
                     end
                     else
                        MessageBox(self.Handle,
                                   PChar('찾으시는 항목이 화면 List에 존재하지 않습니다.'),
                                   PChar('D/B 검색 : 결과내 재검색 결과 알림'),
                                   MB_OK + MB_ICONWARNING)
                                   ;
                  end;

      //[Ctrl + F] : 검색 EditBox 포커싱 @ 2016.11.18 LSH
      Ord('F') :
                  if (ssCtrl in Shift) then
                  begin
                     // 결과내 재검색 True 적용
                     if not fcb_ReScan.Checked then
                        fcb_ReScan.Checked := True;

                     fed_Analysis.SetFocus;
                  end;

   end;

end;

procedure TMainDlg.asg_DBDetailKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  Res : TPoint;
begin
   case (Key) of

      //[Ctrl + C] : 선택한 Cell 클립보드로 복사
      Ord('C') : if (ssCtrl in Shift) then
                 begin
                    asg_DBDetail.CopySelectionToClipboard;
                 end;

      //[Ctrl + A] : 해당 Grid Cell 전체내역 클립보드로 복사
      Ord('A') : if (ssCtrl in Shift) then
                 begin
                    asg_DBDetail.CopyToClipBoard;
                 end;

      //[F3 Key] : 다음 찾기
      Ord(VK_F3) :
                  begin
                     Res := asg_DBDetail.FindNext;

                     if (Res.X >= 0) and
                        (Res.Y >= 0) then
                     begin
                        asg_DBDetail.Col := Res.X;
                        asg_DBDetail.Row := Res.Y;
                     end
                     else
                        MessageBox(self.Handle,
                                   PChar('찾으시는 항목이 화면 List에 존재하지 않습니다.'),
                                   PChar('D/B 검색 : 결과내 재검색 결과 알림'),
                                   MB_OK + MB_ICONWARNING)
                                   ;
                  end;

      //[Ctrl + F] : 검색 EditBox 포커싱 @ 2016.11.18 LSH
      Ord('F') :
                  if (ssCtrl in Shift) then
                  begin
                     // 결과내 재검색 True 적용
                     if not fcb_ReScan.Checked then
                        fcb_ReScan.Checked := True;

                     fed_Analysis.SetFocus;
                  end;

   end;

end;



procedure TMainDlg.UpdateImage(in_Gubun,
                               in_FileName,
                               in_HideFile : String;
                               in_NowRow : Integer);
var
   TpSetMast   : TTpSvc;
   i, iCnt     : Integer;
   vType1,
   vLocate,
   vUserId,
   vUserNm,
   vDeptCd,
   vDutyFlag,
   vMobile,
   vEmail,
   vUserIp,
   vDelDate,
   vRegDate,
   vConText,
   vNickNm,
   vAttachNm,
   vHideFile,
   vEditIp : Variant;
   sType   : String;
begin

   iCnt  := 0;

   // 비상연락망 프로필 이미지 등록
   if (in_Gubun = 'PROFILE') then
   begin

      sType := '11';


      //-------------------------------------------------------------------
      // 1-1. Create Variables
      //-------------------------------------------------------------------
      with asg_Network do
      begin

         for i := in_NowRow to in_NowRow do
         begin
            //------------------------------------------------------------------
            // 1-2. Append Variables
            //------------------------------------------------------------------
            AppendVariant(vType1   ,   sType                  );
            AppendVariant(vLocate  ,   Cells[C_NW_LOCATE,   i]);
            AppendVariant(vUserNm  ,   Cells[C_NW_USERNM,   i]);
            AppendVariant(vUserIp  ,   Cells[C_NW_USERIP,   i]);
            AppendVariant(vAttachNm,   in_FileName            );
            AppendVariant(vHideFile,   in_HideFile            );
            AppendVariant(vEditIp  ,   FsUserIp               );
            AppendVariant(vDeptCd  ,   'ITT'                  );

            Inc(iCnt);
         end;
      end;
   end
   // 업무일지 작성 (신규 or 수정)
   else if (in_Gubun = 'WORKRPT') then
   begin

      sType := '12';

      //-------------------------------------------------------------------
      // 2-1. Create Variables
      //-------------------------------------------------------------------
      with asg_MultiIns do
      begin

         for i := in_NowRow to in_NowRow do
         begin
            //------------------------------------------------------------------
            // 2-2. Append Variables
            //------------------------------------------------------------------
            AppendVariant(vType1   ,   sType           );
            AppendVariant(vLocate  ,   Cells[i, 0]     );
            AppendVariant(vUserNm  ,   FsUserNm        );
            AppendVariant(vRegDate ,   Cells[i, 1]     );
            AppendVariant(vConText ,   Cells[i, 2]     );
            AppendVariant(vEditIp  ,   FsUserIp        );

            Inc(iCnt);
         end;
      end;
   end
   //------------------------------------------------------------
   // 주요업무공유 작성 (신규 or 수정) @ 2014.08.08 LSH
   //------------------------------------------------------------
   else if (in_Gubun = 'WORKCONN') then
   begin

      sType := '14';

      //-------------------------------------------------------------------
      // 2-1. Create Variables
      //-------------------------------------------------------------------
      with asg_MultiIns do
      begin
         for i := in_NowRow to in_NowRow do
         begin
            //------------------------------------------------------------------
            // 2-2. Append Variables
            //------------------------------------------------------------------
            AppendVariant(vType1   ,   sType           );
            AppendVariant(vLocate  ,   Cells[i, 0]     );
            AppendVariant(vUserNm  ,   FsUserNm        );
            AppendVariant(vRegDate ,   Cells[i, 1]     );
            AppendVariant(vConText ,   Cells[i, 2]     );
            AppendVariant(vDelDate ,   ''              );
            AppendVariant(vEditIp  ,   FsUserIp        );

            Inc(iCnt);
         end;
      end;
   end
   //------------------------------------------------------------
   // 주요업무공유 작성 삭제 @ 2014.08.12 LSH
   //------------------------------------------------------------
   else if (in_Gubun = 'WORKCONNDEL') then
   begin

      sType := '14';

      //-------------------------------------------------------------------
      // 2-1. Create Variables
      //-------------------------------------------------------------------
      with asg_WorkConn do
      begin
         for i := in_NowRow to in_NowRow do
         begin
            //------------------------------------------------------------------
            // 2-2. Append Variables
            //------------------------------------------------------------------
            AppendVariant(vType1   ,   sType           );
            AppendVariant(vLocate  ,   Cells[C_WC_LOCATE, Row]);
            AppendVariant(vUserNm  ,   FsUserNm        );
            AppendVariant(vRegDate ,   Cells[C_WC_REGDATE,Row]);
            AppendVariant(vConText ,   ''     );
            AppendVariant(vDelDate ,   in_FileName     );
            AppendVariant(vEditIp  ,   FsUserIp        );

            Inc(iCnt);
         end;
      end;
   end;



   if iCnt <= 0 then
      Exit;

   //-------------------------------------------------------------------
   // 3. Insert by TpSvc
   //-------------------------------------------------------------------
   TpSetMast := TTpSvc.Create;
   TpSetMast.Init(Self);


   Screen.Cursor := crHourGlass;


   try
      if TpSetMast.PutSvc('MD_KUMCM_M1',
                          [
                           'S_TYPE1'   , vType1
                         , 'S_STRING1' , vLocate
                         , 'S_STRING2' , vUserId
                         , 'S_STRING3' , vUserNm
                         , 'S_STRING4' , vDutyFlag
                         , 'S_STRING5' , vMobile
                         , 'S_STRING6' , vEmail
                         , 'S_STRING7' , vUserIp
                         , 'S_STRING8' , vEditIp
                         , 'S_STRING9' , vDeptCd
                         , 'S_STRING10', vDelDate
                         , 'S_STRING20', vRegDate
                         , 'S_STRING28', vContext
                         , 'S_STRING29', vAttachNm
                         , 'S_STRING30', vHideFile
                         , 'S_STRING41', vNickNm
                         ] ) then
      begin

         // 비상연락망 프로필 이미지 등록
         if (in_Gubun = 'PROFILE') then
         begin
            MessageBox(self.Handle,
                       PChar('이미지정보가 정상적으로 [업데이트]되었습니다.'),
                       '[KUMC 다이얼로그] 담당자 프로필 업데이트 알림 ',
                       MB_OK + MB_ICONINFORMATION);

            // Refresh
            SelGridInfo('DIALOG');
         end
         else if (in_Gubun = 'WORKRPT') then
         begin
            MessageBox(self.Handle,
                       PChar('업무일지가 정상적으로 [업데이트]되었습니다.'),
                       '[KUMC 다이얼로그] 업무일지 업데이트 알림 ',
                       MB_OK + MB_ICONINFORMATION);

            // 작성 Panel Close
            SetPanelStatus('WORKRPT', 'OFF');

            // Refresh
            SelGridInfo('WORKRPT');
         end
         else if (in_Gubun = 'WORKCONN') or
                 (in_Gubun = 'WORKCONNDEL')  then
         begin
            MessageBox(self.Handle,
                       PChar('주요 업무공유가 정상적으로 [업데이트]되었습니다.'),
                       '[KUMC 다이얼로그] 업무공유 업데이트 알림 ',
                       MB_OK + MB_ICONINFORMATION);

            // 작성 Panel Close
            SetPanelStatus('WORKCONN', 'OFF');

            // Refresh
            SelGridInfo('WORKCONN');
         end;
      end
      else
      begin
         ShowMessageM(GetTxMsg);
      end;

   finally
      FreeAndNil(TpSetMast);
      Screen.Cursor  := crDefault;
   end;

end;



procedure TMainDlg.asg_UserProfileDblClickCell(Sender: TObject; ARow,
  ACol: Integer);
var
   sFileName, sHideFile, sServerIp : String;
begin
   if ARow = R_PR_PHOTO then
   begin
      // 유효성 체크
      if PosByte(asg_UserProfile.Cells[1, R_PR_USERNM] , asg_NetWork.Cells[C_NW_USERNM, asg_NetWork.Row]) = 0 then
      begin
         MessageBox(self.Handle,
                    PChar('현재 프로필 조회중인 담당자정보가 선택된 정보와 일치하지 않습니다.' + #13#10 + '담당자 이름을 재클릭후, 이미지 등록 진행해주세요.'),
                    '비상연락망 담당자 프로필 이미지 등록전 확인',
                    MB_OK + MB_ICONERROR);
         Exit;
      end;


      // 이미지 등록 버튼 삭제
      asg_UserProfile.RemoveButton(1, R_PR_PHOTO);


      if od_File.Execute then
      begin
         // 첨부파일 Upload
         if (Trim(ExtractFileName(od_File.FileName)) <> '') then
         begin
            // 파일명 및 암호화명 Set
            sFileName := ExtractFileName(Trim(ExtractFileName(od_File.FileName)));
            sHideFile := 'KDIALAPPEND' + 'PRIMG' + FormatDateTime('YYYYMMDDHHNNSS', Now) + '.KDIALFILE';


            if Not FileUpLoad(sHideFile, sFileName, sServerIp) then
            begin
               Showmessage('첨부파일 UpLoad 중 에러가 발생했습니다.' + #13#10 + #13#10 +
                           '다시한번 시도해 주시기 바랍니다.');
               Exit;
            end
            else
            begin
               // 담당자 마스터에 이미지 정보 Update
               UpdateImage('PROFILE',
                           sFileName,
                           sHideFile,
                           asg_Network.Row);

            end;
         end;
      end;
   end;
end;

procedure TMainDlg.mi_InsWorkRptClick(Sender: TObject);
begin
   // 업무일지 작성 Panel-On
   SetPanelStatus('WORKRPT', 'ON');

end;

procedure TMainDlg.asg_MultiInsGetEditorType(Sender: TObject; ACol,
  ARow: Integer; var AEditor: TEditorType);
begin
   with asg_MultiIns do
   begin
      if PosByte('업무일지', apn_MultiIns.Caption.Text) > 0 then
      begin
         if (ARow = 1) then
         begin
            AEditor := edDateEdit;
         end
         else if (ARow = 2) then
         begin
            AEditor := edNormal;
         end;
      end;
   end;
end;


procedure TMainDlg.asg_MultiInsEditingDone(Sender: TObject);
begin
   with asg_MultiIns do
   begin
      if PosByte('업무일지', apn_MultiIns.Caption.Text) > 0 then
      begin
         if (Trim(Cells[1, 0]) <> '') and
            (Trim(Cells[1, 1]) <> '') and
            (Trim(Cells[1, 2]) <> '')
         then
         begin
            AddButton(1,
                      7,
                      ColWidths[1]-5,
                      20,
                      '업무일지 등록',
                      haBeforeText,
                      vaCenter);
         end
         else
            RemoveButton(1, 7);
      end;
   end;
end;

procedure TMainDlg.asg_MultiInsCanEditCell(Sender: TObject; ARow,
  ACol: Integer; var CanEdit: Boolean);
begin
   // 업무일지 작성시, 근무처(소속)는 Edit 불가
   if PosByte('업무일지', apn_MultiIns.Caption.Text) > 0 then
      CanEdit := (ARow = 1) or
                 (ARow = 2);

end;

procedure TMainDlg.asg_MultiInsButtonClick(Sender: TObject; ACol,
  ARow: Integer);
begin
   // 업무일지 등록(신규 or 수정)
   if PosByte('업무일지', apn_MultiIns.Caption.Text) > 0 then
   begin
      // 기존 셀의 TMemo 내용을 복사하고, 컴포넌트 삭제.
      if asg_MultiIns.Objects[1,2] <> nil then
      begin
         asg_MultiIns.Cells[1,2] := TMemo(asg_MultiIns.Objects[1,2]).Text;

         asg_MultiIns.Objects[1,2].Free;
         asg_MultiIns.Objects[1,2] := Nil;
      end;

      // 업무일지 작성내역 없으면, Msg 처리.
      if asg_MultiIns.Cells[1,2] = '' then
      begin
         MessageBox(self.Handle,
                    PChar('업데이트 할 업무일지 작성내역이 없습니다.'),
                    '업무일지 작성내역 Update 오류',
                    MB_OK + MB_ICONERROR);

         asg_MultiIns.SetFocus;

         Exit;
      end;


      // 업무일지 내역 Update
      UpdateImage('WORKRPT',
                  '',
                  '',
                  1);

   end
   // 주요업무공유 등록 (신규 or 수정)
   else if PosByte('주요업무공유', apn_MultiIns.Caption.Text) > 0 then
   begin
      // 기존 셀의 TMemo 내용을 복사하고, 컴포넌트 삭제.
      if asg_MultiIns.Objects[1,2] <> nil then
      begin
         asg_MultiIns.Cells[1,2] := TMemo(asg_MultiIns.Objects[1,2]).Text;

         asg_MultiIns.Objects[1,2].Free;
         asg_MultiIns.Objects[1,2] := Nil;
      end;

      // 주요업무공유 작성내역 없으면, Msg 처리.
      if asg_MultiIns.Cells[1,2] = '' then
      begin
         MessageBox(self.Handle,
                    PChar('업데이트 할 주요업무공유 작성내역이 없습니다.'),
                    '주요업무공유 작성내역 Update 오류',
                    MB_OK + MB_ICONERROR);

         asg_MultiIns.SetFocus;

         Exit;
      end;


      // 주요업무공유 내역 Update
      UpdateImage('WORKCONN',
                  '',
                  '',
                  1);

   end;
end;

procedure TMainDlg.asg_MultiInsGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   if ACol = 0 then
   begin
      HAlign := taCenter;
      VAlign := vtaCenter;
   end;
end;



//------------------------------------------------------------------------------
// [통합출력] FTP 이미지 통합 출력 함수
//
// Date   : 2014.02.25
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.KDialPrint(in_ImgCode  : String;
                              in_GridName : TAdvStringGrid);
var
//   FForm : TForm;
   tmpDocTitle : String;

   SetKDialPrt : TFtpPrint;
   i : Integer;
   //---------------------------------------------------------------------------
   //  함수개요 : 오라클함수중 LPAD와 같은 기능을 한다.
   //             원하는 SIZE만큼 왼쪽을 원하는 문자열로 채운다.
   //  인 자 값 : 1.ASourceStr(String) : 대상문자열
   //             2.PadLen(String)     : 채울 size
   //             3.PadStr(String)     : 채울문자
   //  리턴구분 : LPAD한 String
   //  사 용 예 : A := LPAD('123',7,'0') --> '0000123'
   //  함수출처 : SQCOMCLS.pas
   //---------------------------------------------------------------------------
   function LPad(ASourceStr : String;
                 PadLen     : Integer;
                 PadStr     : String): String;
   var
      i : Integer;
      ResultStr : String;
   begin
      ResultStr := ASourceStr;

      for i := 1 to (PadLen - LengthByte(ASourceStr)) do
         ResultStr := CopyByte(PadStr,1,1) + ResultStr;

      Result := ResultStr;
   end;
begin
   //---------------------------------------------------------------
   // Printer 세팅 Check
   //---------------------------------------------------------------
   //if (IsNotPrinterReady) then exit;



   Screen.Cursor := crHourglass;



   try
      //------------------------------------------------------------
      // 1. 시스템 유지보수 지급 요청서 출력
      //------------------------------------------------------------
      if (in_ImgCode = 'SYSTEM') then
      begin
         //------------------------------------------------------------------
         //
         //------------------------------------------------------------------
         SetKDialPrt := TFtpPrint.Create(Self);

         if SetKDialPrt = nil then
            Application.CreateForm(TFtpPrint, SetKDialPrt);


         with SetKDialPrt do
         begin
            qrlb_ImgCode.Caption := in_ImgCode;



            with in_GridName do
            begin
               // 회계년도 (yyyy-mm)
               if CopyByte(FormatDateTime('mm', Date), 1, 1) = '0' then
                  tmpDocTitle := FormatDateTime('yyyy', Date) + '년 ' + CopyByte(FormatDateTime('mm', Date - 29), 2, 1) + '월 ' + Cells[C_SD_DOCTITLE, Row]
               else
                  tmpDocTitle := FormatDateTime('yyyy', Date) + '년 ' + FormatDateTime('mm', Date - 29) + '월 ' + Cells[C_SD_DOCTITLE, Row];


               qrlb_DocTitle.Caption   := tmpDocTitle;                           // 문서제목
               qrlb_Company.Caption    := Cells[C_SD_RELDEPT,  Row];             // 계약처
               qrlb_DocName.Caption    := Cells[C_SD_DOCTITLE, Row];             // 계약명
               qrlb_Period.Caption     := Cells[C_SD_DOCRMK,   Row];             // 계약기간
               qrlb_TotAmt1.Caption    := Cells[C_SD_TOTRMK,   Row];             // 월간 총 지급액
               qrlb_HqAmt.Caption      := Cells[C_SD_HQREMARK, Row];             // 월간 의료원 지급액
               qrlb_AaAmt.Caption      := Cells[C_SD_AAREMARK, Row];             // 월간 안암 지급액
               qrlb_GrAmt.Caption      := Cells[C_SD_GRREMARK, Row];             // 월간 구로 지급액
               qrlb_AsAmt.Caption      := Cells[C_SD_ASREMARK, Row];             // 월간 안산 지급액
               qrlb_TotAmt2.Caption    := Cells[C_SD_TOTRMK,   Row];             // 월간 총 지급액
               qrlb_DutyUser.Caption   := FsUserNm;                              // 문서 기안자
               qrlb_Manager.Caption    := FsMngrNm;                              // 문서 관리자 (팀장) @ 2014.07.18 LSH
            end;

            qr_KDial.Print;

            Close;
         end;
      end
      //------------------------------------------------------------
      // 2. 시스템 유지보수 지급 요청서 일괄출력
      //       - 유효한 계약만 일괄 출력
      //------------------------------------------------------------
      else if (in_ImgCode = 'SYSTEMALL') then
      begin
         //------------------------------------------------------------------
         //
         //------------------------------------------------------------------
         SetKDialPrt := TFtpPrint.Create(Self);


         if SetKDialPrt = nil then
            Application.CreateForm(TFtpPrint, SetKDialPrt);


         with SetKDialPrt do
         begin
            qrlb_ImgCode.Caption := CopyByte(in_ImgCode, 1, 6);



            with in_GridName do
            begin

               for i := 1 to RowCount do
               begin
                  // 계약 만료된 항목은 출력 Skip.
                  if LengthByte(Trim(CopyByte(Cells[C_SD_DOCRMK,  i], 19, 2))) > 1 then
                  begin
                     if ((CopyByte(Cells[C_SD_DOCRMK,  i], 14, 4) +
                          CopyByte(Cells[C_SD_DOCRMK,  i], 19, 2) +
                          CopyByte(Cells[C_SD_DOCRMK,  i], 23, 4)) < FormatDateTime('yyyymmdd', Date)) then
                     begin
                        Continue;
                     end;
                  end
                  else
                  begin
                     if ((CopyByte(Cells[C_SD_DOCRMK,       i], 14, 4) + '0' +
                          Trim(CopyByte(Cells[C_SD_DOCRMK,  i], 19, 2)) +
                          CopyByte(Cells[C_SD_DOCRMK,       i], 23, 4)) < FormatDateTime('yyyymmdd', Date)) then
                     begin
                        Continue;
                     end;
                  end;


                  // 회계년도 (yyyy-mm)
                  if CopyByte(FormatDateTime('mm', Date), 1, 1) = '0' then
                     tmpDocTitle := FormatDateTime('yyyy', Date) + '년 ' + CopyByte(FormatDateTime('mm', Date - 29), 2, 1) + '월 ' + Cells[C_SD_DOCTITLE, i]
                  else
                     tmpDocTitle := FormatDateTime('yyyy', Date) + '년 ' + FormatDateTime('mm', Date - 29) + '월 ' + Cells[C_SD_DOCTITLE, i];



                  qrlb_DocTitle.Caption   := tmpDocTitle;                           // 문서제목
                  qrlb_Company.Caption    := Cells[C_SD_RELDEPT,  i];               // 계약처
                  qrlb_DocName.Caption    := Cells[C_SD_DOCTITLE, i];               // 계약명
                  qrlb_Period.Caption     := Cells[C_SD_DOCRMK,   i];               // 계약기간
                  qrlb_TotAmt1.Caption    := Cells[C_SD_TOTRMK,   i];               // 월간 총 지급액
                  qrlb_HqAmt.Caption      := Cells[C_SD_HQREMARK, i];               // 월간 의료원 지급액
                  qrlb_AaAmt.Caption      := Cells[C_SD_AAREMARK, i];               // 월간 안암 지급액
                  qrlb_GrAmt.Caption      := Cells[C_SD_GRREMARK, i];               // 월간 구로 지급액
                  qrlb_AsAmt.Caption      := Cells[C_SD_ASREMARK, i];               // 월간 안산 지급액
                  qrlb_TotAmt2.Caption    := Cells[C_SD_TOTRMK,   i];               // 월간 총 지급액
                  qrlb_DutyUser.Caption   := FsUserNm;                              // 문서 기안자
                  qrlb_Manager.Caption    := FsMngrNm;                              // 문서 관리자 (팀장) @ 2014.07.18 LSH

                  qr_KDial.Print;
               end;
            end;

            Close;

         end;
      end;

   finally
      screen.Cursor  :=  crDefault;
   end;
end;


procedure TMainDlg.mi_SystemPrintClick(Sender: TObject);
begin
   // 시스템 유지보수 지급 양식 출력 (FTP)
   KDialPrint('SYSTEM',
              asg_SelDoc);
end;




procedure TMainDlg.pm_DocPopup(Sender: TObject);
begin
   with asg_SelDoc do
   begin
      //--------------------------------------------------
      // 시스템 유지보수 지급요청서 출력
      //--------------------------------------------------
      if (Cells[C_SD_DOCLIST, Row] = '계약처') then
      begin
         mi_SystemPrint.Visible     := True;
         mi_SystemAllPrint.Visible  := True;
         mi_CRADelUp.Visible        := False;
         mi_CopyDocInfo.Visible     := True;
      end
      else if (Cells[C_SD_DOCLIST, Row] = 'CRA/CRC') then
      begin
         mi_SystemPrint.Visible     := False;
         mi_SystemAllPrint.Visible  := False;
         mi_CRADelUp.Visible        := True;
         mi_CopyDocInfo.Visible     := True;
      end
      else
      begin
         mi_SystemPrint.Visible     := False;
         mi_SystemAllPrint.Visible  := False;
         mi_CRADelUp.Visible        := False;
         mi_CopyDocInfo.Visible     := True;
      end;
   end;
end;

procedure TMainDlg.mi_SMSClick(Sender: TObject);
var
   iStartRow, iEndRow, i{, temp} : integer;
begin

   {
   // SMS 권한 테스트용
   if (PosByte('안암도메인5.40', FsUserIp) = 0) then
   begin
      MessageDlg('권한이 불충분합니다.', Dialogs.mtError, [Dialogs.mbOK], 0);
      Exit;
   end;
   }


//   i         := 0;
   iStartRow := 0;
   iEndRow   := 0;


   // 1. 변수 세팅 (멀티Row 시작-끝)
   if (iSelRow1 > iSelRow2) then
   begin
      iStartRow := iSelRow2;
      iEndRow   := iSelRow1;
   end
   else if (iSelRow1 <= iSelRow2) then
   begin
      iStartRow := iSelRow1;
      iEndRow   := iSelRow2;
   end;




   asg_RecvList.RowCount := 2;



   // 2. 일괄입력 유효성체크
   //  -- 마우스 Multi-Dragging만 사용할 경우.
   for i := iStartRow to iEndRow do
   begin
      with asg_RecvList do
      begin
         // 비상연락망
         if (ftc_Dialog.ActiveTab = AT_DIALOG) then
         begin
            Cells[0, i - iStartRow + 1] := asg_Network.Cells[C_NW_USERNM, i];
            Cells[1, i - iStartRow + 1] := asg_Network.Cells[C_NW_MOBILE, i];
         end
         // 다이얼 Book
         else if (ftc_Dialog.ActiveTab = AT_DIALBOOK) then
         begin
            if (
                  (CopyByte(asg_DialList.Cells[C_DL_CALLNO, asg_DialList.Row], 1, 3) = '010') and
                  (TokenStr(asg_DialList.Cells[C_DL_CALLNO, asg_DialList.Row], #13#10, 2) = '')
               ) then
            begin
               Cells[0, i - iStartRow + 1] := asg_DialList.Cells[C_DL_DUTYUSER, i];
               Cells[1, i - iStartRow + 1] := asg_DialList.Cells[C_DL_CALLNO, i];
            end
            else if
               (
                  (CopyByte(asg_MyDial.Cells[C_MD_MOBILE, asg_MyDial.Row], 1, 3) = '010') and
                  (TokenStr(asg_MyDial.Cells[C_MD_MOBILE, asg_MyDial.Row], #13#10, 2) = '')
               ) then
            begin
               Cells[0, i - iStartRow + 1] := asg_MyDial.Cells[C_MD_DUTYUSER, i];
               Cells[1, i - iStartRow + 1] := asg_MyDial.Cells[C_MD_MOBILE, i];
            end
            else if
               (  // SMS 발송 & 구글 Photos 연동 위해 추가 @ 2017.07.11 LSH
                  PosByte('010', TokenStr(asg_DialList.Cells[C_DL_CALLNO, asg_DialList.Row], #13#10, 2)) > 0
               ) then
            begin
               Cells[0, i - iStartRow + 1] := asg_DialList.Cells[C_DL_DUTYUSER, i];
               Cells[1, i - iStartRow + 1] := Trim(TokenStr(asg_DialList.Cells[C_DL_CALLNO, asg_DialList.Row], #13#10, 2));
            end
            else if
               (  // SMS 발송 & 구글 Photos 연동 위해 추가 @ 2017.07.11 LSH
                  PosByte('010', TokenStr(asg_MyDial.Cells[C_MD_MOBILE, asg_MyDial.Row], #13#10, 2)) > 0
               ) then
            begin
               Cells[0, i - iStartRow + 1] := asg_MyDial.Cells[C_MD_DUTYUSER, i];
               Cells[1, i - iStartRow + 1] := Trim(TokenStr(asg_MyDial.Cells[C_MD_MOBILE, asg_MyDial.Row], #13#10, 2));
            end
            // 그 외 Case 적용 @ 2017.10.26 LSH
            else
            begin
               if (CopyByte(asg_DialList.Cells[C_DL_CALLNO, asg_DialList.Row], 1, 3) = '010') then
               begin
                  Cells[0, i - iStartRow + 1] := asg_DialList.Cells[C_DL_DUTYUSER, i];
                  Cells[1, i - iStartRow + 1] := Trim(asg_DialList.Cells[C_DL_CALLNO, i]);
               end
               else
               begin
                  Cells[0, i - iStartRow + 1] := asg_MyDial.Cells[C_MD_DUTYUSER, i];
                  Cells[1, i - iStartRow + 1] := Trim(asg_MyDial.Cells[C_MD_MOBILE, asg_MyDial.Row]);
               end;
            end;
         end;

         RowCount := RowCount + 1;

      end;
   end;


   { Ctrl 키를 활용한 multi Selection 테스트 ..
   for i := 1 to asg_Network.RowCount do
   begin
      if asg_NetWork.IsSelected(C_NW_MOBILE, i) then
      begin
         with asg_RecvList do
         begin
            Cells[0, RowCount - 1] := asg_Network.Cells[C_NW_USERNM, i];
            Cells[1, RowCount - 1] := asg_Network.Cells[C_NW_MOBILE, i];

            RowCount := RowCount + 1;
         end;
      end;
   end;
   }


   {  -- FormShow에서 로그인시점에 가져오기 때문에 주석처리 @ 2014.04.30 LSH
   for j := 1 to asg_Network.RowCount do
   begin
      if asg_Network.Cells[C_NW_USERNM, j] = FsUserNm then
      begin
         FsUserMobile := asg_NetWork.Cells[C_NW_MOBILE, j];

         if (asg_Network.Cells[C_NW_LOCATE, j] = '의료원') or
            (asg_Network.Cells[C_NW_LOCATE, j] = '안암') then
            FsUserCallNo := '02-920-' + asg_Network.Cells[C_NW_CALLNO, j]
         else if (asg_Network.Cells[C_NW_LOCATE, j] = '구로') then
            FsUserCallNo := '02-2626-' + asg_Network.Cells[C_NW_CALLNO, j]
         else if (asg_Network.Cells[C_NW_LOCATE, j] = '안산') then
            FsUserCallNo := '031-412-' + asg_Network.Cells[C_NW_CALLNO, j];

         Break;
      end;
   end;
   }


   // 최종 RowCount 정리
   asg_RecvList.RowCount := asg_RecvList.RowCount - 1;



   fm_SMS.Lines.Text    := '내용입력은 80byte (한글기준 40자정도)까지 가능하며, 발신자 번호는 사무실 번호로 자동 세팅됩니다. 입력후 우측 [전송]버튼을 눌러주세요.';
   apn_SMS.Caption.Text := ' 업무용 SMS';

   apn_SMS.Top       := 136;
   apn_SMS.Left      := 72;
   apn_SMS.Collaps   := True;
   apn_SMS.Visible   := True;
   apn_SMS.Collaps   := False;
   apn_SMS.BringToFront;


   iSelRow1 := iStartRow;
   iSelRow2 := iEndRow;

end;

procedure TMainDlg.asg_NetworkMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
   NowCol, NowRow : integer;
begin
   //---------------------------------------------------------------
   // 마우스 좌클릭(on)시 row-index 가져온다
   //---------------------------------------------------------------
   if (Button = mbLeft) then
   begin
      iSelRow1 := 0;

      with asg_Network do
      begin
         // Mouse 포인터 위치 받아오기
         MouseToCell(X, Y, NowCol, NowRow);

         iSelRow1 := NowRow;
      end
   end;
end;

procedure TMainDlg.asg_NetworkMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
   NowCol, NowRow : integer;
begin
   //---------------------------------------------------------------
   // 마우스 좌클릭(off)시 row-index 가져온다
   //---------------------------------------------------------------
   if (Button = mbLeft) then
   begin
      iSelRow2 := 0;

      with asg_Network do
      begin
         // Mouse 포인터 위치 받아오기
         MouseToCell(X, Y, NowCol, NowRow);

         iSelRow2 := NowRow;
      end
   end;
end;



procedure TMainDlg.fm_SMSClick(Sender: TObject);
begin
   if PosByte('80byte (한글기준 40자정도)' , fm_SMS.Lines.Text) > 0 then
      fm_SMS.Lines.Clear;
end;


// 업무용 SMS 발송 (등록)
procedure TMainDlg.fsbt_SendMsgClick(Sender: TObject);
var
  {Send_Cnt,} i, j, iCnt : Integer;

  {l_hh,l_mm, }sType, sReceiver, sSender, sRsvtime, sMessage : String;
  vType1   ,
  vSender  ,
  vReceiver,
  vContext ,
  vRegTime ,
  vUserIp  : Variant;
  TpSetSMS : TTpSvc;

begin
   {
   if Trim(me_SmsMsg.Text) = '' then
   begin
      ShowMessage('메시지를 입력하십시오.');
      me_SmsMsg.SetFocus;
      Exit;
   end;
   }

   {
   if MessageDlg('선택된 수신자들에게 메시지를 발송하시겠습니까?',
                  mtConfirmation, [mbYes, mbNo], 0) <> mrYes then Exit;
   }


   iCnt := 0;


   for i := 1 to asg_RecvList.RowCount - 1 do
   begin




        {
        // [예약전송] 시
        if rbt_Type2.Checked = True then
        begin
           if (cbx_hh.Text = '') or (cbx_mm.Text = '') then
           begin
             ShowMessage('예약시간/분을 입력해주십시오.');
             Exit;
           end;


           l_hh := cbx_hh.Text;
           l_mm := cbx_mm.Text;


           if (DelChar(FormatDateTime('yyyy-mm-dd',dtp_Reserve.Date),'-') + l_hh + l_mm + '00')  < FormatDateTime('yyyymmddhh24miss',Now) then
           begin
             ShowMessage('예약전송은 현재시각 이후부터 가능합니다.');
             Exit;
           end;

           // 예약시간 설정.
           sRsvtime := DelChar(FormatDateTime('yyyy-mm-dd',dtp_Reserve.Date),'-') + l_hh + l_mm + '00';


        end
        else
        begin
           sRsvtime := '';
        end;


        // 병원명 추가 (구로병원 오픈에 따른), 2012.02.17 LSH
        if (G_Locate = 'AA') then
           sLocateNm    := '[고대안암병원] '
        else if (G_Locate = 'GR') then
           sLocateNm    := '[고대구로병원] '
        else if (G_Locate = 'AS') then
           sLocateNm    := '[고대안산병원] ';
        }


      if PosByte('80byte (한글기준 40자정도)' , fm_SMS.Lines.Text) > 0 then
         if Application.MessageBox('작성하신 메세지의 수정사항이 없습니다. 계속 진행하시겠습니까?',
                                   PChar('업무용 SMS 발송전 확인'), MB_OKCANCEL) <> IDOK then
            Exit;


      if Trim(fm_SMS.Lines.Text) = '' then
      begin
         MessageBox(self.Handle,
                    PChar('입력하신 메세지 내용이 없습니다. 확인해주세요.'),
                    '업무용 SMS 발송전 확인',
                    MB_OK + MB_ICONINFORMATION);

         fm_SMS.SetFocus;
         Exit;
      end;


      if Application.MessageBox(PChar('작성하신 메세지를 수신자에게 [발송]하시겠습니까?' + #13#10 + #13#10 + fm_SMS.Lines.Text),
                                PChar('업무용 SMS 발송전 확인'),
                                MB_OKCANCEL) <> IDOK then
         Exit;



      // 송, 수신자 및 메세지 설정.
      sSender      := FsUserCallNo;
      sReceiver    := DelChar(asg_RecvList.Cells[1, i],'-');
      sMessage     := fm_SMS.Lines.Text;
      sRsvTime     := '';
      sType        := '13';



      //------------------------------------------------------------------
      // 2-1. Append Variables
      //------------------------------------------------------------------
      AppendVariant(vType1   ,   sType                );
      AppendVariant(vSender  ,   FsUserCallNo         );
      AppendVariant(vReceiver,   DelChar(asg_RecvList.Cells[1, i],'-'));
      AppendVariant(vContext ,   fm_SMS.Lines.Text    );
      AppendVariant(vRegTime ,   sRsvTime             );
      AppendVariant(vUserIp  ,   FsUserIp             );

      Inc(iCnt);


      if iCnt <= 0 then
         Exit;


   end;




   // iCnt 만큼 메세지 전송
   for j := 0 to iCnt - 1 do
   begin

      //-------------------------------------------------------------------
      // 3. Insert by TpSvc
      //-------------------------------------------------------------------
      TpSetSMS := TTpSvc.Create;
      TpSetSMS.Init(Self);


      Screen.Cursor := crHourGlass;


      try
         if TpSetSMS.PutSvc('MD_KUMCM_M1',
                            [
                              'S_TYPE1'   , vType1
                            , 'S_STRING7' , vUserIp
                            , 'S_STRING20', vRegTime
                            , 'S_STRING28', vContext
                            , 'S_STRING46', vReceiver
                            , 'S_STRING47', vSender
                            ] ) then
         begin
            // Send Messages
         end
         else
         begin
            ShowMessageM(GetTxMsg);
         end;


      finally
         FreeAndNil(TpSetSMS);

         Screen.Cursor  := crDefault;
      end;
   end;

   // SMS 전송완료 메세지
   MessageBox(self.Handle,
                       PChar(IntToStr(iCnt) + ' 건의 SMS가 발송[등록] 되었습니다.'),
                       '[KUMC 다이얼로그] 업무용 SMS 발송 알림 ',
                       MB_OK + MB_ICONINFORMATION);

   apn_SMS.Collaps := True;
   apn_SMS.Visible := False;

end;

procedure TMainDlg.asg_RecvListGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   HAlign := taCenter;
   VAlign := vtaCenter;
end;

// Grid to Excel 변환
procedure TMainDlg.mi_ExcelClick(Sender: TObject);
begin
   if ftc_Dialog.ActiveTab = AT_DIALOG then
      ExcelSave(asg_Network, '정보전산팀 비상연락망 [Last Updated  : ' + asg_NwUpdate.Cells[0, 1] + ']', 0)
   else if ftc_Dialog.ActiveTab = AT_DOC then
      ExcelSave(asg_Release, '개발자 릴리즈대장 [' + fmed_RegFrDt.Text + '~' + fmed_RegToDt.Text + ']', 0)
   else if ftc_Dialog.ActiveTab = AT_ANALYSIS then
   begin
      // 주간리포트 Excel 변환 기능 추가 @ 2016.12.08 LSH
      if (apn_Weekly.Visible) then
         ExcelSave(asg_WeeklyRpt, '업무 History 리포트 [' + fmed_AnalFrom.Text + '~' + fmed_AnalTo.Text + ']', 0)
      else
         ExcelSave(asg_WorkRpt, '업무일지 [' + fmed_AnalFrom.Text + '~' + fmed_AnalTo.Text + ']', 0);
   end
   else if ftc_Dialog.ActiveTab = AT_WORKCONN then
      ExcelSave(asg_WorkConn, '주요업무공유', 0);

end;


procedure TMainDlg.ExcelSave(ExcelGrid : TAdvStringGrid; Title : String; Disp : Integer);
var
   ExcelAP1 : TExcelApplication;
   ExcelWS1 : TExcelWorksheet;
   ExcelWB1 : TExcelWorkbook;
   ExcelAP2 : Variant;
   SaveD1 : TSaveDialog;
   FName1 : String;
   i,j : LongInt;
   range1 : ExcelRange;
begin
   if Application.MessageBox('Excel화일로 저장하시겠습니까?', 'Excel저장', MB_OKCANCEL) <> IDOK then
      exit;

   if Disp = 1 then
   begin
      SaveD1 := TSaveDialog.Create(nil);
      SaveD1.InitialDir := 'C:\Users\Administrator\Desktop';
      SaveD1.DefaultExt := 'xls';
      SaveD1.FileName := Title;
      SaveD1.Filter := '엑셀화일(*.xls)|*.xls|Ms';

      if SaveD1.Execute = False then
      begin
         FreeAndNil(SaveD1);
         exit;
      end
      else
         FName1 := SaveD1.FileName;

      FreeAndNil(SaveD1);
   end;



   Screen.Cursor := crHourglass;



   if Disp = 1 then
   begin
      try
         ExcelAP1 := TExcelApplication.Create(nil);
         ExcelWB1 := TExcelWorkbook.Create(nil);
         ExcelWS1 := TExcelWorksheet.Create(nil);

         ExcelAP1.Connect;
         ExcelAP1.Workbooks.Add(xlWBATWorksheet,0);
         ExcelWB1.ConnectTo(ExcelAP1.Workbooks.Item[1]);
         EXcelWS1.ConnectTo(ExcelAP1.Sheets[1] as _Worksheet);

      except
         Screen.Cursor := crDefault;

         showmessage('Excel을 찾을수 없습니다.');

         ExcelAP1.Disconnect;

         FreeAndNil(ExcelWS1);
         FreeAndNil(ExcelWB1);
         FreeAndNil(ExcelAP1);
         exit;
      end;

      with EXcelWS1 do
      begin
         // 타이틀
         range1 := range[cells.item[3,2],cells.item[ExcelGrid.Colcount +3,ExcelGrid.RowCount+2]];
         Cells.Item[1,3] := Title;

         // 처리 Routine (grid로 받음)
         for i := 0 to ExcelGrid.RowCount - 1 do
         begin
            for j := 0 to ExcelGrid.Colcount - 1 do
            begin
                Cells.Numberformatlocal := '@';  // 문자로 인식해야할 경우
                Cells.Item[i + 3, j + 2] := string(ExcelGrid.Cells[j,i]);
            end;
         end;

         if FileExists(FName1) then
         begin
            DeleteFile(FName1);
         end;

         range1.Columns.AutoFit;

         SaveAs(FName1);

      end;

      ExcelAP1.Disconnect;
      ExcelAP1.Quit;

      FreeAndNil(ExcelWS1);
      FreeAndNil(ExcelWB1);
      FreeAndNil(ExcelAP1);

      Screen.Cursor := crDefault;

      showmessage('Excel화일로 저장되었습니다.');

   end
   else
   begin
       try
          ExcelAP2 := CreateOLEObject('Excel.Application');
          ExcelAP2.WorkBooks.add;
          ExcelAP2.visible := true;

          // 처리 Routine (grid로 받음)
          // 타이틀
          ExcelAP2.worksheets[1].cells[1,3] := Title;


          for i := 0 to ExcelGrid.RowCount - 1 do
          begin
             for j := 0 to ExcelGrid.ColCount - 1 do
             begin
                ExcelAP2.worksheets[1].Cells[i + 3, j + 2] := '''' + String(ExcelGrid.Cells[j,i]);
             end;
          end;

          //Range(세로,세로)에서 Range(가로,세로)로 변경함 -2011.03.30 smpark
          //ExcelAP2.Range['A1',CHR(64+ ExcelGrid.RowCount+1)+IntToStr(i + 3)].select;
          ExcelAP2.Range['A1',CHR(64+ (j + 3))+ inttostr(i+2)].select;
          ExcelAP2.Selection.Font.Name:='굴림체';
          ExcelAP2.Selection.Font.Size:=10;
          ExcelAP2.selection.Columns.AutoFit;
          ExcelAP2.Range['A1','A1'].select;

       except
          Screen.Cursor := crDefault;

          ExcelAP2.Quit;

          showmessage('Excel을 찾을수 없습니다.');

          exit;
       end;

       Screen.Cursor := crDefault;

   end;
end;

procedure TMainDlg.pm_WorkPopup(Sender: TObject);
begin
   if (IsLogonUser) then
   begin
      if ftc_Dialog.ActiveTab = AT_DOC then
      begin
         mi_InsWorkRpt.Visible := False;
         mi_Excel.Visible      := True;
         mi_GridPrint.Visible  := False;
         mi_Modify.Visible     := False;
         mi_Delete.Visible     := False;

         // 릴리즈대장 On
         if (apn_Release.Visible) then
         begin
            mi_RlzScript.Visible  := True;
            mi_BPL2Delphi.Visible := True;   // BPL --> 델파이 자동등록 메뉴 추가 @ 2016.11.02 LSH
            mi_DelRelease.Visible := True;   // 릴리즈 이력 삭제 @ 2018.06.12 LSH
         end
         // 그 외 (근무표, 문서관리 등)
         else
         begin
            mi_RlzScript.Visible  := False;
            mi_BPL2Delphi.Visible := False;   // BPL --> 델파이 자동등록 메뉴 추가 @ 2016.11.02 LSH
            mi_DelRelease.Visible := False;   // 릴리즈 이력 삭제 @ 2018.06.12 LSH
         end;
      end
      else if ftc_Dialog.ActiveTab = AT_ANALYSIS then
      begin
         // 주간리포트 Excel 변환 팝업메뉴 추가 @ 2016.12.08 LSH
         if (apn_Weekly.Visible) then
         begin
            mi_InsWorkRpt.Visible := False;
            mi_Excel.Visible      := True;
            mi_GridPrint.Visible  := False;
            mi_Modify.Visible     := False;
            mi_Delete.Visible     := False;
            mi_RlzScript.Visible  := False;
            mi_BPL2Delphi.Visible := False;   // BPL --> 델파이 자동등록 메뉴 추가 @ 2016.11.02 LSH
         end
         else
         begin
            mi_InsWorkRpt.Visible := True;
            mi_Excel.Visible      := True;
            mi_GridPrint.Visible  := False;
            mi_Modify.Visible     := False;
            mi_Delete.Visible     := False;
            mi_RlzScript.Visible  := False;
            mi_BPL2Delphi.Visible := False;   // BPL --> 델파이 자동등록 메뉴 추가 @ 2016.11.02 LSH
         end;
      end
      else if ftc_Dialog.ActiveTab = AT_DIALBOOK then
      begin
         mi_InsWorkRpt.Visible := False;
         mi_Excel.Visible      := False;
         mi_GridPrint.Visible  := True;
         mi_Modify.Visible     := False;
         mi_Delete.Visible     := False;
         mi_RlzScript.Visible  := False;
         mi_BPL2Delphi.Visible := False;   // BPL --> 델파이 자동등록 메뉴 추가 @ 2016.11.02 LSH
      end
      // 주간업무공유 (현재 미사용)
      else if ftc_Dialog.ActiveTab = AT_WORKCONN then
      begin
         mi_InsWorkRpt.Visible := False;
         mi_Excel.Visible      := True;
         mi_GridPrint.Visible  := False;
         mi_Modify.Visible     := True;
         mi_Delete.Visible     := True;
         mi_RlzScript.Visible  := False;
         mi_BPL2Delphi.Visible := False;   // BPL --> 델파이 자동등록 메뉴 추가 @ 2016.11.02 LSH
      end;
   end;
end;

procedure TMainDlg.mi_GridPrintClick(Sender: TObject);
var
   varResult : String;
begin
   if not IsLogonUser then
   begin
      MessageBox(self.Handle,
                 PChar('[비상연락망] - [담당자정보]에 담당자 IP 등록후 사용하실 수 있습니다.'),
                 PChar(Self.Caption + ' : 등록되지 않은 User 제한 알림'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;


   //-------------------------------------------------------------------
   // 2. Set Print Options
   //-------------------------------------------------------------------
   SetPrintOptions;


   if (apn_DialMap.Visible) then
   begin


      //-------------------------------------------------------------------
      // 3. Print
      //-------------------------------------------------------------------
      asg_DialMap.Print;



      // 로그 Update
      UpdateLog('DIALMAP',
                'PRINT',
                FsUserIp,
                'P',
                FsUserSpec,
                FsUserNm,
                varResult
                );


      // Comments
      lb_DialScan.Caption := '▶ 다이얼Map 출력이 완료되었습니다.';

   end



end;

procedure TMainDlg.asg_DialMapPrintPage(Sender: TObject; Canvas: TCanvas;
  PageNr, PageXSize, PageYSize: Integer);
var
   savefont : tfont;
   ts,tw    : integer;
   RecName, addInfo  : String;
const
   myowntitle : string = ' ';
begin
   if asg_DialMap.PrintColStart <> 0 then exit;


   // Report Name 정의
   RecName := '정보전산실 업무배치도 (다이얼Map)';
   addInfo := ' Fax no. [안암] 02-920-5672 [구로] 02-2626-2239 [안산] 031-412-5999';


   // 의료원명 Display
   with canvas do
   begin
      savefont := TFont.Create;
      savefont.assign(font);

      font.name   := 'Arial';
      font.Size   := 10;
      font.Style  := [];
      font.height := asg_DialMap.mapfontheight(10);
      font.Color  := clBlack;

      ts := asg_DialMap.Printcoloffset[0];
      tw := asg_DialMap.Printpagewidth;
      ts := ts+ ((tw-textwidth('고려대학교의료원')) shr 1);

      Textout(ts, -15, '고려대학교의료원');

      font.assign(savefont);

      savefont.free;
   end;


   // 기록지명 Display
   with canvas do
   begin
      savefont := TFont.Create;
      savefont.assign(font);

      font.name   := 'Arial';
      font.Size   := 20;
      font.Style  := [fsBold];
      font.height := asg_DialMap.mapfontheight(20);
      font.Color  := clBlack;

      ts := asg_DialMap.Printcoloffset[0];
      tw := asg_DialMap.Printpagewidth;
      ts := ts+ ((tw-textwidth(RecName)) shr 1);

      Textout(ts, -50, RecName);

      font.assign(savefont);

      savefont.free;
   end;

   // 최종수정 ver. Display
   with canvas do
   begin
      savefont := TFont.Create;
      savefont.assign(font);

      font.name   := 'Arial';
      font.Size   := 12;
      font.Style  := [];
      font.height := asg_DialMap.mapfontheight(12);
      font.Color  := clBlack;

      ts := asg_DialMap.Printcoloffset[0];
      tw := asg_DialMap.Printpagewidth;
      ts := ts+ ((tw-textwidth('[' + FormatDateTime('yyyy-mm-dd', Date) + ' Printed]')) shr 1);

      Textout(ts, -135, '[' + FormatDateTime('yyyy-mm-dd', Date) + ' Printed]');

      font.assign(savefont);

      savefont.free;
   end;


   // 병원별 Fax. Info Display
   with canvas do
   begin
      savefont := TFont.Create;
      savefont.assign(font);

      font.name   := 'Arial';
      font.Size   := 12;
      font.Style  := [];
      font.height := asg_DialMap.mapfontheight(12);
      font.Color  := clBlack;

      ts := asg_DialMap.Printcoloffset[0];
      tw := asg_DialMap.Printpagewidth;
      ts := ts+ ((tw-textwidth(addInfo)) shr 1);

      Textout(ts + 300, -200, addInfo);

      font.assign(savefont);

      savefont.free;
   end;

end;

procedure TMainDlg.asg_DialMapPrintStart(Sender: TObject;
  NrOfPages: Integer; var FromPage, ToPage: Integer);
begin
   Printdialog1.FromPage   := Frompage;
   Printdialog1.ToPage     := ToPage;
   Printdialog1.Maxpage    := ToPage;
   Printdialog1.minpage    := 1;

   if Printdialog1.execute then
   begin
      Frompage := Printdialog1.FromPage;
      Topage   := Printdialog1.ToPage;
   end
   else
   begin
      Frompage := 0;
      Topage   := 0;
   end;
end;

procedure TMainDlg.asg_DialMapPrintSetColumnWidth(Sender: TObject;
  ACol: Integer; var Width: Integer);
begin
   Width := 200;
end;

procedure TMainDlg.asg_DialMapPrintSetRowHeight(Sender: TObject;
  ARow: Integer; var Height: Integer);
begin
   Height := 200;
end;

procedure TMainDlg.asg_DialListMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
   NowCol, NowRow : integer;
begin
   //---------------------------------------------------------------
   // 마우스 좌클릭(on)시 row-index 가져온다
   //---------------------------------------------------------------
   if (Button = mbLeft) then
   begin
      iSelRow1 := 0;

      with asg_DialList do
      begin
         // Mouse 포인터 위치 받아오기
         MouseToCell(X, Y, NowCol, NowRow);

         iSelRow1 := NowRow;
      end
   end;

end;

procedure TMainDlg.asg_DialListMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
   NowCol, NowRow : integer;
begin
   //---------------------------------------------------------------
   // 마우스 좌클릭(off)시 row-index 가져온다
   //---------------------------------------------------------------
   if (Button = mbLeft) then
   begin
      iSelRow2 := 0;

      with asg_DialList do
      begin
         // Mouse 포인터 위치 받아오기
         MouseToCell(X, Y, NowCol, NowRow);

         iSelRow2 := NowRow;
      end
   end;

end;

procedure TMainDlg.asg_MyDialMouseUp(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
var
   NowCol, NowRow : integer;
begin
   //---------------------------------------------------------------
   // 마우스 좌클릭(off)시 row-index 가져온다
   //---------------------------------------------------------------
   if (Button = mbLeft) then
   begin
      iSelRow2 := 0;

      with asg_MyDial do
      begin
         // Mouse 포인터 위치 받아오기
         MouseToCell(X, Y, NowCol, NowRow);

         iSelRow2 := NowRow;
      end
   end;

end;

procedure TMainDlg.asg_MyDialMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
   NowCol, NowRow : integer;
begin
   //---------------------------------------------------------------
   // 마우스 좌클릭(on)시 row-index 가져온다
   //---------------------------------------------------------------
   if (Button = mbLeft) then
   begin
      iSelRow1 := 0;

      with asg_MyDial do
      begin
         // Mouse 포인터 위치 받아오기
         MouseToCell(X, Y, NowCol, NowRow);

         iSelRow1 := NowRow;
      end
   end;

end;

procedure TMainDlg.FormPaint(Sender: TObject);
//var
//  DC: hDC;
//  Soo: TBitmap;
begin
{     -- test
    DC := GetWindowDC(Handle);
    Soo := TBitmap.Create;
    Soo.LoadFromFile('c:\140508.bmp');
    BitBlt(DC, 0, 0, 686, 603, Soo.Canvas.Handle, 10, 20, SRCCOPY);

    ReleaseDC(Handle, DC);
    Soo.Free;
}
end;



//------------------------------------------------------------------------------
// AdvStringGrid onMouseMove Event Handler
//       - 다이얼Book 검색된 User 프로필 확인
//
// Date   : 2014.06.26
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_DialListMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
var
   NowCol, NowRow : Integer;
   ServerFileName, ClientFileName: String;
   FTP_SVRIP, FTP_USERID, FTP_PASSWD, FTP_HOSTNAME, FTP_DIR : String;
begin
   with asg_DialList do
   begin
      // 현 Mouse 포인터 좌표 받아오기
      MouseToCell(X,
                  Y,
                  NowCol,
                  NowRow);


      // 다이얼 Book 담당자 프로필 조회
      if (NowCol = C_DL_DUTYUSER) and
         (asg_DialList.Cells[C_DL_DUTYUSER, NowRow] <> '') and
         (IsLogonUser)  then
      begin

         if (NowRow > 0) then
         begin

            //lb_DialScan.Caption := '▶ ' + Trim(asg_DialList.Cells[C_DL_DUTYUSER, NowRow]) + '님의 담당자 프로필 조회중.';

            if apn_UserProfile.Visible = False then
            begin
               // Location & Display
               apn_UserProfile.Top        := asg_DialList.CellRect(C_DL_LINKCNT, NowRow).Top + 53;
               apn_UserProfile.Left       := 212;

               apn_UserProfile.Caption.Color := $0046D9F3;
               asg_UserProfile.FixedColor    := $0046D9F3;
               apn_UserProfile.Collaps       := True;
               apn_UserProfile.Visible       := True;
               apn_UserProfile.Collaps       := False;
            end;



            // 프로필 내역 assign
            asg_UserProfile.Cells[1, R_PR_USERNM]     := Cells[C_DL_DUTYUSER, NowRow];
            asg_UserProfile.Cells[1, R_PR_DUTYPART]   := Cells[C_DL_LOCATE,   NowRow]  + ' ' +
                                                         Cells[C_DL_DEPTNM,   NowRow];

            asg_UserProfile.Cells[1, R_PR_REMARK]     := Cells[C_DL_DEPTSPEC, NowRow];
            asg_UserProfile.Cells[1, R_PR_CALLNO]     := Cells[C_DL_CALLNO,   NowRow];

            asg_UserProfile.AddButton(1, R_PR_DUTYSCH, asg_UserProfile.ColWidths[0]+45, 20, 'S/R 요청현황(1년)', haBeforeText, vaCenter);
            asg_UserProfile.Cells[1, R_PR_USERID]     := Cells[C_DL_DTYUSRID,   NowRow];

            // 이미지 Stream 초기화 (JPEG Error #41 방지위함) @ 2015.06.03 LSH
            asg_UserProfile.RemovePicture(1, R_PR_PHOTO);

            // 프로필 이미지 FTP 파일내역이 존재하면, 이미지 Download
            if (Cells[C_DL_DTYUSRID, NowRow] <> '') then
            begin
               // 담당자 프로필 이미지 파일명
               if Cells[C_DL_DTYUSRID, NowRow] = 'XXXXX' then
               begin
                  // 프로필 이미지 Remove
                  asg_UserProfile.RemovePicture(1, R_PR_PHOTO);

                  // 이미지 등록 버튼 생성 --> 다이얼Book 에서는 프로필 이미지 등록 제한
                  //asg_UserProfile.AddButton(1, R_PR_PHOTO, asg_UserProfile.ColWidths[0]+45, asg_UserProfile.RowHeights[R_PR_PHOTO]-5, '프로필 이미지 등록', haBeforeText, vaCenter);

                  Exit;
               end
               else
                  // 담당자 HRM 프로필 이미지 파일명
                  ServerFileName := Cells[C_DL_DTYUSRID, NowRow] + '.jpg';


                  // 사진 이미지가 없는경우 Skip (JPEG ERROR #41 방지) @ 2015.06.03 LSH
                  if AnsiUpperCase(ServerFileName) = '.JPG' then
                     Exit;


                  // Local 파일이름 Set
                  // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
                  ClientFileName := G_HOMEDIR + 'TEMP\SPOOL\' + ServerFileName;


                  // Local에 해당 Image 존재유무 체크
                  // Local file의 Size 체크 추가 (Local에 해당 사번.jpg가 0 size로 존재하면 JPEG Error #41 떨어짐) @ 2015.06.03 LSH
                  if (FileExists(ClientFileName)) and
                     (CMsg.GetFileSize(ClientFileName) >  0) then
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
                     FTP_USERID := '';
                     FTP_PASSWD := '';
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

                           Exit;
                        end;

                     except
                        Exit;       // 아직 FTP 서버에 사번.jpg 이미지 미등록된 경우 try~except 거쳐서 Exit
                     end;



                     // 해당 Image Variant로 받음.
                     AppendVariant(DownFile, ClientFileName);


                     // 다운로드 횟수 체크
                     DownCnt := DownCnt + 1;
                  end;


                  // 수신한 Image 파일을 Grid에 표기
                  try
                     asg_UserProfile.CreatePicture(1,
                                                   R_PR_PHOTO,
                                                   False,
                                                   ShrinkWithAspectRatio,
                                                   0,
                                                   haLeft,
                                                   vaTop).LoadFromFile(ClientFileName);
                  except
                     Exit;
                  end;
            end;
         end;
      end
      else
      begin
         //lb_DialScan.Caption := '';

         apn_UserProfile.Collaps := True;
         apn_UserProfile.Visible := False;
      end;
   end;
end;



//------------------------------------------------------------------------------
// AdvStringGrid onMouseMove Event Handler
//       - S/R 요청 User 프로필 확인
//
// Date   : 2014.06.27
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_AnalysisMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
var
   NowCol, NowRow : Integer;
   ServerFileName, ClientFileName: String;
   FTP_SVRIP, FTP_USERID, FTP_PASSWD, FTP_HOSTNAME, FTP_DIR : String;
begin
   with asg_Analysis do
   begin
      // 현 Mouse 포인터 좌표 받아오기
      MouseToCell(X,
                  Y,
                  NowCol,
                  NowRow);


      // 담당자 프로필 조회
      if (NowCol = C_AN_REQUSER) and
         (Cells[C_AN_REQUSER, NowRow] <> '') and
         (IsLogonUser) then
      begin
         if (NowRow > 0) then
         begin

            lb_Analysis.Caption := '▶ ' + Trim(Cells[C_AN_REQUSER, NowRow]) + '님의 요청자 프로필 조회중.';

            if apn_UserProfile.Visible = False then
            begin
               // Location & Display
               if (Height - CellRect(C_AN_REQUSER, NowRow).Top) > (apn_UserProfile.Height + 195) then
               begin
                  apn_UserProfile.Top     := CellRect(C_AN_REQUSER, NowRow).Top + 102;
                  apn_UserProfile.Left    := 147;
               end
               else
               begin
                  apn_UserProfile.Top     := CellRect(C_AN_REQUSER, NowRow).Top - apn_UserProfile.Height - 150;
                  apn_UserProfile.Left    := 147;
               end;

               apn_UserProfile.Caption.Color := $00FB97AD;
               asg_UserProfile.FixedColor    := $00FB97AD;
               apn_UserProfile.Collaps       := True;
               apn_UserProfile.Visible       := True;
               apn_UserProfile.Collaps       := False;
            end;

            // 프로필 내역 assign
            asg_UserProfile.Cells[1, R_PR_USERNM]     := Cells[C_AN_REQUSER,  NowRow];
            asg_UserProfile.Cells[1, R_PR_DUTYPART]   := Cells[C_AN_REQDEPTNM,NowRow];
            asg_UserProfile.Cells[1, R_PR_REMARK]     := Cells[C_AN_REQSPEC,  NowRow];
            asg_UserProfile.Cells[1, R_PR_CALLNO]     := Cells[C_AN_REQTELNO, NowRow];
            asg_UserProfile.Cells[1, R_PR_USERID]     := Cells[C_AN_USERID,   NowRow];
            asg_UserProfile.AddButton(1, R_PR_DUTYSCH, asg_UserProfile.ColWidths[0]+45, 20, 'S/R 요청현황(1년)', haBeforeText, vaCenter);


            // 프로필 이미지 FTP 파일내역이 존재하면, 이미지 Download
            if (Cells[C_AN_USERID, NowRow] <> '') then
            begin
               // 담당자 프로필 이미지 파일명
               if Cells[C_AN_USERID, NowRow] = 'XXXXX' then
               begin
                  // 프로필 이미지 Remove
                  asg_UserProfile.RemovePicture(1, R_PR_PHOTO);

                  // 이미지 등록 버튼 생성 --> 다이얼Book 에서는 프로필 이미지 등록 제한
                  //asg_UserProfile.AddButton(1, R_PR_PHOTO, asg_UserProfile.ColWidths[0]+45, asg_UserProfile.RowHeights[R_PR_PHOTO]-5, '프로필 이미지 등록', haBeforeText, vaCenter);

                  Exit;
               end
               else
                  // 담당자 HRM 프로필 이미지 파일명
                  ServerFileName := Cells[C_AN_USERID, NowRow] + '.jpg';

                  // Local 파일이름 Set
                  // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
                  ClientFileName := G_HOMEDIR + 'TEMP\SPOOL\' + ServerFileName;


                  // Local에 해당 Image 존재유무 체크
                  // Local file의 Size 체크 추가 (Local에 해당 사번.jpg가 0 size로 존재하면 JPEG Error #41 떨어짐) @ 2015.06.03 LSH
                  if (FileExists(ClientFileName)) and
                     (CMsg.GetFileSize(ClientFileName) >  0) then
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
                     FTP_USERID := '';
                     FTP_PASSWD := '';
                     FTP_DIR    := '/kuhrm/app/hrm/photo/';


                     // HRM Image 다운로드
                     if Not GetBINFTP(FTP_SVRIP,
                                      FTP_USERID,
                                      FTP_PASSWD,
                                      FTP_DIR + ServerFileName,
                                      ClientFileName,
                                      False) then
                     begin
                        //Showmessage('이미지 다운로드에 실패했습니다. 다시한번 시도해 주십시오');

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
                     asg_UserProfile.CreatePicture(1,
                                                   R_PR_PHOTO,
                                                   False,
                                                   ShrinkWithAspectRatio,
                                                   0,
                                                   haLeft,
                                                   vaTop).LoadFromFile(ClientFileName);
                  except
                     Exit;
                  end;
            end;
         end;
      end
      else
      begin
         // 이전 Action 메세지 보존 위해 주석 @ 2016.11.18 LSH
         //lb_Analysis.Caption := '';

         apn_UserProfile.Collaps := True;
         apn_UserProfile.Visible := False;
      end;
   end;
end;




//------------------------------------------------------------------------------
// [HTML팝업] 다이얼 Pop 메세지 실시간 Pop-up
//       - 구분(in_Gubun)별 Pop-up 메세지 Set.
//
// Date   : 2014.07.02
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
function TMainDlg.CreatePopup(in_Gubun : String) : Boolean;
begin

   Result := False;

   if (in_Gubun = 'LOGIN') then
   begin
      Result := True;

      // 중복 인스턴스 생성(메모리 누수) 방지 @ 2018.07.16 LSH
      if (HTMLPopup = nil) then
         HTMLPopup := THtmlPopup.Create(self);

      htmlpopup.AlwaysOnTop      := True;
      htmlpopup.AutoSize         := True;
      htmlpopup.PopupWidth       := 200;
      htmlpopup.PopupHeight      := 200;
      htmlpopup.ShadeEnable      := True;
      htmlpopup.ShadeSteps       := 50;
      htmlpopup.ShadeStartColor  := clWhite;
      htmlpopup.ShadeEndColor    := $0076BBF3;
      htmlpopup.BorderSize       := 5;
      htmlpopup.PopupTop         := 0;
      htmlpopup.PopupLeft        := 200;
      htmlpopup.Font.Name        := 'Tahoma';
      htmlpopup.Font.Style       := [fsBold];
      htmlpopup.OnAnchorClick    := AnchorClick;

      htmlpopup.Text.Add('<font color="clMaroon"><font size = "9"> :: 다이얼 POP 리얼타임 메세지 :: </font><BR><BR>');

      FsPopMsg := FormatDateTime('yyyy-mm-dd hh:nn:ss', Now);

      htmlpopup.Text.Add('<font color="clBlack">[로그인] ');
      htmlpopup.Text.Add(FsUserSpec + ' ' + FsUserNm + '님');
      htmlpopup.Text.Add(' </font><BR><BR><font color = "clMaroon">');
      htmlpopup.Text.Add(FsPopMsg);
      htmlpopup.Text.Add('</font><BR><BR><font color="clGray">Press here to <a href="Close">닫기</a></font>');

      htmlpopup.RollUp;
   end
   else if (in_Gubun = 'DIET') then
   begin
      Result := True;

      // 식단알림은, 최초 30초동안 Pop-up
      tm_DialPop.Interval := 30000;

      // 중복 인스턴스 생성(메모리 누수) 방지 @ 2018.07.16 LSH
      if (HTMLPopup = nil) then
         HTMLPopup := THtmlPopup.Create(self);

      htmlpopup.AlwaysOnTop      := True;
      htmlpopup.AutoSize         := True;
      htmlpopup.ShadeEnable      := True;
      htmlpopup.ShadeSteps       := 50;
      htmlpopup.ShadeStartColor  := clWhite;
      htmlpopup.ShadeEndColor    := $0076BB9D;
      htmlpopup.BorderSize       := 5;
      htmlpopup.PopupTop         := 0;
      htmlpopup.PopupLeft        := 200;
      htmlpopup.Font.Name        := 'Tahoma';
      htmlpopup.Font.Style       := [fsBold];
      htmlpopup.OnAnchorClick    := AnchorClick;

      htmlpopup.Text.Add('<font color="clMaroon"><font size = "9"> :: 다이얼 POP 리얼타임 메세지 :: </font><BR><BR>');
      htmlpopup.Text.Add('<font color="clBlack">[식단알림] ');
      htmlpopup.Text.Add(' </font><BR><BR><font color = "clMaroon">');
      htmlpopup.Text.Add(FsPopMsg);
      htmlpopup.Text.Add('</font><BR><BR><font color="clGray">Press here to <a href="Close">닫기</a></font>');

      // 식이끼니별 Anchor 분기
      if (PosByte('조식', FsPopMsg) > 0) or
         (PosByte('아침', FsPopMsg) > 0) then
         htmlpopup.Text.Add('<BR><BR><font color="clGray">※ 이 팝업 그만보기 ☞ <a href="Diet1Off">확인</a></font>')
      else if (PosByte('중식', FsPopMsg) > 0) or
              (PosByte('점심', FsPopMsg) > 0) then
         htmlpopup.Text.Add('<BR><BR><font color="clGray">※ 이 팝업 그만보기 ☞ <a href="Diet2Off">확인</a></font>')
      else if (PosByte('석식', FsPopMsg) > 0) or
              (PosByte('저녁', FsPopMsg) > 0) then
         htmlpopup.Text.Add('<BR><BR><font color="clGray">※ 이 팝업 그만보기 ☞ <a href="Diet3Off">확인</a></font>');

      htmlpopup.RollUp;
   end
   else if (in_Gubun = 'RTCHECK') then
   begin
      Result := True;

      // 실시간 알람체크는, 최초 10초동안 Pop-up
      tm_DialPop.Interval := 10000;

      // 중복 인스턴스 생성(메모리 누수) 방지 @ 2018.07.16 LSH
      if (HTMLPopup = nil) then
         HTMLPopup := THtmlPopup.Create(self);

      htmlpopup.AlwaysOnTop      := True;
      htmlpopup.AutoSize         := True;
      htmlpopup.PopupWidth       := 200;
      htmlpopup.PopupHeight      := 200;
      htmlpopup.ShadeEnable      := True;
      htmlpopup.ShadeSteps       := 50;
      htmlpopup.ShadeStartColor  := clWhite;
      htmlpopup.ShadeEndColor    := $00FBBB9D;
      htmlpopup.BorderSize       := 5;
      htmlpopup.PopupTop         := 0;
      htmlpopup.PopupLeft        := 200;
      htmlpopup.Font.Name        := 'Tahoma';
      htmlpopup.Font.Style       := [fsBold];
      htmlpopup.OnAnchorClick    := AnchorClick;

      htmlpopup.Text.Add('<font color="clPurple"><font size = "9"> :: 다이얼 POP 리얼타임 메세지 :: </font><BR><BR>');
      htmlpopup.Text.Add('<font color = "clMaroon">');
      htmlpopup.Text.Add(FsPopMsg);
      htmlpopup.Text.Add('</font>');
      htmlpopup.Text.Add('<BR><BR><font color="clGray">Press here to <a href="Close">닫기</a></font>');

      htmlpopup.RollUp;
   end
   else if (in_Gubun = 'NOPOP') then
      Exit;

end;


//------------------------------------------------------------------------------
// Form onCreate Event Handler
//
// Date   : 2014.07.02
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.FormCreate(Sender: TObject);
var
   sRootPath : String;
begin
   // Memory 누수 체크 적용시 주석풀것 @ 2015.05.12 LSH
   //memchk;

   //
   // 현 App. 실행 Root 디렉토리
   sRootPath := UpperCase(TokenStr(ExtractFileDir(ParamStr(0)), '\', 1)) +
                '\'                                                      +
                UpperCase(TokenStr(ExtractFileDir(ParamStr(0)), '\', 2)) +
                '\';

   // App. 실행경로 잡아주기 (추가)
   G_HOMEDIR := Format('%sEXE\', [sRootPath]);

//
//   // test
//   showmessage(G_HomeDir);
//
//   //
//   G_ENV_FILENAME := G_HomeDir + 'ENV\TMAX.ENV';

   // Form 생성시, HTML Pop-up 초기화
   HTMLPopup := nil;

   // 이미지(프로필, 자료 등) 관련 Temp 폴더 필수 생성
   ForceDirectories(G_HOMEDIR + 'Temp\SPOOL');

end;


//------------------------------------------------------------------------------
// [HTML팝업] Pop-up Anchor onClick Event Handler
//       - 구분(Anchor)별 Close Action 분기
//
// Date   : 2014.07.02
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.AnchorClick(sender: TObject; Anchor: string);
var
   varResult : String;
begin
   // 일반 닫기
   if Anchor = 'Close' then
   begin
      htmlpopup.RollDown;
      htmlpopup.Free;
      htmlpopup := nil;

      // 일반 알람off후, 체크 Interval 10초로 Reset.
      tm_DialPop.Interval := 10000;
   end
   // 식단정보
   else if PosByte('Diet', Anchor) > 0 then
   begin
      htmlpopup.RollDown;
      htmlpopup.Free;
      htmlpopup := nil;

      // 식단끼니별 Pop-up 확인 Log update
      UpdateLog('POPUP',
                Anchor,
                FsUserIp,
                'C',
                FormatDateTime('yyyy-mm-dd', Date),
                FsUserNm,
                varResult
                );

      // 식단알람 off후, 체크 Interval 10초로 Reset.
      tm_DialPop.Interval := 10000;
   end;

   //  메모리 누수 의심부분 주석 @ 2018.07.16 LSH
//   else
//      ShellExecute(0,
//                   'open',
//                   Pchar(Anchor),
//                   nil,
//                   nil,
//                   SW_NORMAL);
end;



procedure TMainDlg.HTMLPopupAnchorClick(Sender: TObject; Anchor: String);
begin
   ShellExecute(0,
                'open',
                pchar(Anchor),
                nil,
                nil,
                SW_NORMAL);

   htmlpopup.RollDown;

   htmlpopup.Free;
   htmlpopup := nil;
end;


//------------------------------------------------------------------------------
// [타이머] 다이얼 Pop onTimer Event Handler
//
// Date   : 2014.07.02
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.tm_DialPopTimer(Sender: TObject);
begin
   // 다이얼 Pop-up 미확인시, 10초마다 자동 Roll-Down
   if Assigned(htmlpopup) then
   begin
      htmlpopup.RollDown;
      htmlpopup.Free;
      htmlpopup := nil;
   end;

   //--------------------------------------------------------------------
   // 개발계 Tmax 연결 확인시에만, 실시간 Pop-up 생성
   //--------------------------------------------------------------------
   if (TxInit(TokenStr(String(CONST_ENV_FILENAME), '.', 1) + '_D0.ENV', '01')) then
   begin
      // 다이얼 Pop 신규 메세지 유무 Check
      if CreatePopup(CheckDialPop('REALTIME')) then
      begin
         //
      end
      else
         Exit;
   end
   else
      MessageBox(self.Handle,
                 PChar('현재 Server와의 연결이 끊겼습니다.' + #13#10 + #13#10 + '90초마다 자동으로 연결을 복구하오니, 잠시만 기다려 주십시오.'),
                 PChar(Self.Caption + ' : TMAX 접속끊김 알림'),
                 MB_OK + MB_ICONWARNING);

end;



//------------------------------------------------------------------------------
// [함수] 다이얼 Pop 메세지 실시간 Check
//
// Date   : 2014.07.02
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
function TMainDlg.CheckDialPop(in_Gubun : String) : String;
var
   TpGetPopUp : TTpSvc;
   sType1, sType2, sType3, sType4, sType5, sType6 : String;
   i, iRowCnt : Integer;
begin
   // Init. Return Values
   Result := 'NOPOP';


   //-------------------------------------------------------------
   // 1. 로그인 팝업 설정
   //-------------------------------------------------------------
   if (in_Gubun = 'LOGIN') then
   begin
      Result := in_Gubun;

      Exit;
   end;


   //-------------------------------------------------------------
   // 2. 시간대별 맞춤 끼니 정보 조회
   //-------------------------------------------------------------
   if ((FormatDateTime('hh:nn', Now) > '07:00') and (FormatDateTime('hh:nn', Now) <= '08:30') and
       (not CheckUserAlarm('Diet1Off',
                           FsUserNm,
                           FsUserIp,
                           'POPUP'))) or
      ((FormatDateTime('hh:nn', Now) > '11:00') and (FormatDateTime('hh:nn', Now) <= '14:00') and
       (not CheckUserAlarm('Diet2Off',
                           FsUserNm,
                           FsUserIp,
                           'POPUP'))) or
      ((FormatDateTime('hh:nn', Now) > '17:00') and (FormatDateTime('hh:nn', Now) <= '19:00') and
       (not CheckUserAlarm('Diet3Off',
                           FsUserNm,
                           FsUserIp,
                           'POPUP'))) then
   begin
      FsPopMsg := GetDietInfo(FormatDateTime('dd', Date),
                              GetDayofWeek(Date, 'HS'));

      Result   := 'DIET';


      Exit;
   end;


   //-----------------------------------------------------------------
   // 3. Set Params.
   //-----------------------------------------------------------------
   sType1 := '37';
   sType2 := '';
   sType3 := '';
   sType4 := '';
   sType5 := '';


   //-----------------------------------------------------------------
   // 4-1. 조회
   //-----------------------------------------------------------------
   Screen.Cursor := crHourGlass;





   //-----------------------------------------------------------------
   // 4-2. Load Data by TpSvc
   //-----------------------------------------------------------------
   TpGetPopUp := TTpSvc.Create;
   TpGetPopUp.Init(Self);


   Screen.Cursor := crHourglass;


   try
      TpGetPopUp.CountField  := 'S_CODE3';
      TpGetPopUp.ShowMsgFlag := False;

      if TpGetPopUp.GetSvc('MD_KUMCM_S1',
                          ['S_TYPE1'  , sType1
                         , 'S_TYPE2'  , sType2
                         , 'S_TYPE3'  , sType3
                         , 'S_TYPE4'  , sType4
                         , 'S_TYPE5'  , sType5
                         , 'S_TYPE6'  , sType6
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
                          ]) then

         if TpGetPopUp.RowCount <= 0 then
         begin
            Exit;
         end;


      // 구분 Flag 별 Return Value 분기
      if (in_Gubun = 'REALTIME') then
      begin
         iRowCnt := TpGetPopUp.RowCount;

         Result := 'RTCHECK';

         FsPopMsg := '';

         if iRowCnt > 1 then
         begin
            for i := 0 to iRowCnt - 1 do
            begin
               FsPopMsg := FsPopMsg +  TpGetPopUp.GetOutputDataS('sNickNm',  i) + #13#10 + #13#10 +
                                       TpGetPopUp.GetOutputDataS('sLocate',  i) + ' ' +
                                       TpGetPopUp.GetOutputDataS('sDutyPart',i) + ' ' +
                                       '[' + TpGetPopUp.GetOutputDataS('sUserNm',  i) + ']' + #13#10 +
                                       TpGetPopUp.GetOutputDataS('sDutyRmk', i) + #13#10 + #13#10;
            end;
         end
         else
            FsPopMsg :=    TpGetPopUp.GetOutputDataS('sNickNm',  0) + #13#10 + #13#10 +
                           TpGetPopUp.GetOutputDataS('sLocate',  0) + ' ' +
                           TpGetPopUp.GetOutputDataS('sDutyPart',0) + ' ' +
                           '[' + TpGetPopUp.GetOutputDataS('sUserNm',  0) + ']'+ #13#10 +
                           TpGetPopUp.GetOutputDataS('sDutyRmk', 0) + #13#10 + #13#10;
      end;


   finally
      FreeAndNil(TpGetPopUp);
      Screen.Cursor := crDefault;
   end;
end;


procedure TMainDlg.asg_ReleaseClickCell(Sender: TObject; ARow,
  ACol: Integer);
begin
   if (ACol = C_RL_SRREQNO) then
   begin
      if (Trim(asg_Release.Cells[C_RL_SRREQNO, ARow]) <> '') then
         lb_RegDoc.Caption := '▶ S/R 번호를 더블클릭하시면, 상세 요청내역을 보실 수 있습니다.'
      else
         lb_RegDoc.Caption := '';
   end;

end;

procedure TMainDlg.asg_SelDocDblClick(Sender: TObject);
begin
   with asg_SelDoc do
   begin
      if PosByte('계약', Cells[C_SD_DOCLIST, Row]) > 0 then
      begin
         asg_DocSpec.Cells[0, R_DS_USERID]   := '계약명';
         asg_DocSpec.Cells[0, R_DS_USERNM]   := '계약업체명';
         asg_DocSpec.Cells[0, R_DS_DEPTCD]   := '계약기간';
         asg_DocSpec.Cells[0, R_DS_DEPTNM]   := '';
         asg_DocSpec.Cells[0, R_DS_SOCNO]    := '분담액(총액)';
         asg_DocSpec.Cells[0, R_DS_STARTDT]  := '분담액(HQ)';
         asg_DocSpec.Cells[0, R_DS_ENDDT]    := '분담액(AA)';
         asg_DocSpec.Cells[0, R_DS_MOBILE]   := '분담액(GR)';
         asg_DocSpec.Cells[0, R_DS_GRPID]    := '분담액(AS)';
         asg_DocSpec.Cells[0, R_DS_USELVL]   := '문서위치';

         if (FontColors[C_SD_DOCTITLE,  Row] = clRed) then
         begin
            asg_DocSpec.Cells[1, R_DS_USERID]      := Cells[C_SD_DOCTITLE,Row];
            asg_DocSpec.FontColors[1, R_DS_USERID] := clRed;

            asg_DocSpec.Cells[1, R_DS_DEPTCD]      := Cells[C_SD_DOCRMK,Row] + ' [계약종료]';
            asg_DocSpec.FontColors[1, R_DS_DEPTCD] := clRed;
         end
         else
         begin
            asg_DocSpec.Cells[1, R_DS_USERID]   := Cells[C_SD_DOCTITLE,Row];
            asg_DocSpec.FontColors[1, R_DS_USERID] := clBlack;

            asg_DocSpec.Cells[1, R_DS_DEPTCD]   := Cells[C_SD_DOCRMK,Row];
            asg_DocSpec.FontColors[1, R_DS_DEPTCD] := clBlack;
         end;

         asg_DocSpec.Cells[1, R_DS_USERNM]   := Cells[C_SD_RELDEPT, Row];
         asg_DocSpec.Cells[1, R_DS_DEPTNM]   := '';
         asg_DocSpec.Cells[1, R_DS_SOCNO]    := Cells[C_SD_TOTRMK, Row];
         asg_DocSpec.Cells[1, R_DS_STARTDT]  := Cells[C_SD_HQREMARK, Row];
         asg_DocSpec.Cells[1, R_DS_ENDDT]    := Cells[C_SD_AAREMARK, Row];
         asg_DocSpec.Cells[1, R_DS_MOBILE]   := Cells[C_SD_GRREMARK, Row];
         asg_DocSpec.Cells[1, R_DS_GRPID]    := Cells[C_SD_ASREMARK, Row];
         asg_DocSpec.Cells[1, R_DS_USELVL]   := Cells[C_SD_DOCLOC, Row];

         apn_DocSpec.Top     := 155;
         apn_DocSpec.Left    := 140;
         apn_DocSpec.Collaps := True;
         apn_DocSpec.Visible := True;
         apn_DocSpec.Collaps := False;

         apn_DocSpec.Caption.Text := Cells[C_SD_DOCLIST, Row] + ' 유지보수 상세정보 조회';
      end;
   end;
end;



procedure TMainDlg.asg_SelDocClick(Sender: TObject);
begin
   if (apn_DocSpec.Visible) and
      (PosByte('조회', apn_DocSpec.Caption.Text) > 0) then
   begin
      apn_DocSpec.Caption.Text := '';
      apn_DocSpec.Collaps      := True;
      apn_DocSpec.Visible      := False;
   end;
end;


procedure TMainDlg.asg_MultiInsClickCell(Sender: TObject; ARow,
  ACol: Integer);
begin
   with asg_MultiIns do
   begin
      // [업무일지 작성내역] Click시
      if (ACol = 1) and
         (ARow = 2) then
      begin
         // TMemo 컴포넌트를 Cell Grid내 Object로 생성
         Objects[ACol,ARow] := TMemo.Create(Application);

         // TMemo 컴포넌트 상세 속성 정의
         with TMemo(Objects[ACol,ARow]) do
         begin
            Parent     := Self;
            Left       := apn_MultiIns.Left + CellRect(ACol,ARow).Left + 1;
            Top        := apn_MultiIns.Top + CellRect(ACol,ARow).Top   + 1;
            Width      := asg_MultiIns.CellRect(ACol,ARow).Right  - asg_MultiIns.CellRect(ACol,ARow).Left;
            Height     := CellRect(ACol,ARow).Bottom - CellRect(ACol,ARow).Top + 2;

            // 기존 작성내역 있으면, TMemo 값으로 가져온다
            if (Cells[1,2] <> '') then
               Text := Cells[1,2];

            SetFocus;
         end;
      end
      else
      begin
         // 기존 셀의 TMemo는 삭제.
         if Objects[1,2] <> nil Then
         begin
            Cells[1,2] := TMemo(Objects[1,2]).Text;

            Objects[1,2].Free;
            Objects[1,2] := Nil;
         end;
      end;
   end;
end;


procedure TMainDlg.asg_MultiInsTopLeftChanged(Sender: TObject);
begin
   with asg_MultiIns do
   begin
      // 기존 셀의 TMemo는 삭제.
      if Objects[1,2] <> nil Then
      begin
         Objects[1,2].Free;
         Objects[1,2] := Nil;
      end;
   end;
end;


procedure TMainDlg.asg_MultiInsDblClick(Sender: TObject);
begin
   // 업무일지 작성(TMemo 동적 컴포넌트) Panel Close @ 2014.07.31 LSH
   SetPanelStatus('WORKRPT', 'OFF');
end;

procedure TMainDlg.fsbt_ConnWriteClick(Sender: TObject);
begin
   if (PosByte('실장',   FsUserSpec) > 0) or
      (PosByte('팀장',   FsUserSpec) > 0) or
      (PosByte('파트장', FsUserSpec) > 0) or
      (PosByte('P/L',    FsUserSpec) > 0) then
   begin
      // 업무공유 등록 Panel-On
      SetPanelStatus('WORKCONN', 'ON');
   end
   else
      MessageBox(self.Handle,
                 PChar('각 병원 관리자 및 파트 P/L 이상만 작성 가능합니다.'),
                 PChar('업무공유 작성권한 제한알림'),
                 MB_OK + MB_ICONWARNING);
end;

procedure TMainDlg.asg_WorkConnGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   if (ARow = 0) or
      ((ARow > 0) and (ACol <> C_WC_CONTEXT)) then
   HAlign := taCenter;

   VAlign := vtaCenter;
end;

procedure TMainDlg.fmed_ConnFromChange(Sender: TObject);
begin
   if (IsLogonUser) then
      if (ftc_Dialog.ActiveTab = AT_WORKCONN) and
         (
            (
               (CopyByte(Trim(fmed_ConnFrom.Text), 1, 1)  <> '') and
               (CopyByte(Trim(fmed_ConnFrom.Text), 2, 1)  <> '') and
               (CopyByte(Trim(fmed_ConnFrom.Text), 3, 1)  <> '') and
               (CopyByte(Trim(fmed_ConnFrom.Text), 4, 1)  <> '') and
               (CopyByte(Trim(fmed_ConnFrom.Text), 6, 1)  <> '') and
               (CopyByte(Trim(fmed_ConnFrom.Text), 7, 1)  <> '') and
               (CopyByte(Trim(fmed_ConnFrom.Text), 9, 1)  <> '') and
               (CopyByte(Trim(fmed_ConnFrom.Text), 10, 1) <> '')
            )
          )
          then
         SelGridInfo('WORKCONN');
end;

procedure TMainDlg.fmed_ConnToChange(Sender: TObject);
begin

   if (IsLogonUser) then
      if (ftc_Dialog.ActiveTab = AT_WORKCONN) and
         (
            (
               (CopyByte(Trim(fmed_ConnTo.Text), 1, 1)  <> '') and
               (CopyByte(Trim(fmed_ConnTo.Text), 2, 1)  <> '') and
               (CopyByte(Trim(fmed_ConnTo.Text), 3, 1)  <> '') and
               (CopyByte(Trim(fmed_ConnTo.Text), 4, 1)  <> '') and
               (CopyByte(Trim(fmed_ConnTo.Text), 6, 1)  <> '') and
               (CopyByte(Trim(fmed_ConnTo.Text), 7, 1)  <> '') and
               (CopyByte(Trim(fmed_ConnTo.Text), 9, 1)  <> '') and
               (CopyByte(Trim(fmed_ConnTo.Text), 10, 1) <> '')
            )
          )
          then
         SelGridInfo('WORKCONN');
end;

procedure TMainDlg.mi_DeleteClick(Sender: TObject);
begin
   if asg_WorkConn.Cells[C_WC_DUTYUSER, asg_WorkConn.Row] <> FsUserNm then
   begin
      MessageBox(self.Handle,
                 PChar('본인이 작성한 내역만 삭제가 가능합니다.'),
                 PChar('주요업무공유 삭제전 본인 확인'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;

   // 삭제 Confirm Message.
   if Application.MessageBox('선택한 게시글을 [삭제]하시겠습니까?',
                             PChar('주요업무공유 삭제 알림'), MB_OKCANCEL) <> IDOK then
      Exit;

   // 주요업무공유 내역 삭제 Update
   UpdateImage('WORKCONNDEL',
               FormatDateTime('yyyy-mm-dd', Date),
               '',
               1);
end;

// [팝업] - [릴리즈 이력 삭제] @ 2018.06.12 LSH
procedure TMainDlg.mi_DelReleaseClick(Sender: TObject);
begin
   UpdReleaseDelDt;
end;

// 릴리즈 히스토리(이력) 삭제 @ 2018.06.12 LSH
procedure TMainDlg.mi_ModifyClick(Sender: TObject);
begin
   if asg_WorkConn.Cells[C_WC_DUTYUSER, asg_WorkConn.Row] <> FsUserNm then
   begin
      MessageBox(self.Handle,
                 PChar('본인이 작성한 내역만 수정이 가능합니다.'),
                 PChar('주요업무공유 수정전 본인 확인'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;

   // 주요업무공유 내역 수정 Update
   SetPanelStatus('WORKCONNMOD', 'ON');

end;



procedure TMainDlg.mi_SystemAllPrintClick(Sender: TObject);
begin
   // 시스템 유지보수 지급 양식 일괄출력 (FTP)
   KDialPrint('SYSTEMALL',
              asg_SelDoc);
end;



//------------------------------------------------------------------------------
// CRA/CRC ID 종료일자 업데이트 @ 2014.09.03 LSH
//------------------------------------------------------------------------------
procedure TMainDlg.mi_CRADelUpClick(Sender: TObject);
var
   TpActDoc   : TTpSvc;
   i, iCnt    : Integer;
   vType1   ,
   vLocate  ,
   vDocList ,
   vDocYear ,
   vDocSeq  ,
   vDocTitle,
   vRegDate ,
   vRegUser ,
   vRelDept ,
   vDocRmk  ,
   vDocLoc  ,
   vDelDate ,
   vEditIp,
   vDeptCd,
   vSocNo,
   vMobile,
   vGrpId,
   vUseLvl : Variant;
   sType, sLocateNm, sDelDate : String;

begin

   if not IsLogonUser then
   begin
      MessageBox(self.Handle,
                 PChar('[비상연락망] - [담당자정보]에 담당자 IP 등록후 사용하실 수 있습니다.'),
                 PChar(Self.Caption + ' : 등록되지 않은 User 제한 알림'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;




   sType    := '3';
   sDelDate := FormatDateTime('yyyy-mm-dd', Date);



   // 접속자 IP 식별후, 근무처 Assign
   if PosByte('안암도메인', FsUserIp) > 0 then
      sLocateNm := '안암'
   else if PosByte('구로도메인', FsUserIp) > 0 then
      sLocateNm := '구로'
   else if PosByte('안산도메인', FsUserIp) > 0 then
      sLocateNm := '안산';



   iCnt  := 0;


   //-------------------------------------------------------------------
   // 2. Create Variables
   //-------------------------------------------------------------------
   with asg_SelDoc do
   begin
      for i := Row to Row do
      begin
         //------------------------------------------------------------------
         // 2-1. Append Variables
         //------------------------------------------------------------------
         AppendVariant(vType1   ,   sType                    );
         AppendVariant(vLocate  ,   sLocateNm                );
         AppendVariant(vDocList ,   Cells[C_SD_DOCLIST,    i]);
         AppendVariant(vDocYear ,   asg_RegDoc.Cells[C_RD_DOCYEAR, asg_RegDoc.Row]);
         AppendVariant(vDocSeq  ,   Cells[C_SD_DOCSEQ,     i]);
         AppendVariant(vDocTitle,   Cells[C_SD_DOCTITLE,   i]);
         AppendVariant(vRegDate ,   Cells[C_SD_REGDATE,    i]);
         AppendVariant(vRegUser ,   Cells[C_SD_REGUSER,    i]);
         AppendVariant(vRelDept ,   Cells[C_SD_RELDEPT,    i]);
         AppendVariant(vDocRmk  ,   Cells[C_SD_DOCRMK,     i]);
         AppendVariant(vDocLoc  ,   Cells[C_SD_DOCLOC,     i]);
         AppendVariant(vEditIp  ,   FsUserIp                 );
         AppendVariant(vDelDate ,   sDelDate                 );

         Inc(iCnt);
      end;
   end;


   // CRA/CRC ID 종료전 확인 Pop-Up
   if Application.MessageBox(PChar(asg_SelDoc.Cells[C_SD_REGUSER, asg_SelDoc.Row] + '(' + asg_SelDoc.Cells[C_SD_DOCRMK, asg_SelDoc.Row] + ') 님의 ID를 [종료]하시겠습니까?'),
                             PChar('CRA/CRC ID 종료 등록전 확인'), MB_OKCANCEL) <> IDOK then
      Exit;



   if iCnt <= 0 then
      Exit;


   //-------------------------------------------------------------------
   // 3. Insert by TpSvc
   //-------------------------------------------------------------------
   TpActDoc := TTpSvc.Create;
   TpActDoc.Init(Self);


   Screen.Cursor := crHourGlass;


   try
      if TpActDoc.PutSvc('MD_KUMCM_M1',
                          [
                           'S_TYPE1'   , vType1
                         , 'S_STRING1' , vLocate
                         , 'S_STRING5' , vMobile
                         , 'S_STRING8' , vEditIp
                         , 'S_STRING9' , vDeptCd
                         , 'S_STRING10', vDelDate
                         , 'S_STRING16', vDocList
                         , 'S_STRING17', vDocYear
                         , 'S_STRING18', vDocSeq
                         , 'S_STRING19', vDocTitle
                         , 'S_STRING20', vRegDate
                         , 'S_STRING21', vRegUser
                         , 'S_STRING22', vRelDept
                         , 'S_STRING23', vDocRmk
                         , 'S_STRING24', vDocLoc
                         , 'S_STRING32', vUseLvl
                         , 'S_STRING33', vGrpId
                         , 'S_STRING44', vSocNo
                         ] ) then
      begin
         {
         MessageBox(self.Handle,
                    '선택하신 CRA/CRC ID가 정상적으로 [종료]되었습니다.',
                    '[KUMC 다이얼로그] CRA/CRC ID 종료 업데이트 알림 ',
                    MB_OK + MB_ICONINFORMATION);
         }

         // Refresh
         SelGridInfo('DOCREFRESH');

      end
      else
      begin
         ShowMessage(GetTxMsg);
      end;


   finally
      FreeAndNil(TpActDoc);
      Screen.Cursor  := crDefault;
   end;
end;



// [문서관리] - [마우스 팝업] - [문서정보 수정] @ 2014.09.17 LSH
procedure TMainDlg.mi_CopyDocInfoClick(Sender: TObject);
begin
   with asg_RegDoc do
   begin
      Cells[C_RD_GUBUN,    Row] := '수정';
      Cells[C_RD_DOCLIST,  Row] := asg_SelDoc.Cells[C_SD_DOCLIST,  asg_SelDoc.Row];


      if (CopyByte(asg_SelDoc.Cells[C_SD_REGDATE, asg_SelDoc.Row], 6, 2) = '01') or
         (CopyByte(asg_SelDoc.Cells[C_SD_REGDATE, asg_SelDoc.Row], 6, 2) = '02') then
         Cells[C_RD_DOCYEAR, Row] := IntToStr(StrToInt((CopyByte(asg_SelDoc.Cells[C_SD_REGDATE, asg_SelDoc.Row], 1, 4))) - 1)
      else
         Cells[C_RD_DOCYEAR, Row] := CopyByte(asg_SelDoc.Cells[C_SD_REGDATE, asg_SelDoc.Row], 1, 4);


      Cells[C_RD_DOCSEQ,   Row] := asg_SelDoc.Cells[C_SD_DOCSEQ,   asg_SelDoc.Row];
      Cells[C_RD_DOCTITLE, Row] := asg_SelDoc.Cells[C_SD_DOCTITLE, asg_SelDoc.Row];
      Cells[C_RD_REGDATE,  Row] := asg_SelDoc.Cells[C_SD_REGDATE,  asg_SelDoc.Row];
      Cells[C_RD_REGUSER,  Row] := asg_SelDoc.Cells[C_SD_REGUSER,  asg_SelDoc.Row];
      Cells[C_RD_RELDEPT,  Row] := asg_SelDoc.Cells[C_SD_RELDEPT,  asg_SelDoc.Row];
      Cells[C_RD_DOCRMK,   Row] := asg_SelDoc.Cells[C_SD_DOCRMK,   asg_SelDoc.Row];
      Cells[C_RD_DOCLOC,   Row] := asg_SelDoc.Cells[C_SD_DOCLOC,   asg_SelDoc.Row];

      AddButton(C_RD_BUTTON,  Row, ColWidths[C_RD_BUTTON]-5,  20, 'Update', haBeforeText, vaCenter);
   end;
end;


//------------------------------------------------------------------------------
// [커뮤니티] 방향키 / Enter키 입력시 Jog-Control 연동
//
// Date   : 2015.03.19
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.asg_BoardKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   case (Key) of

      VK_RETURN : asg_Board_DblClick(Sender);

      VK_LEFT   : if fsbt_ForWard.Visible then
                     fsbt_ForWardClick(Sender);

      VK_RIGHT  : if fsbt_BackWard.Visible then
                     fsbt_BackWardClick(Sender);

      //[Ctrl + C] : 선택한 Cell 클립보드로 복사 @ 2016.11.18 LSH
      Ord('C') :
                  if (ssCtrl in Shift) then
                  begin
                     asg_Board.CopySelectionToClipboard;
                  end;

      //[Ctrl + A] : 해당 Grid Cell 전체내역 클립보드로 복사 @ 2016.11.18 LSH
      Ord('A') :
                  if (ssCtrl in Shift) then
                  begin
                     asg_Board.CopyToClipBoard;
                  end;

      //[Ctrl + F] : 검색화면 호출
      Ord('F') :
                  if (ssCtrl in Shift) then
                  begin
                     fsbt_SearchClick(Sender);
                  end;
   end;
end;



//------------------------------------------------------------------------------
// [커뮤니티] 글 검색모드 Title 키워드 입력후 자동 검색
//
// Date   : 2015.03.19
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.fed_TitleKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   if fsbt_WriteFind.Visible then
   begin
      // 글 검색 모드에서, [제목] 란에 키워드 입력후 <Enter> 누르면 자동 검색 실시
      if Key = VK_RETURN then
         fsbt_WriteFindClick(Sender);
   end;
end;


procedure TMainDlg.SetUpdLogOut(Sender : TObject);
var
   varResult : String;
begin

   try
      Screen.Cursor := crHourGlass;

      // 해당 Chat - User 로그아웃 정보 업데이트
      UpdateLog('START',
                '',
                Trim(pn_ChatUserIp.Caption),
                'O',
                FormatDateTime('yyyy-mm-dd hh:nn', Now),
                '',
                varResult
                );
   except
      Screen.Cursor := crDefault;
   end;

end;


procedure TMainDlg.SetUpdChatLog(Sender : TObject);
var
   varResult : String;
begin

   try
      Screen.Cursor := crHourGlass;

      // 해당 Chat - User 로그아웃 정보 업데이트
      UpdateLog('CBLOG',
                Trim(pn_ChatUserIp.Caption),
                FsUserIp,
                'R',
                '',
                '',
                varResult
                );
   except
      Screen.Cursor := crDefault;
   end;

end;

procedure TMainDlg.SetUpdChatSearch(Sender : TObject);
var
   varResult : String;
begin

   try
      Screen.Cursor := crHourGlass;

      // 로그 Update
      UpdateLog('DIALCHAT',
                'KEYWORD',
                FsUserIp,
                'F',
                Trim(pn_ChatSearchKey.Caption),
                '',
                varResult
                );

   except
      Screen.Cursor := crDefault;
   end;

end;

//--------------------------------------------------------------------------
// 문자열 처음부터 KeyStr이 나오기 직전까지를 잘라냄
//      - 소스출처 : MComfunc.pas
//
// Date   : 2015.03.31
// Author : Lee, Se-Ha
//--------------------------------------------------------------------------
function TMainDlg.FindStr(Fstr : String; KeyStr : String) : String;
var
    Cnt,idx : Integer;
    ResultStr: String;
begin
   ResultStr := '';

   Cnt := LengthByte(FStr);

   if (Cnt = 0) then
   begin
      Result := '';
      Exit;
   end;


   for idx := 1 to Cnt do
   begin
      if (FStr[idx] = KeyStr) then
      begin
         Result := ResultStr;

         Exit;
      end
      else
      begin
         ResultStr := ResultStr + FStr[idx];
      end;
   end;
end;


// 사용자 프로필 이미지 동적 FTP 다운로드 @ 2015.03.31 LSH
function TMainDlg.GetFTPImage(in_PhotoFile, in_HiddenFile, in_UserId : String) : String;
var
   ServerFileName, ClientFileName: String;
   FTP_SVRIP, FTP_USERID, FTP_PASSWD, FTP_HOSTNAME, FTP_DIR : String;
   sRemoteFile, sLocalFile : String;
begin

   Result := '';


   // 프로필 이미지 FTP 파일내역이 존재하면, 이미지 Download
   if (in_PhotoFile <> '') then
   begin
      // MIS 기본 이미지 정보인 경우.
      if (in_HiddenFile = '') then
      begin

         // 담당자 프로필 이미지 파일명
         ServerFileName := in_PhotoFile;


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
            FTP_USERID := '';
            FTP_PASSWD := '';
            FTP_DIR    := '/kumis/app/mis/media/cq/photo/';



            // Image 다운로드
            if Not GetBINFTP(FTP_SVRIP,
                             FTP_USERID,
                             FTP_PASSWD,
                             FTP_DIR + ServerFileName,
                             ClientFileName,
                             False) then
            begin
               //Showmessage('이미지 다운로드에 실패했습니다. 다시한번 시도해 주십시오');

               TUXFTP := nil;

               Exit;
            end;



            // 해당 Image Variant로 받음.
            AppendVariant(DownFile, ClientFileName);



            // 다운로드 횟수 체크
            DownCnt := DownCnt + 1;


            Result := ClientFileName;

         end
         else
            Result := ClientFileName;
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
         if PosByte('/ftpspool/KDIALFILE/', in_HiddenFile) > 0 then
            sRemoteFile := in_HiddenFile
         else
            sRemoteFile := '/ftpspool/KDIALFILE/' + in_HiddenFile;

         // C:\DEV --> G_HOMEDIR 전환 @ 2016.07.26 LSH
         sLocalFile  := G_HOMEDIR + 'TEMP\SPOOL\' + in_PhotoFile;



         if (GetBINFTP(FTP_SVRIP,
                       FTP_USERID,
                       FTP_PASSWD,
                       sRemoteFile,
                       sLocalFile,
                       False)) then
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

            Result := sLocalFile;
      end;
   end
   else
   begin
      // 담당자 프로필 이미지 파일명
      if in_UserId = 'XXXXX' then
      begin
         Result := '';
      end
      else
      begin
         // 담당자 HRM 프로필 이미지 파일명
         ServerFileName := in_UserId + '.jpg';

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
            FTP_USERID := '';
            FTP_PASSWD := '';
            FTP_DIR    := '/kuhrm/app/hrm/photo/';


            // HRM Image 다운로드
            if Not GetBINFTP(FTP_SVRIP,
                             FTP_USERID,
                             FTP_PASSWD,
                             FTP_DIR + ServerFileName,
                             ClientFileName,
                             False) then
            begin
               //Showmessage('이미지 다운로드에 실패했습니다. 다시한번 시도해 주십시오');

               TUXFTP := nil;

               Exit;
            end;


            // 해당 Image Variant로 받음.
            AppendVariant(DownFile, ClientFileName);


            // 다운로드 횟수 체크
            DownCnt := DownCnt + 1;

            Result := ClientFileName;

         end
         else
            Result := ClientFileName;
      end;
   end;
end;





procedure TMainDlg.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
//var
//   FForm : TForm;
begin
{
   case (Key) of

      //[Ctrl + C] : 선택한 Cell 클립보드로 복사
      Ord('L') : if (ssCtrl in Shift) then
                 begin
                     //-----------------------------------------------------
                     // 등록된 User 다이얼 Chat 리스트 팝업 @ 2015.03.30 LSH
                     //-----------------------------------------------------
                     if (IsLogonUser) then
                     begin
                        FForm := BplFormCreate('KDIALCHAT', True);

                        //if FForm <> nil then
                        begin
                           SetBplStrProp(FForm, 'prop_UserNm', FsUserNm);

                           FForm.Height := (self.Height);
                           FForm.Top    := (screen.Height - FForm.Height) - 48;
                           FForm.Left   := 8;

                           try


                              FForm.Show;



                           except
                              on e : Exception do
                              showmessage('Show Form Error(KDIALCHAT) : ' + e.Message);
                           end;


                           self.SetFocus;

                        end;
                     end;
                 end;
      end;
   }
end;


procedure TMainDlg.fsbt_DialChatClick(Sender: TObject);
//var
//   FForm : TForm;
begin
   //-----------------------------------------------------
   // 등록된 User 다이얼 Chat 리스트 팝업 @ 2015.03.30 LSH
   //-----------------------------------------------------
   {  -- Chat XE7 개선전까지 주석 @ 2017.10.31 LSH
   if (IsLogonUser) then
   begin
      FForm := BplFormCreate('KDIALCHAT', True);


      if FForm <> nil then
      begin
         SetBplStrProp(FForm, 'prop_UserNm', FsUserNm);

         FForm.Height := (self.Height);
         FForm.Top    := (screen.Height - FForm.Height) - 48;
         FForm.Left   := 8;

         try
            FForm.Show;

         except
            on e : Exception do
            showmessage('Show Form Error(KDIALCHAT) : ' + e.Message);
         end;

         self.SetFocus;

      end;
   end
   }

   //-----------------------------------------------------
   // 등록된 User 다이얼 Chat 리스트 팝업 @ 2015.03.30 LSH
   //-----------------------------------------------------
   {  -- Chat XE7 개선전까지 주석 @ 2017.10.31 LSH
   if (IsLogonUser) then
   begin
      //FForm := BplFormCreate('KDIALCHAT', True);
      FForm := FindForm('KDIALCHAT');


      if FForm = nil then
      begin
         FForm := BplFormCreate('KDIALCHAT', True);

         SetBplStrProp(FForm, 'prop_UserNm', FsUserNm);

         FForm.Height := (self.Height);
         FForm.Top    := (screen.Height - FForm.Height) - 48;
         FForm.Left   := 8;

         try
            FForm.Show;

         except
            on e : Exception do
            showmessage('Show Form Error(KDIALCHAT) : ' + e.Message);
         end;

         self.SetFocus;

      end;
   end
   else
      MessageBox(self.Handle,
                 PChar('[비상연락망] - [담당자정보]에 담당자 IP 등록후 사용하실 수 있습니다.'),
                 PChar(Self.Caption + ' : 등록되지 않은 User 제한 알림'),
                 MB_ICONWARNING + MB_OK);
   }
end;




//------------------------------------------------------------------------------
// String내의 특정 문자열을 다른 문자열로 대체
//
// Date   : 2015.04.02
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
function TMainDlg.ReplaceChar(TextStr,OrigStr,ChgStr: String):String;
var
   iPos : Integer;
begin
   while PosByte(OrigStr, TextStr) > 0 do
   begin
      iPos := PosByte(OrigStr, TextStr);
      System.Delete(TextStr,iPos,LengthByte(OrigStr));
      System.Insert(ChgStr, TextStr, iPos);
   end;

   Result := TextStr;

end;

{$ENDIF}


procedure TMainDlg.fsbt_LotRefreshClick(Sender: TObject);
begin
   // Refresh
   SelGridInfo('BIGDATA');
end;

procedure TMainDlg.asg_BigDataGetAlignment(Sender: TObject; ARow,
  ACol: Integer; var HAlign: TAlignment; var VAlign: TVAlignment);
begin
   HAlign := taCenter;
   VAlign := vtaCenter;
end;


procedure TMainDlg.asg_LikeMouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
var
   NowCol, NowRow : Integer;
begin
   try
      // 현재 Cell 좌표 가져오기
      asg_Like.MouseToCell(X, Y, NowCol, NowRow);

      if (NowCol <= asg_Like.ColCount-1) then
      begin

         with asg_Like do
         begin
            ShowHint := True;

            // 추천인 Name Hint 표기
            Hint := Cells[NowCol, NowRow + 1]
         end;
      end
      else
         ShowHint := False;

   except
      // Grid.ColCount 벗어나는 Mousemoving시 out of index 팝업 예외처리 위함.
      asg_Like.Hint     := '';
      asg_Like.ShowHint := False;
   end;

end;

//------------------------------------------------------------------------------
// [문서관리] - [개발자 릴리즈대장] - [서비스 릴리즈 스크립트 자동생성] (Memo to Grid)
//------------------------------------------------------------------------------
procedure TMainDlg.mi_RlzScriptClick(Sender: TObject);
var
   sGijunDt, sNowTime : String;
   SearchRow : Integer;
   i, j, k, l, iMaxCheck, iStrTokenCnt : Integer;

   //--------------------------------------------------------------------
   // String내에 특정 Char가 몇개 있는지 Count 하는 함수
   //--------------------------------------------------------------------
   function CountChar(in_InpString, in_Token : String) : Integer;
   var
      i, iCnt      : Integer;
   begin
      iCnt := 0;

      Result := iCnt;

      // Check Validation
      for i := 1 to LengthByte(in_InpString) do
      begin
         if in_Token = CopyByte(in_InpString, i, 1) then
         begin
            Inc(iCnt);
         end;
      end;

      Result := iCnt;

   end;

begin

   // 메모 Log 초기화
   mm_Log.Lines.Clear;


   with asg_Release do
   begin
      // 첫번째 Row의 등록일자를 기준 Date로 fetch
      if (Cells[C_RL_REGDATE, Row] <> '') then
      begin
         //sGijunDt    := CopyByte(Cells[C_RL_REGDATE, 1], 1, 10);

         // 릴리즈 스크립트 생성 기준일 체크 분기 @ 2016.11.02 LSH (월요일 릴리즈시, 금요일부터)
         if FormatDateTime('ddd', Date) = '월' then
            sGijunDt := FormatDateTime('yyyy-mm-dd', Date - 3)
         else
            sGijunDt := FormatDateTime('yyyy-mm-dd', Date - 2);

         // 현 릴리즈 기준 시간 Set @ 2016.11.16 LSH
         sNowTime := FormatDateTime('yyyy-mm-dd hh:nn', Now);

         SearchRow   := Row;
      end
      else
      begin
         MessageBox(self.Handle,
                    '스크립트 생성 대상 릴리즈 내역을 찾을 수 없습니다.',
                    '릴리즈 스크립트 생성 오류 알림',
                    MB_OK + MB_ICONERROR);

         Exit;
      end;

      // 기준일에 해당하는 서버(서비스) 릴리즈 내역을 Memo장으로 복사
      while CopyByte(Cells[C_RL_REGDATE, SearchRow], 1, 10) >= sGijunDt do
      begin
         // 서비스 릴리즈 내역 있을때만, Looping
         if (Trim(Cells[C_RL_SERVESRC, SearchRow]) <> '') and
            (Trim(Cells[C_RL_RELEASDT, SearchRow]) =  '') and           // 릴리즈 일시 없는 대상만 fetch @ 2016.11.16 LSH
            (Trim(Cells[C_RL_REGDATE,  SearchRow]) < sNowTime) then     // 현 릴리즈 기준시간 이전에 작성된 대상만 fetch @ 2016.11.16 LSH
         begin
            if PosByte('신규', Cells[C_RL_SERVESRC,  SearchRow]) > 0 then
            begin
               mm_Log.Lines.Add(Lowercase(ReplaceChar(ReplaceChar(ReplaceChar(ReplaceChar(DelChar(DelChar(DelChar(DelChar(DelChar(Cells[C_RL_SERVESRC,  SearchRow], '-'), '<'), '>'), '('), ')'), '신규', ''), ' ' , ''),#13, ''), #10, '|')));
               mm_Log.Lines.Add(Lowercase(ReplaceChar(ReplaceChar(ReplaceChar(ReplaceChar(DelChar(DelChar(DelChar(DelChar(DelChar(Cells[C_RL_SERVER,    SearchRow], '-'), '<'), '>'), '('), ')'), '신규', ''), ' ' , ''),#13, ''), #10, '.def|')) + '.def');
            end
            else
            begin
               mm_Log.Lines.Add(ReplaceChar(ReplaceChar(DelChar(DelChar(DelChar(Cells[C_RL_SERVESRC,  SearchRow], '-'), '<'), '>'), #13, ''), #10, '|'));
            end;
         end;

         Inc(SearchRow);

      end;
   end;



   asg_Memo.ClearRows(1, asg_Memo.RowCount);
   asg_Memo.RowCount := 2;



   // Looping Index 초기화
   iMaxCheck := 2;
   i         := 1;            // 메모장 Line - index
   k         := 0;            // Token 추가 Line - index
   l         := 1;            // 현 Grid Line - index


   while (i <= iMaxCheck) do
   begin

      asg_Memo.RowCount := iMaxCheck + k;

      if Trim(mm_Log.Lines.Strings[i-1]) <> '' then
      begin
         // 하나의 Line에 붙어있는 Token 개수 확인
         iStrTokenCnt := CountChar(mm_Log.Lines.Strings[i-1], '|');

         // 만약 하나의 String 라인내 Token(|)이 1개이상 있으면, Token 수만큼 다시 분리해서 Row 단위로 Line 추가
         if iStrTokenCnt > 0 then
         begin
            // 실제 Token수 만큼의 릴리즈 내역을 row 단위로 assign
            for j := 0 to iStrTokenCnt do
            begin
               asg_Memo.Cells[0, l + j + k] := TokenStr(mm_Log.Lines.Strings[i-1], '|', j+1);
            end;

            iMaxCheck := iMaxCheck + iStrTokenCnt{ + 1};

            k         := k + iStrTokenCnt {+ 1};
         end
         else
         begin
            asg_Memo.Cells[0, l+k] := mm_Log.Lines.Strings[i-1];

            Inc(IMaxCheck);
         end;

         Inc(l);

      end;

      Inc(i);

   end;


   asg_Memo.RowCount := asg_Memo.RowCount - 1;

   apn_Memo.Caption.Text := '릴리즈 (SVC/CLT) 목록';

   apn_Memo.Top     := 119;
   apn_Memo.Left    := 143;
   apn_Memo.Visible := True;
   apn_Memo.Collaps := True;
   apn_Memo.Collaps := False;

   asg_MemoClickCell(sender, 0, 0);

end;


//------------------------------------------------------------------------------
// 릴리즈 스크립트 생성 (메모 Grid 더블클릭시)
//------------------------------------------------------------------------------
procedure TMainDlg.asg_MemoDblClick(Sender: TObject);
var
   sRlzGroup   : String;
   sGroupPath  : String;
   sCLTSrcPath : String;

   i, iRlzCLTCnt : Integer;
begin
   // 메모 Log 초기화
   mm_Log.Lines.Clear;


   // SMS 템플릿 목록 연동 추가 @ 2017.07.11 LSH
   if PosByte('SMS', apn_Memo.Caption.Text) > 0 then
   begin
      // SMS 입력 Memo 란에 선택한 Template 내용 복사
      fm_Sms.Lines.Clear;
      fm_Sms.Lines.Text := asg_Memo.Cells[1, asg_Memo.Row];

      apn_Memo.Visible := False;
      apn_Memo.Collaps := True;
   end
   else
   begin
      // CLT 릴리즈 대상 Count 초기화
      iRlzCLTCnt := 0;

      // CRLF 제거후, 각 Line별 릴리즈 대상을 하나의 Line-Group으로 묶음
      for i := 1 to asg_Memo.RowCount do
      begin
         // 각 Row 단위(서비스 또는 서버)별 릴리즈 명령어 생성 진행
         if (asg_Memo.Cells[0, i-1] <> '') and
            (CopyByte(Trim(asg_Memo.Cells[0, i-1]), 7, 1) <> 'L') and     // BPL (Package) 목록은 제외 (예: MDN110L1_P01) @ 2016.12.16 LSH
            (PosByte('_', asg_Memo.Cells[0, i-1]) > 0) then
         begin
            // 이전 Row의 그룹명(ex. md / mh / mt ..)이 현 Row의 그룹명과 같을때는
            // 계속 하나의 Line으로 생성
            if (CopyByte(asg_Memo.Cells[0, i], 1, 2) = CopyByte(asg_Memo.Cells[0, i-1], 1, 2)) then
            begin
               sRlzGroup   := asg_Memo.Cells[0, i-1] + ' ' + sRlzGroup;
               sGroupPath  := CopyByte(asg_Memo.Cells[0, i-1], 1, 2);
            end
            else
            // 이전 Row의 그룹명(ex. md / mh / mt ..)이 현 Row의 그룹명과 다를때는
            // 바로 이전 Row(i-1)까지 추가하여 릴리즈 명령어 Line 최종 생성
            begin
               sRlzGroup   := asg_Memo.Cells[0, i-1] + ' ' + sRlzGroup;
               sGroupPath  := CopyByte(asg_Memo.Cells[0, i-1], 1, 2);

               mm_Log.Lines.Add('');
               mm_Log.Lines.Add('---- [테스트계] cp from upsrc to SrcPath ----');
               mm_Log.Lines.Add('cp ' + sRlzGroup + ' ../src/' + sGroupPath);
               mm_Log.Lines.Add('---- [테스트계->릴리즈계] rocscp from users to user_release ----');
               mm_Log.Lines.Add('rocscp ' + sRlzGroup);

               // 다른 그룹 명령어 Line 생성위해 초기화
               sRlzGroup   := '';
               sGroupPath  := '';
            end;
         end
         // Package (BPL) --> 델파이 자동등록(Add To Project) 적용 @ 2016.11.02 LSH
         else if (asg_Memo.Cells[0, i-1] <> '') and
                 (
                     (CopyByte(Trim(asg_Memo.Cells[0, i-1]), 7, 1) = 'L') or     // BPL
                     (PosByte('DLL', Trim(asg_Memo.Cells[0, i-1])) > 0)          // DLL
                 ) then
         begin
            // BPL 개발 Source 경로 Set.
            if (CopyByte(Trim(asg_Memo.Cells[0, i-1]), 7, 1) = 'L') then
            begin
               if (CopyByte(Trim(asg_Memo.Cells[0, i-1]), 1, 2) = 'MN') then
                  sCLTSrcPath := 'MN\' + Trim(asg_Memo.Cells[0, i-1]) + '\'
               else if (CopyByte(Trim(asg_Memo.Cells[0, i-1]), 1, 2) = 'MA') then
                  sCLTSrcPath := 'MA\' + Trim(asg_Memo.Cells[0, i-1]) + '\'
               else if (CopyByte(Trim(asg_Memo.Cells[0, i-1]), 1, 2) = 'MH') then
                  sCLTSrcPath := 'MH\' + Trim(asg_Memo.Cells[0, i-1]) + '\'
               else if (CopyByte(Trim(asg_Memo.Cells[0, i-1]), 1, 2) = 'MJ') then
                  sCLTSrcPath := 'MJ\' + Trim(asg_Memo.Cells[0, i-1]) + '\'
               else if (CopyByte(Trim(asg_Memo.Cells[0, i-1]), 1, 2) = 'MT') then
                  sCLTSrcPath := 'MT\' + Trim(asg_Memo.Cells[0, i-1]) + '\'
               else if (CopyByte(Trim(asg_Memo.Cells[0, i-1]), 1, 2) = 'MX') then
                  sCLTSrcPath := 'MX\' + Trim(asg_Memo.Cells[0, i-1]) + '\'
               else if (CopyByte(Trim(asg_Memo.Cells[0, i-1]), 1, 2) = 'MO') then
                  sCLTSrcPath := 'MO\' + Trim(asg_Memo.Cells[0, i-1]) + '\'
               else if (CopyByte(Trim(asg_Memo.Cells[0, i-1]), 1, 2) = 'ER') then
                  sCLTSrcPath := 'ER\' + Trim(asg_Memo.Cells[0, i-1]) + '\'
               else
                  sCLTSrcPath := 'MD\' + Trim(asg_Memo.Cells[0, i-1]) + '\';
            end
            // DLL 개발 Source 경로 Set.
            else if (PosByte('DLL', Trim(asg_Memo.Cells[0, i-1])) > 0) then
            begin
               if (PosByte('MDINFDLL', Trim(asg_Memo.Cells[0, i-1])) > 0) then
                  sCLTSrcPath := 'MDDLL\간호정보조사지\'
               else
                  sCLTSrcPath := 'MDDLL\' + Trim(asg_Memo.Cells[0, i-1]) + '\';
            end;


            // 각 경로의 .dproj를 찾아서 실행함과 동시에 Delphi XE7 Project Manager에 자동등록
            ShellExecute(HANDLE, 'open',
                         PCHAR(Trim(asg_Memo.Cells[0, i-1]) + '.dproj'),
                         PCHAR(''),
                         PCHAR('C:\KUMC_DEV\Src\SrcM\' + sCLTSrcPath),
                         SW_NORMAL);

            // 추가된 CLT Count ++
            Inc(iRlzCLTCnt);

            // 확인 Message
            MessageBox(self.Handle,
                       PChar(IntToStr(iRlzCLTCnt) + '번째 BPL : ' + Trim(asg_Memo.Cells[0, i-1]) + ' --> Project Manager 등록완료.'),
                       '릴리즈 대상 Package 자동추가',
                       MB_ICONINFORMATION + MB_OK)


         end;
      end;
   end;
end;



procedure TMainDlg.asg_MemoClickCell(Sender: TObject; ARow, ACol: Integer);
begin
   asg_Memo.SortSettings.Show := True;
end;



//------------------------------------------------------------------------------
// 문서관리 - 근무 및 근태현황 - 근무표 생성(관리자)
//
// Author : Lee, Se-Ha
// Date   : 2015.06.12
//------------------------------------------------------------------------------
procedure TMainDlg.fcb_DutyMakeClick(Sender: TObject);
begin
   // 근무표 생성 Check시
   if (fcb_DutyMake.Checked = True) then
   begin
      if GetMsgInfo('INFORM',
                    'DUTYLIST') = FsUserNm then
         SelGridInfo('DUTYMAKE');
   end
   else
   // 근무표 생성 미 Check시, 기존 확정이력(MKUMCDTY) 조회
   begin
      SelGridInfo('DUTYSEARCH');
   end;

   // 근무표 확정(저장) Enabled 조건
   fsbt_DutyInsert.Enabled := fcb_DutyMake.Checked;

end;

//------------------------------------------------------------------------------
// 당직근무표 확정/저장
//
// Author : Lee, Se-Ha
// Date   : 2015.06.12
//------------------------------------------------------------------------------
procedure TMainDlg.fsbt_DutyInsertClick(Sender: TObject);
var
   TpInsDuty : TTpSvc;
   i, iCnt   : Integer;
   vType1    ,
   vDutyDate ,
   vDutyYoil ,
   vDutyFlag ,
   vDutyUser ,
   vDayRmk   ,
   vDelDate  ,
   vEditId   ,
   vEditIp   : Variant;
   sType     : String;
begin

   // 당직근무표 확정/저장
   sType := '16';


   iCnt  := 0;


   //-------------------------------------------------------------------
   // 2. Create Variables
   //-------------------------------------------------------------------
   with asg_Duty do
   begin
      for i := 1 to RowCount - 1 do
      begin
         //------------------------------------------------------------------
         // 2-1. Append Variables
         //------------------------------------------------------------------
         AppendVariant(vType1    ,   sType                  );
         AppendVariant(vDutyDate ,   Cells[C_DT_DUTYDATE  ,i]);
         AppendVariant(vDutyYoil ,   Cells[C_DT_YOIL      ,i]);
         AppendVariant(vDutyFlag ,   Cells[C_DT_DUTYFLAG  ,i]);
         AppendVariant(vDutyUser ,   Cells[C_DT_DUTYUSER  ,i]);
         AppendVariant(vDayRmk   ,   Cells[C_DT_REMARK    ,i]);
         AppendVariant(vDelDate  ,   '');
         AppendVariant(vEditId   ,   FsUserNm);
         AppendVariant(vEditIp   ,   FsUserIp);

         Inc(iCnt);
      end;
   end;


   if iCnt <= 0 then
      Exit;


   //-------------------------------------------------------------------
   // 3. Insert by TpSvc
   //-------------------------------------------------------------------
   TpInsDuty := TTpSvc.Create;
   TpInsDuty.Init(Self);


   Screen.Cursor := crHourGlass;


   try
      if TpInsDuty.PutSvc('MD_KUMCM_M1',
                          [
                           'S_TYPE1'   , vType1
                         , 'S_STRING2' , vEditId
                         , 'S_STRING4' , vDutyYoil
                         , 'S_STRING8' , vEditIp
                         , 'S_STRING10', vDelDate
                         , 'S_STRING11', vDutyFlag
                         , 'S_STRING12', vDayRmk
                         , 'S_STRING13', vDutyUser
                         , 'S_STRING20', vDutyDate

                         { -- 확장성 고려해서 남겨둠..
                         , 'S_STRING21', vReqUser
                         , 'S_STRING22', vReqDept
                         , 'S_STRING28', vContext
                         , 'S_STRING37', vClient
                         , 'S_STRING38', vServer
                         , 'S_STRING44', vSrreqNo
                         , 'S_STRING46', vClienSrc
                         , 'S_STRING47', vServeSrc
                         , 'S_STRING48', vCofmUser
                         , 'S_STRING49', vReleasDt
                         , 'S_STRING50', vTestDate
                         , 'S_TYPE2'   , vReleasUser
                         }
                         ] ) then
      begin
         MessageBox(self.Handle,
                    PChar(IntToStr(iCnt) +'건의 당직근무이력이 정상적으로 [등록] 되었습니다.'),
                    '[KUMC 다이얼로그] 당직근무표 업데이트 알림 ',
                    MB_OK + MB_ICONINFORMATION);

         // 당직근무표 생성(관리자) Disabled
         fcb_DutyMake.Checked := False;

         // Refresh
         SelGridInfo('DUTYSEARCH');
      end
      else
      begin
         ShowMessageM(GetTxMsg);
      end;


   finally
      FreeAndNil(TpInsDuty);
      Screen.Cursor  := crDefault;
   end;
end;


//------------------------------------------------------------------------------
// [근무 및 근태] TAdvStringGrid onCanEditCell Event Hander
//
// Author : Lee, Se-Ha
// Date   : 2015.06.12
//------------------------------------------------------------------------------
procedure TMainDlg.asg_DutyCanEditCell(Sender: TObject; ARow,
  ACol: Integer; var CanEdit: Boolean);
begin
   // 당직자, 근태사항만 수정가능
   CanEdit := ACol in [C_DT_DUTYUSER, C_DT_REMARK];
end;

//------------------------------------------------------------------------------
// [근무 및 근태] TAdvStringGrid onKeyDown Event Hander
//
// Author : Lee, Se-Ha
// Date   : 2015.06.12
//------------------------------------------------------------------------------
procedure TMainDlg.asg_DutyKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   with asg_Duty do
   begin
      // 당직자 컬럼 변경사항 발생시(Enter 키) 확인 Msg
      if (Col = C_DT_DUTYUSER) and
         (Cells[C_DT_DUTYUSER, Row] <> Cells[C_DT_ORGDTYUSR, Row]) and
         (Key = 13) then
      begin
         if MessageBox(self.Handle,
                       PChar('당직 근무를 [' + Cells[C_DT_ORGDTYUSR, Row] + ' -> ' + Cells[C_DT_DUTYUSER, Row] + '](으)로 수정하시겠습니까?'),
                       PChar(Cells[C_DT_DUTYDATE,  Row] + ' (' + Cells[C_DT_YOIL,  Row] + ') 근무자 변경전 확인'),
                       MB_YESNO + MB_ICONQUESTION) = ID_YES then
         begin
            UpdateDuty('DUTYUSER');
         end
         else
            Cells[C_DT_DUTYUSER, Row] := Cells[C_DT_ORGDTYUSR, Row];

      end
      // 근태(휴가) 특기사항 변경 발생시(Enter키) 확인 Msg
      else if (Col = C_DT_REMARK) and
              (Trim(Cells[C_DT_REMARK, Row]) <> Trim(Cells[C_DT_ORGREMARK, Row])) and
              (Key = 13) then
      begin
         if MessageBox(self.Handle,
                       PChar('근태(휴가) 및 특기사항 변경사항을 저장하시겠습니까?'),
                       PChar(Cells[C_DT_DUTYDATE,  Row] + ' (' + Cells[C_DT_YOIL,  Row] + ') 근태(휴가) 및 특기사항 변경전 확인'),
                       MB_YESNO + MB_ICONQUESTION) = ID_YES then
         begin
            UpdateDuty('DUTYUSER');
         end
         else
         begin
            if Trim(Cells[C_DT_ORGREMARK, Row]) <> '' then
               Cells[C_DT_REMARK, Row] := Cells[C_DT_ORGREMARK, Row]
            else
               Cells[C_DT_REMARK, Row] := '';
         end;

      end;
   end;
end;



//------------------------------------------------------------------------------
// [근무 및 근태] 근무-근태-당직일지 변경사항 Update
//
// Author : Lee, Se-Ha
// Date   : 2015.06.12
//------------------------------------------------------------------------------
procedure TMainDlg.UpdateDuty(in_UpdFlag : String);
var
   TpUpdDuty : TTpSvc;
   i, iCnt   : Integer;
   vType1    ,
   vDutyDate ,
   vDutyYoil ,
   vSeqno    ,
   vDutyFlag ,
   vDutyUser ,
   vOrgDtyUsr,
   vDayRmk   ,
   vOrgRemark,
   vDelDate  ,
   vEditId   ,
   vEditIp   : Variant;
   sType     : String;
begin
   //---------------------------------------------------------------
   // 당직근무자 및 근태(휴가) 특기사항 수정
   //---------------------------------------------------------------
   if (in_UpdFlag = 'DUTYUSER') then
      sType := '17'
   //---------------------------------------------------------------
   // 당직근무 Seq 및 평일/휴일 근무시작일 수정
   //---------------------------------------------------------------
   else if (in_UpdFlag = 'DUTYSEQ') then
      sType := '18'
   //---------------------------------------------------------------
   // 당직일지 작성(등록)
   //---------------------------------------------------------------
   else if (in_UpdFlag = 'DUTYRPT') then
      sType := '19';


   iCnt  := 0;


   //-------------------------------------------------------------------
   // 2. Create Variables
   //-------------------------------------------------------------------
   with asg_Duty do
   begin
      //------------------------------------------------------------------
      // 2-1. Append Variables
      //------------------------------------------------------------------
      if (in_UpdFlag = 'DUTYUSER') then
      begin
         AppendVariant(vType1    ,   sType                  );
         AppendVariant(vDutyDate ,   Cells[C_DT_DUTYDATE  ,Row]);
         AppendVariant(vDutyYoil ,   Cells[C_DT_YOIL      ,Row]);
         AppendVariant(vSeqno    ,   Cells[C_DT_SEQNO     ,Row]);
         AppendVariant(vDutyFlag ,   Cells[C_DT_DUTYFLAG  ,Row]);
         AppendVariant(vDutyUser ,   Cells[C_DT_DUTYUSER  ,Row]);
         AppendVariant(vOrgDtyUsr,   Cells[C_DT_ORGDTYUSR ,Row]);
         AppendVariant(vDayRmk   ,   Cells[C_DT_REMARK    ,Row]);
         AppendVariant(vOrgRemark,   Cells[C_DT_ORGREMARK ,Row]);
         AppendVariant(vDelDate  ,   '');
         AppendVariant(vEditId   ,   FsUserNm);
         AppendVariant(vEditIp   ,   FsUserIp);

         Inc(iCnt);
      end
      else if (in_UpdFlag = 'DUTYSEQ') then
      begin
         for i := 1 to asg_DutyList.RowCount - 1 do
         begin
            AppendVariant(vType1    ,   sType            );
            AppendVariant(vSeqno    ,   asg_DutyList.Cells[0, i]    );
            AppendVariant(vDutyUser ,   asg_DutyList.Cells[1, i]    );
            AppendVariant(vDutyDate ,   asg_DutyList.Cells[2, i]    );    // 평일근무 시작 Date
            AppendVariant(vDutyFlag ,   asg_DutyList.Cells[3, i]    );    // 휴일근무 시작 Date
            AppendVariant(vDelDate  ,   ''               );
            AppendVariant(vEditId   ,   ''               );
            AppendVariant(vEditIp   ,   FsUserIp         );

            Inc(iCnt);
         end;
      end
      else if (in_UpdFlag = 'DUTYRPT') then
      begin
         // 진행예정
      end;
   end;


   if iCnt <= 0 then
      Exit;


   //-------------------------------------------------------------------
   // 3. Insert by TpSvc
   //-------------------------------------------------------------------
   TpUpdDuty := TTpSvc.Create;
   TpUpdDuty.Init(Self);


   Screen.Cursor := crHourGlass;


   try
      if TpUpdDuty.PutSvc('MD_KUMCM_M1',
                          [
                           'S_TYPE1'   , vType1
                         , 'S_STRING2' , vEditId
                         , 'S_STRING4' , vDutyYoil
                         , 'S_STRING8' , vEditIp
                         , 'S_STRING10', vDelDate
                         , 'S_STRING11', vDutyFlag
                         , 'S_STRING12', vDayRmk
                         , 'S_STRING13', vDutyUser
                         , 'S_STRING18', vSeqno
                         , 'S_STRING20', vDutyDate
                         , 'S_STRING21', vOrgDtyUsr
                         , 'S_STRING28', vOrgRemark

                         { -- 확장성 고려 남겨둠..
                         , 'S_STRING37', vClient
                         , 'S_STRING38', vServer
                         , 'S_STRING44', vSrreqNo
                         , 'S_STRING46', vClienSrc
                         , 'S_STRING47', vServeSrc
                         , 'S_STRING48', vCofmUser
                         , 'S_STRING49', vReleasDt
                         , 'S_STRING50', vTestDate
                         , 'S_TYPE2'   , vReleasUser
                         }
                         ] ) then
      begin
         MessageBox(self.Handle,
                    PChar(IntToStr(iCnt) +'건의 당직근무 변경사항이 정상적으로 [업데이트] 되었습니다.'),
                    '[KUMC 다이얼로그] 당직근무 내용 업데이트 알림 ',
                    MB_OK + MB_ICONINFORMATION);

         // Refresh
         if (in_UpdFlag = 'DUTYUSER') then
            SelGridInfo('DUTYSEARCH')
         else if (in_UpdFlag = 'DUTYSEQ') then
            SelGridInfo('DUTYLIST');

      end
      else
      begin
         ShowMessageM(GetTxMsg);
      end;


   finally
      FreeAndNil(TpUpdDuty);
      Screen.Cursor  := crDefault;
   end;
end;




// 당직일지 작성화면 닫기
procedure TMainDlg.fsbt_DutyRptCancelClick(Sender: TObject);
begin
   fpn_DutyRpt.Top     := 999;
   fpn_DutyRpt.Left    := 999;
   fpn_DutyRpt.Visible := False;
end;




procedure TMainDlg.asg_DutyDblClick(Sender: TObject);
begin
   if not IsLogonUser then
   begin
      MessageBox(self.Handle,
                 PChar('[비상연락망] - [담당자정보]에 담당자 IP 등록후 사용하실 수 있습니다.'),
                 PChar(Self.Caption + ' : 등록되지 않은 User 제한 알림'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;


   fpn_DutyRpt.Top     := 0;
   fpn_DutyRpt.Left    := 0;
   fpn_DutyRpt.Visible := True;

   fpn_DutyRpt.BringToFront;


   with asg_Duty do
   begin
      fed_DutyDate.Text := CopyByte(Cells[C_DT_DUTYDATE, Row], 1, 4) + '년 ' +
                           CopyByte(Cells[C_DT_DUTYDATE, Row], 6, 2) + '월 ' +
                           CopyByte(Cells[C_DT_DUTYDATE, Row], 9, 2) + '일 ';

      fed_DutyUser.Text := Cells[C_DT_DUTYUSER, Row];
   end;

   {
   fed_BoardSeq.Text    := '';
   fcb_CateUp.Enabled   := True;
   fed_Writer.Enabled   := True;
   fmm_Text.Enabled     := True;
   fed_Title.Enabled    := True;
   fed_Title.Text       := '';


   if FsNickNm <> '' then
      fed_Writer.Text      := FsNickNm
   else
      fed_Writer.Text      := FsUserNm;


   fmm_Text.Lines.Text  := '';
   fed_Attached.Text    := '';
   fcb_CateUp.Text      := '';
   fcb_CateUp.SetFocus;


   fed_CateDownNm.Visible  := False;
   fcb_CateDown.Visible    := False;
   fed_AlarmFro.Visible    := False;
   fed_AlarmTo.Visible     := False;
   fmed_AlarmFro.Visible   := False;
   fmed_AlarmTo.Visible    := False;
   }

   {
   fsbt_FileOpen.Enabled  := True;
   fsbt_FileClear.Enabled := True;
   }

   fsbt_DutyRptInsert.Visible := True;
   fsbt_DutyRptCancel.Visible := True;

   {
   fed_Title.Width    := 510;
   fmm_Text.Width     := 510;
   fmm_Text.MaxLength := 500;
   fmm_Text.Hint      := '신규 작성글은 한글기준 15줄(500자 정도)이내로 제한합니다.';
   }

   lb_DutyRpt.Caption    := '▶ 당직일지 작성(신규/수정) 모드입니다.';
   lb_DutyRpt.Font.Color := $006C8518;
   lb_DutyRpt.Font.Style := [fsBold];

end;

//------------------------------------------------------------------------------
// [문서관리] - [개발자 릴리즈대장] - [BPL --> 델파이 자동추가]
//
// Author : Lee, Se-Ha
// Date   : 2016.11.02
//------------------------------------------------------------------------------
procedure TMainDlg.mi_BPL2DelphiClick(Sender: TObject);
var
   sGijunDt, sNowTime : String;
   SearchRow : Integer;
   i, j, k, l, iMaxCheck, iStrTokenCnt : Integer;

   //--------------------------------------------------------------------
   // String내에 특정 Char가 몇개 있는지 Count 하는 함수
   //--------------------------------------------------------------------
   function CountChar(in_InpString, in_Token : String) : Integer;
   var
      i, iCnt      : Integer;
   begin
      iCnt := 0;

      Result := iCnt;

      // Check Validation
      for i := 1 to LengthByte(in_InpString) do
      begin
         if in_Token = CopyByte(in_InpString, i, 1) then
         begin
            Inc(iCnt);
         end;
      end;

      Result := iCnt;

   end;
begin
   // 메모 Log 초기화
   mm_Log.Lines.Clear;

   // 첫번째 Row의 등록일자를 기준 Date로 fetch
   if (asg_Release.Cells[C_RL_REGDATE, asg_Release.Row] <> '') then
   begin
      //sGijunDt    := CopyByte(Cells[C_RL_REGDATE, 1], 1, 10);

      // 릴리즈 스크립트 생성 기준일 체크 분기 (월요일 릴리즈시, 금요일부터)
      if FormatDateTime('ddd', Date) = '월' then
         sGijunDt := FormatDateTime('yyyy-mm-dd', Date - 3)
      else
         sGijunDt := FormatDateTime('yyyy-mm-dd', Date - 2);

      // 현 릴리즈 기준 시간 Set @ 2016.11.16 LSH
      sNowTime := FormatDateTime('yyyy-mm-dd hh:nn', Now);

      SearchRow   := asg_Release.Row;
   end
   else
   begin
      MessageBox(self.Handle,
                 'Package(BPL) 릴리즈 대상 내역을 찾을 수 없습니다.',
                 'Package (BPL) 릴리즈 대상 없음 알림',
                 MB_OK + MB_ICONERROR);

      Exit;
   end;

   // 기준일에 해당하는 Client(BPL) 릴리즈 내역을 찾아서, 개발 Source 경로의 해당 [프로젝트명.dproj]를 ShellExecute 하여 add to Project 적용.
   while CopyByte(asg_Release.Cells[C_RL_REGDATE, SearchRow], 1, 10) >= sGijunDt do
   begin
      if (Trim(asg_Release.Cells[C_RL_CLIENT,   SearchRow]) <> '') and           // CLT 릴리즈 내역 있을때만, Looping
         (Trim(asg_Release.Cells[C_RL_RELEASDT, SearchRow]) =  '') and           // 릴리즈 일시 없는 대상만 fetch @ 2016.11.16 LSH
         (Trim(asg_Release.Cells[C_RL_REGDATE,  SearchRow]) < sNowTime) then     // 현 릴리즈 기준시간 이전에 작성된 대상만 fetch @ 2016.11.16 LSH
      begin
         if PosByte('신규', asg_Release.Cells[C_RL_CLIENT,  SearchRow]) > 0 then
         begin
            mm_Log.Lines.Add(Lowercase(ReplaceChar(ReplaceChar(ReplaceChar(ReplaceChar(DelChar(DelChar(DelChar(DelChar(DelChar(asg_Release.Cells[C_RL_CLIENT, SearchRow], '-'), '<'), '>'), '('), ')'), '신규', ''), ' ' , ''),#13, ''), #10, '|')));
         end
         else
         begin
            mm_Log.Lines.Add(ReplaceChar(ReplaceChar(DelChar(DelChar(DelChar(asg_Release.Cells[C_RL_CLIENT,  SearchRow], '-'), '<'), '>'), #13, ''), #10, '|'));
         end;
      end;

      Inc(SearchRow);

   end;


   asg_Memo.ClearRows(1, asg_Memo.RowCount);
   asg_Memo.RowCount := 2;



   // Looping Index 초기화
   iMaxCheck := 2;
   i         := 1;            // 메모장 Line - index
   k         := 0;            // Token 추가 Line - index
   l         := 1;            // 현 Grid Line - index


   while (i <= iMaxCheck) do
   begin

      asg_Memo.RowCount := iMaxCheck + k;

      if Trim(mm_Log.Lines.Strings[i-1]) <> '' then
      begin
         // 하나의 Line에 붙어있는 Token 개수 확인
         iStrTokenCnt := CountChar(mm_Log.Lines.Strings[i-1], '|');

         // 만약 하나의 String 라인내 Token(|)이 1개이상 있으면, Token 수만큼 다시 분리해서 Row 단위로 Line 추가
         if iStrTokenCnt > 0 then
         begin
            // 실제 Token수 만큼의 릴리즈 내역을 row 단위로 assign
            for j := 0 to iStrTokenCnt do
            begin
               asg_Memo.Cells[0, l + j + k] := AnsiUpperCase(TokenStr(mm_Log.Lines.Strings[i-1], '|', j+1));
            end;

            iMaxCheck := iMaxCheck + iStrTokenCnt{ + 1};

            k         := k + iStrTokenCnt {+ 1};
         end
         else
         begin
            asg_Memo.Cells[0, l+k] := AnsiUpperCase(mm_Log.Lines.Strings[i-1]);

            Inc(IMaxCheck);
         end;

         Inc(l);

      end;

      Inc(i);

   end;


   asg_Memo.RowCount := asg_Memo.RowCount - 1;

   apn_Memo.Top     := 119;
   apn_Memo.Left    := 143;
   apn_Memo.Visible := True;
   apn_Memo.Collaps := True;
   apn_Memo.Collaps := False;

   asg_MemoClickCell(sender, 0, 0);


end;

//------------------------------------------------------------------------------
// [문서관리] - [개발자 릴리즈대장] - [검색기간 시작일] onKeyDown Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2016.11.18
//------------------------------------------------------------------------------
procedure TMainDlg.fmed_RegFrDtKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   // 릴리즈 내역 시작일 변경 Comp.
   if (Key = VK_RETURN) then
   begin
      // 릴리즈 내역 조회
      if IsLogonUser then
      begin
         if fcb_ReScanRelease.Checked then
            fcb_ReScanRelease.Checked := False;

         // 검색어 Clear --> 검색어 존재하는 경우, 기간만 변경해서 검색 시작 @ 2017.10.26 LSH
         //fed_Release.Text := '';

         SelGridInfo('RELEASESCAN');

         fcb_ReScanRelease.Checked := True;

         fed_Release.SetFocus;
      end;
   end;
end;


//------------------------------------------------------------------------------
// [문서관리] - [개발자 릴리즈대장] - [검색기간 종료일] onKeyDown Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2016.11.18
//------------------------------------------------------------------------------
procedure TMainDlg.fmed_RegToDtKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   // 릴리즈 내역 종료일 변경 Comp.
   if (Key = VK_RETURN) then
   begin
      // 릴리즈 내역 조회
      if IsLogonUser then
      begin
         if fcb_ReScanRelease.Checked then
            fcb_ReScanRelease.Checked := False;

         // 검색어 Clear --> 검색어 존재하는 경우, 기간만 변경해서 검색 시작 @ 2017.10.26 LSH
         //fed_Release.Text := '';

         SelGridInfo('RELEASESCAN');

         fcb_ReScanRelease.Checked := True;

         fed_Release.SetFocus;
      end;
   end;
end;

//------------------------------------------------------------------------------
// [업무분석] - [S/R조회] - [검색기간 시작일] onKeyDown Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2016.11.18
//------------------------------------------------------------------------------
procedure TMainDlg.fmed_AnalFromKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   // S/R 시작일 변경 Comp.
   if (Key = VK_RETURN) then
   begin
      // S/R 이력 조회
      if IsLogonUser then
      begin
         lb_Analysis.Caption := '▶ 조회중입니다....';

         if fcb_ReScan.Checked then
            fcb_ReScan.Checked := False;

         SelGridInfo('ANALYSIS')
      end;
   end;
end;

//------------------------------------------------------------------------------
// [업무분석] - [S/R조회] - [검색기간 종료일] onKeyDown Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2016.11.18
//------------------------------------------------------------------------------
procedure TMainDlg.fmed_AnalToKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   // S/R 종료일 변경 Comp.
   if (Key = VK_RETURN) then
   begin
      // S/R 이력 조회
      if IsLogonUser then
      begin
         lb_Analysis.Caption := '▶ 조회중입니다....';

         if fcb_ReScan.Checked then
            fcb_ReScan.Checked := False;

         SelGridInfo('ANALYSIS')
      end;
   end;
end;


//------------------------------------------------------------------------------
// SMS 자동 템플릿 목록 조회
//          - 숨쉬는 것처럼 모든 일상의 업무를 자동화 시켜버리자..
//          - 자주 전화오는 FAQ 문의들을 D/B에 넣어넣고, 구글 Photos 경로 추출하여
//            SMS 템플릿으로 관리하자.
//
// Author : Lee, Se-Ha
// Date   : 2017.07.11
//------------------------------------------------------------------------------
procedure TMainDlg.mi_Sms_TemplateClick(Sender: TObject);
begin
   apn_Memo.Caption.Text := 'SMS 템플릿 목록';

   apn_Memo.Top     := apn_SMS.Top;
   apn_Memo.Left    := apn_SMS.Left + apn_SMS.Width + 1;
   apn_Memo.Visible := True;
   apn_Memo.Collaps := True;
   apn_Memo.Collaps := False;

   asg_Memo.ClearRows(1, asg_Memo.RowCount);
   asg_Memo.RowCount := 2;

   // SMS 템플릿 목록 D/B 조회
   GetSMSTemplate;

end;

//--------------------------------------------------------------------------
// [SMS] SMS 템플릿 목록 D/B 조회
//      - 2017.07.11 LSH
//--------------------------------------------------------------------------
procedure TMainDlg.GetSMSTemplate;
var
   iTempIdx     : Integer;
   TpGetSMSTemp : TTpSvc;
begin

   Screen.Cursor := crHourglass;


   //-----------------------------------------------------------------
   // 1. Clear Grid
   //-----------------------------------------------------------------
   asg_Memo.ClearRows(1, asg_Memo.RowCount);
   asg_Memo.RowCount := 2;

   //-----------------------------------------------------------------
   // 2. Load Data by TpSvc
   //-----------------------------------------------------------------
   TpGetSMSTemp := TTpSvc.Create;
   TpGetSMSTemp.Init(Self);


   Screen.Cursor := crHourglass;


   try
      TpGetSMSTemp.CountField  := 'S_STRING1';
      TpGetSMSTemp.ShowMsgFlag := False;

      if TpGetSMSTemp.GetSvc('MD_MCOMC_L1',
                          ['S_TYPE1'  , 'J'
                         , 'S_CODE1'  , 'KDIAL'
                         , 'S_CODE2'  , 'SMSTEMPL'
                         , 'S_CODE3'  , ''
                          ],
                          [
                           'L_CNT1'    , 'iCitem06'
                         , 'L_CNT2'    , 'iCitem07'
                         , 'L_CNT3'    , 'iCitem08'
                         , 'L_CNT4'    , 'iCitem09'
                         , 'L_CNT5'    , 'iCitem10'
                         , 'L_SEQNO1'  , 'iDispseq'
                         , 'S_STRING1' , 'sComcd1'
                         , 'S_STRING2' , 'sComcd2'
                         , 'S_STRING3' , 'sComcd3'
                         , 'S_STRING4' , 'sComcdnm1'
                         , 'S_STRING5' , 'sComcdnm2'
                         , 'S_STRING6' , 'sComcdnm3'
                         , 'S_STRING7' , 'sCitem01'
                         , 'S_STRING8' , 'sCitem02'
                         , 'S_STRING9' , 'sCitem03'
                         , 'S_STRING10', 'sCitem04'
                         , 'S_STRING11', 'sCitem05'
                         , 'S_STRING12', 'sCitem11'
                         , 'S_STRING13', 'sCitem12'
                         , 'S_STRING14', 'sDeldate'
                          ]) then
      begin
         if TpGetSMSTemp.RowCount = 0 then
         begin
            Exit;
         end;
      end
      else
      begin
         ShowMessage(GetTxMsg);
         Exit;
      end;



      //----------------------------------------------------
      // 2. SMS 템플릿 목록 List 표시
      //----------------------------------------------------
      with asg_Memo do
      begin

         ColWidths[0] := 100;
         ColWidths[1] := 265;

         RowCount := TpGetSMSTemp.RowCount + FixedRows + 1;

         for iTempIdx := 0 to TpGetSMSTemp.RowCount - 1 do
         begin
            // 구분
            Cells[0, iTempIdx+FixedRows] := TpGetSMSTemp.GetOutputDataS('sComcdnm2', iTempIdx);        // 템플릿 제목
            Cells[1, iTempIdx+FixedRows] := TpGetSMSTemp.GetOutputDataS('sComcdnm3', iTempIdx);        // SMS 템플릿 Context
         end;

         // 그 외 Hidden Info 처리
         ColWidths[2] := 0;
         ColWidths[3] := 0;
         ColWidths[4] := 0;

         RowCount := RowCount - 1;

      end;


   finally
      FreeAndNil(TpGetSMSTemp);

      screen.Cursor  :=  crDefault;
   end;

end;

//------------------------------------------------------------------------------
// 개발관련 문서(릴리즈대장) 이력 삭제 (Deldate Updated)
//
// Author : Lee, Se-Ha
// Date   : 2018.06.12
//------------------------------------------------------------------------------
procedure TMainDlg.UpdReleaseDeldt;
var
   TpUpdRel  : TTpSvc;
   i, iCnt : Integer;
   vType1,
   vRegDate  ,
   vDutySpec ,
   vDutyUser ,
   vContext  ,
   vDelDate  ,
   vEditIp   : Variant;
   sType     : String;
begin


   sType := '19';

   iCnt  := 0;


   //-------------------------------------------------------------------
   // 2. Create Variables
   //-------------------------------------------------------------------

   // 선택되어 있는 것 일괄 Update
   for i:= asg_Release.Selection.Top to asg_Release.Selection.Bottom do
   begin
      //------------------------------------------------------------------
      // 2-1. Append Variables
      //------------------------------------------------------------------
      AppendVariant(vType1    ,   sType                  );
      AppendVariant(vRegDate  ,   CopyByte(Trim(asg_Release.Cells[C_RL_REGDATE  ,i]), 1, 16));
      AppendVariant(vDutySpec ,   asg_Release.Cells[C_RL_DUTYSPEC ,i]);
      AppendVariant(vDutyUser ,   asg_Release.Cells[C_RL_DUTYUSER ,i]);
      AppendVariant(vContext  ,   asg_Release.Cells[C_RL_CONTEXT  ,i]);
      AppendVariant(vDelDate  ,   FormatDateTime('yyyy-mm-dd', Date));
      AppendVariant(vEditIp   ,   FsUserIp);

      Inc(iCnt);
   end;

   if iCnt <= 0 then
      Exit;


   //-------------------------------------------------------------------
   // 3. Insert by TpSvc
   //-------------------------------------------------------------------
   TpUpdRel := TTpSvc.Create;
   TpUpdRel.Init(Self);


   Screen.Cursor := crHourGlass;


   try
      if TpUpdRel.PutSvc('MD_KUMCM_M1',
                          [
                           'S_TYPE1'   , vType1
                         , 'S_STRING8' , vEditIp
                         , 'S_STRING10', vDelDate
                         , 'S_STRING11', vDutySpec
                         , 'S_STRING13', vDutyUser
                         , 'S_STRING20', vRegDate
                         , 'S_STRING28', vContext
                         ] ) then
      begin
         MessageBox(self.Handle,
                    PChar(IntToStr(iCnt) +'건의 릴리즈내역이 정상적으로 [삭제] 되었습니다.'),
                    '[KUMC 다이얼로그] 개발문서(릴리즈대장) 업데이트 알림 ',
                    MB_OK + MB_ICONINFORMATION);


         // Refresh
         SelGridInfo('RELEASE');
      end
      else
      begin
         ShowMessageM(GetTxMsg);
      end;


   finally
      FreeAndNil(TpUpdRel);
      Screen.Cursor  := crDefault;
   end;

end;

end.






