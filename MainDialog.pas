{===============================================================================
   Program ID    : DIALOG_XE7.exe [D5 >> XE7]
   Program Name  : ����Ƿ�� ���������� ��󿬶��� �� ��������
   Program Desc. : 1. ��������� ���������� ������Ʈ�� ���� ��󿬶��� ����
                   2. �߼�/���� �� �볻�� ���� ���� (�˻�/���/����/����)
                   3. ��Ÿ Ŀ�´����̼� ����

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
   // �ֿ�������� Columns
   C_WC_LOCATE    = 0;
   C_WC_DUTYPART  = 1;
   C_WC_CONTEXT   = 2;
   C_WC_REGDATE   = 3;
   C_WC_DUTYUSER  = 4;
   C_WC_CONFIRM   = 5;

   // D/B ������(��)
   C_DBD_SECCD    = 0;
   C_DBD_SECCDNM  = 1;
   C_DBD_3RDCD    = 2;
   C_DBD_3RDCDNM  = 3;
   C_DBD_WORKINFO = 4;

   // D/B ������
   C_DBM_FSTCD    = 0;
   C_DBM_FSTCDNM  = 1;
   C_DBM_WORKINFO = 2;

   // �ٹ� �� �ް� ��Ȳ Columns
   C_DT_DUTYDATE  = 0;
   C_DT_YOIL      = 1;
   C_DT_DUTYFLAG  = 2;
   C_DT_DUTYUSER  = 3;
   C_DT_REMARK    = 4;
   C_DT_SEQNO     = 5;
   C_DT_ORGDTYUSR = 6;
   C_DT_ORGREMARK = 7;

   // ���߰��ù���(���������) Columns
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


   // �������� ID ������ Rows
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

   // ����� ������ Rows
   R_PR_USERNM    = 0;
   R_PR_DUTYPART  = 1;
   R_PR_REMARK    = 2;
   R_PR_CALLNO    = 3;
   R_PR_DUTYSCH   = 4;
   R_PR_PHOTO     = 5;
   R_PR_USERID    = 6;

   // MIS FTP ���� �ּ�
   C_MIS_FTP_IP   = '���� MIS FTP �ּ�';

   // K-Dial FTP ���� �ּ� --> ISMS ���� �ļ���ġ ���� ���߰� IP ���� @ 2017.10.31 LSH
   C_KDIAL_FTP_IP = '��ü FTP �ּ�';


   // �Խ���(Ŀ�´�Ƽ) Page MaxRow �� Block ����
   C_PAGE_MAXROW_UNIT = 20;
   C_PAGE_BLOCK_UNIT  = 10;


   // ����Ʈ Columns
   C_WK_DATE      = 0;
   C_WK_GUBUN     = 1;
   C_WK_CONTEXT   = 2;
   C_WK_STEP      = 3;


   // ��������(KUMC) Columns
   C_WR_LOCATE    = 0;
   C_WR_DUTYUSER  = 1;
   C_WR_DUTYPART  = 2;
   C_WR_REGDATE   = 3;
   C_WR_CONTEXT   = 4;
   C_WR_CONFIRM   = 5;


   // ����(S/R) �� ����Ʈ Columns
   C_AL_TITLE     = 0;
   C_AL_CONTEXT   = 1;
   C_AL_ADDINFO   = 2;


   // ����(S/R) �м� Columns
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


   // ��ũ�� ���̾� Book ����Ʈ Coloumns
   C_LL_USERNM    = 0;
   C_LL_CALLNO    = 1;
   C_LL_DUTYUSER  = 2;
   C_LL_MOBILE    = 3;
   C_LL_LINKDB    = 4;
   C_LL_LINKSEQ   = 5;
   C_LL_DUTYRMK   = 6;
   C_LL_DTYUSRID  = 7;


   // My ���̾� Book Columns
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


   // ���̾� Book �˻�����Ʈ Columns
   C_DL_LOCATE    = 0;
   C_DL_DEPTNM    = 1;
   C_DL_DEPTSPEC  = 2;
   C_DL_CALLNO    = 3;
   C_DL_DUTYUSER  = 4;
   C_DL_LINKDB    = 5;
   C_DL_LINKCNT   = 6;
   C_DL_LINKSEQ   = 7;
   C_DL_DTYUSRID  = 8;


   // ���̾� Book �ֱٰ˻� Rows
   C_DS_KEYWORD = 0;


   // flat Tab-Contrl Active Tab �ε���
   AT_DIALOG    = 0;

   { -- ��󿬶��� Tab ���տ� ���� �ּ� @ 2013.12.30 LSH
   AT_DETAIL   = 1;
   AT_MASTER   = 2;
   }

   AT_DIALBOOK = 1;
   AT_BOARD    = 2;
   AT_DOC      = 3;
   AT_ANALYSIS = 4;
   AT_WORKCONN = 5;


   // �Ĵ� �޴� Columns
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


   // Chat �ڽ� Columns
   C_CH_MYCOL     = 0;
   C_CH_TEXT      = 1;
   C_CH_YOURCOL   = 2;
   C_CH_HIDDEN    = 3;


   // Ŀ�´�Ƽ(�Խ���) Columns
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


   // ���� ���/���� Columns
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


   // ���� ��ȸ Columns
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


   // ��󿬶��� Columns
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

   // �ֱ� ������Ʈ Columns
   C_NU_EDITDATE = 0;
   C_NU_EDITTEXT = 1;
   C_NU_EDITIP   = 2;


   // ���������� Columns
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


   // ��������� Columns
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

   // ������ Lab. Columns @ 2015.04.02 LSH
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
   // Variant Array ������ �ڸ��� ��ĭ �ϳ� ����
   //    - �ҽ���ó : CComFunc.pas
   //----------------------------------------------------------------
   function AppendVariant(var in_var : Variant; in_str : String ) : Integer;


   //----------------------------------------------------------------
   // String ���� Ư�� ���ڿ� ����
   //    - �ҽ���ó : MComFunc.pas
   //----------------------------------------------------------------
   function DeleteStr(OrigStr, DelStr : String) : String;



   //----------------------------------------------------------------
   // TAdvStringGrid ����Ʈ �Ӽ� ����
   //    - 2013.08.08 LSH
   //----------------------------------------------------------------
   procedure SetPrintOptions;


   //---------------------------------------------------------------------------
   // [��ȸ] Welcome �޼��� ���� ��������
   //    - 2013.08.08 LSH
   //---------------------------------------------------------------------------
   function GetMsgInfo(in_ComCd2, in_ComCd3 : String) : String;


   //---------------------------------------------------------------------------
   // [��ȸ] �������� �˾��޼��� ���� ��������
   //    - 2013.08.22 LSH
   //---------------------------------------------------------------------------
   function GetAlarmPopup(in_Gubun, in_CheckDate : String) : String;


   //---------------------------------------------------------------------------
   // [��ȸ] �� ������ Max Seqno ����
   //    - 2013.08.13 LSH
   //---------------------------------------------------------------------------
   function GetMaxDocSeq(in_DocList, in_DocYear : String) : String;


   //---------------------------------------------------------------------------
   // [��ȸ] ����� IP �� ����
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
   // [��ȸ] ���� ����
   //    - �ҽ���ó : MComFunc.pas
   //---------------------------------------------------------------------------
   function GetDayofWeek(Adate : TDatetime; Type1 : String): String;


   //---------------------------------------------------------------------------
   // [������Ʈ] �Խ��� (Ŀ�´�Ƽ) Update
   //    - 2013.08.19 LSH
   //---------------------------------------------------------------------------
   procedure UpdateBoard(in_Type,     in_BoardSeq, in_UserIp,    in_RegDate, in_RegUser,
                         in_CateUp,   in_CateDown, in_Title,     in_Context, in_AttachNm,
                         in_HideFile, in_ServerIp, in_HeadTail,  in_HeadSeq, in_TailSeq,
                         in_AlarmFro, in_AlarmTo,  in_DelDate,   in_EditIp : String);


   //---------------------------------------------------------------------------
   // [��ȸ] �Խ��� (Ŀ�´�Ƽ) ��� ����ȸ
   //    - 2013.08.20 LSH
   //---------------------------------------------------------------------------
   procedure SetReplyList(in_HeaderRow : Integer;
                          in_BoardSeq,
                          in_UserIp,
                          in_RegDate : String);


   //---------------------------------------------------------------------------
   // [������Ʈ] Log ������Ʈ (�Խñ� ��õ ��)
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
   // �ѿ�Ű�� �ѱ�Ű�� ��ȯ
   //    - �ҽ���ó : MComFunc.pas
   //---------------------------------------------------------------------------
   procedure HanKeyChg(Handle1:THandle);



   //---------------------------------------------------------------------------
   // [��ȸ] �Խ��� (Ŀ�´�Ƽ) ��õ List ����ȸ
   //    - 2013.08.27 LSH
   //---------------------------------------------------------------------------
   procedure GetLikeList(in_HeaderRow : Integer;
                         in_BoardSeq,
                         in_UserIp,
                         in_RegDate : String);




   // Client �޼��� ����
   function SendToAllClients( s : String) : boolean;



   //---------------------------------------------------------------------------
   // FTP ���ε�
   //    - �ҽ���ó : MDN191F2.pas
   //---------------------------------------------------------------------------
   function FileUpLoad(TargetFile, DelFile: String; var S_IP: String): Boolean;



   //---------------------------------------------------------------------------
   // [�Լ�] TCP Port ���¿��� Check
   //    - 2013.08.30 LSH
   //    - �ҽ���ó : Google
   //---------------------------------------------------------------------------
   function PortTCPIsOpen(dwPort : Word; ipAddressStr:string) : boolean;


   //---------------------------------------------------------------------------
   // Data Loading Bar Controller
   //    - 2013.09.06 LSH
   //---------------------------------------------------------------------------
   procedure SetLoadingBar(AsStatus : String);


   //----------------------------------------------------------------------
   // ���� ���� Grid�� �ҷ�����
   //    - �ҽ���ó : dkChoi
   //    - 2013.09.13 LSH
   //----------------------------------------------------------------------
   procedure LoadExcelFile(in_FileName : String);



   //---------------------------------------------------------------------------
   // [��ȸ] �� �ٹ�ó�� ������ ����
   //    - 2013.09.16 LSH
   //---------------------------------------------------------------------------
   function GetDietInfo(in_Gubun, in_DayGubun : String) : String;


   //---------------------------------------------------------------------------
   // [������Ʈ] ���� ���̾� Book Update
   //    - 2013.09.26 LSH
   //---------------------------------------------------------------------------
   procedure UpdateMyDial(in_Type,     in_Locate,     in_UserIp,  in_DeptNm,  in_DeptSpec,
                          in_CallNo,   in_DutyUser,   in_Mobile,  in_DutyRmk, in_DataBase,
                          in_LinkSeq,  in_DelDate,    in_DtyUsrId : String);



   //---------------------------------------------------------------------------
   // [��ȸ] ���̾� Book ��ũ List ����ȸ
   //    - 2013.09.27 LSH
   //---------------------------------------------------------------------------
   procedure GetLinkList(in_Gubun,
                         in_DialLocate,
                         in_DialLinkDb,
                         in_DialLinkSeq,
                         in_UserIp : String);

   //---------------------------------------------------------------------------
   // User�� ������ �˶� Check
   //    - 2013.12.06 LSH
   //---------------------------------------------------------------------------
   function CheckUserAlarm(in_Gubun,
                           in_UserNm,
                           in_UserIp,
                           in_Option : String) : Boolean;

   //---------------------------------------------------------------------------
   // Panel ���� �����ϱ�
   //    - 2014.01.29 LSH
   //---------------------------------------------------------------------------
   procedure SetPanelStatus(in_Gubun,
                            in_Status : String);

   //---------------------------------------------------------------------------
   // [��ȸ] FlatComboBox ���� Code ��ȸ
   //    - 2014.02.03 LSH
   //---------------------------------------------------------------------------
   procedure GetComboBoxList(in_Type1,
                             in_Type2 : String;
                             NmCombo  : TFlatComboBox);

   //---------------------------------------------------------------------------
   // [���] ���к� ÷�� �̹��� ���
   //    - 2014.02.06 LSH
   //---------------------------------------------------------------------------
   procedure UpdateImage(in_Gubun,
                         in_FileName,
                         in_HideFile : String;
                         in_NowRow : Integer);

   //---------------------------------------------------------------------------
   // [����] Grid - Excel ��ȯ
   //    - 2014.04.07 LSH
   //---------------------------------------------------------------------------
   procedure ExcelSave(ExcelGrid : TAdvStringGrid; Title : String; Disp : Integer);


   //---------------------------------------------------------------------------
   // ���̾� Pop �޼��� �ǽð� Check
   //    - 2014.07.01 LSH
   //---------------------------------------------------------------------------
   function CheckDialPop(in_Gubun : String) : String;



  public
    //--------------------------------------------------------------------------
    // [��ȸ] ���� IP ���� User ���� ���� ��������
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
    // String ���� Ư�� ���� ����
    //    - �ҽ���ó : MComFunc.pas
    //--------------------------------------------------------------------------
    function DelChar( const Str : String ; DelC : Char ) : String;


    //--------------------------------------------------------------------------
    // [��ȸ] Grid ���Ǻ� �� procedure
    //    - 2013.08.08 LSH
    //--------------------------------------------------------------------------
    procedure SelGridInfo(in_SearchFlag : String);


    //--------------------------------------------------------------------------
    // ���߰��� ����(���������) ���� ���ε� �ϰ����
    //    - 2014.01.02 LSH
    //--------------------------------------------------------------------------
    procedure InsReleaseList;

    //--------------------------------------------------------------------------
    // ������ Lab. ���� ���ε� �ϰ����
    //    - 2015.04.02 LSH
    //--------------------------------------------------------------------------
    procedure InsBigDataList;

    //--------------------------------------------------------------------------
    // Grid �÷� ���� ����
    //    - 2014.02.03 LSH
    //--------------------------------------------------------------------------
    procedure SetGridCol(in_Flag : String);

    //--------------------------------------------------------------------------
    // [�������] FTP �̹��� ���� ��� �Լ�
    //--------------------------------------------------------------------------
    procedure KDialPrint(in_ImgCode  : String;
                         in_GridName : TAdvStringGrid);

    //--------------------------------------------------------------------------
    // [���̾�Pop] Pop-up �޼��� Anchor �̺�Ʈ
    //     - 2014.07.01 LSH
    //--------------------------------------------------------------------------
    procedure AnchorClick(sender: TObject; Anchor: string);

    //--------------------------------------------------------------------------
    // [���̾�Pop] Pop-up �޼��� �˾� �̺�Ʈ
    //     - 2014.07.01 LSH
    //--------------------------------------------------------------------------
    function CreatePopup(in_Gubun : String) : Boolean;


    //--------------------------------------------------------------------------
    // ���ڿ� ó������ KeyStr�� ������ ���������� �߶� @ 2015.03.31 LSH
    //      - �ҽ���ó : MComfunc.pas
    //--------------------------------------------------------------------------
    function FindStr(Fstr : String; KeyStr : String) : String;

    //--------------------------------------------------------------------------
    // String���� Ư�� ���ڿ��� �ٸ� ���ڿ��� ��ü @ 2015.04.02 LSH
    //      - �ҽ���ó : MComfunc.pas
    //      - StringLib.pas�� ReplaceStr�� �浹���� ������ �߰���.
    //--------------------------------------------------------------------------
    function ReplaceChar(TextStr, OrigStr, ChgStr : String) : String;

    //--------------------------------------------------------------------------
    // [�ٹ� �� ����] �ٹ�-����-�������� ������� Update
    //      - 2015.06.12 LSH
    //--------------------------------------------------------------------------
    procedure UpdateDuty(in_UpdFlag : String);

    //--------------------------------------------------------------------------
    // [SMS] SMS ���ø� ��� D/B ��ȸ
    //      - 2017.07.11 LSH
    //--------------------------------------------------------------------------
    procedure GetSMSTemplate;

    //--------------------------------------------------------------------------
    // ���߰��� ����(���������) �̷� ���� (Deldate Updated)
    //    - 2018.06.12 LSH
    //--------------------------------------------------------------------------
    procedure UpdReleaseDeldt;

  published
      // Chat - Client���� Log ������Ʈ ���� Call-back �Լ� @ 2015.03.27 LSH
      procedure SetUpdLogOut(Sender : TObject);
      procedure SetUpdChatLog(Sender : TObject);
      procedure SetUpdChatSearch(Sender : TObject);

      // ����� ������ �̹��� ���� FTP �ٿ�ε� @ 2015.03.31 LSH
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
   //AsPrev,  // AdvStrGrid ��� �̸����� ���
   TpSvc,   // TpCall Agent ������
   Excel2000,
   Imm,     // HIMC �������߰� @ 2013.08.22 LSH
   ShellApi,// ShellExec ������ �߰� @ 2013.08.22 LSH
   Winsock, // Windows Sockets API Unit
   ComObj,
   Qrctrls,
   QuickRpt,
   IdFTPCommon   // ������ XE ���������� Id_FTP --> IndyFTP ��ȯ @ 2014.04.07 LSH
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
   // ���߰� (AAOCSDEV) Param. Set.
   //----------------------------------------------------------
   begin
      sNewHospital := 'D0';
//      G_SVRID      := '01';
   end;


   //----------------------------------------------------------
   // ���߰� (�Ǵ� ���, ���� Ȯ�强 ���) ����
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




   // test : ���߻���� IP ����
   {
   if (FsUserIp = '������IP') then
      FsUserIp := '�׽�ƮIP';
   }







   //---------------------------------------------------------
   // [�ֱپ�����Ʈ] TabSheet Frame Set.
   //---------------------------------------------------------
   with asg_NwUpdate do
   begin
      ColWidths[C_NU_EDITDATE]   := 150;
      ColWidths[C_NU_EDITTEXT]   := 350;
      ColWidths[C_NU_EDITIP]     := 110;
   end;


   //---------------------------------------------------------
   // [��󿬶���] TabSheet Frame Set.
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
   // [����������] TabSheet Frame Set.
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
                 '��ü������ ����ϴ� ������ �󼼳�����,'#13'�Է��Ͻ� �� �ֽ��ϴ�.');

      AddComment(C_D_DUTYPTNR,
                 0,
                 '���������� [������Ʈ]�� ��� �ʼ��Է� �׸�����,'#13'��Ʈ�� P/L �Ǵ� �Ƿ�� �Ǵ� ��Ʈ�ʽ�(���¾�ü) ���θ� ǥ���մϴ�.');

      AddComment(C_D_DELDATE,
                 0,
                 '��󿬶������� ������ ������� ���������� ������ �����ϱ� ���Ͻø�,'#13'�ش� �������ڸ� ����Ŭ������ �������� [Update] ���ּ���.'#13'(����Ŭ���� �������� Toggle)');

      AddComment(C_D_USERADD,
                 0,
                 '�ű� ������������ �߰��ϱ� ���Ͻø� [+]�� �����ּ���.');
   end;



   //---------------------------------------------------------
   // [���������] TabSheet Frame Set.
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
                 '��󿬶������� ������ ����������� �����ϱ� ���Ͻø�,'#13'�ش� �������ڸ� ����Ŭ������ �������� [Update] ���ּ���.'#13'(����Ŭ���� �������� Toggle)');

      AddComment(C_M_USERADD,
                 0,
                 '�ű� ����ڸ� �߰��ϱ� ���Ͻø� [+]�� �����ּ���.');

      AddComment(C_M_USERIP,
                 0,
                 '������ Cell�� ����Ŭ���Ͻø�, ����ġ�� IP ������ �ڵ���ȸ �մϴ�.');
   end;




   //---------------------------------------------------------
   // [�߼�/������] �˻� �� ��� TabSheet Frame Set.
   //---------------------------------------------------------
   with asg_RegDoc do
   begin
      AddComment(C_RD_DOCYEAR,
                 0,
                 'ȸ��⵵ (��� 3��~ �ͳ� 2��)�� [yyyy] 4�ڸ��� �Է����ּ���.');

      AddComment(C_RD_DOCSEQ,
                 0,
                 '[���]�� ����Ŭ�� �Ͻø�, ������ȣ�� �ڵ����� ä���˴ϴ�.'#13'�˻�/����/������ ���� ������ȣ 3�ڸ� �Է����ּ���.');

      AddComment(C_RD_DOCLOC,
                 0,
                 '���� �������ο� ������(�Ǵ� ��������)�� ����ö�� �̸��� �����ּ���.');
   end;


   //---------------------------------------------------------
   // [�߼�/������] ��ȸ��� TabSheet Frame Set.
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
   // [Ŀ�´�Ƽ] TabSheet Frame Set.
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
                 '�Խñ� �� 1ȸ�� ��õ�� �����ϸ�,'#13'��õȽ���� 5ȸ �̻��� ��� Best �Խñ۷� 7�ϰ� Page ��ܿ� �ǵ�˴ϴ�.');
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
   // [Chat ����] Frame Set.
   //---------------------------------------------------------
   with asg_ChatSend do
   begin
      ColWidths[C_SD_TEXT]   := 318;
//      ColWidths[C_SD_BUTTON  := 227;

      AddButton(C_SD_BUTTON, 0, ColWidths[C_SD_BUTTON]-5, 20, '����', haBeforeText, vaCenter);     // ���� Button
   end;


   //---------------------------------------------------------
   // [���̾� Book �˻�] Frame Set.
   //---------------------------------------------------------
   with asg_DialList do
   begin
      ColWidths[C_DL_LINKSEQ]   := 0;
      ColWidths[C_DL_DTYUSRID]  := 0;
   end;


   //---------------------------------------------------------
   // [���� ���̾� Book] Frame Set.
   //---------------------------------------------------------
   with asg_MyDial do
   begin
      ColWidths[C_MD_LINKDB]    := 0;
      ColWidths[C_MD_LINKSEQ]   := 0;
   end;


   //---------------------------------------------------------
   // [���� ���̾� Book] Frame Set.
   //---------------------------------------------------------
   with asg_LinkList do
   begin
      ColWidths[C_LL_LINKDB]    := 0;
      ColWidths[C_LL_LINKSEQ]   := 0;
   end;


   //---------------------------------------------------------
   // [����(S/R) �м�] Frame Set.
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
   // [����(S/R) �м� ��] Frame Set.
   //---------------------------------------------------------
   with asg_AnalList do
   begin
      ColWidths[C_AL_ADDINFO]   := 0;
   end;


   //---------------------------------------------------------
   // [��������(KUMC)] Frame Set.
   //---------------------------------------------------------
   with asg_WorkRpt do
   begin
      ColWidths[C_WR_CONFIRM]   := 0;
   end;


   //---------------------------------------------------------
   // [��󿬶���] ����� ������ Set.
   //---------------------------------------------------------
   with asg_UserProfile do
   begin
      RowHeights[R_PR_PHOTO] := 132;
      RowHeights[R_PR_USERID]:= 0;
   end;

   //---------------------------------------------------------
   // [�ֿ��������] Frame Set.
   //---------------------------------------------------------
   {  -- �̻������ �ּ� @ 2017.10.31 LSH
   with asg_WorkConn do
   begin
      ColWidths[C_WC_CONFIRM]   := 0;

      fmed_ConnFrom.Text   := FormatDateTime('yyyy-mm-dd', DATE - 30);   // �� 1���� �� �̷º��� �˻�
      fmed_ConnTo.Text     := FormatDateTime('yyyy-mm-dd', DATE);
   end;
   }



   //---------------------------------------------------------
   // K-Dialog ����� User ���� ��ȸ
   //---------------------------------------------------------
   IsLogonUser := CheckUserInfo(FsUserIp, FsUserNm, FsNickNm, FsUserPart, FsUserSpec, FsUserMobile, FsUserCallNo, FsMngrNm);





   // ����� ���� Ȯ��
   if (PosByte('����',   FsUserSpec) > 0) or
      (PosByte('����',   FsUserSpec) > 0) or
      (PosByte('��Ʈ��', FsUserSpec) > 0) then
      sUserJikGup := FsUserSpec  + '��'
   else
      sUserJikGup := '������';


   // Welcome �޼��� ���� ��ȸ
   sWcMsg := GetMsgInfo('WELCOME',
                        GetDayofWeek(Date, 'HS'));


   // Version üũ ���� ��ȸ
   sVerMsg := GetMsgInfo('INFORM',
                         'VERSION');


   // Message ���� #1
   if (sVerMsg <> '') then
      sTotalMsg := sWcMsg + #13#10 + #13#10 + #13#10 + sVerMsg
   else
      sTotalMsg := sWcMsg;

   // ������ �޴� ���� ��ȸ
   sDietMsg := GetDietInfo(FormatDateTime('dd', Date),
                           GetDayofWeek(Date, 'HS'));


   // Message ���� #2
   if (sDietMsg <> '') then
      sTotalMsg := sTotalMsg + #13#10 + #13#10 + #13#10 + sDietMsg
   else
      sTotalMsg := sTotalMsg;



   // Alarm �˾��޼��� ���� ��ȸ
   sInformMsg := GetAlarmPopup('POPUP',
                               FormatDateTime('yyyy-mm-dd', DATE));


   // Message ���� #3
   if (sInformMsg <> '') then
      sTotalMsg := sTotalMsg + #13#10 + #13#10 + #13#10 + sInformMsg
   else
      sTotalMsg := sTotalMsg;


   //-----------------------------------------------------
   // ���̾�α� Tab Skin Color
   //-----------------------------------------------------
   ftc_Dialog.Color     := clWhite;
   fpn_DialBook.Color   := clWhite;
   fpn_Analysis.Color   := clWhite;
   apn_Network.Color    := clWhite;
   apn_Detail.Color     := clWhite;
   apn_Master.Color     := clWhite;
   apn_Network.Color    := clWhite;


   //-----------------------------------------------------
   // Welcome �޼��� ó��
   //-----------------------------------------------------
   if (PosByte('������IP', FsUserIp) > 0) then
   begin
      MessageBox(self.Handle,
                 PChar('ȯ���մϴ�, ' + FsUserNm + ' �����ڴ�.' + #13#10 + #13#10 + sTotalMsg),
                 '�ڡ� KUMC ���̾�α� �ڡ�',
                 MB_OK + MB_ICONINFORMATION);

      ftc_Dialog.ActiveTab := AT_DOC;

      fsbt_releaseClick(Sender);
      //fsbt_WorkRptClick(Sender);
      //fsbt_NetworkClick(sender);
      //fsbt_ComDocClick(Sender);
      //fsbt_DBMasterClick(Sender);
      //fsbt_DutyClick(Sender);

      //---------------------------------------------------
      // [������] �ְ� �Ĵ�����(-> Big-Data lab) �Է� ��� On
      //       - �Ĵܾȳ� ��� ��ü ���� @ 2017.10.26 LSH
      //---------------------------------------------------
      //fsbt_Menu.Visible := True;


   end
   else if (PosByte('����',   FsUserSpec) > 0) or
           (PosByte('����',   FsUserSpec) > 0) or
           (PosByte('��Ʈ��', FsUserSpec) > 0) then
   begin
      MessageBox(self.Handle,
                 PChar('ȯ���մϴ�, ' + FsUserNm + ' ' +  sUserJikGup + '.' + #13#10 + #13#10 + sTotalMsg),
                 'Welcome to KUMC ���̾�α�',
                 MB_OK + MB_ICONINFORMATION);

      ftc_Dialog.ActiveTab := AT_BOARD;
   end
   else if (FsUserNm <> '') and (FsUserPart <> '') then
   begin
      MessageBox(self.Handle,
                 PChar('ȯ���մϴ�, ' + FsUserNm + ' ' +  sUserJikGup + '.' + #13#10 + #13#10 + sTotalMsg),
                 'Welcome to KUMC ���̾�α�',
                 MB_OK + MB_ICONINFORMATION);


      // ������Ʈ�� [�����м�], �� �� ��Ʈ�� [Ŀ�´�Ƽ] ù Page.
      // ������Ʈ�� [�Խ���]���� ~ ^^ @ 2014.06.26 LSH
      if PosByte('����', FsUserPart) > 0 then
      begin
         ftc_Dialog.ActiveTab := AT_BOARD;
      end
      else
         ftc_Dialog.ActiveTab := AT_BOARD;

   end
   else if (FsUserNm <> '') and (FsUserPart = '') then
   begin
      MessageBox(self.Handle,
                 PChar('ȯ���մϴ�, ' + FsUserNm + ' ' +  sUserJikGup + '.' + #13#10 + #13#10 + sTotalMsg + #13#10 + #13#10 + '[����������]�� ������Ʈ ���ּž� ��󿬶����� �ڵ�ǥ��(����) �˴ϴ�.' + #13#10 + #13#10 + '�ۼ� Page�� �̵��մϴ�.'),
                 'Welcome to KUMC ���̾�α�',
                 MB_OK + MB_ICONINFORMATION);


      ftc_Dialog.ActiveTab := AT_DIALOG;

      fsbt_DetailClick(Sender);

   end
   else
   begin
      MessageBox(self.Handle,
                 PChar('ȯ���մϴ�!' + #13#10 + #13#10 + sTotalMsg + #13#10 + #13#10 + '��󿬶���(�Ǵ� ����� IP) �̵�� ���̽ó׿�.' + #13#10 + '[���������] �Է���, [����������] ������Ʈ ��Ź�帳�ϴ�.'),
                 'Welcome to KUMC ���̾�α�',
                 MB_OK + MB_ICONINFORMATION);


      ftc_Dialog.ActiveTab := AT_DIALOG;

      fsbt_MasterClick(Sender);
   end;



//-----------------------------------------------------
   // �������� Port ���� �� Enabled
   //-----------------------------------------------------
   ServerSocket.Port := StrToInt(DelChar(CopyByte(FsUserIp, LengthByte(FsUserIp)-3, 4), '.'));


   try
      ServerSocket.Active := True;

   except
       MessageBox(self.Handle,
                  PChar('�����Ͻ� IP/Port ������ �̹� ������Դϴ�.' + #13#10 + #13#10 + '������ ���α׷��� ���� �������� �ƴ��� Ȯ���� �ּ���.'),
                  PChar(Self.Caption + ' : ���α׷� �ߺ����� ����'),
                  MB_OK + MB_ICONWARNING);
   end;



   //-----------------------------------------------------
   // Caption ����� ���� ������Ʈ
   //-----------------------------------------------------
   Self.Caption := Self.Caption + ' :: ' + FsUserNm + ' ' + sUserJikGup + ', ȯ���մϴ�!';



   //-----------------------------------------------------
   // Chat �ڽ� Location Disabled
   //-----------------------------------------------------
   apn_ChatBox.Top   := -1;
   apn_ChatBox.Left  := 999;//460;   --> ����ó��.
   apn_ChatBox.Caption.Color     := $000066C9;
   apn_ChatBox.Caption.ColorTo   := clMaroon;



   //-----------------------------------------------------
   // ���� �� �̸�Ƽ�� ���۸�� Off
   //-----------------------------------------------------
   IsEmotiSend := False;
   IsFileSend  := False;



   //-----------------------------------------------------
   // Tmax ���߰� �ڵ� ���� Timer On
   //-----------------------------------------------------
   tm_TxInit.Enabled := True;





   //-----------------------------------------------------
   // ��ϵ� User ���̾�Pop �޼��� Timer On
   //-----------------------------------------------------
   if (IsLogonUser) then
   begin
      tm_DialPop.Enabled := True;

      //-----------------------------------------------------
      // ���̾� Pop �α��� �޼��� Check
      //-----------------------------------------------------
      CreatePopUp(CheckDialPop('REALTIME'));
   end
   else
      tm_DialPop.Enabled := False;


   // Test�� �Ʒ� �ּ� �� ������� ��..
   // �α� Update
   UpdateLog('START',
             CopyByte(self.Caption, PosByte('[', self.Caption) + 1, 6),
             FsUserIp,
             'R',
             '',
             FsUserNm,
             varResult
             );

   //-----------------------------------------------------
   // User�� ������ �˶� ���� [����]
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
   // User�� ������ �˶� ���� [���ΰ���] @ 2015.04.08 LSH
   //       - �̹���, Ÿ��Ʋ, ���� D/B ��������
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
   // �ֽ� Ver. ����ȭ�鿡 �ڵ� �ٿ�ε�(������Ʈ)
   //          - ������� �� version �浹 ������ �ּ� @ 2017.12.21 LSH
   //------------------------------------------------------
   {
   if (IsLogonUser) then
   begin
      sTargetTitle   := FsVersion + '.exe';

      // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
      sTargetFile    := G_HOMEDIR + 'Exe\' + sTargetTitle;


      try
         Screen.Cursor := crHourGlass;


         // �ֽ� ������ ������ ���, skip
         if FileExists(sTargetFile) and
            // Local file�� Size üũ �߰� (Local�� �ش� ���.jpg�� 0 size�� �����ϸ� JPEG Error #41 ������) @ 2015.06.03 LSH
            (CMsg.GetFileSize(sTargetFile) <>  0) then
         begin
            // ���� �ֽ������� �����ϸ� O.K.
         end
         else
         begin
            //------------------------------------------------------------------
            // Version üũ��, �� ver. �̸� �ֽ� ver. �ٿ�ε� ����
            //       - ver. üũ��, 'v' ���� ���� ���ڸ� Check ���� @ 2015.03.05 LSH
            //------------------------------------------------------------------
            if (CopyByte(self.Caption, PosByte('[', self.Caption) + 1, 6) < 'ver' + DelChar(FsVersion, 'v')) then
            begin
               //  -- ������ FTP ��������, ���߰�FTP ���� �������� �ּ�ó�� @ 2015.03.04 LSH
//               // ���� ��/�ٿ�ε带 ���� ���� ��ȸ
//               if not GetBinUploadInfo(sServerIp,
//                                       sFtpUserID,
//                                       sFtpPasswd,
//                                       sFtpHostName,
//                                       sFtpDir) then
//               begin
//                  MessageDlg('���� ������ ���� ����� ���� ��ȸ��, ������ �߻��߽��ϴ�.',
//                             Dialogs.mtError,
//                             [Dialogs.mbOK],
//                             0);
//                  Exit;
//               end;



               sServerIp   := C_KDIAL_FTP_IP;
               sFtpUserId  := 'bingo';
               sFtpPasswd  := 'bingo123';



               // ���� ������ ����Ǿ� �ִ� ���ϸ� ����
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
                     //	�������� FTP ���� �ٿ��� �ȵȰ��, Msg.
                     MessageBox(self.Handle,
                                PChar('KUMC ���̾�α� �ֽ�Ver. FTP �ٿ�ε� ������ ������ �߻��Ͽ����ϴ�.' + #13#10 + #13#10 +
                                      '���α׷� ���� �� �ٽ� ������ �ֽñ� �ٶ��ϴ�.'),
                                'KUMC ���̾�α� �ֽ�Ver. FTP �ٿ�ε� ����',
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
                                PChar('���̾�α� [' + sTargetTitle + '] �ֽŹ����� [����ȭ��]�� �ٿ�ε� �Ǿ����ϴ�.' + #13#10 + #13#10 +
                                      '������ ����� ������ Ȱ���� ������.' ),
                                'KUMC ���̾�α� �ֽ�Ver. �ڵ� �ٿ�ε� �Ϸ�',
                                MB_OK + MB_ICONINFORMATION);
                  end;


               except
                  MessageBox(self.Handle,
                             PChar('KUMC ���̾�α� �ֽ�Ver. ������Ʈ ������ ������ �߻��Ͽ����ϴ�.' + #13#10 + #13#10 +
                                   '���α׷� ���� �� �ٽ� ������ �ֽñ� �ٶ��ϴ�.'),
                             'KUMC ���̾�α� �ֽ�Ver. �ڵ� ������Ʈ ����',
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
   // ��ϵ� User ���̾� Chat ����Ʈ �˾� @ 2015.03.30 LSH
   //-----------------------------------------------------
   {  -- Chatting Ŭ���̾�Ʈ �ε��ӵ� �����۾��� ���� ����!!
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
   // Login ���� Ȯ�� ���� ���� Ȯ�强 ��� ���ܵ�
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
            G_LOCATENM := '����Ƿ�� ' + combx_locate.Text + '(����)'
         else
            G_LOCATENM := '����Ƿ�� ' + combx_locate.Text;


         G_USERIP := GetIp;

         iUserChk := mrOk;

         ccusermt.Free;
      end
      else
      begin
         ccusermt.Free;

         MakeMsg(CF_D620, D_PASSWORD, NF290); //��й�ȣ�� ��ġ�����ʽ��ϴ�.

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
// [�������������] TAdvStringGrid onButtonClick Event Handler
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
   // ���κ� �������� Update
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
                       PChar('[�ٹ�ó/�̸�/��������]�� �ʼ��Է� �׸��Դϴ�.'),
                       PChar(Self.Caption + ' : ��������� �ʼ��Է� �׸� �˸�'),
                       MB_OK + MB_ICONWARNING);

            Exit;
         end;


         if ((PosByte('�', Cells[C_M_DUTYPART, ARow]) > 0) or
             (PosByte('����', Cells[C_M_DUTYPART, ARow]) > 0) or
             (PosByte('��ȹ', Cells[C_M_DUTYPART, ARow]) > 0)) and
            ({(Cells[C_M_USERID, ARow] = '') or}   // ������Ʈ ���¾�ü ��� ���� ��� Pass @ 2013.08.28 LSH
             (Cells[C_M_MOBILE, ARow] = '') or
             (Cells[C_M_USERIP, ARow] = '')) then
         begin
            MessageBox(self.Handle,
                       PChar('�/����/��ȹ��Ʈ�� ��� ��Ȱ�� ���������� ���� [�޴���/���IP] �ʼ��Է��� ��Ź�帳�ϴ�.'),
                       PChar(Self.Caption + ' : ��������� �ʼ��Է� �׸� �˸�'),
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
                       PChar(asg_Master.Cells[C_M_USERNM, asg_Master.Row] + ' ���� ������ ���������� [������Ʈ]�Ǿ����ϴ�.'),
                       '[KUMC ���̾�α�] ��������� ������Ʈ �˸� ',
                       MB_OK + MB_ICONINFORMATION);

            // ����� Logon ���� Update
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
      asg_Master.AddButton(C_M_USERADD, ARow + 1, asg_Master.ColWidths[C_M_USERADD]-5, 20, '+', haBeforeText, vaCenter);                // ����� Add
   end;

end;




//------------------------------------------------------------------------------
// [�Լ�] Variant �Ӽ� ���� �Լ�
//       - �ҽ���ó : CComFunc.pas
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

         AddComboString('�Ƿ��');
         AddComboString('�Ⱦ�');
         AddComboString('����');
         AddComboString('�Ȼ�');
      end;

      if (ARow > 0) and (PosByte('�Ƿ��', Cells[C_M_LOCATE, ARow]) > 0) and
                        (ACol = C_M_DUTYPART) then
      begin
         AEditor := edComboList;

         ClearComboString;

         AddComboString('������');
         AddComboString('Ȩ������');
      end;


      if (ARow > 0) and (PosByte('�Ⱦ�', Cells[C_M_LOCATE, ARow]) > 0) and
                        (ACol = C_M_DUTYPART) then
      begin
         AEditor := edComboList;

         ClearComboString;

         AddComboString('���Ʈ');
         AddComboString('������Ʈ');
         AddComboString('��ȹ��Ʈ');
         AddComboString('OP');
         AddComboString('PC/AS');
         AddComboString('PACS');
      end;


      if (ARow > 0) and ((PosByte('����', Cells[C_M_LOCATE, ARow]) > 0) or
                         (PosByte('�Ȼ�', Cells[C_M_LOCATE, ARow]) > 0)) and
                        (ACol = C_M_DUTYPART) then
      begin
         AEditor := edComboList;

         ClearComboString;

         AddComboString('���Ʈ');
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
      groupBox2.Caption := '�ֱ� ������ ������Ʈ (' + FormatDateTime('yyyy-mm-dd hh:nn', Now) + ' ����)';

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
      lb_RegDoc.Caption := '�� �μ����빮��(�߼�/���� ��)�� �˻� �Ǵ� ���(����)�Ͻ� �� �ֽ��ϴ�.';

      SelGridInfo('DOC');

      ts_Dialog.Highlighted := False;
      ts_Detail.Highlighted := False;
      ts_Master.Highlighted := False;
      ts_Doc.Highlighted    := True;

   end;
   }

end;



//------------------------------------------------------------------------------
// [��ȸ] Grid ���Ǻ� �� procedure
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
   // 1. Grid �ʱ�ȭ
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
      asg_Detail.AddButton(C_D_USERADD, 1, asg_Detail.ColWidths[C_D_USERADD]-5, 20, '+', haBeforeText, vaCenter);                // ����� Add

      sType1 := '2';
      sType2 := '';
      sType3 := '';
   end
   else if (in_SearchFlag = 'MASTER') then
   begin
      asg_Master.ClearRows(1, asg_Master.RowCount);
      asg_Master.RowCount := 2;
      asg_Master.Cells[C_M_DELDATE, 1] := '';
      asg_Master.AddButton(C_M_USERADD, 1, asg_Master.ColWidths[C_M_USERADD]-5, 20, '+', haBeforeText, vaCenter);                // ����� Add

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

      // ������ IP �ĺ���, �ٹ�ó Assign
      if PosByte('�Ⱦϵ�����', FsUserIp) > 0 then
         sType3 := '�Ⱦ�'
      else if PosByte('���ε�����', FsUserIp) > 0 then
         sType3 := '����'
      else if PosByte('�Ȼ굵����', FsUserIp) > 0 then
         sType3 := '�Ȼ�';
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

      // ������ IP �ĺ���, �ٹ�ó Assign
      if PosByte('�Ⱦϵ�����', FsUserIp) > 0 then
         sType3 := '�Ⱦ�'
      else if PosByte('���ε�����', FsUserIp) > 0 then
         sType3 := '����'
      else if PosByte('�Ȼ굵����', FsUserIp) > 0 then
         sType3 := '�Ȼ�';

      sType4 := asg_RegDoc.Cells[C_RD_DOCLIST, asg_RegDoc.Row];
   end
   else if (in_SearchFlag = 'BOARD') then
   begin
      asg_Board.ClearRows(1, asg_Board.RowCount);
      asg_Board.RowCount := 2;

      sType1 := '6';
      sType2 := Trim(fcb_Page.Text);

      // ������ Page �ܿ� Row ���������� End-Rownum�� Set.
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
      {  -- Chat XE7 ���������� �ӽ� �ּ� @ 2017.10.31 LSH
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


      if (fcb_Scan.Text = '���հ˻�') then
      begin
         sType1 := '16';
         sType2 := Trim(fed_Scan.Text);

         // ������ IP �ĺ���, �ٹ�ó�� order by ���� @ 2014.07.18 LSH
         if PosByte('�Ⱦϵ�����', FsUserIp) > 0 then
            sType3 := '�Ⱦ�'
         else if PosByte('���ε�����', FsUserIp) > 0 then
            sType3 := '����'
         else if PosByte('�Ȼ굵����', FsUserIp) > 0 then
            sType3 := '�Ȼ�';

         sType4 := '';
      end
      else if (fcb_Scan.Text = '�μ���') then
      begin
         sType1 := '17';
         sType2 := Trim(fed_Scan.Text);
         sType3 := '';
         sType4 := '';
      end
      else if (fcb_Scan.Text = '�μ���') then
      begin
         sType1 := '17';
         sType2 := '';
         sType3 := Trim(fed_Scan.Text);
         sType4 := '';
      end
      else if (fcb_Scan.Text = '����ó') then
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

      if (fcb_Analysis.Text = '���հ˻�') then
      begin
         sType1 := '21';
         sType2 := '';
         sType3 := Trim(fmed_AnalFrom.Text);
         sType4 := Trim(fmed_AnalTo.Text);
         sType5 := TokenStr(fed_Analysis.Text, ' ', 1);     // ���� Ű���� �˻��� 1st Ű���� @ 2015.05.20 LSH
         sType6 := TokenStr(fed_Analysis.Text, ' ', 2);     // ���� Ű���� �˻��� 2nd Ű���� @ 2015.05.20 LSH
         sType7 := '';
      end
      else if (fcb_Analysis.Text = 'S/R����') then
      begin
         sType1 := '22';
         sType2 := Trim(fed_Analysis.Text);
         sType3 := '';
         sType4 := '';
         sType5 := '';
         sType6 := '';
         sType7 := '';
      end
      else if (fcb_Analysis.Text = '�μ�') then
      begin
         sType1 := '22';
         sType2 := '';
         sType3 := Trim(fed_Analysis.Text);
         sType4 := '';
         sType5 := '';
         sType6 := '';
         sType7 := '';
      end
      else if (fcb_Analysis.Text = '��û��') then
      begin
         sType1 := '22';
         sType2 := '';
         sType3 := '';
         sType4 := Trim(fed_Analysis.Text);
         sType5 := '';
         sType6 := '';
         sType7 := '';
      end
      else if (fcb_Analysis.Text = '�����') then
      begin
         sType1 := '22';
         sType2 := '';
         sType3 := '';
         sType4 := '';
         sType5 := Trim(fed_Analysis.Text);
         sType6 := '';
         sType7 := '';
      end
      else if (fcb_Analysis.Text = 'ó������') then
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
      else if (fcb_WorkArea.Text = '�Ⱦ�') then
         sType2 := 'AA'
      else if (fcb_WorkArea.Text = '����') then
         sType2 := 'GR'
      else if (fcb_WorkArea.Text = '�Ȼ�') then
         sType2 := 'AS';

      sType3 := Trim(fmed_AnalFrom.Text);
      sType4 := Trim(fmed_AnalTo.Text);

      {
      if (fcb_WorkGubun.Text = '') then
         sType5 := ''
      else if (fcb_WorkGubun.Text = '�') then
         sType5 := '���Ʈ'
      else if (fcb_WorkGubun.Text = '����') then
         sType5 := '������Ʈ'
      else if (fcb_WorkGubun.Text = '��ȹ') then
         sType5 := '��ȹ��Ʈ'
      else if (fcb_WorkGubun.Text = '����') then
         sType5 := FsUserNm;
      }

      sType5 := Trim(fed_Analysis.Text);


   end
   else if (in_SearchFlag = 'WEEKLYRPT') then
   begin
      asg_WeeklyRpt.ClearRows(1, asg_WeeklyRpt.RowCount);
      asg_WeeklyRpt.RowCount := 2;


      // Title
      apn_Weekly.Caption.Text := Trim(fed_Analysis.Text) + '�� ��������Ʈ';

      sType1 := '25';

      // ������ IP �ĺ���, �ٹ�ó Assign
      if PosByte('�Ⱦ�', Trim(fcb_WorkArea.Text)) > 0 then
      begin
         sType2 := 'AA';
      end
      else if PosByte('����', Trim(fcb_WorkArea.Text)) > 0 then
      begin
         sType2 := 'GR';
      end
      else if PosByte('�Ȼ�', Trim(fcb_WorkArea.Text)) > 0 then
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
   // �ٹ�-���� ǥ ����
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

      if (fcb_WorkArea.Text = '�Ⱦ�') then
         tmpLocate := '@AAOCS'
      else if (fcb_WorkArea.Text = '����') then
         tmpLocate := '@GROCS'
      else if (fcb_WorkArea.Text = '�Ȼ�') then
         tmpLocate := '@ASOCS';

      if (fcb_DutyPart.Text = '����') then
         tmpDBLink := tmpLocate + '_ORAA1_REAL'
      else if (fcb_DutyPart.Text = '����') then
         tmpDBLink := tmpLocate + '_ORAM1_REAL'
      else if (fcb_DutyPart.Text = '��������') then
         tmpDBLink := tmpLocate + '_ORAS1_REAL'
      else if (fcb_DutyPart.Text = '�������') then
         tmpDBLink := tmpLocate + '_ORAC1_REAL'
      else if (fcb_DutyPart.Text = 'e-Refer') then
         tmpDBLink := tmpLocate + '_ORAE1'
      else if (fcb_DutyPart.Text = 'HRM') then
         tmpDBLink := '@AAINT_HRMINF'
      else if (fcb_DutyPart.Text = '�Ϲݰ���') then
         tmpDBLink := '@AAINT_ORAG1';

      if (fcb_WorkGubun.Text = '�����ڵ�') then
      begin
         if (fcb_DutyPart.Text = '����') then
         begin
            tmpFlag        := 'A';
            tmpRefTable    := 'CCCOMCDT' + tmpDBLink;
            tmpFstColNm    := 'LARGCD';
            tmpFstColValue := '0000';
         end
         else if (fcb_DutyPart.Text = '����') then
         begin
            tmpFlag        := 'M';
            tmpRefTable    := 'MDMCOMCT' + tmpDBLink;
            tmpFstColNm    := 'COMCD1';
            tmpFstColValue := '000';
         end
         else if (fcb_DutyPart.Text = '��������') then
         begin
            tmpFlag        := 'S';
            tmpRefTable    := 'SDCOMCDT' + tmpDBLink;
            tmpFstColNm    := 'LARGCD';
            tmpFstColValue := 'SD';
         end
         else if (fcb_DutyPart.Text = '�Ϲݰ���') then
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
         if (fcb_DutyPart.Text = '����') then
            tmpFlag        := 'ORAA1'
         else if (fcb_DutyPart.Text = '����') then
            tmpFlag        := 'ORAM1'
         else if (fcb_DutyPart.Text = '��������') then
            tmpFlag        := 'ORAS1'
         else if (fcb_DutyPart.Text = '�������') then
            tmpFlag        := 'ORAC1'
         else if (fcb_DutyPart.Text = 'e-Refer') then
            tmpFlag        := 'ORAE1'
         else if (fcb_DutyPart.Text = 'HRM') then
            tmpFlag        := 'HRMINF'
         else if (fcb_DutyPart.Text = '�Ϲݰ���') then
            tmpFlag        := 'ORAG1';


         tmpRefTable    := 'ALL_OBJECTS' + tmpDBLink + ' a';

         if (fcb_DutyPart.Text = 'HRM') or
            (fcb_DutyPart.Text = '�Ϲݰ���') then
            tmpFstColNm    := 'USER_TAB_COMMENTS' + tmpDBLink + ' b'
         else
            tmpFstColNm    := 'USER_TAB_COMMENTS' + tmpLocate + '_' + tmpFlag + ' b';

         tmpFstColValue := AnsiUpperCase(fcb_WorkGubun.Text);

         sType1 := '33';
         sType6 := 'TABLE_NAME';

      end
      else if (fcb_WorkGubun.Text = 'View') then
      begin
         if (fcb_DutyPart.Text = '����') then
            tmpFlag        := 'ORAA1'
         else if (fcb_DutyPart.Text = '����') then
            tmpFlag        := 'ORAM1'
         else if (fcb_DutyPart.Text = '��������') then
            tmpFlag        := 'ORAS1'
         else if (fcb_DutyPart.Text = '�������') then
            tmpFlag        := 'ORAC1'
         else if (fcb_DutyPart.Text = 'e-Refer') then
            tmpFlag        := 'ORAE1'
         else if (fcb_DutyPart.Text = 'HRM') then
            tmpFlag        := 'HRMINF'
         else if (fcb_DutyPart.Text = '�Ϲݰ���') then
            tmpFlag        := 'ORAG1';


         tmpRefTable    := 'ALL_OBJECTS' + tmpDBLink + ' a';

         if (fcb_DutyPart.Text = 'HRM') or
            (fcb_DutyPart.Text = '�Ϲݰ���') then
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
         if (fcb_DutyPart.Text = '����') then
            tmpFlag        := 'ORAA1'
         else if (fcb_DutyPart.Text = '����') then
            tmpFlag        := 'ORAM1'
         else if (fcb_DutyPart.Text = '��������') then
            tmpFlag        := 'ORAS1'
         else if (fcb_DutyPart.Text = '�������') then
            tmpFlag        := 'ORAC1'
         else if (fcb_DutyPart.Text = 'e-Refer') then
            tmpFlag        := 'ORAE1'
         else if (fcb_DutyPart.Text = 'HRM') then
            tmpFlag        := 'HRMINF'
         else if (fcb_DutyPart.Text = '�Ϲݰ���') then
            tmpFlag        := 'ORAG1';


         tmpRefTable    := 'ALL_OBJECTS' + tmpDBLink + ' a';

         if (fcb_DutyPart.Text = 'HRM') or
            (fcb_DutyPart.Text = '�Ϲݰ���')then
            tmpFstColNm    := 'USER_PROCEDURES' + tmpDBLink + ' b'
         else
            tmpFstColNm    := 'USER_PROCEDURES' + tmpLocate + '_' + tmpFlag + ' b';

         tmpFstColValue := AnsiUpperCase(fcb_WorkGubun.Text);

         sType1 := '34';
         sType6 := 'OBJECT_NAME';
      end
      else if (fcb_WorkGubun.Text = 'Slip�ڵ�') then
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

      // Grid �÷� ���� ����
      SetGridCol(AnsiUpperCase(fcb_WorkGubun.Text));

   end
   else if (in_SearchFlag = 'DBCATEDET') then
   begin
      asg_DBDetail.ClearRows(1, asg_DBDetail.RowCount);
      asg_DBDetail.RowCount := 2;

      if (fcb_WorkArea.Text = '�Ⱦ�') then
         tmpLocate := '@AAOCS'
      else if (fcb_WorkArea.Text = '����') then
         tmpLocate := '@GROCS'
      else if (fcb_WorkArea.Text = '�Ȼ�') then
         tmpLocate := '@ASOCS';

      if (fcb_DutyPart.Text = '����') then
         tmpDBLink := tmpLocate + '_ORAA1_REAL'
      else if (fcb_DutyPart.Text = '����') then
         tmpDBLink := tmpLocate + '_ORAM1_REAL'
      else if (fcb_DutyPart.Text = '��������') then
         tmpDBLink := tmpLocate + '_ORAS1_REAL'
      else if (fcb_DutyPart.Text = '�������') then
         tmpDBLink := tmpLocate + '_ORAC1_REAL'
      else if (fcb_DutyPart.Text = 'e-Refer') then
         tmpDBLink := tmpLocate + '_ORAE1'
      else if (fcb_DutyPart.Text = '�Ϲݰ���') then
         tmpDBLink := '@AAINT_ORAG1'
      else if (fcb_DutyPart.Text = 'HRM') then
         tmpDBLink := '@AAINT_HRMINF';

      if (fcb_WorkGubun.Text = '�����ڵ�') then
      begin
         if (fcb_DutyPart.Text = '����') then
         begin
            tmpFlag        := 'A';
            tmpRefTable    := 'CCCOMCDT' + tmpDBLink;
            tmpFstColNm    := 'LARGCD';
            tmpFstColValue := Trim(asg_DBMaster.Cells[0, asg_DBMaster.Row]);
         end
         else if (fcb_DutyPart.Text = '����') then
         begin
            tmpFlag        := 'M';
            tmpRefTable    := 'MDMCOMCT' + tmpDBLink;
            tmpFstColNm    := 'COMCD1';
            tmpFstColValue := Trim(asg_DBMaster.Cells[0, asg_DBMaster.Row]);
         end
         else if (fcb_DutyPart.Text = '��������') then
         begin
            tmpFlag        := 'S';
            tmpRefTable    := 'SDCOMCDT' + tmpDBLink;
            tmpFstColNm    := 'LARGCD';
            tmpFstColValue := Trim(asg_DBMaster.Cells[0, asg_DBMaster.Row]);
         end
         else if (fcb_DutyPart.Text = '�Ϲݰ���') then
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
      else if (fcb_WorkGubun.Text = 'Slip�ڵ�') then
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

      // Grid �÷� ���� ����
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

      if (fcb_WorkArea.Text = '�Ⱦ�') then
         tmpLocate := '@AAOCS'
      else if (fcb_WorkArea.Text = '����') then
         tmpLocate := '@GROCS'
      else if (fcb_WorkArea.Text = '�Ȼ�') then
         tmpLocate := '@ASOCS';

      if (fcb_DutyPart.Text = '����') then
         tmpDBLink := tmpLocate + '_ORAA1_REAL'
      else if (fcb_DutyPart.Text = '����') then
         tmpDBLink := tmpLocate + '_ORAM1_REAL'
      else if (fcb_DutyPart.Text = '��������') then
         tmpDBLink := tmpLocate + '_ORAS1_REAL'
      else if (fcb_DutyPart.Text = '�������') then
         tmpDBLink := tmpLocate + '_ORAC1_REAL'
      else if (fcb_DutyPart.Text = 'e-Refer') then
         tmpDBLink := tmpLocate + '_ORAE1'
      else if (fcb_DutyPart.Text = 'HRM') then
         tmpDBLink := '@AAINT_HRMINF'
      else if (fcb_DutyPart.Text = '�Ϲݰ���') then
         tmpDBLink := '@AAINT_ORAG1';




      if (fcb_DutyPart.Text = '����') then
         tmpFlag        := 'ORAA1'
      else if (fcb_DutyPart.Text = '����') then
         tmpFlag        := 'ORAM1'
      else if (fcb_DutyPart.Text = '��������') then
         tmpFlag        := 'ORAS1'
      else if (fcb_DutyPart.Text = '�������') then
         tmpFlag        := 'ORAC1'
      else if (fcb_DutyPart.Text = 'e-Refer') then
         tmpFlag        := 'ORAE1'
      else if (fcb_DutyPart.Text = 'HRM') then
         tmpFlag        := 'HRMINF'
      else if (fcb_DutyPart.Text = '�Ϲݰ���') then
         tmpFlag        := 'ORAG1';


      tmpRefTable    := 'all_objects' + tmpDBLink + ' a';


      if (fcb_DutyPart.Text = 'HRM') or
         (fcb_DutyPart.Text = '�Ϲݰ���') then
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


      // Grid �÷� ���� ����
      SetGridCol(in_SearchFlag);

   end
   else if (in_SearchFlag = 'SRHISTORY') then
   begin
      asg_WeeklyRpt.ClearRows(1, asg_WeeklyRpt.RowCount);
      asg_WeeklyRpt.RowCount := 2;

      // Title
      apn_Weekly.Caption.Text := Trim(fed_Analysis.Text) + '�� S/R ��û��Ȳ';

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
   // �����ٹ�-����(Ȯ��) �̷� ��ȸ
   else if (in_SearchFlag = 'DUTYSEARCH') then
   begin
      asg_Duty.ClearRows(1, asg_Duty.RowCount);
      asg_Duty.RowCount := 2;

      sType1 := '45';
      sType2 := Trim(fmed_DutyFrDt.Text);
      sType3 := Trim(fmed_DutyToDt.Text);
   end;




   //-----------------------------------------------------------------
   // Dynamic Column ���� ���� Index �ʱ�ȭ
   //-----------------------------------------------------------------
   iLstRowCnt := 0;





   //-----------------------------------------------------------------
   // 2. ��ȸ
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
            lb_NetWork.Caption   := '�� ��ȸ������ �������� �ʽ��ϴ�.';
            lb_SelDoc.Caption    := '�� ��ȸ������ �������� �ʽ��ϴ�.';
            lb_Board.Caption     := '�� ��ȸ������ �������� �ʽ��ϴ�.';
            lb_DialScan.Caption  := '�� ��ȸ������ �������� �ʽ��ϴ�.';
            lb_MyDial.Caption    := '';
            lb_Analysis.Caption  := '�� ��ȸ������ �������� �ʽ��ϴ�.';
            //lb_WorkConn.Caption  := '�� ��ȸ������ �������� �ʽ��ϴ�.';


            if (in_SearchFlag = 'RELEASE') or
               (in_SearchFlag = 'RELEASESCAN') then
               lb_RegDoc.Caption    := '�� ��ȸ������ �������� �ʽ��ϴ�.';

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


            // Merge ����
            iBaseRow       := 0;    // ��з� Merge Start Index
            iRowSpan       := 1;    // ��з� Span  Index
            jBaseRow       := 0;    // �ߺз� Merge Start Index
            jRowSpan       := 1;    // �ߺз� Span  Index
            kBaseRow       := 0;    // �Һз� Merge Start Index
            kRowSpan       := 1;    // �Һз� Span  Index

            sPrevUptitle    := '';   // ��з� ���� Title
            sPrevMidtitle   := '';   // �ߺз� ���� Title
            sPrevDowntitle  := '';   // �Һз� ���� Title




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
               Cells[C_NW_USERID,   RowCount - 1] := TpGetSearch.GetOutputDataS('sUserId'    , i);     // ����� ID �߰� @ 2014.06.26 LSH

               //---------------------------------------------------------------
               // �����Ϸ��� Server ��밡�� ���� (���� Session On/Off) Check
               //---------------------------------------------------------------
               if (TpGetSearch.GetOutputDataS('sLogYn', i) = 'Y') then
                  FontColors[C_NW_USERNM, RowCount - 1] := $0000CDA3
               else
                  FontColors[C_NW_USERNM, RowCount - 1] := clBlack;


               // ���� ������ User�� Chat-List ��ġ�� ����صд� @ 2015.03.30 LSH
               if TpGetSearch.GetOutputDataS('sUserNm', i) = FsUserNm then
                  iUserRowId := RowCount - 1;

               // �ش� ȭ�� 1-Cycle �ڵ��� ������ Index�� ���� RowCount �� ���
               if (i = iRowCnt) then
                  iLstRowCnt := i+1
               else
                  iLstRowCnt := iRowCnt;



               // ��з� �׸� MergeCell üũ
               if TpGetSearch.GetOutputDataS('sLocate', i) <> sPrevUptitle then
               begin
                  // test
                  //Memo1.Lines.Add('--- Test Logging Start ---');


                  // ���� ��з��� �ٸ� ���
                  if i > 0 then
                  begin
                     MergeCells(0, iBaseRow, 1, iRowSpan);
                     iRowSpan := 1;
                  end;

                  // �� Key Point
                  iBaseRow := RowCount - 1;


                  // �ߺз� �׸� MergeCell üũ
                  if TpGetSearch.GetOutputDataS('sDutyPart', i) <> sPrevMidTitle then
                  begin
                     // ���� �ߺз��� �ٸ� ���
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;


                     // �� Key Point
                     jBaseRow := RowCount - 1;



                     // �Һз� �׸� MergeCell üũ
                     if TpGetSearch.GetOutputDataS('sDutySpec', i) <> sPrevDownTitle then
                     begin
                        // ���� �Һз��� �ٸ� ���
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // �� Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // ���� �Һз��� �ٸ� ���
                        Inc(kRowSpan);


                  end
                  else
                  begin
                     // ���� �ߺз��� ���� ���
                     Inc(jRowSpan);
                  end;
               end
               else
               begin

                  // �ߺз� �׸� MergeCell üũ
                  if TpGetSearch.GetOutputDataS('sDutyPart', i) <> sPrevMidTitle then
                  begin
                     // ���� �ߺз��� �ٸ� ���
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;

                     // �� Key Point
                     jBaseRow := RowCount - 1;


                     // �Һз� �׸� MergeCell üũ
                     if TpGetSearch.GetOutputDataS('sDutySpec', i) <> sPrevDownTitle then
                     begin
                        // ���� �Һз��� ���� ���
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // �� Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // ���� �Һз��� �ٸ� ���
                        Inc(kRowSpan);
                  end
                  else
                  begin
                     // �ߺз� span Index ++
                     Inc(jRowSpan);


                     // �Һз� �׸� MergeCell üũ
                     if TpGetSearch.GetOutputDataS('sDutySpec', i) <> sPrevDownTitle then
                     begin
                        // ���� �Һз��� �ٸ� ���
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // �� Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // ���� �Һз��� ������ ���
                        Inc(kRowSpan);

                  end;


                  // ��з� Span index ++
                  Inc(iRowSpan);

               end;


               // Set previous value
               sPrevUptitle   := TpGetSearch.GetOutputDataS('sLocate'  , i);
               sPrevmidtitle  := TpGetSearch.GetOutputDataS('sDutyPart', i);
               sPrevDowntitle := TpGetSearch.GetOutputDataS('sDutySpec', i);


               // RowCount �����ٷ� �ѱ�.
               RowCount := RowCount + 1;

            end;


            // �� ���� Code List-up�� ������ Row Merge ó��
            asg_NetWork.MergeCells(0, iBaseRow, 1, iRowSpan);
            asg_NetWork.MergeCells(1, jBaseRow, 1, jRowSpan);
            asg_NetWork.MergeCells(2, kBaseRow, 1, kRowSpan);
            asg_NetWork.RowCount := iLstRowCnt + 1;

            // Comments
            lb_Network.Caption := '�� [���������]�� [����������]�� ������Ʈ �Ͻø� ��󿬶����� �ڵ���ȸ �˴ϴ�.';


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
                                    TpGetSearch.GetOutputDataS('sUserNm',   i) + '�� "' +
                                    TpGetSearch.GetOutputDataS('sFlag'  ,   i) + '" ������Ʈ��'
                  else
                     tmpEditTxt :=  TpGetSearch.GetOutputDataS('sLocate',   i) + ' '   +
                                    TpGetSearch.GetOutputDataS('sDutyPart', i) + ' '   +
                                    TpGetSearch.GetOutputDataS('sUserNm',   i) + '�� "' +
                                    TpGetSearch.GetOutputDataS('sFlag'  ,   i) + '" ����(����)��';

                  Cells[C_NU_EDITDATE, i+FixedRows] := TpGetSearch.GetOutputDataS('sEditDate', i);
                  Cells[C_NU_EDITTEXT, i+FixedRows] := tmpEditTxt;
                  Cells[C_NU_EDITIP,   i+FixedRows] := TpGetSearch.GetOutputDataS('sEditIp'  , i);
               end;
            end;
         end;

         // RowCount ����
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

                  AddButton(C_D_USERADD, i+FixedRows, ColWidths[C_D_USERADD]-5, 20, '+', haBeforeText, vaCenter);                // ����� Add
               end;
            end;
         end;


         // RowCount ����
         asg_Detail.RowCount := asg_Detail.RowCount - 1;

         // Comments
         lb_Network.Caption := '�� ' + IntToStr(iRowCnt) + '���� ������������ ��ȸ�Ͽ����ϴ�.';


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


                  AddButton(C_M_USERADD, i+FixedRows, ColWidths[C_M_USERADD]-5, 20, '+', haBeforeText, vaCenter);                // ����� Add



                  //------------------------------------------------------------
                  // �����Ϸ��� Server ��밡�� ���� (���� Session On/Off) Check
                  //------------------------------------------------------------
                  if (TpGetSearch.GetOutputDataS('sLogYn', i) = 'Y') then
                     Colors[C_M_USERNM, i+FixedRows] := $0000CDA3
                  else
                     Colors[C_M_USERNM, i+FixedRows] := clWhite;

               end;
            end;
         end;


         // RowCount ����
         asg_Master.RowCount := asg_Master.RowCount - 1;

         // Comments
         lb_Network.Caption := '�� ' + IntToStr(iRowCnt) + '���� ����������� ��ȸ�Ͽ����ϴ�.';


      end
      else if (in_SearchFlag = 'DOC') or
              (in_SearchFlag = 'DOCREFRESH') then
      begin
         with asg_SelDoc do
         begin
            iRowCnt  := TpGetSearch.RowCount;
            //RowCount := iRowCnt + FixedRows + 1;


            // Merge ����
            iBaseRow       := 0;    // ��з� Merge Start Index
            iRowSpan       := 1;    // ��з� Span  Index
            jBaseRow       := 0;    // �ߺз� Merge Start Index
            jRowSpan       := 1;    // �ߺз� Span  Index
            kBaseRow       := 0;    // �Һз� Merge Start Index
            kRowSpan       := 1;    // �Һз� Span  Index

            sPrevUptitle    := '';   // ��з� ���� Title
            sPrevMidtitle   := '';   // �ߺз� ���� Title
            sPrevDowntitle  := '';   // �Һз� ���� Title


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

               // ��ุ��� �׸��� FontColor = clRed ������Ʈ @ 2014.06.02 LSH
               if (Cells[C_SD_DOCLIST, RowCount - 1] = '���ó') then
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


               // �ش� ȭ�� 1-Cycle �ڵ��� ������ Index�� ���� RowCount �� ���
               if (i = iRowCnt) then
                  iLstRowCnt := i+1
               else
                  iLstRowCnt := iRowCnt;




               // ��з� �׸� MergeCell üũ
               if TpGetSearch.GetOutputDataS('sDocList', i) <> sPrevUptitle then
               begin
                  // test
                  //Memo1.Lines.Add('--- Test Logging Start ---');


                  // ���� ��з��� �ٸ� ���
                  if i > 0 then
                  begin
                     MergeCells(0, iBaseRow, 1, iRowSpan);
                     iRowSpan := 1;
                  end;

                  // �� Key Point
                  iBaseRow := RowCount - 1;


                  // �ߺз� �׸� MergeCell üũ

                     // ���� �ߺз��� �ٸ� ���
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;


                     // �� Key Point
                     jBaseRow := RowCount - 1;



                     // �Һз� �׸� MergeCell üũ
                     if TpGetSearch.GetOutputDataS('sDocTitle', i) <> sPrevDownTitle then
                     begin
                        // ���� �Һз��� �ٸ� ���
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // �� Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // ���� �Һз��� �ٸ� ���
                        Inc(kRowSpan);

               end
               else
               begin

                  // �ߺз� �׸� MergeCell üũ
                  if TpGetSearch.GetOutputDataS('sDocSeq', i) <> sPrevMidTitle then
                  begin
                     // ���� �ߺз��� �ٸ� ���
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;

                     // �� Key Point
                     jBaseRow := RowCount - 1;


                     // �Һз� �׸� MergeCell üũ
                     if TpGetSearch.GetOutputDataS('sDocTitle', i) <> sPrevDownTitle then
                     begin
                        // ���� �Һз��� ���� ���
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // �� Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // ���� �Һз��� �ٸ� ���
                        Inc(kRowSpan);
                  end
                  else
                  begin
                     // �ߺз� span Index ++
                     Inc(jRowSpan);


                     // �Һз� �׸� MergeCell üũ
                     if TpGetSearch.GetOutputDataS('sDocTitle', i) <> sPrevDownTitle then
                     begin
                        // ���� �Һз��� �ٸ� ���
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // �� Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // ���� �Һз��� ������ ���
                        Inc(kRowSpan);

                  end;


                  // ��з� Span index ++
                  Inc(iRowSpan);

               end;


               // Set previous value
               sPrevUptitle   := TpGetSearch.GetOutputDataS('sDocList',   i);
               sPrevmidtitle  := TpGetSearch.GetOutputDataS('sDocSeq',   i);
               sPrevDowntitle := TpGetSearch.GetOutputDataS('sDocTitle', i);


               // RowCount �����ٷ� �ѱ�.
               RowCount := RowCount + 1;


            end;


            // �� ���� Code List-up�� ������ Row Merge ó��
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

         // RowCount ����
         asg_SelDoc.RowCount := asg_SelDoc.RowCount - 1;
         }

         end;


         // Comments
         if (asg_RegDoc.Cells[C_RD_DOCYEAR, asg_RegDoc.Row] = '') or
            (asg_RegDoc.Cells[C_RD_DOCLIST, asg_RegDoc.Row] = '') then
            lb_SelDoc.Caption := '�� ' + IntToStr(iRowCnt) + '���� �� ȸ��⵵ ��ü ������ �˻��Ͽ����ϴ�.'
         else
            lb_SelDoc.Caption := '�� ' + IntToStr(iRowCnt) + '���� [' + asg_RegDoc.Cells[C_RD_DOCYEAR, asg_RegDoc.Row] + ' ' +
                                                                   asg_RegDoc.Cells[C_RD_DOCLIST, asg_RegDoc.Row] + '] ' +
                                                             '������ �˻��Ͽ����ϴ�.';

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
                  // Reply ī��Ʈ ǥ��
                  //------------------------------------------------------------
                  if TpGetSearch.GetOutputDataS('sReplyCnt', i) > '0' then
                     Cells[C_B_TITLE,  i+FixedRows] := TpGetSearch.GetOutputDataS('sDocTitle', i) + '  [' + TpGetSearch.GetOutputDataS('sReplyCnt', i) + ']'
                  else
                     Cells[C_B_TITLE,  i+FixedRows] := TpGetSearch.GetOutputDataS('sDocTitle', i);


                  //------------------------------------------------------------
                  // ÷�� ���� ǥ��
                  //------------------------------------------------------------
                  if TpGetSearch.GetOutputDataS('sAttachNm', i) <> '' then
                     AddButton(C_B_ATTACH,  i+FixedRows, ColWidths[C_B_ATTACH]-5,  20, '��'{'��','��'}, haBeforeText, vaCenter);


                  //------------------------------------------------------------
                  // ���� ������ Cell Index ���� (��� ���� ������� Set)
                  //------------------------------------------------------------
                  if (i = iRowCnt - 1) then
                  begin
                     Cells[C_B_CLOSEYN,   i+FixedRows] := 'E';
                     iNowBoardCnt := i + 1;
                  end
                  else
                     Cells[C_B_CLOSEYN,   i+FixedRows] := 'C';


                  //------------------------------------------------------------
                  // �������� Coloring
                  //------------------------------------------------------------
                  if (Cells[C_B_HEADTAIL,  i+FixedRows] = 'A') then
                     for k := C_B_BOARDSEQ to C_B_ATTACH do
                        asg_Board.Colors[j, i+FixedRows]   := $00BDABFF;


                  //------------------------------------------------------------
                  // Best ��õ�� Coloring
                  //------------------------------------------------------------
                  if PosByte('[��BEST��]', Cells[C_B_TITLE,  i+FixedRows]) > 0 then
                     for k := C_B_BOARDSEQ to C_B_ATTACH do
                        asg_Board.Colors[j, i+FixedRows]   := $00D5D7AF;

               end;
            end;
         end;


         // RowCount ����
         asg_Board.RowCount := asg_Board.RowCount - 1;

         // Comments
         lb_Board.Caption := '�� ' + IntToStr(iRowCnt) + '���� �Խñ��� ��ȸ�Ͽ����ϴ�.';


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

         // ��/�� �Խñ� Paging �˻� ��ư
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
      // --> ChatUser XE7 ���������� ������..
      else if (in_SearchFlag = 'CHATUSER') then
      begin
         {  -- XE7 ���������� �ӽ��ּ� @ 2017.10.31 LSH
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


                  // ���� ������ IP �� Port ����
                  if (Cells[1, i+FixedRows] <> '') then
                  begin
                     sTargetServerIp    := Cells[1, i+FixedRows];
                     iTargetServerPort  := StrToInt(DelChar(CopyByte(Cells[1, i+FixedRows], LengthByte(Cells[1, i+FixedRows])-3, 4), '.'));


                     //---------------------------------------------------------
                     // �����Ϸ��� Server Port ��밡�� ���� Check
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

            // RowCount ����
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
                  // ���� Ű���� ��Ī ���� @ 2015.06.03 LSH
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
                  Cells[C_DL_DTYUSRID,    i+FixedRows]   := TpGetSearch.GetOutputDataS('sUserId',  i);   // ����� ID �߰� @ 2014.06.26 LSH


               end;
            end;


            // RowCount ����
            RowCount := RowCount - 1;


            // Comments
            lb_DialScan.Caption  := '�� ' + IntToStr(iRowCnt) + '���� ����ó ������ ��ȸ�Ͽ����ϴ�.';


         end;
      end
      else if (in_SearchFlag = 'MYDIAL') then
      begin
          with asg_MyDial do
          begin
            iRowCnt  := TpGetSearch.RowCount;
            //RowCount := iRowCnt + FixedRows + 1;


            // Merge ����
            iBaseRow       := 0;    // ��з� Merge Start Index
            iRowSpan       := 1;    // ��з� Span  Index
            jBaseRow       := 0;    // �ߺз� Merge Start Index
            jRowSpan       := 1;    // �ߺз� Span  Index
            kBaseRow       := 0;    // �Һз� Merge Start Index
            kRowSpan       := 1;    // �Һз� Span  Index

            sPrevUptitle    := '';   // ��з� ���� Title
            sPrevMidtitle   := '';   // �ߺз� ���� Title
            sPrevDowntitle  := '';   // �Һз� ���� Title


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
               Cells[C_MD_DTYUSRID,  RowCount - 1]  := TpGetSearch.GetOutputDataS('sUserId',     i);  // ��ϵ� UserId ���� @ 2015.04.13 LSH



               //AutoSizeRow(RowCount - 1);



               // �ش� ȭ�� 1-Cycle �ڵ��� ������ Index�� ���� RowCount �� ���
               if (i = iRowCnt) then
                  iLstRowCnt := i+1
               else
                  iLstRowCnt := iRowCnt;




               // ��з� �׸� MergeCell üũ
               if TpGetSearch.GetOutputDataS('sLocate', i) <> sPrevUptitle then
               begin
                  // test
                  //Memo1.Lines.Add('--- Test Logging Start ---');


                  // ���� ��з��� �ٸ� ���
                  if i > 0 then
                  begin
                     MergeCells(0, iBaseRow, 1, iRowSpan);
                     iRowSpan := 1;
                  end;

                  // �� Key Point
                  iBaseRow := RowCount - 1;


                  // �ߺз� �׸� MergeCell üũ

                  { -- ��з��� ������ �ٸ����, �ߺз��� ������, ���� Grid line�� �׸��� ���� �ּ�ó�� @ 2013.10.14 LSH
                  if TpGetSearch.GetOutputDataS('sDutyRmk', i) <> sPrevMidTitle then
                  begin
                  }
                     // ���� �ߺз��� �ٸ� ���
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;


                     // �� Key Point
                     jBaseRow := RowCount - 1;



                     // �Һз� �׸� MergeCell üũ
                     if TpGetSearch.GetOutputDataS('sDeptSpec', i) <> sPrevDownTitle then
                     begin
                        // ���� �Һз��� �ٸ� ���
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // �� Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // ���� �Һз��� �ٸ� ���
                        Inc(kRowSpan);

                  { -- ��з��� ������ �ٸ����, �ߺз��� ������, ���� Grid line�� �׸��� ���� �ּ�ó�� @ 2013.10.14 LSH
                  end
                  else
                  begin
                     // ���� �ߺз��� ���� ���
                     Inc(jRowSpan);
                  end;
                  }
               end
               else
               begin

                  // �ߺз� �׸� MergeCell üũ
                  if TpGetSearch.GetOutputDataS('sDeptNm', i) <> sPrevMidTitle then
                  begin
                     // ���� �ߺз��� �ٸ� ���
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;

                     // �� Key Point
                     jBaseRow := RowCount - 1;


                     // �Һз� �׸� MergeCell üũ
                     if TpGetSearch.GetOutputDataS('sDeptSpec', i) <> sPrevDownTitle then
                     begin
                        // ���� �Һз��� ���� ���
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // �� Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // ���� �Һз��� �ٸ� ���
                        Inc(kRowSpan);
                  end
                  else
                  begin
                     // �ߺз� span Index ++
                     Inc(jRowSpan);


                     // �Һз� �׸� MergeCell üũ
                     if TpGetSearch.GetOutputDataS('sDeptSpec', i) <> sPrevDownTitle then
                     begin
                        // ���� �Һз��� �ٸ� ���
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // �� Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // ���� �Һз��� ������ ���
                        Inc(kRowSpan);

                  end;


                  // ��з� Span index ++
                  Inc(iRowSpan);

               end;


               // Set previous value
               sPrevUptitle   := TpGetSearch.GetOutputDataS('sLocate',   i);
               sPrevmidtitle  := TpGetSearch.GetOutputDataS('sDeptNm',   i);
               sPrevDowntitle := TpGetSearch.GetOutputDataS('sDeptSpec', i);


               // RowCount �����ٷ� �ѱ�.
               RowCount := RowCount + 1;


            end;


            // �� ���� Code List-up�� ������ Row Merge ó��
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


            // RowCount ����
            RowCount := RowCount - 1;
            }


            // Comments
            lb_MyDial.Caption  := IntToStr(iRowCnt) + '���� ����ó ������ ��ȸ�Ͽ����ϴ�.';


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


            // RowCount ����
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
                  Cells[C_AN_SRTITLE,    i+FixedRows]  := '[����]' + TpGetSearch.GetOutputDataS('sDocTitle',   i)
               else if (TpGetSearch.GetOutputDataS('sHeadTail',   i) = 'M') then
                  Cells[C_AN_SRTITLE,    i+FixedRows]  := '[����]' + TpGetSearch.GetOutputDataS('sDocTitle',   i)
               else if (TpGetSearch.GetOutputDataS('sHeadTail',   i) = 'S') then
                  Cells[C_AN_SRTITLE,    i+FixedRows]  := '[����]' + TpGetSearch.GetOutputDataS('sDocTitle',   i)
               else if (TpGetSearch.GetOutputDataS('sHeadTail',   i) = 'C') then
                  Cells[C_AN_SRTITLE,    i+FixedRows]  := '[����]' + TpGetSearch.GetOutputDataS('sDocTitle',   i)
               else if (TpGetSearch.GetOutputDataS('sHeadTail',   i) = 'G') then
                  Cells[C_AN_SRTITLE,    i+FixedRows]  := '[�Ϲ�]' + TpGetSearch.GetOutputDataS('sDocTitle',   i)
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


               if (PosByte('������', Cells[C_AN_PROCESS,    i+FixedRows]) > 0) or
                  (PosByte('��û�Ⱒ', Cells[C_AN_PROCESS,    i+FixedRows]) > 0) then
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
               else if (PosByte('��û�Ϸ�',   Cells[C_AN_PROCESS,    i+FixedRows]) > 0) or
                       (PosByte('��û����',   Cells[C_AN_PROCESS,    i+FixedRows]) > 0) or
                       (PosByte('��û������', Cells[C_AN_PROCESS,    i+FixedRows]) > 0) then
                  FontColors[C_AN_PROCESS,  i+FixedRows] := clRed
               else if (PosByte('��ȹ',   Cells[C_AN_PROCESS,    i+FixedRows]) > 0) then
                  FontColors[C_AN_PROCESS,  i+FixedRows] := clBlue
               else if (PosByte('���',   Cells[C_AN_PROCESS,    i+FixedRows]) > 0) then
                  FontColors[C_AN_PROCESS,  i+FixedRows] := clPurple
               else if (PosByte('�Է���',   Cells[C_AN_PROCESS,    i+FixedRows]) > 0) then
                  FontColors[C_AN_PROCESS,  i+FixedRows] := clGreen;

            end;


            // RowCount ����
            RowCount := RowCount - 1;


            // Comments
            if (iRowCnt = -1) then
               lb_Analysis.Caption  := '�� ' + GetTxMsg + '.'     // �˻��Ǽ� MAXROWCNT �ʰ� �޼��� �˸� @ 2016.11.18 LSH
            else
               lb_Analysis.Caption  := '�� ' + IntToStr(iRowCnt) + '���� S/R ������ ��ȸ�Ͽ����ϴ�.';

         end;
      end
      else if (in_SearchFlag = 'WORKRPT') then
      begin
         with asg_WorkRpt do
         begin
            iRowCnt  := TpGetSearch.RowCount;
            //RowCount := iRowCnt + FixedRows + 1;


            // Merge ����
            iBaseRow       := 0;    // ��з� Merge Start Index
            iRowSpan       := 1;    // ��з� Span  Index
            jBaseRow       := 0;    // �ߺз� Merge Start Index
            jRowSpan       := 1;    // �ߺз� Span  Index
            kBaseRow       := 0;    // �Һз� Merge Start Index
            kRowSpan       := 1;    // �Һз� Span  Index

            sPrevUptitle    := '';   // ��з� ���� Title
            sPrevMidtitle   := '';   // �ߺз� ���� Title
            sPrevDowntitle  := '';   // �Һз� ���� Title



            for i := 0 to iRowCnt - 1 do
            begin

               Cells[C_WR_LOCATE,   RowCount - 1] := TpGetSearch.GetOutputDataS('sLocate'    , i);
               Cells[C_WR_DUTYUSER, RowCount - 1] := StringReplace(TpGetSearch.GetOutputDataS('sDutyPart'  , i), #13#10, ' ', [rfReplaceAll]);
               Cells[C_WR_DUTYPART, RowCount - 1] := TpGetSearch.GetOutputDataS('sDutyUser'  , i);
               Cells[C_WR_CONTEXT,  RowCount - 1] := Trim(TpGetSearch.GetOutputDataS('sContext'   , i));


               // �ۼ��� �Ʒ� ������ �� �ָ�ǥ�� �� Coloring
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



               // �ش� ȭ�� 1-Cycle �ڵ��� ������ Index�� ���� RowCount �� ���
               if (i = iRowCnt) then
                  iLstRowCnt := i+1
               else
                  iLstRowCnt := iRowCnt;




               // ��з� �׸� MergeCell üũ
               if TpGetSearch.GetOutputDataS('sLocate', i) <> sPrevUptitle then
               begin
                  // test
                  //Memo1.Lines.Add('--- Test Logging Start ---');


                  // ���� ��з��� �ٸ� ���
                  if i > 0 then
                  begin
                     MergeCells(0, iBaseRow, 1, iRowSpan);
                     iRowSpan := 1;
                  end;

                  // �� Key Point
                  iBaseRow := RowCount - 1;


                  // �ߺз� �׸� MergeCell üũ

                  { -- ��з��� ������ �ٸ����, �ߺз��� ������, ���� Grid line�� �׸��� ���� �ּ�ó�� @ 2013.10.14 LSH
                  if TpGetSearch.GetOutputDataS('sDutyRmk', i) <> sPrevMidTitle then
                  begin
                  }
                     // ���� �ߺз��� �ٸ� ���
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;


                     // �� Key Point
                     jBaseRow := RowCount - 1;



                     // �Һз� �׸� MergeCell üũ
                     if TpGetSearch.GetOutputDataS('sDutyUser', i) <> sPrevDownTitle then
                     begin
                        // ���� �Һз��� �ٸ� ���
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // �� Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // ���� �Һз��� �ٸ� ���
                        Inc(kRowSpan);

                  { -- ��з��� ������ �ٸ����, �ߺз��� ������, ���� Grid line�� �׸��� ���� �ּ�ó�� @ 2013.10.14 LSH
                  end
                  else
                  begin
                     // ���� �ߺз��� ���� ���
                     Inc(jRowSpan);
                  end;
                  }
               end
               else
               begin

                  // �ߺз� �׸� MergeCell üũ
                  if TpGetSearch.GetOutputDataS('sDutyPart', i) <> sPrevMidTitle then
                  begin
                     // ���� �ߺз��� �ٸ� ���
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;

                     // �� Key Point
                     jBaseRow := RowCount - 1;


                     // �Һз� �׸� MergeCell üũ
                     if TpGetSearch.GetOutputDataS('sDutyUser', i) <> sPrevDownTitle then
                     begin
                        // ���� �Һз��� ���� ���
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // �� Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // ���� �Һз��� �ٸ� ���
                        Inc(kRowSpan);
                  end
                  else
                  begin
                     // �ߺз� span Index ++
                     Inc(jRowSpan);


                     // �Һз� �׸� MergeCell üũ
                     if TpGetSearch.GetOutputDataS('sDutyUser', i) <> sPrevDownTitle then
                     begin
                        // ���� �Һз��� �ٸ� ���
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // �� Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // ���� �Һз��� ������ ���
                        Inc(kRowSpan);

                  end;


                  // ��з� Span index ++
                  Inc(iRowSpan);

               end;


               // Set previous value
               sPrevUptitle   := TpGetSearch.GetOutputDataS('sLocate', i);
               sPrevmidtitle  := TpGetSearch.GetOutputDataS('sDutyPart', i);
               sPrevDowntitle := TpGetSearch.GetOutputDataS('sDutyUser', i);


               // RowCount �����ٷ� �ѱ�.
               RowCount := RowCount + 1;


            end;


            // �� ���� Code List-up�� ������ Row Merge ó��
            asg_WorkRpt.MergeCells(0, iBaseRow, 1, iRowSpan);
            asg_WorkRpt.MergeCells(1, jBaseRow, 1, jRowSpan);
            asg_WorkRpt.MergeCells(2, kBaseRow, 1, kRowSpan);
            asg_WorkRpt.RowCount := iLstRowCnt + 1;


            //------------------------------------------------------------------
            // �α��� User�� Row�� ȭ�� �������� ���� �߰� @ 2016.09.02 LSH
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


               // �ۼ��� �Ʒ� ������ �� �ָ�ǥ�� �� Coloring
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


         // RowCount ����
         asg_WorkRpt.RowCount := asg_WorkRpt.RowCount - 1;
         }


         // Comments
         if (iRowCnt = -1) then
            lb_Analysis.Caption := '�� ' + GetTxMsg + '.'      // �˻��Ǽ� MAXROWCNT �ʰ� �˸� @ 2016.11.18 LSH
         else
            lb_Analysis.Caption := '�� ' + IntToStr(iRowCnt) + '���� ���������� ��ȸ�Ͽ����ϴ�.';

      end
      else if (in_SearchFlag = 'WEEKLYRPT') then
      begin
         with asg_WeeklyRpt do
         begin
            iRowCnt  := TpGetSearch.RowCount;

            //RowCount := iRowCnt + FixedRows + 1;

            // Merge ����
            iBaseRow       := 0;    // ��з� Merge Start Index
            iRowSpan       := 1;    // ��з� Span  Index
            jBaseRow       := 0;    // �ߺз� Merge Start Index
            jRowSpan       := 1;    // �ߺз� Span  Index
            kBaseRow       := 0;    // �Һз� Merge Start Index
            kRowSpan       := 1;    // �Һз� Span  Index

            sPrevUptitle    := '';   // ��з� ���� Title
            sPrevMidtitle   := '';   // �ߺз� ���� Title
            sPrevDowntitle  := '';   // �Һз� ���� Title


            for i := 0 to iRowCnt - 1 do
            begin
               Cells[C_WK_DATE,     RowCount - 1] := TpGetSearch.GetOutputDataS('sRegDate'    , i) +
                                                      '(' + GetDayofWeek(StrToDate(TpGetSearch.GetOutputDataS('sRegDate'   , i)), 'HS') + ')';
               Cells[C_WK_GUBUN,    RowCount - 1] := StringReplace(Trim(TpGetSearch.GetOutputDataS('sDutyRmk'    , i)), #13#10, ' ', [rfReplaceAll]);
               Cells[C_WK_CONTEXT,  RowCount - 1] := Trim(TpGetSearch.GetOutputDataS('sContext'    , i));

               // �ӽú��� �ʱ�ȭ
               tmpWkLine1   := '';
               tmpWkLine2   := '';
               tmpWkLine3   := '';
               tmpWkProc1   := '';
               tmpWkProc2   := '';
               tmpWkProc3   := '';
               tmpWkProcAll := '';

               //---------------------------------------------------------------
               // �������� [����ܰ�] ǥ��
               //    - �ش� Cell�� Text�� CRLF ������ Parsing �Ͽ�
               //      ����/����/�Ϸ��� 3�ܰ�� �����Ȳ�� ǥ��
               //---------------------------------------------------------------
               if (Cells[C_WK_GUBUN,    RowCount - 1] = '��������') then
               begin
                  // CRLF ������ Trim ����
                  tmpWkLine1 := CopyByte(Trim(Cells[C_WK_CONTEXT,  RowCount - 1]), 1, PosByte(#13#10, Trim(Cells[C_WK_CONTEXT,  RowCount - 1])));
                  tmpWkLine2 := CopyByte(Trim(Cells[C_WK_CONTEXT,  RowCount - 1]), PosByte(#13#10, Trim(Cells[C_WK_CONTEXT,  RowCount - 1])) + 1, LengthByte(Trim(Cells[C_WK_CONTEXT,  RowCount - 1])));
                  tmpWkLine3 := CopyByte(Trim(tmpWkLine2), PosByte(#13#10, Trim(tmpWkLine2)), LengthByte(Trim(tmpWkLine2)));

                  // �ٽ� Line by Line���� Trim ����
                  tmpWkLine2 := CopyByte(Trim(tmpWkLine2), 1, PosByte(#13#10, Trim(tmpWkLine2)) - 1);
                  tmpWkLine3 := CopyByte(Trim(tmpWkLine3), 1, LengthByte(Trim(tmpWkLine3)));


                  if PosByte('����', Trim(tmpWkLine1)) > 0 then
                     tmpWkProc1 := '����'
                  else if PosByte('����', Trim(tmpWkLine1)) > 0 then
                     tmpWkProc1 := '����'
                  else if Trim(tmpWkLine1) <> '' then
                     tmpWkProc1 := '�Ϸ�'
                  else
                     tmpWkProc1 := '';

                  if PosByte('����', Trim(tmpWkLine2)) > 0 then
                     tmpWkProc2 := '����'
                  else if PosByte('����', Trim(tmpWkLine2)) > 0 then
                     tmpWkProc2 := '����'
                  else if Trim(tmpWkLine2) <> '' then
                     tmpWkProc2 := '�Ϸ�'
                  else
                     tmpWkProc2 := '';

                  if PosByte('����', Trim(tmpWkLine3)) > 0 then
                     tmpWkProc3 := '����'
                  else if PosByte('����', Trim(tmpWkLine3)) > 0 then
                     tmpWkProc3 := '����'
                  else if Trim(tmpWkLine3) <> '' then
                     tmpWkProc3 := '�Ϸ�'
                  else
                     tmpWkProc3 := '';


                  // ����ܰ� ����
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


               // ����ܰ�  Coloring
               if PosByte('����', Cells[C_WK_STEP,        RowCount - 1]) > 0 then
               begin
                  Colors[C_WK_CONTEXT,    RowCount - 1] := $0074D7AF;
                  Colors[C_WK_STEP,       RowCount - 1] := $0074D7AF;

                  AddComment(C_WK_STEP,
                             RowCount - 1,
                             '���������� [����]��� �ܾ ���Ե� ���,' + #13#10 +
                             '�Ǵ� S/R���� ��û Ư�̻��׿� ��翹���ڰ� ���� ��쿡 ������� ǥ��˴ϴ�.');
               end
               else if PosByte('����', Cells[C_WK_STEP,   RowCount - 1]) > 0 then
               begin
                  Colors[C_WK_CONTEXT,    RowCount - 1] := $0074A9FF;
                  Colors[C_WK_STEP,       RowCount - 1] := $0074A9FF;

                  AddComment(C_WK_STEP,
                             RowCount - 1,
                             '���������� [����]�̶�� �ܾ ���Ե� ���,' + #13#10 +
                             '�Ǵ� S/R���� ����� ������ ������ο�û ���� ������ ��쿡 ���������� ǥ��˴ϴ�.');
               end
               else if (PosByte('�߼�', Cells[C_WK_STEP,   RowCount - 1]) > 0) or
                       (PosByte('����', Cells[C_WK_STEP,   RowCount - 1]) > 0) or
                       (PosByte('ȸ��', Cells[C_WK_STEP,   RowCount - 1]) > 0) or
                       (PosByte('���ó', Cells[C_WK_STEP, RowCount - 1]) > 0) or
                       (PosByte('OCS',  Cells[C_WK_STEP,   RowCount - 1]) > 0) or
                       (PosByte('CRA/CRC',Cells[C_WK_STEP, RowCount - 1]) > 0) then
               begin
                  Colors[C_WK_CONTEXT,    RowCount - 1] := $00C0D5B5;
                  Colors[C_WK_STEP,       RowCount - 1] := $00C0D5B5;

                  AddComment(C_WK_STEP,
                             RowCount - 1,
                             '�������� Tab���� ��ȣ�� ��ȣ�� �����ؼ� �������� ������ġ�� Ȯ���Ͻ� �� �ֽ��ϴ�.');
               end
               else if ((PosByte('�Ϸ�', Cells[C_WK_STEP,   RowCount - 1]) > 0) and
                        (PosByte('����', Cells[C_WK_STEP,   RowCount - 1]) = 0) and
                        (PosByte('����', Cells[C_WK_STEP,   RowCount - 1]) = 0)) or
                        (PosByte('����', Cells[C_WK_STEP,   RowCount - 1]) > 0) or
                        (PosByte('���', Cells[C_WK_STEP,   RowCount - 1]) > 0)
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


               // �ش� ȭ�� 1-Cycle �ڵ��� ������ Index�� ���� RowCount �� ���
               if (i = iRowCnt) then
                  iLstRowCnt := i+1
               else
                  iLstRowCnt := iRowCnt;




               // ��з� �׸� MergeCell üũ
               if TpGetSearch.GetOutputDataS('sRegDate', i) <> sPrevUptitle then
               begin
                  // test
                  //Memo1.Lines.Add('--- Test Logging Start ---');


                  // ���� ��з��� �ٸ� ���
                  if i > 0 then
                  begin
                     MergeCells(0, iBaseRow, 1, iRowSpan);
                     iRowSpan := 1;
                  end;

                  // �� Key Point
                  iBaseRow := RowCount - 1;


                  // �ߺз� �׸� MergeCell üũ

                  { -- ��з��� ������ �ٸ����, �ߺз��� ������, ���� Grid line�� �׸��� ���� �ּ�ó�� @ 2013.10.14 LSH
                  if TpGetSearch.GetOutputDataS('sDutyRmk', i) <> sPrevMidTitle then
                  begin
                  }
                     // ���� �ߺз��� �ٸ� ���
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;


                     // �� Key Point
                     jBaseRow := RowCount - 1;



                     // �Һз� �׸� MergeCell üũ
                     if TpGetSearch.GetOutputDataS('sContext', i) <> sPrevDownTitle then
                     begin
                        // ���� �Һз��� �ٸ� ���
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // �� Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // ���� �Һз��� �ٸ� ���
                        Inc(kRowSpan);

                  { -- ��з��� ������ �ٸ����, �ߺз��� ������, ���� Grid line�� �׸��� ���� �ּ�ó�� @ 2013.10.14 LSH
                  end
                  else
                  begin
                     // ���� �ߺз��� ���� ���
                     Inc(jRowSpan);
                  end;
                  }
               end
               else
               begin

                  // �ߺз� �׸� MergeCell üũ
                  if TpGetSearch.GetOutputDataS('sDutyRmk', i) <> sPrevMidTitle then
                  begin
                     // ���� �ߺз��� �ٸ� ���
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;

                     // �� Key Point
                     jBaseRow := RowCount - 1;


                     // �Һз� �׸� MergeCell üũ
                     if TpGetSearch.GetOutputDataS('sContext', i) <> sPrevDownTitle then
                     begin
                        // ���� �Һз��� ���� ���
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // �� Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // ���� �Һз��� �ٸ� ���
                        Inc(kRowSpan);
                  end
                  else
                  begin
                     // �ߺз� span Index ++
                     Inc(jRowSpan);


                     // �Һз� �׸� MergeCell üũ
                     if TpGetSearch.GetOutputDataS('sContext', i) <> sPrevDownTitle then
                     begin
                        // ���� �Һз��� �ٸ� ���
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // �� Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // ���� �Һз��� ������ ���
                        Inc(kRowSpan);

                  end;


                  // ��з� Span index ++
                  Inc(iRowSpan);

               end;


               // Set previous value
               sPrevUptitle   := TpGetSearch.GetOutputDataS('sRegDate', i);
               sPrevmidtitle  := TpGetSearch.GetOutputDataS('sDutyRmk', i);
               sPrevDowntitle := TpGetSearch.GetOutputDataS('sContext', i);


               // RowCount �����ٷ� �ѱ�.
               RowCount := RowCount + 1;


            end;


            // �� ���� Code List-up�� ������ Row Merge ó��
            asg_WeeklyRpt.MergeCells(0, iBaseRow, 1, iRowSpan);
            asg_WeeklyRpt.MergeCells(1, jBaseRow, 1, jRowSpan);
            asg_WeeklyRpt.RowCount := iLstRowCnt + 1;




            { -- Cell-Merge ������ Logic
               Cells[C_WK_DATE,     i + FixedRows] := TpGetSearch.GetOutputDataS('sRegDate'    , i) +
                                                      '(' + GetDayofWeek(StrToDate(TpGetSearch.GetOutputDataS('sRegDate'   , i)), 'HS') + ')';
               Cells[C_WK_GUBUN,    i + FixedRows] := TpGetSearch.GetOutputDataS('sDutyRmk'    , i);
               Cells[C_WK_CONTEXT,  i + FixedRows] := TpGetSearch.GetOutputDataS('sContext'    , i);
               Cells[C_WK_STEP,     i + FixedRows] := TpGetSearch.GetOutputDataS('sDutySpec'   , i);



               // ǥ�� �� Coloring
               if Cells[C_WK_STEP,        i + FixedRows] = '����' then
               begin
                  Colors[C_WK_STEP,       i + FixedRows] := $00AF9BE7;
               end
               else if Cells[C_WK_STEP,   i + FixedRows] = '����' then
               begin
                  Colors[C_WR_REGDATE,    i + FixedRows] := $00AFCFAF;
               end
               else if Cells[C_WK_STEP,   i + FixedRows] = '�Ϸ�' then
               begin
                  Colors[C_WR_REGDATE,   2 i + FixedRows] := clSilver;
               end
               else
                  Colors[C_WR_REGDATE,    i + FixedRows] := clWhite;

               AutoSizeRow(i+FixedRows);


            end;
            }

         end;


         // RowCount ����
         //asg_WeeklyRpt.RowCount := asg_WeeklyRpt.RowCount - 1;


         // Comments
         lb_Analysis.Caption := '�� [' + Trim(fed_Analysis.Text) + '] �� ������ ' + IntToStr(iRowCnt) + '���� �ۼ��Ͽ����ϴ�.';

      end
      else if (in_SearchFlag = 'DIALMAP') then
      begin
         iRowCnt  := TpGetSearch.RowCount;

            with asg_DialMap do
            begin

               for i := 0 to iRowCnt - 1 do
               begin
                  // ����
                  if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5141') then
                  begin
                     Cells[7, 0]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[7, 0]   := $0093D7AB;
                  end
                  // ���¾�ü PM
                  {
                  else if (TpGetSearch.GetOutputDataS('sDutyPart'  , i) = '������') and
                          (TpGetSearch.GetOutputDataS('sDutySpec'  , i) = '����')   and
                          (TpGetSearch.GetOutputDataS('sCallNo'    , i) <> '5141')  then
                  begin
                     Cells[1, 0]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[1, 0]   := $0093D7AB;
                  end
                  }
                  // ���Ʈ��
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5142') then
                  begin
                     Cells[5, 0]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[5, 0]   := $0093D7AB;
                  end
                  // ��ȹ��Ʈ��
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5861') then
                  begin
                     Cells[3, 0]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[3, 0]   := $0093D7AB;
                  end
                  // ������Ʈ�� @ 2016.07.26 LSH
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5144') then
                  begin
                     Cells[1, 0]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[1, 0]   := $0093D7AB;
                  end
                  {
                  // ����������Ʈ(1)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6214') then
                  begin
                     Cells[0, 1]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[0, 1]   := $00DDD7E9;
                  end
                  }
                  // ����������Ʈ(1) --> ����(1)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6225') then
                  begin
                     Cells[0, 1]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[0, 1]   := $00DDD7E9;
                  end
                  // ����������Ʈ(2) --> ����(2)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6213') then
                  begin
                     Cells[1, 1]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[1, 1]   := $00DDD7E9;
                  end
                  // ����������Ʈ(3) --> ����(3)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6228') then
                  begin
                     Cells[0, 2]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[0, 2]   := $00DDD7E9;
                  end
                  // ����������Ʈ(4) --> ����(4)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6230') then
                  begin
                     Cells[0, 3]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[0, 3]   := $00DDD7E9;
                  end
                  // ����������Ʈ(5) --> ����(5)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6227') then
                  begin
                     Cells[1, 2]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[1, 2]   := $00DDD7E9;
                  end
                  // ����������Ʈ(6) --> ����(6)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5037') then
                  begin
                     Cells[1, 3]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[1, 3]   := $00DDD7E9;
                  end
                  // ����������Ʈ(7)
                  // �ڸ���ġ �������� �߰� @ 2015.06.05 LSH
                  // --> �ӽ��ڸ��� ���� @ 2016.07.26 LSH
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = 'XXXX') then
                  begin
                     Cells[2, 3]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[2, 3]   := $00DDD7E9;
                  end
                  // ������Ʈ(1)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5564') then
                  begin
                     Cells[0, 4]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[0, 4]   := $00FFD7A5;
                  end
                  // ������Ʈ(2)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6219') then
                  begin
                     Cells[1, 4]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[1, 4]   := $00FFD7A5;
                  end
                  // ������Ʈ(3)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6224') then
                  begin
                     Cells[0, 5]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[0, 5]   := $00FFD7A5;
                  end
                  // ������Ʈ(4)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5143') then
                  begin
                     Cells[1, 5]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[1, 5]   := $00FFD7A5;
                  end
                  // ������Ʈ(5)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6218') then
                  begin
                     Cells[2, 4]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[2, 4]   := $00FFD7A5;
                  end
                  // ������Ʈ(6)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6217') then
                  begin
                     Cells[3, 4]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[3, 4]   := $00FFD7A5;
                  end
                  // ������Ʈ(7)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6223') then
                  begin
                     Cells[2, 5]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[2, 5]   := $00FFD7A5;
                  end
                  // ������Ʈ(8)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6216') then
                  begin
                     Cells[3, 5]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[3, 5]   := $00FFD7A5;
                  end
                  // ������Ʈ(1) --> ����(1)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6222') then
                  begin
                     Cells[0, 6]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[0, 6]   := $009DA1FF;
                  end
                  // ������Ʈ(2) --> ����(2)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6221') then
                  begin
                     Cells[1, 6]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[1, 6]   := $009DA1FF;
                  end
                  // ������Ʈ(3) --> ����(3)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6211') then
                  begin
                     Cells[1, 7]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[1, 7]   := $009DA1FF;
                  end
                  // ������Ʈ(4) --> ����(4)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6214') then
                  begin
                     Cells[2, 6]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[2, 6]   := $009DA1FF;
                  end
                  // ������Ʈ(5) --> ����(5)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6212') then
                  begin
                     Cells[2, 7]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[2, 7]   := $009DA1FF;
                  end
                  // ������Ʈ(6) --> ����(6)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6220') then
                  begin
                     Cells[3, 6]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[3, 6]   := $009DA1FF;
                  end
                  // ������Ʈ(7) --> ����(7)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6215') then
                  begin
                     Cells[3, 7]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[3, 7]   := $009DA1FF;
                  end
                  // ���Ʈ(1)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5565') then
                  begin
                     Cells[6, 1]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[6, 1]   := $009DB1C1;
                  end
                  // ���Ʈ(2)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5146') then
                  begin
                     Cells[7, 1]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[7, 1]   := $009DB1C1;
                  end
                  // ���Ʈ(3)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5563') then
                  begin
                     Cells[7, 2]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[7, 2]   := $009DB1C1;
                  end
                  // ���Ʈ(4) - �ý��� ����� �߰� @ 2014.07.17 LSH
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5145') then
                  begin
                     Cells[6, 2]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[6, 2]   := $009DB1C1;
                  end
                  // ��ȹ��Ʈ(1)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5863') then
                  begin
                     Cells[3, 1]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[3, 1]   := $00C5B1C1;
                  end
                  // ��ȹ��Ʈ(2)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '5862') then
                  begin
                     Cells[4, 1]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[4, 1]   := $00C5B1C1;
                  end
                  // �Ϲݰ�����Ʈ(1)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6245') then
                  begin
                     Cells[7, 4]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[7, 4]   := $0085DFF1;
                  end
                  // �Ϲݰ�����Ʈ(2)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6234') then
                  begin
                     Cells[6, 5]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[6, 5]   := $0085DFF1;
                  end
                  // �Ϲݰ�����Ʈ(3)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6233') then
                  begin
                     Cells[7, 5]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[7, 5]   := $0085DFF1;
                  end
                  // �Ϲݰ�����Ʈ(4)
                  else if (TpGetSearch.GetOutputDataS('sCallNo'    , i) = '6232') then
                  begin
                     Cells[6, 4]    := TpGetSearch.GetOutputDataS('sUserNm', i) + #13#10 + TpGetSearch.GetOutputDataS('sCallNo', i);
                     Colors[6, 4]   := $0085DFF1;
                  end
                  // OP ������
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

               // Ư����� Coloring
               if (Cells[C_RL_REMARK,    i+FixedRows] <> '') then
                  FontColors[C_RL_REMARK,  i+FixedRows] := $0089A1FF;

            end;


            // RowCount ����
            RowCount := RowCount - 1;


            // Comments
            if (iRowCnt = -1) then
               lb_RegDoc.Caption  := '�� '  + GetTxMsg + '.'      // �˻��ڷ� MAXROWCNT �ʰ� �޼��� ǥ�� @ 2016.11.18 LSH
            else
               lb_RegDoc.Caption  := '�� ' + IntToStr(iRowCnt) + '���� ������ ������ ��ȸ�Ͽ����ϴ�.';

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


                  // �ٹ�-����(Ȯ��) ��ȸ vs. �ٹ�ǥ ���� �б� @ 2015.06.12 LSH
                  if (in_SearchFlag = 'DUTYMAKE') then
                  begin
                     // �ָ�(���������� ����)
                     if TpGetSearch.GetOutputDataS('sFlag', i) = 'WEEK' then
                     begin
                        Cells[C_DT_DUTYFLAG,  RowCount - 1] := 'W'
                     end
                     // ����(�߰�)
                     else if TpGetSearch.GetOutputDataS('sFlag', i) = '' then
                     begin
                        Cells[C_DT_DUTYFLAG,  RowCount - 1] := 'P'
                     end
                     // ����������(�ָ�����)
                     else
                     begin
                        Cells[C_DT_DUTYFLAG,  RowCount - 1] := TpGetSearch.GetOutputDataS('sFlag', i);
                     end;
                  end
                  else if (in_SearchFlag = 'DUTYSEARCH') then
                  begin
                     if TpGetSearch.GetOutputDataS('sFlag', i) = 'P' then
                        Cells[C_DT_DUTYFLAG,  RowCount - 1] := '����(�߰�)'
                     else if TpGetSearch.GetOutputDataS('sFlag', i) = 'W' then
                        Cells[C_DT_DUTYFLAG,  RowCount - 1] := '����(�ָ�)'
                     else
                        Cells[C_DT_DUTYFLAG,  RowCount - 1{i+FixedRows}] := '����(' + TpGetSearch.GetOutputDataS('sFlag', i) + ')'
                  end;


                  Cells[C_DT_DUTYUSER,  RowCount - 1{i+FixedRows}] := TpGetSearch.GetOutputDataS('sDutyUser'  , i);
                  Cells[C_DT_REMARK,    RowCount - 1{i+FixedRows}] := TpGetSearch.GetOutputDataS('sDutyRmk'   , i);
                  Cells[C_DT_SEQNO ,    RowCount - 1{i+FixedRows}] := TpGetSearch.GetOutputDataS('sDocSeq'    , i);     // ������ �������� Seqno @ 2015.06.12 LSH
                  Cells[C_DT_ORGDTYUSR, RowCount - 1{i+FixedRows}] := TpGetSearch.GetOutputDataS('sDutyUser'  , i);     // ������ ������ Orig. �ٹ��� @ 2015.06.12 LSH
                  Cells[C_DT_ORGREMARK, RowCount - 1{i+FixedRows}] := TpGetSearch.GetOutputDataS('sDutyRmk'   , i);     // ���� Ư�̻��� ������ Orig. Ư�̻��� @ 2015.06.12 LSH


                  //AddButton(C_M_USERADD, i+FixedRows, ColWidths[C_M_USERADD]-5, 20, '+', haBeforeText, vaCenter);                // ����� Add


                  //------------------------------------------------------------
                  // ���ϱٹ� Coloring
                  //------------------------------------------------------------
                  if (TpGetSearch.GetOutputDataS('sFlag', i) <> 'P') then
                  begin
                     FontColors[C_DT_YOIL,      RowCount - 1{i+FixedRows}] := clRed;
                     FontColors[C_DT_DUTYFLAG,  RowCount - 1{i+FixedRows}] := clRed;
                     FontColors[C_DT_DUTYDATE,  RowCount - 1{i+FixedRows}] := clRed;
                  end;


                  // �ٹ�-����(Ȯ��) ��ȸ��, Coloring @ 2015.06.12 LSH
                  if (in_SearchFlag = 'DUTYSEARCH') then
                  begin

                     //------------------------------------------------------------
                     // �������� Coloring
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
                     // ���αٹ� Coloring
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


         // RowCount ����
         asg_Duty.RowCount := asg_Duty.RowCount - 1;


         // Comments
         if (in_SearchFlag = 'DUTYSEARCH') then
            lb_RegDoc.Caption := '�� ���ñⰣ�� �ٹ� �� ������Ȳ�� ��ȸ�Ͽ����ϴ�.'
         else if (in_SearchFlag = 'DUTYMAKE') then
            lb_RegDoc.Caption := '�� ���ñⰣ�� (Ȯ�������) �ٹ�ǥ�� ��ȸ�Ͽ����ϴ�.';


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
                  Cells[2,  i+FixedRows] := TpGetSearch.GetOutputDataS('sRegDate'   , i);    // (�ٹ����� ��������) ���ϱٹ� ������ @ 2015.06.12 LSH
                  Cells[3,  i+FixedRows] := TpGetSearch.GetOutputDataS('sTestDate'  , i);    // (�ٹ����� ��������) ���ϱٹ� ������ @ 2015.06.12 LSH
               end;
            end;
         end;


         // RowCount ����
         asg_DutyList.RowCount := asg_DutyList.RowCount - 1;

         // Comments
         lb_RegDoc.Caption := '�� �μ� �����ٹ�ǥ ���� �� ������ ��ȸ�Ͽ����ϴ�.';
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


         // RowCount ����
         asg_DBMaster.RowCount := asg_DBMaster.RowCount - 1;

         // Comments
         lb_Analysis.Caption := '�� ' + fcb_WorkArea.Text + ' ' + fcb_DutyPart.Text + '��Ʈ ' + fcb_WorkGubun.Text + '�� ��ȸ�Ͽ����ϴ�.';
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


         // RowCount ����
         asg_DBDetail.RowCount := asg_DBDetail.RowCount - 1;

         // Comments
         lb_Analysis.Caption := '�� ' + fcb_WorkArea.Text + ' ' + fcb_DutyPart.Text + '��Ʈ ' + fcb_WorkGubun.Text + ' [' + asg_DBMaster.Cells[0, asg_DBMaster.Row] + '] ��ȸ�Ͽ����ϴ�.';
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
                  // ���� Ű���� ��Ī ���� @ 2015.06.03 LSH
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


         // RowCount ����
         asg_DBDetail.RowCount := asg_DBDetail.RowCount - 1;

         // Comments
         lb_Analysis.Caption := '�� ' + fcb_WorkArea.Text + ' ' + fcb_DutyPart.Text + ' Ű���� �˻� [' + Trim(fed_Analysis.Text) + '] ' + IntToStr(iRowCnt) + '���� ��ȸ�Ͽ����ϴ�.';
      end
      else if (in_SearchFlag = 'SRHISTORY') then
      begin
         with asg_WeeklyRpt do
         begin
            iRowCnt  := TpGetSearch.RowCount;

            //RowCount := iRowCnt + FixedRows + 1;

            // Merge ����
            iBaseRow       := 0;    // ��з� Merge Start Index
            iRowSpan       := 1;    // ��з� Span  Index
            jBaseRow       := 0;    // �ߺз� Merge Start Index
            jRowSpan       := 1;    // �ߺз� Span  Index
            kBaseRow       := 0;    // �Һз� Merge Start Index
            kRowSpan       := 1;    // �Һз� Span  Index

            sPrevUptitle    := '';   // ��з� ���� Title
            sPrevMidtitle   := '';   // �ߺз� ���� Title
            sPrevDowntitle  := '';   // �Һз� ���� Title


            for i := 0 to iRowCnt - 1 do
            begin
               Cells[C_WK_DATE,     RowCount - 1] := TpGetSearch.GetOutputDataS('sRegDate'    , i);
               Cells[C_WK_GUBUN,    RowCount - 1] := TpGetSearch.GetOutputDataS('sDutyRmk'    , i);
               Cells[C_WK_CONTEXT,  RowCount - 1] := Trim(TpGetSearch.GetOutputDataS('sContext'    , i));

               Cells[C_WK_STEP,     RowCount - 1] := TpGetSearch.GetOutputDataS('sDutySpec'   , i);


               // ����ܰ�  Coloring
               if PosByte('����', Cells[C_WK_STEP,        RowCount - 1]) > 0 then
               begin
                  Colors[C_WK_CONTEXT,    RowCount - 1] := $0074D7AF;
                  Colors[C_WK_STEP,       RowCount - 1] := $0074D7AF;
                  AddComment(C_WK_STEP,
                             RowCount - 1,
                             '���������� [����]��� �ܾ ���Ե� ���,' + #13#10 +
                             '�Ǵ� S/R���� ��û Ư�̻��׿� ��翹���ڰ� ���� ��쿡 ������� ǥ��˴ϴ�.');
               end
               else if PosByte('����', Cells[C_WK_STEP,   RowCount - 1]) > 0 then
               begin
                  Colors[C_WK_CONTEXT,    RowCount - 1] := $0074A9FF;
                  Colors[C_WK_STEP,       RowCount - 1] := $0074A9FF;
                  AddComment(C_WK_STEP,
                             RowCount - 1,
                             '���������� [����]�̶�� �ܾ ���Ե� ���,' + #13#10 +
                             '�Ǵ� S/R���� ����� ������ ������ο�û ���� ������ ��쿡 ���������� ǥ��˴ϴ�.');
               end
               else if (PosByte('�Ϸ�', Cells[C_WK_STEP,   RowCount - 1]) > 0) and
                       (PosByte('����', Cells[C_WK_STEP,   RowCount - 1]) = 0) and
                       (PosByte('����', Cells[C_WK_STEP,   RowCount - 1]) = 0) then
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


               // �ش� ȭ�� 1-Cycle �ڵ��� ������ Index�� ���� RowCount �� ���
               if (i = iRowCnt) then
                  iLstRowCnt := i+1
               else
                  iLstRowCnt := iRowCnt;




               // ��з� �׸� MergeCell üũ
               if TpGetSearch.GetOutputDataS('sRegDate', i) <> sPrevUptitle then
               begin
                  // test
                  //Memo1.Lines.Add('--- Test Logging Start ---');


                  // ���� ��з��� �ٸ� ���
                  if i > 0 then
                  begin
                     MergeCells(0, iBaseRow, 1, iRowSpan);
                     iRowSpan := 1;
                  end;

                  // �� Key Point
                  iBaseRow := RowCount - 1;


                  // �ߺз� �׸� MergeCell üũ

                  { -- ��з��� ������ �ٸ����, �ߺз��� ������, ���� Grid line�� �׸��� ���� �ּ�ó�� @ 2013.10.14 LSH
                  if TpGetSearch.GetOutputDataS('sDutyRmk', i) <> sPrevMidTitle then
                  begin
                  }
                     // ���� �ߺз��� �ٸ� ���
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;


                     // �� Key Point
                     jBaseRow := RowCount - 1;



                     // �Һз� �׸� MergeCell üũ
                     if TpGetSearch.GetOutputDataS('sContext', i) <> sPrevDownTitle then
                     begin
                        // ���� �Һз��� �ٸ� ���
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // �� Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // ���� �Һз��� �ٸ� ���
                        Inc(kRowSpan);

                  { -- ��з��� ������ �ٸ����, �ߺз��� ������, ���� Grid line�� �׸��� ���� �ּ�ó�� @ 2013.10.14 LSH
                  end
                  else
                  begin
                     // ���� �ߺз��� ���� ���
                     Inc(jRowSpan);
                  end;
                  }
               end
               else
               begin

                  // �ߺз� �׸� MergeCell üũ
                  if TpGetSearch.GetOutputDataS('sDutyRmk', i) <> sPrevMidTitle then
                  begin
                     // ���� �ߺз��� �ٸ� ���
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;

                     // �� Key Point
                     jBaseRow := RowCount - 1;


                     // �Һз� �׸� MergeCell üũ
                     if TpGetSearch.GetOutputDataS('sContext', i) <> sPrevDownTitle then
                     begin
                        // ���� �Һз��� ���� ���
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // �� Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // ���� �Һз��� �ٸ� ���
                        Inc(kRowSpan);
                  end
                  else
                  begin
                     // �ߺз� span Index ++
                     Inc(jRowSpan);


                     // �Һз� �׸� MergeCell üũ
                     if TpGetSearch.GetOutputDataS('sContext', i) <> sPrevDownTitle then
                     begin
                        // ���� �Һз��� �ٸ� ���
                        if i > 0 then
                        begin
                           MergeCells(2, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // �� Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // ���� �Һз��� ������ ���
                        Inc(kRowSpan);

                  end;


                  // ��з� Span index ++
                  Inc(iRowSpan);

               end;


               // Set previous value
               sPrevUptitle   := TpGetSearch.GetOutputDataS('sRegDate', i);
               sPrevmidtitle  := TpGetSearch.GetOutputDataS('sDutyRmk', i);
               sPrevDowntitle := TpGetSearch.GetOutputDataS('sContext', i);


               // RowCount �����ٷ� �ѱ�.
               RowCount := RowCount + 1;


            end;


            // �� ���� Code List-up�� ������ Row Merge ó��
            asg_WeeklyRpt.MergeCells(0, iBaseRow, 1, iRowSpan);
            asg_WeeklyRpt.MergeCells(1, jBaseRow, 1, jRowSpan);
            asg_WeeklyRpt.RowCount := iLstRowCnt + 1;
         end;

         // Comments
         //lb_.Caption := '�� ' + Trim(fed_Analysis.Text) + '�� ������ ' + IntToStr(iRowCnt) + '���� �ۼ��Ͽ����ϴ�.';

      end

      else if (in_SearchFlag = 'WORKCONN') then
      begin
         with asg_WorkConn do
         begin
            iRowCnt  := TpGetSearch.RowCount;
            //RowCount := iRowCnt + FixedRows + 1;


            // Merge ����
            iBaseRow       := 0;    // ��з� Merge Start Index
            iRowSpan       := 1;    // ��з� Span  Index
            jBaseRow       := 0;    // �ߺз� Merge Start Index
            jRowSpan       := 1;    // �ߺз� Span  Index
            kBaseRow       := 0;    // �Һз� Merge Start Index
            kRowSpan       := 1;    // �Һз� Span  Index

            sPrevUptitle    := '';   // ��з� ���� Title
            sPrevMidtitle   := '';   // �ߺз� ���� Title
            sPrevDowntitle  := '';   // �Һз� ���� Title



            for i := 0 to iRowCnt - 1 do
            begin

               Cells[C_WC_LOCATE,   RowCount - 1] := TpGetSearch.GetOutputDataS('sLocate'    , i);

               if PosByte('����', TpGetSearch.GetOutputDataS('sDutyPart'  , i)) > 0 then
                  Cells[C_WC_DUTYPART, RowCount - 1] := TpGetSearch.GetOutputDataS('sDutyPart'  , i) + #13#10 + '(' + TpGetSearch.GetOutputDataS('sDutySpec'  , i) + ')'
               else
                  Cells[C_WC_DUTYPART, RowCount - 1] := TpGetSearch.GetOutputDataS('sDutyPart'  , i);

               Cells[C_WC_CONTEXT,  RowCount - 1] := Trim(TpGetSearch.GetOutputDataS('sContext' , i));
               Cells[C_WC_DUTYUSER, RowCount - 1] := TpGetSearch.GetOutputDataS('sDutyUser'  , i);
               Cells[C_WC_REGDATE,  RowCount - 1] := TpGetSearch.GetOutputDataS('sRegDate'   , i);


               AutoSizeRow(RowCount - 1);



               // �ش� ȭ�� 1-Cycle �ڵ��� ������ Index�� ���� RowCount �� ���
               if (i = iRowCnt) then
                  iLstRowCnt := i+1
               else
                  iLstRowCnt := iRowCnt;




               // ��з� �׸� MergeCell üũ
               if TpGetSearch.GetOutputDataS('sLocate', i) <> sPrevUptitle then
               begin
                  // test
                  //Memo1.Lines.Add('--- Test Logging Start ---');


                  // ���� ��з��� �ٸ� ���
                  if i > 0 then
                  begin
                     MergeCells(0, iBaseRow, 1, iRowSpan);
                     iRowSpan := 1;
                  end;

                  // �� Key Point
                  iBaseRow := RowCount - 1;


                  // �ߺз� �׸� MergeCell üũ

                  { -- ��з��� ������ �ٸ����, �ߺз��� ������, ���� Grid line�� �׸��� ���� �ּ�ó�� @ 2013.10.14 LSH
                  if TpGetSearch.GetOutputDataS('sDutyRmk', i) <> sPrevMidTitle then
                  begin
                  }
                     // ���� �ߺз��� �ٸ� ���
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;


                     // �� Key Point
                     jBaseRow := RowCount - 1;



                     // �Һз� �׸� MergeCell üũ
                     if TpGetSearch.GetOutputDataS('sDutyUser', i) <> sPrevDownTitle then
                     begin
                        // ���� �Һз��� �ٸ� ���
                        if i > 0 then
                        begin
                           MergeCells(4, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // �� Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // ���� �Һз��� �ٸ� ���
                        Inc(kRowSpan);

                  { -- ��з��� ������ �ٸ����, �ߺз��� ������, ���� Grid line�� �׸��� ���� �ּ�ó�� @ 2013.10.14 LSH
                  end
                  else
                  begin
                     // ���� �ߺз��� ���� ���
                     Inc(jRowSpan);
                  end;
                  }
               end
               else
               begin

                  // �ߺз� �׸� MergeCell üũ
                  if TpGetSearch.GetOutputDataS('sDutySpec', i) <> sPrevMidTitle then
                  begin
                     // ���� �ߺз��� �ٸ� ���
                     if i > 0 then
                     begin
                        MergeCells(1, jBaseRow, 1, jRowSpan);
                        jRowSpan := 1;
                     end;

                     // �� Key Point
                     jBaseRow := RowCount - 1;


                     // �Һз� �׸� MergeCell üũ
                     if TpGetSearch.GetOutputDataS('sDutyUser', i) <> sPrevDownTitle then
                     begin
                        // ���� �Һз��� ���� ���
                        if i > 0 then
                        begin
                           MergeCells(4, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // �� Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // ���� �Һз��� �ٸ� ���
                        Inc(kRowSpan);
                  end
                  else
                  begin
                     // �ߺз� span Index ++
                     Inc(jRowSpan);


                     // �Һз� �׸� MergeCell üũ
                     if TpGetSearch.GetOutputDataS('sDutyUser', i) <> sPrevDownTitle then
                     begin
                        // ���� �Һз��� �ٸ� ���
                        if i > 0 then
                        begin
                           MergeCells(4, kBaseRow, 1, kRowSpan);
                           kRowSpan := 1;
                        end;


                        // �� Key Point
                        kBaseRow := RowCount - 1;
                     end
                     else
                        // ���� �Һз��� ������ ���
                        Inc(kRowSpan);

                  end;


                  // ��з� Span index ++
                  Inc(iRowSpan);

               end;


               // Set previous value
               sPrevUptitle   := TpGetSearch.GetOutputDataS('sLocate', i);
               sPrevmidtitle  := TpGetSearch.GetOutputDataS('sDutySpec', i);
               sPrevDowntitle := TpGetSearch.GetOutputDataS('sDutyUser', i);


               // RowCount �����ٷ� �ѱ�.
               RowCount := RowCount + 1;


            end;


            // �� ���� Code List-up�� ������ Row Merge ó��
            asg_WorkConn.MergeCells(0, iBaseRow, 1, iRowSpan);
            asg_WorkConn.MergeCells(1, jBaseRow, 1, jRowSpan);
            asg_WorkConn.MergeCells(5, kBaseRow, 1, kRowSpan);
            asg_WorkConn.RowCount := iLstRowCnt + 1;

         end;

         // Comments
         lb_WorkConn.Caption := '�� ' + IntToStr(iRowCnt) + '���� �ֿ�������� ������ ��ȸ�Ͽ����ϴ�.';

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

                  // No1. Range ������ Coloring
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


                  // No2. Range ������ Coloring
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


                  // No3. Range ������ Coloring
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


                  // No4. Range ������ Coloring
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


                  // No5. Range ������ Coloring
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


                  // No6. Range ������ Coloring
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

                  // No (���ʽ�). Range ������ Coloring
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


            // RowCount ����
            RowCount := RowCount - 1;


            // Comments
            //lb_RegDoc.Caption  := '�� ' + IntToStr(iRowCnt) + '���� ������ ������ ��ȸ�Ͽ����ϴ�.';

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
      // �������� Toggle ���
      if (ACol = C_M_DELDATE) then
         if (Cells[ACol, ARow] <> '') then
            Cells[ACol, ARow] := ''
         else
            Cells[ACol, ARow] := FormatDateTime('yyyy-mm-dd', Date);

      // ���� ����� IP ���� ��ȸ
      if (ACol = C_M_USERIP) then
         Cells[ACol, ARow] := FsUserIp;

   end;
end;



procedure TMainDlg.asg_Master_CanEditCell(Sender: TObject; ARow,
  ACol: Integer; var CanEdit: Boolean);
begin
   // ��������, Update, ����� �߰��� Edit ����
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
         lb_Network.Caption := '�� ����� IP�� ����Ͻø�, ��Ȱ�� ��������(��: �ƹθ޸�,ê�ڽ� ��)�� ������ ������ �� �ֽ��ϴ�.'
      else if (ACol = C_M_USERID) then
         lb_Network.Caption := '�� ����� ������ ���� �Է� ���ϼŵ� �������ϴ�.'
      else if (ACol = C_M_DELDATE) then
         lb_Network.Caption := '�� ������ ����������� �����ϱ� ���Ͻø�, �������ڸ� [����Ŭ��]�Ͽ� ������Ʈ ���ּ���.'
      else
         lb_Network.Caption := '';
   end;
end;



//------------------------------------------------------------------------------
// [���] �ÿ��� ���� List ���
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
                 PChar('[��󿬶���] - [���������]�� ����� IP ����� ����Ͻ� �� �ֽ��ϴ�.'),
                 PChar(Self.Caption + ' : ��ϵ��� ���� User ���� �˸�'),
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


   // �α� Update
   UpdateLog('MASTER',
             'PRINT',
             FsUserIp,
             'P',
             FsUserSpec,
             FsUserNm,
             varResult
             );


   // Comments
   lb_Network.Caption := '�� ��������� ����� �Ϸ�Ǿ����ϴ�.';
end;



//------------------------------------------------------------------------------
// [����Ʈ�ɼ�] Set Print Options Handler
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


   // Report Name ����
   RecName := '���������� ����� ����';


   // �Ƿ���� Display
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
      ts := ts+ ((tw-textwidth('������б��Ƿ��')) shr 1);

      Textout(ts, -15, '������б��Ƿ��');

      font.assign(savefont);

      savefont.free;
   end;


   // ������� Display
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

   // ���κ� �������� Update
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
                       PChar('[�ٹ�ó/��������/�����/����ó]�� �ʼ��Է� �׸��Դϴ�.'),
                       PChar(Self.Caption + ' : ���������� �ʼ��Է� �׸� �˸�'),
                       MB_OK + MB_ICONWARNING);

            Exit;
         end;


         if ((PosByte('�', Cells[C_D_DUTYPART, ARow]) > 0) or
             (PosByte('��ȹ', Cells[C_D_DUTYPART, ARow]) > 0)) and
            (Cells[C_D_DUTYSPEC, ARow] = '') then
         begin
            MessageBox(self.Handle,
                       PChar('[������]�� �Է��� �ֽñ� �ٶ��ϴ�.'),
                       PChar(Self.Caption + ' : ���������� �ʼ��Է� �׸� �˸�'),
                       MB_OK + MB_ICONWARNING);

            Exit;
         end;

         if (PosByte('����', Cells[C_D_DUTYPART, ARow]) > 0) and
            ((Cells[C_D_DUTYSPEC, ARow] = '') or
             (Cells[C_D_DUTYRMK,  ARow] = '') or
             (Cells[C_D_DUTYPTNR, ARow] = '')) then
         begin
            MessageBox(self.Handle,
                       PChar('������Ʈ�� ��� [������/Remark/��Ʈ�ʽ�]�� �ʼ��Է� �׸��Դϴ�.'),
                       PChar(Self.Caption + ' : ���������� �ʼ��Է� �׸� �˸�'),
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
                             asg_Detail.Cells[C_D_DUTYUSER, asg_Detail.Row] +' ���� ������ ���������� [������Ʈ]�Ǿ����ϴ�.'),
                       '[KUMC ���̾�α�] ���������� ������Ʈ �˸� ',
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
      asg_Detail.AddButton(C_D_USERADD, ARow + 1, asg_Detail.ColWidths[C_D_USERADD]-5, 20, '+', haBeforeText, vaCenter);                // ����� Add
   end;

end;




procedure TMainDlg.asg_Detail_CanEditCell(Sender: TObject; ARow,
  ACol: Integer; var CanEdit: Boolean);
begin
   // ��������, Update, ����� �߰��� Edit ����
   CanEdit := (ACol <> C_D_DELDATE) and
              (ACol <> C_D_BUTTON)  and
              (ACol <> C_D_USERADD);
end;



procedure TMainDlg.asg_Detail_DblClickCell(Sender: TObject; ARow,
  ACol: Integer);
begin
   with asg_Detail do
   begin
      // �������� Toggle ���
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

         AddComboString('�Ƿ��');
         AddComboString('�Ⱦ�');
         AddComboString('����');
         AddComboString('�Ȼ�');
      end;

      if (ARow > 0) and (PosByte('�Ƿ��', Cells[C_D_LOCATE, ARow]) > 0) and
                        (ACol = C_D_DUTYPART) then
      begin
         AEditor := edComboList;

         ClearComboString;

         AddComboString('������');
         AddComboString('Ȩ������');
      end;


      if (ARow > 0) and (PosByte('�Ⱦ�', Cells[C_D_LOCATE, ARow]) > 0) and
                        (ACol = C_D_DUTYPART) then
      begin
         AEditor := edComboList;

         ClearComboString;

         AddComboString('���Ʈ');
         AddComboString('������Ʈ');
         AddComboString('��ȹ��Ʈ');
         AddComboString('OP');
         AddComboString('PC/AS');
         AddComboString('PACS');
      end;


      if (ARow > 0) and ((PosByte('����', Cells[C_D_LOCATE, ARow]) > 0) or
                         (PosByte('�Ȼ�', Cells[C_D_LOCATE, ARow]) > 0)) and
                        (ACol = C_D_DUTYPART) then
      begin
         AEditor := edComboList;

         ClearComboString;

         AddComboString('���Ʈ');
         AddComboString('OP');
         AddComboString('PC/AS');
         AddComboString('PACS');
      end;

      if (ARow > 0) and (PosByte('�', Cells[C_D_DUTYPART, ARow]) > 0)
                    and (ACol = C_D_DUTYSPEC) then
      begin
         AEditor := edComboList;

         ClearComboString;

         AddComboString('��Ʈ��');
         AddComboString('PC����');
         AddComboString('��Ʈ��ũ');
         AddComboString('�ý���');
         AddComboString('D/B');
      end;

      if (ARow > 0) and (PosByte('����', Cells[C_D_DUTYPART, ARow]) > 0)
                    and (ACol = C_D_DUTYSPEC) then
      begin
         AEditor := edComboList;

         ClearComboString;

         AddComboString('��Ʈ��');  // ������Ʈ�� �߰� @ 2016.07.26 LSH
         AddComboString('����');
         AddComboString('����');
         AddComboString('��������');
         AddComboString('�Ϲݰ���');
         AddComboString('EMR');
         AddComboString('��Ÿ');
      end;

      if (ARow > 0) and (PosByte('����',  Cells[C_D_DUTYPART, ARow]) > 0)
                    and (ACol = C_D_DUTYPTNR) then
      begin
         AEditor := edComboList;

         ClearComboString;

         AddComboString('P/L');
         AddComboString('�Ƿ��');
         AddComboString('��Ʈ�ʽ�');
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
         lb_Network.Caption := '�� ����ϰ� �ִ� �����󼼳���(��: ������Ʈ - û��/�ܷ�,������ȣ/�౹ ��)�� �����ּ���.'
      else if (ACol = C_D_DUTYPTNR) then
         lb_Network.Caption := '�� ���������� [������Ʈ] �ʼ��Է��׸�����, P/L-�Ƿ��-��Ʈ�ʽ�(���¾�ü)�� ǥ���մϴ�.'
      else if (ACol = C_D_DELDATE) then
         lb_Network.Caption := '�� ������ ������������ �����ϱ� ���Ͻø�, �������ڸ� [����Ŭ��]�Ͽ� ������Ʈ ���ּ���.'
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
                 PChar('[��󿬶���] - [���������]�� ����� IP ����� ����Ͻ� �� �ֽ��ϴ�.'),
                 PChar(Self.Caption + ' : ��ϵ��� ���� User ���� �˸�'),
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


   // �α� Update
   UpdateLog('DETAIL',
             'PRINT',
             FsUserIp,
             'P',
             FsUserSpec,
             FsUserNm,
             varResult
             );


   // Comments
   lb_Network.Caption := '�� ���������� ����� �Ϸ�Ǿ����ϴ�.';

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


   // Report Name ����
   RecName := '���������� ���� ������';


   // �Ƿ���� Display
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
      ts := ts+ ((tw-textwidth('������б��Ƿ��')) shr 1);

      Textout(ts, -15, '������б��Ƿ��');

      font.assign(savefont);

      savefont.free;
   end;


   // ������� Display
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
   //  Data Loading bar ���̱�
   //------------------------------------------------------------------
   //SetLoadingBar('ON');


   SelGridInfo('MASTER');


   //------------------------------------------------------------------
   //  Data Loading bar ���߱�
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
// [���] Button onClick Event Handler
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
                 PChar('[��󿬶���] - [���������]�� ����� IP ����� ����Ͻ� �� �ֽ��ϴ�.'),
                 PChar(Self.Caption + ' : ��ϵ��� ���� User ���� �˸�'),
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



      // �α� Update
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
      lb_Network.Caption := '�� ��󿬶��� ����� �Ϸ�Ǿ����ϴ�.';

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


      // �α� Update
      UpdateLog('DETAIL',
                'PRINT',
                FsUserIp,
                'P',
                FsUserSpec,
                FsUserNm,
                varResult
                );


      // Comments
      lb_Network.Caption := '�� ���������� ����� �Ϸ�Ǿ����ϴ�.';

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


      // �α� Update
      UpdateLog('MASTER',
                'PRINT',
                FsUserIp,
                'P',
                FsUserSpec,
                FsUserNm,
                varResult
                );


      // Comments
      lb_Network.Caption := '�� ��������� ����� �Ϸ�Ǿ����ϴ�.';

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


   // Report Name ����
   RecName := '��������� ��󿬶���';
   addInfo := ' Fax no. [�Ⱦ�] 02-920-5672 [����] 02-2626-2239 [�Ȼ�] 031-412-5999';


   // �Ƿ���� Display
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
      ts := ts+ ((tw-textwidth('������б��Ƿ��')) shr 1);

      Textout(ts, -15, '������б��Ƿ��');

      font.assign(savefont);

      savefont.free;
   end;


   // ������� Display
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

   // �������� ver. Display
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


   // ������ Fax. Info Display
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

   // �������� BPL �׽�Ʈ
   //try
      {
      FForm := BplFormCreate('INFORM');

      FForm.Position := poMainFormCenter;

      FForm.Show;
      }

   //-----------------------------------------------------
   // ��ϵ� User ���̾� Chat ����Ʈ �˾� @ 2015.03.30 LSH
   //-----------------------------------------------------
   {  -- Chat XE7 ���������� �ּ� @ 2017.10.31 LSH
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
// [��ȸ] �� Grid �� �޼��� ����
//    - ���� �����ڵ�(MDMCOMCT.COMCD1 = 'KDIAL',
//                             COMCD2 : �� ȭ�� ������
//                             COMCD3 : �޼��� ���� ������
//
// Author : Lee, Se-Ha
// Date   : 2013.08.12
//------------------------------------------------------------------------------
function TMainDlg.GetMsgInfo(in_ComCd2, in_ComCd3 : String) : String;
var
//   i : integer;
   TpGetColumn : TTpSvc;
begin
   // ���ϰ� �ʱ�ȭ
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
         // Ư�� �Ⱓ ������ Inform�� ���, �����Ⱓ ��ȿ�� Check
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
            // Ư�� �Ⱓ ������ Inform �� �ƴѰ��, Version üũ
            //-------------------------------------------------------------
            if (in_ComCd3 = 'VERSION') then
            begin
               // ���α׷� �ֽ� Version ����
               FsVersion := 'v' + TpGetColumn.GetOutputDataS('sComcdnm3', 0);

               //-------------------------------------------------------------
               // Version üũ��, �� ver. �̸� ���� �߰�
               //-------------------------------------------------------------
               if (CopyByte(self.Caption, PosByte('[', self.Caption) + 1, 6) < 'ver' + TpGetColumn.GetOutputDataS('sComcdnm3', 0)) then
                  Result := '�� [' + TpGetColumn.GetOutputDataS('sComcdnm3', 0) + '] �ٿ�ε� �ȳ�: [Ŀ�´�Ƽ] -> �ֽ� Version�� ����غ�����.'
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
// [��ȸ] �� �ٹ�ó�� ������ ����
//       - ���� ������ �����Ĵ� Excel ���ε�� ������, ������ Table ��ȸ�� ��ȯ.
//
// Author : Lee, Se-Ha
// Date   : 2013.09.16
//------------------------------------------------------------------------------
function TMainDlg.GetDietInfo(in_Gubun, in_DayGubun : String) : String;
var
   TpGetMenu : TTpSvc;
   sType1, sType2, sType3, sType4, sType5, sLocateNm, sDietTime : String;
begin
   // ���ϰ� �ʱ�ȭ
   GetDietInfo := '';



   Screen.Cursor := crHourglass;


   //----------------------------------------------------
   // ������ ���Ϻ� ������ ��ȸ Var.
   //----------------------------------------------------
   sType1 := '15';
   sType2 := 'DIET';

   // ������ IP �ĺ���, �ٹ�ó Assign
   if PosByte('�Ⱦϵ�����', FsUserIp) > 0 then
   begin
      sType3 := '�ȾϽĴ����̺���ƮDB';
      sType4 := '�ȾϽĴ����̺���ƮDB';
      sLocateNm := '�Ⱦ�';
   end
   else if PosByte('���ε�����', FsUserIp) > 0 then
   begin
      sType3 := '���νĴ����̺���ƮDB';
      sType4 := '���νĴ����̺���ƮDB';
      sLocateNm := '����';
   end
   else if PosByte('�Ȼ굵����', FsUserIp) > 0 then
   begin
      sType3 := '�Ȼ�Ĵ����̺���ƮDB';
      sType4 := '�Ȼ�Ĵ����̺���ƮDB';
      sLocateNm := '�Ȼ�';
   end
   // ���� IP�� �ƴѰ��(��: ��Ʈ��, �ǻ��� ������ ��), ����
   else
   begin
      Exit;
   end;

   // ���� Flag
   if (FormatDateTime('hh:nn', Now) > '00:00') and (FormatDateTime('hh:nn', Now) <= '08:30') then
   begin
      sType5      := '1';
      sDietTime   := '��ħ';
   end
   else if (FormatDateTime('hh:nn', Now) > '08:30') and (FormatDateTime('hh:nn', Now) <= '14:00') then
   begin
      sType5      := '2';
      sDietTime   := '����';
   end
   else if (FormatDateTime('hh:nn', Now) > '14:00') and (FormatDateTime('hh:nn', Now) <= '19:00') then
   begin
      sType5      := '3';
      sDietTime   := '����';
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
                           'S_CODE6'  , 'sDutyRmk'       // ���Ϻ� �������� (����)
                         { -- ������ ���Ϻ� ������ Table �ڵ���ȸ ���� �ּ� @ 2014.08.26 LSH
                         , 'S_STRING5', 'sAttachNm'      // ����
                         , 'S_STRING6', 'sHideFile'      // ����
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
         // ���� �ð��뺰 ���� ���� ���� Return
         //-------------------------------------------------------------
         { -- ������ ���Ϻ� ������ Table �ڵ���ȸ ���� �ּ� @ 2014.08.26 LSH
         if (FormatDateTime('hh:nn', Now) > '00:00') and (FormatDateTime('hh:nn', Now) <= '08:30') then
            Result := '�� [' +  sType2 + '] ������ (����) �޴� ���� ��' + #13#10 + #13#10 + TpGetMenu.GetOutputDataS('sDutyRmk', 0)
         else if (FormatDateTime('hh:nn', Now) > '08:30') and (FormatDateTime('hh:nn', Now) <= '14:00') then
            Result := '�� [' +  sType2 + '] ������ (�߽�) �޴� ���� ��' + #13#10 + #13#10 + TpGetMenu.GetOutputDataS('sAttachNm', 0)
         else if (FormatDateTime('hh:nn', Now) > '14:00') and (FormatDateTime('hh:nn', Now) <= '19:00') then
            Result := '�� [' +  sType2 + '] ������ (����) �޴� ���� ��' + #13#10 + #13#10 + TpGetMenu.GetOutputDataS('sHideFile', 0)
         else
            Result := '';
         }


         Result := '�� [' +  sLocateNm + '] ������ (' + sDietTime + ') �޴� ���� ��' + #13#10 + #13#10 + TpGetMenu.GetOutputDataS('sDutyRmk', 0);

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

         AddComboString('�˻�');
         AddComboString('���');
         AddComboString('����');
         AddComboString('����');
      end;

      if (ARow > 0) and (ACol = C_RD_DOCLIST) then
      begin
         AEditor := edComboList;

         ClearComboString;

         AddComboString('�߼�');
         AddComboString('����');
         AddComboString('ȸ��');
         AddComboString('���ó');
         AddComboString('OCS����');
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
      // ��ư ���� �б�
      if (
            (Col = C_RD_GUBUN) or
            (Col = C_RD_DOCLIST)
         ) and (Cells[C_RD_GUBUN, Row] <> '�˻�') then
      begin
         tmpGubun := Cells[C_RD_GUBUN, Row];

         // �űԵ�Ͻ�, ���� �ֱ� �ش� Category �̷� �ڵ���ȸ @ 2015.05.11 LSH
         if (Cells[C_RD_GUBUN, Row] = '���') then
         begin
            fsbt_ComDoc.Tag          := 9999;
            Cells[C_RD_GUBUN,   Row] := '�˻�';
            Cells[C_RD_DOCYEAR, Row] := FormatDateTime('yyyy', Date);

            AddButton(C_RD_BUTTON,  Row, ColWidths[C_RD_BUTTON]-5,  20, 'Search', haBeforeText, vaCenter);           // Search

            asg_RegDoc_ButtonClick(Sender, C_RD_BUTTON, Row);

            fsbt_ComDoc.Tag          := 0;
         end;

         Cells[C_RD_GUBUN,    Row] := tmpGubun;

         AddButton(C_RD_BUTTON,  Row, ColWidths[C_RD_BUTTON]-5,  20, 'Update', haBeforeText, vaCenter);           // Update

         // ������ �׸� �ʱ�ȭ
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
               ) and (Cells[C_RD_GUBUN, Row] = '�˻�') then
      begin
         AddButton(C_RD_BUTTON,  Row, ColWidths[C_RD_BUTTON]-5,  20, 'Search', haBeforeText, vaCenter);           // Search

         // ������ �׸� �ʱ�ȭ
         Cells[C_RD_DOCSEQ,   Row] := '';
         Cells[C_RD_DOCTITLE, Row] := '';
         Cells[C_RD_REGDATE,  Row] := '';
         Cells[C_RD_REGUSER,  Row] := '';
         Cells[C_RD_RELDEPT,  Row] := '';
         Cells[C_RD_DOCRMK,   Row] := '';
         Cells[C_RD_DOCLOC,   Row] := '';

         {  -- ���ϴ� ������� �۵����� �ʾƼ� �ϴ� ���� ... @ 2015.05.12 LSH
         // �˻��� �⵵ ���� ������, �ڵ� �˻� Start @ 2015.05.11 LSH
         if (Col = C_RD_DOCYEAR) and (Cells[C_RD_DOCYEAR, Row] <> '') then
            asg_RegDoc_ButtonClick(Sender, C_RD_BUTTON, Row);
         }
      end;


   end;

end;


//------------------------------------------------------------------------------
// [�������� Update] TAdvStringGrid onButtonClick Event Handler
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
                 PChar('[��󿬶���] - [���������]�� ����� IP ����� ����Ͻ� �� �ֽ��ϴ�.'),
                 PChar(Self.Caption + ' : ��ϵ��� ���� User ���� �˸�'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;




   // ���κ� �������� Update
   if (ACol = C_RD_BUTTON) then
   begin

      // ��û���к� Input variables Assign
      if (asg_RegDoc.Cells[C_RD_GUBUN, ARow] = '�˻�') then
      begin
         sType    := '6';
         sDelDate := '';
      end
      else if ((asg_RegDoc.Cells[C_RD_GUBUN, ARow] = '���') or
               (asg_RegDoc.Cells[C_RD_GUBUN, ARow] = '����')) then
      begin
         sType    := '3';
         sDelDate := '';
      end
      else if  (asg_RegDoc.Cells[C_RD_GUBUN, ARow] = '����') then
      begin
         sType    := '3';
         sDelDate := FormatDateTime('yyyy-mm-dd', Date);
      end;


      // ������ IP �ĺ���, �ٹ�ó Assign
      if PosByte('�Ȼ굵����', FsUserIp) > 0 then
         sLocateNm := '�Ⱦ�'
      else if PosByte('���ε�����', FsUserIp) > 0 then
         sLocateNm := '����'
      else if PosByte('�Ȼ굵����', FsUserIp) > 0 then
         sLocateNm := '�Ȼ�';



      iCnt  := 0;


      //-------------------------------------------------------------------
      // 2. Create Variables
      //-------------------------------------------------------------------
      if (asg_RegDoc.Cells[C_RD_GUBUN, ARow] <> '�˻�') then
      begin
         with asg_RegDoc do
         begin

            if (Cells[C_RD_GUBUN, ARow] = '���') and
               ((Cells[C_RD_DOCLIST, ARow] = '') or
                (Cells[C_RD_DOCYEAR, ARow] = '') or
                (Cells[C_RD_DOCSEQ,  ARow] = '') or
                (Cells[C_RD_DOCTITLE,ARow] = '') or
                (Cells[C_RD_REGDATE, ARow] = '') or
                (Cells[C_RD_REGUSER, ARow] = '')) then
            begin
               MessageBox(self.Handle,
                          PChar('���� ��Ͻ� [����/�⵵/��ȣ/����/�����/�����]�� �ʼ��Է� �׸��Դϴ�.'),
                          PChar(Self.Caption + ' : �����ʼ��Է� �׸� �˸�'),
                          MB_OK + MB_ICONWARNING);

               Exit;
            end;

            if ((Cells[C_RD_GUBUN, ARow] = '����') or
                (Cells[C_RD_GUBUN, ARow] = '����')) and
               ((Cells[C_RD_DOCLIST, ARow] = '') or
                (Cells[C_RD_DOCYEAR, ARow] = '') or
                (Cells[C_RD_DOCSEQ,  ARow] = '')) then
            begin
               MessageBox(self.Handle,
                          PChar('���� ����/������ [����/�⵵/��ȣ]�� �ʼ��Է� �׸��Դϴ�.'),
                          PChar(Self.Caption + ' : �����ʼ�üũ �׸� �˸�'),
                          MB_OK + MB_ICONWARNING);

               Exit;
            end;



            // ���ó ���(����) �б� ���� @ 2014.07.18 LSH
            if Cells[C_RD_DOCLIST, ARow] = '���ó' then
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
            if (asg_RegDoc.Cells[C_RD_DOCLIST, ARow] = '���ó') then
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
                                   ' ������ ���������� [' + asg_RegDoc.Cells[C_RD_GUBUN,  asg_RegDoc.Row] + '] �Ǿ����ϴ�.'),
                             '[KUMC ���̾�α�] �������� ������Ʈ �˸� ',
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
                                   ' ������ ���������� [' + asg_RegDoc.Cells[C_RD_GUBUN,  asg_RegDoc.Row] + '] �Ǿ����ϴ�.'),
                             '[KUMC ���̾�α�] �������� ������Ʈ �˸� ',
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
      else if (asg_RegDoc.Cells[C_RD_GUBUN, ARow] = '�˻�') then
      begin

         // ��� �Ǵ� ������ �ش� Category �ӽ���ȸ Tag (9999)�ÿ��� Log ��� ���� @ 2015.05.11 LSH
         if (fsbt_ComDoc.Tag <> 9999) and
            (
               (asg_RegDoc.Cells[C_RD_DOCLIST, ARow] = '') or
               (asg_RegDoc.Cells[C_RD_DOCYEAR, ARow] = '')
             ) then
         begin
            MessageBox(self.Handle,
                       PChar('���� �˻��� [����/�⵵]�� �ʼ��Է� �׸��Դϴ�.'),
                       PChar(Self.Caption + ' : �����˻� �ʼ��Է� �׸� �˸�'),
                       MB_OK + MB_ICONWARNING);

            Exit;
         end;



         asg_SelDoc.ClearRows(1, asg_SelDoc.RowCount);
         asg_SelDoc.RowCount := 2;



         //-----------------------------------------------------------------
         // 1. �Է� ���� Assign
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
                  lb_SelDoc.Caption := '�� �ش� ȸ��⵵�� ���� �˻��ڷᰡ �����ϴ�.';
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

                  // ��ุ��� �׸��� FontColor = clRed ������Ʈ @ 2014.06.02 LSH
                  if (Cells[C_SD_DOCLIST, i+FixedRows] = '���ó') then
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

            // RowCount ����
            asg_SelDoc.RowCount := asg_SelDoc.RowCount - 1;

            // Comments
            lb_SelDoc.Caption := '�� ' + IntToStr(iCnt) + '���� ' + '[' + asg_RegDoc.Cells[C_RD_DOCYEAR, asg_RegDoc.Row] + ' ' +
                                                                          asg_RegDoc.Cells[C_RD_DOCLIST, asg_RegDoc.Row] + '] ' +
                                                                    '������ �˻��Ͽ����ϴ�.';

         finally
            FreeAndNil(TpGetDoc);
            Screen.Cursor := crDefault;
         end;


         // ��� �Ǵ� ������ �ش� Category �ӽ���ȸ Tag (9999)�ÿ��� Log ��� ���� @ 2015.05.11 LSH
         if fsbt_ComDoc.Tag <> 9999 then
         begin
            // �α� Update
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
// [��ȸ] �� ������ Max Seqno ����
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
   // 1. �Է� ���� Assign
   //-----------------------------------------------------------------
   sType1   := '7';
   sDocList := in_DocList;
   sDocYear := in_DocYear;

   // CRA/CRC �λ� �� ���α׷� �ڵ��Է� ���߿Ϸ��̳�, � D/B Transaction ����(ora-24777) �ȵǾ� �ϴ� �ּ�ó�� @ 2013.12.06 LSH
   // non-XA ���� �̰��۾���, �ϱ� ���μ��� ���� @ 2014.01.10 LSH
   // ID ������ �Է� Panel on��, max ID ä������ ��ȸ���� �б�
   if (
         (PosByte('CRA', sDocList) > 0) or
         (PosByte('CRC', sDocList) > 0)
      ) then

      sType1 := '26';



   // ������ IP �ĺ���, �ٹ�ó Assign
   if PosByte('�Ⱦϵ�����', FsUserIp) > 0 then
      sLocate := '�Ⱦ�'
   else if PosByte('���ε�����', FsUserIp) > 0 then
      sLocate := '����'
   else if PosByte('�Ȼ굵����', FsUserIp) > 0 then
      sLocate := '�Ȼ�';




   //-----------------------------------------------------------------
   // 2. ��ȸ
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
         // ID ������ �Է� Panel on��, max ID ä���� assign
         //----------------------------------------------------
         if (sType1 = '26') then
         begin
            asg_RegDoc.Cells[C_RD_DOCSEQ, asg_RegDoc.Row] := TpGetDocSeq.GetOutputDataS('sDocSeq', 0);
            Result := TpGetDocSeq.GetOutputDataS('sDocRmk', 0);
         end
         else if (sType1 = '7') then
            //----------------------------------------------------
            // �ֽ� Max ������ȣ Return
            //----------------------------------------------------
            Result := TpGetDocSeq.GetOutputDataS('sDocSeq', 0);





   finally
      FreeAndNil(TpGetDocSeq);
      Screen.Cursor := crDefault;
   end;
end;



//------------------------------------------------------------------------------
// [��ȸ] IP �� ����� �� ����
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
   varMngrNm      := '';   // ���� ������ (����) @ 2014.07.18 LSH



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
      // ���� IP + ���߻���� �� ���, ����� ����ȭ�� Call
      //------------------------------------------------------------
      if TpGetUser.RowCount > 1 then
      begin

         GetUser := TSelUser.Create(self);

         try
            // ���� User ���� ǥ��
            GetUser.lb_UserCnt.Caption := IntToStr(TpGetUser.RowCount) + '���� ����ڰ� �����մϴ�.';


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
         varMngrNm      := TpGetUser.GetOutputDataS('sDutyUser',  0);   // ���� ������ (����) @ 2014.07.18 LSH
      end;


      Result := True;


   finally
      FreeAndNil(TpGetUser);
      screen.Cursor  :=  crDefault;
   end;
end;



//------------------------------------------------------------------------------
// [��ȸ] ���� ���� �������� �Լ�
//       - �ҽ���ó : MComFunc.pas
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
      days[1] := '�Ͽ���';
      days[2] := '������';
      days[3] := 'ȭ����';
      days[4] := '������';
      days[5] := '�����';
      days[6] := '�ݿ���';
      days[7] := '�����';
   end
   else if Type1 = 'HS' then
   begin
      days[1] := '��';
      days[2] := '��';
      days[3] := 'ȭ';
      days[4] := '��';
      days[5] := '��';
      days[6] := '��';
      days[7] := '��';
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

            lb_RegDoc.Caption := '�� 2010�� �������� D/B�� �������Դϴ�. 4�ڸ� ȸ��⵵�� �Է����ּ���.';
         end
         else if (ACol = C_RD_GUBUN) then
            lb_RegDoc.Caption := '�� �μ����빮��(�߼�/���� ��)�� �˻� �Ǵ� ���(����)�Ͻ� �� �ֽ��ϴ�.'
         else if (ACol = C_RD_DOCSEQ) then
            lb_RegDoc.Caption := '�� [Ŭ��]�Ͻø� �ڵ����� ������ȣ�� ä���մϴ�. (����/������ ������ȣ�� �ʿ�)'
         else if (ACol = C_RD_DOCTITLE) and (Cells[C_RD_GUBUN, ARow] = '�˻�') then
            lb_RegDoc.Caption := '�� �˻��� ������ ���Ե� �ܾ�(��: ��Ʈ��ũ, ���� ��)�� �����˻��� �����մϴ�.'
         else if (ACol = C_RD_REGUSER) then
         begin
            if Cells[C_RD_GUBUN, ARow] = '���' then
               Cells[C_RD_REGUSER, ARow] := FsUserNm;

            lb_RegDoc.Caption := '�� �߼��� ��� ���� �����, ������ ��� ������(�Ǵ� ����������)�� �����ּ���.'
         end
         else if (ACol = C_RD_RELDEPT) then
            lb_RegDoc.Caption := '�� �ش� ������ �޴� �μ���(�Ǵ� �����μ�)��, �Ǵ� ����� �μ����� �����ּ���.'
         else if (ACol = C_RD_DOCRMK) then
            lb_RegDoc.Caption := '�� �ش� ������ ���õ� Tag �Ǵ� �ٹ�ó/�����/����ó �� Ư�̻����� �����ּ���.'
         else if (ACol = C_RD_DOCLOC) then
            lb_RegDoc.Caption := '�� ������ ��Ϲ����� ������ġ�� default�� �����ּ���.'
         else
            lb_RegDoc.Caption := '';




         // �ű� [���]�� ���, ������ȣ �ֽ� Seqno ��������
         if (ACol = C_RD_DOCSEQ) and
            (Cells[C_RD_GUBUN, ARow] = '���') then
         begin
            Cells[ACol, ARow] := GetMaxDocSeq(Cells[C_RD_DOCLIST, ARow],
                                              Cells[C_RD_DOCYEAR, ARow]);

         end;


         // ora-24777 : use of non-migratable data base link not allowed ] ���� (XA�� DB link ���û�� �ȵ�)��
         // ���߿Ϸ���, �ش� ����(md_kumcm_m1) non-XA ���� �̰��۾���, ���� @ 2014.01.10 LSH
         if (ACol = C_RD_DOCSEQ) or
            (ACol = C_RD_DOCYEAR) then
         begin
            // �ӻ󿬱���(CRA/CRC) ID ���(�Ǵ� ����)��, ������ �Է� Panel Open
            if (
                  (Cells[C_RD_GUBUN, ARow] = '���') or
                  (Cells[C_RD_GUBUN, ARow] = '����')
               ) and
               (
                  (PosByte('CRA', Cells[C_RD_DOCLIST, ARow]) > 0) or
                  (PosByte('CRC', Cells[C_RD_DOCLIST, ARow]) > 0)
               ) then
            begin
               asg_DocSpec.Cells[0, R_DS_USERID]   := 'ID';
               asg_DocSpec.Cells[0, R_DS_USERNM]   := '����';
               asg_DocSpec.Cells[0, R_DS_DEPTCD]   := '�Ҽ��ڵ�';
               asg_DocSpec.Cells[0, R_DS_DEPTNM]   := '�ҼӸ�';
               asg_DocSpec.Cells[0, R_DS_SOCNO]    := '�ֹε�Ϲ�ȣ';
               asg_DocSpec.Cells[0, R_DS_STARTDT]  := '��������';
               asg_DocSpec.Cells[0, R_DS_ENDDT]    := '��������';
               asg_DocSpec.Cells[0, R_DS_MOBILE]   := '�޴���';
               asg_DocSpec.Cells[0, R_DS_GRPID]    := '���α׷�ID';
               asg_DocSpec.Cells[0, R_DS_USELVL]   := '���α׷�����';


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

               apn_DocSpec.Caption.Text := Cells[C_RD_DOCLIST, ARow] + ' ID ������ ���(����)';
            end
            // ���ó ������ �Է� Panel Open
            else if
                     (ACol = C_RD_DOCSEQ) and
                     (
                        (Cells[C_RD_GUBUN, ARow] = '���') or
                        (Cells[C_RD_GUBUN, ARow] = '����')
                     )  and
                     (
                        (PosByte('���ó', Cells[C_RD_DOCLIST, ARow]) > 0)
                     ) then
            begin
               asg_DocSpec.Cells[0, R_DS_USERID]   := '����';
               asg_DocSpec.Cells[0, R_DS_USERNM]   := '����ü��';
               asg_DocSpec.Cells[0, R_DS_DEPTCD]   := '���Ⱓ';
               asg_DocSpec.Cells[0, R_DS_DEPTNM]   := '';
               asg_DocSpec.Cells[0, R_DS_SOCNO]    := '�д��(�Ѿ�)';
               asg_DocSpec.Cells[0, R_DS_STARTDT]  := '�д��(HQ)';
               asg_DocSpec.Cells[0, R_DS_ENDDT]    := '�д��(AA)';
               asg_DocSpec.Cells[0, R_DS_MOBILE]   := '�д��(GR)';
               asg_DocSpec.Cells[0, R_DS_GRPID]    := '�д��(AS)';
               asg_DocSpec.Cells[0, R_DS_USELVL]   := '������ġ';


               asg_DocSpec.Cells[1, R_DS_USERID]   := '��) �������� S/W �������� ����';
               asg_DocSpec.Cells[1, R_DS_USERNM]   := '��) (��)��ť��';
               asg_DocSpec.Cells[1, R_DS_DEPTCD]   := '��) 2014. 3. 1 ~ 2015. 2. 28 (12����)';
               asg_DocSpec.Cells[1, R_DS_DEPTNM]   := '';
               asg_DocSpec.Cells[1, R_DS_SOCNO]    := '��) 375,000��';
               asg_DocSpec.Cells[1, R_DS_STARTDT]  := '��) 0��';
               asg_DocSpec.Cells[1, R_DS_ENDDT]    := '��) 146,300��';
               asg_DocSpec.Cells[1, R_DS_MOBILE]   := '��) 142,500��';
               asg_DocSpec.Cells[1, R_DS_GRPID]    := '��) 86,200��';
               asg_DocSpec.Cells[1, R_DS_USELVL]   := '��) PDF, ���ó';

               apn_DocSpec.Top     := 155;
               apn_DocSpec.Left    := 140;
               apn_DocSpec.Collaps := True;
               apn_DocSpec.Visible := True;
               apn_DocSpec.Collaps := False;

               apn_DocSpec.Caption.Text := Cells[C_RD_DOCLIST, ARow] + ' �������� ������ ���(����)';
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
      if (asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row]   <> '��') and
         (asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row+1] <> '��') and
         (asg_Board.Cells[C_B_CLOSEYN,  asg_Board.Row] = 'C') then
      begin
         asg_Board.InsertRows(asg_Board.Row + 1, 1 + StrToInt(asg_Board.Cells[C_B_REPLY, asg_Board.Row]));
         asg_Board.Cells[C_B_CLOSEYN, asg_Board.Row]     := 'O';


         // ÷������ ������ ����, �� Row Merege ���� �б�
         if (asg_Board.Cells[C_B_HIDEFILE, asg_Board.Row] = '') then
            asg_Board.MergeCells(C_B_CATEUP,  asg_Board.Row + 1, 4, 1)
         else
         begin
            asg_Board.MergeCells(C_B_CATEUP,  asg_Board.Row + 1, 2, 1);
            asg_Board.MergeCells(C_B_REGDATE, asg_Board.Row + 1, 2, 1);
         end;

         // ÷�� �̹��� FTP ���ϳ����� �����ϸ�, �̹��� Download
         if (asg_Board.Cells[C_B_HIDEFILE, asg_Board.Row] <> '') then
         begin
            // ���� ��/�ٿ�ε带 ���� ���� ��ȸ
            if not GetBinUploadInfo(sServerIp, sFtpUserID, sFtpPasswd, sFtpHostName, sFtpDir) then
            begin
               MessageDlg('���� ������ ���� ����� ���� ��ȸ��, ������ �߻��߽��ϴ�.', Dialogs.mtError, [Dialogs.mbOK], 0);
               exit;
            end;


            // ��������� ���� IP
            sServerIp := C_KDIAL_FTP_IP;


            // ���� ������ ����Ǿ� �ִ� ���ϸ� ����
            if PosByte('/ftpspool/KDIALFILE/',asg_Board.Cells[C_B_HIDEFILE, asg_Board.Row]) > 0 then
               sRemoteFile := asg_Board.Cells[C_B_HIDEFILE, asg_Board.Row]
            else
               sRemoteFile := '/ftpspool/KDIALFILE/' + asg_Board.Cells[C_B_HIDEFILE, asg_Board.Row];

            // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
            sLocalFile  := G_HOMEDIR + 'TEMP\SPOOL\' + asg_Board.Cells[C_B_ATTACH, asg_Board.Row];



            if (GetBINFTP(sServerIp,
                          sFtpUserID,
                          sFtpPasswd,
                          sRemoteFile,
                          sLocalFile,
                          False)) then
            begin
               //	�������� FTP �ٿ�ε�
            end;



            // Local�� �ش� Image �������� üũ (Local file�� Size üũ �߰� : Local�� �ش� ���.jpg�� 0 size�� �����ϸ� JPEG Error #41/#42 ������) @ 2019.12.03 LSH
            if (FileExists(sLocalFile)) and
               (CMsg.GetFileSize(sLocalFile) > 0) then
            begin
               // �̹����� Preview ����
               if (PosByte('.bmp', sLocalFile) > 0) or
                  (PosByte('.BMP', sLocalFile) > 0) or
                  (PosByte('.jpg', sLocalFile) > 0) or
                  (PosByte('.JPG', sLocalFile) > 0) or
                  {  -- TMS Grid���� png Ÿ�� ���ν� ������ �ּ� @ 2017.11.01 LSH
                  (PosByte('.png', sLocalFile) > 0) or
                  (PosByte('.PNG', sLocalFile) > 0) or
                  }
                  (PosByte('.gif', sLocalFile) > 0) or
                  (PosByte('.GIF', sLocalFile) > 0) then
               begin
                  // ������ Image ������ Grid�� ǥ��
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
            // ������ �̹��� Remove
            asg_Board.RemovePicture(C_B_REGDATE, asg_Board.Row + 1);
         end;




         asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row + 1]   := '��';
         asg_Board.Cells[C_B_TITLE,    asg_Board.Row + 1]   := asg_Board.Cells[C_B_CONTEXT, asg_Board.Row];
         asg_Board.AddButton(C_B_LIKE, asg_Board.Row + 1, asg_Board.ColWidths[C_B_LIKE]-5,  20, '��', haBeforeText, vaCenter);
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

         // �α� Update
         UpdateLog('BOARD',
                   asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row],
                   asg_Board.Cells[C_B_USERIP,   asg_Board.Row],
                   'S',
                   '',
                   '',
                   varResult
                   );


      end
      else if (asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row]   <> '��') and
              (asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row+1] = '��') and
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
         asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row + 1]   := '��';
         asg_Board.Cells[C_B_TITLE,    asg_Board.Row + 1]   := asg_Board.Cells[C_B_CONTEXT, asg_Board.Row];
         asg_Board.AddButton(C_B_LIKE, asg_Board.Row + 1, asg_Board.ColWidths[C_B_LIKE]-5,  20, '��', haBeforeText, vaCenter);
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


         // �α� Update
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




      // [Best] �� �ԽñⰣ Update
      if (PosByte('[��BEST��]', asg_Board.Cells[C_B_TITLE, asg_Board.Row]) > 0) then
      begin
         // �Խ��� (Ŀ�´�Ƽ) Best �ԽñⰣ(����Ŭ�� ~ 7�ϰ�) Update.
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
   // ���� Panel Off
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
            // ��ϵ� User ���̾� Chat ����Ʈ �˾� @ 2015.03.30 LSH
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


      { -- ��󿬶��� Tab ���տ� ���� �ּ� @ 2013.12.30 LSH
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

            // ���� [��������] Ŭ���ÿ��� Release ���� ��ȸ�ϰ�(Tag = 0)
            // ���� ���ʹ� ���������� Ȱ��ȭ �ߴ� Panel�� ��쵵�� ���� @ 2015.06.15 LSH
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

            // �Խ��� Page Block ���� ��ȸ
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

            // [�Ⱦ�] ������ IP �ĺ���, �ٹ�ó ���̾� Map ����ǥ�� @ 2013.11.05 LSH
            if PosByte('�Ⱦϵ�����', FsUserIp) > 0 then
               fsbt_DialMap.Visible := True
            else
               fsbt_DialMap.Visible := False;

            // ���� ���̾� Book ��ȸ
            SelGridInfo('MYDIAL');

            tm_Master.Enabled := False;

            fcb_Scan.Text := '���հ˻�';
            fed_Scan.Text := '';
            fed_Scan.CanFocus;
            fed_Scan.SetFocus;

            lb_DialScan.Caption := '�� [��ȭ��ȣ��], [���̾�α�] �� [S/R��û����]�� ������ ���հ˻��� �����մϴ� ^^b';

            //-----------------------------------------------------
            // ��ϵ� User ���̾� Chat ����Ʈ �˾� @ 2015.03.30 LSH
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


            // D/B ������ �������� [�/������Ʈ]�� ����
            if (PosByte('�', FsUserPart) > 0) or
               (PosByte('����', FsUserPart) > 0) then
               fsbt_DBMaster.Visible := True
            else
               fsbt_DBMaster.Visible := False;


            {
            asg_WorkRpt.Top      := 999;
            asg_WorkRpt.Left     := 999;
            asg_WorkRpt.Visible  := False;


            fmed_AnalFrom.Text   := FormatDateTime('yyyy-mm-dd', DATE - 1095);   // �� 3��ġ �˻�
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

            lb_Analysis.Caption := '�� S/R ��û���� �м��� ���� ȿ������ �˻��� �����մϴ�.';
            }

            fcb_Analysis.Text := '���հ˻�';
            fed_Analysis.Text := '';
            fed_Analysis.CanFocus;
            fed_Analysis.SetFocus;


            //fsbt_AnalysisClick(Sender);
            fsbt_WorkRptClick(Sender);


            tm_Master.Enabled := False;



         end;

      //------------------------------------------------------------------------
      // KKP ����� order�� ������ �������� ���/��ȸ ���� @ 2014.08.08 LSH
      //------------------------------------------------------------------------
      {  -- ���� ���Ϸ� ���� @ 2017.10.31 LSH
      AT_WORKCONN :
         begin
            fpn_WorkConn.Top  := 23;
            fpn_WorkConn.Left := 7;
            fpn_WorkConn.BringToFront;


//            fcb_Analysis.Text := '���հ˻�';
//            fed_Analysis.Text := '';
//            fed_Analysis.CanFocus;
//            fed_Analysis.SetFocus;
//
//
//            //fsbt_AnalysisClick(Sender);
//            fsbt_WorkRptClick(Sender);



            // �Ⱓ�� �ֿ�������� ��ȸ
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
                 PChar('[��󿬶���] - [���������]�� ����� IP ����� ����Ͻ� �� �ֽ��ϴ�.'),
                 PChar(Self.Caption + ' : ��ϵ��� ���� User ���� �˸�'),
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
   fmm_Text.Hint      := '�ű� �ۼ����� �ѱ۱��� 15��(500�� ����)�̳��� �����մϴ�.';

   lb_Write.Caption    := '�� �ű� �� �ۼ� ����Դϴ�.';
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
                 PChar('�� �ű� ��Ͻ� [�з�]�׸��� �ʼ��Է� �Դϴ�.'),
                 PChar(Self.Caption + ' : Ŀ�´�Ƽ �ʼ��Է� �׸� �˸�'),
                 MB_OK + MB_ICONWARNING);

      fcb_CateUp.SetFocus;

      Exit;
   end;

   {  -- ���Ѿ��� �������� ������ �� �ְ� �ϱ� ���� �ּ� @ 2018.03.20 LSH
   if (lb_HeadTail.Caption = 'H') and
      (PosByte('����', fcb_CateUp.Text) > 0) and
      (FsUserIp <> '������IP') then
   begin
      MessageBox(self.Handle,
                 PChar('[��������] ����� �����ڸ� �����մϴ�. ���� ��Ź����� ^^'),
                 PChar(Self.Caption + ' : Ŀ�´�Ƽ �������� ������� �˸�'),
                 MB_OK + MB_ICONWARNING);

      fcb_CateUp.SetFocus;

      Exit;
   end;
   }

   if (lb_HeadTail.Caption = 'H') and
      (PosByte('����', fcb_CateUp.Text) > 0) and
      ((fmed_AlarmFro.Text = '    -  -  ') or (fmed_AlarmTo.Text = '    -  -  '))  then
   begin
      MessageBox(self.Handle,
                 PChar('[��������] ��Ͻ�, �����Ⱓ (From~To) �ʼ��Է� �Դϴ�.'),
                 PChar(Self.Caption + ' : Ŀ�´�Ƽ �������� �Ⱓ��� �˸�'),
                 MB_OK + MB_ICONWARNING);

      fmed_AlarmFro.SetFocus;

      Exit;
   end;

   if (lb_HeadTail.Caption = 'H') and
      (PosByte('����', fcb_CateUp.Text) > 0) and
      (fmed_AlarmFro.Text  > fmed_AlarmTo.Text)  then
   begin
      MessageBox(self.Handle,
                 PChar('[��������] ��Ͻ�, �����Ⱓ To�� From���� �ռ� �� �����ϴ�.'),
                 PChar(Self.Caption + ' : Ŀ�´�Ƽ �������� �Ⱓ��� �˸�'),
                 MB_OK + MB_ICONWARNING);

      fmed_AlarmTo.SetFocus;

      Exit;
   end;


   if (fed_Writer.Text = '') then
   begin
      MessageBox(self.Handle,
                 PChar('�ۼ���(�г���)�� �Է��� �ֽñ� �ٶ��ϴ�.'),
                 PChar(Self.Caption + ' : Ŀ�´�Ƽ �ʼ��Է� �׸� �˸�'),
                 MB_OK + MB_ICONWARNING);

      fed_Writer.SetFocus;

      Exit;
   end;

   if (fed_Title.Text = '') then
   begin
      MessageBox(self.Handle,
                 PChar('[�� ����]�� �Է��� �ֽñ� �ٶ��ϴ�.'),
                 PChar(Self.Caption + ' : Ŀ�´�Ƽ �ʼ��Է� �׸� �˸�'),
                 MB_OK + MB_ICONWARNING);

      fed_Title.SetFocus;

      Exit;
   end;

   if (fmm_Text.Lines.Text = '') then
   begin
      MessageBox(self.Handle,
                 PChar('[����]�� �Է��Ͻ� ������ �����ϴ�.' + #13#10 + '�ƹ� �����̶� �ϳ� �����ּ��� ^^'),
                 PChar(Self.Caption + ' : Ŀ�´�Ƽ �ʼ��Է� �׸� �˸�'),
                 MB_OK + MB_ICONWARNING);

      fmm_Text.SetFocus;

      Exit;
   end;

   if (lb_HeadTail.Caption = 'H') then
   begin
      if (fmm_Text.Lines.Count > 15) then
      begin
         MessageBox(self.Handle,
                    PChar('�ű� ���ۼ��� 15�� �̳��� �����մϴ�.'),
                    PChar(Self.Caption + ' : Ŀ�´�Ƽ �� �ۼ� ���� �˸�'),
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
                    PChar('����ۼ��� ��ũ���� �й��� �����ϱ����� 5�� �̳��� �����մϴ�.'),
                    PChar(Self.Caption + ' : Ŀ�´�Ƽ �� �ۼ� ���� �˸�'),
                    MB_OK + MB_ICONWARNING);

         fmm_Text.SetFocus;

         Exit;
      end;
   end;


   // �������� Flag
   if (PosByte('����', fcb_CateUp.Text) > 0) and
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


   // ÷������ Upload
   if (Trim(fed_Attached.Text) <> '') then
   begin

      sFileName := ExtractFileName(Trim(fed_Attached.Text));
      sHideFile := 'KDIALAPPEND' + Trim(fed_BoardSeq.Text) + FormatDateTime('YYYYMMDDHHNNSS', Now) + '.KDIALFILE';


      if Not FileUpLoad(sHideFile, asg_Board.Cells[C_B_ATTACH, asg_Board.Row], sServerIp) then
      begin
         Showmessage('÷������ UpLoad �� ������ �߻��߽��ϴ�.' + #13#10 + #13#10 +
                     '�ٽ��ѹ� �õ��� �ֽñ� �ٶ��ϴ�.');
         Exit;
      end;
   end;



   // �Խ��� (Ŀ�´�Ƽ) Update.
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
                 PChar('[��󿬶���] - [���������]�� ����� IP ����� ����Ͻ� �� �ֽ��ϴ�.'),
                 PChar(Self.Caption + ' : ��ϵ��� ���� User ���� �˸�'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;

   fpn_Write.Top     := 0;
   fpn_Write.Left    := 0;
   fpn_Write.Visible := True;
   fpn_Write.BringToFront;


   fed_BoardSeq.Text    := '';
   fed_Title.Text       := '[���] ' + asg_Board.Cells[C_B_TITLE, asg_Board.Row];
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
   fmm_Text.Hint      := '��ũ���� �й��� ���ϱ����� ����� 4~5��(�ѱ۱��� 200��)�̳��� �����մϴ�.';

   lb_Write.Caption    := '�� ��� �ۼ� ����Դϴ�.';
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
                 PChar('[��󿬶���] - [���������]�� ����� IP ����� ����Ͻ� �� �ֽ��ϴ�.'),
                 PChar(Self.Caption + ' : ��ϵ��� ���� User ���� �˸�'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;

   fed_BoardSeq.Text    := asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row];
   fed_Title.Text       := '';
   lb_HeadTail.Caption  := '';
   lb_HeadSeq.Caption   := '';
   lb_TailSeq.Caption   := '';


   // �ۼ��� ���� IP Ȯ��
   if (FsUserIp <> asg_Board.Cells[C_B_USERIP, asg_Board.Row]) then
   begin
      MessageBox(self.Handle,
                 PChar('�ۼ��� ������ �Խñ۸� ������ �����մϴ�.'),
                 '[KUMC ���̾�α�] Ŀ�´�Ƽ ���� Ȯ�� �˸� ',
                 MB_OK + MB_ICONINFORMATION);

      Exit;
   end;


   // ���� Confirm Message.
   if Application.MessageBox('������ �Խñ��� [����]�Ͻðڽ��ϱ�?',
                             PChar(self.Caption + ' : Ŀ�´�Ƽ ���� �˸�'), MB_OKCANCEL) <> IDOK then
      Exit;


   if (asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row] <> '��') and
      (asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row] <> '')   then
   begin
      // �Խ��� (Ŀ�´�Ƽ) ������ ���� Update.
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
      // �Խ��� (Ŀ�´�Ƽ) ��� ���� Update.
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
// [��ȸ] �Խ��� (Ŀ�´�Ƽ) Update Procedure
//       - �ű�/���/���� Upd.
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
                          PChar('[' + in_Title + '] ' + ' �Խñ��� ���������� [������Ʈ] �Ǿ����ϴ�.'),
                          '[KUMC ���̾�α�] Ŀ�´�Ƽ ������Ʈ �˸� ',
                          MB_OK + MB_ICONINFORMATION)
            else
               MessageBox(self.Handle,
                          PChar('�Խñ��� ���������� [����] �Ǿ����ϴ�.'),
                          '[KUMC ���̾�α�] Ŀ�´�Ƽ ������Ʈ �˸� ',
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
// [��ȸ] �Խ��� (Ŀ�´�Ƽ) ��� ����ȸ
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
            //stb_Message.Panels[0].Text := '�ش�Ⱓ�� �˻��� �ڷᰡ �����ϴ�.';
            Exit;
         end;



      with asg_Board do
      begin
         iRowCnt  := TpGetReply.RowCount;

         for i := 0 to iRowCnt - 1 do
         begin
            Cells[C_B_BOARDSEQ,  in_HeaderRow + 2 + i] := '';
            Cells[C_B_CATEUP,    in_HeaderRow + 2 + i] := '��';
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
            // ��õ ��ư ǥ��
            //------------------------------------------------------------
            //AddButton(C_B_ATTACH, in_HeaderRow + 2 + i, ColWidths[C_B_ATTACH]-5,  20, '��', haBeforeText, vaCenter);
            AddButton(C_B_LIKE, in_HeaderRow + 2 + i, ColWidths[C_B_LIKE]-5,  20, '��', haCenter, vaUnderText);


            //------------------------------------------------------------
            // ÷�� ���� ǥ��
            //------------------------------------------------------------
            if TpGetReply.GetOutputDataS('sAttachNm', i) <> '' then
               AddButton(C_B_ATTACH,  in_HeaderRow + 2 + i, ColWidths[C_B_ATTACH]-5,  20, '��', haBeforeText, vaCenter);


            // ÷�����ϸ� ���̸�ŭ, AutoSizing �Ǵ� ������ �����ϱ� ���� �ּ�ó�� --> ��õȽ�� + ��ưǥ�� ���� �ٽ� �ּ� ����
            AutoSizeRow(in_HeaderRow + 2 + i);
         end;
      end;




   finally
      FreeAndNil(TpGetReply);
      Screen.Cursor := crDefault;
   end;

end;




//------------------------------------------------------------------------------
// [Update] �Խ��� Log ����
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
                       PChar('������ �Խñ��� ���������� [��õ] �Ǿ����ϴ�.'),
                       '[KUMC ���̾�α�] Ŀ�´�Ƽ ������Ʈ �˸� ',
                       MB_OK + MB_ICONINFORMATION);

         varResult := 'Y';
      end
      else
      begin
         // SVC �ܿ��� GetTxMsg ó����. ���� ó�� �ʿ� ����.
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
      (asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row] = '��') then
      fsbt_Reply.Enabled := False
   else
      fsbt_Reply.Enabled := True;


   if (asg_Board.Cells[C_B_USERIP, asg_Board.Row] = '') then
      fsbt_Delete.Enabled := False
   else
      fsbt_Delete.Enabled := True;

   if (asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row] <> '') and
      (asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row] <> '��') then
      lb_Board.Caption := '�� No.' + asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row] + ' ' + asg_Board.Cells[C_B_TITLE, asg_Board.Row];

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
                 PChar('[��󿬶���] - [���������]�� ����� IP ����� ����Ͻ� �� �ֽ��ϴ�.'),
                 PChar(Self.Caption + ' : ��ϵ��� ���� User ���� �˸�'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;


   with asg_Board do
   begin
      if (ACol = C_B_LIKE) then
         // or ((ACol = C_B_ATTACH) and (Cells[C_B_BOARDSEQ, ARow] = '')) then
      begin
         if Cells[C_B_BOARDSEQ, ARow] = '��' then
            UpdateLog('BOARD',
                      Cells[C_B_BOARDSEQ, ARow - 1],
                      FsUserIp{Cells[C_B_USERIP,   ARow - 1]},
                      'L',
                      '',
                      FsUserNm,
                      varUpResult
                      )
         else if Cells[C_B_CATEUP, ARow] = '��' then
            UpdateLog('BOARD',
                      Cells[C_B_CONTEXT, ARow],
                      FsUserIp{Cells[C_B_USERIP,  ARow]},
                      'L',
                      '',
                      FsUserNm,
                      varUpResult);

      end;



      //---------------------------------------------
      // ÷������ Click�� Event ó��
      //---------------------------------------------
      if ((ACol = C_B_ATTACH) and (Cells[C_B_HIDEFILE, ARow] <> '')) then
      begin
         // ���� ��/�ٿ�ε带 ���� ���� ��ȸ
         if not GetBinUploadInfo(sServerIp, sFtpUserID, sFtpPasswd, sFtpHostName, sFtpDir) then
         begin
            MessageDlg('���� ������ ���� ����� ���� ��ȸ��, ������ �߻��߽��ϴ�.', Dialogs.mtError, [Dialogs.mbOK], 0);
            exit;
         end;


         sServerIp := Cells[C_B_SERVERIP, ARow]; //lb_ServerIp.Caption;    // ��������� ���� IP


         // ���� ������ ����Ǿ� �ִ� ���ϸ� ����
         if PosByte('/ftpspool/KDIALFILE/',Cells[C_B_HIDEFILE, ARow]) > 0 then
            sRemoteFile := Cells[C_B_HIDEFILE, ARow]
         else
            sRemoteFile := '/ftpspool/KDIALFILE/' + Cells[C_B_HIDEFILE, ARow];

         // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
         sLocalFile  := G_HOMEDIR + 'TEMP\SPOOL\' + Cells[C_B_ATTACH, ARow]; //lb_Filename.Caption;



         if (GetBINFTP(sServerIp, sFtpUserID, sFtpPasswd, sRemoteFile, sLocalFile, False)) then
         begin
            //		ShowMessage('���������� ����Ǿ����ϴ�.');
         end;



         Screen.Cursor := crDefault;



         try
            if (PosByte('.exe', sLocalFile) > 0) or
               (PosByte('.zip', sLocalFile) > 0) or
               (PosByte('.rar', sLocalFile) > 0) then
            begin
               MessageBox(self.Handle,
                          PChar('÷������(' + sLocalFile + ') �ٿ�ε尡 �Ϸ�Ǿ����ϴ�.' + #13#10 + #13#10 +
                                '�� �ӽ� �ٿ�ε� ���� --> C:\KUMC(_DEV)\TEMP\SPOOL\'),
                          '÷������ �ٿ�ε� ���� �Ϸ�',
                          MB_OK + MB_ICONINFORMATION);
            end
            else
            begin
               ShellExecute(HANDLE, 'open',
                            PCHAR(Cells[C_B_ATTACH, ARow]),
                            PCHAR(''),
                            // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
                            PCHAR(G_HOMEDIR + 'TEMP\SPOOL\'),
                            SW_SHOWNORMAL);
            end;

         except
            MessageBox(self.Handle,
                       PChar('�ش� ���α׷� ������ ������ �߻��Ͽ����ϴ�.' + #13#10 + #13#10 +
                             '���α׷� ���� �� �ٽ� ������ �ֽñ� �ٶ��ϴ�.'),
                       '÷������ �ٿ�ε� ���� ����',
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


   if (PosByte('����', fcb_CateUp.Text) > 0) or
      (PosByte('����', fcb_CateUp.Text) > 0) then
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


   if (PosByte('����', fcb_CateUp.Text) > 0) then
      {and (FsUserIp = '������IP') then}          // �������� ���� ������ �� �� �ֵ��� �ּ� @ 2018.03.20 LSH
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
         // ���α׷� version Ȯ��.
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
                 PChar('[��󿬶���] - [���������]�� ����� IP ����� ����Ͻ� �� �ֽ��ϴ�.'),
                 PChar(Self.Caption + ' : ��ϵ��� ���� User ���� �˸�'),
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
   fed_Title.Hint       := '�� [����] ����Ű���� �˻� ����)' + #13#10 + #13#10 + '����Ư�� (��ĭ ����) ���� (����)';
   fed_Writer.Text      := '';
   fmm_Text.Lines.Text  := '';
   fmm_Text.Hint        := '�� [����] ����Ű���� �˻� ����)' + #13#10 + #13#10 + '���� (��ĭ ����) ���Ǽ� (�Է��� Ctrl + ���ͽ� �ٷΰ˻� ����Ű)';
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

   lb_Write.Caption    := '�� �� �˻� ����Դϴ�.';
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


   // �α� Update
   UpdateLog('BOARD',
             tmpCateUp,
             FsUserIp,
             'F',
             fed_Title.Text,
             fmm_Text.Lines.Text,
             varResult
             );


   fsbt_WriteCancelClick(Sender);

   // ���� Focus�� Grid�� ���� (�ٷ� ����Ű Ctrl+F ��밡�� �ϵ���) @ 2016.11.18 LSH
   asg_Board.SetFocus;
end;

procedure TMainDlg.fmm_TextExit(Sender: TObject);
begin
   if (lb_HeadTail.Caption = 'H') then
   begin
      if (fmm_Text.Lines.Count > 15) then
      begin
         MessageBox(self.Handle,
                    PChar('�ű� ���ۼ��� 15�� �̳��� �����մϴ�.'),
                    PChar(Self.Caption + ' : Ŀ�´�Ƽ �� �ۼ� ���� �˸�'),
                    MB_OK + MB_ICONWARNING);

         fmm_Text.SetFocus;
      end;
   end
   else if (lb_HeadTail.Caption = 'T') then
   begin
      if (fmm_Text.Lines.Count > 5) then
      begin
         MessageBox(self.Handle,
                    PChar('����ۼ��� ��ũ���� �й��� �����ϱ����� 5�� �̳��� �����մϴ�.'),
                    PChar(Self.Caption + ' : Ŀ�´�Ƽ ��� �ۼ� ���� �˸�'),
                    MB_OK + MB_ICONWARNING);

         fmm_Text.SetFocus;
      end;
   end;

   // Event �߻���, ���۽��� [�������] �����ذ����� ó��
   Application.ProcessMessages;

end;



//------------------------------------------------------------------------------
// �ѿ�Ű�� �ѱ�Ű�� ��ȯ
//       - �ҽ���ó : MComFunc.pas
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
      // �� �˻� ��忡��, [�ۼ���] ���� Ű���� �Է��� <Enter> ������ �ڵ� �˻� �ǽ� @ 2016.11.18 LSH
      if Key = VK_RETURN then
         fsbt_WriteFindClick(Sender);
   end;
end;




//------------------------------------------------------------------------------
// [��ȸ] �˶� �˾� �޼��� ����
//       - �Խ��� ��ȿ�Ⱓ�� [��������] ��ȸ
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
   // 2. ��ȸ
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
         // �˶� ����  Return
         //----------------------------------------------------
         tmpInfo := '�� ' + IntToStr(TpGetAlarm.RowCount) + '���� �˸� ��' + #13#10;


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
      // ���� Cell ��ǥ ��������
      MouseToCell(X, Y, NowCol, NowRow);


      if (NowRow > 0) and
         (
            (
               (NowCol = C_B_TITLE) and
               (Cells[C_B_HEADTAIL, NowRow] = 'H')       // �Խ��� ���� @ 2016.11.18 LSH
            ) or
            (NowCol = C_B_ATTACH)                        // ÷�� ���� ��Ī�� @ 2016.11.18 LSH
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


   // ���� ���ε带 ���� ���� ��ȸ
   if not GetBinUploadInfo(sServerIp, sFtpUserID, sFtpPasswd, sFtpHostName, sFtpDir) then
   begin
      ShowMessage('���� ������ ���� �������� ��ȸ��, ������ �߻��߽��ϴ�.');
      Result := False;
      exit;
   end;

   try
      if Id_FTP.Connected then
         Id_FTP.Disconnect;
   except
      Showmessage('�ʱ⼳�� ����');
      Result := False;
      Exit;
   end;


   Id_FTP.Host       := sServerIp;
   Id_FTP.UserName   := sFtpUserId;
   Id_FTP.Password   := sFtpPasswd;


   try
      Id_FTP.Connect;

   except
      Showmessage('FTP ���ῡ��');
      Result := False;
      Exit;
   end;


   try
      // Chat �ڽ� ���� ���۸���̸�, Path ����
      if (IsFileSend) then
         Id_FTP.ChangeDir('/ftpspool/CHATFILE')
      else
         Id_FTP.ChangeDir('/ftpspool/KDIALFILE')

   except
      Showmessage('ChangeDir ����');
      Result := False;
      Exit;
   end;


   try
      Id_FTP.TransferType := ftBinary;

   except
      Showmessage('Mode ���� ����');
      Result := False;
      Exit;
   end;


   {  -- ÷�γ��� ������� ������.. �ּ�ó�� @ 2013.08.28 LSH
   if DelFile <> '' then
   begin
      if Trim(fed_Attached.Text) <> DelFile then
      begin
         try
            Id_FTP.Delete(DelFile);

         except
            Showmessage('���ϻ��� ����');
            Result := False;
            Exit;
         end;
      end;
   end;
   }


   //--------------------------------------------------------
   // Chat �ڽ� ���� ���۸�� �б�
   //--------------------------------------------------------
   if (IsFileSend) then
   begin

      if Trim(asg_ChatSend.Cells[C_SD_TEXT, 0]) <> '' then
      begin
         try
            Id_FTP.Put(Trim(asg_ChatSend.Cells[C_SD_TEXT, 0]), TargetFile);

         except
            Showmessage('�������ۿ���');
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
      // Ŀ�´�Ƽ [÷��]
      if Trim(fed_Attached.Text) <> '' then
      begin
         try
            Id_FTP.Put(Trim(fed_Attached.Text), TargetFile);

         except
            Showmessage('[Ŀ�´�Ƽ] �������ۿ���');
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
      // ��󿬶��� ������ �̹��� [÷��]
      begin
         try
            Id_FTP.Put(DelFile, TargetFile);

         except
            Showmessage('[��Ÿ] �������ۿ���');
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
      Showmessage('��ü���� ����');
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
// [��ȸ] �Խ��� (Ŀ�´�Ƽ) ��õ List ����ȸ
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
         // ��õ�� Photo - AdvStringGrid �������� ��ȯ @ 2015.04.09 LSH
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



            // ��õ User �̹��� ǥ�� ����
            tmp_UsrPhotoInfo :=  GetFTPImage(DeleteStr(TpGetLike.GetOutputDataS('sAttachNm', i), '/media/cq/photo/'),
                                             TpGetLike.GetOutputDataS('sHideFile', i),
                                             TpGetLike.GetOutputDataS('sUserId',   i));

            if (tmp_UsrPhotoInfo <> '') then
            begin
               // ������ Image ������ Grid�� ǥ��
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


         {  -- ���� Flat-ListBox String �߰��κ� �ּ� @ 2015.04.09 LSH
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
      apn_LikeList.Caption.Text := IntToStr(iRowCnt) + '���� ��õ�մϴ�';


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
         if (Cells[C_B_BOARDSEQ, ARow] <> '��') and
            (Cells[C_B_BOARDSEQ, ARow] <> '') then

            GetLikeList(asg_Board.Row,
                        asg_Board.Cells[C_B_BOARDSEQ, asg_Board.Row],
                        asg_Board.Cells[C_B_USERIP,   asg_Board.Row],
                        asg_Board.Cells[C_B_REGDATE,  asg_Board.Row])

         else if Cells[C_B_CATEUP, ARow] = '��' then

            GetLikeList(asg_Board.Row,
                        asg_Board.Cells[C_B_CONTEXT, asg_Board.Row],
                        asg_Board.Cells[C_B_USERIP,  asg_Board.Row],
                        asg_Board.Cells[C_B_REGDATE, asg_Board.Row]);

      end;
   end;
end;




//------------------------------------------------------------------------------
// [��ȸ] ���� IP ���� User ���� ���� ��������
//       - DlgUser.pas���� Callback ������ �����
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
   FsMngrNm       := in_MngrNm;   // ���� ������ (����) @ 2014.07.18 LSH
end;





//------------------------------------------------------------------------------
// [ä��] ä�� Client ȣ��
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
                 PChar('[��󿬶���] - [���������]�� ����� IP ����� ����Ͻ� �� �ֽ��ϴ�.'),
                 PChar(Self.Caption + ' : ��ϵ��� ���� User ���� �˸�'),
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


   {  -- Chat-Client���� Log ������Ʈ �ϹǷ� �ּ� @ 2015.03.30 LSH
   // �α� Update
   UpdateLog('CBLOG',
             tmp_ChatUserIp,
             FsUserIp,
             'R',
             '',
             '',
             varResult
             );
   }


   {  -- Chat - Client (BPL) ���뿡 ���� �ּ�ó�� @ 2015.03.30 LSH
   SetClient := TClient.Create(self);


   try
      // ���Ӵ�� IP�� �г��� ����
      SetClient.EditIp.Text   := asg_Master.Cells[C_M_USERIP, asg_Master.Row];
      SetClient.EditNick.Text := FsUserNm;

      // ���� ������ IP �� Port ����
      SetClient.ClientSocket.Host    := SetClient.EditIp.Text;
      SetClient.ClientSocket.Port    := StrToInt(DelChar(CopyByte(asg_Master.Cells[C_M_USERIP, asg_Master.Row], LengthByte(asg_Master.Cells[C_M_USERIP, asg_Master.Row])-3, 4), '.'));


      //---------------------------------------------------------
      // �����Ϸ��� Server Port ��밡�� ���� Check
      //---------------------------------------------------------
      if PortTCPIsOpen(SetClient.ClientSocket.Port, SetClient.ClientSocket.Host) then
      begin
         SetClient.ClientSocket.Active  := True;
         SetClient.ShowModal;
      end
      else
      begin
         MessageBox(self.Handle,
                    PChar('�����Ϸ��� ��� Port�� �����ֽ��ϴ�.' + #13#10 + #13#10 + '���� ������� ���α׷��� ���������� Ȯ���� �ʿ��մϴ�.'),
                    'ê(Chat) �ڽ� :: ���� ���� Port ����',
                    MB_OK + MB_ICONWARNING);

         SetClient.ClientSocket.Active  := False;
      end;



   finally
      SetClient.Free;
   end;
   }


   // �� ���� (�̹���) ��� ����
   if Trim(asg_NetWork.Cells[C_NW_PHOTOFILE, iUserRowId]) = '' then
   begin
      FsChatMyPhoto := asg_NetWork.Cells[C_NW_USERID, iUserRowId] + '.jpg';
   end
   else
      // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
      FsChatMyPhoto  := G_HOMEDIR + 'TEMP\SPOOL\' + asg_Network.Cells[C_NW_PHOTOFILE, iUserRowId];



   //-------------------------------------------------------
   // ���̾� Chat - Client (BPL) Form ȣ��
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
         // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
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

         //showmessage('Client ȣ��Ʈ ��û�� --> ����ȭ�� call');

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
// String ���� Ư�� ���� ����
//       - �ҽ� ��ó : MComFunc.pas
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
// [Chat] Client�� Message Echo ����
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
                 PChar('�޼����� ���� �� �ְ� ����� User�� ���� �����ϴ�.'),
                 'ê(Chat) :: ���� �������� User ����',
                 MB_OK + MB_ICONWARNING);

      Result := False;

      Exit;
   end;

   for i := 0 to (ServerSocket.Socket.ActiveConnections - 1) do
   begin
      //------------------------------------------------------------
      // ���� ����(��ȭ)���� Client���Ը� Send �޼��� ó��
      //------------------------------------------------------------
      if PosByte(ServerSocket.Socket.Connections[i].RemoteAddress, apn_ChatBox.Caption.Text) > 0 then
         ServerSocket.Socket.Connections[i].SendText(AnsiString(s));
   end;

   Result := True;

end;





//------------------------------------------------------------------------------
// [Chat] Server Socket - Client ���� Event
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
   SendToAllClients('��[��ȭ��]�� �ٸ� User�κ��� �߰����� ��ȭ��û�� �� �� ������, �����ϴ� �޼����� ���� [��ȭ��]�� User���Ը� ���۵˴ϴ١�');


   if PosByte('��ȭ��', apn_ChatBox.Caption.Text) = 0 then
      apn_ChatBox.Caption.Text := 'ê(Chat) :: ' + Socket.RemoteHost + '(' + socket.RemoteAddress + ')' + ' ��ȭ��';



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
// [Chat] Server Socket - Client ������� Event
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
         apn_ChatBox.Caption.Text      := 'ê(Chat) :: ' + Socket.RemoteHost + '(' + socket.RemoteAddress + ')' + ' ��ȭ����';
         apn_ChatBox.Collaps           := True;
         apn_ChatBox.Top               := 540;
         apn_ChatBox.Left              := 304;
         apn_ChatBox.Caption.Color     := $000066C9;
         apn_ChatBox.Caption.ColorTo   := clMaroon;
      end;



      { ------------> onClientError �̺�Ʈ���� �������� Off->On ó�� ������ !!
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
// [Chat] Server Socket - Client �޼��� Read Event
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
   // Window API : â�� ��Ȱ��ȭ --> Ȱ��ȭ @ 2014.02.28 LSH
   //     - ������ ����
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
   asg_Chat.RowCount := iChatRowCnt;   // Grid Index out of Range �������� �ݵ�� �ʿ� !


   //-----------------------------------------------------------
   // ���� �˶� Text ���Ÿ��
   //-----------------------------------------------------------
   if (PosByte('just connected',     sReceived) > 0) or
      (PosByte('just disconnected',  sReceived) > 0) or
      (PosByte('[��ȭ��]',           sReceived) > 0) then
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
   // �̸�Ƽ��(BMP �̹���) ���Ÿ��
   //-----------------------------------------------------------
   else if (PosByte('[E]',  sReceived) > 0) then
   begin

      // ���� ���ų��� Temp ���Ͽ� ���� (Client User �̸� ��������)
      tmpOrgReceived := sReceived;

      DataSize := 0;
      sl       := '';


      if not Reciving then
      begin
         // Now we need to get the LengthByte of the data stream.
         SetLength(sl, StrLen(PChar(sReceived)) +1);           // +1 for the null terminator




         StrLCopy(@sl[1], PChar(sReceived), {PosByte('[', PChar(sReceived))-1)}LengthByte(sl)- 13);   // �̸��±�(8) + &(1) + [E] (3) + ���� Null (1)




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
               // 1�� �߰�
               asg_Chat.InsertRows(iChatRowCnt, 1);

            except
               on e : Exception do
               showmessage('Server Read Emoti ERROR(InsertRows) : ' + e.Message);
            end;



            try
               // ������ Data-Stream�� Grid�� ǥ��
               asg_Chat.CreatePicture(C_CH_TEXT, iChatRowCnt, True, ShrinkWithAspectRatio, 0, haRight, vaTop).Bitmap.LoadFromStream(Data);

            except
               asg_Chat.Cells[C_CH_TEXT, iChatRowCnt] := '[�̹��� ���ſ���]';
            end;

            // Data �۽��� ǥ��
            //asg_Chat.Cells[C_CH_YOURCOL,   iChatRowCnt] := CopyByte(tmpOrgReceived, PosByte('&', tmpOrgReceived) + 1, 8);


            // Chat ��� ������ �̹��� ������, ���� ǥ��
            if FileExists(CopyByte(FindStr(tmpOrgReceived, '%'), PosByte('&', FindStr(tmpOrgReceived, '%')) + 2, LengthByte(FindStr(tmpOrgReceived, '%')) - PosByte('&', FindStr(tmpOrgReceived, '%')) - 2)) then
            begin
               // Chat ���κ��� ���� ���� ������ �̹��� ����.
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


            {  -- ���� Text User �� �����ϴ� �κ� �ּ� @ 2015.03.31 LSH
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


            // Scroll ����
            asg_Chat.Row := iChatRowCnt;

            // �۽��� �÷� Align
            asg_Chat.Alignments[C_CH_YOURCOL, iChatRowCnt] := taLeftJustify;

            // �ش� Row Auto-Sizing
            asg_Chat.AutoSizeRow(iChatRowCnt);

            // �޸� ����
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
   // ���� ���Ÿ��
   //-----------------------------------------------------------
   else if (PosByte('[A]',  sReceived) > 0) then
   begin
      // Client�� Msg Echo
      //SendToAllClients(sReceived);


      asg_Chat.InsertRows(iChatRowCnt, 1);

      asg_Chat.AddButton(C_CH_MYCOL,iChatRowCnt, asg_Chat.ColWidths[C_CH_MYCOL],  20, '�ٿ�', haBeforeText, vaCenter);

      asg_Chat.Cells[C_CH_TEXT,     iChatRowCnt] := CopyByte(sReceived, 1, PosByte('$', sReceived) - 1);  //CopyByte(sReceived, PosByte('$', sReceived) + 1, PosByte('[', sReceived) - PosByte('$', sReceived) -1); //

      //asg_Chat.Cells[C_CH_YOURCOL,  iChatRowCnt] := CopyByte(sReceived, PosByte('&', sReceived) + 1, 8);


      // Chat ��� ������ �̹��� ������, ���� ǥ��
      if FileExists(CopyByte(FindStr(sReceived, '%'), PosByte('&', FindStr(sReceived, '%')) + 2, LengthByte(FindStr(sReceived, '%')) - PosByte('&', FindStr(sReceived, '%')) - 2)) then
      begin
         // Chat ���κ��� ���� ���� ������ �̹��� ����.
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
   // �Ϲ� Text ���Ÿ��
   //-----------------------------------------------------------
   else
   begin
      // Client�� Msg Echo
      //SendToAllClients(sReceived);

      asg_Chat.InsertRows(iChatRowCnt, 1);


      //asg_Chat.Cells[C_CH_YOURCOL,  iChatRowCnt] := CopyByte(sReceived, PosByte('&', sReceived) + 1, 10);


      // test
      //showmessage('MainDlg���� ���Ž� : ' + sReceived + #13#10#13#10 + CopyByte(FindStr(sReceived, '%'), PosByte('&', FindStr(sReceived, '%')) + 2, LengthByte(FindStr(sReceived, '%')) - PosByte('&', FindStr(sReceived, '%')) - 2) + #13#10#13#10 + CopyByte(sReceived, PosByte('%', sReceived) + 2, LengthByte(sReceived) - PosByte('%', sReceived) - 2));

      // Chat ��� ������ �̹��� ������, ���� ǥ��
      if FileExists(CopyByte(FindStr(sReceived, '%'), PosByte('&', FindStr(sReceived, '%')) + 2, LengthByte(FindStr(sReceived, '%')) - PosByte('&', FindStr(sReceived, '%')) - 2)) then
      begin
         // Chat ���κ��� ���� ���� ������ �̹��� ����.
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
// [Chat] Server Socket �޼��� ���� Event
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

   // ����� �̹��� ǥ�� @ 2015.03.30 LSH
   if FileExists(FsChatMyPhoto) then
   begin
      try
         // ��ȭ��� Image ������ Grid�� ǥ��
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
      //showmessage('text ���۽� : ' + FsChatMyPhoto);

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
// [Chat] Server Socket �޼��� ���� onButtonClick Event Handler
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
      asg_Chat.RowCount := iChatRowCnt;   // Grid Index out of Range �������� �ݵ�� �ʿ� !


      //----------------------------------------------------------
      // �̸�Ƽ��(Bitmap) ���۸��
      //----------------------------------------------------------
      if (IsEmotiSend) then
      begin

         ms := TMemoryStream.Create;

         try
            // ���� �ҷ��� �̹����� Memory-Stream���� ����
            asg_ChatSend.GetPicture(C_SD_TEXT, 0).Bitmap.SaveToStream(ms);


            if ms.Size > 3000 then
            begin
               //
               MessageBox(self.Handle,
                          PChar('�̸�Ƽ�ܿ� Bitmap �̹����� 3KB ���ϸ� �����մϴ�.'),
                          'ê(Chat) :: �̸�Ƽ��(Bitmap) ���� Size ����',
                          MB_OK + MB_ICONERROR);

               // �޸� ����
               ms.Free;

               // �Է�â �̹��� ����
               asg_ChatSend.RemovePicture(C_SD_TEXT, 0);

               // ���� Chagne
               IsEmotiSend := False;

               Exit;
            end;



            ms.Position := 0;


            // Data Stream ������ ���� �� ���� Flag + ������ �̸�
            // Data Stream ������ ���� + Chat ������ �̹���(Ÿ��Ʋ) ���� @ 2015.03.31 LSH
            if FileExists(FsChatMyPhoto) then
               SendToAllClients(IntToStr(ms.Size) + '[E]' + '&[' + FsChatMyPhoto + ']' + #0)
            else
               SendToAllClients(IntToStr(ms.Size) + '[E]' + '&[' + FsUserNm + ']' + #0);



            // ���� ����� Client�� Data-Stream ����
            ServerSocket.Socket.Connections[ServerSocket.Socket.ActiveConnections-1].SendStream(ms);



            try
               // 1�� �߰�
               asg_Chat.InsertRows(iChatRowCnt, 1);

            except
               on e : Exception do
               showmessage('Server Send ERROR(InsertRows) : ' + e.Message);
            end;




            try
               // ����â Row�� �����Ϸ��� �̹����� ǥ��
               asg_Chat.CreatePicture(C_CH_TEXT, iChatRowCnt, True, ShrinkWithAspectRatio, 0, haleft, vaTop).Bitmap.LoadFromFile(OpenPictureDialog1.FileName);

            except
               asg_Chat.Cells[C_CH_TEXT, iChatRowCnt] := '[�̹��� ���۽���]';
            end;

            // ������ ǥ��
            //asg_Chat.Cells[C_CH_MYCOL, iChatRowCnt]      := '[' + FsUserNm + '] ';
            if FileExists(FsChatMyPhoto) then
            begin
               try
                  // ��ȭ��� Image ������ Grid�� ǥ��
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

            // Scroll ����
            asg_Chat.Row   := iChatRowCnt;

            // �Է�â �̹��� ����
            asg_ChatSend.RemovePicture(C_SD_TEXT, 0);

            // ���� Chagne
            IsEmotiSend := False;


         except
            ms.Free;
         end;

      end
      //----------------------------------------------------------
      // File ���۸��
      //----------------------------------------------------------
      else if (IsFileSend) then
      begin

         sFileName   := '';
         sHideFile   := '';
         sServerIp   := '';


         // ÷������ Upload
         if (Trim(asg_ChatSend.Cells[C_SD_TEXT, 0]) <> '') then
         begin

            sFileName := ExtractFileName(Trim(asg_ChatSend.Cells[C_SD_TEXT, 0]));
            sHideFile := 'CHATAPPEND' + DelChar(CopyByte(asg_Master.Cells[C_M_USERIP, asg_Master.Row], LengthByte(asg_Master.Cells[C_M_USERIP, asg_Master.Row])-3, 4), '.') +
                          FormatDateTime('YYYYMMDDHHNNSS', Now) + '.KDIALFILE';


            if Not FileUpLoad(sHideFile, asg_ChatSend.Cells[C_SD_TEXT, 0], sServerIp) then
            begin
               Showmessage('÷������ UpLoad �� ������ �߻��߽��ϴ�.' + #13#10 + #13#10 +
                           '�ٽ��ѹ� �õ��� �ֽñ� �ٶ��ϴ�.');
               Exit;
            end;
         end;

         // ÷������ + Chat ������ �̹���(Ÿ��Ʋ) ����
         if FileExists(FsChatMyPhoto) then
            SendToAllClients(sFileName + '$' + sHideFile + '[A]' + '&[' + FsChatMyPhoto + ']')
         else
            SendToAllClients(sFileName + '$' + sHideFile + '[A]' + '&[' + FsUserNm + ']');



         asg_Chat.InsertRows(iChatRowCnt, 1);

         //asg_Chat.Cells[C_CH_MYCOL,    iChatRowCnt]      := '[' + FsUserNm + ']';

         // ������ �̹��� ǥ�� @ 2015.03.30 LSH
         if FileExists(FsChatMyPhoto) then
         begin
            try
               // ��ȭ��� Image ������ Grid�� ǥ��
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
         asg_Chat.Cells[C_CH_YOURCOL,  iChatRowCnt]      := '<���۵�>';

         //asg_Chat.AddButton(C_CH_YOURCOL,  iChatRowCnt, asg_Chat.ColWidths[C_CH_YOURCOL]-20,  20, '�ٿ�', haBeforeText, vaCenter);

         asg_Chat.Row                                 := iChatRowCnt;
         asg_Chat.Alignments[C_CH_TEXT, iChatRowCnt]  := taLeftJustify;

         asg_ChatSend.Cells[C_SD_TEXT, 0] := '';

         IsFileSend := False;

      end
      //----------------------------------------------------------
      // �Ϲ� Text ���۸��
      //----------------------------------------------------------
      else
      begin
         // test
         //showmessage('MainDlg���� Text ���۽� : ' + FsChatMyPhoto);

         if FileExists(FsChatMyPhoto) then
            SendToAllClients(asg_ChatSend.Cells[C_SD_TEXT, 0] + '&[' + FsChatMyPhoto + ']')
         else
            SendToAllClients(asg_ChatSend.Cells[C_SD_TEXT, 0] + '&[' + FsUserNm + ']');


         asg_Chat.InsertRows(iChatRowCnt, 1);

         //asg_Chat.Cells[C_CH_MYCOL, iChatRowCnt] := '[' + FsUserNm + '] ';

         // ������ �̹��� ǥ�� @ 2015.03.30 LSH
         if FileExists(FsChatMyPhoto) then
         begin
            try
               // ��ȭ��� Image ������ Grid�� ǥ��
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
            //showmessage('�Ϲ� Text ���۸�� : ' + FsChatMyPhoto);

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
// [Chat] Server Socket �޼��� ���� onKeyPress Event Handler
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

      // ������ȣ Cell���� 'Enter' �Է½�, max-seq �ڵ�ä�� @ 2015.05.13 LSH
      if asg_RegDoc.Col = C_RD_DOCSEQ then
         asg_RegDoc_ClickCell(Sender, asg_RegDoc.Row, asg_RegDoc.Col);

   end;
end;



procedure TMainDlg.pm_MasterPopup(Sender: TObject);
begin

   if (IsLogonUser) then
   begin
      //------------------------------------------------------------
      // �����Ϸ��� Server ��밡�� ���� (���� Session On/Off) Check
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
      // �г��� �Է�/������ [����� ������]������ ����
      //------------------------------------------------------------
      if (apn_Master.Visible) then
         mi_NickNm.Visible := True
      else
         mi_NickNm.Visible := False;

      //------------------------------------------------------------
      // SMS�� [��󿬶���], [���̾�Book]������ ������ (���� Ȯ�뿹��)
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
// [�Լ�] TCP Port ���¿��� Check
//    - 2013.08.30 LSH
//    - �ҽ���ó : Google
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
      // Asynch ���� ������, �������� Off-> On
      ServerSocket.Active := False;
      ServerSocket.Active := True;

      {
      // ErrorCode = 0���� �θ�, '�������'���� ���� �����.. ���� ��ƾ� ��..
      //ErrorCode := 1;
      }

      // �����޼��� ��� ����.
      ErrorCode := 0;


      {
      MessageBox(self.Handle,
                 PChar('ServerSocketError(' + FsUserIP + '): ' + Socket.RemoteHost + '(' + Socket.RemoteAddress + ') �κ��� �߸��� Connection Error �߻�(�����ڵ鰪: ' + IntToStr(Socket.SocketHandle) + ')'),
                 'ê(Chat) :: �߸��� Socket ��� ����',
                 MB_OK + MB_ICONERROR);
      }

   end;


end;






//------------------------------------------------------------------------------
// [����] Form onClose Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2013.09.02
//------------------------------------------------------------------------------
procedure TMainDlg.FormClose(Sender: TObject; var Action: TCloseAction);
var
   i : Integer;
   varResult : String;
begin



   // �α� Update
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

   // Ÿ�̸� ���� �߰� @ 2014.06.03 LSH
   for i:= 0 to ComponentCount - 1 do
   begin
      if Components[i] is TTimer then
         TTimer(Components[i]).Enabled := False;
   end;

   Action := caFree;
end;



//------------------------------------------------------------------------------
// [Chat ÷��] AdvStringGrid onButtonClick Event Handler
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
   // ÷������ Click�� Event ó��
   //---------------------------------------------
   with asg_Chat do
   begin
      if ((ACol = C_CH_MYCOL) and (Cells[C_CH_HIDDEN, ARow] <> '')) then
      begin
         // ���� ��/�ٿ�ε带 ���� ���� ��ȸ
         if not GetBinUploadInfo(sServerIp, sFtpUserID, sFtpPasswd, sFtpHostName, sFtpDir) then
         begin
            MessageDlg('���� ������ ���� ����� ���� ��ȸ��, ������ �߻��߽��ϴ�.', Dialogs.mtError, [Dialogs.mbOK], 0);
            exit;
         end;



         //���� ������ ����Ǿ� �ִ� ���ϸ� ����
         if PosByte('/ftpspool/CHATFILE/',Cells[C_CH_HIDDEN, ARow]) > 0 then
            sRemoteFile := Cells[C_CH_HIDDEN, ARow]
         else
            sRemoteFile := '/ftpspool/CHATFILE/' + Cells[C_CH_HIDDEN, ARow];

         // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
         sLocalFile  := G_HOMEDIR + 'TEMP\SPOOL\' + Cells[C_CH_TEXT, ARow]; //lb_Filename.Caption;



         if (GetBINFTP(sServerIp, sFtpUserID, sFtpPasswd, sRemoteFile, sLocalFile, False)) then
         begin
            //	ShowMessage('���������� ����Ǿ����ϴ�.');
         end;



         Screen.Cursor := crDefault;



         try

            if (PosByte('.exe', sLocalFile) > 0) or
               (PosByte('.zip', sLocalFile) > 0) or
               (PosByte('.rar', sLocalFile) > 0) then
            begin
               MessageBox(self.Handle,
                          PChar('÷������(' + sLocalFile + ') �ٿ�ε尡 �Ϸ�Ǿ����ϴ�.' + #13#10 + #13#10 +
                                '�� �ӽ� �ٿ�ε� ���� --> C:\KUMC(_DEV)\TEMP\SPOOL\'),
                          '÷������ �ٿ�ε� ���� �Ϸ�',
                          MB_OK + MB_ICONINFORMATION);
            end
            else
            begin
               ShellExecute(HANDLE, 'open',
                            PCHAR(Cells[C_CH_TEXT, ARow]),
                            PCHAR(''),
                            // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
                            PCHAR(G_HOMEDIR + 'TEMP\SPOOL\'),
                            SW_SHOWNORMAL);
            end;

         except
            MessageBox(self.Handle,
                       PChar('�ش� ���α׷� �ٿ�ε��� ������ �߻��Ͽ����ϴ�.' + #13#10 + #13#10 +
                             '���α׷� ���� �� �ٽ� ������ �ֽñ� �ٶ��ϴ�.'),
                       '÷������ �ٿ�ε� �� ����',
                       MB_OK + MB_ICONERROR);

            Exit;
         end;
      end;
   end;
end;






//------------------------------------------------------------------------------
// [�˾��޴�] Chat ���� ����
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
      // ���� ���۸�� True
      IsFileSend := True;

      // �ڵ� ����
      asg_ChatSendButtonClick(Sender, C_SD_BUTTON, 0);
   end;

end;



//------------------------------------------------------------------------------
// [��ü����] FlatMemo onKeyUp Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2013.09.03
//------------------------------------------------------------------------------
procedure TMainDlg.fmm_TextKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   // �ۼ�����, [Ctrl + A]�� ��ü �� ���õǵ���.
   if(ssCtrl in Shift) and (Key = 65) then
	   fmm_Text.SelectAll;


   if fsbt_WriteFind.Visible then
   begin
      // �� �˻� ��忡��, [Ctrl + Enter]�� �˻� �ڵ����� @ 2016.11.18 LSH
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
      // ������ ǥ�� ���̱�
      pn_Loading.Left    := 128;
      pn_Loading.Top     := 200;
      pn_Loading.Visible := True;

      pn_Loading.Repaint;

   end
   else if AsStatus = 'OFF' then
   begin
      // ������ ǥ�� ����

      pb_DataLoading.Position := 0;
      pn_Loading.Visible      := False;

   end;
end;



//------------------------------------------------------------------------------
// [Ÿ�̸�] ����� ���� �ڵ� Refresh ���� Timer
//
// Author : Lee, Se-Ha
// Date   : 2013.09.06
//------------------------------------------------------------------------------
procedure TMainDlg.tm_MasterTimer(Sender: TObject);
begin
   // Interval = 60,000 (1��) ���� �ڵ� Refresh
   bbt_NetworkRefreshClick(Sender);
end;




//------------------------------------------------------------------------------
// [�˾��޴�] Chat �̸�Ƽ�� ����
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
      // �̸�Ƽ�� ���۸�� True
      IsEmotiSend := True;

      // �ڵ� ����
      asg_ChatSendButtonClick(Sender, C_SD_BUTTON, 0);
   end;
end;




//------------------------------------------------------------------------------
// [Ÿ�̸�] Timer onTimer Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2013.09.12
//------------------------------------------------------------------------------
procedure TMainDlg.tm_TxInitTimer(Sender: TObject);
begin
   //--------------------------------------------------------------------
   // 90�� ����, ���߰� Tmax ���� Ȯ�� �� ����
   //--------------------------------------------------------------------
   if not (TxInit(TokenStr(String(CONST_ENV_FILENAME), '.', 1) + '_D0.ENV', '01')) then
   begin
      MessageBox(self.Handle,
                 PChar('���� Server���� ������ ������ϴ�.' + #13#10 + #13#10 + '90�ʸ��� �ڵ����� ������ �����Ͽ���, ��ø� ��ٷ� �ֽʽÿ�.'),
                 PChar(Self.Caption + ' : TMAX ���Ӳ��� �˸�'),
                 MB_OK + MB_ICONWARNING);
   end;
end;




procedure TMainDlg.fsbt_MenuClick(Sender: TObject);
begin
   // �� ���� �Ĵ� �޴� �ȳ� �˾� @ 2017.10.26 LSH
   MessageBox(self.Handle,
              PChar(GetDietInfo(FormatDateTime('dd', Date),
                                GetDayofWeek(Date, 'HS'))),
              '���� �� ���ְ� �弼�� :-)',
              MB_OK + MB_ICONINFORMATION);

   try
      // �����ڸ� ��� �����(?)�� ���Ӱ��� @ 2017.10.26 LSH
      if (FsUserIp = '������IP') then
      begin
         // ���� �������� ���ε� Panel�� ������ �м����� ��ȯ @ 2015.04.02 LSH
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
         // ������ IP �ĺ���, �ٹ�ó Assign
         if PosByte('�Ⱦϵ�����', FsUserIp) > 0 then
            fcb_MenuLoc.Text := '�ȾϺ���'
         else if PosByte('���ε�����', FsUserIp) > 0 then
            fcb_MenuLoc.Text := '���κ���'
         else if PosByte('�Ȼ굵����', FsUserIp) > 0 then
            fcb_MenuLoc.Text := '�Ȼ꺴��';

         }

         //asg_MenuBar.AddButton(1, 0, asg_MenuBar.ColWidths[1]-5, 20, '���', haBeforeText, vaCenter);    // ��� Button


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
   // 1. ���� ����(.CSV, .XLS) �� Data Loading
   //--------------------------------------------------------------------
   if opendlg_File.Execute then
   begin
      if opendlg_File.FileName = '' then
        Exit;


      sFileExt  := AnsiUpperCase(ExtractFileExt(opendlg_File.FileName));


      // Excel ������ ���, ���� Procedure (LoadExcelFile) Call
      if (sFileExt = '.XLSX') or (sFileExt = '.XLS') Then
      begin
         if opendlg_File.FileName <> '' then
         begin
            LoadExcelFile(opendlg_File.FileName);
         end;
      end
      else
      // CSV �� �Ʒ� ���� ����.
      begin
         // �غ���..
      end;
   end;
end;



//------------------------------------------------------------------------------
// [����] Excel Load Event Handler
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
   iCol_ReleasUser, // ������ ����� �߰� @ 2014.06.19 LSH
   iRow  : Integer;

   // ������ Lab.���� �׸� �߰� @ 2015.04.02 LSH
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

   // ������ Lab.���� �׸� �߰� @ 2015.04.02 LSH
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
      oXL.WorkBooks.Open(xlsfilename, 0, true);          // �б� �������� �б�

      oWK    := oXL.WorkBooks.Item[1];
      oSheet := oWK.ActiveSheet;                         // ���õ� ��Ʈ ��������

      //XLRows := oXL.ActiveSheet.UsedRange.Rows.Count;    // �� Rows �� ���ϱ�.
      //XLRows := 100;

      XLCols := oXL.ActiveSheet.UsedRange.Columns.Count; // �� Columns �� ���ϱ�.

      // ���� ���������� �ۼ��� Row�� Index�� fetch @ 2017.10.25 LSH
      XLLastUsedRows := oXL.ActiveSheet.Cells.Find('*', SearchOrder:=xlByRows, LookIn:=xlValues, SearchDirection:=xlPrevious).Row;


      if (ftc_Dialog.ActiveTab = AT_BOARD) then
      begin
         {  -- ���� ������ �������� update �ּ�ó�� @ 2015.04.02 LSH
         //--------------------------------------------------------------------
         // 2-1. Excel Data Loading
         //--------------------------------------------------------------------
         asg_Menu.BeginUpdate;



         // �������� ������ Xls ������ �ٸ�.
         if PosByte('�Ⱦ�', fcb_MenuLoc.Text) > 0 then
         begin

            for i := 3 to XLRows do
            begin

               for j := 2 to XLCols do
               begin
                  //---------------------------------------------------------------
                  // Excel sheet�� data�� Grid�� �ѷ��ֱ�
                  //---------------------------------------------------------------


                  case i of
                     3 :
                        asg_Menu.Cells[j - 2, i - 3] := oXL.Cells[i, j];

                     4..8 :
                        begin
                           // ��ħ����
                           asg_Menu.Cells[j - 2,  i - 3] := oXL.Cells[i, j];
                           asg_Menu.Colors[j - 2, i - 3] := $00C5B189;
                        end;


                     10..15 :
                        begin
                           // ���ɳ���
                           asg_Menu.CellS[j - 2,  i - 3] := oXL.Cells[i, j];
                           asg_Menu.Colors[j - 2, i - 3] := $0070CDBD;
                        end;


                     17..21 :
                        begin
                           // ���᳢��
                           asg_Menu.CellS[j - 2,  i - 3] := oXL.Cells[i, j];
                           asg_Menu.Colors[j - 2, i - 3] := $0070AFF7;
                        end;
                  end;
               end;
            end;

            // RowCount ����
            asg_Menu.RowCount := XLRows - 5;

         end
         else if PosByte('����', fcb_MenuLoc.Text) > 0 then
         begin
            for i := 3 to XLRows do
            begin
               for j := 2 to XLCols do
               begin
                  //---------------------------------------------------------------
                  // Excel sheet�� data�� Grid�� �ѷ��ֱ�
                  //---------------------------------------------------------------
                  case i of
                     3 :
                        asg_Menu.CellS[j - 2, i - 3] := oXL.Cells[i, j];

                     4..9 :
                        begin
                           // ��ħ����
                           asg_Menu.CellS[j - 2, i - 3]  := oXL.Cells[i, j];
                           asg_Menu.Colors[j - 2, i - 3] := $00C5B189;
                        end;


                     11..16 :
                        begin
                           // ���ɳ���
                           asg_Menu.CellS[j - 2, i - 4]  := oXL.Cells[i, j];
                           asg_Menu.Colors[j - 2, i - 4] := $0070CDBD;
                        end;


                     18..23 :
                        begin
                           // ���᳢��
                           asg_Menu.CellS[j - 2, i - 5]  := oXL.Cells[i, j];
                           asg_Menu.Colors[j - 2, i - 5] := $0070AFF7;
                        end;
                  end;
               end;
            end;


            // RowCount ����
            asg_Menu.RowCount := XLRows - 11;


         end
         else if PosByte('�Ȼ�', fcb_MenuLoc.Text) > 0 then
         begin
            for i := 3 to XLRows do
            begin
               for j := 2 to XLCols do
               begin
                  //---------------------------------------------------------------
                  // Excel sheet�� data�� Grid�� �ѷ��ֱ�
                  //---------------------------------------------------------------
                  case i of
                     3 :
                        asg_Menu.CellS[j - 2, i - 3] := oXL.Cells[i, j];

                     5..8 :
                        begin
                           // ��ħ����
                           asg_Menu.CellS[j - 2, i - 4]  := oXL.Cells[i, j];
                           asg_Menu.Colors[j - 2, i - 4] := $00C5B189;
                        end;


                     13..16 :
                        begin
                           // ���ɳ���
                           asg_Menu.CellS[j - 2, i - 8]  := oXL.Cells[i, j];
                           asg_Menu.Colors[j - 2, i - 8] := $0070CDBD;
                        end;


                     21..24 :
                        begin
                           // ���᳢��
                           asg_Menu.CellS[j - 2, i - 12]  := oXL.Cells[i, j];
                           asg_Menu.Colors[j - 2, i - 12] := $0070AFF7;
                        end;
                  end;
               end;
            end;

            // RowCount ����
            asg_Menu.RowCount := XLRows - 17;

         end;


         asg_Menu.EndUpdate;

         asg_Menu.Refresh;
         }



         //--------------------------------------------------------------------
         // 1. Data Loading bar ���̱�
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
            // ������ Excel sheet�� Title �׸��� Index ���� ��������
            //---------------------------------------------------------------------
            if i = 1 then
            begin
               for k := 1 to XLCols do
               begin
                  if Trim(oXL.Cells[i, k]) = '�⵵' then
                     iCol_Year      := k
                  else if Trim(oXL.Cells[i, k]) = 'ȸ��' then
                     iCol_Seqno     := k
                  else if Trim(oXL.Cells[i, k]) = '������' then
                     iCol_ShowDate  := k
                  else if Trim(oXL.Cells[i, k]) = '�ο�1' then
                     iCol_1stCount  := k
                  else if Trim(oXL.Cells[i, k]) = '�ݾ�1' then
                     iCol_1stRcpamt := k
                  else if Trim(oXL.Cells[i, k]) = '�ο�2' then
                     iCol_2ndCount  := k
                  else if Trim(oXL.Cells[i, k]) = '�ݾ�2' then
                     iCol_2ndRcpamt := k
                  else if Trim(oXL.Cells[i, k]) = '�ο�3' then
                     iCol_3thCount  := k
                  else if Trim(oXL.Cells[i, k]) = '�ݾ�3' then
                     iCol_3thRcpamt := k
                  else if Trim(oXL.Cells[i, k]) = '�ο�4' then
                     iCol_4thCount  := k
                  else if Trim(oXL.Cells[i, k]) = '�ݾ�4' then
                     iCol_4thRcpamt := k
                  else if Trim(oXL.Cells[i, k]) = '�ο�5' then
                     iCol_5thCount  := k
                  else if Trim(oXL.Cells[i, k]) = '�ݾ�5' then
                     iCol_5thRcpamt := k
                  else if Trim(oXL.Cells[i, k]) = '�ѹ�1' then
                     iCol_No1 := k
                  else if Trim(oXL.Cells[i, k]) = '�ѹ�2' then
                     iCol_No2 := k
                  else if Trim(oXL.Cells[i, k]) = '�ѹ�3' then
                     iCol_No3 := k
                  else if Trim(oXL.Cells[i, k]) = '�ѹ�4' then
                     iCol_No4 := k
                  else if Trim(oXL.Cells[i, k]) = '�ѹ�5' then
                     iCol_No5 := k
                  else if Trim(oXL.Cells[i, k]) = '�ѹ�6' then
                     iCol_No6 := k
                  else if Trim(oXL.Cells[i, k]) = '���ʽ�' then
                     iCol_NoBonus := k;

               end;
            end;



            for j := 1 to XLCols do
            begin
               if i > 1 then
               begin
                  //---------------------------------------------------------------
                  // ������ Excel sheet�� Title�� data�� Grid�� �ѷ��ֱ�
                  //---------------------------------------------------------------
                  begin
                     // ����
                     if (j mod iCol_Year = 0) and (j div iCol_Year = 1) then
                     begin
                        asg_BigData.Cells[C_BD_YEAR, i - 1] := CopyByte(oXL.Cells[i, j], 1, 16);
                     end
                     // ȸ��
                     else if (j mod iCol_Seqno = 0) and (j div iCol_Seqno = 1) then
                        asg_BigData.Cells[C_BD_SEQNO,   i - 1] := oXL.Cells[i, j]
                     // ������
                     else if (j mod iCol_ShowDate = 0) and (j div iCol_ShowDate = 1) then
                        asg_BigData.Cells[C_BD_SHOWDT,i - 1] := ReplaceChar(CopyByte(oXL.Cells[i, j], 1, 10), '.', '-')
                     // �ο�1
                     else if (j mod iCol_1stCount = 0) and (j div iCol_1stCount = 1) then
                        asg_BigData.Cells[C_BD_1STCNT,   i - 1] := oXL.Cells[i, j]
                     // �ݾ�1
                     else if (j mod iCol_1stRcpamt = 0) and (j div iCol_1stRcpamt = 1) then
                        asg_BigData.Cells[C_BD_1STAMT,   i - 1] := oXL.Cells[i, j]
                     // �ο�2
                     else if (j mod iCol_2ndCount = 0) and (j div iCol_2ndCount = 1) then
                        asg_BigData.Cells[C_BD_2NDCNT,   i - 1] := oXL.Cells[i, j]
                     // �ݾ�2
                     else if (j mod iCol_2ndRcpamt = 0) and (j div iCol_2ndRcpamt = 1) then
                        asg_BigData.Cells[C_BD_2NDAMT,   i - 1] := oXL.Cells[i, j]
                     // �ο�3
                     else if (j mod iCol_3thCount = 0) and (j div iCol_3thCount = 1) then
                        asg_BigData.Cells[C_BD_3THCNT,   i - 1] := oXL.Cells[i, j]
                     // �ݾ�3
                     else if (j mod iCol_3thRcpamt = 0) and (j div iCol_3thRcpamt = 1) then
                        asg_BigData.Cells[C_BD_3THAMT,   i - 1] := oXL.Cells[i, j]
                     // �ο�4
                     else if (j mod iCol_4thCount = 0) and (j div iCol_4thCount = 1) then
                        asg_BigData.Cells[C_BD_4THCNT,   i - 1] := oXL.Cells[i, j]
                     // �ݾ�4
                     else if (j mod iCol_4thRcpamt = 0) and (j div iCol_4thRcpamt = 1) then
                        asg_BigData.Cells[C_BD_4THAMT,   i - 1] := oXL.Cells[i, j]
                     // �ο�5
                     else if (j mod iCol_5thCount = 0) and (j div iCol_5thCount = 1) then
                        asg_BigData.Cells[C_BD_5THCNT,   i - 1] := oXL.Cells[i, j]
                     // �ݾ�5
                     else if (j mod iCol_5thRcpamt = 0) and (j div iCol_5thRcpamt = 1) then
                        asg_BigData.Cells[C_BD_5THAMT,   i - 1] := oXL.Cells[i, j]
                     // �ѹ�1
                     else if (j mod iCol_No1 = 0) and (j div iCol_No1 = 1) then
                        asg_BigData.Cells[C_BD_NO1,   i - 1] := oXL.Cells[i, j]
                     // �ѹ�2
                     else if (j mod iCol_No2 = 0) and (j div iCol_No2 = 1) then
                        asg_BigData.Cells[C_BD_NO2,   i - 1] := oXL.Cells[i, j]
                     // �ѹ�3
                     else if (j mod iCol_No3 = 0) and (j div iCol_No3 = 1) then
                        asg_BigData.Cells[C_BD_NO3,   i - 1] := oXL.Cells[i, j]
                     // �ѹ�4
                     else if (j mod iCol_No4 = 0) and (j div iCol_No4 = 1) then
                        asg_BigData.Cells[C_BD_NO4,   i - 1] := oXL.Cells[i, j]
                     // �ѹ�5
                     else if (j mod iCol_No5 = 0) and (j div iCol_No5 = 1) then
                        asg_BigData.Cells[C_BD_NO5,   i - 1] := oXL.Cells[i, j]
                     // �ѹ�6
                     else if (j mod iCol_No6 = 0) and (j div iCol_No6 = 1) then
                        asg_BigData.Cells[C_BD_NO6,   i - 1] := oXL.Cells[i, j]
                     // ���ʽ�
                     else if (j mod iCol_NoBonus = 0) and (j div iCol_NoBonus = 1) then
                        asg_BigData.Cells[C_BD_NOBONUS,   i - 1] := oXL.Cells[i, j];


                     // �ش� Thread�� CPU���� �ϴ� ���� �������� ����ó��
                     Application.ProcessMessages;

                  end;
               end;
            end;

            // Update Loading Bar
            pb_DataLoading.StepIt;

         end;



         // RowCount ����
         asg_BigData.RowCount := i - 1;

         //asg_BigData.EndUpdate;

         asg_BigData.Refresh;


         //--------------------------------------------------------------------
         // 3. Data Loading bar �����
         //--------------------------------------------------------------------
         SetLoadingBar('OFF');


         // Comments
         //lb_RegDoc.Caption := '�� ' + IntToStr(i) + '���� �������ε� ������ D/B�� ������Դϴ�.';




         // �������ε� ��󳻿��� 1���̻� �����, D/B Insert ����.
         if (i > 0) then
            InsBigDataList;

      end
      else if (ftc_Dialog.ActiveTab = AT_DOC) then
      begin
         //--------------------------------------------------------------------
         // 1. Data Loading bar ���̱�
         //--------------------------------------------------------------------
         SetLoadingBar('ON');


         // Maximum value of progress status
         //pb_DataLoading.Max := XLRows;
         // ��ȿ�ϰ� 100�� ������ Upload ���� @ 2017.10.25 LSH
         pb_DataLoading.Max := 100;


         //--------------------------------------------------------------------
         // 2-2. Excel Data Loading
         //--------------------------------------------------------------------
         //asg_Release.BeginUpdate;


//         iRow := 0;


         //for i := 1 to XLRows do
         // Title- Row�� ���� �������� ������ ����, ��ȿ�� Range�� �Ʒ����� Searching & Upload ���� @ 2017.10.25 LSH
         for i := 1 to 1 do
         begin
            //---------------------------------------------------------------------
            // ���� ��������� Excel sheet�� Title �׸��� Index ���� ��������
            //---------------------------------------------------------------------
            if i = 1 then
            begin
               for k := 1 to XLCols do
               begin
                  if Trim(oXL.Cells[i, k]) = '�ۼ��Ͻ�' then
                     iCol_RegDate   := k
                  else if Trim(oXL.Cells[i, k]) = '�з�' then
                     iCol_DutySpec  := k
                  else if Trim(oXL.Cells[i, k]) = '�ý���' then
                     iCol_Context   := k
                  else if Trim(oXL.Cells[i, k]) = 'Ŭ���̾�Ʈ �ҽ�' then
                     iCol_ClienSrc  := k
                  else if Trim(oXL.Cells[i, k]) = 'Ŭ���̾�Ʈ' then
                     iCol_Client    := k
                  else if Trim(oXL.Cells[i, k]) = '���� �ҽ�' then
                     iCol_ServeSrc  := k
                  else if Trim(oXL.Cells[i, k]) = '����' then
                     iCol_Server    := k
                  else if PosByte('���� SR', Trim(oXL.Cells[i, k])) > 0 then
                     iCol_SrReqNo   := k
                  else if Trim(oXL.Cells[i, k]) = '�����' then
                     iCol_DutyUser  := k
                  else if Trim(oXL.Cells[i, k]) = '��û��' then
                     iCol_ReqUser   := k
                  else if Trim(oXL.Cells[i, k]) = '������' then
                     iCol_CofmUser  := k
                  else if Trim(oXL.Cells[i, k]) = '�׽�Ʈ�Ͻ�' then
                     iCol_TestDate  := k
                  else if Trim(oXL.Cells[i, k]) = '������ �Ͻ�' then
                     iCol_ReleasDt  := k
                  else if Trim(oXL.Cells[i, k]) = '������' then
                     iCol_ReleasUser:= k
                  else if PosByte('�۾����', Trim(oXL.Cells[i, k])) > 0 then
                     iCol_Remark    := k;
               end;
            end;
         end;



         // Title-Row�� ������ ���� ��ȿ�� ������ Row-Index - 100 Line ������ ������Ʈ ���� @ 2017.10.25 LSH (���ε� �ӵ����� �̽�)
         for i := XLLastUsedRows - 100 to XLLastUsedRows + 1 do
         begin
            for j := 1 to XLCols do
            begin
               if i > 1 then
               begin
                  //---------------------------------------------------------------
                  // ��������� Excel sheet�� Title�� data�� Grid�� �ѷ��ֱ�
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

                     // �ۼ��Ͻ�
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
                     // �з�
                     else if (j mod iCol_DutySpec = 0) and (j div iCol_DutySpec = 1) then
                        asg_Release.Cells[C_RL_DUTYSPEC,   i - 1] := oXL.Cells[i, j]
                     // �������
                     else if (j mod iCol_Context = 0) and (j div iCol_Context = 1) then
                        asg_Release.Cells[C_RL_CONTEXT,   i - 1] := oXL.Cells[i, j]
                     // ���� S/R
                     else if (j mod iCol_SrReqNo = 0) and (j div iCol_SrReqNo = 1) then
                        asg_Release.Cells[C_RL_SRREQNO,   i - 1] := oXL.Cells[i, j]
                     // ��û��
                     else if (j mod iCol_ReqUser = 0) and (j div iCol_ReqUser = 1) then
                        asg_Release.Cells[C_RL_REQUSER,   i - 1] := oXL.Cells[i, j]
                     // �����
                     else if (j mod iCol_DutyUser = 0) and (j div iCol_DutyUser = 1) then
                        asg_Release.Cells[C_RL_DUTYUSER,   i - 1] := oXL.Cells[i, j]
                     // �������Ͻ�
                     else if (j mod iCol_ReleasDt = 0) and (j div iCol_ReleasDt = 1) then
                        asg_Release.Cells[C_RL_RELEASDT,   i - 1] := CopyByte(oXL.Cells[i, j], 1, 16)
                     // ������
                     else if (j mod iCol_CofmUser = 0) and (j div iCol_CofmUser = 1) then
                        asg_Release.Cells[C_RL_COFMUSER,   i - 1] := oXL.Cells[i, j]
                     // �׽�Ʈ�Ͻ�
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
                     // ������ �����
                     else if (j mod iCol_ReleasUser = 0) and (j div iCol_ReleasUser = 1) then
                        asg_Release.Cells[C_RL_RELEASUSER,   i - 1] := oXL.Cells[i, j]
                     // Remark
                     else if (j mod iCol_Remark = 0) and (j div iCol_Remark = 1) then
                        asg_Release.Cells[C_RL_REMARK,   i - 1] := oXL.Cells[i, j];


                     // �ش� Thread�� CPU���� �ϴ� ���� �������� ����ó��
                     Application.ProcessMessages;

                  end;
               end;
            end;

            // Update Loading Bar
            pb_DataLoading.StepIt;

         end;



         // RowCount ����
         asg_Release.RowCount := i - 2;

         //asg_Release.EndUpdate;

         asg_Release.Refresh;


         //--------------------------------------------------------------------
         // 3. Data Loading bar �����
         //--------------------------------------------------------------------
         SetLoadingBar('OFF');


         // Comments
         lb_RegDoc.Caption := '�� ' + IntToStr(i) + '���� �������ε� ������ D/B�� ������Դϴ�.';


         // �������ε� ��󳻿��� 1���̻� �����, D/B Insert ����.
         if (i > 0) then
            InsReleaseList;

      end;



      // Excel Close
      oXL.WorkBooks.Close;


      // ���� ���ε� �Ϸ�� �ֿ� Timer On @ 2014.07.23 LSH
      tm_DialPop.Enabled := True;
      tm_TxInit.Enabled  := True;


   except
      // ���� ���ε� ����ó���� �ֿ� Timer On @ 2014.07.23 LSH
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
// [�Ĵܵ��] AdvStringGrid onButtonClick Event Handler
//    - ������.
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
   // ������ �޴� Update
   if (ACol = 1) then
   begin


      sType := '7';

      iCnt  := 0;


      // ������ IP �ĺ���, �ٹ�ó Assign
      if PosByte('�Ⱦϵ�����', FsUserIp) > 0 then
         sLocateNm := '�Ⱦ�'
      else if PosByte('���ε�����', FsUserIp) > 0 then
         sLocateNm := '����'
      else if PosByte('�Ȼ굵����', FsUserIp) > 0 then
         sLocateNm := '�Ȼ�';



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
               // �Ĵ� ����
               sDietDay   := CopyByte(Cells[j, 0], 1, 2);

               // �Ĵ� ����
               sDayGubun  := CopyByte(Cells[j, 0], 4, 2);

               // ��ħ
               if (i >= 1) and (i <= 6) then
               begin
                  if (i = 1) then
                     sDiet1 := Cells[j, i]
                  else if Cells[j, i] <> '' then
                     sDiet1 := sDiet1 + ',' + Cells[j, i];
               end;

               // ����
               if (i >= 7) and (i <= 12) then
               begin
                  if (i = 7) then
                     sDiet2 := Cells[j, i]
                  else if Cells[j, i] <> '' then
                     sDiet2 := sDiet2 + ',' + Cells[j, i];
               end;

               // ����
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
            AppendVariant(vLocate  ,   sLocateNm              );  // �Ҽ�
            AppendVariant(vUserID  ,   sDietDay               );  // ����(dd)
            AppendVariant(vUserNm  ,   sDayGubun              );  // ����
            AppendVariant(vDiet1   ,   sDiet1                 );  // ��ħ
            AppendVariant(vDiet2   ,   sDiet2                 );  // ����
            AppendVariant(vDiet3   ,   sDiet3                 );  // ����
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
                       PChar('�ְ� �Ĵ������� ���������� [������Ʈ]�Ǿ����ϴ�.'),
                       '[KUMC ���̾�α�] �ְ� �Ĵ����� ������Ʈ �˸� ',
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
   // 1. ���� ����(.CSV, .XLS) �� Data Loading
   //--------------------------------------------------------------------
   if opendlg_File.Execute then
   begin
      if opendlg_File.FileName = '' then
        Exit;


      sFileExt  := AnsiUpperCase(ExtractFileExt(opendlg_File.FileName));


      // Excel ������ ���, ���� Procedure (LoadExcelFile) Call
      if (sFileExt = '.XLSX') or (sFileExt = '.XLS') Then
      begin
         if opendlg_File.FileName <> '' then
         begin
            LoadExcelFile(opendlg_File.FileName);
         end;
      end
      else
      // CSV �� �Ʒ� ���� ����.
      begin
         // �غ���..
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
   // ������ �޴� Update
   sType := '7';

   iCnt  := 0;


   // ������ IP �ĺ���, �ٹ�ó Assign
   if PosByte('�Ⱦ�', fcb_MenuLoc.Text) > 0 then
      sLocateNm := '�Ⱦ�'
   else if PosByte('����', fcb_MenuLoc.Text) > 0 then
      sLocateNm := '����'
   else if PosByte('�Ȼ�', fcb_MenuLoc.Text) > 0 then
      sLocateNm := '�Ȼ�';



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

         // �������� ������ xls ������ �ٸ���.. -_-;
         // ������ ���� �б�
         if (sLocateNm = '�Ⱦ�') then
         begin
            for i := 0 to RowCount - 1 do
            begin
               // �Ĵ� ����
               sDietDay   := CopyByte(Cells[j, 0], 1, 2);

               // �Ĵ� ����
               sDayGubun  := CopyByte(Cells[j, 0], 4, 2);

               // ��ħ
               if (i >= 1) and (i <= 6) then
               begin
                  if (i = 1) then
                     sDiet1 := Cells[j, i]
                  else if Cells[j, i] <> '' then
                     sDiet1 := sDiet1 + ',' + Cells[j, i];
               end;

               // ����
               if (i >= 7) and (i <= 12) then
               begin
                  if (i = 7) then
                     sDiet2 := Cells[j, i]
                  else if Cells[j, i] <> '' then
                     sDiet2 := sDiet2 + ',' + Cells[j, i];
               end;

               // ����
               if (i >= 13) then
               begin
                  if (i = 13) then
                     sDiet3 := Cells[j, i]
                  else if Cells[j, i] <> '' then
                     sDiet3 := sDiet3 + ',' + Cells[j, i];
               end;
            end;
         end
         else if (sLocateNm = '����') then
         begin
            for i := 0 to RowCount - 1 do
            begin
               // �Ĵ� ����
               sDietDay   := CopyByte(Cells[j, 0], 1, 2);

               // �Ĵ� ����
               sDayGubun  := CopyByte(Cells[j, 0], 4, 2);

               // ��ħ
               if (i >= 1) and (i <= 6) then
               begin
                  if (i = 1) then
                     sDiet1 := Cells[j, i]
                  else if Cells[j, i] <> '' then
                     sDiet1 := sDiet1 + ',' + Cells[j, i];
               end;

               // ����
               if (i >= 7) and (i <= 12) then
               begin
                  if (i = 7) then
                     sDiet2 := Cells[j, i]
                  else if Cells[j, i] <> '' then
                     sDiet2 := sDiet2 + ',' + Cells[j, i];
               end;

               // ����
               if (i >= 13) then
               begin
                  if (i = 13) then
                     sDiet3 := Cells[j, i]
                  else if Cells[j, i] <> '' then
                     sDiet3 := sDiet3 + ',' + Cells[j, i];
               end;
            end;
         end
         else if (sLocateNm = '�Ȼ�') then
         begin
            for i := 0 to RowCount - 1 do
            begin
               // �Ĵ� ����
               sDietDay   := CopyByte(Cells[j, 0], 1, 2);

               // �Ĵ� ����
               sDayGubun  := CopyByte(Cells[j, 0], 4, 2);

               // ��ħ
               if (i >= 1) and (i <= 4) then
               begin
                  if (i = 1) then
                     sDiet1 := Cells[j, i]
                  else if Cells[j, i] <> '' then
                     sDiet1 := sDiet1 + ',' + Cells[j, i];
               end;

               // ����
               if (i >= 5) and (i <= 8) then
               begin
                  if (i = 5) then
                     sDiet2 := Cells[j, i]
                  else if Cells[j, i] <> '' then
                     sDiet2 := sDiet2 + ',' + Cells[j, i];
               end;

               // ����
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
         AppendVariant(vLocate  ,   sLocateNm              );  // �Ҽ�
         AppendVariant(vUserID  ,   sDietDay               );  // ����(dd)
         AppendVariant(vUserNm  ,   sDayGubun              );  // ����
         AppendVariant(vDiet1   ,   sDiet1                 );  // ��ħ
         AppendVariant(vDiet2   ,   sDiet2                 );  // ����
         AppendVariant(vDiet3   ,   sDiet3                 );  // ����
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
                    PChar('[' + sLocateNm + '] �ְ� �Ĵ������� ���������� [������Ʈ]�Ǿ����ϴ�.'),
                    '[KUMC ���̾�α�] �ְ� �Ĵ����� ������Ʈ �˸� ',
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


      // �α� Update
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
   // �ѱ��Է� �ʱ�ȭ ���� @ 2015.06.03 LSH
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

   lb_DialScan.Caption := '�� ���̾�Book �˻� �Է´����....';

   // �ֱٰ˻� Key ��ȸ
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
                 PChar('[��󿬶���] - [���������]�� ����� IP ����� ����Ͻ� �� �ֽ��ϴ�.'),
                 PChar(Self.Caption + ' : ��ϵ��� ���� User ���� �˸�'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;





   with asg_DialList do
   begin

      if (Cells[C_DL_LOCATE, Row] = '') then
      begin
         MessageBox(self.Handle,
                    PChar('����Ͻ� ����ó ������ ���õ��� �ʾҽ��ϴ�.'),
                    PChar('���� ���̾� Book ��� ���� �˸�'),
                    MB_OK + MB_ICONWARNING);

         Exit;
      end;

   {
      if CopyByte(Cells[C_DL_LINKSEQ, Row], 1, 4) = '(02)' then
         tmpLinkSeq := CopyByte(Cells[C_DL_LINKSEQ, Row], 5, 9)
      else
         tmpLinkSeq := Trim(Cells[C_DL_LINKSEQ, Row]);
   }




      { --  ���� D/B �����, data�� ǰ���� ���Ͻ�ų �� �ִ� ������ ����.
        --  �� ������ �����Ǳ� �������� ���� D/B (��ȭ��ȣ��/���̾�α�)�� ��ϵǵ��� ��.
      if (Cells[C_DL_LINKDB, Row] = '��ȭ��ȣ��') or
         (Cells[C_DL_LINKDB, Row] = '���̾�α�') then
         tmpUserNm := FsUserNm
      else
         tmpUserNm := Trim(Cells[C_DL_LINKDB, Row]);
      }


      if (apn_LinkList.Visible) then
         // ���� ��ũ�� List���� ���� ���̾� Book�� ������ ����ó ���
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
                      Cells[C_DL_DTYUSRID,   Row]        // ��ϵ� UserId @ 2015.04.13 LSH
                      )


      else
         // ���� �˻� List���� ���� ���̾� Book�� ������ ����ó ���
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
                      Cells[C_DL_DTYUSRID,   Row]        // ��ϵ� UserId @ 2015.04.13 LSH
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
                 PChar('[��󿬶���] - [���������]�� ����� IP ����� ����Ͻ� �� �ֽ��ϴ�.'),
                 PChar(Self.Caption + ' : ��ϵ��� ���� User ���� �˸�'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;




   with asg_MyDial do
   begin

      if (Cells[C_MD_LOCATE, Row] = '') then
      begin
         MessageBox(self.Handle,
                    PChar('�����Ͻ� ����ó ������ ���õ��� �ʾҽ��ϴ�.'),
                    PChar('���� ���̾� Book ���� ���� �˸�'),
                    MB_OK + MB_ICONWARNING);

         Exit;
      end;

      // ���� ���̾� Book�� ������ ����ó ����
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
// [������Ʈ] ���� ���̾� Book Update
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

   // User �̸� filetering @ 2015.04.15 LSH
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
   AppendVariant(vDtyUsrId,  in_DtyUsrId);   // ��ϵ� UserId @ 2015.04.13 LSH



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
               //lb_MyDial.Caption := '����ó [���] �Ϸ�'


               MessageBox(self.Handle,
                          PChar('�����Ͻ� ����ó�� ���� ���̾� Book�� [���] �Ǿ����ϴ�.'),
                          '[KUMC ���̾�α�] ���̾� Book ������Ʈ �˸� ',
                          MB_OK + MB_ICONINFORMATION)

            else
               //lb_MyDial.Caption := '����ó [����] �Ϸ�';

               MessageBox(self.Handle,
                          PChar('�����Ͻ� ����ó�� ���� ���̾� Book���� [����] �Ǿ����ϴ�.'),
                          '[KUMC ���̾�α�] ���̾� Book ������Ʈ �˸� ',
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
   // D/B ��ó�� ���� Update ��ư ���� (�� ID�� D/B ���� ����ʴ���, ū �ǹ� ����)
   if (asg_MyDial.Cells[C_MD_LINKDB, asg_MyDial.Row] = '��ȭ��ȣ��') or
      (asg_MyDial.Cells[C_MD_LINKDB, asg_MyDial.Row] = '���̾�α�') or
      (asg_MyDial.Cells[C_MD_LINKDB, asg_MyDial.Row] = 'Ȩ������') or
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
                       PChar('����Ͻ� ����ó ������ ���õ��� �ʾҽ��ϴ�.'),
                       PChar('���� ���̾� Book ��� ���� �˸�'),
                       MB_OK + MB_ICONWARNING);

            Exit;
         end;



         // ���� ���̾� Book�� ������ ����ó ������� ���
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

   // �Ҽ�, �μ���, �μ��󼼴� Edit ����
   CanEdit := (ACol <> C_MD_LOCATE) and
              (ACol <> C_MD_DEPTNM) and
              (ACol <> C_MD_DEPTSPEC);

{
   if (ARow = asg_MyDial.Row) then
   begin

      // �Ҽ�, �μ���, �μ��󼼴� Edit ����
      CanEdit := (ACol <> C_MD_LOCATE) and
                 (ACol <> C_MD_DEPTNM) and
                 (ACol <> C_MD_DEPTSPEC);

      // D/B ��ó�� ���� ���� ���� (���, �� ID�� D/B ���� �������� �ʴ��� ū �ǹ̴� ����)
      if (asg_MyDial.Cells[C_MD_LINKDB, asg_MyDial.Row] = '��ȭ��ȣ��') or
         (asg_MyDial.Cells[C_MD_LINKDB, asg_MyDial.Row] = '���̾�α�') or
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
// [��ȸ] ���̾� Book / ����(S/R) �м� ��ũ List ����ȸ
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


   // �˻� ����
   if (in_Gubun = 'DIALBOOK') then
   begin
      sType := '19';

      if (PosByte('Book', asg_DialList.Cells[C_DL_LINKDB, asg_DialList.Row]) > 0) and
         (asg_DialList.Cells[C_DL_LOCATE, asg_DialList.Row] <> '���¾�ü') then
         sLinkDb := '��ȭ��ȣ��'
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
      // ���̾� Book : ��ũ�� ����Ʈ ��ȸ
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
               Cells[C_LL_DTYUSRID, i + FixedRows] := TpGetLink.GetOutputDataS('sUserId',     i);  // ��ϵ� UserId @ 2015.04.13 LSH
            end;

            RowCount := RowCount - 1;

         end;



         // Comments
         apn_LinkList.Caption.Text := IntToStr(iRowCnt) + '���� �ش� ����ó�� �Ʒ��� ���� [���� ���̾� Book]�� ���/����� �Դϴ�.';


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
      // ����(S/R) �м� : ��ũ�� �󼼳��� ��ȸ
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
                                       TpGetLink.GetOutputDataS('sRegDate', 0) + ' ��û "' +
                                       TpGetLink.GetOutputDataS('sDocRmk', 0) + '" �󼼳���';


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
      // S/R ����ó �˻��� S/R �󼼳��� ��ȸ
      if PosByte('S/R', Cells[C_DL_LINKDB,   Row]) > 0 then
      begin
         // ���񽺴ܿ��� LINKSEQ�� CQSRREQT.REQNO�� ������, �ڵ� ��ȸ �����ϳ�,
         // �ֱ� 1���̳� S/R ��û���� ��� ��ȸ�� �Ǿ, ��ȿ�� data�� �ߺ����� ��ȸ�Ǵ� �κк��ٴ�
         // ���� unique��(distinct) ����ó�� �����ִ� ���� ���ڴٴ� �Ǵ�.
         GetLinkList('ANALYSIS',
                     Trim(Cells[C_DL_LINKSEQ, Row]),
                     '',
                     '',
                     FsUserIp);
      end
      // ��Ÿ D/B�� ���, ���� Link ����� �󼼳��� ��ȸ
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
   // ������ �ֱٰ˻� Keyword assign
   fed_Scan.Text := asg_DialScan.Cells[C_DS_KEYWORD, asg_DialScan.Row];


   if apn_LinkList.Visible then
   begin
      apn_LinkList.Collaps := True;
      apn_LinkList.Visible := False;
   end;


   // Refresh
   SelGridInfo('DIALLIST');


   // �α� Update
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
   // ������ �˻�â Click�ÿ��� �� Login User �̸� Default Set. @ 2016.12.08 LSH
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
   // S/R ���� �˻�(D/B)
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


         // �α� Update
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

         asg_WeeklyRpt.Cells[C_WK_GUBUN, 0] := '����';
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

         // �α� Update
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
   // ����� ��˻� (Grid)
   else if (fed_Analysis.Text <> '') and
           (fcb_ReScan.Checked)      and
           (Key = #13) then
   begin

      if (asg_Analysis.Visible) then
      begin
         asg_Analysis.Options := [goFixedVertLine,goFixedHorzLine,goVertLine,goHorzLine,goColSizing];

         FndParam := [];

         FndParam := FndParam + {[fnMatchCase] +} [fnDirectionLeftRight];     // Character InSensitive �˻����� fnMatchCase �ּ� @ 2016.11.18 LSH

         ResPoint := asg_Analysis.FindFirst(Trim(fed_Analysis.Text), FndParam);

         if ResPoint.X >= 0 then
         begin
            asg_Analysis.Col   := ResPoint.X;
            asg_Analysis.Row   := ResPoint.Y;
         end
         else
            MessageBox(self.Handle,
                       PChar('ã���ô� �׸��� ȭ�� List�� �������� �ʽ��ϴ�.'),
                       PChar('S/R �˻� : ����� ��˻� ��� �˸�'),
                       MB_OK + MB_ICONWARNING)
                       ;

         asg_Analysis.SetFocus;
      end
      else if (apn_Weekly.Visible) then
      begin
         //asg_WorkRpt.Options := [goFixedVertLine,goFixedHorzLine,goVertLine,goHorzLine,goColSizing];

         FndParam := [];

         FndParam := FndParam + {[fnMatchCase] +} [fnDirectionLeftRight];     // Character InSensitive �˻����� fnMatchCase �ּ� @ 2016.11.18 LSH

         ResPoint := asg_WeeklyRpt.FindFirst(Trim(fed_Analysis.Text), FndParam);

         if ResPoint.X >= 0 then
         begin
            asg_WeeklyRpt.Col   := ResPoint.X;
            asg_WeeklyRpt.Row   := ResPoint.Y;
         end
         else
            MessageBox(self.Handle,
                       PChar('ã���ô� �׸��� ȭ�� List�� �������� �ʽ��ϴ�.'),
                       PChar('S/R �˻� : ����� ��˻� ��� �˸�'),
                       MB_OK + MB_ICONWARNING)
                       ;

         asg_WeeklyRpt.SetFocus;
      end
      else if (asg_WorkRpt.Visible) then
      begin
         FndParam := [];

         FndParam := FndParam + {[fnMatchCase] +} [fnDirectionLeftRight];     // Character InSensitive �˻����� fnMatchCase �ּ� @ 2016.11.18 LSH

         ResPoint := asg_WorkRpt.FindFirst(Trim(fed_Analysis.Text), FndParam);

         if ResPoint.X >= 0 then
         begin
            asg_WorkRpt.Col   := ResPoint.X;
            asg_WorkRpt.Row   := ResPoint.Y;
         end
         else
            MessageBox(self.Handle,
                       PChar('ã���ô� �׸��� ȭ�� List�� �������� �ʽ��ϴ�.'),
                       PChar('S/R �˻� : ����� ��˻� ��� �˸�'),
                       MB_OK + MB_ICONWARNING)
                       ;

         asg_WorkRpt.SetFocus;
      end
      else if (apn_DBMaster.Visible) then
      begin
         asg_DBMaster.Options := [goFixedVertLine,goFixedHorzLine,goVertLine,goHorzLine,goColSizing];

         FndParam := [];

         FndParam := FndParam + {[fnMatchCase] +} [fnDirectionLeftRight];     // Character InSensitive �˻����� fnMatchCase �ּ� @ 2016.11.18 LSH

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

            FndParam := FndParam + {[fnMatchCase] +} [fnDirectionLeftRight];  // Character InSensitive �˻����� fnMatchCase �ּ� @ 2016.11.18 LSH

            ResPoint := asg_DBDetail.FindFirst(Trim(fed_Analysis.Text), FndParam);

            if ResPoint.X >= 0 then
            begin
               asg_DBDetail.Col   := ResPoint.X;
               asg_DBDetail.Row   := ResPoint.Y;

               asg_DBDetail.SetFocus;
            end
            else
               MessageBox(self.Handle,
                          PChar('ã���ô� �׸��� ȭ�� List�� �������� �ʽ��ϴ�.'),
                          PChar('D/B �˻� : ����� ��˻� ��� �˸�'),
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

   // S/R �󼼳��� Grid ��Ŀ�� @ 2016.11.18 LSH
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
      // ÷�� File �� �ƴ� �κи� Event ó��
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
         // S/R ÷�� ���ϸ�
         ServerFileName := Trim(Cells[Col, Row]);


         // Local �����̸� Set
         // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
         ClientFileName := G_HOMEDIR + 'TEMP\SPOOL\' + ServerFileName;



         // Local�� �ش� ���� �������� üũ
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
            FTP_USERID := '';
            FTP_PASSWD := '';
            FTP_DIR    := '/kumis/app/mis/media/cq/';



            // S/R ÷������ �ٿ�ε�
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



            // �ش� ���� Variant�� ����.
            AppendVariant(DownFile, ClientFileName);



            // �ٿ�ε� Ƚ�� üũ
            DownCnt := DownCnt + 1;

         end;

         try
            if (PosByte('.exe', ClientFileName) > 0) or
               (PosByte('.zip', ClientFileName) > 0) or
               (PosByte('.rar', ClientFileName) > 0) then
            begin
               MessageBox(self.Handle,
                          PChar('÷������(' + ClientFileName + ') �ٿ�ε尡 �Ϸ�Ǿ����ϴ�.' + #13#10 + #13#10 +
                                '�� �ӽ� �ٿ�ε� ���� --> C:\KUMC(_DEV)\TEMP\SPOOL\'),
                          'S/R ÷������ �ٿ�ε� ���� �Ϸ�',
                          MB_OK + MB_ICONINFORMATION);
            end
            else
            begin
               ShellExecute(HANDLE, 'open',
                            PCHAR(Trim(Cells[Col, Row])),
                            PCHAR(''),
                            // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
                            PCHAR(G_HOMEDIR + 'TEMP\SPOOL\'),
                            SW_SHOWNORMAL);
            end;

         except
            MessageBox(self.Handle,
                       PChar('�ش� ÷������ ������ ������ �߻��Ͽ����ϴ�.' + #13#10 + #13#10 +
                             '�ٽ� ������ �ֽñ� �ٶ��ϴ�.'),
                       'S/R ÷������ �ٿ�ε� ���� ����',
                       MB_OK + MB_ICONERROR);

            Exit;
         end;
      end;
   end;

   // S/R �����м� ��ȸ Panel - on��, �ش� Grid ��Ŀ�� ���� @ 2016.11.18 LSH
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

      //[Ctrl + C] : ������ Cell Ŭ������� ����
      Ord('C') : if (ssCtrl in Shift) then
                 begin
                    asg_AnalList.CopySelectionToClipboard;
                 end;

      //[Ctrl + A] : �ش� Grid Cell ��ü���� Ŭ������� ����
      Ord('A') : if (ssCtrl in Shift) then
                 begin
                    asg_AnalList.CopyToClipBoard;
                 end;

      //[ESC] : �ش� S/R �󼼳��� Panel-Off @ 2016.11.18 LSH
      Ord(VK_ESCAPE) :
                     begin
                        asg_AnalListDblClick(Sender);
                     end;
      //[Enter] : ÷�� File ����
      Ord(VK_RETURN) :
                     begin
                        // ÷�� File Row-Index [Enter] �Է½� ���� @ 2016.11.18 LSH
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
      // ���� Cell ��ǥ ��������
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
   // �������� �ۼ�(TMemo ���� ������Ʈ) Panel Close @ 2014.07.31 LSH
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

   fed_Analysis.Text       := '';         // �������� �˻���� �������� null �ʱ�ȭ @ 2016.09.02 LSH (fed_Analysis.Text       := FsUserNm;)
   fed_Analysis.Hint       := '������������ �˻��Ϸ��� Ű����(��: ����, ����, �ڷ�����..)�� ���� Enter�� ����������.';

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

   fed_DateTitle.Text      := '�Ⱓ';
   label8.Left             := 124;
   fmed_AnalFrom.Left      := 57;
   fmed_AnalTo.Left        := 135;

   lb_Analysis.Caption := '�� ��������(KUMC) �ٹ�ó��/�������к�/�Ⱓ�� ��ȸ ����.';

   fmed_AnalFrom.Text   := FormatDateTime('yyyy-mm-dd', DATE - 1);
   fmed_AnalTo.Text     := FormatDateTime('yyyy-mm-dd', DATE);

   fcb_WorkGubun.Top     := 999;
   fcb_WorkGubun.Left    := 999;
   fcb_WorkGubun.Visible := False;

   // ������ IP �ĺ���, �ٹ�ó Assign
   if PosByte('�Ⱦϵ�����', FsUserIp) > 0 then
      fcb_WorkArea.Text := '�Ⱦ�'
   else if PosByte('���ε�����', FsUserIp) > 0 then
      fcb_WorkArea.Text := '����'
   else if PosByte('�Ȼ굵����', FsUserIp) > 0 then
      fcb_WorkArea.Text := '�Ȼ�';

   fsbt_Print.Visible := True;
   fsbt_Print.ColorFocused := $0087D5DB;


   // �������� ���� ��ȸ
   if IsLogonUser then
      SelGridInfo('WORKRPT');

end;

procedure TMainDlg.fsbt_AnalysisClick(Sender: TObject);
var
   vKey : Char;
begin
   // �������� �ۼ�(TMemo ���� ������Ʈ) Panel Close @ 2014.07.31 LSH
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
   fed_Analysis.Hint       := 'S/R ����, ��û�� ���Ե� ����(�ڵ�, ó��� ��), ��û��, ȯ�ڹ�ȣ, ÷�����ϸ� �� �ƹ��ų� �˻��غ�����.';


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

   fed_DateTitle.Text      := '�Ⱓ';
   label8.Left             := 124;
   fmed_AnalFrom.Left      := 57;
   fmed_AnalTo.Left        := 135;

   lb_Analysis.Caption := '�� S/R ��û���� �м��� ���� ȿ������ �˻��� �����մϴ�.';

   fmed_AnalFrom.Text   := FormatDateTime('yyyy-mm-dd', DATE - 730);   // �� 3��ġ �˻� --> 2��ġ �˻� ��ȯ @ 2016.11.18 LSH
   fmed_AnalTo.Text     := FormatDateTime('yyyy-mm-dd', DATE);

   fcb_WorkGubun.Top     := 999;
   fcb_WorkGubun.Left    := 999;
   fcb_WorkGubun.Visible := False;

   fsbt_Print.Visible   := False;

   // S/R �������� �ڵ���ȸ
   vKey := #13;
   fed_Analysis.Text := '��û������';
   fed_AnalysisKeyPress(Sender, vKey);


   // S/R �̷� ��ȸ Grid ��Ŀ�� @ 2016.11.18 LSH
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
      {  -- S/R �Ⱓ����� �ǽð� ��ȸ ���� @ 2016.11.18 LSH
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
      {  -- ������ ���� �Ⱓ����� �������� �ǽð� ��ȸ ���� @ 2016.11.18 LSH
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
      {  -- S/R �Ⱓ����� �ǽð� ��ȸ ���� @ 2016.11.18 LSH
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

      //[Ctrl + C] : ������ Cell Ŭ������� ����
      Ord('C') : if (ssCtrl in Shift) then
                 begin
                    asg_WorkRpt.CopySelectionToClipboard;
                 end;

      //[Ctrl + A] : �ش� Grid Cell ��ü���� Ŭ������� ����
      Ord('A') : if (ssCtrl in Shift) then
                 begin
                    asg_WorkRpt.CopyToClipBoard;
                 end;

      //[F3 Key] : ���� ã��
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
                                   PChar('ã���ô� �׸��� ȭ�� List�� �������� �ʽ��ϴ�.'),
                                   PChar('�������� : ����� ��˻� ��� �˸�'),
                                   MB_OK + MB_ICONWARNING)
                                   ;
                  end;

      //[Ctrl + F] : �˻� EditBox ��Ŀ�� @ 2016.11.18 LSH
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
         // D/B �˻� ī�װ� ������, [����� ��˻�] üũ �ڵ����� @ 2016.11.18 LSH
         if fcb_ReScan.Checked then
            fcb_ReScan.Checked := False;


         SelGridInfo('DBCATEGORY');

         // D/B �˻� ī�װ� ������, [����� ��˻�] üũ �ڵ����� @ 2016.11.18 LSH
         fcb_ReScan.Checked := True;

         // �˻� EditBox ��Ŀ�� @ 2016.11.18 LSH
         fed_Analysis.SetFocus;

      end;
end;

procedure TMainDlg.fsbt_WeeklyClick(Sender: TObject);
begin
   // �������� �ۼ�(TMemo ���� ������Ʈ) Panel Close @ 2014.07.31 LSH
   SetPanelStatus('WORKRPT', 'OFF');


   //------------------------------------------------------------------
   // ���̾�Book / S/R�˻� ������ : S/R �ֱ� 1�Ⱓ History ��ȸ
   //------------------------------------------------------------------
   if (ftc_Dialog.ActiveTab = AT_DIALBOOK) then
   begin
      //
      fed_Analysis.Text    := asg_UserProfile.Cells[1, R_PR_USERID];
      apn_Weekly.Top       := 54;
      apn_Weekly.Left      := 13;

      apn_UserProfile.Collaps := True;
      apn_UserProfile.Visible := False;

      lb_DialScan.Caption   := '�� ' + Trim(asg_UserProfile.Cells[1, R_PR_USERNM]) + '���� �ֱ� 1�Ⱓ S/R ��û������ ��ȸ�Ͽ����ϴ�. [����Ŭ����, ���̾�Book ����]';

      apn_Weekly.Collaps := True;
      apn_Weekly.Visible := True;
      asg_WeeklyRpt.ClearRows(1, asg_WeeklyRpt.RowCount);
      asg_WeeklyRpt.RowCount := 1;
      apn_Weekly.Collaps := False;

      asg_WeeklyRpt.Cells[C_WK_GUBUN, 0] := 'S/R���';


      // ��������Ʈ ��ȸ
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

      lb_Analysis.Caption   := '�� ' + Trim(asg_UserProfile.Cells[1, R_PR_USERNM]) + '���� �ֱ� 1�Ⱓ S/R ��û������ ��ȸ�Ͽ����ϴ�. [����Ŭ����, S/R�˻� ����]';

      apn_Weekly.Collaps := True;
      apn_Weekly.Visible := True;
      asg_WeeklyRpt.ClearRows(1, asg_WeeklyRpt.RowCount);
      asg_WeeklyRpt.RowCount := 1;
      apn_Weekly.Collaps := False;

      asg_WeeklyRpt.Cells[C_WK_GUBUN, 0] := 'S/R���';
      //fcb_ReScan.Visible := False;


      // ��������Ʈ ��ȸ
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
      fed_Analysis.Hint       := '�������� �ۼ��Ϸ��� ����� �̸�(��: �̼���) + �˻� Ű����(��: MDO110F1)�� �Է��� ��(��: �̼��� MDO110F1) Enter�� ����������.';

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


      fed_DateTitle.Text      := '�Ⱓ';
      label8.Left             := 124;
      fmed_AnalFrom.Left      := 57;
      fmed_AnalTo.Left        := 135;

      lb_Analysis.Caption := '�� S/R ������ ��������(KUMC) �м��� ���� ����Ʈ�� �ۼ��մϴ�.';

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


      // ������ IP �ĺ���, �ٹ�ó Assign
      if (ftc_Dialog.ActiveTab = AT_DIALOG) then
      begin
         if PosByte('�Ⱦ�', asg_UserProfile.Cells[1, R_PR_DUTYPART]) > 0 then
         begin
            fcb_WorkArea.Text := '�Ⱦ�';
         end
         else if PosByte('����', asg_UserProfile.Cells[1, R_PR_DUTYPART]) > 0 then
         begin
            fcb_WorkArea.Text := '����';
         end
         else if PosByte('�Ȼ�', asg_UserProfile.Cells[1, R_PR_DUTYPART]) > 0 then
         begin
            fcb_WorkArea.Text := '�Ȼ�';
         end;

         fed_Analysis.Text    := Trim(asg_UserProfile.Cells[1, R_PR_USERNM]);

         apn_Weekly.Top       := 55;
         apn_Weekly.Left      := 15;

         lb_Network.Caption   := '�� ' + Trim(asg_UserProfile.Cells[1, R_PR_USERNM]) + '���� �ֱ� ������������ ��ȸ�Ͽ����ϴ�. [����Ŭ����, ��󿬶������� ����]'

      end
      else if (ftc_Dialog.ActiveTab = AT_ANALYSIS) then
      begin
         if PosByte('�Ⱦϵ�����', FsUserIp) > 0 then
         begin
            fcb_WorkArea.Text := '�Ⱦ�';
         end
         else if PosByte('���ε�����', FsUserIp) > 0 then
         begin
            fcb_WorkArea.Text := '����';
         end
         else if PosByte('�Ȼ굵����', FsUserIp) > 0 then
         begin
            fcb_WorkArea.Text := '�Ȼ�';
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

      asg_WeeklyRpt.Cells[C_WK_GUBUN, 0] := '����';

      fsbt_Print.Visible := True;
      fsbt_Print.ColorFocused := $008B8FE3;

      // ��������Ʈ ��ȸ
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

      //[Ctrl + C] : ������ Cell Ŭ������� ����
      Ord('C') : if (ssCtrl in Shift) then
                 begin
                    asg_WeeklyRpt.CopySelectionToClipboard;
                 end;

      //[Ctrl + A] : �ش� Grid Cell ��ü���� Ŭ������� ����
      Ord('A') : if (ssCtrl in Shift) then
                 begin
                    asg_WeeklyRpt.CopyToClipBoard;
                 end;

      //[Ctrl + F] : �˻� EditBox ��Ŀ�� @ 2016.12.08 LSH
      Ord('F') :
                  if (ssCtrl in Shift) then
                  begin
                     fed_Analysis.SetFocus;
                  end;

      //[F3 Key] : ���� ã��
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
                                   PChar('ã���ô� �׸��� ȭ�� List�� �������� �ʽ��ϴ�.'),
                                   PChar('����Ʈ�ۼ� : ����� ��˻� ��� �˸�'),
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

      //[Ctrl + C] : ������ Cell Ŭ������� ����
      Ord('C') : if (ssCtrl in Shift) then
                 begin
                    asg_SelDoc.CopySelectionToClipboard;
                 end;

      //[Ctrl + A] : �ش� Grid Cell ��ü���� Ŭ������� ����
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


   // Report Name ����
   RecName := '���������� ��������';



   // Title�� �ۼ����� + ���� ǥ��
   if (fmed_AnalTo.Text > fmed_AnalFrom.Text) then
      RecName := RecName + ' [' + fmed_AnalFrom.Text + ', '  + GetDayofWeek(StrToDate(fmed_AnalFrom.Text), 'HS') + ' ~ ' + fmed_AnalTo.Text + ', ' + GetDayofWeek(StrToDate(fmed_AnalTo.Text), 'HS') +']'
   else
      RecName := RecName + ' [' + fmed_AnalFrom.Text + ', '  + GetDayofWeek(StrToDate(fmed_AnalFrom.Text), 'HS') + ']';



   // �Ƿ���� Display
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
      //ts := ts+ ((tw - textwidth('������б��Ƿ��')) shr 1);
      ts := ts +  10;

      Textout(ts, -15, '������б��Ƿ��');

      font.assign(savefont);

      savefont.free;
   end;


   // ������� Display
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
                 PChar('[��󿬶���] - [���������]�� ����� IP ����� ����Ͻ� �� �ֽ��ϴ�.'),
                 PChar(Self.Caption + ' : ��ϵ��� ���� User ���� �˸�'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;


   //-------------------------------------------------------------------
   // 2. Set Print Options
   //-------------------------------------------------------------------
   if (asg_WorkRpt.Visible) then
   begin
      // ��ȸ �������� Temp ������ �ӽú���
      tmp_AnalTo := fmed_AnalTo.Text;

      // ��ȸ���� ���ϸ� ��µǵ��� ���� @ 2014.07.17 LSH
      fmed_AnalTo.Text := fmed_AnalFrom.Text;


      // AdvStringGrid.PrintSettings����
      // AutoSizeRows ���� üũ �����ؾ�, �Ʒ� Hide Column�� ������ ��� ������.
      asg_WorkRpt.InsertRows(asg_WorkRpt.RowCount, 4);
      asg_WorkRpt.Cells[0, asg_WorkRpt.RowCount - 4] := '�� ��';
      asg_WorkRpt.Cells[0, asg_WorkRpt.RowCount - 3] := '����';
      asg_WorkRpt.Cells[1, asg_WorkRpt.RowCount - 3] := '�� ��';
      asg_WorkRpt.Cells[1, asg_WorkRpt.RowCount - 2] := '��Ʈ��';
      asg_WorkRpt.Cells[1, asg_WorkRpt.RowCount - 1] := '�� ��';

      // ��� �׸� Merge Cell ����
      asg_WorkRpt.MergeCells(0, asg_WorkRpt.RowCount - 4, 3, 1);
      asg_WorkRpt.MergeCells(3, asg_WorkRpt.RowCount - 4, 2, 1);

      // ���� �׸� Merge Cell ����
      asg_WorkRpt.MergeCells(0, asg_WorkRpt.RowCount - 3, 1, 3);

      // ���/��Ʈ��/���� �׸� Merge Cell ����
      asg_WorkRpt.MergeCells(1, asg_WorkRpt.RowCount - 3, 2, 1);
      asg_WorkRpt.MergeCells(3, asg_WorkRpt.RowCount - 3, 2, 1);
      asg_WorkRpt.MergeCells(1, asg_WorkRpt.RowCount - 2, 2, 1);
      asg_WorkRpt.MergeCells(3, asg_WorkRpt.RowCount - 2, 2, 1);
      asg_WorkRpt.MergeCells(1, asg_WorkRpt.RowCount - 1, 2, 1);
      asg_WorkRpt.MergeCells(3, asg_WorkRpt.RowCount - 1, 2, 1);

      // Ȯ�ζ� Visible
      asg_WorkRpt.ColWidths[C_WR_CONFIRM]   := 30;

      // �ۼ����ڴ� Invisible
      asg_WorkRpt.HideColumn(C_WR_REGDATE);




      SetPrintOptions;




      {   -- �̸����� Text ��
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
         // Grid Index Out of Range ���� �߻��Ǵ� �κ�.. ��¹� �̻����,
         // AdvStringGrid ���׷� �����Ǵµ�, ���� �м��� �� �ȵ�.
      end;




      // �α� Update
      UpdateLog('WORKRPT',
                'PRINT',
                FsUserIp,
                'P',
                FsUserSpec,
                FsUserNm,
                varResult
                );



      // Comments
      lb_Analysis.Caption := '�� �������� ����� �Ϸ�Ǿ����ϴ�.';

      // ��ȸ������ ���󺹱�
      fmed_AnalTo.Text := tmp_AnalTo;

   end
   else if (apn_Weekly.Visible) then
   begin

      SetPrintOptions;




      {   -- �̸����� Text ��
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







      // �α� Update
      UpdateLog('WEEKRPT',
                'PRINT',
                FsUserIp,
                'P',
                FsUserSpec,
                FsUserNm,
                varResult
                );



      // Comments
      lb_Analysis.Caption := '�� ��������Ʈ ����� �Ϸ�Ǿ����ϴ�.';

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


   // Report Name ����
   if (Trim(fed_Analysis.Text) <> '') then
      RecName := Trim(fed_Analysis.Text) + '�� ��������Ʈ'
   else
      RecName := '��������Ʈ';



   // Title�� �ۼ����� + ���� ǥ��
   if (fmed_AnalTo.Text > fmed_AnalFrom.Text) then
      RecName := RecName + ' [' + fmed_AnalFrom.Text + ', '  + GetDayofWeek(StrToDate(fmed_AnalFrom.Text), 'HS') + ' ~ ' + fmed_AnalTo.Text + ', ' + GetDayofWeek(StrToDate(fmed_AnalTo.Text), 'HS') +']'
   else
      RecName := RecName + ' [' + fmed_AnalFrom.Text + ', '  + GetDayofWeek(StrToDate(fmed_AnalFrom.Text), 'HS') + ']';



   // �Ƿ���� Display
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
      //ts := ts+ ((tw - textwidth('������б��Ƿ��')) shr 1);
      ts := ts +  10;

      Textout(ts, -15, '������б��Ƿ�� ����������');

      font.assign(savefont);

      savefont.free;
   end;


   // ������� Display
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

   // ���� �ٹ�ó ���̾� MAP ��ȸ
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

      //[F3 Key] : ���� ã��
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
                                   PChar('ã���ô� �׸��� ȭ�� List�� �������� �ʽ��ϴ�.'),
                                   PChar('S/R �˻� : ����� ��˻� ��� �˸�'),
                                   MB_OK + MB_ICONWARNING)
                                   ;
                  end;

      //[Ctrl + F] : �˻� EditBox ��Ŀ�� @ 2016.11.18 LSH
      Ord('F') :
                  if (ssCtrl in Shift) then
                  begin
                     fcb_ReScan.Checked := True;

                     fed_Analysis.SetFocus;
                  end;


      //[Enter] : �ش� S/R �󼼳��� Panel-On @ 2016.11.18 LSH
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
      // ���� Cell ��ǥ ��������
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
                 PChar('[��󿬶���] - [���������]�� ����� IP ����� ����Ͻ� �� �ֽ��ϴ�.'),
                 PChar(Self.Caption + ' : ��ϵ��� ���� User ���� �˸�'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;

   with asg_AddMyDial do
   begin
      ClearCols(1, ColCount);
      RowCount  := 8;

      Cells[1, 0] := '���¾�ü';
      Cells[1, 1] := '';
      Cells[1, 2] := '';
      Cells[1, 3] := '';
      Cells[1, 4] := '';
      Cells[1, 5] := '';
      Cells[1, 6] := '';
      Cells[1, 7] := '�ʼ��׸� �Է��� �� ĭ�� ��Ϲ�ư�� ������.';

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
   // ���� �Է½�, ���� Cell ���࿩�� �б�
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
                   '���� ���̾� Book �űԵ��',
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

   // 1-1. �ʼ��Է� �׸� Ȯ��
   with asg_AddMyDial do
   begin
      if (Trim(Cells[Col, 0]) = '') or
         (Trim(Cells[Col, 1]) = '') or
         (Trim(Cells[Col, 2]) = '') or
         (Trim(Cells[Col, 3]) = '')
      then
      begin
         MessageBox(Self.Handle,
                    PChar('[����ó] ������ �ʼ��Է� �׸��Դϴ�.'),
                    '���� ���̾� Book �ű� ����� �ʼ��Է� Ȯ��',
                    MB_OK + MB_ICONWARNING);

         Exit;
      end;
   end;




   // 1-2. ���� ��� ���� Ȯ��
   iChoice := MessageBox(Self.Handle,
                         PChar('�ۼ��Ͻ� �ű� ���̾��� ����Ͻðڽ��ϱ�?'),
                         '���� ���̾� Book �ű� ��� Ȯ��',
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
                    '���� ���̾� Book�� ���������� [���]�Ǿ����ϴ�.',
                    '���� ���̾� Book ��ϿϷ� �˸�',
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
   // 4. ���� ���̾� Book List ��ȸ
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
      // �� Mouse ������ ��ǥ �޾ƿ���
      MouseToCell(X,
                  Y,
                  NowCol,
                  NowRow);


      // ����� ������ ��ȸ
      if NowCol = C_NW_USERNM then
      begin

         if (NowRow > 0) and
            (IsLogonUser) then
         begin
            lb_Network.Caption := '�� ' + Trim(asg_UserProfile.Cells[1, R_PR_USERNM]) + '���� ����� ������ ��ȸ��.';

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

            // ������ ���� assign
            asg_UserProfile.Cells[1, R_PR_USERNM]     := Cells[C_NW_USERNM,   NowRow];
            asg_UserProfile.Cells[1, R_PR_DUTYPART]   := Cells[C_NW_LOCATE,   NowRow]  + ' ' +
                                                         Cells[C_NW_DUTYPART, NowRow]  + ' ' +
                                                         Cells[C_NW_DUTYSPEC, NowRow];
            asg_UserProfile.Cells[1, R_PR_REMARK]     := Cells[C_NW_DUTYRMK,  NowRow];
            asg_UserProfile.Cells[1, R_PR_CALLNO]     := Cells[C_NW_CALLNO,   NowRow];

            asg_UserProfile.AddButton(1, R_PR_DUTYSCH, asg_UserProfile.ColWidths[0]+45, 20, '�ֱ� ���� History ����', haBeforeText, vaCenter);


            // ������ �̹��� FTP ���ϳ����� �����ϸ�, �̹��� Download
            if (Cells[C_NW_PHOTOFILE, NowRow] <> '') then
            begin
               // MIS �⺻ �̹��� ������ ���.
               if Cells[C_NW_HIDEFILE, NowRow] = '' then
               begin

                  // ����� ������ �̹��� ���ϸ�
                  ServerFileName := Cells[C_NW_PHOTOFILE, NowRow];


                  // Local �����̸� Set
                  // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
                  ClientFileName := G_HOMEDIR + 'TEMP\SPOOL\' + ServerFileName;



                  // Local�� �ش� Image �������� üũ
                  // Local file�� Size üũ �߰� (Local�� �ش� ���.jpg�� 0 size�� �����ϸ� JPEG Error #41 ������) @ 2015.06.03 LSH
                  if (FileExists(ClientFileName)) and
                     (CMsg.GetFileSize(ClientFileName) >  0) then
                  begin
                     // FTP �������� ��������
                     if Not GetBinUploadInfo(FTP_SVRIP,
                                             FTP_USERID,
                                             FTP_PASSWD,
                                             FTP_HOSTNAME,
                                             FTP_DIR) then
                     begin
                        //ShowMessage('�ٿ�ε尡 �����Ͽ� ������ �� �����ϴ�.');
                        TUXFTP := nil;
                        Exit;
                     end;


                     // FTP ���� IP Set
                     FTP_SVRIP        := C_MIS_FTP_IP;


                     // FTP �������� Set
                     FTP_USERID := '';
                     FTP_PASSWD := '';
                     FTP_DIR    := '/kumis/app/mis/media/cq/photo/';



                     // Image �ٿ�ε�
                     if Not GetBINFTP(FTP_SVRIP,
                                      FTP_USERID,
                                      FTP_PASSWD,
                                      FTP_DIR + ServerFileName,
                                      ClientFileName,
                                      False) then
                     begin

                        //Showmessage('�̹��� �ٿ�ε忡 �����߽��ϴ�. �ٽ��ѹ� �õ��� �ֽʽÿ�');

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
                  if PosByte('/ftpspool/KDIALFILE/', Cells[C_NW_HIDEFILE, NowRow]) > 0 then
                     sRemoteFile := Cells[C_NW_HIDEFILE, NowRow]
                  else
                     sRemoteFile := '/ftpspool/KDIALFILE/' + Cells[C_NW_HIDEFILE, NowRow];

                  // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
                  sLocalFile  := G_HOMEDIR + 'TEMP\SPOOL\' + Cells[C_NW_PHOTOFILE, NowRow];



                  if (GetBINFTP(FTP_SVRIP, FTP_USERID, FTP_PASSWD, sRemoteFile, sLocalFile, False)) then
                  begin
                     //	�������� FTP �ٿ�ε�
                  end;

                  if (FileExists(sLocalFile)) and
                     (CMsg.GetFileSize(sLocalFile) > 0) then
                  begin
                     // �̹����� Preview ����
                     if (PosByte('.bmp', sLocalFile) > 0) or
                        (PosByte('.BMP', sLocalFile) > 0) or
                        (PosByte('.jpg', sLocalFile) > 0) or
                        (PosByte('.JPG', sLocalFile) > 0) or
                        {  -- TMS Grid���� png Ÿ�� ���ν� ������ �ּ� @ 2017.11.01 LSH
                        (PosByte('.png', sLocalFile) > 0) or
                        (PosByte('.PNG', sLocalFile) > 0) or
                        }
                        (PosByte('.gif', sLocalFile) > 0) or
                        (PosByte('.GIF', sLocalFile) > 0) then
                        // ������ Image ������ Grid�� ǥ��
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
               // ����� ������ �̹��� ���ϸ�
               if Cells[C_NW_USERID, NowRow] = 'XXXXX' then
               begin
                  // ������ �̹��� Remove
                  asg_UserProfile.RemovePicture(1, R_PR_PHOTO);

                  // �̹��� ��� ��ư ����
                  asg_UserProfile.AddButton(1, R_PR_PHOTO, asg_UserProfile.ColWidths[0]+45, asg_UserProfile.RowHeights[R_PR_PHOTO]-5, '������ �̹��� ���', haBeforeText, vaCenter);

                  Exit;
               end
               else
                  // ����� HRM ������ �̹��� ���ϸ�
                  ServerFileName := Cells[C_NW_USERID, NowRow] + '.jpg';

                  // Local �����̸� Set
                  // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
                  ClientFileName := G_HOMEDIR + 'TEMP\SPOOL\' + ServerFileName;


                  // Local�� �ش� Image �������� üũ
                  // Local file�� Size üũ �߰� (Local�� �ش� ���.jpg�� 0 size�� �����ϸ� JPEG Error #41 ������) @ 2015.06.03 LSH
                  if (FileExists(ClientFileName)) and
                     (CMsg.GetFileSize(ClientFileName) >  0) then
                  begin
                     // FTP �������� ��������
                     if Not GetBinUploadInfo(FTP_SVRIP,
                                             FTP_USERID,
                                             FTP_PASSWD,
                                             FTP_HOSTNAME,
                                             FTP_DIR) then
                     begin
                        //ShowMessage('�ٿ�ε尡 �����Ͽ� ������ �� �����ϴ�.');
                        TUXFTP := nil;
                        Exit;
                     end;


                     // FTP ���� IP Set
                     FTP_SVRIP        := C_MIS_FTP_IP;

                     // FTP �������� Set
                     FTP_USERID := '';
                     FTP_PASSWD := '';
                     FTP_DIR    := '/kuhrm/app/hrm/photo/';


                     // HRM Image �ٿ�ε�
                     if Not GetBINFTP(FTP_SVRIP,
                                      FTP_USERID,
                                      FTP_PASSWD,
                                      FTP_DIR + ServerFileName,
                                      ClientFileName,
                                      False) then
                     begin
                        //Showmessage('�̹��� �ٿ�ε忡 �����߽��ϴ�. �ٽ��ѹ� �õ��� �ֽʽÿ�');

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
         lb_Network.Caption := '�� [���������]�� [����������]�� ������Ʈ �Ͻø� ��󿬶����� �ڵ���ȸ �˴ϴ�.';

         apn_UserProfile.Collaps := True;
         apn_UserProfile.Visible := False;
      end;
   end;

end;




//----------------------------------------------------------------
// String ���� Ư�� ���ڿ� ����
//    - �ҽ���ó : MComFunc.pas
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
   if (PosByte('���', apn_DocSpec.Caption.Text) > 0) or
      (PosByte('����', apn_DocSpec.Caption.Text) > 0) then
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
                      asg_RegDoc.Cells[C_RD_DOCLIST, asg_RegDoc.Row] + ' ���� ���',
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
      // CRA/CRC ID ����
      if (PosByte('CRA', Cells[C_RD_DOCLIST, Row]) > 0) or
         (PosByte('CRC', Cells[C_RD_DOCLIST, Row]) > 0) then
      begin
         if (asg_DocSpec.Cells[1, R_DS_DEPTCD] = 'CTC') then
            Cells[C_RD_DOCTITLE, Row] := 'CRA OCS ������ ��û'
         else
            Cells[C_RD_DOCTITLE, Row] := 'CRC OCS ������ ��û';

         Cells[C_RD_REGDATE, Row] := FormatDateTime('yyyy-mm-dd', Date);
         Cells[C_RD_REGUSER, Row] := asg_DocSpec.Cells[1, R_DS_USERNM];
         Cells[C_RD_RELDEPT, Row] := asg_DocSpec.Cells[1, R_DS_DEPTNM];
         Cells[C_RD_DOCRMK,  Row] := asg_DocSpec.Cells[1, R_DS_USERID];

      end
      // �������� ���ó ����
      else if (PosByte('���', Cells[C_RD_DOCLIST, Row]) > 0) then
      begin
         Cells[C_RD_DOCTITLE, Row] := asg_DocSpec.Cells[1, R_DS_USERID];   // ����
         Cells[C_RD_REGDATE,  Row] := FormatDateTime('yyyy-mm-dd', Date);  // �����
         Cells[C_RD_REGUSER,  Row] := FsUserNm;                            // �����
         Cells[C_RD_RELDEPT,  Row] := asg_DocSpec.Cells[1, R_DS_USERNM];   // ��ü��
         Cells[C_RD_DOCRMK,   Row] := asg_DocSpec.Cells[1, R_DS_DEPTCD];   // ���Ⱓ
         Cells[C_RD_DOCLOC,   Row] := asg_DocSpec.Cells[1, R_DS_USELVL];   // ������ ������ġ(ex. PDF, ���ó..)
      end;
   end;

   apn_DocSpec.Collaps := True;
   apn_DocSpec.Visible := False;

end;


//-----------------------------------------------------
// User�� ������ �˶� Check
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
   // 2. ��ȸ
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



      // ���� Flag �� Return Value �б�
      if (in_Gubun = 'BORN') then      // ���� ����
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

            fpn_Title.Caption := '���� �����մϴ� ��';

            fpn_Photo.Visible := True;

            fani_Photo.BringToFront;

            for i := 0 to iRowCnt - 1 do
            begin
               asg_Congrats.Cells[0, i + asg_Congrats.FixedRows] := TpGetAlarm.GetOutputDataS('sLocate',  i) +
                                                                    TpGetAlarm.GetOutputDataS('sDutyPart',i) + ' ' +
                                                                    TpGetAlarm.GetOutputDataS('sUserNm',  i) + '��';

               asg_Congrats.AddButton(1, i + asg_Congrats.FixedRows, asg_Congrats.ColWidths[1] - 5, 20, '�޼��� ������', haBeforeText, vaCenter);

               asg_Congrats.Cells[2, i + asg_Congrats.FixedRows] := TpGetAlarm.GetOutputDataS('sMobile',  i);
            end;
         end;

         asg_Congrats.RowCount := asg_Congrats.RowCount - 1;

      end
      // ���Ϻ� ��/�� ����
      else if (in_Gubun = 'Diet1Off') or
              (in_Gubun = 'Diet2Off') or
              (in_Gubun = 'Diet3Off') then
      begin
         // �ش� ���̳����� �˶�Off-Log�� ������(Y), True ����.
         if (TpGetAlarm.GetOutputDataS('sLogYn', 0) = 'Y') then
            Result := True
         else
            Result := False;
      end
      // ���ΰ���
      else if (in_Gubun = 'NOTI') then
      begin
         // �α�Ȯ�� �Ǵ� ��ȿ�� Noti. ���� ������, Exit.
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


            // ���� Title ������ ǥ��, ������ null
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


            // Local �����̸� Set
            // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
            ClientFileName := G_HOMEDIR + 'TEMP\SPOOL\' + TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 9);


            // Local�� �ش� Image �������� üũ
            // Local file�� Size üũ �߰� (Local�� �ش� ���.jpg�� 0 size�� �����ϸ� JPEG Error #41 ������) @ 2015.06.03 LSH
            if (FileExists(ClientFileName)) and
               (CMsg.GetFileSize(ClientFileName) >  0) then
            begin
               // FTP �������� ��������
               if Not GetBinUploadInfo(FTP_SVRIP,
                                       FTP_USERID,
                                       FTP_PASSWD,
                                       FTP_HOSTNAME,
                                       FTP_DIR) then
               begin
                  ShowMessage('FTP �������� ��ȸ�� �����Ͽ� ������ �� �����ϴ�.');
                  TUXFTP := nil;
                  Exit;
               end;


               // FTP �������� Set
               FTP_USERID := '';
               FTP_PASSWD := '';
               FTP_DIR    := '/ftpspool/KDIALIMG/';


               // Image �ٿ�ε�
               if Not GetBINFTP(C_KDIAL_FTP_IP,
                                FTP_USERID,
                                FTP_PASSWD,
                                FTP_DIR + TokenStr(TpGetAlarm.GetOutputDataS('sContext',  0), '|', 9),
                                ClientFileName,
                                False) then
               begin
                  Showmessage('���� �̹��� �ٿ�ε忡 �����߽��ϴ�. �ٽ��ѹ� �õ��� �ֽʽÿ�');

                  TUXFTP := nil;

                  Exit;
               end;
            end;

            try
               // D/B���� ������ �� FTP �������� �ٿ�ε� �� ǥ��
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
      // �α��ν� Noti. �α׾�����Ʈ @ 2015.04.08 LSH
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
   lb_RegDoc.Caption := '�� �μ����빮��(�߼�/���� ��)�� �˻� �Ǵ� ���(����)�Ͻ� �� �ֽ��ϴ�.';

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

   // ��������� ���ε�
   fsbt_Upload.Visible        := False;
   fsbt_Insert.Visible        := False;

   if IsLogonUser then
      SelGridInfo('DOC');

   tm_Master.Enabled := False;

end;

procedure TMainDlg.fsbt_ReleaseClick(Sender: TObject);
begin
   // Comments
   lb_RegDoc.Caption := '�� ���߰��ù���(���������)�� �˻� �Ǵ� ���(����)�Ͻ� �� �ֽ��ϴ�.';

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

   // ��������� ���ε� �޴�
   fsbt_Upload.Visible        := True;
   fsbt_Upload.ColorFocused   := $00AFC3BD;

   // Ư�� ������ ��� : Ư���Ⱓ ��������� ������ ���ε� ���
   if (PosByte('������IP', FsUserIp) > 0) then
   begin
      fsbt_Insert.ColorFocused   := $00AFC3BD;
      fsbt_Insert.Visible        := True;
   end;

   // Default
   fcb_DutySpec.Text := '����';

   fmed_RegFrDt.Text := FormatDateTime('yyyy-mm-dd', Date - 3);
   fmed_RegToDt.Text := FormatDateTime('yyyy-mm-dd', Date    );   // Date - 1���� Date (������ ����)������ ���� (H �����, ������ ���� ������ �����ۼ�...) @ 2016.11.16 LSH

   fed_Release.Clear;


   if IsLogonUser then
   begin
      // �����ڴ� �Ҽ� ������Ʈ�� ������ ��Ȳ Scan
      if PosByte('����', FsUserPart) > 0 then
         SelGridInfo('RELEASE')
      else
         SelGridInfo('RELEASESCAN');
   end;

   tm_Master.Enabled := False;

   // ���� [��������] Click�ÿ��� ������ ���� ��ȸ���� Tagging ó�� @ 2015.06.15 LSH
   fsbt_Release.Tag := 1;

   // �˻� �Ϸ��� Grid ��Ŀ�� @ 2016.11.18 LSH
   asg_Release.SetFocus;

end;

procedure TMainDlg.fsbt_UploadClick(Sender: TObject);
var
   sFileExt : String;
begin
   //--------------------------------------------------------------------
   // 1. ���� ����(.CSV, .XLS) �� Data Loading
   //--------------------------------------------------------------------
   if opendlg_File.Execute then
   begin
      if opendlg_File.FileName = '' then
        Exit;


      sFileExt  := AnsiUpperCase(ExtractFileExt(opendlg_File.FileName));


      // Excel ������ ���, ���� Procedure (LoadExcelFile) Call
      if (sFileExt = '.XLSX') or (sFileExt = '.XLS') Then
      begin
         if opendlg_File.FileName <> '' then
         begin
            // ���� ���ε�� �ֿ� Timer Off (�ӵ�����) @ 2014.07.23 LSH
            tm_DialPop.Enabled := False;
            tm_TxInit.Enabled  := False;

            LoadExcelFile(opendlg_File.FileName);
         end;
      end
      else
      // CSV �� �Ʒ� ���� ����.
      begin
         // �ʿ��ϸ� �غ�.
      end;
   end;
end;


//--------------------------------------------------------------------------
// ������ Lab. ���� ���ε� �ϰ����
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
   vReleasUser,        // ������ ����� �߰� @ 2014.06.19 LSH
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
   // 1. ������ Raw - data ���
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
            AppendVariant(vDutyPart ,   Cells[C_BD_YEAR, i]    );                   // ����
            AppendVariant(vDutySpec ,   Cells[C_BD_SEQNO ,i]   );                   // ȸ��
            AppendVariant(vRegDate  ,   CopyByte(Trim(Cells[C_BD_SHOWDT  ,i]), 1, 10)); // ������
            AppendVariant(vCofmUser ,   Cells[C_BD_1STCNT,i]   );                   // �ο�1
            AppendVariant(vContext  ,   Cells[C_BD_1STAMT,i]   );                   // �ݾ�1
            AppendVariant(vReleasDt ,   Cells[C_BD_2NDCNT,i]   );                   // �ο�2
            AppendVariant(vClient	,   Cells[C_BD_2NDAMT,i]   );                   // �ݾ�2
            AppendVariant(vTestDate ,   Cells[C_BD_3THCNT,i]   );                   // �ο�3
            AppendVariant(vServer	,   Cells[C_BD_3THAMT,i]   );                   // �ݾ�3
            AppendVariant(vSrreqNo  ,   Cells[C_BD_4THCNT,i]   );                   // �ο�4
            AppendVariant(vClienSrc ,   Cells[C_BD_4THAMT,i]   );                   // �ݾ�4
            AppendVariant(vRemark   ,   Cells[C_BD_5THCNT,i]   );                   // �ο�5
            AppendVariant(vServeSrc ,   Cells[C_BD_5THAMT,i]   );                   // �ݾ�5
            AppendVariant(vReqDept  ,   Cells[C_BD_NO1,i]      );                   // �ѹ�1
            AppendVariant(vReleasUser,  Cells[C_BD_NO2,i]      );                   // �ѹ�2
            AppendVariant(vCateUp   ,   Cells[C_BD_NO3,i]      );                   // �ѹ�3
            AppendVariant(vCateDown ,   Cells[C_BD_NO4,i]      );                   // �ѹ�4
            AppendVariant(vAttachNm ,   Cells[C_BD_NO5,i]      );                   // �ѹ�5
            AppendVariant(vHideFile ,   Cells[C_BD_NO6,i]      );                   // �ѹ�6
            AppendVariant(vServerIp ,   Cells[C_BD_NOBONUS,i]  );                   // �ѹ�(���ʽ�)
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
                    PChar(IntToStr(iCnt) +'���� �����Ͱ� ���������� [������Ʈ] �Ǿ����ϴ�.'),
                    '[KUMC ���̾�α�] ������ Lab. ������Ʈ �˸� ',
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
// ���߰��� ����(���������) ���� ���ε� �ϰ����
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
   vReleasUser,        // ������ ����� �߰� @ 2014.06.19 LSH
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
                    PChar('[�ٹ�ó/��������/�����/����ó]�� �ʼ��Է� �׸��Դϴ�.'),
                    PChar(Self.Caption + ' : ���������� �ʼ��Է� �׸� �˸�'),
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
                    PChar(IntToStr(iCnt) +'���� ��������� ���������� [������Ʈ] �Ǿ����ϴ�.'),
                    '[KUMC ���̾�α�] ���߹���(���������) ������Ʈ �˸� ',
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
   if Application.MessageBox('������ �Ⱓ�� ��������� [����] �Ͻðڽ��ϱ�?',
                             PChar('������� ���ñⰣ�� ���'), MB_OKCANCEL) <> IDOK then
      Exit;

   // ���ñⰣ�� D/B ���
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
   // ������ ���� �˻�(D/B)
   if (fed_Release.Text <> '') and
      (not fcb_ReScanRelease.Checked)  and
      (Key = #13) then
   begin
      // Comments
      lb_RegDoc.Caption := '�� "' + Trim(fed_Release.Text) + '"�� ������ ������ �˻��� �Դϴ�....';


      // Refresh
      if IsLogonUser then
         SelGridInfo('RELEASESCAN');


      // �α� Update
      UpdateLog('RELEASE',
                'KEYWORD',
                FsUserIp,
                'F',
                Trim(fed_Release.Text),
                '',
                varResult
                );
   end
   // ����� ��˻� (Grid)
   else if (fed_Release.Text <> '') and
           (fcb_ReScanRelease.Checked)      and
           (Key = #13) then
   begin
      FndParam := [];

      FndParam := FndParam + {[fnMatchCase] +} [fnDirectionLeftRight];  // Character InSensitive �˻����� fnMatchCase �ּ� @ 2016.11.18 LSH

      ResPoint := asg_Release.FindFirst(Trim(fed_Release.Text), FndParam);

      if ResPoint.X >= 0 then
      begin
         asg_Release.Col   := ResPoint.X;
         asg_Release.Row   := ResPoint.Y;
      end
      else
         MessageBox(self.Handle,
                    PChar('ã���ô� �׸��� ȭ�� List�� �������� �ʽ��ϴ�.'),
                    PChar('������� ����� ��˻� ��� �˸�'),
                    MB_OK + MB_ICONWARNING)
                    ;

      asg_Release.SetFocus;
   end;
end;

procedure TMainDlg.fcb_DutySpecChange(Sender: TObject);
begin
   // Comments
   lb_RegDoc.Caption := '�� "' + Trim(fcb_DutySpec.Text) + '"OCS ������ ������ �˻��� �Դϴ�....';

   // ������ ���� ��ȸ
   if IsLogonUser then
      SelGridInfo('RELEASESCAN')
end;



procedure TMainDlg.asg_ReleaseKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var
  Res : TPoint;
begin

   case (Key) of

      //[F3 Key] : ���� ã��
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
                                   PChar('ã���ô� �׸��� ȭ�� List�� �������� �ʽ��ϴ�.'),
                                   PChar('������� ����� ��˻� ��� �˸�'),
                                   MB_OK + MB_ICONWARNING)
                                   ;
                  end;

      //[Ctrl + C] : ������ Cell Ŭ������� ���� @ 2016.11.18 LSH
      Ord('C') :
                  if (ssCtrl in Shift) then
                  begin
                     asg_Release.CopySelectionToClipboard;
                  end;

      //[Ctrl + A] : �ش� Grid Cell ��ü���� Ŭ������� ���� @ 2016.11.18 LSH
      Ord('A') :
                  if (ssCtrl in Shift) then
                  begin
                     asg_Release.CopyToClipBoard;
                  end;

      //[Ctrl + F] : �˻� EditBox ��Ŀ�� @ 2016.11.18 LSH
      Ord('F') :
                  if (ssCtrl in Shift) then
                  begin
                     // ����� ��˻� True ����
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
   // 1. ��, ������ �� �޼��� ����.
   //------------------------------------------------------------------
   sType        := '13';
   sSender      := FsUserCallNo;
   sReceiver    := DelChar(asg_Congrats.Cells[2, asg_Congrats.Row], '-');
   sMessage     := '[' + FsUserNm + '���� Happy-Birth SMS �����˸�] ������ �������� �����մϴ� ^ ^/';
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
              PChar('���� SMS�� �߼ۿϷ� �Ǿ����ϴ�.'),
              '[KUMC ���̾�α�] ���� ���� SMS�߼� �˸� ',
              MB_OK + MB_ICONINFORMATION);

end;

procedure TMainDlg.fsbt_DutyClick(Sender: TObject);
begin
   // Comments
   lb_RegDoc.Caption := '�� �μ� �ٹ�/����(�ް�) ��Ȳ�� ��ȸ �� ����Ͻ� �� �ֽ��ϴ�.';

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

   // D/B ���ѿ� ���� �ٹ�ǥ ���� CheckBox ���� @ 2015.06.12 LSH
   if GetMsgInfo('INFORM',
                 'DUTYLIST') = FsUserNm then
   begin
      fcb_DutyMake.Enabled    := True;
   end
   else
   begin
      fcb_DutyMake.Enabled    := False;
   end;

   // �ٹ�ǥ Ȯ��/���� Visible ���� @ 2015.06.12 LSH
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

   apn_DutyList.Caption.Text := '�ٹ�ǥ ���� ����';

   // �ٹ�ǥ ���� ��������
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
   // �ٹ������� Edit ���� --> �ű� ������Ʈ �����, ������ �����ؼ� �Է� �����ϵ��� �ϱ� ���� �ּ�ó��
   // D/B ���ѿ� ���� CanEdit ���� ���� @ 2015.06.12 LSH
   if GetMsgInfo('INFORM',
                 'DUTYLIST') = FsUserNm then
   begin
      // ����(SEQNO) ������ ������ ���� Enabled
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
                 '�����ٹ� ���� ������ ���� User �Դϴ�.',
                 '�����ٹ� ���� ���� �̴�� �˸�',
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
                     TokenStr(Trim(Cells[C_RL_SRREQNO, Row]) + #13#10, #13#10, 1),  // S/R ��ȣ 2���̻��� ��� ���� ���� S/R ��ȣ�� �߶� ����ȸ @ 2016.11.18 LSH
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
               (ftc_Dialog.ActiveTab = AT_DIALBOOK)     or  // [���̾� Book] ����� ������ : S/R �󼼺��� ���
               (
                  (PosByte('S/R',    Cells[C_WK_GUBUN, Row]) = 0) and
                  (ftc_Dialog.ActiveTab = AT_ANALYSIS)        and
                  (asg_Analysis.Visible)
               ) or  // [��û�� ������] : S/R��û��Ȳ �󼼺��� ���
               (
                  (PosByte('������', Cells[C_WK_GUBUN, Row]) > 0) and
                  (PosByte('<', Cells[C_WK_CONTEXT, Row]) > 0)    and
                  (PosByte('>', Cells[C_WK_CONTEXT, Row]) > 0)
               )    // [�ְ�����Ʈ] > [������]������ S/R �� ���� @ 2016.12.08 LSH
            ) then
         begin
            GetLinkList('ANALYSIS',
                        CopyByte(Cells[C_WK_CONTEXT, Row], PosByte('<', Cells[C_WK_CONTEXT, Row]) + 1, LengthByte(Cells[C_WK_CONTEXT, Row]) - PosByte('<', Cells[C_WK_CONTEXT, Row]) - 1),
                        '',
                        '',
                        FsUserIp);

            // �˻����� Clear
            fed_Analysis.Text := '';
         end
         else if (ftc_Dialog.ActiveTab = AT_DIALOG)   or
                 (ftc_Dialog.ActiveTab = AT_DIALBOOK) or
                 ((ftc_Dialog.ActiveTab = AT_ANALYSIS) and (asg_Analysis.Visible)) then
         begin
            apn_Weekly.Collaps := True;
            apn_Weekly.Visible := False;

            // �˻����� Clear
            fed_Analysis.Text  := '';

            // ����� �˻� visible
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

      // �˻����� Clear
      fed_Analysis.Text := '';

      // ����� �˻� visible
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
   showmessage('�ٹ� �� �ް� Ű����˻� ���� �غ���.');

   // �α� Update
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
                 PChar('[��󿬶���] - [���������]�� ����� IP ����� ����Ͻ� �� �ֽ��ϴ�.'),
                 PChar(Self.Caption + ' : ��ϵ��� ���� User ���� �˸�'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;

   if ((ACol = C_B_REGDATE) or
       (ACol = C_B_REGUSER))  and
      (asg_Board.Cells[C_B_BOARDSEQ, ARow] = '��') then
   begin

      // ÷�� �̹��� FTP ���ϳ����� �����ϸ�, �̹��� Download
      if (asg_Board.Cells[C_B_HIDEFILE, ARow - 1] <> '') then
      begin
         // ���� ��/�ٿ�ε带 ���� ���� ��ȸ
         if not GetBinUploadInfo(sServerIp, sFtpUserID, sFtpPasswd, sFtpHostName, sFtpDir) then
         begin
            MessageDlg('���� ������ ���� ����� ���� ��ȸ��, ������ �߻��߽��ϴ�.', Dialogs.mtError, [Dialogs.mbOK], 0);
            exit;
         end;


         // ��������� ���� IP
         sServerIp := C_KDIAL_FTP_IP;


         // ���� ������ ����Ǿ� �ִ� ���ϸ� ����
         if PosByte('/ftpspool/KDIALFILE/',asg_Board.Cells[C_B_HIDEFILE, ARow - 1]) > 0 then
            sRemoteFile := asg_Board.Cells[C_B_HIDEFILE, ARow]
         else
            sRemoteFile := '/ftpspool/KDIALFILE/' + asg_Board.Cells[C_B_HIDEFILE, ARow - 1];

         // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
         sLocalFile  := G_HOMEDIR + 'TEMP\SPOOL\' + asg_Board.Cells[C_B_ATTACH, ARow - 1];



         if (GetBINFTP(sServerIp, sFtpUserID, sFtpPasswd, sRemoteFile, sLocalFile, False)) then
         begin
            //	�������� FTP �ٿ�ε�
         end;


         try
            if (PosByte('.exe', sLocalFile) > 0) or
               (PosByte('.zip', sLocalFile) > 0) or
               (PosByte('.rar', sLocalFile) > 0) then
            begin
               MessageBox(self.Handle,
                          PChar('÷������(' + sLocalFile + ') �ٿ�ε尡 �Ϸ�Ǿ����ϴ�.' + #13#10 + #13#10 +
                                '�� �ӽ� �ٿ�ε� ���� --> C:\KUMC(_DEV)\TEMP\SPOOL\'),
                          '÷������ �ٿ�ε� ���� �Ϸ�',
                          MB_OK + MB_ICONINFORMATION);
            end
            else
            begin
               ShellExecute(HANDLE, 'open',
                            PCHAR(asg_Board.Cells[C_B_ATTACH, ARow - 1]),
                            PCHAR(''),
                            // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
                            PCHAR(G_HOMEDIR + 'TEMP\SPOOL\'),
                            SW_SHOWNORMAL);
            end;

         except
            MessageBox(self.Handle,
                       PChar('�ش� ÷�� ���� ������ ������ �߻��Ͽ����ϴ�.' + #13#10 + #13#10 +
                             '���� �� �ٽ� ������ �ֽñ� �ٶ��ϴ�.'),
                       '÷������ �ٿ�ε� ���� ����',
                       MB_OK + MB_ICONERROR);

            Exit;
         end;


         {
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
      // ���� ����Ʈ(������) ��ȸ
      fsbt_WeeklyClick(Sender)
   else if ARow = R_PR_PHOTO then
   begin
      // ��ȿ�� üũ
      if PosByte(asg_UserProfile.Cells[1, R_PR_USERNM] , asg_NetWork.Cells[C_NW_USERNM, asg_NetWork.Row]) = 0 then
      begin
         MessageBox(self.Handle,
                    PChar('���� ������ ��ȸ���� ����������� ���õ� ������ ��ġ���� �ʽ��ϴ�.' + #13#10 + '����� �̸��� ��Ŭ����, �̹��� ��� �������ּ���.'),
                    '��󿬶��� ����� ������ �̹��� ����� Ȯ��',
                    MB_OK + MB_ICONERROR);
         Exit;
      end;


      // �̹��� ��� ��ư ����
      asg_UserProfile.RemoveButton(1, R_PR_PHOTO);


      if od_File.Execute then
      begin
         // ÷������ Upload
         if (Trim(ExtractFileName(od_File.FileName)) <> '') then
         begin

            sFileName := ExtractFileName(Trim(ExtractFileName(od_File.FileName)));
            sHideFile := 'KDIALAPPEND' + 'IMG' + FormatDateTime('YYYYMMDDHHNNSS', Now) + '.KDIALFILE';


            if Not FileUpLoad(sHideFile, sFileName, sServerIp) then
            begin
               Showmessage('÷������ UpLoad �� ������ �߻��߽��ϴ�.' + #13#10 + #13#10 +
                           '�ٽ��ѹ� �õ��� �ֽñ� �ٶ��ϴ�.');
               Exit;
            end
            else
            begin
            {
               MessageBox(self.Handle,
                          PChar('�̹����� ���������� ��ϵǾ����ϴ�.' + #13#10 + '����� ������ ����ȸ�� ��ϵ� �̹����� ���� �� �ֽ��ϴ�.'),
                          '��󿬶��� ����� ������ �̹��� ���',
                          MB_OK + MB_ICONINFORMATION);
             }

               // ����� �����Ϳ� �̹��� ���� Update
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
// Panel ���� �����ϱ�
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

         // �������� ������Ʈ ���� Free @ 2014.07.31 LSH
         with asg_MultiIns do
         begin
            // [�������� �ۼ�����] ���� TMemo�� ����.
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

         // �������� ������Ʈ ���� Free @ 2014.07.31 LSH
         with asg_MultiIns do
         begin
            // [�������� �ۼ�����] ���� TMemo�� ����.
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

         // �������� ������Ʈ ���� Free @ 2014.07.31 LSH
         with asg_MultiIns do
         begin
            // [�ۼ�����] ���� TMemo�� ����.
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
            Cells[0, 0] := '�ٹ�ó(�Ҽ�)';
            Cells[0, 1] := '�ۼ�����';
            Cells[0, 2] := '��������';
            Cells[0, 8] := '';

            // �ٹ�ó Set
            if PosByte('�Ⱦϵ�����', FsUserIp) > 0 then
            begin
               if PosByte('��ȹ', FsUserPart) > 0 then
                  Cells[1, 0] := 'MIS'
               else
                  Cells[1, 0] := 'AA';
            end
            else if PosByte('���ε�����', FsUserIp) > 0 then
               Cells[1, 0] := 'GR'
            else if PosByte('�Ȼ굵����', FsUserIp) > 0 then
               Cells[1, 0] := 'AS';


            // �ۼ�����, �������� Set
            // ���� ������ ���� ���� Set @ 2017.10.26 LSH
            Cells[1, 1] := CopyByte(asg_WorkRpt.Cells[C_WR_REGDATE, asg_WorkRpt.Row], 1, 10); //FormatDateTime('yyyy-mm-dd', Date);


            // �������� �������� �� �ۼ����� ������, �������� �ۼ� Memo �÷����� ���� @ 2014.07.31 LSH
            //if (PosByte(FormatDateTime('yyyy-mm-dd', Date), asg_WorkRpt.Cells[C_WR_REGDATE, asg_WorkRpt.Row]) > 0) then
            //begin
               // ���� ������ ������ �������� ���� Set @ 2017.10.26 LSH
               Cells[1, 2] := asg_WorkRpt.Cells[C_WR_CONTEXT, asg_WorkRpt.Row];
            //end;

            AddComment(0,
                       1,
                       '�������� �ۼ� Panel (�Ǵ� Fixed-Column) ����Ŭ����, �ۼ�ȭ���� ������ �� �ֽ��ϴ�.');

            AddComment(0,
                       2,
                       '�������� �ۼ��� �Ʒ� [���]��ư�� ������ ����˴ϴ�.');

            // Merge Cells
            MergeCells(0, 2, 1, 6);
            MergeCells(1, 2, 1, 6);


            AddButton(1,
                      8,
                      ColWidths[1]-5,
                      20,
                      '�������� ���',
                      haBeforeText,
                      vaCenter);

         end;

         // Panel Display
         apn_MultiIns.Caption.Text  := '�������� �ۼ�';
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
            Cells[0, 0] := '�ٹ�ó(�Ҽ�)';
            Cells[0, 1] := '�������';
            Cells[0, 2] := '�ֿ����';
            Cells[0, 8] := '';

            // �ٹ�ó Set
            if PosByte('�Ⱦϵ�����', FsUserIp) > 0 then
            begin
               if FsUserNm = FsMngrNm then
                  Cells[1, 0] := '�Ƿ��'
               else
                  Cells[1, 0] := '�Ⱦ�';
            end
            else if PosByte('���ε�����', FsUserIp) > 0 then
               Cells[1, 0] := '����'
            else if PosByte('�Ȼ굵����', FsUserIp) > 0 then
               Cells[1, 0] := '�Ȼ�';


            // �������, �ֿ���� Set
            Cells[1, 1] := FormatDateTime('yyyy-mm-dd', Date);


            // �������� �ֿ���� �ۼ����� ������, �ֿ���� �ۼ� Memo �÷����� ���� @ 2014.07.31 LSH
            if (PosByte(FormatDateTime('yyyy-mm-dd', Date), asg_WorkConn.Cells[C_WC_REGDATE, asg_WorkConn.Row]) > 0) then
            begin
               Cells[1, 2] := asg_WorkConn.Cells[C_WC_CONTEXT, asg_WorkConn.Row];
            end;

            AddComment(0,
                       1,
                       '�ֿ�������� �ۼ� Panel ����Ŭ����, �ۼ�ȭ���� �����մϴ�.');

            AddComment(0,
                       2,
                       '�ֿ�������� �ۼ��� �Ʒ� ��Ϲ�ư�� ������ ����˴ϴ�.');

            // Merge Cells
            MergeCells(0, 2, 1, 6);
            MergeCells(1, 2, 1, 6);


            AddButton(1,
                      8,
                      ColWidths[1]-5,
                      20,
                      '�ֿ�������� ���',
                      haBeforeText,
                      vaCenter);

         end;

         // Panel Display
         apn_MultiIns.Caption.Text  := '�ֿ�������� �ۼ�';
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
            Cells[0, 0] := '�ٹ�ó(�Ҽ�)';
            Cells[0, 1] := '�������';
            Cells[0, 2] := '�ֿ����';
            Cells[0, 8] := '';

            // �ٹ�ó Set
            if PosByte('�Ⱦϵ�����', FsUserIp) > 0 then
            begin
               if FsUserNm = FsMngrNm then
                  Cells[1, 0] := '�Ƿ��'
               else
                  Cells[1, 0] := '�Ⱦ�';
            end
            else if PosByte('���ε�����', FsUserIp) > 0 then
               Cells[1, 0] := '����'
            else if PosByte('�Ȼ굵����', FsUserIp) > 0 then
               Cells[1, 0] := '�Ȼ�';


            // ���� �����Ϸ��� ���� �ۼ����� ��ȸ
            Cells[1, 1] := asg_WorkConn.Cells[C_WC_REGDATE, asg_WorkConn.Row];
            Cells[1, 2] := asg_WorkConn.Cells[C_WC_CONTEXT, asg_WorkConn.Row];


            AddComment(0,
                       1,
                       '�ֿ�������� �ۼ� Panel ����Ŭ����, �ۼ�ȭ���� �����մϴ�.');

            AddComment(0,
                       2,
                       '�ֿ�������� �ۼ��� �Ʒ� ��Ϲ�ư�� ������ ����˴ϴ�.');

            // Merge Cells
            MergeCells(0, 2, 1, 6);
            MergeCells(1, 2, 1, 6);


            AddButton(1,
                      8,
                      ColWidths[1]-5,
                      20,
                      '�ֿ�������� ���',
                      haBeforeText,
                      vaCenter);

         end;

         // Panel Display
         apn_MultiIns.Caption.Text  := '�ֿ�������� ����';
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
   // �������� �ۼ�(TMemo ���� ������Ʈ) Panel Close @ 2014.07.31 LSH
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
   fed_Analysis.Hint       := 'D/B���� ��ȸ�Ϸ��� Ű����(��: Table��, Column��, proc./func./trigger ��)�� �Է���  Enter�� ����������.';


   fsbt_Analysis.ColorFocused := $00D7BF8B;
   fsbt_WorkRpt.ColorFocused  := $00D7BF8B;
   fsbt_Weekly.ColorFocused   := $00D7BF8B;
   fsbt_DBMaster.ColorFocused := $00D7BF8B;

   fed_DateTitle.ColorFlat := $00D7BF8B;
   fmed_AnalFrom.ColorFlat := $00D7BF8B;
   fmed_AnalTo.ColorFlat   := $00D7BF8B;

   fed_DateTitle.Text      := ' �˻�';
   label8.Left             := 999;
   fmed_AnalFrom.Left      := 999;
   fmed_AnalTo.Left        := 999;

   asg_Analysis.Top      := 999;
   asg_Analysis.Left     := 999;
   asg_Analysis.Visible  := False;


   lb_Analysis.Caption := '�� D/B �м��� ����, D/B ������!';

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


   // [��������] ������� Code ��ȸ
   GetComboBoxList('KDIAL',
                   'LOCATE',
                   fcb_WorkArea);

   // [�˻�����] ������� Code ��ȸ
   GetComboBoxList('KDIAL',
                   'DBLIST',
                   fcb_WorkGubun);

   // [��Ʈ����] ������� Code ��ȸ
   GetComboBoxList('KDIAL',
                   'DUTYPART',
                   fcb_DutyPart);

   // ������ IP �ĺ���, �������� Assign
   if PosByte('�Ⱦϵ�����', FsUserIp) > 0 then
      fcb_WorkArea.Text := '�Ⱦ�'
   else if PosByte('���ε�����', FsUserIp) > 0 then
      fcb_WorkArea.Text := '����'
   else if PosByte('�Ȼ굵����', FsUserIp) > 0 then
      fcb_WorkArea.Text := '�Ȼ�';

   // [�˻�����] �⺻ ���ǰ�
   fcb_WorkGubun.Text := fcb_WorkGubun.Items.Strings[0];

   // [��Ʈ����] �Ҽ���Ʈ�� ����, �⺻ ������ "����"
   if PosByte('����', FsUserPart) > 0 then
   begin
      if PosByte('P/L', FsUserSpec) > 0 then
         fcb_DutyPart.Text := CopyByte(FsUserSpec, 1, PosByte('(', FsUserSpec) - 1)
      else
         fcb_DutyPart.Text := FsUserSpec;
   end
   else
      fcb_DutyPart.Text := '����';


   // D/B ������ ī�װ� ��ȸ
   if IsLogonUser then
   begin
      SelGridInfo('DBCATEGORY');

      if not fcb_ReScan.Checked then
         fcb_ReScan.Checked := True;

      // �˻� EditBox ��Ŀ�� @ 2016.11.18 LSH
      fed_Analysis.SetFocus;
   end;


end;



//---------------------------------------------------------------------------
// [��ȸ] FlatComboBox ���� Code ��ȸ
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
      // 2. ComboBox List ���
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
   // D/B ������ ī�װ� �� ��ȸ
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
// Grid �÷� ���� ����
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
         Cells[C_DBM_FSTCDNM, 0]    := 'Table��';
         Cells[C_DBM_WORKINFO,0]    := '������';

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
         Cells[C_DBD_SECCDNM, 0]    := 'Column��';
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
         Cells[C_DBM_WORKINFO,0]    := '������';

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
         Cells[C_DBM_WORKINFO,0]    := '������';

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
         Cells[C_DBM_WORKINFO,0]    := '������';

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
         Cells[C_DBM_WORKINFO,0]    := '������';

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
   else if (in_Flag = '�����ڵ�') then
   begin
      with asg_DBMaster do
      begin
         ColWidths[C_DBM_FSTCD]     := 55;
         ColWidths[C_DBM_FSTCDNM]   := 190;
         ColWidths[C_DBM_WORKINFO]  := 47;

         Cells[C_DBM_FSTCD,   0]    := '��з�';
         Cells[C_DBM_FSTCDNM, 0]    := '��з���';
         Cells[C_DBM_WORKINFO,0]    := '�۾���';

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

         Cells[C_DBD_SECCD,   0]    := '�ߺз�';
         Cells[C_DBD_SECCDNM, 0]    := '�ߺз���';
         Cells[C_DBD_3RDCD,   0]    := '�Һз�';
         Cells[C_DBD_3RDCDNM, 0]    := '�Һз���';
         Cells[C_DBD_WORKINFO,0]    := '�۾���';

         Left  := 312;
         Width := 332;

      end;
   end
   else if (in_Flag = 'SLIP�ڵ�') then
   begin
      with asg_DBMaster do
      begin
         ColWidths[C_DBM_FSTCD]     := 55;
         ColWidths[C_DBM_FSTCDNM]   := 190;
         ColWidths[C_DBM_WORKINFO]  := 47;

         Cells[C_DBM_FSTCD,   0]    := '�׷�Slip';
         Cells[C_DBM_FSTCDNM, 0]    := '�׷��';
         Cells[C_DBM_WORKINFO,0]    := '�۾���';

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

         Cells[C_DBD_SECCD,   0]    := '��Slip';
         Cells[C_DBD_SECCDNM, 0]    := '��Slip��';
         Cells[C_DBD_3RDCD,   0]    := '�μ�';
         Cells[C_DBD_3RDCDNM, 0]    := 'Type';
         Cells[C_DBD_WORKINFO,0]    := '�۾���';

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
         Cells[C_DBM_WORKINFO,0]    := '������';

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
         Cells[C_DBD_SECCDNM, 0]    := '���̺��';
         Cells[C_DBD_3RDCD,   0]    := 'Comments';
         Cells[C_DBD_3RDCDNM, 0]    := '�÷���';
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

      //[Ctrl + C] : ������ Cell Ŭ������� ����
      Ord('C') : if (ssCtrl in Shift) then
                 begin
                    asg_DBMaster.CopySelectionToClipboard;
                 end;

      //[Ctrl + A] : �ش� Grid Cell ��ü���� Ŭ������� ����
      Ord('A') : if (ssCtrl in Shift) then
                 begin
                    asg_DBMaster.CopyToClipBoard;
                 end;

      //[F3 Key] : ���� ã��
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
                                   PChar('ã���ô� �׸��� ȭ�� List�� �������� �ʽ��ϴ�.'),
                                   PChar('D/B �˻� : ����� ��˻� ��� �˸�'),
                                   MB_OK + MB_ICONWARNING)
                                   ;
                  end;

      //[Ctrl + F] : �˻� EditBox ��Ŀ�� @ 2016.11.18 LSH
      Ord('F') :
                  if (ssCtrl in Shift) then
                  begin
                     // ����� ��˻� True ����
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

      //[Ctrl + C] : ������ Cell Ŭ������� ����
      Ord('C') : if (ssCtrl in Shift) then
                 begin
                    asg_DBDetail.CopySelectionToClipboard;
                 end;

      //[Ctrl + A] : �ش� Grid Cell ��ü���� Ŭ������� ����
      Ord('A') : if (ssCtrl in Shift) then
                 begin
                    asg_DBDetail.CopyToClipBoard;
                 end;

      //[F3 Key] : ���� ã��
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
                                   PChar('ã���ô� �׸��� ȭ�� List�� �������� �ʽ��ϴ�.'),
                                   PChar('D/B �˻� : ����� ��˻� ��� �˸�'),
                                   MB_OK + MB_ICONWARNING)
                                   ;
                  end;

      //[Ctrl + F] : �˻� EditBox ��Ŀ�� @ 2016.11.18 LSH
      Ord('F') :
                  if (ssCtrl in Shift) then
                  begin
                     // ����� ��˻� True ����
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

   // ��󿬶��� ������ �̹��� ���
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
   // �������� �ۼ� (�ű� or ����)
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
   // �ֿ�������� �ۼ� (�ű� or ����) @ 2014.08.08 LSH
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
   // �ֿ�������� �ۼ� ���� @ 2014.08.12 LSH
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

         // ��󿬶��� ������ �̹��� ���
         if (in_Gubun = 'PROFILE') then
         begin
            MessageBox(self.Handle,
                       PChar('�̹��������� ���������� [������Ʈ]�Ǿ����ϴ�.'),
                       '[KUMC ���̾�α�] ����� ������ ������Ʈ �˸� ',
                       MB_OK + MB_ICONINFORMATION);

            // Refresh
            SelGridInfo('DIALOG');
         end
         else if (in_Gubun = 'WORKRPT') then
         begin
            MessageBox(self.Handle,
                       PChar('���������� ���������� [������Ʈ]�Ǿ����ϴ�.'),
                       '[KUMC ���̾�α�] �������� ������Ʈ �˸� ',
                       MB_OK + MB_ICONINFORMATION);

            // �ۼ� Panel Close
            SetPanelStatus('WORKRPT', 'OFF');

            // Refresh
            SelGridInfo('WORKRPT');
         end
         else if (in_Gubun = 'WORKCONN') or
                 (in_Gubun = 'WORKCONNDEL')  then
         begin
            MessageBox(self.Handle,
                       PChar('�ֿ� ���������� ���������� [������Ʈ]�Ǿ����ϴ�.'),
                       '[KUMC ���̾�α�] �������� ������Ʈ �˸� ',
                       MB_OK + MB_ICONINFORMATION);

            // �ۼ� Panel Close
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
      // ��ȿ�� üũ
      if PosByte(asg_UserProfile.Cells[1, R_PR_USERNM] , asg_NetWork.Cells[C_NW_USERNM, asg_NetWork.Row]) = 0 then
      begin
         MessageBox(self.Handle,
                    PChar('���� ������ ��ȸ���� ����������� ���õ� ������ ��ġ���� �ʽ��ϴ�.' + #13#10 + '����� �̸��� ��Ŭ����, �̹��� ��� �������ּ���.'),
                    '��󿬶��� ����� ������ �̹��� ����� Ȯ��',
                    MB_OK + MB_ICONERROR);
         Exit;
      end;


      // �̹��� ��� ��ư ����
      asg_UserProfile.RemoveButton(1, R_PR_PHOTO);


      if od_File.Execute then
      begin
         // ÷������ Upload
         if (Trim(ExtractFileName(od_File.FileName)) <> '') then
         begin
            // ���ϸ� �� ��ȣȭ�� Set
            sFileName := ExtractFileName(Trim(ExtractFileName(od_File.FileName)));
            sHideFile := 'KDIALAPPEND' + 'PRIMG' + FormatDateTime('YYYYMMDDHHNNSS', Now) + '.KDIALFILE';


            if Not FileUpLoad(sHideFile, sFileName, sServerIp) then
            begin
               Showmessage('÷������ UpLoad �� ������ �߻��߽��ϴ�.' + #13#10 + #13#10 +
                           '�ٽ��ѹ� �õ��� �ֽñ� �ٶ��ϴ�.');
               Exit;
            end
            else
            begin
               // ����� �����Ϳ� �̹��� ���� Update
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
   // �������� �ۼ� Panel-On
   SetPanelStatus('WORKRPT', 'ON');

end;

procedure TMainDlg.asg_MultiInsGetEditorType(Sender: TObject; ACol,
  ARow: Integer; var AEditor: TEditorType);
begin
   with asg_MultiIns do
   begin
      if PosByte('��������', apn_MultiIns.Caption.Text) > 0 then
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
      if PosByte('��������', apn_MultiIns.Caption.Text) > 0 then
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
                      '�������� ���',
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
   // �������� �ۼ���, �ٹ�ó(�Ҽ�)�� Edit �Ұ�
   if PosByte('��������', apn_MultiIns.Caption.Text) > 0 then
      CanEdit := (ARow = 1) or
                 (ARow = 2);

end;

procedure TMainDlg.asg_MultiInsButtonClick(Sender: TObject; ACol,
  ARow: Integer);
begin
   // �������� ���(�ű� or ����)
   if PosByte('��������', apn_MultiIns.Caption.Text) > 0 then
   begin
      // ���� ���� TMemo ������ �����ϰ�, ������Ʈ ����.
      if asg_MultiIns.Objects[1,2] <> nil then
      begin
         asg_MultiIns.Cells[1,2] := TMemo(asg_MultiIns.Objects[1,2]).Text;

         asg_MultiIns.Objects[1,2].Free;
         asg_MultiIns.Objects[1,2] := Nil;
      end;

      // �������� �ۼ����� ������, Msg ó��.
      if asg_MultiIns.Cells[1,2] = '' then
      begin
         MessageBox(self.Handle,
                    PChar('������Ʈ �� �������� �ۼ������� �����ϴ�.'),
                    '�������� �ۼ����� Update ����',
                    MB_OK + MB_ICONERROR);

         asg_MultiIns.SetFocus;

         Exit;
      end;


      // �������� ���� Update
      UpdateImage('WORKRPT',
                  '',
                  '',
                  1);

   end
   // �ֿ�������� ��� (�ű� or ����)
   else if PosByte('�ֿ��������', apn_MultiIns.Caption.Text) > 0 then
   begin
      // ���� ���� TMemo ������ �����ϰ�, ������Ʈ ����.
      if asg_MultiIns.Objects[1,2] <> nil then
      begin
         asg_MultiIns.Cells[1,2] := TMemo(asg_MultiIns.Objects[1,2]).Text;

         asg_MultiIns.Objects[1,2].Free;
         asg_MultiIns.Objects[1,2] := Nil;
      end;

      // �ֿ�������� �ۼ����� ������, Msg ó��.
      if asg_MultiIns.Cells[1,2] = '' then
      begin
         MessageBox(self.Handle,
                    PChar('������Ʈ �� �ֿ�������� �ۼ������� �����ϴ�.'),
                    '�ֿ�������� �ۼ����� Update ����',
                    MB_OK + MB_ICONERROR);

         asg_MultiIns.SetFocus;

         Exit;
      end;


      // �ֿ�������� ���� Update
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
// [�������] FTP �̹��� ���� ��� �Լ�
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
   //  �Լ����� : ����Ŭ�Լ��� LPAD�� ���� ����� �Ѵ�.
   //             ���ϴ� SIZE��ŭ ������ ���ϴ� ���ڿ��� ä���.
   //  �� �� �� : 1.ASourceStr(String) : ����ڿ�
   //             2.PadLen(String)     : ä�� size
   //             3.PadStr(String)     : ä�﹮��
   //  ���ϱ��� : LPAD�� String
   //  �� �� �� : A := LPAD('123',7,'0') --> '0000123'
   //  �Լ���ó : SQCOMCLS.pas
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
   // Printer ���� Check
   //---------------------------------------------------------------
   //if (IsNotPrinterReady) then exit;



   Screen.Cursor := crHourglass;



   try
      //------------------------------------------------------------
      // 1. �ý��� �������� ���� ��û�� ���
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
               // ȸ��⵵ (yyyy-mm)
               if CopyByte(FormatDateTime('mm', Date), 1, 1) = '0' then
                  tmpDocTitle := FormatDateTime('yyyy', Date) + '�� ' + CopyByte(FormatDateTime('mm', Date - 29), 2, 1) + '�� ' + Cells[C_SD_DOCTITLE, Row]
               else
                  tmpDocTitle := FormatDateTime('yyyy', Date) + '�� ' + FormatDateTime('mm', Date - 29) + '�� ' + Cells[C_SD_DOCTITLE, Row];


               qrlb_DocTitle.Caption   := tmpDocTitle;                           // ��������
               qrlb_Company.Caption    := Cells[C_SD_RELDEPT,  Row];             // ���ó
               qrlb_DocName.Caption    := Cells[C_SD_DOCTITLE, Row];             // ����
               qrlb_Period.Caption     := Cells[C_SD_DOCRMK,   Row];             // ���Ⱓ
               qrlb_TotAmt1.Caption    := Cells[C_SD_TOTRMK,   Row];             // ���� �� ���޾�
               qrlb_HqAmt.Caption      := Cells[C_SD_HQREMARK, Row];             // ���� �Ƿ�� ���޾�
               qrlb_AaAmt.Caption      := Cells[C_SD_AAREMARK, Row];             // ���� �Ⱦ� ���޾�
               qrlb_GrAmt.Caption      := Cells[C_SD_GRREMARK, Row];             // ���� ���� ���޾�
               qrlb_AsAmt.Caption      := Cells[C_SD_ASREMARK, Row];             // ���� �Ȼ� ���޾�
               qrlb_TotAmt2.Caption    := Cells[C_SD_TOTRMK,   Row];             // ���� �� ���޾�
               qrlb_DutyUser.Caption   := FsUserNm;                              // ���� �����
               qrlb_Manager.Caption    := FsMngrNm;                              // ���� ������ (����) @ 2014.07.18 LSH
            end;

            qr_KDial.Print;

            Close;
         end;
      end
      //------------------------------------------------------------
      // 2. �ý��� �������� ���� ��û�� �ϰ����
      //       - ��ȿ�� ��ุ �ϰ� ���
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
                  // ��� ����� �׸��� ��� Skip.
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


                  // ȸ��⵵ (yyyy-mm)
                  if CopyByte(FormatDateTime('mm', Date), 1, 1) = '0' then
                     tmpDocTitle := FormatDateTime('yyyy', Date) + '�� ' + CopyByte(FormatDateTime('mm', Date - 29), 2, 1) + '�� ' + Cells[C_SD_DOCTITLE, i]
                  else
                     tmpDocTitle := FormatDateTime('yyyy', Date) + '�� ' + FormatDateTime('mm', Date - 29) + '�� ' + Cells[C_SD_DOCTITLE, i];



                  qrlb_DocTitle.Caption   := tmpDocTitle;                           // ��������
                  qrlb_Company.Caption    := Cells[C_SD_RELDEPT,  i];               // ���ó
                  qrlb_DocName.Caption    := Cells[C_SD_DOCTITLE, i];               // ����
                  qrlb_Period.Caption     := Cells[C_SD_DOCRMK,   i];               // ���Ⱓ
                  qrlb_TotAmt1.Caption    := Cells[C_SD_TOTRMK,   i];               // ���� �� ���޾�
                  qrlb_HqAmt.Caption      := Cells[C_SD_HQREMARK, i];               // ���� �Ƿ�� ���޾�
                  qrlb_AaAmt.Caption      := Cells[C_SD_AAREMARK, i];               // ���� �Ⱦ� ���޾�
                  qrlb_GrAmt.Caption      := Cells[C_SD_GRREMARK, i];               // ���� ���� ���޾�
                  qrlb_AsAmt.Caption      := Cells[C_SD_ASREMARK, i];               // ���� �Ȼ� ���޾�
                  qrlb_TotAmt2.Caption    := Cells[C_SD_TOTRMK,   i];               // ���� �� ���޾�
                  qrlb_DutyUser.Caption   := FsUserNm;                              // ���� �����
                  qrlb_Manager.Caption    := FsMngrNm;                              // ���� ������ (����) @ 2014.07.18 LSH

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
   // �ý��� �������� ���� ��� ��� (FTP)
   KDialPrint('SYSTEM',
              asg_SelDoc);
end;




procedure TMainDlg.pm_DocPopup(Sender: TObject);
begin
   with asg_SelDoc do
   begin
      //--------------------------------------------------
      // �ý��� �������� ���޿�û�� ���
      //--------------------------------------------------
      if (Cells[C_SD_DOCLIST, Row] = '���ó') then
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
   // SMS ���� �׽�Ʈ��
   if (PosByte('�Ⱦϵ�����5.40', FsUserIp) = 0) then
   begin
      MessageDlg('������ ������մϴ�.', Dialogs.mtError, [Dialogs.mbOK], 0);
      Exit;
   end;
   }


//   i         := 0;
   iStartRow := 0;
   iEndRow   := 0;


   // 1. ���� ���� (��ƼRow ����-��)
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



   // 2. �ϰ��Է� ��ȿ��üũ
   //  -- ���콺 Multi-Dragging�� ����� ���.
   for i := iStartRow to iEndRow do
   begin
      with asg_RecvList do
      begin
         // ��󿬶���
         if (ftc_Dialog.ActiveTab = AT_DIALOG) then
         begin
            Cells[0, i - iStartRow + 1] := asg_Network.Cells[C_NW_USERNM, i];
            Cells[1, i - iStartRow + 1] := asg_Network.Cells[C_NW_MOBILE, i];
         end
         // ���̾� Book
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
               (  // SMS �߼� & ���� Photos ���� ���� �߰� @ 2017.07.11 LSH
                  PosByte('010', TokenStr(asg_DialList.Cells[C_DL_CALLNO, asg_DialList.Row], #13#10, 2)) > 0
               ) then
            begin
               Cells[0, i - iStartRow + 1] := asg_DialList.Cells[C_DL_DUTYUSER, i];
               Cells[1, i - iStartRow + 1] := Trim(TokenStr(asg_DialList.Cells[C_DL_CALLNO, asg_DialList.Row], #13#10, 2));
            end
            else if
               (  // SMS �߼� & ���� Photos ���� ���� �߰� @ 2017.07.11 LSH
                  PosByte('010', TokenStr(asg_MyDial.Cells[C_MD_MOBILE, asg_MyDial.Row], #13#10, 2)) > 0
               ) then
            begin
               Cells[0, i - iStartRow + 1] := asg_MyDial.Cells[C_MD_DUTYUSER, i];
               Cells[1, i - iStartRow + 1] := Trim(TokenStr(asg_MyDial.Cells[C_MD_MOBILE, asg_MyDial.Row], #13#10, 2));
            end
            // �� �� Case ���� @ 2017.10.26 LSH
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


   { Ctrl Ű�� Ȱ���� multi Selection �׽�Ʈ ..
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


   {  -- FormShow���� �α��ν����� �������� ������ �ּ�ó�� @ 2014.04.30 LSH
   for j := 1 to asg_Network.RowCount do
   begin
      if asg_Network.Cells[C_NW_USERNM, j] = FsUserNm then
      begin
         FsUserMobile := asg_NetWork.Cells[C_NW_MOBILE, j];

         if (asg_Network.Cells[C_NW_LOCATE, j] = '�Ƿ��') or
            (asg_Network.Cells[C_NW_LOCATE, j] = '�Ⱦ�') then
            FsUserCallNo := '02-920-' + asg_Network.Cells[C_NW_CALLNO, j]
         else if (asg_Network.Cells[C_NW_LOCATE, j] = '����') then
            FsUserCallNo := '02-2626-' + asg_Network.Cells[C_NW_CALLNO, j]
         else if (asg_Network.Cells[C_NW_LOCATE, j] = '�Ȼ�') then
            FsUserCallNo := '031-412-' + asg_Network.Cells[C_NW_CALLNO, j];

         Break;
      end;
   end;
   }


   // ���� RowCount ����
   asg_RecvList.RowCount := asg_RecvList.RowCount - 1;



   fm_SMS.Lines.Text    := '�����Է��� 80byte (�ѱ۱��� 40������)���� �����ϸ�, �߽��� ��ȣ�� �繫�� ��ȣ�� �ڵ� ���õ˴ϴ�. �Է��� ���� [����]��ư�� �����ּ���.';
   apn_SMS.Caption.Text := ' ������ SMS';

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
   // ���콺 ��Ŭ��(on)�� row-index �����´�
   //---------------------------------------------------------------
   if (Button = mbLeft) then
   begin
      iSelRow1 := 0;

      with asg_Network do
      begin
         // Mouse ������ ��ġ �޾ƿ���
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
   // ���콺 ��Ŭ��(off)�� row-index �����´�
   //---------------------------------------------------------------
   if (Button = mbLeft) then
   begin
      iSelRow2 := 0;

      with asg_Network do
      begin
         // Mouse ������ ��ġ �޾ƿ���
         MouseToCell(X, Y, NowCol, NowRow);

         iSelRow2 := NowRow;
      end
   end;
end;



procedure TMainDlg.fm_SMSClick(Sender: TObject);
begin
   if PosByte('80byte (�ѱ۱��� 40������)' , fm_SMS.Lines.Text) > 0 then
      fm_SMS.Lines.Clear;
end;


// ������ SMS �߼� (���)
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
      ShowMessage('�޽����� �Է��Ͻʽÿ�.');
      me_SmsMsg.SetFocus;
      Exit;
   end;
   }

   {
   if MessageDlg('���õ� �����ڵ鿡�� �޽����� �߼��Ͻðڽ��ϱ�?',
                  mtConfirmation, [mbYes, mbNo], 0) <> mrYes then Exit;
   }


   iCnt := 0;


   for i := 1 to asg_RecvList.RowCount - 1 do
   begin




        {
        // [��������] ��
        if rbt_Type2.Checked = True then
        begin
           if (cbx_hh.Text = '') or (cbx_mm.Text = '') then
           begin
             ShowMessage('����ð�/���� �Է����ֽʽÿ�.');
             Exit;
           end;


           l_hh := cbx_hh.Text;
           l_mm := cbx_mm.Text;


           if (DelChar(FormatDateTime('yyyy-mm-dd',dtp_Reserve.Date),'-') + l_hh + l_mm + '00')  < FormatDateTime('yyyymmddhh24miss',Now) then
           begin
             ShowMessage('���������� ����ð� ���ĺ��� �����մϴ�.');
             Exit;
           end;

           // ����ð� ����.
           sRsvtime := DelChar(FormatDateTime('yyyy-mm-dd',dtp_Reserve.Date),'-') + l_hh + l_mm + '00';


        end
        else
        begin
           sRsvtime := '';
        end;


        // ������ �߰� (���κ��� ���¿� ����), 2012.02.17 LSH
        if (G_Locate = 'AA') then
           sLocateNm    := '[���ȾϺ���] '
        else if (G_Locate = 'GR') then
           sLocateNm    := '[��뱸�κ���] '
        else if (G_Locate = 'AS') then
           sLocateNm    := '[���Ȼ꺴��] ';
        }


      if PosByte('80byte (�ѱ۱��� 40������)' , fm_SMS.Lines.Text) > 0 then
         if Application.MessageBox('�ۼ��Ͻ� �޼����� ���������� �����ϴ�. ��� �����Ͻðڽ��ϱ�?',
                                   PChar('������ SMS �߼��� Ȯ��'), MB_OKCANCEL) <> IDOK then
            Exit;


      if Trim(fm_SMS.Lines.Text) = '' then
      begin
         MessageBox(self.Handle,
                    PChar('�Է��Ͻ� �޼��� ������ �����ϴ�. Ȯ�����ּ���.'),
                    '������ SMS �߼��� Ȯ��',
                    MB_OK + MB_ICONINFORMATION);

         fm_SMS.SetFocus;
         Exit;
      end;


      if Application.MessageBox(PChar('�ۼ��Ͻ� �޼����� �����ڿ��� [�߼�]�Ͻðڽ��ϱ�?' + #13#10 + #13#10 + fm_SMS.Lines.Text),
                                PChar('������ SMS �߼��� Ȯ��'),
                                MB_OKCANCEL) <> IDOK then
         Exit;



      // ��, ������ �� �޼��� ����.
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




   // iCnt ��ŭ �޼��� ����
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

   // SMS ���ۿϷ� �޼���
   MessageBox(self.Handle,
                       PChar(IntToStr(iCnt) + ' ���� SMS�� �߼�[���] �Ǿ����ϴ�.'),
                       '[KUMC ���̾�α�] ������ SMS �߼� �˸� ',
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

// Grid to Excel ��ȯ
procedure TMainDlg.mi_ExcelClick(Sender: TObject);
begin
   if ftc_Dialog.ActiveTab = AT_DIALOG then
      ExcelSave(asg_Network, '���������� ��󿬶��� [Last Updated  : ' + asg_NwUpdate.Cells[0, 1] + ']', 0)
   else if ftc_Dialog.ActiveTab = AT_DOC then
      ExcelSave(asg_Release, '������ ��������� [' + fmed_RegFrDt.Text + '~' + fmed_RegToDt.Text + ']', 0)
   else if ftc_Dialog.ActiveTab = AT_ANALYSIS then
   begin
      // �ְ�����Ʈ Excel ��ȯ ��� �߰� @ 2016.12.08 LSH
      if (apn_Weekly.Visible) then
         ExcelSave(asg_WeeklyRpt, '���� History ����Ʈ [' + fmed_AnalFrom.Text + '~' + fmed_AnalTo.Text + ']', 0)
      else
         ExcelSave(asg_WorkRpt, '�������� [' + fmed_AnalFrom.Text + '~' + fmed_AnalTo.Text + ']', 0);
   end
   else if ftc_Dialog.ActiveTab = AT_WORKCONN then
      ExcelSave(asg_WorkConn, '�ֿ��������', 0);

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
   if Application.MessageBox('Excelȭ�Ϸ� �����Ͻðڽ��ϱ�?', 'Excel����', MB_OKCANCEL) <> IDOK then
      exit;

   if Disp = 1 then
   begin
      SaveD1 := TSaveDialog.Create(nil);
      SaveD1.InitialDir := 'C:\Users\Administrator\Desktop';
      SaveD1.DefaultExt := 'xls';
      SaveD1.FileName := Title;
      SaveD1.Filter := '����ȭ��(*.xls)|*.xls|Ms';

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

         showmessage('Excel�� ã���� �����ϴ�.');

         ExcelAP1.Disconnect;

         FreeAndNil(ExcelWS1);
         FreeAndNil(ExcelWB1);
         FreeAndNil(ExcelAP1);
         exit;
      end;

      with EXcelWS1 do
      begin
         // Ÿ��Ʋ
         range1 := range[cells.item[3,2],cells.item[ExcelGrid.Colcount +3,ExcelGrid.RowCount+2]];
         Cells.Item[1,3] := Title;

         // ó�� Routine (grid�� ����)
         for i := 0 to ExcelGrid.RowCount - 1 do
         begin
            for j := 0 to ExcelGrid.Colcount - 1 do
            begin
                Cells.Numberformatlocal := '@';  // ���ڷ� �ν��ؾ��� ���
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

      showmessage('Excelȭ�Ϸ� ����Ǿ����ϴ�.');

   end
   else
   begin
       try
          ExcelAP2 := CreateOLEObject('Excel.Application');
          ExcelAP2.WorkBooks.add;
          ExcelAP2.visible := true;

          // ó�� Routine (grid�� ����)
          // Ÿ��Ʋ
          ExcelAP2.worksheets[1].cells[1,3] := Title;


          for i := 0 to ExcelGrid.RowCount - 1 do
          begin
             for j := 0 to ExcelGrid.ColCount - 1 do
             begin
                ExcelAP2.worksheets[1].Cells[i + 3, j + 2] := '''' + String(ExcelGrid.Cells[j,i]);
             end;
          end;

          //Range(����,����)���� Range(����,����)�� ������ -2011.03.30 smpark
          //ExcelAP2.Range['A1',CHR(64+ ExcelGrid.RowCount+1)+IntToStr(i + 3)].select;
          ExcelAP2.Range['A1',CHR(64+ (j + 3))+ inttostr(i+2)].select;
          ExcelAP2.Selection.Font.Name:='����ü';
          ExcelAP2.Selection.Font.Size:=10;
          ExcelAP2.selection.Columns.AutoFit;
          ExcelAP2.Range['A1','A1'].select;

       except
          Screen.Cursor := crDefault;

          ExcelAP2.Quit;

          showmessage('Excel�� ã���� �����ϴ�.');

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

         // ��������� On
         if (apn_Release.Visible) then
         begin
            mi_RlzScript.Visible  := True;
            mi_BPL2Delphi.Visible := True;   // BPL --> ������ �ڵ���� �޴� �߰� @ 2016.11.02 LSH
            mi_DelRelease.Visible := True;   // ������ �̷� ���� @ 2018.06.12 LSH
         end
         // �� �� (�ٹ�ǥ, �������� ��)
         else
         begin
            mi_RlzScript.Visible  := False;
            mi_BPL2Delphi.Visible := False;   // BPL --> ������ �ڵ���� �޴� �߰� @ 2016.11.02 LSH
            mi_DelRelease.Visible := False;   // ������ �̷� ���� @ 2018.06.12 LSH
         end;
      end
      else if ftc_Dialog.ActiveTab = AT_ANALYSIS then
      begin
         // �ְ�����Ʈ Excel ��ȯ �˾��޴� �߰� @ 2016.12.08 LSH
         if (apn_Weekly.Visible) then
         begin
            mi_InsWorkRpt.Visible := False;
            mi_Excel.Visible      := True;
            mi_GridPrint.Visible  := False;
            mi_Modify.Visible     := False;
            mi_Delete.Visible     := False;
            mi_RlzScript.Visible  := False;
            mi_BPL2Delphi.Visible := False;   // BPL --> ������ �ڵ���� �޴� �߰� @ 2016.11.02 LSH
         end
         else
         begin
            mi_InsWorkRpt.Visible := True;
            mi_Excel.Visible      := True;
            mi_GridPrint.Visible  := False;
            mi_Modify.Visible     := False;
            mi_Delete.Visible     := False;
            mi_RlzScript.Visible  := False;
            mi_BPL2Delphi.Visible := False;   // BPL --> ������ �ڵ���� �޴� �߰� @ 2016.11.02 LSH
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
         mi_BPL2Delphi.Visible := False;   // BPL --> ������ �ڵ���� �޴� �߰� @ 2016.11.02 LSH
      end
      // �ְ��������� (���� �̻��)
      else if ftc_Dialog.ActiveTab = AT_WORKCONN then
      begin
         mi_InsWorkRpt.Visible := False;
         mi_Excel.Visible      := True;
         mi_GridPrint.Visible  := False;
         mi_Modify.Visible     := True;
         mi_Delete.Visible     := True;
         mi_RlzScript.Visible  := False;
         mi_BPL2Delphi.Visible := False;   // BPL --> ������ �ڵ���� �޴� �߰� @ 2016.11.02 LSH
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
                 PChar('[��󿬶���] - [���������]�� ����� IP ����� ����Ͻ� �� �ֽ��ϴ�.'),
                 PChar(Self.Caption + ' : ��ϵ��� ���� User ���� �˸�'),
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



      // �α� Update
      UpdateLog('DIALMAP',
                'PRINT',
                FsUserIp,
                'P',
                FsUserSpec,
                FsUserNm,
                varResult
                );


      // Comments
      lb_DialScan.Caption := '�� ���̾�Map ����� �Ϸ�Ǿ����ϴ�.';

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


   // Report Name ����
   RecName := '��������� ������ġ�� (���̾�Map)';
   addInfo := ' Fax no. [�Ⱦ�] 02-920-5672 [����] 02-2626-2239 [�Ȼ�] 031-412-5999';


   // �Ƿ���� Display
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
      ts := ts+ ((tw-textwidth('������б��Ƿ��')) shr 1);

      Textout(ts, -15, '������б��Ƿ��');

      font.assign(savefont);

      savefont.free;
   end;


   // ������� Display
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

   // �������� ver. Display
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


   // ������ Fax. Info Display
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
   // ���콺 ��Ŭ��(on)�� row-index �����´�
   //---------------------------------------------------------------
   if (Button = mbLeft) then
   begin
      iSelRow1 := 0;

      with asg_DialList do
      begin
         // Mouse ������ ��ġ �޾ƿ���
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
   // ���콺 ��Ŭ��(off)�� row-index �����´�
   //---------------------------------------------------------------
   if (Button = mbLeft) then
   begin
      iSelRow2 := 0;

      with asg_DialList do
      begin
         // Mouse ������ ��ġ �޾ƿ���
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
   // ���콺 ��Ŭ��(off)�� row-index �����´�
   //---------------------------------------------------------------
   if (Button = mbLeft) then
   begin
      iSelRow2 := 0;

      with asg_MyDial do
      begin
         // Mouse ������ ��ġ �޾ƿ���
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
   // ���콺 ��Ŭ��(on)�� row-index �����´�
   //---------------------------------------------------------------
   if (Button = mbLeft) then
   begin
      iSelRow1 := 0;

      with asg_MyDial do
      begin
         // Mouse ������ ��ġ �޾ƿ���
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
//       - ���̾�Book �˻��� User ������ Ȯ��
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
      // �� Mouse ������ ��ǥ �޾ƿ���
      MouseToCell(X,
                  Y,
                  NowCol,
                  NowRow);


      // ���̾� Book ����� ������ ��ȸ
      if (NowCol = C_DL_DUTYUSER) and
         (asg_DialList.Cells[C_DL_DUTYUSER, NowRow] <> '') and
         (IsLogonUser)  then
      begin

         if (NowRow > 0) then
         begin

            //lb_DialScan.Caption := '�� ' + Trim(asg_DialList.Cells[C_DL_DUTYUSER, NowRow]) + '���� ����� ������ ��ȸ��.';

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



            // ������ ���� assign
            asg_UserProfile.Cells[1, R_PR_USERNM]     := Cells[C_DL_DUTYUSER, NowRow];
            asg_UserProfile.Cells[1, R_PR_DUTYPART]   := Cells[C_DL_LOCATE,   NowRow]  + ' ' +
                                                         Cells[C_DL_DEPTNM,   NowRow];

            asg_UserProfile.Cells[1, R_PR_REMARK]     := Cells[C_DL_DEPTSPEC, NowRow];
            asg_UserProfile.Cells[1, R_PR_CALLNO]     := Cells[C_DL_CALLNO,   NowRow];

            asg_UserProfile.AddButton(1, R_PR_DUTYSCH, asg_UserProfile.ColWidths[0]+45, 20, 'S/R ��û��Ȳ(1��)', haBeforeText, vaCenter);
            asg_UserProfile.Cells[1, R_PR_USERID]     := Cells[C_DL_DTYUSRID,   NowRow];

            // �̹��� Stream �ʱ�ȭ (JPEG Error #41 ��������) @ 2015.06.03 LSH
            asg_UserProfile.RemovePicture(1, R_PR_PHOTO);

            // ������ �̹��� FTP ���ϳ����� �����ϸ�, �̹��� Download
            if (Cells[C_DL_DTYUSRID, NowRow] <> '') then
            begin
               // ����� ������ �̹��� ���ϸ�
               if Cells[C_DL_DTYUSRID, NowRow] = 'XXXXX' then
               begin
                  // ������ �̹��� Remove
                  asg_UserProfile.RemovePicture(1, R_PR_PHOTO);

                  // �̹��� ��� ��ư ���� --> ���̾�Book ������ ������ �̹��� ��� ����
                  //asg_UserProfile.AddButton(1, R_PR_PHOTO, asg_UserProfile.ColWidths[0]+45, asg_UserProfile.RowHeights[R_PR_PHOTO]-5, '������ �̹��� ���', haBeforeText, vaCenter);

                  Exit;
               end
               else
                  // ����� HRM ������ �̹��� ���ϸ�
                  ServerFileName := Cells[C_DL_DTYUSRID, NowRow] + '.jpg';


                  // ���� �̹����� ���°�� Skip (JPEG ERROR #41 ����) @ 2015.06.03 LSH
                  if AnsiUpperCase(ServerFileName) = '.JPG' then
                     Exit;


                  // Local �����̸� Set
                  // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
                  ClientFileName := G_HOMEDIR + 'TEMP\SPOOL\' + ServerFileName;


                  // Local�� �ش� Image �������� üũ
                  // Local file�� Size üũ �߰� (Local�� �ش� ���.jpg�� 0 size�� �����ϸ� JPEG Error #41 ������) @ 2015.06.03 LSH
                  if (FileExists(ClientFileName)) and
                     (CMsg.GetFileSize(ClientFileName) >  0) then
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
                     FTP_USERID := '';
                     FTP_PASSWD := '';
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

                           Exit;
                        end;

                     except
                        Exit;       // ���� FTP ������ ���.jpg �̹��� �̵�ϵ� ��� try~except ���ļ� Exit
                     end;



                     // �ش� Image Variant�� ����.
                     AppendVariant(DownFile, ClientFileName);


                     // �ٿ�ε� Ƚ�� üũ
                     DownCnt := DownCnt + 1;
                  end;


                  // ������ Image ������ Grid�� ǥ��
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
//       - S/R ��û User ������ Ȯ��
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
      // �� Mouse ������ ��ǥ �޾ƿ���
      MouseToCell(X,
                  Y,
                  NowCol,
                  NowRow);


      // ����� ������ ��ȸ
      if (NowCol = C_AN_REQUSER) and
         (Cells[C_AN_REQUSER, NowRow] <> '') and
         (IsLogonUser) then
      begin
         if (NowRow > 0) then
         begin

            lb_Analysis.Caption := '�� ' + Trim(Cells[C_AN_REQUSER, NowRow]) + '���� ��û�� ������ ��ȸ��.';

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

            // ������ ���� assign
            asg_UserProfile.Cells[1, R_PR_USERNM]     := Cells[C_AN_REQUSER,  NowRow];
            asg_UserProfile.Cells[1, R_PR_DUTYPART]   := Cells[C_AN_REQDEPTNM,NowRow];
            asg_UserProfile.Cells[1, R_PR_REMARK]     := Cells[C_AN_REQSPEC,  NowRow];
            asg_UserProfile.Cells[1, R_PR_CALLNO]     := Cells[C_AN_REQTELNO, NowRow];
            asg_UserProfile.Cells[1, R_PR_USERID]     := Cells[C_AN_USERID,   NowRow];
            asg_UserProfile.AddButton(1, R_PR_DUTYSCH, asg_UserProfile.ColWidths[0]+45, 20, 'S/R ��û��Ȳ(1��)', haBeforeText, vaCenter);


            // ������ �̹��� FTP ���ϳ����� �����ϸ�, �̹��� Download
            if (Cells[C_AN_USERID, NowRow] <> '') then
            begin
               // ����� ������ �̹��� ���ϸ�
               if Cells[C_AN_USERID, NowRow] = 'XXXXX' then
               begin
                  // ������ �̹��� Remove
                  asg_UserProfile.RemovePicture(1, R_PR_PHOTO);

                  // �̹��� ��� ��ư ���� --> ���̾�Book ������ ������ �̹��� ��� ����
                  //asg_UserProfile.AddButton(1, R_PR_PHOTO, asg_UserProfile.ColWidths[0]+45, asg_UserProfile.RowHeights[R_PR_PHOTO]-5, '������ �̹��� ���', haBeforeText, vaCenter);

                  Exit;
               end
               else
                  // ����� HRM ������ �̹��� ���ϸ�
                  ServerFileName := Cells[C_AN_USERID, NowRow] + '.jpg';

                  // Local �����̸� Set
                  // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
                  ClientFileName := G_HOMEDIR + 'TEMP\SPOOL\' + ServerFileName;


                  // Local�� �ش� Image �������� üũ
                  // Local file�� Size üũ �߰� (Local�� �ش� ���.jpg�� 0 size�� �����ϸ� JPEG Error #41 ������) @ 2015.06.03 LSH
                  if (FileExists(ClientFileName)) and
                     (CMsg.GetFileSize(ClientFileName) >  0) then
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
                     FTP_USERID := '';
                     FTP_PASSWD := '';
                     FTP_DIR    := '/kuhrm/app/hrm/photo/';


                     // HRM Image �ٿ�ε�
                     if Not GetBINFTP(FTP_SVRIP,
                                      FTP_USERID,
                                      FTP_PASSWD,
                                      FTP_DIR + ServerFileName,
                                      ClientFileName,
                                      False) then
                     begin
                        //Showmessage('�̹��� �ٿ�ε忡 �����߽��ϴ�. �ٽ��ѹ� �õ��� �ֽʽÿ�');

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
         // ���� Action �޼��� ���� ���� �ּ� @ 2016.11.18 LSH
         //lb_Analysis.Caption := '';

         apn_UserProfile.Collaps := True;
         apn_UserProfile.Visible := False;
      end;
   end;
end;




//------------------------------------------------------------------------------
// [HTML�˾�] ���̾� Pop �޼��� �ǽð� Pop-up
//       - ����(in_Gubun)�� Pop-up �޼��� Set.
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

      // �ߺ� �ν��Ͻ� ����(�޸� ����) ���� @ 2018.07.16 LSH
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

      htmlpopup.Text.Add('<font color="clMaroon"><font size = "9"> :: ���̾� POP ����Ÿ�� �޼��� :: </font><BR><BR>');

      FsPopMsg := FormatDateTime('yyyy-mm-dd hh:nn:ss', Now);

      htmlpopup.Text.Add('<font color="clBlack">[�α���] ');
      htmlpopup.Text.Add(FsUserSpec + ' ' + FsUserNm + '��');
      htmlpopup.Text.Add(' </font><BR><BR><font color = "clMaroon">');
      htmlpopup.Text.Add(FsPopMsg);
      htmlpopup.Text.Add('</font><BR><BR><font color="clGray">Press here to <a href="Close">�ݱ�</a></font>');

      htmlpopup.RollUp;
   end
   else if (in_Gubun = 'DIET') then
   begin
      Result := True;

      // �Ĵܾ˸���, ���� 30�ʵ��� Pop-up
      tm_DialPop.Interval := 30000;

      // �ߺ� �ν��Ͻ� ����(�޸� ����) ���� @ 2018.07.16 LSH
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

      htmlpopup.Text.Add('<font color="clMaroon"><font size = "9"> :: ���̾� POP ����Ÿ�� �޼��� :: </font><BR><BR>');
      htmlpopup.Text.Add('<font color="clBlack">[�Ĵܾ˸�] ');
      htmlpopup.Text.Add(' </font><BR><BR><font color = "clMaroon">');
      htmlpopup.Text.Add(FsPopMsg);
      htmlpopup.Text.Add('</font><BR><BR><font color="clGray">Press here to <a href="Close">�ݱ�</a></font>');

      // ���̳��Ϻ� Anchor �б�
      if (PosByte('����', FsPopMsg) > 0) or
         (PosByte('��ħ', FsPopMsg) > 0) then
         htmlpopup.Text.Add('<BR><BR><font color="clGray">�� �� �˾� �׸����� �� <a href="Diet1Off">Ȯ��</a></font>')
      else if (PosByte('�߽�', FsPopMsg) > 0) or
              (PosByte('����', FsPopMsg) > 0) then
         htmlpopup.Text.Add('<BR><BR><font color="clGray">�� �� �˾� �׸����� �� <a href="Diet2Off">Ȯ��</a></font>')
      else if (PosByte('����', FsPopMsg) > 0) or
              (PosByte('����', FsPopMsg) > 0) then
         htmlpopup.Text.Add('<BR><BR><font color="clGray">�� �� �˾� �׸����� �� <a href="Diet3Off">Ȯ��</a></font>');

      htmlpopup.RollUp;
   end
   else if (in_Gubun = 'RTCHECK') then
   begin
      Result := True;

      // �ǽð� �˶�üũ��, ���� 10�ʵ��� Pop-up
      tm_DialPop.Interval := 10000;

      // �ߺ� �ν��Ͻ� ����(�޸� ����) ���� @ 2018.07.16 LSH
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

      htmlpopup.Text.Add('<font color="clPurple"><font size = "9"> :: ���̾� POP ����Ÿ�� �޼��� :: </font><BR><BR>');
      htmlpopup.Text.Add('<font color = "clMaroon">');
      htmlpopup.Text.Add(FsPopMsg);
      htmlpopup.Text.Add('</font>');
      htmlpopup.Text.Add('<BR><BR><font color="clGray">Press here to <a href="Close">�ݱ�</a></font>');

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
   // Memory ���� üũ ����� �ּ�Ǯ�� @ 2015.05.12 LSH
   //memchk;

   //
   // �� App. ���� Root ���丮
   sRootPath := UpperCase(TokenStr(ExtractFileDir(ParamStr(0)), '\', 1)) +
                '\'                                                      +
                UpperCase(TokenStr(ExtractFileDir(ParamStr(0)), '\', 2)) +
                '\';

   // App. ������ ����ֱ� (�߰�)
   G_HOMEDIR := Format('%sEXE\', [sRootPath]);

//
//   // test
//   showmessage(G_HomeDir);
//
//   //
//   G_ENV_FILENAME := G_HomeDir + 'ENV\TMAX.ENV';

   // Form ������, HTML Pop-up �ʱ�ȭ
   HTMLPopup := nil;

   // �̹���(������, �ڷ� ��) ���� Temp ���� �ʼ� ����
   ForceDirectories(G_HOMEDIR + 'Temp\SPOOL');

end;


//------------------------------------------------------------------------------
// [HTML�˾�] Pop-up Anchor onClick Event Handler
//       - ����(Anchor)�� Close Action �б�
//
// Date   : 2014.07.02
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.AnchorClick(sender: TObject; Anchor: string);
var
   varResult : String;
begin
   // �Ϲ� �ݱ�
   if Anchor = 'Close' then
   begin
      htmlpopup.RollDown;
      htmlpopup.Free;
      htmlpopup := nil;

      // �Ϲ� �˶�off��, üũ Interval 10�ʷ� Reset.
      tm_DialPop.Interval := 10000;
   end
   // �Ĵ�����
   else if PosByte('Diet', Anchor) > 0 then
   begin
      htmlpopup.RollDown;
      htmlpopup.Free;
      htmlpopup := nil;

      // �Ĵܳ��Ϻ� Pop-up Ȯ�� Log update
      UpdateLog('POPUP',
                Anchor,
                FsUserIp,
                'C',
                FormatDateTime('yyyy-mm-dd', Date),
                FsUserNm,
                varResult
                );

      // �Ĵܾ˶� off��, üũ Interval 10�ʷ� Reset.
      tm_DialPop.Interval := 10000;
   end;

   //  �޸� ���� �ǽɺκ� �ּ� @ 2018.07.16 LSH
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
// [Ÿ�̸�] ���̾� Pop onTimer Event Handler
//
// Date   : 2014.07.02
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.tm_DialPopTimer(Sender: TObject);
begin
   // ���̾� Pop-up ��Ȯ�ν�, 10�ʸ��� �ڵ� Roll-Down
   if Assigned(htmlpopup) then
   begin
      htmlpopup.RollDown;
      htmlpopup.Free;
      htmlpopup := nil;
   end;

   //--------------------------------------------------------------------
   // ���߰� Tmax ���� Ȯ�νÿ���, �ǽð� Pop-up ����
   //--------------------------------------------------------------------
   if (TxInit(TokenStr(String(CONST_ENV_FILENAME), '.', 1) + '_D0.ENV', '01')) then
   begin
      // ���̾� Pop �ű� �޼��� ���� Check
      if CreatePopup(CheckDialPop('REALTIME')) then
      begin
         //
      end
      else
         Exit;
   end
   else
      MessageBox(self.Handle,
                 PChar('���� Server���� ������ ������ϴ�.' + #13#10 + #13#10 + '90�ʸ��� �ڵ����� ������ �����Ͽ���, ��ø� ��ٷ� �ֽʽÿ�.'),
                 PChar(Self.Caption + ' : TMAX ���Ӳ��� �˸�'),
                 MB_OK + MB_ICONWARNING);

end;



//------------------------------------------------------------------------------
// [�Լ�] ���̾� Pop �޼��� �ǽð� Check
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
   // 1. �α��� �˾� ����
   //-------------------------------------------------------------
   if (in_Gubun = 'LOGIN') then
   begin
      Result := in_Gubun;

      Exit;
   end;


   //-------------------------------------------------------------
   // 2. �ð��뺰 ���� ���� ���� ��ȸ
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
   // 4-1. ��ȸ
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


      // ���� Flag �� Return Value �б�
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
         lb_RegDoc.Caption := '�� S/R ��ȣ�� ����Ŭ���Ͻø�, �� ��û������ ���� �� �ֽ��ϴ�.'
      else
         lb_RegDoc.Caption := '';
   end;

end;

procedure TMainDlg.asg_SelDocDblClick(Sender: TObject);
begin
   with asg_SelDoc do
   begin
      if PosByte('���', Cells[C_SD_DOCLIST, Row]) > 0 then
      begin
         asg_DocSpec.Cells[0, R_DS_USERID]   := '����';
         asg_DocSpec.Cells[0, R_DS_USERNM]   := '����ü��';
         asg_DocSpec.Cells[0, R_DS_DEPTCD]   := '���Ⱓ';
         asg_DocSpec.Cells[0, R_DS_DEPTNM]   := '';
         asg_DocSpec.Cells[0, R_DS_SOCNO]    := '�д��(�Ѿ�)';
         asg_DocSpec.Cells[0, R_DS_STARTDT]  := '�д��(HQ)';
         asg_DocSpec.Cells[0, R_DS_ENDDT]    := '�д��(AA)';
         asg_DocSpec.Cells[0, R_DS_MOBILE]   := '�д��(GR)';
         asg_DocSpec.Cells[0, R_DS_GRPID]    := '�д��(AS)';
         asg_DocSpec.Cells[0, R_DS_USELVL]   := '������ġ';

         if (FontColors[C_SD_DOCTITLE,  Row] = clRed) then
         begin
            asg_DocSpec.Cells[1, R_DS_USERID]      := Cells[C_SD_DOCTITLE,Row];
            asg_DocSpec.FontColors[1, R_DS_USERID] := clRed;

            asg_DocSpec.Cells[1, R_DS_DEPTCD]      := Cells[C_SD_DOCRMK,Row] + ' [�������]';
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

         apn_DocSpec.Caption.Text := Cells[C_SD_DOCLIST, Row] + ' �������� ������ ��ȸ';
      end;
   end;
end;



procedure TMainDlg.asg_SelDocClick(Sender: TObject);
begin
   if (apn_DocSpec.Visible) and
      (PosByte('��ȸ', apn_DocSpec.Caption.Text) > 0) then
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
      // [�������� �ۼ�����] Click��
      if (ACol = 1) and
         (ARow = 2) then
      begin
         // TMemo ������Ʈ�� Cell Grid�� Object�� ����
         Objects[ACol,ARow] := TMemo.Create(Application);

         // TMemo ������Ʈ �� �Ӽ� ����
         with TMemo(Objects[ACol,ARow]) do
         begin
            Parent     := Self;
            Left       := apn_MultiIns.Left + CellRect(ACol,ARow).Left + 1;
            Top        := apn_MultiIns.Top + CellRect(ACol,ARow).Top   + 1;
            Width      := asg_MultiIns.CellRect(ACol,ARow).Right  - asg_MultiIns.CellRect(ACol,ARow).Left;
            Height     := CellRect(ACol,ARow).Bottom - CellRect(ACol,ARow).Top + 2;

            // ���� �ۼ����� ������, TMemo ������ �����´�
            if (Cells[1,2] <> '') then
               Text := Cells[1,2];

            SetFocus;
         end;
      end
      else
      begin
         // ���� ���� TMemo�� ����.
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
      // ���� ���� TMemo�� ����.
      if Objects[1,2] <> nil Then
      begin
         Objects[1,2].Free;
         Objects[1,2] := Nil;
      end;
   end;
end;


procedure TMainDlg.asg_MultiInsDblClick(Sender: TObject);
begin
   // �������� �ۼ�(TMemo ���� ������Ʈ) Panel Close @ 2014.07.31 LSH
   SetPanelStatus('WORKRPT', 'OFF');
end;

procedure TMainDlg.fsbt_ConnWriteClick(Sender: TObject);
begin
   if (PosByte('����',   FsUserSpec) > 0) or
      (PosByte('����',   FsUserSpec) > 0) or
      (PosByte('��Ʈ��', FsUserSpec) > 0) or
      (PosByte('P/L',    FsUserSpec) > 0) then
   begin
      // �������� ��� Panel-On
      SetPanelStatus('WORKCONN', 'ON');
   end
   else
      MessageBox(self.Handle,
                 PChar('�� ���� ������ �� ��Ʈ P/L �̻� �ۼ� �����մϴ�.'),
                 PChar('�������� �ۼ����� ���Ѿ˸�'),
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
                 PChar('������ �ۼ��� ������ ������ �����մϴ�.'),
                 PChar('�ֿ�������� ������ ���� Ȯ��'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;

   // ���� Confirm Message.
   if Application.MessageBox('������ �Խñ��� [����]�Ͻðڽ��ϱ�?',
                             PChar('�ֿ�������� ���� �˸�'), MB_OKCANCEL) <> IDOK then
      Exit;

   // �ֿ�������� ���� ���� Update
   UpdateImage('WORKCONNDEL',
               FormatDateTime('yyyy-mm-dd', Date),
               '',
               1);
end;

// [�˾�] - [������ �̷� ����] @ 2018.06.12 LSH
procedure TMainDlg.mi_DelReleaseClick(Sender: TObject);
begin
   UpdReleaseDelDt;
end;

// ������ �����丮(�̷�) ���� @ 2018.06.12 LSH
procedure TMainDlg.mi_ModifyClick(Sender: TObject);
begin
   if asg_WorkConn.Cells[C_WC_DUTYUSER, asg_WorkConn.Row] <> FsUserNm then
   begin
      MessageBox(self.Handle,
                 PChar('������ �ۼ��� ������ ������ �����մϴ�.'),
                 PChar('�ֿ�������� ������ ���� Ȯ��'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;

   // �ֿ�������� ���� ���� Update
   SetPanelStatus('WORKCONNMOD', 'ON');

end;



procedure TMainDlg.mi_SystemAllPrintClick(Sender: TObject);
begin
   // �ý��� �������� ���� ��� �ϰ���� (FTP)
   KDialPrint('SYSTEMALL',
              asg_SelDoc);
end;



//------------------------------------------------------------------------------
// CRA/CRC ID �������� ������Ʈ @ 2014.09.03 LSH
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
                 PChar('[��󿬶���] - [���������]�� ����� IP ����� ����Ͻ� �� �ֽ��ϴ�.'),
                 PChar(Self.Caption + ' : ��ϵ��� ���� User ���� �˸�'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;




   sType    := '3';
   sDelDate := FormatDateTime('yyyy-mm-dd', Date);



   // ������ IP �ĺ���, �ٹ�ó Assign
   if PosByte('�Ⱦϵ�����', FsUserIp) > 0 then
      sLocateNm := '�Ⱦ�'
   else if PosByte('���ε�����', FsUserIp) > 0 then
      sLocateNm := '����'
   else if PosByte('�Ȼ굵����', FsUserIp) > 0 then
      sLocateNm := '�Ȼ�';



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


   // CRA/CRC ID ������ Ȯ�� Pop-Up
   if Application.MessageBox(PChar(asg_SelDoc.Cells[C_SD_REGUSER, asg_SelDoc.Row] + '(' + asg_SelDoc.Cells[C_SD_DOCRMK, asg_SelDoc.Row] + ') ���� ID�� [����]�Ͻðڽ��ϱ�?'),
                             PChar('CRA/CRC ID ���� ����� Ȯ��'), MB_OKCANCEL) <> IDOK then
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
                    '�����Ͻ� CRA/CRC ID�� ���������� [����]�Ǿ����ϴ�.',
                    '[KUMC ���̾�α�] CRA/CRC ID ���� ������Ʈ �˸� ',
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



// [��������] - [���콺 �˾�] - [�������� ����] @ 2014.09.17 LSH
procedure TMainDlg.mi_CopyDocInfoClick(Sender: TObject);
begin
   with asg_RegDoc do
   begin
      Cells[C_RD_GUBUN,    Row] := '����';
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
// [Ŀ�´�Ƽ] ����Ű / EnterŰ �Է½� Jog-Control ����
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

      //[Ctrl + C] : ������ Cell Ŭ������� ���� @ 2016.11.18 LSH
      Ord('C') :
                  if (ssCtrl in Shift) then
                  begin
                     asg_Board.CopySelectionToClipboard;
                  end;

      //[Ctrl + A] : �ش� Grid Cell ��ü���� Ŭ������� ���� @ 2016.11.18 LSH
      Ord('A') :
                  if (ssCtrl in Shift) then
                  begin
                     asg_Board.CopyToClipBoard;
                  end;

      //[Ctrl + F] : �˻�ȭ�� ȣ��
      Ord('F') :
                  if (ssCtrl in Shift) then
                  begin
                     fsbt_SearchClick(Sender);
                  end;
   end;
end;



//------------------------------------------------------------------------------
// [Ŀ�´�Ƽ] �� �˻���� Title Ű���� �Է��� �ڵ� �˻�
//
// Date   : 2015.03.19
// Author : Lee, Se-Ha
//------------------------------------------------------------------------------
procedure TMainDlg.fed_TitleKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   if fsbt_WriteFind.Visible then
   begin
      // �� �˻� ��忡��, [����] ���� Ű���� �Է��� <Enter> ������ �ڵ� �˻� �ǽ�
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

      // �ش� Chat - User �α׾ƿ� ���� ������Ʈ
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

      // �ش� Chat - User �α׾ƿ� ���� ������Ʈ
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

      // �α� Update
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
// ���ڿ� ó������ KeyStr�� ������ ���������� �߶�
//      - �ҽ���ó : MComfunc.pas
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


// ����� ������ �̹��� ���� FTP �ٿ�ε� @ 2015.03.31 LSH
function TMainDlg.GetFTPImage(in_PhotoFile, in_HiddenFile, in_UserId : String) : String;
var
   ServerFileName, ClientFileName: String;
   FTP_SVRIP, FTP_USERID, FTP_PASSWD, FTP_HOSTNAME, FTP_DIR : String;
   sRemoteFile, sLocalFile : String;
begin

   Result := '';


   // ������ �̹��� FTP ���ϳ����� �����ϸ�, �̹��� Download
   if (in_PhotoFile <> '') then
   begin
      // MIS �⺻ �̹��� ������ ���.
      if (in_HiddenFile = '') then
      begin

         // ����� ������ �̹��� ���ϸ�
         ServerFileName := in_PhotoFile;


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
            FTP_USERID := '';
            FTP_PASSWD := '';
            FTP_DIR    := '/kumis/app/mis/media/cq/photo/';



            // Image �ٿ�ε�
            if Not GetBINFTP(FTP_SVRIP,
                             FTP_USERID,
                             FTP_PASSWD,
                             FTP_DIR + ServerFileName,
                             ClientFileName,
                             False) then
            begin
               //Showmessage('�̹��� �ٿ�ε忡 �����߽��ϴ�. �ٽ��ѹ� �õ��� �ֽʽÿ�');

               TUXFTP := nil;

               Exit;
            end;



            // �ش� Image Variant�� ����.
            AppendVariant(DownFile, ClientFileName);



            // �ٿ�ε� Ƚ�� üũ
            DownCnt := DownCnt + 1;


            Result := ClientFileName;

         end
         else
            Result := ClientFileName;
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
         if PosByte('/ftpspool/KDIALFILE/', in_HiddenFile) > 0 then
            sRemoteFile := in_HiddenFile
         else
            sRemoteFile := '/ftpspool/KDIALFILE/' + in_HiddenFile;

         // C:\DEV --> G_HOMEDIR ��ȯ @ 2016.07.26 LSH
         sLocalFile  := G_HOMEDIR + 'TEMP\SPOOL\' + in_PhotoFile;



         if (GetBINFTP(FTP_SVRIP,
                       FTP_USERID,
                       FTP_PASSWD,
                       sRemoteFile,
                       sLocalFile,
                       False)) then
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

            Result := sLocalFile;
      end;
   end
   else
   begin
      // ����� ������ �̹��� ���ϸ�
      if in_UserId = 'XXXXX' then
      begin
         Result := '';
      end
      else
      begin
         // ����� HRM ������ �̹��� ���ϸ�
         ServerFileName := in_UserId + '.jpg';

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
            FTP_USERID := '';
            FTP_PASSWD := '';
            FTP_DIR    := '/kuhrm/app/hrm/photo/';


            // HRM Image �ٿ�ε�
            if Not GetBINFTP(FTP_SVRIP,
                             FTP_USERID,
                             FTP_PASSWD,
                             FTP_DIR + ServerFileName,
                             ClientFileName,
                             False) then
            begin
               //Showmessage('�̹��� �ٿ�ε忡 �����߽��ϴ�. �ٽ��ѹ� �õ��� �ֽʽÿ�');

               TUXFTP := nil;

               Exit;
            end;


            // �ش� Image Variant�� ����.
            AppendVariant(DownFile, ClientFileName);


            // �ٿ�ε� Ƚ�� üũ
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

      //[Ctrl + C] : ������ Cell Ŭ������� ����
      Ord('L') : if (ssCtrl in Shift) then
                 begin
                     //-----------------------------------------------------
                     // ��ϵ� User ���̾� Chat ����Ʈ �˾� @ 2015.03.30 LSH
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
   // ��ϵ� User ���̾� Chat ����Ʈ �˾� @ 2015.03.30 LSH
   //-----------------------------------------------------
   {  -- Chat XE7 ���������� �ּ� @ 2017.10.31 LSH
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
   // ��ϵ� User ���̾� Chat ����Ʈ �˾� @ 2015.03.30 LSH
   //-----------------------------------------------------
   {  -- Chat XE7 ���������� �ּ� @ 2017.10.31 LSH
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
                 PChar('[��󿬶���] - [���������]�� ����� IP ����� ����Ͻ� �� �ֽ��ϴ�.'),
                 PChar(Self.Caption + ' : ��ϵ��� ���� User ���� �˸�'),
                 MB_ICONWARNING + MB_OK);
   }
end;




//------------------------------------------------------------------------------
// String���� Ư�� ���ڿ��� �ٸ� ���ڿ��� ��ü
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
      // ���� Cell ��ǥ ��������
      asg_Like.MouseToCell(X, Y, NowCol, NowRow);

      if (NowCol <= asg_Like.ColCount-1) then
      begin

         with asg_Like do
         begin
            ShowHint := True;

            // ��õ�� Name Hint ǥ��
            Hint := Cells[NowCol, NowRow + 1]
         end;
      end
      else
         ShowHint := False;

   except
      // Grid.ColCount ����� Mousemoving�� out of index �˾� ����ó�� ����.
      asg_Like.Hint     := '';
      asg_Like.ShowHint := False;
   end;

end;

//------------------------------------------------------------------------------
// [��������] - [������ ���������] - [���� ������ ��ũ��Ʈ �ڵ�����] (Memo to Grid)
//------------------------------------------------------------------------------
procedure TMainDlg.mi_RlzScriptClick(Sender: TObject);
var
   sGijunDt, sNowTime : String;
   SearchRow : Integer;
   i, j, k, l, iMaxCheck, iStrTokenCnt : Integer;

   //--------------------------------------------------------------------
   // String���� Ư�� Char�� � �ִ��� Count �ϴ� �Լ�
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

   // �޸� Log �ʱ�ȭ
   mm_Log.Lines.Clear;


   with asg_Release do
   begin
      // ù��° Row�� ������ڸ� ���� Date�� fetch
      if (Cells[C_RL_REGDATE, Row] <> '') then
      begin
         //sGijunDt    := CopyByte(Cells[C_RL_REGDATE, 1], 1, 10);

         // ������ ��ũ��Ʈ ���� ������ üũ �б� @ 2016.11.02 LSH (������ �������, �ݿ��Ϻ���)
         if FormatDateTime('ddd', Date) = '��' then
            sGijunDt := FormatDateTime('yyyy-mm-dd', Date - 3)
         else
            sGijunDt := FormatDateTime('yyyy-mm-dd', Date - 2);

         // �� ������ ���� �ð� Set @ 2016.11.16 LSH
         sNowTime := FormatDateTime('yyyy-mm-dd hh:nn', Now);

         SearchRow   := Row;
      end
      else
      begin
         MessageBox(self.Handle,
                    '��ũ��Ʈ ���� ��� ������ ������ ã�� �� �����ϴ�.',
                    '������ ��ũ��Ʈ ���� ���� �˸�',
                    MB_OK + MB_ICONERROR);

         Exit;
      end;

      // �����Ͽ� �ش��ϴ� ����(����) ������ ������ Memo������ ����
      while CopyByte(Cells[C_RL_REGDATE, SearchRow], 1, 10) >= sGijunDt do
      begin
         // ���� ������ ���� ��������, Looping
         if (Trim(Cells[C_RL_SERVESRC, SearchRow]) <> '') and
            (Trim(Cells[C_RL_RELEASDT, SearchRow]) =  '') and           // ������ �Ͻ� ���� ��� fetch @ 2016.11.16 LSH
            (Trim(Cells[C_RL_REGDATE,  SearchRow]) < sNowTime) then     // �� ������ ���ؽð� ������ �ۼ��� ��� fetch @ 2016.11.16 LSH
         begin
            if PosByte('�ű�', Cells[C_RL_SERVESRC,  SearchRow]) > 0 then
            begin
               mm_Log.Lines.Add(Lowercase(ReplaceChar(ReplaceChar(ReplaceChar(ReplaceChar(DelChar(DelChar(DelChar(DelChar(DelChar(Cells[C_RL_SERVESRC,  SearchRow], '-'), '<'), '>'), '('), ')'), '�ű�', ''), ' ' , ''),#13, ''), #10, '|')));
               mm_Log.Lines.Add(Lowercase(ReplaceChar(ReplaceChar(ReplaceChar(ReplaceChar(DelChar(DelChar(DelChar(DelChar(DelChar(Cells[C_RL_SERVER,    SearchRow], '-'), '<'), '>'), '('), ')'), '�ű�', ''), ' ' , ''),#13, ''), #10, '.def|')) + '.def');
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



   // Looping Index �ʱ�ȭ
   iMaxCheck := 2;
   i         := 1;            // �޸��� Line - index
   k         := 0;            // Token �߰� Line - index
   l         := 1;            // �� Grid Line - index


   while (i <= iMaxCheck) do
   begin

      asg_Memo.RowCount := iMaxCheck + k;

      if Trim(mm_Log.Lines.Strings[i-1]) <> '' then
      begin
         // �ϳ��� Line�� �پ��ִ� Token ���� Ȯ��
         iStrTokenCnt := CountChar(mm_Log.Lines.Strings[i-1], '|');

         // ���� �ϳ��� String ���γ� Token(|)�� 1���̻� ������, Token ����ŭ �ٽ� �и��ؼ� Row ������ Line �߰�
         if iStrTokenCnt > 0 then
         begin
            // ���� Token�� ��ŭ�� ������ ������ row ������ assign
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

   apn_Memo.Caption.Text := '������ (SVC/CLT) ���';

   apn_Memo.Top     := 119;
   apn_Memo.Left    := 143;
   apn_Memo.Visible := True;
   apn_Memo.Collaps := True;
   apn_Memo.Collaps := False;

   asg_MemoClickCell(sender, 0, 0);

end;


//------------------------------------------------------------------------------
// ������ ��ũ��Ʈ ���� (�޸� Grid ����Ŭ����)
//------------------------------------------------------------------------------
procedure TMainDlg.asg_MemoDblClick(Sender: TObject);
var
   sRlzGroup   : String;
   sGroupPath  : String;
   sCLTSrcPath : String;

   i, iRlzCLTCnt : Integer;
begin
   // �޸� Log �ʱ�ȭ
   mm_Log.Lines.Clear;


   // SMS ���ø� ��� ���� �߰� @ 2017.07.11 LSH
   if PosByte('SMS', apn_Memo.Caption.Text) > 0 then
   begin
      // SMS �Է� Memo ���� ������ Template ���� ����
      fm_Sms.Lines.Clear;
      fm_Sms.Lines.Text := asg_Memo.Cells[1, asg_Memo.Row];

      apn_Memo.Visible := False;
      apn_Memo.Collaps := True;
   end
   else
   begin
      // CLT ������ ��� Count �ʱ�ȭ
      iRlzCLTCnt := 0;

      // CRLF ������, �� Line�� ������ ����� �ϳ��� Line-Group���� ����
      for i := 1 to asg_Memo.RowCount do
      begin
         // �� Row ����(���� �Ǵ� ����)�� ������ ��ɾ� ���� ����
         if (asg_Memo.Cells[0, i-1] <> '') and
            (CopyByte(Trim(asg_Memo.Cells[0, i-1]), 7, 1) <> 'L') and     // BPL (Package) ����� ���� (��: MDN110L1_P01) @ 2016.12.16 LSH
            (PosByte('_', asg_Memo.Cells[0, i-1]) > 0) then
         begin
            // ���� Row�� �׷��(ex. md / mh / mt ..)�� �� Row�� �׷��� ��������
            // ��� �ϳ��� Line���� ����
            if (CopyByte(asg_Memo.Cells[0, i], 1, 2) = CopyByte(asg_Memo.Cells[0, i-1], 1, 2)) then
            begin
               sRlzGroup   := asg_Memo.Cells[0, i-1] + ' ' + sRlzGroup;
               sGroupPath  := CopyByte(asg_Memo.Cells[0, i-1], 1, 2);
            end
            else
            // ���� Row�� �׷��(ex. md / mh / mt ..)�� �� Row�� �׷��� �ٸ�����
            // �ٷ� ���� Row(i-1)���� �߰��Ͽ� ������ ��ɾ� Line ���� ����
            begin
               sRlzGroup   := asg_Memo.Cells[0, i-1] + ' ' + sRlzGroup;
               sGroupPath  := CopyByte(asg_Memo.Cells[0, i-1], 1, 2);

               mm_Log.Lines.Add('');
               mm_Log.Lines.Add('---- [�׽�Ʈ��] cp from upsrc to SrcPath ----');
               mm_Log.Lines.Add('cp ' + sRlzGroup + ' ../src/' + sGroupPath);
               mm_Log.Lines.Add('---- [�׽�Ʈ��->�������] rocscp from users to user_release ----');
               mm_Log.Lines.Add('rocscp ' + sRlzGroup);

               // �ٸ� �׷� ��ɾ� Line �������� �ʱ�ȭ
               sRlzGroup   := '';
               sGroupPath  := '';
            end;
         end
         // Package (BPL) --> ������ �ڵ����(Add To Project) ���� @ 2016.11.02 LSH
         else if (asg_Memo.Cells[0, i-1] <> '') and
                 (
                     (CopyByte(Trim(asg_Memo.Cells[0, i-1]), 7, 1) = 'L') or     // BPL
                     (PosByte('DLL', Trim(asg_Memo.Cells[0, i-1])) > 0)          // DLL
                 ) then
         begin
            // BPL ���� Source ��� Set.
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
            // DLL ���� Source ��� Set.
            else if (PosByte('DLL', Trim(asg_Memo.Cells[0, i-1])) > 0) then
            begin
               if (PosByte('MDINFDLL', Trim(asg_Memo.Cells[0, i-1])) > 0) then
                  sCLTSrcPath := 'MDDLL\��ȣ����������\'
               else
                  sCLTSrcPath := 'MDDLL\' + Trim(asg_Memo.Cells[0, i-1]) + '\';
            end;


            // �� ����� .dproj�� ã�Ƽ� �����԰� ���ÿ� Delphi XE7 Project Manager�� �ڵ����
            ShellExecute(HANDLE, 'open',
                         PCHAR(Trim(asg_Memo.Cells[0, i-1]) + '.dproj'),
                         PCHAR(''),
                         PCHAR('C:\KUMC_DEV\Src\SrcM\' + sCLTSrcPath),
                         SW_NORMAL);

            // �߰��� CLT Count ++
            Inc(iRlzCLTCnt);

            // Ȯ�� Message
            MessageBox(self.Handle,
                       PChar(IntToStr(iRlzCLTCnt) + '��° BPL : ' + Trim(asg_Memo.Cells[0, i-1]) + ' --> Project Manager ��ϿϷ�.'),
                       '������ ��� Package �ڵ��߰�',
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
// �������� - �ٹ� �� ������Ȳ - �ٹ�ǥ ����(������)
//
// Author : Lee, Se-Ha
// Date   : 2015.06.12
//------------------------------------------------------------------------------
procedure TMainDlg.fcb_DutyMakeClick(Sender: TObject);
begin
   // �ٹ�ǥ ���� Check��
   if (fcb_DutyMake.Checked = True) then
   begin
      if GetMsgInfo('INFORM',
                    'DUTYLIST') = FsUserNm then
         SelGridInfo('DUTYMAKE');
   end
   else
   // �ٹ�ǥ ���� �� Check��, ���� Ȯ���̷�(MKUMCDTY) ��ȸ
   begin
      SelGridInfo('DUTYSEARCH');
   end;

   // �ٹ�ǥ Ȯ��(����) Enabled ����
   fsbt_DutyInsert.Enabled := fcb_DutyMake.Checked;

end;

//------------------------------------------------------------------------------
// �����ٹ�ǥ Ȯ��/����
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

   // �����ٹ�ǥ Ȯ��/����
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

                         { -- Ȯ�强 ����ؼ� ���ܵ�..
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
                    PChar(IntToStr(iCnt) +'���� �����ٹ��̷��� ���������� [���] �Ǿ����ϴ�.'),
                    '[KUMC ���̾�α�] �����ٹ�ǥ ������Ʈ �˸� ',
                    MB_OK + MB_ICONINFORMATION);

         // �����ٹ�ǥ ����(������) Disabled
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
// [�ٹ� �� ����] TAdvStringGrid onCanEditCell Event Hander
//
// Author : Lee, Se-Ha
// Date   : 2015.06.12
//------------------------------------------------------------------------------
procedure TMainDlg.asg_DutyCanEditCell(Sender: TObject; ARow,
  ACol: Integer; var CanEdit: Boolean);
begin
   // ������, ���»��׸� ��������
   CanEdit := ACol in [C_DT_DUTYUSER, C_DT_REMARK];
end;

//------------------------------------------------------------------------------
// [�ٹ� �� ����] TAdvStringGrid onKeyDown Event Hander
//
// Author : Lee, Se-Ha
// Date   : 2015.06.12
//------------------------------------------------------------------------------
procedure TMainDlg.asg_DutyKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   with asg_Duty do
   begin
      // ������ �÷� ������� �߻���(Enter Ű) Ȯ�� Msg
      if (Col = C_DT_DUTYUSER) and
         (Cells[C_DT_DUTYUSER, Row] <> Cells[C_DT_ORGDTYUSR, Row]) and
         (Key = 13) then
      begin
         if MessageBox(self.Handle,
                       PChar('���� �ٹ��� [' + Cells[C_DT_ORGDTYUSR, Row] + ' -> ' + Cells[C_DT_DUTYUSER, Row] + '](��)�� �����Ͻðڽ��ϱ�?'),
                       PChar(Cells[C_DT_DUTYDATE,  Row] + ' (' + Cells[C_DT_YOIL,  Row] + ') �ٹ��� ������ Ȯ��'),
                       MB_YESNO + MB_ICONQUESTION) = ID_YES then
         begin
            UpdateDuty('DUTYUSER');
         end
         else
            Cells[C_DT_DUTYUSER, Row] := Cells[C_DT_ORGDTYUSR, Row];

      end
      // ����(�ް�) Ư����� ���� �߻���(EnterŰ) Ȯ�� Msg
      else if (Col = C_DT_REMARK) and
              (Trim(Cells[C_DT_REMARK, Row]) <> Trim(Cells[C_DT_ORGREMARK, Row])) and
              (Key = 13) then
      begin
         if MessageBox(self.Handle,
                       PChar('����(�ް�) �� Ư����� ��������� �����Ͻðڽ��ϱ�?'),
                       PChar(Cells[C_DT_DUTYDATE,  Row] + ' (' + Cells[C_DT_YOIL,  Row] + ') ����(�ް�) �� Ư����� ������ Ȯ��'),
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
// [�ٹ� �� ����] �ٹ�-����-�������� ������� Update
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
   // �����ٹ��� �� ����(�ް�) Ư����� ����
   //---------------------------------------------------------------
   if (in_UpdFlag = 'DUTYUSER') then
      sType := '17'
   //---------------------------------------------------------------
   // �����ٹ� Seq �� ����/���� �ٹ������� ����
   //---------------------------------------------------------------
   else if (in_UpdFlag = 'DUTYSEQ') then
      sType := '18'
   //---------------------------------------------------------------
   // �������� �ۼ�(���)
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
            AppendVariant(vDutyDate ,   asg_DutyList.Cells[2, i]    );    // ���ϱٹ� ���� Date
            AppendVariant(vDutyFlag ,   asg_DutyList.Cells[3, i]    );    // ���ϱٹ� ���� Date
            AppendVariant(vDelDate  ,   ''               );
            AppendVariant(vEditId   ,   ''               );
            AppendVariant(vEditIp   ,   FsUserIp         );

            Inc(iCnt);
         end;
      end
      else if (in_UpdFlag = 'DUTYRPT') then
      begin
         // ���࿹��
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

                         { -- Ȯ�强 ��� ���ܵ�..
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
                    PChar(IntToStr(iCnt) +'���� �����ٹ� ��������� ���������� [������Ʈ] �Ǿ����ϴ�.'),
                    '[KUMC ���̾�α�] �����ٹ� ���� ������Ʈ �˸� ',
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




// �������� �ۼ�ȭ�� �ݱ�
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
                 PChar('[��󿬶���] - [���������]�� ����� IP ����� ����Ͻ� �� �ֽ��ϴ�.'),
                 PChar(Self.Caption + ' : ��ϵ��� ���� User ���� �˸�'),
                 MB_OK + MB_ICONWARNING);

      Exit;
   end;


   fpn_DutyRpt.Top     := 0;
   fpn_DutyRpt.Left    := 0;
   fpn_DutyRpt.Visible := True;

   fpn_DutyRpt.BringToFront;


   with asg_Duty do
   begin
      fed_DutyDate.Text := CopyByte(Cells[C_DT_DUTYDATE, Row], 1, 4) + '�� ' +
                           CopyByte(Cells[C_DT_DUTYDATE, Row], 6, 2) + '�� ' +
                           CopyByte(Cells[C_DT_DUTYDATE, Row], 9, 2) + '�� ';

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
   fmm_Text.Hint      := '�ű� �ۼ����� �ѱ۱��� 15��(500�� ����)�̳��� �����մϴ�.';
   }

   lb_DutyRpt.Caption    := '�� �������� �ۼ�(�ű�/����) ����Դϴ�.';
   lb_DutyRpt.Font.Color := $006C8518;
   lb_DutyRpt.Font.Style := [fsBold];

end;

//------------------------------------------------------------------------------
// [��������] - [������ ���������] - [BPL --> ������ �ڵ��߰�]
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
   // String���� Ư�� Char�� � �ִ��� Count �ϴ� �Լ�
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
   // �޸� Log �ʱ�ȭ
   mm_Log.Lines.Clear;

   // ù��° Row�� ������ڸ� ���� Date�� fetch
   if (asg_Release.Cells[C_RL_REGDATE, asg_Release.Row] <> '') then
   begin
      //sGijunDt    := CopyByte(Cells[C_RL_REGDATE, 1], 1, 10);

      // ������ ��ũ��Ʈ ���� ������ üũ �б� (������ �������, �ݿ��Ϻ���)
      if FormatDateTime('ddd', Date) = '��' then
         sGijunDt := FormatDateTime('yyyy-mm-dd', Date - 3)
      else
         sGijunDt := FormatDateTime('yyyy-mm-dd', Date - 2);

      // �� ������ ���� �ð� Set @ 2016.11.16 LSH
      sNowTime := FormatDateTime('yyyy-mm-dd hh:nn', Now);

      SearchRow   := asg_Release.Row;
   end
   else
   begin
      MessageBox(self.Handle,
                 'Package(BPL) ������ ��� ������ ã�� �� �����ϴ�.',
                 'Package (BPL) ������ ��� ���� �˸�',
                 MB_OK + MB_ICONERROR);

      Exit;
   end;

   // �����Ͽ� �ش��ϴ� Client(BPL) ������ ������ ã�Ƽ�, ���� Source ����� �ش� [������Ʈ��.dproj]�� ShellExecute �Ͽ� add to Project ����.
   while CopyByte(asg_Release.Cells[C_RL_REGDATE, SearchRow], 1, 10) >= sGijunDt do
   begin
      if (Trim(asg_Release.Cells[C_RL_CLIENT,   SearchRow]) <> '') and           // CLT ������ ���� ��������, Looping
         (Trim(asg_Release.Cells[C_RL_RELEASDT, SearchRow]) =  '') and           // ������ �Ͻ� ���� ��� fetch @ 2016.11.16 LSH
         (Trim(asg_Release.Cells[C_RL_REGDATE,  SearchRow]) < sNowTime) then     // �� ������ ���ؽð� ������ �ۼ��� ��� fetch @ 2016.11.16 LSH
      begin
         if PosByte('�ű�', asg_Release.Cells[C_RL_CLIENT,  SearchRow]) > 0 then
         begin
            mm_Log.Lines.Add(Lowercase(ReplaceChar(ReplaceChar(ReplaceChar(ReplaceChar(DelChar(DelChar(DelChar(DelChar(DelChar(asg_Release.Cells[C_RL_CLIENT, SearchRow], '-'), '<'), '>'), '('), ')'), '�ű�', ''), ' ' , ''),#13, ''), #10, '|')));
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



   // Looping Index �ʱ�ȭ
   iMaxCheck := 2;
   i         := 1;            // �޸��� Line - index
   k         := 0;            // Token �߰� Line - index
   l         := 1;            // �� Grid Line - index


   while (i <= iMaxCheck) do
   begin

      asg_Memo.RowCount := iMaxCheck + k;

      if Trim(mm_Log.Lines.Strings[i-1]) <> '' then
      begin
         // �ϳ��� Line�� �پ��ִ� Token ���� Ȯ��
         iStrTokenCnt := CountChar(mm_Log.Lines.Strings[i-1], '|');

         // ���� �ϳ��� String ���γ� Token(|)�� 1���̻� ������, Token ����ŭ �ٽ� �и��ؼ� Row ������ Line �߰�
         if iStrTokenCnt > 0 then
         begin
            // ���� Token�� ��ŭ�� ������ ������ row ������ assign
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
// [��������] - [������ ���������] - [�˻��Ⱓ ������] onKeyDown Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2016.11.18
//------------------------------------------------------------------------------
procedure TMainDlg.fmed_RegFrDtKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   // ������ ���� ������ ���� Comp.
   if (Key = VK_RETURN) then
   begin
      // ������ ���� ��ȸ
      if IsLogonUser then
      begin
         if fcb_ReScanRelease.Checked then
            fcb_ReScanRelease.Checked := False;

         // �˻��� Clear --> �˻��� �����ϴ� ���, �Ⱓ�� �����ؼ� �˻� ���� @ 2017.10.26 LSH
         //fed_Release.Text := '';

         SelGridInfo('RELEASESCAN');

         fcb_ReScanRelease.Checked := True;

         fed_Release.SetFocus;
      end;
   end;
end;


//------------------------------------------------------------------------------
// [��������] - [������ ���������] - [�˻��Ⱓ ������] onKeyDown Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2016.11.18
//------------------------------------------------------------------------------
procedure TMainDlg.fmed_RegToDtKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   // ������ ���� ������ ���� Comp.
   if (Key = VK_RETURN) then
   begin
      // ������ ���� ��ȸ
      if IsLogonUser then
      begin
         if fcb_ReScanRelease.Checked then
            fcb_ReScanRelease.Checked := False;

         // �˻��� Clear --> �˻��� �����ϴ� ���, �Ⱓ�� �����ؼ� �˻� ���� @ 2017.10.26 LSH
         //fed_Release.Text := '';

         SelGridInfo('RELEASESCAN');

         fcb_ReScanRelease.Checked := True;

         fed_Release.SetFocus;
      end;
   end;
end;

//------------------------------------------------------------------------------
// [�����м�] - [S/R��ȸ] - [�˻��Ⱓ ������] onKeyDown Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2016.11.18
//------------------------------------------------------------------------------
procedure TMainDlg.fmed_AnalFromKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   // S/R ������ ���� Comp.
   if (Key = VK_RETURN) then
   begin
      // S/R �̷� ��ȸ
      if IsLogonUser then
      begin
         lb_Analysis.Caption := '�� ��ȸ���Դϴ�....';

         if fcb_ReScan.Checked then
            fcb_ReScan.Checked := False;

         SelGridInfo('ANALYSIS')
      end;
   end;
end;

//------------------------------------------------------------------------------
// [�����м�] - [S/R��ȸ] - [�˻��Ⱓ ������] onKeyDown Event Handler
//
// Author : Lee, Se-Ha
// Date   : 2016.11.18
//------------------------------------------------------------------------------
procedure TMainDlg.fmed_AnalToKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
   // S/R ������ ���� Comp.
   if (Key = VK_RETURN) then
   begin
      // S/R �̷� ��ȸ
      if IsLogonUser then
      begin
         lb_Analysis.Caption := '�� ��ȸ���Դϴ�....';

         if fcb_ReScan.Checked then
            fcb_ReScan.Checked := False;

         SelGridInfo('ANALYSIS')
      end;
   end;
end;


//------------------------------------------------------------------------------
// SMS �ڵ� ���ø� ��� ��ȸ
//          - ������ ��ó�� ��� �ϻ��� ������ �ڵ�ȭ ���ѹ�����..
//          - ���� ��ȭ���� FAQ ���ǵ��� D/B�� �־�ְ�, ���� Photos ��� �����Ͽ�
//            SMS ���ø����� ��������.
//
// Author : Lee, Se-Ha
// Date   : 2017.07.11
//------------------------------------------------------------------------------
procedure TMainDlg.mi_Sms_TemplateClick(Sender: TObject);
begin
   apn_Memo.Caption.Text := 'SMS ���ø� ���';

   apn_Memo.Top     := apn_SMS.Top;
   apn_Memo.Left    := apn_SMS.Left + apn_SMS.Width + 1;
   apn_Memo.Visible := True;
   apn_Memo.Collaps := True;
   apn_Memo.Collaps := False;

   asg_Memo.ClearRows(1, asg_Memo.RowCount);
   asg_Memo.RowCount := 2;

   // SMS ���ø� ��� D/B ��ȸ
   GetSMSTemplate;

end;

//--------------------------------------------------------------------------
// [SMS] SMS ���ø� ��� D/B ��ȸ
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
      // 2. SMS ���ø� ��� List ǥ��
      //----------------------------------------------------
      with asg_Memo do
      begin

         ColWidths[0] := 100;
         ColWidths[1] := 265;

         RowCount := TpGetSMSTemp.RowCount + FixedRows + 1;

         for iTempIdx := 0 to TpGetSMSTemp.RowCount - 1 do
         begin
            // ����
            Cells[0, iTempIdx+FixedRows] := TpGetSMSTemp.GetOutputDataS('sComcdnm2', iTempIdx);        // ���ø� ����
            Cells[1, iTempIdx+FixedRows] := TpGetSMSTemp.GetOutputDataS('sComcdnm3', iTempIdx);        // SMS ���ø� Context
         end;

         // �� �� Hidden Info ó��
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
// ���߰��� ����(���������) �̷� ���� (Deldate Updated)
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

   // ���õǾ� �ִ� �� �ϰ� Update
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
                    PChar(IntToStr(iCnt) +'���� ��������� ���������� [����] �Ǿ����ϴ�.'),
                    '[KUMC ���̾�α�] ���߹���(���������) ������Ʈ �˸� ',
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






