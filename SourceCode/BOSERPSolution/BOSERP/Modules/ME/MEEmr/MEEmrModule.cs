#region using
using AutoMapper;
using BOSCommon;
using BOSComponent;
using BOSERP.Modules.ME;
using BOSERP.Modules.ME.Helpers;
using BOSERP.Modules.MEEmr.UI;
using BOSERP.Utilities;
using BOSLib;
using BOSLib.DataAccess;
using Clas.Business.EmrStore;
using Clas.Business.Ftp;
using Clas.Emr.Core;
using Clas.Emr.Intergration;
using Clas.Emr.Ipc.Net.Messaging;
using Clas.Emr.Model;
using Clas.Model.Domain;
using Clas.Model.Mongo;
using DevExpress.Skins;
using DevExpress.XtraBars;
using DevExpress.XtraBars.Docking;
using DevExpress.XtraBars.Ribbon;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Filtering.Templates;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraPdfViewer;
using DevExpress.XtraPrinting;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Layout;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit.Commands;
using DevExpress.XtraRichEdit.Services;
using DevExpress.XtraTreeList;
using DevExpress.XtraTreeList.Nodes;
using Emr;
using Emr.Ca;
using Emr.Ca.Core;
using Emr.Devices.GeV100;
using Emr.Document.Pdf;
using Emr.Pluggable.Interface;
using Emr.Pluggable.Plugin;
using Emr.Workflow.Client;
using Emr.Workflow.Client.Models;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Localization;
using MongoDB.Bson;
using MongoDB.Bson.Serialization;
using MongoDB.Bson.Serialization.Serializers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Activities;
using System.Activities.XamlIntegration;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Path = System.IO.Path;
using Clas.Model.Middle;
using Emr.FingerPrint;
using System.Xml.Linq;
using signotec.STPadLibNet;
using Emr.SignPad;
using System.Security;
using Microsoft.AspNet.SignalR.Client;
using System.Reflection;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using System.Runtime.InteropServices;
using BOSBase.Dto;
using DevExpress.DataProcessing;
using DevExpress.XtraCharts;
using Gecko.Net;
using System.Threading;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
#endregion

namespace BOSERP.Modules.MEEmr
{
    public class MEEmrModule : BaseModuleERP
    {
        #region Constant
        private const string _templateContentCtrlName = "richEditCtrl";
        private const string _fld_lbc_MEEmrsName = "fld_lbc_MEEmrs";
        private const string _fld_lbl_MEPatientNoName = "fld_lbl_MEPatientNo";
        private const string _fld_lbl_MEPatientNameName = "fld_lbl_MEPatientName";
        private const string _fld_lbl_MEEmrDepartment = "fld_lbl_MEEmrDepartment";
        private const string _fld_lbl_MEPatientBirthdayName = "fld_lbl_MEPatientBirthday";

        private const string _fld_lbl_RichEditMsgName = "fld_lbl_RichEditMsg";
        private const string _fld_dgc_MEEmrsName = "fld_dgcMEEmr";
        private const string _fld_lkeFK_MEPatientID1Name = "fld_lkeFK_MEPatientID1";

        private MEPatientsController _patientCtrl;
        private MEEmrsController _emrCtrl;
        private MEEmrEntities _entity;
        private METemplatesController _templateCtrl;
        private string _documentPath;
        private RichEditControl _richEditCtrl;
        private RichEditControl _tempRichEditCtrl;
        private MEEmrDocumentsController _emrDocumentCtrl;
        private MEEmrDocumentSignsController _emrDocumentSignCtrl;
        private MEEmrActionsController _actionsController;
        private string _receiveChannel;
        private FileTemplateManager _ftpFileMng;

        private EmrDocumentManager _emrDocumentMng;
        private string _msgMaximumPage;
        private Dictionary<string, string> _userAbbrevs;
        private MEEmrTypesController _emrTypeCtrl;
        private MEEmrTypeActionsController _emrTypeActionCtrl;
        private MEEmrSymbolsController _symbolCtrl;
        private bool _overrideEditPermission = true;
        private bool _moveBetweenTagByTabEnabled = true;
        private bool _clickAndAutoMoveToNextTagConfig = false;
        private bool _allowEditOuterEmrTag = false;
        private bool _allowNormalUserHideDocument = false;
        private bool _notAllowCopyEmrTag = false;
        private bool _highlightModeEmrTagBorder = false;
        private bool _highlightModeEmrTagFill = false;
        private bool _highlightEmrTagUserConfig = true;
        private bool _pagePainterVer2 = false;
        private string _pagePrintPainterVer = string.Empty;
        private object _appPreData;
        private HashProvider _md5Hasher;
        private bool _useZKTeco = Convert.ToBoolean(BOSApp.GetSystemConfigValue(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.FINGER_PRINT_USE_ZKTeco));
        private bool _showPopupUpdateInfoPatient = Convert.ToBoolean(BOSApp.GetSystemConfigValue(EmrProcess.GROUP, SysCfgConsts.SHOW_POPUP_UPDATE_INFO_PATIENT));

        private bool _useSignPadSignotec = Convert.ToBoolean(BOSApp.GetSystemConfigValue(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.SIGN_PAD_USE_SIGNOTEC));
        private bool _kydientuHashWaterMark = Convert.ToBoolean(BOSApp.GetSystemConfigValue(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.KYDIENTU_PRINT_HASH_WATERMARK));
        private bool _appToAppVB = Convert.ToBoolean(BOSApp.GetSystemConfigValue(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.APP_TO_APP_VB));

        private bool _checkSystem = false;
        private string _hospitalProject = BOSApp.GetSystemConfigValue(SysCfgConsts.PRIVATE, SysCfgConsts.HOSPITAL_PROJECT);
        #endregion

        #region Variable
        private IpcListener _hisChannelListener;
        private ParamPool _requestParamPool;

        private ApiHelper _api;
        private SqlHelper _sqlHelper;
        private BOSSearchResultsGridControl _searchGrid;
        private LabelControl _msgNotification;
        private Dictionary<string, string> _patientDataMappings;
        private MEEmrTemplateActionsController _templateActionCtrl;
        private string _signedContent;
        private MEEmrActionRelationsController _actionRelationsController;
        private HREmployeesController _employeeCtrl;
        private ADUserConfigsController _userConfigCtrl;
        private EmrAction _emrActionHelper;

        private EmrDocumentHelper _emrDocumentHelper;
        private DataAccess _dataHelper;
        private EmrParser _emrParser;
        private BOSMemoEdit _msgLogs;
        private DockPanel _dpnMsg;
        private bool _notificationTab;
        private bool _logConfig;
        private string _msgLogsTemp;

        private PdfViewer _pdfViewer;
        private DockPanel _dpnRichEdit;
        private DockPanel _dpnViewPdf;

        private MEParamLookupDatasController _paramLookupDataCtrl;
        private MEParamLookupsController _paramLookupCtrl;
        private MEEmrImagesController _imageCtrl;
        private MEEmrImagePatternsController _imagePatternCtrl;
        private MEEmrImageParamsController _imageParamCtrl;
        private HRDepartmentsController _departmentCtrl;

        private BOSDbUtil _dbUtil;
        private GeV100SerialConn _v100SerialPort;

        private EmrChartHelper _chartHelper;
        private List<DevExpress.Office.Utils.OfficeImage> _signedImgs;
        private string _blockWriteUrl;
        private string _blockVerifyUrl;
        private HashProvider _hashProvider;
        private ApiHelper _apiBlock;
        private bool _useMessageBox;

        private MEEmrTransferHistoriesController _transferCtrl;
        private MEEmrShareHistoriesController _shareCtrl;
        private string _macAddress;
        private string _ipAddress;
        private string _hostName;
        private DigitalSignatureProvider _caProvider;
        private IDigitalSignatureBase _digitalSig;

        private IPdfProcessor _pdfProcessor;
        private EmrArchiveHelper _emrArchiveHelper;
        private METemplateParamsController _templateParamCtrl;
        private METemplateIndexsController _templateIndexsCtrl;
        private EmrDocumentSortHelper _emrDocumentSortHelper;

        private ApiHelper _apiEmr;
        private bool _isRichEditInputingMode = false;
        private List<System.Threading.Thread> _backgroundJobThreads;
        private readonly bool _fingerPrintHashWatermark;
        private readonly ReaderHandler _fingerPrint;
        private ReaderHandlerZK _fingerPrintZK;
        private string _fileLog;
        private DevExpress.XtraBars.BarEditItem _qatNotificationMsg;

        private bool _sysnSave = true;

        private List<MEEmrMergeHistoriesInfo> _mergeHistories;
        private MainHelper _helper;
        private SystemHelper _sysHelper;
        private int _stateAppHISKV = 0;

        private string _dataInitFromHIS;
        private IModel _chanelExchangeHisToEmr;
        private IModel _chanelQueueHisToEmr;
        private bool _waittingForHIS;
        private string _dataJsonFromHIS;

        // 0: none, error
        // 1: init
        // 2: received data
        #endregion

        #region Public
        public DockManager DockManager { get; set; }
        public bool IsRichEditInputMode
        {
            get { return _isRichEditInputingMode; }
            set { _isRichEditInputingMode = value; }
        }
        #endregion
        public MEEmrModule()
        {
            Name = "MEEmr";
            CurrentModuleEntity = new MEEmrEntities();
            CurrentModuleEntity.Module = this;
            InitializeModule();
            _backgroundJobThreads = new List<System.Threading.Thread>();

            _fingerPrintHashWatermark = _entity.GetConfigFingerPrintHashWatermark();
            _fingerPrint = new ReaderHandler();
            _fingerPrintZK = new ReaderHandlerZK();
            this._helper = new MainHelper();
        }
        public override void InitializeModule()
        {
            var begin = DateTime.Now;
            _symbolCtrl = new MEEmrSymbolsController();
            base.InitializeModule();
            //InitializeModule();
            this._entity = CurrentModuleEntity as MEEmrEntities;
            this._patientCtrl = new MEPatientsController();
            this._emrCtrl = new MEEmrsController();
            this._emrTypeCtrl = new MEEmrTypesController();
            this._emrTypeActionCtrl = new MEEmrTypeActionsController();
            this._emrDocumentCtrl = new MEEmrDocumentsController();
            this._templateCtrl = new METemplatesController();
            this._actionsController = new MEEmrActionsController();
            this._actionRelationsController = new MEEmrActionRelationsController();
            this._templateActionCtrl = new MEEmrTemplateActionsController();
            this._employeeCtrl = new HREmployeesController();
            this._userConfigCtrl = new ADUserConfigsController();
            this._sysHelper = new SystemHelper();

            if (_sysHelper.AllowThread())
            {
                //uthv co the chay am tham ben duoi ma ko anh huong
                var thGetAllTemplate = new System.Threading.Thread(() => GetAllTemplate());
                thGetAllTemplate.Start();
            }
            else
            {
                GetAllTemplate();
            }

            Configuration configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            _checkSystem = SystemMemCache.GetSystemConfigValue(SysCfgConsts.PRIVATE, SysCfgConsts.PRIVATE_ALLOW_CHECK_SYSTEM).ToLower() == "true";
            this._documentPath = SystemMemCache.GetSystemConfigValue(SysCfgConsts.PRIVATE, SysCfgConsts.PRIVATE_TEMPLATE_SERVER_PATH);
            if (!_documentPath.Contains(":\\")) // đường dẫn tương đối
                this._documentPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\" + _documentPath;

            this._dpnViewPdf = this.Controls["dpnViewPdf"] as DockPanel;
            this._dpnRichEdit = this.Controls["dpnRichEdit"] as DockPanel;
            this._richEditCtrl = this.Controls[_templateContentCtrlName] as RichEditControl;
            this._pdfViewer = this.Controls["fld_pdfViewer"] as PdfViewer;
            if (this._richEditCtrl == null) MessageBox.Show("RichEditControl is null");

            this._searchGrid = this.Controls[_fld_dgc_MEEmrsName] as BOSSearchResultsGridControl;
            SetupNotificationControl();

            this._dpnMsg = this.Controls["dpnMsg"] as DockPanel;
            this._msgLogs = this.Controls["msgLogs"] as BOSMemoEdit;

            #region APP TO APP
            this._receiveChannel = SystemMemCache.GetSystemConfigValue(SysCfgConsts.PRIVATE, SysCfgConsts.PRIVATE_EMR_RECEIVE_CHANNEL);// configuration.AppSettings.Settings["emr_receive_channel"].Value.ToString();
            if (!string.IsNullOrEmpty(_receiveChannel))
            {
                // HIS FPT,...
                // create an instance of the listener object
                this._hisChannelListener = new IpcListener();
                // attach the message handler
                this._hisChannelListener.MessageReceived += new IpcListener.IpcMessageHandler(listener_MessageFromHisReceived);
                // register the channels we want to listen on
                this._hisChannelListener.RegisterChannel(_receiveChannel);
            }
            #endregion

            //this._api = new ApiHelper(SystemMemCache.GetSystemConfigValue(SysCfgConsts.PRIVATE, SysCfgConsts.PRIVATE_HIS_API_ENDPOINT), BOSApp.ApiToken);
            this._api = new ApiHelper(SqlDatabaseHelper._HIS_API_ENDPOINT, BOSApp.ApiToken);
            this._sqlHelper = new SqlHelper();

            this.InitPatientInfoMappingForExternal();
            this._tempRichEditCtrl = new RichEditControl();
            this._emrDocumentHelper = new EmrDocumentHelper(this._richEditCtrl);
            this._emrParser = new EmrParser(this._richEditCtrl, this._msgNotification, this._msgLogs, this._emrDocumentHelper);
            this._emrActionHelper = new EmrAction(this._richEditCtrl, this._emrDocumentHelper, this._emrParser);
            this._dataHelper = new DataAccess();
            this._ftpFileMng = new FileTemplateManager();

            this._richEditCtrl.Options.Authentication.UserName = BOSApp.CurrentUsersInfo.ADUserName;
            this._richEditCtrl.Options.Authentication.Password = this._emrDocumentHelper.ShareEmrPassword;

            var myCommandFactory = new CustomRichEditCommandFactoryService(this, this._richEditCtrl, this._richEditCtrl.GetService<IRichEditCommandFactoryService>());
            this._richEditCtrl.ReplaceService<IRichEditCommandFactoryService>(myCommandFactory);

            this._paramLookupDataCtrl = new MEParamLookupDatasController();
            this._paramLookupCtrl = new MEParamLookupsController();
            this._imageCtrl = new MEEmrImagesController();
            this._imagePatternCtrl = new MEEmrImagePatternsController();
            this._imageParamCtrl = new MEEmrImageParamsController();
            if (_sysHelper.AllowThread())
            {
                var thDocMongo = new System.Threading.Thread(() =>
                {
                    this._emrDocumentMng = new EmrDocumentManager();
                });
                thDocMongo.Start();
            }
            else
            {
                this._emrDocumentMng = new EmrDocumentManager();
            }
            this._emrDocumentSignCtrl = new MEEmrDocumentSignsController();
            this._departmentCtrl = new HRDepartmentsController();

            if (_sysHelper.AllowThread())
            {
                //uthv co the chay am tham ben duoi ma ko anh huong
                var thInitAutoCorrect = new System.Threading.Thread(() => InitAutoCorrect());
                thInitAutoCorrect.Start();
            }
            else
            {
                InitAutoCorrect();
            }
            this._richEditCtrl.AutoCorrect += richEditControl1_AutoCorrect;
            this._dbUtil = new BOSDbUtil();

            var tree = Controls["fld_trlDocumentDataView"] as DevExpress.XtraTreeList.TreeList;
            Skin skin = GridSkins.GetSkin(tree.LookAndFeel);
            skin.Properties[GridSkins.OptShowTreeLine] = true;

            var currentView = _richEditCtrl.ActiveView as PageBasedRichEditView;
            currentView.PageCountChanged += new EventHandler(_richEditCtrl_PageBasedRichEditView_PageCountChanged);

            _richEditCtrl.Document.DefaultCharacterProperties.FontName = "Times New Roman";
            _richEditCtrl.Document.DefaultCharacterProperties.FontSize = 13;

            _allowNormalUserHideDocument = _entity.GetConfigAllowNormalUserHideDocument();
            if (_allowNormalUserHideDocument)
            {
                Controls["fld_btn_HidingEmrDocument"].Visible = true;
                Controls["fld_btn_HidingPatientDocument"].Visible = true;
            }
            else
            {
                Controls["fld_btn_HidingEmrDocument"].Visible = (BOSApp.CurrentUserGroupInfo.ADUserGroupRole == UserGroupRole.admin.ToString());
                Controls["fld_btn_HidingPatientDocument"].Visible = (BOSApp.CurrentUserGroupInfo.ADUserGroupRole == UserGroupRole.admin.ToString());
            }

            _richEditCtrl.KeyDown += _richEditCtrl_KeyDown;

            //_richEditCtrl.UnhandledException += richEditCtrl_UnhandledException;
            //_richEditCtrl.DocumentLayout.DocumentFormatted += DocumentLayout_DocumentFormatted;

            _chartHelper = new EmrChartHelper();
            Plugin.UseCache = SystemMemCache.GetSystemConfigValue(SysCfgConsts.PRIVATE, SysCfgConsts.PRIVATE_CACHE_PLUGINS).ToLower() == "true";

            _useMessageBox = SystemMemCache.GetSystemConfigValue(SysCfgConsts.PRIVATE, SysCfgConsts.PRIVATE_USE_MESSAGE_BOX_FOR_ALERT).ToLower() == "true";

            _transferCtrl = new MEEmrTransferHistoriesController();
            _shareCtrl = new MEEmrShareHistoriesController();
            ConfigBlock(configuration);
            SetMachineInfo();
            ConfigDigitalSignature();
            BackgroundLoadCache();
            LoadUserUIConfig();

            _pdfProcessor = new PdfProcessor(_documentPath);
            _emrArchiveHelper = new EmrArchiveHelper(_documentPath, _pdfProcessor, _ftpFileMng, _hashProvider, _digitalSig);

            _templateParamCtrl = new METemplateParamsController();
            _templateIndexsCtrl = new METemplateIndexsController();
            _emrDocumentSortHelper = new EmrDocumentSortHelper();

            if (_sysHelper.AllowThread())
            {
                var thClearEmrFiles = new System.Threading.Thread(() => ClearEmrFiles());
                thClearEmrFiles.Start();
            }
            else
            {
                ClearEmrFiles();
            }

            //var emrEndpoint = BOSApp.GetSystemConfigValue(SysCfgConsts.PRIVATE, SysCfgConsts.PRIVATE_EMR_API_ENDPOINT);
            var emrEndpoint = SqlDatabaseHelper._EMR_API_ENDPOINT;
            if (!string.IsNullOrEmpty(emrEndpoint) && !string.IsNullOrEmpty(BOSApp.EmrApiAuthToken))
            {
                var timeout = BOSApp.GetSystemConfigValueInt(SysCfgConsts.PRIVATE, SysCfgConsts.PRIVATE_EMR_API_ENDPOINT_TIMEOUT, 180000);
                _apiEmr = new ApiHelper(emrEndpoint, BOSApp.EmrApiAuthToken, "EMR", timeout);
            }
            _allowEditOuterEmrTag = _entity.GetConfigAllowEditOuterEmrTag();

            _notAllowCopyEmrTag = _entity.GetConfigNotAllowCopyEmrTag();

            _highlightModeEmrTagBorder = _entity.GetConfigHighlightModeEmrTagBorder();
            _highlightModeEmrTagFill = _entity.GetConfigHighlightModeEmrTagFill();

            if (_highlightModeEmrTagBorder || _highlightModeEmrTagFill)
            {
                if (_highlightEmrTagUserConfig)
                    _richEditCtrl.Options.Fields.HighlightMode = FieldsHighlightMode.Always;
            }

            _richEditCtrl.ReadOnly = true;
            _pagePainterVer2 = _entity.GetConfigPagePainterVersion();
            _pagePrintPainterVer = _entity.GetConfigPagePrintPainterVersion();
            _richEditCtrl.BeforePagePaint += RichEditControl_BeforePagePaint;

            _fileLog = SanitizedFileName.Sanitize($"{BOSApp.CurrentUser}_{_macAddress}_{_ipAddress}_", "_") + DateTime.Now.ToString("yyyyMMdd_HHmmssfff") + ".txt";

            _logConfig = BOSApp.GetSystemConfigValue(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.SYSTEM_CONFIGS_GET_CLIENT_EMR_LOG)?.ToUpper() == "TRUE";

            _md5Hasher = new HashProvider("MD5");

            ShowFlashNotification("Khởi tạo module trong (ms): " + (int)(DateTime.Now - begin).TotalMilliseconds, 10000);
            AppMemCache.InitEmrModuleSessionAsync();
            _tempRichEditCtrl.BeforePagePaint += RichEditTempControl_BeforePagePaint;

            #region rabbitMQ
            if (BOSApp._isLoginByHIS)
            {
                InitExchangeHisToEmr();
                InitQueueHisToEmr();
            }
            #endregion
        }
        private void SetMachineInfo()
        {
            _macAddress = BOSApp.GetMachineMac();
            _ipAddress = BOSApp.GetMachineIp();
            _hostName = System.Net.Dns.GetHostName();
        }
        private void ConfigBlock(System.Configuration.Configuration configuration)
        {
            this._blockWriteUrl = SystemMemCache.GetSystemConfigValue(SysCfgConsts.PRIVATE, SysCfgConsts.PRIVATE_BLOCK_WRITE_URL);
            this._blockVerifyUrl = SystemMemCache.GetSystemConfigValue(SysCfgConsts.PRIVATE, SysCfgConsts.PRIVATE_BLOCK_VERIFY_URL);

            _hashProvider = new HashProvider(SystemMemCache.GetSystemConfigValue(SysCfgConsts.PRIVATE, SysCfgConsts.PRIVATE_HASH_ALGORITHM));
            _apiBlock = new ApiHelper(string.Empty, string.Empty);
        }
        private void ConfigDigitalSignature()
        {
            if (string.IsNullOrEmpty(BOSApp.CurrentUsersInfo.ADUserCaIdentity)) return;
            _caProvider = new DigitalSignatureProvider(BOSApp.CurrentCompanyInfo.CSCompanyCaProvider);
            _digitalSig = _caProvider.GetInstance();
        }
        private void BackgroundLoadCache()
        {
            if (_sysHelper.AllowThread())
            {
                //uthv co the chay am tham ben duoi ma ko anh huong
                var thGetEmrProcessSystemConfigs = new System.Threading.Thread(() => _entity.GetEmrProcessSystemConfigs());
                thGetEmrProcessSystemConfigs.Start();
                //AppMemCache.GetEmrTemplates(); // No use
            }
            else
            {
                _entity.GetEmrProcessSystemConfigs();
            }
        }

        #region RabibitMQ 

        private void InitQueueHisToEmr()
        {
            var queueName = RabbitMqConnectionManager.GetQueueName("EMRModule_Queue");
            _chanelQueueHisToEmr = RabbitMqConnectionManager.CreateChannel();
            _chanelQueueHisToEmr.QueueDeclare(queue: queueName,
                     durable: false, exclusive: false, autoDelete: true, arguments: null);
            var consumer = new EventingBasicConsumer(_chanelQueueHisToEmr);
            consumer.Received += (model, ea) =>
            {
                byte[] body = ea.Body.ToArray();
                var dataJson = Encoding.UTF8.GetString(body);
                ActionHisToEmrData data = null;
                try
                {
                    data = JsonConvert.DeserializeObject<ActionHisToEmrData>(dataJson);

                }
                catch
                {

                }

                switch (data.Action)
                {
                    case "HisSearch":
                        {
                            SetSearchParamValueInvoke(ParentScreen.SearchQuickContainer, "fld_txtMEEmrNo", data.EmrNo);
                            SetSearchParamValueInvoke(ParentScreen.SearchQuickContainer, "chkSearchByEmrCode", true);
                            SetSearchParamValueInvoke(ParentScreen.SearchQuickContainer, "fld_dteSearchToMEEmrCreatedDate", null);
                            SetSearchParamValueInvoke(ParentScreen.SearchQuickContainer, "fld_dteSearchFromMEEmrCreatedDate", null);
                            //ParentScreen.ShowDialog(); 

                            //ParentScreen.Invoke(new Action(() =>
                            //{
                            //    var handle = Process.GetCurrentProcess().MainWindowHandle; // Or use: Process.GetCurrentProcess().MainWindowHandle
                            //    ShowWindow(handle, SW_RESTORE);       // Restore if minimized
                            //    SetForegroundWindow(handle);
                            //}));
                            QuickSearch();

                            //query  tờ điều trị

                            //select tờ điều trị

                            break;
                        }
                }

                try
                {
                    var task = Task.Run(() =>
                    {
                        //ParentScreen.Invoke(new Action(() =>
                        //{
                        //}));

                        var handle = Process.GetCurrentProcess().MainWindowHandle; // Or use: Process.GetCurrentProcess().MainWindowHandle
                        ShowWindow(handle, SW_RESTORE);       // Restore if minimized
                        SetForegroundWindow(handle);
                    });
                    task.Wait();

                }
                catch (Exception ex)
                {

                }
            };

            _chanelQueueHisToEmr.BasicConsume(queue: queueName, noAck: true, consumer: consumer);
        }
        private void InitExchangeHisToEmr()
        {
            _chanelExchangeHisToEmr = RabbitMqConnectionManager.CreateChannel();
            _chanelExchangeHisToEmr.ExchangeDeclare(exchange: RabbitMqExchange.HisToEmr_Exchange,
                       type: ExchangeType.Direct,
                       durable: false,
                       autoDelete: false, null);
            // declare a server-named queue
            QueueDeclareOk queueDeclareResult = _chanelExchangeHisToEmr.QueueDeclare();
            string routingKey = RabbitMqConnectionManager.GetQueueName("EMRModule");
            string queueName = queueDeclareResult.QueueName;
            _chanelExchangeHisToEmr.QueueBind(queue: queueName, exchange: RabbitMqExchange.HisToEmr_Exchange, routingKey: routingKey);

            var consumer = new EventingBasicConsumer(_chanelExchangeHisToEmr);
            consumer.Received += ReceiveDataActionFromHis;
            _chanelExchangeHisToEmr.BasicConsume(queueName, noAck: true, consumer: consumer);
        }

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private const int SW_RESTORE = 3;
        private const int SW_MINIMIZE = 6;
        private void ReceiveDataActionFromHis(object model, BasicDeliverEventArgs ea)
        {
            byte[] body = ea.Body.ToArray();
            var dataJson = Encoding.UTF8.GetString(body);
            ActionHisToEmrData data = null;
            try
            {
                data = JsonConvert.DeserializeObject<ActionHisToEmrData>(dataJson);

            }
            catch
            {

            }

            if (data != null)
            {


                switch (data.Action)
                {
                    case "HisReturnDataDieuTri":
                        {
                            _dataJsonFromHIS = data.Data;
                            //_waittingForHIS = false;
                            BOSApp._hisWaitHandle.Set();
                        }
                        break;
                    case "HisReturnDataTongKetBenhAn":
                        {
                            _dataJsonFromHIS = data.Data;
                            //_waittingForHIS = false;
                            BOSApp._hisWaitHandle.Set();
                        }
                        break;
                    case "HisSearch":
                        {
                            SetSearchParamValueInvoke(ParentScreen.SearchQuickContainer, "fld_txtMEEmrNo", data.EmrNo);
                            SetSearchParamValueInvoke(ParentScreen.SearchQuickContainer, "chkSearchByEmrCode", true);
                            SetSearchParamValueInvoke(ParentScreen.SearchQuickContainer, "fld_dteSearchToMEEmrCreatedDate", null);
                            SetSearchParamValueInvoke(ParentScreen.SearchQuickContainer, "fld_dteSearchFromMEEmrCreatedDate", null);
                            //ParentScreen.ShowDialog(); 

                            //ParentScreen.Invoke(new Action(() =>
                            //{
                            //    var handle = Process.GetCurrentProcess().MainWindowHandle; // Or use: Process.GetCurrentProcess().MainWindowHandle
                            //    ShowWindow(handle, SW_RESTORE);       // Restore if minimized
                            //    SetForegroundWindow(handle);
                            //}));
                            QuickSearch();

                            //query  tờ điều trị

                            //select tờ điều trị

                            try
                            {
                                var task = Task.Run(() =>
                                {
                                    //ParentScreen.Invoke(new Action(() =>
                                    //{
                                    //}));

                                    var handle = Process.GetCurrentProcess().MainWindowHandle; // Or use: Process.GetCurrentProcess().MainWindowHandle
                                    ShowWindow(handle, SW_RESTORE);       // Restore if minimized
                                    SetForegroundWindow(handle);
                                });
                                task.Wait();

                            }
                            catch (Exception ex)
                            {

                            }
                            break;
                        }
                }

               


            }
        }

        #endregion
        public override void Show()
        {
            base.Show();
            CurrentModuleEntity.UpdateSearchObjectBindingSource();
            this.OpenGuidance();
            EnableEmrField(false);
            SetStatusControl(0);
            if (BOSApp.CurrentUserGroupInfo.ADUserGroupRole == UserGroupRole.admin.ToString())
            {
                ParentScreen.SetEnableOfToolbarButton("EmrForceReleaseDocument", true);
            }
            else
            {
                ParentScreen.SetEnableOfToolbarButton("EmrForceReleaseDocument", false);
            }

            _dpnMsg.Visible = false;
            if (_notificationTab)
            {
                _dpnMsg.Visible = true;
            }
            RemainEmrReopenButNotClose();
        }
        internal void RunActionsOnHisUpdated(MEEmrDocumentsInfo doc)
        {
            var actions = this._templateActionCtrl.GetAllByTemplateID(doc.FK_METemplateID).Where(o => o.MEEmrTemplateActionWhen == EmrTemplateActionWhen.HisUpdated.ToString()).OrderBy(o => o.MEEmrTemplateActionOrder).ToList();
            foreach (var act in actions)
            {
                var action = this._actionsController.GetObjectByID(act.FK_MEEmrActionID) as MEEmrActionsInfo;
                if (action != null)
                {
                    if (_checkSystem)
                    {
                        _sysHelper.LogTxt("information", $"Bắt đầu chạy chức năng {action.MEEmrActionNo}.");
                        var watchAct = Stopwatch.StartNew();
                        CallEmrAction($"{action.MEEmrActionNo}{EmrParam.TagCodeSeparator}{EmrParam.GuidTag}={doc.MEEmrDocumentGuid}", this._richEditCtrl.Document.Range);
                        watchAct.Stop();
                        var elapsedAct = watchAct.ElapsedMilliseconds / 1000.0;
                        _sysHelper.LogTxt("information", $"{elapsedAct} giây. Hoàn tất chạy chức năng {action.MEEmrActionNo}.");
                    }
                    else
                    {
                        CallEmrAction($"{action.MEEmrActionNo}{EmrParam.TagCodeSeparator}{EmrParam.GuidTag}={doc.MEEmrDocumentGuid}", this._richEditCtrl.Document.Range);
                    }
                }
            }
        }

        /// <summary>
        /// Chạy các action được cấu hình auto run khi mở tài liệu
        /// </summary>
        /// <param name="doc"></param>
        internal void RunActionsOnOpen(MEEmrDocumentsInfo doc)
        {
            var actions = this._templateActionCtrl.GetAllByTemplateID(doc.FK_METemplateID)
                .Where(o => o.MEEmrTemplateActionWhen == EmrTemplateActionWhen.Open.ToString())
                .OrderBy(o => o.MEEmrTemplateActionOrder).ToList();
            foreach (var act in actions)
            {
                var action = this._actionsController.GetObjectByID(act.FK_MEEmrActionID) as MEEmrActionsInfo;
                if (action != null)
                {
                    if (_checkSystem)
                    {
                        _sysHelper.LogTxt("information", $"Bắt đầu chạy chức năng {action.MEEmrActionNo}.");
                        var watchAct = Stopwatch.StartNew();
                        CallEmrAction($"{action.MEEmrActionNo}{EmrParam.TagCodeSeparator}{EmrParam.GuidTag}={doc.MEEmrDocumentGuid}", this._richEditCtrl.Document.Range);
                        watchAct.Stop();
                        var elapsedAct = watchAct.ElapsedMilliseconds / 1000.0;
                        _sysHelper.LogTxt("information", $"{elapsedAct} giây. Hoàn tất chạy chức năng {action.MEEmrActionNo}.");
                    }
                    else
                    {
                        CallEmrAction($"{action.MEEmrActionNo}{EmrParam.TagCodeSeparator}{EmrParam.GuidTag}={doc.MEEmrDocumentGuid}", this._richEditCtrl.Document.Range);
                    }
                }
            }
        }
        private void RichEditControl_BeforePagePaint(object sender, BeforePagePaintEventArgs e)
        {
            if (e.CanvasOwnerType == CanvasOwnerType.Printer)
                return;
            if (_pagePainterVer2 == false)
                e.Painter = new PagePainterV1(_richEditCtrl, _highlightEmrTagUserConfig, _highlightModeEmrTagBorder, _highlightModeEmrTagFill);
            else
                e.Painter = new PagePainterV2(_richEditCtrl, _highlightEmrTagUserConfig, _highlightModeEmrTagBorder, _highlightModeEmrTagFill);
        }
        private void RichEditTempControl_BeforePagePaint(object sender, BeforePagePaintEventArgs e)
        {
            if (string.IsNullOrEmpty(_pagePrintPainterVer)) return;
            if (_pagePrintPainterVer == "V1")
                e.Painter = new PagePrintPainterBaseV1(_tempRichEditCtrl);
            else
                e.Painter = new PagePrintPainterBaseV2(_tempRichEditCtrl);
        }
        private void richEditCtrl_UnhandledException(object sender, RichEditUnhandledExceptionEventArgs e)
        {
            e.Handled = true;
            // Add your code to handle the exception. 
            _richEditCtrl.BeginInvoke(new MethodInvoker(delegate ()
            {
                MessageBox.Show(e.Exception.Message, "The command is not completed!");
            }));
        }
        internal void ClickThenMoveNextTag()
        {
            var doc = _richEditCtrl.Document;
            if (doc.Selections.Count == 1 && doc.Selections[0].Length == 0)
                GotoNextParam(Keys.Right);
        }
        internal void ClickThenMoveToBeginTag()
        {
            var doc = _richEditCtrl.Document;
            if (doc.Selections.Count == 1 && doc.Selections[0].Length == 0)
            {
                var caretPosition = doc.CaretPosition.ToInt();
                if (caretPosition < 1) return;
                var text = doc.GetText(doc.CreateRange(caretPosition - 1, 2));
                if (string.IsNullOrEmpty(text)) return;
                if (text[0] == '\t' && text[1] == EmrParam.EndTag[0])
                    doc.CaretPosition = doc.CreatePosition(caretPosition - 1);
            }
        }
        public bool IsPositionInsideEmrTag(int position)
        {
            var doc = _richEditCtrl.Document;
            var field = _emrDocumentHelper.GetFieldAtPosition(position);
            if (field == null)
            {
                return false;
            }
            else
            {
                var text = doc.GetText(field.ResultRange);
                if (!text.StartsWith(EmrParam.BeginTag)) return false;
                return true;
            }
        }
        public bool AllowEditAtPointOuterEmrTag()
        {
            //Console.WriteLine("CHECK");
            if (_allowEditOuterEmrTag) return true;
            return IsPositionInsideEmrTag(_richEditCtrl.Document.CaretPosition.ToInt());
        }
        public bool AllowEditRangeOuterEmrTag()
        {
            if (_allowEditOuterEmrTag) return true;
            var doc = _richEditCtrl.Document;
            foreach (var selection in doc.Selections)
            {
                var fields = doc.Fields.Get(selection);
                if (fields.Count > 0) return false;
            }
            return AllowEditAtPointOuterEmrTag();
        }

        private void _richEditCtrl_KeyDown(object sender, KeyEventArgs e)
        {
            //Console.WriteLine(e.KeyCode);
            if (IsDirectionKey(e.KeyCode))
                _isRichEditInputingMode = false;

            if (e.KeyCode == Keys.Left
                || e.KeyCode == Keys.Right
                || e.KeyCode == Keys.Up
                || e.KeyCode == Keys.Down)
            {
                return;
            }
            if (e.Control
                && e.KeyCode != Keys.X
                && e.KeyCode != Keys.C
                && e.KeyCode != Keys.V)
            {
                return;
            }
            var doc = _richEditCtrl.Document;
            if (e.KeyCode == Keys.Insert)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }
            else if (e.KeyCode == Keys.Tab)
            {
                if (_moveBetweenTagByTabEnabled)
                {
                    var direction = e.Shift ? Keys.Left : Keys.Right;
                    GotoNextParam(direction);
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    return;
                }
            }

            if (Control.IsKeyLocked(Keys.Insert))
            {
                ShowFlashNotification("Chế độ INSERT đang BẬT. Vui lòng tắt để tiếp tục.", 4000);
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }

            var msg = _allowEditOuterEmrTag ? "Thao tác xóa thẻ này KHÔNG HỢP LỆ. Quét chọn TOÀN BỘ thẻ và bấm DELETE nếu muốn xóa bỏ thẻ." : "Thao tác xóa thẻ KHÔNG ĐƯỢC PHÉP";
            if (doc.Selections.Count == 1 && doc.Selections[0].Length == 0)
            {
                //uthv dùng phương pháp này để tăng tốc độ nhập dữ liệu
                if (!_isRichEditInputingMode)
                    _isRichEditInputingMode = AllowEditAtPointOuterEmrTag();

                if (!_isRichEditInputingMode)
                {
                    ShowFlashNotification("Chỉ được nhập thông tin vào trong thẻ dữ liệu", 3000);
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    return;
                }

                if (!_emrDocumentHelper.IsAllowEditField(doc.CaretPosition))
                {
                    ShowFlashNotification("Thẻ dữ liệu đã bị chặn sửa thủ công", 3000);
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    return;
                }

                if (e.KeyCode == Keys.Delete)
                {
                    var pos = doc.CaretPosition.ToInt();
                    if (pos == doc.Range.End.ToInt()) return;
                    var text = doc.GetText(doc.CreateRange(pos, 1));
                    //Console.WriteLine("pos: " + pos + " text: " + text);
                    if (text.Contains(EmrParam.BeginTag) || text.Contains(EmrParam.EndTag))
                    {
                        ShowFlashNotification(msg, 4000);
                        e.Handled = true;
                        e.SuppressKeyPress = true;
                    }
                }
                else if (e.KeyCode == Keys.Back)
                {
                    var pos = doc.CaretPosition.ToInt();
                    if (pos <= 0) return;
                    var prePos = doc.CreatePosition(pos - 1);
                    var field = doc.Fields.Any(f => f.ResultRange.Start == prePos);
                    if (field)
                    {
                        ShowFlashNotification(msg, 3000);
                        e.Handled = true;
                        e.SuppressKeyPress = true;
                    }
                    field = doc.Fields.Any(f => f.ResultRange.End == prePos);
                    if (field)
                    {
                        ShowFlashNotification(msg, 3000);
                        e.Handled = true;
                        e.SuppressKeyPress = true;
                    }
                }
                else
                {
                    // Console.WriteLine(e.KeyCode);
                }
            }
            else
            {
                if (e.Control && e.KeyCode == Keys.C)
                {
                    if (_notAllowCopyEmrTag)
                        CopyJsonValueOfRange(e);
                    return;
                }

                if (e.Control && e.KeyCode == Keys.V)
                {
                    if (_notAllowCopyEmrTag)
                        PasteParamValueAtRanges(e);
                    else
                        PasteParamValue(e);
                    return;
                }

                if (!AllowEditRangeOuterEmrTag())
                {
                    ShowFlashNotification("Chỉ được sửa thông tin trong thẻ dữ liệu", 3000);
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    return;
                }

                if (!_emrDocumentHelper.IsAllowEditField(doc.CaretPosition))
                {
                    ShowFlashNotification("Thẻ dữ liệu đã bị chặn sửa thủ công", 3000);
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    return;
                }

                if (!_emrDocumentHelper.IsAllowEditSelections(doc.Selections))
                {
                    ShowFlashNotification("Vùng chọn hiện có chứa thẻ dữ liệu đã bị chặn sửa thủ công. Vui lòng chọn lại", 3000);
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    return;
                }
                var text = string.Join("", doc.Selections.Select(o => doc.GetText(o)));
                if (IsInputKey(e.KeyCode))
                {
                    if (text.Contains(EmrParam.BeginTag) || text.Contains(EmrParam.EndTag))
                    {
                        ShowFlashNotification(msg, 3000);
                        e.Handled = true;
                        e.SuppressKeyPress = true;
                        return;
                    }
                }
                else if (e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back || e.KeyCode == Keys.Tab || e.KeyCode == Keys.Enter)
                {
                    if (text.Count(c => c == EmrParam.BeginTag[0]) != text.Count(c => c == EmrParam.EndTag[0]))
                    {
                        ShowFlashNotification(msg, 3000);
                        e.Handled = true;
                        e.SuppressKeyPress = true;
                        return;
                    }
                }

            }
        }

        private bool IsInputKey(Keys keyCode)
        {
            if (keyCode >= Keys.A && keyCode <= Keys.Z)
                return true;
            if (keyCode >= Keys.D0 && keyCode <= Keys.D9)
                return true;
            if (keyCode >= Keys.NumPad0 && keyCode <= Keys.Divide)
                return true;
            if (keyCode >= Keys.OemSemicolon && keyCode <= Keys.Oem102)
                return true;
            if (keyCode == Keys.Space) return true;
            return false;
        }
        private bool IsDirectionKey(Keys keyCode)
        {
            if (keyCode >= Keys.Prior && keyCode <= Keys.Down)
                return true;
            if (keyCode >= Keys.LButton && keyCode <= Keys.XButton2)
                return true;
            return false;
        }

        #region Notification
        private void ClearAllLogAndNotification()
        {
            _msgNotification.Text = string.Empty;
            if (!_notificationTab)
            {
                _msgLogs.Text = string.Empty;
            }
            _qatNotificationMsg.EditValue = string.Empty;
        }

        private void ShowFlashNotification(string msg, int timeout, int waitTime = 0, LabelControl contrl = null)
        {
            if (contrl == null) contrl = this._msgNotification;
            if (waitTime == 0)
            {
                if (contrl.InvokeRequired)
                {
                    contrl.BeginInvoke((MethodInvoker)delegate () { contrl.Text = msg; });
                }
                else
                {
                    contrl.Text = msg;
                }
            }
            else
            {
                AssignFlashNotification(msg, waitTime, contrl);
            }
            AssignFlashNotification(string.Empty, timeout + waitTime, contrl);
        }
        private void AssignFlashNotification(string msg, int time, LabelControl contrl)
        {
            try
            {
                System.Threading.Timer waitTimer = null;
                waitTimer = new System.Threading.Timer((obj) =>
                {
                    if (contrl.InvokeRequired)
                    {
                        contrl.BeginInvoke((MethodInvoker)delegate () { contrl.Text = msg; });
                    }
                    else
                    {
                        contrl.Text = msg;
                    }
                    waitTimer.Dispose();
                }, null, time, System.Threading.Timeout.Infinite);
            }
            catch (Exception ex)
            {
                _sysHelper.LogTxt("error", ex.ToString());
            }
        }
        private void ShowFlashNotificationInToolbar(string msg, int timeout, int waitTime = 0)
        {
            // Hot fix always show
            if (_msgNotification.InvokeRequired)
            {
                _msgNotification.BeginInvoke((MethodInvoker)delegate () { _qatNotificationMsg.EditValue = msg; });
            }
            else
            {
                _qatNotificationMsg.EditValue = msg;
            }
            //if (waitTime == 0)
            //{
            //    if (_msgNotification.InvokeRequired)
            //    {
            //        _msgNotification.BeginInvoke((MethodInvoker)delegate () { _qatNotificationMsg.EditValue = msg; });
            //    }
            //    else
            //    {
            //        _qatNotificationMsg.EditValue = msg;
            //    }
            //}
            //else
            //{
            //    AssignFlashNotificationInToolbar(msg, waitTime);
            //}
            //AssignFlashNotificationInToolbar(string.Empty, timeout + waitTime);
        }
        private void AssignFlashNotificationInToolbar(string msg, int time)
        {
            System.Threading.Timer waitTimerToolbar = null;
            waitTimerToolbar = new System.Threading.Timer((obj) =>
            {
                if (_msgNotification.InvokeRequired)
                {
                    _msgNotification.BeginInvoke((MethodInvoker)delegate () { _qatNotificationMsg.EditValue = msg; });
                }
                else
                {
                    _qatNotificationMsg.EditValue = msg;
                }
                waitTimerToolbar.Dispose();
            }, null, time, System.Threading.Timeout.Infinite);
        }
        #endregion

        private void _richEditCtrl_PageBasedRichEditView_PageCountChanged(object sender, EventArgs e)
        {
            if (_entity.METemplate == null) return;

            PageBasedRichEditView currentView = null;
            if (_entity.METemplate.METemplateMaximumPage <= 0)
                return;
            else
            {
                currentView = sender as PageBasedRichEditView;
                if (currentView != null && currentView.PageCount > _entity.METemplate.METemplateMaximumPage)
                {
                    _msgNotification.Text = $"Tờ bệnh án vượt số trang quy định ({_entity.METemplate.METemplateMaximumPage} trang).";
                }
                else
                {
                    if (_msgNotification.Text.StartsWith("Tờ bệnh án vượt số trang quy định"))
                        _msgNotification.Text = string.Empty;
                }
            }

            if (_entity.METemplatePrimaryHeader != null)
            {
                currentView = sender as PageBasedRichEditView;
                //Them header cho to benh an tu trang 2 tro di #363
                if (currentView != null && currentView.PageCount == 2)
                {
                    Section firstSection = _richEditCtrl.Document.Sections[0];
                    var headerContent = false;
                    if (firstSection.HasHeader(HeaderFooterType.Primary))
                    {
                        SubDocument newHeader = firstSection.BeginUpdateHeader();
                        if (newHeader.Length > 1)
                            headerContent = true;
                        //headerContent = newHeader.GetText(newHeader.Range).Trim();
                        firstSection.EndUpdateHeader(newHeader);
                    }
                    if (!headerContent)
                    {
                        var formatDt = "yyyyMMddHHmmssfff";
                        string filePath = string.Format(@"{0}\Template\{1}_{2}.docx", _documentPath, _entity.METemplatePrimaryHeader.METemplateNo, _entity.METemplatePrimaryHeader.AAUpdatedDate.ToString(formatDt));
                        if (!File.Exists(filePath))
                        {
                            _ftpFileMng.DownloadFile("/Template/", _entity.METemplatePrimaryHeader.METemplateNo + ".docx", filePath);
                        }

                        if (!File.Exists(filePath))
                        {
                            MessageBox.Show("Không tải được file. File mẫu hearder không tồn tại ở địa chỉ. " + filePath);
                        }
                        else
                        {
                            _emrDocumentHelper.InsertPrimaryHeader(_richEditCtrl.Document, filePath);
                        }
                    }
                    else
                    {
                        //PrintMgsLog("HEADER_TU_DONG", "Tờ bệnh án đã có sẵn header. Không thể thêm header mới. Vui lòng xóa nội dung header cũ");
                    }
                }
            }
        }
        private void GetAllTemplate()
        {
            try
            {
                this._entity.METemplateList = this._templateCtrl.GetTemplatesByType(TemplateType.ProgressNote.ToString(), BOSApp.CurrentUserGroupInfo.ADUserGroupID);
            }
            catch (Exception ex)
            {
                _sysHelper.LogTxt("error", ex.ToString());
            }
        }
        bool IsSeparator(char ch)
        {
            return ch == '\r' || ch == '\n' || Char.IsPunctuation(ch) || Char.IsSeparator(ch) || Char.IsWhiteSpace(ch);
        }
        private void richEditControl1_AutoCorrect(object sender, DevExpress.XtraRichEdit.AutoCorrectEventArgs e)
        {
            AutoCorrectInfo info = e.AutoCorrectInfo;
            e.AutoCorrectInfo = null;

            if (info.Text.Length <= 0)
                return;

            if (!IsSeparator(info.Text[info.Text.Length - 1]))
                return;
            for (; ; )
            {
                if (!info.DecrementStartPosition())
                    return;
                if (IsSeparator(info.Text[0]) && info.Text.Length > 3)
                {
                    var key = info.Text.Substring(1, info.Text.Length - 2);
                    if (string.IsNullOrEmpty(key)) return;
                    //Console.WriteLine(key);
                    if (!_userAbbrevs.ContainsKey(key)) return;
                    var document = _richEditCtrl.Document;
                    document.BeginUpdate();
                    DocumentRange range = document.CreateRange(document.CaretPosition.ToInt() - key.Length - 1, key.Length);
                    var text = _userAbbrevs[key];
                    if (!text.StartsWith("{\\rtf1"))
                        document.Replace(range, text);
                    else
                    {
                        document.Replace(range, string.Empty);
                        range = document.InsertRtfText(range.Start, text);
                        document.Replace(document.CreateRange(document.CreatePosition(range.End.ToInt() - 1), 1), string.Empty);
                    }
                    document.EndUpdate();
                    return;
                }
            }
        }
        private void InitAutoCorrect()
        {
            var autoCorrectService = this._richEditCtrl.GetService<IAutoCorrectService>();
            var correctionOptions = this._richEditCtrl.Options.AutoCorrect;
            correctionOptions.ReplaceTextAsYouType = true;
            LoadUserAbbrevs();
            // if (autoCorrectService != null)
            //autoCorrectService.SetReplaceTable(LoadUserAbbrevs());
        }

        public override void InitializeScreens()
        {
            base.InitializeScreens();
        }

        internal void HideAllParam()
        {
            this._emrDocumentHelper.HideAllParam();
        }

        internal void ShowAllParam()
        {
            this._emrDocumentHelper.ShowAllParam();
        }
        internal void GotoNextParam(Keys keyCode)
        {
            this._emrDocumentHelper.GotoNextParam(keyCode);
            _richEditCtrl.ScrollToCaret();
        }
        internal void GotoNextParamV0(Keys keyCode)
        {
            this._emrDocumentHelper.GotoNextParamV0(keyCode);
            _richEditCtrl.ScrollToCaret();
        }
        public bool IsNothingToSaveDocumentContent()
        {
            if (!this._richEditCtrl.Modified) return true;

            if (IsEmrReadOnly()) return true;

            MEEmrsInfo emrCurr = _entity.MainObject as MEEmrsInfo;
            if (this._richEditCtrl.Modified)
            {
                var confirm = MessageBox.Show("Bạn có muốn lưu thay đổi?", "Nội dung tờ bệnh án đã thay đổi", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                if (confirm == DialogResult.Yes)
                {
                    var command = this._richEditCtrl.CreateCommand(RichEditCommandId.FileSave);
                    command.Execute();
                    return true;
                }
                else if (confirm == DialogResult.Cancel)
                {
                    for (int i = 0; i < _entity.MEEmrsList.Count; i++)
                    {
                        if (_entity.MEEmrsList[i].MEEmrID == emrCurr.MEEmrID)
                        {
                            _entity.MEEmrsList.GridView.FocusedRowHandle = _entity.MEEmrsList.GridView.GetRowHandle(i);
                            break;
                        }
                    }
                    return false;
                }
            }
            return true;
        }

        public override void Invalidate(int iObjectId)
        {
            if (!IsNothingToSaveDocumentContent()) return;

            if (iObjectId > 0)
                CreateEmrLocalDir(iObjectId);

            _emrDocumentCtrl.ReleaseAllEditing(BOSApp.CurrentEmployeesInfo.HREmployeeID, _macAddress);
            _richEditCtrl.ReadOnly = true;
            this._emrDocumentHelper.ClearCacheBindingFields();

            base.Invalidate(iObjectId);
            _entity.MEEmrArchiveList.Invalidate(iObjectId);

            LogFile(false); // Save log detail before clear
            ClearAllLogAndNotification();
            var tabEmrProfiles = (Controls["tabEmrProfiles"] as DevExpress.XtraTab.XtraTabPage);
            tabEmrProfiles.PageVisible = false;
            var tabPatientProfile = (Controls["tabPatientProfiles"] as DevExpress.XtraTab.XtraTabPage);
            tabPatientProfile.PageVisible = false;
            var tabEmrRelations = (Controls["tabEmrRelations"] as DevExpress.XtraTab.XtraTabPage);
            tabEmrRelations.PageVisible = false;
            var tacTransferAndShare = (Controls["tacTransferAndShare"] as DevExpress.XtraTab.XtraTabControl);
            tacTransferAndShare.Visible = false;

            var dpnDocTree = Controls["dpnDocTree"] as DockPanel;
            dpnDocTree.Text = string.Empty;

            if (iObjectId > 0)
            {
                if (Screen.PrimaryScreen.Bounds.Width < 1200)
                {
                    dpnDocTree.Visibility = DockVisibility.AutoHide;
                }
                if (dpnDocTree.Visibility == DockVisibility.AutoHide)
                {
                    dpnDocTree.Show();
                }
                tabEmrProfiles.PageVisible = true;
                tabPatientProfile.PageVisible = true;
                tacTransferAndShare.Visible = true;

                var meEmr = _entity.MainObject as MEEmrsInfo;

                #region Emr Stuck Initing
                StuckAtInitingState();
                #endregion

                #region Tac vu chay ngam
                GetEmrDocumentsBackground();
                #endregion

                dpnDocTree.Text = "Hồ sơ " + meEmr.MEEmrNo;
                var label = (dpnDocTree.CustomHeaderButtons[0] as CustomHeaderButton);
                if (meEmr.MEEmrTypeProfile == EmrTypeProfile.Patient.ToString())
                {
                    label.Caption = "HỒ SƠ";
                    label.Appearance.ForeColor = Color.Blue;
                }
                else if (meEmr.MEEmrTypeProfile == EmrTypeProfile.Out.ToString())
                {
                    label.Caption = "NGOẠI TRÚ";
                    label.Appearance.ForeColor = Color.DarkRed;
                }
                else if (meEmr.MEEmrTypeProfile == EmrTypeProfile.In.ToString())
                {
                    label.Caption = "NỘI TRÚ";
                    label.Appearance.ForeColor = Color.Black;
                }
                else
                {
                    var caption = ADConfigValueUtility.GetTextFromKey("EmrTypeProfile" + meEmr.MEEmrTypeProfile);
                    label.Caption = string.IsNullOrEmpty(caption) ? "BỆNH ÁN" : caption.ToUpper();
                    label.Appearance.ForeColor = Color.Black;
                }

                _entity.MEPatient = _patientCtrl.GetObjectByID(meEmr.FK_MEPatientID) as MEPatientsInfo;
                (this.Controls[_fld_lbl_MEPatientNoName] as Label).Text = _entity.MEPatient.MEPatientNo;
                (this.Controls[_fld_lbl_MEPatientNameName] as Label).Text = _entity.MEPatient.MEPatientName.ToUpper();
                // Cham Cuu request
                var ageTxt = XConvert.Age(_entity.MEPatient.MEPatientBirthday);
                if (_entity.MEPatient.MEPatientBirthdayOnlyYear)
                {
                    (this.Controls[_fld_lbl_MEPatientBirthdayName] as Label).Text = $"{_entity.MEPatient.MEPatientBirthday.Year} {ageTxt}";
                }
                else
                {
                    (this.Controls[_fld_lbl_MEPatientBirthdayName] as Label).Text = $"{_entity.MEPatient.MEPatientBirthday.ToString("dd/MM/yyyy")} {ageTxt}";
                }
                var emrPatientProfile = _emrCtrl.GetEmrPatientProfile(meEmr.FK_MEPatientID);

                SetStatusControl(meEmr.MEEmrID);

                if (emrPatientProfile != null)
                {
                    _entity.InvalidateModuleObject(emrPatientProfile);
                    InvalidatePatientDocuments(emrPatientProfile.MEEmrID);
                    if (emrPatientProfile.MEEmrID == iObjectId)
                    {
                        //neu dong dang click la ho so benh nhan > an tab ho so benh an, 
                        tabEmrProfiles.PageVisible = false;
                        tacTransferAndShare.Visible = false;
                        EnableEmrField(false);
                        dpnDocTree.Text = "Hồ sơ bệnh nhân " + emrPatientProfile.MEEmrNo;
                        //khong cho phep sua HSBN
                        ParentScreen.SetEnableOfToolbarButton("Edit", false);
                        ParentScreen.SetEnableOfToolbarButton("Close", false);
                        ParentScreen.SetEnableOfToolbarButton("MergeEmr", false);
                    }
                    ParentScreen.SetEnableOfToolbarButton("NewPatientProfile", false);
                }
                else
                {
                    //an tab ho so benh nhan
                    tabPatientProfile.PageVisible = false;
                    _entity.SetDefaultModuleObject(TableName.MEEmrsTableName);
                    _entity.MEEmrPatientDocumentsList.SetDefaultListAndRefreshGridControl();
                    ParentScreen.SetEnableOfToolbarButton("NewPatientProfile", true);
                }

                //active tab page tuy theo loai benh an hay hs benh nhan
                if (emrPatientProfile != null && emrPatientProfile.MEEmrID == iObjectId)
                    tabPatientProfile.TabControl.SelectedTabPage = tabPatientProfile;
                else
                    tabEmrProfiles.TabControl.SelectedTabPage = tabEmrProfiles;

                // đang chọn xem hồ sơ bệnh nhân thì không load danh sách bệnh án liên quan
                // đang chọn xem hồ sơ bệnh án thì load các bệnh án liên quan
                // BA Lien Ket hien thi khi Hồ sơ bệnh án.
                if (emrPatientProfile == null || emrPatientProfile.MEEmrID != iObjectId)
                {
                    InvalidateEmrsOfPatient();

                    InvalidateEmrDocumentList(iObjectId);
                    _msgNotification.Text = string.Empty;
                    CheckBgJobStatusDocumentList();
                    EnableEmrField(false);

                    // HTSS - LKBA
                    EmrRelationInit();
                }

                if (_sysHelper.AllowThread())
                {
                    //uthv co the chay am tham ben duoi ma ko anh huong
                    System.Threading.Thread thInitRequestParamPool = new System.Threading.Thread(() => InitRequestParamPool());
                    thInitRequestParamPool.Start();
                }
                else
                {
                    InitRequestParamPool();
                }

                _entity.InvalidateModuleObject(_entity.MEPatient);

                if (emrPatientProfile == null || emrPatientProfile.MEEmrID != iObjectId)
                {
                    if (_showPopupUpdateInfoPatient)
                        this.AskIfPatientInfoChanged();
                    this.AskForInitEmr();
                }

                // XuanTM - US 263
                if (!string.IsNullOrEmpty(meEmr.MEEmrNotifyMsg))
                {
                    MessageBox.Show(meEmr.MEEmrNotifyMsg, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                // XuanTM - Bug 866 - No use SetStatusControl(meEmr.MEEmrID) more complex; First code here, optimize later.
                if (meEmr.MEEmrStatus == EmrStatus.Closed.ToString())
                {
                    Controls["fld_btnTransfer"].Enabled = false;
                }
                if (meEmr.MEEmrHasNote)
                {
                    MessageBox.Show("Bệnh án có ghi chú.\nVui lòng sử dụng chức năng Xem ghi chú để hiển thị chi tiết.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                if (_sysHelper.AllowThread())
                {
                    //uthv co the chay am tham ben duoi ma ko anh huong
                    System.Threading.Thread thShowCurrentDepartment = new System.Threading.Thread(() => ShowCurrentDepartment());
                    thShowCurrentDepartment.Start();
                }
                else
                {
                    ShowCurrentDepartment();
                }

                RunActionsOnOpenEmr(meEmr);
            }
            else
            {
                _entity.MEPatient = new MEPatientsInfo();
                (this.Controls[_fld_lbl_MEPatientNoName] as Label).Text = _entity.MEPatient.MEPatientNo;
                (this.Controls[_fld_lbl_MEPatientNameName] as Label).Text = _entity.MEPatient.MEPatientName.ToUpper();
                (this.Controls[_fld_lbl_MEPatientBirthdayName] as Label).Text = _entity.MEPatient.MEPatientBirthdayOnlyYear ?
                    _entity.MEPatient.MEPatientBirthday.Year.ToString() :
                    _entity.MEPatient.MEPatientBirthday.ToString("dd/MM/yyyy");
                EnableEmrField(false);
                SetStatusControl(0);
                (this.Controls[_fld_lbl_MEEmrDepartment] as Label).Text = string.Empty;
            }

            if (Controls["pdfViewerEmrArchive"] != null)
            {
                (Controls["pdfViewerEmrArchive"] as PdfViewer).CloseDocument();
            }

            // Reset data merge history
            // MergeEmrRefresh(); // can be slow
            var gridControlMergeHistories = this.Controls["fld_dgcMEEmrMergeHistories"] as MEEmrMergeHistoriesGridControl;
            if (gridControlMergeHistories != null)
            {
                gridControlMergeHistories.DataSource = new List<MEEmrMergeHistoriesInfo>();
                gridControlMergeHistories.RefreshDataSource();
                gridControlMergeHistories.Refresh();
            }

            InvalidateGEObjectHistory();

            this.OpenGuidance();
        }
        /// <summary>
        /// use for ShowCurrentDepartment only, to prevent double popup
        /// </summary>
        private int _preObjectID = 0;

        /// <summary>
        /// Ut trong mot so truong hop thong tin table Emr bi doi
        /// can thuc hien load lai thong tin moi nhat
        /// </summary>
        internal bool ReloadIfEmrChangedByOther(MEEmrDocumentsInfo currentDoc)
        {
            if (currentDoc.MEEmrDocumentStatus == EmrDocumentStatus.Closed.ToString()) return false;
            if (currentDoc.MEEmrDocumentFileExt == EmrDocumentFileExtention.pdf.ToString()) return false;
            return ReloadIfEmrChangedByOther(currentDoc.FK_MEEmrID);
        }
        internal bool ReloadIfEmrChangedByOther(int emrID, bool showMessageBox = false)
        {
            var emr = _entity.MainObject as MEEmrsInfo;
            var lastUpdate = _emrCtrl.GetLastChangedTime(emrID);
            if (lastUpdate > emr.AAUpdatedDate)
            {
                var msg = "Thông tin bệnh án này đã bị thay đổi lúc " + lastUpdate.ToString("dd/MM/yyyy HH:mm:ss") + ". Đã làm mới thông tin.";
                if (showMessageBox)
                    MessageBox.Show(msg, "THÔNG TIN BỆNH ÁN ĐÃ THAY ĐỔI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                else
                    ShowFlashNotification(msg, 5000, 5000);

                Invalidate(emrID);
                emr = _entity.MainObject as MEEmrsInfo;
                RefreshCurrentEmrSearchResultsControl(emr);
                return true;
            }
            return false;
        }
        private void ShowCurrentDepartment()
        {
            var meEmr = _entity.MainObject as MEEmrsInfo;

            var dept = _departmentCtrl.GetObjectByID(meEmr.FK_HRDepartmentID) as HRDepartmentsInfo;
            if (dept != null)
            {
                var lblDept = (this.Controls[_fld_lbl_MEEmrDepartment] as Label);
                if (lblDept.InvokeRequired)
                    lblDept.Invoke((MethodInvoker)(() => lblDept.Text = dept.HRDepartmentName.ToUpper()));
                else
                    lblDept.Text = dept.HRDepartmentName.ToUpper();
                var type = meEmr.MEEmrTypeProfile == EmrTypeProfile.Patient.ToString() ? "Hồ sơ" : "Bệnh án";

                if (_preObjectID == meEmr.MEEmrID) return;

                _preObjectID = meEmr.MEEmrID;
                if (_useMessageBox)
                    MessageBox.Show($"{type} {meEmr.MEEmrNo} và bệnh nhân {_entity.MEPatient.MEPatientName} đang quản lý bởi khoa\n\n{dept.HRDepartmentNo} - {dept.HRDepartmentName.ToUpper()}",
                        $"KHOA ĐANG QUẢN LÝ {type.ToUpper()}", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0x40000);
            }
            _preObjectID = meEmr.MEEmrID;

        }

        /// <summary>
        /// his goi api emr va cap nhat thong tin benh nhan, dong thoi set MEPatientMatchCode01Combo = "UpdatedOnHis"
        /// </summary>
        private void AskIfPatientInfoChanged()
        {
            if (_entity.MEPatient.MEPatientMatchCode01Combo == "UpdatedOnHis" && !IsEmrReadOnly(true))
            {
                var confirm = MessageBox.Show("Thông tin hành chính của bệnh nhân có sự thay đổi.\n"
                        + "- Yes: Tự động cập nhật trên các tờ bệnh án. Sẽ mất chút thời gian\n"
                        + "- No: Tôi sẽ tự cập nhật bằng cách click vào các thẻ chức năng\n"
                        + "- Cancel: Tôi không biết. Sẽ cập nhật sau.",
                        "Bạn có muốn cập nhật thay đổi?", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                var emr = _entity.MainObject as MEEmrsInfo;
                if (confirm == DialogResult.Yes)
                {
                    try
                    {
                        var synsApply = emr.MEEmrStatus != EmrStatus.Initing.ToString() ? true : false;
                        var documents = _emrDocumentCtrl.GetByEmrId(emr.MEEmrID);
                        var orderedDocs = _emrDocumentSortHelper.SortDocumentTree(documents, emr);
                        _entity.MEEmrDocumentsList.Invalidate(orderedDocs);
                        var count = _entity.MEEmrDocumentsList.Count;
                        for (int i = 0; i < count; i++)
                        {
                            var item = _entity.MEEmrDocumentsList[i];
                            if (item.MEEmrDocumentStatus == EmrStatus.Closed.ToString() || item.MEEmrDocumentFileExt != EmrDocumentFileExtention.docx.ToString())
                                continue;

                            if (synsApply)
                            {
                                InvalidateDocument(item, true, false);
                                _sysnSave = false;
                            }
                            else
                            {
                                InvalidateDocument(item);
                            }

                            RunActionsOnHisUpdated(item);
                            var command = this._richEditCtrl.CreateCommand(RichEditCommandId.FileSave);
                            command.Execute();
                        }
                    }
                    catch (Exception ex)
                    {
                        _sysnSave = true;
                        Trace.TraceError("AskIfPatientInfoChanged ERROR: {0}:{1}:{2}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), BOSApp.CurrentUser, ex);
                        Trace.Flush();
                        MessageBox.Show($"Lỗi phát sinh trong quá trình lưu tờ bệnh án.",
                             "CÓ LỖI XẢY RA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    finally
                    {
                        _sysnSave = true;
                    }
                }
                if (confirm == DialogResult.Yes || confirm == DialogResult.No)
                {
                    var patient = _entity.MEPatient.Clone() as MEPatientsInfo;
                    patient.MEPatientMatchCode01Combo = string.Empty;
                    patient.AAUpdatedUser = BOSApp.CurrentUser;
                    _patientCtrl.UpdateObject(patient);
                    _entity.MEPatient = patient;
                }
                if (confirm == DialogResult.Yes)
                {
                    this.Invalidate(emr.MEEmrID);
                }
            }
        }

        public void UpdateHisChanges()
        {
            if (IsEmrReadOnly(true)) return;
            var confirm = MessageBox.Show("Thông tin sẽ được lấy từ HIS và cập nhật lại. Các thay đổi trên tờ bệnh án bạn đã thực hiện có thể bị cập nhật ghi đè bởi thao tác này.",
                "Bạn có muốn cập nhật thay đổi?", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            var emr = _entity.MainObject as MEEmrsInfo;
            try
            {
                if (confirm == DialogResult.OK)
                {
                    var documents = _emrDocumentCtrl.GetByEmrId(emr.MEEmrID);
                    var orderedDocs = _emrDocumentSortHelper.SortDocumentTree(documents, emr);
                    _entity.MEEmrDocumentsList.Invalidate(orderedDocs);
                    var count = _entity.MEEmrDocumentsList.Count;
                    for (int i = 0; i < count; i++)
                    {
                        var item = _entity.MEEmrDocumentsList[i];

                        if (item.MEEmrDocumentStatus == EmrStatus.Closed.ToString() || item.MEEmrDocumentFileExt != EmrDocumentFileExtention.docx.ToString())
                            continue;
                        InvalidateDocument(item);
                        RunActionsOnHisUpdated(item);
                        _sysnSave = false;
                        var command = this._richEditCtrl.CreateCommand(RichEditCommandId.FileSave);
                        command.Execute();
                    }
                    _entity.MEPatient = _patientCtrl.GetObjectByID(emr.FK_MEPatientID) as MEPatientsInfo;
                    if (_entity.MEPatient.MEPatientMatchCode01Combo == "UpdatedOnHis")
                    {
                        var patient = _entity.MEPatient.Clone() as MEPatientsInfo;
                        patient.MEPatientMatchCode01Combo = string.Empty;
                        patient.AAUpdatedUser = BOSApp.CurrentUser;
                        _patientCtrl.UpdateObject(patient);
                        _entity.MEPatient = patient;
                    }
                }
                if (confirm == DialogResult.OK)
                {
                    this.Invalidate(emr.MEEmrID);
                }
            }
            catch (Exception ex)
            {
                _sysnSave = true;
                MessageBox.Show("Có lỗi xảy ra. Xin thử lại.\nChi tiết: " + ex.Message,
                "Có lỗi xảy ra", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _sysnSave = true;
                this.ClearDocumentSession();
            }
        }
        private void AskForInitEmr()
        {
            if (!Toolbar.IsNullOrNoneAction()) return;
            var meEmr = _entity.MainObject as MEEmrsInfo;
            if (meEmr.FK_MEEmrTypeID > 0 && meEmr.MEEmrDocumentCount == 0 && meEmr.MEEmrStatus == EmrStatus.InProgress.ToString())
            {
                ShowInitAskControls(true);
                this._dpnRichEdit.Visible = true;
            }
            else
            {
                ShowInitAskControls(false);
            }
        }
        private void ShowInitAskControls(bool state)
        {
            Controls["btnInitEmr"].Visible = state;
            Controls["lblAskForInit"].Visible = state;
        }
        public void InitEmr()
        {
            if (!Toolbar.IsNullOrNoneAction()) return;
            var meEmr = _entity.MainObject as MEEmrsInfo;
            if (!IsExist(meEmr.MEEmrID)) return;
            if (meEmr.FK_MEEmrTypeID > 0 && meEmr.MEEmrDocumentCount == 0)
            {
                //if (MessageBox.Show("Bệnh án chưa có thông tin. Bạn có muốn khởi tạo bệnh án này?", "Khởi tạo bệnh án", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                //    return;
                InitEmr(meEmr, string.Empty, false);
                ShowInitAskControls(false);
            }
        }
        private void SetStatusControl(int emrId)
        {
            var edit = false;
            if (emrId > 0)
            {
                var editCap = BOSApp.GetUserEmrViewPermission();

                //[Vai trò = Quản trị] + [Quyền truy xuất = Toàn bộ] thì mới toàn quyền edit
                if (BOSApp.CurrentUserGroupInfo.ADUserGroupRole != UserGroupRole.admin.ToString())
                    editCap = UserEmrView.DEPARTMENT;

                var emr = _emrCtrl.CheckEditPermission(emrId, BOSApp.CurrentEmployeesInfo.HREmployeeID, BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID, editCap);
                if (emr != null)
                {
                    if (IsStatePermission(emr))
                    {
                        if (IsAllowEdit(emr))
                            edit = true;
                        else
                            edit = false;
                    }
                    else
                    {
                        if (emr.MEEmrStatus == EmrStatus.InProgress.ToString())
                        {
                            edit = true;
                        }
                        else
                        {
                            edit = false;
                        }
                    }
                }
            }
            SetStatusControl(edit);
        }
        private void SetStatusControl(bool edit)
        {
            ParentScreen.SetEnableOfToolbarButton("Approve", edit);
            ParentScreen.SetEnableOfToolbarButton("Close", edit);
            ParentScreen.SetEnableOfToolbarButton("Edit", edit);
            ParentScreen.SetEnableOfToolbarButton("MergeEmr", edit);
            ParentScreen.SetEnableOfToolbarButton("UnsignEmr", edit);
            ParentScreen.SetEnableOfToolbarButton("UnsignEmrGroup", edit);
            ParentScreen.SetEnableOfToolbarButton("UnsignEmrStrikeThrough", edit);
            ParentScreen.SetEnableOfToolbarButton("UpdateHisChanges", edit);
            ParentScreen.SetEnableOfToolbarButton("NewPatientProfile", edit);
            ParentScreen.SetEnableOfToolbarButton("EmrModuleChangeEmrTemplate", edit);
            ParentScreen.SetEnableOfToolbarButton("UpdateEmrTemplate", edit);
            Controls["fld_btn_DeleteEmrDocument"].Enabled = edit;
            Controls["fld_btn_AddNewEmrDocument"].Enabled = edit;
            Controls["fld_btn_HidingEmrDocument"].Enabled = edit;
            Controls["fld_btnTransfer"].Enabled = edit;
            var canShareEmr = CanShareEmr();
            var isCloseEmr = IsCloseEmr();
            var isCloseOrWaitCloseEmr = IsCloseOrWaitCloseEmr();
            Controls["fld_btnSelectDepartment"].Enabled = isCloseEmr ? false : edit;// edit;
            Controls["fld_btnSelectEmployee"].Enabled = canShareEmr; //edit;
            //Controls["fld_btnSelectDepartment"].Visible = !isCloseEmr;

            Controls["fld_tbnSaveShareEmrClose"].Visible = false;
            if (isCloseEmr && canShareEmr)
            {
                Controls["fld_tbnSaveShareEmrClose"].Visible = true;
            }

            ParentScreen.SetEnableOfToolbarButton("MergeEmrRollback", edit);
            if (isCloseOrWaitCloseEmr)
            {
                ParentScreen.SetEnableOfToolbarButton("MergeEmrRollback", false);
            }

            _overrideEditPermission = edit;
            (_entity.MEEmrShareList.GridView).OptionsBehavior.AllowAddRows = edit ? DevExpress.Utils.DefaultBoolean.True : DevExpress.Utils.DefaultBoolean.False;
            ShowInitAskControls(false);

            ParentScreen.SetEnableOfToolbarButton("EmrForceReleaseDocument", edit);
            ParentScreen.SetEnableOfToolbarButton("CloseDocument", edit);
            ParentScreen.SetEnableOfToolbarButton("TakeInitPermission", edit);
            ParentScreen.SetEnableOfToolbarButton("ChangeEmrType", edit);
            //ParentScreen.SetEnableOfToolbarButton("OpenTheClosedEmr", edit);
            //ParentScreen.SetEnableOfToolbarButton("EmrModuleArchiveEmr", edit);
            ParentScreen.SetEnableOfToolbarButton("EmrModuleCheckupEmr", edit);
            //ParentScreen.SetEnableOfToolbarButton("EmrStateMachine", edit);
            //ParentScreen.SetEnableOfToolbarButton("ToWaitClose", edit);
            //ParentScreen.SetEnableOfToolbarButton("BackToDept", edit);
            //ParentScreen.SetEnableOfToolbarButton("DocumentStateMachine", edit);
            ParentScreen.SetEnableOfToolbarButton("EmrModuleChangeEmrTemplate", edit);
            ParentScreen.SetEnableOfToolbarButton("UpdateEmrTemplate", edit);
            ParentScreen.SetEnableOfToolbarButton("EmrModuleGroupEmrDocuments", edit);
            //ParentScreen.SetEnableOfToolbarButton("EmrRelation", isCloseEmr ? false : edit);
            ParentScreen.SetEnableOfToolbarButton("EmrRenameEmrNo", isCloseEmr ? false : edit);
            var ribbon = this.Controls["richEditRibbonControl"] as RibbonControl;
            var pageEmr = ribbon.Pages.GetPageByName("ribbonPageEmr");
            if (pageEmr != null)
            {
                foreach (RibbonPageGroup group in pageEmr.Groups)
                {
                    group.Enabled = edit;
                }
            }
            var dpnDocTree = Controls["dpnDocTree"] as DockPanel;
            var lblReadOnly = (dpnDocTree.CustomHeaderButtons[1] as CustomHeaderButton);
            lblReadOnly.Visible = !edit;
        }
        public void InitRequestParamPool(MEEmrDocumentsInfo doc = null)
        {
            try
            {
                _requestParamPool = new ParamPool();
                //get one document for example
                MEEmrsInfo emr = GetCurrentMainObject();
                MEPatientsInfo patient = _entity.MEPatient;
                if (!string.IsNullOrEmpty(patient.MEPatientNo))
                    _requestParamPool.AddFromDataRow(_emrCtrl.GetRequestPoolParams(
                        patient.MEPatientNo,
                        emr != null ? emr.MEEmrNo : "",
                        doc != null ? doc.MEEmrDocumentNo : "",
                        BOSApp.CurrentUsersInfo.ADUserID,
                        BOSApp.CurrentEmployeesInfo.HREmployeeID,
                        BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID
                        ));

                _requestParamPool.AddOrUpdate(new KeyValuePair<string, object>("ADUserID", BOSApp.CurrentUsersInfo.ADUserID));
                _requestParamPool.AddOrUpdate(new KeyValuePair<string, object>("ADUserName", BOSApp.CurrentUsersInfo.ADUserName));
                _requestParamPool.AddOrUpdate(new KeyValuePair<string, object>("ADUserHISID", BOSApp.CurrentUsersInfo.ADUserHISID));

                _requestParamPool.AddOrUpdate(new KeyValuePair<string, object>("HREmployeeID", BOSApp.CurrentEmployeesInfo.HREmployeeID));
                _requestParamPool.AddOrUpdate(new KeyValuePair<string, object>("HREmployeeNo", BOSApp.CurrentEmployeesInfo.HREmployeeNo));
                _requestParamPool.AddOrUpdate(new KeyValuePair<string, object>("HREmployeeName", BOSApp.CurrentEmployeesInfo.HREmployeeName));

                _requestParamPool.AddOrUpdate(new KeyValuePair<string, object>("FK_HRDepartmentID", BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID));
            }
            catch (Exception ex)
            {
                _sysHelper.LogTxt("error", ex.ToString());
            }
        }
        internal void Invalidate(MEEmrsInfo mEEmrsInfo)
        {
            if (!IsNothingToSaveDocumentContent()) return;
            _msgNotification.Text = string.Empty;
            if (mEEmrsInfo != null)
            {
                mEEmrsInfo = CopyPatientInfoToEmr(mEEmrsInfo);
                Controls["dpnDocTree"].Text = "Hồ sơ bệnh án " + mEEmrsInfo.MEEmrNo;
                //var emr = _entity.MainObject as MEEmrsInfo;
                //if (emr != null && mEEmrsInfo.MEEmrID == emr.MEEmrID) return;
                var gridView = (_searchGrid.MainView as GridView);
                var tb = this.Toolbar.ObjectCollection.Tables[0] as DataTable;
                bool has = false, handled = false;
                for (int i = 0; i < tb.Rows.Count; i++)
                {
                    if (tb.Rows[i]["MEEmrID"].ToString() == mEEmrsInfo.MEEmrID.ToString())
                    {
                        //TODO: hien tai chua giai quyet dc focus row chinh xac nen cach tot nhat la ClearColumnsFilter
                        gridView.ClearColumnsFilter();
                        int newRowIdx = gridView.GetRowHandle(i);
                        handled = gridView.FocusedRowHandle != newRowIdx;
                        has = true;
                        gridView.FocusedRowHandle = newRowIdx;
                        break;
                    }
                }
                if (!has)
                {
                    var row = tb.NewRow();
                    tb.Rows.Add(_emrCtrl.GetDataRowFromBusinessObject(row, mEEmrsInfo));
                    gridView.FocusedRowHandle = gridView.GetRowHandle(tb.Rows.Count - 1);
                    //gridView.FocusedRowHandle = tb.Rows.Count - 1;
                }
            }
        }
        internal void InvalidateEmrsOfPatient()
        {
            List<MEEmrsInfo> emrs;
            var meEmr = _entity.MainObject as MEEmrsInfo;
            //khoa quan ly va cau hinh cho phep xem toan bo benh an cua benh nhan dang dieu tri
            if (BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID == meEmr.FK_HRDepartmentID && _entity.AllowAccessAllTheEmrsOfPatientAtManageDept)
            {
                if (meEmr.MEEmrTypeProfile == EmrTypeProfile.Out.ToString())
                    // đang chọn xem hồ sơ bệnh án ngoại trú thì không load bệnh án nội trú
                    emrs = _emrCtrl.GetEmrByPatientWithoutPatientProfile(_entity.MEPatient.MEPatientID, EmrTypeProfile.Out.ToString());
                else
                    // đang chọn xem hồ sơ bệnh án nội trú thì load tất cả bệnh án
                    emrs = _emrCtrl.GetEmrByPatientWithoutPatientProfile(_entity.MEPatient.MEPatientID, null);
            }
            else
            {
                var stateConds = string.Empty;
                if (_entity.AllowSharingTheClosedEmr)
                    stateConds = this.GetStatePermiQueryConditionStrSharedCase();
                else
                    stateConds = base.GetStatePermiQueryConditionStr(TableName.MEEmrsTableName, false);

                emrs = _emrCtrl.GetEmrByPatientWithSharedConds(
                        _entity.MEPatient.MEPatientID
                        , meEmr.MEEmrTypeProfile == EmrTypeProfile.Out.ToString() ? EmrTypeProfile.Out.ToString() : null
                        , BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID
                        , BOSApp.CurrentEmployeesInfo.HREmployeeID
                        , stateConds);
            }
            emrs = emrs.OrderByDescending(e => e.MEEmrID).ToList();
            _entity.MEEmrsList.Invalidate(emrs);
            _entity.MEEmrsList.GridControl.Refresh();
            _entity.MEEmrsList.GridView.FocusInvalidRow();
            for (int i = 0; i < _entity.MEEmrsList.Count; i++)
            {
                if (_entity.MEEmrsList[i].MEEmrID == meEmr.MEEmrID)
                {
                    _entity.MEEmrsList.GridView.FocusedRowHandle = _entity.MEEmrsList.GridView.GetRowHandle(i);
                    break;
                }
            }
        }
        private void ClearDocumentSession()
        {
            _entity.SetDefaultModuleObject(TableName.MEEmrDocumentsTableName);
            this._richEditCtrl.CreateNewDocument(false);
        }
        public void InvalidateEmrDocumentList()
        {
            var meEmr = _entity.MainObject as MEEmrsInfo;
            var gridView = _entity.MEEmrDocumentsList.GridView;
            gridView.ActiveFilter.Clear();
            InvalidateEmrDocumentList(meEmr.MEEmrID);
        }
        private void InvalidateEmrDocumentList(int emrID)
        {
            var meEmr = _entity.MainObject as MEEmrsInfo;
            var listDocs = new List<MEEmrDocumentsInfo>();
            if (_allowNormalUserHideDocument)
            {
                //LK yêu cầu người dùng được thấy toàn bộ tờ bệnh án bao gồm cả các tờ ẩn
                listDocs = _emrDocumentCtrl.GetListBusinessObjects<MEEmrDocumentsInfo>(_emrDocumentCtrl.GetByMEEmrID(emrID));
            }
            else
            {
                //chi co admin moi thay danh sach cac to benh an da bi an
                if (BOSApp.CurrentUserGroupInfo.ADUserGroupRole != UserGroupRole.admin.ToString())
                {
                    listDocs = _emrDocumentCtrl.GetByUserGroupWrite(emrID, BOSApp.CurrentUserGroupInfo.ADUserGroupID);
                    listDocs = listDocs.Where(o => o.MEEmrDocumentStatus != EmrDocumentStatus.Hidden.ToString()).ToList();
                }
                else
                {
                    //Admin đc quyền xem tất tần tật
                    listDocs = _emrDocumentCtrl.GetListBusinessObjects<MEEmrDocumentsInfo>(_emrDocumentCtrl.GetByMEEmrID(emrID));
                }
            }

            var orderedDocs = _emrDocumentSortHelper.SortDocumentTree(listDocs, meEmr);
            var gridView = _entity.MEEmrDocumentsList.GridView;
            var preDocumentId = (_entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo).MEEmrDocumentID;
            var preFocusDoc = gridView.LocateByValue("MEEmrDocumentID", preDocumentId);
            _entity.MEEmrDocumentsList.Invalidate(orderedDocs);
            _entity.MEEmrDocumentsList.GridView.FocusedRowHandle = -1;
            if (preFocusDoc != DevExpress.XtraGrid.GridControl.InvalidRowHandle)
            {
                int rowHandle = gridView.LocateByValue("MEEmrDocumentID", preDocumentId);
                if (rowHandle == DevExpress.XtraGrid.GridControl.InvalidRowHandle)
                {
                    //neu ko tim thay to tren cay benh an, kha nang dang bi filter
                    gridView.ActiveFilter.Clear();
                    rowHandle = gridView.LocateByValue("MEEmrDocumentID", preDocumentId);
                }
                if (rowHandle != DevExpress.XtraGrid.GridControl.InvalidRowHandle)
                {
                    gridView.FocusedRowHandle = rowHandle;
                    var newDoc = gridView.GetRow(rowHandle) as MEEmrDocumentsInfo;
                    _entity.InvalidateModuleObject(newDoc);
                }
                else
                {
                    _entity.SetDefaultModuleObject(TableName.MEEmrDocumentsTableName);
                    this._richEditCtrl.CreateNewDocument(false);
                }
            }
            else
            {
                _entity.SetDefaultModuleObject(TableName.MEEmrDocumentsTableName);
                this._richEditCtrl.CreateNewDocument(false);
            }

            meEmr.MEEmrDocumentViewCount = _entity.MEEmrDocumentsList.Count;
            meEmr.MEEmrDocumentCount = _emrDocumentCtrl.GetCountByEmrId(meEmr.MEEmrID);
            ShowInitAskControls(meEmr.MEEmrDocumentCount == 0);
            //cho phep thay doi loai benh an neu chua co to benh an nao
            Controls["fld_lkeFK_MEEmrTypeID"].Enabled = (meEmr.MEEmrDocumentCount == 0);
        }
        private void InvalidateEmrDocumentListForAdmin(MEEmrsInfo emr)
        {
            //Get all for admin
            var listDocs = _emrDocumentCtrl.GetListBusinessObjects<MEEmrDocumentsInfo>(_emrDocumentCtrl.GetByMEEmrID(emr.MEEmrID));
            var orderedDocs = _emrDocumentSortHelper.SortDocumentTree(listDocs, emr);
            _entity.MEEmrDocumentsList.Invalidate(orderedDocs);
            _entity.MEEmrDocumentsList.GridView.FocusedRowHandle = -1;
        }
        private void ReorderDocumentList(BOSList<MEEmrDocumentsInfo> list, MEEmrsInfo emr, MEEmrDocumentsInfo focusDoc)
        {
            var gridView = list.GridView;
            var orderedDocs = _emrDocumentSortHelper.SortDocumentTree(list, emr);
            list.Invalidate(orderedDocs);
            gridView.FocusedRowHandle = -1;
            int rowHandle = gridView.LocateByValue("MEEmrDocumentID", focusDoc.MEEmrDocumentID);
            if (rowHandle == DevExpress.XtraGrid.GridControl.InvalidRowHandle)
            {
                //neu ko tim thay to tren cay benh an, kha nang dang bi filter
                gridView.ActiveFilter.Clear();
                rowHandle = gridView.LocateByValue("MEEmrDocumentID", focusDoc.MEEmrDocumentID);
            }
            if (rowHandle != DevExpress.XtraGrid.GridControl.InvalidRowHandle)
            {
                gridView.FocusedRowHandle = rowHandle;
                focusDoc = gridView.GetRow(rowHandle) as MEEmrDocumentsInfo;
                _entity.InvalidateModuleObject(focusDoc);
            }
            else
            {
                MessageBox.Show("Nội dung tờ bệnh án đã được lưu. " +
                    "\nTuy nhiên tờ bệnh án bạn đang thao tác [không còn hiển thị] trên cây bệnh án. " +
                    "\nCó thể đã bị xóa/ẩn. Vui lòng thao tác trên tờ khác.", "TỜ BỆNH ÁN KHÔNG TÌM THẤY TRÊN CÂY", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        public override bool BeforeClosing()
        {
            var ok = base.BeforeClosing();
            if (ok)
                if (this._richEditCtrl.Modified)
                {
                    var confirm = MessageBox.Show("Bạn có muốn lưu thay đổi?", "Nội dung tờ bệnh án đã thay đổi", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    if (confirm == DialogResult.Yes)
                    {
                        var command = this._richEditCtrl.CreateCommand(RichEditCommandId.FileSave);
                        command.Execute();
                        ok = true;
                    }
                    else if (confirm == DialogResult.Cancel)
                    {
                        ok = false;
                    }
                }
            return ok;
        }
        public override void BeforeClose()
        {
            base.BeforeClose();
            try
            {
                // Fix 1743
                ParentScreen.SearchContainer.Visibility = DockVisibility.Hidden;
                if (!string.IsNullOrEmpty(_receiveChannel))
                {
                    this._hisChannelListener.ReleaseHandle();
                }
                this.SaveUserConfigs();
                LogFile(true);
                foreach (var thread in _backgroundJobThreads)
                {
                    if (thread.IsAlive)
                        thread.Abort();
                }
            }
            catch (Exception ex)
            {
                _sysHelper.LogTxt("error", ex.ToString());
            }
        }
        internal bool IsExist(int emrID)
        {
            if (!_emrCtrl.IsExist(emrID))
            {
                MessageBox.Show("Bệnh án không còn tồn tại. Có thể đã bị xóa hoặc trộn vào bệnh án khác.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
            return true;
        }
        internal void InvalidateDocument(MEEmrDocumentsInfo document, bool mustDownload = true, bool asyncMode = false)
        {
            if (!IsExist(document.FK_MEEmrID)) return;
            if (mustDownload) BOSProgressBar.Start("Đang tải xuống và mở tập tin");
            else BOSProgressBar.Start("Đang mở tập tin");
            this._dpnViewPdf.Visible = false;
            this._dpnRichEdit.Visible = false;
            try
            {
                BackgroudLoadDocumentRelativeThings(document);
                var objEmr = _emrCtrl.GetObjectByID(document.FK_MEEmrID) as MEEmrsInfo;
                if (objEmr.MEEmrTypeProfile != EmrTypeProfile.Patient.ToString())
                {
                    foreach (var item in _entity.MEEmrDocumentsList)
                    {
                        if (item.MEEmrDocumentID == document.MEEmrDocumentID) continue;
                        ClearCurrentEditingUser(item);
                    }
                }
                else
                {
                    foreach (var item in _entity.MEEmrPatientDocumentsList)
                    {
                        if (item.MEEmrDocumentID == document.MEEmrDocumentID) continue;
                        ClearCurrentEditingUser(item);
                    }
                }

                _msgNotification.Text = string.Empty;
                if (document.MEEmrDocumentStatus == EmrDocumentStatus.Closed.ToString()
                    || document.MEEmrDocumentFileExt == EmrDocumentFileExtention.pdf.ToString())
                {
                    this._dpnViewPdf.Visible = true;
                    this.DockManager.ActivePanel = this._dpnViewPdf;
                    _entity.InvalidateModuleObject(document);
                    OpenEmrPdf(document.FK_MEEmrID, document.MEEmrDocumentFile, mustDownload);
                    return;
                }

                this._dpnRichEdit.Visible = true;
                this.DockManager.ActivePanel = this._dpnRichEdit;
                if (this.TakeEditingPermission(document, IsEmrReadOnly()))
                {
                    document = SetCurrentEditingUser(document);
                    _entity.InvalidateModuleObject(document);
                    if (!true)
                    {
                        OpenEmrTemplate(document.MEEmrDocumentNo);
                        Dictionary<string, object> data = ConvertXMLToDic(document);
                        BindingDataToEmrDocument(data, document.MEEmrDocumentGuid, string.Empty);
                    }
                    else
                        OpenEmrDocument(document.FK_MEEmrID, document.MEEmrDocumentFile, document.MEEmrDocumentStatus == EmrDocumentStatus.Closed.ToString(), mustDownload, asyncMode);
                }
                else
                {
                    _entity.InvalidateModuleObject(document);
                    OpenEmrDocument(document.FK_MEEmrID, document.MEEmrDocumentFile, true, mustDownload, asyncMode);
                }
                if (IsEmrReadOnly())
                    _msgNotification.Text = "Tờ bệnh án được mở ở chế độ CHỈ ĐỌC";


                if (_entity.METemplate != null && _entity.METemplate.METemplateID == document.FK_METemplateID && _entity.METemplate.METemplateHighlightDataTag)
                    _richEditCtrl.Options.Fields.HighlightMode = FieldsHighlightMode.Always;
                else
                    _richEditCtrl.Options.Fields.HighlightMode = FieldsHighlightMode.Never;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                BOSProgressBar.Close();
            }
        }
        private void BackgroudLoadDocumentRelativeThings(MEEmrDocumentsInfo document)
        {
            _entity.METemplatePrimaryHeader = null;
            //uthv khong bat dong bo o day dc nua vi co rang buoc cac the ko duoc edit
            InitDocumentSession(document.FK_METemplateID);

            //bat dong bo
            AppMemCache.InitDocumentSessionAsync(document.MEEmrDocumentID, document.FK_METemplateID);

            if (_sysHelper.AllowThread())
            {
                System.Threading.Thread thInitRequestParamPool = new System.Threading.Thread(() => InitRequestParamPool(document));
                thInitRequestParamPool.Start();
            }
            else
            {
                InitRequestParamPool(document);
            }

            if (_sysHelper.AllowThread())
            {
                //TODO move to AppMemCache
                System.Threading.Thread thread2 = new System.Threading.Thread(() =>
                {
                    _entity.METemplate = this._templateCtrl.GetObjectByID(document.FK_METemplateID) as METemplatesInfo;
                    _entity.METemplatePrimaryHeader = this._templateCtrl.GetFirstTemplateByParentAndType(TemplateType.PrimaryHeader.ToString(), _entity.METemplate.METemplateID);
                });
                thread2.Start();
            }
            else
            {
                _entity.METemplate = this._templateCtrl.GetObjectByID(document.FK_METemplateID) as METemplatesInfo;
                _entity.METemplatePrimaryHeader = this._templateCtrl.GetFirstTemplateByParentAndType(TemplateType.PrimaryHeader.ToString(), _entity.METemplate.METemplateID);
            }
            if (_sysHelper.AllowThread())
            {
                System.Threading.Thread thInvalidateChartList = new System.Threading.Thread(() => InvalidateChartList(document));
                thInvalidateChartList.Start();
            }
            else
            {
                InvalidateChartList(document);
            }
        }
        private void InvalidateChartList(MEEmrDocumentsInfo document)
        {
            try
            {
                var ctrl = new METemplateChartsController();
                _entity.TemplateChartList = ctrl.GetListBusinessObjects<METemplateChartsInfo>(ctrl.GetAllDataByForeignColumn("FK_METemplateID", document.FK_METemplateID));
            }
            catch (Exception ex)
            {
                _sysHelper.LogTxt("error", ex.ToString());
            }
        }
        private void InitDocumentSession(int templateId)
        {
            AppMemCache.InitDocumentSession(templateId);
            _emrDocumentHelper.SetTemplateParamList(
                AppMemCache.GetTemplateParams(templateId),
                AppMemCache.GetTemplateParamsDictPath(templateId)
             );
        }

        /// <summary>
        /// lay quyen edit tai lieu nay
        /// neu ko lay dc quyen thi mo bang mode read only
        /// </summary>
        /// <param name="mEEmrDocumentsInfo"></param>
        /// <returns></returns>
        private bool TakeEditingPermission(MEEmrDocumentsInfo mEEmrDocumentsInfo, bool readOnly = false)
        {
            if (readOnly)
                return false;
            var mETemplateList = _entity.METemplateList;
            if (mETemplateList.Where(x => x.METemplateNo == mEEmrDocumentsInfo.MEEmrDocumentNo).Count() == 0)
            {
                var msg = $"Tờ bệnh án được mở ở chế độ CHỈ ĐỌC vì người dùng thuộc nhóm không có quyền trên mẫu bệnh án này.";
                _msgNotification.Text = msg;
                return false;
            }
            var machineInfo = _hostName + "/" + _ipAddress;
            var owner = this._emrDocumentCtrl.TakeEditingPermission(mEEmrDocumentsInfo.MEEmrDocumentID, BOSApp.CurrentUsersInfo.FK_HREmployeeID, _macAddress, machineInfo);
            if (string.IsNullOrEmpty(owner))
                return true;
            else
            {
                var msg = $"Tờ bệnh án được mở ở chế độ CHỈ ĐỌC vì đang được soạn bởi {owner}.";
                _msgNotification.Text = msg;
                if (_useMessageBox)
                    MessageBox.Show(msg, "ĐANG MỞ Ở CHẾ ĐỘ CHỈ ĐỌC", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, (MessageBoxOptions)0x40000);

            }
            return false;

        }
        public override void ActionNew()
        {
            if (!IsNothingToSaveDocumentContent()) return;
            base.ActionNew();
            this.ActivateScreen("DMMEEMR101");
            var emr = this._entity.MainObject as MEEmrsInfo;
            emr.MEEmrStatus = EmrStatus.InProgress.ToString();
            emr.MEEmrCreatedDate = DateTime.Now;
            emr.FK_HRDepartmentID = BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID;
            emr.FK_HREmployeeCreatedID = BOSApp.CurrentEmployeesInfo.HREmployeeID;
            _entity.SetDefaultModuleObject(TableName.MEEmrTransferHistoriesTableName);

            SetStatusControl(emr.MEEmrID);
            EnableEmrField(true);
        }
        public override int ActionSave()
        {
            var emr = _entity.MainObject as MEEmrsInfo;
            if (IsEmrReadOnly(true))
            {
                MessageBox.Show("Bệnh án đã đóng. Không thể lưu.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 0;
            }
            if (emr.FK_MEEmrTypeID <= 0)
            {
                MessageBox.Show("Loại bệnh án không được để trống", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 0;
            };
            foreach (var share in _entity.MEEmrShareList)
            {
                if ((share.MEEmrShareHistoryToDate - share.MEEmrShareHistoryFromDate).TotalMinutes < 15)
                {
                    MessageBox.Show("Thời gian chia sẻ không hợp lệ. Tối thiểu 15 phút", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return 0;
                };
            }
            var isNew = false;
            var autoMerge = false;
            var sourcesMerge = new List<MEEmrsInfo>();
            if (emr.MEEmrID <= 0) isNew = true;
            if (isNew)
            {
                // uthv 17/04/2019 TWCT kiem tra moi benh nhan chi co 1 benh an ngoai tru dang mo, muon tao benh an ngoai tru moi phai dong benh an hien tai
                // uthv 22/01/2010 DKLK cho phep benh nhan co nhieu benh an ngoai tru
                var type = _emrTypeCtrl.GetObjectByID(emr.FK_MEEmrTypeID) as MEEmrTypesInfo;
                if (type.MEEmrTypeOneEmrOnePhase)
                {
                    // xuantm 02/04/2021 kiểm tra chỉ 1 bệnh án 1 đợt điều trị.
                    // auto trộn các bệnh án cũ vào bệnh án mới
                    sourcesMerge = _emrCtrl.GetAllOpeningEmrByType(emr.FK_MEPatientID, type.MEEmrTypeID).Where(m => m.MEEmrStatus.Equals(EmrStatus.InProgress.ToString())).ToList();
                    if (sourcesMerge.Count() > 0)
                    {
                        autoMerge = true;
                    }
                }
                else if (type.MEEmrTypeMaxCountPerPatient > 0)
                {
                    string customSp = _entity.GetConfigCustomRuleForNewEmr();
                    if (string.IsNullOrEmpty(customSp))
                    {
                        var openings = _emrCtrl.GetAllOpeningEmrByType(emr.FK_MEPatientID, type.MEEmrTypeID);
                        if (openings.Count >= type.MEEmrTypeMaxCountPerPatient)
                        {
                            var nos = openings.Select(e => e.MEEmrNo).ToList();
                            MessageBox.Show($"Bệnh nhân này có [{openings.Count}] bệnh án [{type.MEEmrTypeName}]: \n - {string.Join("\n - ", nos)} đang mở. " +
                                $"\n\nChỉ được tạo [{type.MEEmrTypeMaxCountPerPatient}] bệnh án [{type.MEEmrTypeName}] đang mở cho bệnh nhân.",
                                 $"BỆNH NHÂN NÀY CÓ BỆNH ÁN [{type.MEEmrTypeName}] ĐANG MỞ",
                                 MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            return 0;
                        }
                    }
                    else
                    {
                        var openings = _emrCtrl.GetAllOpeningEmrByTypeCustomSp(customSp, emr.FK_MEPatientID, type.MEEmrTypeNo);
                        if (openings.Count >= type.MEEmrTypeMaxCountPerPatient)
                        {
                            var nos = openings.Select(e => e.MEEmrFullName).ToList();
                            MessageBox.Show($"Bệnh nhân này có [{openings.Count}] bệnh án: \n - {string.Join("\n - ", nos)} đang mở. " +
                                $"\n\nKhông thể tạo thêm bệnh án [{type.MEEmrTypeName}] cho bệnh nhân.",
                                 $"BỆNH NHÂN NÀY CÓ BỆNH ÁN ĐANG MỞ",
                                 MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            return 0;
                        }
                    }
                }
                else
                {
                    if (!type.MEEmrTypeIsTmp)
                    {
                        var maxCountAll = _entity.GetConfigMaxCountEmrPerPatient("ALL");
                        //maxCountAll lớn hơn 0 thì mới xét
                        if (maxCountAll > 0)
                        {
                            DialogResult confirm = DialogResult.None;
                            string lastestNo = string.Empty;
                            var openingEmrs = _emrCtrl.GetAllOpeningEmrWithoutMaxCountPerPatient(emr.FK_MEPatientID);

                            var outNos = openingEmrs.Where(e => e.MEEmrTypeProfile == EmrTypeProfile.Out.ToString()).Select(e => e.MEEmrNo).ToArray();
                            var inNos = openingEmrs.Where(e => e.MEEmrTypeProfile == EmrTypeProfile.In.ToString()).Select(e => e.MEEmrNo).ToArray();

                            if (openingEmrs.Count >= maxCountAll)
                            {
                                lastestNo = openingEmrs.OrderByDescending(e => e.MEEmrID).First().MEEmrNo;
                                confirm = MessageBox.Show($"Bệnh nhân này có [{openingEmrs.Count}] bệnh án đang mở:" +
                                            $"{(outNos.Length > 0 ? "\nNGOẠI TRÚ\n - " + string.Join("\n - ", outNos) : string.Empty)}" +
                                            $"{(inNos.Length > 0 ? "\nNỘI TRÚ\n - " + string.Join("\n - ", inNos) : string.Empty)}" +
                                            $"\n\nBệnh nhân chỉ có tối đa [{maxCountAll}] bệnh án đang mở" +
                                            $"\nOK để xem bệnh án [{lastestNo}]" +
                                            $"\nCANCLE để hủy thao tác",
                                            "VƯỢT QUÁ SỐ BỆNH ÁN TỐI ĐA CHO PHÉP",
                                            MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                            }
                            else
                            {
                                if (type.MEEmrTypeProfile == EmrTypeProfile.Out.ToString())
                                {
                                    // check số lượng bệnh án ngoại trú đang mở KHÔNG BAO GỒM các bệnh án có ràng buột MEEmrTypeMaxCountPerPatient
                                    if (outNos.Length > 0)
                                    {
                                        var maxCount = _entity.GetConfigMaxCountEmrPerPatient(EmrTypeProfile.Out.ToString());
                                        lastestNo = openingEmrs.Where(e => e.MEEmrTypeProfile == EmrTypeProfile.Out.ToString()).OrderByDescending(e => e.MEEmrID).First().MEEmrNo;
                                        if (maxCount == 0 || outNos.Length < maxCount)
                                        // maxCount = 0 là ko có cấu hình => là tạo bao nhiêu bệnh án đều được chỉ cảnh báo
                                        {
                                            confirm = MessageBox.Show($"Bệnh nhân này có [{outNos.Length}] bệnh án NGOẠI TRÚ đang mở:" +
                                                                        "\n - " + string.Join("\n - ", outNos) +
                                                                        $"\n\nVẪN TIẾP TỤC TẠO BỆNH ÁN MỚI?" +
                                                                        $"\n\nYES để tạo bệnh án mới" +
                                                                        $"\nNO để xem bệnh án [{lastestNo}]" +
                                                                        $"\nCANCLE để hủy thao tác",
                                                                        $"BỆNH NHÂN NÀY CÓ BỆNH ÁN ĐANG MỞ",
                                                                        MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                                        }
                                        else
                                        {
                                            confirm = MessageBox.Show($"Bệnh nhân này có [{outNos.Length}] bệnh án NGOẠI TRÚ đang mở:" +
                                                                       "\n - " + string.Join("\n - ", outNos) +
                                                                       $"\n\nBệnh nhân chỉ có tối đa [{maxCount}] bệnh án ngoại trú đang mở" +
                                                                        $"\nOK để xem bệnh án [{lastestNo}]" +
                                                                        $"\nCANCLE để hủy thao tác",
                                                                        "VƯỢT QUÁ SỐ BỆNH ÁN TỐI ĐA CHO PHÉP",
                                                                        MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                                        }
                                    }
                                }
                                else if (type.MEEmrTypeProfile == EmrTypeProfile.In.ToString())
                                {
                                    //11/06/2020 DKLK cho phép nhiều bệnh án nội trú mở cùng lúc
                                    //11/06/2020 PSHN chỉ cho phép 01 bệnh án nội trú mở cùng lúc
                                    var maxCount = _entity.GetConfigMaxCountEmrPerPatient(EmrTypeProfile.In.ToString());
                                    if (maxCount > 0)
                                    {
                                        if (inNos.Length >= maxCount)
                                        {
                                            lastestNo = openingEmrs.Where(e => e.MEEmrTypeProfile == EmrTypeProfile.In.ToString()).OrderByDescending(e => e.MEEmrID).First().MEEmrNo;
                                            confirm = MessageBox.Show($"Bệnh nhân này có [{inNos.Length}] bệnh án NỘI TRÚ đang mở:" +
                                                                      "\n - " + string.Join("\n - ", inNos) +
                                                                       $"\n\nChỉ được tạo [{maxCount}] bệnh án nội trú đang mở cho bệnh nhân." +
                                                                       $"\nOK để xem bệnh án [{lastestNo}]" +
                                                                       $"\nCANCLE để hủy thao tác",
                                                                       "VƯỢT QUÁ SỐ BỆNH ÁN TỐI ĐA CHO PHÉP",
                                                                       MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                                        }
                                    }
                                }
                            }
                            if (confirm == DialogResult.No || confirm == DialogResult.OK)
                            {
                                if (!string.IsNullOrEmpty(lastestNo))
                                {
                                    ActionCancel();
                                    SetSearchParamValue(ParentScreen.SearchQuickContainer, "fld_txtMEEmrNo", lastestNo);
                                    SetSearchParamValue(ParentScreen.SearchQuickContainer, "chkSearchByEmrCode", true);
                                    QuickSearch();
                                }
                                return 0;
                            }
                            else if (confirm == DialogResult.Cancel)
                            {
                                return 0;
                            }
                        }
                    }
                }
            }

            var tran = _entity.ModuleObjects[TableName.MEEmrTransferHistoriesTableName] as MEEmrTransferHistoriesInfo;
            var caseNo = tran.MEEmrTransferHistoryNo;
            emr = CopyPatientInfoToEmr(emr);

            int mEEmrID = base.ActionSave();
            if (mEEmrID <= 0) return 0;
            emr.MEEmrID = mEEmrID;
            if (isNew)
            {
                HistoryEmr(_entity.MainObject as MEEmrsInfo, cstObjectHistoryActionNew, "Khởi tạo bệnh án từ EMR");

                InitEmr(_entity.MainObject as MEEmrsInfo, caseNo, true);

                if (autoMerge)
                {
                    foreach (var source in sourcesMerge)
                    {
                        MergeEmrAction(source, _entity.MainObject as MEEmrsInfo);
                    }
                }
            }
            return mEEmrID;
        }
        private void InitEmr(MEEmrsInfo emr, string caseNo, bool willTranfer)
        {
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                BOSProgressBar.Start("Đang khởi tạo bệnh án");

                PrintMgsLog("BAT-DAU-KHOI-TAO-BENH-AN", emr.MEEmrNo);
                CreateDocumentsFromEmrType(emr, _entity.MEEmrDocumentsList);
                var backupDocs = _entity.MEEmrDocumentsList.Select(o => new MEEmrDocumentsInfo()
                {
                    MEEmrDocumentFile = o.MEEmrDocumentFile,
                    MEEmrDocumentJson = o.MEEmrDocumentJson
                }).ToList();

                _entity.UpdateModuleObjectBindingSource(TableName.MEEmrDocumentsTableName);
                _entity.MEEmrDocumentsList.SaveItemObjects();

                foreach (var doc in _entity.MEEmrDocumentsList)
                {
                    var documentNo = doc.MEEmrDocumentNo;
                    if (string.IsNullOrEmpty(documentNo))
                    {
                        // Get template no
                        var template = this._templateCtrl.GetObjectByID(doc.FK_METemplateID) as METemplatesInfo;
                        documentNo = template.METemplateNo;
                    }

                    BackgroundEmrDocumentValidates(emr, doc.FK_METemplateID, doc.MEEmrDocumentID);
                    //HistoryEmr(emr, cstObjectHistoryActionChange, $"{documentNo} cập nhật thành công tờ bệnh án. [EMR-APP tạo từ mẫu bệnh án]");
                }

                this.InvalidateEmrDocumentList(emr.MEEmrID);
                foreach (var doc in _entity.MEEmrDocumentsList)
                {   //save to mongo db
                    doc.MEEmrDocumentJson = backupDocs.Where(o => o.MEEmrDocumentFile == doc.MEEmrDocumentFile).FirstOrDefault()?.MEEmrDocumentJson;
                    InsertMongoDocument(emr, doc);
                }
                _entity.MEEmrDocumentsList.SaveItemObjects();

                if (willTranfer)
                {
                    //Thêm thông tin chuyển khoa
                    var tran = _entity.ModuleObjects[TableName.MEEmrTransferHistoriesTableName] as MEEmrTransferHistoriesInfo;
                    tran.MEEmrTransferHistoryID = 0;
                    tran.FK_HRDepartmentFromID = 0;
                    tran.MEEmrTransferHistoryDate = DateTime.Now;
                    tran.FK_HREmployeeFromID = 0;
                    tran.FK_HREmployeeToID = BOSApp.CurrentEmployeesInfo.HREmployeeID;
                    tran.MEEmrTransferHistoryRemark = "Khởi tạo bệnh án";
                    tran.FK_MEEmrID = emr.MEEmrID;
                    tran.MEEmrTransferHistoryCurrent = true;
                    tran.FK_HRDepartmentToID = BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID;
                    if (string.IsNullOrEmpty(caseNo))
                        tran.MEEmrTransferHistoryNo = emr.MEEmrNo + ".1";
                    else
                        tran.MEEmrTransferHistoryNo = caseNo;

                    _entity.MEEmrTranfersList.AddObjectToList();
                    _entity.MEEmrTranfersList.SaveItemObjects();

                    var dept = _departmentCtrl.GetObjectByID(BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID) as HRDepartmentsInfo;
                    CreateShareAll(emr.MEEmrID, dept);
                }

                PrintMgsLog("KHOI-TAO-BENH-AN-THANH-CONG", emr.MEEmrNo);

                this.ActivateScreen("DMMEEMR100");
                this.DockManager.ActivePanel = this._dpnRichEdit;
                this.OpenGuidance();

                if (_sysHelper.AllowThread())
                {
                    //uthv co the chay am tham ben duoi ma ko anh huong
                    var thSendEmrToHis = new System.Threading.Thread(() => SendEmrToHis());
                    thSendEmrToHis.Start();
                }
                else
                {
                    SendEmrToHis();
                }
            }
            catch (Exception ex)
            {
                HistoryEmr(emr, cstObjectHistoryActionError, $"{StringExt.Truncate(ex.ToString(), 256)}... [EMR-APP tạo từ mẫu bệnh án]");
                throw;
            }
            finally
            {
                BOSProgressBar.Close();
                Cursor.Current = Cursors.Default;
            }
        }

        private void CreateDocumentsFromEmrType(MEEmrsInfo emr, BOSList<MEEmrDocumentsInfo> documentList)
        {
            CreateEmrDir(emr.MEEmrID);
            var templates = AppMemCache.GetEmrTypeTemplatesFromDict(emr.FK_MEEmrTypeID)
                   .Where(o => o.MEEmrTypeTemplateRequired).OrderBy(o => o.MEEmrTypeTemplateOrder).ToList();
            foreach (var item in templates)
            {
                var template = this._templateCtrl.GetObjectByID(item.FK_METemplateID) as METemplatesInfo;
                if (template != null)
                {
                    var order = AppMemCache.GetTemplateIndexById(item.FK_METemplateIndexID);

                    BOSProgressBar.SetText("Khởi tạo: " + template.METemplateName);
                    _entity.SetDefaultModuleObject(TableName.MEEmrDocumentsTableName);
                    var doc = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
                    doc.FK_METemplateID = template.METemplateID;
                    doc.MEEmrDocumentCreatedDate = DateTime.Now;
                    doc.MEEmrDocumentStatus = EmrDocumentStatus.InProgress.ToString();
                    doc.MEEmrDocumentFile = template.METemplateNo + DateTime.Now.ToString("_ddMMyyyy_HHmmssffff_") + Guid.NewGuid().ToString().Replace('-', '_');
                    doc.MEEmrDocumentNo = template.METemplateNo;
                    doc.MEEmrDocumentCode = template.METemplateNo;
                    doc.MEEmrDocumentGuid = template.METemplateGuid;
                    doc.MEEmrDocumentOrder = (order != null ? order.METemplateIndexOrder : 999);
                    doc.MEEmrDocumentGroup = (order != null ? order.METemplateIndexName : "Khác");
                    doc.MEEmrDocumentSubOrder = GetDocumentSubOrder(emr.MEEmrID, doc.MEEmrDocumentGroup);
                    doc.FK_MEEmrID = emr.MEEmrID;
                    doc.FK_EditingUserID = 0;
                    doc.MEEmrDocumentHoldMachineIp = string.Empty;
                    doc.MEEmrDocumentHoldMachineMac = string.Empty;
                    doc.FK_HRDepartmentID = BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID;
                    doc.MEEmrDocumentFileExt = EmrDocumentFileExtention.docx.ToString();
                    doc.FK_HREmployeeCreatedID = BOSApp.CurrentEmployeesInfo.HREmployeeID;
                    BackgroudLoadDocumentRelativeThings(doc);
                    this.CreateDocumentFileFromTemplate(emr, doc, template.METemplateGuid);
                    this._richEditCtrl.CreateNewDocument(false);
                    documentList.AddObjectToList();
                }
            }
        }

        private int GetDocumentSubOrder(int emrId, string group)
        {
            return _emrDocumentCtrl.GetDocumentSubOrder(emrId, group);
        }
        private void CreateShareAll(int mEEmrID, HRDepartmentsInfo dept)
        {
            //chi nhung khoa chia se benh an moi chia se cho moi nguoi
            if (dept.HRDepartmentEmrShared)
            {
                _entity.SetDefaultModuleObject(TableName.MEEmrShareHistoriesTableName);
                var share = _entity.ModuleObjects[TableName.MEEmrShareHistoriesTableName] as MEEmrShareHistoriesInfo;
                share.FK_MEEmrID = mEEmrID;
                share.FK_HRDepartmentID = 0; //chia se tat ca
                share.FK_HREmployeeID = 0;
                share.FK_HREmployeeShareByID = BOSApp.CurrentEmployeesInfo.HREmployeeID;
                share.MEEmrShareHistoryActive = true;
                share.MEEmrShareHistoryDate = DateTime.Now;
                share.MEEmrShareHistoryFromDate = DateTime.Now.Date;
                share.MEEmrShareHistoryToDate = DateTime.Now.Date.AddYears(100);
                share.MEEmrShareHistoryRemark = "Khởi tạo bệnh án";
                share.MEEmrShareHistoryMode = EmrShareHistoryMode.Edit.ToString();
                _entity.MEEmrShareList.AddObjectToList();
                _entity.MEEmrShareList.SaveItemObjects();
            }
        }
        private void CreateRemainShare(int mEEmrID, int departmentID)
        {
            var dept = _departmentCtrl.GetObjectByID(departmentID) as HRDepartmentsInfo;
            if ((decimal)dept.HRDepartmentAutoShareAfterTranf > 0)
            {
                _entity.SetDefaultModuleObject(TableName.MEEmrShareHistoriesTableName);
                var share = _entity.ModuleObjects[TableName.MEEmrShareHistoriesTableName] as MEEmrShareHistoriesInfo;
                share.FK_MEEmrID = mEEmrID;
                share.FK_HRDepartmentID = departmentID; //chia se tat ca
                share.FK_HREmployeeID = 0;
                share.FK_HREmployeeShareByID = BOSApp.CurrentEmployeesInfo.HREmployeeID;
                share.MEEmrShareHistoryActive = true;
                share.MEEmrShareHistoryDate = DateTime.Now;
                share.MEEmrShareHistoryFromDate = DateTime.Now;
                share.MEEmrShareHistoryToDate = DateTime.Now.AddHours((double)dept.HRDepartmentAutoShareAfterTranf);
                share.MEEmrShareHistoryRemark = $"Chia sẻ {dept.HRDepartmentAutoShareAfterTranf}h sau khi chuyển khoa";
                _entity.MEEmrShareList.AddObjectToList();
                _entity.MEEmrShareList.SaveItemObjects();
            }
        }

        public override void ActionEdit()
        {
            if (!IsNothingToSaveDocumentContent()) return;

            //hard set for fix bug https://trello.com/c/gB6irgIm
            _richEditCtrl.Modified = false;

            base.ActionEdit();
            EnableEmrField(false);
        }
        private void EnableEmrField(bool status)
        {
            Controls["fld_txtMEPatientNo"].Enabled = status;
            Controls["fld_tbnSearchPatientFromHis"].Enabled = status;
            Controls["fld_tbnSearchPatientLocal"].Enabled = status;
            Controls["fld_txtMEEmrNo2"].Enabled = status;
            Controls["fld_dteAACreatedDate"].Enabled = status;

            //https://trello.com/c/43cTg0FB/639-cho-ph%C3%A9p-s%E1%BB%ADa-lo%E1%BA%A1i-ba-trong-module-ba%C4%91t-khi-b%E1%BB%87nh-%C3%A1n-ch%C6%B0a-c%C3%B3-t%E1%BB%9D-n%C3%A0o
            if ((_entity.MainObject as MEEmrsInfo).MEEmrDocumentCount > 0)
                Controls["fld_lkeFK_MEEmrTypeID"].Enabled = false;
            else
                Controls["fld_lkeFK_MEEmrTypeID"].Enabled = true;

            Controls["fld_medMEEmrDesc1"].Enabled = status;
        }
        private void CreateDocumentFileFromTemplate(MEEmrsInfo emr, MEEmrDocumentsInfo documentFile, string mETemplateGuid)
        {
            var fileName = DownloadTemplate(documentFile);
            if (string.IsNullOrEmpty(fileName)) return;

            string toFileName = string.Format(@"{0}\Emr\{1}\{2}.docx", _documentPath, emr.MEEmrID, documentFile.MEEmrDocumentFile);
            File.Copy(fileName, toFileName);
            _richEditCtrl.LoadDocument(toFileName);
            this.InitFirstValueForEmrDocument(emr, documentFile.FK_METemplateID, mETemplateGuid);
            var doc = this._richEditCtrl.Document;
            doc.EndUpdate();
            _richEditCtrl.SaveDocument(toFileName, DocumentFormat.OpenXml);
            using (OfficeOpenXmlCrypto.OfficeCryptoStream stream = OfficeOpenXmlCrypto.OfficeCryptoStream.Open(toFileName, this._emrDocumentHelper.ShareEmrPassword))
            {
                stream.Password = this._emrDocumentHelper.ShareEmrPassword;
                stream.Save();
            }
            var item = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            item.MEEmrDocumentJson = this.ParserDocumentToJson();

            var data = JsonConvert.DeserializeObject(item.MEEmrDocumentJson) as JToken;
            item = ExtractDataFromContentToDocumentInfo(item, data);
            emr = UpdateEmrInfoByDocumentContent(emr, data);

            _ftpFileMng.UploadFile($"/Emr/{item.FK_MEEmrID}/", item.MEEmrDocumentFile + ".docx", toFileName);
        }

        internal void ShowNewEmrDocumentDialog()
        {
            var emr = _entity.MainObject as MEEmrsInfo;
            if (emr.MEEmrStatus == EmrStatus.Initing.ToString())
            {
                MessageBox.Show($"Bệnh án đang được khởi tạo, vui lòng không thêm tờ cho đến khi trạng thái bệnh án là Đang mở", "THÔNG BÁO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var emrDb = _emrCtrl.GetObjectByID(emr.MEEmrID) as MEEmrsInfo;
            if (emrDb.MEEmrStatus == EmrStatus.Initing.ToString())
            {
                MessageBox.Show($"Bệnh án đang được khởi tạo, vui lòng không thêm tờ cho đến khi trạng thái bệnh án là Đang mở", "THÔNG BÁO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (IsEmrReadOnly(true)) return;
            if (!IsNothingToSaveDocumentContent()) return;
            _entity.MENewEmrDocument = null;

            if (_sysHelper.AllowThread())
            {
                var thCreateEmrDir = new System.Threading.Thread(() => CreateEmrDir(emr.MEEmrID));
                thCreateEmrDir.Start();
            }
            else
            {
                CreateEmrDir(emr.MEEmrID);
            }

            var isFilterTemplate = false;
            var emrType = _emrTypeCtrl.GetObjectByID(emr.FK_MEEmrTypeID) as MEEmrTypesInfo;
            if (emrType != null)
            {
                isFilterTemplate = emrType.MEEmrTypeFilterTemplate;
            }
            var mETemplateList = _entity.METemplateList;
            if (isFilterTemplate)
            {
                mETemplateList = this._templateCtrl.GetTemplatesByEmrType(TemplateType.ProgressNote.ToString(), BOSApp.CurrentUserGroupInfo.ADUserGroupID, emr.FK_MEEmrTypeID);
            }
            var gui = new DSMEEMR100(mETemplateList) { Module = this };
            if (gui.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    AddNewEmrDocument(null, asyncMode: true);
                }
                catch (Exception ex)
                {
                    this.InvalidateEmrDocumentList(emr.MEEmrID);
                    this.ClearDocumentSession();
                    MessageBox.Show("Có lỗi khi thêm tờ bệnh án. Xin thử lại.", "CÓ LỖI KHI TẠO TỜ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    PrintMgsLog("LOI-ADD-NEW-EMR-DOCUMENT-MANUAL", ex.ToString());
                }
            }
        }
        private MEEmrDocumentsInfo AddNewDocument(MEEmrsInfo emr, BOSList<MEEmrDocumentsInfo> documentList, Dictionary<string, object> initData)
        {
            BOSProgressBar.Start("Đang khởi tạo tờ bệnh án");
            PrintMgsLog("BAT-DAU-KHOI-TAO-MAU", _entity.MENewEmrDocument.MEEmrDocumentFile);
            var emrTypeTemplates = AppMemCache.GetEmrTypeTemplatesFromDict(emr.FK_MEEmrTypeID);

            var filePath = _entity.MENewEmrDocument.MEEmrDocumentExternalFile;

            var ext = string.IsNullOrEmpty(filePath) ? string.Empty : Path.GetExtension(filePath);
            if (!string.IsNullOrEmpty(filePath))
            {
                var file = new FileInfo(filePath);
                if (!file.Exists)
                {
                    MessageBox.Show("Tập tin không tồn tại ở địa chỉ. " + filePath);
                    return null;
                }
                if (file.Length == 0)
                {
                    MessageBox.Show($"Tập tin [BỊ RỖNG]. Không chứa dữ liệu.", "TẬP TIN ĐÃ CHỌN BỊ LỖI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
                if (file.IsReadOnly)
                {
                    MessageBox.Show($"Tập tin [CHỈ ĐỌC] không thể tải lên.", "TẬP TIN ĐÃ CHỌN BỊ LỖI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
                try
                {
                    //#1043 
                    if (ext.Equals("." + EmrDocumentFileExtention.pdf.ToString()))
                        using (var reader = new PdfReader(filePath)) { };
                    //uthv #1198
                    using (FileStream stream = new FileStream(filePath, FileMode.Open)) { }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Tập tin lỗi không thể tải lên. \n\nChi tiết: " + ex.ToString(), "TẬP TIN ĐÃ CHỌN BỊ LỖI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }
            }
            BOSProgressBar.Start("Đang khởi tạo tờ bệnh án");
            // create new file from template or external file
            if (_entity.MENewEmrDocument.FK_METemplateID > 0)
            {
                var templType = emrTypeTemplates.Where(o => o.FK_METemplateID == _entity.MENewEmrDocument.FK_METemplateID).FirstOrDefault();
                METemplateIndexsInfo order = null;
                if (templType != null)
                    order = AppMemCache.GetTemplateIndexById(templType.FK_METemplateIndexID);

                _entity.SetDefaultModuleObject(TableName.MEEmrDocumentsTableName);
                var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
                document.AACreatedUser = BOSApp.CurrentUser;
                document.FK_METemplateID = _entity.MENewEmrDocument.FK_METemplateID;
                document.MEEmrDocumentCreatedDate = DateTime.Now;
                document.MEEmrDocumentStatus = EmrDocumentStatus.InProgress.ToString();
                document.MEEmrDocumentDesc = _entity.MENewEmrDocument.MEEmrDocumentDesc;
                document.MEEmrDocumentFile = _entity.MENewEmrDocument.MEEmrDocumentFile;
                document.MEEmrDocumentNo = _entity.MENewEmrDocument.MEEmrDocumentNo;
                document.MEEmrDocumentCode = _entity.MENewEmrDocument.MEEmrDocumentCode;
                document.MEEmrDocumentGuid = _entity.MENewEmrDocument.MEEmrDocumentGuid;
                document.FK_MEEmrID = emr.MEEmrID;
                document = SetCurrentEditingUser(document);
                document.FK_HRDepartmentID = BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID;
                document.FK_HREmployeeCreatedID = BOSApp.CurrentEmployeesInfo.HREmployeeID;
                document.MEEmrDocumentOrder = (order != null ? order.METemplateIndexOrder : 999);
                document.MEEmrDocumentGroup = (order != null ? order.METemplateIndexName : "Khác");
                document.MEEmrDocumentSubOrder = GetDocumentSubOrder(emr.MEEmrID, document.MEEmrDocumentGroup);
                string toFileName = string.Empty;
                if (string.IsNullOrEmpty(filePath) || ext.Equals("." + EmrDocumentFileExtention.docx.ToString()))
                {
                    BackgroudLoadDocumentRelativeThings(document);

                    document.MEEmrDocumentFileExt = EmrDocumentFileExtention.docx.ToString();
                    document.MEEmrDocumentInitFileExt = document.MEEmrDocumentFileExt;

                    if (string.IsNullOrEmpty(filePath))
                        this.CopyTemplateToDocument(_entity.MENewEmrDocument.FK_METemplateID, document.MEEmrDocumentFile, document.FK_MEEmrID);
                    else
                        this.CopyExternalFileToDocument(filePath, document.MEEmrDocumentFile, document.FK_MEEmrID);

                    toFileName = string.Format(@"{0}\Emr\{1}\{2}.{3}", _documentPath, document.FK_MEEmrID, document.MEEmrDocumentFile, document.MEEmrDocumentFileExt);
                    this.OpenEmrDocument(document.FK_MEEmrID, document.MEEmrDocumentFile, false, false);

                    PrintMgsLog("BAT-DAU-INIT-GIA-TRI-FILE", toFileName);
                    //to benh an tao tu dong, lay gia tri tu to chinh phai do du lieu truoc khi goi api
                    if (initData != null)
                        BindingDataToEmrDocument(initData, document.MEEmrDocumentGuid, string.Empty);

                    BOSProgressBar.SetText("Điền dữ liệu vào tờ bệnh án");

                    this.InitFirstValueForEmrDocument(emr, document.FK_METemplateID, document.MEEmrDocumentGuid);
                    var doc = this._richEditCtrl.Document;
                    doc.EndUpdate();

                    PrintMgsLog("BAT-DAU-MA-HOA-FILE", toFileName);
                    _richEditCtrl.SaveDocument(toFileName, DocumentFormat.OpenXml);
                    using (OfficeOpenXmlCrypto.OfficeCryptoStream stream =
                        OfficeOpenXmlCrypto.OfficeCryptoStream.Open(toFileName, this._emrDocumentHelper.ShareEmrPassword))
                    {
                        stream.Password = this._emrDocumentHelper.ShareEmrPassword;
                        stream.Save();
                    }
                    PrintMgsLog("KET-THUC-MA-HOA-FILE", toFileName);
                    document.MEEmrDocumentJson = this.ParserDocumentToJson();
                    this._richEditCtrl.Modified = false;

                    // background TDT

                }
                else
                {
                    document.MEEmrDocumentFileExt = filePath.Substring(filePath.LastIndexOf(".") + 1);
                    document.MEEmrDocumentInitFileExt = document.MEEmrDocumentFileExt;

                    toFileName = string.Format(@"{0}\Emr\{1}\{2}.{3}", _documentPath, document.FK_MEEmrID, document.MEEmrDocumentFile, document.MEEmrDocumentFileExt);
                    this.CopyFileToEmr(filePath, toFileName);
                    if (document.MEEmrDocumentFileExt == EmrDocumentFileExtention.pdf.ToString())
                    {

                    }
                    else if (document.MEEmrDocumentFileExt == EmrDocumentFileExtention.jpg.ToString())
                    {

                    }
                }

                BOSProgressBar.SetText("Đang lưu và upload tờ bệnh án");
                PrintMgsLog("BAT-DAU-UPLOAD-FILE", toFileName);
                _ftpFileMng.UploadFile($"/Emr/{document.FK_MEEmrID}/", document.MEEmrDocumentFile + "." + document.MEEmrDocumentFileExt, toFileName);
                PrintMgsLog("KET-THUC-UPLOAD-FILE", toFileName);
                return document;
            }
            //append to current file
            else if (ext.Equals("." + EmrDocumentFileExtention.docx.ToString()))
            {
                var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
                if (document == null)
                {
                    MessageBox.Show("Chọn tập tin (.docx) để thêm vào tờ bệnh án", "Chọn 1 tập tin");
                    return null;
                }
                var doc = this._richEditCtrl.Document;

                doc.InsertDocumentContent(doc.CaretPosition, filePath, DocumentFormat.OpenXml, InsertOptions.KeepSourceFormatting);
                var command = this._richEditCtrl.CreateCommand(RichEditCommandId.FileSave);
                command.Execute();
            }
            PrintMgsLog("KHOI-TAO-MAU-THANH-CONG", _entity.MENewEmrDocument.MEEmrDocumentFile);
            return null;
        }
        private void AddNewEmrDocument(Dictionary<string, object> initData, bool asyncMode = false)
        {
            Cursor.Current = Cursors.WaitCursor;
            var emr = _entity.MainObject as MEEmrsInfo;
            try
            {
                if (_entity.MENewEmrDocument != null)
                {
                    // Check TDT
                    if (!AllowCreateTDT(emr, _entity.MENewEmrDocument.FK_METemplateID))
                    {
                        MessageBox.Show("Tờ điều trị đã tạo gần nhất chưa thực hiện xong. Không thể thêm Tờ điều trị mới. Vui lòng thực hiện tiếp ở Tờ điều trị đã tạo gần nhất", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    var document = this.AddNewDocument(emr, _entity.MEEmrDocumentsList, initData);
                    if (document != null)
                    {
                        var data = JsonConvert.DeserializeObject(document.MEEmrDocumentJson) as JToken;
                        document = ExtractDataFromContentToDocumentInfo(document, data);

                        emr = UpdateEmrInfoByDocumentContent(emr, data);
                        _entity.MEEmrDocumentsList.AddObjectToList();
                        _emrDocumentCtrl.CreateObject(document);
                        //_entity.MEEmrDocumentsList.SaveItemObjects();

                        BackgroundEmrDocumentValidates(emr, document.FK_METemplateID, document.MEEmrDocumentID);

                        var documentCode = getDocumentNo(document);
                        HistoryEmr(emr, "Change", $"{documentCode} thêm tờ bệnh án.");

                        this.InvalidateEmrDocumentList(emr.MEEmrID);

                        var grid = _entity.MEEmrDocumentsList.GridView;
                        for (int i = 0; i < _entity.MEEmrDocumentsList.Count; i++)
                        {
                            if (_entity.MEEmrDocumentsList[i].MEEmrDocumentFile == document.MEEmrDocumentFile)
                            {
                                grid.FocusedRowHandle = grid.GetRowHandle(i);
                                _entity.MEEmrDocumentsList[i].MEEmrDocumentJson = document.MEEmrDocumentJson;
                                document = _entity.MEEmrDocumentsList[i];
                                //save to mongo db
                                InsertMongoDocument(emr, document);
                                document.AAUpdatedUser = BOSApp.CurrentUser;
                                this._emrDocumentCtrl.UpdateObject(document);

                                var documentCodeU = getDocumentNo(document);
                                HistoryEmr(emr, "Change", $"cập nhật thông tin tờ bệnh án [{documentCodeU}].");

                                break;
                            }
                        }
                        InvalidateDocument(document, mustDownload: false, asyncMode: asyncMode);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                BOSProgressBar.Close();
                _entity.MENewEmrDocument = null;
                Cursor.Current = Cursors.Default;
            }
        }
        /// <summary>
        /// Dung cho truong hop sinh to tu dong de tang performance
        /// </summary>
        /// <param name="initData"></param>
        private MEEmrDocumentsInfo AddNewEmrDocumentAuto(Dictionary<string, object> initData)
        {
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                if (_entity.MENewEmrDocument != null)
                {
                    var emr = _entity.MainObject as MEEmrsInfo;
                    var document = this.AddNewDocument(emr, _entity.MEEmrDocumentsList, initData);
                    if (document != null)
                    {
                        var data = JsonConvert.DeserializeObject(document.MEEmrDocumentJson) as JToken;
                        document = ExtractDataFromContentToDocumentInfo(document, data);
                        emr = UpdateEmrInfoByDocumentContent(emr, data);
                        _entity.MEEmrDocumentsList.AddObjectToList();
                        _emrDocumentCtrl.CreateObject(document);

                        //save to mongo db
                        InsertMongoDocument(emr, document);
                        document.AAUpdatedUser = BOSApp.CurrentUser;
                        _emrDocumentCtrl.UpdateObject(document);

                        var documentCodeU = getDocumentNo(document);
                        HistoryEmr(emr, "Change", $"cập nhật thông tin tờ bệnh án [{documentCodeU}].");

                        return document;
                    }
                }
            }
            catch (Exception) { throw; }
            finally
            {
                BOSProgressBar.Close();
                _entity.MENewEmrDocument = null;
                Cursor.Current = Cursors.Default;
            }
            return null;
        }
        /// <summary>
        /// copy content 
        /// </summary>
        /// <param name="mEEmrDocumentExternalFile"></param>
        /// <param name="mEEmrDocumentFile"></param>
        private void CopyExternalFileToDocument(string mEEmrDocumentExternalFile, string mEEmrDocumentFile, int emrId)
        {
            if (!File.Exists(mEEmrDocumentExternalFile))
            {
                MessageBox.Show("File không tồn tại ở địa chỉ. " + mEEmrDocumentExternalFile);
                return;
            }
            string toFileName = string.Format(@"{0}\Emr\{1}\{2}.docx", _documentPath, emrId, mEEmrDocumentFile);
            File.Copy(mEEmrDocumentExternalFile, toFileName);
            using (OfficeOpenXmlCrypto.OfficeCryptoStream stream = OfficeOpenXmlCrypto.OfficeCryptoStream.Open(toFileName))
            {
                stream.Password = this._emrDocumentHelper.ShareEmrPassword;
                stream.Save();
            }
            //string pathTemplate = string.Format(@"{0}\Emr\{1}.docx", _documentPath, mEEmrDocumentFile);
            _ftpFileMng.UploadFile($"/Emr/{emrId}/", mEEmrDocumentFile + ".docx", toFileName);
        }
        private Dictionary<string, object> GetHardParamList(MEEmrsInfo emr)
        {
            var paramList = new Dictionary<string, object>
            {
                { "Day", DateTime.Now.ToString("dd") },
                { "Month", DateTime.Now.ToString("MM") },
                { "Year", DateTime.Now.Year },
                { "Hour", DateTime.Now.Hour },
                { "Minute", DateTime.Now.Minute },
                { "Second", DateTime.Now.Second },
                { "DayMonth", DateTime.Now.ToString("dd/MM", CultureInfo.InvariantCulture) },
                { "MonthYear", DateTime.Now.ToString("MM/yyyy", CultureInfo.InvariantCulture) },
                { "Date", DateTime.Now.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) },
                { "DateTime", DateTime.Now.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture) },
                { "Time", DateTime.Now.ToString("HH:mm") },
                { "TimeSpan", DateTime.Now.ToString("HH:mm:ss") },

                { "day", DateTime.Now.ToString("dd") },
                { "month", DateTime.Now.ToString("MM") },
                { "year", DateTime.Now.Year },
                { "hour", DateTime.Now.Hour },
                { "minute", DateTime.Now.Minute },
                { "second", DateTime.Now.Second },
                { "daymonth", DateTime.Now.ToString("dd/MM", CultureInfo.InvariantCulture) },
                { "monthyear", DateTime.Now.ToString("MM/yyyy", CultureInfo.InvariantCulture) },
                { "date", DateTime.Now.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) },
                { "datetime", DateTime.Now.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture) },
                { "time", DateTime.Now.ToString("HH:mm") },
                { "timespan", DateTime.Now.ToString("HH:mm:ss") },

                { "EmrNo", emr.MEEmrNo },
                { "PatientNo", this._entity.MEPatient.MEPatientNo },
                { "Username", BOSApp.CurrentEmployeesInfo.HREmployeeName }
            };
            return paramList;
        }
        private void InitFirstValueForEmrDocument(MEEmrsInfo emr, int FK_METemplateID, string group)
        {
            var paramList = GetHardParamList(emr);
            BindingDataToEmrDocument(paramList, group, string.Empty);
            var actions = this._templateActionCtrl.GetAllByTemplateID(FK_METemplateID).Where(o => o.MEEmrTemplateActionWhen == EmrTemplateActionWhen.Init.ToString())
                .OrderBy(o => o.MEEmrTemplateActionOrder).ToList();
            foreach (var act in actions)
            {
                var action = this._actionsController.GetObjectByID(act.FK_MEEmrActionID) as MEEmrActionsInfo;
                //TODO cac action yeu cau Range se ko chay dc
                if (action != null)
                {
                    if (_checkSystem)
                    {
                        _sysHelper.LogTxt("information", $"Bắt đầu chạy chức năng {action.MEEmrActionNo}.");
                        var watchAct = Stopwatch.StartNew();
                        CallEmrAction($"{action.MEEmrActionNo}{EmrParam.TagCodeSeparator}{EmrParam.GuidTag}={group}", this._richEditCtrl.Document.Range, cacheTimeOut: act.MEEmrTemplateActionCacheExpire);
                        watchAct.Stop();
                        var elapsedAct = watchAct.ElapsedMilliseconds / 1000.0;
                        _sysHelper.LogTxt("information", $"{elapsedAct} giây. Hoàn tất chạy chức năng {action.MEEmrActionNo}.");
                    }
                    else
                    {
                        CallEmrAction($"{action.MEEmrActionNo}{EmrParam.TagCodeSeparator}{EmrParam.GuidTag}={group}", this._richEditCtrl.Document.Range, cacheTimeOut: act.MEEmrTemplateActionCacheExpire);
                    }
                }
            }
        }
        private void BeforeSaveFileDocument()
        {
            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            // 20201013 khong dung nua, thay bang plugin
            /*foreach (var para in _entity.METemplateParamList)
            {
                if (para.METemplateParamPath.ToLower() == "bmi")
                {
                    _emrDocumentHelper.CalculateParamValueBeforeSave(document.MEEmrDocumentGuid);
                    break;
                }
            }*/
            var doc = _richEditCtrl.Document;
            var renderChart = _entity.TemplateChartList.Where(c => c.METemplateChartRenderAtSave).ToList();
            if (renderChart.Count > 0)
            {
                try
                {
                    BOSProgressBar.Start("Đang cập nhật dữ liệu cho các biểu đồ");
                    var fields = this._emrDocumentHelper.GetAllDataFieldInRangeOrDocument();
                    var json = this._emrParser.ParserFieldsToJson(fields, AppMemCache.GetTemplateParams(document.FK_METemplateID));
                    var data = JsonConvert.DeserializeObject(json) as JObject;
                    var serieCtrl = new METemplateChartSeriesController();

                    foreach (var chartConfig in renderChart)
                    {
                        var listChartSeries = serieCtrl.GetListBusinessObjects<METemplateChartSeriesInfo>(serieCtrl.GetAllDataByForeignColumn("FK_METemplateChartID", chartConfig.METemplateChartID));
                        var gui = new guiChartConfig(data, chartConfig, listChartSeries, AppMemCache.GetTemplateParamsDictPath(document.FK_METemplateID), true)
                        {
                            AllowSaveConfig = false,
                            Module = this
                        };
                        if (gui.ShowDialog() == DialogResult.OK)
                        {
                            var field = _emrDocumentHelper.GetFirstFieldByPath(chartConfig.METemplateChartContainerParam, document.MEEmrDocumentGuid);
                            var position = this._richEditCtrl.Document.CaretPosition;
                            if (field != null)
                            {
                                doc.Replace(field.ResultRange, EmrParam.BeginTag + EmrParam.EndTag);
                                position = doc.CreatePosition(field.ResultRange.Start.ToInt() + 1);
                            }
                            this._emrDocumentHelper.InsertImageToDocument(gui.ChartImage, position, 0, 0, false, false);
                        }
                    }
                }
                catch (Exception)
                {

                    throw;
                }
                finally
                {
                    BOSProgressBar.Close();
                }
            }
        }
        internal bool IsForceRevokePermission()
        {
            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            // to benh an moi tao co the id = 0
            if (document.MEEmrDocumentID == 0) return false;
            var docInDb = _emrDocumentCtrl.GetObjectByID(document.MEEmrDocumentID) as MEEmrDocumentsInfo;
            if (docInDb == null) return false;
            if (document.FK_EditingUserID != docInDb.FK_EditingUserID || docInDb.MEEmrDocumentHoldMachineMac != _macAddress)
            {
                var emp = _employeeCtrl.GetObjectByID(docInDb.FK_EditingUserID) as HREmployeesInfo;
                var msg = "Có thể Admin đã lấy lại quyền vì bạn giữ tờ bệnh án quá lâu.";
                if (emp != null)
                {
                    msg = $"Ai đó đã dùng tài khoản của bạn trên máy khác. {emp.HREmployeeName} ({docInDb.MEEmrDocumentHoldMachineIp}:{docInDb.MEEmrDocumentHoldMachineMac})";
                }
                else if (docInDb.FK_EditingUserID == 0)
                {
                    if (document.MEEmrDocumentGroup == docInDb.MEEmrDocumentGroup && document.MEEmrDocumentOrder == docInDb.MEEmrDocumentOrder)
                    {
                        var serverHash = _hashProvider.ComputeHash(DownloadFtpFileTemp(document));
                        string currentFile = string.Format(@"{0}\Emr\{1}\{2}.{3}", _documentPath, document.FK_MEEmrID, document.MEEmrDocumentFile, document.MEEmrDocumentFileExt);
                        var currentHash = _hashProvider.ComputeHash(currentFile);
                        if (currentHash == serverHash)
                        {
                            return false;
                        }
                    }
                    var revoker = _geObjHistoryCtrl.GetLatestHistoryByObjectNameAndObjectId(TableName.MEEmrDocumentsTableName, document.MEEmrDocumentID, cstObjectHistoryActionRevokeEditPer);
                    if (revoker != null)
                    {
                        emp = _employeeCtrl.GetEmployeeByUserID(revoker.ADUserID) as HREmployeesInfo;
                        if (emp != null)
                            msg = $"{emp.HREmployeeName} đã thực hiện lấy lại quyền lúc: {revoker.GEObjectHistoryDate.ToString("dd/MM/yyyy HH:mm:ss")} vì bạn giữ tờ bệnh án quá lâu.";
                    }
                }

                MessageBox.Show("Không thể lưu tờ bệnh án. Bạn đã mất quyền sửa trên tờ bệnh án này." +
                    "\n" + msg +
                    "\nNếu muốn tiếp tục, bạn vui lòng đăng nhập lại.",
                    "Mất quyền sửa tờ bệnh án", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return true;
            }
            return false;
        }
        internal bool SaveFileDocument()
        {
            //moi thao tac luu bang code cung se ko dc thuc hien
            //_overrideEditPermission hieu luc cao hon cac quyen con lai
            if (this._richEditCtrl.ReadOnly || (_overrideEditPermission == false)) return false;

            if (IsViolateSignRule())
            {
                MessageBox.Show("Bạn đã sửa nội dung đã được ký của người khác, " +
                    "bao gồm việc thay đổi trật tự nội dung. Mọi thay đổi sẽ không thể lưu.",
                    "Vi phạm quy định ký", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }

            if (!IsShareEdit())
            {
                MessageBox.Show("Bạn được chia sẻ CHỈ ĐỌC bệnh án.",
                    "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }

            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;

            BeforeSaveFileDocument();

            BOSProgressBar.Start("Đang lưu và upload tập tin");

            var doc = this._richEditCtrl.Document;
            doc.EndUpdate();
            string fileName = string.Format(@"{0}\Emr\{1}\{2}.docx", _documentPath, document.FK_MEEmrID, document.MEEmrDocumentFile);
            _richEditCtrl.SaveDocument(fileName, DocumentFormat.OpenXml);

            var currentView = _richEditCtrl.ActiveView as PageBasedRichEditView;
            if (currentView != null && currentView.PageCount > _entity.METemplate.METemplateMaximumPage && _entity.METemplate.METemplateMaximumPage > 0)
            {
                _msgMaximumPage = ($"Tờ bệnh án vượt số trang quy định ({_entity.METemplate.METemplateMaximumPage} trang).");
            }
            return true;

        }

        // Revert 26/10/2020 3:35:26 PM
        internal bool EncryptFileDocument()
        {
            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            string fileName = string.Format(@"{0}\Emr\{1}\{2}.docx", _documentPath, document.FK_MEEmrID, document.MEEmrDocumentFile);
            string backupName = $"{document.MEEmrDocumentFile}.{BOSApp.CurrentUsersInfo.ADUserID}.{document.AAUpdatedDate.ToString("ddMMyyyyHHmmss")}.encrypt.bk.docx";
            string backupPath = string.Format(@"{0}\Emr\{1}\{2}", _documentPath, document.FK_MEEmrID, backupName);
            try
            {
                if (_checkSystem)
                {
                    BOSProgressBar.SetText($"Dự phòng {fileName} sang {backupPath}");
                }
                File.Copy(fileName, backupPath);
                PrintMgsLog("ENCRYPT-FILE", fileName);
                if (_checkSystem)
                {
                    BOSProgressBar.SetText($"Xử lý mở mã hóa {fileName}");
                }
                using (OfficeOpenXmlCrypto.OfficeCryptoStream stream = OfficeOpenXmlCrypto.OfficeCryptoStream.Open(fileName, this._emrDocumentHelper.ShareEmrPassword))
                {
                    stream.Password = this._emrDocumentHelper.ShareEmrPassword;
                    stream.Save();
                }
                PrintMgsLog("VERIFY-ENCRYPT-FILE", fileName);
                OfficeOpenXmlCrypto.OfficeCryptoStream verifyStream;
                //5.71% CPU TODO optimization
                var verify = OfficeOpenXmlCrypto.OfficeCryptoStream.TryOpen(fileName, this._emrDocumentHelper.ShareEmrPassword, out verifyStream);
                if (verifyStream != null) verifyStream.Close();
                if (!verify)
                {
                    if (_checkSystem)
                    {
                        BOSProgressBar.SetText($"Thất bại mở mã hóa {fileName}. Tiến hành thử lại 01 lần nữa!");
                    }
                    //try again 1 time
                    File.Delete(fileName);
                    using (OfficeOpenXmlCrypto.OfficeCryptoStream stream = OfficeOpenXmlCrypto.OfficeCryptoStream.Open(backupPath, this._emrDocumentHelper.ShareEmrPassword))
                    {
                        stream.Password = this._emrDocumentHelper.ShareEmrPassword;
                        stream.SaveAs(fileName);
                    }
                    verify = OfficeOpenXmlCrypto.OfficeCryptoStream.TryOpen(fileName, this._emrDocumentHelper.ShareEmrPassword, out verifyStream);
                    if (verifyStream != null) verifyStream.Close();

                    Trace.TraceError("OfficeOpenXmlCrypto ERROR + TRY AGAIN: {0}:{1}:{2}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), BOSApp.CurrentUser, fileName);
                    Trace.Flush();
                }

                if (!verify)
                {
                    if (_checkSystem)
                    {
                        BOSProgressBar.SetText($"Không thể mở mã hóa {fileName}.");
                    }
                    _ftpFileMng.UploadFile($"/Emr/{document.FK_MEEmrID}/", backupName, backupPath);
                    throw new Exception("Không thể mã hóa tờ bệnh án");
                }
                else
                {
                    if (_checkSystem)
                    {
                        BOSProgressBar.SetText($"Xóa dự phòng {backupPath}");
                    }
                    //delete backup file
                    File.Delete(backupPath);
                }
                if (_checkSystem)
                {
                    BOSProgressBar.SetText($"Đã mở mã hóa {fileName}");
                }
                PrintMgsLog("KET-THUC-VERIFY-ENCRYPT-FILE", fileName);
                return true;
            }
            catch (OfficeOpenXmlCrypto.InvalidPasswordException ex)
            {
                var msg = "OfficeOpenXmlCrypto.InvalidPasswordException " + ex.ToString();
                MessageBox.Show("Không mã hóa được tờ bệnh án. Vui lòng liên hệ Admin. Tờ bệnh án (chưa mã hóa) được sao lưu ở: " + backupPath,
                    "Không mã hóa được tờ bệnh án", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _richEditCtrl.Modified = true;
                Trace.TraceError("OfficeOpenXmlCrypto ERROR: {0}:{1}:{2}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), BOSApp.CurrentUser, msg);
                Trace.Flush();
                return false;
            }
            catch (Exception ex)
            {
                var msg = ex.ToString();
                MessageBox.Show("Có lỗi xảy ra, vui lòng liên hệ Admin. Tờ bệnh án (chưa mã hóa) sao lưu ở: " + backupPath + "\n" + ex.ToString(),
                    "Không mã hóa được tờ bệnh án", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _richEditCtrl.Modified = true;
                Trace.TraceError("OfficeOpenXmlCrypto ERROR: {0}:{1}:{2}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), BOSApp.CurrentUser, msg);
                Trace.Flush();
                return false;
            }
        }
        /// <summary>
        /// Save xuong db
        /// </summary>
        internal bool SaveEmrDocumentInfoAndUploadFile()
        {
            if (_checkSystem)
            {
                var watchSaveDoc = Stopwatch.StartNew();
                var result = SaveDocumentAndUpload();
                watchSaveDoc.Stop();
                var elapsedMs = watchSaveDoc.ElapsedMilliseconds / 1000.0; // seconds
                _sysHelper.LogTxt("information", $"{elapsedMs} giây SaveDocumentAndUpload - total.");
                return result;
            }
            else
            {
                return SaveDocumentAndUpload();
            }
        }

        private bool SaveDocumentAndUpload()
        {
            var item = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            var docx = "." + EmrDocumentFileExtention.docx.ToString();
            var dir = Path.Combine(_documentPath, "Emr", item.FK_MEEmrID.ToString());
            string fileName = Path.Combine(dir, item.MEEmrDocumentFile + docx);
            item.MEEmrDocumentHash = _md5Hasher.ComputeHash(fileName);
            // SQL , Mongo
            item = SaveEmrDocumentInfo(item, true);
            // Mix
            RunActionsOnSaveDocument(item);
            TDTUpdateEmrDocumentValidatesWhenSave(item);
            // End Mix
            // FTP
            var uploaded = UploadDocumentFile(item);
            // Local
            var name = item.MEEmrDocumentFile;
            Task.Run(() =>
            {
                var di = new DirectoryInfo(dir);
                foreach (var file in di.GetFiles($"{name}_MD5*{docx}", SearchOption.TopDirectoryOnly))
                {
                    file.Delete();
                }
                File.Copy(fileName, Path.Combine(dir, $"{name}_MD5{item.MEEmrDocumentHash}{docx}"));
            });
            if (!uploaded) return false;
            _msgNotification.Text += string.IsNullOrEmpty(_msgMaximumPage) ? "" : $" {_msgMaximumPage}";
            return true;
        }

        // Fine todo apply EmrModule.
        internal bool SaveEmrDocumentInfoAndUploadFileSafe()
        {
            var configSaveDocument = BOSApp.GetSystemConfigValue(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.EMR_DOCUMENT_SAVE_SAFE)?.ToUpper() == "TRUE";
            if (!configSaveDocument)
            {
                if (!EncryptFileDocument()) return false;
                return SaveEmrDocumentInfoAndUploadFile();
            }
            // 1. [Encrypt]
            // 2. [Upload] file ten temp // [Check Encrypt] => Rename
            // 3. [Check Hash]
            // 2 vs 3 ok => SaveEmrDocumentInfo().
            var msg = string.Empty;
            var docx = "." + EmrDocumentFileExtention.docx.ToString();
            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            string serverPath = $"/Emr/{document.FK_MEEmrID}/";
            var dir = Path.Combine(_documentPath, "Emr", document.FK_MEEmrID.ToString());
            var fileName = $"{document.MEEmrDocumentFile}.docx";
            string fileFullPath = Path.Combine(dir, fileName);
            string backupName = $"{document.MEEmrDocumentFile}.{BOSApp.CurrentUsersInfo.ADUserID}.{DateTime.Now.ToString("ddMMyyyyHHmmss")}.bk{docx}";
            string backupFullPath = Path.Combine(dir, backupName);
            // store file encrypted backup use upload (failed) & VerifyOfficeCryto.
            string encryptedBackupName = $"{document.MEEmrDocumentFile}.{BOSApp.CurrentUsersInfo.ADUserID}.{DateTime.Now.ToString("ddMMyyyyHHmmss")}.encrypted.bk{docx}";
            string encryptedBackupFullPath = Path.Combine(dir, encryptedBackupName);
            try
            {
                // 1. [LOCAL] Copy file sang 1 file khác cùng thư mục : tên file: [ FileName.{ADUserID}.{ddMMyyyyHHmmss}.bk.docx ];
                File.Copy(fileFullPath, backupFullPath);
                PrintMgsLog("ENCRYPT-FILE", fileFullPath);
                // 2. [LOCAL] Cập nhật nội dung vô file.
                using (OfficeOpenXmlCrypto.OfficeCryptoStream stream = OfficeOpenXmlCrypto.OfficeCryptoStream.Open(fileFullPath, this._emrDocumentHelper.ShareEmrPassword))
                {
                    stream.Password = this._emrDocumentHelper.ShareEmrPassword;
                    stream.Save();
                }

                // 3. [LOCAL] Copy file => [ FileName.{ADUserID}.{ddMMyyyyHHmmss}.encrypted.bk{docx} ];
                // File.Copy faster stream.SaveAs
                File.Copy(fileFullPath, encryptedBackupFullPath);

                // 4. [DB] HASH MD5 file + save nội dung vô MEEmrDocuments (MEEmrDocumentHash)
                document.MEEmrDocumentHash = _md5Hasher.ComputeHash(fileFullPath);
                document = SaveEmrDocumentInfo(document, true);

                RunActionsOnSaveDocument(document);

                TDTUpdateEmrDocumentValidatesWhenSave(document);

                // 5. [FTP] [PROCESS 1] Upload file [ FileName.{ADUserID}.{ddMMyyyyHHmmss}.encrypted.bk{docx} ] lên ftp với filename: [ FileName_yyyyMMddHHmmssfff.docx] 
                //  - KHÔNG ĐỤNG TỚI FILE ĐANG CÓ (FILENAME.docx)
                // Use file encrypted backup check. Avoid another process
                var fileUpload = UploadDocumentFileSafe(serverPath, fileName, fileFullPath, document.FK_MEEmrID.ToString(), encryptedBackupFullPath);
                #region UploadFileXMLToServer
                //if (!true)
                //{
                //    if (!true)
                //    {
                //        var templateParams = _templateParamCtrl.GetAllTemplateParamObjectByTemplateID(document.FK_METemplateID);
                //        var paramCtrl = new MEParamsController();
                //        var listMEParamInfo = new List<MEParamsInfo>();
                //        foreach (var templateParam in templateParams)
                //        {
                //            var param = paramCtrl.GetObjectByID(templateParam.FK_MEParamID) as MEParamsInfo;
                //            listMEParamInfo.Add(param);
                //        }
                //        var xmlParamList = new Dictionary<string, string>();
                //        MEEmrsInfo emr = _entity.MainObject as MEEmrsInfo;
                //        var paramList = GetParams(emr, document, listMEParamInfo);
                //        string localPath = string.Format(@"{0}\Emr\{1}\{2}.{3}", _documentPath, document.FK_MEEmrID, document.MEEmrDocumentFile, "xml");

                //        foreach (var xItem in paramList)
                //        {
                //            try
                //            {
                //                var xKey = xItem.Key; // MEParamNo
                //                object xItemValue = xItem.Value;
                //                string json = JsonConvert.SerializeObject(xItemValue, Newtonsoft.Json.Formatting.None);
                //                string temp = json.Replace("[", "").Replace("]", "").Replace("\"", "");
                //                if (xItemValue != null)
                //                {
                //                    xmlParamList.Add(xKey, temp);
                //                }
                //            }
                //            catch (Exception ex)
                //            {
                //                xmlParamList.Add(xItem.Key, string.Empty);
                //            }
                //        }
                //        bool isGeneralXML = GeneralXML(xmlParamList, listMEParamInfo, document.MEEmrDocumentNo, $"{document.MEEmrDocumentFile}.xml", localPath, serverPath);
                //    }
                //    else
                //    {
                //        MEEmrsInfo emr = _entity.MainObject as MEEmrsInfo;
                //        var ob = GetJsonMongo(emr, document);
                //        string localPath = string.Format(@"{0}\Emr\{1}\{2}.{3}", _documentPath, document.FK_MEEmrID, document.MEEmrDocumentFile, "xml");
                //        bool isConvert = ConvertJsonToXML(ob.ToString(), document.MEEmrDocumentNo, $"{document.MEEmrDocumentFile}.xml", localPath, serverPath);
                //    }
                //}
                #endregion
                // 6. [LOCAL] Xóa file [ {name}_MD5*{docx} ] và tạo file MD5 mới 
                var name = document.MEEmrDocumentFile;
                Task.Run(() =>
                {
                    var di = new DirectoryInfo(dir);
                    foreach (var file in di.GetFiles($"{name}_MD5*{docx}", SearchOption.TopDirectoryOnly))
                    {
                        file.Delete();
                    }
                    File.Copy(fileFullPath, Path.Combine(dir, $"{name}_MD5{document.MEEmrDocumentHash}{docx}"));
                });

                // 7. Kiểm tra PROCESS 1. Nếu ok trả về filename [ FileName_yyyyMMddHHmmssfff.docx ] thì tiếp bước 9. # return false.
                if (string.IsNullOrEmpty(fileUpload)) return false;

                // 8. [LOCAL] [PROCESS 2A] KIỂM TRA FILE MÃ HÓA. NẾU FILE MÃ HÓA OK THÌ XÓA File backup ở 1. [ FileName.{ADUserID}.{ddMMyyyyHHmmss}.bk.docx ]
                // 9. [FTP] [PROCEES 2B] Kiểm tra md5 file trên ftp ở bước 5 với mã md5 truyền lên. 
                if (_sysnSave)
                {
                    var verifyCrytoTask = Task.Factory.StartNew<bool>(() =>
                    {
                        return VerifyOfficeCryto(encryptedBackupFullPath, false, backupFullPath, serverPath);
                    });
                    var verifyUploadFileTask = Task.Factory.StartNew<bool>(() =>
                    {
                        return VerifyUploadFile(fileFullPath, fileUpload, document.FK_MEEmrID.ToString(), document.MEEmrDocumentHash);
                    });
                    Task.WaitAll(verifyCrytoTask, verifyUploadFileTask);

                    // PROCESS 2A và 2B thành công
                    if (verifyCrytoTask.Result && verifyUploadFileTask.Result)
                    {
                        File.Delete(encryptedBackupFullPath);
                        _ftpFileMng.Rename(serverPath, fileUpload, fileName);
                        _msgNotification.Text += string.IsNullOrEmpty(_msgMaximumPage) ? "" : $" {_msgMaximumPage}";
                        return true;
                    }
                }
                else
                {
                    var verifyCrytoResult = VerifyOfficeCryto(encryptedBackupFullPath, false, backupFullPath, serverPath);
                    var verifyUploadFileResult = VerifyUploadFile(fileFullPath, fileUpload, document.FK_MEEmrID.ToString(), document.MEEmrDocumentHash);
                    if (verifyCrytoResult && verifyUploadFileResult)
                    {
                        File.Delete(encryptedBackupFullPath);
                        _ftpFileMng.Rename(serverPath, fileUpload, fileName);
                        _msgNotification.Text += string.IsNullOrEmpty(_msgMaximumPage) ? "" : $" {_msgMaximumPage}";
                        return true;
                    }
                }

                return false;
            }
            catch (OfficeOpenXmlCrypto.InvalidPasswordException ex)
            {
                msg = "OfficeOpenXmlCrypto.InvalidPasswordException " + ex.ToString();
                MessageBox.Show("Không mã hóa được tờ bệnh án. Vui lòng liên hệ Admin. Tờ bệnh án (chưa mã hóa) được sao lưu ở: " + backupFullPath,
                    "Không mã hóa được tờ bệnh án", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                msg = ex.ToString();
                MessageBox.Show("Vui lòng liên hệ Admin. Chi tiết \n" + msg,
                    "Lỗi tờ bệnh án", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            _richEditCtrl.Modified = true;
            Trace.TraceError("ERROR: {0}:{1}:{2}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), BOSApp.CurrentUser, msg);
            Trace.Flush();
            return false;
        }
        // Revert 26/10/2020 3:35:26 PM
        internal bool UploadDocumentFile(MEEmrDocumentsInfo item)
        {
            //var item = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            string fileName = string.Format(@"{0}\Emr\{1}\{2}.docx", _documentPath, item.FK_MEEmrID, item.MEEmrDocumentFile);
            try
            {
                if (_checkSystem)
                {
                    BOSProgressBar.SetText($"Bắt đầu upload tờ {fileName} lên máy chủ FTP...");
                    _sysHelper.LogTxt("information", $"Bắt đầu upload tờ {fileName} lên máy chủ FTP...");
                    var watchFTP = Stopwatch.StartNew();
                    PrintMgsLog("BAT-DAU-UPLOAD-FILE", fileName);
                    _ftpFileMng.UploadFile($"/Emr/{item.FK_MEEmrID}/", item.MEEmrDocumentFile + ".docx", fileName);
                    PrintMgsLog("KET-THUC-UPLOAD-FILE", fileName);
                    watchFTP.Stop();
                    var elapsedFTP = watchFTP.ElapsedMilliseconds / 1000.0; // seconds
                    _sysHelper.LogTxt("information", $"{elapsedFTP} giây. Hoàn tất upload tờ {fileName} lên máy chủ FTP.");
                    BOSProgressBar.SetText($"Hoàn tất upload tờ {fileName} lên máy chủ FTP.");
                }
                else
                {
                    PrintMgsLog("BAT-DAU-UPLOAD-FILE", fileName);
                    _ftpFileMng.UploadFile($"/Emr/{item.FK_MEEmrID}/", item.MEEmrDocumentFile + ".docx", fileName);
                    PrintMgsLog("KET-THUC-UPLOAD-FILE", fileName);
                }
            }
            catch (Exception ex)
            {
                string backupName = string.Format(@"{0}\Emr\{1}\{2}.{3}.{4}.upload.bk.docx", _documentPath, item.FK_MEEmrID, item.MEEmrDocumentFile,
                    BOSApp.CurrentUsersInfo.ADUserID, item.AAUpdatedDate.ToString("ddMMyyyyHHmmss"));
                File.Copy(fileName, backupName, true);
                Trace.TraceError("FtpFileMng.UploadFile ERROR: {0}:{1}:{2}:{3}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), BOSApp.CurrentUser, ex, fileName);
                Trace.TraceError("FtpFileMng.UploadFile ERROR BACKUP AT: {0}:{1}:{2}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), BOSApp.CurrentUser, backupName);
                Trace.Flush();
                if (MessageBox.Show(
                     "LỖI: " + ex.Message +
                     "\n\nCÓ LỖI KHI TẢI LÊN TỜ BỆNH ÁN, MỌI THAY ĐỔI ĐÃ KHÔNG LƯU ĐƯỢC." +
                     "\nNếu vẫn có lỗi vui lòng liên hệ Admin. " +
                     "\nTờ bệnh án (đã mã hóa) được sao lưu ở: " + backupName +
                     "\n\n XIN THỬ LƯU LẠI VÀI LẦN NỮA. " +
                     "\n RETRY để thử lại. CANCEL để hủy.",
                     "[NGHIÊM TRỌNG] KHÔNG THỂ TẢI LÊN TỜ BỆNH ÁN", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error) == DialogResult.Retry)
                {
                    if (item.FK_EditingUserID > 0 && item.FK_EditingUserID != BOSApp.CurrentEmployeesInfo.HREmployeeID)
                    {
                        ShowDocumentEditingUser(item);
                        return false;
                    }
                    // Check edited
                    var docInDb = _emrDocumentCtrl.GetObjectByID(item.MEEmrDocumentID) as MEEmrDocumentsInfo;
                    if (docInDb.FK_EditingUserID != item.FK_EditingUserID || docInDb.MEEmrDocumentHoldMachineMac != _macAddress)
                    {
                        MessageBox.Show("Không thể lưu tờ bệnh án. Bạn đã mất quyền sửa trên tờ bệnh án này." +
                        "\nNếu muốn tiếp tục, bạn vui lòng mở lại tờ này lần nữa, hoặc thao tác trên tờ bệnh án khác và quay lại sau.",
                        "Mất quyền sửa tờ bệnh án", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return false;
                    }
                    return UploadDocumentFile(item);
                }
                else
                {
                    _richEditCtrl.Modified = true;
                    return false;
                }
            }
            return true;
        }

        internal string UploadDocumentFileSafe(string serverPath, string fileName, string localFullPath, string emrId, string backup)
        {
            var result = string.Empty;
            try
            {
                PrintMgsLog("BAT-DAU-UPLOAD-FILE", localFullPath);
                result = _ftpFileMng.UploadFileSafe(serverPath, fileName, localFullPath);
                PrintMgsLog("KET-THUC-UPLOAD-FILE", localFullPath);
            }
            catch (Exception ex)
            {
                Trace.TraceError("FtpFileMng.UploadFile ERROR: {0}:{1}:{2}:{3}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), BOSApp.CurrentUser, ex, localFullPath);
                Trace.TraceError("FtpFileMng.UploadFile ERROR BACKUP AT: {0}:{1}:{2}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), BOSApp.CurrentUser, backup);
                Trace.Flush();
                if (MessageBox.Show(
                     "LỖI: " + ex.Message +
                     "\n\nCÓ LỖI KHI TẢI LÊN TỜ BỆNH ÁN, MỌI THAY ĐỔI ĐÃ KHÔNG LƯU ĐƯỢC." +
                     "\nNếu vẫn có lỗi vui lòng liên hệ Admin. " +
                     "\nTờ bệnh án (đã mã hóa) được sao lưu ở: " + backup +
                     "\n\n XIN THỬ LƯU LẠI VÀI LẦN NỮA. " +
                     "\n RETRY để thử lại. CANCEL để hủy.",
                     "[NGHIÊM TRỌNG] KHÔNG THỂ TẢI LÊN TỜ BỆNH ÁN", MessageBoxButtons.RetryCancel, MessageBoxIcon.Error) == DialogResult.Retry)
                {
                    return UploadDocumentFileSafe(serverPath, fileName, localFullPath, emrId, backup);
                }
                else
                {
                    _richEditCtrl.Modified = true;
                    return result;
                }
            }
            return result;
        }
        internal MEEmrDocumentsInfo ExtractDataFromContentToDocumentInfo(MEEmrDocumentsInfo item, JToken data)
        {
            if (data == null) return item;
            var list = AppMemCache.GetTemplateParams(item.FK_METemplateID).Where(o => !string.IsNullOrEmpty(o.METemplateParamUpdateTo))
                .GroupBy(o => o.METemplateParamUpdateTo,
                (k, g) => new { Group = k, List = g.ToList() });

            var defaultVal = new MEEmrDocumentsInfo();
            foreach (var cfg in list)
            {
                if (!string.IsNullOrEmpty(cfg.Group))
                {
                    if (cfg.List.Count == 0) continue;
                    var property = item.GetType().GetProperty(cfg.List.First().METemplateParamUpdateTo);
                    List<List<object>> valueArr = new List<List<object>>();
                    foreach (var p in cfg.List)
                    {
                        List<object> values = new List<object>();
                        var path = "$." + p.METemplateParamPath;
                        var param = AppMemCache.GetParamFromDictKeyID(p.FK_MEParamID);
                        var tokens = data.SelectTokens(path).ToList();

                        var formatType = param.MEParamFormatType;
                        var formatStr = param.MEParamFormatString;
                        if (!string.IsNullOrEmpty(p.MEParamFormatString))
                            formatStr = p.MEParamFormatString;
                        if (!string.IsNullOrEmpty(p.MEParamFormatType))
                            formatType = p.MEParamFormatType;
                        if (tokens.Count() == 1)
                        {
                            var token = tokens.First();
                            if (token is JValue)
                            {
                                var v = (tokens.First() as JValue)?.Value;
                                if (v != null)
                                {
                                    values = ExtractDataFromContentAddToValues(property, values, formatType, formatStr, v);
                                }
                            }
                            else if (token is JArray)
                            {
                                foreach (JValue v in (token as JArray))
                                    if (v != null)
                                    {
                                        values = ExtractDataFromContentAddToValues(property, values, formatType, formatStr, v);
                                    }
                            }
                        }
                        else if (tokens.Count() > 1)
                        {
                            foreach (JValue v in tokens)
                                if (v != null)
                                    values.Add(this._emrDocumentHelper.GetStringFromDataValue(formatType, formatStr, v.Value, string.Empty));
                        }
                        values = values.Where(v => !string.IsNullOrEmpty(v.ToString())).Distinct().ToList();
                        valueArr.Add(values);
                    }

                    if (valueArr.Count > 0)
                    {
                        object value = null;
                        if (valueArr.Count == 1 && valueArr[0].Count == 1)
                        {
                            value = valueArr[0][0];
                        }
                        else
                        {
                            value = string.Empty;
                            for (int i = 0; i < valueArr[0].Count; i++)
                            {
                                for (int j = 0; j < valueArr.Count; j++)
                                {
                                    if (valueArr[j].Count > i)
                                        value += valueArr[j][i].ToString() + ", ";
                                }
                                value = (value as string).TrimEnd(' ', ',');
                                value += EmrConsts.DATA_FROM_CONTENT_SEPARATOR;
                            }
                        }
                        if (typeof(String) == value.GetType())
                            value = value.ToString().TrimEnd(' ', ';', ',');
                        try
                        {
                            if (property.PropertyType == typeof(String))
                            {
                                OnEmrDocumentPropertyChange(item, property.Name, property?.GetValue(item), value.ToString());
                                property?.SetValue(item, value.ToString(), null);
                            }
                            else
                            {
                                //TODO OnEmrDocumentPropertyChange
                                var type = Convert.GetTypeCode(property.GetValue(item));
                                property?.SetValue(item, Convert.ChangeType(value, type), null);
                            }
                        }
                        catch (Exception ex)
                        {
                            try { property?.SetValue(item, property.GetValue(defaultVal), null); } catch { /*just for safe, do nothing*/ }
                            PrintMgsLog("LOI-GAN-GIA-TRI-VAO-TO-BENH-AN", ex.ToString());
                        }
                    }
                }
            }
            return item;
        }

        private List<object> ExtractDataFromContentAddToValues(System.Reflection.PropertyInfo property, List<object> values, string formatType, string formatStr, object v)
        {
            var typeV = v.GetType();
            if (property.PropertyType == typeV)
                values.Add(v);
            else if (property.PropertyType == typeof(int) && typeV == typeof(double))
                values.Add(v);
            else if (property.PropertyType == typeof(int) && typeV == typeof(float))
                values.Add(v);
            else
                values.Add(this._emrDocumentHelper.GetStringFromDataValue(formatType, formatStr, v, string.Empty));

            return values;
        }

        private void OnEmrDocumentPropertyChange(MEEmrDocumentsInfo document, string propertyName, object oldVal, string newVal)
        {
            if (oldVal?.ToString() == newVal) return;
            switch (propertyName)
            {
                case "MEEmrDocumentGroup":
                    document.MEEmrDocumentSubOrder = GetDocumentSubOrder(document.FK_MEEmrID, newVal);
                    break;
                default:
                    break;
            }
        }

        private MEEmrsInfo UpdateEmrInfoByDocumentContent(MEEmrsInfo emr, JToken data)
        {
            emr.AllowPropertyChangedEvent = false;
            try
            {
                var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
                var newEmr = ExtractDataFromContentToEmrInfo(emr, data, document);
                //TH khong co thay doi gi thi ko can luu
                if (newEmr != null)
                {
                    emr = newEmr;
                    emr.AAUpdatedUser = BOSApp.CurrentUser;
                    _emrCtrl.UpdateObject(emr);
                    HistoryEmr(emr, cstObjectHistoryActionChange, $"Thay đổi thông tin bệnh án - từ tờ bệnh án [{document.MEEmrDocumentNo}].");
                    RefreshCurrentEmrSearchResultsControl(emr);
                }
            }
            catch (Exception ex)
            {
                _sysHelper.LogTxt("error", ex.ToString());
                throw;
            }
            finally
            {
                emr.AllowPropertyChangedEvent = true;
            }
            return emr;
        }
        internal MEEmrsInfo ExtractDataFromContentToEmrInfo(MEEmrsInfo emr, JToken data, MEEmrDocumentsInfo document)
        {
            if (data == null) return null;
            var list = AppMemCache.GetTemplateParams(document.FK_METemplateID).Where(o => !string.IsNullOrEmpty(o.METemplateParamUpdToEmr))
                .GroupBy(o => o.METemplateParamUpdToEmr,
                (k, g) => new { Group = k, List = g.ToList() });
            if (list.Count() == 0) return null;
            var defaultVal = new MEEmrsInfo();
            foreach (var cfg in list)
            {
                if (!string.IsNullOrEmpty(cfg.Group))
                {
                    if (cfg.List.Count == 0) continue;
                    var property = emr.GetType().GetProperty(cfg.List.First().METemplateParamUpdToEmr);
                    List<List<object>> valueArr = new List<List<object>>();
                    foreach (var p in cfg.List)
                    {
                        List<object> values = new List<object>();
                        var path = "$." + p.METemplateParamPath;
                        var param = AppMemCache.GetParamFromDictKeyID(p.FK_MEParamID);
                        var tokens = data.SelectTokens(path).ToList();

                        var formatType = param.MEParamFormatType;
                        var formatStr = param.MEParamFormatString;
                        if (!string.IsNullOrEmpty(p.MEParamFormatString))
                            formatStr = p.MEParamFormatString;
                        if (!string.IsNullOrEmpty(p.MEParamFormatType))
                            formatType = p.MEParamFormatType;

                        if (tokens.Count() == 1)
                        {
                            var token = tokens.First();
                            if (token is JValue)
                            {
                                var v = (tokens.First() as JValue)?.Value;
                                if (v != null)
                                {
                                    values = ExtractDataFromContentAddToValues(property, values, formatType, formatStr, v);
                                }
                            }
                            else if (token is JArray)
                            {
                                foreach (JValue v in (token as JArray))
                                    if (v != null)
                                    {
                                        values = ExtractDataFromContentAddToValues(property, values, formatType, formatStr, v);
                                    }
                            }
                        }
                        else if (tokens.Count() > 1)
                        {
                            foreach (JValue v in tokens)
                                if (v != null)
                                    values.Add(this._emrDocumentHelper.GetStringFromDataValue(formatType, formatStr, v.Value, string.Empty));
                        }
                        values = values.Where(v => !string.IsNullOrEmpty(v.ToString())).Distinct().ToList();
                        valueArr.Add(values);
                    }

                    if (valueArr.Count > 0)
                    {
                        object value = null;
                        if (valueArr.Count == 1 && valueArr[0].Count == 1)
                        {
                            value = valueArr[0][0];
                        }
                        else
                        {
                            value = string.Empty;
                            for (int i = 0; i < valueArr[0].Count; i++)
                            {
                                for (int j = 0; j < valueArr.Count; j++)
                                {
                                    if (valueArr[j].Count > i)
                                        value += valueArr[j][i].ToString() + ", ";
                                }
                                value = (value as string).TrimEnd(' ', ',');
                                value += EmrConsts.DATA_FROM_CONTENT_SEPARATOR;
                            }
                        }
                        if (typeof(String) == value.GetType())
                            value = value.ToString().TrimEnd(' ', ';', ',');
                        try
                        {
                            if (property.PropertyType == typeof(String))
                            {
                                property?.SetValue(emr, value.ToString(), null);
                            }
                            else
                            {
                                var type = Convert.GetTypeCode(property.GetValue(emr));
                                property?.SetValue(emr, Convert.ChangeType(value, type), null);
                            }
                        }
                        catch (Exception ex)
                        {
                            try { property?.SetValue(emr, property.GetValue(defaultVal), null); } catch { /*just for safe, do nothing*/ }
                            PrintMgsLog("LOI-GAN-GIA-TRI-VAO-BENH-AN", ex.ToString());
                        }
                    }
                }
            }
            return emr;
        }
        private MEEmrsInfo CopyPatientInfoToEmr(MEEmrsInfo emr)
        {
            var patient = _entity.ModuleObjects[TableName.MEPatientsTableName] as MEPatientsInfo;
            emr.MEPatientNo = patient.MEPatientNo;
            emr.MEPatientName = patient.MEPatientName;
            emr.MEPatientBirthday = patient.MEPatientBirthday;
            emr.MEPatientBirthYear = patient.MEPatientBirthday.Year;
            emr.FK_HRDepartmentShortID = emr.FK_HRDepartmentID;
            return emr;
        }
        private void RefreshCurrentEmrSearchResultsControl(MEEmrsInfo emr)
        {
            emr = CopyPatientInfoToEmr(emr);
            MEEmrsInfo main = _entity.MainObject as MEEmrsInfo;
            // nếu đang soạn hồ sơ bệnh nhân thì ko cần load lại lưới
            if (main.MEEmrID == emr.MEEmrID)
                InvalidateSearchResultsControl(null, string.Empty);
        }
        protected override void BeforeInvalidateSearchResultsControl()
        {
            var emr = this._entity.MainObject as MEEmrsInfo;
            emr = CopyPatientInfoToEmr(emr);
        }
        internal MEEmrDocumentsInfo SaveEmrDocumentInfo(MEEmrDocumentsInfo document, bool isHistory)
        {
            MEEmrsInfo emr = _entity.MainObject as MEEmrsInfo;
            if (document.FK_MEEmrID != emr.MEEmrID)
            {
                // patient profile
                emr = _entity.ModuleObjects[TableName.MEEmrsTableName] as MEEmrsInfo;
            }
            document.MEEmrDocumentJson = this.ParserDocumentToJson();
            var data = JsonConvert.DeserializeObject(document.MEEmrDocumentJson) as JToken;
            var isUpdateRefNo = true;
            var configRefNo = BOSApp.GetSystemConfigValue(DocumentProcess.GROUP, DocumentProcess.API_UPDATE_REFNO);
            if (!string.IsNullOrEmpty(configRefNo) && configRefNo.ToUpper() == "FALSE")
                isUpdateRefNo = false;

            var tempNoRef = string.Empty;
            if (!isUpdateRefNo)
            {
                var docDb = _emrDocumentCtrl.GetObjectByID(document.MEEmrDocumentID) as MEEmrDocumentsInfo;
                tempNoRef = docDb.MEEmrDocumentRefNo;
            }

            document = ExtractDataFromContentToDocumentInfo(document, data);

            if (!isUpdateRefNo && !string.IsNullOrEmpty(tempNoRef))
            {
                document.MEEmrDocumentRefNo = tempNoRef;
            }

            document = SetCurrentEditingUser(document);
            document.AAUpdatedUser = BOSApp.CurrentUser;

            if (_checkSystem)
            {
                BOSProgressBar.SetText($"Xử lý lưu thông tin lên máy chủ SQL, bệnh án ID {emr.MEEmrID} dữ liệu tờ bệnh án ID {document.MEEmrDocumentID}.");
                var watchSQL = Stopwatch.StartNew();
                _emrDocumentCtrl.UpdateObject(document);
                emr = UpdateEmrInfoByDocumentContent(emr, data);
                watchSQL.Stop();
                var elapsedSQL = watchSQL.ElapsedMilliseconds / 1000.0; // seconds
                _sysHelper.LogTxt("information", $"{elapsedSQL} giây. Hoàn tất lưu thông tin lên máy chủ SQL, bệnh án ID {emr.MEEmrID} dữ liệu tờ bệnh án ID {document.MEEmrDocumentID}.");
                BOSProgressBar.SetText($"Hoàn tất lưu thông tin lên máy chủ SQL, bệnh án ID {emr.MEEmrID} dữ liệu tờ bệnh án ID {document.MEEmrDocumentID}.");
            }
            else
            {
                _emrDocumentCtrl.UpdateObject(document);
                emr = UpdateEmrInfoByDocumentContent(emr, data);
            }

            UpdateMongoDocument(emr, document);

            if (document.FK_MEEmrID == (_entity.MainObject as MEEmrsInfo).MEEmrID)
            {
                ReorderDocumentList(_entity.MEEmrDocumentsList, emr, document);
            }
            else
            {
                ReorderDocumentList(_entity.MEEmrPatientDocumentsList, emr, document);
            }
            return document;
        }

        #region MongoDB
        internal void InsertMongoDocument(MEEmrsInfo emr, MEEmrDocumentsInfo emrDoc)
        {
            var mongoDoc = Mapper.Map<Clas.Model.Mongo.EmrDocument>(emrDoc);
            mongoDoc.MEEmr = Mapper.Map<Clas.Model.Mongo.Emr>(emr);
            if (!string.IsNullOrEmpty(emrDoc.MEEmrDocumentJson) && emrDoc.MEEmrDocumentJson.Length > 4)
            {
                var converter = new Newtonsoft.Json.Converters.ExpandoObjectConverter();
                var obj = JsonConvert.DeserializeObject<System.Dynamic.ExpandoObject>(emrDoc.MEEmrDocumentJson, converter);
                mongoDoc.MEEmrDocumentContent = (obj == null ? null : obj.ToBsonDocument());
            }
            mongoDoc.AAUpdatedDate = mongoDoc.AACreatedDate;
            emrDoc.MEEmrDocumentMongoID = _emrDocumentMng.Insert(mongoDoc, emrDoc.MEEmrDocumentNo);
        }
        internal void UpdateMongoDocument(MEEmrsInfo emr, MEEmrDocumentsInfo emrDoc, List<string> excludedFields = null)
        {
            if (_checkSystem)
            {
                BOSProgressBar.SetText($"Xử lý lưu thông tin lên máy chủ Mongo, bệnh án ID {emr.MEEmrID} dữ liệu tờ bệnh án ID {emrDoc.MEEmrDocumentID}.");
                var watchMG = Stopwatch.StartNew();
                UpdateMongoDocumentEmr(emr, emrDoc, excludedFields);
                watchMG.Stop();
                var elapsedMG = watchMG.ElapsedMilliseconds / 1000.0; // seconds
                _sysHelper.LogTxt("information", $"{elapsedMG} giây. Hoàn tất lưu thông tin lên máy chủ Mongo, bệnh án ID {emr.MEEmrID} dữ liệu tờ bệnh án ID {emrDoc.MEEmrDocumentID}.");
                BOSProgressBar.SetText($"Hoàn tất lưu thông tin lên máy chủ Mongo, bệnh án ID {emr.MEEmrID} dữ liệu tờ bệnh án ID {emrDoc.MEEmrDocumentID}.");
            }
            else
            {
                UpdateMongoDocumentEmr(emr, emrDoc, excludedFields);
            }
        }

        private void UpdateMongoDocumentEmr(MEEmrsInfo emr, MEEmrDocumentsInfo emrDoc, List<string> excludedFields)
        {
            if (string.IsNullOrEmpty(emrDoc.MEEmrDocumentMongoID))
            {
                InsertMongoDocument(emr, emrDoc);
            }
            else
            {
                var mongoDoc = Mapper.Map<Clas.Model.Mongo.EmrDocument>(emrDoc);
                mongoDoc.MEEmr = Mapper.Map<Clas.Model.Mongo.Emr>(emr);
                if (!string.IsNullOrEmpty(emrDoc.MEEmrDocumentJson) && emrDoc.MEEmrDocumentJson.Length > 4)
                {
                    var converter = new Newtonsoft.Json.Converters.ExpandoObjectConverter();
                    var obj = JsonConvert.DeserializeObject<System.Dynamic.ExpandoObject>(emrDoc.MEEmrDocumentJson, converter);
                    mongoDoc.MEEmrDocumentContent = (obj == null ? null : obj.ToBsonDocument());
                }
                mongoDoc.FK_MEEmrID = emrDoc.FK_MEEmrID;
                mongoDoc.AAUpdatedDate = DateTime.Now.ToLocalTime();
                mongoDoc.AAUpdatedUser = BOSApp.CurrentUser;
                _emrDocumentMng.Update(emrDoc.MEEmrDocumentMongoID, mongoDoc, emrDoc.MEEmrDocumentNo, excludedFields);
            }
        }

        internal void InsertMongoLog(Clas.Model.Base.LogMongo logMongo, string collection)
        {
            var now = DateTime.Now;
            logMongo.AAStatus = "Alive";
            logMongo.AACreatedUser = BOSApp.CurrentUser;
            logMongo.AAUpdatedUser = BOSApp.CurrentUser;
            logMongo.AACreatedDate = now;
            logMongo.AAUpdatedDate = now;
            _emrDocumentMng.InsertLog(logMongo, collection);
        }
        internal List<Clas.Model.Base.LogMongo> FindMongoLog(Dictionary<string, object> filters, Dictionary<string, string> fields, string collection)
        {
            return _emrDocumentMng.GetLogList(filters, fields, collection);
        }
        #endregion

        #region Sign
        internal void SignEmrDocument()
        {
            if (this._richEditCtrl.Modified)
            {
                var confirm = MessageBox.Show("Lưu thay đổi trước khi thực hiện ký", "Nội dung tờ bệnh án đã thay đổi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (IsEmrReadOnly()) return;
            if (IsForceRevokePermission()) return;

            var doc = this._richEditCtrl.Document;
            var range = doc.Range;
            var selectedRanges = doc.Selections;
            if (selectedRanges.Count == 1 && selectedRanges[0].Length == 0)
            {
                var confirm = MessageBox.Show("Bạn muốn ký toàn bộ tập tin bệnh án này? Thao tác sẽ không thể hoàn tác.", "XÁC NHẬN KÝ TOÀN BỘ TỜ BỆNH ÁN", MessageBoxButtons.OKCancel);
                if (confirm == DialogResult.Cancel) return;
                selectedRanges.Clear();
                //uthv fix bug lỗi trên tờ Bệnh án Sản khoa LK
                //System.InvalidOperationException: 'The last cell in the selected range continues the vertical merge, which is not allowed in a selection collection.'
                this._richEditCtrl.Document.SelectAll();
                selectedRanges = doc.Selections;
            }
            else
            {
                var ok = true;
                foreach (var selectedRange in selectedRanges)
                {
                    var beginCell = doc.Tables.GetTableCell(selectedRange.Start);
                    var endCell = doc.Tables.GetTableCell(selectedRange.End);
                    if (beginCell != null && endCell != null)
                    {
                        if (beginCell.Table != endCell.Table)
                        {
                            //không duoc phep phap sinh lenh ky cho loai nay
                            ok = false;
                        }
                    }
                    else if (beginCell == null && endCell != null)
                    {
                        //không duoc phep phap sinh lenh ky cho loai nay
                        ok = false;
                    }
                }
                if (!ok)
                {
                    var confirm = MessageBox.Show("Vùng chọn phải trong 1 bảng, bao trọn 1 bảng hay bao trọn nhiều bảng.", "Vùng chọn không hợp lệ", MessageBoxButtons.OK);
                    return;
                }
            }
            try
            {
                PrintMgsLog("KIEM-TRA-VUNG-KY-HOP-LE", string.Empty);

                var rangePermissions = doc.BeginUpdateRangePermissions();
                //chỉ được ký bao trên nội dung người khác đã ký
                foreach (var selectedRange in selectedRanges)
                {
                    foreach (var rangePermis in rangePermissions)
                    {
                        range = rangePermis.Range;
                        if ((selectedRange.Start > range.Start && selectedRange.End < range.End)
                            || (selectedRange.Start > range.Start && selectedRange.Start < range.End)
                            || (selectedRange.End > range.Start && selectedRange.End < range.End))
                        {
                            MessageBox.Show("Chỉ được ký trên nội dung chưa ký, hoặc ký đè toàn bộ một nội dung đã ký", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                }
                if (IsViolateSignRule())
                {
                    MessageBox.Show("Bạn đã sửa nội dung đã được ký của người khác, bao gồm việc thay đổi trật tự nội dung. Mọi thay đổi sẽ không thể lưu.", "Vi phạm quy định ký", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }
                var signFields = GetSignerRoleFields(selectedRanges);
                if (signFields == null) return;

                PrintMgsLog("VUNG-KY-HOP-LE", string.Empty);
                guiSign guiSign = new guiSign("Bạn muốn ký phần này? Thao tác sẽ không thể hoàn tác.", _fingerPrintZK);
                if (guiSign.ShowDialog() != DialogResult.OK) return;

                PrintMgsLog("BAT-DAU-KY", string.Empty);
                if (signFields.Count > 0)
                {
                    BOSProgressBar.Start("Đang chèn tên và chữ ký");
                    PrintMgsLog("CHEN-CHU-KY", string.Empty);
                    this.InsertSignature(signFields);
                    PrintMgsLog("CHEN-CHU-KY-DONE", string.Empty);
                }
                var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
                BOSProgressBar.Start("Đang xử lý dữ liệu");
                var cloneStream = new MemoryStream();
                this._richEditCtrl.SaveDocument(cloneStream, DocumentFormat.OpenXml);
                var ranges = selectedRanges.Select(r => new { r.Start, r.Length }).ToArray();
                var extractTask = Task.Factory.StartNew<string>(() =>
                {
                    try
                    {
                        using (var cloneRich = new RichEditDocumentServer())
                        {
                            cloneStream.Position = 0;
                            cloneRich.LoadDocument(cloneStream, DocumentFormat.OpenXml);
                            using (var toRich = new RichEditDocumentServer())
                            {
                                var cloneDoc = cloneRich.Document;
                                var toDoc = toRich.Document;
                                range = toDoc.Range;
                                toDoc.BeginUpdate();
                                for (int i = 0; i < ranges.Length; i++)
                                {
                                    var fromRange = cloneDoc.CreateRange(ranges[i].Start, ranges[i].Length);
                                    range = toDoc.InsertDocumentContent(range.End, fromRange, InsertOptions.KeepSourceFormatting);
                                }
                                toDoc.EndUpdate();
                                //Lưu tài liệu lên server
                                PrintMgsLog("XU-LY-PHAN-KY-TEN", string.Empty);
                                string toFileName = document.MEEmrDocumentFile + "_sign" + DateTime.Now.ToBinary();
                                string toFileFullname = toFileName + ".docx";
                                string toFilePath = string.Format(@"{0}\Emr\Partials\{1}\{2}", _documentPath, document.FK_MEEmrID, toFileFullname);
                                CreateSignLocalDir(document.FK_MEEmrID);
                                this.RemoveAllTagForPrint(toRich, _entity.METemplate);
                                this._emrDocumentHelper.RemoveAllParamMarkup(toRich.Document);
                                toRich.SaveDocument(toFilePath, DocumentFormat.OpenXml);
                                PrintMgsLog("UPLOAD-PHAN-KY-TEN", string.Empty);
                                _ftpFileMng.UploadFile($"/Emr/Partials/{document.FK_MEEmrID}/", toFileFullname, toFilePath);
                                return toFileName;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        PrintMgsLog("XU-LY-PHAN-KY-TEN", "ERROR: " + ex.ToString());
                        return string.Empty;
                    }
                    finally
                    {
                        cloneStream.Close();
                        cloneStream.Dispose();
                    }
                });

                var hash = string.Empty;
                var blockAddress = string.Empty;
                var start = DateTime.Now;
                for (int i = 0; i < selectedRanges.Count; i++)
                {
                    this._emrDocumentHelper.SignRange(selectedRanges[i], (i == (selectedRanges.Count - 1)));
                }
                PrintMgsLog("SignRange", "Time: " + (DateTime.Now - start).TotalMilliseconds);
                string fileName = string.Format(@"{0}\Emr\{1}\{2}.docx", _documentPath, document.FK_MEEmrID, document.MEEmrDocumentFile);

                BOSProgressBar.SetText("Đang lưu dữ liệu");

                _richEditCtrl.SaveDocument(fileName, DocumentFormat.OpenXml);
                if (!EncryptFileDocument())
                {
                    BOSProgressBar.Close();
                    return;
                }
                if (!SaveEmrDocumentInfoAndUploadFile())
                {
                    BOSProgressBar.Close();
                    return;
                }
                //save sign history
                PrintMgsLog("LUU-LICH-SU-KY", string.Empty);
                Task.WaitAll(extractTask);
                #region Extra function
                var signFile = extractTask.Result;
                //co cau hinh block chain thi moi thuc hien
                if (!string.IsNullOrEmpty(this._blockWriteUrl) && !string.IsNullOrEmpty(signFile))
                {
                    PrintMgsLog("GOI-API-BLOCK", string.Empty);
                    BOSProgressBar.Start("Đang tạo hash và block");
                    hash = _hashProvider.ComputeHash(signFile);
                    var blockParams = new Dictionary<string, object>
                        {
                            { "digitalFingerprint", hash },
                            { "signature", BOSApp.CurrentUser }
                        };
                    var dataBlock = (JObject)_apiBlock.Post(_blockWriteUrl, null, blockParams);
                    if (dataBlock != null && dataBlock.ContainsKey("TransactionHash"))
                    {
                        blockAddress = dataBlock["TransactionHash"].ToString();
                        PrintMgsLog("TAO-BLOCK-THANH_CONG", blockAddress);
                    }
                    else
                    {
                        _msgNotification.Text = "Có lỗi khi gọi api tạo block. Xem thông báo để biết chi tiết";
                    }
                }
                #endregion
                var sign = new MEEmrDocumentSignsInfo()
                {
                    FK_MEEmrDocumentID = document.MEEmrDocumentID,
                    FK_HRDepartmentID = BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID,
                    FK_HREmployeeID = BOSApp.CurrentEmployeesInfo.HREmployeeID,
                    MEEmrDocumentSignTime = DateTime.Now,
                    MEEmrDocumentSignFile = signFile,
                    MEEmrDocumentSignFileExt = EmrDocumentFileExtention.docx.ToString(),
                    MEEmrDocumentSignType = EmrDocumentSignType.Signed.ToString(),
                    MEEmrDocumentSignRemark = string.Empty,
                    MEEmrDocumentSignHash = hash,
                    MEEmrDocumentSignUser = BOSApp.CurrentUser,
                    MEEmrDocumentSignBlockAddr = blockAddress,
                    AACreatedUser = BOSApp.CurrentUser
                };
                this._emrDocumentSignCtrl.CreateObject(sign);

                CloneSignedRangeForTracking();
                //Clear history để không đc undo.
                ((DevExpress.XtraRichEdit.Model.DocumentModel)this._richEditCtrl.Model).History.Clear();
                _richEditCtrl.Modified = false;
                PrintMgsLog("DA-KY", string.Empty);

                HistoryEmr(_entity.MainObject as MEEmrsInfo, "Change", $"Ký tên tờ {document.MEEmrDocumentFile}.{document.MEEmrDocumentFileExt}");
            }
            catch (Exception ex)
            {
                _msgNotification.Text = "Có lỗi khi ký. Xem thông báo để biết chi tiết";
                PrintMgsLog("KY-TEN-LOI: ", ex.ToString());
            }
            finally
            {
                BOSProgressBar.Close();
            }
        }
        private List<Field> GetSignerRoleFields(IEnumerable<DocumentRange> selectedRanges)
        {
            var roles = GetSignerRoles(selectedRanges, true);
            return roles == null ? null : roles.SelectMany(r => r.Value).Select(f => f.Field).ToList();
        }
        private Dictionary<string, List<EmrField>> GetSignerRoles(IEnumerable<DocumentRange> selectedRanges, bool multiSelect)
        {
            List<string> selectedRoles = new List<string>();
            var roles = this._emrDocumentHelper.GetAllSignerRoles(_richEditCtrl.Document, selectedRanges);
            var doc = _richEditCtrl.Document;
            var config = BOSApp.GetSystemConfigValue(DocumentProcess.GROUP, DocumentProcess.AUTO_VALIDATE_SIGN_RANGE);
            if (config?.ToUpper() == "TRUE")
            {
                if (roles.Count == 0)
                {
                    if (MessageBox.Show("Vùng quét chọn không chứa thẻ chữ ký. \n\n[OK] để tiếp tục ký \n[Cancel] để chọn lại vùng ký khác",
                        "[CẢNH BÁO] VÙNG CHỌN KHÔNG CHỨA THẺ CHỮ KÝ", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.Cancel)
                        return null;
                }
                else
                {
                    //verify vùng chọn đúng hay không
                    var msg = _emrDocumentHelper.VerifySignRange(_richEditCtrl.Document, roles);
                    if (!string.IsNullOrEmpty(msg))
                    {
                        MessageBox.Show(msg, "[CẢNH BÁO] VÙNG CHỌN BỊ THIẾU THẺ", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return null;
                    }
                    var requireParam = _emrDocumentHelper.VerifySignAsOneGroup(_richEditCtrl.Document, roles, selectedRanges);
                    if (requireParam != null)
                    {
                        var param = AppMemCache.GetParamFromDictKeyID(requireParam.FK_MEParamID);
                        if (param != null)
                        {
                            MessageBox.Show($"Phải chọn kèm thẻ [{param.MEParamName} - {param.MEParamCaption}]", "[CẢNH BÁO] CHƯA QUÉT CHỌN ĐỦ THẺ", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            return null;
                        }
                    }
                }
            }
            //TH trong vung chon khong co the chu ky thi roles.Count = 0
            if (roles.Keys.Count > 1)
            {
                var guiSelectRole = new guiSignerRolesSelection(roles, multiSelect) { Module = this };
                guiSelectRole.StartPosition = FormStartPosition.CenterParent;
                // huy khong ky nua
                if (guiSelectRole.ShowDialog() != DialogResult.OK) return null;
                selectedRoles = guiSelectRole.SelectedRoles;
            }
            else if (roles.Keys.Count == 1)
            {
                selectedRoles = new List<string>() { roles.Keys.First() };
            }

            return roles.Where(r => selectedRoles.Contains(r.Key)).ToDictionary(t => t.Key, t => t.Value);
        }
        internal void UnsignEmrDocumentStrikeThrough()
        {
            UnsignEmrDocument(true);
        }
        internal void UnsignEmrDocument()
        {
            UnsignEmrDocument(false);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="strikeThrough">true neu huy ky gach ngang/false neu huy ky xoa bo</param>
        internal void UnsignEmrDocument(bool strikeThrough)
        {
            if (this._richEditCtrl.Modified)
            {
                var confirm = MessageBox.Show("Lưu thay đổi trước khi thực hiện hủy ký", "Nội dung tờ bệnh án đã thay đổi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (IsEmrReadOnly()) return;
            var doc = this._richEditCtrl.Document;
            var selectedRanges = doc.Selections;
            var rangePermissions = doc.BeginUpdateRangePermissions();
            var listRange = new List<RangePermission>();
            //khong quet chon
            if (selectedRanges.Count == 1 && selectedRanges[0].Length == 0)
            {
                foreach (var rangePermis in rangePermissions)
                    if (doc.CaretPosition.ToInt() > rangePermis.Range.Start.ToInt() && doc.CaretPosition.ToInt() < rangePermis.Range.End.ToInt())
                        listRange.Add(rangePermis);
            }
            else
            {
                foreach (var selectedRange in selectedRanges)
                    foreach (var rangePermis in rangePermissions)
                    {
                        var r = rangePermis.Range;
                        if ((selectedRange.Start.ToInt() <= r.Start.ToInt() && selectedRange.End.ToInt() >= r.End.ToInt()))
                            listRange.Add(rangePermis);
                    }
            }
            for (int i = 0; i < listRange.Count; i++)
            {
                var p = listRange[i].Range;
                for (int j = 0; j < listRange.Count; j++)
                {
                    var c = listRange[j].Range;
                    if (p != c)
                        if (p.Start.ToInt() <= c.Start.ToInt() && p.End.ToInt() >= c.End.ToInt())
                        {
                            //range cha bao gom range con
                            listRange.Remove(listRange[j]);
                            j--;
                            i--;
                        }
                }
                if (i < 0) i = 0;
            }
            if (listRange.Count == 0)
            {
                MessageBox.Show("Đặt con trỏ vào giữa vùng ký hay quét chọn vùng mà bạn muốn hủy.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            //todo khong lay duoc range con
            //foreach (var rangePermis in rangePermissions)
            //{
            //    if (rangePermis != range
            //        && range.Range.Start.ToInt() < rangePermis.Range.Start.ToInt()
            //        && range.Range.End.ToInt() > rangePermis.Range.End.ToInt())
            //    {
            //        MessageBox.Show("Không thể hủy ký vì có chữ ký khác thuộc trên nội dung này.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //        return;
            //    }
            //}
            //if (BOSApp.CurrentUser != BOSApp.CurrentUser)
            //{
            //    MessageBox.Show("Bạn không có quyền hủy ký trên vùng dữ liệu này.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}
            doc.EndUpdateRangePermissions(rangePermissions);

            //Kiem tra user huy ky co phai la user da ky
            if (!UserIsAdmin()) //Neu user co quyen admin thi bo qua kiem tra
            {
                bool bOK = false;
                foreach (var range in selectedRanges)
                {
                    var listEmrField = this._emrDocumentHelper.GetAllFields(doc, range);
                    foreach (var field in listEmrField)
                    {
                        var value = this._emrDocumentHelper.GetFieldValue(doc, field.Field);
                        var caretPosition = field.Field.ResultRange.Start.ToInt() + 1;
                        var properties = this._emrDocumentHelper.GetParamInfos(doc, caretPosition);
                        if (properties != null)
                        {
                            if (BOSApp.CurrentUsersInfo.ADUserName.ToLower() != properties.LastOrDefault(pro => pro.Key == "Người ký").Value?.ToLower()) //So sanh user ky cuoi cung vi co the ky de len
                            {
                                MessageBox.Show("Không thể hủy ký của người dùng khác!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                this._richEditCtrl.Modified = false;
                                return;
                            }
                            else bOK = true;
                        }
                    }
                }
                if (!bOK)
                {
                    var caretPosition = doc.CaretPosition.ToInt();
                    var properties = this._emrDocumentHelper.GetParamInfos(doc, caretPosition);
                    if (properties != null)
                    {
                        if (BOSApp.CurrentUsersInfo.ADUserName.ToLower() != properties.LastOrDefault(pro => pro.Key == "Người ký").Value?.ToLower()) //So sanh user ky cuoi cung vi co the ky de len
                        {
                            MessageBox.Show("Không thể hủy ký của người dùng khác!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            this._richEditCtrl.Modified = false;
                            return;
                        }
                    }
                }
            }

            var gui = new guiGetOneStr("Nhập lý do vì sao hủy ký phần này.");
            if (gui.ShowDialog() != DialogResult.OK) return;
            string unsignRemark = gui.Value;
            guiSign guiSign = new guiSign("Bạn muốn hủy ký phần này? Thao tác sẽ không thể hoàn tác.", _fingerPrintZK);
            if (guiSign.ShowDialog() != DialogResult.OK) return;

            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;

            BOSProgressBar.Start("Đang xử lý dữ liệu");
            this._tempRichEditCtrl.CreateNewDocument(false);
            var toDoc = this._tempRichEditCtrl.Document;
            var fields = new List<Field>();
            foreach (var item in listRange)
                toDoc.InsertDocumentContent(toDoc.Range.End, item.Range, InsertOptions.KeepSourceFormatting);

            //Lưu tài liệu lên server
            BOSProgressBar.SetText("Đang lưu và upload tập tin");
            doc.EndUpdate();
            //upload sign partial to ftp fo sign
            toDoc.EndUpdate();
            string toFileName = document.MEEmrDocumentFile + "_unsign" + DateTime.Now.ToBinary();
            string toFileFullname = toFileName + ".docx";
            string toFile = string.Format(@"{0}\Emr\Partials\{1}\{2}", _documentPath, document.FK_MEEmrID, toFileFullname);
            Directory.CreateDirectory(string.Format(@"{0}\Emr\Partials", _documentPath));
            Directory.CreateDirectory(string.Format(@"{0}\Emr\Partials\{1}", _documentPath, document.FK_MEEmrID));

            this.RemoveAllTagForPrint(this._tempRichEditCtrl, _entity.METemplate);
            this._emrDocumentHelper.RemoveAllParamMarkup(this._tempRichEditCtrl.Document);
            _tempRichEditCtrl.SaveDocument(toFile, DocumentFormat.OpenXml);
            _ftpFileMng.UploadFile($"/Emr/Partials/{document.FK_MEEmrID}/", toFileFullname, toFile);

            var ok = false;
            foreach (var item in listRange)
                if (strikeThrough)
                    ok = this._emrActionHelper.UnsignRangeStrikeThrough(item, BOSApp.CurrentUser);
                else
                    ok = this._emrActionHelper.UnsignRange(item, BOSApp.CurrentUser);
            if (!ok) return;
            string fileName = string.Format(@"{0}\Emr\{1}\{2}.docx", _documentPath, document.FK_MEEmrID, document.MEEmrDocumentFile);
            _richEditCtrl.SaveDocument(fileName, DocumentFormat.OpenXml);
            if (!EncryptFileDocument())
            {
                BOSProgressBar.Close();
                return;
            }
            if (!SaveEmrDocumentInfoAndUploadFile())
            {
                BOSProgressBar.Close();
                return;
            }
            //save sign history
            SaveSignHistory(toFileName, EmrDocumentFileExtention.docx, EmrDocumentSignType.Unsigned.ToString(), unsignRemark);
            //Clear history để không đc undo.
            ((DevExpress.XtraRichEdit.Model.DocumentModel)this._richEditCtrl.Model).History.Clear();
            CloneSignedRangeForTracking();
            _richEditCtrl.Modified = false;
            BOSProgressBar.Close();
        }
        /// <summary>
        /// chen chu ky image vao vi tri chu ky cuoi cung
        /// </summary>
        /// <param name="fields">cac field sign</param>
        private void InsertSignature(List<Field> fields)
        {
            if (BOSApp.CurrentEmployeesInfo.HREmployeeSignature == null)
                _msgNotification.Text = "Vui lòng cập nhật hình ảnh chữ ký vào thông tin nhân viên.";

            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            //web se truyen them role, nhung win thi chi can fields, vi fields da filter theo sign
            this._emrDocumentHelper.InsertSignature(fields,
                BOSApp.CurrentEmployeesInfo.HREmployeeSignature,
                BOSApp.CurrentEmployeesInfo.HREmployeeName.ToUpper(),
                BOSApp.CurrentEmployeesInfo.HREmployeeShortSignature,
                string.Empty,
                BOSApp.CurrentUser,
                AppMemCache.GetTemplateParams(document.FK_METemplateID));
        }
        private void SaveSignHistory(string toFileName, EmrDocumentFileExtention ext, string type, string remark)
        {
            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            var sign = new MEEmrDocumentSignsInfo()
            {
                FK_MEEmrDocumentID = document.MEEmrDocumentID,
                FK_HRDepartmentID = BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID,
                FK_HREmployeeID = BOSApp.CurrentEmployeesInfo.HREmployeeID,
                MEEmrDocumentSignTime = DateTime.Now,
                MEEmrDocumentSignFile = toFileName,
                MEEmrDocumentSignFileExt = ext.ToString(),
                MEEmrDocumentSignType = type,
                MEEmrDocumentSignRemark = remark,
                AACreatedUser = BOSApp.CurrentUser
            };
            this._emrDocumentSignCtrl.CreateObject(sign);
        }
        internal void GetDocumentSignHistory()
        {
            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            _entity.MEEmrDocumentSignsList.Invalidate(document.MEEmrDocumentID);
        }

        internal void ViewSignedDocument(MEEmrDocumentSignsInfo sign)
        {
            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            string fileName = string.Format(@"{0}\Emr\Signed\{1}\{2}{3}.{4}", _documentPath, document.FK_MEEmrID, sign.MEEmrDocumentSignFile, DateTime.Now.ToBinary().ToString(), sign.MEEmrDocumentSignFileExt);
            Directory.CreateDirectory(string.Format(@"{0}\Emr\Signed", _documentPath));
            Directory.CreateDirectory(string.Format(@"{0}\Emr\Signed\{1}", _documentPath, document.FK_MEEmrID));

            try
            {
                if (sign.MEEmrDocumentSignType != EmrDocumentSignType.DigitalSigned.ToString())
                {
                    _ftpFileMng.DownloadFile($"/Emr/Partials/{document.FK_MEEmrID}/", sign.MEEmrDocumentSignFile + "." + sign.MEEmrDocumentSignFileExt, fileName);
                }
                else
                {
                    if (sign.MEEmrDocumentSignFileExt == EmrDocumentFileExtention.pdf.ToString())
                    {
                        _ftpFileMng.DownloadFile($"/Emr/{document.FK_MEEmrID}/", sign.MEEmrDocumentSignFile + "." + sign.MEEmrDocumentSignFileExt, fileName);
                    }
                    else
                    {
                        _ftpFileMng.DownloadFile($"/Emr/Signed/{document.FK_MEEmrID}/", sign.MEEmrDocumentSignFile + "." + sign.MEEmrDocumentSignFileExt, fileName);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!File.Exists(fileName))
            {
                _msgNotification.Text = ("File bệnh án không tồn tại ở địa chỉ. " + fileName);
                _richEditCtrl.Visible = false;
                return;
            }
            // verify status
            var openFile = DialogResult.OK;
            if (sign.MEEmrDocumentSignType == EmrDocumentSignType.DigitalSigned.ToString())
            {
                var valid = false;
                var location = BOSApp.CurrentCompanyInfo.CSCompanyCaProvider == CaProviders.ESIGN_CA ? "LOCAL" : $"MÁY CHỦ {BOSApp.CurrentCompanyInfo.CSCompanyCaProvider} KÝ SỐ TRẢ VỀ";
                try
                {
                    BOSProgressBar.Start("Đang xác thực chữ ký số");
                    var configCA = BOSApp.GetSystemConfigValue(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.SYSTEM_CONFIGS_CA_METHOD);
                    var methodCA = !string.IsNullOrEmpty(configCA) ? configCA : string.Empty;
                    valid = _digitalSig.Verify(File.ReadAllBytes(fileName), sign.MEEmrDocumentSignFileExt, methodCA);
                }
                catch (SignerCustomException ex)
                {
                    BOSProgressBar.Close();
                    MessageBox.Show(location + ": XÁC THỰC KHÔNG THÀNH CÔNG" +
                        "\n Mã lỗi: " + ex.Code +
                        "\n Chi tiết: " + ex.Message,
                        "Xác thực không thành công", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    if (_notificationTab)
                    {
                        _msgLogs.Text += "\r\n DIGITAL VERIFY FAILED: " + ex.Code + "-" + ex.ToString();
                    }
                    else if (_logConfig)
                    {
                        _msgLogsTemp += "\r\n DIGITAL VERIFY FAILED: " + ex.Code + "-" + ex.ToString();
                    }
                }
                catch (Exception ex)
                {
                    BOSProgressBar.Close();
                    var notificationStr = _notificationTab ? " Xem chi tiết ở thông báo" : string.Empty;
                    MessageBox.Show(ex.ToString(), $"Lỗi không xác định khi gọi xác thực.{notificationStr}", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    if (_notificationTab)
                    {
                        _msgLogs.Text += "\r\n DIGITAL VERIFY FAILED: " + ex.ToString();
                    }
                    else if (_logConfig)
                    {
                        _msgLogsTemp += "\r\n DIGITAL VERIFY FAILED: " + ex.ToString();
                    }
                }
                finally
                {
                    BOSProgressBar.Close();
                }
                if (valid)
                {
                    var cerDetail = string.Empty;
                    if (sign.MEEmrDocumentSignFileExt == EmrDocumentFileExtention.pdf.ToString())
                        cerDetail = _digitalSig.GetCerDetailFromPdf(fileName);
                    else if (sign.MEEmrDocumentSignFileExt == EmrDocumentFileExtention.docx.ToString())
                        cerDetail = _digitalSig.GetCerDetailFromDocx(fileName);
                    openFile = MessageBox.Show(location + ": \u2714 TOÀN VẸN DỮ LIỆU VÀ CHỮ KÝ HỢP LỆ.\n\n" + cerDetail +
                        "\n OK: Để tải xuống và mở tập tin", "Xác thực toàn vẹn dữ liệu và thông tin chữ ký số", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                }
            }
            if (!string.IsNullOrEmpty(this._blockVerifyUrl))
            {
                var hash = _hashProvider.ComputeHash(fileName);

                PrintMgsLog("GOI-API-BLOCK", string.Empty);
                BOSProgressBar.Start("Đang xác thực block");
                var blockParams = new Dictionary<string, object>
                    {
                        { "digitalFingerprint", hash },
                    };
                try
                {
                    var dataBlock = (JObject)_apiBlock.Post(_blockVerifyUrl, null, blockParams);
                    if (dataBlock != null && dataBlock.ContainsKey("signature"))
                    {
                        MessageBox.Show("Nội dung ký đã được XÁC THỰC"
                            + "\n - Người ký: " + sign.MEEmrDocumentSignUser
                            + "\n - Hash: " + sign.MEEmrDocumentSignHash
                            + "\n - Block: " + sign.MEEmrDocumentSignBlockAddr
                            , "Xác thực nội dung đã ký", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Nội dung ký đã BỊ THAY ĐỔI"
                         + "\n - Người ký: " + sign.MEEmrDocumentSignUser
                         + "\n - Hash cũ: " + sign.MEEmrDocumentSignHash
                         + "\n - Hash mới: " + hash
                         + "\n - Block: " + sign.MEEmrDocumentSignBlockAddr
                         , "Nội dung ký không được xác thực", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex)
                {
                    PrintMgsLog("GOI-API-BLOCK-LOI", ex.ToString());
                }
                finally
                {
                    BOSProgressBar.Close();
                }
            }

            if (openFile != DialogResult.OK) return;

            var proc = new System.Diagnostics.Process
            {
                EnableRaisingEvents = false
            };
            proc.StartInfo.FileName = fileName;
            proc.Start();
        }
        #endregion Sign
        public override void ActionDelete()
        {
            if (IsEmrReadOnly())
            {
                MessageBox.Show("Chỉ được xóa các bệnh án ở trạng thái 'Đang nhập'");
                return;
            }
            var emr = this._entity.MainObject as MEEmrsInfo;

            var transfers = _transferCtrl.GetByEmrId(emr.MEEmrID);
            if (transfers.Count > 0)
            {
                MessageBox.Show("Tồn tại lịch sử chuyển khoa của bệnh án này. Cần xóa tất cả lịch sử chuyển khoa trước khi xóa bệnh án.", "Không xóa được", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            var shares = _shareCtrl.GetByEmrId(emr.MEEmrID);
            if (shares.Count > 0)
            {
                MessageBox.Show("Tồn tại lịch sử chia sẻ của bệnh án này. Cần xóa tất cả lịch sử chia sẻ trước khi xóa bệnh án.", "Không xóa được", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            var documents = _emrDocumentCtrl.GetByEmrId(emr.MEEmrID);
            if (documents.Count > 0)
            {
                MessageBox.Show("Tồn tại tờ bệnh án thuộc bệnh án này. Cần xóa tất cả tờ bệnh án trước khi xóa bệnh án.", "Không xóa được", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            guiSign guiSign = new guiSign("Bạn chắc chắn xóa bệnh án này. Thao tác sẽ không thể hoàn tác.", _fingerPrintZK);
            if (guiSign.ShowDialog() == DialogResult.OK)
            {
                base.ActionDelete();
                var tb = this.Toolbar.ObjectCollection.Tables[0];
                var gridView = (_searchGrid.MainView as GridView);
                if (tb.Rows.Count > 0)
                {
                    gridView.FocusedRowHandle = 0;
                    this.Invalidate(int.Parse(tb.Rows[0]["MEEmrID"].ToString()));
                }
                else
                {
                    this.Invalidate(0);
                    InvalidateEmrDocumentList(0);
                }
            }
        }
        public override bool Delete(int iObjectId)
        {
            SaveObjectHistory(cstObjectHistoryActionDelete, Toolbar.CurrentObjectID);
            CurrentModuleEntity.Delete(iObjectId);
            QuickSearch();
            return true;
        }
        internal void DeleteEmrDocument(MEEmrDocumentsInfo row)
        {
            if (row == null) return;
            if (IsEmrReadOnly()) return;
            if (row.FK_HREmployeeCreatedID != BOSApp.CurrentEmployeesInfo.HREmployeeID)
            {
                MessageBox.Show("Không thể xóa tờ bệnh án do người khác tạo. ", "Không thể xóa tờ bệnh án này", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            GetDocumentSignHistory();
            foreach (var item in _entity.MEEmrDocumentSignsList)
            {
                if (item.FK_HREmployeeID != BOSApp.CurrentEmployeesInfo.HREmployeeID)
                {
                    MessageBox.Show("Đã có người dùng khác ký trên tờ bệnh án này. Không thể xóa tờ bệnh án.", "Không thể xóa tờ bệnh án này", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }
            }
            GetDocumentDataHistory();
            foreach (var item in _entity.MEEmrDocumentDataHistoriesList)
            {
                if (!string.IsNullOrEmpty(item.AAUpdatedUser) && item.AAUpdatedUser != BOSApp.CurrentUser)
                {
                    MessageBox.Show("Đã có người dùng khác thao tác trên tờ bệnh án này. Không thể xóa tờ bệnh án.", "Không thể xóa tờ bệnh án này", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }
            }
            guiSign guiSign = new guiSign("Đang thực hiện xóa tờ bệnh án. Sẽ không thể hoàn tác. Vẫn thực hiện?", _fingerPrintZK);
            if (guiSign.ShowDialog() == DialogResult.OK)
            {
                var documentCode = getDocumentNo(row);
                var remark = $"xóa tờ bệnh án [{documentCode}]";

                BackgroundEmrDocumentValidatesDelete(_entity.MainObject as MEEmrsInfo, row.FK_METemplateID, row.MEEmrDocumentID);

                _emrDocumentCtrl.DeleteObject(row.MEEmrDocumentID);

                HistoryEmr(_entity.MainObject as MEEmrsInfo, cstObjectHistoryActionChange, remark);

                this.OpenGuidance();
                this._emrDocumentMng.Delete(row.MEEmrDocumentMongoID, row.MEEmrDocumentNo, BOSApp.CurrentUsersInfo.ADUserName);
                this.InvalidateEmrDocumentList(row.FK_MEEmrID);
            }
        }
        private void CloseEmrDocument(MEEmrDocumentsInfo item, bool isHistory)
        {
            item.MEEmrDocumentEndDate = DateTime.Now;
            item.MEEmrDocumentPreStatus = item.MEEmrDocumentStatus;
            item.MEEmrDocumentStatus = EmrDocumentStatus.Closed.ToString();
            if (item.MEEmrDocumentFileExt == EmrDocumentFileExtention.docx.ToString())
                item.MEEmrDocumentFileExt = EmrDocumentFileExtention.pdf.ToString();
            item.MEEmrDocumentJson = string.Empty;//this.ParserDocumentToJson(); uthv cham qua roi

            item = ClearCurrentEditingUser(item);

            var emr = GetCurrentMainObject();
            UpdateMongoDocument(emr, item, new List<string>() { "MEEmrDocumentContent" });
            item.AAUpdatedUser = BOSApp.CurrentUser;
            _emrDocumentCtrl.UpdateObject(item);

            var documentCodeU = getDocumentNo(item);
            if (isHistory)
            {
                HistoryEmr(emr, "Change", $"{documentCodeU} đóng tờ");
            }
        }

        private void UndoCloseEmrDocument(MEEmrDocumentsInfo currentItem)
        {
            _emrDocumentCtrl.UpdateObject(currentItem);

            var documentCodeU = getDocumentNo(currentItem);
            HistoryEmr(GetCurrentMainObject(), "Change", $"{documentCodeU} khôi phục đóng tờ");
        }

        public void CloseCurrentEmrDocument()
        {
            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            if (document.MEEmrDocumentID == 0)
            {
                MessageBox.Show("Vui lòng chọn một tờ bệnh án", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (IsEmrReadOnly()) return;
            if (!TakeEditingPermission(document)) return;
            if (!IsNothingToSaveDocumentContent())
                return;
            //truong hop nay la khi nguoi dung bam NO hoac YES nhung luu khong duoc, can mo lai to benh an
            if (this._richEditCtrl.Modified)
            {
                MessageBox.Show("Lưu tờ bệnh án trước khi thực hiện đóng. Nếu vi phạm quy định ký cần hoàn tác các nội dung vi phạm.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var missingDataParam = ValidateDocumentRequired(document, true);
            if (missingDataParam.Count > 0)
            {
                var msgValidates = string.Join("\n", missingDataParam.Select(x => "<" + x.Value + ">").ToArray());
                MessageBox.Show($"{msgValidates} chưa có dữ liệu.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var template = _templateCtrl.GetObjectByID(document.FK_METemplateID) as METemplatesInfo;
            if (!CheckDataIntegrityFromLastDigitalSignDocx(document, template.METemplateName)) return;

            // Check File
            string fileNameLocal = string.Format(@"{0}\Emr\{1}\{2}.{3}", _documentPath, document.FK_MEEmrID, document.MEEmrDocumentFile, document.MEEmrDocumentFileExt);
            if (!File.Exists(fileNameLocal))
            {
                MessageBox.Show("Tờ bệnh án không tìm thấy.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            long length = new System.IO.FileInfo(fileNameLocal).Length;
            if (length == 0)
            {
                MessageBox.Show("Tờ bệnh án lỗi.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            guiSign guiSign = new guiSign("Đóng tờ bệnh án? Thao tác sẽ không thể hoàn tác.", _fingerPrintZK);
            if (guiSign.ShowDialog() == DialogResult.OK)
            {
                if (_entity.ModuleObjects[TableName.MEEmrDocumentsTableName] == null) return;
                var extension = document.MEEmrDocumentFileExt;
                if (extension == EmrDocumentFileExtention.docx.ToString())
                {
                    this.RemoveAllForPrint(this._richEditCtrl, document.FK_METemplateID);
                    string fileName = string.Format(@"{0}\Emr\{1}\{2}.{3}", _documentPath, document.FK_MEEmrID, document.MEEmrDocumentFile, EmrDocumentFileExtention.pdf.ToString());
                    this._richEditCtrl.ExportToPdf(fileName);
                    _ftpFileMng.UploadFile($"/Emr/{document.FK_MEEmrID}/", document.MEEmrDocumentFile + "." + EmrDocumentFileExtention.pdf.ToString(), fileName);
                }
                this.CloseEmrDocument(document, true);

                _entity.MEEmrDocumentsList.ChangeObjectFromList();
                _entity.MEEmrDocumentsList.GridControl.Refresh();
                InvalidateDocument(document, false);

                _msgNotification.Text = "Tờ bệnh án này đã đóng. Mọi thay đổi sẽ không thể lưu được.";
            }
        }

        /// <summary>
        /// dong toan bo cac file benh an trong emr nay
        /// </summary>
        public void CloseAllEmr()
        {
            if (!IsNothingToSaveDocumentContent())
            {
                return;
            }
            var emr = this._entity.MainObject as MEEmrsInfo;
            if (IsEmrReadOnly(true)) return;
            InvalidateEmrDocumentListForAdmin(emr);
            //_entity.MEEmrDocumentsList.Invalidate(emr.MEEmrID);
            foreach (var doc in _entity.MEEmrDocumentsList)
            {
                if (doc.FK_EditingUserID > 0 && doc.FK_EditingUserID != BOSApp.CurrentEmployeesInfo.HREmployeeID)
                {
                    ShowDocumentEditingUser(doc);
                    return;
                }
            }

            if (!ValidateBeforeActionEmr(EmrStatus.Closed.ToString()))
            {
                return;
            }
            guiSign guiSign = new guiSign("Đang thực hiện đóng bệnh án và chuyển sang pdf. Sẽ không thể hoàn tác. Vẫn thực hiện?", _fingerPrintZK);
            if (guiSign.ShowDialog() == DialogResult.OK)
            {
                #region LKBA - HTSS. Important, config enable?
                try
                {
                    RelationSubmit(emr);

                    InvalidateEmrDocumentListForAdmin(emr);
                }
                catch (Exception ex)
                {
                    if (_notificationTab)
                    {
                        _msgLogs.Text = $"Lỗi xử lý bệnh án liên kết {emr.MEEmrNo}. Chi tiết lỗi: {Environment.NewLine}{ex}";
                    }
                    else if (_logConfig)
                    {
                        _msgLogsTemp += $"{Environment.NewLine}Lỗi xử lý bệnh án liên kết {emr.MEEmrNo}. Chi tiết lỗi: {Environment.NewLine}{ex}";
                    }
                }
                #endregion

                this._dpnViewPdf.Visible = false;
                this._dpnRichEdit.Visible = true;
                this.DockManager.ActivePanel = this._dpnRichEdit;
                BOSProgressBar.Start("Đang đóng bệnh án");
                emr.AllowPropertyChangedEvent = false;
                var currentItem = new MEEmrDocumentsInfo(); // use for UndoCloseEmrDocument()
                try
                {
                    var successMsg = "Bệnh án này đã đóng. Mọi thay đổi sẽ không thể lưu được.";
                    var isActionLocal = false;
                    if (_apiEmr != null)
                    {
                        var actionUri = BOSApp.GetSystemConfigValue(SysCfgConsts.EMR_API_ENDPOINT, SysCfgConsts.EMR_CLOSE);
                        if (!string.IsNullOrEmpty(actionUri))
                        {
                            PrintMgsLog("BAT-DAU-GOI-API", actionUri);
                            var body = new { emrId = emr.MEEmrID };
                            var response = _apiEmr.Post<Emr.Base.Models.Abp.AjaxResponse, IDictionary<string, object>>(actionUri, null, body);
                            PrintMgsLog("KET-THUC-GOI-API", actionUri);
                            if (response != null)
                            {
                                if (response.Success)
                                {
                                    _msgNotification.Text = successMsg;
                                    _entity.MEEmrDocumentsList.Invalidate(emr.MEEmrID);
                                    emr.MEEmrStatus = response.Result["MEEmrStatus"].ToString();
                                    InactiveShareEmr(emr.MEEmrID);
                                }
                                else
                                {
                                    if (_notificationTab)
                                    {
                                        _msgLogs.Text = (response.Error.Message);
                                    }
                                    else if (_logConfig)
                                    {
                                        _msgLogsTemp += $"\r\n {(response.Error.Message)}";
                                    }

                                    isActionLocal = true;
                                    if (MessageBox.Show("Tiếp tục thực hiện đóng bệnh án trên máy? \nĐóng bệnh án thất bại trên máy chủ EMR API. Xem chi tiết ở thông báo.", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.Cancel)
                                        return;
                                }
                            }
                            else
                            {
                                if (_notificationTab)
                                {
                                    _msgLogs.Text = ("Không có dữ liệu trả về từ api. Liên hệ quản trị viên để biết thêm chi tiết");
                                }
                                else if (_logConfig)
                                {
                                    _msgLogsTemp += $"\r\n {("Không có dữ liệu trả về từ api. Liên hệ quản trị viên để biết thêm chi tiết")}";
                                }
                                isActionLocal = true;
                            }
                        }
                        else
                        {
                            //_msgLogs.Text = ("Không cấu hình đóng bệnh án bằng API");
                            isActionLocal = true;
                        }
                    }
                    else
                    {
                        //_msgLogs.Text = ("Không kết nối được đến máy chủ EMR API");
                        isActionLocal = true;
                    }

                    if (isActionLocal)
                    {
                        int totalClosed = 0;
                        var serverPath = $"/Emr/{emr.MEEmrID}/";
                        var serverFiles = _ftpFileMng.GetNameListing(serverPath);
                        var localPath = string.Format(@"{0}\Emr\{1}\", _documentPath, emr.MEEmrID);
                        DirectoryInfo localDir = new DirectoryInfo(localPath);
                        var pdf = EmrDocumentFileExtention.pdf.ToString();
                        var docx = EmrDocumentFileExtention.docx.ToString();
                        var hashFiles = new List<string>();
                        foreach (var item in _entity.MEEmrDocumentsList)
                        {
                            //InitDocumentSession(item.FK_METemplateID);
                            //_entity.InvalidateModuleObject(item);
                            // bo qua cac to an va huy
                            if (item.MEEmrDocumentStatus == EmrDocumentStatus.Discarded.ToString()
                                || item.MEEmrDocumentStatus == EmrDocumentStatus.Hidden.ToString())
                            {
                                hashFiles.Add(item.MEEmrDocumentFile);
                                totalClosed++;
                                continue;
                            }
                            var fileNamePdf = $"{item.MEEmrDocumentFile}.{pdf}";
                            var fileNameDocx = $"{item.MEEmrDocumentFile}.{docx}";
                            var fullFileNamePdfLocal = string.Format(@"{0}\{1}", localPath, fileNamePdf);
                            var fullFileNameDocxLocal = string.Format(@"{0}\{1}", localPath, fileNameDocx);
                            if (item.MEEmrDocumentFileExt == docx)
                            {
                                _ftpFileMng.DownloadFile(serverPath, fileNameDocx, fullFileNameDocxLocal);

                                if (!CheckDataIntegrityFromLastDigitalSignDocx(item, item.MEEmrDocumentFile)) continue;

                                currentItem = item.Clone() as MEEmrDocumentsInfo;

                                var hash = _md5Hasher.ComputeHash(fullFileNameDocxLocal);
                                var filePdfHash = $"{item.MEEmrDocumentFile}_MD5{hash}.{pdf}";
                                var existFileHash = serverFiles.Where(stringToCheck => stringToCheck.Contains(filePdfHash));
                                if (existFileHash.Count() == 0)
                                {
                                    var fi = new FileInfo(Path.Combine(localPath, filePdfHash));
                                    if (fi != null && fi.Exists && fi.Length > 10000)
                                    {
                                        File.Delete(fullFileNamePdfLocal);
                                        File.Move(Path.Combine(localPath, filePdfHash), fullFileNamePdfLocal);
                                    }
                                    else
                                    {
                                        this.OpenEmrDocument(item.FK_MEEmrID, item.MEEmrDocumentFile, false, false);
                                        this.RemoveAllForPrint(this._richEditCtrl, item.FK_METemplateID);
                                        this._richEditCtrl.ExportToPdf(fullFileNamePdfLocal);
                                    }
                                    _ftpFileMng.UploadFile(serverPath, fileNamePdf, fullFileNamePdfLocal);
                                }
                                else
                                {
                                    _ftpFileMng.DownloadFile(serverPath, filePdfHash, fullFileNamePdfLocal);
                                    var fi = new FileInfo(Path.Combine(localPath, fileNamePdf));
                                    if (fi != null && fi.Exists && fi.Length > 10240)
                                    {
                                        _ftpFileMng.MoveFile(serverPath, filePdfHash, fileNamePdf);
                                    }
                                    else
                                    {
                                        this.OpenEmrDocument(item.FK_MEEmrID, item.MEEmrDocumentFile, false, false);
                                        this.RemoveAllForPrint(this._richEditCtrl, item.FK_METemplateID);
                                        this._richEditCtrl.ExportToPdf(fullFileNamePdfLocal);
                                        _ftpFileMng.UploadFile(serverPath, fileNamePdf, fullFileNamePdfLocal);
                                    }
                                }
                                this.CloseEmrDocument(item, false);
                                hashFiles.Add(item.MEEmrDocumentFile);
                            }
                            else
                            {
                                if (!_ftpFileMng.FileExists($"/Emr/{item.FK_MEEmrID}/", fileNamePdf))
                                {
                                    throw new FtpException();
                                }

                                // Bỏ qua TH file pdf người dùng upload < 10 kbs
                                _ftpFileMng.DownloadFile(serverPath, fileNamePdf, fullFileNamePdfLocal);
                                var fi = new FileInfo(Path.Combine(localPath, fileNamePdf));
                                if (fi != null && fi.Exists && fi.Length < 10240)
                                {
                                    if (_ftpFileMng.FileExists($"/Emr/{item.FK_MEEmrID}/", $"{item.MEEmrDocumentFile}.{docx}"))
                                    {
                                        this.OpenEmrDocument(item.FK_MEEmrID, item.MEEmrDocumentFile, false, false);
                                        this.RemoveAllForPrint(this._richEditCtrl, item.FK_METemplateID);
                                        this._richEditCtrl.ExportToPdf(fullFileNamePdfLocal);
                                        _ftpFileMng.UploadFile(serverPath, fileNamePdf, fullFileNamePdfLocal);
                                    }
                                }

                                item.MEEmrDocumentEndDate = DateTime.Now;
                                item.MEEmrDocumentPreStatus = item.MEEmrDocumentStatus;
                                item.MEEmrDocumentStatus = EmrDocumentStatus.Closed.ToString();
                                if (item.MEEmrDocumentFileExt == EmrDocumentFileExtention.docx.ToString())
                                    item.MEEmrDocumentFileExt = EmrDocumentFileExtention.pdf.ToString();
                                item.AAUpdatedUser = BOSApp.CurrentUser;
                                _emrDocumentCtrl.UpdateObject(item);

                                var documentCodeU = getDocumentNo(item);
                            }
                            totalClosed++;
                        }
                        _entity.MEEmrDocumentsList.Invalidate(emr.MEEmrID);
                        if (totalClosed == _entity.MEEmrDocumentsList.Count)
                        {
                            #region Change End Date
                            var emrEndDate = DateTime.Now;
                            //var lastHistoryReOpen = _geObjHistoryCtrl.GetLatestHistoryByObjectNameAndObjectId(TableName.MEEmrsTableName, emr.MEEmrID, cstObjectHistoryActionReOpen);
                            if (emr.MEEmrEndDate.Ticks < 3155378975999970000)
                            {
                                BOSProgressBar.Close();
                                var emrEndDateDb = emr.MEEmrEndDate;
                                var gui = new guiUpdateEmrEndDate(emrEndDate)
                                {
                                    Module = this
                                };
                                gui.InitializeControls(gui.Controls);
                                gui.StartPosition = FormStartPosition.CenterParent;
                                if (gui.ShowDialog() == DialogResult.OK)
                                {
                                    emrEndDate = gui._emrEndDate;
                                }
                                else
                                {
                                    emrEndDate = emrEndDateDb;
                                }
                            }
                            #endregion
                            // emr = _emrCtrl.GetObjectByID(emr.MEEmrID) as MEEmrsInfo;
                            // them cot nguoi dong benh an theo yeu cau cua DKLK
                            emr.FK_HREmployeeClosedID = BOSApp.CurrentEmployeesInfo.HREmployeeID;
                            emr.MEEmrStatus = EmrStatus.Closed.ToString();
                            emr.MEEmrEndDate = emrEndDate;
                            emr.AAUpdatedUser = BOSApp.CurrentUser;
                            this._emrCtrl.UpdateObject(emr);
                            InactiveShareEmr(emr.MEEmrID);

                            #region Histories: Update if ReOpen
                            var historyReOpen = _geObjHistoryCtrl.GetLatestHistoryByActionStatus(TableName.MEEmrsTableName, emr.MEEmrID, cstObjectHistoryActionReOpen, "New");
                            if (historyReOpen != null)
                            {
                                historyReOpen.GEObjectHistoryStatus = "Inactive";
                                _geObjHistoryCtrl.UpdateObject(historyReOpen);
                            }
                            #endregion

                            #region Histories
                            var hisLog = "Đóng bệnh án";
                            if (emr.MEEmrEndDate.Date != DateTime.Now.Date)
                            {
                                hisLog += $" - Ngày đóng {emr.MEEmrEndDate.ToString("dd/MM/yyyy HH:mm:ss")}";
                            }
                            HistoryEmr(emr, cstObjectHistoryActionChange, $"{hisLog}");
                            #endregion

                            _msgNotification.Text = successMsg;
                        }
                        else
                        {
                            MessageBox.Show("Bệnh án chưa được đóng vì có tờ bệnh án chưa đóng." +
                                "\nThực hiện đóng từng tờ bệnh án còn mở và sau đó tiếp tục đóng bệnh án.",
                                "Bệnh án chưa được đóng", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }

                        // Delete file hash
                        serverFiles = _ftpFileMng.GetNameListing(serverPath); // get newest
                        foreach (var hashFile in hashFiles)
                        {
                            ClearHashFileServer(serverPath, serverFiles, hashFile, string.Empty);
                            ClearHashFileLocal(localDir, hashFile);
                        }
                    }

                    #region TTBA
                    if (_sysHelper.AllowThread())
                    {
                        var thSum = new System.Threading.Thread(() => SummaryEmr(emr));
                        thSum.Start();
                    }
                    else
                    {
                        SummaryEmr(emr);
                    }
                    #endregion

                    RefreshCurrentEmrSearchResultsControl(emr);
                    this.Invalidate(emr.MEEmrID);
                }
                catch (Exception ex)
                {
                    UndoCloseEmrDocument(currentItem);
                    if (_notificationTab)
                    {
                        _msgLogs.Text += "\r\n CLOSE EMR FAILED: " + ex.ToString();
                    }
                    else if (_logConfig)
                    {
                        _msgLogsTemp += "\r\n CLOSE EMR FAILED: " + ex.ToString();
                    }
                    _msgNotification.Text = "Đóng bệnh án không thành công. Thực hiện đóng từng tờ bệnh án còn mở. Xem chi tiết ở thông báo.";
                    this.InvalidateEmrDocumentList(emr.MEEmrID);
                }
                finally
                {
                    emr.AllowPropertyChangedEvent = true;
                    BOSProgressBar.Close();
                }
            }
        }

        public string GetDocumentJsonData()
        {
            var docInfo = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            if (docInfo.MEEmrDocumentStatus == EmrDocumentStatus.Closed.ToString()
                || docInfo.MEEmrDocumentFileExt == EmrDocumentFileExtention.pdf.ToString())
            {
                return docInfo.MEEmrDocumentJson;
            }
            return this.ParserDocumentToJson();
        }
        /// <summary>
        /// duyet toan bo emr trong benh an nay
        /// </summary>
        public string ParserDocumentToJson()
        {
            try
            {
                var doc = this._richEditCtrl.Document;
                var docInfo = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
                //du lieu nay da dc cache khi load module neu cache=true InitEmrModuleSessionAsync
                var dictParams = AppMemCache.GetParamsDictKeyNo();
                //du lieu nay da dc lay tu cache/db la do AppMemCache quan ly
                var templateParams = AppMemCache.GetTemplateParamsDictPath(docInfo.FK_METemplateID);
                return this._emrParser.ParserFieldsToJsonV3(this._emrDocumentHelper.GetAllDataFieldInRangeOrDocument(), dictParams, templateParams);
            }
            catch (Exception)
            {
                _msgNotification.Text = ("Có lỗi khi trích xuất dữ liệu. Vui lòng kiểm tra thông báo lỗi");
            }
            return string.Empty;
        }

        public bool IsAllowEditField(DocumentPosition position)
        {
            return _emrDocumentHelper.IsAllowEditField(position);
        }
        public bool IsAllowEditPositionWithFlashNotification(DocumentPosition position)
        {
            var allow = _emrDocumentHelper.IsAllowEditField(position);
            if (!allow)
            {
                ShowFlashNotification("Thẻ dữ liệu đã bị chặn sửa thủ công", 3000);
            }
            return allow;
        }
        /// <summary>
        /// kiem tra emr nay co phai dang bi dong
        /// kiem tra emr co dang con ton tai 
        /// </summary>
        /// <returns></returns>
        public bool IsEmrReadOnly(bool onlyCheckEmr = false)
        {
            var emr = GetCurrentMainObject();
            if (emr.MEEmrID == 0) return false;
            //lay tu db de dam bao khong co truong hop them to vao benh an da dong boi user khac
            emr = _emrCtrl.GetObjectByID(emr.MEEmrID) as MEEmrsInfo;
            if (emr == null)
            {
                MessageBox.Show("Bệnh án không còn tồn tại. Có thể đã bị xóa hoặc trộn vào bệnh án khác.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return true;
            }
            if (emr.MEEmrStatus == EmrStatus.Closed.ToString())
                return true;
            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;

            if (!onlyCheckEmr && document != null && document.MEEmrDocumentStatus == EmrDocumentStatus.Closed.ToString())
                return true;
            if (!IsAllowEdit(emr)) return true;
            return false;
        }

        private void CopyTemplateToDocument(int templateId, string mEEmrDocumentFile, int emrId)
        {
            var template = _templateCtrl.GetObjectByID(templateId) as METemplatesInfo;
            var formatDt = "yyyyMMddHHmmssfff";
            string fileName = string.Format(@"{0}\Template\{1}_{2}.docx", _documentPath, template.METemplateNo, template.AAUpdatedDate.ToString(formatDt));
            if (!File.Exists(fileName))
            {
                try
                {
                    PrintMgsLog("BAT-DAU-DOWNLOAD-FILE", fileName);
                    _ftpFileMng.DownloadFile("/Template/", template.METemplateNo + ".docx", fileName);
                    PrintMgsLog("KET-THUC-DOWNLOAD-FILE", fileName);
                }
                catch (Exception ex)
                {
                    PrintMgsLog("LOI-KHONG-TAI-DUOC-FILE", fileName);
                    Trace.TraceError("FtpFileMng.DownloadFile ERROR: {0}:{1}:{2}:{3}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), BOSApp.CurrentUser, ex, fileName);
                    Trace.Flush();
                    MessageBox.Show($"Không tải được tờ bệnh án từ máy chủ.",
                         "CÓ LỖI XẢY RA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            if (!File.Exists(fileName))
            {
                MessageBox.Show("File mẫu bệnh án không tồn tại ở địa chỉ. " + fileName);
                return;
            }
            string toFileName = string.Format(@"{0}\Emr\{1}\{2}.docx", _documentPath, emrId, mEEmrDocumentFile);
            File.Copy(fileName, toFileName);
        }
        private void CopyFileToEmr(string filePath, string toFileName)
        {
            if (!File.Exists(filePath))
            {
                MessageBox.Show("File không tồn tại ở địa chỉ. " + filePath);
                return;
            }
            File.Copy(filePath, toFileName);
        }
        private void OpenEmrTemplate(string mEEmrDocumentNo)
        {
            string fileName = string.Format(@"{0}\Template\{1}.docx", _documentPath, mEEmrDocumentNo);
            _ftpFileMng.DownloadFile("/Template/", mEEmrDocumentNo + ".docx", fileName);
            if (!File.Exists(fileName))
            {
                MessageBox.Show("File mẫu bệnh án không tồn tại ở địa chỉ. " + fileName);
                return;
            }
            _richEditCtrl.LoadDocument(fileName);
        }
        private void OpenEmrDocument(int emrId, string mEEmrDocumentFile, bool readOnly = false, bool mustDowload = true, bool asyncMode = false)
        {
            PrintMgsLog("BAT-DAU-MO-FILE", mEEmrDocumentFile);

            this._emrDocumentHelper.ClearCacheBindingFields();

            _richEditCtrl.Visible = true;
            _richEditCtrl.CreateNewDocument(false);
            _msgMaximumPage = string.Empty;
            this._pdfViewer.CloseDocument();
            string fileName = string.Format(@"{0}\Emr\{1}\{2}.docx", _documentPath, emrId, mEEmrDocumentFile);
            if (mustDowload)
            {
                PrintMgsLog("BAT-DAU-DOWNLOAD-FILE", fileName);
                try
                {
                    _ftpFileMng.DownloadFile($"/Emr/{emrId}/", mEEmrDocumentFile + ".docx", fileName);
                    var fi = new FileInfo(fileName);
                    if (fi.Length == 0)
                    {
                        RequestRecoveryDocument(emrId, mEEmrDocumentFile);
                    }
                }
                catch (Exception)
                {
                    if (!RequestRecoveryDocument(emrId, mEEmrDocumentFile))
                    {
                        throw;
                    }
                }
                PrintMgsLog("KET-THUC-DOWNLOAD-FILE", fileName);
            }

            /*if (!File.Exists(fileName))
            {
                _msgNotification.Text = ("File bệnh án không tồn tại ở địa chỉ. " + fileName);
                _richEditCtrl.Visible = false;
                return;
            }*/
            try
            {
                PrintMgsLog("BAT-DAU-GIAI-MA-FILE", fileName);
                using (OfficeOpenXmlCrypto.OfficeCryptoStream stream = OfficeOpenXmlCrypto.OfficeCryptoStream.Open(fileName, this._emrDocumentHelper.ShareEmrPassword))
                {
                    _richEditCtrl.LoadDocument(stream, DocumentFormat.OpenXml);
                    PrintMgsLog("KET-THUC-GIAI-MA-FILE", fileName);
                    if (readOnly) _richEditCtrl.ReadOnly = true;
                    else
                    {
                        _richEditCtrl.ReadOnly = false;
                        var start = DateTime.Now;
                        var tg = this._emrDocumentHelper.GenCacheBindingFields(stream);
                        //truong hop tao moi to benh an, thao tac nay phai sync
                        if (!asyncMode)
                            tg.Wait(10000);
                        PrintMgsLog("GenCacheBindingFields", (DateTime.Now - start).TotalMilliseconds.ToString());

                        start = DateTime.Now;
                        var ts = this.CloneSignedRangeForTracking(stream);
                        if (!asyncMode && ts != null)
                            ts.Wait(10000);
                        PrintMgsLog("CloneSignedRangeForTracking", (DateTime.Now - start).TotalMilliseconds.ToString());
                    }
                    //quyen nay override moi quyen khac
                    //chi app dung cho benh an, ho so benh nhan duoc quyen edit
                    if (_overrideEditPermission == false && emrId == (_entity.MainObject as MEEmrsInfo).MEEmrID)
                    {
                        ShowFlashNotification("Tờ bệnh án được mở ở chế độ CHỈ ĐỌC. Không thể chỉnh sửa vì không có quyền.", 5000);
                        _richEditCtrl.ReadOnly = true;
                    }
                }
            }
            catch (OfficeOpenXmlCrypto.InvalidPasswordException)
            {
                _msgNotification.Text = ("Không giải mã được tờ bệnh án. Đã có lỗi trong quá trình mã hóa và upload tờ bệnh án. Không mở được tập tin.");
                RequestRecoveryDocument(emrId, mEEmrDocumentFile);
            }
            catch
            {
                RequestRecoveryDocument(emrId, mEEmrDocumentFile);
            }
            PrintMgsLog("MO-FILE-THANH-CONG", mEEmrDocumentFile);
        }

        private bool RequestRecoveryDocument(int emrId, string file)
        {
            if (MessageBox.Show("Vui lòng kiểm tra kết nối. " +
                        "\n\nNếu kết nối vẫn đang ổn định vui lòng bấm [RETRY] ĐỂ [GỞI YÊU CẦU PHỤC HỒI] TỜ BỆNH ÁN và quay lại sau ít phút.",
                        "KHÔNG TẢI ĐƯỢC TỜ BỆNH ÁN TỪ MÁY CHỦ", MessageBoxButtons.RetryCancel, MessageBoxIcon.Asterisk) == DialogResult.Retry)
            {
                var document = _emrDocumentCtrl.GetDocumentByFileName(emrId, file);
                var request = BOSApp.InvokeSignalHub("RequestRecoveryDocument", emrId, file, document.MEEmrDocumentHash, BOSApp.CurrentUser, _macAddress, _ipAddress);
                //TODO write to server to get when click manual if signalr not connect
                Task.Factory.StartNew(async () =>
                {
                    var result = await BOSApp.OnRequestRecoveryDocument(emrId, file, document.MEEmrDocumentHash);
                    if (result == -1) return;
                    if (result == 1)
                    {
                        BOSApp.OnCommonSignalMessage(EmrConsts.SYS_HUB_SENDER, "Đã phục hồi thành công");
                    }
                });
                return request;
            }
            return false;
        }

        private void OpenEmrPdf(int emrId, string file, bool mustDownload = true)
        {
            string filePath = string.Format(@"{0}\Emr\{1}\{2}.{3}", _documentPath, emrId, file, EmrDocumentFileExtention.pdf.ToString());
            if (mustDownload)
            {
                _ftpFileMng.DownloadFile($"/Emr/{emrId}/", file + "." + EmrDocumentFileExtention.pdf.ToString(), filePath);
            }
            if (!File.Exists(filePath))
            {
                if (_notificationTab)
                {
                    _msgLogs.Text += "\r\n" + ("File bệnh án không tồn tại ở địa chỉ. " + filePath);
                }
                else if (_logConfig)
                {
                    _msgLogsTemp += "\r\n" + ("File bệnh án không tồn tại ở địa chỉ. " + filePath);
                }
                return;
            }
            using (FileStream stream = new FileStream(filePath, FileMode.Open))
            {
                _pdfViewer.DetachStreamAfterLoadComplete = true;
                _pdfViewer.RotationAngle = 0;
                _pdfViewer.LoadDocument(stream);
            }
            this._richEditCtrl.CreateNewDocument(false);
        }
        #region Tracking change on signed range
        /// <summary>
        /// uthv
        /// Clone signed range for tracking change
        /// all signed range can't be update by other people
        /// </summary>
        private Task<int> CloneSignedRangeForTracking(MemoryStream inStream = null)
        {
            _signedContent = string.Empty;
            _signedImgs = new List<DevExpress.Office.Utils.OfficeImage>();
            var doc = this._richEditCtrl.Document;
            var rangePermissions = doc.BeginUpdateRangePermissions();
            var countRange = rangePermissions.Count;
            doc.EndUpdateRangePermissions(rangePermissions);
            if (countRange == 0) return null;
            var imgs = doc.Images.ToArray();
            var stream = new MemoryStream();
            if (inStream != null)
            {
                inStream.Position = 0;
                inStream.CopyTo(stream);
            }
            else
            {
                this._richEditCtrl.SaveDocument(stream, DocumentFormat.OpenXml);
            }
            var taskCode = Task.Factory.StartNew<int>(() =>
            {
                using (var rich = new RichEditDocumentServer())
                {
                    stream.Position = 0;
                    rich.LoadDocument(stream, DocumentFormat.OpenXml);
                    var cloneDoc = rich.Document;
                    rangePermissions = cloneDoc.BeginUpdateRangePermissions();
                    foreach (var rang in rangePermissions)
                    {
                        //khong cho nguoi do sua noi dung cua chinh ho
                        //if (rang.UserName != BOSApp.CurrentUsersInfo.ADUserName)
                        this._signedContent += cloneDoc.GetText(rang.Range);
                        foreach (var img in imgs)
                        {
                            if (img.Range.Start >= rang.Range.Start
                            && img.Range.End <= rang.Range.End)
                            {
                                _signedImgs.Add(img.Image);
                            }
                        }
                    }
                    cloneDoc.EndUpdateRangePermissions(rangePermissions);
                    foreach (var item in cloneDoc.Comments)
                    {
                        SubDocument sub = item.BeginUpdate();
                        this._signedContent += sub.GetText(sub.Range);
                        item.EndUpdate(sub);
                    }
                }
                stream.Close();
                stream.Dispose();
                return 0;
            });
            return taskCode;
            //_signedContent = string.Empty;
            //_signedImgs = new List<DevExpress.Office.Utils.OfficeImage>();
            //var doc = this._richEditCtrl.Document;
            //var rangePermissions = doc.BeginUpdateRangePermissions();
            //foreach (var rang in rangePermissions)
            //{
            //    //khong cho nguoi do sua noi dung cua chinh ho
            //    //if (rang.UserName != BOSApp.CurrentUsersInfo.ADUserName)
            //    this._signedContent += doc.GetText(rang.Range);
            //    foreach (var img in doc.Images)
            //    {
            //        if (img.Range.Start >= rang.Range.Start && img.Range.End <= rang.Range.End)
            //        {
            //            _signedImgs.Add(img.Image);
            //        }
            //    }
            //}
            //doc.EndUpdateRangePermissions(rangePermissions);

            //foreach (var item in _richEditCtrl.Document.Comments)
            //{
            //    SubDocument sub = item.BeginUpdate();
            //    this._signedContent += sub.GetText(sub.Range);
            //    item.EndUpdate(sub);
            //}
        }

        private bool IsViolateSignRule()
        {
            if (string.IsNullOrEmpty(this._signedContent)) return false;
            var doc = this._richEditCtrl.Document;
            var rangePermissions = doc.BeginUpdateRangePermissions();
            var newContent = string.Empty;
            foreach (var rang in rangePermissions)
            {
                //khong cho nguoi do sua noi dung cua chinh ho
                //if (rang.UserName != BOSApp.CurrentUsersInfo.ADUserName)
                newContent += doc.GetText(rang.Range);
            }

            //doc.EndUpdateRangePermissions(rangePermissions);

            foreach (var item in _richEditCtrl.Document.Comments)
            {
                SubDocument sub = item.BeginUpdate();
                newContent += sub.GetText(sub.Range);
                item.EndUpdate(sub);
            }
            if (!newContent.Equals(this._signedContent))
                return true;

            foreach (var img in _signedImgs)
            {
                if (doc.Images.Any(i => i.Image == img))
                    continue;
                else
                    return true;
            }
            return false;
        }
        #endregion

        #region Data and binding to Template
        /// <summary>
        /// uthv
        /// thuc thi action goi api, goi app, goi sql hay bat ky
        /// transactionAction: use for AppToApp HIS KV
        /// </summary>
        /// <param name="navigateUri"></param>
        internal object CallEmrAction(string navigateUri, DocumentRange actionRange, string groupIn = "", object data = null,
            Func<object, string, string, string> beforeBinding = null, object preData = null, int cacheTimeOut = 0, string meEmrTemplateActionDo = "", string transactionAction = "")
        {
            var errCount = 0;
            try
            {
                _msgNotification.Text = $"Đang cập nhật...";
                var emr = GetCurrentMainObject();
                Cursor.Current = Cursors.WaitCursor;
                if (IsEmrReadOnly())
                {
                    _msgNotification.Text = "Bệnh án này đã đóng. Mọi thay đổi sẽ không thể lưu được.";
                    return null;
                };
                var uris = navigateUri.Split(EmrParam.TagCodeSeparator);
                if (uris.Length == 0) return null;

                if (actionRange == null)
                    actionRange = this._emrDocumentHelper.GetActionRange(navigateUri, groupIn);

                var action = _actionsController.GetObjectByNo(uris[0]) as MEEmrActionsInfo;
                if (action == null)
                {
                    _msgNotification.Text = "Không tìm thấy thẻ chức năng tương ứng. Kiểm tra lại cấu hình";
                    return null;
                };

                var notify = $"Đang thực thi chức năng [{action.MEEmrActionName}]";
                _msgNotification.Text = notify;

                //case RemoteCase
                if (action.MEEmrActionType == EmrActionTypes.RemoteCase.ToString())
                {
                    InsertRemoteCaseData(action, actionRange, uris.Length > 1 ? uris[1] : null);
                    _msgNotification.Text = "Dữ liệu đã được cập nhật";
                    return null;
                }
                //get action gid
                var group = groupIn;
                if (string.IsNullOrEmpty(group))
                {
                    group = uris.Where(t => t.Contains($"{EmrParam.GuidTag}=")).FirstOrDefault();
                    group = string.IsNullOrEmpty(group) ? string.Empty : group.Replace($"{EmrParam.GuidTag}=", "");
                }

                if (action.MEEmrActionType == EmrActionTypes.Hard.ToString())
                {
                    this.PerformHardAction(action, actionRange, group);
                    _msgNotification.Text = _msgNotification.Text.Replace(notify, string.Empty);
                    return null;
                }
                else if (action.MEEmrActionType == EmrActionTypes.Composition.ToString())
                {
                    this.CallCompositionAction(action, actionRange, group, cacheTimeOut, meEmrTemplateActionDo);
                    return null;
                }
                else if (action.MEEmrActionType == EmrActionTypes.AutoValue.ToString())
                {
                    this.BindingDataToEmrDocument(GetHardParamList(emr), group, string.Empty);
                    return null;
                }

                // var key = this.GetDataKeyByGroup(group);
                var tid = uris.Where(t => t.Contains($"{EmrParam.TransactionIdTag}=")).FirstOrDefault();
                tid = tid == null ? string.Empty : tid.Replace($"{EmrParam.TransactionIdTag}=", "");

                var allParams = AppMemCache.GetActionParamsFromDict(action.MEEmrActionID);
                var updateParams = allParams.Where(o => o.MEEmrActionParamRequest == false && o.FK_MEParamID > 0).ToList();

                //sub template
                if (action.MEEmrActionType == EmrActionTypes.SubTemplate.ToString())
                {
                    DocumentPosition pos = actionRange.Start;
                    if (action.MEEmrActionNewGuid) group = Guid.NewGuid().ToString();
                    InsertSubTemplate(action, actionRange, updateParams, group);
                    _msgNotification.Text = "Dữ liệu đã được cập nhật";
                    return null;
                }
                if (action.MEEmrActionType == EmrActionTypes.DinamapProV100.ToString()
                    && action.MEEmrActionUri.ToUpper() == "PRINT")
                {
                    return PrintV100Snap();
                }

                var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;

                var emrNo = _entity.GetRelativeEmrNo(emr.MEEmrNo);
                var paramList = new Dictionary<string, object>
                {
                    // params cung, luc nao cung truyen xuong
                    { "patientNo", this._entity.MEPatient.MEPatientNo },
                    { "emrNo", emrNo },
                    { "documentNo", document.MEEmrDocumentFile },
                    { "documentDate", document.MEEmrDocumentCreatedDate },
                    //{ "documentTime", document.MEEmrDocumentCreatedDate },
                    { EmrParam.TransactionIdTag, tid },
                    { EmrParam.GuidTag, group },
                    { "caseNo", this.GetCaseNo() }
                };

                //hard action params
                var hardParams = _requestParamPool.GetActionParam(action.MEEmrActionSendParams);
                //doc action params
                foreach (var item in hardParams)
                {
                    if (!paramList.ContainsKey(item.Key))
                        paramList.Add(item.Key, item.Value);
                    else
                    {
                        //Cho phep lay gia tri tu to benh an cho bat ky param nao
                        if (action.MEEmrActionAllowOverrideReqParamValue)
                            paramList[item.Key] = item.Value;
                    }
                }
                //doc action params
                var requestParams = this._emrDocumentHelper.GetActionParamFromDoc(allParams.Where(o => o.MEEmrActionParamRequest).ToList(), action, actionRange, group);
                requestParams = _emrActionHelper.MapToRequestParams(action.MEEmrActionSendParams, requestParams);

                if (action.MEEmrActionType == EmrActionTypes.OpenWebBrowser.ToString())
                {
                    this.OpenWebBrowser(requestParams, action);
                    return null;
                }

                foreach (var item in requestParams)
                {
                    if (!paramList.ContainsKey(item.Key))
                        paramList.Add(item.Key, item.Value);
                    else
                    {
                        //Cho phep lay gia tri tu to benh an cho bat ky param nao
                        if (action.MEEmrActionAllowOverrideReqParamValue)
                            paramList[item.Key] = item.Value;
                    }
                }

                if (updateParams == null || updateParams.Count == 0)
                {
                    _msgNotification.Text = ("Chức năng này không định nghĩa cập nhật thông tin nào. Kiểm tra lại cấu hình chức năng này.");
                    return null;
                }

                if (action.MEEmrActionType == EmrActionTypes.Sql.ToString())
                {
                    paramList["documentDate"] = document.MEEmrDocumentCreatedDate.ToString("dd/MM/yyyy HH:mm:ss");
                    data = GetSqlData(cacheTimeOut, action, paramList);
                    if (data != null)
                    {
                        if (beforeBinding != null)
                            errCount = BindingDataToEmrDocument(data, action, group, updateParams, paramList, (d, path) =>
                            {
                                return beforeBinding(d, group, path);
                            }, preData);
                        else
                            errCount = BindingDataToEmrDocument(data, action, group, updateParams, paramList, preData: preData);
                    }
                }
                else if (action.MEEmrActionType == EmrActionTypes.Api.ToString())
                {
                    data = GetApiData(cacheTimeOut, action, paramList);
                    if (data == null)
                    {
                        _msgNotification.Text = ("Không có dữ liệu trả về từ api. Liên hệ quản trị viên để biết thêm chi tiết");
                        return null;
                    }
                    PrintMgsLogJson($"DU-LIEU-API {action.MEEmrActionName}", data);
                    PrintMgsLog("BAT-DAU-CAP-NHAT-DU-LIEU", action.MEEmrActionUri);
                    if (beforeBinding != null)
                        errCount = BindingDataToEmrDocument(data, action, group, updateParams, paramList, (d, path) =>
                        {
                            return beforeBinding(d, group, path);
                        }, preData);
                    else
                        errCount = BindingDataToEmrDocument(data, action, group, updateParams, paramList, preData: preData);
                    PrintMgsLog("KET-THUC-CAP-NHAT-DU-LIEU", action.MEEmrActionUri);
                }
                else if (action.MEEmrActionType == EmrActionTypes.DataPlugin.ToString()
                    || action.MEEmrActionType == EmrActionTypes.CalcPlugin.ToString())
                {
                    PrintMgsLog("BAT-DAU-GOI-PLUGIN", action.MEEmrActionPlugin);
                    data = CallDllPlugin(action, group, updateParams, paramList);
                    PrintMgsLog("KET-THUC-GOI-PLUGIN", action.MEEmrActionPlugin);
                    if (data == null)
                    {
                        _msgNotification.Text = ("Không có dữ liệu trả về từ Trình cắm. Liên hệ quản trị viên để biết thêm chi tiết");
                        return null;
                    }
                    PrintMgsLogJson($"DU-LIEU-PLUGIN {action.MEEmrActionName}", data);
                    PrintMgsLog("BAT-DAU-CAP-NHAT-DU-LIEU", action.MEEmrActionPlugin);
                    errCount = BindingDataToEmrDocument(data, action, group, updateParams, paramList, preData: preData);
                    PrintMgsLog("KET-THUC-CAP-NHAT-DU-LIEU", action.MEEmrActionPlugin);
                }
                #region App
                else if (action.MEEmrActionType == EmrActionTypes.App.ToString())
                {
                    try
                    {
                        var template = this._templateCtrl.GetObjectByID(document.FK_METemplateID) as METemplatesInfo;
                        string msg = "";
                        var message = new Dictionary<string, object>(paramList);
                        message["documentDate"] = document.MEEmrDocumentCreatedDate.ToString("dd/MM/yyyy HH:mm:ss");
                        message.Add("documentName", template.METemplateName);

                        message.Add("preData", preData);
                        _appPreData = preData;

                        message.Add("msg", action.MEEmrActionNo);
                        if (action.MEEmrActionDataType == EmrActionDataTypes.Xml.ToString())
                        {
                            msg = JsonConvert.SerializeObject(new { root = message }, Newtonsoft.Json.Formatting.Indented);
                            msg = JsonConvert.DeserializeXmlNode(msg).InnerXml;
                        }
                        else
                        {
                            msg = JsonConvert.SerializeObject(message, Newtonsoft.Json.Formatting.Indented);
                        }

                        if (!string.IsNullOrEmpty(_receiveChannel))
                        {
                            // HIS FPT, Other...
                            var icp = new IpcHelper();
                            icp.SendMessage(action.MEEmrActionNo, msg, action.MEEmrActionDataType);
                            PrintMgsLog("DA-GOI-YEU-CAU-DEN-HIS", msg);
                            _msgNotification.Text = "Đã gởi dữ liệu đến HIS. Đang chờ dữ liệu trả về...";
                            return null;
                        }
                        else if (_appToAppVB)
                        {
                            string thongbao = string.Empty;
                            var rs = AppToAppVB(message);
                            Type myType = rs.GetType();
                            //IList<PropertyInfo> props = new List<PropertyInfo>(myType.GetProperties());
                            //if (props != null && props.Count > 0)
                            if (rs is JObject jObject)
                            {
                                //foreach (PropertyInfo prop in props)
                                //{
                                //    if (prop.Name == "ThongBao")
                                //    {
                                //        string propValue = prop.GetValue(rs, null).ToString();
                                //        thongbao = !string.IsNullOrEmpty(propValue) ? propValue : "";
                                //    }
                                //    else if (prop.Name == "Data")
                                //    {
                                //        object propValue = prop.GetValue(rs, null);
                                //        data = propValue != null ? JArray.FromObject(propValue) : null;
                                //    }
                                //}
                                try
                                {
                                    // Lấy giá trị ThongBao
                                    string value = (string)jObject["ThongBao"];
                                    thongbao = !string.IsNullOrEmpty(value) ? value : "";

                                    // Lấy giá trị Data
                                    var dataToken = jObject["Data"];
                                    data = dataToken != null ? JArray.FromObject(dataToken) : null;
                                }
                                catch (Exception ex)
                                {
                                    // Xử lý lỗi nếu cần
                                }

                                // debug data from file txt
                                //string test = System.IO.File.ReadAllText(Path.Combine("D:/", "datatest.txt"));
                                //data = JsonConvert.DeserializeObject(test);
                                if (data != null)
                                {
                                    errCount = BindingDataToEmrDocument(data, action, group, updateParams, paramList);
                                    _msgNotification.Text = "Dữ liệu đã được cập nhật";
                                    PrintMgsLogJson($"DU-LIEU-APP-TO-APP {action.MEEmrActionName}", data);
                                    var responseString = JsonConvert.SerializeObject(data, Newtonsoft.Json.Formatting.Indented);
                                    // Store value in Mongo
                                    var logMongo = new Clas.Model.Base.LogMongo
                                    {
                                        Transaction = message["method"].ToString(),
                                        ValObj = responseString
                                    };
                                    InsertMongoLog(logMongo, "apptoapp_vb_result");
                                }
                                else
                                {
                                    _msgNotification.Text = "Không có dữ liệu trả về";
                                    PrintMgsLog($"DU-LIEU-APP-TO-APP {action.MEEmrActionName}", thongbao);
                                }
                            }
                            else
                                _msgNotification.Text = "Không có dữ liệu trả về";
                            return null;
                        }
                        else
                        {
                            // HIS KV
                            SignalRApp(action, msg, cacheTimeOut, group, updateParams, paramList, preData, transactionAction);
                            // Write mongo or public variable...

                            return null;
                        }
                    }
                    catch (Exception exApp)
                    {
                        Trace.TraceError("AppToApp ERROR: {0}:{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), exApp);
                        _sysHelper.LogTxt("error", exApp.ToString());
                        throw;
                    }
                }
                #endregion App
                else if (action.MEEmrActionType == EmrActionTypes.Replace.ToString())
                {
                    DocumentPosition pos = null;
                    //chuc nang nay yeu cau phat sinh ma nhom moi de co lap du lieu
                    if (action.MEEmrActionNewGuid) group = Guid.NewGuid().ToString();
                    foreach (var param in updateParams)
                    {
                        pos = this._emrActionHelper.ReplaceEmrParam(param, action, actionRange, group);
                    }
                    if (pos != null)
                        this._emrActionHelper.ReplaceEmrAction(pos, action, actionRange, group);
                }
                // gan gia tri cho the
                else if (action.MEEmrActionType == EmrActionTypes.Assign.ToString())
                {
                    foreach (var param in updateParams)
                    {
                        this._emrActionHelper.AssignEmrParamValue(param, action, actionRange, group);
                    }
                }
                // them dong du lieu moi cho table
                else if (action.MEEmrActionType == EmrActionTypes.AddRow.ToString())
                {
                    try
                    {
                        BOSProgressBar.Start("Đang xử lý dữ liệu");

                        var rowTDT = 0;
                        var maxRow = BOSApp.GetSystemConfigValueInt(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.EMR_DOCUMENT_TDT_DIENBIENVAYLENH_ROW, 0);
                        if (maxRow > 0)
                        {
                            var templateE = _templateCtrl.GetObjectByID(document.FK_METemplateID) as METemplatesInfo;
                            if (templateE.METemplateNo == "TDT")
                            {
                                rowTDT = GetTableEmrDocument("todieutri_dienbienvaylenh", 1, 1);
                                if (rowTDT >= maxRow)
                                {
                                    BOSProgressBar.Close();
                                    _msgNotification.Text = ("Số dòng Diễn biến, Y lệnh đã đủ. Không thể thêm dòng mới. Vui lòng tạo Tờ điều trị mới để thao tác.");
                                    UpdateEmrDocumentValidates(emr, document.FK_METemplateID, document.MEEmrDocumentID, "TRUE", $"Dòng trên tờ: {rowTDT} - Dòng quy định tối đa: {maxRow}");
                                    MessageBox.Show("Số dòng Diễn biến, Y lệnh đã đủ. Không thể thêm dòng mới. Vui lòng tạo Tờ điều trị mới để thao tác.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    return null;
                                }
                            }
                        }

                        var start = DateTime.Now;
                        var newGroup = this._emrActionHelper.AddTableRow(action, actionRange, group);
                        PrintMgsLog("AddTableRow: ", (DateTime.Now - start).TotalMilliseconds.ToString());

                        if (maxRow > 0)
                        {
                            rowTDT += 1;
                            if (rowTDT >= maxRow)
                            {
                                UpdateEmrDocumentValidates(emr, document.FK_METemplateID, document.MEEmrDocumentID, "TRUE", $"Dòng trên tờ: {rowTDT} - Dòng quy định tối đa: {maxRow}");
                            }
                            else
                            {
                                UpdateEmrDocumentValidates(emr, document.FK_METemplateID, document.MEEmrDocumentID, "FALSE", $"Dòng trên tờ: {rowTDT} - Dòng quy định tối đa: {maxRow}");
                            }
                        }

                        start = DateTime.Now;
                        this.BindingDataToEmrDocument(GetHardParamList(emr), newGroup, string.Empty);
                        PrintMgsLog("BindingDataToEmrDocument: ", (DateTime.Now - start).TotalMilliseconds.ToString());

                        if (beforeBinding != null)
                            data = beforeBinding(data as JToken, newGroup, string.Empty);
                        AutoAddHeaderAndFooter();
                    }
                    catch (Exception exAddRow)
                    {
                        _sysHelper.LogTxt("error", exAddRow.ToString());
                        throw;
                    }
                    finally
                    {
                        BOSProgressBar.Close();
                    }

                }
                // them cot du lieu moi cho table
                else if (action.MEEmrActionType == EmrActionTypes.AddCol.ToString())
                {
                    this._emrActionHelper.AddTableCol(action, actionRange, group);
                }
                #region Mongo
                else if (action.MEEmrActionType == EmrActionTypes.Lookup.ToString())
                {
                    if (_checkSystem)
                    {
                        _sysHelper.LogTxt("information", $"Action: {action.MEEmrActionType}");
                    }
                    var template = this._templateCtrl.GetObjectByID(action.FK_MESourceTemplateID) as METemplatesInfo;
                    var filters = new Dictionary<string, object>();
                    if (action.MEEmrActionScope != EmrActionScopes.All.ToString())
                    {
                        if (action.MEEmrActionScope == EmrActionScopes.Patient.ToString())
                            filters.Add("MEEmr.FK_MEPatientID", new Tuple<MongoFilter, object>(MongoFilter.Eq, this._entity.MEPatient.MEPatientID));
                        if (action.MEEmrActionScope == EmrActionScopes.Emr.ToString())
                            filters.Add("FK_MEEmrID", new Tuple<MongoFilter, object>(MongoFilter.Eq, emr.MEEmrID));
                        if (action.MEEmrActionScope == EmrActionScopes.Department.ToString())
                            filters.Add("MEEmr.FK_HRDepartmentID", new Tuple<MongoFilter, object>(MongoFilter.Eq, BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID));
                        //if (action.MEEmrActionScope == EmrActionScopes.Document.ToString() && !string.IsNullOrEmpty(document.MEEmrDocumentMongoID)) 
                        //    filters.Add("_id", new Tuple<MongoFilter, object>(MongoFilter.Eq, ObjectId.Parse(document.MEEmrDocumentMongoID)));
                        if (action.MEEmrActionScope == EmrActionScopes.Document.ToString())
                        {
                            if (string.IsNullOrEmpty(document.MEEmrDocumentMongoID))
                            {
                                filters.Add("_id", new Tuple<MongoFilter, object>(MongoFilter.Eq, ObjectId.GenerateNewId()));
                            }
                            else
                            {
                                filters.Add("_id", new Tuple<MongoFilter, object>(MongoFilter.Eq, ObjectId.Parse(document.MEEmrDocumentMongoID)));
                            }
                        }
                        if (action.MEEmrActionScope == EmrActionScopes.DocumentBeforeSaving.ToString())
                        {
                            filters.Add("RefId", new Tuple<MongoFilter, object>(MongoFilter.Eq, ObjectId.Parse(document.MEEmrDocumentMongoID)));
                        }
                    }
                    //cac scrope khac ko can xet trong to hien tai
                    if (action.MEEmrActionScope != EmrActionScopes.Document.ToString() && !string.IsNullOrEmpty(document.MEEmrDocumentMongoID))
                        filters.Add("_id", new Tuple<MongoFilter, object>(MongoFilter.Ne, ObjectId.Parse(document.MEEmrDocumentMongoID)));

                    filters.Add("MEEmrDocumentStatus", new Tuple<MongoFilter, object>(MongoFilter.Nin, new string[] { EmrDocumentStatus.Hidden.ToString(), EmrDocumentStatus.Discarded.ToString() }));

                    //hard action params
                    foreach (var item in hardParams)
                    {
                        var key = "MEEmrDocumentContent." + item.Key;
                        if (!filters.ContainsKey(key))
                            filters.Add(key, new Tuple<MongoFilter, object>(MongoFilter.Eq, item.Value));
                    }
                    //doc action params
                    foreach (var item in requestParams)
                    {
                        var key = "MEEmrDocumentContent." + item.Key;
                        if (!filters.ContainsKey(key))
                            filters.Add(key, new Tuple<MongoFilter, object>(MongoFilter.Eq, item.Value));
                    }
                    var fields = new Dictionary<string, string>();
                    foreach (var item in allParams.Where(o => o.FK_MEParamID == 0))
                    {
                        if (!string.IsNullOrEmpty(item.MEEmrActionParamSourcePath))
                        {
                            var sourcePath = item.MEEmrActionParamSourcePath.Replace("[*]", string.Empty);
                            var path = "$MEEmrDocumentContent." + sourcePath;
                            if (!fields.ContainsKey(sourcePath))
                                fields.Add(sourcePath, path);
                        }
                    }
                    foreach (var item in updateParams)
                    {
                        var dest = AppMemCache.GetParamFromDictKeyID(item.FK_MEParamID);
                        var sourcePath = item.MEEmrActionParamSourcePath.Replace("[*]", string.Empty);
                        if (dest != null && !string.IsNullOrEmpty(sourcePath))
                        {
                            var path = "$MEEmrDocumentContent." + sourcePath;
                            if (!fields.ContainsKey(dest.MEParamNo))
                                fields.Add(dest.MEParamNo, path);
                        }
                    }
                    if (_checkSystem)
                    {
                        _sysHelper.LogTxt("information", "Bắt đầu truy suất dữ liệu mongo...");
                    }
                    if (action.MEEmrActionScope == EmrActionScopes.DocumentBeforeSaving.ToString())
                    {
                        data = this._emrDocumentMng.FindLast(filters, fields, template.METemplateNo + "_revisions");
                    }
                    else
                    {
                        data = this._emrDocumentMng.Find(filters, fields, template.METemplateNo);
                    }

                    if (_checkSystem)
                    {
                        _sysHelper.LogTxt("information", $"Dữ liệu mongo: {data}");
                    }

                    if (data is JArray)
                        data = MergeChildArrayData(data as JArray);
                    else
                        data = MergeChildArrayData(data as JToken);
                    if (action.MEEmrActionAllowNullResult)
                    {
                        if (data == null)
                        {
                            if (string.IsNullOrEmpty(action.MEEmrActionPlugin))
                            {
                                _msgNotification.Text = ("Không tìm thấy dữ liệu phù hợp và không có cấu hình [Trình cắm]");
                                return data;
                            }
                            else
                            {
                                data = JToken.FromObject(new object());
                            }
                        }
                    }
                    else
                    {
                        if (data == null)
                        {
                            _msgNotification.Text = ("Không tìm thấy dữ liệu phù hợp");
                            return data;
                        }
                    }
                    //uthv chuyen thanh task de tang toc do
                    PrintMgsLogJson($"DU-LIEU-TRUY-VAN {action.MEEmrActionName}", data);
                    if (beforeBinding != null)
                        errCount = BindingDataToEmrDocument(data, action, group, updateParams, paramList, (d, path) =>
                        {
                            return beforeBinding(d, group, path);
                        }, preData, meEmrTemplateActionDo);
                    else
                        errCount = BindingDataToEmrDocument(data, action, group, updateParams, paramList, preData: preData, meEmrTemplateActionDo: meEmrTemplateActionDo);
                }
                #endregion
                else if (action.MEEmrActionType == EmrActionTypes.AutoAddDoc.ToString())
                {
                    paramList["documentDate"] = document.MEEmrDocumentCreatedDate.ToString("dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
                    data = this.AutoAddDocument(allParams, data, action, group, paramList);
                }
                else if (action.MEEmrActionType == EmrActionTypes.AddExtFileSharedDoc.ToString())
                {
                    this.AddExtFileSharedDocuments(paramList, allParams, data, action, group);
                }
                else if (action.MEEmrActionType == EmrActionTypes.AddImageFileShared.ToString())
                {
                    this.InsertImageFilesShared(paramList, allParams, data, action, group);
                }
                else if (action.MEEmrActionType == EmrActionTypes.DinamapProV100.ToString())
                {
                    errCount = this.InsertDinamapProV100DataAsync(paramList, allParams, data, action, group);
                    return null;
                }
                _msgNotification.Text = (errCount == 0) ? "Dữ liệu đã được cập nhật" : "Có lỗi khi cập nhật dữ liệu. Vui lòng xem thông báo lỗi";

                //chỉ chạy chức năng con trong trường hợp app-to-app các trường hợp khác dùng chức năng phức hợp thay thế
                //if (action.MEEmrActionType == EmrActionTypes.Sql.ToString()
                //    || action.MEEmrActionType == EmrActionTypes.Sql.ToString()
                //    || action.MEEmrActionType == EmrActionTypes.Api.ToString()
                //    || action.MEEmrActionType == EmrActionTypes.Assign.ToString()
                //    || action.MEEmrActionType == EmrActionTypes.Lookup.ToString())
                //{
                //    //TODO trả về danh sách
                //    data = CallChildEmrAction(action, actionRange, group, data);
                //}
            }
            catch (Exception exCall)
            {
                _sysHelper.LogTxt("error", exCall.ToString());
                MessageBox.Show(exCall.ToString(), "Có lỗi xảy ra", MessageBoxButtons.OK, MessageBoxIcon.Error);
                PrintMgsLog($"LOI_GOI_CHUC_NANG_{navigateUri}", exCall.ToString());
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }

            try
            {
                ParentScreen.Activate();
                _dpnRichEdit.Focus();
            }
            catch (Exception exP)
            {
                _sysHelper.LogTxt("error", exP.ToString());
            }

            return data;
        }

        private object GetSqlData(int cacheTimeOut, MEEmrActionsInfo action, Dictionary<string, object> paramList)
        {
            var hasCache = false;
            var key = string.Empty;
            object data = null;
            PrintMgsLog("BAT-DAU-DU-LIEU-SQL", action.MEEmrActionName + ": " + action.MEEmrActionUri);
            if (cacheTimeOut > 0)
            {
                key = GenCacheKey(action.MEEmrActionUri, paramList);
                var point = DateTime.Now.AddMilliseconds(-cacheTimeOut);
                hasCache = AppMemCache.GetActionDataFromMemCache(key, point, ref data);
                if (hasCache)
                {
                    PrintMgsLog("DU_LIEU_LAY_TU_CACHE", key);
                    PrintMgsLogJson($"DU-LIEU-SQL {action.MEEmrActionName}", data);
                }
            }
            if (!hasCache)
            {
                data = _sqlHelper.GetList(action.MEEmrActionUri, paramList);
                if (data != null)
                {
                    string str = JsonConvert.SerializeObject(data, Newtonsoft.Json.Formatting.Indented);
                    PrintMgsLog("DU-LIEU-SQL", str);
                    data = JsonConvert.DeserializeObject(str);
                    if (cacheTimeOut > 0)
                    {
                        AppMemCache.SetActionDataToMemCache(key, DateTime.Now, data);
                    }
                }
            }
            return data;
        }
        private object GetApiData(int cacheTimeOut, MEEmrActionsInfo action, Dictionary<string, object> paramList)
        {
            try
            {
                if (action != null && !string.IsNullOrEmpty(action.MEEmrActionUri))
                {
                    var hasCache = false;
                    var key = string.Empty;
                    object data = null;
                    if (cacheTimeOut > 0)
                    {
                        key = GenCacheKey(action.MEEmrActionUri, paramList);
                        var point = DateTime.Now.AddMilliseconds(-cacheTimeOut);
                        hasCache = AppMemCache.GetActionDataFromMemCache(key, point, ref data);
                        if (hasCache) PrintMgsLog("DU_LIEU_LAY_TU_CACHE", key);
                    }
                    if (!hasCache)
                    {
                        PrintMgsLog("BAT-DAU-GOI-API-HIS", action.MEEmrActionUri);
                        var start = DateTime.Now;
                        if (_msgNotification.InvokeRequired)
                            _msgNotification.BeginInvoke((MethodInvoker)delegate { _msgNotification.Text = ($"[Dữ liệu từ HIS] {action.MEEmrActionName}"); });
                        else
                            _msgNotification.Text = ($"[Dữ liệu từ HIS] {action.MEEmrActionName}");

                        if (_checkSystem)
                        {
                            var prString = string.Join(", ", paramList.Select(kv => $"{kv.Key}={kv.Value}"));
                            _sysHelper.LogTxt("information", $"Uri: {action.MEEmrActionUri} - Method: {action.MEEmrActionHttpMethod} - Param: {prString}");
                        }
                        if (action.MEEmrActionHttpMethod == HttpMethod.POST.ToString())
                            data = _api.Post(action.MEEmrActionUri, null, paramList);
                        else
                            data = _api.Get(action.MEEmrActionUri, paramList);
                        PrintMgsLog("KET-THUC-GOI-API-HIS", $"{action.MEEmrActionUri} [{(DateTime.Now - start).TotalMilliseconds} ms]");

                        if (cacheTimeOut > 0)
                        {
                            AppMemCache.SetActionDataToMemCache(key, DateTime.Now, data);
                        }
                    }
                    return data;
                }
                else
                {
                    PrintMgsLog("ERROR-API-HIS", $"{action.MEEmrActionNo}: Không tìm thấy thông tin đường dẫn API.");
                    if (_checkSystem)
                    {
                        _sysHelper.LogTxt("error", $"{action.MEEmrActionNo}: Không tìm thấy thông tin đường dẫn API.");
                    }
                    return null;
                }
            }
            catch (Exception exApi)
            {
                _sysHelper.LogTxt("error", exApi.Message);
                return null;
            }
        }

        private string GenCacheKey(string mEEmrActionUri, Dictionary<string, object> paramList)
        {
            var key = mEEmrActionUri;
            if (paramList.ContainsKey("patientNo"))
                key += $"/{paramList["patientNo"].ToString()}";
            if (paramList.ContainsKey("emrNo"))
                key += $"/{paramList["emrNo"].ToString()}";
            return key;
        }

        private BOSList<MEEmrDocumentsInfo> GetCurrentDocumentList(int emrId)
        {
            MEEmrsInfo profile = _entity.ModuleObjects[TableName.MEEmrsTableName] as MEEmrsInfo;
            if (emrId == profile.MEEmrID)
                return _entity.MEEmrPatientDocumentsList;
            else
                return _entity.MEEmrDocumentsList;
        }
        private MEEmrsInfo GetCurrentMainObject()
        {
            MEEmrsInfo profile = _entity.ModuleObjects[TableName.MEEmrsTableName] as MEEmrsInfo;
            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            if (document.FK_MEEmrID > 0 && document.FK_MEEmrID == profile.MEEmrID)
                return profile;
            else
                return _entity.MainObject as MEEmrsInfo;
        }
        private bool IsPatientProfileEditing()
        {
            MEEmrsInfo profile = _entity.ModuleObjects[TableName.MEEmrsTableName] as MEEmrsInfo;
            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            if (document.FK_MEEmrID == profile.MEEmrID)
                return true;
            else
                return false;
        }
        private JArray DecreaseArrLevel(JArray arr)
        {
            var data = new JArray();
            foreach (JToken v in arr)
            {
                if (v.Type == JTokenType.Array)
                {
                    var subArr = v as JArray;
                    if (subArr.Count() > 0 && subArr.Any(i => i is JArray))
                    {
                        foreach (var item in DecreaseArrLevel(subArr))
                        {
                            data.Add(item);
                        }
                    }
                    else
                    {
                        foreach (JToken s in subArr)
                        {
                            data.Add(s);
                        }
                    }
                }
                else
                {
                    data.Add(v);
                }
            }
            return data;
        }
        private JToken MergeChildArrayData(JToken data)
        {
            if (data == null) return null;
            foreach (JProperty pro in data)
            {
                if (pro.Value.Type == JTokenType.Array
                    && pro.Value.Count() > 0
                    && pro.Value.Any(i => i is JArray))
                {
                    pro.Value = DecreaseArrLevel(pro.Value as JArray);
                }
            }
            return data;
        }
        private JArray MergeChildArrayData(JArray data)
        {
            if (data == null) return null;
            var values = new JArray();
            var first = new JObject();
            foreach (JToken item in data)
            {
                var obj = new JObject();
                foreach (JProperty pro in item)
                {
                    if (pro.Value.Type == JTokenType.Array)
                    {
                        if (pro.Value.Count() > 0 && pro.Value.Any(i => i is JArray))
                        {
                            pro.Value = DecreaseArrLevel(pro.Value as JArray);
                        }
                        if (!first.ContainsKey(pro.Name))
                            first.Add(pro);
                        else
                        {
                            foreach (var v in pro.Value)
                                (first[pro.Name] as JArray).Add(v);
                        }
                    }
                    else
                    {
                        obj.Add(pro);
                    }
                }
                if (obj.Count > 1)
                    values.Add(obj);
            }
            if (first.Count > 0)
                values.Insert(0, first);
            return values;
        }

        /// <summary>
        /// Trả về mã chuyển khoa hiện tại
        /// Nếu không có mã chuyển khoa thì lấy mã emrNo
        /// </summary>
        /// <returns></returns>
        private string GetCaseNo()
        {
            var current = _entity.MEEmrTranfersList.Where(o => o.MEEmrTransferHistoryCurrent).FirstOrDefault();
            if (current != null)
                return current.MEEmrTransferHistoryNo;
            var emr = GetCurrentMainObject();
            return emr.MEEmrNo;
        }

        /// <summary>
        /// Tự động thêm tờ bệnh án theo data trả về
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private object AutoAddDocument(List<MEEmrActionParamsInfo> paramList, object data, MEEmrActionsInfo action, string group, Dictionary<string, object> requestParams)
        {
            if (paramList.Count == 0) return data;
            //chac chan la luu tai lieu truoc
            var command = this._richEditCtrl.CreateCommand(RichEditCommandId.FileSave);
            command.Execute();
            var values = this._dataHelper.FlattenObjDataWithoutChangeName(string.Empty, data as JObject);
            MEParamsInfo param = AppMemCache.GetParamFromDictKeyID(paramList[0].FK_MEParamID);
            var docNoParam = string.Empty;
            if (paramList.Count > 1)
                docNoParam = (AppMemCache.GetParamFromDictKeyID(paramList[1].FK_MEParamID))?.MEParamNo;

            var listDocNo = new List<string>();
            // ko lay data tu action truoc, lay data tu tai lieu
            if (data == null)
            {
                var dataFields = this._emrDocumentHelper.GetFieldsByGroup(group);
                var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
                data = JObject.FromObject(this._emrParser.ParserFieldsToDict(dataFields, AppMemCache.GetTemplateParams(document.FK_METemplateID)));
                values = this._dataHelper.FlattenObjDataWithoutChangeName(string.Empty, data as JObject);
            }
            var emr = GetCurrentMainObject();
            var documentList = GetCurrentDocumentList(emr.MEEmrID);
            MEEmrDocumentsInfo lastDocument = null;
            var listDocs = new List<string>();
            int success = 0, failed = 0, duplicate = 0;

            //chi truong hop goi SP tao to ngam moi can
            Dictionary<string, object> spParams = null;
            var createByJob = !string.IsNullOrEmpty(action.MEEmrActionUri);
            if (createByJob)
            {
                spParams = requestParams.Take(7).ToDictionary(x => x.Key, x => x.Value);
                spParams.Remove(param.MEParamNo);
                spParams.Remove(docNoParam);
                spParams.Add("departmentNo", BOSApp.CurrentDepartmentInfo.HRDepartmentNo);
                spParams.Add("createdUser", BOSApp.CurrentUser);
            }

            foreach (var prop in values)
            {
                if (prop.Key.EndsWith("." + param.MEParamNo))
                {
                    var reportNo = prop.Value.ToString();
                    var templates = this._templateCtrl.GetTemplateByParamAndValue(param.MEParamID, reportNo);
                    foreach (var template in templates)
                    {
                        string emrDocNo = string.Empty;
                        try
                        {
                            var value = (prop.Value as JToken).Parent;
                            if (value != null && value.Type == JTokenType.Property)
                            {
                                var rootPath = prop.Key.Substring(0, prop.Key.Length - param.MEParamNo.Length);
                                var path = rootPath + docNoParam;
                                var dict = value.Parent.ToObject<Dictionary<string, object>>();
                                var docNo = values.ContainsKey(path) ? values[path].ToString() : string.Empty;
                                emrDocNo = $"{template.METemplateNo}-{reportNo}-{docNo}";
                                if (documentList.Any(o => o.MEEmrDocumentCode.Equals(emrDocNo)))
                                {
                                    duplicate++;
                                    var msg = $"[ĐÃ CÓ] {emrDocNo} - {template.METemplateName}. Nên không được thêm lại.";
                                    PrintMgsLog("TU_DONG_THEM_TO_BENH_AN", msg);
                                    listDocs.Add(msg);
                                    _msgNotification.Text = msg;
                                    continue;
                                }
                                var list = value.Parent.Parent;
                                if (!string.IsNullOrEmpty(docNoParam) && !string.IsNullOrEmpty(docNo)
                                    && list != null && list.Type == JTokenType.Array)
                                {
                                    if (listDocNo.Contains(emrDocNo)) continue;
                                    var arr = new JArray();
                                    listDocNo.Add(emrDocNo);
                                    foreach (JObject item in list)
                                    {
                                        if (item.ContainsKey(docNoParam)
                                            && item[docNoParam].ToString() == docNo
                                            && item[param.MEParamNo].ToString() == reportNo)
                                            arr.Add(item);
                                    }
                                    var parentPath = (list.Parent as JProperty)?.Name;
                                    if (!string.IsNullOrEmpty(parentPath))
                                        dict.Add(parentPath, arr);

                                }
                                if (!createByJob)
                                {
                                    _entity.MENewEmrDocument = new MEEmrDocumentsInfo()
                                    {
                                        FK_METemplateID = template.METemplateID,
                                        MEEmrDocumentFile = template.METemplateNo + DateTime.Now.ToString("_ddMMyyyy_HHmmssffff_") + Guid.NewGuid().ToString().Replace('-', '_'),
                                        MEEmrDocumentNo = template.METemplateNo,
                                        MEEmrDocumentCode = emrDocNo,
                                        MEEmrDocumentGuid = template.METemplateGuid,
                                        MEEmrDocumentDesc = "Thêm tự động",
                                        FK_HREmployeeCreatedID = BOSApp.CurrentEmployeesInfo.HREmployeeID
                                    };
                                    var start = DateTime.Now;
                                    if (IsPatientProfileEditing())
                                        //TODO toi uu toc do nhu AddNewEmrDocumentAuto
                                        lastDocument = AddNewPatientDocument(dict);
                                    else
                                        lastDocument = AddNewEmrDocumentAuto(dict);

                                    listDocs.Add($"[ĐÃ TẠO] {emrDocNo} - {template.METemplateName}");
                                    PrintMgsLog("THEM_TO_TU_DONG", (DateTime.Now - start).TotalMilliseconds.ToString());
                                    success++;
                                }
                                else
                                {
                                    var start = DateTime.Now;
                                    var prs = new Dictionary<string, object>(spParams)
                                    {
                                        { "templateNo", template.METemplateNo },
                                        { param.MEParamNo, reportNo},
                                        { docNoParam, docNo},
                                    };
                                    prs = GetRequestParams(paramList.Skip(2).Take(int.MaxValue).ToArray(), prs, data as JToken);
                                    var sqlResult = GetSqlData(0, action, prs) as JToken;
                                    var message = "TẠO NGẦM";
                                    if (sqlResult != null)
                                        message = sqlResult.SelectToken("$." + UserFriendlyException.MESSAGE_KEY.ToLower(), false)?.ToString();
                                    listDocs.Add($"[{message}] {emrDocNo} - {template.METemplateName}");
                                    PrintMgsLog("THEM_TO_TU_DONG_NGAM", (DateTime.Now - start).TotalMilliseconds.ToString());
                                    success++;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            failed++;
                            listDocs.Add($"[LỖI] {emrDocNo} - {template.METemplateName}. Chi tiết lỗi: {ex.Message}");
                            PrintMgsLog("LOI_THEM_TO_TU_DONG", ex.ToString());
                        }
                    }
                }
            }
            if (lastDocument != null && lastDocument.MEEmrDocumentID > 0)
            {
                this.TakeEditingPermission(lastDocument);
                this.InvalidateEmrDocumentList(emr.MEEmrID);
                var gridView = _entity.MEEmrDocumentsList.GridView;
                gridView.FocusedRowHandle = gridView.LocateByValue("MEEmrDocumentID", lastDocument.MEEmrDocumentID);
                lastDocument = SetCurrentEditingUser(lastDocument);
                InvalidateDocument(lastDocument, false);
            }
            else
            {
                if (success > 0 || failed > 0)
                {
                    this.InvalidateEmrDocumentList(emr.MEEmrID);
                    this.ClearDocumentSession();
                }
            }

            if (listDocs.Count > 0)
                MessageBox.Show($"[{success}] TỜ ĐƯỢC TẠO {(createByJob ? "NGẦM" : string.Empty)}, [{failed}] LỖI, [{duplicate}] ĐÃ CÓ.\n\n" +
                    $"- {string.Join("\n- ", listDocs)}\n\n" +
                    (failed > 0 ? "CHẠY LẠI CHỨC NĂNG ĐỂ THỬ TẠO LẠI CÁC TỜ BỊ LỖI\n" : "") +
                    $"Ctrl+C để copy nội dung thông báo này.",
                    $"THỐNG KÊ {(createByJob ? "TÁC VỤ NGẦM, VUI LÒNG CHỜ ÍT PHÚT ĐỂ TẠO" : string.Empty)}", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return null;
        }
        private Dictionary<string, object> GetRequestParams(MEEmrActionParamsInfo[] paramList, Dictionary<string, object> requestParams, JToken data)
        {
            foreach (var param in paramList)
            {
                var value = data.SelectToken("$." + param.MEEmrActionParamSourcePath);
                var no = (AppMemCache.GetParamFromDictKeyID(param.FK_MEParamID))?.MEParamNo;
                if (!string.IsNullOrEmpty(no))
                {
                    requestParams.Remove(no);
                    requestParams.Add(no, value);
                }
            }
            return requestParams;
        }
        private object AddExtFileSharedDocuments(Dictionary<string, object> sendParams, List<MEEmrActionParamsInfo> paramList, object data, MEEmrActionsInfo action, string group)
        {
            if (paramList.Count == 0) return data;
            //chac chan la luu tai lieu truoc
            var command = this._richEditCtrl.CreateCommand(RichEditCommandId.FileSave);
            command.Execute();
            var values = this._dataHelper.FlattenObjDataWithoutChangeName(string.Empty, data as JObject);
            // ko lay data tu action truoc, lay data tu tai lieu
            if (data == null)
            {
                var dataFields = this._emrDocumentHelper.GetFieldsByGroup(group);
                var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
                data = JObject.FromObject(this._emrParser.ParserFieldsToDict(dataFields, AppMemCache.GetTemplateParams(document.FK_METemplateID)));
                values = this._dataHelper.FlattenObjDataWithoutChangeName(string.Empty, data as JObject);
            }
            var filePathProperty = "path";
            if (paramList.Any(p => !p.MEEmrActionParamRequest))
            {
                var file = AppMemCache.GetParamFromDictKeyID(paramList.Where(p => !p.MEEmrActionParamRequest).First().FK_MEParamID);
                filePathProperty = file.MEParamNo;
            }
            var reqParams = paramList.Where(p => p.MEEmrActionParamRequest).ToList();
            MEParamsInfo param = AppMemCache.GetParamFromDictKeyID(reqParams[0].FK_MEParamID);
            var docNoParam = string.Empty;
            if (paramList.Count > 1)
                docNoParam = (AppMemCache.GetParamFromDictKeyID(reqParams[1].FK_MEParamID))?.MEParamNo;

            var listDocNo = new List<string>();
            foreach (var prop in values)
            {
                if (prop.Key == param.MEParamNo || prop.Key.EndsWith("." + param.MEParamNo))
                {
                    var reportNo = prop.Value.ToString();
                    if (string.IsNullOrEmpty(reportNo)) continue;
                    var value = (prop.Value as JToken).Parent;
                    if (value != null && value.Type == JTokenType.Property)
                    {
                        var path = prop.Key.Substring(0, prop.Key.Length - param.MEParamNo.Length) + docNoParam;
                        var docNo = values.ContainsKey(path) ? values[path].ToString() : string.Empty;
                        if (!string.IsNullOrEmpty(docNo))
                        {
                            if (listDocNo.Contains(reportNo + docNo)) continue;
                            listDocNo.Add(reportNo + docNo);
                            sendParams[param.MEParamNo] = reportNo;
                            sendParams[docNoParam] = docNo;
                            PrintMgsLog("BAT-DAU-GOI-API", action.MEEmrActionUri);

                            if (action.MEEmrActionHttpMethod == HttpMethod.POST.ToString())
                                data = _api.Post(action.MEEmrActionUri, null, sendParams);
                            else
                                data = _api.Get(action.MEEmrActionUri, sendParams);

                            PrintMgsLog("KET-THUC-GOI-API", action.MEEmrActionUri);
                            if (data != null)
                            {
                                AddExtFileSharedDocument(param, reportNo, filePathProperty, docNo, data, action, group);
                            }
                        }

                    }

                }
            }
            return data;
        }

        /// <summary>
        /// Thêm tờ bệnh án từ hệ thống khác thông qua API
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private object AddExtFileSharedDocument(MEParamsInfo param, string reportNo, string filePathProperty, string docNo, object data, MEEmrActionsInfo action, string group)
        {
            var values = this._dataHelper.FlattenObjDataWithoutChangeName(string.Empty, data as JObject);
            METemplatesInfo template = null;
            template = this._templateCtrl.GetTemplateByParamAndValue(param.MEParamID, reportNo).FirstOrDefault();
            if (template == null)
            {
                PrintMgsLog("THEM_TO_BENH_AN_FILE_SHARE", "Không tìm thấy mẫu bệnh án để phát sinh tự động. Kiểm tra lại cấu hình");
                _msgNotification.Text = $"Không tìm thấy mẫu bệnh án để phát sinh tự động. Kiểm tra lại cấu hình";
                return data;
            }
            var emr = GetCurrentMainObject();
            var documentList = GetCurrentDocumentList(emr.MEEmrID);

            var emrDocNo = $"{template.METemplateNo}-{reportNo}-{docNo}";
            if (documentList.Any(o => o.MEEmrDocumentCode.Equals(emrDocNo)))
            {
                PrintMgsLog("THEM_TO_BENH_AN_FILE_SHARE", $"Đã tồn tại tờ: {emrDocNo}. Tờ này sẽ không được thêm lại.");
                _msgNotification.Text = $"Xem thông báo lỗi";
                return data;
            }
            var error = 0;
            foreach (var prop in values)
            {
                if (prop.Key.Contains(filePathProperty))
                {
                    var path = prop.Value.ToString();
                    if (File.Exists(path))
                    {
                        try
                        {
                            _entity.MENewEmrDocument = new MEEmrDocumentsInfo()
                            {
                                FK_METemplateID = template.METemplateID,
                                MEEmrDocumentFile = template.METemplateNo + DateTime.Now.ToString("_ddMMyyyy_HHmmssffff_") + Guid.NewGuid().ToString().Replace('-', '_'),
                                MEEmrDocumentNo = template.METemplateNo,
                                MEEmrDocumentCode = emrDocNo,
                                MEEmrDocumentGuid = template.METemplateGuid,
                                MEEmrDocumentDesc = "Tờ bệnh án được lấy từ hệ thống ngoài",
                                MEEmrDocumentExternalFile = path,
                                FK_HREmployeeCreatedID = BOSApp.CurrentEmployeesInfo.HREmployeeID
                            };
                            if (IsPatientProfileEditing())
                                AddNewPatientDocument(null);
                            else
                                AddNewEmrDocument(null);
                        }
                        catch (Exception ex)
                        {
                            error++;
                            PrintMgsLog($"LOI_THEM_TO_BENH_AN_FILE_SHARE {_entity.MENewEmrDocument.MEEmrDocumentExternalFile}", ex.ToString());
                        }
                    }
                }
            }

            if (error > 0)
            {
                if (IsPatientProfileEditing())
                {
                    this.InvalidatePatientDocuments(emr.MEEmrID);
                }
                else
                {
                    this.InvalidateEmrDocumentList(emr.MEEmrID);
                }
                this.ClearDocumentSession();
                MessageBox.Show("Có lỗi khi thêm tờ bệnh án. Xin thử lại. Chi tiết ở màn hình [Thông báo]", "CÓ LỖI KHI TẠO TỜ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return data;
        }

        private object CallChildEmrAction(MEEmrActionsInfo action, DocumentRange actionRange, string group, object parentData)
        {
            var childActions = this._actionRelationsController.GetAllByParentID(action.MEEmrActionID).OrderBy(c => c.MEEmrActionRelationOrder).ToList();
            object data = null;
            var serialized = JsonConvert.SerializeObject(parentData);
            if (childActions.Count > 0)
            {
                for (int i = 0; i < childActions.Count; i++)
                {
                    var childAction = this._actionsController.GetObjectByID(childActions[i].FK_MEEmrActionChildID) as MEEmrActionsInfo;
                    if (childAction == null) continue;
                    data = CallEmrAction(childAction.MEEmrActionNo, actionRange, group, JsonConvert.DeserializeObject(serialized), preData: data);
                }
            }
            return data;
        }
        private string BeforeBindingData(List<MEEmrActionRelationsInfo> childActs, string group, object data, string path)
        {
            foreach (var c in childActs)
            {
                var child = this._actionsController.GetObjectByID(c.FK_MEEmrActionChildID) as MEEmrActionsInfo;
                if (child == null) continue;
                var query = !string.IsNullOrEmpty(c.MEEmrActionRelationDependOnData) ? "$." + c.MEEmrActionRelationDependOnData : null;
                if (child.MEEmrActionType == EmrActionTypes.AddRow.ToString())
                {
                    var parram = AppMemCache.GetParamFromDictKeyNo(c.MEEmrActionRelationDependOnData);
                    //so luong row can them la so luong dong du lieu cua data
                    if (path == c.MEEmrActionRelationDependOnData && data != null)
                    {
                        var list = JArray.FromObject(data).ToArray();
                        var rowCount = list.Length;
                        if (parram != null)
                            rowCount = list.Length - parram.MEParamSampleCount;

                        var exitedRow = _emrDocumentHelper.GetFirstEmrFieldByPrefix($"{c.MEEmrActionRelationDependOnData}{EmrParam.CodeSeparator}1");
                        if (exitedRow != null)
                            rowCount = list.Length;
                        else
                        {
                            var table = _emrDocumentHelper.GetFirstEmrFieldByCode(c.MEEmrActionRelationDependOnData);
                            //Neu da duoc update truoc do
                            if (table != null && table.Tokens.Where(t => t.StartsWith($"{EmrParam.UserUpdate}=")).FirstOrDefault() != null)
                            {
                                rowCount = list.Length;
                            }
                        }

                        for (int i = 0; i < rowCount; i++)
                        {
                            var t = list[i];
                            if (_checkSystem)
                            {
                                _sysHelper.LogTxt("information", $"Bắt đầu chạy chức năng {child.MEEmrActionNo}.");
                                var watchAct = Stopwatch.StartNew();
                                CallEmrAction(child.MEEmrActionNo, null, group, t,
                                   (d, g, p) =>
                                   {
                                       return BeforeBindingData(c.Children, g, d, p);
                                   });
                                watchAct.Stop();
                                var elapsedAct = watchAct.ElapsedMilliseconds / 1000.0;
                                _sysHelper.LogTxt("information", $"{elapsedAct} giây. Hoàn tất chạy chức năng {child.MEEmrActionNo}.");
                            }
                            else
                            {
                                CallEmrAction(child.MEEmrActionNo, null, group, t,
                                   (d, g, p) =>
                                   {
                                       return BeforeBindingData(c.Children, g, d, p);
                                   });
                            }
                        }
                    }
                    else //thuc thi 1 lan
                    {
                        // Do later
                    }
                }
                else if (child.MEEmrActionType == EmrActionTypes.SubTemplate.ToString())
                {

                }
                else if (child.MEEmrActionType == EmrActionTypes.Replace.ToString())
                {
                    if (query == null) continue;
                    var value = (data as JToken).SelectToken(query);
                    if (value.HasValues)
                    {
                        if (_checkSystem)
                        {
                            _sysHelper.LogTxt("information", $"Bắt đầu chạy chức năng {child.MEEmrActionNo}.");
                            var watchAct = Stopwatch.StartNew();
                            CallEmrAction(child.MEEmrActionNo, null, group, data as JToken,
                              (d, g, p) =>
                              {
                                  return BeforeBindingData(c.Children, g, d, p);
                              });
                            watchAct.Stop();
                            var elapsedAct = watchAct.ElapsedMilliseconds / 1000.0;
                            _sysHelper.LogTxt("information", $"{elapsedAct} giây. Hoàn tất chạy chức năng {child.MEEmrActionNo}.");
                        }
                        else
                        {
                            CallEmrAction(child.MEEmrActionNo, null, group, data as JToken,
                              (d, g, p) =>
                              {
                                  return BeforeBindingData(c.Children, g, d, p);
                              });
                        }
                    }
                }
            }
            return group;
        }
        private void CallCompositionAction(MEEmrActionsInfo action, DocumentRange actionRange, string group, int cacheTimeOut, string meEmrTemplateActionDo = "")
        {
            var acts = this._actionRelationsController.GetAllByParentID(action.MEEmrActionID).OrderBy(o => o.MEEmrActionRelationOrder).ToList();
            acts = this._actionRelationsController.GetTree(acts, 0);
            object preData = null;
            object data = null;
            var modeApp = false;
            var transactionAction = string.Empty;
            foreach (var a in acts)
            {
                // không truyền data vi các action nay ngang cấp, chỉ các action kế thừa mới pass data cho nhau
                var act = this._actionsController.GetObjectByID(a.FK_MEEmrActionChildID) as MEEmrActionsInfo;
                if (act == null) continue;
                if (act.MEEmrActionType == EmrActionTypes.App.ToString())
                {
                    transactionAction = $"{act.MEEmrActionID}_{DateTime.Now.ToString("yyyyMMddHHmmss")}";
                }
                if (_stateAppHISKV != 0 && modeApp)
                {
                    // _stateAppHISKV
                    // 0: none, error
                    // 1: init
                    // 2: received data
                    // Move to task: error do later
                    var runtime = BOSApp.GetSystemConfigValueInt(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.HIS_KV_RUN_TIME, 90000);
                    try
                    {
                        preData = CallNextActionComposition(actionRange, group, cacheTimeOut, meEmrTemplateActionDo, preData, data, transactionAction, a, act, runtime);
                    }
                    catch (Exception exTask)
                    {
                        _sysHelper.LogTxt("error", exTask.Message);
                    }
                    //Task.Run(() =>
                    //{
                    //    try
                    //    {
                    //        preData = CallNextActionComposition(actionRange, group, cacheTimeOut, meEmrTemplateActionDo, preData, data, transactionAction, a, act, runtime);
                    //    }
                    //    catch (Exception exTask) 
                    //    {
                    //        _sysHelper.LogTxt("error", exTask.Message);
                    //    }
                    //});
                }
                else
                {
                    preData = CallCompositionChildAction(actionRange, group, cacheTimeOut, meEmrTemplateActionDo, preData, data, a, act, transactionAction);
                }
                // XuanTM
                modeApp = act.MEEmrActionType == EmrActionTypes.App.ToString() ? true : false;
            }
        }

        private object CallNextActionComposition(DocumentRange actionRange, string group, int cacheTimeOut, string meEmrTemplateActionDo, object preData, object data, string transactionAction, MEEmrActionRelationsInfo a, MEEmrActionsInfo act, int runtime)
        {
            var iTime = 0;
            var quickRun = false;
            while (!quickRun)
            {
                if (_checkSystem)
                {
                    _sysHelper.LogTxt("information", $"Run time: {iTime}");
                }
                if (_stateAppHISKV == 0 || iTime == runtime)
                {
                    break;
                }
                if (_stateAppHISKV == 2)
                {
                    _stateAppHISKV = 0;
                    preData = CallCompositionChildAction(actionRange, group, cacheTimeOut, meEmrTemplateActionDo, preData, data, a, act, transactionAction);
                    break;
                    // No need
                    //quickRun = GetLog(transactionAction);
                    //if (quickRun)
                    //{
                    //    _stateAppHISKV = 0;
                    //    preData = CallCompositionChildAction(actionRange, group, cacheTimeOut, meEmrTemplateActionDo, preData, data, a, act, transactionAction);
                    //    break;
                    //}
                }
                iTime++;
            }

            return preData;
        }

        public bool GetLog(string transactionAction)
        {
            var filters = new Dictionary<string, object>
                        {
                            { "Transaction", new Tuple<MongoFilter, object>(MongoFilter.Eq, transactionAction) }
                        };
            var fields = new Dictionary<string, string>();
            var logs = FindMongoLog(filters, fields, "apptoapp_kv");
            if (logs != null && logs.Count() > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private object CallCompositionChildAction(DocumentRange actionRange, string group, int cacheTimeOut, string meEmrTemplateActionDo, object preData, object data, MEEmrActionRelationsInfo a, MEEmrActionsInfo act, string transactionAction)
        {
            if (a.Children.Count > 0)
            {
                preData = CallEmrAction(act.MEEmrActionNo, actionRange, group, data,
                   (d, g, path) =>
                   {
                       return BeforeBindingData(a.Children, g, d, path);
                   }, preData, cacheTimeOut: cacheTimeOut, meEmrTemplateActionDo, transactionAction);
            }
            else
            {
                preData = CallEmrAction(act.MEEmrActionNo, actionRange, group, data, null, preData, cacheTimeOut: cacheTimeOut, meEmrTemplateActionDo, transactionAction);
            }

            return preData;
        }

        /// <summary> 
        ///Thực thi các thẻ chức năng cứng
        /// </summary>
        /// <param name="action"></param>
        /// <param name="actionRange"></param>
        /// <param name="group"></param>
        private void PerformHardAction(MEEmrActionsInfo action, DocumentRange actionRange, string group)
        {
            var methodE = GetType().GetMethods().Where(m => action.MEEmrActionNo.ToUpper().Equals(m.Name.ToUpper())).FirstOrDefault();
            if (methodE != null)
            {
                methodE.Invoke(this, null);
            }
            else
            {
                switch (action.MEEmrActionNo)
                {
                    case "CLOSE-EMR":
                        this.CloseAllEmr();
                        break;
                    case "SaveDocument":
                        this.SaveDocument();
                        break;
                    case "validate-emr-document":
                        var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
                        if (document.MEEmrDocumentID > 0)
                        {
                            this.ValidateEmrDocument(document, true);
                        }
                        else
                        {
                            var emr = _entity.MainObject as MEEmrsInfo;
                            var documents = _emrDocumentCtrl.GetByEmrId(emr.MEEmrID);
                            foreach (var item in documents)
                            {
                                this.ValidateEmrDocument(item, false);
                            }
                        }
                        break;
                    case "CLOSE-DOC":
                        break;
                    default:
                        break;
                }
            }
        }

        public void SaveDocument()
        {
            var command = this._richEditCtrl.CreateCommand(RichEditCommandId.FileSave);
            command.Execute();
        }

        public void ValidateEmrDocument(MEEmrDocumentsInfo document, bool current)
        {
            var result = GetTemplateParamsValidateByDocument(document, current);
            if (result.Count() > 0)
            {
                var gui = new guiMETemplateParams("CẢNH BÁO DANH SÁCH THẺ [CHƯA NHẬP DỮ LIỆU]", result)
                {
                    Module = this,
                    StartPosition = FormStartPosition.CenterParent
                };
                BOSProgressBar.Close();
                gui.ShowDialog();
            }
        }

        private List<METemplateParamsInfo> GetTemplateParamsValidateByDocument(MEEmrDocumentsInfo document, bool current)
        {
            var result = new List<METemplateParamsInfo>();

            var requiredTemplateParams = AppMemCache.GetTemplateParams(document.FK_METemplateID).Where(m => m.METemplateParamRequired.Equals(true)).ToList();
            if (requiredTemplateParams != null && requiredTemplateParams.Count > 0)
            {
                object data = null;
                if (current)
                {
                    data = JsonConvert.DeserializeObject(document.MEEmrDocumentJson) as JToken;
                    if (data == null)
                    {
                        current = false;
                    }
                }

                if (!current)
                {
                    var filters = new Dictionary<string, object>
                        {
                            { "MEEmrDocumentID", new Tuple<MongoFilter, object>(MongoFilter.Eq, document.MEEmrDocumentID) }
                        };
                    var fields = new Dictionary<string, string>();
                    foreach (var item in requiredTemplateParams)
                    {
                        var mETemplateParamPath = item.METemplateParamPath;
                        if (mETemplateParamPath.Contains('*'))
                        {
                            mETemplateParamPath = mETemplateParamPath.Split('.')[0].Replace("[*]", "");
                        }
                        var sourcePath = mETemplateParamPath;
                        var path = "$MEEmrDocumentContent." + mETemplateParamPath;
                        if (!fields.ContainsKey(sourcePath))
                        {
                            fields.Add(sourcePath, path);
                        }
                    }

                    data = _emrDocumentMng.Find(filters, fields, document.MEEmrDocumentNo);
                    if (data is JArray)
                        data = MergeChildArrayData(data as JArray);
                    else
                        data = MergeChildArrayData(data as JToken);
                }

                if (data != null)
                {
                    if (((JToken)data).Type != JTokenType.Array)
                    {
                        var dataObj = (data as JObject).ToObject<IDictionary<string, object>>();
                        foreach (var require in requiredTemplateParams)
                        {
                            if (!ValidateObject(dataObj, require.METemplateParamPath, require.FK_MEParamID, new List<string>()))
                            {
                                result.Add(require);
                            }
                        }
                    }
                }
            }

            return result;
        }

        private bool ValidateObject(IDictionary<string, object> data, string key, int id, List<string> datakeys)
        {
            var keys = key.Split('.');
            var keyRequire = keys[0].Replace("[*]", "");
            if (data.TryGetValue(keyRequire, out object value))
            {
                if (value.GetType() == typeof(JObject))
                {
                    var childDataObj = (value as JObject).ToObject<IDictionary<string, object>>();
                    if (childDataObj.Count() == 0)
                    {
                        return false;
                    }
                    var param = AppMemCache.GetParamFromDictKeyID(id);
                    if (param != null && (param.MEParamControlType == EmrParamControlTypes.Radio.ToString() || param.MEParamControlType == EmrParamControlTypes.Checkbox.ToString()))
                    {
                        var result = false;
                        foreach (var child in childDataObj)
                        {
                            if (ValidateObject(childDataObj, child.Key, id, datakeys))
                            {
                                result = true;
                                break;
                            }
                        }
                        return result;
                        //var childs = AppMemCache.GetParamRelationsFromDict(param.MEParamID);
                        //var result = false;
                        //foreach (var child in childs)
                        //{
                        //    var paramChild = AppMemCache.GetParamFromDictKeyID(child.FK_MEParamChildID);
                        //    if (ValidateObject(childDataObj, paramChild.MEParamNo, id))
                        //    {
                        //        result = true;
                        //        break;
                        //    }
                        //}
                        //return result;
                    }
                    else if (param != null && keys.Count() > 1)
                    {
                        return ValidateObject(childDataObj, key.Replace(keys[0] + ".", ""), id, datakeys);
                    }

                    var keyChildRequire = key.Split('.').Last();
                    return ValidateObject(childDataObj, keyChildRequire, id, childDataObj.Select(item => (string)item.Key).ToList());
                }
                else if (value.GetType() == typeof(JArray))
                {
                    if (JArray.FromObject(value).Count() == 0)
                    {
                        return false;
                    }
                    if (keys.Count() > 1 && !string.IsNullOrEmpty(key.Replace(keys[0], "")))
                    {
                        var result = true;
                        var values = JArray.FromObject(value);
                        foreach (var itemVal in values)
                        {
                            var dict = itemVal.Cast<JProperty>()
                                .ToDictionary(item => item.Name,
                                            item => (object)item.Value); //can be modified for your desired type

                            if (!ValidateObject(dict, key.Replace(keys[0] + ".", ""), id, datakeys))
                            {
                                result = false;
                                break;
                            }
                        }
                        return result;
                        //var dict = JArray.FromObject(value).First() //First() is only necessary if embedded in a json object
                        //        .Cast<JProperty>()
                        //        .ToDictionary(item => item.Name,
                        //                    item => (object)item.Value); //can be modified for your desired type

                        //return ValidateObject(dict, key.Replace(keys[0] + ".", ""), id);
                    }
                }
                else
                {
                    if (string.IsNullOrEmpty(value.ToString()))
                    {
                        return false;
                    }
                }
            }

            if (value == null && datakeys.Count() > 0)
            {
                // Check child have data is ok
                var result = false;
                foreach (var datakey in datakeys)
                {
                    if (ValidateObject(data, datakey, id, new List<string>()))
                    {
                        result = true;
                        break;
                    }
                }
                return result;
            }
            return true;
        }

        /// <summary>
        /// The delegate which processes all cross AppDomain messages and writes them to screen.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listener_MessageFromHisReceived(object sender, IpcMessageEventArgs e)
        {
            BOSApp.MainScreen.Activate();
            PrintMgsLog("THONG_DIEP_TU_HIS", e.DataGram.Message);
            int errCount = 0;
            if (e.DataGram.Channel == _receiveChannel)
            {
                var response = new Dictionary<string, object>();
                var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
                try
                {
                    if (e.DataGram.DataType == EmrActionDataTypes.Xml.ToString())
                    {
                        //very hard
                        var docXml = new XmlDocument();
                        docXml.LoadXml(e.DataGram.Message);
                        var json = JsonConvert.SerializeXmlNode(docXml);
                        response = (JsonConvert.DeserializeObject<Dictionary<string, object>>(json)["root"] as JObject).ToObject<Dictionary<string, object>>();
                    }
                    else
                    {
                        response = JsonConvert.DeserializeObject<Dictionary<string, object>>(e.DataGram.Message);
                    }
                }
                catch (Exception ex)
                {
                    Trace.TraceError("IPC ERROR: {0}:{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), ex);
                    ShowFlashNotification("Có lỗi xảy ra", 3000);
                    return;
                    //throw;
                }
                if (!response.ContainsKey("msg")) return;
                switch (response["msg"].ToString())
                {
                    case "CREATE_EMR":
                        CreateEmrFromHis(response);
                        return;
                    case "CREATE_DOC":
                        CreateDocFromHis(response);
                        return;
                    default:
                        break;
                }
                string patientNo = response.ContainsKey("patientNo") ? response["patientNo"].ToString() : response["patientno"].ToString();
                if (patientNo != this._entity.MEPatient.MEPatientNo)
                {
                    _msgNotification.Text = ("Ứng dụng nhận được dữ liệu nhưng của bệnh nhân có mã: " + response["patientNo"].ToString() + " Vui lòng chọn đúng bệnh nhân, bệnh án và mẫu bệnh án");
                    return;
                }

                //truong hop benh an ngoai tru, goi app de lay du lieu cua 1 benh an khac nen khong check rang buoc nay nua 13/05/2019
                //var emr = GetCurrentMainObject();// this._entity.MainObject as MEEmrsInfo;

                //string emrNo = response.ContainsKey("emrNo") ? response["emrNo"].ToString() : response["emrno"].ToString();
                //if (emrNo != emr.MEEmrNo)
                //{
                //    _msgNotification.Text = ("Ứng dụng nhận được dữ liệu nhưng của bệnh án có mã: " + response["emrNo"].ToString() + " Vui lòng chọn đúng bệnh nhân, bệnh án và mẫu bệnh án");
                //    return;
                //}
                string documentNo = response.ContainsKey("documentNo") ? response["documentNo"].ToString() : response["documentno"].ToString();
                string documentName = response.ContainsKey("documentName") ? response["documentName"].ToString() : response["documentname"].ToString();
                string documentDate = response.ContainsKey("documentDate") ? response["documentDate"].ToString() : response["documentdate"].ToString();
                if (documentNo != document.MEEmrDocumentFile)
                {
                    _msgNotification.Text = string.Format("Ứng dụng nhận được dữ liệu nhưng của mẫu bệnh án: '{0}' ngày {1} {2}"
                        , documentName
                        , documentDate
                        , "Vui lòng chọn đúng bệnh nhân, bệnh án và mẫu bệnh án");
                    return;
                }
                if (response == null || !response.ContainsKey("data") || response["data"] == null)
                {
                    _msgNotification.Text = ("Không có dữ liệu trả về từ HIS. Liên hệ quản trị viên để biết thêm chi tiết");
                    return;
                }
                var action = _actionsController.GetObjectByNo(response["msg"].ToString()) as MEEmrActionsInfo;
                if (action != null)
                {
                    var group = string.Empty;
                    if (response.ContainsKey(EmrParam.GuidTag))
                        group = response[EmrParam.GuidTag].ToString();

                    //TODO, BUG trong truong hop nguoi dung chon document khac truoc khi his tra ve du lieu
                    if (string.IsNullOrEmpty(group))
                        group = document.MEEmrDocumentGuid;

                    var allParams = AppMemCache.GetActionParamsFromDict(action.MEEmrActionID);
                    var updateParams = allParams.Where(o => o.MEEmrActionParamRequest == false && o.FK_MEParamID > 0).ToList();

                    var requestParams = new Dictionary<string, object>(response);
                    requestParams.Remove("data");

                    object preData = null;
                    if (_appPreData != null)
                        preData = _appPreData;

                    if (response.ContainsKey("preData"))
                        preData = response["preData"];

                    errCount = BindingDataToEmrDocument(response["data"], action, group, updateParams, requestParams, preData: preData);

                    //HIS co the goi du lieu nhieu lan cho 1 req tu EMR
                    //_appPreData = null;

                    CallChildEmrAction(action, null, group, response["data"]);
                }
                _msgNotification.Text = errCount == 0 ? "Dữ liệu đã được cập nhật" : "Có lỗi khi cập nhật dữ liệu. Vui lòng xem thông báo lỗi";
            }
        }
        private void ShowPluginMessage(JToken msg)
        {
            if (msg == null) return;
            if (msg.Type == JTokenType.String)
            {
                MessageBox.Show(msg.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else if (msg.Type == JTokenType.Object)
            {
                var type = msg.SelectToken("$." + UserFriendlyException.TYPE_KEY, false)?.ToString();
                var message = msg.SelectToken("$." + UserFriendlyException.MESSAGE_KEY, false)?.ToString();
                var title = msg.SelectToken("$." + UserFriendlyException.TITLE_KEY, false)?.ToString();
                var timeOut = msg.SelectToken("$." + UserFriendlyException.TIME_OUT, false)?.ToString();
                var timeOutInt = string.IsNullOrEmpty(timeOut) ? 0 : Convert.ToInt32(timeOut);
                switch (type)
                {
                    case UserFriendlyException.TYPE_LOG:
                        PrintMgsLog(title == null ? "THONG_BAO_TU_PLUGIN" : title, message);
                        break;
                    case UserFriendlyException.TYPE_HIGHLIGHT:
                        ShowFlashNotification(title == null ? message : $"[{title}] {message}", timeOutInt, 1000);
                        break;
                    case UserFriendlyException.TYPE_IN_EMR_SESSION:
                        ShowFlashNotificationInToolbar(title == null ? message : $"[{title}] {message}", timeOutInt);
                        break;
                    default:
                        MessageBox.Show(message, title == null ? "Thông báo" : title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        break;
                }
            }
        }

        private int BindingDataToEmrDocument(object data, MEEmrActionsInfo action, string group, List<MEEmrActionParamsInfo> updateParams,
           Dictionary<string, object> requestParams, Func<object, string, string> beforeBinding = null, object preData = null, string meEmrTemplateActionDo = "")
        {
            DeleteAllPageKeepFirst(); /*return 0;*/
            var bindingData = true;
            if (action.MEEmrActionType != EmrActionTypes.DataPlugin.ToString()
                && action.MEEmrActionType != EmrActionTypes.CalcPlugin.ToString()
                && !String.IsNullOrEmpty(action.MEEmrActionPlugin))
            {
                try
                {
                    // tranform chi thay doi duoc cac property cua data
                    // https://stackoverflow.com/questions/8708632/passing-objects-by-reference-or-value-in-c-sharp
                    if (_checkSystem)
                    {
                        var tranformPlugin = Plugin.CreatePlugin<IPluginBase>(Path.Combine(BOSApp.AppLocation, "plugins", action.MEEmrActionNo), action.MEEmrActionPlugin);
                        _sysHelper.LogTxt("information", $"Bắt đầu gọi plugin: {action.MEEmrActionType} - {action.MEEmrActionPlugin}");
                        var pluginTimeOut = BOSApp.GetSystemConfigValueInt(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.PLUGIN_TIMEOUT, 5);
                        TimeSpan maxDuration = TimeSpan.FromSeconds(pluginTimeOut);
                        var donePlugin = false;
                        var swPlugin = Stopwatch.StartNew();
                        while (!donePlugin)
                        {
                            if (swPlugin.Elapsed > maxDuration)
                            {
                                swPlugin.Stop();
                                throw new TimeoutException($"Timeout plugin, max time allow {pluginTimeOut} seconds: {action.MEEmrActionType} - {action.MEEmrActionPlugin}");
                            }
                            if (tranformPlugin is ITransformPluginV2)
                            {
                                data = (tranformPlugin as ITransformPluginV2).Tranform(data as JToken, requestParams, preData as JToken);
                            }
                            else
                            {
                                data = (tranformPlugin as ITransformPlugin).Tranform(data as JToken, requestParams);
                            }
                            donePlugin = true;
                            swPlugin.Stop();
                        }
                        _sysHelper.LogTxt("information", $"Hoàn tất plugin: {action.MEEmrActionType} - {action.MEEmrActionPlugin}");
                        _sysHelper.LogTxt("information", $"Data plugin {action.MEEmrActionType} - {action.MEEmrActionPlugin}: {data}");
                    }
                    else
                    {
                        var tranformPlugin = Plugin.CreatePlugin<IPluginBase>(Path.Combine(BOSApp.AppLocation, "plugins", action.MEEmrActionNo), action.MEEmrActionPlugin);
                        if (tranformPlugin is ITransformPluginV2)
                        {
                            data = (tranformPlugin as ITransformPluginV2).Tranform(data as JToken, requestParams, preData as JToken);
                        }
                        else
                        {
                            data = (tranformPlugin as ITransformPlugin).Tranform(data as JToken, requestParams);
                        }
                    }

                    if (data == null)
                    {
                        MessageBox.Show("Không có dữ liệu từ [Trình cắm]", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return 0;
                    }
                    // Notification plugin error message - US 1064
                    var validateCatchKey = Emr.Pluggable.Interface.Constants.USER_FRIENDLY_MESSAGE_KEY;
                    if (((JToken)data).Type == JTokenType.Array)
                    {
                        var msgs = (data as JArray).SelectTokens($"$..{validateCatchKey}");
                        if (msgs != null && msgs.Count() > 0)
                            ShowPluginMessage(msgs.FirstOrDefault());
                        bindingData = (data as JArray).SelectTokens($"$..{Emr.Pluggable.Interface.Constants.NOT_BINDING_PLUGIN_DATA_TO_DOC}").Count() == 0;
                    }
                    else
                    {
                        var msg = (data as JObject).SelectToken($"$.{validateCatchKey}");
                        if (msg != null && msg.Count() > 0)
                            ShowPluginMessage(msg);

                        bindingData = (data as JObject).SelectToken($"$.{Emr.Pluggable.Interface.Constants.NOT_BINDING_PLUGIN_DATA_TO_DOC}") == null;
                    }
                    PrintMgsLogJson($"DU-LIEU-PLUGIN: {action.MEEmrActionName}", data);
                }
                catch (TimeoutException exTimeOut)
                {
                    MessageBox.Show(exTimeOut.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    if (_checkSystem)
                    {
                        _sysHelper.LogTxt("error", exTimeOut.Message);
                    }
                    return 1;
                }
                catch (UserFriendlyException exUF)
                {
                    MessageBox.Show(exUF.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    if (_checkSystem)
                    {
                        _sysHelper.LogTxt("error", exUF.Message);
                    }
                    return 1;
                }
                catch (Exception ex)
                {
                    PrintMgsLog("Có lỗi khi gọi Trình cắm chuyển đổi dữ liệu", ex.ToString());
                    if (_checkSystem)
                    {
                        _sysHelper.LogTxt("error", ex.ToString());
                    }
                    return 1;
                }
            }

            // XuanTM - hotfix PSHN 1698.23 XML5
            if (meEmrTemplateActionDo == EmrTemplateActionDo.Notify.ToString())
            {
                bindingData = false;
            }
            // End XuanTM

            var error = 0;
            if (!bindingData) return 0;
            // if array will show select popup
            if (((JToken)data).Type == JTokenType.Array)
            {
                if ((data as JToken).Count() > 1)
                {
                    var gui = new guiDataSelection(data as JArray, updateParams)
                    {
                        Module = this,
                        StartPosition = FormStartPosition.CenterParent
                    };
                    if (gui.ShowDialog() == DialogResult.OK)
                    {
                        if (gui.SelectedRow != null)
                        {
                            error = BindingDataToEmrDocument((gui.SelectedRow as JObject).ToObject<IDictionary<string, object>>(), action, group, updateParams, beforeBinding);
                        }
                    }
                }
                else if((data as JToken).Count() > 0)
                {
                    error = BindingDataToEmrDocument((data as JToken)[0].ToObject<IDictionary<string, object>>(), action, group, updateParams, beforeBinding);
                }
            }
            else
            {
                error = BindingDataToEmrDocument((data as JObject).ToObject<IDictionary<string, object>>(), action, group, updateParams, beforeBinding);
            }

            AutoAddHeaderAndFooter();

            return error;
        }
        /// <summary>
        /// binding du lieu nhan dc tu api, app, sql
        /// </summary>
        /// <param name="data"></param>
        /// <param name="action"></param>
        private int BindingDataToEmrDocument(IDictionary<string, object> data,
            MEEmrActionsInfo action, string group, List<MEEmrActionParamsInfo> updateParams,
            Func<object, string, string> beforeBinding = null)
        {
            int errorCount = 0;
            var value = this._dataHelper.GetValueToBinding(data, EmrParam.TransactionIdTag);
            var tid = value != null ? value.ToString() : string.Empty;
            var showPrgbar = false;
            if (!BOSProgressBar.IsShowing())
            {
                showPrgbar = true;
                BOSProgressBar.Start("Điền dữ liệu vào tờ bệnh án");
            }

            var doc = this._richEditCtrl.Document;
            if (_richEditCtrl.InvokeRequired)
            {
                _richEditCtrl.BeginInvoke((Action)(() =>
                {
                    doc = this._richEditCtrl.Document;
                    doc.BeginUpdate();
                }));
            }
            else
            {
                doc.BeginUpdate();
            }
            var exlFields = new HashSet<Field>();
            try
            {
                foreach (var f in updateParams)
                {
                    var param = AppMemCache.GetParamFromDictKeyID(f.FK_MEParamID);
                    if (param == null) continue;
                    if (data.ContainsKey(param.MEParamNo))
                    {
                        //get value from return data
                        value = this._dataHelper.GetValueToBinding(data, param.MEParamNo);
                        if (value != null)
                        {
                            if (value is JArray && param.MEParamType == EmrParamTypes.Single.ToString())
                            {
                                var arr = value as JArray;
                                if (arr.Count > 1)
                                {
                                    BOSProgressBar.Close();
                                    Cursor.Current = Cursors.Default;
                                    var gui = new guiDataSelection(value as JToken, param);
                                    gui.Module = this;
                                    gui.StartPosition = FormStartPosition.CenterParent;
                                    if (gui.ShowDialog() == DialogResult.OK)
                                        value = gui.SelectedRow;
                                    else
                                        value = null;
                                }
                                else
                                {
                                    value = arr.FirstOrDefault();
                                }

                            }
                            //cho phep cau hinh hien thi hop thoai chon
                            else if (value is JArray && f.MEEmrActionParamPopupSelect)
                            {
                                var arr = value as JArray;
                                if (arr.Count > 1 || f.MEEmrActionParamPopupSelectChild)
                                {
                                    BOSProgressBar.Close();
                                    Cursor.Current = Cursors.Default;
                                    var gui = new guiDataSelection(value as JToken, param, true, f.MEEmrActionParamPopupSelectChild)
                                    {
                                        Module = this,
                                        StartPosition = FormStartPosition.CenterParent
                                    };
                                    if (gui.ShowDialog() == DialogResult.OK)
                                        value = gui.SelectedRows;
                                    else
                                        value = null;
                                }
                                else
                                {
                                    //do nothing
                                    //value = new List<JToken>(arr.FirstOrDefault());
                                }

                            }
                            //BOSProgressBar.Start("Đang cập nhật dữ liệu");
                            var formatStyle = this._dataHelper.GetValueToBinding(data, param.MEParamNo + EmrConsts.FORMAT_STYLE_PREFIX)?.ToString();
                            if (value != null)
                            {
                                beforeBinding?.Invoke(value, param.MEParamNo);
                                //get update field from doc can have multi fields
                                var updateFields = this._emrDocumentHelper.GetBindingFields(group, param.MEParamNo, tid, !f.MEEmrActionParamUpdateFirst, true, exlFields);
                                var prefix = this._emrDocumentHelper.GetFieldPrefix(updateFields.FirstOrDefault());
                                foreach (var updateField in updateFields)
                                {
                                    this._emrDocumentHelper.BindingDataToFieldV2(param, updateField, value, prefix, tid, group, f, formatStyle, exlFields, allowRecursive: true);
                                    exlFields.Add(updateField);
                                }
                            }
                        }
                    }
                    else
                    {
                        if (_notificationTab)
                        {
                            _msgLogs.Text += "\r\n" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + " Dữ liệu trả về không chứa thẻ " + param.MEParamNo + " - " + param.MEParamCaption;
                        }
                        else if (_logConfig)
                        {
                            _msgLogsTemp += "\r\n" + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss") + " Dữ liệu trả về không chứa thẻ " + param.MEParamNo + " - " + param.MEParamCaption;
                        }
                        errorCount++;
                    }
                }
            }
            catch (Exception ex)
            {
                Trace.TraceError("BINDING ERROR: {0}:{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), ex);
                throw;
            }
            finally
            {
                if (showPrgbar)
                    BOSProgressBar.Close();
                Cursor.Current = Cursors.Default;
                if (_richEditCtrl.InvokeRequired)
                {
                    _richEditCtrl.BeginInvoke((Action)(() =>
                    {
                        doc.EndUpdate();
                    }));
                }
                else
                {
                    doc.EndUpdate();
                }
            }
            return errorCount;
        }
        /// <summary>
        /// Binding du lieu cung khi khoi tao document
        /// </summary>
        /// <param name="data"></param>
        private void BindingDataToEmrDocument(IDictionary<string, object> data, string group, string prefix, bool paramUnit = true)
        {
            var doc = this._richEditCtrl.Document;
            var value = this._dataHelper.GetValueToBinding(data, EmrParam.TransactionIdTag);
            var tid = value != null ? value.ToString() : string.Empty;
            doc.BeginUpdate();
            //field needed update
            var fields = data.Keys;
            try
            {
                Field last = null;
                foreach (var f in fields)
                {
                    var codes = f.Split(EmrParam.CodeSeparator);
                    var param = AppMemCache.GetParamFromDictKeyNo(codes.Last());
                    if (param == null) continue;
                    //get value from return data
                    value = this._dataHelper.GetValueToBinding(data, f);
                    if (value != null)
                    {
                        //get update field from doc can have multi fields
                        List<Field> updateFields;
                        if (string.IsNullOrEmpty(prefix))
                            updateFields = this._emrDocumentHelper.GetBindingFields(group, f, tid, true, true);
                        else
                            updateFields = this._emrDocumentHelper.GetBindingFields(group, $"{prefix}{EmrParam.CodeSeparator}{f}", tid, false, true);

                        var newPrefix = string.Empty;
                        if (string.IsNullOrEmpty(prefix))
                            newPrefix = this._emrDocumentHelper.GetFieldPrefix(updateFields.FirstOrDefault());
                        else
                            newPrefix = prefix;

                        var formatStyle = this._dataHelper.GetValueToBinding(data, f + EmrConsts.FORMAT_STYLE_PREFIX)?.ToString();

                        foreach (var updateField in updateFields)
                        {
                            this._emrDocumentHelper.BindingDataToFieldV2(param, updateField, value, newPrefix, tid, group, null, formatStyle, allowRecursive: true, paramUnit: paramUnit);
                        }
                        last = updateFields.LastOrDefault();
                    }
                }
                if (last != null)
                {
                    this._emrDocumentHelper.GotoEndOfParam(doc, last);
                }
            }
            catch (Exception ex)
            {
                Trace.TraceError("BINDING ERROR: {0}:{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), ex);
                throw;
            }
            finally
            {

            }
            doc.EndUpdate();
        }
        #endregion

        #region Get Patient from external system
        internal void SearchPatientExternal(string patientNo)
        {
            var gui = new guiSearchPatientExternal(patientNo);
            gui.Module = this;
            var result = gui.ShowDialog();
            if (result == DialogResult.OK)
            {
                var gridView = this._entity.MEPatientsSearchList.GridView;
                var index = gridView.FocusedRowHandle;
                if (index < 0) return;
                var patient = gridView.GetRow(gridView.FocusedRowHandle) as MEPatientsInfo;
                if (patient != null)
                {
                    var emr = (this._entity.MainObject as MEEmrsInfo);
                    if (!string.IsNullOrEmpty(patient.MEEmrNo))
                    {
                        emr.MEEmrNo = patient.MEEmrNo;
                        var tran = _entity.ModuleObjects[TableName.MEEmrTransferHistoriesTableName] as MEEmrTransferHistoriesInfo;
                        tran.MEEmrTransferHistoryNo = patient.MEEmrNo;
                    }
                    patient = CreatePatientOfExt(patient);
                    _entity.InvalidateModuleObject(patient);
                    Controls["fld_lkeFK_MEEmrTypeID"].Focus();
                    emr.FK_MEPatientID = patient.MEPatientID;
                }
            }
        }
        private MEPatientsInfo CreatePatientOfExt(MEPatientsInfo patient)
        {
            var exist = this._patientCtrl.GetObjectByNo(patient.MEPatientNo) as MEPatientsInfo;
            if (exist != null) return exist;
            patient.AACreatedUser = BOSApp.CurrentUser;
            int patientId = this._patientCtrl.CreateObject(patient);
            int customerId = this._patientCtrl.CreateCustomerFromPatientInfo(patient, BOSApp.CurrentBranchInfo.BRBranchID);
            return this._patientCtrl.GetObjectByID(patientId) as MEPatientsInfo;
        }
        internal void SearchPatientFromExternal(string criteria)
        {
            this._entity.MEPatientsSearchList.Clear();
            var query = new Dictionary<string, object>();
            query.Add("query", criteria);
            List<Dictionary<string, object>> data;
            try
            {
                BOSProgressBar.Start("Đang tìm hồ sơ bệnh nhân");
                if (!string.IsNullOrEmpty(BOSApp.CurrentCompanyInfo.CSCompanyEmrApiSearchPatient))
                {
                    if (BOSApp.CurrentCompanyInfo.CSCompanyEmrApiSearchPatientMethod == HttpMethod.POST.ToString())
                        data = _api.PostAndGetList(BOSApp.CurrentCompanyInfo.CSCompanyEmrApiSearchPatient, null, query);
                    else
                        data = _api.GetList(BOSApp.CurrentCompanyInfo.CSCompanyEmrApiSearchPatient, query);

                }
                else
                {
                    data = _sqlHelper.Get(BOSApp.CurrentCompanyInfo.CSCompanyEmrSPSearchPatient, query);

                }
                if (data != null)
                {
                    foreach (var item in data)
                    {
                        this._entity.MEPatientsSearchList.Add(new MEPatientsInfo()
                        {
                            MEEmrNo = GetExternalPatientStringFieldData(item, "MEEmrNo"),
                            MEPatientNo = GetExternalPatientStringFieldData(item, "MEPatientNo"),
                            MEPatientName = GetExternalPatientStringFieldData(item, "MEPatientName").ToUpper(),
                            MEGender = GetExternalPatientStringFieldData(item, "MEGender") == "Nam" ? "Male" : "Female",
                            MEPatientBirthday = GetExternalPatientDateFieldData(item, "MEPatientBirthday"),
                            MEPatientContactCellPhone = GetExternalPatientStringFieldData(item, "MEPatientCellPhone")
                        });
                    }
                }
                this._entity.MEPatientsSearchList.GridControl.RefreshDataSource();
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                BOSProgressBar.Close();
            }
        }
        private DateTime GetExternalPatientDateFieldData(Dictionary<string, object> item, string property)
        {
            var value = GetExternalPatientStringFieldData(item, property);
            if (string.IsNullOrEmpty(value)) return DateTime.MaxValue;
            var provider = CultureInfo.InvariantCulture;
            var date = DateTime.Now;
            if (value.ToString().Length == 4)
            {
                date = new DateTime(int.Parse(value.ToString()), 1, 1);
            }
            else
                DateTime.TryParseExact(value, "dd/MM/yyyy", provider, DateTimeStyles.None, out date);
            return date;
        }
        private string GetExternalPatientStringFieldData(Dictionary<string, object> item, string property)
        {
            if (_patientDataMappings.ContainsKey(property))
            {
                if (item.ContainsKey(_patientDataMappings[property]))
                {
                    return item[_patientDataMappings[property]].ToString();
                }
            }
            return string.Empty;
        }
        /// <summary>
        /// uthv
        /// khoi tao danh sach mapping thong tin benh nhan lay tu HIS
        /// Neu khong cau hinh co che mapping BOSApp.CurrentCompanyInfo.CSCompanyEmrPatientMapping se mac dinh cac api HIS phai tra ve ten giong nhu ten của table Emr
        /// </summary>
        private void InitPatientInfoMappingForExternal()
        {
            if (BOSApp.CurrentCompanyInfo == null) return;
            _patientDataMappings = new Dictionary<string, string>();
            if (!string.IsNullOrEmpty(BOSApp.CurrentCompanyInfo.CSCompanyEmrPatientMapping))
            {
                var pairs = BOSApp.CurrentCompanyInfo.CSCompanyEmrPatientMapping.Split('|');
                foreach (var pair in pairs)
                {
                    var parames = pair.Split('=');
                    if (parames.Length == 2)
                        if (!_patientDataMappings.ContainsKey(parames[0].Trim()))
                            _patientDataMappings.Add(parames[0].Trim(), parames[1].Trim());
                }
            }
            else
            {
                //lay tat ca cac properties cua patient info 
                foreach (var prop in typeof(MEPatientsInfo).GetProperties())
                {
                    _patientDataMappings.Add(prop.Name, prop.Name);
                }

            }
        }
        #endregion

        #region RemoteCase
        public void CreateRemoteCasePatient()
        {
            //get again to make sure every thing is ok
            var patient = this._patientCtrl.GetObjectByID(this._entity.MEPatient.MEPatientID) as MEPatientsInfo;
            if (patient != null)
            {
                var api = new RemoteCaseApiHelper();
                var paramList = new Dictionary<string, object>();
                //tự tao id card number neu chua ton tai
                //lưu lại và có thể sửa sau này
                //Tạm thời - uthv
                patient.MEPatientIDCard = string.IsNullOrEmpty(patient.MEPatientIDCard) ? Guid.NewGuid().ToString().ToUpper() : patient.MEPatientIDCard;
                var body = new
                {
                    TenantKey = "A05807CD-FDE0-469D-B244-B1351D20493F", //tạm thời hard code mà mã hóa exe ko đưa ra config
                    UserNamePatientPortal = "Ext_RemoteCase_EMR",
                    PasswordEncStr = "WjCy4UUhUcrAHI3LBw2RuBypGC6VH3IMgN59efxAqrQoH7syVV1buH4qkMLCdcWPlhothSQDI4YdNoUFmuJxNtRac8jM1mGNfOuXmPO+q6jfdUUS/eNB5Fn0mM/zWtjuhiICiyilaabdH0CKpumxKirGT8HXKfVNuFysB86+73P+hznZfQTVTEO0hvlU33rh7QRW6oBDnmGnfR/okBnzYTVfDDP29IU0Jebbspy2vYWm/gDWX7iVsWMuPljuga4Y7zt11gddDQxiIvHXA4riivuvPKa20yhG27ppDLEOiRkUtgnWOP39sZhJwaKOGarMz3F/mRhnWtsTGEm1A57/lA==",
                    Patient = new
                    {
                        PatientId = string.Empty,
                        DateOfBirth = patient.MEPatientBirthday,
                        IdentityDocumentNumberType = 1000,//NIC
                        Sex = patient.MEGender == "Male" ? 0 : 1,
                        IdentityDocumentNumber = patient.MEPatientIDCard,
                        LastName = patient.MEPatientName,
                        AddressCountryCode = 1233,//VN
                        PhoneNumber = patient.MEPatientContactCellPhone,
                        Email = patient.MEPatientContactEmail,
                        BloodGroup = 1000,

                    }
                };
                var response = api.Post<ServiceCreatePatientExtResponse>("http://medcubes.clas.mobi:33333/api/PatientExternalService/CreatePatient", paramList, body);
                if (response != null)
                {
                    if (response.Success)
                    {
                        patient.MEPatientRemoteCaseID = response.PatientIdCreated.ToString();
                        patient.AAUpdatedUser = BOSApp.CurrentUser;
                        this._patientCtrl.UpdateObject(patient);
                        this._entity.MEPatient = patient;
                        MessageBox.Show("Bạn có thể tìm bệnh nhân trên RemoteCase để thăm khám và lấy kết quả vào EMR.", "Tạo mã thành công");
                        return;
                    }
                    else
                    {
                        if (response.ErrorCode.ToString() == "B300")
                        {
                            MessageBox.Show($"Bệnh nhân đã tồn tại trên RemoteCase với ID = {response.ExistingPatientId.Value.ToString()} ", "Bệnh nhân tồn tại trên RemoteCase");
                            return;
                        }
                    }
                }
                MessageBox.Show("Có lỗi xảy ra, xin thử lại hoặc liên hệ quản trị viên");
            }
        }
        private void InsertRemoteCaseData(MEEmrActionsInfo action, DocumentRange actionRange, string audioFile = null)
        {
            var patient = this._patientCtrl.GetObjectByID(this._entity.MEPatient.MEPatientID) as MEPatientsInfo;
            //get again to make sure every thing is ok
            if (string.IsNullOrEmpty(patient.MEPatientRemoteCaseID))
            {
                MessageBox.Show("Bệnh nhân chưa có mã RemoteCase. Bấm tạo mã RemoteCase trước");
                return;
            }
            this._entity.MEPatient = patient;
            switch (action.MEEmrActionNo)
            {
                case EmrActionRemoteCase.StethoscopeInsert:
                    InsertRemoteCaseStethoscopeData(action, actionRange);
                    break;
                case EmrActionRemoteCase.StethoscopeView:
                    OpenAudioFile(audioFile);
                    break;
                default:
                    break;
            }
        }
        private void InsertRemoteCaseStethoscopeData(MEEmrActionsInfo action, DocumentRange actionRange)
        {
            var api = new RemoteCaseApiHelper();
            this._entity.SelectedStethoscopesDataList = new List<StethoscopeData>();
            var paramList = new Dictionary<string, object>();
            paramList.Add("PatientId", this._entity.MEPatient.MEPatientRemoteCaseID);
            paramList.Add("CustomerId", 1);//config later
            paramList.Add("TenantId", 1);//config later
            var data = api.Post<StethoscopeResponse>("MedCubes.TeleAssessment.Server.Services.BacsiExternal.BacsiExternalService.svc/RESTSECURE/ReadStethoscopeData", null, paramList);
            if (data == null) return;
            if (data.Success && data.Values != null)
            {
                var doc = _richEditCtrl.Document;
                doc.BeginUpdate();
                var linkRange = actionRange;
                try
                {
                    foreach (var item in data.Values)
                    {
                        item.LocalizationName = EmrStethoscopeLocation.LocationMap[item.Localization];

                    }
                    this._entity.StethoscopesDataList = data.Values;
                    var gui = new DSMEEMR102();
                    gui.Module = this;
                    var result = gui.ShowDialog();
                    if (result == DialogResult.OK)
                    {
                        foreach (var item in this._entity.SelectedStethoscopesDataList)
                        {
                            var time = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");// item.CreatedAt.ToString("dd/MM/yyyy HH:mm:ss");
                            string audioFile = $"stethoscope_audio_{item.PatientId}_{DateTime.Now.ToString("ddMMyyyy_hhmmssffff")}.mp3";
                            //Save audio file
                            string path = string.Format(@"{0}\Audio\{1}", _documentPath, audioFile);
                            File.WriteAllBytes(path, item.AudioFile);
                            var caption = $"{item.LocalizationName} {time}";
                            linkRange = doc.InsertText(linkRange.End, "\n");
                            linkRange = doc.InsertText(linkRange.End, caption);
                            //Convert the inserted text to the field 
                            var link = doc.Hyperlinks.Create(linkRange);
                            link.Anchor = "Click vào để nghe âm thanh";
                            link.NavigateUri = $"{EmrActionRemoteCase.StethoscopeView}{EmrParam.TagCodeSeparator}{audioFile}";
                            link.ToolTip = "Click vào để nghe âm thanh";
                        }
                        this._entity.SelectedStethoscopesDataList.Clear();
                    }
                }
                catch (Exception)
                {

                    throw;
                }
                finally
                {
                    doc.EndUpdate();
                }
            }
        }
        private void OpenAudioFile(string name)
        {
            string path = string.Format(@"{0}\Audio\{1}", _documentPath, name);
            FileTemplateManager fileMng = new FileTemplateManager();
            fileMng.DownloadFile("/Audio/", name, path);
            if (!File.Exists(path))
            {
                MessageBox.Show("File audio không tồn tại ở địa chỉ. " + path);
                return;
            }
            Process.Start(path);
        }
        #endregion

        #region Param info
        internal void ViewParamInfo()
        {
            var doc = _richEditCtrl.Document;
            var caretPosition = doc.CaretPosition.ToInt();
            var properties = this._emrDocumentHelper.GetParamInfos(doc, caretPosition);
            if (properties != null)
            {
                var gui = new guiParamProperties(properties);
                gui.ShowDialog();
            }
            else
                MessageBox.Show("Đặt con trỏ giữa một thẻ dữ liệu.", "Không tìm thấy thẻ tương ứng", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        public void ViewParamValues()
        {
            var doc = _richEditCtrl.Document;
            if (_hospitalProject != "ChamCuu" && !_emrDocumentHelper.IsAllowEditField(doc.CaretPosition))
            {
                ShowFlashNotification("Thẻ dữ liệu đã bị chặn sửa thủ công", 3000);
                return;
            }
            var field = this._emrActionHelper.GetEmrtFieldAtPosition(doc.CaretPosition);
            string valueField = "";
            if (field != null)
            {
                if (field.FieldCode.StartsWith("HYPERLINK")) return;
                var group = field.Gid;

                var paramCtrl = new MEParamsController();
                var param = paramCtrl.GetObjectByNo(field.ParamNo) as MEParamsInfo;
                if (param == null)
                {
                    MessageBox.Show("Không tìm thấy thẻ dữ liệu này trong hệ thống. Kiểm tra lại mã thẻ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                if (param.MEParamExtend)
                {
                    var fieldAtPosition = _emrDocumentHelper.GetFieldAtPosition(doc.CaretPosition.ToInt());
                    if (fieldAtPosition != null)
                        valueField = _emrParser.GetFieldValue(fieldAtPosition);
                }
                var config = _paramLookupCtrl.GetObjectByID(param.FK_MEParamLookupID) as MEParamLookupsInfo;
                var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;

                var fieldPath = _emrDocumentHelper.GetTemplateFieldPath(field.CodeLevelStr);
                var templateParams = AppMemCache.GetTemplateParams(document.FK_METemplateID);
                var tempParam = templateParams.Where(o => o.METemplateParamPath == fieldPath).FirstOrDefault();
                if (tempParam != null)
                {
                    if (!string.IsNullOrEmpty(tempParam.MEParamFormatString))
                        param.MEParamFormatString = tempParam.MEParamFormatString;
                    if (!string.IsNullOrEmpty(tempParam.MEParamFormatType))
                        param.MEParamFormatType = tempParam.MEParamFormatType;
                }
                List<MEParamLookupDatasInfo> lookupList = _paramLookupDataCtrl.GetAllParamLookupDatas(param.FK_MEParamLookupID,
                    document.MEEmrDocumentID, document.FK_MEEmrID, BOSApp.CurrentUsersInfo.ADUserID,
                    BOSApp.CurrentEmployeesInfo.HREmployeeID, BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID);

                lookupList = lookupList.OrderBy(o => o.MEParamLookupDataOrder).ToList();
                var isListParam = param.MEParamType == EmrParamTypes.List.ToString();
                if (lookupList != null && lookupList.Count > 0)
                {
                    var selectedVals = new List<string>();
                    if (isListParam)
                        selectedVals = _emrParser.GetArrayValFromParam<string>(field.CodeLevelArr, field.Gid, templateParams);
                    DSMEEMR106 form = new DSMEEMR106(config, lookupList, isListParam, selectedVals, param, valueField)
                    {
                        Module = this,
                        StartPosition = FormStartPosition.CenterParent
                    };
                    if (form.ShowDialog() == DialogResult.OK && form.SelectedData.Count > 0)
                    {
                        var prefix = this.GetParamPrefixByLevelWithoutIndexing(field.CodeLevelStr);
                        var isPart = config.MEParamLookupPart;
                        var relationParams = new List<MEParamsInfo>();
                        if (isPart)
                        {
                            relationParams = paramCtrl.GetAllByTemplateAndLookupDataParams(param.FK_METemplateID, param.FK_MEParamLookupID);
                            relationParams = relationParams.Where(m => !string.IsNullOrEmpty(m.MEParamMap)).ToList();
                        }

                        if (param.MEParamType == EmrParamTypes.List.ToString())
                        {
                            Dictionary<string, object> data = new Dictionary<string, object>();
                            if (isPart && !string.IsNullOrEmpty(param.MEParamMap)) // if param not config
                            {
                                DataAddValues(data, relationParams, form.SelectedData);
                            }
                            else
                            {
                                List<string> selectedValues = new List<string>();
                                foreach (var t in form.SelectedData)
                                    selectedValues.Add(t.MEParamLookupDataValue);
                                data.Add(param.MEParamNo, new JArray(selectedValues.ToArray()));
                            }
                            try
                            {
                                if (param.MEParamExtend)
                                {
                                    foreach (var item in data)
                                    {
                                        JArray valueNew = (JArray)item.Value;
                                        valueNew.Add("");
                                    }
                                }
                            }
                            catch
                            {

                            }

                            BindingDataToEmrDocument(data, group, prefix);
                        }
                        else
                        {
                            Dictionary<string, object> data = new Dictionary<string, object>();
                            if (isPart && !string.IsNullOrEmpty(param.MEParamMap)) // if param config
                            {
                                foreach (var paramI in relationParams)
                                {
                                    DataAddValue(data, paramI, form.SelectedData[0]);
                                    BindingDataToEmrDocument(data, group, prefix);
                                    //UpdateReleatedParams(field, group, paramI, data, tempParam);
                                }
                            }
                            else
                            {
                                data.Add(param.MEParamNo, form.SelectedData[0].MEParamLookupDataValue);
                                BindingDataToEmrDocument(data, group, prefix);
                                UpdateReleatedParams(field, group, param, data, tempParam);
                            }
                        }
                    }
                }
                else if (param.MEParamFormatType == ParamFormatType.DateTime.ToString()
                    || param.MEParamFormatType == ParamFormatType.Date.ToString())
                {
                    if (_hospitalProject == "ChamCuu" && !_emrDocumentHelper.IsAllowEditField(doc.CaretPosition))
                    {
                        ShowFlashNotification("Thẻ dữ liệu đã bị chặn sửa thủ công", 3000);
                        return;
                    }
                    var form = new guiDateTimeSelection(param.MEParamFormatType, param.MEParamFormatString, _dataHelper)
                    {
                        Module = this,
                        StartPosition = FormStartPosition.CenterParent
                    };
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        var prefix = this.GetParamPrefixByLevelWithoutIndexing(field.CodeLevelStr);
                        Dictionary<string, object> data = new Dictionary<string, object>
                        {
                            { param.MEParamNo, form.InputValue }
                        };
                        BindingDataToEmrDocument(data, group, prefix);
                        UpdateReleatedParams(field, group, param, data, tempParam);
                    }
                }
                else
                    MessageBox.Show("Thẻ không có danh mục để chọn", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
                MessageBox.Show("Đặt con trỏ giữa một thẻ dữ liệu.", "Không tìm thấy thẻ tương ứng", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void DataAddValues(Dictionary<string, object> data, List<MEParamsInfo> relationParams, List<MEParamLookupDatasInfo> selectedData)
        {
            foreach (var param in relationParams)
            {
                string colMap = param.MEParamMap;
                List<string> selectedValues = new List<string>();
                foreach (var t in selectedData)
                {
                    switch (colMap)
                    {
                        case "MEParamLookupDataKey":
                            selectedValues.Add(t.MEParamLookupDataKey);
                            break;
                        case "MEParamLookupDataText":
                            selectedValues.Add(t.MEParamLookupDataText);
                            break;
                        case "MEParamLookupDataGroup":
                            selectedValues.Add(t.MEParamLookupDataGroup);
                            break;
                        default:
                            selectedValues.Add(t.MEParamLookupDataValue);
                            break;
                    }
                }
                data.Add(param.MEParamNo, new JArray(selectedValues.ToArray()));
            }
        }

        private void DataAddValue(Dictionary<string, object> data, MEParamsInfo param, MEParamLookupDatasInfo selectedData)
        {
            string colMap = param.MEParamMap;
            switch (colMap)
            {
                case "MEParamLookupDataKey":
                    data.Add(param.MEParamNo, selectedData.MEParamLookupDataKey);
                    break;
                case "MEParamLookupDataText":
                    data.Add(param.MEParamNo, selectedData.MEParamLookupDataText);
                    break;
                case "MEParamLookupDataGroup":
                    data.Add(param.MEParamNo, selectedData.MEParamLookupDataGroup);
                    break;
                default:
                    data.Add(param.MEParamNo, selectedData.MEParamLookupDataValue);
                    break;
            }
        }

        private void UpdateReleatedParams(EmrField field, string group, MEParamsInfo param, Dictionary<string, object> data, METemplateParamsInfo tempParam)
        {
            //uthv https://trello.com/c/JHqukqtT
            if (field.CodeLevelArr.Length > 1 && int.TryParse(field.CodeLevelArr[field.CodeLevelArr.Length - 2], out int y))
            {
                var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
                var templateParams = AppMemCache.GetTemplateParams(document.FK_METemplateID);
                var releatedParams = templateParams.Where(o =>
                o != tempParam &&
                o.METemplateParamPath.EndsWith("[*]." + param.MEParamNo)).ToList();
                var msg = string.Empty;
                foreach (var item in releatedParams)
                {
                    var paths = item.METemplateParamPath.Split('.');
                    var p = AppMemCache.GetParamFromDictKeyNo(paths[paths.Length - 2].Replace("[*]", string.Empty)) as MEParamsInfo;
                    msg += $"+ {p.MEParamName} - {param.MEParamName}[{y + 1}]\n";
                }
                if (releatedParams.Count > 0)
                {
                    if (MessageBox.Show($"Cập nhật dữ liệu cho các thẻ liên quan: \n{msg}", "Cập nhật thẻ liên quan", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        foreach (var item in releatedParams)
                        {
                            var path = item.METemplateParamPath.Replace('.', '-').Replace("[*]", "-0").Replace("0-" + param.MEParamNo, y.ToString());
                            BindingDataToEmrDocument(data, group, path);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// remove last index if exist
        /// benhnhan-chandoan-1 => benhnhan
        /// </summary>
        /// <param name="v"></param>
        /// <returns></returns>
        private string GetParamPrefixByLevelWithoutIndexing(string codeByLevel)
        {
            string[] codes = codeByLevel.Split(EmrParam.CodeSeparator);
            int index;
            if (int.TryParse(codes.Last(), out index))
            {
                return string.Join(EmrParam.CodeSeparator.ToString(), codes.Take(codes.Length - 2));
            }
            else
            {
                return string.Join(EmrParam.CodeSeparator.ToString(), codes.Take(codes.Length - 1));
            }
        }
        #endregion

        #region Chart
        internal void ViewChartDesigner()
        {
            if (!IsAllowEditPositionWithFlashNotification(_richEditCtrl.Document.CaretPosition)) return;

            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            if (_entity.TemplateChartList.Count == 1)
            {
                this.ViewChartForm(_entity.TemplateChartList[0]);
                return;
            }
            var guiSelection = new DSMEEMR105(_entity.TemplateChartList);
            guiSelection.Module = this;
            if (guiSelection.ShowDialog() == DialogResult.OK)
            {
                this.ViewChartForm(guiSelection.SelectedChart);
            }

        }

        private void GetAllListParamAndData(JObject data, Dictionary<string, object> results)
        {
            if (data == null) return;
            foreach (var token in data)
            {
                if (token.Value.Type == JTokenType.Array)
                {
                    if (!results.ContainsKey(token.Key))
                        results.Add(token.Key, token.Value);
                }
                else if (token.Value.Type == JTokenType.Object)
                {
                    GetAllListParamAndData(token.Value as JObject, results);
                }
            }
        }

        internal void ViewChartForm(METemplateChartsInfo chartConfig)
        {
            var doc = _richEditCtrl.Document;
            var selectedRange = doc.Selection;
            var fields = this._emrDocumentHelper.GetAllDataFieldInRangeOrDocument(selectedRange);
            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            var templateParams = AppMemCache.GetTemplateParams(document.FK_METemplateID);
            var json = this._emrParser.ParserFieldsToJson(fields, templateParams);
            var data = JsonConvert.DeserializeObject(json) as JObject;
            var ctrl = new METemplateChartSeriesController();
            var listChartSeries = ctrl.GetListBusinessObjects<METemplateChartSeriesInfo>(ctrl.GetAllDataByForeignColumn("FK_METemplateChartID", chartConfig.METemplateChartID));

            var gui = new guiChartConfig(data, chartConfig, listChartSeries, AppMemCache.GetTemplateParamsDictPath(document.FK_METemplateID))
            {
                AllowSaveConfig = false,
                Module = this
            };
            if (gui.ShowDialog() == DialogResult.OK)
            {
                this._emrDocumentHelper.InsertImageToDocumentAtParam(gui.ChartImage, chartConfig.METemplateChartContainerParam, document.MEEmrDocumentGuid, 0, 0, false, false);
            }
        }
        #region Ve hinh
        internal void OpenDrawTool(int imageID = 0, DocumentImage oldImage = null, int width = 0, int height = 0)
        {
            if (!IsAllowEditPositionWithFlashNotification(_richEditCtrl.Document.CaretPosition)) return;

            byte[] oldByteImage = null;
            if (oldImage != null)
            {
                var format = DevExpress.Office.Utils.OfficeImageFormat.Png;
                if (oldImage.Image.CanGetImageBytes(format))
                {
                    oldByteImage = oldImage.Image.GetImageBytes(format);
                    width = (int)oldImage.Size.Width;
                    height = (int)oldImage.Size.Height;
                }
            }
            var form = new guiDraw(imageID, oldByteImage, width, height)
            {
                Module = this
            };
            form.InitializeControls(form.Controls);
            form.StartPosition = FormStartPosition.CenterParent;
            if (form.ShowDialog() != DialogResult.OK) return;

            var doc = _richEditCtrl.Document;
            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            var position = this._richEditCtrl.Document.CaretPosition;
            try
            {
                ((DevExpress.XtraRichEdit.Model.DocumentModel)this._richEditCtrl.Model).History.BeginTransaction();
                DocumentImage image = null;
                if (form.EmrImage != null && form.EmrImage.FK_MEParamContainerID > 0)
                {
                    var param = AppMemCache.GetParamFromDictKeyID(form.EmrImage.FK_MEParamContainerID);
                    if (param != null)
                    {
                        var imgIdStr = form.EmrImage.MEEmrImageID.ToString();
                        image = this._emrDocumentHelper.InsertImageToDocumentAtParamWithHidenContent(form.ToImage(),
                            param.MEParamNo, document.MEEmrDocumentGuid, imgIdStr,
                            form.ImgWidth, form.ImgHeight);
                        // image tag <[imgIdStr][image]>
                        var imageField = this._emrDocumentHelper.GetFirstFieldByPath(param.MEParamNo, document.MEEmrDocumentGuid);
                        if (imageField != null)
                        {
                            var imageRange = doc.InsertSingleLineText(imageField.Range.End, string.Empty);
                            this._emrActionHelper.AssignEmrImageParamValue(_entity.MEImageParamList.ToList(), imageRange, document.MEEmrDocumentGuid);
                        }
                    }
                }

                if (image == null)
                {
                    image = this._emrDocumentHelper.InsertImageToDocument(form.ToImage(), position, form.ImgWidth, form.ImgHeight);
                    this._emrActionHelper.AssignEmrImageParamValue(_entity.MEImageParamList.ToList(), image.Range, document.MEEmrDocumentGuid);
                }

                if (oldImage != null)
                {
                    doc.Replace(oldImage.Range, string.Empty);
                }
            }
            catch (Exception) { throw; }
            finally
            {
                ((DevExpress.XtraRichEdit.Model.DocumentModel)this._richEditCtrl.Model).History.EndTransaction();
            }
        }
        internal void EditImage()
        {
            var exportCf = new DevExpress.XtraRichEdit.Export.PlainTextDocumentExporterOptions() { ExportHiddenText = true };
            var position = this._richEditCtrl.Document.CaretPosition;
            var doc = _richEditCtrl.Document;
            var docImage = doc.Images.Where(i => i.Range.End == position).FirstOrDefault();
            if (docImage != null)
            {
                int emrImageID = 0;
                var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
                int.TryParse(_emrDocumentHelper.GetFieldValueAtPosition(doc, position.ToInt()), out emrImageID);
                _entity.MEImageParamList.Clear();
                if (emrImageID > 0)
                {
                    _entity.MEImageParamList.Invalidate(emrImageID);
                    foreach (var item in _entity.MEImageParamList)
                    {
                        // TODO so hardly but it run
                        var param = AppMemCache.GetParamFromDictKeyID(item.FK_MEParamID);
                        var text = _emrDocumentHelper.GetFieldValue(doc, _emrDocumentHelper.GetFirstFieldByPath(param.MEParamNo, document.MEEmrDocumentGuid));
                        if (!string.IsNullOrEmpty(text))
                        {
                            var values = text.Split(':');
                            item.MEEmrImageParamCaption = values[0].Trim();
                            if (values.Length > 0)
                                item.MEEmrImageParamValue = values[1].Trim();
                        }
                    }

                }
                OpenDrawTool(emrImageID, docImage);
            }
        }

        #endregion

        public DocumentImage InsertImageToDocument(Image image, int width = 0, int height = 0, bool newLine = true)
        {
            return this._emrDocumentHelper.InsertImageToDocument(image, this._richEditCtrl.Document.CaretPosition, width, height, newLine);
        }
        #endregion

        #region Print
        private RichEditControl CloneDocumentToPrint()
        {
            if (!IsNothingToSaveDocumentContent()) return null;
            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            if (string.IsNullOrEmpty(document.MEEmrDocumentFile)) return null;
            string fromFile = string.Format(@"{0}\Emr\{1}\{2}.docx", _documentPath, document.FK_MEEmrID, document.MEEmrDocumentFile);
            string toFile = string.Format(@"{0}\Temp\{1}{2}.docx", _documentPath, document.MEEmrDocumentFile, Guid.NewGuid().ToString());
            Directory.CreateDirectory(string.Format(@"{0}\Temp", _documentPath));
            File.Copy(fromFile, toFile);
            try
            {
                using (OfficeOpenXmlCrypto.OfficeCryptoStream stream = OfficeOpenXmlCrypto.OfficeCryptoStream.Open(toFile, this._emrDocumentHelper.ShareEmrPassword))
                {
                    this._tempRichEditCtrl.LoadDocument(stream, DocumentFormat.OpenXml);
                }
            }
            catch (OfficeOpenXmlCrypto.InvalidPasswordException ex)
            {
                MessageBox.Show("Không giải mã được tờ bệnh án", "Có thể đã có lỗi trong quá trình mã hóa và upload tờ bệnh án. Không mở được tập tin.");
            }
            RemoveAllForPrint(this._tempRichEditCtrl, document.FK_METemplateID);
            return this._tempRichEditCtrl;
        }
        public void RemoveAllForPrint(RichEditControl richContrl, int templateID)
        {
            var template = _entity.METemplateList.Where(t => t.METemplateID == templateID).FirstOrDefault();
            if (template == null)
                template = _templateCtrl.GetObjectByID(templateID) as METemplatesInfo;

            //_entity.METemplate = template; uthv remove 20200831

            this.RemoveAllTagForPrint(richContrl, template);
            this.RemoveAllCommentForPrint(richContrl);
            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            var templateParams = AppMemCache.GetTemplateParams(document.FK_METemplateID);
            if (templateParams.FirstOrDefault()?.FK_METemplateID == templateID)
                this._emrDocumentHelper.RemoveAllHiddenDataForPrint(richContrl, templateID);
            else
            {
                templateParams = _templateParamCtrl.GetAllTemplateParamObjectByTemplateID(templateID);
                this._emrDocumentHelper.RemoveAllHiddenDataForPrint(richContrl, templateID, templateParams);
            }
            _emrDocumentHelper.ProgressHyperlinksParamForPrint(richContrl.Document);

            if (template.METemplateSignatureNotAlone)
                _emrDocumentHelper.SignatureAndContentOnTheSamePage(richContrl, richContrl.Document, richContrl.DocumentLayout);

            if (template.METemplateNo.ToUpper() == "TDT")
            {
                var table = richContrl.Document.Tables.OrderByDescending(t => t.Range.Length).FirstOrDefault();
                if (table != null && table.Rows.Count > 2)
                {
                    try
                    {
                        int pageCount = richContrl.DocumentLayout.GetFormattedPageCount();
                        for (int i = 0; i < pageCount; i++)
                        {
                            var collector = new TableCellLayoutVisitor(richContrl);
                            collector.Visit(richContrl.DocumentLayout.GetPage(i));
                            foreach (var range in collector.Ranges)
                            {
                                var cell = richContrl.Document.Tables.GetTableCell(richContrl.Document.CreatePosition(range.Start));
                                if (cell != null && table == cell.Table)
                                {
                                    var pos = richContrl.Document.CreatePosition(cell.ContentRange.End.ToInt() - 1);
                                    var p = richContrl.Document.Paragraphs.Insert(pos);
                                    p = richContrl.Document.Paragraphs.Insert(pos);
                                    p = richContrl.Document.Paragraphs.Insert(pos);
                                    p = richContrl.Document.Paragraphs.Insert(pos);
                                    p = richContrl.Document.Paragraphs.Insert(pos);
                                }
                            }
                        }
                    }
                    catch (Exception) {/*do nothing to make sure not error*/}
                }
            }
        }
        /// <summary>
        /// remove all hyperlink and tag to print
        /// TODO: not remove info hyperlink ext: rc link...
        /// </summary>
        /// <param name="richContrl"></param>
        private void RemoveAllTagForPrint(RichEditControl richContrl, METemplatesInfo template)
        {
            var doc = richContrl.Document;
            RemoveAllTagForPrint(doc, template);
        }
        private void RemoveAllTagForPrint(RichEditDocumentServer server, METemplatesInfo template)
        {
            var doc = server.Document;
            RemoveAllTagForPrint(doc, template);
        }
        private void RemoveAllTagForPrint(DevExpress.XtraRichEdit.API.Native.Document doc, METemplatesInfo template)
        {
            doc.BeginUpdate();
            doc.ReplaceAll(EmrParam.BeginTag, " ", SearchOptions.None);
            doc.ReplaceAll(EmrParam.EndTag, " ", SearchOptions.None);
            for (int i = 0; i < doc.Hyperlinks.Count; i++)
            {
                var link = doc.Hyperlinks[i];
                try
                {
                    if (RemoveBreakingHyperlink(doc, link))
                        continue;

                    /* if (template.METemplateNo.ToUpper() == "TDT")
                     {
                         if (link.NavigateUri.StartsWith("thaythethe_chandoan|gid="))
                         {
                             //14 = length('Thêm chẩn đoán')
                             doc.Replace(doc.CreateRange(doc.CreatePosition(link.Range.End.ToInt() - 14), 14), " ");
                             continue;
                         }
                     }
                     // still not find solution, hard code...
                     if (template.METemplateNo.ToUpper() == "BASK")
                     {
                         if (link.NavigateUri.StartsWith("bask_phuchop_thongtindathai|gid="))
                         {
                             //12 = length('Chọn số thai')
                             doc.Replace(doc.CreateRange(doc.CreatePosition(link.Range.End.ToInt() - 12), 12), " ");
                             continue;
                         }
                     }
                     */
                }
                catch (Exception) {/*do nothing*/}
                link.ToolTip = null;
                var pos = link.Range.Start;
                doc.Replace(link.Range, string.Empty);
                doc.Hyperlinks.Remove(link);
                if (_entity.METemplate != null && _entity.METemplate.METemplateRemoveEmptyParagraph)
                {
                    var parag = doc.Paragraphs.Get(pos);
                    if (parag != null)
                    {
                        var text = doc.GetText(parag.Range);
                        if (text != null) text = text.Trim();
                        if (string.IsNullOrEmpty(text))
                            doc.Delete(parag.Range);
                    }
                }
                i--;
            }
            doc.EndUpdate();
        }
        private void RemoveAllCommentForPrint(RichEditControl richContrl)
        {
            var doc = richContrl.Document;
            if (doc.Comments.Count == 0) return;
            doc.BeginUpdate();
            for (int i = 0; i < doc.Comments.Count; i++)
            {
                doc.Comments.Remove(doc.Comments[i]);
                i--;
            }
            doc.EndUpdate();
        }
        /// <summary>
        /// Process hyperlink sticky with data tag
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="link"></param>
        /// <returns></returns>
        private bool RemoveBreakingHyperlink(DevExpress.XtraRichEdit.API.Native.Document doc, Hyperlink link)
        {
            var endLink = link.Range.End.ToInt();
            var startLink = link.Range.Start.ToInt();
            var cp = doc.BeginUpdateCharacters(doc.CreateRange(endLink - 1, 1));
            var underline = cp.Underline;
            var foreColor = cp.ForeColor;
            doc.EndUpdateCharacters(cp);
            string text = null;
            if (underline == UnderlineType.None && foreColor != Color.Blue)
            {
                text = doc.GetText(link.Range);
                var nIdx = 0;
                while (true)
                {
                    if (nIdx + 1 == text.Length || nIdx == -1) break;
                    nIdx = text.IndexOf(" ", nIdx + 1);
                    cp = doc.BeginUpdateCharacters(doc.CreateRange(startLink + nIdx + 2, 1));
                    underline = cp.Underline;
                    doc.EndUpdateCharacters(cp);
                    if (underline == UnderlineType.None)
                    {
                        while (nIdx > 0)
                        {
                            nIdx--;
                            cp = doc.BeginUpdateCharacters(doc.CreateRange(startLink + nIdx, 1));
                            underline = cp.Underline;
                            doc.EndUpdateCharacters(cp);
                            if (underline != UnderlineType.None)
                                break;
                        }
                        doc.Replace(doc.CreateRange(doc.CreatePosition(startLink), nIdx + 1), string.Empty);
                        break;
                    }
                }
                return true;
            }
            else
            {
                cp = doc.BeginUpdateCharacters(doc.CreateRange(link.Range.Start, 1));
                underline = cp.Underline;
                foreColor = cp.ForeColor;
                doc.EndUpdateCharacters(cp);
                if (underline == UnderlineType.None && foreColor != Color.Blue)
                {
                    if (string.IsNullOrEmpty(text))
                        text = doc.GetText(link.Range);
                    for (int nIdx = text.Length - 1; nIdx >= 0; nIdx--)
                    {
                        if (text[nIdx] != ' ') continue;
                        cp = doc.BeginUpdateCharacters(doc.CreateRange(endLink - (text.Length - nIdx) - 2, 1));
                        underline = cp.Underline;
                        doc.EndUpdateCharacters(cp);
                        if (underline == UnderlineType.None)
                        {
                            nIdx = text.Length - nIdx;
                            while (nIdx > 0)
                            {
                                nIdx--;
                                cp = doc.BeginUpdateCharacters(doc.CreateRange(endLink - nIdx, 1));
                                underline = cp.Underline;
                                doc.EndUpdateCharacters(cp);
                                if (underline != UnderlineType.None)
                                    break;
                            }
                            doc.Replace(doc.CreateRange(doc.CreatePosition(endLink - nIdx), nIdx), string.Empty);
                            break;
                        }
                    }
                    return true;
                }
            }

            return false;
        }
        public void ShowPrintDialog(RichEditControl richContrl)
        {
            if (richContrl != null)
            {
                PrintableComponentLink link = new PrintableComponentLink(new PrintingSystem());
                link.Component = richContrl;
                // Disable warnings.
                link.PrintingSystem.ShowMarginsWarning = false;
                link.PrintingSystem.ShowPrintStatusDialog = true;
                link.PrintDlg();
            }
        }
        internal void ShowPrintDialog()
        {
            var richContrl = this.CloneDocumentToPrint();
            ShowPrintDialog(richContrl);
        }
        internal void ShowPreviewPrintDialog()
        {
            var richContrl = this.CloneDocumentToPrint();
            ShowPreviewPrintDialog(richContrl);
        }
        internal void ShowPreviewPrintDialog(RichEditControl richContrl)
        {
            if (richContrl != null)
            {
                var printSys = new PrintingSystem();
                PrintableComponentLink link = new PrintableComponentLink(printSys);
                link.Component = richContrl;
                // Disable warnings.
                link.PrintingSystem.ShowMarginsWarning = false;
                link.PrintingSystem.ShowPrintStatusDialog = true;
                link.ShowPreview();
            }
        }
        internal void QuickPrintDocument(RichEditControl richContrl)
        {
            if (richContrl != null)
            {
                PrintableComponentLink link = new PrintableComponentLink(new PrintingSystem());
                link.Component = richContrl;
                // Disable warnings.
                link.PrintingSystem.ShowMarginsWarning = false;
                link.PrintingSystem.ShowPrintStatusDialog = true;
                link.Print();
            }
        }
        internal void QuickPrintDocument()
        {
            var documentNo = getDocumentNo(_entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo);
            HistoryEmr(_entity.MainObject as MEEmrsInfo, cstObjectHistoryActionChange, $"In tờ bệnh án {documentNo}.");
            var richContrl = this.CloneDocumentToPrint();
            QuickPrintDocument(richContrl);
        }
        #endregion

        #region Search
        protected override DataSet GetSearchData(ref string searchQuery)
        {
            var stateConds = GetStatePermiQueryConditionStr(TableName.MEEmrsTableName, false);
            var search = CurrentModuleEntity.SearchObject as MEEmrsInfo;
            object from = null; object to = null;
            foreach (Control ctrl in SearchScreen.CriteriaSection.Controls)
                if (ctrl.Tag != null && ctrl.Tag.ToString() == BOSScreen.SearchControl)
                {
                    if (ctrl.Name == "fld_dteMEEmrCreatedDateSearchFrom")
                        from = ((BaseEdit)ctrl).EditValue;
                    if (ctrl.Name == "fld_dteMEEmrCreatedDateSearchTo")
                        to = ((BaseEdit)ctrl).EditValue;
                }

            var patientGroup = string.Empty;
            DateTime? fromDateOut = null;
            DateTime? toDateOut = null;
            // 1523 Đã ký CA
            var mEEmrArchiveStatus = 0;
            if (BOSApp.CurrentUserGroupInfo.ADUserGroupRole == UserGroupRole.admin.ToString())
            {
                mEEmrArchiveStatus = 1;
            }

            var paramValues = new object[]
            {
                search.MEEmrNo,
                search.MEEmrStatus,
                from,
                to,
                search.FK_MEPatientID,
                search.FK_MEEmrTypeID,
                patientGroup,
                fromDateOut,
                toDateOut,
                search.FK_HRDepartmentID,
                BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID,
                BOSApp.CurrentEmployeesInfo.HREmployeeID,
                mEEmrArchiveStatus,
                BOSApp.GetUserEmrViewPermission(),
                stateConds
            };
            return _emrCtrl.Search(paramValues);
        }
        public override void ResetSearchObject()
        {
            base.ResetSearchObject();
        }
        #endregion

        #region Draw image
        internal List<MEEmrImagesInfo> GetListImageByDepartmentAndGroup(int departmentID, int groupID)
        {
            return _imageCtrl.GetListImageByDepartmentAndGroup(departmentID, groupID, BOSApp.CurrentEmployeesInfo.HREmployeeID);
        }
        internal List<MEEmrImagePatternsInfo> GetPatternByImage(int mEEmrImageID)
        {
            return _imagePatternCtrl.GetListBusinessObjects<MEEmrImagePatternsInfo>(
                _imagePatternCtrl.GetAllDataByForeignColumn("FK_MEEmrImageID", mEEmrImageID))
                .OrderBy(o => o.MEEmrImagePatternOrder).ToList();
        }
        internal void InvalidateImageParam(int mEEmrImageID)
        {
            _entity.MEImageParamList.Invalidate(mEEmrImageID);
        }
        #endregion

        #region Transfer
        internal void TransferDepartment()
        {
            var emr = _entity.MainObject as MEEmrsInfo;
            if (BOSApp.CurrentUserGroupInfo.ADUserGroupRole != UserGroupRole.admin.ToString())
            {
                if (emr.FK_HRDepartmentID != BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID)
                {
                    MessageBox.Show("Chỉ có khoa đang quản lý bệnh án mới có quyền thực hiện chuyển khoa.", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    return;
                }
            }

            var docs = _emrDocumentCtrl.GetListBusinessObjects<MEEmrDocumentsInfo>(_emrDocumentCtrl.GetAllDataByForeignColumn("FK_MEEmrID", emr.MEEmrID));
            foreach (var doc in docs)
            {
                if (doc.FK_EditingUserID > 0 && doc.MEEmrDocumentHoldMachineMac != _macAddress)
                {
                    ShowDocumentEditingUser(doc);
                    return;
                }
            }

            this.ActionEdit();
            // Truong hop thay doi tren to benh an va bam Cancle khi hoi luu
            if (this.Toolbar.IsNullOrNoneAction()) return;

            var gui = new guiTransfer();
            gui.Module = this;
            _entity.SetDefaultModuleObject(TableName.MEEmrTransferHistoriesTableName);

            gui.InitializeControls(gui.Controls);
            gui.StartPosition = FormStartPosition.CenterParent;

            var lastTransfer = _entity.MEEmrTranfersList.LastOrDefault();

            var tran = _entity.ModuleObjects[TableName.MEEmrTransferHistoriesTableName] as MEEmrTransferHistoriesInfo;
            tran.MEEmrTransferHistoryID = 0;
            if (lastTransfer != null)
            {
                tran.FK_HRDepartmentFromID = lastTransfer.FK_HRDepartmentToID;
            }
            else
            {
                tran.FK_HRDepartmentFromID = BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID;
            }
            tran.MEEmrTransferHistoryDate = DateTime.Now;
            tran.FK_HREmployeeFromID = BOSApp.CurrentEmployeesInfo.HREmployeeID;

            // mac dinh sinh ma chuyen khoa nhung cho phep thay doi
            tran.MEEmrTransferHistoryNo = emr.MEEmrNo + "." + (_entity.MEEmrTranfersList.Count + 1);

            if (gui.ShowDialog() == DialogResult.OK)
            {
                tran.FK_MEEmrID = emr.MEEmrID;
                if (string.IsNullOrEmpty(tran.MEEmrTransferHistoryNo))
                    tran.MEEmrTransferHistoryNo = emr.MEEmrNo + "." + (_entity.MEEmrTranfersList.Count + 1);

                tran.MEEmrTransferHistoryCurrent = true;

                foreach (var item in _entity.MEEmrTranfersList)
                    item.MEEmrTransferHistoryCurrent = false;
                var tranmsg = $"khoa cũ {emr.FK_HRDepartmentID} - khoa mới {tran.FK_HRDepartmentToID}";
                emr.FK_HRDepartmentID = tran.FK_HRDepartmentToID;
                _entity.MEEmrTranfersList.AddObjectToList();
                _entity.MEEmrTranfersList.SaveItemObjects();
                emr.AAUpdatedUser = BOSApp.CurrentUser;
                _emrCtrl.UpdateObject(emr);
                HistoryEmr(emr, cstObjectHistoryActionChange, $"thay đổi thông tin bệnh án - chuyển khoa {tranmsg}.");
                _entity.MainObject = emr;

                // tat tat ca share trước đó
                foreach (var item in _entity.MEEmrShareList)
                    item.MEEmrShareHistoryActive = false;
                _entity.MEEmrShareList.SaveItemObjects();

                //share all neu can
                var dept = _departmentCtrl.GetObjectByID(tran.FK_HRDepartmentToID) as HRDepartmentsInfo;
                CreateShareAll(emr.MEEmrID, dept);

                CreateRemainShare(emr.MEEmrID, tran.FK_HRDepartmentFromID);

                this.ActionSave();
            }
            this.ActionCancel();
        }
        #endregion

        #region Camera

        internal void OpenCameraTool()
        {
            if (!IsAllowEditPositionWithFlashNotification(_richEditCtrl.Document.CaretPosition)) return;

            var gui = new Clas.Cap.guiCaptureCam();
            if (gui.ShowDialog() == DialogResult.OK)
            {
                if (gui.Images != null)
                    foreach (var img in gui.Images)
                    {
                        this.InsertImageToDocument(img.Img);
                    }
            }
            //gui.Dispose();
        }

        #endregion
        void PrintMgsLog(string title, string msg)
        {
            var str = "\r\n" + title + " - " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss:fff");
            str += "\r\n" + msg;
            str += "\r\n";

            if (_notificationTab)
            {
                if (_msgLogs.InvokeRequired)
                    _msgLogs.BeginInvoke((MethodInvoker)delegate { _msgLogs.Text += str; });
                else
                    _msgLogs.Text += str;
            }
            else if (_logConfig)
            {
                _msgLogsTemp += $"\r\n {str}";
            }
        }
        void PrintMgsLogJson(string title, object data)
        {
            if (data == null) return;
            Task.Run(() =>
            {
                try
                {
                    string str = JsonConvert.SerializeObject(data, Newtonsoft.Json.Formatting.Indented);
                    PrintMgsLog(title, str);
                }
                catch (Exception) {/*do nothing*/}
            });
        }
        #region History

        internal void GetDocumentDataHistory()
        {
            var emr = GetCurrentMainObject();// this._entity.MainObject as MEEmrsInfo;
            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName].Clone() as MEEmrDocumentsInfo;
            if (string.IsNullOrEmpty(document.MEEmrDocumentMongoID)) return;
            _entity.MEEmrDocumentDataHistoriesList.SetDefaultListAndRefreshGridControl();
            var template = this._templateCtrl.GetObjectByID(document.FK_METemplateID) as METemplatesInfo;
            var filters = new Dictionary<string, object>
            {
                { "RefId", new Tuple<MongoFilter, object>(MongoFilter.Eq, ObjectId.Parse(document.MEEmrDocumentMongoID)) }
            };
            var fields = new Dictionary<string, string>();
            var list = this._emrDocumentMng.GetList(filters, fields, template.METemplateNo + "_revisions");
            foreach (var item in list)
            {
                var doc = Mapper.Map<MEEmrDocumentsInfo>(item);
                if (doc.AAUpdatedDate.Year == 9999)
                    doc.AAUpdatedDate = doc.AACreatedDate;
                if (item.MEEmrDocumentContent != null)
                    doc.MEEmrDocumentJson = JToken.FromObject(item.MEEmrDocumentContent.ToDictionary()).ToString();
                //doc.MEEmrDocumentJson = item.MEEmrDocumentContent.ToJson(new MongoDB.Bson.IO.JsonWriterSettings() { OutputMode = MongoDB.Bson.IO.JsonOutputMode.Strict });
                _entity.MEEmrDocumentDataHistoriesList.Add(doc);
            }
            document.MEEmrDocumentJson = GetDocumentJsonData();
            document.AAUpdatedDate = DateTime.Now;
            document.MEEmrDocumentDesc = "Hiện tại trên tờ";
            document.AAUpdatedUser = string.Empty;
            _entity.MEEmrDocumentDataHistoriesList.Add(document);

            filters = new Dictionary<string, object>
            {
                { "_id", new Tuple<MongoFilter, object>(MongoFilter.Eq, ObjectId.Parse(document.MEEmrDocumentMongoID)) }
            };
            var current = this._emrDocumentMng.GetList(filters, fields, template.METemplateNo);
            if (current != null && current.Count > 0)
            {
                var inMongo = current.First();
                var doc = Mapper.Map<MEEmrDocumentsInfo>(inMongo);
                if (inMongo.MEEmrDocumentContent != null)
                    doc.MEEmrDocumentJson = JToken.FromObject(inMongo.MEEmrDocumentContent.ToDictionary()).ToString();
                doc.MEEmrDocumentDesc = "Hiện tại từ MongoDB";
                _entity.MEEmrDocumentDataHistoriesList.Add(doc);
            }
            _entity.MEEmrDocumentDataHistoriesList.GridControl.RefreshDataSource();
        }
        internal void ViewDocumentDataTree(MEEmrDocumentsInfo doc)
        {
            var tree = Controls["fld_trlDocumentDataView"] as DevExpress.XtraTreeList.TreeList;
            tree.ClearNodes();
            tree.BeginUnboundLoad();
            var data = JsonConvert.DeserializeObject(doc.MEEmrDocumentJson, new JsonSerializerSettings()
            {
                DateTimeZoneHandling = DateTimeZoneHandling.Local
            });
            var listEmrParams = new List<MEParamsInfo>();
            this.GetParamsTree(data as JToken, tree, null, listEmrParams);
            tree.EndUnboundLoad();
            tree.ExpandAll();
        }
        public MEParamsInfo GetParamByNo(string v, List<MEParamsInfo> listEmrParams)
        {
            var p = listEmrParams.Where(o => o.MEParamNo == v).FirstOrDefault();
            if (p != null) return p;
            p = AppMemCache.GetParamFromDictKeyNo(v);
            if (p != null) listEmrParams.Add(p);
            return p;
        }
        private void GetParamsTree(JToken data, TreeList list, TreeListNode parent, List<MEParamsInfo> listEmrParams)
        {
            if (data == null) return;
            if (data.Type == JTokenType.Object)
            {
                foreach (JToken child in data.Children())
                    this.GetParamsTree(child, list, parent, listEmrParams);
                return;
            }
            if (data.Type == JTokenType.Property)
            {
                string paramNo = data.GetType().GetProperty("Name").GetValue(data).ToString();
                var info = GetParamByNo(paramNo, listEmrParams);
                if (info != null)
                {
                    var prop = data as JProperty;
                    var value = prop != null ? prop.Value : null;
                    var str = value != null && value.HasValues ? string.Empty : value.ToString();
                    var obj = new object[] { info.MEParamName, str, prop };
                    var node = list.AppendNode(obj, parent);
                    this.GetParamsTree(value, list, node, listEmrParams);
                }
                else this.GetParamsTree(data.FirstOrDefault(), list, parent, listEmrParams);
                return;
            }
            if (data.Type == JTokenType.Array)
            {
                for (int i = 0; i < data.Count(); i++)
                {
                    var value = data[i];
                    var obj = new object[] { "#" + (i + 1), string.Empty, value };
                    if (value.Type != JTokenType.Object && value.Type != JTokenType.Array && value.Type != JTokenType.Property)
                        obj[1] = value.ToString();
                    var node = list.AppendNode(obj, parent);
                    this.GetParamsTree(value, list, node, listEmrParams);
                }
                return;
            }
        }
        #endregion

        #region AutoCorrect

        public void LoadUserAbbrevs()
        {
            _entity.MEEmrAbbrevList.Invalidate(BOSApp.CurrentUsersInfo.FK_HREmployeeID);
            this._userAbbrevs = new Dictionary<string, string>();
            foreach (var item in _entity.MEEmrAbbrevList)
            {
                if (!string.IsNullOrEmpty(item.MEEmrAbbrevNo) && !this._userAbbrevs.ContainsKey(item.MEEmrAbbrevNo))
                {
                    this._userAbbrevs.Add(item.MEEmrAbbrevNo, item.MEEmrAbbrevContent);
                }
            }
        }
        #endregion

        #region Guidance
        private void OpenGuidance()
        {
            string fileName = string.Format(@"{0}\{1}.docx", _documentPath, "guidance");
            if (File.Exists(fileName))
            {
                _richEditCtrl.LoadDocument(fileName, DocumentFormat.OpenXml);
                _richEditCtrl.ReadOnly = true;
            }
            else
                this._richEditCtrl.CreateNewDocument(false);

            this._pdfViewer.CloseDocument();
        }
        #endregion

        #region Share Emr
        internal void ChangeShareEmployeeOrDepartment(MEEmrShareHistoriesInfo row)
        {
            if (row.FK_HREmployeeID > 0)
            {
                var emp = _employeeCtrl.GetObjectByID(row.FK_HREmployeeID) as HREmployeesInfo;
                var dept = _departmentCtrl.GetObjectByID(emp.FK_HRDepartmentID) as HRDepartmentsInfo;
                if (dept != null)
                {
                    row.FK_HRDepartmentID = dept.HRDepartmentID;
                }
            }
        }

        internal void SelectEmployeeForShare()
        {
            if (Toolbar.IsNullOrNoneAction())
            {
                if (!IsCloseEmr())
                {
                    this.ActionEdit();
                }
            }

            var dataSource = ((DataSet)BOSApp.LookupTables[TableName.HREmployeesTableName]).Tables[0];
            var gui = new guiShareEmployeeSelection(dataSource, IsCloseOrWaitCloseEmr())
            {
                Module = this
            };
            gui.InitializeControls(gui.Controls);
            gui.StartPosition = FormStartPosition.CenterParent;
            if (gui.ShowDialog() == DialogResult.OK)
            {
                var emr = _entity.MainObject as MEEmrsInfo;
                if (gui.SelectedIDs.Count == 0)
                    this.ActionCancel();

                var emrShareHistoryMode = gui.emrShareHistoryMode;
                foreach (var id in gui.SelectedIDs)
                {
                    _entity.SetDefaultModuleObject(TableName.MEEmrShareHistoriesTableName);
                    var share = _entity.ModuleObjects[TableName.MEEmrShareHistoriesTableName] as MEEmrShareHistoriesInfo;
                    var emp = _employeeCtrl.GetObjectByID(id) as HREmployeesInfo;
                    var dept = _departmentCtrl.GetObjectByID(emp.FK_HRDepartmentID) as HRDepartmentsInfo;
                    if (dept != null)
                        share.FK_HRDepartmentID = dept.HRDepartmentID;

                    share.FK_MEEmrID = emr.MEEmrID;
                    share.FK_HREmployeeID = id;//chia se tat ca
                    share.MEEmrShareHistoryDate = DateTime.Now;
                    share.MEEmrShareHistoryFromDate = DateTime.Now.AddMinutes(-1);
                    share.MEEmrShareHistoryToDate = share.MEEmrShareHistoryFromDate.AddHours(BOSApp.CurrentCompanyInfo.CSCompanyEmrShareHours);
                    share.MEEmrShareHistoryActive = true;
                    share.FK_HREmployeeShareByID = BOSApp.CurrentEmployeesInfo.HREmployeeID;
                    share.MEEmrShareHistoryMode = emrShareHistoryMode;
                    _entity.MEEmrShareList.AddObjectToList();
                }
            }
        }
        internal void SelectDepartmentForShare()
        {
            if (Toolbar.IsNullOrNoneAction())
                this.ActionEdit();
            if (this.Toolbar.IsNullOrNoneAction()) return;

            var dataSource = ((DataSet)BOSApp.LookupTables[TableName.HRDepartmentsTableName]).Tables[0];
            var gui = new guiShareDepartmentSelection(dataSource, IsCloseOrWaitCloseEmr())
            {
                Module = this
            };
            gui.InitializeControls(gui.Controls);
            gui.StartPosition = FormStartPosition.CenterParent;
            if (gui.ShowDialog() == DialogResult.OK)
            {
                var emr = _entity.MainObject as MEEmrsInfo;
                if (gui.SelectedIDs.Count == 0)
                    this.ActionCancel();

                var emrShareHistoryMode = gui.emrShareHistoryMode;
                foreach (var id in gui.SelectedIDs)
                {
                    _entity.SetDefaultModuleObject(TableName.MEEmrShareHistoriesTableName);
                    var share = _entity.ModuleObjects[TableName.MEEmrShareHistoriesTableName] as MEEmrShareHistoriesInfo;
                    share.FK_MEEmrID = emr.MEEmrID;
                    share.FK_HRDepartmentID = id;
                    share.FK_HREmployeeID = 0;//chia se tat ca

                    share.MEEmrShareHistoryDate = DateTime.Now;
                    share.MEEmrShareHistoryFromDate = DateTime.Now.AddMinutes(-1);
                    share.MEEmrShareHistoryToDate = share.MEEmrShareHistoryFromDate.AddHours(BOSApp.CurrentCompanyInfo.CSCompanyEmrShareHours);
                    share.MEEmrShareHistoryActive = true;
                    share.FK_HREmployeeShareByID = BOSApp.CurrentEmployeesInfo.HREmployeeID;
                    share.MEEmrShareHistoryMode = emrShareHistoryMode;
                    _entity.MEEmrShareList.AddObjectToList();
                }
            }
        }
        internal void DeleteShareHistoryList()
        {
            _entity.MEEmrShareList.RemoveSelectedRowObjectFromList();
        }
        public bool CanShareEmr()
        {
            var emr = _entity.MainObject as MEEmrsInfo;
            if (emr.MEEmrStatus == EmrStatus.Closed.ToString())
            {
                if (!_entity.AllowSharingTheClosedEmr) return false;
                //benh an da dong thi Phải có đủ 3 điều kiện: [Vai trò = Quản trị] + [Quyền truy xuất = Toàn bộ] + [Phân quyền theo trạng thái Xem = Closed] thì mới có thể chia sẻ đc
                if (BOSApp.CurrentUserGroupInfo.ADUserGroupRole != UserGroupRole.admin.ToString()) return false;
                if (BOSApp.GetUserEmrViewPermission() != UserEmrView.ALL) return false;
                return true;
            }
            else //benh an dang mo
            {
                //thi chi admin moi duoc chia se
                if (BOSApp.CurrentUserGroupInfo.ADUserGroupRole == UserGroupRole.admin.ToString()) return true;
                //khoa quan ly moi duoc chia se
                if (emr.FK_HRDepartmentID == BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID) return true;
            }
            return false;
        }

        public string GetStatePermiQueryConditionStrSharedCase()
        {
            var viewStatePerms = BOSApp.GetObjectStatePermisions(TableName.MEEmrsTableName, ObjectStatePermissionAction.View);
            if (viewStatePerms == null) return string.Empty;
            var stateConds = new StringBuilder();
            var tbPrefix = string.Empty;
            foreach (var perm in viewStatePerms)
            {
                var states = string.Join("', '", perm.Value);
                // khi cho phep chia se benh an da dong thi moi them rule nay cho nguoi dung
                if (perm.Key == "MEEmrStatus")
                    states += "', 'Closed";

                stateConds.Append($" AND {tbPrefix}[{perm.Key}] IN ('{states}')");
                stateConds.AppendLine();
            }
            return stateConds.ToString();
        }
        #endregion

        #region Create IPC HIS
        private void CreateDocFromHis(Dictionary<string, object> response)
        {
            var icp = new Clas.Emr.Intergration.IpcHelper();
            string err = string.Empty;
            var msg = string.Empty;
            var emr = GetCurrentMainObject();// this._entity.MainObject as MEEmrsInfo;
            MEEmrsInfo emrOld = null;
            METemplatesInfo template = null;
            if (string.IsNullOrEmpty(response["MEEmrNo"].ToString().Trim()))
                err = "Tạo tờ bệnh án từ HIS. Mã bệnh án không được để trống";
            else if (string.IsNullOrEmpty(response["FK_METemplateNo"].ToString().Trim()))
                err = "Tạo tờ bệnh án từ HIS. Mã mẫu không được để trống";
            if (string.IsNullOrEmpty(err))
            {
                emrOld = _emrCtrl.GetObjectByNo(response["MEEmrNo"].ToString().Trim()) as MEEmrsInfo;
                if (emrOld == null)
                    err = "Tạo tờ bệnh án từ HIS. Mã bệnh án không tồn tại";
            }
            if (string.IsNullOrEmpty(err))
            {
                template = _templateCtrl.GetObjectByNo(response["FK_METemplateNo"].ToString().Trim()) as METemplatesInfo;
                if (template == null)
                    err = "Tạo tờ bệnh án từ HIS. Mẫu không tồn tại";
            }
            if (!string.IsNullOrEmpty(err))
            {
                PrintMgsLog("TAO_TO_BENH_AN_TU_HIS", err);
                msg = JsonConvert.SerializeObject(new
                {
                    msg = response["msg"] + "_FAILED",
                    error = err
                }, Newtonsoft.Json.Formatting.Indented);
                icp.SendMessage(response["msg"].ToString(), msg, "Json");
                return;
            }
            if (emr.MEEmrID != emrOld.MEEmrID)
            {
                this.ActionCancel();
                this.Invalidate(emrOld.MEEmrID);
            }
            _entity.MENewEmrDocument = new MEEmrDocumentsInfo()
            {
                FK_METemplateID = template.METemplateID,
                MEEmrDocumentFile = template.METemplateNo + DateTime.Now.ToString("_ddMMyyyy_hhmmssfff_") + Guid.NewGuid().ToString().Replace('-', '_'),
                MEEmrDocumentNo = template.METemplateNo,
                MEEmrDocumentCode = template.METemplateNo,
                MEEmrDocumentGuid = template.METemplateGuid,
                MEEmrDocumentDesc = response["MEEmrDocumentDesc"].ToString(),
                FK_HREmployeeCreatedID = BOSApp.CurrentEmployeesInfo.HREmployeeID
            };
            try
            {
                if (IsPatientProfileEditing())
                    AddNewPatientDocument(null);
                else
                    AddNewEmrDocument(null);

                PrintMgsLog("TAO_TO_BENH_AN_TU_HIS", emr.MEEmrNo);
                msg = JsonConvert.SerializeObject(new
                {
                    msg = response["msg"] + "_SUCCESSFULLY",
                    error = err,
                    data = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo
                }, Newtonsoft.Json.Formatting.Indented);
            }
            catch (Exception ex)
            {
                if (IsPatientProfileEditing())
                    this.InvalidatePatientDocuments(emr.MEEmrID);
                else
                    this.InvalidateEmrDocumentList(emr.MEEmrID);
                this.ClearDocumentSession();
                MessageBox.Show("Có lỗi khi thêm tờ bệnh án. Xin thử lại.", "CÓ LỖI KHI TẠO TỜ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                PrintMgsLog("LOI_TAO_TO_BENH_AN_TU_HIS", ex.ToString());
                msg = JsonConvert.SerializeObject(new
                {
                    msg = response["msg"] + "_FAILED",
                    error = ex.ToString(),
                }, Newtonsoft.Json.Formatting.Indented);
            }
            icp.SendMessage(response["msg"].ToString(), msg, "Json");
        }

        private void CreateEmrFromHis(Dictionary<string, object> response)
        {
            var icp = new Clas.Emr.Intergration.IpcHelper();
            var type = _emrTypeCtrl.GetObjectByNo(response["FK_MEEmrTypeNo"].ToString()) as MEEmrTypesInfo;
            string err = string.Empty;
            var msg = string.Empty;
            if (type == null)
                err = "Tạo mới bệnh án từ HIS. Không tìm thấy loại bệnh án";
            else if (string.IsNullOrEmpty(response["MEPatientNo"].ToString().Trim()))
                err = "Tạo mới bệnh án từ HIS. Mã bệnh nhân không được để trống";
            else if (string.IsNullOrEmpty(response["MEPatientName"].ToString().Trim()))
                err = "Tạo mới bệnh án từ HIS. Tên bệnh nhân không được để trống";
            else if (string.IsNullOrEmpty(response["MEEmrNo"].ToString().Trim()))
                err = "Tạo mới bệnh án từ HIS. Mã bệnh án không được để trống";
            else
            {
                var emrOld = _emrCtrl.GetObjectByNo(response["MEEmrNo"].ToString().Trim()) as MEEmrsInfo;
                if (emrOld != null)
                    err = "Tạo mới bệnh án từ HIS. Mã bệnh án đã tồn tại";
            }
            if (!string.IsNullOrEmpty(err))
            {
                PrintMgsLog("TAO_BENH_AN_TU_HIS", err);
                msg = JsonConvert.SerializeObject(new
                {
                    msg = response["msg"] + "_FAILED",
                    error = err
                }, Newtonsoft.Json.Formatting.Indented);
                icp.SendMessage(response["msg"].ToString(), msg, "Json");
                return;
            }
            this.ActionCancel();
            this.ActionNew();
            var emr = this._entity.MainObject as MEEmrsInfo;
            var patient = new MEPatientsInfo()
            {
                MEPatientNo = response["MEPatientNo"].ToString().Trim(),
                MEPatientName = response["MEPatientName"].ToString().Trim().ToUpper(),
                MEGender = response["MEGender"].ToString(),
                MEPatientBirthday = (DateTime)response["MEPatientBirthday"],
                MEPatientContactCellPhone = response["MEPatientContactCellPhone"].ToString(),
                MEPatientPermanentResidence = response["MEPatientPermanentResidence"].ToString(),
                MEPatientContactAddressLine1 = response["MEPatientContactAddressLine1"].ToString(),
                MEPatientIDCard = response["MEPatientIDCard"].ToString(),
                MEPatientIDCardDate = (DateTime)response["MEPatientIDCardDate"],
                MEPatientIDCardStateProvinces = response["MEPatientIDCardStateProvinces"].ToString(),
                MEMarital = response["MEMarital"].ToString(),
            };
            patient = CreatePatientOfExt(patient);
            _entity.InvalidateModuleObject(patient);
            //var lookup = (Controls[_fld_lkeFK_MEPatientID1Name] as BOSLookupEdit);
            //lookup.EditValue = patient.MEPatientID;
            //lookup.Refresh();
            //lookup.Text = patient.MEPatientName;

            emr.FK_MEPatientID = patient.MEPatientID;
            emr.MEEmrNo = response["MEEmrNo"].ToString().Trim();
            emr.FK_MEEmrTypeID = type.MEEmrTypeID;
            emr.MEEmrDesc = response["MEEmrDesc"].ToString();
            var tran = _entity.ModuleObjects[TableName.MEEmrTransferHistoriesTableName] as MEEmrTransferHistoriesInfo;
            tran.MEEmrTransferHistoryNo = patient.MEEmrNo;
            this.ActionSave();
            PrintMgsLog("TAO_BENH_AN_TU_HIS", emr.MEEmrNo);
            msg = JsonConvert.SerializeObject(new
            {
                msg = response["msg"] + "_SUCCESSFULLY",
                error = err,
                data = _emrCtrl.GetObjectByID(emr.MEEmrID) as MEEmrsInfo
            }, Newtonsoft.Json.Formatting.Indented);
            icp.SendMessage(response["msg"].ToString(), msg, "Json");
        }
        #endregion

        #region Release Handled Doc
        public override void AfterClose()
        {
            base.AfterClose();
            this._emrDocumentCtrl.ReleaseAllEditing(BOSApp.CurrentUsersInfo.FK_HREmployeeID, _macAddress);
        }
        #endregion

        internal void PreviewDocument(MEEmrDocumentsInfo doc)
        {
            var current = (_entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo);
            if (current.MEEmrDocumentID == doc.MEEmrDocumentID) return;
            if (doc.MEEmrDocumentFileExt == EmrDocumentFileExtention.docx.ToString())
            {
                var gui = new guiDocumentPreview(doc, this._emrDocumentHelper.ShareEmrPassword);
                gui.StartPosition = FormStartPosition.WindowsDefaultLocation;
                gui.Module = this;
                //gui.WindowState = FormWindowState.Maximized;
                gui.Show();
            }
            else if (doc.MEEmrDocumentFileExt == EmrDocumentFileExtention.pdf.ToString())
            {
                var gui = new guiPreviewPdf(doc);
                gui.StartPosition = FormStartPosition.WindowsDefaultLocation;
                gui.Module = this;
                //gui.WindowState = FormWindowState.Maximized;
                gui.Show();
            }
        }

        #region Search
        public override void ResetSearch()
        {
            //base.ResetSearch();
            InitSearchCriteria();
        }
        private void InitSearchCriteria()
        {
            var search = CurrentModuleEntity.SearchObject as MEEmrsInfo;
            search.MEEmrStatus = string.Empty;
            search.FK_MEEmrTypeID = 0;
            search.FK_MEPatientID = 0;
            search.MEEmrNo = string.Empty;
            CurrentModuleEntity.UpdateSearchObjectBindingSource();

            SetSearchParamValue(SearchScreen.CriteriaSection, "fld_ccbeFK_HRDepartmentID1", BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID);
            SetSearchParamValue(SearchScreen.CriteriaSection, "FK_MEPatientID", 0);
            SetSearchParamValue(SearchScreen.CriteriaSection, "fld_dteMEEmrCreatedDateSearchFrom", DateTime.Now.AddYears(-1));
            SetSearchParamValue(SearchScreen.CriteriaSection, "fld_dteMEEmrCreatedDateSearchTo", DateTime.Now.AddYears(1));
            SetSearchParamValue(SearchScreen.CriteriaSection, "MEEmrNo", null);
            SetSearchParamValue(SearchScreen.CriteriaSection, "MEEmrStatus", string.Empty);
            SetSearchParamValue(SearchScreen.CriteriaSection, "FK_MEEmrTypeID", 0);

            //uthv tim kiem nhanh
            SetSearchParamValue(ParentScreen.SearchQuickContainer, "fld_ccbeFK_HRDepartmentID", BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID);
            SetSearchParamValue(ParentScreen.SearchQuickContainer, "MEEmrCreatedDateFrom", DateTime.Now.Date);
            SetSearchParamValue(ParentScreen.SearchQuickContainer, "MEEmrCreatedDateTo", DateTime.Now.Date.AddDays(1).AddMilliseconds(-1));

        }
        public override void SearchAll(string query)
        {
            ResetSearch();
            Search();
        }
        public override void Search()
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                var stateConds = GetStatePermiQueryConditionStr(TableName.MEEmrsTableName, false);
                var view = BOSApp.GetUserEmrViewPermission();
                var shared = (bool)GetSearchParamValue(SearchScreen.CriteriaSection, "chkEmrCurrentShared");
                // 1523 Đã ký CA
                var mEEmrArchiveStatus = 0;
                if (BOSApp.CurrentUserGroupInfo.ADUserGroupRole == UserGroupRole.admin.ToString())
                {
                    mEEmrArchiveStatus = 1;
                }
                var excludedEmrTypes = GetExcludedEmrTypes();

                DataSet ds;
                if (!shared)
                {
                    var department = GetSearchParamValue(SearchScreen.CriteriaSection, "fld_ccbeFK_HRDepartmentID1").ToString();
                    //department = ", " + department + ", ";
                    department = department.Replace(", " + BOSApp.CurrentEmployeesInfo.FK_HRDepartmentRoomID + ", ", ", ");
                    var fromDate = ((DateTime)GetSearchParamValue(SearchScreen.CriteriaSection, "fld_dteMEEmrCreatedDateSearchFrom")).Date;
                    var toDate = ((DateTime)GetSearchParamValue(SearchScreen.CriteriaSection, "fld_dteMEEmrCreatedDateSearchTo")).Date.AddDays(1).AddMilliseconds(-1);

                    var patientGroup = string.Empty;
                    DateTime? fromDateOut = null;
                    DateTime? toDateOut = null;

                    object[] paramValues = new object[]
                    {
                         GetSearchParamValue(SearchScreen.CriteriaSection ,"MEEmrNo"),
                         GetSearchParamValue(SearchScreen.CriteriaSection ,"MEEmrStatus"),
                         fromDate,
                         toDate,
                         GetSearchParamValue(SearchScreen.CriteriaSection ,"FK_MEPatientID"),
                         GetSearchParamValue(SearchScreen.CriteriaSection ,"FK_MEEmrTypeID"),
                         patientGroup,
                         fromDateOut,
                         toDateOut,
                         department,
                         BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID,
                         BOSApp.CurrentEmployeesInfo.HREmployeeID,
                         mEEmrArchiveStatus,
                         view,
                         stateConds
                    };
                    ds = _emrCtrl.Search(paramValues);
                }
                else
                    ds = _emrCtrl.GetAllActiveShared(BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID, mEEmrArchiveStatus, stateConds, excludedEmrTypes);

                Toolbar.SetToolbar(ds);
                InvalidateAfterSearch(null, string.Empty);
                Cursor.Current = Cursors.Default;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }
        private string GetExcludedEmrTypes()
        {
            var excludedEmrTypes = string.Empty;
            if (_entity.HideTheTempEmrOnQuickSearch)
            {
                var emrTypes = _emrTypeCtrl.GetListBusinessObjects<MEEmrTypesInfo>((DataSet)BOSApp.LookupTables[TableName.MEEmrTypesTableName]);
                excludedEmrTypes = string.Join(", ", emrTypes.Where(t => t.MEEmrTypeIsTmp).Select(d => d.MEEmrTypeID).ToArray());
            }
            return excludedEmrTypes;
        }
        public override void QuickSearch()
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                var stateConds = base.GetStatePermiQueryConditionStr(TableName.MEEmrsTableName, false);
                var patientOnly = (bool)GetSearchParamValue(ParentScreen.SearchQuickContainer, "chkSearchByPatient");
                var emrNoOnly = (bool)GetSearchParamValue(ParentScreen.SearchQuickContainer, "chkSearchByEmrCode");
                var shareOnly = (bool)GetSearchParamValue(ParentScreen.SearchQuickContainer, "chkSearchByShare");
                var assignOnly = (bool)GetSearchParamValue(ParentScreen.SearchQuickContainer, "chkUserAssigned");
                var userName = string.Empty;
                // 1523 Đã ký CA
                var isIncludeArchiveStatus = 0;
                if (BOSApp.CurrentUserGroupInfo.ADUserGroupRole == UserGroupRole.admin.ToString())
                {
                    isIncludeArchiveStatus = 1;
                }
                var excludedEmrTypes = GetExcludedEmrTypes();
                var ds = new DataSet();
                var view = BOSApp.GetUserEmrViewPermission();
                if (patientOnly)
                {
                    var id = (int)GetSearchParamValue(ParentScreen.SearchQuickContainer, "fld_lkeFK_MEPatientID");
                    if (id <= 0)
                    {
                        MessageBox.Show("Vui lòng nhập chọn một bệnh nhân", "CHƯA CHỌN BỆNH NHÂN", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }
                    ds = _emrCtrl.QuickSearchWithOneCriteria(null, id,
                        BOSApp.CurrentEmployeesInfo.HREmployeeID,
                        BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID,
                        null,
                        isIncludeArchiveStatus,
                        view,
                        stateConds.ToString(),
                        excludedEmrTypes
                        );
                }
                else if (emrNoOnly)
                {
                    var no = ((string)GetSearchParamValue(ParentScreen.SearchQuickContainer, "fld_txtMEEmrNo"))?.Trim();
                    if (string.IsNullOrEmpty(no))
                    {
                        MessageBox.Show("Vui lòng nhập mã bệnh án", "CHƯA NHẬP MÃ BỆNH ÁN", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }
                    ds = _emrCtrl.QuickSearchWithOneCriteria(no, null,
                        BOSApp.CurrentEmployeesInfo.HREmployeeID,
                        BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID,
                        null,
                        isIncludeArchiveStatus,
                        view,
                        stateConds.ToString(),
                        excludedEmrTypes
                        );
                }
                else if (shareOnly)
                {
                    var index = GetSearchParamValue(ParentScreen.SearchQuickContainer, "fld_cmbChooseShare");
                    //tim benh an dang duoc chi se boi khoa toi
                    if (index.ToString() == "Khoa tôi chia sẻ")
                        ds = _emrCtrl.GetAllActiveShared(BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID, isIncludeArchiveStatus, stateConds.ToString(), excludedEmrTypes);
                    //tim benh an dc chi se cho toi hay khoa toi khong bao gom khoa mo
                    else
                    {
                        var deps = _departmentCtrl.GetListBusinessObjects<HRDepartmentsInfo>((DataSet)BOSApp.LookupTables[TableName.HRDepartmentsTableName]);
                        //lay tat ca cac khoa khac ngoai tru khoa chia se va khoa cua toi
                        var depList = deps.Where(d => !d.HRDepartmentEmrShared && d.HRDepartmentID != BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID).Select(d => d.HRDepartmentID).ToArray();
                        var departments = string.Join(", ", depList);

                        // khi cho phep chia se benh an da dong thi moi them rule nay cho nguoi dung
                        var openDeptStr = string.Empty;
                        if (_entity.AllowSharingTheClosedEmr)
                        {
                            stateConds = this.GetStatePermiQueryConditionStrSharedCase();
                            var openDeptList = deps.Where(d => d.HRDepartmentEmrShared && d.HRDepartmentID != BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID).Select(d => d.HRDepartmentID).ToList();
                            openDeptStr = string.Join(", ", openDeptList);
                        }

                        // User được phép thấy bệnh án Chờ Đóng đã chia sẻ cho mình
                        if (!string.IsNullOrEmpty(stateConds))
                        {
                            if (stateConds.Contains(')'))
                            {
                                var stateCondsArr = stateConds.Split(')');
                                var stateCondsArr0 = stateCondsArr[0] + $", '{EmrStatus.WaitClose}'";
                                stateConds = stateCondsArr0 + ")" + stateCondsArr[1];
                            }
                        }

                        object[] paramValues = new object[]
                        {
                            new DateTime(1900,1,1), new DateTime(2999,1,1),
                            BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID,
                            departments,
                            openDeptStr,
                            BOSApp.CurrentEmployeesInfo.HREmployeeID,
                            isIncludeArchiveStatus,
                            view,
                            stateConds.ToString(),
                            excludedEmrTypes
                        };
                        ds = _emrCtrl.QuickSharedSearch(paramValues);
                    }
                }
                else if (assignOnly)
                {
                    userName = BOSApp.CurrentUser;
                    object[] paramValues = new object[]
                    {
                        new DateTime(1900,1,1),
                        new DateTime(2999,1,1),
                        userName
                    };
                    ds = _emrCtrl.QuickUserAssignedSearch(paramValues);
                }
                else
                {
                    #region VIEW EMR SHARED MODE
                    // [Mode xem ba] mặc định ở LK: false | PSHN: true (SEE EMR SHARED)
                    // NO SEE EMR SHARED WHEN SETTING = FALSE && chkDepartmentShared false.
                    var modeShare = _entity.GetConfigModeShareEmrViewAll();
                    if (!modeShare)
                    {
                        // LK
                        try
                        {
                            var isShareDepartment = (bool)GetSearchParamValue(ParentScreen.SearchQuickContainer, "chkDepartmentShared");
                            if (isShareDepartment)
                            {
                                modeShare = true;
                            }
                        }
                        catch (Exception ex) { }
                    }
                    #endregion

                    var departments = GetSearchParamValue(ParentScreen.SearchQuickContainer, "fld_ccbeFK_HRDepartmentID").ToString();
                    //departments = ", " + departments + ", ";
                    var fromDate = ((DateTime)GetSearchParamValue(ParentScreen.SearchQuickContainer, "MEEmrCreatedDateFrom")).Date;
                    if (fromDate.Year == 9999)
                        fromDate = DateTime.MinValue;

                    var toDate = ((DateTime)GetSearchParamValue(ParentScreen.SearchQuickContainer, "MEEmrCreatedDateTo")).Date;
                    if (toDate.Year != 9999)
                        toDate = toDate.AddDays(1).AddMilliseconds(-1);

                    object[] paramValues = new object[]
                    {
                         fromDate,
                         toDate,
                         BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID,
                         departments,
                         BOSApp.CurrentEmployeesInfo.HREmployeeID,
                         isIncludeArchiveStatus,
                         view,
                         stateConds.ToString(),
                         excludedEmrTypes,
                         modeShare
                    };
                    ds = _emrCtrl.QuickSearch(paramValues);
                }
                InvalidateSearchResult(ds);
                Cursor.Current = Cursors.Default;
            }
            catch (Exception e)
            {
                if (e is System.Data.SqlClient.SqlException)
                {
                    MessageBox.Show("Có lỗi trong quá trình lấy dữ liệu. Vui lòng thử lại.");
                    return;
                }

                MessageBox.Show(e.ToString());
            }
        }
        private void InvalidateSearchResult(DataSet ds)
        {
            var preIndex = Toolbar.CurrentIndex;
            Toolbar.SetToolbar(ds);
            InvalidateAfterSearch(null, string.Empty);
            var searchResultControl = Controls.Values.Cast<Control>().FirstOrDefault(ctrl => (string)ctrl.Tag == BOSScreen.SearchResultControl);
            if (searchResultControl != null)
            {
                var control = searchResultControl as BOSSearchResultsGridControl;
                if (control != null)
                    control.InvalidateLookupEditColumns();
            }
            var currEmrID = (_entity.MainObject as MEEmrsInfo)?.MEEmrID;
            if (preIndex == 0 && Toolbar.CurrentIndex == 0 && Toolbar.CurrentObjectID != currEmrID)
            {
                // force invalidate in case current index not change
                // invoke required due to cross thread exception
                InvokeInvalidate();
            }
        }
        delegate void InvokeRequiredInvalidate();
        private void InvokeInvalidate()
        {
            if (_notificationTab && this._msgLogs.InvokeRequired)
            {
                InvokeRequiredInvalidate d = new InvokeRequiredInvalidate(InvokeInvalidate);
                _msgLogs.Invoke(d);
            }
            else
            {
                Toolbar.Invalidate();
            }
        }
        private void SetSearchParamValue(PanelControl panel, string strColumnName, object value)
        {
            foreach (Control ctrl in panel.Controls)
                if (ctrl.Tag != null && ctrl.Tag.ToString() == BOSScreen.SearchControl)
                {
                    if (strColumnName == _dbUtil.GetPropertyStringValue(ctrl, BOSScreen.cstDataMemberPropertyName)
                        || strColumnName == ctrl.Name)
                    {
                        if (ctrl is CheckedComboBoxEdit)
                            ((CheckedComboBoxEdit)ctrl).SetEditValue(value);
                        //else if (ctrl is MultiColCheckedComboBoxEdit)
                        //    ((MultiColCheckedComboBoxEdit)ctrl).EditValue = (value);
                        else
                            ((BaseEdit)ctrl).EditValue = value;
                        return;
                    }
                }
                else if (ctrl.Controls.Count > 0)
                {
                    foreach (Control c in ctrl.Controls)
                    {
                        if (c.Tag != null && c.Tag.ToString() == BOSScreen.SearchControl)
                            if (strColumnName == _dbUtil.GetPropertyStringValue(c, BOSScreen.cstDataMemberPropertyName)
                                || strColumnName == ctrl.Name)
                            {
                                if (ctrl is CheckedComboBoxEdit)
                                    ((CheckedComboBoxEdit)ctrl).SetEditValue(value);
                                //else if (ctrl is MultiColCheckedComboBoxEdit)
                                //    ((MultiColCheckedComboBoxEdit)ctrl).EditValue = (value);
                                else
                                    ((BaseEdit)c).EditValue = value;
                                return;
                            }
                    }
                }
            return;
        }
        private void SetSearchParamValueInvoke(PanelControl panel, string strColumnName, object value)
        {
            foreach (Control ctrl in panel.Controls)
                if (ctrl.Tag != null && ctrl.Tag.ToString() == BOSScreen.SearchControl)
                {
                    if (strColumnName == _dbUtil.GetPropertyStringValue(ctrl, BOSScreen.cstDataMemberPropertyName)
                        || strColumnName == ctrl.Name)
                    {
                        if (ctrl is CheckedComboBoxEdit)
                            ((CheckedComboBoxEdit)ctrl).Invoke(new Action(() =>
                            {
                                ((CheckedComboBoxEdit)ctrl).SetEditValue(value);
                            }
                            ));
                        else if (ctrl is BOSDateEdit)
                        {
                            ((BOSDateEdit)ctrl).Invoke(new Action(() =>
                            {
                                ((BOSDateEdit)ctrl).EditValue = (value);
                            }
                            ));
                        }
                        //else if (ctrl is MultiColCheckedComboBoxEdit)
                        //    ((MultiColCheckedComboBoxEdit)ctrl).EditValue = (value);
                        else
                            ((BaseEdit)ctrl).Invoke(new Action(() =>
                            {
                                ((BaseEdit)ctrl).EditValue = value;
                            }
                            ));
                        return;
                    }
                }
                else if (ctrl.Controls.Count > 0)
                {
                    foreach (Control c in ctrl.Controls)
                    {
                        if (c.Tag != null && c.Tag.ToString() == BOSScreen.SearchControl)
                            if (strColumnName == _dbUtil.GetPropertyStringValue(c, BOSScreen.cstDataMemberPropertyName)
                                || strColumnName == ctrl.Name)
                            {
                                if (ctrl is CheckedComboBoxEdit)
                                    ((CheckedComboBoxEdit)ctrl).Invoke(new Action(() =>
                                    {
                                        ((CheckedComboBoxEdit)ctrl).SetEditValue(value);
                                    }
                                    ));
                                //else if (ctrl is MultiColCheckedComboBoxEdit)
                                //    ((MultiColCheckedComboBoxEdit)ctrl).EditValue = (value);
                                else
                                    ((BaseEdit)ctrl).Invoke(new Action(() =>
                                    {
                                        ((BaseEdit)ctrl).EditValue = value;
                                    }
                                    ));
                                return;
                            }
                    }
                }
            return;
        }
        private object GetSearchParamValue(PanelControl panel, string strColumnName)
        {
            foreach (Control ctrl in panel.Controls)
                if (ctrl.Tag != null && ctrl.Tag.ToString() == BOSScreen.SearchControl)
                {
                    if (strColumnName == _dbUtil.GetPropertyStringValue(ctrl, BOSScreen.cstDataMemberPropertyName)
                        || strColumnName == ctrl.Name)
                        return ((DevExpress.XtraEditors.BaseEdit)ctrl).EditValue;
                }
                else if (ctrl.Controls.Count > 0)
                {
                    foreach (Control c in ctrl.Controls)
                    {
                        if (c.Tag != null && c.Tag.ToString() == BOSScreen.SearchControl)
                            if (strColumnName == _dbUtil.GetPropertyStringValue(c, BOSScreen.cstDataMemberPropertyName)
                                || strColumnName == ctrl.Name)
                                return ((DevExpress.XtraEditors.BaseEdit)ctrl).EditValue;
                    }
                }
            return null;
        }

        #endregion

        #region Benh an tam
        public void CreateEmrTemp()
        {
            var gui = new guiCreateNewPatient();
            gui.Module = this;
            _entity.SetDefaultModuleObject(TableName.MEPatientsTableName);
            gui.InitializeControls(gui.Controls);
            gui.StartPosition = FormStartPosition.CenterParent;
            var patient = _entity.ModuleObjects[TableName.MEPatientsTableName] as MEPatientsInfo;
            patient.MEPatientNo = "***NEW***";
            patient.MEGender = "Male";
            patient.MEPatientType = CustomerType.Temp.ToString();
            if (gui.ShowDialog() == DialogResult.OK)
            {
                this.ActionNew();
                patient.MEPatientNo = BOSApp.GetMainObjectNo(ModuleName.MEPatient);
                patient.MEPatientName = patient.MEPatientName.Trim().ToUpper();
                patient.AACreatedUser = BOSApp.CurrentUser;
                int patientId = this._patientCtrl.CreateObject(patient);
                BOSApp.UpdateObjectNumbering(ModuleName.MEPatient);
                int customerId = this._patientCtrl.CreateCustomerFromPatientInfo(patient, BOSApp.CurrentBranchInfo.BRBranchID);
                patient = this._patientCtrl.GetObjectByID(patientId) as MEPatientsInfo;
                var emr = (this._entity.MainObject as MEEmrsInfo);
                _entity.InvalidateModuleObject(patient);
                //var lookup = (Controls[_fld_lkeFK_MEPatientID1Name] as BOSLookupEdit);
                //lookup.InvalidateDataSourceToLookupEdit();
                var emrType = this._emrTypeCtrl.GetTmpByDeparment(BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID);
                if (emrType != null)
                {
                    (Controls["fld_lkeFK_MEEmrTypeID"] as BOSLookupEdit).EditValue = emrType.MEEmrTypeID;
                    emr.FK_MEEmrTypeID = emrType.MEEmrTypeID;
                    Controls["fld_medMEEmrDesc1"].Focus();
                }
                else
                    Controls["fld_lkeFK_MEEmrTypeID"].Focus();
                emr.FK_MEPatientID = patient.MEPatientID;
                emr.MEEmrCreatedDate = DateTime.Now;
                emr.FK_HREmployeeCreatedID = BOSApp.CurrentEmployeesInfo.HREmployeeID;
            }
        }
        #endregion

        #region Merge Emr
        public void MergeEmr()
        {
            if (IsEmrReadOnly(true))
            {
                MessageBox.Show("Bệnh án đã đóng. Không thể thực hiện thao tác này.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var gui = new guiSelectEmr
            {
                Module = this,
                StartPosition = FormStartPosition.CenterParent
            };
            if (gui.ShowDialog() == DialogResult.OK)
            {

                int id = gui.SelectedEmrID;
                var source = _emrCtrl.GetObjectByID(id) as MEEmrsInfo;
                var emr = _entity.MainObject as MEEmrsInfo;

                MergeEmrAction(source, emr);
            }
        }

        public void MergeEmrAction(MEEmrsInfo source, MEEmrsInfo emr)
        {
            var msg = string.Empty;
            if (source.MEEmrID == emr.MEEmrID)
            {
                msg = "Bệnh án trùng nhau không thể trộn.";
                MergeEmrLog(source, emr, EmrMergeHistoryStatus.MergeError.ToString(), msg);
                MessageBox.Show(msg, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            //if (source.FK_HRDepartmentID != emr.FK_HRDepartmentID)
            //{
            //    MessageBox.Show("Bệnh án khác khoa không thể trộn", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}
            if (emr.AACreatedDate < source.AACreatedDate)
            {
                msg = "Bệnh án đích có NGÀY TẠO NHỎ HƠN Bệnh án nguồn. Có thể bạn đã chọn nhầm bệnh án ĐÍCH.";
                if (MessageBox.Show($"{msg}. Bạn vẫn tiếp tục muốn trộn bệnh án?", "Thông báo",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.Cancel)
                {
                    MergeEmrLog(source, emr, EmrMergeHistoryStatus.MergeError.ToString(), $"Không xác nhận {msg}");
                    return;
                }
            }

            var type = _emrTypeCtrl.GetObjectByID(source.FK_MEEmrTypeID) as MEEmrTypesInfo;
            if (type.MEEmrTypeIsTmp)
            {
                if (MessageBox.Show("Bạn đang trộn dữ liệu vào một BỆNH ÁN TẠM. Vẫn thực hiện?", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.Cancel)
                {
                    MergeEmrLog(source, emr, EmrMergeHistoryStatus.MergeError.ToString(), "Không xác nhận trộn dữ liệu vào một BỆNH ÁN TẠM.");
                    return;
                }
            }

            var patient = _patientCtrl.GetObjectByID(source.FK_MEPatientID) as MEPatientsInfo;
            if (patient.MEPatientType == CustomerType.Temp.ToString())
                if (MessageBox.Show("Bạn đang trộn bệnh án của BỆNH NHÂN TẠM " + patient.MEPatientNo +
                    "\n Sau khi trộn. Bệnh nhân này sẽ bị XÓA. Vẫn thực hiện?", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.Cancel)
                {
                    MergeEmrLog(source, emr, EmrMergeHistoryStatus.MergeError.ToString(), "Không xác nhận trộn bệnh án của BỆNH NHÂN TẠM.");
                    return;
                }

            var docs = _emrDocumentCtrl.GetListBusinessObjects<MEEmrDocumentsInfo>(_emrDocumentCtrl.GetAllDataByForeignColumn("FK_MEEmrID", source.MEEmrID));
            foreach (var doc in docs)
            {
                if (doc.FK_EditingUserID > 0 && doc.FK_EditingUserID != BOSApp.CurrentEmployeesInfo.HREmployeeID)
                {
                    msg = $"Tờ bệnh án: {doc.MEEmrDocumentGroup} đang được soạn bởi người khác ({doc.MEEmrDocumentHoldMachineIp}:{doc.MEEmrDocumentHoldMachineMac}) từ lúc {doc.MEEmrDocumentHoldFrom.ToString("HH:mm dd/MM/yyyy")}";
                    MergeEmrLog(source, emr, EmrMergeHistoryStatus.MergeError.ToString(), msg);
                    ShowDocumentEditingUser(doc);
                    return;
                }
            }

            var docsEmr = _emrDocumentCtrl.GetListBusinessObjects<MEEmrDocumentsInfo>(_emrDocumentCtrl.GetAllDataByForeignColumn("FK_MEEmrID", emr.MEEmrID));
            foreach (var doc in docsEmr)
            {
                if (doc.FK_EditingUserID > 0 && doc.MEEmrDocumentHoldMachineMac != _macAddress)
                {
                    msg = $"Tờ bệnh án: {doc.MEEmrDocumentGroup} đang được soạn bởi người khác ({doc.MEEmrDocumentHoldMachineIp}:{doc.MEEmrDocumentHoldMachineMac}) từ lúc {doc.MEEmrDocumentHoldFrom.ToString("HH:mm dd/MM/yyyy")}";
                    MergeEmrLog(source, emr, EmrMergeHistoryStatus.MergeError.ToString(), msg);
                    ShowDocumentEditingUser(doc);
                    return;
                }
            }

            BOSProgressBar.Start("Đang trộn bệnh án");
            try
            {
                var isActionLocal = false;
                if (_apiEmr != null)
                {
                    var actionUri = BOSApp.GetSystemConfigValue(SysCfgConsts.EMR_API_ENDPOINT, SysCfgConsts.EMR_MERGE);
                    if (!string.IsNullOrEmpty(actionUri))
                    {
                        PrintMgsLog("BAT-DAU-GOI-API", actionUri);
                        var body = new
                        {
                            MEEmrSourceNo = source.MEEmrNo,
                            MEEmrTargetNo = emr.MEEmrNo,
                            AACreatedUser = BOSApp.CurrentUser
                        };
                        var response = _apiEmr.Post<Emr.Base.Models.Abp.AjaxResponse, JValue>(actionUri, null, body);
                        PrintMgsLog("KET-THUC-GOI-API", actionUri);
                        if (response != null)
                        {
                            if (!response.Success)
                            {
                                if (response.Error != null)
                                {
                                    if (response.Error.Code != 1)
                                    {
                                        isActionLocal = true;
                                    }
                                    else
                                    {
                                        msg = response.Error.Message;
                                        MergeEmrLog(source, emr, EmrMergeHistoryStatus.MergeError.ToString(), msg);
                                        MessageBox.Show(msg, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                        return;
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (_notificationTab)
                            {
                                _msgLogs.Text = ("Không có dữ liệu trả về từ api. Liên hệ quản trị viên để biết thêm chi tiết");
                            }
                            else if (_logConfig)
                            {
                                _msgLogsTemp += $"\r\n {("Không có dữ liệu trả về từ api. Liên hệ quản trị viên để biết thêm chi tiết")}";
                            }
                            isActionLocal = true;
                        }
                    }
                    else
                    {
                        if (_notificationTab)
                        {
                            _msgLogs.Text = ("Không cấu hình trộn bệnh án bằng API");
                        }
                        else if (_logConfig)
                        {
                            _msgLogsTemp += $"\r\n {("Không cấu hình trộn bệnh án bằng API")}";
                        }
                        isActionLocal = true;
                    }
                }
                else
                {
                    if (_notificationTab)
                    {
                        _msgLogs.Text = ("Không kết nối được đến máy chủ EMR API");
                    }
                    else if (_logConfig)
                    {
                        _msgLogsTemp += $"\r\n {("Không kết nối được đến máy chủ EMR API")}";
                    }
                    isActionLocal = true;
                }

                if (isActionLocal)
                {
                    // UI: Lỗi api => auto use local
                    //if (MessageBox.Show("Tiếp tục thực hiện trộn bệnh án trên máy? \n Trộn bệnh án thất bại trên máy chủ EMR API.", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.Cancel)
                    //    return;

                    var emrTypeTemplates = AppMemCache.GetEmrTypeTemplatesFromDict(emr.FK_MEEmrTypeID);
                    _ftpFileMng.CreateDirectory($"/Emr/{emr.MEEmrID}");

                    #region History Merge
                    var mergeHistoryDetails = new List<MEEmrMergeHistoryDetailsInfo>();
                    #endregion

                    foreach (var doc in docs)
                    {
                        var typeTmpl = emrTypeTemplates.Where(o => o.FK_METemplateID == doc.FK_METemplateID).FirstOrDefault();
                        if (typeTmpl != null)
                        {
                            var order = AppMemCache.GetTemplateIndexById(typeTmpl.FK_METemplateIndexID);
                            if (order != null)
                            {
                                doc.MEEmrDocumentOrder = order.METemplateIndexOrder;
                                doc.MEEmrDocumentGroup = order.METemplateIndexName;
                            }
                        }
                        doc.FK_MEEmrID = emr.MEEmrID;

                        var from = $"/Emr/{source.MEEmrID}/{doc.MEEmrDocumentFile}.{doc.MEEmrDocumentFileExt}";
                        var to = $"/Emr/{emr.MEEmrID}/{doc.MEEmrDocumentFile}.{doc.MEEmrDocumentFileExt}";
                        var localPath = string.Format(@"{0}\Emr\{1}\{2}.{3}", _documentPath, emr.MEEmrID, doc.MEEmrDocumentFile, doc.MEEmrDocumentFileExt);
                        _ftpFileMng.CopyFile(from, to, localPath);

                        if (doc.MEEmrDocumentStatus == EmrDocumentStatus.Closed.ToString())
                        {
                            from = $"/Emr/{source.MEEmrID}/{doc.MEEmrDocumentFile}.{EmrDocumentFileExtention.docx.ToString()}";
                            to = $"/Emr/{emr.MEEmrID}/{doc.MEEmrDocumentFile}.{EmrDocumentFileExtention.docx.ToString()}";
                            localPath = string.Format(@"{0}\Emr\{1}\{2}.{3}", _documentPath, emr.MEEmrID, doc.MEEmrDocumentFile, EmrDocumentFileExtention.docx.ToString());
                            _ftpFileMng.CopyFile(from, to, localPath);
                        }

                        doc.AAUpdatedUser = BOSApp.CurrentUser;
                        _emrDocumentCtrl.UpdateObject(doc);

                        var documentCode = getDocumentNo(doc);

                        UpdateMongoDocument(emr, doc, new List<string>() { "MEEmrDocumentContent" });

                        #region History Merge
                        mergeHistoryDetails.Add(new MEEmrMergeHistoryDetailsInfo()
                        {
                            FK_MEEmrDocumentID = doc.MEEmrDocumentID
                        });
                        #endregion
                    }

                    // Signed files
                    var fromPath = $"/Emr/Signed/{source.MEEmrID}/";
                    var toPath = $"/Emr/Signed/{emr.MEEmrID}/";
                    var localDir = string.Format(@"{0}\Emr\Signed\{1}\", _documentPath, emr.MEEmrID);
                    _ftpFileMng.CopyAllFiles(fromPath, toPath, localDir);

                    // Partials files
                    fromPath = $"/Emr/Partials/{source.MEEmrID}/";
                    toPath = $"/Emr/Partials/{emr.MEEmrID}/";
                    localDir = string.Format(@"{0}\Emr\Partials\{1}\", _documentPath, emr.MEEmrID);
                    _ftpFileMng.CopyAllFiles(fromPath, toPath, localDir);

                    var transfers = _transferCtrl.GetByEmrId(source.MEEmrID).ToList();
                    foreach (var tran in transfers)
                    {
                        tran.FK_MEEmrID = emr.MEEmrID;
                        tran.MEEmrTransferHistoryNo = source.MEEmrNo + "/" + tran.MEEmrTransferHistoryID;
                        tran.MEEmrTransferHistoryRemark = tran.MEEmrTransferHistoryRemark + $" > Trộn {source.MEEmrNo} vào {emr.MEEmrNo}";
                        tran.MEEmrTransferHistoryCurrent = false;
                        tran.AAUpdatedUser = BOSApp.CurrentUser;
                        _transferCtrl.UpdateObject(tran);
                    }

                    HistoryEmr(emr, "Change", $"Trộn các tờ từ bệnh án id {source.MEEmrID} vào bệnh án.");

                    HistoryEmr(source, "Delete", $"Xoá bệnh án sau khi trộn vào bệnh án id {emr.MEEmrID}.");

                    if (patient.MEPatientType == CustomerType.Temp.ToString())
                        _patientCtrl.DeleteObject(patient.MEPatientID);
                    _emrCtrl.DeleteObject(source.MEEmrID);

                    #region History Merge
                    var emrMergeHistoryDetailCtrl = new MEEmrMergeHistoryDetailsController();
                    // Insert history
                    var mergeHistoryId = MergeEmrLog(source, emr, EmrMergeHistoryStatus.Merge.ToString(), string.Empty);
                    // Insert detail
                    foreach (var mergeHistoryDetail in mergeHistoryDetails)
                    {
                        mergeHistoryDetail.FK_MEEmrMergeHistoryID = mergeHistoryId;
                        mergeHistoryDetail.AACreatedUser = BOSApp.CurrentUser;
                        emrMergeHistoryDetailCtrl.CreateObject(mergeHistoryDetail);
                    }
                    #endregion
                }
                this.ActionCancel();
                this.Invalidate(emr.MEEmrID);

                Control searchResultControl = Controls.Values.Cast<Control>().FirstOrDefault(ctrl => (string)ctrl.Tag == BOSScreen.SearchResultControl);
                if (searchResultControl != null)
                {
                    if (searchResultControl is BOSSearchResultsGridControl control)
                    {
                        var gridView = (control.MainView as GridView);
                        int rowHandle = gridView.LocateByValue("MEEmrID", source.MEEmrID);
                        if (rowHandle != DevExpress.XtraGrid.GridControl.InvalidRowHandle)
                            gridView.DeleteRow(rowHandle);
                    }
                }
                HistoryEmr(source, cstObjectHistoryActionMerge, $"Trộn vào bệnh án {emr.MEEmrNo}");
                MessageBox.Show($"Trộn thành công {source.MEEmrNo} vào {emr.MEEmrNo}", "Trộn bệnh án thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MergeEmrLog(source, emr, EmrMergeHistoryStatus.MergeError.ToString(), ex.ToString());
                MessageBox.Show("Có lỗi xảy ra khi trộn bệnh án. \n" + ex.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                BOSProgressBar.Close();
            }
        }

        private int MergeEmrLog(MEEmrsInfo source, MEEmrsInfo emr, string status, string description)
        {
            var emrMergeHistoryCtrl = new MEEmrMergeHistoriesController();
            var mergeHistoryId = emrMergeHistoryCtrl.CreateObject(new MEEmrMergeHistoriesInfo()
            {
                AACreatedUser = BOSApp.CurrentUser,
                FK_MEEmrFromID = source.MEEmrID,
                MEEmrFromNo = source.MEEmrNo,
                FK_MEEmrToID = emr.MEEmrID,
                MEEmrToNo = emr.MEEmrNo,
                MEEmrMergeHistoryDate = DateTime.Now,
                FK_HRDepartmentFromID = source.FK_HRDepartmentID,
                FK_HRDepartmentToID = BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID,
                MEEmrMergeHistoryStatus = status,
                MEEmrMergeHistoryRemark = description
            });

            return mergeHistoryId;
        }

        private void ShowDocumentEditingUser(MEEmrDocumentsInfo doc)
        {
            var emp = _employeeCtrl.GetObjectByID(doc.FK_EditingUserID) as HREmployeesInfo;
            var template = _templateCtrl.GetObjectByID(doc.FK_METemplateID) as METemplatesInfo;
            MessageBox.Show($"Tờ bệnh án:" +
                $"\n{doc.MEEmrDocumentGroup} > STT:{doc.MEEmrDocumentSubOrder}. {template.METemplateName} " +
                $"\nđang được soạn bởi {emp.HREmployeeName} ({doc.MEEmrDocumentHoldMachineIp}:{doc.MEEmrDocumentHoldMachineMac}) " +
                $"từ lúc {doc.MEEmrDocumentHoldFrom.ToString("HH:mm dd/MM/yyyy")}",
                "KHÔNG THỂ THỰC HIỆN ĐƯỢC THAO TÁC NÀY", MessageBoxButtons.OK, MessageBoxIcon.Stop);
        }

        internal void MergeEmrRefresh()
        {
            var emr = _entity.MainObject as MEEmrsInfo;
            _mergeHistories = new List<MEEmrMergeHistoriesInfo>();
            var emrMergeHistoryCtrl = new MEEmrMergeHistoriesController();
            var emrMergeHistories = MergeHistoryGetList(emr.MEEmrID, emrMergeHistoryCtrl);
            if (emrMergeHistories != null && emrMergeHistories.Count() > 0)
            {
                // case no add MEEmrMergeHistoryID on grid
                emrMergeHistories = emrMergeHistories.OrderBy(m => m.MEEmrMergeHistoryID).ToList();
            }
            var gridControl = this.Controls["fld_dgcMEEmrMergeHistories"] as MEEmrMergeHistoriesGridControl;
            if (gridControl != null)
            {
                gridControl.DataSource = emrMergeHistories;
                gridControl.RefreshDataSource();
                gridControl.Refresh();
            }
        }

        private List<MEEmrMergeHistoriesInfo> MergeHistoryGetList(int meEmrToId, MEEmrMergeHistoriesController emrMergeHistoryCtrl)
        {
            var emrMergeHistories = emrMergeHistoryCtrl.GetListBusinessObjects<MEEmrMergeHistoriesInfo>(emrMergeHistoryCtrl.GetAllDataByForeignColumn("FK_MEEmrToID", meEmrToId));
            if (emrMergeHistories != null)
            {
                foreach (var emrMergeHistory in emrMergeHistories)
                {
                    if (_mergeHistories.Count() > 0 && _mergeHistories.Count(m => m.MEEmrMergeHistoryID.Equals(emrMergeHistory.MEEmrMergeHistoryID)) > 0)
                    {
                        continue;
                    }
                    else
                    {
                        _mergeHistories.Add(emrMergeHistory);
                        MergeHistoryGetList(emrMergeHistory.FK_MEEmrFromID, emrMergeHistoryCtrl);
                    }
                }
            }

            return _mergeHistories;
        }

        internal void MergeEmrRollback()
        {
            var grid = this.Controls["fld_dgcMEEmrMergeHistories"] as MEEmrMergeHistoriesGridControl;
            if (grid != null)
            {
                var gridView = (grid.MainView as GridView);
                var rows = gridView.GetSelectedRows();
                if (rows.Length == 0)
                {
                    MessageBox.Show("Chưa chọn bệnh án ở thẻ [Lịch sử trộn] để thực hiện khôi phục.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }
                if (rows.Length > 1)
                {
                    MessageBox.Show("Chọn duy nhất 01 bệnh án ở thẻ [Lịch sử trộn] để thực hiện khôi phục.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                var emrMergeHistoryCtrl = new MEEmrMergeHistoriesController();
                var emrMergeHistoryDetailCtrl = new MEEmrMergeHistoryDetailsController();
                var emr = _entity.MainObject as MEEmrsInfo;

                var mergeHistory = gridView.GetRow(rows[0]) as MEEmrMergeHistoriesInfo;
                if (mergeHistory.MEEmrMergeHistoryStatus != EmrMergeHistoryStatus.Merge.ToString())
                {
                    MessageBox.Show("Chỉ khôi phục nếu [Lịch sử trộn] thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                #region Check newest history
                var nos = new List<string>();
                var newestEmrMerges = emrMergeHistoryCtrl.GetListBusinessObjects<MEEmrMergeHistoriesInfo>(emrMergeHistoryCtrl.GetAllDataByForeignColumn("FK_MEEmrToID", emr.MEEmrID));
                if (newestEmrMerges != null && newestEmrMerges.Count() > 0)
                {
                    foreach (var newestEmrMerge in newestEmrMerges)
                    {
                        nos.Add(newestEmrMerge.MEEmrFromNo);
                    }

                    if (!nos.Contains(mergeHistory.MEEmrFromNo))
                    {
                        MessageBox.Show("Chọn 1 bệnh án được trộn gần nhất ở thẻ [Lịch sử trộn] để thực hiện khôi phục.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }
                }
                #endregion

                // Khong khoi phục ba đã đc tạo lại
                var existEmr = _emrCtrl.GetObjectByNo(mergeHistory.MEEmrFromNo) as MEEmrsInfo;
                if (existEmr != null)
                {
                    MessageBox.Show($"Bệnh án {mergeHistory.MEEmrFromNo} đang tồn tại trên EMR. Không thể khôi phục lại bệnh án này.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                var emrDb = _emrCtrl.GetObjectByID(emr.MEEmrID) as MEEmrsInfo;
                if (emrDb.MEEmrStatus == EmrStatus.Closed.ToString() || emrDb.MEEmrStatus == EmrStatus.WaitClose.ToString())
                {
                    MessageBox.Show($"Không thể thực hiện vì thông tin bệnh án đã bị thay đổi. Vui lòng làm mới.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                var confirm = MessageBox.Show($"Khôi phục bệnh án [{mergeHistory.MEEmrFromNo}]. Vui lòng xác nhận?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (confirm == DialogResult.Yes)
                {
                    try
                    {
                        // 1. Update MEEmrDocuments vs MergeHistoryDetails
                        var now = DateTime.Now;
                        var emrMergeHistoryDetails = emrMergeHistoryDetailCtrl.GetListBusinessObjects<MEEmrMergeHistoryDetailsInfo>(emrMergeHistoryDetailCtrl.GetAllDataByForeignColumn("FK_MEEmrMergeHistoryID", mergeHistory.MEEmrMergeHistoryID));
                        var missDocs = new List<MEEmrMergeHistoryDetailsInfo>();
                        var processHistories = new List<MEEmrMergeHistoryDetailsInfo>();
                        var processDocs = new List<MEEmrDocumentsInfo>();
                        foreach (var emrMergeHistoryDetail in emrMergeHistoryDetails)
                        {
                            var doc = _emrDocumentCtrl.GetObjectByID(emrMergeHistoryDetail.FK_MEEmrDocumentID) as MEEmrDocumentsInfo;
                            if (doc != null)
                            {
                                processHistories.Add(emrMergeHistoryDetail);
                                processDocs.Add(doc);
                            }
                            else
                            {
                                missDocs.Add(emrMergeHistoryDetail);
                            }
                        }
                        if (missDocs.Count() > 0)
                        {
                            var confirmYN = MessageBox.Show($"Bệnh án [{mergeHistory.MEEmrFromNo}] có tờ đã xóa. Không thể khôi phục tờ bệnh án đã xóa, bạn muốn tiếp tục khôi phục bệnh án?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (confirmYN == DialogResult.No) return;
                        }

                        foreach (var doc in processDocs)
                        {
                            doc.FK_MEEmrID = mergeHistory.FK_MEEmrFromID;
                            doc.AAUpdatedDate = now;
                            doc.AAUpdatedUser = BOSApp.CurrentUser;
                            _emrDocumentCtrl.UpdateObject(doc);

                            var documentCodeU = getDocumentNo(doc);

                            // TH tạo sau khi đã trộn
                            if (!_ftpFileMng.FileExists($"/Emr/{mergeHistory.FK_MEEmrFromID}/", $"{doc.MEEmrDocumentFile}.{doc.MEEmrDocumentFileExt}"))
                            {
                                var from = $"/Emr/{emr.MEEmrID}/{doc.MEEmrDocumentFile}.{doc.MEEmrDocumentFileExt}";
                                var to = $"/Emr/{mergeHistory.FK_MEEmrFromID}/{doc.MEEmrDocumentFile}.{doc.MEEmrDocumentFileExt}";
                                var localPath = string.Format(@"{0}\Emr\{1}\{2}.{3}", _documentPath, mergeHistory.FK_MEEmrFromID, doc.MEEmrDocumentFile, doc.MEEmrDocumentFileExt);
                                _ftpFileMng.CopyFile(from, to, localPath);
                            }
                        }
                        foreach (var emrMergeHistoryDetail in processHistories)
                        {
                            emrMergeHistoryDetail.AAUpdatedDate = now;
                            emrMergeHistoryDetail.AAUpdatedUser = BOSApp.CurrentUser;
                            emrMergeHistoryDetail.AAStatus = Status.Delete.ToString();
                            emrMergeHistoryDetailCtrl.UpdateObject(emrMergeHistoryDetail);
                        }

                        // 2. Transaction Histories
                        var missMsg = missDocs.Count() > 0 ? $"(Tờ đã xóa id: {string.Join<int>(", ", missDocs.Select(m => m.FK_MEEmrDocumentID))})" : string.Empty;
                        var transfers = _transferCtrl.GetByEmrId(mergeHistory.FK_MEEmrToID).ToList();
                        var transRollback = transfers.Where(m => m.MEEmrTransferHistoryNo.Contains(mergeHistory.MEEmrFromNo));
                        if (transRollback != null && transRollback.Count() > 0)
                        {
                            foreach (var tran in transRollback)
                            {
                                tran.FK_MEEmrID = mergeHistory.FK_MEEmrFromID;
                                //tran.MEEmrTransferHistoryNo = mergeHistory.MEEmrToNo + "/" + tran.MEEmrTransferHistoryID;
                                tran.MEEmrTransferHistoryRemark = tran.MEEmrTransferHistoryRemark + $" > Khôi phục {mergeHistory.MEEmrFromNo} từ {mergeHistory.MEEmrToNo}{missMsg}";
                                tran.MEEmrTransferHistoryCurrent = true;
                                tran.AAUpdatedUser = BOSApp.CurrentUser;
                                _transferCtrl.UpdateObject(tran);
                            }
                        }
                        else
                        {
                            _transferCtrl.CreateObject(new MEEmrTransferHistoriesInfo()
                            {
                                AACreatedUser = BOSApp.CurrentUser,
                                MEEmrTransferHistoryNo = mergeHistory.MEEmrFromNo + ".1",
                                MEEmrTransferHistoryDate = now,
                                FK_MEEmrID = mergeHistory.FK_MEEmrFromID,
                                FK_HRDepartmentFromID = mergeHistory.FK_HRDepartmentToID,
                                FK_HRDepartmentToID = mergeHistory.FK_HRDepartmentFromID,
                                MEEmrTransferHistoryCurrent = true,
                                MEEmrTransferHistoryRemark = $"Khôi phục {mergeHistory.MEEmrFromNo} từ {mergeHistory.MEEmrToNo}{missMsg}"
                            });
                        }

                        // 3. Update status MEEmrs vs MergeHistories
                        var emrActiveId = mergeHistory.FK_MEEmrFromID;
                        _emrCtrl.EmrActive(emrActiveId, BOSApp.CurrentUser);

                        var activeEmr = _emrCtrl.GetObjectByID(emrActiveId) as MEEmrsInfo;
                        HistoryEmr(emr, "Change", $"Khôi phục trộn, tách tờ liên quan qua bệnh án nguồn id {activeEmr.MEEmrID}.");

                        HistoryEmr(activeEmr, "Change", $"Khôi phục trộn, khôi phục tờ liên quan từ bệnh án đích id {emr.MEEmrID}.");

                        // Khôi phục lại Bệnh nhân
                        var emrActive = _emrCtrl.GetObjectByID(emrActiveId) as MEEmrsInfo;
                        _patientCtrl.PatientActive(emrActive.FK_MEPatientID, BOSApp.CurrentUser);

                        mergeHistory.AAStatus = Status.Delete.ToString();
                        mergeHistory.AAUpdatedUser = BOSApp.CurrentUser;
                        mergeHistory.AAUpdatedDate = now;
                        emrMergeHistoryCtrl.UpdateObject(mergeHistory);

                        MessageBox.Show($"Khôi phục bệnh án [{mergeHistory.MEEmrFromNo}] thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Có lỗi khôi phục bệnh án trộn. Xin thử lại.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        PrintMgsLog("ERROR-ROLLBACK_MERGE_EMR", ex.ToString());
                    }
                    finally
                    {
                        MergeEmrRefresh();
                        QuickSearch();
                        // Refresh documents
                        Invalidate(emr.MEEmrID);
                    }
                }
            }
        }
        #endregion

        #region Insert Param
        internal void InsertParam(string paramNo)
        {
            if (!IsAllowEditPositionWithFlashNotification(_richEditCtrl.Document.CaretPosition)) return;
            Cursor.Current = Cursors.WaitCursor;
            var param = AppMemCache.GetParamFromDictKeyNo(paramNo);
            if (param == null) return;
            var doc = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            this._emrDocumentHelper.AddParamToTemplateAtCaretPosition(param, doc.MEEmrDocumentGuid, true);
            Cursor.Current = Cursors.Default;
        }

        #endregion

        #region Symbol
        internal void InitSymbolBarButton(BarSubItem barSubItemSymbol, RibbonControl ribbonEmr)
        {
            var syms = _symbolCtrl.GetListBusinessObjects<MEEmrSymbolsInfo>(_symbolCtrl.GetAllObjects()).OrderBy(o => o.MEEmrSymbolOrder);
            var groups = syms.Select(o => o.MEEmrSymbolGroup).Distinct().ToList().OrderBy(o => o);
            var gID = 0;
            foreach (var g in groups)
            {
                var btns = syms.Where(o => o.MEEmrSymbolMenu && o.MEEmrSymbolGroup == g && !string.IsNullOrEmpty(o.MEEmrSymbolNo))
                    .Select(o => new DevExpress.XtraBars.BarButtonItem()
                    {
                        Manager = ribbonEmr.Manager,
                        Caption = o.MEEmrSymbolIcon != null ? o.MEEmrSymbolName : o.MEEmrSymbolStr,
                        Id = o.MEEmrSymbolID,
                        Name = Guid.NewGuid().ToString(),
                        Glyph = ByteArrayToImage(o.MEEmrSymbolIcon),
                        Tag = o.MEEmrSymbolNo
                    }).ToList();
                foreach (var b in btns)
                    b.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.BarButtonItemSymbol_ItemClick);

                if (btns.Count > 0)
                {
                    ribbonEmr.Items.AddRange(btns.ToArray());
                    gID++;
                    var barSubItem = new BarSubItem()
                    {
                        Manager = ribbonEmr.Manager,
                        Caption = g,
                        Id = gID,
                        Name = Guid.NewGuid().ToString()
                    };
                    ribbonEmr.Items.Add(barSubItem);
                    barSubItem.AddItems(btns.ToArray());
                    barSubItem.LinksPersistInfo.AddRange(btns.Select(b => new DevExpress.XtraBars.LinkPersistInfo(b)).ToArray());
                    barSubItemSymbol.AddItem(barSubItem);
                    barSubItemSymbol.LinksPersistInfo.Add(new DevExpress.XtraBars.LinkPersistInfo(barSubItem));
                }
            }
        }

        private void BarButtonItemSymbol_ItemClick(object sender, ItemClickEventArgs e)
        {
            if (!IsAllowEditPositionWithFlashNotification(_richEditCtrl.Document.CaretPosition)) return;

            var sym = _symbolCtrl.GetObjectByNo(e.Item.Tag.ToString()) as MEEmrSymbolsInfo;
            this._emrDocumentHelper.AddSymbol(sym.MEEmrSymbolFont, sym.MEEmrSymbolChar, sym.MEEmrSymbolStr);
        }

        public Image ByteArrayToImage(byte[] byteArrayIn)
        {
            if (byteArrayIn == null) return null;
            if (byteArrayIn.Length == 0) return null;
            MemoryStream ms = new MemoryStream(byteArrayIn);
            Image returnImage = Image.FromStream(ms);
            return returnImage;

        }
        #endregion

        #region Subtemplate
        internal void InsertSubTemplate()
        {
            if (!IsAllowEditPositionWithFlashNotification(_richEditCtrl.Document.CaretPosition)) return;

            var doc = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            var templates = _templateCtrl.GetSubTemplateByRoleAndParentTemplate(doc.FK_METemplateID,
                BOSApp.CurrentUserGroupInfo.ADUserGroupRole,
                BOSApp.CurrentEmployeesInfo.HREmployeeID, BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID);
            var gui = new guiSelectSubTemplate(templates);
            gui.Module = this;
            if (gui.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    BOSProgressBar.Start("Đang xử lý dữ liệu");
                    Cursor.Current = Cursors.WaitCursor;
                    this._emrActionHelper.InsertTemplate(gui.SelectedSubTemplate, doc.MEEmrDocumentGuid);
                }
                catch (Exception)
                {

                    throw;
                }
                finally
                {
                    Cursor.Current = Cursors.Default;
                    BOSProgressBar.Close();
                }
            }
        }
        private void InsertSubTemplate(MEEmrActionsInfo action, DocumentRange actionRange, List<MEEmrActionParamsInfo> updateParams, string group)
        {
            var doc = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            var templates = _templateCtrl.GetSubTemplateByRoleAndParentTemplate(doc.FK_METemplateID,
                BOSApp.CurrentUserGroupInfo.ADUserGroupRole,
                BOSApp.CurrentEmployeesInfo.HREmployeeID, BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID);
            METemplatesInfo selected = null;
            if (templates.Count > 1)
            {
                var gui = new guiSelectSubTemplate(templates);
                gui.Module = this;
                if (gui.ShowDialog() == DialogResult.OK)
                    selected = gui.SelectedSubTemplate;
            }
            else
                selected = templates.FirstOrDefault();

            if (selected != null)
            {
                if (updateParams.Count > 0)
                    foreach (var item in updateParams)
                    {
                        var param = AppMemCache.GetParamFromDictKeyID(item.FK_MEParamID);
                        this._emrActionHelper.InsertTemplate(selected, group, actionRange.End, param.MEParamNo, item.MEEmrActionParamUpdateFirst);
                    }
                else
                    this._emrActionHelper.InsertTemplate(selected, group, actionRange.End, string.Empty, true);
            }
        }
        #endregion

        #region Query Mongo Data
        internal void QueryDocumentData()
        {
            if (!IsAllowEditPositionWithFlashNotification(_richEditCtrl.Document.CaretPosition)) return;

            var emr = GetCurrentMainObject();
            var documentList = GetCurrentDocumentList(emr.MEEmrID);
            var listDocs = _emrDocumentSortHelper.SortDocumentTree(documentList.Clone() as BOSList<MEEmrDocumentsInfo>, emr);
            var gui = new guiDataQuery(listDocs, true);
            gui.Module = this;
            if (gui.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    BOSProgressBar.Start("Đang xử lý dữ liệu");
                    if (gui.InsertParam)
                    {
                        var doc = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
                        var paramList = new List<string>();
                        var data = new JObject();
                        foreach (JProperty item in gui.SelectedRows)
                        {
                            if (!data.ContainsKey(item.Name))
                            {
                                paramList.Add(item.Name);
                                data.Add(item);
                            }
                        }
                        string group = doc.MEEmrDocumentGuid;
                        var prefix = string.Empty;
                        var range = this._emrDocumentHelper.AddParamsToTemplateAtCaretPosition(paramList, ref group, ref prefix, true);
                        //TODO: BUG khi ben trong su dung cac ham binding du lieu lai goi get field ma lai ko check range 20/03/2019
                        this._emrDocumentHelper.BindingDataToRange(data, range, group, prefix);
                    }
                    else
                    {
                        var values = new List<string>();
                        foreach (JProperty item in gui.SelectedRows)
                            GetParamsValue(item, values);
                        var doc = this._richEditCtrl.Document;
                        doc.BeginUpdate();
                        doc.InsertText(doc.CaretPosition, string.Join("; ", values.Where(s => !string.IsNullOrEmpty(s)).ToList()));
                        doc.EndUpdate();
                    }
                }
                catch (Exception)
                {

                    throw;
                }
                finally
                {
                    BOSProgressBar.Close();
                }
            }
        }
        private void GetParamsValue(JToken data, List<string> values)
        {
            if (data.Type == JTokenType.Object)
            {
                foreach (JToken child in data.Children())
                    GetParamsValue(child, values);
                return;
            }
            else if (data.Type == JTokenType.Property)
            {
                var p = (data as JProperty);
                if (p.Value is JObject || p.Value is JArray)
                    GetParamsValue(p.Value, values);
                else
                    values.Add((data as JProperty).Value.ToString());
                return;
            }
            else if (data.Type == JTokenType.Array)
            {
                for (int i = 0; i < data.Count(); i++)
                    GetParamsValue(data[i], values);
                return;
            }
            else
            {
                values.Add(data.ToString());
            }
        }
        internal JToken GetMongoDataByQuery(MEEmrDocumentsInfo document)
        {
            if (string.IsNullOrEmpty(document.MEEmrDocumentMongoID)) return JToken.FromObject(new object());
            var emr = GetCurrentMainObject();// this._entity.MainObject as MEEmrsInfo;
            var template = this._templateCtrl.GetObjectByID(document.FK_METemplateID) as METemplatesInfo;
            var filters = new Dictionary<string, object>();
            filters.Add("_id", new Tuple<MongoFilter, object>(MongoFilter.Eq, ObjectId.Parse(document.MEEmrDocumentMongoID)));
            // 1704 => filters.Add("MEEmrDocumentStatus", new Tuple<MongoFilter, object>(MongoFilter.Nin, new string[] { EmrDocumentStatus.Hidden.ToString(), EmrDocumentStatus.Discarded.ToString() }));
            filters.Add("MEEmrDocumentStatus", new Tuple<MongoFilter, object>(MongoFilter.Ne, EmrDocumentStatus.Hidden.ToString()));
            var fields = new Dictionary<string, string>();
            fields.Add("Content", "$MEEmrDocumentContent");
            JToken data = this._emrDocumentMng.Find(filters, fields, template.METemplateNo) as JToken;
            if (data != null)
                return data["Content"] as JToken;
            return null;
        }
        #endregion

        #region Extra Function
        internal void DoubleStrikeThroughAllParam()
        {
            this._emrDocumentHelper.DoubleStrikeThroughAllParam();
        }

        internal void InsertCircleToRound()
        {
            this._emrDocumentHelper.InsertCircleToRound();

        }
        internal void MedicationInfutionTakenNote()
        {
            try
            {
                var doc = this._richEditCtrl.Document;
                if (doc.Selections.Count == 1 && doc.Selections[0].Length == 0)
                {
                    MessageBox.Show("Vui lòng quét chọn các thẻ cần ghi nhận dịch truyền", "CHỌN THẺ PHÙ HỢP", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                BOSProgressBar.Start("Đang xử lý dữ liệu");
                Cursor.Current = Cursors.WaitCursor;
                var countRow = this._emrDocumentHelper.MedicationInfutionTakenNote();
                Cursor.Current = Cursors.Default;
                BOSProgressBar.Close();
                if (doc.Selections.Count > countRow)
                {
                    MessageBox.Show("Nhập lượng dịch truyền vào một trong các thẻ dữ liệu của dòng được chọn.", "VUI LÒNG NHẬP LƯỢNG DỊCH TRUYỀN", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                Cursor.Current = Cursors.Default;
                BOSProgressBar.Close();
            }
        }
        internal void ClearMedicationInfutionTakenNote()
        {
            try
            {
                BOSProgressBar.Start("Đang xử lý dữ liệu");
                Cursor.Current = Cursors.WaitCursor;
                this._emrDocumentHelper.ClearMedicationInfutionTakenNote();
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                Cursor.Current = Cursors.Default;
                BOSProgressBar.Close();
            }
        }
        #endregion

        #region Send data to HIS
        private void SendEmrToHis()
        {
            try
            {
                var confg = (new ADConfigValuesController()).GetObjectByConfigKey("ApiOfHis-SendEmrToHis");
                //khong goi dua lieu neu de trong
                if (string.IsNullOrWhiteSpace(confg.ADConfigKeyValue)) return;

                var emr = GetCurrentMainObject();// this._entity.MainObject as MEEmrsInfo;
                var emrDto = Mapper.Map<MEEmrsDto>(_emrCtrl.GetObjectByID(emr.MEEmrID) as MEEmrsInfo);
                var documents = _emrDocumentCtrl.GetListBusinessObjects<MEEmrDocumentsInfo>(_emrDocumentCtrl.GetByMEEmrID(emr.MEEmrID));
                foreach (var doc in documents)
                {
                    var docDto = Mapper.Map<MEEmrDocumentsDto>(doc);
                    docDto.MEEmrDocumentContent = GetMongoDataByQuery(doc);
                    emrDto.MEEmrDocuments.Add(docDto);
                }
                emrDto.MEPatient = Mapper.Map<MEPatientsDto>(_patientCtrl.GetObjectByID(emr.FK_MEPatientID) as MEPatientsInfo);

                var shareCtrl = new MEEmrShareHistoriesController();
                var shareds = shareCtrl.GetListBusinessObjects<MEEmrShareHistoriesInfo>(shareCtrl.GetAllDataByForeignColumn("FK_MEEmrID", emr.MEEmrID));
                foreach (var s in shareds)
                    emrDto.MEEmrShareHistories.Add(Mapper.Map<MEEmrShareHistoriesDto>(s));

                var transferCtrl = new MEEmrTransferHistoriesController();
                var transfers = transferCtrl.GetListBusinessObjects<MEEmrTransferHistoriesInfo>(transferCtrl.GetAllDataByForeignColumn("FK_MEEmrID", emr.MEEmrID));
                foreach (var t in transfers)
                    emrDto.MEEmrTransferHistories.Add(Mapper.Map<MEEmrTransferHistoriesDto>(t));


                var data = _api.Post(confg.ADConfigKeyValue, new Dictionary<string, object>(), emrDto);
                if (data != null)
                    PrintMgsLog("GOI_BENH_AN_DEN_HIS", JsonConvert.SerializeObject(data));
                else
                {
                    PrintMgsLog("GOI_BENH_AN_DEN_HIS", "Có lỗi xảy ra");
                    ShowNotification("Có lỗi xảy ra khi gởi dữ liệu bệnh án sang HIS");
                }
            }
            catch (Exception ex)
            {
                _sysHelper.LogTxt("error", ex.ToString());
            }
        }
        private void ShowNotification(string str)
        {
            if (_msgNotification.InvokeRequired)
                _msgNotification.BeginInvoke((MethodInvoker)delegate { _msgNotification.Text = str; });
            else
                _msgNotification.Text = str;
        }
        #endregion

        #region Nhom va sap xep du lieu
        internal void OrderAndGroupData()
        {
            var doc = _richEditCtrl.Document;
            //_richEditCtrl.Options.FormattingMarkVisibility.HiddenText = RichEditFormattingMarkVisibility.Visible;
            var field = _emrActionHelper.GetEmrtFieldAtPosition(doc.CaretPosition);
            if (field == null)
            {
                MessageBox.Show("Đặt trỏ chuột vào một thẻ con trong danh sách.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            int y, idx = 0;
            for (int i = field.CodeLevelArr.Length - 1; i >= 0; i--)
                if (int.TryParse(field.CodeLevelArr[i], out y))
                {
                    idx = i; break;
                }
            var prefix = string.Join(EmrParam.CodeSeparator.ToString(), field.CodeLevelArr.Take(idx));
            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            var templateParams = AppMemCache.GetTemplateParams(document.FK_METemplateID);
            var json = this._emrParser.GetDataByPrefix(field.Gid, prefix, templateParams);
            if (string.IsNullOrEmpty(prefix) || json == "[]")
            {
                MessageBox.Show("Đặt trỏ chuột vào một thẻ con trong danh sách.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var data = JsonConvert.DeserializeObject(json) as JToken;
            var param = AppMemCache.GetParamFromDictKeyNo(field.CodeLevelArr[idx - 1]);
            if (param != null)
            {
                var child = AppMemCache.GetParamRelationsFromDict(param.MEParamID);
                var codes = field.CodeLevelArr.Take(idx).Select(s =>
                {
                    return !int.TryParse(s, out y) ? s : "[0]";
                });
                var raw = data.SelectToken(string.Join(".", codes)) as JArray;
                if (raw == null)
                {
                    MessageBox.Show("Không tìm thấy dữ liệu để sắp xếp.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                var gui = new guiOrderAndGroup(raw, child);
                if (gui.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        BOSProgressBar.Start("Đang xử lý dữ liệu");
                        var value = new JObject
                        {
                            { param.MEParamNo, gui.FormatedData }
                        };
                        prefix = string.Join(EmrParam.CodeSeparator.ToString(), field.CodeLevelArr.Take(idx - 1));
                        doc.BeginUpdate();
                        if (field.Field.Parent != null)
                            this._emrDocumentHelper.BindingDataToRange(value, field.Field.Parent.Range, field.Gid, prefix);
                        else
                            this._emrDocumentHelper.BindingDataToRange(value, doc.Range, field.Gid, prefix);

                        if (idx >= 2 && int.TryParse(field.CodeLevelArr[idx - 2], out y))
                            prefix = string.Join(EmrParam.CodeSeparator.ToString(), field.CodeLevelArr.Take(idx));

                        this._emrDocumentHelper.GroupingList(gui.FormatedData as JArray, EmrParam.MedicationGroup, gui.FormatType, prefix, param, field.Gid);
                        doc.EndUpdate();
                    }
                    catch (Exception)
                    {
                        throw;
                    }
                    finally
                    {
                        BOSProgressBar.Close();
                    }
                }
            }
        }
        #endregion

        #region An to benh an
        internal void HidingEmrDocument(MEEmrDocumentsInfo row)
        {
            guiSign guiSign = new guiSign("Đang thực hiện ẩn tờ bệnh án. Sẽ không thể hoàn tác. Vẫn thực hiện?", _fingerPrintZK);
            if (guiSign.ShowDialog() == DialogResult.OK)
            {
                var preStatus = string.IsNullOrEmpty(row.MEEmrDocumentPreStatus) ? EmrDocumentStatus.InProgress.ToString() : row.MEEmrDocumentPreStatus;
                row.MEEmrDocumentPreStatus = row.MEEmrDocumentStatus;

                if (row.MEEmrDocumentStatus != EmrDocumentStatus.Hidden.ToString())
                {
                    row.MEEmrDocumentStatus = EmrDocumentStatus.Hidden.ToString();
                    BackgroundEmrDocumentValidatesDelete(_entity.MainObject as MEEmrsInfo, row.FK_METemplateID, row.MEEmrDocumentID);
                }
                else
                {
                    row.MEEmrDocumentStatus = preStatus;
                    BackgroundEmrDocumentValidatesActive(_entity.MainObject as MEEmrsInfo, row.FK_METemplateID, row.MEEmrDocumentID);
                }
                row.AAUpdatedUser = BOSApp.CurrentUser;
                _emrDocumentCtrl.UpdateObject(row);



                var documentCodeU = getDocumentNo(row);
                HistoryEmr(GetCurrentMainObject(), "Change", $"{documentCodeU} ẩn tờ.");

                var mongoDoc = Mapper.Map<Clas.Model.Mongo.EmrDocument>(row);
                mongoDoc.AAUpdatedDate = DateTime.Now.ToLocalTime();
                mongoDoc.AAUpdatedUser = BOSApp.CurrentUser;
                var excludedFields = new List<string>() { "RefId", "v", "MEEmrDocumentContent", "MEEmr" };
                _emrDocumentMng.Update(row.MEEmrDocumentMongoID, mongoDoc, row.MEEmrDocumentNo, excludedFields);
                this.InvalidateEmrDocumentList(row.FK_MEEmrID);
            }
        }
        #endregion

        #region Ho so benh nhan
        internal void ShowNewPatientDocumentDialog()
        {
            _entity.MENewEmrDocument = null;
            var patientProfile = _entity.ModuleObjects[TableName.MEEmrsTableName] as MEEmrsInfo;

            if (_sysHelper.AllowThread())
            {
                var thCreateEmrDir = new System.Threading.Thread(() => CreateEmrDir(patientProfile.MEEmrID));
                thCreateEmrDir.Start();
            }
            else
            {
                CreateEmrDir(patientProfile.MEEmrID);
            }
            var isFilterTemplate = false;
            var emrType = _emrTypeCtrl.GetObjectByID(patientProfile.FK_MEEmrTypeID) as MEEmrTypesInfo;
            if (emrType != null)
            {
                isFilterTemplate = emrType.MEEmrTypeFilterTemplate;
            }
            var mETemplateList = _entity.METemplateList;
            if (isFilterTemplate)
            {
                mETemplateList = this._templateCtrl.GetTemplatesByEmrType(TemplateType.ProgressNote.ToString(), BOSApp.CurrentUserGroupInfo.ADUserGroupID, patientProfile.FK_MEEmrTypeID);
            }
            var gui = new DSMEEMR100(mETemplateList)
            {
                Module = this,
                Text = "Thêm tờ vào hồ sơ bệnh nhân " + patientProfile.MEEmrNo
            };
            gui.ShowDialog();
            if (gui.DialogResult == DialogResult.OK)
            {
                try
                {
                    AddNewPatientDocument(null);
                }
                catch (Exception ex)
                {
                    this.InvalidatePatientDocuments(patientProfile.MEEmrID);
                    this.ClearDocumentSession();
                    MessageBox.Show("Có lỗi khi thêm tờ bệnh án. Xin thử lại.", "CÓ LỖI KHI TẠO TỜ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    PrintMgsLog("LOI-ADD-NEW-PATIENT-DOCUMENT-MANUAL", ex.ToString());
                }
            }
        }
        private MEEmrDocumentsInfo AddNewPatientDocument(Dictionary<string, object> initData)
        {
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                if (_entity.MENewEmrDocument != null)
                {
                    var emr = _entity.ModuleObjects[TableName.MEEmrsTableName] as MEEmrsInfo;
                    var document = this.AddNewDocument(emr, _entity.MEEmrPatientDocumentsList, initData);
                    if (document != null)
                    {
                        var data = JsonConvert.DeserializeObject(document.MEEmrDocumentJson) as JToken;
                        document = ExtractDataFromContentToDocumentInfo(document, data);

                        emr = UpdateEmrInfoByDocumentContent(emr, data);

                        _entity.MEEmrPatientDocumentsList.AddObjectToList();
                        _entity.MEEmrPatientDocumentsList.SaveItemObjects();
                        this.InvalidatePatientDocuments(emr.MEEmrID);

                        var grid = _entity.MEEmrPatientDocumentsList.GridView;
                        for (int i = 0; i < _entity.MEEmrPatientDocumentsList.Count; i++)
                        {
                            if (_entity.MEEmrPatientDocumentsList[i].MEEmrDocumentFile == document.MEEmrDocumentFile)
                            {
                                grid.FocusedRowHandle = grid.GetRowHandle(i);
                                _entity.MEEmrPatientDocumentsList[i].MEEmrDocumentJson = document.MEEmrDocumentJson;
                                document = _entity.MEEmrPatientDocumentsList[i];
                                //save to mongo db
                                InsertMongoDocument(emr, document);
                                document.AAUpdatedUser = BOSApp.CurrentUser;
                                this._emrDocumentCtrl.UpdateObject(document);

                                var documentCodeU = getDocumentNo(document);
                                HistoryEmr(GetCurrentMainObject(), "Change", $"cập nhật thông tin tờ bệnh án [{documentCodeU}].");

                                break;
                            }
                        }
                        InvalidateDocument(document, false);
                        return document;
                    }
                }
            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {
                BOSProgressBar.Close();
                _entity.MENewEmrDocument = null;
            }

            Cursor.Current = Cursors.Default;
            return null;
        }
        public void CreateEmrWithTypePatientProfile()
        {
            if (!IsNothingToSaveDocumentContent()) return;

            var emrPatientProfile = _emrCtrl.GetEmrPatientProfile(_entity.MEPatient.MEPatientID);
            if (emrPatientProfile != null)
            {
                MessageBox.Show("Bệnh nhân đã có hồ sơ", "Mỗi bệnh nhân chỉ có một hồ sơ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var type = _emrTypeCtrl.GetPatientProfileType();
            if (type == null)
            {
                MessageBox.Show("Chưa cấu hình", "Chưa cấu hình loại bệnh án 'Hồ sơ bệnh nhân'", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var patient = _entity.MEPatient;
            var emr = new MEEmrsInfo()
            {
                MEEmrNo = patient.MEPatientNo,
                MEEmrDesc = "Hồ sơ bệnh nhân " + patient.MEPatientNo,
                MEEmrStatus = EmrStatus.InProgress.ToString(),
                MEEmrCreatedDate = DateTime.Now,
                FK_HRDepartmentID = BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID,
                FK_HREmployeeCreatedID = BOSApp.CurrentEmployeesInfo.HREmployeeID,
                FK_MEEmrTypeID = type.MEEmrTypeID,
                FK_MEPatientID = patient.MEPatientID,
                MEEmrTypeProfile = EmrTypeProfile.Patient.ToString(),
                AACreatedUser = BOSApp.CurrentUser
            };
            _emrCtrl.CreateObject(emr);
            emr = _emrCtrl.GetEmrPatientProfile(_entity.MEPatient.MEPatientID);
            _entity.InvalidateModuleObject(emr);

            var tab = (Controls["tabPatientProfiles"] as DevExpress.XtraTab.XtraTabPage);
            tab.PageVisible = true;
            tab.TabControl.SelectedTabPage = tab;

            InitPatientProfile(emr);
        }
        private void InitPatientProfile(MEEmrsInfo emr)
        {
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                BOSProgressBar.Start("Đang khởi tạo hồ sơ bệnh nhân");
                PrintMgsLog("BAT-DAU-KHOI-TAO-HO-SO-BENH-NHAN", emr.MEEmrNo);

                CreateDocumentsFromEmrType(emr, _entity.MEEmrPatientDocumentsList);

                var backupDocs = _entity.MEEmrPatientDocumentsList.Select(o => new MEEmrDocumentsInfo()
                {
                    MEEmrDocumentFile = o.MEEmrDocumentFile,
                    MEEmrDocumentJson = o.MEEmrDocumentJson
                }).ToList();

                _entity.UpdateModuleObjectBindingSource(TableName.MEEmrDocumentsTableName);
                _entity.MEEmrPatientDocumentsList.SaveItemObjects();
                this.InvalidatePatientDocuments(emr.MEEmrID);
                foreach (var doc in _entity.MEEmrPatientDocumentsList)
                {   //save to mongo db
                    doc.MEEmrDocumentJson = backupDocs.Where(o => o.MEEmrDocumentFile == doc.MEEmrDocumentFile).FirstOrDefault()?.MEEmrDocumentJson;
                    InsertMongoDocument(emr, doc);
                }
                _entity.MEEmrPatientDocumentsList.SaveItemObjects();

                PrintMgsLog("KHOI-TAO-HO-SO-BENH-NHAN-THANH-CONG", emr.MEEmrNo);

                this.ActivateScreen("DMMEEMR100");
                this.DockManager.ActivePanel = this._dpnRichEdit;
                this.OpenGuidance();

                if (_sysHelper.AllowThread())
                {
                    //uthv co the chay am tham ben duoi ma ko anh huong
                    var thSendEmrToHis = new System.Threading.Thread(() => SendEmrToHis());
                    thSendEmrToHis.Start();
                }
                else
                {
                    SendEmrToHis();
                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                BOSProgressBar.Close();
                Cursor.Current = Cursors.Default;
            }
        }
        private void InvalidatePatientDocuments(int emrId)
        {
            var _focusedRowHandle = _entity.MEEmrPatientDocumentsList.GridView.FocusedRowHandle;
            var listDocs = _emrDocumentCtrl.GetListBusinessObjects<MEEmrDocumentsInfo>(_emrDocumentCtrl.GetByMEEmrID(emrId));
            //chi co admin moi thay danh sach cac to benh an da bi an
            if (BOSApp.CurrentUserGroupInfo.ADUserGroupRole != UserGroupRole.admin.ToString())
            {
                listDocs = listDocs.Where(o => o.MEEmrDocumentStatus != EmrDocumentStatus.Hidden.ToString()).ToList();
            }
            _entity.MEEmrPatientDocumentsList.Invalidate(listDocs.OrderBy(o => o.MEEmrDocumentOrder).ThenByDescending(o => o.MEEmrDocumentCreatedDate).ToList());
            _entity.MEEmrPatientDocumentsList.GridView.FocusedRowHandle = -1;

            var meEmr = _entity.MainObject as MEEmrsInfo;
            meEmr.MEPatientDocumentCount = _entity.MEEmrPatientDocumentsList.Count;
        }

        #endregion

        #region Chen hinh anh tu file shared
        private object InsertImageFilesShared(Dictionary<string, object> sendParams, List<MEEmrActionParamsInfo> paramList, object data, MEEmrActionsInfo action, string group)
        {
            if (paramList.Count == 0) return data;
            //chac chan la luu tai lieu truoc
            var command = this._richEditCtrl.CreateCommand(RichEditCommandId.FileSave);
            command.Execute();
            var imagesParam = paramList.Where(p => !p.MEEmrActionParamRequest).ToList();
            PrintMgsLog("BAT-DAU-GOI-API", action.MEEmrActionUri);

            if (action.MEEmrActionHttpMethod == HttpMethod.POST.ToString())
                data = _api.Post(action.MEEmrActionUri, null, sendParams);
            else
                data = _api.Get(action.MEEmrActionUri, sendParams);

            PrintMgsLog("KET-THUC-GOI-API", action.MEEmrActionUri);
            if (data != null)
            {
                InsertImageFilesShared(imagesParam, data, action, group);
            }
            return data;
        }
        private object InsertImageFilesShared(List<MEEmrActionParamsInfo> imagesParam, object data, MEEmrActionsInfo action, string group)
        {
            var values = this._dataHelper.FlattenObjDataWithoutChangeName(string.Empty, data as JObject);
            foreach (var item in imagesParam)
            {
                var param = AppMemCache.GetParamFromDictKeyID(item.FK_MEParamID);
                var paths = new List<string>();
                var regex = new System.Text.RegularExpressions.Regex(param.MEParamNo + @"\[\d+\]");
                foreach (var prop in values)
                {
                    var r = regex.Match(prop.Key);
                    if (r.Success || prop.Key.EndsWith(param.MEParamNo))
                    {
                        if (prop.Value is JArray)
                            foreach (var path in (prop.Value as JArray))
                            {
                                paths.Add(path.ToString());
                            }
                        else
                            paths.Add(prop.Value.ToString());
                    }
                }
                this._emrDocumentHelper.InsertImageToParam(param, action, group, paths);
            }
            return data;
        }
        #endregion

        #region Dinamap Pro V100
        private int InsertDinamapProV100DataAsync(Dictionary<string, object> paramList,
            List<MEEmrActionParamsInfo> allParams, object data, MEEmrActionsInfo action, string group)
        {
            DetectAndGetV100Port();
            if (_v100SerialPort == null)
            {
                ShowFlashNotification("Không có kết nối với GE Dinamap Pro V100", 3000);
                return 1;
            }
            try
            {
                _v100SerialPort.Open();
            }
            catch (Exception ex)
            {
                ShowFlashNotification("Có lỗi, vui lòng kiểm tra kết nối đến Dinamap V100", 3000);
                PrintMgsLog("GE DINAMAP PRO V100 CONNECT ERROR", ex.ToString());
                return 1;
            }

            var state = Task.Run(async () => await _v100SerialPort.GetStateAsync()).ConfigureAwait(false).GetAwaiter().GetResult();
            var bindingData = new Dictionary<string, object>();
            foreach (var p in allParams)
            {
                var param = AppMemCache.GetParamFromDictKeyID(p.FK_MEParamID) as MEParamsInfo;
                switch (p.MEEmrActionParamSourcePath)
                {
                    case "Systolic":
                        if (string.IsNullOrEmpty(state.BloodPrsMsg))
                            bindingData.Add(param.MEParamNo, state.Systolic);
                        break;
                    case "Diastolic":
                        if (string.IsNullOrEmpty(state.BloodPrsMsg))
                            bindingData.Add(param.MEParamNo, state.Diastolic);
                        break;
                    case "Oxygen":
                        if (string.IsNullOrEmpty(state.OxygenMsg))
                            bindingData.Add(param.MEParamNo, state.Oxygen);
                        break;
                    case "HeartRate":
                        if (string.IsNullOrEmpty(state.HeartRateMsg))
                            bindingData.Add(param.MEParamNo, state.HeartRate);
                        break;
                    case "Temperature":
                        if (string.IsNullOrEmpty(state.TemperatureMsg))
                            bindingData.Add(param.MEParamNo, state.Temperature);
                        break;

                    default:
                        break;
                }
            }
            if (bindingData.Count > 0)
            {
                BindingDataToEmrDocument(bindingData, group, string.Empty);
                var hard = GetHardParamList(GetCurrentMainObject());
                BindingDataToEmrDocument(hard, group, string.Empty);
            }
            var msg = !string.IsNullOrEmpty(state.BloodPrsMsg) ? state.BloodPrsMsg + "; " : string.Empty;
            msg += !string.IsNullOrEmpty(state.OxygenMsg) ? state.OxygenMsg + "; " : string.Empty;
            msg += !string.IsNullOrEmpty(state.HeartRateMsg) ? state.HeartRateMsg + "; " : string.Empty;
            msg += !string.IsNullOrEmpty(state.TemperatureMsg) ? state.TemperatureMsg + "; " : string.Empty;
            if (msg.Length > 1)
            {
                ShowFlashNotification(msg, 10000);
                return 1;
            }
            _v100SerialPort.Close();
            ShowFlashNotification("Dữ liệu đã được cập nhật", 3000);
            return 0;
        }

        private void DetectAndGetV100Port()
        {
            if (_v100SerialPort != null) return;

            if (!string.IsNullOrEmpty(BOSApp.V100COMPort) && _v100SerialPort == null)
                _v100SerialPort = new GeV100SerialConn(BOSApp.V100COMPort);
            else
            {
                if (MessageBox.Show("Cắm cáp USB đến Dinamap GE V100 vào PC sau đó bấm 'OK' và thử lại.", "Không tìm thấy thiết bị", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
                {
                    Task.Run(async () =>
                    {
                        var port = await Emr.Devices.GeV100.GeV100SerialConn.ComPortDetectAsync();
                        if (string.IsNullOrEmpty(port))
                            MessageBox.Show("Không tìm thấy thiết bị", "Không tìm thấy thiết bị", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        else
                            MessageBox.Show("Thiết bị đã được cắm vào: " + port, "Có thiết bị mới", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        BOSApp.V100COMPort = port;
                    });
                }
                else
                {
                    return;
                }
            }
        }

        private object PrintV100Snap()
        {
            //in snap v100
            DetectAndGetV100Port();
            if (_v100SerialPort != null)
                _v100SerialPort.WriteCommand(" PS!F");
            return null;
        }

        internal void MonitoringVitalSign()
        {
            DetectAndGetV100Port();
            if (_v100SerialPort == null) return;
            var gui = new guiGeV100Monitor(_v100SerialPort);
            gui.Module = this;
            gui.WindowState = FormWindowState.Maximized;
            gui.Show();
        }

        #endregion

        #region Admin nhả quyền edit
        public void ForceReleaseDocument()
        {
            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            var doc = _emrDocumentCtrl.GetObjectByID(document.MEEmrDocumentID) as MEEmrDocumentsInfo;
            if (doc != null)
            {
                if (doc.FK_EditingUserID > 0)
                {
                    var emp = _employeeCtrl.GetObjectByID(doc.FK_EditingUserID) as HREmployeesInfo;
                    if (MessageBox.Show($"Tờ bệnh án này đang được giữ bởi {emp.HREmployeeName} ({doc.MEEmrDocumentHoldMachineIp}:{doc.MEEmrDocumentHoldMachineMac}) từ lúc {doc.MEEmrDocumentHoldFrom.ToString("HH:mm dd/MM/yyyy")}",
                        "Sẽ lấy lại quyền sửa trên tờ bệnh án này?", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                    {
                        this._emrDocumentCtrl.ForceReleaseEditingPermission(document.MEEmrDocumentID, doc.FK_EditingUserID);
                        var objGeObjectHistoryInfo = new GEObjectHistoryInfo
                        {
                            ADUserID = BOSApp.CurrentUsersInfo.ADUserID,
                            ADUserName = BOSApp.CurrentUser,
                            GEObjectHistoryAction = cstObjectHistoryActionRevokeEditPer,
                            GEObjectHistoryObjectID = document.MEEmrDocumentID,
                            GEObjectHistoryObjectName = TableName.MEEmrDocumentsTableName,
                            GEObjectHistoryObjectNumber = document.MEEmrDocumentFile,
                            GEObjectHistoryDate = DateTime.Now
                        };
                        _geObjHistoryCtrl.CreateObject(objGeObjectHistoryInfo);
                        MessageBox.Show($"Đã lấy lại quyền.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }
        #endregion

        #region Loại hồ sơ
        internal void ChangeEmrType(int typeId)
        {
            if (typeId <= 0) return;
            if (Toolbar.IsNullOrNoneAction()) return;
            var type = _emrTypeCtrl.GetObjectByID(typeId) as MEEmrTypesInfo;
            var emr = _entity.MainObject as MEEmrsInfo;
            emr.MEEmrTypeProfile = type.MEEmrTypeProfile;
            emr.FK_MEEmrTypeID = typeId;
            _entity.UpdateMainObjectBindingSource();
        }
        #endregion

        #region Print toan bo Emr
        public void PrintEmr()
        {
            var emr = _entity.MainObject as MEEmrsInfo;
            var listDocs = _emrDocumentSortHelper.SortDocumentTree(_entity.MEEmrDocumentsList.Clone() as BOSList<MEEmrDocumentsInfo>, emr);
            var gui = new guiPrintEmr(listDocs)
            {
                Module = this
            };
            gui.ShowDialog();
        }
        internal void PreviewDocumentForPrint(MEEmrDocumentsInfo doc)
        {
            if (doc.MEEmrDocumentFileExt == EmrDocumentFileExtention.docx.ToString())
            {
                var richContrl = this.PrepareDocumentToPrint(doc);
                ShowPreviewPrintDialog(richContrl);
            }
            else if (doc.MEEmrDocumentFileExt == EmrDocumentFileExtention.pdf.ToString())
            {
                var gui = new guiPreviewPdf(doc);
                gui.StartPosition = FormStartPosition.WindowsDefaultLocation;
                gui.Module = this;
                gui.Show();
            }
        }
        internal void QuickPrintDocument(MEEmrDocumentsInfo doc)
        {
            if (doc.MEEmrDocumentFileExt == EmrDocumentFileExtention.docx.ToString())
            {
                var richContrl = this.PrepareDocumentToPrint(doc);
                QuickPrintDocument(richContrl);
            }
            else if (doc.MEEmrDocumentFileExt == EmrDocumentFileExtention.pdf.ToString())
            {
                var filePath = DownloadFtpFile(doc);
                QuickPrintPdf(filePath);
            }
        }
        internal void QuickPrintPdf(string filePath)
        {
            if (!string.IsNullOrEmpty(filePath))
            {
                var viewer = new PdfViewer();
                var settings = new DevExpress.Pdf.PdfPrinterSettings(new PrinterSettings());
                using (FileStream stream = new FileStream(filePath, FileMode.Open))
                {
                    viewer.ShowPrintStatusDialog = false;
                    viewer.DetachStreamAfterLoadComplete = true;
                    viewer.LoadDocument(stream);
                    viewer.Print(settings);
                }
            }
        }
        private RichEditControl PrepareDocumentToPrint(MEEmrDocumentsInfo document)
        {
            DownAndLoadDocx(this._tempRichEditCtrl, document);
            RemoveAllForPrint(this._tempRichEditCtrl, document.FK_METemplateID);
            return this._tempRichEditCtrl;
        }
        public void DownAndLoadDocx(RichEditControl richCtrl, MEEmrDocumentsInfo document)
        {
            try
            {
                var filePath = DownloadFtpFile(document);
                if (!File.Exists(filePath))
                {
                    MessageBox.Show(("File bệnh án không tồn tại ở địa chỉ. " + filePath), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                using (OfficeOpenXmlCrypto.OfficeCryptoStream stream = OfficeOpenXmlCrypto.OfficeCryptoStream.Open(filePath, this._emrDocumentHelper.ShareEmrPassword))
                {
                    richCtrl.LoadDocument(stream, DocumentFormat.OpenXml);
                }
            }
            catch (OfficeOpenXmlCrypto.InvalidPasswordException ex)
            {
                MessageBox.Show("Có thể đã có lỗi trong quá trình mã hóa và upload tờ bệnh án. Không mở được tập tin.", "Không giải mã được tờ bệnh án", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        public string DownloadFtpFile(MEEmrDocumentsInfo document)
        {
            string filePath = string.Format(@"{0}\Emr\{1}\{2}.{3}", _documentPath, document.FK_MEEmrID, document.MEEmrDocumentFile, document.MEEmrDocumentFileExt);
            _ftpFileMng.DownloadFile($"/Emr/{document.FK_MEEmrID}/", document.MEEmrDocumentFile + "." + document.MEEmrDocumentFileExt, filePath);
            return filePath;
        }

        public string DownloadFtpFileTemp(MEEmrDocumentsInfo document)
        {
            string filePath = string.Format(@"{0}\Emr\{1}\{2}.{3}", _documentPath, document.FK_MEEmrID, $"{document.MEEmrDocumentFile}_Temp", document.MEEmrDocumentFileExt);
            try
            {
                _ftpFileMng.DownloadFile($"/Emr/{document.FK_MEEmrID}/", document.MEEmrDocumentFile + "." + document.MEEmrDocumentFileExt, filePath);
                return filePath;
            }
            catch (Exception ex)
            {
                Trace.TraceError("FtpFileMng.DownloadFile ERROR: {0}:{1}:{2}:{3}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), BOSApp.CurrentUser, ex, filePath);
                Trace.Flush();
                MessageBox.Show($"Không tải được tờ bệnh án từ máy chủ.",
                     "CÓ LỖI XẢY RA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return string.Empty;
            }
        }
        #endregion

        #region Emr Dir
        private void CreateEmrDir(int emrId)
        {
            try
            {
                CreateEmrLocalDir(emrId);
                CreateEmrFtpDir(emrId);
            }
            catch (Exception ex)
            {
                _sysHelper.LogTxt("error", ex.ToString());
            }
        }
        private void CreateEmrLocalDir(int emrId)
        {
            string localDir = string.Format(@"{0}\Emr\{1}\", _documentPath, emrId);
            if (!Directory.Exists(localDir))
                Directory.CreateDirectory(localDir);

            localDir = string.Format(@"{0}\Emr\Partials\{1}\", _documentPath, emrId);
            if (!Directory.Exists(localDir))
                Directory.CreateDirectory(localDir);

            localDir = string.Format(@"{0}\Emr\Signed\{1}\", _documentPath, emrId);
            if (!Directory.Exists(localDir))
                Directory.CreateDirectory(localDir);
        }
        private void CreateEmrFtpDir(int emrId)
        {
            _ftpFileMng.CreateDirectory("/Emr/" + emrId);
            _ftpFileMng.CreateDirectory("/Emr/Partials/" + emrId);
            _ftpFileMng.CreateDirectory("/Emr/Signed/" + emrId);
        }
        #endregion

        #region Dll Plugin
        private object CallDllPlugin(MEEmrActionsInfo action, string group, List<MEEmrActionParamsInfo> updateParams, Dictionary<string, object> requestParams, MEEmrDocumentsInfo meEmrDocumentsInfo = null)
        {
            JToken data = null;
            try
            {
                if (action.MEEmrActionType == EmrActionTypes.CalcPlugin.ToString())
                {
                    var plugin = Plugin.CreatePlugin<ICalcPlugin>(Path.Combine(BOSApp.AppLocation, "plugins", action.MEEmrActionNo), action.MEEmrActionPlugin);
                    var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
                    var templateParams = AppMemCache.GetTemplateParams(document.FK_METemplateID > 0 ? document.FK_METemplateID : meEmrDocumentsInfo.FK_METemplateID);
                    var documentData = this._emrParser.ParserFieldsToJToken(this._emrDocumentHelper.GetAllDataFieldInRangeOrDocument(), templateParams);
                    data = plugin.Calculate(action.MEEmrActionNo, group, requestParams, documentData);
                }
                else
                {
                    var plugin = Plugin.CreatePlugin<IDataPlugin>(Path.Combine(BOSApp.AppLocation, "plugins", action.MEEmrActionNo), action.MEEmrActionPlugin);
                    switch (action.MEEmrActionDataType)
                    {
                        case "Json":
                            data = plugin.GetJToken(action.MEEmrActionNo, group, requestParams);
                            break;
                        case "Xml":
                            break;
                        case "Image":
                            break;
                        default:
                            break;
                    }
                }
            }
            catch (UserFriendlyException ex)
            {
                MessageBox.Show(ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                PrintMgsLog("Có lỗi khi gọi Trình cắm", ex.ToString());
            }
            return data;
        }
        #endregion

        #region Copy data in doc to place
        internal void CopyFromRowOrColumn()
        {
            if (!IsAllowEditPositionWithFlashNotification(_richEditCtrl.Document.CaretPosition)) return;
            var msg = "Đặt trỏ chuột vào một thẻ con trong danh sách [thẻ đích đến] > Click Copy dòng/cột > Chọn dữ liệu nguồn cần copy [thẻ nguồn].";
            var doc = _richEditCtrl.Document;
            var field = _emrActionHelper.GetEmrtFieldAtPosition(doc.CaretPosition);
            if (field == null)
            {
                MessageBox.Show(msg, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            int currentIdx = -1, idx = -1, parentIdx = -1, currentParentIdx = -1;
            for (int i = field.CodeLevelArr.Length - 1; i >= 0; i--)
                if (int.TryParse(field.CodeLevelArr[i], out currentIdx))
                {
                    idx = i; break;
                }
            for (int i = idx - 1; i >= 0; i--)
                if (int.TryParse(field.CodeLevelArr[i], out currentParentIdx))
                {
                    parentIdx = i; break;
                }
            var prefix = string.Join(EmrParam.CodeSeparator.ToString(), field.CodeLevelArr.Take(idx));
            //kiem tra cha co phai la danh sach hay khong.
            if (parentIdx > 0)
            {
                var parentParam = AppMemCache.GetParamFromDictKeyNo(field.CodeLevelArr[parentIdx - 1]) as MEParamsInfo;
                if (parentParam.MEParamType == EmrParamTypes.List.ToString())
                    prefix = string.Join(EmrParam.CodeSeparator.ToString(), field.CodeLevelArr.Take(parentIdx));
            }
            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            var templateParams = AppMemCache.GetTemplateParams(document.FK_METemplateID);
            var json = this._emrParser.GetDataByPrefix(field.Gid, prefix, templateParams);
            if (string.IsNullOrEmpty(prefix) || json == "[]")
            {
                MessageBox.Show(msg, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var data = JsonConvert.DeserializeObject(json) as JToken;
            var param = AppMemCache.GetParamFromDictKeyNo(field.CodeLevelArr[idx - 1]) as MEParamsInfo;
            if (param != null)
            {
                var codes = field.CodeLevelArr.Take(idx).Select(s =>
                {
                    return !int.TryParse(s, out int y) ? s : "[*]";
                });
                var raw = new JArray();
                foreach (var item in data.SelectTokens(string.Join(".", codes)))
                {
                    raw.Add(item.DeepClone());
                }
                if (raw.Count == 0)
                {
                    MessageBox.Show("Không tìm thấy dữ liệu để copy.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                //xac dinh currentIdx trong truong hop the cha la danh sach
                if (parentIdx > 0)
                {
                    currentIdx = raw.Take(currentParentIdx).Sum(o => o.Count()) + currentIdx;
                }
                raw = DecreaseArrLevel(raw);
                //currentIdx is row index where pointer at. can't select that row
                var gui = new guiDataSelection(raw, param, false, currentIdx)
                {
                    Module = this,
                    StartPosition = FormStartPosition.CenterParent
                };
                if (gui.ShowDialog() == DialogResult.OK)
                {
                    var dataPrefix = string.Join(EmrParam.CodeSeparator.ToString(), field.CodeLevelArr.Take(idx + 1));
                    var bindingData = _dataHelper.FlattenObjData(string.Empty, (gui.SelectedRow.DeepClone() as JObject));
                    for (int i = 0; i < bindingData.Count; i++)
                    {
                        if (bindingData.Keys.ElementAt(i).EndsWith(EmrParam.ChuKyHinh)
                            || bindingData.Keys.ElementAt(i).EndsWith(EmrParam.ChuKyHoTen)
                            || bindingData.Keys.ElementAt(i).EndsWith(EmrParam.ChuKyTen)
                            || bindingData.Keys.ElementAt(i).EndsWith(EmrParam.ChuKyNguoiDung)
                            || bindingData.Keys.ElementAt(i).EndsWith(EmrParam.ChuKyThoiGian))
                        {
                            bindingData.Remove(bindingData.Keys.ElementAt(i));
                            i--;
                        }
                    }
                    //#1958: Copy dòng cột hiển thị luôn cả đơn vị tính => Không hiển thị
                    BindingDataToEmrDocument(bindingData, field.Gid, dataPrefix, false);
                }
            }
        }

        internal void PasteParamValueAtRanges(KeyEventArgs e)
        {
            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            var ok = _emrDocumentHelper.PasteParamValue(document.MEEmrDocumentGuid);
            //UtHV chặn không cho thực hiện quét chọn và paste raw text
            e.Handled = true;
            e.SuppressKeyPress = true;
        }
        internal void PasteParamValue(KeyEventArgs e)
        {
            var ok = PasteParamValue();
            if (ok)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }
        public bool PasteParamValue()
        {
            if (Clipboard.ContainsText(TextDataFormat.UnicodeText))
            {
                string clipboardText = Clipboard.GetText(TextDataFormat.UnicodeText);
                var doc = _richEditCtrl.Document;
                var listField = new List<Field>();
                foreach (var selection in doc.Selections)
                {
                    listField.AddRange(_emrDocumentHelper.GetAllDataFieldInRange(selection));
                }
                if (listField.Count > 0)
                {
                    doc.BeginUpdate();
                    foreach (var field in listField)
                    {
                        doc.Replace(field.ResultRange, $"{EmrParam.BeginTag}{clipboardText}{EmrParam.EndTag}");
                    }
                    doc.EndUpdate();
                    return true;
                }
            }
            return false;
        }

        internal void CopyParamValue()
        {
            var doc = _richEditCtrl.Document;
            if (doc.Selections.Count == 1 && doc.Selections[0].Length == 0)
            {
                var field = _emrActionHelper.GetEmrtFieldAtPosition(doc.CaretPosition);
                if (field == null)
                {
                    ShowFlashNotification("Đặt trỏ chuột vào một thẻ và Ctrl+C để copy nội dung thẻ.", 4000);
                    return;
                }
                var value = _emrParser.GetFieldValue(field.Field);
                if (!string.IsNullOrEmpty(value))
                    Clipboard.SetText(value);
            }
        }
        internal void CopyJsonValueOfRange(KeyEventArgs e)
        {
            var doc = _richEditCtrl.Document;
            if (doc.Selections.Count > 0 && doc.Selections[0].Length > 0)
            {
                List<Field> fields = new List<Field>();
                foreach (var selection in doc.Selections)
                {
                    fields.AddRange(_emrDocumentHelper.GetAllDataFieldInRange(selection));
                }
                if (fields.Count > 0)
                {
                    var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
                    var json = _emrParser.ParserFieldsToJson(fields, AppMemCache.GetTemplateParams(document.FK_METemplateID));
                    if (!string.IsNullOrEmpty(json))
                    {
                        Clipboard.SetText("JSON:" + json);
                        ShowFlashNotification("Đã copy dữ liệu của thẻ. Chọn vùng tương đương và thực hiện Ctrl+V.", 3000);
                    }

                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
            }
        }
        #endregion

        #region Checkbox on doc
        internal void CheckEmrParam()
        {
            var doc = _richEditCtrl.Document;
            var field = _emrActionHelper.GetEmrtFieldAtPosition(doc.CaretPosition);
            if (field == null) return;
            if (field.CodeLevelArr.Length < 2) return; //included HYPERLINK
            var parentCode = field.CodeLevelArr[field.CodeLevelArr.Length - 2];
            if (int.TryParse(parentCode, out int y))
                parentCode = field.CodeLevelArr[field.CodeLevelArr.Length - 3];
            var parent = AppMemCache.GetParamFromDictKeyNo(parentCode);
            if (parent != null)
            {
                var param = AppMemCache.GetParamFromDictKeyNo(field.CodeLevelArr.Last());
                _emrDocumentHelper.CheckRadioOrCheckbox(param, parent, field);
            }

        }
        #endregion

        public bool IsCloseEmr()
        {
            var emr = _entity.MainObject as MEEmrsInfo;
            if (emr.MEEmrStatus == EmrStatus.Closed.ToString()) return true;

            return false;
        }

        public bool IsCloseOrWaitCloseEmr()
        {
            var emr = _entity.MainObject as MEEmrsInfo;
            if (emr.MEEmrStatus == EmrStatus.Closed.ToString() || emr.MEEmrStatus == EmrStatus.WaitClose.ToString()) return true;

            return false;
        }
        #region header and footer
        private void DeleteAllPageKeepFirst()
        {
            _emrDocumentHelper.DeleteAllPageKeepFirst(_richEditCtrl.Document, _richEditCtrl.DocumentLayout);

        }
        private void AutoAddHeaderAndFooter()
        {
            _emrDocumentHelper.AutoAddHeaderAndFooter(_richEditCtrl.Document, _richEditCtrl.DocumentLayout, _richEditCtrl.LayoutUnit);
        }
        #endregion

        public void TakeInitPermission()
        {
            var emr = _entity.MainObject as MEEmrsInfo;
            emr = _emrCtrl.GetObjectByID(emr.MEEmrID) as MEEmrsInfo;
            if (emr.MEEmrStatus == EmrStatus.Initing.ToString())
            {
                if (MessageBox.Show("Trước khi thực hiện thao tác này.\nBạn cần đảm bảo rằng không có Background Job đang khởi tạo bệnh án này.", "Lấy quyền khởi tạo", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
                {
                    emr.MEEmrStatus = EmrStatus.InProgress.ToString();
                    emr.AAUpdatedUser = BOSApp.CurrentUser;
                    _emrCtrl.UpdateObject(emr);
                    HistoryEmr(emr, cstObjectHistoryActionChange, $"thay đổi thông tin bệnh án - lấy quyền khởi tạo.");
                    MessageBox.Show("Đã lấy lại quyền khởi tạo thành công.", "Lấy quyền khởi tạo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Invalidate(emr.MEEmrID);
                }

            }
        }
        internal void DeleteTransferHistoryList()
        {
            if (BOSApp.CurrentUserGroupInfo.ADUserGroupRole != UserGroupRole.admin.ToString())
            {
                MessageBox.Show("Chỉ Admin mới có quyền thực hiện chức năng này.", "Thiếu quyền", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            _entity.MEEmrTranfersList.RemoveSelectedRowObjectFromList();
        }

        #region Thay doi loai benh an / Mo lai benh an da dong / Thay doi gay benh an
        public void ChangeEmrType()
        {
            if (IsEmrReadOnly(true))
            {
                MessageBox.Show($"Bệnh án ĐÃ ĐÓNG. Không thể thực hiện chức năng này.", "Không thể thực hiện chức năng này", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            };

            var emr = _entity.MainObject as MEEmrsInfo;
            //show dialog and map changes
            var gui = new guiChangeEmrType(emr.MEEmrID)
            {
                Module = this
            };
            gui.InitializeControls(gui.Controls);
            gui.StartPosition = FormStartPosition.CenterParent;
            if (gui.ShowDialog() == DialogResult.OK)
            {
                BOSProgressBar.Start("Đang lưu dữ liệu");
                try
                {
                    foreach (var document in gui.LastDocuments)
                    {
                        var newGrpID = gui.MappingList.Where(m => m.MEEmrTypeTemplateGroup == document.MEEmrDocumentGroup).First().MEEmrTypeTemplateGroupNewID;
                        var newGroup = gui.TemplateIndexs.Where(g => g.METemplateIndexID == newGrpID).FirstOrDefault();
                        if (newGroup != null)
                        {
                            document.MEEmrDocumentGroup = newGroup.METemplateIndexName;
                            document.MEEmrDocumentOrder = newGroup.METemplateIndexOrder;
                        }
                        else
                        {
                            document.MEEmrDocumentGroup = "Khác";
                            document.MEEmrDocumentOrder = 999;
                        }
                        document.AAUpdatedUser = BOSApp.CurrentUser;
                        _emrDocumentCtrl.UpdateObject(document);

                        //var documentCodeU = getDocumentNo(document);
                        //HistoryEmr(GetCurrentMainObject(), "Change", $"{documentCodeU} cập nhật thông tin loại bệnh án.");
                        //TODO considering update mongo
                    }
                    emr = _emrCtrl.GetObjectByID(emr.MEEmrID) as MEEmrsInfo;
                    emr.AAUpdatedUser = BOSApp.CurrentUser;
                    var emrType = _emrTypeCtrl.GetObjectByID(gui.FK_MEEmrTypeID) as MEEmrTypesInfo;
                    var changeMsg = $"Id cũ {emr.FK_MEEmrTypeID} - mới {gui.FK_MEEmrTypeID}";
                    emr.FK_MEEmrTypeID = gui.FK_MEEmrTypeID;
                    emr.MEEmrTypeProfile = emrType.MEEmrTypeProfile;
                    _emrCtrl.UpdateObject(emr);

                    HistoryEmr(emr, cstObjectHistoryActionChange, $"thay đổi thông tin bệnh án - loại bệnh án {changeMsg}.");

                    BOSProgressBar.Close();
                    Invalidate(emr.MEEmrID);
                    MessageBox.Show($"Loại bệnh án đã đổi thành công", "Đổi loại bệnh án thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    if (_notificationTab)
                    {
                        _msgLogs.Text += "\r\n CHANGE EMR TYPE FAILED: " + ex.ToString();
                    }
                    else if (_logConfig)
                    {
                        _msgLogsTemp += "\r\n CHANGE EMR TYPE FAILED: " + ex.ToString();
                    }
                    _msgNotification.Text = "Đổi loại bệnh án không thành công. Xem chi tiết ở thông báo.";
                }
                finally
                {
                    BOSProgressBar.Close();
                }

            }
        }
        public void OpenTheClosedEmr()
        {
            if (!IsEmrReadOnly(true))
            {
                MessageBox.Show($"Bệnh án ĐANG MỞ. Không thể thực hiện chức năng này.", "Không thể thực hiện chức năng này", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            };
            guiSign guiSign = new guiSign("Đang thực hiện mở lại bệnh án. Vẫn thực hiện?", _fingerPrintZK);
            if (guiSign.ShowDialog() == DialogResult.OK)
            {
                var gui = new guiGetOneStr("Nhập lý do vì sao mở lại bệnh án này.");
                if (gui.ShowDialog() != DialogResult.OK) return;
                string remark = gui.Value.Trim();
                BOSProgressBar.Start("Đang mở bệnh án");
                try
                {
                    var emr = _entity.MainObject as MEEmrsInfo;
                    emr = _emrCtrl.GetObjectByID(emr.MEEmrID) as MEEmrsInfo;
                    emr.AAUpdatedUser = BOSApp.CurrentUser;
                    emr.MEEmrStatus = EmrStatus.InProgress.ToString();
                    emr.MEEmrDesc = $"{emr.MEEmrDesc} [{DateTime.Now.ToString("dd/MM/yyyy HH:mm")}] {BOSApp.CurrentEmployeesInfo.HREmployeeName} mở lại bệnh án (Lý do: {remark})";
                    _emrCtrl.UpdateObject(emr);
                    HistoryEmr(emr, cstObjectHistoryActionReOpen, $"Mở lại bệnh án đóng.");
                    #region TTBA
                    if (_sysHelper.AllowThread())
                    {
                        var thUpdateEmrSumWhenOpenEmr = new System.Threading.Thread(() => UpdateEmrSumWhenOpenEmr(emr));
                        thUpdateEmrSumWhenOpenEmr.Start();
                    }
                    else
                    {
                        UpdateEmrSumWhenOpenEmr(emr);
                    }
                    #endregion
                    BOSProgressBar.Close();
                    Invalidate(emr.MEEmrID);
                }
                catch (Exception ex)
                {
                    if (_notificationTab)
                    {
                        _msgLogs.Text += "\r\n OPEN EMR FAILED: " + ex.ToString();
                    }
                    else if (_logConfig)
                    {
                        _msgLogsTemp += "\r\n OPEN EMR FAILED: " + ex.ToString();
                    }
                    _msgNotification.Text = "Mở bệnh án không thành công. Xem chi tiết ở thông báo.";
                }
                finally
                {
                    BOSProgressBar.Close();
                }
            }
        }
        #endregion

        #region Current Editing User
        private MEEmrDocumentsInfo SetCurrentEditingUser(MEEmrDocumentsInfo document)
        {
            document.FK_EditingUserID = BOSApp.CurrentEmployeesInfo.HREmployeeID;
            document.MEEmrDocumentHoldMachineMac = _macAddress;
            document.MEEmrDocumentHoldMachineIp = _hostName + "/" + _ipAddress;
            document.MEEmrDocumentHoldFrom = DateTime.Now;
            return document;
        }

        private MEEmrDocumentsInfo ClearCurrentEditingUser(MEEmrDocumentsInfo document)
        {
            document.FK_EditingUserID = 0;
            document.MEEmrDocumentHoldMachineMac = string.Empty;
            document.MEEmrDocumentHoldMachineIp = string.Empty;
            document.MEEmrDocumentHoldFrom = DateTime.MaxValue;
            return document;
        }
        #endregion

        #region Digital Signature
        internal void DigitalSignEmrDocument()
        {
            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            if (string.IsNullOrEmpty(document.MEEmrDocumentFile))
            {
                MessageBox.Show("Cần mở tờ bệnh án trước khi thực hiện ký số.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            var emr = this._entity.MainObject as MEEmrsInfo;
            var isDocx = document.MEEmrDocumentFileExt == EmrDocumentFileExtention.docx.ToString();
            METemplatesInfo template = _entity.METemplate;
            if (isDocx)
            {
                if (IsEmrReadOnly(true)) return;
                if (this._richEditCtrl.Modified)
                {
                    var confirm = MessageBox.Show("Lưu thay đổi trước khi thực hiện ký", "Nội dung tờ bệnh án đã thay đổi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (IsForceRevokePermission()) return;

                template = _templateCtrl.GetObjectByID(document.FK_METemplateID) as METemplatesInfo;
                if (!CheckDataIntegrityFromLastDigitalSignDocx(document, template.METemplateName)) return;
            }

            guiSign guiSign = new guiSign("Đang thực hiện ký toàn tờ bệnh án bằng chứng thư số CA. Sẽ không thể hoàn tác. Thực hiện?", _fingerPrintZK);
            if (guiSign.ShowDialog() != DialogResult.OK) return;
            try
            {
                if (isDocx)
                {
                    DigitalSignDocxDocument(document, template);
                }
                else if (document.MEEmrDocumentFileExt == EmrDocumentFileExtention.pdf.ToString())
                {
                    DigitalSignPdfDocument(document);
                }
            }
            catch (SignerCustomException ex)
            {
                BOSProgressBar.Close();
                var notificationStr = _notificationTab ? " Xem chi tiết ở thông báo" : string.Empty;
                if (ex.Code == -1)
                {
                    MessageBox.Show(ex.Message, $"[KẾT QUẢ TỪ MÁY CHỦ KÝ SỐ CA] {notificationStr}", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(ex.Code + " - " + ex.Message, $"[LỖI TỪ MÁY CHỦ KÝ SỐ CA] {notificationStr}", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                if (_notificationTab)
                {
                    _msgLogs.Text += "\r\n DIGITAL SIGN FAILED: " + ex.Code + "-" + ex.ToString();
                }
                else if (_logConfig)
                {
                    _msgLogsTemp += "\r\n DIGITAL SIGN FAILED: " + ex.Code + "-" + ex.ToString();
                }
                if (isDocx) InvalidateDocument(document);
            }
            catch (Exception ex)
            {
                BOSProgressBar.Close();
                var notificationStr = _notificationTab ? " Xem chi tiết ở thông báo" : string.Empty;
                MessageBox.Show(ex.ToString(), $"Lỗi không xác định.{notificationStr}", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (_notificationTab)
                {
                    _msgLogs.Text += "\r\n DIGITAL SIGN FAILED: " + ex.ToString();
                }
                else if (_logConfig)
                {
                    _msgLogsTemp += "\r\n DIGITAL SIGN FAILED: " + ex.ToString();
                }
                if (isDocx) InvalidateDocument(document);
            }
            finally
            {
                BOSProgressBar.Close();
            }
        }

        private void CreateSignLocalDir(int emrID)
        {
            Directory.CreateDirectory(string.Format(@"{0}\Emr\Partials", _documentPath));
            Directory.CreateDirectory(string.Format(@"{0}\Emr\Partials\{1}", _documentPath, emrID));
        }
        private void DigitalSignDocxDocument(MEEmrDocumentsInfo document, METemplatesInfo template)
        {
            var docx = EmrDocumentFileExtention.docx.ToString();
            var doc = _richEditCtrl.Document;
            // chon vai tro ky ten
            doc.SelectAll();
            var signFields = GetSignerRoleFields(doc.Selections);
            if (signFields == null) return;
            if (signFields.Count > 0)
            {
                BOSProgressBar.Start("Đang chèn tên và chữ ký");
                PrintMgsLog("CHEN-CHU-KY", string.Empty);
                this.InsertSignature(signFields);
            }
            // luu file to benh an chinh sau khi ky
            string fileName = string.Format(@"{0}\Emr\{1}\{2}.docx", _documentPath, document.FK_MEEmrID, document.MEEmrDocumentFile);
            this._emrDocumentHelper.SignRange(doc.Range, true, "CA Digital Signature", "Ký bằng chứng thư số bởi");

            int signWidth = template.METemplateDgtSignatureWidth > 0 ? template.METemplateDgtSignatureWidth : 100;
            int signHeght = template.METemplateDgtSignatureHeight > 0 ? template.METemplateDgtSignatureHeight : 50;
            byte[] signImage = ResizeImage(BOSApp.CurrentEmployeesInfo.HREmployeeSignature, signWidth, signHeght);

            string signature = template.METemplateDgtSignatureImage ? Convert.ToBase64String(signImage) : string.Empty;
            // thuc hien ky so CA tren toan bo file
            BOSProgressBar.Start("Đang thực hiện ký số CA");
            string toFileName = $"{document.MEEmrDocumentFile}_digitalsigned{DateTime.Now.ToBinary()}";
            string toFileFullname = $"{toFileName}.{docx}";
            string toFilePath = string.Format(@"{0}\Emr\Partials\{1}\{2}", _documentPath, document.FK_MEEmrID, toFileFullname);
            CreateSignLocalDir(document.FK_MEEmrID);
            // clone ra mot file khac
            _richEditCtrl.SaveDocument(toFilePath, DocumentFormat.OpenXml);
            _tempRichEditCtrl.LoadDocument(toFilePath, DocumentFormat.OpenXml);
            RemoveAllTagForPrint(_tempRichEditCtrl, template);
            // do not RemoveAllHiddenDataForPrint 
            _emrDocumentHelper.RemoveAllParamMarkup(_tempRichEditCtrl.Document);
            RemoveAllCommentForPrint(_tempRichEditCtrl);
            _emrDocumentHelper.RemoveAllPermisionRanges(_tempRichEditCtrl.Document);
            _tempRichEditCtrl.SaveDocument(toFilePath, DocumentFormat.OpenXml);

            var byteContent = File.ReadAllBytes(toFilePath);
            _tempRichEditCtrl.CreateNewDocument(false);
            var base64Resp = _digitalSig.FileSignerVinCa(byteContent, docx,
               new DigitalSignature()
               {
                   userName = BOSApp.CurrentEmployeesInfo.HREmployeeName,
                   userFullName = BOSApp.CurrentEmployeesInfo.HREmployeeName,
                   userDesc = "",
                   signatureType = "3",
                   signatureName = BOSApp.CurrentEmployeesInfo.HREmployeeName,
                   base64Pdf = Convert.ToBase64String(byteContent),
                   base64Signature = signature,
                   dateSigned = DateTime.Now,
                   appId = BOSApp.CurrentUsersInfo.ADUserCaPasscode, //  Cryptographier.Decrypt(BOSApp.CurrentUsersInfo.ADUserCaPasscode),
                   secret = BOSApp.CurrentUsersInfo.ADUserCaIdentity,
                   pdfFileName = toFileFullname,
                   locations = null
               }); ;

            //// Ký theo kiểu VNPT 
            //var base64Resp = _digitalSig.FileSigner(byteContent, docx,
            //    new SignerParammeterBase()
            //    {
            //        SerialNumber = BOSApp.CurrentUsersInfo.ADUserCaIdentity,
            //        AgreementUUID = BOSApp.CurrentUsersInfo.ADUserCaIdentity,
            //        AuthorizeCode = Emr.Cryptographier.Decrypt(BOSApp.CurrentUsersInfo.ADUserCaPasscode),
            //        IdentityNumber = BOSApp.CurrentEmployeesInfo.HREmployeeIDNumber,
            //        FileName = toFileFullname
            //    });  
            var byteArr = Convert.FromBase64String(base64Resp);
            File.WriteAllBytes(toFilePath, byteArr);
            _ftpFileMng.UploadFile($"/Emr/Signed/{document.FK_MEEmrID}/", toFileFullname, toFilePath);

            // neu bi loi se ko den buoc nay va khong luu
            BOSProgressBar.Start("Đang lưu tờ bệnh án gốc");
            _richEditCtrl.SaveDocument(fileName, DocumentFormat.OpenXml);
            if (!EncryptFileDocument()) return;
            if (!SaveEmrDocumentInfoAndUploadFile()) return;
            // tinh hash tren file goc
            var hash = _hashProvider.ComputeHash(fileName);

            var sign = new MEEmrDocumentSignsInfo()
            {
                FK_MEEmrDocumentID = document.MEEmrDocumentID,
                FK_HRDepartmentID = BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID,
                FK_HREmployeeID = BOSApp.CurrentEmployeesInfo.HREmployeeID,
                MEEmrDocumentSignTime = DateTime.Now,
                MEEmrDocumentSignFile = toFileName,
                MEEmrDocumentSignFileExt = docx,
                MEEmrDocumentSignType = EmrDocumentSignType.DigitalSigned.ToString(),
                MEEmrDocumentSignRemark = "Ký bằng chứng thư số",
                MEEmrDocumentSignHash = hash, // hash nay dc tinh truoc khi goi ky CA
                MEEmrDocumentSignUser = BOSApp.CurrentUser,
                MEEmrDocumentSignBlockAddr = string.Empty,
                AACreatedUser = BOSApp.CurrentUser
            };
            this._emrDocumentSignCtrl.CreateObject(sign);
            CloneSignedRangeForTracking();
            //Clear history để không đc undo.
            ((DevExpress.XtraRichEdit.Model.DocumentModel)this._richEditCtrl.Model).History.Clear();
            _richEditCtrl.Modified = false;
            _msgNotification.Text = "Ký bằng chứng thư số CA thành công. Xem lịch sử và nội dung ký ở tab Lịch sử ký tên";
        }
        private byte[] ResizeImage(byte[] imageBytes, int width, int height)
        {
            //using (var inputStream = new MemoryStream(imageBytes))
            //using (var image = Image.FromStream(inputStream))
            //using (var destImage = new Bitmap(width, height))
            //{
            //    destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            //    using (var graphics = Graphics.FromImage(destImage))
            //    {
            //        graphics.CompositingMode = CompositingMode.SourceCopy;
            //        graphics.CompositingQuality = CompositingQuality.HighQuality;
            //        graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            //        graphics.SmoothingMode = SmoothingMode.HighQuality;
            //        graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

            //        using (var wrapMode = new ImageAttributes())
            //        {
            //            wrapMode.SetWrapMode(WrapMode.TileFlipXY);
            //            graphics.DrawImage(image, new System.Drawing.Rectangle(0, 0, width, height));
            //        }
            //    }

            //    using (var outputStream = new MemoryStream())
            //    {
            //        // Convert thành PNG giúp giữ trong suốt
            //        destImage.Save(outputStream, ImageFormat.Png);
            //        return outputStream.ToArray();
            //    }
            //}

            Image image = null;

            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                image = Image.FromStream(ms);
            }

            var destRect = new System.Drawing.Rectangle(0, 0, width, height);
            var destImage = new Bitmap(width, height);

            destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

            using (var graphics = Graphics.FromImage(destImage))
            {
                graphics.CompositingMode = CompositingMode.SourceCopy;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                using (var wrapMode = new ImageAttributes())
                {
                    wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, destRect, 0, 0, width, height, GraphicsUnit.Pixel, wrapMode);
                }
            }
            byte[] imageBytesReturn = null;
            using (MemoryStream ms2 = new MemoryStream())
            {
                destImage.Save(ms2, image.RawFormat);
                imageBytesReturn = ms2.ToArray();
            }
            return imageBytesReturn;

            //using (var inputStream = new MemoryStream(imageBytes))
            //using (var sourceImage = Image.FromStream(inputStream))
            //{
            //    // Tạo bitmap mới theo size mong muốn
            //    var newBitmap = new Bitmap(width, height);

            //    using (var graphics = Graphics.FromImage(newBitmap))
            //    {
            //        graphics.CompositingQuality = CompositingQuality.HighQuality;
            //        graphics.SmoothingMode = SmoothingMode.HighQuality;
            //        graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;

            //        // Vẽ ảnh đã resize
            //        graphics.DrawImage(sourceImage, 0, 0, width, height);
            //    }

            //    using (var outputStream = new MemoryStream())
            //    {
            //        // Lưu ra định dạng JPEG/PNG
            //        newBitmap.Save(outputStream, ImageFormat.Jpeg);
            //        return outputStream.ToArray();
            //    }
            //}
        }
        private void DigitalSignPdfDocument(MEEmrDocumentsInfo document)
        {
            string fileName = $"{document.MEEmrDocumentFile}.{EmrDocumentFileExtention.pdf}";
            string filePath = Path.Combine(_documentPath, "Emr", document.FK_MEEmrID.ToString(), fileName);
            if (!File.Exists(filePath))
            {
                _msgNotification.Text = ("Cần mở lại tờ bệnh án. File bệnh án pdf không tồn tại ở địa chỉ. " + filePath);
                return;
            }
            var template = _templateCtrl.GetObjectByID(document.FK_METemplateID) as METemplatesInfo;
            var reason = string.Empty;
            if (_digitalSig.WillShowReason())
            {
                var guiOneStr = new guiGetOneStr("Nhập lý do ký số CA", !string.IsNullOrEmpty(template.METemplateDgtSignatureReason) ? template.METemplateDgtSignatureReason : _digitalSig.GetDefaultReason());
                if (guiOneStr.ShowDialog() != DialogResult.OK) return;
                reason = guiOneStr.Value;
            }
            BOSProgressBar.Start("Đang thực hiện ký số CA");
            var page = template.METemplateDgtSignaturePage <= 0 ? 1 : template.METemplateDgtSignaturePage;

            // ky o trang cuoi tai lieu
            if (_pdfViewer.PageCount < template.METemplateDgtSignaturePage)
                page = _pdfViewer.PageCount;

            var byteContent = File.ReadAllBytes(filePath);
            var hash = _hashProvider.ComputeHash(byteContent);

            int signWidth = template.METemplateDgtSignatureWidth > 0 ? template.METemplateDgtSignatureWidth : 100;
            int signHeght = template.METemplateDgtSignatureHeight > 0 ? template.METemplateDgtSignatureHeight : 50;
            byte[] signImage = ResizeImage(BOSApp.CurrentEmployeesInfo.HREmployeeSignature, signWidth, signHeght);

            string signature = template.METemplateDgtSignatureImage ? Convert.ToBase64String(signImage) : string.Empty;

            var configCA = BOSApp.GetSystemConfigValue(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.SYSTEM_CONFIGS_CA_METHOD);
            var methodCA = !string.IsNullOrEmpty(configCA) ? configCA : string.Empty;
            string transactionDesc = string.Empty;

            if (BOSApp.CurrentCompanyInfo.CSCompanyCaProvider == CaProviders.VNPT_CA)
            {
                if (_entity.MEPatient != null)
                    transactionDesc = $"{_entity.MEPatient.MEPatientNo} - {template.METemplateNo}";
                else
                    transactionDesc = $"{template.METemplateNo}";
                BOSProgressBar.SetText("Vui lòng xử lý tiếp trên ứng dụng VNPT CA...");
            }
            List<DigitalSignatureLocation> locations = new List<DigitalSignatureLocation>();

            DigitalSignatureLocation digitalSignature = new DigitalSignatureLocation();
            List<DigitalSignatureRect> lstRect = new List<DigitalSignatureRect>();
            DigitalSignatureRect rect = new DigitalSignatureRect();
            rect.StartX = template.METemplateDgtSignatureX;
            rect.StartY = template.METemplateDgtSignatureY;
            rect.EndX = template.METemplateDgtSignatureX + signWidth;
            rect.EndY = template.METemplateDgtSignatureY - signHeght;

            lstRect.Add(rect);
            digitalSignature.pageSign = page;
            digitalSignature.lstRect = lstRect;
            locations.Add(digitalSignature);
            var base64Resp = _digitalSig.PdfSignerVinCa(byteContent, "PDF", new DigitalSignature()
            {
                userName = BOSApp.CurrentEmployeesInfo.HREmployeeName,
                userFullName = BOSApp.CurrentEmployeesInfo.HREmployeeName,
                userDesc = "",
                signatureType = "3",
                signatureName = BOSApp.CurrentEmployeesInfo.HREmployeeName,
                base64Pdf = Convert.ToBase64String(byteContent),
                base64Signature = signature,
                dateSigned = DateTime.Now,
                appId = Cryptographier.Decrypt(BOSApp.CurrentUsersInfo.ADUserCaPasscode),
                secret = BOSApp.CurrentUsersInfo.ADUserCaIdentity,
                pdfFileName = fileName,
                locations = locations
            }); ;

            //// Ký theo VNPT  
            //var base64Resp = _digitalSig.PdfSigner(byteContent,
            //new SignerParammeterBase()
            //{
            //    SerialNumber = BOSApp.CurrentUsersInfo.ADUserCaIdentity,
            //    AgreementUUID = BOSApp.CurrentUsersInfo.ADUserCaIdentity,
            //    AuthorizeCode = Emr.Cryptographier.Decrypt(BOSApp.CurrentUsersInfo.ADUserCaPasscode),
            //    IdentityNumber = BOSApp.CurrentEmployeesInfo.HREmployeeIDNumber,
            //    FileName = BOSApp.CurrentCompanyInfo.CSCompanyCaProvider == CaProviders.VNPT_CA ? transactionDesc : fileName,
            //    Method = methodCA
            //},
            //new PdfSignerPropertyBase()
            //{
            //    Visible = template.METemplateDgtSignatureVisible,
            //    Width = template.METemplateDgtSignatureWidth,
            //    Height = template.METemplateDgtSignatureHeight,
            //    Reason = reason,
            //    Page = page,
            //    CoordinateX = template.METemplateDgtSignatureX, // góc Dưới - Trái
            //    CoordinateY = template.METemplateDgtSignatureY, // góc Dưới - Trái
            //    SignatureImage = signature,
            //    TextColor = template.METemplateDgtSignatureTextColor,
            //    FontSize = template.METemplateDgtSignatureFontSize,
            //});
            if (BOSApp.CurrentCompanyInfo.CSCompanyCaProvider == CaProviders.VNPT_CA)
            {
                BOSProgressBar.SetText("Đang xử lý kết quả từ VNPT CA...");
            }
            var byteArr = Convert.FromBase64String(base64Resp);
            //_msgLogs.Text += $"\r\n Result API CA: {base64Resp}";
            if (BOSApp.CurrentCompanyInfo.CSCompanyCaProvider == CaProviders.VNPT_CA)
            {
                BOSProgressBar.SetText("Đang lưu tài liệu đã ký số...");
            }
            File.WriteAllBytes(filePath, byteArr);
            _ftpFileMng.UploadFile($"/Emr/{document.FK_MEEmrID}/", fileName, filePath);

            var sign = new MEEmrDocumentSignsInfo()
            {
                FK_MEEmrDocumentID = document.MEEmrDocumentID,
                FK_HRDepartmentID = BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID,
                FK_HREmployeeID = BOSApp.CurrentEmployeesInfo.HREmployeeID,
                MEEmrDocumentSignTime = DateTime.Now,
                MEEmrDocumentSignFile = document.MEEmrDocumentFile,
                MEEmrDocumentSignFileExt = EmrDocumentFileExtention.pdf.ToString(),
                MEEmrDocumentSignType = EmrDocumentSignType.DigitalSigned.ToString(),
                MEEmrDocumentSignRemark = reason,
                MEEmrDocumentSignHash = hash, // hash nay dc tinh truoc khi goi ky CA
                MEEmrDocumentSignUser = BOSApp.CurrentUser,
                MEEmrDocumentSignBlockAddr = string.Empty,
                AACreatedUser = BOSApp.CurrentUser
            };
            this._emrDocumentSignCtrl.CreateObject(sign);

            using (MemoryStream stream = new MemoryStream(byteArr))
            {
                _pdfViewer.DetachStreamAfterLoadComplete = true;
                _pdfViewer.LoadDocument(stream);
            }
            BOSProgressBar.Close();
            MessageBox.Show("Ký bằng chứng thư số CA thành công. Xem lịch sử và nội dung ký ở tab Lịch sử ký tên", "Ký bằng chứng thư số CA thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }



        /// <summary>
        /// Kiem tra to benh an co thay doi 
        /// </summary>
        /// <returns></returns>
        private bool CheckDataIntegrityFromLastDigitalSignDocx(MEEmrDocumentsInfo document, string METemplateName)
        {
            var lastDigSign = _emrDocumentSignCtrl.GetLastDigitalSignDocx(document.MEEmrDocumentID);
            if (lastDigSign == null) return true;

            string fileName = string.Format(@"{0}\Emr\{1}\{2}.docx", _documentPath, document.FK_MEEmrID, document.MEEmrDocumentFile);
            var currentHash = _hashProvider.ComputeHash(fileName);
            if (currentHash != lastDigSign.MEEmrDocumentSignHash)
            {
                var result = MessageBox.Show($"NỘI DUNG TỜ BỆNH ÁN ĐÃ BỊ THAY ĐỔI KỂ TỪ LẦN KÝ SỐ GẦN NHẤT"
                    + $"\n\n \u2756 Tờ bệnh án: {document.MEEmrDocumentGroup}/ {document.MEEmrDocumentSubOrder}. {METemplateName}"
                    + "\n \u2756 Ký bởi: " + lastDigSign.MEEmrDocumentSignUser
                    + "\n \u2756 Ký lúc: " + lastDigSign.MEEmrDocumentSignTime.ToString("dd/MM/yyyy HH:mm:ss")
                    + "\n \u2756 Mã băm cũ: " + lastDigSign.MEEmrDocumentSignHash
                    + "\n\n \u2756 Mã băm hiện tại: " + currentHash
                    + "\n\n Yes: Để tiếp tục, No: Hủy",
                    "Nội dung tờ bệnh án đã bị thay đổi kể từ lần ký số gần nhất", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                return result == DialogResult.Yes;
            }
            return true;
        }
        #endregion

        #region User UI config, Quick Access Toolbar
        /// <summary>
        /// Move _msgNotification to parent screen, needing a space for Quick Access Toolbar
        /// </summary>
        private void SetupNotificationControl()
        {
            this._msgNotification = this.Controls[_fld_lbl_RichEditMsgName] as LabelControl;
            this._msgNotification.Location = new Point(351, 3);
            this.ParentScreen.ScreenContainer.Controls.Add(this._msgNotification);
            this._msgNotification.Width = this.ParentScreen.ScreenContainer.Width - 265;
            this._msgNotification.BringToFront();
        }
        private string GetQuickAccessToolbarConfigBase64()
        {
            var ribbon = this.Controls["richEditRibbonControl"] as RibbonControl;
            using (var mem = new MemoryStream())
            {
                ribbon.Toolbar.SaveLayoutToStream(mem);
                return Convert.ToBase64String(mem.ToArray());
            }
        }

        private void LoadUserUIConfig()
        {
            var ribbon = this.Controls["richEditRibbonControl"] as RibbonControl;
            var config = BOSApp.GetQuickAccessToolbarConfig();
            if (config != null)
            {
                var bytes = Convert.FromBase64String(config.ADUserConfigValue);
                using (var mem = new MemoryStream(bytes))
                {
                    ribbon.Toolbar.RestoreLayoutFromStream(mem);
                }
            }
            config = BOSApp.GetAutoHideRibbonConfig();
            ribbon.Minimized = (config?.ADUserConfigValue == "TRUE");

            config = BOSApp.GetNotificationTabConfig();
            _notificationTab = (config?.ADUserConfigValue == "TRUE");

            LoadMoveBetweenTagByTabBtnConfig();

            var edit = new DevExpress.XtraEditors.Repository.RepositoryItemTextEdit
            {
                AllowFocused = false,
                BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
            };
            edit.Appearance.ForeColor = Color.DarkRed;
            edit.Appearance.BackColor = Color.Transparent;
            //edit.Appearance.Font = new Font("Tahoma", 8.25F, FontStyle.Bold);
            edit.Appearance.Font = new Font("Tahoma", 7.5F, FontStyle.Bold);
            _qatNotificationMsg = new BarEditItem
            {
                Caption = "",
                Edit = edit,
                Id = 9999,
                Name = "qatNotificationMsg",
                AutoFillWidthInMenu = DevExpress.Utils.DefaultBoolean.True,
                AutoFillWidth = true,
                EditWidth = this.ParentScreen.ScreenContainer.Width - 250, //ribbon.Width / 3,
                Enabled = false,
                EditValue = ""
            };
            //ribbon.Manager.Items.Add(_qatNotificationMsg);
            ribbon.Toolbar.ItemLinks.Add(_qatNotificationMsg);
        }
        private void SaveUserConfigs()
        {
            SaveQuickAccessToolbarConfig();
            SaveAutoHideRibbonConfig();
            BOSApp.GetAllUserConfigsAsync(BOSApp.CurrentUsersInfo.ADUserID);
        }
        private void SaveQuickAccessToolbarConfig()
        {
            var config = BOSApp.GetQuickAccessToolbarConfig();
            if (config == null)
            {
                config = new ADUserConfigsInfo()
                {
                    FK_ADUserID = BOSApp.CurrentUsersInfo.ADUserID,
                    AAStatus = "Alive",
                    IsActive = true,
                    ADUserConfigGroup = UserCfgConsts.EMR_MODULE_GROUP,
                    ADUserConfigKey = UserCfgConsts.EMR_QUICK_ACCESS_TOOLBAR,
                    ADUserConfigValue = GetQuickAccessToolbarConfigBase64(),
                    ADUserConfigText = "Cấu hình Quick Access Toolbar theo người dùng",
                    ADUserConfigDesc = "Cấu hình Quick Access Toolbar trong màn hình soạn thảo bệnh án"
                };
                _userConfigCtrl.CreateObject(config);
            }
            else
            {
                config.ADUserConfigValue = GetQuickAccessToolbarConfigBase64();
                _userConfigCtrl.UpdateObject(config);
            }
        }
        private void SaveAutoHideRibbonConfig()
        {
            var ribbon = this.Controls["richEditRibbonControl"] as RibbonControl;
            var configAutoHide = BOSApp.GetAutoHideRibbonConfig();
            var autoHide = ribbon.Minimized.ToString().ToUpper();
            if (configAutoHide == null)
            {
                configAutoHide = new ADUserConfigsInfo()
                {
                    FK_ADUserID = BOSApp.CurrentUsersInfo.ADUserID,
                    AAStatus = "Alive",
                    IsActive = true,
                    ADUserConfigGroup = UserCfgConsts.EMR_MODULE_GROUP,
                    ADUserConfigKey = UserCfgConsts.EMR_MINIMIZE_RIBBON,
                    ADUserConfigValue = autoHide,
                    ADUserConfigText = "Cấu hình Auto Hide Ribbon theo người dùng",
                    ADUserConfigDesc = "Cấu hình Auto Hide Ribbon trong màn hình soạn thảo bệnh án"
                };
                _userConfigCtrl.CreateObject(configAutoHide);
            }
            else
            {
                configAutoHide.ADUserConfigValue = autoHide;
                _userConfigCtrl.UpdateObject(configAutoHide);
            }

        }
        public bool LoadMoveBetweenTagByTabBtnConfig()
        {
            _moveBetweenTagByTabEnabled = BOSApp.GetBooleanUserConfig(UserCfgConsts.MOVE_BETWEEN_TAG_BY_TAB, true);
            return _moveBetweenTagByTabEnabled;
        }
        public bool LoadClickThenAutoMoveToNextTagConfig()
        {
            _clickAndAutoMoveToNextTagConfig = BOSApp.GetBooleanUserConfig(UserCfgConsts.CLICK_THEN_MOVE_TO_NEXT_TAG, false);
            return _clickAndAutoMoveToNextTagConfig;
        }

        public bool LoadHighlightEmrTagUserConfig()
        {
            _highlightEmrTagUserConfig = BOSApp.GetBooleanUserConfig(UserCfgConsts.HIGHLIGHT_EMR_TAG, true);
            if (_richEditCtrl != null)
                _richEditCtrl.Options.Fields.HighlightMode = _highlightEmrTagUserConfig ? FieldsHighlightMode.Always : FieldsHighlightMode.Never;
            return _highlightEmrTagUserConfig;
        }
        #endregion

        #region Emr Archives
        private void CreateArchiveDir(int emrId)
        {
            var localDir = Path.Combine(_documentPath, _emrArchiveHelper.StorageDir, emrId.ToString());
            if (!Directory.Exists(localDir))
                Directory.CreateDirectory(localDir);
        }
        internal void InvalidateEmrArchives()
        {
            var emr = _entity.MainObject as MEEmrsInfo;
            _entity.MEEmrArchiveList.Invalidate(emr.MEEmrID);
        }

        internal void InvalidateGEObjectHistory()
        {
            var emr = _entity.MainObject as MEEmrsInfo;
            var ds = _geObjHistoryCtrl.GetListBusinessObjects<GEObjectHistoryInfo>(_geObjHistoryCtrl.GetGEObjectHistoryByObjectNameAndObjectID(TableName.MEEmrsTableName, emr.MEEmrID));
            var gridControlEmrRelations = this.Controls["fld_dgcGEObjectHistory"] as GEObjectHistoryGridControl;
            if (gridControlEmrRelations != null)
            {
                gridControlEmrRelations.DataSource = ds;
                gridControlEmrRelations.RefreshDataSource();
                gridControlEmrRelations.Refresh();
            }
        }

        public void ArchiveEmr()
        {
            if (!IsEmrReadOnly(true))
            {
                MessageBox.Show($"Bệnh án CHƯA ĐÓNG. Không thể thực hiện chức năng này.", "Không thể thực hiện chức năng này", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            };
            try
            {
                var emr = _entity.MainObject as MEEmrsInfo;
                _entity.MEEmrArchiveList.Invalidate(emr.MEEmrID);
                if (_entity.MEEmrArchiveList.Count > 0)
                    if (MessageBox.Show($"Bệnh án này đã được lưu trữ trước đó. \n\n Vẫn tiếp tục tạo bảng lưu trữ mới?", "Đã tồn tại bộ lưu trữ",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                    {
                        return;
                    }

                if (string.IsNullOrEmpty(emr.MEEmrArchiveNo))
                    if (MessageBox.Show($"SỐ LƯU TRỮ đang để trống. \n\n Vẫn tiếp tục?", "Chưa nhập Mã lưu trữ",
                          MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                    {
                        return;
                    }

                if (emr.MEEmrDateOut == null || emr.MEEmrDateOut.Year == 9999)
                    if (MessageBox.Show($"NGÀY RA VIỆN đang để trống. \n\n Vẫn tiếp tục?", "Chưa nhập Ngày ra viện",
                          MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                    {
                        return;
                    }

                BOSProgressBar.Start("Đang xử lý dữ liệu");
                CreateArchiveDir(emr.MEEmrID);
                InvalidateEmrDocumentListForAdmin(emr);
                var archive = _emrArchiveHelper.MergeAndArchived(emr, _entity.MEEmrDocumentsList);

                if (archive == null)
                {
                    MessageBox.Show($"Bệnh án không có tờ nào để lưu trữ.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }

                _entity.MEEmrArchiveList.Add(archive);
                _entity.MEEmrArchiveList.SaveItemObjects();
                _entity.MEEmrArchiveList.Invalidate(emr.MEEmrID);
                _entity.MEEmrArchiveList.GridView.FocusedRowHandle = 0;
                ViewArchivePdf(archive);
                MessageBox.Show("Đã lưu trữ thành công. \nChọn chức năng Ký số CA để thực hiện ký số nếu cần.", "Lưu trữ thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xảy ra khi xử lý dữ liệu. \n" + ex.ToString(), "Có lỗi xảy ra", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                BOSProgressBar.Close();
            }
        }
        internal void ViewArchivePdf(MEEmrArchivesInfo archive)
        {
            if (archive == null) return;
            try
            {
                BOSProgressBar.Start("Đang tải và mở tập tin");
                var emr = _entity.MainObject as MEEmrsInfo;
                CreateArchiveDir(emr.MEEmrID);
                var dotExt = "." + archive.MEEmrArchiveFileExt;
                string fileName = archive.MEEmrArchiveFile + dotExt;
                string filePath = Path.Combine(_documentPath, _emrArchiveHelper.StorageDir, emr.MEEmrID.ToString(), fileName);
                if (_ftpFileMng.FileExists($"/{_emrArchiveHelper.StorageDir}/{emr.MEEmrID}/", fileName))
                {
                    _ftpFileMng.DownloadFile($"/{_emrArchiveHelper.StorageDir}/{emr.MEEmrID}/", fileName, filePath);
                }
                else
                {
                    _ftpFileMng.DownloadFile($"/Emr/{emr.MEEmrID}/", fileName, filePath); // backward compatible
                }

                OpenArchivePdfFile(filePath);

                // Bug 1526: Không check ký số khi user chưa cấu hình ký số khi bấm vào file lưu trữ đã ký số
                if (!string.IsNullOrEmpty(BOSApp.CurrentUsersInfo.ADUserCaIdentity))
                {
                    var resultViewDigitalSignatureInfo = _emrArchiveHelper.ViewDigitalSignatureInfo(archive, filePath, _logConfig, _notificationTab, _msgLogs);
                    if (_logConfig && !_notificationTab)
                    {
                        _msgLogsTemp += $"\r\n {resultViewDigitalSignatureInfo}";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xảy ra khi xử lý dữ liệu. \n" + ex.ToString(), "Có lỗi xảy ra", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                BOSProgressBar.Close();
            }
        }
        private void OpenArchivePdfFile(string filePath)
        {
            var pdfViewer = (Controls["pdfViewerEmrArchive"] as PdfViewer);
            using (FileStream stream = new FileStream(filePath, FileMode.Open))
            {
                pdfViewer.DetachStreamAfterLoadComplete = true;
                pdfViewer.LoadDocument(stream);
            }
        }
        public void DigitalSignEmrArchivePdf(MEEmrArchivesInfo archive)
        {
            if (archive == null) return;
            var pdfViewer = (Controls["pdfViewerEmrArchive"] as PdfViewer);
            var emr = _entity.MainObject as MEEmrsInfo;
            var patient = _patientCtrl.GetObjectByID(emr.FK_MEPatientID) as MEPatientsInfo;
            emr.MEPatientNo = patient.MEPatientNo;
            var dotExt = "." + archive.MEEmrArchiveFileExt;
            string fileName = archive.MEEmrArchiveFile + dotExt;
            string filePath = Path.Combine(_documentPath, _emrArchiveHelper.StorageDir, emr.MEEmrID.ToString(), fileName);
            if (!File.Exists(filePath))
            {
                _msgNotification.Text = ("Click để mở lại tập tin lưu trữ. Tập tin bệnh án .pdf không tồn tại ở địa chỉ. " + filePath);
                return;
            }
            var emrType = _emrTypeCtrl.GetObjectByID(emr.FK_MEEmrTypeID) as MEEmrTypesInfo;
            if (emrType == null) return;

            if (archive.MEEmrArchiveStatus == EmrArchiveStatus.DigitalSigned.ToString())
            {
                if (MessageBox.Show($"Tệp tin lưu trữ đã được ký số bởi {archive.AAUpdatedUser} lúc {archive.MEEmrArchiveSignTime}.\n Bạn vẫn muốn ký lại?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                    return;
            }
            var reason = string.Empty;
            if (_digitalSig.WillShowReason())
            {
                var guiOneStr = new guiGetOneStr("Nhập lý do ký số CA", !string.IsNullOrEmpty(emrType.MEEmrTypeDgtSignatureReason) ? emrType.MEEmrTypeDgtSignatureReason : _digitalSig.GetDefaultReason());
                if (guiOneStr.ShowDialog() != DialogResult.OK) return;
                reason = guiOneStr.Value;
            }
            try
            {
                BOSProgressBar.Start("Đang thực hiện ký số CA");
                var page = emrType.MEEmrTypeDgtSignaturePage <= 0 ? 1 : emrType.MEEmrTypeDgtSignaturePage;

                // ky o trang cuoi tai lieu
                if (pdfViewer.PageCount < emrType.MEEmrTypeDgtSignaturePage)
                    page = pdfViewer.PageCount;

                archive = _emrArchiveHelper.DigitalSignEmrArchivePdf(emr, archive, emrType, page, reason);

                _entity.MEEmrArchiveList.GridControl.RefreshDataSource();
                OpenArchivePdfFile(filePath);

                BOSProgressBar.Close();
                MessageBox.Show("Ký bằng chứng thư số CA thành công", "Ký bằng chứng thư số CA thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (SignerCustomException ex)
            {
                BOSProgressBar.Close();
                if (ex.Code == -1)
                {
                    MessageBox.Show(ex.Message, $"[KẾT QUẢ TỪ MÁY CHỦ KÝ SỐ CA]", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(ex.Code + " - " + ex.Message, $"[LỖI TỪ MÁY CHỦ KÝ SỐ CA]", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xảy ra khi xử lý dữ liệu. \n" + ex.ToString(), "Có lỗi xảy ra", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                BOSProgressBar.Close();
            }
        }
        #endregion

        #region Open Web Browser
        private void OpenWebBrowser(Dictionary<string, object> requestParams, MEEmrActionsInfo action)
        {
            foreach (var link in requestParams)
            {
                var url = link.Value.ToString();
                if (string.IsNullOrEmpty(url)) continue;

                if (!string.IsNullOrEmpty(action.MEEmrActionPlugin))
                    Process.Start(action.MEEmrActionPlugin, url);
                else
                    // open in default browser
                    Process.Start(url);
            }
            if (string.IsNullOrEmpty(action.MEEmrActionPlugin))
            {
                ShowFlashNotification("Đang mở bằng trình duyệt mặc định của máy. Thẻ chức năng này chưa được cấu hình trình duyệt.", 5000);
            }
            else
            {
                _msgNotification.Text = "Đã mở trình duyệt";
            }
        }
        public void OpenWebBrowser(string url)
        {
            var browser = BOSApp.GetSystemConfigValue(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.SYSTEM_CONFIGS_DEFAULT_BROWSER);
            if (!string.IsNullOrEmpty(browser))
                Process.Start(browser, url);
            else
                // open in default browser
                Process.Start(url);

            if (string.IsNullOrEmpty(browser))
            {
                ShowFlashNotification("Đang mở bằng trình duyệt mặc định của máy.", 5000);
            }
            else
            {
                _msgNotification.Text = "Đã mở trình duyệt";
            }
        }
        #endregion

        #region Kiem duyet Emr
        public void CheckupEmr()
        {
            if (this._richEditCtrl.ReadOnly)
            {
                MessageBox.Show($"Tờ bệnh án được mở ở chế độ CHỈ ĐỌC. Vui lòng mở lại tờ bệnh án trước khi thao tác.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!IsNothingToSaveDocumentContent()) return;
            var emr = _entity.MainObject as MEEmrsInfo;
            var docs = _emrDocumentCtrl.GetByEmrAndStatusCondition(emr.MEEmrID);
            foreach (var document in docs)
            {
                if (document.FK_EditingUserID > 0 && document.FK_EditingUserID != BOSApp.CurrentEmployeesInfo.HREmployeeID)
                {
                    ShowDocumentEditingUser(document);
                    return;
                }
            }
            var mainDocument = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            var actions = this._templateActionCtrl.GetAllByTemplateIdAndWhen(mainDocument.FK_METemplateID, EmrTemplateActionWhen.Checkup.ToString())
                .OrderBy(o => o.MEEmrTemplateActionOrder).ToList();
            if (actions.Count == 0)
            {
                MessageBox.Show($"Không tìm thấy chức năng tự chạy khi kiểm duyệt dành cho tờ bệnh án này. " +
                    $"\nVui lòng chọn đúng tờ bệnh án (vd: Tờ đầu bệnh án).",
                   "Vui lòng chọn đúng tờ bệnh án", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            BOSProgressBar.Start("Đang xuất PDF các tờ bệnh án. Có thể mất chút thời gian.");
            string localPath = string.Format(@"{0}\Emr\{1}\", _documentPath, emr.MEEmrID);
            DirectoryInfo localDir = new DirectoryInfo(localPath);
            var pdf = EmrDocumentFileExtention.pdf.ToString();
            var docx = EmrDocumentFileExtention.docx.ToString();
            try
            {
                var isActionLocal = true;
                var documentsDto = new List<MEEmrDocumentsDto>();
                if (_apiEmr != null)
                {
                    var actionUri = BOSApp.GetSystemConfigValue(SysCfgConsts.EMR_API_ENDPOINT, SysCfgConsts.EMR_CHECKUP);
                    if (!string.IsNullOrEmpty(actionUri))
                    {
                        PrintMgsLog("BAT-DAU-GOI-API-KIEM-DUYET", actionUri);
                        var body = new { emrId = emr.MEEmrID, userId = BOSApp.CurrentUsersInfo.ADUserID };
                        var response = _apiEmr.Post<Emr.Base.Models.Abp.AjaxResponse, List<MEEmrDocumentsDto>>(actionUri, null, body);
                        PrintMgsLog("KET-THUC-GOI-API-KIEM-DUYET", actionUri);
                        if (response != null && response.Success)
                        {
                            documentsDto = response.Result;
                            isActionLocal = false;
                        }
                        else
                        {
                            var msg = "Lỗi không xác định.";
                            if (response.Error != null)
                            {
                                var document = _entity.MEEmrDocumentsList.Where(d => d.MEEmrDocumentFile == response.Error.Message).FirstOrDefault();
                                if (document != null)
                                {
                                    //TODO lấy số STT tờ theo gridview
                                    var template = _templateCtrl.GetObjectByID(document.FK_METemplateID) as METemplatesInfo;
                                    msg = $"Gáy: [{document.MEEmrDocumentOrder}.{document.MEEmrDocumentGroup}]\nTờ: [{document.MEEmrDocumentSubOrder}.{template.METemplateName}. {document.MEEmrDocumentFile}]";
                                    msg += $"\n\n\nChi tiết: {response.Error?.Details.ToString()}";
                                }
                            }
                            //MessageBox.Show($"Có lỗi khi xử lý kiểm duyệt bệnh án.\n\n" + msg, "KHÔNG THỂ THỰC HIỆN KIỂM DUYỆT BẰNG [EMR API]", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            if (MessageBox.Show("Tiếp tục kiểm duyệt bệnh án trên [MÁY CLIENT NÀY]? \nKiểm duyệt bệnh án thất bại trên [MÁY CHỦ EMR API].\n\nOK: Để tiếp tục \nCancel: Hủy thao tác",
                        "THỰC HIỆN KIỂM DUYỆT TRÊN MÁY CLIENT NÀY", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.Cancel)
                                return;
                        }
                    }
                    else
                    {
                        //ShowFlashNotification("Không cấu hình kiểm duyệt bệnh án bằng EMR API. Đang kiểm duyệt ở [máy client này]", 3000);
                    }
                }
                else
                {
                    //ShowFlashNotification("Không kết nối được đến máy chủ EMR API. Đang kiểm duyệt ở [máy client này]", 3000);
                }

                if (isActionLocal)
                {
                    string serverPath = $"/Emr/{emr.MEEmrID}/";
                    var serverFiles = _ftpFileMng.GetNameListing(serverPath);
                    var processBgFiles = new List<string>();
                    foreach (var document in _entity.MEEmrDocumentsList)
                    {
                        try
                        {
                            if (document.MEEmrDocumentStatus == EmrDocumentStatus.InProgress.ToString() || document.MEEmrDocumentStatus == EmrDocumentStatus.Closed.ToString())
                            {
                                string filePdfPath = Path.Combine(localPath, $"{document.MEEmrDocumentFile}.{pdf}");
                                if (document.MEEmrDocumentFileExt == docx)
                                {
                                    var localFileDocx = DownloadFtpFile(document);
                                    var hash = _md5Hasher.ComputeHash(localFileDocx);
                                    var filePdfHash = $"{document.MEEmrDocumentFile}_MD5{hash}.{pdf}";
                                    File.Delete(filePdfPath);
                                    var fi = new FileInfo(Path.Combine(localPath, filePdfHash));
                                    if (fi != null && fi.Exists && fi.Length > 10000)
                                    {
                                        File.Move(Path.Combine(localPath, filePdfHash), filePdfPath);
                                        //khong up len lai server
                                        //processBgFiles.Add(filePdfHash);
                                    }
                                    else
                                    {
                                        var existFileHash = serverFiles.Where(stringToCheck => stringToCheck.Contains(filePdfHash));
                                        if (existFileHash.Count() > 0)
                                        {
                                            //co file hash tren server thi tai ve
                                            _ftpFileMng.DownloadFile(serverPath, filePdfHash, filePdfPath);
                                            fi = new FileInfo(filePdfPath);
                                            if (fi != null && fi.Exists && fi.Length < 10000)
                                                File.Delete(filePdfPath);
                                        }
                                    }

                                    try
                                    {
                                        //neu file pdf co loi thi se xuat lai
                                        document.MEEmrDocumentPageCount = GetNumberOfPdfPages(filePdfPath);
                                    }
                                    catch (Exception)
                                    {
                                        using (OfficeOpenXmlCrypto.OfficeCryptoStream stream = OfficeOpenXmlCrypto.OfficeCryptoStream.Open(localFileDocx, this._emrDocumentHelper.ShareEmrPassword))
                                        {
                                            _tempRichEditCtrl.LoadDocument(stream, DocumentFormat.OpenXml);
                                        }
                                        RemoveAllForPrint(this._tempRichEditCtrl, document.FK_METemplateID);

                                        _tempRichEditCtrl.ExportToPdf(filePdfPath);

                                        document.MEEmrDocumentPageCount = GetNumberOfPdfPages(filePdfPath);
                                        if (!processBgFiles.Contains(filePdfHash))
                                        {
                                            processBgFiles.Add(filePdfHash);
                                        }
                                    }

                                    ClearHashFileLocal(localDir, document.MEEmrDocumentFile);
                                    //trong tat ca cac truong hop deu giu lai file hash o local cho lan kiem duyet tiep theo neu co
                                    File.Copy(filePdfPath, Path.Combine(localPath, filePdfHash), true);
                                }
                                else if (document.MEEmrDocumentFileExt == pdf)
                                {
                                    _ftpFileMng.DownloadFile(serverPath, $"{document.MEEmrDocumentFile}.{pdf}", filePdfPath);
                                    document.MEEmrDocumentPageCount = GetNumberOfPdfPages(filePdfPath);
                                }
                                document.MEEmrDocumentDuplexCount = (int)Math.Ceiling(document.MEEmrDocumentPageCount / 2F);
                            }
                        }
                        catch (Exception ex)
                        {
                            var mirrorDocument = (MEEmrDocumentsInfo)document.Clone();
                            mirrorDocument.MEEmrDocumentDesc = $"{ex.ToString()}";
                            var noticeDocs = new List<MEEmrDocumentsInfo>
                            {
                                mirrorDocument
                            };
                            var guiNotice = new guiNotificationEmrDocument("KHÔNG THỂ THỰC HIỆN KIỂM DUYỆT - Có lỗi khi xử lý tờ bệnh án.", noticeDocs)
                            {
                                Module = this
                            };
                            guiNotice.ShowDialog();
                            return;
                        }
                    }

                    if (_sysHelper.AllowThread())
                    {
                        var thProcessHashFileFtp = new System.Threading.Thread(() => ProcessHashFileFtp(localPath, serverPath, processBgFiles));
                        thProcessHashFileFtp.Start();
                    }
                    else
                    {
                        ProcessHashFileFtp(localPath, serverPath, processBgFiles);
                    }
                }

                var mainFilePath = Path.Combine(_documentPath, "Emr", mainDocument.FK_MEEmrID.ToString(), mainDocument.MEEmrDocumentFile + "." + pdf);
                if (!isActionLocal)
                {
                    foreach (var document in _entity.MEEmrDocumentsList)
                    {
                        if (document.MEEmrDocumentStatus == EmrDocumentStatus.InProgress.ToString() || document.MEEmrDocumentStatus == EmrDocumentStatus.Closed.ToString())
                        {
                            var docDto = documentsDto.FirstOrDefault(m => m.MEEmrDocumentID.Equals(document.MEEmrDocumentID));
                            document.MEEmrDocumentPageCount = docDto.MEEmrDocumentPageCount;
                            document.MEEmrDocumentDuplexCount = (int)Math.Ceiling(document.MEEmrDocumentPageCount / 2F);
                        }
                        // Xoá pdf kiểm duyệt trước đó nếu có
                        if (document.MEEmrDocumentFileExt == docx)
                        {
                            string filePdfPath = Path.Combine(localPath, $"{document.MEEmrDocumentFile}.{pdf}");
                            File.Delete(filePdfPath);
                        }
                    }
                    #region Download files. Now download each file when click
                    //string localDir = Path.Combine(_documentPath, "Emr", emr.MEEmrID.ToString());
                    //var pathWp = Path.Combine($"\\Workspace", BOSApp.CurrentUsersInfo.ADUserID.ToString(), "Emr", emr.MEEmrID.ToString());
                    //var threadDownWPEmrsPdf = new System.Threading.Thread(() => DownWpEmrPdf(localDir, pathWp, Path.GetFileName(mainFilePath)));
                    //threadDownWPEmrsPdf.Start();
                    #endregion
                }

                var rowHandled = _entity.MEEmrDocumentsList.GridView.FocusedRowHandle;
                _entity.MEEmrDocumentsList.SaveItemObjects(); // TODO: Lost focus grid
                _entity.MEEmrDocumentsList.GridView.FocusedRowHandle = rowHandled;

                if (actions.Count > 0)
                {
                    if (mainDocument.MEEmrDocumentFileExt == EmrDocumentFileExtention.docx.ToString())
                    {
                        foreach (var act in actions)
                        {
                            var action = this._actionsController.GetObjectByID(act.FK_MEEmrActionID) as MEEmrActionsInfo;
                            if (action != null)
                            {
                                if (_checkSystem)
                                {
                                    _sysHelper.LogTxt("information", $"Bắt đầu chạy chức năng {action.MEEmrActionNo}.");
                                    var watchAct = Stopwatch.StartNew();
                                    CallEmrAction($"{action.MEEmrActionNo}{EmrParam.TagCodeSeparator}{EmrParam.GuidTag}={mainDocument.MEEmrDocumentGuid}", this._richEditCtrl.Document.Range);
                                    watchAct.Stop();
                                    var elapsedAct = watchAct.ElapsedMilliseconds / 1000.0;
                                    _sysHelper.LogTxt("information", $"{elapsedAct} giây. Hoàn tất chạy chức năng {action.MEEmrActionNo}.");
                                }
                                else
                                {
                                    CallEmrAction($"{action.MEEmrActionNo}{EmrParam.TagCodeSeparator}{EmrParam.GuidTag}={mainDocument.MEEmrDocumentGuid}", this._richEditCtrl.Document.Range);
                                }
                            }
                        }
                        var memStream = new MemoryStream(_richEditCtrl.Document.OpenXmlBytes);
                        _tempRichEditCtrl.LoadDocument(memStream, DocumentFormat.OpenXml);
                        RemoveAllForPrint(_tempRichEditCtrl, mainDocument.FK_METemplateID);
                        _tempRichEditCtrl.ExportToPdf(mainFilePath);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xảy ra khi xử lý dữ liệu. \n" + ex.ToString(), "KHÔNG THỂ THỰC HIỆN KIỂM DUYỆT", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            finally
            {
                BOSProgressBar.Close();
            }

            var gui = new guiCheckup(_entity.MEEmrDocumentsList.Clone() as BOSList<MEEmrDocumentsInfo>) { Module = this };
            gui.StartPosition = FormStartPosition.CenterParent;
            if (gui.ShowDialog() == DialogResult.OK)
            {
                var command = this._richEditCtrl.CreateCommand(RichEditCommandId.FileSave);
                command.Execute();
                MessageBox.Show("Đã lưu dữ liệu", "Lưu thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                // Undo change on document
                InvalidateDocument(mainDocument, false);
            }
        }

        private void ProcessHashFileFtp(string localPath, string serverPath, List<string> processBgFiles)
        {
            try
            {
                var serverFiles = _ftpFileMng.GetNameListing(serverPath);
                foreach (var hashFile in processBgFiles)
                {
                    var documentFile = Path.GetFileNameWithoutExtension(hashFile).Split(new string[] { "_MD5" }, StringSplitOptions.None)[0];
                    var fileLocalPath = Path.Combine(localPath, $"{documentFile}.pdf");
                    ClearHashFileServer(serverPath, serverFiles, documentFile, string.Empty);
                    UploadFileFtp(serverPath, hashFile, fileLocalPath);
                }
            }
            catch (Exception ex)
            {
                _sysHelper.LogTxt("error", ex.ToString());
            }
        }

        private static void ClearHashFileLocal(DirectoryInfo localDir, string mEEmrDocumentFile)
        {
            var checkMD5HashSt = $"{mEEmrDocumentFile}_MD5";
            foreach (FileInfo file in localDir.GetFiles())
            {
                if (file.Name.Contains(checkMD5HashSt))
                {
                    file.Delete();
                }
            }
        }

        private void ClearHashFileServer(string serverPath, string[] serverFiles, string mEEmrDocumentFile, string exceptFile)
        {
            var deleteFilePart = $"{mEEmrDocumentFile}_MD5";
            var deleteFiles = serverFiles.Where(m => m.Contains(deleteFilePart));
            if (!string.IsNullOrEmpty(exceptFile))
            {
                deleteFiles = deleteFiles.Where(m => !m.Contains(exceptFile));
            }
            foreach (var file in deleteFiles)
            {
                if (Path.GetExtension(file) == ".pdf")
                {
                    _ftpFileMng.Delete(serverPath, Path.GetFileName(file));
                }
            }
        }

        private void UploadFileFtp(string serverPath, string serverFile, string localPath)
        {
            try
            {
                _ftpFileMng.UploadFile(serverPath, serverFile, localPath);
            }
            catch (Exception ex)
            {
                PrintMgsLog("Có lỗi khi tải lên tờ bệnh án tạm. Lỗi không ảnh hưởng.", ex.ToString());
            }
        }

        internal void ViewLocalPdfFile(MEEmrDocumentsInfo document, PdfViewer pdfViewer)
        {
            string filePath = Path.Combine(_documentPath, "Emr", document.FK_MEEmrID.ToString(), document.MEEmrDocumentFile + "." + EmrDocumentFileExtention.pdf.ToString());
            if (!File.Exists(filePath))
            {
                var serverPath = $"/Emr/{document.FK_MEEmrID}/";
                var filesServer = _ftpFileMng.GetNameListing(serverPath);
                var checkUpFilePart = $"{document.MEEmrDocumentFile}_MD5";
                var checkUpFile = filesServer.FirstOrDefault(stringToCheck => stringToCheck.Contains(checkUpFilePart));
                if (checkUpFile != null)
                {
                    _ftpFileMng.DownloadFile(serverPath, Path.GetFileName(checkUpFile), filePath);
                }
                else
                {
                    if (_notificationTab)
                    {
                        _msgLogs.Text += "\r\n" + ("File bệnh án không tồn tại ở địa chỉ. " + filePath);
                    }
                    else if (_logConfig)
                    {
                        _msgLogsTemp += "\r\n" + ("File bệnh án không tồn tại ở địa chỉ. " + filePath);
                    }
                    return;
                }
            }
            using (FileStream stream = new FileStream(filePath, FileMode.Open))
            {
                pdfViewer.DetachStreamAfterLoadComplete = true;
                pdfViewer.LoadDocument(stream);
            }
        }
        private static int GetNumberOfPdfPages(string path)
        {
            using (var reader = new iTextSharp.text.pdf.PdfReader(path))
                return reader.NumberOfPages;
        }
        private void DownWpEmrPdf(string localDir, string remotePath, string exceptFile)
        {
            _ftpFileMng.DownloadFiles(localDir, remotePath, exceptFile);
        }
        #endregion

        #region Template Emr
        public void ChangeEmrTemplate()
        {
            if (IsEmrReadOnly(true))
            {
                MessageBox.Show($"Chức năng chỉ thực hiện với bệnh án đang mở.", "Không thể thực hiện chức năng này", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            };

            var emr = _entity.MainObject as MEEmrsInfo;

            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            if (document.MEEmrDocumentID == 0)
            {
                MessageBox.Show("Vui lòng chọn một tờ bệnh án", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var docInDb = _emrDocumentCtrl.GetObjectByID(document.MEEmrDocumentID) as MEEmrDocumentsInfo;

            if (docInDb.FK_EditingUserID > 0 && docInDb.MEEmrDocumentHoldMachineMac != _macAddress && docInDb.MEEmrDocumentFileExt == EmrDocumentFileExtention.docx.ToString())
            {
                var noticeDocs = new List<MEEmrDocumentsInfo>
                {
                    document
                };
                var guiNotice = new guiNotificationEmrDocument("Tờ bệnh án đang được soạn trên máy khác", noticeDocs)
                {
                    Module = this
                };
                guiNotice.ShowDialog();
                return;
            }

            var gui = new guiChangeEmrTemplate(emr.FK_MEEmrTypeID, document.MEEmrDocumentOrder)
            {
                Module = this
            };
            gui.InitializeControls(gui.Controls);
            gui.StartPosition = FormStartPosition.CenterParent;
            if (gui.ShowDialog() == DialogResult.OK)
            {
                BOSProgressBar.Start("Đang lưu dữ liệu");
                try
                {
                    var currentDocumentTemplate = $"{document.MEEmrDocumentOrder}. {document.MEEmrDocumentGroup}";
                    document.MEEmrDocumentOrder = gui.METemplateIndexOrder;
                    document.MEEmrDocumentGroup = gui.METemplateIndexName;
                    document.AAUpdatedUser = BOSApp.CurrentUser;
                    _emrDocumentCtrl.UpdateObject(document);

                    var documentCodeU = getDocumentNo(document);

                    HistoryEmr(GetCurrentMainObject(), cstObjectHistoryActionChange, $"Tờ {documentCodeU} chuyển gáy {gui.METemplateIndexOrder}. {gui.METemplateIndexName}.");

                    BOSProgressBar.Close();
                    Invalidate(emr.MEEmrID);
                    MessageBox.Show($"Gáy bệnh án đã chuyển thành công", "Chuyển gáy bệnh án thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    if (_notificationTab)
                    {
                        _msgLogs.Text += "\r\n CHANGE EMR TEMPLATE FAILED: " + ex.ToString();
                    }
                    else if (_logConfig)
                    {
                        _msgLogsTemp += "\r\n CHANGE EMR TEMPLATE FAILED: " + ex.ToString();
                    }
                    _msgNotification.Text = "Chuyển gáy bệnh án không thành công. Xem chi tiết ở thông báo.";
                }
                finally
                {
                    BOSProgressBar.Close();
                }

            }
        }

        internal void UpdateEmrTemplate()
        {
            var emr = _entity.MainObject as MEEmrsInfo;
            if (IsEditEmr(emr)) return;

            var documents = _emrDocumentCtrl.GetListBusinessObjects<MEEmrDocumentsInfo>(_emrDocumentCtrl.GetAllDataByForeignColumn("FK_MEEmrID", emr.MEEmrID));
            if (documents != null && documents.Count > 0)
            {
                var gui = new guiUpdateEmrTemplate(emr.FK_MEEmrTypeID, documents)
                {
                    Module = this
                };
                gui.InitializeControls(gui.Controls);
                gui.StartPosition = FormStartPosition.CenterParent;
                if (gui.ShowDialog() == DialogResult.OK)
                {
                    foreach (var templateIndex in gui.MappingList)
                    {
                        if (templateIndex.METemplateIndexIDNew > 0)
                        {
                            var objTemplateIndex = _templateIndexsCtrl.GetObjectByID(templateIndex.METemplateIndexIDNew) as METemplateIndexsInfo;
                            if (templateIndex.METemplateIndexName != objTemplateIndex.METemplateIndexName || templateIndex.METemplateIndexOrder != objTemplateIndex.METemplateIndexOrder)
                            {
                                var result = _emrDocumentCtrl.UpdateByEmrTemplateIndex(emr.MEEmrID, templateIndex.METemplateIndexName,
                                    templateIndex.METemplateIndexOrder, objTemplateIndex.METemplateIndexName, objTemplateIndex.METemplateIndexOrder);
                                if (!result)
                                {
                                    MessageBox.Show($"Cập nhật tới gáy {templateIndex.METemplateIndexOrder}. {templateIndex.METemplateIndexName} không thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return;
                                }
                            }
                        }
                    }

                    MessageBox.Show("Bệnh án đã cập nhật gáy thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Invalidate(emr.MEEmrID);
                }
            }
            else
            {
                MessageBox.Show("Bệnh án chưa có thông tin tờ bệnh án.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        private bool IsEditEmr(MEEmrsInfo emr)
        {
            emr = _emrCtrl.GetObjectByID(emr.MEEmrID) as MEEmrsInfo;
            if (emr == null)
            {
                MessageBox.Show("Bệnh án không còn tồn tại. Có thể đã bị xóa hoặc trộn vào bệnh án khác.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return true;
            }
            if (emr.MEEmrStatus != EmrStatus.InProgress.ToString())
            {
                MessageBox.Show($"Chức năng chỉ thực hiện với bệnh án đang mở.", "Không thể thực hiện chức năng này", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return true;
            }
            if (!IsAllowEdit(emr)) return true;

            var documents = _emrDocumentCtrl.GetListBusinessObjects<MEEmrDocumentsInfo>(_emrDocumentCtrl.GetAllDataByForeignColumn("FK_MEEmrID", emr.MEEmrID));
            if (documents != null && documents.Count > 0)
            {
                foreach (var document in documents)
                {
                    if (document.FK_EditingUserID > 0 && document.MEEmrDocumentHoldMachineMac != _macAddress && document.MEEmrDocumentFileExt == EmrDocumentFileExtention.docx.ToString())
                    {
                        ShowDocumentEditingUser(document);
                        return true;
                    }
                }
            }
            return false;
        }
        #endregion

        #region FILES: ClearEmrFiles, Logs
        private void ClearEmrFiles()
        {
            var enable = ConfigurationManager.AppSettings["clear-file-local"] == "true";
            if (enable)
            {
                try
                {
                    int numday = Convert.ToInt32(BOSApp.GetSystemConfigValue(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.EMR_FILE_LIFE_DAY));
                    var dateCurrent = DateTime.Now.Date;
                    var focusDirectories = new List<string> {
                    Path.Combine(_documentPath, _emrArchiveHelper.StorageDir),
                    Path.Combine(_documentPath, "Emr")
                };
                    var exceptFiles = new List<string> { ".bk" };
                    var exceptFolders = new List<string> { "Partials", "Signed" };

                    DirectoryInfo templateDir = new DirectoryInfo(_documentPath);
                    var noteFile = templateDir.GetFiles("clear_*", SearchOption.TopDirectoryOnly).OrderByDescending(p => p.CreationTime).FirstOrDefault();
                    var isDelete = true;
                    if (noteFile != null)
                    {
                        var dateCheck = File.GetCreationTime(noteFile.FullName).Date.AddDays(numday * -1);
                        if (dateCheck > dateCurrent)
                        {
                            isDelete = false;
                        }
                    }
                    if (isDelete)
                    {
                        foreach (var focusDirectory in focusDirectories)
                        {
                            Directory.CreateDirectory(focusDirectory);
                            foreach (DirectoryInfo dir in new DirectoryInfo(focusDirectory).GetDirectories())
                            {
                                if (!exceptFolders.Contains(dir.Name))
                                {
                                    DeletePhysical(exceptFiles, numday, dateCurrent, dir, false);
                                }
                                else
                                {
                                    foreach (DirectoryInfo dirSub in dir.GetDirectories())
                                    {
                                        DeletePhysical(exceptFiles, numday, dateCurrent, dir, false);
                                    }
                                }
                            }
                        }

                        // Delete note file
                        foreach (FileInfo notefile in templateDir.GetFiles("clear_*", SearchOption.TopDirectoryOnly))
                        {
                            notefile.Delete();
                        }
                        string nameNoteFileNew = string.Format(@"{0}\clear_{1}.txt", _documentPath, DateTime.Now.ToString("ddMMyyyy"));
                        if (!File.Exists(nameNoteFileNew))
                        {
                            File.Create(nameNoteFileNew).Dispose();
                        }
                    }
                }
                catch (Exception ex)
                {
                    PrintMgsLog("Có lỗi khi dọn dẹp tờ bệnh án.", ex.ToString());
                    _sysHelper.LogTxt("error", ex.ToString());
                }
            }
        }

        private static void DeletePhysical(List<string> exceptFiles, int numday, DateTime dateCurrent, DirectoryInfo dir, bool deleteDir)
        {
            foreach (FileInfo file in dir.GetFiles())
            {
                if (!exceptFiles.Any(file.Name.Contains) && file.CreationTime < dateCurrent.AddDays(numday))
                {
                    file.Delete();
                }
            }
            if (deleteDir)
            {
                if (!File.Exists(dir.FullName))
                {
                    dir.Delete(true);
                }
            }
        }

        private void LogFile(bool upload)
        {
            if (_logConfig)
            {
                if (_notificationTab && !string.IsNullOrEmpty(_msgLogs.Text))
                {
                    _sysHelper.LogTxt(string.Empty, _msgLogs.Text, _fileLog, upload);
                }
                else if (!string.IsNullOrEmpty(_msgLogsTemp))
                {
                    _sysHelper.LogTxt(string.Empty, _msgLogsTemp, _fileLog, upload);
                    _msgLogsTemp = string.Empty;
                }
            }
        }

        private bool VerifyOfficeCryto(string fileCheck, bool fixMode, string backupFullPath, string uploadPath)
        {
            var result = false;
            try
            {
                PrintMgsLog("VERIFY-ENCRYPT-FILE", fileCheck);
                var fileName = Path.GetFileName(fileCheck);
                var fileBackUpName = !string.IsNullOrEmpty(backupFullPath) ? Path.GetFileName(backupFullPath) : string.Empty;

                OfficeOpenXmlCrypto.OfficeCryptoStream verifyStream;
                //5.71% CPU TODO optimization
                var verify = OfficeOpenXmlCrypto.OfficeCryptoStream.TryOpen(fileCheck, this._emrDocumentHelper.ShareEmrPassword, out verifyStream);
                if (verifyStream != null) verifyStream.Close();

                //try again 1 time. Current set false cause task // upload.
                if (!verify && fixMode && !string.IsNullOrEmpty(backupFullPath))
                {
                    File.Delete(fileCheck);
                    using (OfficeOpenXmlCrypto.OfficeCryptoStream stream = OfficeOpenXmlCrypto.OfficeCryptoStream.Open(backupFullPath, this._emrDocumentHelper.ShareEmrPassword))
                    {
                        stream.Password = this._emrDocumentHelper.ShareEmrPassword;
                        stream.SaveAs(fileCheck); // ?? todo
                    }
                    verify = OfficeOpenXmlCrypto.OfficeCryptoStream.TryOpen(fileCheck, this._emrDocumentHelper.ShareEmrPassword, out verifyStream);
                    if (verifyStream != null) verifyStream.Close();

                    Trace.TraceError("OfficeOpenXmlCrypto ERROR + TRY AGAIN: {0}:{1}:{2}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), BOSApp.CurrentUser, fileName);
                    Trace.Flush();
                }

                if (!verify)
                {
                    //_ftpFileMng.UploadFile(uploadPath, fileBackUpName, backupFullPath); File not change on safe mode
                    throw new Exception($"Không thể mã hóa tờ bệnh án {fileName}");
                }
                else
                {
                    if (!string.IsNullOrEmpty(backupFullPath))
                    {
                        File.Delete(backupFullPath);
                    }
                }
                PrintMgsLog("KET-THUC-VERIFY-ENCRYPT-FILE", fileCheck);
                result = true;
            }
            catch (Exception ex)
            {
                PrintMgsLog("VERIFY-ENCRYPT-FILE", ex.Message);
            }

            return result;
        }

        private bool VerifyUploadFile(string source, string fileName, string emrId, string hashLocal)
        {
            if (_apiEmr != null)
            {
                var actionUri = BOSApp.GetSystemConfigValue(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.SYSTEM_CONFIGS_VERIFY_HASH);
                if (!string.IsNullOrEmpty(actionUri))
                {
                    var body = new
                    {
                        hashLocal,
                        emrId,
                        destinateFileName = fileName
                    };
                    var response = _apiEmr.Post<Emr.Base.Models.Abp.AjaxResponse, JValue>(actionUri, null, body);
                    if (response != null)
                    {
                        if (!response.Success)
                        {
                            if (response.Error != null)
                            {
                                MessageBox.Show(response.Error.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                PrintMgsLog("VERIFY-UPLOAD-FILE-ERROR", response.Error.Message);
                                return false;
                            }
                        }
                    }
                    else
                    {
                        PrintMgsLog("VERIFY-UPLOAD-FILE", "Không có dữ liệu trả về từ api.");
                    }
                }
            }

            return true;
        }
        #endregion

        #region Nhom va sap xep to benh an
        internal void GroupEmrDocuments()
        {
            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            var meEmr = _entity.MainObject as MEEmrsInfo;
            if (document.MEEmrDocumentID == 0)
            {
                MessageBox.Show("Bạn phải chọn vào tờ bệnh án để thực hiện chức năng này", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Get document base template
            var data = _emrDocumentCtrl.GetAllByEmrAndGroup(document.FK_MEEmrID, document.MEEmrDocumentGroup);
            if (data != null)
            {
                foreach (var doc in data)
                {
                    if (doc.FK_EditingUserID > 0 && doc.FK_EditingUserID != BOSApp.CurrentEmployeesInfo.HREmployeeID)
                    {
                        ShowDocumentEditingUser(doc);
                        return;
                    }
                }

                var gui = new guiOrderAndGroupDocument(data, meEmr)
                {
                    Module = this
                };
                if (gui.ShowDialog() == DialogResult.OK)
                {
                    if (gui.isChange)
                    {
                        var listDocs = gui._data;
                        foreach (var item in listDocs)
                        {
                            _emrDocumentCtrl.UpdateObject(item);

                            var documentCodeU = getDocumentNo(item);
                            HistoryEmr(GetCurrentMainObject(), "Change", $"gom nhóm tờ bệnh án {documentCodeU}.");
                        }
                        InvalidateEmrDocumentList(document.FK_MEEmrID);
                    }
                }
            }
            else
            {
                MessageBox.Show("Không tìm thấy tờ bệnh án để thực hiện chức năng này", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }
        #endregion

        public override void TriggerWorkflow(object sender, string eventName, object toolbar)
        {
            // Check data: US 985
            var action = toolbar as STToolbarsInfo;
            var parentBar = _stToolbarsController.GetObjectByID(action.STToolbarParentID) as STToolbarsInfo;
            if (parentBar == null)
            {
                MessageBox.Show("Cấu hình thanh công cụ sai. Vui lòng kiểm tra cấu hình", "Cấu hình sai", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            var actionStatus = "ToWaitClose";
            if (action.STToolbarTag == actionStatus)
            {
                if (!ValidateBeforeActionEmr(actionStatus))
                {
                    return;
                }
            }

            base.TriggerWorkflow(sender, eventName, toolbar);

            // history number
            var emr = GetCurrentMainObject();
            var gEObjectHistory = _geObjHistoryCtrl.GetLatestHistoryByObjectNameAndObjectId(TableName.MEEmrsTableName, emr.MEEmrID, action.STToolbarTag);
            if (gEObjectHistory != null && gEObjectHistory.GEObjectHistoryID > 0)
            {
                gEObjectHistory.GEObjectHistoryObjectNumber = emr.MEEmrNo;
                _geObjHistoryCtrl.UpdateObject(gEObjectHistory);
            }
        }

        internal void ShareEmrClose()
        {
            var meEmr = _entity.MainObject as MEEmrsInfo;
            var sharesDb = _shareCtrl.GetByEmrId(meEmr.MEEmrID);
            _entity.MEEmrShareList.EndCurrentEdit();
            foreach (var share in _entity.MEEmrShareList)
            {
                if (share.MEEmrShareHistoryActive)
                {
                    if ((share.MEEmrShareHistoryToDate - share.MEEmrShareHistoryFromDate).TotalMinutes < 15)
                    {
                        MessageBox.Show("Thời gian chia sẻ không hợp lệ. Tối thiểu 15 phút", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    };
                    if (share.FK_HRDepartmentID == 0 || share.FK_HREmployeeID == 0)
                    {
                        MessageBox.Show("Chỉ chia sẻ cho nhân viên cụ thể.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }

                if (share.MEEmrShareHistoryID > 0)
                {
                    _shareCtrl.UpdateObject(share);
                }
                else
                {
                    _shareCtrl.CreateObject(share);
                }
            }

            var ids = _entity.MEEmrShareList.Select(m => m.MEEmrShareHistoryID).ToList();
            foreach (var share in sharesDb)
            {
                if (!ids.Contains(share.MEEmrShareHistoryID))
                {
                    share.AAStatus = Status.Delete.ToString();
                    share.AAUpdatedUser = BOSApp.CurrentUsersInfo.ADUserName;
                    _shareCtrl.UpdateObject(share);
                }
            }

            MessageBox.Show("Lưu chia sẻ thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            base.ActionComplete();
        }

        private bool InactiveShareEmr(int emrID)
        {
            return _shareCtrl.UpdateAllInactive(emrID);
        }

        #region Tac Vu Chay Ngam

        public void GetEmrDocumentsBackground()
        {
            var meEmr = _entity.MainObject as MEEmrsInfo;
            var mEEmrNo = meEmr.MEEmrNo;
            if (_apiEmr != null)
            {
                var actionUri = BOSApp.GetSystemConfigValue(SysCfgConsts.EMR_API_ENDPOINT, SysCfgConsts.EMR_DOCUMENTS_BACKGROUND_ALL_BY_EMR);
                var paramList = new Dictionary<string, object>
                {
                    { "emrNo", mEEmrNo }
                };
                if (!string.IsNullOrEmpty(actionUri))
                {
                    try
                    {
                        var listMdAutoGenDocumentDto = new List<MdAutoGenDocumentDto>();
                        PrintMgsLog("BAT-DAU-GOI-API", actionUri);
                        var response = _apiEmr.Post<Emr.Base.Models.Abp.AjaxResponse, JArray>(actionUri, null, paramList);
                        PrintMgsLog("KET-THUC-GOI-API", actionUri);
                        if (response != null)
                        {
                            if (response.Success)
                            {
                                if (response.Result != null && response.Result.Count > 0)
                                {
                                    listMdAutoGenDocumentDto = response.Result.ToObject<List<MdAutoGenDocumentDto>>();
                                    // group base VENDOC_NO
                                    var groupVendocs = listMdAutoGenDocumentDto.GroupBy(m => m.VENDOR_DOC_NO).Select((n) => new { Key = n.Key, Items = n.ToList() });
                                    var countMiss = 0;
                                    foreach (var groupVendoc in groupVendocs)
                                    {
                                        var maxE = groupVendoc.Items.OrderByDescending(m => m.ID).FirstOrDefault();
                                        if (maxE == null) continue;

                                        if (maxE.STATE != AutoGenDocumentStatus.CREATED.ToString())
                                        {
                                            countMiss++;
                                        }
                                    }

                                    if (countMiss > 0)
                                    {
                                        ShowFlashNotification($"Bệnh án {mEEmrNo} có {countMiss} tác vụ ngầm chưa thành công. Xem chi tiết ở [Tác vụ chạy ngầm]", 5000, 1000);
                                    }
                                }
                            }
                        }

                        List<ADConfigValuesInfo> adConfigs = _objConfigValuesController.GetConfigValuesByGroup("MdAutoGenDocumentsStatus");
                        var finalList = new List<MdAutoGenDocumentDto>();
                        foreach (var item in listMdAutoGenDocumentDto)
                        {
                            var configE = adConfigs.Where(m => m.ADConfigKeyValue.Equals(item.STATE)).FirstOrDefault();
                            item.STATE = configE.ADConfigText;
                            finalList.Add(item);
                        }

                        var gridControl = this.Controls["fld_dgcMdAutoGenDocumentDto"] as MdAutoGenDocumentDtoGridControl;
                        if (gridControl != null)
                            gridControl.LoadDataToGridMdAutoGenDocumentDto(finalList);
                    }
                    catch (Exception ex)
                    {
                        ShowFlashNotification($"[Tác vụ chạy ngầm] Có lỗi khi gọi EMR API " + ex.Message, 4000);
                    }
                }
            }
        }
        internal void RunAgainEmrDocumentsBackground(List<MdAutoGenDocumentDto> selects)
        {
            List<ADConfigValuesInfo> adConfigs = _objConfigValuesController.GetConfigValuesByGroup("MdAutoGenDocumentsStatus");
            var conditionStatus = new List<string>()
            {
                AutoGenDocumentStatus.DISCARDED.ToString()
            };
            var conditionAdConfigs = adConfigs.Where(m => conditionStatus.Contains(m.ADConfigKeyValue)).Select(m => m.ADConfigText).ToList(); //.FirstOrDefault().ADConfigText;

            foreach (var item in selects)
            {
                if (!conditionAdConfigs.Contains(item.STATE))
                {
                    MessageBox.Show($"Chức năng chỉ áp dụng với tờ bệnh án có trạng thái = [{String.Join("] hoặc [", conditionAdConfigs.ToArray())}]. \nVui lòng chọn lại", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }
            }

            var groups = selects.GroupBy(m => m.VENDOR_DOC_NO).Select((n) => new { Key = n.Key, Items = n.ToList() });
            var notices = new List<string>();
            var newestSelects = new List<MdAutoGenDocumentDto>();
            foreach (var group in groups)
            {
                if (group.Items.Count > 1)
                {
                    var sortItems = group.Items.OrderByDescending(m => m.ID);
                    var newestItem = sortItems.FirstOrDefault();
                    var othersItem = group.Items.Where(m => !m.ID.Equals(newestItem.ID)).Select(m => m.ID).ToList();
                    notices.Add($"Mã tờ HIS [{group.Key}]: ID mới nhất [{newestItem.ID}] - cũ [{string.Join(";", othersItem)}]");
                    newestSelects.Add(newestItem);
                }
                else
                {
                    var item = group.Items.FirstOrDefault();
                    if (item != null)
                    {
                        newestSelects.Add(item);
                    }
                }
            }

            if (notices.Count() > 0)
            {
                MessageBox.Show(string.Join("\n", notices) + "\n\n Vui lòng chỉ chọn các dòng mới nhất",
                    "[Cảnh báo] Các dòng cũ hệ thống sẽ tự động bỏ qua [không chạy lại]", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            #region Apply newest id
            if (_apiEmr != null)
            {
                var actionUri = BOSApp.GetSystemConfigValue(SysCfgConsts.EMR_API_ENDPOINT, SysCfgConsts.EMR_DOCUMENTS_BACKGROUND_ALL_BY_EMR);
                if (!string.IsNullOrEmpty(actionUri))
                {
                    var emrNo = newestSelects.FirstOrDefault()?.EMR_NO;
                    var paramList = new Dictionary<string, object>{
                        { "emrNo", emrNo }
                    };
                    var listMdAutoGenDocumentDto = new List<MdAutoGenDocumentDto>();
                    var response = _apiEmr.Post<Emr.Base.Models.Abp.AjaxResponse, JArray>(actionUri, null, paramList);
                    if (response != null)
                    {
                        if (response.Success)
                        {
                            if (response.Result != null && response.Result.Count > 0)
                            {
                                listMdAutoGenDocumentDto = response.Result.ToObject<List<MdAutoGenDocumentDto>>();
                            }
                        }
                    }
                    foreach (var document in newestSelects)
                    {
                        var maxE = listMdAutoGenDocumentDto.Where(d => d.VENDOR_DOC_NO == document.VENDOR_DOC_NO).OrderByDescending(m => m.ID).FirstOrDefault();
                        if (maxE == null) continue;
                        if (maxE.ID != document.ID)
                        {
                            if (maxE.STATE == AutoGenDocumentStatus.DISCARDED.ToString())
                            {
                                MessageBox.Show($"{document.VENDOR_DOC_NO} đang chọn ID cũ [{document.ID}], có ID mới hơn [{maxE.ID}]. " +
                                $"\nVui lòng tải lại dữ liệu và chọn ID mới nhất", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                return;
                            }
                            else if (maxE.STATE == AutoGenDocumentStatus.CREATED.ToString())
                            {
                                MessageBox.Show($"{document.VENDOR_DOC_NO} [Đã tạo] thành công với ID [{maxE.ID}].", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                return;
                            }
                            else if (maxE.STATE == AutoGenDocumentStatus.SCHEDULED.ToString() || maxE.STATE == AutoGenDocumentStatus.RETRYING.ToString())
                            {
                                MessageBox.Show($"{document.VENDOR_DOC_NO} đã được lên lịch [Chờ thực hiện] / [Đang thử lại] với ID [{maxE.ID}].", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                return;
                            }
                            else
                            {
                                MessageBox.Show($"{document.VENDOR_DOC_NO} [Đang tạo] với ID [{maxE.ID}].", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                                return;
                            }
                        }
                    }
                }
            }
            #endregion
            if (newestSelects.Count <= 0) return;
            var ids = newestSelects.Select(m => m.ID).Distinct().ToList();
            var vendorDocs = newestSelects.Select(m => m.VENDOR_DOC_NO).Distinct().ToList();
            if (_apiEmr != null)
            {
                var actionUri = BOSApp.GetSystemConfigValue(SysCfgConsts.EMR_API_ENDPOINT, SysCfgConsts.MD_AUTO_GEN_DOCUMENTS_UPDATE);
                if (!string.IsNullOrEmpty(actionUri))
                {
                    var paramList = new Dictionary<string, object>
                    {
                        { "ids",  ids },
                        { "userName", BOSApp.CurrentUser },
                    };
                    var response = _apiEmr.Post<Emr.Base.Models.Abp.AjaxResponse, JValue>(actionUri, null, paramList);
                    if (response != null)
                    {
                        if (response.Success)
                        {
                            MessageBox.Show($"Đã kích hoạt thành công.\n{string.Join("\n", vendorDocs)}" +
                                $"\nTác vụ ký ngầm của các tờ này cũng đã được kích hoạt." +
                                "\n\nXin chờ trong ít phút để hệ thống tạo lại tờ.\n[Tải lại] để xem tình trạng.", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            var notificationStr = _notificationTab ? " Xem chi tiết ở thông báo" : string.Empty;
                            MessageBox.Show($"{response.Error.Message}{notificationStr}", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            if (_notificationTab)
                            {
                                _msgLogs.Text += "\r\n TÁC VỤ CHẠY NGẦM - THỰC HIỆN LẠI: " + response.Error.Details;
                            }
                            else if (_logConfig)
                            {
                                _msgLogsTemp += "\r\n TÁC VỤ CHẠY NGẦM - THỰC HIỆN LẠI: " + response.Error.Details;
                            }
                        }
                    }
                }
            }
        }
        private void CheckBgJobStatusDocumentList()
        {
            Task.Factory.StartNew(() =>
            {
                foreach (var item in _entity.MEEmrDocumentsList)
                {
                    if (item.MEEmrDocumentStatus != EmrDocumentStatus.InProgress.ToString()) continue;
                    if (item.MEEmrDocumentBgJobStatus == EmrDocumentBgJobStatus.CreatedWithErr.ToString())
                    {
                        var msg = "Có tờ bệnh án lỗi khi tạo ngầm. Xem ở cột [Tác vụ ngầm] trên [Cây bệnh án]. Chi tiết vui lòng liên hệ Admin";
                        if (string.IsNullOrEmpty(_msgNotification.Text))
                            ShowFlashNotification(msg, 10000, 0);
                        else
                            ShowFlashNotification(msg, 10000, 5000);
                        break;
                    }
                }
            });
        }
        #endregion

        public void ReCloseEmrDocuments()
        {
            if (BOSApp.CurrentUserGroupInfo.ADUserGroupRole != UserGroupRole.admin.ToString())
                return;

            var emr = this._entity.MainObject as MEEmrsInfo;
            if (emr.MEEmrStatus != EmrStatus.Closed.ToString())
            {
                MessageBox.Show("Chức năng chỉ thực hiện được trên bệnh án đã đóng.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            };

            if (BOSApp.CurrentUserGroupInfo.ADUserGroupRole != UserGroupRole.admin.ToString())
                return;

            this._dpnViewPdf.Visible = false;
            this._dpnRichEdit.Visible = true;
            this.DockManager.ActivePanel = this._dpnRichEdit;

            var template = _templateCtrl.GetObjectByNo("TDT") as METemplatesInfo;
            if (template == null) return;
            InitDocumentSession(template.METemplateID);
            _entity.METemplate = template;
            var listDocs = _emrDocumentCtrl.GetListBusinessObjects<MEEmrDocumentsInfo>(_emrDocumentCtrl.GetByMEEmrID(emr.MEEmrID));
            BOSProgressBar.Start("Đang xử lý");
            emr.AllowPropertyChangedEvent = false;
            try
            {
                foreach (var item in listDocs)
                {
                    if (item.FK_METemplateID == template.METemplateID)
                    {
                        if (item.MEEmrDocumentFileExt == EmrDocumentFileExtention.pdf.ToString())
                        {
                            // bo qua cac to an va huy
                            if (item.MEEmrDocumentStatus == EmrDocumentStatus.Discarded.ToString()
                                || item.MEEmrDocumentStatus == EmrDocumentStatus.Hidden.ToString())
                            {
                                continue;
                            }
                            _richEditCtrl.Visible = true;
                            _richEditCtrl.CreateNewDocument(false);
                            _richEditCtrl.ReadOnly = false;

                            string fileDocx = string.Format(@"{0}\Emr\{1}\{2}.docx", _documentPath, item.FK_MEEmrID, item.MEEmrDocumentFile);
                            _ftpFileMng.DownloadFile($"/Emr/{item.FK_MEEmrID}/", item.MEEmrDocumentFile + ".docx", fileDocx);
                            using (OfficeOpenXmlCrypto.OfficeCryptoStream stream = OfficeOpenXmlCrypto.OfficeCryptoStream.Open(fileDocx, this._emrDocumentHelper.ShareEmrPassword))
                            {
                                _richEditCtrl.LoadDocument(stream, DocumentFormat.OpenXml);
                            }

                            this.RemoveAllForPrint(this._richEditCtrl, item.FK_METemplateID);
                            string fileName = string.Format(@"{0}\Emr\{1}\{2}.{3}", _documentPath, item.FK_MEEmrID, item.MEEmrDocumentFile, EmrDocumentFileExtention.pdf.ToString());
                            this._richEditCtrl.ExportToPdf(fileName);
                            _ftpFileMng.UploadFile($"/Emr/{item.FK_MEEmrID}/", item.MEEmrDocumentFile + "." + EmrDocumentFileExtention.pdf.ToString(), fileName);
                            item.AAUpdatedUser = BOSApp.CurrentUser;
                            item.MEEmrDocumentDesc += $" [{DateTime.Now.ToString("dd/MM/yyyy HH:mm")}] Chuyển PDF sau khi đóng.";
                            _emrDocumentCtrl.UpdateObject(item);

                            var documentCodeU = getDocumentNo(item);
                            HistoryEmr(GetCurrentMainObject(), "Change", $"{documentCodeU} khôi phục đóng tờ bệnh án.");
                            //_msgLogs.Text += "\r\n EXPORT OK: " + item.MEEmrDocumentFile;
                        }
                    }
                }
                _entity.MEEmrDocumentsList.Invalidate(emr.MEEmrID);
                BOSProgressBar.Close();
                MessageBox.Show("Đã xong. Vui lòng thực hiện [Lưu trữ] lại.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                if (_notificationTab)
                {
                    _msgLogs.Text += "\r\n CLOSE FAILED: " + ex.ToString();
                }
                else if (_logConfig)
                {
                    _msgLogsTemp += "\r\n CLOSE FAILED: " + ex.ToString();
                }
                _msgNotification.Text = "Xử lý dữ liệu không thành công. Xem chi tiết ở thông báo.";
            }
            finally
            {
                emr.AllowPropertyChangedEvent = true;
                BOSProgressBar.Close();
            }
        }

        #region Checkup Local by Threading
        public void ExportToPdf()
        {
            var pdfExt = EmrDocumentFileExtention.pdf.ToString();
            var emr = _entity.MainObject as MEEmrsInfo;
            if (_sysHelper.AllowThread())
            {
                MessageBox.Show("Chức năng sẽ thực hiện kết xuất PDF ngầm.\n\n[Cứ tiếp tục công việc của bạn].\nHệ thống sẽ hiển thị thông báo sau khi hoàn tất.",
                "KẾT XUẤT PDF", MessageBoxButtons.OK, MessageBoxIcon.Error);
                try
                {
                    var thread = new System.Threading.Thread(() =>
                    {
                        var documents = _emrDocumentCtrl.GetListBusinessObjects<MEEmrDocumentsInfo>(_emrDocumentCtrl.GetByMEEmrID(emr.MEEmrID));
                        RichEditControl richContrl = new RichEditControl();
                        EmrDocumentHelper emrDocumentHelper = new EmrDocumentHelper(richContrl);
                        var docx = "." + EmrDocumentFileExtention.docx.ToString();
                        foreach (var d in documents)
                        {
                            try
                            {
                                var document = _emrDocumentCtrl.GetObjectByID(d.MEEmrDocumentID) as MEEmrDocumentsInfo;
                                if (document.FK_EditingUserID > 0 && document.MEEmrDocumentHoldMachineMac != _macAddress)
                                {
                                    //to dang bi giu boi nguoi khac thi khong thuc hien
                                    continue;
                                }
                                if (document.MEEmrDocumentStatus == EmrDocumentStatus.InProgress.ToString())
                                {
                                    if (document.MEEmrDocumentFileExt == EmrDocumentFileExtention.docx.ToString())
                                    {
                                        var emrhelper = new EmrHelper(_documentPath, _pdfProcessor, _ftpFileMng, _hashProvider, _digitalSig);
                                        emrhelper.ExportPdf(document, true, true, true);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                PrintMgsLog("KẾT XUẤT PDF", $"[CÓ LỖI] khi xuất PDF bệnh án: {emr.MEEmrNo} - {d.MEEmrDocumentFile}\n" + ex.ToString());
                            }
                        }
                        var patient = _patientCtrl.GetObjectByID(emr.FK_MEPatientID) as MEPatientsInfo;
                        MessageBox.Show($"Xuất PDF thành công bệnh án [{emr.MEEmrNo} - {patient.MEPatientName}]. \n\nTiếp tục [kiểm duyệt] để hoàn tất quy trình",
                            "KẾT XUẤT PDF THÀNH CÔNG", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    });
                    thread.Start();
                    _backgroundJobThreads.Add(thread);
                }
                catch (Exception ex)
                {
                    _sysHelper.LogTxt("error", ex.ToString());
                }
            }
            else
            {
                var documents = _emrDocumentCtrl.GetListBusinessObjects<MEEmrDocumentsInfo>(_emrDocumentCtrl.GetByMEEmrID(emr.MEEmrID));
                RichEditControl richContrl = new RichEditControl();
                EmrDocumentHelper emrDocumentHelper = new EmrDocumentHelper(richContrl);
                var docx = "." + EmrDocumentFileExtention.docx.ToString();
                foreach (var d in documents)
                {
                    try
                    {
                        var document = _emrDocumentCtrl.GetObjectByID(d.MEEmrDocumentID) as MEEmrDocumentsInfo;
                        if (document.FK_EditingUserID > 0 && document.MEEmrDocumentHoldMachineMac != _macAddress)
                        {
                            //to dang bi giu boi nguoi khac thi khong thuc hien
                            continue;
                        }
                        if (document.MEEmrDocumentStatus == EmrDocumentStatus.InProgress.ToString())
                        {
                            if (document.MEEmrDocumentFileExt == EmrDocumentFileExtention.docx.ToString())
                            {
                                var emrhelper = new EmrHelper(_documentPath, _pdfProcessor, _ftpFileMng, _hashProvider, _digitalSig);
                                emrhelper.ExportPdf(document, true, true, true);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        PrintMgsLog("KẾT XUẤT PDF", $"[CÓ LỖI] khi xuất PDF bệnh án: {emr.MEEmrNo} - {d.MEEmrDocumentFile}\n" + ex.ToString());
                    }
                }
                var patient = _patientCtrl.GetObjectByID(emr.FK_MEPatientID) as MEPatientsInfo;
                MessageBox.Show($"Xuất PDF thành công bệnh án [{emr.MEEmrNo} - {patient.MEPatientName}]. \n\nTiếp tục [kiểm duyệt] để hoàn tất quy trình",
                    "KẾT XUẤT PDF THÀNH CÔNG", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        /// <summary>
        /// threading
        /// </summary>
        /// <param name="richContrl"></param>
        /// <param name="template"></param>
        /// <param name="emrDocumentHelper"></param>
        public void RemoveAllForExportPdf(RichEditControl richContrl, METemplatesInfo template, EmrDocumentHelper emrDocumentHelper)
        {
            this.RemoveAllTagForExportPdf(richContrl, template);
            this.RemoveAllCommentForPrint(richContrl);

            var templateParams = _templateParamCtrl.GetAllTemplateParamObjectByTemplateID(template.METemplateID);
            emrDocumentHelper.RemoveAllHiddenDataForPrint(richContrl, template.METemplateID, templateParams);
            emrDocumentHelper.ProgressHyperlinksParamForPrint(richContrl.Document);

            if (template.METemplateSignatureNotAlone)
                emrDocumentHelper.SignatureAndContentOnTheSamePage(richContrl, richContrl.Document, richContrl.DocumentLayout);

            if (template.METemplateNo.ToUpper() == "TDT")
            {
                var table = richContrl.Document.Tables.OrderByDescending(t => t.Range.Length).FirstOrDefault();
                if (table != null && table.Rows.Count > 2)
                {
                    try
                    {
                        int pageCount = richContrl.DocumentLayout.GetFormattedPageCount();
                        for (int i = 0; i < pageCount; i++)
                        {
                            var collector = new TableCellLayoutVisitor(richContrl);
                            collector.Visit(richContrl.DocumentLayout.GetPage(i));
                            foreach (var range in collector.Ranges)
                            {
                                var cell = richContrl.Document.Tables.GetTableCell(richContrl.Document.CreatePosition(range.Start));
                                if (cell != null && table == cell.Table)
                                {
                                    var pos = richContrl.Document.CreatePosition(cell.ContentRange.End.ToInt() - 1);
                                    var p = richContrl.Document.Paragraphs.Insert(pos);
                                    p = richContrl.Document.Paragraphs.Insert(pos);
                                    p = richContrl.Document.Paragraphs.Insert(pos);
                                    p = richContrl.Document.Paragraphs.Insert(pos);
                                    p = richContrl.Document.Paragraphs.Insert(pos);
                                }
                            }
                        }
                    }
                    catch (Exception) {/*do nothing to make sure not error*/}
                }
            }
        }
        /// <summary>
        /// threading
        /// </summary>
        /// <param name="richContrl"></param>
        /// <param name="template"></param>
        private void RemoveAllTagForExportPdf(RichEditControl richContrl, METemplatesInfo template)
        {
            var doc = richContrl.Document;
            doc.BeginUpdate();
            for (int i = 0; i < doc.Hyperlinks.Count; i++)
            {
                var link = doc.Hyperlinks[i];
                try
                {
                    if (RemoveBreakingHyperlink(doc, link))
                        continue;
                    /*if (template.METemplateNo.ToUpper() == "TDT")
                    {
                        if (link.NavigateUri.StartsWith("thaythethe_chandoan|gid="))
                        {
                            //14 = length('Thêm chẩn đoán')
                            doc.Replace(doc.CreateRange(doc.CreatePosition(link.Range.End.ToInt() - 14), 14), " ");
                            continue;
                        }
                    }*/
                }
                catch (Exception) {/*do nothing*/}
                link.ToolTip = null;
                doc.Replace(link.Range, string.Empty);
                doc.Hyperlinks.Remove(link);
                i--;
            }
            doc.ReplaceAll(EmrParam.BeginTag, " ", SearchOptions.None);
            doc.ReplaceAll(EmrParam.EndTag, " ", SearchOptions.None);
            doc.EndUpdate();
        }
        #endregion

        private bool ValidateBeforeActionEmr(string action)
        {
            var config = BOSApp.GetSystemConfigValue(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.EMR_DOCUMENT_STATUS_VERIFY).ToUpper();
            var actionConfig = config.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries).Select(p => p.Trim()).ToList();
            if (!string.IsNullOrEmpty(config) && actionConfig.Contains(action.ToUpper()))
            {
                var emr = _entity.MainObject as MEEmrsInfo;
                var documentsMissData = new List<MEEmrDocumentsInfo>();
                var docs = _emrDocumentCtrl.GetByEmrAndStatusInProgress(emr.MEEmrID);
                foreach (var document in docs)
                {
                    if (document.MEEmrDocumentFileExt != EmrDocumentFileExtention.docx.ToString()) continue;

                    var missingDataParam = ValidateDocumentRequired(document, false);
                    if (missingDataParam.Count > 0)
                    {
                        var msgValidates = string.Join("\n", missingDataParam.Select(x => "<" + x.Value + ">").ToArray());
                        document.MEEmrDocumentDesc = $"{msgValidates} chưa có dữ liệu.";
                        documentsMissData.Add(document);
                    }
                }

                if (documentsMissData != null && documentsMissData.Count > 0)
                {
                    var gui = new guiNotificationEmrDocument("DANH SÁCH TỜ BỆNH ÁN THIẾU DỮ LIỆU", documentsMissData)
                    {
                        Module = this
                    };
                    BOSProgressBar.Close();
                    gui.ShowDialog();
                    return false;
                }
            }
            return true;
        }

        private Dictionary<string, string> ValidateDocumentRequired(MEEmrDocumentsInfo document, bool current)
        {
            var result = new Dictionary<string, string>();
            var validates = GetTemplateParamsValidateByDocument(document, current);
            foreach (var item in validates)
            {
                var paramNo = item.METemplateParamPath.Split('.').Last();
                var paramE = AppMemCache.GetParamFromDictKeyNo(paramNo);
                if (paramE != null)
                {
                    if (!result.ContainsKey(item.METemplateParamPath))
                    {
                        result.Add(item.METemplateParamPath, paramE.MEParamName);
                    }
                }
            }
            return result;
        }

        #region Run Action EMR
        internal void RunActionsOnOpenEmr(MEEmrsInfo emr)
        {
            var actions = _emrTypeActionCtrl.GetAllByTypeIdAndWhen(emr.FK_MEEmrTypeID, EmrTypeActionWhen.Open.ToString());
            foreach (var act in actions)
            {
                var action = this._actionsController.GetObjectByID(act.FK_MEEmrActionID) as MEEmrActionsInfo;
                if (action == null) continue;

                if (action.MEEmrActionScope == EmrActionScopes.Document.ToString()) continue;

                if (_checkSystem)
                {
                    _sysHelper.LogTxt("information", $"Bắt đầu chạy chức năng {action.MEEmrActionNo}.");
                    var watchAct = Stopwatch.StartNew();
                    CallEmrAction($"{action.MEEmrActionNo}{EmrParam.TagCodeSeparator}{EmrParam.GuidTag}={emr.MEEmrID}", this._tempRichEditCtrl.Document.Range);
                    watchAct.Stop();
                    var elapsedAct = watchAct.ElapsedMilliseconds / 1000.0;
                    _sysHelper.LogTxt("information", $"{elapsedAct} giây. Hoàn tất chạy chức năng {action.MEEmrActionNo}.");
                }
                else
                {
                    CallEmrAction($"{action.MEEmrActionNo}{EmrParam.TagCodeSeparator}{EmrParam.GuidTag}={emr.MEEmrID}", this._tempRichEditCtrl.Document.Range);
                }
            }
        }

        internal void RunActionsOnSaveDocument(MEEmrDocumentsInfo doc)
        {
            var actions = this._templateActionCtrl.GetAllByTemplateID(doc.FK_METemplateID)
                .Where(o => o.MEEmrTemplateActionWhen == EmrTemplateActionWhen.Save.ToString())
                .OrderBy(o => o.MEEmrTemplateActionOrder).ToList();
            foreach (var act in actions)
            {
                var action = this._actionsController.GetObjectByID(act.FK_MEEmrActionID) as MEEmrActionsInfo;
                if (action != null)
                {
                    if (_checkSystem)
                    {
                        _sysHelper.LogTxt("information", $"Bắt đầu chạy chức năng {action.MEEmrActionNo}.");
                        var watchAct = Stopwatch.StartNew();
                        CallEmrAction($"{action.MEEmrActionNo}{EmrParam.TagCodeSeparator}{EmrParam.GuidTag}={doc.MEEmrDocumentGuid}", this._richEditCtrl.Document.Range, meEmrTemplateActionDo: act.MEEmrTemplateActionDo);
                        watchAct.Stop();
                        var elapsedAct = watchAct.ElapsedMilliseconds / 1000.0;
                        _sysHelper.LogTxt("information", $"{elapsedAct} giây. Hoàn tất chạy chức năng {action.MEEmrActionNo}.");
                    }
                    else
                    {
                        CallEmrAction($"{action.MEEmrActionNo}{EmrParam.TagCodeSeparator}{EmrParam.GuidTag}={doc.MEEmrDocumentGuid}", this._richEditCtrl.Document.Range, meEmrTemplateActionDo: act.MEEmrTemplateActionDo);
                    }
                }
            }
        }
        #endregion

        #region Test new parser
        public void UnitTestNewParser()
        {
            int count = 0;
            for (int emrId = 65512; emrId < 65602; emrId++) //65602
            {
                Invalidate(emrId);
                var emr = _entity.MainObject as MEEmrsInfo;
                if (emr.MEEmrID == 0) continue;
                PrintMgsLog("UNIT_TEST_PARSER", "EMR: " + emr.MEEmrNo);
                var documents = _emrDocumentCtrl.GetListBusinessObjects<MEEmrDocumentsInfo>(_emrDocumentCtrl.GetByMEEmrID(emrId));
                foreach (var document in documents)
                {
                    if (document.MEEmrDocumentFileExt != "docx") continue;
                    InvalidateDocument(document);
                    TestNewParser(document);
                }
                _richEditCtrl.CreateNewDocument(false);
                LogFile(false);
                count++;
            }
            MessageBox.Show("DONE: " + count);
        }
        private bool TestNewParser(MEEmrDocumentsInfo document)
        {
            try
            {
                var templateParams = AppMemCache.GetTemplateParams(document.FK_METemplateID);
                var begin = DateTime.Now;
                var oldData = JsonConvert.DeserializeObject(this._emrParser.ParserFieldsToJson(this._emrDocumentHelper.GetAllDataFieldInRangeOrDocument(), templateParams)) as JToken;
                PrintMgsLog("UNIT_TEST_PARSER", "Old: " + (DateTime.Now - begin).TotalMilliseconds);
                begin = DateTime.Now;
                var dictParams = AppMemCache.GetParamsDictKeyNo();
                var templateParamsPath = AppMemCache.GetTemplateParamsDictPath(document.FK_METemplateID);
                JToken newData = JsonConvert.DeserializeObject(this._emrParser.ParserFieldsToJsonV3(this._emrDocumentHelper.GetAllDataFieldInRangeOrDocument(), dictParams, templateParamsPath)) as JToken;
                PrintMgsLog("UNIT_TEST_PARSER", "New: " + (DateTime.Now - begin).TotalMilliseconds);

                var oldDataFlat = _dataHelper.FlattenObjDataWithoutChangeName(string.Empty, oldData as JObject);
                var newDataFlat = _dataHelper.FlattenObjDataWithoutChangeName(string.Empty, newData as JObject);

                var error = false;
                foreach (var item in oldDataFlat)
                {
                    try
                    {
                        if (!newDataFlat.ContainsKey(item.Key))
                        {
                            PrintMgsLog("UNIT_TEST_PARSER_ERR", "Missed: " + item.Key);
                            error = true;
                            continue;
                        }
                        var newProp = newDataFlat[item.Key] as JToken;
                        var oldProp = oldDataFlat[item.Key] as JToken;
                        switch (oldProp.Type)
                        {
                            case JTokenType.Object:
                                PrintMgsLog("UNIT_TEST_PARSER_ERR", "JTokenType.Object: " + item.Key);
                                error = true;
                                break;
                            case JTokenType.Array:
                                PrintMgsLog("UNIT_TEST_PARSER", "JTokenType.Array: " + item.Key);
                                if (newProp.Count() != oldProp.Count())
                                    error = true;
                                break;
                            case JTokenType.Property:
                                PrintMgsLog("UNIT_TEST_PARSER_ERR", "JTokenType.Property: " + item.Key);
                                error = true;
                                break;
                            case JTokenType.Integer:
                                error = (int)oldProp != (int)newProp;
                                break;
                            case JTokenType.Float:
                                error = (float)oldProp != (float)newProp;
                                break;
                            case JTokenType.String:
                                error = oldProp.ToString() != newProp.ToString();
                                break;
                            case JTokenType.Boolean:
                                error = (bool)oldProp != (bool)newProp;
                                break;
                            case JTokenType.Date:
                                error = ((DateTime)oldProp).ToBinary() != ((DateTime)newProp).ToBinary();
                                break;
                            case JTokenType.TimeSpan:
                                error = ((TimeSpan)oldProp).TotalMilliseconds != ((TimeSpan)newProp).TotalMilliseconds;
                                break;
                            default:
                                error = oldProp.ToString() != newProp.ToString();
                                break;
                        }
                        if (error)
                        {
                            PrintMgsLog("UNIT_TEST_PARSER_ERR", $"Diff: {item.Key} {oldProp.ToString()} {newProp.ToString()}");
                        }
                    }
                    catch (Exception ex)
                    {
                        PrintMgsLog("UNIT_TEST_PARSER_ERR", $"Exp: {item.Key} {ex.ToString()}");
                        error = true;
                    }
                }
                return error;
            }
            catch (Exception ex)
            {
                PrintMgsLog("UNIT_TEST_PARSER_ERR", $"Exp: {document.FK_MEEmrID}/{document.MEEmrDocumentFile} {ex.ToString()}");
            }
            return true;
        }
        #endregion

        #region Chuyen dong TDT len xuong
        public void SortTableRow()
        {
            if (this._richEditCtrl.Modified)
            {
                MessageBox.Show("Lưu thay đổi trước khi thực hiện chức năng này",
                    "NỘI DUNG TỜ BỆNH ÁN ĐÃ THAY ĐỔI", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var doc = _richEditCtrl.Document;
            var validSelection = true;
            if (doc.Selections.Count > 1)
                validSelection = false;
            TableCell selectedCell = null;
            if (validSelection)
            {
                selectedCell = doc.Tables.GetTableCell(doc.CaretPosition);
                validSelection = selectedCell != null;
            }
            if (!validSelection)
            {
                MessageBox.Show("Đặt trỏ chuột vào một thẻ con trong danh sách/trong bảng.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var field = _emrActionHelper.GetEmrtFieldAtPosition(doc.CaretPosition);
            if (field == null)
            {
                MessageBox.Show("Đặt trỏ chuột vào một thẻ con trong danh sách/trong bảng.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            int y, idx = 0;
            for (int i = field.CodeLevelArr.Length - 1; i >= 0; i--)
                if (int.TryParse(field.CodeLevelArr[i], out y))
                {
                    idx = i; break;
                }
            var prefix = string.Join(EmrParam.CodeSeparator.ToString(), field.CodeLevelArr.Take(idx));
            var fields = _emrDocumentHelper.GetAllDataFieldInRange(selectedCell.Table.Range);
            var document = (_entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo);
            var templateParams = AppMemCache.GetTemplateParams(document.FK_METemplateID);
            var data = _emrParser.ParserFieldsToJToken(fields, templateParams) as JToken;
            var codes = field.CodeLevelArr.Take(idx).Select(s => !int.TryParse(s, out y) ? s : "[0]");
            var raw = data.SelectToken(string.Join(".", codes)) as JArray;
            if (raw == null)
            {
                MessageBox.Show("Chức năng sắp xếp không hỗ trợ cho danh sách/bảng này.\nĐặt trỏ chuột vào một thẻ con trong danh sách/trong bảng.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var param = AppMemCache.GetParamFromDictKeyNo(field.CodeLevelArr[idx - 1]) as MEParamsInfo;
            if (param != null)
            {
                var childs = AppMemCache.GetParamRelationsFromDict(param.MEParamID).OrderBy(c => c.MEParamRelationOrder).ToList();
                var gui = new guiSortTableRows(param, raw, childs);
                if (gui.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var table = selectedCell.Table;
                        BOSProgressBar.Start("Đang xử lý dữ liệu");
                        var firstRowIdx = _emrDocumentHelper.FindTableRowIdx(table, prefix, string.Empty);
                        var secondRowIdx = _emrDocumentHelper.FindTableRowIdx(table, prefix, string.Empty, firstRowIdx + 1);
                        var padding = secondRowIdx - firstRowIdx;
                        var idxMapping = _emrDocumentHelper.GetTableRowIdxs(table, prefix, string.Empty);
                        var hasHiddenRow = padding > 1;
                        for (int newIdx = 0; newIdx < gui.FormatedData.Count; newIdx++)
                        {
                            var oldIdx = (int)(gui.FormatedData[newIdx] as JObject).Property("OriginalOrderNum").Value;
                            var oldRowIdx = _emrDocumentHelper.HasTableRow(table, prefix, idxMapping.ElementAt(oldIdx - 1).Key, new string[] { string.Empty });
                            var newRowIdx = firstRowIdx + (newIdx * padding);
                            _emrDocumentHelper.MoveTableRow(table, oldRowIdx, newRowIdx, hasHiddenRow);
                        }
                        if (hasHiddenRow)
                        {
                            var lastRow = table.Rows.Last();
                            foreach (var cell in lastRow.Cells)
                            {
                                cell.Borders.Bottom.LineStyle = TableBorderLineStyle.Single;
                                cell.Borders.Bottom.LineColor = Color.Black;
                                cell.Borders.Bottom.LineThickness = 1;
                            }
                            var rangePermissions = doc.BeginUpdateRangePermissions();
                            var rangePermissCount = rangePermissions.Count;
                            for (int i = 0; i < rangePermissCount; i++)
                            {
                                var rangePermis = rangePermissions[i];
                                var range = rangePermis.Range;
                                if (range.Start >= lastRow.Range.Start && range.End <= lastRow.Range.End)
                                {
                                    rangePermissions.Remove(rangePermis);
                                }
                            }
                            doc.EndUpdateRangePermissions(rangePermissions);
                        }
                        CloneSignedRangeForTracking()?.Wait(10000);

                        BOSProgressBar.Close();
                        MessageBox.Show("Nội dung đã được sắp xếp.\nVui lòng kiểm tra lại kết quả sắp xếp \n\nvà [THỰC HIỆN LƯU NGAY]. \n\nMở lại tờ nếu muốn undo.", "VUI LÒNG KIỂM TRA LẠI KẾT QUẢ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        ShowFlashNotification("Nội dung đã được sắp xếp.", 3000);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Có lỗi xảy ra khi sắp xếp. Mở lại tờ và thử lại?.\n\nChi tiết: " + ex.Message,
                            "CÓ LỖI XẢY RA. MỞ LẠI TỜ",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        InvalidateDocument(document, true);
                    }
                    finally
                    {
                        BOSProgressBar.Close();
                    }
                }
            }
        }

        #endregion

        #region Share Mode
        private bool IsShareEdit()
        {
            var editCap = BOSApp.GetUserEmrViewPermission();
            //[Vai trò = Quản trị] + [Quyền truy xuất = Toàn bộ] thì mới toàn quyền edit
            if (BOSApp.CurrentUserGroupInfo.ADUserGroupRole != UserGroupRole.admin.ToString())
                editCap = UserEmrView.DEPARTMENT;
            var emr = _entity.MainObject as MEEmrsInfo;
            var share = _emrCtrl.CheckEditPermission(emr.MEEmrID, BOSApp.CurrentEmployeesInfo.HREmployeeID, BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID, editCap);
            return share != null ? true : false;
        }
        #endregion

        #region Filter by Room
        public void FilterEmrsByRoom()
        {
            var notInRoom = "CHƯA XẾP PHÒNG";
            var rooms = _emrCtrl.DistinctRoomByDept(BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID);
            if (rooms.Any(r => string.IsNullOrEmpty(r)))
                rooms.Add(notInRoom);

            rooms = rooms.Where(r => !string.IsNullOrEmpty(r)).ToList();
            var all = "TẤT CẢ";
            rooms.Add(all);

            var gui = new guiSelectRoom(rooms)
            {
                Module = this
            };
            gui.InitializeControls(gui.Controls);
            gui.StartPosition = FormStartPosition.CenterParent;
            if (gui.ShowDialog() == DialogResult.OK)
            {
                var seletedRoom = gui.SelectedRoom;
                if (seletedRoom == all) seletedRoom = string.Empty;
                if (seletedRoom == notInRoom) seletedRoom = "NOT_IN_ROOM";
                try
                {
                    Cursor.Current = Cursors.WaitCursor;
                    DataSet ds = _emrCtrl.StayingByRoom(BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID, seletedRoom);
                    InvalidateSearchResult(ds);
                    var leftPanel = ParentScreen.SearchContainer;
                    if (leftPanel.Visibility == DockVisibility.AutoHide)
                    {
                        leftPanel.Show();
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }
                finally
                {
                    Cursor.Current = Cursors.Default;
                }
            }
        }
        #endregion

        #region Benh an mo lai ma quen dong
        private void RemainEmrReopenButNotClose()
        {
            //chi co user co quyen dong benh an moi chay chuc nang nay
            var barbtn = ParentScreen.GetToolbarButton("Close");
            if (barbtn == null) return;
            if (barbtn.Visibility == BarItemVisibility.Never) return;
            Task.Run(() =>
            {
                int days = _entity.GetConfigDaysRemainCloseEmr();
                if (days <= 0) return;
                var forgetCloses = _emrCtrl.GetCountEmrsForgetClose(days);
                if (forgetCloses > 0)
                {
                    if (MessageBox.Show($"Bệnh án được [mở lại] nhưng chưa đóng sau {days} ngày." +
                        $"\n- [Yes]: Tìm danh sách bệnh án chưa đóng và hiển thị trên [Danh sách đối tượng]" +
                        $"\n- [No]: Bỏ qua", "CÓ BỆNH ÁN [MỞ LẠI] NHƯNG CHƯA ĐÓNG", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        var ds = _emrCtrl.GetEmrsForgetClose(days);
                        InvalidateSearchResult(ds);
                    }
                }
            });
        }
        #endregion

        #region Document Notes
        internal void AddNewDocumentNote()
        {
            if (this._richEditCtrl.Modified)
            {
                var confirm = MessageBox.Show("Lưu thay đổi trước khi thực hiện ghi chú", "Nội dung tờ bệnh án đã thay đổi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var doc = _richEditCtrl.Document;
            var content = string.Join("", doc.Selections.Select(s => doc.GetText(s)));
            var gui = new guiDocumentNoteNew(content)
            {
                Module = this
            };
            _entity.SetDefaultModuleObject(TableName.MEEmrDocumentNotesTableName);
            var note = _entity.ModuleObjects[TableName.MEEmrDocumentNotesTableName] as MEEmrDocumentNotesInfo;
            note.FK_HRDepartmentID = BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID;
            note.FK_HREmployeeID = BOSApp.CurrentEmployeesInfo.HREmployeeID;
            gui.InitializeControls(gui.Controls);
            gui.StartPosition = FormStartPosition.CenterParent;
            if (gui.ShowDialog() == DialogResult.OK)
            {
                var bookmark = string.Empty;
                if (!string.IsNullOrEmpty(content))
                {
                    bookmark = $"doc-note-{BOSApp.CurrentUser}-{DateTime.Now.ToString("ddMMyyyy-hhmmssfff")}";
                    doc.Bookmarks.Create(doc.Selections.First(), bookmark);
                    var command = this._richEditCtrl.CreateCommand(RichEditCommandId.FileSave);
                    command.Execute();
                }
                var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
                var emr = _entity.MainObject as MEEmrsInfo;
                var documentNoteCtrl = new MEEmrDocumentNotesController();
                note.MEEmrDocumentNoteTime = DateTime.Now;
                note.FK_MEEmrDocumentID = document.MEEmrDocumentID;
                note.MEEmrDocumentNoteBookmark = bookmark;
                note.AACreatedUser = BOSApp.CurrentUser;
                note.FK_MEEmrID = emr.MEEmrID;
                note.FK_METemplateID = document.FK_METemplateID;
                documentNoteCtrl.CreateObject(note);
                MessageBox.Show("Đã thêm", "Thêm ghi chú", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        public void ShowDocumentNotes()
        {
            var emr = _entity.MainObject as MEEmrsInfo;
            var gui = new guiDocumentNoteList(emr.MEEmrID)
            {
                Module = this
            };
            gui.InitializeControls(gui.Controls);
            gui.StartPosition = FormStartPosition.Manual;
            gui.Location = new Point(0, 0);
            gui.Show(this.ParentScreen);
        }
        public void GotoBookmark(string name, int documentId = 0)
        {
            var doc = _richEditCtrl.Document;
            var bookmark = doc.Bookmarks.Where(b => b.Name == name).FirstOrDefault();
            if (bookmark != null)
            {
                doc.Bookmarks.Select(bookmark);
                _richEditCtrl.ScrollToCaret();
            }
            else if (documentId > 0)
            {
                var emrDoc = _emrDocumentCtrl.GetObjectByID(documentId) as MEEmrDocumentsInfo;
                InvalidateDocument(emrDoc, asyncMode: true);
                GotoBookmark(name);
            }
            else
                MessageBox.Show("Không tìm thấy vị trí đánh dấu", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }
        public void DeleteNote()
        {
            var documentNoteCtrl = new MEEmrDocumentNotesController();
            MEEmrsInfo emr = _entity.MainObject as MEEmrsInfo;
            if (emr.MEEmrID > 0 && emr.MEEmrHasNote)
            {
                if (MessageBox.Show($"Muốn xóa ghi chú bệnh án {emr.MEEmrNo}?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    documentNoteCtrl.DeleteByForeignColumn("FK_MEEmrID", emr.MEEmrID);
                    MessageBox.Show($"Xóa ghi chú bệnh án {emr.MEEmrNo} thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
        #endregion

        private void StuckAtInitingState()
        {
            var meEmr = _entity.MainObject as MEEmrsInfo;
            if (meEmr.MEEmrStatus != EmrStatus.Initing.ToString()) return;

            var minuteStuckMax = BOSApp.GetSystemConfigValue(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.MAX_STUCK_AT_INITING_STATE_INTERVAL);
            if (string.IsNullOrEmpty(minuteStuckMax)) return;

            try
            {
                var minuteStuckMaxDouble = Convert.ToDouble(minuteStuckMax);
                if (minuteStuckMaxDouble == 0) return;

                var minuteStuck = DateTime.Now.Subtract(meEmr.AACreatedDate).TotalMinutes;
                if (minuteStuck > minuteStuckMaxDouble)
                {
                    var confirm = MessageBox.Show("Bệnh án đang treo ở trạng thái Đang được khởi tạo.\n"
                                + "Thực hiện khởi tạo bệnh án thủ công?\n"
                                + "- Yes: Khởi tạo\n"
                                + "- No: Để sau\n",
                                "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (confirm == DialogResult.Yes)
                    {
                        // Update emr InProgress (TakeInitPermission())
                        meEmr = _emrCtrl.GetObjectByID(meEmr.MEEmrID) as MEEmrsInfo;
                        // Check status again, case changed in db => Thong bao nguoi dung
                        if (meEmr.MEEmrStatus != EmrStatus.Initing.ToString())
                        {
                            var msg = "Thông tin bệnh án này đã bị thay đổi lúc " + meEmr.AAUpdatedDate.ToString("dd/MM/yyyy HH:mm:ss") + ". Vui lòng làm mới dữ liệu.";
                            MessageBox.Show(msg, "THÔNG TIN BỆNH ÁN ĐÃ THAY ĐỔI", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }

                        meEmr.MEEmrStatus = EmrStatus.InProgress.ToString();
                        meEmr.AAUpdatedUser = BOSApp.CurrentUser;
                        _emrCtrl.UpdateObject(meEmr);
                        HistoryEmr(meEmr, cstObjectHistoryActionChange, $"thay đổi thông tin bệnh án - khởi tạo bệnh án thủ công.");
                        _entity.MainObject = meEmr;

                        // Do action Init (btnInitEmr_Click)
                        InitEmr();
                    }
                }
            }
            catch (Exception ex)
            {
                Trace.TraceError("StuckAtInitingState ERROR: {0}:{1}:{2}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), BOSApp.CurrentUser, ex);
                Trace.Flush();
            }
        }

        #region Lien Ket Benh An - HTSS
        public void EmrRelation()
        {
            var gui = new guiSelectEmrRelation
            {
                Module = this,
                StartPosition = FormStartPosition.CenterParent
            };
            if (gui.ShowDialog() == DialogResult.OK)
            {
                var selectedEmrID = gui.SelectedEmrID;
                var selectedEmrNo = gui.SelectedEmrNo;
                var selectedEmrRelationFromName = gui.SelectedEmrRelationFromName;
                var selectedEmrRelationToName = gui.SelectedEmrRelationToName;
                if (selectedEmrID == 0)
                {
                    var emrSelected = _emrCtrl.GetObjectByNo(selectedEmrNo) as MEEmrsInfo;
                    selectedEmrID = emrSelected.MEEmrID;
                }

                var emrRelationCtrl = new MEEmrRelationsController();
                var meEmr = _entity.MainObject as MEEmrsInfo;
                if (meEmr.MEEmrID == selectedEmrID)
                {
                    MessageBox.Show($"Bệnh án trùng nhau, không thể liên kết. Vui lòng chọn bệnh án khác!!!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var relations = emrRelationCtrl.GetListBusinessObjects<MEEmrRelationsInfo>(emrRelationCtrl.GetAllDataByForeignColumn("FK_MEEmrFromID", meEmr.MEEmrID));
                if (relations.Count() > 0)
                {
                    var existE = relations.FirstOrDefault(m => m.FK_MEEmrToID.Equals(selectedEmrID));
                    if (existE != null)
                    {
                        MessageBox.Show($"Bệnh án liên kết {selectedEmrNo} đã được liên kết.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }

                var now = DateTime.Now;
                var relation = new MEEmrRelationsInfo()
                {
                    FK_MEEmrFromID = meEmr.MEEmrID,
                    FK_MEEmrToID = selectedEmrID,
                    MEEmrRelationFromName = selectedEmrRelationFromName,
                    MEEmrRelationToName = selectedEmrRelationToName,
                    MEEmrRelationDate = now,
                    MEEmrRelationRemark = string.Empty,
                    AACreatedUser = BOSApp.CurrentUser
                };

                emrRelationCtrl.CreateObject(relation);

                var emrRelation = _emrCtrl.GetObjectByID(selectedEmrID) as MEEmrsInfo;
                HistoryEmr(meEmr, "Change", $"Liên kết với bệnh án Id {selectedEmrID}");
                HistoryEmr(emrRelation, "Change", $"Liên kết với bệnh án Id {meEmr.MEEmrID}");

                Invalidate(meEmr.MEEmrID);
            }
        }

        private void EmrRelationInit()
        {
            var meEmr = _entity.MainObject as MEEmrsInfo;
            var relationEmrs = _emrCtrl.GetListBusinessObjects<MEEmrsInfo>(_emrCtrl.GetAllWithRelation(meEmr.MEEmrID));

            if (relationEmrs.Count() > 0)
            {
                var tabEmrRelations = (Controls["tabEmrRelations"] as DevExpress.XtraTab.XtraTabPage);
                tabEmrRelations.PageVisible = true;

                var gridControlEmrRelations = this.Controls["fld_dgcMEEmrRelations"] as MEEmrRelationsGridControl;
                if (gridControlEmrRelations != null)
                {
                    gridControlEmrRelations.DataSource = relationEmrs;
                    gridControlEmrRelations.RefreshDataSource();
                    gridControlEmrRelations.Refresh();
                }

            // Clean
            (this.Controls["fld_lblRelation"] as Label).Text = string.Empty;
                var gridControlEmrRelationDocuments = this.Controls["fld_dgcMEEmrRelationDocuments"] as MEEmrRelationDocumentsGridControl;
                if (gridControlEmrRelationDocuments != null)
                {
                    gridControlEmrRelationDocuments.DataSource = new List<MEEmrDocumentsInfo>();
                    gridControlEmrRelationDocuments.RefreshDataSource();
                    gridControlEmrRelationDocuments.Refresh();
                }
            }
        }

        public void EmrRelationAction(int emrRelationId)
        {
            var emrRelationName = EmrRelationName(emrRelationId);
            (this.Controls["fld_lblRelation"] as Label).Text = $"Bệnh án - {emrRelationName}";
            var gridControlEmrRelationDocuments = this.Controls["fld_dgcMEEmrRelationDocuments"] as MEEmrRelationDocumentsGridControl;
            if (gridControlEmrRelationDocuments != null)
            {
                var emrRelationDocuments = _emrDocumentCtrl.GetListBusinessObjects<MEEmrDocumentsInfo>(_emrDocumentCtrl.GetAllDataByForeignColumn("FK_MEEmrID", emrRelationId));
                gridControlEmrRelationDocuments.DataSource = emrRelationDocuments;
                gridControlEmrRelationDocuments.RefreshDataSource();
                gridControlEmrRelationDocuments.Refresh();
            }
        }

        private void RelationSubmit(MEEmrsInfo emr)
        {
            var emrRelationCtrl = new MEEmrRelationsController();
            var relations = emrRelationCtrl.GetListBusinessObjects<MEEmrRelationsInfo>(emrRelationCtrl.GetAllDataByForeignColumn("FK_MEEmrFromID", emr.MEEmrID));
            if (relations.Count() > 0)
            {
                var msg = $"YES: Copy tất cả các tờ bệnh án từ bệnh án liên quan sang bệnh án hiện tại và đóng đồng thời các bệnh án này cùng lúc. {Environment.NewLine}";
                msg += $"NO: Không copy và đóng độc lập các bệnh án.";
                var confirm = MessageBox.Show(msg, "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (confirm == DialogResult.Yes)
                {
                    #region Gáy gốc
                    // Trường hợp nhiều mối quan hệ: hiện tại chưa phát sinh nên lấy First().
                    // [METemplateIndexRelationCombine] false : thêm mối quan hệ trước gáy.
                    // [METemplateIndexRelationCombine] true : No action. Mục đích map với gáy Bệnh án quan hệ
                    var relationFromNameText = ADConfigValueGetText("EmrRelationFromName", relations.First().MEEmrRelationFromName);
                    var documents = _emrDocumentCtrl.GetByEmrId(emr.MEEmrID);
                    var mapSourceTemplateIndexs = new List<METemplateIndexsInfo>(); // Relation cùng check
                    foreach (var document in documents)
                    {
                        var templateIndex = _templateIndexsCtrl.GetObjectInfoByNameAndEmrType(document.MEEmrDocumentGroup, emr.FK_MEEmrTypeID);
                        if (templateIndex != null)
                        {
                            if (!templateIndex.METemplateIndexRelationCombine)
                            {
                                document.MEEmrDocumentGroup = $"{relationFromNameText}: {document.MEEmrDocumentGroup}";
                                _emrDocumentCtrl.UpdateObject(document);

                                var documentCode = getDocumentNo(document);
                                HistoryEmr(emr, "Change", $"{documentCode} thay đổi nhóm tờ bệnh án. [Mối quan hệ]");
                            }
                            else
                            {
                                // Relation cùng check
                                if (!mapSourceTemplateIndexs.Any(item => item.METemplateIndexID == templateIndex.METemplateIndexID))
                                {
                                    mapSourceTemplateIndexs.Add(templateIndex);
                                }
                            }
                        }
                    }
                    #endregion

                    foreach (var relation in relations)
                    {
                        // STEP 1: Copy tờ bệnh án qua bệnh án hiện tại. 
                        // STEP 2: Sắp xếp đúng gáy.
                        var relationID = relation.FK_MEEmrToID;
                        var relationEmr = _emrCtrl.GetObjectByID(relationID) as MEEmrsInfo;
                        if (relationEmr != null)
                        {
                            // TẬN DỤNG PDF.
                            var msgCloseRelation = CloseEmrRelation(relationEmr);
                            if (msgCloseRelation != EmrStatus.Closed.ToString())
                            {
                                if (_notificationTab)
                                {
                                    _msgLogs.Text += $"\r\n {msgCloseRelation}";
                                }
                                else if (_logConfig)
                                {
                                    _msgLogsTemp += $"\r\n {msgCloseRelation}";
                                }
                            }
                            else
                            {
                                RefreshCurrentEmrSearchResultsControl(relationEmr);
                            }

                            var relationToNameText = ADConfigValueGetText("EmrRelationToName", relation.MEEmrRelationToName);
                            var relationDocuments = _emrDocumentCtrl.GetByEmrId(relationEmr.MEEmrID);
                            foreach (var relationDocument in relationDocuments)
                            {
                                if (!documents.Any(m => m.MEEmrDocumentFile.Equals(relationDocument.MEEmrDocumentFile)))
                                {
                                    if (!string.IsNullOrEmpty(relationDocument.MEEmrDocumentGroup)) // FK_METemplateID
                                    {
                                        var indexTemplate = _templateIndexsCtrl.GetObjectInfoByNameAndEmrType(relationDocument.MEEmrDocumentGroup, relationEmr.FK_MEEmrTypeID);
                                        if (indexTemplate != null)
                                        {
                                            // Exception: Manual on top
                                            if (indexTemplate.METemplateIndexRelationOrder > 0 || !string.IsNullOrEmpty(indexTemplate.METemplateIndexRelationName))
                                            {
                                                relationDocument.MEEmrDocumentOrder = indexTemplate.METemplateIndexRelationOrder;
                                                relationDocument.MEEmrDocumentGroup = $"{indexTemplate.METemplateIndexRelationName}";
                                            }
                                            else
                                            {
                                                if (!indexTemplate.METemplateIndexRelationCombine)
                                                {
                                                    relationDocument.MEEmrDocumentGroup = $"{relationToNameText}: {relationDocument.MEEmrDocumentGroup}";
                                                }
                                                else
                                                {
                                                    // Map with source: bỏ vô gáy vợ. Mà cả hai sẽ giống nhau => Ko làm gì
                                                    // Nếu ko map thì gắn gáy quan hệ phía trước.
                                                    if (!mapSourceTemplateIndexs.Any(item => item.METemplateIndexName.Trim().Equals(relationDocument.MEEmrDocumentGroup.Trim())))
                                                    {
                                                        relationDocument.MEEmrDocumentGroup = $"{relationToNameText}: {relationDocument.MEEmrDocumentGroup}";
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    relationDocument.MEEmrDocumentID = 0;
                                    relationDocument.FK_MEEmrID = emr.MEEmrID;
                                    relationDocument.AACreatedUser = BOSApp.CurrentUser;

                                    var from = $"/Emr/{relationID}/{relationDocument.MEEmrDocumentFile}.{relationDocument.MEEmrDocumentFileExt}";
                                    var to = $"/Emr/{emr.MEEmrID}/{relationDocument.MEEmrDocumentFile}.{relationDocument.MEEmrDocumentFileExt}";
                                    var localPath = string.Format(@"{0}\Emr\{1}\{2}.{3}", _documentPath, emr.MEEmrID, relationDocument.MEEmrDocumentFile, relationDocument.MEEmrDocumentFileExt);
                                    _ftpFileMng.CopyFile(from, to, localPath);

                                    _emrDocumentCtrl.CreateObject(relationDocument);

                                    var documentCode = getDocumentNo(relationDocument);
                                    HistoryEmr(emr, "Change", $"{documentCode} thêm tờ bệnh án từ bệnh án {relation.MEEmrRelationFromName}. [Mối quan hệ]");
                                }
                            }
                        }
                    }
                }
            }
        }

        private MEEmrRelationsInfo EmrRelationInfo(int relationEmrId)
        {
            var meEmr = _entity.MainObject as MEEmrsInfo;
            var emrRelationCtrl = new MEEmrRelationsController();

            // From
            var relationFroms = emrRelationCtrl.GetListBusinessObjects<MEEmrRelationsInfo>(emrRelationCtrl.GetAllDataByForeignColumn("FK_MEEmrFromID", relationEmrId));
            var relationFrom = relationFroms.FirstOrDefault(m => m.FK_MEEmrToID.Equals(meEmr.MEEmrID));
            if (relationFrom != null) return relationFrom;

            // To
            var relationTos = emrRelationCtrl.GetListBusinessObjects<MEEmrRelationsInfo>(emrRelationCtrl.GetAllDataByForeignColumn("FK_MEEmrToID", relationEmrId));
            var relationTo = relationTos.FirstOrDefault(m => m.FK_MEEmrFromID.Equals(meEmr.MEEmrID));
            return relationTo;
        }

        private string EmrRelationName(int relationEmrId)
        {
            var emrRelation = EmrRelationInfo(relationEmrId);
            if (emrRelation == null) return string.Empty;

            if (relationEmrId == emrRelation.FK_MEEmrFromID)
            {
                return ADConfigValueGetText("EmrRelationFromName", emrRelation.MEEmrRelationFromName);
            }
            else
            {
                return ADConfigValueGetText("EmrRelationToName", emrRelation.MEEmrRelationToName);
            }
        }

        private string ADConfigValueGetText(string group, string key)
        {
            List<ADConfigValuesInfo> adConfigs = _objConfigValuesController.GetConfigValuesByGroup(group);
            var configE = adConfigs.Where(m => m.ADConfigKeyValue.Equals(key)).FirstOrDefault();
            return configE == null ? key : configE.ADConfigText;
        }
        #endregion

        private string CloseEmrRelation(MEEmrsInfo emr)
        {
            #region IsEmrReadOnly(true) mix winform vs api
            // BA đã đóng thì copy
            if (emr.MEEmrStatus == EmrStatus.Closed.ToString()) return EmrStatus.Closed.ToString();

            if (emr.MEEmrTypeProfile == EmrTypeProfile.Patient.ToString()) return "Không thể đóng hồ sơ bệnh nhân.";

            #endregion

            var listDocs = _emrDocumentCtrl.GetListBusinessObjects<MEEmrDocumentsInfo>(_emrDocumentCtrl.GetByMEEmrID(emr.MEEmrID));
            foreach (var doc in listDocs)
            {
                if (doc.FK_EditingUserID > 0 && doc.FK_EditingUserID != BOSApp.CurrentEmployeesInfo.HREmployeeID)
                {
                    return $"Tờ bệnh án {doc.MEEmrDocumentGroup} của bệnh án {emr.MEEmrNo} đang được soạn bởi người khác hoặc máy khác. Không xử lý bệnh án này.";
                }
            }

            #region ValidateBeforeActionEmr
            var config = BOSApp.GetSystemConfigValue(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.EMR_DOCUMENT_STATUS_VERIFY).ToUpper();
            var actionConfig = config.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries).Select(p => p.Trim()).ToList();
            if (!string.IsNullOrEmpty(config) && actionConfig.Contains(EmrStatus.Closed.ToString().ToUpper()))
            {
                var documentsMissData = new List<MEEmrDocumentsInfo>();
                var docs = listDocs.Where(m => m.MEEmrDocumentStatus.Equals(EmrStatus.InProgress.ToString())).ToList();
                foreach (var document in docs)
                {
                    if (document.MEEmrDocumentFileExt != EmrDocumentFileExtention.docx.ToString()) continue;

                    var missingDataParam = ValidateDocumentRequired(document, false);
                    if (missingDataParam.Count > 0)
                    {
                        var msgValidates = string.Join("\n", missingDataParam.Select(x => "<" + x.Value + ">").ToArray());
                        document.MEEmrDocumentDesc = $"{msgValidates} chưa có dữ liệu.";
                        documentsMissData.Add(document);
                    }
                }

                if (documentsMissData != null && documentsMissData.Count > 0)
                {
                    return $"Bệnh án {emr.MEEmrNo} có tờ thiếu dữ liệu. Không xử lý bệnh án này.";
                }
            }
            #endregion


            this._dpnViewPdf.Visible = false;
            this._dpnRichEdit.Visible = true;
            this.DockManager.ActivePanel = this._dpnRichEdit;
            emr.AllowPropertyChangedEvent = false;
            var currentItem = new MEEmrDocumentsInfo(); // use for UndoCloseEmrDocument()
            try
            {
                int totalClosed = 0;
                var serverPath = $"/Emr/{emr.MEEmrID}/";
                var serverFiles = _ftpFileMng.GetNameListing(serverPath);
                var localPath = string.Format(@"{0}\Emr\{1}\", _documentPath, emr.MEEmrID);
                DirectoryInfo localDir = new DirectoryInfo(localPath);
                var pdf = EmrDocumentFileExtention.pdf.ToString();
                var docx = EmrDocumentFileExtention.docx.ToString();
                var hashFiles = new List<string>();
                foreach (var item in listDocs)
                {
                    if (item.MEEmrDocumentStatus == EmrDocumentStatus.Discarded.ToString()
                        || item.MEEmrDocumentStatus == EmrDocumentStatus.Hidden.ToString())
                    {
                        hashFiles.Add(item.MEEmrDocumentFile);
                        totalClosed++;
                        continue;
                    }
                    var fileNamePdf = $"{item.MEEmrDocumentFile}.{pdf}";
                    var fileNameDocx = $"{item.MEEmrDocumentFile}.{docx}";
                    var fullFileNamePdfLocal = string.Format(@"{0}\{1}", localPath, fileNamePdf);
                    var fullFileNameDocxLocal = string.Format(@"{0}\{1}", localPath, fileNameDocx);
                    if (item.MEEmrDocumentFileExt == docx)
                    {
                        _ftpFileMng.DownloadFile(serverPath, fileNameDocx, fullFileNameDocxLocal);

                        if (!CheckDataIntegrityFromLastDigitalSignDocx(item, item.MEEmrDocumentFile)) continue;

                        currentItem = item.Clone() as MEEmrDocumentsInfo;

                        #region CloseEmrDocument(item)
                        item.MEEmrDocumentEndDate = DateTime.Now;
                        item.MEEmrDocumentPreStatus = item.MEEmrDocumentStatus;
                        item.MEEmrDocumentStatus = EmrDocumentStatus.Closed.ToString();
                        if (item.MEEmrDocumentFileExt == EmrDocumentFileExtention.docx.ToString())
                            item.MEEmrDocumentFileExt = EmrDocumentFileExtention.pdf.ToString();
                        item.MEEmrDocumentJson = string.Empty;//this.ParserDocumentToJson(); uthv cham qua roi

                        ClearCurrentEditingUser(item);

                        UpdateMongoDocument(emr, item, new List<string>() { "MEEmrDocumentContent" });
                        item.AAUpdatedUser = BOSApp.CurrentUser;
                        _emrDocumentCtrl.UpdateObject(item);
                        #endregion

                        var hash = _md5Hasher.ComputeHash(fullFileNameDocxLocal);
                        var filePdfHash = $"{item.MEEmrDocumentFile}_MD5{hash}.{pdf}";
                        var existFileHash = serverFiles.Where(stringToCheck => stringToCheck.Contains(filePdfHash));
                        if (existFileHash.Count() == 0)
                        {
                            if (File.Exists(Path.Combine(localPath, filePdfHash)))
                            {
                                File.Delete(fullFileNamePdfLocal);
                                File.Move(Path.Combine(localPath, filePdfHash), fullFileNamePdfLocal);
                            }
                            else
                            {
                                this.OpenEmrDocument(item.FK_MEEmrID, item.MEEmrDocumentFile, false, false);
                                this.RemoveAllForPrint(this._richEditCtrl, item.FK_METemplateID);
                                this._richEditCtrl.ExportToPdf(fullFileNamePdfLocal);
                            }
                            _ftpFileMng.UploadFile(serverPath, fileNamePdf, fullFileNamePdfLocal);
                        }
                        else
                        {
                            _ftpFileMng.MoveFile(serverPath, filePdfHash, fileNamePdf);
                        }

                        hashFiles.Add(item.MEEmrDocumentFile);
                    }
                    else
                    {
                        if (!_ftpFileMng.FileExists($"/Emr/{item.FK_MEEmrID}/", fileNamePdf))
                        {
                            throw new FtpException();
                        }

                        item.MEEmrDocumentEndDate = DateTime.Now;
                        item.MEEmrDocumentPreStatus = item.MEEmrDocumentStatus;
                        item.MEEmrDocumentStatus = EmrDocumentStatus.Closed.ToString();
                        if (item.MEEmrDocumentFileExt == EmrDocumentFileExtention.docx.ToString())
                            item.MEEmrDocumentFileExt = EmrDocumentFileExtention.pdf.ToString();
                        item.AAUpdatedUser = BOSApp.CurrentUser;
                        _emrDocumentCtrl.UpdateObject(item);
                    }

                    var documentCodeU = getDocumentNo(item);
                    HistoryEmr(GetCurrentMainObject(), "Change", $"{documentCodeU} đóng tờ bệnh án.");

                    totalClosed++;
                }

                listDocs = _emrDocumentCtrl.GetListBusinessObjects<MEEmrDocumentsInfo>(_emrDocumentCtrl.GetByMEEmrID(emr.MEEmrID));
                if (totalClosed == listDocs.Count)
                {
                    // them cot nguoi dong benh an theo yeu cau cua DKLK
                    emr.FK_HREmployeeClosedID = BOSApp.CurrentEmployeesInfo.HREmployeeID;
                    emr.MEEmrStatus = EmrStatus.Closed.ToString();
                    emr.MEEmrEndDate = DateTime.Now;
                    emr.AAUpdatedUser = BOSApp.CurrentUser;
                    this._emrCtrl.UpdateObject(emr);
                    HistoryEmr(emr, cstObjectHistoryActionChange, $"đóng bệnh án.");
                    InactiveShareEmr(emr.MEEmrID);

                    // Delete file hash
                    serverFiles = _ftpFileMng.GetNameListing(serverPath); // get newest
                    foreach (var hashFile in hashFiles)
                    {
                        ClearHashFileServer(serverPath, serverFiles, hashFile, string.Empty);
                        ClearHashFileLocal(localDir, hashFile);
                    }
                    return EmrStatus.Closed.ToString();
                }
                else
                {
                    return "Đóng bệnh án không thành công. Thực hiện đóng từng tờ bệnh án còn mở.";
                }
            }
            catch (Exception ex)
            {
                UndoCloseEmrDocument(currentItem);
                return $"Đóng bệnh án không thành công. Chi tiết lỗi: {Environment.NewLine} {ex}.";
            }
        }

        #region TEMP - TOOL
        public void ReOpenEmrDocumentEmr()
        {
            var gridView = _entity.MEEmrDocumentsList.GridView;
            if (gridView.FocusedRowHandle < 0)
            {
                MessageBox.Show("Vui lòng chọn tờ bệnh án");
                return;
            }
            var document = gridView.GetRow(gridView.FocusedRowHandle) as MEEmrDocumentsInfo;
            if (document.MEEmrDocumentStatus == EmrDocumentStatus.InProgress.ToString())
            {
                MessageBox.Show("Không áp dụng tờ bệnh án đang mở.");
                return;
            }
            if (document.MEEmrDocumentStatus == EmrDocumentStatus.Hidden.ToString())
            {
                MessageBox.Show("Không áp dụng tờ bệnh án đang ẩn.");
                return;
            }

            guiSign guiSign = new guiSign("Đang thực hiện mở lại tờ bệnh án + bệnh án. Sẽ không thể hoàn tác. Vẫn thực hiện?", _fingerPrintZK);
            if (guiSign.ShowDialog() == DialogResult.OK)
            {
                string serverPath = $"/Emr/{document.FK_MEEmrID}/";
                string localPath = string.Format(@"{0}\Emr\{1}\", _documentPath, document.FK_MEEmrID);
                var pdf = EmrDocumentFileExtention.pdf.ToString();
                var docx = EmrDocumentFileExtention.docx.ToString();
                // Find docx
                var fileNamePdf = $"{document.MEEmrDocumentFile}.{pdf}";
                var fileNameDocx = $"{document.MEEmrDocumentFile}.{docx}";

                if (!_ftpFileMng.FileExists(serverPath, fileNameDocx))
                {
                    if (document.MEEmrDocumentInitFileExt == docx)
                    {
                        MessageBox.Show($"Không thể thực hiện mở lại tờ bệnh án này.{Environment.NewLine}Không tìm thấy tờ bệnh án docx.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show($"Không thể thực hiện mở lại tờ bệnh án này. {Environment.NewLine}Tờ bệnh án gốc không phải file docx.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    return;
                }

                // backup pdf vs docx
                string filePdfLocalFullPath = Path.Combine(localPath, fileNamePdf);
                var fromCp = Path.Combine(serverPath, fileNamePdf);
                var toCp = Path.Combine(serverPath, $"{document.MEEmrDocumentFile}_bkReOpen{DateTime.Now.ToString("MM-dd-yyyy")}.{pdf}");
                _ftpFileMng.CopyFile(fromCp, toCp, filePdfLocalFullPath);

                string fileDocxLocalFullPath = Path.Combine(localPath, fileNameDocx);
                var fromDocxCp = Path.Combine(serverPath, fileNameDocx);
                var toDocxCp = Path.Combine(serverPath, $"{document.MEEmrDocumentFile}_bkReOpen{DateTime.Now.ToString("MM-dd-yyyy")}.{docx}");
                _ftpFileMng.CopyFile(fromDocxCp, toDocxCp, fileDocxLocalFullPath);

                document.MEEmrDocumentPreStatus = document.MEEmrDocumentStatus;
                document.MEEmrDocumentStatus = EmrDocumentStatus.InProgress.ToString();
                document.MEEmrDocumentFileExt = docx;
                document.AAUpdatedUser = BOSApp.CurrentUser;
                _emrDocumentCtrl.UpdateObject(document);

                //var documentCodeU = getDocumentNo(document);
                //HistoryEmr(GetCurrentMainObject(), "Change", $"{documentCodeU} mở lại tờ bệnh án.");

                // Open emr
                var emr = _emrCtrl.GetObjectByID(document.FK_MEEmrID) as MEEmrsInfo;
                emr.AAUpdatedUser = BOSApp.CurrentUser;
                emr.MEEmrStatus = EmrStatus.InProgress.ToString();
                emr.MEEmrDesc = $"{emr.MEEmrDesc}. {BOSApp.CurrentEmployeesInfo.HREmployeeName} mở lại tờ bệnh án";
                _emrCtrl.UpdateObject(emr);

                HistoryEmr(emr, cstObjectHistoryActionReOpen, $"Mở lại bệnh án - làm mới tờ bệnh án.");

                #region TTBA
                if (_sysHelper.AllowThread())
                {
                    var threadSum = new System.Threading.Thread(() => UpdateEmrSumWhenOpenEmr(emr));
                    threadSum.Start();
                }
                else
                {
                    UpdateEmrSumWhenOpenEmr(emr);
                }
                #endregion

                // Check archive
                var archivesCtrl = new MEEmrArchivesController();
                var archive = archivesCtrl.GetObjectLastestByEmrId(emr.MEEmrID);
                if (archive != null)
                {
                    MessageBox.Show($"Bệnh án {emr.MEEmrNo} có thông tin lưu trữ {archive.MEEmrArchiveStatus.ToUpper()}.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                MessageBox.Show($"Mở lại tờ bệnh án, bệnh án thành công.{Environment.NewLine}Bấm vô tờ bệnh án để thao tác.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);

                _entity.MainObject = emr;
                RefreshCurrentEmrSearchResultsControl(emr);
                Invalidate(emr.MEEmrID);
                InvalidateEmrDocumentList(document.FK_MEEmrID);
            }
        }

        #endregion

        #region Benh nhan ky van tay và ky signpad 
        internal void PatientSign()
        {
            if (this._richEditCtrl.Modified)
            {
                var confirm = MessageBox.Show("Lưu thay đổi trước khi thực hiện ký", "Nội dung tờ bệnh án đã thay đổi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (IsEmrReadOnly()) return;
            if (IsForceRevokePermission()) return;

            // Divide fingerprint or signpad
            var stPad = new STPadLib(); // SignPad
            DPUruNet.Reader reader; // FingerPrint

            var signMode = 0; // 0: FingerPrint; 1 : SignPad
            if (!_useZKTeco && _useSignPadSignotec)
            {
                signMode = 1;
            }
            if (_useZKTeco && _useSignPadSignotec)
            {
                // get ready devide
                var signPadDevide = true;
                var fingerPrintDevide = true;
                int signPadCount = stPad.DeviceGetCount();
                if (signPadCount <= 0)
                {
                    signPadDevide = false;
                }

                if (!_useZKTeco)
                {
                    reader = _fingerPrint.GetFirst();
                    if (reader == null)
                    {
                        fingerPrintDevide = false;
                    }
                }
                else if (_fingerPrintZK.GetDeviceCount() <= 0)
                {
                    fingerPrintDevide = false;
                }

                if (fingerPrintDevide && !signPadDevide)
                {
                    signMode = 0;
                }
                else if (!fingerPrintDevide && signPadDevide)
                {
                    signMode = 1;
                }
                else
                {
                    var confirm = MessageBox.Show("Vui lòng chọn thiết bị để ký.\n"
                        + "- Yes: Sử dụng thiết bị ký vân tay - FingerPrint\n"
                        + "- No: Sử dụng thiết bị ký tên - SignPad\n"
                        + "- Cancel: Không tiếp tục thao tác.",
                        "Thông báo", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    if (confirm == DialogResult.Yes)
                    {
                        signMode = 0;
                    }
                    else if (confirm == DialogResult.No)
                    {
                        signMode = 1;
                    }
                    else
                    {
                        return;
                    }
                }
            }

            if (signMode == 0)
            {
                try
                {
                    if (!_useZKTeco)
                    {
                        reader = _fingerPrint.GetFirst();
                        if (reader == null)
                        {
                            MessageBox.Show("Vui lòng gắn thiết bị đọc vân tay vào cổng USB.", "KHÔNG TÌM THẤY THIẾT BỊ ĐỌC VÂN TAY", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            return;
                        }
                        _fingerPrint.SetReader(reader);
                    }
                    else if (_fingerPrintZK.GetDeviceCount() <= 0)
                    {
                        MessageBox.Show("Vui lòng gắn thiết bị đọc vân tay vào cổng USB.", "KHÔNG TÌM THẤY THIẾT BỊ ĐỌC VÂN TAY", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }
                    else
                    {
                        _fingerPrintZK.OpenDevice(0);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Chức năng không khả dụng. Vui lòng gắn thiết bị và cài đặt SDK. \n" + ex.Message, "CHỨC NĂNG KHÔNG KHẢ DỤNG", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    return;
                }
            }
            else
            {
                // _useSignPadSignotec
                if (_useSignPadSignotec)
                {
                    int deviceCount = stPad.DeviceGetCount();
                    if (deviceCount <= 0)
                    {
                        MessageBox.Show("Vui lòng gắn thiết bị ký tên SignPad vào cổng USB.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        return;
                    }
                }
            }

            var doc = this._richEditCtrl.Document;
            var range = doc.Range;
            var selectedRanges = new List<DocumentRange>(doc.Selections);
            if (selectedRanges.Count == 1 && selectedRanges[0].Length == 0)
            {
                selectedRanges.Clear();
                selectedRanges.Add(doc.Range);
            }
            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;

            try
            {
                //neu co chon mot the
                //neu co quet chon chu ky
                //neu ko quet chon thi hien popup de chon
                var signRoles = GetSignerRoles(selectedRanges, false);
                if (signRoles == null) return;
                var defaultSignerRela = string.Empty;
                var alternativeSign = false;
                if (signRoles.Count <= 0)
                {
                    //neu khong co chu ky nao tren mau thi chen vao vi tri con tro
                    if (MessageBox.Show("Chèn vân tay vào vị trí con trỏ?", "KHÔNG TÌM THẤY THẺ CHỮ KÝ", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK)
                        return;
                }
                else
                {
                    defaultSignerRela = AppMemCache.GetParamFromDictKeyNo(signRoles.First().Key)?.MEParamValue;
                    var code = signRoles.First().Value?.Select(f => f.FieldCode).FirstOrDefault();
                    if (!string.IsNullOrEmpty(code))
                    {
                        var path = code.Split(EmrParam.TagCodeSeparator)[0];
                        var idx = path.LastIndexOf(EmrParam.CodeSeparator);
                        if (idx > 0)
                            path = path.Substring(0, idx);
                        path = _emrDocumentHelper.GetTemplateFieldPath(path);
                        if (AppMemCache.GetTemplateParamsDictPath(document.FK_METemplateID).TryGetValue(path, out METemplateParamsInfo config))
                            alternativeSign = config.METemplateParamAlternativeSign;
                    }
                }

                var contentHash = string.Empty;
                if (signMode == 0 && _fingerPrintHashWatermark)
                {
                    contentHash = ComputateContentHash();
                }
                if (signMode == 1 && _kydientuHashWaterMark)
                {
                    contentHash = ComputateContentHash();
                }

                var signer = new SignerNameDto("BỆNH NHÂN", _entity.MEPatient.MEPatientName, _entity.MEPatient.MEPatientNo);
                var listName = new List<SignerNameDto> { signer };

                var emr = _entity.MainObject as MEEmrsInfo;
                var relationEmrs = _emrCtrl.GetListBusinessObjects<MEEmrsInfo>(_emrCtrl.GetAllWithRelation(emr.MEEmrID));
                foreach (var item in relationEmrs)
                {
                    var relation = ADConfigValueUtility.GetConfigTextByGroupAndValue("EmrRelationFromName", item.MEEmrRelationFromName);
                    var relativeSigner = new SignerNameDto(relation.ToUpper(), item.MEPatientName, item.MEPatientNo);
                    listName.Add(relativeSigner);
                    if (item.MEEmrRelationToName.Equals(defaultSignerRela))
                        signer = relativeSigner;
                }
                // Defaut tên bệnh nhân
                var fullSignerName = _entity.MEPatient.MEPatientName;

                if (signMode == 1)
                {
                    // Relations
                    var guiRelation = new guiSelectPatientRelation(listName, signer, alternativeSign)
                    {
                        Module = this,
                        StartPosition = FormStartPosition.CenterParent
                    };

                    if (guiRelation.ShowDialog() == DialogResult.OK)
                    {
                        //signer = guiRelation.SelectedSigner;
                        fullSignerName = guiRelation.SelectedName;
                    }

                    // size signature
                    var sigpadWidth = Convert.ToInt32(BOSApp.GetSystemConfigValue(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.SIGN_PAD_WIDTH));
                    var sigpadHeight = Convert.ToInt32(BOSApp.GetSystemConfigValue(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.SIGN_PAD_HEIGHT));
                    var termsSignPad = BOSApp.GetSystemConfigValue(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.SIGN_PAD_TERMS);
                    using (var signPadForm = new MainWindow(stPad, contentHash, listName, signer, alternativeSign, sigpadWidth, sigpadHeight, termsSignPad))
                    {
                        if (signPadForm.ShowDialog() != DialogResult.OK) return;

                        byte[] byteImage; 

                        using (var ms = new MemoryStream())
                        {
                            signPadForm.Signature.Save(ms, signPadForm.Signature.RawFormat);
                            byteImage = ms.ToArray();
                        }
                        byteImage = ResizeImage(byteImage, 100, 140);

                        Image signImage;
                        using (MemoryStream ms = new MemoryStream(byteImage))
                        {
                            ms.Position = 0;
                            signImage = Image.FromStream(ms);
                        }

                        //var img = signPadForm.Signature;
                        var img = signImage;
                        if (img == null) return;

                        var signRanges = new List<DocumentRange>();
                        if (signRoles.Count > 0)
                        {
                            var fields = signRoles.SelectMany(r => r.Value).Select(f => f.Field).ToList();
                            this._emrDocumentHelper.InsertSignature(fields, img,
                                                                        fullSignerName, fullSignerName,
                                                                        string.Empty, _entity.MEPatient.MEPatientNo,
                                                                        AppMemCache.GetTemplateParams(document.FK_METemplateID),
                                                                        false);
                            signRanges = fields.Select(f => f.Range).ToList();
                        }
                        else
                        {
                            var signRange = this.InsertImageToDocument(img, 300, 0, false)?.Range;
                            //them dau nhay de co the huy ky
                            var pos = doc.InsertText(signRange.End, " ").End;
                            signRange = doc.CreateRange(signRange.Start, signRange.Length + 1);
                            if (!string.IsNullOrEmpty(fullSignerName))
                            {
                                pos = doc.InsertText(pos, "\n").End;
                                var fullNameRange = doc.InsertText(pos, signPadForm.SelectedName);
                                signRange = doc.CreateRange(signRange.Start, fullNameRange.End.ToInt() - signRange.Start.ToInt());
                            }
                            signRanges.Add(signRange);
                        }
                        if (signRanges.Count > 0)
                        {
                            var rangePermissions = doc.BeginUpdateRangePermissions();
                            doc.BeginUpdate();
                            try
                            {
                                for (int i = 0; i < signRanges.Count; i++)
                                {
                                    if (i == (signRanges.Count - 1))
                                    {
                                        Comment comment = doc.Comments.Create(signRanges[i], _entity.MEPatient.MEPatientID.ToString(), DateTime.Now);
                                        comment.Name = "Ký điện tử";
                                        SubDocument commentDocument = comment.BeginUpdate();
                                        commentDocument.InsertText(commentDocument.CreatePosition(0), $"Được ký bởi: {fullSignerName} lúc {DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")}");
                                        comment.EndUpdate(commentDocument);
                                    }
                                    rangePermissions.AddRange(_emrDocumentHelper.CreateRangePermissions(signRanges[i], "Patient", _entity.MEPatient.MEPatientID.ToString()));
                                }
                            }
                            catch (Exception)
                            {
                                throw;
                            }
                            finally
                            {
                                doc.EndUpdateRangePermissions(rangePermissions);
                                doc.EndUpdate();
                            }
                        }
                    }
                }
                else if (signMode == 0 && !_useZKTeco)
                {
                    using (var cap = new Emr.FingerPrint.Capture(_fingerPrint, contentHash, listName, signer, alternativeSign))
                    {
                        if (cap.ShowDialog() != DialogResult.OK) return;

                        byte[] byteImage;

                        using (var ms = new MemoryStream())
                        {
                            cap.GetImage().Save(ms, cap.GetImage().RawFormat);
                            byteImage = ms.ToArray();
                        }
                        byteImage = ResizeImage(byteImage, 100, 140);

                        Image signImage;
                        using (MemoryStream ms = new MemoryStream(byteImage))
                        {
                            ms.Position = 0;
                            signImage = Image.FromStream(ms);
                        }

                        var img = signImage;
                        //var img = cap.GetImage();
                        if (img == null) return;
                        signer = cap.SelectedSigner;
                        fullSignerName = cap.SelectedName;
                        var signRanges = new List<DocumentRange>();
                        if (signRoles.Count > 0)
                        {
                            var fields = signRoles.SelectMany(r => r.Value).Select(f => f.Field).ToList();
                            this._emrDocumentHelper.InsertSignature(fields, img,
                                                                     cap.SelectedName, cap.SelectedName,
                                                                     string.Empty, _entity.MEPatient.MEPatientNo,
                                                                     AppMemCache.GetTemplateParams(document.FK_METemplateID),
                                                                     false);
                            signRanges = fields.Select(f => f.Range).ToList();
                        }
                        else
                        {
                            var signRange = this.InsertImageToDocument(img, 300, 0, false)?.Range;
                            //them dau nhay de co the huy ky
                            var pos = doc.InsertText(signRange.End, " ").End;
                            signRange = doc.CreateRange(signRange.Start, signRange.Length + 1);
                            if (!string.IsNullOrEmpty(cap.SelectedName))
                            {
                                pos = doc.InsertText(pos, "\n").End;
                                var fullNameRange = doc.InsertText(pos, cap.SelectedName);
                                signRange = doc.CreateRange(signRange.Start, fullNameRange.End.ToInt() - signRange.Start.ToInt());
                            }
                            signRanges.Add(signRange);
                        }
                        if (signRanges.Count > 0)
                        {
                            var rangePermissions = doc.BeginUpdateRangePermissions();
                            doc.BeginUpdate();
                            try
                            {
                                for (int i = 0; i < signRanges.Count; i++)
                                {
                                    if (i == (signRanges.Count - 1))
                                    {
                                        Comment comment = doc.Comments.Create(signRanges[i], _entity.MEPatient.MEPatientID.ToString(), DateTime.Now);
                                        comment.Name = "Ký vân tay";
                                        SubDocument commentDocument = comment.BeginUpdate();
                                        commentDocument.InsertText(commentDocument.CreatePosition(0), $"Được ký bởi: {fullSignerName} lúc {DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")}");
                                        comment.EndUpdate(commentDocument);
                                    }
                                    rangePermissions.AddRange(_emrDocumentHelper.CreateRangePermissions(signRanges[i], "Patient", _entity.MEPatient.MEPatientID.ToString()));
                                }
                            }
                            catch (Exception)
                            {
                                throw;
                            }
                            finally
                            {
                                doc.EndUpdateRangePermissions(rangePermissions);
                                doc.EndUpdate();
                            }
                        }
                    }
                }
                else
                {
                    using (var cap = new CaptureZKTeco(_fingerPrintZK, contentHash, listName, signer, alternativeSign))
                    {
                        if (cap.ShowDialog() != DialogResult.OK) return;

                        byte[] byteImage;

                        using (var ms = new MemoryStream())
                        {
                            ImageCodecInfo codec = ImageCodecInfo.GetImageEncoders().FirstOrDefault(c => c.MimeType.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase));
                            EncoderParameters encoderParams = new EncoderParameters(1);
                            encoderParams.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 90L);

                            cap.GetImage().Save(ms, codec, encoderParams);
                            byteImage = ms.ToArray();
                        }
                        byteImage = ResizeImage(byteImage, 100, 140);

                        Image signImage;
                        using (MemoryStream ms = new MemoryStream(byteImage))
                        {
                            ms.Position = 0;
                            signImage = Image.FromStream(ms);
                        }

                        //var img = signPadForm.Signature;
                        var img = signImage;
                        //var img = cap.GetImage();
                        if (img == null) return;
                        signer = cap.SelectedSigner;
                        fullSignerName = cap.SelectedName;
                        var signRanges = new List<DocumentRange>();
                        if (signRoles.Count > 0)
                        {
                            var fields = signRoles.SelectMany(r => r.Value).Select(f => f.Field).ToList();
                            this._emrDocumentHelper.InsertSignature(fields, img,
                                                                     cap.SelectedName, cap.SelectedName,
                                                                     string.Empty, _entity.MEPatient.MEPatientNo,
                                                                     AppMemCache.GetTemplateParams(document.FK_METemplateID),
                                                                     false);
                            signRanges = fields.Select(f => f.Range).ToList();
                        }
                        else
                        {
                            var signRange = this.InsertImageToDocument(img, 300, 0, false)?.Range;
                            //them dau nhay de co the huy ky
                            var pos = doc.InsertText(signRange.End, " ").End;
                            signRange = doc.CreateRange(signRange.Start, signRange.Length + 1);
                            if (!string.IsNullOrEmpty(cap.SelectedName))
                            {
                                pos = doc.InsertText(pos, "\n").End;
                                var fullNameRange = doc.InsertText(pos, cap.SelectedName);
                                signRange = doc.CreateRange(signRange.Start, fullNameRange.End.ToInt() - signRange.Start.ToInt());
                            }
                            signRanges.Add(signRange);
                        }
                        if (signRanges.Count > 0)
                        {
                            var rangePermissions = doc.BeginUpdateRangePermissions();
                            doc.BeginUpdate();
                            try
                            {
                                for (int i = 0; i < signRanges.Count; i++)
                                {
                                    if (i == (signRanges.Count - 1))
                                    {
                                        Comment comment = doc.Comments.Create(signRanges[i], _entity.MEPatient.MEPatientID.ToString(), DateTime.Now);
                                        comment.Name = "Ký vân tay";
                                        SubDocument commentDocument = comment.BeginUpdate();
                                        commentDocument.InsertText(commentDocument.CreatePosition(0), $"Được ký bởi: {fullSignerName} lúc {DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")}");
                                        comment.EndUpdate(commentDocument);
                                    }
                                    rangePermissions.AddRange(_emrDocumentHelper.CreateRangePermissions(signRanges[i], "Patient", _entity.MEPatient.MEPatientID.ToString()));
                                }
                            }
                            catch (Exception)
                            {
                                throw;
                            }
                            finally
                            {
                                doc.EndUpdateRangePermissions(rangePermissions);
                                doc.EndUpdate();
                            }
                        }
                    }
                }


                BOSProgressBar.Start("Đang xử lý dữ liệu");
                var docx = "." + EmrDocumentFileExtention.docx.ToString();
                var cloneStream = new MemoryStream();
                this._richEditCtrl.SaveDocument(cloneStream, DocumentFormat.OpenXml);
                var extractTask = Task.Factory.StartNew<string>(() =>
                {
                    try
                    {
                        using (var cloneRich = new RichEditDocumentServer())
                        {
                            cloneStream.Position = 0;
                            cloneRich.LoadDocument(cloneStream, DocumentFormat.OpenXml);
                            string toFileName = document.MEEmrDocumentFile + "_patient_sign" + DateTime.Now.ToBinary();
                            string toFileFullname = toFileName + docx;
                            string toFilePath = Path.Combine(_documentPath, "Emr", "Partials", document.FK_MEEmrID.ToString(), toFileFullname);
                            CreateSignLocalDir(document.FK_MEEmrID);
                            this.RemoveAllTagForPrint(cloneRich, _entity.METemplate);
                            this._emrDocumentHelper.RemoveAllParamMarkup(cloneRich.Document);
                            cloneRich.SaveDocument(toFilePath, DocumentFormat.OpenXml);
                            _ftpFileMng.UploadFile($"/Emr/Partials/{document.FK_MEEmrID}/", toFileFullname, toFilePath);
                            return toFileName;
                        }
                    }
                    catch (Exception ex)
                    {
                        PrintMgsLog("XU-LY-PHAN-KY-TEN-DIGITAL-PRINT", "ERROR: " + ex.ToString());
                        return string.Empty;
                    }
                    finally
                    {
                        cloneStream.Close();
                        cloneStream.Dispose();
                    }
                });

                string fileName = Path.Combine(_documentPath, "Emr", document.FK_MEEmrID.ToString(), document.MEEmrDocumentFile + docx);
                BOSProgressBar.SetText("Đang lưu dữ liệu");
                _richEditCtrl.SaveDocument(fileName, DocumentFormat.OpenXml);
                if (!EncryptFileDocument())
                {
                    BOSProgressBar.Close();
                    return;
                }
                if (!SaveEmrDocumentInfoAndUploadFile())
                {
                    BOSProgressBar.Close();
                    return;
                }
                Task.WaitAll(extractTask);
                var sign = new MEEmrDocumentSignsInfo()
                {
                    FK_MEEmrDocumentID = document.MEEmrDocumentID,
                    FK_HRDepartmentID = BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID,
                    FK_HREmployeeID = BOSApp.CurrentEmployeesInfo.HREmployeeID,
                    MEEmrDocumentSignTime = DateTime.Now,
                    MEEmrDocumentSignFile = extractTask.Result,
                    MEEmrDocumentSignFileExt = EmrDocumentFileExtention.docx.ToString(),
                    MEEmrDocumentSignType = EmrDocumentSignType.FingerPrintSigned.ToString(),
                    MEEmrDocumentSignRemark = "Xác nhận ký điện tử bởi: " + fullSignerName,
                    MEEmrDocumentSignHash = contentHash,
                    MEEmrDocumentSignUser = signer != null ? signer.SignerNo : _entity.MEPatient.MEEmrNo,
                    MEEmrDocumentSignBlockAddr = string.Empty,
                    AACreatedUser = BOSApp.CurrentUser
                };
                this._emrDocumentSignCtrl.CreateObject(sign);
                CloneSignedRangeForTracking();
                ((DevExpress.XtraRichEdit.Model.DocumentModel)this._richEditCtrl.Model).History.Clear();
                _richEditCtrl.Modified = false;
            }
            catch (Exception ex)
            {
                _msgNotification.Text = "Có lỗi khi ký. Xin thử lại.";
                PrintMgsLog("KY-TEN-LOI-DIGITAL-PRINT: ", ex.ToString());
            }
            finally
            {
                BOSProgressBar.Close();
            }
        }
        private string ComputateContentHash()
        {
            var exportCf = new DevExpress.XtraRichEdit.Export.PlainTextDocumentExporterOptions() { ExportHiddenText = false };
            var textContent = _richEditCtrl.Document.GetText(_richEditCtrl.Document.Range, exportCf);
            var hasher = new HashProvider("SHA1");

            using (MemoryStream stream = new MemoryStream())
            using (BinaryWriter binWriter = new BinaryWriter(stream))
            {
                var bytes = Encoding.UTF8.GetBytes(textContent);
                stream.Position = 0;
                binWriter.Write(bytes);
                foreach (var image in _richEditCtrl.Document.Images)
                {
                    bytes = image.Image.RootImage.GetImageBytesSafe(image.Image.RawFormat);
                    binWriter.Write(bytes);
                }
                foreach (var image in _richEditCtrl.Document.Shapes)
                {
                    if (image.Picture == null) continue;
                    bytes = image.Picture.RootImage.GetImageBytesSafe(image.Picture.RawFormat);
                    binWriter.Write(bytes);
                }
                stream.Position = 0;
                return hasher.ComputeHash(stream);
            }
        }
        #endregion

        #region Tool edit emrNo, patientNo
        public void EmrRenameEmrNo()
        {
            // Check save current
            if (this._richEditCtrl.Modified)
            {
                MessageBox.Show($"Vui lòng lưu lại thông tin trước khi thao tác.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            var emr = this._entity.MainObject as MEEmrsInfo;
            if (string.IsNullOrEmpty(emr.MEEmrNo))
            {
                MessageBox.Show($"Vui lòng chọn 1 bệnh án.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            var gui = new guiUpdateEmrNo(emr.MEEmrNo)
            {
                Module = this
            };
            gui.InitializeControls(gui.Controls);
            gui.StartPosition = FormStartPosition.CenterParent;
            if (gui.ShowDialog() == DialogResult.OK)
            {
                var currentEmrNo = emr.MEEmrNo;
                var newEmrNo = gui._emrNoNew;
                // Update emrNo
                var emrDb = _emrCtrl.GetObjectByID(emr.MEEmrID) as MEEmrsInfo;
                var changeMsg = $"cũ {emrDb.MEEmrNo} - mới {newEmrNo}";
                emrDb.MEEmrNo = newEmrNo;
                emrDb.AAUpdatedUser = BOSApp.CurrentUser;
                _emrCtrl.UpdateObject(emrDb);

                _entity.MainObject = emrDb;

                HistoryEmr(emr, cstObjectHistoryActionChange, $"thay đổi thông tin bệnh án - mã bệnh án {changeMsg}.");

                MessageBox.Show($"Sửa mã bệnh án thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                gui.Close();

                RefreshByEmr(emrDb);
            }
        }

        public void EmrRenamePatientNo()
        {
            // Check save current
            if (this._richEditCtrl.Modified)
            {
                MessageBox.Show($"Vui lòng lưu lại thông tin trước khi thao tác.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            var patientE = _entity.MEPatient;
            if (string.IsNullOrEmpty(patientE.MEPatientNo))
            {
                MessageBox.Show($"Vui lòng chọn 1 bệnh nhân.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            var gui = new guiUpdatePatientNo(patientE.MEPatientNo)
            {
                Module = this
            };
            gui.InitializeControls(gui.Controls);
            gui.StartPosition = FormStartPosition.CenterParent;
            if (gui.ShowDialog() == DialogResult.OK)
            {
                var currentNo = patientE.MEPatientNo;
                var newNo = gui._patientNoNew;
                var remark = gui._patientRemarkNew;
                var userUpdate = BOSApp.CurrentUser;
                // Update emrNo
                // Patient
                remark = $"Đổi mã, lý do: {remark}";
                var patientDb = _patientCtrl.GetObjectByID(patientE.MEPatientID) as MEPatientsInfo;
                patientDb.MEPatientNo = newNo;
                patientDb.MEPatientDesc = remark;
                patientDb.AAUpdatedUser = userUpdate;
                _patientCtrl.UpdateObject(patientDb);
                // update current entity
                _entity.ModuleObjects[TableName.MEPatientsTableName] = patientDb;
                // update current emr
                var emr = this._entity.MainObject as MEEmrsInfo;
                emr.MEPatientNo = patientDb.MEPatientNo;

                HistoryEmr(emr, cstObjectHistoryActionChange, $"Đổi mã bệnh nhân {currentNo} thành {newNo}. Lý do {remark}");

                MessageBox.Show($"Sửa mã bệnh nhân thành công.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                gui.Close();

                RefreshByEmr(emr);
            }
        }

        private void RefreshByEmr(MEEmrsInfo emr)
        {
            RefreshCurrentEmrSearchResultsControl(emr);
            Invalidate(emr.MEEmrID);
        }

        #endregion

        #region Histories
        private void HistoryEmr(MEEmrsInfo emr, string action, string remark)
        {
            var objGeObjectHistoryInfo = new GEObjectHistoryInfo
            {
                ADUserID = BOSApp.CurrentUsersInfo.ADUserID,
                ADUserName = BOSApp.CurrentUser,
                GEObjectHistoryObjectName = TableName.MEEmrsTableName,
                GEObjectHistoryObjectID = emr.MEEmrID,
                GEObjectHistoryObjectNumber = emr.MEEmrNo,
                GEObjectHistoryAction = action,
                GEObjectHistoryDate = DateTime.Now,
                GEObjectHistoryRemark = $"{BOSApp.CurrentEmployeesInfo.HREmployeeName} - {remark}."
            };

            _geObjHistoryCtrl.CreateObject(objGeObjectHistoryInfo);
        }

        private string getDocumentNo(MEEmrDocumentsInfo emrDoc)
        {
            if (emrDoc == null)
                return string.Empty;

            var docNo = emrDoc.MEEmrDocumentNo;
            if (string.IsNullOrEmpty(docNo))
            {
                var template = _templateCtrl.GetObjectByID(emrDoc.FK_METemplateID) as METemplatesInfo;
                docNo = template.METemplateNo;
            }
            return docNo;
        }

        internal void PrintHistory()
        {
            HistoryEmr(_entity.MainObject as MEEmrsInfo, "Change", "In bệnh án.");
        }
        #endregion

        #region TDT
        private int GetTableEmrDocument(string paramNo, int headerRow, int blankRow)
        {
            try
            {
                // Add 2 rows: 1 row blank
                var paramCtrl = new MEParamsController();
                var paramE = paramCtrl.GetObjectByNo(paramNo) as MEParamsInfo;
                var field = _emrDocumentHelper.GetFirstEmrFieldByCode(paramE.MEParamNo);
                var updateField = _emrDocumentHelper.GetFirstFieldByPath(paramE.MEParamNo, field.Gid);
                var table = _emrDocumentHelper.GetTableFromDocument(_richEditCtrl, updateField);
                var rows = table.Rows.Count();
                var rowsNoHead = rows - headerRow;
                var rowsTrue = rowsNoHead / (1 + blankRow);
                return rowsTrue;
            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        private bool AllowCreateTDT(MEEmrsInfo emr, int templateId)
        {
            var maxRow = BOSApp.GetSystemConfigValueInt(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.EMR_DOCUMENT_TDT_DIENBIENVAYLENH_ROW, 0);
            if (maxRow > 0)
            {
                var templateE = this._templateCtrl.GetObjectByID(templateId) as METemplatesInfo;
                if (templateE != null && templateE.METemplateNo == "TDT")
                {
                    var validateCtrl = new MEEmrDocumentValidatesController();
                    var validateE = validateCtrl.GetObjectInfoByEmrAndTemplateAndModeAndValue(emr.MEEmrID, templateId, DocumentValidateMode.Create, "FALSE");
                    if (validateE != null)
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        private void BackgroundEmrDocumentValidates(MEEmrsInfo emr, int templateId, int documentId)
        {
            if (_sysHelper.AllowThread())
            {
                var threadEmrDocumentValidates = new System.Threading.Thread(() => EmrDocumentValidates(emr, templateId, documentId));
                threadEmrDocumentValidates.Start();
            }
            else
            {
                EmrDocumentValidates(emr, templateId, documentId);
            }
        }

        private void BackgroundEmrDocumentValidatesDelete(MEEmrsInfo emr, int templateId, int documentId)
        {
            if (_sysHelper.AllowThread())
            {
                var threadEmrDocumentValidates = new System.Threading.Thread(() => EmrDocumentValidatesDelete(emr, templateId, documentId));
                threadEmrDocumentValidates.Start();
            }
            else
            {
                EmrDocumentValidatesDelete(emr, templateId, documentId);
            }
        }

        private void BackgroundEmrDocumentValidatesActive(MEEmrsInfo emr, int templateId, int documentId)
        {
            if (_sysHelper.AllowThread())
            {
                var threadEmrDocumentValidates = new System.Threading.Thread(() => EmrDocumentValidatesActive(emr, templateId, documentId));
                threadEmrDocumentValidates.Start();
            }
            else
            {
                EmrDocumentValidatesActive(emr, templateId, documentId);
            }
        }

        private void EmrDocumentValidates(MEEmrsInfo emr, int templateId, int documentId)
        {
            try
            {
                var maxRow = BOSApp.GetSystemConfigValueInt(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.EMR_DOCUMENT_TDT_DIENBIENVAYLENH_ROW, 0);
                if (maxRow > 0)
                {
                    var templateE = this._templateCtrl.GetObjectByID(templateId) as METemplatesInfo;
                    if (templateE != null && templateE.METemplateNo == "TDT")
                    {
                        var initRow = BOSApp.GetSystemConfigValueInt(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.EMR_DOCUMENT_TDT_DIENBIENVAYLENH_INIT_ROW, 1);
                        var statusAllow = "FALSE";
                        if (initRow >= maxRow)
                            statusAllow = "TRUE";

                        UpdateEmrDocumentValidates(emr, templateId, documentId, statusAllow, $"Dòng trên tờ: {initRow} - Dòng quy định tối đa: {maxRow}");
                    }
                }
            }
            catch (Exception ex)
            {
                _sysHelper.LogTxt("error", ex.ToString());
            }
        }

        private void EmrDocumentValidatesDelete(MEEmrsInfo emr, int templateId, int documentId)
        {
            try
            {
                var maxRow = BOSApp.GetSystemConfigValueInt(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.EMR_DOCUMENT_TDT_DIENBIENVAYLENH_ROW, 0);
                if (maxRow > 0)
                {
                    var templateE = this._templateCtrl.GetObjectByID(templateId) as METemplatesInfo;
                    if (templateE != null && templateE.METemplateNo == "TDT")
                    {
                        var validateCtrl = new MEEmrDocumentValidatesController();
                        var validateObj = validateCtrl.GetObjectInfoByDocumentAndMode(documentId, DocumentValidateMode.Create);
                        if (validateObj != null && validateObj.MEEmrDocumentValidateID > 0)
                        {
                            validateCtrl.DeleteObject(validateObj.MEEmrDocumentValidateID);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _sysHelper.LogTxt("error", ex.ToString());
            }
        }

        private void EmrDocumentValidatesActive(MEEmrsInfo emr, int templateId, int documentId)
        {
            try
            {
                var maxRow = BOSApp.GetSystemConfigValueInt(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.EMR_DOCUMENT_TDT_DIENBIENVAYLENH_ROW, 0);
                if (maxRow > 0)
                {
                    var templateE = this._templateCtrl.GetObjectByID(templateId) as METemplatesInfo;
                    if (templateE != null && templateE.METemplateNo == "TDT")
                    {
                        var validateCtrl = new MEEmrDocumentValidatesController();
                        var validateObj = validateCtrl.GetObjectInfoByDocumentAndModeAndInActive(documentId, DocumentValidateMode.Create);
                        if (validateObj != null && validateObj.MEEmrDocumentValidateID > 0)
                        {
                            validateObj.AAStatus = "Alive";
                            validateObj.AAUpdatedUser = BOSApp.CurrentUser;
                            validateObj.AAUpdatedDate = DateTime.Now;
                            validateCtrl.UpdateObject(validateObj);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _sysHelper.LogTxt("error", ex.ToString());
            }
        }

        private void UpdateEmrDocumentValidates(MEEmrsInfo emr, int templateId, int documentId, string value, string remark)
        {
            // avoid documentId = 0, some time documentId = 0
            if (documentId == 0)
            {
                var currentDoc = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
                var docDb = _emrDocumentCtrl.GetObjectByEmrAndFile(emr.MEEmrID, currentDoc.MEEmrDocumentFile, currentDoc.MEEmrDocumentFileExt);
                if (docDb != null && docDb.MEEmrDocumentID > 0)
                {
                    documentId = docDb.MEEmrDocumentID;
                }
                else
                {
                    return;
                }
            }

            var validateCtrl = new MEEmrDocumentValidatesController();
            var validateE = validateCtrl.GetObjectInfoByEmrAndTemplateAndDocumentAndMode(emr.MEEmrID, templateId, documentId, DocumentValidateMode.Create);
            if (validateE == null)
            {
                validateCtrl.CreateObject(new MEEmrDocumentValidatesInfo()
                {
                    AACreatedUser = BOSApp.CurrentUser,
                    FK_MEEmrID = emr.MEEmrID,
                    FK_METemplateID = templateId,
                    FK_MEEmrDocumentID = documentId,
                    MEEmrDocumentValidateMode = DocumentValidateMode.Create,
                    MEEmrDocumentValidateValue = value,
                    MEEmrDocumentValidateRemark = remark
                });
            }
            else
            {
                validateE.AAUpdatedUser = BOSApp.CurrentUser;
                validateE.MEEmrDocumentValidateValue = value;
                validateE.MEEmrDocumentValidateRemark = remark;
                validateCtrl.UpdateObject(validateE);
            }
        }

        private void TDTUpdateEmrDocumentValidatesWhenSave(MEEmrDocumentsInfo document)
        {
            var maxRow = BOSApp.GetSystemConfigValueInt(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.EMR_DOCUMENT_TDT_DIENBIENVAYLENH_ROW, 0);
            if (maxRow > 0)
            {
                var emr = _entity.MainObject as MEEmrsInfo;
                var templateE = _templateCtrl.GetObjectByID(document.FK_METemplateID) as METemplatesInfo;
                if (templateE.METemplateNo == "TDT")
                {
                    var rowTDT = GetTableEmrDocument("todieutri_dienbienvaylenh", 1, 1);
                    if (rowTDT < maxRow)
                    {
                        UpdateEmrDocumentValidates(emr, document.FK_METemplateID, document.MEEmrDocumentID, "FALSE", $"Dòng trên tờ: {rowTDT} - Dòng quy định tối đa: {maxRow}");
                    }
                }
            }
        }

        //Tam thoi chua su dung
        public object CallBackEmrAction(string navigateUri, DocumentRange actionRange, MEEmrDocumentsInfo document, string groupIn = "", object data = null,
            Func<object, string, string, string> beforeBinding = null, object preData = null, int cacheTimeOut = 0)
        {
            BackgroudLoadDocumentRelativeThings(document);
            OpenEmrDocument(document.FK_MEEmrID, document.MEEmrDocumentFile, document.MEEmrDocumentStatus == EmrDocumentStatus.Closed.ToString(), true, true);
            var errCount = 0;
            try
            {
                _msgNotification.Text = $"Đang cập nhật...";
                //var emr = GetCurrentMainObject();
                MEEmrsInfo emr = (MEEmrsInfo)_emrCtrl.GetObjectByID(document.FK_MEEmrID);
                Cursor.Current = Cursors.WaitCursor;
                //if (IsEmrReadOnly())
                //{
                //    _msgNotification.Text = "Bệnh án này đã đóng. Mọi thay đổi sẽ không thể lưu được.";
                //    return null;
                //};
                var uris = navigateUri.Split(EmrParam.TagCodeSeparator);
                if (uris.Length == 0) return null;

                if (actionRange == null)
                    actionRange = this._emrDocumentHelper.GetActionRange(navigateUri, groupIn);

                var action = _actionsController.GetObjectByNo(uris[0]) as MEEmrActionsInfo;
                if (action == null)
                {
                    _msgNotification.Text = "Không tìm thấy thẻ chức năng tương ứng. Kiểm tra lại cấu hình";
                    return null;
                };

                var notify = $"Đang thực thi chức năng [{action.MEEmrActionName}]";
                _msgNotification.Text = notify;

                //case RemoteCase
                if (action.MEEmrActionType == EmrActionTypes.RemoteCase.ToString())
                {
                    InsertRemoteCaseData(action, actionRange, uris.Length > 1 ? uris[1] : null);
                    _msgNotification.Text = "Dữ liệu đã được cập nhật";
                    return null;
                }
                //get action gid
                var group = groupIn;
                if (string.IsNullOrEmpty(group))
                {
                    group = uris.Where(t => t.Contains($"{EmrParam.GuidTag}=")).FirstOrDefault();
                    group = string.IsNullOrEmpty(group) ? string.Empty : group.Replace($"{EmrParam.GuidTag}=", "");
                }

                if (action.MEEmrActionType == EmrActionTypes.Hard.ToString())
                {
                    this.PerformHardAction(action, actionRange, group);
                    _msgNotification.Text = _msgNotification.Text.Replace(notify, string.Empty);
                    return null;
                }
                else if (action.MEEmrActionType == EmrActionTypes.Composition.ToString())
                {
                    this.CallCompositionAction(action, actionRange, group, cacheTimeOut);
                    return null;
                }
                else if (action.MEEmrActionType == EmrActionTypes.AutoValue.ToString())
                {
                    this.BindingDataToEmrDocument(GetHardParamList(emr), group, string.Empty);
                    return null;
                }

                // var key = this.GetDataKeyByGroup(group);
                var tid = uris.Where(t => t.Contains($"{EmrParam.TransactionIdTag}=")).FirstOrDefault();
                tid = tid == null ? string.Empty : tid.Replace($"{EmrParam.TransactionIdTag}=", "");

                var allParams = AppMemCache.GetActionParamsFromDict(action.MEEmrActionID);
                var updateParams = allParams.Where(o => o.MEEmrActionParamRequest == false && o.FK_MEParamID > 0).ToList();

                //sub template
                if (action.MEEmrActionType == EmrActionTypes.SubTemplate.ToString())
                {
                    DocumentPosition pos = actionRange.Start;
                    if (action.MEEmrActionNewGuid) group = Guid.NewGuid().ToString();
                    InsertSubTemplate(action, actionRange, updateParams, group);
                    _msgNotification.Text = "Dữ liệu đã được cập nhật";
                    return null;
                }
                if (action.MEEmrActionType == EmrActionTypes.DinamapProV100.ToString()
                    && action.MEEmrActionUri.ToUpper() == "PRINT")
                {
                    return PrintV100Snap();
                }

                //var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;

                var emrNo = _entity.GetRelativeEmrNo(emr.MEEmrNo);
                var paramList = new Dictionary<string, object>
                {
                    // params cung, luc nao cung truyen xuong
                    { "patientNo", this._entity.MEPatient.MEPatientNo },
                    { "emrNo", emrNo },
                    { "documentNo", document.MEEmrDocumentFile },
                    { "documentDate", document.MEEmrDocumentCreatedDate },
                    { EmrParam.TransactionIdTag, tid },
                    { EmrParam.GuidTag, group },
                    { "caseNo", this.GetCaseNo() }
                };

                //hard action params
                var hardParams = _requestParamPool.GetActionParam(action.MEEmrActionSendParams);
                //doc action params
                foreach (var item in hardParams)
                {
                    if (!paramList.ContainsKey(item.Key))
                        paramList.Add(item.Key, item.Value);
                    else
                    {
                        //Cho phep lay gia tri tu to benh an cho bat ky param nao
                        if (action.MEEmrActionAllowOverrideReqParamValue)
                            paramList[item.Key] = item.Value;
                    }
                }
                //doc action params
                var requestParams = this._emrDocumentHelper.GetActionParamFromDoc(allParams.Where(o => o.MEEmrActionParamRequest).ToList(), action, actionRange, group);
                requestParams = _emrActionHelper.MapToRequestParams(action.MEEmrActionSendParams, requestParams);

                if (action.MEEmrActionType == EmrActionTypes.OpenWebBrowser.ToString())
                {
                    this.OpenWebBrowser(requestParams, action);
                    return null;
                }

                foreach (var item in requestParams)
                {
                    if (!paramList.ContainsKey(item.Key))
                        paramList.Add(item.Key, item.Value);
                    else
                    {
                        //Cho phep lay gia tri tu to benh an cho bat ky param nao
                        if (action.MEEmrActionAllowOverrideReqParamValue)
                            paramList[item.Key] = item.Value;
                    }
                }

                if (updateParams == null || updateParams.Count == 0)
                {
                    _msgNotification.Text = ("Chức năng này không định nghĩa cập nhật thông tin nào. Kiểm tra lại cấu hình chức năng này.");
                    return null;
                }

                if (action.MEEmrActionType == EmrActionTypes.Sql.ToString())
                {
                    paramList["documentDate"] = document.MEEmrDocumentCreatedDate.ToString("dd/MM/yyyy HH:mm:ss");
                    data = GetSqlData(cacheTimeOut, action, paramList);
                    if (data != null)
                    {
                        if (beforeBinding != null)
                            errCount = BindingDataToEmrDocument(data, action, group, updateParams, paramList, (d, path) =>
                            {
                                return beforeBinding(d, group, path);
                            }, preData);
                        else
                            errCount = BindingDataToEmrDocument(data, action, group, updateParams, paramList, preData: preData);
                    }
                }
                else if (action.MEEmrActionType == EmrActionTypes.Api.ToString())
                {
                    data = GetApiData(cacheTimeOut, action, paramList);
                    if (data == null)
                    {
                        _msgNotification.Text = ("Không có dữ liệu trả về từ api. Liên hệ quản trị viên để biết thêm chi tiết");
                        return null;
                    }
                    PrintMgsLogJson($"DU-LIEU-API {action.MEEmrActionName}", data);
                    PrintMgsLog("BAT-DAU-CAP-NHAT-DU-LIEU", action.MEEmrActionUri);
                    if (beforeBinding != null)
                        errCount = BindingDataToEmrDocument(data, action, group, updateParams, paramList, (d, path) =>
                        {
                            return beforeBinding(d, group, path);
                        }, preData);
                    else
                        errCount = BindingDataToEmrDocument(data, action, group, updateParams, paramList, preData: preData);
                    PrintMgsLog("KET-THUC-CAP-NHAT-DU-LIEU", action.MEEmrActionUri);
                }
                else if (action.MEEmrActionType == EmrActionTypes.DataPlugin.ToString()
                    || action.MEEmrActionType == EmrActionTypes.CalcPlugin.ToString())
                {
                    PrintMgsLog("BAT-DAU-GOI-PLUGIN", action.MEEmrActionPlugin);
                    data = CallDllPlugin(action, group, updateParams, paramList, document);
                    PrintMgsLog("KET-THUC-GOI-PLUGIN", action.MEEmrActionPlugin);
                    if (data == null)
                    {
                        _msgNotification.Text = ("Không có dữ liệu trả về từ Trình cắm. Liên hệ quản trị viên để biết thêm chi tiết");
                        return null;
                    }
                    PrintMgsLogJson($"DU-LIEU-PLUGIN {action.MEEmrActionName}", data);
                    PrintMgsLog("BAT-DAU-CAP-NHAT-DU-LIEU", action.MEEmrActionPlugin);
                    errCount = BindingDataToEmrDocument(data, action, group, updateParams, paramList, preData: preData);
                    PrintMgsLog("KET-THUC-CAP-NHAT-DU-LIEU", action.MEEmrActionPlugin);
                }
                #region App
                else if (action.MEEmrActionType == EmrActionTypes.App.ToString())
                {
                    try
                    {
                        var icp = new IpcHelper();
                        var template = this._templateCtrl.GetObjectByID(document.FK_METemplateID) as METemplatesInfo;
                        string msg = "";
                        var message = new Dictionary<string, object>(paramList);
                        message["documentDate"] = document.MEEmrDocumentCreatedDate.ToString("dd/MM/yyyy HH:mm:ss");
                        message.Add("documentName", template.METemplateName);

                        message.Add("preData", preData);
                        _appPreData = preData;

                        message.Add("msg", action.MEEmrActionNo);
                        if (action.MEEmrActionDataType == EmrActionDataTypes.Xml.ToString())
                        {
                            msg = JsonConvert.SerializeObject(new { root = message }, Newtonsoft.Json.Formatting.Indented);
                            msg = JsonConvert.DeserializeXmlNode(msg).InnerXml;
                        }
                        else
                        {
                            msg = JsonConvert.SerializeObject(message, Newtonsoft.Json.Formatting.Indented);
                        }
                        icp.SendMessage(action.MEEmrActionNo, msg, action.MEEmrActionDataType);
                        PrintMgsLog("DA-GOI-YEU-CAU-DEN-HIS", msg);
                        _msgNotification.Text = "Đã gởi dữ liệu đến HIS. Đang chờ dữ liệu trả về...";
                        return null;
                    }
                    catch (Exception ex)
                    {
                        Trace.TraceError("IPC ERROR: {0}:{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), ex);
                        throw;
                    }
                }
                #endregion App
                else if (action.MEEmrActionType == EmrActionTypes.Replace.ToString())
                {
                    DocumentPosition pos = null;
                    //chuc nang nay yeu cau phat sinh ma nhom moi de co lap du lieu
                    if (action.MEEmrActionNewGuid) group = Guid.NewGuid().ToString();
                    foreach (var param in updateParams)
                    {
                        pos = this._emrActionHelper.ReplaceEmrParam(param, action, actionRange, group);
                    }
                    if (pos != null)
                        this._emrActionHelper.ReplaceEmrAction(pos, action, actionRange, group);
                }
                // gan gia tri cho the
                else if (action.MEEmrActionType == EmrActionTypes.Assign.ToString())
                {
                    foreach (var param in updateParams)
                    {
                        this._emrActionHelper.AssignEmrParamValue(param, action, actionRange, group);
                    }
                }
                // them dong du lieu moi cho table
                else if (action.MEEmrActionType == EmrActionTypes.AddRow.ToString())
                {
                    try
                    {
                        BOSProgressBar.Start("Đang xử lý dữ liệu");

                        var rowTDT = 0;
                        var maxRow = BOSApp.GetSystemConfigValueInt(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.EMR_DOCUMENT_TDT_DIENBIENVAYLENH_ROW, 0);
                        if (maxRow > 0)
                        {
                            var templateE = _templateCtrl.GetObjectByID(document.FK_METemplateID) as METemplatesInfo;
                            if (templateE.METemplateNo == "TDT")
                            {
                                rowTDT = GetTableEmrDocument("todieutri_dienbienvaylenh", 1, 1);
                                if (rowTDT >= maxRow)
                                {
                                    BOSProgressBar.Close();
                                    _msgNotification.Text = ("Số dòng Diễn biến, Y lệnh đã đủ. Không thể thêm dòng mới. Vui lòng tạo Tờ điều trị mới để thao tác.");
                                    UpdateEmrDocumentValidates(emr, document.FK_METemplateID, document.MEEmrDocumentID, "TRUE", $"Dòng trên tờ: {rowTDT} - Dòng quy định tối đa: {maxRow}");
                                    MessageBox.Show("Số dòng Diễn biến, Y lệnh đã đủ. Không thể thêm dòng mới. Vui lòng tạo Tờ điều trị mới để thao tác.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    return null;
                                }
                            }
                        }

                        var start = DateTime.Now;
                        var newGroup = this._emrActionHelper.AddTableRow(action, actionRange, group);
                        PrintMgsLog("AddTableRow: ", (DateTime.Now - start).TotalMilliseconds.ToString());

                        if (maxRow > 0)
                        {
                            rowTDT += 1;
                            if (rowTDT >= maxRow)
                            {
                                UpdateEmrDocumentValidates(emr, document.FK_METemplateID, document.MEEmrDocumentID, "TRUE", $"Dòng trên tờ: {rowTDT} - Dòng quy định tối đa: {maxRow}");
                            }
                            else
                            {
                                UpdateEmrDocumentValidates(emr, document.FK_METemplateID, document.MEEmrDocumentID, "FALSE", $"Dòng trên tờ: {rowTDT} - Dòng quy định tối đa: {maxRow}");
                            }
                        }

                        start = DateTime.Now;
                        this.BindingDataToEmrDocument(GetHardParamList(emr), newGroup, string.Empty);
                        PrintMgsLog("BindingDataToEmrDocument: ", (DateTime.Now - start).TotalMilliseconds.ToString());

                        if (beforeBinding != null)
                            data = beforeBinding(data as JToken, newGroup, string.Empty);
                    }
                    catch (Exception)
                    {
                        throw;
                    }
                    finally
                    {
                        BOSProgressBar.Close();
                    }

                }
                // them cot du lieu moi cho table
                else if (action.MEEmrActionType == EmrActionTypes.AddCol.ToString())
                {
                    this._emrActionHelper.AddTableCol(action, actionRange, group);
                }
                #region Mongo
                else if (action.MEEmrActionType == EmrActionTypes.Lookup.ToString())
                {
                    var template = this._templateCtrl.GetObjectByID(action.FK_MESourceTemplateID) as METemplatesInfo;
                    var filters = new Dictionary<string, object>();
                    if (action.MEEmrActionScope != EmrActionScopes.All.ToString())
                    {
                        if (action.MEEmrActionScope == EmrActionScopes.Patient.ToString())
                            filters.Add("MEEmr.FK_MEPatientID", new Tuple<MongoFilter, object>(MongoFilter.Eq, this._entity.MEPatient.MEPatientID));
                        if (action.MEEmrActionScope == EmrActionScopes.Emr.ToString())
                            filters.Add("FK_MEEmrID", new Tuple<MongoFilter, object>(MongoFilter.Eq, emr.MEEmrID));
                        if (action.MEEmrActionScope == EmrActionScopes.Department.ToString())
                            filters.Add("MEEmr.FK_HRDepartmentID", new Tuple<MongoFilter, object>(MongoFilter.Eq, BOSApp.CurrentEmployeesInfo.FK_HRDepartmentID));
                        //if (action.MEEmrActionScope == EmrActionScopes.Document.ToString() && !string.IsNullOrEmpty(document.MEEmrDocumentMongoID)) 
                        //    filters.Add("_id", new Tuple<MongoFilter, object>(MongoFilter.Eq, ObjectId.Parse(document.MEEmrDocumentMongoID)));
                        if (action.MEEmrActionScope == EmrActionScopes.Document.ToString())
                        {
                            if (string.IsNullOrEmpty(document.MEEmrDocumentMongoID))
                            {
                                filters.Add("_id", new Tuple<MongoFilter, object>(MongoFilter.Eq, ObjectId.GenerateNewId()));
                            }
                            else
                            {
                                filters.Add("_id", new Tuple<MongoFilter, object>(MongoFilter.Eq, ObjectId.Parse(document.MEEmrDocumentMongoID)));
                            }
                        }

                    }
                    //cac scrope khac ko can xet trong to hien tai
                    if (action.MEEmrActionScope != EmrActionScopes.Document.ToString() && !string.IsNullOrEmpty(document.MEEmrDocumentMongoID))
                        filters.Add("_id", new Tuple<MongoFilter, object>(MongoFilter.Ne, ObjectId.Parse(document.MEEmrDocumentMongoID)));

                    filters.Add("MEEmrDocumentStatus", new Tuple<MongoFilter, object>(MongoFilter.Nin, new string[] { EmrDocumentStatus.Hidden.ToString(), EmrDocumentStatus.Discarded.ToString() }));

                    //hard action params
                    foreach (var item in hardParams)
                    {
                        var key = "MEEmrDocumentContent." + item.Key;
                        if (!filters.ContainsKey(key))
                            filters.Add(key, new Tuple<MongoFilter, object>(MongoFilter.Eq, item.Value));
                    }
                    //doc action params
                    foreach (var item in requestParams)
                    {
                        var key = "MEEmrDocumentContent." + item.Key;
                        if (!filters.ContainsKey(key))
                            filters.Add(key, new Tuple<MongoFilter, object>(MongoFilter.Eq, item.Value));
                    }
                    var fields = new Dictionary<string, string>();
                    foreach (var item in allParams.Where(o => o.FK_MEParamID == 0))
                    {
                        if (!string.IsNullOrEmpty(item.MEEmrActionParamSourcePath))
                        {
                            var sourcePath = item.MEEmrActionParamSourcePath.Replace("[*]", string.Empty);
                            var path = "$MEEmrDocumentContent." + sourcePath;
                            if (!fields.ContainsKey(sourcePath))
                                fields.Add(sourcePath, path);
                        }
                    }
                    foreach (var item in updateParams)
                    {
                        var dest = AppMemCache.GetParamFromDictKeyID(item.FK_MEParamID);
                        var sourcePath = item.MEEmrActionParamSourcePath.Replace("[*]", string.Empty);
                        if (dest != null && !string.IsNullOrEmpty(sourcePath))
                        {
                            var path = "$MEEmrDocumentContent." + sourcePath;
                            if (!fields.ContainsKey(dest.MEParamNo))
                                fields.Add(dest.MEParamNo, path);
                        }
                    }
                    data = this._emrDocumentMng.Find(filters, fields, template.METemplateNo);
                    if (data is JArray)
                        data = MergeChildArrayData(data as JArray);
                    else
                        data = MergeChildArrayData(data as JToken);
                    if (action.MEEmrActionAllowNullResult)
                    {
                        if (data == null)
                        {
                            if (string.IsNullOrEmpty(action.MEEmrActionPlugin))
                            {
                                _msgNotification.Text = ("Không tìm thấy dữ liệu phù hợp và không có cấu hình [Trình cắm]");
                                return data;
                            }
                            else
                            {
                                data = JToken.FromObject(new object());
                            }
                        }
                    }
                    else
                    {
                        if (data == null)
                        {
                            _msgNotification.Text = ("Không tìm thấy dữ liệu phù hợp");
                            return data;
                        }
                    }
                    //uthv chuyen thanh task de tang toc do
                    PrintMgsLogJson($"DU-LIEU-TRUY-VAN {action.MEEmrActionName}", data);

                    if (beforeBinding != null)
                        errCount = BindingDataToEmrDocument(data, action, group, updateParams, paramList, (d, path) =>
                        {
                            return beforeBinding(d, group, path);
                        }, preData);
                    else
                        errCount = BindingDataToEmrDocument(data, action, group, updateParams, paramList, preData: preData);
                }
                #endregion
                else if (action.MEEmrActionType == EmrActionTypes.AutoAddDoc.ToString())
                {
                    paramList["documentDate"] = document.MEEmrDocumentCreatedDate.ToString("dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
                    data = this.AutoAddDocument(allParams, data, action, group, paramList);
                }
                else if (action.MEEmrActionType == EmrActionTypes.AddExtFileSharedDoc.ToString())
                {
                    this.AddExtFileSharedDocuments(paramList, allParams, data, action, group);
                }
                else if (action.MEEmrActionType == EmrActionTypes.AddImageFileShared.ToString())
                {
                    this.InsertImageFilesShared(paramList, allParams, data, action, group);
                }
                else if (action.MEEmrActionType == EmrActionTypes.DinamapProV100.ToString())
                {
                    errCount = this.InsertDinamapProV100DataAsync(paramList, allParams, data, action, group);
                    return null;
                }
                _msgNotification.Text = (errCount == 0) ? "Dữ liệu đã được cập nhật" : "Có lỗi khi cập nhật dữ liệu. Vui lòng xem thông báo lỗi";

                //chỉ chạy chức năng con trong trường hợp app-to-app các trường hợp khác dùng chức năng phức hợp thay thế
                //if (action.MEEmrActionType == EmrActionTypes.Sql.ToString()
                //    || action.MEEmrActionType == EmrActionTypes.Sql.ToString()
                //    || action.MEEmrActionType == EmrActionTypes.Api.ToString()
                //    || action.MEEmrActionType == EmrActionTypes.Assign.ToString()
                //    || action.MEEmrActionType == EmrActionTypes.Lookup.ToString())
                //{
                //    //TODO trả về danh sách
                //    data = CallChildEmrAction(action, actionRange, group, data);
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Có lỗi xảy ra", MessageBoxButtons.OK, MessageBoxIcon.Error);
                PrintMgsLog($"LOI_GOI_CHUC_NANG_{navigateUri}", ex.ToString());
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
            try
            {
                ParentScreen.Activate();
                _dpnRichEdit.Focus();
            }
            catch (Exception) { }
            return data;
        }
        #endregion

        #region TTBA
        private void SummaryEmr(MEEmrsInfo emr)
        {
            try
            {
                var sumTemplate = BOSApp.GetSystemConfigValue(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.EMR_TEMPLATE_SUMMARY_TEMPLATE);
                if (!string.IsNullOrEmpty(sumTemplate))
                {
                    var sumTypeExcept = BOSApp.GetSystemConfigValue(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.EMR_TEMPLATE_SUMMARY_TYPE_EXCEPT);
                    if (!string.IsNullOrEmpty(sumTypeExcept))
                    {
                        if (sumTypeExcept.Split(';').Select(Int32.Parse).ToList().Contains(emr.FK_MEEmrTypeID))
                            return;
                    }

                    var now = DateTime.Now;
                    var emrSumCtrl = new MEEmrSumsController();
                    var existE = emrSumCtrl.GetFirstObjectByForeignColumn("FK_MEEmrID", emr.MEEmrID) as MEEmrSumsInfo;
                    if (existE != null && existE.MEEmrSumID > 0)
                    {
                        //emrSumCtrl.DeleteObject(existE.MEEmrSumID); // For history. fast, future move to History db
                        // Log
                        var emrSumLogCtrl = new MEEmrSumLogsController();
                        emrSumLogCtrl.CreateObject(new MEEmrSumLogsInfo
                        {
                            AACreatedUser = BOSApp.CurrentUser,
                            FK_MEEmrID = existE.FK_MEEmrID,
                            MEEmrSumCode = existE.MEEmrSumCode,
                            MEEmrSumStoreCode = existE.MEEmrSumStoreCode,
                            MEEmrSumDate = existE.MEEmrSumDate,
                            MEEmrSumRemark = existE.MEEmrSumRemark,
                            MEEmrSumFileName = existE.MEEmrSumFileName,
                            MEEmrSumMongoID = existE.MEEmrSumMongoID,
                            MEEmrSumStatus = existE.MEEmrSumStatus,
                            MEEmrSumXMLStatus = existE.MEEmrSumXMLStatus,
                            MEEmrSumXMLDesc = existE.MEEmrSumXMLDesc,
                            MEEmrSumXMLDate = existE.MEEmrSumXMLDate,
                            MEEmrSumXMLSentNum = existE.MEEmrSumXMLSentNum
                        });
                        // End Log
                        existE.MEEmrSumDate = now;
                        existE.MEEmrSumRemark = string.Empty;
                        existE.MEEmrSumFileName = string.Empty;
                        //existE.MEEmrSumMongoID = string.Empty;
                        existE.MEEmrSumStatus = EmrSumStatus.Active.ToString();
                        existE.MEEmrSumXMLStatus = EmrSumXMLStatus.None.ToString();
                        existE.MEEmrSumXMLDesc = string.Empty;
                        existE.MEEmrSumXMLDate = DateTime.MaxValue;
                        existE.MEEmrSumXMLSentNum = 0;
                        emrSumCtrl.UpdateObject(existE);
                    }
                    else
                    {
                        var sumE = new MEEmrSumsInfo
                        {
                            AACreatedUser = BOSApp.CurrentUser,
                            FK_MEEmrID = emr.MEEmrID,
                            MEEmrSumCode = BOSApp.GetMainObjectNo("MEEmrSum"),
                            MEEmrSumStoreCode = string.Empty,
                            MEEmrSumDate = now,
                            MEEmrSumRemark = string.Empty,
                            MEEmrSumFileName = string.Empty,
                            MEEmrSumMongoID = string.Empty,
                            MEEmrSumStatus = EmrSumStatus.Active.ToString(),
                            MEEmrSumXMLStatus = EmrSumXMLStatus.None.ToString(),
                            MEEmrSumXMLDesc = string.Empty,
                            MEEmrSumXMLDate = DateTime.MaxValue,
                            MEEmrSumXMLSentNum = 0
                        };
                        emrSumCtrl.CreateObject(sumE);
                        BOSApp.UpdateObjectNumbering("MEEmrSum");
                    }
                }
            }
            catch (Exception ex)
            {
                _sysHelper.LogTxt("error", ex.ToString());
            }
        }

        internal void ViewEmrSum()
        {
            var emr = _entity.MainObject as MEEmrsInfo;
            var emrSumCtrl = new MEEmrSumsController();
            var emrSum = emrSumCtrl.GetObjectByEmr(emr.MEEmrID);
            if (emrSum == null)
            {
                MessageBox.Show("Không tìm thấy thông tin Tóm tắt bệnh án.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            if (emrSum.MEEmrSumXMLStatus == EmrSumXMLStatus.None.ToString())
            {
                MessageBox.Show("Tóm tắt bệnh án đang chờ khởi tạo.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            var gui = new guiPreviewPdf(new MEEmrDocumentsInfo
            {
                FK_MEEmrID = emr.MEEmrID,
                MEEmrDocumentFile = emrSum.MEEmrSumFileName,
                MEEmrDocumentFileExt = "pdf"
            }, "Report", emrSum.MEEmrSumXMLCombine);
            gui.StartPosition = FormStartPosition.WindowsDefaultLocation;
            gui.Module = this;
            //gui.WindowState = FormWindowState.Maximized;
            gui.Show();
            BOSProgressBar.Close();
        }

        private void UpdateEmrSumWhenOpenEmr(MEEmrsInfo emr)
        {
            var now = DateTime.Now;
            var emrSumCtrl = new MEEmrSumsController();
            var existE = emrSumCtrl.GetFirstObjectByForeignColumn("FK_MEEmrID", emr.MEEmrID) as MEEmrSumsInfo;
            if (existE != null && existE.MEEmrSumID > 0)
            {
                //emrSumCtrl.DeleteObject(existE.MEEmrSumID); // For history. fast, future move to History db
                // Log
                var emrSumLogCtrl = new MEEmrSumLogsController();
                emrSumLogCtrl.CreateObject(new MEEmrSumLogsInfo
                {
                    AACreatedUser = BOSApp.CurrentUser,
                    FK_MEEmrID = existE.FK_MEEmrID,
                    MEEmrSumCode = existE.MEEmrSumCode,
                    MEEmrSumStoreCode = existE.MEEmrSumStoreCode,
                    MEEmrSumDate = existE.MEEmrSumDate,
                    MEEmrSumRemark = existE.MEEmrSumRemark,
                    MEEmrSumFileName = existE.MEEmrSumFileName,
                    MEEmrSumMongoID = existE.MEEmrSumMongoID,
                    MEEmrSumStatus = existE.MEEmrSumStatus,
                    MEEmrSumXMLStatus = existE.MEEmrSumXMLStatus,
                    MEEmrSumXMLDesc = existE.MEEmrSumXMLDesc,
                    MEEmrSumXMLDate = existE.MEEmrSumXMLDate,
                    MEEmrSumXMLSentNum = existE.MEEmrSumXMLSentNum
                });
                // End Log
                existE.MEEmrSumDate = now;
                existE.MEEmrSumRemark = string.Empty;
                existE.MEEmrSumFileName = string.Empty;
                //existE.MEEmrSumMongoID = string.Empty;
                existE.MEEmrSumStatus = EmrSumStatus.Hide.ToString();
                existE.MEEmrSumXMLStatus = EmrSumXMLStatus.None.ToString();
                existE.MEEmrSumXMLDesc = string.Empty;
                existE.MEEmrSumXMLDate = DateTime.MaxValue;
                existE.MEEmrSumXMLSentNum = 0;
                emrSumCtrl.UpdateObject(existE);
            }
        }
        private object GetJsonMongo(MEEmrsInfo emr, MEEmrDocumentsInfo document)
        {
            try
            {
                if (document != null)
                {
                    var filters = new Dictionary<string, object>
                        {
                            { "FK_MEEmrID", new Tuple<MongoFilter, object>(MongoFilter.Eq, emr.MEEmrID) }
                        };
                    var fields = new Dictionary<string, string>();
                    if (!fields.ContainsKey("MEEmrDocumentContent"))
                    {
                        fields.Add("MEEmrDocumentContent", "$MEEmrDocumentContent");
                    }
                    filters.Add("_id", new Tuple<MongoFilter, object>(MongoFilter.Eq, ObjectId.Parse(document.MEEmrDocumentMongoID)));
                    var dataMongo = _emrDocumentMng.Find(filters, fields, document.MEEmrDocumentNo);
                    dataMongo = dataMongo is JArray ? _helper.MergeChildArrayData(dataMongo as JArray) : _helper.MergeChildArrayData(dataMongo as JToken);
                    return dataMongo;
                }
                else
                    return null;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        private Dictionary<string, object> GetParams(MEEmrsInfo emr, MEEmrDocumentsInfo document, List<MEParamsInfo> reportParams)
        {
            //GetHardParamList(emr); // default data
            try
            {
                var paramList = new Dictionary<string, object>();
                foreach (var param in reportParams)
                {
                    var valueParams = new List<object>();
                    object valueParamReport = null;
                    #region MEParamReportMap
                    if (document != null)
                    {
                        var filters = new Dictionary<string, object>
                        {
                            { "FK_MEEmrID", new Tuple<MongoFilter, object>(MongoFilter.Eq, emr.MEEmrID) }
                        };
                        var fields = new Dictionary<string, string>();
                        if (!fields.ContainsKey("MEEmrDocumentContent"))
                        {
                            fields.Add("MEEmrDocumentContent", "$MEEmrDocumentContent");
                        }
                        var dataMongo = _emrDocumentMng.Find(filters, fields, document.MEEmrDocumentNo);
                        dataMongo = dataMongo is JArray ? _helper.MergeChildArrayData(dataMongo as JArray) : _helper.MergeChildArrayData(dataMongo as JToken);

                        valueParamReport = ((JToken)dataMongo).Type == JTokenType.Array
                            ? (dataMongo as JArray).SelectTokens($"$..{param.MEParamNo}")
                            : (dataMongo as JObject).SelectToken($"$..{param.MEParamNo}");

                        if (valueParamReport == null || string.IsNullOrEmpty(valueParamReport.ToString()))
                        {
                            valueParamReport = null;
                        }
                    }
                    #endregion

                    if (valueParamReport != null)
                        valueParams.Add(valueParamReport);
                    paramList.Add(param.MEParamNo, valueParams);
                }
                return paramList;
            }
            catch (Exception ex)
            {
                return new Dictionary<string, object>();
            }
        }
        private bool GeneralXML(Dictionary<string, string> paramList, List<MEParamsInfo> paramsReport, string typeXML, string xmlFile, string localPath, string serverPath)
        {
            var xmlList = new List<XMLModel>();
            foreach (KeyValuePair<string, string> entry in paramList)
            {
                // do something with entry.Value or entry.Key
                var tag = string.Empty;
                var xmlValue = string.Empty;
                var xmlGroup = 1;
                var xmlLevel = 1;
                var xmlOrder = 1;
                var paramE = paramsReport.FirstOrDefault(m => m.MEParamNo.Equals(entry.Key));
                if (paramE != null)
                {
                    tag = paramE.MEParamXMLTag;
                    if (string.IsNullOrEmpty(tag))
                    {
                        tag = paramE.MEParamNo;
                    }
                    xmlList.Add(new XMLModel
                    {
                        Tag = tag,
                        Value = entry.Value,
                        Group = xmlGroup,
                        Level = xmlLevel,
                        Order = xmlOrder
                    });
                }
            }

            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                XmlDeclaration xmlDeclaration = xmlDoc.CreateXmlDeclaration("1.0", "utf-8", string.Empty);
                xmlDoc.AppendChild(xmlDeclaration);

                XmlElement parentNode = xmlDoc.CreateElement(typeXML);
                XmlAttribute xsd = xmlDoc.CreateAttribute("xmlns:xsd");
                xsd.Value = "http://www.w3.org/2001/XMLSchema";
                parentNode.Attributes.Append(xsd);
                XmlAttribute xsi = xmlDoc.CreateAttribute("xmlns:xsi");
                xsi.Value = "http://www.w3.org/2001/XMLSchema-instance";
                parentNode.Attributes.Append(xsi);
                foreach (var item in xmlList.OrderBy(m => m.Order))
                {
                    // Group, Level do later. analytics
                    XmlElement itemNode = xmlDoc.CreateElement(item.Tag);
                    itemNode.InnerText = item.Value;
                    parentNode.AppendChild(itemNode);
                }
                xmlDoc.AppendChild(parentNode);

                var fullXMLLocal = string.Format(@"{0}", localPath);
                xmlDoc.Save(fullXMLLocal);
                _ftpFileMng.UploadFile(serverPath, xmlFile, fullXMLLocal);

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        private bool ConvertJsonToXML(string json, string typeXML, string xmlFile, string localPath, string serverPath)
        {
            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc = JsonConvert.DeserializeXmlNode(json, typeXML);

                var fullXMLLocal = string.Format(@"{0}", localPath);
                xmlDoc.Save(fullXMLLocal);
                _ftpFileMng.UploadFile(serverPath, xmlFile, fullXMLLocal);
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        private Dictionary<string, object> ConvertXMLToDic(MEEmrDocumentsInfo document)
        {
            string serverPath = $"/Emr/{document.FK_MEEmrID}/";
            string localPath = string.Format(@"{0}\Emr\{1}\{2}.{3}", _documentPath, document.FK_MEEmrID, document.MEEmrDocumentFile, "xml");
            _ftpFileMng.DownloadFile(serverPath, $"{document.MEEmrDocumentFile}.xml", localPath);
            XDocument doc = XDocument.Load(localPath);
            Dictionary<string, object> dataDictionary = new Dictionary<string, object>();

            foreach (XElement element in doc.Descendants().Where(p => p.HasElements == false))
            {
                int keyInt = 0;
                string keyName = element.Name.LocalName;

                while (dataDictionary.ContainsKey(keyName))
                {
                    keyName = element.Name.LocalName + "_" + keyInt++;
                }

                dataDictionary.Add(keyName, element.Value);
            }
            return dataDictionary;
        }
        #endregion

        private void SignalRApp(MEEmrActionsInfo action, string msg, int cacheTimeOut, string group, List<MEEmrActionParamsInfo> updateParams, Dictionary<string, object> paramList, object preData, string transactionAction)
        {
            var serverUri = BOSApp.GetSystemConfigValue(SysCfgConsts.SYSTEM_CONFIGS, SysCfgConsts.HIS_SIGNALR_HUB);
            if (!string.IsNullOrEmpty(serverUri))
            {
                if (_checkSystem)
                {
                    _sysHelper.LogTxt("information", $"App to App HIS KV. Dữ liệu: {msg}");
                }
                var dataHISReceived = string.Empty;
                var userName = BOSApp.CurrentUsersInfo.ADUserName;
                if (_checkSystem)
                {
                    _sysHelper.LogTxt("information", $"Khởi tạo HubConnection {serverUri}?user={userName}");
                }
                var hubConnection = new HubConnection(serverUri, $"user={userName}");
                IHubProxy hubProxy = hubConnection.CreateHubProxy("SignalRHub");
                hubProxy.On<string, string>("ReceiveMessage", (user, message) =>
                {
                    if (_checkSystem)
                    {
                        _sysHelper.LogTxt("information", $"ReceiveMessage at {user}: {message}.");
                    }
                    dataHISReceived = message;
                    hubConnection.Stop();
                    listener_MessageFromHisReceivedSignalR(dataHISReceived, cacheTimeOut, paramList, transactionAction);
                });
                try
                {
                    hubConnection.Start().GetAwaiter().GetResult();
                    if (_checkSystem)
                    {
                        _sysHelper.LogTxt("information", $"Connected to server at {serverUri}");
                    }
                }
                catch (Exception ex)
                {
                    if (_checkSystem)
                    {
                        _sysHelper.LogTxt("error", $"Unable to connect to server: Start server before connecting clients.\n{ex.Message}");
                    }
                    _stateAppHISKV = 0;
                    return;
                }
                try
                {
                    if (_checkSystem)
                    {
                        _sysHelper.LogTxt("information", $"SendHISMessage user: {userName}, info: {msg}");
                    }
                    hubProxy.Invoke("SendHISMessage", userName, msg).Wait();
                    _stateAppHISKV = 1;
                    if (_checkSystem)
                    {
                        _sysHelper.LogTxt("information", "Đã gởi dữ liệu đến HIS KV. Đang chờ dữ liệu trả về...");
                    }
                }
                catch (Exception exSH)
                {
                    _sysHelper.LogTxt("error", $"SendHISMessage: {exSH.Message}");
                    _stateAppHISKV = 0;
                    return;
                }
                PrintMgsLog("DA-GOI-YEU-CAU-DEN-HIS-KV", msg);
                _msgNotification.Text = "Đã gởi dữ liệu đến HIS. Đang chờ dữ liệu trả về...";
            }
            else
            {
                PrintMgsLog("KHONG-CAU-HINH-HUB-HIS-KV", msg);
                _msgNotification.Text = "Không gởi được dữ liệu đến HIS. Không cấu hình đường dẫn hub HIS KV...";
                if (_checkSystem)
                {
                    _sysHelper.LogTxt("information", "Không gởi được dữ liệu đến HIS. Không cấu hình đường dẫn hub HIS KV...");
                }
                _stateAppHISKV = 0;
                return;
            }
        }

        private void listener_MessageFromHisReceivedSignalR(string message, int cacheTimeOut, Dictionary<string, object> paramList, string transactionAction)
        {
            if (BOSApp.MainScreen.InvokeRequired)
                BOSApp.MainScreen.BeginInvoke((Action)(() =>
                {
                    BOSApp.MainScreen.Activate();
                }));
            else
            {
                BOSApp.MainScreen.Activate();
            }
            // Gui Json ko data.
            PrintMgsLog("THONG_DIEP_TU_HIS", message);
            if (string.IsNullOrEmpty(message))
            {
                if (_checkSystem)
                {
                    _sysHelper.LogTxt("information", $"THONG_DIEP_TU_HIS: Không nhận được dữ liệu.");
                }
                if (_msgNotification.InvokeRequired)
                {
                    _msgNotification.BeginInvoke((MethodInvoker)delegate ()
                    {
                        _msgNotification.Text = "Không nhận được dữ liệu từ HIS.";
                    });
                }
                else
                {
                    _msgNotification.Text = "Không nhận được dữ liệu từ HIS.";
                }
                _stateAppHISKV = 0;
                return;
            }
            // CALL API GET DATA
            var response = new Dictionary<string, object>();
            var document = _entity.ModuleObjects[TableName.MEEmrDocumentsTableName] as MEEmrDocumentsInfo;
            try
            {
                if (_sysHelper.IsJson(message))
                {
                    response = JsonConvert.DeserializeObject<Dictionary<string, object>>(message);
                }
                else
                {
                    //very hard
                    var docXml = new XmlDocument();
                    docXml.LoadXml(message);
                    var json = JsonConvert.SerializeXmlNode(docXml);
                    response = (JsonConvert.DeserializeObject<Dictionary<string, object>>(json)["root"] as JObject).ToObject<Dictionary<string, object>>();
                }
            }
            catch (Exception ex)
            {
                Trace.TraceError("HIS KV APPTOAPP ERROR: {0}:{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), ex);
                ShowFlashNotification("Có lỗi xảy ra", 3000);
                _stateAppHISKV = 0;
                return;
                //throw;
            }
            if (_checkSystem)
            {
                _sysHelper.LogTxt("information", $"THONG_TIN_HIS: {response}");
            }
            if (!response.ContainsKey("msg"))
            {
                _stateAppHISKV = 0;
                if (_msgNotification.InvokeRequired)
                {
                    _msgNotification.BeginInvoke((MethodInvoker)delegate ()
                    {
                        _msgNotification.Text = "Có lỗi khi cập nhật dữ liệu. Vui lòng xem thông báo lỗi.";
                    });
                }
                else
                {
                    _msgNotification.Text = "Có lỗi khi cập nhật dữ liệu. Vui lòng xem thông báo lỗi.";
                }
                return;
            }
            string patientNo = response.ContainsKey("patientNo") ? response["patientNo"].ToString() : response["patientno"].ToString();
            if (patientNo != this._entity.MEPatient.MEPatientNo)
            {
                _stateAppHISKV = 0;
                if (_msgNotification.InvokeRequired)
                {
                    _msgNotification.BeginInvoke((MethodInvoker)delegate ()
                    {
                        _msgNotification.Text = ("Ứng dụng nhận được dữ liệu nhưng của bệnh nhân có mã: " + response["patientNo"].ToString() + " Vui lòng chọn đúng bệnh nhân, bệnh án và mẫu bệnh án");
                    });
                }
                else
                {
                    _msgNotification.Text = ("Ứng dụng nhận được dữ liệu nhưng của bệnh nhân có mã: " + response["patientNo"].ToString() + " Vui lòng chọn đúng bệnh nhân, bệnh án và mẫu bệnh án");
                }
                return;
            }
            string documentNo = response.ContainsKey("documentNo") ? response["documentNo"].ToString() : response["documentno"].ToString();
            string documentName = response.ContainsKey("documentName") ? response["documentName"].ToString() : response["documentname"].ToString();
            string documentDate = response.ContainsKey("documentDate") ? response["documentDate"].ToString() : response["documentdate"].ToString();
            if (documentNo != document.MEEmrDocumentFile)
            {
                _stateAppHISKV = 0;
                if (_msgNotification.InvokeRequired)
                {
                    _msgNotification.BeginInvoke((MethodInvoker)delegate ()
                    {
                        _msgNotification.Text = string.Format("Ứng dụng nhận được dữ liệu nhưng của mẫu bệnh án: '{0}' ngày {1} {2}"
                                                , documentName
                                                , documentDate
                                                , "Vui lòng chọn đúng bệnh nhân, bệnh án và mẫu bệnh án");
                    });
                }
                else
                {
                    _msgNotification.Text = string.Format("Ứng dụng nhận được dữ liệu nhưng của mẫu bệnh án: '{0}' ngày {1} {2}"
                    , documentName
                    , documentDate
                    , "Vui lòng chọn đúng bệnh nhân, bệnh án và mẫu bệnh án");
                }
                return;
            }
            if (response == null)
            {
                _stateAppHISKV = 0;
                if (_msgNotification.InvokeRequired)
                {
                    _msgNotification.BeginInvoke((MethodInvoker)delegate ()
                    {
                        _msgNotification.Text = ("Không có dữ liệu trả về từ HIS. Liên hệ quản trị viên để biết thêm chi tiết");
                    });
                }
                else
                {
                    _msgNotification.Text = ("Không có dữ liệu trả về từ HIS. Liên hệ quản trị viên để biết thêm chi tiết");
                }
                return;
            }

            // Store value in Mongo
            var responseString = string.Join(", ", response.Select(kv => $"{kv.Key}={kv.Value}"));
            var logMongo = new Clas.Model.Base.LogMongo
            {
                Transaction = transactionAction,
                ValObj = responseString
            };
            InsertMongoLog(logMongo, "apptoapp_kv");
            _stateAppHISKV = 2;

            #region Bind Data. Comment now... use next action
            //int errCount = 0;
            //var action = _actionsController.GetObjectByNo(response["msg"].ToString()) as ActionsInfo;
            //if (action != null)
            //{
            //    var group = string.Empty;
            //    if (response.ContainsKey(EmrParam.GuidTag))
            //        group = response[EmrParam.GuidTag].ToString();

            //    //TODO, BUG trong truong hop nguoi dung chon document khac truoc khi his tra ve du lieu
            //    if (string.IsNullOrEmpty(group))
            //        group = document.MEEmrDocumentGuid;

            //    var allParams = AppMemCache.GetActionParamsFromDict(action.MEEmrActionID);
            //    var updateParams = allParams.Where(o => o.MEEmrActionParamRequest == false && o.FK_MEParamID > 0).ToList();

            //    var requestParams = new Dictionary<string, object>(response);
            //    try
            //    {
            //        requestParams.Remove("data");
            //    }
            //    catch (Exception exRm)
            //    {
            //        _sysHelper.LogTxt("error", $"Lỗi không ảnh hưởng hệ thống: {exRm}");
            //    }


            //    object preData = null;
            //    if (_appPreData != null)
            //        preData = _appPreData;

            //    if (response.ContainsKey("preData"))
            //        preData = response["preData"];

            //    var hisId = response.ContainsKey("HIS_ID") ? response["HIS_ID"].ToString() : string.Empty;
            //    if (string.IsNullOrEmpty(hisId))
            //    {
            //        if (_checkSystem)
            //        {
            //            _sysHelper.LogTxt("information", $"HIS ID không tìm thấy.");
            //        }
            //        return;
            //    }

            //    //var paramHis = new Dictionary<string, object>
            //    //    {
            //    //        { "HIS_ID", hisId }
            //    //    };
            //    if (_checkSystem)
            //    {
            //        var prString = string.Join(", ", paramList.Select(kv => $"{kv.Key}={kv.Value}"));
            //        _sysHelper.LogTxt("information", $"THONG_TIN_HIS Action: {action} - Param: HIS_ID = {hisId}");
            //    }
            //    var data = GetApiData(cacheTimeOut, action, paramList);
            //    if (_checkSystem)
            //    {
            //        _sysHelper.LogTxt("information", $"API-HIS Mã chức năng: {action.MEEmrActionNo} - uri: {action.MEEmrActionUri} - Data: {data}");
            //    }
            //    if (data == null)
            //    {
            //        _msgNotification.Text = $"Không có dữ liệu trả về từ api. Mã chức năng: {action.MEEmrActionNo} - uri: {action.MEEmrActionUri}";
            //        if (_checkSystem)
            //        {
            //            _sysHelper.LogTxt("information", $"Không có dữ liệu trả về từ api. Mã chức năng: {action.MEEmrActionNo} - uri: {action.MEEmrActionUri}");
            //        }
            //        return;
            //    }
            //    if (_checkSystem)
            //    {
            //        _sysHelper.LogTxt("information", "BindingDataToEmrDocument - Bắt đầu đưa dữ liệu vào EMR.");
            //    }

            //    // Test Happy Case
            //    //data = JObject.Parse("{\"danhsachchidinh\":[{\"sophieu\":\"RI.20.0001203\",\"bacsichidinh\":\"lname 1 fname 1\",\"paraid\":13008,\"report\":\"CT-Scanner\",\"report_id\":6,\"tenchidinh\":\"CT scaner s\u1ecd n\u00e3o\"}]}");

            //    errCount = BindingDataToEmrDocument(data, action, group, updateParams, requestParams, preData: preData);

            //    //HIS co the goi du lieu nhieu lan cho 1 req tu EMR
            //    //_appPreData = null;
            //    CallChildEmrAction(action, null, group, data);
            //}
            //if (_msgNotification.InvokeRequired)
            //{
            //    _msgNotification.BeginInvoke((MethodInvoker)delegate () {
            //        _msgNotification.Text = errCount == 0 ? "Dữ liệu đã được cập nhật" : "Có lỗi khi cập nhật dữ liệu. Vui lòng xem thông báo lỗi";
            //    });
            //}
            //else
            //{
            //    _msgNotification.Text = errCount == 0 ? "Dữ liệu đã được cập nhật" : "Có lỗi khi cập nhật dữ liệu. Vui lòng xem thông báo lỗi";
            //}
            #endregion
        }
        private object AppToAppVB(Dictionary<string, object> message)
        {
            object result = new object();
            string method = string.Empty;
            try
            {
                var emrNo = string.Empty;
                var patientNo = string.Empty;
                var documentNo = string.Empty;
                var userID = 0;

                if (message.ContainsKey("method"))
                    method = message["method"].ToString();
                else
                    throw new Exception("Thẻ chức năng không có method!");

                if (message.ContainsKey("emrNo"))
                    emrNo = message["emrNo"].ToString();
                else
                    throw new Exception("Không có emrNo!");

                if (message.ContainsKey("patientNo"))
                    patientNo = message["patientNo"].ToString();
                else
                    throw new Exception("Không có patientNo!");

                if (message.ContainsKey("documentNo"))
                    documentNo = message["documentNo"].ToString();
                else
                    throw new Exception("Không có documentNo!");

                if (!string.IsNullOrEmpty(BOSApp.CurrentUsersInfo.ADUserHISID))
                    userID = Convert.ToInt32(BOSApp.CurrentUsersInfo.ADUserHISID);
                else
                    throw new Exception("Tài khoản không có HISID!");

                string sophieuthuoc = message.ContainsKey("id_don_thuoc") ? message["id_don_thuoc"].ToString() : string.Empty;
                string sochidinh = message.ContainsKey("id_phieu_chi_dinh") ? message["id_phieu_chi_dinh"].ToString() : string.Empty;
                DateTime date = message.ContainsKey("Date") ? Convert.ToDateTime(message["Date"].ToString()) : DateTime.Now;
                DateTime time = message.ContainsKey("time") ? Convert.ToDateTime(message["time"].ToString()) : DateTime.Now;

                DateTime tdtThoiGian = new DateTime(date.Year, date.Month, date.Day, time.Hour, time.Minute, time.Second);
                var sContractor = new sContractor()
                {
                    NsdID = userID,
                    MaBN = patientNo,
                    TdtThoiGian = tdtThoiGian,
                    SoBA = emrNo,
                    MaTDT = documentNo,
                    soPhieu = method == "ChiDinh" ? sochidinh : sophieuthuoc,
                    kieu = method

                };
                sContractor.formAction = message.ContainsKey("form") ? message["form"].ToString() : "";
                using (var chanelSend = RabbitMqConnectionManager.CreateChannel())
                {
                    var routingKey = RabbitMqConnectionManager.GetQueueName("OpenForm");
                    chanelSend.ExchangeDeclare(exchange: RabbitMqExchange.EmrToHis_Exchange, type: ExchangeType.Direct, durable: false, autoDelete: false, null);
                    chanelSend.BasicPublish(exchange: RabbitMqExchange.EmrToHis_Exchange, routingKey: routingKey, null, body: Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(sContractor)));
                }
                _waittingForHIS = true;
                //Process.GetCurrentProcess().StartInfo.WindowStyle = ProcessWindowStyle.Minimized;
                Process currentProcess = Process.GetCurrentProcess();
                IntPtr hWnd = currentProcess.MainWindowHandle;

                // Minimize the window
                ShowWindow(hWnd, SW_MINIMIZE);
                //while (_waittingForHIS) ;
                BOSApp._hisWaitHandle.Reset();
                BOSApp._hisWaitHandle.WaitOne(); // <-- BLOCK tại đây nhưng không chiếm CPU

                try
                {
                    var handle = Process.GetCurrentProcess().MainWindowHandle; // Or use: Process.GetCurrentProcess().MainWindowHandle
                    ShowWindow(handle, SW_RESTORE);       // Restore if minimized
                    SetForegroundWindow(handle);
                }
                catch (Exception ex)
                {

                }

                if (!string.IsNullOrEmpty(_dataJsonFromHIS))
                {
                    result = JsonConvert.DeserializeObject<dynamic>(_dataJsonFromHIS);
                }

                return result;
            }
            catch (Exception ex)
            {
                // Store value in Mongo
                var logMongo = new Clas.Model.Base.LogMongo
                {
                    Transaction = method,
                    ValObj = ex.Message
                };
                InsertMongoLog(logMongo, "apptoapp_vb_error");
                Trace.TraceError("AppToApp ERROR: {0}:{1}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), ex);
                _sysHelper.LogTxt("error", ex.ToString());
                throw;
            }
        }

        #region Common : End of module
        private string DownloadTemplate(MEEmrDocumentsInfo document)
        {
            string fileName = string.Format(@"{0}\Template\{1}.docx", _documentPath, document.MEEmrDocumentNo);
            try
            {
                _ftpFileMng.DownloadFile("/Template/", document.MEEmrDocumentNo + ".docx", fileName);
            }
            catch (Exception ex)
            {
                Trace.TraceError("FtpFileMng.DownloadFile ERROR: {0}:{1}:{2}:{3}", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"), BOSApp.CurrentUser, ex, fileName);
                Trace.Flush();
                MessageBox.Show($"Không tải được tờ bệnh án từ máy chủ.",
                        "CÓ LỖI XẢY RA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return string.Empty;
            }

            if (!File.Exists(fileName))
            {
                MessageBox.Show("File mẫu bệnh án không tồn tại ở địa chỉ. " + fileName);
                return string.Empty;
            }
            return fileName;
        }

        public string DownloadFtpFileCommon(string serverPath, string fileName, string localPart)
        {
            var localPath = Path.Combine(_documentPath, localPart);
            Directory.CreateDirectory(localPath);
            var fullFileLocal = Path.Combine(localPath, fileName);
            _ftpFileMng.DownloadFile(serverPath, fileName, fullFileLocal);
            return fullFileLocal;
        }
        #endregion        
    }
}
