using com.amtec.action;
using com.amtec.configurations;
using com.amtec.model;
using com.itac.mes.imsapi.domain.container;
using com.itac.oem.common.container.imsapi.utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Ports;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Timers;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Linq;
using System.Runtime.InteropServices;
using ScreenPrinter.com.amtec.forms;
using com.amtec.device;

namespace com.amtec.forms
{
    public partial class MainView : Form
    {
        public ApplicationConfiguration config;
        IMSApiSessionContextStruct sessionContext;
        public bool isScanProcessEnabled = false;
        public InitModel initModel;
        public LanguageResources res;
        public string UserName = "";
        private string indata = "";
        private DateTime PFCStartTime = DateTime.Now;
        List<SerialNumberData> serialNumberArray = new List<SerialNumberData>();
        public delegate void HandleInterfaceUpdateTopMostDelegate(string sn, string message);
        public HandleInterfaceUpdateTopMostDelegate topmostHandle;
        public TopMostForm topmostform = null;
        CommonModel commonModel = null;
        public string CaptionName = "";
        private System.Timers.Timer SendTrigerTimer = new System.Timers.Timer();
        public SocketClientHandler cSocket = null;
        private System.Timers.Timer CheckConnectTimer = null;
        private System.Timers.Timer CheckIPITimer = null;
        private System.Timers.Timer RestoreMaterialTimer = null;
        string Supervisor_OPTION = "1";
        string IPQC_OPTION = "1";
        private SocketClientHandler2 checklist_cSocket = null;
        bool isStartLineCheck = true;//开线点检已经获取=true. 过程点检=false
        string workorderType = "NULL";

        #region Init
        public MainView(string userName, DateTime dTime, IMSApiSessionContextStruct _sessionContext)
        {
            InitializeComponent();
            sessionContext = _sessionContext;
            UserName = userName;
            commonModel = ReadIhasFileData.getInstance();
            this.lblLoginTime.Text = dTime.ToString("yyyy/MM/dd HH:mm:ss");
            this.lblUser.Text = userName == "" ? commonModel.Station : userName;
            this.lblStationNO.Text = commonModel.Station;
        }

        private void MainView_Shown(object sender, EventArgs e)
        {
            BackgroundWorker bgWorker = new BackgroundWorker();
            bgWorker.DoWork += new DoWorkEventHandler(bgWorkerInit);
            bgWorker.RunWorkerAsync();
        }
        bool isOK = true;
        private void bgWorkerInit(object sender, DoWorkEventArgs args)
        {
            errorHandler(0, "Application start...", "");
            errorHandler(0, "Version :" + Assembly.GetExecutingAssembly().GetName().Version.ToString(), "");
            res = new LanguageResources();
            config = new ApplicationConfiguration(sessionContext, this);
            InitializeMainGUI init = new InitializeMainGUI(sessionContext, config, this, res);
            initModel = init.Initialize();
            this.InvokeEx(x =>
            {
                //this.tabDocument.Parent = null;
                this.Text = res.MAIN_TITLE + " (" + Assembly.GetExecutingAssembly().GetName().Version.ToString() + ")";
                CaptionName = res.MAIN_TITLE + System.Environment.NewLine + config.StationNumber;
                #region add by qy
                SystemVariable.CurrentLangaugeCode = config.Language;
                InitCintrolLanguage(this);
                #endregion
                //if (this.txbCDAMONumber.Text == "")
                //{
                //    errorHandler(3, Message("msg_No activated work order"), "");
                //    return;
                //}
                
                //显示工单信息
                InitWorkOrderList();

                SetWorkorderGridStatus();
                this.tabWOActived.Refresh();
                this.cmbLayer.Text = "T";

                //工单转换
                if (config.IsNeedTransWO != "Y")
                {
                    //tableLayoutPanel4.ColumnStyles.Remove(tableLayoutPanel4.ColumnStyles[0]);
                    //panel14.Visible = false;
                    panel17.Visible = false;
                }
                else
                {
                    InitGetHasTransWO();
                    this.dgvTransWorkOrder.RowsDefaultCellStyle.BackColor = Color.LightGray;
                }

                //生产检查
                if (config.IsNeedProductionInspection != "ENABLE")
                {
                    panel16.Visible = false;
                }

                //既不需要生产检查又不需要工单转换
                if (config.IsNeedProductionInspection != "ENABLE" && config.IsNeedTransWO != "Y")
                {
                    tableLayoutPanel4.ColumnStyles.Remove(tableLayoutPanel4.ColumnStyles[0]);
                    panel14.Visible = false;
                }


                if (config.AUTH_CHECKLIST_APP_TEAM != "" && config.AUTH_CHECKLIST_APP_TEAM != null)
                {
                    string[] teams = config.AUTH_CHECKLIST_APP_TEAM.Split(';');
                    string[] items = teams[0].Split(',');
                    string Super = items[0];
                    Supervisor_OPTION = items[1];
                    string[] IPQCitems = teams[1].Split(',');
                    string IP = IPQCitems[0];
                    IPQC_OPTION = IPQCitems[1];
                }

                if (config.LAYER_DISPLAY == "")
                {
                    this.btnShowPCB.Visible = false;
                }
                Application.DoEvents();

                cSocket = new SocketClientHandler(this);
                bool isOK2 = cSocket.connect(config.IPAddress, config.Port);
                if (isOK2)
                {
                    //持续发送PING信号
                    GetTimerStart();

                    LoadYield();

                    //上料
                    InitSetupGrid();

                    //InitTaskData();
                    InitFailureMapTable();
                    InitEquipmentGridEXT();
                    //InitSqueegeeValue();
                    InitDocumentGrid();
                    InitIPIWorkOrderType();
                    //StrippedEquipmentFromStation();
                    ShowTopWindow();
                    this.txbCDADataInput.Focus();
                    InitShift(txbCDAMONumber.Text);
                    if (config.RESTORE_TIME != "" && config.RESTORE_TREAD_TIMER != "")
                    {
                        GetRestoreTimerStart();
                        ReadRestoreFile();
                    }
                    SetTipMessage(MessageType.OK, Message("msg_Initialize Success"));
                    if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")//20161208 edit by qy
                    {
                        InitShift2(txbCDAMONumber.Text);
                        InitWorkOrderType();
                        this.tabCheckList.Parent = null;
                        checklist_cSocket = new SocketClientHandler2(this);
                        isOK = checklist_cSocket.connect(config.CHECKLIST_IPAddress, config.CHECKLIST_Port);
                        if (isOK)
                        {
                            if (!CheckShiftChange2())
                            {
                                InitTaskData_SOCKET("开线点检;设备点检");
                                isStartLineCheck = true;
                            }
                            else
                            {
                                if (!ReadCheckListFile())//20161214 edit by qy
                                {
                                    InitTaskData_SOCKET("开线点检");
                                    isStartLineCheck = true;
                                }
                            }
                        }
                    }
                    else
                    {
                        InitTaskData();
                        this.tabCheckListTable.Parent = null;
                    }
                    if (config.DataOutputInterface != null && config.DataOutputInterface != "")
                        SendSN(config.LIGHT_CHANNEL_ON);
                }
            });
        }

        #region add by qy
        public void InitCintrolLanguage(Form form)
        {
            MutiLanguages lang = new MutiLanguages();
            foreach (Control ctl in form.Controls)
            {
                lang.InitLangauge(ctl);
                if (ctl is TabControl)
                {
                    lang.InitLangaugeForTabControl((TabControl)ctl);
                }
            }

            //Controls不包含ContextMenuStrip，可用以下方法获得
            System.Reflection.FieldInfo[] fieldInfo = this.GetType().GetFields(System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            for (int i = 0; i < fieldInfo.Length; i++)
            {
                switch (fieldInfo[i].FieldType.Name)
                {
                    case "ContextMenuStrip":
                        ContextMenuStrip contextMenuStrip = (ContextMenuStrip)fieldInfo[i].GetValue(this);
                        lang.InitLangauge(contextMenuStrip);
                        break;
                }
            }
        }

        public string Message(string messageId)
        {
            return MutiLanguages.ParserString("$" + messageId);
        }
        #endregion
        #endregion

        #region delegate
        public delegate void errorHandlerDel(int typeOfError, String logMessage, String labelMessage);
        public void errorHandler(int typeOfError, String logMessage, String labelMessage)
        {
            if (txtConsole.InvokeRequired)
            {
                errorHandlerDel errorDel = new errorHandlerDel(errorHandler);
                Invoke(errorDel, new object[] { typeOfError, logMessage, labelMessage });
            }
            else
            {
                String errorBuilder = null;
                String isSucces = null;
                switch (typeOfError)
                {
                    case 0:
                        isSucces = "SUCCESS";
                        txtConsole.SelectionColor = Color.Black;
                        errorBuilder = "# " + DateTime.Now.ToString("HH:mm:ss") + " >> " + isSucces + " >< " + logMessage + "\n";
                        SetTipMessage(MessageType.OK, logMessage);
                        LogHelper.Info(logMessage);
                        break;
                    case 1:
                        isSucces = "SUCCESS";
                        txtConsole.SelectionColor = Color.Blue;
                        errorBuilder = "# " + DateTime.Now.ToString("HH:mm:ss") + " >> " + isSucces + " >< " + logMessage + "\n";
                        SetTipMessage(MessageType.OK, logMessage);
                        LogHelper.Info(logMessage);
                        break;
                    case 2:
                        isSucces = "FAIL";
                        txtConsole.SelectionColor = Color.Red;
                        errorBuilder = "# " + DateTime.Now.ToString("HH:mm:ss") + " >> " + isSucces + " >< " + logMessage + "\n";
                        SetTipMessage(MessageType.Error, logMessage);
                        LogHelper.Error(logMessage);
                        break;
                    case 3:
                        isSucces = "FAIL";
                        txtConsole.SelectionColor = Color.Red;
                        errorBuilder = "# " + DateTime.Now.ToString("HH:mm:ss") + " >> " + isSucces + " >< " + logMessage + "\n";
                        SetTipMessage(MessageType.Error, logMessage);
                        LogHelper.Error(logMessage);
                        break;
                    default:
                        isSucces = "FAIL";
                        txtConsole.SelectionColor = Color.Red;
                        errorBuilder = "# " + DateTime.Now.ToString("HH:mm:ss") + " >> " + isSucces + " >< " + logMessage + "\n";
                        LogHelper.Error(logMessage);
                        break;
                }
                SetStatusLabelText(logMessage);
                txtConsole.AppendText(errorBuilder);
                txtConsole.ScrollToCaret();
            }
        }

        public delegate void SetTipMessageDel(MessageType strType, string strMessage);
        private void SetTipMessage(MessageType strType, string strMessage)
        {
            if (this.messageControl1.InvokeRequired)
            {
                SetTipMessageDel messageDel = new SetTipMessageDel(SetTipMessage);
                Invoke(messageDel, new object[] { strType, strMessage });
            }
            else
            {
                switch (strType)
                {
                    case MessageType.OK:
                        this.messageControl1.BackColor = Color.FromArgb(184, 255, 160);
                        this.messageControl1.PicType = @"pic\ok.png";
                        this.messageControl1.Title = "OK";
                        this.messageControl1.Content = strMessage;
                        break;
                    case MessageType.Error:
                        this.messageControl1.BackColor = Color.Red;
                        this.messageControl1.PicType = @"pic\Close.png";
                        this.messageControl1.Title = "Error Message";
                        this.messageControl1.Content = strMessage;
                        break;
                    case MessageType.Instruction:
                        this.messageControl1.BackColor = Color.FromArgb(184, 255, 160);
                        this.messageControl1.PicType = @"pic\Instruction.png";
                        this.messageControl1.Title = "Instruction";
                        this.messageControl1.Content = strMessage;
                        break;
                    default:
                        this.messageControl1.BackColor = Color.FromArgb(184, 255, 160);
                        this.messageControl1.PicType = @"pic\ok.png";
                        this.messageControl1.Title = "OK";
                        this.messageControl1.Content = strMessage;
                        break;
                }
            }
        }

        public delegate void SetConnectionTextDel(int typeOfError, string strMessage);
        public void SetConnectionText(int typeOfError, string strMessage)
        {
            if (txtConnection.InvokeRequired)
            {
                SetConnectionTextDel connectDel = new SetConnectionTextDel(SetConnectionText);
                Invoke(connectDel, new object[] { typeOfError, strMessage });
            }
            else
            {
                String errorBuilder = null;
                String isSucces = null;
                switch (typeOfError)
                {
                    case 0:
                        isSucces = "SUCCESS";
                        txtConnection.SelectionColor = Color.Black;
                        errorBuilder = "# " + DateTime.Now.ToString("HH:mm:ss") + " >> " + isSucces + " >< " + strMessage + "\n";
                        //LogHelper.Info(strMessage);
                        break;
                    case 1:
                        isSucces = "FAIL";
                        txtConnection.SelectionColor = Color.Red;
                        errorBuilder = "# " + DateTime.Now.ToString("HH:mm:ss") + " >> " + isSucces + " >< " + strMessage + "\n";
                        LogHelper.Error(strMessage);
                        break;
                    default:
                        isSucces = "FAIL";
                        txtConnection.SelectionColor = Color.Red;
                        errorBuilder = "# " + DateTime.Now.ToString("HH:mm:ss") + " >> " + isSucces + " >< " + strMessage + "\n";
                        break;
                }

                txtConnection.AppendText(errorBuilder);
                txtConnection.ScrollToCaret();
            }
        }

        public void SetStatusLabelText(string strText)
        {
            this.InvokeEx(x => this.lblStatus.Text = strText);
        }

        public string GetWorkOrderValue()
        {
            string str = "";
            this.InvokeEx(x => str = this.txbCDAMONumber.Text);
            return str;
        }

        public string GetPartNumberValue()
        {
            string str = "";
            this.InvokeEx(x => str = this.txbCDAPartNumber.Text);
            return str;
        }

        public TextBox getFieldPartNumber()
        {
            return this.txbCDAPartNumber;
        }

        public TextBox getFieldWorkorder()
        {
            return this.txbCDAMONumber;
        }

        public Label getFieldLabelUser()
        {
            return lblUser;
        }

        public Label getFieldLabelTime()
        {
            return lblLoginTime;
        }
        #endregion

        #region Data process function
        public void DataRecivedHeandler(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = (SerialPort)sender;
            try
            {
                if (isFormOutPoump)//登出状态不做操作
                {
                    return;
                }
                if (!VerifyCheckList())
                {
                    return;
                }
                Thread.Sleep(200);
                Byte[] bt = new Byte[sp.BytesToRead];
                sp.Read(bt, 0, sp.BytesToRead);
                indata = System.Text.Encoding.ASCII.GetString(bt).Trim();
                //string abs = CodeConversionManager.StringToHexString(indata, Encoding.UTF8, ' ');
                //indata = sp.ReadLine();
                //indata = sp.ReadExisting();
                LogHelper.Info("Scan number(original): " + indata);
                indata = indata.Replace("?", "").Replace("K", "").Replace("X", "").Replace("H", "").Replace("Q", "").Replace("T1", "").Replace("D1", "").Trim();
                if (indata.Length <= 2)
                {
                    initModel.scannerHandler.handler().DiscardInBuffer();
                    return;
                }

                this.Invoke(new MethodInvoker(delegate
                {
                    this.txbCDADataInput.Text = indata;
                }));
                if (indata.TrimEnd() == config.NoRead)
                {
                    if (config.DataOutputInterface != null && config.DataOutputInterface != "")
                        SendSN(config.LIGHT_CHANNEL_ON);
                    initModel.scannerHandler.sendLowExt();
                    errorHandler(2, Message("msg_NO READ"), "");
                    initModel.scannerHandler.handler().DiscardInBuffer();
                    return;
                }
                //match material bin number
                Match match = Regex.Match(indata, config.MBNExtractPattern);
                if (match.Success)
                {
                    this.Invoke(new MethodInvoker(delegate
                    {
                        this.tabControl1.SelectedTab = this.tabSetup;
                        SetTipMessage(MessageType.OK, Message("msg_Scan material bin number"));
                    }));
                    if (config.AutoNextMaterial.ToUpper() == "ENABLE")
                        ProcessMaterialBinNo(match.ToString());
                    else
                        ProcessMaterialBinNoEXT(match.ToString());
                    return;
                }
                //match equipment
                match = Regex.Match(indata, config.EquipmentExtractPattern);
                if (match.Success)
                {
                    this.Invoke(new MethodInvoker(delegate
                    {
                        this.tabControl1.SelectedTab = this.tabEquipment;
                        SetTipMessage(MessageType.OK, Message("msg_Scan equipment number"));
                    }));
                    ProcessEquipmentDataEXT(match.ToString());
                    return;
                }
                //match serial number
                match = Regex.Match(indata, config.DLExtractPattern);
                if (match.Success)
                {
                    string serialNumber = match.ToString();
                    if (match.Groups.Count > 1)
                        serialNumber = match.Groups[1].ToString();
                    if (!CheckEquipmentSetup() || !CheckMaterialSetUp())
                    {
                        return;
                    }
                    if (ProcessSerialNumberEXT(serialNumber))
                    {
                        if (!config.IPI_WORKORDERTYPE_CHECK.Contains(workorderType))
                        {
                            if (config.IPI_STATUS_CHECK == "ENABLE" && IPIStatus == 0)
                            {
                                IPITimerStart();
                                this.Invoke(new MethodInvoker(delegate
                                {
                                    IPIFormPoup();
                                }));
                            }
                        }
                    }
                    return;
                }
                else
                {
                    errorHandler(2, "条码不匹配", "");
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message + ";" + ex.StackTrace);
            }
            finally
            {
                initModel.scannerHandler.handler().DiscardInBuffer();
            }
        }
        #endregion

        #region Event
        private void MainView_Load(object sender, EventArgs e)
        {
            string filePath = Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName;
            string _appDir = Path.GetDirectoryName(filePath) + @"\pic\Chart_Column_Silver.png";
            NetworkChange.NetworkAvailabilityChanged += AvailabilityChanged;
        }

        private void MainView_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult dr = MessageBox.Show(Message("msg_Do you want to close the application"), Message("msg_Quit Application"), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
            if (dr == DialogResult.OK)
            {
                if (this.txbCDAMONumber.Text != "")
                {
                    SaveEquAndMaterial();
                    SaveCheckList();
                    EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
                    foreach (DataGridViewRow row in dgvEquipment.Rows)
                    {
                        string equipmentNo = row.Cells["EquipNo"].Value.ToString();
                        string equipmentIndex = row.Cells["EquipmentIndex"].Value.ToString();
                        if (string.IsNullOrEmpty(equipmentNo))
                            continue;
                        int errorCode = eqManager.UpdateEquipmentData(equipmentIndex, equipmentNo, 1);
                        RemoveAttributeForEquipment(equipmentNo, equipmentIndex, "attribEquipmentHasRigged");
                    }

                    SetUpManager setupHandler = new SetUpManager(sessionContext, initModel, this);
                    setupHandler.SetupStateChange(this.txbCDAMONumber.Text, 2);
                }
                LogHelper.Info("Application end...");
                System.Environment.Exit(0);
            }
            else
            {
                e.Cancel = true;
            }
        }

        private void txbCDADataInput_KeyUp(object sender, KeyEventArgs e)
        {
          if (e.KeyData == Keys.Enter)
            {
                if (!VerifyCheckList())
                {
                    this.txbCDADataInput.SelectAll();
                    this.txbCDADataInput.Focus();
                    //errorHandler(3, Message("$msg_checklist_first"), "");
                    return;
                }
                indata = this.txbCDADataInput.Text.Trim();
                //match material bin number
                Match match = Regex.Match(indata, config.MBNExtractPattern);
                if (match.Success)
                {
                    this.Invoke(new MethodInvoker(delegate
                    {
                        this.tabControl1.SelectedTab = this.tabSetup;
                        SetTipMessage(MessageType.OK, Message("msg_Scan material bin number"));
                    }));
                    if (config.AutoNextMaterial.ToUpper() == "ENABLE")
                        ProcessMaterialBinNo(match.ToString());
                    else
                        ProcessMaterialBinNoEXT(match.ToString());
                    this.txbCDADataInput.SelectAll();
                    this.txbCDADataInput.Focus();
                    return;
                }
                //match equipment
                match = Regex.Match(indata, config.EquipmentExtractPattern);
                if (match.Success)
                {
                    this.Invoke(new MethodInvoker(delegate
                    {
                        this.tabControl1.SelectedTab = this.tabEquipment;
                        SetTipMessage(MessageType.OK, Message("msg_Scan equipment number"));
                    }));
                    ProcessEquipmentDataEXT(match.ToString());
                    this.txbCDADataInput.SelectAll();
                    this.txbCDADataInput.Focus();
                    return;
                }
                //match serial number
                match = Regex.Match(indata, config.DLExtractPattern);
                if (match.Success)
                {
                    if (!CheckEquipmentSetup() || !CheckMaterialSetUp())
                    {
                        this.txbCDADataInput.SelectAll();
                        this.txbCDADataInput.Focus();
                        return;
                    }
                    this.Invoke(new MethodInvoker(delegate
                    {
                        //this.tabControl1.SelectedTab = this.tabShipping;
                        SetTipMessage(MessageType.OK, Message("msg_Scan serial number"));
                    }));
                    if (ProcessSerialNumberEXT(match.ToString()))
                    {
                        if (!config.IPI_WORKORDERTYPE_CHECK.Contains(workorderType))
                        {
                            if (config.IPI_STATUS_CHECK == "ENABLE" && IPIStatus != 1)
                            {
                                IPITimerStart();
                                this.Invoke(new MethodInvoker(delegate
                                {
                                    IPIFormPoup();
                                }));
                            }
                        }
                    }
                    this.txbCDADataInput.SelectAll();
                    this.txbCDADataInput.Focus();
                    return;
                }
                this.txbCDADataInput.SelectAll();
                this.txbCDADataInput.Focus();
                errorHandler(3, Message("msg_wrong barcode"), "");
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.tabControl1.SelectedTab.Name == "tabSetup")
            {
                this.gridSetup.ClearSelection();
            }
            else if (this.tabControl1.SelectedTab.Name == "tabEquipment")
            {
                this.dgvEquipment.ClearSelection();
            }
            else if (this.tabControl1.SelectedTab.Name == "tabDocument")
            {
                this.gridDocument.ClearSelection();
            }
            else if (this.tabControl1.SelectedTab.Name == "tabSNLog")
            {
                this.gridSNLog.ClearSelection();
            }
            else if (this.tabControl1.SelectedTab.Name == "tabWOActived")
            {
                this.gridWorkorder.ClearSelection();
                SetWorkorderGridStatus();
            }
        }

        private void gridDocument_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1)
                return;
            long documentID = Convert.ToInt64(gridDocument.Rows[e.RowIndex].Cells[0].Value.ToString());
            string fileName = gridDocument.Rows[e.RowIndex].Cells[1].Value.ToString();
            SetDocumentControlForDoc(documentID, fileName);

        }

        private void btnRemove_Click(object sender, EventArgs e)
        {
            if (this.dgvEquipment.SelectedRows.Count > 0)
            {
                DataGridViewRow row = this.dgvEquipment.SelectedRows[0];
                string equipmentNo = row.Cells["EquipNo"].Value.ToString();
                string equipmentIndex = row.Cells["EquipmentIndex"].Value.ToString();
                row.Cells["NextMaintenance"].Value = "";
                row.Cells["UsCount"].Value = "";
                row.Cells["EquipNo"].Value = "";
                row.Cells["EquipmentIndex"].Value = "";
                row.Cells["ScanTime"].Value = "";
                row.Cells["Status"].Value = ScreenPrinter.Properties.Resources.Close;
                row.Cells["eqPartNumber"].Style.BackColor = Color.FromArgb(255, 255, 255);

                //Strip down equipment
                if (string.IsNullOrEmpty(equipmentNo))
                    return;
                EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
                int errorCode = eqManager.UpdateEquipmentData(equipmentIndex, equipmentNo, 1);
                //remove attribute "attribEquipmentHasRigged"
                RemoveAttributeForEquipment(equipmentNo, equipmentIndex, "attribEquipmentHasRigged");
                this.dgvEquipment.ClearSelection();
                SaveEquAndMaterial();
            }
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            GetCurrentWorkorder currentWorkorder = new GetCurrentWorkorder(sessionContext, initModel, this);
            initModel.currentSettings = currentWorkorder.GetCurrentWorkorderResultCall();
            if (initModel.currentSettings != null && initModel.currentSettings.workorderNumber != this.txbCDAMONumber.Text)
            {
                this.gridSetup.Rows.Clear();
                //this.dgvEquipment.Rows.Clear();
                this.txbCDAMONumber.Text = initModel.currentSettings.workorderNumber;
                this.txbCDAPartNumber.Text = initModel.currentSettings.partNumber;
                LoadYield();
                InitSetupGrid();
                //InitEquipmentGrid();
                ShowTopWindow();
                SetTipMessage(MessageType.OK, Message("msg_Refresh Success"));
                this.txbCDADataInput.Focus();
            }
            if (initModel.currentSettings == null)
            {
                this.gridSetup.Rows.Clear();
                //this.dgvEquipment.Rows.Clear();
                this.txbCDAMONumber.Text = "";
                this.txbCDAPartNumber.Text = "";
            }
        }

        int iIndexItem = -1;
        private void dgvEquipment_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (this.dgvEquipment.Rows.Count == 0)
                    return;
                this.dgvEquipment.ContextMenuStrip = contextMenuStrip1;
                iIndexItem = ((DataGridView)sender).CurrentRow.Index;
            }
        }

        private void removeEquipment_Click(object sender, EventArgs e)
        {
            if (iIndexItem > -1)
            {
                DataGridViewRow row = this.dgvEquipment.Rows[iIndexItem];
                string equipmentNo = row.Cells["EquipNo"].Value.ToString();
                string equipmentIndex = row.Cells["EquipmentIndex"].Value.ToString();
                row.Cells["NextMaintenance"].Value = "";
                row.Cells["ScanTime"].Value = "";
                row.Cells["UsCount"].Value = "";
                row.Cells["EquipNo"].Value = "";
                row.Cells["ScanTime"].Value = "";
                row.Cells["Status"].Value = ScreenPrinter.Properties.Resources.Close;
                row.Cells["eqPartNumber"].Style.BackColor = Color.FromArgb(255, 255, 255);

                //Strip down equipment
                if (string.IsNullOrEmpty(equipmentNo))
                    return;
                EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
                int errorCode = eqManager.UpdateEquipmentData(equipmentIndex, equipmentNo, 1);
                RemoveAttributeForEquipment(equipmentNo, equipmentIndex, "attribEquipmentHasRigged");
                this.dgvEquipment.ClearSelection();
                SaveEquAndMaterial();
            }
        }

        private void btnPassBoard_Click(object sender, EventArgs e)
        {
            PassBoard();
        }
        #endregion

        #region Listen file
        string SqueegeeForce = "";
        string SqueegeeBaseSpeed = "";
        private void InitSqueegeeValue()
        {
            string filePath = config.LogFileFolder + @"\Camx.log";
            if (!File.Exists(filePath))
            {
                errorHandler(2, Message("msg_Camx.log file not exist"), "");
            }
            else
            {
                string valueRegex = @">(.*)<";
                string[] lines = File.ReadAllLines(filePath);
                foreach (var line in lines)
                {
                    if (line.Contains("<squeegee_force>"))
                    {
                        Match match = Regex.Match(line.Trim(), valueRegex);
                        if (match.Success)
                        {
                            if (match.Groups.Count > 1)
                                SqueegeeForce = match.Groups[1].ToString();
                        }
                    }
                    if (line.Contains("<squeegee_base_speed>"))
                    {
                        Match match = Regex.Match(line.Trim(), valueRegex);
                        if (match.Success)
                        {
                            if (match.Groups.Count > 1)
                                SqueegeeBaseSpeed = match.Groups[1].ToString();
                            break;
                        }
                    }
                }
            }
        }

        public void ListenFile(string filePath)
        {
            LogHelper.Info("start analysis camx.log");
            if (!File.Exists(filePath))
            {
                errorHandler(0, Message("msg_Process logfile ") + filePath + Message("msg_SUCCESS"), "");
                return;
            }
            string valueRegex = @">(.*)<";
            string[] lines = File.ReadAllLines(filePath);
            foreach (var line in lines)
            {
                if (line.Contains("<squeegee_force>"))
                {
                    Match match = Regex.Match(line.Trim(), valueRegex);
                    if (match.Success)
                    {
                        if (match.Groups.Count > 1)
                            SqueegeeForce = match.Groups[1].ToString();
                    }
                }
                if (line.Contains("<squeegee_base_speed>"))
                {
                    Match match = Regex.Match(line.Trim(), valueRegex);
                    if (match.Success)
                    {
                        if (match.Groups.Count > 1)
                            SqueegeeBaseSpeed = match.Groups[1].ToString();
                        break;
                    }
                }
            }
            LogHelper.Info("squeegee_force: " + SqueegeeForce);
            LogHelper.Info("squeegee_base_speed: " + SqueegeeBaseSpeed);
            LogHelper.Info("end analysis camx.log");
        }

        private void MoveFileToOKFolder(string filepath)
        {
            string OkFolder = config.LogTransOK;
            string strDir = Path.GetDirectoryName(filepath) + @"\";
            string strDirCopy = Path.GetDirectoryName(filepath);
            string strDestDir = "";
            try
            {
                if (strDir == config.LogFileFolder)//move file to ok folder
                {
                    FileInfo fInfo = new FileInfo(@"" + filepath);
                    string fileNameOnly = Path.GetFileNameWithoutExtension(filepath);
                    string extension = Path.GetExtension(filepath);
                    string newFullPath = null;
                    if (config.ChangeFileName.ToUpper() == "ENABLE")
                    {
                        newFullPath = Path.Combine(OkFolder, fileNameOnly + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + extension);
                    }
                    else
                    {
                        newFullPath = Path.Combine(OkFolder, fileNameOnly + extension);
                    }
                    if (!Directory.Exists(OkFolder)) Directory.CreateDirectory(OkFolder);
                    if (File.Exists(newFullPath))
                    {
                        File.Delete(newFullPath);
                    }

                    fInfo.MoveTo(@"" + newFullPath);
                }
                else//move Directory to ok folder
                {
                    string strDirName = strDirCopy.Substring(strDirCopy.LastIndexOf(@"\") + 1);
                    strDestDir = config.LogTransOK + strDirName + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                    if (!Directory.Exists(OkFolder)) Directory.CreateDirectory(OkFolder);
                    if (Directory.Exists(strDestDir))
                    {
                        Directory.Delete(strDestDir, true);
                    }
                    Directory.Move(strDir, strDestDir);
                }
                errorHandler(1, Message("msg_Move file:") + filepath + Message(" msg_to OK folder success."), "");
            }
            catch (Exception e)
            {
                errorHandler(2, Message("msg_move file error ") + e.Message, "");
            }
        }

        private void MoveFileToErrorFolder(string filepath, string errorMsg)
        {
            string errorFolder = config.LogTransError;
            string strDir = Path.GetDirectoryName(filepath) + @"\";
            string strDirCopy = Path.GetDirectoryName(filepath);
            string strDestDir = "";
            try
            {
                if (strDir == config.LogFileFolder)//move file to error folder
                {
                    FileInfo fInfo = new FileInfo(@"" + filepath);
                    string fileNameOnly = Path.GetFileNameWithoutExtension(filepath);
                    string extension = Path.GetExtension(filepath);
                    string newFullPath = null;
                    if (config.ChangeFileName.ToUpper() == "ENABLE")
                    {
                        newFullPath = Path.Combine(errorFolder, fileNameOnly + "_" + errorMsg + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + extension);
                    }
                    else
                    {
                        newFullPath = Path.Combine(errorFolder, fileNameOnly + extension);
                    }
                    if (!Directory.Exists(errorFolder)) Directory.CreateDirectory(errorFolder);
                    if (File.Exists(newFullPath))
                    {
                        File.Delete(newFullPath);
                    }
                    fInfo.MoveTo(@"" + newFullPath);
                }
                else//move Directory to error folder
                {
                    string strDirName = strDirCopy.Substring(strDirCopy.LastIndexOf(@"\") + 1);
                    strDestDir = errorFolder + strDirName + "_" + errorMsg + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                    if (!Directory.Exists(errorFolder)) Directory.CreateDirectory(errorFolder);
                    if (Directory.Exists(strDestDir))
                    {
                        Directory.Delete(strDestDir, true);
                    }
                    Directory.Move(strDir, strDestDir);
                }
                errorHandler(1, Message("msg_Move file:") + filepath + Message("msg_to OK folder success."), "");
            }
            catch (Exception e)
            {
                errorHandler(2, Message("msg_move file error ") + e.Message, "");
            }
        }
        #endregion

        #region Other functions
        private void LoadYield()
        {
            GetProductQuantity getProductHandler = new GetProductQuantity(sessionContext, initModel, this);
            if (!string.IsNullOrEmpty(this.txbCDAMONumber.Text))
            {
                ProductEntity entity = getProductHandler.GetProductQty(Convert.ToInt32(initModel.currentSettings.processLayer), this.txbCDAMONumber.Text);
                if (entity != null)
                {
                    int totalQty = Convert.ToInt32(entity.QUANTITY_PASS) + Convert.ToInt32(entity.QUANTITY_FAIL) + Convert.ToInt32(entity.QUANTITY_SCRAP);
                    this.lblPass.Text = entity.QUANTITY_PASS;
                    this.lblFail.Text = entity.QUANTITY_FAIL;
                    this.lblScrap.Text = entity.QUANTITY_SCRAP;
                    this.lblYield.Text = "0%";
                    //this.lblAllCount.Text = totalQty + "";
                    if (totalQty > 0)
                    {
                        this.lblYield.Text = Math.Round(Convert.ToDecimal(lblPass.Text) / Convert.ToDecimal(totalQty) * 100, 2) + "%";
                    }
                }
            }
        }

        private void ShowTopWindow()
        {
            if (topmostform == null)
            {
                topmostform = new TopMostForm(this);
                topmostHandle = new HandleInterfaceUpdateTopMostDelegate(topmostform.UpdateData);
                topmostform.Show();
            }
        }

        private void SetTopWindowMessage(string text, string errorMsg)
        {
            if (topmostform != null)
            {
                this.Invoke(topmostHandle, new string[] { text, errorMsg });
            }
            else
            {
                topmostform = new TopMostForm(this);
                topmostHandle = new HandleInterfaceUpdateTopMostDelegate(topmostform.UpdateData);
                topmostform.Show();
                this.Invoke(topmostHandle, new string[] { text, errorMsg });
            }
        }

        private void StrippedEquipmentFromStation()
        {
            EquipmentManager equipmentHandler = new EquipmentManager(sessionContext, initModel, this);
            List<EquipmentEntityExt> entityList = equipmentHandler.GetSetupEquipmentDataByStation(config.StationNumber);
            if (entityList != null && entityList.Count > 0)
            {
                foreach (var item in entityList)
                {
                    equipmentHandler.UpdateEquipmentData(item.EQUIPMENT_INDEX, item.EQUIPMENT_NUMBER, 1);
                    RemoveAttributeForEquipment(item.EQUIPMENT_NUMBER, item.EQUIPMENT_INDEX, "attribEquipmentHasRigged");
                }
            }
        }

        Dictionary<string, string> dicAttris = new Dictionary<string, string>();
        private void GetWorkOrderAttris()
        {
            if (string.IsNullOrEmpty(this.txbCDAMONumber.Text))
                return;
            GetAttributeValue getAttriHandler = new GetAttributeValue(sessionContext, initModel, this);
            dicAttris = getAttriHandler.GetAllAttributeValuesForWO(this.txbCDAMONumber.Text);
        }

        private void AddDataToSNGrid(object[] values)
        {
            this.Invoke(new MethodInvoker(delegate
            {
                this.gridSNLog.Rows.Insert(0, values);
                if (this.gridSNLog.Rows.Count > 100)
                {
                    this.gridSNLog.Rows.RemoveAt(100);
                }
                this.gridSNLog.ClearSelection();
            }));
        }

        private void InitSetupGrid()
        {
            this.gridSetup.Rows.Clear();
            GetMaterialBinData getMaterial = new GetMaterialBinData(sessionContext, initModel, this);
            DataTable dt = getMaterial.GetBomMaterialData(this.txbCDAMONumber.Text);
            if (dt != null && dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    this.gridSetup.Rows.Add(new object[10] { ScreenPrinter.Properties.Resources.Close, "", row["PartNumber"], row["PartDesc"], "", "", row["CompName"], row["Quantity"], "", "" });
                }
                this.gridSetup.ClearSelection();
            }
        }
        //verify the serial number's part number is equals the current part number
        private bool VerifySerialNumber(string serialNumber)
        {
            bool isValid = true;
            GetSerialNumberInfo getSNInfoHandler = new GetSerialNumberInfo(sessionContext, initModel, this);
            string[] snValues = getSNInfoHandler.GetSNInfo(serialNumber);//"PART_DESC", "PART_NUMBER", "WORKORDER_NUMBER" 
            if (snValues != null && snValues.Length > 0)
            {
                string snPartNumber = snValues[1];
                if (snPartNumber == this.txbCDAPartNumber.Text.Trim())
                { }
                else
                {
                    errorHandler(3, Message("msg_The serial number's part number is not equals the current part number"), "");
                    isValid = false;
                }
            }
            else
            {
                errorHandler(3, Message("msg_The serial number: ") + serialNumber + Message("msg_ is invalid"), "");
                SetTopWindowMessage(serialNumber, "The serial number is invalid.");
                isValid = false;
            }
            return isValid;
        }
        string snWorkOrder = "";
        private bool VerifySerialNumberByWo(string serialNumber)
        {
            bool isValid = true;
            GetSerialNumberInfo getSNInfoHandler = new GetSerialNumberInfo(sessionContext, initModel, this);
            string[] snValues = getSNInfoHandler.GetSNInfo(serialNumber);//"PART_DESC", "PART_NUMBER", "WORKORDER_NUMBER" 
            if (snValues != null && snValues.Length > 0)
            {
                snWorkOrder = snValues[2];

                if (snWorkOrder == this.txbCDAMONumber.Text.Trim())
                { }
                else
                {
                    if (config.Auto_Work_Order_Change == "ENABLE")
                        errorHandler(3, Message("msg_WO_NotMatch"), "");
                    isValid = false;
                }
            }
            else
            {
                errorHandler(3, Message("msg_The serial number: ") + serialNumber + Message("msg_ is invalid"), "");
                SetTopWindowMessage(serialNumber, "The serial number is invalid.");
                isValid = false;
            }
            return isValid;
        }

        private void PassBoard()
        {
            initModel.scannerHandler.sendHighExt();
            Thread.Sleep(Convert.ToInt32(config.GateKeeperTimer));
            initModel.scannerHandler.sendLowExt();
            //initModel.scannerHandler.sendHigh();
        }

        private void ProcessMaterialBinNo(string materialBinNo)
        {
            GetMaterialBinData getMaterialHandler = new GetMaterialBinData(sessionContext, initModel, this);
            AppendAttribute appendAttriHandler = new AppendAttribute(sessionContext, initModel, this);
            string[] values = getMaterialHandler.GetMaterialBinDataDetails(materialBinNo);
            //"MATERIAL_BIN_NUMBER", "MATERIAL_BIN_PART_NUMBER", "MATERIAL_BIN_QTY_ACTUAL", "MATERIAL_BIN_QTY_TOTAL", "PART_DESC", "MSL_FLOOR_LIFETIME_REMAIN" ,"EXPIRATION_DATE"
            if (values != null && values.Length > 0)
            {
                string strPartNumber = values[1];
                string strActualQty = values[2];
                string strLifeTime = values[5];
                string lockState = values[7];
                DateTime dtExpiry = Convert.ToDateTime(ConvertDateFromStamp(values[6]));

                if (lockState == "-1")
                {
                    errorHandler(2, Message("msg_TheContainerIsLocked"), "");
                    return;
                }

                if (!VerifyActivatedWO() || !VerifyMaterialBinData(materialBinNo, strPartNumber) || !VerifyMaterialBinData24And48(materialBinNo))
                    return;

                ////judge whether this material bin is expiry
                if (dtExpiry < DateTime.Now)
                {
                    errorHandler(2, Message("msg_The solder paste has expiry."), "");
                    return;
                }
                //get attribute value FloorLifeExpiry
                GetAttributeValue getAttriHandler = new GetAttributeValue(sessionContext, initModel, this);
                //string[] valuesAttri = getAttriHandler.GetAttributeValueForContainer("FLOOR_LIFE_EXPIRY", materialBinNo);
                //if (valuesAttri.Length > 0)
                //{
                //    dtExpiry = Convert.ToDateTime(valuesAttri[1]);
                //}
                ////judge whether this material bin is FloorLifeExpiry
                //if (dtExpiry < DateTime.Now)
                //{
                //    errorHandler(2, Message("msg_The solder paste has expiry."), "");
                //    return;
                //}

                string[] valuesAttriOpenTime = getAttriHandler.GetAttributeValueForContainer("VISCOSITY_SUCCESS", materialBinNo);
                if (valuesAttriOpenTime.Length <= 0)
                {
                    errorHandler(2, Message("msg_The solder paste has no viscosity test."), "");
                    return;
                }
                bool isMatchPN = false;
                foreach (DataGridViewRow row in gridSetup.Rows)
                {
                    if (row.Cells["PartNumber"].Value.ToString() == strPartNumber)
                    {
                        if (row.Cells["MaterialBinNo"].Value == null || row.Cells["MaterialBinNo"].Value.ToString() == "")
                        {
                            //setup material
                            SetUpManager setupHandler = new SetUpManager(sessionContext, initModel, this);
                            setupHandler.UpdateMaterialSetUpByBin(initModel.currentSettings.processLayer, this.txbCDAMONumber.Text, materialBinNo, strActualQty, strPartNumber, config.StationNumber + "_01", "01");//todo
                            setupHandler.SetupStateChange(this.txbCDAMONumber.Text, 0);

                            //add attribute, update material bin EXPIRATION_DATE
                            //if (valuesAttri == null || valuesAttri.Length == 0)
                            //{
                            int error = appendAttriHandler.AppendAttributeValuesForContainer("USED_FLAG", "Y", materialBinNo);
                            if (error == 0)
                            {
                                //DateTime dtExpiration = DateTime.Now.AddHours(Convert.ToDouble(config.SolderPasteValidity));
                                //getMaterialHandler.ChangeMaterialBinData(materialBinNo, ConvertDateToStamp(dtExpiration));
                                //appendAttriHandler.AppendAttributeValuesForContainer("FloorLifeExpiry", dtExpiration.ToString("yyyy/MM/dd HH:mm:ss"), materialBinNo);
                                //dtExpiry = dtExpiration;
                            }
                            //}

                            row.Cells["MaterialBinNo"].Value = materialBinNo;
                            row.Cells["Qty"].Value = Convert.ToDouble(strActualQty);
                            row.Cells["MScanTime"].Value = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                            row.Cells["ExpiryTime"].Value = dtExpiry.ToString("yyyy/MM/dd HH:mm:ss");
                            row.Cells["MSL"].Value = ConverToHourAndMin(Convert.ToInt32("100"));
                            row.Cells["Checked"].Value = ScreenPrinter.Properties.Resources.ok;
                            row.Cells["MaterialBinNo"].Style.BackColor = Color.FromArgb(0, 192, 0);
                            SetTipMessage(MessageType.OK, Message("msg_Process material bin number") + materialBinNo + Message("SUCCESS"));
                            isMatchPN = true;
                            SaveEquAndMaterial();
                            if (CheckMaterialSetUp() && CheckEquipmentSetup())
                            {
                                if (config.DataOutputInterface != null && config.DataOutputInterface != "")
                                    SendSN(config.LIGHT_CHANNEL_OFF);
                            }
                            break;
                        }
                        else
                        {
                            if (CheckMaterialBinHasSetup(materialBinNo))
                                return;
                            //add attribute, update material bin EXPIRATION_DATE
                            //if (valuesAttri == null || valuesAttri.Length == 0)
                            //{
                            int error = appendAttriHandler.AppendAttributeValuesForContainer("USED_FLAG", "Y", materialBinNo);
                            if (error == 0)
                            {
                                //DateTime dtExpiration = DateTime.Now.AddHours(Convert.ToDouble(config.SolderPasteValidity));
                                //getMaterialHandler.ChangeMaterialBinData(materialBinNo, ConvertDateToStamp(dtExpiration));
                                //appendAttriHandler.AppendAttributeValuesForContainer("FloorLifeExpiry", dtExpiration.ToString("yyyy/MM/dd HH:mm:ss"), materialBinNo);
                                //dtExpiry = dtExpiration;
                            }
                            //}
                            this.Invoke(new MethodInvoker(delegate
                            {
                                gridSetup.Rows.Add();
                                DataGridViewRow newRow = gridSetup.Rows[gridSetup.Rows.Count - 1];
                                newRow.Cells["Checked"].Value = ScreenPrinter.Properties.Resources.ok;
                                newRow.Cells["MaterialBinNo"].Value = materialBinNo;
                                newRow.Cells["PartNumber"].Value = row.Cells["PartNumber"].Value;
                                newRow.Cells["PartDesc"].Value = row.Cells["PartDesc"].Value;
                                newRow.Cells["Qty"].Value = Convert.ToDouble(strActualQty);
                                newRow.Cells["MSL"].Value = ConverToHourAndMin(Convert.ToInt32("100"));
                                newRow.Cells["MaterialBinNo"].Style.BackColor = Color.FromArgb(0, 192, 0);
                                newRow.Cells["MScanTime"].Value = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                                newRow.Cells["ExpiryTime"].Value = dtExpiry.ToString("yyyy/MM/dd HH:mm:ss");
                                newRow.Cells["PPNQty"].Value = row.Cells["PPNQty"].Value.ToString();
                            }));
                            isMatchPN = true;
                            break;
                        }
                    }
                }

                if (isMatchPN == false)
                    SetTipMessage(MessageType.Error, Message("msg_MBN not match"));
            }
        }

        private void ProcessMaterialBinNoEXT(string materialBinNo)
        {
            GetMaterialBinData getMaterialHandler = new GetMaterialBinData(sessionContext, initModel, this);
            AppendAttribute appendAttriHandler = new AppendAttribute(sessionContext, initModel, this);
            ProcessMaterialBinData materialHandler = new ProcessMaterialBinData(sessionContext, initModel, this);
            string[] values = getMaterialHandler.GetMaterialBinDataDetails(materialBinNo);
            //"MATERIAL_BIN_NUMBER", "MATERIAL_BIN_PART_NUMBER", "MATERIAL_BIN_QTY_ACTUAL", "MATERIAL_BIN_QTY_TOTAL", "PART_DESC", "MSL_FLOOR_LIFETIME_REMAIN" ,"EXPIRATION_DATE"
            if (values != null && values.Length > 0)
            {
                GetAttributeValue getAttriHandler = new GetAttributeValue(sessionContext, initModel, this);
                //判断锡膏是否为空Use_Empty 20160831
                string[] valuesEmptyAttri = getAttriHandler.GetAttributeValueForContainer("Use_Empty", materialBinNo);
                if (valuesEmptyAttri != null && valuesEmptyAttri.Length > 0)
                {
                    if (valuesEmptyAttri[1] == "Y")
                    {
                        errorHandler(2, Message("msg_The solder paste has empty please select another solder"), "");
                        return;
                    }
                }
                else
                {
                    appendAttriHandler.AppendAttributeValuesForContainer("Use_Empty", "N", materialBinNo);
                }

                string strPartNumber = values[1];
                string strActualQty = values[2];
                string strLifeTime = values[5];
                string lockState = values[7];
                DateTime dtExpiry = Convert.ToDateTime(ConvertDateFromStamp(values[6]));

                if (lockState == "-1")
                {
                    errorHandler(2, Message("msg_TheContainerIsLocked"), "");
                    return;
                }

                if (!VerifyActivatedWO() || !VerifyMaterialBinData(materialBinNo, strPartNumber) || !VerifyMaterialBinData24And48(materialBinNo))
                    return;

                ////judge whether this material bin is expiry
                if (dtExpiry < DateTime.Now)
                {
                    errorHandler(2, Message("msg_The solder paste has expiry."), "");
                    return;
                }
                //get attribute value FloorLifeExpiry 8小时
                //string[] valuesAttri = getAttriHandler.GetAttributeValueForContainer("FLOOR_LIFE_EXPIRY", materialBinNo);
                //if (valuesAttri.Length > 0)
                //{
                //    dtExpiry = Convert.ToDateTime(valuesAttri[1]);
                //}
                ////judge whether this material bin is FloorLifeExpiry
                //if (dtExpiry < DateTime.Now)
                //{
                //    errorHandler(2, Message("msg_The solder paste has expiry."), "");
                //    return;
                //}
                //粘度测试
                string[] valuesAttriOpenTime = getAttriHandler.GetAttributeValueForContainer("VISCOSITY_SUCCESS", materialBinNo);
                if (valuesAttriOpenTime.Length <= 0)
                {
                    errorHandler(2, Message("msg_The solder paste has no viscosity test."), "");
                    return;
                }
                bool isMatchPN = false;
                foreach (DataGridViewRow row in gridSetup.Rows)
                {
                    if (row.Cells["PartNumber"].Value.ToString() == strPartNumber)
                    {
                        //if (row.Cells["MaterialBinNo"].Value == null || row.Cells["MaterialBinNo"].Value.ToString() == "")
                        //{
                        //在上另一罐锡膏之前弹出框提示是否前一罐已用完
                        if (row.Cells["MaterialBinNo"].Value.ToString() != "" && row.Cells["MaterialBinNo"].Value.ToString() != materialBinNo)
                        {
                            DialogResult dr = MessageBox.Show(Message("msg_The previous solder paste is run out"), Message("msg_Warning"), MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
                            if (dr == DialogResult.Yes)
                            {
                                appendAttriHandler.AppendAttributeValuesForContainer("Use_Empty", "Y", row.Cells["MaterialBinNo"].Value.ToString());
                                int iQty = Convert.ToInt32(row.Cells[4].Value.ToString());
                                materialHandler.UpdateMaterialBinBooking(row.Cells["MaterialBinNo"].Value.ToString(), this.txbCDAMONumber.Text, -iQty);
                            }
                            else
                            {

                            }
                        }

                        //setup material
                        SetUpManager setupHandler = new SetUpManager(sessionContext, initModel, this);
                        setupHandler.SetupStateChange(this.txbCDAMONumber.Text, 2);

                        setupHandler.UpdateMaterialSetUpByBin(initModel.currentSettings.processLayer, this.txbCDAMONumber.Text, materialBinNo, strActualQty, strPartNumber, config.StationNumber + "_01", "01");//todo
                        setupHandler.SetupStateChange(this.txbCDAMONumber.Text, 0);

                        //add attribute, update material bin EXPIRATION_DATE
                        //if (valuesAttri == null || valuesAttri.Length == 0)
                        //{
                        int error = appendAttriHandler.AppendAttributeValuesForContainer("USED_FLAG", "Y", materialBinNo);
                        if (error == 0)
                        {
                            //DateTime dtExpiration = DateTime.Now.AddHours(Convert.ToDouble(config.SolderPasteValidity));
                            //getMaterialHandler.ChangeMaterialBinData(materialBinNo, ConvertDateToStamp(dtExpiration));
                            //appendAttriHandler.AppendAttributeValuesForContainer("FloorLifeExpiry", dtExpiration.ToString("yyyy/MM/dd HH:mm:ss"), materialBinNo);
                            //dtExpiry = dtExpiration;
                        }
                        //}

                        row.Cells["MaterialBinNo"].Value = materialBinNo;
                        row.Cells["Qty"].Value = Convert.ToDouble(strActualQty);
                        row.Cells["MScanTime"].Value = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                        row.Cells["ExpiryTime"].Value = dtExpiry.ToString("yyyy/MM/dd HH:mm:ss");
                        row.Cells["MSL"].Value = ConverToHourAndMin(Convert.ToInt32("100"));
                        row.Cells["Checked"].Value = ScreenPrinter.Properties.Resources.ok;
                        row.Cells["MaterialBinNo"].Style.BackColor = Color.FromArgb(0, 192, 0);
                        SetTipMessage(MessageType.OK, Message("msg_Process material bin number") + materialBinNo + Message("SUCCESS"));
                        isMatchPN = true;
                        SaveEquAndMaterial();
                        if (CheckMaterialSetUp() && CheckEquipmentSetup())
                        {
                            if (config.DataOutputInterface != null && config.DataOutputInterface != "")
                                SendSN(config.LIGHT_CHANNEL_OFF);
                        }
                        break;
                        //}
                        //else if (row.Cells["MaterialBinNo"].Value != null && row.Cells["MaterialBinNo"].Value.ToString() == materialBinNo)
                        //{
                        //    errorHandler(2, Message("msg_the material has been setup"), "");
                        //    isMatchPN = true;
                        //    break;
                        //}
                    }
                }

                if (isMatchPN == false)
                    SetTipMessage(MessageType.Error, Message("msg_MBN not match"));
            }
        }

        private bool ProcessSerialNumber(string serialNumber)
        {
            //verify material&equipment is ok
            if (!VerifyEquipment())
            {
                return false;
            }
            if (!VerifySerialNumberByWo(serialNumber))
            {
                return false;
            }
            if (!VerifyIPIStatus())
            {
                this.BeginInvoke(new Action(() =>
                {
                    IPIFormPoup();
                }));
                return false;
            }
            //gate keeper,check serial state
            CheckSerialNumberState checkHandler = new CheckSerialNumberState(sessionContext, initModel, this);
            bool isOK = checkHandler.CheckSNState(serialNumber);
            if (isOK)
            {
                UploadProcessResult updateHandler = new UploadProcessResult(sessionContext, initModel, this);
                int errorCode = updateHandler.UploadProcessResultCall(serialNumber);
                if (errorCode == 0)
                {
                    //AppendAttributeToSN(serialNumber);

                    SetTopWindowMessage(serialNumber, "");
                    //update material,equipment data on UI
                    UpdateGridDataAfterUploadState();

                    //set IPI Attribute
                    SetIPI(serialNumber);

                    //add data to SN Log Grid
                    AddDataToSNGrid(new object[6] { serialNumber, GetWorkOrderValue(), GetPartNumberValue(), DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "OK", "Upload state success" });
                    SetTipMessage(MessageType.OK, Message("msg_Process serial number ") + serialNumber + Message("msg_SUCCESS"));
                    this.Invoke(new MethodInvoker(delegate
                    {
                        LoadYield();
                    }));
                    if (config.OpenControlBox == "Enable")
                    {
                        PassBoard();
                    }
                    if (config.IPI_STATUS_CHECK == "ENABLE" && IPIStatus != 1)
                    {
                        IPITimerStart();
                        IPIFormPoup();
                    }
                    return true;
                }
                else
                {
                    AddDataToSNGrid(new object[6] { serialNumber, GetWorkOrderValue(), GetPartNumberValue(), DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "OK", "Upload state error" });
                    SetTopWindowMessage(serialNumber, Message("msg_Upload state error."));
                    SetTipMessage(MessageType.Error, Message("msg_Process serial number") + serialNumber + " fail.");
                    return false;
                }
            }
            else
            {
                SetTopWindowMessage(serialNumber, "Check Serial Number State Error.");
                AddDataToSNGrid(new object[6] { serialNumber, GetWorkOrderValue(), GetPartNumberValue(), DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "NG", "Check Serial Number State Error" });
                return false;
            }
        }

        string worktepnumber = "";
        private bool ProcessSerialNumberEXT(string serialNumber)
        {
            try
            {
                //verify material&equipment is ok
                if (!VerifyEquipment())
                {
                    return false;
                }
                if (!VerifyMaterialBinData24And48())
                {
                    return false;
                }
                worktepnumber = GetWorkStepNumberBySN(serialNumber);
                #region 原来流程
                if (config.Auto_Work_Order_Change == "ENABLE")//如果没有自动转换工单功能，检测到工单不一致就报错，不往下执行
                {
                    if (!VerifySerialNumberByWo(serialNumber))
                    {
                        return false;
                    }
                    if (!VerifyIPIStatusEXT(serialNumber, worktepnumber))
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            IPIFormPoup();
                        }));
                        return false;
                    }

                    //gate keeper,check serial state
                    CheckSerialNumberState checkHandler = new CheckSerialNumberState(sessionContext, initModel, this);
                    bool isOK = checkHandler.trCheckSNStateNextStep(serialNumber);
                    if (isOK)
                    {
                        if (config.DataOutputInterface != null && config.DataOutputInterface != "")
                            SendSN(config.LIGHT_CHANNEL_OFF);
                        UploadProcessResult updateHandler = new UploadProcessResult(sessionContext, initModel, this);
                        int errorCode = updateHandler.UploadProcessResultCall(serialNumber);
                        if (errorCode == 0)
                        {
                            //AppendAttributeToSN(serialNumber);

                            SetTopWindowMessage(serialNumber, "");
                            //update material,equipment data on UI
                            UpdateGridDataAfterUploadState();

                            //set IPI Attribute
                            SetIPIEXT(serialNumber, worktepnumber);

                            //add data to SN Log Grid
                            AddDataToSNGrid(new object[6] { serialNumber, GetWorkOrderValue(), GetPartNumberValue(), DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "OK", Message("msg_Upload state success") });
                            SetTipMessage(MessageType.OK, Message("msg_Process serial number ") + serialNumber + Message("msg_SUCCESS"));
                            this.Invoke(new MethodInvoker(delegate
                            {
                                LoadYield();
                            }));
                            if (config.OpenControlBox == "Enable")
                            {
                                PassBoard();
                            }
                            if (config.IsNeedProductionInspection == "ENABLE")
                            {
                                if (!config.IPI_WORKORDERTYPE_CHECK.Contains(workorderType))
                                {
                                    if (!CheckShiftChange())//验证是否为换班后的第一块板，是就给serialnumber加上属性
                                    {
                                        LogHelper.Debug("Check shift change production inspection.");
                                        if (IPIStatus == 1 || IPIStatus == 3)
                                        {
                                            IPITimerStart();
                                            //set IPI Attribute
                                            SetIPIProductionEXT(serialNumber, worktepnumber);
                                            WriteIntoShift();
                                            errorHandler(0, Message("msg_SN require to perform Production Inspection") + serialNumber, "");
                                        }
                                    }
                                    if (NextSelectSN)
                                    {
                                        LogHelper.Debug("Check next production inspection.");
                                        if (IPIStatus == 1 || IPIStatus == 3)
                                        {
                                            IPITimerStart();
                                            //set IPI Attribute
                                            SetIPIProductionEXT(serialNumber, worktepnumber);
                                            NextSelectSN = false;
                                            errorHandler(0, Message("msg_serial number selected for Inspection") + serialNumber, "");
                                        }
                                    }
                                    if (!CheckLastProduct())
                                    {
                                        LogHelper.Debug("Check last production inspection.");
                                        if (IPIStatus == 1 || IPIStatus == 3)
                                        {
                                            //set IPI Attribute
                                            SetLastIPIProductionEXT(serialNumber, worktepnumber);
                                            errorHandler(0, Message("msg_SN require to perform Last Production Inspection") + serialNumber, "");
                                        }
                                    }
                                }
                            }

                            return true;
                        }
                        else
                        {
                            AddDataToSNGrid(new object[6] { serialNumber, GetWorkOrderValue(), GetPartNumberValue(), DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "OK", Message("msg_Upload state error") });
                            SetTopWindowMessage(serialNumber, Message("msg_Upload state error."));
                            SetTipMessage(MessageType.Error, Message("msg_Process serial number") + serialNumber + Message("msg_FAIL"));
                            return false;
                        }
                    }
                    else
                    {
                        if (config.DataOutputInterface != null && config.DataOutputInterface != "")
                            SendSN(config.LIGHT_CHANNEL_ON);
                        SetTopWindowMessage(serialNumber, Message("msg_Check Serial Number State Error"));
                        AddDataToSNGrid(new object[6] { serialNumber, GetWorkOrderValue(), GetPartNumberValue(), DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "NG", Message("msg_Check Serial Number State Error") });
                        return false;
                    }
                }
                #endregion

                #region 转换工单流程
                if (config.Auto_Work_Order_Change == "DISABLE" && config.IsNeedTransWO == "Y")//"Disable", 判断是否为转换工单，是的话就执行转换功能
                {
                    bool IsTheSameWO = true;
                    if (!VerifySerialNumberByWo(serialNumber))
                    {
                        if (!VerifyTransWOExist())
                        {
                            errorHandler(2, Message("msg_Serial Number’s Work Order transfer not allow"), "");
                            return false;
                        }
                        IsTheSameWO = false;
                    }

                    if (!VerifyIPIStatusEXT(serialNumber, worktepnumber))
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            IPIFormPoup();
                        }));
                        return false;
                    }

                    //gate keeper,check serial state
                    CheckSerialNumberState checkHandler = new CheckSerialNumberState(sessionContext, initModel, this);
                    bool isOK = checkHandler.trCheckSNStateNextStep(serialNumber);
                    if (isOK)
                    {
                        if (config.DataOutputInterface != null && config.DataOutputInterface != "")
                            SendSN(config.LIGHT_CHANNEL_OFF);
                        UploadProcessResult updateHandler = new UploadProcessResult(sessionContext, initModel, this);
                        GetSerialNumberInfo getSNHandler = new GetSerialNumberInfo(sessionContext, initModel, this);
                        AssignSerialNumber assignHandler = new AssignSerialNumber(sessionContext, initModel, this);
                        int assignCode = -1;
                        #region//工单不同时才要执行以下操作
                        if (!IsTheSameWO)
                        {
                            SerialNumberData[] snArray = getSNHandler.GetSerialNumberBySNRef(serialNumber);//获取原工单所有小板和位置
                            if (snArray.Length != initModel.numberOfSingleBoards)
                            {
                                errorHandler(2, Message("msg_serial number's board is not match"), "");
                                return false;
                            }


                            string temSerialnumber = serialNumber;
                            if (serialNumber.Length < snWorkOrder.Length + 7)
                            {
                                temSerialnumber = serialNumber + "001";
                            }
                            string[] bookingResultValues = getSNHandler.GetSNHistoryInfo(temSerialnumber);//获取原来镭雕机过站时间
                            long LMbookdate = 0;
                            string LMstationnumber = "";
                            if (bookingResultValues != null && bookingResultValues.Length > 0)
                            {
                                if (bookingResultValues.Length / 2 >= 2)
                                {
                                    errorHandler(2, Message("msg_serial number has pass SP"), "");
                                    return false;
                                }
                                LMbookdate = long.Parse(bookingResultValues[0]);
                                LMstationnumber = bookingResultValues[1];
                            }

                            GetProductQuantity getProductHandler = new GetProductQuantity(sessionContext, initModel, this);
                            if (!string.IsNullOrEmpty(this.txbCDAMONumber.Text))
                            {
                                ProductEntity entity = getProductHandler.GetProductQty(Convert.ToInt32(initModel.currentSettings.processLayer), this.txbCDAMONumber.Text, LMstationnumber);
                                if (entity != null)
                                {
                                    string passQty = entity.QUANTITY_PASS;
                                    int leftQty = initModel.currentSettings.QuantityMO - Convert.ToInt32(passQty);
                                    if (leftQty < initModel.numberOfSingleBoards)//如果剩余数量不足，则不允许转工单
                                    {
                                        errorHandler(2, Message("msg_left qty is not enough"), "");
                                        return false;
                                    }
                                }
                            }
                            string datetime = DateTime.Now.ToString("yyyyMMddHHmmss");
                            getSNHandler.SwitchSerialNumber(snArray, datetime);//把原序列号Switch
                            updateHandler.ScrapProcessResultCall(serialNumber + "_" + datetime, LMstationnumber);//把原来的序列号报废

                            string sn1 = snArray[0].serialNumber;
                            string refsn = sn1.Substring(0, sn1.Length - 3);
                            //将序列号重新添加到新的工单并过镭雕站
                            if (snArray.Count() == 1)
                            {
                                assignCode = assignHandler.AssignSerialNumberResultCallForSingle(snArray, GetWorkOrderValue(), LMstationnumber);
                            }
                            else
                            {
                                assignCode = assignHandler.AssignSerialNumberResultCallForMul(refsn, snArray, GetWorkOrderValue(), LMstationnumber);
                            }

                            updateHandler.UploadPreProcessResultCall(sn1, LMstationnumber, LMbookdate);
                        }
                        #endregion
                        int errorCode = updateHandler.UploadProcessResultCall(serialNumber);
                        if (errorCode == 0)
                        {
                            //AppendAttributeToSN(serialNumber);

                            SetTopWindowMessage(serialNumber, "");
                            //update material,equipment data on UI
                            UpdateGridDataAfterUploadState();

                            //set IPI Attribute
                            SetIPIEXT(serialNumber, worktepnumber);

                            //add data to SN Log Grid
                            AddDataToSNGrid(new object[6] { serialNumber, GetWorkOrderValue(), GetPartNumberValue(), DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "OK", Message("msg_Upload state success") });
                            SetTipMessage(MessageType.OK, Message("msg_Process serial number ") + serialNumber + Message("msg_SUCCESS"));
                            this.Invoke(new MethodInvoker(delegate
                            {
                                LoadYield();
                            }));
                            if (config.OpenControlBox == "Enable")
                            {
                                PassBoard();
                            }
                            //if (config.IPI_STATUS_CHECK == "ENABLE" && IPIStatus != 1)
                            //{
                            //    IPITimerStart();
                            //    IPIFormPoup();
                            //}
                            if (config.IsNeedProductionInspection == "ENABLE")
                            {
                                if (!config.IPI_WORKORDERTYPE_CHECK.Contains(workorderType))
                                {
                                    if (!CheckShiftChange())//验证是否为换班后的第一块板，是就给serialnumber加上属性
                                    {
                                        LogHelper.Debug("Check shift change production inspection.");
                                        if (IPIStatus == 1 || IPIStatus == 3)
                                        {
                                            IPITimerStart();
                                            //set IPI Attribute
                                            SetIPIProductionEXT(serialNumber, worktepnumber);
                                            WriteIntoShift();
                                            errorHandler(0, Message("msg_SN require to perform Production Inspection") + serialNumber, "");
                                        }
                                    }
                                    if (NextSelectSN)
                                    {
                                        LogHelper.Debug("Check next production inspection.");
                                        if (IPIStatus == 1 || IPIStatus == 3)
                                        {
                                            IPITimerStart();
                                            //set IPI Attribute
                                            SetIPIProductionEXT(serialNumber, worktepnumber);
                                            NextSelectSN = false;
                                            errorHandler(0, Message("msg_serial number selected for Inspection") + serialNumber, "");
                                        }
                                    }
                                    if (!CheckLastProduct())
                                    {
                                        LogHelper.Debug("Check last production inspection.");
                                        if (IPIStatus == 1 || IPIStatus == 3)
                                        {
                                            //set IPI Attribute
                                            SetLastIPIProductionEXT(serialNumber, worktepnumber);
                                            errorHandler(0, Message("msg_SN require to perform Last Production Inspection") + serialNumber, "");
                                        }
                                    }
                                }
                            }
                            return true;
                        }
                        else
                        {
                            AddDataToSNGrid(new object[6] { serialNumber, GetWorkOrderValue(), GetPartNumberValue(), DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "OK", Message("msg_Upload state error") });
                            SetTopWindowMessage(serialNumber, Message("msg_Upload state error."));
                            SetTipMessage(MessageType.Error, Message("msg_Process serial number") + serialNumber + Message("msg_FAIL"));
                            return false;
                        }
                    }
                    else
                    {
                        if (config.DataOutputInterface != null && config.DataOutputInterface != "")
                            SendSN(config.LIGHT_CHANNEL_ON);
                        SetTopWindowMessage(serialNumber, Message("msg_Check Serial Number State Error"));
                        AddDataToSNGrid(new object[6] { serialNumber, GetWorkOrderValue(), GetPartNumberValue(), DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "NG", Message("msg_Check Serial Number State Error") });
                        return false;
                    }
                }
                #endregion

                return true;
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message, ex);
                return false;
            }
        }

        private void AppendAttributeToSN(string serialNumber)
        {
            AppendAttribute appendAttriHandler = new AppendAttribute(sessionContext, initModel, this);
            appendAttriHandler.AppendAttributeValuesForSN("Squeegee_Force_Attrib", SqueegeeForce, "", serialNumber);
            appendAttriHandler.AppendAttributeValuesForSN("Squeegee_Base_Speed_Attrib", SqueegeeBaseSpeed, "", serialNumber);
        }

        private void UpdateGridDataAfterUploadState()
        {
            if (dgvEquipment.Rows.Count > 0)
            {
                foreach (DataGridViewRow row in dgvEquipment.Rows)
                {
                    string equipmentNo = row.Cells["EquipNo"].Value.ToString();
                    Match matchStencil = Regex.Match(equipmentNo, config.StencilPrefix);
                    if (matchStencil.Success)
                    {
                        if (row.Cells["UsCount"].Value != null && row.Cells["UsCount"].Value.ToString().Length > 0)
                        {
                            int iQty = Convert.ToInt32(row.Cells["UsCount"].Value.ToString());
                            if (config.ReduceEquType == "1")
                            {
                                row.Cells["UsCount"].Value = iQty - 1;
                                int usagecount = 0;
                                GetAttributeValue getAttribHandler = new GetAttributeValue(sessionContext, initModel, this);
                                string[] valuesusage = getAttribHandler.GetAttributeValueForEquipment("USAGE_COUNT", equipmentNo, "0");
                                if (valuesusage != null && valuesusage.Length != 0)
                                {
                                    usagecount = Convert.ToInt32(valuesusage[1]);
                                }

                                AppendAttribute appendAttri = new AppendAttribute(sessionContext, initModel, this);
                                appendAttri.AppendAttributeValuesForEquipment("USAGE_COUNT", Convert.ToString(usagecount + 1), equipmentNo);
                            }
                            else
                            {
                                row.Cells["UsCount"].Value = iQty - initModel.numberOfSingleBoards;//
                            }

                        }
                    }
                    else
                    {
                        if (row.Cells["UsCount"].Value != null && row.Cells["UsCount"].Value.ToString().Length > 0)
                        {
                            int iQty = Convert.ToInt32(row.Cells["UsCount"].Value.ToString());
                            row.Cells["UsCount"].Value = iQty - initModel.numberOfSingleBoards;//
                        }
                    }
                }
            }

            //string materialBinNo = FindMaterialBinNumber();
            //if (materialBinNo != null && materialBinNo != "")
            //{
            //    Double iPPNQty = Convert.ToDouble(Convert.ToDecimal(gridSetup.Rows[0].Cells["PPNQty"].Value));//* initModel.numberOfSingleBoards;//todo
            //    LogHelper.Info("Consumption material bin number:" + materialBinNo);
            //    LogHelper.Info("Consumption quantity:" + materialBinNo);
            //    UpdateMaterialGridData(materialBinNo, iPPNQty);
            //}
        }

        private void UpdateMaterialGridData(string materialBinNumber, double qty)
        {
            ProcessMaterialBinData materialHandler = new ProcessMaterialBinData(sessionContext, initModel, this);
            for (int i = 0; i < this.gridSetup.Rows.Count; i++)
            {
                if (gridSetup.Rows[i].Cells["MaterialBinNo"].Value.ToString() == materialBinNumber)
                {
                    double iQty = Convert.ToDouble(gridSetup.Rows[i].Cells["Qty"].Value);
                    if (iQty >= qty)
                    {
                        gridSetup.Rows[i].Cells["Qty"].Value = iQty - qty;
                        int errorMaterial = materialHandler.UpdateMaterialBinBooking(materialBinNumber, this.txbCDAMONumber.Text, -qty);
                        if (iQty == qty)//update 2015/6/24
                        {
                            if (i + 1 < this.gridSetup.Rows.Count)
                            {
                                string nextMaterialBinNo = gridSetup.Rows[i + 1].Cells["MaterialBinNo"].Value.ToString();
                                string nextPartNumber = gridSetup.Rows[i + 1].Cells["PartNumber"].Value.ToString();
                                string nextQty = gridSetup.Rows[i + 1].Cells["Qty"].Value.ToString();
                                //setup material
                                SetUpManager setupHandler = new SetUpManager(sessionContext, initModel, this);
                                setupHandler.UpdateMaterialSetUpByBin(initModel.currentSettings.processLayer, this.txbCDAMONumber.Text, nextMaterialBinNo, nextQty, nextPartNumber, config.StationNumber + "_01", "01");
                                setupHandler.SetupStateChange(this.txbCDAMONumber.Text, 0);

                                gridSetup.Rows.Remove(gridSetup.Rows[i]);
                            }
                            else
                            {
                                gridSetup.Rows[i].Cells["Qty"].Value = "";
                                gridSetup.Rows[i].Cells["MaterialBinNo"].Value = "";
                                //gridSetup.Rows[i].Cells["LotNumber"].Value = "";
                                gridSetup.Rows[i].Cells["MScanTime"].Value = "";
                                gridSetup.Rows[i].Cells["ExpiryTime"].Value = "";
                                gridSetup.Rows[i].Cells["Checked"].Value = ScreenPrinter.Properties.Resources.Close;
                                gridSetup.Rows[i].Cells["MaterialBinNo"].Style.BackColor = Color.White;
                                errorHandler(3, Message("msg_The solder paste not enough."), "");
                            }
                        }
                        break;
                    }
                    else
                    {
                        gridSetup.Rows[i].Cells["Qty"].Value = 0;
                        int errorMaterial = materialHandler.UpdateMaterialBinBooking(materialBinNumber, this.txbCDAMONumber.Text, -iQty);
                        if (i + 1 < this.gridSetup.Rows.Count)
                        {
                            string nextMaterialBinNo = gridSetup.Rows[i + 1].Cells["MaterialBinNo"].Value.ToString();
                            string nextPartNumber = gridSetup.Rows[i + 1].Cells["PartNumber"].Value.ToString();
                            string nextQty = gridSetup.Rows[i + 1].Cells["Qty"].Value.ToString();
                            //setup material
                            SetUpManager setupHandler = new SetUpManager(sessionContext, initModel, this);
                            setupHandler.UpdateMaterialSetUpByBin(initModel.currentSettings.processLayer, this.txbCDAMONumber.Text, nextMaterialBinNo, nextQty, nextPartNumber, config.StationNumber + "_01", "01");
                            setupHandler.SetupStateChange(this.txbCDAMONumber.Text, 0);
                            UpdateMaterialGridData(nextMaterialBinNo, qty - iQty);
                        }
                        else
                        {
                            //warming no material
                            errorHandler(3, Message("msg_The solder paste not enough."), "");
                        }
                    }
                }
            }
        }

        private long ConvertDateToStamp(DateTime dt)
        {
            DateTime dtStart = TimeZone.CurrentTimeZone.ToLocalTime(new DateTime(1970, 1, 1));
            TimeSpan toNow = dt.Subtract(dtStart);
            return Convert.ToInt64(toNow.TotalMilliseconds);
        }

        private string FindMaterialBinNumber()
        {
            string materialBinNo = "";
            if (gridSetup.Rows.Count > 0)
            {
                for (int i = 0; i < this.gridSetup.Rows.Count; i++)
                {
                    if (Convert.ToDouble(gridSetup.Rows[i].Cells["Qty"].Value) > 0)
                    {
                        materialBinNo = gridSetup.Rows[i].Cells["MaterialBinNo"].Value.ToString();
                        break;
                    }

                }
            }
            return materialBinNo;
        }

        private bool CheckMaterialBinHasSetup(string materailBinNo)
        {
            bool isExist = false;
            foreach (DataGridViewRow row in gridSetup.Rows)
            {
                if (row.Cells["MaterialBinNo"].Value.ToString() == materailBinNo)
                {
                    isExist = true;
                    break;
                }
            }
            return isExist;
        }

        private bool VerifyMaterialBinData24And48()
        {
            bool isValid = true;

            if (config.ThawingCheck != "Enable")
            {
                return true;
            }
            foreach (DataGridViewRow item in this.dgvEquipment.Rows)
            {
                string materilBin = item.Cells["MaterialBinNo"].Value.ToString();
                //有效期24小时&48小时
                GetAttributeValue getAttriHandler = new GetAttributeValue(sessionContext, initModel, this);
                string[] values = getAttriHandler.GetAttributeValueForContainer("THAW_EXPIRY", materilBin);
                if (values != null && values.Length > 0)
                {
                    DateTime dtCompleteThawing = Convert.ToDateTime(values[1]);
                    if (DateTime.Now > dtCompleteThawing)
                    {
                        if (config.DataOutputInterface != null && config.DataOutputInterface != "")
                            SendSN(config.LIGHT_CHANNEL_ON);
                        isValid = false;
                        errorHandler(3, Message("msg_The solder paste thawing scrap."), "");
                        break;
                    }
                }
                else
                {
                    if (config.DataOutputInterface != null && config.DataOutputInterface != "")
                        SendSN(config.LIGHT_CHANNEL_ON);
                    isValid = false;
                    errorHandler(3, Message("msg_The solder paste thawing scrap."), "");
                    break;
                }
            }
            return isValid;
        }

        private bool VerifyEquipment()
        {
            bool isValid = true;
            GetAttributeValue getAttribHandler = new GetAttributeValue(sessionContext, initModel, this);
            EquipmentManager equipmentHandler = new EquipmentManager(sessionContext, initModel, this);
            int errorCode = equipmentHandler.CheckEquipmentData(this.txbCDAMONumber.Text);
            if (errorCode != 0)
            {
                if (config.DataOutputInterface != null && config.DataOutputInterface != "")
                    SendSN(config.LIGHT_CHANNEL_ON);
                //errorHandler(3, Message("msg_Check equipment data error :") + errorCode, "");
                return false;
            }
            foreach (DataGridViewRow item in this.dgvEquipment.Rows)
            {
                string equipmentNo = item.Cells["EquipNo"].Value.ToString();
                Match matchStencil = Regex.Match(equipmentNo, config.StencilPrefix);
                if (matchStencil.Success)
                {
                    string[] valuesTest = getAttribHandler.GetAttributeValueForEquipment("NEXT_TEST_DATE", equipmentNo, "0");
                    DateTime NextTestDate = DateTime.Now;
                    if (valuesTest != null && valuesTest.Length != 0)
                        NextTestDate = Convert.ToDateTime(valuesTest[1]);
                    if (Convert.ToDateTime(item.Cells["NextMaintenance"].Value) <= DateTime.Now)// || NextTestDate < DateTime.Now
                    {
                        if (config.DataOutputInterface != null && config.DataOutputInterface != "")
                            SendSN(config.LIGHT_CHANNEL_ON);
                        isValid = false;
                        item.Cells["eqPartNumber"].Style.BackColor = Color.FromArgb(255, 255, 255);
                        errorHandler(3, Message("msg_Stencil need to clean"), "");
                        //add attribute lock time
                        //add attribute to equipment
                        AppendAttribute appendAttri = new AppendAttribute(sessionContext, initModel, this);
                        appendAttri.AppendAttributeValuesForEquipment("STENCIL_LOCK_TIME", DateTime.Now.AddHours(Convert.ToDouble(config.LockTime)).ToString("yyyy/MM/dd HH:mm:ss"), equipmentNo);
                        break;
                    }
                }

                if (Convert.ToInt32(item.Cells["UsCount"].Value) <= Convert.ToInt32(config.WarningQty))
                {
                    if (config.DataOutputInterface != null && config.DataOutputInterface != "")
                        SendSN(config.LIGHT_CHANNEL_ON);
                    isValid = false;
                    item.Cells["eqPartNumber"].Style.BackColor = Color.FromArgb(255, 255, 255);
                    errorHandler(3, Message("msg_Stencil usage count cann't less then 0"), "");
                    break;
                }
                else if (Convert.ToDateTime(item.Cells["NextMaintenance"].Value) <= DateTime.Now)
                {
                    if (config.DataOutputInterface != null && config.DataOutputInterface != "")
                        SendSN(config.LIGHT_CHANNEL_ON);
                    isValid = false;
                    item.Cells["eqPartNumber"].Style.BackColor = Color.FromArgb(255, 255, 255);
                    errorHandler(3, Message("msg_equipment is expired"), "");
                    //add attribute lock time
                    //add attribute to equipment
                    //AppendAttribute appendAttri = new AppendAttribute(sessionContext, initModel, this);
                    //appendAttri.AppendAttributeValuesForEquipment("STENCIL_LOCK_TIME", DateTime.Now.AddHours(Convert.ToDouble(config.LockTime)).ToString("yyyy/MM/dd HH:mm:ss"), equipmentNo);
                    break;
                }
            }
            if (isValid)//continue check material expiry date
            {
                foreach (DataGridViewRow itemM in this.gridSetup.Rows)
                {
                    DateTime dtExpiry = Convert.ToDateTime(itemM.Cells["ExpiryTime"].Value);
                    if (DateTime.Now > dtExpiry)
                    {
                        if (config.DataOutputInterface != null && config.DataOutputInterface != "")
                            SendSN(config.LIGHT_CHANNEL_ON);
                        isValid = false;
                        errorHandler(3, Message("msg_The solder paste has expiry."), "");
                        break;
                    }
                }
            }
            return isValid;
        }

        private bool VerifyActivatedWO()
        {
            bool isValid = true;
            GetCurrentWorkorder getActivatedWOHandler = new GetCurrentWorkorder(sessionContext, initModel, this);
            GetStationSettingModel stationSetting = getActivatedWOHandler.GetCurrentWorkorderResultCall();
            if (stationSetting != null && stationSetting.workorderNumber != null)
            {
                if (stationSetting.workorderNumber == this.txbCDAMONumber.Text)
                {
                    isValid = true;
                }
                else
                {
                    isValid = false;
                    errorHandler(2, Message("msg_The current activated work order has changed, please refresh."), "");
                }
            }
            return isValid;
        }

        private bool VerifyMaterialBinData(string materialBinNo, string partNumber)
        {
            bool isValid = true;
            if (config.ThawingCheck != "Enable")
            {
                return true;
            }
            //是否回温
            GetAttributeValue getAttriHandler = new GetAttributeValue(sessionContext, initModel, this);
            string[] values = getAttriHandler.GetAttributeValueForContainer("SP_THAW_COMPLETE", materialBinNo);
            if (values != null && values.Length > 0)
            {
                DateTime dtCompleteThawing = Convert.ToDateTime(values[1]);//.AddMinutes(GetThawingTime(partNumber))
                if (DateTime.Now < dtCompleteThawing)
                {
                    isValid = false;
                    errorHandler(3, Message("msg_The solder paste thawing not complete."), "");
                }
            }
            else
            {
                isValid = false;
                errorHandler(3, Message("msg_The solder paste thawing not complete."), "");
            }
            return isValid;
        }
        private bool VerifyMaterialBinData24And48(string materialBinNo)
        {
            bool isValid = true;
            if (config.ThawingCheck != "Enable")
            {
                return true;
            }
            //有效期24小时&48小时
            GetAttributeValue getAttriHandler = new GetAttributeValue(sessionContext, initModel, this);
            string[] values = getAttriHandler.GetAttributeValueForContainer("THAW_EXPIRY", materialBinNo);
            if (values != null && values.Length > 0)
            {
                DateTime dtCompleteThawing = Convert.ToDateTime(values[1]);//.AddMinutes(GetThawingTime(partNumber))
                if (DateTime.Now > dtCompleteThawing)
                {
                    isValid = false;
                    errorHandler(3, Message("msg_The solder paste thawing scrap."), "");
                }
                else
                {
                    //if (dtCompleteThawing > DateTime.Now.AddHours(24))
                    //{
                    //    string retutnThawing = DateTime.Now.AddHours(24).ToString("yyyy/MM/dd HH:mm:ss");
                    //    AppendAttribute appendAttriHandler = new AppendAttribute(sessionContext, initModel, this);
                    //    appendAttriHandler.AppendAttributeValuesForContainer("THAW_EXPIRY", retutnThawing, materialBinNo);
                    //}
                }
            }
            else
            {
                isValid = false;
                errorHandler(3, Message("msg_The solder paste thawing scrap."), "");
            }
            return isValid;
        }

        private double GetThawingTime(string partNumber)
        {
            double iValue = 0;
            //get solder paste part number attribute
            GetAttributeValue getAttriHandler = new GetAttributeValue(sessionContext, initModel, this);
            string[] attriCodes = new string[] { "SPThawingTime" };
            Dictionary<string, string> attriValues = getAttriHandler.GetAttributeValueForPart(attriCodes, partNumber);
            if (attriValues != null)
            {
                foreach (var key in attriValues.Keys)
                {
                    if (!string.IsNullOrEmpty(attriValues[key]))
                    {
                        if (key == "SPThawingTime")
                        {
                            iValue = Convert.ToDouble(attriValues[key]);
                        }
                    }
                }
            }
            if (iValue == 0)
            {
                iValue = Convert.ToDouble(config.ThawingDuration);
            }
            return iValue;
        }

        private string ConvertDateFromStamp(string timeStamp)
        {
            double d = Convert.ToDouble(timeStamp);
            DateTime start = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            DateTime date = start.AddMilliseconds(d).ToLocalTime();
            return date.ToString();
        }

        private string ConverToHourAndMin(int number)
        {
            int iHour = number / 60;
            int iMin = number % 60;
            return iHour + "hr " + iMin + "min";
        }

        private bool CheckMaterialSetUp()
        {
            bool isValid = true;
            double iQty = 0;
            foreach (DataGridViewRow row in gridSetup.Rows)
            {
                if (row.Cells["MaterialBinNo"].Value == null || row.Cells["MaterialBinNo"].Value.ToString().Length == 0)
                {
                    errorHandler(3, Message("msg_Material setup required."), "");
                    isValid = false;
                    break;
                }
                iQty += Convert.ToDouble(row.Cells["Qty"].Value);
            }
            if (iQty <= Convert.ToInt32(config.WarningQty) && gridSetup.Rows.Count > 0)
            {
                if (config.DataOutputInterface != null && config.DataOutputInterface != "")
                    SendSN(config.LIGHT_CHANNEL_ON);
                errorHandler(3, Message("msg_Material quantity is not enough."), "");
                isValid = false;
            }
            return isValid;
        }

        #region Compare
        private class RowComparer : System.Collections.IComparer
        {
            private static int sortOrderModifier = 1;

            public RowComparer(SortOrder sortOrder)
            {
                if (sortOrder == SortOrder.Descending)
                {
                    sortOrderModifier = -1;
                }
                else if (sortOrder == SortOrder.Ascending) { sortOrderModifier = 1; }
            }
            public int Compare(object x, object y)
            {
                DataGridViewRow DataGridViewRow1 = (DataGridViewRow)x;
                DataGridViewRow DataGridViewRow2 = (DataGridViewRow)y;
                // Try to sort based on the Scan time column.
                string value1 = DataGridViewRow1.Cells["colItemCode"].Value.ToString();
                string value2 = DataGridViewRow2.Cells["colItemCode"].Value.ToString();
                string type1 = DataGridViewRow1.Cells["colType"].Value.ToString();
                string type2 = DataGridViewRow2.Cells["colType"].Value.ToString();
                int CompareResult = 0;
                if (type1 == type2)
                {
                    CompareResult = value1.CompareTo(value2);
                }
                else
                {
                    CompareResult = type1.CompareTo(type2);
                }
                return CompareResult * sortOrderModifier;
            }
        }
        #endregion

        #region Document
        static string cachePN = "";
        private void InitDocumentGrid()
        {
            if (config.FilterByFileName == "disable") //by station
            {
                if (gridDocument.Rows.Count <= 0)
                {
                    GetDocumentData getDocument = new GetDocumentData(sessionContext, initModel, this);
                    List<DocumentEntity> listDoc = getDocument.GetDocumentDataByStation();
                    if (listDoc != null && listDoc.Count > 0)
                    {
                        foreach (DocumentEntity item in listDoc)
                        {
                            gridDocument.Rows.Add(new object[2] { item.MDA_DOCUMENT_ID, item.MDA_FILE_NAME });
                        }
                    }
                }
            }
            else //by station & filename(partno)
            {
                if (this.txbCDAPartNumber.Text == "" || cachePN == this.txbCDAPartNumber.Text)
                    return;
                cachePN = this.txbCDAPartNumber.Text;
                gridDocument.Rows.Clear();
                this.Invoke(new MethodInvoker(delegate
                {
                    webBrowser1.Navigate("about:blank");
                }));
                GetDocumentData getDocument = new GetDocumentData(sessionContext, initModel, this);
                List<DocumentEntity> listDoc = getDocument.GetDocumentDataByStation();
                if (listDoc != null && listDoc.Count > 0)
                {
                    foreach (DocumentEntity item in listDoc)
                    {
                        string filename = item.MDA_FILE_NAME;
                        Match name = Regex.Match(filename, config.FileNamePattern);
                        if (name.Success)
                        {
                            if (name.Groups.Count > 1)
                            {
                                string partno = name.Groups[1].ToString();
                                if (partno == this.txbCDAPartNumber.Text)
                                {
                                    gridDocument.Rows.Add(new object[2] { item.MDA_DOCUMENT_ID, item.MDA_FILE_NAME });
                                }
                            }
                        }
                    }
                }
            }
        }

        private void GetDocumentCollections()
        {
            GetDocumentData getDocument = new GetDocumentData(sessionContext, initModel, this);
            //get advice id
            Advice[] adviceArray = getDocument.GetAdviceByStationAndPN(this.txbCDAPartNumber.Text);
            if (adviceArray != null && adviceArray.Length > 0)
            {
                int iAdviceID = adviceArray[0].id;
                List<DocumentEntity> list = getDocument.GetDocumentDataByAdvice(iAdviceID);
                if (list != null && list.Count > 0)
                {
                    foreach (var item in list)
                    {
                        string docID = item.MDA_DOCUMENT_ID;
                        string fileName = item.MDA_FILE_NAME;
                        SetDocumentControl(docID, fileName);
                        break;
                    }
                }
            }
        }

        private void SetDocumentControl(string docID, string fileName)
        {
            GetDocumentData documentHandler = new GetDocumentData(sessionContext, initModel, this);
            byte[] content = documentHandler.GetDocumnetContentByID(Convert.ToInt64(docID));
            if (content != null)
            {
                string path = config.MDAPath;
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);
                string filePath = path + @"/" + fileName;
                FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate);
                Encoding.GetEncoding("gb2312");
                fs.Write(content, 0, content.Length);
                fs.Flush();
                fs.Close();
            }
        }

        private void SetDocumentControlForDoc(long documentID, string fileName)
        {
            string path = config.MDAPath;
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            string filePath = path + @"/" + fileName;
            if (!File.Exists(filePath))
            {
                GetDocumentData documentHandler = new GetDocumentData(sessionContext, initModel, this);
                byte[] content = documentHandler.GetDocumnetContentByID(documentID);
                if (content != null)
                {
                    FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate);
                    fs.Write(content, 0, content.Length);
                    fs.Flush();
                    fs.Close();
                }
            }
            this.webBrowser1.Navigate(filePath);
        }
        #endregion

        #region Equipment
        private void InitEquipmentGrid()
        {
            if (string.IsNullOrEmpty(this.txbCDAMONumber.Text))
                return;

            EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
            foreach (DataGridViewRow row in dgvEquipment.Rows)
            {
                string equipmentNo = row.Cells["EquipNo"].Value.ToString();
                string equipmentIndex = row.Cells["EquipmentIndex"].Value.ToString();
                if (string.IsNullOrEmpty(equipmentNo))
                    continue;
                int errorCode = eqManager.UpdateEquipmentData(equipmentIndex, equipmentNo, 1);
                RemoveAttributeForEquipment(equipmentNo, equipmentIndex, "attribEquipmentHasRigged");
            }

            this.dgvEquipment.Rows.Clear();
            List<EquipmentEntity> listEntity = eqManager.GetRequiredEquipmentData(this.txbCDAMONumber.Text);
            if (listEntity != null)
            {
                foreach (var item in listEntity)
                {
                    this.dgvEquipment.Rows.Add(new object[8] { ScreenPrinter.Properties.Resources.Close, item.PART_NUMBER, item.EQUIPMENT_DESCRIPTION, "", "", "", "", "0" });
                }
            }
            this.dgvEquipment.ClearSelection();
        }

        public bool CheckEquipmentSetup()
        {
            foreach (DataGridViewRow row in this.dgvEquipment.Rows)
            {
                if (row.Cells["UsCount"].Value != null && row.Cells["UsCount"].Value.ToString().Length == 0)
                {
                    errorHandler(3, Message("msg_Equipment setup required."), "");
                    return false;
                }
            }
            return true;
        }

        public void ProcessEquipmentData(string equipmentNo)
        {
            if (!VerifyActivatedWO())
                return;
            EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
            string[] values = eqManager.GetEquipmentDetailData(equipmentNo);
            string ePartNumber = "";
            string eIndex = "0";
            int uasedcount = 0;
            if (values != null && values.Length > 0)
            {
                uasedcount = Convert.ToInt32(values[5]);
            }

            if (!CheckEquipmentDuplication(values, ref ePartNumber, ref eIndex))
            {
                errorHandler(3, Message("msg_The equipment") + equipmentNo + Message("msg_ has more Available states."), "");
                return;
            }
            if (!CheckEquipmentValid(ePartNumber))
            {
                errorHandler(3, Message("msg_The equipment is invalid"), "");
                return;
            }
            //check equipment number  whether need to setup?
            if (!CheckEquipmentIsExist(ePartNumber))
                return;
            //check equipment number has rigged on others station
            if (CheckEquipmentHasSetup(equipmentNo, eIndex, "attribEquipmentHasRigged"))
            {
                errorHandler(3, Message("msg_The equipment has rigged on others station."), "");
                return;
            }
            string strEquipmentIndex = eIndex;
            //check the equipment whether is cleaning?
            int equicount = 0;
            Match matchStencil = Regex.Match(equipmentNo, config.StencilPrefix);
            if (matchStencil.Success)
            {
                GetAttributeValue getAttribHandler = new GetAttributeValue(sessionContext, initModel, this);
                string[] valuesTest = getAttribHandler.GetAttributeValueForEquipment("TEST_ITEM", equipmentNo, strEquipmentIndex);
                if (valuesTest == null || valuesTest.Length == 0)
                {
                    errorHandler(3, Message("msg_The equipment is not test"), "");
                    return;
                }

                string[] valuesEquip = getAttribHandler.GetAttributeValueForEquipment("STENCIL_LOCK_TIME", equipmentNo, strEquipmentIndex);
                if (valuesEquip != null && valuesEquip.Length > 0)
                {
                    try
                    {
                        //if (Convert.ToDateTime(valuesEquip[1]) > DateTime.Now)
                        //{
                        errorHandler(3, Message("msg_The equipment is cleaning"), "");
                        return;
                        //}
                    }
                    catch (Exception exx)
                    {

                        LogHelper.Error(exx);
                    }
                }

                int usagecount = 0;
                int maxcount = 0;
                if (config.ReduceEquType == "1")
                {
                    string[] valuesusage = getAttribHandler.GetAttributeValueForEquipment("USAGE_COUNT", equipmentNo, strEquipmentIndex);
                    if (valuesusage != null && valuesusage.Length != 0)
                    {
                        usagecount = Convert.ToInt32(valuesusage[1]);
                    }
                    string[] valuesmax = getAttribHandler.GetAttributeValueForEquipment("MAX_USAGE", equipmentNo, strEquipmentIndex);
                    if (valuesmax != null && valuesmax.Length != 0)
                    {
                        maxcount = Convert.ToInt32(valuesmax[1]);
                    }
                    equicount = maxcount - usagecount;
                    if (equicount <= 0)
                    {
                        errorHandler(3, Message("msg_Stencil usage count cann't less then 0"), "");
                        return;
                    }
                }
                else
                {
                    if (uasedcount <= 0)
                    {
                        errorHandler(3, Message("msg_Stencil usage count cann't less then 0"), "");
                        return;
                    }
                }
            }
            else
            {
                if (uasedcount <= 0)
                {
                    errorHandler(3, Message("msg_Stencil usage count cann't less then 0"), "");
                    return;
                }
            }
            int errorCode = eqManager.UpdateEquipmentData(strEquipmentIndex, equipmentNo, 0);
            if (errorCode == 0)//1301 Equipment is already set up
            {
                //add attribue command the equipment is uesd
                AppendAttributeForEquipment(equipmentNo, strEquipmentIndex, "attribEquipmentHasRigged");
                EquipmentEntityExt entityExt = eqManager.GetSetupEquipmentData(equipmentNo);
                if (entityExt != null)
                {
                    entityExt.PART_NUMBER = ePartNumber;
                    entityExt.EQUIPMENT_INDEX = strEquipmentIndex;
                    SetEquipmentGridData(entityExt, equicount, values[6]);
                    SetTipMessage(MessageType.OK, Message("msg_Process equipment number ") + equipmentNo + Message("msg_SUCCESS"));
                }
            }
        }

        private void SetEquipmentGridData(EquipmentEntityExt entityExt, int usagecount, string expireddate)
        {
            foreach (DataGridViewRow row in this.dgvEquipment.Rows)
            {
                if (row.Cells["eqPartNumber"].Value != null && row.Cells["eqPartNumber"].Value.ToString() == entityExt.PART_NUMBER
                    && (row.Cells["EquipNo"].Value == null || row.Cells["EquipNo"].Value.ToString() == ""))
                {
                    Match matchStencil = Regex.Match(entityExt.EQUIPMENT_NUMBER, config.StencilPrefix);
                    if (matchStencil.Success)
                    {
                        if (config.ReduceEquType == "1")
                            row.Cells["UsCount"].Value = usagecount;
                        else
                        {
                            row.Cells["UsCount"].Value = entityExt.USAGES_BEFORE_EXPIRATION;
                        }


                        row.Cells["NextMaintenance"].Value = DateTime.Now.AddHours(Convert.ToDouble(config.UsageTime)).ToString("yyyy/MM/dd HH:mm:ss");
                    }
                    else
                    {
                        row.Cells["UsCount"].Value = entityExt.USAGES_BEFORE_EXPIRATION;
                        //row.Cells["NextMaintenance"].Value = DateTime.Now.AddSeconds(Convert.ToDouble(entityExt.SECONDS_BEFORE_EXPIRATION)).ToString("yyyy/MM/dd HH:mm:ss");
                        row.Cells["NextMaintenance"].Value = Convert.ToDateTime("1970-01-01 08:00:00").AddMilliseconds(Convert.ToDouble(expireddate)).ToString("yyyy/MM/dd HH:mm:ss");
                    }
                    row.Cells["ScanTime"].Value = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    row.Cells["EquipNo"].Value = entityExt.EQUIPMENT_NUMBER;
                    row.Cells["EquipmentIndex"].Value = entityExt.EQUIPMENT_INDEX;
                    row.Cells["Status"].Value = ScreenPrinter.Properties.Resources.ok;
                    row.Cells["eqPartNumber"].Style.BackColor = Color.FromArgb(0, 192, 0);
                }
            }
        }

        private bool CheckEquipmentIsExist(string partNumber)
        {
            foreach (DataGridViewRow row in this.dgvEquipment.Rows)
            {
                if (row.Cells["eqPartNumber"].Value != null && row.Cells["eqPartNumber"].Value.ToString() == partNumber
                    && row.Cells["EquipNo"].Value.ToString() != "")
                {
                    errorHandler(3, Message("msg_The equipment already exist."), "");
                    return false;
                }
            }
            return true;
        }

        private bool CheckEquipmentValid(string ePartNumber)
        {
            bool isValid = false;
            if (string.IsNullOrEmpty(ePartNumber))
                return false;
            foreach (DataGridViewRow item in this.dgvEquipment.Rows)
            {
                if (item.Cells["eqPartNumber"].Value.ToString() == ePartNumber)// "EQUIPMENT_STATE", "ERROR_CODE", "PART_NUMBER"
                {
                    isValid = true;
                    break;
                }
            }
            return isValid;
        }

        private bool CheckEquipmentDuplication(string[] values, ref string ePartNumber, ref string eIndex)
        {
            int iCount = 0;
            ePartNumber = "";
            eIndex = "0";
            for (int i = 0; i < values.Length; i += 6)
            {
                if (values[i] == "0")
                {
                    ePartNumber = values[i + 2];
                    eIndex = values[i + 3];
                    iCount++;
                }

            }
            if (iCount > 1)
                return false;
            else
                return true;
        }

        private bool CheckEquipmentHasSetup(string equipmentNumber, string equipmentIndex, string attributeCode)
        {
            bool hasSetup = false;
            GetAttributeValue getAttributeHandler = new GetAttributeValue(sessionContext, initModel, this);
            string[] values = getAttributeHandler.GetAttributeValueForAll(15, equipmentNumber, equipmentIndex, attributeCode);
            if (values != null && values.Length > 0)
            {
                hasSetup = true;
            }
            return hasSetup;
        }

        private void AppendAttributeForEquipment(string equipmentNumber, string equipmentIndex, string attributeCode)
        {
            AppendAttribute appendAttriHandler = new AppendAttribute(sessionContext, initModel, this);
            appendAttriHandler.AppendAttributeForAll(15, equipmentNumber, equipmentIndex, attributeCode, "Y");
        }

        private void RemoveAttributeForEquipment(string equipmentNumber, string equipmentIndex, string attributeCode)
        {
            RemoveAttributeValue removeAttriHandler = new RemoveAttributeValue(sessionContext, initModel, this);
            removeAttriHandler.RemoveAttributeForAll(15, equipmentNumber, equipmentIndex, attributeCode);
        }
        #endregion
        #endregion

        #region Network status
        private string strNetMsg = "Network Connected";
        private void picNet_MouseHover(object sender, EventArgs e)
        {
            this.toolTip1.Show(strNetMsg, this.picNet);
        }

        private void AvailabilityChanged(object sender, NetworkAvailabilityEventArgs e)
        {
            if (e.IsAvailable)
            {
                this.picNet.Image = ScreenPrinter.Properties.Resources.NetWorkConnectedGreen24x24;
                this.toolTip1.Show("Network Connected", this.picNet);
                strNetMsg = "Network Connected";
            }
            else
            {
                this.picNet.Image = ScreenPrinter.Properties.Resources.NetWorkDisconnectedRed24x24;
                this.toolTip1.Show("Network Disconnected", this.picNet);
                strNetMsg = "Network Disconnected";
            }
        }
        #endregion

        #region CheckList
        private void btnAddTask_Click(object sender, EventArgs e)
        {
            int iHour = DateTime.Now.Hour;
            if (8 <= iHour && iHour <= 18)
            {
                gridCheckList.Rows.Add(new object[] { this.gridCheckList.Rows.Count + 1, DateTime.Now.ToString("yyyy/MM/dd"), "白班", "", "", "", "", "", "", "", "" });
            }
            else
            {
                gridCheckList.Rows.Add(new object[] { this.gridCheckList.Rows.Count + 1, DateTime.Now.ToString("yyyy/MM/dd"), "晚班", "", "", "", "", "", "", "", "" });
            }
            gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clResult1"].ReadOnly = true;
            gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clSeq"].ReadOnly = true;
            gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clDate"].ReadOnly = true;
            gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clShift"].ReadOnly = true;
            gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clStatus"].ReadOnly = true;
            gridCheckList.ClearSelection();
        }

        private void btnConfirm_Click(object sender, EventArgs e)
        {
            //if (!VerifyCheckList())
            //{
            //    //errorHandler(2, Message("$msg_checklist_first"), "");
            //    return;
            //}
            CheckListsCreate();
            #region
            //if (gridCheckList.Rows.Count > 0)
            //{
            //    string targetFileName = "";
            //    string shortFileName = config.StationNumber + "_" + this.gridCheckList.Rows[0].Cells["clShift"].Value.ToString() + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            //    bool isOK = CreateTemplate(shortFileName, ref targetFileName);
            //    if (isOK)
            //    {
            //        Excel.Application xlsApp = null;
            //        Excel._Workbook xlsBook = null;
            //        Excel._Worksheet xlsSheet = null;
            //        try
            //        {
            //            GC.Collect();
            //            xlsApp = new Excel.Application();
            //            xlsApp.DisplayAlerts = false;
            //            xlsApp.Workbooks.Open(targetFileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            //            xlsBook = xlsApp.ActiveWorkbook;
            //            xlsSheet = (Excel._Worksheet)xlsBook.ActiveSheet;

            //            int iBeginIndex = 7;
            //            Excel.Range range = null;
            //            foreach (DataGridViewRow row in gridCheckList.Rows)
            //            {
            //                range = (Excel.Range)xlsSheet.Rows[iBeginIndex, Missing.Value];
            //                range.Rows.Insert(Excel.XlDirection.xlDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            //                string strSeq = row.Cells["clSeq"].Value.ToString();
            //                string strItemName = row.Cells["clItemName"].Value.ToString();
            //                string strItemPoint = row.Cells["clPoint"].Value.ToString();
            //                string strItemStandard = row.Cells["clStandard"].Value.ToString();
            //                string strItemMethod = row.Cells["clMethod"].Value.ToString();
            //                string strItemResult = GetCheckItemResult(row.Cells["clResult1"].Value.ToString(), row.Cells["clResult2"].Value.ToString());
            //                string strCheckDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            //                string strException = row.Cells["clException"].Value == null ? "" : row.Cells["clException"].Value.ToString();
            //                string strHappendTime = row.Cells["clChangeDate"].Value == null ? "" : row.Cells["clChangeDate"].Value.ToString();
            //                string strProcessContent = row.Cells["clContent"].Value == null ? "" : row.Cells["clContent"].Value.ToString();
            //                string strProcessPersion = row.Cells["clPersion"].Value == null ? "" : row.Cells["clPersion"].Value.ToString();
            //                string strOperator = row.Cells["clOperator"].Value == null ? "" : row.Cells["clOperator"].Value.ToString();
            //                string strLeader = row.Cells["clLeader"].Value == null ? "" : row.Cells["clLeader"].Value.ToString();
            //                xlsSheet.Cells[iBeginIndex, 1] = strSeq;
            //                xlsSheet.Cells[iBeginIndex, 2] = strItemName;
            //                xlsSheet.Cells[iBeginIndex, 3] = strItemPoint;
            //                xlsSheet.Cells[iBeginIndex, 4] = strItemStandard;
            //                xlsSheet.Cells[iBeginIndex, 5] = strItemMethod;
            //                xlsSheet.Cells[iBeginIndex, 6] = strItemResult;
            //                xlsSheet.Cells[iBeginIndex, 7] = strCheckDate;
            //                xlsSheet.Cells[iBeginIndex, 8] = strException;
            //                xlsSheet.Cells[iBeginIndex, 9] = strHappendTime;
            //                xlsSheet.Cells[iBeginIndex, 10] = strProcessContent;
            //                xlsSheet.Cells[iBeginIndex, 11] = strProcessPersion;
            //                xlsSheet.Cells[iBeginIndex, 12] = strOperator;
            //                xlsSheet.Cells[iBeginIndex, 13] = strLeader;
            //                iBeginIndex++;
            //            }
            //            xlsBook.Save();
            //            errorHandler(0, "Save Production Check List success.(" + targetFileName + ")", "");
            //        }
            //        catch (Exception ex)
            //        {
            //            LogHelper.Error(ex);
            //        }
            //        finally
            //        {
            //            xlsBook.Close(false, Type.Missing, Type.Missing);
            //            xlsApp.Quit();
            //            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsApp);
            //            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsBook);
            //            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsSheet);

            //            xlsSheet = null;
            //            xlsBook = null;
            //            xlsApp = null;

            //            GC.Collect();
            //            GC.WaitForPendingFinalizers();
            //        }
            //    }
            //}
            #endregion
        }

        #region add by qy
        private void CheckListsCreate()
        {
            if (gridCheckList.Rows.Count > 0)
            {
                string targetFileName = "";
                string shortFileName = config.StationNumber + "_ICT_" + DateTime.Now.ToString("yyyyMM");
                bool isOK = CreateTemplate(shortFileName, ref targetFileName);
                if (isOK)
                {
                    Excel.Application xlsApp = null;
                    Excel._Workbook xlsBook = null;
                    Excel._Worksheet xlsSheet = null;
                    try
                    {
                        GC.Collect();
                        xlsApp = new Excel.Application();
                        xlsApp.DisplayAlerts = false;
                        xlsApp.Workbooks.Open(targetFileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                        xlsBook = xlsApp.ActiveWorkbook;
                        xlsSheet = (Excel._Worksheet)xlsBook.ActiveSheet;
                        int count = xlsSheet.UsedRange.Cells.Rows.Count;

                        int iBeginIndex = count;
                        Excel.Range range = null;
                        foreach (DataGridViewRow row in gridCheckList.Rows)
                        {
                            range = (Excel.Range)xlsSheet.Rows[iBeginIndex, Missing.Value];
                            range.Rows.Insert(Excel.XlDirection.xlDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                            string strSeq = row.Cells["clSeq"].Value.ToString();
                            string strItemName = row.Cells["clItemName"].Value.ToString();
                            string strItemPoint = row.Cells["clPoint"].Value.ToString();
                            string strItemStandard = row.Cells["clStandard"].Value.ToString();
                            string strItemMethod = row.Cells["clMethod"].Value.ToString();
                            string strItemResult = GetCheckItemResult(row.Cells["clResult1"].Value.ToString(), row.Cells["clResult2"].Value.ToString());
                            string strCheckDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                            string strShift = row.Cells["clShift"].Value.ToString();
                            string strException = row.Cells["clException"].Value == null ? "" : row.Cells["clException"].Value.ToString();
                            string strHappendTime = row.Cells["clChangeDate"].Value == null ? "" : row.Cells["clChangeDate"].Value.ToString();
                            string strProcessContent = row.Cells["clContent"].Value == null ? "" : row.Cells["clContent"].Value.ToString();
                            string strProcessPersion = row.Cells["clPersion"].Value == null ? "" : row.Cells["clPersion"].Value.ToString();
                            string strOperator = row.Cells["clOperator"].Value == null ? "" : row.Cells["clOperator"].Value.ToString();
                            string strLeader = row.Cells["clLeader"].Value == null ? "" : row.Cells["clLeader"].Value.ToString();

                            xlsSheet.Cells[iBeginIndex, 1] = iBeginIndex - 7;
                            xlsSheet.Cells[iBeginIndex, 2] = strItemName;
                            xlsSheet.Cells[iBeginIndex, 3] = strItemPoint;
                            xlsSheet.Cells[iBeginIndex, 4] = strItemStandard;
                            xlsSheet.Cells[iBeginIndex, 5] = strItemMethod;
                            xlsSheet.Cells[iBeginIndex, 6] = strItemResult;
                            xlsSheet.Cells[iBeginIndex, 7] = strShift;
                            xlsSheet.Cells[iBeginIndex, 8] = strCheckDate;
                            xlsSheet.Cells[iBeginIndex, 9] = strException;
                            xlsSheet.Cells[iBeginIndex, 10] = strHappendTime;
                            xlsSheet.Cells[iBeginIndex, 11] = strProcessContent;
                            xlsSheet.Cells[iBeginIndex, 12] = strProcessPersion;
                            xlsSheet.Cells[iBeginIndex, 13] = strOperator;
                            xlsSheet.Cells[iBeginIndex, 14] = strLeader;

                            iBeginIndex++;
                        }
                        xlsBook.Save();
                        errorHandler(0, Message("msg_Save_CheckList_Success") + ".(" + targetFileName + ")", "");
                        SetTipMessage(MessageType.OK, Message("msg_Save_CheckList_Success") + ".(" + targetFileName + ")");
                    }
                    catch (Exception ex)
                    {
                        LogHelper.Error(ex);
                    }
                    finally
                    {
                        xlsBook.Close(false, Type.Missing, Type.Missing);
                        xlsApp.Quit();
                        KillSpecialExcel(xlsApp);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsApp);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsBook);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsSheet);

                        xlsSheet = null;
                        xlsBook = null;
                        xlsApp = null;

                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                }
            }
        }

        private bool CheckCheckList()
        {
            bool result = true;
            foreach (DataGridViewRow row in gridCheckList.Rows)
            {
                string status = row.Cells["clStatus"].Value.ToString();
                if (status != "OK")
                {
                    result = false;
                    errorHandler(2, Message("$msg_Verify_CheckList"), "");
                    break;
                }
            }
            return result;
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
        private static void KillSpecialExcel(Excel.Application m_objExcel)
        {
            try
            {
                if (m_objExcel != null)
                {
                    int lpdwProcessId;
                    GetWindowThreadProcessId(new IntPtr(m_objExcel.Hwnd), out lpdwProcessId);
                    System.Diagnostics.Process.GetProcessById(lpdwProcessId).Kill();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        private string GetCheckItemResult(string result1, string result2)
        {
            if (string.IsNullOrEmpty(result1))
                return result2;
            if (string.IsNullOrEmpty(result2))
                return result1;
            else
                return "NA";
        }

        private void InitTaskData()
        {
            try
            {
                gridCheckList.Rows.Clear();
                Dictionary<string, List<CheckListItemEntity>> dicTask = new Dictionary<string, List<CheckListItemEntity>>();
                XDocument xdc = XDocument.Load("TaskFile.xml");
                var stationNodes = from item in xdc.Descendants("StationNumber")
                                   where item.Attribute("value").Value == config.StationNumber
                                   select item;
                XElement stationNode = stationNodes.FirstOrDefault();
                var tasks = from item in stationNode.Descendants("shift")
                            select item;
                foreach (XElement node in tasks.ToList())
                {
                    string shiftValue = GetNoteAttributeValues(node, "value");
                    List<CheckListItemEntity> itemList = new List<CheckListItemEntity>();
                    var items = from item in node.Descendants("Item")
                                select item;
                    foreach (XElement subItem in items.ToList())
                    {
                        CheckListItemEntity entity = new CheckListItemEntity();
                        entity.ItemName = GetNoteAttributeValues(subItem, "name");
                        entity.ItemPoint = GetNoteAttributeValues(subItem, "point");
                        entity.ItemStandard = GetNoteAttributeValues(subItem, "standard");
                        entity.ItemMethod = GetNoteAttributeValues(subItem, "method");
                        entity.ItemInputType = GetNoteAttributeValues(subItem, "inputType");
                        itemList.Add(entity);
                    }
                    if (!dicTask.ContainsKey(shiftValue))
                    {
                        dicTask[shiftValue] = itemList;
                    }
                }
                //init check list grid
                string strInputValue = GetNoteDescendantsValues(stationNode, "DataInputType");
                string[] strInputValues = strInputValue.Split(new char[] { ',' });
                DataTable dtInput = new DataTable();
                dtInput.Columns.Add("name");
                dtInput.Columns.Add("value");
                DataRow rowEmpty = dtInput.NewRow();
                rowEmpty["name"] = "";
                rowEmpty["value"] = "";
                dtInput.Rows.Add(rowEmpty);
                foreach (var strValues in strInputValues)
                {
                    DataRow row = dtInput.NewRow();
                    row["name"] = strValues;
                    row["value"] = strValues;
                    dtInput.Rows.Add(row);
                }
                ((DataGridViewComboBoxColumn)this.gridCheckList.Columns["clResult2"]).DataSource = dtInput;
                ((DataGridViewComboBoxColumn)this.gridCheckList.Columns["clResult2"]).DisplayMember = "Name";
                ((DataGridViewComboBoxColumn)this.gridCheckList.Columns["clResult2"]).ValueMember = "Value";

                int iHour = DateTime.Now.Hour;
                int seq = 1;
                if (8 <= iHour && iHour <= 18)
                {
                    if (dicTask.ContainsKey("白班"))
                    {
                        List<CheckListItemEntity> itemList = dicTask["白班"];
                        if (itemList != null && itemList.Count > 0)
                        {
                            foreach (var item in itemList)
                            {
                                object[] objValues = new object[11] { seq, DateTime.Now.ToString("yyyy/MM/dd"), "白班", item.ItemName, item.ItemPoint, item.ItemStandard, item.ItemMethod, "", "", "", item.ItemInputType };
                                this.gridCheckList.Rows.Add(objValues);
                                seq++;
                            }
                            SetCheckListInputStatus();
                            this.gridCheckList.ClearSelection();
                        }
                    }
                }
                else
                {
                    if (dicTask.ContainsKey("晚班"))
                    {
                        List<CheckListItemEntity> itemList = dicTask["晚班"];
                        if (itemList != null && itemList.Count > 0)
                        {
                            foreach (var item in itemList)
                            {
                                object[] objValues = new object[11] { seq, DateTime.Now.ToString("yyyy/MM/dd"), "晚班", item.ItemName, item.ItemPoint, item.ItemStandard, item.ItemMethod, "", "", "", item.ItemInputType };
                                this.gridCheckList.Rows.Add(objValues);
                                seq++;
                            }
                            SetCheckListInputStatus();
                            this.gridCheckList.ClearSelection();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex);
            }
        }

        private string GetNoteAttributeValues(XElement node, string attributename)
        {
            string strValue = "";
            try
            {
                strValue = node.Attribute(attributename).Value;
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex);
            }
            return strValue;
        }

        private string GetNoteDescendantsValues(XElement node, string attributename)
        {
            string strValue = "";
            try
            {
                strValue = node.Descendants(attributename).First().Value;
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex);
            }
            return strValue;
        }

        private string GetNoteDescendantsAttributeValues(XElement node, string nodeName, string attributeName)
        {
            string strValue = "";
            try
            {
                strValue = node.Descendants(nodeName).First().Attribute(attributeName).Value;
            }
            catch (Exception ex)
            {
                //MeasuredOctet 
                strValue = node.Descendants("RepairAction").First().Attribute("repairKey").Value;
                LogHelper.Info(node.ToString());
                LogHelper.Info("Node Name: " + nodeName);
                LogHelper.Info("Attribute Name: " + attributeName);
                LogHelper.Error(ex);
            }
            return strValue;
        }

        private void SetCheckListInputStatus()
        {
            foreach (DataGridViewRow row in this.gridCheckList.Rows)
            {
                if (row.Cells["clInputType"].Value.ToString() == "1")
                {
                    row.Cells["clResult1"].ReadOnly = true;
                }
                else if (row.Cells["clInputType"].Value.ToString() == "2")
                {
                    row.Cells["clResult2"].ReadOnly = true;
                }
                row.Cells["clSeq"].ReadOnly = true;
                row.Cells["clDate"].ReadOnly = true;
                row.Cells["clShift"].ReadOnly = true;
                row.Cells["clStatus"].ReadOnly = true;
            }
        }

        private void gridCheckList_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1)
                return;
            if (this.gridCheckList.Columns[e.ColumnIndex].Name == "clResult1" && this.gridCheckList.Rows[e.RowIndex].Cells["clResult1"].Value.ToString() != "")
            {
                //verify the input value
                string strRegex = @"^(\d{0,9}.\d{0,9})-(\d{0,9}.\d{0,9}).*$";
                string strResult1 = this.gridCheckList.Rows[e.RowIndex].Cells["clResult1"].Value.ToString();
                string strStandard = this.gridCheckList.Rows[e.RowIndex].Cells["clStandard"].Value.ToString();
                Match match = Regex.Match(strStandard, strRegex);
                if (match.Success)
                {
                    if (match.Groups.Count > 2)
                    {
                        double iMin = Convert.ToDouble(match.Groups[1].Value);
                        double iMax = Convert.ToDouble(match.Groups[2].Value);
                        double iResult = Convert.ToDouble(strResult1);
                        if (iResult >= iMin && iResult <= iMax)
                        {
                            this.gridCheckList.Rows[e.RowIndex].Cells["clStatus"].Style.BackColor = Color.FromArgb(0, 192, 0);
                            this.gridCheckList.Rows[e.RowIndex].Cells["clStatus"].Value = "OK";
                        }
                        else
                        {
                            this.gridCheckList.Rows[e.RowIndex].Cells["clStatus"].Style.BackColor = Color.Red;
                            this.gridCheckList.Rows[e.RowIndex].Cells["clStatus"].Value = "NG";
                        }
                    }
                }
                else
                {
                    this.gridCheckList.Rows[e.RowIndex].Cells["clStatus"].Style.BackColor = Color.Red;
                    this.gridCheckList.Rows[e.RowIndex].Cells["clStatus"].Value = "NG";
                }
            }
            //else if (this.gridCheckList.Columns[e.ColumnIndex].Name == "clResult2" && this.gridCheckList.Rows[e.RowIndex].Cells["clResult2"].Value != null
            //    && this.gridCheckList.Rows[e.RowIndex].Cells["clResult2"].Value.ToString() != "")
            //{
            //    //verify the select value
            //    this.gridCheckList.Rows[e.RowIndex].Cells["clStatus"].Style.BackColor = Color.FromArgb(0, 192, 0);
            //    this.gridCheckList.Rows[e.RowIndex].Cells["clStatus"].Value = "OK";
            //}
        }

        #region Grid ComboBox
        int iRowIndex = -1;
        private void gridCheckList_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            DataGridView dgv = sender as DataGridView;
            if (dgv.CurrentCell.GetType().Name == "DataGridViewComboBoxCell" && dgv.CurrentCell.RowIndex != -1)
            {
                iRowIndex = dgv.CurrentCell.RowIndex;
                (e.Control as ComboBox).SelectedIndexChanged += new EventHandler(ComboBox_SelectedIndexChanged);
            }
        }

        public void ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox combox = sender as ComboBox;
            combox.Leave += new EventHandler(combox_Leave);
            try
            {
                if (combox.SelectedItem != null && combox.Text != "")
                {
                    if (OKlist.Contains(combox.Text))
                    {
                        this.gridCheckList.Rows[iRowIndex].Cells["clStatus"].Style.BackColor = Color.FromArgb(0, 192, 0);
                        this.gridCheckList.Rows[iRowIndex].Cells["clStatus"].Value = "OK";
                    }
                    else
                    {
                        this.gridCheckList.Rows[iRowIndex].Cells["clStatus"].Style.BackColor = Color.Red;
                        this.gridCheckList.Rows[iRowIndex].Cells["clStatus"].Value = "NG";
                    }
                }
                else
                {
                    this.gridCheckList.Rows[iRowIndex].Cells["clStatus"].Style.BackColor = Color.White;
                    this.gridCheckList.Rows[iRowIndex].Cells["clStatus"].Value = "";
                }
                Thread.Sleep(100);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void combox_Leave(object sender, EventArgs e)
        {
            ComboBox combox = sender as ComboBox;
            combox.SelectedIndexChanged -= new EventHandler(ComboBox_SelectedIndexChanged);
        }
        #endregion

        int iIndexCheckList = -1;
        private void gridCheckList_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (this.gridCheckList.Rows.Count == 0)
                    return;
                this.gridCheckList.ContextMenuStrip = contextMenuStrip2;
                iIndexCheckList = ((DataGridView)sender).CurrentRow.Index;
                ((DataGridView)sender).CurrentRow.Selected = true;
            }
        }

        private bool CreateTemplate(string strFileName, ref string targetFileName)
        {
            bool bFlag = true;
            targetFileName = "";
            string filePath = Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName;
            string _appDir = Path.GetDirectoryName(filePath);
            string strExportPath = _appDir + @"\CheckListFiles\";
            //临时文件目录
            if (Directory.Exists(strExportPath) == false)
            {
                Directory.CreateDirectory(strExportPath);
            }
            string strSourceFileName = strExportPath + @"CheckListTemplate.xls";
            string strTargetFileName = config.CheckListFolder + strFileName + ".xls";
            targetFileName = strTargetFileName;
            if (!Directory.Exists(config.CheckListFolder))
                Directory.CreateDirectory(config.CheckListFolder);
            if (File.Exists(targetFileName))
            {
                return true;
            }
            if (System.IO.File.Exists(strSourceFileName))
            {
                try
                {
                    System.IO.File.Copy(strSourceFileName, strTargetFileName, true);
                    //去掉文件Readonly,避免不可写
                    FileInfo file = new FileInfo(strTargetFileName);
                    if ((file.Attributes & FileAttributes.ReadOnly) > 0)
                    {
                        file.Attributes ^= FileAttributes.ReadOnly;
                    }
                }
                catch (Exception ex)
                {
                    bFlag = false;
                    LogHelper.Error(ex);
                    throw ex;
                }
            }
            else
            {
                bFlag = false;
            }

            return bFlag;
        }

        private bool VerifyCheckList()
        {
            if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
            {
                //if (!CheckShiftChange2())
                //{
                //    if (this.dgvCheckListTable.Rows.Count <= 0 && dgvCheckListTable.Rows[0].Cells["tabdjclass"].Value.ToString() != "开线点检")
                //    {
                //        InitTaskData_SOCKET("开线点检");
                //    }

                //}
                foreach (DataGridViewRow row in this.dgvCheckListTable.Rows)
                {
                    if (row.Cells["tabStatus"].Value.ToString() != "OK")
                    {
                        errorHandler(2, Message("msg_Verify_CheckList"), "");
                        return false;
                    }
                }
                if (this.dgvCheckListTable.Rows.Count > 0)
                {
                    if (!Supervisor)
                    {
                        errorHandler(2, Message("msg_Superivisor_check_fail"), "");
                        return false;
                    }
                    if (!IPQC)
                    {

                        errorHandler(2, Message("msg_IPQC_check_fail"), "");
                        return false;
                    }
                }

                return true;
            }
            else
            {
                foreach (DataGridViewRow row in gridCheckList.Rows)
                {
                    if (row.Cells["clStatus"].Value.ToString() != "OK")
                    {

                        errorHandler(2, Message("msg_Verify_CheckList"), "");
                        return false;
                    }
                }
                //if (this.gridCheckList.Rows.Count > 0)
                //{
                //    if (!Supervisor)
                //    {

                //        errorHandler(2, Message("msg_Superivisor_check_fail"), "");
                //        return false;
                //    }
                //    if (!IPQC)
                //    {

                //        errorHandler(2, Message("msg_IPQC_check_fail"), "");
                //        return false;
                //    }
                //}
                return true;
            }
        }

        private void checkListAdd_Click(object sender, EventArgs e)
        {
            if (iIndexCheckList > -1)
            {
                int iHour = DateTime.Now.Hour;
                if (8 <= iHour && iHour <= 18)
                {
                    gridCheckList.Rows.Add(new object[] { this.gridCheckList.Rows.Count + 1, DateTime.Now.ToString("yyyy/MM/dd"), "白班", "", "", "", "", "", "", "", "" });
                }
                else
                {
                    gridCheckList.Rows.Add(new object[] { this.gridCheckList.Rows.Count + 1, DateTime.Now.ToString("yyyy/MM/dd"), "晚班", "", "", "", "", "", "", "", "" });
                }
                gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clResult1"].ReadOnly = true;
                gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clSeq"].ReadOnly = true;
                gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clDate"].ReadOnly = true;
                gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clShift"].ReadOnly = true;
                gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clStatus"].ReadOnly = true;
                gridCheckList.ClearSelection();
            }
        }

        private void checkListDelete_Click(object sender, EventArgs e)
        {
            if (iIndexCheckList > -1)
            {
                this.gridCheckList.Rows.RemoveAt(iIndexCheckList);
                int seq = 1;
                foreach (DataGridViewRow row in this.gridCheckList.Rows)
                {
                    row.Cells["clSeq"].Value = seq;
                    seq++;
                }
                this.gridCheckList.ClearSelection();
            }
        }

        //20161214 add by qy
        bool IsGetShiftCheckList = false;
        private void InitShiftCheckList()
        {
            if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
            {
                //InitShift2(txbCDAMONumber.Text);
                if (!CheckShiftChange2())
                {
                    if (this.dgvCheckListTable.Rows.Count <= 0 || (this.dgvCheckListTable.Rows.Count > 0 && !isStartLineCheck))//!IsShiftCheck()
                    {
                        InitTaskData_SOCKET("开线点检;设备点检");
                        isStartLineCheck = true;
                    }
                }
            }
        }
        private bool IsShiftCheck()//true 表示已经带出开线点检的内容了
        {
            bool isValid = false;
            foreach (DataGridViewRow row in this.dgvCheckListTable.Rows)
            {
                if (row.Cells["tabdjclass"].Value.ToString() == "开线点检")
                {
                    isValid = true;
                    break;
                }
            }
            return isValid;
        }
        #endregion

        #region PFC
        private object _lock1 = new Object();
        public void ProcessPFCMessage(string pfcMsg)
        {
            //#!PONGCRLF
            //#!BCTOPBARCODECRLF
            //#!BCBOTTOMBARCODECRLF
            //#!BOARDAVCRLF
            //#!TRANSFERBARCODECRLF
            lock (_lock1)
            {
                if (isFormOutPoump)//登出状态不做操作
                {
                    return;
                }
                if (IPIform != null && IPIform.Visible == true)
                    return; ;
                //errorHandler(0, "Receive message from PFC " + pfcMsg.TrimEnd(), "");
                SetConnectionText(0, "Receive message from PFC " + pfcMsg.TrimEnd());
                LogHelper.Info("Receive message from PFC " + pfcMsg.TrimEnd());
                if (pfcMsg.Length >= 10)
                {
                    bool isOK = true;
                    string messageType = pfcMsg.Substring(2, 8).TrimEnd();
                    switch (messageType)
                    {
                        case "PONG":
                            PFCStartTime = DateTime.Now;
                            break;
                        case "BCTOP":
                            if (!VerifyCheckList())
                            {
                                break;
                            }
                            string serialNumber = pfcMsg.Substring(10).Replace("#!PONG", "").TrimEnd();
                            isOK = ProcessSerialNumberEXT(serialNumber);
                            if (isOK)
                            {
                                SendMsessageToPFC(PFCMessage.GO, serialNumber);
                                //if (config.IPI_STATUS_CHECK == "ENABLE" && IPIStatus != 1)
                                //{
                                //    IPITimerStart();
                                //    this.Invoke(new MethodInvoker(delegate
                                //    {
                                //        IPIFormPoupFirst();
                                //    }));
                                //}
                            }
                            break;
                        case "BCBOTTOM":
                            if (!VerifyCheckList())
                            {
                                break;
                            }
                            string serialNumber1 = pfcMsg.Substring(10).Replace("#!PONG", "").TrimEnd();
                            isOK = ProcessSerialNumberEXT(serialNumber1);
                            if (isOK)
                            {
                                SendMsessageToPFC(PFCMessage.GO, serialNumber1);
                            }
                            break;
                        case "BOARDAV"://todo
                            //BoardCome = true;
                            ////isOK = ProcessSerialNumberData();
                            //if (isOK)
                            //{
                            //    SendMsessageToPFC(PFCMessage.GO, "");
                            //    BoardCome = false;
                            //}
                            break;
                        case "TRANSFER":
                            string serialNumber2 = pfcMsg.Substring(10).Replace("#!PONG", "").TrimEnd();
                            SendMsessageToPFC(PFCMessage.COMPLETE, serialNumber2);

                            if (config.IPI_STATUS_CHECK == "ENABLE")
                            {
                                if (!config.IPI_WORKORDERTYPE_CHECK.Contains(workorderType))
                                {
                                    if (IPIStatus == 0 || IPIStatus == 4)
                                    {
                                        IPITimerStart();
                                        //this.Invoke(new MethodInvoker(delegate
                                        //{
                                        this.BeginInvoke(new Action(() =>
                                        {
                                            IPIFormPoup();
                                        }));
                                    }
                                }
                            }
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    MessageBox.Show("Receive message length less then 10.");
                }
            }
        }

        private object _lockSend = new Object();
        public void SendMsessageToPFC(PFCMessage msgType, string serialNumber)
        {
            lock (_lockSend)
            {
                //#!PINGCRLF
                //#!GOBARCODECRLF
                //#!COMPLETEBARCODECRLF
                //if (IPIform == null || !IPIform.Visible)
                //{
                string prefix = "#!";
                string suffix = HexToStr1("0D") + HexToStr1("0A");
                string sendMessage = "";
                switch (msgType)
                {
                    case PFCMessage.PING:
                        sendMessage = prefix + PFCMessage.PING.ToString().PadRight(8, ' ') + suffix;
                        break;
                    case PFCMessage.GO:
                        sendMessage = prefix + PFCMessage.GO.ToString().PadRight(8, ' ') + serialNumber + suffix;
                        break;
                    case PFCMessage.COMPLETE:
                        sendMessage = prefix + PFCMessage.COMPLETE.ToString().PadRight(8, ' ') + serialNumber + suffix;
                        break;
                    case PFCMessage.CONFIRM:
                        sendMessage = prefix + PFCMessage.CONFIRM.ToString().PadRight(8, ' ') + serialNumber + suffix;
                        break;
                    default:
                        sendMessage = prefix + PFCMessage.PING.ToString().PadRight(8, ' ') + suffix;
                        break;
                }
                //send message through socket
                try
                {
                    if (DateTime.Now.Subtract(PFCStartTime).Seconds >= 20)
                    {
                        if (cSocket.tcpsend.Connected)
                        {
                            cSocket.send(prefix + PFCMessage.PING.ToString().PadRight(8, ' ') + suffix);
                            PFCStartTime = DateTime.Now;
                            Thread.Sleep(1000);
                        }
                    }
                    bool isOK = true;
                    if (cSocket.tcpsend.Connected)
                    {
                        isOK = cSocket.send(sendMessage);
                        if (isOK)
                        {
                            //errorHandler(1, "Send message to PFC:" + sendMessage.TrimEnd(), "");
                            SetConnectionText(0, "Send message to PFC:" + sendMessage.TrimEnd());
                            LogHelper.Info("Send message to PFC:" + sendMessage.TrimEnd());
                        }
                    }
                    else
                    {

                        //errorHandler(2, "Send message to PFC:" + sendMessage.TrimEnd(), "");
                        SetConnectionText(1, "Send message to PFC:" + sendMessage.TrimEnd());
                        bool isConnectOK = cSocket.connect(config.IPAddress, config.Port);
                        if (isConnectOK)
                        {

                            isOK = cSocket.send(sendMessage);
                            if (isOK)
                            {
                                SetConnectionText(0, "Send message to PFC:" + sendMessage.TrimEnd());
                            }
                            else
                            {
                                SetConnectionText(1, "Send message to PFC:" + sendMessage.TrimEnd());
                            }
                        }
                        else
                        {
                            SetConnectionText(1, "Conncet to PFC error");
                        }
                    }


                }
                catch (Exception ex)
                {
                    cSocket.send(prefix + PFCMessage.PING.ToString().PadRight(8, ' ') + suffix);
                    bool isOK = cSocket.send(sendMessage);
                    if (isOK)
                    {
                        //errorHandler(0, "Send message to PFC:" + sendMessage.TrimEnd(), "");
                        SetConnectionText(1, "Send message to PFC:" + sendMessage.TrimEnd());
                    }
                    else
                    {
                        SetConnectionText(1, "Send message to PFC:" + sendMessage.TrimEnd());
                    }
                    LogHelper.Error(ex.Message, ex);
                }
            }
            //}
        }

        public static string HexToStr1(string mHex) // 返回十六进制代表的字符串
        {
            mHex = mHex.Replace(" ", "");
            if (mHex.Length <= 0) return "";
            byte[] vBytes = new byte[mHex.Length / 2];
            for (int i = 0; i < mHex.Length; i += 2)
                if (!byte.TryParse(mHex.Substring(i, 2), NumberStyles.HexNumber, null, out vBytes[i / 2]))
                    vBytes[i / 2] = 0;
            return ASCIIEncoding.Default.GetString(vBytes);
        }

        public void GetTimerStart()
        {
            CheckConnectTimer = new System.Timers.Timer();
            if (CheckConnectTimer.Enabled)
                return;
            // 循环间隔时间(1分钟)
            CheckConnectTimer.Interval = Convert.ToInt32(config.GateKeeperTimer);
            // 允许Timer执行
            CheckConnectTimer.Enabled = true;
            // 定义回调
            CheckConnectTimer.Elapsed += new ElapsedEventHandler(CheckConnectTimer_Elapsed);
            // 定义多次循环
            CheckConnectTimer.AutoReset = true;
        }

        private void CheckConnectTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (isFormOutPoump)//登出状态不做操作
            {
                return;
            }
            SendMsessageToPFC(PFCMessage.PING, "");
        }
        #endregion

        #region Active wo add by qy
        private void btnActivate_Click(object sender, EventArgs e)
        {
            //if (!VerifyCheckList())
            //{
            //    return;
            //}
            try
            {
                string workorder = "";
                strShiftChecklist = "";
                bool isInitChecklist = false;
                string WorkorderPre = txbCDAMONumber.Text;
                int processlayerPre = initModel.currentSettings == null ? -1 : initModel.currentSettings.processLayer;
                if (this.gridWorkorder.SelectedRows.Count > 0)
                {
                    workorder = this.gridWorkorder.SelectedRows[0].Cells["columnWoNumber"].Value.ToString();
                    ActivateWorkorder activateHandler = new ActivateWorkorder(sessionContext, initModel, this);
                    int error = activateHandler.ActivateWorkorderExt(initModel.configHandler.StationNumber, workorder, 1, ConvertProcessLayerToString(this.cmbLayer.Text));//1 = Activate work order for the station only
                    if (error == 0)
                    {
                        this.txbCDAMONumber.Text = this.gridWorkorder.SelectedRows[0].Cells["columnWoNumber"].Value.ToString();
                        this.txbCDAPartNumber.Text = this.gridWorkorder.SelectedRows[0].Cells["columnPn"].Value.ToString();
                        GetCurrentWorkorder getCurrentHandler = new GetCurrentWorkorder(sessionContext, initModel, this);
                        GetStationSettingModel model = getCurrentHandler.GetCurrentWorkorderResultCall();
                        initModel.currentSettings = model;
                        if (model != null && model.workorderNumber != null)
                        {
                            GetNumbersOfSingleBoards getNumBoard = new GetNumbersOfSingleBoards(sessionContext, initModel, this);
                            List<MdataGetPartData> listData = getNumBoard.GetNumbersOfSingleBoardsResultCall(model.partNumber);
                            if (listData != null && listData.Count > 0)
                            {
                                MdataGetPartData mData = listData[0];
                                initModel.numberOfSingleBoards = mData.quantityMultipleBoard;
                            }
                        }
                        LoadYield();

                        if (workorder != WorkorderPre || processlayerPre != initModel.currentSettings.processLayer)
                        {
                            strShiftChecklist = "";
                            InitWorkOrderType();
                            InitShift2(WorkorderPre);
                            if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
                            {
                                if (!CheckShiftChange2())
                                {
                                    InitTaskData_SOCKET("开线点检;设备点检");
                                    isStartLineCheck = true;
                                }
                                else
                                {
                                    InitTaskData_SOCKET("开线点检");
                                    isStartLineCheck = true;
                                }

                            }
                            isInitChecklist = true;
                        }
                        InitSetupGrid();
                        InitWorkOrderList();
                        SetWorkorderGridStatus();
                        InitEquipmentGridEXT();
                        InitDocumentGrid();
                        InitIPIWorkOrderType();
                        SetTipMessage(MessageType.OK, Message("msg_Activated work order success."));
                        InitShift(txbCDAMONumber.Text);
                        if (config.IsNeedTransWO == "Y")
                        {
                            InitGetHasTransWO();
                        }
                        ReadRestoreFile();
                        //this.BeginInvoke(new Action(() =>
                        //{
                        //    InitIPIStatus();
                        //}));
                        if (!isInitChecklist)
                        {
                            if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
                            {
                                InitShift2(WorkorderPre);
                                if (!CheckShiftChange2())
                                {
                                    InitTaskData_SOCKET("开线点检;设备点检");
                                }
                                else
                                {
                                    if (!ReadCheckListFile())
                                    {
                                        InitTaskData_SOCKET("开线点检");
                                        isStartLineCheck = true;
                                    }
                                }
                            }
                            else
                            {
                                InitTaskData();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.Debug(ex.Message, ex);
            }


        }
        private void btnActivateExt_Click(object sender, EventArgs e)
        {
            //if (!VerifyCheckList())
            //{
            //    return;
            //}
            try
            {
                string workorder = "";
                strShiftChecklist = "";
                bool isInitChecklist = false;
                string WorkorderPre = txbCDAMONumber.Text;
                int processlayerPre = initModel.currentSettings == null ? -1 : initModel.currentSettings.processLayer;
                if (this.gridWorkorder.SelectedRows.Count > 0)
                {
                    workorder = this.gridWorkorder.SelectedRows[0].Cells["columnWoNumber"].Value.ToString();
                    ActivateWorkorder activateHandler = new ActivateWorkorder(sessionContext, initModel, this);
                    int error = activateHandler.ActivateWorkorderExt(initModel.configHandler.StationNumber, workorder, 2, ConvertProcessLayerToString(this.cmbLayer.Text));//2 = Activate work order for entire line
                    if (error == 0)
                    {
                        this.txbCDAMONumber.Text = this.gridWorkorder.SelectedRows[0].Cells["columnWoNumber"].Value.ToString();
                        this.txbCDAPartNumber.Text = this.gridWorkorder.SelectedRows[0].Cells["columnPn"].Value.ToString();
                        GetCurrentWorkorder getCurrentHandler = new GetCurrentWorkorder(sessionContext, initModel, this);
                        GetStationSettingModel model = getCurrentHandler.GetCurrentWorkorderResultCall();
                        initModel.currentSettings = model;
                        if (model != null && model.workorderNumber != null)
                        {
                            GetNumbersOfSingleBoards getNumBoard = new GetNumbersOfSingleBoards(sessionContext, initModel, this);
                            List<MdataGetPartData> listData = getNumBoard.GetNumbersOfSingleBoardsResultCall(model.partNumber);
                            if (listData != null && listData.Count > 0)
                            {
                                MdataGetPartData mData = listData[0];
                                initModel.numberOfSingleBoards = mData.quantityMultipleBoard;
                            }
                        }
                        LoadYield();

                        if (workorder != WorkorderPre || processlayerPre != initModel.currentSettings.processLayer)
                        {
                            strShiftChecklist = "";
                            InitWorkOrderType();
                            InitShift2(WorkorderPre);
                            if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
                            {
                                if (!CheckShiftChange2())
                                {
                                    InitTaskData_SOCKET("开线点检;设备点检");
                                    isStartLineCheck = true;
                                }
                                else
                                {
                                    InitTaskData_SOCKET("开线点检");
                                    isStartLineCheck = true;
                                }
                            }
                            isInitChecklist = true;
                        }
                        InitSetupGrid();
                        InitWorkOrderList();
                        SetWorkorderGridStatus();
                        InitEquipmentGridEXT();
                        InitDocumentGrid();
                        InitIPIWorkOrderType();
                        SetTipMessage(MessageType.OK, Message("msg_Activated work order success."));
                        if (config.IsNeedTransWO == "Y")
                        {
                            InitGetHasTransWO();
                        }
                        InitShift(txbCDAMONumber.Text);
                        ReadRestoreFile();
                        //this.BeginInvoke(new Action(() =>
                        //{
                        //    InitIPIStatus();
                        //}));
                        if (!isInitChecklist)
                        {
                            if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
                            {
                                InitShift2(WorkorderPre);
                                if (!CheckShiftChange2())
                                {
                                    InitTaskData_SOCKET("开线点检;设备点检");
                                    isStartLineCheck = true;
                                }
                                else
                                {
                                    if (!ReadCheckListFile())
                                    {
                                        InitTaskData_SOCKET("开线点检");
                                        isStartLineCheck = true;
                                    }
                                }
                            }
                            else
                            {
                                InitTaskData();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.Debug(ex.Message, ex);
            }

        }
        private void button1_Click(object sender, EventArgs e)
        {
            //if (!VerifyCheckList())
            //{
            //    return;
            //}
            try
            {
                string WorkorderPre = txbCDAMONumber.Text;
                int processlayerPre = initModel.currentSettings == null ? -1 : initModel.currentSettings.processLayer;
                strShiftChecklist = "";//20161215 add by qy
                GetCurrentWorkorder getCurrentHandler = new GetCurrentWorkorder(sessionContext, initModel, this);
                GetStationSettingModel model = getCurrentHandler.GetCurrentWorkorderResultCall();
                initModel.currentSettings = model;
                if (model != null && model.workorderNumber != null)
                {
                    GetNumbersOfSingleBoards getNumBoard = new GetNumbersOfSingleBoards(sessionContext, initModel, this);
                    List<MdataGetPartData> listData = getNumBoard.GetNumbersOfSingleBoardsResultCall(model.partNumber);
                    if (listData != null && listData.Count > 0)
                    {
                        MdataGetPartData mData = listData[0];
                        initModel.numberOfSingleBoards = mData.quantityMultipleBoard;
                    }
                    this.txbCDAMONumber.Text = model.workorderNumber;
                    this.txbCDAPartNumber.Text = model.partNumber;
                    LoadYield();

                    bool isInitChecklist = false;
                    if (model.workorderNumber != WorkorderPre || processlayerPre != initModel.currentSettings.processLayer)
                    {
                        strShiftChecklist = "";
                        InitWorkOrderType();
                        InitShift2(WorkorderPre);//20161215 add by qy
                        if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
                        {
                            if (!CheckShiftChange2())
                            {
                                InitTaskData_SOCKET("开线点检;设备点检");
                                isStartLineCheck = true;
                            }
                            else
                            {
                                InitTaskData_SOCKET("开线点检");
                                isStartLineCheck = true;
                            }
                        }
                        isInitChecklist = true;
                    }

                    InitSetupGrid();
                    InitWorkOrderList();
                    SetWorkorderGridStatus();
                    InitEquipmentGridEXT();
                    InitDocumentGrid();
                    InitIPIWorkOrderType();
                    ShowTopWindow();
                    if (config.IsNeedTransWO == "Y")
                    {
                        InitGetHasTransWO();
                    }
                    InitShift(txbCDAMONumber.Text);
                    ReadRestoreFile();
                    //this.BeginInvoke(new Action(() =>
                    //{
                    //    InitIPIStatus();
                    //}));
                    if (!isInitChecklist)
                    {
                        if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
                        {
                            InitShift2(WorkorderPre);
                            if (!CheckShiftChange2())
                            {
                                InitTaskData_SOCKET("开线点检;设备点检");
                                isStartLineCheck = true;
                            }
                            else
                            {
                                if (!ReadCheckListFile())
                                {
                                    InitTaskData_SOCKET("开线点检");
                                    isStartLineCheck = true;
                                }
                            }
                        }
                        else
                        {
                            InitTaskData();
                        }
                    }
                }
                else
                {
                    this.txbCDAMONumber.Text = "";
                    this.txbCDAPartNumber.Text = "";
                }
            }
            catch (Exception ex)
            {
                LogHelper.Debug(ex.Message, ex);
            }

        }
        private void InitWorkOrderList()
        {
            GetWorkOrder getWOHandler = new GetWorkOrder(sessionContext, initModel, this);
            DataTable dt = getWOHandler.GetAllWorkordersExt();
            DataView dv = dt.DefaultView;
            dv.Sort = "Info desc";
            dt = dv.Table;
            if (dt != null)
            {
                this.gridWorkorder.DataSource = dt;
                this.gridWorkorder.ClearSelection();

                if (config.IsNeedTransWO == "Y")
                {
                    DataTable transDT = new DataTable();
                    transDT.Columns.Add(new DataColumn("WONumber", typeof(string)));
                    foreach (DataGridViewRow row in gridWorkorder.Rows)
                    {
                        if (row.Cells["columnPn"].Value.ToString() == this.txbCDAPartNumber.Text)
                        {
                            DataRow transrow = transDT.NewRow();
                            transrow["WONumber"] = row.Cells["columnWoNumber"].Value.ToString();
                            transDT.Rows.Add(transrow);
                        }
                    }
                    this.txtTransWO.DataSource = transDT;
                    this.txtTransWO.SelectedIndex = -1;
                    this.txtTransWO.DisplayMember = "WONumber";
                    this.txtTransWO.ValueMember = "WONumber";
                    this.txtTransWO.Text = "";
                }
            }
            for (int i = 0; i < gridWorkorder.Rows.Count; i++)
            {
                gridWorkorder.Rows[i].Cells["columnRunId"].Value = i + 1 + "";
            }
        }

        private delegate void SetWorkorderGridStatusHandle();
        private void SetWorkorderGridStatus()
        {
            if (this.gridWorkorder.InvokeRequired)
            {
                SetWorkorderGridStatusHandle setStatusDel = new SetWorkorderGridStatusHandle(SetWorkorderGridStatus);
                Invoke(setStatusDel, new object[] { });
            }
            else
            {
                for (int i = 0; i < this.gridWorkorder.Rows.Count; i++)
                {
                    if (this.txbCDAMONumber.Text.Trim() == this.gridWorkorder.Rows[i].Cells["columnWoNumber"].Value.ToString())
                    {
                        ((DataGridViewImageCell)gridWorkorder.Rows[i].Cells["Activated"]).Value = ScreenPrinter.Properties.Resources.ok;
                    }
                    else
                    {
                        ((DataGridViewImageCell)gridWorkorder.Rows[i].Cells["Activated"]).Value = ScreenPrinter.Properties.Resources.Close;
                    }
                }
            }
        }

        #endregion

        #region IPI
        public int IPIStatus = 0; //0 start process, -1 fail, 1 pass
        private void InitIPIStatus()
        {
            if (config.IPI_STATUS_CHECK == "ENABLE")
            {
                GetAttributeValue getAttriHandler = new GetAttributeValue(sessionContext, initModel, this);
                string[] valuesAttri = getAttriHandler.GetAttributeValueForAll(1, this.txbCDAMONumber.Text, "-1", "IPI_STATUS_WS_" + worktepnumber);
                if (valuesAttri != null && valuesAttri.Length > 0)
                {
                    string status = valuesAttri[1];
                    if (status == "0")
                    {
                        IPIStatus = 0;
                        IPITimerStart();
                        IPIFormPoup();
                    }
                    else if (status == "-1")
                    {
                        IPIStatus = -1;
                        IPITimerStart();
                        IPIFormPoup();
                    }
                    else if (status == "-2")
                    {
                        IPIStatus = -2;
                        IPITimerStart();
                        IPIFormPoup();
                    }
                    else
                    {
                        IPIStatus = 1;
                        errorHandler(0, Message("msg_IPI SUCCESS"), "");
                    }
                }
                else
                {
                    IPIStatus = 0;
                    errorHandler(0, Message("msg_scan sn to IPI"), "");
                }
            }

        }
        private void SetIPI(string serialnumber)
        {
            if (config.IPI_STATUS_CHECK == "ENABLE")
            {
                if (IPIStatus == 0)
                {
                    GetSerialNumberInfo getHandle = new GetSerialNumberInfo(sessionContext, initModel, this);
                    int error = getHandle.GetSerialNumberByref(serialnumber);
                    if (error == -203)
                        serialnumber = serialnumber.Substring(0, serialnumber.Length - 3);
                    AppendAttribute appendAttri = new AppendAttribute(sessionContext, initModel, this);
                    appendAttri.AppendAttributeForAll(1, this.txbCDAMONumber.Text, "-1", "IPI_STATUS", "0");
                    appendAttri.AppendAttributeForAll(0, serialnumber, "-1", "IPI", "Y");
                }
            }

        }
        public void RemoveIPI()
        {
            RemoveAttributeValue removeAttri = new RemoveAttributeValue(sessionContext, initModel, this);
            removeAttri.RemoveAttributeForAll(1, this.txbCDAMONumber.Text, "-1", "IPI_STATUS_WS_" + worktepnumber);
        }

        IPIForm IPIform = null;
        private void IPIFormPoup()
        {
            //if (IPIform != null)
            //{
            //    IPIform.Hide();
            //}
            //cSocket.CloseSocket();
            //CheckConnectTimer.Stop();
            IPIform = new IPIForm(IPIStatus.ToString(), this);
            IPIform.ShowDialog(this);
        }
        private void IPIFormPoupFirst()
        {
            IPIform = new IPIForm(IPIStatus.ToString(), this);
            IPIform.ShowDialog(this);
        }

        public bool VerifyIPIStatus()
        {
            bool isValid = true;
            if (config.IPI_STATUS_CHECK == "ENABLE")
            {
                GetAttributeValue getAttriHandler = new GetAttributeValue(sessionContext, initModel, this);
                string[] valuesAttri = getAttriHandler.GetAttributeValueForAll(1, this.txbCDAMONumber.Text, "-1", "IPI_STATUS_WS_" + worktepnumber);
                if (valuesAttri != null && valuesAttri.Length > 0)
                {
                    string status = valuesAttri[1];
                    if (status == "0") //首件检查开始
                    {
                        IPIStatus = 0;
                        IPITimerStart();
                        isValid = false;
                    }
                    if (status == "1") //首件检查成功
                    {
                        IPIStatus = 1;
                        isValid = true;
                    }
                    if (status == "-1") //首件检查失败
                    {
                        IPIStatus = -1;
                        IPITimerStart();
                        isValid = false;
                    }
                    if (status == "2") //生产检查开始
                    {
                        IPIStatus = 2;
                        IPITimerStart();
                        isValid = true;
                    }
                    if (status == "3") //生产检查成功
                    {
                        IPIStatus = 3;
                        IPITimerStart();
                        isValid = true;
                    }
                    if (status == "-3") //生产检查失败
                    {
                        IPIStatus = -3;
                        IPITimerStart();
                        isValid = false;
                    }
                    if (status == "4") //末件检查开始
                    {
                        IPIStatus = 4;
                        IPITimerStart();
                        isValid = true;
                    }
                    if (status == "5") //末件检查成功
                    {
                        IPIStatus = 5;
                        IPITimerStart();
                        isValid = true;
                    }
                    if (status == "-5") //末件检查失败
                    {
                        IPIStatus = -5;
                        IPITimerStart();
                        isValid = false;
                    }
                }
                else
                {
                    IPIStatus = 0;
                }
            }

            return isValid;
        }

        public void IPITimerStart()
        {
            if (CheckIPITimer != null && CheckIPITimer.Enabled)
                return;
            CheckIPITimer = new System.Timers.Timer();
            // 循环间隔时间(1分钟)
            CheckIPITimer.Interval = Convert.ToInt32(config.IPI_STATUS_CHECK_INTERVAL) * 1000;
            // 允许Timer执行
            CheckIPITimer.Enabled = true;
            // 定义回调
            CheckIPITimer.Elapsed += new ElapsedEventHandler(CheckIPITimer_Elapsed);
            // 定义多次循环
            CheckIPITimer.AutoReset = true;
        }
        public void IPITimerStop()
        {
            CheckIPITimer.Stop();
        }
        
        private void CheckIPITimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            LogHelper.Debug("IPI Timer:" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            GetAttributeValue getAttriHandler = new GetAttributeValue(sessionContext, initModel, this);
            string[] valuesAttri = getAttriHandler.GetAttributeValueForAll(1, this.txbCDAMONumber.Text, "-1", "IPI_STATUS_WS_" + worktepnumber);
            if (valuesAttri != null && valuesAttri.Length > 0)
            {
                string status = valuesAttri[1];
                this.Invoke(new MethodInvoker(delegate
               {
                   if (status == "0")
                   {
                       IPIStatus = 0;
                       if (IPIform != null && IPIform.Visible)
                       {
                           IPIform.Text = Message("msg_IPI form title");
                           IPIform.lbltime.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                           LogHelper.Info(Message("msg_Initial Product Inspection in process"));
                           IPIform.panel1.BackColor = Color.LemonChiffon;
                           IPIform.btnAllowProduction.Visible = false;
                       }
                   }
                   if (status == "1")
                   {
                       IPIStatus = 1;
                       if (IPIform != null && IPIform.Visible)
                       {
                           IPIform.Text = Message("msg_IPI form title");
                           IPIform.lblIPITitle.Text = Message("msg_Initial Product Inspection") + Message("msg_SUCCESS");
                           IPIform.lbltime.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                           LogHelper.Info(Message("msg_Initial Product Inspection") + Message("msg_SUCCESS"));
                           IPIform.panel1.BackColor = Color.LightGreen;
                           IPIform.btnOk.Visible = true;
                           IPIform.btnAllowProduction.Visible = false;
                       }
                   }
                   if (status == "-1")
                   {
                       IPIStatus = -1;
                       if (IPIform != null && IPIform.Visible)
                       {
                           IPIform.Text = Message("msg_IPI form title");
                           IPIform.lblIPITitle.Text = Message("msg_Initial Product Inspection") + Message("msg_FAIL");
                           IPIform.lbltime.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                           LogHelper.Error(Message("msg_Initial Product Inspection") + Message("msg_FAIL"));
                           IPIform.panel1.BackColor = Color.Red;
                           IPIform.btnAnother.Visible = true;
                           IPIform.btnOk.Visible = true;
                           IPIform.btnOk.Enabled = false;
                           IPIform.btnAllowProduction.Visible = false;
                       }
                   }
                   if (status == "-3")
                   {
                       IPIStatus = -2;
                       //if (IPIform != null)
                       //{
                       //    IPIform.Hide();
                       //}
                       //IPIFormPoup();
                       if (IPIform != null && IPIform.Visible)
                       {
                           IPIform.Text = Message("msg_ProductionInspection form title");
                           IPIform.lblIPITitle.Text = Message("msg_Product Inspection") + Message("msg_FAIL");
                           IPIform.lbltime.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                           LogHelper.Error(Message("msg_Product Inspection") + Message("msg_FAIL"));
                           IPIform.panel1.BackColor = Color.Red;
                           IPIform.btnAnother.Visible = false;
                           IPIform.btnOk.Visible = false;
                           IPIform.btnAllowProduction.Visible = true;
                       }
                   }
                   if (status == "5")
                   {
                       IPIStatus = 5;
                       if (IPIform != null && IPIform.Visible)
                       {
                           IPIform.Text = Message("msg_LastProductionInspection form title");
                           IPIform.lblIPITitle.Text = Message("msg_Last Product Inspection") + Message("msg_SUCCESS");
                           IPIform.lbltime.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                           LogHelper.Info(Message("msg_Last Product Inspection") + Message("msg_SUCCESS"));
                           IPIform.panel1.BackColor = Color.LightGreen;
                           IPIform.btnOk.Visible = true;
                           IPIform.btnAllowProduction.Visible = false;
                       }
                   }
                   if (status == "-5")
                   {
                       IPIStatus = -5;
                       if (IPIform != null && IPIform.Visible)
                       {
                           IPIform.Text = Message("msg_LastProductionInspection form title");
                           IPIform.lblIPITitle.Text = Message("msg_Last Product Inspection") + Message("msg_FAIL");
                           IPIform.lbltime.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                           LogHelper.Error(Message("msg_Last Product Inspection") + Message("msg_FAIL"));
                           IPIform.panel1.BackColor = Color.Red;
                           IPIform.btnAnother.Visible = false;
                           IPIform.btnOk.Visible = false;
                           IPIform.btnAllowProduction.Visible = true;
                       }
                   }
               }));
            }
        }

        #region add by qy 160823 首件检查两个面次
        private string GetWorkStepNumberBySN(string serialnumber)
        {
            CheckSerialNumberState getHandle = new CheckSerialNumberState(sessionContext, initModel, this);
            string workstrpnumber = "";
            workstrpnumber = getHandle.GetWorkStepNumberbyWorkPlan(serialnumber);
            return workstrpnumber;
        }
        public bool VerifyIPIStatusEXT(string SN, string workstepnumber)
        {
            bool isValid = true;
            if (config.IPI_WORKORDERTYPE_CHECK.Contains(workorderType))
            {
                return true;
            }
            if (config.IPI_STATUS_CHECK == "ENABLE")
            {
                GetAttributeValue getAttriHandler = new GetAttributeValue(sessionContext, initModel, this);
                string[] valuesAttri = getAttriHandler.GetAttributeValueForAll(1, this.txbCDAMONumber.Text, "-1", "IPI_STATUS_WS_" + workstepnumber);
                if (valuesAttri != null && valuesAttri.Length > 0)
                {
                    string status = valuesAttri[1];
                    if (status == "0") //首件检查开始
                    {
                        IPIStatus = 0;
                        IPITimerStart();
                        isValid = false;
                    }
                    if (status == "1") //首件检查成功
                    {
                        IPIStatus = 1;
                        isValid = true;
                    }
                    if (status == "-1") //首件检查失败
                    {
                        IPIStatus = -1;
                        IPITimerStart();
                        isValid = false;
                    }
                    if (status == "2") //生产检查开始
                    {
                        IPIStatus = 2;
                        IPITimerStart();
                        isValid = true;
                    }
                    if (status == "3") //生产检查成功
                    {
                        IPIStatus = 3;
                        IPITimerStart();
                        isValid = true;
                    }
                    if (status == "-3") //生产检查失败
                    {
                        IPIStatus = -3;
                        IPITimerStart();
                        isValid = false;
                    }
                    if (status == "4") //末件检查开始
                    {
                        IPIStatus = 4;
                        IPITimerStart();
                        isValid = true;
                    }
                    if (status == "5") //末件检查成功
                    {
                        IPIStatus = 5;
                        IPITimerStart();
                        isValid = true;
                    }
                    if (status == "-5") //末件检查失败
                    {
                        IPIStatus = -5;
                        IPITimerStart();
                        isValid = false;
                    }
                }
                else
                {
                    IPIStatus = 0;
                }
            }

            return isValid;
        }

        private void SetIPIEXT(string serialnumber, string workstepnumber)
        {
            if (config.IPI_WORKORDERTYPE_CHECK.Contains(workorderType))
                return;
            if (config.IPI_STATUS_CHECK == "ENABLE")
            {
                if (IPIStatus == 0)
                {
                    GetSerialNumberInfo getHandle = new GetSerialNumberInfo(sessionContext, initModel, this);
                    int error = getHandle.GetSerialNumberByref(serialnumber);
                    if (error == -203)
                        serialnumber = serialnumber.Substring(0, serialnumber.Length - 3);
                    AppendAttribute appendAttri = new AppendAttribute(sessionContext, initModel, this);
                    appendAttri.AppendAttributeForAll(1, this.txbCDAMONumber.Text, "-1", "IPI_STATUS_WS_" + workstepnumber, "0");
                    appendAttri.AppendAttributeForAll(0, serialnumber, "-1", "IPI", workstepnumber);
                }
            }
        }
        #endregion

        #endregion

        #region transfer workorder
        private void btnAllow_Click(object sender, EventArgs e)
        {
            if (this.txbCDAMONumber.Text == "")
            {
                errorHandler(2, Message("msg_no active wo"), "");
                return;
            }
            if (this.txtTransWO.Text.Trim() == "")
            {
                errorHandler(2, Message("msg_please input the transfor workorder"), "");
                return;
            }
            if (this.txtTransWO.Text.Trim() == this.txbCDAMONumber.Text)
            {
                errorHandler(2, Message("msg_trans wo can not same as current wo"), "");
                return;
            }
            if (!VerifyTransWO(this.txtTransWO.Text.Trim()))
            {
                return;
            }
            if (config.LogInType == "COM" && initModel.scannerHandler.handler().IsOpen)
                initModel.scannerHandler.handler().Close();
            LoginForm LogForm = new LoginForm(1, this, "");
            LogForm.ShowDialog();
        }
        public bool AddTransWO()
        {
            string pre_Value = "";
            GetAttributeValue getAttriHandler = new GetAttributeValue(sessionContext, initModel, this);
            string[] valuesAttri = getAttriHandler.GetAttributeValueForAll(1, this.txbCDAMONumber.Text, "-1", "APPROVED_TRANSFER_WO");
            if (valuesAttri != null && valuesAttri.Length > 0)
            {
                pre_Value = valuesAttri[1];
            }

            if (pre_Value == "")
            {
                pre_Value = this.txtTransWO.Text.Trim();
            }
            else
            {
                if (pre_Value.Contains(this.txtTransWO.Text.Trim()))
                {
                    errorHandler(2, Message("msg_this transwo has been added"), "");
                    return false;
                }
                pre_Value = pre_Value + ";" + this.txtTransWO.Text.Trim();
            }

            AppendAttribute appendAttri = new AppendAttribute(sessionContext, initModel, this);
            appendAttri.AppendAttributeForAll(1, this.txbCDAMONumber.Text, "-1", "APPROVED_TRANSFER_WO", pre_Value);

            this.dgvTransWorkOrder.Rows.Insert(0, new object[] { this.txtTransWO.Text.Trim() });

            this.dgvTransWorkOrder.ClearSelection();
            this.txtTransWO.Text = "";
            return true;
        }
        public bool RemoveTransWO(string workorder)
        {
            string pre_Value = "";
            GetAttributeValue getAttriHandler = new GetAttributeValue(sessionContext, initModel, this);
            string[] valuesAttri = getAttriHandler.GetAttributeValueForAll(1, this.txbCDAMONumber.Text, "-1", "APPROVED_TRANSFER_WO");
            if (valuesAttri != null && valuesAttri.Length > 0)
            {
                string[] values = valuesAttri[1].Split(';');
                foreach (var wo in values)
                {
                    if (wo != workorder)
                    {
                        if (pre_Value == "")
                            pre_Value = wo;
                        else
                            pre_Value = pre_Value + ";" + wo;
                    }

                }
                AppendAttribute appendAttri = new AppendAttribute(sessionContext, initModel, this);
                appendAttri.AppendAttributeForAll(1, this.txbCDAMONumber.Text, "-1", "APPROVED_TRANSFER_WO", pre_Value);

                this.dgvTransWorkOrder.ClearSelection();
            }
            else
            {
                return false;
            }

            return true;
        }

        private bool VerifyTransWO(string workorder)
        {
            bool isValid = true;
            GetWorkOrder getHandle = new GetWorkOrder(sessionContext, initModel, this);
            string[] mdataGetWorkordersResultValues = getHandle.GetPartNumberByWorkOrder(workorder);
            if (mdataGetWorkordersResultValues != null && mdataGetWorkordersResultValues.Length > 0)
            {
                string workordersate = mdataGetWorkordersResultValues[1];
                string partnumber = mdataGetWorkordersResultValues[0];
                if (workordersate != "F" && workordersate != "S")
                {
                    errorHandler(2, Message("msg_transworkorder is not valid"), "");
                    isValid = false;
                }
                else
                {
                    if (partnumber != this.txbCDAPartNumber.Text)
                    {
                        errorHandler(2, Message("msg_The transfer wo's part number is not equals the current part number"), "");
                        isValid = false;
                    }
                }

            }
            else
            {
                errorHandler(2, Message("msg_transworkorder is not valid"), "");
                isValid = false;
            }

            return isValid;
        }

        private bool VerifyTransWOExist()
        {
            bool isValid = false;
            foreach (DataGridViewRow row in this.dgvTransWorkOrder.Rows)
            {
                string transWO = row.Cells[0].Value.ToString();
                if (snWorkOrder == transWO)
                {
                    isValid = true;
                    break;
                }
            }

            return isValid;
        }
        public void InitGetHasTransWO()
        {
            this.dgvTransWorkOrder.Rows.Clear();
            if (this.txbCDAMONumber.Text == "")
            {
                errorHandler(2, Message("msg_no active wo"), "");
                return;
            }
            GetAttributeValue getAttriHandler = new GetAttributeValue(sessionContext, initModel, this);
            string[] valuesAttri = getAttriHandler.GetAttributeValueForAll(1, this.txbCDAMONumber.Text, "-1", "APPROVED_TRANSFER_WO");
            if (valuesAttri != null && valuesAttri.Length > 0)
            {
                string[] attributes = valuesAttri[1].Split(';');

                foreach (var wo in attributes)
                {
                    if (wo != "")
                    {
                        this.dgvTransWorkOrder.Rows.Insert(0, new object[] { wo });
                    }
                    this.dgvTransWorkOrder.ClearSelection();
                }

            }
        }

        private void btnRemoveTransWO_Click(object sender, EventArgs e)
        {
            if (this.txbCDAMONumber.Text == "")
            {
                errorHandler(2, Message("msg_no active wo"), "");
                return;
            }

            if (this.dgvTransWorkOrder.SelectedRows.Count > 0)
            {
                DataGridViewRow row = this.dgvTransWorkOrder.SelectedRows[0];
                string WO = row.Cells["LTransferWorkOrder"].Value.ToString();
                if (string.IsNullOrEmpty(WO))
                    return;

                if (config.LogInType == "COM" && initModel.scannerHandler.handler().IsOpen)
                    initModel.scannerHandler.handler().Close();

                LoginForm LogForm = new LoginForm(2, this, WO);
                LogForm.ShowDialog();

                //dgvTransWorkOrder.Rows.Remove(row);
                //RemoveTransWO(WO);
            }
            else
            {
                errorHandler(2, Message("msg_please select trans wo"), "");
            }
        }
        private void txtTransWO_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in this.dgvTransWorkOrder.Rows)
            {
                if (row.Cells[0].Value.ToString() == this.txtTransWO.Text)
                {
                    errorHandler(2, Message("msg_this transwo has been added"), "");
                    break;
                }
            }
        }

        #endregion

        #region production inspection
        string strShift = "";
        bool NextSelectSN = false;
        private void btnPInspection_Click(object sender, EventArgs e)//只要点击了下一块板的按钮,那么下一块板就自动转为生产检查的板
        {
            NextSelectSN = true;
            errorHandler(0, Message("msg_Next sn production inspection"), "");
        }

        private void InitShift(string wo)
        {
            if (config.IsNeedProductionInspection == "ENABLE")
            {
                string path = @"ShiftTemp.txt";
                if (File.Exists(path))
                {
                    string[] content = File.ReadAllLines(path);

                    foreach (var item in content)
                    {
                        if (item != "")
                        {
                            string[] items = item.Split(';');
                            if (items[1] == wo)
                            {
                                strShift = items[0];
                                break;
                            }
                        }

                    }
                }
            }

        }

        private void WriteIntoShift()
        {
            string datetime = DateTime.Now.ToString("yyyy/MM/dd HH:mm");
            strShift = datetime;
            string path = @"ShiftTemp.txt";
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(datetime + ";" + this.txbCDAMONumber.Text);
            FileStream fs = new FileStream(path, FileMode.OpenOrCreate);
            byte[] bt = Encoding.UTF8.GetBytes(sb.ToString());
            fs.Seek(0, SeekOrigin.Begin);
            fs.Write(bt, 0, bt.Length);
            fs.Flush();
            fs.Close();
        }

        //检查有没有到换班时间，如果到换班时间，换班后的第一块板都要做production inspection
        private bool CheckShiftChange()
        {
            try
            {
                bool isValid = false;
                if (strShift == "")
                    return false;

                string[] shifchangetimes = config.SHIFT_CHANGE_TIME.Split(';');
                List<string> shiftList = new List<string>();
                string nowDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm");
                for (int i = 0; i < shifchangetimes.Length; i++)
                {
                    shiftList.Add(DateTime.Now.ToString("yyyy/MM/dd ") + shifchangetimes[i].Substring(0, 2) + ":" + shifchangetimes[i].Substring(2, 2));
                }

                shiftList.Sort();

                for (int j = shiftList.Count - 1; j < shiftList.Count; j--)
                {
                    if (Convert.ToDateTime(nowDate) > Convert.ToDateTime(shiftList[j]))
                    {
                        if (Convert.ToDateTime(strShift) > Convert.ToDateTime(shiftList[j]))
                        {
                            isValid = true;
                        }
                        break;
                    }
                }

                return isValid;
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message, ex);
                return true;
            }

        }

        public void ResetIPIStatus()
        {
            AppendAttribute appendAttri = new AppendAttribute(sessionContext, initModel, this);
            string status = "";
            if (IPIStatus == -3)
                status = "3";
            if (IPIStatus == -5)
                status = "5";
            appendAttri.AppendAttributeForAll(1, this.txbCDAMONumber.Text, "-1", "IPI_STATUS_WS_" + worktepnumber, status);
            errorHandler(0, Message("msg_resume production"), "");
        }

        private void SetIPIProduction(string serialnumber)
        {
            if (config.IsNeedProductionInspection == "ENABLE")
            {
                GetSerialNumberInfo getHandle = new GetSerialNumberInfo(sessionContext, initModel, this);
                int error = getHandle.GetSerialNumberByref(serialnumber);
                if (error == -203)
                    serialnumber = serialnumber.Substring(0, serialnumber.Length - 3);
                AppendAttribute appendAttri = new AppendAttribute(sessionContext, initModel, this);
                appendAttri.AppendAttributeForAll(0, serialnumber, "-1", "IPI", "Y");
            }

        }

        //检测是否为最后一块板
        private bool CheckLastProduct()
        {
            try
            {
                bool isValid = true;
                int wototal = initModel.currentSettings.QuantityMO;
                int pass = Convert.ToInt32(this.lblPass.Text);
                int fail = Convert.ToInt32(this.lblFail.Text);
                int scrap = Convert.ToInt32(this.lblScrap.Text);

                int board = initModel.numberOfSingleBoards; //连扳数
                int total = pass + fail + scrap;
                if ((wototal / board) - (total / board) <= 1)
                {
                    isValid = false;//false 为最后一块
                }

                return isValid;
            }
            catch (Exception ex)
            {
                LogHelper.Debug(ex.Message, ex);
                return true;
            }
        }

        public void OpenScanPort()
        {
            initModel.scannerHandler = new ScannerHeandler(initModel, this);
            initModel.scannerHandler.handler().DataReceived += new SerialDataReceivedEventHandler(DataRecivedHeandler);
            initModel.scannerHandler.handler().Open();
        }

        #region 生产检查两个面次
        private void SetIPIProductionEXT(string serialnumber, string workstepnumber)
        {
            if (config.IsNeedProductionInspection == "ENABLE")
            {
                GetSerialNumberInfo getHandle = new GetSerialNumberInfo(sessionContext, initModel, this);
                int error = getHandle.GetSerialNumberByref(serialnumber);
                if (error == -203)
                    serialnumber = serialnumber.Substring(0, serialnumber.Length - 3);
                AppendAttribute appendAttri = new AppendAttribute(sessionContext, initModel, this);
                appendAttri.AppendAttributeForAll(1, this.txbCDAMONumber.Text, "-1", "IPI_STATUS_WS_" + workstepnumber, "2");
                appendAttri.AppendAttributeForAll(0, serialnumber, "-1", "IPI", workstepnumber);
                IPIStatus = 2;
            }

        }
        private void SetLastIPIProductionEXT(string serialnumber, string workstepnumber)
        {
            if (config.IsNeedProductionInspection == "ENABLE")
            {
                GetSerialNumberInfo getHandle = new GetSerialNumberInfo(sessionContext, initModel, this);
                int error = getHandle.GetSerialNumberByref(serialnumber);
                if (error == -203)
                    serialnumber = serialnumber.Substring(0, serialnumber.Length - 3);
                AppendAttribute appendAttri = new AppendAttribute(sessionContext, initModel, this);
                appendAttri.AppendAttributeForAll(1, this.txbCDAMONumber.Text, "-1", "IPI_STATUS_WS_" + workstepnumber, "4");
                appendAttri.AppendAttributeForAll(0, serialnumber, "-1", "IPI", workstepnumber);
                IPIStatus = 4;
            }

        }
        #endregion
        #endregion

        #region
        string OKlist = "";
        string NGlist = "";
        private void InitFailureMapTable()
        {
            string[] LineList = File.ReadAllLines("CheckResultMappingFile.txt", Encoding.Default);
            foreach (var line in LineList)
            {
                if (!string.IsNullOrEmpty(line))
                {
                    string[] strs = line.Split(new char[] { ';' });
                    if (strs[0] == "OK")
                    {
                        OKlist = OKlist + "," + strs[1];
                    }
                    else
                    {
                        NGlist = NGlist + "," + strs[1];
                    }
                }
            }
        }
        #endregion

        #region restore equipment
        //如果assign 的是设备，就不能上其他设备
        List<string> equipmentList = new List<string>();
        private void InitEquipmentGridEXT()
        {
            if (string.IsNullOrEmpty(this.txbCDAMONumber.Text))
                return;

            EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
            foreach (DataGridViewRow row in dgvEquipment.Rows)
            {
                string equipmentNo = row.Cells["EquipNo"].Value.ToString();
                string equipmentIndex = row.Cells["EquipmentIndex"].Value.ToString();
                if (string.IsNullOrEmpty(equipmentNo))
                    continue;
                int errorCode = eqManager.UpdateEquipmentData(equipmentIndex, equipmentNo, 1);
                RemoveAttributeForEquipment(equipmentNo, equipmentIndex, "attribEquipmentHasRigged");
            }

            this.dgvEquipment.Rows.Clear();
            equipmentList.Clear();
            List<string> equipList = new List<string>();
            Dictionary<string, List<EquipmentEntity>> DiclistEntity = eqManager.GetRequiredEquipmentDataDic(this.txbCDAMONumber.Text);
            if (DiclistEntity != null)
            {
                foreach (var key in DiclistEntity.Keys)
                {
                    List<EquipmentEntity> listEntity = DiclistEntity[key.ToString()];
                    foreach (var item in listEntity)
                    {
                        equipmentList.Add(item.EQUIPMENT_NUMBER);
                        if (equipList.Contains(item.PART_NUMBER))
                            continue;
                        equipList.Add(item.PART_NUMBER);

                        string partdesc = eqManager.GePartDescData(item.PART_NUMBER);
                        this.dgvEquipment.Rows.Add(new object[8] { ScreenPrinter.Properties.Resources.Close, item.PART_NUMBER, partdesc, "", "", "", "", "0" });
                    }
                }
            }
            this.dgvEquipment.ClearSelection();
        }

        public void ProcessEquipmentDataEXT(string equipmentNo)
        {
            if (!VerifyActivatedWO())
                return;
            if (!equipmentList.Contains(equipmentNo))
            {
                string rightequipment = "";
                foreach (var equ in equipmentList)
                {
                    rightequipment += equ + ";";
                }
                errorHandler(2, Message("msg_equ is not belong to the list") + rightequipment.TrimEnd(';'), "");
                return;
            }
            EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
            string[] values = eqManager.GetEquipmentDetailData(equipmentNo);
            string ePartNumber = "";
            string eIndex = "0";
            int uasedcount = 0;
            if (values != null && values.Length > 0)
            {
                uasedcount = Convert.ToInt32(values[5]);
            }

            if (!CheckEquipmentDuplication(values, ref ePartNumber, ref eIndex))
            {
                errorHandler(3, Message("msg_The equipment") + equipmentNo + Message("msg_ has more Available states."), "");
                return;
            }
            if (!CheckEquipmentValid(ePartNumber))
            {
                errorHandler(3, Message("msg_The equipment is invalid"), "");
                return;
            }
            //check equipment number  whether need to setup?
            if (!CheckEquipmentIsExist(ePartNumber))
                return;
            //check equipment number has rigged on others station
            if (CheckEquipmentHasSetup(equipmentNo, eIndex, "attribEquipmentHasRigged"))
            {
                errorHandler(3, Message("msg_The equipment has rigged on others station."), "");
                return;
            }
            string strEquipmentIndex = eIndex;
            //check the equipment whether is cleaning?
            int equicount = 0;
            Match matchStencil = Regex.Match(equipmentNo, config.StencilPrefix);
            if (matchStencil.Success)
            {
                GetAttributeValue getAttribHandler = new GetAttributeValue(sessionContext, initModel, this);
                string[] valuesTest = getAttribHandler.GetAttributeValueForEquipment("TEST_ITEM", equipmentNo, strEquipmentIndex);
                if (valuesTest == null || valuesTest.Length == 0)
                {
                    errorHandler(3, Message("msg_The equipment is not test"), "");
                    return;
                }

                string[] valuesEquip = getAttribHandler.GetAttributeValueForEquipment("STENCIL_LOCK_TIME", equipmentNo, strEquipmentIndex);
                if (valuesEquip != null && valuesEquip.Length > 0)
                {
                    try
                    {
                        //if (Convert.ToDateTime(valuesEquip[1]) > DateTime.Now)
                        //{
                        errorHandler(3, Message("msg_The equipment is cleaning"), "");
                        return;
                        //}
                    }
                    catch (Exception exx)
                    {

                        LogHelper.Error(exx);
                    }
                }
                string[] valuesTest2 = getAttribHandler.GetAttributeValueForEquipment("NEXT_TEST_DATE", equipmentNo, "0");
                DateTime NextTestDate = DateTime.Now;
                if (valuesTest2 != null && valuesTest2.Length != 0)
                    NextTestDate = Convert.ToDateTime(valuesTest2[1]);
                if (NextTestDate < DateTime.Now)
                {
                    errorHandler(3, Message("msg_Stencil need to clean"), "");
                    AppendAttribute appendAttri = new AppendAttribute(sessionContext, initModel, this);
                    appendAttri.AppendAttributeValuesForEquipment("STENCIL_LOCK_TIME", DateTime.Now.AddHours(Convert.ToDouble(config.LockTime)).ToString("yyyy/MM/dd HH:mm:ss"), equipmentNo);
                    return;
                }

                int usagecount = 0;
                int maxcount = 0;
                if (config.ReduceEquType == "1")
                {
                    string[] valuesusage = getAttribHandler.GetAttributeValueForEquipment("USAGE_COUNT", equipmentNo, strEquipmentIndex);
                    if (valuesusage != null && valuesusage.Length != 0)
                    {
                        usagecount = Convert.ToInt32(valuesusage[1]);
                    }
                    string[] valuesmax = getAttribHandler.GetAttributeValueForEquipment("MAX_USAGE", equipmentNo, strEquipmentIndex);
                    if (valuesmax != null && valuesmax.Length != 0)
                    {
                        maxcount = Convert.ToInt32(valuesmax[1]);
                    }
                    equicount = maxcount - usagecount;
                    if (equicount <= 0)
                    {
                        errorHandler(3, Message("msg_Stencil usage count cann't less then 0"), "");
                        return;
                    }
                }
                else
                {
                    if (uasedcount <= 0)
                    {
                        errorHandler(3, Message("msg_Stencil usage count cann't less then 0"), "");
                        return;
                    }
                }
            }
            else
            {
                if (uasedcount <= 0)
                {
                    errorHandler(3, Message("msg_Stencil usage count cann't less then 0"), "");
                    return;
                }
            }
            int errorCode = eqManager.UpdateEquipmentData(strEquipmentIndex, equipmentNo, 0);
            if (errorCode == 0)//1301 Equipment is already set up
            {
                //add attribue command the equipment is uesd
                AppendAttributeForEquipment(equipmentNo, strEquipmentIndex, "attribEquipmentHasRigged");
                EquipmentEntityExt entityExt = eqManager.GetSetupEquipmentData(equipmentNo);
                if (entityExt != null)
                {
                    entityExt.PART_NUMBER = ePartNumber;
                    entityExt.EQUIPMENT_INDEX = strEquipmentIndex;
                    SetEquipmentGridData(entityExt, equicount, values[6]);
                    SetTipMessage(MessageType.OK, Message("msg_Process equipment number ") + equipmentNo + Message("msg_SUCCESS"));
                    SaveEquAndMaterial();
                    if (CheckMaterialSetUp() && CheckEquipmentSetup())
                    {
                        if (config.DataOutputInterface != null && config.DataOutputInterface != "")
                            SendSN(config.LIGHT_CHANNEL_OFF);
                    }
                }
            }
        }

        public void GetRestoreTimerStart()
        {

            if (RestoreMaterialTimer != null && RestoreMaterialTimer.Enabled)
                return;
            RestoreMaterialTimer = new System.Timers.Timer();
            // 循环间隔时间(1分钟)
            RestoreMaterialTimer.Interval = Convert.ToInt32(config.RESTORE_TREAD_TIMER) * 1000;
            // 允许Timer执行
            RestoreMaterialTimer.Enabled = true;
            // 定义回调
            RestoreMaterialTimer.Elapsed += new ElapsedEventHandler(RestoreMaterialTimer_Elapsed);
            // 定义多次循环
            RestoreMaterialTimer.AutoReset = true;
        }

        private void RestoreMaterialTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            SaveEquAndMaterial();
            SaveCheckList();
            InitProductionChecklist();
            InitShiftCheckList();
        }

        private void SaveEquAndMaterial()
        {
            try
            {
                string path = @"restore.txt";
                string datetime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                string equipment = "";
                string material = "";
                StringBuilder sb = new StringBuilder();
                sb.AppendLine(datetime);
                sb.AppendLine(txbCDAMONumber.Text + ";" + initModel.currentSettings.processLayer);
                foreach (DataGridViewRow row in dgvEquipment.Rows)
                {
                    string equipmentNo = row.Cells["EquipNo"].Value.ToString();
                    if (string.IsNullOrEmpty(equipmentNo))
                        continue;
                    equipment += equipmentNo + ";";
                }
                sb.AppendLine(equipment.TrimEnd(';'));
                foreach (DataGridViewRow row in gridSetup.Rows)
                {
                    string materialbin = row.Cells["MaterialBinNo"].Value.ToString();
                    if (string.IsNullOrEmpty(materialbin))
                        continue;
                    material += materialbin + ";";
                }
                sb.AppendLine(material.TrimEnd(';'));
                FileStream fs = new FileStream(path, FileMode.Create);
                byte[] bt = Encoding.UTF8.GetBytes(sb.ToString());
                fs.Seek(0, SeekOrigin.Begin);
                fs.Write(bt, 0, bt.Length);
                fs.Flush();
                fs.Close();
                LogHelper.Info("Save restore file success.");
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message, ex);
            }
        }

        private void ReadRestoreFile()
        {
            string path = @"restore.txt";
            if (File.Exists(path))
            {
                string[] linelist = File.ReadAllLines(path);
                string datetimespan = linelist[0];
                string workorder = linelist[1];
                string equipment = linelist[2];
                string material = linelist[3];

                TimeSpan span = DateTime.Now - Convert.ToDateTime(datetimespan);

                if (span.TotalMinutes > Convert.ToInt32(config.RESTORE_TIME))//判断是否大于10分钟，大于10分钟则不自动上料和设备
                {

                }
                else
                {
                    string[] workorders = workorder.Split(';');
                    if (workorders.Length > 1)
                    {
                        if (workorders[0] == this.txbCDAMONumber.Text)//判断工单是否有变化，无变化则自动上料和设备
                        {
                            if (workorders[1] == initModel.currentSettings.processLayer.ToString())//判断面次是否有变化，无变化则自动上料和设备
                            {
                                bool isOK = false;
                                #region setup equ
                                EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
                                string[] equs = equipment.Split(';');
                                if (equipment.Replace(";", "").Trim() != "")
                                {
                                    foreach (var equipmentNo in equs)
                                    {
                                        string equipmentnumber = equipmentNo.ToString();
                                        if (string.IsNullOrEmpty(equipmentnumber))
                                            continue;
                                        int errorCode = eqManager.UpdateEquipmentData("0", equipmentnumber, 1);
                                        RemoveAttributeForEquipment(equipmentnumber, "0", "attribEquipmentHasRigged");
                                        ProcessEquipmentDataEXT(equipmentnumber);
                                    }
                                    isOK = true;
                                }

                                #endregion

                                #region setup material
                                if (material.Replace(";", "").Trim() != "")
                                {
                                    SetUpManager setupHandler = new SetUpManager(sessionContext, initModel, this);
                                    setupHandler.SetupStateChange(this.txbCDAMONumber.Text, 2);
                                    string[] materials = material.Split(';');
                                    foreach (var materialbin in materials)
                                    {
                                        string materialbinnumber = materialbin.ToString();
                                        if (string.IsNullOrEmpty(materialbinnumber))
                                            continue;
                                        if (config.AutoNextMaterial.ToUpper() == "ENABLE")
                                            ProcessMaterialBinNo(materialbinnumber);
                                        else
                                            ProcessMaterialBinNoEXT(materialbinnumber);
                                    }
                                    isOK = true;
                                }
                                #endregion
                                if (isOK)
                                    errorHandler(0, Message("msg_Material and Equipment setup has been restored"), "");
                            }

                        }
                    }

                }
            }
        }
        #endregion
        private string ConvertProcessLayerToString(string str)
        {
            string iValue = "";
            switch (str)
            {
                case "T":
                    iValue = "0";
                    break;
                case "B":
                    iValue = "1";
                    break;
                default:
                    iValue = "2";
                    break;
            }
            return iValue;
        }

        #region checklist from OA
        bool Supervisor = false;
        bool IPQC = true;
        private void InitTaskData_SOCKET(string djclass)
        {
            try
            {
                string PartNumber = this.txbCDAPartNumber.Text;
                if (PartNumber == "")
                {
                    errorHandler(2, Message("msg_no active wo"), "");
                    return;
                }
                this.Invoke(new MethodInvoker(delegate
                {
                    try
                    {
                        this.dgvCheckListTable.Rows.Clear();

                        Supervisor = false;
                        IPQC = true;
                        GetWorkPlanData handle = new GetWorkPlanData(sessionContext, initModel, this);
                        int firstSN = int.Parse(this.lblPass.Text) + int.Parse(this.lblFail.Text) + int.Parse(this.lblScrap.Text);
                        if (firstSN == 0)
                        {
                            djclass = djclass + ";首末件点检";
                        }

                        string workstep_text = handle.GetWorkStepInfobyWorkPlan(this.txbCDAMONumber.Text, initModel.currentSettings.processLayer);
                        if (workstep_text != "")
                        {
                            GetAttributeValue getAttriHandler = new GetAttributeValue(sessionContext, initModel, this);
                            string[] processCode = getAttriHandler.GetAttributeValueForAll(1, this.txbCDAMONumber.Text, "-1", "TTE_PROCESS_CODE");
                            if (processCode != null && processCode.Length > 0)
                            {
                                string process = processCode[1];
                                string sedmessage = "{getCheckListItem;" + PartNumber + ";" + process + ";[" + workstep_text + "];[" + djclass + "];" + "}";
                                string returnMsg = checklist_cSocket.SendData(sedmessage);

                                if (returnMsg != "" && returnMsg != null)
                                {
                                    string[] values = returnMsg.TrimEnd(';').Replace("{", "").Replace("}", "").Replace("#", "").Split(new string[] { ";" }, StringSplitOptions.None);
                                    string status = values[1];
                                    if (status == "0")//“0” , or “-1” (error)  
                                    {
                                        int seq = 1;
                                        string itemregular = @"\{[^\{\}]+\}"; //@"\[[^\[\]]+\]";
                                        MatchCollection match = Regex.Matches(returnMsg.TrimStart('{').Substring(0, returnMsg.Length - 2), itemregular);
                                        if (match.Count <= 0)
                                        {
                                            errorHandler(2, Message("msg_No checklist data"), "");
                                            return;
                                        }
                                        for (int i = 0; i < match.Count; i++)
                                        {
                                            string data = match[i].ToString().TrimStart('{').TrimEnd('}');
                                            //string[] datas = data.Split(';');
                                            string[] datas = Regex.Split(data, "#!#", RegexOptions.IgnoreCase);
                                            string sourceclass = datas[4];//数据来源
                                            string formno = datas[0];//对应单号
                                            string itemno = datas[1];//机种品号
                                            string itemnname = datas[2];//机种品名
                                            string sbno = datas[5];//设备编号
                                            string sbname = datas[6];//设备名称
                                            string gcno = datas[7];//过程编号
                                            string gcname = datas[8];//过程名称
                                            string lbclass = datas[9];//类别
                                            string djxmname = datas[10];//点检项目
                                            string specvalue = datas[11];//规格值
                                            string djkind = datas[12];//点检类型
                                            string maxvalue = datas[14];//上限值
                                            string minvalue = datas[13];//下限值
                                            string djclase = datas[15];//点检类别
                                            string djversion = datas[3];//版本
                                            string dataclass = datas[16];//状态

                                            object[] objValues = new object[] { seq, djclase, djxmname, gcname, specvalue, "", "", "", djkind, gcno, maxvalue, minvalue, lbclass, sourceclass, formno, itemno, itemnname, sbno, sbname, djversion, dataclass, "" };
                                            this.dgvCheckListTable.Rows.Add(objValues);
                                            seq++;
                                            SetCheckListInputStatusTable();

                                            if (djkind == "判断值")
                                            {
                                                string[] strInputValues = new string[] { "Y", "N" };
                                                DataTable dtInput = new DataTable();
                                                dtInput.Columns.Add("name");
                                                dtInput.Columns.Add("value");
                                                DataRow rowEmpty = dtInput.NewRow();
                                                rowEmpty["name"] = "";
                                                rowEmpty["value"] = "";
                                                dtInput.Rows.Add(rowEmpty);
                                                foreach (var strValues in strInputValues)
                                                {
                                                    DataRow row = dtInput.NewRow();
                                                    row["name"] = strValues;
                                                    row["value"] = strValues;
                                                    dtInput.Rows.Add(row);
                                                }

                                                DataGridViewComboBoxCell ComboBoxCell = new DataGridViewComboBoxCell();
                                                ComboBoxCell.DataSource = dtInput;
                                                ComboBoxCell.DisplayMember = "Name";
                                                ComboBoxCell.ValueMember = "Value";
                                                dgvCheckListTable.Rows[this.dgvCheckListTable.Rows.Count - 1].Cells["tabResult2"] = ComboBoxCell;
                                            }

                                            this.dgvCheckListTable.ClearSelection();
                                        }
                                    }
                                    else
                                    {
                                        string errormsg = values[1];
                                        errorHandler(2, errormsg, "");
                                    }
                                }
                                else
                                {
                                    isOK = checklist_cSocket.connect(config.CHECKLIST_IPAddress, config.CHECKLIST_Port);
                                    returnMsg = checklist_cSocket.SendData(sedmessage);

                                    if (returnMsg != "" && returnMsg != null)
                                    {
                                        string[] values = returnMsg.TrimEnd(';').Replace("{", "").Replace("}", "").Replace("#", "").Split(new string[] { ";" }, StringSplitOptions.None);
                                        string status = values[1];
                                        if (status == "0")//“0” , or “-1” (error)  
                                        {
                                            int seq = 1;
                                            string itemregular = @"\{[^\{\}]+\}";
                                            MatchCollection match = Regex.Matches(returnMsg.TrimStart('{').Substring(0, returnMsg.Length - 2), itemregular);
                                            if (match.Count <= 0)
                                            {
                                                errorHandler(2, Message("msg_No checklist data"), "");
                                                return;
                                            }
                                            for (int i = 0; i < match.Count; i++)
                                            {
                                                string data = match[i].ToString().TrimStart('{').TrimEnd('}');
                                                //string[] datas = data.Split(';');
                                                string[] datas = Regex.Split(data, "#!#", RegexOptions.IgnoreCase);
                                                string sourceclass = datas[4];//数据来源
                                                string formno = datas[0];//对应单号
                                                string itemno = datas[1];//机种品号
                                                string itemnname = datas[2];//机种品名
                                                string sbno = datas[5];//设备编号
                                                string sbname = datas[6];//设备名称
                                                string gcno = datas[7];//过程编号
                                                string gcname = datas[8];//过程名称
                                                string lbclass = datas[9];//类别
                                                string djxmname = datas[10];//点检项目
                                                string specvalue = datas[11];//规格值
                                                string djkind = datas[12];//点检类型
                                                string maxvalue = datas[14];//上限值
                                                string minvalue = datas[13];//下限值
                                                string djclase = datas[15];//点检类别
                                                string djversion = datas[3];//版本
                                                string dataclass = datas[16];//状态

                                                object[] objValues = new object[] { seq, djclase, djxmname, gcname, specvalue, "", "", "", djkind, gcno, maxvalue, minvalue, lbclass, sourceclass, formno, itemno, itemnname, sbno, sbname, djversion, dataclass, "" };
                                                this.dgvCheckListTable.Rows.Add(objValues);
                                                seq++;
                                                SetCheckListInputStatusTable();

                                                if (djkind == "判断值")
                                                {
                                                    string[] strInputValues = new string[] { "Y", "N" };
                                                    DataTable dtInput = new DataTable();
                                                    dtInput.Columns.Add("name");
                                                    dtInput.Columns.Add("value");
                                                    DataRow rowEmpty = dtInput.NewRow();
                                                    rowEmpty["name"] = "";
                                                    rowEmpty["value"] = "";
                                                    dtInput.Rows.Add(rowEmpty);
                                                    foreach (var strValues in strInputValues)
                                                    {
                                                        DataRow row = dtInput.NewRow();
                                                        row["name"] = strValues;
                                                        row["value"] = strValues;
                                                        dtInput.Rows.Add(row);
                                                    }

                                                    DataGridViewComboBoxCell ComboBoxCell = new DataGridViewComboBoxCell();
                                                    ComboBoxCell.DataSource = dtInput;
                                                    ComboBoxCell.DisplayMember = "Name";
                                                    ComboBoxCell.ValueMember = "Value";
                                                    dgvCheckListTable.Rows[this.dgvCheckListTable.Rows.Count - 1].Cells["tabResult2"] = ComboBoxCell;
                                                }

                                                this.dgvCheckListTable.ClearSelection();
                                            }
                                        }
                                        else
                                        {
                                            string errormsg = values[1];
                                            errorHandler(2, errormsg, "");
                                        }
                                    }
                                }
                            }
                            else
                            {
                                errorHandler(2, Message("msg_no TTE_PROCESS_CODE"), "");//20161213 edit by qy
                                return;
                            }
                        }
                        else
                        {
                            errorHandler(2, Message("msg_no workstep text"), "");
                            return;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogHelper.Error(ex.Message, ex);
                    }
                }));

            }
            catch (Exception ex)
            {
                //20161208 edit by qy
                LogHelper.Error(ex.Message, ex);
            }
        }

        private void SetCheckListInputStatusTable()
        {
            foreach (DataGridViewRow row in this.dgvCheckListTable.Rows)
            {
                if (row.Cells["tabdjkind"].Value.ToString() == "判断值")
                {
                    row.Cells["tabResult1"].ReadOnly = true;
                }
                else if (row.Cells["tabdjkind"].Value.ToString() == "输入值" || row.Cells["tabdjkind"].Value.ToString() == "范围值")
                {
                    row.Cells["tabResult2"].ReadOnly = true;
                }
                row.Cells["tabNo"].ReadOnly = true;
                row.Cells["tabStatus"].ReadOnly = true;
            }
        }

        private void btnSupervisor_Click(object sender, EventArgs e)
        {
            if (gridCheckList.RowCount <= 0)
            {
                return;
            }
            if (config.LogInType == "COM" && initModel.scannerHandler.handler().IsOpen)
                initModel.scannerHandler.handler().Close();

            LoginForm LogForm = new LoginForm(4, this, "");
            LogForm.ShowDialog();
        }

        private void btnIPQC_Click(object sender, EventArgs e)
        {
            if (gridCheckList.RowCount <= 0)
            {
                return;
            }
            if (config.LogInType == "COM" && initModel.scannerHandler.handler().IsOpen)
                initModel.scannerHandler.handler().Close();

            LoginForm LogForm = new LoginForm(5, this, "");
            LogForm.ShowDialog();
        }

        public void SupervisorConfirm(string user)//班长确认
        {
            DialogResult dr = MessageBox.Show(Message("msg_produtc or not"), Message("msg_Warning"), MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
            if (dr == DialogResult.Yes)
            {
                Supervisor = true;
                errorHandler(0, Message("msg_supervisor confirm OK"), "");
            }
            else
            {
                Supervisor = false;
                errorHandler(2, Message("msg_supervisor confirm NG"), "");
            }
            if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
            {
                SaveCheckList();
                string result = "N";
                if (Supervisor)
                    result = "Y";
                string endsendmessage = "{updateCheckListResult;1;" + user + ";" + result + ";" + sequece + "}";
                checklist_cSocket.SendData(endsendmessage);
            }

        }

        public void IPQCConfirm(string user)//IPQC巡检
        {
            DialogResult dr = MessageBox.Show(Message("msg_IPQC produtc or not"), Message("msg_Warning"), MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
            if (dr == DialogResult.Yes)
            {
                IPQC = true;
                errorHandler(0, Message("msg_IPQC confirm OK"), "");
            }
            else
            {
                IPQC = false;
                errorHandler(2, Message("msg_IPQC confirm NG"), "");
            }
            if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
            {
                SaveCheckList();
                string result = "N";
                if (Supervisor)
                    result = "Y";
                string endsendmessage = "{updateCheckListResult;2;" + user + ";" + result + ";" + sequece + "}";
                checklist_cSocket.SendData(endsendmessage);
            }
        }

        private void btnAddCheckListTable_Click(object sender, EventArgs e)
        {
            dgvCheckListTable.Rows.Add(new object[] { this.dgvCheckListTable.Rows.Count + 1, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" });

            dgvCheckListTable.Rows[this.dgvCheckListTable.Rows.Count - 1].Cells["tabResult1"].ReadOnly = true;
            dgvCheckListTable.Rows[this.dgvCheckListTable.Rows.Count - 1].Cells["tabNo"].ReadOnly = true;
            dgvCheckListTable.Rows[this.dgvCheckListTable.Rows.Count - 1].Cells["tabStatus"].ReadOnly = true;
            dgvCheckListTable.ClearSelection();
        }
        string sequece = "";
        private void btnConfirmTable_Click(object sender, EventArgs e)
        {
            try
            {

                string PartNumber = this.txbCDAPartNumber.Text;
                if (PartNumber == "")
                {
                    errorHandler(2, Message("msg_no active wo"), "");
                    return;
                }
                foreach (DataGridViewRow row in this.dgvCheckListTable.Rows)
                {
                    if (row.Cells["tabStatus"].Value == null || row.Cells["tabStatus"].Value.ToString() == "")
                    {
                        errorHandler(2, Message("msg_Verify_CheckList"), "");
                        return;
                    }
                }

                string headmessage = "{appendCheckListResult;" + PartNumber;
                string sedmessage = "";
                string date = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                foreach (DataGridViewRow row in this.dgvCheckListTable.Rows)
                {
                    string gdcode = this.txbCDAMONumber.Text;
                    string itemno = PartNumber;
                    string itemname = initModel.currentSettings.partdesc;
                    string gczcode = initModel.configHandler.StationNumber;
                    string gczname = "";
                    string lineclass = "";
                    string lbclass = row.Cells["tablbclass"].Value.ToString();
                    string djxmname = row.Cells["tabdjxmname"].Value.ToString();
                    string specvalue = "";
                    if (row.Cells["tabResult1"].Value.ToString() != "")
                        specvalue = row.Cells["tabResult1"].Value.ToString();
                    else
                        specvalue = row.Cells["tabResult2"].Value.ToString();
                    string djkind = row.Cells["tabdjkind"].Value.ToString();
                    string maxvalues = row.Cells["tabmaxvalue"].Value.ToString();
                    string minvalues = row.Cells["tabminvalue"].Value.ToString();
                    string djclass = row.Cells["tabdjclass"].Value.ToString();
                    string djversion = row.Cells["tabdjversion"].Value.ToString();
                    string djuser = lblUser.Text;
                    string djremark = "";
                    string djdate = date;
                    string jcuser = lblUser.Text;
                    string qruser = "";
                    string pguser = "";

                    string msgrow = "{" + gdcode + "#!#" + itemno + "#!#" + itemname + "#!#" + gczcode + "#!#" + gczname + "#!#" + lineclass + "#!#" + lbclass + "#!#" + djxmname + "#!#" + specvalue + "#!#" + djkind + "#!#" + maxvalues + "#!#" + minvalues + "#!#" + djclass + "#!#" + djversion + "#!#" + djuser + "#!#" + djremark + "#!#" + djdate + "#!#" + jcuser + "#!#" + qruser + "#!#" + pguser + "}";
                    if (sedmessage == "")
                        sedmessage = msgrow;
                    else
                        sedmessage = sedmessage + ";" + msgrow;
                }
                if (sedmessage == "")
                {
                    errorHandler(2, Message("smg_No checklist data"), "");
                    return;
                }
                string endsendmessage = headmessage + ";" + sedmessage + "}";
                string returnMsg = checklist_cSocket.SendData(endsendmessage);
                if (returnMsg != null && returnMsg != "")
                {
                    returnMsg = returnMsg.TrimStart('{').TrimEnd('}');
                    string[] Msgs = returnMsg.Split(';');
                    if (Msgs[1] == "0")
                    {
                        if (Supervisor_OPTION == "1")
                        {
                            Supervisor = true;
                            errorHandler(0, Message("msg_Send_CheckList_Success"), "");
                        }
                        else
                        {
                            errorHandler(0, Message("msg_Send_CheckList_Success,please supervisor confirm"), "");
                        }

                        sequece = Msgs[3];
                        SaveCheckList();
                        WriteIntoShift2();
                        InitShift2(txbCDAMONumber.Text);
                    }
                    else
                    {
                        errorHandler(2, Message("msg_Send_CheckList_fail"), "");
                    }
                }
                else
                {
                    isOK = checklist_cSocket.connect(config.CHECKLIST_IPAddress, config.CHECKLIST_Port);
                    returnMsg = checklist_cSocket.SendData(endsendmessage);
                    if (returnMsg != null && returnMsg != "")
                    {
                        returnMsg = returnMsg.TrimStart('{').TrimEnd('}');
                        string[] Msgs = returnMsg.Split(';');
                        if (Msgs[1] == "0")
                        {
                            if (Supervisor_OPTION == "1")
                            {
                                Supervisor = true;
                                errorHandler(0, Message("msg_Send_CheckList_Success"), "");
                            }
                            else
                            {
                                errorHandler(0, Message("msg_Send_CheckList_Success,please supervisor confirm"), "");
                            }

                            sequece = Msgs[3];
                            SaveCheckList();
                            WriteIntoShift2();
                            InitShift2(txbCDAMONumber.Text);
                        }
                        else
                        {
                            errorHandler(2, Message("msg_Send_CheckList_fail"), "");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //20161208 edit by qy
                LogHelper.Error(ex.Message, ex);
            }
        }

        private void btnSupervisorTable_Click(object sender, EventArgs e)
        {
            if (sequece == "")
            {
                return;
            }
            if (config.LogInType == "COM" && initModel.scannerHandler.handler().IsOpen)
                initModel.scannerHandler.handler().Close();

            LoginForm LogForm = new LoginForm(4, this, "");
            LogForm.ShowDialog();
        }

        private void btnIPQCTable_Click(object sender, EventArgs e)
        {
            if (sequece == "")
            {
                return;
            }
            if (config.LogInType == "COM" && initModel.scannerHandler.handler().IsOpen)
                initModel.scannerHandler.handler().Close();

            LoginForm LogForm = new LoginForm(5, this, "");
            LogForm.ShowDialog();
        }
        private void dgvCheckListTable_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //20161208 edit by qy
            try
            {
                if (e.RowIndex == -1)
                    return;
                if (this.dgvCheckListTable.Columns[e.ColumnIndex].Name == "tabResult1" && this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabResult1"].Value.ToString() != "")
                {
                    if (this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabdjkind"].Value.ToString() == "范围值")
                    {
                        //verify the input value
                        string strRegex = @"^\d{0,9}\.\d{0,9}|-\d{0,9}\.\d{0,9}";//@"^(\d{0,9}.\d{0,9})～(\d{0,9}.\d{0,9}).*$";"^(\-|\+?\d{0,9}.\d{0,9})～(\-|\+?\d{0,9}.\d{0,9})$"
                        string strResult1 = this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabResult1"].Value.ToString();
                        string strStandard = this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabspecname"].Value.ToString().Replace("（", "").Replace("）", "");
                        string strMax = this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabmaxvalue"].Value.ToString();
                        string strMin = this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabminvalue"].Value.ToString();
                        Match match1 = Regex.Match(strMax, strRegex);
                        Match match2 = Regex.Match(strMin, strRegex);
                        if (match1.Success && match2.Success)
                        {
                            //if (match.Groups.Count > 2)
                            //{
                            //double iMin = Convert.ToDouble(match.Groups[1].Value);
                            //double iMax = Convert.ToDouble(match.Groups[2].Value);
                            double iMin = Convert.ToDouble(match2.ToString());
                            double iMax = Convert.ToDouble(match1.ToString());
                            double iResult = Convert.ToDouble(strResult1);
                            if (iResult >= iMin && iResult <= iMax)
                            {
                                this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Style.BackColor = Color.FromArgb(0, 192, 0);
                                this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Value = "OK";
                            }
                            else
                            {
                                this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Style.BackColor = Color.Red;
                                this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Value = "NG";
                            }
                            //}
                        }
                        else
                        {
                            this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Style.BackColor = Color.Red;
                            this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Value = "NG";
                        }
                    }
                    else if (this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabdjkind"].Value.ToString() == "输入值")
                    {
                        if (this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabResult1"].Value.ToString() == this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabspecname"].Value.ToString())
                        {
                            this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Style.BackColor = Color.FromArgb(0, 192, 0);
                            this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Value = "OK";
                        }
                        else
                        {
                            this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Style.BackColor = Color.Red;
                            this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Value = "NG";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.Debug(ex.Message, ex);
            }

        }

        private void dgvCheckListTable_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            DataGridView dgv = sender as DataGridView;
            if (dgv.CurrentCell.GetType().Name == "DataGridViewComboBoxCell" && dgv.CurrentCell.RowIndex != -1)
            {
                iRowIndex = dgv.CurrentCell.RowIndex;
                (e.Control as ComboBox).SelectedIndexChanged += new EventHandler(ComboBoxTable_SelectedIndexChanged);
            }
        }

        public void ComboBoxTable_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox combox = sender as ComboBox;
            combox.Leave += new EventHandler(comboxtable_Leave);
            try
            {
                if (combox.SelectedItem != null && combox.Text != "")
                {
                    if (OKlist.Contains(combox.Text))
                    {
                        this.dgvCheckListTable.Rows[iRowIndex].Cells["tabStatus"].Style.BackColor = Color.FromArgb(0, 192, 0);
                        this.dgvCheckListTable.Rows[iRowIndex].Cells["tabStatus"].Value = "OK";
                    }
                    else
                    {
                        this.dgvCheckListTable.Rows[iRowIndex].Cells["tabStatus"].Style.BackColor = Color.Red;
                        this.dgvCheckListTable.Rows[iRowIndex].Cells["tabStatus"].Value = "NG";
                    }
                }
                else
                {
                    this.dgvCheckListTable.Rows[iRowIndex].Cells["tabStatus"].Style.BackColor = Color.White;
                    this.dgvCheckListTable.Rows[iRowIndex].Cells["tabStatus"].Value = "";
                }
                Thread.Sleep(100);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public void comboxtable_Leave(object sender, EventArgs e)
        {
            ComboBox combox = sender as ComboBox;
            combox.SelectedIndexChanged -= new EventHandler(ComboBoxTable_SelectedIndexChanged);
        }

        int iIndexCheckListTable = -1;
        private void dgvCheckListTable_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (this.dgvCheckListTable.Rows.Count == 0)
                    return;
                ((DataGridView)sender).CurrentRow.Selected = true;
                iIndexCheckListTable = ((DataGridView)sender).CurrentRow.Index;
                this.dgvCheckListTable.ContextMenuStrip = contextMenuStrip2;

                if (iIndexCheckListTable == -1)
                    this.dgvCheckListTable.ContextMenuStrip = null;

            }
        }
        private void dgvCheckListTable_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            //if (e.Button == MouseButtons.Right)
            //{
            //    if (this.dgvCheckListTable.Rows.Count == 0)
            //        return;

            //    if (e.RowIndex == -1)
            //    {
            //        this.dgvCheckListTable.ContextMenuStrip = null;
            //        return;
            //    }

            //    iIndexCheckListTable = ((DataGridView)sender).CurrentRow.Index;
            //    this.dgvCheckListTable.ContextMenuStrip = contextMenuStrip2;
            //    ((DataGridView)sender).CurrentRow.Selected = true;
            //}
        }

        private void SaveCheckList()
        {
            try
            {
                string path = @"CheckList.txt";
                string datetime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                StringBuilder sb = new StringBuilder();
                sb.AppendLine(datetime);
                sb.AppendLine(txbCDAMONumber.Text + ";" + initModel.currentSettings.processLayer);
                sb.AppendLine(Supervisor.ToString());
                sb.AppendLine(IPQC.ToString());
                sb.AppendLine(sequece);
                foreach (DataGridViewRow row in dgvCheckListTable.Rows)
                {
                    string sourceclass = row.Cells["tabsourceclass"].Value.ToString();//数据来源
                    string formno = row.Cells["tabformno"].Value.ToString();//对应单号
                    string itemno = row.Cells["tabitemno"].Value.ToString();//机种品号
                    string itemnname = row.Cells["tabitemname"].Value.ToString();//机种品名
                    string sbno = row.Cells["tabsbno"].Value.ToString();//设备编号
                    string sbname = row.Cells["tabsnname"].Value.ToString();//设备名称
                    string gcno = row.Cells["tabgcno"].Value.ToString();//过程编号
                    string gcname = row.Cells["tabgcname"].Value.ToString();//过程名称
                    string lbclass = row.Cells["tablbclass"].Value.ToString();//类别
                    string djxmname = row.Cells["tabdjxmname"].Value.ToString();//点检项目
                    string specvalue = row.Cells["tabspecname"].Value.ToString();//规格值
                    string result1 = row.Cells["tabResult1"].Value.ToString();
                    string result2 = row.Cells["tabResult2"].Value == null ? "" : row.Cells["tabResult2"].Value.ToString();// row.Cells["tabResult2"].Value.ToString();
                    string status = row.Cells["tabstatus"].Value.ToString();//结果
                    string djkind = row.Cells["tabdjkind"].Value.ToString();//点检类型
                    string maxvalue = row.Cells["tabmaxvalue"].Value.ToString();//上限值
                    string minvalue = row.Cells["tabminvalue"].Value.ToString();//下限值
                    string djclase = row.Cells["tabdjclass"].Value.ToString();//点检类别
                    string djversion = row.Cells["tabdjversion"].Value.ToString();//版本
                    string dataclass = row.Cells["tabdataclass"].Value.ToString();//状态

                    //string cell13 = row.Cells[13].Value == null ? "" : row.Cells[13].Value.ToString();
                    string linedata = sourceclass + "￥" + formno + "￥" + itemno + "￥" + itemnname + "￥" + sbno + "￥" + sbname + "￥" + gcno + "￥" + gcname + "￥" + lbclass + "￥" + djxmname + "￥" + specvalue + "￥" + result1 + "￥" + result2 + "￥" + status + "￥" + djkind + "￥" + maxvalue + "￥" + minvalue + "￥" + djclase + "￥" + djversion + "￥" + dataclass;
                    //string linedata = row.Cells[1].Value.ToString() + ";" + row.Cells[2].Value.ToString() + ";" + row.Cells[3].Value.ToString() + ";" + row.Cells[4].Value.ToString() + ";" + row.Cells[5].Value.ToString() + ";" + row.Cells[6].Value.ToString() + ";" + row.Cells[7].Value.ToString() + ";" + row.Cells[8].Value.ToString() + ";" + row.Cells[9].Value.ToString() + ";" + row.Cells[10].Value.ToString() + ";" + row.Cells[11].Value.ToString() + ";" + row.Cells[12].Value.ToString() + ";" + cell13 + ";" + row.Cells[14].Value.ToString() + ";" + row.Cells[15].Value.ToString() + ";" + row.Cells[16].Value.ToString() + ";" + row.Cells[17].Value.ToString() + ";" + row.Cells[18].Value.ToString() + ";" + row.Cells[19].Value.ToString() + ";" + row.Cells[20].Value.ToString() + ";" + djkind;
                    sb.AppendLine(linedata);
                }

                FileStream fs = new FileStream(path, FileMode.Create);
                byte[] bt = Encoding.UTF8.GetBytes(sb.ToString());
                fs.Seek(0, SeekOrigin.Begin);
                fs.Write(bt, 0, bt.Length);
                fs.Flush();
                fs.Close();
                LogHelper.Info("Save checklist file success.");
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message, ex);
            }
        }

        private bool ReadCheckListFile()
        {
            try
            {
                string path = @"CheckList.txt";
                if (File.Exists(path))
                {
                    string[] linelist = File.ReadAllLines(path);
                    string datetimespan = linelist[0];
                    string workorder = linelist[1];
                    Supervisor = Convert.ToBoolean(linelist[2]);
                    IPQC = Convert.ToBoolean(linelist[3]);
                    sequece = linelist[4];
                    TimeSpan span = DateTime.Now - Convert.ToDateTime(datetimespan);

                    if (span.TotalMinutes > Convert.ToInt32(config.RESTORE_TIME))//判断是否大于10分钟，大于10分钟则不自动点检
                    {
                        return false;
                    }
                    else
                    {
                        string[] workorders = workorder.Split(';');
                        if (workorders.Length > 1)
                        {
                            if (workorders[0] == this.txbCDAMONumber.Text)//判断工单是否有变化，无变化则自动点检
                            {
                                //if (workorders[1] == initModel.currentSettings.processLayer.ToString())//判断面次是否有变化
                                //{
                                #region setup checklist
                                int seq = 1;
                                if (linelist.Count() <= 6)
                                    return false;
                                this.dgvCheckListTable.Rows.Clear();
                                for (int i = 5; i < linelist.Count(); i++)
                                {
                                    string line = linelist[i];
                                    if (string.IsNullOrEmpty(line.Trim()))
                                        continue;

                                    string[] datas = line.Split('￥');
                                    object[] objValues = new object[] { seq, datas[17], datas[9], datas[7], datas[10], datas[11], "", datas[13], datas[14], datas[6], datas[15], datas[16], datas[8], datas[0], datas[1], datas[2], datas[3], datas[4], datas[5], datas[18], datas[19], "" };
                                    this.dgvCheckListTable.Rows.Add(objValues);
                                    seq++;
                                    SetCheckListInputStatusTable();
                                    if (datas[14] == "判断值")
                                    {
                                        string[] strInputValues = new string[] { "Y", "N" };
                                        DataTable dtInput = new DataTable();
                                        dtInput.Columns.Add("name");
                                        dtInput.Columns.Add("value");
                                        DataRow rowEmpty = dtInput.NewRow();
                                        rowEmpty["name"] = "";
                                        rowEmpty["value"] = "";
                                        dtInput.Rows.Add(rowEmpty);
                                        foreach (var strValues in strInputValues)
                                        {
                                            DataRow row = dtInput.NewRow();
                                            row["name"] = strValues;
                                            row["value"] = strValues;
                                            dtInput.Rows.Add(row);
                                        }

                                        DataGridViewComboBoxCell ComboBoxCell = new DataGridViewComboBoxCell();
                                        ComboBoxCell.DataSource = dtInput;
                                        ComboBoxCell.DisplayMember = "Name";
                                        ComboBoxCell.ValueMember = "Value";
                                        dgvCheckListTable.Rows[this.dgvCheckListTable.Rows.Count - 1].Cells["tabResult2"] = ComboBoxCell;
                                    }
                                    dgvCheckListTable.Rows[this.dgvCheckListTable.Rows.Count - 1].Cells["tabResult2"].Value = datas[12];
                                    this.dgvCheckListTable.ClearSelection();

                                }
                                foreach (DataGridViewRow row in dgvCheckListTable.Rows)
                                {
                                    if (row.Cells["tabStatus"].Value.ToString() == "OK")
                                    {
                                        row.Cells["tabStatus"].Style.BackColor = Color.FromArgb(0, 192, 0);
                                    }
                                    else if ((row.Cells["tabStatus"].Value.ToString() == "NG"))
                                    {
                                        row.Cells["tabStatus"].Style.BackColor = Color.Red;
                                    }

                                }
                                return true;
                                #endregion
                                //}
                                //else
                                //{
                                //    return false;
                                //}
                            }
                            else
                            {
                                return false;
                            }
                        }
                        else
                        {
                            return false;
                        }
                    }
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message, ex);
                return false;
            }
        }

        private void WriteIntoShift2()
        {
            string datetime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            strShift = datetime;
            string path = @"CheckListShiftTemp.txt";
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(datetime + ";" + this.txbCDAMONumber.Text);
            FileStream fs = new FileStream(path, FileMode.OpenOrCreate);
            byte[] bt = Encoding.UTF8.GetBytes(sb.ToString());
            fs.Seek(0, SeekOrigin.Begin);
            fs.Write(bt, 0, bt.Length);
            fs.Flush();
            fs.Close();
        }

        //检查有没有到换班时间，如果到换班时间
        string strShiftChecklist = "";
        private bool CheckShiftChange2()
        {
            try
            {
                bool isValid = false;
                if (strShiftChecklist == "")
                    return false;

                string[] shifchangetimes = config.SHIFT_CHANGE_TIME.Split(';');
                List<string> shiftList = new List<string>();
                string nowDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                for (int i = 0; i < shifchangetimes.Length; i++)
                {

                    shiftList.Add(DateTime.Now.ToString("yyyy/MM/dd ") + shifchangetimes[i].Substring(0, 2) + ":" + shifchangetimes[i].Substring(2, 2));

                }

                shiftList.Sort();

                for (int j = shiftList.Count - 1; j < shiftList.Count; j--)
                {
                    if (j == -1)
                        break;
                    LogHelper.Debug("shift time: " + shiftList[j]);
                    string shitftime = shiftList[j];

                    if (Convert.ToDateTime(nowDate) > Convert.ToDateTime(shiftList[j])) //当前时间与设定的时间做比较，如果到换班时间则比较上次点检的时间
                    {
                        if (Convert.ToDateTime(strShiftChecklist) > Convert.ToDateTime(shitftime))
                        {
                            isValid = true;
                        }
                        break;
                    }
                    else
                    {
                        if (Convert.ToDateTime(strShiftChecklist).ToString("yyyy/MM/dd") != Convert.ToDateTime(nowDate).ToString("yyyy/MM/dd"))//add by qy
                        {
                            string covert_datetime = nowDate;
                            if (j == shiftList.Count - 1)
                            {
                                covert_datetime = shiftList[j - 1];
                            }
                            else if (j == 0)
                            {
                                covert_datetime = shiftList[j];
                            }
                            if (Convert.ToDateTime(strShiftChecklist) < Convert.ToDateTime(nowDate) && Convert.ToDateTime(nowDate) < Convert.ToDateTime(covert_datetime))
                            {
                                isValid = true;
                            }
                            break;
                        }

                        //if (Convert.ToDateTime(strShiftChecklist).ToString("yyyy/MM/dd") != Convert.ToDateTime(nowDate).ToString("yyyy/MM/dd"))//add by qy
                        //{
                        //    shitftime = Convert.ToDateTime(shitftime).AddDays(-1).ToString("yyyy/MM/dd HH:mm:ss");

                        //    if (Convert.ToDateTime(strShiftChecklist) > Convert.ToDateTime(shitftime))
                        //    {
                        //        isValid = true;
                        //    }
                        //    break;
                        //}
                    }
                }

                return isValid;
            }
            catch (Exception ex)
            {
                LogHelper.Debug(ex.Message, ex);
                return false;
            }

        }

        private void InitShift2(string wo)
        {
            string path = @"CheckListShiftTemp.txt";
            if (File.Exists(path))
            {
                string[] content = File.ReadAllLines(path);

                foreach (var item in content)
                {
                    if (item != "")
                    {
                        string[] items = item.Split(';');
                        //if (items[1] == wo)
                        //{
                        strShiftChecklist = items[0];
                        break;
                        //}
                    }
                }
            }
        }
        DateTime next_checklist_time = DateTime.Now;
        string checklist_freq_time = "";
        private void InitWorkOrderType()
        {

            Dictionary<string, string> dicfreq = new Dictionary<string, string>();
            string CHECKLIST_FREQ = config.CHECKLIST_FREQ;
            string[] freqs = CHECKLIST_FREQ.Split(';');
            foreach (var item in freqs)
            {
                string[] items = item.Split(',');
                string key = items[0];
                if (key == "")
                    key = "OTHERS";
                dicfreq[key] = items[1];
            }

            GetAttributeValue getAttriHandler = new GetAttributeValue(sessionContext, initModel, this);
            string[] valuesAttri = getAttriHandler.GetAttributeValueForAll(1, this.txbCDAMONumber.Text, "-1", "WORKORDER_TYPE");
            if (valuesAttri != null && valuesAttri.Length > 0)
            {
                string value = valuesAttri[1];
                if (CHECKLIST_FREQ.Contains(value))
                {
                    checklist_freq_time = dicfreq[value];
                }
                else
                {
                    checklist_freq_time = dicfreq["OTHERS"];
                }
            }
            else
            {
                checklist_freq_time = dicfreq["OTHERS"];
            }
            if (strShiftChecklist != "")
            {
                next_checklist_time = Convert.ToDateTime(strShiftChecklist).AddMinutes(double.Parse(checklist_freq_time) * 60);
            }
            else
            {
                next_checklist_time = DateTime.Now.AddMinutes(double.Parse(checklist_freq_time) * 60);
            }

        }

        private void InitIPIWorkOrderType()
        {
            workorderType = "NULL";
            GetAttributeValue getAttriHandler = new GetAttributeValue(sessionContext, initModel, this);
            string[] valuesAttri = getAttriHandler.GetAttributeValueForAll(1, this.txbCDAMONumber.Text, "-1", "WORKORDER_TYPE");
            if (valuesAttri != null && valuesAttri.Length > 0)
            {
                workorderType = valuesAttri[1];
            }
        }

        private void InitProductionChecklist()
        {
            if (DateTime.Now > next_checklist_time)
            {
                if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
                {
                    InitTaskData_SOCKET("过程点检");
                    isStartLineCheck = false;
                    next_checklist_time = DateTime.Now.AddMinutes(double.Parse(checklist_freq_time) * 60);
                }
            }
        }
        #endregion

        #region IO TOWER LIGHT
        private bool SendSN(string serialNumber)
        {
            try
            {
                //Thread.Sleep(20);
                if (config.DataOutputInterface == "COM")
                {
                    try
                    {
                        if (config.IOSerialPort != "" && config.IOSerialPort != null)
                        {
                            initModel.scannerHandler.handler2().Write(strToToHexByte(serialNumber), 0, strToToHexByte(serialNumber).Length);
                            LogHelper.Info("Send command:" + serialNumber);
                        }
                        return true;
                    }
                    catch (Exception e)
                    {
                        LogHelper.Error(e);
                        return false;
                    }
                }
                else
                {
                    if (config.OutputEnter == "1")
                    {
                        if (Control.IsKeyLocked(Keys.CapsLock))
                        {
                            SendKeys.SendWait("{CAPSLOCK}" + serialNumber + "\r"); //大写键总是被按起。。。。
                        }
                        else
                        {
                            SendKeys.SendWait(serialNumber + "\r");
                        }
                    }
                    else
                    {
                        if (Control.IsKeyLocked(Keys.CapsLock))
                        {
                            SendKeys.SendWait("{CAPSLOCK}" + serialNumber);
                        }
                        else
                        {
                            SendKeys.SendWait(serialNumber);
                        }
                    }
                    SendKeys.Flush();
                    Thread.Sleep(300);
                    return true;
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex);
                return false;
            }
        }
        private static byte[] strToToHexByte(string hexString)
        {
            hexString = hexString.Replace(" ", "");
            if ((hexString.Length % 2) != 0)
                hexString += " ";
            byte[] returnBytes = new byte[hexString.Length / 2];
            for (int i = 0; i < returnBytes.Length; i++)
                returnBytes[i] = Convert.ToByte(hexString.Substring(i * 2, 2), 16);
            return returnBytes;
        }
        #endregion

        #region Panel Position and Direction graphics
        PosGraghicsForm frm = null;
        private void btnShowPCB_Click(object sender, EventArgs e)
        {
            if (frm != null && frm.pictureBox1.Image != null)
            {
                frm.Hide();
            }
            if (txbCDAPartNumber.Text == "")
            {
                errorHandler(2, Message("msg_No activated work order"), "");
                return;
            }
            frm = new PosGraghicsForm(this, sessionContext, initModel, txbCDAPartNumber.Text, initModel.currentSettings == null ? "-1" : initModel.currentSettings.processLayer.ToString());
            frm.Show();
            if (frm.pictureBox1.Image == null)
                frm.Hide();
        }
        #endregion


        public bool isFormOutPoump = false;

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            this.txbCDADataInput.Text = "";
            LogoutForm frmOut = new LogoutForm(UserName, this, initModel,sessionContext);
            DialogResult dr = frmOut.ShowDialog();

            if (dr == DialogResult.OK)
            {
                UserName = frmOut.UserName;
                lblUser.Text = UserName;
                lblLoginTime.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                sessionContext = frmOut.sessionContext;
                if (config.LogInType == "COM")
                {
                    SerialPort serialPort = new SerialPort();
                    serialPort.PortName = config.SerialPort;
                    serialPort.BaudRate = int.Parse(config.BaudRate);
                    serialPort.Parity = (Parity)int.Parse(config.Parity);
                    serialPort.StopBits = (StopBits)1;
                    serialPort.Handshake = Handshake.None;
                    serialPort.DataBits = int.Parse(config.DataBits);
                    serialPort.NewLine = "\r";
                    serialPort.DataReceived += new SerialDataReceivedEventHandler(DataRecivedHeandler);
                    serialPort.Open();
                    initModel.scannerHandler.SetSerialPortData(serialPort);
                }
            }
        }

    }
}
