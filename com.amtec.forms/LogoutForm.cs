using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using com.amtec.configurations;
using com.itac.mes.imsapi.domain.container;
using com.itac.mes.imsapi.client.dotnet;
using com.itac.oem.common.container.imsapi.utils;
using com.amtec.forms;
using com.amtec.model;
using com.amtec.action;
using System.IO.Ports;
using System.Threading;
using System.Text.RegularExpressions;

namespace com.amtec.forms
{
    public partial class LogoutForm : Form
    {
        private IMSApiSessionValidationStruct sessionValidationStruct;
        public IMSApiSessionContextStruct sessionContext = null;
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private LanguageResources res;
        public string UserName = "";
        public int LoginResult = 0;
        public bool isCanLogin = false;
        MainView view;
        private InitModel init;
        public LogoutForm(string userName, MainView _view, InitModel _init, IMSApiSessionContextStruct _sessionContext)
        {
            InitializeComponent();
            sessionContext = _sessionContext;
            this.txtUserName.Text = userName;
            view = _view;
            init = _init;
            SetControlStatus(false);
            this.progressBar1.Value = 0;
            this.progressBar1.Maximum = 100;
            this.progressBar1.Step = 1;

            this.timer1.Interval = 100;
            this.timer1.Tick += new EventHandler(timer_Tick);

            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.DoWork += new DoWorkEventHandler(worker_DoWork);
            backgroundWorker1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(worker_RunWorkerCompleted);
        }

        #region add by liuxue for scan user & psw when login
        private SerialPort serialPort;
        private ApplicationConfiguration config;
        private void LoginForm_Load(object sender, EventArgs e)
        {
            config = new ApplicationConfiguration();
            if (config.LogInType == "COM")
            {
                this.txtPassword.ReadOnly = true;
                this.txtUserName.ReadOnly = true;
                
            }
            else
            {
                this.txtPassword.ReadOnly = false;
                this.txtUserName.ReadOnly = false;
            }
        }
        private void InitSerialPort()
        {
            serialPort = new SerialPort();
            serialPort.PortName = config.SerialPort;
            serialPort.BaudRate = int.Parse(config.BaudRate);
            serialPort.Parity = (Parity)int.Parse(config.Parity);
            serialPort.StopBits = (StopBits)1;
            serialPort.Handshake = Handshake.None;
            serialPort.DataBits = int.Parse(config.DataBits);
            serialPort.NewLine = "\r";

            serialPort.DataReceived += new SerialDataReceivedEventHandler(DataRecivedHeandler);
            serialPort.Open();
        }
        public void DataRecivedHeandler(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = (SerialPort)sender;
            try
            {
                Thread.Sleep(200);
                Byte[] bt = new Byte[sp.BytesToRead];
                sp.Read(bt, 0, sp.BytesToRead);
                string indata = System.Text.Encoding.ASCII.GetString(bt).Trim();
                Match match = Regex.Match(indata, config.LoadExtractPattern);
                if (match.Success)
                {
                    SetUserControlText(match.Groups[1].ToString());
                    SetPasswordControlText(match.Groups[2].ToString());
                    this.Invoke(new MethodInvoker(delegate
                    {
                        btnOK_Click(null, null);
                    }));
                }
                else
                {
                    SetStatusLabelText("条码错误", 1);
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex);
            }
        }
        public void SetUserControlText(string strText)
        {
            this.InvokeEx(x => this.txtUserName.Text = strText);
        }

        public void SetPasswordControlText(string strText)
        {
            this.InvokeEx(x => this.txtPassword.Text = strText);
        }
        #endregion

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (!VerifyLoginInfo())
                return;
            LogHelper.Info("Login start...");
            backgroundWorker1.RunWorkerAsync();
            this.lblErrorMsg.Text = "Loading application....";
            this.timer1.Start();
            SetControlStatus(false);
            view.isFormOutPoump = false;
        }
        private void btnLogout_Click(object sender, EventArgs e)
        {
            int errorCode = imsapi.regLogout(sessionContext);
            LogHelper.Info("Api regLogout result code =" + errorCode);
            this.txtUserName.Text = "";
            this.txtPassword.Text = "";
            SetControlStatus(true);
            view.isFormOutPoump = true;
            if (config.LogInType == "COM" && init.scannerHandler.handler().IsOpen)
            {
                init.scannerHandler.handler().Close();
                InitSerialPort();
            }
        }

        private void timer_Tick(object sender, EventArgs e)
        {
            if (this.progressBar1.Value < this.progressBar1.Maximum - 5)
            {
                this.progressBar1.Value++;
            }
        }

        private void worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            timer1.Stop();
            this.progressBar1.Value = this.progressBar1.Maximum;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            view.isFormOutPoump = false;
            this.Close();
        }

        private void txtPassword_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnOK_Click(null, null);
            }
        }

        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            //添加你初始化的代码
            res = new LanguageResources();
            sessionValidationStruct = new IMSApiSessionValidationStruct();
            sessionValidationStruct.stationNumber = config.StationNumber;
            sessionValidationStruct.stationPassword = "";
            sessionValidationStruct.user = this.txtUserName.Text.Trim();
            sessionValidationStruct.password = this.txtPassword.Text.Trim();
            sessionValidationStruct.client = config.Client;
            sessionValidationStruct.registrationType = config.RegistrationType;
            sessionValidationStruct.systemIdentifier = config.StationNumber;
            UserName = this.txtUserName.Text.Trim();

            LoginResult = imsapi.regLogin(sessionValidationStruct, out sessionContext);
            LogHelper.Info("Api regLogin (result code=" + LoginResult + ")");
            LogHelper.Info("Login end...");
            if (LoginResult != IMSApiDotNetConstants.RES_OK)
            {
                this.Invoke(new MethodInvoker(delegate
                {
                    SetStatusLabelText("Api regLogin error(error code=" + LoginResult + ")", 1);
                    SetControlStatus(true);
                }));
                return;
            }
            else
            {
                if (!VerifyTeamNumber())
                {
                    this.Invoke(new MethodInvoker(delegate
                    {
                        SetStatusLabelText("User not belong to the userteam.[" + config.UserTeam + "]", 1);
                        SetControlStatus(true);
                    }));
                    return;
                }
                this.Invoke(new MethodInvoker(delegate
                {
                    if (config.LogInType == "COM" && serialPort.IsOpen)
                        serialPort.Close();
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }));
            }
        }

        public delegate void SetStatusLabelTextDel(string strText, int iCase);
        public void SetStatusLabelText(string strText, int iCase)
        {
            if (this.lblErrorMsg.InvokeRequired)
            {
                SetStatusLabelTextDel setText = new SetStatusLabelTextDel(SetStatusLabelText);
                Invoke(setText, new object[] { strText, iCase });
            }
            else
            {
                this.lblErrorMsg.Text = strText;
                if (iCase == 0)
                {
                    this.lblErrorMsg.ForeColor = Color.Black;
                }
                else if (iCase == 1)
                {
                    this.lblErrorMsg.ForeColor = Color.Red;
                }
            }
        }

        private void SetControlStatus(bool isOK)
        {
            this.btnOK.Enabled = isOK;
            this.btnLogout.Enabled = !isOK;
            this.txtPassword.Enabled = isOK;
            this.txtUserName.Enabled = isOK;
        }

        private bool VerifyLoginInfo()
        {
            bool isValidate = true;
            if (string.IsNullOrEmpty(this.txtUserName.Text.Trim()) || string.IsNullOrEmpty(this.txtPassword.Text.Trim()))
            {
                SetStatusLabelText("Pls input user name/password.", 1);
                isValidate = false;
            }
            return isValidate;
        }

        private bool VerifyTeamNumber()
        {
            bool isValid = false;

            GetUserData getUser = new GetUserData(sessionContext);
            string[] mdataGetUserDataValues = getUser.mdataGetUserData(this.txtUserName.Text.Trim(), this.txtPassword.Text.Trim(), config.StationNumber);
            if (mdataGetUserDataValues != null && mdataGetUserDataValues.Length > 0)
            {
                string teamNo = mdataGetUserDataValues[2];
                if (string.IsNullOrEmpty(config.UserTeam))
                {
                    isValid = true;
                }
                else if (string.IsNullOrEmpty(teamNo))
                {
                    isValid = false;
                }
                else
                {
                    if (config.UserTeam.Contains(teamNo))
                    {
                        isValid = true;
                    }
                }
            }

            return isValid;
        }
    }
}
