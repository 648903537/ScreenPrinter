using com.amtec.action;
using com.amtec.forms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ScreenPrinter.com.amtec.forms
{
    public partial class IPIForm : Form
    {
        string IPIStatus = "";
        MainView view;
        public IPIForm(string _IPIStatus, MainView _view)
        {
            InitializeComponent();
            IPIStatus = _IPIStatus;
            view = _view;
        }
        private void IPIForm_Load(object sender, EventArgs e)
        {
            view.InitCintrolLanguage(this);

            if (IPIStatus == "0")
            {
                this.Text = view.Message("msg_IPI form title");
                lblIPITitle.Text = view.Message("msg_Initial Product Inspection in process");
                lbltime.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                LogHelper.Info(view.Message("msg_Initial Product Inspection in process"));
                this.panel1.BackColor = Color.LemonChiffon;
                this.btnAllowProduction.Visible = false;
            }
            else if (IPIStatus == "1")
            {
                this.Text = view.Message("msg_IPI form title");
                lblIPITitle.Text = view.Message("msg_Initial Product Inspection") + view.Message("msg_SUCCESS");
                lbltime.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                LogHelper.Info(view.Message("msg_Initial Product Inspection") + view.Message("msg_SUCCESS"));
                this.panel1.BackColor = Color.LightGreen;
                this.btnOk.Visible = true;
                this.btnAllowProduction.Visible = false;
            }
            else if (IPIStatus == "-1")
            {
                this.Text = view.Message("msg_IPI form title");
                lblIPITitle.Text = view.Message("msg_Initial Product Inspection") + view.Message("msg_FAIL");
                lbltime.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                LogHelper.Error(view.Message("msg_Initial Product Inspection") + view.Message("msg_FAIL"));
                this.panel1.BackColor = Color.Red;
                this.btnAnother.Visible = true;
                this.btnOk.Visible = true;
                this.btnOk.Enabled = false;
                this.btnAllowProduction.Visible = false;
            }
            else if (IPIStatus == "-3")
            {
                this.Text = view.Message("msg_ProductionInspection form title");
                lblIPITitle.Text = view.Message("msg_Product Inspection") + view.Message("msg_FAIL");
                lbltime.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                LogHelper.Error(view.Message("msg_Product Inspection") + view.Message("msg_FAIL"));
                this.panel1.BackColor = Color.Red;
                this.btnAnother.Visible = false;
                this.btnOk.Visible = false;
                this.btnAllowProduction.Visible = true;
            }
            if (IPIStatus == "4")
            {
                this.Text = view.Message("msg_LastProductionInspection form title");
                lblIPITitle.Text = view.Message("msg_Last Product Inspection in process");
                lbltime.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                LogHelper.Info(view.Message("msg_Last Product Inspection in process"));
                this.panel1.BackColor = Color.LemonChiffon;
                this.btnAllowProduction.Visible = false;
            }
            else if (IPIStatus == "-5")
            {
                this.Text = view.Message("msg_LastProductionInspection form title");
                lblIPITitle.Text = view.Message("msg_Last Product Inspection") + view.Message("msg_FAIL");
                lbltime.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                LogHelper.Error(view.Message("msg_Last Product Inspection") + view.Message("msg_FAIL"));
                this.panel1.BackColor = Color.Red;
                this.btnAnother.Visible = false;
                this.btnOk.Visible = false;
                this.btnAllowProduction.Visible = true;
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (view.IPIStatus == 1)
                view.errorHandler(0, view.Message("msg_IPI SUCCESS"), "");
            if (view.IPIStatus == 5)
                view.errorHandler(0, view.Message("msg_IPI Last SUCCESS"), "");
            view.IPITimerStop();
            this.Hide();
        }

        private void btnAnother_Click(object sender, EventArgs e)
        {
            view.RemoveIPI();
            view.errorHandler(2, view.Message("msg_IPI FAIL"), "");
            view.IPITimerStop();
            this.Hide();
        }

        private void IPIForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (view.config.LogInType == "COM" && view.initModel.scannerHandler.handler().IsOpen)
            {
                view.initModel.scannerHandler.handler().Close();
                view.OpenScanPort();
            }
            view.IPITimerStop();
            //this.Invoke(new MethodInvoker(delegate
            //{
            //    if (!view.cSocket.tcpsend.Connected)
            //    {
            //        view.cSocket = new SocketClientHandler(view);
            //        view.cSocket.connect(view.config.IPAddress, view.config.Port);
            //    }
            //    view.GetTimerStart();
            //}));
            this.Hide();
        }

        private void btnAllowProduction_Click(object sender, EventArgs e)
        {
            if (view.config.LogInType == "COM" && view.initModel.scannerHandler.handler().IsOpen)
                view.initModel.scannerHandler.handler().Close();
            LoginForm LogForm = new LoginForm(3, view, "");
            LogForm.ShowDialog(view);
            view.VerifyIPIStatus();
            if (view.IPIStatus == 3 || view.IPIStatus == 5)
            {
                view.IPITimerStop();
                this.Hide();
            }
        }
    }
}
