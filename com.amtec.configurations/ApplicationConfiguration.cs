using com.amtec.action;
using com.amtec.forms;
using com.amtec.model;
using com.itac.mes.imsapi.domain.container;
using System;
using System.Collections.Generic;
using System.Xml.Linq;
using System.Xml;
using System.Linq;

namespace com.amtec.configurations
{
    public class ApplicationConfiguration
    {
        public String StationNumber { get; set; }

        public String Client { get; set; }

        public String RegistrationType { get; set; }

        public String SerialPort { get; set; }

        public String BaudRate { get; set; }

        public String Parity { get; set; }

        public String StopBits { get; set; }

        public String DataBits { get; set; }

        public String NewLineSymbol { get; set; }

        public String High { get; set; }

        public String Low { get; set; }

        public String EndCommand { get; set; }

        public String DLExtractPattern { get; set; }

        public String MBNExtractPattern { get; set; }

        public String MDAPath { get; set; }

        public String EquipmentExtractPattern { get; set; }

        public String OpacityValue { get; set; }

        public String LocationXY { get; set; }

        public String ThawingDuration { get; set; }

        public String LockTime { get; set; }

        public String UsageTime { get; set; }

        public String ThawingCheck { get; set; }

        public String GateKeeperTimer { get; set; }

        public String SolderPasteValidity { get; set; }

        public String IPAddress { get; set; }

        public String Port { get; set; }

        public String OpenControlBox { get; set; }

        public String StencilPrefix { get; set; }

        public String TimerSpan { get; set; }

        public String StartTrigerStr { get; set; }

        public String EndTrigerStr { get; set; }

        public String NoRead { get; set; }

        public String LogFileFolder { get; set; }

        public String LogTransOK { get; set; }

        public String LogTransError { get; set; }

        public String ChangeFileName { get; set; }

        public String CheckListFolder { get; set; }

        public String LogInType { get; set; }

        public String LoadExtractPattern { get; set; }

        public String Language { get; set; }

        public String ReduceEquType { get; set; }

        public String UserTeam { get; set; }

        public String FileNamePattern { get; set; }

        public String FilterByFileName { get; set; }

        public String IPI_STATUS_CHECK { get; set; }

        public String IPI_STATUS_CHECK_INTERVAL { get; set; }

        public String WarningQty { get; set; }

        public String Authorized_Seria_Number_Transfer { get; set; }

        public String Auto_Work_Order_Change { get; set; }

        public String IsNeedTransWO { get; set; }

        public String SHIFT_CHANGE_TIME { get; set; }

        public String Authorized_Allow_Production { get; set; }

        public String IsNeedProductionInspection { get; set; }

        public String RESTORE_TREAD_TIMER { get; set; }

        public String RESTORE_TIME { get; set; }

        public String AutoNextMaterial { get; set; }

        public String IPI_WORKORDERTYPE_CHECK { get; set; }

        #region checklist
        public String CHECKLIST_IPAddress { get; set; }
        public String CHECKLIST_Port { get; set; }
        public String CHECKLIST_SOURCE { get; set; }
        public String AUTH_CHECKLIST_APP_TEAM { get; set; }
        public String CHECKLIST_FREQ { get; set; }
        #endregion

        public String LIGHT_CHANNEL_ON { get; set; }
        public String LIGHT_CHANNEL_OFF { get; set; }
        public String IO_BOX_CONNECT { get; set; }
        public String IOSerialPort { get; set; }
        public String IOBaudRate { get; set; }
        public String IOParity { get; set; }
        public String IOStopBits { get; set; }
        public String IODataBits { get; set; }
        public String OutputEnter { get; set; }
        public String DataOutputInterface { get; set; }

        #region postion display
        public String LAYER_DISPLAY { get; set; }
        #endregion
        Dictionary<string, string> dicConfig = null;

        public ApplicationConfiguration()
        {
            try
            {
                CommonModel commonModel = ReadIhasFileData.getInstance();
                XDocument config = XDocument.Load("ApplicationConfig.xml");
                StationNumber = commonModel.Station;
                Client = commonModel.Client;
                RegistrationType = commonModel.RegisterType;
                SerialPort = GetDescendants("SerialPort", config);//config.Descendants("SerialPort").First().Value;
                BaudRate = GetDescendants("BaudRate", config);//config.Descendants("BaudRate").First().Value;
                Parity = GetDescendants("Parity", config);//config.Descendants("Parity").First().Value;
                StopBits = GetDescendants("StopBits", config);//config.Descendants("StopBits").First().Value;
                DataBits = GetDescendants("DataBits", config);//config.Descendants("DataBits").First().Value;
                NewLineSymbol = GetDescendants("NewLineSymbol", config);//config.Descendants("NewLineSymbol").First().Value;
                High = GetDescendants("High", config);//config.Descendants("High").First().Value;
                Low = GetDescendants("Low", config);// config.Descendants("Low").First().Value;
                EndCommand = GetDescendants("EndCommand", config);//config.Descendants("EndCommand").First().Value;
                DLExtractPattern = GetDescendants("DLExtractPattern", config);//config.Descendants("DLExtractPattern").First().Value;
                MBNExtractPattern = GetDescendants("MBNExtractPattern", config);//config.Descendants("MBNExtractPattern").First().Value;
                EquipmentExtractPattern = GetDescendants("EquipmentExtractPattern", config);//config.Descendants("EquipmentExtractPattern").First().Value;
                OpacityValue = GetDescendants("OpacityValue", config);//config.Descendants("OpacityValue").First().Value;
                LocationXY = GetDescendants("LocationXY", config);//config.Descendants("LocationXY").First().Value;
                ThawingDuration = GetDescendants("ThawingDuration", config);//config.Descendants("ThawingDuration").First().Value;
                ThawingCheck = GetDescendants("ThawingCheck", config);//config.Descendants("ThawingCheck").First().Value;
                LockTime = GetDescendants("LockOutTime", config);// config.Descendants("LockOutTime").First().Value;
                UsageTime = GetDescendants("UsageDurationSetting", config);//config.Descendants("UsageDurationSetting").First().Value;
                GateKeeperTimer = GetDescendants("GateKeeperTimer", config);//config.Descendants("GateKeeperTimer").First().Value;
                SolderPasteValidity = GetDescendants("SolderPasteValidity", config);// config.Descendants("SolderPasteValidity").First().Value;
                OpenControlBox = GetDescendants("OpenControlBox", config);// config.Descendants("OpenControlBox").First().Value;
                StencilPrefix = GetDescendants("StencilPrefix", config);//config.Descendants("StencilPrefix").First().Value;
                TimerSpan = GetDescendants("TimerSpan", config);//config.Descendants("TimerSpan").First().Value;
                StartTrigerStr = GetDescendants("StartTrigerStr", config);//config.Descendants("StartTrigerStr").First().Value;
                EndTrigerStr = GetDescendants("EndTrigerStr", config);//config.Descendants("EndTrigerStr").First().Value;
                NoRead = GetDescendants("NoRead", config);//config.Descendants("NoRead").First().Value;
                LogFileFolder = GetDescendants("LogFileFolder", config);//config.Descendants("LogFileFolder").First().Value;
                LogTransOK = GetDescendants("LogTransOK", config);// config.Descendants("LogTransOK").First().Value;
                LogTransError = GetDescendants("LogTransError", config);// config.Descendants("LogTransError").First().Value;
                ChangeFileName = GetDescendants("ChangeFileName", config);//config.Descendants("ChangeFileName").First().Value;
                CheckListFolder = GetDescendants("CheckListFolder", config);//config.Descendants("CheckListFolder").First().Value;
                LoadExtractPattern = GetDescendants("LoadExtractPattern", config);// config.Descendants("LoadExtractPattern").First().Value;
                LogInType = GetDescendants("LogInType", config);//config.Descendants("LogInType").First().Value;
                Language = GetDescendants("Language", config);//config.Descendants("Language").First().Value;
                MDAPath = GetDescendants("MDAPath", config);//config.Descendants("MDAPath").First().Value;
                IPAddress = GetDescendants("IPAddress", config);//config.Descendants("IPAddress").First().Value;
                Port = GetDescendants("Port", config);//config.Descendants("Port").First().Value;
                ReduceEquType = GetDescendants("ReduceEquType", config);//config.Descendants("ReduceEquType").First().Value;
                UserTeam = GetDescendants("AUTH_TEAM", config);//config.Descendants("UserTeam").First().Value;

                FilterByFileName = GetDescendants("FilterByFileName", config);//config.Descendants("FilterByFileName").First().Value;
                FileNamePattern = GetDescendants("FileNamePattern", config);//config.Descendants("FileNamePattern").First().Value;

                IPI_STATUS_CHECK = GetDescendants("IPI_STATUS_CHECK", config);
                IPI_STATUS_CHECK_INTERVAL = GetDescendants("IPI_STATUS_CHECK_INTERVAL", config);
                WarningQty = GetDescendants("WarningQty", config);
                Authorized_Seria_Number_Transfer = GetDescendants("Authorized_Seria_Number_Transfer", config);
                Auto_Work_Order_Change = GetDescendants("Auto_Work_Order_Change", config);
                IsNeedTransWO = GetDescendants("IsNeedTransWO", config);
                SHIFT_CHANGE_TIME = GetDescendants("SHIFT_CHANGE_TIME", config);
                Authorized_Allow_Production = GetDescendants("Authorized_Allow_Production", config);
                IsNeedProductionInspection = GetDescendants("Production_Inspection_CHECK", config);

                RESTORE_TREAD_TIMER = GetDescendants("RESTORE_TREAD_TIMER", config);
                RESTORE_TIME = GetDescendants("RESTORE_TIME", config);

                AutoNextMaterial = GetDescendants("MATERIAL_SPLICING", config);

                CHECKLIST_IPAddress = GetDescendants("CHECKLIST_IPAddress", config);
                CHECKLIST_Port = GetDescendants("CHECKLIST_Port", config);
                CHECKLIST_SOURCE = GetDescendants("CHECKLIST_SOURCE", config);
                AUTH_CHECKLIST_APP_TEAM = GetDescendants("AUTH_CHECKLIST_APP_TEAM", config);
                CHECKLIST_FREQ = GetDescendants("CHECKLIST_FREQ", config);

                LIGHT_CHANNEL_ON = GetDescendants("LIGHT_CHANNEL_ON", config);
                LIGHT_CHANNEL_OFF = GetDescendants("LIGHT_CHANNEL_OFF", config);
                IO_BOX_CONNECT = GetDescendants("IO_BOX_CONNECT", config);
                if (IO_BOX_CONNECT != null && IO_BOX_CONNECT.Split(';').Length >= 6)
                {
                    string[] infos = IO_BOX_CONNECT.Split(';');
                    IOSerialPort = "COM" + infos[0];
                    IOBaudRate = infos[1];
                    IOStopBits = infos[4];
                    IODataBits = infos[2];
                    IOParity = infos[3];
                }
                OutputEnter = GetDescendants("OutputEnter", config);
                DataOutputInterface = GetDescendants("DataOutputInterface", config);
                LAYER_DISPLAY = GetDescendants("DataOutputInterface", config);

                IPI_WORKORDERTYPE_CHECK = GetDescendants("IPI_WORKORDERTYPE_CHECK", config);
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex);
            }
        }

        public ApplicationConfiguration(IMSApiSessionContextStruct sessionContext, MainView view)
        {
            try
            {
                dicConfig = new Dictionary<string, string>();
                ConfigManage configHandler = new ConfigManage(sessionContext, view);
                CommonModel commonModel = ReadIhasFileData.getInstance();
                if (commonModel.UpdateConfig == "L")
                {
                    XDocument config = XDocument.Load("ApplicationConfig.xml");
                    StationNumber = commonModel.Station;
                    Client = commonModel.Client;
                    RegistrationType = commonModel.RegisterType;
                    SerialPort = GetDescendants("SerialPort", config);//config.Descendants("SerialPort").First().Value;
                    BaudRate = GetDescendants("BaudRate", config);//config.Descendants("BaudRate").First().Value;
                    Parity = GetDescendants("Parity", config);//config.Descendants("Parity").First().Value;
                    StopBits = GetDescendants("StopBits", config);//config.Descendants("StopBits").First().Value;
                    DataBits = GetDescendants("DataBits", config);//config.Descendants("DataBits").First().Value;
                    NewLineSymbol = GetDescendants("NewLineSymbol", config);//config.Descendants("NewLineSymbol").First().Value;
                    High = GetDescendants("High", config);//config.Descendants("High").First().Value;
                    Low = GetDescendants("Low", config);//config.Descendants("Low").First().Value;
                    EndCommand = GetDescendants("EndCommand", config);//config.Descendants("EndCommand").First().Value;
                    DLExtractPattern = GetDescendants("DLExtractPattern", config);//config.Descendants("DLExtractPattern").First().Value;
                    MBNExtractPattern = GetDescendants("MBNExtractPattern", config);// config.Descendants("MBNExtractPattern").First().Value;
                    EquipmentExtractPattern = GetDescendants("EquipmentExtractPattern", config);//config.Descendants("EquipmentExtractPattern").First().Value;
                    OpacityValue = GetDescendants("OpacityValue", config);// config.Descendants("OpacityValue").First().Value;
                    LocationXY = GetDescendants("LocationXY", config);//config.Descendants("LocationXY").First().Value;
                    ThawingDuration = GetDescendants("ThawingDuration", config);//config.Descendants("ThawingDuration").First().Value;
                    ThawingCheck = GetDescendants("ThawingCheck", config);//config.Descendants("ThawingCheck").First().Value;
                    LockTime = GetDescendants("LockOutTime", config);// config.Descendants("LockOutTime").First().Value;
                    UsageTime = GetDescendants("UsageDurationSetting", config);//config.Descendants("UsageDurationSetting").First().Value;
                    GateKeeperTimer = GetDescendants("GateKeeperTimer", config);//config.Descendants("GateKeeperTimer").First().Value;
                    SolderPasteValidity = GetDescendants("SolderPasteValidity", config);//config.Descendants("SolderPasteValidity").First().Value;
                    OpenControlBox = GetDescendants("OpenControlBox", config);//config.Descendants("OpenControlBox").First().Value;
                    StencilPrefix = GetDescendants("StencilPrefix", config);//config.Descendants("StencilPrefix").First().Value;
                    TimerSpan = GetDescendants("TimerSpan", config);//config.Descendants("TimerSpan").First().Value;
                    StartTrigerStr = GetDescendants("StartTrigerStr", config);//config.Descendants("StartTrigerStr").First().Value;
                    EndTrigerStr = GetDescendants("EndTrigerStr", config);// config.Descendants("EndTrigerStr").First().Value;
                    NoRead = GetDescendants("NoRead", config);//config.Descendants("NoRead").First().Value;
                    LogFileFolder = GetDescendants("LogFileFolder", config);//config.Descendants("LogFileFolder").First().Value;
                    LogTransOK = GetDescendants("LogTransOK", config);// config.Descendants("LogTransOK").First().Value;
                    LogTransError = GetDescendants("LogTransError", config);// config.Descendants("LogTransError").First().Value;
                    ChangeFileName = GetDescendants("ChangeFileName", config);// config.Descendants("ChangeFileName").First().Value;
                    CheckListFolder = GetDescendants("CheckListFolder", config);// config.Descendants("CheckListFolder").First().Value;
                    MDAPath = GetDescendants("MDAPath", config);// config.Descendants("MDAPath").First().Value;
                    IPAddress = GetDescendants("IPAddress", config);//config.Descendants("IPAddress").First().Value;
                    Port = GetDescendants("Port", config);// config.Descendants("Port").First().Value;
                    ReduceEquType = GetDescendants("ReduceEquType", config);// config.Descendants("ReduceEquType").First().Value;
                    UserTeam = GetDescendants("AUTH_TEAM", config);//config.Descendants("UserTeam").First().Value;
                    FilterByFileName = GetDescendants("FilterByFileName", config);//config.Descendants("FilterByFileName").First().Value;
                    FileNamePattern = GetDescendants("FileNamePattern", config);//config.Descendants("FileNamePattern").First().Value;
                    IPI_STATUS_CHECK = GetDescendants("IPI_STATUS_CHECK", config);
                    IPI_STATUS_CHECK_INTERVAL = GetDescendants("IPI_STATUS_CHECK_INTERVAL", config);
                    WarningQty = GetDescendants("WarningQty", config);
                    Authorized_Seria_Number_Transfer = GetDescendants("Authorized_Seria_Number_Transfer", config);
                    Auto_Work_Order_Change = GetDescendants("Auto_Work_Order_Change", config);
                    IsNeedTransWO = GetDescendants("IsNeedTransWO", config);
                    SHIFT_CHANGE_TIME = GetDescendants("SHIFT_CHANGE_TIME", config);
                    Authorized_Allow_Production = GetDescendants("Authorized_Allow_Production", config);
                    IsNeedProductionInspection = GetDescendants("Production_Inspection_CHECK", config);
                    LogInType = GetDescendants("LogInType", config);//config.Descendants("LogInType").First().Value;
                    RESTORE_TREAD_TIMER = GetDescendants("RESTORE_TREAD_TIMER", config);
                    RESTORE_TIME = GetDescendants("RESTORE_TIME", config);
                    AutoNextMaterial = GetDescendants("MATERIAL_SPLICING", config);

                    CHECKLIST_IPAddress = GetDescendants("CHECKLIST_IPAddress", config);
                    CHECKLIST_Port = GetDescendants("CHECKLIST_Port", config);
                    CHECKLIST_SOURCE = GetDescendants("CHECKLIST_SOURCE", config);
                    AUTH_CHECKLIST_APP_TEAM = GetDescendants("AUTH_CHECKLIST_APP_TEAM", config);
                    CHECKLIST_FREQ = GetDescendants("CHECKLIST_FREQ", config);

                    LIGHT_CHANNEL_ON = GetDescendants("LIGHT_CHANNEL_ON", config);
                    LIGHT_CHANNEL_OFF = GetDescendants("LIGHT_CHANNEL_OFF", config);
                    IO_BOX_CONNECT = GetDescendants("IO_BOX_CONNECT", config);
                    if (IO_BOX_CONNECT != null && IO_BOX_CONNECT.Split(';').Length >= 6)
                    {
                        string[] infos = IO_BOX_CONNECT.Split(';');
                        IOSerialPort = "COM" + infos[0];
                        IOBaudRate = infos[1];
                        IOStopBits = infos[4];
                        IODataBits = infos[2];
                        IOParity = infos[3];
                    }
                    OutputEnter = GetDescendants("OutputEnter", config);
                    DataOutputInterface = GetDescendants("DataOutputInterface", config);
                    IPI_WORKORDERTYPE_CHECK = GetDescendants("IPI_WORKORDERTYPE_CHECK", config);
                }
                else
                {
                    if (commonModel.UpdateConfig == "Y")
                    {
                        //int error = configHandler.DeleteConfigParameters(commonModel.APPTYPE);
                        //if (error == 0 || error == -3303 || error == -3302)
                        //{
                        //    WriteParameterToiTac(configHandler);
                        //}
                        string[] parametersValue = configHandler.GetParametersForScope(commonModel.APPTYPE);
                        if (parametersValue != null && parametersValue.Length > 0)
                        {
                            foreach (var parameterID in parametersValue)
                            {
                                configHandler.DeleteConfigParametersExt(parameterID);
                            }
                        }
                        WriteParameterToiTac(configHandler);
                    }
                    List<ConfigEntity> getvalues = configHandler.GetConfigData(commonModel.APPID, commonModel.APPTYPE, commonModel.Cluster, commonModel.Station);
                    if (getvalues != null)
                    {
                        foreach (var item in getvalues)
                        {
                            if (item != null)
                            {
                                string[] strs = item.PARAMETER_NAME.Split(new char[] { '.' });
                                dicConfig.Add(strs[strs.Length - 1], item.CONFIG_VALUE);
                            }
                        }
                    }

                    StationNumber = commonModel.Station;
                    Client = commonModel.Client;
                    RegistrationType = commonModel.RegisterType;
                    SerialPort = GetParameterValue("SerialPort");
                    BaudRate = GetParameterValue("BaudRate");
                    Parity = GetParameterValue("Parity");
                    StopBits = GetParameterValue("StopBits");
                    DataBits = GetParameterValue("DataBits");
                    NewLineSymbol = GetParameterValue("NewLineSymbol");
                    High = GetParameterValue("High");
                    Low = GetParameterValue("Low");
                    EndCommand = GetParameterValue("EndCommand");
                    DLExtractPattern = GetParameterValue("DLExtractPattern");
                    MBNExtractPattern = GetParameterValue("MBNExtractPattern");
                    EquipmentExtractPattern = GetParameterValue("EquipmentExtractPattern");
                    OpacityValue = GetParameterValue("OpacityValue");
                    LocationXY = GetParameterValue("LocationXY");
                    ThawingDuration = GetParameterValue("ThawingDuration");
                    ThawingCheck = GetParameterValue("ThawingCheck");
                    LockTime = GetParameterValue("LockOutTime");
                    UsageTime = GetParameterValue("UsageDurationSetting");
                    GateKeeperTimer = GetParameterValue("GateKeeperTimer");
                    SolderPasteValidity = GetParameterValue("SolderPasteValidity");
                    OpenControlBox = GetParameterValue("OpenControlBox");
                    StencilPrefix = GetParameterValue("StencilPrefix");
                    TimerSpan = GetParameterValue("TimerSpan");
                    StartTrigerStr = GetParameterValue("StartTrigerStr");
                    EndTrigerStr = GetParameterValue("EndTrigerStr");
                    NoRead = GetParameterValue("NoRead");
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message, ex);
            }
        }

        private string GetParameterValue(string parameterName)
        {
            if (dicConfig.ContainsKey(parameterName))
            {
                return dicConfig[parameterName];
            }
            else
            {
                return "";
            }
        }

        private void WriteParameterToiTac(ConfigManage configHandler)
        {
            GetApplicationDatas getData = new GetApplicationDatas();
            List<ParameterEntity> entityList = getData.GetApplicationEntity();
            string[] strs = GetParameterString(entityList);
            string[] strvalues = GetValueString(entityList);
            if (strs != null && strs.Length > 0)
            {
                int errorCode = configHandler.CreateConfigParameter(strs);
                if (errorCode == 0 || errorCode == 5)
                {
                    CommonModel commonModel = ReadIhasFileData.getInstance();
                    int re = configHandler.UpdateParameterValues(commonModel.APPID, commonModel.APPTYPE, commonModel.Cluster, commonModel.Station, strvalues);
                }
            }

            //if (entityList.Count > 0)
            //{
            //    List<ParameterEntity> entitySubList = null;
            //    CommonModel commonModel = ReadIhasFileData.getInstance();
            //    foreach (var entity in entityList)
            //    {
            //        entitySubList = new List<ParameterEntity>();
            //        entitySubList.Add(entity);
            //        string[] strs = GetParameterString(entitySubList);
            //        string[] strvalues = GetValueString(entitySubList);
            //        if (strs != null && strs.Length > 0)
            //        {
            //            int errorCode = configHandler.CreateConfigParameter(strs);
            //            if (errorCode == 0 || errorCode == 5)
            //            {                           
            //                int re = configHandler.UpdateParameterValues(commonModel.APPID, commonModel.APPTYPE, commonModel.Cluster, commonModel.Station, strvalues);
            //            }
            //            else if (errorCode == -3301)//Parameter already exists
            //            {
            //                int re = configHandler.UpdateParameterValues(commonModel.APPID, commonModel.APPTYPE, commonModel.Cluster, commonModel.Station, strvalues);
            //            }
            //        }
            //    }
            //}
        }

        private string[] GetParameterString(List<ParameterEntity> entityList)
        {
            List<string> strList = new List<string>();
            foreach (var entity in entityList)
            {
                strList.Add(entity.PARAMETER_DESCRIPTION);
                strList.Add(entity.PARAMETER_DIMPATH);
                strList.Add(entity.PARAMETER_DISPLAYNAME);
                strList.Add(entity.PARAMETER_NAME);
                strList.Add(entity.PARAMETER_PARENT_NAME);
                strList.Add(entity.PARAMETER_SCOPE);
                strList.Add(entity.PARAMETER_TYPE_NAME);
            }
            return strList.ToArray();
        }

        private string[] GetValueString(List<ParameterEntity> entityList)
        {
            List<string> strList = new List<string>();
            foreach (var entity in entityList)
            {
                if (entity.PARAMETER_VALUE == "")
                    continue;
                strList.Add(entity.PARAMETER_VALUE);
                strList.Add(entity.PARAMETER_NAME);

            }
            return strList.ToArray();
        }

        private string GetDescendants(string parameter, XDocument _config)
        {
            try
            {
                string value = _config.Descendants(parameter).First().Value;

                return value;
            }
            catch
            {
                LogHelper.Error("Parameter is not exist." + parameter);
                return "";
            }
        }
    }
}
