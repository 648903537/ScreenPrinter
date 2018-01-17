using com.amtec.forms;
using com.amtec.model;
using com.itac.mes.imsapi.client.dotnet;
using com.itac.mes.imsapi.domain.container;
using System.Collections.Generic;

namespace com.amtec.action
{
    public class GetSerialNumberInfo
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private IMSApiSessionContextStruct sessionContext;
        private InitModel init;
        private MainView view;

        public GetSerialNumberInfo(IMSApiSessionContextStruct sessionContext, InitModel init, MainView view)
        {
            this.sessionContext = sessionContext;
            this.init = init;
            this.view = view;
        }

        public string[] GetSNInfo(string serialNumber)
        {
            int error = 0;
            string errorMsg = "";
            string[] serialNumberResultKeys = new string[] { "PART_DESC", "PART_NUMBER", "WORKORDER_NUMBER" };
            string[] serialNumberResultValues = new string[] { };
            LogHelper.Info("begin api trGetSerialNumberInfo (serial number =" + serialNumber + ")");
            error = imsapi.trGetSerialNumberInfo(sessionContext, init.configHandler.StationNumber, serialNumber, "-1", serialNumberResultKeys, out serialNumberResultValues);
            LogHelper.Info("end api trGetSerialNumberInfo (result code = " + error + ")");
            //imsapi.imsapiGetErrorText(sessionContext, error, out errorMsg);
            errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            if (error == 0)
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trGetSerialNumberInfo " + error, "");
            }
            else
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " trGetSerialNumberInfo " + error + "," + errorMsg, "");
            }
            return serialNumberResultValues;
        }

        public string[] GetSNHistoryInfo(string serialNumber)
        {
            int error = 0;
            string[] bookingResultKeys = new string[] { "BOOK_DATE", "STATION_NUMBER" };
            string[] bookingResultValues = new string[] { };
            string[] failureDataResultKeys = new string[] { };
            string[] failureDataResultValues = new string[] { };
            string[] failureSlipDataResultKeys = new string[] { };
            string[] failureSlipDataResultValues = new string[] { };
            string[] measureDataResultKeys = new string[] { };
            string[] measureDataResultValues = new string[] { };
            string workOrderNumber = "";
            string partNumber = "";
            string customerPartNumber = "";
            string partDesc = "";
            string quantity = "";
            long lastReportDate = 0;
            string lotNumber = "";
            int isLocked = 0;
            error = imsapi.trGetSerialNumberHistoryData(sessionContext, init.configHandler.StationNumber, serialNumber, "-1", init.currentSettings.processLayer, 0, 0, bookingResultKeys, out bookingResultValues, failureDataResultKeys, out failureDataResultValues, failureSlipDataResultKeys, out failureSlipDataResultValues, measureDataResultKeys, out measureDataResultValues, out workOrderNumber, out partNumber, out customerPartNumber, out partDesc, out quantity, out lastReportDate, out lotNumber, out isLocked);
            LogHelper.Info("end api trGetSerialNumberHistoryData (serialNumber=" + serialNumber + ",result code = " + error + ")");
            return bookingResultValues;
        }
        public int SwitchSerialNumber(SerialNumberData[] snArray, string datetime)
        {
            List<SwitchSerialNumberData> listsnData = new List<SwitchSerialNumberData>();
            string refSerialnumber = "";
            foreach (var item in snArray)
            {
                refSerialnumber = item.serialNumber.Substring(0, item.serialNumber.Length - 3);
                SwitchSerialNumberData snData = new SwitchSerialNumberData(0, item.serialNumber + "_" + datetime, "-1", item.serialNumber, 0);
                listsnData.Add(snData);
            }
            if (snArray.Length > 1)
            {
                SwitchSerialNumberData refsnData = new SwitchSerialNumberData(0, refSerialnumber + "_" + datetime, "-1", refSerialnumber, 0);
                listsnData.Add(refsnData);
            }

            SwitchSerialNumberData[] serialNumberArray = new SwitchSerialNumberData[] { };
            serialNumberArray = listsnData.ToArray();
            int error = imsapi.trSwitchSerialNumber(sessionContext, init.configHandler.StationNumber, "-1", "-1", ref serialNumberArray);
            LogHelper.Info("end api trSwitchSerialNumber (serialNumber+" + refSerialnumber + "result code = " + error + ")");
            return error;
        }

        public SerialNumberData[] GetSerialNumberBySNRef(string serialNumberRef)
        {
            int error = 0;
            SerialNumberData[] serialNumberArray = new SerialNumberData[] { };
            error = imsapi.trGetSerialNumberBySerialNumberRef(sessionContext, init.configHandler.StationNumber, serialNumberRef, "-1", out serialNumberArray);
            LogHelper.Info("api trGetSerialNumberBySerialNumberRef (serial number ref = " + serialNumberRef + ",result code = " + error);
            return serialNumberArray;
        }

        public int GetSerialNumberByref(string serialNumber)
        {
            SerialNumberData[] serialNumArray = new SerialNumberData[] { };
            int errorSubSN = imsapi.trGetSerialNumberBySerialNumberRef(sessionContext, init.configHandler.StationNumber, serialNumber, "-1", out serialNumArray);
            LogHelper.Info("trGetSerialNumberBySerialNumberRef(): SN:" + serialNumber + ",ResultCode :" + errorSubSN + "");
            return errorSubSN;
        }
    }
}
