using com.amtec.forms;
using com.amtec.model;
using com.itac.mes.imsapi.client.dotnet;
using com.itac.mes.imsapi.domain.container;

namespace com.amtec.action
{
    public class SetUpManager
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private IMSApiSessionContextStruct sessionContext;
        private InitModel init;
        private MainView view;

        public SetUpManager(IMSApiSessionContextStruct sessionContext, InitModel init, MainView view)
        {
            this.sessionContext = sessionContext;
            this.init = init;
            this.view = view;
        }

        public int UpdateMaterialSetUpByBin(int processLayer, string workorderNumber, string materialBinNumber, string materialQty, string partNumber, string setupName, string setupPos)
        {
            int error = 0;
            string[] materialSetupUploadKeys = new string[] { "ERROR_CODE", "MATERIAL_BIN_NUMBER", "MATERIAL_BIN_QTY_TOTAL", "PART_NUMBER", "SETUP_POSITION", "SETUP_STATE" };
            string[] materialSetupUploadValues = new string[] { "0", materialBinNumber, materialQty, partNumber, setupPos, "0" };
            string[] compPositionsUploadKeys = new string[] { "COMP_REFERENCE" };
            string[] compPositionsUploadValues = new string[] { };
            string[] materialSetupResultValues = new string[] { };
            string[] compPositionsResultValues = new string[] { };
            LogHelper.Info("begin api setupUpdateMaterialSetup (material bin number:" + materialBinNumber + ")");
            error = imsapi.setupUpdateMaterialSetup(sessionContext, init.configHandler.StationNumber, processLayer, workorderNumber, "-1", setupName, materialSetupUploadKeys
                , materialSetupUploadValues, compPositionsUploadKeys, compPositionsUploadValues, out materialSetupResultValues, out compPositionsResultValues);
            LogHelper.Info("end api setupUpdateMaterialSetup (result code = " + error + ")");
            if (error != 0)
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " setupUpdateMaterialSetup " + error, "");
            }
            else
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " setupUpdateMaterialSetup " + error, "");
            }
            return error;
        }

        public int SetupStateChange(string workorder, int activateFlag)
        {
            int error = 0;
            //0 = Activate setup
            //1 = Deactivate setup
            //2 = Delete setup
            error = imsapi.setupStateChange(sessionContext, init.configHandler.StationNumber, init.currentSettings.processLayer, workorder, "-1", -1, activateFlag);
            if (error != 0)
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " setupStateChange " + error, "");
            }
            else
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " setupStateChange " + error, "");
            }
            return error;
        }
    }
}
