using com.amtec.action;
using com.amtec.forms;
using com.amtec.model;
using com.itac.mes.imsapi.client.dotnet;
using com.itac.mes.imsapi.domain.container;

namespace com.amtec.configurations
{
    public class SessionContextHeandler
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private IMSApiSessionValidationStruct sessionValidationStruct;
        private IMSApiSessionContextStruct sessionContext = null;
        private int initResult;
        private LoginForm mainView;

        public SessionContextHeandler(CommonModel config, LoginForm mainView)
        {
            this.mainView = mainView;
            initResult = imsapi.imsapiInit();

            if (initResult != 0)
            {
                mainView.SetStatusLabelText("Conncection to DMS failed", 1);
                mainView.isCanLogin = false;
                LogHelper.Info("Conncection to DMS failed");
            }
            else
            {
                mainView.SetStatusLabelText("Conncection to DMS established", 0);
                mainView.isCanLogin = true;
                LogHelper.Info("Conncection to DMS established");
            }
        }

        public IMSApiSessionContextStruct getSessionContext()
        {
            if (initResult != IMSApiDotNetConstants.RES_OK)
            {
                return null;
            }
            else
            {
                int result = imsapi.regLogin(sessionValidationStruct, out sessionContext);
                if (result != IMSApiDotNetConstants.RES_OK)
                {
                    return null;
                }
                else
                {
                    return sessionContext;
                }
            }
        }
    }
}
