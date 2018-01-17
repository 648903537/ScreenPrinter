using com.amtec.action;
using com.amtec.forms;
using com.amtec.model;
using System;
using System.IO.Ports;

namespace com.amtec.device
{
    public class ScannerHeandler
    {
        private SerialPort serialPort;
        private InitModel init;
        private MainView view;
        private SerialPort IOserialPort;

        public ScannerHeandler(InitModel init, MainView view)
        {
            this.init = init;
            this.view = view;
            if (init.configHandler.SerialPort != "" && init.configHandler.SerialPort != null)
            {
                serialPort = new SerialPort();
                serialPort.PortName = init.configHandler.SerialPort;
                serialPort.BaudRate = int.Parse(init.configHandler.BaudRate);
                serialPort.Parity = (Parity)int.Parse(init.configHandler.Parity);
                serialPort.StopBits = (StopBits)1;
                serialPort.Handshake = Handshake.None;
                serialPort.DataBits = int.Parse(init.configHandler.DataBits);
                serialPort.NewLine = "\r";
            }

            if (init.configHandler.IOSerialPort != "" && init.configHandler.IOSerialPort != null)
            {
                IOserialPort = new SerialPort();
                IOserialPort.PortName = init.configHandler.IOSerialPort;
                IOserialPort.BaudRate = int.Parse(init.configHandler.IOBaudRate);
                IOserialPort.Parity = (Parity)int.Parse(init.configHandler.IOParity);
                IOserialPort.StopBits = (StopBits)1;
                IOserialPort.Handshake = Handshake.None;
                IOserialPort.DataBits = int.Parse(init.configHandler.IODataBits);
                IOserialPort.NewLine = "\r";
            }
        }
        public void SetSerialPortData(SerialPort setSP)
        {
            serialPort = setSP;
        }
        public SerialPort handler()
        {
            return serialPort;
        }

        public SerialPort handler2()
        {
            return IOserialPort;
        }

        public void endCommand()
        {
            char[] charArray;
            String text = init.configHandler.EndCommand;
            String tmpString = text.Trim();
            tmpString = tmpString.Replace("ESC", "*");
            charArray = tmpString.ToCharArray();

            for (int i = 0; i < charArray.Length; i++)
            {
                if (charArray[i].Equals((char)42))
                {
                    charArray[i] = (char)27;
                }
            }

            try
            {
                serialPort.Write(charArray, 0, charArray.Length);
                LogHelper.Info("Send end command:" + text);
                view.errorHandler(0, "Send end command", "Send end command");
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex);
                view.errorHandler(0, "Send end command error", "Send end command error");
            }
        }

        public void sendHigh()
        {
            char[] charArray;
            String text = init.configHandler.High;
            String tmpString = text.Trim();
            tmpString = tmpString.Replace("ESC", "*");
            charArray = tmpString.ToCharArray();

            for (int i = 0; i < charArray.Length; i++)
            {
                if (charArray[i].Equals((char)42))
                {
                    charArray[i] = (char)27;
                }
            }

            try
            {
                serialPort.Write(charArray, 0, charArray.Length);
                view.errorHandler(0, "SEND HIGH", "SEND HIGH");
                LogHelper.Info("SEND HIGH:" + text);
            }
            catch (Exception ex)
            {
                view.errorHandler(2, "SEND HIGH ERROR", "SEND HIGH ERROR");
                LogHelper.Error(ex);
            }
        }

        public void sendHighExt()
        {
            String text = init.configHandler.High;
            String tmpString = text.Trim();
            byte[] bytes = null;//System.Text.Encoding.UTF8.GetBytes(tmpString);

            try
            {
                bytes = CodeConversionManager.strToToHexByte(tmpString);
                serialPort.Write(bytes, 0, bytes.Length);
                view.errorHandler(0, "SEND HIGH", "SEND HIGH");
                LogHelper.Info("SEND HIGH:" + text);
            }
            catch (Exception ex)
            {
                view.errorHandler(2, "SEND HIGH ERROR", "SEND HIGH ERROR");
                LogHelper.Error(ex);
            }
        }

        public void sendLowExt()
        {
            String text = init.configHandler.Low;
            String tmpString = text.Trim();
            byte[] bytes = null;//System.Text.Encoding.UTF8.GetBytes(tmpString);

            try
            {
                bytes = CodeConversionManager.strToToHexByte(tmpString);
                serialPort.Write(bytes, 0, bytes.Length);
                view.errorHandler(0, "SEND LOW", "SEND LOW");
                LogHelper.Info("SEND LOW:" + text);
            }
            catch (Exception ex)
            {
                view.errorHandler(2, "SEND LOW ERROR", "SEND LOW ERROR");
                LogHelper.Error(ex);
            }
        }


        public void sendLow()
        {
            char[] charArray;
            String text = init.configHandler.Low;
            String tmpString = text.Trim();
            tmpString = tmpString.Replace("ESC", "*");
            charArray = tmpString.ToCharArray();
            for (int i = 0; i < charArray.Length; i++)
            {
                if (charArray[i].Equals((char)42))
                {
                    charArray[i] = (char)27;
                }
            }

            try
            {
                serialPort.Write(charArray, 0, charArray.Length);
                view.errorHandler(0, "SEND LOW", "SEND LOW");
                LogHelper.Info("SEND LOW:" + text);
            }
            catch (Exception ex)
            {
                view.errorHandler(2, "SEND LOW ERROR", "SEND LOW ERROR");
                LogHelper.Error(ex);
            }
        }

        public void sendStartTrigger()
        {
            char[] charArray;
            String text = init.configHandler.StartTrigerStr;
            String tmpString = text.Trim();
            tmpString = tmpString.Replace("ESC", "*");
            charArray = tmpString.ToCharArray();
            for (int i = 0; i < charArray.Length; i++)
            {
                if (charArray[i].Equals((char)42))
                {
                    charArray[i] = (char)27;
                }
            }

            try
            {
                serialPort.Write(charArray, 0, charArray.Length);
                view.errorHandler(0, "SEND START TRIGGER", "START TRIGGER");
                LogHelper.Info("SEND START TRIGGRT:" + text);
            }
            catch
            {
                view.errorHandler(2, "SEND START TRIGGER ERROR", "SEND START TRIGGER ERROR");
                LogHelper.Info("SEND START TRIGGER ERROR");
            }
        }

        public void sendEndTrigger()
        {
            char[] charArray;
            String text = init.configHandler.EndTrigerStr;
            String tmpString = text.Trim();
            tmpString = tmpString.Replace("ESC", "*");
            charArray = tmpString.ToCharArray();
            for (int i = 0; i < charArray.Length; i++)
            {
                if (charArray[i].Equals((char)42))
                {
                    charArray[i] = (char)27;
                }
            }

            try
            {
                serialPort.Write(charArray, 0, charArray.Length);
                view.errorHandler(0, "SEND END TRIGGER", "END TRIGGER");
                LogHelper.Info("SEND END TRIGGRT:" + text);
            }
            catch
            {
                view.errorHandler(2, "SEND END TRIGGRT ERROR", "SEND END TRIGGRT ERROR");
                LogHelper.Info("SEND END TRIGGRT ERROR");
            }
        }

        public void sendBlockBoard()
        {
            string signlText = "1B 5B 43 1B 5B 42 1B 4F 41 30 1B 49 41 23";
            try
            {
                serialPort.Write(signlText);
                view.errorHandler(0, "SEND BLOCK BOARD SIGNAL", "END TRIGGER");
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex);
                view.errorHandler(2, "SEND BLOCK BOARD SIGNAL ERROR", "END TRIGGER");
            }
        }

        public void sendRequireBoard()
        {
            string signlText = "1B 5B 43 1B 5B 42 1B 4F 41 31 1B 49 41 23";
            try
            {
                serialPort.Write(signlText);
                view.errorHandler(0, "SEND REQUIRE BOARD SIGNAL", "END TRIGGER");
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex);
                view.errorHandler(2, "SEND REQUIRE BOARD SIGNAL ERROR", "END TRIGGER");
            }
        }
    }
}
