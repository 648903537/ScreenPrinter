using com.amtec.configurations;
using com.amtec.forms;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace com.amtec.action
{
    class SocketServerHandle
    {
        private MainView mv;
        public SocketServerHandle(MainView mv1)
        {
            mv = mv1;
        }

        public IPEndPoint tcplisener;//将网络端点表示为IP地址和端口号
        public Socket read;
        public Thread accept;//创建并控制线程
        public ManualResetEvent AcceptDone = new ManualResetEvent(false); //连接的信号
        public string _HostName = "";
        public string _IpAddress = "";

        internal string IpAddress
        {
            get
            {
                if (!string.IsNullOrEmpty(_IpAddress))
                    return _IpAddress;

                var ips = Dns.GetHostAddresses(HostName);
                foreach (var ip in ips)
                {
                    if (ip.IsIPv6LinkLocal)
                        continue;

                    return ip.ToString();
                }
                return "";
            }
        }

        internal string HostName
        {
            get
            {
                return !string.IsNullOrEmpty(_HostName) ? _HostName : Dns.GetHostName();
            }
        }

        /// <summary>
        /// Open port
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void OpenPort(ApplicationConfiguration config)
        {
            string ipaddress = config.IPAddress;
            string port = config.Port;
            IPAddress ip = IPAddress.Parse(ipaddress.Trim());
            //用指定ip和端口号初始化
            tcplisener = new IPEndPoint(ip, Convert.ToInt32(port.Trim()));
            //创建一个socket对象
            read = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            try
            {
                read.Bind(tcplisener); //绑定
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message + "," + ex.StackTrace);
                mv.errorHandler(2, "Socket server run error", "Error");
                return;
            }
            ///收到客户端的连接，建立新的socket并接受信息
            read.Listen(500); //开始监听            
            mv.errorHandler(0, "Server run success,  Wait client connection", "Success");
            accept = new Thread(new ThreadStart(Listen));
            accept.Start();
            GC.Collect();//即使垃圾回收
            GC.WaitForPendingFinalizers();
        }

        public void Listen()
        {
            Thread.CurrentThread.IsBackground = true; //后台线程
            try
            {
                while (true)
                {
                    AcceptDone.Reset();
                    read.BeginAccept(new AsyncCallback(AcceptCallback), read);  //异步调用                    
                    AcceptDone.WaitOne();
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message + ";" + ex.StackTrace);
                mv.errorHandler(3, "ReadCallback error", "Error");
            }
        }

        List<StateObject> clientList = new List<StateObject>();
        public void AcceptCallback(IAsyncResult ar) //accpet的回调处理函数
        {
            try
            {
                AcceptDone.Set();
                Socket temp_socket = (Socket)ar.AsyncState;
                Socket client = temp_socket.EndAccept(ar); //获取远程的客户端
                Control.CheckForIllegalCrossThreadCalls = false;
                IPEndPoint remotepoint = (IPEndPoint)client.RemoteEndPoint;//获取远程的端口
                string remoteaddr = remotepoint.Address.ToString();        //获取远程端口的ip地址   
                mv.errorHandler(0, "IP-" + remotepoint.Address + " Port-" + remotepoint.Port + " has connection.", "Success");
                StateObject state = new StateObject();
                state.workSocket = client;
                if (!clientList.Contains(state))
                    clientList.Add(state);
                client.BeginReceive(state.buffer, 0, StateObject.BufferSize, 0, new AsyncCallback(ReadCallback), state);
                
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message + ";" + ex.StackTrace);
                mv.errorHandler(2, ex.Message + ";" + ex.StackTrace, "");
            }
        }

        public void ReadCallback(IAsyncResult ar)
        {
            String content = String.Empty;
            string responseStr = "";
            // Retrieve the state object and the handler socket
            // from the asynchronous state object.
            StateObject state = (StateObject)ar.AsyncState;
            Socket handler = state.workSocket;
            IPEndPoint remotepoint = (IPEndPoint)handler.RemoteEndPoint;
            // Read data from the client socket. 
            int bytesRead = 0;
            try
            {
                bytesRead = handler.EndReceive(ar);
            }
            catch (Exception)
            {
                if (!handler.Connected)
                {
                    mv.errorHandler(0, "IP-" + remotepoint.Address + " Port-" + remotepoint.Port + " stop connect.", "Success");
                    return;
                }
                return;
            }

            //client close pop msg todo
            if (bytesRead > 0)
            {
                try
                {
                    string contentstr = ""; //接收到的数据
                    contentstr += Encoding.GetEncoding("UTF-8").GetString(state.buffer, 0, bytesRead);
                    //WriteSocketToFile(contentstr);
                    string strSN = contentstr;//.Replace(strPex, "").Replace(strBex, "").Trim(); 
                    mv.errorHandler(0, "IP-" + remotepoint.Address + " Port-" + remotepoint.Port + " receive sn-" + strSN, "Success");
                    //responseStr = mv.ProcessSocketData(contentstr);
                    byte[] byteData = Encoding.GetEncoding("UTF-8").GetBytes(responseStr);//回发信息
                    handler.BeginSend(byteData, 0, byteData.Length, 0, new AsyncCallback(SendCallback), handler);

                    mv.errorHandler(0, "Send output." + responseStr, "success");//Scanned output更改为Send output                
                    handler.BeginReceive(state.buffer, 0, StateObject.BufferSize, 0, new AsyncCallback(ReadCallback), state);
                }
                catch (Exception ex)
                {
                    LogHelper.Error(ex.Message + ";" + ex.StackTrace);
                }
            }
        }

        private string[] ReadSocketFile()
        {
            string filePath = Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName;
            string _appDir = Path.GetDirectoryName(filePath);
            string[] strs = null;
            if (File.Exists(_appDir + @"\SOCKET.txt"))
            {
                strs = File.ReadAllLines(_appDir + @"\SOCKET.txt");
            }
            if (strs.Length > 0)
            {
                try
                {
                    File.Delete(_appDir + @"\SOCKET.txt");
                }
                catch (Exception ex)
                {

                    LogHelper.Error(ex.Message);
                }

            }
            return strs;
        }

        private void WriteSocketToFile(string strText)
        {
            string filePath = Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName;
            string _appDir = Path.GetDirectoryName(filePath);
            FileStream fs;
            StreamWriter sw;
            if (!File.Exists(_appDir + @"\SOCKET.txt"))
            {
                fs = new FileStream(_appDir + @"\SOCKET.txt", FileMode.OpenOrCreate);
            }
            else
            {
                fs = new FileStream(_appDir + @"\SOCKET.txt", FileMode.Append);
            }
            sw = new StreamWriter(fs);
            try
            {

                sw.WriteLine(strText);
                sw.Flush();
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message);
            }
            finally
            {
                sw.Close();
                fs.Close();
            }
        }

        private void Send(Socket handler, String data)
        {
            try
            {
                // Convert the string data to byte data using ASCII encoding.
                byte[] byteData = Encoding.GetEncoding("GB2312").GetBytes(data);
                // Begin sending the data to the remote device.
                handler.BeginSend(byteData, 0, byteData.Length, 0,
                    new AsyncCallback(SendCallback), handler);
            }
            catch (Exception ex)
            {
                mv.errorHandler(2, ex.Message + ";" + ex.StackTrace, "");
            }

        }

        private void SendCallback(IAsyncResult ar)
        {
            try
            {
                try
                {
                    // Retrieve the socket from the state object.
                    //Socket handler = (Socket)ar.AsyncState;

                    // Complete sending the data to the remote device.
                    //int bytesSent = handler.EndSend(ar);
                    //Console.WriteLine("Sent {0} bytes to client.", bytesSent);
                    //handler.Shutdown(SocketShutdown.Both);
                    //handler.Close();
                }
                catch (Exception ex)
                {
                    //mv.errorHandler(2, ex.Message + ";" + ex.StackTrace, "");
                    LogHelper.Error(ex.Message + ";" + ex.StackTrace);
                }
            }

            catch (Exception e)
            {
                //MessageBox.Show(e.ToString());
            }
        }

        public class StateObject
        {
            // Client  socket.
            public Socket workSocket = null;
            // Size of receive buffer.
            public const int BufferSize = 1024;
            // Receive buffer.
            public byte[] buffer = new byte[BufferSize];
            // Received data string.
            public StringBuilder sb = new StringBuilder();
        }
    }
}
