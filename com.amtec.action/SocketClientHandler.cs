using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Sockets;
using System.Threading;
using System.Net;
using com.amtec.forms;
using ShippingClient;

namespace com.amtec.action
{
    public class SocketClientHandler
    {
        public Socket tcpsend; //发送创建套接字 
        public bool connect_flag = true;
        byte[] buffer = new byte[1024];//设置一个缓冲区，用来保存数据
        public ManualResetEvent connectDone = new ManualResetEvent(false); //连接的信号     
        MainView view = null;

        public SocketClientHandler(MainView _view)
        {
            view = _view;
        }

        public bool connect(string address, string port)
        {
            try
            {
                tcpsend = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);//初始化套接字
                IPEndPoint remotepoint = new IPEndPoint(IPAddress.Parse(address), Convert.ToInt32(port));//根据ip地址和端口号创建远程终结点
                EndPoint end = (EndPoint)remotepoint;
                view.errorHandler(0, "Start Server connection. IP:" + remotepoint.Address + " Port:" + remotepoint.Port, "");
                //tcpsend.BeginConnect(end, new AsyncCallback(ConnectedCallback), tcpsend); //调用回调函数
                tcpsend.Connect(end);
                view.errorHandler(0, "Connected to Server success", "");
                Thread.Sleep(200);
                tcpsend.BeginReceive(buffer, 0, buffer.Length, SocketFlags.None, new AsyncCallback(ReceiveMessage), tcpsend);
                //connectDone.WaitOne();
                return true;
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message, ex);
                view.errorHandler(3, "Connected to Server error.", "");
                return false;
            }
        }

        private void ConnectedCallback(IAsyncResult ar)
        {
            Socket client = (Socket)ar.AsyncState;
            try
            {
                client.EndConnect(ar);
                view.errorHandler(0, "Connected to Server.", "");
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message + "," + ex.StackTrace);
                view.errorHandler(2, "Connection Error. IP Cannot be reached.", "");
            }
        }

        public void ReceiveMessage(IAsyncResult ar)
        {
            try
            {
                if (tcpsend != null && tcpsend.Connected)
                {
                    var socket = ar.AsyncState as Socket;
                    var length = socket.EndReceive(ar);
                    if (length == 0)
                    {
                        view.SetConnectionText(1, "Disconnect to server, please check");
                        return;
                    }
                    var message = Encoding.GetEncoding("UTF-8").GetString(buffer, 0, length);
                    view.ProcessPFCMessage(message);
                    //接收下一个消息(因为这是一个递归的调用，所以这样就可以一直接收消息了）
                    socket.BeginReceive(buffer, 0, buffer.Length, SocketFlags.None, new AsyncCallback(ReceiveMessage), socket);
                }

            }
            catch (Exception ex)
            {
                //LogHelper.Error(ex.Message, ex);
            }
        }

        public bool send(string data)
        {
            try
            {
                int length = data.Length;
                Byte[] Bysend = new byte[length];
                Bysend = System.Text.Encoding.GetEncoding("UTF-8").GetBytes(data); //将字符串指定到指定Byte数组
                tcpsend.BeginSend(Bysend, 0, Bysend.Length, 0, new AsyncCallback(SendCallback), tcpsend); //异步发送数据
                return true;
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message, ex);
                return false;
            }
        }

        private void SendCallback(IAsyncResult ar) //发送的回调函数
        {
            Socket client = (Socket)ar.AsyncState;
            int bytesSend = client.EndSend(ar); //完成发送
        }

        public void CloseSocket()
        {
            tcpsend.Close();
        }
    }
}
