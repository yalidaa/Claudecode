using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Sockets;
using System.Diagnostics;
using System.Windows.Forms;
using System.Net.Http;

namespace MagneticField
{
    internal class PowerTcpClientManager
    {
        private TcpClient _client;
        private string _ip;
        private int _port;
        public string _coilconstant;
        public string _zerofieldcurrent;
        private bool _isConnected = false;
        public string _xyz { get; set; }

        // 添加电压和电流属性
        public double LatestVoltage { get; set; }
        public double LatestCurrent { get; set; }


        public bool IsConnected => _isConnected;

        public PowerTcpClientManager(string ip, int port, string coilconstant, string zerofieldcurrent)
        {
            _ip = ip;
            _port = port;
            _coilconstant = coilconstant;
            _zerofieldcurrent = zerofieldcurrent;
            _client = new TcpClient();
        }

        public void ConnectAsync()
        {
            try
            {
                if (_isConnected)
                {
                    Disconnect();
                }

                //_client.Connect(_ip, _port);

                if (_client == null)
                {
                    _client = new TcpClient();
                }

                var connectTask = _client.ConnectAsync(_ip, _port);
                // 2秒超时
                if (!connectTask.Wait(2000))
                {
                    _client.Close();
                    _client = null; // 建议添加：超时关闭后置空
                    return;
                }

                _isConnected = true;
            }
            catch (Exception ex)
            {
                _isConnected = false;
                // 如果连接异常，确保清理
                if (_client != null)
                {
                    _client.Close();
                    _client = null;
                }
                MessageBox.Show($"连接 {_ip}:{_port} 失败: {ex.Message}");
            }
        }

        public void Disconnect()
        {
            if (_client != null)
            {
                try
                {
                    _client.Close();
                }
                catch (Exception)
                {
                    // 忽略关闭时的错误
                }
                finally
                {
                    _client = null; // 关键：置空防止再次使用或重复关闭
                }
            }
            _isConnected = false;
        }

        public async Task<string> SendCommandAndReceiveResponseAsync(string command)
        {
            if (!_isConnected || _client == null || !_client.Connected)
                return null;

            try
            {
                // 确保命令以换行符结尾
                if (!command.EndsWith("\n"))
                    command += "\n";

                NetworkStream stream = _client.GetStream();
                byte[] data = Encoding.Default.GetBytes(command);
                await stream.WriteAsync(data, 0, data.Length);

                // 读取响应
                byte[] buffer = new byte[1024];
                int bytesRead = await stream.ReadAsync(buffer, 0, buffer.Length);
                return Encoding.UTF8.GetString(buffer, 0, bytesRead).Trim();
            }
            catch (Exception ex)
            {
                Disconnect();
                return null;
            }
        }
        public async Task<string> SendCommandAndReceiveResponseAsyncForHex(string hexCommand)
        {
            if (!_isConnected || _client == null || !_client.Connected)
                return null;

            try
            {
                NetworkStream stream = _client.GetStream();

                hexCommand = hexCommand.Replace(" ", "").Replace("-", "");

                if (hexCommand.Length % 2 != 0)
                {
                    throw new ArgumentException("十六进制字符串长度必须是偶数");
                }

                byte[] bytes = new byte[hexCommand.Length / 2];
                for (int i = 0; i < bytes.Length; i++)
                {
                    string byteValue = hexCommand.Substring(i * 2, 2);
                    bytes[i] = Convert.ToByte(byteValue, 16);
                }

                await stream.WriteAsync(bytes, 0, bytes.Length);

                // 读取响应
                byte[] buffer = new byte[1024];
                int bytesRead = await stream.ReadAsync(buffer, 0, buffer.Length);
                return Encoding.UTF8.GetString(buffer, 0, bytesRead).Trim();
            }
            catch (Exception ex)
            {
                Disconnect();
                return null;
            }
        }

        public void SendCommand(string command)
        {
            if (!_isConnected || _client == null || !_client.Connected)
                return;
            
            try
            {
                // 确保命令以换行符结尾
                if (!command.EndsWith("\n"))
                    command += "\n";

                NetworkStream stream = _client.GetStream();
                byte[] data = Encoding.Default.GetBytes(command);
                stream.WriteAsync(data, 0, data.Length);
            }
            catch (Exception ex)
            {
                Disconnect();
            }
        }
    }
}
