using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MagneticField
{
    //翠海采集卡管理类
    internal class ProbeTcpClientCH : IDisposable
    {
        private TcpClient _tcpClient;
        private NetworkStream _networkStream;
        private readonly string _ip;
        private readonly int _port;
        private readonly Dictionary<int, ProbeInfo> _channelProbeMap = new Dictionary<int, ProbeInfo>();
        private readonly List<ProbeInfo> _probeList;
        private Thread _receiveThread;
        private bool _isRunning;

        public ProbeTcpClientCH(string ip, int port, List<ProbeInfo> probeList)
        {
            _ip = ip;
            _port = port;
            _probeList = probeList;
        }

        public void AssignProbe(int channel, ProbeInfo probe)
        {
            if (probe == null) return;
            _channelProbeMap[channel] = probe;
        }

        public bool Connect()
        {
            //try
            //{
            //    _tcpClient = new TcpClient();
            //    _tcpClient.Connect(_ip, _port);
            //    _networkStream = _tcpClient.GetStream();

            //    _isRunning = true;
            //    _receiveThread = new Thread(ReceiveData)
            //    {
            //        IsBackground = true
            //    };
            //    _receiveThread.Start();

            //    return true;
            //}
            //catch
            //{
            //    return false;
            //}
            try
            {
                _tcpClient = new TcpClient();
                var connectTask = _tcpClient.ConnectAsync(_ip, _port);

                // 添加3秒超时
                if (!connectTask.Wait(3000))
                {
                    _tcpClient.Close();
                    return false;
                }

                if (!_tcpClient.Connected)
                    return false;

                _networkStream = _tcpClient.GetStream();

                _isRunning = true;
                _receiveThread = new Thread(ReceiveData)
                {
                    IsBackground = true
                };
                _receiveThread.Start();

                return true;
            }
            catch
            {
                _tcpClient?.Close();
                return false;
            }
        }

        private void ReceiveData()
        {
            while (_isRunning)
            {
                try
                {
                    // 读取64字节数据帧
                    byte[] buffer = new byte[64];
                    int bytesRead = _networkStream.Read(buffer, 0, buffer.Length);

                    if (bytesRead == 64)
                    {
                        UpdateProbeData(buffer);
                    }
                }
                catch
                {
                    // 连接异常处理
                    Thread.Sleep(1000);
                }
            }
        }

        private void UpdateProbeData(byte[] buffer)
        {

            // 每12字节对应一个探头的XYZ值 (4字节X + 4字节Y + 4字节Z)
            for (int probeOffset = 0; probeOffset < 4; probeOffset++)
            {
                int channel = probeOffset + 1; // 通道1-4对应探头1-4或5-8
                if (!_channelProbeMap.TryGetValue(channel, out var probe)) continue;

                // 计算当前探头的字节偏移量 (每个探头12字节)
                int byteOffset = probeOffset * 12;

                // 调试输出原始字节
                byte[] xBytes = new byte[] { buffer[byteOffset], buffer[byteOffset + 1], buffer[byteOffset + 2], buffer[byteOffset + 3] };
                byte[] yBytes = new byte[] { buffer[byteOffset + 4], buffer[byteOffset + 5], buffer[byteOffset + 6], buffer[byteOffset + 7] };
                byte[] zBytes = new byte[] { buffer[byteOffset + 8], buffer[byteOffset + 9], buffer[byteOffset + 10], buffer[byteOffset + 11] };

                // 解析为整数
                int xValue = BitConverter.ToInt32(xBytes, 0);
                int yValue = BitConverter.ToInt32(yBytes, 0);
                int zValue = BitConverter.ToInt32(zBytes, 0);


                // 解析X值 (小端模式)
                probe.X = xValue * 107.5 / 2147483647.0 * 1000;

                // 解析Y值
                probe.Y = yValue * 107.5 / 2147483647.0 * 1000;

                // 解析Z值
                probe.Z = zValue * 107.5 / 2147483647.0 * 1000;

            }
        }

        public bool SendStartCommand(bool isMaster, int sampleRate)
        {
            try
            {
                var command = CreateStartCommand(isMaster, sampleRate);
                _networkStream.Write(command, 0, command.Length);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool SendStopCommand()
        {
            try
            {
                var command = CreateStopCommand();
                _networkStream.Write(command, 0, command.Length);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static byte[] CreateStartCommand(bool isMaster, int sampleRate)
        {
            byte[] command = new byte[8];
            command[0] = 0x9F; // 启动采集指令

            // 设置主从卡标志
            command[4] = isMaster ? (byte)0x00 : (byte)0x01;

            // 设置采样率 (小端模式)
            command[6] = (byte)(sampleRate & 0xFF);        // N1 低字节
            command[7] = (byte)((sampleRate >> 8) & 0xFF); // N2 高字节

            return command;
        }

        public static byte[] CreateStopCommand()
        {
            return new byte[] { 0x38 }; // 停止指令可以只发1个字节
        }

        public void Disconnect()
        {
            _isRunning = false;
            _receiveThread?.Join(500);
            _networkStream?.Close();
            _tcpClient?.Close();
        }

        public void Dispose()
        {
            Disconnect();
        }

        public void ClearChannelMappings()
        {
            _channelProbeMap.Clear();
        }
    }

}
