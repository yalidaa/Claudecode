using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;

namespace MagneticField
{
    //华舜采集卡管理类
    internal class ProbeTcpClientHS
    {
        private TcpClient _tcpClient;
        private NetworkStream _networkStream;
        private readonly string _ip;
        private readonly int _port;
        private List<ProbeInfo> _probesInThisConnection;
        private Thread _receiveThread;
        private bool _isRunning;

        //奇数通道对应传感器1，偶数通道对应传感器2
        private ProbeInfo _sensor1Probe;
        private ProbeInfo _sensor2Probe;

        private string xTransformationType = "X";
        private string yTransformationType = "Y";
        private string zTransformationType = "Z";


        // 新增：数据推送接口
        private readonly IDataSink _dataSink;
        // 公开属性，便于主窗体获取探头信息
        public ProbeInfo Sensor1Probe => _sensor1Probe;
        public ProbeInfo Sensor2Probe => _sensor2Probe;


        public void ChangeTransform(string x, string y, string z)
        {
            xTransformationType = x;
            yTransformationType = y;
            zTransformationType = z;
        }


        public ProbeTcpClientHS(string ip, int port, List<ProbeInfo> probes, IDataSink dataSink)
        {
            _ip = ip;
            _port = port;
            _probesInThisConnection = probes;
            _dataSink = dataSink;

            // 奇数通道对应传感器1，偶数通道对应传感器2
            foreach (var probe in _probesInThisConnection)
            {
                if (probe.ProbeChannel % 2 == 1) // 奇数
                {
                    _sensor1Probe = probe;
                }
                else // 偶数
                {
                    _sensor2Probe = probe;
                }
            }
        }

        public bool Connect()
        {
            try
            {
                _tcpClient = new TcpClient();
                var connectTask = _tcpClient.ConnectAsync(_ip, _port);

                // 1秒超时
                if (!connectTask.Wait(1000))
                {
                    _tcpClient.Close();
                    return false;
                }

                if (!_tcpClient.Connected)
                    return false;

                _networkStream = _tcpClient.GetStream();
                _isRunning = true;



                SendCommand("start\n");

                _receiveThread = new Thread(ReceiveData);
                _receiveThread.IsBackground = true;
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
            byte[] buffer = new byte[1024];

            while (_isRunning)
            {
                try
                {
                    if (_networkStream.DataAvailable)
                    {
                        int bytesRead = _networkStream.Read(buffer, 0, buffer.Length);
                        if (bytesRead > 0)
                        {
                            byte[] receivedData = new byte[bytesRead];
                            Array.Copy(buffer, receivedData, bytesRead);
                            ParseFluxgateData(receivedData);
                        }
                    }
                    else
                    {
                        //Thread.Sleep(10); // 减少CPU占用
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"接收数据异常: {ex.Message}");
                    Thread.Sleep(1000);
                }
            }
        }

        private void ParseFluxgateData(byte[] data)
        {
            // 检查帧头
            if (data[0] != 0x46 || data[1] != 0x4D)
            {
                Console.WriteLine("无效帧头");
                return;
            }

            // 解析递增序列 (字节4-5，小端模式)
            ushort sequence = (ushort)(data[4] | (data[5] << 8));

            // 传感器1的数据 (奇数通道探头)
            if (_sensor1Probe != null)
            {
                ParseSensorData(data, 8, _sensor1Probe, sequence);
            }

            // 传感器2的数据 (偶数通道探头)
            if (_sensor2Probe != null)
            {
                ParseSensorData(data, 20, _sensor2Probe, sequence);
            }

        }

        private void ParseSensorData(byte[] data, int startIndex, ProbeInfo probe, ushort sequence)
        {
            double kc = 10; // 校准系数

            // 解析X通道数据 (24位，小端)
            int xData = data[startIndex] | (data[startIndex + 1] << 8) | (data[startIndex + 2] << 16);
            double xValue = CalculateMagneticField(xData, kc);

            // 解析Y通道数据
            int yData = data[startIndex + 4] | (data[startIndex + 5] << 8) | (data[startIndex + 6] << 16);
            double yValue = CalculateMagneticField(yData, kc);

            // 解析Z通道数据
            int zData = data[startIndex + 8] | (data[startIndex + 9] << 8) | (data[startIndex + 10] << 16);
            double zValue = CalculateMagneticField(zData, kc);

            // 更新写入探头值
            //probe.X = xValue; 
            //probe.Y = yValue;
            //probe.Z = zValue;
            double[] newProbeXYZ = GetValueByType(xTransformationType, yTransformationType, zTransformationType, xValue, yValue, zValue);
            probe.X = newProbeXYZ[0];
            probe.Y = newProbeXYZ[1];
            probe.Z = newProbeXYZ[2];


            // 推送 XYZ 给记录器
            var record = new DataRecord
            {
                Sequence = sequence,
                X = newProbeXYZ[0],
                Y = newProbeXYZ[1],
                Z = newProbeXYZ[2],
                Timestamp = DateTime.Now
            };

            _dataSink?.OnDataReceived(probe.ProbeName, record);

        }

        private double CalculateMagneticField(int rawData, double kc)
        {
            const int maxPositive = 8388607; // 2^23 - 1
            const double fullScale = 8388607.0; // 2^24 / 2
            const double factor = 2.5 / fullScale * 1000 * 4;

            // 24位有符号转换
            double signedValue;
            if (rawData < maxPositive)
            {
                signedValue = rawData;
            }
            else
            {
                signedValue = rawData - 16777216; // 2^24
            }

            // 计算磁场值: Mag = data * (2.5/8388607) * 1000 * 4 * Kc
            return signedValue * factor * kc;
        }


        //通道坐标转换
        private double[] GetValueByType(string type1, string type2, string type3, double x, double y, double z)
        {
            double[] redouble = new double[3] { 0.0, 0.0, 0.0 };
            if (string.IsNullOrEmpty(type1)|| string.IsNullOrEmpty(type2)|| string.IsNullOrEmpty(type3))
                return redouble;

            //type = type.ToUpper();

            if (type1.StartsWith("-"))
            {
                string axis = type1.Substring(1);
                if (axis == "X") redouble[0] = -x;
                if (axis == "Y") redouble[1] = -x;
                if (axis == "Z") redouble[2] = -x;
            }
            if (type1 == "X") redouble[0] = x;
            if (type1 == "Y") redouble[1] = x;
            if (type1 == "Z") redouble[2] = x;

            if (type2.StartsWith("-"))
            {
                string axis = type2.Substring(1);
                if (axis == "X") redouble[0] = -y;
                if (axis == "Y") redouble[1] = -y;
                if (axis == "Z") redouble[2] = -y;
            }
            if (type2 == "X") redouble[0] = y;
            if (type2 == "Y") redouble[1] = y;
            if (type2 == "Z") redouble[2] = y;

            if (type3.StartsWith("-"))
            {
                string axis = type3.Substring(1);
                if (axis == "X") redouble[0] = -z;
                if (axis == "Y") redouble[1] = -z;
                if (axis == "Z") redouble[2] = -z;
            }
            if (type3 == "X") redouble[0] = z;
            if (type3 == "Y") redouble[1] = z;
            if (type3 == "Z") redouble[2] = z;

            return redouble;
        }





        public void SendCommand(string command)
        {
            if (!_isRunning || _tcpClient == null || !_tcpClient.Connected)
                return;

            try
            {
                // 确保命令以换行符结尾
                if (!command.EndsWith("\n"))
                    command += "\n";

                NetworkStream stream = _tcpClient.GetStream();
                byte[] data = Encoding.Default.GetBytes(command);
                stream.Write(data, 0, data.Length);
            }
            catch (Exception ex)
            {
                Disconnect();
            }
        }



        public void Disconnect()
        {
            _isRunning = false;

            // 等待接收线程结束
            _receiveThread?.Join(1000);

            _networkStream?.Close();
            _tcpClient?.Close();
        }

        public void Dispose()
        {
            Disconnect();
        }

    }


    // 主机从机信息类
    public class HostInfo
    {
        public string HostId { get; set; }  // 主机标识，如 "HS1", "HS2" 等
        public string Address { get; set; } // IP:Port
        public bool IsMainHost { get; set; } // 是否为主主机，否则为从主机
        public TcpClient Client { get; set; }
        public NetworkStream Stream { get; set; }
    }


    // 数据记录结构
    public struct DataRecord
    {
        public ushort Sequence { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
        public DateTime Timestamp { get; set; }
    }
}
