using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MagneticField
{
    //探头数据CSV写入类
    internal class CsvDataRecorder
    {
        private string _csvFilePath;
        private StreamWriter _csvWriter;
        private bool _isRecording;
        private readonly object _lockObject = new object();
        private List<ProbeInfo> _allProbes;
        private DateTime _startTime;

        public CsvDataRecorder(List<ProbeInfo> allProbes)
        {
            _allProbes = allProbes.OrderBy(p => p.ProbeChannel).ToList();
        }

        public bool StartRecording(string basePath = null)
        {
            try
            {
                if (basePath == null)
                {
                    basePath = AppDomain.CurrentDomain.BaseDirectory;
                }

                // 创建Data文件夹
                string dataDir = Path.Combine(basePath, "Data");
                if (!Directory.Exists(dataDir))
                {
                    Directory.CreateDirectory(dataDir);
                }

                // 生成带时间戳的文件名
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                _csvFilePath = Path.Combine(dataDir, $"Probes_{timestamp}.csv");

                _csvWriter = new StreamWriter(_csvFilePath, false, Encoding.UTF8);
                _startTime = DateTime.Now;

                // 写入CSV表头
                WriteCsvHeader();

                _isRecording = true;
                Console.WriteLine($"开始记录数据到: {_csvFilePath}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建CSV文件失败: {ex.Message}");
                return false;
            }
        }

        private void WriteCsvHeader()
        {
            lock (_lockObject)
            {
                var headerFields = new List<string>();

                // 按通道顺序添加所有探头的XYZ字段
                foreach (var probe in _allProbes)
                {
                    headerFields.Add($"探头{probe.ProbeChannel}X");
                    headerFields.Add($"探头{probe.ProbeChannel}Y");
                    headerFields.Add($"探头{probe.ProbeChannel}Z");
                }

                // 写入表头
                _csvWriter.WriteLine(string.Join(",", headerFields));
                _csvWriter.Flush();
            }
        }

        public void RecordData()
        {
            if (!_isRecording) return;

            lock (_lockObject)
            {
                try
                {
                    var dataFields = new List<string>();

                    // 按通道顺序获取所有探头的XYZ值
                    foreach (var probe in _allProbes)
                    {
                        dataFields.Add(probe.X.ToString("F6"));
                        dataFields.Add(probe.Y.ToString("F6"));
                        dataFields.Add(probe.Z.ToString("F6"));
                    }

                    // 写入一行数据
                    _csvWriter.WriteLine(string.Join(",", dataFields));

                    // 每100行刷新一次，平衡性能和数据安全
                    if (DateTime.Now.Second % 10 == 0) // 每10秒刷新一次
                    {
                        _csvWriter.Flush();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"记录数据失败: {ex.Message}");
                }
            }
        }

        public void StopRecording()
        {
            _isRecording = false;

            lock (_lockObject)
            {
                try
                {
                    _csvWriter?.Flush();
                    _csvWriter?.Close();
                    _csvWriter?.Dispose();
                    _csvWriter = null;

                    Console.WriteLine($"数据记录已停止，文件保存到: {_csvFilePath}");

                    // 显示记录统计
                    if (File.Exists(_csvFilePath))
                    {
                        var fileInfo = new FileInfo(_csvFilePath);
                        var lineCount = File.ReadLines(_csvFilePath).Count();
                        Console.WriteLine($"文件大小: {fileInfo.Length / 1024.0:F2} KB, 总行数: {lineCount}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"停止记录时发生错误: {ex.Message}");
                }
            }
        }

        public bool IsRecording => _isRecording;
    }
}
