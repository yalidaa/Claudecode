using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml;

namespace MagneticField
{
    internal class ExcelProbeHelper
    {
        private string _filePath;

        public ExcelProbeHelper(string filePath)
        {
            _filePath = filePath;
        }

        /// <summary>
        /// 读取Sheet1数据到DataTable
        /// </summary>
        public DataTable ReadSheet1Data(out int channelCount)
        {
            channelCount = 3;
            var dataTable = new DataTable();

            try
            {
                if (!File.Exists(_filePath))
                {
                    MessageBox.Show("Excel 文件不存在！");
                    return dataTable;
                }

                using (var package = new ExcelPackage(new FileInfo(_filePath)))
                {
                    var worksheet = package.Workbook.Worksheets["数据"] ??
                           package.Workbook.Worksheets[0];

                    if (worksheet.Dimension == null)
                    {
                        MessageBox.Show("工作表为空！");
                        return dataTable;
                    }

                    int rowCount = worksheet.Dimension.Rows;
                    int colCount = worksheet.Dimension.Columns;

                    if (colCount > 24)
                    {
                        // channel数量 = (总列数 - 1) / 6 （减去A列，每6列一个channel）
                        channelCount = (colCount - 1) / 6;
                    }

                    // 从第7行开始读取数据
                    int dataStartRow = 7;

                    // 获取通道标题（第5行）
                    List<string> channelHeaders = GetChannelHeaders(worksheet, colCount);

                    if (channelHeaders.Count == 0)
                    {
                        MessageBox.Show("未找到通道数据列");
                        return dataTable;
                    }

                    // 创建DataTable列
                    CreateDataTableColumns(ref dataTable, channelHeaders);

                    int rowsAdded = 0;
                    // 提取数据并添加到DataTable
                    for (int row = dataStartRow; row <= rowCount; row++)
                    {
                        // 第一列是number（从1开始的序号）
                        int sampleNumber = row - dataStartRow + 1;

                        DataRow newRow = dataTable.NewRow();
                        newRow["number"] = sampleNumber;

                        // 读取每个通道的数据（B列、D列、F列...）
                        for (int i = 0; i < channelHeaders.Count; i++)
                        {
                            // 数据列位置：B列(2)、D列(4)、F列(6)...
                            int dataCol = 2 + (i * 2);

                            if (dataCol > colCount) break;

                            var dataCell = worksheet.Cells[row, dataCol].Text?.Trim();
                            if (!string.IsNullOrEmpty(dataCell))
                            {
                                string numberStr = dataCell.Replace(",", ".");
                                if (double.TryParse(numberStr, System.Globalization.NumberStyles.Any,
                                    System.Globalization.CultureInfo.InvariantCulture, out double value))
                                {
                                    newRow[channelHeaders[i]] = value;
                                }
                            }
                        }

                        dataTable.Rows.Add(newRow);
                        rowsAdded++;
                    }

                    MessageBox.Show($"成功读取 {rowsAdded} 行数据");

                    // 调试：显示DataTable结构
                    //DebugDataTableStructure(dataTable);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"读取Sheet1数据时出错: {ex.Message}");
            }

            Debug.WriteLine(dataTable.Rows.Count);

            return dataTable;
        }

        public Dictionary<string, string> ReadSheet2Data()
        {
            var data = new Dictionary<string, string>();

            try
            {
                if (!File.Exists(_filePath))
                {
                    MessageBox.Show("Excel 文件不存在！");
                    return data;
                }

                using (var package = new ExcelPackage(new FileInfo(_filePath)))
                {
                    // 检查Sheet2是否存在
                    if (package.Workbook.Worksheets.Count < 2)
                        return data;

                    var worksheet = package.Workbook.Worksheets["零磁场磁矩参数"] ??
                           package.Workbook.Worksheets[1];

                    if (worksheet.Dimension == null)
                    {
                        MessageBox.Show("工作表为空！");
                        return data;
                    }


                    // 从第2行开始读取到末尾
                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        string paramName = worksheet.Cells[row, 2].Text?.Trim(); // B列：参数名称
                        string paramValue = worksheet.Cells[row, 3].Text?.Trim(); // C列：参数值

                        if (!string.IsNullOrEmpty(paramName) && !string.IsNullOrEmpty(paramValue))
                        {
                            data[paramName] = paramValue;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"读取Sheet2数据时出错: {ex.Message}");
            }

            return data;
        }


        private List<string> GetChannelHeaders(ExcelWorksheet worksheet, int totalColumns)
        {
            var headers = new List<string>();
            int headerRow = 5;

            // 从B列开始，每隔一列读取通道标题
            for (int col = 2; col <= totalColumns; col += 2)
            {
                string headerText = worksheet.Cells[headerRow, col].Text?.Trim();
                if (!string.IsNullOrEmpty(headerText) && headerText.StartsWith("channel"))
                {
                    // 去掉"channel "前缀，只保留数字
                    string channelNumber = headerText.Replace("channel", "").Trim();
                    headers.Add(channelNumber);
                }
                else
                {
                    break; // 遇到非通道标题就停止
                }
            }

            return headers;
        }
        private void CreateDataTableColumns(ref DataTable dataTable, List<string> channelHeaders)
        {
            // 第一列：number
            if (!dataTable.Columns.Contains("number"))
            {
                dataTable.Columns.Add("number", typeof(int));
            }

            // 后续列：通道数据
            foreach (string channel in channelHeaders)
            {
                if (!dataTable.Columns.Contains(channel))
                {
                    dataTable.Columns.Add(channel, typeof(double));
                }
            }
        }

        /// <summary>
        /// 调试方法：显示DataTable结构
        /// </summary>
        private void DebugDataTableStructure(DataTable dataTable)
        {
            StringBuilder debugInfo = new StringBuilder();
            debugInfo.AppendLine("DataTable结构：");
            debugInfo.AppendLine($"总行数：{dataTable.Rows.Count}");
            debugInfo.AppendLine($"总列数：{dataTable.Columns.Count}");

            if (dataTable.Columns.Count > 0)
            {
                debugInfo.Append("列顺序：");
                foreach (DataColumn col in dataTable.Columns)
                {
                    debugInfo.Append($"{col.ColumnName} ");
                }
                debugInfo.AppendLine();
            }

            if (dataTable.Rows.Count > 0)
            {
                debugInfo.AppendLine("前3行数据示例：");
                for (int i = 0; i < Math.Min(3, dataTable.Rows.Count); i++)
                {
                    debugInfo.Append($"行{i}: ");
                    for (int j = 0; j < Math.Min(5, dataTable.Columns.Count); j++)
                    {
                        debugInfo.Append($"{dataTable.Columns[j].ColumnName}={dataTable.Rows[i][j]} ");
                    }
                    debugInfo.AppendLine();
                }
            }

            MessageBox.Show(debugInfo.ToString());
        }




        public bool WriteExcelData(Dictionary<string, string> dataToWrite)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(_filePath)))
                {
                    // 检查Sheet2是否存在，不存在则创建
                    var worksheet = package.Workbook.Worksheets["零磁场磁矩参数"];
                    if (worksheet == null)
                    {
                        worksheet = package.Workbook.Worksheets.Add("零磁场磁矩参数");

                        // 创建表头
                        worksheet.Cells[1, 1].Value = "";
                        worksheet.Cells[1, 2].Value = "参数名称";
                        worksheet.Cells[1, 3].Value = "参数值";

                        // 填充数据，包含序号
                        int row = 2;
                        int index = 0;
                        foreach (var item in dataToWrite)
                        {
                            worksheet.Cells[row, 1].Value = index;      // 序号0,1,2...
                            worksheet.Cells[row, 2].Value = item.Key;   // 参数名称
                            worksheet.Cells[row, 3].Value = item.Value; // 参数值
                            row++;
                            index++;
                        }
                    }
                    else
                    {
                        // Sheet2已存在，更新数据
                        for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                        {
                            string paramName = worksheet.Cells[row, 2].Text;
                            if (dataToWrite.ContainsKey(paramName))
                            {
                                worksheet.Cells[row, 3].Value = dataToWrite[paramName];
                            }
                        }
                    }

                    package.Save();
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"写入Excel数据时出错: {ex.Message}");
                return false;
            }
        }


    }




    internal class ExcelHelper
    {
        private string _filePath;

        public ExcelHelper(string filePath)
        {
            _filePath = filePath;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public Dictionary<string, string> ReadExcelData()
        {
            var data = new Dictionary<string, string>();

            try
            {
                if (!File.Exists(_filePath))
                {
                    MessageBox.Show("Excel 文件不存在！");
                    return data;
                }

                using (var package = new ExcelPackage(new FileInfo(_filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[1];

                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        string paramName = worksheet.Cells[row, 2].Text;
                        string paramValue = worksheet.Cells[row, 3].Text;
                        data[paramName] = paramValue;
                    }
                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"读取Excel数据时出错: {ex.Message}");
            }

            return data;
        }

        public bool WriteExcelData(Dictionary<string, string> dataToWrite)
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(_filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[0];

                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        string paramName = worksheet.Cells[row, 2].Text;
                        if (dataToWrite.ContainsKey(paramName))
                        {
                            worksheet.Cells[row, 3].Value = dataToWrite[paramName];
                        }
                    }

                    package.Save();
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"写入Excel数据时出错: {ex.Message}");
                return false;
            }
        }

        // 坐标转换辅助方法
        public static int GetCoordinateIndex(string value)
        {
            string[] coordinateOptions = { "X", "Y", "Z", "-X", "-Y", "-Z" };
            int index = int.Parse(value);
            return index >= 0 && index < coordinateOptions.Length ? index : 0;
        }

        public static string GetCoordinateValue(string selectedText)
        {
            switch (selectedText)
            {
                case "X": return "0";
                case "Y": return "1";
                case "Z": return "2";
                case "-X": return "3";
                case "-Y": return "4";
                case "-Z": return "5";
                default: return "1";
            }
        }
    }
}
