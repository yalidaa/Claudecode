using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MagneticField
{
    internal class MomentCalculator
    {

        // 常量定义
        private const int MAGNETOMETER_ROWS = 4;
        private const int MAGNETOMETER_COLS = 3;
        private const int MAG_FIELD_ROWS = 37;  // 36+1行（包含本底）
        private const int MAG_FIELD_COLS = 12;
        private const int MAX_ROWS = 36;
        private const int SINGLE_COLS = 3;
        private const double PI = Math.PI;
        private const double MU0_NT = 4 * PI * 1e-7 * 1e9;  // 真空磁导率，单位nT

        public static double[,] CalculateMoment(
            double[,] magnetometer,
            int angleStep,
            double[,] magField,
            int moveLine,
            int jiance)
        {
            // 检查数组维度
            if (magnetometer.GetLength(0) != MAGNETOMETER_ROWS || magnetometer.GetLength(1) != MAGNETOMETER_COLS)
                throw new ArgumentException("磁强计数组维度不正确");

            if (magField.GetLength(0) != MAG_FIELD_ROWS || magField.GetLength(1) != MAG_FIELD_COLS)
                throw new ArgumentException("磁场数组维度不正确");

            if (angleStep != 10 && angleStep != 20)
                throw new ArgumentException("angleStep只能为10或20");

            // 调用主计算函数
            double[,] magMoment = new double[9, 4];
            MagMomentChidao2(magnetometer, angleStep, magField, magMoment, moveLine, jiance);

            return magMoment;
        }

        private static void MagMomentChidao2(double[,] magnetometer, int angleStep, double[,] magField,
            double[,] magMoment, int moveLine, int jiance)
        {
            // 提取磁强计位置
            double r1 = magnetometer[0, 1];
            double r2 = magnetometer[1, 1];
            double r3 = magnetometer[2, 1];

            // 确定磁场值行数
            int rowNumb = (angleStep == 10) ? 36 : 18;

            // 提取磁场值
            double[][,] B1 = new double[36][,];
            double[][,] B2 = new double[36][,];
            double[][,] B3 = new double[36][,];
            FieldExtract(magField, rowNumb, jiance, B1, B2, B3, moveLine);

            // 一次偶极化
            double[][,] B1f = new double[36][,];
            double[][,] B2f = new double[36][,];
            double[][,] B3f = new double[36][,];
            First(B1, B1f, rowNumb);
            First(B2, B2f, rowNumb);
            First(B3, B3f, rowNumb);

            // 计算一次偶极化后的磁矩
            double[] m1 = new double[4];
            double[] m2 = new double[4];
            double[] m3 = new double[4];
            MagMomentCal(r1, angleStep, B1f, m1, rowNumb);
            MagMomentCal(r2, angleStep, B2f, m2, rowNumb);
            MagMomentCal(r3, angleStep, B3f, m3, rowNumb);

            // 二次偶极化
            double[][,] B12s = new double[36][,];
            double[][,] B23s = new double[36][,];
            double[][,] B13s = new double[36][,];
            Second(r1, r2, B1f, B2f, B12s, rowNumb);
            Second(r2, r3, B2f, B3f, B23s, rowNumb);
            Second(r1, r3, B1f, B3f, B13s, rowNumb);

            // 计算二次偶极化后的磁矩
            double[] m12 = new double[4];
            double[] m23 = new double[4];
            double[] m13 = new double[4];
            MagMomentCal(r1, angleStep, B12s, m12, rowNumb);
            MagMomentCal(r2, angleStep, B23s, m23, rowNumb);
            MagMomentCal(r1, angleStep, B13s, m13, rowNumb);

            // 三次偶极化
            double[][,] B23t = new double[36][,];
            Third(r2, r3, B12s, B13s, B23t, rowNumb);

            // 计算三次偶极化后的磁矩
            double[] m23t = new double[4];
            MagMomentCal(r1, angleStep, B23t, m23t, rowNumb);

            // 存储结果
            for (int i = 0; i < 4; i++)
            {
                magMoment[0, i] = m1[i];
                magMoment[1, i] = m2[i];
                magMoment[2, i] = m3[i];

                magMoment[3, i] = (m1[i] + m2[i] + m3[i]) / 3.0;

                magMoment[4, i] = m12[i];
                magMoment[5, i] = m23[i];
                magMoment[6, i] = m13[i];

                magMoment[7, i] = (m12[i] + m23[i] + m13[i]) / 3.0;

                magMoment[8, i] = m23t[i];
            }
        }

        private static void FieldExtract(double[,] magField, int rowNumb, int jiance,
            double[][,] B1, double[][,] B2, double[][,] B3, int moveLine)
        {
            int rowA = magField.GetLength(0);
            int colA = magField.GetLength(1);

            double[,] AField = new double[rowA, colA];

            // 复制输入矩阵
            for (int i = 0; i < rowA; i++)
            {
                for (int j = 0; j < colA; j++)
                {
                    AField[i, j] = magField[i, j];
                }
            }

            // 监测探头降噪
            if (jiance == 1 && colA == 12)
            {
                for (int i = 0; i < rowA; i++)
                {
                    for (int j = 0; j < 3; j++)
                    {
                        AField[i, j] -= AField[i, 9 + j];     // 1号减4号
                        AField[i, 3 + j] -= AField[i, 9 + j]; // 2号减4号
                        AField[i, 6 + j] -= AField[i, 9 + j]; // 3号减4号
                    }
                }
            }

            // 去本底
            if (moveLine == 1 && rowA == (rowNumb + 1))
            {
                for (int i = 0; i < rowNumb; i++)
                {
                    for (int j = 0; j < colA; j++)
                    {
                        AField[i, j] -= AField[rowA - 1, j];
                    }
                }
            }

            // 输出结果
            for (int i = 0; i < rowNumb; i++)
            {
                B1[i] = new double[1, 3];
                B2[i] = new double[1, 3];
                B3[i] = new double[1, 3];

                for (int j = 0; j < 3; j++)
                {
                    B1[i][0, j] = AField[i, j];
                    B2[i][0, j] = AField[i, 3 + j];
                    B3[i][0, j] = AField[i, 6 + j];
                }
            }
        }

        private static void First(double[][,] magB, double[][,] BFirst, int rowNumb)
        {
            // 扩展磁场数据
            double[][,] magB2 = new double[72][,];

            for (int i = 0; i < 72; i++)
            {
                int srcIndex = i % rowNumb;
                magB2[i] = new double[1, 3];
                for (int j = 0; j < 3; j++)
                {
                    magB2[i][0, j] = magB[srcIndex][0, j];
                }
            }

            // 计算偶极化结果
            for (int i = 0; i < rowNumb; i++)
            {
                BFirst[i] = new double[1, 3];
                // X向: 0.5*(magB2[i] - magB2[i+18])
                BFirst[i][0, 0] = 0.5 * (magB2[i][0, 0] - magB2[i + 18][0, 0]);
                // Y向: 0.5*(magB2[i] - magB2[i+18])
                BFirst[i][0, 1] = 0.5 * (magB2[i][0, 1] - magB2[i + 18][0, 1]);
                // Z向: 0.5*(magB2[i] + magB2[i+18])
                BFirst[i][0, 2] = 0.5 * (magB2[i][0, 2] + magB2[i + 18][0, 2]);
            }
        }

        private static void Second(double r1, double r2, double[][,] b1, double[][,] b2,
            double[][,] BSecond, int rowNumb)
        {
            double ratio = Math.Pow(r2 / r1, 5);
            double denom = Math.Pow(r2 / r1, 2) - 1;

            for (int i = 0; i < rowNumb; i++)
            {
                BSecond[i] = new double[1, 3];
                BSecond[i][0, 0] = (ratio * b2[i][0, 0] - b1[i][0, 0]) / denom;
                BSecond[i][0, 1] = (ratio * b2[i][0, 1] - b1[i][0, 1]) / denom;
                BSecond[i][0, 2] = (ratio * b2[i][0, 2] - b1[i][0, 2]) / denom;
            }
        }

        private static void Third(double r2, double r3, double[][,] b2, double[][,] b3,
            double[][,] BThird, int rowNumb)
        {
            double ratio = Math.Pow(r3 / r2, 2);
            double denom = ratio - 1;

            for (int i = 0; i < rowNumb; i++)
            {
                BThird[i] = new double[1, 3];
                BThird[i][0, 0] = (ratio * b3[i][0, 0] - b2[i][0, 0]) / denom;
                BThird[i][0, 1] = (ratio * b3[i][0, 1] - b2[i][0, 1]) / denom;
                BThird[i][0, 2] = (ratio * b3[i][0, 2] - b2[i][0, 2]) / denom;
            }
        }

        private static void MagMomentCal(double r1, int angleStep, double[][,] magB,
            double[] memont, int rowNumb)
        {
            // 设置角度参数
            double[] theta = new double[rowNumb];
            int a = rowNumb;

            double angleInRadians = angleStep * PI / 180.0;
            for (int i = 0; i < rowNumb; i++)
            {
                theta[i] = i * angleInRadians;
            }

            // 计算求和部分
            double mx = 0.0, my = 0.0, mz = 0.0;

            for (int i = 0; i < a; i++)
            {
                mx += magB[i][0, 0] * Math.Cos(theta[i]) + 0.5 * magB[i][0, 1] * Math.Sin(theta[i]);
                my += -magB[i][0, 0] * Math.Sin(theta[i]) + 0.5 * magB[i][0, 1] * Math.Cos(theta[i]);
                mz += magB[i][0, 2];
            }

            // 计算最终磁矩
            double factor = (4 * PI) * Math.Pow(r1, 3) / (MU0_NT * a);
            mx = -factor * mx;
            my = factor * my;
            mz = -factor * mz;

            // 存储结果
            memont[0] = mx;
            memont[1] = my;
            memont[2] = mz;
            memont[3] = Math.Sqrt(mx * mx + my * my + mz * mz);
        }



        //新增XYZ转换方法
        public static double[,] TransformMagneticMoments(double[,] mag_moment, string comboBoxX, string comboBoxY, string comboBoxZ)
        {
            if (mag_moment.GetLength(0) != 9 || mag_moment.GetLength(1) != 4)
            {
                throw new ArgumentException("输入数组必须是9行4列的二维数组");
            }

            double[,] new_moment = new double[9, 4];

            // 将第4列复制
            for (int i = 0; i < 9; i++)
            {
                new_moment[i, 3] = mag_moment[i, 3];
            }

            // 解析comboBox的值
            var xConfig = ParseComboBoxValue(comboBoxX);
            var yConfig = ParseComboBoxValue(comboBoxY);
            var zConfig = ParseComboBoxValue(comboBoxZ);

            // 填充新的数组
            for (int i = 0; i < 9; i++)
            {
                new_moment[i, 0] = xConfig.sign * mag_moment[i, xConfig.columnIndex];
                new_moment[i, 1] = yConfig.sign * mag_moment[i, yConfig.columnIndex];
                new_moment[i, 2] = zConfig.sign * mag_moment[i, zConfig.columnIndex];
            }

            return new_moment;
        }

        // 存储列索引和符号
        private struct ColumnConfig
        {
            public int columnIndex;
            public double sign;
        }

        // 解析comboBox的值，返回列索引和符号
        private static ColumnConfig ParseComboBoxValue(string comboBoxValue)
        {
            ColumnConfig config = new ColumnConfig();
            string value = comboBoxValue.ToUpper();

            if (value == "X")
            {
                config.columnIndex = 0;
                config.sign = 1.0;
            }
            else if (value == "Y")
            {
                config.columnIndex = 1;
                config.sign = 1.0;
            }
            else if (value == "Z")
            {
                config.columnIndex = 2;
                config.sign = 1.0;
            }
            else if (value == "-X")
            {
                config.columnIndex = 0;
                config.sign = -1.0;
            }
            else if (value == "-Y")
            {
                config.columnIndex = 1;
                config.sign = -1.0;
            }
            else if (value == "-Z")
            {
                config.columnIndex = 2;
                config.sign = -1.0;
            }
            else
            {
                throw new ArgumentException($"无效的comboBox值: {comboBoxValue}");
            }

            return config;
        }


    }
}
