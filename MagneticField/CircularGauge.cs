using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MagneticField
{
    public partial class CircularGauge : UserControl
    {
        private float _currentAngle = 0;
        private readonly Font _scaleFont = new Font("Arial", 8);
        private readonly Font _centerFont = new Font("Arial", 12, FontStyle.Bold);

        public CircularGauge()
        {
            InitializeComponent();

            this.DoubleBuffered = true; // 避免闪烁
            this.Size = new Size(300, 300);
        }

        // 当前角度属性（0-359.9）
        public float CurrentAngle
        {
            get => _currentAngle;
            set
            {
                _currentAngle = value % 360;
                this.Invalidate(); // 触发重绘
            }
        }

protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        Graphics g = e.Graphics;
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

        // 1. 绘制表盘外圆
        Rectangle rect = new Rectangle(20, 20, Width - 40, Height - 40);
        g.DrawEllipse(new Pen(Color.DarkGray, 3), rect);

        // 2. 绘制刻度（每5°小刻度，每20°显示数字）
        for (int angle = 0; angle < 360; angle += 5)
        {
            float mathAngle = angle + 90; // 转换坐标系（0°在正下方）
            bool isMajorTick = angle % 20 == 0;

            // 刻度线长度区分
            float innerRadius = rect.Width / 2 - (isMajorTick ? 25 : 20);
            PointF start = GetPointOnCircle(mathAngle, rect.Width / 2 - 10);
            PointF end = GetPointOnCircle(mathAngle, innerRadius);
            g.DrawLine(new Pen(Color.Black, isMajorTick ? 2 : 1), start, end);

            // 每20°显示数字
            if (isMajorTick)
            {
                PointF textPos = GetPointOnCircle(mathAngle, rect.Width / 2 - 35);
                string text = angle.ToString();
                SizeF textSize = g.MeasureString(text, _scaleFont);
                g.DrawString(text, _scaleFont, Brushes.Black,
                    textPos.X - textSize.Width / 2, textPos.Y - textSize.Height / 2);
            }
        }

        // 3. 绘制指针
        float needleAngle = _currentAngle + 90;
        PointF needleEnd = GetPointOnCircle(needleAngle, rect.Width / 2 - 35);
        g.DrawLine(new Pen(Color.Red, 4),
            new PointF(rect.Width / 2 + 20, rect.Height / 2 + 20),
            needleEnd);

        // 4. 中心显示当前角度
        string angleText = $"{_currentAngle:F1}°";
        SizeF angletextSize = g.MeasureString(angleText, _centerFont);
        g.DrawString(angleText, _centerFont, Brushes.Blue,
            Width / 2 - angletextSize.Width / 2,
            Height / 2 - angletextSize.Height / 2);
    }

    private PointF GetPointOnCircle(float angle, float radius)
    {
        float radian = angle * (float)Math.PI / 180;
        float x = Width / 2 + radius * (float)Math.Cos(radian);
        float y = Height / 2 + radius * (float)Math.Sin(radian);
        return new PointF(x, y);
    }

    }
}
