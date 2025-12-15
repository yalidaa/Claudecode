using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using System.Runtime.InteropServices;
using System.Windows.Forms.DataVisualization.Charting;
using System.Net;
using System.Net.Sockets;
using System.IO;
using System.Threading;
using System.Reflection;
using SkiaSharp;
using System.Media;
using System.Collections.Concurrent;
using System.Diagnostics; // 提示音
using ScottPlot;
using ScottPlot.Plottables;//formsPlot控件使用
using OfficeOpenXml;
using OpenTK.Graphics.ES11;
using DocumentFormat.OpenXml.ExtendedProperties;
using VBIDE;
using NationalInstruments.Examples.ContAcqVoltageSamples_IntClk;
using System.Timers;
using ScottPlot.WinForms;
using Color = System.Drawing.Color;
using System.Windows.Forms.VisualStyles;
using Rectangle = System.Drawing.Rectangle;
using DocumentFormat.OpenXml.Drawing.Charts;
using DataTable = System.Data.DataTable;
using Size = System.Drawing.Size;
using DocumentFormat.OpenXml.Vml.Office;
using ComboBox = System.Windows.Forms.ComboBox;
using Label = System.Windows.Forms.Label;
using TextBox = System.Windows.Forms.TextBox;
using Button = System.Windows.Forms.Button;
using Ookii.Dialogs.WinForms;
using System.Runtime.InteropServices.ComTypes;



namespace MagneticField
{
    public partial class Form1 : Form,IDataSink
    {

        

        public Form1()
        {
            InitializeComponent();

            // 禁止用户拖拽调整窗口大小
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            // 隐藏最大化按钮（但保留最小化和关闭）
            this.MaximizeBox = false;

            tabControl1.SelectedIndex = 5;
            tabControl5.SelectedIndex = 0;
            start11();
            start12();

            start21();
            start22();

            start3();

            start61();
            start62();

        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.groupBox21 = new System.Windows.Forms.GroupBox();
            this.button_11_2 = new System.Windows.Forms.Button();
            this.button_11_5 = new System.Windows.Forms.Button();
            this.button_11_4 = new System.Windows.Forms.Button();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.button_12_1 = new System.Windows.Forms.Button();
            this.button_11_1 = new System.Windows.Forms.Button();
            this.tabControl11 = new System.Windows.Forms.TabControl();
            this.tabPage11 = new System.Windows.Forms.TabPage();
            this.label43 = new System.Windows.Forms.Label();
            this.textBox_11_1 = new System.Windows.Forms.TextBox();
            this.label38 = new System.Windows.Forms.Label();
            this.tabPage12 = new System.Windows.Forms.TabPage();
            this.label37 = new System.Windows.Forms.Label();
            this.numericUpDown13 = new System.Windows.Forms.NumericUpDown();
            this.label36 = new System.Windows.Forms.Label();
            this.numericUpDown12 = new System.Windows.Forms.NumericUpDown();
            this.label35 = new System.Windows.Forms.Label();
            this.numericUpDown11 = new System.Windows.Forms.NumericUpDown();
            this.groupBox_12_2 = new System.Windows.Forms.GroupBox();
            this.button_12_004 = new System.Windows.Forms.Button();
            this.button_12_003 = new System.Windows.Forms.Button();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.pictureBox11 = new System.Windows.Forms.PictureBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.textBox_11_3 = new System.Windows.Forms.TextBox();
            this.checkedListBox2 = new System.Windows.Forms.CheckedListBox();
            this.textBox_11_2 = new System.Windows.Forms.TextBox();
            this.label39 = new System.Windows.Forms.Label();
            this.groupBox8 = new System.Windows.Forms.GroupBox();
            this.button_11_9 = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.formsPlot1 = new ScottPlot.WinForms.FormsPlot();
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.chart1 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.button_11_8 = new System.Windows.Forms.Button();
            this.button_11_7 = new System.Windows.Forms.Button();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.groupBox24 = new System.Windows.Forms.GroupBox();
            this.textBox_23_3 = new System.Windows.Forms.TextBox();
            this.textBox_23_2 = new System.Windows.Forms.TextBox();
            this.textBox_23_1 = new System.Windows.Forms.TextBox();
            this.label152 = new System.Windows.Forms.Label();
            this.label151 = new System.Windows.Forms.Label();
            this.comboBox_23_1 = new System.Windows.Forms.ComboBox();
            this.groupBox18 = new System.Windows.Forms.GroupBox();
            this.tabControl23 = new System.Windows.Forms.TabControl();
            this.tabPage231 = new System.Windows.Forms.TabPage();
            this.tabControl4 = new System.Windows.Forms.TabControl();
            this.tabPage10 = new System.Windows.Forms.TabPage();
            this.textBox_22_3 = new System.Windows.Forms.TextBox();
            this.textBox_22_4 = new System.Windows.Forms.TextBox();
            this.label140 = new System.Windows.Forms.Label();
            this.textBox_22_1 = new System.Windows.Forms.TextBox();
            this.label138 = new System.Windows.Forms.Label();
            this.label141 = new System.Windows.Forms.Label();
            this.tabPage15 = new System.Windows.Forms.TabPage();
            this.groupBox20 = new System.Windows.Forms.GroupBox();
            this.radioButton_22_4 = new System.Windows.Forms.RadioButton();
            this.radioButton_22_3 = new System.Windows.Forms.RadioButton();
            this.textBox_22_2 = new System.Windows.Forms.TextBox();
            this.label139 = new System.Windows.Forms.Label();
            this.groupBox19 = new System.Windows.Forms.GroupBox();
            this.radioButton_22_2 = new System.Windows.Forms.RadioButton();
            this.radioButton_22_1 = new System.Windows.Forms.RadioButton();
            this.tabPage232 = new System.Windows.Forms.TabPage();
            this.label142 = new System.Windows.Forms.Label();
            this.textBox_22_5 = new System.Windows.Forms.TextBox();
            this.label143 = new System.Windows.Forms.Label();
            this.textBox_22_6 = new System.Windows.Forms.TextBox();
            this.label144 = new System.Windows.Forms.Label();
            this.textBox_22_7 = new System.Windows.Forms.TextBox();
            this.comboBox_22_1 = new System.Windows.Forms.ComboBox();
            this.formsPlot3 = new ScottPlot.WinForms.FormsPlot();
            this.label136 = new System.Windows.Forms.Label();
            this.button_22_2 = new System.Windows.Forms.Button();
            this.button_22_1 = new System.Windows.Forms.Button();
            this.groupBox17 = new System.Windows.Forms.GroupBox();
            this.formsPlot2 = new ScottPlot.WinForms.FormsPlot();
            this.comboBox_21_7 = new System.Windows.Forms.ComboBox();
            this.button_21_2 = new System.Windows.Forms.Button();
            this.label127 = new System.Windows.Forms.Label();
            this.label133 = new System.Windows.Forms.Label();
            this.label128 = new System.Windows.Forms.Label();
            this.button_21_1 = new System.Windows.Forms.Button();
            this.comboBox_21_8 = new System.Windows.Forms.ComboBox();
            this.label132 = new System.Windows.Forms.Label();
            this.label129 = new System.Windows.Forms.Label();
            this.label131 = new System.Windows.Forms.Label();
            this.comboBox_21_9 = new System.Windows.Forms.ComboBox();
            this.textBox_21_8_3 = new System.Windows.Forms.TextBox();
            this.label130 = new System.Windows.Forms.Label();
            this.textBox_21_8_2 = new System.Windows.Forms.TextBox();
            this.textBox_21_7_1 = new System.Windows.Forms.TextBox();
            this.textBox_21_8_1 = new System.Windows.Forms.TextBox();
            this.textBox_21_7_2 = new System.Windows.Forms.TextBox();
            this.textBox_21_7_3 = new System.Windows.Forms.TextBox();
            this.groupBox16 = new System.Windows.Forms.GroupBox();
            this.textBox_21_4_1 = new System.Windows.Forms.TextBox();
            this.comboBox_21_1 = new System.Windows.Forms.ComboBox();
            this.label119 = new System.Windows.Forms.Label();
            this.comboBox_21_2 = new System.Windows.Forms.ComboBox();
            this.label120 = new System.Windows.Forms.Label();
            this.comboBox_21_3 = new System.Windows.Forms.ComboBox();
            this.label121 = new System.Windows.Forms.Label();
            this.comboBox_21_4 = new System.Windows.Forms.ComboBox();
            this.label122 = new System.Windows.Forms.Label();
            this.comboBox_21_5 = new System.Windows.Forms.ComboBox();
            this.label123 = new System.Windows.Forms.Label();
            this.comboBox_21_6 = new System.Windows.Forms.ComboBox();
            this.label124 = new System.Windows.Forms.Label();
            this.label125 = new System.Windows.Forms.Label();
            this.textBox_21_1_1 = new System.Windows.Forms.TextBox();
            this.textBox_21_2_1 = new System.Windows.Forms.TextBox();
            this.textBox_21_3_1 = new System.Windows.Forms.TextBox();
            this.textBox_21_5_1 = new System.Windows.Forms.TextBox();
            this.textBox_21_6_1 = new System.Windows.Forms.TextBox();
            this.textBox_21_1_2 = new System.Windows.Forms.TextBox();
            this.label126 = new System.Windows.Forms.Label();
            this.textBox_21_2_2 = new System.Windows.Forms.TextBox();
            this.textBox_21_6_2 = new System.Windows.Forms.TextBox();
            this.textBox_21_3_2 = new System.Windows.Forms.TextBox();
            this.textBox_21_5_2 = new System.Windows.Forms.TextBox();
            this.textBox_21_4_2 = new System.Windows.Forms.TextBox();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.tabControl2 = new System.Windows.Forms.TabControl();
            this.tabPage7 = new System.Windows.Forms.TabPage();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.label29 = new System.Windows.Forms.Label();
            this.label30 = new System.Windows.Forms.Label();
            this.label31 = new System.Windows.Forms.Label();
            this.label32 = new System.Windows.Forms.Label();
            this.label33 = new System.Windows.Forms.Label();
            this.textBox_31_33 = new System.Windows.Forms.TextBox();
            this.textBox_31_34 = new System.Windows.Forms.TextBox();
            this.textBox_31_35 = new System.Windows.Forms.TextBox();
            this.textBox_31_36 = new System.Windows.Forms.TextBox();
            this.label34 = new System.Windows.Forms.Label();
            this.textBox_31_32 = new System.Windows.Forms.TextBox();
            this.label21 = new System.Windows.Forms.Label();
            this.label22 = new System.Windows.Forms.Label();
            this.label23 = new System.Windows.Forms.Label();
            this.label24 = new System.Windows.Forms.Label();
            this.label25 = new System.Windows.Forms.Label();
            this.label26 = new System.Windows.Forms.Label();
            this.label27 = new System.Windows.Forms.Label();
            this.textBox_31_17 = new System.Windows.Forms.TextBox();
            this.textBox_31_18 = new System.Windows.Forms.TextBox();
            this.textBox_31_19 = new System.Windows.Forms.TextBox();
            this.textBox_31_20 = new System.Windows.Forms.TextBox();
            this.textBox_31_21 = new System.Windows.Forms.TextBox();
            this.textBox_31_22 = new System.Windows.Forms.TextBox();
            this.textBox_31_23 = new System.Windows.Forms.TextBox();
            this.textBox_31_24 = new System.Windows.Forms.TextBox();
            this.textBox_31_25 = new System.Windows.Forms.TextBox();
            this.textBox_31_26 = new System.Windows.Forms.TextBox();
            this.textBox_31_27 = new System.Windows.Forms.TextBox();
            this.textBox_31_28 = new System.Windows.Forms.TextBox();
            this.textBox_31_29 = new System.Windows.Forms.TextBox();
            this.textBox_31_30 = new System.Windows.Forms.TextBox();
            this.textBox_31_31 = new System.Windows.Forms.TextBox();
            this.label28 = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.label19 = new System.Windows.Forms.Label();
            this.label18 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.textBox_31_16 = new System.Windows.Forms.TextBox();
            this.textBox_31_15 = new System.Windows.Forms.TextBox();
            this.textBox_31_14 = new System.Windows.Forms.TextBox();
            this.textBox_31_13 = new System.Windows.Forms.TextBox();
            this.textBox_31_12 = new System.Windows.Forms.TextBox();
            this.textBox_31_11 = new System.Windows.Forms.TextBox();
            this.textBox_31_10 = new System.Windows.Forms.TextBox();
            this.textBox_31_9 = new System.Windows.Forms.TextBox();
            this.textBox_31_5 = new System.Windows.Forms.TextBox();
            this.textBox_31_6 = new System.Windows.Forms.TextBox();
            this.textBox_31_7 = new System.Windows.Forms.TextBox();
            this.textBox_31_8 = new System.Windows.Forms.TextBox();
            this.textBox_31_3 = new System.Windows.Forms.TextBox();
            this.textBox_31_4 = new System.Windows.Forms.TextBox();
            this.textBox_31_2 = new System.Windows.Forms.TextBox();
            this.textBox_31_1 = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.button_31_3 = new System.Windows.Forms.Button();
            this.label12 = new System.Windows.Forms.Label();
            this.comboBox_31_7 = new System.Windows.Forms.ComboBox();
            this.label11 = new System.Windows.Forms.Label();
            this.comboBox_31_6 = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.comboBox_31_5 = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.checkBox_31_1 = new System.Windows.Forms.CheckBox();
            this.radioButton_31_2 = new System.Windows.Forms.RadioButton();
            this.label8 = new System.Windows.Forms.Label();
            this.radioButton_31_1 = new System.Windows.Forms.RadioButton();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label7 = new System.Windows.Forms.Label();
            this.button_31_2 = new System.Windows.Forms.Button();
            this.textBox_31_r3 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.textBox_31_r2 = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.textBox_31_r1 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.comboBox_31_4 = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.comboBox_31_3 = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.comboBox_31_2 = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.comboBox_31_1 = new System.Windows.Forms.ComboBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.textBox_30_1 = new System.Windows.Forms.TextBox();
            this.comboBox_30_1 = new System.Windows.Forms.ComboBox();
            this.radioButton_30_2 = new System.Windows.Forms.RadioButton();
            this.radioButton_30_1 = new System.Windows.Forms.RadioButton();
            this.button_31_1 = new System.Windows.Forms.Button();
            this.tabPage8 = new System.Windows.Forms.TabPage();
            this.groupBox15 = new System.Windows.Forms.GroupBox();
            this.label99 = new System.Windows.Forms.Label();
            this.textBox40 = new System.Windows.Forms.TextBox();
            this.label100 = new System.Windows.Forms.Label();
            this.textBox41 = new System.Windows.Forms.TextBox();
            this.label101 = new System.Windows.Forms.Label();
            this.textBox42 = new System.Windows.Forms.TextBox();
            this.groupBox14 = new System.Windows.Forms.GroupBox();
            this.radioButton_32_9 = new System.Windows.Forms.RadioButton();
            this.radioButton_32_8 = new System.Windows.Forms.RadioButton();
            this.label96 = new System.Windows.Forms.Label();
            this.textBox37 = new System.Windows.Forms.TextBox();
            this.label97 = new System.Windows.Forms.Label();
            this.textBox38 = new System.Windows.Forms.TextBox();
            this.label98 = new System.Windows.Forms.Label();
            this.textBox39 = new System.Windows.Forms.TextBox();
            this.label95 = new System.Windows.Forms.Label();
            this.label109 = new System.Windows.Forms.Label();
            this.label110 = new System.Windows.Forms.Label();
            this.label111 = new System.Windows.Forms.Label();
            this.label112 = new System.Windows.Forms.Label();
            this.label113 = new System.Windows.Forms.Label();
            this.label114 = new System.Windows.Forms.Label();
            this.label115 = new System.Windows.Forms.Label();
            this.textBox61 = new System.Windows.Forms.TextBox();
            this.textBox62 = new System.Windows.Forms.TextBox();
            this.textBox63 = new System.Windows.Forms.TextBox();
            this.textBox64 = new System.Windows.Forms.TextBox();
            this.textBox65 = new System.Windows.Forms.TextBox();
            this.textBox66 = new System.Windows.Forms.TextBox();
            this.textBox67 = new System.Windows.Forms.TextBox();
            this.textBox68 = new System.Windows.Forms.TextBox();
            this.textBox69 = new System.Windows.Forms.TextBox();
            this.textBox70 = new System.Windows.Forms.TextBox();
            this.textBox71 = new System.Windows.Forms.TextBox();
            this.textBox72 = new System.Windows.Forms.TextBox();
            this.groupBox13 = new System.Windows.Forms.GroupBox();
            this.label73 = new System.Windows.Forms.Label();
            this.label74 = new System.Windows.Forms.Label();
            this.label75 = new System.Windows.Forms.Label();
            this.label76 = new System.Windows.Forms.Label();
            this.label77 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.label78 = new System.Windows.Forms.Label();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.label79 = new System.Windows.Forms.Label();
            this.label80 = new System.Windows.Forms.Label();
            this.label81 = new System.Windows.Forms.Label();
            this.label82 = new System.Windows.Forms.Label();
            this.label83 = new System.Windows.Forms.Label();
            this.label84 = new System.Windows.Forms.Label();
            this.label85 = new System.Windows.Forms.Label();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.textBox7 = new System.Windows.Forms.TextBox();
            this.textBox8 = new System.Windows.Forms.TextBox();
            this.textBox9 = new System.Windows.Forms.TextBox();
            this.textBox10 = new System.Windows.Forms.TextBox();
            this.textBox11 = new System.Windows.Forms.TextBox();
            this.textBox12 = new System.Windows.Forms.TextBox();
            this.textBox13 = new System.Windows.Forms.TextBox();
            this.textBox14 = new System.Windows.Forms.TextBox();
            this.textBox15 = new System.Windows.Forms.TextBox();
            this.textBox16 = new System.Windows.Forms.TextBox();
            this.textBox17 = new System.Windows.Forms.TextBox();
            this.textBox18 = new System.Windows.Forms.TextBox();
            this.textBox19 = new System.Windows.Forms.TextBox();
            this.textBox20 = new System.Windows.Forms.TextBox();
            this.label86 = new System.Windows.Forms.Label();
            this.label87 = new System.Windows.Forms.Label();
            this.label88 = new System.Windows.Forms.Label();
            this.label89 = new System.Windows.Forms.Label();
            this.label90 = new System.Windows.Forms.Label();
            this.label91 = new System.Windows.Forms.Label();
            this.label92 = new System.Windows.Forms.Label();
            this.label93 = new System.Windows.Forms.Label();
            this.textBox21 = new System.Windows.Forms.TextBox();
            this.textBox22 = new System.Windows.Forms.TextBox();
            this.textBox23 = new System.Windows.Forms.TextBox();
            this.textBox24 = new System.Windows.Forms.TextBox();
            this.textBox25 = new System.Windows.Forms.TextBox();
            this.textBox26 = new System.Windows.Forms.TextBox();
            this.textBox27 = new System.Windows.Forms.TextBox();
            this.textBox28 = new System.Windows.Forms.TextBox();
            this.textBox29 = new System.Windows.Forms.TextBox();
            this.textBox30 = new System.Windows.Forms.TextBox();
            this.textBox31 = new System.Windows.Forms.TextBox();
            this.textBox32 = new System.Windows.Forms.TextBox();
            this.textBox33 = new System.Windows.Forms.TextBox();
            this.textBox34 = new System.Windows.Forms.TextBox();
            this.textBox35 = new System.Windows.Forms.TextBox();
            this.textBox36 = new System.Windows.Forms.TextBox();
            this.label94 = new System.Windows.Forms.Label();
            this.groupBox12 = new System.Windows.Forms.GroupBox();
            this.radioButton_32_7 = new System.Windows.Forms.RadioButton();
            this.radioButton_32_6 = new System.Windows.Forms.RadioButton();
            this.radioButton_32_5 = new System.Windows.Forms.RadioButton();
            this.checkBox_32_4 = new System.Windows.Forms.CheckBox();
            this.checkBox_32_3 = new System.Windows.Forms.CheckBox();
            this.checkBox_32_2 = new System.Windows.Forms.CheckBox();
            this.button_32_5 = new System.Windows.Forms.Button();
            this.label72 = new System.Windows.Forms.Label();
            this.label71 = new System.Windows.Forms.Label();
            this.groupBox9 = new System.Windows.Forms.GroupBox();
            this.button_32_4 = new System.Windows.Forms.Button();
            this.radioButton_32_4 = new System.Windows.Forms.RadioButton();
            this.radioButton_32_3 = new System.Windows.Forms.RadioButton();
            this.label70 = new System.Windows.Forms.Label();
            this.label67 = new System.Windows.Forms.Label();
            this.comboBox_32_13 = new System.Windows.Forms.ComboBox();
            this.label68 = new System.Windows.Forms.Label();
            this.comboBox_32_12 = new System.Windows.Forms.ComboBox();
            this.label69 = new System.Windows.Forms.Label();
            this.comboBox_32_11 = new System.Windows.Forms.ComboBox();
            this.label64 = new System.Windows.Forms.Label();
            this.comboBox_32_10 = new System.Windows.Forms.ComboBox();
            this.label65 = new System.Windows.Forms.Label();
            this.comboBox_32_9 = new System.Windows.Forms.ComboBox();
            this.label66 = new System.Windows.Forms.Label();
            this.comboBox_32_8 = new System.Windows.Forms.ComboBox();
            this.label63 = new System.Windows.Forms.Label();
            this.label59 = new System.Windows.Forms.Label();
            this.label60 = new System.Windows.Forms.Label();
            this.textBox_32_z1 = new System.Windows.Forms.TextBox();
            this.label61 = new System.Windows.Forms.Label();
            this.textBox_32_y1 = new System.Windows.Forms.TextBox();
            this.label62 = new System.Windows.Forms.Label();
            this.textBox_32_x1 = new System.Windows.Forms.TextBox();
            this.groupBox10 = new System.Windows.Forms.GroupBox();
            this.radioButton_32_2 = new System.Windows.Forms.RadioButton();
            this.button_32_3 = new System.Windows.Forms.Button();
            this.radioButton_32_1 = new System.Windows.Forms.RadioButton();
            this.label49 = new System.Windows.Forms.Label();
            this.label45 = new System.Windows.Forms.Label();
            this.checkBox_32_1 = new System.Windows.Forms.CheckBox();
            this.comboBox_32_7 = new System.Windows.Forms.ComboBox();
            this.label58 = new System.Windows.Forms.Label();
            this.label46 = new System.Windows.Forms.Label();
            this.label57 = new System.Windows.Forms.Label();
            this.comboBox_32_6 = new System.Windows.Forms.ComboBox();
            this.label50 = new System.Windows.Forms.Label();
            this.label47 = new System.Windows.Forms.Label();
            this.textBox_32_r3 = new System.Windows.Forms.TextBox();
            this.comboBox_32_5 = new System.Windows.Forms.ComboBox();
            this.label51 = new System.Windows.Forms.Label();
            this.label48 = new System.Windows.Forms.Label();
            this.textBox_32_r2 = new System.Windows.Forms.TextBox();
            this.label52 = new System.Windows.Forms.Label();
            this.textBox_32_r1 = new System.Windows.Forms.TextBox();
            this.label53 = new System.Windows.Forms.Label();
            this.comboBox_32_4 = new System.Windows.Forms.ComboBox();
            this.label54 = new System.Windows.Forms.Label();
            this.comboBox_32_3 = new System.Windows.Forms.ComboBox();
            this.label55 = new System.Windows.Forms.Label();
            this.comboBox_32_2 = new System.Windows.Forms.ComboBox();
            this.label56 = new System.Windows.Forms.Label();
            this.comboBox_32_1 = new System.Windows.Forms.ComboBox();
            this.groupBox11 = new System.Windows.Forms.GroupBox();
            this.button_32_2 = new System.Windows.Forms.Button();
            this.button_32_1 = new System.Windows.Forms.Button();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.comboBox_4_1 = new System.Windows.Forms.ComboBox();
            this.label_4_1 = new System.Windows.Forms.Label();
            this.tabPage5 = new System.Windows.Forms.TabPage();
            this.tabPage6 = new System.Windows.Forms.TabPage();
            this.tabControl5 = new System.Windows.Forms.TabControl();
            this.tabPage13 = new System.Windows.Forms.TabPage();
            this.groupBox23 = new System.Windows.Forms.GroupBox();
            this.checkBox_61_2 = new System.Windows.Forms.CheckBox();
            this.label148 = new System.Windows.Forms.Label();
            this.comboBox_61_8 = new System.Windows.Forms.ComboBox();
            this.label149 = new System.Windows.Forms.Label();
            this.comboBox_61_7 = new System.Windows.Forms.ComboBox();
            this.label150 = new System.Windows.Forms.Label();
            this.comboBox_61_6 = new System.Windows.Forms.ComboBox();
            this.groupBox22 = new System.Windows.Forms.GroupBox();
            this.checkBox_61_1 = new System.Windows.Forms.CheckBox();
            this.label147 = new System.Windows.Forms.Label();
            this.comboBox_61_5 = new System.Windows.Forms.ComboBox();
            this.label146 = new System.Windows.Forms.Label();
            this.comboBox_61_4 = new System.Windows.Forms.ComboBox();
            this.label145 = new System.Windows.Forms.Label();
            this.comboBox_61_3 = new System.Windows.Forms.ComboBox();
            this.label44 = new System.Windows.Forms.Label();
            this.comboBox_61_2 = new System.Windows.Forms.ComboBox();
            this.label42 = new System.Windows.Forms.Label();
            this.label41 = new System.Windows.Forms.Label();
            this.label40 = new System.Windows.Forms.Label();
            this.comboBox_61_1 = new System.Windows.Forms.ComboBox();
            this.button_61_3 = new System.Windows.Forms.Button();
            this.button_61_2 = new System.Windows.Forms.Button();
            this.button_61_1 = new System.Windows.Forms.Button();
            this.dataGridView_61_1 = new System.Windows.Forms.DataGridView();
            this.dataGridView_61_2 = new System.Windows.Forms.DataGridView();
            this.tabPage14 = new System.Windows.Forms.TabPage();
            this.textBox_62_9_1 = new System.Windows.Forms.TextBox();
            this.label137 = new System.Windows.Forms.Label();
            this.textBox_62_8_2 = new System.Windows.Forms.TextBox();
            this.textBox_62_8_1 = new System.Windows.Forms.TextBox();
            this.label135 = new System.Windows.Forms.Label();
            this.textBox_62_7_2 = new System.Windows.Forms.TextBox();
            this.textBox_62_7_1 = new System.Windows.Forms.TextBox();
            this.label134 = new System.Windows.Forms.Label();
            this.textBox_62_6_4 = new System.Windows.Forms.TextBox();
            this.textBox_62_5_4 = new System.Windows.Forms.TextBox();
            this.textBox_62_4_4 = new System.Windows.Forms.TextBox();
            this.textBox_62_3_4 = new System.Windows.Forms.TextBox();
            this.textBox_62_2_4 = new System.Windows.Forms.TextBox();
            this.textBox_62_1_4 = new System.Windows.Forms.TextBox();
            this.textBox_62_6_3 = new System.Windows.Forms.TextBox();
            this.textBox_62_5_3 = new System.Windows.Forms.TextBox();
            this.textBox_62_4_3 = new System.Windows.Forms.TextBox();
            this.textBox_62_3_3 = new System.Windows.Forms.TextBox();
            this.textBox_62_2_3 = new System.Windows.Forms.TextBox();
            this.textBox_62_1_3 = new System.Windows.Forms.TextBox();
            this.textBox_62_6_2 = new System.Windows.Forms.TextBox();
            this.textBox_62_5_2 = new System.Windows.Forms.TextBox();
            this.textBox_62_4_2 = new System.Windows.Forms.TextBox();
            this.textBox_62_3_2 = new System.Windows.Forms.TextBox();
            this.textBox_62_2_2 = new System.Windows.Forms.TextBox();
            this.textBox_62_1_2 = new System.Windows.Forms.TextBox();
            this.textBox_62_6_1 = new System.Windows.Forms.TextBox();
            this.textBox_62_5_1 = new System.Windows.Forms.TextBox();
            this.textBox_62_4_1 = new System.Windows.Forms.TextBox();
            this.textBox_62_3_1 = new System.Windows.Forms.TextBox();
            this.textBox_62_2_1 = new System.Windows.Forms.TextBox();
            this.label118 = new System.Windows.Forms.Label();
            this.label117 = new System.Windows.Forms.Label();
            this.label116 = new System.Windows.Forms.Label();
            this.label108 = new System.Windows.Forms.Label();
            this.label107 = new System.Windows.Forms.Label();
            this.label106 = new System.Windows.Forms.Label();
            this.label105 = new System.Windows.Forms.Label();
            this.label104 = new System.Windows.Forms.Label();
            this.label103 = new System.Windows.Forms.Label();
            this.label102 = new System.Windows.Forms.Label();
            this.textBox_62_1_1 = new System.Windows.Forms.TextBox();
            this.circularGauge1 = new MagneticField.CircularGauge();
            this.circularGauge2 = new MagneticField.CircularGauge();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.groupBox21.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.tabControl11.SuspendLayout();
            this.tabPage11.SuspendLayout();
            this.tabPage12.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown13)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown12)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown11)).BeginInit();
            this.groupBox_12_2.SuspendLayout();
            this.groupBox7.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox11)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.groupBox6.SuspendLayout();
            this.groupBox8.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).BeginInit();
            this.tabPage2.SuspendLayout();
            this.groupBox24.SuspendLayout();
            this.groupBox18.SuspendLayout();
            this.tabControl23.SuspendLayout();
            this.tabPage231.SuspendLayout();
            this.tabControl4.SuspendLayout();
            this.tabPage10.SuspendLayout();
            this.tabPage15.SuspendLayout();
            this.groupBox20.SuspendLayout();
            this.groupBox19.SuspendLayout();
            this.tabPage232.SuspendLayout();
            this.groupBox17.SuspendLayout();
            this.groupBox16.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.tabControl2.SuspendLayout();
            this.tabPage7.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabPage8.SuspendLayout();
            this.groupBox15.SuspendLayout();
            this.groupBox14.SuspendLayout();
            this.groupBox13.SuspendLayout();
            this.groupBox12.SuspendLayout();
            this.groupBox9.SuspendLayout();
            this.groupBox10.SuspendLayout();
            this.groupBox11.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.tabPage6.SuspendLayout();
            this.tabControl5.SuspendLayout();
            this.tabPage13.SuspendLayout();
            this.groupBox23.SuspendLayout();
            this.groupBox22.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_61_1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_61_2)).BeginInit();
            this.tabPage14.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Alignment = System.Windows.Forms.TabAlignment.Left;
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Controls.Add(this.tabPage5);
            this.tabControl1.Controls.Add(this.tabPage6);
            this.tabControl1.ImageList = this.imageList1;
            this.tabControl1.ItemSize = new System.Drawing.Size(80, 80);
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Multiline = true;
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1200, 800);
            this.tabControl1.SizeMode = System.Windows.Forms.TabSizeMode.Fixed;
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.groupBox21);
            this.tabPage1.Controls.Add(this.groupBox5);
            this.tabPage1.Controls.Add(this.tabControl11);
            this.tabPage1.Controls.Add(this.groupBox_12_2);
            this.tabPage1.Controls.Add(this.groupBox7);
            this.tabPage1.Controls.Add(this.groupBox6);
            this.tabPage1.Controls.Add(this.groupBox8);
            this.tabPage1.ImageIndex = 5;
            this.tabPage1.Location = new System.Drawing.Point(84, 4);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1112, 792);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // groupBox21
            // 
            this.groupBox21.Controls.Add(this.button_11_2);
            this.groupBox21.Controls.Add(this.button_11_5);
            this.groupBox21.Controls.Add(this.button_11_4);
            this.groupBox21.Location = new System.Drawing.Point(303, 8);
            this.groupBox21.Name = "groupBox21";
            this.groupBox21.Size = new System.Drawing.Size(465, 65);
            this.groupBox21.TabIndex = 4;
            this.groupBox21.TabStop = false;
            this.groupBox21.Text = "记录探头采集Excel";
            // 
            // button_11_2
            // 
            this.button_11_2.Location = new System.Drawing.Point(40, 20);
            this.button_11_2.Name = "button_11_2";
            this.button_11_2.Size = new System.Drawing.Size(100, 30);
            this.button_11_2.TabIndex = 1;
            this.button_11_2.Text = "选择文件夹";
            this.button_11_2.UseVisualStyleBackColor = true;
            this.button_11_2.Click += new System.EventHandler(this.button_11_2_Click);
            // 
            // button_11_5
            // 
            this.button_11_5.Location = new System.Drawing.Point(330, 20);
            this.button_11_5.Name = "button_11_5";
            this.button_11_5.Size = new System.Drawing.Size(100, 30);
            this.button_11_5.TabIndex = 4;
            this.button_11_5.Text = "停止记录";
            this.button_11_5.UseVisualStyleBackColor = true;
            this.button_11_5.Click += new System.EventHandler(this.button_11_5_Click);
            // 
            // button_11_4
            // 
            this.button_11_4.Location = new System.Drawing.Point(183, 20);
            this.button_11_4.Name = "button_11_4";
            this.button_11_4.Size = new System.Drawing.Size(100, 30);
            this.button_11_4.TabIndex = 3;
            this.button_11_4.Text = "开始记录";
            this.button_11_4.UseVisualStyleBackColor = true;
            this.button_11_4.Click += new System.EventHandler(this.button_11_4_Click);
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.button_12_1);
            this.groupBox5.Controls.Add(this.button_11_1);
            this.groupBox5.Location = new System.Drawing.Point(6, 8);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(282, 65);
            this.groupBox5.TabIndex = 0;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "探头连接";
            // 
            // button_12_1
            // 
            this.button_12_1.Location = new System.Drawing.Point(157, 20);
            this.button_12_1.Name = "button_12_1";
            this.button_12_1.Size = new System.Drawing.Size(100, 30);
            this.button_12_1.TabIndex = 5;
            this.button_12_1.Text = "断开探头";
            this.button_12_1.UseVisualStyleBackColor = true;
            this.button_12_1.Click += new System.EventHandler(this.button_12_1_Click);
            // 
            // button_11_1
            // 
            this.button_11_1.Location = new System.Drawing.Point(25, 20);
            this.button_11_1.Name = "button_11_1";
            this.button_11_1.Size = new System.Drawing.Size(100, 30);
            this.button_11_1.TabIndex = 0;
            this.button_11_1.Text = "连接探头";
            this.button_11_1.UseVisualStyleBackColor = true;
            this.button_11_1.Click += new System.EventHandler(this.button_11_1_Click);
            // 
            // tabControl11
            // 
            this.tabControl11.Controls.Add(this.tabPage11);
            this.tabControl11.Controls.Add(this.tabPage12);
            this.tabControl11.Location = new System.Drawing.Point(6, 79);
            this.tabControl11.Name = "tabControl11";
            this.tabControl11.SelectedIndex = 0;
            this.tabControl11.Size = new System.Drawing.Size(200, 80);
            this.tabControl11.SizeMode = System.Windows.Forms.TabSizeMode.Fixed;
            this.tabControl11.TabIndex = 0;
            // 
            // tabPage11
            // 
            this.tabPage11.Controls.Add(this.label43);
            this.tabPage11.Controls.Add(this.textBox_11_1);
            this.tabPage11.Controls.Add(this.label38);
            this.tabPage11.Location = new System.Drawing.Point(8, 39);
            this.tabPage11.Name = "tabPage11";
            this.tabPage11.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage11.Size = new System.Drawing.Size(184, 33);
            this.tabPage11.TabIndex = 0;
            this.tabPage11.Text = "周期采集";
            this.tabPage11.UseVisualStyleBackColor = true;
            // 
            // label43
            // 
            this.label43.AutoSize = true;
            this.label43.Location = new System.Drawing.Point(160, 30);
            this.label43.Name = "label43";
            this.label43.Size = new System.Drawing.Size(34, 24);
            this.label43.TabIndex = 2;
            this.label43.Text = "秒";
            // 
            // textBox_11_1
            // 
            this.textBox_11_1.Location = new System.Drawing.Point(12, 27);
            this.textBox_11_1.Name = "textBox_11_1";
            this.textBox_11_1.Size = new System.Drawing.Size(140, 35);
            this.textBox_11_1.TabIndex = 1;
            this.textBox_11_1.Text = "1";
            this.textBox_11_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label38
            // 
            this.label38.AutoSize = true;
            this.label38.Location = new System.Drawing.Point(10, 10);
            this.label38.Name = "label38";
            this.label38.Size = new System.Drawing.Size(106, 24);
            this.label38.TabIndex = 0;
            this.label38.Text = "保存周期";
            // 
            // tabPage12
            // 
            this.tabPage12.Controls.Add(this.label37);
            this.tabPage12.Controls.Add(this.numericUpDown13);
            this.tabPage12.Controls.Add(this.label36);
            this.tabPage12.Controls.Add(this.numericUpDown12);
            this.tabPage12.Controls.Add(this.label35);
            this.tabPage12.Controls.Add(this.numericUpDown11);
            this.tabPage12.Location = new System.Drawing.Point(8, 39);
            this.tabPage12.Name = "tabPage12";
            this.tabPage12.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage12.Size = new System.Drawing.Size(184, 33);
            this.tabPage12.TabIndex = 1;
            this.tabPage12.Text = "角度触发";
            this.tabPage12.UseVisualStyleBackColor = true;
            // 
            // label37
            // 
            this.label37.AutoSize = true;
            this.label37.Location = new System.Drawing.Point(130, 10);
            this.label37.Name = "label37";
            this.label37.Size = new System.Drawing.Size(106, 24);
            this.label37.TabIndex = 5;
            this.label37.Text = "记录次数";
            // 
            // numericUpDown13
            // 
            this.numericUpDown13.Location = new System.Drawing.Point(126, 25);
            this.numericUpDown13.Maximum = new decimal(new int[] {
            359,
            0,
            0,
            0});
            this.numericUpDown13.Name = "numericUpDown13";
            this.numericUpDown13.Size = new System.Drawing.Size(60, 35);
            this.numericUpDown13.TabIndex = 4;
            this.numericUpDown13.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.numericUpDown13.UpDownAlign = System.Windows.Forms.LeftRightAlignment.Left;
            // 
            // label36
            // 
            this.label36.AutoSize = true;
            this.label36.Location = new System.Drawing.Point(70, 10);
            this.label36.Name = "label36";
            this.label36.Size = new System.Drawing.Size(106, 24);
            this.label36.TabIndex = 3;
            this.label36.Text = "步进角度";
            // 
            // numericUpDown12
            // 
            this.numericUpDown12.Location = new System.Drawing.Point(66, 25);
            this.numericUpDown12.Minimum = new decimal(new int[] {
            100,
            0,
            0,
            -2147483648});
            this.numericUpDown12.Name = "numericUpDown12";
            this.numericUpDown12.Size = new System.Drawing.Size(60, 35);
            this.numericUpDown12.TabIndex = 2;
            this.numericUpDown12.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.numericUpDown12.UpDownAlign = System.Windows.Forms.LeftRightAlignment.Left;
            // 
            // label35
            // 
            this.label35.AutoSize = true;
            this.label35.Location = new System.Drawing.Point(10, 10);
            this.label35.Name = "label35";
            this.label35.Size = new System.Drawing.Size(106, 24);
            this.label35.TabIndex = 1;
            this.label35.Text = "初始角度";
            // 
            // numericUpDown11
            // 
            this.numericUpDown11.Location = new System.Drawing.Point(6, 25);
            this.numericUpDown11.Maximum = new decimal(new int[] {
            359,
            0,
            0,
            0});
            this.numericUpDown11.Name = "numericUpDown11";
            this.numericUpDown11.Size = new System.Drawing.Size(60, 35);
            this.numericUpDown11.TabIndex = 0;
            this.numericUpDown11.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.numericUpDown11.UpDownAlign = System.Windows.Forms.LeftRightAlignment.Left;
            // 
            // groupBox_12_2
            // 
            this.groupBox_12_2.Controls.Add(this.button_12_004);
            this.groupBox_12_2.Controls.Add(this.button_12_003);
            this.groupBox_12_2.Location = new System.Drawing.Point(784, 8);
            this.groupBox_12_2.Name = "groupBox_12_2";
            this.groupBox_12_2.Size = new System.Drawing.Size(322, 65);
            this.groupBox_12_2.TabIndex = 7;
            this.groupBox_12_2.TabStop = false;
            this.groupBox_12_2.Text = "记录帧数据文档";
            // 
            // button_12_004
            // 
            this.button_12_004.Location = new System.Drawing.Point(187, 20);
            this.button_12_004.Name = "button_12_004";
            this.button_12_004.Size = new System.Drawing.Size(100, 30);
            this.button_12_004.TabIndex = 5;
            this.button_12_004.Text = "停止记录";
            this.button_12_004.UseVisualStyleBackColor = true;
            this.button_12_004.Click += new System.EventHandler(this.button_12_004_Click);
            // 
            // button_12_003
            // 
            this.button_12_003.Location = new System.Drawing.Point(41, 20);
            this.button_12_003.Name = "button_12_003";
            this.button_12_003.Size = new System.Drawing.Size(100, 30);
            this.button_12_003.TabIndex = 4;
            this.button_12_003.Text = "开始记录";
            this.button_12_003.UseVisualStyleBackColor = true;
            this.button_12_003.Click += new System.EventHandler(this.button_12_003_Click);
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.pictureBox11);
            this.groupBox7.Controls.Add(this.circularGauge1);
            this.groupBox7.Controls.Add(this.dataGridView1);
            this.groupBox7.Location = new System.Drawing.Point(209, 79);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(897, 244);
            this.groupBox7.TabIndex = 2;
            this.groupBox7.TabStop = false;
            // 
            // pictureBox11
            // 
            this.pictureBox11.BackColor = System.Drawing.Color.Gray;
            this.pictureBox11.Location = new System.Drawing.Point(616, 22);
            this.pictureBox11.Name = "pictureBox11";
            this.pictureBox11.Size = new System.Drawing.Size(20, 20);
            this.pictureBox11.TabIndex = 2;
            this.pictureBox11.TabStop = false;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(6, 13);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 82;
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(564, 224);
            this.dataGridView1.TabIndex = 0;
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.textBox_11_3);
            this.groupBox6.Controls.Add(this.checkedListBox2);
            this.groupBox6.Controls.Add(this.textBox_11_2);
            this.groupBox6.Controls.Add(this.label39);
            this.groupBox6.Location = new System.Drawing.Point(5, 157);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(200, 591);
            this.groupBox6.TabIndex = 1;
            this.groupBox6.TabStop = false;
            // 
            // textBox_11_3
            // 
            this.textBox_11_3.Enabled = false;
            this.textBox_11_3.Location = new System.Drawing.Point(6, 15);
            this.textBox_11_3.Name = "textBox_11_3";
            this.textBox_11_3.Size = new System.Drawing.Size(189, 35);
            this.textBox_11_3.TabIndex = 7;
            // 
            // checkedListBox2
            // 
            this.checkedListBox2.Enabled = false;
            this.checkedListBox2.FormattingEnabled = true;
            this.checkedListBox2.Location = new System.Drawing.Point(5, 45);
            this.checkedListBox2.Name = "checkedListBox2";
            this.checkedListBox2.Size = new System.Drawing.Size(189, 484);
            this.checkedListBox2.TabIndex = 6;
            // 
            // textBox_11_2
            // 
            this.textBox_11_2.Location = new System.Drawing.Point(64, 559);
            this.textBox_11_2.Name = "textBox_11_2";
            this.textBox_11_2.Size = new System.Drawing.Size(130, 35);
            this.textBox_11_2.TabIndex = 3;
            this.textBox_11_2.Text = "0";
            // 
            // label39
            // 
            this.label39.AutoSize = true;
            this.label39.Location = new System.Drawing.Point(10, 563);
            this.label39.Name = "label39";
            this.label39.Size = new System.Drawing.Size(106, 24);
            this.label39.TabIndex = 2;
            this.label39.Text = "采集次数";
            // 
            // groupBox8
            // 
            this.groupBox8.Controls.Add(this.button_11_9);
            this.groupBox8.Controls.Add(this.listBox1);
            this.groupBox8.Controls.Add(this.formsPlot1);
            this.groupBox8.Controls.Add(this.checkedListBox1);
            this.groupBox8.Controls.Add(this.chart1);
            this.groupBox8.Controls.Add(this.button_11_8);
            this.groupBox8.Controls.Add(this.button_11_7);
            this.groupBox8.Location = new System.Drawing.Point(209, 322);
            this.groupBox8.Name = "groupBox8";
            this.groupBox8.Size = new System.Drawing.Size(897, 420);
            this.groupBox8.TabIndex = 3;
            this.groupBox8.TabStop = false;
            // 
            // button_11_9
            // 
            this.button_11_9.Location = new System.Drawing.Point(762, 13);
            this.button_11_9.Name = "button_11_9";
            this.button_11_9.Size = new System.Drawing.Size(120, 30);
            this.button_11_9.TabIndex = 7;
            this.button_11_9.Text = "清除图像";
            this.button_11_9.UseVisualStyleBackColor = true;
            this.button_11_9.Click += new System.EventHandler(this.button_11_9_Click);
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 24;
            this.listBox1.Location = new System.Drawing.Point(762, 46);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(120, 292);
            this.listBox1.TabIndex = 6;
            // 
            // formsPlot1
            // 
            this.formsPlot1.DisplayScale = 0F;
            this.formsPlot1.Location = new System.Drawing.Point(6, 13);
            this.formsPlot1.Name = "formsPlot1";
            this.formsPlot1.Size = new System.Drawing.Size(747, 396);
            this.formsPlot1.TabIndex = 4;
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Location = new System.Drawing.Point(812, 89);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(69, 132);
            this.checkedListBox1.TabIndex = 5;
            // 
            // chart1
            // 
            chartArea1.Name = "ChartArea1";
            this.chart1.ChartAreas.Add(chartArea1);
            legend1.Enabled = false;
            legend1.Name = "Legend1";
            this.chart1.Legends.Add(legend1);
            this.chart1.Location = new System.Drawing.Point(6, 20);
            this.chart1.Name = "chart1";
            series1.ChartArea = "ChartArea1";
            series1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            series1.Legend = "Legend1";
            series1.Name = "Series1";
            this.chart1.Series.Add(series1);
            this.chart1.Size = new System.Drawing.Size(660, 389);
            this.chart1.TabIndex = 4;
            this.chart1.Text = "chart1";
            // 
            // button_11_8
            // 
            this.button_11_8.Location = new System.Drawing.Point(762, 379);
            this.button_11_8.Name = "button_11_8";
            this.button_11_8.Size = new System.Drawing.Size(120, 30);
            this.button_11_8.TabIndex = 2;
            this.button_11_8.Text = "全部隐藏";
            this.button_11_8.UseVisualStyleBackColor = true;
            this.button_11_8.Click += new System.EventHandler(this.button_11_8_Click);
            // 
            // button_11_7
            // 
            this.button_11_7.Location = new System.Drawing.Point(762, 344);
            this.button_11_7.Name = "button_11_7";
            this.button_11_7.Size = new System.Drawing.Size(120, 30);
            this.button_11_7.TabIndex = 1;
            this.button_11_7.Text = "全部可见";
            this.button_11_7.UseVisualStyleBackColor = true;
            this.button_11_7.Click += new System.EventHandler(this.button_11_7_Click);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.groupBox24);
            this.tabPage2.Controls.Add(this.groupBox18);
            this.tabPage2.Controls.Add(this.groupBox17);
            this.tabPage2.Controls.Add(this.groupBox16);
            this.tabPage2.ImageIndex = 4;
            this.tabPage2.Location = new System.Drawing.Point(84, 4);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1112, 792);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // groupBox24
            // 
            this.groupBox24.Controls.Add(this.textBox_23_3);
            this.groupBox24.Controls.Add(this.textBox_23_2);
            this.groupBox24.Controls.Add(this.textBox_23_1);
            this.groupBox24.Controls.Add(this.label152);
            this.groupBox24.Controls.Add(this.label151);
            this.groupBox24.Controls.Add(this.comboBox_23_1);
            this.groupBox24.Location = new System.Drawing.Point(672, 3);
            this.groupBox24.Name = "groupBox24";
            this.groupBox24.Size = new System.Drawing.Size(421, 155);
            this.groupBox24.TabIndex = 48;
            this.groupBox24.TabStop = false;
            // 
            // textBox_23_3
            // 
            this.textBox_23_3.Location = new System.Drawing.Point(258, 110);
            this.textBox_23_3.Name = "textBox_23_3";
            this.textBox_23_3.Size = new System.Drawing.Size(120, 35);
            this.textBox_23_3.TabIndex = 5;
            this.textBox_23_3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_23_2
            // 
            this.textBox_23_2.Location = new System.Drawing.Point(141, 110);
            this.textBox_23_2.Name = "textBox_23_2";
            this.textBox_23_2.Size = new System.Drawing.Size(120, 35);
            this.textBox_23_2.TabIndex = 4;
            this.textBox_23_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_23_1
            // 
            this.textBox_23_1.Location = new System.Drawing.Point(24, 110);
            this.textBox_23_1.Name = "textBox_23_1";
            this.textBox_23_1.Size = new System.Drawing.Size(120, 35);
            this.textBox_23_1.TabIndex = 3;
            this.textBox_23_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label152
            // 
            this.label152.AutoSize = true;
            this.label152.Location = new System.Drawing.Point(24, 79);
            this.label152.Name = "label152";
            this.label152.Size = new System.Drawing.Size(130, 24);
            this.label152.TabIndex = 2;
            this.label152.Text = "探头磁场值";
            // 
            // label151
            // 
            this.label151.AutoSize = true;
            this.label151.Location = new System.Drawing.Point(103, 45);
            this.label151.Name = "label151";
            this.label151.Size = new System.Drawing.Size(106, 24);
            this.label151.TabIndex = 1;
            this.label151.Text = "探头选择";
            // 
            // comboBox_23_1
            // 
            this.comboBox_23_1.FormattingEnabled = true;
            this.comboBox_23_1.Location = new System.Drawing.Point(173, 42);
            this.comboBox_23_1.Name = "comboBox_23_1";
            this.comboBox_23_1.Size = new System.Drawing.Size(114, 32);
            this.comboBox_23_1.TabIndex = 0;
            this.comboBox_23_1.SelectedIndexChanged += new System.EventHandler(this.comboBox_23_1_SelectedIndexChanged);
            // 
            // groupBox18
            // 
            this.groupBox18.Controls.Add(this.tabControl23);
            this.groupBox18.Controls.Add(this.comboBox_22_1);
            this.groupBox18.Controls.Add(this.formsPlot3);
            this.groupBox18.Controls.Add(this.label136);
            this.groupBox18.Controls.Add(this.button_22_2);
            this.groupBox18.Controls.Add(this.button_22_1);
            this.groupBox18.Location = new System.Drawing.Point(3, 458);
            this.groupBox18.Name = "groupBox18";
            this.groupBox18.Size = new System.Drawing.Size(1090, 297);
            this.groupBox18.TabIndex = 47;
            this.groupBox18.TabStop = false;
            this.groupBox18.Text = "电源7-8";
            // 
            // tabControl23
            // 
            this.tabControl23.Controls.Add(this.tabPage231);
            this.tabControl23.Controls.Add(this.tabPage232);
            this.tabControl23.Location = new System.Drawing.Point(16, 20);
            this.tabControl23.Name = "tabControl23";
            this.tabControl23.SelectedIndex = 0;
            this.tabControl23.Size = new System.Drawing.Size(190, 270);
            this.tabControl23.TabIndex = 57;
            // 
            // tabPage231
            // 
            this.tabPage231.Controls.Add(this.tabControl4);
            this.tabPage231.Location = new System.Drawing.Point(8, 39);
            this.tabPage231.Name = "tabPage231";
            this.tabPage231.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage231.Size = new System.Drawing.Size(174, 223);
            this.tabPage231.TabIndex = 0;
            this.tabPage231.Text = "退磁电源";
            this.tabPage231.UseVisualStyleBackColor = true;
            // 
            // tabControl4
            // 
            this.tabControl4.Controls.Add(this.tabPage10);
            this.tabControl4.Controls.Add(this.tabPage15);
            this.tabControl4.Location = new System.Drawing.Point(3, 6);
            this.tabControl4.Name = "tabControl4";
            this.tabControl4.SelectedIndex = 0;
            this.tabControl4.Size = new System.Drawing.Size(176, 235);
            this.tabControl4.TabIndex = 57;
            // 
            // tabPage10
            // 
            this.tabPage10.Controls.Add(this.textBox_22_3);
            this.tabPage10.Controls.Add(this.textBox_22_4);
            this.tabPage10.Controls.Add(this.label140);
            this.tabPage10.Controls.Add(this.textBox_22_1);
            this.tabPage10.Controls.Add(this.label138);
            this.tabPage10.Controls.Add(this.label141);
            this.tabPage10.Location = new System.Drawing.Point(8, 39);
            this.tabPage10.Name = "tabPage10";
            this.tabPage10.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage10.Size = new System.Drawing.Size(160, 188);
            this.tabPage10.TabIndex = 0;
            this.tabPage10.Text = "基础设置";
            this.tabPage10.UseVisualStyleBackColor = true;
            // 
            // textBox_22_3
            // 
            this.textBox_22_3.Location = new System.Drawing.Point(90, 84);
            this.textBox_22_3.Name = "textBox_22_3";
            this.textBox_22_3.Size = new System.Drawing.Size(70, 35);
            this.textBox_22_3.TabIndex = 51;
            this.textBox_22_3.Text = "1";
            // 
            // textBox_22_4
            // 
            this.textBox_22_4.Location = new System.Drawing.Point(90, 133);
            this.textBox_22_4.Name = "textBox_22_4";
            this.textBox_22_4.Size = new System.Drawing.Size(70, 35);
            this.textBox_22_4.TabIndex = 56;
            this.textBox_22_4.Text = "1";
            // 
            // label140
            // 
            this.label140.AutoSize = true;
            this.label140.Location = new System.Drawing.Point(11, 88);
            this.label140.Name = "label140";
            this.label140.Size = new System.Drawing.Size(58, 24);
            this.label140.TabIndex = 50;
            this.label140.Text = "频率";
            // 
            // textBox_22_1
            // 
            this.textBox_22_1.Location = new System.Drawing.Point(90, 34);
            this.textBox_22_1.Name = "textBox_22_1";
            this.textBox_22_1.Size = new System.Drawing.Size(70, 35);
            this.textBox_22_1.TabIndex = 47;
            this.textBox_22_1.Text = "1";
            // 
            // label138
            // 
            this.label138.AutoSize = true;
            this.label138.Location = new System.Drawing.Point(11, 39);
            this.label138.Name = "label138";
            this.label138.Size = new System.Drawing.Size(82, 24);
            this.label138.TabIndex = 44;
            this.label138.Text = "电流值";
            // 
            // label141
            // 
            this.label141.AutoSize = true;
            this.label141.Location = new System.Drawing.Point(11, 138);
            this.label141.Name = "label141";
            this.label141.Size = new System.Drawing.Size(118, 24);
            this.label141.TabIndex = 55;
            this.label141.Text = "总时间(s)";
            // 
            // tabPage15
            // 
            this.tabPage15.Controls.Add(this.groupBox20);
            this.tabPage15.Controls.Add(this.textBox_22_2);
            this.tabPage15.Controls.Add(this.label139);
            this.tabPage15.Controls.Add(this.groupBox19);
            this.tabPage15.Location = new System.Drawing.Point(8, 39);
            this.tabPage15.Name = "tabPage15";
            this.tabPage15.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage15.Size = new System.Drawing.Size(160, 188);
            this.tabPage15.TabIndex = 1;
            this.tabPage15.Text = "高级设置";
            this.tabPage15.UseVisualStyleBackColor = true;
            // 
            // groupBox20
            // 
            this.groupBox20.Controls.Add(this.radioButton_22_4);
            this.groupBox20.Controls.Add(this.radioButton_22_3);
            this.groupBox20.Location = new System.Drawing.Point(9, 88);
            this.groupBox20.Name = "groupBox20";
            this.groupBox20.Size = new System.Drawing.Size(150, 50);
            this.groupBox20.TabIndex = 46;
            this.groupBox20.TabStop = false;
            this.groupBox20.Text = "波形";
            // 
            // radioButton_22_4
            // 
            this.radioButton_22_4.AutoSize = true;
            this.radioButton_22_4.Location = new System.Drawing.Point(80, 20);
            this.radioButton_22_4.Name = "radioButton_22_4";
            this.radioButton_22_4.Size = new System.Drawing.Size(113, 28);
            this.radioButton_22_4.TabIndex = 1;
            this.radioButton_22_4.Text = "矩形波";
            this.radioButton_22_4.UseVisualStyleBackColor = true;
            // 
            // radioButton_22_3
            // 
            this.radioButton_22_3.AutoSize = true;
            this.radioButton_22_3.Checked = true;
            this.radioButton_22_3.Location = new System.Drawing.Point(20, 20);
            this.radioButton_22_3.Name = "radioButton_22_3";
            this.radioButton_22_3.Size = new System.Drawing.Size(113, 28);
            this.radioButton_22_3.TabIndex = 0;
            this.radioButton_22_3.TabStop = true;
            this.radioButton_22_3.Text = "正弦波";
            this.radioButton_22_3.UseVisualStyleBackColor = true;
            // 
            // textBox_22_2
            // 
            this.textBox_22_2.Location = new System.Drawing.Point(82, 157);
            this.textBox_22_2.Name = "textBox_22_2";
            this.textBox_22_2.Size = new System.Drawing.Size(80, 35);
            this.textBox_22_2.TabIndex = 49;
            this.textBox_22_2.Text = "0.1";
            // 
            // label139
            // 
            this.label139.AutoSize = true;
            this.label139.Location = new System.Drawing.Point(12, 162);
            this.label139.Name = "label139";
            this.label139.Size = new System.Drawing.Size(106, 24);
            this.label139.TabIndex = 48;
            this.label139.Text = "衰减系数";
            // 
            // groupBox19
            // 
            this.groupBox19.Controls.Add(this.radioButton_22_2);
            this.groupBox19.Controls.Add(this.radioButton_22_1);
            this.groupBox19.Location = new System.Drawing.Point(9, 20);
            this.groupBox19.Name = "groupBox19";
            this.groupBox19.Size = new System.Drawing.Size(150, 50);
            this.groupBox19.TabIndex = 45;
            this.groupBox19.TabStop = false;
            this.groupBox19.Text = "衰减方式";
            // 
            // radioButton_22_2
            // 
            this.radioButton_22_2.AutoSize = true;
            this.radioButton_22_2.Location = new System.Drawing.Point(80, 20);
            this.radioButton_22_2.Name = "radioButton_22_2";
            this.radioButton_22_2.Size = new System.Drawing.Size(89, 28);
            this.radioButton_22_2.TabIndex = 1;
            this.radioButton_22_2.Text = "线性";
            this.radioButton_22_2.UseVisualStyleBackColor = true;
            // 
            // radioButton_22_1
            // 
            this.radioButton_22_1.AutoSize = true;
            this.radioButton_22_1.Checked = true;
            this.radioButton_22_1.Location = new System.Drawing.Point(20, 20);
            this.radioButton_22_1.Name = "radioButton_22_1";
            this.radioButton_22_1.Size = new System.Drawing.Size(89, 28);
            this.radioButton_22_1.TabIndex = 0;
            this.radioButton_22_1.TabStop = true;
            this.radioButton_22_1.Text = "指数";
            this.radioButton_22_1.UseVisualStyleBackColor = true;
            // 
            // tabPage232
            // 
            this.tabPage232.Controls.Add(this.label142);
            this.tabPage232.Controls.Add(this.textBox_22_5);
            this.tabPage232.Controls.Add(this.label143);
            this.tabPage232.Controls.Add(this.textBox_22_6);
            this.tabPage232.Controls.Add(this.label144);
            this.tabPage232.Controls.Add(this.textBox_22_7);
            this.tabPage232.Location = new System.Drawing.Point(8, 39);
            this.tabPage232.Name = "tabPage232";
            this.tabPage232.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage232.Size = new System.Drawing.Size(174, 223);
            this.tabPage232.TabIndex = 1;
            this.tabPage232.Text = "充磁电源";
            this.tabPage232.UseVisualStyleBackColor = true;
            // 
            // label142
            // 
            this.label142.AutoSize = true;
            this.label142.Location = new System.Drawing.Point(10, 64);
            this.label142.Name = "label142";
            this.label142.Size = new System.Drawing.Size(166, 24);
            this.label142.TabIndex = 52;
            this.label142.Text = "磁场保持时间s";
            // 
            // textBox_22_5
            // 
            this.textBox_22_5.Location = new System.Drawing.Point(100, 59);
            this.textBox_22_5.Name = "textBox_22_5";
            this.textBox_22_5.Size = new System.Drawing.Size(70, 35);
            this.textBox_22_5.TabIndex = 53;
            this.textBox_22_5.Text = "1";
            // 
            // label143
            // 
            this.label143.AutoSize = true;
            this.label143.Location = new System.Drawing.Point(10, 117);
            this.label143.Name = "label143";
            this.label143.Size = new System.Drawing.Size(130, 24);
            this.label143.TabIndex = 54;
            this.label143.Text = "磁场强度mT";
            // 
            // textBox_22_6
            // 
            this.textBox_22_6.Location = new System.Drawing.Point(100, 112);
            this.textBox_22_6.Name = "textBox_22_6";
            this.textBox_22_6.Size = new System.Drawing.Size(70, 35);
            this.textBox_22_6.TabIndex = 55;
            this.textBox_22_6.Text = "1";
            // 
            // label144
            // 
            this.label144.AutoSize = true;
            this.label144.Location = new System.Drawing.Point(10, 173);
            this.label144.Name = "label144";
            this.label144.Size = new System.Drawing.Size(178, 24);
            this.label144.TabIndex = 56;
            this.label144.Text = "磁场变化率mT/s";
            // 
            // textBox_22_7
            // 
            this.textBox_22_7.Location = new System.Drawing.Point(100, 168);
            this.textBox_22_7.Name = "textBox_22_7";
            this.textBox_22_7.Size = new System.Drawing.Size(70, 35);
            this.textBox_22_7.TabIndex = 57;
            this.textBox_22_7.Text = "1";
            // 
            // comboBox_22_1
            // 
            this.comboBox_22_1.FormattingEnabled = true;
            this.comboBox_22_1.Items.AddRange(new object[] {
            "关闭",
            "电源7",
            "电源8"});
            this.comboBox_22_1.Location = new System.Drawing.Point(226, 44);
            this.comboBox_22_1.Name = "comboBox_22_1";
            this.comboBox_22_1.Size = new System.Drawing.Size(80, 32);
            this.comboBox_22_1.TabIndex = 26;
            this.comboBox_22_1.Text = "关闭";
            this.comboBox_22_1.SelectedIndexChanged += new System.EventHandler(this.comboBox_22_1_SelectedIndexChanged);
            // 
            // formsPlot3
            // 
            this.formsPlot3.DisplayScale = 0F;
            this.formsPlot3.Location = new System.Drawing.Point(336, 20);
            this.formsPlot3.Name = "formsPlot3";
            this.formsPlot3.Size = new System.Drawing.Size(748, 262);
            this.formsPlot3.TabIndex = 54;
            // 
            // label136
            // 
            this.label136.AutoSize = true;
            this.label136.ForeColor = System.Drawing.SystemColors.MenuHighlight;
            this.label136.Location = new System.Drawing.Point(226, 24);
            this.label136.Name = "label136";
            this.label136.Size = new System.Drawing.Size(106, 24);
            this.label136.TabIndex = 27;
            this.label136.Text = "电源选择";
            // 
            // button_22_2
            // 
            this.button_22_2.Location = new System.Drawing.Point(228, 239);
            this.button_22_2.Name = "button_22_2";
            this.button_22_2.Size = new System.Drawing.Size(75, 23);
            this.button_22_2.TabIndex = 53;
            this.button_22_2.Text = "停止输出";
            this.button_22_2.UseVisualStyleBackColor = true;
            this.button_22_2.Click += new System.EventHandler(this.button_22_2_Click);
            // 
            // button_22_1
            // 
            this.button_22_1.Location = new System.Drawing.Point(228, 185);
            this.button_22_1.Name = "button_22_1";
            this.button_22_1.Size = new System.Drawing.Size(75, 23);
            this.button_22_1.TabIndex = 52;
            this.button_22_1.Text = "开始输出";
            this.button_22_1.UseVisualStyleBackColor = true;
            this.button_22_1.Click += new System.EventHandler(this.button_22_1_Click);
            // 
            // groupBox17
            // 
            this.groupBox17.Controls.Add(this.formsPlot2);
            this.groupBox17.Controls.Add(this.comboBox_21_7);
            this.groupBox17.Controls.Add(this.button_21_2);
            this.groupBox17.Controls.Add(this.label127);
            this.groupBox17.Controls.Add(this.label133);
            this.groupBox17.Controls.Add(this.label128);
            this.groupBox17.Controls.Add(this.button_21_1);
            this.groupBox17.Controls.Add(this.comboBox_21_8);
            this.groupBox17.Controls.Add(this.label132);
            this.groupBox17.Controls.Add(this.label129);
            this.groupBox17.Controls.Add(this.label131);
            this.groupBox17.Controls.Add(this.comboBox_21_9);
            this.groupBox17.Controls.Add(this.textBox_21_8_3);
            this.groupBox17.Controls.Add(this.label130);
            this.groupBox17.Controls.Add(this.textBox_21_8_2);
            this.groupBox17.Controls.Add(this.textBox_21_7_1);
            this.groupBox17.Controls.Add(this.textBox_21_8_1);
            this.groupBox17.Controls.Add(this.textBox_21_7_2);
            this.groupBox17.Controls.Add(this.textBox_21_7_3);
            this.groupBox17.Location = new System.Drawing.Point(3, 164);
            this.groupBox17.Name = "groupBox17";
            this.groupBox17.Size = new System.Drawing.Size(1090, 288);
            this.groupBox17.TabIndex = 46;
            this.groupBox17.TabStop = false;
            this.groupBox17.Text = "电源1-6";
            // 
            // formsPlot2
            // 
            this.formsPlot2.DisplayScale = 0F;
            this.formsPlot2.Location = new System.Drawing.Point(336, 15);
            this.formsPlot2.Name = "formsPlot2";
            this.formsPlot2.Size = new System.Drawing.Size(748, 262);
            this.formsPlot2.TabIndex = 26;
            // 
            // comboBox_21_7
            // 
            this.comboBox_21_7.FormattingEnabled = true;
            this.comboBox_21_7.Items.AddRange(new object[] {
            "X",
            "Y",
            "Z",
            "-X",
            "-Y",
            "-Z"});
            this.comboBox_21_7.Location = new System.Drawing.Point(36, 55);
            this.comboBox_21_7.Name = "comboBox_21_7";
            this.comboBox_21_7.Size = new System.Drawing.Size(80, 32);
            this.comboBox_21_7.TabIndex = 27;
            this.comboBox_21_7.Text = "X";
            // 
            // button_21_2
            // 
            this.button_21_2.Location = new System.Drawing.Point(53, 230);
            this.button_21_2.Name = "button_21_2";
            this.button_21_2.Size = new System.Drawing.Size(100, 23);
            this.button_21_2.TabIndex = 44;
            this.button_21_2.Text = "查询电压电流";
            this.button_21_2.UseVisualStyleBackColor = true;
            this.button_21_2.Click += new System.EventHandler(this.button_21_2_Click);
            // 
            // label127
            // 
            this.label127.AutoSize = true;
            this.label127.Location = new System.Drawing.Point(37, 35);
            this.label127.Name = "label127";
            this.label127.Size = new System.Drawing.Size(106, 24);
            this.label127.TabIndex = 28;
            this.label127.Text = "坐标转换";
            // 
            // label133
            // 
            this.label133.AutoSize = true;
            this.label133.Location = new System.Drawing.Point(226, 203);
            this.label133.Name = "label133";
            this.label133.Size = new System.Drawing.Size(106, 24);
            this.label133.TabIndex = 43;
            this.label133.Text = "零场电源";
            // 
            // label128
            // 
            this.label128.AutoSize = true;
            this.label128.Location = new System.Drawing.Point(16, 58);
            this.label128.Name = "label128";
            this.label128.Size = new System.Drawing.Size(22, 24);
            this.label128.TabIndex = 29;
            this.label128.Text = "X";
            // 
            // button_21_1
            // 
            this.button_21_1.Location = new System.Drawing.Point(215, 230);
            this.button_21_1.Name = "button_21_1";
            this.button_21_1.Size = new System.Drawing.Size(75, 23);
            this.button_21_1.TabIndex = 42;
            this.button_21_1.Text = "磁场置零";
            this.button_21_1.UseVisualStyleBackColor = true;
            this.button_21_1.Click += new System.EventHandler(this.button_21_1_Click);
            // 
            // comboBox_21_8
            // 
            this.comboBox_21_8.FormattingEnabled = true;
            this.comboBox_21_8.Items.AddRange(new object[] {
            "X",
            "Y",
            "Z",
            "-X",
            "-Y",
            "-Z"});
            this.comboBox_21_8.Location = new System.Drawing.Point(36, 96);
            this.comboBox_21_8.Name = "comboBox_21_8";
            this.comboBox_21_8.Size = new System.Drawing.Size(80, 32);
            this.comboBox_21_8.TabIndex = 30;
            this.comboBox_21_8.Text = "Y";
            // 
            // label132
            // 
            this.label132.AutoSize = true;
            this.label132.Location = new System.Drawing.Point(237, 35);
            this.label132.Name = "label132";
            this.label132.Size = new System.Drawing.Size(106, 24);
            this.label132.TabIndex = 41;
            this.label132.Text = "设置电流";
            // 
            // label129
            // 
            this.label129.AutoSize = true;
            this.label129.Location = new System.Drawing.Point(16, 99);
            this.label129.Name = "label129";
            this.label129.Size = new System.Drawing.Size(22, 24);
            this.label129.TabIndex = 31;
            this.label129.Text = "Y";
            // 
            // label131
            // 
            this.label131.AutoSize = true;
            this.label131.Location = new System.Drawing.Point(137, 35);
            this.label131.Name = "label131";
            this.label131.Size = new System.Drawing.Size(106, 24);
            this.label131.TabIndex = 40;
            this.label131.Text = "设置磁场";
            // 
            // comboBox_21_9
            // 
            this.comboBox_21_9.FormattingEnabled = true;
            this.comboBox_21_9.Items.AddRange(new object[] {
            "X",
            "Y",
            "Z",
            "-X",
            "-Y",
            "-Z"});
            this.comboBox_21_9.Location = new System.Drawing.Point(36, 135);
            this.comboBox_21_9.Name = "comboBox_21_9";
            this.comboBox_21_9.Size = new System.Drawing.Size(80, 32);
            this.comboBox_21_9.TabIndex = 32;
            this.comboBox_21_9.Text = "Z";
            // 
            // textBox_21_8_3
            // 
            this.textBox_21_8_3.Location = new System.Drawing.Point(236, 135);
            this.textBox_21_8_3.Name = "textBox_21_8_3";
            this.textBox_21_8_3.Size = new System.Drawing.Size(80, 35);
            this.textBox_21_8_3.TabIndex = 39;
            this.textBox_21_8_3.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox_21_8_3_KeyDown);
            // 
            // label130
            // 
            this.label130.AutoSize = true;
            this.label130.Location = new System.Drawing.Point(16, 138);
            this.label130.Name = "label130";
            this.label130.Size = new System.Drawing.Size(22, 24);
            this.label130.TabIndex = 33;
            this.label130.Text = "Z";
            // 
            // textBox_21_8_2
            // 
            this.textBox_21_8_2.Location = new System.Drawing.Point(236, 96);
            this.textBox_21_8_2.Name = "textBox_21_8_2";
            this.textBox_21_8_2.Size = new System.Drawing.Size(80, 35);
            this.textBox_21_8_2.TabIndex = 38;
            this.textBox_21_8_2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox_21_8_2_KeyDown);
            // 
            // textBox_21_7_1
            // 
            this.textBox_21_7_1.Location = new System.Drawing.Point(136, 55);
            this.textBox_21_7_1.Name = "textBox_21_7_1";
            this.textBox_21_7_1.Size = new System.Drawing.Size(80, 35);
            this.textBox_21_7_1.TabIndex = 34;
            this.textBox_21_7_1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox_21_7_1_KeyDown);
            // 
            // textBox_21_8_1
            // 
            this.textBox_21_8_1.Location = new System.Drawing.Point(236, 55);
            this.textBox_21_8_1.Name = "textBox_21_8_1";
            this.textBox_21_8_1.Size = new System.Drawing.Size(80, 35);
            this.textBox_21_8_1.TabIndex = 37;
            this.textBox_21_8_1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox_21_8_1_KeyDown);
            // 
            // textBox_21_7_2
            // 
            this.textBox_21_7_2.Location = new System.Drawing.Point(136, 96);
            this.textBox_21_7_2.Name = "textBox_21_7_2";
            this.textBox_21_7_2.Size = new System.Drawing.Size(80, 35);
            this.textBox_21_7_2.TabIndex = 35;
            this.textBox_21_7_2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox_21_7_2_KeyDown);
            // 
            // textBox_21_7_3
            // 
            this.textBox_21_7_3.Location = new System.Drawing.Point(136, 135);
            this.textBox_21_7_3.Name = "textBox_21_7_3";
            this.textBox_21_7_3.Size = new System.Drawing.Size(80, 35);
            this.textBox_21_7_3.TabIndex = 36;
            this.textBox_21_7_3.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox_21_7_3_KeyDown);
            // 
            // groupBox16
            // 
            this.groupBox16.Controls.Add(this.textBox_21_4_1);
            this.groupBox16.Controls.Add(this.comboBox_21_1);
            this.groupBox16.Controls.Add(this.label119);
            this.groupBox16.Controls.Add(this.comboBox_21_2);
            this.groupBox16.Controls.Add(this.label120);
            this.groupBox16.Controls.Add(this.comboBox_21_3);
            this.groupBox16.Controls.Add(this.label121);
            this.groupBox16.Controls.Add(this.comboBox_21_4);
            this.groupBox16.Controls.Add(this.label122);
            this.groupBox16.Controls.Add(this.comboBox_21_5);
            this.groupBox16.Controls.Add(this.label123);
            this.groupBox16.Controls.Add(this.comboBox_21_6);
            this.groupBox16.Controls.Add(this.label124);
            this.groupBox16.Controls.Add(this.label125);
            this.groupBox16.Controls.Add(this.textBox_21_1_1);
            this.groupBox16.Controls.Add(this.textBox_21_2_1);
            this.groupBox16.Controls.Add(this.textBox_21_3_1);
            this.groupBox16.Controls.Add(this.textBox_21_5_1);
            this.groupBox16.Controls.Add(this.textBox_21_6_1);
            this.groupBox16.Controls.Add(this.textBox_21_1_2);
            this.groupBox16.Controls.Add(this.label126);
            this.groupBox16.Controls.Add(this.textBox_21_2_2);
            this.groupBox16.Controls.Add(this.textBox_21_6_2);
            this.groupBox16.Controls.Add(this.textBox_21_3_2);
            this.groupBox16.Controls.Add(this.textBox_21_5_2);
            this.groupBox16.Controls.Add(this.textBox_21_4_2);
            this.groupBox16.Location = new System.Drawing.Point(3, 3);
            this.groupBox16.Name = "groupBox16";
            this.groupBox16.Size = new System.Drawing.Size(654, 155);
            this.groupBox16.TabIndex = 45;
            this.groupBox16.TabStop = false;
            this.groupBox16.Text = "电源连接";
            // 
            // textBox_21_4_1
            // 
            this.textBox_21_4_1.Location = new System.Drawing.Point(352, 79);
            this.textBox_21_4_1.Name = "textBox_21_4_1";
            this.textBox_21_4_1.ReadOnly = true;
            this.textBox_21_4_1.Size = new System.Drawing.Size(80, 35);
            this.textBox_21_4_1.TabIndex = 16;
            this.textBox_21_4_1.Text = "0";
            // 
            // comboBox_21_1
            // 
            this.comboBox_21_1.FormattingEnabled = true;
            this.comboBox_21_1.Items.AddRange(new object[] {
            "关闭",
            "恒场X",
            "恒场Y",
            "恒场Z",
            "零场X",
            "零场Y",
            "零场Z"});
            this.comboBox_21_1.Location = new System.Drawing.Point(53, 42);
            this.comboBox_21_1.Name = "comboBox_21_1";
            this.comboBox_21_1.Size = new System.Drawing.Size(80, 32);
            this.comboBox_21_1.TabIndex = 0;
            this.comboBox_21_1.Text = "关闭";
            this.comboBox_21_1.SelectedIndexChanged += new System.EventHandler(this.comboBox_21_1_SelectedIndexChanged_1);
            // 
            // label119
            // 
            this.label119.AutoSize = true;
            this.label119.Location = new System.Drawing.Point(53, 22);
            this.label119.Name = "label119";
            this.label119.Size = new System.Drawing.Size(70, 24);
            this.label119.TabIndex = 1;
            this.label119.Text = "电源1";
            // 
            // comboBox_21_2
            // 
            this.comboBox_21_2.FormattingEnabled = true;
            this.comboBox_21_2.Items.AddRange(new object[] {
            "关闭",
            "恒场X",
            "恒场Y",
            "恒场Z",
            "零场X",
            "零场Y",
            "零场Z"});
            this.comboBox_21_2.Location = new System.Drawing.Point(153, 42);
            this.comboBox_21_2.Name = "comboBox_21_2";
            this.comboBox_21_2.Size = new System.Drawing.Size(80, 32);
            this.comboBox_21_2.TabIndex = 2;
            this.comboBox_21_2.Text = "关闭";
            this.comboBox_21_2.SelectedIndexChanged += new System.EventHandler(this.comboBox_21_2_SelectedIndexChanged_1);
            // 
            // label120
            // 
            this.label120.AutoSize = true;
            this.label120.Location = new System.Drawing.Point(153, 22);
            this.label120.Name = "label120";
            this.label120.Size = new System.Drawing.Size(70, 24);
            this.label120.TabIndex = 3;
            this.label120.Text = "电源2";
            // 
            // comboBox_21_3
            // 
            this.comboBox_21_3.FormattingEnabled = true;
            this.comboBox_21_3.Items.AddRange(new object[] {
            "关闭",
            "恒场X",
            "恒场Y",
            "恒场Z",
            "零场X",
            "零场Y",
            "零场Z"});
            this.comboBox_21_3.Location = new System.Drawing.Point(253, 42);
            this.comboBox_21_3.Name = "comboBox_21_3";
            this.comboBox_21_3.Size = new System.Drawing.Size(80, 32);
            this.comboBox_21_3.TabIndex = 4;
            this.comboBox_21_3.Text = "关闭";
            this.comboBox_21_3.SelectedIndexChanged += new System.EventHandler(this.comboBox_21_3_SelectedIndexChanged_1);
            // 
            // label121
            // 
            this.label121.AutoSize = true;
            this.label121.Location = new System.Drawing.Point(253, 22);
            this.label121.Name = "label121";
            this.label121.Size = new System.Drawing.Size(70, 24);
            this.label121.TabIndex = 5;
            this.label121.Text = "电源3";
            // 
            // comboBox_21_4
            // 
            this.comboBox_21_4.FormattingEnabled = true;
            this.comboBox_21_4.Items.AddRange(new object[] {
            "关闭",
            "恒场X",
            "恒场Y",
            "恒场Z",
            "零场X",
            "零场Y",
            "零场Z"});
            this.comboBox_21_4.Location = new System.Drawing.Point(353, 42);
            this.comboBox_21_4.Name = "comboBox_21_4";
            this.comboBox_21_4.Size = new System.Drawing.Size(80, 32);
            this.comboBox_21_4.TabIndex = 6;
            this.comboBox_21_4.Text = "关闭";
            this.comboBox_21_4.SelectedIndexChanged += new System.EventHandler(this.comboBox_21_4_SelectedIndexChanged_1);
            // 
            // label122
            // 
            this.label122.AutoSize = true;
            this.label122.Location = new System.Drawing.Point(353, 22);
            this.label122.Name = "label122";
            this.label122.Size = new System.Drawing.Size(70, 24);
            this.label122.TabIndex = 7;
            this.label122.Text = "电源4";
            // 
            // comboBox_21_5
            // 
            this.comboBox_21_5.FormattingEnabled = true;
            this.comboBox_21_5.Items.AddRange(new object[] {
            "关闭",
            "恒场X",
            "恒场Y",
            "恒场Z",
            "零场X",
            "零场Y",
            "零场Z"});
            this.comboBox_21_5.Location = new System.Drawing.Point(453, 42);
            this.comboBox_21_5.Name = "comboBox_21_5";
            this.comboBox_21_5.Size = new System.Drawing.Size(80, 32);
            this.comboBox_21_5.TabIndex = 8;
            this.comboBox_21_5.Text = "关闭";
            this.comboBox_21_5.SelectedIndexChanged += new System.EventHandler(this.comboBox_21_5_SelectedIndexChanged_1);
            // 
            // label123
            // 
            this.label123.AutoSize = true;
            this.label123.Location = new System.Drawing.Point(453, 22);
            this.label123.Name = "label123";
            this.label123.Size = new System.Drawing.Size(70, 24);
            this.label123.TabIndex = 9;
            this.label123.Text = "电源5";
            // 
            // comboBox_21_6
            // 
            this.comboBox_21_6.FormattingEnabled = true;
            this.comboBox_21_6.Items.AddRange(new object[] {
            "关闭",
            "恒场X",
            "恒场Y",
            "恒场Z",
            "零场X",
            "零场Y",
            "零场Z"});
            this.comboBox_21_6.Location = new System.Drawing.Point(553, 42);
            this.comboBox_21_6.Name = "comboBox_21_6";
            this.comboBox_21_6.Size = new System.Drawing.Size(80, 32);
            this.comboBox_21_6.TabIndex = 10;
            this.comboBox_21_6.Text = "关闭";
            this.comboBox_21_6.SelectedIndexChanged += new System.EventHandler(this.comboBox_21_6_SelectedIndexChanged_1);
            // 
            // label124
            // 
            this.label124.AutoSize = true;
            this.label124.Location = new System.Drawing.Point(553, 22);
            this.label124.Name = "label124";
            this.label124.Size = new System.Drawing.Size(70, 24);
            this.label124.TabIndex = 11;
            this.label124.Text = "电源6";
            // 
            // label125
            // 
            this.label125.AutoSize = true;
            this.label125.Location = new System.Drawing.Point(22, 86);
            this.label125.Name = "label125";
            this.label125.Size = new System.Drawing.Size(58, 24);
            this.label125.TabIndex = 12;
            this.label125.Text = "电压";
            // 
            // textBox_21_1_1
            // 
            this.textBox_21_1_1.Location = new System.Drawing.Point(53, 82);
            this.textBox_21_1_1.Name = "textBox_21_1_1";
            this.textBox_21_1_1.ReadOnly = true;
            this.textBox_21_1_1.Size = new System.Drawing.Size(80, 35);
            this.textBox_21_1_1.TabIndex = 13;
            this.textBox_21_1_1.Text = "0";
            // 
            // textBox_21_2_1
            // 
            this.textBox_21_2_1.Location = new System.Drawing.Point(152, 82);
            this.textBox_21_2_1.Name = "textBox_21_2_1";
            this.textBox_21_2_1.ReadOnly = true;
            this.textBox_21_2_1.Size = new System.Drawing.Size(80, 35);
            this.textBox_21_2_1.TabIndex = 14;
            this.textBox_21_2_1.Text = "0";
            // 
            // textBox_21_3_1
            // 
            this.textBox_21_3_1.Location = new System.Drawing.Point(254, 82);
            this.textBox_21_3_1.Name = "textBox_21_3_1";
            this.textBox_21_3_1.ReadOnly = true;
            this.textBox_21_3_1.Size = new System.Drawing.Size(80, 35);
            this.textBox_21_3_1.TabIndex = 15;
            this.textBox_21_3_1.Text = "0";
            // 
            // textBox_21_5_1
            // 
            this.textBox_21_5_1.Location = new System.Drawing.Point(452, 79);
            this.textBox_21_5_1.Name = "textBox_21_5_1";
            this.textBox_21_5_1.ReadOnly = true;
            this.textBox_21_5_1.Size = new System.Drawing.Size(80, 35);
            this.textBox_21_5_1.TabIndex = 17;
            this.textBox_21_5_1.Text = "0";
            // 
            // textBox_21_6_1
            // 
            this.textBox_21_6_1.Location = new System.Drawing.Point(552, 79);
            this.textBox_21_6_1.Name = "textBox_21_6_1";
            this.textBox_21_6_1.ReadOnly = true;
            this.textBox_21_6_1.Size = new System.Drawing.Size(80, 35);
            this.textBox_21_6_1.TabIndex = 18;
            this.textBox_21_6_1.Text = "0";
            // 
            // textBox_21_1_2
            // 
            this.textBox_21_1_2.Location = new System.Drawing.Point(52, 122);
            this.textBox_21_1_2.Name = "textBox_21_1_2";
            this.textBox_21_1_2.ReadOnly = true;
            this.textBox_21_1_2.Size = new System.Drawing.Size(80, 35);
            this.textBox_21_1_2.TabIndex = 19;
            this.textBox_21_1_2.Text = "0";
            // 
            // label126
            // 
            this.label126.AutoSize = true;
            this.label126.Location = new System.Drawing.Point(22, 126);
            this.label126.Name = "label126";
            this.label126.Size = new System.Drawing.Size(58, 24);
            this.label126.TabIndex = 25;
            this.label126.Text = "电流";
            // 
            // textBox_21_2_2
            // 
            this.textBox_21_2_2.Location = new System.Drawing.Point(152, 122);
            this.textBox_21_2_2.Name = "textBox_21_2_2";
            this.textBox_21_2_2.ReadOnly = true;
            this.textBox_21_2_2.Size = new System.Drawing.Size(80, 35);
            this.textBox_21_2_2.TabIndex = 20;
            this.textBox_21_2_2.Text = "0";
            // 
            // textBox_21_6_2
            // 
            this.textBox_21_6_2.Location = new System.Drawing.Point(552, 119);
            this.textBox_21_6_2.Name = "textBox_21_6_2";
            this.textBox_21_6_2.ReadOnly = true;
            this.textBox_21_6_2.Size = new System.Drawing.Size(80, 35);
            this.textBox_21_6_2.TabIndex = 24;
            this.textBox_21_6_2.Text = "0";
            // 
            // textBox_21_3_2
            // 
            this.textBox_21_3_2.Location = new System.Drawing.Point(254, 122);
            this.textBox_21_3_2.Name = "textBox_21_3_2";
            this.textBox_21_3_2.ReadOnly = true;
            this.textBox_21_3_2.Size = new System.Drawing.Size(80, 35);
            this.textBox_21_3_2.TabIndex = 21;
            this.textBox_21_3_2.Text = "0";
            // 
            // textBox_21_5_2
            // 
            this.textBox_21_5_2.Location = new System.Drawing.Point(452, 119);
            this.textBox_21_5_2.Name = "textBox_21_5_2";
            this.textBox_21_5_2.ReadOnly = true;
            this.textBox_21_5_2.Size = new System.Drawing.Size(80, 35);
            this.textBox_21_5_2.TabIndex = 23;
            this.textBox_21_5_2.Text = "0";
            // 
            // textBox_21_4_2
            // 
            this.textBox_21_4_2.Location = new System.Drawing.Point(352, 119);
            this.textBox_21_4_2.Name = "textBox_21_4_2";
            this.textBox_21_4_2.ReadOnly = true;
            this.textBox_21_4_2.Size = new System.Drawing.Size(80, 35);
            this.textBox_21_4_2.TabIndex = 22;
            this.textBox_21_4_2.Text = "0";
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.tabControl2);
            this.tabPage3.ImageIndex = 6;
            this.tabPage3.Location = new System.Drawing.Point(84, 4);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(1112, 792);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // tabControl2
            // 
            this.tabControl2.Alignment = System.Windows.Forms.TabAlignment.Left;
            this.tabControl2.Controls.Add(this.tabPage7);
            this.tabControl2.Controls.Add(this.tabPage8);
            this.tabControl2.ImageList = this.imageList1;
            this.tabControl2.ItemSize = new System.Drawing.Size(80, 80);
            this.tabControl2.Location = new System.Drawing.Point(6, 6);
            this.tabControl2.Multiline = true;
            this.tabControl2.Name = "tabControl2";
            this.tabControl2.SelectedIndex = 0;
            this.tabControl2.Size = new System.Drawing.Size(1100, 760);
            this.tabControl2.SizeMode = System.Windows.Forms.TabSizeMode.Fixed;
            this.tabControl2.TabIndex = 0;
            // 
            // tabPage7
            // 
            this.tabPage7.Controls.Add(this.groupBox4);
            this.tabPage7.Controls.Add(this.groupBox3);
            this.tabPage7.Controls.Add(this.groupBox2);
            this.tabPage7.Controls.Add(this.groupBox1);
            this.tabPage7.ImageIndex = 10;
            this.tabPage7.Location = new System.Drawing.Point(84, 4);
            this.tabPage7.Name = "tabPage7";
            this.tabPage7.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage7.Size = new System.Drawing.Size(1012, 752);
            this.tabPage7.TabIndex = 0;
            this.tabPage7.UseVisualStyleBackColor = true;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.label29);
            this.groupBox4.Controls.Add(this.label30);
            this.groupBox4.Controls.Add(this.label31);
            this.groupBox4.Controls.Add(this.label32);
            this.groupBox4.Controls.Add(this.label33);
            this.groupBox4.Controls.Add(this.textBox_31_33);
            this.groupBox4.Controls.Add(this.textBox_31_34);
            this.groupBox4.Controls.Add(this.textBox_31_35);
            this.groupBox4.Controls.Add(this.textBox_31_36);
            this.groupBox4.Controls.Add(this.label34);
            this.groupBox4.Controls.Add(this.textBox_31_32);
            this.groupBox4.Controls.Add(this.label21);
            this.groupBox4.Controls.Add(this.label22);
            this.groupBox4.Controls.Add(this.label23);
            this.groupBox4.Controls.Add(this.label24);
            this.groupBox4.Controls.Add(this.label25);
            this.groupBox4.Controls.Add(this.label26);
            this.groupBox4.Controls.Add(this.label27);
            this.groupBox4.Controls.Add(this.textBox_31_17);
            this.groupBox4.Controls.Add(this.textBox_31_18);
            this.groupBox4.Controls.Add(this.textBox_31_19);
            this.groupBox4.Controls.Add(this.textBox_31_20);
            this.groupBox4.Controls.Add(this.textBox_31_21);
            this.groupBox4.Controls.Add(this.textBox_31_22);
            this.groupBox4.Controls.Add(this.textBox_31_23);
            this.groupBox4.Controls.Add(this.textBox_31_24);
            this.groupBox4.Controls.Add(this.textBox_31_25);
            this.groupBox4.Controls.Add(this.textBox_31_26);
            this.groupBox4.Controls.Add(this.textBox_31_27);
            this.groupBox4.Controls.Add(this.textBox_31_28);
            this.groupBox4.Controls.Add(this.textBox_31_29);
            this.groupBox4.Controls.Add(this.textBox_31_30);
            this.groupBox4.Controls.Add(this.textBox_31_31);
            this.groupBox4.Controls.Add(this.label28);
            this.groupBox4.Controls.Add(this.label20);
            this.groupBox4.Controls.Add(this.label19);
            this.groupBox4.Controls.Add(this.label18);
            this.groupBox4.Controls.Add(this.label17);
            this.groupBox4.Controls.Add(this.label16);
            this.groupBox4.Controls.Add(this.label15);
            this.groupBox4.Controls.Add(this.label14);
            this.groupBox4.Controls.Add(this.textBox_31_16);
            this.groupBox4.Controls.Add(this.textBox_31_15);
            this.groupBox4.Controls.Add(this.textBox_31_14);
            this.groupBox4.Controls.Add(this.textBox_31_13);
            this.groupBox4.Controls.Add(this.textBox_31_12);
            this.groupBox4.Controls.Add(this.textBox_31_11);
            this.groupBox4.Controls.Add(this.textBox_31_10);
            this.groupBox4.Controls.Add(this.textBox_31_9);
            this.groupBox4.Controls.Add(this.textBox_31_5);
            this.groupBox4.Controls.Add(this.textBox_31_6);
            this.groupBox4.Controls.Add(this.textBox_31_7);
            this.groupBox4.Controls.Add(this.textBox_31_8);
            this.groupBox4.Controls.Add(this.textBox_31_3);
            this.groupBox4.Controls.Add(this.textBox_31_4);
            this.groupBox4.Controls.Add(this.textBox_31_2);
            this.groupBox4.Controls.Add(this.textBox_31_1);
            this.groupBox4.Controls.Add(this.label13);
            this.groupBox4.Location = new System.Drawing.Point(444, 6);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(522, 655);
            this.groupBox4.TabIndex = 3;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "总磁矩";
            // 
            // label29
            // 
            this.label29.AutoSize = true;
            this.label29.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label29.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label29.Location = new System.Drawing.Point(30, 514);
            this.label29.Name = "label29";
            this.label29.Size = new System.Drawing.Size(118, 24);
            this.label29.TabIndex = 66;
            this.label29.Text = "123号探头";
            // 
            // label30
            // 
            this.label30.AutoSize = true;
            this.label30.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label30.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label30.Location = new System.Drawing.Point(375, 494);
            this.label30.Name = "label30";
            this.label30.Size = new System.Drawing.Size(46, 24);
            this.label30.TabIndex = 65;
            this.label30.Text = "SUM";
            // 
            // label31
            // 
            this.label31.AutoSize = true;
            this.label31.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label31.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label31.Location = new System.Drawing.Point(295, 494);
            this.label31.Name = "label31";
            this.label31.Size = new System.Drawing.Size(34, 24);
            this.label31.TabIndex = 64;
            this.label31.Text = "Mz";
            // 
            // label32
            // 
            this.label32.AutoSize = true;
            this.label32.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label32.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label32.Location = new System.Drawing.Point(210, 494);
            this.label32.Name = "label32";
            this.label32.Size = new System.Drawing.Size(34, 24);
            this.label32.TabIndex = 63;
            this.label32.Text = "My";
            // 
            // label33
            // 
            this.label33.AutoSize = true;
            this.label33.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label33.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label33.Location = new System.Drawing.Point(125, 494);
            this.label33.Name = "label33";
            this.label33.Size = new System.Drawing.Size(34, 24);
            this.label33.TabIndex = 62;
            this.label33.Text = "Mx";
            // 
            // textBox_31_33
            // 
            this.textBox_31_33.Location = new System.Drawing.Point(90, 511);
            this.textBox_31_33.Name = "textBox_31_33";
            this.textBox_31_33.ReadOnly = true;
            this.textBox_31_33.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_33.TabIndex = 61;
            this.textBox_31_33.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_34
            // 
            this.textBox_31_34.Location = new System.Drawing.Point(176, 511);
            this.textBox_31_34.Name = "textBox_31_34";
            this.textBox_31_34.ReadOnly = true;
            this.textBox_31_34.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_34.TabIndex = 60;
            this.textBox_31_34.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_35
            // 
            this.textBox_31_35.Location = new System.Drawing.Point(262, 511);
            this.textBox_31_35.Name = "textBox_31_35";
            this.textBox_31_35.ReadOnly = true;
            this.textBox_31_35.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_35.TabIndex = 59;
            this.textBox_31_35.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_36
            // 
            this.textBox_31_36.Location = new System.Drawing.Point(348, 511);
            this.textBox_31_36.Name = "textBox_31_36";
            this.textBox_31_36.ReadOnly = true;
            this.textBox_31_36.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_36.TabIndex = 58;
            this.textBox_31_36.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label34
            // 
            this.label34.AutoSize = true;
            this.label34.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label34.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label34.Location = new System.Drawing.Point(50, 466);
            this.label34.Name = "label34";
            this.label34.Size = new System.Drawing.Size(110, 24);
            this.label34.TabIndex = 57;
            this.label34.Text = "三个探头";
            // 
            // textBox_31_32
            // 
            this.textBox_31_32.Location = new System.Drawing.Point(348, 403);
            this.textBox_31_32.Name = "textBox_31_32";
            this.textBox_31_32.ReadOnly = true;
            this.textBox_31_32.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_32.TabIndex = 56;
            this.textBox_31_32.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label21.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label21.Location = new System.Drawing.Point(35, 366);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(106, 24);
            this.label21.TabIndex = 55;
            this.label21.Text = "13号探头";
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label22.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label22.Location = new System.Drawing.Point(35, 331);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(106, 24);
            this.label22.TabIndex = 54;
            this.label22.Text = "23号探头";
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label23.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label23.Location = new System.Drawing.Point(35, 296);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(106, 24);
            this.label23.TabIndex = 53;
            this.label23.Text = "12号探头";
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label24.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label24.Location = new System.Drawing.Point(375, 276);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(46, 24);
            this.label24.TabIndex = 52;
            this.label24.Text = "SUM";
            // 
            // label25
            // 
            this.label25.AutoSize = true;
            this.label25.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label25.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label25.Location = new System.Drawing.Point(295, 276);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(34, 24);
            this.label25.TabIndex = 51;
            this.label25.Text = "Mz";
            // 
            // label26
            // 
            this.label26.AutoSize = true;
            this.label26.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label26.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label26.Location = new System.Drawing.Point(210, 276);
            this.label26.Name = "label26";
            this.label26.Size = new System.Drawing.Size(34, 24);
            this.label26.TabIndex = 50;
            this.label26.Text = "My";
            // 
            // label27
            // 
            this.label27.AutoSize = true;
            this.label27.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label27.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label27.Location = new System.Drawing.Point(125, 276);
            this.label27.Name = "label27";
            this.label27.Size = new System.Drawing.Size(34, 24);
            this.label27.TabIndex = 49;
            this.label27.Text = "Mx";
            // 
            // textBox_31_17
            // 
            this.textBox_31_17.Location = new System.Drawing.Point(90, 293);
            this.textBox_31_17.Name = "textBox_31_17";
            this.textBox_31_17.ReadOnly = true;
            this.textBox_31_17.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_17.TabIndex = 48;
            this.textBox_31_17.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_18
            // 
            this.textBox_31_18.Location = new System.Drawing.Point(176, 293);
            this.textBox_31_18.Name = "textBox_31_18";
            this.textBox_31_18.ReadOnly = true;
            this.textBox_31_18.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_18.TabIndex = 47;
            this.textBox_31_18.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_19
            // 
            this.textBox_31_19.Location = new System.Drawing.Point(262, 293);
            this.textBox_31_19.Name = "textBox_31_19";
            this.textBox_31_19.ReadOnly = true;
            this.textBox_31_19.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_19.TabIndex = 46;
            this.textBox_31_19.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_20
            // 
            this.textBox_31_20.Location = new System.Drawing.Point(348, 293);
            this.textBox_31_20.Name = "textBox_31_20";
            this.textBox_31_20.ReadOnly = true;
            this.textBox_31_20.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_20.TabIndex = 45;
            this.textBox_31_20.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_21
            // 
            this.textBox_31_21.Location = new System.Drawing.Point(90, 325);
            this.textBox_31_21.Name = "textBox_31_21";
            this.textBox_31_21.ReadOnly = true;
            this.textBox_31_21.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_21.TabIndex = 44;
            this.textBox_31_21.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_22
            // 
            this.textBox_31_22.Location = new System.Drawing.Point(176, 325);
            this.textBox_31_22.Name = "textBox_31_22";
            this.textBox_31_22.ReadOnly = true;
            this.textBox_31_22.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_22.TabIndex = 43;
            this.textBox_31_22.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_23
            // 
            this.textBox_31_23.Location = new System.Drawing.Point(262, 325);
            this.textBox_31_23.Name = "textBox_31_23";
            this.textBox_31_23.ReadOnly = true;
            this.textBox_31_23.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_23.TabIndex = 42;
            this.textBox_31_23.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_24
            // 
            this.textBox_31_24.Location = new System.Drawing.Point(348, 325);
            this.textBox_31_24.Name = "textBox_31_24";
            this.textBox_31_24.ReadOnly = true;
            this.textBox_31_24.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_24.TabIndex = 41;
            this.textBox_31_24.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_25
            // 
            this.textBox_31_25.Location = new System.Drawing.Point(90, 363);
            this.textBox_31_25.Name = "textBox_31_25";
            this.textBox_31_25.ReadOnly = true;
            this.textBox_31_25.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_25.TabIndex = 40;
            this.textBox_31_25.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_26
            // 
            this.textBox_31_26.Location = new System.Drawing.Point(176, 363);
            this.textBox_31_26.Name = "textBox_31_26";
            this.textBox_31_26.ReadOnly = true;
            this.textBox_31_26.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_26.TabIndex = 39;
            this.textBox_31_26.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_27
            // 
            this.textBox_31_27.Location = new System.Drawing.Point(262, 363);
            this.textBox_31_27.Name = "textBox_31_27";
            this.textBox_31_27.ReadOnly = true;
            this.textBox_31_27.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_27.TabIndex = 38;
            this.textBox_31_27.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_28
            // 
            this.textBox_31_28.Location = new System.Drawing.Point(348, 363);
            this.textBox_31_28.Name = "textBox_31_28";
            this.textBox_31_28.ReadOnly = true;
            this.textBox_31_28.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_28.TabIndex = 37;
            this.textBox_31_28.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_29
            // 
            this.textBox_31_29.Location = new System.Drawing.Point(90, 403);
            this.textBox_31_29.Name = "textBox_31_29";
            this.textBox_31_29.ReadOnly = true;
            this.textBox_31_29.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_29.TabIndex = 36;
            this.textBox_31_29.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_30
            // 
            this.textBox_31_30.Location = new System.Drawing.Point(176, 403);
            this.textBox_31_30.Name = "textBox_31_30";
            this.textBox_31_30.ReadOnly = true;
            this.textBox_31_30.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_30.TabIndex = 35;
            this.textBox_31_30.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_31
            // 
            this.textBox_31_31.Location = new System.Drawing.Point(262, 403);
            this.textBox_31_31.Name = "textBox_31_31";
            this.textBox_31_31.ReadOnly = true;
            this.textBox_31_31.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_31.TabIndex = 34;
            this.textBox_31_31.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label28
            // 
            this.label28.AutoSize = true;
            this.label28.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label28.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label28.Location = new System.Drawing.Point(50, 248);
            this.label28.Name = "label28";
            this.label28.Size = new System.Drawing.Size(110, 24);
            this.label28.TabIndex = 32;
            this.label28.Text = "两个探头";
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label20.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label20.Location = new System.Drawing.Point(35, 148);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(94, 24);
            this.label20.TabIndex = 31;
            this.label20.Text = "3号探头";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label19.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label19.Location = new System.Drawing.Point(35, 113);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(94, 24);
            this.label19.TabIndex = 30;
            this.label19.Text = "2号探头";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label18.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label18.Location = new System.Drawing.Point(35, 78);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(94, 24);
            this.label18.TabIndex = 29;
            this.label18.Text = "1号探头";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label17.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label17.Location = new System.Drawing.Point(375, 58);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(46, 24);
            this.label17.TabIndex = 28;
            this.label17.Text = "SUM";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label16.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label16.Location = new System.Drawing.Point(295, 58);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(34, 24);
            this.label16.TabIndex = 27;
            this.label16.Text = "Mz";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label15.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label15.Location = new System.Drawing.Point(210, 58);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(34, 24);
            this.label15.TabIndex = 26;
            this.label15.Text = "My";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label14.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label14.Location = new System.Drawing.Point(125, 58);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(34, 24);
            this.label14.TabIndex = 25;
            this.label14.Text = "Mx";
            // 
            // textBox_31_16
            // 
            this.textBox_31_16.Location = new System.Drawing.Point(348, 185);
            this.textBox_31_16.Name = "textBox_31_16";
            this.textBox_31_16.ReadOnly = true;
            this.textBox_31_16.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_16.TabIndex = 24;
            this.textBox_31_16.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_15
            // 
            this.textBox_31_15.Location = new System.Drawing.Point(262, 185);
            this.textBox_31_15.Name = "textBox_31_15";
            this.textBox_31_15.ReadOnly = true;
            this.textBox_31_15.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_15.TabIndex = 23;
            this.textBox_31_15.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_14
            // 
            this.textBox_31_14.Location = new System.Drawing.Point(176, 185);
            this.textBox_31_14.Name = "textBox_31_14";
            this.textBox_31_14.ReadOnly = true;
            this.textBox_31_14.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_14.TabIndex = 22;
            this.textBox_31_14.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_13
            // 
            this.textBox_31_13.Location = new System.Drawing.Point(90, 185);
            this.textBox_31_13.Name = "textBox_31_13";
            this.textBox_31_13.ReadOnly = true;
            this.textBox_31_13.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_13.TabIndex = 21;
            this.textBox_31_13.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_12
            // 
            this.textBox_31_12.Location = new System.Drawing.Point(348, 145);
            this.textBox_31_12.Name = "textBox_31_12";
            this.textBox_31_12.ReadOnly = true;
            this.textBox_31_12.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_12.TabIndex = 20;
            this.textBox_31_12.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_11
            // 
            this.textBox_31_11.Location = new System.Drawing.Point(262, 145);
            this.textBox_31_11.Name = "textBox_31_11";
            this.textBox_31_11.ReadOnly = true;
            this.textBox_31_11.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_11.TabIndex = 19;
            this.textBox_31_11.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_10
            // 
            this.textBox_31_10.Location = new System.Drawing.Point(176, 145);
            this.textBox_31_10.Name = "textBox_31_10";
            this.textBox_31_10.ReadOnly = true;
            this.textBox_31_10.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_10.TabIndex = 18;
            this.textBox_31_10.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_9
            // 
            this.textBox_31_9.Location = new System.Drawing.Point(90, 145);
            this.textBox_31_9.Name = "textBox_31_9";
            this.textBox_31_9.ReadOnly = true;
            this.textBox_31_9.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_9.TabIndex = 17;
            this.textBox_31_9.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_5
            // 
            this.textBox_31_5.Location = new System.Drawing.Point(90, 110);
            this.textBox_31_5.Name = "textBox_31_5";
            this.textBox_31_5.ReadOnly = true;
            this.textBox_31_5.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_5.TabIndex = 16;
            this.textBox_31_5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_6
            // 
            this.textBox_31_6.Location = new System.Drawing.Point(176, 110);
            this.textBox_31_6.Name = "textBox_31_6";
            this.textBox_31_6.ReadOnly = true;
            this.textBox_31_6.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_6.TabIndex = 15;
            this.textBox_31_6.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_7
            // 
            this.textBox_31_7.Location = new System.Drawing.Point(262, 110);
            this.textBox_31_7.Name = "textBox_31_7";
            this.textBox_31_7.ReadOnly = true;
            this.textBox_31_7.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_7.TabIndex = 14;
            this.textBox_31_7.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_8
            // 
            this.textBox_31_8.Location = new System.Drawing.Point(348, 110);
            this.textBox_31_8.Name = "textBox_31_8";
            this.textBox_31_8.ReadOnly = true;
            this.textBox_31_8.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_8.TabIndex = 13;
            this.textBox_31_8.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_3
            // 
            this.textBox_31_3.Location = new System.Drawing.Point(262, 75);
            this.textBox_31_3.Name = "textBox_31_3";
            this.textBox_31_3.ReadOnly = true;
            this.textBox_31_3.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_3.TabIndex = 12;
            this.textBox_31_3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_4
            // 
            this.textBox_31_4.Location = new System.Drawing.Point(348, 75);
            this.textBox_31_4.Name = "textBox_31_4";
            this.textBox_31_4.ReadOnly = true;
            this.textBox_31_4.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_4.TabIndex = 11;
            this.textBox_31_4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_2
            // 
            this.textBox_31_2.Location = new System.Drawing.Point(176, 75);
            this.textBox_31_2.Name = "textBox_31_2";
            this.textBox_31_2.ReadOnly = true;
            this.textBox_31_2.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_2.TabIndex = 10;
            this.textBox_31_2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox_31_1
            // 
            this.textBox_31_1.Location = new System.Drawing.Point(90, 75);
            this.textBox_31_1.Name = "textBox_31_1";
            this.textBox_31_1.ReadOnly = true;
            this.textBox_31_1.Size = new System.Drawing.Size(80, 35);
            this.textBox_31_1.TabIndex = 9;
            this.textBox_31_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label13.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label13.Location = new System.Drawing.Point(50, 30);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(110, 24);
            this.label13.TabIndex = 2;
            this.label13.Text = "一个探头";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.button_31_3);
            this.groupBox3.Controls.Add(this.label12);
            this.groupBox3.Controls.Add(this.comboBox_31_7);
            this.groupBox3.Controls.Add(this.label11);
            this.groupBox3.Controls.Add(this.comboBox_31_6);
            this.groupBox3.Controls.Add(this.label10);
            this.groupBox3.Controls.Add(this.comboBox_31_5);
            this.groupBox3.Controls.Add(this.label9);
            this.groupBox3.Controls.Add(this.checkBox_31_1);
            this.groupBox3.Controls.Add(this.radioButton_31_2);
            this.groupBox3.Controls.Add(this.label8);
            this.groupBox3.Controls.Add(this.radioButton_31_1);
            this.groupBox3.Location = new System.Drawing.Point(6, 278);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(400, 150);
            this.groupBox3.TabIndex = 2;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "设置计算参数";
            // 
            // button_31_3
            // 
            this.button_31_3.Location = new System.Drawing.Point(297, 41);
            this.button_31_3.Name = "button_31_3";
            this.button_31_3.Size = new System.Drawing.Size(80, 23);
            this.button_31_3.TabIndex = 15;
            this.button_31_3.Text = "计算总磁矩";
            this.button_31_3.UseVisualStyleBackColor = true;
            this.button_31_3.Click += new System.EventHandler(this.button_31_3_Click);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label12.Location = new System.Drawing.Point(147, 119);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(26, 28);
            this.label12.TabIndex = 11;
            this.label12.Text = "Z";
            // 
            // comboBox_31_7
            // 
            this.comboBox_31_7.FormattingEnabled = true;
            this.comboBox_31_7.Items.AddRange(new object[] {
            "X",
            "Y",
            "Z",
            "-X",
            "-Y",
            "-Z"});
            this.comboBox_31_7.Location = new System.Drawing.Point(165, 115);
            this.comboBox_31_7.Name = "comboBox_31_7";
            this.comboBox_31_7.Size = new System.Drawing.Size(40, 32);
            this.comboBox_31_7.TabIndex = 10;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label11.Location = new System.Drawing.Point(147, 88);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(26, 28);
            this.label11.TabIndex = 9;
            this.label11.Text = "Y";
            // 
            // comboBox_31_6
            // 
            this.comboBox_31_6.FormattingEnabled = true;
            this.comboBox_31_6.Items.AddRange(new object[] {
            "X",
            "Y",
            "Z",
            "-X",
            "-Y",
            "-Z"});
            this.comboBox_31_6.Location = new System.Drawing.Point(165, 84);
            this.comboBox_31_6.Name = "comboBox_31_6";
            this.comboBox_31_6.Size = new System.Drawing.Size(40, 32);
            this.comboBox_31_6.TabIndex = 8;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label10.Location = new System.Drawing.Point(147, 57);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(26, 28);
            this.label10.TabIndex = 7;
            this.label10.Text = "X";
            // 
            // comboBox_31_5
            // 
            this.comboBox_31_5.FormattingEnabled = true;
            this.comboBox_31_5.Items.AddRange(new object[] {
            "X",
            "Y",
            "Z",
            "-X",
            "-Y",
            "-Z"});
            this.comboBox_31_5.Location = new System.Drawing.Point(165, 53);
            this.comboBox_31_5.Name = "comboBox_31_5";
            this.comboBox_31_5.Size = new System.Drawing.Size(40, 32);
            this.comboBox_31_5.TabIndex = 6;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(150, 30);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(106, 24);
            this.label9.TabIndex = 5;
            this.label9.Text = "坐标关系";
            // 
            // checkBox_31_1
            // 
            this.checkBox_31_1.AutoSize = true;
            this.checkBox_31_1.CheckAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.checkBox_31_1.Location = new System.Drawing.Point(14, 30);
            this.checkBox_31_1.Name = "checkBox_31_1";
            this.checkBox_31_1.Padding = new System.Windows.Forms.Padding(0, 0, 0, 5);
            this.checkBox_31_1.Size = new System.Drawing.Size(110, 84);
            this.checkBox_31_1.TabIndex = 4;
            this.checkBox_31_1.Text = "监测干扰\r\n\r\n";
            this.checkBox_31_1.UseVisualStyleBackColor = true;
            // 
            // radioButton_31_2
            // 
            this.radioButton_31_2.AutoSize = true;
            this.radioButton_31_2.Location = new System.Drawing.Point(83, 88);
            this.radioButton_31_2.Name = "radioButton_31_2";
            this.radioButton_31_2.Size = new System.Drawing.Size(89, 28);
            this.radioButton_31_2.TabIndex = 3;
            this.radioButton_31_2.TabStop = true;
            this.radioButton_31_2.Text = "20度";
            this.radioButton_31_2.UseVisualStyleBackColor = true;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(77, 30);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(106, 24);
            this.label8.TabIndex = 2;
            this.label8.Text = "旋转角度";
            // 
            // radioButton_31_1
            // 
            this.radioButton_31_1.AutoSize = true;
            this.radioButton_31_1.Location = new System.Drawing.Point(83, 57);
            this.radioButton_31_1.Name = "radioButton_31_1";
            this.radioButton_31_1.Size = new System.Drawing.Size(89, 28);
            this.radioButton_31_1.TabIndex = 0;
            this.radioButton_31_1.TabStop = true;
            this.radioButton_31_1.Text = "10度";
            this.radioButton_31_1.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.button_31_2);
            this.groupBox2.Controls.Add(this.textBox_31_r3);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.textBox_31_r2);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.textBox_31_r1);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.comboBox_31_4);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.comboBox_31_3);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.comboBox_31_2);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.comboBox_31_1);
            this.groupBox2.Location = new System.Drawing.Point(6, 142);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(400, 130);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "选择探头序号并设置探头位置";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(294, 104);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(34, 24);
            this.label7.TabIndex = 13;
            this.label7.Text = "r3";
            // 
            // button_31_2
            // 
            this.button_31_2.Location = new System.Drawing.Point(296, 20);
            this.button_31_2.Name = "button_31_2";
            this.button_31_2.Size = new System.Drawing.Size(80, 23);
            this.button_31_2.TabIndex = 14;
            this.button_31_2.Text = "保存参数";
            this.button_31_2.UseVisualStyleBackColor = true;
            this.button_31_2.Click += new System.EventHandler(this.button_31_2_Click);
            // 
            // textBox_31_r3
            // 
            this.textBox_31_r3.Location = new System.Drawing.Point(317, 101);
            this.textBox_31_r3.Name = "textBox_31_r3";
            this.textBox_31_r3.Size = new System.Drawing.Size(60, 35);
            this.textBox_31_r3.TabIndex = 12;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(171, 104);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(34, 24);
            this.label6.TabIndex = 11;
            this.label6.Text = "r2";
            // 
            // textBox_31_r2
            // 
            this.textBox_31_r2.Location = new System.Drawing.Point(194, 101);
            this.textBox_31_r2.Name = "textBox_31_r2";
            this.textBox_31_r2.Size = new System.Drawing.Size(60, 35);
            this.textBox_31_r2.TabIndex = 10;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(47, 104);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(34, 24);
            this.label5.TabIndex = 9;
            this.label5.Text = "r1";
            // 
            // textBox_31_r1
            // 
            this.textBox_31_r1.Location = new System.Drawing.Point(70, 101);
            this.textBox_31_r1.Name = "textBox_31_r1";
            this.textBox_31_r1.Size = new System.Drawing.Size(60, 35);
            this.textBox_31_r1.TabIndex = 8;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(276, 69);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(70, 24);
            this.label4.TabIndex = 7;
            this.label4.Text = "探头3";
            // 
            // comboBox_31_4
            // 
            this.comboBox_31_4.FormattingEnabled = true;
            this.comboBox_31_4.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5"});
            this.comboBox_31_4.Location = new System.Drawing.Point(317, 66);
            this.comboBox_31_4.Name = "comboBox_31_4";
            this.comboBox_31_4.Size = new System.Drawing.Size(60, 32);
            this.comboBox_31_4.TabIndex = 6;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(153, 69);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(70, 24);
            this.label3.TabIndex = 5;
            this.label3.Text = "探头2";
            // 
            // comboBox_31_3
            // 
            this.comboBox_31_3.FormattingEnabled = true;
            this.comboBox_31_3.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5"});
            this.comboBox_31_3.Location = new System.Drawing.Point(194, 66);
            this.comboBox_31_3.Name = "comboBox_31_3";
            this.comboBox_31_3.Size = new System.Drawing.Size(60, 32);
            this.comboBox_31_3.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(29, 69);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(70, 24);
            this.label2.TabIndex = 3;
            this.label2.Text = "探头1";
            // 
            // comboBox_31_2
            // 
            this.comboBox_31_2.FormattingEnabled = true;
            this.comboBox_31_2.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5"});
            this.comboBox_31_2.Location = new System.Drawing.Point(70, 66);
            this.comboBox_31_2.Name = "comboBox_31_2";
            this.comboBox_31_2.Size = new System.Drawing.Size(60, 32);
            this.comboBox_31_2.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 33);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(106, 24);
            this.label1.TabIndex = 1;
            this.label1.Text = "干扰探头";
            // 
            // comboBox_31_1
            // 
            this.comboBox_31_1.FormattingEnabled = true;
            this.comboBox_31_1.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5"});
            this.comboBox_31_1.Location = new System.Drawing.Point(70, 30);
            this.comboBox_31_1.Name = "comboBox_31_1";
            this.comboBox_31_1.Size = new System.Drawing.Size(60, 32);
            this.comboBox_31_1.TabIndex = 0;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.textBox_30_1);
            this.groupBox1.Controls.Add(this.comboBox_30_1);
            this.groupBox1.Controls.Add(this.radioButton_30_2);
            this.groupBox1.Controls.Add(this.radioButton_30_1);
            this.groupBox1.Controls.Add(this.button_31_1);
            this.groupBox1.Location = new System.Drawing.Point(6, 6);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(400, 130);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "提取数据文件";
            // 
            // textBox_30_1
            // 
            this.textBox_30_1.Location = new System.Drawing.Point(70, 80);
            this.textBox_30_1.Name = "textBox_30_1";
            this.textBox_30_1.ReadOnly = true;
            this.textBox_30_1.Size = new System.Drawing.Size(207, 35);
            this.textBox_30_1.TabIndex = 49;
            this.textBox_30_1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // comboBox_30_1
            // 
            this.comboBox_30_1.FormattingEnabled = true;
            this.comboBox_30_1.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5"});
            this.comboBox_30_1.Location = new System.Drawing.Point(70, 36);
            this.comboBox_30_1.Name = "comboBox_30_1";
            this.comboBox_30_1.Size = new System.Drawing.Size(207, 32);
            this.comboBox_30_1.TabIndex = 19;
            this.comboBox_30_1.SelectedIndexChanged += new System.EventHandler(this.comboBox_30_1_SelectedIndexChanged);
            // 
            // radioButton_30_2
            // 
            this.radioButton_30_2.AutoSize = true;
            this.radioButton_30_2.Checked = true;
            this.radioButton_30_2.Location = new System.Drawing.Point(24, 83);
            this.radioButton_30_2.Name = "radioButton_30_2";
            this.radioButton_30_2.Size = new System.Drawing.Size(27, 26);
            this.radioButton_30_2.TabIndex = 18;
            this.radioButton_30_2.TabStop = true;
            this.radioButton_30_2.UseVisualStyleBackColor = true;
            // 
            // radioButton_30_1
            // 
            this.radioButton_30_1.AutoSize = true;
            this.radioButton_30_1.Location = new System.Drawing.Point(24, 39);
            this.radioButton_30_1.Name = "radioButton_30_1";
            this.radioButton_30_1.Size = new System.Drawing.Size(27, 26);
            this.radioButton_30_1.TabIndex = 17;
            this.radioButton_30_1.TabStop = true;
            this.radioButton_30_1.UseVisualStyleBackColor = true;
            this.radioButton_30_1.CheckedChanged += new System.EventHandler(this.radioButton_30_1_CheckedChanged);
            // 
            // button_31_1
            // 
            this.button_31_1.Location = new System.Drawing.Point(297, 78);
            this.button_31_1.Name = "button_31_1";
            this.button_31_1.Size = new System.Drawing.Size(80, 23);
            this.button_31_1.TabIndex = 16;
            this.button_31_1.Text = "读取数据";
            this.button_31_1.UseVisualStyleBackColor = true;
            this.button_31_1.Click += new System.EventHandler(this.button_31_1_Click);
            // 
            // tabPage8
            // 
            this.tabPage8.Controls.Add(this.groupBox15);
            this.tabPage8.Controls.Add(this.groupBox14);
            this.tabPage8.Controls.Add(this.groupBox13);
            this.tabPage8.Controls.Add(this.groupBox12);
            this.tabPage8.Controls.Add(this.groupBox9);
            this.tabPage8.Controls.Add(this.groupBox10);
            this.tabPage8.Controls.Add(this.groupBox11);
            this.tabPage8.ImageIndex = 9;
            this.tabPage8.Location = new System.Drawing.Point(84, 4);
            this.tabPage8.Name = "tabPage8";
            this.tabPage8.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage8.Size = new System.Drawing.Size(1012, 752);
            this.tabPage8.TabIndex = 1;
            this.tabPage8.UseVisualStyleBackColor = true;
            // 
            // groupBox15
            // 
            this.groupBox15.Controls.Add(this.label99);
            this.groupBox15.Controls.Add(this.textBox40);
            this.groupBox15.Controls.Add(this.label100);
            this.groupBox15.Controls.Add(this.textBox41);
            this.groupBox15.Controls.Add(this.label101);
            this.groupBox15.Controls.Add(this.textBox42);
            this.groupBox15.Location = new System.Drawing.Point(462, 635);
            this.groupBox15.Name = "groupBox15";
            this.groupBox15.Size = new System.Drawing.Size(522, 100);
            this.groupBox15.TabIndex = 9;
            this.groupBox15.TabStop = false;
            this.groupBox15.Text = "剩磁矩";
            // 
            // label99
            // 
            this.label99.AutoSize = true;
            this.label99.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label99.Location = new System.Drawing.Point(319, 37);
            this.label99.Name = "label99";
            this.label99.Size = new System.Drawing.Size(22, 24);
            this.label99.TabIndex = 44;
            this.label99.Text = "Z";
            // 
            // textBox40
            // 
            this.textBox40.Location = new System.Drawing.Point(342, 34);
            this.textBox40.Name = "textBox40";
            this.textBox40.ReadOnly = true;
            this.textBox40.Size = new System.Drawing.Size(60, 35);
            this.textBox40.TabIndex = 43;
            // 
            // label100
            // 
            this.label100.AutoSize = true;
            this.label100.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label100.Location = new System.Drawing.Point(220, 37);
            this.label100.Name = "label100";
            this.label100.Size = new System.Drawing.Size(22, 24);
            this.label100.TabIndex = 42;
            this.label100.Text = "Y";
            // 
            // textBox41
            // 
            this.textBox41.Location = new System.Drawing.Point(243, 34);
            this.textBox41.Name = "textBox41";
            this.textBox41.ReadOnly = true;
            this.textBox41.Size = new System.Drawing.Size(60, 35);
            this.textBox41.TabIndex = 41;
            // 
            // label101
            // 
            this.label101.AutoSize = true;
            this.label101.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label101.Location = new System.Drawing.Point(122, 37);
            this.label101.Name = "label101";
            this.label101.Size = new System.Drawing.Size(22, 24);
            this.label101.TabIndex = 40;
            this.label101.Text = "X";
            // 
            // textBox42
            // 
            this.textBox42.Location = new System.Drawing.Point(145, 34);
            this.textBox42.Name = "textBox42";
            this.textBox42.ReadOnly = true;
            this.textBox42.Size = new System.Drawing.Size(60, 35);
            this.textBox42.TabIndex = 39;
            // 
            // groupBox14
            // 
            this.groupBox14.Controls.Add(this.radioButton_32_9);
            this.groupBox14.Controls.Add(this.radioButton_32_8);
            this.groupBox14.Controls.Add(this.label96);
            this.groupBox14.Controls.Add(this.textBox37);
            this.groupBox14.Controls.Add(this.label97);
            this.groupBox14.Controls.Add(this.textBox38);
            this.groupBox14.Controls.Add(this.label98);
            this.groupBox14.Controls.Add(this.textBox39);
            this.groupBox14.Controls.Add(this.label95);
            this.groupBox14.Controls.Add(this.label109);
            this.groupBox14.Controls.Add(this.label110);
            this.groupBox14.Controls.Add(this.label111);
            this.groupBox14.Controls.Add(this.label112);
            this.groupBox14.Controls.Add(this.label113);
            this.groupBox14.Controls.Add(this.label114);
            this.groupBox14.Controls.Add(this.label115);
            this.groupBox14.Controls.Add(this.textBox61);
            this.groupBox14.Controls.Add(this.textBox62);
            this.groupBox14.Controls.Add(this.textBox63);
            this.groupBox14.Controls.Add(this.textBox64);
            this.groupBox14.Controls.Add(this.textBox65);
            this.groupBox14.Controls.Add(this.textBox66);
            this.groupBox14.Controls.Add(this.textBox67);
            this.groupBox14.Controls.Add(this.textBox68);
            this.groupBox14.Controls.Add(this.textBox69);
            this.groupBox14.Controls.Add(this.textBox70);
            this.groupBox14.Controls.Add(this.textBox71);
            this.groupBox14.Controls.Add(this.textBox72);
            this.groupBox14.Location = new System.Drawing.Point(462, 411);
            this.groupBox14.Name = "groupBox14";
            this.groupBox14.Size = new System.Drawing.Size(522, 219);
            this.groupBox14.TabIndex = 8;
            this.groupBox14.TabStop = false;
            this.groupBox14.Text = "感磁矩";
            // 
            // radioButton_32_9
            // 
            this.radioButton_32_9.AutoSize = true;
            this.radioButton_32_9.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.radioButton_32_9.Location = new System.Drawing.Point(208, 188);
            this.radioButton_32_9.Name = "radioButton_32_9";
            this.radioButton_32_9.Size = new System.Drawing.Size(185, 28);
            this.radioButton_32_9.TabIndex = 47;
            this.radioButton_32_9.TabStop = true;
            this.radioButton_32_9.Text = "显示参考结果";
            this.radioButton_32_9.UseVisualStyleBackColor = true;
            // 
            // radioButton_32_8
            // 
            this.radioButton_32_8.AutoSize = true;
            this.radioButton_32_8.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.radioButton_32_8.Location = new System.Drawing.Point(50, 188);
            this.radioButton_32_8.Name = "radioButton_32_8";
            this.radioButton_32_8.Size = new System.Drawing.Size(185, 28);
            this.radioButton_32_8.TabIndex = 46;
            this.radioButton_32_8.TabStop = true;
            this.radioButton_32_8.Text = "显示默认结果";
            this.radioButton_32_8.UseVisualStyleBackColor = true;
            // 
            // label96
            // 
            this.label96.AutoSize = true;
            this.label96.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label96.Location = new System.Drawing.Point(319, 29);
            this.label96.Name = "label96";
            this.label96.Size = new System.Drawing.Size(22, 24);
            this.label96.TabIndex = 38;
            this.label96.Text = "Z";
            // 
            // textBox37
            // 
            this.textBox37.Location = new System.Drawing.Point(342, 26);
            this.textBox37.Name = "textBox37";
            this.textBox37.ReadOnly = true;
            this.textBox37.Size = new System.Drawing.Size(60, 35);
            this.textBox37.TabIndex = 37;
            // 
            // label97
            // 
            this.label97.AutoSize = true;
            this.label97.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label97.Location = new System.Drawing.Point(220, 29);
            this.label97.Name = "label97";
            this.label97.Size = new System.Drawing.Size(22, 24);
            this.label97.TabIndex = 36;
            this.label97.Text = "Y";
            // 
            // textBox38
            // 
            this.textBox38.Location = new System.Drawing.Point(243, 26);
            this.textBox38.Name = "textBox38";
            this.textBox38.ReadOnly = true;
            this.textBox38.Size = new System.Drawing.Size(60, 35);
            this.textBox38.TabIndex = 35;
            // 
            // label98
            // 
            this.label98.AutoSize = true;
            this.label98.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label98.Location = new System.Drawing.Point(122, 29);
            this.label98.Name = "label98";
            this.label98.Size = new System.Drawing.Size(22, 24);
            this.label98.TabIndex = 34;
            this.label98.Text = "X";
            // 
            // textBox39
            // 
            this.textBox39.Location = new System.Drawing.Point(145, 26);
            this.textBox39.Name = "textBox39";
            this.textBox39.ReadOnly = true;
            this.textBox39.Size = new System.Drawing.Size(60, 35);
            this.textBox39.TabIndex = 33;
            // 
            // label95
            // 
            this.label95.AutoSize = true;
            this.label95.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label95.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label95.Location = new System.Drawing.Point(30, 29);
            this.label95.Name = "label95";
            this.label95.Size = new System.Drawing.Size(106, 24);
            this.label95.TabIndex = 32;
            this.label95.Text = "地磁场值";
            // 
            // label109
            // 
            this.label109.AutoSize = true;
            this.label109.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label109.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label109.Location = new System.Drawing.Point(30, 155);
            this.label109.Name = "label109";
            this.label109.Size = new System.Drawing.Size(94, 24);
            this.label109.TabIndex = 31;
            this.label109.Text = "3号探头";
            // 
            // label110
            // 
            this.label110.AutoSize = true;
            this.label110.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label110.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label110.Location = new System.Drawing.Point(30, 120);
            this.label110.Name = "label110";
            this.label110.Size = new System.Drawing.Size(94, 24);
            this.label110.TabIndex = 30;
            this.label110.Text = "2号探头";
            // 
            // label111
            // 
            this.label111.AutoSize = true;
            this.label111.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label111.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label111.Location = new System.Drawing.Point(30, 85);
            this.label111.Name = "label111";
            this.label111.Size = new System.Drawing.Size(94, 24);
            this.label111.TabIndex = 29;
            this.label111.Text = "1号探头";
            // 
            // label112
            // 
            this.label112.AutoSize = true;
            this.label112.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label112.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label112.Location = new System.Drawing.Point(393, 65);
            this.label112.Name = "label112";
            this.label112.Size = new System.Drawing.Size(34, 24);
            this.label112.TabIndex = 28;
            this.label112.Text = "Zy";
            // 
            // label113
            // 
            this.label113.AutoSize = true;
            this.label113.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label113.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label113.Location = new System.Drawing.Point(313, 65);
            this.label113.Name = "label113";
            this.label113.Size = new System.Drawing.Size(22, 24);
            this.label113.TabIndex = 27;
            this.label113.Text = "Y";
            // 
            // label114
            // 
            this.label114.AutoSize = true;
            this.label114.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label114.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label114.Location = new System.Drawing.Point(228, 65);
            this.label114.Name = "label114";
            this.label114.Size = new System.Drawing.Size(34, 24);
            this.label114.TabIndex = 26;
            this.label114.Text = "Zx";
            // 
            // label115
            // 
            this.label115.AutoSize = true;
            this.label115.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label115.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label115.Location = new System.Drawing.Point(143, 65);
            this.label115.Name = "label115";
            this.label115.Size = new System.Drawing.Size(22, 24);
            this.label115.TabIndex = 25;
            this.label115.Text = "X";
            // 
            // textBox61
            // 
            this.textBox61.Location = new System.Drawing.Point(366, 152);
            this.textBox61.Name = "textBox61";
            this.textBox61.ReadOnly = true;
            this.textBox61.Size = new System.Drawing.Size(80, 35);
            this.textBox61.TabIndex = 20;
            this.textBox61.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox62
            // 
            this.textBox62.Location = new System.Drawing.Point(280, 152);
            this.textBox62.Name = "textBox62";
            this.textBox62.ReadOnly = true;
            this.textBox62.Size = new System.Drawing.Size(80, 35);
            this.textBox62.TabIndex = 19;
            this.textBox62.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox63
            // 
            this.textBox63.Location = new System.Drawing.Point(194, 152);
            this.textBox63.Name = "textBox63";
            this.textBox63.ReadOnly = true;
            this.textBox63.Size = new System.Drawing.Size(80, 35);
            this.textBox63.TabIndex = 18;
            this.textBox63.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox64
            // 
            this.textBox64.Location = new System.Drawing.Point(108, 152);
            this.textBox64.Name = "textBox64";
            this.textBox64.ReadOnly = true;
            this.textBox64.Size = new System.Drawing.Size(80, 35);
            this.textBox64.TabIndex = 17;
            this.textBox64.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox65
            // 
            this.textBox65.Location = new System.Drawing.Point(108, 117);
            this.textBox65.Name = "textBox65";
            this.textBox65.ReadOnly = true;
            this.textBox65.Size = new System.Drawing.Size(80, 35);
            this.textBox65.TabIndex = 16;
            this.textBox65.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox66
            // 
            this.textBox66.Location = new System.Drawing.Point(194, 117);
            this.textBox66.Name = "textBox66";
            this.textBox66.ReadOnly = true;
            this.textBox66.Size = new System.Drawing.Size(80, 35);
            this.textBox66.TabIndex = 15;
            this.textBox66.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox67
            // 
            this.textBox67.Location = new System.Drawing.Point(280, 117);
            this.textBox67.Name = "textBox67";
            this.textBox67.ReadOnly = true;
            this.textBox67.Size = new System.Drawing.Size(80, 35);
            this.textBox67.TabIndex = 14;
            this.textBox67.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox68
            // 
            this.textBox68.Location = new System.Drawing.Point(366, 117);
            this.textBox68.Name = "textBox68";
            this.textBox68.ReadOnly = true;
            this.textBox68.Size = new System.Drawing.Size(80, 35);
            this.textBox68.TabIndex = 13;
            this.textBox68.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox69
            // 
            this.textBox69.Location = new System.Drawing.Point(280, 82);
            this.textBox69.Name = "textBox69";
            this.textBox69.ReadOnly = true;
            this.textBox69.Size = new System.Drawing.Size(80, 35);
            this.textBox69.TabIndex = 12;
            this.textBox69.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox70
            // 
            this.textBox70.Location = new System.Drawing.Point(366, 82);
            this.textBox70.Name = "textBox70";
            this.textBox70.ReadOnly = true;
            this.textBox70.Size = new System.Drawing.Size(80, 35);
            this.textBox70.TabIndex = 11;
            this.textBox70.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox71
            // 
            this.textBox71.Location = new System.Drawing.Point(194, 82);
            this.textBox71.Name = "textBox71";
            this.textBox71.ReadOnly = true;
            this.textBox71.Size = new System.Drawing.Size(80, 35);
            this.textBox71.TabIndex = 10;
            this.textBox71.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox72
            // 
            this.textBox72.Location = new System.Drawing.Point(108, 82);
            this.textBox72.Name = "textBox72";
            this.textBox72.ReadOnly = true;
            this.textBox72.Size = new System.Drawing.Size(80, 35);
            this.textBox72.TabIndex = 9;
            this.textBox72.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // groupBox13
            // 
            this.groupBox13.Controls.Add(this.label73);
            this.groupBox13.Controls.Add(this.label74);
            this.groupBox13.Controls.Add(this.label75);
            this.groupBox13.Controls.Add(this.label76);
            this.groupBox13.Controls.Add(this.label77);
            this.groupBox13.Controls.Add(this.textBox1);
            this.groupBox13.Controls.Add(this.textBox2);
            this.groupBox13.Controls.Add(this.textBox3);
            this.groupBox13.Controls.Add(this.textBox4);
            this.groupBox13.Controls.Add(this.label78);
            this.groupBox13.Controls.Add(this.textBox5);
            this.groupBox13.Controls.Add(this.label79);
            this.groupBox13.Controls.Add(this.label80);
            this.groupBox13.Controls.Add(this.label81);
            this.groupBox13.Controls.Add(this.label82);
            this.groupBox13.Controls.Add(this.label83);
            this.groupBox13.Controls.Add(this.label84);
            this.groupBox13.Controls.Add(this.label85);
            this.groupBox13.Controls.Add(this.textBox6);
            this.groupBox13.Controls.Add(this.textBox7);
            this.groupBox13.Controls.Add(this.textBox8);
            this.groupBox13.Controls.Add(this.textBox9);
            this.groupBox13.Controls.Add(this.textBox10);
            this.groupBox13.Controls.Add(this.textBox11);
            this.groupBox13.Controls.Add(this.textBox12);
            this.groupBox13.Controls.Add(this.textBox13);
            this.groupBox13.Controls.Add(this.textBox14);
            this.groupBox13.Controls.Add(this.textBox15);
            this.groupBox13.Controls.Add(this.textBox16);
            this.groupBox13.Controls.Add(this.textBox17);
            this.groupBox13.Controls.Add(this.textBox18);
            this.groupBox13.Controls.Add(this.textBox19);
            this.groupBox13.Controls.Add(this.textBox20);
            this.groupBox13.Controls.Add(this.label86);
            this.groupBox13.Controls.Add(this.label87);
            this.groupBox13.Controls.Add(this.label88);
            this.groupBox13.Controls.Add(this.label89);
            this.groupBox13.Controls.Add(this.label90);
            this.groupBox13.Controls.Add(this.label91);
            this.groupBox13.Controls.Add(this.label92);
            this.groupBox13.Controls.Add(this.label93);
            this.groupBox13.Controls.Add(this.textBox21);
            this.groupBox13.Controls.Add(this.textBox22);
            this.groupBox13.Controls.Add(this.textBox23);
            this.groupBox13.Controls.Add(this.textBox24);
            this.groupBox13.Controls.Add(this.textBox25);
            this.groupBox13.Controls.Add(this.textBox26);
            this.groupBox13.Controls.Add(this.textBox27);
            this.groupBox13.Controls.Add(this.textBox28);
            this.groupBox13.Controls.Add(this.textBox29);
            this.groupBox13.Controls.Add(this.textBox30);
            this.groupBox13.Controls.Add(this.textBox31);
            this.groupBox13.Controls.Add(this.textBox32);
            this.groupBox13.Controls.Add(this.textBox33);
            this.groupBox13.Controls.Add(this.textBox34);
            this.groupBox13.Controls.Add(this.textBox35);
            this.groupBox13.Controls.Add(this.textBox36);
            this.groupBox13.Controls.Add(this.label94);
            this.groupBox13.Location = new System.Drawing.Point(462, 6);
            this.groupBox13.Name = "groupBox13";
            this.groupBox13.Size = new System.Drawing.Size(522, 399);
            this.groupBox13.TabIndex = 7;
            this.groupBox13.TabStop = false;
            this.groupBox13.Text = "总磁矩";
            // 
            // label73
            // 
            this.label73.AutoSize = true;
            this.label73.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label73.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label73.Location = new System.Drawing.Point(24, 371);
            this.label73.Name = "label73";
            this.label73.Size = new System.Drawing.Size(118, 24);
            this.label73.TabIndex = 66;
            this.label73.Text = "123号探头";
            // 
            // label74
            // 
            this.label74.AutoSize = true;
            this.label74.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label74.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label74.Location = new System.Drawing.Point(393, 351);
            this.label74.Name = "label74";
            this.label74.Size = new System.Drawing.Size(46, 24);
            this.label74.TabIndex = 65;
            this.label74.Text = "SUM";
            // 
            // label75
            // 
            this.label75.AutoSize = true;
            this.label75.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label75.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label75.Location = new System.Drawing.Point(313, 351);
            this.label75.Name = "label75";
            this.label75.Size = new System.Drawing.Size(34, 24);
            this.label75.TabIndex = 64;
            this.label75.Text = "Mz";
            // 
            // label76
            // 
            this.label76.AutoSize = true;
            this.label76.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label76.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label76.Location = new System.Drawing.Point(228, 351);
            this.label76.Name = "label76";
            this.label76.Size = new System.Drawing.Size(34, 24);
            this.label76.TabIndex = 63;
            this.label76.Text = "My";
            // 
            // label77
            // 
            this.label77.AutoSize = true;
            this.label77.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label77.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label77.Location = new System.Drawing.Point(143, 351);
            this.label77.Name = "label77";
            this.label77.Size = new System.Drawing.Size(34, 24);
            this.label77.TabIndex = 62;
            this.label77.Text = "Mx";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(108, 368);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(80, 35);
            this.textBox1.TabIndex = 61;
            this.textBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(194, 368);
            this.textBox2.Name = "textBox2";
            this.textBox2.ReadOnly = true;
            this.textBox2.Size = new System.Drawing.Size(80, 35);
            this.textBox2.TabIndex = 60;
            this.textBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(280, 368);
            this.textBox3.Name = "textBox3";
            this.textBox3.ReadOnly = true;
            this.textBox3.Size = new System.Drawing.Size(80, 35);
            this.textBox3.TabIndex = 59;
            this.textBox3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(366, 368);
            this.textBox4.Name = "textBox4";
            this.textBox4.ReadOnly = true;
            this.textBox4.Size = new System.Drawing.Size(80, 35);
            this.textBox4.TabIndex = 58;
            this.textBox4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label78
            // 
            this.label78.AutoSize = true;
            this.label78.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label78.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label78.Location = new System.Drawing.Point(16, 348);
            this.label78.Name = "label78";
            this.label78.Size = new System.Drawing.Size(110, 24);
            this.label78.TabIndex = 57;
            this.label78.Text = "三个探头";
            // 
            // textBox5
            // 
            this.textBox5.Location = new System.Drawing.Point(366, 317);
            this.textBox5.Name = "textBox5";
            this.textBox5.ReadOnly = true;
            this.textBox5.Size = new System.Drawing.Size(80, 35);
            this.textBox5.TabIndex = 56;
            this.textBox5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label79
            // 
            this.label79.AutoSize = true;
            this.label79.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label79.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label79.Location = new System.Drawing.Point(30, 280);
            this.label79.Name = "label79";
            this.label79.Size = new System.Drawing.Size(106, 24);
            this.label79.TabIndex = 55;
            this.label79.Text = "13号探头";
            // 
            // label80
            // 
            this.label80.AutoSize = true;
            this.label80.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label80.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label80.Location = new System.Drawing.Point(30, 245);
            this.label80.Name = "label80";
            this.label80.Size = new System.Drawing.Size(106, 24);
            this.label80.TabIndex = 54;
            this.label80.Text = "23号探头";
            // 
            // label81
            // 
            this.label81.AutoSize = true;
            this.label81.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label81.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label81.Location = new System.Drawing.Point(30, 210);
            this.label81.Name = "label81";
            this.label81.Size = new System.Drawing.Size(106, 24);
            this.label81.TabIndex = 53;
            this.label81.Text = "12号探头";
            // 
            // label82
            // 
            this.label82.AutoSize = true;
            this.label82.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label82.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label82.Location = new System.Drawing.Point(393, 190);
            this.label82.Name = "label82";
            this.label82.Size = new System.Drawing.Size(46, 24);
            this.label82.TabIndex = 52;
            this.label82.Text = "SUM";
            // 
            // label83
            // 
            this.label83.AutoSize = true;
            this.label83.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label83.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label83.Location = new System.Drawing.Point(313, 190);
            this.label83.Name = "label83";
            this.label83.Size = new System.Drawing.Size(34, 24);
            this.label83.TabIndex = 51;
            this.label83.Text = "Mz";
            // 
            // label84
            // 
            this.label84.AutoSize = true;
            this.label84.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label84.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label84.Location = new System.Drawing.Point(228, 190);
            this.label84.Name = "label84";
            this.label84.Size = new System.Drawing.Size(34, 24);
            this.label84.TabIndex = 50;
            this.label84.Text = "My";
            // 
            // label85
            // 
            this.label85.AutoSize = true;
            this.label85.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label85.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label85.Location = new System.Drawing.Point(143, 190);
            this.label85.Name = "label85";
            this.label85.Size = new System.Drawing.Size(34, 24);
            this.label85.TabIndex = 49;
            this.label85.Text = "Mx";
            // 
            // textBox6
            // 
            this.textBox6.Location = new System.Drawing.Point(108, 207);
            this.textBox6.Name = "textBox6";
            this.textBox6.ReadOnly = true;
            this.textBox6.Size = new System.Drawing.Size(80, 35);
            this.textBox6.TabIndex = 48;
            this.textBox6.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox7
            // 
            this.textBox7.Location = new System.Drawing.Point(194, 207);
            this.textBox7.Name = "textBox7";
            this.textBox7.ReadOnly = true;
            this.textBox7.Size = new System.Drawing.Size(80, 35);
            this.textBox7.TabIndex = 47;
            this.textBox7.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox8
            // 
            this.textBox8.Location = new System.Drawing.Point(280, 207);
            this.textBox8.Name = "textBox8";
            this.textBox8.ReadOnly = true;
            this.textBox8.Size = new System.Drawing.Size(80, 35);
            this.textBox8.TabIndex = 46;
            this.textBox8.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox9
            // 
            this.textBox9.Location = new System.Drawing.Point(366, 207);
            this.textBox9.Name = "textBox9";
            this.textBox9.ReadOnly = true;
            this.textBox9.Size = new System.Drawing.Size(80, 35);
            this.textBox9.TabIndex = 45;
            this.textBox9.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox10
            // 
            this.textBox10.Location = new System.Drawing.Point(108, 239);
            this.textBox10.Name = "textBox10";
            this.textBox10.ReadOnly = true;
            this.textBox10.Size = new System.Drawing.Size(80, 35);
            this.textBox10.TabIndex = 44;
            this.textBox10.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox11
            // 
            this.textBox11.Location = new System.Drawing.Point(194, 239);
            this.textBox11.Name = "textBox11";
            this.textBox11.ReadOnly = true;
            this.textBox11.Size = new System.Drawing.Size(80, 35);
            this.textBox11.TabIndex = 43;
            this.textBox11.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox12
            // 
            this.textBox12.Location = new System.Drawing.Point(280, 239);
            this.textBox12.Name = "textBox12";
            this.textBox12.ReadOnly = true;
            this.textBox12.Size = new System.Drawing.Size(80, 35);
            this.textBox12.TabIndex = 42;
            this.textBox12.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox13
            // 
            this.textBox13.Location = new System.Drawing.Point(366, 239);
            this.textBox13.Name = "textBox13";
            this.textBox13.ReadOnly = true;
            this.textBox13.Size = new System.Drawing.Size(80, 35);
            this.textBox13.TabIndex = 41;
            this.textBox13.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox14
            // 
            this.textBox14.Location = new System.Drawing.Point(108, 277);
            this.textBox14.Name = "textBox14";
            this.textBox14.ReadOnly = true;
            this.textBox14.Size = new System.Drawing.Size(80, 35);
            this.textBox14.TabIndex = 40;
            this.textBox14.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox15
            // 
            this.textBox15.Location = new System.Drawing.Point(194, 277);
            this.textBox15.Name = "textBox15";
            this.textBox15.ReadOnly = true;
            this.textBox15.Size = new System.Drawing.Size(80, 35);
            this.textBox15.TabIndex = 39;
            this.textBox15.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox16
            // 
            this.textBox16.Location = new System.Drawing.Point(280, 277);
            this.textBox16.Name = "textBox16";
            this.textBox16.ReadOnly = true;
            this.textBox16.Size = new System.Drawing.Size(80, 35);
            this.textBox16.TabIndex = 38;
            this.textBox16.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox17
            // 
            this.textBox17.Location = new System.Drawing.Point(366, 277);
            this.textBox17.Name = "textBox17";
            this.textBox17.ReadOnly = true;
            this.textBox17.Size = new System.Drawing.Size(80, 35);
            this.textBox17.TabIndex = 37;
            this.textBox17.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox18
            // 
            this.textBox18.Location = new System.Drawing.Point(108, 317);
            this.textBox18.Name = "textBox18";
            this.textBox18.ReadOnly = true;
            this.textBox18.Size = new System.Drawing.Size(80, 35);
            this.textBox18.TabIndex = 36;
            this.textBox18.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox19
            // 
            this.textBox19.Location = new System.Drawing.Point(194, 317);
            this.textBox19.Name = "textBox19";
            this.textBox19.ReadOnly = true;
            this.textBox19.Size = new System.Drawing.Size(80, 35);
            this.textBox19.TabIndex = 35;
            this.textBox19.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox20
            // 
            this.textBox20.Location = new System.Drawing.Point(280, 317);
            this.textBox20.Name = "textBox20";
            this.textBox20.ReadOnly = true;
            this.textBox20.Size = new System.Drawing.Size(80, 35);
            this.textBox20.TabIndex = 34;
            this.textBox20.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label86
            // 
            this.label86.AutoSize = true;
            this.label86.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label86.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label86.Location = new System.Drawing.Point(16, 185);
            this.label86.Name = "label86";
            this.label86.Size = new System.Drawing.Size(110, 24);
            this.label86.TabIndex = 32;
            this.label86.Text = "两个探头";
            // 
            // label87
            // 
            this.label87.AutoSize = true;
            this.label87.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label87.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label87.Location = new System.Drawing.Point(30, 122);
            this.label87.Name = "label87";
            this.label87.Size = new System.Drawing.Size(94, 24);
            this.label87.TabIndex = 31;
            this.label87.Text = "3号探头";
            // 
            // label88
            // 
            this.label88.AutoSize = true;
            this.label88.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label88.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label88.Location = new System.Drawing.Point(30, 87);
            this.label88.Name = "label88";
            this.label88.Size = new System.Drawing.Size(94, 24);
            this.label88.TabIndex = 30;
            this.label88.Text = "2号探头";
            // 
            // label89
            // 
            this.label89.AutoSize = true;
            this.label89.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label89.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label89.Location = new System.Drawing.Point(30, 52);
            this.label89.Name = "label89";
            this.label89.Size = new System.Drawing.Size(94, 24);
            this.label89.TabIndex = 29;
            this.label89.Text = "1号探头";
            // 
            // label90
            // 
            this.label90.AutoSize = true;
            this.label90.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label90.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label90.Location = new System.Drawing.Point(393, 32);
            this.label90.Name = "label90";
            this.label90.Size = new System.Drawing.Size(46, 24);
            this.label90.TabIndex = 28;
            this.label90.Text = "SUM";
            // 
            // label91
            // 
            this.label91.AutoSize = true;
            this.label91.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label91.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label91.Location = new System.Drawing.Point(313, 32);
            this.label91.Name = "label91";
            this.label91.Size = new System.Drawing.Size(34, 24);
            this.label91.TabIndex = 27;
            this.label91.Text = "Mz";
            // 
            // label92
            // 
            this.label92.AutoSize = true;
            this.label92.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label92.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label92.Location = new System.Drawing.Point(228, 32);
            this.label92.Name = "label92";
            this.label92.Size = new System.Drawing.Size(34, 24);
            this.label92.TabIndex = 26;
            this.label92.Text = "My";
            // 
            // label93
            // 
            this.label93.AutoSize = true;
            this.label93.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label93.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label93.Location = new System.Drawing.Point(143, 32);
            this.label93.Name = "label93";
            this.label93.Size = new System.Drawing.Size(34, 24);
            this.label93.TabIndex = 25;
            this.label93.Text = "Mx";
            // 
            // textBox21
            // 
            this.textBox21.Location = new System.Drawing.Point(366, 159);
            this.textBox21.Name = "textBox21";
            this.textBox21.ReadOnly = true;
            this.textBox21.Size = new System.Drawing.Size(80, 35);
            this.textBox21.TabIndex = 24;
            this.textBox21.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox22
            // 
            this.textBox22.Location = new System.Drawing.Point(280, 159);
            this.textBox22.Name = "textBox22";
            this.textBox22.ReadOnly = true;
            this.textBox22.Size = new System.Drawing.Size(80, 35);
            this.textBox22.TabIndex = 23;
            this.textBox22.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox23
            // 
            this.textBox23.Location = new System.Drawing.Point(194, 159);
            this.textBox23.Name = "textBox23";
            this.textBox23.ReadOnly = true;
            this.textBox23.Size = new System.Drawing.Size(80, 35);
            this.textBox23.TabIndex = 22;
            this.textBox23.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox24
            // 
            this.textBox24.Location = new System.Drawing.Point(108, 159);
            this.textBox24.Name = "textBox24";
            this.textBox24.ReadOnly = true;
            this.textBox24.Size = new System.Drawing.Size(80, 35);
            this.textBox24.TabIndex = 21;
            this.textBox24.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox25
            // 
            this.textBox25.Location = new System.Drawing.Point(366, 119);
            this.textBox25.Name = "textBox25";
            this.textBox25.ReadOnly = true;
            this.textBox25.Size = new System.Drawing.Size(80, 35);
            this.textBox25.TabIndex = 20;
            this.textBox25.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox26
            // 
            this.textBox26.Location = new System.Drawing.Point(280, 119);
            this.textBox26.Name = "textBox26";
            this.textBox26.ReadOnly = true;
            this.textBox26.Size = new System.Drawing.Size(80, 35);
            this.textBox26.TabIndex = 19;
            this.textBox26.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox27
            // 
            this.textBox27.Location = new System.Drawing.Point(194, 119);
            this.textBox27.Name = "textBox27";
            this.textBox27.ReadOnly = true;
            this.textBox27.Size = new System.Drawing.Size(80, 35);
            this.textBox27.TabIndex = 18;
            this.textBox27.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox28
            // 
            this.textBox28.Location = new System.Drawing.Point(108, 119);
            this.textBox28.Name = "textBox28";
            this.textBox28.ReadOnly = true;
            this.textBox28.Size = new System.Drawing.Size(80, 35);
            this.textBox28.TabIndex = 17;
            this.textBox28.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox29
            // 
            this.textBox29.Location = new System.Drawing.Point(108, 84);
            this.textBox29.Name = "textBox29";
            this.textBox29.ReadOnly = true;
            this.textBox29.Size = new System.Drawing.Size(80, 35);
            this.textBox29.TabIndex = 16;
            this.textBox29.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox30
            // 
            this.textBox30.Location = new System.Drawing.Point(194, 84);
            this.textBox30.Name = "textBox30";
            this.textBox30.ReadOnly = true;
            this.textBox30.Size = new System.Drawing.Size(80, 35);
            this.textBox30.TabIndex = 15;
            this.textBox30.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox31
            // 
            this.textBox31.Location = new System.Drawing.Point(280, 84);
            this.textBox31.Name = "textBox31";
            this.textBox31.ReadOnly = true;
            this.textBox31.Size = new System.Drawing.Size(80, 35);
            this.textBox31.TabIndex = 14;
            this.textBox31.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox32
            // 
            this.textBox32.Location = new System.Drawing.Point(366, 84);
            this.textBox32.Name = "textBox32";
            this.textBox32.ReadOnly = true;
            this.textBox32.Size = new System.Drawing.Size(80, 35);
            this.textBox32.TabIndex = 13;
            this.textBox32.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox33
            // 
            this.textBox33.Location = new System.Drawing.Point(280, 49);
            this.textBox33.Name = "textBox33";
            this.textBox33.ReadOnly = true;
            this.textBox33.Size = new System.Drawing.Size(80, 35);
            this.textBox33.TabIndex = 12;
            this.textBox33.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox34
            // 
            this.textBox34.Location = new System.Drawing.Point(366, 49);
            this.textBox34.Name = "textBox34";
            this.textBox34.ReadOnly = true;
            this.textBox34.Size = new System.Drawing.Size(80, 35);
            this.textBox34.TabIndex = 11;
            this.textBox34.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox35
            // 
            this.textBox35.Location = new System.Drawing.Point(194, 49);
            this.textBox35.Name = "textBox35";
            this.textBox35.ReadOnly = true;
            this.textBox35.Size = new System.Drawing.Size(80, 35);
            this.textBox35.TabIndex = 10;
            this.textBox35.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox36
            // 
            this.textBox36.Location = new System.Drawing.Point(108, 49);
            this.textBox36.Name = "textBox36";
            this.textBox36.ReadOnly = true;
            this.textBox36.Size = new System.Drawing.Size(80, 35);
            this.textBox36.TabIndex = 9;
            this.textBox36.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label94
            // 
            this.label94.AutoSize = true;
            this.label94.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label94.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label94.Location = new System.Drawing.Point(20, 25);
            this.label94.Name = "label94";
            this.label94.Size = new System.Drawing.Size(110, 24);
            this.label94.TabIndex = 2;
            this.label94.Text = "一个探头";
            // 
            // groupBox12
            // 
            this.groupBox12.Controls.Add(this.radioButton_32_7);
            this.groupBox12.Controls.Add(this.radioButton_32_6);
            this.groupBox12.Controls.Add(this.radioButton_32_5);
            this.groupBox12.Controls.Add(this.checkBox_32_4);
            this.groupBox12.Controls.Add(this.checkBox_32_3);
            this.groupBox12.Controls.Add(this.checkBox_32_2);
            this.groupBox12.Controls.Add(this.button_32_5);
            this.groupBox12.Controls.Add(this.label72);
            this.groupBox12.Controls.Add(this.label71);
            this.groupBox12.Location = new System.Drawing.Point(6, 574);
            this.groupBox12.Name = "groupBox12";
            this.groupBox12.Size = new System.Drawing.Size(450, 161);
            this.groupBox12.TabIndex = 6;
            this.groupBox12.TabStop = false;
            this.groupBox12.Text = "设置剩磁矩参数";
            // 
            // radioButton_32_7
            // 
            this.radioButton_32_7.AutoSize = true;
            this.radioButton_32_7.Location = new System.Drawing.Point(297, 75);
            this.radioButton_32_7.Name = "radioButton_32_7";
            this.radioButton_32_7.Size = new System.Drawing.Size(101, 28);
            this.radioButton_32_7.TabIndex = 45;
            this.radioButton_32_7.TabStop = true;
            this.radioButton_32_7.Text = "探头3";
            this.radioButton_32_7.UseVisualStyleBackColor = true;
            // 
            // radioButton_32_6
            // 
            this.radioButton_32_6.AutoSize = true;
            this.radioButton_32_6.Location = new System.Drawing.Point(211, 75);
            this.radioButton_32_6.Name = "radioButton_32_6";
            this.radioButton_32_6.Size = new System.Drawing.Size(101, 28);
            this.radioButton_32_6.TabIndex = 44;
            this.radioButton_32_6.TabStop = true;
            this.radioButton_32_6.Text = "探头2";
            this.radioButton_32_6.UseVisualStyleBackColor = true;
            // 
            // radioButton_32_5
            // 
            this.radioButton_32_5.AutoSize = true;
            this.radioButton_32_5.Location = new System.Drawing.Point(124, 75);
            this.radioButton_32_5.Name = "radioButton_32_5";
            this.radioButton_32_5.Size = new System.Drawing.Size(101, 28);
            this.radioButton_32_5.TabIndex = 43;
            this.radioButton_32_5.TabStop = true;
            this.radioButton_32_5.Text = "探头1";
            this.radioButton_32_5.UseVisualStyleBackColor = true;
            // 
            // checkBox_32_4
            // 
            this.checkBox_32_4.AutoSize = true;
            this.checkBox_32_4.Location = new System.Drawing.Point(297, 40);
            this.checkBox_32_4.Name = "checkBox_32_4";
            this.checkBox_32_4.Size = new System.Drawing.Size(102, 28);
            this.checkBox_32_4.TabIndex = 42;
            this.checkBox_32_4.Text = "探头3";
            this.checkBox_32_4.UseVisualStyleBackColor = true;
            // 
            // checkBox_32_3
            // 
            this.checkBox_32_3.AutoSize = true;
            this.checkBox_32_3.Location = new System.Drawing.Point(210, 40);
            this.checkBox_32_3.Name = "checkBox_32_3";
            this.checkBox_32_3.Size = new System.Drawing.Size(102, 28);
            this.checkBox_32_3.TabIndex = 41;
            this.checkBox_32_3.Text = "探头2";
            this.checkBox_32_3.UseVisualStyleBackColor = true;
            // 
            // checkBox_32_2
            // 
            this.checkBox_32_2.AutoSize = true;
            this.checkBox_32_2.Location = new System.Drawing.Point(125, 40);
            this.checkBox_32_2.Name = "checkBox_32_2";
            this.checkBox_32_2.Size = new System.Drawing.Size(102, 28);
            this.checkBox_32_2.TabIndex = 40;
            this.checkBox_32_2.Text = "探头1";
            this.checkBox_32_2.UseVisualStyleBackColor = true;
            // 
            // button_32_5
            // 
            this.button_32_5.Location = new System.Drawing.Point(352, 116);
            this.button_32_5.Name = "button_32_5";
            this.button_32_5.Size = new System.Drawing.Size(80, 23);
            this.button_32_5.TabIndex = 39;
            this.button_32_5.Text = "计算剩磁矩";
            this.button_32_5.UseVisualStyleBackColor = true;
            // 
            // label72
            // 
            this.label72.AutoSize = true;
            this.label72.Location = new System.Drawing.Point(20, 77);
            this.label72.Name = "label72";
            this.label72.Size = new System.Drawing.Size(130, 24);
            this.label72.TabIndex = 17;
            this.label72.Text = "感磁矩结果";
            // 
            // label71
            // 
            this.label71.AutoSize = true;
            this.label71.Location = new System.Drawing.Point(20, 41);
            this.label71.Name = "label71";
            this.label71.Size = new System.Drawing.Size(106, 24);
            this.label71.TabIndex = 16;
            this.label71.Text = "磁矩探头";
            // 
            // groupBox9
            // 
            this.groupBox9.Controls.Add(this.button_32_4);
            this.groupBox9.Controls.Add(this.radioButton_32_4);
            this.groupBox9.Controls.Add(this.radioButton_32_3);
            this.groupBox9.Controls.Add(this.label70);
            this.groupBox9.Controls.Add(this.label67);
            this.groupBox9.Controls.Add(this.comboBox_32_13);
            this.groupBox9.Controls.Add(this.label68);
            this.groupBox9.Controls.Add(this.comboBox_32_12);
            this.groupBox9.Controls.Add(this.label69);
            this.groupBox9.Controls.Add(this.comboBox_32_11);
            this.groupBox9.Controls.Add(this.label64);
            this.groupBox9.Controls.Add(this.comboBox_32_10);
            this.groupBox9.Controls.Add(this.label65);
            this.groupBox9.Controls.Add(this.comboBox_32_9);
            this.groupBox9.Controls.Add(this.label66);
            this.groupBox9.Controls.Add(this.comboBox_32_8);
            this.groupBox9.Controls.Add(this.label63);
            this.groupBox9.Controls.Add(this.label59);
            this.groupBox9.Controls.Add(this.label60);
            this.groupBox9.Controls.Add(this.textBox_32_z1);
            this.groupBox9.Controls.Add(this.label61);
            this.groupBox9.Controls.Add(this.textBox_32_y1);
            this.groupBox9.Controls.Add(this.label62);
            this.groupBox9.Controls.Add(this.textBox_32_x1);
            this.groupBox9.Location = new System.Drawing.Point(6, 318);
            this.groupBox9.Name = "groupBox9";
            this.groupBox9.Size = new System.Drawing.Size(450, 250);
            this.groupBox9.TabIndex = 5;
            this.groupBox9.TabStop = false;
            this.groupBox9.Text = "设置感磁矩参数";
            // 
            // button_32_4
            // 
            this.button_32_4.Location = new System.Drawing.Point(352, 193);
            this.button_32_4.Name = "button_32_4";
            this.button_32_4.Size = new System.Drawing.Size(80, 23);
            this.button_32_4.TabIndex = 38;
            this.button_32_4.Text = "计算感磁矩";
            this.button_32_4.UseVisualStyleBackColor = true;
            // 
            // radioButton_32_4
            // 
            this.radioButton_32_4.AutoSize = true;
            this.radioButton_32_4.Location = new System.Drawing.Point(150, 191);
            this.radioButton_32_4.Name = "radioButton_32_4";
            this.radioButton_32_4.Size = new System.Drawing.Size(89, 28);
            this.radioButton_32_4.TabIndex = 37;
            this.radioButton_32_4.TabStop = true;
            this.radioButton_32_4.Text = "20度";
            this.radioButton_32_4.UseVisualStyleBackColor = true;
            // 
            // radioButton_32_3
            // 
            this.radioButton_32_3.AutoSize = true;
            this.radioButton_32_3.Location = new System.Drawing.Point(88, 191);
            this.radioButton_32_3.Name = "radioButton_32_3";
            this.radioButton_32_3.Size = new System.Drawing.Size(89, 28);
            this.radioButton_32_3.TabIndex = 35;
            this.radioButton_32_3.TabStop = true;
            this.radioButton_32_3.Text = "10度";
            this.radioButton_32_3.UseVisualStyleBackColor = true;
            // 
            // label70
            // 
            this.label70.AutoSize = true;
            this.label70.Location = new System.Drawing.Point(13, 193);
            this.label70.Name = "label70";
            this.label70.Size = new System.Drawing.Size(106, 24);
            this.label70.TabIndex = 36;
            this.label70.Text = "旋转角度";
            // 
            // label67
            // 
            this.label67.AutoSize = true;
            this.label67.Location = new System.Drawing.Point(273, 140);
            this.label67.Name = "label67";
            this.label67.Size = new System.Drawing.Size(70, 24);
            this.label67.TabIndex = 34;
            this.label67.Text = "探头6";
            // 
            // comboBox_32_13
            // 
            this.comboBox_32_13.FormattingEnabled = true;
            this.comboBox_32_13.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5"});
            this.comboBox_32_13.Location = new System.Drawing.Point(308, 137);
            this.comboBox_32_13.Name = "comboBox_32_13";
            this.comboBox_32_13.Size = new System.Drawing.Size(60, 32);
            this.comboBox_32_13.TabIndex = 33;
            // 
            // label68
            // 
            this.label68.AutoSize = true;
            this.label68.Location = new System.Drawing.Point(163, 140);
            this.label68.Name = "label68";
            this.label68.Size = new System.Drawing.Size(70, 24);
            this.label68.TabIndex = 32;
            this.label68.Text = "探头5";
            // 
            // comboBox_32_12
            // 
            this.comboBox_32_12.FormattingEnabled = true;
            this.comboBox_32_12.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5"});
            this.comboBox_32_12.Location = new System.Drawing.Point(198, 137);
            this.comboBox_32_12.Name = "comboBox_32_12";
            this.comboBox_32_12.Size = new System.Drawing.Size(60, 32);
            this.comboBox_32_12.TabIndex = 31;
            // 
            // label69
            // 
            this.label69.AutoSize = true;
            this.label69.Location = new System.Drawing.Point(53, 140);
            this.label69.Name = "label69";
            this.label69.Size = new System.Drawing.Size(70, 24);
            this.label69.TabIndex = 30;
            this.label69.Text = "探头4";
            // 
            // comboBox_32_11
            // 
            this.comboBox_32_11.FormattingEnabled = true;
            this.comboBox_32_11.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5"});
            this.comboBox_32_11.Location = new System.Drawing.Point(88, 137);
            this.comboBox_32_11.Name = "comboBox_32_11";
            this.comboBox_32_11.Size = new System.Drawing.Size(60, 32);
            this.comboBox_32_11.TabIndex = 29;
            // 
            // label64
            // 
            this.label64.AutoSize = true;
            this.label64.Location = new System.Drawing.Point(273, 104);
            this.label64.Name = "label64";
            this.label64.Size = new System.Drawing.Size(70, 24);
            this.label64.TabIndex = 28;
            this.label64.Text = "探头3";
            // 
            // comboBox_32_10
            // 
            this.comboBox_32_10.FormattingEnabled = true;
            this.comboBox_32_10.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5"});
            this.comboBox_32_10.Location = new System.Drawing.Point(308, 101);
            this.comboBox_32_10.Name = "comboBox_32_10";
            this.comboBox_32_10.Size = new System.Drawing.Size(60, 32);
            this.comboBox_32_10.TabIndex = 27;
            // 
            // label65
            // 
            this.label65.AutoSize = true;
            this.label65.Location = new System.Drawing.Point(163, 104);
            this.label65.Name = "label65";
            this.label65.Size = new System.Drawing.Size(70, 24);
            this.label65.TabIndex = 26;
            this.label65.Text = "探头2";
            // 
            // comboBox_32_9
            // 
            this.comboBox_32_9.FormattingEnabled = true;
            this.comboBox_32_9.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5"});
            this.comboBox_32_9.Location = new System.Drawing.Point(198, 101);
            this.comboBox_32_9.Name = "comboBox_32_9";
            this.comboBox_32_9.Size = new System.Drawing.Size(60, 32);
            this.comboBox_32_9.TabIndex = 25;
            // 
            // label66
            // 
            this.label66.AutoSize = true;
            this.label66.Location = new System.Drawing.Point(53, 104);
            this.label66.Name = "label66";
            this.label66.Size = new System.Drawing.Size(70, 24);
            this.label66.TabIndex = 24;
            this.label66.Text = "探头1";
            // 
            // comboBox_32_8
            // 
            this.comboBox_32_8.FormattingEnabled = true;
            this.comboBox_32_8.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5"});
            this.comboBox_32_8.Location = new System.Drawing.Point(88, 101);
            this.comboBox_32_8.Name = "comboBox_32_8";
            this.comboBox_32_8.Size = new System.Drawing.Size(60, 32);
            this.comboBox_32_8.TabIndex = 23;
            // 
            // label63
            // 
            this.label63.AutoSize = true;
            this.label63.Location = new System.Drawing.Point(9, 75);
            this.label63.Name = "label63";
            this.label63.Size = new System.Drawing.Size(154, 24);
            this.label63.TabIndex = 22;
            this.label63.Text = "选择计算探头";
            // 
            // label59
            // 
            this.label59.AutoSize = true;
            this.label59.Location = new System.Drawing.Point(9, 36);
            this.label59.Name = "label59";
            this.label59.Size = new System.Drawing.Size(154, 24);
            this.label59.TabIndex = 21;
            this.label59.Text = "输入地磁场值";
            // 
            // label60
            // 
            this.label60.AutoSize = true;
            this.label60.Location = new System.Drawing.Point(295, 36);
            this.label60.Name = "label60";
            this.label60.Size = new System.Drawing.Size(22, 24);
            this.label60.TabIndex = 20;
            this.label60.Text = "Z";
            // 
            // textBox_32_z1
            // 
            this.textBox_32_z1.Location = new System.Drawing.Point(318, 33);
            this.textBox_32_z1.Name = "textBox_32_z1";
            this.textBox_32_z1.Size = new System.Drawing.Size(60, 35);
            this.textBox_32_z1.TabIndex = 19;
            // 
            // label61
            // 
            this.label61.AutoSize = true;
            this.label61.Location = new System.Drawing.Point(196, 36);
            this.label61.Name = "label61";
            this.label61.Size = new System.Drawing.Size(22, 24);
            this.label61.TabIndex = 18;
            this.label61.Text = "Y";
            // 
            // textBox_32_y1
            // 
            this.textBox_32_y1.Location = new System.Drawing.Point(219, 33);
            this.textBox_32_y1.Name = "textBox_32_y1";
            this.textBox_32_y1.Size = new System.Drawing.Size(60, 35);
            this.textBox_32_y1.TabIndex = 17;
            // 
            // label62
            // 
            this.label62.AutoSize = true;
            this.label62.Location = new System.Drawing.Point(98, 36);
            this.label62.Name = "label62";
            this.label62.Size = new System.Drawing.Size(22, 24);
            this.label62.TabIndex = 16;
            this.label62.Text = "X";
            // 
            // textBox_32_x1
            // 
            this.textBox_32_x1.Location = new System.Drawing.Point(121, 33);
            this.textBox_32_x1.Name = "textBox_32_x1";
            this.textBox_32_x1.Size = new System.Drawing.Size(60, 35);
            this.textBox_32_x1.TabIndex = 15;
            // 
            // groupBox10
            // 
            this.groupBox10.Controls.Add(this.radioButton_32_2);
            this.groupBox10.Controls.Add(this.button_32_3);
            this.groupBox10.Controls.Add(this.radioButton_32_1);
            this.groupBox10.Controls.Add(this.label49);
            this.groupBox10.Controls.Add(this.label45);
            this.groupBox10.Controls.Add(this.checkBox_32_1);
            this.groupBox10.Controls.Add(this.comboBox_32_7);
            this.groupBox10.Controls.Add(this.label58);
            this.groupBox10.Controls.Add(this.label46);
            this.groupBox10.Controls.Add(this.label57);
            this.groupBox10.Controls.Add(this.comboBox_32_6);
            this.groupBox10.Controls.Add(this.label50);
            this.groupBox10.Controls.Add(this.label47);
            this.groupBox10.Controls.Add(this.textBox_32_r3);
            this.groupBox10.Controls.Add(this.comboBox_32_5);
            this.groupBox10.Controls.Add(this.label51);
            this.groupBox10.Controls.Add(this.label48);
            this.groupBox10.Controls.Add(this.textBox_32_r2);
            this.groupBox10.Controls.Add(this.label52);
            this.groupBox10.Controls.Add(this.textBox_32_r1);
            this.groupBox10.Controls.Add(this.label53);
            this.groupBox10.Controls.Add(this.comboBox_32_4);
            this.groupBox10.Controls.Add(this.label54);
            this.groupBox10.Controls.Add(this.comboBox_32_3);
            this.groupBox10.Controls.Add(this.label55);
            this.groupBox10.Controls.Add(this.comboBox_32_2);
            this.groupBox10.Controls.Add(this.label56);
            this.groupBox10.Controls.Add(this.comboBox_32_1);
            this.groupBox10.Location = new System.Drawing.Point(6, 112);
            this.groupBox10.Name = "groupBox10";
            this.groupBox10.Size = new System.Drawing.Size(450, 200);
            this.groupBox10.TabIndex = 4;
            this.groupBox10.TabStop = false;
            this.groupBox10.Text = "设置总磁矩参数";
            // 
            // radioButton_32_2
            // 
            this.radioButton_32_2.AutoSize = true;
            this.radioButton_32_2.Location = new System.Drawing.Point(328, 123);
            this.radioButton_32_2.Name = "radioButton_32_2";
            this.radioButton_32_2.Size = new System.Drawing.Size(89, 28);
            this.radioButton_32_2.TabIndex = 3;
            this.radioButton_32_2.TabStop = true;
            this.radioButton_32_2.Text = "20度";
            this.radioButton_32_2.UseVisualStyleBackColor = true;
            // 
            // button_32_3
            // 
            this.button_32_3.Location = new System.Drawing.Point(352, 163);
            this.button_32_3.Name = "button_32_3";
            this.button_32_3.Size = new System.Drawing.Size(80, 23);
            this.button_32_3.TabIndex = 15;
            this.button_32_3.Text = "计算总磁矩";
            this.button_32_3.UseVisualStyleBackColor = true;
            // 
            // radioButton_32_1
            // 
            this.radioButton_32_1.AutoSize = true;
            this.radioButton_32_1.Location = new System.Drawing.Point(275, 123);
            this.radioButton_32_1.Name = "radioButton_32_1";
            this.radioButton_32_1.Size = new System.Drawing.Size(89, 28);
            this.radioButton_32_1.TabIndex = 0;
            this.radioButton_32_1.TabStop = true;
            this.radioButton_32_1.Text = "10度";
            this.radioButton_32_1.UseVisualStyleBackColor = true;
            // 
            // label49
            // 
            this.label49.AutoSize = true;
            this.label49.Location = new System.Drawing.Point(214, 125);
            this.label49.Name = "label49";
            this.label49.Size = new System.Drawing.Size(106, 24);
            this.label49.TabIndex = 2;
            this.label49.Text = "旋转角度";
            // 
            // label45
            // 
            this.label45.AutoSize = true;
            this.label45.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label45.Location = new System.Drawing.Point(224, 169);
            this.label45.Name = "label45";
            this.label45.Size = new System.Drawing.Size(26, 28);
            this.label45.TabIndex = 11;
            this.label45.Text = "Z";
            // 
            // checkBox_32_1
            // 
            this.checkBox_32_1.AutoSize = true;
            this.checkBox_32_1.Location = new System.Drawing.Point(99, 124);
            this.checkBox_32_1.Name = "checkBox_32_1";
            this.checkBox_32_1.Size = new System.Drawing.Size(138, 28);
            this.checkBox_32_1.TabIndex = 16;
            this.checkBox_32_1.Text = "监测干扰";
            this.checkBox_32_1.UseVisualStyleBackColor = true;
            // 
            // comboBox_32_7
            // 
            this.comboBox_32_7.FormattingEnabled = true;
            this.comboBox_32_7.Items.AddRange(new object[] {
            "X",
            "Y",
            "Z",
            "-X",
            "-Y",
            "-Z"});
            this.comboBox_32_7.Location = new System.Drawing.Point(242, 165);
            this.comboBox_32_7.Name = "comboBox_32_7";
            this.comboBox_32_7.Size = new System.Drawing.Size(40, 32);
            this.comboBox_32_7.TabIndex = 10;
            // 
            // label58
            // 
            this.label58.AutoSize = true;
            this.label58.Location = new System.Drawing.Point(5, 125);
            this.label58.Name = "label58";
            this.label58.Size = new System.Drawing.Size(154, 24);
            this.label58.TabIndex = 15;
            this.label58.Text = "设置计算参数";
            // 
            // label46
            // 
            this.label46.AutoSize = true;
            this.label46.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label46.Location = new System.Drawing.Point(147, 169);
            this.label46.Name = "label46";
            this.label46.Size = new System.Drawing.Size(26, 28);
            this.label46.TabIndex = 9;
            this.label46.Text = "Y";
            // 
            // label57
            // 
            this.label57.AutoSize = true;
            this.label57.Location = new System.Drawing.Point(9, 84);
            this.label57.Name = "label57";
            this.label57.Size = new System.Drawing.Size(154, 24);
            this.label57.TabIndex = 14;
            this.label57.Text = "输入距离参数";
            // 
            // comboBox_32_6
            // 
            this.comboBox_32_6.FormattingEnabled = true;
            this.comboBox_32_6.Items.AddRange(new object[] {
            "X",
            "Y",
            "Z",
            "-X",
            "-Y",
            "-Z"});
            this.comboBox_32_6.Location = new System.Drawing.Point(165, 165);
            this.comboBox_32_6.Name = "comboBox_32_6";
            this.comboBox_32_6.Size = new System.Drawing.Size(40, 32);
            this.comboBox_32_6.TabIndex = 8;
            // 
            // label50
            // 
            this.label50.AutoSize = true;
            this.label50.Location = new System.Drawing.Point(295, 84);
            this.label50.Name = "label50";
            this.label50.Size = new System.Drawing.Size(34, 24);
            this.label50.TabIndex = 13;
            this.label50.Text = "r3";
            // 
            // label47
            // 
            this.label47.AutoSize = true;
            this.label47.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label47.Location = new System.Drawing.Point(71, 169);
            this.label47.Name = "label47";
            this.label47.Size = new System.Drawing.Size(26, 28);
            this.label47.TabIndex = 7;
            this.label47.Text = "X";
            // 
            // textBox_32_r3
            // 
            this.textBox_32_r3.Location = new System.Drawing.Point(318, 81);
            this.textBox_32_r3.Name = "textBox_32_r3";
            this.textBox_32_r3.Size = new System.Drawing.Size(60, 35);
            this.textBox_32_r3.TabIndex = 12;
            // 
            // comboBox_32_5
            // 
            this.comboBox_32_5.FormattingEnabled = true;
            this.comboBox_32_5.Items.AddRange(new object[] {
            "X",
            "Y",
            "Z",
            "-X",
            "-Y",
            "-Z"});
            this.comboBox_32_5.Location = new System.Drawing.Point(89, 165);
            this.comboBox_32_5.Name = "comboBox_32_5";
            this.comboBox_32_5.Size = new System.Drawing.Size(40, 32);
            this.comboBox_32_5.TabIndex = 6;
            // 
            // label51
            // 
            this.label51.AutoSize = true;
            this.label51.Location = new System.Drawing.Point(196, 84);
            this.label51.Name = "label51";
            this.label51.Size = new System.Drawing.Size(34, 24);
            this.label51.TabIndex = 11;
            this.label51.Text = "r2";
            // 
            // label48
            // 
            this.label48.AutoSize = true;
            this.label48.Location = new System.Drawing.Point(6, 170);
            this.label48.Name = "label48";
            this.label48.Size = new System.Drawing.Size(106, 24);
            this.label48.TabIndex = 5;
            this.label48.Text = "坐标关系";
            // 
            // textBox_32_r2
            // 
            this.textBox_32_r2.Location = new System.Drawing.Point(219, 81);
            this.textBox_32_r2.Name = "textBox_32_r2";
            this.textBox_32_r2.Size = new System.Drawing.Size(60, 35);
            this.textBox_32_r2.TabIndex = 10;
            // 
            // label52
            // 
            this.label52.AutoSize = true;
            this.label52.Location = new System.Drawing.Point(98, 84);
            this.label52.Name = "label52";
            this.label52.Size = new System.Drawing.Size(34, 24);
            this.label52.TabIndex = 9;
            this.label52.Text = "r1";
            // 
            // textBox_32_r1
            // 
            this.textBox_32_r1.Location = new System.Drawing.Point(121, 81);
            this.textBox_32_r1.Name = "textBox_32_r1";
            this.textBox_32_r1.Size = new System.Drawing.Size(60, 35);
            this.textBox_32_r1.TabIndex = 8;
            // 
            // label53
            // 
            this.label53.AutoSize = true;
            this.label53.Location = new System.Drawing.Point(350, 33);
            this.label53.Name = "label53";
            this.label53.Size = new System.Drawing.Size(70, 24);
            this.label53.TabIndex = 7;
            this.label53.Text = "探头3";
            // 
            // comboBox_32_4
            // 
            this.comboBox_32_4.FormattingEnabled = true;
            this.comboBox_32_4.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5"});
            this.comboBox_32_4.Location = new System.Drawing.Point(385, 30);
            this.comboBox_32_4.Name = "comboBox_32_4";
            this.comboBox_32_4.Size = new System.Drawing.Size(60, 32);
            this.comboBox_32_4.TabIndex = 6;
            // 
            // label54
            // 
            this.label54.AutoSize = true;
            this.label54.Location = new System.Drawing.Point(240, 33);
            this.label54.Name = "label54";
            this.label54.Size = new System.Drawing.Size(70, 24);
            this.label54.TabIndex = 5;
            this.label54.Text = "探头2";
            // 
            // comboBox_32_3
            // 
            this.comboBox_32_3.FormattingEnabled = true;
            this.comboBox_32_3.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5"});
            this.comboBox_32_3.Location = new System.Drawing.Point(275, 30);
            this.comboBox_32_3.Name = "comboBox_32_3";
            this.comboBox_32_3.Size = new System.Drawing.Size(60, 32);
            this.comboBox_32_3.TabIndex = 4;
            // 
            // label55
            // 
            this.label55.AutoSize = true;
            this.label55.Location = new System.Drawing.Point(130, 33);
            this.label55.Name = "label55";
            this.label55.Size = new System.Drawing.Size(70, 24);
            this.label55.TabIndex = 3;
            this.label55.Text = "探头1";
            // 
            // comboBox_32_2
            // 
            this.comboBox_32_2.FormattingEnabled = true;
            this.comboBox_32_2.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5"});
            this.comboBox_32_2.Location = new System.Drawing.Point(165, 30);
            this.comboBox_32_2.Name = "comboBox_32_2";
            this.comboBox_32_2.Size = new System.Drawing.Size(60, 32);
            this.comboBox_32_2.TabIndex = 2;
            // 
            // label56
            // 
            this.label56.AutoSize = true;
            this.label56.Location = new System.Drawing.Point(5, 33);
            this.label56.Name = "label56";
            this.label56.Size = new System.Drawing.Size(106, 24);
            this.label56.TabIndex = 1;
            this.label56.Text = "干扰探头";
            // 
            // comboBox_32_1
            // 
            this.comboBox_32_1.FormattingEnabled = true;
            this.comboBox_32_1.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5"});
            this.comboBox_32_1.Location = new System.Drawing.Point(60, 30);
            this.comboBox_32_1.Name = "comboBox_32_1";
            this.comboBox_32_1.Size = new System.Drawing.Size(60, 32);
            this.comboBox_32_1.TabIndex = 0;
            // 
            // groupBox11
            // 
            this.groupBox11.Controls.Add(this.button_32_2);
            this.groupBox11.Controls.Add(this.button_32_1);
            this.groupBox11.Location = new System.Drawing.Point(6, 6);
            this.groupBox11.Name = "groupBox11";
            this.groupBox11.Size = new System.Drawing.Size(450, 100);
            this.groupBox11.TabIndex = 3;
            this.groupBox11.TabStop = false;
            this.groupBox11.Text = "提取数据文件";
            // 
            // button_32_2
            // 
            this.button_32_2.Location = new System.Drawing.Point(166, 50);
            this.button_32_2.Name = "button_32_2";
            this.button_32_2.Size = new System.Drawing.Size(80, 23);
            this.button_32_2.TabIndex = 14;
            this.button_32_2.Text = "保存参数";
            this.button_32_2.UseVisualStyleBackColor = true;
            // 
            // button_32_1
            // 
            this.button_32_1.Location = new System.Drawing.Point(60, 50);
            this.button_32_1.Name = "button_32_1";
            this.button_32_1.Size = new System.Drawing.Size(80, 23);
            this.button_32_1.TabIndex = 16;
            this.button_32_1.Text = "读取数据";
            this.button_32_1.UseVisualStyleBackColor = true;
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "171.png");
            this.imageList1.Images.SetKeyName(1, "164.png");
            this.imageList1.Images.SetKeyName(2, "309.png");
            this.imageList1.Images.SetKeyName(3, "大转台.png");
            this.imageList1.Images.SetKeyName(4, "电源控制.png");
            this.imageList1.Images.SetKeyName(5, "数据采集.png");
            this.imageList1.Images.SetKeyName(6, "数据计算.png");
            this.imageList1.Images.SetKeyName(7, "系统配置.png");
            this.imageList1.Images.SetKeyName(8, "小转台.png");
            this.imageList1.Images.SetKeyName(9, "制作图标 (1).png");
            this.imageList1.Images.SetKeyName(10, "制作图标.png");
            this.imageList1.Images.SetKeyName(11, "制作图标 (2).png");
            this.imageList1.Images.SetKeyName(12, "制作图标 (3).png");
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.comboBox_4_1);
            this.tabPage4.Controls.Add(this.label_4_1);
            this.tabPage4.Controls.Add(this.circularGauge2);
            this.tabPage4.ImageIndex = 3;
            this.tabPage4.Location = new System.Drawing.Point(84, 4);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage4.Size = new System.Drawing.Size(1112, 792);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // comboBox_4_1
            // 
            this.comboBox_4_1.FormattingEnabled = true;
            this.comboBox_4_1.Items.AddRange(new object[] {
            "Dev0",
            "Dev1",
            "Dev2",
            "Dev3"});
            this.comboBox_4_1.Location = new System.Drawing.Point(139, 26);
            this.comboBox_4_1.Name = "comboBox_4_1";
            this.comboBox_4_1.Size = new System.Drawing.Size(140, 32);
            this.comboBox_4_1.TabIndex = 4;
            this.comboBox_4_1.Text = "Dev1";
            // 
            // label_4_1
            // 
            this.label_4_1.AutoSize = true;
            this.label_4_1.Location = new System.Drawing.Point(75, 29);
            this.label_4_1.Name = "label_4_1";
            this.label_4_1.Size = new System.Drawing.Size(58, 24);
            this.label_4_1.TabIndex = 3;
            this.label_4_1.Text = "Dev:";
            // 
            // tabPage5
            // 
            this.tabPage5.ImageIndex = 8;
            this.tabPage5.Location = new System.Drawing.Point(84, 4);
            this.tabPage5.Name = "tabPage5";
            this.tabPage5.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage5.Size = new System.Drawing.Size(1112, 792);
            this.tabPage5.TabIndex = 4;
            this.tabPage5.UseVisualStyleBackColor = true;
            // 
            // tabPage6
            // 
            this.tabPage6.Controls.Add(this.tabControl5);
            this.tabPage6.ImageIndex = 7;
            this.tabPage6.Location = new System.Drawing.Point(84, 4);
            this.tabPage6.Name = "tabPage6";
            this.tabPage6.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage6.Size = new System.Drawing.Size(1112, 792);
            this.tabPage6.TabIndex = 5;
            this.tabPage6.UseVisualStyleBackColor = true;
            // 
            // tabControl5
            // 
            this.tabControl5.Alignment = System.Windows.Forms.TabAlignment.Left;
            this.tabControl5.Controls.Add(this.tabPage13);
            this.tabControl5.Controls.Add(this.tabPage14);
            this.tabControl5.ImageList = this.imageList1;
            this.tabControl5.Location = new System.Drawing.Point(6, 6);
            this.tabControl5.Multiline = true;
            this.tabControl5.Name = "tabControl5";
            this.tabControl5.SelectedIndex = 0;
            this.tabControl5.Size = new System.Drawing.Size(1100, 760);
            this.tabControl5.SizeMode = System.Windows.Forms.TabSizeMode.Fixed;
            this.tabControl5.TabIndex = 0;
            // 
            // tabPage13
            // 
            this.tabPage13.Controls.Add(this.groupBox23);
            this.tabPage13.Controls.Add(this.groupBox22);
            this.tabPage13.Controls.Add(this.label44);
            this.tabPage13.Controls.Add(this.comboBox_61_2);
            this.tabPage13.Controls.Add(this.label42);
            this.tabPage13.Controls.Add(this.label41);
            this.tabPage13.Controls.Add(this.label40);
            this.tabPage13.Controls.Add(this.comboBox_61_1);
            this.tabPage13.Controls.Add(this.button_61_3);
            this.tabPage13.Controls.Add(this.button_61_2);
            this.tabPage13.Controls.Add(this.button_61_1);
            this.tabPage13.Controls.Add(this.dataGridView_61_1);
            this.tabPage13.Controls.Add(this.dataGridView_61_2);
            this.tabPage13.ImageIndex = 12;
            this.tabPage13.Location = new System.Drawing.Point(77, 4);
            this.tabPage13.Name = "tabPage13";
            this.tabPage13.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage13.Size = new System.Drawing.Size(1019, 752);
            this.tabPage13.TabIndex = 0;
            this.tabPage13.UseVisualStyleBackColor = true;
            // 
            // groupBox23
            // 
            this.groupBox23.Controls.Add(this.checkBox_61_2);
            this.groupBox23.Controls.Add(this.label148);
            this.groupBox23.Controls.Add(this.comboBox_61_8);
            this.groupBox23.Controls.Add(this.label149);
            this.groupBox23.Controls.Add(this.comboBox_61_7);
            this.groupBox23.Controls.Add(this.label150);
            this.groupBox23.Controls.Add(this.comboBox_61_6);
            this.groupBox23.Location = new System.Drawing.Point(506, 649);
            this.groupBox23.Name = "groupBox23";
            this.groupBox23.Size = new System.Drawing.Size(400, 60);
            this.groupBox23.TabIndex = 13;
            this.groupBox23.TabStop = false;
            this.groupBox23.Text = "CH设备转换";
            // 
            // checkBox_61_2
            // 
            this.checkBox_61_2.AutoSize = true;
            this.checkBox_61_2.Location = new System.Drawing.Point(307, 27);
            this.checkBox_61_2.Name = "checkBox_61_2";
            this.checkBox_61_2.Size = new System.Drawing.Size(138, 28);
            this.checkBox_61_2.TabIndex = 18;
            this.checkBox_61_2.Text = "允许修改";
            this.checkBox_61_2.UseVisualStyleBackColor = true;
            this.checkBox_61_2.CheckedChanged += new System.EventHandler(this.checkBox_61_2_CheckedChanged);
            // 
            // label148
            // 
            this.label148.AutoSize = true;
            this.label148.Location = new System.Drawing.Point(212, 28);
            this.label148.Name = "label148";
            this.label148.Size = new System.Drawing.Size(22, 24);
            this.label148.TabIndex = 16;
            this.label148.Text = "Z";
            // 
            // comboBox_61_8
            // 
            this.comboBox_61_8.Enabled = false;
            this.comboBox_61_8.FormattingEnabled = true;
            this.comboBox_61_8.Items.AddRange(new object[] {
            "X",
            "Y",
            "Z",
            "-X",
            "-Y",
            "-Z"});
            this.comboBox_61_8.Location = new System.Drawing.Point(229, 25);
            this.comboBox_61_8.Name = "comboBox_61_8";
            this.comboBox_61_8.Size = new System.Drawing.Size(50, 32);
            this.comboBox_61_8.TabIndex = 15;
            this.comboBox_61_8.Text = "Z";
            // 
            // label149
            // 
            this.label149.AutoSize = true;
            this.label149.Location = new System.Drawing.Point(117, 28);
            this.label149.Name = "label149";
            this.label149.Size = new System.Drawing.Size(22, 24);
            this.label149.TabIndex = 14;
            this.label149.Text = "Y";
            // 
            // comboBox_61_7
            // 
            this.comboBox_61_7.Enabled = false;
            this.comboBox_61_7.FormattingEnabled = true;
            this.comboBox_61_7.Items.AddRange(new object[] {
            "X",
            "Y",
            "Z",
            "-X",
            "-Y",
            "-Z"});
            this.comboBox_61_7.Location = new System.Drawing.Point(134, 25);
            this.comboBox_61_7.Name = "comboBox_61_7";
            this.comboBox_61_7.Size = new System.Drawing.Size(50, 32);
            this.comboBox_61_7.TabIndex = 13;
            this.comboBox_61_7.Text = "Y";
            // 
            // label150
            // 
            this.label150.AutoSize = true;
            this.label150.Location = new System.Drawing.Point(23, 28);
            this.label150.Name = "label150";
            this.label150.Size = new System.Drawing.Size(22, 24);
            this.label150.TabIndex = 12;
            this.label150.Text = "X";
            // 
            // comboBox_61_6
            // 
            this.comboBox_61_6.Enabled = false;
            this.comboBox_61_6.FormattingEnabled = true;
            this.comboBox_61_6.Items.AddRange(new object[] {
            "X",
            "Y",
            "Z",
            "-X",
            "-Y",
            "-Z"});
            this.comboBox_61_6.Location = new System.Drawing.Point(40, 25);
            this.comboBox_61_6.Name = "comboBox_61_6";
            this.comboBox_61_6.Size = new System.Drawing.Size(50, 32);
            this.comboBox_61_6.TabIndex = 11;
            this.comboBox_61_6.Text = "X";
            // 
            // groupBox22
            // 
            this.groupBox22.Controls.Add(this.checkBox_61_1);
            this.groupBox22.Controls.Add(this.label147);
            this.groupBox22.Controls.Add(this.comboBox_61_5);
            this.groupBox22.Controls.Add(this.label146);
            this.groupBox22.Controls.Add(this.comboBox_61_4);
            this.groupBox22.Controls.Add(this.label145);
            this.groupBox22.Controls.Add(this.comboBox_61_3);
            this.groupBox22.Location = new System.Drawing.Point(506, 572);
            this.groupBox22.Name = "groupBox22";
            this.groupBox22.Size = new System.Drawing.Size(400, 60);
            this.groupBox22.TabIndex = 12;
            this.groupBox22.TabStop = false;
            this.groupBox22.Text = "HS设备转换";
            // 
            // checkBox_61_1
            // 
            this.checkBox_61_1.AutoSize = true;
            this.checkBox_61_1.Location = new System.Drawing.Point(307, 27);
            this.checkBox_61_1.Name = "checkBox_61_1";
            this.checkBox_61_1.Size = new System.Drawing.Size(138, 28);
            this.checkBox_61_1.TabIndex = 17;
            this.checkBox_61_1.Text = "允许修改";
            this.checkBox_61_1.UseVisualStyleBackColor = true;
            this.checkBox_61_1.CheckedChanged += new System.EventHandler(this.checkBox_61_1_CheckedChanged);
            // 
            // label147
            // 
            this.label147.AutoSize = true;
            this.label147.Location = new System.Drawing.Point(212, 28);
            this.label147.Name = "label147";
            this.label147.Size = new System.Drawing.Size(22, 24);
            this.label147.TabIndex = 16;
            this.label147.Text = "Z";
            // 
            // comboBox_61_5
            // 
            this.comboBox_61_5.Enabled = false;
            this.comboBox_61_5.FormattingEnabled = true;
            this.comboBox_61_5.Items.AddRange(new object[] {
            "X",
            "Y",
            "Z",
            "-X",
            "-Y",
            "-Z"});
            this.comboBox_61_5.Location = new System.Drawing.Point(229, 25);
            this.comboBox_61_5.Name = "comboBox_61_5";
            this.comboBox_61_5.Size = new System.Drawing.Size(50, 32);
            this.comboBox_61_5.TabIndex = 15;
            this.comboBox_61_5.Text = "-X";
            // 
            // label146
            // 
            this.label146.AutoSize = true;
            this.label146.Location = new System.Drawing.Point(117, 28);
            this.label146.Name = "label146";
            this.label146.Size = new System.Drawing.Size(22, 24);
            this.label146.TabIndex = 14;
            this.label146.Text = "Y";
            // 
            // comboBox_61_4
            // 
            this.comboBox_61_4.Enabled = false;
            this.comboBox_61_4.FormattingEnabled = true;
            this.comboBox_61_4.Items.AddRange(new object[] {
            "X",
            "Y",
            "Z",
            "-X",
            "-Y",
            "-Z"});
            this.comboBox_61_4.Location = new System.Drawing.Point(134, 25);
            this.comboBox_61_4.Name = "comboBox_61_4";
            this.comboBox_61_4.Size = new System.Drawing.Size(50, 32);
            this.comboBox_61_4.TabIndex = 13;
            this.comboBox_61_4.Text = "Y";
            // 
            // label145
            // 
            this.label145.AutoSize = true;
            this.label145.Location = new System.Drawing.Point(23, 28);
            this.label145.Name = "label145";
            this.label145.Size = new System.Drawing.Size(22, 24);
            this.label145.TabIndex = 12;
            this.label145.Text = "X";
            // 
            // comboBox_61_3
            // 
            this.comboBox_61_3.Enabled = false;
            this.comboBox_61_3.FormattingEnabled = true;
            this.comboBox_61_3.Items.AddRange(new object[] {
            "X",
            "Y",
            "Z",
            "-X",
            "-Y",
            "-Z"});
            this.comboBox_61_3.Location = new System.Drawing.Point(40, 25);
            this.comboBox_61_3.Name = "comboBox_61_3";
            this.comboBox_61_3.Size = new System.Drawing.Size(50, 32);
            this.comboBox_61_3.TabIndex = 11;
            this.comboBox_61_3.Text = "Z";
            // 
            // label44
            // 
            this.label44.AutoSize = true;
            this.label44.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label44.Location = new System.Drawing.Point(747, 465);
            this.label44.Name = "label44";
            this.label44.Size = new System.Drawing.Size(0, 28);
            this.label44.TabIndex = 10;
            // 
            // comboBox_61_2
            // 
            this.comboBox_61_2.FormattingEnabled = true;
            this.comboBox_61_2.Items.AddRange(new object[] {
            "192.168.1.11:6001",
            "192.168.1.12:6002",
            "192.168.1.13:6003",
            "192.168.1.14:6004"});
            this.comboBox_61_2.Location = new System.Drawing.Point(750, 490);
            this.comboBox_61_2.Name = "comboBox_61_2";
            this.comboBox_61_2.Size = new System.Drawing.Size(160, 32);
            this.comboBox_61_2.TabIndex = 9;
            // 
            // label42
            // 
            this.label42.AutoSize = true;
            this.label42.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label42.Location = new System.Drawing.Point(503, 465);
            this.label42.Name = "label42";
            this.label42.Size = new System.Drawing.Size(0, 28);
            this.label42.TabIndex = 8;
            // 
            // label41
            // 
            this.label41.AutoSize = true;
            this.label41.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label41.Location = new System.Drawing.Point(503, 13);
            this.label41.Name = "label41";
            this.label41.Size = new System.Drawing.Size(0, 28);
            this.label41.TabIndex = 7;
            // 
            // label40
            // 
            this.label40.AutoSize = true;
            this.label40.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label40.Location = new System.Drawing.Point(43, 13);
            this.label40.Name = "label40";
            this.label40.Size = new System.Drawing.Size(0, 28);
            this.label40.TabIndex = 6;
            // 
            // comboBox_61_1
            // 
            this.comboBox_61_1.FormattingEnabled = true;
            this.comboBox_61_1.Items.AddRange(new object[] {
            "0.2",
            "0.5",
            "1",
            "10",
            "50",
            "100",
            "200",
            "400",
            "1000",
            "2000"});
            this.comboBox_61_1.Location = new System.Drawing.Point(506, 490);
            this.comboBox_61_1.Name = "comboBox_61_1";
            this.comboBox_61_1.Size = new System.Drawing.Size(121, 32);
            this.comboBox_61_1.TabIndex = 5;
            // 
            // button_61_3
            // 
            this.button_61_3.Location = new System.Drawing.Point(380, 253);
            this.button_61_3.Name = "button_61_3";
            this.button_61_3.Size = new System.Drawing.Size(100, 30);
            this.button_61_3.TabIndex = 4;
            this.button_61_3.Text = "应用";
            this.button_61_3.UseVisualStyleBackColor = true;
            this.button_61_3.Click += new System.EventHandler(this.button_61_3_Click);
            // 
            // button_61_2
            // 
            this.button_61_2.Location = new System.Drawing.Point(216, 696);
            this.button_61_2.Name = "button_61_2";
            this.button_61_2.Size = new System.Drawing.Size(120, 30);
            this.button_61_2.TabIndex = 3;
            this.button_61_2.Text = "全部空选";
            this.button_61_2.UseVisualStyleBackColor = true;
            this.button_61_2.Click += new System.EventHandler(this.button_61_2_Click);
            // 
            // button_61_1
            // 
            this.button_61_1.Location = new System.Drawing.Point(66, 696);
            this.button_61_1.Name = "button_61_1";
            this.button_61_1.Size = new System.Drawing.Size(120, 30);
            this.button_61_1.TabIndex = 2;
            this.button_61_1.Text = "全部选中";
            this.button_61_1.UseVisualStyleBackColor = true;
            this.button_61_1.Click += new System.EventHandler(this.button_61_1_Click);
            // 
            // dataGridView_61_1
            // 
            this.dataGridView_61_1.AllowUserToAddRows = false;
            this.dataGridView_61_1.AllowUserToDeleteRows = false;
            this.dataGridView_61_1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView_61_1.Location = new System.Drawing.Point(506, 40);
            this.dataGridView_61_1.Name = "dataGridView_61_1";
            this.dataGridView_61_1.ReadOnly = true;
            this.dataGridView_61_1.RowHeadersWidth = 82;
            this.dataGridView_61_1.RowTemplate.Height = 23;
            this.dataGridView_61_1.Size = new System.Drawing.Size(450, 400);
            this.dataGridView_61_1.TabIndex = 1;
            // 
            // dataGridView_61_2
            // 
            this.dataGridView_61_2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView_61_2.Location = new System.Drawing.Point(46, 40);
            this.dataGridView_61_2.Name = "dataGridView_61_2";
            this.dataGridView_61_2.RowHeadersWidth = 82;
            this.dataGridView_61_2.RowTemplate.Height = 23;
            this.dataGridView_61_2.Size = new System.Drawing.Size(305, 640);
            this.dataGridView_61_2.TabIndex = 0;
            // 
            // tabPage14
            // 
            this.tabPage14.Controls.Add(this.textBox_62_9_1);
            this.tabPage14.Controls.Add(this.label137);
            this.tabPage14.Controls.Add(this.textBox_62_8_2);
            this.tabPage14.Controls.Add(this.textBox_62_8_1);
            this.tabPage14.Controls.Add(this.label135);
            this.tabPage14.Controls.Add(this.textBox_62_7_2);
            this.tabPage14.Controls.Add(this.textBox_62_7_1);
            this.tabPage14.Controls.Add(this.label134);
            this.tabPage14.Controls.Add(this.textBox_62_6_4);
            this.tabPage14.Controls.Add(this.textBox_62_5_4);
            this.tabPage14.Controls.Add(this.textBox_62_4_4);
            this.tabPage14.Controls.Add(this.textBox_62_3_4);
            this.tabPage14.Controls.Add(this.textBox_62_2_4);
            this.tabPage14.Controls.Add(this.textBox_62_1_4);
            this.tabPage14.Controls.Add(this.textBox_62_6_3);
            this.tabPage14.Controls.Add(this.textBox_62_5_3);
            this.tabPage14.Controls.Add(this.textBox_62_4_3);
            this.tabPage14.Controls.Add(this.textBox_62_3_3);
            this.tabPage14.Controls.Add(this.textBox_62_2_3);
            this.tabPage14.Controls.Add(this.textBox_62_1_3);
            this.tabPage14.Controls.Add(this.textBox_62_6_2);
            this.tabPage14.Controls.Add(this.textBox_62_5_2);
            this.tabPage14.Controls.Add(this.textBox_62_4_2);
            this.tabPage14.Controls.Add(this.textBox_62_3_2);
            this.tabPage14.Controls.Add(this.textBox_62_2_2);
            this.tabPage14.Controls.Add(this.textBox_62_1_2);
            this.tabPage14.Controls.Add(this.textBox_62_6_1);
            this.tabPage14.Controls.Add(this.textBox_62_5_1);
            this.tabPage14.Controls.Add(this.textBox_62_4_1);
            this.tabPage14.Controls.Add(this.textBox_62_3_1);
            this.tabPage14.Controls.Add(this.textBox_62_2_1);
            this.tabPage14.Controls.Add(this.label118);
            this.tabPage14.Controls.Add(this.label117);
            this.tabPage14.Controls.Add(this.label116);
            this.tabPage14.Controls.Add(this.label108);
            this.tabPage14.Controls.Add(this.label107);
            this.tabPage14.Controls.Add(this.label106);
            this.tabPage14.Controls.Add(this.label105);
            this.tabPage14.Controls.Add(this.label104);
            this.tabPage14.Controls.Add(this.label103);
            this.tabPage14.Controls.Add(this.label102);
            this.tabPage14.Controls.Add(this.textBox_62_1_1);
            this.tabPage14.ImageIndex = 11;
            this.tabPage14.Location = new System.Drawing.Point(77, 4);
            this.tabPage14.Name = "tabPage14";
            this.tabPage14.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage14.Size = new System.Drawing.Size(1019, 752);
            this.tabPage14.TabIndex = 1;
            this.tabPage14.UseVisualStyleBackColor = true;
            // 
            // textBox_62_9_1
            // 
            this.textBox_62_9_1.Location = new System.Drawing.Point(620, 440);
            this.textBox_62_9_1.Name = "textBox_62_9_1";
            this.textBox_62_9_1.Size = new System.Drawing.Size(140, 35);
            this.textBox_62_9_1.TabIndex = 43;
            this.textBox_62_9_1.Text = "18.4854";
            // 
            // label137
            // 
            this.label137.AutoSize = true;
            this.label137.Location = new System.Drawing.Point(530, 443);
            this.label137.Name = "label137";
            this.label137.Size = new System.Drawing.Size(130, 24);
            this.label137.TabIndex = 42;
            this.label137.Text = "BI系数mT/A";
            // 
            // textBox_62_8_2
            // 
            this.textBox_62_8_2.Location = new System.Drawing.Point(347, 461);
            this.textBox_62_8_2.Name = "textBox_62_8_2";
            this.textBox_62_8_2.Size = new System.Drawing.Size(100, 35);
            this.textBox_62_8_2.TabIndex = 41;
            this.textBox_62_8_2.Text = "5025";
            // 
            // textBox_62_8_1
            // 
            this.textBox_62_8_1.Location = new System.Drawing.Point(166, 461);
            this.textBox_62_8_1.Name = "textBox_62_8_1";
            this.textBox_62_8_1.Size = new System.Drawing.Size(140, 35);
            this.textBox_62_8_1.TabIndex = 40;
            this.textBox_62_8_1.Text = "192.168.0.10";
            // 
            // label135
            // 
            this.label135.AutoSize = true;
            this.label135.Location = new System.Drawing.Point(89, 464);
            this.label135.Name = "label135";
            this.label135.Size = new System.Drawing.Size(70, 24);
            this.label135.TabIndex = 39;
            this.label135.Text = "电源8";
            // 
            // textBox_62_7_2
            // 
            this.textBox_62_7_2.Location = new System.Drawing.Point(347, 416);
            this.textBox_62_7_2.Name = "textBox_62_7_2";
            this.textBox_62_7_2.Size = new System.Drawing.Size(100, 35);
            this.textBox_62_7_2.TabIndex = 36;
            this.textBox_62_7_2.Text = "5025";
            // 
            // textBox_62_7_1
            // 
            this.textBox_62_7_1.Location = new System.Drawing.Point(166, 416);
            this.textBox_62_7_1.Name = "textBox_62_7_1";
            this.textBox_62_7_1.Size = new System.Drawing.Size(140, 35);
            this.textBox_62_7_1.TabIndex = 35;
            this.textBox_62_7_1.Text = "192.168.123.44";
            // 
            // label134
            // 
            this.label134.AutoSize = true;
            this.label134.Location = new System.Drawing.Point(89, 419);
            this.label134.Name = "label134";
            this.label134.Size = new System.Drawing.Size(70, 24);
            this.label134.TabIndex = 34;
            this.label134.Text = "电源7";
            // 
            // textBox_62_6_4
            // 
            this.textBox_62_6_4.Location = new System.Drawing.Point(671, 347);
            this.textBox_62_6_4.Name = "textBox_62_6_4";
            this.textBox_62_6_4.Size = new System.Drawing.Size(140, 35);
            this.textBox_62_6_4.TabIndex = 33;
            this.textBox_62_6_4.Text = "6";
            // 
            // textBox_62_5_4
            // 
            this.textBox_62_5_4.Location = new System.Drawing.Point(671, 297);
            this.textBox_62_5_4.Name = "textBox_62_5_4";
            this.textBox_62_5_4.Size = new System.Drawing.Size(140, 35);
            this.textBox_62_5_4.TabIndex = 32;
            this.textBox_62_5_4.Text = "5";
            // 
            // textBox_62_4_4
            // 
            this.textBox_62_4_4.Location = new System.Drawing.Point(671, 247);
            this.textBox_62_4_4.Name = "textBox_62_4_4";
            this.textBox_62_4_4.Size = new System.Drawing.Size(140, 35);
            this.textBox_62_4_4.TabIndex = 31;
            this.textBox_62_4_4.Text = "4";
            // 
            // textBox_62_3_4
            // 
            this.textBox_62_3_4.Location = new System.Drawing.Point(671, 197);
            this.textBox_62_3_4.Name = "textBox_62_3_4";
            this.textBox_62_3_4.Size = new System.Drawing.Size(140, 35);
            this.textBox_62_3_4.TabIndex = 30;
            this.textBox_62_3_4.Text = "3";
            // 
            // textBox_62_2_4
            // 
            this.textBox_62_2_4.Location = new System.Drawing.Point(671, 147);
            this.textBox_62_2_4.Name = "textBox_62_2_4";
            this.textBox_62_2_4.Size = new System.Drawing.Size(140, 35);
            this.textBox_62_2_4.TabIndex = 29;
            this.textBox_62_2_4.Text = "2";
            // 
            // textBox_62_1_4
            // 
            this.textBox_62_1_4.Location = new System.Drawing.Point(671, 97);
            this.textBox_62_1_4.Name = "textBox_62_1_4";
            this.textBox_62_1_4.Size = new System.Drawing.Size(140, 35);
            this.textBox_62_1_4.TabIndex = 28;
            this.textBox_62_1_4.Text = "1";
            // 
            // textBox_62_6_3
            // 
            this.textBox_62_6_3.Location = new System.Drawing.Point(491, 347);
            this.textBox_62_6_3.Name = "textBox_62_6_3";
            this.textBox_62_6_3.Size = new System.Drawing.Size(140, 35);
            this.textBox_62_6_3.TabIndex = 27;
            this.textBox_62_6_3.Text = "2000";
            // 
            // textBox_62_5_3
            // 
            this.textBox_62_5_3.Location = new System.Drawing.Point(491, 297);
            this.textBox_62_5_3.Name = "textBox_62_5_3";
            this.textBox_62_5_3.Size = new System.Drawing.Size(140, 35);
            this.textBox_62_5_3.TabIndex = 26;
            this.textBox_62_5_3.Text = "1000";
            // 
            // textBox_62_4_3
            // 
            this.textBox_62_4_3.Location = new System.Drawing.Point(491, 247);
            this.textBox_62_4_3.Name = "textBox_62_4_3";
            this.textBox_62_4_3.Size = new System.Drawing.Size(140, 35);
            this.textBox_62_4_3.TabIndex = 25;
            this.textBox_62_4_3.Text = "2000";
            // 
            // textBox_62_3_3
            // 
            this.textBox_62_3_3.Location = new System.Drawing.Point(491, 197);
            this.textBox_62_3_3.Name = "textBox_62_3_3";
            this.textBox_62_3_3.Size = new System.Drawing.Size(140, 35);
            this.textBox_62_3_3.TabIndex = 24;
            this.textBox_62_3_3.Text = "1000";
            // 
            // textBox_62_2_3
            // 
            this.textBox_62_2_3.Location = new System.Drawing.Point(491, 147);
            this.textBox_62_2_3.Name = "textBox_62_2_3";
            this.textBox_62_2_3.Size = new System.Drawing.Size(140, 35);
            this.textBox_62_2_3.TabIndex = 23;
            this.textBox_62_2_3.Text = "2000";
            // 
            // textBox_62_1_3
            // 
            this.textBox_62_1_3.Location = new System.Drawing.Point(491, 97);
            this.textBox_62_1_3.Name = "textBox_62_1_3";
            this.textBox_62_1_3.Size = new System.Drawing.Size(140, 35);
            this.textBox_62_1_3.TabIndex = 22;
            this.textBox_62_1_3.Text = "1000";
            // 
            // textBox_62_6_2
            // 
            this.textBox_62_6_2.Location = new System.Drawing.Point(347, 347);
            this.textBox_62_6_2.Name = "textBox_62_6_2";
            this.textBox_62_6_2.Size = new System.Drawing.Size(100, 35);
            this.textBox_62_6_2.TabIndex = 21;
            this.textBox_62_6_2.Text = "5025";
            // 
            // textBox_62_5_2
            // 
            this.textBox_62_5_2.Location = new System.Drawing.Point(347, 297);
            this.textBox_62_5_2.Name = "textBox_62_5_2";
            this.textBox_62_5_2.Size = new System.Drawing.Size(100, 35);
            this.textBox_62_5_2.TabIndex = 20;
            this.textBox_62_5_2.Text = "5025";
            // 
            // textBox_62_4_2
            // 
            this.textBox_62_4_2.Location = new System.Drawing.Point(347, 247);
            this.textBox_62_4_2.Name = "textBox_62_4_2";
            this.textBox_62_4_2.Size = new System.Drawing.Size(100, 35);
            this.textBox_62_4_2.TabIndex = 19;
            this.textBox_62_4_2.Text = "5025";
            // 
            // textBox_62_3_2
            // 
            this.textBox_62_3_2.Location = new System.Drawing.Point(347, 197);
            this.textBox_62_3_2.Name = "textBox_62_3_2";
            this.textBox_62_3_2.Size = new System.Drawing.Size(100, 35);
            this.textBox_62_3_2.TabIndex = 18;
            this.textBox_62_3_2.Text = "5025";
            // 
            // textBox_62_2_2
            // 
            this.textBox_62_2_2.Location = new System.Drawing.Point(347, 147);
            this.textBox_62_2_2.Name = "textBox_62_2_2";
            this.textBox_62_2_2.Size = new System.Drawing.Size(100, 35);
            this.textBox_62_2_2.TabIndex = 17;
            this.textBox_62_2_2.Text = "5025";
            // 
            // textBox_62_1_2
            // 
            this.textBox_62_1_2.Location = new System.Drawing.Point(347, 97);
            this.textBox_62_1_2.Name = "textBox_62_1_2";
            this.textBox_62_1_2.Size = new System.Drawing.Size(100, 35);
            this.textBox_62_1_2.TabIndex = 16;
            this.textBox_62_1_2.Text = "5025";
            // 
            // textBox_62_6_1
            // 
            this.textBox_62_6_1.Location = new System.Drawing.Point(166, 347);
            this.textBox_62_6_1.Name = "textBox_62_6_1";
            this.textBox_62_6_1.Size = new System.Drawing.Size(140, 35);
            this.textBox_62_6_1.TabIndex = 15;
            this.textBox_62_6_1.Text = "192.168.0.10";
            // 
            // textBox_62_5_1
            // 
            this.textBox_62_5_1.Location = new System.Drawing.Point(166, 297);
            this.textBox_62_5_1.Name = "textBox_62_5_1";
            this.textBox_62_5_1.Size = new System.Drawing.Size(140, 35);
            this.textBox_62_5_1.TabIndex = 14;
            this.textBox_62_5_1.Text = "192.168.0.10";
            // 
            // textBox_62_4_1
            // 
            this.textBox_62_4_1.Location = new System.Drawing.Point(166, 247);
            this.textBox_62_4_1.Name = "textBox_62_4_1";
            this.textBox_62_4_1.Size = new System.Drawing.Size(140, 35);
            this.textBox_62_4_1.TabIndex = 13;
            this.textBox_62_4_1.Text = "192.168.0.10";
            // 
            // textBox_62_3_1
            // 
            this.textBox_62_3_1.Location = new System.Drawing.Point(166, 197);
            this.textBox_62_3_1.Name = "textBox_62_3_1";
            this.textBox_62_3_1.Size = new System.Drawing.Size(140, 35);
            this.textBox_62_3_1.TabIndex = 12;
            this.textBox_62_3_1.Text = "192.168.0.10";
            // 
            // textBox_62_2_1
            // 
            this.textBox_62_2_1.Location = new System.Drawing.Point(166, 147);
            this.textBox_62_2_1.Name = "textBox_62_2_1";
            this.textBox_62_2_1.Size = new System.Drawing.Size(140, 35);
            this.textBox_62_2_1.TabIndex = 11;
            this.textBox_62_2_1.Text = "192.168.0.10";
            // 
            // label118
            // 
            this.label118.AutoSize = true;
            this.label118.Location = new System.Drawing.Point(89, 350);
            this.label118.Name = "label118";
            this.label118.Size = new System.Drawing.Size(70, 24);
            this.label118.TabIndex = 10;
            this.label118.Text = "电源6";
            // 
            // label117
            // 
            this.label117.AutoSize = true;
            this.label117.Location = new System.Drawing.Point(89, 300);
            this.label117.Name = "label117";
            this.label117.Size = new System.Drawing.Size(70, 24);
            this.label117.TabIndex = 9;
            this.label117.Text = "电源5";
            // 
            // label116
            // 
            this.label116.AutoSize = true;
            this.label116.Location = new System.Drawing.Point(89, 250);
            this.label116.Name = "label116";
            this.label116.Size = new System.Drawing.Size(70, 24);
            this.label116.TabIndex = 8;
            this.label116.Text = "电源4";
            // 
            // label108
            // 
            this.label108.AutoSize = true;
            this.label108.Location = new System.Drawing.Point(89, 200);
            this.label108.Name = "label108";
            this.label108.Size = new System.Drawing.Size(70, 24);
            this.label108.TabIndex = 7;
            this.label108.Text = "电源3";
            // 
            // label107
            // 
            this.label107.AutoSize = true;
            this.label107.Location = new System.Drawing.Point(89, 150);
            this.label107.Name = "label107";
            this.label107.Size = new System.Drawing.Size(70, 24);
            this.label107.TabIndex = 6;
            this.label107.Text = "电源2";
            // 
            // label106
            // 
            this.label106.AutoSize = true;
            this.label106.Location = new System.Drawing.Point(89, 100);
            this.label106.Name = "label106";
            this.label106.Size = new System.Drawing.Size(70, 24);
            this.label106.TabIndex = 5;
            this.label106.Text = "电源1";
            // 
            // label105
            // 
            this.label105.AutoSize = true;
            this.label105.Location = new System.Drawing.Point(708, 55);
            this.label105.Name = "label105";
            this.label105.Size = new System.Drawing.Size(118, 24);
            this.label105.TabIndex = 4;
            this.label105.Text = "零场电流A";
            // 
            // label104
            // 
            this.label104.AutoSize = true;
            this.label104.Location = new System.Drawing.Point(519, 55);
            this.label104.Name = "label104";
            this.label104.Size = new System.Drawing.Size(154, 24);
            this.label104.TabIndex = 3;
            this.label104.Text = "线圈常数nT/A";
            // 
            // label103
            // 
            this.label103.AutoSize = true;
            this.label103.Location = new System.Drawing.Point(379, 55);
            this.label103.Name = "label103";
            this.label103.Size = new System.Drawing.Size(58, 24);
            this.label103.TabIndex = 2;
            this.label103.Text = "端口";
            // 
            // label102
            // 
            this.label102.AutoSize = true;
            this.label102.Location = new System.Drawing.Point(220, 55);
            this.label102.Name = "label102";
            this.label102.Size = new System.Drawing.Size(58, 24);
            this.label102.TabIndex = 1;
            this.label102.Text = "地址";
            // 
            // textBox_62_1_1
            // 
            this.textBox_62_1_1.Location = new System.Drawing.Point(166, 97);
            this.textBox_62_1_1.Name = "textBox_62_1_1";
            this.textBox_62_1_1.Size = new System.Drawing.Size(140, 35);
            this.textBox_62_1_1.TabIndex = 0;
            this.textBox_62_1_1.Text = "192.168.0.10";
            // 
            // circularGauge1
            // 
            this.circularGauge1.CurrentAngle = 0F;
            this.circularGauge1.Location = new System.Drawing.Point(616, 0);
            this.circularGauge1.Name = "circularGauge1";
            this.circularGauge1.Size = new System.Drawing.Size(250, 250);
            this.circularGauge1.TabIndex = 1;
            // 
            // circularGauge2
            // 
            this.circularGauge2.CurrentAngle = 0F;
            this.circularGauge2.Location = new System.Drawing.Point(279, 115);
            this.circularGauge2.Name = "circularGauge2";
            this.circularGauge2.Size = new System.Drawing.Size(500, 500);
            this.circularGauge2.TabIndex = 2;
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(1184, 761);
            this.Controls.Add(this.tabControl1);
            this.Name = "Form1";
            this.Text = "磁试验数据采集及处理系统";
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.groupBox21.ResumeLayout(false);
            this.groupBox5.ResumeLayout(false);
            this.tabControl11.ResumeLayout(false);
            this.tabPage11.ResumeLayout(false);
            this.tabPage11.PerformLayout();
            this.tabPage12.ResumeLayout(false);
            this.tabPage12.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown13)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown12)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown11)).EndInit();
            this.groupBox_12_2.ResumeLayout(false);
            this.groupBox7.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox11)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            this.groupBox8.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).EndInit();
            this.tabPage2.ResumeLayout(false);
            this.groupBox24.ResumeLayout(false);
            this.groupBox24.PerformLayout();
            this.groupBox18.ResumeLayout(false);
            this.groupBox18.PerformLayout();
            this.tabControl23.ResumeLayout(false);
            this.tabPage231.ResumeLayout(false);
            this.tabControl4.ResumeLayout(false);
            this.tabPage10.ResumeLayout(false);
            this.tabPage10.PerformLayout();
            this.tabPage15.ResumeLayout(false);
            this.tabPage15.PerformLayout();
            this.groupBox20.ResumeLayout(false);
            this.groupBox20.PerformLayout();
            this.groupBox19.ResumeLayout(false);
            this.groupBox19.PerformLayout();
            this.tabPage232.ResumeLayout(false);
            this.tabPage232.PerformLayout();
            this.groupBox17.ResumeLayout(false);
            this.groupBox17.PerformLayout();
            this.groupBox16.ResumeLayout(false);
            this.groupBox16.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.tabControl2.ResumeLayout(false);
            this.tabPage7.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabPage8.ResumeLayout(false);
            this.groupBox15.ResumeLayout(false);
            this.groupBox15.PerformLayout();
            this.groupBox14.ResumeLayout(false);
            this.groupBox14.PerformLayout();
            this.groupBox13.ResumeLayout(false);
            this.groupBox13.PerformLayout();
            this.groupBox12.ResumeLayout(false);
            this.groupBox12.PerformLayout();
            this.groupBox9.ResumeLayout(false);
            this.groupBox9.PerformLayout();
            this.groupBox10.ResumeLayout(false);
            this.groupBox10.PerformLayout();
            this.groupBox11.ResumeLayout(false);
            this.tabPage4.ResumeLayout(false);
            this.tabPage4.PerformLayout();
            this.tabPage6.ResumeLayout(false);
            this.tabControl5.ResumeLayout(false);
            this.tabPage13.ResumeLayout(false);
            this.tabPage13.PerformLayout();
            this.groupBox23.ResumeLayout(false);
            this.groupBox23.PerformLayout();
            this.groupBox22.ResumeLayout(false);
            this.groupBox22.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_61_1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_61_2)).EndInit();
            this.tabPage14.ResumeLayout(false);
            this.tabPage14.PerformLayout();
            this.ResumeLayout(false);

        }

        #region 选项1数据采集


        #region 选项1数据采集-1实时采集-初始化
        //初始化
        private readonly List<ProbeInfo> _probeList = new List<ProbeInfo>();
        private readonly Dictionary<string, CircularBuffer> _plotData = new Dictionary<string, CircularBuffer>();
        private readonly object _dataLock = new object();


        // 定义HS主机配置
        private readonly Dictionary<string, string> _hsHostConfigs = new Dictionary<string, string>
            {
                 { "HS1", "192.168.1.11:6001" },
                 { "HS2", "192.168.1.12:6002" },
                 { "HS3", "192.168.1.13:6003" },
                 { "HS4", "192.168.1.14:6004" }
             };
        private Dictionary<string, HostInfo> _connectedHosts = new Dictionary<string, HostInfo>();//用于连接HS主机和从机tcp
        private HostInfo _mainhsHost;
        private TcpClient _HScontrolClient = new TcpClient();
        private NetworkStream _HScontrolStream;
        private List<ProbeTcpClientHS> _HSClients = new List<ProbeTcpClientHS>();
        //获取坐标转换规则
        private string HSxTransformationType = "X";
        private string HSyTransformationType = "Y";
        private string HSzTransformationType = "Z";
        private List<ProbeTcpClientCH> _CHtcpClients = new List<ProbeTcpClientCH>();
        //获取坐标转换规则
        private string CHxTransformationType = "X";
        private string CHyTransformationType = "Y";
        private string CHzTransformationType = "Z";


        private AngleDetector _angledetector = new AngleDetector(); //大转台角度类
        DataTable angledataTable = new DataTable();
        private DataColumn[] angledataColumn = null;
        private System.Windows.Forms.DataGrid angleacquisitionDataGrid = new DataGrid();


        private void start11()
        {
            button_11_4.Enabled = true;
            button_11_5.Enabled = false;

            formsPlot1.Plot.Legend.IsVisible = false;//折线图隐藏图例
            formsPlot1.UserInputProcessor.Disable(); // 禁用所有鼠标交互

            //添加自定义图例，包含颜色设置一一对应
            InitializeColorPalette();
            listBox1.DrawMode = DrawMode.OwnerDrawFixed;
            listBox1.ItemHeight = 22;
            listBox1.SelectionMode = SelectionMode.None; // 或 One，避免选中高亮干扰
            listBox1.DrawItem += ListBox1_DrawItem;
            listBox1.MouseDown += ListBox1_MouseDown; // 用于点击 checkbox 切换



            InitializeDataGridView1();

            DingSound_Load();//提示音加载

            CircularGauge1set(0);//仪表盘角度

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        private void InitializeDataGridView1()//初始化dataGridView1
        {

            dataGridView1.AutoGenerateColumns = false;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.Rows.Clear();

            dataGridView1.Columns.Clear();
            dataGridView1.Columns.Add("ProbeName", "探头名称");
            dataGridView1.Columns["ProbeName"].Width = 130;
            dataGridView1.Columns.Add("X", "X (nT)");
            dataGridView1.Columns["X"].Width = 130;
            dataGridView1.Columns.Add("Y", "Y (nT)");
            dataGridView1.Columns["Y"].Width = 130;
            dataGridView1.Columns.Add("Z", "Z (nT)");
            dataGridView1.Columns["Z"].Width = 130;
            //dataGridView1.Columns.Add("DeviceAddress", "设备地址");
            //dataGridView1.Columns["DeviceAddress"].Width = 150;
        }

        private void button_11_1_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView_61_1.Rows.Count < 1)
                {
                    MessageBox.Show("请先选择要连接的探头");
                    tabControl1.SelectedIndex = 5;
                    tabControl5.SelectedIndex = 0;
                    return;
                }

                bool isTcpConnected = _CHtcpClients.Count > 0 || _HSClients.Count > 0;

                if (isTcpConnected)
                {
                    var result = MessageBox.Show(
                        "是否停止现有数据连接？",
                        "存在活跃连接",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        StopDataMonitoring();
                        DisconnectAllProbes();
                    }
                    else
                    {
                        return;
                    }
                }


                HSxTransformationType = comboBox_61_3.SelectedItem.ToString() ?? "X";
                HSyTransformationType = comboBox_61_4.SelectedItem.ToString() ?? "Y";
                HSzTransformationType = comboBox_61_5.SelectedItem.ToString() ?? "Z";


                //checkedListBox1.Items.Clear();
                listBox1.Items.Clear();
                InitializeDataStructures();

                // 使用Task.Run避免阻塞UI
                Task.Run(() =>
                {
                    if (_probeList.Any(p => p.DeviceType.StartsWith("CH")))
                        ProcessCHDevices();

                    if (_probeList.Any(p => p.DeviceType.StartsWith("HS")))
                        ProcessHSDevices();


                });

                this.BeginInvoke((MethodInvoker)delegate
                {
                    //InitializeCheckedListBox1();
                    InitializeListBox1();
                    StartDataMonitoring();
                    StartProbeMonitoring();
                });

                ConnectAngleDetector();

            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化失败: {ex.Message}");
            }
        }

        private void ConnectAngleDetector()
        {
            try
            {
                //连接大转台角度获取
                _angledetector.AngleChanged += OnAngleChanged;
                //报告转速事件
                _angledetector.ReportRotationSpeed += OnReportRotationSpeed;
                String angleDevice = comboBox_4_1.Text;
                //初始化大转台连接
                _angledetector.Start(angleDevice + "/ai0:3", -10, 10, 1000, 20, angledataTable, angledataColumn, angleacquisitionDataGrid);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"连接大转台失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void button_12_1_Click(object sender, EventArgs e)
        {
            DisconnectAllProbes();
            StopProbeMonitoring();

            await Task.Run(() => _angledetector.Stop());
            pictureBox11.BackColor = System.Drawing.Color.Gray;


            MessageBox.Show("已断开所有探头连接");
        }


        private void InitializeListBox1()
        {
            listBox1.Items.Clear();
            int index = 1; // 从1开始编号

            foreach (var probe in _probeList)
            {
                listBox1.Items.Add(new CheckableItem { Text = $"{probe.ProbeName}X", Checked = true, DisplayIndex = index++ });
                listBox1.Items.Add(new CheckableItem { Text = $"{probe.ProbeName}Y", Checked = true, DisplayIndex = index++ });
                listBox1.Items.Add(new CheckableItem { Text = $"{probe.ProbeName}Z", Checked = true, DisplayIndex = index++ });
            }

            UpdatePlot1Display();
        }
        private void InitializeCheckedListBox1()// 初始化checkedListBox1的方法
        {
            // 临时禁用事件
            checkedListBox1.ItemCheck -= CheckedListBox1_ItemCheck;


            // 清除现有项（但不清除历史数据）
            checkedListBox1.Items.Clear();

            // 添加新的探头项
            //foreach (var probe in _probeList)
            for (int i = 0; i < _probeList.Count; i++)
            {
                var probe = _probeList[i];
                checkedListBox1.Items.Add($"{probe.ProbeName}X", true);
                checkedListBox1.Items.Add($"{probe.ProbeName}Y", true);
                checkedListBox1.Items.Add($"{probe.ProbeName}Z", true);
            }


            checkedListBox1.Refresh();

            // 重新启用事件
            checkedListBox1.ItemCheck += CheckedListBox1_ItemCheck;
        }

        private void InitializeDataStructures()//初始化图表
        {
            lock (_dataLock)
            {
                // 1. 清空现有数据
                _probeList.Clear();
                dataGridView1.Rows.Clear();

                _plotData.Clear();
                formsPlot1.Plot.Clear();

                // 2. 从dataGridView_61_1获取选中的探头
                foreach (DataGridViewRow row in dataGridView_61_1.Rows)
                {
                    if (row.IsNewRow ||
            row.Cells["探头序号"].Value == null ||
            row.Cells["设备号"].Value == null ||
            row.Cells["探头"].Value == null ||
            row.Cells["地址"].Value == null)
                    {
                        continue;
                    }

                    var probe = new ProbeInfo
                    {
                        ProbeName = row.Cells["探头序号"].Value?.ToString(),
                        DeviceType = row.Cells["设备号"].Value?.ToString(),
                        DeviceAddress = row.Cells["地址"].Value?.ToString(),
                        ProbeChannel = Convert.ToInt32(row.Cells["探头"].Value),
                        X = 0,
                        Y = 0,
                        Z = 0
                    };

                    // 3. 添加到_probeList
                    _probeList.Add(probe);

                    // 4. 初始化DataGridView行
                    dataGridView1.Rows.Add(
                        probe.ProbeName,
                        "0",
                        "0",
                        "0");
                    //probe.DeviceAddress);

                    //// 5. 初始化checkedListBox1
                    //checkedListBox1.Items.Add($"{probe.ProbeName}X", true);
                    //checkedListBox1.Items.Add($"{probe.ProbeName}Y", true);
                    //checkedListBox1.Items.Add($"{probe.ProbeName}Z", true);

                    // 6. 初始化绘图数据
                    //_plotData[$"{probe.ProbeName}X"] = new List<double>();
                    //_plotData[$"{probe.ProbeName}Y"] = new List<double>();
                    //_plotData[$"{probe.ProbeName}Z"] = new List<double>();
                    _plotData[$"{probe.ProbeName}X"] = new CircularBuffer(MaxformsPlot1Points);
                    _plotData[$"{probe.ProbeName}Y"] = new CircularBuffer(MaxformsPlot1Points);
                    _plotData[$"{probe.ProbeName}Z"] = new CircularBuffer(MaxformsPlot1Points);
                }

                //// 7. 设置checkedListBox1的事件处理
                //checkedListBox1.ItemCheck += (s, e) =>
                //{
                //    UpdatePlot1Display();
                //};
                //checkedListBox1.ItemCheck += CheckedListBox1_ItemCheck; // 重新添加事件
            }
        }

        private void DisconnectAllProbes() // 断开所有探头连接
        {
            try
            {
                // 停止数据监控
                StopDataMonitoring();


                // 断开CH设备
                foreach (var client in _CHtcpClients)
                {
                    client.ClearChannelMappings();
                    client.Dispose();
                }
                _CHtcpClients.Clear();

                // 断开HS设备
                foreach (var client in _HSClients)
                {
                    client.Dispose();
                    //client.StopRecording();
                }
                _HSClients.Clear();


                // 关闭HS控制客户端
                StopAllHSConnections();
                //StopHScontrol();

                StopRecording();// 停止记录Excel

                _CHtcpClients.Clear();
                _HSClients.Clear();

                // 清空探头列表
                _probeList.Clear();

                // 清空checkedListBox1（先移除事件处理避免触发）
                //checkedListBox1.ItemCheck -= CheckedListBox1_ItemCheck; // 临时移除事件
                //checkedListBox1.Items.Clear();
                //checkedListBox1.ItemCheck += CheckedListBox1_ItemCheck; // 重新添加事件
            }
            catch (Exception ex)
            {
                MessageBox.Show($"断开连接时出错: {ex.Message}");
            }
        }

        // 模拟原来的 CheckedListBox1_ItemCheck 事件
        private void OnItemCheckSimulated(int index, bool oldValue, bool newValue)
        {
            this.BeginInvoke((MethodInvoker)(() => UpdatePlot1Display()));
        }
        private void CheckedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            this.BeginInvoke((MethodInvoker)(() => UpdatePlot1Display()));
        }
        private void button_11_7_Click(object sender, EventArgs e)// 全选所有项
        {
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                ((CheckableItem)listBox1.Items[i]).Checked = true;
            }
            listBox1.Invalidate();
            //for (int i = 0; i < checkedListBox1.Items.Count; i++)
            //{
            //    checkedListBox1.SetItemChecked(i, true);
            //}
            UpdatePlot1Display();
        }
        private void button_11_8_Click(object sender, EventArgs e)// 全不选所有项
        {
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                ((CheckableItem)listBox1.Items[i]).Checked = false;
            }
            listBox1.Invalidate();
            //for (int i = 0; i < checkedListBox1.Items.Count; i++)
            //{
            //    checkedListBox1.SetItemChecked(i, false);
            //}
            UpdatePlot1Display();
        }
        private void button_11_9_Click(object sender, EventArgs e)//清图
        {
            ClearFormsPlot1();
        }

        private void ClearFormsPlot1()
        {
            lock (_dataLock)//保留个位数值，确保图像不突兀
            {
                foreach (var buffer in _plotData.Values)
                {
                    int currentCount = buffer.Count;
                    if (currentCount == 0) continue;

                    int remainder = currentCount % 10;
                    int keepCount = remainder == 0 ? 10 : remainder;

                    buffer.KeepLast(keepCount);
                }
            }
            formsPlot1.Plot.Clear();
            formsPlot1.Refresh();
        }



        //bool RotationOpen = false;
        //private async void button_11_9_Click(object sender, EventArgs e)
        //{
        //    if (RotationOpen)
        //    {
        //        button_11_9.Enabled = false;
        //        //_angledetector.Stop();
        //        await Task.Run(() => _angledetector.Stop());
        //        pictureBox11.BackColor = System.Drawing.Color.Gray;
        //        button_11_9.Enabled = true;
        //        RotationOpen = false;
        //        button_11_9.Text = "开启";
        //    }
        //    else
        //    {
        //        button_11_9.Enabled = false;
        //        //连接大转台角度获取
        //        _angledetector.AngleChanged += OnAngleChanged;
        //        //报告转速事件
        //        _angledetector.ReportRotationSpeed += OnReportRotationSpeed;
        //        //初始化大转台连接
        //        _angledetector.Start("Dev1/ai0:3", -10, 10, 1000, 20, angledataTable, angledataColumn, angleacquisitionDataGrid);

        //        button_11_9.Enabled = true;
        //        RotationOpen = true;
        //        button_11_9.Text = "关闭";
        //    }
        //}

        //模拟仪表盘展示角度
        private void CircularGauge1set(int angle)
        {
            //circularGauge1.CurrentAngle = angle; // 指针指向220°位置

            // 1. 增加跨线程安全检查：如果不在UI线程，则通过Invoke封送
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action<int>(CircularGauge1set), angle);
                return;
            }

            // 2. 同时更新两个仪表盘，确保无论在哪个页面都能看到变化
            if (circularGauge1 != null) circularGauge1.CurrentAngle = angle; // Tab1 采集页面的圆盘
            if (circularGauge2 != null) circularGauge2.CurrentAngle = angle; // Tab4 配置页面的圆盘
        }

        //报告转速事件
        private void OnReportRotationSpeed(int consecutiveChannelCount)
        {
            if (pictureBox11.InvokeRequired)
            {
                //如果不在UI线程，则切换调用自身
                pictureBox11.Invoke(new Action<int>(OnReportRotationSpeed), consecutiveChannelCount);
                return;
            }

            if (consecutiveChannelCount > 3)
            {
                pictureBox11.BackColor = System.Drawing.Color.Green;
            }
            else
            {
                pictureBox11.BackColor = System.Drawing.Color.Red;
            }
        }



        #region 自定义ListBox1绘图
        private void ListBox1_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= listBox1.Items.Count) return;

            var item = (CheckableItem)listBox1.Items[e.Index];
            int itemNumber = e.Index + 1; // 假设第0项是"1"

            // 获取颜色（基于编号）
            //Color backColor = GetItemColor(itemNumber);
            Color backColor = Color.White;
            Color textColor = Color.Black;

            e.Graphics.FillRectangle(new SolidBrush(backColor), e.Bounds);

            ButtonState state = item.Checked ? ButtonState.Checked : ButtonState.Normal;
            ControlPaint.DrawCheckBox(e.Graphics,
                new Rectangle(e.Bounds.Left + 2, e.Bounds.Top + 3, 14, 14),
                state);


            // 获取用于绘制文本的矩形区域，留出空间给颜色方块
            Rectangle textRect = new Rectangle(
                e.Bounds.Left + 22,
                e.Bounds.Top,
                e.Bounds.Width - 56, // 减少宽度以留出空间给颜色方块和间距
                e.Bounds.Height);

            TextRenderer.DrawText(e.Graphics, item.Text, e.Font, textRect, textColor,
                          TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine);
            //TextRenderer.DrawText(e.Graphics, item.Text, e.Font,
            //    new Rectangle(e.Bounds.Left + 22, e.Bounds.Top, e.Bounds.Width - 22, e.Bounds.Height),
            //    textColor, Color.Transparent,
            //    TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine);

            // 计算颜色方块的位置
            Rectangle colorBoxRect = new Rectangle(
                e.Bounds.Right - 32, // 方块距离右边界的距离
                e.Bounds.Top + 3,
                16, // 方块的大小
                16);

            // 绘制颜色方块
            using (SolidBrush brush = new SolidBrush(GetItemColor(itemNumber)))
            {
                e.Graphics.FillRectangle(brush, colorBoxRect);
                e.Graphics.DrawRectangle(Pens.Black, colorBoxRect); // 可选：给方块加边框
            }

        }
        private void ListBox1_MouseDown(object sender, MouseEventArgs e)
        {
            int index = listBox1.IndexFromPoint(e.Location);
            if (index < 0 || index >= listBox1.Items.Count)
                return;

            // ✅ 关键：获取该项在当前视图中的真实矩形（已包含滚动偏移）
            Rectangle itemRect = listBox1.GetItemRectangle(index);

            // CheckBox 区域：相对于 itemRect 的左上角
            Rectangle checkBoxRect = new Rectangle(
                itemRect.Left + 2,
                itemRect.Top + 3,
                14,
                14
            );

            // 检查鼠标是否点击在 CheckBox 区域内
            if (checkBoxRect.Contains(e.Location))
            {
                var item = (CheckableItem)listBox1.Items[index];
                bool oldChecked = item.Checked;
                item.Checked = !item.Checked;

                // 触发逻辑
                OnItemCheckSimulated(index, oldChecked, item.Checked);

                // 仅重绘该项（高效）
                listBox1.Invalidate(itemRect);
            }
        }


        private Color[] _itemColors;
        private void InitializeColorPalette()
        {
            const int totalColors = 96;
            _itemColors = new Color[totalColors];

            // 选择一个与 96 互质的步长（推荐 35, 37, 41, 43...）
            int step = 35; // gcd(35,96)=1，互质

            for (int i = 0; i < totalColors; i++)
            {
                // 跳跃索引，打乱顺序
                int hueIndex = (i * step) % totalColors;

                // 将 hueIndex 映射到 0~360 度
                float hue = (hueIndex / (float)totalColors) * 360f;

                // 使用高饱和度、中等明度，确保鲜艳且易区分
                _itemColors[i] = FromHsv(hue, 0.9f, 0.95f);
            }
        }
        public static Color FromHsv(float h, float s, float v)
        {
            // 标准 HSV to RGB 算法
            double r = 0, g = 0, b = 0;

            if (s == 0)
            {
                r = g = b = v;
            }
            else
            {
                h = h % 360;
                if (h < 0) h += 360;

                int i = (int)Math.Floor(h / 60);
                double f = h / 60 - i;
                double p = v * (1 - s);
                double q = v * (1 - s * f);
                double t = v * (1 - s * (1 - f));

                switch (i)
                {
                    case 0: r = v; g = t; b = p; break;
                    case 1: r = q; g = v; b = p; break;
                    case 2: r = p; g = v; b = t; break;
                    case 3: r = p; g = q; b = v; break;
                    case 4: r = t; g = p; b = v; break;
                    case 5: r = v; g = p; b = q; break;
                }
            }

            return Color.FromArgb(
                (int)(r * 255),
                (int)(g * 255),
                (int)(b * 255)
            );
        }
        public Color GetItemColor(int number)
        {
            // number 从 1 到 96
            if (number < 1 || number > 96)
                return Color.Gray; // 或抛出异常

            return _itemColors[number - 1]; // 转为 0-based 索引
        }
        // 将System.Drawing.Color转换为ScottPlot.Color
        private ScottPlot.Color GetScottPlotColor(int number)
        {
            if (number < 0 || number >= 96)
                return ScottPlot.Colors.Black; // 默认颜色

            System.Drawing.Color drawingColor = GetItemColor(number);
            return ScottPlot.Color.FromARGB(drawingColor.ToArgb());
        }
        #endregion



        #endregion

        #region 选项1数据采集-1实时采集-TCP连接



        // 第二级：处理CH设备连接和指令发送
        private void ProcessCHDevices()
        {
            // 1. 清空现有连接
            _CHtcpClients.Clear();

            // 2. 获取CH设备并按地址分组
            var chProbes = _probeList.Where(p => p.DeviceType.StartsWith("CH")).ToList();
            var addressGroups = chProbes.GroupBy(p => p.DeviceAddress).ToList();

            if (addressGroups.Count == 0) return;

            // 3. 建立TCP连接
            foreach (var group in addressGroups)
            {
                string address = group.Key;
                string[] parts = address.Split(':');
                if (parts.Length != 2 || !int.TryParse(parts[1], out int port))
                {
                    MessageBox.Show($"无效的地址格式: {address}", "连接错误",
                                 MessageBoxButtons.OK, MessageBoxIcon.Error);
                    continue;
                }

                string ip = parts[0];
                var probesInGroup = group.ToList();

                var client = new ProbeTcpClientCH(ip, port, probesInGroup);

                // 在后台线程执行连接
                var connectTask = Task.Run(() => client.Connect());

                if (connectTask.Wait(3000)) // 等待最多3秒
                {
                    if (connectTask.Result)
                    {
                        _CHtcpClients.Add(client);
                        foreach (var probe in probesInGroup)
                        {
                            int channel = (probe.ProbeChannel - 1) % 4 + 1;
                            client.AssignProbe(channel, probe);
                        }
                    }
                    else
                    {
                        MessageBox.Show($"{address} 连接失败", "连接错误",
                                     MessageBoxButtons.OK, MessageBoxIcon.Error);
                        client.Dispose();
                    }
                }
                else
                {
                    MessageBox.Show($"{address} 连接超时", "连接错误",
                                 MessageBoxButtons.OK, MessageBoxIcon.Error);
                    client.Dispose();
                }
            }

            // 4. 发送采集指令
            SendCHCollectionCommands();
        }

        // 第三级：发送CH设备采集指令
        private void SendCHCollectionCommands()
        {
            if (_CHtcpClients.Count == 0) return;

            // 获取采样率
            int sampleRate = GetSelectedSampleRate();
            byte[] rateBytes = BitConverter.GetBytes(sampleRate);

            if (_CHtcpClients.Count == 1)
            {
                // 单设备-主卡
                var masterCommand = CreateStartCommand(true, sampleRate);
                _CHtcpClients[0].SendStartCommand(true, sampleRate);
            }
            else
            {
                // 多设备-先发从卡
                for (int i = 1; i < _CHtcpClients.Count; i++)
                {
                    _CHtcpClients[i].SendStartCommand(false, sampleRate);
                }

                // 再发主卡
                _CHtcpClients[0].SendStartCommand(true, sampleRate);
            }
        }
        // 辅助方法：从comboBox获取采样率
        private int GetSelectedSampleRate()
        {
            // 如果是在UI线程中调用，直接获取值
            if (comboBox_61_1.InvokeRequired == false)
            {
                string rateText = comboBox_61_1.SelectedItem?.ToString();
                if (string.IsNullOrEmpty(rateText)) return 1000;

                if (int.TryParse(rateText, out int rate))
                {
                    return rate;
                }
                return 1000;
            }
            // 如果是在后台线程中调用，使用Invoke
            else
            {
                return (int)comboBox_61_1.Invoke(new Func<int>(() =>
                {
                    string rateText = comboBox_61_1.SelectedItem?.ToString();
                    if (string.IsNullOrEmpty(rateText)) return 1000;

                    if (int.TryParse(rateText, out int rate))
                    {
                        return rate;
                    }
                    return 1000;
                }));
            }
        }

        // 辅助方法：创建启动命令字节数组
        private byte[] CreateStartCommand(bool isMaster, int sampleRate)
        {
            byte[] command = new byte[8];
            command[0] = 0x9F; // 启动指令
            command[4] = isMaster ? (byte)0x00 : (byte)0x01; // 主从标志

            // 采样率（小端）
            command[6] = (byte)(sampleRate & 0xFF);        // 低字节
            command[7] = (byte)((sampleRate >> 8) & 0xFF); // 高字节

            return command;
        }



        // 第二级：处理HS设备连接和控制
        private void ProcessHSDevices()
        {
            // 1. 清空现有连接
            _HSClients.Clear();
            _connectedHosts.Clear();
            _mainhsHost = null;


            // 2. 获取HS设备
            var hsProbes = _probeList.Where(p => p.DeviceType.StartsWith("HS")).ToList();
            if (hsProbes.Count == 0) return;

            // 3. 判断模式
            var distinctDeviceTypes = hsProbes.Select(p => p.DeviceType).Distinct().ToList();
            bool isSingleMode = true;
            if (distinctDeviceTypes.Count > 1)
            {
                isSingleMode = false;
            }
            Console.WriteLine($"HS设备模式: {(isSingleMode ? "单机模式" : "组网模式")}");


            // 4. 建立所有HS探头的TCP连接
            // 按地址分组,每个地址对应2个探头
            var addressGroups = hsProbes.GroupBy(p => p.DeviceAddress).ToList();
            foreach (var probe in addressGroups)
            {
                string address = probe.Key;
                string[] parts = address.Split(':');
                if (parts.Length != 2 || !int.TryParse(parts[1], out int port))
                {
                    MessageBox.Show($"无效的地址格式: {address}", "连接错误",
                                 MessageBoxButtons.OK, MessageBoxIcon.Error);
                    continue;
                }

                string ip = parts[0];
                var probesInGroup = probe.ToList();
                var client = new ProbeTcpClientHS(ip, port, probesInGroup, this);
                client.ChangeTransform(HSxTransformationType, HSyTransformationType, HSzTransformationType);//添加坐标转换规则

                var connectTask = Task.Run(() => client.Connect());

                if (connectTask.Wait(1000))
                {
                    if (connectTask.Result)
                    {
                        _HSClients.Add(client);
                    }
                    else
                    {
                        MessageBox.Show($"{address} 连接失败", "连接错误",
                                     MessageBoxButtons.OK, MessageBoxIcon.Error);
                        client.Dispose();
                    }
                }
                else
                {
                    MessageBox.Show($"{address} 连接超时", "连接错误",
                                 MessageBoxButtons.OK, MessageBoxIcon.Error);
                    client.Dispose();
                }
            }

            // 5. 连接控制主机
            if (isSingleMode)
            {
                // 单机模式 - 使用原有的ConnectHSControlClient
                this.BeginInvoke((MethodInvoker)delegate
                {
                    ConnectHSControlClient();
                });
            }
            else
            {
                // 组网模式 - 使用新的组网逻辑
                ConnectHSControlHosts(hsProbes);
            }
        }


        // 第三级：连接HS控制客户端并发送指令
        private void ConnectHSControlClient()
        {
            try
            {
                string controlAddress = comboBox_61_2.SelectedItem.ToString();
                if (string.IsNullOrEmpty(controlAddress)) return;

                var addressParts = controlAddress.Split(':');
                if (addressParts.Length != 2) return;

                string ip = addressParts[0];
                int port = int.Parse(addressParts[1]);

                // 连接控制客户端
                _HScontrolClient = new TcpClient();
                //_HScontrolClient.Connect(ip, port);
                // 使用异步连接并设置超时
                var connectTask = _HScontrolClient.ConnectAsync(ip, port);
                // 设置3秒超时
                if (!connectTask.Wait(3000))
                {
                    _HScontrolClient.Close();
                    _HScontrolClient = null;
                    MessageBox.Show("控制客户端" + controlAddress + "连接超时", "连接错误",
                                 MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (!_HScontrolClient.Connected)
                {
                    _HScontrolClient.Close();
                    _HScontrolClient = null;
                    MessageBox.Show("控制客户端" + controlAddress + "连接失败", "连接错误",
                                 MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                _HScontrolStream = _HScontrolClient.GetStream();

                // 发送采样率设置
                string sampleRate = comboBox_61_1.SelectedItem.ToString();
                if (!string.IsNullOrEmpty(sampleRate))
                {
                    byte[] rateCommand = Encoding.ASCII.GetBytes($"rate:{sampleRate}");
                    _HScontrolStream.Write(rateCommand, 0, rateCommand.Length);

                    // 延时100ms后发送开始指令
                    Thread.Sleep(100);
                    byte[] startCommand = Encoding.ASCII.GetBytes("start");
                    _HScontrolStream.Write(startCommand, 0, startCommand.Length);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"HS控制客户端连接失败: {ex.Message}");
            }
        }



        private void ConnectHSControlHosts(List<ProbeInfo> hsProbes)
        {
            this.BeginInvoke((MethodInvoker)delegate
            {
                try
                {
                    // 1. 获取需要连接的主机
                    var neededHosts = GetNeededHosts(hsProbes);
                    if (neededHosts.Count == 0)
                    {
                        MessageBox.Show("未找到需要连接的主机");
                        return;
                    }

                    // 2. 连接所有需要的主机
                    ConnectAllHosts(neededHosts);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"主机连接失败: {ex.Message}");
                }
            });
        }
        // 获取需要连接的主机列表
        private Dictionary<string, string> GetNeededHosts(List<ProbeInfo> hsProbes)
        {
            var neededHosts = new Dictionary<string, string>();

            // 根据探头设备类型确定需要连接的主机
            var deviceTypes = hsProbes.Select(p => p.DeviceType).Distinct();
            foreach (var deviceType in deviceTypes)
            {
                if (_hsHostConfigs.ContainsKey(deviceType))
                {
                    neededHosts[deviceType] = _hsHostConfigs[deviceType];
                }
            }

            return neededHosts;
        }

        // 连接所有主机
        private async void ConnectAllHosts(Dictionary<string, string> neededHosts)
        {
            try
            {
                // 连接所有主机
                foreach (var host in neededHosts)
                {
                    bool connected = await ConnectHostAsync(host.Key, host.Value);
                    if (!connected)
                    {
                        MessageBox.Show($"{host.Key} 连接失败: {host.Value}");
                        return;
                    }
                }

                // 确定主主机
                string mainHostAddress = comboBox_61_2.SelectedItem?.ToString();
                _mainhsHost = _connectedHosts.Values.FirstOrDefault(h => h.Address == mainHostAddress);

                if (_mainhsHost == null)
                {
                    MessageBox.Show("未找到对应的主主机");
                    return;
                }

                // 发送指令：先发送从主机start，再发送主主机rate+start
                await SendControlCommands();

            }
            catch (Exception ex)
            {
                MessageBox.Show($"主机控制指令发送失败: {ex.Message}");
            }
        }
        // 连接单个主机
        private async Task<bool> ConnectHostAsync(string hostId, string address)
        {
            try
            {
                var addressParts = address.Split(':');
                if (addressParts.Length != 2) return false;

                string ip = addressParts[0];
                int port = int.Parse(addressParts[1]);

                var client = new TcpClient();
                var connectTask = client.ConnectAsync(ip, port);

                // 3秒超时
                if (await Task.WhenAny(connectTask, Task.Delay(3000)) != connectTask || !client.Connected)
                {
                    client.Close();
                    return false;
                }

                var hostInfo = new HostInfo
                {
                    HostId = hostId,
                    Address = address,
                    Client = client,
                    Stream = client.GetStream(),
                    IsMainHost = false // 稍后确定
                };

                _connectedHosts[hostId] = hostInfo;
                Console.WriteLine($"{hostId} 连接成功: {address}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{hostId} 连接异常: {ex.Message}");
                return false;
            }
        }

        private async Task SendControlCommands()
        {
            try
            {
                // 1. 先对所有从主机发送"start"
                var slaveHosts = _connectedHosts.Values.Where(h => h != _mainhsHost).ToList();
                foreach (var host in slaveHosts)
                {
                    await SendStartCommand(host);
                }

                // 2. 对主主机发送"rate:" + 采样率 和 "start"
                string sampleRate = comboBox_61_1.SelectedItem?.ToString();
                if (!string.IsNullOrEmpty(sampleRate))
                {
                    // 发送采样率设置
                    byte[] rateCommand = Encoding.ASCII.GetBytes($"rate:{sampleRate}");
                    await _mainhsHost.Stream.WriteAsync(rateCommand, 0, rateCommand.Length);

                    // 短暂延时
                    await Task.Delay(100);

                    // 发送开始指令
                    await SendStartCommand(_mainhsHost);

                }

                Console.WriteLine("所有控制指令发送完成");
            }
            catch (Exception ex)
            {
                throw new Exception($"发送控制指令失败: {ex.Message}");
            }
        }

        // 发送start命令
        private async Task SendStartCommand(HostInfo host)
        {
            try
            {
                byte[] startCommand = Encoding.ASCII.GetBytes("start");
                await host.Stream.WriteAsync(startCommand, 0, startCommand.Length);
                Console.WriteLine($"已发送start指令到 {host.HostId}");
            }
            catch (Exception ex)
            {
                throw new Exception($"发送start指令到 {host.HostId} 失败: {ex.Message}");
            }
        }



        private void StopAllHSConnections()
        {
            try
            {
                // 停止数据采集连接
                foreach (var client in _HSClients)
                {
                    client.Disconnect();
                }
                _HSClients.Clear();

                // 停止控制连接（单机模式）
                if (_HScontrolClient?.Connected == true)
                {
                    try
                    {
                        byte[] stopCommand = Encoding.ASCII.GetBytes("stop");
                        _HScontrolStream?.Write(stopCommand, 0, stopCommand.Length);
                        Thread.Sleep(50);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"停止单机控制连接异常: {ex.Message}");
                    }
                    finally
                    {
                        _HScontrolStream?.Close();
                        _HScontrolClient?.Close();
                        _HScontrolStream = null;
                        _HScontrolClient = null;
                    }
                }

                // 停止控制连接（组网模式）
                StopHScontrol();

                Console.WriteLine("所有HS连接已停止");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"停止HS连接失败: {ex.Message}");
            }
        }
        // 停止HS设备采集(组网模式)
        private void StopHScontrol()
        {
            try
            {
                // 停止所有主机
                foreach (var host in _connectedHosts.Values)
                {
                    if (host.Client?.Connected == true)
                    {
                        try
                        {
                            byte[] stopCommand = Encoding.ASCII.GetBytes("stop");
                            host.Stream.Write(stopCommand, 0, stopCommand.Length);
                            Thread.Sleep(50);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"停止 {host.HostId} 异常: {ex.Message}");
                        }
                        finally
                        {
                            host.Stream?.Close();
                            host.Client?.Close();
                        }
                        Console.WriteLine($"已停止 {host.HostId}");
                    }
                }

                _connectedHosts.Clear();
                _mainhsHost = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"停止HS设备失败: {ex.Message}");
            }
        }


        //// 窗体关闭时释放资源
        //protected override void OnFormClosing(FormClosingEventArgs e)
        //{
        //    base.OnFormClosing(e);

        //    DisconnectAllProbes();
        //}









        #endregion

        #region 选项1数据采集-1实时采集-数据展示

        private System.Threading.Timer _dataGridTimer;
        private System.Threading.Timer _dataCollectionTimer;
        private System.Threading.Timer _plotUpdateTimer;
        private volatile bool _isMonitoring = false;
        private const int MaxformsPlot1Points = 2000; // 最大显示点数
        private List<Color> checkedListBox1ItemColors = new List<Color>();

        private readonly Stopwatch _stopwatch = Stopwatch.StartNew();// 高精度时间源
        // 下一次采样/绘图的绝对时间（单位：毫秒）
        private double _nextSampleTime = 0;
        private double _nextPlotTime = 0;
        // 定时器
        private System.Threading.Timer _highFreqTimer;
        private int dataCollectionInterval = 100;//采集率

        // 启动数据监控线程
        private void StartDataMonitoring()
        {
            _isMonitoring = true;

            // 获取采样率
            int sampleRate = GetSelectedSampleRate();
            //int dataCollectionInterval = 100 /*= sampleRate > 0 ? 1000 / sampleRate : 100*/;
            // UI刷新固定为200ms，避免过快刷新
            int uiRefreshInterval = 1000;

            if (sampleRate >= 10)
            {
                dataCollectionInterval = 100;
            }
            else if (sampleRate <= 1)
            {
                dataCollectionInterval = 1000;
            }

            // 初始化DataGridView更新定时器（每秒更新一次）
            _dataGridTimer = new System.Threading.Timer(
                _ => UpdateDataGridView1Safe(),
                null,
                0,
                1000);

            //// 数据收集定时器（根据采样率）
            //_dataCollectionTimer = new System.Threading.Timer(
            //    _ => CollectPlotData(),
            //    null,
            //    0,
            //    dataCollectionInterval);

            //// 初始化Plot更新定时器（根据采样率更新）
            //_plotUpdateTimer = new System.Threading.Timer(
            //    _ => UpdatePlotUISafe(),
            //    null,
            //    0,
            //    uiRefreshInterval);

            // 启动高精度计时器
            _stopwatch.Restart();

            _nextSampleTime = 0;
            _nextPlotTime = 0;
            _highFreqTimer = new System.Threading.Timer(_ =>
            {
                var elapsed = _stopwatch.Elapsed.TotalMilliseconds;

                if (elapsed >= _nextSampleTime)
                {
                    CollectPlotData();
                    _nextSampleTime += dataCollectionInterval; // 0.1s
                }

                if (elapsed >= _nextPlotTime)
                {
                    UpdatePlotUISafe();
                    _nextPlotTime += uiRefreshInterval; // 1.0s
                }
            }, null, 0, 10); // 10ms 检查周期

        }

        // 停止数据监控
        private void StopDataMonitoring()
        {
            _isMonitoring = false;
            _dataGridTimer?.Dispose();
            _plotUpdateTimer?.Dispose();

            _dataGridTimer = null;
            _plotUpdateTimer = null;

            _highFreqTimer?.Dispose();
            _highFreqTimer = null;
            _stopwatch.Stop();
        }

        // 更新DataGridView数据
        private void UpdateDataGridView1Safe()
        {
            if (!_isMonitoring) return;

            try
            {
                if (dataGridView1.InvokeRequired)
                {
                    dataGridView1.BeginInvoke(new Action(UpdateDataGridView1));
                }
                else
                {
                    UpdateDataGridView1();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"安全更新DataGridView失败: {ex.Message}");
            }
        }

        //数据收集
        private void CollectPlotData()
        {
            if (!_isMonitoring) return;

            try
            {
                // 快速收集数据到环形缓冲区
                lock (_dataLock)
                {
                    foreach (var probe in _probeList)
                    {
                        string xKey = $"{probe.ProbeName}X";
                        string yKey = $"{probe.ProbeName}Y";
                        string zKey = $"{probe.ProbeName}Z";

                        if (_plotData.ContainsKey(xKey) && !(probe.X == 0 && probe.Y == 0 && probe.Z == 0))
                        {
                            _plotData[xKey].Add(probe.X);
                            _plotData[yKey].Add(probe.Y);
                            _plotData[zKey].Add(probe.Z);
                            //Random random = new Random();
                            //_plotData[xKey].Add(random.NextDouble() * 10);
                            //_plotData[yKey].Add(random.NextDouble() * 30);
                            //_plotData[zKey].Add(random.NextDouble() * 50);
                        }
                    }
                }


                //// 快速收集数据，尽量减少锁的时间
                //lock (_dataLock)
                //{
                //    foreach (var probe in _probeList)
                //    {
                //        string xKey = $"{probe.ProbeName}X";
                //        string yKey = $"{probe.ProbeName}Y";
                //        string zKey = $"{probe.ProbeName}Z";

                //        // 快速添加数据
                //        _plotData[xKey].Add(probe.X);
                //        _plotData[yKey].Add(probe.Y);
                //        _plotData[zKey].Add(probe.Z);

                //        // 限制数据量
                //        if (_plotData[xKey].Count > MaxformsPlot1Points)
                //        {
                //            _plotData[xKey].RemoveAt(0);
                //            _plotData[yKey].RemoveAt(0);
                //            _plotData[zKey].RemoveAt(0);
                //        }
                //    }
                //}
            }
            catch (Exception ex)
            {
                Console.WriteLine($"数据收集失败: {ex.Message}");
            }
        }
        private void UpdateDataGridView1()
        {
            if (dataGridView1.InvokeRequired)
            {
                dataGridView1.Invoke(new Action(UpdateDataGridView1));
                return;
            }

            lock (_dataLock)
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.IsNewRow) continue;

                    string probeName = row.Cells["ProbeName"].Value?.ToString();
                    if (string.IsNullOrEmpty(probeName)) continue;

                    var probe = _probeList.FirstOrDefault(p => p.ProbeName == probeName);
                    if (probe == null) continue;

                    // 更新X,Y,Z值
                    row.Cells["X"].Value = probe.X.ToString("F3");
                    row.Cells["Y"].Value = probe.Y.ToString("F3");
                    row.Cells["Z"].Value = probe.Z.ToString("F3");
                }
            }
        }

        // 更新Plot数据
        private void UpdatePlotUISafe()
        {
            if (!_isMonitoring) return;

            if (formsPlot1.InvokeRequired)
            {
                formsPlot1.BeginInvoke(new Action(UpdatePlot1Display));
            }
            else
            {
                UpdatePlot1Display();
            }
        }

        private void UpdatePlot1Display()//刷新formsPlot1线段显示隐藏
        {
            try
            {
                formsPlot1.Plot.Clear();

                foreach (CheckableItem item in listBox1.Items)
                {
                    if (!item.Checked) continue;

                    if (_plotData.TryGetValue(item.Text, out var dataRef))
                    {
                        double[] yData;
                        lock (_dataLock)
                        {
                            yData = dataRef.GetData();
                        }

                        if (yData.Length == 0) continue;

                        //double[] xData = Enumerable.Range(0, yData.Length).Select(x => (double)x).ToArray();
                        double[] xData = new double[yData.Length];
                        for (int i = 0; i < yData.Length; i++)
                        {
                            xData[i] = i * dataCollectionInterval / 1000.0; // 每点间隔 0.1 秒
                        }

                        var signal = formsPlot1.Plot.Add.SignalXY(xData, yData);
                        signal.LegendText = item.Text;
                        signal.Color = GetScottPlotColor(item.DisplayIndex);
                    }
                }

                //// 快速获取选中项
                //var selectedItems = new List<string>();
                //for (int i = 0; i < checkedListBox1.Items.Count; i++)
                //{
                //    if (checkedListBox1.GetItemChecked(i))
                //    {
                //        selectedItems.Add(checkedListBox1.Items[i].ToString());
                //    }
                //}

                //// 使用字典存储数据快照
                //var dataSnapshot = new Dictionary<string, double[]>();

                //// 快速获取数据（尽量减少锁的时间）
                //lock (_dataLock)
                //{
                //    foreach (var itemName in selectedItems)
                //    {
                //        if (_plotData.ContainsKey(itemName))
                //        {
                //            dataSnapshot[itemName] = _plotData[itemName].GetData();
                //        }
                //    }
                //}

                //// 绘制图表
                //foreach (var itemName in selectedItems)
                //{
                //    if (dataSnapshot.ContainsKey(itemName) && dataSnapshot[itemName].Length > 0)
                //    {
                //        double[] yData = dataSnapshot[itemName];
                //        //double[] xData = Enumerable.Range(0, yData.Length).Select(x => (double)x).ToArray();

                //        var scatter = formsPlot1.Plot.Add.Signal(yData);
                //        scatter.LegendText = itemName;
                //    }
                //}

                formsPlot1.Plot.Axes.AutoScale();
                formsPlot1.Refresh();
            }
            catch
            {
                Console.WriteLine($"更新图表显示失败");
            }


            //lock (_dataLock)
            //{
            //    formsPlot1.Plot.Clear();

            //    // 获取当前选中的项
            //    var selectedItems = new List<string>();

            //    // 只显示勾选的曲线
            //    for (int i = 0; i < checkedListBox1.Items.Count; i++)
            //    {
            //        if (checkedListBox1.GetItemChecked(i))
            //        {
            //            selectedItems.Add(checkedListBox1.Items[i].ToString());


            //        }
            //    }

            //    for (int i = 0; i < selectedItems.Count; i++)
            //    {
            //        string itemName = selectedItems[i].ToString();
            //        var data = _plotData[itemName];

            //        if (data.Count > 0)
            //        {
            //            double[] xData = Enumerable.Range(0, data.Count)
            //                .Select(x => (double)x).ToArray();

            //            var scatter = formsPlot1.Plot.Add.Scatter(xData, data.ToArray());
            //            scatter.LegendText = itemName;
            //        }
            //    }

            //    formsPlot1.Plot.Axes.AutoScale();
            //    formsPlot1.Refresh();
            //}
        }



        #endregion

        #region 选项1数据采集-1实时采集-文档记录

        private string selectedFolderPath = null; // 用于保存选中文件夹的完整路径

        private FileInfo _currentExcelFile;
        private List<FileInfo> _currentExcelfileList = new List<FileInfo>(); // 存储文件列表
        private bool _isRecording = false;
        private int _sampleCount = 0;
        // 高精度计时相关变量
        private System.Threading.Timer _recordingTimer;
        private volatile int _currentSample = 0; // 使用volatile保证多线程可见性
        private DateTime _recordingStartTime;
        private readonly object _fileLock = new object(); // 文件操作锁
        //记录数据至文档内容
        private SoundPlayer _dingPlayer;

        // 新增成员变量用于角度触发模式
        private bool isAngleTriggerMode = true;
        private bool _isAngleMode = false;
        int _initialAngleS = 0;
        int _stepAngleR = 0;
        private int _recordCountN = 0;
        private int _currentRecordCount = 0;
        private HashSet<int> _targetAngles = new HashSet<int>();
        private readonly object _angleLock = new object();

        //private void button_11_2_Click(object sender, EventArgs e)
        //{
        //    string _excelDataFolderPath = Path.Combine(System.Windows.Forms.Application.StartupPath, "ExcelData");
        //    if (!Directory.Exists(_excelDataFolderPath))
        //    {
        //        Directory.CreateDirectory(_excelDataFolderPath);
        //    }

        //    using (var saveFileDialog = new SaveFileDialog())
        //    {
        //        saveFileDialog.Filter = "Excel文件|*.xlsx";
        //        saveFileDialog.InitialDirectory = _excelDataFolderPath;
        //        saveFileDialog.Title = "创建新记录文件";
        //        string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
        //        saveFileDialog.FileName = timestamp + ".xlsx"; // 默认文件名

        //        if (saveFileDialog.ShowDialog() == DialogResult.OK)
        //        {
        //            _currentExcelFile = new FileInfo(saveFileDialog.FileName);
        //            CreateNewExcelFile(_currentExcelFile);

        //            // 添加到文件列表
        //            AddExcelFileToList(_currentExcelFile);

        //            MessageBox.Show($"已创建新文件: {_currentExcelFile.Name}");
        //        }

        //        button_11_4.Enabled = true;
        //        button_11_5.Enabled = false;
        //    }
        //}
        private void button_11_2_Click(object sender, EventArgs e)
        {
            var dialog = new VistaFolderBrowserDialog
            {
                Description = "请选择一个文件夹",
                UseDescriptionForTitle = true, // 将描述显示在标题栏
                SelectedPath = selectedFolderPath ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) // 初始路径
            };

            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                string folderPath = dialog.SelectedPath;

                // 如果路径不存在（用户输入了新路径），则创建
                if (!Directory.Exists(folderPath))
                {
                    try
                    {
                        Directory.CreateDirectory(folderPath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"无法创建文件夹：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                // 保存路径
                selectedFolderPath = folderPath;
                CalculateDataPath = selectedFolderPath;

                // 设置 textBox_11_3 的 Text 为文件夹名称
                textBox_11_3.Text = Path.GetFileName(folderPath);

                // 清空 checkedListBox2 并加载 Excel 文件
                checkedListBox2.Items.Clear();

                var excelExtensions = new[] { ".xls", ".xlsx", ".xlsm" };
                var excelFiles = Directory.GetFiles(folderPath)
                    .Where(file => excelExtensions.Contains(Path.GetExtension(file).ToLowerInvariant()))
                    .Select(Path.GetFileName)
                    .ToArray();

                foreach (string fileName in excelFiles)
                {
                    checkedListBox2.Items.Add(fileName);
                }

                // 无 Excel 文件
                if (excelFiles.Length == 0)
                {
                    // MessageBox.Show("所选文件夹中没有 Excel 文件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        //private void button_11_3_Click(object sender, EventArgs e)
        //{
        //    string _excelDataFolderPath = Path.Combine(System.Windows.Forms.Application.StartupPath, "ExcelData");
        //    if (!Directory.Exists(_excelDataFolderPath))
        //    {
        //        Directory.CreateDirectory(_excelDataFolderPath);
        //    }
        //    using (var openFileDialog = new OpenFileDialog())
        //    {
        //        openFileDialog.Filter = "Excel文件|*.xlsx";
        //        openFileDialog.Title = "选择记录文件";
        //        openFileDialog.InitialDirectory = _excelDataFolderPath;

        //        if (openFileDialog.ShowDialog() == DialogResult.OK)
        //        {
        //            _currentExcelFile = new FileInfo(openFileDialog.FileName);

        //            // 添加到文件列表
        //            AddExcelFileToList(_currentExcelFile);

        //            MessageBox.Show($"已选择文件: {_currentExcelFile.FullName}");
        //        }

        //        button_11_4.Enabled = true;
        //        button_11_5.Enabled = false;
        //    }
        //}

        private void button_11_4_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(selectedFolderPath) || !Directory.Exists(selectedFolderPath))
            {
                MessageBox.Show("请先选择一个有效的文件夹。");
                return;
            }

            if (tabControl11.SelectedIndex == 0) // 定时周期模式
            {
                isAngleTriggerMode = false;
            }
            else if (tabControl11.SelectedIndex == 1) // 角度触发模式
            {
                isAngleTriggerMode = true;
                _initialAngleS = (int)numericUpDown11.Value;
                _stepAngleR = (int)numericUpDown12.Value;
                _recordCountN = (int)numericUpDown13.Value;

                //// 参数验证
                if (_initialAngleS < 0 || _initialAngleS >= 360)
                {
                    MessageBox.Show("初始角度S必须在0-359之间");
                    return;
                }
                if (_stepAngleR <= 0 || _stepAngleR >= 360)
                {
                    MessageBox.Show("步进角度R必须在1-359之间");
                    return;
                }
                if (_recordCountN <= 0)
                {
                    MessageBox.Show("记录次数N必须大于0");
                    return;
                }
            }
            else
            {
                MessageBox.Show("未知的记录模式");
                return;
            }

            // 创建自定义小窗口
            var form = new Form
            {
                Text = "Excel 文件操作",
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                Width = 380,
                Height = 180,
                ShowIcon = false,
                ShowInTaskbar = false
            };

            // 输入框
            var textBox = new TextBox
            {
                Location = new Point(15, 20),
                Width = 200,
                Height = 23,
                Text = "新建报表" // 默认名称
            };

            // “确定”按钮
            var btnConfirm = new Button
            {
                Text = "确定",
                Location = new Point(225, 19),
                Width = 60,
                Height = 25
            };

            // “打开”图标按钮
            var btnOpen = new Button
            {
                Text = "📂", // 可用图标字体
                Font = new Font("Segoe UI Emoji", 12),
                Location = new Point(290, 19),
                Width = 35,
                Height = 25,
            };

            // 勾选框
            var chkAppend = new System.Windows.Forms.CheckBox
            {
                Text = "添加至上一个文件？",
                Location = new Point(15, 55),
                Width = 150
            };

            // 下拉框
            var cmbExistingFiles = new ComboBox
            {
                Location = new Point(15, 80),
                Width = 240,
                Height = 25,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Enabled = false // 初始禁用
            };



            // 加载当前文件夹下的 Excel 文件到 ComboBox
            void LoadExcelFilesToComboBox()
            {
                cmbExistingFiles.Items.Clear();
                var excelFiles = Directory.GetFiles(selectedFolderPath, "*.xlsx")
                    .Select(Path.GetFileName)
                    .OrderBy(f => f)
                    .ToArray();

                if (excelFiles.Length > 0)
                {
                    cmbExistingFiles.Items.AddRange(excelFiles);
                    // 尝试默认选中 _currentExcelFile（如果存在且在列表中）
                    if (_currentExcelFile != null && _currentExcelFile.DirectoryName == selectedFolderPath)
                    {
                        int idx = Array.IndexOf(excelFiles, _currentExcelFile.Name);
                        if (idx >= 0)
                            cmbExistingFiles.SelectedIndex = idx;
                        else if (cmbExistingFiles.Items.Count > 0)
                            cmbExistingFiles.SelectedIndex = 0; // 或保持 -1
                    }
                    else if (cmbExistingFiles.Items.Count > 0)
                    {
                        cmbExistingFiles.SelectedIndex = 0;
                    }
                }
                else
                {
                    cmbExistingFiles.Items.Add("(无 Excel 文件)");
                    cmbExistingFiles.SelectedIndex = 0;
                    cmbExistingFiles.Enabled = false; // 无文件时仍禁用
                }
            }

            // CheckBox 状态切换
            chkAppend.CheckedChanged += (s, ev) =>
            {
                bool appendMode = chkAppend.Checked;
                textBox.Enabled = !appendMode;
                cmbExistingFiles.Enabled = appendMode && cmbExistingFiles.Items.Count > 0;

                if (appendMode)
                {
                    LoadExcelFilesToComboBox();
                }
            };





            // === 事件绑定 ===

            btnConfirm.Click += (s, ev) =>
            {
                if (chkAppend.Checked)
                {
                    if (cmbExistingFiles.SelectedItem == null ||
                cmbExistingFiles.SelectedItem.ToString() == "(无 Excel 文件)")
                    {
                        MessageBox.Show("请选择一个有效的 Excel 文件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    string selectedFileName = cmbExistingFiles.SelectedItem.ToString();
                    string fullPath = Path.Combine(selectedFolderPath, selectedFileName);
                    _currentExcelFile = new FileInfo(fullPath);
                    _currentExcelFile.Refresh(); // 确保 Exists 正确

                    form.DialogResult = DialogResult.OK;
                    form.Close();
                }
                else
                {
                    string inputName = textBox.Text.Trim();
                    if (string.IsNullOrEmpty(inputName))
                    {
                        MessageBox.Show("请输入文件名。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // 自动补全 .xlsx 后缀（如果没写）
                    if (!inputName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) &&
                        !inputName.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                    {
                        inputName += ".xlsx";
                    }

                    // 检查非法字符
                    if (inputName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                    {
                        MessageBox.Show("文件名包含非法字符。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    string fullPath = Path.Combine(selectedFolderPath, inputName);
                    if (File.Exists(fullPath))
                    {
                        MessageBox.Show("该文件已存在，请换一个名称。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    // 新建文件方法
                    var fileInfo = new FileInfo(fullPath);
                    CreateNewExcelFile(fileInfo);
                    checkedListBox2.Items.Add(inputName);

                    form.DialogResult = DialogResult.OK;
                    form.Close();
                }


            };

            btnOpen.Click += (s, ev) =>
            {
                // 打开文件方法
                SelectExistingExcelFile();

                form.DialogResult = DialogResult.OK;
                form.Close();
            };

            // 回车键支持（在输入框按回车 = 点击新建）
            textBox.KeyDown += (s, ev) =>
            {
                if (ev.KeyCode == Keys.Enter && !chkAppend.Checked)
                {
                    btnConfirm.PerformClick();
                    ev.SuppressKeyPress = true;
                }
            };

            // 初始加载一次文件列表（虽未启用，但可预加载）
            LoadExcelFilesToComboBox();

            // 添加控件（注意顺序：后添加的在上层）
            form.Controls.AddRange(new Control[]
            {
        textBox, btnConfirm, btnOpen,
        chkAppend, cmbExistingFiles
            });

            form.DialogResult = DialogResult.Cancel;
            // 显示模态窗口
            DialogResult result = form.ShowDialog(this);

            if (result == DialogResult.OK)
            {
                if (_currentExcelFile != null && _currentExcelFile.Exists)
                {
                    StartWritingExcel();
                }
            }

        }

        private void button_11_5_Click(object sender, EventArgs e)
        {
            StopRecording();
            MessageBox.Show($"记录已暂停，共记录{_sampleCount}条数据");

            UpdateButton_11_States(false);
        }
        private void UpdateButton_11_States(bool bl)
        {
            button_11_1.Enabled = !bl;
            button_12_1.Enabled = !bl;
            button_11_2.Enabled = !bl;
            //button_11_3.Enabled = !bl;
            button_11_4.Enabled = !bl;
            button_11_5.Enabled = bl;
            button_12_003.Enabled = !bl;

            if (bl)
            {
                button_11_4.Text = "记录中...";
            }
            else
            {
                button_11_4.Text = "开始记录";
            }
        }

        // 创建新Excel文件
        private void CreateNewExcelFile(FileInfo file)
        {
            using (var package = new ExcelPackage(file))
            {
                var worksheet = package.Workbook.Worksheets.Add("数据");

                // 第一行：Title
                worksheet.Cells[1, 1].Value = "Title:" + Path.GetFileNameWithoutExtension(file.Name);
                //worksheet.Cells[1, 2].Value = Path.GetFileNameWithoutExtension(file.Name);

                // 第二行：Comment
                worksheet.Cells[2, 1].Value = "Comment:";
                //worksheet.Cells[2, 2].Value = "";

                // 第三行：Date and Time
                worksheet.Cells[3, 1].Value = "Date and Time:" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff");
                //worksheet.Cells[3, 2].Value = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff");

                // 第五行：探头通道标题（每个通道占2列）
                int col = 2; // 从第二列开始（跳过Sample列）
                for (int i = 0; i < _probeList.Count; i++)
                {
                    int probeNum = i + 1;
                    worksheet.Cells[5, col].Value = $"channel {probeNum}01"; // X
                    worksheet.Cells[5, col + 2].Value = $"channel {probeNum}02"; // Y
                    worksheet.Cells[5, col + 4].Value = $"channel {probeNum}03"; // Z
                    col += 6; // 每个探头占6列（XYZ各2列）
                }

                // 第六行：数据列标题
                worksheet.Cells[6, 1].Value = "Sample";
                col = 2;
                for (int i = 0; i < _probeList.Count; i++)
                {
                    worksheet.Cells[6, col++].Value = "XData";
                    worksheet.Cells[6, col++].Value = "Time";
                    worksheet.Cells[6, col++].Value = "YData";
                    worksheet.Cells[6, col++].Value = "Time";
                    worksheet.Cells[6, col++].Value = "ZData";
                    worksheet.Cells[6, col++].Value = "Time";
                }

                // 自动调整列宽
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                package.Save();
            }

            file.Refresh();
            _currentExcelFile = file;
        }

        //打开已有excel文件
        private void SelectExistingExcelFile()
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel文件|*.xlsx";
                openFileDialog.Title = "选择记录文件";
                openFileDialog.InitialDirectory = selectedFolderPath;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    _currentExcelFile = new FileInfo(openFileDialog.FileName);

                    MessageBox.Show($"已选择文件: {_currentExcelFile.FullName}");
                }

                //button_11_4.Enabled = true;
                //button_11_5.Enabled = false;
            }
        }

        //开始记录数据到excel
        private void StartWritingExcel()
        {
            if (_currentExcelFile == null)
            {
                MessageBox.Show("请先创建或选择Excel文件");
                return;
            }
            _currentExcelFile.Refresh(); // 强制刷新状态

            if (!_currentExcelFile.Exists)
            {
                MessageBox.Show("读取文件路径失败");
                return;
            }

            if (_HSClients.Count == 0)
            {
                MessageBox.Show("请先连接探头");
                return;
            }

            ClearFormsPlot1();//清空图像

            // 判断当前选中页
            if (!isAngleTriggerMode) // 定时周期模式
            {
                StartTimedRecordingMode();
            }
            else  // 角度触发模式
            {
                StartAngleTriggerMode();
            }

            UpdateButton_11_States(true);
        }








        // 定时周期模式
        private void StartTimedRecordingMode()
        {
            if (!int.TryParse(textBox_11_1.Text, out int intervalMs) || intervalMs <= 0)
            {
                MessageBox.Show("请输入有效的时间间隔(秒)");
                return;
            }
            intervalMs = intervalMs * 1000;//转毫秒

            _isAngleMode = false;
            _isRecording = true;
            _currentSample = GetLastSampleCount();
            _recordingStartTime = DateTime.Now;

            // 立即记录第一条数据
            RecordData(TimeSpan.Zero);

            // 启动定时器
            _recordingTimer = new System.Threading.Timer(
                callback: TimedRecordingCallback,
                state: null,
                dueTime: intervalMs,
                period: intervalMs);

            button_11_4.Enabled = false;
            button_11_5.Enabled = true;

            MessageBox.Show("已启动定时周期记录模式");
        }
        private void TimedRecordingCallback(object state)// 定时周期模式回调
        {
            if (!_isRecording || _isAngleMode) return;

            TimeSpan elapsed = DateTime.Now - _recordingStartTime;
            RecordData(elapsed);
        }

        // 角度触发模式
        private void StartAngleTriggerMode()
        {
            // 获取角度参数
            //int _initialAngleS = (int)numericUpDown11.Value;
            //int _stepAngleR = (int)numericUpDown12.Value;
            //_recordCountN = (int)numericUpDown13.Value;

            //// 参数验证
            //if (_initialAngleS < 0 || _initialAngleS >= 360)
            //{
            //    MessageBox.Show("初始角度S必须在0-359之间");
            //    return;
            //}
            //if (_stepAngleR <= 0 || _stepAngleR >= 360)
            //{
            //    MessageBox.Show("步进角度R必须在1-359之间");
            //    return;
            //}
            //if (_recordCountN <= 0)
            //{
            //    MessageBox.Show("记录次数N必须大于0");
            //    return;
            //}

            _isAngleMode = true;
            _isRecording = true;
            _currentSample = GetLastSampleCount();
            _currentRecordCount = 0;
            _recordingStartTime = DateTime.Now;

            // 预计算所有目标角度
            _targetAngles.Clear();
            for (int i = 0; i < _recordCountN; i++)
            {
                int targetAngle = (_initialAngleS + i * _stepAngleR) % 360;
                _targetAngles.Add(targetAngle);
            }

            button_11_4.Enabled = false;
            button_11_5.Enabled = true;

            Console.Write($"已启动角度触发模式，将在以下角度记录数据：\n{string.Join("、", _targetAngles.OrderBy(a => a))}");
        }

        // 外部角度值传入方法（需要在角度更新时调用）
        public void OnAngleChanged(int currentAngle)
        {
            //if (!_isRecording || !_isAngleMode) return;

            CircularGauge1set(currentAngle);//仪表盘角度

            lock (_angleLock)
            {
                // 角度归一化到0-359
                int normalizedAngle = currentAngle % 360;
                if (normalizedAngle < 0) normalizedAngle += 360;

                // 检查是否为目标角度且未记录过
                if (_targetAngles.Contains(normalizedAngle) && _currentRecordCount < _recordCountN)
                {
                    TimeSpan elapsed = DateTime.Now - _recordingStartTime;
                    RecordAngleData(elapsed, normalizedAngle);

                    // 从目标集合中移除已记录的角度
                    _targetAngles.Remove(normalizedAngle);
                    _currentRecordCount++;

                    // 更新界面显示
                    this.BeginInvoke((MethodInvoker)delegate
                    {
                        textBox_11_2.Text = _currentRecordCount.ToString();
                    });

                    // 检查是否完成所有记录
                    if (_currentRecordCount >= _recordCountN)
                    {
                        this.BeginInvoke((MethodInvoker)delegate
                        {
                            MessageBox.Show($"角度触发记录完成，共记录{_recordCountN}条数据");
                            StopRecording();
                        });
                    }
                }
            }
        }

        // 角度触发模式的数据记录
        private void RecordAngleData(TimeSpan elapsed, int triggerAngle)
        {
            lock (_fileLock)
            {
                try
                {
                    int sampleNumber = Interlocked.Increment(ref _currentSample);
                    string timeStr = $"{elapsed.Minutes:D1}:{elapsed.Seconds:D2}.{elapsed.Milliseconds:D3}";

                    this.BeginInvoke((MethodInvoker)delegate
                    {
                        try
                        {
                            using (var package = new ExcelPackage(_currentExcelFile))
                            {
                                var worksheet = package.Workbook.Worksheets["数据"];
                                int newRow = (worksheet.Dimension?.End.Row ?? 6) + 1;

                                // 记录Sample编号
                                worksheet.Cells[newRow, 1].Value = sampleNumber;

                                // 获取最新探头数据
                                var probes = _probeList.ToArray();

                                int col = 2;
                                foreach (var probe in probes)
                                {
                                    worksheet.Cells[newRow, col].Value = probe.X;
                                    worksheet.Cells[newRow, col + 1].Value = timeStr;
                                    worksheet.Cells[newRow, col + 2].Value = probe.Y;
                                    worksheet.Cells[newRow, col + 3].Value = timeStr;
                                    worksheet.Cells[newRow, col + 4].Value = probe.Z;
                                    worksheet.Cells[newRow, col + 5].Value = timeStr;
                                    col += 6;
                                }

                                //在注释列记录触发角度
                                //worksheet.Cells[newRow, worksheet.Dimension.End.Column + 1].Value = $"触发角度: {triggerAngle}°";

                                package.Save();

                                PlayDingSound();//记录音效
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"角度记录失败: {ex.Message}");
                            StopRecording();
                        }
                    });
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"角度记录后台错误: {ex.Message}");
                }
            }
        }


        // 线程安全的记录方法
        private void RecordData(TimeSpan elapsed)
        {
            lock (_fileLock)
            {
                try
                {
                    int sampleNumber = Interlocked.Increment(ref _currentSample);
                    string timeStr = $"{elapsed.Minutes:D1}:{elapsed.Seconds:D2}.{elapsed.Milliseconds:D3}";

                    this.BeginInvoke((MethodInvoker)delegate
                    {
                        try
                        {
                            using (var package = new ExcelPackage(_currentExcelFile))
                            {
                                var worksheet = package.Workbook.Worksheets["数据"];
                                int newRow = (worksheet.Dimension?.End.Row ?? 6) + 1;

                                // 确保Sample列正确递增
                                worksheet.Cells[newRow, 1].Value = sampleNumber;

                                // 获取最新探头数据（确保使用实时数据）
                                var probes = _probeList.ToArray(); // 创建数据快照

                                int col = 2;
                                foreach (var probe in probes)
                                {
                                    worksheet.Cells[newRow, col].Value = probe.X;
                                    worksheet.Cells[newRow, col + 1].Value = timeStr;
                                    worksheet.Cells[newRow, col + 2].Value = probe.Y;
                                    worksheet.Cells[newRow, col + 3].Value = timeStr;
                                    worksheet.Cells[newRow, col + 4].Value = probe.Z;
                                    worksheet.Cells[newRow, col + 5].Value = timeStr;
                                    col += 6;
                                }

                                package.Save();
                                PlayDingSound();
                            }
                            textBox_11_2.Text = sampleNumber.ToString();
                            _sampleCount = sampleNumber;
                            //textBox_11_2.Text = (_currentSample - GetLastSampleCount() + _currentRecordCount).ToString();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"记录失败: {ex.Message}");
                            StopRecording();
                        }
                    });
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"后台记录错误: {ex.Message}");
                }
            }
        }






        private void DingSound_Load()
        {
            // 初始化音效播放器
            try
            {
                string audioPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sounds", "ding.wav");
                _dingPlayer = new SoundPlayer(audioPath);
                _dingPlayer.LoadAsync(); // 预加载
                Console.WriteLine($"加载音效成功");
            }
            catch { Console.WriteLine($"加载音效失败"); }
        }
        private void PlayDingSound()
        {
            try
            {
                if (_dingPlayer != null)
                {
                    _dingPlayer.Play(); // 异步播放
                }
                else
                {
                    SystemSounds.Beep.Play(); // 回退到系统音
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"播放音效失败: {ex.Message}");
            }
        }


        private int GetLastSampleCount() // 获取文件中最后的Sample计数
        {
            if (_currentExcelFile == null || !_currentExcelFile.Exists) return 0;

            try
            {
                using (var package = new ExcelPackage(_currentExcelFile))
                {
                    var worksheet = package.Workbook.Worksheets["数据"];
                    if (worksheet == null) return 0;

                    int lastRow = worksheet.Dimension?.End.Row ?? 6;
                    return lastRow - 6; // 减去标题行
                }
            }
            catch
            {
                return 0;
            }
        }
        private void StopRecording()// 停止记录
        {
            _isRecording = false;
            _isAngleMode = false;
            _recordingTimer?.Dispose();
            _targetAngles.Clear();
        }

        //// 窗体关闭时停止记录
        //protected override void OnFormClosing(FormClosingEventArgs e)
        //{
        //    base.OnFormClosing(e);
        //    _recordTimer?.Stop();
        //    _recordTimer?.Dispose();
        //}


        #endregion

        #region 探头数据帧记录文档

        private UnifiedDataRecorder _unifiedRecorder;
        //private List<ProbeInfo> _allProbes = new List<ProbeInfo>(); // 缓存所有探头

        private void start12()
        {
            button_12_003.Enabled = true;
            button_12_004.Enabled = false;
        }


        private void button_12_003_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(selectedFolderPath) || !Directory.Exists(selectedFolderPath))
            {
                MessageBox.Show("请先选择一个有效的文件夹。");
                return;
            }
            if (_HSClients == null || _HSClients.Count == 0)
            {
                MessageBox.Show("请先连接探头");
                return;
            }

            // 确保之前的记录器已完全停止和释放
            if (_unifiedRecorder != null)
            {
                try { _unifiedRecorder.Dispose(); } catch { }
                _unifiedRecorder = null;
            }
            // 修改：直接从 _probeList 中筛选当前的 HS 探头
            // 这样可以避免 _allProbes 缓存导致的断开重连后数据不更新问题，
            // 同时也避免了因 _HSClients 中 SensorProbe 属性可能为空导致的“未检测到有效探头”错误。
            //_allProbes = _probeList
            //    .Where(p => p.DeviceType != null && p.DeviceType.StartsWith("HS"))
            //    .ToList();
            //if (_allProbes.Count == 0)
            //{
            //    MessageBox.Show("未检测到有效探头");
            //    return;
            //}

            var currentProbes = new List<ProbeInfo>();
            foreach (var client in _HSClients)
            {
                if (client.Sensor1Probe != null) currentProbes.Add(client.Sensor1Probe);
                if (client.Sensor2Probe != null) currentProbes.Add(client.Sensor2Probe);
            }

            if (currentProbes.Count == 0)
            {
                MessageBox.Show("未检测到有效探头");
                return;
            }

            try
            {
                var newRecorder = new UnifiedDataRecorder(selectedFolderPath, currentProbes);
                _unifiedRecorder = newRecorder;

                UpdateButton_12_States(true);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"启动记录失败: {ex.Message}");
                _unifiedRecorder = null;
                UpdateButton_12_States(false); // 确保按钮状态复位
            }

        }

        private void button_12_004_Click(object sender, EventArgs e)
        {
            var recorder = _unifiedRecorder;
            if (recorder != null)
            {
                _unifiedRecorder = null;
                recorder.Dispose();
                UpdateButton_12_States(false);
            }
            else
            {
                MessageBox.Show("未在记录状态");
            }
        }

        private void UpdateButton_12_States(bool bl)
        {
            button_11_1.Enabled = !bl;
            button_12_1.Enabled = !bl;
            button_11_2.Enabled = !bl;
            //button_11_3.Enabled = !bl;
            button_11_4.Enabled = !bl;
            button_12_003.Enabled = !bl;
            button_12_004.Enabled = bl;

            if (bl)
            {
                button_12_003.Text = "记录中...";
            }
            else
            {
                button_12_003.Text = "开始记录";
            }
        }

        // === 实现接口 ===
        public void OnDataReceived(string probeName, DataRecord record)
        {
            _unifiedRecorder?.SubmitData(probeName, record);
        }


        #endregion

        #endregion

        #region 选项2电源控制-2电源1/2/3/4/5/6

        private PowerTcpClientManager[] _PowerTcpManagers = new PowerTcpClientManager[6];
        private System.Threading.Timer _measurementTimer;
        private const int MeasurementInterval = 1000; // 1秒

        private void start21()
        {
            StartPowerRequestTimer();
            InitializePlot2();
            button_21_2.Enabled = false;
        }
        private void InitializePowerTcpManagers(int i)
        {
            // 获取对应的TextBox控件
            var ipTextBox = Controls.Find($"textBox_62_{i + 1}_1", true).FirstOrDefault() as System.Windows.Forms.TextBox;
            var portTextBox = Controls.Find($"textBox_62_{i + 1}_2", true).FirstOrDefault() as System.Windows.Forms.TextBox;
            var coilTextBox = Controls.Find($"textBox_62_{i + 1}_3", true).FirstOrDefault() as System.Windows.Forms.TextBox;
            var zerofieldTextBox = Controls.Find($"textBox_62_{i + 1}_4", true).FirstOrDefault() as System.Windows.Forms.TextBox;

            if (ipTextBox != null && portTextBox != null && coilTextBox != null && zerofieldTextBox != null)
            {
                string ip = ipTextBox.Text;
                int port = int.TryParse(portTextBox.Text, out port) ? port : 0;
                string coil = coilTextBox.Text;
                string zerofield = zerofieldTextBox.Text;

                _PowerTcpManagers[i] = new PowerTcpClientManager(ip, port, coil, zerofield);
                _PowerTcpManagers[i].ConnectAsync();
                //_PowerTcpManagers[i]._xyz = "关闭";
            }
        }


        private void HandleComboBoxChange(int index)
        {
            System.Windows.Forms.ComboBox currentComboBox = Controls.Find($"comboBox_21_{index}", true).FirstOrDefault() as System.Windows.Forms.ComboBox;
            if (currentComboBox == null) return;

            // 获取所有6个ComboBox
            List<System.Windows.Forms.ComboBox> allComboBoxes = new List<System.Windows.Forms.ComboBox>();
            for (int i = 1; i <= 6; i++)
            {
                var cb = Controls.Find($"comboBox_21_{i}", true).FirstOrDefault() as System.Windows.Forms.ComboBox;
                if (cb != null) allComboBoxes.Add(cb);
            }
            string selectedValue = currentComboBox.SelectedItem?.ToString();

            if (selectedValue == "关闭")
            {
                if (_PowerTcpManagers[index - 1] != null)
                {
                    _PowerTcpManagers[index - 1].SendCommand("OUTP 0\n");
                    _PowerTcpManagers[index - 1].Disconnect();
                    _PowerTcpManagers[index - 1]._xyz = selectedValue;
                }

                var voltTextBox = Controls.Find($"textBox_21_{index}_1", true).FirstOrDefault() as System.Windows.Forms.TextBox;
                voltTextBox.Text = "0";
                var currTextBox = Controls.Find($"textBox_21_{index}_2", true).FirstOrDefault() as System.Windows.Forms.TextBox;
                currTextBox.Text = "0";

            }
            else
            {
                //// 检查其他ComboBox是否已经选择了相同的值
                //bool valueAlreadySelected = allComboBoxes
                //    .Where(cb => cb != currentComboBox)
                //    .Any(cb => cb.SelectedItem?.ToString() == selectedValue);

                //if (valueAlreadySelected)
                //{
                //    // 如果值已被其他ComboBox选择，恢复当前ComboBox的选中项
                //    currentComboBox.SelectedIndex = 0;
                //    MessageBox.Show($"'{selectedValue}'已经被其他电源选择，请选择其他选项。");
                //    return;
                //}
                //else
                {
                    InitializePowerTcpManagers(index - 1);
                    if (_PowerTcpManagers[index - 1] != null)
                    {
                        _PowerTcpManagers[index - 1]._xyz = selectedValue;
                        _PowerTcpManagers[index - 1].SendCommand("CURR 0\n");
                        _PowerTcpManagers[index - 1].SendCommand("OUTP 1\n");
                    }
                }

            }

            // 检查是否有任何连接，决定是否启动定时器
            bool anyConnected = _PowerTcpManagers.Any(m => m != null && m.IsConnected);
            if (anyConnected)
            {
                _measurementTimer.Change(0, MeasurementInterval);
            }
            else
            {
                _measurementTimer.Change(Timeout.Infinite, MeasurementInterval);
            }
        }


        private void comboBox_21_1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            HandleComboBoxChange(1);
        }

        private void comboBox_21_2_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            HandleComboBoxChange(2);
        }

        private void comboBox_21_3_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            HandleComboBoxChange(3);
        }

        private void comboBox_21_4_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            HandleComboBoxChange(4);
        }

        private void comboBox_21_5_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            HandleComboBoxChange(5);
        }

        private void comboBox_21_6_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            HandleComboBoxChange(6);
        }

        private void StartPowerRequestTimer()
        {
            _measurementTimer = new System.Threading.Timer(async _ =>
            {
                await SendPeriodicRequests();
            }, null, Timeout.Infinite, MeasurementInterval); // 每隔MeasurementInterval秒执行
        }

        private async Task SendPeriodicRequests()
        {
            _measurementCount++;
            bool anyConnected = false;
            double[] latestCurrents = new double[6];//这三个管Plot2更新

            for (int i = 0; i < _PowerTcpManagers.Length; i++)
            {
                var manager = _PowerTcpManagers[i];
                if (manager == null) continue;

                if (manager.IsConnected)
                {
                    anyConnected = true;

                    // 读取电压
                    string voltageResponse = await manager.SendCommandAndReceiveResponseAsync("MEAS:VOLT?\n");
                    if (double.TryParse(voltageResponse, out double voltage))
                    {
                        manager.LatestVoltage = voltage;
                    }
                    else
                    {
                        manager.LatestVoltage = 0;
                    }

                    // 读取电流
                    string currentResponse = await manager.SendCommandAndReceiveResponseAsync("MEAS:CURR?\n");
                    if (double.TryParse(currentResponse, out double current))
                    {
                        manager.LatestCurrent = current;
                    }
                    else
                    {
                        manager.LatestCurrent = 0;
                    }

                    latestCurrents[i] = manager.LatestCurrent;
                }
                else
                {
                    latestCurrents[i] = 0;
                }
            }

            // 更新UI
            this.Invoke((MethodInvoker)delegate
            {
                for (int i = 0; i < _PowerTcpManagers.Length; i++)
                {
                    if (_PowerTcpManagers[i] == null) continue;

                    // 更新电压TextBox
                    var voltageTextBox = Controls.Find($"textBox_21_{i + 1}_1", true).FirstOrDefault() as System.Windows.Forms.TextBox;
                    voltageTextBox.Text = (_PowerTcpManagers[i].LatestVoltage.ToString("F2").TrimEnd('0').TrimEnd('.'));

                    // 更新电流TextBox
                    var currentTextBox = Controls.Find($"textBox_21_{i + 1}_2", true).FirstOrDefault() as System.Windows.Forms.TextBox;
                    currentTextBox.Text = (_PowerTcpManagers[i].LatestCurrent.ToString("F2").TrimEnd('0').TrimEnd('.'));
                }


                // 如果有连接则更新图表
                if (anyConnected)
                {
                    UpdatePlot2(_measurementCount, latestCurrents);
                }
            });



        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _measurementTimer?.Dispose();
            _demagnetizationTimer?.Stop();
            _demagnetizationTimer?.Dispose();
            _magnetizationTimer?.Stop();
            _magnetizationTimer?.Dispose();

            base.OnFormClosing(e);
            foreach (var manager in _PowerTcpManagers)
            {
                manager?.Disconnect();
            }
            foreach (var manager in _PowerTcpManagers78)
            {
                manager?.Disconnect();
            }
        }




        // 用于存储绘图数据
        private List<double>[] _currentDataSeries2 = new List<double>[6];
        private int _measurementCount = 0;
        private const int MaxPowerPoints = 100; // 最大显示数据点数

        private void InitializePlot2()
        {
            // 初始化6条数据线
            for (int i = 0; i < 6; i++)
            {
                _currentDataSeries2[i] = new List<double>();
            }

            // 初始化6条曲线，设置不同颜色
            var colors = new ScottPlot.Color[]
            {
        Colors.Red,
        Colors.Green,
        Colors.Blue,
        Colors.Orange,
        Colors.Purple,
        Colors.Brown
            };

            for (int i = 0; i < 6; i++)
            {
                var scatter = formsPlot2.Plot.Add.Scatter(new double[0], new double[0]);
                scatter.LegendText = $"{i + 1}";
                scatter.Color = colors[i];
            }

            // 显示图例
            formsPlot2.Plot.ShowLegend();
            formsPlot2.Refresh();
            formsPlot2.UserInputProcessor.Disable();
        }

        // 添加新数据的方法
        public void UpdatePlot2(int xValue, double[] yValues)
        {
            formsPlot2.Plot.Clear();

            // 添加新数据点
            for (int i = 0; i < 6; i++)
            {
                _currentDataSeries2[i].Add(yValues[i]);

                // 限制数据点数量，保持图表清晰
                if (_currentDataSeries2[i].Count > MaxPowerPoints)
                {
                    _currentDataSeries2[i].RemoveAt(0);
                }
            }

            // 准备X轴数据（从1开始的序列）
            double[] xData = Enumerable.Range(
                Math.Max(1, xValue - _currentDataSeries2[0].Count + 1),
                _currentDataSeries2[0].Count).Select(x => (double)x).ToArray();

            // 更新每条曲线
            for (int i = 0; i < 6; i++)
            {
                //var scatter = formsPlot2.Plot.Plots[i] as ScottPlot.Plots.Scatter;
                //scatter.Data.X = xData;
                //scatter.Data.Y = _currentDataSeries2[i].ToArray();

                // 新版ScottPlot的绘图方式
                var scatter = formsPlot2.Plot.Add.Scatter(xData, _currentDataSeries2[i].ToArray());
                scatter.LegendText = (i + 1).ToString();
            }

            // 自动调整坐标轴范围
            formsPlot2.Plot.Axes.AutoScale();

            // 刷新图表
            formsPlot2.Refresh();

        }



        //设置电源磁场

        private void textBox_21_7_1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (textBox_21_7_1.Text == "")
                {
                    formsPlot2.Focus();
                    return;
                }
                string input = textBox_21_7_1.Text;
                // 使用正则表达式验证是否为正整数（不含零）
                // bool isValid = System.Text.RegularExpressions.Regex.IsMatch(input, @"^[1-9]\d*$");
                bool isValid = System.Text.RegularExpressions.Regex.IsMatch(input, @"^(0|-?[1-9]\d*)$");

                if (isValid)
                {
                    int num = int.Parse(input);
                    string textxyz = comboBox_21_7.Text.Replace("-", "");
                    PowerTcpClientManager ptcm = PowerTcpFind(textxyz);
                    if (ptcm != null)
                    {
                        if (double.TryParse(ptcm._coilconstant, out double coilConstant) && coilConstant != 0)
                        {
                            double changedCurrent = (double)num / coilConstant;
                            //if (comboBox_21_7.Text.Contains("-")) changedCurrent = -changedCurrent;

                            textBox_21_8_1.Text = changedCurrent.ToString();
                            //textBox_21_8_1_Leave(textBox_21_8_1, EventArgs.Empty);
                            var args = new KeyEventArgs(Keys.Enter);
                            textBox_21_8_1_KeyDown(textBox_21_8_1, args);
                        }
                        else
                        {
                            MessageBox.Show("请添加或修改参数并重新连接");
                            tabControl1.SelectedIndex = 5;
                            tabControl5.SelectedIndex = 1;
                        }

                    }
                }
                else
                {
                    MessageBox.Show($"输入无效：{input} 不是正常数据");
                    textBox_21_7_1.Text = "";
                }
                formsPlot2.Focus();
            }
        }

        private void textBox_21_7_2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (textBox_21_7_2.Text == "")
                {
                    formsPlot2.Focus();
                    return;
                }
                string input = textBox_21_7_2.Text;

                bool isValid = System.Text.RegularExpressions.Regex.IsMatch(input, @"^(0|-?[1-9]\d*)$");

                if (isValid)
                {
                    int num = int.Parse(input);
                    string textxyz = comboBox_21_8.Text.Replace("-", "");
                    PowerTcpClientManager ptcm = PowerTcpFind(textxyz);
                    if (ptcm != null)
                    {
                        if (double.TryParse(ptcm._coilconstant, out double coilConstant) && coilConstant != 0)
                        {
                            double changedCurrent = (double)num / coilConstant;
                            //if (comboBox_21_8.Text.Contains("-")) changedCurrent = -changedCurrent;
                            textBox_21_8_2.Text = changedCurrent.ToString();
                            //textBox_21_8_2_Leave(textBox_21_8_2, EventArgs.Empty);
                            var args = new KeyEventArgs(Keys.Enter);
                            textBox_21_8_2_KeyDown(textBox_21_8_2, args);
                        }
                        else
                        {
                            MessageBox.Show("请添加或修改参数并重新连接");
                            tabControl1.SelectedIndex = 5;
                            tabControl5.SelectedIndex = 1;
                        }
                    }
                }
                else
                {
                    MessageBox.Show($"输入无效：{input} 不是正常数据");
                    textBox_21_7_2.Text = "";
                }
                formsPlot2.Focus();
            }
        }

        private void textBox_21_7_3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (textBox_21_7_3.Text == "")
                {
                    formsPlot2.Focus();
                    return;
                }
                string input = textBox_21_7_3.Text;

                bool isValid = System.Text.RegularExpressions.Regex.IsMatch(input, @"^(0|-?[1-9]\d*)$");


                if (isValid)
                {
                    int num = int.Parse(input);
                    string textxyz = comboBox_21_9.Text.Replace("-", "");
                    PowerTcpClientManager ptcm = PowerTcpFind(textxyz);
                    if (ptcm != null)
                    {
                        if (double.TryParse(ptcm._coilconstant, out double coilConstant) && coilConstant != 0)
                        {
                            double changedCurrent = (double)num / coilConstant;
                            //if (comboBox_21_9.Text.Contains("-")) changedCurrent = -changedCurrent;

                            textBox_21_8_3.Text = changedCurrent.ToString();
                            //textBox_21_8_3_Leave(textBox_21_8_3, EventArgs.Empty);
                            var args = new KeyEventArgs(Keys.Enter);
                            textBox_21_8_3_KeyDown(textBox_21_8_3, args);
                        }
                        else
                        {
                            MessageBox.Show("请添加或修改参数并重新连接");
                            tabControl1.SelectedIndex = 5;
                            tabControl5.SelectedIndex = 1;
                        }
                    }
                }
                else
                {
                    MessageBox.Show($"输入无效：{input} 不是正常数据");
                    textBox_21_7_3.Text = "";
                }
                formsPlot2.Focus();
            }
        }

        //设置电源电流
        private void textBox_21_8_1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (textBox_21_8_1.Text == "")
                {
                    formsPlot2.Focus();
                    return;
                }

                if (!double.TryParse(textBox_21_8_1.Text, out double current))
                {
                    MessageBox.Show("请输入有效的电流值。");
                    formsPlot2.Focus();
                    return;
                }

                // 取绝对值，并根据ComboBox的符号决定正负
                current = Math.Abs(current);
                if (comboBox_21_7.Text.Contains("-"))
                {
                    current = -current;
                }

                string textxyz = comboBox_21_7.Text.Replace("-", "");
                PowerTcpClientManager ptcm = PowerTcpFind(textxyz);
                if (ptcm != null)
                {
                    //ptcm.SendCommand("CURR " + textBox_21_8_1.Text + "\n");
                    ptcm.SendCommand("CURR " + current.ToString() + "\n");
                }
                else
                {
                    MessageBox.Show($"找不到对应电源TCP 恒场" + textxyz);
                }
                formsPlot2.Focus();
            }
        }

        private void textBox_21_8_2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (textBox_21_8_2.Text == "")
                {
                    formsPlot2.Focus();
                    return;
                }

                if (!double.TryParse(textBox_21_8_2.Text, out double current))
                {
                    MessageBox.Show("请输入有效的电流值。");
                    formsPlot2.Focus();
                    return;
                }

                // 取绝对值，并根据ComboBox的符号决定正负
                current = Math.Abs(current);
                if (comboBox_21_8.Text.Contains("-"))
                {
                    current = -current;
                }

                string textxyz = comboBox_21_8.Text.Replace("-", "");
                PowerTcpClientManager ptcm = PowerTcpFind(textxyz);
                if (ptcm != null)
                {
                    //ptcm.SendCommand("CURR " + textBox_21_8_2.Text + "\n");
                    ptcm.SendCommand("CURR " + current.ToString() + "\n");
                }
                else
                {
                    MessageBox.Show($"找不到对应电源TCP 恒场" + textxyz);
                }
                formsPlot2.Focus();
            }
        }

        private void textBox_21_8_3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (textBox_21_8_3.Text == "")
                {
                    formsPlot2.Focus();
                    return;
                }

                if (!double.TryParse(textBox_21_8_3.Text, out double current))
                {
                    MessageBox.Show("请输入有效的电流值。");
                    formsPlot2.Focus();
                    return;
                }

                // 取绝对值，并根据ComboBox的符号决定正负
                current = Math.Abs(current);
                if (comboBox_21_9.Text.Contains("-"))
                {
                    current = -current;
                }

                string textxyz = comboBox_21_9.Text.Replace("-", "");
                PowerTcpClientManager ptcm = PowerTcpFind(textxyz);
                if (ptcm != null)
                {
                    //ptcm.SendCommand("CURR " + textBox_21_8_3.Text + "\n");
                    ptcm.SendCommand("CURR " + current.ToString() + "\n");
                }
                else
                {
                    MessageBox.Show($"找不到对应电源TCP 恒场" + textxyz);
                }
                formsPlot2.Focus();
            }
        }
        //定位对应电源
        private PowerTcpClientManager PowerTcpFind(string sxyz)
        {
            PowerTcpClientManager ptcm = null;
            for (int i = 0; i < 6; i++)
            {
                if (_PowerTcpManagers[i] == null) continue;
                if (_PowerTcpManagers[i]._xyz == "恒场" + sxyz)
                {
                    ptcm = _PowerTcpManagers[i];
                }
            }
            return ptcm;
        }


        //发送设置零场电流
        private void button_21_1_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 6; i++)
            {
                if (_PowerTcpManagers[i] == null) continue;
                if (_PowerTcpManagers[i]._xyz.Contains("零场"))
                {
                    _PowerTcpManagers[i].SendCommand("CURR " + _PowerTcpManagers[i]._zerofieldcurrent + "\n");
                }
            }
        }

        private async void button_21_2_Click(object sender, EventArgs e)
        {
            // 禁用按钮防止重复点击
            button_21_2.Enabled = false;

            try
            {
                // 调用异步方法
                await SendPeriodicRequests();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 恢复按钮状态
                button_21_2.Enabled = true;
            }
        }






        #endregion

        #region 选项2电源控制-2电源7/8


        private PowerTcpClientManager[] _PowerTcpManagers78 = new PowerTcpClientManager[2];
        private int _MaxCount78;
        private int _measurementPoint = 40;//每秒默认点数，和波形类里的点数同步

        // 退磁电源相关变量
        private MagPowerWaveCalculator _MPWC;
        private double[] _demagnetizationWave;
        private double _demagnetizationMeasurementPoint;
        private System.Windows.Forms.Timer _demagnetizationTimer;
        private int _demagnetizationCurrentIndex = 0;
        private List<double> _demagnetizationXData = new List<double>();
        private List<double> _demagnetizationYData = new List<double>();
        // 充磁电源相关变量
        private MagPowerMagnetizationWaveCalculator _MPMWC;
        private double[] _magnetizationWave;
        private double _magnetizationMeasurementPoint;
        private System.Windows.Forms.Timer _magnetizationTimer;
        private int _magnetizationCurrentIndex = 0;
        private List<double> _magnetizationXData = new List<double>();
        private List<double> _magnetizationYData = new List<double>();

        private void start22()
        {
            InitializePlot3();
            StartPowerRequestTimer78();
            button_22_1.Enabled = false;
            button_22_2.Enabled = false;
        }

        private void InitializePowerTcpManagers78(int i)
        {
            // 获取对应的TextBox控件
            var ipTextBox = Controls.Find($"textBox_62_{i + 7}_1", true).FirstOrDefault() as System.Windows.Forms.TextBox;
            var portTextBox = Controls.Find($"textBox_62_{i + 7}_2", true).FirstOrDefault() as System.Windows.Forms.TextBox;

            if (ipTextBox != null && portTextBox != null)
            {
                string ip = ipTextBox.Text;
                int port = int.TryParse(portTextBox.Text, out port) ? port : 0;
                string coil = "0";
                string zerofield = "0";

                _PowerTcpManagers78[i] = new PowerTcpClientManager(ip, port, coil, zerofield);
                _PowerTcpManagers78[i].ConnectAsync();
            }
        }


        private void comboBox_22_1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedValue = comboBox_22_1.SelectedItem?.ToString();

            if (selectedValue == "电源7")
            {
                if (_PowerTcpManagers78[1] != null && _PowerTcpManagers78[1].IsConnected)
                {
                    _PowerTcpManagers78[1].SendCommand("OUTP 0\n");
                    _PowerTcpManagers78[1].Disconnect();
                }

                InitializePowerTcpManagers78(0);
                if (_PowerTcpManagers78[0].IsConnected)
                {
                    //_PowerTcpManagers78[2]._xyz = selectedValue;
                    //_PowerTcpManagers78[0].SendCommand("OUTP 1\n");
                    button_22_1.Enabled = true;
                }
                else
                {
                    comboBox_22_1.SelectedIndex = 0;
                }
            }
            else if (selectedValue == "电源8")
            {
                if (_PowerTcpManagers78[0] != null && _PowerTcpManagers78[0].IsConnected)
                {
                    _PowerTcpManagers78[0].SendCommand("OUTP 0\n");
                    _PowerTcpManagers78[0].Disconnect();
                }

                InitializePowerTcpManagers78(1);

                if (_PowerTcpManagers78[1].IsConnected)
                {
                    //_PowerTcpManagers78[2]._xyz = selectedValue;
                    //_PowerTcpManagers78[1].SendCommand("OUTP 1\n");
                    button_22_1.Enabled = true;
                }
                else
                {
                    comboBox_22_1.SelectedIndex = 0;
                }
            }
            else
            {
                foreach (var manager in _PowerTcpManagers78)
                {
                    if (manager != null && manager.IsConnected)
                    {
                        manager.SendCommand("OUTP 0\n");
                        manager.Disconnect();
                    }
                }
                button_22_1.Enabled = false;
            }
        }


        private void StartPowerRequestTimer78()
        {
            // 初始化退磁电源定时器
            _demagnetizationTimer = new System.Windows.Forms.Timer();
            _demagnetizationTimer.Tick += DemagnetizationTimer_Tick;

            // 初始化充磁电源定时器
            _magnetizationTimer = new System.Windows.Forms.Timer();
            _magnetizationTimer.Tick += MagnetizationTimer_Tick;
        }

        private void InitializePlot3()
        {
            var scatter = formsPlot3.Plot.Add.Scatter(new double[0], new double[0]);

            formsPlot3.Refresh();
            formsPlot3.UserInputProcessor.Disable();
        }

        private async void button_22_1_Click(object sender, EventArgs e)
        {
            // 禁用按钮防止重复点击
            button_22_1.Enabled = false;

            try
            {
                if (tabControl23.SelectedIndex == 0)
                {
                    // 退磁电源
                    Demagnetization();
                }
                else if (tabControl23.SelectedIndex == 1)
                {
                    // 充磁电源
                    Magnetizing();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 恢复按钮状态
                //button_22_1.Enabled = true;

                button_able(true);
            }
        }


        #region 退磁电源方法
        private async void Demagnetization()
        {
            _magnetizationTimer?.Stop();

            _MPWC = new MagPowerWaveCalculator();
            _measurementPoint = _MPWC.PointPreScd;

            double A = 1;
            if (double.TryParse(textBox_22_1.Text, out double value1))
            {
                A = value1;
                _MPWC.DemagnetizationAmplitude = A;
            }
            double K = 1;
            if (double.TryParse(textBox_22_2.Text, out double value2))
            {
                K = value2;
                _MPWC.DemagnetizationDecayCoefficient = K;
            }
            double T = 1;
            if (double.TryParse(textBox_22_3.Text, out double value3))
            {
                double f = value3;
                T = 1 / f;
                _MPWC.DemagnetizationPeriod = T;
            }

            //计算IL值
            if (int.TryParse(textBox_22_4.Text, out int value4))
            {
                _MPWC.TimeTotal = value4;
                _MaxCount78 = _MPWC.GetTimeTotal();
                if (_MaxCount78 > _MPWC.PointMax)
                {
                    _MaxCount78 = _MPWC.PointMax;
                }
            }

            if (radioButton_22_1.Checked)
            {
                if (radioButton_22_3.Checked)
                {
                    _MPWC.DemagnetizationWaveType = 0;
                }
                else
                {
                    _MPWC.DemagnetizationWaveType = 2;
                }
            }
            else
            {
                if (radioButton_22_3.Checked)
                {
                    _MPWC.DemagnetizationWaveType = 1;
                }
                else
                {
                    _MPWC.DemagnetizationWaveType = 3;
                }
            }

            _demagnetizationWave = _MPWC.GetDemagnetizationWave();
            _demagnetizationMeasurementPoint = _MPWC.TimeInterval();
            //Console.Write(_demagnetizationWave);


            for (int i = 0; i < _PowerTcpManagers78.Length; i++)
            {
                //continue;
                var manager = _PowerTcpManagers78[i];
                if (manager == null) continue;

                if (manager.IsConnected)
                {
                    string ERRstring = await manager.SendCommandAndReceiveResponseAsync("SYST:ERR?");
                    if (ERRstring == null || ERRstring == "")
                    {
                        MessageBox.Show(ERRstring);
                        return;
                    }
                    await Task.Delay(100);

                    //发送命令获取校准参数
                    string responseCALIbrate = await manager.SendCommandAndReceiveResponseAsync("CALIbrate:CURRent:PARAmeter?");
                    if (string.IsNullOrEmpty(responseCALIbrate))
                    {
                        MessageBox.Show("获取校准参数失败");
                        return;
                    }
                    //处理返回的数据
                    string[] values = responseCALIbrate.Split(',');
                    if (values.Length < 2)
                    {
                        MessageBox.Show("返回数据格式不正确");
                        return;
                    }
                    //解析数组数量
                    if (!int.TryParse(values[0], out int arrayCount))
                    {
                        MessageBox.Show("解析数组数量失败");
                        return;
                    }
                    //初始化数组
                    float[] real_data = new float[arrayCount + 1];
                    float[] set_dac_data = new float[arrayCount + 1];
                    real_data[0] = 0;
                    set_dac_data[0] = 0;

                    //填充数组
                    for (int j = 0; j < arrayCount; j++)
                    {
                        if (!float.TryParse(values[j * 2 + 1], out real_data[j + 1]) ||
                            !float.TryParse(values[j * 2 + 2], out set_dac_data[j + 1]))
                        {
                            MessageBox.Show($"解析第{j + 1}组数据失败");
                            return;
                        }
                    }

                    await Task.Delay(100);

                    //确保电源处于非输出状态
                    manager.SendCommand("OUTP 0");

                    await Task.Delay(100);

                    //设置功能模式为CC
                    manager.SendCommand("FUNC:MODE CC");
                    await Task.Delay(100);
                    //周期运行设定值
                    manager.SendCommand("FUNC:RUN 0");

                    await Task.Delay(100);

                    //计算IL值
                    //if (!int.TryParse(textBox_22_4.Text, out int textBoxValue))
                    //{
                    //    MessageBox.Show("请输入有效的数字");
                    //    return;
                    //}
                    //int IL = textBoxValue * 10 + 1;
                    //if (IL > 4096)
                    //{
                    //    IL = 4096;
                    //}

                    //设置采样率
                    manager.SendCommand($"FUNC:SAMPle {_MaxCount78}");

                    await Task.Delay(100);
                    //设置功能长度
                    manager.SendCommand($"FUNC:LENgth {_MaxCount78}");

                    await Task.Delay(100);
                    //设置参数为电流
                    manager.SendCommand("FUNC:PARAmeter CURR");

                    await Task.Delay(100);
                    manager.SendCommand("FUNC:STRIng LIST");

                    //计算所有区间的a和b值
                    List<(double a, double b)> coefficients = CalculateCoefficients(real_data, set_dac_data);

                    List<double> ldy = new List<double>();//设计的电流值
                    List<double> ldx = new List<double>();//实际转换后的电流值
                    List<int> myf = new List<int>();//记录电流0为正1为负
                    StringBuilder CURRhex = new StringBuilder();//电流值十六进制
                    StringBuilder TIMEhex = new StringBuilder();//时间差十六进制

                    for (int j = 0; j < _MaxCount78; j++)
                    {
                        double WC = _demagnetizationWave[j];
                        //Debug.WriteLine("函数my值：" + WC);
                        if (WC < 0) { myf.Add(1); }
                        else { myf.Add(0); }
                        ldy.Add(Math.Abs(WC));
                    }
                    foreach (double my in ldy)
                    {
                        //根据my值找到对应的区间和系数
                        var (a, b) = FindCoefficientsForValue(my, real_data, coefficients);

                        //计算mx值
                        double mx = (my - b) / a;
                        ldx.Add(mx);
                        //Debug.WriteLine("计算结果mx值：" + mx);

                        int intValue = (int)(mx * 65535 / (60 * 1.02));
                        string hexValue = intValue.ToString("X4"); // 4位16进制
                        //Debug.WriteLine("* 65535/(60 * 1.02)转16进制结果：" + hexValue);
                        CURRhex.Append(hexValue);
                    }
                    foreach (int my in myf)
                    {
                        string hexValue = TimeToHex(_demagnetizationMeasurementPoint, my);
                        TIMEhex.Append(hexValue);
                    }


                    //分段发送CURRhex
                    bool allSegmentsSentSuccessfully = await SendHexSegments(manager, CURRhex);

                    if (!allSegmentsSentSuccessfully)
                    {
                        manager.SendCommand("FUNC:END");
                        button_22_1.Enabled = true;
                        Debug.WriteLine("电流值发送多次未成功，跳出");
                        return;
                    }

                    //发送成功，继续发送TIMEhex
                    manager.SendCommand("FUNC:PARAmeter TIME");
                    await Task.Delay(100);


                    //分段发送TIMEhex
                    bool allTimeSegmentsSentSuccessfully = await SendHexSegments(manager, TIMEhex);

                    if (allTimeSegmentsSentSuccessfully)
                    {
                        manager.SendCommand("FUNC:END");
                        await Task.Delay(100);
                        manager.SendCommand("OUTPUT ON");
                    }
                    else
                    {
                        manager.SendCommand("FUNC:END");
                        button_22_1.Enabled = true;
                        Debug.WriteLine("时间值发送多次未成功，跳出");
                        return;
                    }
                    await Task.Delay(100);
                    //设置电源为输出状态，输出上述设定
                    manager.SendCommand("OUTP 1");

                    await Task.Delay(100);
                }
            }

            // 重置状态
            _demagnetizationCurrentIndex = 0;
            _demagnetizationXData.Clear();
            _demagnetizationYData.Clear();
            // 清空图表
            formsPlot3.Plot.Clear();
            // 设置定时器间隔
            int intervalMs = (int)(_demagnetizationMeasurementPoint * 1000);
            _demagnetizationTimer.Interval = Math.Max(1, intervalMs); // 确保至少1ms
            // 启动定时器
            _demagnetizationTimer.Start();
        }
        private void DemagnetizationTimer_Tick(object sender, EventArgs e)
        {
            UpdateDemagnetizationPlot();
        }

        private void UpdateDemagnetizationPlot()
        {
            if (_demagnetizationCurrentIndex < _demagnetizationWave.Length)
            {
                // 添加新数据点
                _demagnetizationXData.Add((double)_demagnetizationCurrentIndex / _measurementPoint);
                _demagnetizationYData.Add(_demagnetizationWave[_demagnetizationCurrentIndex]);

                // 更新图表
                formsPlot3.Plot.Clear();
                var scatter = formsPlot3.Plot.Add.Scatter(_demagnetizationXData.ToArray(), _demagnetizationYData.ToArray());

                formsPlot3.Plot.Axes.AutoScale();
                formsPlot3.Refresh();
                _demagnetizationCurrentIndex++;
            }
            else
            {
                // 数据已全部绘制完成，停止定时器
                _demagnetizationTimer.Stop();
                AutoTriggerStopButton();
            }
        }
        #endregion

        #region 充磁电源方法
        private async void Magnetizing()
        {
            _demagnetizationTimer?.Stop();

            _MPMWC = new MagPowerMagnetizationWaveCalculator();
            _measurementPoint = _MPMWC.PointPreScd;

            double T = 1;
            if (double.TryParse(textBox_22_5.Text, out double value1))
            {
                T = value1;
                _MPMWC.MagHoldTime = T;
            }
            double B = 1;
            if (double.TryParse(textBox_22_6.Text, out double value2))
            {
                B = value2;
                _MPMWC.MagMax = B;
            }
            double K = 1;
            if (double.TryParse(textBox_22_7.Text, out double value3))
            {
                K = value3;
                _MPMWC.MagGradient = K;
            }
            double BI = 1;
            if (int.TryParse(textBox_62_9_1.Text, out int value4))
            {
                BI = value4;
                _MPMWC.MagBI = BI;
            }

            //计算IL值
            if (int.TryParse(textBox_22_5.Text, out int value5))
            {
                _MaxCount78 = _MPMWC.GetTimeTotal();
                if (_MaxCount78 > _MPMWC.PointMax)
                {
                    _MaxCount78 = _MPMWC.PointMax;
                }
            }

            _magnetizationWave = _MPMWC.GetMagnetizationWave();
            _magnetizationMeasurementPoint = _MPMWC.TimeInterval();


            for (int i = 0; i < _PowerTcpManagers78.Length; i++)
            {
                //continue;
                var manager = _PowerTcpManagers78[i];
                if (manager == null) continue;

                if (manager.IsConnected)
                {
                    string ERRstring = await manager.SendCommandAndReceiveResponseAsync("SYST:ERR?");
                    if (ERRstring == null || ERRstring == "")
                    {
                        //MessageBox.Show(ERRstring);
                        return;
                    }
                    await Task.Delay(100);

                    //发送命令获取校准参数
                    string responseCALIbrate = await manager.SendCommandAndReceiveResponseAsync("CALIbrate:CURRent:PARAmeter?");
                    if (string.IsNullOrEmpty(responseCALIbrate))
                    {
                        MessageBox.Show("获取校准参数失败");
                        return;
                    }
                    //处理返回的数据
                    string[] values = responseCALIbrate.Split(',');
                    if (values.Length < 2)
                    {
                        MessageBox.Show("返回数据格式不正确");
                        return;
                    }
                    //解析数组数量
                    if (!int.TryParse(values[0], out int arrayCount))
                    {
                        MessageBox.Show("解析数组数量失败");
                        return;
                    }
                    //初始化数组
                    float[] real_data = new float[arrayCount + 1];
                    float[] set_dac_data = new float[arrayCount + 1];
                    real_data[0] = 0;
                    set_dac_data[0] = 0;

                    //填充数组
                    for (int j = 0; j < arrayCount; j++)
                    {
                        if (!float.TryParse(values[j * 2 + 1], out real_data[j + 1]) ||
                            !float.TryParse(values[j * 2 + 2], out set_dac_data[j + 1]))
                        {
                            MessageBox.Show($"解析第{j + 1}组数据失败");
                            return;
                        }
                    }

                    await Task.Delay(100);

                    //确保电源处于非输出状态
                    manager.SendCommand("OUTP 0");

                    await Task.Delay(100);

                    //设置功能模式为CC
                    manager.SendCommand("FUNC:MODE CC");
                    await Task.Delay(100);
                    //周期运行设定值
                    manager.SendCommand("FUNC:RUN 0");

                    await Task.Delay(100);
                    ////计算IL值
                    //if (!int.TryParse(textBox_22_4.Text, out int textBoxValue))
                    //{
                    //    MessageBox.Show("请输入有效的数字");
                    //    return;
                    //}
                    //int IL = textBoxValue * 10 + 1;
                    //if (IL > 4096)
                    //{
                    //    IL = 4096;
                    //}
                    //设置采样率
                    manager.SendCommand($"FUNC:SAMPle {_MaxCount78}");

                    await Task.Delay(100);
                    //设置功能长度
                    manager.SendCommand($"FUNC:LENgth {_MaxCount78}");

                    await Task.Delay(100);
                    //设置参数为电流
                    manager.SendCommand("FUNC:PARAmeter CURR");

                    await Task.Delay(100);
                    manager.SendCommand("FUNC:STRIng LIST");

                    //计算所有区间的a和b值
                    List<(double a, double b)> coefficients = CalculateCoefficients(real_data, set_dac_data);

                    List<double> ldy = new List<double>();//设计的电流值
                    List<double> ldx = new List<double>();//实际转换后的电流值
                    List<int> myf = new List<int>();//记录电流0为正1为负
                    StringBuilder CURRhex = new StringBuilder();//电流值十六进制
                    StringBuilder TIMEhex = new StringBuilder();//时间差十六进制

                    for (int j = 0; j < _MaxCount78; j++)
                    {
                        double WC = _magnetizationWave[j];
                        //Debug.WriteLine("函数my值：" + WC);
                        if (WC < 0) { myf.Add(1); }
                        else { myf.Add(0); }
                        ldy.Add(Math.Abs(WC));
                    }
                    foreach (double my in ldy)
                    {
                        //根据my值找到对应的区间和系数
                        var (a, b) = FindCoefficientsForValue(my, real_data, coefficients);

                        //计算mx值
                        double mx = (my - b) / a;
                        ldx.Add(mx);
                        //Debug.WriteLine("计算结果mx值：" + mx);
                        //转换为hex数据
                        int intValue = (int)(mx * 65535 / (60 * 1.02));
                        string hexValue = intValue.ToString("X4"); // 4位16进制
                        //Debug.WriteLine("* 65535/(60 * 1.02)转16进制结果：" + hexValue);
                        CURRhex.Append(hexValue);
                    }
                    foreach (int my in myf)
                    {
                        string hexValue = TimeToHex(_magnetizationMeasurementPoint, my);
                        TIMEhex.Append(hexValue);
                    }


                    //分段发送CURRhex
                    bool allSegmentsSentSuccessfully = await SendHexSegments(manager, CURRhex);

                    if (!allSegmentsSentSuccessfully)
                    {
                        manager.SendCommand("FUNC:END");
                        button_22_1.Enabled = true;
                        Debug.WriteLine("电流值发送多次未成功，跳出");
                        return;
                    }

                    //发送成功，继续发送TIMEhex
                    manager.SendCommand("FUNC:PARAmeter TIME");
                    await Task.Delay(100);


                    //分段发送TIMEhex
                    bool allTimeSegmentsSentSuccessfully = await SendHexSegments(manager, TIMEhex);

                    if (allTimeSegmentsSentSuccessfully)
                    {
                        manager.SendCommand("FUNC:END");
                        await Task.Delay(100);
                        manager.SendCommand("OUTPUT ON");
                    }
                    else
                    {
                        manager.SendCommand("FUNC:END");
                        button_22_1.Enabled = true;
                        Debug.WriteLine("时间值发送多次未成功，跳出");
                        return;
                    }
                    await Task.Delay(100);
                    //设置电源为输出状态，输出上述设定
                    manager.SendCommand("OUTP 1");

                    await Task.Delay(100);
                }
            }

            // 重置状态
            _magnetizationCurrentIndex = 0;
            _magnetizationXData.Clear();
            _magnetizationYData.Clear();
            // 清空图表
            formsPlot3.Plot.Clear();
            // 设置定时器间隔
            int intervalMs = (int)(_magnetizationMeasurementPoint * 1000);
            _magnetizationTimer.Interval = Math.Max(1, intervalMs); // 确保至少1ms
            // 启动定时器
            _magnetizationTimer.Start();
        }
        private void MagnetizationTimer_Tick(object sender, EventArgs e)
        {
            UpdateMagnetizationPlot();
        }

        private void UpdateMagnetizationPlot()
        {
            if (_magnetizationCurrentIndex < _magnetizationWave.Length)
            {
                // 添加新数据点
                _magnetizationXData.Add((double)_magnetizationCurrentIndex / _measurementPoint);
                _magnetizationYData.Add(_magnetizationWave[_magnetizationCurrentIndex]);

                // 更新图表
                formsPlot3.Plot.Clear();
                var scatter = formsPlot3.Plot.Add.Scatter(_magnetizationXData.ToArray(), _magnetizationYData.ToArray());

                formsPlot3.Plot.Axes.AutoScale();
                formsPlot3.Refresh();
                _magnetizationCurrentIndex++;
            }
            else
            {
                // 数据已全部绘制完成，停止定时器
                _magnetizationTimer.Stop();
                AutoTriggerStopButton();
            }
        }
        #endregion

        // 自动触发停止按钮
        private void AutoTriggerStopButton()
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(AutoTriggerStopButton));
                return;
            }

            // 模拟点击停止按钮
            button_22_2_Click(null, EventArgs.Empty);
        }


        private void button_22_2_Click(object sender, EventArgs e)
        {
            _demagnetizationTimer?.Stop();
            _magnetizationTimer?.Stop();
            button_able(false);

            for (int i = 0; i < _PowerTcpManagers78.Length; i++)
            {
                var manager = _PowerTcpManagers78[i];
                if (manager == null) continue;

                if (manager.IsConnected)
                {
                    manager.SendCommand("OUTP 0\n");
                }

            }
        }
        // 计算所有区间的a和b值
        private List<(double a, double b)> CalculateCoefficients(float[] real_data, float[] set_dac_data)
        {
            List<(double a, double b)> coefficients = new List<(double, double)>();

            for (int i = 0; i < real_data.Length - 1; i++)
            {
                double x1 = set_dac_data[i];
                double x2 = set_dac_data[i + 1];
                double y1 = real_data[i];
                double y2 = real_data[i + 1];

                double a = (y2 - y1) / (x2 - x1);
                double b = y1 - a * x1;


                Debug.WriteLine("第" + i + "段a值：" + a + "  第" + i + "段b值：" + b);
                coefficients.Add((a, b));
            }

            // 最后一个区间使用上一行的值
            if (coefficients.Count > 0)
            {
                coefficients.Add(coefficients[coefficients.Count - 1]);
            }

            return coefficients;
        }

        // 根据值找到对应的系数
        private (double a, double b) FindCoefficientsForValue(double value, float[] real_data, List<(double a, double b)> coefficients)
        {
            if (value < real_data[0])
            {
                return (1, 0); // 默认值，或者您可能需要特殊处理
            }

            for (int i = 0; i < real_data.Length - 1; i++)
            {
                if (value >= real_data[i] && value < real_data[i + 1])
                {
                    return coefficients[i];
                }
            }

            // 大于最后一个值
            return coefficients[coefficients.Count - 1];
        }

        public static string TimeToHex(double a, int c)
        {
            // 计算表达式 A*1000/0.005
            double result = a / 0.00005;

            // 将结果转换为整数（截断小数部分）
            long integerResult = (long)result;

            // 转换为16进制字符串（大写）
            string hexResult = integerResult.ToString("X");

            // 确保至少7位，不足则前补0
            hexResult = hexResult.PadLeft(7, '0');

            // 根据C的值添加前缀
            if (c == 0)
            {
                hexResult = "4" + hexResult;
            }
            else if (c == 1)
            {
                hexResult = "5" + hexResult;
            }

            return hexResult;
        }

        private async Task<bool> SendHexSegments(PowerTcpClientManager manager, StringBuilder hexData)
        {
            const int segmentSize = 128; // 每段100字节
            const int maxRetries = 3;    // 最大重试次数
            const int timeoutMs = 2000;  // 2秒超时

            // 将hexData按128字节分段
            List<string> segments = new List<string>();
            for (int i = 0; i < hexData.Length; i += segmentSize)
            {
                int length = Math.Min(segmentSize, hexData.Length - i);
                segments.Add(hexData.ToString(i, length));
            }

            // 发送每个分段
            for (int segmentIndex = 0; segmentIndex < segments.Count; segmentIndex++)
            {
                string currentSegment = segments[segmentIndex];
                int expectedBytes = currentSegment.Length / 2; // 16进制字符串的字节数
                int retryCount = 0;
                bool segmentSuccess = false;

                while (!segmentSuccess)
                {
                    // 1. 首先尝试发送FUNCLISTNEXT
                    if (retryCount == 0)
                    {
                        //string nextCmd = "FUNCLISTNEXT " + currentSegment;
                        string nextCmd = "46554E434C4953544E455854" + currentSegment + "0D0A";
                        var (success, _) = await SendCommandWithValidation(manager, nextCmd, expectedBytes, timeoutMs);

                        if (success)
                        {
                            segmentSuccess = true;
                            continue; // 成功则处理下一个分段
                        }
                    }

                    // 2. 发送失败后进入FUNCLISTBACK重试循环
                    retryCount++;
                    if (retryCount > maxRetries)
                    {
                        return false; // 超过最大重试次数
                    }

                    //string backCmd = "FUNCLISTBACK " + currentSegment;
                    string backCmd = "46554E434C4953544241434B" + currentSegment + "0D0A";
                    var (backSuccess, responseValid) = await SendCommandWithValidation(manager, backCmd, expectedBytes, timeoutMs);

                    if (responseValid)
                    {
                        // BACK成功且返回值正确，可以重新尝试NEXT
                        retryCount = 0; // 重置计数器
                    }
                    // 如果BACK失败则继续循环
                }
            }

            return true; // 所有分段发送成功
        }
        private async Task<(bool success, bool responseValid)> SendCommandWithValidation(PowerTcpClientManager manager, string command, int expectedLength, int timeoutMs)
        {
            try
            {
                var responseTask = manager.SendCommandAndReceiveResponseAsyncForHex(command);

                // 等待响应或超时
                if (await Task.WhenAny(responseTask, Task.Delay(timeoutMs)) == responseTask)
                {
                    string response = responseTask.Result;
                    bool isValid = response == expectedLength.ToString();
                    return (true, isValid); // 收到响应，返回验证结果
                }

                return (false, false); // 超时
            }
            catch
            {
                return (false, false); // 发生异常
            }
        }

        private void button_able(bool bl)
        {
            radioButton_22_1.Enabled = !bl;
            radioButton_22_2.Enabled = !bl;
            radioButton_22_3.Enabled = !bl;
            radioButton_22_4.Enabled = !bl;
            textBox_22_1.Enabled = !bl;
            textBox_22_2.Enabled = !bl;
            textBox_22_3.Enabled = !bl;
            textBox_22_4.Enabled = !bl;
            comboBox_22_1.Enabled = !bl;
            button_22_1.Enabled = !bl;
            button_22_2.Enabled = bl;
            tabControl23.Enabled = !bl;
        }

        #endregion

        #region 选项2探头磁场值显示

        private System.Threading.Timer _refreshTimer;
        private ProbeInfo _currentProbe;


        private void StartProbeMonitoring()
        {
            try
            {
                // 1. 初始化ComboBox
                InitializeComboBox_23_1();

                // 2. 启动定时刷新
                if (_refreshTimer == null)
                {
                    _refreshTimer = new System.Threading.Timer(
                        callback: RefreshProbeValues,
                        state: null,
                        dueTime: 0,
                        period: 1000);
                }

                //MessageBox.Show("探头监控已启动！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                //MessageBox.Show($"启动失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void StopProbeMonitoring()
        {
            try
            {
                // 1. 停止定时器
                if (_refreshTimer != null)
                {
                    _refreshTimer.Dispose();
                    _refreshTimer = null;
                }

                // 2. 清空ComboBox
                comboBox_23_1.Items.Clear();
                comboBox_23_1.Text = string.Empty;

                // 3. 清空文本框
                ClearTextBoxes();

                // 4. 重置当前探头
                _currentProbe = null;

                // 5. 移除事件绑定（避免重复绑定）
                comboBox_23_1.SelectedIndexChanged -= comboBox_23_1_SelectedIndexChanged;

                //MessageBox.Show("探头监控已停止！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                //MessageBox.Show($"停止失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void InitializeComboBox_23_1()
        {
            comboBox_23_1.Items.Clear();
            comboBox_23_1.SelectedIndexChanged -= comboBox_23_1_SelectedIndexChanged;

            // 检查是否有数据
            if (_probeList == null || _probeList.Count == 0)
            {
                return;
            }

            // 添加所有探头名称
            foreach (var probe in _probeList)
            {
                comboBox_23_1.Items.Add(probe.ProbeName);
            }

            //// 如果有探头，默认选择第一个
            //if (comboBox_23_1.Items.Count > 0)
            //{
            //    comboBox_23_1.SelectedIndex = 0;
            //}

            // 绑定选择改变事件
            comboBox_23_1.SelectedIndexChanged += comboBox_23_1_SelectedIndexChanged;

        }
        private void comboBox_23_1_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateSelectedProbeInfo();
        }

        private void RefreshProbeValues(object state)
        {
            // 需要在UI线程上更新控件
            if (comboBox_23_1.InvokeRequired)
            {
                comboBox_23_1.Invoke(new Action(() => RefreshProbeValues(state)));
                return;
            }

            // 如果有选中的探头，更新显示
            if (_currentProbe != null)
            {
                UpdateTextBoxValues();
            }
        }

        private void UpdateSelectedProbeInfo()
        {
            if (comboBox_23_1.SelectedItem == null)
            {
                _currentProbe = null;
                ClearTextBoxes();
                return;
            }

            string selectedProbeName = comboBox_23_1.SelectedItem.ToString();

            // 从列表中找到对应的探头
            _currentProbe = _probeList.FirstOrDefault(p => p.ProbeName == selectedProbeName);

            if (_currentProbe != null)
            {
                UpdateTextBoxValues();
            }
            else
            {
                ClearTextBoxes();
            }
        }

        private void UpdateTextBoxValues()
        {
            if (_currentProbe == null) return;

            // 安全更新文本框
            if (textBox_23_1.InvokeRequired)
            {
                textBox_23_1.Invoke(new MethodInvoker(UpdateTextBoxValues));
                return;
            }

            // 更新文本框显示
            textBox_23_1.Text = _currentProbe.X.ToString("F3");
            textBox_23_2.Text = _currentProbe.Y.ToString("F3");
            textBox_23_3.Text = _currentProbe.Z.ToString("F3");
        }

        private void ClearTextBoxes()
        {
            if (textBox_23_1.InvokeRequired)
            {
                textBox_23_1.Invoke(new MethodInvoker(ClearTextBoxes));
                return;
            }

            textBox_23_1.Text = string.Empty;
            textBox_23_2.Text = string.Empty;
            textBox_23_3.Text = string.Empty;
        }

        #endregion


        #region 选项3数据计算-1零磁场磁矩

        private string CalculateDataPath = null;

        private ExcelProbeHelper excelProbeHelper;
        private string excelFilePath = Path.Combine(System.Windows.Forms.Application.StartupPath, "数据计算参数.xlsx");

        private string conStr = "";
        private MySqlDataAdapter dataAdapter;
        private DataTable exceldataTable = new DataTable();
        private Dictionary<string, string> excelData = new Dictionary<string, string>();
        double[,] magnetometer;
        double[,] mag_moment = new double[9, 4];
        int moveline = 1;
        int jiance = 0;

        private void start3()
        {
            //_excelHelper = new ExcelHelper(excelFilePath);
        }

        private void radioButton_30_1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton_30_1.Checked)
            {
                button_31_1.Enabled = false;
                comboBox_30_1.Enabled = true;
                InitializeComboBox_30_1();
            }
            else
            {
                button_31_1.Enabled = true;
                comboBox_30_1.Enabled = false;
            }
        }

        //初始化comboBox_30_1
        private void InitializeComboBox_30_1()
        {
            try
            {
                comboBox_30_1.Items.Clear();

                // 检查CalculateDataPath路径是否存在
                if (!Directory.Exists(CalculateDataPath))
                {
                    comboBox_30_1.SelectedText = "";
                    return;
                }

                // 搜索所有Excel文件
                string[] excelFiles = Directory.GetFiles(CalculateDataPath, "*.xlsx", SearchOption.TopDirectoryOnly)
                    .Concat(Directory.GetFiles(CalculateDataPath, "*.xls", SearchOption.TopDirectoryOnly))
                    .ToArray();

                if (excelFiles.Length == 0)
                {
                    comboBox_30_1.SelectedText = "";
                    return;
                }

                // 添加文件到ComboBox
                foreach (string filePath in excelFiles)
                {
                    string fileName = Path.GetFileName(filePath);
                    comboBox_30_1.Items.Add(fileName);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"获取Excel文件列表失败：{ex.Message}", "错误",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void comboBox_30_1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                // 检查是否有选中项
                if (comboBox_30_1.SelectedItem == null)
                {
                    excelFilePath = string.Empty;
                    return;
                }

                string selectedFileName = comboBox_30_1.SelectedItem.ToString();
                excelFilePath = Path.Combine(CalculateDataPath, selectedFileName);

                if (!File.Exists(excelFilePath))
                {
                    MessageBox.Show($"文件不存在：{excelFilePath}", "警告",
                                  MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    excelFilePath = string.Empty;
                    return;
                }

                excelProbeHelper = new ExcelProbeHelper(excelFilePath);
                // 读取Excel数据到DataTable，并获取channel数量
                int channelCount;
                exceldataTable.Clear();
                exceldataTable = excelProbeHelper.ReadSheet1Data(out channelCount);
                ComboBox[] comboBoxes = { comboBox_31_1, comboBox_31_2, comboBox_31_3, comboBox_31_4 };
                // 为每个comboBox添加选项
                foreach (var comboBox in comboBoxes)
                {
                    if (comboBox != null)
                    {
                        comboBox.Items.Clear();
                        // 添加channel选项（1, 2, 3, ...）
                        for (int i = 1; i <= channelCount; i++)
                        {
                            comboBox.Items.Add(i.ToString());
                        }
                        // 默认选择第一个（如果之前有选择则保持原选择）
                        if (comboBox.Items.Count > 0 && comboBox.SelectedIndex == -1)
                        {
                            comboBox.SelectedIndex = 0;
                        }
                    }
                }

                // 读取Sheet2参数
                excelData = excelProbeHelper.ReadSheet2Data();


            }
            catch (Exception ex)
            {
                MessageBox.Show($"切换文件失败：{ex.Message}", "错误",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        private void SetConStr()
        {
            string ip = "127.0.0.1";
            string port = "3306";
            string user = "root";
            string password = "123456";
            string database = "datademo1";
            conStr = "server=" + ip + ";port=" + port + ";user=" + user + ";password=" + password + ";database=" + database + ";";
        }


        //读取数据库数据并展示
        private void button_31_1_Click(object sender, EventArgs e)
        {
            try
            {

                ReadExcelDataToDataTable();


                //var excelData = _excelHelper.ReadExcelData();

                if (excelData.ContainsKey("监测干扰"))
                {
                    checkBox_31_1.Checked = excelData["监测干扰"] == "是";
                }

                if (excelData.ContainsKey("旋转角度"))
                {
                    radioButton_31_1.Checked = excelData["旋转角度"] == "10";
                    radioButton_31_2.Checked = !radioButton_31_1.Checked;
                }

                if (excelData.ContainsKey("干扰探头"))
                {
                    comboBox_31_1.SelectedItem = excelData["干扰探头"];
                }

                if (excelData.ContainsKey("探头1"))
                {
                    comboBox_31_2.SelectedItem = excelData["探头1"];
                }

                if (excelData.ContainsKey("探头2"))
                {
                    comboBox_31_3.SelectedItem = excelData["探头2"];
                }

                if (excelData.ContainsKey("探头3"))
                {
                    comboBox_31_4.SelectedItem = excelData["探头3"];
                }

                if (excelData.ContainsKey("距离1"))
                {
                    textBox_31_r1.Text = excelData["距离1"];
                }

                if (excelData.ContainsKey("距离2"))
                {
                    textBox_31_r2.Text = excelData["距离2"];
                }

                if (excelData.ContainsKey("距离3"))
                {
                    textBox_31_r3.Text = excelData["距离3"];
                }

                if (excelData.ContainsKey("坐标关系X"))
                {
                    comboBox_31_5.SelectedIndex = ExcelHelper.GetCoordinateIndex(excelData["坐标关系X"]);
                }

                if (excelData.ContainsKey("坐标关系Y"))
                {
                    comboBox_31_6.SelectedIndex = ExcelHelper.GetCoordinateIndex(excelData["坐标关系Y"]);
                }

                if (excelData.ContainsKey("坐标关系Z"))
                {
                    comboBox_31_7.SelectedIndex = ExcelHelper.GetCoordinateIndex(excelData["坐标关系Z"]);
                }

                //MessageBox.Show("数据导入完成！");
                return;
                //以下为数据库读取方式

                SetConStr();
                // 创建连接对象
                using (MySqlConnection connection = new MySqlConnection(conStr))
                {
                    // 打开连接
                    connection.Open();

                    // 读取“监测干扰”的参数值
                    string query1 = "SELECT 参数值 FROM 零磁场磁矩参数 WHERE 参数名称 = '监测干扰'";
                    using (MySqlCommand command1 = new MySqlCommand(query1, connection))
                    {
                        object result1 = command1.ExecuteScalar();
                        if (result1 != null && result1.ToString() == "是")
                        {
                            checkBox_31_1.Checked = true;
                        }
                        else
                        {
                            checkBox_31_1.Checked = false;
                        }
                    }

                    // 读取“旋转角度”的参数值
                    string query2 = "SELECT 参数值 FROM 零磁场磁矩参数 WHERE 参数名称 = '旋转角度'";
                    using (MySqlCommand command2 = new MySqlCommand(query2, connection))
                    {
                        object result2 = command2.ExecuteScalar();
                        if (result2 != null)
                        {
                            string value = result2.ToString();
                            if (value == "10")
                            {
                                radioButton_31_1.Checked = true;
                            }
                            else if (value == "20")
                            {
                                radioButton_31_2.Checked = true;
                            }
                        }
                    }

                    // 处理探头相关的下拉框
                    string[] probeParameterNames = { "干扰探头", "探头1", "探头2", "探头3" };
                    System.Windows.Forms.ComboBox[] comboBoxes = { comboBox_31_1, comboBox_31_2, comboBox_31_3, comboBox_31_4 };

                    for (int i = 0; i < probeParameterNames.Length; i++)
                    {
                        string query = "SELECT 参数值 FROM 零磁场磁矩参数 WHERE 参数名称 = @ParameterName";
                        using (MySqlCommand command = new MySqlCommand(query, connection))
                        {
                            command.Parameters.AddWithValue("@ParameterName", probeParameterNames[i]);

                            using (MySqlDataReader reader = command.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    string parameterValue = reader["参数值"].ToString();
                                    // 找到对应的值并选中
                                    for (int j = 0; j < comboBoxes[i].Items.Count; j++)
                                    {
                                        if (comboBoxes[i].Items[j].ToString() == parameterValue)
                                        {
                                            comboBoxes[i].SelectedIndex = j;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }


                    // 处理距离相关的文本框
                    string[] parameterNames = { "距离1", "距离2", "距离3" };
                    System.Windows.Forms.TextBox[] textBoxes = { textBox_31_r1, textBox_31_r2, textBox_31_r3 };

                    for (int i = 0; i < parameterNames.Length; i++)
                    {
                        string query = "SELECT 参数值 FROM 零磁场磁矩参数 WHERE 参数名称 = @ParameterName";
                        using (MySqlCommand command = new MySqlCommand(query, connection))
                        {
                            command.Parameters.AddWithValue("@ParameterName", parameterNames[i]);

                            using (MySqlDataReader reader = command.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    textBoxes[i].Text = reader["参数值"].ToString();
                                }
                            }
                        }
                    }

                    // 处理坐标关系相关的下拉框
                    string[] coordinateParameterNames = { "坐标关系X", "坐标关系Y", "坐标关系Z" };
                    System.Windows.Forms.ComboBox[] coordinateComboBoxes = { comboBox_31_5, comboBox_31_6, comboBox_31_7 };
                    //string[] coordinateValues = { "X", "Y", "Z", "-X", "-Y", "-Z" };

                    for (int i = 0; i < coordinateParameterNames.Length; i++)
                    {
                        string query = "SELECT 参数值 FROM 零磁场磁矩参数 WHERE 参数名称 = @ParameterName";
                        using (MySqlCommand command = new MySqlCommand(query, connection))
                        {
                            command.Parameters.AddWithValue("@ParameterName", coordinateParameterNames[i]);

                            using (MySqlDataReader reader = command.ExecuteReader())
                            {
                                if (reader.Read())
                                {
                                    string parameterValue = reader["参数值"].ToString();
                                    int index;
                                    if (int.TryParse(parameterValue, out index) && index >= 1 && index <= 6)
                                    {
                                        coordinateComboBoxes[i].SelectedIndex = index - 1;
                                    }
                                    else
                                    {
                                        MessageBox.Show($"获取的 {coordinateParameterNames[i]} 参数值无效！");
                                    }
                                }
                            }
                        }
                    }


                    GetTestDate();//读取测试数据

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }


        //保存数据至数据库
        private void button_31_2_Click(object sender, EventArgs e)
        {
            try
            {
                var dataToWrite = new Dictionary<string, string>
                {
                    ["监测干扰"] = checkBox_31_1.Checked ? "是" : "否",
                    ["旋转角度"] = radioButton_31_1.Checked ? "10" : "20",
                    ["干扰探头"] = comboBox_31_1.SelectedItem?.ToString(),
                    ["探头1"] = comboBox_31_2.SelectedItem?.ToString(),
                    ["探头2"] = comboBox_31_3.SelectedItem?.ToString(),
                    ["探头3"] = comboBox_31_4.SelectedItem?.ToString(),
                    ["距离1"] = textBox_31_r1.Text,
                    ["距离2"] = textBox_31_r2.Text,
                    ["距离3"] = textBox_31_r3.Text,
                    ["坐标关系X"] = ExcelHelper.GetCoordinateValue(comboBox_31_5.SelectedItem?.ToString()),
                    ["坐标关系Y"] = ExcelHelper.GetCoordinateValue(comboBox_31_6.SelectedItem?.ToString()),
                    ["坐标关系Z"] = ExcelHelper.GetCoordinateValue(comboBox_31_7.SelectedItem?.ToString())
                };

                if (excelProbeHelper.WriteExcelData(dataToWrite))
                {
                    MessageBox.Show("数据导出成功！");
                }
                return;
                //以下为数据库保存方式

                SetConStr();

                using (MySqlConnection connection = new MySqlConnection(conStr))
                {
                    connection.Open();

                    // 保存距离相关的文本框数据
                    string[] distanceParameterNames = { "距离1", "距离2", "距离3" };
                    System.Windows.Forms.TextBox[] textBoxes = { textBox_31_r1, textBox_31_r2, textBox_31_r3 };

                    for (int i = 0; i < distanceParameterNames.Length; i++)
                    {
                        string updateQuery = "UPDATE 零磁场磁矩参数 SET 参数值 = @Value WHERE 参数名称 = @ParameterName";
                        using (MySqlCommand command = new MySqlCommand(updateQuery, connection))
                        {
                            command.Parameters.AddWithValue("@Value", textBoxes[i].Text);
                            command.Parameters.AddWithValue("@ParameterName", distanceParameterNames[i]);
                            command.ExecuteNonQuery();
                        }
                    }

                    // 保存探头相关的下拉框数据
                    string[] probeParameterNames = { "干扰探头", "探头1", "探头2", "探头3" };
                    System.Windows.Forms.ComboBox[] comboBoxes = { comboBox_31_1, comboBox_31_2, comboBox_31_3, comboBox_31_4 };

                    for (int i = 0; i < probeParameterNames.Length; i++)
                    {
                        string updateQuery = "UPDATE 零磁场磁矩参数 SET 参数值 = @Value WHERE 参数名称 = @ParameterName";
                        using (MySqlCommand command = new MySqlCommand(updateQuery, connection))
                        {
                            command.Parameters.AddWithValue("@Value", comboBoxes[i].SelectedItem.ToString());
                            command.Parameters.AddWithValue("@ParameterName", probeParameterNames[i]);
                            command.ExecuteNonQuery();
                        }
                    }

                    // 保存监测干扰复选框数据
                    string monitorInterferenceUpdateQuery = "UPDATE 零磁场磁矩参数 SET 参数值 = @Value WHERE 参数名称 = '监测干扰'";
                    using (MySqlCommand monitorInterferenceCommand = new MySqlCommand(monitorInterferenceUpdateQuery, connection))
                    {
                        monitorInterferenceCommand.Parameters.AddWithValue("@Value", checkBox_31_1.Checked ? "是" : "否");
                        monitorInterferenceCommand.ExecuteNonQuery();
                    }

                    // 保存旋转角度单选框数据
                    string rotationAngleUpdateQuery = "UPDATE 零磁场磁矩参数 SET 参数值 = @Value WHERE 参数名称 = '旋转角度'";
                    using (MySqlCommand rotationAngleCommand = new MySqlCommand(rotationAngleUpdateQuery, connection))
                    {
                        string rotationAngleValue = radioButton_31_1.Checked ? "10" : "20";
                        rotationAngleCommand.Parameters.AddWithValue("@Value", rotationAngleValue);
                        rotationAngleCommand.ExecuteNonQuery();
                    }



                    // 保存坐标关系相关的下拉框数据
                    string[] coordinateParameterNames = { "坐标关系X", "坐标关系Y", "坐标关系Z" };
                    System.Windows.Forms.ComboBox[] coordinateComboBoxes = { comboBox_31_5, comboBox_31_6, comboBox_31_7 };
                    string[] coordinateValues = { "X", "Y", "Z", "-X", "-Y", "-Z" };

                    for (int i = 0; i < coordinateParameterNames.Length; i++)
                    {
                        string selectedValue = coordinateComboBoxes[i].SelectedItem.ToString();
                        int index = Array.IndexOf(coordinateValues, selectedValue) + 1;

                        string updateQuery = "UPDATE 零磁场磁矩参数 SET 参数值 = @Value WHERE 参数名称 = @ParameterName";
                        using (MySqlCommand command = new MySqlCommand(updateQuery, connection))
                        {
                            command.Parameters.AddWithValue("@Value", index);
                            command.Parameters.AddWithValue("@ParameterName", coordinateParameterNames[i]);
                            command.ExecuteNonQuery();
                        }
                    }

                    MessageBox.Show("数据已成功保存到数据库！");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }


        //读取测试数据
        private void GetTestDate()
        {
            exceldataTable.Clear();
            try
            {
                using (MySqlConnection connection = new MySqlConnection(conStr))
                {
                    connection.Open();
                    string query = "SELECT * FROM 测试数据";

                    // 创建 MySQL 命令对象
                    using (MySqlCommand command = new MySqlCommand(query, connection))
                    {
                        // 创建 MySQL 数据适配器
                        using (MySqlDataAdapter adapter = new MySqlDataAdapter(command))
                        {
                            // 使用数据适配器填充 DataTable
                            adapter.Fill(exceldataTable);
                            //MessageBox.Show(TestdataTable.Rows.Count.ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }

        }

        //读取excel文件数据放入TestdataTable中
        private void ReadExcelDataToDataTable()
        {
            exceldataTable.Clear();
            try
            {
                // 选择Excel文件
                OpenFileDialog openFileDialog = new OpenFileDialog();
                //openFileDialog.InitialDirectory = Path.Combine(System.Windows.Forms.Application.StartupPath, "ExcelData");
                if (string.IsNullOrEmpty(CalculateDataPath) || !Directory.Exists(CalculateDataPath))
                {
                    openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                }
                else
                {
                    openFileDialog.InitialDirectory = CalculateDataPath;
                }


                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog.Title = "选择Excel数据文件";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    CalculateDataPath = Path.GetDirectoryName(openFileDialog.FileName);

                    string excelFilePath = openFileDialog.FileName;

                    excelProbeHelper = new ExcelProbeHelper(excelFilePath);
                    // 读取Excel数据到DataTable，并获取channel数量
                    int channelCount;
                    exceldataTable = excelProbeHelper.ReadSheet1Data(out channelCount);
                    ComboBox[] comboBoxes = { comboBox_31_1, comboBox_31_2, comboBox_31_3, comboBox_31_4 };
                    // 为每个comboBox添加选项
                    foreach (var comboBox in comboBoxes)
                    {
                        if (comboBox != null)
                        {
                            comboBox.Items.Clear();
                            // 添加channel选项（1, 2, 3, ...）
                            for (int i = 1; i <= channelCount; i++)
                            {
                                comboBox.Items.Add(i.ToString());
                            }
                            // 默认选择第一个（如果之前有选择则保持原选择）
                            if (comboBox.Items.Count > 0 && comboBox.SelectedIndex == -1)
                            {
                                comboBox.SelectedIndex = 0;
                            }
                        }
                    }


                    // 读取Sheet2参数
                    excelData = excelProbeHelper.ReadSheet2Data();

                    textBox_30_1.Text = openFileDialog.FileName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Excel读取错误: " + ex.Message);
            }
        }






        //计算总磁矩
        private void button_31_3_Click(object sender, EventArgs e)
        {
            int[] randomNumbers = new int[4];
            randomNumbers[0] = int.Parse(comboBox_31_2.Text);
            randomNumbers[1] = int.Parse(comboBox_31_3.Text);
            randomNumbers[2] = int.Parse(comboBox_31_4.Text);
            randomNumbers[3] = int.Parse(comboBox_31_1.Text);

            //重新按照探头顺序排序
            List<string> columnNames = new List<string>();
            foreach (int num in randomNumbers)
            {
                columnNames.AddRange(new string[] { $"{num}01", $"{num}02", $"{num}03" });
            }

            int rowCount = exceldataTable.Rows.Count;
            int colCount = columnNames.Count;
            double[,] resultArray = new double[rowCount, colCount];

            for (int colIndex = 0; colIndex < colCount; colIndex++)
            {
                string columnName = columnNames[colIndex];
                DataColumn column = exceldataTable.Columns[columnName];
                for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
                {
                    try
                    {
                        resultArray[rowIndex, colIndex] = Convert.ToDouble(exceldataTable.Rows[rowIndex][column]);
                    }
                    catch (FormatException)
                    {
                        MessageBox.Show($"列 {columnName} 中存在无法转换为 double 的值");
                        return;
                    }
                }
            }


            // 创建新的二维数组，存储第 2 行到第 37 行和最后一行的数据
            double[,] finalArray = new double[37, colCount];


            int newRowIndex = 0;
            // 复制第 1 行到第 36 行的数据
            for (int i = 0; i < 36; i++)
            {
                for (int j = 0; j < colCount; j++)
                {
                    finalArray[newRowIndex, j] = resultArray[i, j];
                }
                newRowIndex++;
            }
            // 复制最后一行的数据
            for (int j = 0; j < colCount; j++)
            {
                finalArray[newRowIndex, j] = resultArray[rowCount - 1, j];
            }


            int angle_step = 10;
            if (radioButton_31_1.Checked)
            {
                angle_step = 10;
            }
            else
            {
                angle_step = 20;
            }

            if (checkBox_31_1.Checked)
            {
                jiance = 1;
            }
            else
            {
                jiance = 0;
            }

            SetMagnetometer();

            Debug.WriteLine(finalArray[0, 0]);
            double[,] mag_moment = MomentCalculator.CalculateMoment(
                magnetometer,
                angle_step,
                finalArray,
                moveline,
                jiance);

            // 输出结果
            for (int i = 0; i < mag_moment.GetLength(0); i++)
            {
                Console.Write($"m{i + 1}: [");
                for (int j = 0; j < mag_moment.GetLength(1); j++)
                {
                    Console.Write($"{mag_moment[i, j]:F12}, ");
                }
                Console.WriteLine("]");
            }

            //进行XYZ转换
            string selectedX = comboBox_31_5.SelectedItem?.ToString();
            string selectedY = comboBox_31_6.SelectedItem?.ToString();
            string selectedZ = comboBox_31_7.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(selectedX) || string.IsNullOrEmpty(selectedY) || string.IsNullOrEmpty(selectedZ))
            {
                MessageBox.Show("请为所有坐标轴选择方向");
                return;
            }
            double[,] new_moment = MomentCalculator.TransformMagneticMoments(mag_moment, selectedX, selectedY, selectedZ);


            FillTextBoxes_31(new_moment);
        }

        private void SetMagnetometer()
        {
            double result1;
            double result2;
            double result3;

            if (double.TryParse(textBox_31_r1.Text, out double value1))
            {
                result1 = value1 / 100;
            }
            else
            {
                result1 = 0;
                MessageBox.Show("r1 输入不是有效的数字，请重新输入。");
            }

            if (double.TryParse(textBox_31_r2.Text, out double value2))
            {
                result2 = value2 / 100;
            }
            else
            {
                result2 = 0;
                MessageBox.Show("r2 输入不是有效的数字，请重新输入。");
            }

            if (double.TryParse(textBox_31_r3.Text, out double value3))
            {
                result3 = value3 / 100;
            }
            else
            {
                result3 = 0;
                MessageBox.Show("r3 输入不是有效的数字，请重新输入。");
            }
            magnetometer = new double[4, 3] { { 0, result1, 0 }, { 0, result2, 0 }, { 0, result3, 0 }, { 0, 0, 0 } };
        }

        private void FillTextBoxes_31(double[,] array)
        {
            int textBoxIndex = 1;
            for (int i = 0; i < array.GetLength(0); i++)
            {
                for (int j = 0; j < array.GetLength(1); j++)
                {
                    // 获取对应的 TextBox 控件
                    System.Windows.Forms.TextBox textBox = Controls.Find($"textbox_31_{textBoxIndex}", true)[0] as System.Windows.Forms.TextBox;
                    if (textBox != null)
                    {
                        // 将数组元素的值转换为字符串并赋值给 TextBox
                        textBox.Text = (array[i, j] * 1000).ToString("F1");
                    }
                    textBoxIndex++;
                }
            }
        }




        #endregion

        #region 选项3数据计算-2地磁场磁矩

        #endregion


        #region 选项6系统配置-1磁强计配置

        // 新增数据存储类和线程安全集合
        private double _currentCollectionInterval = 1.0; // 默认1秒

        private void start61()
        {
            dataGridView_61_1.RowHeadersVisible = false;
            dataGridView_61_1.Columns.Add("探头序号", "探头序号");
            dataGridView_61_1.Columns.Add("设备号", "设备号");
            dataGridView_61_1.Columns.Add("探头", "探头");
            dataGridView_61_1.Columns.Add("地址", "地址");
            //dataGridView_61_1.Columns[0].Width = 80;
            //dataGridView_61_1.Columns[1].Width = 80;
            //dataGridView_61_1.Columns[2].Width = 80;
            dataGridView_61_1.Columns[3].Width = 160;
            dataGridView_61_2_Load();

            comboBox_61_1.SelectedIndexChanged += ComboBox_61_1_SelectedIndexChanged;
        }


        private void dataGridView_61_2_Load()
        {
            // 隐藏行标题列
            dataGridView_61_2.RowHeadersVisible = false;

            dataGridView_61_2.AllowUserToAddRows = false;
            dataGridView_61_2.AllowUserToDeleteRows = false;

            // 设置列
            dataGridView_61_2.Columns.Add("设备名", "设备名");
            dataGridView_61_2.Columns.Add("探头", "探头");

            // 添加复选框列
            DataGridViewCheckBoxColumn enableColumn = new DataGridViewCheckBoxColumn();
            enableColumn.Name = "启用";
            enableColumn.HeaderText = "启用";
            dataGridView_61_2.Columns.Add(enableColumn);

            //华舜默认探头
            for (int i = 1; i <= 4; i++)
            {
                for (int j = 1; j <= 8; j++)
                {
                    dataGridView_61_2.Rows.Add($"HS机箱{i}", $"探头{j}", false);
                }
            }

            //翠海默认探头
            //for (int i = 1; i <= 2; i++)
            //{
            //    for (int j = 1; j <= 8; j++)
            //    {
            //        dataGridView_61_2.Rows.Add($"CH机箱{i}", $"探头{j}", false);
            //    }
            //}
        }



        private void button_61_1_Click(object sender, EventArgs e)
        {
            // 确保"启用"列存在
            if (dataGridView_61_2.Columns.Contains("启用"))
            {
                int enableColumnIndex = dataGridView_61_2.Columns["启用"].Index;

                // 遍历所有行
                foreach (DataGridViewRow row in dataGridView_61_2.Rows)
                {
                    // 跳过新行（如果有）
                    if (!row.IsNewRow)
                    {
                        row.Cells[enableColumnIndex].Value = true;
                    }
                }
            }
        }

        private void button_61_2_Click(object sender, EventArgs e)
        {
            // 确保"启用"列存在
            if (dataGridView_61_2.Columns.Contains("启用"))
            {
                int enableColumnIndex = dataGridView_61_2.Columns["启用"].Index;

                // 遍历所有行
                foreach (DataGridViewRow row in dataGridView_61_2.Rows)
                {
                    // 跳过新行（如果有）
                    if (!row.IsNewRow)
                    {
                        row.Cells[enableColumnIndex].Value = false;
                    }
                }
            }
        }

        private void button_61_3_Click(object sender, EventArgs e)
        {
            dataGridView_61_1.Rows.Clear();

            var selectedRowsHS = new List<DataGridViewRow>();//华舜
            var selectedRowsCH = new List<DataGridViewRow>();//翠海
            int enableColumnIndex = dataGridView_61_2.Columns["启用"].Index;

            foreach (DataGridViewRow row in dataGridView_61_2.Rows)
            {
                if (!row.IsNewRow && row.Cells[enableColumnIndex].Value != null &&
                    (bool)row.Cells[enableColumnIndex].Value &&
                    row.Cells["设备名"].Value.ToString().Contains("HS"))
                {
                    selectedRowsHS.Add(row);
                }
            }
            foreach (DataGridViewRow row in dataGridView_61_2.Rows)
            {
                if (!row.IsNewRow && row.Cells[enableColumnIndex].Value != null &&
                    (bool)row.Cells[enableColumnIndex].Value &&
                    row.Cells["设备名"].Value.ToString().Contains("CH"))
                {
                    selectedRowsCH.Add(row);
                }
            }

            if (selectedRowsHS.Count == 0 && selectedRowsCH.Count == 0)
            {
                MessageBox.Show("请至少选择一项");
                return;
            }

            int HSandCH = 0;
            // 检查所有选中的HS设备名是否相同
            bool allSameDeviceHS = selectedRowsHS.All(row =>
                row.Cells["设备名"].Value.ToString() == selectedRowsHS[0].Cells["设备名"].Value.ToString());
            // HS按顺序添加选中项到右侧表格
            for (int i = 0; i < selectedRowsHS.Count; i++)
            {
                var row = selectedRowsHS[i];
                string fullDeviceName = row.Cells["设备名"].Value.ToString();
                string channel = row.Cells["探头"].Value.ToString().Replace("探头", "");

                // 处理设备号（去掉"机箱"二字）
                string deviceNumber = fullDeviceName.Replace("机箱", "");

                // 生成探头序号
                string probeNumber = $"探头{i + 1}";

                // 生成IP地址和端口
                string address;
                int channelNum = int.Parse(channel.Replace("探头", ""));

                if (allSameDeviceHS)
                {
                    // 情况1：所有设备名相同
                    address = $"192.168.1.{21 + (int)Math.Ceiling(channelNum / 2.0) - 1}:{5001 + (int)Math.Ceiling(channelNum / 2.0) - 1}";
                    //address = $"192.168.123.48:{5001 + (int)Math.Ceiling(channelNum / 2.0) - 1}";
                }
                else
                {
                    // 情况2：设备名不同

                    int ipBase = 21;
                    int portBase = 5001;

                    if (deviceNumber.Contains("1"))
                    {
                        ipBase = 21;
                        portBase = 5001;
                    }
                    else if (deviceNumber.Contains("2"))
                    {
                        ipBase = 25;
                        portBase = 5005;
                    }
                    else if (deviceNumber.Contains("3"))
                    {
                        ipBase = 29;
                        portBase = 5009;
                    }
                    else if (deviceNumber.Contains("4"))
                    {
                        ipBase = 33;
                        portBase = 5013;
                    }

                    address = $"192.168.1.{ipBase + (int)Math.Ceiling(channelNum / 2.0) - 1}:{portBase + (int)Math.Ceiling(channelNum / 2.0) - 1}";
                }

                // 添加到右侧表格
                dataGridView_61_1.Rows.Add(probeNumber, deviceNumber, channel, address);
                HSandCH++;
            }


            for (int i = 0; i < selectedRowsCH.Count; i++)
            {
                var row = selectedRowsCH[i];
                string fullDeviceName = row.Cells["设备名"].Value.ToString();
                string channel = row.Cells["探头"].Value.ToString().Replace("探头", "");

                // 处理设备号（去掉"机箱"二字）
                string deviceNumber = fullDeviceName.Replace("机箱", "");

                // 生成探头序号
                string probeNumber = $"探头{HSandCH + i + 1}";

                // 生成IP地址和端口
                string address;
                int channelNum = int.Parse(channel.Replace("探头", ""));

                if (deviceNumber.Contains("1") && channelNum < 5)
                {
                    address = $"192.168.0.168:1024";
                    //address = $"192.168.123.48:1024";
                }
                else if (channelNum >= 5 && deviceNumber.Contains("1"))
                {
                    address = $"192.168.0.169:1024";
                    //address = $"192.168.123.48:1025";
                }
                else if (deviceNumber.Contains("2") && channelNum < 5)
                {
                    address = $"192.168.0.170:1024";
                    //address = $"192.168.123.48:1026";
                }
                else
                {
                    address = $"192.168.0.171:1024";
                    //address = $"192.168.123.48:1027";
                }
                // 添加到右侧表格
                dataGridView_61_1.Rows.Add(probeNumber, deviceNumber, channel, address);
            }

        }


        private void ComboBox_61_1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (double.TryParse(comboBox_61_1.SelectedItem?.ToString(), out double interval))
            {
                _currentCollectionInterval = 1 / interval;
            }
        }

        private void checkBox_61_1_CheckedChanged(object sender, EventArgs e)
        {
            comboBox_61_3.Enabled = checkBox_61_1.Checked;
            comboBox_61_4.Enabled = checkBox_61_1.Checked;
            comboBox_61_5.Enabled = checkBox_61_1.Checked;
        }

        private void checkBox_61_2_CheckedChanged(object sender, EventArgs e)
        {
            comboBox_61_6.Enabled = checkBox_61_2.Checked;
            comboBox_61_7.Enabled = checkBox_61_2.Checked;
            comboBox_61_8.Enabled = checkBox_61_2.Checked;
        }



        #endregion

        #region 选项6系统配置-2电源配置

        void start62()
        {
            this.Load += Load62;
            this.FormClosing += Closing62;
        }



        // 自动保存所有TextBox
        private void AutoSaveAllTextBoxValues()
        {
            try
            {
                Properties.Settings.Default.comboBox_61_1_SelectedItem = comboBox_61_1.SelectedItem.ToString();
                Properties.Settings.Default.comboBox_61_2_SelectedItem = comboBox_61_2.SelectedItem.ToString();
                Properties.Settings.Default.comboBox_61_3_SelectedItem = comboBox_61_3.SelectedItem.ToString();
                Properties.Settings.Default.comboBox_61_4_SelectedItem = comboBox_61_4.SelectedItem.ToString();
                Properties.Settings.Default.comboBox_61_5_SelectedItem = comboBox_61_5.SelectedItem.ToString();
                Properties.Settings.Default.comboBox_61_6_SelectedItem = comboBox_61_6.SelectedItem.ToString();
                Properties.Settings.Default.comboBox_61_7_SelectedItem = comboBox_61_7.SelectedItem.ToString();
                Properties.Settings.Default.comboBox_61_8_SelectedItem = comboBox_61_8.SelectedItem.ToString();
                Properties.Settings.Default.textBox_62_1_1 = textBox_62_1_1.Text;
                Properties.Settings.Default.textBox_62_1_2 = textBox_62_1_2.Text;
                Properties.Settings.Default.textBox_62_1_3 = textBox_62_1_3.Text;
                Properties.Settings.Default.textBox_62_1_4 = textBox_62_1_4.Text;
                Properties.Settings.Default.textBox_62_2_1 = textBox_62_2_1.Text;
                Properties.Settings.Default.textBox_62_2_2 = textBox_62_2_2.Text;
                Properties.Settings.Default.textBox_62_2_3 = textBox_62_2_3.Text;
                Properties.Settings.Default.textBox_62_2_4 = textBox_62_2_4.Text;
                Properties.Settings.Default.textBox_62_3_1 = textBox_62_3_1.Text;
                Properties.Settings.Default.textBox_62_3_2 = textBox_62_3_2.Text;
                Properties.Settings.Default.textBox_62_3_3 = textBox_62_3_3.Text;
                Properties.Settings.Default.textBox_62_3_4 = textBox_62_3_4.Text;
                Properties.Settings.Default.textBox_62_4_1 = textBox_62_4_1.Text;
                Properties.Settings.Default.textBox_62_4_2 = textBox_62_4_2.Text;
                Properties.Settings.Default.textBox_62_4_3 = textBox_62_4_3.Text;
                Properties.Settings.Default.textBox_62_4_4 = textBox_62_4_4.Text;
                Properties.Settings.Default.textBox_62_5_1 = textBox_62_5_1.Text;
                Properties.Settings.Default.textBox_62_5_2 = textBox_62_5_2.Text;
                Properties.Settings.Default.textBox_62_5_3 = textBox_62_5_3.Text;
                Properties.Settings.Default.textBox_62_5_4 = textBox_62_5_4.Text;
                Properties.Settings.Default.textBox_62_6_1 = textBox_62_6_1.Text;
                Properties.Settings.Default.textBox_62_6_2 = textBox_62_6_2.Text;
                Properties.Settings.Default.textBox_62_6_3 = textBox_62_6_3.Text;
                Properties.Settings.Default.textBox_62_6_4 = textBox_62_6_4.Text;
                Properties.Settings.Default.textBox_62_7_1 = textBox_62_7_1.Text;
                Properties.Settings.Default.textBox_62_7_2 = textBox_62_7_2.Text;
                Properties.Settings.Default.textBox_62_8_1 = textBox_62_8_1.Text;
                Properties.Settings.Default.textBox_62_8_2 = textBox_62_8_2.Text;
                Properties.Settings.Default.textBox_62_9_1 = textBox_62_9_1.Text;
                Properties.Settings.Default.textBox_22_1 = textBox_22_1.Text;
                Properties.Settings.Default.textBox_22_2 = textBox_22_2.Text;
                Properties.Settings.Default.textBox_22_3 = textBox_22_3.Text;
                Properties.Settings.Default.textBox_22_4 = textBox_22_4.Text;
                Properties.Settings.Default.textBox_22_5 = textBox_22_5.Text;
                Properties.Settings.Default.textBox_22_6 = textBox_22_6.Text;
                Properties.Settings.Default.textBox_22_7 = textBox_22_7.Text;
                Properties.Settings.Default.radioButton_22_1 = radioButton_22_1.Checked;
                Properties.Settings.Default.radioButton_22_2 = radioButton_22_2.Checked;
                Properties.Settings.Default.radioButton_22_3 = radioButton_22_3.Checked;
                Properties.Settings.Default.radioButton_22_4 = radioButton_22_4.Checked;


            }
            catch (Exception ex)
            {
                MessageBox.Show($"自动保存时出错: {ex.Message}", "错误",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 自动加载所有TextBox
        private void AutoLoadAllTextBoxValues()
        {
            try
            {
                comboBox_61_1.SelectedItem = Properties.Settings.Default.comboBox_61_1_SelectedItem;
                comboBox_61_2.SelectedItem = Properties.Settings.Default.comboBox_61_2_SelectedItem;
                comboBox_61_3.SelectedItem = Properties.Settings.Default.comboBox_61_3_SelectedItem;
                comboBox_61_4.SelectedItem = Properties.Settings.Default.comboBox_61_4_SelectedItem;
                comboBox_61_5.SelectedItem = Properties.Settings.Default.comboBox_61_5_SelectedItem;
                comboBox_61_6.SelectedItem = Properties.Settings.Default.comboBox_61_6_SelectedItem;
                comboBox_61_7.SelectedItem = Properties.Settings.Default.comboBox_61_7_SelectedItem;
                comboBox_61_8.SelectedItem = Properties.Settings.Default.comboBox_61_8_SelectedItem;
                textBox_62_1_1.Text = Properties.Settings.Default.textBox_62_1_1;
                textBox_62_1_2.Text = Properties.Settings.Default.textBox_62_1_2;
                textBox_62_1_3.Text = Properties.Settings.Default.textBox_62_1_3;
                textBox_62_1_4.Text = Properties.Settings.Default.textBox_62_1_4;
                textBox_62_2_1.Text = Properties.Settings.Default.textBox_62_2_1;
                textBox_62_2_2.Text = Properties.Settings.Default.textBox_62_2_2;
                textBox_62_2_3.Text = Properties.Settings.Default.textBox_62_2_3;
                textBox_62_2_4.Text = Properties.Settings.Default.textBox_62_2_4;
                textBox_62_3_1.Text = Properties.Settings.Default.textBox_62_3_1;
                textBox_62_3_2.Text = Properties.Settings.Default.textBox_62_3_2;
                textBox_62_3_3.Text = Properties.Settings.Default.textBox_62_3_3;
                textBox_62_3_4.Text = Properties.Settings.Default.textBox_62_3_4;
                textBox_62_4_1.Text = Properties.Settings.Default.textBox_62_4_1;
                textBox_62_4_2.Text = Properties.Settings.Default.textBox_62_4_2;
                textBox_62_4_3.Text = Properties.Settings.Default.textBox_62_4_3;
                textBox_62_4_4.Text = Properties.Settings.Default.textBox_62_4_4;
                textBox_62_5_1.Text = Properties.Settings.Default.textBox_62_5_1;
                textBox_62_5_2.Text = Properties.Settings.Default.textBox_62_5_2;
                textBox_62_5_3.Text = Properties.Settings.Default.textBox_62_5_3;
                textBox_62_5_4.Text = Properties.Settings.Default.textBox_62_5_4;
                textBox_62_6_1.Text = Properties.Settings.Default.textBox_62_6_1;
                textBox_62_6_2.Text = Properties.Settings.Default.textBox_62_6_2;
                textBox_62_6_3.Text = Properties.Settings.Default.textBox_62_6_3;
                textBox_62_6_4.Text = Properties.Settings.Default.textBox_62_6_4;
                textBox_62_7_1.Text = Properties.Settings.Default.textBox_62_7_1;
                textBox_62_7_2.Text = Properties.Settings.Default.textBox_62_7_2;
                textBox_62_8_1.Text = Properties.Settings.Default.textBox_62_8_1;
                textBox_62_8_2.Text = Properties.Settings.Default.textBox_62_8_2;
                textBox_62_9_1.Text = Properties.Settings.Default.textBox_62_9_1;
                textBox_22_1.Text = Properties.Settings.Default.textBox_22_1;
                textBox_22_2.Text = Properties.Settings.Default.textBox_22_2;
                textBox_22_3.Text = Properties.Settings.Default.textBox_22_3;
                textBox_22_4.Text = Properties.Settings.Default.textBox_22_4;
                textBox_22_5.Text = Properties.Settings.Default.textBox_22_5;
                textBox_22_6.Text = Properties.Settings.Default.textBox_22_6;
                textBox_22_7.Text = Properties.Settings.Default.textBox_22_7;
                radioButton_22_1.Checked = Properties.Settings.Default.radioButton_22_1;
                radioButton_22_2.Checked = Properties.Settings.Default.radioButton_22_2;
                radioButton_22_3.Checked = Properties.Settings.Default.radioButton_22_3;
                radioButton_22_4.Checked = Properties.Settings.Default.radioButton_22_4;

            }
            catch (Exception ex)
            {
                MessageBox.Show($"自动加载时出错: {ex.Message}", "错误",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Load62(object sender, EventArgs e)
        {
            LoadDataGridView();
            AutoLoadAllTextBoxValues();
        }

        private void Closing62(object sender, FormClosingEventArgs e)
        {
            SaveDataGridView();
            AutoSaveAllTextBoxValues();
            // 保存设置
            Properties.Settings.Default.Save();
        }

        #region DataGridView 保存和加载方法

        // 保存 DataGridView 数据
        private void SaveDataGridView()
        {
            try
            {
                if (dataGridView_61_1 == null) return;

                // 保存表格结构信息
                Properties.Settings.Default.dataGridView_61_1_ColumnCount = dataGridView_61_1.Columns.Count;
                Properties.Settings.Default.dataGridView_61_1_RowCount = dataGridView_61_1.Rows.Count;

                // 保存表格数据（使用CSV格式）
                StringBuilder sb = new StringBuilder();

                // 保存列头
                for (int col = 0; col < dataGridView_61_1.Columns.Count; col++)
                {
                    if (col > 0) sb.Append("|");
                    sb.Append(EscapeCsv(dataGridView_61_1.Columns[col].HeaderText ?? ""));
                }
                sb.AppendLine();

                // 保存行数据
                foreach (DataGridViewRow row in dataGridView_61_1.Rows)
                {
                    // 跳过新行（未提交的行）
                    if (row.IsNewRow) continue;

                    for (int col = 0; col < dataGridView_61_1.Columns.Count; col++)
                    {
                        if (col > 0) sb.Append("|");
                        object cellValue = row.Cells[col].Value;
                        sb.Append(EscapeCsv(cellValue?.ToString() ?? ""));
                    }
                    sb.AppendLine();
                }

                Properties.Settings.Default.dataGridView_61_1_Data = sb.ToString();
                Console.WriteLine("dataGridView_61_1 数据保存完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"保存 dataGridView_61_1 时出错: {ex.Message}");
            }
        }

        // 加载 DataGridView 数据
        private void LoadDataGridView()
        {
            try
            {
                if (dataGridView_61_1 == null) return;

                string savedData = Properties.Settings.Default.dataGridView_61_1_Data;
                if (string.IsNullOrEmpty(savedData)) return;

                string[] lines = savedData.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                if (lines.Length < 1) return;

                // 清空现有数据
                dataGridView_61_1.Rows.Clear();

                // 解析列头（第一行）
                string[] headers = lines[0].Split('|');
                for (int i = 0; i < headers.Length; i++)
                {
                    headers[i] = UnescapeCsv(headers[i]);
                }

                // 确保列数匹配
                while (dataGridView_61_1.Columns.Count < headers.Length)
                {
                    dataGridView_61_1.Columns.Add($"Column{dataGridView_61_1.Columns.Count + 1}", $"Column{dataGridView_61_1.Columns.Count + 1}");
                }

                // 设置列头
                for (int col = 0; col < headers.Length && col < dataGridView_61_1.Columns.Count; col++)
                {
                    dataGridView_61_1.Columns[col].HeaderText = headers[col];
                }

                // 加载行数据
                for (int lineIndex = 1; lineIndex < lines.Length; lineIndex++)
                {
                    string[] cells = lines[lineIndex].Split('|');
                    int rowIndex = dataGridView_61_1.Rows.Add();

                    for (int col = 0; col < cells.Length && col < dataGridView_61_1.Columns.Count; col++)
                    {
                        dataGridView_61_1.Rows[rowIndex].Cells[col].Value = UnescapeCsv(cells[col]);
                    }
                }

                Console.WriteLine("dataGridView_61_1 数据加载完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"加载 dataGridView_61_1 时出错: {ex.Message}");
            }
        }

        // CSV 转义处理
        private string EscapeCsv(string input)
        {
            if (string.IsNullOrEmpty(input)) return "";

            // 如果包含分隔符或换行符，进行转义
            if (input.Contains("|") || input.Contains("\r") || input.Contains("\n") || input.Contains("\""))
            {
                return "\"" + input.Replace("\"", "\"\"") + "\"";
            }
            return input;
        }

        // CSV 反转义处理
        private string UnescapeCsv(string input)
        {
            if (string.IsNullOrEmpty(input)) return "";

            if (input.StartsWith("\"") && input.EndsWith("\""))
            {
                input = input.Substring(1, input.Length - 2);
                input = input.Replace("\"\"", "\"");
            }
            return input;
        }








        #endregion

        #endregion

        #region



        #endregion

    }


    //新建ListBox1专用item类
    class CheckableItem
    {
        public string Text { get; set; }
        public bool Checked { get; set; }
        public int DisplayIndex { get; set; } // 新增：1-based 编号（1~96）
        public override string ToString() => Text; // 用于简单显示
    }


    //新建探头类
    public class ProbeInfo
    {
        public string ProbeName { get; set; }
        public string DeviceAddress { get; set; }  //地址=IP:port
        public string DeviceType { get; set; }     // 对应"设备号"列 (CH1/CH2)
        public int ProbeChannel { get; set; }      // 对应"探头"列 (1-8)
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }

        public override string ToString()
        {
            return $"{ProbeName},{X:F3},{Y:F3},{Z:F3}";
        }
    }


    // 图表数据批量更新类
    public class PlotUpdatePackage
    {
        public Dictionary<string, List<double>> NewPoints { get; } = new Dictionary<string, List<double>>();
        public Dictionary<string, (double X, double Y, double Z)> LatestValues { get; } = new Dictionary<string, (double, double, double)>();
    }

    // 探头数据结构
    public class ProbeData
    {
        public string DeviceId { get; set; }
        public int ProbeNumber { get; set; }
        public string Address { get; set; }
        public int ProbeIndex { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
        public DateTime LastUpdate { get; set; }
    }

    // 新增InvokeAsync扩展方法
    public static class ControlExtensions
    {
        public static Task InvokeAsync(this Control control, Action action)
        {
            var tcs = new TaskCompletionSource<object>();
            control.BeginInvoke((MethodInvoker)(() =>
            {
                try { action(); tcs.SetResult(null); }
                catch (Exception ex) { tcs.SetException(ex); }
            }));
            return tcs.Task;
        }
    }
}
