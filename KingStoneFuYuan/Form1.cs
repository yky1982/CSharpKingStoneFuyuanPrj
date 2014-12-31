using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Runtime.InteropServices;
using CommDriveBaseXML;

namespace KingStoneFuYuan
{
    public partial class Form1 : Form
    {

        #region Com
        [DllImport("CommDriveBaseXML.dll", EntryPoint = "Init", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int Init(string XMLFilePath, string ProtocolType, string DriveType);

        [DllImport("CommDriveBaseXML.dll", EntryPoint = "GetCommState", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void GetCommState(ref bool CommState);

        [DllImport("CommDriveBaseXML.dll", EntryPoint = "WriteBlockData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int WriteBlockData(int BlockNum, byte[] data);

        [DllImport("CommDriveBaseXML.dll", EntryPoint = "ReadData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadData(string RegName, ref bool data, ref DateTime dt);

        [DllImport("CommDriveBaseXML.dll", EntryPoint = "ReadData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadData(string RegName, ref byte data, ref DateTime dt);

        [DllImport("CommDriveBaseXML.dll", EntryPoint = "ReadData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadData(string RegName, ref sbyte data, ref DateTime dt);

        [DllImport("CommDriveBaseXML.dll", EntryPoint = "ReadData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadData(string RegName, ref Int16 data, ref DateTime dt);

        [DllImport("CommDriveBaseXML.dll", EntryPoint = "ReadData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadData(string RegName, ref UInt16 data, ref DateTime dt);

        [DllImport("CommDriveBaseXML.dll", EntryPoint = "ReadData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadData(string RegName, ref Int32 data, ref DateTime dt);

        [DllImport("CommDriveBaseXML.dll", EntryPoint = "ReadData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadData(string RegName, ref UInt32 data, ref DateTime dt);

        [DllImport("CommDriveBaseXML.dll", EntryPoint = "ReadData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int ReadData(string RegName, ref float data, ref DateTime dt);

        [DllImport("CommDriveBaseXML.dll", EntryPoint = "WriteData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int WriteData(string RegName, bool Data);

        [DllImport("CommDriveBaseXML.dll", EntryPoint = "WriteData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int WriteData(string RegName, byte Data);

        [DllImport("CommDriveBaseXML.dll", EntryPoint = "WriteData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int WriteData(string RegName, sbyte Data);

        [DllImport("CommDriveBaseXML.dll", EntryPoint = "WriteData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int WriteData(string RegName, Int16 Data);

        [DllImport("CommDriveBaseXML.dll", EntryPoint = "WriteData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int WriteData(string RegName, UInt16 Data);

        [DllImport("CommDriveBaseXML.dll", EntryPoint = "WriteData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int WriteData(string RegName, Int32 Data);

        [DllImport("CommDriveBaseXML.dll", EntryPoint = "WriteData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int WriteData(string RegName, UInt32 Data);

        [DllImport("CommDriveBaseXML.dll", EntryPoint = "WriteData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern int WriteData(string RegName, float Data);
        #endregion
        public IComm m_PLCCommHandle = new IComm();

        private string m_XMLFilePath = Application.StartupPath + "\\SystemFile\\xmlConfig.xml";
        private string m_DriveType = "";
        private string m_ProtocolType = "";

        private Form[] m_FromHandle;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Form.CheckForIllegalCrossThreadCalls = false;
            CommInit();
            FormInit();
        }


        private void CommInit()
        {
            m_ProtocolType = "S7200";
            m_DriveType = "Socket";
            int code = m_PLCCommHandle.Init(m_XMLFilePath, m_ProtocolType, m_DriveType);
            if (code != 1)
            {
                MessageBox.Show("初始化失败, 错误代码：" + code.ToString());
            }
        }

        private void FormInit()
        {
            m_FromHandle = new Form[] { new Bench_1(this, Grp_No1), new Bench_2(this, Grp_No2) };
            m_FromHandle[0].TopLevel = false;
            Grp_No1.Controls.Add(m_FromHandle[0]);
            m_FromHandle[0].Show();
            m_FromHandle[1].TopLevel = false;
            Grp_No2.Controls.Add(m_FromHandle[1]);
            m_FromHandle[1].Show();
            m_FromHandle[1].TopLevel = false;
        }
    }
}
