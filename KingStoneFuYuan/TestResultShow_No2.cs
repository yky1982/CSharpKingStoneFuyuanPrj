using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Runtime.InteropServices;
using System.IO;

namespace KingStoneFuYuan
{
    public partial class TestResultShow_No2 : Form
    {
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filepath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retval, int size, string filePath);

        private Bench_2 m_BenchHandle;
        private Form1 m_MainFrameHandle;

        private System.Windows.Forms.Timer m_TimerFunc = new System.Windows.Forms.Timer();

        public TestResultShow_No2(Bench_2 Handle, Form1 MainHandle)
        {
            InitializeComponent();
            m_BenchHandle = Handle;
            m_MainFrameHandle = MainHandle;
        }

        private string strFilePath = Application.StartupPath + @"\2#\TestConfig.ini";//获取INI文件路径
        private string strSec = "TestConfig"; //INI文件名
        private bool m_isSelectSavePattern = false;//是否按了保存按钮和取消按钮
        private void TestResultShow_No2_Load(object sender, EventArgs e)
        {
            m_color = button_Save.BackColor;
            strSec = Path.GetFileNameWithoutExtension(strFilePath);
            textBox_TestNo.Text = ContentValue(strSec, "TestNo_Bench2");

            m_TimerFunc.Interval = 100;
            m_TimerFunc.Tick += new EventHandler(TimerFun);
            m_TimerFunc.Start();

            m_isSelectSavePattern = m_BenchHandle.m_isSaveDataInTestResultShowForm;
            if (m_isSelectSavePattern)
            {
                button_Save.Enabled = false;
                button_Cancel.Enabled = false;
            }
            else
            {
                button_Save.Enabled = true;
                button_Cancel.Enabled = true;
            }
        }

        private string ContentValue(string Section, string key)
        {
            StringBuilder temp = new StringBuilder(1024);
            GetPrivateProfileString(Section, key, "", temp, 1024, strFilePath);
            return temp.ToString();
        }

        private int m_TimerCount = 0;
        private void TimerFun(object o, EventArgs e)
        {
            m_TimerCount++;
            if (m_TimerCount % 10 == 0)
            {
                SampleDataFun();
            }
            if (m_TimerCount % 3 == 0)
            {
                SampleSaveFun();
            }
            if (m_TimerCount >= 10)
            {
                m_TimerCount = 0;
            }
        }

        private float m_InitPressure_No1 = 0;
        private float m_InitPressure_No2 = 0;
        private float m_InitPressure_No3 = 0;
        private float m_EndPressure_No1 = 0;
        private float m_EndPressure_No2 = 0;
        private float m_EndPressure_No3 = 0;
        private float m_DropPressure_No1 = 0;
        private float m_DropPressure_No2 = 0;
        private float m_DropPressure_No3 = 0;
        private float m_KeepPressureTime_No1 = 0;
        private float m_KeepPressureTime_No2 = 0;
        private float m_KeepPressureTime_No3 = 0;
        private float m_TestTimes = 0;
        private bool m_ButtonBackColor = false;
        private void SampleDataFun()
        {
            float InitPressure_No1 = default(float);
            DateTime dt = default(DateTime);
            int code = m_MainFrameHandle.m_PLCCommHandle.ReadData("InitPressure_No1_Bench2", ref InitPressure_No1, ref dt);
            if (code != 1)
            {
                return;
            }
            m_InitPressure_No1 = InitPressure_No1;
            textBox_InitPressure_No1.Text = InitPressure_No1.ToString("0.00");

            float InitPressure_No2 = default(float);
            dt = default(DateTime);
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("InitPressure_No2_Bench2", ref InitPressure_No2, ref dt);
            if (code != 1)
            {
                return;
            }
            m_InitPressure_No2 = InitPressure_No2;
            textBox_InitPressure_No2.Text = InitPressure_No2.ToString("0.00");

            float InitPressure_No3 = default(float);
            dt = default(DateTime);
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("InitPressure_No3_Bench2", ref InitPressure_No3, ref dt);
            if (code != 1)
            {
                return;
            }
            m_InitPressure_No3 = InitPressure_No3;
            textBox_InitPressure_No3.Text = InitPressure_No3.ToString("0.00");

            float EndPressure_No1 = default(float);
            dt = default(DateTime);
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("EndPressure_No1_Bench2", ref EndPressure_No1, ref dt);
            if (code != 1)
            {
                return;
            }
            m_EndPressure_No1 = EndPressure_No1;
            textBox_EndPressure_No1.Text = EndPressure_No1.ToString("0.00");

            float EndPressure_No2 = default(float);
            dt = default(DateTime);
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("EndPressure_No2_Bench2", ref EndPressure_No2, ref dt);
            if (code != 1)
            {
                return;
            }
            m_EndPressure_No2 = EndPressure_No2;
            textBox_EndPressure_No2.Text = EndPressure_No2.ToString("0.00");

            float EndPressure_No3 = default(float);
            dt = default(DateTime);
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("EndPressure_No3_Bench2", ref EndPressure_No3, ref dt);
            if (code != 1)
            {
                return;
            }
            m_EndPressure_No3 = EndPressure_No3;
            textBox_EndPressure_No3.Text = EndPressure_No3.ToString("0.00");

            float DropPressure_No1 = default(float);
            dt = default(DateTime);
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("DropPressure_No1_Bench2", ref DropPressure_No1, ref dt);
            if (code != 1)
            {
                return;
            }
            m_DropPressure_No1 = DropPressure_No1;
            textBox_DropPressure_No1.Text = DropPressure_No1.ToString("0.00");

            float DropPressure_No2 = default(float);
            dt = default(DateTime);
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("DropPressure_No2_Bench2", ref DropPressure_No2, ref dt);
            if (code != 1)
            {
                return;
            }
            m_DropPressure_No2 = DropPressure_No2;
            textBox_DropPressure_No2.Text = DropPressure_No2.ToString("0.00");

            float DropPressure_No3 = default(float);
            dt = default(DateTime);
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("DropPressure_No3_Bench2", ref DropPressure_No3, ref dt);
            if (code != 1)
            {
                return;
            }
            m_DropPressure_No3 = DropPressure_No3;
            textBox_DropPressure_No3.Text = DropPressure_No3.ToString("0.00");

            float KeepPressureTime_No1 = default(float);
            dt = default(DateTime);
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepTime_No1_Bench2", ref KeepPressureTime_No1, ref dt);
            if (code != 1)
            {
                return;
            }
            m_KeepPressureTime_No1 = KeepPressureTime_No1;
            textBox_KeepPressureTime_No1.Text = KeepPressureTime_No1.ToString("0.00");

            float KeepPressureTime_No2 = default(float);
            dt = default(DateTime);
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepTime_No2_Bench2", ref KeepPressureTime_No2, ref dt);
            if (code != 1)
            {
                return;
            }
            m_KeepPressureTime_No2 = KeepPressureTime_No2;
            textBox_KeepPressureTime_No2.Text = KeepPressureTime_No2.ToString("0.00");

            float KeepPressureTime_No3 = default(float);
            dt = default(DateTime);
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepTime_No3_Bench2", ref KeepPressureTime_No3, ref dt);
            if (code != 1)
            {
                return;
            }
            m_KeepPressureTime_No3 = KeepPressureTime_No3;
            textBox_KeepPressureTime_No3.Text = KeepPressureTime_No3.ToString("0.00");

            UInt16 TestTimes = default(UInt16);
            dt = default(DateTime);
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("TestTimes_No1_Bench2", ref TestTimes, ref dt);
            if (code != 1)
            {
                return;
            }
            m_TestTimes = TestTimes;
            textBox_Times.Text = TestTimes.ToString();
            if (TestTimes == m_BenchHandle.m_SetTestTimes)
            {
                //m_SampleDateTimer.Stop();
                //m_SampleSaveTimer.Stop();
            }

            //保存按钮闪烁
            if (m_ButtonBackColor)
            {
                button_Save.BackColor = Color.YellowGreen;
                m_ButtonBackColor = false;
            }
            else
            {
                m_ButtonBackColor = true;
                button_Save.BackColor = Color.LimeGreen;
            }
            button_Save.BackColor = Color.Green;
            DateTime dt1 = m_BenchHandle.m_StartTime;
            textBox_Time.Text = dt1.ToString("yyyy/MM/dd  HH:mm:ss");
        }

        private void SampleSaveFun()
        {
            bool saveFlag = false;
            DateTime dt = default(DateTime);

            int code = m_MainFrameHandle.m_PLCCommHandle.ReadData("ReadSaveFlag_Bench2", ref saveFlag, ref dt);

            if (saveFlag)
            {
                m_TimerFunc.Stop();
                SavePara();
                saveFlag = false;
                code = m_MainFrameHandle.m_PLCCommHandle.WriteData("ReadSaveFlag_Bench2", saveFlag);
                m_BenchHandle.Show();
                this.Hide();
                //this.Dispose();
            }
            bool CancelSaveFlag = false;
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("SaveFlag_Bench2", ref CancelSaveFlag, ref dt);
            if (CancelSaveFlag)
            {
                m_TimerFunc.Stop();
                m_BenchHandle.m_PointFArrays.Clear();
                //m_BenchHandle.m_PointFLists.Clear();
                CancelSaveFlag = false;
                code = m_MainFrameHandle.m_PLCCommHandle.WriteData("SaveFlag_Bench2", CancelSaveFlag);
                m_BenchHandle.Show();
                this.Hide();
                //this.Dispose();
            }
        }

        private Color m_color;

        private void button_Save_Click(object sender, EventArgs e)
        {
            if (m_BenchHandle.m_isTesting)
            {
                return;
            }

            m_TimerFunc.Stop();
            SavePara();
            button_Save.BackColor = m_color;
            button_Save.Enabled = false;
            button_Cancel.Enabled = false;
            m_isSelectSavePattern = true;
            m_BenchHandle.m_isSaveDataInTestResultShowForm = m_isSelectSavePattern;
            //button_Return.Enabled = true;
        }

        private void button_Return_Click(object sender, EventArgs e)
        {
            if (m_BenchHandle.m_isTesting)
            {
                m_BenchHandle.Show();
                this.Close();
                return;
            }
            if (!m_isSelectSavePattern)
            {
                MessageBox.Show("请选择“保存”或是“取消保存”", "Info", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                return;
            }
            m_BenchHandle.Show();
            //button_Return.Enabled = false;
            this.Close();
            //this.Dispose();
        }

        private void SavePara()
        {
            string testNo = textBox_TestNo.Text;
            int TestSequence = m_BenchHandle.m_TestSequence;
            if (testNo == m_BenchHandle.m_TestNo)
            {
                // m_BenchHandle.AddPonitFtoList();
                if (m_TestTimes >= 1)
                {
                    string No = TestSequence.ToString() + "--1";
                    m_BenchHandle.m_TestResultLists.Add(new Bench_2.TestResultList(No, m_InitPressure_No1, m_EndPressure_No1, m_KeepPressureTime_No1, m_DropPressure_No1));
                }
                if (m_TestTimes >= 2)
                {
                    string No = TestSequence.ToString() + "--2";
                    m_BenchHandle.m_TestResultLists.Add(new Bench_2.TestResultList(No, m_InitPressure_No2, m_EndPressure_No2, m_KeepPressureTime_No2, m_DropPressure_No2));
                }
                if (m_TestTimes >= 3)
                {
                    string No = TestSequence.ToString() + "--3";
                    m_BenchHandle.m_TestResultLists.Add(new Bench_2.TestResultList(No, m_InitPressure_No3, m_EndPressure_No3, m_KeepPressureTime_No3, m_DropPressure_No3));
                }
                if (m_TestTimes >= 4)
                {
                    //不存在

                }
                TestSequence++;
            }
            else
            {
                m_BenchHandle.m_TestNo = testNo;
                m_BenchHandle.m_TestResultLists.Clear();
                m_BenchHandle.m_DrawCurveHandle.ClearAllPointArrays();
                m_BenchHandle.m_TestSequence = 1;

                //m_BenchHandle.AddPonitFtoList();
                if (m_TestTimes >= 1)
                {
                    string No = TestSequence.ToString() + "--1";
                    m_BenchHandle.m_TestResultLists.Add(new Bench_2.TestResultList(No, m_InitPressure_No1, m_EndPressure_No1, m_KeepPressureTime_No1, m_DropPressure_No1));
                }
                if (m_TestTimes >= 2)
                {
                    string No = TestSequence.ToString() + "--2";
                    m_BenchHandle.m_TestResultLists.Add(new Bench_2.TestResultList(No, m_InitPressure_No2, m_EndPressure_No2, m_KeepPressureTime_No2, m_DropPressure_No2));
                }
                if (m_TestTimes >= 3)
                {
                    string No = TestSequence.ToString() + "--3";
                    m_BenchHandle.m_TestResultLists.Add(new Bench_2.TestResultList(No, m_InitPressure_No3, m_EndPressure_No3, m_KeepPressureTime_No3, m_DropPressure_No3));
                }
                if (m_TestTimes >= 4)
                {

                }
                TestSequence++;
            }
            m_BenchHandle.m_TestSequence = TestSequence;

            //m_BenchHandle.m_DrawCurveHandle.SavePointListsToArray();//将本次试验的试验点添加至列表中
            //m_BenchHandle.m_DrawCurveHandle.ClearSourcePointF();
            m_BenchHandle.SavePointsToArray();
            m_BenchHandle.m_PointFArrays.Clear();//清除本次试验点
        }

        private void button_Cancel_Click(object sender, EventArgs e)
        {
            if (m_BenchHandle.m_isTesting)
            {
                return;
            }
            m_TimerFunc.Stop();
            button_Save.BackColor = m_color;
            button_Save.Enabled = false;
            button_Cancel.Enabled = false;
            m_isSelectSavePattern = true;
            m_BenchHandle.m_isSaveDataInTestResultShowForm = m_isSelectSavePattern;
            m_BenchHandle.m_PointFArrays.Clear();
            m_BenchHandle.ClearPointList();
        }
    }
}
