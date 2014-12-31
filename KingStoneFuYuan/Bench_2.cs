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
using DrawCurve_BaseFunction;

namespace KingStoneFuYuan
{
    public partial class Bench_2 : Form
    {
        #region INI
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filepath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retval, int size, string filePath);
        #endregion

        #region DrawCurve
        [DllImport("DrawCurve.dll", EntryPoint = "GDIInit", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void GDIInit(Panel panel, string CurveType, Color col, float LineWidth, int Gridwidth, int Gridheight);

        [DllImport("DrawCurve.dll", EntryPoint = "OpenGLInit", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void OpenGLInit(Point StartPos, Point EndPos, int Length, int Height, float MaxLength, float MaxHeight, string CurveType, float LineWidth);

        [DllImport("DrawCurve.dll", EntryPoint = "SetModule", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void SetModule(int DrawType);

        [DllImport("DrawCurve.dll", EntryPoint = "SaveSourcePointF", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void SaveSourcePointF(PointF[] pt);

        [DllImport("DrawCurve.dll", EntryPoint = "ClearSourcePointF", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void ClearSourcePointF();

        [DllImport("DrawCurve.dll", EntryPoint = "GetSourcePointF", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void GetSourcePointF(out PointF[] pt);

        [DllImport("DrawCurve.dll", EntryPoint = "DrawCurve", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void DrawCurve(Panel panel, Color col, float xMax, float yMax);

        [DllImport("DrawCurve.dll", EntryPoint = "ZoomOutBasePoint", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void ZoomOutBasePoint(Panel panel, Point StartPt, Point EndPt);

        [DllImport("DrawCurve.dll", EntryPoint = "ZoomInBasePoint", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void ZoomInBasePoint(Panel panel, Point StartPt, Point EndPt);

        [DllImport("DrawCurve.dll", EntryPoint = "ZoomOutBaseXY", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void ZoomOutBaseXY(Panel panel, float StartX, float StartY, float EndX, float EndY);

        [DllImport("DrawCurve.dll", EntryPoint = "ZoomInBaseXY", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void ZoomInBaseXY(Panel panel, float StartX, float StartY, float EndX, float EndY);

        #endregion

        private string strFilePath = Application.StartupPath + @"\2#\TestConfig.ini";//获取INI文件路径
        private string strSec = "TestConfig"; //INI文件名

        private Form1 m_MainFrameHandle;
        private GroupBox m_GrpParent;
        private SettingForm_No2 m_SettingFormHandle;
        private TestResultShow_No2 m_TestResultHandle;
        private CreateTestResult_No2 m_CreateTestResult_No1Handle;
        public UInt16 m_SetTestTimes = 0;
        public UInt32 m_yMax = 150;

        public string m_TestNo = default(string);
        public string m_GetTestNo
        {
            get
            {
                return m_TestNo;
            }
            set
            {
                m_TestNo = value;
            }
        }
        public int m_TestSequence = 1;//试验顺序；

        public DrawCurvesInOneMap m_DrawCurveHandle = new DrawCurvesInOneMap();

        private System.Windows.Forms.Timer m_FuncTimer = new System.Windows.Forms.Timer();

        public class PointFArray
        {
            public PointF m_Pt;
            public PointFArray(PointF pt)
            {
                m_Pt = pt;
            }
        }
        public List<PointFArray> m_PointFArrays = new List<PointFArray>();
        private object O_LockPointFArray = new object();
        private bool m_isFirtSampleDate = true;

        public class TestResultList
        {
            public string m_No;
            public float m_InitPressure;
            public float m_EndPressure;
            public float m_KeepTime;
            public float m_DropPressure;
            public TestResultList(string No, float IPre, float EPre, float KTime, float DPre)
            {
                m_No = No;
                m_InitPressure = IPre;
                m_EndPressure = EPre;
                m_KeepTime = KTime;
                m_DropPressure = DPre;
            }
        }
        public List<TestResultList> m_TestResultLists = new List<TestResultList>();

        public Bench_2(Form1 Handle, GroupBox grp)
        {
            InitializeComponent();
            Bench_1.CheckForIllegalCrossThreadCalls = false;
            m_MainFrameHandle = Handle;
            m_GrpParent = grp;
        }

        private void Bench_2_Load(object sender, EventArgs e)
        {
            m_SettingFormHandle = new SettingForm_No2(m_MainFrameHandle, this);
            m_SettingFormHandle.TopLevel = false;
            m_GrpParent.Controls.Add(m_SettingFormHandle);


            m_DrawCurveHandle.GDIInit(panel_No1, "Line", Color.Red, 1.0f, 20, 10);
            m_DrawCurveHandle.SetModule(1);//单条曲线

            m_FuncTimer.Interval = 100;
            m_FuncTimer.Tick += new EventHandler(TimerFun);
            m_FuncTimer.Enabled = true;

            string strSec = Path.GetFileNameWithoutExtension(strFilePath);
            m_TestNo = ContentValue(strSec, "TestNo_Bench1");

            BT_Start.Enabled = false;

            textBox_yMax.Text = m_yMax.ToString();
            label5.Visible = false;
            textBox_S1.Visible = false;
        }

        private void Bench_2_Shown(object sender, EventArgs e)
        {

        }

        private string ContentValue(string Section, string key)
        {
            StringBuilder temp = new StringBuilder(1024);
            GetPrivateProfileString(Section, key, "", temp, 1024, strFilePath);
            return temp.ToString();
        }

        private int m_TimerCount = 0;//时间计数器
        private bool m_isStartFun = true;
        private bool m_isStopFun = false;
        private bool m_isSampleDataFun = false;
        private bool m_isDrawCurveFun = false;
        private bool m_isSampleStateFun = true;
        private void TimerFun(object o, EventArgs e)
        {
            m_TimerCount++;
            //采集开始信号,150ms
            if (m_TimerCount % 10 == 0 && m_isStartFun)
            {
                SampleStartFun();
            }

            //采集停止信号
            if (m_TimerCount % 11 == 0 && m_isStopFun)
            {
                SampleStopFun();
            }

            //采集状态
            if (m_TimerCount % 13 == 0 && m_isSampleStateFun)
            {
                SampleSystemStatusFun();
            }

            //采集数据
            if (m_TimerCount % 2 == 0 && m_isSampleDataFun)
            {
                SampleDataFun();
            }

            //绘图
            if (m_TimerCount % 16 == 0 && m_isDrawCurveFun)
            {
                DrawLineFun();
            }

            if (m_TimerCount >= 16)
            {
                m_TimerCount = 0;
            }
        }


        private bool m_isStartSample = true;
        private void SampleDataFun()
        {
            float TestData = 0;
            DateTime dt = DateTime.Now;

            float setPressure = 0.0f;
            Int16 currentStage = 0;
            int code = m_MainFrameHandle.m_PLCCommHandle.ReadData("CurrentStage_Bench2", ref currentStage, ref dt);
            if (code != 1)
            {
                return;
            }
            if (currentStage == 1)
            {
                code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepPressure_No1_Bench2", ref setPressure, ref dt);
                if (code != 1)
                {
                    return;
                }
            }
            if (currentStage == 2)
            {
                code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepPressure_No2_Bench2", ref setPressure, ref dt);
                if (code != 1)
                {
                    return;
                }
            }
            if (currentStage == 3)
            {
                code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepPressure_No3_Bench2", ref setPressure, ref dt);
                if (code != 1)
                {
                    return;
                }
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("PressureTest_Bench2", ref TestData, ref dt);
            if (code != 1)
            {
                return;
            }

            if (m_isFirtSampleDate)
            {
                m_BaseTime = dt;
                m_isFirtSampleDate = false;
                TestData = 0.3f;
                m_DrawCurveTimeSpan = 30;
            }
            if (m_isStartSample)
            {
                TestData = 0.3f;
                m_DrawCurveTimeSpan = 30;
                m_isStartSample = false;
            }

            if (TestData <= 0.3)
            {
                TestData = 0.3f;
            }

            if ((TestData - setPressure) > (0.1 * setPressure))
            {
                return;
            }

            float TotalTime = 0;
            GetTimeSpan(dt, m_BaseTime, ref TotalTime);
            PointF temp = default(PointF);
            temp.X = TotalTime;
            temp.Y = TestData;
            lock (O_LockPointFArray)
            {
                m_PointFArrays.Add(new PointFArray(temp));
            }
            lock (o_EndTimeLock)
            {
                m_EndTime = dt;
            }

            float KeepPressureTime = 0.0f;
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepPressureTime_Bench2", ref KeepPressureTime, ref dt);
            if (code != 1)
            {
                return;
            }
            TextBox_KeepPressureTime.Text = KeepPressureTime.ToString("0.00");

            float DropPressure = 0.0f;
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("DropPressure_Bench2", ref DropPressure, ref dt);
            if (code != 1)
            {
                return;
            }
            TextBox_KeepPressureDrop.Text = DropPressure.ToString("0.00");
        }

        public void ReadEndTime(ref DateTime EndTime)
        {
            lock (o_EndTimeLock)
            {
                EndTime = m_EndTime;
            }
        }

        public DateTime m_BaseTime;
        public bool m_isSetFlag = false;
        private void SampleSystemStatusFun()
        {
            bool isCommunication = false;
            m_MainFrameHandle.m_PLCCommHandle.GetCommState(ref isCommunication);
            if (!isCommunication)
            {
                TextBox_Status.Text = "通信故障";
                TextBox_Status.BackColor = Color.Red;
                return;
            }
            else
            {
                TextBox_Status.BackColor = Color.Black;
            }

            bool SystemStatusManual = false;
            DateTime dt = DateTime.Now;
            int code = m_MainFrameHandle.m_PLCCommHandle.ReadData("SystemStatusManual_No1_Bench2", ref SystemStatusManual, ref dt);
            if (code != 1)
            {
                return;
            }
            if (SystemStatusManual)
            {
                TextBox_Status.Text = "手动";
            }

            bool SystemStatusEmergencyStop = false;
            dt = DateTime.Now;
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("SystemStatusEmergencyStop_No1_Bench2", ref SystemStatusEmergencyStop, ref dt);
            if (code != 1)
            {
                return;
            }
            if (SystemStatusEmergencyStop)
            {
                TextBox_Status.Text = "急停";
            }

            bool SystemStatusStop = false;
            dt = DateTime.Now;
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("SystemStatusStop_No1_Bench2", ref SystemStatusStop, ref dt);
            if (code != 1)
            {
                return;
            }
            if (SystemStatusStop)
            {
                TextBox_Status.Text = "停止";
                BT_Start.Enabled = true;
            }

            bool SystemStatusAddPressure = false;
            dt = DateTime.Now;
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("SystemStatusAddPressure_No1_Bench2", ref SystemStatusAddPressure, ref dt);
            if (code != 1)
            {
                return;
            }
            if (SystemStatusAddPressure)
            {
                TextBox_Status.Text = "加压";
            }

            bool SystemStatusKeepPressure = false;
            dt = DateTime.Now;
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("SystemStatusKeepPressure_No1_Bench2", ref SystemStatusKeepPressure, ref dt);
            if (code != 1)
            {
                return;
            }
            if (SystemStatusKeepPressure)
            {
                TextBox_Status.Text = "保压";
            }

            bool SystemStatusDropPressure = false;
            dt = DateTime.Now;
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("SystemStatusDropPressure_No1_Bench2", ref SystemStatusDropPressure, ref dt);
            if (code != 1)
            {
                return;
            }
            if (SystemStatusDropPressure)
            {
                TextBox_Status.Text = "泄压";
            }

            float PressureTest = 0.0f;
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("PressureTest_Bench2", ref PressureTest, ref dt);
            if (code != 1)
            {
                return;
            }
            TextBox_TestPressure.Text = PressureTest.ToString("0.00");

            if (m_isSetFlag)
            {
                BT_Start.Enabled = true;
                m_isSetFlag = false;
            }

            ReadSettingInfo();
            ReadChanelInfo();
        }

        private void GetTimeSpan(DateTime now, DateTime BaseTime, ref float TotalTime)
        {
            TimeSpan tp = now - BaseTime;
            TotalTime = (tp.Days * 24 * 3600) +
                        (tp.Hours * 3600) +
                        (tp.Minutes * 60) +
                        (tp.Seconds) +
                        tp.Milliseconds / 1000.0f;
        }

        public bool m_isTesting = false;//是否正在测试
        private void SampleStartFun()
        {
            bool startFlag = false;
            DateTime dt = DateTime.Now;
            int code = m_MainFrameHandle.m_PLCCommHandle.ReadData("ReadStartTest_Bench2", ref startFlag, ref dt);
            if (code != 1)
            {
                return;
            }

            if (startFlag)
            {
                m_isSampleDataFun = true;
                //for (int i = 0; i < 300000000; i++) ;
                m_isDrawCurveFun = true;
                m_isStartFun = false;
                m_isStopFun = true;

                if (m_isTestStartBaseTestNo)
                {
                    m_StartTime = DateTime.Now;
                    m_isTestStartBaseTestNo = false;
                }

                label_X1.Text = m_BaseTime.ToString("HH:mm:ss");
                DateTime d1 = m_BaseTime;
                d1.AddMinutes(m_DrawCurveTimeSpan / 60);
                label_X7.Text = d1.ToString("HH:mm:ss");
                BT_Start.Enabled = false;
                m_isTesting = true;
                m_isSaveDataInTestResultShowForm = false;
            }

            startFlag = false;
            code = m_MainFrameHandle.m_PLCCommHandle.WriteData("ReadStartTest_Bench2", startFlag);
        }

        private void SampleStopFun()
        {
            //采集PLC停止命令
            bool stopFlag = false;
            DateTime dt = DateTime.Now;
            int code = m_MainFrameHandle.m_PLCCommHandle.ReadData("ReadStopTest_Bench2", ref stopFlag, ref dt);
            if (code != 1)
            {
                return;
            }

            if (stopFlag)
            {
                m_isDrawCurveFun = false;
                m_isSampleDataFun = false;
                m_isStopFun = false;
                m_isStartFun = true;
                m_isFirtSampleDate = true;
                m_isStartSample = true;
                //label_X2.Text = DateTime.Now.ToString("HH:mm:ss");
                //CreatResultForm();
                //ShowFun();    

            }
            bool stopFlag1 = false;
            code = m_MainFrameHandle.m_PLCCommHandle.WriteData("ReadStopTest_Bench2", stopFlag1);


            if (stopFlag)
            {
                m_isTesting = false;
                BT_Start.Enabled = true;
                MethodInvoker invoke = new MethodInvoker(CreatResultForm);
                BeginInvoke(invoke);
            }
        }

        public int m_DrawCurveTimeSpan = 30;//画曲线时间宽度,初始化值为30s
        private void DrawLineFun()
        {
            int len = 0;
            m_DrawCurveHandle.SetModule(1);//单条曲线
            lock (O_LockPointFArray)
            {
                len = m_PointFArrays.Count;
            }
            if (len < 4)
            {
                return;
            }
            PointF[] p = new PointF[len];
            lock (O_LockPointFArray)
            {
                for (int i = 0; i < len; i++)
                {
                    p[i] = m_PointFArrays[i].m_Pt;
                }
                m_PointFArrays.Clear();
            }
            m_DrawCurveHandle.SaveSourcePointF(p);
            //if (m_DrawCurveTimeSpan < 5)
            //{
            //    m_DrawCurveTimeSpan = 10;
            //}
            m_DrawCurveHandle.DrawCurve(panel_No1, Color.Red, m_DrawCurveTimeSpan, m_yMax);
            if (p[len - 1].X > m_DrawCurveTimeSpan - 10)
            {
                m_DrawCurveTimeSpan += 30;//如果超过，则自加30s宽度；
            }

            label_X1.Text = m_BaseTime.ToString("HH:mm:ss");

            int xMax = m_DrawCurveTimeSpan;
            int H = 0;
            int M = 0;
            int S = 0;

            GetlableTime(xMax / 6, ref H, ref M, ref S);
            S += m_BaseTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += m_BaseTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += m_BaseTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X2.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(xMax * 2 / 6, ref H, ref M, ref S);
            S += m_BaseTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += m_BaseTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += m_BaseTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X3.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(xMax * 3 / 6, ref H, ref M, ref S);
            S += m_BaseTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += m_BaseTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += m_BaseTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X4.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(xMax * 4 / 6, ref H, ref M, ref S);
            S += m_BaseTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += m_BaseTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += m_BaseTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X5.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(xMax * 5 / 6, ref H, ref M, ref S);
            S += m_BaseTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += m_BaseTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += m_BaseTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X6.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(xMax * 6 / 6, ref H, ref M, ref S);
            S += m_BaseTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += m_BaseTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += m_BaseTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X7.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
        }

        private void GetlableTime(int t, ref int H, ref int M, ref int S)
        {
            H = (int)(t / 3600);
            M = (int)((t - H * 3600) / 60);
            S = (int)((t - H * 3600) % 60);
        }

        public DateTime m_StartTime;//本轮试验起始时间
        public bool m_isTestStartBaseTestNo = true;//基于试验编号的试验起始标志
        public DateTime m_EndTime;//本轮试压结束时间
        object o_EndTimeLock = new object();

        public bool m_isSaveDataInTestResultShowForm = false;
        private void BT_Start_Click(object sender, EventArgs e)
        {
            bool StartSignal = true;
            m_isSampleDataFun = true;
            int code = m_MainFrameHandle.m_PLCCommHandle.WriteData("StartTest_Bench2", StartSignal);
            if (code != 1)
            {
                MessageBox.Show("启动失败，请重新启动", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            TextBox_Status.Text = "开始";
            m_isSampleDataFun = true;
            m_isSaveDataInTestResultShowForm = false;

            //m_PointFArrays.Clear();//清采集数据
            if (m_isTestStartBaseTestNo)
            {
                m_StartTime = DateTime.Now;
                m_isTestStartBaseTestNo = false;
            }

            label_X1.Text = m_BaseTime.ToString("HH:mm:ss");
            DateTime d1 = m_BaseTime;

            BT_Start.Enabled = false;
        }

        private void BT_Stop_Click(object sender, EventArgs e)
        {
            bool StopSignal = true;
            int code = m_MainFrameHandle.m_PLCCommHandle.WriteData("StopTest_Bench2", StopSignal);
            if (code != 1)
            {
                MessageBox.Show("停止失败，请重新启动", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void BT_Setting_Click(object sender, EventArgs e)
        {
            if (m_isTesting)
            {
                MessageBox.Show("请先停止测试。", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            this.Hide();
            m_SettingFormHandle.Show();
            m_SettingFormHandle.FreshForm();
        }

        private void BT_Result_Click(object sender, EventArgs e)
        {
            CreatResultForm();
            this.Hide();  
        }

        public void CreatResultForm()
        {
            if (m_TestResultHandle != null)
            {
                m_TestResultHandle.Dispose();
                m_TestResultHandle = null;
            }
            m_TestResultHandle = new TestResultShow_No2(this, m_MainFrameHandle);
            m_TestResultHandle.TopLevel = false;
            m_GrpParent.Controls.Add(m_TestResultHandle);
            m_TestResultHandle.Show();
            this.Hide();
        }

        public void SavePointsToArray()
        {
            PointF[] pt;
            m_DrawCurveHandle.GetListPoint(out pt);
            if (pt.Length < 2)
            {
                return;
            }
            float xSpan = (float)(m_BaseTime - m_StartTime).TotalSeconds;

            int len = pt.Length;

            for (int i = 0; i < len; i++)
            {
                pt[i].X += xSpan;
            }
            m_DrawCurveHandle.ClearSourcePointF();
            m_DrawCurveHandle.SaveSourcePointF(pt);
            m_DrawCurveHandle.SavePointListsToArray();
            m_DrawCurveHandle.ClearSourcePointF();
        }

        public void ClearPointList()
        {
            m_DrawCurveHandle.ClearSourcePointF();
        }

        public void ClearPointBuffer()
        {
            m_DrawCurveHandle.ClearSourcePointF();
            m_DrawCurveHandle.ClearAllPointArrays();
        }

        private void BT_Report_Click(object sender, EventArgs e)
        {
            m_CreateTestResult_No1Handle = new CreateTestResult_No2(this);
            m_CreateTestResult_No1Handle.ShowDialog();

            m_CreateTestResult_No1Handle.Close();
            m_CreateTestResult_No1Handle.Dispose();
        }

        private void ReadSettingData()
        {
            bool ChanelSelect_No1 = false;
            DateTime dt = default(DateTime);
            m_MainFrameHandle.m_PLCCommHandle.ReadData("ChanelSelect_No1_Bench2", ref ChanelSelect_No1, ref dt);

            bool ChanelSelect_No2 = false;
            m_MainFrameHandle.m_PLCCommHandle.ReadData("ChanelSelect_No2_Bench2", ref ChanelSelect_No2, ref dt);

            bool ChanelSelect_No3 = false;
            m_MainFrameHandle.m_PLCCommHandle.ReadData("ChanelSelect_No3_Bench2", ref ChanelSelect_No3, ref dt);

            float ContinueTime = 0.0f;
            if (ChanelSelect_No3)
            {
                float time3 = 0;
                m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepTime_No3_Bench2", ref time3, ref dt);
                ContinueTime += time3;

                float Press3 = 0;
                m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepPressure_No3_Bench2", ref Press3, ref dt);

                WritePrivateProfileString(strSec, "KeepPressure_No3_Bench2", Press3.ToString(), strFilePath);
                WritePrivateProfileString(strSec, "KeepTime_No3_Bench2", time3.ToString(), strFilePath);
                WritePrivateProfileString(strSec, "KeepSelect_No3_Bench2", "ON", strFilePath);
            }

            if (ChanelSelect_No2)
            {
                float time2 = 0;
                m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepTime_No2_Bench2", ref time2, ref dt);
                ContinueTime += time2;

                float Press2 = 0;
                m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepPressure_No2_Bench2", ref Press2, ref dt);

                WritePrivateProfileString(strSec, "KeepPressure_No2_Bench2", Press2.ToString(), strFilePath);
                WritePrivateProfileString(strSec, "KeepTime_No2_Bench2", time2.ToString(), strFilePath);
                WritePrivateProfileString(strSec, "KeepSelect_No2_Bench2", "ON", strFilePath);
            }

            if (ChanelSelect_No1)
            {
                float time1 = 0;
                m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepTime_No1_Bench2", ref time1, ref dt);
                ContinueTime += time1;

                float Press1 = 0;
                m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepPressure_No1_Bench2", ref Press1, ref dt);

                WritePrivateProfileString(strSec, "KeepPressure_No1_Bench2", Press1.ToString(), strFilePath);
                WritePrivateProfileString(strSec, "KeepTime_No1_Bench2", time1.ToString(), strFilePath);
                WritePrivateProfileString(strSec, "KeepSelect_No1_Bench2", "ON", strFilePath);
            }

            m_DrawCurveTimeSpan = (int)(ContinueTime * 60);
        }

        private void ReadSettingInfo()
        {
            float HighBumpStartPress = 0;
            float SensorLength = 0;
            float SensorOffSet = 0.0f;
            UInt16 StabilityTime = 0;
            UInt16 OpenValveTime = 0;
            float DropPressSelect = 0;
            UInt16 TestPressInterval = 0;
            float EarlyPre = 0.0f;

            DateTime dt = default(DateTime);

            int code = m_MainFrameHandle.m_PLCCommHandle.ReadData("HighBumpStartPress_Bench2", ref HighBumpStartPress, ref dt);
            if (code != 1)
            {
                return;
            }
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("SensorLength_Bench2", ref SensorLength, ref dt);
            if (code != 1)
            {
                return;
            }
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("SensorOffSet_Bench2", ref SensorOffSet, ref dt);
            if (code != 1)
            {
                return;
            }
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepPressStabilityTime_Bench2", ref StabilityTime, ref dt);
            if (code != 1)
            {
                return;
            }
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("OpenValveTime_Bench2", ref OpenValveTime, ref dt);
            if (code != 1)
            {
                return;
            }
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("DropPressSelect_Bench2", ref DropPressSelect, ref dt);
            if (code != 1)
            {
                return;
            }
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("TestPressInterval_Bench2", ref TestPressInterval, ref dt);
            if (code != 1)
            {
                return;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("PreEarly_Bench2", ref EarlyPre, ref dt);
            if (code != 1)
            {
                return;
            }

            textBox_S1.Text = HighBumpStartPress.ToString("0.00");
            textBox_S2.Text = SensorLength.ToString("0.00");
            textBox_S3.Text = SensorOffSet.ToString("0.00");
            textBox_S4.Text = StabilityTime.ToString();
            textBox_S5.Text = OpenValveTime.ToString();
            textBox_S6.Text = DropPressSelect.ToString("0.00");
            textBox_S7.Text = TestPressInterval.ToString();
            textBox_EarlyPre.Text = EarlyPre.ToString();
        }

        private void ReadChanelInfo()
        {
            float KeepPressure1 = 0.0f;
            float KeepPressure2 = 0.0f;
            float KeepPressure3 = 0.0f;
            float KeepTime1 = 0.0f;
            float KeepTime2 = 0.0f;
            float KeepTime3 = 0.0f;
            bool ChanelSelect1 = false;
            bool ChanelSelect2 = false;
            bool ChanelSelect3 = false;

            DateTime dt = default(DateTime);

            int code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepPressure_No1_Bench2", ref KeepPressure1, ref dt);
            if (code != 1)
            {
                return;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepPressure_No2_Bench2", ref KeepPressure2, ref dt);
            if (code != 1)
            {
                return;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepPressure_No3_Bench2", ref KeepPressure3, ref dt);
            if (code != 1)
            {
                return;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepTime_No1_Bench2", ref KeepTime1, ref dt);
            if (code != 1)
            {
                return;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepTime_No2_Bench2", ref KeepTime2, ref dt);
            if (code != 1)
            {
                return;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepTime_No3_Bench2", ref KeepTime3, ref dt);
            if (code != 1)
            {
                return;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("ChanelSelect_No1_Bench2", ref ChanelSelect1, ref dt);
            if (code != 1)
            {
                return;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("ChanelSelect_No2_Bench2", ref ChanelSelect2, ref dt);
            if (code != 1)
            {
                return;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("ChanelSelect_No3_Bench2", ref ChanelSelect3, ref dt);
            if (code != 1)
            {
                return;
            }

            if (ChanelSelect1 == true)
            {
                textBox_C1.Visible = true;
                textBox_C1.Text = KeepPressure1.ToString("0.00") + "MPa; " + KeepTime1.ToString("0.00") + "Min";
            }
            else
            {
                textBox_C1.Visible = false;
            }

            if (ChanelSelect2 == true)
            {
                textBox_C2.Visible = true;
                textBox_C2.Text = KeepPressure2.ToString("0.00") + "MPa; " + KeepTime2.ToString("0.00") + "Min";
            }
            else
            {
                textBox_C2.Visible = false;
            }

            if (ChanelSelect3 == true)
            {
                textBox_C3.Visible = true;
                textBox_C3.Text = KeepPressure3.ToString("0.00") + "MPa; " + KeepTime3.ToString("0.00") + "Min";
            }
            else
            {
                textBox_C3.Visible = false;
            }
        }

        private void textBox_yMax_TextChanged(object sender, EventArgs e)
        {
            try
            {
                m_yMax = Convert.ToUInt32(textBox_yMax.Text);
            }
            catch (Exception)
            {
                //MessageBox.Show("纵轴最大值输入错误，请核对", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                m_yMax = 250;
                return;
            }
            if (m_yMax < 1)
            {
                m_yMax = 250;
            }
            UInt32 t1 = m_yMax * 1 / 5;
            label9.Text = t1.ToString() + "MPa";

            UInt32 t2 = m_yMax * 2 / 5;
            label8.Text = t2.ToString() + "MPa";

            UInt32 t3 = m_yMax * 3 / 5;
            label7.Text = t3.ToString() + "MPa";

            UInt32 t4 = m_yMax * 4 / 5;
            label4.Text = t4.ToString() + "MPa";
        }
    }
}
