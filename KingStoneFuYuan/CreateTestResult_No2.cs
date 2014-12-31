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
using Microsoft.Office.Interop.Word;
using AccessDLL__AddBuffer;

namespace KingStoneFuYuan
{
    public partial class CreateTestResult_No2 : Form
    {
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filepath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retval, int size, string filePath);

        #region AccessDll
        [DllImport("AccessDLL__AddBuffer.dll", EntryPoint = "Init", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void Init(string FileName);

        [DllImport("AccessDLL__AddBuffer.dll", EntryPoint = "CreatFile", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void CreatFile();

        [DllImport("AccessDLL__AddBuffer.dll", EntryPoint = "AddTable", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void AddTable(string TableName);

        [DllImport("AccessDLL__AddBuffer.dll", EntryPoint = "AddColumn", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void AddColumn(string TableName, string ColName, string Datatype);

        [DllImport("AccessDLL__AddBuffer.dll", EntryPoint = "AddNewRow", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void AddNewRow(string TableName);

        [DllImport("AccessDLL__AddBuffer.dll", EntryPoint = "DeleteRow", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void DeleteRow(string TableName, int index);

        [DllImport("AccessDLL__AddBuffer.dll", EntryPoint = "WriteFloatsData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void WriteFloatsData(string TableName, int index, string ColName, float[] data);

        [DllImport("AccessDLL__AddBuffer.dll", EntryPoint = "WriteIntsData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void WriteIntsData(string TableName, int index, string ColName, int[] data);

        [DllImport("AccessDLL__AddBuffer.dll", EntryPoint = "WriteSigleData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void WriteSigleData(string TableName, int index, string ColName, object data);

        [DllImport("AccessDLL__AddBuffer.dll", EntryPoint = "WriteFloatsToExistCells", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void WriteFloatsToExistCells(string TableName, int index, string ColName, float[] data);

        [DllImport("AccessDLL__AddBuffer.dll", EntryPoint = "WriteIntsToExistCells", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void WriteIntsToExistCells(string TableName, int index, string ColName, int[] data);

        [DllImport("AccessDLL__AddBuffer.dll", EntryPoint = "WriteSigleDataToExistCells", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void WriteSigleDataToExistCells(string TableName, int index, string ColName, object data);

        [DllImport("AccessDLL__AddBuffer.dll", EntryPoint = "ReadFloatsData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void ReadFloatsData(string TableName, int index, string ColName, out float[] data);

        [DllImport("AccessDLL__AddBuffer.dll", EntryPoint = "ReadIntsData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void ReadIntsData(string TableName, int index, string ColName, out int[] data);

        [DllImport("AccessDLL__AddBuffer.dll", EntryPoint = "ReadSigleData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void ReadSigleData(string TableName, int index, string ColName, ref object data);

        [DllImport("AccessDLL__AddBuffer.dll", EntryPoint = "QueryData", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void QueryData(string TableName, int index, string ColName, ref object data);

        [DllImport("AccessDLL__AddBuffer.dll", EntryPoint = "RefreshBuffer", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void RefreshBuffer();

        [DllImport("AccessDLL__AddBuffer.dll", EntryPoint = "SaveDataToBuffer", ExactSpelling = false, CallingConvention = CallingConvention.Cdecl)]
        public static extern void SaveDataToBuffer();

        #endregion

        private Bench_2 m_BenchHandle;
        private string I_ProductName;
        private string I_ProductNo;
        private string I_ProductModel;
        private string I_TestStandard;
        private string I_TestNiuju;
        private string I_EquipmentNo;
        private string I_CompanyName;
        private string I_Operator;
        private string I_Tester;
        private string I_Confirmation;
        private string I_Date;
        private string m_TestNo;

        public CreateTestResult_No2(Bench_2 handle)
        {
            InitializeComponent();
            m_BenchHandle = handle;
        }

        private UInt32 m_yMax = 0;
        private float m_xMax = 0.0f;
        private string m_BaseFilePath = System.Windows.Forms.Application.StartupPath + "\\2#";
        private PointF[] m_HistoryPointF;
        private DateTime m_HistoryStartTime;
        private DateTime m_HistoryEndTime;

        public AccessLibInterface m_AccessHandle = new AccessLibInterface();
        public string m_DataBaseFilePath = System.Windows.Forms.Application.StartupPath + "\\2#\\DataBase.mdb";

        private SelectHistoryData_No2 m_SelectHistoryData_No2Handle;
        public bool m_isNoSelect = true;
        public int m_RecordId = 0;
        private bool m_isOpenDataBase = false;
        private System.Windows.Forms.Timer m_TimerFun = new System.Windows.Forms.Timer();
        private void CreateTestResult_No2_Load(object sender, EventArgs e)
        {
            m_AccessHandle.Init(m_DataBaseFilePath, "TestResult_1", 1);

            //m_BenchHandle.m_DrawCurveHandle.SetModule(0);
            m_TestNo = m_BenchHandle.m_TestNo;

            ///调整数据单元格显示格式
            dataGridView_DateList.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView_DateList_e.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView_DateList.ReadOnly = false;
            dataGridView_DateList_e.ReadOnly = false;


            //界面信息
            m_yMax = m_BenchHandle.m_yMax;
            textBox_yMax.Text = m_yMax.ToString();
            textBox_yMax_e.Text = m_yMax.ToString();

            //初始化时，要读取的数据
            ReadTestInfo();
            ReadTestDate();
            ReadOfficeInfo();

            I_Date = DateTime.Now.ToString("yyyy-MM-dd");
            TB_Date.Text = I_Date;
            TB_Date_e.Text = I_Date;
            m_TimerFun.Interval = 1000;
            m_TimerFun.Tick += new EventHandler(DrawCurveFun);
            m_TimerFun.Start();

        }

        private void CreateTestResult_No2_Shown(object sender, EventArgs e)
        {
            for (int i = 0; i < 100000000; i++) ;
            DrawCurveCh();
            SavePicCh();
            bool isTesting = m_BenchHandle.m_isTesting;
            if (isTesting)
            {
                button_Save_Path.Enabled = false;
                button3.Enabled = false;
                button_CreatReport.Enabled = false;
                button2.Enabled = false;
            }
        }

        private int m_DrawCurveTimeCount = 0;
        private void DrawCurveFun(object o, EventArgs e)
        {
            m_DrawCurveTimeCount++;
            DrawCurveCh();
            SavePicCh();
            if (m_DrawCurveTimeCount >= 3)
            {
                m_TimerFun.Stop();
            }
        }

        private void DrawCurveCh()
        {
            bool isTesting = m_BenchHandle.m_isTesting;
            if (isTesting)
            {
                return;
            }
            m_BenchHandle.m_DrawCurveHandle.SetModule(0);
            PointF[] pt;
            m_BenchHandle.m_DrawCurveHandle.GetSourcePointF(out pt);
            if (pt.Length <= 2)
            {
                return;
            }
            int len = pt.Length;
            if (pt[len - 1].X == 0)
            {
                m_xMax = pt[len - 2].X;
            }
            else
            {
                m_xMax = pt[len - 2].X;
            }
            m_BenchHandle.m_DrawCurveHandle.DrawCurve(panel_Result, Color.Red, m_xMax, m_yMax);

            DateTime StartTime = m_BenchHandle.m_StartTime;
            DateTime EndTime = default(DateTime);
            m_BenchHandle.ReadEndTime(ref EndTime);
            TimeSpan tp = EndTime - StartTime;
            int second = (int)tp.TotalSeconds;
            label_X1.Text = StartTime.ToString("HH:mm:ss");
            textBox_StartTime.Text = label_X1.Text;

            int H = 0;
            int M = 0;
            int S = 0;
            GetlableTime(second / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X2.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(second * 2 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X3.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(second * 3 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X4.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(second * 4 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X5.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(second * 5 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X6.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(second * 6 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X7.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            label_X7.Text = EndTime.ToString("HH:mm:ss");
            textBox_EndTime.Text = label_X7.Text;

            UInt32 t1 = m_yMax * 1 / 5;
            label32.Text = t1.ToString() + "MPa";

            UInt32 t2 = m_yMax * 2 / 5;
            label33.Text = t2.ToString() + "MPa";

            UInt32 t3 = m_yMax * 3 / 5;
            label34.Text = t3.ToString() + "MPa";

            UInt32 t4 = m_yMax * 4 / 5;
            label36.Text = t4.ToString() + "MPa";

        }

        private void GetlableTime(int t, ref int H, ref int M, ref int S)
        {
            H = (int)(t / 3600);
            M = (int)((t - H * 3600) / 60);
            S = (int)((t - H * 3600) % 60);
        }

        private void SavePicCh()
        {
            int width = 630;
            int height = 335;
            Bitmap image = new Bitmap(width, height);

            Graphics g = Graphics.FromImage(image);
            Point p1 = default(Point);
            Point p2 = default(Point);
            p1.X = 0;
            p1.Y = 0;
            p2.X = this.Left + 58;
            p2.Y = this.Top + 337;

            g.CopyFromScreen(p2, p1, image.Size);

            string FileName = m_BaseFilePath + "\\Presure.png";
            image.Save(FileName);

            return;
        }

        private void DrawHisCurveCh()
        {
            m_BenchHandle.m_DrawCurveHandle.SetModule(2);
            m_BenchHandle.m_DrawCurveHandle.ClearALLPointF();
            m_BenchHandle.m_DrawCurveHandle.SaveAllPointF(m_HistoryPointF);
            float xMax = 0;
            float yMax = m_yMax;
            int len = m_HistoryPointF.Length;
            if (len < 2)
            {
                return;
            }
            if (m_HistoryPointF[len - 1].X == 0)
            {
                xMax = m_HistoryPointF[len - 2].X;
            }
            else
            {
                xMax = m_HistoryPointF[len - 2].X;
            }
            m_BenchHandle.m_DrawCurveHandle.DrawCurve(panel_Result, Color.Red, xMax, yMax);


            DateTime startTime = default(DateTime);
            DateTime endTime = default(DateTime);
            startTime = m_HistoryStartTime;
            endTime = m_HistoryEndTime;

            TimeSpan tp = endTime - startTime;
            int second = (int)tp.TotalSeconds;
            label_X1.Text = startTime.ToString("HH:mm:ss");
            textBox_StartTime.Text = label_X1.Text;

            int H = 0;
            int M = 0;
            int S = 0;
            GetlableTime(second / 6, ref H, ref M, ref S);
            S += startTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += startTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += startTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X2.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 2 / 6, ref H, ref M, ref S);
            S += startTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += startTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += startTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X3.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 3 / 6, ref H, ref M, ref S);
            S += startTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += startTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += startTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X4.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 4 / 6, ref H, ref M, ref S);
            S += startTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += startTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += startTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X5.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 5 / 6, ref H, ref M, ref S);
            S += startTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += startTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += startTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X6.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 6 / 6, ref H, ref M, ref S);
            S += startTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += startTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += startTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X7.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            label_X7.Text = endTime.ToString("HH:mm:ss");
            textBox_EndTime.Text = label_X7.Text;

            UInt32 t1 = m_yMax * 1 / 5;
            label32.Text = t1.ToString() + "MPa";

            UInt32 t2 = m_yMax * 2 / 5;
            label33.Text = t2.ToString() + "MPa";

            UInt32 t3 = m_yMax * 3 / 5;
            label34.Text = t3.ToString() + "MPa";

            UInt32 t4 = m_yMax * 4 / 5;
            label36.Text = t4.ToString() + "MPa";
        }

        private string strFilePath = System.Windows.Forms.Application.StartupPath + @"\2#\TestInfoConfig.ini";//获取INI文件路径
        private string strSec = "TestInfoConfig"; //INI文件名
        private void ReadTextBoxData()
        {
            I_ProductName = TB_Name.Text;
            I_ProductNo = TB_ProductID.Text;
            I_ProductModel = TB_Module.Text;
            I_TestStandard = TB_Standard.Text;
            I_TestNiuju = TB_Niuju.Text;
            I_EquipmentNo = TB_EquipNum.Text;
            I_CompanyName = TB_CompanyName.Text;
            I_Operator = TB_Operator.Text;
            I_Tester = TB_Tester.Text;
            I_Confirmation = TB_Conformation.Text;

            WritePrivateProfileString(strSec, "ProductName", I_ProductName, strFilePath);
            WritePrivateProfileString(strSec, "ProductNo", I_ProductNo, strFilePath);
            WritePrivateProfileString(strSec, "ProductModel", I_ProductModel, strFilePath);
            WritePrivateProfileString(strSec, "TestStandard", I_TestStandard, strFilePath);
            WritePrivateProfileString(strSec, "TestNiuju", I_TestNiuju, strFilePath);
            WritePrivateProfileString(strSec, "EquipmentNo", I_EquipmentNo, strFilePath);
            WritePrivateProfileString(strSec, "CompanyName", I_CompanyName, strFilePath);
            WritePrivateProfileString(strSec, "Operator", I_Operator, strFilePath);
            WritePrivateProfileString(strSec, "Tester", I_Tester, strFilePath);
            WritePrivateProfileString(strSec, "Confirmation", I_Confirmation, strFilePath);
            WritePrivateProfileString(strSec, "Date", I_Date, strFilePath);
        }

        private bool SaveDataToDataBase()
        {
            int code = 0;
            int index = 0;
            m_AccessHandle.AddNewRow("TestResult_1", ref index);
            try
            {
                code = m_AccessHandle.WriteSigleData("TestResult_1", index, "TestNo", m_TestNo);
                if (code != 1)
                {
                    m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                DateTime dt1 = m_BenchHandle.m_StartTime;
                string StartTime = dt1.ToString("yyyy-MM-dd  HH:mm:ss");
                code = m_AccessHandle.WriteSigleData("TestResult_1", index, "StartTime_s", StartTime);
                if (code != 1)
                {
                    m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                code = m_AccessHandle.WriteSigleData("TestResult_1", index, "StartTime", m_BenchHandle.m_StartTime);
                if (code != 1)
                {
                    m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                DateTime endTime = default(DateTime);
                m_BenchHandle.ReadEndTime(ref endTime);
                code = m_AccessHandle.WriteSigleData("TestResult_1", index, "EndTime", endTime);
                if (code != 1)
                {
                    m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                code = m_AccessHandle.WriteSigleData("TestResult_1", index, "ProductName", I_ProductName);
                if (code != 1)
                {
                    m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }
                code = m_AccessHandle.WriteSigleData("TestResult_1", index, "ProductNo", I_ProductNo);
                if (code != 1)
                {
                    m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }
                code = m_AccessHandle.WriteSigleData("TestResult_1", index, "ProductModel", I_ProductModel);
                if (code != 1)
                {
                    m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                code = m_AccessHandle.WriteSigleData("TestResult_1", index, "TestStandard", I_TestStandard);
                if (code != 1)
                {
                    m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                code = m_AccessHandle.WriteSigleData("TestResult_1", index, "TestNiuju", I_TestNiuju);
                if (code != 1)
                {
                    m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                code = m_AccessHandle.WriteSigleData("TestResult_1", index, "EquipmentNo", I_EquipmentNo);
                if (code != 1)
                {
                    m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                code = m_AccessHandle.WriteSigleData("TestResult_1", index, "CompanyName", I_CompanyName);
                if (code != 1)
                {
                    m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }
                code = m_AccessHandle.WriteSigleData("TestResult_1", index, "Operator", I_Operator);
                if (code != 1)
                {
                    m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }
                code = m_AccessHandle.WriteSigleData("TestResult_1", index, "Tester", I_Tester);
                if (code != 1)
                {
                    m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }
                code = m_AccessHandle.WriteSigleData("TestResult_1", index, "Confirmation", I_Confirmation);
                if (code != 1)
                {
                    m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }
                code = m_AccessHandle.WriteSigleData("TestResult_1", index, "TestDate", I_Date);
                if (code != 1)
                {
                    m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                int rows = dataGridView_DateList.Rows.Count;
                float[] testDate = new float[(rows - 1) * 5];
                for (int i = 1; i < rows - 1; i++)
                {
                    for (int j = 0; j < 5; j++)
                    {
                        if (j == 0)
                        {
                            string No = dataGridView_DateList[j, i].Value.ToString();
                            char[] a = No.ToCharArray();
                            string temp = "";
                            for (int m = 0; m < a.Length; m++)
                            {
                                if (a[m] == '-')
                                {
                                    continue;
                                }
                                temp += a[m];
                            }
                            testDate[(i - 1) * 5 + j] = Convert.ToSingle(temp);
                        }
                        else
                        {

                            testDate[(i - 1) * 5 + j] = Convert.ToSingle(dataGridView_DateList[j, i].Value.ToString());
                        }
                    }
                }

                //数量
                code = m_AccessHandle.WriteSigleData("TestResult_1", index, "ResultsNum", rows - 2);//减去新建的行和标题行
                if (code != 1)
                {
                    m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                code = m_AccessHandle.WriteSigleData("TestResult_1", index, "xMax", (int)m_xMax);
                if (code != 1)
                {
                    m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                code = m_AccessHandle.WriteSigleData("TestResult_1", index, "yMax", m_yMax);
                if (code != 1)
                {
                    m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                code = m_AccessHandle.WriteFloatsData("TestResult_1", index, "TestResult", testDate);
                if (code != 1)
                {
                    m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                PointF[] pt;
                m_BenchHandle.m_DrawCurveHandle.GetSourcePointF(out pt);
                float[] PointF_x = new float[pt.Length];
                float[] PointF_y = new float[pt.Length];
                int len = pt.Length;
                for (int i = 0; i < len; i++)
                {
                    PointF_x[i] = pt[i].X;
                    PointF_y[i] = pt[i].Y;
                }

                code = m_AccessHandle.WriteFloatsData("TestResult_1", index, "CuverData_X", PointF_x);
                if (code != 1)
                {
                    m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }

                code = m_AccessHandle.WriteFloatsData("TestResult_1", index, "CuverData_Y", PointF_y);
                if (code != 1)
                {
                    m_AccessHandle.DeleteRow("TestResult_1", index);
                    return false;
                }
                code = m_AccessHandle.SaveDataToBuffer();
                code = m_AccessHandle.SaveDateToDataBase();
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        private int D_delaytime = 0;
        private bool CreatOfficeReport()
        {
            string FileName = m_BaseFilePath + @"\\Report\\Module.doc";

            if (File.Exists(FileName))
            {
                string newFileName = m_BaseFilePath + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".doc";
                File.Copy(FileName, newFileName);
                FileName = newFileName;
            }

            //创建Word文档
            Object oMissing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word._Application WordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word._Document WordDoc = WordApp.Documents.Open(FileName, ref oMissing, ref oMissing, ref oMissing);
            //打开office文档
            if (File.Exists(FileName))
            {
                //File.Open(FileName,FileMode.Open);
                System.Diagnostics.Process.Start(FileName);
            }

            //设置格式
            WordApp.Selection.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceAtLeast;//1倍行距
            WordApp.Selection.ParagraphFormat.SpaceBefore = float.Parse("0");
            WordApp.Selection.ParagraphFormat.SpaceBeforeAuto = 0;
            WordApp.Selection.ParagraphFormat.SpaceAfter = float.Parse("0");//段后间距
            WordApp.Selection.ParagraphFormat.SpaceAfterAuto = 0;
            WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
            try
            {
                for (int i = 0; i < D_delaytime; i++) ;
                WordDoc.Paragraphs.Last.Range.Font.Name = "Times New Roman";
                WordDoc.Paragraphs.Last.Range.Font.Size = 9;
                string strContent = m_OfficeInfo + "\n";
                WordDoc.Paragraphs.Last.Range.Text = strContent;
                for (int i = 0; i < D_delaytime; i++) ;


                WordDoc.Paragraphs.Last.Range.Font.Name = "宋体";
                WordDoc.Paragraphs.Last.Range.Font.Size = 11;
                //WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
                strContent = "压力试验报告/Hydrostatic Test Chart\n";
                //WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
                WordDoc.Paragraphs.Last.Range.Text = strContent;

                for (int i = 0; i < D_delaytime; i++) ;

                WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;//水平居中
                WordDoc.Paragraphs.Last.Range.Font.Name = "宋体";
                WordDoc.Paragraphs.Last.Range.Font.Size = 9;
                strContent = "产品信息/Product Information\n";
                WordDoc.Paragraphs.Last.Range.Text = strContent;

                for (int i = 0; i < D_delaytime; i++) ;

                //添加表格
                Microsoft.Office.Interop.Word.Table table1 = WordDoc.Tables.Add(WordDoc.Paragraphs.Last.Range, 2, 7, ref oMissing, ref oMissing);
                //设置表格样式
                table1.Range.Font.Name = "宋体";
                table1.Range.Font.Size = 8;
                table1.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleThickThinLargeGap;
                table1.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;

                table1.Cell(1, 1).Range.Text = "产品图号及名称/Draw No";
                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(1, 2).Range.Text = "产品编号/Serial No";
                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(1, 3).Range.Text = "试压标准/Test Standard";
                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(1, 4).Range.Text = "扭矩/Test Torque(N.m)";
                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(1, 5).Range.Text = "设备编号/Equipment No";
                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(1, 6).Range.Text = "单位名称/Company Name";
                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(1, 7).Range.Text = "合同号/Contract No";
                for (int i = 0; i < D_delaytime; i++) ;

                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(2, 1).Range.Text = TB_Name.Text;
                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(2, 2).Range.Text = TB_ProductID.Text;
                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(2, 3).Range.Text = TB_Module.Text;
                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(2, 4).Range.Text = TB_Standard.Text;
                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(2, 5).Range.Text = TB_Niuju.Text;
                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(2, 6).Range.Text = TB_EquipNum.Text;
                for (int i = 0; i < D_delaytime; i++) ;
                table1.Cell(2, 7).Range.Text = TB_CompanyName.Text;
                for (int i = 0; i < D_delaytime; i++) ;

                //插入文本，试验数据
                strContent = "\n";
                WordDoc.Paragraphs.Last.Range.Text = strContent;
                for (int i = 0; i < D_delaytime; i++) ;

                WordDoc.Paragraphs.Last.Range.Font.Name = "宋体";
                WordDoc.Paragraphs.Last.Range.Font.Size = 9;
                //WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
                strContent = "试验数据/Test Parameter\n";
                WordDoc.Paragraphs.Last.Range.Text = strContent;
                //for (int i = 0; i < D_delaytime; i++) ;

                for (int i = 0; i < D_delaytime; i++) ;
                //第二个表格
                int rows = dataGridView_DateList.Rows.Count;
                Microsoft.Office.Interop.Word.Table table2 = WordDoc.Tables.Add(WordDoc.Paragraphs.Last.Range, rows - 1, 5, ref oMissing, ref oMissing);
                //表格样式
                table2.Range.Font.Size = 8;
                table2.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleThickThinLargeGap;
                table2.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;

                for (int i = 1; i < rows; i++)
                {
                    for (int j = 1; j <= 5; j++)
                    {
                        table2.Cell(i, j).Range.Text = dataGridView_DateList[j - 1, i - 1].Value.ToString();
                        //for (int t = 0; t < D_delaytime; t++) ;
                    }
                }

                for (int i = 0; i < D_delaytime; i++) ;

                strContent = "\n";
                WordDoc.Paragraphs.Last.Range.Text = strContent;

                //插入图片
                FileName = m_BaseFilePath + "\\Presure.png";//图片所在路径
                object LinkToFile = false;
                object SaveWithDocument = true;
                object Anchor = WordDoc.Paragraphs.Last.Range;
                //object Anchor = WordDoc.Application.Selection.Range;
                //WordDoc.Paragraphs.Last.Range = WordDoc.Application.Selection.Range;

                WordDoc.Application.ActiveDocument.InlineShapes.AddPicture(FileName, ref LinkToFile, ref SaveWithDocument, ref Anchor);
                Microsoft.Office.Interop.Word.Shape s = WordDoc.Application.ActiveDocument.InlineShapes[1].ConvertToShape();
                s.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapTopBottom;
                object format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault;

                object oEndOfDoc = "\\endofdoc";
                Microsoft.Office.Interop.Word.Paragraph return_pragraph;
                object myrange2 = WordDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                return_pragraph = WordDoc.Content.Paragraphs.Add(ref myrange2);
                return_pragraph.Range.InsertParagraphAfter(); //插入一个空白行  

                for (int i = 0; i < D_delaytime; i++) ;
                //插入试验结果
                WordDoc.Paragraphs.Last.Range.Font.Name = "宋体";
                WordDoc.Paragraphs.Last.Range.Font.Size = 11;
                //WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
                strContent = "试验结果：在保压周期内产品无可见渗漏\n";
                WordDoc.Paragraphs.Last.Range.Text = strContent;

                for (int i = 0; i < D_delaytime; i++) ;
                //插入试验结果
                WordDoc.Paragraphs.Last.Range.Font.Name = "Times New Roman";
                WordDoc.Paragraphs.Last.Range.Font.Size = 9;
                //WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
                strContent = "Test Result：no visibale leakage during each holding period\n";
                WordDoc.Paragraphs.Last.Range.Text = strContent;

                for (int i = 0; i < D_delaytime; i++) ;
                //第三个表格
                Microsoft.Office.Interop.Word.Table table3 = WordDoc.Tables.Add(WordDoc.Paragraphs.Last.Range, 2, 4, ref oMissing, ref oMissing);
                //表格样式
                table3.Range.Font.Name = "宋体";
                table3.Range.Font.Size = 8;
                table3.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleThickThinLargeGap;
                table3.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;

                for (int i = 0; i < D_delaytime; i++) ;
                table3.Cell(1, 1).Range.Text = "试验人员/Tested by";
                for (int i = 0; i < D_delaytime; i++) ;
                table3.Cell(1, 2).Range.Text = "检验人员/Inspected by";
                for (int i = 0; i < D_delaytime; i++) ;
                table3.Cell(1, 3).Range.Text = "审核/Review by";
                for (int i = 0; i < D_delaytime; i++) ;
                table3.Cell(1, 4).Range.Text = "日期/Date";
                for (int i = 0; i < D_delaytime; i++) ;
                table3.Cell(2, 1).Range.Text = TB_Operator.Text;
                for (int i = 0; i < D_delaytime; i++) ;
                table3.Cell(2, 2).Range.Text = TB_Tester.Text;
                for (int i = 0; i < D_delaytime; i++) ;
                table3.Cell(2, 3).Range.Text = TB_Conformation.Text;
                for (int i = 0; i < D_delaytime; i++) ;
                table3.Cell(2, 4).Range.Text = TB_Date.Text;
                for (int i = 0; i < D_delaytime; i++) ;

                string SubName = DateTime.Now.ToString("yyyyMMddHHmmss") + ".doc";
                //format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault;
                object filename = m_BaseFilePath + "\\Report\\" + SubName;
                for (int i = 0; i < D_delaytime; i++) ;

                WordDoc.SaveAs(ref filename, ref format, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                //关闭wordDoc 文档对象
                //WordDoc.Close(ref oMissing, ref oMissing, ref oMissing);
                //关闭wordApp 组件对象
                //WordApp.Quit(ref oMissing, ref oMissing, ref oMissing);
                FileName = filename.ToString();
            }
            catch (Exception)
            {
                //关闭wordDoc 文档对象
                WordDoc.Close(ref oMissing, ref oMissing, ref oMissing);
                //关闭wordApp 组件对象
                WordApp.Quit(ref oMissing, ref oMissing, ref oMissing);

                return false;
            }
            return true;
        }

        private bool ReadCurveDataFromDataBase()
        {
            int code = 0;
            int index = 0;
            code = m_AccessHandle.GetIndexBaseKeyWord("TestResult_1", m_RecordId, ref index);
            if (code != 1)
            {
                MessageBox.Show("数据不存在", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            float[] X;
            float[] Y;

            code = m_AccessHandle.ReadFloatsData("TestResult_1", index, "CuverData_X", out X);
            if (code != 1)
            {
                return false;
            }

            code = m_AccessHandle.ReadFloatsData("TestResult_1", index, "CuverData_Y", out Y);
            if (code != 1)
            {
                return false;
            }

            PointF[] p = new PointF[X.Length];
            m_HistoryPointF = new PointF[X.Length];
            int len = X.Length;
            for (int i = 0; i < len; i++)
            {
                p[i].X = X[i];
                p[i].Y = Y[i];
                m_HistoryPointF[i].X = X[i];
                m_HistoryPointF[i].Y = Y[i];
            }

            DateTime startTime = default(DateTime);
            DateTime endTime = default(DateTime);

            object Obj = default(object);
            code = m_AccessHandle.ReadSigleData("TestResult_1", index, "StartTime", ref Obj);
            if (code != 1)
            {
                return false;
            }
            startTime = Convert.ToDateTime(Obj);
            m_HistoryStartTime = startTime;

            code = m_AccessHandle.ReadSigleData("TestResult_1", index, "EndTime", ref Obj);
            if (code != 1)
            {
                return false;
            }
            endTime = Convert.ToDateTime(Obj);
            m_HistoryEndTime = endTime;

            return true;
        }

        private void button_Return_Click(object sender, EventArgs e)
        {
            this.Close();
            this.Dispose();
        }

        private bool m_isCreatReport = false;

        private void button_CreatReport_Click(object sender, EventArgs e)
        {
            bool ret = false;
            SavePicCh();
            ReadTextBoxData();
            ret = CreatOfficeReport();
            if (!ret)
            {
                MessageBox.Show("创建报表失败", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //this.Cursor = System.Windows.Forms.Cursors.Arrow;
                return;
            }

            if (!m_isOpenDataBase)
            {
                if (!m_isCreatReport)
                {
                    ret = SaveDataToDataBase();
                    if (!ret)
                    {
                        MessageBox.Show("创建报表失败,请检查数据信息", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        this.Cursor = System.Windows.Forms.Cursors.Arrow;
                        return;
                    }
                    m_isCreatReport = true;
                }
            }
        }

        private void button_Save_Path_Click(object sender, EventArgs e)
        {
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            m_SelectHistoryData_No2Handle = new SelectHistoryData_No2(this, m_BenchHandle);
            m_SelectHistoryData_No2Handle.ControlBox = false;
            m_SelectHistoryData_No2Handle.ShowDialog();
            m_DisStartTime = default(DateTime);
            m_DisEndTime = default(DateTime);

            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            if (m_isNoSelect)
            {
                this.Cursor = System.Windows.Forms.Cursors.Arrow;
                return;
            }

            bool ret = ReadInfoDataFromDataBase();
            if (!ret)
            {
                MessageBox.Show("读取设备信息失败", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Cursor = System.Windows.Forms.Cursors.Arrow;
                return;
            }

            ret = ReadTestDataFromDataBase();
            if (!ret)
            {
                MessageBox.Show("读取试验数据失败", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Cursor = System.Windows.Forms.Cursors.Arrow;
                return;
            }

            ret = ReadCurveDataFromDataBase();
            if (!ret)
            {
                MessageBox.Show("读取曲线信息失败", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Cursor = System.Windows.Forms.Cursors.Arrow;
                return;
            }
            m_isOpenDataBase = true;
            DrawHisCurveCh();
            this.Cursor = System.Windows.Forms.Cursors.Arrow;
        }

        /// <summary>
        /// 读取测试信息，如测试人员--公共
        /// </summary>        
        private void ReadTestInfo()
        {
            string strSecfileName = Path.GetFileNameWithoutExtension(strFilePath);
            I_ProductName = ContentValue(strSecfileName, "ProductName");
            TB_Name.Text = I_ProductName;
            TB_Name_e.Text = I_ProductName;

            I_ProductNo = ContentValue(strSecfileName, "ProductNo");
            TB_ProductID.Text = I_ProductNo;
            TB_ProductID_e.Text = I_ProductNo;

            I_ProductModel = ContentValue(strSecfileName, "ProductModel");
            TB_Module.Text = I_ProductModel;
            TB_Module_e.Text = I_ProductModel;

            I_TestStandard = ContentValue(strSecfileName, "TestStandard");
            TB_Standard.Text = I_TestStandard;
            TB_Standard_e.Text = I_TestStandard;

            I_TestNiuju = ContentValue(strSecfileName, "TestNiuju");
            TB_Niuju.Text = I_TestNiuju;
            TB_Niuju_e.Text = I_TestNiuju;

            I_EquipmentNo = ContentValue(strSecfileName, "EquipmentNo");
            TB_EquipNum.Text = I_EquipmentNo;
            TB_EquipNum_e.Text = I_EquipmentNo;

            I_CompanyName = ContentValue(strSecfileName, "CompanyName");
            TB_CompanyName.Text = I_CompanyName;
            TB_CompanyName_e.Text = I_CompanyName;

            I_Operator = ContentValue(strSecfileName, "Operator");
            TB_Operator.Text = I_Operator;
            TB_Operator_e.Text = I_Operator;

            I_Tester = ContentValue(strSecfileName, "Tester");
            TB_Tester.Text = I_Tester;
            TB_Tester_e.Text = I_Tester;

            I_Confirmation = ContentValue(strSecfileName, "Confirmation");
            TB_Conformation.Text = I_Confirmation;
            TB_Conformation_e.Text = I_Confirmation;

            I_Date = ContentValue(strSecfileName, "Date");
            TB_Date.Text = I_Date;
            TB_Date_e.Text = I_Date;
        }
        private string ContentValue(string Section, string key)
        {
            StringBuilder temp = new StringBuilder(1024);
            GetPrivateProfileString(Section, key, "", temp, 1024, strFilePath);
            return temp.ToString();
        }

        /// <summary>
        /// 读取试验数据
        /// </summary>
        private void ReadTestDate()
        {
            dataGridView_DateList.ColumnCount = 5;
            dataGridView_DateList_e.ColumnCount = 5;
            int len = m_BenchHandle.m_TestResultLists.Count;
            for (int i = 0; i < len + 1; i++)    //
            {
                dataGridView_DateList.Rows.Add();
                dataGridView_DateList_e.Rows.Add();
            }
            dataGridView_DateList.Rows[0].Cells[0].Value = "级数/段数/Stage No";
            dataGridView_DateList.Rows[0].Cells[1].Value = "初始压力/Start Pressure（MPa）";
            dataGridView_DateList.Rows[0].Cells[2].Value = "终止压力/Final Pressure（MPa）";
            dataGridView_DateList.Rows[0].Cells[3].Value = "保压时间/Hold Period（Min）";
            dataGridView_DateList.Rows[0].Cells[4].Value = "保压压降/Pressure Reduction（MPa）";

            dataGridView_DateList_e.Rows[0].Cells[0].Value = "Stage No";
            dataGridView_DateList_e.Rows[0].Cells[1].Value = "Start Pressure(MPa)";
            dataGridView_DateList_e.Rows[0].Cells[2].Value = "Final Pressure(MPa)";
            dataGridView_DateList_e.Rows[0].Cells[3].Value = "Hold Period(Min)";
            dataGridView_DateList_e.Rows[0].Cells[4].Value = "Pressure Reduction(MPa)";

            for (int i = 1; i < len + 1; i++)
            {
                //中文报告
                dataGridView_DateList.Rows[i].Cells[0].Value = m_BenchHandle.m_TestResultLists[i - 1].m_No; ;
                dataGridView_DateList.Rows[i].Cells[1].Value = Double.Parse(m_BenchHandle.m_TestResultLists[i - 1].m_InitPressure.ToString("0.00"));
                dataGridView_DateList.Rows[i].Cells[2].Value = Double.Parse(m_BenchHandle.m_TestResultLists[i - 1].m_EndPressure.ToString("0.00"));
                dataGridView_DateList.Rows[i].Cells[3].Value = Double.Parse(m_BenchHandle.m_TestResultLists[i - 1].m_KeepTime.ToString("0.00"));
                dataGridView_DateList.Rows[i].Cells[4].Value = Double.Parse(m_BenchHandle.m_TestResultLists[i - 1].m_DropPressure.ToString("0.00"));
                //英文报告
                dataGridView_DateList_e.Rows[i].Cells[0].Value = m_BenchHandle.m_TestResultLists[i - 1].m_No;
                dataGridView_DateList_e.Rows[i].Cells[1].Value = Double.Parse(m_BenchHandle.m_TestResultLists[i - 1].m_InitPressure.ToString("0.00"));
                dataGridView_DateList_e.Rows[i].Cells[2].Value = Double.Parse(m_BenchHandle.m_TestResultLists[i - 1].m_EndPressure.ToString("0.00"));
                dataGridView_DateList_e.Rows[i].Cells[3].Value = Double.Parse(m_BenchHandle.m_TestResultLists[i - 1].m_KeepTime.ToString("0.00"));
                dataGridView_DateList_e.Rows[i].Cells[4].Value = Double.Parse(m_BenchHandle.m_TestResultLists[i - 1].m_DropPressure.ToString("0.00"));
            }
        }

        private bool ReadInfoDataFromDataBase()
        {
            int code = 0;
            int index = 0;
            code = m_AccessHandle.GetIndexBaseKeyWord("TestResult_1", m_RecordId, ref index);
            if (code != 1)
            {
                MessageBox.Show("数据不存在", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            object Obj = default(object);
            code = m_AccessHandle.ReadSigleData("TestResult_1", index, "ProductName", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_ProductName = Obj.ToString();

            code = m_AccessHandle.ReadSigleData("TestResult_1", index, "ProductNo", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_ProductNo = Obj.ToString();

            code = m_AccessHandle.ReadSigleData("TestResult_1", index, "ProductModel", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_ProductModel = Obj.ToString();

            code = m_AccessHandle.ReadSigleData("TestResult_1", index, "TestStandard", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_TestStandard = Obj.ToString();

            code = m_AccessHandle.ReadSigleData("TestResult_1", index, "TestNiuju", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_TestNiuju = Obj.ToString();

            code = m_AccessHandle.ReadSigleData("TestResult_1", index, "EquipmentNo", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_EquipmentNo = Obj.ToString();

            code = m_AccessHandle.ReadSigleData("TestResult_1", index, "CompanyName", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_CompanyName = Obj.ToString();

            code = m_AccessHandle.ReadSigleData("TestResult_1", index, "Operator", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_Operator = Obj.ToString();

            code = m_AccessHandle.ReadSigleData("TestResult_1", index, "Tester", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_Tester = Obj.ToString();

            code = m_AccessHandle.ReadSigleData("TestResult_1", index, "Confirmation", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_Confirmation = Obj.ToString();

            code = m_AccessHandle.ReadSigleData("TestResult_1", index, "TestDate", ref Obj);
            if (code != 1)
            {
                return false;
            }
            I_Date = Obj.ToString();

            TB_Name.Text = I_ProductName;
            TB_ProductID.Text = I_ProductNo;
            TB_Module.Text = I_ProductModel;
            TB_Standard.Text = I_TestStandard;
            TB_Niuju.Text = I_TestNiuju;
            TB_EquipNum.Text = I_EquipmentNo;
            TB_CompanyName.Text = I_CompanyName;
            TB_Operator.Text = I_Operator;
            TB_Tester.Text = I_Tester;
            TB_Conformation.Text = I_Confirmation;

            TB_Name_e.Text = I_ProductName;
            TB_ProductID_e.Text = I_ProductNo;
            TB_Module_e.Text = I_ProductModel;
            TB_Standard_e.Text = I_TestStandard;
            TB_Niuju_e.Text = I_TestNiuju;
            TB_EquipNum_e.Text = I_EquipmentNo;
            TB_CompanyName_e.Text = I_CompanyName;
            TB_Operator_e.Text = I_Operator;
            TB_Tester_e.Text = I_Tester;
            TB_Conformation_e.Text = I_Confirmation;

            TB_Result.Text = "在保压周期内无可见渗漏";
            TB_Result_e.Text = "no visibale leakage during each holding period";
            return true;
        }

        private bool ReadTestDataFromDataBase()
        {
            int code = 0;
            int index = 0;
            code = m_AccessHandle.GetIndexBaseKeyWord("TestResult_1", m_RecordId, ref index);
            if (code != 1)
            {
                MessageBox.Show("数据不存在", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            int TestDataNum = 0;
            object Obj = default(object);
            code = m_AccessHandle.ReadSigleData("TestResult_1", index, "ResultsNum", ref Obj);
            if (code != 1)
            {
                return false;
            }
            TestDataNum = Convert.ToInt32(Obj);

            code = m_AccessHandle.ReadSigleData("TestResult_1", index, "xMax", ref Obj);
            if (code != 1)
            {
                return false;
            }
            m_xMax = Convert.ToInt32(Obj);

            code = m_AccessHandle.ReadSigleData("TestResult_1", index, "yMax", ref Obj);
            if (code != 1)
            {
                return false;
            }
            m_yMax = Convert.ToUInt32(Obj);
            textBox_yMax.Text = m_yMax.ToString();
            textBox_yMax_e.Text = m_yMax.ToString();

            float[] TestData;
            code = m_AccessHandle.ReadFloatsData("TestResult_1", index, "TestResult", out TestData);
            if (code != 1)
            {
                return false;
            }

            dataGridView_DateList.Rows.Clear();
            dataGridView_DateList_e.Rows.Clear();
            dataGridView_DateList.ColumnCount = 5;
            dataGridView_DateList_e.ColumnCount = 5;
            for (int i = 0; i < TestDataNum + 1; i++)    //
            {
                dataGridView_DateList.Rows.Add();
                dataGridView_DateList_e.Rows.Add();
            }
            dataGridView_DateList.Rows[0].Cells[0].Value = "级数/段数/Stage No";
            dataGridView_DateList.Rows[0].Cells[1].Value = "初始压力/Start Pressure（MPa）";
            dataGridView_DateList.Rows[0].Cells[2].Value = "终止压力/Final Pressure（MPa）";
            dataGridView_DateList.Rows[0].Cells[3].Value = "保压时间/Hold Period（Min）";
            dataGridView_DateList.Rows[0].Cells[4].Value = "保压压降/Pressure Reduction（MPa）";

            dataGridView_DateList_e.Rows[0].Cells[0].Value = "Stage No";
            dataGridView_DateList_e.Rows[0].Cells[1].Value = "Start Pressure(MPa)";
            dataGridView_DateList_e.Rows[0].Cells[2].Value = "Final Pressure(MPa)";
            dataGridView_DateList_e.Rows[0].Cells[3].Value = "Hold Period(Min)";
            dataGridView_DateList_e.Rows[0].Cells[4].Value = "Pressure Reduction(MPa)";

            for (int i = 1; i < TestDataNum + 1; i++)
            {
                for (int j = 0; j < 5; j++)
                {
                    if (j == 0)
                    {
                        string No = ((int)TestData[(i - 1) * 5 + j]).ToString();
                        char[] a = No.ToCharArray();
                        if (a.Length == 2)
                        {
                            dataGridView_DateList.Rows[i].Cells[j].Value = a[0] + "--" + a[1];
                            dataGridView_DateList_e.Rows[i].Cells[j].Value = a[0] + "--" + a[1];
                        }
                        if (a.Length == 3)
                        {
                            dataGridView_DateList.Rows[i].Cells[j].Value = a[0] + a[1] + "--" + a[2];
                            dataGridView_DateList_e.Rows[i].Cells[j].Value = a[0] + a[1] + "--" + a[2];

                        }
                        if (a.Length == 4)
                        {
                            dataGridView_DateList.Rows[i].Cells[j].Value = a[0] + a[1] + a[2] + "--" + a[3];
                            dataGridView_DateList_e.Rows[i].Cells[j].Value = a[0] + a[1] + a[2] + "--" + a[3];
                        }
                        if (a.Length == 5)
                        {
                            dataGridView_DateList.Rows[i].Cells[j].Value = a[0] + a[1] + a[2] + a[3] + "--" + a[4];
                            dataGridView_DateList_e.Rows[i].Cells[j].Value = a[0] + a[1] + a[2] + a[3] + "--" + a[4];
                        }
                    }
                    else
                    {
                        dataGridView_DateList.Rows[i].Cells[j].Value = TestData[(i - 1) * 5 + j];
                        dataGridView_DateList_e.Rows[i].Cells[j].Value = TestData[(i - 1) * 5 + j];
                    }
                }
            }

            return true;
        }

        private void TB_Name_TextChanged(object sender, EventArgs e)
        {
            TB_Name_e.Text = TB_Name.Text;
        }

        private void TB_Name_e_TextChanged(object sender, EventArgs e)
        {
            TB_Name.Text = TB_Name_e.Text;
        }

        private void TB_ProductID_TextChanged(object sender, EventArgs e)
        {
            TB_ProductID_e.Text = TB_ProductID.Text;
        }

        private void TB_ProductID_e_TextChanged(object sender, EventArgs e)
        {
            TB_ProductID.Text = TB_ProductID_e.Text;
        }

        private void TB_Module_TextChanged(object sender, EventArgs e)
        {
            TB_Module_e.Text = TB_Module.Text;
        }

        private void TB_Module_e_TextChanged(object sender, EventArgs e)
        {
            TB_Module.Text = TB_Module_e.Text;
        }

        private void TB_Standard_TextChanged(object sender, EventArgs e)
        {
            TB_Standard_e.Text = TB_Standard.Text;
        }

        private void TB_Standard_e_TextChanged(object sender, EventArgs e)
        {
            TB_Standard.Text = TB_Standard_e.Text;
        }

        private void TB_Niuju_TextChanged(object sender, EventArgs e)
        {
            TB_Niuju_e.Text = TB_Niuju.Text;
        }

        private void TB_Niuju_e_TextChanged(object sender, EventArgs e)
        {
            TB_Niuju.Text = TB_Niuju_e.Text;
        }

        private void TB_EquipNum_TextChanged(object sender, EventArgs e)
        {
            TB_EquipNum_e.Text = TB_EquipNum.Text;
        }

        private void TB_EquipNum_e_TextChanged(object sender, EventArgs e)
        {
            TB_EquipNum.Text = TB_EquipNum_e.Text;
        }

        private void TB_CompanyName_TextChanged(object sender, EventArgs e)
        {
            TB_CompanyName_e.Text = TB_CompanyName.Text;
        }

        private void TB_CompanyName_e_TextChanged(object sender, EventArgs e)
        {
            TB_CompanyName.Text = TB_CompanyName_e.Text;
        }

        private void TB_Operator_TextChanged(object sender, EventArgs e)
        {
            TB_Operator_e.Text = TB_Operator.Text;
        }

        private void TB_Operator_e_TextChanged(object sender, EventArgs e)
        {
            TB_Operator.Text = TB_Operator_e.Text;
        }

        private void TB_Tester_TextChanged(object sender, EventArgs e)
        {
            TB_Tester_e.Text = TB_Tester.Text;
        }

        private void TB_Tester_e_TextChanged(object sender, EventArgs e)
        {
            TB_Tester.Text = TB_Tester_e.Text;
        }

        private void TB_Conformation_TextChanged(object sender, EventArgs e)
        {
            TB_Conformation_e.Text = TB_Conformation.Text;
        }

        private void TB_Conformation_e_TextChanged(object sender, EventArgs e)
        {
            TB_Conformation.Text = TB_Conformation_e.Text;
        }

        //////////////////////////////////////////////////////////////////////////
        //////////////////////////////////////////////////////////////////////////

        private void DrawCurveEn()
        {
            bool isTesting = m_BenchHandle.m_isTesting;
            if (isTesting)
            {
                return;
            }
            m_BenchHandle.m_DrawCurveHandle.SetModule(0);
            PointF[] pt;
            m_BenchHandle.m_DrawCurveHandle.GetSourcePointF(out pt);
            if (pt.Length <= 2)
            {
                return;
            }
            int len = pt.Length;
            if (pt[len - 1].X == 0)
            {
                m_xMax = pt[len - 2].X;
            }
            else
            {
                m_xMax = pt[len - 2].X;
            }
            m_BenchHandle.m_DrawCurveHandle.DrawCurve(panel_Result_e, Color.Red, m_xMax, m_yMax);

            DateTime StartTime = m_BenchHandle.m_StartTime;
            DateTime EndTime = default(DateTime);
            m_BenchHandle.ReadEndTime(ref EndTime);
            TimeSpan tp = EndTime - StartTime;
            int second = (int)tp.TotalSeconds;
            label_X1_e.Text = StartTime.ToString("HH:mm:ss");
            textBox_StartTime_e.Text = label_X1_e.Text;

            int H = 0;
            int M = 0;
            int S = 0;
            GetlableTime(second / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X2_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 2 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X3_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 3 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X4_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 4 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X5_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 5 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X6_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 6 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X7_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            label_X7_e.Text = EndTime.ToString("HH:mm:ss");
            textBox_EndTime_e.Text = label_X7_e.Text;

            UInt32 t1 = m_yMax * 1 / 5;
            label45.Text = t1.ToString() + "MPa";

            UInt32 t2 = m_yMax * 2 / 5;
            label44.Text = t2.ToString() + "MPa";

            UInt32 t3 = m_yMax * 3 / 5;
            label43.Text = t3.ToString() + "MPa";

            UInt32 t4 = m_yMax * 4 / 5;
            label41.Text = t4.ToString() + "MPa";
        }

        private void DrawHisCurveEn()
        {
            m_BenchHandle.m_DrawCurveHandle.SetModule(2);
            m_BenchHandle.m_DrawCurveHandle.ClearALLPointF();
            m_BenchHandle.m_DrawCurveHandle.SaveAllPointF(m_HistoryPointF);
            float xMax = 0;
            float yMax = m_yMax;
            int len = m_HistoryPointF.Length;
            if (len < 2)
            {
                return;
            }
            if (m_HistoryPointF[len - 1].X == 0)
            {
                xMax = m_HistoryPointF[len - 2].X;
            }
            else
            {
                xMax = m_HistoryPointF[len - 2].X;
            }
            m_BenchHandle.m_DrawCurveHandle.DrawCurve(panel_Result_e, Color.Red, xMax, yMax);

            DateTime startTime = default(DateTime);
            DateTime endTime = default(DateTime);
            startTime = m_HistoryStartTime;
            endTime = m_HistoryEndTime;

            TimeSpan tp = endTime - startTime;
            int second = (int)tp.TotalSeconds;
            label_X1_e.Text = startTime.ToString("HH:mm:ss");
            textBox_StartTime_e.Text = label_X1_e.Text;

            int H = 0;
            int M = 0;
            int S = 0;
            GetlableTime(second / 6, ref H, ref M, ref S);
            S += startTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += startTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += startTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X2_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 2 / 6, ref H, ref M, ref S);
            S += startTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += startTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += startTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X3_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 3 / 6, ref H, ref M, ref S);
            S += startTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += startTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += startTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X4_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 4 / 6, ref H, ref M, ref S);
            S += startTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += startTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += startTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X5_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 5 / 6, ref H, ref M, ref S);
            S += startTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += startTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += startTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X6_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 6 / 6, ref H, ref M, ref S);
            S += startTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += startTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += startTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X7_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            label_X7_e.Text = endTime.ToString("HH:mm:ss");
            textBox_EndTime_e.Text = label_X7_e.Text;

            UInt32 t1 = m_yMax * 1 / 5;
            label45.Text = t1.ToString() + "MPa";

            UInt32 t2 = m_yMax * 2 / 5;
            label44.Text = t2.ToString() + "MPa";

            UInt32 t3 = m_yMax * 3 / 5;
            label43.Text = t3.ToString() + "MPa";

            UInt32 t4 = m_yMax * 4 / 5;
            label41.Text = t4.ToString() + "MPa";
        }

        /// <summary>
        /// 将英文产品信息存入到本地缓存
        /// </summary>
        private void ReadTextBoxDataEn()
        {
            I_ProductName = TB_Name_e.Text;
            I_ProductNo = TB_ProductID_e.Text;
            I_ProductModel = TB_Module_e.Text;
            I_TestStandard = TB_Standard_e.Text;
            I_TestNiuju = TB_Niuju_e.Text;
            I_EquipmentNo = TB_EquipNum_e.Text;
            I_CompanyName = TB_CompanyName_e.Text;
            I_Operator = TB_Operator_e.Text;
            I_Tester = TB_Tester_e.Text;
            I_Confirmation = TB_Conformation_e.Text;

            WritePrivateProfileString(strSec, "ProductName", I_ProductName, strFilePath);
            WritePrivateProfileString(strSec, "ProductNo", I_ProductNo, strFilePath);
            WritePrivateProfileString(strSec, "ProductModel", I_ProductModel, strFilePath);
            WritePrivateProfileString(strSec, "TestStandard", I_TestStandard, strFilePath);
            WritePrivateProfileString(strSec, "TestNiuju", I_TestNiuju, strFilePath);
            WritePrivateProfileString(strSec, "EquipmentNo", I_EquipmentNo, strFilePath);
            WritePrivateProfileString(strSec, "CompanyName", I_CompanyName, strFilePath);
            WritePrivateProfileString(strSec, "Operator", I_Operator, strFilePath);
            WritePrivateProfileString(strSec, "Tester", I_Tester, strFilePath);
            WritePrivateProfileString(strSec, "Confirmation", I_Confirmation, strFilePath);
            WritePrivateProfileString(strSec, "Date", I_Date, strFilePath);
        }

        private bool CreatOfficeReportEn()
        {
            string FileName = m_BaseFilePath + @"\Report\Module.doc";
            //创建Word文档
            Object oMissing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word._Application WordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word._Document WordDoc = WordApp.Documents.Open(FileName, ref oMissing, ref oMissing, ref oMissing);
            //打开office文档
            if (File.Exists(FileName))
            {
                //File.Open(FileName,FileMode.Open);
                System.Diagnostics.Process.Start(FileName);
            }

            //设置格式
            WordApp.Selection.ParagraphFormat.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceAtLeast;//1.5倍行距
            WordApp.Selection.ParagraphFormat.SpaceBefore = float.Parse("0");
            WordApp.Selection.ParagraphFormat.SpaceBeforeAuto = 0;
            WordApp.Selection.ParagraphFormat.SpaceAfter = float.Parse("0");//段后间距
            WordApp.Selection.ParagraphFormat.SpaceAfterAuto = 0;
            WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;//水平居中

            try
            {
                WordDoc.Paragraphs.Last.Range.Font.Name = "Times New Roman";
                WordDoc.Paragraphs.Last.Range.Font.Size = 9;
                //WordDoc.Paragraphs.Last.Range           
                string strContent = m_OfficeInfo + "\n";
                WordDoc.Paragraphs.Last.Range.Text = strContent;
                for (int i = 0; i < D_delaytime; i++) ;

                WordDoc.Paragraphs.Last.Range.Font.Name = "Times New Roman";
                WordDoc.Paragraphs.Last.Range.Font.Size = 11;
                //WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
                strContent = "Hydrostatic Test Chart\n";
                WordDoc.Paragraphs.Last.Range.Text = strContent;

                WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;//水平居中
                WordDoc.Paragraphs.Last.Range.Font.Name = "Times New Roman";
                WordDoc.Paragraphs.Last.Range.Font.Size = 10;
                //WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;//水平居中
                strContent = "Product Information\n";
                WordDoc.Paragraphs.Last.Range.Text = strContent;

                //添加表格
                Microsoft.Office.Interop.Word.Table table = WordDoc.Tables.Add(WordDoc.Paragraphs.Last.Range, 2, 7, ref oMissing, ref oMissing);
                //设置表格样式
                table.Range.Font.Name = "Times New Roman";
                table.Range.Font.Size = 8;
                table.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleThickThinLargeGap;
                table.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;

                table.Cell(1, 1).Range.Text = "Draw No";
                table.Cell(1, 2).Range.Text = "Serial No";
                table.Cell(1, 3).Range.Text = "Test Standard";
                table.Cell(1, 4).Range.Text = "Test Torque(N.M)";
                table.Cell(1, 5).Range.Text = "Equipment No";
                table.Cell(1, 6).Range.Text = "Company Name";
                table.Cell(1, 7).Range.Text = "Contract No";

                table.Cell(2, 1).Range.Text = TB_Name.Text;
                table.Cell(2, 2).Range.Text = TB_ProductID.Text;
                table.Cell(2, 3).Range.Text = TB_Module.Text;
                table.Cell(2, 4).Range.Text = TB_Standard.Text;
                table.Cell(2, 5).Range.Text = TB_Niuju.Text;
                table.Cell(2, 6).Range.Text = TB_EquipNum.Text;
                table.Cell(2, 7).Range.Text = TB_CompanyName.Text;

                //插入文本，试验数据
                strContent = "\n";
                WordDoc.Paragraphs.Last.Range.Text = strContent;

                WordDoc.Paragraphs.Last.Range.Font.Name = "Times New Roman";
                WordDoc.Paragraphs.Last.Range.Font.Size = 9;
                //WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
                strContent = "Test Parameter\n";
                WordDoc.Paragraphs.Last.Range.Text = strContent;

                //第二个表格
                int rows = dataGridView_DateList.Rows.Count;
                table = WordDoc.Tables.Add(WordDoc.Paragraphs.Last.Range, rows - 1, 5, ref oMissing, ref oMissing);
                //表格样式
                table.Range.Font.Size = 8;
                table.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleThickThinLargeGap;
                table.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;

                for (int i = 1; i < rows; i++)
                {
                    for (int j = 1; j <= 5; j++)
                    {
                        table.Cell(i, j).Range.Text = dataGridView_DateList[j - 1, i - 1].Value.ToString();
                        for (int t = 0; t < D_delaytime; t++) ;
                    }
                }
                table.Cell(1, 1).Range.Text = "Stage No";
                table.Cell(1, 2).Range.Text = "Start Pressure(MPa)";
                table.Cell(1, 3).Range.Text = "Final Pressure(MPa)";
                table.Cell(1, 4).Range.Text = "Hold Period(Min)";
                table.Cell(1, 5).Range.Text = "Pressure Reduction(MPa)";
                strContent = "\n";
                WordDoc.Paragraphs.Last.Range.Text = strContent;

                //插入图片
                FileName = m_BaseFilePath + "\\Presure.png";//图片所在路径
                object LinkToFile = false;
                object SaveWithDocument = true;
                object Anchor = WordDoc.Paragraphs.Last.Range;
                //object Anchor = WordDoc.Application.Selection.Range;
                //WordDoc.Paragraphs.Last.Range = WordDoc.Application.Selection.Range;

                WordDoc.Application.ActiveDocument.InlineShapes.AddPicture(FileName, ref LinkToFile, ref SaveWithDocument, ref Anchor);
                Microsoft.Office.Interop.Word.Shape s = WordDoc.Application.ActiveDocument.InlineShapes[1].ConvertToShape();
                s.WrapFormat.Type = Microsoft.Office.Interop.Word.WdWrapType.wdWrapTopBottom;
                object format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault;

                object oEndOfDoc = "\\endofdoc";
                Microsoft.Office.Interop.Word.Paragraph return_pragraph;
                object myrange2 = WordDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                return_pragraph = WordDoc.Content.Paragraphs.Add(ref myrange2);
                return_pragraph.Range.InsertParagraphAfter(); //插入一个空白行  
                //插入试验结果
                WordDoc.Paragraphs.Last.Range.Font.Name = "Times New Roman";
                WordDoc.Paragraphs.Last.Range.Font.Size = 9;
                //WordApp.Selection.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;//水平居中
                strContent = "Test Result：no visibale leakage during each holding period\n";
                WordDoc.Paragraphs.Last.Range.Text = strContent;

                //第三个表格
                table = WordDoc.Tables.Add(WordDoc.Paragraphs.Last.Range, 2, 4, ref oMissing, ref oMissing);
                //表格样式
                table.Range.Font.Size = 8;
                table.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleThickThinLargeGap;
                table.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;

                table.Cell(1, 1).Range.Text = "Tested By";
                table.Cell(1, 2).Range.Text = "Inspected by";
                table.Cell(1, 3).Range.Text = "Review By";
                table.Cell(1, 4).Range.Text = "Date";
                table.Cell(2, 1).Range.Text = TB_Operator.Text;
                table.Cell(2, 2).Range.Text = TB_Tester.Text;
                table.Cell(2, 3).Range.Text = TB_Conformation.Text;
                table.Cell(2, 4).Range.Text = TB_Date.Text;

                string SubName = "EReport" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".doc";
                //format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault;
                object filename = m_BaseFilePath + "\\Report\\" + SubName;
                for (int i = 0; i < D_delaytime; i++) ;

                WordDoc.SaveAs(ref filename, ref format, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                //关闭wordDoc 文档对象
                //WordDoc.Close(ref oMissing, ref oMissing, ref oMissing);
                //关闭wordApp 组件对象
                //WordApp.Quit(ref oMissing, ref oMissing, ref oMissing);
            }
            catch (Exception)
            {
                //关闭wordDoc 文档对象
                WordDoc.Close(ref oMissing, ref oMissing, ref oMissing);
                //关闭wordApp 组件对象
                WordApp.Quit(ref oMissing, ref oMissing, ref oMissing);
                return false;
            }
            return true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            bool ret = false;
            SavePicCh();
            ReadTextBoxDataEn();

            ret = CreatOfficeReportEn();
            if (!ret)
            {
                MessageBox.Show("创建报表失败", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Cursor = System.Windows.Forms.Cursors.Arrow;
                return;
            }

            if (!m_isOpenDataBase)
            {
                if (!m_isCreatReport)
                {
                    ret = SaveDataToDataBase();
                    if (!ret)
                    {
                        MessageBox.Show("创建报表失败,请检查数据信息", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        this.Cursor = System.Windows.Forms.Cursors.Arrow;
                        return;
                    }
                    m_isCreatReport = true;
                }
            }

            //this.Cursor = System.Windows.Forms.Cursors.Arrow;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            m_SelectHistoryData_No2Handle = new SelectHistoryData_No2(this, m_BenchHandle);
            m_SelectHistoryData_No2Handle.ControlBox = false;
            m_SelectHistoryData_No2Handle.ShowDialog();


            this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            if (m_isNoSelect)
            {
                this.Cursor = System.Windows.Forms.Cursors.Arrow;
                return;
            }

            bool ret = ReadInfoDataFromDataBase();
            if (!ret)
            {
                MessageBox.Show("读取测试信息失败", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Cursor = System.Windows.Forms.Cursors.Arrow;
                return;
            }

            ret = ReadTestDataFromDataBase();
            if (!ret)
            {
                MessageBox.Show("读取测试数据失败", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Cursor = System.Windows.Forms.Cursors.Arrow;
                return;
            }

            ret = ReadCurveDataFromDataBase();
            if (!ret)
            {
                MessageBox.Show("读取曲线信息失败", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Cursor = System.Windows.Forms.Cursors.Arrow;
                return;
            }
            m_isOpenDataBase = true;
            //button_CreatReport.Enabled = false;
            //button2.Enabled = false;
            //m_DrawHistoryCurveTimer_e.Start();
            DrawHisCurveEn();
            this.Cursor = System.Windows.Forms.Cursors.Arrow;
        }

        private void TableControl_English_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (TableControl_English.SelectedIndex == 0)
            {
                if (m_isOpenDataBase)
                {
                    DrawHisCurveCh();
                }
                else
                {
                    DrawCurveCh();
                }
            }

            if (TableControl_English.SelectedIndex == 1)
            {
                if (m_isOpenDataBase)
                {
                    DrawHisCurveEn();
                }
                else
                {
                    DrawCurveEn();
                }
            }
            SavePicCh();
        }

        private string m_OfficeInfo = "";//office信息
        private void ReadOfficeInfo()
        {
            string strFilePath1 = System.Windows.Forms.Application.StartupPath + @"\OfficeInfo.ini";//获取INI文件路径
            string strSecfileName = Path.GetFileNameWithoutExtension(strFilePath1);
            string doc = ContentValue1(strSecfileName, "Doc No", strFilePath1);
            m_OfficeInfo += "Doc No:" + doc + ", ";
            string ed = ContentValue1(strSecfileName, "ED", strFilePath1);
            m_OfficeInfo += "ED:" + ed + ", ";
            string date = ContentValue1(strSecfileName, "Date", strFilePath1);
            m_OfficeInfo += "Date:" + date + ", ";
            string approve = ContentValue1(strSecfileName, "Approved by", strFilePath1);
            m_OfficeInfo += "Approved by:" + approve + ".";
        }
        private string ContentValue1(string Section, string key, string path)
        {
            StringBuilder temp = new StringBuilder(1024);
            GetPrivateProfileString(Section, key, "", temp, 1024, path);
            return temp.ToString();
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
            textBox_yMax_e.Text = m_yMax.ToString();
            if (m_isOpenDataBase)
            {
                //m_DrawHistoryCurveTimer.Start();
                DrawHisCurveCh();
            }
            else
            {
                //m_DrawCurve.Start();
                DrawCurveCh();
            }
            SavePicCh();
        }

        private void textBox_yMax_e_TextChanged(object sender, EventArgs e)
        {
            try
            {
                m_yMax = Convert.ToUInt32(textBox_yMax_e.Text);
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
            textBox_yMax_e.Text = m_yMax.ToString();
            if (m_isOpenDataBase)
            {

                //m_DrawHistoryCurveTimer_e.Start();
                DrawHisCurveEn();
            }
            else
            {
                //m_DrawCurve_e.Start();
                DrawCurveEn();
            }
            SavePicCh();
        }

        private void textBox_StartTime_KeyDown(object sender, KeyEventArgs e)
        {
            int keyCode = e.KeyValue;
            if (keyCode != 13)
            {
                return;
            }

            m_DisStartTime = Convert.ToDateTime(textBox_StartTime.Text);
            //m_DisEndTime = Convert.ToDateTime(textBox_EndTime.Text);
            textBox_StartTime.Text = m_DisStartTime.ToLongTimeString();
            textBox_StartTime.BorderStyle = BorderStyle.None;
            textBox_StartTime.BackColor = textBox_StartTime_e.BackColor;
            textBox_StartTime.Size = new Size(54, 14);
            m_isTextBoxSelect = false;

            DateTime OldBaseTime = default(DateTime);
            DateTime OldEndTime = default(DateTime);
            DateTime startTime = default(DateTime);
            DateTime endTime = default(DateTime);
            try
            {
                if (m_isOpenDataBase)
                {
                    OldBaseTime = m_HistoryStartTime;
                    OldEndTime = m_HistoryEndTime;
                    startTime = m_DisStartTime;
                    if (m_DisEndTime == default(DateTime))
                    {
                        m_DisEndTime = m_HistoryEndTime;
                    }
                    endTime = m_DisEndTime;
                    if ((startTime - endTime).TotalSeconds > 0 || (startTime - OldEndTime).TotalSeconds > 0)
                    {
                        startTime = OldBaseTime;
                        endTime = OldEndTime;
                        m_DisStartTime = startTime;
                        m_DisEndTime = endTime;
                        textBox_StartTime.Text = startTime.ToLongTimeString();
                        textBox_EndTime.Text = endTime.ToLongTimeString();
                    }
                    if ((startTime - OldBaseTime).TotalSeconds < 0)
                    {
                        startTime = OldBaseTime;
                        m_DisStartTime = startTime;
                        textBox_StartTime.Text = startTime.ToLongTimeString();
                    }
                }
                else
                {
                    OldBaseTime = m_BenchHandle.m_StartTime;
                    OldEndTime = m_BenchHandle.m_EndTime;
                    startTime = m_DisStartTime;
                    if (m_DisEndTime == default(DateTime))
                    {
                        m_DisEndTime = m_BenchHandle.m_EndTime; ;
                    }
                    endTime = m_DisEndTime;

                    if ((startTime - endTime).TotalSeconds > 0 || (startTime - OldEndTime).TotalSeconds > 0)
                    {
                        startTime = OldBaseTime;
                        endTime = OldEndTime;
                        m_DisStartTime = startTime;
                        m_DisEndTime = endTime;
                        textBox_StartTime.Text = startTime.ToLongTimeString();
                        textBox_EndTime.Text = endTime.ToLongTimeString();
                    }
                    if ((startTime - OldBaseTime).TotalSeconds < 0)
                    {
                        startTime = OldBaseTime;
                        m_DisStartTime = startTime;
                        textBox_StartTime.Text = startTime.ToLongTimeString();
                    }
                }
            }
            catch (Exception)
            {
                startTime = OldBaseTime;
                endTime = OldEndTime;
                m_DisStartTime = startTime;
                m_DisEndTime = endTime;
                textBox_StartTime.Text = startTime.ToLongTimeString();
                textBox_EndTime.Text = endTime.ToLongTimeString();
            }
            float xStart = (float)(startTime - OldBaseTime).TotalSeconds;
            float xEnd = (float)(endTime - OldBaseTime).TotalSeconds;
            m_BenchHandle.m_DrawCurveHandle.ZoomOutBaseXY(panel_Result, xStart, 0, xEnd, m_yMax);
            RefreshTimeLabel(startTime, endTime);
        }

        private DateTime m_DisStartTime = default(DateTime);
        private bool m_isTextBoxSelect = false;
        private void textBox_StartTime_MouseClick(object sender, MouseEventArgs e)
        {
            if (m_isTextBoxSelect)
            {
                return;
            }
            textBox_StartTime.Size = new Size(120, 14);
            DateTime startTime = default(DateTime);
            if (m_isOpenDataBase)
            {
                startTime = m_HistoryStartTime;
                textBox_StartTime.Text = startTime.ToShortDateString() + " " + startTime.ToLongTimeString();
            }
            else
            {
                startTime = m_BenchHandle.m_StartTime;
                textBox_StartTime.Text = startTime.ToShortDateString() + " " + startTime.ToLongTimeString();
            }
            textBox_StartTime.BorderStyle = BorderStyle.Fixed3D;
            textBox_StartTime.BackColor = Color.Red;
            m_isTextBoxSelect = true;
        }

        private void textBox_EndTime_KeyDown(object sender, KeyEventArgs e)
        {
            int keyCode = e.KeyValue;
            if (keyCode != 13)
            {
                return;
            }

            //m_DisStartTime = Convert.ToDateTime(textBox_StartTime.Text);
            m_DisEndTime = Convert.ToDateTime(textBox_EndTime.Text);
            textBox_EndTime.Text = m_DisEndTime.ToLongTimeString();
            textBox_EndTime.BorderStyle = BorderStyle.None;
            textBox_EndTime.BackColor = textBox_EndTime_e.BackColor;
            textBox_EndTime.Size = new Size(54, 14);
            m_isTextBoxSelect = false;

            DateTime OldBaseTime = default(DateTime);
            DateTime OldEndTime = default(DateTime);
            DateTime startTime = default(DateTime);
            DateTime endTime = default(DateTime);
            try
            {
                if (m_isOpenDataBase)
                {
                    OldBaseTime = m_HistoryStartTime;
                    OldEndTime = m_HistoryEndTime;
                    if (m_DisStartTime == default(DateTime))
                    {
                        m_DisStartTime = m_HistoryStartTime;
                    }
                    startTime = m_DisStartTime;
                    endTime = m_DisEndTime;
                    if ((endTime - startTime).TotalSeconds < 0 || (endTime - OldBaseTime).TotalSeconds < 0)
                    {
                        startTime = OldBaseTime;
                        endTime = OldEndTime;
                        m_DisStartTime = startTime;
                        m_DisEndTime = endTime;
                        textBox_StartTime.Text = startTime.ToLongTimeString();
                        textBox_EndTime.Text = endTime.ToLongTimeString();
                    }
                    if ((endTime - OldEndTime).TotalSeconds > 0)
                    {
                        endTime = OldEndTime;
                        m_DisEndTime = endTime;
                        textBox_EndTime.Text = endTime.ToLongTimeString();
                    }
                }
                else
                {
                    OldBaseTime = m_BenchHandle.m_StartTime;
                    OldEndTime = m_BenchHandle.m_EndTime;
                    if (m_DisStartTime == default(DateTime))
                    {
                        m_DisStartTime = m_BenchHandle.m_StartTime;
                    }
                    startTime = m_DisStartTime;
                    endTime = m_DisEndTime;
                    if ((endTime - startTime).TotalSeconds < 0 || (endTime - OldBaseTime).TotalSeconds < 0)
                    {
                        startTime = OldBaseTime;
                        endTime = OldEndTime;
                        m_DisStartTime = startTime;
                        m_DisEndTime = endTime;
                        textBox_StartTime.Text = startTime.ToLongTimeString();
                        textBox_EndTime.Text = endTime.ToLongTimeString();
                    }
                    if ((endTime - OldEndTime).TotalSeconds > 0)
                    {
                        endTime = OldEndTime;
                        m_DisEndTime = endTime;
                        textBox_EndTime.Text = endTime.ToLongTimeString();
                    }
                }
            }
            catch (Exception)
            {
                startTime = OldBaseTime;
                endTime = OldEndTime;
                m_DisStartTime = startTime;
                m_DisEndTime = endTime;
                textBox_StartTime.Text = startTime.ToLongTimeString();
                textBox_EndTime.Text = endTime.ToLongTimeString();
            }

            float xStart = (float)(startTime - OldBaseTime).TotalSeconds;
            float xEnd = (float)(endTime - OldBaseTime).TotalSeconds;
            m_BenchHandle.m_DrawCurveHandle.ZoomOutBaseXY(panel_Result, xStart, 0, xEnd, m_yMax);
            RefreshTimeLabel(startTime, endTime);
        }

        private DateTime m_DisEndTime = default(DateTime);
        private void textBox_EndTime_MouseClick(object sender, MouseEventArgs e)
        {
            if (m_isTextBoxSelect)
            {
                return;
            }
            textBox_EndTime.Size = new Size(120, 14);
            DateTime EndTime = default(DateTime);
            if (m_isOpenDataBase)
            {
                EndTime = m_HistoryEndTime;
                textBox_EndTime.Text = EndTime.ToShortDateString() + " " + EndTime.ToLongTimeString();
            }
            else
            {
                EndTime = m_BenchHandle.m_EndTime;
                textBox_EndTime.Text = EndTime.ToShortDateString() + " " + EndTime.ToLongTimeString();
            }
            textBox_EndTime.BorderStyle = BorderStyle.Fixed3D;
            textBox_EndTime.BackColor = Color.Red;
            m_isTextBoxSelect = true;
        }

        private void textBox_StartTime_e_KeyDown(object sender, KeyEventArgs e)
        {
            int keyCode = e.KeyValue;
            if (keyCode != 13)
            {
                return;
            }

            m_DisStartTime = Convert.ToDateTime(textBox_StartTime_e.Text);
            //m_DisEndTime = Convert.ToDateTime(textBox_EndTime.Text);
            textBox_StartTime_e.Text = m_DisStartTime.ToLongTimeString();
            textBox_StartTime_e.BorderStyle = BorderStyle.None;
            textBox_StartTime_e.BackColor = textBox_StartTime.BackColor;
            textBox_StartTime_e.Size = new Size(54, 14);
            m_isTextBoxSelect = false;

            DateTime OldBaseTime = default(DateTime);
            DateTime OldEndTime = default(DateTime);
            DateTime startTime = default(DateTime);
            DateTime endTime = default(DateTime);
            try
            {
                if (m_isOpenDataBase)
                {
                    OldBaseTime = m_HistoryStartTime;
                    OldEndTime = m_HistoryEndTime;
                    startTime = m_DisStartTime;
                    if (m_DisEndTime == default(DateTime))
                    {
                        m_DisEndTime = m_HistoryEndTime;
                    }
                    endTime = m_DisEndTime;
                    if ((startTime - endTime).TotalSeconds > 0 || (startTime - OldEndTime).TotalSeconds > 0)
                    {
                        startTime = OldBaseTime;
                        endTime = OldEndTime;
                        m_DisStartTime = startTime;
                        m_DisEndTime = endTime;
                        textBox_StartTime_e.Text = startTime.ToLongTimeString();
                        textBox_EndTime_e.Text = endTime.ToLongTimeString();
                    }
                    if ((startTime - OldBaseTime).TotalSeconds < 0)
                    {
                        startTime = OldBaseTime;
                        m_DisStartTime = startTime;
                        textBox_StartTime_e.Text = startTime.ToLongTimeString();
                    }
                }
                else
                {
                    OldBaseTime = m_BenchHandle.m_StartTime;
                    OldEndTime = m_BenchHandle.m_EndTime;
                    startTime = m_DisStartTime;
                    if (m_DisEndTime == default(DateTime))
                    {
                        m_DisEndTime = m_BenchHandle.m_EndTime; ;
                    }
                    endTime = m_DisEndTime;
                    if ((startTime - endTime).TotalSeconds > 0 || (startTime - OldEndTime).TotalSeconds > 0)
                    {
                        startTime = OldBaseTime;
                        endTime = OldEndTime;
                        m_DisStartTime = startTime;
                        m_DisEndTime = endTime;
                        textBox_StartTime_e.Text = startTime.ToLongTimeString();
                        textBox_EndTime_e.Text = endTime.ToLongTimeString();
                    }
                    if ((startTime - OldBaseTime).TotalSeconds < 0)
                    {
                        startTime = OldBaseTime;
                        m_DisStartTime = startTime;
                        textBox_StartTime_e.Text = startTime.ToLongTimeString();
                    }
                }
            }
            catch (Exception)
            {
                startTime = OldBaseTime;
                endTime = OldEndTime;
                m_DisStartTime = startTime;
                m_DisEndTime = endTime;
                textBox_StartTime_e.Text = startTime.ToLongTimeString();
                textBox_EndTime_e.Text = endTime.ToLongTimeString();
            }
            float xStart = (float)(startTime - OldBaseTime).TotalSeconds;
            float xEnd = (float)(endTime - OldBaseTime).TotalSeconds;
            m_BenchHandle.m_DrawCurveHandle.ZoomOutBaseXY(panel_Result_e, xStart, 0, xEnd, m_yMax);
            RefreshTimeLabel_e(startTime, endTime);
        }

        private void textBox_StartTime_e_MouseClick(object sender, MouseEventArgs e)
        {
            if (m_isTextBoxSelect)
            {
                return;
            }
            textBox_StartTime_e.Size = new Size(120, 14);
            DateTime startTime = default(DateTime);
            if (m_isOpenDataBase)
            {
                startTime = m_HistoryStartTime;
                textBox_StartTime_e.Text = startTime.ToShortDateString() + " " + startTime.ToLongTimeString();
            }
            else
            {
                startTime = m_BenchHandle.m_StartTime;
                textBox_StartTime_e.Text = startTime.ToShortDateString() + " " + startTime.ToLongTimeString();
            }
            textBox_StartTime_e.BorderStyle = BorderStyle.Fixed3D;
            textBox_StartTime_e.BackColor = Color.Red;
            m_isTextBoxSelect = true;
        }

        private void textBox_EndTime_e_KeyDown(object sender, KeyEventArgs e)
        {
            int keyCode = e.KeyValue;
            if (keyCode != 13)
            {
                return;
            }

            //m_DisStartTime = Convert.ToDateTime(textBox_StartTime.Text);
            m_DisEndTime = Convert.ToDateTime(textBox_EndTime_e.Text);
            textBox_EndTime_e.Text = m_DisEndTime.ToLongTimeString();
            textBox_EndTime_e.BorderStyle = BorderStyle.None;
            textBox_EndTime_e.BackColor = textBox_EndTime.BackColor;
            textBox_EndTime_e.Size = new Size(54, 14);
            m_isTextBoxSelect = false;

            DateTime OldBaseTime = default(DateTime);
            DateTime OldEndTime = default(DateTime);
            DateTime startTime = default(DateTime);
            DateTime endTime = default(DateTime);
            try
            {
                if (m_isOpenDataBase)
                {
                    OldBaseTime = m_HistoryStartTime;
                    OldEndTime = m_HistoryEndTime;
                    if (m_DisStartTime == default(DateTime))
                    {
                        m_DisStartTime = m_HistoryStartTime;
                    }
                    startTime = m_DisStartTime;
                    endTime = m_DisEndTime;
                    if ((endTime - startTime).TotalSeconds < 0 || (endTime - OldBaseTime).TotalSeconds < 0)
                    {
                        startTime = OldBaseTime;
                        endTime = OldEndTime;
                        m_DisStartTime = startTime;
                        m_DisEndTime = endTime;
                        textBox_StartTime_e.Text = startTime.ToLongTimeString();
                        textBox_EndTime_e.Text = endTime.ToLongTimeString();
                    }
                    if ((endTime - OldEndTime).TotalSeconds > 0)
                    {
                        endTime = OldEndTime;
                        m_DisEndTime = endTime;
                        textBox_EndTime_e.Text = endTime.ToLongTimeString();
                    }
                }
                else
                {
                    OldBaseTime = m_BenchHandle.m_StartTime;
                    OldEndTime = m_BenchHandle.m_EndTime;
                    if (m_DisStartTime == default(DateTime))
                    {
                        m_DisStartTime = m_BenchHandle.m_StartTime;
                    }
                    startTime = m_DisStartTime;
                    endTime = m_DisEndTime;
                    if ((endTime - startTime).TotalSeconds < 0 || (endTime - OldBaseTime).TotalSeconds < 0)
                    {
                        startTime = OldBaseTime;
                        endTime = OldEndTime;
                        m_DisStartTime = startTime;
                        m_DisEndTime = endTime;
                        textBox_StartTime_e.Text = startTime.ToLongTimeString();
                        textBox_EndTime_e.Text = endTime.ToLongTimeString();
                    }
                    if ((endTime - OldEndTime).TotalSeconds > 0)
                    {
                        endTime = OldEndTime;
                        m_DisEndTime = endTime;
                        textBox_EndTime_e.Text = endTime.ToLongTimeString();
                    }
                }
            }
            catch (Exception)
            {
                startTime = OldBaseTime;
                endTime = OldEndTime;
                m_DisStartTime = startTime;
                m_DisEndTime = endTime;
                textBox_StartTime_e.Text = startTime.ToLongTimeString();
                textBox_EndTime_e.Text = endTime.ToLongTimeString();
            }

            float xStart = (float)(startTime - OldBaseTime).TotalSeconds;
            float xEnd = (float)(endTime - OldBaseTime).TotalSeconds;
            m_BenchHandle.m_DrawCurveHandle.ZoomOutBaseXY(panel_Result_e, xStart, 0, xEnd, m_yMax);
            RefreshTimeLabel_e(startTime, endTime);
        }

        private void textBox_EndTime_e_MouseClick(object sender, MouseEventArgs e)
        {
            if (m_isTextBoxSelect)
            {
                return;
            }
            textBox_StartTime.Size = new Size(120, 14);
            DateTime startTime = default(DateTime);
            if (m_isOpenDataBase)
            {
                startTime = m_HistoryStartTime;
                textBox_StartTime.Text = startTime.ToShortDateString() + " " + startTime.ToLongTimeString();
            }
            else
            {
                startTime = m_BenchHandle.m_StartTime;
                textBox_StartTime.Text = startTime.ToShortDateString() + " " + startTime.ToLongTimeString();
            }
            textBox_StartTime.BorderStyle = BorderStyle.Fixed3D;
            textBox_StartTime.BackColor = Color.Red;
            m_isTextBoxSelect = true;
        }

        private void RefreshTimeLabel(DateTime StartTime, DateTime EndTime)
        {
            TimeSpan tp = EndTime - StartTime;
            int second = (int)tp.TotalSeconds;
            label_X1.Text = StartTime.ToString("HH:mm:ss");
            //textBox_StartTime.Text = label_X1.Text;

            int H = 0;
            int M = 0;
            int S = 0;
            GetlableTime(second / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X2.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(second * 2 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X3.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(second * 3 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X4.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(second * 4 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X5.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(second * 5 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X6.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();

            GetlableTime(second * 6 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X7.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            label_X7.Text = EndTime.ToString("HH:mm:ss");
            //textBox_EndTime.Text = label_X7.Text;
        }

        private void RefreshTimeLabel_e(DateTime StartTime, DateTime EndTime)
        {
            TimeSpan tp = EndTime - StartTime;
            int second = (int)tp.TotalSeconds;
            label_X1_e.Text = StartTime.ToString("HH:mm:ss");
            //textBox_StartTime_e.Text = label_X1_e.Text;

            int H = 0;
            int M = 0;
            int S = 0;
            GetlableTime(second / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X2_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 2 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X3_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 3 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X4_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 4 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X5_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 5 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X6_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            GetlableTime(second * 6 / 6, ref H, ref M, ref S);
            S += StartTime.Second;
            if (S > 59)
            {
                M++;
                S -= 60;
            }
            M += StartTime.Minute;
            if (M > 59)
            {
                H++;
                M -= 60;
            }
            H += StartTime.Hour;
            if (H > 23)
            {
                H = H - 24;
            }
            label_X7_e.Text = H.ToString() + ":" + M.ToString() + ":" + S.ToString();
            label_X7_e.Text = EndTime.ToString("HH:mm:ss");
            //textBox_EndTime_e.Text = label_X7_e.Text;
        }











    }
}
