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
    public partial class SettingForm_No1 : Form
    {
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filepath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retval, int size, string filePath);


        private Form1 m_MainFrameHandle;
        private Bench_1 m_ParentFormHandle;
        private BaseSetting_No1 m_BaseSettingHandle;
        private string strFilePath = Application.StartupPath + @"\1#\TestConfig.ini";//获取INI文件路径
        private string strSec = "TestConfig"; //INI文件名

        private bool m_KeepSelect_No1 = true;
        private bool m_KeepSelect_No2 = false;
        private bool m_KeepSelect_No3 = false;

        public SettingForm_No1(Form1 Handle, Bench_1 ParentForHandle)
        {
            InitializeComponent();
            m_MainFrameHandle = Handle;
            m_ParentFormHandle = ParentForHandle;
        }

        private void SettingForm_No1_Load(object sender, EventArgs e)
        {
            m_TextBoxBKOld = Color.Black;
            Load_INI();
            if (Button_Selection_1.Text == "开")
            {
                TextBox_KeepPressure_1.Enabled = true;
                TextBox_KeepTime_1.Enabled = true;
                m_KeepSelect_No1 = true;
            }
            else
            {
                TextBox_KeepPressure_1.Enabled = false;
                TextBox_KeepTime_1.Enabled = false;
                m_KeepSelect_No1 = false;
            }

            if (Button_Selection_2.Text == "开")
            {
                TextBox_KeepPressure_2.Enabled = true;
                TextBox_KeepTime_2.Enabled = true;
                m_KeepSelect_No2 = true;
            }
            else
            {
                TextBox_KeepPressure_2.Enabled = false;
                TextBox_KeepTime_2.Enabled = false;
                m_KeepSelect_No2 = false;
            }

            if (Button_Selection_3.Text == "开")
            {
                TextBox_KeepPressure_3.Enabled = true;
                TextBox_KeepTime_3.Enabled = true;
                m_KeepSelect_No3 = true;
            }
            else
            {
                TextBox_KeepPressure_3.Enabled = false;
                TextBox_KeepTime_3.Enabled = false;
                m_KeepSelect_No3 = false;
            }
        }

        private void Load_INI()
        {
            if (File.Exists(strFilePath))//读取时先要判读INI文件是否存在
            {
                strSec = Path.GetFileNameWithoutExtension(strFilePath);
                TextBox_KeepPressure_1.Text = ContentValue(strSec, "KeepPressure_No1_Bench1");
                TextBox_KeepTime_1.Text = ContentValue(strSec, "KeepTime_No1_Bench1");
                string ButtonState = ContentValue(strSec, "KeepSelect_No1_Bench1");
                if (ButtonState == "OFF")
                {
                    Button_Selection_1.BackColor = Color.Red;
                    Button_Selection_1.Text = "关";

                }
                else
                {
                    Button_Selection_1.BackColor = Color.Green;
                    Button_Selection_1.Text = "开";
                }
                TextBox_KeepPressure_2.Text = ContentValue(strSec, "KeepPressure_No2_Bench1");
                TextBox_KeepTime_2.Text = ContentValue(strSec, "KeepTime_No2_Bench1");
                ButtonState = ContentValue(strSec, "KeepSelect_No2_Bench1");
                if (ButtonState == "OFF")
                {
                    Button_Selection_2.BackColor = Color.Red;
                    Button_Selection_2.Text = "关";
                }
                else
                {
                    Button_Selection_2.BackColor = Color.Green;
                    Button_Selection_2.Text = "开";
                }
                TextBox_KeepPressure_3.Text = ContentValue(strSec, "KeepPressure_No3_Bench1");
                TextBox_KeepTime_3.Text = ContentValue(strSec, "KeepTime_No3_Bench1");
                ButtonState = ContentValue(strSec, "KeepSelect_No3_Bench1");
                if (ButtonState == "OFF")
                {
                    Button_Selection_3.BackColor = Color.Red;
                    Button_Selection_3.Text = "关";
                }
                else
                {
                    Button_Selection_3.BackColor = Color.Green;
                    Button_Selection_3.Text = "开";
                }

                TextBox_Text_No.Text = ContentValue(strSec, "TestNo_Bench1");

            }
            else
            {
                MessageBox.Show("INI文件不存在");
            }
        }

        public void FreshForm()
        {
            bool ret = ReadPara();
            if (!ret)
            {
                MessageBox.Show("读取失败，请重试", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private Color m_TextBoxBKOld;
        private void SettingForm_No1_Shown(object sender, EventArgs e)
        {
            bool ret = ReadPara();
            if (!ret)
            {
                MessageBox.Show("读取失败，请重试", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private bool ReadPara()
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

            int code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepPressure_No1_Bench1", ref KeepPressure1, ref dt);
            if (code != 1)
            {
                return false;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepPressure_No2_Bench1", ref KeepPressure2, ref dt);
            if (code != 1)
            {
                return false;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepPressure_No3_Bench1", ref KeepPressure3, ref dt);
            if (code != 1)
            {
                return false;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepTime_No1_Bench1", ref KeepTime1, ref dt);
            if (code != 1)
            {
                return false;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepTime_No2_Bench1", ref KeepTime2, ref dt);
            if (code != 1)
            {
                return false;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepTime_No3_Bench1", ref KeepTime3, ref dt);
            if (code != 1)
            {
                return false;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("ChanelSelect_No1_Bench1", ref ChanelSelect1, ref dt);
            if (code != 1)
            {
                return false;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("ChanelSelect_No2_Bench1", ref ChanelSelect2, ref dt);
            if (code != 1)
            {
                return false;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("ChanelSelect_No3_Bench1", ref ChanelSelect3, ref dt);
            if (code != 1)
            {
                return false;
            }

            if (ChanelSelect1 == true)
            {
                TextBox_KeepPressure_1.Enabled = true;
                TextBox_KeepTime_1.Enabled = true;
                TextBox_KeepPressure_1.Text = KeepPressure1.ToString("0.00");
                TextBox_KeepTime_1.Text = KeepTime1.ToString("0.00");
                Button_Selection_1.Text = "开";
                Button_Selection_1.BackColor = Color.Green;
                m_KeepSelect_No1 = true;
            }
            else
            {
            }

            if (ChanelSelect2 == true)
            {
                TextBox_KeepPressure_2.Enabled = true;
                TextBox_KeepTime_2.Enabled = true;
                TextBox_KeepPressure_2.Text = KeepPressure2.ToString("0.00");
                TextBox_KeepTime_2.Text = KeepTime2.ToString("0.00");
                Button_Selection_2.Text = "开";
                Button_Selection_2.BackColor = Color.Green;
                m_KeepSelect_No2 = true;
            }
            else
            {
                TextBox_KeepPressure_2.Enabled = false;
                TextBox_KeepTime_2.Enabled = false;
                Button_Selection_2.Text = "关";
                Button_Selection_2.BackColor = Color.Red;
                m_KeepSelect_No2 = false;
            }

            if (ChanelSelect3 == true)
            {
                TextBox_KeepPressure_3.Enabled = true;
                TextBox_KeepTime_3.Enabled = true;
                TextBox_KeepPressure_3.Text = KeepPressure3.ToString("0.00");
                TextBox_KeepTime_3.Text = KeepTime3.ToString("0.00");
                Button_Selection_3.Text = "开";
                Button_Selection_3.BackColor = Color.Green;
                m_KeepSelect_No3 = true;
            }
            else
            {
                TextBox_KeepPressure_3.Enabled = false;
                TextBox_KeepTime_3.Enabled = false;
                Button_Selection_3.Text = "关";
                Button_Selection_3.BackColor = Color.Red;
                m_KeepSelect_No3 = false;
            }
            return true;
        }

        private string ContentValue(string Section, string key)
        {
            StringBuilder temp = new StringBuilder(1024);
            GetPrivateProfileString(Section, key, "", temp, 1024, strFilePath);
            return temp.ToString();
        }

        private void Button_Return_Click(object sender, EventArgs e)
        {
            this.Hide();
            m_ParentFormHandle.m_TestNo = TextBox_Text_No.Text;
            m_ParentFormHandle.Show();
        }

        private void TextBox_KeepPressure_1_TextChanged(object sender, EventArgs e)
        {
            float Press1 = 0;
            try
            {
                Press1 = Convert.ToSingle(TextBox_KeepPressure_1.Text);
            }
            catch (Exception)
            {
                TextBox_KeepPressure_1.BackColor = Color.Red;
                //MessageBox.Show("保压压力1输入错误，请重新检查输入参数", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            TextBox_KeepPressure_1.BackColor = m_TextBoxBKOld;
            int code = m_MainFrameHandle.m_PLCCommHandle.WriteData("KeepPressure_No1_Bench1", Press1);
            if (code != 1)
            {
                //m_MainFrameHandle.m_LogWR.WriteLog(m_LogFilePath, "SetParamentForm");
                TextBox_KeepPressure_1.BackColor = Color.Red;
                MessageBox.Show("保压压力1写入失败,错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            WritePrivateProfileString(strSec, "KeepPressure_No1_Bench1", Press1.ToString(), strFilePath);

        }

        private void TextBox_KeepTime_1_TextChanged(object sender, EventArgs e)
        {
            float Time1 = 0;
            try
            {
                Time1 = Convert.ToSingle(TextBox_KeepTime_1.Text);
            }
            catch (Exception)
            {
                TextBox_KeepPressure_1.BackColor = Color.Red;
                //MessageBox.Show("保压时间1输入错误，请重新检查输入参数", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            TextBox_KeepPressure_1.BackColor = m_TextBoxBKOld;
            int code = m_MainFrameHandle.m_PLCCommHandle.WriteData("KeepTime_No1_Bench1", Time1);
            if (code != 1)
            {
                TextBox_KeepPressure_1.BackColor = Color.Red;
                //m_MainFrameHandle.m_LogWR.WriteLog(m_LogFilePath, "SetParamentForm");
                //MessageBox.Show("保压时间1写入失败,错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            WritePrivateProfileString(strSec, "KeepTime_No1_Bench1", Time1.ToString(), strFilePath);
        }

        private void Button_Selection_1_Click(object sender, EventArgs e)
        {
            bool bSelect1 = true;
            int code = m_MainFrameHandle.m_PLCCommHandle.WriteData("ChanelSelect_No1_Bench1", bSelect1);
            if (code != 1)
            {
                //m_MainFrameHandle.m_LogWR.WriteLog(m_LogFilePath, "SetParamentForm");
                MessageBox.Show("保压选择1写入失败,错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (Button_Selection_1.Text == "开")
            {
                //TextBox_KeepPressure_1.Enabled = false;
                //TextBox_KeepTime_1.Enabled = false;
                if (m_KeepSelect_No3 == false && m_KeepSelect_No2 == false)
                {
                    MessageBox.Show("不能关闭，必须保留一组测试", "Info");
                    return;
                }
            }
            else
            {
                //TextBox_KeepPressure_1.Enabled = true;
                //TextBox_KeepTime_1.Enabled = true;
            }
            WritePrivateProfileString(strSec, "KeepSelect_No1_Bench1", "ON", strFilePath);
        }

        private void TextBox_KeepPressure_2_TextChanged(object sender, EventArgs e)
        {
            float Press2 = 0;
            try
            {
                Press2 = Convert.ToSingle(TextBox_KeepPressure_2.Text);
            }
            catch (Exception)
            {
                TextBox_KeepPressure_2.BackColor = Color.Red;
                //MessageBox.Show("保压压力1输入错误，请重新检查输入参数", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            TextBox_KeepPressure_2.BackColor = m_TextBoxBKOld;
            int code = m_MainFrameHandle.m_PLCCommHandle.WriteData("KeepPressure_No2_Bench1", Press2);
            if (code != 1)
            {
                //m_MainFrameHandle.m_LogWR.WriteLog(m_LogFilePath, "SetParamentForm");
                TextBox_KeepPressure_2.BackColor = Color.Red;
                MessageBox.Show("保压压力1写入失败,错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            WritePrivateProfileString(strSec, "KeepPressure_No2_Bench1", Press2.ToString(), strFilePath);

        }

        private void TextBox_KeepTime_2_TextChanged(object sender, EventArgs e)
        {
            float Time2 = 0;
            try
            {
                Time2 = Convert.ToSingle(TextBox_KeepTime_2.Text);
            }
            catch (Exception)
            {
                TextBox_KeepPressure_2.BackColor = Color.Red;
                //MessageBox.Show("保压时间1输入错误，请重新检查输入参数", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            TextBox_KeepPressure_2.BackColor = m_TextBoxBKOld;
            int code = m_MainFrameHandle.m_PLCCommHandle.WriteData("KeepTime_No2_Bench1", Time2);
            if (code != 1)
            {
                TextBox_KeepPressure_2.BackColor = Color.Red;
                //m_MainFrameHandle.m_LogWR.WriteLog(m_LogFilePath, "SetParamentForm");
                //MessageBox.Show("保压时间1写入失败,错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            WritePrivateProfileString(strSec, "KeepTime_No2_Bench1", Time2.ToString(), strFilePath);
        }

        private void Button_Selection_2_Click(object sender, EventArgs e)
        {
            string sSelet = "OFF";
            int code = 0;
            if (Button_Selection_2.Text == "开")
            {
                if (m_KeepSelect_No3 == false && m_KeepSelect_No1)
                {
                    TextBox_KeepPressure_2.Enabled = false;
                    TextBox_KeepTime_2.Enabled = false;
                    m_KeepSelect_No2 = false;
                    Button_Selection_2.BackColor = Color.Red;
                    Button_Selection_2.Text = "关";
                    sSelet = "OFF";
                    bool bSelect2 = false;
                    code = m_MainFrameHandle.m_PLCCommHandle.WriteData("ChanelSelect_No2_Bench1", bSelect2);
                    if (code != 1)
                    {
                        //m_MainFrameHandle.m_LogWR.WriteLog(m_LogFilePath, "SetParamentForm");
                        MessageBox.Show("保压选择2写入失败,错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("请先关闭序列3", "Info");
                }
            }
            else
            {
                TextBox_KeepPressure_2.Enabled = true;
                TextBox_KeepTime_2.Enabled = true;
                if (m_KeepSelect_No3 == false && m_KeepSelect_No1)
                {
                    m_KeepSelect_No2 = true;
                    Button_Selection_2.BackColor = Color.Green;
                    Button_Selection_2.Text = "开";
                    sSelet = "ON";
                    bool bSelect2 = true;
                    code = m_MainFrameHandle.m_PLCCommHandle.WriteData("ChanelSelect_No2_Bench1", bSelect2);
                    if (code != 1)
                    {
                        //m_MainFrameHandle.m_LogWR.WriteLog(m_LogFilePath, "SetParamentForm");
                        MessageBox.Show("保压选择2写入失败,错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            WritePrivateProfileString(strSec, "KeepSelect_No2_Bench1", sSelet, strFilePath);
        }

        private void TextBox_KeepPressure_3_TextChanged(object sender, EventArgs e)
        {
            float Press3 = 0;
            try
            {
                Press3 = Convert.ToSingle(TextBox_KeepPressure_3.Text);
            }
            catch (Exception)
            {
                TextBox_KeepPressure_3.BackColor = Color.Red;
                //MessageBox.Show("保压压力1输入错误，请重新检查输入参数", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            TextBox_KeepPressure_3.BackColor = m_TextBoxBKOld;
            int code = m_MainFrameHandle.m_PLCCommHandle.WriteData("KeepPressure_No3_Bench1", Press3);
            if (code != 1)
            {
                //m_MainFrameHandle.m_LogWR.WriteLog(m_LogFilePath, "SetParamentForm");
                TextBox_KeepPressure_3.BackColor = Color.Red;
                MessageBox.Show("保压压力1写入失败,错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            WritePrivateProfileString(strSec, "KeepPressure_No3_Bench1", Press3.ToString(), strFilePath);

        }

        private void TextBox_KeepTime_3_TextChanged(object sender, EventArgs e)
        {
            float Time3 = 0;
            try
            {
                Time3 = Convert.ToSingle(TextBox_KeepTime_3.Text);
            }
            catch (Exception)
            {
                TextBox_KeepPressure_3.BackColor = Color.Red;
                //MessageBox.Show("保压时间1输入错误，请重新检查输入参数", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            TextBox_KeepPressure_3.BackColor = m_TextBoxBKOld;
            int code = m_MainFrameHandle.m_PLCCommHandle.WriteData("KeepTime_No3_Bench1", Time3);
            if (code != 1)
            {
                TextBox_KeepPressure_3.BackColor = Color.Red;
                //m_MainFrameHandle.m_LogWR.WriteLog(m_LogFilePath, "SetParamentForm");
                //MessageBox.Show("保压时间1写入失败,错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            WritePrivateProfileString(strSec, "KeepTime_No3_Bench1", Time3.ToString(), strFilePath);
        }

        private void Button_Selection_3_Click(object sender, EventArgs e)
        {
            string sSelect = "OFF";
            int code = 0;
            if (Button_Selection_3.Text == "开")
            {
                if (m_KeepSelect_No2 && m_KeepSelect_No1)
                {
                    TextBox_KeepPressure_3.Enabled = false;
                    TextBox_KeepTime_3.Enabled = false;
                    m_KeepSelect_No3 = false;
                    Button_Selection_3.BackColor = Color.Red;
                    Button_Selection_3.Text = "关";
                    sSelect = "OFF";
                    bool bSelect3 = false;
                    code = m_MainFrameHandle.m_PLCCommHandle.WriteData("ChanelSelect_No3_Bench1", bSelect3);
                    if (code != 1)
                    {
                        //m_MainFrameHandle.m_LogWR.WriteLog(m_LogFilePath, "SetParamentForm");
                        MessageBox.Show("保压选择3写入失败,错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            else
            {
                if (m_KeepSelect_No2 && m_KeepSelect_No1)
                {
                    TextBox_KeepPressure_3.Enabled = true;
                    TextBox_KeepTime_3.Enabled = true;
                    m_KeepSelect_No3 = true;
                    Button_Selection_3.BackColor = Color.Green;
                    Button_Selection_3.Text = "开";
                    sSelect = "ON";
                    bool bSelect3 = true;
                    code = m_MainFrameHandle.m_PLCCommHandle.WriteData("ChanelSelect_No3_Bench1", bSelect3);
                    if (code != 1)
                    {
                        //m_MainFrameHandle.m_LogWR.WriteLog(m_LogFilePath, "SetParamentForm");
                        MessageBox.Show("保压选择3写入失败,错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            WritePrivateProfileString(strSec, "KeepSelect_No3_Bench1", sSelect, strFilePath);
        }

        private void Button_SavePara_Click(object sender, EventArgs e)
        {
            //如果设置新的实验标号，则清除内存
            bool bSelect1 = true;
            int code = m_MainFrameHandle.m_PLCCommHandle.WriteData("ChanelSelect_No1_Bench1", bSelect1);
            if (code != 1)
            {
                //m_MainFrameHandle.m_LogWR.WriteLog(m_LogFilePath, "SetParamentForm");
                MessageBox.Show("保压选择1写入失败,错误代码：" + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string TestNo = TextBox_Text_No.Text;
            if (TestNo != m_ParentFormHandle.m_GetTestNo)
            {
                m_ParentFormHandle.m_PointFArrays.Clear();
                m_ParentFormHandle.m_TestResultLists.Clear();
                m_ParentFormHandle.m_TestSequence = 1;
                m_ParentFormHandle.m_isTestStartBaseTestNo = true;
                m_ParentFormHandle.ClearPointBuffer();
            }
            m_ParentFormHandle.m_GetTestNo = TestNo;   //保存试验编号
            WritePrivateProfileString(strSec, "TestNo_Bench1", TextBox_Text_No.Text, strFilePath);
            m_ParentFormHandle.m_isSetFlag = true;
            ushort setTestTimes = 0;
            if (Button_Selection_3.Text == "开")
            {
                setTestTimes = 3;
            }
            if (Button_Selection_2.Text == "开")
            {
                setTestTimes = 2;
            }
            if (Button_Selection_2.Text == "开")
            {
                setTestTimes = 1;
            }
            m_ParentFormHandle.m_SetTestTimes = setTestTimes;
            MessageBox.Show("设置成功", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Button_BaseSetting_Click(object sender, EventArgs e)
        {
            m_BaseSettingHandle = new BaseSetting_No1(m_MainFrameHandle, this);
            m_BaseSettingHandle.ShowDialog();
            m_BaseSettingHandle.Close();
            m_BaseSettingHandle.Dispose();
            m_BaseSettingHandle = null;
        }

        private void Button_PLC_Read_Click(object sender, EventArgs e)
        {
            bool ret = ReadPara();
            if (!ret)
            {
                MessageBox.Show("读取失败，请重试", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
    }
}
