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
    public partial class BaseSetting_No1 : Form
    {
        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filepath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retval, int size, string filePath);

        private Form1 m_MainFrameHandle;
        private SettingForm_No1 m_SettingFormHandle;
        private string strFilePath = Application.StartupPath + @"\1#\BaseSettingConfig.ini";//获取INI文件路径
        private string strSec = "BaseSettingConfig"; //INI文件名
        public BaseSetting_No1(Form1 Handle, SettingForm_No1 handle)
        {
            InitializeComponent();
            m_MainFrameHandle = Handle;
            m_SettingFormHandle = handle;
        }

        private void BaseSetting_No1_Load(object sender, EventArgs e)
        {
            this.ControlBox = false;
        }

        private void BaseSetting_No1_Shown(object sender, EventArgs e)
        {
            bool flag = ReadPara();
            if (!flag)
            {
                MessageBox.Show("参数读取失败", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void TB_HighBumpStartPress_TextChanged(object sender, EventArgs e)
        {
            float HighBumpStartPress = 0;
            //byte SensorAdj = 1;
            try
            {
                HighBumpStartPress = Convert.ToSingle(TB_HighBumpStartPress.Text);
            }
            catch (Exception)
            {
                TB_HighBumpStartPress.BackColor = Color.Red;
                //MessageBox.Show("输入的参数错误，请查证。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            TB_HighBumpStartPress.BackColor = Color.Black;

            int code = m_MainFrameHandle.m_PLCCommHandle.WriteData("HighBumpStartPress_Bench1", HighBumpStartPress);
            if (code != 1)
            {
                TB_HighBumpStartPress.BackColor = Color.Red;
                //MessageBox.Show("高压泵启动压力设置失败，ErrorCode = " + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void TB_SensorLength_TextChanged(object sender, EventArgs e)
        {
            float SensorLength = 0;
            try
            {
                SensorLength = Convert.ToSingle(TB_SensorLength.Text);
            }
            catch (Exception)
            {
                TB_SensorLength.BackColor = Color.Red;
                //MessageBox.Show("输入的参数错误，请查证。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            TB_SensorLength.BackColor = Color.Black;

            int code = m_MainFrameHandle.m_PLCCommHandle.WriteData("SensorLength_Bench1", SensorLength);
            if (code != 1)
            {
                TB_SensorLength.BackColor = Color.Red;
                //MessageBox.Show("传感器量程设置失败，ErrorCode = " + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void TB_SensorOffset_TextChanged(object sender, EventArgs e)
        {
            float SensorOffSet = 0.0f;
            try
            {
                SensorOffSet = Convert.ToSingle(TB_SensorOffset.Text);
            }
            catch (Exception)
            {
                TB_SensorOffset.BackColor = Color.Red;
                //MessageBox.Show("输入的参数错误，请查证。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            TB_SensorOffset.BackColor = Color.Black;
            int code = m_MainFrameHandle.m_PLCCommHandle.WriteData("SensorOffSet_Bench1", SensorOffSet);
            if (code != 1)
            {
                TB_SensorOffset.BackColor = Color.Red;
                //MessageBox.Show("传感器偏置设置失败，ErrorCode = " + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void textBox_Bench1_PreEarly_TextChanged(object sender, EventArgs e)
        {
            float EarlyPre = 0.0f;
            try
            {
                EarlyPre = Convert.ToSingle(textBox_Bench1_PreEarly.Text);
            }
            catch (Exception)
            {
                {
                    textBox_Bench1_PreEarly.BackColor = Color.Red;
                    return;
                }
            }
            textBox_Bench1_PreEarly.BackColor = Color.Black;
            int code = m_MainFrameHandle.m_PLCCommHandle.WriteData("PreEarly_Bench1", EarlyPre);
            if (code != 1)
            {
                textBox_Bench1_PreEarly.BackColor = Color.Red;
                //MessageBox.Show("传感器偏置设置失败，ErrorCode = " + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void TB_StabilityTime_TextChanged(object sender, EventArgs e)
        {
            UInt16 StabilityTime = 0;
            try
            {
                StabilityTime = Convert.ToUInt16(TB_StabilityTime.Text);
            }
            catch (Exception)
            {
                TB_StabilityTime.BackColor = Color.Red;
                //MessageBox.Show("输入的参数错误，请查证。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            TB_StabilityTime.BackColor = Color.Black;

            int code = m_MainFrameHandle.m_PLCCommHandle.WriteData("KeepPressStabilityTime_Bench1", StabilityTime);
            if (code != 1)
            {
                TB_StabilityTime.BackColor = Color.Red;
                //MessageBox.Show("保压稳定时间失败，ErrorCode = " + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void TB_OpenValveTime_TextChanged(object sender, EventArgs e)
        {
            UInt16 OpenValveTime = 0;
            try
            {
                OpenValveTime = Convert.ToUInt16(TB_OpenValveTime.Text);
            }
            catch (Exception)
            {
                TB_OpenValveTime.BackColor = Color.Red;
                //MessageBox.Show("输入的参数错误，请查证。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            TB_OpenValveTime.BackColor = Color.Black;

            int code = m_MainFrameHandle.m_PLCCommHandle.WriteData("OpenValveTime_Bench1", OpenValveTime);
            if (code != 1)
            {
                TB_OpenValveTime.BackColor = Color.Red;
                //MessageBox.Show("开阀操作时间设置失败，ErrorCode = " + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void TB_DropPressSelect_TextChanged(object sender, EventArgs e)
        {
            float DropPressSelect = 0;
            try
            {
                DropPressSelect = Convert.ToSingle(TB_DropPressSelect.Text);
            }
            catch (Exception)
            {
                TB_DropPressSelect.BackColor = Color.Red;
                //MessageBox.Show("输入的参数错误，请查证。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            TB_DropPressSelect.BackColor = Color.Black;

            int code = m_MainFrameHandle.m_PLCCommHandle.WriteData("DropPressSelect_Bench1", DropPressSelect);
            if (code != 1)
            {
                TB_DropPressSelect.BackColor = Color.Red;
                //MessageBox.Show("泄压判断压力设置失败，ErrorCode = " + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void TB_TestPressIntervalTime_TextChanged(object sender, EventArgs e)
        {
            UInt16 TestPressInterval = 0;
            //byte SensorAdj = 1;
            try
            {
                TestPressInterval = Convert.ToUInt16(TB_TestPressIntervalTime.Text);
            }
            catch (Exception)
            {
                TB_TestPressIntervalTime.BackColor = Color.Red;
                //MessageBox.Show("输入的参数错误，请查证。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            TB_TestPressIntervalTime.BackColor = Color.Black;

            int code = m_MainFrameHandle.m_PLCCommHandle.WriteData("TestPressInterval_Bench1", TestPressInterval);
            if (code != 1)
            {
                TB_TestPressIntervalTime.BackColor = Color.Red;
                //MessageBox.Show("试压间隔时间设置失败，ErrorCode = " + code.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }


        private void button_Return_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private bool ReadPara()
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

            int code = m_MainFrameHandle.m_PLCCommHandle.ReadData("HighBumpStartPress_Bench1", ref HighBumpStartPress, ref dt);
            if (code != 1)
            {
                return false;
            }
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("SensorLength_Bench1", ref SensorLength, ref dt);
            if (code != 1)
            {
                return false;
            }
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("SensorOffSet_Bench1", ref SensorOffSet, ref dt);
            if (code != 1)
            {
                return false;
            }
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("KeepPressStabilityTime_Bench1", ref StabilityTime, ref dt);
            if (code != 1)
            {
                return false;
            }
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("OpenValveTime_Bench1", ref OpenValveTime, ref dt);
            if (code != 1)
            {
                return false;
            }
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("DropPressSelect_Bench1", ref DropPressSelect, ref dt);
            if (code != 1)
            {
                return false;
            }
            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("TestPressInterval_Bench1", ref TestPressInterval, ref dt);
            if (code != 1)
            {
                return false;
            }

            code = m_MainFrameHandle.m_PLCCommHandle.ReadData("PreEarly_Bench1", ref EarlyPre, ref dt);
            if (code != 1)
            {
                return false;
            }

            TB_HighBumpStartPress.Text = HighBumpStartPress.ToString("0.00");
            TB_SensorLength.Text = SensorLength.ToString("0.00");
            TB_SensorOffset.Text = SensorOffSet.ToString("0.00");
            TB_StabilityTime.Text = StabilityTime.ToString();
            TB_OpenValveTime.Text = OpenValveTime.ToString();
            TB_DropPressSelect.Text = DropPressSelect.ToString("0.00");
            TB_TestPressIntervalTime.Text = TestPressInterval.ToString();
            textBox_Bench1_PreEarly.Text = EarlyPre.ToString("0.00");
            return true;
        }

        private void button_Set_Click(object sender, EventArgs e)
        {
            float HighBumpStartPress = 0;
            float SensorLength = 0;
            float SensorOffSet = 0.0f;
            UInt16 StabilityTime = 0;
            UInt16 OpenValveTime = 0;
            float DropPressSelect = 0;
            UInt16 TestPressInterval = 0;
            float earlyPre = 0.0f;
            try
            {
                HighBumpStartPress = Convert.ToSingle(TB_HighBumpStartPress.Text);
                SensorLength = Convert.ToSingle(TB_SensorLength.Text);
                SensorOffSet = Convert.ToSingle(TB_SensorOffset.Text);
                StabilityTime = Convert.ToUInt16(TB_StabilityTime.Text);
                OpenValveTime = Convert.ToUInt16(TB_OpenValveTime.Text);
                DropPressSelect = Convert.ToSingle(TB_DropPressSelect.Text);
                TestPressInterval = Convert.ToUInt16(TB_TestPressIntervalTime.Text);
                earlyPre = Convert.ToSingle(textBox_Bench1_PreEarly.Text);
            }
            catch (Exception)
            {
                MessageBox.Show("输入的参数错误，请查证。", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            WritePrivateProfileString(strSec, "HighBumpStartPress_Bench1", HighBumpStartPress.ToString("0.00"), strFilePath);
            WritePrivateProfileString(strSec, "SensorLength_Bench1", SensorLength.ToString("0.00"), strFilePath);
            WritePrivateProfileString(strSec, "SensorOffSet_Bench1", SensorOffSet.ToString("0.00"), strFilePath);
            WritePrivateProfileString(strSec, "KeepPressStabilityTime_Bench1", StabilityTime.ToString(), strFilePath);
            WritePrivateProfileString(strSec, "OpenValveTime_Bench1", OpenValveTime.ToString(), strFilePath);
            WritePrivateProfileString(strSec, "DropPressSelect_Bench1", DropPressSelect.ToString("0.00"), strFilePath);
            WritePrivateProfileString(strSec, "TestPressInterval_Bench1", TestPressInterval.ToString(), strFilePath);
            WritePrivateProfileString(strSec, "EarlyPre_Bench1", earlyPre.ToString(), strFilePath);
            MessageBox.Show("参数设置成功.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
        }







    }
}
