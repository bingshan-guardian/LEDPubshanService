using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using EQ2008_DataStruct;
using Microsoft.Win32;

namespace LedPublishService {
    public partial class LedPublishService : ServiceBase {

        //==========================1、节目操作函数======================//
        //添加节目
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern int User_AddProgram(int CardNum, Boolean bWaitToEnd, int iPlayTime);

        //删除所有节目
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern Boolean User_DelAllProgram(int CardNum);

        //添加单行文本区
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern int User_AddSingleText(int CardNum, ref User_SingleText pSingleText, int iProgramIndex);

        //添加文本区
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern int User_AddText(int CardNum, ref User_Text pText, int iProgramIndex);

        //添加时间区
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern int User_AddTime(int CardNum, ref User_DateTime pdateTime, int iProgramIndex);

        //添加图文区
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern int User_AddBmpZone(int CardNum, ref User_Bmp pBmp, int iProgramIndex);

        //指定图像句柄添加图片
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern bool User_AddBmp(int CardNum, int iBmpPartNum, IntPtr hBitmap, ref User_MoveSet pMoveSet,
            int iProgramIndex);

        //指定图像路径添加图片
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern bool User_AddBmpFile(int CardNum, int iBmpPartNum, string strFileName,
            ref User_MoveSet pMoveSet, int iProgramIndex);

        //添加RTF区
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern int User_AddRTF(int CardNum, ref User_RTF pRTF, int iProgramIndex);

        //添加计时区
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern int User_AddTimeCount(int CardNum, ref User_Timer pTimeCount, int iProgramIndex);

        //添加温度区
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern int User_AddTemperature(int CardNum, ref User_Temperature pTemperature, int iProgramIndex);

        //发送数据
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern Boolean User_SendToScreen(int CardNum);
        //====================================================================//       

        //=======================2、实时发送数据（高频率发送）=================//
        //实时建立连接
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern Boolean User_RealtimeConnect(int CardNum);

        //实时发送图片数据
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern Boolean User_RealtimeSendData(int CardNum, int x, int y, int iWidth, int iHeight,
            IntPtr hBitmap);

        //实时发送图片文件
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern Boolean User_RealtimeSendBmpData(int CardNum, int x, int y, int iWidth, int iHeight,
            string strFileName);

        //实时发送文本
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern Boolean User_RealtimeSendText(int CardNum, int x, int y, int iWidth, int iHeight,
            string strText, ref User_FontSet pFontInfo);

        //实时关闭连接
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern Boolean User_RealtimeDisConnect(int CardNum);

        //实时发送清屏
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern Boolean User_RealtimeScreenClear(int CardNum);
        //====================================================================//

        //==========================3、显示屏控制函数组=======================//
        //校正时间
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern Boolean User_AdjustTime(int CardNum);

        //开屏
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern Boolean User_OpenScreen(int CardNum);

        //关屏
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern Boolean User_CloseScreen(int CardNum);

        //亮度调节
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern Boolean User_SetScreenLight(int CardNum, int iLightDegreen);

        //Reload参数文件
        [DllImport("EQ2008_Dll.dll", CharSet = CharSet.Ansi)]
        public static extern void User_ReloadIniFile(string strEQ2008_Dll_Set_Path);

        [System.Runtime.InteropServices.DllImport("gdi32.dll")]
        public static extern bool DeleteObject(IntPtr hObject);

        //====================================================================//

        System.Timers.Timer publishTimer = new System.Timers.Timer();
        String publishText = "";
        RegistryKey folder = null;
        int iCardNum = 1;
        int iProgramIndex = 0;

        public LedPublishService() {
            InitializeComponent();
        }


        protected override void OnStart(string[] args) {
            this.WriteLog("Service started");
            //检查注册表项
            this.CheckRegistry();
            this.CheckIniFile();

            publishTimer.Interval = Convert.ToInt32(this.GetRegistryFolder().GetValue("publishInterval"));
            publishTimer.AutoReset = true;
            publishTimer.Enabled = true;
            publishTimer.Start();
            publishTimer.Elapsed += new System.Timers.ElapsedEventHandler(doPublishTxt);
            this.WriteLog("周期发送设置完成，间隔时间为：" + publishTimer.Interval);
            publishTxt();
        }

        private void doPublishTxt(object source, System.Timers.ElapsedEventArgs e) {
            this.WriteLog("周期发送时间到，执行发送");
            publishTxt();
        }
        protected override void OnStop() {
            this.WriteLog("Service stopped");
        }

        private void publishTxt() {
            //判断从哪种方式获取天气信息
            int sourceEnum = Convert.ToInt16(this.GetRegistryFolder().GetValue("sourceType"));
            //
            if (0.Equals(sourceEnum)) { // 从Url获取天气txt
                this.getTxtFromUrl();
                this.WriteLog("通过Url更新文本：\r\n" + this.publishText);
            } else if (1.Equals(sourceEnum)) { // 从文件获取天气txt
                this.getTxtFromFile();
                this.WriteLog("通过文件更新文本：\r\n" + this.publishText);
            } else { // 从网络天气服务获取
                this.getWeatherFromUrl();
                this.WriteLog("通过Api更新文本：\r\n" + this.publishText);
            }
            //加载屏幕参数
            String iniFilePath = this.GetRegistryFolder().GetValue("iniFile").ToString().Replace(".\\", System.Threading.Thread.GetDomain().BaseDirectory);
            User_ReloadIniFile(iniFilePath);
            this.WriteLog("加载屏幕配置文件：" + iniFilePath);
            this.iCardNum = Convert.ToInt16(this.GetRegistryFolder().GetValue("iCardNum"));
            if (!User_RealtimeConnect(this.iCardNum)) {
                this.WriteLog("连接控制屏失败 iCardNum:" + iCardNum);
                return;
            }
            this.WriteLog("连接屏幕成功 iCardNum:" + iCardNum);
            if (!User_OpenScreen(this.iCardNum)) {
                this.WriteLog("打开显示屏失败");
            } else {
                this.WriteLog("打开显示屏成功");
            }
            if (!User_DelAllProgram(this.iCardNum)) {
                this.WriteLog("清除原节目失败");
                return;
            }
            this.WriteLog("清除原节目成功");
            this.iProgramIndex = User_AddProgram(this.iCardNum, true, 40);

            User_Text Text = new User_Text();
            Text.BkColor = Convert.ToInt16(this.GetRegistryFolder().GetValue("BkColor"));
            Text.chContent = this.publishText;

            Text.PartInfo.FrameColor = Convert.ToInt16(this.GetRegistryFolder().CreateSubKey("PartInfo").GetValue("FrameColor"));
            Text.PartInfo.iFrameMode = Convert.ToInt16(this.GetRegistryFolder().CreateSubKey("PartInfo").GetValue("iFrameMode"));
            Text.PartInfo.iHeight = Convert.ToInt16(this.GetRegistryFolder().CreateSubKey("PartInfo").GetValue("iHeight"));
            Text.PartInfo.iWidth = Convert.ToInt16(this.GetRegistryFolder().CreateSubKey("PartInfo").GetValue("iWidth"));
            Text.PartInfo.iX = Convert.ToInt32(this.GetRegistryFolder().CreateSubKey("PartInfo").GetValue("iX"));
            Text.PartInfo.iY = Convert.ToInt32(this.GetRegistryFolder().CreateSubKey("PartInfo").GetValue("iY"));

            Text.FontInfo.bFontBold = Convert.ToBoolean(this.GetRegistryFolder().CreateSubKey("FontInfo").GetValue("bFontBold"));
            Text.FontInfo.bFontItaic = Convert.ToBoolean(this.GetRegistryFolder().CreateSubKey("FontInfo").GetValue("bFontItaic"));
            Text.FontInfo.bFontUnderline = Convert.ToBoolean(this.GetRegistryFolder().CreateSubKey("FontInfo").GetValue("bFontUnderline"));
            Text.FontInfo.colorFont = Convert.ToInt16(this.GetRegistryFolder().CreateSubKey("FontInfo").GetValue("colorFont"));
            Text.FontInfo.iFontSize = Convert.ToInt16(this.GetRegistryFolder().CreateSubKey("FontInfo").GetValue("iFontSize"));
            Text.FontInfo.strFontName = this.GetRegistryFolder().CreateSubKey("FontInfo").GetValue("strFontName").ToString();
            Text.FontInfo.iAlignStyle = Convert.ToInt16(this.GetRegistryFolder().CreateSubKey("FontInfo").GetValue("iAlignStyle"));
            Text.FontInfo.iVAlignerStyle = Convert.ToInt16(this.GetRegistryFolder().CreateSubKey("FontInfo").GetValue("iVAlignerStyle"));
            Text.FontInfo.iRowSpace = Convert.ToInt16(this.GetRegistryFolder().CreateSubKey("FontInfo").GetValue("iRowSpace"));

            Text.MoveSet.bClear = Convert.ToBoolean(this.GetRegistryFolder().CreateSubKey("MoveSet").GetValue("bClear"));
            Text.MoveSet.iActionSpeed = Convert.ToInt16(this.GetRegistryFolder().CreateSubKey("MoveSet").GetValue("iActionSpeed"));
            Text.MoveSet.iActionType = Convert.ToInt16(this.GetRegistryFolder().CreateSubKey("MoveSet").GetValue("iActionType"));
            Text.MoveSet.iHoldTime = Convert.ToInt16(this.GetRegistryFolder().CreateSubKey("MoveSet").GetValue("iHoldTime"));
            Text.MoveSet.iClearActionType = Convert.ToInt16(this.GetRegistryFolder().CreateSubKey("MoveSet").GetValue("iClearActionType"));
            Text.MoveSet.iClearSpeed = Convert.ToInt16(this.GetRegistryFolder().CreateSubKey("MoveSet").GetValue("iClearSpeed"));
            Text.MoveSet.iFrameTime = Convert.ToInt16(this.GetRegistryFolder().CreateSubKey("MoveSet").GetValue("iFrameTime"));

            this.WriteLog("节目建立成功");
            if (-1 == User_AddText(this.iCardNum, ref Text, this.iProgramIndex)) {
                this.WriteLog("添加文本失败");
            } else if (!User_SendToScreen(this.iCardNum)) {
                this.WriteLog("发送到屏幕失败");
            } else {
                this.WriteLog("发送到屏幕成功");
            }

            if (!User_RealtimeDisConnect(this.iCardNum)) {
                this.WriteLog("关闭连接失败");
            } else {
                this.WriteLog("关闭连接成功");
            }
        }

        protected void getTxtFromUrl() {
            //从注册表读取获得txt文件的地址
            String txtUrl = this.GetRegistryFolder().GetValue("sourceUrl").ToString();
            String downloadedText = this.HttpDownloadFile(txtUrl);
            if (downloadedText.Length > 0) {
                this.publishText = downloadedText;
            }
        }

        protected void getTxtFromFile() {
            //从注册表读取获得txt文件的路径
            String filePathString = this.GetRegistryFolder().GetValue("sourceFile").ToString();
            FileStream fileStream = new FileStream(filePathString, FileMode.Open, FileAccess.Read);
            StreamReader reader = new StreamReader(fileStream, Encoding.GetEncoding(this.GetRegistryFolder().GetValue("encode").ToString()));
            StringBuilder builder = new StringBuilder();
            while (reader.Peek() != -1) {
                builder.Append(reader.ReadLine() + "\r\n");
            }
            reader.Close();
            fileStream.Close();
            if (builder.Length > 0) {
                this.publishText = builder.ToString();
            }
        }

        protected void getWeatherFromUrl() {

        }

        public string HttpDownloadFile(string url) {
            // 设置参数
            HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;

            //发送请求并获取相应回应数据
            HttpWebResponse response = request.GetResponse() as HttpWebResponse;
            //直到request.GetResponse()程序才开始向目标网页发送Post请求
            Stream responseStream = response.GetResponseStream();
            //byte[] bArr = new byte[1024];
            //int size = responseStream.Read(bArr, 0, (int)bArr.Length);
            //responseStream.Close();
            //if (size > 0)
            //{
            //    string inputString = Encoding.GetEncoding(this.GetRegistryFolder().GetValue("encode").ToString()).GetString(bArr);
            //    return inputString;
            //}
            //else
            //{
            //    return "";
            //}
            string content = "";
            using (StreamReader sr = new StreamReader(responseStream, Encoding.GetEncoding(this.GetRegistryFolder().GetValue("encode").ToString())))
            {
                content = sr.ReadToEnd();
            }
            return content;
        }

        private void WriteLog(String message) {
            try {
                String fileName = DateTime.Now.Year.ToString() + "_" + DateTime.Now.Month.ToString() + "_" + DateTime.Now.Day.ToString();
                if (!File.Exists(System.Threading.Thread.GetDomain().BaseDirectory + "logs\\" + DateTime.Now.Year.ToString() + "\\" + fileName + ".log")) {
                    if (!Directory.Exists(System.Threading.Thread.GetDomain().BaseDirectory + "logs")) {
                        Directory.CreateDirectory(System.Threading.Thread.GetDomain().BaseDirectory + "logs");
                    }
                    if (!Directory.Exists(System.Threading.Thread.GetDomain().BaseDirectory + "logs\\" + DateTime.Now.Year.ToString())) {
                        Directory.CreateDirectory(System.Threading.Thread.GetDomain().BaseDirectory + "logs\\" + DateTime.Now.Year.ToString());
                    }
                    File.Create(System.Threading.Thread.GetDomain().BaseDirectory + "logs\\" + DateTime.Now.Year.ToString() + "\\" + fileName + ".log");
                    File.Create(System.Threading.Thread.GetDomain().BaseDirectory + "logs\\" + DateTime.Now.Year.ToString() + "\\" + fileName + ".log").Dispose();
                }
                using (System.IO.StreamWriter streamWriter = new System.IO.StreamWriter(System.Threading.Thread.GetDomain().BaseDirectory + "logs\\" + DateTime.Now.Year.ToString() + "\\" + fileName + ".log", true)) {
                    streamWriter.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " : " + message);
                }
            } catch (Exception e) {
                System.Console.WriteLine(e.Message);
            }
        }

        /// <summary>
        /// 检查注册表项
        /// \LOCALMACHINE\SOFTWARE\BingshanGuardian\LEDPublish 文件夹
        /// sourceType 0：从url获取txt；1：从file获取txt；其他：从API获取天气信息
        /// sourceUrl   从url获取txt的地址
        /// sourceFile  从file获取txt的文件路径
        /// iniFile EQ2008_Dll_set.ini文件路径
        /// publishInterval 发送间隔毫秒数
        /// encode  文本编码格式
        /// iCardNum 1
        /// iProgramIndex 0
        /// 子文件夹PartInfo
        /// FrameColor 0
        /// iFrameMode 0
        /// iHeight 64
        /// iWidth 320
        /// iX 0
        /// iY 0
        /// 子文件夹FontInfo
        /// bFontBold false
        /// bFontItaic false
        /// bFontUnderline false
        /// colorFont 0x00FF
        /// iFontSize 22
        /// strFontName 宋体
        /// iAlignStyle 1
        /// iRowSpace 1
        /// 子文件夹MoveSet
        /// bClear true
        /// iActionSpeed 5
        /// iActionType 5
        /// iHoldTime 100
        /// iClearActionType 5
        /// iClearSpeed 5
        /// iFrameTime 20
        /// </summary>
        private void CheckRegistry() {
            this.folder = this.GetRegistryFolder();
            this.CheckMakeRegistry(this.folder, "sourceType", "0");
            this.CheckMakeRegistry(this.folder, "sourceUrl", "http://10.0.0.171/moji/weather.txt");
            this.CheckMakeRegistry(this.folder, "sourceFile", ".\\weather.txt");
            this.CheckMakeRegistry(this.folder, "iniFile", ".\\EQ2008_Dll_Set.ini");
            this.CheckMakeRegistry(this.folder, "publishInterval", "600000");
            this.CheckMakeRegistry(this.folder, "encode", "utf-8");
            this.CheckMakeRegistry(this.folder, "iCardNum", "1"); 

            this.CheckMakeRegistry(this.folder, "BkColor", "0"); 
             RegistryKey partInfo = this.folder.CreateSubKey("PartInfo", true);
            this.CheckMakeRegistry(partInfo, "FrameColor", "0");
            this.CheckMakeRegistry(partInfo, "iFrameMode", "0");
            this.CheckMakeRegistry(partInfo, "iHeight", "64");
            this.CheckMakeRegistry(partInfo, "iWidth", "320");
            this.CheckMakeRegistry(partInfo, "iX", "0");
            this.CheckMakeRegistry(partInfo, "iY", "0");

            RegistryKey fontInfo = this.folder.CreateSubKey("FontInfo", true);
            this.CheckMakeRegistry(fontInfo, "bFontBold", "false");
            this.CheckMakeRegistry(fontInfo, "bFontItaic", "false");
            this.CheckMakeRegistry(fontInfo, "bFontUnderline", "false");
            this.CheckMakeRegistry(fontInfo, "colorFont", "255");
            this.CheckMakeRegistry(fontInfo, "iFontSize", "22");
            this.CheckMakeRegistry(fontInfo, "strFontName", "宋体");
            this.CheckMakeRegistry(fontInfo, "iAlignStyle", "1");
            this.CheckMakeRegistry(fontInfo, "iRowSpace", "1");

            RegistryKey moveSet = this.folder.CreateSubKey("MoveSet", true);
            this.CheckMakeRegistry(moveSet, "bClear", "true");
            this.CheckMakeRegistry(moveSet, "iActionSpeed", "5");
            this.CheckMakeRegistry(moveSet, "iActionType", "5");
            this.CheckMakeRegistry(moveSet, "iHoldTime", "100");
            this.CheckMakeRegistry(moveSet, "iClearActionType", "5");
            this.CheckMakeRegistry(moveSet, "iClearSpeed", "5");
            this.CheckMakeRegistry(moveSet, "iFrameTime", "20");

        }


        private void CheckMakeRegistry(RegistryKey key, String name, String value) {
            try {
                String regValue = key.GetValue(name).ToString();
                if (regValue.Trim().Length == 0) {
                    key.SetValue(name, value);
                    this.WriteLog("RegistryKey " + name + " is empty, set value : " + value);
                } else {
                    this.WriteLog("RegistryKey " + name + " found, value : " + regValue);
                }
            } catch (Exception e) {
                key.SetValue(name, value);
                this.WriteLog("RegistryKey " + name + " Not found, set value : " + value);
            }
        }

        /*
         * EQ2008-I/II控制卡动态库参数配置文件:
         *   1、控制卡地址"[地址：n]"和"CardAddress" 范围为：0~1023;
         *   2、控制卡类型"CardType"的取值为：
         *      EQ3002-I=0
         *      EQ3002-II=1
         *      EQ3002-III=2
         *      EQ2008-I/II=3
         *      EQ2010-I=4
         *      EQ2008-IE=5
         *      EQ2011=7
         *      EQ2012=8
         *      EQ2008-M=9
         *      EQ2013=21
         *      EQ2023=22
         *      EQ2033=23
         *   3、控制卡通讯模式“CommunicationMode”的取值为：
         *      串口通讯=0
         *      网路通讯=1
         *   4、显示屏的宽度和高度分别为“ScreemWidth”和“ScreemHeight”，取值为：
         *      ScreemWidth=8的倍数
         *   5、串口波特率和串口号分别为“SerialBaud”和“SerialNum”，取值为：
         *      SerialBaud=(9600，19200，57600，115200);
         *       (注：当CardType=EQ2013/EQ2023/EQ2033时,波特率只能为9600或57600)
         *      SerialNum =(1为COM1口，2为COM2口);
         *   6、网络端口号“NetPort”必须为5005;
         *   7、参数“IpAddressn”为IP地址：默认值为192.168.1.236
         *   8、ColorStyle:显示屏颜色类型:0--单色屏，1--双色屏。
         * 注意：
         *   *地址的个数可以根据实际显示屏的个数添加；
         *   *不要修改本文件的文件名及后缀；
         *   *本文件必须和应用程序放在同一个目录下。
         */
        private void CheckIniFile() {
            if (!File.Exists(System.Threading.Thread.GetDomain().BaseDirectory + "EQ2008_Dll_Set.ini")) {                
                FileStream fs = new FileStream(System.Threading.Thread.GetDomain().BaseDirectory + "EQ2008_Dll_Set.ini", FileMode.OpenOrCreate, FileAccess.ReadWrite); // 创建文件
                StreamWriter sw = new StreamWriter(fs, Encoding.GetEncoding("gb2312")); // 创建写入流
                sw.WriteLine("[地址：0]\r\n"
                    + "CardType=21\r\n"
                    + "CardAddress=0\r\n"
                    + "CommunicationMode=1\r\n"
                    + "ScreemHeight=64\r\n"
                    + "ScreemWidth=320\r\n"
                    + "SerialBaud=57600\r\n"
                    + "SerialNum=1\r\n"
                    + "NetPort=5005\r\n"
                    + "IpAddress0=10\r\n"
                    + "IpAddress1=1\r\n"
                    + "IpAddress2=175\r\n"
                    + "IpAddress3=100\r\n"
                    + "ColorStyle=0"
                    ); // 写入
                sw.Close(); //关闭文件
                this.WriteLog("重建EQ2008_Dll_Set.ini");
            }
        }

        private RegistryKey GetRegistryFolder() {
            return Registry.LocalMachine.CreateSubKey("SOFTWARE", true).CreateSubKey("BingshanGuardian", true).CreateSubKey("LEDPublish", true);
        }
    }
}
