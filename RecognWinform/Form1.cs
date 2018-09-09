using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Speech.Recognition;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace RecognWinform
{
    public partial class Form1 : Form
    {

        private SpeechRecognitionEngine SRE = new SpeechRecognitionEngine();

        public Form1()
        {
            InitializeComponent();
        }

        public void Form1_Load(object sender, EventArgs e)
        {
            SRE.SetInputToDefaultAudioDevice();         //<======= 默认的语音输入设备，你可以设定为去识别一个WAV文件。
            GrammarBuilder GB = new GrammarBuilder();
            //GB.Append("选择");
            GB.Append(new Choices(new string[] { "红色","橙色", "黄色", "绿色", "青色", "蓝色", "紫色",  "往上翻", "上一页", "往下翻", "下一页" }));
            Grammar G = new Grammar(GB);
            G.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(G_SpeechRecognized);
            SRE.LoadGrammar(G);
            SRE.RecognizeAsync(RecognizeMode.Multiple); //<======= 异步调用识别引擎，允许多次识别（否则程序只响应你的一句话）

        }


        /// <summary>
        /// 定义一个代理
        /// </summary>
        private delegate void CrossThreadOperationControl();

        private void setLabelText(String sre_text)
        {
            CrossThreadOperationControl CrossDelete = delegate ()
            {
                label1.Text = sre_text;

            };
            label1.Invoke(CrossDelete);
        }

        private String getComboboxText()
        {
            String combo_str = "";
            CrossThreadOperationControl CrossDelete = delegate ()
            {
                combo_str = comboBox1.Text;

            };
            comboBox1.Invoke(CrossDelete);
            return combo_str;
        }

        private String getCombobox2Text()
        {
            String combo_str = "";
            CrossThreadOperationControl CrossDelete = delegate ()
            {
                combo_str = comboBox2.Text;

            };
            comboBox2.Invoke(CrossDelete);
            return combo_str;
        }

        private void setFormBcolor(Color color)
        {
            CrossThreadOperationControl CrossDelete = delegate ()
            {
                this.BackColor = color;

            };
            this.Invoke(CrossDelete);

        }


        //private void setFocus(){ }

        [DllImport("USER32.DLL", CharSet = CharSet.Auto)]
        private static extern int ShowWindow(System.IntPtr hWnd, int nCmdShow);

        [DllImport("USER32.DLL", CharSet = CharSet.Auto)]
        private static extern bool SetForegroundWindow(System.IntPtr hWnd);
        private const int SW_NORMAL = 1;



        [DllImport("user32.dll")]
        public static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);

        public const byte vbKeyLeft = 0x25;      // LEFT ARROW 键
        public const byte vbKeyUp = 0x26;        // UP ARROW 键
        public const byte vbKeyRight = 0x27;     // RIGHT ARROW 键
        public const byte vbKeyDown = 0x28;      // DOWN ARROW 键

        //判断是否存在名为processName的进程(程序)，存在的话使它获得焦点
        //例 if (CheckExistProcess("JobExec") == false)
        private static bool CheckExistProcess(string processName)
        {
            bool aRet = false;
            System.Diagnostics.Process[] arrProcess =
                System.Diagnostics.Process.GetProcessesByName(processName);
            if (arrProcess.Length > 0) { aRet = true; }

            if (aRet)
            {
                try
                {
                    foreach (System.Diagnostics.Process hProcess in arrProcess)
                    {
                        ShowWindow(hProcess.MainWindowHandle, SW_NORMAL);
                        SetForegroundWindow(hProcess.MainWindowHandle);
                        break;
                    }
                }
                catch { }
            }
            return aRet;
        }


        private void changePage(int speed, byte vbKey)
        {
            if (getComboboxText() == "福昕阅读器")
            {
                
                if (CheckExistProcess("FoxitReader") == false)
                {
                    Console.WriteLine("false");
                    setLabelText("福昕阅读器没打开");
                }
                else
                {
                    Console.WriteLine("true");

                    for (int i = 0; i < speed; i++)
                    {
                        //模拟按下键
                        keybd_event(vbKey, 0, 0, 0);
                        //松开按键
                        keybd_event(vbKey, 0, 2, 0);
                    }
                    
                }
                

            }
            else if (getComboboxText() == "360浏览器")
            {
                if (CheckExistProcess("360se") == false)
                {
                    Console.WriteLine("false");
                    setLabelText("360浏览器没打开");
                }
                else
                {
                    Console.WriteLine("true");

                    for (int i = 0; i < speed; i++)
                    {
                        //模拟按下键
                        keybd_event(vbKey, 0, 0, 0);
                        //松开按键
                        keybd_event(vbKey, 0, 2, 0);
                    }


                }
            }
            else
            {

            }
        }
        int speed = 0;
        void G_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            if (e.Result.Text.Length > 0)
            {
                setLabelText(e.Result.Text);    //设置lable文字

                switch (e.Result.Text)
                {
                    case "往下翻":
                        speed = Convert.ToInt32(getCombobox2Text());
                        changePage(speed, vbKeyDown);

                        break;
                    case "往上翻":
                        speed = Convert.ToInt32(getCombobox2Text());
                        changePage(speed, vbKeyUp);

                        break;
                    case "上一页":
                        speed = Convert.ToInt32(getCombobox2Text());
                        changePage(speed * 2, vbKeyUp);

                        break;
                    case "下一页":
                        speed = Convert.ToInt32(getCombobox2Text());
                        changePage(speed * 2, vbKeyDown);

                        break;


                    case "红色":
                        setFormBcolor(Color.Red);
                        break;
                    case "橙色":
                        setFormBcolor(Color.Orange);
                        break;
                    case "黄色":
                        setFormBcolor(Color.Yellow);
                        break;
                    case "绿色":
                        setFormBcolor(Color.Green);
                        break;
                    case "青色":
                        setFormBcolor(Color.Cyan);
                        break;
                    case "蓝色":
                        setFormBcolor(Color.Blue);
                        break;
                    case "紫色":
                        setFormBcolor(Color.Purple);
                        break;
                    
                        //default:
                        //    setLabelText("未识别！");
                        //    break;
                }


}
            
            
            
        }

        private void label1_Click(object sender, EventArgs e)
{

}

        private void button1_Click(object sender, EventArgs e)
        {
            Form2 fm2 = new Form2();
            fm2.Show();
        }
    }
}
