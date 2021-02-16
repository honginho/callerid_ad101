﻿using System;
using System.Data.OleDb;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public struct AD101DEVICEPARAMETER
    {
        public int nRingOn;
        public int nRingOff;
        public int nHookOn;
        public int nHookOff;
        public int nStopCID;
        public int nNoLine;         // Add this parameter in new AD101(MCU Version is 6.0)
    }

    public partial class CallID : Form
    {
        #region
        [DllImport("AD101Device.dll", EntryPoint = "AD101_GetCallerID", CharSet = CharSet.Ansi)]
        public static extern int AD101_GetCallerID(int nLine, StringBuilder szCallerIDBuffer, StringBuilder szName, StringBuilder szTime);

        // Control led 
        [DllImport("AD101Device.dll", EntryPoint = "AD101_SetLED")]
        public static extern int AD101_SetLED(int nLine, int enumLed);

        [DllImport("AD101Device.dll", EntryPoint = "AD101_GetParameter")]
        public static extern int AD101_GetParameter(int nLine, ref AD101DEVICEPARAMETER tagParameter);

        [DllImport("AD101Device.dll", EntryPoint = "AD101_GetCPUVersion")]
        public static extern int AD101_GetCPUVersion(int nLine, StringBuilder szCPUVersion);

        [DllImport("AD101Device.dll", EntryPoint = "AD101_InitDevice")]
        public static extern int AD101_InitDevice(int hWnd);

        // Get talking time
        [DllImport("AD101Device.dll", EntryPoint = "AD101_GetTalkTime")]
        public static extern int AD101_GetTalkTime(int nLine);

        [DllImport("AD101Device.dll", EntryPoint = "AD101_GetCPUID")]
        public static extern int AD101_GetCPUID(int nLine, StringBuilder szCPUID);


        [DllImport("AD101Device.dll", EntryPoint = "AD101_ReadParameter")]
        public static extern int AD101_ReadParameter(int nLine);

        [DllImport("AD101Device.dll", EntryPoint = "AD101_GetDialDigit")]
        public static extern int AD101_GetDialDigit(int nLine, StringBuilder szDialDigitBuffer);

        public const int MCU_BACKCID = 0x09;		// Return Device CallerID
        public const int WM_USBLINEMSG = 1024 + 180;
        public const int MCU_BACKDIGIT = 0x0A;
        public const int MCU_BACKTALK = 0xBB;
        public const int MCU_BACKPARAM = 0x0C;
        public const int MCU_BACKID = 0x07;	// Return Device ID
        //public const int MCU_BACKSTATE = 0x08;	// Return Device State
        public const int MCU_BACKCPUID = 0x0D;	// Return Device CPU ID
        //public const int MCU_BACKPARAM = 0x0C;	// Return Device Paramter
        //public const int MCU_BACKDEVICE = 0x0B;	// Return Device Back Device ID
        public const int MCU_BACKDISABLE = 0xFF;    // Return Device Init
        // LED Status 
        enum LEDTYPE
        {
            LED_CLOSE = 1,
            LED_RED,
            LED_GREEN,
            LED_YELLOW,
            LED_REDSLOW,
            LED_GREENSLOW,
            LED_YELLOWSLOW,
            LED_REDQUICK,
            LED_GREENQUICK,
            LED_YELLOWQUICK,
        };
        //////////////////////////////////////////////////////////////////////////////////////////////

        // Line Status 
        enum ENUMLINEBUSY
        {
            LINEBUSY = 0,
            LINEFREE,
        };


        public const int HKONSTATEPRA = 0x01; // hook on pr+  HOOKON_PRA
        public const int HKONSTATEPRB = 0x02;  // hook on pr-  HOOKON_PRR
        public const int HKONSTATENOPR = 0x03;  // have pr  HAVE_PR
        public const int HKOFFSTATEPRA = 0x04;   // hook off pr+  HOOKOFF_PRA
        public const int HKOFFSTATEPRB = 0x05;  // hook off pr-  HOOKOFF_PRR
        public const int NO_LINE = 0x06; // no line  NULL_LINE
        public const int RINGONSTATE = 0x07;  // ring on  RING_ON
        public const int RINGOFFSTATE = 0x08;  // ring off RING_OFF
        public const int NOHKPRA = 0x09; // NOHOOKPRA= 0x09, // no hook pr+
        public const int NOHKPRB = 0x0a; // NOHOOKPRR= 0x0a, // no hook pr-
        public const int NOHKNOPR = 0x0b; // NOHOOKNPR= 0x0b, // no hook no pr

        //public const int WM_USBLINEMSG = 1024 + 180;
        #endregion

        public CallID()
        {
            InitializeComponent();
        }

        private void InsertCallerDetailsToDBF(int index, string phoneNumber, DateTime dialDatetime)
        {
            string DBSource = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"a03_tel_in.dbf");
            using (OleDbConnection dbfcon = new OleDbConnection("Provider=vfpoledb;Data Source=" + DBSource + ";Collating Sequence=machine;Mode=ReadWrite;"))
            {
                string query = @"INSERT INTO A03_tel_in(Datetime1, C_line, C_tel) VALUES(?,?,?)";
                OleDbCommand cmd = new OleDbCommand(query, dbfcon);
                cmd.Parameters.AddWithValue("@Datetime1", dialDatetime);
                cmd.Parameters.AddWithValue("@C_line", index.ToString());
                cmd.Parameters.AddWithValue("@C_tel", phoneNumber);

                dbfcon.Open();
                new OleDbCommand("set null off", dbfcon).ExecuteNonQuery();
                cmd.ExecuteNonQuery();
                dbfcon.Close();
            }
        }

        #region
        private void OnDeviceMsg(IntPtr wParam, IntPtr Lparam)
        {
            try
            {
                int nMsg = new int();
                int nLine = new int();

                nMsg = wParam.ToInt32() % 65536;
                nLine = wParam.ToInt32() / 65536;

                switch (nMsg)
                {
                    // Init
                    case MCU_BACKDISABLE:
                        listView1.Items[nLine].SubItems[1].Text = "-";
                        listView1.Items[nLine].SubItems[2].Text = "-";
                        listView1.Items[nLine].SubItems[3].Text = "-";
                        listView1.Items[nLine].SubItems[4].Text = "-";
                        listView1.Items[nLine].SubItems[5].Text = "-";
                        break;
                    case MCU_BACKID:
                        {
                            StringBuilder szCPUVersion = new StringBuilder(32);
                            listView1.Items[nLine].SubItems[1].Text = "啟用";
                            AD101_GetCPUVersion(nLine, szCPUVersion);
                            listView1.Items[nLine].SubItems[3].Text = szCPUVersion.ToString();
                        }
                        break;
                    case MCU_BACKCID:
                        {
                            StringBuilder szCallerID = new StringBuilder(128);
                            StringBuilder szName = new StringBuilder(128);
                            StringBuilder szTime = new StringBuilder(128);

                            int nLen = AD101_GetCallerID(nLine, szCallerID, szName, szTime);
                            string check = "";
                            check = szCallerID.ToString();

                            this.BeginInvoke((System.Windows.Forms.MethodInvoker)delegate ()
                            {
                                string phoneNumber = "";
                                DateTime dialDatetime = DateTime.Now;
                                string strDialDatetime = dialDatetime.ToString("yyyy/MM/dd HH:mm:ss");
                                if (check == null)
                                {
                                    phoneNumber = "國際電話號碼";
                                }
                                else
                                {
                                    string str = szCallerID.ToString();
                                    string newString = Regex.Replace(str, "[^.0-9]", "");
                                    phoneNumber = newString;
                                }
                                listView1.Items[nLine].SubItems[4].Text = phoneNumber;
                                listView1.Items[nLine].SubItems[5].Text = strDialDatetime;
                                InsertCallerDetailsToDBF(nLine, phoneNumber, dialDatetime);
                            });
                        }
                        break;
                    case MCU_BACKTALK:
                        {
                            string strTalk;
                            strTalk = string.Format("{0:D2}", Lparam) + "S";
                            //DeviceStateLabel0.Text = strTalk;
                            //listView1.Items[nLine].SubItems[7].Text = strTalk;
                        }
                        break;
                    case MCU_BACKPARAM:
                        {
                            AD101DEVICEPARAMETER tagParameter = new AD101DEVICEPARAMETER();
                            AD101_GetParameter(nLine, ref tagParameter);
                        }
                        break;
                    case MCU_BACKCPUID:
                        {
                            StringBuilder szCPUID = new StringBuilder(4);
                            AD101_GetCPUID(nLine, szCPUID);
                            listView1.Items[nLine].SubItems[2].Text = szCPUID.ToString();
                        }
                        break;
                    //case MCU_BACKDIGIT:
                    //    {
                    //        StringBuilder szDialDigit = new StringBuilder(128);
                    //        int nLen = AD101_GetDialDigit(nLine, szDialDigit);
                    //        //szDialDigit.ToString(); // get which number this machine dial
                    //    }
                    //    break;
                    //case MCU_BACKPARAM:
                    //    {
                    //        AD101DEVICEPARAMETER tagParameter = new AD101DEVICEPARAMETER();

                    //        AD101_GetParameter(nLine, ref tagParameter);
                    //    }
                    //    break;
                    //case MCU_BACKDEVICE:
                    //    {
                    //        StringBuilder szCPUVersion = new StringBuilder(32);

                    //        AD101_GetCPUVersion(nLine, szCPUVersion);


                    //    }
                    //    break;
                    default:
                        break;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }



        //one more function
        protected override void DefWndProc(ref System.Windows.Forms.Message m)
        {
            try
            {
                switch (m.Msg)
                {
                    case WM_USBLINEMSG:
                        OnDeviceMsg(m.WParam, m.LParam);
                        break;
                    default:
                        base.DefWndProc(ref m);
                        break;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }
        #endregion

        private void CallID_Load(object sender, EventArgs e)
        {
            listView1.Columns.Clear();
            listView1.Columns.Add("#", 30, HorizontalAlignment.Left);
            listView1.Columns.Add("機器狀態", 90, HorizontalAlignment.Left);
            listView1.Columns.Add("機器編號", 110, HorizontalAlignment.Left);
            listView1.Columns.Add("CPU 版本", 130, HorizontalAlignment.Left);
            listView1.Columns.Add("電話號碼", 130, HorizontalAlignment.Left);
            listView1.Columns.Add("來電時間", 200, HorizontalAlignment.Left);

            listView1.Items.Add("1");
            listView1.Items.Add("2");
            listView1.Items.Add("3");
            listView1.Items.Add("4");

            for (int i = 0; i < 4; i++)
            {
                listView1.Items[i].SubItems.Add("-");
                listView1.Items[i].SubItems.Add("-");
                listView1.Items[i].SubItems.Add("-");
                listView1.Items[i].SubItems.Add("-");
                listView1.Items[i].SubItems.Add("-");
                //listView1.Items[i].SubItems[1].Text = "停用";
            }

            try
            {
                if (AD101_InitDevice(Handle.ToInt32()) == 0)
                {
                    return;
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message + " " + ex.InnerException.Message); }
            AD101_SetLED(0, 3);
        }
    }
}