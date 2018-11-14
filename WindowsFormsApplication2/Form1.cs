using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace RS232_monitor
{
    public partial class FormMain : Form
    {
        private bool o_cd1, o_dsr1, o_dtr1, o_rts1, o_cts1;
        private bool o_cd2, o_dsr2, o_dtr2, o_rts2, o_cts2;
        private bool o_cd3, o_dsr3, o_dtr3, o_rts3, o_cts3;
        private bool o_cd4, o_dsr4, o_dtr4, o_rts4, o_cts4;
        public DataTable CSVdataTable = new DataTable("Logs");
        private List<string> portName = new List<string>();
        private string[] altPortName = new string[4];
        private bool[] displayhex = new bool[4] { false, false, false, false };
        private int LineBreakTimeout = 100;
        private int CSVLineNumberLimit = 0;
        private string CSVFileName = "";
        private int CSVLineCount = 0;
        private int LogLinesLimit = 100;
        private int LineLengthLimit = 200;
        private Logger datalog = new Logger();
        private Thread t;

        public static System.Timers.Timer aTimer;

        public FormMain()
        {
            InitializeComponent();
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.DoubleBuffer, true);
            aTimer = new System.Timers.Timer();
            aTimer.Interval = RS232_monitor2.Properties.Settings.Default.GUIRefreshPeriod;
            aTimer.Elapsed += CollectBuffer;
            aTimer.AutoReset = true;
            aTimer.Enabled = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            toolStripTextBox_CSVLinesNumber.LostFocus += ToolStripTextBox_CSVLinesNumber_Leave;
            LineBreakToolStripTextBox1.LostFocus += LineBreakToolStripTextBox1_Leave;

            dataGridView.DataSource = CSVdataTable;
            //create columns
            DataColumn colDate;
            colDate = new DataColumn("Date", typeof(string));
            DataColumn colTime;
            colTime = new DataColumn("Time", typeof(string));
            DataColumn colMilis;
            colMilis = new DataColumn("Milis", typeof(string));
            DataColumn colPort;
            colPort = new DataColumn("Port", typeof(string));
            DataColumn colDir;
            colDir = new DataColumn("Dir", typeof(string));
            DataColumn colData;
            colData = new DataColumn("Data", typeof(string));
            DataColumn colSig;
            colSig = new DataColumn("Signal", typeof(string));
            DataColumn colMark;
            colMark = new DataColumn("Mark", typeof(bool));
            //add columns to the table
            CSVdataTable.Columns.AddRange(new DataColumn[] { colDate, colTime, colMilis, colPort, colDir, colData, colSig, colMark });

            DataGridViewColumn column = dataGridView.Columns[0];
            //column.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            column.Resizable = DataGridViewTriState.True;
            column.MinimumWidth = 70;
            column.Width = 70;

            column = dataGridView.Columns[1];
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            column.Resizable = DataGridViewTriState.True;
            column.MinimumWidth = 55;
            column.Width = 55;

            column = dataGridView.Columns[2];
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            column.Resizable = DataGridViewTriState.True;
            column.MinimumWidth = 30;
            column.Width = 30;

            column = dataGridView.Columns[3];
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            column.Resizable = DataGridViewTriState.True;
            column.MinimumWidth = 30;
            column.Width = 40;

            column = dataGridView.Columns[4];
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            column.Resizable = DataGridViewTriState.True;
            column.MinimumWidth = 30;
            column.Width = 30;

            column = dataGridView.Columns[5];
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            column.Resizable = DataGridViewTriState.True;
            column.MinimumWidth = 200;
            column.Width = 250;

            column = dataGridView.Columns[6];
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            column.Resizable = DataGridViewTriState.True;
            column.MinimumWidth = 60;
            column.Width = 60;

            column = dataGridView.Columns[7];
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;
            column.Resizable = DataGridViewTriState.True;
            column.MinimumWidth = 30;
            column.Width = 30;

            //load settings
            textBox_command.Text = RS232_monitor2.Properties.Settings.Default.DefaultCommand;
            checkBox_commandhex.Checked = RS232_monitor2.Properties.Settings.Default.DefaultCommandHex;
            textBox_params.Text = RS232_monitor2.Properties.Settings.Default.DefaultParameter;
            checkBox_paramhex.Checked = RS232_monitor2.Properties.Settings.Default.DefaultParamHex;
            checkBox_cr.Checked = RS232_monitor2.Properties.Settings.Default.addCR;
            checkBox_lf.Checked = RS232_monitor2.Properties.Settings.Default.addLF;
            checkBox_suff.Checked = RS232_monitor2.Properties.Settings.Default.addSuff;
            textBox_suff.Text = RS232_monitor2.Properties.Settings.Default.SuffText;
            checkBox_suffhex.Checked = RS232_monitor2.Properties.Settings.Default.DefaultSuffHex;
            checkBox_insPin.Checked = RS232_monitor2.Properties.Settings.Default.LogSignal;
            checkBox_insTime.Checked = RS232_monitor2.Properties.Settings.Default.LogTime;
            checkBox_insDir.Checked = RS232_monitor2.Properties.Settings.Default.LogDir;
            checkBox_portName.Checked = RS232_monitor2.Properties.Settings.Default.LogPortName;
            checkBox_displayPort1hex.Checked = RS232_monitor2.Properties.Settings.Default.HexPort1;
            checkBox_displayPort2hex.Checked = RS232_monitor2.Properties.Settings.Default.HexPort2;
            checkBox_displayPort3hex.Checked = RS232_monitor2.Properties.Settings.Default.HexPort3;
            checkBox_displayPort4hex.Checked = RS232_monitor2.Properties.Settings.Default.HexPort4;
            textBox_port1Name.Text = RS232_monitor2.Properties.Settings.Default.Port1Name;
            textBox_port2Name.Text = RS232_monitor2.Properties.Settings.Default.Port2Name;
            textBox_port3Name.Text = RS232_monitor2.Properties.Settings.Default.Port3Name;
            textBox_port4Name.Text = RS232_monitor2.Properties.Settings.Default.Port4Name;
            logToGridToolStripMenuItem.Checked = RS232_monitor2.Properties.Settings.Default.LogGrid;
            autoscrollToolStripMenuItem.Checked = RS232_monitor2.Properties.Settings.Default.AutoScroll;
            lineWrapToolStripMenuItem.Checked = RS232_monitor2.Properties.Settings.Default.LineWrap;
            autosaveTXTToolStripMenuItem1.Checked = RS232_monitor2.Properties.Settings.Default.AutoLogTXT;
            terminaltxtToolStripMenuItem1.Text = RS232_monitor2.Properties.Settings.Default.TXTlogFile;
            autosaveCSVToolStripMenuItem1.Checked = RS232_monitor2.Properties.Settings.Default.AutoLogCSV;
            LineBreakToolStripTextBox1.Text = RS232_monitor2.Properties.Settings.Default.LineBreakTimeout.ToString();
            LineBreakTimeout = RS232_monitor2.Properties.Settings.Default.LineBreakTimeout;
            CSVLineNumberLimit = RS232_monitor2.Properties.Settings.Default.CSVMaxLineNumber;
            toolStripTextBox_CSVLinesNumber.Text = CSVLineNumberLimit.ToString();
            LogLinesLimit = RS232_monitor2.Properties.Settings.Default.LogLinesLimit;
            LineLengthLimit = RS232_monitor2.Properties.Settings.Default.LineLengthLimit;
            terminaltxtToolStripMenuItem1.Enabled = autosaveTXTToolStripMenuItem1.Checked;

            //set the codepage to COM-port
            serialPort1.Encoding = Encoding.GetEncoding(RS232_monitor2.Properties.Settings.Default.CodePage);
            serialPort2.Encoding = Encoding.GetEncoding(RS232_monitor2.Properties.Settings.Default.CodePage);
            serialPort3.Encoding = Encoding.GetEncoding(RS232_monitor2.Properties.Settings.Default.CodePage);
            serialPort4.Encoding = Encoding.GetEncoding(RS232_monitor2.Properties.Settings.Default.CodePage);
            SerialPopulate();
        }

        private void ToolStripTextBox_CSVLinesNumber_LostFocus(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        private void Button_openport_Click(object sender, EventArgs e)
        {
            datalog.Clear();
            altPortName[0] = comboBox_portname1.Text;
            altPortName[1] = comboBox_portname2.Text;
            altPortName[2] = comboBox_portname3.Text;
            altPortName[3] = comboBox_portname4.Text;
            portName.Clear();
            portName.Add("");
            portName.Add("");
            portName.Add("");
            portName.Add("");

            checkBox_DTR1.Checked = false;
            checkBox_DTR2.Checked = false;
            checkBox_DTR3.Checked = false;
            checkBox_DTR4.Checked = false;
            checkBox_RTS1.Checked = false;
            checkBox_RTS2.Checked = false;
            checkBox_RTS3.Checked = false;
            checkBox_RTS4.Checked = false;
            CSVFileName = DateTime.Today.ToShortDateString() + "_" + DateTime.Now.ToLongTimeString() + "_" + DateTime.Now.Millisecond.ToString("D3") + ".csv";
            CSVFileName = CSVFileName.Replace(':', '-').Replace('\\', '-').Replace('/', '-');
            CSVLineCount = 0;
            if (comboBox_portname1.SelectedIndex != 0)
            {
                comboBox_portname1.Enabled = false;
                comboBox_portspeed1.Enabled = false;
                comboBox_handshake1.Enabled = false;
                comboBox_databits1.Enabled = false;
                comboBox_parity1.Enabled = false;
                comboBox_stopbits1.Enabled = false;

                comboBox_portname2.Enabled = false;
                comboBox_portspeed2.Enabled = false;
                comboBox_handshake2.Enabled = false;
                comboBox_databits2.Enabled = false;
                comboBox_parity2.Enabled = false;
                comboBox_stopbits2.Enabled = false;

                comboBox_portname3.Enabled = false;
                comboBox_portspeed3.Enabled = false;
                comboBox_handshake3.Enabled = false;
                comboBox_databits3.Enabled = false;
                comboBox_parity3.Enabled = false;
                comboBox_stopbits3.Enabled = false;

                comboBox_portname4.Enabled = false;
                comboBox_portspeed4.Enabled = false;
                comboBox_handshake4.Enabled = false;
                comboBox_databits4.Enabled = false;
                comboBox_parity4.Enabled = false;
                comboBox_stopbits4.Enabled = false;

                serialPort1.PortName = comboBox_portname1.Text;
                serialPort1.BaudRate = Convert.ToInt32(comboBox_portspeed1.Text);
                serialPort1.DataBits = Convert.ToUInt16(comboBox_databits1.Text);
                serialPort1.Handshake = (Handshake)Enum.Parse(typeof(Handshake), comboBox_handshake1.Text);
                serialPort1.Parity = (Parity)Enum.Parse(typeof(Parity), comboBox_parity1.Text);
                serialPort1.StopBits = (StopBits)Enum.Parse(typeof(StopBits), comboBox_stopbits1.Text);
                serialPort1.ReadTimeout = RS232_monitor2.Properties.Settings.Default.ReceiveTimeOut;
                serialPort1.WriteTimeout = RS232_monitor2.Properties.Settings.Default.SendTimeOut;
                serialPort1.ReadBufferSize = 8192;
                try
                {
                    serialPort1.Open();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error opening port " + serialPort1.PortName + ": " + ex.Message);
                    comboBox_portname1.Enabled = true;
                    comboBox_portspeed1.Enabled = true;
                    comboBox_handshake1.Enabled = true;
                    comboBox_databits1.Enabled = true;
                    comboBox_parity1.Enabled = true;
                    comboBox_stopbits1.Enabled = true;

                    comboBox_portname2.Enabled = true;
                    comboBox_portspeed2.Enabled = true;
                    comboBox_handshake2.Enabled = true;
                    comboBox_databits2.Enabled = true;
                    comboBox_parity2.Enabled = true;
                    comboBox_stopbits2.Enabled = true;

                    comboBox_portname3.Enabled = true;
                    comboBox_portspeed3.Enabled = true;
                    comboBox_handshake3.Enabled = true;
                    comboBox_databits3.Enabled = true;
                    comboBox_parity3.Enabled = true;
                    comboBox_stopbits3.Enabled = true;

                    comboBox_portname4.Enabled = true;
                    comboBox_portspeed4.Enabled = true;
                    comboBox_handshake4.Enabled = true;
                    comboBox_databits4.Enabled = true;
                    comboBox_parity4.Enabled = true;
                    comboBox_stopbits4.Enabled = true;

                    return;
                }
                if (checkBox_insPin.Checked) serialPort1.PinChanged += SerialPort_PinChanged;
                serialPort1.DataReceived += SerialPort_DataReceived;
                portName[0] = serialPort1.PortName;
                button_refresh.Enabled = false;
                button_closeport.Enabled = true;
                button_openport.Enabled = false;
                o_cd1 = serialPort1.CDHolding;
                checkBox_CD1.Checked = o_cd1;
                o_dsr1 = serialPort1.DsrHolding;
                checkBox_DSR1.Checked = o_dsr1;
                o_dtr1 = serialPort1.DtrEnable;
                checkBox_DTR1.Checked = o_dtr1;
                o_cts1 = serialPort1.CtsHolding;
                checkBox_CTS1.Checked = o_cts1;
                checkBox_DTR1.Enabled = true;

                if (serialPort1.Handshake == Handshake.RequestToSend || serialPort1.Handshake == Handshake.RequestToSendXOnXOff)
                {
                    checkBox_RTS1.Enabled = false;
                }
                else
                {
                    o_rts1 = serialPort1.RtsEnable;
                    checkBox_RTS1.Checked = o_rts1;
                    checkBox_RTS1.Enabled = true;
                }
            }
            if (comboBox_portname2.SelectedIndex != 0)
            {
                comboBox_portname1.Enabled = false;
                comboBox_portspeed1.Enabled = false;
                comboBox_handshake1.Enabled = false;
                comboBox_databits1.Enabled = false;
                comboBox_parity1.Enabled = false;
                comboBox_stopbits1.Enabled = false;

                comboBox_portname2.Enabled = false;
                comboBox_portspeed2.Enabled = false;
                comboBox_handshake2.Enabled = false;
                comboBox_databits2.Enabled = false;
                comboBox_parity2.Enabled = false;
                comboBox_stopbits2.Enabled = false;

                comboBox_portname3.Enabled = false;
                comboBox_portspeed3.Enabled = false;
                comboBox_handshake3.Enabled = false;
                comboBox_databits3.Enabled = false;
                comboBox_parity3.Enabled = false;
                comboBox_stopbits3.Enabled = false;

                comboBox_portname4.Enabled = false;
                comboBox_portspeed4.Enabled = false;
                comboBox_handshake4.Enabled = false;
                comboBox_databits4.Enabled = false;
                comboBox_parity4.Enabled = false;
                comboBox_stopbits4.Enabled = false;

                serialPort2.PortName = comboBox_portname2.Text;
                serialPort2.BaudRate = Convert.ToInt32(comboBox_portspeed2.Text);
                serialPort2.DataBits = Convert.ToUInt16(serialPort2.DataBits);
                serialPort2.Handshake = (Handshake)Enum.Parse(typeof(Handshake), comboBox_handshake2.Text);
                serialPort2.Parity = (Parity)Enum.Parse(typeof(Parity), comboBox_parity2.Text);
                serialPort2.StopBits = (StopBits)Enum.Parse(typeof(StopBits), comboBox_stopbits2.Text);
                serialPort2.ReadTimeout = RS232_monitor2.Properties.Settings.Default.ReceiveTimeOut;
                serialPort2.WriteTimeout = RS232_monitor2.Properties.Settings.Default.SendTimeOut;
                serialPort2.ReadBufferSize = 8192;
                try
                {
                    serialPort2.Open();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error opening port " + serialPort2.PortName + ": " + ex.Message);
                    comboBox_portname1.Enabled = true;
                    comboBox_portspeed1.Enabled = true;
                    comboBox_handshake1.Enabled = true;
                    comboBox_databits1.Enabled = true;
                    comboBox_parity1.Enabled = true;
                    comboBox_stopbits1.Enabled = true;

                    comboBox_portname2.Enabled = true;
                    comboBox_portspeed2.Enabled = true;
                    comboBox_handshake2.Enabled = true;
                    comboBox_databits2.Enabled = true;
                    comboBox_parity2.Enabled = true;
                    comboBox_stopbits2.Enabled = true;

                    comboBox_portname3.Enabled = true;
                    comboBox_portspeed3.Enabled = true;
                    comboBox_handshake3.Enabled = true;
                    comboBox_databits3.Enabled = true;
                    comboBox_parity3.Enabled = true;
                    comboBox_stopbits3.Enabled = true;

                    comboBox_portname4.Enabled = true;
                    comboBox_portspeed4.Enabled = true;
                    comboBox_handshake4.Enabled = true;
                    comboBox_databits4.Enabled = true;
                    comboBox_parity4.Enabled = true;
                    comboBox_stopbits4.Enabled = true;
                    return;
                }
                if (checkBox_insPin.Checked) serialPort2.PinChanged += SerialPort_PinChanged;
                serialPort2.DataReceived += SerialPort_DataReceived;
                portName[1] = serialPort2.PortName;
                button_refresh.Enabled = false;
                button_closeport.Enabled = true;
                button_openport.Enabled = false;
                o_cd2 = serialPort2.CDHolding;
                checkBox_CD2.Checked = o_cd2;
                o_dsr2 = serialPort2.DsrHolding;
                checkBox_DSR2.Checked = o_dsr2;
                o_dtr2 = serialPort2.DtrEnable;
                checkBox_DTR2.Checked = o_dtr2;
                o_cts2 = serialPort2.CtsHolding;
                checkBox_CTS2.Checked = o_cts2;
                checkBox_DTR2.Enabled = true;
                if (serialPort2.Handshake == Handshake.RequestToSend || serialPort2.Handshake == Handshake.RequestToSendXOnXOff)
                {
                    checkBox_RTS2.Enabled = false;
                }
                else
                {
                    o_rts2 = serialPort2.RtsEnable;
                    checkBox_RTS2.Checked = o_rts2;
                    checkBox_RTS2.Enabled = true;
                }
            }
            if (comboBox_portname3.SelectedIndex != 0)
            {
                comboBox_portname1.Enabled = false;
                comboBox_portspeed1.Enabled = false;
                comboBox_handshake1.Enabled = false;
                comboBox_databits1.Enabled = false;
                comboBox_parity1.Enabled = false;
                comboBox_stopbits1.Enabled = false;

                comboBox_portname2.Enabled = false;
                comboBox_portspeed2.Enabled = false;
                comboBox_handshake2.Enabled = false;
                comboBox_databits2.Enabled = false;
                comboBox_parity2.Enabled = false;
                comboBox_stopbits2.Enabled = false;

                comboBox_portname3.Enabled = false;
                comboBox_portspeed3.Enabled = false;
                comboBox_handshake3.Enabled = false;
                comboBox_databits3.Enabled = false;
                comboBox_parity3.Enabled = false;
                comboBox_stopbits3.Enabled = false;

                comboBox_portname4.Enabled = false;
                comboBox_portspeed4.Enabled = false;
                comboBox_handshake4.Enabled = false;
                comboBox_databits4.Enabled = false;
                comboBox_parity4.Enabled = false;
                comboBox_stopbits4.Enabled = false;

                serialPort3.PortName = comboBox_portname3.Text;
                serialPort3.BaudRate = Convert.ToInt32(comboBox_portspeed3.Text);
                serialPort3.DataBits = Convert.ToUInt16(serialPort3.DataBits);
                serialPort3.Handshake = (Handshake)Enum.Parse(typeof(Handshake), comboBox_handshake3.Text);
                serialPort3.Parity = (Parity)Enum.Parse(typeof(Parity), comboBox_parity3.Text);
                serialPort3.StopBits = (StopBits)Enum.Parse(typeof(StopBits), comboBox_stopbits3.Text);
                serialPort3.ReadTimeout = RS232_monitor2.Properties.Settings.Default.ReceiveTimeOut;
                serialPort3.WriteTimeout = RS232_monitor2.Properties.Settings.Default.SendTimeOut;
                serialPort3.ReadBufferSize = 8192;
                try
                {
                    serialPort3.Open();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error opening port " + serialPort3.PortName + ": " + ex.Message);
                    comboBox_portname1.Enabled = true;
                    comboBox_portspeed1.Enabled = true;
                    comboBox_handshake1.Enabled = true;
                    comboBox_databits1.Enabled = true;
                    comboBox_parity1.Enabled = true;
                    comboBox_stopbits1.Enabled = true;

                    comboBox_portname2.Enabled = true;
                    comboBox_portspeed2.Enabled = true;
                    comboBox_handshake2.Enabled = true;
                    comboBox_databits2.Enabled = true;
                    comboBox_parity2.Enabled = true;
                    comboBox_stopbits2.Enabled = true;

                    comboBox_portname3.Enabled = true;
                    comboBox_portspeed3.Enabled = true;
                    comboBox_handshake3.Enabled = true;
                    comboBox_databits3.Enabled = true;
                    comboBox_parity3.Enabled = true;
                    comboBox_stopbits3.Enabled = true;

                    comboBox_portname4.Enabled = true;
                    comboBox_portspeed4.Enabled = true;
                    comboBox_handshake4.Enabled = true;
                    comboBox_databits4.Enabled = true;
                    comboBox_parity4.Enabled = true;
                    comboBox_stopbits4.Enabled = true;
                    return;
                }
                if (checkBox_insPin.Checked) serialPort3.PinChanged += SerialPort_PinChanged;
                serialPort3.DataReceived += SerialPort_DataReceived;
                portName[2] = serialPort3.PortName;
                button_refresh.Enabled = false;
                button_closeport.Enabled = true;
                button_openport.Enabled = false;
                o_cd3 = serialPort3.CDHolding;
                checkBox_CD3.Checked = o_cd3;
                o_dsr3 = serialPort3.DsrHolding;
                checkBox_DSR3.Checked = o_dsr3;
                o_dtr3 = serialPort3.DtrEnable;
                checkBox_DTR3.Checked = o_dtr3;
                o_cts3 = serialPort3.CtsHolding;
                checkBox_CTS3.Checked = o_cts3;
                checkBox_DTR3.Enabled = true;
                if (serialPort3.Handshake == Handshake.RequestToSend || serialPort3.Handshake == Handshake.RequestToSendXOnXOff)
                {
                    checkBox_RTS3.Enabled = false;
                }
                else
                {
                    o_rts3 = serialPort3.RtsEnable;
                    checkBox_RTS3.Checked = o_rts3;
                    checkBox_RTS3.Enabled = true;
                }
            }
            if (comboBox_portname4.SelectedIndex != 0)
            {
                comboBox_portname1.Enabled = false;
                comboBox_portspeed1.Enabled = false;
                comboBox_handshake1.Enabled = false;
                comboBox_databits1.Enabled = false;
                comboBox_parity1.Enabled = false;
                comboBox_stopbits1.Enabled = false;

                comboBox_portname2.Enabled = false;
                comboBox_portspeed2.Enabled = false;
                comboBox_handshake2.Enabled = false;
                comboBox_databits2.Enabled = false;
                comboBox_parity2.Enabled = false;
                comboBox_stopbits2.Enabled = false;

                comboBox_portname3.Enabled = false;
                comboBox_portspeed3.Enabled = false;
                comboBox_handshake3.Enabled = false;
                comboBox_databits3.Enabled = false;
                comboBox_parity3.Enabled = false;
                comboBox_stopbits3.Enabled = false;

                comboBox_portname4.Enabled = false;
                comboBox_portspeed4.Enabled = false;
                comboBox_handshake4.Enabled = false;
                comboBox_databits4.Enabled = false;
                comboBox_parity4.Enabled = false;
                comboBox_stopbits4.Enabled = false;

                serialPort4.PortName = comboBox_portname4.Text;
                serialPort4.BaudRate = Convert.ToInt32(comboBox_portspeed4.Text);
                serialPort4.DataBits = Convert.ToUInt16(serialPort4.DataBits);
                serialPort4.Handshake = (Handshake)Enum.Parse(typeof(Handshake), comboBox_handshake4.Text);
                serialPort4.Parity = (Parity)Enum.Parse(typeof(Parity), comboBox_parity4.Text);
                serialPort4.StopBits = (StopBits)Enum.Parse(typeof(StopBits), comboBox_stopbits4.Text);
                serialPort4.ReadTimeout = RS232_monitor2.Properties.Settings.Default.ReceiveTimeOut;
                serialPort4.WriteTimeout = RS232_monitor2.Properties.Settings.Default.SendTimeOut;
                serialPort4.ReadBufferSize = 8192;
                try
                {
                    serialPort4.Open();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error opening port " + serialPort4.PortName + ": " + ex.Message);
                    comboBox_portname1.Enabled = true;
                    comboBox_portspeed1.Enabled = true;
                    comboBox_handshake1.Enabled = true;
                    comboBox_databits1.Enabled = true;
                    comboBox_parity1.Enabled = true;
                    comboBox_stopbits1.Enabled = true;

                    comboBox_portname2.Enabled = true;
                    comboBox_portspeed2.Enabled = true;
                    comboBox_handshake2.Enabled = true;
                    comboBox_databits2.Enabled = true;
                    comboBox_parity2.Enabled = true;
                    comboBox_stopbits2.Enabled = true;

                    comboBox_portname3.Enabled = true;
                    comboBox_portspeed3.Enabled = true;
                    comboBox_handshake3.Enabled = true;
                    comboBox_databits3.Enabled = true;
                    comboBox_parity3.Enabled = true;
                    comboBox_stopbits3.Enabled = true;

                    comboBox_portname4.Enabled = true;
                    comboBox_portspeed4.Enabled = true;
                    comboBox_handshake4.Enabled = true;
                    comboBox_databits4.Enabled = true;
                    comboBox_parity4.Enabled = true;
                    comboBox_stopbits4.Enabled = true;
                    return;
                }
                if (checkBox_insPin.Checked) serialPort4.PinChanged += SerialPort_PinChanged;
                serialPort4.DataReceived += SerialPort_DataReceived;
                portName[3] = serialPort4.PortName;
                button_refresh.Enabled = false;
                button_closeport.Enabled = true;
                button_openport.Enabled = false;
                o_cd4 = serialPort4.CDHolding;
                checkBox_CD4.Checked = o_cd4;
                o_dsr4 = serialPort4.DsrHolding;
                checkBox_DSR4.Checked = o_dsr4;
                o_dtr4 = serialPort4.DtrEnable;
                checkBox_DTR4.Checked = o_dtr4;
                o_cts4 = serialPort4.CtsHolding;
                checkBox_CTS4.Checked = o_cts4;
                checkBox_DTR4.Enabled = true;
                if (serialPort4.Handshake == Handshake.RequestToSend || serialPort4.Handshake == Handshake.RequestToSendXOnXOff)
                {
                    checkBox_RTS4.Enabled = false;
                }
                else
                {
                    o_rts4 = serialPort4.RtsEnable;
                    checkBox_RTS4.Checked = o_rts4;
                    checkBox_RTS4.Enabled = true;
                }
            }

            toolStripStatusLabel1.Text = "Idle...";
            toolStripStatusLabel1.BackColor = Color.White;

            if (checkBox_sendPort1.Checked == false && checkBox_sendPort2.Checked == false && checkBox_sendPort3.Checked == false && checkBox_sendPort4.Checked == false) button_send.Enabled = false;
            else if (serialPort1.IsOpen == true || serialPort2.IsOpen == true || serialPort3.IsOpen == true || serialPort4.IsOpen == true) button_send.Enabled = true;
            CheckBox_portName_CheckedChanged(this, EventArgs.Empty);
            aTimer.Interval = RS232_monitor2.Properties.Settings.Default.GUIRefreshPeriod;
            aTimer.Elapsed += CollectBuffer;
            aTimer.AutoReset = true;
            aTimer.Enabled = true;
            aTimer.Start();
        }

        private void Button_closeport_Click(object sender, EventArgs e)
        {
            if (t != null && t.IsAlive)
            {
                MessageBox.Show("Please wait till the data processing is finished.");
                return;
            }
            toolStripStatusLabel1.Text = "Idle...";

            aTimer.Elapsed -= CollectBuffer;
            aTimer.AutoReset = false;
            aTimer.Enabled = false;
            aTimer.Stop();

            toolStripStatusLabel1.Text = "";
            toolStripStatusLabel1.BackColor = Color.White;

            try
            {
                serialPort1.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error closing port " + serialPort1.PortName + ": " + ex.Message);
            }
            try
            {
                serialPort2.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error closing port " + serialPort2.PortName + ": " + ex.Message);
            }
            try
            {
                serialPort3.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error closing port " + serialPort3.PortName + ": " + ex.Message);
            }
            try
            {
                serialPort4.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error closing port " + serialPort4.PortName + ": " + ex.Message);
            }
            //flush logs base
            datalog.Clear();
            serialPort1.DataReceived -= SerialPort_DataReceived;
            serialPort1.PinChanged -= SerialPort_PinChanged;
            serialPort2.DataReceived -= SerialPort_DataReceived;
            serialPort2.PinChanged -= SerialPort_PinChanged;
            serialPort3.DataReceived -= SerialPort_DataReceived;
            serialPort3.PinChanged -= SerialPort_PinChanged;
            serialPort4.DataReceived -= SerialPort_DataReceived;
            serialPort4.PinChanged -= SerialPort_PinChanged;

            comboBox_portname1.Enabled = true;
            comboBox_portspeed1.Enabled = true;
            comboBox_handshake1.Enabled = true;
            comboBox_databits1.Enabled = true;
            comboBox_parity1.Enabled = true;
            comboBox_stopbits1.Enabled = true;

            comboBox_portname2.Enabled = true;
            comboBox_portspeed2.Enabled = true;
            comboBox_handshake2.Enabled = true;
            comboBox_databits2.Enabled = true;
            comboBox_parity2.Enabled = true;
            comboBox_stopbits2.Enabled = true;

            comboBox_portname3.Enabled = true;
            comboBox_portspeed3.Enabled = true;
            comboBox_handshake3.Enabled = true;
            comboBox_databits3.Enabled = true;
            comboBox_parity3.Enabled = true;
            comboBox_stopbits3.Enabled = true;

            comboBox_portname4.Enabled = true;
            comboBox_portspeed4.Enabled = true;
            comboBox_handshake4.Enabled = true;
            comboBox_databits4.Enabled = true;
            comboBox_parity4.Enabled = true;
            comboBox_stopbits4.Enabled = true;

            button_send.Enabled = false;
            button_refresh.Enabled = true;
            button_openport.Enabled = true;
            button_closeport.Enabled = false;

            checkBox_DTR1.Enabled = false;
            checkBox_RTS1.Enabled = false;

            checkBox_DTR2.Enabled = false;
            checkBox_RTS2.Enabled = false;

            checkBox_DTR3.Enabled = false;
            checkBox_RTS3.Enabled = false;

            checkBox_DTR4.Enabled = false;
            checkBox_RTS4.Enabled = false;
        }

        private void Button_send_Click(object sender, EventArgs e)
        {
            DateTime time = DateTime.Now;
            byte[] data = Accessory.ConvertHexToByteArray(textBox_senddata.Text);
            if (textBox_senddata.Text != "")
            {
                if (checkBox_sendPort1.Checked == true && serialPort1.IsOpen)
                {
                    try
                    {
                        serialPort1.BaseStream.Write(data, 0, data.Length);
                        time = DateTime.Now;
                    }
                    catch (Exception ex)
                    {
                        datalog.Add(serialPort1.PortName, Logger.DirectionType.Error, DateTime.Now, null, "Error writing port: " + ex.Message, checkBox_Mark.Checked);
                    }
                    datalog.Add(serialPort1.PortName, Logger.DirectionType.DataOut, time, data, "", checkBox_Mark.Checked);
                }
                if (checkBox_sendPort2.Checked == true && serialPort2.IsOpen)
                {
                    try
                    {
                        serialPort2.BaseStream.Write(data, 0, data.Length);
                        time = DateTime.Now;
                    }
                    catch (Exception ex)
                    {
                        datalog.Add(serialPort2.PortName, Logger.DirectionType.Error, DateTime.Now, null, "Error writing port: " + ex.Message, checkBox_Mark.Checked);
                    }
                    datalog.Add(serialPort2.PortName, Logger.DirectionType.DataOut, time, data, "", checkBox_Mark.Checked);
                }
                if (checkBox_sendPort3.Checked == true && serialPort3.IsOpen)
                {
                    try
                    {
                        serialPort3.BaseStream.Write(data, 0, data.Length);
                        time = DateTime.Now;
                    }
                    catch (Exception ex)
                    {
                        datalog.Add(serialPort3.PortName, Logger.DirectionType.Error, DateTime.Now, null, "Error writing port: " + ex.Message, checkBox_Mark.Checked);
                    }
                    datalog.Add(serialPort3.PortName, Logger.DirectionType.DataOut, time, data, "", checkBox_Mark.Checked);
                }
                if (checkBox_sendPort4.Checked == true && serialPort4.IsOpen)
                {
                    try
                    {
                        serialPort4.BaseStream.Write(data, 0, data.Length);
                        time = DateTime.Now;
                    }
                    catch (Exception ex)
                    {
                        datalog.Add(serialPort4.PortName, Logger.DirectionType.Error, DateTime.Now, null, "Error writing port: " + ex.Message, checkBox_Mark.Checked);
                    }
                    datalog.Add(serialPort4.PortName, Logger.DirectionType.DataOut, time, data, "", checkBox_Mark.Checked);
                }
            }
        }

        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            DateTime time = DateTime.Now;
            var serialPort = (SerialPort)sender;
            List<byte> rx = new List<byte>();
            try
            {
                while (serialPort.BytesToRead > 0)
                {
                    rx.Add((byte)serialPort.BaseStream.ReadByte());
                }
            }
            catch (Exception ex)
            {
                datalog.Add(serialPort.PortName, Logger.DirectionType.Error, DateTime.Now, null, "Error reading port: " + ex.Message, checkBox_Mark.Checked);
            }

            datalog.Add(serialPort.PortName, Logger.DirectionType.DataIn, time, rx.ToArray(), "", checkBox_Mark.Checked);
        }

        private void SerialPort_PinChanged(object sender, SerialPinChangedEventArgs e)
        {
            DateTime time = DateTime.Now;
            var serialPort = (SerialPort)sender;

            string outStr = "";
            if (serialPort.CDHolding == true && o_cd1 == false)
            {
                o_cd1 = true;
                outStr += "<DCD^>";
            }
            else if (serialPort.CDHolding == false && o_cd1 == true)
            {
                o_cd1 = false;
                outStr += "<DCDv>";
            }

            if (serialPort.DsrHolding == true && o_dsr1 == false)
            {
                o_dsr1 = true;
                outStr += "<DSR^>";
            }
            else if (serialPort.DsrHolding == false && o_dsr1 == true)
            {
                o_dsr1 = false;
                outStr += "<DSRv>";
            }

            if (serialPort.CtsHolding == true && o_cts1 == false)
            {
                o_cts1 = true;
                outStr += "<CTS^>";
            }
            else if (serialPort.CtsHolding == false && o_cts1 == true)
            {
                o_cts1 = false;
                outStr += "<CTSv>";
            }

            if (e.EventType.Equals(SerialPinChange.Ring))
            {
                outStr += "<RING>";
            }

            datalog.Add(serialPort.PortName, Logger.DirectionType.SignalIn, time, null, outStr, checkBox_Mark.Checked);
            int portNum = portName.IndexOf(serialPort.PortName);
            SetPinCD(portNum, serialPort.CDHolding);
            SetPinDSR(portNum, serialPort.DsrHolding);
            SetPinCTS(portNum, serialPort.CtsHolding);
            SetPinRING(portNum, true);
        }

        private void SerialPort_ErrorReceived(object sender, SerialErrorReceivedEventArgs e)
        {
            DateTime time = DateTime.Now;
            var serialPort = (SerialPort)sender;
            datalog.Add(serialPort.PortName, Logger.DirectionType.Error, time, null, e.EventType.ToString(), checkBox_Mark.Checked);
        }

        private void SetPinCD(int portNum, bool setPin)
        {
            this.Invoke((MethodInvoker)delegate
            {
                switch (portNum)
                {
                    case 1:
                        checkBox_CD1.Checked = setPin;
                        break;
                    case 2:
                        checkBox_CD2.Checked = setPin;
                        break;
                    case 3:
                        checkBox_CD3.Checked = setPin;
                        break;
                    case 4:
                        checkBox_CD4.Checked = setPin;
                        break;
                }
            });
        }
        private void SetPinDSR(int portNum, bool setPin)
        {
            this.Invoke((MethodInvoker)delegate
            {
                switch (portNum)
                {
                    case 1:
                        checkBox_DSR1.Checked = setPin;
                        break;
                    case 2:
                        checkBox_DSR2.Checked = setPin;
                        break;
                    case 3:
                        checkBox_DSR3.Checked = setPin;
                        break;
                    case 4:
                        checkBox_DSR4.Checked = setPin;
                        break;
                }
            });
        }
        private void SetPinCTS(int portNum, bool setPin)
        {
            this.Invoke((MethodInvoker)delegate
            {
                switch (portNum)
                {
                    case 1:
                        checkBox_CTS1.Checked = setPin;
                        break;
                    case 2:
                        checkBox_CTS2.Checked = setPin;
                        break;
                    case 3:
                        checkBox_CTS3.Checked = setPin;
                        break;
                    case 4:
                        checkBox_CTS4.Checked = setPin;
                        break;
                }
            });
        }
        private void SetPinRING(int portNum, bool setPin)
        {
            this.Invoke((MethodInvoker)delegate
            {
                switch (portNum)
                {
                    case 1:
                        checkBox_RI1.Checked = setPin;
                        break;
                    case 2:
                        checkBox_RI2.Checked = setPin;
                        break;
                    case 3:
                        checkBox_RI3.Checked = setPin;
                        break;
                    case 4:
                        checkBox_RI4.Checked = setPin;
                        break;
                }
            });
        }

        private void CheckBox_DTR1_CheckedChanged(object sender, EventArgs e)
        {
            DateTime time = DateTime.Now;
            try
            {
                serialPort1.DtrEnable = checkBox_DTR1.Checked;
                time = DateTime.Now;
            }
            catch (Exception ex)
            {
                datalog.Add(serialPort1.PortName, Logger.DirectionType.Error, DateTime.Now, null, "Error setting DTR: " + ex.Message, checkBox_Mark.Checked);
            }
            string outStr = "";
            if (serialPort1.DtrEnable == true && o_dtr1 == false)
            {
                o_dtr1 = true;
                outStr += "<DTR^>";
            }
            else if (serialPort1.DtrEnable == false && o_dtr1 == true)
            {
                o_dtr1 = false;
                outStr += "<DTRv>";
            }
            datalog.Add(serialPort1.PortName, Logger.DirectionType.SignalOut, time, null, outStr, checkBox_Mark.Checked);
        }

        private void CheckBox_DTR2_CheckedChanged(object sender, EventArgs e)
        {
            DateTime time = DateTime.Now;
            try
            {
                serialPort2.DtrEnable = checkBox_DTR1.Checked;
                time = DateTime.Now;
            }
            catch (Exception ex)
            {
                datalog.Add(serialPort2.PortName, Logger.DirectionType.Error, DateTime.Now, null, "Error setting DTR: " + ex.Message, checkBox_Mark.Checked);
            }
            string outStr = "";
            if (serialPort2.DtrEnable == true && o_dtr2 == false)
            {
                o_dtr2 = true;
                outStr += "<DTR^>";
            }
            else if (serialPort2.DtrEnable == false && o_dtr2 == true)
            {
                o_dtr2 = false;
                outStr += "<DTRv>";
            }
            datalog.Add(serialPort2.PortName, Logger.DirectionType.SignalOut, time, null, outStr, checkBox_Mark.Checked);
        }

        private void CheckBox_DTR3_CheckedChanged(object sender, EventArgs e)
        {
            DateTime time = DateTime.Now;
            try
            {
                serialPort3.DtrEnable = checkBox_DTR1.Checked;
                time = DateTime.Now;
            }
            catch (Exception ex)
            {
                datalog.Add(serialPort3.PortName, Logger.DirectionType.Error, DateTime.Now, null, "Error setting DTR: " + ex.Message, checkBox_Mark.Checked);
            }
            string outStr = "";
            if (serialPort3.DtrEnable == true && o_dtr3 == false)
            {
                o_dtr3 = true;
                outStr += "<DTR^>";
            }
            else if (serialPort3.DtrEnable == false && o_dtr3 == true)
            {
                o_dtr3 = false;
                outStr += "<DTRv>";
            }
            datalog.Add(serialPort3.PortName, Logger.DirectionType.SignalOut, time, null, outStr, checkBox_Mark.Checked);
        }

        private void CheckBox_DTR4_CheckedChanged(object sender, EventArgs e)
        {
            DateTime time = DateTime.Now;
            try
            {
                serialPort4.DtrEnable = checkBox_DTR1.Checked;
                time = DateTime.Now;
            }
            catch (Exception ex)
            {
                datalog.Add(serialPort4.PortName, Logger.DirectionType.Error, DateTime.Now, null, "Error setting DTR: " + ex.Message, checkBox_Mark.Checked);
            }
            string outStr = "";
            if (serialPort4.DtrEnable == true && o_dtr4 == false)
            {
                o_dtr4 = true;
                outStr += "<DTR^>";
            }
            else if (serialPort4.DtrEnable == false && o_dtr4 == true)
            {
                o_dtr4 = false;
                outStr += "<DTRv>";
            }
            datalog.Add(serialPort4.PortName, Logger.DirectionType.SignalOut, time, null, outStr, checkBox_Mark.Checked);
        }

        private void CheckBox_RTS1_CheckedChanged(object sender, EventArgs e)
        {
            DateTime time = DateTime.Now;
            try
            {
                serialPort1.DtrEnable = checkBox_DTR1.Checked;
                time = DateTime.Now;
            }
            catch (Exception ex)
            {
                datalog.Add(serialPort1.PortName, Logger.DirectionType.Error, DateTime.Now, null, "Error setting DTR: " + ex.Message, checkBox_Mark.Checked);
            }
            string outStr = "";
            if (serialPort1.RtsEnable == true && o_rts1 == false && serialPort1.Handshake != Handshake.RequestToSend && serialPort1.Handshake != Handshake.RequestToSendXOnXOff)
            {
                o_rts1 = true;
                outStr += "<RTS^>";
            }
            else if (serialPort1.RtsEnable == false && o_rts1 == true)
            {
                o_rts1 = false;
                outStr += "<RTSv>";
            }
            datalog.Add(serialPort1.PortName, Logger.DirectionType.SignalOut, time, null, outStr, checkBox_Mark.Checked);
        }

        private void CheckBox_RTS2_CheckedChanged(object sender, EventArgs e)
        {
            DateTime time = DateTime.Now;
            try
            {
                serialPort2.DtrEnable = checkBox_DTR1.Checked;
                time = DateTime.Now;
            }
            catch (Exception ex)
            {
                datalog.Add(serialPort2.PortName, Logger.DirectionType.Error, DateTime.Now, null, "Error setting DTR: " + ex.Message, checkBox_Mark.Checked);
            }
            string outStr = "";
            if (serialPort2.RtsEnable == true && o_rts2 == false)
            {
                o_rts2 = true;
                outStr += "<RTS^>";
            }
            else if (serialPort2.RtsEnable == false && o_rts2 == true)
            {
                o_rts2 = false;
                outStr += "<RTSv>";
            }
            datalog.Add(serialPort2.PortName, Logger.DirectionType.SignalOut, time, null, outStr, checkBox_Mark.Checked);
        }

        private void CheckBox_RTS3_CheckedChanged(object sender, EventArgs e)
        {
            DateTime time = DateTime.Now;
            try
            {
                serialPort3.DtrEnable = checkBox_DTR1.Checked;
                time = DateTime.Now;
            }
            catch (Exception ex)
            {
                datalog.Add(serialPort3.PortName, Logger.DirectionType.Error, DateTime.Now, null, "Error setting DTR: " + ex.Message, checkBox_Mark.Checked);
            }
            string outStr = "";
            if (serialPort3.RtsEnable == true && o_rts3 == false)
            {
                o_rts3 = true;
                outStr += "<RTS^>";
            }
            else if (serialPort3.RtsEnable == false && o_rts3 == true)
            {
                o_rts3 = false;
                outStr += "<RTSv>";
            }
            datalog.Add(serialPort3.PortName, Logger.DirectionType.SignalOut, time, null, outStr, checkBox_Mark.Checked);
        }

        private void CheckBox_RTS4_CheckedChanged(object sender, EventArgs e)
        {
            DateTime time = DateTime.Now;
            try
            {
                serialPort4.DtrEnable = checkBox_DTR1.Checked;
                time = DateTime.Now;
            }
            catch (Exception ex)
            {
                datalog.Add(serialPort4.PortName, Logger.DirectionType.Error, DateTime.Now, null, "Error setting DTR: " + ex.Message, checkBox_Mark.Checked);
            }
            string outStr = "";
            if (serialPort4.RtsEnable == true && o_rts4 == false)
            {
                o_rts4 = true;
                outStr += "<RTS^>";
            }
            else if (serialPort4.RtsEnable == false && o_rts4 == true)
            {
                o_rts4 = false;
                outStr += "<RTSv>";
            }
            datalog.Add(serialPort4.PortName, Logger.DirectionType.SignalOut, time, null, outStr, checkBox_Mark.Checked);
        }

        private void TextBox_custom_command_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (checkBox_commandhex.Checked == true)
            {
                char c = e.KeyChar;
                if (c != '\b' && !((c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f') || (c >= '0' && c <= '9') || c == 0x08 || c == ' '))
                {
                    e.Handled = true;
                }
            }
        }

        private void TextBox_params_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (checkBox_paramhex.Checked == true)
            {
                char c = e.KeyChar;
                if (c != '\b' && !((c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f') || (c >= '0' && c <= '9') || c == 0x08 || c == ' '))
                {
                    e.Handled = true;
                }
            }
        }

        private void TextBox_suff_KeyPress(object sender, KeyPressEventArgs e)
        {
            char c = e.KeyChar;
            if (c != '\b' && !((c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f') || (c >= '0' && c <= '9') || c == 0x08 || c == ' '))
            {
                e.Handled = true;
            }
        }

        private void CheckBox_suff_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_suff.Checked == true) textBox_suff.Enabled = false;
            else textBox_suff.Enabled = true;
            SendStringCollect();
        }

        private void Button_Refresh_Click(object sender, EventArgs e)
        {
            SerialPopulate();
        }

        private void Button_clear1_Click(object sender, EventArgs e)
        {
            textBox_terminal.Clear();
            CSVdataTable.Rows.Clear();
        }

        private void ComboBox_portname1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_portname1.SelectedIndex != 0 && comboBox_portname1.SelectedIndex == comboBox_portname2.SelectedIndex) comboBox_portname1.SelectedIndex = 0;
            if (comboBox_portname1.SelectedIndex != 0 && comboBox_portname1.SelectedIndex == comboBox_portname3.SelectedIndex) comboBox_portname1.SelectedIndex = 0;
            if (comboBox_portname1.SelectedIndex != 0 && comboBox_portname1.SelectedIndex == comboBox_portname4.SelectedIndex) comboBox_portname1.SelectedIndex = 0;
            if (comboBox_portname1.SelectedIndex == 0)
            {
                comboBox_portspeed1.Enabled = false;
                comboBox_handshake1.Enabled = false;
                comboBox_databits1.Enabled = false;
                comboBox_parity1.Enabled = false;
                comboBox_stopbits1.Enabled = false;
                checkBox_sendPort1.Enabled = false;
                checkBox_sendPort1.Checked = false;
                checkBox_displayPort1hex.Enabled = false;
            }
            else
            {
                if (comboBox_portname2.SelectedIndex > 0)
                {
                    comboBox_portspeed1.SelectedIndex = comboBox_portspeed2.SelectedIndex;
                    comboBox_handshake1.SelectedIndex = comboBox_handshake2.SelectedIndex;
                    comboBox_databits1.SelectedIndex = comboBox_databits2.SelectedIndex;
                    comboBox_parity1.SelectedIndex = comboBox_parity2.SelectedIndex;
                    comboBox_stopbits1.SelectedIndex = comboBox_stopbits2.SelectedIndex;
                }
                else
                {
                    comboBox_portspeed1.SelectedIndex = 0;
                    comboBox_handshake1.SelectedIndex = 0;
                    comboBox_databits1.SelectedIndex = 0;
                    comboBox_parity1.SelectedIndex = 2;
                    comboBox_stopbits1.SelectedIndex = 1;
                }
                comboBox_portspeed1.Enabled = true;
                comboBox_handshake1.Enabled = true;
                comboBox_databits1.Enabled = true;
                comboBox_parity1.Enabled = true;
                comboBox_stopbits1.Enabled = true;
                checkBox_sendPort1.Enabled = true;
                checkBox_displayPort1hex.Enabled = true;
            }
            if (comboBox_portname1.SelectedIndex == 0 && comboBox_portname2.SelectedIndex == 0 && comboBox_portname3.SelectedIndex == 0 && comboBox_portname4.SelectedIndex == 0) button_openport.Enabled = false;
            else button_openport.Enabled = true;
        }

        private void ComboBox_portname2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_portname2.SelectedIndex != 0 && comboBox_portname2.SelectedIndex == comboBox_portname1.SelectedIndex) comboBox_portname2.SelectedIndex = 0;
            if (comboBox_portname2.SelectedIndex != 0 && comboBox_portname2.SelectedIndex == comboBox_portname3.SelectedIndex) comboBox_portname2.SelectedIndex = 0;
            if (comboBox_portname2.SelectedIndex != 0 && comboBox_portname2.SelectedIndex == comboBox_portname4.SelectedIndex) comboBox_portname2.SelectedIndex = 0;
            if (comboBox_portname2.SelectedIndex == 0)
            {
                comboBox_portspeed2.Enabled = false;
                comboBox_handshake2.Enabled = false;
                comboBox_databits2.Enabled = false;
                comboBox_parity2.Enabled = false;
                comboBox_stopbits2.Enabled = false;
                checkBox_sendPort2.Enabled = false;
                checkBox_sendPort2.Checked = false;
                checkBox_displayPort2hex.Enabled = false;
            }
            else
            {
                if (comboBox_portname1.SelectedIndex > 0)
                {

                    comboBox_portspeed2.SelectedIndex = comboBox_portspeed1.SelectedIndex;
                    comboBox_handshake2.SelectedIndex = comboBox_handshake1.SelectedIndex;
                    comboBox_databits2.SelectedIndex = comboBox_databits1.SelectedIndex;
                    comboBox_parity2.SelectedIndex = comboBox_parity1.SelectedIndex;
                    comboBox_stopbits2.SelectedIndex = comboBox_stopbits1.SelectedIndex;
                }
                else
                {
                    comboBox_portspeed2.SelectedIndex = 0;
                    comboBox_handshake2.SelectedIndex = 0;
                    comboBox_databits2.SelectedIndex = 0;
                    comboBox_parity2.SelectedIndex = 2;
                    comboBox_stopbits2.SelectedIndex = 1;
                }
                comboBox_portspeed2.Enabled = true;
                comboBox_handshake2.Enabled = true;
                comboBox_databits2.Enabled = true;
                comboBox_parity2.Enabled = true;
                comboBox_stopbits2.Enabled = true;
                checkBox_sendPort2.Enabled = true;
                checkBox_displayPort2hex.Enabled = true;
            }
            if (comboBox_portname1.SelectedIndex == 0 && comboBox_portname2.SelectedIndex == 0 && comboBox_portname3.SelectedIndex == 0 && comboBox_portname4.SelectedIndex == 0) button_openport.Enabled = false;
            else button_openport.Enabled = true;
        }

        private void ComboBox_portname3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_portname3.SelectedIndex != 0 && comboBox_portname3.SelectedIndex == comboBox_portname1.SelectedIndex) comboBox_portname3.SelectedIndex = 0;
            if (comboBox_portname3.SelectedIndex != 0 && comboBox_portname3.SelectedIndex == comboBox_portname2.SelectedIndex) comboBox_portname3.SelectedIndex = 0;
            if (comboBox_portname3.SelectedIndex != 0 && comboBox_portname3.SelectedIndex == comboBox_portname4.SelectedIndex) comboBox_portname3.SelectedIndex = 0;
            if (comboBox_portname3.SelectedIndex == 0)
            {
                comboBox_portspeed3.Enabled = false;
                comboBox_handshake3.Enabled = false;
                comboBox_databits3.Enabled = false;
                comboBox_parity3.Enabled = false;
                comboBox_stopbits3.Enabled = false;
                checkBox_sendPort3.Enabled = false;
                checkBox_sendPort3.Checked = false;
                checkBox_displayPort3hex.Enabled = false;
            }
            else
            {
                if (comboBox_portname4.SelectedIndex > 0)
                {
                    comboBox_portspeed3.SelectedIndex = comboBox_portspeed4.SelectedIndex;
                    comboBox_handshake3.SelectedIndex = comboBox_handshake4.SelectedIndex;
                    comboBox_databits3.SelectedIndex = comboBox_databits4.SelectedIndex;
                    comboBox_parity3.SelectedIndex = comboBox_parity4.SelectedIndex;
                    comboBox_stopbits3.SelectedIndex = comboBox_stopbits4.SelectedIndex;
                }
                else
                {
                    comboBox_portspeed3.SelectedIndex = 0;
                    comboBox_handshake3.SelectedIndex = 0;
                    comboBox_databits3.SelectedIndex = 0;
                    comboBox_parity3.SelectedIndex = 2;
                    comboBox_stopbits3.SelectedIndex = 1;
                }
                comboBox_portspeed3.Enabled = true;
                comboBox_handshake3.Enabled = true;
                comboBox_databits3.Enabled = true;
                comboBox_parity3.Enabled = true;
                comboBox_stopbits3.Enabled = true;
                checkBox_sendPort3.Enabled = true;
                checkBox_displayPort3hex.Enabled = true;
            }
            if (comboBox_portname1.SelectedIndex == 0 && comboBox_portname2.SelectedIndex == 0 && comboBox_portname3.SelectedIndex == 0 && comboBox_portname4.SelectedIndex == 0) button_openport.Enabled = false;
            else button_openport.Enabled = true;
        }

        private void ComboBox_portname4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_portname4.SelectedIndex != 0 && comboBox_portname4.SelectedIndex == comboBox_portname1.SelectedIndex) comboBox_portname4.SelectedIndex = 0;
            if (comboBox_portname4.SelectedIndex != 0 && comboBox_portname4.SelectedIndex == comboBox_portname2.SelectedIndex) comboBox_portname4.SelectedIndex = 0;
            if (comboBox_portname4.SelectedIndex != 0 && comboBox_portname4.SelectedIndex == comboBox_portname3.SelectedIndex) comboBox_portname4.SelectedIndex = 0;
            if (comboBox_portname4.SelectedIndex == 0)
            {
                comboBox_portspeed4.Enabled = false;
                comboBox_handshake4.Enabled = false;
                comboBox_databits4.Enabled = false;
                comboBox_parity4.Enabled = false;
                comboBox_stopbits4.Enabled = false;
                checkBox_sendPort4.Enabled = false;
                checkBox_sendPort4.Checked = false;
                checkBox_displayPort4hex.Enabled = false;
            }
            else
            {
                if (comboBox_portname3.SelectedIndex > 0)
                {

                    comboBox_portspeed4.SelectedIndex = comboBox_portspeed3.SelectedIndex;
                    comboBox_handshake4.SelectedIndex = comboBox_handshake3.SelectedIndex;
                    comboBox_databits4.SelectedIndex = comboBox_databits3.SelectedIndex;
                    comboBox_parity4.SelectedIndex = comboBox_parity3.SelectedIndex;
                    comboBox_stopbits4.SelectedIndex = comboBox_stopbits3.SelectedIndex;
                }
                else
                {
                    comboBox_portspeed4.SelectedIndex = 0;
                    comboBox_handshake4.SelectedIndex = 0;
                    comboBox_databits4.SelectedIndex = 0;
                    comboBox_parity4.SelectedIndex = 2;
                    comboBox_stopbits4.SelectedIndex = 1;
                }
                comboBox_portspeed4.Enabled = true;
                comboBox_handshake4.Enabled = true;
                comboBox_databits4.Enabled = true;
                comboBox_parity4.Enabled = true;
                comboBox_stopbits4.Enabled = true;
                checkBox_sendPort4.Enabled = true;
                checkBox_displayPort4hex.Enabled = true;
            }
            if (comboBox_portname1.SelectedIndex == 0 && comboBox_portname2.SelectedIndex == 0 && comboBox_portname3.SelectedIndex == 0 && comboBox_portname4.SelectedIndex == 0) button_openport.Enabled = false;
            else button_openport.Enabled = true;
        }

        private void CheckBox_commandhex_CheckedChanged(object sender, EventArgs e)
        {
            string tmpstr = textBox_command.Text;
            if (checkBox_commandhex.Checked == true) textBox_command.Text = Accessory.ConvertStringToHex(tmpstr);
            else textBox_command.Text = Accessory.ConvertHexToString(tmpstr);
        }

        private void CheckBox_paramhex_CheckedChanged(object sender, EventArgs e)
        {
            string tmpstr = textBox_params.Text;
            if (checkBox_paramhex.Checked == true) textBox_params.Text = Accessory.ConvertStringToHex(tmpstr);
            else textBox_params.Text = Accessory.ConvertHexToString(tmpstr);
        }

        private void CheckBox_send_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_sendPort1.Checked == false && checkBox_sendPort2.Checked == false && checkBox_sendPort3.Checked == false && checkBox_sendPort4.Checked == false) button_send.Enabled = false;
            else if (serialPort1.IsOpen == true || serialPort2.IsOpen == true || serialPort3.IsOpen == true || serialPort4.IsOpen == true) button_send.Enabled = true;
        }

        private void TextBox_command_Leave(object sender, EventArgs e)
        {
            if (checkBox_commandhex.Checked == true) textBox_command.Text = Accessory.CheckHexString(textBox_command.Text);
            SendStringCollect();
        }

        private void TextBox_params_Leave(object sender, EventArgs e)
        {
            if (checkBox_paramhex.Checked == true) textBox_params.Text = Accessory.CheckHexString(textBox_params.Text);
            SendStringCollect();
        }

        private void TextBox_suff_Leave(object sender, EventArgs e)
        {
            if (checkBox_suffhex.Checked == true) textBox_suff.Text = Accessory.CheckHexString(textBox_suff.Text);
            SendStringCollect();
        }

        private void CheckBox_cr_CheckedChanged(object sender, EventArgs e)
        {
            SendStringCollect();
        }

        private void CheckBox_lf_CheckedChanged(object sender, EventArgs e)
        {
            SendStringCollect();
        }

        private void CheckBox_suffhex_CheckedChanged(object sender, EventArgs e)
        {
            string tmpstr = textBox_suff.Text;
            if (checkBox_suffhex.Checked == true) textBox_suff.Text = Accessory.ConvertStringToHex(tmpstr);
            else textBox_suff.Text = Accessory.ConvertHexToString(tmpstr);
        }

        private void CheckBox_portName_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_portName.Checked == true)
            {
                if (button_closeport.Enabled == true)
                {
                    textBox_port1Name.Enabled = false;
                    textBox_port2Name.Enabled = false;
                    textBox_port3Name.Enabled = false;
                    textBox_port4Name.Enabled = false;
                }
                checkBox_sendPort1.Text = textBox_port1Name.Text;
                checkBox_sendPort2.Text = textBox_port2Name.Text;
                checkBox_sendPort3.Text = textBox_port3Name.Text;
                checkBox_sendPort4.Text = textBox_port4Name.Text;
                checkBox_displayPort1hex.Text = textBox_port1Name.Text;
                checkBox_displayPort2hex.Text = textBox_port2Name.Text;
                checkBox_displayPort3hex.Text = textBox_port3Name.Text;
                checkBox_displayPort4hex.Text = textBox_port4Name.Text;
            }
            else
            {
                textBox_port1Name.Enabled = true;
                textBox_port2Name.Enabled = true;
                textBox_port3Name.Enabled = true;
                textBox_port4Name.Enabled = true;
                checkBox_sendPort1.Text = comboBox_portname1.Text;
                checkBox_sendPort2.Text = comboBox_portname2.Text;
                checkBox_sendPort3.Text = comboBox_portname3.Text;
                checkBox_sendPort4.Text = comboBox_portname4.Text;
                checkBox_displayPort1hex.Text = comboBox_portname1.Text;
                checkBox_displayPort2hex.Text = comboBox_portname2.Text;
                checkBox_displayPort3hex.Text = comboBox_portname3.Text;
                checkBox_displayPort4hex.Text = comboBox_portname4.Text;
            }
        }

        private void SaveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            if (saveFileDialog.Title == "Save .TXT log as...")
            {
                try
                {
                    File.WriteAllText(saveFileDialog.FileName, textBox_terminal.Text);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error writing to file " + saveFileDialog.FileName + ": " + ex.Message);
                }
            }
            if (saveFileDialog.Title == "Save .CSV log as...")
            {
                int columnCount = dataGridView.ColumnCount;
                string output = "";
                for (int i = 0; i < columnCount; i++)
                {
                    output += dataGridView.Columns[i].Name.ToString() + ",";
                }
                output += "\r\n";
                for (int i = 1; (i - 1) < dataGridView.RowCount; i++)
                {
                    for (int j = 0; j < columnCount; j++)
                    {
                        output += dataGridView.Rows[i - 1].Cells[j].Value.ToString() + ",";
                    }
                    output += "\r\n";
                }
                try
                {
                    File.WriteAllText(saveFileDialog.FileName, output, Encoding.GetEncoding(RS232_monitor2.Properties.Settings.Default.CodePage));
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error writing to file " + saveFileDialog.FileName + ": " + ex.Message);
                }
            }
        }

        private void SaveTXTToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog.Title = "Save .TXT log as...";
            saveFileDialog.DefaultExt = "txt";
            saveFileDialog.Filter = "Text files|*.txt|All files|*.*";
            saveFileDialog.FileName = "terminal_" + DateTime.Today.ToShortDateString().Replace("/", "_") + ".txt";
            saveFileDialog.ShowDialog();
        }

        private void SaveCSVToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog.Title = "Save .CSV log as...";
            saveFileDialog.DefaultExt = "csv";
            saveFileDialog.Filter = "CSV files|*.csv|All files|*.*";
            saveFileDialog.FileName = "terminal_" + DateTime.Today.ToShortDateString().Replace("/", "_") + ".csv";
            saveFileDialog.ShowDialog();
        }

        private void AboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("RS232 Monitor2\r\n(c) Kalugin Andrey\r\nContact: jekyll@mail.ru");
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void SaveParametersToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            RS232_monitor2.Properties.Settings.Default.DefaultCommand = textBox_command.Text;
            RS232_monitor2.Properties.Settings.Default.DefaultCommandHex = checkBox_commandhex.Checked;
            RS232_monitor2.Properties.Settings.Default.DefaultParameter = textBox_params.Text;
            RS232_monitor2.Properties.Settings.Default.DefaultParamHex = checkBox_paramhex.Checked;
            RS232_monitor2.Properties.Settings.Default.addCR = checkBox_cr.Checked;
            RS232_monitor2.Properties.Settings.Default.addLF = checkBox_lf.Checked;
            RS232_monitor2.Properties.Settings.Default.addSuff = checkBox_suff.Checked;
            RS232_monitor2.Properties.Settings.Default.SuffText = textBox_suff.Text;
            RS232_monitor2.Properties.Settings.Default.DefaultSuffHex = checkBox_suffhex.Checked;
            RS232_monitor2.Properties.Settings.Default.LogSignal = checkBox_insPin.Checked;
            RS232_monitor2.Properties.Settings.Default.LogTime = checkBox_insTime.Checked;
            RS232_monitor2.Properties.Settings.Default.LogDir = checkBox_insDir.Checked;
            RS232_monitor2.Properties.Settings.Default.LogPortName = checkBox_portName.Checked;
            RS232_monitor2.Properties.Settings.Default.HexPort1 = checkBox_displayPort1hex.Checked;
            RS232_monitor2.Properties.Settings.Default.HexPort2 = checkBox_displayPort2hex.Checked;
            RS232_monitor2.Properties.Settings.Default.HexPort3 = checkBox_displayPort3hex.Checked;
            RS232_monitor2.Properties.Settings.Default.HexPort4 = checkBox_displayPort4hex.Checked;
            RS232_monitor2.Properties.Settings.Default.Port1Name = textBox_port1Name.Text;
            RS232_monitor2.Properties.Settings.Default.Port2Name = textBox_port2Name.Text;
            RS232_monitor2.Properties.Settings.Default.Port3Name = textBox_port3Name.Text;
            RS232_monitor2.Properties.Settings.Default.Port4Name = textBox_port4Name.Text;
            RS232_monitor2.Properties.Settings.Default.LogGrid = logToGridToolStripMenuItem.Checked;
            RS232_monitor2.Properties.Settings.Default.LogText = logToTextToolStripMenuItem.Checked;
            RS232_monitor2.Properties.Settings.Default.AutoScroll = autoscrollToolStripMenuItem.Checked;
            RS232_monitor2.Properties.Settings.Default.LineWrap = lineWrapToolStripMenuItem.Checked;
            RS232_monitor2.Properties.Settings.Default.AutoLogTXT = autosaveTXTToolStripMenuItem1.Checked;
            RS232_monitor2.Properties.Settings.Default.TXTlogFile = terminaltxtToolStripMenuItem1.Text;
            RS232_monitor2.Properties.Settings.Default.AutoLogCSV = autosaveCSVToolStripMenuItem1.Checked;
            RS232_monitor2.Properties.Settings.Default.LineBreakTimeout = LineBreakTimeout / 10000;
            RS232_monitor2.Properties.Settings.Default.CSVMaxLineNumber = CSVLineNumberLimit;
            RS232_monitor2.Properties.Settings.Default.LineLengthLimit = LineLengthLimit;

            RS232_monitor2.Properties.Settings.Default.Save();
        }

        private void AutosaveTXTToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            autosaveTXTToolStripMenuItem1.Checked = !autosaveTXTToolStripMenuItem1.Checked;
            terminaltxtToolStripMenuItem1.Enabled = !autosaveTXTToolStripMenuItem1.Checked;
        }

        private void LineWrapToolStripMenuItem_Click(object sender, EventArgs e)
        {
            lineWrapToolStripMenuItem.Checked = !lineWrapToolStripMenuItem.Checked;
            textBox_terminal.WordWrap = lineWrapToolStripMenuItem.Checked;
        }

        private void AutoscrollToolStripMenuItem_Click(object sender, EventArgs e)
        {
            autoscrollToolStripMenuItem.Checked = !autoscrollToolStripMenuItem.Checked;
        }

        private void LogToTextToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (logToTextToolStripMenuItem.Checked == true)
            {
                logToTextToolStripMenuItem.Checked = false;
                textBox_terminal.Enabled = false;
                ((Control)tabPage2).Enabled = false;
                if (logToGridToolStripMenuItem.Checked == false)
                {
                    tabControl1.Enabled = false;
                    tabControl1.Visible = false;
                }
            }
            else
            {
                logToTextToolStripMenuItem.Checked = true;
                textBox_terminal.Enabled = true;
                ((Control)tabPage2).Enabled = true;
                tabControl1.Enabled = true;
                tabControl1.Visible = true;
            }
        }

        private void TabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (e.TabPage == tabPage1 && logToTextToolStripMenuItem.Checked == false)
                e.Cancel = true;
            if (e.TabPage == tabPage2 && logToGridToolStripMenuItem.Checked == false)
                e.Cancel = true;
        }

        private void checkBox_displayPort1hex_CheckedChanged(object sender, EventArgs e)
        {
            displayhex[0] = checkBox_displayPort1hex.Checked;
        }

        private void checkBox_displayPort2hex_CheckedChanged(object sender, EventArgs e)
        {
            displayhex[1] = checkBox_displayPort2hex.Checked;
        }

        private void checkBox_displayPort3hex_CheckedChanged(object sender, EventArgs e)
        {
            displayhex[2] = checkBox_displayPort3hex.Checked;
        }

        private void checkBox_displayPort4hex_CheckedChanged(object sender, EventArgs e)
        {
            displayhex[3] = checkBox_displayPort4hex.Checked;
        }

        private void LineLengthToolStripTextBox1_Leave(object sender, EventArgs e)
        {
            int.TryParse(LineBreakToolStripTextBox1.Text, out LineLengthLimit);
            LineBreakToolStripTextBox1.Text = LineLengthLimit.ToString();
        }

        private void LogToGridToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (logToGridToolStripMenuItem.Checked == true)
            {
                logToGridToolStripMenuItem.Checked = false;
                dataGridView.Enabled = false;
                ((Control)tabPage1).Enabled = false;
                if (logToTextToolStripMenuItem.Checked == false)
                {
                    tabControl1.Enabled = false;
                    tabControl1.Visible = false;
                }
            }
            else
            {
                logToGridToolStripMenuItem.Checked = true;
                dataGridView.Enabled = true;
                ((Control)tabPage1).Enabled = true;
                tabControl1.Enabled = true;
                tabControl1.Visible = true;
            }
        }

        private void CheckBox_Mark_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_Mark.Checked == true) checkBox_Mark.Font = new Font(checkBox_Mark.Font, FontStyle.Bold);
            else checkBox_Mark.Font = new Font(checkBox_Mark.Font, FontStyle.Regular);
        }

        private void TextBox_command_KeyUp(object sender, KeyEventArgs e)
        {
            if (button_send.Enabled == true)
                if (e.KeyData == Keys.Return)
                    Button_send_Click(textBox_command, EventArgs.Empty);
        }

        private void ToolStripMenuItem_onlyData_Click(object sender, EventArgs e)
        {
            toolStripMenuItem_onlyData.Checked = !toolStripMenuItem_onlyData.Checked;

            if (toolStripMenuItem_onlyData.Checked == false)
            {
                serialPort1.ErrorReceived += new System.IO.Ports.SerialErrorReceivedEventHandler(SerialPort_ErrorReceived);
                serialPort2.ErrorReceived += new System.IO.Ports.SerialErrorReceivedEventHandler(SerialPort_ErrorReceived);
                serialPort3.ErrorReceived += new System.IO.Ports.SerialErrorReceivedEventHandler(SerialPort_ErrorReceived);
                serialPort4.ErrorReceived += new System.IO.Ports.SerialErrorReceivedEventHandler(SerialPort_ErrorReceived);
                checkBox_insPin.Checked = true;
            }
            else
            {
                serialPort1.ErrorReceived -= SerialPort_ErrorReceived;
                serialPort2.ErrorReceived -= SerialPort_ErrorReceived;
                serialPort3.ErrorReceived -= SerialPort_ErrorReceived;
                serialPort4.ErrorReceived -= SerialPort_ErrorReceived;
                checkBox_insPin.Checked = false;
            }
        }

        private void SerialPopulate()
        {
            comboBox_portname1.Items.Clear();
            comboBox_handshake1.Items.Clear();
            comboBox_parity1.Items.Clear();
            comboBox_stopbits1.Items.Clear();

            comboBox_portname2.Items.Clear();
            comboBox_handshake2.Items.Clear();
            comboBox_parity2.Items.Clear();
            comboBox_stopbits2.Items.Clear();

            comboBox_portname3.Items.Clear();
            comboBox_handshake3.Items.Clear();
            comboBox_parity3.Items.Clear();
            comboBox_stopbits3.Items.Clear();

            comboBox_portname4.Items.Clear();
            comboBox_handshake4.Items.Clear();
            comboBox_parity4.Items.Clear();
            comboBox_stopbits4.Items.Clear();

            //Serial settings populate
            comboBox_portname1.Items.Add("-None-");
            comboBox_portname2.Items.Add("-None-");
            comboBox_portname3.Items.Add("-None-");
            comboBox_portname4.Items.Add("-None-");

            //Add ports
            foreach (string s in SerialPort.GetPortNames())
            {
                comboBox_portname1.Items.Add(s);
                comboBox_portname2.Items.Add(s);
                comboBox_portname3.Items.Add(s);
                comboBox_portname4.Items.Add(s);
            }
            //Add handshake methods
            foreach (string s in Enum.GetNames(typeof(Handshake)))
            {
                comboBox_handshake1.Items.Add(s);
                comboBox_handshake2.Items.Add(s);
                comboBox_handshake3.Items.Add(s);
                comboBox_handshake4.Items.Add(s);
            }
            //Add parity
            foreach (string s in Enum.GetNames(typeof(Parity)))
            {
                comboBox_parity1.Items.Add(s);
                comboBox_parity2.Items.Add(s);
                comboBox_parity3.Items.Add(s);
                comboBox_parity4.Items.Add(s);
            }
            //Add stopbits
            foreach (string s in Enum.GetNames(typeof(StopBits)))
            {
                comboBox_stopbits1.Items.Add(s);
                comboBox_stopbits2.Items.Add(s);
                comboBox_stopbits3.Items.Add(s);
                comboBox_stopbits4.Items.Add(s);
            }

            if (comboBox_portname1.Items.Count > 1)
            {
                comboBox_portname1.SelectedIndex = 1;
                comboBox_portspeed1.SelectedIndex = 0;
                comboBox_handshake1.SelectedIndex = 0;
                comboBox_databits1.SelectedIndex = 0;
                comboBox_parity1.SelectedIndex = 2;
                comboBox_stopbits1.SelectedIndex = 1;
                checkBox_sendPort1.Enabled = true;
                checkBox_displayPort1hex.Enabled = true;
            }
            else
            {
                comboBox_portname1.SelectedIndex = 0;
                checkBox_sendPort1.Enabled = false;
                checkBox_displayPort1hex.Enabled = false;
            }

            if (comboBox_portname2.Items.Count > 2)
            {
                comboBox_portname2.SelectedIndex = 2;
                comboBox_portspeed2.SelectedIndex = 0;
                comboBox_handshake2.SelectedIndex = 0;
                comboBox_databits2.SelectedIndex = 0;
                comboBox_parity2.SelectedIndex = 2;
                comboBox_stopbits2.SelectedIndex = 1;
                checkBox_sendPort2.Enabled = true;
                checkBox_displayPort2hex.Enabled = true;
            }
            else
            {
                comboBox_portname2.SelectedIndex = 0;
                checkBox_sendPort2.Enabled = false;
                checkBox_displayPort2hex.Enabled = false;
            }

            if (comboBox_portname3.Items.Count > 3)
            {
                comboBox_portname3.SelectedIndex = 3;
                comboBox_portspeed3.SelectedIndex = 0;
                comboBox_handshake3.SelectedIndex = 0;
                comboBox_databits3.SelectedIndex = 0;
                comboBox_parity3.SelectedIndex = 2;
                comboBox_stopbits3.SelectedIndex = 1;
                checkBox_sendPort3.Enabled = true;
                checkBox_displayPort3hex.Enabled = true;
            }
            else
            {
                comboBox_portname3.SelectedIndex = 0;
                checkBox_sendPort3.Enabled = false;
                checkBox_displayPort3hex.Enabled = false;
            }

            if (comboBox_portname4.Items.Count > 4)
            {
                comboBox_portname4.SelectedIndex = 4;
                comboBox_portspeed4.SelectedIndex = 0;
                comboBox_handshake4.SelectedIndex = 0;
                comboBox_databits4.SelectedIndex = 0;
                comboBox_parity4.SelectedIndex = 2;
                comboBox_stopbits4.SelectedIndex = 1;
                checkBox_sendPort4.Enabled = true;
                checkBox_displayPort4hex.Enabled = true;
            }
            else
            {
                comboBox_portname4.SelectedIndex = 0;
                checkBox_sendPort4.Enabled = false;
                checkBox_displayPort4hex.Enabled = false;
            }

            if (comboBox_portname1.SelectedIndex == 0 && comboBox_portname2.SelectedIndex == 0 && comboBox_portname3.SelectedIndex == 0 && comboBox_portname4.SelectedIndex == 0) button_openport.Enabled = false;
            else button_openport.Enabled = true;
            CheckBox_portName_CheckedChanged(this, EventArgs.Empty);
        }

        private void SendStringCollect()
        {
            string tmpStr;
            if (checkBox_commandhex.Checked == true) tmpStr = textBox_command.Text.Trim();
            else tmpStr = Accessory.ConvertStringToHex(textBox_command.Text).Trim();
            if (checkBox_paramhex.Checked == true) tmpStr += " " + textBox_params.Text.Trim();
            else tmpStr += " " + Accessory.ConvertStringToHex(textBox_params.Text).Trim();
            if (checkBox_cr.Checked == true) tmpStr += " 0D";
            if (checkBox_lf.Checked == true) tmpStr += " 0A";
            if (checkBox_suff.Checked == true)
            {
                if (checkBox_suffhex.Checked == true) tmpStr += " " + textBox_suff.Text.Trim();
                else tmpStr += " " + Accessory.ConvertStringToHex(textBox_suff.Text).Trim();
            }
            textBox_senddata.Text = Accessory.CheckHexString(tmpStr);
        }

        private void CheckBox_insPin_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_insPin.Checked == true)
            {
                serialPort1.PinChanged += SerialPort_PinChanged;
                serialPort2.PinChanged += SerialPort_PinChanged;
                serialPort3.PinChanged += SerialPort_PinChanged;
                serialPort4.PinChanged += SerialPort_PinChanged;
            }
            else
            {
                serialPort1.PinChanged -= SerialPort_PinChanged;
                serialPort2.PinChanged -= SerialPort_PinChanged;
                serialPort3.PinChanged -= SerialPort_PinChanged;
                serialPort4.PinChanged -= SerialPort_PinChanged;
            }
        }

        private void ToolStripTextBox_CSVLinesNumber_Leave(object sender, EventArgs e)
        {
            Int32.TryParse(toolStripTextBox_CSVLinesNumber.Text, out CSVLineNumberLimit);
            if (CSVLineNumberLimit < 1)
            {
                CSVLineNumberLimit = 1;
            }
            toolStripTextBox_CSVLinesNumber.Text = CSVLineNumberLimit.ToString();
        }

        private void LineBreakToolStripTextBox1_Leave(object sender, EventArgs e)
        {
            int.TryParse(LineBreakToolStripTextBox1.Text, out LineBreakTimeout);
            LineBreakToolStripTextBox1.Text = LineBreakTimeout.ToString();
        }

        private void AutosaveCSVToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            autosaveCSVToolStripMenuItem1.Checked = !autosaveCSVToolStripMenuItem1.Checked;
        }

        //base function run by timer
        public void CollectBuffer(object sender, EventArgs e)
        {
            if (t != null && t.IsAlive) return;
            aTimer.Enabled = false;
            t = new Thread(new ThreadStart(ManageLog));
            t.Start();
        }

        //separate thread to get log records and toss
        public void ManageLog()
        {
            if (datalog.QueueSize() <= 0)
            {
                this.Invoke((MethodInvoker)delegate
                {
                    toolStripStatusLabel1.Text = "Idle...";
                    toolStripStatusLabel1.BackColor = Color.White;
                });

                aTimer.Enabled = true;
                return;
            }
            this.Invoke((MethodInvoker)delegate
            {
                toolStripStatusLabel1.Text = "Processing data...";
                toolStripStatusLabel1.BackColor = Color.Red;
            });


            List<Logger.LogRecord> tmp = new List<Logger.LogRecord>();
            tmp.AddRange(datalog.GetLog());

            //sort list of records by time just in case
            //tmp.Sort((x, y) => x.dateTime.CompareTo(y.dateTime));

            //combine/split data based on "LineBreakTimeout" and "LineLengthLimit"

            DateTime tmpTime = tmp[0].dateTime;
            for (int i = 0; i < tmp.Count - 1; i++)
            {
                if (tmp[i].portName == tmp[i + 1].portName &&
                    tmp[i].direction == tmp[i + 1].direction &&
                    (tmp[i].message.Length + tmp[i + 1].message.Length) <= LineLengthLimit &&
                    tmp[i + 1].dateTime.Subtract(tmpTime).Milliseconds <= LineBreakTimeout)
                {
                    //move message and signalPin from i+1 to i
                    tmpTime = tmp[i + 1].dateTime;
                    Logger.LogRecord t = tmp[i];
                    t.message = CombineByteArrays(tmp[i].message, tmp[i + 1].message);
                    t.signalPin = tmp[i].signalPin + tmp[i + 1].signalPin;
                    tmp[i] = t;
                    tmp.RemoveAt(i + 1);
                    i--;
                }
                else tmpTime = tmp[i + 1].dateTime;
            }

            //put records array into the GridView
            if (logToGridToolStripMenuItem.Checked || autosaveCSVToolStripMenuItem1.Checked)
            {
                List<DataRow> DataRowArray = new List<DataRow>();
                foreach (Logger.LogRecord record in tmp)
                {
                    //get number of port from portname
                    int portnum = portName.IndexOf(record.portName);
                    //make record for DataTable
                    DataRow tempRow = CSVdataTable.NewRow();
                    tempRow["Date"] = record.dateTime.ToShortDateString();
                    tempRow["Time"] = record.dateTime.ToLongTimeString();
                    tempRow["Milis"] = record.dateTime.Millisecond.ToString("D3");
                    if (checkBox_portName.Checked) tempRow["Port"] = altPortName[portnum];
                    else tempRow["Port"] = record.portName;
                    tempRow["Dir"] = datalog.DirectionMark[(int)record.direction];
                    tempRow["Data"] = Accessory.ConvertByteArrayToHex(record.message);
                    tempRow["Signal"] = record.signalPin;
                    tempRow["Mark"] = record.mark.ToString();
                    DataRowArray.Add(tempRow);
                }

                if (logToGridToolStripMenuItem.Checked) SendToGridView(DataRowArray);
                //save record to CSV if needed
                if (autosaveCSVToolStripMenuItem1.Checked)
                {
                    CSVcollectBuffer(DataRowArray);
                }
            }

            if (logToTextToolStripMenuItem.Checked || autosaveTXTToolStripMenuItem1.Checked)
            {
                //make string array for textbox
                List<string> newText = new List<string>();
                foreach (Logger.LogRecord record in tmp)
                {
                    //get number of port from portname
                    int portnum = portName.IndexOf(record.portName);
                    //create text strings
                    string tmpBuffer = "";

                    //Date+time+millis
                    if (checkBox_insTime.Checked)
                    {
                        tmpBuffer += record.dateTime.ToShortDateString() + " ";
                        tmpBuffer += record.dateTime.ToLongTimeString() + ",";
                        tmpBuffer += record.dateTime.Millisecond.ToString("D3") + " ";
                    }

                    //port
                    if (checkBox_portName.Checked) tmpBuffer += altPortName[portnum];
                    else tmpBuffer += record.portName;

                    //dir
                    if (checkBox_insDir.Checked) tmpBuffer += datalog.DirectionMark[(int)record.direction] + " ";

                    //signal
                    if (checkBox_insPin.Checked && record.signalPin != "") tmpBuffer += record.signalPin + " ";

                    //data
                    if (displayhex[portnum]) tmpBuffer += Accessory.ConvertByteArrayToHex(record.message);
                    else tmpBuffer += Accessory.ConvertByteArrayToString(record.message);

                    //skip "mark" field in the text
                    tmpBuffer += "\r\n";
                    newText.Add(tmpBuffer);
                }

                //put string array into the TextBox
                if (logToTextToolStripMenuItem.Checked) SendToTextBox(newText.ToArray());

                //save string array to TXT if needed
                if (autosaveTXTToolStripMenuItem1.Checked)
                {
                    try
                    {
                        foreach (string s in newText)
                        {
                            File.AppendAllText(terminaltxtToolStripMenuItem1.Text, s, Encoding.GetEncoding(RS232_monitor2.Properties.Settings.Default.CodePage));
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("\r\nError writing file " + terminaltxtToolStripMenuItem1.Text + ": " + ex.Message);
                    }
                }
            }
            aTimer.Enabled = true;
        }

        //send new text to GridView caring the string number limit
        public void SendToGridView(List<DataRow> tmpDataRow)
        {
            this.Invoke((MethodInvoker)delegate
            {
                dataGridView.SuspendLayout();
                dataGridView.Enabled = false;
            });

            while (CSVdataTable.Rows.Count > 0 &&
                CSVdataTable.Rows.Count + tmpDataRow.Count > LogLinesLimit)
            {
                CSVdataTable.Rows.RemoveAt(0);
            }
            foreach (var r in tmpDataRow)
            {
                CSVdataTable.Rows.Add(r);
            }

            this.Invoke((MethodInvoker)delegate
            {
                dataGridView.ResumeLayout();
                dataGridView.Enabled = true;
            });
        }

        //save text to CSV file caring the string number limit
        public void CSVcollectBuffer(List<DataRow> tmpDataRow)
        {
            if (tmpDataRow == null || tmpDataRow.Count == 0) return;

            foreach (var r in tmpDataRow)
            {
                //create CSV strings
                string tmpBuffer =
                    r["Date"].ToString() + RS232_monitor2.Properties.Settings.Default.CSVDelimiter +
                    r["Time"].ToString() + RS232_monitor2.Properties.Settings.Default.CSVDelimiter +
                    r["Milis"].ToString() + RS232_monitor2.Properties.Settings.Default.CSVDelimiter +
                    r["Port"].ToString() + RS232_monitor2.Properties.Settings.Default.CSVDelimiter +
                    r["Dir"].ToString() + RS232_monitor2.Properties.Settings.Default.CSVDelimiter +
                    r["Data"].ToString() + RS232_monitor2.Properties.Settings.Default.CSVDelimiter +
                    r["Signal"].ToString() + RS232_monitor2.Properties.Settings.Default.CSVDelimiter +
                    r["Mark"].ToString() + RS232_monitor2.Properties.Settings.Default.CSVDelimiter + "\r\n";

                //save strings to file
                if (CSVLineCount >= CSVLineNumberLimit)
                {
                    CSVFileName = DateTime.Today.ToShortDateString() + "_" + DateTime.Now.ToLongTimeString() + "_" + DateTime.Now.Millisecond.ToString("D3") + ".csv";
                    CSVFileName = CSVFileName.Replace(':', '-').Replace('\\', '-').Replace('/', '-');
                    CSVLineCount = 0;
                }
                try
                {
                    File.AppendAllText(CSVFileName, tmpBuffer, Encoding.GetEncoding(RS232_monitor2.Properties.Settings.Default.CodePage));
                    CSVLineCount++;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("\r\nError writing file " + CSVFileName + ": " + ex.Message);
                }
            }
        }

        //send new text to TextBox caring the string number limit
        private void SendToTextBox(string[] text)
        {
            if (text == null || text.Length == 0) return;
            for (int i = 0; i < text.Length; i++)
            {
                text[i] = Accessory.FilterZeroChar(text[i]);
            }

            int pos = 0;
            int l = 0;
            this.Invoke((MethodInvoker)delegate
            {
                pos = textBox_terminal.SelectionStart;
                //get number of lines in new text
                l = textBox_terminal.Lines.Length + text.Length;
            });

            StringBuilder tmp = new StringBuilder();
            if (l > LogLinesLimit)
            {
                for (int i = l - LogLinesLimit; i < textBox_terminal.Lines.Length; i++)
                {
                    tmp.Append(textBox_terminal.Lines[i]);
                    tmp.Append("\r\n");
                }
                foreach (string s in text) tmp.Append(s);
                this.Invoke((MethodInvoker)delegate
                {
                    textBox_terminal.Text = tmp.ToString();
                });
            }
            else
            {
                this.Invoke((MethodInvoker)delegate
                {
                    foreach (string s in text) textBox_terminal.Text += s;
                });
            }

            this.Invoke((MethodInvoker)delegate
            {

                if (autoscrollToolStripMenuItem.Checked)
                {
                    textBox_terminal.SelectionStart = textBox_terminal.Text.Length;
                    textBox_terminal.ScrollToCaret();
                }
                else
                {
                    textBox_terminal.SelectionStart = pos;
                    textBox_terminal.ScrollToCaret();
                }
            });
        }

        public byte[] CombineByteArrays(byte[] first, byte[] second)
        {
            byte[] ret = new byte[first.Length + second.Length];
            Buffer.BlockCopy(first, 0, ret, 0, first.Length);
            Buffer.BlockCopy(second, 0, ret, first.Length, second.Length);
            return ret;
        }

    }
}
