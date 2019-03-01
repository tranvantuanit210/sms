using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SMSManager
{
    public partial class Form1 : Form
    {
        //congfig sms 
        SerialPort port = new SerialPort();
        clsSMS objclsSMS = new clsSMS();
        ShortMessageCollection objShortMessageCollection = new ShortMessageCollection();
        public bool isModemConnect = false;
        public bool hasFileExcel = false;
        IList<SmsInfor> smsDatas = new List<SmsInfor>();
        string filePath = AppDomain.CurrentDomain.BaseDirectory + "\\content.txt";

        public Form1()
        {
            InitializeComponent();
            btnDeleteFile.Visible = false;
            dataGridView1.DataSource = smsDatas;
        }
        

        private void btn_connect_Click(object sender, EventArgs e)
        {
            if (isModemConnect)
            {
                DisableConnect();
                isModemConnect = false;
            }
            else
            {
                var port = txt_port.Text;
                if (string.IsNullOrEmpty(port))
                {
                    var str_Port_Connect = FindPortAndConnect();
                    NotifiStatusConnect(!string.IsNullOrEmpty(str_Port_Connect), str_Port_Connect);
                }
                else
                {
                    this.port = objclsSMS.OpenPort(port, Convert.ToInt32(115200), Convert.ToInt32(8), Convert.ToInt32(100), Convert.ToInt32(100));
                    bool isConnect = objclsSMS.CheckPortSendSMS(this.port);
                    NotifiStatusConnect(isConnect, port);
                }
                
            }
        }

        public string FindPortAndConnect()
        {
            string result = null;
            string[] ports = SerialPort.GetPortNames();
            foreach (string port in ports)
            {
                this.port = objclsSMS.OpenPort(port, Convert.ToInt32(9600), Convert.ToInt32(8), Convert.ToInt32(100), Convert.ToInt32(100));
                bool isConnect = objclsSMS.CheckPortSendSMS(this.port);
                if (isConnect)
                {
                    return port;

                }

            }
            return result;
        }

        private void btn_send_Click(object sender, EventArgs e)
        {
            Task.Factory.StartNew(new Action(() =>
            {
                try
                {
                    btn_send.BeginInvoke(new Action(() =>
                    {
                        btn_send.Enabled = false;

                    }));

                    if (!isModemConnect)
                    {
                        lbl_status.BeginInvoke(new Action(() =>
                        {
                            lbl_status.Text = "Trạng thái: Không tìm thấy thiết bị Modem GMS!";
                            lbl_status.ForeColor = Color.Red;
                        }));

                        return;
                    }

                    lbl_status.BeginInvoke(new Action(() =>
                    {
                        lbl_status.Text = "Đang gởi tin nhắn...";
                        lbl_status.ForeColor = Color.Blue;

                    }));

                    if (!hasFileExcel)
                    {
                        if (!string.IsNullOrEmpty(txt_phone.Text))
                        {
                            smsDatas.Insert(0, new SmsInfor()
                            {
                                Name = null,
                                Number = txt_phone.Text,
                                Content = txt_message.Text,
                                Money = null,
                                Field01 = null,
                                Field02 = null,
                                Field03 = null,
                            });
                        }                        
                    }

                    var firstSms = smsDatas.FirstOrDefault();
                    if (firstSms == null || string.IsNullOrEmpty(firstSms.Number) || string.IsNullOrEmpty(txt_message.Text))
                    {
                        lbl_status.BeginInvoke(new Action(() =>
                        {
                            lbl_status.Text = "Vui lòng nhập số điện thoại và nội dung tin nhắn";
                            lbl_status.ForeColor = Color.Red;
                        }));
                        
                        return;
                    }

                    var timeout = Convert.ToInt32(txtTimeout.Value);
                    foreach (var item in smsDatas)
                    {
                        var content = txt_message.Text;
                        item.Content = content.Replace("[name]", item.Name).Replace("[money]", item.Money).Replace("[field01]", item.Field01)
                            .Replace("[field02]", item.Field02).Replace("[field03]", item.Field03);
                        var number = item.Number.Substring(item.Number.Length - 9);
                        item.Success = objclsSMS.sendMsg(this.port, "0" + number, item.Content, timeout * 1000);
                    }

                    dataGridView1.BeginInvoke(new Action(() =>
                    {
                        dataGridView1.DataSource = null;
                        dataGridView1.Refresh();
                        dataGridView1.DataSource = smsDatas;

                    }));

                    lbl_status.BeginInvoke(new Action(() =>
                    {
                        lbl_status.Text = "Đã gởi tin nhắn!";
                        lbl_status.ForeColor = Color.Blue;

                    }));

                }
                catch (Exception ex)
                {
                    lbl_status.BeginInvoke(new Action(() =>
                    {
                        lbl_status.Text = ex.Message;
                        lbl_status.ForeColor = Color.Red;
                        
                    }));

                    Debug.Write(ex.Message);
                }
                finally
                {
                    btn_send.BeginInvoke(new Action(() =>
                    {
                        btn_send.Enabled = true;

                    }));
                }
            }));
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = "Mở file Excel";
            openFileDialog.Filter = "Excel Files|*.xls; *xlsx";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                hasFileExcel = true;
                var fileName = openFileDialog.FileName;

                //...first fetch excel file from project folder or any others location…
                var fi = new FileInfo(fileName);

                using (var package = new OfficeOpenXml.ExcelPackage(fi))
                {
                    var workbook = package.Workbook;
                    //..In Excel file you can add more than one ExcelSheet but here we get first excelsheet…
                    var worksheet = workbook.Worksheets.First();
                    //...here we have decrement of 1 count because do not calculate header row...
                    int noOfRow = worksheet.Dimension.End.Row;
                    //... we know our data start from second row so set row 2..
                    int row = 1;
                    for (int a = 0; a < noOfRow; a++)
                    {
                        var number = worksheet.GetValue(row, 1);
                        var name = worksheet.GetValue(row, 2);
                        var money = worksheet.GetValue(row, 3);
                        var field01 = worksheet.GetValue(row, 4);
                        var field02 = worksheet.GetValue(row, 5);
                        var field03 = worksheet.GetValue(row, 6);
                        var sms = new SmsInfor
                        {
                            Number = number == null ? null : number.ToString(),
                            Name = name == null ? null : name.ToString(),
                            Money = money == null ? null : money.ToString(),
                            Field01 = field01 == null ? null : field01.ToString(),
                            Field02 = field02 == null ? null : field02.ToString(),
                            Field03 = field03 == null ? null : field03.ToString(),
                        };

                        smsDatas.Add(sms);
                        row++;
                    }
                }

                btnDeleteFile.Visible = true;
                dataGridView1.DataSource = null;
                dataGridView1.Refresh();
                dataGridView1.DataSource = smsDatas;
            }
            else return;
        }

        private void btnDeleteFile_Click(object sender, EventArgs e)
        {
            smsDatas = new List<SmsInfor>();
            dataGridView1.DataSource = smsDatas;
            btnDeleteFile.Visible = false;
            hasFileExcel = false;
        }

        private void DisableConnect()
        {
            try
            {
                objclsSMS.ClosePort(this.port);
                lbl_status.Text = "Đã ngắt kết nối đến Modem GMS port " + txt_port.Text;
                btn_connect.Text = "Kết nối";


                lbl_status.ForeColor = Color.Red;

            }
            catch (Exception ex)
            {
                Debug.Write(ex.Message);
            }
        }

        private void NotifiStatusConnect(bool isNotConnect, string port)
        {
            if (!isNotConnect)
            {
                txt_port.Text = "";
                btn_connect.Text = "Kết nối";


                lbl_status.Text = "Trạng thái: Không tìm thấy thiết bị Modem GMS!";
                lbl_status.ForeColor = Color.Red;

                isModemConnect = false;

            }
            else
            {
                txt_port.Text = port;
                btn_connect.Text = "Ngắt kết nối";

                lbl_status.ForeColor = Color.Blue;
                lbl_status.Text = "Trạng thái: Đang kết nối Modem GMS cổng " + port;

                isModemConnect = true;
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (File.Exists(filePath))
            {
                StreamWriter sw = new StreamWriter(filePath, false);
                sw.WriteLine(txt_message.Text);
                sw.Close();
            }
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            if (File.Exists(filePath))
            {
                StreamReader sr = new StreamReader(filePath);
                txt_message.Text = sr.ReadToEnd();
                sr.Close();
            }
        }
    }
}
