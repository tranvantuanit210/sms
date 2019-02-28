namespace SMSManager
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txt_port = new System.Windows.Forms.TextBox();
            this.txt_phone = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txt_message = new System.Windows.Forms.TextBox();
            this.btn_send = new System.Windows.Forms.Button();
            this.lbl_status = new System.Windows.Forms.Label();
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.btnDeleteFile = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.label4 = new System.Windows.Forms.Label();
            this.txtTimeout = new System.Windows.Forms.NumericUpDown();
            this.label5 = new System.Windows.Forms.Label();
            this.btn_connect = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtTimeout)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(28, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(67, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "COM PORT:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(22, 78);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(73, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Số điện thoại:";
            // 
            // txt_port
            // 
            this.txt_port.Location = new System.Drawing.Point(101, 21);
            this.txt_port.Name = "txt_port";
            this.txt_port.Size = new System.Drawing.Size(147, 20);
            this.txt_port.TabIndex = 2;
            // 
            // txt_phone
            // 
            this.txt_phone.Location = new System.Drawing.Point(101, 78);
            this.txt_phone.Multiline = true;
            this.txt_phone.Name = "txt_phone";
            this.txt_phone.Size = new System.Drawing.Size(262, 20);
            this.txt_phone.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(42, 112);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Nội dung:";
            // 
            // txt_message
            // 
            this.txt_message.Location = new System.Drawing.Point(101, 112);
            this.txt_message.Multiline = true;
            this.txt_message.Name = "txt_message";
            this.txt_message.Size = new System.Drawing.Size(262, 160);
            this.txt_message.TabIndex = 5;
            // 
            // btn_send
            // 
            this.btn_send.Location = new System.Drawing.Point(288, 278);
            this.btn_send.Name = "btn_send";
            this.btn_send.Size = new System.Drawing.Size(75, 23);
            this.btn_send.TabIndex = 7;
            this.btn_send.Text = "Gởi tin nhắn";
            this.btn_send.UseVisualStyleBackColor = true;
            this.btn_send.Click += new System.EventHandler(this.btn_send_Click);
            // 
            // lbl_status
            // 
            this.lbl_status.AutoSize = true;
            this.lbl_status.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_status.ForeColor = System.Drawing.Color.Red;
            this.lbl_status.Location = new System.Drawing.Point(98, 310);
            this.lbl_status.Name = "lbl_status";
            this.lbl_status.Size = new System.Drawing.Size(72, 15);
            this.lbl_status.TabIndex = 15;
            this.lbl_status.Text = "Trạng thái";
            // 
            // btnSelectFile
            // 
            this.btnSelectFile.Location = new System.Drawing.Point(101, 278);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(69, 23);
            this.btnSelectFile.TabIndex = 16;
            this.btnSelectFile.Text = "Chọn File";
            this.btnSelectFile.UseVisualStyleBackColor = true;
            this.btnSelectFile.Click += new System.EventHandler(this.btnSelectFile_Click);
            // 
            // btnDeleteFile
            // 
            this.btnDeleteFile.Location = new System.Drawing.Point(176, 278);
            this.btnDeleteFile.Name = "btnDeleteFile";
            this.btnDeleteFile.Size = new System.Drawing.Size(106, 23);
            this.btnDeleteFile.TabIndex = 17;
            this.btnDeleteFile.Text = "Xóa File";
            this.btnDeleteFile.UseVisualStyleBackColor = true;
            this.btnDeleteFile.Visible = false;
            this.btnDeleteFile.Click += new System.EventHandler(this.btnDeleteFile_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(394, 19);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(447, 282);
            this.dataGridView1.TabIndex = 18;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(22, 49);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(75, 13);
            this.label4.TabIndex = 1;
            this.label4.Text = "Thời gian chờ:";
            // 
            // txtTimeout
            // 
            this.txtTimeout.Location = new System.Drawing.Point(101, 47);
            this.txtTimeout.Minimum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.txtTimeout.Name = "txtTimeout";
            this.txtTimeout.Size = new System.Drawing.Size(115, 20);
            this.txtTimeout.TabIndex = 19;
            this.txtTimeout.Value = new decimal(new int[] {
            10,
            0,
            0,
            0});
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(222, 49);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(26, 13);
            this.label5.TabIndex = 20;
            this.label5.Text = "giây";
            // 
            // btn_connect
            // 
            this.btn_connect.Location = new System.Drawing.Point(265, 18);
            this.btn_connect.Name = "btn_connect";
            this.btn_connect.Size = new System.Drawing.Size(98, 23);
            this.btn_connect.TabIndex = 6;
            this.btn_connect.Text = "Kết Nối";
            this.btn_connect.UseVisualStyleBackColor = true;
            this.btn_connect.Click += new System.EventHandler(this.btn_connect_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(855, 346);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtTimeout);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.btnDeleteFile);
            this.Controls.Add(this.btnSelectFile);
            this.Controls.Add(this.lbl_status);
            this.Controls.Add(this.btn_send);
            this.Controls.Add(this.btn_connect);
            this.Controls.Add(this.txt_message);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txt_phone);
            this.Controls.Add(this.txt_port);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "SMS";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtTimeout)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txt_port;
        private System.Windows.Forms.TextBox txt_phone;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txt_message;
        private System.Windows.Forms.Button btn_send;
        private System.Windows.Forms.Label lbl_status;
        private System.Windows.Forms.Button btnSelectFile;
        private System.Windows.Forms.Button btnDeleteFile;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.NumericUpDown txtTimeout;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btn_connect;
    }
}

