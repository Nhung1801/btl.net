namespace baocao
{
    partial class Hieusuat
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
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea7 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend7 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series7 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea8 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend8 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series8 = new System.Windows.Forms.DataVisualization.Charting.Series();
            this.label1 = new System.Windows.Forms.Label();
            this.gpbThoigian = new System.Windows.Forms.GroupBox();
            this.dtpKT = new System.Windows.Forms.DateTimePicker();
            this.dtpBD = new System.Windows.Forms.DateTimePicker();
            this.lblNam2 = new System.Windows.Forms.Label();
            this.lblNam1 = new System.Windows.Forms.Label();
            this.cboNamKT = new System.Windows.Forms.ComboBox();
            this.cboKT = new System.Windows.Forms.ComboBox();
            this.lblKT = new System.Windows.Forms.Label();
            this.cboNamBD = new System.Windows.Forms.ComboBox();
            this.cboBD = new System.Windows.Forms.ComboBox();
            this.lblBD = new System.Windows.Forms.Label();
            this.btnHienthi = new System.Windows.Forms.Button();
            this.cboThoigian = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.dgridXephang = new System.Windows.Forms.DataGridView();
            this.dgridBaocao = new System.Windows.Forms.DataGridView();
            this.btnLammoi = new System.Windows.Forms.Button();
            this.btnXuatexcel = new System.Windows.Forms.Button();
            this.btnDong = new System.Windows.Forms.Button();
            this.chart = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.chart1 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.label4 = new System.Windows.Forms.Label();
            this.gpbThoigian.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgridXephang)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgridBaocao)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chart)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(58, 465);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(231, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "Nhân viên có doanh thu cao nhất";
            // 
            // gpbThoigian
            // 
            this.gpbThoigian.Controls.Add(this.dtpKT);
            this.gpbThoigian.Controls.Add(this.dtpBD);
            this.gpbThoigian.Controls.Add(this.lblNam2);
            this.gpbThoigian.Controls.Add(this.lblNam1);
            this.gpbThoigian.Controls.Add(this.cboNamKT);
            this.gpbThoigian.Controls.Add(this.cboKT);
            this.gpbThoigian.Controls.Add(this.lblKT);
            this.gpbThoigian.Controls.Add(this.cboNamBD);
            this.gpbThoigian.Controls.Add(this.cboBD);
            this.gpbThoigian.Controls.Add(this.lblBD);
            this.gpbThoigian.Controls.Add(this.btnHienthi);
            this.gpbThoigian.Controls.Add(this.cboThoigian);
            this.gpbThoigian.Controls.Add(this.label3);
            this.gpbThoigian.Location = new System.Drawing.Point(61, 94);
            this.gpbThoigian.Name = "gpbThoigian";
            this.gpbThoigian.Size = new System.Drawing.Size(905, 99);
            this.gpbThoigian.TabIndex = 1;
            this.gpbThoigian.TabStop = false;
            this.gpbThoigian.Text = "Chọn thời gian";
            // 
            // dtpKT
            // 
            this.dtpKT.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpKT.Location = new System.Drawing.Point(406, 65);
            this.dtpKT.Name = "dtpKT";
            this.dtpKT.Size = new System.Drawing.Size(142, 22);
            this.dtpKT.TabIndex = 17;
            this.dtpKT.ValueChanged += new System.EventHandler(this.dtpKT_ValueChanged);
            // 
            // dtpBD
            // 
            this.dtpBD.CustomFormat = "";
            this.dtpBD.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpBD.Location = new System.Drawing.Point(405, 15);
            this.dtpBD.Name = "dtpBD";
            this.dtpBD.Size = new System.Drawing.Size(142, 22);
            this.dtpBD.TabIndex = 16;
            // 
            // lblNam2
            // 
            this.lblNam2.AutoSize = true;
            this.lblNam2.Location = new System.Drawing.Point(575, 71);
            this.lblNam2.Name = "lblNam2";
            this.lblNam2.Size = new System.Drawing.Size(47, 16);
            this.lblNam2.TabIndex = 15;
            this.lblNam2.Text = "Năm :";
            // 
            // lblNam1
            // 
            this.lblNam1.AutoSize = true;
            this.lblNam1.Location = new System.Drawing.Point(575, 18);
            this.lblNam1.Name = "lblNam1";
            this.lblNam1.Size = new System.Drawing.Size(47, 16);
            this.lblNam1.TabIndex = 14;
            this.lblNam1.Text = "Năm :";
            // 
            // cboNamKT
            // 
            this.cboNamKT.FormattingEnabled = true;
            this.cboNamKT.Location = new System.Drawing.Point(631, 63);
            this.cboNamKT.Name = "cboNamKT";
            this.cboNamKT.Size = new System.Drawing.Size(98, 24);
            this.cboNamKT.TabIndex = 13;
            // 
            // cboKT
            // 
            this.cboKT.FormattingEnabled = true;
            this.cboKT.Location = new System.Drawing.Point(405, 63);
            this.cboKT.Name = "cboKT";
            this.cboKT.Size = new System.Drawing.Size(143, 24);
            this.cboKT.TabIndex = 12;
            // 
            // lblKT
            // 
            this.lblKT.AutoSize = true;
            this.lblKT.Location = new System.Drawing.Point(361, 63);
            this.lblKT.Name = "lblKT";
            this.lblKT.Size = new System.Drawing.Size(38, 16);
            this.lblKT.TabIndex = 10;
            this.lblKT.Text = "Đến:";
            // 
            // cboNamBD
            // 
            this.cboNamBD.FormattingEnabled = true;
            this.cboNamBD.Location = new System.Drawing.Point(631, 15);
            this.cboNamBD.Name = "cboNamBD";
            this.cboNamBD.Size = new System.Drawing.Size(98, 24);
            this.cboNamBD.TabIndex = 9;
            // 
            // cboBD
            // 
            this.cboBD.FormattingEnabled = true;
            this.cboBD.Location = new System.Drawing.Point(405, 15);
            this.cboBD.Name = "cboBD";
            this.cboBD.Size = new System.Drawing.Size(143, 24);
            this.cboBD.TabIndex = 8;
            // 
            // lblBD
            // 
            this.lblBD.AutoSize = true;
            this.lblBD.Location = new System.Drawing.Point(366, 15);
            this.lblBD.Name = "lblBD";
            this.lblBD.Size = new System.Drawing.Size(33, 16);
            this.lblBD.TabIndex = 6;
            this.lblBD.Text = "Từ :";
            // 
            // btnHienthi
            // 
            this.btnHienthi.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnHienthi.Location = new System.Drawing.Point(759, 21);
            this.btnHienthi.Name = "btnHienthi";
            this.btnHienthi.Size = new System.Drawing.Size(121, 56);
            this.btnHienthi.TabIndex = 5;
            this.btnHienthi.Text = "Hiển thị";
            this.btnHienthi.UseVisualStyleBackColor = false;
            this.btnHienthi.Click += new System.EventHandler(this.btnHienthi_Click);
            // 
            // cboThoigian
            // 
            this.cboThoigian.FormattingEnabled = true;
            this.cboThoigian.Location = new System.Drawing.Point(138, 43);
            this.cboThoigian.Name = "cboThoigian";
            this.cboThoigian.Size = new System.Drawing.Size(166, 24);
            this.cboThoigian.TabIndex = 4;
            this.cboThoigian.SelectedIndexChanged += new System.EventHandler(this.cboThoigian_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(22, 46);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(110, 16);
            this.label3.TabIndex = 3;
            this.label3.Text = "Thời gian bán :";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(702, 341);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(50, 16);
            this.label2.TabIndex = 2;
            this.label2.Text = "label2";
            // 
            // dgridXephang
            // 
            this.dgridXephang.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgridXephang.Location = new System.Drawing.Point(61, 484);
            this.dgridXephang.Name = "dgridXephang";
            this.dgridXephang.RowHeadersWidth = 51;
            this.dgridXephang.RowTemplate.Height = 24;
            this.dgridXephang.Size = new System.Drawing.Size(976, 125);
            this.dgridXephang.TabIndex = 4;
            // 
            // dgridBaocao
            // 
            this.dgridBaocao.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgridBaocao.Location = new System.Drawing.Point(61, 224);
            this.dgridBaocao.Name = "dgridBaocao";
            this.dgridBaocao.RowHeadersWidth = 51;
            this.dgridBaocao.RowTemplate.Height = 24;
            this.dgridBaocao.Size = new System.Drawing.Size(976, 190);
            this.dgridBaocao.TabIndex = 5;
            // 
            // btnLammoi
            // 
            this.btnLammoi.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnLammoi.Location = new System.Drawing.Point(324, 630);
            this.btnLammoi.Name = "btnLammoi";
            this.btnLammoi.Size = new System.Drawing.Size(157, 48);
            this.btnLammoi.TabIndex = 6;
            this.btnLammoi.Text = "Làm mới";
            this.btnLammoi.UseVisualStyleBackColor = false;
            this.btnLammoi.Click += new System.EventHandler(this.btnLammoi_Click);
            // 
            // btnXuatexcel
            // 
            this.btnXuatexcel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnXuatexcel.Location = new System.Drawing.Point(609, 630);
            this.btnXuatexcel.Name = "btnXuatexcel";
            this.btnXuatexcel.Size = new System.Drawing.Size(157, 48);
            this.btnXuatexcel.TabIndex = 7;
            this.btnXuatexcel.Text = "In báo cáo";
            this.btnXuatexcel.UseVisualStyleBackColor = false;
            this.btnXuatexcel.Click += new System.EventHandler(this.btnXuatexcel_Click);
            // 
            // btnDong
            // 
            this.btnDong.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnDong.Location = new System.Drawing.Point(894, 630);
            this.btnDong.Name = "btnDong";
            this.btnDong.Size = new System.Drawing.Size(157, 48);
            this.btnDong.TabIndex = 8;
            this.btnDong.Text = "Thoát";
            this.btnDong.UseVisualStyleBackColor = false;
            this.btnDong.Click += new System.EventHandler(this.btnDong_Click);
            // 
            // chart
            // 
            chartArea7.Name = "ChartArea1";
            this.chart.ChartAreas.Add(chartArea7);
            legend7.Name = "Legend1";
            this.chart.Legends.Add(legend7);
            this.chart.Location = new System.Drawing.Point(1095, 55);
            this.chart.Name = "chart";
            series7.ChartArea = "ChartArea1";
            series7.Legend = "Legend1";
            series7.Name = "Series1";
            this.chart.Series.Add(series7);
            this.chart.Size = new System.Drawing.Size(334, 206);
            this.chart.TabIndex = 9;
            this.chart.Text = "Tổng số đơn hàng đã bán";
            // 
            // chart1
            // 
            chartArea8.Name = "ChartArea1";
            this.chart1.ChartAreas.Add(chartArea8);
            legend8.Name = "Legend1";
            this.chart1.Legends.Add(legend8);
            this.chart1.Location = new System.Drawing.Point(1095, 366);
            this.chart1.Name = "chart1";
            series8.ChartArea = "ChartArea1";
            series8.Legend = "Legend1";
            series8.Name = "Series1";
            this.chart1.Series.Add(series8);
            this.chart1.Size = new System.Drawing.Size(334, 206);
            this.chart1.TabIndex = 10;
            this.chart1.Text = "Tổng số doanh thu";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.White;
            this.label4.Font = new System.Drawing.Font("Microsoft YaHei UI", 25.8F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.Red;
            this.label4.Location = new System.Drawing.Point(465, 20);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(501, 57);
            this.label4.TabIndex = 11;
            this.label4.Text = "BÁO CÁO NHÂN VIÊN";
            // 
            // Hieusuat
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1455, 699);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.chart1);
            this.Controls.Add(this.chart);
            this.Controls.Add(this.btnDong);
            this.Controls.Add(this.btnXuatexcel);
            this.Controls.Add(this.btnLammoi);
            this.Controls.Add(this.dgridBaocao);
            this.Controls.Add(this.dgridXephang);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.gpbThoigian);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "Hieusuat";
            this.Text = "Hieusuat";
            this.Load += new System.EventHandler(this.Hieusuat_Load);
            this.gpbThoigian.ResumeLayout(false);
            this.gpbThoigian.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgridXephang)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgridBaocao)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chart)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox gpbThoigian;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DataGridView dgridXephang;
        private System.Windows.Forms.DataGridView dgridBaocao;
        private System.Windows.Forms.Button btnHienthi;
        private System.Windows.Forms.ComboBox cboThoigian;
        private System.Windows.Forms.Button btnLammoi;
        private System.Windows.Forms.Button btnXuatexcel;
        private System.Windows.Forms.Button btnDong;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lblBD;
        private System.Windows.Forms.ComboBox cboNamBD;
        private System.Windows.Forms.ComboBox cboBD;
        private System.Windows.Forms.ComboBox cboNamKT;
        private System.Windows.Forms.ComboBox cboKT;
        private System.Windows.Forms.Label lblKT;
        private System.Windows.Forms.Label lblNam1;
        private System.Windows.Forms.Label lblNam2;
        private System.Windows.Forms.DateTimePicker dtpKT;
        private System.Windows.Forms.DateTimePicker dtpBD;
    }
}