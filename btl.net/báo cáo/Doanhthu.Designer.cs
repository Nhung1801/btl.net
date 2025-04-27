namespace baocao
{
    partial class Doanhthu
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.mskDen = new System.Windows.Forms.MaskedTextBox();
            this.mskTu = new System.Windows.Forms.MaskedTextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.mskNgay = new System.Windows.Forms.MaskedTextBox();
            this.rdoKhoang = new System.Windows.Forms.RadioButton();
            this.rdoNgay = new System.Windows.Forms.RadioButton();
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.label2 = new System.Windows.Forms.Label();
            this.txtdoanhthu = new System.Windows.Forms.TextBox();
            this.btnXem = new System.Windows.Forms.Button();
            this.btnThoat = new System.Windows.Forms.Button();
            this.btnIn = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.mskDen);
            this.groupBox1.Controls.Add(this.mskTu);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.mskNgay);
            this.groupBox1.Controls.Add(this.rdoKhoang);
            this.groupBox1.Controls.Add(this.rdoNgay);
            this.groupBox1.Location = new System.Drawing.Point(652, 126);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(599, 122);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Tìm kiếm";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(396, 84);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(34, 16);
            this.label3.TabIndex = 6;
            this.label3.Text = "Đến";
            // 
            // mskDen
            // 
            this.mskDen.Location = new System.Drawing.Point(459, 81);
            this.mskDen.Mask = "0000/00/00";
            this.mskDen.Name = "mskDen";
            this.mskDen.Size = new System.Drawing.Size(112, 22);
            this.mskDen.TabIndex = 5;
            this.mskDen.ValidatingType = typeof(System.DateTime);
            // 
            // mskTu
            // 
            this.mskTu.Location = new System.Drawing.Point(242, 81);
            this.mskTu.Mask = "0000/00/00";
            this.mskTu.Name = "mskTu";
            this.mskTu.Size = new System.Drawing.Size(112, 22);
            this.mskTu.TabIndex = 4;
            this.mskTu.ValidatingType = typeof(System.DateTime);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(202, 87);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(25, 16);
            this.label1.TabIndex = 3;
            this.label1.Text = "Từ";
            // 
            // mskNgay
            // 
            this.mskNgay.Location = new System.Drawing.Point(237, 35);
            this.mskNgay.Mask = "0000/00/00";
            this.mskNgay.Name = "mskNgay";
            this.mskNgay.Size = new System.Drawing.Size(112, 22);
            this.mskNgay.TabIndex = 2;
            this.mskNgay.ValidatingType = typeof(System.DateTime);
            // 
            // rdoKhoang
            // 
            this.rdoKhoang.AutoSize = true;
            this.rdoKhoang.Location = new System.Drawing.Point(74, 83);
            this.rdoKhoang.Name = "rdoKhoang";
            this.rdoKhoang.Size = new System.Drawing.Size(112, 20);
            this.rdoKhoang.TabIndex = 1;
            this.rdoKhoang.TabStop = true;
            this.rdoKhoang.Text = "Theo Tháng";
            this.rdoKhoang.UseVisualStyleBackColor = true;
            this.rdoKhoang.CheckedChanged += new System.EventHandler(this.rdoKhoang_CheckedChanged_1);
            // 
            // rdoNgay
            // 
            this.rdoNgay.AutoSize = true;
            this.rdoNgay.Location = new System.Drawing.Point(74, 38);
            this.rdoNgay.Name = "rdoNgay";
            this.rdoNgay.Size = new System.Drawing.Size(105, 20);
            this.rdoNgay.TabIndex = 0;
            this.rdoNgay.TabStop = true;
            this.rdoNgay.Text = "Theo Ngày";
            this.rdoNgay.UseVisualStyleBackColor = true;
            this.rdoNgay.CheckedChanged += new System.EventHandler(this.rdoNgay_CheckedChanged);
            // 
            // dataGridView
            // 
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView.Location = new System.Drawing.Point(110, 264);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.RowHeadersWidth = 51;
            this.dataGridView.RowTemplate.Height = 24;
            this.dataGridView.Size = new System.Drawing.Size(990, 266);
            this.dataGridView.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(154, 557);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(114, 16);
            this.label2.TabIndex = 2;
            this.label2.Text = "Tổng doanh thu";
            // 
            // txtdoanhthu
            // 
            this.txtdoanhthu.Location = new System.Drawing.Point(305, 551);
            this.txtdoanhthu.Name = "txtdoanhthu";
            this.txtdoanhthu.Size = new System.Drawing.Size(121, 22);
            this.txtdoanhthu.TabIndex = 3;
            // 
            // btnXem
            // 
            this.btnXem.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnXem.Location = new System.Drawing.Point(173, 607);
            this.btnXem.Name = "btnXem";
            this.btnXem.Size = new System.Drawing.Size(161, 49);
            this.btnXem.TabIndex = 4;
            this.btnXem.Text = "Hiển thị";
            this.btnXem.UseVisualStyleBackColor = false;
            this.btnXem.Click += new System.EventHandler(this.btnXem_Click_1);
            // 
            // btnThoat
            // 
            this.btnThoat.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnThoat.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnThoat.Location = new System.Drawing.Point(894, 607);
            this.btnThoat.Name = "btnThoat";
            this.btnThoat.Size = new System.Drawing.Size(151, 49);
            this.btnThoat.TabIndex = 5;
            this.btnThoat.Text = "Thoát";
            this.btnThoat.UseVisualStyleBackColor = false;
            // 
            // btnIn
            // 
            this.btnIn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.btnIn.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnIn.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.btnIn.Location = new System.Drawing.Point(525, 607);
            this.btnIn.Name = "btnIn";
            this.btnIn.Size = new System.Drawing.Size(166, 49);
            this.btnIn.TabIndex = 6;
            this.btnIn.Text = "In báo cáo";
            this.btnIn.UseVisualStyleBackColor = false;
            this.btnIn.Click += new System.EventHandler(this.btnIn_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.White;
            this.label4.Font = new System.Drawing.Font("Microsoft YaHei UI", 25.8F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.Red;
            this.label4.Location = new System.Drawing.Point(320, 40);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(523, 57);
            this.label4.TabIndex = 7;
            this.label4.Text = "BÁO CÁO DOANH THU";
            // 
            // Doanhthu
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1275, 692);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btnIn);
            this.Controls.Add(this.btnThoat);
            this.Controls.Add(this.btnXem);
            this.Controls.Add(this.txtdoanhthu);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dataGridView);
            this.Controls.Add(this.groupBox1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "Doanhthu";
            this.Text = "Doanhthu";
            this.Load += new System.EventHandler(this.Doanhthu_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.MaskedTextBox mskDen;
        private System.Windows.Forms.MaskedTextBox mskTu;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.MaskedTextBox mskNgay;
        private System.Windows.Forms.RadioButton rdoKhoang;
        private System.Windows.Forms.RadioButton rdoNgay;
        private System.Windows.Forms.DataGridView dataGridView;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtdoanhthu;
        private System.Windows.Forms.Button btnXem;
        private System.Windows.Forms.Button btnThoat;
        private System.Windows.Forms.Button btnIn;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
    }
}