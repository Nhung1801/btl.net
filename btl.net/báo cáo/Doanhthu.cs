using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Printing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;


namespace baocao
{
    public partial class Doanhthu : Form
    {
        public Doanhthu()
        {
            InitializeComponent();
        }

        private void ClearDataGridView()
        {
            dataGridView.DataSource = null;
            dataGridView.Rows.Clear();
            dataGridView.Columns.Clear();
        }
        private void ExportExcel(string path)
        {
            try
            {
                // Tạo một đối tượng Excel mới  
                Excel.Application application = new Excel.Application();
                application.Application.Workbooks.Add(Type.Missing);

                // Đặt tiêu đề cho bảng tính  
                Excel.Worksheet worksheet = (Excel.Worksheet)application.ActiveSheet;
                worksheet.Name = "Báo Cáo Doanh Thu";

                // Tính tổng doanh thu  
                decimal totalRevenue = 0;

                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    //  doanh thu là cột thứ 4 (index 3)  
                    totalRevenue += Convert.ToDecimal(dataGridView.Rows[i].Cells[3].Value);
                }

                // Thêm tiêu đề cho báo cáo  
                worksheet.Cells[1, 1] = "BÁO CÁO DOANH THU";
                worksheet.Cells[2, 1] = "Ngày lập báo cáo: " + DateTime.Now.ToString("dd/MM/yyyy");
                worksheet.Cells[3, 1] = "Tổng doanh thu: " + totalRevenue.ToString("C0");

                // Định dạng cho tiêu đề  
                Excel.Range titleRange = worksheet.Range["A1:C1"];
                titleRange.Merge();
                titleRange.Font.Size = 16;
                titleRange.Font.Bold = true;
                titleRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                // Đặt tiêu đề cột  
                for (int i = 0; i < dataGridView.Columns.Count; i++)
                {
                    worksheet.Cells[5, i + 1] = dataGridView.Columns[i].HeaderText; // Bắt đầu từ dòng 5  
                    worksheet.Cells[5, i + 1].Font.Bold = true; // Làm đậm tiêu đề  
                    worksheet.Cells[5, i + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray); // Đổi màu nền  
                }

                // Thêm dữ liệu từ DataGridView vào Excel  
                for (int i = 0; i < dataGridView.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 6, j + 1] = dataGridView.Rows[i].Cells[j].Value; // Bắt đầu từ dòng 6  
                    }
                }

                // Định dạng cho cột doanh thu  
                Excel.Range revenueRange = worksheet.Range["D6:D" + (dataGridView.Rows.Count + 5)]; // Thay D bằng cột của thông tin doanh thu  
                revenueRange.NumberFormat = "#,##0"; // Định dạng số  

                // Tự động điều chỉnh kích thước cột  
                application.Columns.AutoFit();

                // Lưu tệp Excel  
                application.ActiveWorkbook.SaveCopyAs(path);
                application.ActiveWorkbook.Saved = true;
                application.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi xảy ra trong quá trình xuất file: " + ex.Message);
            }
        }
        private void btnThoat_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show(
        "Bạn có chắc muốn thoát không?",
        "Xác nhận thoát",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                this.Close(); // Đóng form hiện tại  
                              // Hoặc Application.Exit(); để thoát toàn ứng dụng  
            }
        }

        private void Validate_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar >= '0' && e.KeyChar <= '9') || e.KeyChar == (char)Keys.Back)
                e.Handled = false;
            else
                e.Handled = true;
        }

        private void btnIn_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Export Excel";
            saveFileDialog.Filter = "Excel (*.xlsx)|*.xlsx|Excel 2003 (*.xlsx)|*.xls";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    ExportExcel(saveFileDialog.FileName);
                    MessageBox.Show("Xuất file thành công!");
                }

                catch
                {
                    MessageBox.Show("Xuất file thất bại!");
                }
            }
        }

        private void rdoKhoang_CheckedChanged_1(object sender, EventArgs e)
        {
            mskNgay.Enabled = false;
            mskNgay.Text = "";
            mskTu.Enabled = true;
            mskDen.Enabled = true;
            txtdoanhthu.Text = "";
            ClearDataGridView();
        }

        private void rdoNgay_CheckedChanged(object sender, EventArgs e)
        {
            mskNgay.Enabled = true;
            mskTu.Enabled = false;
            mskDen.Enabled = false;
            mskTu.Text = "";
            mskDen.Text = "";
            txtdoanhthu.Text = "";
            ClearDataGridView();
        }

        private void Doanhthu_Load(object sender, EventArgs e)
        {
            btnXem.Enabled = true;
            btnIn.Enabled = true;
            btnThoat.Enabled = true;
            txtdoanhthu.Enabled = true;
            rdoNgay.Checked = false;
            rdoKhoang.Checked = false;
        }

        private void btnXem_Click_1(object sender, EventArgs e)
        {
            string connString = "Data Source=DESKTOP-IK88KCU;Initial Catalog=Qlcuahangquanao;Integrated Security=True;Encrypt=False";
            string sql = "";

            if (rdoNgay.Checked)
            {
                DateTime selectedDate;
                if (DateTime.TryParse(mskNgay.Text, out selectedDate))
                {
                    sql = $"SELECT sp.MaQuanAo, sp.TenQuanAo, SUM(ct.SoLuong) AS soluongbanra, SUM(ct.ThanhTien) AS doanhthu " +
                      $"FROM HoaDonBan hdb " +
                      $"JOIN ChiTietHDBan ct ON hdb.SoHDB = ct.SoHDB " +
                      $"JOIN SanPham sp ON sp.MaQuanAo = ct.MaQuanAo " +
                      $"WHERE hdb.NgayBan = '{selectedDate.ToString("yyyy-MM-dd")}' " +
                      $"GROUP BY sp.MaQuanAo, sp.TenQuanAo, hdb.NgayBan " +
                      $"ORDER BY hdb.NgayBan, sp.MaQuanAo";

                }
                else
                {
                    MessageBox.Show("Ngày không hợp lệ, bạn vui lòng nhập lại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
            else if (rdoKhoang.Checked)
            {
                mskNgay.Enabled = false;
                DateTime fromDate, toDate;
                if (DateTime.TryParse(mskTu.Text, out fromDate) && DateTime.TryParse(mskDen.Text, out toDate))
                {


                    sql = $"SELECT sp.MaQuanAo,sp.TenQuanAo,SUM(ct.SoLuong) AS soluongbanra,SUM(ct.ThanhTien) AS doanhthu  FROM  HoaDonban hdb  " +
                    $"JOIN    ChiTietHDBan ct ON hdb.SoHDB = ct.SoHDB " +
                    $"JOIN    SanPham sp ON sp.MaQuanAo = ct.MaQuanAo   " +
                    $"WHERE    hdb.NgayBan BETWEEN '{fromDate.ToString("yyyy-MM-dd")}' AND '{toDate.ToString("yyyy-MM-dd")}'   " +
                    $"GROUP BY    sp.MaQuanAo, sp.TenQuanAo  ORDER BY  sp.MaQuanAo";
                }
                else
                {
                    MessageBox.Show("Ngày không hợp lệ, bạn vui lòng nhập lại!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
            else
            {
                MessageBox.Show("Vui lòng chọn một tùy chọn.", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //Đổ dữ liệu vào DataGridView
            using (SqlConnection connection = new SqlConnection(connString))
            {
                SqlCommand command = new SqlCommand(sql, connection);
                SqlDataAdapter adapter = new SqlDataAdapter(command);
                DataTable table = new DataTable();
                adapter.Fill(table);
                dataGridView.DataSource = table;
            dataGridView.DataSource = table;
                dataGridView.Columns[0].HeaderText = "Mã quần áo";
                dataGridView.Columns[1].HeaderText = "Tên quần áo";
                dataGridView.Columns[2].HeaderText = "Số lượng bán ra";
                dataGridView.Columns[3].HeaderText = "Doanh thu";
                dataGridView.Columns[0].Width = 180;
                dataGridView.Columns[1].Width = 300;
                dataGridView.Columns[2].Width = 180;
                dataGridView.Columns[3].Width = 200;

                decimal totalRevenue = 0;
                foreach (DataRow row in table.Rows)
                {
                    totalRevenue += Convert.ToDecimal(row["doanhthu"]);
                }

                // Hiển thị tổng doanh thu
                txtdoanhthu.Enabled = true;
                NumberFormatInfo nfi = new CultureInfo("vi-VN", false).NumberFormat;
                nfi.CurrencySymbol = "₫";
                txtdoanhthu.Text = totalRevenue.ToString("C", nfi);

            }

        }
    }

}
