using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using COMExcel = Microsoft.Office.Interop.Excel;

namespace baocao
{
    public partial class Banhang: Form
    {
        public Banhang()
        {
            InitializeComponent();
        }


        private void Banhang_Load(object sender, EventArgs e)
        {
            Load_dataGridView();
            Load_dataGridView1();
            LoadComboBoxSoHDB();

        }
        DataTable tblCTHDB;
        DataTable tblCTHDB1;
        private void Load_dataGridView()
        {
            string sql;
            sql = "SELECT hdb.SoHDB, mt.MaQuanAo, mt.TenQuanAo, cthdb.SoLuong, mt.DonGiaBan, cthdb.ThanhTien, hdb.NgayBan, kh.TenKhach, nv.tenNV FROM HoaDonBan hdb JOIN ChiTietHDBan cthdb ON hdb.SoHDB = cthdb.SoHDB JOIN SanPham mt ON cthdb.MaQuanAo = mt.MaQuanAo JOIN KhachHang kh ON hdb.MaKhach = kh.MaKhach JOIN NhanVien nv ON hdb.MaNV = nv.MaNV";
            tblCTHDB = function.GetDataToTable(sql);
            dataGridView.DataSource = tblCTHDB;

            if (dataGridView.Columns.Count >= 9)
            {
                dataGridView.Columns[0].HeaderText = "Mã hóa đơn";
                dataGridView.Columns[0].Width = 100;
                dataGridView.Columns[1].HeaderText = "Mã quần áo";
                dataGridView.Columns[1].Width = 50;
                dataGridView.Columns[2].HeaderText = "Tên quần áo";
                dataGridView.Columns[2].Width = 200;
                dataGridView.Columns[3].HeaderText = "Số lượng bán";
                dataGridView.Columns[3].Width = 50;
                dataGridView.Columns[4].HeaderText = "Đơn giá bán";
                dataGridView.Columns[4].Width = 90;
                dataGridView.Columns[5].HeaderText = "Thành tiền";
                dataGridView.Columns[5].Width = 90;
                dataGridView.Columns[6].HeaderText = "Ngày bán";
                dataGridView.Columns[6].Width = 100;
                dataGridView.Columns[7].HeaderText = "Tên khách hàng";
                dataGridView.Columns[7].Width = 200;
                dataGridView.Columns[8].HeaderText = "Tên nhân viên";
                dataGridView.Columns[8].Width = 200;
            }

            dataGridView.AllowUserToAddRows = false;
            dataGridView.EditMode = DataGridViewEditMode.EditProgrammatically;
        }
        private void Load_dataGridView1()
        {
            string sql;
            sql = "SELECT mt.TenQuanAo as 'Tên quần áo', SUM(cthdb.SoLuong) as 'Tổng số lượng bán' FROM HoaDonBan hdb JOIN ChiTietHDBan cthdb ON hdb.SoHDB = cthdb.SoHDB JOIN SanPham mt ON cthdb.MaQuanAo = mt.MaQuanAo GROUP BY mt.TenQuanAo";
            tblCTHDB1 = function.GetDataToTable(sql);
            dataGridView1.DataSource = tblCTHDB1;

            dataGridView1.Columns[0].HeaderText = "Tên quần áo";
            dataGridView1.Columns[0].Width = 200;
            dataGridView1.Columns[1].HeaderText = "Tổng số lượng bán";
            dataGridView1.Columns[1].Width = 50;

            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.EditMode = DataGridViewEditMode.EditProgrammatically;
        }

        private void btnIn_Click(object sender, EventArgs e)
        {
            COMExcel.Application exApp = new COMExcel.Application();
            COMExcel.Workbook exBook = exApp.Workbooks.Add(COMExcel.XlWBATemplate.xlWBATWorksheet);
            COMExcel.Worksheet exSheet = (COMExcel.Worksheet)exBook.Worksheets[1];

            // Định dạng tiêu đề báo cáo
            COMExcel.Range exRange = (COMExcel.Range)exSheet.Cells[1, 1];
            exRange.Range["E10:G10"].Font.Size = 14;
            exRange.Range["E10:G10"].Font.Name = "Times New Roman";
            exRange.Range["E10:G10"].Font.Bold = true;
            exRange.Range["E10:G10"].Font.ColorIndex = 3; // Màu đỏ
            exRange.Range["E10:G10"].MergeCells = true;
            exRange.Range["E10:G10"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["E10:G10"].Value = "Danh sách Bán hàng";

            // Định dạng tiêu đề cột
            exRange.Range["A12:J12"].Font.Bold = true;
            exRange.Range["A12:J12"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange.Range["A12"].Value = "STT";
            exRange.Range["B12"].Value = "Số hóa đơn";
            exRange.Range["C12"].Value = "Tên sản phẩm";
            exRange.Range["D12"].Value = "Mã sản phẩm";
            exRange.Range["E12"].Value = "Số lượng bán";
            exRange.Range["F12"].Value = "Đơn giá bán";
            exRange.Range["G12"].Value = "Thành tiền";
            exRange.Range["H12"].Value = "Ngày bán";
            exRange.Range["I12"].Value = "Tên khách hàng";
            exRange.Range["J12"].Value = "Tên nhân viên bán";

            // Điền dữ liệu
            for (int row = 0; row < tblCTHDB.Rows.Count; row++)
            {
                ((COMExcel.Range)exSheet.Cells[row + 13, 1]).Value2 = row + 1; // STT
                for (int col = 0; col < tblCTHDB.Columns.Count; col++)
                {
                    if (tblCTHDB.Columns[col].ColumnName == "NgayBan")
                    {
                        DateTime ngayNhap = Convert.ToDateTime(tblCTHDB.Rows[row]["NgayBan"]);
                        ((COMExcel.Range)exSheet.Cells[row + 13, col + 2]).Value2 = ngayNhap.ToShortDateString();
                    }
                    else
                    {
                        ((COMExcel.Range)exSheet.Cells[row + 13, col + 2]).Value2 = tblCTHDB.Rows[row][col].ToString();
                    }
                }
            }

            // Điền dữ liệu vào phần tổng số lượng bán
            exRange = (COMExcel.Range)exSheet.Cells[1, 1]; // Bắt đầu từ cột N
            exRange.Range["N12:O12"].Font.Bold = true;
            exRange.Range["N12:O12"].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignCenter;
            exRange = (COMExcel.Range)exSheet.Cells[12, 14]; // Set exRange to cell N12
            exRange.Value2 = "Tên sản phẩm";
            exRange = (COMExcel.Range)exSheet.Cells[12, 15]; // Set exRange to cell O12
            exRange.Value2 = "Số lượng bán";
            for (int row = 0; row < tblCTHDB1.Rows.Count; row++)
            {
                ((COMExcel.Range)exSheet.Cells[row + 13, 14]).Value2 = tblCTHDB1.Rows[row][0].ToString();
                ((COMExcel.Range)exSheet.Cells[row + 13, 15]).Value2 = tblCTHDB1.Rows[row][1].ToString();
            }
            // Tính tổng thành tiền cột H (thứ 8), bắt đầu từ hàng 13 cho đến hết dữ liệu  
            int rowCount = tblCTHDB.Rows.Count;
            int startDataRow = 13; // bắt đầu từ dòng 13 trong Excel, tương ứng row = 0 trong bảng C#  

            // Đặt chữ "Tổng cộng" tại cột G, dòng sau cùng (rowCount + 13)  
            exSheet.Cells[rowCount + startDataRow, 6].Value2 = "Tổng cộng"; 

            // Tính tổng cột H trong bảng tblCTHDB và ghi vào cột H, cùng dòng với "Tổng cộng"  
            double totalThanhTien = 0;
            for (int i = 0; i < rowCount; i++)
            {
                double tt = 0;
                double.TryParse(tblCTHDB.Rows[i]["ThanhTien"].ToString(), out tt);
                totalThanhTien += tt;
            }
            exSheet.Cells[rowCount + startDataRow, 7].Value2 = totalThanhTien;

            // Căn phải cho tổng tiền  
            exSheet.Cells[rowCount + startDataRow, 7].HorizontalAlignment = COMExcel.XlHAlign.xlHAlignRight;

            // In đậm dòng tổng cộng  
            exRange.Range[$"G{rowCount + startDataRow}:H{rowCount + startDataRow}"].Font.Bold = true;

            // Cuối cùng autofit cột  
            exSheet.Columns.AutoFit();
            exSheet.Columns.AutoFit();
            exApp.Visible = true;
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

        // Hàm tải dữ liệu vào ComboBox số hóa đơn
        private void LoadComboBoxSoHDB()
        {
            //Sử dụng lại function.cs của bạn để lấy dữ liệu
            string sql = "SELECT DISTINCT SoHDB FROM HoaDonBan";
            cboSoHDB.DataSource = function.GetDataToTable(sql);
            cboSoHDB.DisplayMember = "SoHDB";
            cboSoHDB.ValueMember = "SoHDB";

            // Tải tất cả dữ liệu khi form hiển thị lần đầu
            LoadAllData();
        }

        private void cboSoHDB_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Kiểm tra xem có mục nào được chọn không
            if (cboSoHDB.SelectedValue != null)
            {
                string soHDB = cboSoHDB.SelectedValue.ToString();
                LoadDataBySoHDB(soHDB);
            }
        }

        // Hàm tải dữ liệu theo số hóa đơn
        private void LoadDataBySoHDB(string soHDB)
        {
            //Sử dụng lại function.cs của bạn để lấy dữ liệu
            string sqlChiTiet = $@"
        SELECT cthd.SoHDB, cthd.MaQuanAo, qa.TenQuanAo, cthd.SoLuong, qa.DonGiaBan, cthd.ThanhTien,
               hd.NgayBan, kh.TenKhach, nv.tenNV
        FROM ChiTietHDBan cthd
        INNER JOIN HoaDonBan hd ON cthd.SoHDB = hd.SoHDB
        INNER JOIN SanPham qa ON cthd.MaQuanAo = qa.MaQuanAo
        INNER JOIN KhachHang kh ON hd.MaKhach = kh.MaKhach
        INNER JOIN NhanVien nv ON hd.MaNV = nv.MaNV
        WHERE cthd.SoHDB = '{soHDB}'";

            dataGridView.DataSource = function.GetDataToTable(sqlChiTiet);

            string sqlSanLuong = $@"
        SELECT qa.TenQuanAo, SUM(cthd.SoLuong) AS TongSoLuongBan
        FROM ChiTietHDBan cthd
        INNER JOIN SanPham qa ON cthd.MaQuanAo = qa.MaQuanAo
        WHERE cthd.SoHDB = '{soHDB}'
        GROUP BY qa.TenQuanAo";

            dataGridView1.DataSource = function.GetDataToTable(sqlSanLuong);
        }

        // Hàm tải tất cả dữ liệu khi form được hiển thị
        private void LoadAllData()
        {
            //Sử dụng lại function.cs của bạn để lấy dữ liệu
            string sqlChiTiet = @"
        SELECT cthd.SoHDB, cthd.MaQuanAo, qa.TenQuanAo, cthd.SoLuong, qa.DonGiaBan, cthd.ThanhTien,
               hd.NgayBan, kh.TenKhach, nv.tenNV
        FROM ChiTietHDBan cthd
        INNER JOIN HoaDonBan hd ON cthd.SoHDB = hd.SoHDB
        INNER JOIN SanPham qa ON cthd.MaQuanAo = qa.MaQuanAo
        INNER JOIN KhachHang kh ON hd.MaKhach = kh.MaKhach
        INNER JOIN NhanVien nv ON hd.MaNV = nv.MaNV";

            dataGridView.DataSource = function.GetDataToTable(sqlChiTiet);

            string sqlSanLuong = @"
        SELECT qa.TenQuanAo, SUM(cthd.SoLuong) AS TongSoLuongBan
        FROM ChiTietHDBan cthd
        INNER JOIN SanPham qa ON cthd.MaQuanAo = qa.MaQuanAo
        GROUP BY qa.TenQuanAo";

            dataGridView1.DataSource = function.GetDataToTable(sqlSanLuong);
        }

    }
}


