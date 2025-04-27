using System;
using System.Data;
using System.Data.SqlClient;
using System.Runtime.Remoting.Contexts;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace baocao
{
    public class function
    {
        public static SqlConnection conn;  //Khai báo đối tượng kết nối
        public static string connString;   //Khai báo biến chứa chuỗi kết nối

        public static void Connect()
        {
            //Thiết lập giá trị cho chuỗi kết nối
            connString = "Data Source=DESKTOP-IK88KCU;Initial Catalog=Qlcuahangquanao;Integrated Security=True;Encrypt=False";
            conn = new SqlConnection();         		//Cấp phát đối tượng
            conn.ConnectionString = connString; 		//Kết nối
            conn.Open();                        		//Mở kết nối
        }

        public static DataTable GetDataToTable(string sql)
        {
            try
            {
                if (conn == null || conn.State == ConnectionState.Closed)
                    Connect(); // Đảm bảo đã mở kết nối  

                DataTable table = new DataTable();
                SqlDataAdapter Mydata = new SqlDataAdapter(sql, conn);
                Mydata.Fill(table);
                return table;
            }
            catch (SqlException ex)
            {
                MessageBox.Show(ex.Message); // Hiển thị thông báo lỗi  
                return null; // Hoặc xử lý tùy ý  
            }
        }



        // Kiểm tra ID của bản ghi
        public static bool CheckID(string query)
        {
            SqlDataAdapter data = new SqlDataAdapter(query, conn);
            DataTable table = new DataTable();
            data.Fill(table);
            if (table.Rows.Count > 0)
            {
                return true;
            }

            return false;
        }

        // Thực thi query SQL
        public static void RunSQL(string query)
        {
            SqlCommand command = new SqlCommand();
            command.Connection = conn;
            command.CommandText = query;
            try
            {
                command.ExecuteNonQuery();
            }
            catch (System.Exception e)
            {
                MessageBox.Show($"Đã xảy ra lỗi: {e}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            command.Dispose();
        }

        // Thực thi query SQL (xóa)
        public static void RunDeleteSQL(string query)
        {
            SqlCommand command = new SqlCommand();
            command.Connection = conn;
            command.CommandText = query;
            try
            {
                command.ExecuteNonQuery();
            }
            catch (System.Exception)
            {
                MessageBox.Show("Dữ liệu đang được sử dụng, không thể xóa!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            command.Dispose();
        }

        // Lấy dữ liệu từ một query SQL
        public static string GetFieldValues(string query)
        {
            string key = "";
            SqlCommand command = new SqlCommand(query, conn);
            SqlDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                key = reader.GetValue(0).ToString();
            }

            reader.Close();
            return key;
        }

        // Đổ dữ liệu vào ComboBox
        public static void FillCombo(string query, System.Windows.Forms.ComboBox comboBox, string value, string name)
        {
            SqlDataAdapter data = new SqlDataAdapter(query, conn);
            DataTable table = new DataTable();
            data.Fill(table);

            comboBox.DataSource = table;
            comboBox.ValueMember = value; // Trường giá trị
            comboBox.DisplayMember = name; // Trường hiển thị
        }

        public static void FillCombo1(string query, System.Windows.Forms.ComboBox comboBox, string value, string name)
        {
            SqlDataAdapter data = new SqlDataAdapter(query, conn);
            DataTable table = new DataTable();
            data.Fill(table);

            comboBox.DataSource = table;
            comboBox.ValueMember = value; // Trường giá trị
            comboBox.DisplayMember = value; // Trường hiển thị
        }
        // Đổ dữ liệu vào combo box với định dạng mã + tên
        /*public static void FillCombo1(string query, ComboBox comboBox, string value, string displayExpression)
        {
            // Create a SqlDataAdapter and DataTable
            SqlDataAdapter data = new SqlDataAdapter(query, conn);
            DataTable table = new DataTable();
            data.Fill(table);

            // Add a new "DisplayMember" column to the DataTable
            table.Columns.Add("DisplayMember", typeof(string)).Expression = displayExpression;

            // Bind the DataTable to the combobox
            comboBox.DataSource = table;

            // Set the ValueMember and DisplayMember properties
            comboBox.ValueMember = value;
            comboBox.DisplayMember = "DisplayMember";

            // Set the selected index to -1
            comboBox.SelectedIndex = -1;
        }*/
        // Hàm này tương tự như hàm FillCombo nhưng có thêm khả năng định dạng văn bản hiển thị trong combobox.

        // Hàm kiểm tra dữ liệu nhập vào có phải là date không
        public static bool IsDate(string date)
        {
            string[] parts = date.Split('/');
            if ((Convert.ToInt32(parts[0]) >= 1) && (Convert.ToInt32(parts[0]) <= 31) && (Convert.ToInt32(parts[1]) >= 1) && (Convert.ToInt32(parts[1]) <= 12) && (Convert.ToInt32(parts[2]) >= 1900))
                return true;
            else
                return false;
        }

        // Chuyển đổi định dạng ngày tháng năm từ dd/MM/yyyy thành MM/dd/yyyy
        public static string ConvertDateTime(string date)
        {
            string[] parts = date.Split('/');
            string dateTime = String.Format("{0}/{1}/{2}", parts[1], parts[0], parts[2]);
            return dateTime;
        }

        // Chuyển đổi giá tiền từ định dạng số sang định dạng chữ
        public static string ConvertNumericToText(string number)
        {
            int length, digit;
            string result = "";
            string[] numText;

            // Xóa các dấu "," nếu có
            number = number.Replace(",", "");
            numText = "không;một;hai;ba;bốn;năm;sáu;bảy;tám;chín".Split(';');
            length = number.Length - 1; // Trừ 1 vì thứ tự đi từ 0

            for (int i = 0; i <= length; i++)
            {
                digit = Convert.ToInt32(number.Substring(i, 1));
                result = result + " " + numText[digit];

                if (length == i) // Chữ số cuối cùng không cần xét tiếp
                    break;

                switch ((length - i) % 9)
                {
                    case 0:
                        result = result + " tỷ";
                        if (number.Substring(i + 1, 3) == "000")
                            i += 3;
                        if (number.Substring(i + 1, 3) == "000")
                            i += 3;
                        if (number.Substring(i + 1, 3) == "000")
                            i += 3;
                        break;
                    case 6:
                        result = result + " triệu";
                        if (number.Substring(i + 1, 3) == "000")
                            i += 3;
                        if (number.Substring(i + 1, 3) == "000")
                            i += 3;
                        break;
                    case 3:
                        result = result + " nghìn";
                        if (number.Substring(i + 1, 3) == "000")
                            i += 3;
                        break;
                    default:
                        switch ((length - i) % 3)
                        {
                            case 2:
                                result = result + " trăm";
                                break;
                            case 1:
                                result = result + " mươi";
                                break;
                        }

                        break;
                }
            }

            // Loại bỏ trường hợp x00
            result = result.Replace("không mươi không ", "");
            result = result.Replace("không mươi không", "");

            // Loại bỏ trường hợp 00x
            result = result.Replace("không mươi ", "linh ");

            // Loại bỏ trường hợp x0, x>=2
            result = result.Replace("mươi không", "mươi");

            // Fix trường hợp 10
            result = result.Replace("một mươi", "mười");

            // Fix trường hợp x4, x>=2
            result = result.Replace("mươi bốn", "mươi tư");

            // Fix trường hợp x04
            result = result.Replace("linh bốn", "linh tư");

            // Fix trường hợp x5, x>=2
            result = result.Replace("mươi năm", "mươi lăm");

            // Fix trường hợp x1, x>=2
            result = result.Replace("mươi một", "mươi mốt");

            // Fix trường hợp x15
            result = result.Replace("mười năm", "mười lăm");

            // Bỏ ký tự space
            result = result.Trim();

            // Viết hoa ký tự đầu tiên
            result = char.ToUpper(result[0]) + result.Substring(1) + " đồng";

            return result;
        }

        // Tạo ID dựa theo ngày và giờ cho params bất kì
        public static string CreateKey(string param)
        {
            string key = param;

            string[] dateParts;
            dateParts = DateTime.Now.ToShortDateString().Split('/');
            string date = String.Format("{0}{1}{2}", dateParts[0], dateParts[1], dateParts[2]);
            key = key + date;

            string[] timeParts;
            timeParts = DateTime.Now.ToLongTimeString().Split(':');
            if (timeParts[2].Substring(3, 2) == "PM")
                timeParts[0] = ConvertTimeTo24(timeParts[0]);
            if (timeParts[2].Substring(3, 2) == "AM")
                if (timeParts[0].Length == 1)
                    timeParts[0] = "0" + timeParts[0];
            // Xóa ký tự trắng và PM hoặc AM
            timeParts[2] = timeParts[2].Remove(2, 3);
            string time;
            time = String.Format("{0}{1}{2}", timeParts[0], timeParts[1], timeParts[2]);
            key = key + "_" + time;

            return key;
        }

        // Chuyển đổi thời gian từ 12h sang 24h
        public static string ConvertTimeTo24(string hour)
        {
            string hour24 = "";
            switch (hour)
            {
                case "1":
                    hour24 = "13";
                    break;
                case "2":
                    hour24 = "14";
                    break;
                case "3":
                    hour24 = "15";
                    break;
                case "4":
                    hour24 = "16";
                    break;
                case "5":
                    hour24 = "17";
                    break;
                case "6":
                    hour24 = "18";
                    break;
                case "7":
                    hour24 = "19";
                    break;
                case "8":
                    hour24 = "20";
                    break;
                case "9":
                    hour24 = "21";
                    break;
                case "10":
                    hour24 = "22";
                    break;
                case "11":
                    hour24 = "23";
                    break;
                case "12":
                    hour24 = "0";
                    break;
            }

            return hour24;
        }
    }
}
