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
using Excel = Microsoft.Office.Interop.Excel;
using Dapper;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.dataGridViewDatabase.AutoGenerateColumns = false;
        }
      
        private List<Order> listdata = null;
        private List<OrderDetail> listdetaildata = null;

        private string stringconnection = @"Data Source=LAPTOP-MVFVSECF\SQL2019EXPRESS;Initial Catalog=DB_IMS;Integrated Security=True";

        private List<Order> GetAllDataFromOrder()
        {
            IEnumerable<Order> listdata = null;

            try
            {
                using (var conn = new SqlConnection(stringconnection))
                {
                    listdata = conn.Query<Order>("SELECT Orderr.Nomor, convert(varchar, Orderr.Tanggal, 101) AS Tanggal, Orderr.Supplier, COUNT(OrderDetail.NomorUrut) AS TotalItem, Orderr.Keterangan FROM Orderr INNER JOIN OrderDetail ON Orderr.Nomor = OrderDetail.Nomor GROUP BY Orderr.Nomor, Orderr.Tanggal, Orderr.Supplier, Orderr.Keterangan");
                }
                //SqlCommand cmd = new SqlCommand("SELECT Orderr.Nomor, convert(varchar, Orderr.Tanggal, 101) AS Tanggal, Orderr.Supplier, COUNT(OrderDetail.NomorUrut) AS TotalItem, Orderr.Keterangan FROM Orderr INNER JOIN OrderDetail ON Orderr.Nomor = OrderDetail.Nomor GROUP BY Orderr.Nomor, Orderr.Tanggal, Orderr.Supplier, Orderr.Keterangan", conn);               
            }
            catch(Exception)
            {
                throw;
            }
            return listdata?.ToList() ?? null;
        }

        private List<OrderDetail> GetAllDataFromOrderDetail()
        {
            IEnumerable<OrderDetail> listdetaildata = null;
            try
            {
                using (var conn = new SqlConnection(stringconnection))
                {
                    listdetaildata = conn.Query<OrderDetail>("SELECT OrderDetail.Nomor, OrderDetail.KodeBarang, Barang.NamaBarang, OrderDetail.Quantity, Barang.Satuan FROM OrderDetail INNER JOIN Barang ON OrderDetail.KodeBarang = Barang.KodeBarang");
                }
               
            }
            catch (Exception)
            {
                throw;
            }
            return listdetaildata?.ToList() ?? null;
        }

        private void LoadData()
        {
            listdata = GetAllDataFromOrder();
            if (listdata != null)
            {
                this.dataGridViewDatabase.DataSource = listdata;
                this.dataGridViewDatabase.Columns[0].DataPropertyName = nameof(Order.Nomor);
                this.dataGridViewDatabase.Columns[1].DataPropertyName = nameof(Order.Tanggal);
                this.dataGridViewDatabase.Columns[2].DataPropertyName = nameof(Order.Supplier);
                this.dataGridViewDatabase.Columns[3].DataPropertyName = nameof(Order.TotalItem);
                this.dataGridViewDatabase.Columns[4].DataPropertyName = nameof(Order.Keterangan);
            }
            listdetaildata = GetAllDataFromOrderDetail();          
        }

        private void FrmPrint_LoadData(object sender, EventArgs e)
        {
            try
            {             
                LoadData();               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btn_Export_Click(object sender, EventArgs e)
        {          
            try
            {
                Excel.Application app = new Excel.Application();

                Excel.Workbook book = app.Workbooks.Add();

                Excel.Worksheet sheet = book.ActiveSheet as Excel.Worksheet;

                app.Visible = true;
                app.WindowState = Excel.XlWindowState.xlMaximized;

                //Head
                sheet.Cells[1, 1] = "Data Order";

                //subhead
                sheet.Cells[3, 1] = "Nomor";
                sheet.Cells[3, 2] = "Tanggal";
                sheet.Cells[3, 3] = "Supplier";

                sheet.Cells[4, 2] = "KodeBarang";
                sheet.Cells[4, 3] = "NamaBarang";
                sheet.Cells[4, 4] = "Quantity";
                sheet.Cells[4, 5] = "Satuan";

                //pemisalan i,j column untuk load ke excel
                //i = baris , j = kolum
                int i = 1;
                int j = 6;

                //load objdata ke excel sheet
                int count = 0;
                foreach (var order in listdata)
                {
                    sheet.Cells[j, i] = listdata[count].Nomor;
                    sheet.Cells[j, i + 1] = listdata[count].Tanggal;
                    sheet.Cells[j, i + 2] = listdata[count].Supplier;

                    //format kasi rapi
                    sheet.Cells[i, j].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    sheet.Cells[i, j + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    //setelah sudah satu baris maka lanjut
                    j++;
                    int HitungDetail = 0;

                    //isikan listdatadetail yang tadi sudah dibuat ke dalam cell 
                    foreach (var orderDetail in listdetaildata)
                    {
                        if (listdetaildata[HitungDetail].Nomor == listdata[count].Nomor)
                        {
                            sheet.Cells[j, i + 1] = listdetaildata[HitungDetail].KodeBarang;
                            sheet.Cells[j, i + 2] = listdetaildata[HitungDetail].NamaBarang;
                            sheet.Cells[j, i + 3] = listdetaildata[HitungDetail].Quantity;
                            sheet.Cells[j, i + 4] = listdetaildata[HitungDetail].Satuan;

                            //format kasi rapi
                            sheet.Cells[j, i + 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                            j++;
                        }
                        HitungDetail++;
                    }
                    j++;
                    count++;
                }
                //buat lagi format biar rapi
                sheet.Range["A1", "E1"].Font.Bold = true;
                sheet.Range["A1", "E1"].MergeCells = true;
                sheet.Range["A1", "E1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                sheet.Range["A3", "C3"].Font.Bold = true;
                sheet.Range["B4", "E4"].Font.Bold = true;

                sheet.Range["A3", "C3"].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                sheet.Range["A3", "C3"].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 3d;

                sheet.Range["B4", "E4"].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                sheet.Range["B4", "E4"].Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 3d;

                sheet.Columns.AutoFit();
                sheet.Rows.AutoFit();

                sheet.Name = "DataOrder";

                app.UserControl = true;

                //pake proteksi password
                book.Password = "password";
            }

            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }                       
        }

        private void dataGridViewDatabase_RowAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            int record = dataGridViewDatabase.Rows.Count;
            lblcount.Text = record + " Record";
        }
    }
}
