using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfApp1
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void ReadExcel_and_Import_to_DB(object sender, RoutedEventArgs e)
        {
            string connString = "";
            DataTable OTHERS = new DataTable();
            OTHERS = ExceltoDatatableSheet1("prolog.xlsx");

            SqlConnection conn = new SqlConnection(connString);
            conn.Open();
            using (var adapter = new SqlDataAdapter("SELECT * FROM table", conn))
            using (var builder = new SqlCommandBuilder(adapter))
            {
                adapter.InsertCommand = builder.GetInsertCommand();
                adapter.Update(OTHERS);
            }
            conn.Close();
            MessageBox.Show("Imported Successfully");
        }
        public static DataTable ExceltoDatatableSheet1(String Input)
        {
            FileStream fs = File.Open(Input, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            ExcelPackage ep = new ExcelPackage(fs);
            DataTable dt = new DataTable();
            using (ep)
            {
                try
                {
                    var w = ep.Workbook.Worksheets[1]; //instead of 1 you can use sheet name           
                    for (int Col = 1; Col <= w.Dimension.End.Column; Col++)
                    {
                        String cn = (w.Cells[1, Col].Value ?? "").ToString();
                        if (dt.Columns.Contains(cn) || String.IsNullOrEmpty(cn))
                        {
                            cn = String.Format("Column {0}", Col);
                        }
                        dt.Columns.Add(cn);
                    }
                    for (int Row = 2; Row <= w.Dimension.End.Row; Row++)
                    {
                        DataRow dr = dt.NewRow();
                        for (int Col = 1; Col <= w.Dimension.End.Column; Col++)
                        {
                            dr[Col - 1] = (w.Cells[Row, Col].Value ?? "").ToString();
                        }
                        dt.Rows.Add(dr);
                    }

                }
                catch (System.Exception ex)
                {
                    UserInfo.PrintToInfoBox("Sheet1 Not Found");
                    Console.WriteLine(ex);
                }
            }
            return dt;
        }
    }
}
