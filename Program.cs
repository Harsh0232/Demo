using ExcelDataReader;
using NPOI.SS.UserModel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ConsoleApp2
{
    
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
             
            int value;
            Console.WriteLine("1: Export File To selected Excel" +"\n"+
                              "2: Import File To SQl DataBase");

            string input = Console.ReadLine();

            value = Convert.ToInt32(input);
            string connstring = @"Data Source = LOCALHOST\SQLEXPRESS;Initial Catalog= Infrasol;Trusted_Connection=true;";

            switch (value) {
                
                //In case 1 it will send the data to Excel file from the sql server, User can select the excel file 
                //through prompt box.
                case 1:
                    {
                        //Dialog Box to select the file to export

                        //Created filter so only Excel file(.xlsx; .xls) will be displayed to customer to
                        //select other will be filtered out.
                        OpenFileDialog choofdlog = new OpenFileDialog();
                        choofdlog.Filter = "Excel Document|*.xlsx;*.xls";
                        choofdlog.FilterIndex = 1;
                        if (choofdlog.ShowDialog() == DialogResult.OK)
                        {
                            var sFileName = choofdlog.FileName;
                            
                            //Initilizing file path to fileInfo 
                            var file1 = new FileInfo(sFileName);
                            try
                            {
                                //Creating Excel package 
                                using (ExcelPackage excel = new ExcelPackage(file1))
                                {
                                    //To select excel sheet for exporting the data
                                    ExcelWorksheet sheet = excel.Workbook.Worksheets["sheet1"];
                                   
                                    SqlConnection conn = new SqlConnection(connstring);
                                    conn.Open();
                                    var command = new SqlCommand("select TOP 6 * from dbo.Data", conn);
                                   
                                    SqlDataAdapter da = new SqlDataAdapter(command);
                                    DataTable dataTable = new DataTable();
                                   
                                    da.Fill(dataTable);
                                    int count = dataTable.Rows.Count;
                                    
                                    //Fill Data to Excel sheet
                                    sheet.Cells.LoadFromDataTable(dataTable, true);
                                    FileInfo exceFile = new FileInfo(sFileName);
                                    excel.SaveAs(exceFile);

                                }
                            }
                            catch (Exception exception)
                            {
                                Console.WriteLine(exception.Message);
                            }
                        }
                        break;
                    }
                
                //In case 2 it send the data from Excel to Sql server 
                case 2:
                    {
                        try
                        {
                            //File from where data to import to SQL server
                            string excelFilePath = @"E:\Asset.xlsx";
                          
                            //Connection String Forrmate for Excel File
                            String excelConnString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0\"", excelFilePath);
                            //Create Connection to Excel work book 
                            using (OleDbConnection excelConnection = new OleDbConnection(excelConnString))
                            {
                                //Create OleDbCommand to fetch data from Excel 
                                using (OleDbCommand cmd = new OleDbCommand("Select * from [Sheet1$]", excelConnection))
                                {
                                    excelConnection.Open();
                                    using (OleDbDataReader dataReader = cmd.ExecuteReader())
                                    {
                                        using (SqlBulkCopy sqlBulk = new SqlBulkCopy(connstring))
                                        {
                                            // Destination table  
                                            sqlBulk.DestinationTableName = "dbo.Data";
                                            sqlBulk.WriteToServer(dataReader);
                                            Console.WriteLine(sqlBulk);
                                        }
                                    }
                                }
                            }
                        }
                        catch(Exception exception)
                        {
                            Console.WriteLine(exception.Message);
                        }
                        break;
                    }

            }
           
        }
    }
}
