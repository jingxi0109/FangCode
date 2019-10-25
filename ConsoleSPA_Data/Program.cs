using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Reflection;
using System.IO;
using System.Data;
using ExcelDataReader;
using System.Runtime.InteropServices;
using DataTable = System.Data.DataTable;

namespace ConsoleSPA_Data
{
    class Program
    {
        static void Main(string[] args)
        {
            // SPA_Data_Migration.Open_xls_File();
            DataSet dsListPrice;//=new DataSet();
            DataSet dsSPA;
            // dsSPA=ConsoleSPA_Data.SPA_Data_Migration.ExcelRead_DS(SPA_Data_Migration.Fname);
            dsSPA = ConsoleSPA_Data.SPA_Data_Migration.ExcelRead_DS("SPA特价.xlsx");
            dsSPA.WriteXml("cba.xml", XmlWriteMode.WriteSchema);
            dsListPrice = ConsoleSPA_Data.SPA_Data_Migration.ExcelRead_DS("PriceList.xlsx");//SPA_Data_Migration.Fname);

            dsListPrice.WriteXml("abc.xml", XmlWriteMode.WriteSchema);
            DataSet Totalds = new DataSet();
            Totalds.Merge(dsSPA);
            Totalds.Merge(dsListPrice);
            // ds.ReadXml("abc.xml");
           SPA_Data_Migration.removeNullRow(Totalds, "SYS-2029TP-HTR");
            Totalds.WriteXml("Total_DS.xml", XmlWriteMode.WriteSchema);
            Console.WriteLine(SPA_Data_Migration.OutputString);
            Console.WriteLine("done!!!");
            Console.ReadLine();
        }

    }
    public class SPA_Data_Migration
    {

        public SPA_Data_Migration()
        {

        }
        private static string outputString;

        public static void removeNullRow(DataSet dataSet,string ky )
        {
            string keyw = ky.ToUpper();// "SYS-2029TP-HTR";
            //keyw = keyw.ToUpper();
// outputString = outputString+"\n" + "========================" + "\n";
            foreach (DataTable dt in dataSet.Tables)
            {
                foreach(DataRow dr in dt.Rows)
                {

                  if(  dr[dt.Columns[0]].ToString().Contains(keyw))
                    {
                        outputString = outputString + "\n";
                        foreach (var ress in dt.Columns)
                        {
                            outputString = outputString + ress.ToString() + "\t";
                        }
                        outputString = outputString + "\n";
                        foreach (var t in dr.ItemArray)
                        {

                            outputString = outputString + t.ToString() + "\t";
                        }
                        outputString = outputString + "\n";
                    }
                    if (dr[dt.Columns[1]].ToString().Contains(keyw))
                    {
                        outputString = outputString + "\n";
                        foreach (var ress in dt.Columns)
                        {
                            outputString = outputString + ress.ToString() + "\t";
                        }
                        outputString = outputString + "\n";
                        foreach (var t in dr.ItemArray)
                        {

                            outputString = outputString + t.ToString() + "\t";
                        }
                        outputString = outputString + "\n";
                    }
                }
                //////////Console.WriteLine(dt.TableName);
                ////////var dtt = dt.AsEnumerable();
                ////////var res = dtt.Where(a => a.Field<string>(dt.Columns[0])==(keyw));

                ////////foreach (var re in res)
                ////////{

                ////////    //   Console.WriteLine("========================");
                ////////    outputString = outputString + "\n";
                ////////    foreach (var ress in re.Table.Columns)
                ////////    {
                ////////        outputString = outputString + ress.ToString() + "\t";
                ////////    }
                ////////  //  Console.WriteLine();
                ////////    outputString = outputString + "\n";
                ////////    foreach (var t in re.ItemArray)
                ////////    {

                ////////        outputString = outputString + t.ToString() + "\t";
                ////////    }
                ////////}
                ////////// Console.WriteLine();
                ////////outputString = outputString + "\n";
                //////////foreach (DataColumn dc in dt.Columns)
                //////////{
                //////////    Console.Write(dc.ColumnName + "\t");
                //////////}

            }
            //foreach (DataTable dt in dataSet.Tables)
            //{
            //    //Console.WriteLine(dt.TableName);
            //    var dtt = dt.AsEnumerable();
            //    //var res;
            //    //try
            //    //{
            //    //    res = dtt.Where(a => a.Field<string>(dt.Columns[1]) == keyw);
            //    //}
            //    //catch (Exception ex)
            //    //{
            //    // Console.WriteLine(ex.Message);
            //    try
            //    {
            //       // outputString = outputString + "\n" + "========================" + "\n";
            //        //  dt.Columns[1].DataType = typeof(System.String);
            //        var res = dtt.Where(a => a.Field<string>(dt.Columns[1])==(keyw));
            //        foreach (var re in res)
            //        {
            //            //  Console.WriteLine("========================");
            //            // outputString = outputString  + "========================"+ "\n";
            //            outputString = outputString + "\n";
            //            foreach (var ress in re.Table.Columns)
            //            {
            //               // Console.Write(ress.ToString() + "\t");
            //                outputString = outputString + ress.ToString() + "\t";
            //            }
            //           // Console.WriteLine();
            //            outputString = outputString + "\n";
            //            foreach (var t in re.ItemArray)
            //            {

            //               // Console.Write(t.ToString() + "\t");
            //                outputString = outputString + t.ToString() + "\t";
            //            }
            //        }
            //    }catch(Exception ex)
            //    {
            //      //  Console.WriteLine("======================"+ex.Message);
            //    }
                
            //    //}


                

            //}


            Console.WriteLine();
            //foreach (DataColumn dc in dt.Columns)
            //{
            //    Console.Write(dc.ColumnName + "\t");
            //}

        }




        public static void Open_xls_Files()
        {
            // creating Excel Application  
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            // creating new WorkBook within Excel application  
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Open(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location‬‬) + @"\SPA特价.xlsx");
            //app.Workbooks.Add(Type.Missing);
            // creating new Excelsheet in workbook  
            Microsoft.Office.Interop.Excel._Worksheet worksheet = workbook.Sheets[1];
            // see the excel sheet behind the program  
            app.Visible = false;
            // get the reference of first sheet. By default its name is Sheet1.  
            // store its reference to worksheet  
            worksheet = workbook.Sheets["SPA台湾"];
            worksheet = workbook.ActiveSheet;
            // changing the name of active sheet  
            worksheet.Copy(Type.Missing, worksheet);
            Microsoft.Office.Interop.Excel.Worksheet wcSheet = workbook.Sheets[2];
            wcSheet.Name = "ExSPA台湾";


            wcSheet.UsedRange.UnMerge();
            Console.WriteLine(wcSheet.UsedRange.Rows.Count);
            Range Ro1 = null;
            Range Rd1 = null;
            Range Ro2 = null;
            Range Rd2 = null;
            Range tempR = null;
            for (int i = 1; i <= wcSheet.UsedRange.Rows.Count; i++)
            {

                try
                {
                    string s = wcSheet.UsedRange.Cells[i, 1].Value2.ToString();
                    Console.WriteLine(s);
                    if (wcSheet.Cells[i, 1].Value2 != null)//DBNull)
                    {




                        if (s == "SPA: ")
                        {

                            Ro1 = wcSheet.Range[wcSheet.Cells[i, 1], wcSheet.Cells[i, 9]];
                            // Console.WriteLine("=============="+wcSheet.Cells[i, 1].Value2.ToString());
                            Rd1 = wcSheet.Range[wcSheet.Cells[i + 2, 10], wcSheet.Cells[i + 2, 19]];
                            Ro1.Cut(Rd1);



                            /////////////////////////////////////
                            Ro2 = wcSheet.Range[wcSheet.Cells[i + 1, 1], wcSheet.Cells[i + 1, 9]];
                            // Console.WriteLine(Ro.Cells.Count);
                            Rd2 = wcSheet.Range[wcSheet.Cells[i + 4, 10], wcSheet.Cells[i + 4, 19]];
                            Ro2.Cut(Rd2);
                            // Console.WriteLine(wcSheet.UsedRange.Cells[i,1]);


                            i = i + 4;
                        }
                        //Rd2.Copy(wcSheet.Range[wcSheet.Cells[i + 4, 10], wcSheet.Cells[i + 4, 19]]);
                        tempR = wcSheet.Range[wcSheet.Cells[i, 10], wcSheet.Cells[i, 19]];


                        //= //Rd2.Copy(wcSheet.Range[wcSheet.Cells[i + 3, 10], wcSheet.Cells[i + 3, 19]]);
                        Ro2.Copy(tempR);
                        //   tempR.Clear();
                        //  Ro2.Clear();
                    }
                    else if (wcSheet.Cells[i, 1].Value == "")
                    {





                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);

                }

            }

            //for (int i = 1; i <= wcSheet.UsedRange.Rows.Count; i++)
            //{
            //    try
            //    {
            //        string s = wcSheet.UsedRange.Cells[i, 1].Value2.ToString();
            //        Console.WriteLine(s);
            //        if (wcSheet.Cells[i, 1].Value2 != null)//DBNull)
            //        {




            //            if (s == "SPA: ")
            //            {



            //                /////////////////////////////////////
            //              //  Ro2 = wcSheet.Range[wcSheet.Cells[i + 1, 1], wcSheet.Cells[i + 1, 9]];
            //                // Console.WriteLine(Ro.Cells.Count);
            //              //  Rd2 = wcSheet.Range[wcSheet.Cells[i + 4, 10], wcSheet.Cells[i + 4, 19]];
            //                // Ro2.Cut(Rd2);
            //                // Console.WriteLine(wcSheet.UsedRange.Cells[i,1]);

            //                Ro1 = wcSheet.Range[wcSheet.Cells[i, 1], wcSheet.Cells[i, 9]];
            //                // Console.WriteLine(Ro.Cells.Count);
            //                Rd1 = wcSheet.Range[wcSheet.Cells[i + 2, 10], wcSheet.Cells[i + 2, 19]];
            //                Ro1.Cut (Rd1);
            //            }
            //            //Rd2.Copy(wcSheet.Range[wcSheet.Cells[i + 4, 10], wcSheet.Cells[i + 4, 19]]);
            //         //   tempR = wcSheet.Range[wcSheet.Cells[i, 10], wcSheet.Cells[i, 19]];


            //            //= //Rd2.Copy(wcSheet.Range[wcSheet.Cells[i + 3, 10], wcSheet.Cells[i + 3, 19]]);
            //          //  Ro2.Copy(tempR);
            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        Console.WriteLine(ex.Message);
            //    }

            //}
            //// storing header part in Excel  
            ////for (int i = 1; i < this.DGV1.Columns.Count + 1; i++)
            ////{
            ////    worksheet.Cells[1, i] = DGV1.Columns[i - 1].Header;//.HeaderText;
            //                                                       // worksheet.Columns.ColumnWidth 
            //}
            //// storing Each row and column value to excel sheet  
            //for (int i = 0; i < DGV1.Items.Count - 1; i++)
            //{
            //    for (int j = 0; j < DGV1.Columns.Count; j++)
            //    {
            //        //  var drview= DGV1.Items[i];
            //        //   worksheet.Cells[i + 2, j + 1] = .Cells[j].Value.ToString();
            //        // System.Data.DataRowView dataRow = drview;
            //        DataRowView dataRow = (DataRowView)DGV1.Items[i];//.Row.//.Row;
            //        worksheet.Cells[i + 2, j + 1] = dataRow[j].ToString();//dataRow.Row..ToString();
            //                                                              // worksheet.Cells[i + 2, j + 1]
            //                                                              //MessageBox.Show(DGV1.Items[i].ToString());
            //    }
            //}

            // save the application  


            // Microsoft.Office.Interop.Excel.Worksheet ws = app.ActiveWorkbook.Worksheets[1];
            // Microsoft.Office.Interop.Excel.Range range = ws.UsedRange;
            // ws.Columns.AutoFit();
            // ws.Rows.AutoFit();
            try
            {
                Fname = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location‬‬) + @"\temp" + DateTime.Now.ToString().Trim().Replace(':', '_').Replace('/', '_').Trim() + ".xlsx";

                workbook.SaveAs(Fname, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                // Exit from the application  
                // workbook.SaveCopyAs(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location‬‬) + @"\列宽output.xls");
            }
            catch (Exception es)
            {
                //   MessageBox.Show(es.Message);
                Console.WriteLine(es.Message);
            }
            // workbook.Saved = true;
            app.Quit();
            Marshal.ReleaseComObject(app);

            Console.WriteLine("excel文件导出到当前app运行目录下，命名为output.xls,Excel将关闭退出~");

        }
        static string fname;

        public static string Fname { get => fname; set => fname = value; }
        public static string OutputString { get => outputString; set => outputString = value; }

        public static DataSet ExcelRead_DS(string filePath)
        {
            System.Data.DataSet ds;
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // Choose one of either 1 or 2:

                    // 1. Use the reader methods
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        // Gets or sets a value indicating whether to set the DataColumn.DataType 
                        // property in a second pass.
                        UseColumnDataType = true,

                        // Gets or sets a callback to determine whether to include the current sheet
                        // in the DataSet. Called once per sheet before ConfigureDataTable.
                        FilterSheet = (tableReader, sheetIndex) => true,

                        // Gets or sets a callback to obtain configuration options for a DataTable. 
                        ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                        {
                            // Gets or sets a value indicating the prefix of generated column names.


                            // Gets or sets a value indicating whether to use a row from the 
                            // data as column names.
                            UseHeaderRow = true,




                        }
                    });
                    ds = result; //reader.AsDataSet(exc);


                    // The result of each spreadsheet is in result.Tables
                }
            }
            return ds;
        }
        public static void Change_XLS_Cell(int sheet_Num, int other, Microsoft.Office.Interop.Excel.Workbook wb)
        {
            Microsoft.Office.Interop.Excel._Worksheet ws = wb.Sheets[sheet_Num];
            Microsoft.Office.Interop.Excel.Range xr_o = ws.UsedRange;
            Microsoft.Office.Interop.Excel.Range xr_d = ws.UsedRange;
            xr_o.Cut(xr_d);
        }


    }
}