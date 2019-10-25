using ConsoleSPA_Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CCView
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

    


        private void Button_Click(object sender, RoutedEventArgs e)
        {

            SPA_Data_Migration.OutputString = "";
            this.textBox1.Text = "";
            // ds.ReadXml("abc.xml");
            SPA_Data_Migration.removeNullRow(Totalds, this.textBox.Text);
           // Totalds.WriteXml("Total_DS.xml", XmlWriteMode.WriteSchema);
            this.textBox1.Text = SPA_Data_Migration.OutputString;
          //  Console.WriteLine();
        }
        DataSet Totalds = new DataSet();
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            
              //dsSPA=ConsoleSPA_Data.SPA_Data_Migration.ExcelRead_DS(SPA_Data_Migration.Fname);


            DataSet dsListPrice;//=new DataSet();
            DataSet dsSPA;
            
            dsSPA = ConsoleSPA_Data.SPA_Data_Migration.ExcelRead_DS("SPA特价.xlsx");
            dsSPA.WriteXml("cba.xml", XmlWriteMode.WriteSchema);
            dsListPrice = ConsoleSPA_Data.SPA_Data_Migration.ExcelRead_DS("PriceList.xlsx");//SPA_Data_Migration.Fname);

            dsListPrice.WriteXml("abc.xml", XmlWriteMode.WriteSchema);
            
            Totalds.Merge(dsSPA);
            Totalds.Merge(dsListPrice);

            DataTable itemTable=Totalds.Tables["SPA台湾"];
            Model.Data.Clear();
            foreach (DataRow dataRow in itemTable.Rows)
            {
                
                if(dataRow[1]!=null)
                {
                   // addItem(dataRow[1].ToString());
                    Model.Data.Add(dataRow[1].ToString());
                }
            }
        }

        private void Button1_Click(object sender, RoutedEventArgs e)
        {
            SPA_Data_Migration.Open_xls_Files();
            MessageBox.Show("更新完成");
        }
    }
}
