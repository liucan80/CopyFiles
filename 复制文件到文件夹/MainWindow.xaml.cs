using System;
using System.Collections.Generic;
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
using System.Windows.Forms;
using Office;
using Microsoft.Office.Interop.Excel;
using System.IO;






namespace 复制文件到文件夹
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        Workbook workbook1;
        Worksheet worksheet1;
        Microsoft.Office.Interop.Excel.Application excelApp;
        List<string> names = new List<string>();
        List<string> AllFileNames = new List<string>();
        int i = 0;

       int n=1;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog filedia1 = new OpenFileDialog();
            filedia1.Filter ="Excel文件 | *.xls;*.xlsx";
            filedia1.ShowDialog();
            filepath.Text = filedia1.FileName;

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;
            workbook1 = excelApp.Workbooks.Open(filepath.Text);
            worksheet1 = workbook1.ActiveSheet;
            while (!Equals(((Range)worksheet1.Cells[n, 1]).Text, ""))
            {
                string a = ((Range)worksheet1.Cells[n, 1]).Text;
                names.Add(a);
                n++;
            }
            TBLog.Clear();
            foreach (var name in names)
            {
                var temp = false;
              
                foreach (var file in AllFileNames)
                {
                    if (file.Contains(name))
                    {
                        //FileInfo fi1 = new FileInfo(file);
                       // fi1.CopyTo(TBTargetFolderPath+@"\" + System.IO.Path.GetFileName(file));
                        var a = System.IO.Path.GetFileName(file);
                        var b = TBTargetFolderPath.Text +@"\"+ a;
                        System.IO.File.Copy(file, b , true);
                        temp = true;
                        

                    }
                   
                }
                if (!temp)
                {
                    
                    i++;
                    TBLog.AppendText(i+":"+name + "未找到\n");
                     //System.Windows.MessageBox.Show(i + ":" + name + "未找到\n");
                }
               

            }
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog br1 = new FolderBrowserDialog();
            br1.ShowNewFolderButton = false;
            br1.ShowDialog();
            TBSourceFolderPath.Text = br1.SelectedPath;
            
            foreach (string file in Directory.GetFiles(br1.SelectedPath))
            {
                AllFileNames.Add(file);
            }
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog br2 = new FolderBrowserDialog();
            br2.ShowNewFolderButton = true;
            br2.ShowDialog();
            TBTargetFolderPath.Text = br2.SelectedPath;
            
            //br1.
        }
    }
}
