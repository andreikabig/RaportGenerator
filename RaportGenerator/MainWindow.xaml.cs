
using Microsoft.Win32;
using RaportGenerator.Classes;
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
using ExcelDataReader;
using System.Data;
using System.IO;

namespace RaportGenerator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // Path to ExcelFile
        private string pathExcel;

        // List of tables - NOT USED
        private List<ExcelTable> tables;

        // Tables
        DataTableCollection tableCollection = null;

        public MainWindow()
        {
            InitializeComponent();
            // Config the encoding to Russian
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            tables = new List<ExcelTable>();
            
            


        }

        // МЕТОД ЗАГРУЗКИ ДАННЫХ ИЗ EXCEL ФАЙЛА
        public void BtnLoadData_Click(object sender, RoutedEventArgs e)
        {
            // Создаем объект диалогового окна
            OpenFileDialog fileSelect = new OpenFileDialog();

            // Устаналиваем отображаемые файлы на .xlsx в диалоговом окне
            fileSelect.DefaultExt = ".xlsx";

            // Устанавливаем фильтры для диалогового окна
            fileSelect.Filter = "Excel документы (.xlsx)|*.xlsx";

            // Результат взаимодействия с FileDialog
            var result = fileSelect.ShowDialog();

            // Сохраняем путь к файлу, если файл был выбран
            if (result == true)
            {
                pathExcel = fileSelect.FileName;
                //MessageBox.Show(pathExcel);
                OpenExcelFile(pathExcel);
            }

            
        }

        // МЕТОД СЧИТЫВАНИЯ ДАННЫХ С ТАБЛИЦ
        private void OpenExcelFile(string path)
        {
            using (FileStream fs = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(fs))
                {
                    DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration() {
                        ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                        {
                            //UseHeaderRow = true
                        }
                    });
                    tableCollection = db.Tables;


                    ComboBoxPages.Items.Clear();
                    // Отображаем каждую страницу Excel в ComboBoxPages
                    foreach (DataTable dt in tableCollection)
                    {
                        // Добавляем в ComboBoxPages название страницы
                        ComboBoxPages.Items.Add(dt.TableName);
                    }

                    // Устанавливаем выбранной страницей - первую страницу
                    ComboBoxPages.SelectedIndex = 0;
                }
            }
        }

        // МЕТОД ОБНОВЛЕНИЯ DATAGRID ПРИ СМЕНЕ ВЫБРАННОЙ СТРАНИЦЫ
        private void ComboBoxPages_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Выбираем страницу, соответствующую выбранной в ComboBoxPages
            DataTable? table = tableCollection[Convert.ToString(ComboBoxPages.SelectedItem)] ;

            // Если такая страница есть
            if (table != null)
            { 
                // Отображаем ее в DataGrid
                DgvTable.ItemsSource = table.DefaultView;

                // Позволяем автоматически генерировать колонки
                DgvTable.AutoGenerateColumns = true;

                // Запрещаем добавлять новые колонки
                DgvTable.CanUserAddRows = false;
                
            }
        }
    }
}
