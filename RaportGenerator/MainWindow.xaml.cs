
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
using System.Text.Json;
using RaportGenerator.Classes.Settings;

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

            GetData();

            SortTables(tables);


            string nameColumns = "ГОРОД\tСВ1\tСВ2\tСВ3\tСВ4\n";
            string msg = $"{nameColumns}";
            foreach (var entry in tables[0].Entries)
            {
                //msg += $"{entry.Name}\t{entry.QuantityCurrent}\t{entry.QuantityCurrent}\t{entry.DynamicAbs}\t\t{entry.DynamicPersents}\n";

                string properties = $"{entry.Name}";
                
                for (int i = 0; i < entry.Properties.Count(); i++)
                {
                    try
                    {
                        Property<double> property = (Property<double>)entry.Properties[i];
                        properties += $"\t{property.Value}";
                    }
                    catch (Exception ex)
                    {
                        properties += $"\tNULL";
                    }
                }
                msg += $"{properties}\n";


                
            }
            MessageBox.Show(msg);
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
            if (tables != null)
                tables.Clear();
            if (tableCollection != null)
                tableCollection.Clear();
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
            DataTable? table = tableCollection[Convert.ToString(ComboBoxPages.SelectedItem)];

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

        // МЕТОД СБОРА ДАННЫХ ОБЩЕЙ АНАЛИТИКИ СМРЕТНОСТИ
        private void GetData()
        {
            using (FileStream fs = new FileStream(@"..\..\..\appsettings.json", FileMode.OpenOrCreate))
            {
                var settings = JsonSerializer.Deserialize<Root>(fs);

                //DataTable? table = tableCollection[Convert.ToString(ComboBoxPages.SelectedItem)];
                foreach (DataTable table in tableCollection)
                {
                    if (table != null)
                    {
                        // Создание объекта таблицы
                        ExcelTable exTable = new ExcelTable();

                        // Добавление названия таблицы
                        exTable.TableName = table.Rows[settings.Table1.TableName[0]][settings.Table1.TableName[1]].ToString();

                        // Формирование списка объектов таблицы
                        exTable.Entries = new List<Entry>();

                        
                        // Заополнение списка объектов таблицы
                        for (int i = settings.Table1.Properties.RangeData[0]; i <= settings.Table1.Properties.RangeData[1]; i++) // РЕЗУЛЬТИРУЮЩУЮ БУДЕМ ВЫВОДИТЬ САМОСТОЯТЕЛЬНО, ПОСЛЕ СОРТИРОВКИ!
                        {
                            // Создание нового объекта
                            Entry entry = new Entry();
                           
                            // Добавление названия объекта
                            entry.Name = table.Rows[i][settings.Table1.DataName[1]].ToString();

                            /* СВОЙСТВА НЕ ЗАПОЛНЯЮТСЯ */
                            // Заполнение свойств объекта

                            // int _j = settings.Table1.Properties.RangePropertisNames.Range[0]; ПОЧЕМУ ПЕРЕДАЕТСЯ ПО ССЫЛКЕ ? 
                            for (int j = 1; j <= 4; j++)
                            {
                                

                                // Считывание названия свойства
                                string? propName = table.Rows[settings.Table1.Properties.RangePropertisNames.Row][j].ToString();

                                // Считывание значения свойства
                                var value = DoubleConverter(table.Rows[i][j]);

                                // Объявление свойства
                                IProperty property;

                                // Если конвертер вернул double число
                                if (value != null)
                                {
                                    // Создание свойства объекту с числом
                                    property = new Property<double>() { Name = propName, Value = (double)value };
                                }
                                else 
                                {
                                    // Создание свойства объекту без числа
                                    property = new Property<double>() { Name = propName };
                                }

                                // Если свойство есть
                                if (property != null)
                                {
                                    // Добавляем свойство в список свойств
                                    entry.Properties.Add(property);
                                }
                                else
                                    MessageBox.Show("Одно из свойств объекта не было сохранено корректно! Пожалуйста, проверьте правильность составленного документа!");
                                
                                
                            }

                            // ИСПРАВИТЬ НА ОТДЕЛЬНЫЙ МЕТОД
                            //try
                            //{
                            //    entry.QuantityCurrent = (double)table.Rows[i][1];
                            //}
                            //catch (System.InvalidCastException ex)
                            //{
                            //}

                            //try
                            //{
                            //    entry.QuantityLast = (double)table.Rows[i][2];
                            //}
                            //catch (System.InvalidCastException ex)
                            //{
                            //}

                            //try
                            //{
                            //    entry.DynamicAbs = (double)table.Rows[i][3];
                            //}
                            //catch (System.InvalidCastException ex)
                            //{
                            //}
                            //try
                            //{
                            //    entry.DynamicPersents = (double)table.Rows[i][4];
                            //}
                            //catch (System.InvalidCastException ex)
                            //{
                            //}


                            exTable.Entries.Add(entry);
                        }
                        tables.Add(exTable);
                    }
                }
            }

            
        }

        // МЕТОД ПРЕОБРАЗОВАНИЯ ДАННЫХ В DOUBLE
        private double? DoubleConverter(object obj)
        {
            try
            {
                return (double)obj;
            }
            catch (System.InvalidCastException ex)
            {
                return null;
            }
        }

        // МЕТОД  СОРТИРОВКИ ИСХОДНЫХ ДАННЫХ
        private void SortTables(List<ExcelTable> tables)
        {
            //foreach (var table in tables)
            //    table.Entries = table.Entries.OrderBy(e => e.DynamicPersents).ToList();

            // ЖЕСТКО ЗАКОДЕННАЯ СОРТИРОВКА ПО ПОСЛЕДНЕМУ СТОЛБЦУ
            foreach (var table in tables)
                table.Entries = table.Entries.OrderBy(e => e.Properties[^1].Value).ToList(); // НУЖНО УКАЗАТЬ СРАВНЕНИЕ ПО VALUE 
            //tables[0].Entries = tables[0].Entries.OrderBy(e => e.Properties[^1].Value).ToList();
        }

        private void BtnInfo_Click(object sender, RoutedEventArgs e)
        {
            HelpWindow helpWindow = new HelpWindow();
            helpWindow.Show();
        }

       

        private void BtnSettings_Click(object sender, RoutedEventArgs e)
        {
            SettingsWindow settingsWindow = new SettingsWindow();
            settingsWindow.Show();
        }
    }
}
