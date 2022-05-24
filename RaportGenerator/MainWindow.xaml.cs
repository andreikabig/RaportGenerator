using Microsoft.Win32;
using System;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Data;
using System.IO;
using System.Text.Json;
using RaportGenerator.Classes.Settings;
using RaportGenerator.FixedClasses;
using Word = Microsoft.Office.Interop.Word;


namespace RaportGenerator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // List of tables - NOT USED
        private TablesCollection tables1;
        private TablesCollection tables2;

        // Tables
        DataTableCollection? tableCollection1 = null;
        DataTableCollection? tableCollection2 = null;

        // Настройки
        Root? settings;

        // Путь сохранения
        

        public MainWindow()
        {
            InitializeComponent();

            // Переопределение кодировки
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Подгрузка настроек
            LoadSettings();
        }


        // МЕТОД ПОДГРУЗКИ НАСТРОЕК
        private void LoadSettings()
        {
            using (FileStream fs = new FileStream(@"..\..\..\appsettings.json", FileMode.OpenOrCreate))
            {
                settings = JsonSerializer.Deserialize<Root>(fs);
            }
        }


        // МЕТОД КНОПКИ ОТКРЫТИЯ ПЕРВОЙ ТАБЛИЦЫ
        private void BtnLoadFirstTable_Click(object sender, RoutedEventArgs e)
        {
            var res = SaveTableCollection();

            if (tableCollection1 == null || res != null)
                tableCollection1 = res;

            // Установка значений в комбобоксы
            ChangeCombobox(ComboBoxPages, tableCollection1);
        }


        // МЕТОД КНОПКИ ОТКРЫТИЯ ВТОРОЙ ТАБЛИЦЫ
        private void BtnLoadSecondTable_Click(object sender, RoutedEventArgs e)
        {
            var res = SaveTableCollection();

            if (tableCollection2 == null || res != null)
                tableCollection2 = res;

            // Установка значений в комбобоксы
            ChangeCombobox(ComboBoxPages2, tableCollection2);
        }

        // МЕТОД ВЫБОРА ПУТИ СОХРАНЕНИЯ ФАЙЛОВ
        private void BtnSaveFolder_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("На данный момент нет возможности выбора пути сохранения. Ваш отчет будет сохранен на диске D. В противном случае его придется сохранять самостоятельно.");

        }

        // МЕТОД СОХРАНЕНИЯ СЧИТАННОЙ ИНФОРМАЦИИ С EXCEL
        private DataTableCollection? SaveTableCollection()
        {
            var result = OpenFile();

            if (result == null)
            {
                MessageBox.Show("Таблица не была загружена!");
            }

            // Новое значение для коллекции таблиц
            return result;

        }


        // МЕТОД ОТКРЫТИЯ ДИАЛОГОВОГО ОКНА
        private DataTableCollection? OpenFile()
        {
            // Возвращаемые данные
            DataTableCollection? tc = null;
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
                tc = ExcelTableReader.Load(fileSelect.FileName);
            }

            // Возвращаем таблицу
            return tc;
        }


        // МЕТОД ЗАПОЛНЕНИЯ КОМБОБОКСОВ НА ОСНОВЕ СЧИТАННЫХ ДАННЫХ
        private void ChangeCombobox(ComboBox comboBox, DataTableCollection? tc)
        {
            // Подсистка прошлых значений комбобокса
            comboBox.Items.Clear();

            if (tc != null)
            {
                // Отображаем каждую страницу Excel в ComboBoxPages
                foreach (DataTable dt in tc)
                {
                    // Добавляем в ComboBoxPages название страницы
                    comboBox.Items.Add(dt.TableName);
                }

                // Устанавливаем выбранной страницей - первую страницу
                comboBox.SelectedIndex = 0;
            }
        }


        // МЕТОД ОТКРЫТИЯ ИНФОРМАЦИОННОГО ОКНА
        private void BtnInfo_Click(object sender, RoutedEventArgs e)
        {
            HelpWindow helpWindow = new HelpWindow();
            helpWindow.Show();
        }


        // МЕТОД ОТКРЫТИЯ ОКНА НАСТРОЕК
        private void BtnSettings_Click(object sender, RoutedEventArgs e)
        {
            SettingsWindow settingsWindow = new SettingsWindow();
            settingsWindow.Show();
        }


        // МЕТОД ОБНОВЛЕНИЯ СТРАНИЦ ПРИ СМЕНЕ ЗНАЧЕНИЯ КОМБОБОКСА
        private void ComboBoxPages_SelectionChanged(object sender, SelectionChangedEventArgs e) => ShowSelectedTable(tableCollection1, ComboBoxPages);
        private void ComboBoxPages2_SelectionChanged(object sender, SelectionChangedEventArgs e) => ShowSelectedTable(tableCollection2, ComboBoxPages2);

        
        // МЕТОД ЗАПОЛНЕНИЯ DGV
        private void ShowSelectedTable(DataTableCollection dtc, ComboBox comboBox)
        {
            DataTable? table = dtc[Convert.ToString(comboBox.SelectedItem)];

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


        // КНОПКА СОЗДАНИЯ ОТЧЕТА
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (tableCollection1 != null && tableCollection2 != null)
            {
                tables1 = new TablesCollection(tableCollection1, settings.Table1);
                tables2 = new TablesCollection(tableCollection2, settings.Table2);
            }
            else
            {
                MessageBox.Show("Одна или несколько таблиц не были загружены.", "Невозможно сгенерировать отчет", MessageBoxButton.OK);
            }

            // Номер таблицы

            int count = 1;
            // РАБОТА С WORD

            var application = new Word.Application();

            // Добавляем в приложение новый документ
            Word.Document document = application.Documents.Add();
            
            // Range - доступ к тексту

            foreach (var table in tables1.Tables)
            {
                // Название таблицы
                var nameParagraph = document.Paragraphs.Add();
                var nameRange = nameParagraph.Range;
                nameRange.Text = table.Name;
                nameRange.InsertParagraphAfter();
                nameRange.Font.Name = "Times New Roman";

                // Нумерация таблицы
                var tableCountParagraph = document.Paragraphs.Add();
                var tableCountRange = tableCountParagraph.Range;
                tableCountRange.Text = $"Таблица {count}";
                tableCountRange.InsertParagraphAfter();
                tableCountRange.Font.Name = "Times New Roman";

                // Таблица
                var tableParagraph = document.Paragraphs.Add();
                var tableRange = tableParagraph.Range;
                Word.Table t = document.Tables.Add(tableRange, table.Entries.Count + 3, 5); // range, строк, столбцов
                t.Rows.SetHeight(2.5f, Word.WdRowHeightRule.wdRowHeightAuto); // Установка высоты столбцов - допустимо
                

                // Границы
                t.Borders.InsideLineStyle = t.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

                // Выравнивание по центру по вертикали
                t.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                t.Rows[1].Range.Bold = 1;
                t.Rows[1].Range.Font.Name = "Times New Roman";
                t.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                // Выравнивание по центру
                for (int i = 1; i <= t.Rows.Count; i++)
                {
                    t.Rows[i].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }


                // Выравнивание справа
                for (int i = 3; i <= t.Rows.Count; i++)
                {
                    t.Rows[i].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                }

               

                // Информация об определенной ячейки в таблице
                Word.Range cellRange;
                cellRange = t.Cell(1, 1).Range;
                cellRange.Font.Name = "Times New Roman";
                cellRange.Text = "Наименование муниципального района/городского округа";

                t.Cell(1, 1).Merge(t.Cell(2, 1));

                cellRange = t.Cell(1, 2).Range;
                cellRange.Text = table.Name;
                cellRange.Font.Name = "Times New Roman";


                // Объединение ячеек Названия таблицы
                t.Cell(1, 2).Merge(t.Cell(1,3));
                t.Cell(1,2).Merge(t.Cell(1,3));
                t.Cell(1,2).Merge(t.Cell(1,3));
                //t.Cell(1, 4).Merge(t.Cell(1, 5));

                // Заполняем названия колонок свойств 2.2 2.3 2.4 2.5
                for (int i = 0; i < 4; i++)
                {
                    cellRange = t.Cell(2, i + 2).Range;
                    
                    cellRange.Text = table.Entries[0].Properties[i].Name;
                }

                // Заполнение Entries
                for (int i = 0; i < table.Entries.Count; i++)
                { 
                    // Вытаскиваем название Entrie
                    var currentEntrie = table.Entries[i];
                    cellRange = t.Cell(i+3, 1).Range;
                    

                    cellRange.Text = currentEntrie.Name;

                    // Заполнение свойств
                    for (int j = 0; j < table.Entries[i].Properties.Count; j++)
                    { 
                        var currentProperty = table.Entries[i].Properties[j];
                        cellRange = t.Cell(i+3, j+2).Range;
                        
                        //cellRange.Text = Convert.ToString(currentProperty.Value);
                        
                        if (j < table.Entries[i].Properties.Count - 1)
                            cellRange.Text = Convert.ToString(currentProperty.Value);
                        else
                            cellRange.Text = $"{currentProperty.Value:N1}";
                        
                    }
                }

                // Добавление итога
                cellRange = t.Cell(table.Entries.Count + 3, 1).Range;
                cellRange.Text = "Итог";

                // Подсчеты по первому свойству
                // Подсчеты по второму свойству
                int firstValue = 0;
                int secondValue = 0;
                foreach (var entry in table.Entries)
                {
                    try
                    {
                        firstValue += (int)entry.Properties[0].Value;
                    }
                    catch { }
                    try
                    {
                        secondValue += (int)entry.Properties[1].Value;
                    }
                    catch { }
                }

                // Подсчеты по третьему свойству
                int thirdValue = firstValue - secondValue;

                // Подсчеты по четвертому свойству
                double fourthValue = (Convert.ToDouble(firstValue) / Convert.ToDouble(secondValue)) * 100 - 100;


                cellRange = t.Cell(table.Entries.Count + 3, 2).Range;
                cellRange.Text = $"{firstValue}";

                cellRange = t.Cell(table.Entries.Count + 3, 3).Range;
                cellRange.Text = $"{secondValue}";

                cellRange = t.Cell(table.Entries.Count + 3, 4).Range;
                cellRange.Text = $"{thirdValue}";

                cellRange = t.Cell(table.Entries.Count + 3, 5).Range;
                cellRange.Text = $"{fourthValue:N1}";



                document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);


                count++;
            }

            // Создание второй группы таблиц
            foreach (var table in tables2.Tables)
            {
                // Название таблицы
                var nameParagraph = document.Paragraphs.Add();
                var nameRange = nameParagraph.Range;
                nameRange.Text = table.Name;
                nameRange.InsertParagraphAfter();
                nameRange.Font.Name = "Times New Roman";

                // Нумерация таблицы
                var tableCountParagraph = document.Paragraphs.Add();
                var tableCountRange = tableCountParagraph.Range;
                tableCountRange.Text = $"Таблица {count}";
                tableCountRange.InsertParagraphAfter();
                tableCountRange.Font.Name = "Times New Roman";

                // Таблица
                var tableParagraph = document.Paragraphs.Add();
                var tableRange = tableParagraph.Range;
                Word.Table t = document.Tables.Add(tableRange, table.Entries.Count + 2, 3); // range, строк, столбцов
                t.Rows.SetHeight(2.5f, Word.WdRowHeightRule.wdRowHeightAuto); // Установка высоты столбцов - допустимо


                // Границы
                t.Borders.InsideLineStyle = t.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

                // Выравнивание по центру по вертикали заголовков
                t.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                t.Rows[1].Range.Bold = 1;
                t.Rows[1].Range.Font.Name = "Times New Roman";

                for (int a = 1; a <= t.Rows.Count; a++)
                {
                    t.Rows[a].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                }
                
                

                // Формирование заголовков
                Word.Range cellRange;
                cellRange = t.Cell(1, 1).Range;
                cellRange.Font.Name = "Times New Roman";
                cellRange.Text = "Территория";

                cellRange = t.Cell(1, 2).Range;
                cellRange.Font.Name = "Times New Roman";
                cellRange.Text = table.Entries[0].Properties[0].Name;

                cellRange = t.Cell(1, 3).Range;
                cellRange.Font.Name = "Times New Roman";
                cellRange.Text = table.Entries[0].Properties[1].Name;

                // Заполнение Entries
                // Заполнение Entries
                for (int i = 0; i < table.Entries.Count; i++)
                {
                    // Вытаскиваем название Entrie
                    var currentEntrie = table.Entries[i];
                    cellRange = t.Cell(i + 2, 1).Range;


                    cellRange.Text = currentEntrie.Name;

                    // Заполнение свойств
                    for (int j = 0; j < table.Entries[i].Properties.Count; j++)
                    {
                        var currentProperty = table.Entries[i].Properties[j];
                        cellRange = t.Cell(i + 2, j + 2).Range;

                        //cellRange.Text = Convert.ToString(currentProperty.Value);
                        try
                        {
                            cellRange.Text = Convert.ToString((int)currentProperty.Value);
                        }
                        catch { }
                        
                    }
                }


                // Добавление итога
                cellRange = t.Cell(table.Entries.Count + 2, 1).Range;
                cellRange.Text = "Всего по области";

                int firstValue = 0;
                int secondValue = 0;

                foreach (var entry in table.Entries)
                {
                    try
                    {
                        firstValue += (int)entry.Properties[0].Value;
                    }
                    catch { }
                    try
                    {
                        secondValue += (int)entry.Properties[1].Value;
                    }
                    catch { }
                }

                cellRange = t.Cell(table.Entries.Count + 2, 2).Range;
                cellRange.Text = $"{firstValue}";

                cellRange = t.Cell(table.Entries.Count + 2, 3).Range;
                cellRange.Text = $"{secondValue}";

                //document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
            }



            application.Visible = true;
            try
            {

            }
            catch 
            {
                MessageBox.Show(@"Не удалось сохранить отчет по пути D:\\, пожалуйста сохраните отчет самостоятельно.");
            }
            document.SaveAs2(@"D:\Отчет аналитики.docx");
            document.SaveAs2(@"D:\Отчет аналитики.pdf", Word.WdExportFormat.wdExportFormatPDF);

        }


    }
}