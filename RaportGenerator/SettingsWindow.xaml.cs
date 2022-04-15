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
using System.Windows.Shapes;
using System.Text.Json;
using System.IO;
using RaportGenerator.Classes.Settings;

namespace RaportGenerator
{
    /// <summary>
    /// Логика взаимодействия для SettingsWindow.xaml
    /// </summary>
    public partial class SettingsWindow : Window
    {
        private Root? settings;

        public SettingsWindow()
        {
            InitializeComponent();
            ReadSettings();
            ShowSettings();
        }

        public void ReadSettings()
        {
            using (FileStream fs = new FileStream(@"..\..\..\appsettings.json", FileMode.OpenOrCreate))
            {
                settings = JsonSerializer.Deserialize<Root>(fs);
            }
        }

        public async void WriteSettings()
        { 
        
        }

        public void ShowSettings()
        {
            if (settings != null)
            {
                // Общая аналитика - заполнение
                AllRawTableName.Text = settings.Table1.TableName[0].ToString();
                AllColTableName.Text = settings.Table1.TableName[1].ToString();

                AllRawDataName.Text = settings.Table1.DataName[0].ToString();
                AllColDataName.Text = settings.Table1.DataName[1].ToString();

                AllRawPropDesc.Text = settings.Table1.Properties.PropertisDescription[0].ToString();
                AllColPropDesc.Text = settings.Table1.Properties.PropertisDescription[1].ToString();

                AllRangeDataX.Text = settings.Table1.Properties.RangeData[0].ToString();
                AllRangeDataY.Text = settings.Table1.Properties.RangeData[1].ToString();

                AllRangePropNameX.Text = settings.Table1.Properties.RangePropertisNames.Range[0].ToString();
                AllRangePropNameY.Text = settings.Table1.Properties.RangePropertisNames.Range[1].ToString();

                AllPropRaw.Text = settings.Table1.Properties.RangePropertisNames.Row.ToString();

                // Детская аналитика - заполнение
                ChildRawTableName.Text = settings.Table2.TableName[0].ToString();
                ChildColTableName.Text = settings.Table2.TableName[1].ToString();

                ChildRawDataName.Text = settings.Table2.DataName[0].ToString();
                ChildColDataName.Text = settings.Table2.DataName[1].ToString();
                
                // Разобраться с этим свойством, т.к. оно в данном случае null

                //try
                //{
                //    ChildRawPropDesc.Text = settings.Table2.Properties.PropertisDescription[0].ToString();
                //    ChildColPropDesc.Text = settings.Table2.Properties.PropertisDescription[1].ToString();
                //}
                //catch (Exception ex)
                //{
                //}

                ChildRangeDataX.Text = settings.Table2.Properties.RangeData[0].ToString();
                ChildRangeDataY.Text = settings.Table2.Properties.RangeData[1].ToString();

                ChildRangePropNameX.Text = settings.Table2.Properties.RangePropertisNames.Range[0].ToString();
                ChildRangePropNameY.Text = settings.Table2.Properties.RangePropertisNames.Range[1].ToString();

                ChildPropRaw.Text = settings.Table2.Properties.RangePropertisNames.Row.ToString();
            }
        }
    }
}
