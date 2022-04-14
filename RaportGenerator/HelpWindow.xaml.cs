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

namespace RaportGenerator
{
    /// <summary>
    /// Логика взаимодействия для HelpWindow.xaml
    /// </summary>
    public partial class HelpWindow : Window
    {
        public HelpWindow()
        {
            InitializeComponent();
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch (InfoComboBox.SelectedIndex)
            {
                case 0:
                    tbInfo.Text = "Программа еще не создана, поэтому руководства по пользованию нет.";
                    break;
                case 1:
                    tbInfo.Text = "Разработчик: Дюдькин Андрей Александрович\n\nВы можете задать свой вопрос:\ntelegram: @likedzizu\nemail: apps.bigint@gmail.com.\n\n\n\nПРИМЕЧАНИЕ: РАЗРАБОТЧИК НЕ НЕСЕТ ОТВЕТСТВЕННОСТИ ЗА СГЕНЕРИРОВАННЫЕ ОТЧЕТЫ. УБЕДИТЕЛЬНАЯ ПРОСЬБА ПЕРЕД ОТПРАВКОЙ ОТЧЕТА ПРОВЕРИТЬ ЕГО НА СОДЕРЖАНИЕ.";
                    break;
                case 2:
                    tbInfo.Text = "Excel документ должен соответствовать следующим требованиям:\n\n1. Документ должен быть с расширением .xlsx\n2. В excel-документе не должно быть пустых листов\n3. Документ должен соответствовать по содержанию и размещению документу, присланному статистикой";
                    break;
                default:
                    break;
            }
        }
    }
}
