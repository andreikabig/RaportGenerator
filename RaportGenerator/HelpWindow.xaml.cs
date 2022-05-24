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
                    tbInfo.Text = "Правила работы с генератором отчетов: \n\n1. Выберите первую таблицу (общую аналитику)\n2. Выберите вторую таблицу (младенческая аналитика)\n3. Нажмите кнопку создать отчет\n4. Дождитесь выполнения работы программы\n5. В случае неудачных попыток, поробуйте перезагрузить программу и попробовать сгенерировать отчет еще раз\n6. Воспользуйтесь сгенерированным отчетом для создания оригинального отчета (копируйте информацию с него и вставляйте в исходный отчет)\n\nP.S. Программа находится на ранней стадии разработке, в случае сбоев просьба не пугаться, а попробовать перезагрузить программу и попробовать снова!";
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
