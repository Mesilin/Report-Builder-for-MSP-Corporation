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
using Word = Microsoft.Office.Interop.Word;
using Npgsql;
using System.Data.Common;

namespace TestWordApp
{
    public partial class MainWindow : Window
    {
        private Object _missingObj = System.Reflection.Missing.Value;
        private Object _trueObj = true;
        private Object _falseObj = false;
        public int count = 0;
        public MainWindow()
        {
            InitializeComponent();
        }
        //===============================================================================================
        string query_db(string x) 
        {
            string w = null;
            String connectionString = "Server=10.200.50.13;Port=5432;User ID=reader;Password=ti9gMUKFzQ#iXH{aYMB<;Database=cpgu;";
            NpgsqlConnection npgSqlConnection = new NpgsqlConnection(connectionString);
            npgSqlConnection.Open();
            NpgsqlCommand npgSqlCommand = new NpgsqlCommand(x, npgSqlConnection);
            NpgsqlDataReader npgSqlDataReader = npgSqlCommand.ExecuteReader();
            while (npgSqlDataReader.Read())
            w = npgSqlDataReader[0].ToString();
            npgSqlDataReader.Close();
            npgSqlConnection.Close();
            return w;
        }
        public string[,] getRows(int Rows, string x) //
        {
            String connectionString = "Server=10.200.50.13;Port=5432;User ID=reader;Password=ti9gMUKFzQ#iXH{aYMB<;Database=cpgu;";
            NpgsqlConnection connection = new NpgsqlConnection(connectionString);
            using (connection)
            {
                NpgsqlCommand command = new NpgsqlCommand(x, connection);
                connection.Open();
                NpgsqlDataReader reader = command.ExecuteReader();
                string[,] mass=new string[Rows*2,3];
                if (reader.HasRows)
                {
                    int i = 0;
                    while (reader.Read())
                        if ((reader[0].ToString().Length) != 0) //Если ИНН не равен нулю (этим мы отсекаем физлиц)
                        {
                            mass[i, 0] = reader[0].ToString();// ИНН
                            mass[i, 1] = reader[1].ToString();//код услуги
                            mass[i, 2] = reader[2].ToString();//код услуги
                            i++;
                        }
                    count = i;
                }
                reader.Close();
                connection.Close();
                return mass;
            }
        }
        //========================================================================================
        int numberofservice(string service_eid)
        {
            switch (service_eid)
            {
                case "custom2395626": 
                        return 6;
                case "msp100005699":
                        return 6;
                case "custom2395697":
                        return 7; 
                case "msp100000001":
                        return 7;
                case "custom2395404":
                        return 5;
                case "msp100005700":
                        return 5;
                case "custom13995827":
                    return 8;
                case "msp100005896":
                    return 8;
                case "msp100005897"://
                    return 9;//
                case "custom13996126"://
                    return 9;//
                case "custom13731684":
                    return 11;
                case "custom58388298":
                    return 11;
                case "custom83145437": 
                    return 10;
                case "custom6470835":
                    return 12;
                case "custom6470862":
                    return 13;
                default:
                    return 0;
            }
        }
        //========================================================================================
        public string[,] getFullName(int Rows) //
        {
            String connectionString = "Server=10.200.50.13;Port=5432;User ID=reader;Password=ti9gMUKFzQ#iXH{aYMB<;Database=cpgu;";
            NpgsqlConnection connection = new NpgsqlConnection(connectionString);
            using (connection)
            {
                string x = "select id, full_name from cpgu_mfc where id != 109;";
                NpgsqlCommand command = new NpgsqlCommand(x, connection);
                connection.Open();
                NpgsqlDataReader reader = command.ExecuteReader();
                string[,] mass = new string[Rows , 2];
                if (reader.HasRows)
                {
                    int i = 0;
                    while (reader.Read())
                    {
                        mass[i, 0] = reader[0].ToString();//code 
                        mass[i, 1] = reader[1].ToString();//full_name
                        i++;
                    }
                }
                reader.Close();
                connection.Close();
                return mass;
            }
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if ((dp_start.SelectedDate.HasValue) && (dp_finish.SelectedDate.HasValue))
            {
                string date1 = dp_start.Text.Replace(".", "-");
                string date2 = dp_finish.Text.Replace(".", "-");
                char[] tmp = date1.ToCharArray();
                char[] tmp2 = date2.ToCharArray();
                for (int i = 0; i < 4; i++)
                {
                    tmp[i] = date1[i + 6];
                    tmp2[i] = date2[i + 6];
                }
                tmp[4] = '-';
                tmp[5] = date1[3];
                tmp[6] = date1[4];
                tmp[7] = '-';
                tmp[8] = date1[0];
                tmp[9] = date1[1];
                tmp2[4] = '-';
                tmp2[5] = date2[3];
                tmp2[6] = date2[4];
                tmp2[7] = '-';
                tmp2[8] = date2[0];
                tmp2[9] = date2[1];
                date1 = null;
                date2 = null;
                for (int i = 0; i < 10; i++)
                {
                    date1 += tmp[i];
                    date2 += tmp2[i];
                }
                // Теперь дата хранится в стрингах date1 и  date2 в формате гггг-мм-дд
                // Создаём объект документа
                Word.Document doc = null;
                try
                {
                    string UserProfile = Environment.GetEnvironmentVariable("USERPROFILE");
                    Word.Application app = new Word.Application();// Создаём объект приложения
                    string source = UserProfile + @"\corp\Шаблон отчета корпорации.docx"; // Путь до шаблона документа
                    doc = app.Documents.Open(source);
                    doc.Activate();
                    if (CheckBox1.IsChecked == true)
                    {
                        app.Visible = true;
                        Task.Delay(2000).Wait();
                    }
                    Word.Range wRange;
                    //Всего оказано услуг корпорации в период
                    int CountOrders = 0;
                    //CountOrders = Convert.ToInt32(query_db("select count(*) from cpgu_order where(order_date >= '" + date1 + "') and(order_date <= '" + date2 + "') and(service_eid like '%2395626%' or service_eid like 'msp100005699' or service_eid like '%2395697%' or service_eid like 'msp100000001' or service_eid like '%2395404%' or service_eid like 'msp100005700' or service_eid like '%13995827%' or service_eid like 'msp100005896' or  service_eid like '%msp100005897%' or service_eid like '%13996126%' or service_eid like '%13731684%' or service_eid like '%58388298%' or service_eid like '%83145437%'     or service_eid like '%6470835%' or service_eid like '6470862'     )"));
                    CountOrders = Convert.ToInt32(query_db("select count(*) from cpgu_order where(order_date >= '" + date1 + "') and(order_date <= '" + date2 + "') and(service_eid like '%2395626%' or service_eid like 'msp100005699' or service_eid like '%2395697%' or service_eid like 'msp100000001' or service_eid like '%2395404%' or service_eid like 'msp100005700' or service_eid like '%13995827%' or service_eid like 'msp100005896' or  service_eid like '%msp100005897%' or service_eid like '%13996126%' or service_eid like '%13731684%' or service_eid like '%58388298%' or service_eid like '%83145437%'     or service_eid like '%6470835%' or service_eid like '6470862'     )"));
                    //количество строк которые нужно добавить в ворд т.е. кол-во уникальных пар значений: Заявитель=наименование МФЦ
                    //int RowsInWord = Convert.ToInt32(query_db("select count(distinct(requester, reception_place_value)) from cpgu_order where(order_date >= '" + date1 + "') and (order_date <= '" + date2 + "') and(service_eid like '%2395626%' or service_eid like 'msp100005699' or service_eid like '%2395697%' or service_eid like 'msp100000001' or service_eid like '%2395404%' or service_eid like 'msp100005700' or service_eid like '%13995827%' or service_eid like 'msp100005896' or  service_eid like '%msp100005897%' or service_eid like '%13996126%' or service_eid like '%13731684%' or service_eid like '%58388298%' or service_eid like '%83145437%'   or service_eid like '%6470835%' or service_eid like '6470862'     )"));
                    int RowsInWord = Convert.ToInt32(query_db("select count(distinct(requester, reception_place_value)) from cpgu_order where(order_date >= '" + date1 + "') and (order_date <= '" + date2 + "') and(service_eid like '%2395626%' or service_eid like 'msp100005699' or service_eid like '%2395697%' or service_eid like 'msp100000001' or service_eid like '%2395404%' or service_eid like 'msp100005700' or service_eid like '%13995827%' or service_eid like 'msp100005896' or  service_eid like '%msp100005897%' or service_eid like '%13996126%' or service_eid like '%13731684%' or service_eid like '%58388298%' or service_eid like '%83145437%'   or service_eid like '%6470835%' or service_eid like '6470862'     )"));
                    string[,] mass = new string[CountOrders, 3];
                    string x = "select requester.inn, cpgu_order.mfc, cpgu_order.service_eid from cpgu_order join requester on cpgu_order.requester=requester.id AND( (cpgu_order.order_date >= '" + date1 + "') and (cpgu_order.order_date <= '" + date2 + "') and (cpgu_order.service_eid like '%2395626%' or cpgu_order.service_eid like 'msp100005699' or cpgu_order.service_eid like '%2395697%'  or cpgu_order.service_eid like 'msp100000001' or cpgu_order.service_eid like '%2395404%' or cpgu_order.service_eid like 'msp100005700' or cpgu_order.service_eid like '%13995827%' or cpgu_order.service_eid like 'msp100005896' or cpgu_order.service_eid like '%msp100005897%' or cpgu_order.service_eid like '%13996126%' or  cpgu_order.service_eid like '%13731684%' or  cpgu_order.service_eid like '%58388298%' or  cpgu_order.service_eid like '%83145437%' or cpgu_order.service_eid like '%6470835%' or cpgu_order.service_eid like '%6470862%')) order by requester.inn;";
                    mass =getRows(CountOrders, x); //Заполненный значениями массив 
                        //Подменяем в массиве краткое наименование МФЦ на полное
                        //для этого создадим отдельный массив с парами значений (cpgu_order.mfc) | (cpgu_mfc.full_name)
                        int CountMFC = Convert.ToInt32(query_db("select count (*) from cpgu_mfc;")); //Количество строк в таблице со списком мфц
                    string[,] vs = new string[CountMFC + 1, 2];
                    vs = getFullName(CountMFC);
                    //Подменяем в массиве mass значения 
                    for (int i = 0; i < CountOrders; i++)
                        for (int jj = 0; jj < CountMFC; jj++)
                            if ((Convert.ToInt32((mass[i, 1])) == (Convert.ToInt32(vs[jj, 0]))))
                            {
                                mass[i, 1] = vs[jj, 1];
                                break;
                            }
                    Word.Table _table = doc.Tables[1];//Выбираем таблицу в документе. Нумерация таблиц в ворде начинается с 1
                    _table.Rows.Add(ref _missingObj);// Добавляем в таблицу шаблона в ворде пустую строку
                    wRange = _table.Cell(6 , 2).Range;
                    wRange.Text = mass[0, 0];//Заполняем столбец ИНН
                    wRange = _table.Cell(6 , 3).Range;//Заполняем столбец Наименование МФЦ, предоставившего поддержку
                    wRange.Text = mass[0, 1];
                    int j = 0, f = 1, k=0; 
                    wRange = _table.Cell(6 , 4).Range; //заполняем столбец Всего
                    wRange.Text = f.ToString();
                    string inn = mass[0, 0];
                    k = numberofservice(mass[0,2]);
                    wRange = _table.Cell(6, numberofservice(mass[0, 2])).Range; //заполняем столбец 
                    wRange.Text = "1";

                    for (int i = 1; i < CountOrders; i++)
                    {
                        Application.Current.MainWindow.Title =  "Обработано записей: "+i.ToString(); //Чтобы окно не подвисало
                        if (mass[i, 0] != inn) //Если не тот же самый заявитель
                        {
                            inn = mass[i, 0];
                            j++; f = 1;
                            _table.Rows.Add(ref _missingObj);
                            wRange = _table.Cell(6 + j, 2).Range; //заполняем столбец ИНН
                            wRange.Text = mass[i, 0];
                            wRange = _table.Cell(6 + j, 3).Range; //заполняем столбец Наименование МФЦ
                            wRange.Text = mass[i, 1];
                        }
                        else
                            f++;
                        wRange = _table.Cell(6 + j, 4).Range; //заполняем столбец Всего
                        wRange.Text = (f).ToString();
                        k = numberofservice(mass[i, 2]);
                        wRange = _table.Cell(6 + j, k).Range;
                        if ((wRange.Text.Length) == 0)
                            wRange.Text = "1";
                        else
                        {
                            if (wRange.Text.Length != 2)
                                wRange.Text = (1 + Convert.ToInt32(wRange.Text.Substring(0, 2))).ToString();
                            else
                                wRange.Text = "1";
                        }
                    }
                    string SavePath = UserProfile+@"\corp\" + DateTime.Now.ToString("dd'-'MM'-'yyyy';'HH'-'mm'-'ss") + ".docx";
                    doc.SaveAs2(SavePath);
                    doc.Close(); //This is used to close document.
                    app.Quit(); //This is used to quit the Word application.
                    doc = null;
                    MessageBox.Show(@"Готово. Отчет находится в папке "+SavePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Во время выполнения произошла ошибка!");
                    // Если произошла ошибка, то закрываем документ и выводим информацию
                    doc.Close();
                        doc = null;
                }
            }
            else MessageBox.Show("Необходимо выбрать дату!");
        }
    }
}