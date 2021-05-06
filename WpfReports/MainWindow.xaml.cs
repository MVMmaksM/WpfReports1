using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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
using System.IO.Compression;
using System.Windows.Threading;
using System.Xml;
using Tulpep.NotificationWindow;
using NLog;

namespace Reports
{
    public partial class MainWindow : Window
    {
        string pathOutOffline, pathXml, pathZip;
        int reportCount, zipCount, errorCount, timerInterval, nullRptCount;

        DispatcherTimer timer = new DispatcherTimer();
        private static Logger logger = LogManager.GetCurrentClassLogger();

        public void CreateDirectory(string path)
        {
            try
            {
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                    logger.Info("Создана директория: " + path);
                }
            }
            catch (Exception exCreate)
            {
                MessageBox.Show(exCreate.Message.ToString(), "Ошибка в при создании директории", MessageBoxButton.OK, MessageBoxImage.Error);
                ButtonStart.IsEnabled = false;

                logger.Error("Ошибка при выполнении метода CreateDirectory " + exCreate.StackTrace);
            }
        }

        public void CreateDirectory(string path, string path2)
        {
            try
            {
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                    logger.Info("Создана директория: " + path);
                }

                if (!Directory.Exists(path2))
                {
                    Directory.CreateDirectory(path2);
                    logger.Info("Создана директория: " + path2);
                }
            }
            catch (Exception exCreate)
            {
                MessageBox.Show(exCreate.Message.ToString(), "Ошибка в при создании директории", MessageBoxButton.OK, MessageBoxImage.Error);
                ButtonStart.IsEnabled = false;

                logger.Error("Ошибка при выполнении метода CreateDirectory " + exCreate.StackTrace);
            }

        }
        public void TulpepArchives(string nameArchive)
        {
            PopupNotifier Tulpep = new PopupNotifier
            {
                TitleText = "Уведомление",
                ContentText = "Создан архив " + nameArchive,
                Delay = 4000
            };
            Tulpep.Click += PopupClick;
            Tulpep.Popup();
        }
        public bool LoadConfig()
        {
            logger.Info("Открытие и чтение config.xml ");

            string pathConfig = $"{Environment.CurrentDirectory}\\config.xml";
            XmlDocument configXml;

            try
            {
                if (File.Exists(pathConfig))
                {
                    configXml = new XmlDocument();
                    configXml.Load(pathConfig);

                    pathOutOffline = configXml.SelectSingleNode("Base/DirectoriesReports/OutOffline").InnerText;
                    pathXml = configXml.SelectSingleNode("Base/DirectoriesReports/Xml").InnerText + "\\xml";
                    pathZip = configXml.SelectSingleNode("Base/DirectoriesReports/Archives").InnerText + "\\archives";
                    timerInterval = Convert.ToInt32(configXml.SelectSingleNode("Base/Settings/TimerInterval").InnerText);

                    return true;
                }
                else
                {
                    MessageBox.Show("Отсутствует файл config.xml", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    ButtonStart.IsEnabled = false;

                    return false;
                }
            }
            catch (Exception exConfig)
            {
                MessageBox.Show(exConfig.Message.ToString(), "Ошибка в config.xml", MessageBoxButton.OK, MessageBoxImage.Error);
                ButtonStart.IsEnabled = false;

                logger.Error("Ошибка при выполнении метода LoadConfig " + exConfig.StackTrace);

                return false;
            }
        }
        public void ReportsMove()
        {
            logger.Info("Вызов метода ReportsMove");

            try
            {
                string[] nameReports = Directory.GetFiles(pathOutOffline, "*.xml");
                logger.Info("Получение списка отчетов: " + nameReports.Length);

                if (nameReports.Length != 0)
                {
                    string pathXmlSave = pathXml + "\\" + DateTime.Now.ToShortDateString() + "\\" + DateTime.Now.ToString("HHmm");
                    string pathZipSave = pathZip + "\\" + DateTime.Now.ToShortDateString();

                    for (int i = 0; i < nameReports.Length; i++)
                    {
                        string nameXml = System.IO.Path.GetFileName(nameReports[i]); // получаем имя первого файла из массива 

                        try
                        {
                            if (new FileInfo(nameReports[i]).Length != 0) // проверка размера файла
                            {
                                CreateDirectory(pathXmlSave);

                                File.Move(nameReports[i], pathXmlSave + "\\" + nameXml);     // исключение если файл уже существует, переход в catch в начало цикла                               

                                LblReportCount.Content = ++reportCount; // подсчет количества перемещенных отчетов
                            }
                            else
                            {
                                CreateDirectory(pathOutOffline + "\\Пустые отчеты"); 

                                File.Move(nameReports[i], pathOutOffline + "\\Пустые отчеты" + "\\" + nameXml); // перемещение файла с 0 кб в папку "Пустые отчеты"

                                LblNullCount.Content = ++nullRptCount; // подсчет количества перемещенных пустых отчетов
                            }
                        }
                        catch (IOException exIO)
                        {
                            logger.Error("Исключение при выполнении метода ReportsMove " + exIO.StackTrace);
                            continue; // пропускает файл, который вызвал исключение и переходит в начало цикла 
                        }
                    }

                    if (Directory.Exists(pathXmlSave) && Directory.GetFiles(pathXmlSave).Length != 0) // проверка существует директория сохранения xml с текущей датой и временем &&                                                                                               
                    {                                                                                 //если существует то проверяется на файлы
                        CreateDirectory(pathZipSave);

                        ZipFile.CreateFromDirectory(pathXmlSave, pathZipSave + "\\" + DateTime.Now.ToString("HHmm") + ".zip");  // исключение если архив уже существует не обработано отдельно
                        logger.Info("Создан архив: " + DateTime.Now.ToString("HHmm"));

                        LblZipCount.Content = ++zipCount;

                        if (ChkBoxNtf.IsChecked == true)
                        {
                            TulpepArchives(DateTime.Now.ToString("HHmm"));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                timer.Stop();
                logger.Info("Остановка таймера");

                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                logger.Error("Ошибка при выполнении метода ReportsMove " + ex.StackTrace);

                if (ButtonStop.IsEnabled)
                {
                    ButtonStop.IsEnabled = false;
                    ButtonStart.IsEnabled = true;
                }

                TimeStop.Content = DateTime.Now.ToShortTimeString();
                ErrorCount.Content = ++errorCount; // подсчет количества ошибок
            }
        }
        public MainWindow()
        {
            InitializeComponent();
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            logger.Info("Запуск программы ");

            if (LoadConfig())
            {
                CreateDirectory(pathXml, pathZip);
            }

            timer.Interval = TimeSpan.FromSeconds(timerInterval);
            timer.Tick += TimeTick;

            LblReportCount.Content = reportCount;
            LblZipCount.Content = zipCount;
            ErrorCount.Content = errorCount;
            LblNullCount.Content = nullRptCount;
        }

        private void ButtonStart_Click(object sender, RoutedEventArgs e)
        {
            ButtonStart.IsEnabled = false;
            ButtonStop.IsEnabled = true;
            TimeStart.Content = DateTime.Now.ToShortTimeString();
            
            timer.Start();
            logger.Info("Запуск таймера");

            ReportsMove();
        }

        private void TimeTick(object sender, EventArgs e)
        {
            ReportsMove();
        }

        private void PopupClick(object sender, EventArgs e)
        {
            Process.Start(pathZip + "\\archives\\" + DateTime.Now.ToShortDateString());
        }
        private void ButtonStop_Click(object sender, RoutedEventArgs e)
        {
            timer.Stop();
            logger.Info("Остановка таймера пользователем");

            ButtonStop.IsEnabled = false;
            ButtonStart.IsEnabled = true;
            TimeStop.Content = DateTime.Now.ToShortTimeString();
        }
    }
}
