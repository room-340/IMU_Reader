// To Do List:
// Откалиброванные блоки: 2, 4, 7(\m), 8, 9, 10, 11(\m), 12(\m), 13, 14, 15
// Пятый НЕ ОТКАЛИБРОВАН.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using System.IO.Ports;
using System.IO;
using System.Threading;

namespace IMU_Reader
{
    /// <summary>
    /// This file contains almost all code for IMU reader project. Only algorythms ocde excluded.
    /// </summary>
    public partial class MainWindow : Window
    {
        string build_version = "0.82 Kalman (basic axis)";
        bool connected_to_port = false;
        int block_index = 0;
        SerialPort active_com = new SerialPort();
        int[] links = new int[128];
        volatile int selected_record = 0;
        volatile string file2save;

        public MainWindow()
        {
            InitializeComponent();
            mainWindow.Title += " v" + build_version;
            button1.Content = "Считать\nданные";
            button2.Content = "Сохранить\nзапись";
            button3.Content = "Очистить";
            set_stage("stage 0");
            button3.Uid = "0";
        }

        private string connect()
        {
            string valid_com = "COM0";
            string[] ports = SerialPort.GetPortNames();
            addText("\n" + get_time() + " Идет подключение к датчику...");
            byte[] buffer = new byte[4];
            foreach (string port in ports)
            {
                //addText("\nПодключение к "+ port);
                try
                {
                    active_com = new SerialPort(port, 3000000, 0, 8, StopBits.One);
                    active_com.WriteBufferSize = 512;
                    active_com.ReadBufferSize = 8192;
                    active_com.Open();
                    active_com.ReadTimeout = 500;
                    active_com.DiscardInBuffer();
                    active_com.Write("v");
                    active_com.Read(buffer, 0, 4);
                    //addText(" - успешно!");
                    addText("\nДатчик успешно подключен.");
                    valid_com = port;
                }
                catch (Exception)
                {
                    //addText(" - неудачно.");
                }
            }
            if (valid_com != "COM0")
                connected_to_port = true;
            else
                active_com.Close();
            return valid_com;
        }

        private void Disconnect()
        {
            active_com.Close();
            connected_to_port = false;
            addText("\n" + get_time() + " Датчик отключен. Можете выключить блок.");
        }

        private string get_time()
        {
            DateTime current = DateTime.Now;
            //string time_stamp = String.Format("{0:d2}:{1:d2}:{2:d2}", current.Hour, current.Minute, current.Second);
            string time_stamp = String.Format("{0:d2}:{1:d2}", current.Hour, current.Minute);
            return time_stamp;
        }

        public void addText(string text)
        {
            Dispatcher.Invoke(new Action(() => textBox1.Text += text));
        }

        private void textBox1_TextChanged(object sender, TextChangedEventArgs e)
        {
            Dispatcher.Invoke(new Action(() => textBox1.ScrollToEnd()));
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            Thread reading = new Thread(read_thread);
            reading.Start();
        }

        private void set_stage(string stage)
        {
            SolidColorBrush active = new SolidColorBrush(Colors.White);
            SolidColorBrush inactive = new SolidColorBrush(Colors.LightGray);
            SolidColorBrush Red = new SolidColorBrush((Color) ColorConverter.ConvertFromString("#FFEC7878"));
            switch (stage)
            {
                case "stage 0":
                    
                    Dispatcher.Invoke(new Action(() => button1.IsEnabled = true));
                    Dispatcher.Invoke(new Action(() => button2.IsEnabled = false));
                    Dispatcher.Invoke(new Action(() => button3.IsEnabled = true));
                    Dispatcher.Invoke(new Action(() => button1.BorderBrush = new SolidColorBrush(Colors.Orange)));
                    Dispatcher.Invoke(new Action(() => button1.Foreground = new SolidColorBrush(Colors.Orange)));
                    Dispatcher.Invoke(new Action(() => button2.BorderBrush = inactive));
                    Dispatcher.Invoke(new Action(() => button2.Foreground = inactive));
                    Dispatcher.Invoke(new Action(() => button3.BorderBrush = Red));
                    Dispatcher.Invoke(new Action(() => button3.Foreground = Red));
                    break;
                case "stage 1":
                    Dispatcher.Invoke(new Action(() => button1.IsEnabled = false));
                    Dispatcher.Invoke(new Action(() => button2.IsEnabled = false));
                    Dispatcher.Invoke(new Action(() => button3.IsEnabled = false));
                    Dispatcher.Invoke(new Action(() => button1.BorderBrush = inactive));
                    Dispatcher.Invoke(new Action(() => button1.Foreground = inactive));
                    Dispatcher.Invoke(new Action(() => button2.BorderBrush = inactive));
                    Dispatcher.Invoke(new Action(() => button2.Foreground = inactive));
                    Dispatcher.Invoke(new Action(() => button3.BorderBrush = inactive));
                    Dispatcher.Invoke(new Action(() => button3.Foreground = inactive));
                    break;
                case "stage 2":
                    Dispatcher.Invoke(new Action(() => button1.IsEnabled = false));
                    Dispatcher.Invoke(new Action(() => button2.IsEnabled = true));
                    Dispatcher.Invoke(new Action(() => button3.IsEnabled = true));
                    Dispatcher.Invoke(new Action(() => button1.BorderBrush = inactive));
                    Dispatcher.Invoke(new Action(() => button1.Foreground = inactive));
                    Dispatcher.Invoke(new Action(() => button2.BorderBrush = new SolidColorBrush(Colors.Orange)));
                    Dispatcher.Invoke(new Action(() => button2.Foreground = new SolidColorBrush(Colors.Orange)));
                    Dispatcher.Invoke(new Action(() => button3.BorderBrush = Red));
                    Dispatcher.Invoke(new Action(() => button3.Foreground = Red));
                    break;
                case "stage 3":
                    break;
                case "stage 4":
                    break;
            }
        }

        private void read_thread(object none)
        {
            Dispatcher.Invoke(new Action(() => set_stage("stage 1")));
            string com_num = connect();
            byte[] buffer = new byte[4096]; // максимальный пакет 4096 байт, взято с запасом
            int numb = 0;
            char[] tempS = new char[200];
            long numfiles = 0;
            long fcur = 0;
            char[] log_name_c = new char[2048];
            System.Diagnostics.Stopwatch timer = new System.Diagnostics.Stopwatch();
            FileStream fs;
            BinaryWriter bin_wr;
            if (!Directory.Exists("C:\\temp\\"))
                Directory.CreateDirectory("C:\\temp\\");

            try
            {
                if ((connected_to_port) && (com_num != "COM0"))
                {
                    addText("\n" + get_time() + " Идет считывание...");
                    active_com.DiscardInBuffer();
                    timer.Start();
                    active_com.ReadTimeout = 1000;
                    active_com.Write("v");    // запрос версии и номера датчика
                    Thread.Sleep(40);
                    numb = active_com.Read(buffer, 0, 4);
                    block_index = buffer[0] * (int)Math.Pow(2, 24) + buffer[1] * (int)Math.Pow(2, 16) + buffer[2] * (int)Math.Pow(2, 8) + buffer[3];
                    //addText("\nНомер блока - " + block_index);
                    active_com.Write("?");  // запрос на количество записей в памяти
                    Thread.Sleep(50);
                    numb = active_com.Read(buffer, 0, 4);
                    numfiles = buffer[0] * (int)Math.Pow(2, 24) + buffer[1] * (int)Math.Pow(2, 16) + buffer[2] * (int)Math.Pow(2, 8) + buffer[3];
                    //addText("\nКоличество файлов - " + numfiles);
                    active_com.Write("c");
                    Thread.Sleep(50);   // Временная задержка, рабочий вариант
                    numb = active_com.Read(buffer, 0, 4);
                    fcur = buffer[0] * (int)Math.Pow(2, 24) + buffer[1] * (int)Math.Pow(2, 16) + buffer[2] * (int)Math.Pow(2, 8) + buffer[3];
                    //addText("\nТекущий файл - " + fcur);
                    //if (fcur > 1)
                    //    addText("\nСброс счетчика файлов...");
                    while (fcur > 1)
                    {
                        active_com.Write("-");
                        active_com.Read(buffer, 0, 4);
                        Thread.Sleep(20);
                        active_com.Write("c");
                        Thread.Sleep(20);
                        numb = active_com.Read(buffer, 0, 4);
                        fcur = buffer[0] * (int)Math.Pow(2, 24) + buffer[1] * (int)Math.Pow(2, 16) + buffer[2] * (int)Math.Pow(2, 8) + buffer[3];
                    }

                    int fsum = 0;
                    Dispatcher.Invoke(new Action(() => progressBar1.Maximum = 15));
                    Dispatcher.Invoke(new Action(() => progressBar1.Value = 1));
                    int size_counter = 0;
                    int packet_count = 0;
                    timer.Start();
                    for (int q = 0; q < numfiles; q++)
                    {
                        //addText("\n" + get_time() + " Считывание " + (q + 1) + " файла...");
                        fsum = 0;
                        size_counter = 0;
                        active_com.Write("r"); // переход в режим чтения
                        Thread.Sleep(40);
                        numb = active_com.Read(buffer, 0, 4);

                        fs = File.Create("C:\\temp\\log_" + (q + 1) + ".txt", 4096, FileOptions.None);
                        bin_wr = new BinaryWriter(fs);

                        active_com.DiscardInBuffer();
                        active_com.DiscardOutBuffer();

                        while (fsum != 4096 * 255)  // выход из цикла если все значения равны FF
                        {
                            fsum = 0;

                            active_com.Write("n"); // any key
                            Thread.Sleep(40);
                            numb = active_com.Read(buffer, 0, 4096);
                            for (int w = 0; w < numb; w++) // sum array functions?
                            {
                                fsum += buffer[w];
                                //bin_wr.Write(buffer[w]);
                            }
                            bin_wr.Write(buffer, 0, numb);
                            size_counter += numb;
                            packet_count++;
                            if (packet_count % 30 == 0)
                            {
                                Dispatcher.Invoke(new Action(() => progressBar1.Maximum += 1));
                                Dispatcher.Invoke(new Action(() => progressBar1.Value += 1));
                            }
                        }
                        bin_wr.Flush();
                        bin_wr.Close();
                        if (size_counter > 4096)
                        {
                            Dispatcher.Invoke(new Action(() => listBox1.Items.Add("Запись " + (listBox1.Items.Count + 1) + " - " + size_counter / 194700 + " мин.")));
                            links[listBox1.Items.Count - 1] = q + 1;
                        }
                        else
                            File.Delete("C:\\temp\\log_" + (q + 1) + ".txt");
                    }
                    Dispatcher.Invoke(new Action(() => progressBar1.Value = progressBar1.Maximum));
                    //addTextaddText("\n" + get_time() + " Чтение завершено! Всего " + timer.ElapsedMilliseconds / 1000 + " секунд!");
                    addText("\n" + get_time() + " Чтение завершено!");
                    System.Media.SystemSounds.Asterisk.Play();
                    timer.Stop();
                    Disconnect();
                    Dispatcher.Invoke(new Action(() => set_stage("stage 2")));
                }
                else // не удалось подключиться к СОМ порту
                {
                    addText("\nДатчик не найден. Проверьте подключение\nUSB провода и убедитесь что датчик включен.");
                    Dispatcher.Invoke(new Action(() => set_stage("stage 0")));
                }
            }
            catch (Exception crit_error)
            {
                System.Media.SystemSounds.Exclamation.Play();
                MessageBox.Show("Произошла критическая ошибка. Возможные причины ошибки:\n" +
                "• выбранный СОМ порт не связан с датчиком СКВП\n" +
                "• датчик выключен\n" +
                "\n Код ошибки: " + crit_error.Message);
                Disconnect();
                Dispatcher.Invoke(new Action(() => set_stage("stage 0")));
                return;
            }
        }

        private void button3_Click(object sender, RoutedEventArgs e)
        {
            button3.Uid = (Convert.ToInt32(button3.Uid) + 1).ToString();
            if (Convert.ToInt32(button3.Uid) == 1)
                addText("\nДля очистки датчика нажмите кнопку очистки еще раз.");
            else
            {
                Thread clearing = new Thread(clear_thread);
                clearing.Start();
            }
        }

        private void clear_thread(object none)
        {
            Dispatcher.Invoke(new Action(() => set_stage("stage 1")));
            string com_num = connect();
            char[] buffer = new char[8192];
            if ((connected_to_port) && (com_num != "COM0"))
            {
                addText("\n" + get_time() + " Идет очистка датчика...");

                active_com.DiscardInBuffer();
                active_com.Write("f");
                addText("\n" + get_time() + " Очистка закончена. Выключите модуль.\nПрограмма закроется через 2 секунды.");
                System.Media.SystemSounds.Asterisk.Play();
                Disconnect();
                Thread.Sleep(2500);
                Dispatcher.Invoke(new Action(() => this.Close()));
            }
            else // не удалось подключиться к СОМ порту
            {
                addText("\nДатчик не найден. Проверьте подключение\nUSB провода и убедитесь что датчик включен.");
                Dispatcher.Invoke(new Action(() => set_stage("stage 0")));
            }
           
        }

        private void button2_Click(object sender, RoutedEventArgs e)
        {
            if (listBox1.SelectedItem != null)
            {
                Microsoft.Win32.SaveFileDialog dialog = new Microsoft.Win32.SaveFileDialog();
                dialog.FileName = "Блок " + block_index + " запись " + (listBox1.SelectedIndex + 1);
                //dialog.DefaultExt = ".";
                dialog.Filter = "Все файлы (*.*)|*.*";
                dialog.Title = "Выберите файл для сохранения данных";
                Nullable<bool> result = dialog.ShowDialog();
                if (result == true)
                {
                    selected_record = links[listBox1.SelectedIndex];
                    file2save = dialog.FileName;
                    Thread saving = new Thread(save_thread);
                    saving.Start();
                }
                else
                {
                    addText("\n" + get_time() + " Не выбран файл для сохранения.");
                }
            }
            else
                addText("\n" + get_time() + " Не выбрана запись для сохранения.");
        }

        private void save_thread(object none)
        {
            int selected = selected_record;
            byte[] full_file = File.ReadAllBytes("C:\\temp\\log_" + selected + ".txt");
            Dispatcher.Invoke(new Action(() => set_stage("stage 1")));
            uint[] ticks = new uint[full_file.Length / 35 + 1];   // примерное количество IMU пакетов
            uint[] ticks2 = new uint[ticks.Length/15];   // примерное количество GPS пакетов неизвестно,                             
            byte[] type = new byte[ticks.Length];   // но оно примерно в 20 раз меньше (взято с запасом)

            byte[] pack = new byte[32];
            byte[] pack2 = new byte[26];
            int crc;
            int k = 0;
            int k2 = 0;
            int tt = 0;
            int[] counter = new int[ticks.Length];
            double[,] a = new double[ticks.Length, 3];
            double[,] w = new double[ticks.Length, 3];
            double[,] m = new double[ticks.Length, 3];
            double[,] q = new double[ticks.Length, 4];
            double[] anglex = new double[ticks.Length];
            double[] angley = new double[ticks.Length];
            double[] anglez = new double[ticks.Length];
            double[] lat = new double[ticks2.Length];
            double[] lon = new double[ticks2.Length];
            double[] speed = new double[ticks2.Length];
            double[] course = new double[ticks2.Length];
            double[] time = new double[ticks2.Length];
            double[] stat = new double[ticks2.Length];
            double[] date = new double[ticks2.Length];
            byte[] buffer = new byte[2];
            byte[] buffer2 = new byte[4];
            double[] corr_mag = {   1, 2.2391, 1, 2.1586, 2.3980,
                                    1, 1, 1.9201, 1.9963, 2.1538,
                                    1, 1, 2.7105, 1.9782, 1.9013 };
            double[] corr_accl = {   1, 0.9778, 1, 0.8928, 1.3899,
                                     1, 0.9722, 0.9766, 0.9599, 0.9379,
                                     0.9183, 0.9924, 1.2657, 0.9892, 0.9598 };
            double[,] corr_gyr = {   {1, 1, 1, 1, 1, 1, 1, 1, 1}, // 1
                                     {1, 1, 1, 1, 1, 1, 1, 1, 1}, // 2
                                     {1, 1, 1, 1, 1, 1, 1, 1, 1}, // 3
                                     {1, 1, 1, 1, 1, 1, 1, 1, 1}, // 4
                                     {-0.00373042367688305, 1.4558979622744, -0.0258818845609887, 0.00460253725775466, 1.4522057450117,
                                         -0.035372440971534, -0.00301630637615598, 1.46124337866467, 0.0112681875599293}, // 5
                                     {1, 1, 1, 1, 1, 1, 1, 1, 1}, // 6
                                     {1, 1, 1, 1, 1, 1, 1, 1, 1}, // 7
                                     {1, 1, 1, 1, 1, 1, 1, 1, 1}, // 8
                                     {1, 1, 1, 1, 1, 1, 1, 1, 1}, // 9
                                     {1, 1, 1, 1, 1, 1, 1, 1, 1}, // 10
                                     {1, 1, 1, 1, 1, 1, 1, 1, 1}, // 11
                                     {1, 1, 1, 1, 1, 1, 1, 1, 1}, // 12
                                     {1, 1, 1, 1, 1, 1, 1, 1, 1}, // 13
                                     {1, 1, 1, 1, 1, 1, 1, 1, 1}, // 14
                                     {1, 1, 1, 1, 1, 1, 1, 1, 1}, // 15
                  }; // Gyro correction coeffs are listed in XYZ order

           addText("\n" + get_time() + " Начало сохранения файла...");
            for (int i = 0; i < full_file.Length - 30; i++)
            {
                if ((i < full_file.Length - 35) && (full_file[i + 34] == 3) && (full_file[i + 33] == 16) &&
                    (full_file[i] == 16) && (full_file[i + 1] == 49))   // условие начала IMU пакета
                {
                    crc = 0;
                    for (int j = 0; j < 32; j++)
                    {
                        pack[j] = full_file[i + j + 1];
                        if (j < 31)
                            crc = crc ^ pack[j];
                    }
                    if (crc == pack[pack.Length - 1])
                    {
                        ticks[k] = BitConverter.ToUInt32(pack, 1);
                        type[k] = pack[0];
                        if (type[k] == 49)
                        {
                            buffer[0] = pack[7]; buffer[1] = pack[8];
                            a[k, 0] = ((double)BitConverter.ToInt16(buffer, 0)) * 0.001766834114354 *corr_accl[block_index - 1];
                            buffer[0] = pack[5]; buffer[1] = pack[6];
                            a[k, 1] = (double)BitConverter.ToInt16(buffer, 0) * 0.001766834114354 *corr_accl[block_index - 1];
                            buffer[0] = pack[9]; buffer[1] = pack[10];
                            a[k, 2] = -(double)BitConverter.ToInt16(buffer, 0) * 0.001766834114354 *corr_accl[block_index - 1];

                            buffer[0] = pack[13]; buffer[1] = pack[14];
                            w[k, 0] = (double)BitConverter.ToInt16(buffer, 0) * 0.00053264;
                            buffer[0] = pack[11]; buffer[1] = pack[12];
                            w[k, 1] = (double)BitConverter.ToInt16(buffer, 0) * 0.00053264;
                            buffer[0] = pack[15]; buffer[1] = pack[16];
                            w[k, 2] = -(double)BitConverter.ToInt16(buffer, 0) * 0.00053264;
                            w[k, 0] = corr_gyr[block_index - 1, 0] * Math.Pow(w[k, 0], 2) + corr_gyr[block_index - 1, 1] * w[k, 0] + corr_gyr[block_index - 1, 2];
                            w[k, 1] = corr_gyr[block_index - 1, 3] * Math.Pow(w[k, 1], 2) + corr_gyr[block_index - 1, 4] * w[k, 1] + corr_gyr[block_index - 1, 5];
                            w[k, 2] = corr_gyr[block_index - 1, 6] * Math.Pow(w[k, 2], 2) + corr_gyr[block_index - 1, 7] * w[k, 2] + corr_gyr[block_index - 1, 8];

                            buffer[0] = pack[17]; buffer[1] = pack[18];
                            m[k, 0] = ((double)BitConverter.ToInt16(buffer, 0) * 0.00030518) * corr_mag[block_index - 1];
                            buffer[0] = pack[19]; buffer[1] = pack[20];
                            m[k, 1] = -((double)BitConverter.ToInt16(buffer, 0) * 0.00030518) * corr_mag[block_index - 1];
                            buffer[0] = pack[21]; buffer[1] = pack[22];
                            m[k, 2] = ((double)BitConverter.ToInt16(buffer, 0) * 0.00030518) * corr_mag[block_index - 1];

                            buffer[0] = pack[23]; buffer[1] = pack[24];
                            q[k, 0] = (double)BitConverter.ToInt16(buffer, 0);
                            buffer[0] = pack[25]; buffer[1] = pack[26];
                            q[k, 1] = (double)BitConverter.ToInt16(buffer, 0) * 0.00003125;
                            buffer[0] = pack[27]; buffer[1] = pack[28];
                            q[k, 2] = (double)BitConverter.ToInt16(buffer, 0) * 0.00003125;
                            buffer[0] = pack[29]; buffer[1] = pack[30];
                            q[k, 3] = (double)BitConverter.ToInt16(buffer, 0) * 0.00003125;
                        }
                        counter[k] = k2;
                        k++;
                    }
                    else
                        tt++;
                }
                if ((full_file[i + 29] == 3) && (full_file[i + 28] == 16) && (full_file[i] == 16) &&
                    (full_file[i + 1] == 50))   // условие начала GPS пакета
                {
                    crc = 50;
                    for (int j = 0; j < 26; j++)
                    {
                        pack2[j] = full_file[i + j + 2];
                        if (j < 25)
                            crc = crc ^ pack2[j];
                    }
                    if (crc == pack2[pack2.Length - 1])
                    {
                        ticks2[k2] = BitConverter.ToUInt32(pack2, 0);
                        buffer2[0] = pack2[4]; buffer2[1] = pack2[5]; buffer2[2] = pack2[6]; buffer2[3] = pack2[7];
                        lat[k2] = ((double)BitConverter.ToInt32(buffer2, 0)) / 600000;
                        buffer2[0] = pack2[8]; buffer2[1] = pack2[9]; buffer2[2] = pack2[10]; buffer2[3] = pack2[11];
                        lon[k2] = ((double)BitConverter.ToInt32(buffer2, 0)) / 600000;
                        buffer[0] = pack2[12]; buffer[1] = pack2[13];
                        speed[k2] = (double)BitConverter.ToInt16(buffer, 0) / 100;
                        buffer[0] = pack2[14]; buffer[1] = pack2[15];
                        course[k2] = (double)BitConverter.ToInt16(buffer, 0) / 160;
                        buffer2[0] = pack2[16]; buffer2[1] = pack2[17]; buffer2[2] = pack2[18]; buffer2[3] = pack2[19];
                        time[k2] = ((double)BitConverter.ToInt32(buffer2, 0)) / 10;
                        stat[k2] = pack2[20];
                        buffer2[0] = pack2[21]; buffer2[1] = pack2[22]; buffer2[2] = pack2[23]; buffer2[3] = pack2[24];
                        date[k2] = ((double)BitConverter.ToInt32(buffer2, 0));
                        k2++;

                    }
                }
            }
            
            // -------------- Сохранение в IMU/GPS
            
            
            double[] angles = new double[3];
            double[] mw, ma, mm;
            ma = new double[3];
            mw = new double[3];
            mm = new double[3];

            FileStream fs_imu = File.Create(file2save + ".imu", 2048, FileOptions.None);
            BinaryWriter str_imu = new BinaryWriter(fs_imu);
            FileStream fs_gps = File.Create(file2save + ".gps", 2048, FileOptions.None);
            BinaryWriter str_gps = new BinaryWriter(fs_gps);
            Int16 buf16; Byte buf8; Int32 buf32;
            Double bufD; Single bufS; UInt32 bufU32;

            addText("\n" + get_time() + " Сохранение файлов:\n" + file2save + ".imu\n" + file2save + ".gps");
            int pbar_adj = 50;
            Dispatcher.Invoke(new Action(() => progressBar1.Maximum = (k+k2)/pbar_adj + 1));
            Dispatcher.Invoke(new Action(() => progressBar1.Value = 0));

            double[] magn_c = get_magn_coefs(block_index);
            double[] accl_c = get_accl_coefs(block_index);
            double[] gyro_c = new double[12];

            double[] w_helper = new double[ticks.Length];

            // Получение уголов путем интегрирования угловых скоростей (простой вариант)
            //for (int i = 0; i < w_helper.Length; i++) w_helper[i] = w[i, 0];
            //anglex = Signal_processing.Zero_average_corr(w_helper, w_helper.Length);
            //for (int i = 0; i < w_helper.Length; i++) w_helper[i] = w[i, 1];
            //angley = Signal_processing.Zero_average_corr(w_helper, w_helper.Length);
            //for (int i = 0; i < w_helper.Length; i++) w_helper[i] = w[i, 2];
            //anglez = Signal_processing.Zero_average_corr(w_helper, w_helper.Length);
            //-----------------------------------------------------------------------------------------------
            // Получение углов путем использования фильтра Калмана (сложный вариант)
            MathNet.Numerics.LinearAlgebra.Double.DenseVector Magn_coefs =
                new MathNet.Numerics.LinearAlgebra.Double.DenseVector(get_magn_coefs(block_index));
            MathNet.Numerics.LinearAlgebra.Double.DenseVector Accl_coefs =
                new MathNet.Numerics.LinearAlgebra.Double.DenseVector(get_accl_coefs(block_index));
            MathNet.Numerics.LinearAlgebra.Double.DenseVector Gyro_coefs = new MathNet.Numerics.LinearAlgebra.Double.DenseVector(12);
            Kalman_class.Parameters Parameters = new Kalman_class.Parameters(Accl_coefs, Magn_coefs, Gyro_coefs);
            System.Func<int, int, double> filler = (x, y) => 0;
            //Kalman_class.Sensors Sensors = new Kalman_class.Sensors(new MathNet.Numerics.LinearAlgebra.Double.DenseMatrix(1, 3, 0),
            //    new MathNet.Numerics.LinearAlgebra.Double.DenseMatrix(1, 3, 0), new MathNet.Numerics.LinearAlgebra.Double.DenseMatrix(1, 3, 0));
            //MathNet.Numerics.LinearAlgebra.Double.Matrix Initia_quat = new MathNet.Numerics.LinearAlgebra.Double.DenseMatrix(1, 4, 0);
            Kalman_class.Sensors Sensors = new Kalman_class.Sensors(MathNet.Numerics.LinearAlgebra.Double.DenseMatrix.Create(1,3, filler),
                MathNet.Numerics.LinearAlgebra.Double.DenseMatrix.Create(1, 3, filler), MathNet.Numerics.LinearAlgebra.Double.DenseMatrix.Create(1, 3, filler));
            MathNet.Numerics.LinearAlgebra.Double.Matrix Initia_quat = MathNet.Numerics.LinearAlgebra.Double.DenseMatrix.Create(1, 4, filler);
            Initia_quat[0, 0] = 1;
            
            Kalman_class.State State = new Kalman_class.State(Kalman_class.ACCLERATION_NOISE, Kalman_class.MAGNETIC_FIELD_NOISE, Kalman_class.ANGULAR_VELOCITY_NOISE,
                Math.Pow(10, -6), Math.Pow(10, -15), Math.Pow(10, -15), Initia_quat);
            Tuple<MathNet.Numerics.LinearAlgebra.Double.Vector, Kalman_class.Sensors, Kalman_class.State> AHRS_result;
            //----------------------------------------------------------------------------------------------
            System.Diagnostics.Stopwatch timer = new System.Diagnostics.Stopwatch();
            timer.Start();
            for (int i = 0; i < k; i++)
            {
                if (i % pbar_adj == 0)
                    Dispatcher.Invoke(new Action(() => progressBar1.Value++));
                if (i == ((int)(k + k2) / 20))
                {
                    long min = (timer.ElapsedMilliseconds * 16) / 1000 / 60;
                    long sec = (timer.ElapsedMilliseconds * 16 - min * 1000 * 60) / 1000;
                    addText("\n" + get_time() + " Сохранение файлов займет около\n" +
                        min + " минут(ы) и " + sec + " секунд(ы).");
                }
                
                //------------------------------------------------------------------
                // Легкий варинат - интегрирование угловых скоростей
                mm = single_correction(magn_c, m[i, 0], m[i, 1], m[i, 2]);
                ma = single_correction(accl_c, a[i, 0], a[i, 1], a[i, 2]);
                mw = single_correction(gyro_c, w[i, 0], w[i, 1], w[i, 2]);

                //angles[0] = (anglez[i]);
                //angles[1] = (angley[i]);
                //angles[2] = (anglex[i]);
                //----------------------------------------------------------------------
                // Сложный вариант - фильтр Калмана
                for (int j = 0; j < 3; j++)
                {
                    Sensors.a.At(0, j, a[i, j]);
                    Sensors.w.At(0, j, w[i, j]);
                    Sensors.m.At(0, j, m[i, j]);
                }
                //Sensors.a.At(0, 0, a[i, 0]);
                //Sensors.a.At(0, 1, a[i, 1]);
                //Sensors.a.At(0, 2, a[i, 2]);

                //Sensors.w.At(0, 0, w[i, 0]);
                //Sensors.w.At(0, 1, w[i, 1]);
                //Sensors.w.At(0, 2, w[i, 2]);

                //Sensors.m.At(0, 0, m[i, 0]);
                //Sensors.m.At(0, 1, m[i, 1]);
                //Sensors.m.At(0, 2, m[i, 2]);

                AHRS_result = Kalman_class.AHRS_LKF_EULER(Sensors, State, Parameters);
                State = AHRS_result.Item3;

                //ma[0] = AHRS_result.Item2.a[0, 0];
                //ma[1] = AHRS_result.Item2.a[0, 1];
                //ma[2] = AHRS_result.Item2.a[0, 2];

                //mw[0] = w[i, 0];
                //mw[1] = w[i, 1];
                //mw[2] = w[i, 2];
                //ma[0] = a[i, 0];
                //ma[1] = a[i, 1];
                //ma[2] = a[i, 2];
                //mm[0] = m[i, 0];
                //mm[1] = m[i, 1];
                //mm[2] = m[i, 2];

                angles[0] = (AHRS_result.Item1.At(0));
                angles[1] = (AHRS_result.Item1.At(1));
                angles[2] = (AHRS_result.Item1.At(2));
                //------------------------------------------------------------------------
           
                // IMU
                buf16 = (Int16)(angles[0] * 10000);
                str_imu.Write(buf16);
                buf16 = (Int16)(angles[1] * 10000);
                str_imu.Write(buf16);
                buf16 = (Int16)(angles[2] * 10000);
                str_imu.Write(buf16);

                buf16 = (Int16)(mw[0] * 3000);
                str_imu.Write(buf16);
                buf16 = (Int16)(mw[1] * 3000);
                str_imu.Write(buf16);
                buf16 = (Int16)(mw[2] * 3000);
                str_imu.Write(buf16);

                buf16 = (Int16)(ma[0] * 3000);
                str_imu.Write(buf16);
                buf16 = (Int16)(ma[1] * 3000);
                str_imu.Write(buf16);
                buf16 = (Int16)(ma[2] * 3000);
                str_imu.Write(buf16);

                buf16 = (Int16)(mm[0] * 3000);
                str_imu.Write(buf16);
                buf16 = (Int16)(mm[1] * 3000);
                str_imu.Write(buf16);
                buf16 = (Int16)(mm[2] * 3000);
                str_imu.Write(buf16);

                buf16 = (Int16)(q[i, 0]);
                str_imu.Write(buf16);

                buf32 = (Int32)(ticks[i]);
                str_imu.Write(buf32);

                buf8 = (Byte)(0);
                str_imu.Write(buf8);
            }


            for (int i = 0; i < k2; i++)
            {
                if (i % pbar_adj == 0)
                    Dispatcher.Invoke(new Action(() => progressBar1.Value++));
                // GPS
                bufD = (Double)(lat[i]) / ((180 / Math.PI) * 16.66);
                str_gps.Write(bufD);
                bufD = (Double)(lon[i]) / ((180 / Math.PI) * 16.66);
                str_gps.Write(bufD);
                bufD = (Double)(0);
                str_gps.Write(bufD);

                bufS = (Single)(time[i]);
                str_gps.Write(bufS);
                bufS = (Single)(speed[i]);
                str_gps.Write(bufS);
                bufS = (Single)(0);
                str_gps.Write(bufS);
                str_gps.Write(bufS);

                bufU32 = (UInt32)(ticks2[i]);
                str_gps.Write(bufU32);
                buf8 = (Byte)(0);
                str_gps.Write(buf8);
                str_gps.Write(buf8);
                str_gps.Write(buf8);
            }
            // Запись даты в конец gps файла
            int day = (int)date[k2 - 1] / 10000;
            int month = (int)(date[k2 - 1] - day * 10000) / 100;
            int year = (int)(2000 + date[k2 - 1] - day * 10000 - month * 100);
            string datarec = String.Format("{0:d2}.{1:d2}.{2:d4}", day, month, year);
            str_gps.Write(datarec);
            str_imu.Flush();
            str_imu.Close();
            str_gps.Flush();
            str_gps.Close();


            Dispatcher.Invoke(new Action(() => progressBar1.Value = (k+k2)/pbar_adj + 1));
            //addText("\n" + get_time() + " Сохранение завершено! Всего " + timer.ElapsedMilliseconds / 1000 + " секунд!");
            addText("\n" + get_time() + " Сохранение завершено!");
            System.Media.SystemSounds.Asterisk.Play();
            timer.Stop();
            Dispatcher.Invoke(new Action(() => set_stage("stage 2")));
        }

        private double[] single_correction(double[] coefs, double xdata, double ydata, double zdata)
        {
            double[] result = new double[3];
            MathNet.Numerics.LinearAlgebra.Double.Matrix B = new MathNet.Numerics.LinearAlgebra.Double.DiagonalMatrix(3, 3, 1);
            MathNet.Numerics.LinearAlgebra.Double.Matrix A = new MathNet.Numerics.LinearAlgebra.Double.DenseMatrix(3, 3);
            A.At(0, 0, coefs[0]);
            A.At(0, 1, coefs[3]);
            A.At(0, 2, coefs[4]);
            A.At(1, 0, coefs[5]);
            A.At(1, 1, coefs[1]);
            A.At(1, 2, coefs[6]);
            A.At(2, 0, coefs[7]);
            A.At(2, 1, coefs[8]);
            A.At(2, 2, coefs[2]);
            MathNet.Numerics.LinearAlgebra.Double.Matrix B1 = Kalman_class.Matrix_Minus(B, A);
            MathNet.Numerics.LinearAlgebra.Double.Matrix C = new MathNet.Numerics.LinearAlgebra.Double.DenseMatrix(3, 1);
            C.At(0, 0, xdata);
            C.At(1, 0, ydata);
            C.At(2, 0, zdata);
            MathNet.Numerics.LinearAlgebra.Double.Matrix D = new MathNet.Numerics.LinearAlgebra.Double.DenseMatrix(3, 1);
            D.At(0, 0, coefs[9]);
            D.At(1, 0, coefs[10]);
            D.At(2, 0, coefs[11]);
            MathNet.Numerics.LinearAlgebra.Double.Matrix res = new MathNet.Numerics.LinearAlgebra.Double.DenseMatrix(3, 1);
            res = Kalman_class.Matrix_Mult(B1, Kalman_class.Matrix_Minus(C, D));
            result[0] = res.At(0, 0);
            result[1] = res.At(1, 0);
            result[2] = res.At(2, 0);
            return result;
        }

        private double[] get_accl_coefs(int index)
        {
            double[] result = new double[0];
            switch (index)
            {
                case 1:
                    double[] temp1 = { -0.0081, 0.0401, -0.0089, 0.0301,
                                        -0.0111, 0.0323, 0.0093, 0.0104, 0.0085, -0.0921, -0.1201, -0.1454 };
                    result = temp1;
                    break;
                case 2: // good for 11.06.2014
                    double[] temp2 = { 0.297316327969713,   0.050055115710092,   0.000023184287949,   0.193383555058825,
                                       0.098300224691001,  -0.062730186495880,   0.143393748688178,  -0.068685981663446,
                                      -0.104205939288307,  -0.822211858352402,  -0.512780968413496,   0.226647239842711 };
                    result = temp2;
                    break;
                case 3:
                    double[] temp3 = {0.024167075678158,   0.017921002182728,   0.027208980951327,   0.047646758555060,
                      -0.052181643992751,  -0.056215499304436,  -0.011304687635132,   0.087617101877347,
                       0.051573793476529,  -0.045096522590058,  -0.034826544109968,  -0.215775421842763 };
                    result = temp3;
                    break;
                case 4: // good for 11.06.2014
                    double[] temp4 = {-0.087441622673382,  -0.101543457753094,  -0.120921146833179,   0.036088401285170,
                                       0.103290700518614,  -0.007421669272914,  -0.030943901137254,  -0.106600016491935,
                                       0.032278058721998,  -0.097460873376706,   0.017422150007582,  -0.406360302723787 };
                    result = temp4;
                    break;
                case 5: // good for 30.06.2014
                    double[] temp5 = { 0.043661195052852,   0.046390535173552,   0.018747392379679,   0.019726515111855,
                                      -0.020799337631274,  -0.005986909839915,  -0.052822916638544,   0.036782243283644,
                                       0.063732368470725,  -0.170738154446748,  -0.153659654203979,  -0.148925283084755 };
                    result = temp5;
                    break;
                case 6:
                    double[] temp6 = {-0.067103542084971,   0.079590227298170,   0.070742925683350,  -0.041636392993342,
                       0.220644176913664,  -0.170330066249381,   0.073407186641070,  -0.024528044153757,
                      -0.047531378473926,   0.078357526154640,  -0.155837767265987,  -0.238145173236917 };
                    result = temp6;
                    break;
                case 7: // good for 01.07.2014
                    double[] temp7 = { 0.003511318264000,  -0.002066738528440,  -0.024519686869331,   0.003044496728777,
                                       0.033363994009801,   0.009855411257946,   0.023173389139447,  -0.038967928071645,
                                      -0.016442948598913,  -0.174871352881015,  -0.043953441045321,  -0.120393700722034 };
                    result = temp7;
                    break;
                case 8: // good for 02.07.2014
                    double[] temp8 = { 0.031157346242867,   0.015814743270598,  -0.022188885151018,  -0.050398249350561,
                                       0.017906182934314,   0.056748261566553,  -0.016875274988413,  -0.063752150437853,
                                       0.035628001358496,  -0.155645692364133,  -0.160953523993233,  -0.090017114827224 };
                    result = temp8;
                    break;
                case 9: // good for 11.06.2014
                    double[] temp9 = {-0.019099213127202,  -0.034197311501348,  -0.068816008010963,  -0.005128473938109,
                                      -0.010630700171860,   0.032957348681911,   0.027784010147100,   0.018816245028074,
                                      -0.020126711893738,  -0.401482168097592,  -0.234689337296154,  -0.038193703490663 };
                    result = temp9;
                    break;
                case 10: // good for 02.07.2014
                    double[] temp10 = {-0.015273646378505,  -0.050455645322452,  -0.082448187615115,   0.151825462763625,
                                        0.029549145696320,  -0.145910515537700,   0.092437851858749,  -0.040261873753222,
                                       -0.078446720622792,  -0.371080641108828,  -0.287723604108068,   0.082385201956462 };
                    result = temp10;
                    break;
                case 11: // should be good
                    double[] temp11 = {-0.019838245402916,  -0.053192576612497,  -0.084668887844949,   0.111462150286516,
                                        0.004430873977272,   0.153013237226223,  -0.002716376135042,   0.042837775216123,
                                        0.002837299906774,  -0.181151396555036,  -0.020642187578603,   0.296833829037690 };
                    result = temp11;
                    break;
                case 12: // good for 02.07.2014
                    double[] temp12 = { 0.020532482397174,   0.015326231885260,  -0.017210544741571,   0.082700531001178,
                                        0.013225472599005,  -0.054464185565828,  -0.042458179793951,  -0.021878238256829,
                                        0.047329772444054,  -0.069394891613174,   0.015828487748667,  -0.301263091039201 };
                    result = temp12;
                    break;
                case 13: // good for 03.07.2014
                    double[] temp13 = {-0.043965199015748,  -0.041637449971824,  -0.054463437699816,   0.070169713825651,
                                       -0.141031463846135,  -0.021815961491165,  -0.168395100622704,   0.129350452261723,
                                        0.167182626033956,  -0.430141334263552,  -0.339436484886449,  -0.190793954173201 };
                    result = temp13;
                    break;
                case 14: // good for 03.07.2014
                    double[] temp14 = { 0.023983851855217,   0.006432666270029,  -0.010095926161812,   0.070805067030385,
                                        0.126580725447670,  -0.036027333617155,   0.097522786860719,  -0.129864215732971,
                                       -0.079006063030160,  -0.369522167853694,  -0.280230396904738,   0.063757449190316 };
                    result = temp14;
                    break;
                case 15: // good for 02.07.2014
                    double[] temp15 = {-0.028848052110869,  -0.042401865267251,  -0.049672725540312,  -0.081002098824795,
                                        0.078414279912455,   0.123970532940648,   0.069653219752965,  -0.082702147671884,
                                       -0.063120255246608,  -0.296040534818233,  -0.298959292294719,   0.004758321035274 };
                    result = temp15;
                    break;
                default:
                    result = new double[12];
                    break;
            }

            return result;
        }

        private double[] get_magn_coefs(int index)
        {
            double[] result = new double[0];
            switch (index)
            {
                case 1:
                    double[] temp1 = { 0.028299362728460,  0.007032267012805, -0.023200541745735, -0.045213515412935,
                                      -0.000578469883817, -0.028325323775602,  0.000715600536868, -0.001895278941303,
                                      -0.004646794130663,  0.039087537663713, -0.018300888706668,  0.052442950622684 };
                    result = temp1;
                    break;
                case 2: // good for 11.06.2014
                    double[] temp2 = {-0.056161489797236,  -0.093250673930749,  -0.127075257039532,  -0.026275677874421,
                                       0.018401174234315,  -0.010316738659001,   0.013402699560750,  -0.027656089897247,
                                      -0.027462576950282,  -0.211883202917799,   0.033020719455609,   0.547078311993916 };
                    result = temp2;
                    break;
                case 3:
                    double[] temp3 = { 0.031091540838056, 0.043044634483806, -0.064899668582821, -0.025218056440666,
                                      -0.004095734923580, 0.010905999083761,  0.002629550752880, -0.003586717217649,
                                      -0.002187269382124, 0.056199172306864, -0.093523132313623,  0.042641416844719 };
                    result = temp3;
                    break;
                case 4: // good for 11.06.2014
                    double[] temp4 = { -0.097270961860931,  -0.114883970525066,  -0.184555111238087,  -0.061388093301272,
                                       -0.002714393898200,   0.081020780290328,   0.025900426180001,  -0.014421765270503,
                                       -0.009017049715442,   0.347226509437727,  -0.193020318041600,  -0.017132562426564 };
                    result = temp4;
                    break;
                case 5: // good for 30.06.2014
                    double[] temp5 = {-0.316785015342110,  -0.286421054821716,  -0.369881696527811,   0.003831311629333,
                                       0.012198510486341,  -0.034501276393751,   0.073246173975953,  -0.008447819912604,
                                      -0.078060318208142,   0.149977431591378,   0.323489866670562,   0.582660085102686 };
                    result = temp5;
                    break;
                case 6:
                    double[] temp6 = { 0.058139581705881, 0.079931661061508, -0.002864001885243, -0.035195541982848,
                                       0.000314672880480, 0.019253549148174,  0.026179418967180,  0.004510178357884,
                                      -0.024468778350392, 0.211145783181474, -0.145806073707686,  0.145807917877855 };
                    result = temp6;
                    break;
                case 7:
                    double[] temp7 = new double[12];
                    result = temp7;
                    break;
                case 8: // good for 02.07.2014
                    double[] temp8 = {-0.212872032485740, -0.204599759415777,  -0.269244668262141,  -0.050316532909885,
                                       0.038549328254905,  0.021176931214368,   0.039695820982477,  -0.038139072311161,
                                      -0.052005644651056,  0.486843501622204,   0.454123489421682,   0.353142185252624 };
                    result = temp8;
                    break;
                case 9: // good for 11.06.2014
                    double[] temp9 = {-0.239562664926672,  -0.237594798761519,  -0.276540918045356,   0.053808584804510,
                                       0.014826232353189,  -0.065454004262785,  -0.066950070215670,  -0.009203514272912,
                                       0.059782306778317,  -0.200480223165631,   0.468017076025475,  -0.288159443551576 };
                    result = temp9;
                    break;
                case 10: // good for 02.07.2014
                    double[] temp10 = { -0.111131070269000,  -0.080465064713768,  -0.165045096635441,  -0.003987526412547,
                                        -0.001066567359500,  -0.020334123123012,   0.045654391595861,   0.013247803209313,
                                        -0.034551806055709,  -0.059601116667666,   0.136285834728099,   0.618891949444801 };
                    result = temp10;
                    break;
                case 11:
                    double[] temp11 = new double[12];
                    result = temp11;
                    break;
                case 12:
                    double[] temp12 = new double[12];
                    result = temp12;
                    break;
                case 13: // good for 03.07.2014
                    double[] temp13 = {-0.237480932628149,  -0.187532885497650,  -0.209210035778355,   0.004460999672365,
                                        0.002031985298151,  -0.035507392376370,   0.078190471121499,  -0.014423422731184,
                                       -0.084796126816155,   0.467055084642093,   0.027626929404883,   0.558124436821596 };
                    result = temp13;
                    break;
                case 14:// good for 03.07.2014
                    double[] temp14 = {-0.145480579380709,  -0.147459040240400,  -0.181966026158069,  -0.045019949762378,
                                        0.054289406364952,   0.037207275210514,   0.023686151925315,  -0.061393620995850,
                                       -0.035296583809944,   0.390877048540474,  -0.173834207503900,   0.566994213355738 };
                    result = temp14;
                    break;
                case 15: // good for 02.07.2014
                    double[] temp15 = {-0.224959909455451,  -0.194546924244309,  -0.267863768104781,  -0.032857168286001,
                                        0.017820397248977,   0.015525756991128,   0.007398257888608,  -0.024740642255649,
                                       -0.012139892860486,   0.448518797201737,   0.239284923960286,   0.473442337769427 };
                    result = temp15;
                    break;
                default:
                    result = new double[12];
                    break;
            }
            return result;
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Disconnect();
            for (int j = 0; j < listBox1.Items.Count; j++)
            {
                File.Delete("C:\\temp\\log_" + links[j] + ".txt");
            }
        }

    }
}
