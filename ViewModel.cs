using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Input;
using static System.Net.Mime.MediaTypeNames;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace MC_Sim
{
    class ViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private string _code; //標的資產代號
        public string Code
        {
            get { return _code; }
            set
            {
                _code = value;
                OnPropertyChanged("Code");
            }
        }

        private string _s0; //初始價格(float)
        public string S0
        {
            get { return _s0; }
            set
            {
                _s0 = value;
                OnPropertyChanged("S0");
            }
        }

        private string _mu; //年化報酬率(float)
        public string Mu
        {
            get { return _mu; }
            set
            {
                _mu = value;
                OnPropertyChanged("Mu");
            }
        }

        private string _sigma; //年化波動率(float)
        public string Sigma
        {
            get { return _sigma; }
            set
            {
                _sigma = value;
                OnPropertyChanged("Sigma");
            }
        }

        private string _t = "60"; //模擬期限(int)
        public string T
        {
            get { return _t; }
            set
            {
                _t = value;
                OnPropertyChanged("T");
            }
        }

        private string _m; //模擬路徑數量(int)
        public string M
        {
            get { return _m; }
            set
            {
                _m = value;
                OnPropertyChanged("M");
            }
        }

        private string _total_day = "252"; //全年交易天數(int)
        public string Total_day
        {
            get { return _total_day; }
            set
            {
                _total_day = value;
                OnPropertyChanged("Total_day");
            }
        }

        private string _dir = ""; //目錄路徑
        public string Dir
        {
            get { return _dir; }
            set
            {
                _dir = value;
                OnPropertyChanged("Dir");
            }
        }

        public ICommand Tips { get; } //開啟說明視窗
        public ICommand Execute { get; } //執行.exe檔

        public void tips()
        {
            Window1 window1 = new Window1();
            window1.DataContext = this;
            window1.Owner = System.Windows.Application.Current.MainWindow;
            window1.ShowDialog();
        }

        public void execute()
        {
            if (value_check())
            {
                Exe();
            }
            else
            {
                System.Windows.MessageBox.Show("輸入值型態可能有誤!");
            }
        }

        public bool value_check()
        {
            return IsFloat(S0) && IsFloat(Mu) && IsFloat(Sigma) && IsInteger(T) && IsInteger(M) && IsInteger(Total_day);
        }

        public bool IsInteger(string input)
        {
            int result;
            return int.TryParse(input, out result);
        }

        public bool IsFloat(string input)
        {
            float result;
            return float.TryParse(input, out result);
        }

        public void Exe()
        {
            try
            {
                string exeName = "MonteCarlo_sim.exe";
                string exePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, exeName);
                string arguments = $"{Code} {S0} {Mu} {Sigma} {T} {M} {Total_day} {Dir}";
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = exePath,
                    Arguments = arguments,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (Process process = new Process { StartInfo = startInfo })
                {
                    process.Start();
                    string output = process.StandardOutput.ReadToEnd(); // 捕獲標準輸出
                    string error = process.StandardError.ReadToEnd(); // 捕獲標準錯誤輸出
                    process.WaitForExit();
                    System.Windows.MessageBox.Show($"輸入參數: {output}");
                    if (error != "")
                    {
                        System.Windows.MessageBox.Show($"Error: {error}");
                    }
                    else
                    {
                        System.Windows.MessageBox.Show($"資料儲存於{Dir}/output_{Code}.xlsx!");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Error: {ex.Message}");
            }
        } //觸發MC_sim.exe

        public ViewModel()
        {
            Tips = new RelayCommand(tips);
            Execute = new RelayCommand(execute);
        }
    }
}
