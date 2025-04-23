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
using NAudio.Wave;
using System.IO;
using System.Management;
using System.Threading;

namespace WpfApp12
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        private WaveIn waveIn;
        private WaveFileWriter writer;
        private string outputFilePath = "recorded.wav";
        private List<WaveInCapabilities> microphoneDevices;
        private ManagementEventWatcher deviceWatcher;
        private SynchronizationContext uiContext;

        public MainWindow()
        {
            InitializeComponent();
            uiContext = SynchronizationContext.Current;
            InitializeMicrophone();
            LoadMicrophoneDevices();
            StartDeviceWatcher();
        }

        private void InitializeMicrophone()
        {
            // 마이크 장치 초기화
            waveIn = new WaveIn();
            waveIn.WaveFormat = new WaveFormat(44100, 1); // 44.1kHz, 모노
            waveIn.DataAvailable += WaveIn_DataAvailable;
        }

        private void LoadMicrophoneDevices()
        {
            uiContext.Post(_ =>
            {
                try
                {
                    microphoneDevices = new List<WaveInCapabilities>();
                    for (int i = 0; i < WaveIn.DeviceCount; i++)
                    {
                        microphoneDevices.Add(WaveIn.GetCapabilities(i));
                    }

                    MicrophoneDevices.ItemsSource = null;
                    MicrophoneDevices.ItemsSource = microphoneDevices;
                    MicrophoneDevices.DisplayMemberPath = "ProductName";
                    
                    if (microphoneDevices.Count > 0)
                    {
                        MicrophoneDevices.SelectedIndex = 0;
                        StartRecording.IsEnabled = true;
                        StopRecording.IsEnabled = true;
                    }
                    else
                    {
                        MicrophoneDevices.Text = "사용 가능한 마이크가 없습니다";
                        StartRecording.IsEnabled = false;
                        StopRecording.IsEnabled = false;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"마이크 목록 로드 오류: {ex.Message}");
                    StartRecording.IsEnabled = false;
                    StopRecording.IsEnabled = false;
                }
            }, null);
        }

        private void StartDeviceWatcher()
        {
            try
            {
                // 장치 연결/해제 이벤트를 모두 감지
                var query = new WqlEventQuery("SELECT * FROM Win32_DeviceChangeEvent WHERE EventType = 2 OR EventType = 3");
                deviceWatcher = new ManagementEventWatcher(query);
                deviceWatcher.EventArrived += DeviceWatcher_EventArrived;
                deviceWatcher.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"장치 감지 오류: {ex.Message}");
            }
        }

        private void DeviceWatcher_EventArrived(object sender, EventArrivedEventArgs e)
        {
            // 장치 변경이 감지되면 잠시 대기 후 마이크 목록을 갱신
            // (장치가 완전히 연결/해제되기를 기다림)
            Thread.Sleep(1000);
            LoadMicrophoneDevices();
        }

        private void WaveIn_DataAvailable(object sender, WaveInEventArgs e)
        {
            if (writer != null)
            {
                writer.Write(e.Buffer, 0, e.BytesRecorded);
            }
        }

        private void StartRecording_Click(object sender, RoutedEventArgs e)
        {
            if (MicrophoneDevices.SelectedIndex == -1)
            {
                MessageBox.Show("마이크 장치를 선택해주세요.");
                return;
            }

            try
            {
                waveIn.DeviceNumber = MicrophoneDevices.SelectedIndex;
                writer = new WaveFileWriter(outputFilePath, waveIn.WaveFormat);
                waveIn.StartRecording();
                StartRecording.IsEnabled = false;
                StopRecording.IsEnabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"녹음 시작 오류: {ex.Message}");
                writer?.Dispose();
                writer = null;
                StartRecording.IsEnabled = true;
                StopRecording.IsEnabled = false;
            }
        }

        private void StopRecording_Click(object sender, RoutedEventArgs e)
        {
            waveIn.StopRecording();
            writer?.Dispose();
            writer = null;
            StartRecording.IsEnabled = true;
            StopRecording.IsEnabled = false;
        }

        protected override void OnClosed(EventArgs e)
        {
            deviceWatcher?.Stop();
            deviceWatcher?.Dispose();
            waveIn?.Dispose();
            writer?.Dispose();
            base.OnClosed(e);
        }
    }
}
