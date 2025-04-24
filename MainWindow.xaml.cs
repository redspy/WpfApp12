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
using NAudio.CoreAudioApi;
using System.IO;
using System.Management;
using System.Threading;
using System.Runtime.InteropServices;

namespace WpfApp12
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    [ComImport]
    [Guid("870AF99C-171D-4F9E-AF0D-E63DF40C2BC9")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPolicyConfig
    {
        [PreserveSig]
        int GetMixFormat(string pszDeviceName, IntPtr ppFormat);

        [PreserveSig]
        int GetDeviceFormat(string pszDeviceName, int bDefault, IntPtr ppFormat);

        [PreserveSig]
        int SetDeviceFormat(string pszDeviceName, IntPtr pEndpointFormat, IntPtr MixFormat);

        [PreserveSig]
        int GetProcessingPeriod(string pszDeviceName, int bDefault, IntPtr pmftDefaultPeriod, IntPtr pmftMinimumPeriod);

        [PreserveSig]
        int SetProcessingPeriod(string pszDeviceName, IntPtr pmftPeriod);

        [PreserveSig]
        int GetShareMode(string pszDeviceName, IntPtr pMode);

        [PreserveSig]
        int SetShareMode(string pszDeviceName, IntPtr mode);

        [PreserveSig]
        int GetPropertyValue(string pszDeviceName, IntPtr key, IntPtr pv);

        [PreserveSig]
        int SetPropertyValue(string pszDeviceName, IntPtr key, IntPtr pv);

        [PreserveSig]
        int SetDefaultEndpoint(string wszDeviceId, Role eRole);

        [PreserveSig]
        int SetEndpointVisibility(string pszDeviceName, int bVisible);
    }

    [ComImport]
    [Guid("568b9108-44bf-40b4-9006-86afe5b5a620")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPolicyConfigVista
    {
        [PreserveSig]
        int GetMixFormat(string pszDeviceName, IntPtr ppFormat);

        [PreserveSig]
        int GetDeviceFormat(string pszDeviceName, int bDefault, IntPtr ppFormat);

        [PreserveSig]
        int SetDeviceFormat(string pszDeviceName, IntPtr pEndpointFormat, IntPtr MixFormat);

        [PreserveSig]
        int GetProcessingPeriod(string pszDeviceName, int bDefault, IntPtr pmftDefaultPeriod, IntPtr pmftMinimumPeriod);

        [PreserveSig]
        int SetProcessingPeriod(string pszDeviceName, IntPtr pmftPeriod);

        [PreserveSig]
        int GetShareMode(string pszDeviceName, IntPtr pMode);

        [PreserveSig]
        int SetShareMode(string pszDeviceName, IntPtr mode);

        [PreserveSig]
        int GetPropertyValue(string pszDeviceName, IntPtr key, IntPtr pv);

        [PreserveSig]
        int SetPropertyValue(string pszDeviceName, IntPtr key, IntPtr pv);

        [PreserveSig]
        int SetDefaultEndpoint(string wszDeviceId, ERole eRole);

        [PreserveSig]
        int SetEndpointVisibility(string pszDeviceName, int bVisible);
    }

    [ComImport]
    [Guid("294935CE-F637-4E7C-A41B-AB255460B862")]
    internal class CPolicyConfigVistaClient
    {
    }

    public enum ERole
    {
        eConsole,
        eMultimedia,
        eCommunications,
        ERole_enum_count
    }

    public partial class MainWindow : Window
    {
        [DllImport("ole32.dll")]
        private static extern int PropVariantClear(IntPtr pvar);

        [DllImport("ole32.dll")]
        private static extern int PropVariantClear(ref PROPVARIANT pvar);

        [StructLayout(LayoutKind.Sequential)]
        private struct PROPVARIANT
        {
            public ushort vt;
            public ushort wReserved1;
            public ushort wReserved2;
            public ushort wReserved3;
            public IntPtr p;
        }

        private WaveIn waveIn;
        private WaveFileWriter writer;
        private string outputFilePath = "recorded.wav";
        private List<WaveInCapabilities> microphoneDevices;
        private ManagementEventWatcher deviceWatcher;
        private SynchronizationContext uiContext;
        private bool isSettingDefaultDevice = false;

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
                        // 기본 마이크 찾기
                        var enumerator = new MMDeviceEnumerator();
                        var defaultDevice = enumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Console);
                        string defaultDeviceName = defaultDevice.FriendlyName;

                        // 기본 마이크와 일치하는 장치 찾기
                        int defaultDeviceIndex = -1;
                        for (int i = 0; i < microphoneDevices.Count; i++)
                        {
                            if (microphoneDevices[i].ProductName.Contains(defaultDeviceName))
                            {
                                defaultDeviceIndex = i;
                                break;
                            }
                        }

                        // 기본 마이크를 찾았으면 선택, 아니면 첫 번째 장치 선택
                        isSettingDefaultDevice = true;
                        MicrophoneDevices.SelectedIndex = defaultDeviceIndex != -1 ? defaultDeviceIndex : 0;
                        isSettingDefaultDevice = false;
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

        private void MicrophoneDevices_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (isSettingDefaultDevice || MicrophoneDevices.SelectedIndex == -1)
                return;

            try
            {
                var enumerator = new MMDeviceEnumerator();
                var devices = enumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active);
                var selectedDeviceName = microphoneDevices[MicrophoneDevices.SelectedIndex].ProductName;

                foreach (var device in devices)
                {
                    if (device.FriendlyName.Contains(selectedDeviceName))
                    {
                        // 모든 역할에 대해 기본 장치로 설정
                        var policyConfig = (IPolicyConfigVista)new CPolicyConfigVistaClient();
                        policyConfig.SetDefaultEndpoint(device.ID, ERole.eConsole);
                        policyConfig.SetDefaultEndpoint(device.ID, ERole.eMultimedia);
                        policyConfig.SetDefaultEndpoint(device.ID, ERole.eCommunications);
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"기본 장치 설정 오류: {ex.Message}");
            }
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

    [ComImport]
    [Guid("870AF99C-171D-4F9E-AF0D-E63DF40C2BC9")]
    internal class PolicyConfig
    {
    }
}
