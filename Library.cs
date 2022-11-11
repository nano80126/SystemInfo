using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace SystemInfo
{
    /// <summary>
    /// 系統基本資訊
    /// </summary>
    public interface IBasicInfo
    {
        /// <summary>
        /// 作業系統版本
        /// </summary>
        public string OS { get; }

        /// <summary>
        /// 64 / 32 位元
        /// </summary>
        public string Plateform { get; }

        /// <summary>
        /// Process ID
        /// </summary>
        public int PID { get; }

        /// <summary>
        /// 系統時間
        /// </summary>
        public string SystemTime { get; }

        /// <summary>
        /// 系統日期
        /// </summary>
        public string SystemDate { get; }
    }

    /// <summary>
    /// 使用的 SDK Version
    /// </summary>
    public interface ISDKVersion
    {
        #region Properties
        public string DotNetVer { get; }

        public string MongoVer { get; } 
        #endregion

        public void SetMongoVersion(string version);
    }

    public interface IOperationInfo
    {
        #region Properties
        public string Mode { get; }

        public string AutoTime { get; }

        public string TotalAutoTime { get; }

        public int TotalHours { get; }

        public int TotalParts { get; }

        /// <summary>
        /// 閒置時間
        /// </summary>
        public int IdleTime { get; }

        /// <summary>
        /// 是否為 Idle
        /// </summary>
        public bool Idle { get; }
        #endregion

        #region Methods
        public void SetStartTime();

        public void SetTotalAutoTime(int seconds);

        public void SetTotalParts(int parts);

        public void PlusTotalParts();

        public void StartIdleWatch();

        public void StopIdleWatch();

        public int GetAutoTimeInSeconds();
        #endregion
    }


    /// <summary>
    /// 
    /// </summary>
    public class Env : IBasicInfo, ISDKVersion , IOperationInfo, INotifyPropertyChanged, IDisposable
    {
        #region Private
        // 刷新 UI 用 Timer
        private System.Timers.Timer _timer;
        // 閒置計時後用 Timer
        private System.Timers.Timer _idleTimer;
        // Disposed 旗標
        private bool _disposed;

        // Socker for tcp server
        private Socket _socket;
        // Socket list of tcp clients connected
        private List<Socket> _clients = new List<Socket>();
        // Socket Task CancellationTokeSource
        private CancellationTokenSource _cancellationTokenSource;

        //private bool _x64;
        private bool _auto;
        private string _mongoVer = null;
        private DateTime _startTime;
        /// <summary>
        /// 總自動模式時間，每次啟動自動模式時從資料庫讀取
        /// </summary>
        private int _totalAutoTime;
        /// <summary>
        /// 閒置時間計時器
        /// </summary>
        private Stopwatch _stopwatch;
        /// <summary>
        /// 量測總次數
        /// </summary>
        private int _totalParts = 0;
        /// <summary>
        /// idle 旗標
        /// </summary>
        private bool _idle = false;
        #endregion

        #region Properties

        #region BasicInfo
        /// <summary>
        /// 作業系統
        /// </summary>
        [BsonElement(nameof(OS))]
        public string OS => $"{Environment.OSVersion.Version}";

        /// <summary>
        /// 64 / 32 位元
        /// </summary>
        [BsonElement(nameof(Plateform))]
        public string Plateform => Environment.Is64BitProcess ? "64 位元" : "32 位元";

        /// <summary>
        /// Program ID
        /// </summary>
        [BsonIgnore]
        [JsonIgnore]
        public int PID => Environment.ProcessId;

        /// <summary>
        /// 系統時間
        /// </summary>
        [BsonIgnore]
        [JsonIgnore]
        public string SystemTime => $"{DateTime.Now:HH:mm:ss}";

        /// <summary>
        /// 系統日期
        /// </summary>
        [BsonIgnore]
        [JsonIgnore]
        public string SystemDate => $"{DateTime.Now:MM/dd/yy}";
        #endregion

        #region SDK Version
        /// <summary>
        /// .NET 版本
        /// </summary>
        [BsonElement(nameof(DotNetVer))]
        public string DotNetVer => $"{Environment.Version}";

        /// <summary>
        /// MongoDB 版本
        /// </summary>
        [BsonElement(nameof(MongoVer))]
        public string MongoVer => _mongoVer ?? "未連線";

        /// <summary>
        /// Basler Pylon API Versnio Number
        /// </summary>
        [BsonElement(nameof(PylonVer))]
        public string PylonVer => FileVersionInfo.GetVersionInfo("Basler.Pylon.dll").FileVersion;

        /// <summary>
        /// 軟體版本
        /// </summary>
        [BsonElement(nameof(SoftVer))]
        public string SoftVer { get; set; } = "2.1.0";
        #endregion

        #region OperationInfo
        /// <summary>
        /// 建立日期
        /// </summary>
        [BsonIgnore]
        public string BuileTime
        {
            get
            {
                string filePath = Assembly.GetExecutingAssembly().Location;
                return new FileInfo(filePath).LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss");
            }
        }

        /// <summary>
        /// 模式
        /// </summary>
        [BsonIgnore]
        [JsonIgnore]
        public string Mode => _auto ? "自動模式" : "編輯模式";
        /// <summary>
        /// 自動運行時間
        /// </summary>
        [BsonIgnore]
        [JsonIgnore]
        public string AutoTime
        {
            get
            {
                if (_auto)
                {
                    TimeSpan timeSpan = TimeSpan.FromSeconds((DateTime.Now - _startTime).TotalSeconds - (_stopwatch?.Elapsed.TotalSeconds ?? 0));

                    //Debug.WriteLine($"{timeSpan} {(int)timeSpan.TotalHours:00}:{timeSpan.Minutes:00}:{timeSpan.Seconds:00}");
                    //Debug.WriteLine($"{timeSpan.Seconds} {timeSpan.TotalSeconds}");

                    return $"{(int)timeSpan.TotalHours:00}:{timeSpan.Minutes:00}:{timeSpan.Seconds:00}";
                }
                else
                {
                    return "00:00:00";
                }
            }
        }
        /// <summary>
        /// 自動運行時間 (累計)
        /// </summary>
        [BsonElement(nameof(TotalAutoTime))]
        public string TotalAutoTime
        {
            get
            {
                if (_auto)
                {
                    TimeSpan timeSpan = TimeSpan.FromSeconds((DateTime.Now - _startTime).TotalSeconds - (_stopwatch?.Elapsed.TotalSeconds ?? 0) + _totalAutoTime);
                    return $"{(int)timeSpan.TotalHours:00}:{timeSpan.Minutes:00}:{timeSpan.Seconds:00}";
                }
                else
                {
                    TimeSpan timeSpan = TimeSpan.FromSeconds(_totalAutoTime);
                    return $"{(int)timeSpan.TotalHours:00}:{timeSpan.Minutes:00}:{timeSpan.Seconds:00}";
                }
            }
        }
        /// <summary>
        /// 自動運行時間 (小時)，(TotalAutoTime 超過 9999時，由這邊紀錄)
        /// </summary>
        [BsonElement(nameof(TotalHours))]
        public int TotalHours
        {
            get
            {
                if (_auto)
                {
                    TimeSpan timeSpan = TimeSpan.FromSeconds((DateTime.Now - _startTime).TotalSeconds - (_stopwatch?.Elapsed.TotalSeconds ?? 0) + _totalAutoTime);
                    return (int)timeSpan.TotalHours;
                }
                else
                {
                    TimeSpan timeSpan = TimeSpan.FromSeconds(_totalAutoTime);
                    return (int)timeSpan.TotalHours;
                }
            }
        }
        /// <summary>
        /// 總計檢驗數量
        /// </summary>
        [BsonElement(nameof(TotalParts))]
        public int TotalParts
        {
            get => _totalParts;
            set
            {
                if (value != _totalParts)
                {
                    _totalParts = value;
                    OnPropertyChanged();
                }
            }
        }

        /// <summary>
        /// 閒置時間 (seconds) (無條件捨棄)
        /// </summary>
        [BsonIgnore]
        [JsonIgnore]
        public int IdleTime => _stopwatch != null ? (int)(_stopwatch.ElapsedMilliseconds / 1000.0) : 0;

        [BsonIgnore]
        [JsonIgnore]
        public bool Idle
        {
            get => _idle;
            set
            {
                if (value != _idle)
                {
                    _idle = value;
                    OnPropertyChanged();
                    OnIdle(_idle);
                }
            }
        }
        #endregion

        public static ObservableCollection<NetworkInfo> NetworkInfos { get; } = new ObservableCollection<NetworkInfo>();

        #endregion

        #region 單例模式建構子
        private static readonly Env instance = new Env();

        /// <summary>
        /// 改為 private 避免被另外 new 一個實體出來
        /// </summary>
        private Env()
        {
            NetworkInterface[] interfaces = NetworkInterface.GetAllNetworkInterfaces().OrderBy(x => x.Name).ToArray();

            foreach (NetworkInterface @interface in interfaces)
            {
                if (@interface.NetworkInterfaceType == NetworkInterfaceType.Ethernet)
                {
                    UnicastIPAddressInformation unicastIP = @interface.GetIPProperties().UnicastAddresses.FirstOrDefault(x => x.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork);
                    GatewayIPAddressInformation gatewayIP = @interface.GetIPProperties().GatewayAddresses.FirstOrDefault();

                    if (NetworkInfos.Count == 0)
                    {
                        NetworkInfos.Add(new NetworkInfo(@interface.Name, @interface.OperationalStatus)
                        {
                            IP = $"{unicastIP.Address}",
                            //MAC = $"{string.Join("-", Array.ConvertAll(@interface.GetPhysicalAddress().GetAddressBytes(), x => $"{x:X2}"))}",
                            MAC = $"{@interface.GetPhysicalAddress()}",
                            SubMask = $"{unicastIP.IPv4Mask}",
                            DefaultGetway = $"{gatewayIP?.Address}"
                        });
                    }
                    else
                    {
                        NetworkInfo networkInfo = NetworkInfos[0].Clone();

                        networkInfo.Name = @interface.Name;
                        networkInfo.SetOperationStatus(@interface.OperationalStatus);
                        networkInfo.IP = $"{unicastIP.Address}";
                        networkInfo.MAC = $"{@interface.GetPhysicalAddress()}";
                        networkInfo.SubMask = $"{unicastIP.IPv4Mask}";
                        networkInfo.DefaultGetway = $"{gatewayIP?.Address}";
                        NetworkInfos.Add(networkInfo);
                    }
                }
            }
        }

        public static Env GetInstance()
        {
            return instance;
        } 
        #endregion

        #region Methods
        /// <summary>
        /// 設定 Socket Server
        /// </summary>
        public void SetSocketServer()
        {
            _cancellationTokenSource = new CancellationTokenSource();
            Task.Factory.StartNew(() =>
            {
                IPAddress ip = IPAddress.Any;
                IPEndPoint point = new IPEndPoint(ip, 8016);

                _socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                _socket.Bind(point);
                    // max pending connection
                    _socket.Listen(5);

                while (!_cancellationTokenSource.IsCancellationRequested)
                {
                        // 等待有新連線 pending()
                        SpinWait.SpinUntil(() => _socket.Poll(250, SelectMode.SelectRead) || _cancellationTokenSource.IsCancellationRequested);
                    if (_cancellationTokenSource.IsCancellationRequested) break;

                    Socket client = _socket.Accept();
                    _clients.Add(client);

                    Task.Run(() =>
                    {
                        byte[] bytes = new byte[256];
                        while (!_cancellationTokenSource.IsCancellationRequested)
                        {
                            try
                            {
                                    // 1. 有新資料 2. 連線已關閉 3. 工作被取消 ex. 10分鐘沒有新資料
                                    if (SpinWait.SpinUntil(() => client.Poll(100, SelectMode.SelectRead) || _cancellationTokenSource.IsCancellationRequested, 10 * 1000))
                                {
                                    if (_cancellationTokenSource.IsCancellationRequested) { break; }

                                    int i = client.Receive(bytes);
                                        // 遠端 client 關閉連線
                                        if (i == 0) { break; }
                                    string data = System.Text.Encoding.UTF8.GetString(bytes, 0, i);
                                    Debug.WriteLine($"{data}");

                                    byte[] msg = System.Text.Encoding.UTF8.GetBytes($"{DateTime.Now:HH:mm:ss.fff}");
                                    client.Send(msg, msg.Length, SocketFlags.None);
                                }
                                else { break; }
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine(ex.Message);
                            }
                        }
                    }).ContinueWith(t =>
                    {
                        client.Shutdown(SocketShutdown.Both);
                        client.Close();
                        _clients.Remove(client);
                        client.Dispose();
                    });
                }

                return _socket;
            }, TaskCreationOptions.LongRunning).ContinueWith(t =>
            {
                t.Result.Close();
                t.Result.Dispose();
                Debug.WriteLine($"Server socket task status: {t.Status}");
            });
        }

        /// <summary>
        /// 終止 Socket Server
        /// </summary>
        public void EndSocketServer()
        {
            if (_cancellationTokenSource?.IsCancellationRequested == false)
            {
                _cancellationTokenSource.Cancel();
            }
        }

        /// <summary>
        /// 設定 自動/編輯 模式
        /// </summary>
        public void SetMode(bool auto)
        {
            _auto = auto;
        }

        /// <summary>
        /// 設定 Mongo 版本
        /// </summary>
        /// <param name="version">版本</param>
        public void SetMongoVersion(string version = null)
        {
            _mongoVer = version;
        }

        /// <summary>
        /// 設定啟動時間
        /// </summary>
        /// <param name="dateTime">啟動時間</param>
        public void SetStartTime()
        {
            _startTime = DateTime.Now;
        }

        /// <summary>
        /// 設定 Total Auto Time
        /// </summary>
        public void SetTotalAutoTime(int seconds)
        {
            _totalAutoTime = seconds;
        }

        /// <summary>
        /// 設定 Total Parts (主要用於初始化載入)
        /// </summary>
        public void SetTotalParts(int parts)
        {
            _totalParts = parts;
            OnPropertyChanged(nameof(TotalParts));
        }

        /// <summary>
        /// 量測工件計數 +1
        /// </summary>
        public void PlusTotalParts()
        {
            _totalParts += 1;
            OnPropertyChanged(nameof(TotalParts));
        }

        /// <summary>
        /// 開始閒置時間計時器
        /// </summary>
        public void StartIdleWatch()
        {
            if (_stopwatch == null)
            {
                _stopwatch = new Stopwatch();
                _stopwatch.Start();
            }
            else if (_stopwatch?.IsRunning is false)
            {
                _stopwatch.Start();
            }
        }

        /// <summary>
        /// 停止閒置時間計時器
        /// </summary>
        public void StopIdleWatch()
        {
            if (_stopwatch?.IsRunning == true)
            {
                _stopwatch.Stop();
            }
        }

        /// <summary>
        /// 取得 自動運行時間 (秒)
        /// </summary>
        /// <returns></returns>
        public int GetAutoTimeInSeconds()
        {
            if (_auto)
            {
                return (int)((DateTime.Now - _startTime).TotalSeconds - _stopwatch.Elapsed.TotalSeconds);
            }
            else
            {
                return 0;
            }
        }

        /// <summary>
        /// 取得 累計自動運行時間 (秒)
        /// </summary>
        /// <returns></returns>
        public int GetTotalAutoTimeTnSeconds()
        {
            if (_auto)
            {
                return (int)((DateTime.Now - _startTime).TotalSeconds - _stopwatch.Elapsed.TotalSeconds + _totalAutoTime);
            }
            else
            {
                return _totalAutoTime;
            }
        }
        #endregion

        #region 定時執行 UI 刷新 Timer
        /// <summary>
        /// 啟動 Timer (刷新UI用)
        /// </summary>
        public void EnableTimer()
        {
            if (_timer == null)
            {
                _timer = new System.Timers.Timer()
                {
                    Interval = 1000,
                    AutoReset = true,
                };

                _timer.Elapsed += Timer_Elapsed;
                _timer.Start();
            }
            else if (!_timer.Enabled)
            {
                _timer.Start();
            }
        }

        /// <summary>
        /// 計時器 Elapsed 事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            OnPropertyChanged(nameof(SystemTime));

            if (_auto)
            {
                OnPropertyChanged(nameof(AutoTime));
                OnPropertyChanged(nameof(TotalAutoTime));
            }
        }

        /// <summary>
        /// 停止 Timer (刷新UI用)
        /// </summary>
        public void DisableTimer()
        {
            if (_timer?.Enabled == true)
            {
                _timer.Stop();
            }
        }
        #endregion

        #region 閒置計時 Timer
        public void SetIdleTimer(int seconds)
        {
            if (_idleTimer == null)
            {
                _idleTimer = new System.Timers.Timer()
                {
                    Interval = seconds * 1000,
                    AutoReset = false
                };

                _idleTimer.Elapsed += (sender, e) =>
                {
                    Idle = true;

                    StartIdleWatch();
                };
                _idleTimer.Start();
            }
        }

        public void ResetIdlTimer()
        {
            Idle = false;

            StopIdleWatch();

            if (_idleTimer != null)
            {
                _idleTimer.Stop();
                _idleTimer.Start();
            }
        }

        public delegate void IdleChangedEventHandler(object sender, IdleChangedEventArgs e);

        public event IdleChangedEventHandler IdleChanged;

        public class IdleChangedEventArgs : EventArgs
        {
            public bool Idle { get; }

            public IdleChangedEventArgs(bool idle)
            {
                Idle = idle;
            }
        }

        protected void OnIdle(bool idle)
        {
            IdleChanged?.Invoke(this, new IdleChangedEventArgs(idle));
        }
        #endregion

        #region Property Changed Event
        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void PropertyChange(string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion

        #region Dispose
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            if (disposing)
            {
                _timer?.Stop();
                _timer?.Dispose();

                _idleTimer?.Stop();
                _idleTimer?.Dispose();
            }
            _disposed = true;
        } 
        #endregion
    }

    /// <summary>
    /// 網卡資訊
    /// </summary>
    public class NetworkInfo : INotifyPropertyChanged
    {
        #region Private
        private OperationalStatus _status;
        #endregion

        #region Properties
        public string Name { get; set; }

        public string IP { get; set; }

        public string MAC { get; set; }

        public string SubMask { get; set; }

        public string DefaultGetway { get; set; }

        public bool Status => _status == OperationalStatus.Up;
        #endregion

        #region 建構子
        public NetworkInfo(OperationalStatus status)
        {
            _status = status;
        }

        public NetworkInfo(string name, OperationalStatus status)
        {
            Name = name;
            _status = status;
        }
        #endregion

        #region Methods
        public void SetOperationStatus(OperationalStatus status)
        {
            _status = status;
        }

        public NetworkInfo Clone()
        {
            return (NetworkInfo)MemberwiseClone();
        }
        #endregion

        #region Property Changed Event
        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion
    }

    /// <summary>
    /// 網卡資訊集合
    /// </summary>
    [Obsolete]
    public class NetWorkInfoCollection : ObservableCollection<NetworkInfo>
    {
        public NetWorkInfoCollection() : base()
        {
            NetworkInterface[] interfaces = NetworkInterface.GetAllNetworkInterfaces().OrderBy(x => x.Name).ToArray();

            foreach (NetworkInterface @interface in interfaces)
            {
                if (@interface.NetworkInterfaceType == NetworkInterfaceType.Ethernet)
                {
                    UnicastIPAddressInformation unicastIP = @interface.GetIPProperties().UnicastAddresses.FirstOrDefault(x => x.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork);
                    GatewayIPAddressInformation gatewayIP = @interface.GetIPProperties().GatewayAddresses.FirstOrDefault();

                    Add(new NetworkInfo(@interface.Name, @interface.OperationalStatus)
                    {
                        IP = $"{unicastIP.Address}",
                        //MAC = $"{string.Join("-", Array.ConvertAll(@interface.GetPhysicalAddress().GetAddressBytes(), x => $"{x:X2}"))}",
                        MAC = $"{@interface.GetPhysicalAddress()}",
                        SubMask = $"{unicastIP.IPv4Mask}",
                        DefaultGetway = $"{gatewayIP?.Address}"
                    });
                }
            }
        }
    }
}
