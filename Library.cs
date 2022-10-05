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
    /// 
    /// </summary>
    public class Env : INotifyPropertyChanged, IDisposable
    {
        #region Private
        // 刷新 UI 用 Timer
        private System.Timers.Timer _timer;
        // 閒置計時後用 Timer
        private System.Timers.Timer _idleTimer;
        // Disposed 旗標
        private bool _disposed;

        //private TcpClient _tcpClient;
        [Obsolete]
        private TcpListener _tcpListener;
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
        //[BsonId]
        //public ObjectId ObjID { get; set; } 

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

        /// <summary>
        /// 軟體版本
        /// </summary>
        [BsonElement(nameof(SoftVer))]
        public string SoftVer { get; set; } = "2.1.0";

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
        /// 設定 Tcp Listener
        /// </summary>
        [Obsolete("deprecated, will be removed at next version")]
        public void SetTcpListener()
        {
            _cancellationTokenSource = new CancellationTokenSource();
            Task.Factory.StartNew(() =>
            {
                _tcpListener = new TcpListener(System.Net.IPAddress.Parse("0.0.0.0"), 8016);
                _tcpListener.Start();
                byte[] bytes = new byte[256];

                while (!_cancellationTokenSource.IsCancellationRequested)
                {
                    Debug.WriteLine($"Wait for a new connection...");

                        // 等待有 client 連線 || 工作被取消
                        SpinWait.SpinUntil(() => _tcpListener.Pending() || _cancellationTokenSource.IsCancellationRequested);
                        // 若工作被 Cancel，跳出迴圈
                        if (_cancellationTokenSource.IsCancellationRequested) { break; }

                        // 接受 client 連線
                        TcpClient tcpClient = _tcpListener.AcceptTcpClient();

                    Task.Run(() =>
                    {
                        Debug.WriteLine("Connected!");

                        Debug.WriteLine($"{tcpClient.Client.RemoteEndPoint} {tcpClient.Client.LocalEndPoint} {tcpClient.Client.Handle} {Task.CurrentId}");

                            // 取得 NetworkStream
                            NetworkStream networkStream = tcpClient.GetStream();

                            // (i = networkStream.Read(bytes, 0, bytes.Length)).;
                            try
                        {
                                #region 這部分要重寫
#if false
                        //while (_tcpClient.Connected)
                        //{
                        //    Debug.WriteLine($"waiting1");

                        //    // 等待 DataAvailable || 工作被取消
                        //    SpinWait.SpinUntil(() => networkStream.DataAvailable ||  _cancellationTokenSource.IsCancellationRequested, 1000);
                        //    // 若工作被 Cancel，關閉 NetworkStream 與 Client 並跳出迴圈
                        //    if (_cancellationTokenSource.IsCancellationRequested)
                        //    {
                        //        networkStream.Close();
                        //        _tcpClient.Close();
                        //        break;
                        //    }

                        //    Debug.WriteLine($"waiting2");

                        //    int i = networkStream.Read(bytes, 0, bytes.Length);
                        //    string data = System.Text.Encoding.UTF8.GetString(bytes, 0, i);
                        //    Debug.WriteLine($"Data: {data} {_tcpClient.Connected} {i} {networkStream.CanRead}");
                        //}  
#endif

                                //while (!_cancellationTokenSource.IsCancellationRequested)
                                //{
                                //    // 等待 DataAvailabel || 工作被取消；或太久沒有資料傳遞，關閉此連線
                                //    if (SpinWait.SpinUntil(() => networkStream.DataAvailable || _cancellationTokenSource.IsCancellationRequested, 30 * 1000))
                                //    {
                                //        if (_cancellationTokenSource.IsCancellationRequested) { break; }
                                //        i = networkStream.Read(bytes, 0, bytes.Length);
                                //        if (i == 0) { break; }
                                //        string data = System.Text.Encoding.UTF8.GetString(bytes, 0, i);
                                //        Debug.WriteLine($"Data: {data}, i: {i}");
                                //    }
                                //    else
                                //    {
                                //        break;
                                //    }
                                //}

                                //while ((i = networkStream.Read(bytes, 0, bytes.Length)) != 0)
                                //{
                                //    string data = System.Text.Encoding.UTF8.GetString(bytes, 0, i);
                                //    Debug.WriteLine($"Data: {data}, i: {i}");
                                //}

                                #endregion
                            }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"EX: {ex.Message}");
                        }

                        networkStream.Close();
                        networkStream.Dispose();

                        return tcpClient;
                    }).ContinueWith(t =>
                    {
                        Debug.WriteLine(t.Result.Connected);
                        Debug.WriteLine(t.Status);



                        t.Result.Close();
                        t.Dispose();
                    });
                }

                _tcpListener.Stop();
                Debug.WriteLine($"End Server");
            }, TaskCreationOptions.LongRunning);
        }

        /// <summary>
        /// 終止Tcp Listener
        /// </summary>
        [Obsolete("deprecated, will be removed at next version")]
        public void EndTcpListener()
        {
            //_cancellationTokenSource 不為 null 且尚未要求 cancel
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
            // Debug.WriteLine($"{SystemTime}");
            OnPropertyChanged(nameof(SystemTime));

            if (_auto)
            {
                OnPropertyChanged(nameof(AutoTime));
                OnPropertyChanged(nameof(TotalAutoTime));
            }
            // OnPropertyChanged(nameof(IdleTime)); // 之後需移除
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
