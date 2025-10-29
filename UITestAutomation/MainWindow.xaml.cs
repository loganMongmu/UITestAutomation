using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;

namespace MouseMacroWpf
{
    public partial class MainWindow : Window
    {
        // ===== Models =========================================================
        public enum EventType { MouseClick, KeyDown, KeyUp }

        public class InputEvent
        {
            public EventType Type { get; set; } = EventType.MouseClick;
            // Mouse
            public int X { get; set; }
            public int Y { get; set; }
            public int Button { get; set; } = 0; // 0=Left
            // Keyboard
            public int VkCode { get; set; }
            public int ScanCode { get; set; }
            // Common
            public int DelayMs { get; set; } = 1000;
        }

        public class Calibration
        {
            public int OffsetX { get; set; }
            public int OffsetY { get; set; }
            public int ContentW { get; set; }
            public int ContentH { get; set; }

            [JsonIgnore]
            public bool IsValid => ContentW > 0 && ContentH > 0;
        }

        // 새 JSON 루트(캘리브+이벤트). 과거 배열(JSON List)과 호환됨.
        public class MacroFile
        {
            public Calibration? Calibration { get; set; }
            public List<InputEvent> Events { get; set; } = new();
        }

        private List<InputEvent> _events = new();
        private Calibration _calib = new(); // 기본 0 / invalid

        // ===== Recording state ===============================================
        private IntPtr _mouseHookId = IntPtr.Zero;
        private IntPtr _kbdHookId = IntPtr.Zero;
        private LowLevelMouseProc _mouseProc;
        private LowLevelKeyboardProc _kbdProc;
        private IntPtr _targetWindow = IntPtr.Zero;

        // ===== Calibration state ==============================================
        private bool _calibActive = false;
        private POINT? _calibPtTL = null; // client coords
        private POINT? _calibPtBR = null; // client coords

        // ===== Playback state ================================================
        private CancellationTokenSource? _playbackCts;
        private const int HOTKEY_ID = 0x1000;
        private bool _hotkeyRegistered = false;

        public MainWindow()
        {
            InitializeComponent();
            _mouseProc = MouseHookCallback;
            _kbdProc = KeyboardHookCallback;
            UpdateUiState();
            UpdateCalibLabels();
        }

        // ===== UI helpers =====================================================
        private void UI(Action a) => Dispatcher.Invoke(a);

        private void UpdateUiState()
        {
            BtnStopPlay.IsEnabled = false;
            BtnStopRecord.IsEnabled = false;
            BtnRecord.IsEnabled = true;
            BtnPlay.IsEnabled = _events.Count > 0;
            BtnSave.IsEnabled = _events.Count > 0;
        }

        private void UpdateCalibLabels()
        {
            if (_calib != null && _calib.IsValid)
            {
                LblOffX.Text = _calib.OffsetX.ToString();
                LblOffY.Text = _calib.OffsetY.ToString();
                LblCW.Text = _calib.ContentW.ToString();
                LblCH.Text = _calib.ContentH.ToString();
                LblCalibHint.Text = "Calibration: Active";
            }
            else
            {
                LblOffX.Text = LblOffY.Text = LblCW.Text = LblCH.Text = "-";
                LblCalibHint.Text = "Click Start Calibration, then click CONTENT Top-Left and Bottom-Right of target window.";
            }
        }

        private void RefreshListAndJson()
        {
            ListEvents.Items.Clear();
            for (int i = 0; i < _events.Count; i++)
            {
                var ev = _events[i];
                string line = ev.Type switch
                {
                    EventType.MouseClick => $"{i + 1}: MOUSE ({ev.X},{ev.Y}), delay={ev.DelayMs}ms",
                    EventType.KeyDown => $"{i + 1}: KEYDOWN VK={ev.VkCode} scan={ev.ScanCode}, delay={ev.DelayMs}ms",
                    EventType.KeyUp => $"{i + 1}: KEYUP   VK={ev.VkCode} scan={ev.ScanCode}, delay={ev.DelayMs}ms",
                    _ => $"{i + 1}: ?"
                };
                ListEvents.Items.Add(line);
            }

            // JSON 미리보기: 캘리브 유효하면 객체형식, 아니면 레거시 배열
            if (_calib != null && _calib.IsValid)
            {
                var mf = new MacroFile { Calibration = _calib, Events = _events };
                TxtJson.Text = JsonSerializer.Serialize(mf, new JsonSerializerOptions { WriteIndented = true });
            }
            else
            {
                TxtJson.Text = JsonSerializer.Serialize(_events, new JsonSerializerOptions { WriteIndented = true });
            }
        }

        private void TryLoadEventsFromJsonText()
        {
            try
            {
                var text = TxtJson.Text?.Trim();
                if (string.IsNullOrWhiteSpace(text)) return;

                if (text.StartsWith("{"))
                {
                    var mf = JsonSerializer.Deserialize<MacroFile>(text);
                    if (mf != null)
                    {
                        _events = mf.Events ?? new();
                        _calib = mf.Calibration ?? new();
                        UpdateCalibLabels();
                        RefreshListAndJson();
                    }
                }
                else if (text.StartsWith("["))
                {
                    var parsed = JsonSerializer.Deserialize<List<InputEvent>>(text);
                    if (parsed != null)
                    {
                        _events = parsed;
                        // 캘리브는 유지
                        RefreshListAndJson();
                    }
                }
            }
            catch
            {
                // 편집 중 오류 무시
            }
        }

        // ===== Buttons ========================================================
        private void BtnRecord_Click(object sender, RoutedEventArgs e)
        {
            var title = Microsoft.VisualBasic.Interaction.InputBox(
                "window title to be titled (partial strings allowed)\n ex: \"Chrome\" ",
                "Window Title", "");
            title = (title ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(title)) return;

            _targetWindow = FindWindowByTitleContains(title);
            if (_targetWindow == IntPtr.Zero)
            {
                if (MessageBox.Show("Could not find the window. Would you like to record in absolute screen coordinates?",
                    "Window not found", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No)
                    return;
            }
            StartRecording();
        }

        private void BtnStopRecord_Click(object sender, RoutedEventArgs e) => StopRecording();

        private async void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.SaveFileDialog { Filter = "JSON files (*.json)|*.json", DefaultExt = ".json" };
            if (dlg.ShowDialog() == true)
            {
                foreach (var ev in _events) if (ev.DelayMs <= 0) ev.DelayMs = 1000;

                string json;
                if (_calib != null && _calib.IsValid)
                    json = JsonSerializer.Serialize(new MacroFile { Calibration = _calib, Events = _events },
                                                    new JsonSerializerOptions { WriteIndented = true });
                else
                    json = JsonSerializer.Serialize(_events, new JsonSerializerOptions { WriteIndented = true });

                await File.WriteAllTextAsync(dlg.FileName, json);
                UI(() => TxtStatus.Text = $"Saved: {dlg.FileName}");
            }
        }

        private void BtnLoad_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog { Filter = "JSON files (*.json)|*.json", DefaultExt = ".json" };
            if (dlg.ShowDialog() == true)
            {
                try
                {
                    var text = File.ReadAllText(dlg.FileName).Trim();
                    if (text.StartsWith("{"))
                    {
                        var mf = JsonSerializer.Deserialize<MacroFile>(text) ?? new MacroFile();
                        _events = mf.Events ?? new();
                        _calib = mf.Calibration ?? new();
                    }
                    else
                    {
                        _events = JsonSerializer.Deserialize<List<InputEvent>>(text) ?? new();
                        // _calib 유지
                    }
                    UpdateCalibLabels();
                    RefreshListAndJson();
                    UI(() => TxtStatus.Text = $"Loaded: {dlg.FileName}");
                }
                catch (Exception ex) { UI(() => MessageBox.Show("JSON Load Fail: " + ex.Message)); }
            }
        }

        private void BtnPlay_Click(object sender, RoutedEventArgs e)
        {
            if (_events.Count == 0) { UI(() => MessageBox.Show("Please record first or load JSON.")); return; }
            StartPlayback();
        }

        private void BtnStopPlay_Click(object sender, RoutedEventArgs e) => StopPlayback();

        private void ListEvents_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ListEvents.SelectedIndex >= 0 &&
                MessageBox.Show("Delete selected event?", "Delete", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                _events.RemoveAt(ListEvents.SelectedIndex);
                RefreshListAndJson();
                UpdateUiState();
            }
        }

        private void ChkUseAdb_Checked(object s, RoutedEventArgs e) { TxtDevW.IsEnabled = true; TxtDevH.IsEnabled = true; }
        private void ChkUseAdb_Unchecked(object s, RoutedEventArgs e) { TxtDevW.IsEnabled = false; TxtDevH.IsEnabled = false; }

        // ===== Calibration UI =================================================
        private void BtnCalibStart_Click(object sender, RoutedEventArgs e)
        {
            if (_targetWindow == IntPtr.Zero)
            {
                MessageBox.Show("Select the window you want to record (Record button).");
                return;
            }
            _calibActive = true;
            _calibPtTL = null;
            _calibPtBR = null;
            LblCalibHint.Text = "Calibration: Click CONTENT Top-Left, then Bottom-Right (2 clicks).";
            UI(() => TxtStatus.Text = "Calibration waiting for 2 clicks...");
            // 마우스 훅이 이미 설치되지 않았다면 잠시 설치해서 2클릭만 수집하는 방식도 가능하지만,
            // 여기선 사용자가 Record 전/후 어느때나 할 수 있도록 항상 MouseHook에 처리 로직을 둡니다.
            if (_mouseHookId == IntPtr.Zero) _mouseHookId = NativeMethods.SetMouseHook(_mouseProc);
        }

        private void BtnCalibClear_Click(object sender, RoutedEventArgs e)
        {
            _calib = new Calibration();
            UpdateCalibLabels();
            UI(() => TxtStatus.Text = "Calibration cleared.");
        }

        private void FinalizeCalibrationFromPoints(POINT tl, POINT br)
        {
            // tl/br 은 "클라이언트 기준" 좌표
            int x1 = Math.Min(tl.x, br.x);
            int y1 = Math.Min(tl.y, br.y);
            int x2 = Math.Max(tl.x, br.x);
            int y2 = Math.Max(tl.y, br.y);

            _calib = new Calibration
            {
                OffsetX = x1,
                OffsetY = y1,
                ContentW = Math.Max(1, x2 - x1),
                ContentH = Math.Max(1, y2 - y1)
            };
            _calibActive = false;
            UpdateCalibLabels();
            UI(() => TxtStatus.Text = "Calibration saved.");
        }

        // ===== Recording (mouse + keyboard) ==================================
        private void StartRecording()
        {
            _events.Clear();

            if (_mouseHookId == IntPtr.Zero) _mouseHookId = NativeMethods.SetMouseHook(_mouseProc);
            if (_kbdHookId == IntPtr.Zero) _kbdHookId = NativeMethods.SetKeyboardHook(_kbdProc);

            BtnRecord.IsEnabled = false;
            BtnStopRecord.IsEnabled = true;
            BtnSave.IsEnabled = false;
            BtnLoad.IsEnabled = false;
            BtnPlay.IsEnabled = false;
            UI(() => TxtStatus.Text = "Recording... (ESC: stop)");

            var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(100) };
            timer.Tick += (s, ev) =>
            {
                if (Keyboard.IsKeyDown(Key.Escape))
                {
                    timer.Stop();
                    StopRecording();
                }
            };
            timer.Start();
        }

        private void StopRecording()
        {
            if (_mouseHookId != IntPtr.Zero) { NativeMethods.UnhookWindowsHookEx(_mouseHookId); _mouseHookId = IntPtr.Zero; }
            if (_kbdHookId != IntPtr.Zero) { NativeMethods.UnhookWindowsHookEx(_kbdHookId); _kbdHookId = IntPtr.Zero; }

            foreach (var ev in _events) if (ev.DelayMs <= 0) ev.DelayMs = 1000;

            RefreshListAndJson();
            BtnRecord.IsEnabled = true;
            BtnStopRecord.IsEnabled = false;
            BtnSave.IsEnabled = _events.Count > 0;
            BtnLoad.IsEnabled = true;
            BtnPlay.IsEnabled = _events.Count > 0;
            UI(() => TxtStatus.Text = $"Recording stopped. {_events.Count} events.");
        }

        private IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (nCode >= 0)
                {
                    const int WM_LBUTTONDOWN = 0x0201;
                    if (wParam.ToInt32() == WM_LBUTTONDOWN)
                    {
                        var info = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);
                        int sx = info.pt.x, sy = info.pt.y; // screen coords

                        // 대상 창 클라이언트로 변환
                        int cx = sx, cy = sy;
                        if (_targetWindow != IntPtr.Zero)
                        {
                            var pt = new POINT { x = sx, y = sy };
                            if (NativeMethods.ScreenToClient(_targetWindow, ref pt)) { cx = pt.x; cy = pt.y; }
                        }

                        // 캘리브레이션 수집 모드일 때: 두 점만 수집
                        if (_calibActive)
                        {
                            if (_calibPtTL is null) _calibPtTL = new POINT { x = cx, y = cy };
                            else if (_calibPtBR is null) _calibPtBR = new POINT { x = cx, y = cy };

                            if (_calibPtTL is not null && _calibPtBR is not null)
                                UI(() => FinalizeCalibrationFromPoints(_calibPtTL.Value, _calibPtBR.Value));
                            else
                                UI(() => TxtStatus.Text = "Calibration: click Bottom-Right.");
                            return NativeMethods.CallNextHookEx(_mouseHookId, nCode, wParam, lParam);
                        }

                        // 일반 녹화: 좌표 저장 (캘리브 유효 시 컨텐츠 좌표로 정규화)
                        if (_calib.IsValid)
                        {
                            cx -= _calib.OffsetX;
                            cy -= _calib.OffsetY;
                        }

                        _events.Add(new InputEvent { Type = EventType.MouseClick, X = cx, Y = cy, Button = 0, DelayMs = 1000 });
                        UI(RefreshListAndJson);
                    }
                }
            }
            catch { /* ignore */ }

            return NativeMethods.CallNextHookEx(_mouseHookId, nCode, wParam, lParam);
        }

        private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (nCode >= 0)
                {
                    const int WM_KEYDOWN = 0x0100, WM_KEYUP = 0x0101;
                    int msg = wParam.ToInt32();
                    if (msg == WM_KEYDOWN || msg == WM_KEYUP)
                    {
                        var info = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
                        int vk = (int)info.vkCode;

                        // 정지 핫키(Ctrl+Alt+X)는 녹화 제외
                        bool isCtrl = Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl);
                        bool isAlt = Keyboard.IsKeyDown(Key.LeftAlt) || Keyboard.IsKeyDown(Key.RightAlt);
                        bool isX = vk == KeyInterop.VirtualKeyFromKey(Key.X);
                        if (isCtrl && isAlt && isX) return NativeMethods.CallNextHookEx(_kbdHookId, nCode, wParam, lParam);

                        _events.Add(new InputEvent
                        {
                            Type = (msg == WM_KEYDOWN) ? EventType.KeyDown : EventType.KeyUp,
                            VkCode = vk,
                            ScanCode = (int)info.scanCode,
                            DelayMs = 1000
                        });
                        UI(RefreshListAndJson);
                    }
                }
            }
            catch { /* ignore */ }

            return NativeMethods.CallNextHookEx(_kbdHookId, nCode, wParam, lParam);
        }

        // ===== Playback =======================================================
        private void StartPlayback()
        {
            // JSON 편집 반영
            TryLoadEventsFromJsonText();

            // 재생 중 캘리브 수집 금지
            _calibActive = false;

            // UI 상태
            BtnPlay.IsEnabled = false;
            BtnStopPlay.IsEnabled = true;
            BtnRecord.IsEnabled = false;
            BtnLoad.IsEnabled = false;
            BtnSave.IsEnabled = false;
            TxtStatus.Text = "Playing...";

            // 🔒 재생 중 편집 잠금 (열거 수정 방지)
            ListEvents.IsEnabled = false;
            TxtJson.IsReadOnly = true;

            _playbackCts = new CancellationTokenSource();
            var token = _playbackCts.Token;

            // ✔️ 이벤트 스냅샷 (복사본)
            var snapshot = _events.ToArray(); // <- 핵심

            // UI 스냅샷
            bool useAdb = false; int devW = 0, devH = 0;
            useAdb = (ChkUseAdb.IsChecked == true);
            if (useAdb) { int.TryParse(TxtDevW.Text, out devW); int.TryParse(TxtDevH.Text, out devH); }

            RegisterStopHotkey();

            _ = Task.Run(async () =>
            {
                try
                {
                    // ⬇️ 복사본만 전달
                    await PlaybackLoop(snapshot, token, useAdb, devW, devH);
                }
                catch (OperationCanceledException) { }
                catch (Exception ex) { UI(() => MessageBox.Show("Playback error: " + ex.Message)); }
                finally
                {
                    UI(() => UnregisterStopHotkey());
                    UI(StopPlaybackUiCleanup);
                    UI(() => TxtStatus.Text = "Playback stopped.");
                }
            });
        }


        private void StopPlayback()
        {
            _playbackCts?.Cancel();
            UI(() => UnregisterStopHotkey());
            UI(StopPlaybackUiCleanup);
            UI(() => TxtStatus.Text = "Playback stopped by user.");
        }

        private void StopPlaybackUiCleanup()
        {
            BtnPlay.IsEnabled = _events.Count > 0;
            BtnStopPlay.IsEnabled = false;
            BtnRecord.IsEnabled = true;
            BtnLoad.IsEnabled = true;
            BtnSave.IsEnabled = _events.Count > 0;

            // 🔓 재생 종료 후 편집 가능
            ListEvents.IsEnabled = true;
            TxtJson.IsReadOnly = false;
        }


        private async Task PlaybackLoop(IReadOnlyList<InputEvent> events, CancellationToken token, bool useAdb, int devW, int devH)
        {
            // 창 위치/크기
            int winLeft = 0, winTop = 0, clientW = 0, clientH = 0;
            if (_targetWindow != IntPtr.Zero)
            {
                if (NativeMethods.GetWindowRect(_targetWindow, out RECT wr))
                { winLeft = wr.Left; winTop = wr.Top; }
                if (NativeMethods.GetClientRect(_targetWindow, out RECT rc))
                { clientW = rc.Right - rc.Left; clientH = rc.Bottom - rc.Top; }
            }

            foreach (var ev in events)
            {
                token.ThrowIfCancellationRequested();
                await Task.Delay(ev.DelayMs, token);

                switch (ev.Type)
                {
                    case EventType.MouseClick:
                        if (useAdb)
                        {
                            // ADB: 캘리브가 있으면 컨텐츠 기준, 없으면 클라이언트 기준
                            int baseW = _calib.IsValid ? _calib.ContentW : clientW;
                            int baseH = _calib.IsValid ? _calib.ContentH : clientH;
                            if (baseW <= 0 || baseH <= 0) { UI(() => MessageBox.Show("Invalid window/calibre size for ADB mapping.")); return; }

                            int cx = ev.X, cy = ev.Y;
                            if (!_calib.IsValid)
                            {
                                // 클라이언트 좌표로 녹화되었다고 가정
                                // (화면절대좌표로 녹화된 경우엔 ADB가 맞지 않을 수 있음)
                            }

                            int dx = (int)Math.Round((double)cx / baseW * devW);
                            int dy = (int)Math.Round((double)cy / baseH * devH);
                            RunAdb($"shell input tap {dx} {dy}");
                        }
                        else
                        {
                            // PC: 컨텐츠 → 스크린
                            int sx, sy;
                            if (_calib.IsValid)
                            {
                                sx = winLeft + _calib.OffsetX + ev.X;
                                sy = winTop + _calib.OffsetY + ev.Y;
                            }
                            else
                            {
                                sx = winLeft + ev.X;
                                sy = winTop + ev.Y;
                            }
                            ClickAbsVirtual(sx, sy);
                        }
                        break;

                    case EventType.KeyDown:
                    case EventType.KeyUp:
                        if (useAdb)
                        {
                            if (AdbKeyEventForVk(ev.VkCode) is int androidKey)
                                RunAdb($"shell input keyevent {androidKey}");
                        }
                        else
                        {
                            bool down = ev.Type == EventType.KeyDown;
                            SendKey(down, ev.VkCode, ev.ScanCode);
                        }
                        break;
                }
            }
        }

        // ===== Low-level: Mouse/Keyboard send =================================
        private void ClickAbsVirtual(int screenX, int screenY)
        {
            const int SM_XVIRTUALSCREEN = 76, SM_YVIRTUALSCREEN = 77, SM_CXVIRTUALSCREEN = 78, SM_CYVIRTUALSCREEN = 79;
            int vx = NativeMethods.GetSystemMetrics(SM_XVIRTUALSCREEN);
            int vy = NativeMethods.GetSystemMetrics(SM_YVIRTUALSCREEN);
            int vw = NativeMethods.GetSystemMetrics(SM_CXVIRTUALSCREEN);
            int vh = NativeMethods.GetSystemMetrics(SM_CYVIRTUALSCREEN);

            int nx = (int)Math.Round((screenX - vx) * 65535.0 / Math.Max(1, vw - 1));
            int ny = (int)Math.Round((screenY - vy) * 65535.0 / Math.Max(1, vh - 1));

            var inputs = new INPUT[]
            {
                new INPUT{ type = INPUT_MOUSE, U = new InputUnion{ mi = new MOUSEINPUT{ dx = nx, dy = ny, dwFlags = MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_VIRTUALDESK } } },
                new INPUT{ type = INPUT_MOUSE, U = new InputUnion{ mi = new MOUSEINPUT{ dwFlags = MOUSEEVENTF_LEFTDOWN } } },
                new INPUT{ type = INPUT_MOUSE, U = new InputUnion{ mi = new MOUSEINPUT{ dwFlags = MOUSEEVENTF_LEFTUP } } },
            };
            NativeMethods.SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
        }

        private void SendKey(bool down, int vk, int scan = 0)
        {
            const uint INPUT_KEYBOARD = 1;
            const uint KEYEVENTF_KEYUP = 0x0002, KEYEVENTF_SCANCODE = 0x0008;

            uint flags = 0;
            if (!down) flags |= KEYEVENTF_KEYUP;
            if (scan > 0) flags |= KEYEVENTF_SCANCODE;

            var inp = new INPUT
            {
                type = INPUT_KEYBOARD,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = (ushort)vk,
                        wScan = (ushort)scan,
                        dwFlags = flags
                    }
                }
            };
            NativeMethods.SendInput(1, new[] { inp }, Marshal.SizeOf<INPUT>());
        }

        private int? AdbKeyEventForVk(int vk) => vk switch
        {
            0x0D => 66,   // Enter -> KEYCODE_ENTER
            0x1B => 111,  // Esc   -> KEYCODE_ESCAPE
            0x08 => 67,   // Backspace -> KEYCODE_DEL
            0x09 => 61,   // Tab -> KEYCODE_TAB
            _ => null
        };

        private void RunAdb(string args)
        {
            try
            {
                var p = new Process();
                p.StartInfo = new ProcessStartInfo("adb", args)
                {
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = true
                };
                p.Start(); _ = p.StandardOutput.ReadToEnd(); p.WaitForExit(3000);
            }
            catch (Exception ex) { UI(() => MessageBox.Show("adb execution error: " + ex.Message)); }
        }

        // ===== Hotkey (Ctrl+Alt+X) ===========================================
        private void RegisterStopHotkey()
        {
            if (_hotkeyRegistered) return;
            var helper = new System.Windows.Interop.WindowInteropHelper(this);
            var hWnd = helper.Handle;
            var src = System.Windows.Interop.HwndSource.FromHwnd(hWnd);
            src.AddHook(HwndHook);

            const uint MOD_CONTROL = 0x0002, MOD_ALT = 0x0001;
            uint vk = (uint)KeyInterop.VirtualKeyFromKey(Key.X);

            if (!NativeMethods.RegisterHotKey(hWnd, HOTKEY_ID, MOD_CONTROL | MOD_ALT, vk))
                UI(() => MessageBox.Show("Hotkey registration failed (Ctrl+Alt+X)"));
            else _hotkeyRegistered = true;
        }

        private void UnregisterStopHotkey()
        {
            if (!_hotkeyRegistered) return;
            var helper = new System.Windows.Interop.WindowInteropHelper(this);
            NativeMethods.UnregisterHotKey(helper.Handle, HOTKEY_ID);
            _hotkeyRegistered = false;
        }

        private IntPtr HwndHook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            const int WM_HOTKEY = 0x0312;
            if (msg == WM_HOTKEY && wParam.ToInt32() == HOTKEY_ID) { StopPlayback(); handled = true; }
            return IntPtr.Zero;
        }

        // ===== Window search by partial title =================================
        private static IntPtr FindWindowByTitleContains(string text)
        {
            IntPtr found = IntPtr.Zero;
            NativeMethods.EnumWindows((hWnd, lParam) =>
            {
                if (!NativeMethods.IsWindowVisible(hWnd)) return true;
                StringBuilder sb = new(512);
                NativeMethods.GetWindowText(hWnd, sb, sb.Capacity);
                if (sb.ToString().Contains(text, StringComparison.OrdinalIgnoreCase)) { found = hWnd; return false; }
                return true;
            }, IntPtr.Zero);
            return found;
        }

        // ===== Native interop =================================================
        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);
        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)] private struct POINT { public int x, y; }
        [StructLayout(LayoutKind.Sequential)] private struct MSLLHOOKSTRUCT { public POINT pt; public uint mouseData, flags, time, dwExtraInfo; }
        [StructLayout(LayoutKind.Sequential)] private struct KBDLLHOOKSTRUCT { public uint vkCode, scanCode, flags, time, dwExtraInfo; }
        [StructLayout(LayoutKind.Sequential)] private struct RECT { public int Left, Top, Right, Bottom; }

        [StructLayout(LayoutKind.Sequential)] private struct INPUT { public uint type; public InputUnion U; }
        [StructLayout(LayoutKind.Explicit)]
        private struct InputUnion
        {
            [FieldOffset(0)] public MOUSEINPUT mi;
            [FieldOffset(0)] public KEYBDINPUT ki;
        }
        [StructLayout(LayoutKind.Sequential)] private struct MOUSEINPUT { public int dx, dy; public uint mouseData, dwFlags, time; public IntPtr dwExtraInfo; }
        [StructLayout(LayoutKind.Sequential)] private struct KEYBDINPUT { public ushort wVk, wScan; public uint dwFlags, time; public IntPtr dwExtraInfo; }

        private const uint INPUT_MOUSE = 0;
        private const uint MOUSEEVENTF_MOVE = 0x0001, MOUSEEVENTF_ABSOLUTE = 0x8000, MOUSEEVENTF_LEFTDOWN = 0x0002, MOUSEEVENTF_LEFTUP = 0x0004, MOUSEEVENTF_VIRTUALDESK = 0x4000;

        private static class NativeMethods
        {
            public const int WH_MOUSE_LL = 14, WH_KEYBOARD_LL = 13;

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            public static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            public static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            public static extern bool UnhookWindowsHookEx(IntPtr hhk);

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

            [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            public static extern IntPtr GetModuleHandle(string? lpModuleName);

            public static IntPtr SetMouseHook(LowLevelMouseProc proc) => SetWindowsHookEx(WH_MOUSE_LL, proc, GetModuleHandle(null), 0);
            public static IntPtr SetKeyboardHook(LowLevelKeyboardProc proc) => SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(null), 0);

            [DllImport("user32.dll")] public static extern bool ScreenToClient(IntPtr hWnd, ref POINT lpPoint);
            [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
            [DllImport("user32.dll")] public static extern bool GetClientRect(IntPtr hWnd, out RECT rect);
            [DllImport("user32.dll")] public static extern int GetSystemMetrics(int nIndex);

            [DllImport("user32.dll", SetLastError = true)]
            public static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

            // Hotkey
            [DllImport("user32.dll")] public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
            [DllImport("user32.dll")] public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

            // Enum windows
            public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
            [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
            [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
            [DllImport("user32.dll", CharSet = CharSet.Auto)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        }
    }
}
