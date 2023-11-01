using System;
using System.Windows;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Windows.Interop;
using System.Windows.Media.Media3D;
using System.Diagnostics;
using System.Windows.Automation;
using System.Timers;

namespace WPF_Language_Lab_project
{
    public partial class MainWindow : Window
    {
        // External functions for hotkey registration
        [DllImport("user32.dll")]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vlc);
        [DllImport("user32.dll")]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        [DllImport("user32.dll")]
        private static extern bool UnhookWinEvent(IntPtr hWinEventHook); // NOT UNDERSTOOD!
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);
        [DllImport("user32.dll")]
        private static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
        // End of Dll imports
        
        private delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime); // NOT UNDERSTOOD!

        private System.Timers.Timer timer;
        private IntPtr targetWindowHandle;
        private AutomationElement _targetWindowElement;
        private AutomationElement _learningLabElement;
        private IntPtr learningLabhandle = IntPtr.Zero;

        private bool hotkeyActivated = false;
        private bool windowEventIsSet = false;

        private const int HOTKEY_ID = 9000; // arbitrary unique ID
        private const int MOD_CONTROL = 0x0002;
        private const int MOD_SHIFT = 0x0004;
        private const int KEY_O = 0x4F;

        private NotifyIcon notifyIcon = null;

        // Variable that adds a functionality of not closing the application when Clicking on the 'X' (instead it gets minimized to tray)
        private bool forceClose = false;

        private LearningLabActivity learningLabActivity = null;
        private WindowVisualState _lastWindowState;
        private AutomationElement _windowElement;

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        public struct RECT
        {
            public int Left; // Specifies the x-coordinate of the upper-left corner of the rectangle.
            public int Top; // Specifies the y-coordinate of the upper-left corner of the rectangle.
            public int Right; // Specifies the x-coordinate of the lower-right corner of the rectangle.
            public int Bottom; // Specifies the y-coordinate of the lower-right corner of the rectangle.
        }

        private int originalForegroundWidth = 0;
        private int originalForegroundHeight = 0;

        // Call this function when your hotkey is pressed, right before showing the LearningLabActivity window.

        public MainWindow()
        {
            System.Diagnostics.Debug.WriteLine("Entering MainWindow Constructor");
            InitializeComponent();
            InitializeSystemTray();
            Loaded += MainWindow_Loaded;
            Closed += MainWindow_Closed;
        }

        private void HandleEventSubscriptions()
        {
            if (windowEventIsSet)
            {
                Automation.RemoveAutomationPropertyChangedEventHandler(_targetWindowElement, OnWindowPropertyChanged);
                MonitorWindowState(_targetWindowElement, true);
                MonitorWindowState(_learningLabElement, true);

                System.Diagnostics.Debug.WriteLine("Removing TheEventHandler");
                windowEventIsSet = false;
            }
            if (targetWindowHandle != null && !windowEventIsSet)
            {
                _targetWindowElement = AutomationElement.FromHandle(targetWindowHandle);
                _learningLabElement = AutomationElement.FromHandle(learningLabhandle);
                // OnWindowPropertyChanged Handles the repositioning of the attached window (learningLab) to a target window
                Automation.AddAutomationPropertyChangedEventHandler(
                    _targetWindowElement,
                    TreeScope.Element,
                    OnWindowPropertyChanged,
                    AutomationElement.BoundingRectangleProperty);
                MonitorWindowState(_targetWindowElement, false);
                MonitorWindowState(_learningLabElement, false);

                System.Diagnostics.Debug.WriteLine("Adding TheEventHandler");
                windowEventIsSet = true;
            }
        }

        private bool ShouldAdjustWindow()
        {
            IntPtr mainWindowhandle = new WindowInteropHelper(this).Handle;
            return targetWindowHandle != IntPtr.Zero &&
                   targetWindowHandle != mainWindowhandle &&
                   targetWindowHandle != learningLabhandle &&
                   !hotkeyActivated;
        }

        private void AdjustTargetWindowAndShowLearningLab()
        {
            RECT activeWindowRect;
            GetWindowRect(targetWindowHandle, out activeWindowRect);

            originalForegroundWidth = activeWindowRect.Right - activeWindowRect.Left;
            originalForegroundHeight = activeWindowRect.Bottom - activeWindowRect.Top;

            int width = activeWindowRect.Right - activeWindowRect.Left;
            int height = activeWindowRect.Bottom - activeWindowRect.Top;

            MoveWindow(targetWindowHandle, activeWindowRect.Left, activeWindowRect.Top, width - 200, height, true);

            ShowLearningLabPopup();

            if (learningLabActivity != null)
            {
                learningLabActivity.Width = 200;
                learningLabActivity.Height = height;
                learningLabActivity.Left = activeWindowRect.Right - 200;
                learningLabActivity.Top = activeWindowRect.Top;
                learningLabhandle = new WindowInteropHelper(learningLabActivity).Handle;
            }

            hotkeyActivated = true;
        }
        private void AdjustAndPositionWindow()
        {
            System.Diagnostics.Debug.WriteLine("Entering AdjustAndPositionWindow");
            IntPtr activeWindowHandle = GetForegroundWindow();
            targetWindowHandle = activeWindowHandle;

            if (ShouldAdjustWindow())
            {
                AdjustTargetWindowAndShowLearningLab();
            }
            else if (activeWindowHandle == IntPtr.Zero)
            {
                ShowLearningLabPopup();
            }
            else if (activeWindowHandle == learningLabhandle)
            {
                CloseLearningLabActivity(false);
            }

            HandleEventSubscriptions();

        }

        //      ...PropertyChanged Event Functions :
        private void OnWindowPropertyChanged(object sender, AutomationPropertyChangedEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine(e);
            if (e.Property == AutomationElement.BoundingRectangleProperty)
            {
                RECT activeWindowRect;
                RECT learningLabRect;
                GetWindowRect(targetWindowHandle, out activeWindowRect);
                var targetWindowrect = (Rect)e.NewValue;
                System.Diagnostics.Debug.WriteLine("Rect from GetWindowRect : " + WritingRect(activeWindowRect) + "\nRect from EventArgs : " + targetWindowrect);
                // Update the position of the other window according to the target window's position
                //_learningLabElement.Left = targetWindowrect.Left + targetWindowrect.Width;
                //_learningLabElement.Top = targetWindowrect.Top;
                GetWindowRect(learningLabhandle, out learningLabRect);
                System.Diagnostics.Debug.WriteLine((learningLabRect.Right - learningLabRect.Left).ToString());
                MoveWindow(learningLabhandle, activeWindowRect.Right, activeWindowRect.Top, learningLabRect.Right - learningLabRect.Left, activeWindowRect.Bottom - activeWindowRect.Top, true);
            }
        }
        void OnTargetWindowPropertyChanged(object sender, AutomationPropertyChangedEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("e.OldValue : " + e.OldValue);
            System.Diagnostics.Debug.WriteLine("e.NewValue : " + e.NewValue);
            if ((WindowVisualState)e.NewValue == WindowVisualState.Maximized)
            {
                // Adjust the size and position of the learningLab window
                // to fill the entire screen except for 200 pixels on the right.
                AdjustLearningLabAndTargetWindow(true);
            }
        }

        void OnLearningLabPropertyChanged(object sender, AutomationPropertyChangedEventArgs e)
        {
            if ((WindowVisualState)e.NewValue == WindowVisualState.Maximized)
            {
                // Adjust the size and position of the target window
                // to fill the entire screen except for 200 pixels on the left.
                AdjustLearningLabAndTargetWindow(false);
            }
        }

        void AdjustLearningLabAndTargetWindow(bool isTargetWindow)
        {
            // Get the screen size
            var screenSize = System.Windows.Forms.Screen.PrimaryScreen.Bounds;
            if (!isTargetWindow) // if it is the learning lab event handler calling this method
            {
                // Set the learningLab window size and position
                SetWindowPos(learningLabhandle, IntPtr.Zero, 200, 0, screenSize.Width - 200, screenSize.Height, 0x0040);
                SetWindowPos(targetWindowHandle, IntPtr.Zero, 0, 0, 200, screenSize.Height, 0x0040);
            }
            else if (isTargetWindow)
            {
                SetWindowPos(targetWindowHandle, IntPtr.Zero, 0, 0, screenSize.Width - 200, screenSize.Height, 0x0040);
                SetWindowPos(learningLabhandle, IntPtr.Zero, screenSize.Width - 200, 0, 200, screenSize.Height, 0x0040);
            }            

        }

        public void MonitorWindowState(AutomationElement windowElement, bool isSubscribed)
        {
            _lastWindowState = GetWindowVisualState(windowElement); // Problematic line ( can throw 'ElementNotAvailableException' exception )
            _windowElement = windowElement;
            // Start a timer to poll the window state at regular intervals
            timer = new System.Timers.Timer(100);  // Interval in milliseconds
            if (isSubscribed)
            {
                // Unsubscribing from the event
                System.Diagnostics.Debug.WriteLine("Unsubscribing from the event " + windowElement);
                System.Diagnostics.Debug.WriteLine("Current Subscribers List " + WindowElementElapsedEventHandler.GetInvocationList().ToString() + windowElement);

                //timer.Elapsed -= (sender, e) => CheckWindowState(windowElement);
                timer.Stop();
                timer.Elapsed -= WindowElementElapsedEventHandler;
            }
            else if (!isSubscribed)
            {
                // Subscribing to the event
                System.Diagnostics.Debug.WriteLine("Subscribing to the event");
                //timer.Elapsed += (sender, e) => CheckWindowState(windowElement);
                timer.Elapsed += WindowElementElapsedEventHandler;
                timer.Start();
            }
            
        }

        private void WindowElementElapsedEventHandler(object sender, ElapsedEventArgs e)
        {
            CheckWindowState(_windowElement);
        }

        private void CheckWindowState(AutomationElement windowElement)
        {
            var currentWindowState = GetWindowVisualState(windowElement); // Problematic line ( can throw 'ElementNotAvailableException' exception )

            if (currentWindowState != _lastWindowState)
            {
                // The window state has changed
                System.Diagnostics.Debug.WriteLine("Window state changed from" + _lastWindowState + " to " + currentWindowState);
                _lastWindowState = currentWindowState;

                // Handle the window state change as needed
                if (windowElement == _targetWindowElement)
                {
                    AdjustLearningLabAndTargetWindow(true);
                }
                else if (windowElement == _learningLabElement)
                {
                    AdjustLearningLabAndTargetWindow(true);
                }
            }
        }

        private WindowVisualState GetWindowVisualState(AutomationElement windowElement)
        {
            try
            {
                if (windowElement.TryGetCurrentPattern(WindowPattern.Pattern, out object pattern))
                {
                    var windowPattern = (WindowPattern)pattern;
                    return windowPattern.Current.WindowVisualState; // Exception thrown here!! "System.Windows.Automation.ElementNotAvailableException: 'ElementNotAvailable'"
                }
            }            
            catch (ElementNotAvailableException)
            {
                System.Diagnostics.Debug.WriteLine("Exception 'ElementNotAvailableException' was raised");
            }
            return WindowVisualState.Normal;  // Default value
        }

        private string WritingRect(RECT windowRect) {
            return windowRect.Top.ToString() + ";" + windowRect.Bottom.ToString() + ";" + windowRect.Right.ToString() + ";" + windowRect.Left.ToString();
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("Entering OnClosing");
            if (!forceClose)
            {
                System.Diagnostics.Debug.WriteLine("Entering OnClosing");
                e.Cancel = true; // This cancels the close operation
                Hide(); // This hides the window so it's not visible anymore
            }
            notifyIcon.Dispose();
            base.OnClosing(e);
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("Entering MainWindow_Loaded");

            IntPtr handle = new WindowInteropHelper(this).Handle;
            RegisterHotKey(handle, HOTKEY_ID, MOD_CONTROL | MOD_SHIFT, KEY_O);
            SetRunAtStartup(true);
            Hide();
        }

        private void MainWindow_Closed(object sender, EventArgs e)
        {
            IntPtr handle = new WindowInteropHelper(this).Handle;
            UnregisterHotKey(handle, HOTKEY_ID);
            if (_targetWindowElement != null)
            {
                Automation.RemoveAutomationPropertyChangedEventHandler(
                    _targetWindowElement, OnWindowPropertyChanged);
            }
            notifyIcon.Dispose();

        }


        private void InitializeSystemTray()
        {
            notifyIcon = new NotifyIcon
            {
                Icon = System.Drawing.SystemIcons.Application
            };

            ContextMenuStrip contextMenuStrip = new ContextMenuStrip();
            ToolStripMenuItem openItem = new ToolStripMenuItem("Open");
            openItem.Click += (sender, args) => ShowMainWindow();
            contextMenuStrip.Items.Add(openItem);

            ToolStripMenuItem exitItem = new ToolStripMenuItem("Exit");
            exitItem.Click += (sender, args) => CloseApp();
            contextMenuStrip.Items.Add(exitItem);

            notifyIcon.ContextMenuStrip = contextMenuStrip;
            notifyIcon.Visible = true;

            notifyIcon.DoubleClick += (sender, args) => ShowMainWindow();
        }


        protected override void OnSourceInitialized(EventArgs e) // Automatically executes when the underlying Win32 window handle becomes available.
        {
            base.OnSourceInitialized(e);
            HwndSource source = PresentationSource.FromVisual(this) as HwndSource;
            source.AddHook(WndProc);
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            System.Diagnostics.Debug.WriteLine("Entering WndProc");
            if (msg == 0x0312 && wParam.ToInt32() == HOTKEY_ID) // if message is WM_HOTKEY and it's ID is 9000 (the specified Hotkey id for our Ctrl+Shift+O hotkey)
            {
                System.Diagnostics.Debug.WriteLine("Inside if-condition for WM_HOTKEY and HOTKEY_ID");
                AdjustAndPositionWindow();
                //if (learningLabActivity != null && learningLabActivity.IsLoaded)
                //{
                //    CloseLearningLabActivity(false); // Close the Learning Lab Activity
                //} 
                //else
                //{
                //    AdjustAndPositionWindow(); // Show the Learning Lab Activity
                //}
                handled = true;
            }
            return IntPtr.Zero;
        }

        public void LearningLabPopup_Closed(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("LearningLabPopup is being Closed");
            MonitorWindowState(_targetWindowElement, true);
            MonitorWindowState(_learningLabElement, true);
        }

        private void CloseLearningLabActivity(bool invokedFromEvent)
        {
            System.Diagnostics.Debug.WriteLine("Entering CloseLearningLabActivity");
            IntPtr activeWindowHandle = GetForegroundWindow();

            if (activeWindowHandle != IntPtr.Zero && originalForegroundWidth > 0 && originalForegroundHeight > 0) // If there is a foreground window
            {
                System.Diagnostics.Debug.WriteLine("Inside if-condition for valid activeWindowHandle and dimensions");
                RECT activeWindowRect;
                GetWindowRect(activeWindowHandle, out activeWindowRect);

                // Restore original size
                MoveWindow(activeWindowHandle, activeWindowRect.Left, activeWindowRect.Top, originalForegroundWidth, originalForegroundHeight, true);
                hotkeyActivated = false;
            }

            if (learningLabActivity != null) // if learning lab window is already open
            {
                System.Diagnostics.Debug.WriteLine("Inside if-condition for existing learningLabActivity");
                learningLabActivity.ClosedByUser -= CloseLearningLabActivity;
                hotkeyActivated = false;
                if (invokedFromEvent == false) { learningLabActivity.Close(); } // This is done because normal if this function was invoked from an event, it is already closing, 
                learningLabActivity = null;
            }
        }


        private void ShowMainWindow()
        {
            Show();
            WindowState = WindowState.Normal; // Sets it's size to normal, not minimized or maximized
            Activate(); // Puts the window on the foreground/ focused
        }

        private void ShowLearningLabPopup()
        {
            // Check if the LearningLabActivity window is null or closed
            if (learningLabActivity == null || !learningLabActivity.IsLoaded)
            {
                learningLabActivity = new LearningLabActivity();
                learningLabActivity.ClosedByUser += CloseLearningLabActivity; // Subscribe to event
                learningLabActivity.Show();
            }
            else
            {
                // If it's already open, just bring it to the foreground
                learningLabActivity.Activate();
            }
        }

        // Close App from System Tray
        private void CloseApp()
        {
            System.Diagnostics.Debug.WriteLine("Entering CloseApp");
            forceClose = true; // Set the flag to true so that the close isn't cancelled
            notifyIcon.Dispose();

            // Close all open windows
            foreach (Window window in System.Windows.Application.Current.Windows)
            {
                window.Close();
            }
        
        }


        // Run on startup in background
        private void SetRunAtStartup(bool shouldRun)
        {
            System.Diagnostics.Debug.WriteLine("Entering SetRunAtStartup");
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true))
            {
                if (shouldRun)
                {
                    System.Diagnostics.Debug.WriteLine("Inside if-condition for enabling run at startup");
                    key.SetValue("WPF_Language_Lab_project", System.Reflection.Assembly.GetExecutingAssembly().Location);
                }
                else {
                    System.Diagnostics.Debug.WriteLine("Inside else-condition for disabling run at startup");
                    key.DeleteValue("WPF_Language_Lab_project", false);
                }
            }
        }

    }
}
