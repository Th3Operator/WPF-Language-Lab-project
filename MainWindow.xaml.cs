using System;
using System.Windows;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Windows.Interop;
using System.Windows.Media.Media3D;
using System.Diagnostics;
using System.Windows.Automation;


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

            if (_targetWindowElement != null)
            {
                Automation.AddAutomationPropertyChangedEventHandler(
                    _targetWindowElement,
                    TreeScope.Element,
                    OnWindowPropertyChanged,
                    AutomationElement.BoundingRectangleProperty);
            }

        }

        private void HandleEventSubscriptions()
        {
            if (windowEventIsSet)
            {
                Automation.RemoveAutomationPropertyChangedEventHandler(_targetWindowElement, OnWindowPropertyChanged);
                System.Diagnostics.Debug.WriteLine("Removing TheEventHandler");
                windowEventIsSet = false;
            }
            if (targetWindowHandle != null && !windowEventIsSet)
            {
                _targetWindowElement = AutomationElement.FromHandle(targetWindowHandle);
                Automation.AddAutomationPropertyChangedEventHandler(
                    _targetWindowElement,
                    TreeScope.Element,
                    OnWindowPropertyChanged,
                    AutomationElement.BoundingRectangleProperty);
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

            HandleEventSubscriptions();

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
        }
        /*private void AdjustAndPositionWindow()
        {
            System.Diagnostics.Debug.WriteLine("Entering AdjustAndPositionWindow");
            IntPtr activeWindowHandle = GetForegroundWindow();
            targetWindowHandle = activeWindowHandle;
            //learningLabhandle = IntPtr.Zero;
            if (windowEventIsSet) { 
                Automation.RemoveAutomationPropertyChangedEventHandler(_targetWindowElement, OnWindowPropertyChanged);
                System.Diagnostics.Debug.WriteLine("Removing TheEventHandler");
                windowEventIsSet = false;
            }
            if (activeWindowHandle != null && !windowEventIsSet)
            {
                _targetWindowElement = AutomationElement.FromHandle(activeWindowHandle);
                Automation.AddAutomationPropertyChangedEventHandler(
                    _targetWindowElement,
                    TreeScope.Element,
                    OnWindowPropertyChanged,
                    AutomationElement.BoundingRectangleProperty);
                System.Diagnostics.Debug.WriteLine("Adding TheEventHandler");
                windowEventIsSet = true;

            }
            IntPtr mainWindowhandle = new WindowInteropHelper(this).Handle;
            
            if (learningLabActivity != null)
            {
                learningLabhandle = new WindowInteropHelper(learningLabActivity).Handle;
                _learningLabElement = AutomationElement.FromHandle(learningLabhandle);
            }

            // Only adjust if there is a foreground window
            if (activeWindowHandle != IntPtr.Zero && activeWindowHandle != mainWindowhandle && activeWindowHandle != learningLabhandle && !hotkeyActivated) // This might be the problem for the exception being thrown
            {
                System.Diagnostics.Debug.WriteLine("Inside if-condition for valid activeWindowHandle");
                RECT activeWindowRect;
                GetWindowRect(activeWindowHandle, out activeWindowRect);

                originalForegroundWidth = activeWindowRect.Right - activeWindowRect.Left; // Save original width
                originalForegroundHeight = activeWindowRect.Bottom - activeWindowRect.Top; // Save original height

                int width = activeWindowRect.Right - activeWindowRect.Left;
                int height = activeWindowRect.Bottom - activeWindowRect.Top;

                // Adjust the width of the active window by -200px
                MoveWindow(activeWindowHandle, activeWindowRect.Left, activeWindowRect.Top, width - 200, height, true); // change back to true

                // Show Learning Lab Popup
                ShowLearningLabPopup();

                // Position Learning Lab Popup window
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
            //else if (activeWindowHandle != IntPtr.Zero && activeWindowHandle != mainWindowhandle && activeWindowHandle != learningLabhandle && hotkeyActivated) // New Function here
            //{
            //    System.Diagnostics.Debug.WriteLine("Inside if-condition for SECOND valid activeWindowHandle");
            //}
            else if (activeWindowHandle == IntPtr.Zero) // If there is no foreground window
            {
                System.Diagnostics.Debug.WriteLine("Inside if-condition for no foreground window");
                // If there is no foreground window, show the Learning Lab Popup conventionally
                ShowLearningLabPopup();
            }
            else if (activeWindowHandle == learningLabhandle) // If the learning lab window is already open, and it is the foreground window
            {
                System.Diagnostics.Debug.WriteLine("Inside if-condition for learningLabhandle foreground window");
                CloseLearningLabActivity(false);
            }
        }*/

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

        private string WritingRect(RECT windowRect) {
            return windowRect.Top.ToString() + ";" + windowRect.Bottom.ToString() + ";" + windowRect.Right.ToString() + ";" + windowRect.Left.ToString();
        }
        private void RepositionAppNextToTarget()
        {
            // Get the target window's position and size
            GetWindowRect(targetWindowHandle, out RECT rect);

            // Calculate your window's position and size
            int appX = rect.Right;  
            int appY = rect.Top;
            int appWidth = 200;
            int appHeight = rect.Bottom - rect.Top;

            // Get your window's handle
            IntPtr appHandle = new WindowInteropHelper(this).Handle;

            // Set your window's position and size
            MoveWindow(appHandle, appX, appY, appWidth, appHeight, true);
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
            // Find the target window (e.g., the currently active window)
            targetWindowHandle = GetForegroundWindow();

            if (targetWindowHandle != IntPtr.Zero)
            {
                System.Diagnostics.Debug.WriteLine("Inside if-condition for valid targetWindowHandle");

                // Reposition your window initially
                RepositionAppNextToTarget();
            }
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


        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            HwndSource source = PresentationSource.FromVisual(this) as HwndSource;
            source.AddHook(WndProc);
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            System.Diagnostics.Debug.WriteLine("Entering WndProc");
            if (msg == 0x0312 && wParam.ToInt32() == HOTKEY_ID) // if message is WM_HOTKEY and it's ID is 9000, which is the specified Hotkey id for our Ctrl+Shift+O hotkey
            {
                System.Diagnostics.Debug.WriteLine("Inside if-condition for WM_HOTKEY and HOTKEY_ID");
                if (learningLabActivity != null && learningLabActivity.IsLoaded)
                {
                    CloseLearningLabActivity(false); // Close the Learning Lab Activity
                }
                else
                {
                    AdjustAndPositionWindow(); // Show the Learning Lab Activity
                }
                handled = true;
            }
            return IntPtr.Zero;
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
