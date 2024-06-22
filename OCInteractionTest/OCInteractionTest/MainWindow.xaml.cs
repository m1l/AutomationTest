using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Automation;

namespace OCInteractionTest
{    
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        //Gets OC path and updates PathTextBox
        private void GetButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var ocAutomationElement = GetOCAutomationElement();
                if (ocAutomationElement != null)
                {
                    AutomationElement automationElement = GetTextElement(ocAutomationElement, "CurrentPathGet");
                    if (automationElement != null)
                    {
                        ValuePattern valuePattern = automationElement.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                        PathTextBox.Text = valuePattern.Current.Value;
                    }                    
                }
                else
                    MessageBox.Show("Can't find OneCommander window");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //Sets OC path to value from PathTextBox
        //This can also be a path to a file, it will open folder and select it
        private void SetButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var ocAutomationElement = GetOCAutomationElement();
                if (ocAutomationElement != null)
                {
                    AutomationElement automationElement = GetTextElement(ocAutomationElement, "CurrentPathSet");
                    if (automationElement != null)
                    {
                        ValuePattern valuePattern = automationElement.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                        valuePattern.SetValue(PathTextBox.Text);
                    }
                }
                else
                    MessageBox.Show("Can't find OneCommander window");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
               

        private AutomationElement GetOCAutomationElement()
        {
            var ocWindowHandle = GetOCWindowHandle();
            if (ocWindowHandle == IntPtr.Zero)
                return null;
            var ae = AutomationElement.FromHandle(ocWindowHandle);
            return ae;            
        }
        

        private IntPtr GetOCWindowHandle()
        {
            //Check first if the focused window is OneCommander
            IntPtr foregroundWinHandle = GetForegroundWindow();
            if (IsHandleOCWindow(foregroundWinHandle))
                return foregroundWinHandle;

            //no? Let's find OC window
            IntPtr oCWindowHandle = FindWindowHandleByName("OneCommander");
            return oCWindowHandle; //or null
        }

        internal static IntPtr FindWindowHandleByName(string windowName)
        {
            // Get the desktop element, which is the root of the UI Automation tree
            AutomationElement desktop = AutomationElement.RootElement;

            // Create a condition to find the window by its Name property
            var condition = new PropertyCondition(AutomationElement.NameProperty, windowName);

            // Find the first window that matches the condition
            AutomationElement windowElement = desktop.FindFirst(TreeScope.Children, condition);

            if (windowElement != null)
            {
                // Retrieve the window handle from the AutomationElement
                IntPtr windowHandle = new IntPtr(windowElement.Current.NativeWindowHandle);
                return windowHandle;
            }

            return IntPtr.Zero; // Return zero if the window is not found
        }
        private AutomationElement GetTextElement(AutomationElement parentElement, string value)
        {
            System.Windows.Automation.Condition condition = new PropertyCondition(AutomationElement.AutomationIdProperty, value);
            return parentElement.FindFirst(TreeScope.Children | TreeScope.Descendants, condition);
        }

        internal static bool IsHandleOCWindow(IntPtr hWnd)
        {            
            var className = new StringBuilder(512);
            var r = GetClassName(hWnd, className, className.Capacity);
            if (r != 0 && className.ToString().Contains("OneCommander.exe"))
                return true;
            return false;
        }

        [DllImport("user32.dll")]
        internal static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);       
    }
}
