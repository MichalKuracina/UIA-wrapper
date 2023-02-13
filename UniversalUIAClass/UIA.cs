using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using UIAutomationClient;

namespace UniversalUIAClass
{
    public class UIA
    {
        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private readonly CUIAutomation _automation;
        private readonly IUIAutomationElement _root;
        private readonly int _sleep;
        private Dictionary<int, string> _properties;

        public UIA()
        {
            _automation = new CUIAutomation();
            _root = _automation.GetRootElement();
            _properties = new Dictionary<int, string>();
            _sleep = 0;
        }

        public UIA(int sleep)
        {
            _automation = new CUIAutomation();
            _root = _automation.GetRootElement();
            _properties = new Dictionary<int, string>();
            _sleep = sleep;
        }

        public UIA AddProperty(Property propetyId, string value)
        {
            _properties.Add((int)propetyId, value);
            return this;
        }

        public void RemoveProperties()
        {
            _properties.Clear();
        }

        // https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-automation-element-propids
        public enum Property : int
        {
            Name = 30005,
            ClassName = 30012,
            LocalizedControlType = 30004,
            AutomationID = 30011,
            ControlType = 30003,
            AriaRole = 30101,
            FrameworkId = 30024
        }

        // https://learn.microsoft.com/en-us/previous-versions/dd757483(v=vs.85)
        public enum Pattern : int
        {
            Invoke = 10000,
            Value = 10002,
            Selection = 10001,
            ScrollItem = 10017,
            LegacyIAccessible = 10018,
            Window = 10009,
            SelectionItem = 10010,
            VirtualizedItem = 10020,
            TransformPattern = 10016,
            ExpandCollapse = 10005
        }

        private IUIAutomationElement FindFirst()
        {
            int propertiesCount = _properties.Count;

            IUIAutomationCondition[] conditionArray = new IUIAutomationCondition[propertiesCount];
            for (int i = 0; i < propertiesCount; i++)
            {
                conditionArray[i] = _automation.CreatePropertyCondition(_properties.ElementAt(i).Key, _properties.ElementAt(i).Value);
            }
            IUIAutomationCondition elementCondition = _automation.CreateAndConditionFromArray(conditionArray);
            IUIAutomationElement element = _root.FindFirst(TreeScope.TreeScope_Descendants, elementCondition);

            RemoveProperties(); ;

            if (element != null)
            {
                return element;
            }
            else
            {
                throw new NullReferenceException($"No such element found.");
            }

        }

        public bool Exists()
        {
            bool exists = false;

            int propertiesCount = _properties.Count;

            IUIAutomationCondition[] conditionArray = new IUIAutomationCondition[propertiesCount];
            for (int i = 0; i < propertiesCount; i++)
            {
                conditionArray[i] = _automation.CreatePropertyCondition(_properties.ElementAt(i).Key, _properties.ElementAt(i).Value);
            }
            IUIAutomationCondition elementCondition = _automation.CreateAndConditionFromArray(conditionArray);
            IUIAutomationElement element = _root.FindFirst(TreeScope.TreeScope_Descendants, elementCondition);

            RemoveProperties();

            if (element != null)
            {
                exists = true;
            }

            return exists;
        }

        public void Until(int seconds)
        {
            int counter = 0;
            Dictionary<int, string> tempDict = new Dictionary<int, string>(_properties);

            do
            {
                if (Exists())
                {
                    return;
                }
                _properties = new Dictionary<int, string>(tempDict);
                Thread.Sleep(1000);
                counter++;
            }
            while (counter < seconds);

            throw new TimeoutException($"Timeout threshold {seconds} reached.");
        }

        public void Invoke()
        {
            IUIAutomationElement element = FindFirst();
            IUIAutomationInvokePattern invokePattern = element.GetCurrentPattern((int)Pattern.Invoke);
            invokePattern.Invoke();
            Thread.Sleep(_sleep);
        }

        public void Expand()
        {
            IUIAutomationElement element = FindFirst();
            IUIAutomationExpandCollapsePattern expandCollapsePattern = element.GetCurrentPattern((int)Pattern.ExpandCollapse);
            expandCollapsePattern.Expand();
            Thread.Sleep(_sleep);
        }

        public void Collapse()
        {
            IUIAutomationElement element = FindFirst();
            IUIAutomationExpandCollapsePattern expandCollapsePattern = element.GetCurrentPattern((int)Pattern.ExpandCollapse);
            expandCollapsePattern.Collapse();
            Thread.Sleep(_sleep);
        }

        public void SelectionItem()
        {
            IUIAutomationElement element = FindFirst();
            IUIAutomationSelectionItemPattern selectionItem = element.GetCurrentPattern((int)Pattern.SelectionItem);
            selectionItem.Select();
            Thread.Sleep(_sleep);
        }

        public void VizualizedItem()
        {
            IUIAutomationElement element = FindFirst();
            IUIAutomationVirtualizedItemPattern virtualizedItem = element.GetCurrentPattern((int)Pattern.VirtualizedItem);
            virtualizedItem.Realize();
            Thread.Sleep(_sleep);
        }

        public Tuple<bool, bool, bool> Transform(double moveX = double.NaN, double moveY = double.NaN, double resizeWidth = double.NaN, double resizeHeight = double.NaN, double rotate = double.NaN)
        {
            IUIAutomationElement element = FindFirst();
            IUIAutomationTransformPattern transform = element.GetCurrentPattern((int)Pattern.TransformPattern);
            if (!Double.IsNaN(0 / moveX) && !Double.IsNaN(0 / moveY))
            {
                transform.Move(moveX, moveY);
            }
            else if (!Double.IsNaN(0 / resizeWidth) && !Double.IsNaN(0 / resizeHeight))
            {
                transform.Resize(resizeWidth, resizeHeight);
            }
            else if (!Double.IsNaN(0 / rotate))
            {
                transform.Rotate(rotate);
            }

            Thread.Sleep(_sleep);
            return Tuple.Create(Convert.ToBoolean(transform.CurrentCanMove), Convert.ToBoolean(transform.CurrentCanResize), Convert.ToBoolean(transform.CurrentCanRotate));

        }

        public void Write(string text)
        {
            IUIAutomationElement element = FindFirst();
            IUIAutomationValuePattern valuePattern = element.GetCurrentPattern((int)Pattern.Value);
            valuePattern.SetValue(text);
            Thread.Sleep(_sleep);
        }

        /// <summary>
        /// <b>UIAutomationClient.dll</b> overload. You must add element Properties to the object before using this method.<br/>State:<i> min, max, normal</i>
        /// </summary>
        /// <param name="state"></param>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        public void Window(string state)
        {
            IUIAutomationElement element = FindFirst();
            IUIAutomationWindowPattern windowPattern = element.GetCurrentPattern((int)Pattern.Window);
            switch (state.ToLower())
            {
                case ("min"):
                    windowPattern.SetWindowVisualState(WindowVisualState.WindowVisualState_Minimized);
                    Thread.Sleep(_sleep);
                    return;
                case ("max"):
                    windowPattern.SetWindowVisualState(WindowVisualState.WindowVisualState_Maximized);
                    Thread.Sleep(_sleep);
                    return;
                case ("normal"):
                    windowPattern.SetWindowVisualState(WindowVisualState.WindowVisualState_Normal);
                    Thread.Sleep(_sleep);
                    return;
            }
            throw new ArgumentOutOfRangeException($"Visual state '{state}' does not exist. Allowed: Min, Max or Normal");
        }

        /// <summary>
        /// <b>User32.dll</b> overload. You must specify Process name when using this method.<br/>State:<i> min, max, normal</i>
        /// </summary>
        /// <param name="processName"></param>
        /// <param name="state"></param>
        /// <exception cref="ArgumentOutOfRangeException"></exception>
        public void Window(string processName, string state)
        {
            IntPtr windowHandle = GetWindowHandle(processName);
            if (windowHandle == IntPtr.Zero) throw new ArgumentNullException($"Failed to obtain handle to process '{processName}'");

            switch (state.ToLower())
            {
                case ("min"):
                    ShowWindow(windowHandle, (int)WindowControls.SW_MINIMIZE);
                    Thread.Sleep(_sleep);
                    return;
                case ("max"):
                    ShowWindow(windowHandle, (int)WindowControls.SW_MAXIMIZE);
                    Thread.Sleep(_sleep);
                    return;
                case ("normal"):
                    ShowWindow(windowHandle, (int)WindowControls.SW_SHOWNORMAL);
                    Thread.Sleep(_sleep);
                    return;
            }
            throw new ArgumentOutOfRangeException($"Visual state '{state}' does not exist. Allowed: Min, Max or Normal");
        }

        private IntPtr GetWindowHandle(string title)
        {
            IntPtr handle = IntPtr.Zero;
            foreach (Process process in Process.GetProcesses())
            {
                if (Regex.IsMatch(process.ToString(), title, RegexOptions.IgnoreCase))
                {
                    return process.MainWindowHandle;
                }
            }
            return IntPtr.Zero;
        }

        public void CloseWindow(string title)
        {
            foreach (Process process in Process.GetProcesses())
            {
                if (process.MainWindowTitle == title)
                {
                    process.CloseMainWindow();
                }
            }
        }

        public void SendKey(string keys)
        {
            Thread.Sleep(_sleep);
            SendKeys.SendWait(keys);
            Thread.Sleep(_sleep);
        }

        public List<string> Children(Dictionary<int, string> elementIdentificator, Property returnValue)
        {
            List<string> children = new List<string>();
            IUIAutomationElement element = FindFirst();
            int elementIdentificatorCount = elementIdentificator.Count;

            IUIAutomationCondition[] conditionArray = new IUIAutomationCondition[elementIdentificatorCount];
            for (int i = 0; i < elementIdentificatorCount; i++)
            {
                conditionArray[i] = _automation.CreatePropertyCondition(elementIdentificator.ElementAt(i).Key, elementIdentificator.ElementAt(i).Value);
            }

            IUIAutomationCondition elementCondition = _automation.CreateAndConditionFromArray(conditionArray);
            //IUIAutomationElement element = _root.FindFirst(TreeScope.TreeScope_Descendants, elementCondition);


            //---
            // get Options
            IUIAutomationElementArray elementsArray = element.FindAll(TreeScope.TreeScope_Children, elementCondition);
            int elementsArrayLength = elementsArray.Length;
            Console.WriteLine($"Options count: {elementsArrayLength}");

            if (elementsArrayLength > 0)
            {
                for (int i = 0; i < elementsArrayLength; i++)
                {
                    IUIAutomationElement elementOptionElement = elementsArray.GetElement(i);
                    string elementCurrentName = elementOptionElement.CurrentName;

                    //if (elementCurrentName.Contains("Person Michal Litavsky")) // could be "Group chat" or smth
                    //{
                    //    // click on element
                    //    IUIAutomationCondition optionButtonCondition = automation.CreatePropertyCondition((int)PropertyIdEnum.AriaRole, "button");
                    //    IUIAutomationElement buttonElement = listboxOptionElement.FindFirst(TreeScope.TreeScope_Children, optionButtonCondition);
                    //    IUIAutomationLegacyIAccessiblePattern legacyButtonElement = buttonElement.GetCurrentPattern((int)ControlPatternEnum.LegacyIAccessible);
                    //    legacyButtonElement.DoDefaultAction();

                    //    break;
                    //}
                    children.Add(elementCurrentName);
                }
            }

            // ---

            return children;
        }

        enum WindowControls
        {
            // Description: https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindow?redirectedfrom=MSDN
            // Enumeration: http://pinvoke.net/default.aspx/Enums/SHOWWINDOW_FLAGS.html?diff=y
            SW_HIDE = 0,
            SW_SHOWNORMAL = 1,
            SW_NORMAL = 1,
            SW_SHOWMINIMIZED = 2,
            SW_SHOWMAXIMIZED = 3,
            SW_MAXIMIZE = 3,
            SW_SHOWNOACTIVATE = 4,
            SW_SHOW = 5,
            SW_MINIMIZE = 6,
            SW_SHOWMINNOACTIVE = 7,
            SW_SHOWNA = 8,
            SW_RESTORE = 9,
            SW_SHOWDEFAULT = 10,
            SW_FORCEMINIMIZE = 11,
            SW_MAX = 11
        }

        enum ButtonControls
        {
            // https://wiki.winehq.org/List_Of_Windows_Messages
            BM_CLICK = 0x00F5,
            BM_GETCHECK = 0x00f0,
            BM_GETIMAGE = 0x00f6,
            BM_GETSTATE = 0x00f2,
            BM_SETCHECK = 0x00f1,
            BM_SETDONTCLICK = 0x00f8,
            BM_SETIMAGE = 0x00f7,
            BM_SETSTATE = 0x00f3,
            BM_SETSTYLE = 0x00f4
        }

    }
}