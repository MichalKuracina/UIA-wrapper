#UniversalUIAClass

## Example

``` c#
static void Main(string[] args)
{
    ProcessStartInfo startInfo = new ProcessStartInfo(@"C:\Something.exe");
    startInfo.WindowStyle = ProcessWindowStyle.Maximized;
    Process.Start(startInfo);

    UIA app = new UIA(1500);
    app.AddProperty(Property.ClassName, "Button");
    app.AddProperty(Property.Name, "Connect");
    app.Invoke();

    app.AddProperty(Property.AutomationID, "mat-input-0");
    app.Invoke();
}
```


## Methods

FindFirst
Exists
Until
Invoke
Expand
Collapse
SelectionItem
VizualizedItem
Transform
Write
Window
CloseWindow
SendKey
Children

[propids](https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-automation-element-propids)
``` c#
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
```
[controlpattern-ids](https://learn.microsoft.com/en-us/windows/win32/winauto/uiauto-controlpattern-ids)
``` c#
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
```

[Description](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-showwindow)
[Enumeration](http://pinvoke.net/default.aspx/Enums/SHOWWINDOW_FLAGS.html?diff=y)

``` c#
enum WindowControls
{
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
```

[List_Of_Windows_Messages](https://wiki.winehq.org/List_Of_Windows_Messages)

``` c#
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
```
