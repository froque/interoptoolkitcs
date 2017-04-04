namespace $safeprojectname$
{
    using System;
    using System.Diagnostics;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;

    using Microsoft.InteropFormTools;
    using Microsoft.VisualBasic.Devices;
    using Microsoft.Win32;

#if COM_INTEROP_ENABLED

    /* 
     * Implements access to the InteropToolbox in similar fashion to the
     * vb version of the library. Access by using My.InteropToolbox.
     */

    internal static class My
    {
        private static readonly InteropToolbox Toolbox;

        static My()
        {
            Toolbox = new InteropToolbox();
        }

        public static InteropToolbox InteropToolbox
        {
            get
            {
                return Toolbox;
            }
        }
    }

    internal static class ComRegistration
    {
        const int OLEMISC_RECOMPOSEONRESIZE = 1;
        const int OLEMISC_CANTLINKINSIDE = 16;
        const int OLEMISC_INSIDEOUT = 128;
        const int OLEMISC_ACTIVATEWHENVISIBLE = 256;
        const int OLEMISC_SETCLIENTSITEFIRST = 131072;

        public static void RegisterControl(Type t)
        {
            try
            {
                GuardNullType(t, "t");
                GuardTypeIsControl(t);

                // CLSID
                string key = @"CLSID\" + t.GUID.ToString("B");

                using (RegistryKey subkey = Registry.ClassesRoot.OpenSubKey(key, true))
                {

                    // InProcServer32
                    RegistryKey inprocKey = subkey.OpenSubKey("InprocServer32", true);
                    if (inprocKey != null)
                    {
                        inprocKey.SetValue(null, Environment.SystemDirectory + @"\mscoree.dll");
                    }

                    //Control
                    using (subkey.CreateSubKey("Control"))
                    { }

                    //Misc
                    using (RegistryKey miscKey = subkey.CreateSubKey("MiscStatus"))
                    {
                        const int MiscStatusValue = OLEMISC_RECOMPOSEONRESIZE +
                                                    OLEMISC_CANTLINKINSIDE + OLEMISC_INSIDEOUT +
                                                    OLEMISC_ACTIVATEWHENVISIBLE + OLEMISC_SETCLIENTSITEFIRST;

                        miscKey.SetValue("", MiscStatusValue.ToString(), RegistryValueKind.String);
                    }

                    // ToolBoxBitmap32
                    using (RegistryKey bitmapKey = subkey.CreateSubKey("ToolBoxBitmap32"))
                    {
                        // If you want to have different icons for each control in this assembly
                        // you can modify this section to specify a different icon each time.
                        // Each specified icon must be embedded as a win32resource in the
                        // assembly; the default one is at index 101, but you can add additional ones.
                        bitmapKey.SetValue("", Assembly.GetExecutingAssembly().Location + ", 101",
                                           RegistryValueKind.String);
                    }

                    // TypeLib
                    using (RegistryKey typeLibKey = subkey.CreateSubKey("TypeLib"))
                    {
                        Guid libId = Marshal.GetTypeLibGuidForAssembly(t.Assembly);
                        typeLibKey.SetValue("", libId.ToString("B"), RegistryValueKind.String);
                    }

                    // Version
                    using (RegistryKey versionKey = subkey.CreateSubKey("Version"))
                    {
                        int major, minor;
                        Marshal.GetTypeLibVersionForAssembly(t.Assembly, out major, out minor);
                        versionKey.SetValue("", String.Format("{0}.{1}", major, minor));
                    }

                }

                const string Source = "Host .NET Interop UserControl in VB6";
                const string Log = "Application";
                string sEvent = "Registration successful: key = " + key;

                if (!EventLog.SourceExists(Source))
                    EventLog.CreateEventSource(Source, Log);

                EventLog.WriteEntry(Source, sEvent, EventLogEntryType.Warning, 234);
            }
            catch (Exception ex)
            {
                LogAndRethrowException("ComRegisterFunction failed.", t, ex);
            }
        }

        public static void UnregisterControl(Type t)
        {
            try
            {
                GuardNullType(t, "t");
                GuardTypeIsControl(t);

                // CLSID
                string key = @"CLSID\" + t.GUID.ToString("B");
                Registry.ClassesRoot.DeleteSubKeyTree(key);
            }
            catch (Exception ex)
            {
                LogAndRethrowException("ComUnregisterFunction failed.", t, ex);
            }

        }

        private static void GuardNullType(Type t, string param)
        {
            if (null == t)
            {
                throw new ArgumentException("The CLR type must be specified.", param);
            }
        }


        private static void GuardTypeIsControl(Type t)
        {
            if (!typeof(Control).IsAssignableFrom(t))
            {
                throw new ArgumentException("Type argument must be a Windows Forms control.");
            }
        }

        private static void LogAndRethrowException(string message, Type t, Exception ex)
        {
            try
            {
                if (null != t)
                {
                    message += Environment.NewLine + String.Format("CLR class '{0}'", t.FullName);
                }

                throw new ComRegistrationException(message, ex);
            }
            catch (Exception ex2)
            {
                const string Source = "Host .NET Interop UserControl in VB6";
                const string Log = "Application";
                string sEvent = t.GUID.ToString("B") + " registration failed: " + Environment.NewLine + ex2.Message;

                if (!EventLog.SourceExists(Source))
                    EventLog.CreateEventSource(Source, Log);

                EventLog.WriteEntry(Source, sEvent, EventLogEntryType.Warning, 234);
            }
        }
    }

    [Serializable]
    public class ComRegistrationException : Exception
    {
        public ComRegistrationException() { }
        public ComRegistrationException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }



    // Helper functions to convert common COM types to their .NET equivalents
    [ComVisible(false)]
    internal class ActiveXControlHelpers : AxHost
    {
        internal ActiveXControlHelpers()
            : base(null)
        {
        }

        internal static System.Drawing.Color GetColorFromOleColor(int oleColor)
        {
            return GetColorFromOleColor(CIntToUInt(oleColor));
        }

        internal static new int GetOleColorFromColor(System.Drawing.Color color)
        {
            return CUIntToInt(AxHost.GetOleColorFromColor(color));
        }

        internal static int CUIntToInt(uint uiArg)
        {
            if (uiArg <= int.MaxValue)
            {
                return (int)uiArg;
            }

            return (int)(uiArg - unchecked(2 * ((uint)(int.MaxValue) + 1)));
        }

        internal static uint CIntToUInt(int iArg)
        {
            if (iArg < 0)
            {
                return (uint)(uint.MaxValue + iArg + 1);
            }
            return (uint)iArg;
        }

        private const int KEY_PRESSED = 0x1000;

        [DllImport("user32.dll")]
        static extern short GetKeyState(int nVirtKey);

        private static int CheckForAccessorKey()
        {
            Keyboard keyboard = new Keyboard();
            if (keyboard.AltKeyDown)
            {
                for (int i = (int)Keys.A; i <= (int)Keys.Z; i++)
                {
                    if ((GetKeyState(i) != 0 && KEY_PRESSED != 0))
                    {
                        return i;
                    }
                }
            }
            return -1;
        }


        [ComVisible(false)]
        internal static void HandleFocus(UserControl f)
        {
            Keyboard keyboard = new Keyboard();
            if (keyboard.AltKeyDown)
            {
                HandleAccessorKey(f.GetNextControl(null, true), f);
            }
            else
            {
                // Move to the first control that can receive focus, taking into account
                // the possibility that the user pressed <Shift>+<Tab>, in which case we
                // need to start at the end and work backwards.
                Control ctl = f.GetNextControl(null, !keyboard.ShiftKeyDown);
                while (null != ctl)
                {
                    if (ctl.Enabled && ctl.CanSelect)
                    {
                        ctl.Focus();
                        break;
                    }
                    ctl = f.GetNextControl(ctl, !keyboard.ShiftKeyDown);
                }
            }
        }

        private static void HandleAccessorKey(object sender, UserControl f)
        {
            int key = CheckForAccessorKey();
            if (key == -1)
                return;

            Control ctlCurrent = f.GetNextControl((Control)sender, false);

            do
            {
                ctlCurrent = f.GetNextControl(ctlCurrent, true);
                if (ctlCurrent != null && Control.IsMnemonic(Convert.ToChar(key), ctlCurrent.Text))
                {
                    // VB6 handles conflicts correctly already, so if we handle it also we'll end up 
                    // one control past where the focus should be
                    if (!KeyConflict(Convert.ToChar(key), f))
                    {

                        // If we land on a label or other non-selectable control then go to the next 
                        // control in the tab order
                        if (!ctlCurrent.CanSelect)
                        {
                            Control ctlAfterLabel = f.GetNextControl(ctlCurrent, true);
                            if (ctlAfterLabel != null && ctlAfterLabel.CanFocus)
                            {
                                ctlAfterLabel.Focus();
                            }
                        }
                        else
                        {
                            ctlCurrent.Focus();
                        }
                        break;
                    }
                }

                // Loop until we hit the end of the tab order
                // If we've hit the end of the tab order we don't want to loop back because the
                // parent form's controls come next in the tab order.
            } while (ctlCurrent != null);
        }


        private static bool KeyConflict(char key, UserControl u)
        {
            bool flag = false;
            foreach (Control ctl in u.Controls)
            {
                if (Control.IsMnemonic(key, ctl.Text))
                {
                    if (flag)
                    {
                        return true;
                    }
                    flag = true;
                }
            }
            return false;
        }

        // Handles <Tab> and <Shift>+<Tab>
        internal static void TabHandler(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Tab)
            {
                Control ctl = sender as Control;
                UserControl userCtl = GetParentUserControl(ctl);
                Control firstCtl = userCtl.GetNextControl(null, true);
                do
                {
                    firstCtl = userCtl.GetNextControl(firstCtl, true);
                } while (firstCtl != null && !firstCtl.CanSelect);

                Control lastCtl = userCtl.GetNextControl(null, false);
                do
                {
                    lastCtl = userCtl.GetNextControl(lastCtl, false);
                } while (lastCtl != null && lastCtl.CanSelect);

                if (ctl.Equals(lastCtl) || ctl.Equals(firstCtl) || lastCtl.Contains(ctl) || firstCtl.Contains(ctl))
                {
                    userCtl.SelectNextControl((Control)sender, lastCtl.Equals(userCtl.ActiveControl), true, true, true);
                }
            }
        }

        private static UserControl GetParentUserControl(Control ctl)
        {
            if (ctl == null)
            {
                return null;
            }

            do
            {
                ctl = ctl.Parent;
            } while (ctl.Parent != null);

            if (ctl != null)
            {
                return (UserControl)ctl;
            }

            return null;
        }

        internal static void WireUpHandlers(Control ctl, EventHandler validationHandler)
        {
            if (ctl != null)
            {
                ctl.KeyDown += TabHandler;
                ctl.LostFocus += validationHandler;

                if (ctl.HasChildren)
                {
                    foreach (Control child in ctl.Controls)
                    {
                        WireUpHandlers(child, validationHandler);
                    }
                }
            }
        }
    }
#endif
}