using System;
using System.Text;
using System.Windows.Forms;
using Microsoft.VisualBasic.Logging;
using Microsoft.VisualBasic.ApplicationServices;
using Microsoft.VisualBasic.Devices;
using Microsoft.Win32;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace CSMultithreadedControl
{


    internal sealed class ComRegistration
    {
        private ComRegistration()
        {
        }

        const int OLEMISC_RECOMPOSEONRESIZE = 1;
        const int OLEMISC_CANTLINKINSIDE = 16;
        const int OLEMISC_INSIDEOUT = 128;
        const int OLEMISC_ACTIVATEWHENVISIBLE = 256;
        const int OLEMISC_SETCLIENTSITEFIRST = 131072;

        public static void RegisterControl(System.Type t)
        {
            try
            {
                GuardNullType(t, "t");
                GuardTypeIsControl(t);

                //CLSID
                string key = @"CLSID\" + t.GUID.ToString("B");

                using (RegistryKey subkey = Registry.ClassesRoot.OpenSubKey(key, true))
                {

                    //InProcServer32
                    RegistryKey InprocKey = subkey.OpenSubKey("InprocServer32", true);
                    if (InprocKey != null)
                    {
                        InprocKey.SetValue(null, Environment.SystemDirectory + @"\mscoree.dll");
                    }

                    //Control
                    using (RegistryKey controlKey = subkey.CreateSubKey("Control"))
                    { }

                    //Misc
                    using (RegistryKey miscKey = subkey.CreateSubKey("MiscStatus"))
                    {
                        int MiscStatusValue = OLEMISC_RECOMPOSEONRESIZE +
                            OLEMISC_CANTLINKINSIDE + OLEMISC_INSIDEOUT +
                            OLEMISC_ACTIVATEWHENVISIBLE + OLEMISC_SETCLIENTSITEFIRST;

                        miscKey.SetValue("", MiscStatusValue.ToString(), RegistryValueKind.String);
                    }

                    //ToolBoxBitmap32
                    using (RegistryKey bitmapKey = subkey.CreateSubKey("ToolBoxBitmap32"))
                    {
                        //'If you want to have different icons for each control in this assembly
                        //'you can modify this section to specify a different icon each time.
                        //'Each specified icon must be embedded as a win32resource in the
                        //'assembly; the default one is at index 101, but you can add additional ones.
                        bitmapKey.SetValue("", Assembly.GetExecutingAssembly().Location + ", 101",
                                           RegistryValueKind.String);
                    }

                    //TypeLib
                    using (RegistryKey typeLibKey = subkey.CreateSubKey("TypeLib"))
                    {
                        Guid libId = Marshal.GetTypeLibGuidForAssembly(t.Assembly);
                        typeLibKey.SetValue("", libId.ToString("B"), RegistryValueKind.String);
                    }

                    //Version
                    using (RegistryKey versionKey = subkey.CreateSubKey("Version"))
                    {
                        int major, minor;
                        Marshal.GetTypeLibVersionForAssembly(t.Assembly, out major, out minor);
                        versionKey.SetValue("", String.Format("{0}.{1}", major, minor));
                    }

                }
                string sSource;
                string sLog;
                string sEvent;

                sSource = "Host .NET Interop UserControl in VB6";
                sLog = "Application";
                sEvent = "Registration successful: key = " + key;

                if (!EventLog.SourceExists(sSource))
                    EventLog.CreateEventSource(sSource, sLog);

                EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Warning, 234);
                MessageBox.Show("I feel good!=)");
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

                //CLSID
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
            if (!typeof(System.Windows.Forms.Control).IsAssignableFrom(t))
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
                string sSource;
                string sLog;
                string sEvent;

                sSource = "Host .NET Interop UserControl in VB6";
                sLog = "Application";
                sEvent = t.GUID.ToString("B") + " registration failed: " + Environment.NewLine + ex2.Message;

                if (!EventLog.SourceExists(sSource))
                    EventLog.CreateEventSource(sSource, sLog);

                EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Warning, 234);
            }
        }
    }

    [Serializable()]
    public class ComRegistrationException : Exception
    {
        public ComRegistrationException() { }
        public ComRegistrationException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }




}

