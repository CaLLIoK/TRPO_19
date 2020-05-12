using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic;
using System.Security.Permissions;
using System.Drawing;
using System.Collections;
using CSMultithreadedControl;

namespace MyCalendarX
{
    [ComVisible(true), Guid(Calendar.EventsId), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface CalendarEvents
    {
        [DispId(1)]
        void DateChanged(DateTime D);

    }

    [Guid(Calendar.InterfaceId), ComVisible(true)]
    public interface CalendarProperties
    {
        [DispId(1)]
        bool Visible { [DispId(1)] get; [DispId(1)] set; }
        [DispId(2)]
        bool Enabled { [DispId(2)] get; [DispId(2)] set; }
        [DispId(3)]
        void Refresh();
        [DispId(4)]
        DateTime PrevDate { [DispId(4)] get; [DispId(4)] set; }
        

    }
    [Guid(Calendar.ClassId), ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces("MyCalendarX.CalendarEvents")]
    [ComClass(Calendar.ClassId, Calendar.InterfaceId, Calendar.EventsId)]
    public partial class Calendar :  UserControl, CalendarProperties
    {
        #region "COM Registration"

        //These  GUIDs provide the COM identity for this class 
        //and its COM interfaces. If you change them, existing 
        //clients will no longer be able to access the class.

        public const string ClassId = "28C4BD50-F055-4C78-B7DF-671966D97CF4";
        public const string InterfaceId = "E090B40A-C1C2-4D25-89C6-BB1E968A8B3A";
        public const string EventsId = "BAD5F18A-59E8-42ED-8641-90870E0534F4";


        //These routines perform the additional COM registration needed by ActiveX controls
        [EditorBrowsable(EditorBrowsableState.Never)]
        [ComRegisterFunction]
        private static void Register(System.Type t)
        {
            ComRegistration.RegisterControl(t);
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        [ComUnregisterFunction]
        private static void Unregister(System.Type t)
        {
            ComRegistration.UnregisterControl(t);
        }


        #endregion
        

        #region "Methods"

       

        public override void Refresh()
        {
            base.Refresh();
        }

        public new bool Visible
        {
            get { return base.Visible; }
            set { base.Visible = value; }
        }

        public new bool Enabled
        {
            get { return base.Enabled; }
            set
            {
                base.Enabled = value;
            }
        }

        public DateTime PrevDate
        {
            get { return Prev_Date.Date; }
            set { Prev_Date = value; }
        }

        #endregion



        #region "Events delegate"
        public delegate void DateChangedEvenHandle(DateTime D);
        public event DateChangedEvenHandle DateChanged = null;


        #endregion

        private DateTime Prev_Date;

        public Calendar()
        {
            InitializeComponent();
            Prev_Date = new DateTime();
        }

        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {
            if (null != DateChanged)
            {
                DateChanged(e.End.Date);
            }
        }
    }
}
