using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Text;
using System.Threading;
using System.Net.NetworkInformation;

namespace DualisScheduleSync
{
    static class Program
    {
        static DualisScheduleSync d = new DualisScheduleSync();
        static NotifyIcon n = new NotifyIcon();

        static MenuItem syncItem = new MenuItem()
        {
            Index = 0,
            Text = "&Sync"
        };
        static MenuItem nextSync = new MenuItem()
        {
            Index = 1,
            Enabled = false,
            Text = "n/a"
        };
        static MenuItem lastSync = new MenuItem()
        {
            Index = 2,
            Enabled = false,
            Text = "n/a"
        };
        static MenuItem exitItem = new MenuItem()
        {
            Index = 3,
            Text = "E&xit"
        };

        static System.Windows.Forms.Timer t = new System.Windows.Forms.Timer();

        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            ContextMenu contextMenu = new ContextMenu();

            contextMenu.MenuItems.AddRange(new MenuItem[] { syncItem, nextSync, lastSync, exitItem });

            exitItem.Click += new EventHandler(exitItem_Click);
            syncItem.Click += new EventHandler(syncItem_Click);

            n.Text = "DualisScheduleSync";
            n.ContextMenu = contextMenu;
            n.Visible = true;
            n.Icon = new System.Drawing.Icon(@"D:\Icons\cal.ico");
            n.BalloonTipTitle = "Next Sync";
            n.BalloonTipText = "n/a";

            t.Tick += new EventHandler(syncItem_Click);

            syncItem_Click(null, null);

            Application.Run();
            n.Visible = false;
        }

        static void syncItem_Click(object sender, EventArgs e)
        {
            if (t.Enabled)
            {
                t.Stop();
            }
            lastSync.Text = "Last run: " + DateTime.Now.TimeOfDay.ToString("hh\\:mm\\:ss");
            if (HasConnection() && d.Sync())
            {
                n.ShowBalloonTip(3000, "Sync succeeded", "Sync with Dualis done", ToolTipIcon.Info);
                t.Interval = 1000 * 60 * 30;
            }
            else
            {
                n.ShowBalloonTip(3000, "Sync Failed", d.Error ?? "Did not provide a reason", ToolTipIcon.Error);
                t.Interval = 1000 * 60 * 5;
            }
            n.BalloonTipText = DateTime.Now.AddMilliseconds(t.Interval).TimeOfDay.ToString("hh\\:mm\\:ss");
            nextSync.Text = "Next run: " + n.BalloonTipText;
            t.Start();
        }

        private static void exitItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private static bool HasConnection()
        {
            Ping myPing = new Ping();
            String host = "google.com";
            byte[] buffer = new byte[32];
            int timeout = 1000;
            PingOptions pingOptions = new PingOptions();
            PingReply reply = myPing.Send(host, timeout, buffer, pingOptions);
            if (reply.Status == IPStatus.Success)
            {
                return true;
            }
            return false;
        }
    }
}
