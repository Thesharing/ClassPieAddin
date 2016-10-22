using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ClassPieAddin {
    public interface IDanmakuEngine : IDisposable, IEnumerable<IDanmaku> {
        void ShowDanmaku(string text, Color color, Font font);
        bool Hidden { get; set; }
        int DanmakuCount { get; }
    }

    public interface IDanmaku {
        string Text { get; }
        Color Color { get; set; }
        void Close();
    }

    public class DanmakuEngine : IDanmakuEngine {
        List<Danmaku> danmakuList = new List<Danmaku>();
        public IList<Danmaku> DanmakuList => danmakuList;
        public bool Stop { get; set; }
        public int DanmakuCount => danmakuList.Count;

        bool hidden = false;
        public bool Hidden {
            get {
                return hidden;
            }

            set {
                if (hidden == value)
                    return;
                hidden = value;
                for (int i = 0; i < danmakuList.Count; i++) {
                    try {
                        var d = danmakuList[i];
                        if (hidden) {
                            d.BeginInvoke(new MethodInvoker(() => d.Hide()));
                        }
                        else {
                            d.BeginInvoke(new MethodInvoker(() => d.Show()));
                        }
                    }
                    catch (Exception) { } // ignore
                }
            }
        }

        public DanmakuEngine() {
            startEngineThread();
        }

        public void ShowDanmaku(string text, Color color, Font font) {
            if (disposed)
                return;
            new Thread(() => {
                var d = new Danmaku(this, text, color, font);
                setDanmakuStartPosition(d);
                d.FormClosed += (s, e) => {
                    try {
                        danmakuList.Remove(d);
                    }
                    catch (Exception) { } // ignore
                };
                if (hidden)
                    d.Hide();
                danmakuList.Add(d);
                Application.Run(d);
            }) { IsBackground = true }.Start();
        }

        public void ShowDanmaku(string text, Color color, Font font, int left, int top) {
            if (disposed)
                return;
            new Thread(() => {
                var d = new Danmaku(this, text, color, font);
                setDanmakuStartPosition(d, left, top);
                d.FormClosed += (s, e) => {
                    try {
                        danmakuList.Remove(d);
                    }
                    catch (Exception) { } // ignore
                };
                if (hidden)
                    d.Hide();
                danmakuList.Add(d);
                Application.Run(d);
            }) { IsBackground = true }.Start();
        }

#pragma warning disable CS1690 // 访问引用封送类的字段上的成员可能导致运行时异常

        void setDanmakuStartPosition(Danmaku d) {
            d.Location = new Point(d.screenBounds.Right, d.screenBounds.Bottom - 100);
            lock (danmakuList) {
                while (true) {
                    bool hasIntersct = false;
                    for (int i = 0; i < danmakuList.Count; i++) {
                        var d2 = danmakuList[i];
                        var bounds = d2.Bounds;
                        bounds.Width += 20;
                        if (bounds.IntersectsWith(d.Bounds)) {
                            hasIntersct = true;
                            d.Top = d2.Bounds.Bottom;
                            break;
                        }
                    }
                    if (!hasIntersct)
                        break;
                }
            }
        }

        void setDanmakuStartPosition(Danmaku d, int left, int top) {
            d.Location = new Point(left, top);
            lock (danmakuList) {
                while (true) {
                    bool hasIntersct = false;
                    for (int i = 0; i < danmakuList.Count; i++) {
                        var d2 = danmakuList[i];
                        var bounds = d2.Bounds;
                        bounds.Width += 20;
                        if (bounds.IntersectsWith(d.Bounds)) {
                            hasIntersct = true;
                            d.Top = d2.Bounds.Bottom;
                            break;
                        }
                    }
                    if (!hasIntersct)
                        break;
                }
            }
        }

        void startEngineThread() {
            new Thread(() => {
                while (true) {
                    Thread.Sleep(16);
                    if (danmakuList.Count < 0)
                        continue;
                    try {
                        for (int i = danmakuList.Count - 1; i >= 0; i--) {
                            var d = danmakuList[i];
                            if (d.IsShown) {
                                d.BeginInvoke(new MethodInvoker(() => {
                                    d.Left -= Setting.speed;
                                    if (d.Bounds.Right < d.screenBounds.Left) {
                                        d.Close();
                                    }
                                }));
                            }
                        }
                    }
                    catch (Exception) { } // ignore
                }
            }) { IsBackground = true }.Start();
        }

        bool disposed = false;
        public void Dispose() {
            if (disposed)
                return;
            disposed = true;
            for (int i = danmakuList.Count - 1; i >= 0; i--) {
                try {
                    danmakuList[i].Close();
                }
                catch (Exception) { } // ignore
            }
        }
#pragma warning restore CS1690

        public IEnumerator<IDanmaku> GetEnumerator() {
            return danmakuList.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator() {
            return danmakuList.GetEnumerator();
        }
    }
}