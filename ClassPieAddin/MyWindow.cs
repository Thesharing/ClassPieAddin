using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ClassPieAddin {
    public partial class MyWindow : Form {
        public MyWindow() {
            VisibleChanged += MyWindow_VisibleChanged;
            HandleCreated += (s, e) => _handle = Handle;
        }

        bool _Layered = false;
        bool _NoActivate = false;
        bool _Back = false;
        bool _Penetrate = false;
        /// <summary>
        /// 指示窗口是否为分层窗口
        /// </summary>
        public bool Layered {
            get {
                return _Layered;
            }

            set {
                _layeredAlpha = (int)(base.Opacity * 255);
                _Layered = value;
                updateWindowLong();
            }
        }
        /// <summary>
        /// 指示窗口是否能获取焦点
        /// </summary>
        public bool NoActivate {
            get {
                return _NoActivate;
            }

            set {
                _NoActivate = value;
                updateWindowLong();
            }
        }

        public bool Back {
            get {
                return _Back;
            }

            set {
                _Back = value;
                updateWindowLong();
            }
        }
        /// <summary>
        /// 指示窗口是否穿透鼠标
        /// </summary>
        public bool Penetrate {
            get {
                return _Penetrate;
            }

            set {
                _Penetrate = value;
                updateWindowLong();
            }
        }

        private void updateWindowLong() {
            if (IsHandleCreated) {
                Win32.SetWindowLong(Handle, Win32.GWL_EXSTYLE, (uint)CreateParams.ExStyle);
            }
        }

        private void MyWindow_VisibleChanged(object sender, EventArgs e) {
            if (_Back && Visible) {
                BringToFront(); SendToBack();
            }
        }
        protected override CreateParams CreateParams {
            get {
                CreateParams cp = base.CreateParams;
                if (_Layered)
                    cp.ExStyle |= Win32.WS_EX_LAYERED;
                if (_NoActivate)
                    cp.ExStyle |= Win32.WS_EX_NOACTIVATE;
                if (TopMost)
                    cp.ExStyle |= Win32.WS_EX_TOPMOST;
                if (_Penetrate)
                    cp.ExStyle |= Win32.WS_EX_TRANSPARENT;
                return cp;
            }
        }

        int _layeredAlpha = 255;
        public int LayeredAlpha {
            get {
                return _layeredAlpha;
            }
            set {
                if (value <= 0)
                    _layeredAlpha = 1;
                else
                    _layeredAlpha = value;
                if (IsHandleCreated)
                    UpdateLayeredWindow(TrueLayeredAlpha);
            }
        }
        double _opacity = 1;
        public new double Opacity {
            get {
                if (Layered)
                    return LayeredAlpha / 255.0;
                else
                    return _opacity;
            }
            set {
                if (Layered)
                    LayeredAlpha = (int)(value * 255);
                else
                    base.Opacity = TrueOpacity;
            }
        }

        private double fadeOpacity = 1;
        public virtual double TrueOpacity => Opacity * FadeOpacity;
        public int TrueLayeredAlpha => (int)(TrueOpacity * 255);

        public double FadeOpacity {
            get {
                return fadeOpacity;
            }

            set {
                if (value < 0 || value > 1)
                    return;
                fadeOpacity = value;
                if (IsHandleCreated)
                    Opacity = Opacity;
            }
        }
        IntPtr _handle = IntPtr.Zero;

        private void UpdateLayeredWindow(int alpha) {
            if (_handle == IntPtr.Zero) {
                //Debug("UpdateLayeredWindow with No Handle", Log.Level.Warning);
                return;
            }
            Win32.BLENDFUNCTION blendFunc = new Win32.BLENDFUNCTION();

            blendFunc.BlendOp = Win32.AC_SRC_OVER;
            if (alpha > 255)
                alpha = 255;
            blendFunc.SourceConstantAlpha = (byte)alpha;
            blendFunc.AlphaFormat = Win32.AC_SRC_ALPHA;
            blendFunc.BlendFlags = 0;
            try {
                int ret = Win32.UpdateLayeredWindow2(_handle, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, 0, ref blendFunc, Win32.ULW_ALPHA);
                if (ret != 1) {
                    //Debug($"UpdateLayeredWindow Failed({ret}) ({this.GetType()})", Log.Level.Warning);
                }
            }
            finally { }
        }

        public int UpdateLayeredWindow(Bitmap bitmap) => UpdateLayeredWindow(bitmap, TrueLayeredAlpha);
        public int UpdateLayeredWindow(Bitmap bitmap, int alpha) => UpdateLayeredWindow(bitmap, alpha, Size);
        public int UpdateLayeredWindow(Bitmap bitmap, int alpha, Size size) {
            int ret = -1;
            if (_handle == IntPtr.Zero) {
                //Debug("UpdateLayeredWindow with No Handle", Log.Level.Warning);
                return ret;
            }
            if (!Bitmap.IsCanonicalPixelFormat(bitmap.PixelFormat) || !Bitmap.IsAlphaPixelFormat(bitmap.PixelFormat)) {
                return ret;
            }
            IntPtr oldBits = IntPtr.Zero;
            IntPtr screenDC = Win32.GetDC(IntPtr.Zero);
            IntPtr hBitmap = IntPtr.Zero;
            IntPtr memDc = Win32.CreateCompatibleDC(screenDC);

            try {
                Win32.Point topLoc = new Win32.Point(Left, Top);
                Win32.Size bitMapSize = new Win32.Size(size.Width, size.Height);
                Win32.BLENDFUNCTION blendFunc = new Win32.BLENDFUNCTION();
                Win32.Point srcLoc = new Win32.Point(0, 0);

                hBitmap = bitmap.GetHbitmap(Color.FromArgb(0));
                oldBits = Win32.SelectObject(memDc, hBitmap);

                blendFunc.BlendOp = Win32.AC_SRC_OVER;
                blendFunc.SourceConstantAlpha = (byte)alpha;
                blendFunc.AlphaFormat = Win32.AC_SRC_ALPHA;
                blendFunc.BlendFlags = 0;

                ret = Win32.UpdateLayeredWindow(_handle, screenDC, ref topLoc, ref bitMapSize, memDc, ref srcLoc, 0, ref blendFunc, Win32.ULW_ALPHA);
                if (ret != 1) {
                    //Debug($"UpdateLayeredWindow Failed({ret}) ({this.GetType()})", Log.Level.Warning);
                }
            }
            finally {
                if (hBitmap != IntPtr.Zero) {
                    Win32.SelectObject(memDc, oldBits);
                    Win32.DeleteObject(hBitmap);
                }
                Win32.ReleaseDC(IntPtr.Zero, screenDC);
                Win32.DeleteDC(memDc);
            }
            return ret;
        }

        public delegate void FadeCallBack(MyWindow sender, bool done);
        class FadeTask {
            public FadeCallBack callback;
            public double opa;
            public int time;
        }


        public void FadeTo(double opa, int time = 500, FadeCallBack callback = null, bool sync = false) {
            if (opa < 0 || opa > 1 || time < 14) {
                //Debug($"MyWindow.FadeTo() wrong arg (opa={opa} time={time})", Log.Level.Warning);
                return;
            }
            EventWaitHandle ewh = new EventWaitHandle(false, EventResetMode.ManualReset);
            var temp = callback;
            callback = (s, d) => {
                ewh.Set();
                temp?.Invoke(s, d);
            };
            var task = new FadeTask() { callback = callback, opa = opa, time = time };
            lock (_fadeThreadLocker) {
                _currentTask = task;
            }
            if (sync)
                _fadeThread(task);
            else
                new Thread(() => _fadeThread(task)) { Name = "MyWindow FadeThread" }.Start();
        }
        object _fadeThreadLocker = new object();
        FadeTask _currentTask;
        void _fadeThread(FadeTask task) {
            try {
                while (true) {
                    lock (_fadeThreadLocker) {
                        if (task.opa == FadeOpacity) {
                            task.callback?.Invoke(this, true);
                            break;
                        }
                        if (_currentTask != task) {
                            task.callback?.Invoke(this, false);
                            break;
                        }
                    }

                    if (task.time < 14) {
                        _fadeThr_setOpa(task.opa);
                        task.callback?.Invoke(this, true);
                        break;
                    }
                    var minus = task.opa - FadeOpacity;
                    var setopa = FadeOpacity + minus / task.time * 14;
                    _fadeThr_setOpa(setopa);
                    task.time -= 14;
                    Thread.Sleep(14);
                }
            }
            catch (Exception ex) {
                if (ex is ObjectDisposedException)
                    return;
                //Debug("FadeThread Exception.", Log.Level.Warning, ex);
            }
        }
        void _fadeThr_setOpa(double opa) {
            if (Layered)
                FadeOpacity = opa;
            else
                this.Invoke(new MethodInvoker(() => FadeOpacity = opa));
        }
    }
}