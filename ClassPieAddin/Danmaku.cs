using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ClassPieAddin {
    public class Danmaku : MyWindow, IDanmaku {

        public Rectangle screenBounds = Screen.PrimaryScreen.Bounds;
        string text = null;
        DanmakuEngine engine;

        public Danmaku(DanmakuEngine engine, string text, Color color, Font font) {
            NoActivate = true;
            Penetrate = true;
            Layered = true;
            this.engine = engine;
            this.text = text;
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            Font = font;
            ForeColor = color;

            update();
            this.Shown += Danmaku_Shown;
        }

        private void Danmaku_Shown(object sender, EventArgs e) {
            this.TopMost = true;
            IsShown = true;
            Penetrate = true;
        }

        public bool IsShown = false;

        Color IDanmaku.Color {
            get {
                return ForeColor;
            }

            set {
                ForeColor = value;
                update();
            }
        }

        void IDanmaku.Close() {
            try {
                if (IsDisposed == false) {
                    Invoke(new MethodInvoker(() => this.Close()));
                }
            }
            catch (Exception) { } // ignore
        }

        void update() {
            Graphics g = CreateGraphics();
            Size = g.MeasureString(text, Font).ToSize();
            using (Bitmap bitmap = new Bitmap(Width, Height))
            using (Graphics gb = Graphics.FromImage(bitmap)) {
                paint(gb);
                UpdateLayeredWindow(bitmap);
            }
        }

        void paint(Graphics g) {
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
            Brush b = new SolidBrush(ForeColor);
            g.DrawString(text, Font, b, 0, 0);
        }

    }
}