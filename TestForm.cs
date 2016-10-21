using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

namespace ClassPieAddin {
    public partial class TestForm : Form {
        public TestForm() {
            InitializeComponent();
            MessageBox.Show("this", "this");
        }

        private void buttonOK_Click(object sender, EventArgs e) {
            this.Close();
        }
    }
}
