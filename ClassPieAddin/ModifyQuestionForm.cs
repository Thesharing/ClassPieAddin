using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;

namespace ClassPieAddin {
    public partial class ModifyQuestionForm : UserControl {
        public List<TextBox> textboxArray = new List<TextBox>();
        public List<Label> labelArray = new List<Label>();
        public int count = 0;
        public PowerPoint.Slide slide;

        public ModifyQuestionForm() {
            InitializeComponent();
            Init();
        }

        public void Init() {
            try {
                PowerPoint.Application Application = Globals.ThisAddIn.Application;
                PowerPoint.Slides slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;//获取当前应用程序的所有PPT文档
                if (Application.SlideShowWindows.Count > 0) {
                    slide = slides._Index(Application.SlideShowWindows[1].View.Slide.SlideIndex);
                }
                else {
                    slide = slides._Index(Application.ActiveWindow.Selection.SlideRange.SlideIndex);
                }
                if (slide.Tags["Question"] == "Yes") {
                    textBoxQuestion.Text = slide.Shapes["TextBoxQuestion"].TextFrame.TextRange.Text;
                    textBoxQuestion.TextChanged += ChoiceBox_TextChanged;
                    for (int i = 1; i <= 26; i++) {
                        if (slide.Shapes["TextBoxAnswer" + i.ToString()].TextFrame.TextRange.Text != null) {
                            TextBox text = new TextBox();
                            this.panel1.Controls.Add(text);
                            text.Location = new Point(86, 45 + 41 * i);
                            text.Name = "textBox" + i.ToString();
                            text.Size = new System.Drawing.Size(129, 21);
                            text.Text = slide.Shapes["TextBoxAnswer" + i.ToString()].TextFrame.TextRange.Text;
                            Label label = new Label();
                            this.panel1.Controls.Add(label);
                            label.Location = new Point(20, 47 + 40 * i);
                            label.Name = "label" + i.ToString();
                            label.Text = "选项 " + i.ToString();
                            textboxArray.Add(text);
                            labelArray.Add(label);
                        }
                        count++;
                    }
                }
            }
            catch (COMException error) {
            }
            for (int i = 0; i < count; i++) {
                textboxArray[i].TextChanged += ChoiceBox_TextChanged;
            }
        }

        private void ChoiceBox_TextChanged(object sender, EventArgs e) {
            try {
                if (slide.Tags["Question"] == "Yes") {
                    slide.Shapes["TextBoxQuestion"].TextFrame.TextRange.Text = textBoxQuestion.Text;
                    for (int i = 1; i <= 26; i++) {
                        slide.Shapes["TextBoxAnswer" + i.ToString()].TextFrame.TextRange.Text = textboxArray[i - 1].Text;
                    }
                }
            }
            catch {
                return;
            }
        }

        private void buttonAdd_Click(object sender, EventArgs e) {
            count++;
            TextBox text = new TextBox();
            this.panel1.Controls.Add(text);
            text.Location = new Point(86, 45 + 41 * (count));
            text.Name = "textBox" + count.ToString();
            text.Size = new System.Drawing.Size(129, 21);
            Label label = new Label();
            this.panel1.Controls.Add(label);
            label.Location = new Point(20, 47 + 40 * count);
            label.Name = "label" + count.ToString();
            label.Text = "选项 " + count.ToString();
            textboxArray.Add(text);
            var textBoxAnswer = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 150 + count * 50, 600, 50);
            textBoxAnswer.Name = "TextBoxAnswer" + count.ToString();
            text.Text = textBoxAnswer.TextFrame.TextRange.Text = ((char)('A' + count - 1)).ToString() + ". 选项 " + count.ToString() + ".";
            textBoxAnswer.TextFrame.TextRange.Font.Size = 24;//设置文本字体大小
            text.TextChanged += ChoiceBox_TextChanged;
            labelArray.Add(label);
        }

        private void buttonDelete_Click(object sender, EventArgs e) {
            if (count > 0) {
                count--;
                this.panel1.Controls.Remove(textboxArray[count]);
                textboxArray.RemoveAt(count);
                this.panel1.Controls.Remove(labelArray[count]);
                labelArray.RemoveAt(count);
                System.Diagnostics.Debug.WriteLine(count + 1);
                slide.Shapes["TextBoxAnswer" + (count + 1).ToString()].Delete();
            }
            else {
                MessageBox.Show("没有可删除的选项。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
