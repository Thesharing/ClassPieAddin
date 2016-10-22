using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassPieAddin {
    public class Question {
        public int number;
        public int count;
        public string question;
        public List<Choice> choice;
        public Question(int number, string question) {
            this.number = number;
            this.question = question;
            choice = new List<Choice>();
            count = 0;
        }
        public void Add(string choice) {
            this.choice.Add(new Choice(++count,choice));
            count = this.choice.Count;
        }

        public void RemoveAt(int index) {
            this.choice.RemoveAt(index);
            count = this.choice.Count;
        }

        public void RemoveLast() {
            this.choice.RemoveAt(this.choice.Count - 1);
            count = this.choice.Count;
        }
    }
}
