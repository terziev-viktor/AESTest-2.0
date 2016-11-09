using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace AESTest2._0
{
    public partial class MainForm : Form
    {
        private string mainPath = @"C:\data\";
        private int sec = 0;
        private int min = 30;
        private int questionIndex = 0;
        private const int WS_SYSMENU = 0x80000;
        private List<int> deltaYear = new List<int>();
        private List<Question> questions = new List<Question>();
        private List<string> allEgn = new List<string>();

        private Dictionary<char, int> questionsForGroup = new Dictionary<char, int>()
        {
            {'2', 10},
            {'3', 10},
            {'4', 10},
            {'5', 15},
            {'9', 15},
        };

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            this.ShowIcon = false;
            this.MaximumSize = this.Size;
            this.MinimumSize = this.Size;
            stage_2.Enabled = false;
            stage_2.Visible = false;
            stage_3.Visible = false;
            stage_3.Enabled = false;

            this.btnStart.Enabled = false;
            this.labelTime.Location = new System.Drawing.Point(this.Width / 2 - this.labelTime.Width / 2, 0);
            this.stage_1.Size = new Size(this.Size.Width * 2 / 3, 250);
            this.stage_1.Location = new System.Drawing.Point(this.Size.Width / 2 - this.stage_1.Width / 2, this.labelTime.Size.Height + 30);
            this.stage_2.Size = new Size(this.Size.Width, this.Size.Height - (this.labelTime.Height));
            this.stage_2.Location = new System.Drawing.Point(0, this.labelTime.Height);
            this.stage_3.Location = new System.Drawing.Point(this.Width / 2 - this.stage_3.Width / 2, this.Height / 2 - this.stage_3.Height / 2);
            this.btnPrev.Location = new System.Drawing.Point(10, this.Height - (this.btnPrev.Height + 100));
            this.btnNext.Location = new System.Drawing.Point(this.Width - (this.btnNext.Size.Width + 30), this.Height - (this.btnNext.Height + 100));
            this.btnEnd.Location = new System.Drawing.Point(this.Width / 2 - this.btnEnd.Width / 2, this.Height - (this.btnEnd.Height + 100));
            this.pBar.Location = new System.Drawing.Point(this.Width / 2 - this.pBar.Width / 2, this.btnEnd.Location.Y - this.pBar.Height);
            this.lblQuestionText.MaximumSize = new Size(this.stage_2.Width - this.lblQuestion.Width - 20, 500);
            this.lblAnswerA.Location = new System.Drawing.Point(25, this.Height / 4);
            this.lblAnswerA.MaximumSize = new Size(this.Width / 3, this.Height / 3);
            this.lblAnswerB.Location = new System.Drawing.Point(25, this.Height * 2 / 4);
            this.lblAnswerB.MaximumSize = new Size(this.Width / 3, this.Height / 3);
            this.lblAnswerC.Location = new System.Drawing.Point(this.Width * 2 / 3, this.Height / 4);
            this.lblAnswerC.MaximumSize = new Size(this.Width / 3 - 25, this.Height / 3);
            this.lblAnswerD.Location = new System.Drawing.Point(this.Width * 2 / 3, this.Height * 2 / 4);
            this.lblAnswerD.MaximumSize = new Size(this.Width / 3 - 25, this.Height / 3);
            this.btnStart.Height = 50;
            this.cmbNames.Width = stage_1.Width - (lblPosts.Width + 20);
            this.cmbGroups.Width = stage_1.Width - (lblPosts.Width + 20);
            this.cmbPosts.Width = stage_1.Width - (lblPosts.Width + 20);
            this.cmbNames.Enabled = false;
            this.cmbGroups.Enabled = false;
            this.cmbPosts.Enabled = false;

            ExplorerManager.Kill();
            this.EnterStage_1();
        }

        // Fix the ws and range cells disparity
        private void Fill(string pathToTemplate, string saveAsPath,
            string protocolNumber,
            string date,
            string dateplus,
            string fullname,
            string name,
            string sur,
            string famil,
            string post,
            string markInPercent,
            string group)
        {
            
            object misValue = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = app.Workbooks.Open(pathToTemplate);
            Worksheet ws = (Worksheet)wb.Worksheets.get_Item(1);
            Range range = ws.UsedRange;
            this.setProgressBar(1, range.Rows.Count + range.Columns.Count);
            for (int i = 1; i < range.Rows.Count; i++)
            {
                for (int j = 1; j < range.Columns.Count; j++)
                {
                    string content = Convert.ToString((range.Cells[i, j] as Range).Value2);
                    if (content != null && content.Length != 0)
                    {
                        if (content.Contains("<protocol>")) content = content.Replace("<protocol>", protocolNumber);
                        if (content.Contains("<date>")) content = content.Replace("<date>", date);
                        if (content.Contains("<date+>")) content = content.Replace("<date+>", dateplus);
                        if (content.Contains("<fullname>")) content = content.Replace("<fullname>", name);
                        if (content.Contains("<name>")) content = content.Replace("<name>", name);
                        if (content.Contains("<sur>")) content = content.Replace("<sur>", sur);
                        if (content.Contains("<famil>")) content = content.Replace("<famil>",famil);
                        if (content.Contains("<post>")) content = content.Replace("<post>", post);
                        if (content.Contains("<mark>")) content = content.Replace("<mark>", markInPercent);
                        if (content.Contains("<group>")) content = content.Replace("<group>", group);
                        (range.Cells[i, j] as Range).Value2 = content;
                    }
                    this.pBar.PerformStep();
                }
            }
            wb.SaveAs(saveAsPath, XlFileFormat.xlOpenXMLWorkbook);
            wb.Close(true, misValue, misValue);
            app.Quit();
            ReleaseObject(ws);
            ReleaseObject(wb);
            ReleaseObject(app);

        }

        private void ReadDataStage_1()
        {
            this.btnNextTest.Text = "Зареждане...";
            object misValue = System.Reflection.Missing.Value;
            List<String> allNames = new List<string>();
            List<string> allPosts = new List<string>();

            Microsoft.Office.Interop.Excel.Application namesAndPostsApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook namesList = namesAndPostsApp.Workbooks.Open(this.mainPath + "Имена.xlsx");
            Worksheet namesWorkSheet = (Worksheet)namesList.Worksheets.get_Item(1);
            Range namesRange = namesWorkSheet.UsedRange;

            for (int i = 1; i <= namesRange.Cells.Count / 2; i++)
            {
                string name = Convert.ToString((namesRange.Cells[i, 1] as Range).Value2);
                string egn = Convert.ToString((namesRange.Cells[i, 2] as Range).Value2);
                if (name != null && egn != null)
                {
                    allNames.Add(name);
                    this.allEgn.Add(egn);
                }
            }
            // release objects
            namesList.Close(true, misValue, misValue);
            namesAndPostsApp.Quit();
            ReleaseObject(namesWorkSheet);
            ReleaseObject(namesList);
            ReleaseObject(namesAndPostsApp);

            Microsoft.Office.Interop.Excel.Application postsApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook postsList = postsApp.Workbooks.Open(this.mainPath + "Длъжности.xlsx");
            Worksheet postWorkSheet = (Worksheet)postsList.Worksheets.get_Item(1);
            Range postsRange = postWorkSheet.UsedRange;

            for (int i = 1; i <= postsRange.Rows.Count; i++)
            {
                string post = Convert.ToString((postsRange.Cells[i, 1] as Range).Value2);
                string yearAsStr = Convert.ToString((postsRange.Cells[i, 2] as Range).Value2);
                if (post != null) allPosts.Add(post);
                if(yearAsStr != null)
                {
                    int year = int.Parse(yearAsStr);
                    this.deltaYear.Add(year);
                }
            }
            postsList.Close(true, misValue, misValue);
            postsApp.Quit();
            ReleaseObject(postsList);
            ReleaseObject(postWorkSheet);
            ReleaseObject(postsApp);
            this.cmbNames.Items.AddRange(allNames.ToArray());
            this.cmbPosts.Items.AddRange(allPosts.ToArray());
        }

        private void ReadDataStage_2()
        {
            string selectedItem = cmbGroups.SelectedItem.ToString();
            string path = this.GetPathToGroup(selectedItem);
            char selectedItemNumber = selectedItem[selectedItem.Length - 1];
            HashSet<int> asked = new HashSet<int>();
            Random rnd = new Random();

            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Workbook wBook = app.Workbooks.Open(path);
            Worksheet wSheet = (Worksheet)wBook.Worksheets.get_Item(1);
            Range range = wSheet.UsedRange;
            // 
            int index = rnd.Next(2, range.Rows.Count);

            for (int i = 1; i <= this.questionsForGroup[selectedItemNumber]; i++)
            {
                while (asked.Contains(index))
                {
                    index = rnd.Next(2, range.Rows.Count);
                }
                asked.Add(index);
                string currentQuestion = Convert.ToString((range.Cells[index, 1] as Range).Value2);
                string currentAnswer = Convert.ToString((range.Cells[index, 2] as Range).Value2);
                string currentReference = Convert.ToString((range.Cells[index, 3] as Range).Value2);
                if (currentAnswer != null && currentQuestion != null && currentAnswer.Length != 0 && currentQuestion.Length != 0)
                {
                    Question quest = new Question();
                    string[] questionAndItsAnswes = currentQuestion.Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    quest.Questiontext = questionAndItsAnswes[0];
                    quest.AnswerA = questionAndItsAnswes[1];
                    quest.AnswerB = questionAndItsAnswes[2];
                    quest.AnswerC = questionAndItsAnswes[3];
                    quest.AnswerD = questionAndItsAnswes[4];
                    if (currentReference != null && currentReference != "")
                    {
                        quest.ForReference = currentReference;
                    }
                    quest.StudentsAnswer = -1;
                    quest.RightAnswer = int.Parse(currentAnswer.Trim());
                    this.questions.Add(quest);
                }
                else
                {
                    i--;
                }
            }
            //
            wBook.Close(0);
            app.Quit();
            ReleaseObject(wBook);
            ReleaseObject(wSheet);
            ReleaseObject(app);
        }

        private void cmb_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbPosts.SelectedIndex > -1
                && cmbGroups.SelectedIndex > -1
                && cmbNames.SelectedIndex > -1)
            {
                this.btnStart.Enabled = true;
                this.btnStart.Visible = true;
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            this.EnterStage_2();
        }

        private void setQuestion(int index)
        {
            this.UncolorAnswerLabels();
            lblQuestionText.Text = this.questions[index].Questiontext;
            lblAnswerA.Text = this.questions[index].AnswerA;
            lblAnswerB.Text = this.questions[index].AnswerB;
            lblAnswerC.Text = this.questions[index].AnswerC;
            lblAnswerD.Text = this.questions[index].AnswerD;
            if (this.questions[index].StudentsAnswer != -1)
            {
                switch (this.questions[index].StudentsAnswer)
                {
                    case 1:
                        this.lblAnswerA.BackColor = Color.DodgerBlue;
                        break;
                    case 2:
                        this.lblAnswerB.BackColor = Color.DodgerBlue;
                        break;
                    case 3:
                        this.lblAnswerC.BackColor = Color.DodgerBlue;
                        break;
                    case 4:
                        this.lblAnswerD.BackColor = Color.DodgerBlue;
                        break;
                    default: break;
                }
            }
        }

        private string GetPathToGroup(string group)
        {
            switch (this.cmbGroups.SelectedItem.ToString())
            {
                case "изпит за квалификационна група по ПБЗРЕУ  група 2":
                    return this.mainPath + @"Тестове по групи\\Втора Група.xlsx";
                case "изпит за квалификационна група по ПБЗРЕУ  група 3":
                    return this.mainPath + @"Тестове по групи\\Трета Група.xlsx";
                case "изпит за квалификационна група по ПБЗРЕУ  група 4":
                    return this.mainPath + @"Тестове по групи\\Четвърта Група.xlsx";
                case "изпит за квалификационна група по ПБЗРЕУ  група 5":
                    return this.mainPath + "Тестове по групи\\Пета Група.xlsx";
                case "изпит за квалификационна група по ПБЗРНЕУ  група 2":
                    return this.mainPath + "Тестове по групи\\Втора Група НеЕл.xlsx";
                case "изпит за квалификационна група по ПБЗРНЕУ  група 3":
                    return this.mainPath + "Тестове по групи\\Трета Група НеЕл.xlsx";
                case "изпит за квалификационна група по ПБЗРНЕУ  група 4":
                    return this.mainPath + "Тестове по групи\\Четвърта Група НеЕл.xlsx";
                case "изпит за квалификационна група по ПБЗРНЕУ  група 5":
                    return this.mainPath + "Тестове по групи\\Пета Група НеЕл.xlsx";
                case "изпит по наредба 9":
                    return this.mainPath + "Тестове по групи\\наредба 9\\" + this.cmbPosts.SelectedItem.ToString() + ".xlsx";
                default:
                    return "";
            }
        }

        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void Time_Tick(object sender, EventArgs e)
        {

            this.sec--;
            if (this.sec < 0)
            {
                this.sec = 59;
                this.min--;
                if (this.min < 0)
                {
                    MessageBox.Show("Вашето време изтече");
                    this.Close();
                }
            }
            this.labelTime.Text = string.Format("Оставащо време: {0}:{1}", this.min, this.sec);
        }

        private void CheckIfHasToHideBtn()
        {
            if(this.questionIndex == 0)
            {
                this.btnPrev.Enabled = false;
                this.btnPrev.Visible = false;

                this.btnNext.Enabled = true;
                this.btnNext.Visible = true;
            }
            else if(this.questionIndex == this.questions.Count - 1)
            {
                this.btnPrev.Enabled = true;
                this.btnPrev.Visible = true;

                this.btnNext.Enabled = false;
                this.btnNext.Visible = false;

                this.btnEnd.Visible = true;
                this.btnEnd.Enabled = true;
            }
            else
            {
                this.btnEnd.Visible = false;
                this.btnEnd.Enabled = false;
                this.btnPrev.Enabled = true;
                this.btnPrev.Visible = true;
                this.btnNext.Enabled = true;
                this.btnNext.Visible = true;
            }
        }

        private void btnPrev_Click(object sender, EventArgs e)
        {
            this.questionIndex--;
            this.setQuestion(this.questionIndex);
            this.CheckIfHasToHideBtn();
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            this.questionIndex++;
            this.setQuestion(this.questionIndex);
            this.CheckIfHasToHideBtn();
        }

        private void UncolorAnswerLabels()
        {
            this.lblAnswerA.BackColor = Color.Transparent;
            this.lblAnswerB.BackColor = Color.Transparent;
            this.lblAnswerC.BackColor = Color.Transparent;
            this.lblAnswerD.BackColor = Color.Transparent;
        }

        private void lblAnswerA_Click(object sender, EventArgs e)
        {
            this.UncolorAnswerLabels();
            this.lblAnswerA.BackColor = Color.DodgerBlue;
            this.questions[this.questionIndex].StudentsAnswer = 1;
        }

        private void lblAnswerB_Click(object sender, EventArgs e)
        {
            this.UncolorAnswerLabels();
            this.lblAnswerB.BackColor = Color.DodgerBlue;
            this.questions[this.questionIndex].StudentsAnswer = 2;
        }

        private void lblAnswerC_Click(object sender, EventArgs e)
        {
            this.UncolorAnswerLabels();
            this.lblAnswerC.BackColor = Color.DodgerBlue;
            this.questions[this.questionIndex].StudentsAnswer = 3;
        }

        private void lblAnswerD_Click(object sender, EventArgs e)
        {
            this.UncolorAnswerLabels();
            this.lblAnswerD.BackColor = Color.DodgerBlue;
            this.questions[this.questionIndex].StudentsAnswer = 4;
        }

        //
        // Fills the templates and starts stage_3
        private void btnEnd_Click(object sender, EventArgs e)
        {
            this.btnEnd.Enabled = false;
            DialogResult dialogResult = MessageBox.Show("Сигурни ли сте че искате да приключите теста?",
                "Приключи теста",
                MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                // contains only questions with right student's answer
                var rightAnswers =
                    from quest in this.questions
                    where quest.StudentsAnswer == quest.RightAnswer
                    select quest;

                int rightAnswersCount = rightAnswers.Count(); // number of right answers
                string[] selectedItem = cmbGroups.SelectedItem.ToString().Split();
                char group = selectedItem[selectedItem.Length - 1][0]; // student's group as char
                int mark = ((rightAnswersCount * 100) / this.questionsForGroup[group]); // students mark
                int protocol = this.getProtocolNumber(); // number of protocol
                string name = this.cmbNames.SelectedItem.ToString(); // student's name
                string post = this.cmbPosts.SelectedItem.ToString(); // student's post
                string date = DateTime.Now.ToShortDateString();
                string dateplus = DateTime.Now.AddYears(1).ToShortDateString();
                string path = this.GetTemplatePath(mark, name); // Path to main template for current group and post
                string certificatePath = this.getCertificatePath(); // path to certificate for current group
                bool passed = mark >= this.calcScoreNeeded(); // If the student passed the exam
                string groupAsString = "" + group;
                string[] nameSplitted = name.Split(new char[] { ' ' });
                string saveAsPath = this.mainPath + @"Генерирани Документи\" + name;
                saveAsPath += passed ? "_Издържал_" : "_Неиздържал_";
                saveAsPath += this.getGroupTypeString(this.cmbGroups.SelectedItem.ToString()) + ".xlsx";
                this.Fill(this.mainPath + @"Темплейти\" + path, saveAsPath,
                    protocol.ToString(),
                    date, 
                    dateplus,
                    name, 
                    nameSplitted[0],
                    nameSplitted[1],
                    nameSplitted[2], 
                    post, 
                    mark.ToString(), 
                    groupAsString);
                if(passed)
                {
                    string saveCertificatePath = this.mainPath + @"Генерирани Документи\Удостоверения\" + 
                        name +
                        "_" +
                        this.getGroupTypeString(this.cmbGroups.SelectedItem.ToString()) + "_Удостоверение.xlsx;";

                    this.Fill(this.mainPath + @"Темплейти\" + certificatePath, saveCertificatePath,
                        protocol.ToString(),
                        date, 
                        dateplus,
                        name,nameSplitted[0], 
                        nameSplitted[1],
                        nameSplitted[2],
                        post,
                        mark.ToString(),
                        groupAsString);
                    string failedDoc = this.GetFailedDocument();
                    this.RemoveFromFailedDocument(failedDoc, name);
                    this.RemoveFromFailedDocument("Повторно " + failedDoc, name);
                }
                else
                {
                    this.PutInFailedDocument(name);
                }
                this.PutInAreToBeExamined();
                this.GeneratePrivateDocuments(name, mark, protocol);
                this.RemoveCurrentNameFromList(name);
                this.pBar.Visible = false;
                this.pBar.Enabled = false;
                this.EnterStage_3(mark, passed);
                Time.Stop();
            }
            else
            {
                return;
            }
        }
        // --------------------------------
        private void RemoveFromFailedDocument(string path, string nameToRemove)
        {
            List<string> names = new List<string>();
            using (StreamReader reader = new StreamReader(string.Format(this.mainPath + path)))
            {
                while (true)
                {
                    string currentName = reader.ReadLine();
                    if (currentName == null) break;
                    if (currentName != nameToRemove)
                        names.Add(reader.ReadLine());
                }

            }
            File.WriteAllText(this.mainPath + path, String.Empty);
            using (StreamWriter w = new StreamWriter(string.Format(this.mainPath + path)))
            {
                foreach (var name in names)
                {
                    w.WriteLine(name);
                }
            }
        }

        private void RemoveCurrentNameFromList(string name)
        {
            // Create a new Excel document ------------------------------------------------------------
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlWorkBook = xlApp.Workbooks.Open(this.mainPath + "Имена.xlsx");
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            this.setProgressBar(1, this.cmbNames.Items.Count / 2 - 1);
            for (int i = 0; i < this.cmbNames.Items.Count; i++)
            {
                if ((string)this.cmbNames.Items[i] == name)
                {
                    this.cmbNames.Items.RemoveAt(i);
                    this.allEgn.RemoveAt(i);
                    break;
                }
                pBar.PerformStep();
            }

            int index;
            for (index = 1; index <= this.cmbNames.Items.Count; index++)
            {
                (xlWorkSheet.Cells[index, 1] as Range).Value2 = this.cmbNames.Items[index - 1];
                (xlWorkSheet.Cells[index, 2] as Range).Value2 = this.allEgn[index - 1];
                pBar.PerformStep();
            }
            (xlWorkSheet.Cells[index, 1] as Range).Value2 = "";
            (xlWorkSheet.Cells[index, 2] as Range).Value2 = "";

            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            ReleaseObject(xlWorkSheet);
            ReleaseObject(xlWorkBook);
            ReleaseObject(xlApp);
        }

        private void PutInAreToBeExamined()
        {
            using (StreamWriter w = new StreamWriter(this.mainPath + "Предстоят да се явят.txt", true))
            {
                int year = cmbPosts.SelectedIndex == 8 ? DateTime.Now.Year + this.deltaYear[this.cmbPosts.SelectedIndex] : 1;
                w.WriteLine(string.Format("{0} - {1}.{2}.{3}", cmbNames.SelectedItem.ToString(),
                    DateTime.Now.Day,
                    DateTime.Now.Month,
                    year));
            }
        }

        private void EnterStage_1()
        {
            this.cmbNames.Items.Clear();
            this.cmbPosts.Items.Clear();
            this.cmbGroups.SelectedIndex = -1;
            this.stage_3.Enabled = false;
            this.stage_3.Visible = false;
            this.stage_1.Enabled = true;
            this.stage_1.Visible = true;
            this.stage_1.BringToFront();
            this.btnStart.Enabled = false;
            this.btnStart.Visible = false;
            this.ReadDataStage_1();
            this.cmbNames.Enabled = true;
            this.cmbGroups.Enabled = true;
            this.cmbPosts.Enabled = true;
        }

        private void EnterStage_2()
        {
            this.stage_1.Enabled = false;
            this.stage_1.Visible = false;
            this.stage_2.Enabled = true;
            this.stage_2.Visible = true;
            this.stage_2.BringToFront();
            this.cmbNames.Enabled = false;
            this.cmbGroups.Enabled = false;
            this.cmbPosts.Enabled = false;
            this.questions.Clear();
            this.questionIndex = 0;
            this.btnPrev.Enabled = false;
            this.btnPrev.Visible = false;
            this.btnNext.Enabled = true;
            this.btnNext.Visible = true;
            this.btnEnd.Visible = false;
            this.btnEnd.Enabled = false;
            this.lblAnswerA.Text = "Отговор А";
            this.lblAnswerB.Text = "Отговор Б";
            this.lblAnswerC.Text = "Отговор В";
            this.lblAnswerD.Text = "Отговор Г";
            this.lblQuestionText.Text = "...";
            this.ReadDataStage_2();
            this.setQuestion(this.questionIndex);
            this.sec = 0;
            this.min = 30;
            this.Time.Start();
        }

        private void EnterStage_3(int mark, bool passed)
        {
            this.stage_2.Enabled = false;
            this.stage_2.Visible = false;
            this.stage_3.Enabled = true;
            this.stage_3.Visible = true;
            this.stage_3.BringToFront();
            this.btnNextTest.Text = "Нов тест";
            this.YesNoLabel.Text = passed ? "Да" : "Не";
            this.YesNoLabel.BackColor = passed ? Color.Green : Color.Red;
        }

        private void PutInFailedDocument(string name)
        {
            string path = this.GetFailedDocument();

            if (this.StudentHasAlreadyFailed(name)) path = "Повторно " + path;

            using (StreamWriter writer = new StreamWriter(this.mainPath + @"\" + path, true))
            {
                writer.WriteLine(name);
            }
        }

        private string GetFailedDocument()
        {
            string path = "Неиздържали";
            switch (this.getGroupType(this.cmbGroups.SelectedItem.ToString()))
            {
                case 0: path += " Наредба 9.txt"; break;
                case 1: path += " Ел.txt"; break;
                case 2: path += " НеЕл.txt"; break;
            }
            return path;
        }

        private void GeneratePrivateDocuments(string name, int mark, int protocol)
        {
            setProgressBar(1, this.questions.Count);
            using (StreamWriter privateWriter = new StreamWriter(string.Format(this.mainPath + @"Генерирани Документи\_Anatoliy\Отговори на {0}.txt", name)))
            using (StreamWriter publicWriter = new StreamWriter(string.Format(this.mainPath + @"Генерирани Документи\_Отговори\Отговори на {0}.txt", name)))
            {
                int indexOfQuestion = 1;
                publicWriter.WriteLine(string.Format("Дата: {0}-{1}-{2}", DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year));
                publicWriter.WriteLine("Номер на протокола: " + protocol);
                publicWriter.WriteLine("Явил се на " + cmbGroups.SelectedItem.ToString());
                publicWriter.WriteLine("Отговорите на: " + name);
                publicWriter.WriteLine("Резултат: " + mark + "%");
                publicWriter.WriteLine("\n");

                privateWriter.WriteLine(string.Format("Дата: {0}-{1}-{2}", DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year));
                privateWriter.WriteLine("Номер на протокола: " + protocol);
                privateWriter.WriteLine("Явил се на " + cmbGroups.SelectedItem.ToString());
                privateWriter.WriteLine("Отговорите на: " + name);
                privateWriter.WriteLine("Резултат: " + mark + "%");
                privateWriter.WriteLine("\n");

                foreach (var quest in this.questions)
                {
                    privateWriter.WriteLine(indexOfQuestion + ". " + quest.Questiontext);
                    privateWriter.WriteLine(quest.AnswerA);
                    privateWriter.WriteLine(quest.AnswerB);
                    privateWriter.WriteLine(quest.AnswerC);
                    privateWriter.WriteLine(quest.AnswerD);
                    privateWriter.WriteLine("Верен отговор: " + quest.RightAnswer);
                    privateWriter.WriteLine("Даден отговор: " + quest.StudentsAnswer);
                    bool isWrong = quest.RightAnswer != quest.StudentsAnswer;
                    if (isWrong)
                    {
                        privateWriter.WriteLine("За справка: " + quest.ForReference);
                    }
                    privateWriter.WriteLine("-------------------------------------------------------");

                    publicWriter.WriteLine(indexOfQuestion++ + ". " + quest.Questiontext);
                    publicWriter.WriteLine(quest.AnswerA);
                    publicWriter.WriteLine(quest.AnswerB);
                    publicWriter.WriteLine(quest.AnswerC);
                    publicWriter.WriteLine(quest.AnswerD);
                    publicWriter.WriteLine("Даден отговор: " + quest.StudentsAnswer);
                    if (isWrong)
                    {
                        publicWriter.WriteLine("**Грешен**");
                    }
                    publicWriter.WriteLine("-------------------------------------------------------");
                    this.pBar.PerformStep();
                }
            }
        }

        private string getCertificatePath()
        {
            if (this.cmbGroups.SelectedIndex < 4)
            {
                return "certificate_neel.xlsx";
            }
            if (this.cmbGroups.SelectedIndex < 8)
            {
                return "certificate_el.xlsx";
            }
            return "certificate_9.xlsx";
        }

        private string GetTemplatePath(int mark, string name)
        {
            int groupType = this.getGroupType(this.cmbGroups.SelectedItem.ToString());
            bool hasAlreadyFailed = this.StudentHasAlreadyFailed(name);
            int scoreNeeded = this.calcScoreNeeded();

            string path = "Template";
            if (mark >= scoreNeeded)
            {
                path += "Passed";
            }
            else
            {
                path += "NotPassed";
                if (hasAlreadyFailed)
                {
                    path += "SecondTime";
                }
            }
            switch (groupType)
            {
                case 0: path += "_Ordinance_9"; break;
                case 1: path += "El"; break;
                case 2: break;
                default: break;
            }
            path += ".xlsx";
            return path;

        }

        private bool StudentHasAlreadyFailed(string name)
        {
            using (StreamReader r = new StreamReader(this.mainPath + "Неиздържали.txt"))
            {
                while (true)
                {
                    string readName = r.ReadLine();
                    if (readName == null) return false;
                    if (readName == name)
                    {
                        return true;
                    }
                }
            }
        }

        private int calcScoreNeeded()
        {
            if(this.cmbGroups.SelectedIndex == 8) // Ordinance 9
            {
                return 75;
            }
            else
            {
                return 80;
            }
        }

        private int getProtocolNumber()
        {
            int protocolNo;
            string group = this.cmbGroups.SelectedItem.ToString();
            int typeOfGroup = getGroupType(group); // 0 -> наредба 9; 1 -> Ел; 2 -> НеЕл;
            using (StreamReader reader = new StreamReader(this.mainPath + "ПротоколНомер.txt"))
            {
                int[] protocolNumber = new int[3];
                protocolNumber[0] = int.Parse(reader.ReadLine().Trim());
                protocolNumber[1] = int.Parse(reader.ReadLine().Trim());
                protocolNumber[2] = int.Parse(reader.ReadLine().Trim());
                protocolNo = protocolNumber[typeOfGroup];
                reader.Close();
                using (StreamWriter writer = new StreamWriter(this.mainPath + "ПротоколНомер.txt"))
                {
                    protocolNumber[typeOfGroup]++;
                    writer.WriteLine(protocolNumber[0]);
                    writer.WriteLine(protocolNumber[1]);
                    writer.WriteLine(protocolNumber[2]);
                }
            }
            return protocolNo;
        }

        private int getGroupType(string group)
        {
            string[] type = group.Split(' ');
            if (type.Length == 4) return 0; // наредба 9
            if (type[5] == "ПБЗРЕУ") return 1; // ел група
            if (type[5] == "ПБЗРНЕУ") return 2; // не ел група
            throw new System.Exception("Group of test is not supported");

        }
        private string getGroupTypeString(string group)
        {
            string[] type = group.Split(' ');
            if (type.Length == 4) return "Наредба 9"; // наредба 9
            return type[5];
            throw new System.Exception("Group of test is not supported");
        }

        private void labelTime_Click(object sender, EventArgs e)
        {
            var passForm = new PasswordForm();
            passForm.Show();
        }

        private void btnNextTest_Click(object sender, EventArgs e)
        {
            this.EnterStage_1();
        }

        private void setProgressBar(int min, int max)
        {
            this.pBar.Visible = true;
            this.pBar.Enabled = true;
            this.pBar.Minimum = min;
            this.pBar.Maximum = max;
            this.pBar.Value = 1;
            this.pBar.Step = 1;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == System.Windows.Forms.CloseReason.UserClosing)
            {
                e.Cancel = true;
            }
        }
    }
}