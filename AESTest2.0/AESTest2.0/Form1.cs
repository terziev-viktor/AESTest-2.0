using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace AESTest2._0
{
    public partial class MainForm : Form
    {
        private const string MAINPATH = @"C:\data\";
        private const string DATAEXAMS = @"данни_тестове.txt";
        private const string DATASTUDENTS = @"данни_имена.txt";
        private const string DATAPOSTS = @"данни_длъжности.txt";
        private const string GENERATEDDOCS = @"Генерирани Документи\";
        private const string TEMPLATESDOCS = @"Темплейти\";
        private const string TEMPLATESFAILED = @"Скъсани\";
        private const string TEMPLATESFAILEDAGAIN = @"Скъсани Втори Път\";
        private const string TEMPLATESPASSED = @"Преминали\";
        private const string QUESTIONSDOCS = @"Въпросници\";
        private const string TEMPLATESCERTIFICATES = @"Удостоверения\";
        private const string HIDDENDOCS = @"Скрити\";
        private const string PUBLICDOCS = @"Публични\";
        private const string FAILEDDOCS = @"Неиздържали\";
        private const string FAILEDAGAINDOCS = @"Повторно Неиздържали\";
        private const string PROTOCOLDOC = @"Номер на протокола.txt";
        private const string DEFAULTPROTOCOLNUMBER = "100";
        private bool SaveDataToDataSheets = true;
        private string[] templateStrings =
        {
            "<protocol>",
            "<date>",
            "<dateplus>",
            "<fullname>",
            "<name>",
            "<sur>",
            "<famil>",
            "<post>",
            "<mark>",
            "<group>",
            "<pin>",
        };

        private DataHolder dataHolder = new DataHolder();

        private int sec = 0;
        private int min = 30;
        private int questionIndex = 0;
        private const int WS_SYSMENU = 0x80000;

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
            this.YesNoLabel.Enabled = false;
            this.btnStart.Enabled = false;

            this.labelTime.Location = new System.Drawing.Point(this.Width / 2 - this.labelTime.Width / 2, 0);
            this.stageManager.Size = new Size(this.Size.Width, this.Size.Height - (this.labelTime.Height));
            this.stageManager.Location = new System.Drawing.Point(0, this.labelTime.Height);

            this.stage_1.Size = new Size(this.Size.Width * 2 / 3, 700);
            this.rtbAbout.Width = this.stage_1.Width - 10;
            this.rtbAbout.Height = 200;
            this.stage_1.Location = new System.Drawing.Point(this.Size.Width / 2 - this.stage_1.Width / 2, this.labelTime.Size.Height + 30);
            this.btnOpenManagerPanel.Location = new System.Drawing.Point(this.Width - this.btnOpenManagerPanel.Width, this.btnOpenManagerPanel.Height);
            this.stage_2.Size = new Size(this.Size.Width, this.Size.Height - (this.labelTime.Height));
            this.stage_2.Location = new System.Drawing.Point(0, this.labelTime.Height);
            this.stage_3.Location = new System.Drawing.Point(this.Width / 2 - this.stage_3.Width / 2, this.Height / 2 - this.stage_3.Height / 2);
            this.btnPrev.Location = new System.Drawing.Point(10, this.Height - (this.btnPrev.Height + 100));
            this.btnNext.Location = new System.Drawing.Point(this.Width - (this.btnNext.Size.Width + 30), this.Height - (this.btnNext.Height + 100));
            this.btnQuestIndex.Location = new System.Drawing.Point(this.Width / 2 - this.btnQuestIndex.Width / 2, this.Height - (this.btnQuestIndex.Height + 100));
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
            if (!Directory.Exists(MAINPATH))
            {
                // the program is opened for first time. Setup the main directory with all files that should be there
                MessageBox.Show("Вие отворихте програмата за първи път. Файлове с данни ще бъдат генерирани в папка C:/data/. Моля попълнете ги.");
                this.InitSetup();
                this.SaveDataToDataSheets = false;
                System.Windows.Forms.Application.Exit();
            }
            // Kills explorer.exe
            ExplorerManager.Kill();
            this.EnterStage_1();
        }

        private void InitSetup()
        {
            Directory.CreateDirectory(MAINPATH);
            Directory.CreateDirectory(MAINPATH + TEMPLATESDOCS);
            Directory.CreateDirectory(MAINPATH + TEMPLATESDOCS + TEMPLATESPASSED);
            Directory.CreateDirectory(MAINPATH + TEMPLATESDOCS + TEMPLATESFAILED);
            Directory.CreateDirectory(MAINPATH + TEMPLATESDOCS + TEMPLATESFAILED + TEMPLATESFAILEDAGAIN);
            Directory.CreateDirectory(MAINPATH + TEMPLATESDOCS + TEMPLATESCERTIFICATES);
            Directory.CreateDirectory(MAINPATH + QUESTIONSDOCS);
            Directory.CreateDirectory(MAINPATH + GENERATEDDOCS);
            Directory.CreateDirectory(MAINPATH + GENERATEDDOCS + PUBLICDOCS);
            Directory.CreateDirectory(MAINPATH + GENERATEDDOCS + TEMPLATESCERTIFICATES);
            DirectoryInfo dir = Directory.CreateDirectory(MAINPATH + GENERATEDDOCS + HIDDENDOCS);
            dir.Attributes = FileAttributes.Hidden;
            Directory.CreateDirectory(MAINPATH + GENERATEDDOCS + HIDDENDOCS + FAILEDDOCS);
            Directory.CreateDirectory(MAINPATH + GENERATEDDOCS + HIDDENDOCS + FAILEDAGAINDOCS);

            object misValue = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            if (excelApp == null)
            {
                MessageBox.Show("MS Excel не е инсталиран правилно или е стара версия (< 2008).");
                System.Windows.Forms.Application.Exit();
            }
            
            File.Create(MAINPATH + DATAEXAMS).Close();
            File.Create(MAINPATH + DATAPOSTS).Close();
            File.Create(MAINPATH + DATASTUDENTS).Close();

            using (var w = new StreamWriter(MAINPATH + DATAEXAMS))
            {
                w.WriteLine(".");
                w.WriteLine("Изтрийте всичко и попълнете в следния формат:");
                w.WriteLine("<вид изпит> <брой въпроси> <нужен резултата за оценка 'ДА'> <нов ред>");
            }

            using (var w = new StreamWriter(MAINPATH + DATAPOSTS))
            {
                w.WriteLine(".");
                w.WriteLine("Изтрийте всичко и попълнете в следния формат:");
                w.WriteLine("<длъжност> <на колко години да се явява> <нов ред>");
            }
            using (var w = new StreamWriter(MAINPATH + DATASTUDENTS))
            {
                w.WriteLine(".");
                w.WriteLine("Изтрийте всичко и попълнете в следния формат:");
                w.WriteLine("<3 имена> <ЕГН> <нов ред>");
            }

            // file with then number of protocol
            using (var writer = File.CreateText(MAINPATH + PROTOCOLDOC))
            {
                writer.WriteLine(DEFAULTPROTOCOLNUMBER);
            }
        }

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
            string group,
            string egn)
        {
            
            object misValue = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(pathToTemplate);
            this.setProgressBar(1, doc.Words.Count);
            string[] correspondingStrings = { protocolNumber, date, dateplus, fullname, name, sur, famil, post, markInPercent, group, egn};
            for (int i = 0; i < templateStrings.Length; i++)
            {
                Microsoft.Office.Interop.Word.Find findObject = wordApp.Selection.Find;
                findObject.ClearFormatting();
                findObject.Text = templateStrings[i];
                findObject.Replacement.ClearFormatting();
                findObject.Replacement.Text = correspondingStrings[i];

                object replaceAll = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;

                findObject.Execute(ref misValue, ref misValue, ref misValue, ref misValue, ref misValue,
                    ref misValue, ref misValue, ref misValue, ref misValue, ref misValue,
                    ref replaceAll, ref misValue, ref misValue, ref misValue, ref misValue);
            }

            doc.SaveAs2(saveAsPath);
            doc.Close(true);
            ReleaseObject(doc);
            wordApp.Quit();
            ReleaseObject(wordApp);
        }

        private void ReadDataStage_1()
        {
            this.btnNextTest.Text = "Зареждане...";
            this.cmbGroups.Items.Clear();
            this.cmbNames.Items.Clear();
            this.cmbPosts.Items.Clear();
            this.dataHolder.Questions.Clear();
            object misValue = System.Reflection.Missing.Value;

            using (var r = new StreamReader(MAINPATH + DATASTUDENTS))
            {
                string line = r.ReadLine();
                while(line != null && line != "")
                {
                    string[] sp = line.Split(new char[] { ' ' });
                    string egn = sp[sp.Length - 1];
                    string name = string.Join(" ", sp.Take(sp.Length - 1));
                    dataHolder.Students.Add(new Student()
                    {
                        Fullname = name,
                        PIN = egn
                    });
                    cmbNames.Items.Add(name);
                    line = r.ReadLine();
                }
            }

            using (var r = new StreamReader(MAINPATH + DATAEXAMS))
            {
                string line = r.ReadLine();
                while(line != null && line != "")
                {
                    string[] sp = line.Split(new char[] { ' ' });
                    string minScore = sp[sp.Length - 1];
                    string num = sp[sp.Length - 2];
                    string exam = string.Join(" ", sp.Take(sp.Length - 2));
                    dataHolder.Exams.Add(new Exam()
                    {
                        MinScore = int.Parse(minScore),
                        QuestionsCount = int.Parse(num),
                        Title = exam
                    });
                    cmbGroups.Items.Add(exam);
                    line = r.ReadLine();
                }
            }

            using (var r = new StreamReader(MAINPATH + DATAPOSTS))
            {
                string line = r.ReadLine();
                while(line != null && line != "")
                {
                    string[] sp = line.Split(new char[] { ' ' });
                    string deltaYear = sp[sp.Length - 1];
                    string post = string.Join(" ", sp.Take(sp.Length - 1));
                    cmbPosts.Items.Add(post);
                    this.dataHolder.Posts.Add(new Post()
                    {
                        DeltaYear = int.Parse(deltaYear),
                        Title = post
                    });
                    line = r.ReadLine();
                }
            }
            
        }

        private void ReadDataStage_2()
        {
            string selectedItem = cmbGroups.SelectedItem.ToString();
            Exam exam = this.dataHolder.Exams.Where(x => x.Title == selectedItem).ElementAt(0);
            
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook wBook = excelApp.Workbooks.Open(MAINPATH + QUESTIONSDOCS + selectedItem + ".xlsx");
            Worksheet wSheet = (Worksheet)wBook.Worksheets.get_Item(1);
            Microsoft.Office.Interop.Excel.Range range = wSheet.UsedRange;
            
            for (int i = 2; i <= range.Rows.Count; i++)
            {
                string currentQuestion = Convert.ToString((range.Cells[i, 1] as Microsoft.Office.Interop.Excel.Range).Value2);
                string currentAnswer = Convert.ToString((range.Cells[i, 2] as Microsoft.Office.Interop.Excel.Range).Value2);
                string currentReference = Convert.ToString((range.Cells[i, 3] as Microsoft.Office.Interop.Excel.Range).Value2);
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
                    this.dataHolder.Questions.Add(quest);
                }
            }
            Random rnd = new Random();
            this.dataHolder.Questions.OrderBy(x => rnd.Next());
            //
            wBook.Close(false);
            ReleaseObject(wSheet);
            ReleaseObject(wBook);
            excelApp.Quit();
            ReleaseObject(excelApp);
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
            lblQuestionText.Text = this.dataHolder.Questions[index].Questiontext;
            lblAnswerA.Text = this.dataHolder.Questions[index].AnswerA;
            lblAnswerB.Text = this.dataHolder.Questions[index].AnswerB;
            lblAnswerC.Text = this.dataHolder.Questions[index].AnswerC;
            lblAnswerD.Text = this.dataHolder.Questions[index].AnswerD;
            if (this.dataHolder.Questions[index].StudentsAnswer != -1)
            {
                switch (this.dataHolder.Questions[index].StudentsAnswer)
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
                    System.Windows.Forms.Application.Exit();
                }
            }
            this.labelTime.Text = string.Format("Оставащо време: {0}:{1}", this.min, this.sec);
        }

        private void CheckIfHasToHideBtn(int numberOfQuestionsForCurrentExam)
        {
            if (this.questionIndex == 0)
            {
                this.btnPrev.Enabled = false;
                this.btnPrev.Visible = false;

                this.btnNext.Enabled = true;
                this.btnNext.Visible = true;
            }
            else if (this.questionIndex == numberOfQuestionsForCurrentExam - 1)
            {
                this.btnPrev.Enabled = true;
                this.btnPrev.Visible = true;

                this.btnNext.Enabled = false;
                this.btnNext.Visible = false;

                this.btnEnd.Visible = true;
                this.btnEnd.Enabled = true;

                this.btnQuestIndex.Visible = false;
            }
            else
            {
                this.btnEnd.Visible = false;
                this.btnEnd.Enabled = false;
                this.btnQuestIndex.Visible = true;
                this.btnPrev.Enabled = true;
                this.btnPrev.Visible = true;
                this.btnNext.Enabled = true;
                this.btnNext.Visible = true;
            }
        }

        private void btnPrev_Click(object sender, EventArgs e)
        {
            int count = this.dataHolder.Exams[this.dataHolder.CurrentExamIndex].QuestionsCount;
            this.questionIndex--;
            this.btnQuestIndex.Text = (this.questionIndex + 1) + "/" + count;
            this.setQuestion(this.questionIndex);
            this.CheckIfHasToHideBtn(count);
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            int count = this.dataHolder.Exams[this.dataHolder.CurrentExamIndex].QuestionsCount;

            this.questionIndex++;
            this.btnQuestIndex.Text = (this.questionIndex + 1) + "/" + count;
            this.setQuestion(this.questionIndex);
            this.CheckIfHasToHideBtn(count);
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
            this.dataHolder.Questions[this.questionIndex].StudentsAnswer = 1;
        }

        private void lblAnswerB_Click(object sender, EventArgs e)
        {
            this.UncolorAnswerLabels();
            this.lblAnswerB.BackColor = Color.DodgerBlue;
            this.dataHolder.Questions[this.questionIndex].StudentsAnswer = 2;
        }

        private void lblAnswerC_Click(object sender, EventArgs e)
        {
            this.UncolorAnswerLabels();
            this.lblAnswerC.BackColor = Color.DodgerBlue;
            this.dataHolder.Questions[this.questionIndex].StudentsAnswer = 3;
        }

        private void lblAnswerD_Click(object sender, EventArgs e)
        {
            this.UncolorAnswerLabels();
            this.lblAnswerD.BackColor = Color.DodgerBlue;
            this.dataHolder.Questions[this.questionIndex].StudentsAnswer = 4;
        }

        // Fills the templates and enters stage_3
        private void btnEnd_Click(object sender, EventArgs e)
        {
            this.btnEnd.Enabled = false;
            DialogResult dialogResult = MessageBox.Show("Сигурни ли сте че искате да приключите теста?",
                                                        "Приключи теста",
                                                         MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                // contains only questions with right student's answer

                string selectedExamType = cmbGroups.SelectedItem.ToString();
                Exam exam = this.dataHolder.Exams.Where(x => x.Title == selectedExamType).ElementAt(0);
                Student student = this.dataHolder.Students.Where(x => x.Fullname == this.cmbNames.SelectedItem.ToString()).ElementAt(0);
                Post post = this.dataHolder.Posts.Where(x => x.Title == cmbPosts.SelectedItem.ToString()).ElementAt(0);

                int rightAnswersCount = dataHolder.Questions.Take(exam.QuestionsCount).Where(x => x.RightAnswer == x.StudentsAnswer).Count(); // number of right answers
                
                int mark = ((rightAnswersCount * 100) / exam.QuestionsCount); // students mark
                int protocol = this.getProtocolNumber(); // number of protocol
                string date = DateTime.Now.ToShortDateString();
                string dateplus = DateTime.Now.AddYears(post.DeltaYear).ToShortDateString();
                string path = this.GetTemplatePath(mark, student.Fullname, exam); // Path to template for current exam
                string certificatePath = MAINPATH + TEMPLATESDOCS + TEMPLATESCERTIFICATES + exam.Title + ".docx"; // path to certificate for current group
                bool passed = mark >= exam.MinScore; // If the student passed the exam
                string passedStr = passed ? "Издържал" : "Скъсан";
                string[] nameSplitted = student.Fullname.Split(new char[] { ' ' });
                string saveAsPath = MAINPATH + GENERATEDDOCS + student.Fullname;
                saveAsPath += "_" + passedStr + "_";
                saveAsPath += exam.Title + ".docx";
                this.Fill(path, saveAsPath,
                    protocol.ToString(),
                    date,
                    dateplus,
                    student.Fullname,
                    nameSplitted[0],
                    nameSplitted[1],
                    nameSplitted[2],
                    post.Title,
                    mark.ToString(),
                    selectedExamType,
                    student.PIN);
                if (passed)
                {
                    string saveCertificatePath = MAINPATH + GENERATEDDOCS + TEMPLATESCERTIFICATES +
                        student.Fullname +
                        "_" +
                        exam.Title + ".docx";

                    this.Fill(certificatePath, saveCertificatePath,
                        protocol.ToString(),
                        date,
                        dateplus,
                        student.Fullname,
                        nameSplitted[0],
                        nameSplitted[1],
                        nameSplitted[2],
                        post.Title,
                        mark.ToString(),
                        selectedExamType,
                        student.PIN);
                    this.RemoveFromFailedDocument(exam.Title, student.Fullname);
                }
                else
                {
                    this.PutInFailedDocument(student.Fullname, exam.Title);
                }
                this.PutInAreToBeExamined(post, student);
                this.GeneratePrivateDocuments(student.Fullname, mark, protocol, passedStr, exam);
                this.dataHolder.Students.RemoveAt(dataHolder.Students.FindIndex(x => x.Fullname == student.Fullname));
                cmbNames.Items.Clear();
                cmbNames.Items.AddRange(dataHolder.Students.Select(x => x.Fullname).ToArray());
                this.pBar.Visible = false;
                this.pBar.Enabled = false;
                this.EnterStage_3(mark, passed);
                Time.Stop();
            }
        }

        private void WriteDataToDataSheets()
        {

            File.WriteAllText(MAINPATH + DATASTUDENTS, String.Empty);
            using (var w = new StreamWriter(MAINPATH + DATASTUDENTS))
            {
                for (int i = 0; i < dataHolder.Students.Count; i++)
                {
                    w.WriteLine(dataHolder.Students[i].Fullname + " " + dataHolder.Students[i].PIN);
                }
            }

            File.WriteAllText(MAINPATH + DATAEXAMS, String.Empty);
            using (var w = new StreamWriter(MAINPATH + DATAEXAMS))
            {
                for (int i = 0; i < dataHolder.Exams.Count; i++)
                {
                    w.WriteLine(dataHolder.Exams[i].Title + " " + dataHolder.Exams[i].QuestionsCount + " " + dataHolder.Exams[i].MinScore);
                }
            }

            File.WriteAllText(MAINPATH + DATAPOSTS, String.Empty);
            using (var w = new StreamWriter(MAINPATH + DATAPOSTS))
            {
                for (int i = 0; i < dataHolder.Posts.Count; i++)
                {
                    w.WriteLine(dataHolder.Posts[i].Title + " " + dataHolder.Posts[i].DeltaYear);
                }
            }
        }

        // --------------------------------
        private void RemoveFromFailedDocument(string examtype, string nameToRemove)
        {
            string path1 = MAINPATH + GENERATEDDOCS + HIDDENDOCS + FAILEDDOCS + examtype + ".txt";
            string path2 = MAINPATH + GENERATEDDOCS + HIDDENDOCS + FAILEDAGAINDOCS + examtype + ".txt";
            RemoveNameFromDoc(path1, nameToRemove);
            RemoveNameFromDoc(path2, nameToRemove);
        }

        private void RemoveNameFromDoc(string path, string nameToRemove)
        {
            if(!File.Exists(path))
            {
                File.Create(path).Close();
            }
            List<string> names = new List<string>();
            using (StreamReader reader = new StreamReader(path))
            {
                while (true)
                {
                    string currentName = reader.ReadLine();
                    if (currentName == null) break;
                    if (currentName != nameToRemove)
                        names.Add(reader.ReadLine());
                }

            }
            File.WriteAllText(path, String.Empty);
            using (StreamWriter w = new StreamWriter(path))
            {
                foreach (var name in names)
                {
                    w.WriteLine(name);
                }
            }
        }
        
        private void PutInAreToBeExamined(Post post, Student student)
        {
            using (StreamWriter w = new StreamWriter(MAINPATH + "Предстоят да се явят.txt", true))
            {
                int year = post.DeltaYear;
                w.WriteLine(string.Format("{0} - {1}.{2}.{3}", student.Fullname,
                    DateTime.Now.Day,
                    DateTime.Now.Month,
                    year));
            }
        }

        private void EnterStage_1()
        {
            this.cmbGroups.SelectedIndex = -1;
            
            this.stageManager.Enabled = false;
            this.stageManager.Visible = false;
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
            this.dataHolder.Questions.Clear();
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
            this.dataHolder.CurrentExamIndex = this.dataHolder.Exams.IndexOf(this.dataHolder.Exams.Where(x => x.Title == this.cmbGroups.SelectedItem.ToString()).First());
            this.btnQuestIndex.Text = (this.questionIndex + 1) + "/" + this.dataHolder.Exams.Where(x => x.Title == cmbGroups.SelectedItem.ToString()).ElementAt(0).QuestionsCount;
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

        private void PutInFailedDocument(string name, string examtype)
        {
            string path = MAINPATH + GENERATEDDOCS + HIDDENDOCS;

            if (this.StudentHasAlreadyFailed(name, examtype))
            {
                path += FAILEDAGAINDOCS + examtype + ".txt";
            }
            else
            {
                path += FAILEDDOCS + examtype + ".txt";
            }
            if(!File.Exists(path))
            {
                File.Create(path).Close();
            }
            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine(name);
            }
        }

        private void GeneratePrivateDocuments(string name, int mark, int protocol, string passed, Exam exam)
        {
            setProgressBar(1, exam.QuestionsCount);
            using (StreamWriter privateWriter = new StreamWriter(string.Format(MAINPATH + GENERATEDDOCS + HIDDENDOCS + @"Отговори {0}_{1}_{2}.txt", name, passed, exam.Title)))
            using (StreamWriter publicWriter = new StreamWriter(string.Format(MAINPATH + GENERATEDDOCS + PUBLICDOCS + @"Отговори {0}_{1}_{2}.txt", name, passed, exam.Title)))
            {
                int indexOfQuestion = 1;
                publicWriter.WriteLine(string.Format("Дата: {0}-{1}-{2}", DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year));
                publicWriter.WriteLine("Номер на протокола: " + protocol);
                publicWriter.WriteLine("Явил се на " + exam.Title);
                publicWriter.WriteLine("Отговорите на: " + name);
                publicWriter.WriteLine("Резултат: " + mark + "%");
                publicWriter.WriteLine("\n");

                privateWriter.WriteLine(string.Format("Дата: {0}-{1}-{2}", DateTime.Now.Day, DateTime.Now.Month, DateTime.Now.Year));
                privateWriter.WriteLine("Номер на протокола: " + protocol);
                privateWriter.WriteLine("Явил се на " + cmbGroups.SelectedItem.ToString());
                privateWriter.WriteLine("Отговорите на: " + name);
                privateWriter.WriteLine("Резултат: " + mark + "%");
                privateWriter.WriteLine("\n");

                for (int i = 0; i < exam.QuestionsCount; i++)
                {
                    privateWriter.WriteLine(indexOfQuestion + ". " + dataHolder.Questions[i].Questiontext);
                    privateWriter.WriteLine(dataHolder.Questions[i].AnswerA);
                    privateWriter.WriteLine(dataHolder.Questions[i].AnswerB);
                    privateWriter.WriteLine(dataHolder.Questions[i].AnswerC);
                    privateWriter.WriteLine(dataHolder.Questions[i].AnswerD);
                    privateWriter.WriteLine("Верен отговор: " + dataHolder.Questions[i].RightAnswer);
                    privateWriter.WriteLine("Даден отговор: " + dataHolder.Questions[i].StudentsAnswer);
                    bool isWrong = dataHolder.Questions[i].RightAnswer != dataHolder.Questions[i].StudentsAnswer;
                    if (isWrong)
                    {
                        privateWriter.Write("** Грешен **");
                        privateWriter.WriteLine("За справка: " + dataHolder.Questions[i].ForReference);
                    }
                    privateWriter.WriteLine("-------------------------------------------------------");

                    publicWriter.WriteLine(indexOfQuestion++ + ". " + dataHolder.Questions[i].Questiontext);
                    publicWriter.WriteLine(dataHolder.Questions[i].AnswerA);
                    publicWriter.WriteLine(dataHolder.Questions[i].AnswerB);
                    publicWriter.WriteLine(dataHolder.Questions[i].AnswerC);
                    publicWriter.WriteLine(dataHolder.Questions[i].AnswerD);
                    publicWriter.WriteLine("Даден отговор: " + dataHolder.Questions[i].StudentsAnswer);
                    if (isWrong)
                    {
                        publicWriter.WriteLine("**Грешен**");
                    }
                    publicWriter.WriteLine("-------------------------------------------------------");
                    this.pBar.PerformStep();
                }
            }
        }

        private string GetTemplatePath(int mark, string name, Exam exam)
        {
            bool hasAlreadyFailed = this.StudentHasAlreadyFailed(name, exam.Title);
            // TODO: calculate the score needed to pass the exam
            int scoreNeeded = exam.MinScore;

            string path = MAINPATH + TEMPLATESDOCS;
            if (mark >= scoreNeeded)
            {
                path += TEMPLATESPASSED;
            }
            else
            {
                path += TEMPLATESFAILED;
                if (hasAlreadyFailed)
                {
                    path += TEMPLATESFAILEDAGAIN;
                }
            }
            path += exam.Title;
            path += ".docx";
            return path;

        }

        private bool StudentHasAlreadyFailed(string name, string examtype)
        {
            if (!File.Exists(MAINPATH + GENERATEDDOCS + HIDDENDOCS + FAILEDDOCS + examtype + ".txt"))
            {
                File.Create(MAINPATH + GENERATEDDOCS + HIDDENDOCS + FAILEDDOCS + examtype + ".txt").Close();
                return false;
            }
            using (StreamReader r = new StreamReader(MAINPATH + GENERATEDDOCS + HIDDENDOCS + FAILEDDOCS + examtype + ".txt"))
            {
                string readName;
                do
                {
                    readName = r.ReadLine();
                    if (readName == name)
                    {
                        return true;
                    }
                } while (readName != null);
                return false;
            }
        }
        
        private int getProtocolNumber()
        {
            int protocolNo;
            using (StreamReader reader = new StreamReader(MAINPATH + PROTOCOLDOC))
            {
                bool success = int.TryParse(reader.ReadLine(), out protocolNo);
                if (!success)
                {
                    MessageBox.Show("Файлът с номера на протокола е бил променен. Ще бъде използван номер по подразбиране.");
                    return 100;
                }
            }
            using (StreamWriter writer = new StreamWriter(MAINPATH + PROTOCOLDOC))
            {
                writer.Write(protocolNo + 1);
            }
            return protocolNo;
        }

        private string getGroupTypeString(string group)
        {
            string[] type = group.Split(' ');
            if (type.Length == 4) return "Наредба_9"; // наредба 9
            return type[5];
            throw new Exception("Group of test is not supported");
        }

        private void labelTime_Click(object sender, EventArgs e)
        {
            var passForm = new PasswordForm();
            passForm.Show();
        }

        private void btnNextTest_Click(object sender, EventArgs e)
        {
            this.cmbGroups.SelectedIndex = -1;
            this.stageManager.Enabled = false;
            this.stageManager.Visible = false;
            this.stage_3.Enabled = false;
            this.stage_3.Visible = false;
            this.stage_1.Enabled = true;
            this.stage_1.Visible = true;
            this.stage_1.BringToFront();
            this.btnStart.Enabled = false;
            this.btnStart.Visible = false;
            this.cmbNames.Enabled = true;
            this.cmbGroups.Enabled = true;
            this.cmbPosts.Enabled = true;

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
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
            }
            if(this.SaveDataToDataSheets)
            {
                WriteDataToDataSheets();
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnOpenManagerPanel_Click(object sender, EventArgs e)
        {
            this.enterManagerStage();
        }

        private void enterManagerStage()
        {
            this.stage_1.Enabled = false;
            this.stage_1.Visible = false;
            this.stageManager.Enabled = true;
            this.stageManager.Visible = true;
            this.lbExamTypes.Items.Clear();
            this.readDataStageManager();
        }

        private void readDataStageManager()
        {
            foreach (var item in this.dataHolder.Exams)
            {
                this.lbExamTypes.Items.Add(item.Title + " -> " + item.QuestionsCount);
            }
        }

        private void btnAddExamType_Click(object sender, EventArgs e)
        {
            string newExamType = this.tbAddExamType.Text;
            string qcount = this.tbAddExamTypeQCount.Text;
            string minscoreStr = this.tbAddExamTypeMinScore.Text;

            if (this.cmbGroups.Items.Contains(newExamType))
            {
                MessageBox.Show("Вече има такъв изпит.");
                return;
            }
            if (newExamType == null || qcount == null)
            {
                MessageBox.Show("Попълнете полетата.");
                return;
            }
            int count;
            int minscore;
            bool isCountInt = int.TryParse(qcount, out count);
            bool isMinScoreInt = int.TryParse(minscoreStr, out minscore);
            if(!isCountInt)
            {
                MessageBox.Show("Невалиден брой въпроси.");
                return;
            }
            if(!isMinScoreInt)
            {
                MessageBox.Show("Невалиден минимален резултат за оценка 'Да'.");
                return;
            }
            Exam newExam = new Exam()
            {
                Title = newExamType,
                QuestionsCount = count,
                MinScore = minscore
            };
            this.dataHolder.Exams.Add(newExam);
            this.lbExamTypes.Items.Add(newExam.Title + " -> " + newExam.QuestionsCount);
            this.cmbGroups.Items.Add(newExam.Title);
        }

        private void lbExamTypes_SelectedIndexChanged(object sender, EventArgs e)
        {
            int index = this.lbExamTypes.SelectedIndex;
            if (index < 0)
            {
                //this.btnRemoveExamType.Enabled = false;
                //this.btnRemoveExamType.Visible = false;
                this.btnAddQuestionsFile.Enabled = false;
                this.btnAddQuestionsFile.Visible = false;
                this.btnAddPassedTemplate.Enabled = false;
                this.btnAddPassedTemplate.Visible = false;
                this.btnAddFailedTemplate.Enabled = false;
                this.btnAddFailedTemplate.Visible = false;
                this.btnAddTemplateFailedSecondTime.Visible = false;
                this.btnAddTemplateFailedSecondTime.Enabled = false;
                this.btnAddCertificateTemplate.Enabled = false;
                this.btnAddCertificateTemplate.Visible = false;
            }
            else
            {
                //this.btnRemoveExamType.Enabled = true;
                //this.btnRemoveExamType.Visible = true;
                this.btnAddQuestionsFile.Enabled = true;
                this.btnAddQuestionsFile.Visible = true;
                this.btnAddPassedTemplate.Enabled = true;
                this.btnAddPassedTemplate.Visible = true;
                this.btnAddFailedTemplate.Enabled = true;
                this.btnAddTemplateFailedSecondTime.Visible = true;
                this.btnAddTemplateFailedSecondTime.Enabled = true;
                this.btnAddFailedTemplate.Visible = true;
                this.btnAddCertificateTemplate.Enabled = true;
                this.btnAddCertificateTemplate.Visible = true;
            }
        }

        private void btnRemoveExamType_Click(object sender, EventArgs e)
        {
            // TODO: selected item has -> <number of questions> appended to it.
            this.dataHolder.Exams.Remove(this.dataHolder.Exams.Where(x => x.Title == lbExamTypes.SelectedItem.ToString()).First());
            this.cmbGroups.Items.Remove(lbExamTypes.SelectedItem.ToString());
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.CloseManagerStage();
        }

        private void CloseManagerStage()
        {
            this.stageManager.Visible = false;
            this.stageManager.Enabled = false;

            this.stage_1.Visible = true;
            this.stage_1.Enabled = true;
        }

        private void btnAddQuestionsFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel Files|*.xlsx;";
            dialog.Title = "Please select a Excel 2007+ document for the template.";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string item = this.lbExamTypes.SelectedItem.ToString();
                int n = item.IndexOf(" -> ");
                item = item.Remove(n);
                File.Copy(dialog.FileName, MAINPATH + @"Въпросници\" + item + ".xlsx", true);
            }
        }

        private void btnAddPassedTemplate_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Word Files|*.docx;";
            dialog.Title = "Please select a Word 2007+ document for the template.";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string item = this.lbExamTypes.SelectedItem.ToString();
                int n = item.IndexOf(" -> ");
                item = item.Remove(n);
                File.Copy(dialog.FileName, MAINPATH + "Темплейти\\Преминали\\" + item + ".docx", true);
            }
        }

        private void btnAddFailedTemplate_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Word Files|*.docx;";
            dialog.Title = "Please select a Word 2007+ document for the template.";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string item = this.lbExamTypes.SelectedItem.ToString();
                int n = item.IndexOf(" -> ");
                item = item.Remove(n);
                File.Copy(dialog.FileName, MAINPATH + TEMPLATESDOCS + TEMPLATESFAILED + item + ".docx", true);
            }
        }

        private void btnAddCertificateTemplate_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Word Files|*.docx;";
            dialog.Title = "Please select a Word 2007+ document for the template.";

            string item = this.lbExamTypes.SelectedItem.ToString();
            int n = item.IndexOf(" -> ");
            item = item.Remove(n);
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                File.Copy(dialog.FileName, MAINPATH + "Темплейти\\Удостоверения\\" + item + ".docx", true);
            }
        }

        private void btnAddTemplateFailedSecondTime_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Word Files|*.docx;";
            dialog.Title = "Please select a Word 2007+ document for the template.";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string item = this.lbExamTypes.SelectedItem.ToString();
                int n = item.IndexOf(" -> ");
                item = item.Remove(n);
                File.Copy(dialog.FileName, MAINPATH + TEMPLATESDOCS + TEMPLATESFAILED + TEMPLATESFAILEDAGAIN + item + ".docx", true);
            }
        }

        private void btnHelp_Click(object sender, EventArgs e)
        {
            // Open Help Panel
        }
    }
}