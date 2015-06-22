namespace QDBGenerator
{
    using Microsoft.Win32;
    using NPOI.HSSF.UserModel;
    using NPOI.SS.UserModel;
    using NPOI.XSSF.UserModel;
    using System;
    using System.CodeDom.Compiler;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Diagnostics;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Controls.Primitives;
    using System.Windows.Input;
    using System.Windows.Markup;
    using System.Windows.Media;

    public class MainWindow : Window, IComponentConnector, IStyleConnector
    {
        private bool _contentLoaded;
        internal Button btnCombine;
        internal Button btnGenerate;
        internal Button btnOrigin;
        private string combineDirectoryName = string.Empty;
        private List<Course> courseList = null;
        private Dictionary<string, Course> courseNameNodeDic = null;
        private Dictionary<string, HashSet<string>> courseNodeDic = null;
        private bool isFileError = false;
        internal ItemsControl itemsControl;
        private int maxCount = 30;
        private Dictionary<string, ToggleButton> originCheckDic = null;
        private string replaceStr = "※";
        internal ScrollViewer scrList;
        private Dictionary<string, HashSet<string>> specialityNodeDic = null;
        private SortedDictionary<string, List<Student>> stuOriginDic = null;
        private string targetRootPath = string.Empty;
        private string templateFilePath = string.Empty;
        internal TextBox txtArrange;
        internal TextBox txtStudent;
        internal TextBox txtTemplate;

        public MainWindow()
        {
            this.InitializeComponent();
        }

        private void btnCombine_Click(object sender, RoutedEventArgs e)
        {
            if ((this.stuOriginDic == null) || (this.stuOriginDic.Count == 0))
            {
                Note("请放入所需文件，并初始化！");
            }
            else if ((this.originCheckDic == null) || (this.originCheckDic.Count == 0))
            {
                Note("请放入所需文件，并选中需要合并的来源！");
            }
            else
            {
                DateTime now = DateTime.Now;
                List<Student> stuList = new List<Student>();
                foreach (string str in this.originCheckDic.Keys)
                {
                    stuList.AddRange(this.stuOriginDic[str]);
                    this.stuOriginDic.Remove(str);
                }
                stuList.Sort(new Comparison<Student>(MainWindow.CompareStudent));
                this.combineDirectoryName = string.Join("+", this.originCheckDic.Keys);
                try
                {
                    this.GenerateFromList(stuList);
                    this.ForbidCombinedCheckBox();
                    DateTime time2 = DateTime.Now;
                    TimeSpan span = (TimeSpan) (time2 - now);
                    MessageBox.Show(string.Format("开始时间：{0}\r\n结束时间：{1}\r\n总耗时：{2:f1} 秒", now, time2, span.TotalSeconds));
                }
                catch (Exception exception)
                {
                    Note(exception.Message);
                }
                finally
                {
                    this.originCheckDic.Clear();
                    stuList.Clear();
                    stuList.TrimExcess();
                    stuList = null;
                    this.combineDirectoryName = string.Empty;
                }
            }
        }

        private void btnGenerate_Click(object sender, RoutedEventArgs e)
        {
            if ((this.stuOriginDic == null) || (this.stuOriginDic.Count == 0))
            {
                Note("请放入所需文件，并初始化！");
            }
            else
            {
                DateTime now = DateTime.Now;
                try
                {
                    List<string> list = this.stuOriginDic.Keys.ToList<string>();
                    for (int i = 0; i < list.Count; i++)
                    {
                        this.stuOriginDic[list[i]].Sort(new Comparison<Student>(MainWindow.CompareStudent));
                        this.GenerateFromList(this.stuOriginDic[list[i]]);
                        this.stuOriginDic.Remove(list[i]);
                    }
                }
                catch (Exception exception)
                {
                    Note(exception.ToString());
                }
                this.ForbidLeftCheckBox();
                DateTime time2 = DateTime.Now;
                TimeSpan span = (TimeSpan) (time2 - now);
                MessageBox.Show(string.Format("开始时间：{0}\r\n结束时间：{1}\r\n总耗时：{2:f1} 秒", now, time2, span.TotalSeconds));
                this.CollectGarbage();
            }
        }

        private void btnInit_Click(object sender, RoutedEventArgs e)
        {
            if ((this.stuOriginDic != null) && (this.stuOriginDic.Count > 0))
            {
                this.CollectGarbage();
            }
            this.CheckFileFormat(new TextBox[] { this.txtStudent, this.txtArrange, this.txtTemplate });
            if (this.isFileError)
            {
                Note("请拖入Excel格式文件！");
            }
            else
            {
                FileStream stream = null;
                try
                {
                    this.InitContainer();
                    this.InitData(this.txtStudent.Text.Trim(), this.txtArrange.Text.Trim(), this.txtTemplate.Text.Trim());
                    this.itemsControl.ItemsSource = this.stuOriginDic;
                }
                catch (Exception exception)
                {
                    Note(exception.ToString());
                }
                finally
                {
                    if (stream != null)
                    {
                        stream.Dispose();
                    }
                }
            }
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            ToggleButton button = sender as ToggleButton;
            this.originCheckDic[button.Content.ToString()] = button;
        }

        private void CheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            ToggleButton button = sender as ToggleButton;
            this.originCheckDic.Remove(button.Content.ToString());
        }

        private void CheckFileFormat(params TextBox[] tbs)
        {
            foreach (TextBox box in tbs)
            {
                if (!(string.IsNullOrEmpty(box.Text) || (!Path.GetExtension(box.Text).Equals(".xls") && !Path.GetExtension(box.Text).Equals(".xlsx"))))
                {
                    this.isFileError = false;
                    box.BorderBrush = Brushes.Black;
                    box.BorderThickness = new Thickness(0.5);
                }
                else
                {
                    this.isFileError = true;
                    box.BorderBrush = Brushes.Red;
                    box.BorderThickness = new Thickness(2.0);
                }
            }
        }

        private void ClassifyByCourse(List<Student> stuList)
        {
            Dictionary<string, List<Student>> dictionary = new Dictionary<string, List<Student>>();
            HashSet<string> set = new HashSet<string>();
            foreach (Student student in stuList)
            {
                foreach (string str in this.specialityNodeDic[student.Speciality])
                {
                    if (!set.Contains(str))
                    {
                        set.Add(str);
                    }
                }
            }
            foreach (string str2 in set)
            {
                foreach (Student student in stuList)
                {
                    if (this.courseNodeDic[str2].Contains(student.Speciality))
                    {
                        if (!dictionary.ContainsKey(str2))
                        {
                            dictionary.Add(str2, new List<Student>());
                        }
                        dictionary[str2].Add(student);
                    }
                }
            }
            foreach (KeyValuePair<string, List<Student>> pair in dictionary)
            {
                this.GeneratQDG(this.courseNameNodeDic[pair.Key], pair.Value);
            }
            dictionary.Clear();
            dictionary = null;
            set.Clear();
            set.TrimExcess();
            set = null;
        }

        private void CollectGarbage()
        {
            this.stuOriginDic.Clear();
            this.courseList.Clear();
            this.courseList.TrimExcess();
            this.courseNodeDic.Clear();
            this.specialityNodeDic.Clear();
            this.courseNameNodeDic.Clear();
            GC.Collect();
        }

        private static int CompareStudent(Student s1, Student s2)
        {
            if (s1 == null)
            {
                if (s2 == null)
                {
                    return 0;
                }
                return -1;
            }
            if (s2 == null)
            {
                return 1;
            }
            if (s1.Grade.CompareTo(s2.Grade) == 0)
            {
                if (s1.Layer.CompareTo(s2.Layer) == 0)
                {
                    if (s1.Speciality.CompareTo(s2.Speciality) == 0)
                    {
                        return s1.Id.CompareTo(s2.Id);
                    }
                    return s1.Speciality.CompareTo(s2.Speciality);
                }
                return s1.Layer.CompareTo(s2.Layer);
            }
            return s1.Grade.CompareTo(s2.Grade);
        }

        private void ForbidCombinedCheckBox()
        {
            foreach (ToggleButton button in this.originCheckDic.Values)
            {
                button.Background = Brushes.Gray;
                button.IsEnabled = false;
            }
        }

        private void ForbidLeftCheckBox()
        {
            List<ToggleButton> childObjects = GetChildObjects<ToggleButton>(this.itemsControl, "");
            foreach (ToggleButton button in childObjects)
            {
                if (button.IsEnabled)
                {
                    button.Background = Brushes.Silver;
                    button.IsEnabled = false;
                }
            }
        }

        private void GenerateFromList(List<Student> stuList)
        {
            List<Student> range = null;
            int index = 0;
            int count = 0;
            if ((stuList != null) && (stuList.Count > 0))
            {
                count = 1;
                for (int i = 1; i < stuList.Count; i++)
                {
                    if (stuList[i].Grade.Equals(stuList[i - 1].Grade) && stuList[i].Layer.Equals(stuList[i - 1].Layer))
                    {
                        count++;
                    }
                    else
                    {
                        range = stuList.GetRange(index, count);
                        this.ClassifyByCourse(range);
                        index = i;
                        count = 1;
                    }
                }
                range = stuList.GetRange(index, count);
                this.ClassifyByCourse(range);
            }
        }

        private void GeneratQDG(Course course, List<Student> stuList)
        {
            Student student = stuList[0];
            string path = string.Empty;
            if (string.IsNullOrEmpty(this.combineDirectoryName))
            {
                path = Path.Combine(this.targetRootPath, student.Origin, student.Grade, student.Layer);
            }
            else
            {
                path = Path.Combine(this.targetRootPath, this.combineDirectoryName, student.Grade, student.Layer);
            }
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string courseName = course.CourseName;
            if (courseName.Contains("*"))
            {
                courseName = courseName.Replace("*", this.replaceStr);
            }
            List<Student> list2 = null;
            int count = 0;
            char ch = '1';
            bool flag = false;
            list2 = stuList;
            while (true)
            {
                count = (list2.Count <= this.maxCount) ? list2.Count : this.maxCount;
                List<Student> range = list2.GetRange(0, count);
                list2.RemoveRange(0, count);
                string filePath = string.Empty;
                if (!flag)
                {
                    filePath = Path.Combine(path, string.Format("{0}{1}", courseName, ".xls"));
                }
                else
                {
                    object[] args = new object[4];
                    args[0] = courseName;
                    args[1] = "___";
                    ch = (char) (ch + '\x0001');
                    args[2] = ch.ToString();
                    args[3] = ".xls";
                    filePath = Path.Combine(path, string.Format("{0}{1}{2}{3}", args));
                }
                this.SaveToExcelFile(filePath, course.CourseName, course.Time, range);
                count = (list2.Count <= this.maxCount) ? list2.Count : this.maxCount;
                if (count > 0)
                {
                    flag = true;
                }
                else
                {
                    return;
                }
            }
        }

        private void GennerateCourseNode()
        {
            foreach (Course course in this.courseList)
            {
                this.courseNameNodeDic[course.CourseName] = course;
                if (this.courseNodeDic.ContainsKey(course.CourseName))
                {
                    if (!this.courseNodeDic[course.CourseName].Contains(course.Speciality))
                    {
                        this.courseNodeDic[course.CourseName].Add(course.Speciality);
                    }
                }
                else
                {
                    this.courseNodeDic.Add(course.CourseName, new HashSet<string>());
                    this.courseNodeDic[course.CourseName].Add(course.Speciality);
                }
            }
        }

        private void GennerateSpecialityNode()
        {
            foreach (Course course in this.courseList)
            {
                if (this.specialityNodeDic.ContainsKey(course.Speciality))
                {
                    if (!this.specialityNodeDic[course.Speciality].Contains(course.CourseName))
                    {
                        this.specialityNodeDic[course.Speciality].Add(course.CourseName);
                    }
                }
                else
                {
                    this.specialityNodeDic.Add(course.Speciality, new HashSet<string>());
                    this.specialityNodeDic[course.Speciality].Add(course.CourseName);
                }
            }
        }

        private static List<T> GetChildObjects<T>(DependencyObject obj, string name) where T: FrameworkElement
        {
            DependencyObject child = null;
            List<T> list = new List<T>();
            for (int i = 0; i <= (VisualTreeHelper.GetChildrenCount(obj) - 1); i++)
            {
                child = VisualTreeHelper.GetChild(obj, i);
                if ((child is T) && ((((T) child).Name == name) || string.IsNullOrEmpty(name)))
                {
                    list.Add((T) child);
                }
                list.AddRange(GetChildObjects<T>(child, ""));
            }
            return list;
        }

        private void GetCourseFromExcel(string filePath)
        {
            int num2;
            IWorkbook fileWorkbook = GetFileWorkbook(filePath);
            ISheet sheetAt = null;
            IRow row = null;
            if ((fileWorkbook == null) || (fileWorkbook.NumberOfSheets <= 0))
            {
                throw new Exception("考试安排表文件错误！");
            }
            sheetAt = fileWorkbook.GetSheetAt(0);
            IRow row2 = sheetAt.GetRow(1);
            int lastCellNum = row2.LastCellNum;
            Dictionary<string, int> dictionary = new Dictionary<string, int>();
            for (num2 = row2.FirstCellNum; num2 < lastCellNum; num2++)
            {
                if (!((row2.GetCell(num2) == null) | string.IsNullOrEmpty(row2.GetCell(num2).StringCellValue.Trim())))
                {
                    dictionary[row2.GetCell(num2).StringCellValue] = num2;
                }
            }
            this.courseList = new List<Course>();
            int lastRowNum = sheetAt.LastRowNum;
            StringBuilder builder = new StringBuilder();
            for (num2 = sheetAt.FirstRowNum + 2; num2 <= lastRowNum; num2++)
            {
                row = sheetAt.GetRow(num2);
                if (row != null)
                {
                    Course item = new Course {
                        Speciality = (row.GetCell(dictionary["专业名称"]) == null) ? "" : row.GetCell(dictionary["专业名称"]).StringCellValue,
                        CourseName = (row.GetCell(dictionary["课程名称"]) == null) ? "" : row.GetCell(dictionary["课程名称"]).StringCellValue,
                        Property = (row.GetCell(dictionary["考试性质"]) == null) ? "" : row.GetCell(dictionary["考试性质"]).StringCellValue
                    };
                    builder.Append((row.GetCell(dictionary["考试时间"]) == null) ? "" : row.GetCell(dictionary["考试时间"]).StringCellValue);
                    if (!string.IsNullOrEmpty(builder.ToString()))
                    {
                        item.Time = builder.ToString().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)[0];
                        item.Time = Regex.Replace(item.Time, "(.+?)-(.+?)-(.+)", "$1年$2月$3日");
                    }
                    this.courseList.Add(item);
                    builder.Clear();
                }
            }
            dictionary.Clear();
            dictionary = null;
        }

        private static IWorkbook GetFileWorkbook(string filePath)
        {
            FileStream stream = null;
            try
            {
                using (stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    char ch = filePath[filePath.Length - 1];
                    if (ch.Equals('x'))
                    {
                        return new XSSFWorkbook(stream);
                    }
                    return new HSSFWorkbook(stream);
                }
            }
            catch (Exception exception)
            {
                throw new IOException(string.Format("文件{0}无法打开，请检查该文件是否被其他程序锁定！\r\n{1}", filePath, exception.ToString()));
            }
            finally
            {
                if (stream != null)
                {
                    stream.Dispose();
                }
            }
        }

        private void GetStudentFromExcelToDic(string filePath)
        {
            int num2;
            IWorkbook fileWorkbook = GetFileWorkbook(filePath);
            ISheet sheetAt = null;
            IRow row = null;
            if ((fileWorkbook == null) || (fileWorkbook.NumberOfSheets <= 0))
            {
                throw new Exception("学籍表文件错误！");
            }
            sheetAt = fileWorkbook.GetSheetAt(0);
            IRow row2 = sheetAt.GetRow(0);
            int lastCellNum = row2.LastCellNum;
            Dictionary<string, int> dictionary = new Dictionary<string, int>();
            for (num2 = row2.FirstCellNum; num2 < lastCellNum; num2++)
            {
                if (!((row2.GetCell(num2) == null) | string.IsNullOrEmpty(row2.GetCell(num2).StringCellValue.Trim())))
                {
                    dictionary[row2.GetCell(num2).StringCellValue.Trim()] = num2;
                }
            }
            int lastRowNum = sheetAt.LastRowNum;
            for (num2 = sheetAt.FirstRowNum + 1; num2 <= lastRowNum; num2++)
            {
                row = sheetAt.GetRow(num2);
                if (row != null)
                {
                    Student item = new Student {
                        Id = (row.GetCell(dictionary["学号"]) == null) ? "" : row.GetCell(dictionary["学号"]).StringCellValue,
                        Name = (row.GetCell(dictionary["姓名"]) == null) ? "" : row.GetCell(dictionary["姓名"]).StringCellValue,
                        Origin = (row.GetCell(dictionary["来源"]) == null) ? "" : row.GetCell(dictionary["来源"]).StringCellValue,
                        Layer = (row.GetCell(dictionary["层次"]) == null) ? "" : row.GetCell(dictionary["层次"]).StringCellValue,
                        Speciality = (row.GetCell(dictionary["专业名称"]) == null) ? "" : row.GetCell(dictionary["专业名称"]).StringCellValue,
                        Grade = (row.GetCell(dictionary["年级"]) == null) ? "" : row.GetCell(dictionary["年级"]).StringCellValue
                    };
                    if (!this.stuOriginDic.ContainsKey(item.Origin))
                    {
                        this.stuOriginDic[item.Origin] = new List<Student>();
                    }
                    this.stuOriginDic[item.Origin].Add(item);
                }
            }
            dictionary.Clear();
            dictionary = null;
        }

        private IWorkbook GetTemplate(string filePath)
        {
            return GetFileWorkbook(filePath);
        }

        private void InitContainer()
        {
            this.stuOriginDic = new SortedDictionary<string, List<Student>>();
            this.courseList = new List<Course>();
            this.courseNodeDic = new Dictionary<string, HashSet<string>>();
            this.specialityNodeDic = new Dictionary<string, HashSet<string>>();
            this.courseNameNodeDic = new Dictionary<string, Course>();
            this.originCheckDic = new Dictionary<string, ToggleButton>();
        }

        private void InitData(string filePath, string schedulePath, string templatePath)
        {
            this.GetStudentFromExcelToDic(filePath);
            this.GetCourseFromExcel(schedulePath);
            this.GennerateCourseNode();
            this.GennerateSpecialityNode();
            this.targetRootPath = Path.Combine(Directory.GetParent(schedulePath).FullName, string.Format("签到表___{0}", DateTime.Now.ToString("yyyy年MM月dd日HH：mm：ss")));
            this.templateFilePath = templatePath;
        }

        [GeneratedCode("PresentationBuildTasks", "4.0.0.0"), DebuggerNonUserCode]
        public void InitializeComponent()
        {
            if (!this._contentLoaded)
            {
                this._contentLoaded = true;
                Uri resourceLocator = new Uri("/QDBGenerator;component/mainwindow.xaml", UriKind.Relative);
                Application.LoadComponent(this, resourceLocator);
            }
        }

        private static void Note(string msg)
        {
            MessageBox.Show(msg, "警告");
        }

        private void SaveToExcelFile(string filePath, string course, string examDate, List<Student> stuList)
        {
            IWorkbook template = this.GetTemplate(this.txtTemplate.Text);
            ISheet sheetAt = null;
            IRow row = null;
            string str = string.Empty;
            sheetAt = template.GetSheetAt(0);
            row = sheetAt.GetRow(2);
            str = row.Cells[0].StringCellValue.Trim();
            str = string.Format("{0} {1}", str, course);
            row.Cells[0].SetCellValue(str);
            str = row.Cells[5].StringCellValue.Trim().Substring(0, 5);
            str = string.Format("{0} {1}", str, examDate);
            row.Cells[5].SetCellValue(str);
            for (int i = 0; i < stuList.Count; i++)
            {
                row = sheetAt.GetRow(i + 4);
                int num2 = i + 1;
                row.Cells[0].SetCellValue(num2.ToString());
                row.Cells[1].SetCellValue(stuList[i].Speciality);
                row.Cells[2].SetCellValue(stuList[i].Layer);
                row.Cells[3].SetCellValue(stuList[i].Id);
                row.Cells[4].SetCellValue(stuList[i].Name);
            }
            this.WriteIO(filePath, template);
        }

        [EditorBrowsable(EditorBrowsableState.Never), DebuggerNonUserCode, GeneratedCode("PresentationBuildTasks", "4.0.0.0")]
        void IComponentConnector.Connect(int connectionId, object target)
        {
            switch (connectionId)
            {
                case 1:
                    ((MainWindow) target).KeyUp += new KeyEventHandler(this.Window_KeyUp);
                    return;

                case 2:
                    this.txtStudent = (TextBox) target;
                    this.txtStudent.PreviewDragEnter += new DragEventHandler(this.txtStudent_PreviewDragEnter);
                    this.txtStudent.PreviewDragOver += new DragEventHandler(this.txtStudent_PreviewDragEnter);
                    this.txtStudent.PreviewDrop += new DragEventHandler(this.txtStudent_PreviewDrop);
                    this.txtStudent.MouseDoubleClick += new MouseButtonEventHandler(this.txtFile_MouseDoubleClick);
                    return;

                case 3:
                    this.txtArrange = (TextBox) target;
                    this.txtArrange.PreviewDragEnter += new DragEventHandler(this.txtStudent_PreviewDragEnter);
                    this.txtArrange.PreviewDragOver += new DragEventHandler(this.txtStudent_PreviewDragEnter);
                    this.txtArrange.PreviewDrop += new DragEventHandler(this.txtStudent_PreviewDrop);
                    this.txtArrange.MouseDoubleClick += new MouseButtonEventHandler(this.txtFile_MouseDoubleClick);
                    return;

                case 4:
                    this.txtTemplate = (TextBox) target;
                    this.txtTemplate.PreviewDragEnter += new DragEventHandler(this.txtStudent_PreviewDragEnter);
                    this.txtTemplate.PreviewDragOver += new DragEventHandler(this.txtStudent_PreviewDragEnter);
                    this.txtTemplate.PreviewDrop += new DragEventHandler(this.txtStudent_PreviewDrop);
                    this.txtTemplate.MouseDoubleClick += new MouseButtonEventHandler(this.txtFile_MouseDoubleClick);
                    return;

                case 5:
                    this.scrList = (ScrollViewer) target;
                    return;

                case 6:
                    this.itemsControl = (ItemsControl) target;
                    return;

                case 8:
                    this.btnOrigin = (Button) target;
                    this.btnOrigin.Click += new RoutedEventHandler(this.btnInit_Click);
                    return;

                case 9:
                    this.btnCombine = (Button) target;
                    this.btnCombine.Click += new RoutedEventHandler(this.btnCombine_Click);
                    return;

                case 10:
                    this.btnGenerate = (Button) target;
                    this.btnGenerate.Click += new RoutedEventHandler(this.btnGenerate_Click);
                    return;
            }
            this._contentLoaded = true;
        }

        [EditorBrowsable(EditorBrowsableState.Never), GeneratedCode("PresentationBuildTasks", "4.0.0.0"), DebuggerNonUserCode]
        void IStyleConnector.Connect(int connectionId, object target)
        {
            if (connectionId == 7)
            {
                ((ToggleButton) target).Checked += new RoutedEventHandler(this.CheckBox_Checked);
                ((ToggleButton) target).Unchecked += new RoutedEventHandler(this.CheckBox_Unchecked);
            }
        }

        private void txtFile_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog {
                Filter = "Excel文件|*.xls;*.xlsx|所有文件|*.*"
            };
            if (dialog.ShowDialog() == true)
            {
                (sender as TextBox).Text = dialog.FileName;
            }
        }

        private void txtStudent_PreviewDragEnter(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.Copy;
            e.Handled = true;
        }

        private void txtStudent_PreviewDrop(object sender, DragEventArgs e)
        {
            object data = e.Data.GetData(DataFormats.FileDrop);
            TextBox box = sender as TextBox;
            if (box != null)
            {
                box.Text = string.Format("{0}", ((string[]) data)[0]);
            }
            this.CheckFileFormat(new TextBox[] { box });
        }

        private void Window_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.F1)
            {
                MessageBox.Show("注意事项：\r\n1、课程名称不能包含除*号以外的无法成为文件名的字符\r\n2、课程名中所有*号，在签到表名称中将被替换为※\r\n3、学籍表标题必须包含以下项：\r\n    学号、姓名、来源、层次、专业名称、年级\r\n4、考试安排表标题必须包含以下项：\r\n    专业名称、课程名称、考试性质、考试时间\r\n5、生成的签到表保存在考试安排表所在目录下");
            }
        }

        private void WriteIO(string filePath, IWorkbook workbook)
        {
            FileStream stream = null;
            try
            {
                using (stream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(stream);
                }
            }
            catch (Exception exception)
            {
                throw new IOException("在路径" + Path.GetDirectoryName(filePath) + "无法创建新文件，请检查该路径的访问权限！\r\n" + exception.ToString());
            }
            finally
            {
                if (stream != null)
                {
                    stream.Dispose();
                }
            }
        }
    }
}

