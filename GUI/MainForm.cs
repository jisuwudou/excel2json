using FastColoredTextBoxNS;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using System.IO;
using System.Data;

namespace excel2json.GUI
{

    /// <summary>
    /// 主窗口
    /// </summary>
    public partial class MainForm : Form
    {
        // Excel导入数据管理
        private DataManager mDataMgr;
        private string mCurrentXlsx;

        // 支持语法高亮的文本框
        private FastColoredTextBox mJsonTextBox;
        private FastColoredTextBox mCSharpTextBox;


        // 文本框的样式
        private TextStyle mBrownStyle = new TextStyle(Brushes.Brown, null, FontStyle.Regular);
        private TextStyle mMagentaStyle = new TextStyle(Brushes.Magenta, null, FontStyle.Regular);
        private TextStyle mGreenStyle = new TextStyle(Brushes.Green, null, FontStyle.Regular);

        // 导出数据相关的按钮，方便整体Enable/Disable
        private List<ToolStripButton> mExportButtonList;

        // 打开的excel文件名，不包含后缀xlsx。。。
        private String FileName;
        private List<string> records = new List<string>();
        private int record_cur_idx = -1;
        private string record_path = "./../../Record.txt";
        /// <summary>
        /// 构造函数，初始化控件初值；创建文本框
        /// </summary>
        public MainForm()
        {
            InitializeComponent();

            //-- syntax highlight text box
            mJsonTextBox = createTextBoxInTab(this.tabPageJSON);
            mJsonTextBox.Language = Language.Custom;
            mJsonTextBox.TextChanged += new EventHandler<TextChangedEventArgs>(this.jsonTextChanged);

            mCSharpTextBox = createTextBoxInTab(this.tabCSharp);
            mCSharpTextBox.Language = Language.CSharp;


            if (!File.Exists(record_path))
            {
                FileStream fs1 = new FileStream(record_path, FileMode.Create, FileAccess.Write);//创建写入文件 
                StreamWriter sw = new StreamWriter(fs1);
                //sw.WriteLine(this.textBox3.Text.Trim() + "+" + this.textBox4.Text);//开始写入值
                //System.Console.WriteLine()
                //sw.WriteLine("Init Flag");
                sw.Close();
                fs1.Close();
            }
            else
            {
                //StreamReader 简单读取
                StreamReader sr = new StreamReader(record_path, Encoding.Default);//初始化读取 设置编码格式，否则中文会乱码
                
                string read_line = sr.ReadLine();
                string show = "";
                while(read_line != null)
                {
                    show += (" "+read_line);
                    records.Add(show);
                    read_line = sr.ReadLine();
                }
                System.Console.WriteLine(show);
                sr.Close();
                
            }


            //-- componet init states
            this.comboBoxType.SelectedIndex = 0;
            this.comboBoxLowcase.SelectedIndex = 1;
            this.comboBoxHeader.SelectedIndex = 1;
            this.comboBoxDateFormat.SelectedIndex = 0;
            this.comboBoxSheetName.SelectedIndex = 1;
            
            this.comboBoxType.SelectedIndexChanged += comboBoxType_SelectedIndexChanged;

            this.comboBoxEncoding.Items.Clear();
            this.comboBoxEncoding.Items.Add("utf8-nobom");
            foreach (EncodingInfo ei in Encoding.GetEncodings())
            {
                Encoding e = ei.GetEncoding();
                this.comboBoxEncoding.Items.Add(e.HeaderName);
            }
            this.comboBoxEncoding.SelectedIndex = 0;

            //-- button list
            mExportButtonList = new List<ToolStripButton>();
            mExportButtonList.Add(this.btnCopyJSON);
            mExportButtonList.Add(this.btnSaveJson);
            mExportButtonList.Add(this.btnCopyCSharp);
            mExportButtonList.Add(this.btnSaveCSharp);
            enableExportButtons(false);

            //-- data manager
            mDataMgr = new DataManager();
            this.btnReimport.Enabled = false;
        }

        /// <summary>
        /// 设置导出相关的按钮是否可用
        /// </summary>
        /// <param name="enable">是否可用</param>
        private void enableExportButtons(bool enable)
        {
            foreach (var btn in mExportButtonList)
                btn.Enabled = enable;
        }

        /// <summary>
        /// 在一个TabPage中创建Text Box
        /// </summary>
        /// <param name="tab">TabPage容器控件</param>
        /// <returns>新建的Text Box控件</returns>
        private FastColoredTextBox createTextBoxInTab(TabPage tab)
        {
            FastColoredTextBox textBox = new FastColoredTextBox();
            textBox.Dock = DockStyle.Fill;
            textBox.Font = new Font("Microsoft YaHei", 11F);
            tab.Controls.Add(textBox);
            return textBox;
        }

        /// <summary>
        /// 设置Json文本高亮格式
        /// </summary>
        private void jsonTextChanged(object sender, TextChangedEventArgs e)
        {
            e.ChangedRange.ClearStyle(mBrownStyle, mMagentaStyle, mGreenStyle);
            //allow to collapse brackets block
            e.ChangedRange.SetFoldingMarkers("{", "}");
            //string highlighting
            e.ChangedRange.SetStyle(mBrownStyle, @"""""|@""""|''|@"".*?""|(?<!@)(?<range>"".*?[^\\]"")|'.*?[^\\]'");
            //number highlighting
            e.ChangedRange.SetStyle(mGreenStyle, @"\b\d+[\.]?\d*([eE]\-?\d+)?[lLdDfF]?\b|\b0x[a-fA-F\d]+\b");
        }

        /// <summary>
        /// 使用BackgroundWorker加载Excel文件，使用UI中的Options设置
        /// </summary>
        /// <param name="path">Excel文件路径</param>
        private void loadExcelAsync(string path)
        {

            mCurrentXlsx = path;
            FileName = System.IO.Path.GetFileNameWithoutExtension(path);

            //-- update ui
            this.btnReimport.Enabled = true;
            this.labelExcelFile.Text = System.IO.Path.GetFileName(path);
            enableExportButtons(false);

            //读取默认状态
            bool isDone = false;
            for(int i = 0; i< records.Count; i++)
            {
                string[] element = records[i].Split(' ');
                Console.WriteLine(records[i]);
                if (element[0] != this.labelExcelFile.Text) continue;
                
                this.comboBoxType.SelectedIndex = Convert.ToInt32(element[1]);
                record_cur_idx = i;
                isDone = true;
                break;
            }

            if(false == isDone)
            {
                string newre = this.labelExcelFile.Text + " 0";
                records.Add(newre);
                this.record_cur_idx = records.Count - 1;
                this.comboBoxType.SelectedIndex = 0;
                save_record();
            }

            this.statusLabel.IsLink = false;
            this.statusLabel.Text = "Loading Excel ...";

            //-- load options from ui
            Program.Options options = new Program.Options();
            options.ExcelPath = path;
            options.ExportArray = this.comboBoxType.SelectedIndex == 0;
            options.Encoding = this.comboBoxEncoding.SelectedText;
            options.Lowcase = this.comboBoxLowcase.SelectedIndex == 0;
            options.HeaderRows = int.Parse(this.comboBoxHeader.Text);
            options.DateFormat = this.comboBoxDateFormat.Text;
            options.ForceSheetName = this.comboBoxSheetName.SelectedIndex == 0;
            options.ExcludePrefix = this.textBoxExculdePrefix.Text;
            options.CellJson = this.checkBoxCellJson.Checked;
            options.AllString = this.checkBoxAllString.Checked;

            //-- start import
            this.backgroundWorker.RunWorkerAsync(options);
        }


        private void comboBoxType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (record_cur_idx == -1)
                return;

            string element = records[record_cur_idx];
            string[] info = element.Split(' ');
            info[1] = Convert.ToString( this.comboBoxType.SelectedIndex);
            records[record_cur_idx] = info[0] + " " + info[1];
            save_record();
        }

        private void save_record()
        {
            System.IO.File.WriteAllText(record_path, string.Empty);
            FileStream fs1 = new FileStream(record_path, FileMode.Create, FileAccess.Write);//创建写入文件 
            fs1.Position = 0;
            StreamWriter sw = new StreamWriter(fs1);
            
            for (int i = 0; i < records.Count; i++)
            {
                sw.WriteLine(records[i]);
            }

            sw.Close();
            fs1.Close();
        }

        /// <summary>
        /// 接受Excel拖放事件
        /// </summary>
        private void panelExcelDropBox_DragDrop(object sender, DragEventArgs e)
        {
            string[] dropData = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (dropData != null)
            {
                this.loadExcelAsync(dropData[0]);
            }
        }

        /// <summary>
        /// 显示Help文档
        /// </summary>
        private void btnHelp_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://neil3d.github.io/coding/excel2json.html");
        }

        /// <summary>
        /// 判断拖放对象是否是一个.xlsx文件
        /// </summary>
        private void panelExcelDropBox_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] dropData = (string[])e.Data.GetData(DataFormats.FileDrop, false);
                if (dropData != null && dropData.Length > 0)
                {
                    string szPath = dropData[0];
                    string szExt = System.IO.Path.GetExtension(szPath);
                    FileName = System.IO.Path.GetFileNameWithoutExtension(szPath);
                    szExt = szExt.ToLower();
                    if (szExt == ".xlsx")
                    {
                        e.Effect = DragDropEffects.All;
                        return;
                    }
                }
            }//end of if(file)
            e.Effect = DragDropEffects.None;
        }

        /// <summary>
        /// 执行实际的Excel加载
        /// </summary>
        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            lock (this.mDataMgr)
            {
                this.mDataMgr.loadExcel((Program.Options)e.Argument);
            }
        }

        /// <summary>
        /// Excel加载完成
        /// </summary>
        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            // 判断错误信息
            if (e.Error != null)
            {
                showStatus(e.Error.Message, Color.Red);
                return;
            }

            // 更新UI
            lock (this.mDataMgr)
            {
                this.statusLabel.IsLink = false;
                this.statusLabel.Text = "Load completed.";

                mJsonTextBox.Text = mDataMgr.JsonContext;
                mCSharpTextBox.Text = mDataMgr.CSharpCode;

                enableExportButtons(true);
            }
        }

        /// <summary>
        /// 工具栏按钮：Import Excel
        /// </summary>
        private void btnImportExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.RestoreDirectory = true;
            dlg.Filter = "Excel File(*.xlsx)|*.xlsx";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                this.loadExcelAsync(dlg.FileName);
            }
        }

        /// <summary>
        /// 点击状态栏链接
        /// </summary>
        private void statusLabel_Click(object sender, EventArgs e)
        {
            if (this.statusLabel.IsLink)
            {
                System.Diagnostics.Process.Start(this.statusLabel.Text);
            }
        }

        /// <summary>
        /// 保存导出文件
        /// </summary>
        private void saveToFile(int type, string filter)
        {

            try
            {
                SaveFileDialog dlg = new SaveFileDialog();
                dlg.RestoreDirectory = true;
                dlg.Filter = filter;
                dlg.FileName = FileName;
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    lock (mDataMgr)
                    {
                        switch (type)
                        {
                            case 1:
                                mDataMgr.saveJson(dlg.FileName);
                                break;
                            case 2:
                                mDataMgr.saveCSharp(dlg.FileName);
                                break;
                        }
                    }
                    showStatus(string.Format("{0} saved!", dlg.FileName), Color.Black);
                }// end of if
            }
            catch (Exception ex)
            {
                showStatus(ex.Message, Color.Red);
            }
        }

        /// <summary>
        /// 工具栏按钮：Save Json
        /// </summary>
        private void btnSaveJson_Click(object sender, EventArgs e)
        {
            saveToFile(1, "Json File(*.json)|*.json");
        }

        /// <summary>
        /// 工具栏按钮：Copy Json
        /// </summary>
        private void btnCopyJSON_Click(object sender, EventArgs e)
        {
            lock (mDataMgr)
            {
                Clipboard.SetText(mDataMgr.JsonContext);
                showStatus("Json text copyed to clipboard.", Color.Black);
            }
        }

        /// <summary>
        /// 设置状态栏信息
        /// </summary>
        /// <param name="szMessage">信息文字</param>
        /// <param name="color">信息颜色</param>
        private void showStatus(string szMessage, Color color)
        {
            this.statusLabel.Text = szMessage;
            this.statusLabel.ForeColor = color;
            this.statusLabel.IsLink = false;
        }

        /// <summary>
        /// 配置项变更之后，手动重新导入xlsx文件
        /// </summary>
        private void btnReimport_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(mCurrentXlsx))
            {
                loadExcelAsync(mCurrentXlsx);
            }
        }

        private void btnCopyCSharp_Click(object sender, EventArgs e)
        {
            lock (mDataMgr)
            {
                Clipboard.SetText(mDataMgr.CSharpCode);
                showStatus("C# code copyed to clipboard.", Color.Black);
            }
        }

        private void btnSaveCSharp_Click(object sender, EventArgs e)
        {
            saveToFile(2, "C# code file(*.cs)|*.cs");
        }
    }
}
