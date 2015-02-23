using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

namespace ExcelSheetExplorer
{
    
    public partial class Form1 : Form
    {
        private string selectedFilePath=null;//�I�𒆂̃t�@�C���̃t���p�X
        private bool is_ignore_error = true;//�V�[�g���擾���s�𖳎��t���O

        //������ƃX���b�h�֌W�̃f�[�^
        private string search_path = null;//�����p�X
        private string search_keyword = null;//�����L�[���[�h
        //ListView3�ւ̗v�f�ǉ��p�f���Q�[�g
        private delegate void AddListView3ItemDelegate(ListViewItem item);

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = null;

            //�O��I�����̃f�B���N�g�����J��
            string default_path = read_setting(Application.StartupPath + "\\default.txt");
            if(default_path!=null){
                open_folder(default_path);
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            //�I�����̃f�B���N�g�����t�@�C���ɋL�^
            write_setting(Application.StartupPath + "\\default.txt", textBox1.Text);
        }

        private void splitContainer1_Panel2_Paint(object sender, PaintEventArgs e)
        {

        }
 
        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //���ڂ��P���I������Ă��Ȃ��ꍇ
            if (listView1.SelectedItems.Count == 0)
            {
                listView2.Items.Clear();
                return;//�����𔲂���
            }
            
            //StatusStrip�Ƀt�@�C�����܂ł�FullPath��\��
            selectedFilePath = textBox1.Text+"\\" + listView1.SelectedItems[0].Text;
            toolStripStatusLabel1.Text = selectedFilePath;

            //listView2�̍X�V
            listView2.Items.Clear();

            //Excel�t�@�C���̏ꍇ
            if (isXLS(selectedFilePath))
            {
                //�V�[�g���̎擾
                MyExcelSheets xls = new MyExcelSheets(selectedFilePath);
                List<string> sheets = xls.sheetsNameList;
                if (sheets != null)
                {
                    //listView2�ɒǉ�
                    foreach (string sh in sheets)
                    {
                        listView2.Items.Add(sh);
                    }
                }
                else
                {
                    listView2.Items.Add("�V�[�g���̎擾���s");
                }
            }
        }

        private void listView1_DoubleClick(object sender, EventArgs e)
        {
            if (selectedFilePath != null)
            {               
                if(isXLS(selectedFilePath))//Excel�t�@�C���̏ꍇ
                {
                    Process.Start(selectedFilePath);//�V�F���Ɋ֘A�t����ꂽ�A�v���ŋN��
                }
                else if(isDIR(selectedFilePath))//�t�H���_�̏ꍇ
                {
                    open_folder(selectedFilePath);//�W�J
                }
            }
        }

        private void listView2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            string path = folderBrowserDialog1.SelectedPath;
            open_folder(path);
        }

        private bool open_folder(string path)
        {
            try
            {
                //��������f�B���N�g���̓W�J
                DirectoryInfo di = new DirectoryInfo(path);
                if (di.Exists)
                {
                    listView1.Items.Clear();
                    listView2.Items.Clear();

                    //�f�B���N�g���̗�
                    DirectoryInfo[] subDirInfo = di.GetDirectories();

                    foreach (DirectoryInfo sub_di in subDirInfo)
                    {
                        listView1.Items.Add(sub_di.Name,0);
                    }

                    //.xls�t�@�C���̗�
                    foreach (FileInfo file in di.GetFiles())
                    {
                        if ((file.Extension.ToLower() == ".xls"))
                        {
                            listView1.Items.Add(file.Name,1);
                        }
                    }
                    //listView1.Sort();

                    //�\���̍X�V
                    textBox1.Text = path;
                    toolStripStatusLabel1.Text = path;
                    selectedFilePath = path;
                }
            }
            catch (Exception exc)
            {
                //MessageBox.Show(exc.Message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
  
            return true;
        }

        private bool isXLS(string path){
            try
            {
                FileInfo file = new FileInfo(path);
                //Excel�t�@�C���̏ꍇ
                if ((file.Extension.ToLower() == ".xls")) return true;
                else return false;
            }
            catch (System.IO.IOException exc)
            {
                return false;
            }
        }

        private bool isDIR(string path)
        {
            try
            {
                // �t�H���_�̑������擾����
                System.IO.FileAttributes uAttribute = System.IO.File.GetAttributes(path);

                // �{���Ƀf�B���N�g�����ǂ������f���� (�_���ςŔ��f����)
                if ((uAttribute & System.IO.FileAttributes.Directory) == System.IO.FileAttributes.Directory)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (System.IO.IOException exc)
            {
                return false;
            }
            catch (System.ArgumentException)
            {
                return false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (selectedFilePath != null && selectedFilePath != "")
            {
                //�e�t�H���_�Ɉړ�����
                try
                {
                    FileInfo file = new FileInfo(selectedFilePath);
                    if (isDIR(selectedFilePath))
                    {
                        open_folder(file.DirectoryName);
                    }
                    else
                    {
                        open_folder(file.Directory.Parent.FullName);
                    }
                }
                catch (System.IO.IOException exc)
                {
                }
            }
        }

        private void listView1_DragDrop(object sender, DragEventArgs e)
        {
            //Drop���ꂽ�t�@�C�����̎擾
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (files.Length != 0)
            {
                string path= files[0];
                if (isDIR(path))
                {
                    open_folder(path);//�f�B���N�g���̏ꍇ
                }
                else
                {
                    //�t�@�C���̏ꍇ�̓t�@�C��������f�B���N�g�����J��
                    FileInfo file = new FileInfo(path);
                    open_folder(file.DirectoryName);
                }

            }
        }

        private void listView1_DragEnter(object sender, DragEventArgs e)
        {
            //Drop�`�����t�@�C���݂̂Ɍ���
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.All;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void listView1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if(e.KeyChar == (char)Keys.Enter){
                listView1_DoubleClick(sender,e);
            }else if(e.KeyChar == (char)Keys.Back){
                button2_Click(sender, e);
            }
        }

        private string read_setting(string setting_file)
        {
            string default_path;
            StreamReader sr=null;
            try
            {
                sr = new StreamReader(setting_file);

                default_path = sr.ReadLine();
            }
            catch (System.IO.IOException)
            {
                default_path = null;
            }
            finally
            {
                if (sr != null)
                {
                    sr.Close();
                }
            }
            return default_path;
        }

        private bool write_setting(string setting_file, string default_path)
        {
            bool state = false;
            StreamWriter sw=null;
            try
            {
                sw = new StreamWriter(setting_file);

                sw.WriteLine(default_path);
                state = true;
            }
            catch (System.IO.IOException)
            {
                state = false;
            }
            finally
            {
                if (sw != null)
                {
                    sw.Close();
                }
                
            }
            return state;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //�w���v�\��
            Process.Start(Application.StartupPath + "\\help.txt");
        }

        //�����{�^��
        private void button4_Click(object sender, EventArgs e)
        {
            //�����̏�����
            listView3.Items.Clear();
            search_path = textBox1.Text;//�����ꏊ
            search_keyword = textBox2.Text.ToLower();//�����L�[���[�h(�啶������������)
            is_ignore_error = checkBox1.Checked;

            if (!isDIR(search_path))
            {
                MessageBox.Show("�����Ώۃt�H���_���J���܂���", "�G���[");
                return;
            }

            //�e���Ԃ̕ύX
            button4.Enabled = false;
            button5.Enabled = true;
            toolStripStatusLabel1.Text = "�����Ώ�xls�t�@�C���̈ꗗ���쐬���D�D�D";
        
            //�����p�̍�ƃX���b�h�J�n
            backgroundWorker1.RunWorkerAsync();
        }
        
        //���������̎���
        private void searchSheet(object sender, DoWorkEventArgs e)
        {
            //OR�����p�ɃL�[���[�h���g�[�N���ɕ���
            char [] sep = { ' ', '�@' };//���p�A�S�p��
            string[] keywords = search_keyword.Split(sep,StringSplitOptions.RemoveEmptyEntries);

            if (keywords.Length == 0)
            {
                DialogResult ret= MessageBox.Show("�L�[���[�h���󔒂ł��B�V�[�g���ꗗ���擾���܂����H", "�m�F", MessageBoxButtons.YesNo);
                if (ret == DialogResult.Yes)
                {
                    keywords = new string[]{ "" };
                }
                else
                {
                    e.Cancel = true;
                    return;
                }
            }

            //�������猟���J�n
            
            // sender�̒l��bgWorker�̒l�Ɠ���
            BackgroundWorker worker = (BackgroundWorker)sender;
            //ListView3�v�f�ǉ��p�f���Q�[�g
            AddListView3ItemDelegate dlg = new AddListView3ItemDelegate(addListView3Item);

            //xls�t�@�C���̃��X�g�擾
            List<string> xlsList = new List<string>();
            if (getAllXlsList(search_path, xlsList,sender,e))
            {
                int counter = 0;
                int listsize = xlsList.Count;
                //���ׂĂ�xls�t�@�C���ɂ���
                foreach (string xlsFile in xlsList)
                {
                    // �L�����Z������ĂȂ�������I�Ƀ`�F�b�N
                    if (worker.CancellationPending)
                    {
                        e.Cancel = true;
                        return;
                    }

                    //�i���󋵕�
                    worker.ReportProgress((int)(counter*100 / listsize));

                    //�������̃��X�|���X����̂��ߑҋ@���̑��̃X���b�h��D�悷��
                    System.Threading.Thread.Sleep(0);

                    //�t�@�C�����Ƃ̔�r
                    if (isMatchOR(new FileInfo(xlsFile).Name.ToLower(), keywords))
                    {
                        //��v�����ꍇListView�ɒǉ�
                        ListViewItem item = new ListViewItem(new FileInfo(xlsFile).Name);
                        item.SubItems.Add("<�t�@�C�����Ɉ�v>");
                        item.SubItems.Add(xlsFile);
                        //listView3.Items.Add(item);
                        this.Invoke(dlg,new object[]{item});//�f���Q�[�g�o�R�Œǉ�
                    }


                    //�V�[�g���̗�
                    MyExcelSheets xls = new MyExcelSheets(xlsFile);
                    if (xls != null)
                    {
                        if (xls.sheetsNameList != null)
                        {
                            foreach (string sheet_name in xls.sheetsNameList)
                            {
                                //�V�[�g���̔�r(�啶�������������A������v)
                                if (isMatchOR(sheet_name.ToLower(), keywords))
                                {
                                    //��v�����ꍇListView�ɒǉ�
                                    ListViewItem item = new ListViewItem(new FileInfo(xlsFile).Name);
                                    item.SubItems.Add(sheet_name);
                                    item.SubItems.Add(xlsFile);
                                    //listView3.Items.Add(item);
                                    this.Invoke(dlg, new object[] { item });//�f���Q�[�g�o�R�Œǉ�
                                }
                            }
                        }
                        else if(!is_ignore_error)
                        {
                            //�V�[�g���̎擾�Ɏ��s
                            //ListView�ɒǉ�
                            ListViewItem item = new ListViewItem(new FileInfo(xlsFile).Name);
                            item.SubItems.Add("<�V�[�g���̎擾���s>");
                            item.SubItems.Add(xlsFile);
                            //listView3.Items.Add(item);
                            this.Invoke(dlg, new object[] { item });//�f���Q�[�g�o�R�Œǉ�
                        }
                    }
                    
                    counter++;
                }
            }

            
        }

        //���ׂĂ�xls�t�@�C���̃��X�g�擾
        bool getAllXlsList(string root, List<string> result, object sender, DoWorkEventArgs e)
        {
            // sender�̒l��bgWorker�̒l�Ɠ���
            BackgroundWorker worker = (BackgroundWorker)sender;
            // �L�����Z������ĂȂ�������I�Ƀ`�F�b�N
            if (worker.CancellationPending)
            {
                e.Cancel = true;
                return false;
            }

            //�������̃��X�|���X����̂��ߑҋ@���̑��̃X���b�h��D�悷��
            System.Threading.Thread.Sleep(0);
            try
            {
                //�f�B���N�g�����̃t�@�C��
                foreach (string f in Directory.GetFiles(root, "*.xls"))
                {
                    result.Add(f);
                }

                //�T�u�t�H���_���ċA�I�ɓW�J
                foreach (string d in Directory.GetDirectories(root))
                {
                    if (!getAllXlsList(d, result,sender,e)) return false;
                }
                
            }
            catch (System.Exception)
            {
            }

            return true;
        }

        //������Ɍ����L�[�z�񂪊܂܂�邩��r����
        private bool isMatchOR(string str, string[] keys)
        {
            bool ret = false;
            foreach (string key in keys)
            {
                if (str.Contains(key))
                {
                    ret = true;//OR����
                    break;
                }
            }
            return ret;
        }

        //ListView3�ւ̗v�f�ǉ��p�f���Q�[�g�֐�
        private void addListView3Item(ListViewItem item)
        {
            listView3.Items.Add(item);
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                button4_Click(sender, e);
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                if (isDIR(textBox1.Text))//�t�H���_�̏ꍇ
                {
                    open_folder(textBox1.Text);//�W�J
                }
            }
        }

        private void listView3_DoubleClick(object sender, EventArgs e)
        {
            //�������ʂ���Excel�t�@�C�����J��
            if (listView3.SelectedItems.Count != 0)
            {
                string xlsFile = listView3.SelectedItems[0].SubItems[2].Text;
                Process.Start(xlsFile);//�V�F���Ɋ֘A�t����ꂽ�A�v���ŋN��
            }
        }

        //�������f�{�^��
        private void button5_Click(object sender, EventArgs e)
        {
            backgroundWorker1.CancelAsync();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            searchSheet(sender,e);//���������̎���
        }

        //�i����
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            toolStripStatusLabel1.Text = "�������D�D�D" + e.ProgressPercentage + "%����";
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // ��������
            string msg = null;
            if (e.Cancelled)
            {
                msg = "���������f����܂���";
            }
            else
            {
                msg = "�������ʁF" + listView3.Items.Count + "���̃f�[�^��������܂���";
            }
            
            //�I������
            toolStripStatusLabel1.Text = msg;
            MessageBox.Show(msg, "��������");
            button4.Enabled = true;
            button5.Enabled = false;
        }

 
    }
}