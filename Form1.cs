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
        private string selectedFilePath=null;//選択中のファイルのフルパス
        private bool is_ignore_error = true;//シート名取得失敗を無視フラグ

        //検索作業スレッド関係のデータ
        private string search_path = null;//検索パス
        private string search_keyword = null;//検索キーワード
        //ListView3への要素追加用デリゲート
        private delegate void AddListView3ItemDelegate(ListViewItem item);

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = null;

            //前回終了時のディレクトリを開く
            string default_path = read_setting(Application.StartupPath + "\\default.txt");
            if(default_path!=null){
                open_folder(default_path);
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            //終了時のディレクトリをファイルに記録
            write_setting(Application.StartupPath + "\\default.txt", textBox1.Text);
        }

        private void splitContainer1_Panel2_Paint(object sender, PaintEventArgs e)
        {

        }
 
        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //項目が１つも選択されていない場合
            if (listView1.SelectedItems.Count == 0)
            {
                listView2.Items.Clear();
                return;//処理を抜ける
            }
            
            //StatusStripにファイル名までのFullPathを表示
            selectedFilePath = textBox1.Text+"\\" + listView1.SelectedItems[0].Text;
            toolStripStatusLabel1.Text = selectedFilePath;

            //listView2の更新
            listView2.Items.Clear();

            //Excelファイルの場合
            if (isXLS(selectedFilePath))
            {
                //シート名の取得
                MyExcelSheets xls = new MyExcelSheets(selectedFilePath);
                List<string> sheets = xls.sheetsNameList;
                if (sheets != null)
                {
                    //listView2に追加
                    foreach (string sh in sheets)
                    {
                        listView2.Items.Add(sh);
                    }
                }
                else
                {
                    listView2.Items.Add("シート名の取得失敗");
                }
            }
        }

        private void listView1_DoubleClick(object sender, EventArgs e)
        {
            if (selectedFilePath != null)
            {               
                if(isXLS(selectedFilePath))//Excelファイルの場合
                {
                    Process.Start(selectedFilePath);//シェルに関連付けられたアプリで起動
                }
                else if(isDIR(selectedFilePath))//フォルダの場合
                {
                    open_folder(selectedFilePath);//展開
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
                //ここからディレクトリの展開
                DirectoryInfo di = new DirectoryInfo(path);
                if (di.Exists)
                {
                    listView1.Items.Clear();
                    listView2.Items.Clear();

                    //ディレクトリの列挙
                    DirectoryInfo[] subDirInfo = di.GetDirectories();

                    foreach (DirectoryInfo sub_di in subDirInfo)
                    {
                        listView1.Items.Add(sub_di.Name,0);
                    }

                    //.xlsファイルの列挙
                    foreach (FileInfo file in di.GetFiles())
                    {
                        if ((file.Extension.ToLower() == ".xls"))
                        {
                            listView1.Items.Add(file.Name,1);
                        }
                    }
                    //listView1.Sort();

                    //表示の更新
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
                //Excelファイルの場合
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
                // フォルダの属性を取得する
                System.IO.FileAttributes uAttribute = System.IO.File.GetAttributes(path);

                // 本当にディレクトリかどうか判断する (論理積で判断する)
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
                //親フォルダに移動する
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
            //Dropされたファイル名の取得
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (files.Length != 0)
            {
                string path= files[0];
                if (isDIR(path))
                {
                    open_folder(path);//ディレクトリの場合
                }
                else
                {
                    //ファイルの場合はファイルがあるディレクトリを開く
                    FileInfo file = new FileInfo(path);
                    open_folder(file.DirectoryName);
                }

            }
        }

        private void listView1_DragEnter(object sender, DragEventArgs e)
        {
            //Drop形式をファイルのみに限定
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
            //ヘルプ表示
            Process.Start(Application.StartupPath + "\\help.txt");
        }

        //検索ボタン
        private void button4_Click(object sender, EventArgs e)
        {
            //検索の初期化
            listView3.Items.Clear();
            search_path = textBox1.Text;//検索場所
            search_keyword = textBox2.Text.ToLower();//検索キーワード(大文字小文字無視)
            is_ignore_error = checkBox1.Checked;

            if (!isDIR(search_path))
            {
                MessageBox.Show("検索対象フォルダが開けません", "エラー");
                return;
            }

            //各種状態の変更
            button4.Enabled = false;
            button5.Enabled = true;
            toolStripStatusLabel1.Text = "検索対象xlsファイルの一覧を作成中．．．";
        
            //検索用の作業スレッド開始
            backgroundWorker1.RunWorkerAsync();
        }
        
        //検索処理の実装
        private void searchSheet(object sender, DoWorkEventArgs e)
        {
            //OR検索用にキーワードをトークンに分割
            char [] sep = { ' ', '　' };//半角、全角空白
            string[] keywords = search_keyword.Split(sep,StringSplitOptions.RemoveEmptyEntries);

            if (keywords.Length == 0)
            {
                DialogResult ret= MessageBox.Show("キーワードが空白です。シート名一覧を取得しますか？", "確認", MessageBoxButtons.YesNo);
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

            //ここから検索開始
            
            // senderの値はbgWorkerの値と同じ
            BackgroundWorker worker = (BackgroundWorker)sender;
            //ListView3要素追加用デリゲート
            AddListView3ItemDelegate dlg = new AddListView3ItemDelegate(addListView3Item);

            //xlsファイルのリスト取得
            List<string> xlsList = new List<string>();
            if (getAllXlsList(search_path, xlsList,sender,e))
            {
                int counter = 0;
                int listsize = xlsList.Count;
                //すべてのxlsファイルについて
                foreach (string xlsFile in xlsList)
                {
                    // キャンセルされてないか定期的にチェック
                    if (worker.CancellationPending)
                    {
                        e.Cancel = true;
                        return;
                    }

                    //進捗状況報告
                    worker.ReportProgress((int)(counter*100 / listsize));

                    //検索中のレスポンス向上のため待機中の他のスレッドを優先する
                    System.Threading.Thread.Sleep(0);

                    //ファイル名との比較
                    if (isMatchOR(new FileInfo(xlsFile).Name.ToLower(), keywords))
                    {
                        //一致した場合ListViewに追加
                        ListViewItem item = new ListViewItem(new FileInfo(xlsFile).Name);
                        item.SubItems.Add("<ファイル名に一致>");
                        item.SubItems.Add(xlsFile);
                        //listView3.Items.Add(item);
                        this.Invoke(dlg,new object[]{item});//デリゲート経由で追加
                    }


                    //シート名の列挙
                    MyExcelSheets xls = new MyExcelSheets(xlsFile);
                    if (xls != null)
                    {
                        if (xls.sheetsNameList != null)
                        {
                            foreach (string sheet_name in xls.sheetsNameList)
                            {
                                //シート名の比較(大文字小文字無視、部分一致)
                                if (isMatchOR(sheet_name.ToLower(), keywords))
                                {
                                    //一致した場合ListViewに追加
                                    ListViewItem item = new ListViewItem(new FileInfo(xlsFile).Name);
                                    item.SubItems.Add(sheet_name);
                                    item.SubItems.Add(xlsFile);
                                    //listView3.Items.Add(item);
                                    this.Invoke(dlg, new object[] { item });//デリゲート経由で追加
                                }
                            }
                        }
                        else if(!is_ignore_error)
                        {
                            //シート名の取得に失敗
                            //ListViewに追加
                            ListViewItem item = new ListViewItem(new FileInfo(xlsFile).Name);
                            item.SubItems.Add("<シート名の取得失敗>");
                            item.SubItems.Add(xlsFile);
                            //listView3.Items.Add(item);
                            this.Invoke(dlg, new object[] { item });//デリゲート経由で追加
                        }
                    }
                    
                    counter++;
                }
            }

            
        }

        //すべてのxlsファイルのリスト取得
        bool getAllXlsList(string root, List<string> result, object sender, DoWorkEventArgs e)
        {
            // senderの値はbgWorkerの値と同じ
            BackgroundWorker worker = (BackgroundWorker)sender;
            // キャンセルされてないか定期的にチェック
            if (worker.CancellationPending)
            {
                e.Cancel = true;
                return false;
            }

            //検索中のレスポンス向上のため待機中の他のスレッドを優先する
            System.Threading.Thread.Sleep(0);
            try
            {
                //ディレクトリ内のファイル
                foreach (string f in Directory.GetFiles(root, "*.xls"))
                {
                    result.Add(f);
                }

                //サブフォルダを再帰的に展開
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

        //文字列に検索キー配列が含まれるか比較する
        private bool isMatchOR(string str, string[] keys)
        {
            bool ret = false;
            foreach (string key in keys)
            {
                if (str.Contains(key))
                {
                    ret = true;//OR検索
                    break;
                }
            }
            return ret;
        }

        //ListView3への要素追加用デリゲート関数
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
                if (isDIR(textBox1.Text))//フォルダの場合
                {
                    open_folder(textBox1.Text);//展開
                }
            }
        }

        private void listView3_DoubleClick(object sender, EventArgs e)
        {
            //検索結果からExcelファイルを開く
            if (listView3.SelectedItems.Count != 0)
            {
                string xlsFile = listView3.SelectedItems[0].SubItems[2].Text;
                Process.Start(xlsFile);//シェルに関連付けられたアプリで起動
            }
        }

        //検索中断ボタン
        private void button5_Click(object sender, EventArgs e)
        {
            backgroundWorker1.CancelAsync();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            searchSheet(sender,e);//検索処理の実装
        }

        //進捗報告
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            toolStripStatusLabel1.Text = "検索中．．．" + e.ProgressPercentage + "%完了";
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // 検索結果
            string msg = null;
            if (e.Cancelled)
            {
                msg = "検索が中断されました";
            }
            else
            {
                msg = "検索結果：" + listView3.Items.Count + "件のデータが見つかりました";
            }
            
            //終了処理
            toolStripStatusLabel1.Text = msg;
            MessageBox.Show(msg, "検索結果");
            button4.Enabled = true;
            button5.Enabled = false;
        }

 
    }
}