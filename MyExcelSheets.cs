using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace ExcelSheetExplorer
{
    //MyExcelSheetsクラスはコンストラクタにExcelファイルのフルパスを渡すことで
    //バイナリを直接解析し、シート名の配列sheetsNameListを取得します
    //ファイルがExcelファイルであるかのチェックは行われません
    //読み込みに失敗した場合sheetNameListプロパティはnullを返します
    //既に開いているファイルの読み込みは失敗します
    class MyExcelSheets
    {
        private string fullPath;
        private List<string> sheets;

        public List<string> sheetsNameList
        {
            get { return sheets; }
        }

        //コンストラクタ
        
        public MyExcelSheets(string fileName)
        {
            sheets=new List<string>();
            fullPath=fileName;

            if (!readSheetsName(fullPath))
            {
                //読み込み失敗
                //MessageBox.Show("シート名の読み込みに失敗しました:"+fullPath);
                sheets = null;
            }
        }

        //シート名の読み込み
        private bool readSheetsName(string xlsFile)
        {
            FileStream br = null;
            bool stateFlg = false;
            //List<string> log = new List<string>();
            try
            {
                //ファイルを開く
                br = new FileStream(xlsFile, FileMode.Open, FileAccess.Read);

                int b1, b2, b3, b4;
                bool nextSheetFlg = false;
                int counter=0;
                //BOF(0x0809)を探す。上位と下位バイトは反転して格納されている。
                //先頭2バイトがレコードタイプ、次の2バイトがレコード長
                //log.Add("find BOF start");
                do{
                    int seek = (++counter)*256;
                    br.Seek(seek, SeekOrigin.Begin);
                    b1 = br.ReadByte();
                    b2 = br.ReadByte();
                    b3 = br.ReadByte();
                    b4 = br.ReadByte();
                    if (b1 == -1 || b2 == -1 || b3==-1 || b4==-1) throw new System.Exception("Unknown File Format: BOF is not found.");
                } while (!(b1 == 0x09 && b2 == 0x08));

                //log.Add("BOF is detected at:" + (br.Position-4).ToString("X"));
                //BOUNDSHEET(0x0085)レコードを探す
                while (!stateFlg)
                {
                    //log.Add("find BOUNDSHEET start");
                    //次のレコードの読み込み
                    do{
                        int seek = b4*256 + b3;
                        br.Seek(seek, SeekOrigin.Current);
                        string pos = "Position:" + br.Position.ToString("X");
                        b1 = br.ReadByte();
                        b2 = br.ReadByte();
                        b3 = br.ReadByte();
                        b4 = br.ReadByte();
                        //log.Add(pos + "\tRecord:" + b1.ToString("X") + " " + b2.ToString("X") + "=" + (b2*256+b1) );
                        if (b1 == -1 || b2 == -1 || b3 == -1 || b4 == -1) throw new System.Exception("Unknown File Format: BOUNDSHEET is not found.");
                        if (nextSheetFlg && !(b1 == 0x85 && b2 == 0x00))
                        {
                            //シート名が連続したレコードに格納されていると仮定し
                            //以降のレコードを読み込まない
                            stateFlg = true;
                            //log.Add("stateFlg=true && end of BOUNDSHEET");
                            break;
                        }
                    }while (!(b1 == 0x85 && b2 == 0x00));

                    if (stateFlg) break;
                    //log.Add("BOUNDSHEET is detected");
                    //シート名の読み込み
                    //Excel97以降のBIFF8形式
                    long save_pos = br.Position;
                    br.Seek(6, SeekOrigin.Current);
                    int sheetNameLength = br.ReadByte();//バイト数ではなく文字数
                    int unicodeType = br.ReadByte();//Unicode圧縮されているか？

                    Byte[] sheetName = new Byte[sheetNameLength * 2];

                    //シート名はUTF16リトルエンディアン(上下位逆順)で格納されている
                    for (int i = 0; i < sheetNameLength; i++)
                    {
                        sheetName[i * 2] = (Byte)br.ReadByte();
                        if (unicodeType == 0)
                        {
                            //Unicode圧縮されている場合
                            sheetName[i * 2 + 1] = 0;
                        }
                        else if (unicodeType == 1)
                        {
                            //Unicode圧縮されていない場合
                            sheetName[i * 2 + 1] = (Byte)br.ReadByte();
                        }
                        else
                        {
                            throw new System.Exception("Unknown File Format: Undifined UnicodeType(0x" + unicodeType.ToString("X") + ")is detected.");
                        }
                    }
                    //リストに追加
                    sheets.Add(byte2string(sheetName));
                    //log.Add("Add sheet:" + byte2string(sheetName));
                    //連続したシート名をチェック
                    nextSheetFlg = true;
                    //次のループのためにファイルポインタを戻す
                    br.Seek(save_pos, SeekOrigin.Begin);
                }
            }
            catch (System.IO.IOException e)
            {
                //MessageBox.Show(e.Message, "ファイルの読み込みに失敗しました");
                stateFlg = false;
            }
            catch (System.Exception e)
            {
                //MessageBox.Show(e.Message, "例外が発生しました");
                stateFlg = false;
                //log.Add("Exception Position:" + br.Position.ToString("X"));
            }
            finally
            {
                if (br != null)
                    br.Close();
                //StreamWriter sw = new StreamWriter("debuglog.txt");
                //foreach (string msg in log)
                //{
                //    sw.WriteLine(msg);
                //}
                //sw.Close();
            }
            return stateFlg;
        }

        /*
        //シート名の読み込み
        private bool readSheetsName(string xlsFile)
        {
            FileStream br=null;
            bool stateFlg = false;
            List<string> log = new List<string>();
            try
            {
                log.Add("start open file");
                //ファイルを開く
                br = new FileStream(xlsFile, FileMode.Open, FileAccess.Read);
                log.Add("end open file");

                int b1, b2;
                bool nextSheetFlg = false;
                while ((b1 = br.ReadByte()) != -1)
                {
                    log.Add("Position:" + br.Position.ToString("X") + ":" + b1.ToString("X"));
                    //Sheet名が格納されているレコード番号8500h
                    if (b1 == 0x85)
                    {
                        log.Add("0x85 found");
                        if ((b2 = br.ReadByte()) == -1)
                            break;
                        log.Add("Position:" + br.Position.ToString("X")+":"+b2.ToString("X"));
                        if (b2 == 0x00)
                        {
                            log.Add("0x00 found");
                            //シート名の読み込み
                            //Excel97以降のBIFF8形式
                            br.Seek(8, SeekOrigin.Current);
                            int sheetNameLength = br.ReadByte();//バイト数ではなく文字数
                            int unicodeType = br.ReadByte();//Unicode圧縮されているか？

                            Byte[] sheetName = new Byte[sheetNameLength * 2];

                            log.Add("Length=Byte*2=" + sheetNameLength + " unicode=" + unicodeType);
                            //シート名はUTF16リトルエンディアン(上下位逆順)で格納されている
                            for (int i = 0; i < sheetNameLength; i++)
                            {
                                sheetName[i * 2] = (Byte)br.ReadByte();
                                log.Add("Position:" + br.Position.ToString("X")+":"+sheetName[i*2].ToString("X"));
                                if (unicodeType == 0)
                                {
                                    //Unicode圧縮されている場合
                                    sheetName[i * 2 + 1] = 0;
                                }
                                else
                                {
                                    //Unicode圧縮されていない場合
                                    sheetName[i * 2 + 1] = (Byte)br.ReadByte();
                                    log.Add("Position:" + br.Position.ToString("X") + ":" + sheetName[i * 2+1].ToString("X"));
                                }
                            }
                            //リストに追加
                            sheets.Add(byte2string(sheetName));
                            log.Add("sheets add=" + byte2string(sheetName));
                            //連続したシート名をチェック
                            nextSheetFlg = true;
                            log.Add("set nextSheetsFlg true");

                        }
                        else if (nextSheetFlg)
                        {
                            log.Add("nextSheetFlg is false. End of sheets");
                            //すべてのシート名読み取り終了
                            stateFlg = true;
                            break;
                        }

                    }
                    else if (nextSheetFlg)
                    {
                        log.Add("nextSheetFlg is false. End of sheets");
                        //すべてのシート名読み取り終了
                        stateFlg = true;
                        break;
                    }

                }

            }
            catch (System.IO.IOException e)
            {
                MessageBox.Show(e.Message, "ファイルの読み込みに失敗しました");
                stateFlg = false;
            }
            catch (System.Exception e)
            {
                MessageBox.Show(e.Message, "例外が発生しました");
                stateFlg = false;
            }
            finally
            {
                if (br != null)
                    br.Close();
                if (true)
                {
                    StreamWriter sw = new StreamWriter("debuglog.txt");
                    foreach (string msg in log)
                    {
                        sw.WriteLine(msg);
                    }
                    sw.Close();
                }
            }
            return stateFlg;
        }
        */

        //Unicodeバイト列をstringに変換
        private string byte2string(Byte[] bytes)
        {
            Encoding e = Encoding.Unicode;
            return e.GetString(bytes);
        }

    }
}
