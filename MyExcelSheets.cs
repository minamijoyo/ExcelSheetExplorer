using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace ExcelSheetExplorer
{
    //MyExcelSheets�N���X�̓R���X�g���N�^��Excel�t�@�C���̃t���p�X��n�����Ƃ�
    //�o�C�i���𒼐ډ�͂��A�V�[�g���̔z��sheetsNameList���擾���܂�
    //�t�@�C����Excel�t�@�C���ł��邩�̃`�F�b�N�͍s���܂���
    //�ǂݍ��݂Ɏ��s�����ꍇsheetNameList�v���p�e�B��null��Ԃ��܂�
    //���ɊJ���Ă���t�@�C���̓ǂݍ��݂͎��s���܂�
    class MyExcelSheets
    {
        private string fullPath;
        private List<string> sheets;

        public List<string> sheetsNameList
        {
            get { return sheets; }
        }

        //�R���X�g���N�^
        
        public MyExcelSheets(string fileName)
        {
            sheets=new List<string>();
            fullPath=fileName;

            if (!readSheetsName(fullPath))
            {
                //�ǂݍ��ݎ��s
                //MessageBox.Show("�V�[�g���̓ǂݍ��݂Ɏ��s���܂���:"+fullPath);
                sheets = null;
            }
        }

        //�V�[�g���̓ǂݍ���
        private bool readSheetsName(string xlsFile)
        {
            FileStream br = null;
            bool stateFlg = false;
            //List<string> log = new List<string>();
            try
            {
                //�t�@�C�����J��
                br = new FileStream(xlsFile, FileMode.Open, FileAccess.Read);

                int b1, b2, b3, b4;
                bool nextSheetFlg = false;
                int counter=0;
                //BOF(0x0809)��T���B��ʂƉ��ʃo�C�g�͔��]���Ċi�[����Ă���B
                //�擪2�o�C�g�����R�[�h�^�C�v�A����2�o�C�g�����R�[�h��
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
                //BOUNDSHEET(0x0085)���R�[�h��T��
                while (!stateFlg)
                {
                    //log.Add("find BOUNDSHEET start");
                    //���̃��R�[�h�̓ǂݍ���
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
                            //�V�[�g�����A���������R�[�h�Ɋi�[����Ă���Ɖ��肵
                            //�ȍ~�̃��R�[�h��ǂݍ��܂Ȃ�
                            stateFlg = true;
                            //log.Add("stateFlg=true && end of BOUNDSHEET");
                            break;
                        }
                    }while (!(b1 == 0x85 && b2 == 0x00));

                    if (stateFlg) break;
                    //log.Add("BOUNDSHEET is detected");
                    //�V�[�g���̓ǂݍ���
                    //Excel97�ȍ~��BIFF8�`��
                    long save_pos = br.Position;
                    br.Seek(6, SeekOrigin.Current);
                    int sheetNameLength = br.ReadByte();//�o�C�g���ł͂Ȃ�������
                    int unicodeType = br.ReadByte();//Unicode���k����Ă��邩�H

                    Byte[] sheetName = new Byte[sheetNameLength * 2];

                    //�V�[�g����UTF16���g���G���f�B�A��(�㉺�ʋt��)�Ŋi�[����Ă���
                    for (int i = 0; i < sheetNameLength; i++)
                    {
                        sheetName[i * 2] = (Byte)br.ReadByte();
                        if (unicodeType == 0)
                        {
                            //Unicode���k����Ă���ꍇ
                            sheetName[i * 2 + 1] = 0;
                        }
                        else if (unicodeType == 1)
                        {
                            //Unicode���k����Ă��Ȃ��ꍇ
                            sheetName[i * 2 + 1] = (Byte)br.ReadByte();
                        }
                        else
                        {
                            throw new System.Exception("Unknown File Format: Undifined UnicodeType(0x" + unicodeType.ToString("X") + ")is detected.");
                        }
                    }
                    //���X�g�ɒǉ�
                    sheets.Add(byte2string(sheetName));
                    //log.Add("Add sheet:" + byte2string(sheetName));
                    //�A�������V�[�g�����`�F�b�N
                    nextSheetFlg = true;
                    //���̃��[�v�̂��߂Ƀt�@�C���|�C���^��߂�
                    br.Seek(save_pos, SeekOrigin.Begin);
                }
            }
            catch (System.IO.IOException e)
            {
                //MessageBox.Show(e.Message, "�t�@�C���̓ǂݍ��݂Ɏ��s���܂���");
                stateFlg = false;
            }
            catch (System.Exception e)
            {
                //MessageBox.Show(e.Message, "��O���������܂���");
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
        //�V�[�g���̓ǂݍ���
        private bool readSheetsName(string xlsFile)
        {
            FileStream br=null;
            bool stateFlg = false;
            List<string> log = new List<string>();
            try
            {
                log.Add("start open file");
                //�t�@�C�����J��
                br = new FileStream(xlsFile, FileMode.Open, FileAccess.Read);
                log.Add("end open file");

                int b1, b2;
                bool nextSheetFlg = false;
                while ((b1 = br.ReadByte()) != -1)
                {
                    log.Add("Position:" + br.Position.ToString("X") + ":" + b1.ToString("X"));
                    //Sheet�����i�[����Ă��郌�R�[�h�ԍ�8500h
                    if (b1 == 0x85)
                    {
                        log.Add("0x85 found");
                        if ((b2 = br.ReadByte()) == -1)
                            break;
                        log.Add("Position:" + br.Position.ToString("X")+":"+b2.ToString("X"));
                        if (b2 == 0x00)
                        {
                            log.Add("0x00 found");
                            //�V�[�g���̓ǂݍ���
                            //Excel97�ȍ~��BIFF8�`��
                            br.Seek(8, SeekOrigin.Current);
                            int sheetNameLength = br.ReadByte();//�o�C�g���ł͂Ȃ�������
                            int unicodeType = br.ReadByte();//Unicode���k����Ă��邩�H

                            Byte[] sheetName = new Byte[sheetNameLength * 2];

                            log.Add("Length=Byte*2=" + sheetNameLength + " unicode=" + unicodeType);
                            //�V�[�g����UTF16���g���G���f�B�A��(�㉺�ʋt��)�Ŋi�[����Ă���
                            for (int i = 0; i < sheetNameLength; i++)
                            {
                                sheetName[i * 2] = (Byte)br.ReadByte();
                                log.Add("Position:" + br.Position.ToString("X")+":"+sheetName[i*2].ToString("X"));
                                if (unicodeType == 0)
                                {
                                    //Unicode���k����Ă���ꍇ
                                    sheetName[i * 2 + 1] = 0;
                                }
                                else
                                {
                                    //Unicode���k����Ă��Ȃ��ꍇ
                                    sheetName[i * 2 + 1] = (Byte)br.ReadByte();
                                    log.Add("Position:" + br.Position.ToString("X") + ":" + sheetName[i * 2+1].ToString("X"));
                                }
                            }
                            //���X�g�ɒǉ�
                            sheets.Add(byte2string(sheetName));
                            log.Add("sheets add=" + byte2string(sheetName));
                            //�A�������V�[�g�����`�F�b�N
                            nextSheetFlg = true;
                            log.Add("set nextSheetsFlg true");

                        }
                        else if (nextSheetFlg)
                        {
                            log.Add("nextSheetFlg is false. End of sheets");
                            //���ׂẴV�[�g���ǂݎ��I��
                            stateFlg = true;
                            break;
                        }

                    }
                    else if (nextSheetFlg)
                    {
                        log.Add("nextSheetFlg is false. End of sheets");
                        //���ׂẴV�[�g���ǂݎ��I��
                        stateFlg = true;
                        break;
                    }

                }

            }
            catch (System.IO.IOException e)
            {
                MessageBox.Show(e.Message, "�t�@�C���̓ǂݍ��݂Ɏ��s���܂���");
                stateFlg = false;
            }
            catch (System.Exception e)
            {
                MessageBox.Show(e.Message, "��O���������܂���");
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

        //Unicode�o�C�g���string�ɕϊ�
        private string byte2string(Byte[] bytes)
        {
            Encoding e = Encoding.Unicode;
            return e.GetString(bytes);
        }

    }
}
