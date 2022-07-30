using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using System.Collections;

namespace WordWolf
{
    public partial class FormMain : Form
    {
        int PLAYERS = 6;

        int RANDOMWORDNUM = 1;
        int RANDOMWOLFNUM = 1;

        Random WORDNUM;
        Random WOLFNUM;

        string SAFEWORD;
        string OUTWORD;

        int MASTERCOUNTTIME = 180;

        int COUNTTIME = 0;

        bool WORDOK = false;

        bool PSETSTATE = true;

        bool RESULT = true;

        bool WINCOUNT = true;


        //メール関連
        string MAILHOST = "smtp.gmail.com";
        int MAILPORT = 587;
        string USERNAME;
        string PASSWORD;

        public FormMain()
        {
            InitializeComponent();

            this.RESULT = true;
            this.WINCOUNT = true;

            this.WORDNUM = new System.Random();
            this.WOLFNUM = new System.Random();

            this.LblPlayer1.Text = string.Empty;
            this.LblPlayer2.Text = string.Empty;
            this.LblPlayer3.Text = string.Empty;
            this.LblPlayer4.Text = string.Empty;
            this.LblPlayer5.Text = string.Empty;
            this.LblPlayer6.Text = string.Empty;

            this.ReadSetFile();

        }

        private void BtnAddP3_Click(object sender, EventArgs e)
        {
            //プレイ人数指定
            this.PLAYERS = 3;

            //プレイヤー、ワード欄をリセット

            //プレイヤー、ワード欄を指定した人数
            this.TbxPlayer4.Visible = false;
            this.TbxPlayer5.Visible = false;
            this.TbxPlayer6.Visible = false;

            this.TxbP4Mail.Visible = false;
            this.TxbP5Mail.Visible = false;
            this.TxbP6Mail.Visible = false;

            this.LblP4Word.Visible = false;
            this.LblP5Word.Visible = false;
            this.LblP6Word.Visible = false;

            this.LblPlayer4.Visible = false;
            this.LblPlayer5.Visible = false;
            this.LblPlayer6.Visible = false;

            this.LblWin4.Visible = false;
            this.LblWin5.Visible = false;
            this.LblWin6.Visible = false;

            this.BtnUp4.Visible = false;
            this.BtnUp5.Visible = false;
            this.BtnUp6.Visible = false;

            this.BtnDown4.Visible = false;
            this.BtnDown5.Visible = false;
            this.BtnDown6.Visible = false;
        }

        private void BtnAddP4_Click(object sender, EventArgs e)
        {
            //プレイ人数指定
            this.PLAYERS = 4;

            //プレイヤー、ワード欄をリセット
            this.TbxPlayer4.Visible = true;
            this.TxbP4Mail.Visible = true;
            this.LblP4Word.Visible = true;
            this.LblPlayer4.Visible = true;
            this.LblWin4.Visible = true;
            this.BtnUp4.Visible = true;
            this.BtnDown4.Visible = true;

            //プレイヤー、ワード欄を指定した人数
            this.TbxPlayer5.Visible = false;
            this.TxbP5Mail.Visible = false;
            this.LblP5Word.Visible = false;
            this.LblPlayer5.Visible = false;
            this.LblWin5.Visible = false;
            this.BtnUp5.Visible = false;
            this.BtnDown5.Visible = false;

            //プレイヤー、ワード欄を指定した人数
            this.TbxPlayer6.Visible = false;
            this.TxbP6Mail.Visible = false;
            this.LblP6Word.Visible = false;
            this.LblPlayer6.Visible = false;
            this.LblWin6.Visible = false;
            this.BtnUp6.Visible = false;
            this.BtnDown6.Visible = false;
        }

        private void BtnAddP5_Click(object sender, EventArgs e)
        {
            //プレイ人数指定
            this.PLAYERS = 5;

            //プレイヤー、ワード欄をリセット
            this.TbxPlayer4.Visible = true;
            this.TbxPlayer5.Visible = true;
            this.TbxPlayer6.Visible = false;

            this.TxbP4Mail.Visible = true;
            this.TxbP5Mail.Visible = true;
            this.TxbP6Mail.Visible = false;

            this.LblP4Word.Visible = true;
            this.LblP5Word.Visible = true;
            this.LblP6Word.Visible = false;

            this.LblPlayer4.Visible = true;
            this.LblPlayer5.Visible = true;
            this.LblPlayer6.Visible = false;

            this.LblWin4.Visible = true;
            this.LblWin5.Visible = true;
            this.LblWin6.Visible = false;

            this.BtnUp4.Visible = true;
            this.BtnUp5.Visible = true;
            this.BtnUp6.Visible = false;

            this.BtnDown4.Visible = true;
            this.BtnDown5.Visible = true;
            this.BtnDown6.Visible = false;

            //プレイヤー、ワード欄を指定した人数

        }

        private void BtnAddP6_Click(object sender, EventArgs e)
        {
            //プレイ人数指定
            this.PLAYERS = 6;

            //プレイヤー、ワード欄をリセット
            this.TbxPlayer4.Visible = true;
            this.TbxPlayer5.Visible = true;
            this.TbxPlayer6.Visible = true;

            this.TxbP4Mail.Visible = true;
            this.TxbP5Mail.Visible = true;
            this.TxbP6Mail.Visible = true;

            this.LblP4Word.Visible = true;
            this.LblP5Word.Visible = true;
            this.LblP6Word.Visible = true;

            this.LblPlayer4.Visible = true;
            this.LblPlayer5.Visible = true;
            this.LblPlayer6.Visible = true;

            this.LblWin4.Visible = true;
            this.LblWin5.Visible = true;
            this.LblWin6.Visible = true;

            this.BtnUp4.Visible = true;
            this.BtnUp5.Visible = true;
            this.BtnUp6.Visible = true;

            this.BtnDown4.Visible = true;
            this.BtnDown5.Visible = true;
            this.BtnDown6.Visible = true;

            //プレイヤー、ワード欄を指定した人数
        }

        private void BtnWordOK_Click(object sender, EventArgs e)
        {
            this.WORDOK = true;

            if (this.RESULT == true)
            {
                if (this.WINCOUNT == true)
                {
                    this.LblP1Word.ForeColor = Color.Black;
                    this.LblP2Word.ForeColor = Color.Black;
                    this.LblP3Word.ForeColor = Color.Black;
                    this.LblP4Word.ForeColor = Color.Black;
                    this.LblP5Word.ForeColor = Color.Black;
                    this.LblP6Word.ForeColor = Color.Black;

                    this.RANDOMWORDNUM = this.WORDNUM.Next(1, 3);

                    if (RANDOMWORDNUM == 1)
                    {
                        this.SAFEWORD = this.TbxWord1.Text;
                        this.OUTWORD = this.TbxWord2.Text;
                    }
                    else
                    {
                        this.OUTWORD = this.TbxWord1.Text;
                        this.SAFEWORD = this.TbxWord2.Text;
                    }

                    if (this.PLAYERS == 3)
                    {
                        this.RANDOMWOLFNUM = this.WOLFNUM.Next(1, 4);

                        switch (this.RANDOMWOLFNUM)
                        {
                            case 1:

                                this.LblP1Word.Text = this.OUTWORD;
                                this.LblP2Word.Text = this.SAFEWORD;
                                this.LblP3Word.Text = this.SAFEWORD;

                                this.LblP1Word.ForeColor = Color.Red;

                                break;
                            case 2:

                                this.LblP1Word.Text = this.SAFEWORD;
                                this.LblP2Word.Text = this.OUTWORD;
                                this.LblP3Word.Text = this.SAFEWORD;

                                this.LblP2Word.ForeColor = Color.Red;

                                break;
                            case 3:

                                this.LblP1Word.Text = this.SAFEWORD;
                                this.LblP2Word.Text = this.SAFEWORD;
                                this.LblP3Word.Text = this.OUTWORD;

                                this.LblP3Word.ForeColor = Color.Red;

                                break;
                        }
                    }
                    else if (this.PLAYERS == 4)
                    {
                        this.RANDOMWOLFNUM = this.WOLFNUM.Next(1, 5);

                        switch (this.RANDOMWOLFNUM)
                        {
                            case 1:

                                this.LblP1Word.Text = this.OUTWORD;
                                this.LblP2Word.Text = this.SAFEWORD;
                                this.LblP3Word.Text = this.SAFEWORD;
                                this.LblP4Word.Text = this.SAFEWORD;

                                this.LblP1Word.ForeColor = Color.Red;

                                break;
                            case 2:

                                this.LblP1Word.Text = this.SAFEWORD;
                                this.LblP2Word.Text = this.OUTWORD;
                                this.LblP3Word.Text = this.SAFEWORD;
                                this.LblP4Word.Text = this.SAFEWORD;

                                this.LblP2Word.ForeColor = Color.Red;

                                break;
                            case 3:

                                this.LblP1Word.Text = this.SAFEWORD;
                                this.LblP2Word.Text = this.SAFEWORD;
                                this.LblP3Word.Text = this.OUTWORD;
                                this.LblP4Word.Text = this.SAFEWORD;

                                this.LblP3Word.ForeColor = Color.Red;

                                break;
                            case 4:

                                this.LblP1Word.Text = this.SAFEWORD;
                                this.LblP2Word.Text = this.SAFEWORD;
                                this.LblP3Word.Text = this.SAFEWORD;
                                this.LblP4Word.Text = this.OUTWORD;

                                this.LblP4Word.ForeColor = Color.Red;

                                break;
                        }
                    }
                    else if (this.PLAYERS == 5)
                    {
                        this.RANDOMWOLFNUM = this.WOLFNUM.Next(1, 6);

                        switch (this.RANDOMWOLFNUM)
                        {
                            case 1:

                                this.LblP1Word.Text = this.OUTWORD;
                                this.LblP2Word.Text = this.SAFEWORD;
                                this.LblP3Word.Text = this.SAFEWORD;
                                this.LblP4Word.Text = this.SAFEWORD;
                                this.LblP5Word.Text = this.SAFEWORD;

                                this.LblP1Word.ForeColor = Color.Red;

                                break;
                            case 2:

                                this.LblP1Word.Text = this.SAFEWORD;
                                this.LblP2Word.Text = this.OUTWORD;
                                this.LblP3Word.Text = this.SAFEWORD;
                                this.LblP4Word.Text = this.SAFEWORD;
                                this.LblP5Word.Text = this.SAFEWORD;

                                this.LblP2Word.ForeColor = Color.Red;

                                break;
                            case 3:

                                this.LblP1Word.Text = this.SAFEWORD;
                                this.LblP2Word.Text = this.SAFEWORD;
                                this.LblP3Word.Text = this.OUTWORD;
                                this.LblP4Word.Text = this.SAFEWORD;
                                this.LblP5Word.Text = this.SAFEWORD;

                                this.LblP3Word.ForeColor = Color.Red;

                                break;
                            case 4:

                                this.LblP1Word.Text = this.SAFEWORD;
                                this.LblP2Word.Text = this.SAFEWORD;
                                this.LblP3Word.Text = this.SAFEWORD;
                                this.LblP4Word.Text = this.OUTWORD;
                                this.LblP5Word.Text = this.SAFEWORD;

                                this.LblP4Word.ForeColor = Color.Red;

                                break;
                            case 5:

                                this.LblP1Word.Text = this.SAFEWORD;
                                this.LblP2Word.Text = this.SAFEWORD;
                                this.LblP3Word.Text = this.SAFEWORD;
                                this.LblP4Word.Text = this.SAFEWORD;
                                this.LblP5Word.Text = this.OUTWORD;

                                this.LblP5Word.ForeColor = Color.Red;

                                break;
                        }

                    }
                    else if (this.PLAYERS == 6)
                    {
                        this.RANDOMWOLFNUM = this.WOLFNUM.Next(1, 7);

                        switch (this.RANDOMWOLFNUM)
                        {
                            case 1:

                                this.LblP1Word.Text = this.OUTWORD;
                                this.LblP2Word.Text = this.SAFEWORD;
                                this.LblP3Word.Text = this.SAFEWORD;
                                this.LblP4Word.Text = this.SAFEWORD;
                                this.LblP5Word.Text = this.SAFEWORD;
                                this.LblP6Word.Text = this.SAFEWORD;

                                this.LblP1Word.ForeColor = Color.Red;

                                break;
                            case 2:

                                this.LblP1Word.Text = this.SAFEWORD;
                                this.LblP2Word.Text = this.OUTWORD;
                                this.LblP3Word.Text = this.SAFEWORD;
                                this.LblP4Word.Text = this.SAFEWORD;
                                this.LblP5Word.Text = this.SAFEWORD;
                                this.LblP6Word.Text = this.SAFEWORD;

                                this.LblP2Word.ForeColor = Color.Red;

                                break;
                            case 3:

                                this.LblP1Word.Text = this.SAFEWORD;
                                this.LblP2Word.Text = this.SAFEWORD;
                                this.LblP3Word.Text = this.OUTWORD;
                                this.LblP4Word.Text = this.SAFEWORD;
                                this.LblP5Word.Text = this.SAFEWORD;
                                this.LblP6Word.Text = this.SAFEWORD;

                                this.LblP3Word.ForeColor = Color.Red;

                                break;
                            case 4:

                                this.LblP1Word.Text = this.SAFEWORD;
                                this.LblP2Word.Text = this.SAFEWORD;
                                this.LblP3Word.Text = this.SAFEWORD;
                                this.LblP4Word.Text = this.OUTWORD;
                                this.LblP5Word.Text = this.SAFEWORD;
                                this.LblP6Word.Text = this.SAFEWORD;

                                this.LblP4Word.ForeColor = Color.Red;

                                break;
                            case 5:

                                this.LblP1Word.Text = this.SAFEWORD;
                                this.LblP2Word.Text = this.SAFEWORD;
                                this.LblP3Word.Text = this.SAFEWORD;
                                this.LblP4Word.Text = this.SAFEWORD;
                                this.LblP5Word.Text = this.OUTWORD;
                                this.LblP6Word.Text = this.SAFEWORD;

                                this.LblP5Word.ForeColor = Color.Red;

                                break;
                            case 6:

                                this.LblP1Word.Text = this.SAFEWORD;
                                this.LblP2Word.Text = this.SAFEWORD;
                                this.LblP3Word.Text = this.SAFEWORD;
                                this.LblP4Word.Text = this.SAFEWORD;
                                this.LblP5Word.Text = this.SAFEWORD;
                                this.LblP6Word.Text = this.OUTWORD;

                                this.LblP6Word.ForeColor = Color.Red;

                                break;
                        }



                    }
                }
                else
                {
                    //メッセージボックスを表示する
                    MessageBox.Show("勝利数を更新してください。","確認",MessageBoxButtons.OK,MessageBoxIcon.Information);
                }
            }
            else
            {
                this.LblP1Word.ForeColor = Color.Black;
                this.LblP2Word.ForeColor = Color.Black;
                this.LblP3Word.ForeColor = Color.Black;
                this.LblP4Word.ForeColor = Color.Black;
                this.LblP5Word.ForeColor = Color.Black;
                this.LblP6Word.ForeColor = Color.Black;

                this.RANDOMWORDNUM = this.WORDNUM.Next(1, 3);

                if (RANDOMWORDNUM == 1)
                {
                    this.SAFEWORD = this.TbxWord1.Text;
                    this.OUTWORD = this.TbxWord2.Text;
                }
                else
                {
                    this.OUTWORD = this.TbxWord1.Text;
                    this.SAFEWORD = this.TbxWord2.Text;
                }

                if (this.PLAYERS == 3)
                {
                    this.RANDOMWOLFNUM = this.WOLFNUM.Next(1, 4);

                    switch (this.RANDOMWOLFNUM)
                    {
                        case 1:

                            this.LblP1Word.Text = this.OUTWORD;
                            this.LblP2Word.Text = this.SAFEWORD;
                            this.LblP3Word.Text = this.SAFEWORD;

                            this.LblP1Word.ForeColor = Color.Red;

                            break;
                        case 2:

                            this.LblP1Word.Text = this.SAFEWORD;
                            this.LblP2Word.Text = this.OUTWORD;
                            this.LblP3Word.Text = this.SAFEWORD;

                            this.LblP2Word.ForeColor = Color.Red;

                            break;
                        case 3:

                            this.LblP1Word.Text = this.SAFEWORD;
                            this.LblP2Word.Text = this.SAFEWORD;
                            this.LblP3Word.Text = this.OUTWORD;

                            this.LblP3Word.ForeColor = Color.Red;

                            break;
                    }
                }
                else if (this.PLAYERS == 4)
                {
                    this.RANDOMWOLFNUM = this.WOLFNUM.Next(1, 5);

                    switch (this.RANDOMWOLFNUM)
                    {
                        case 1:

                            this.LblP1Word.Text = this.OUTWORD;
                            this.LblP2Word.Text = this.SAFEWORD;
                            this.LblP3Word.Text = this.SAFEWORD;
                            this.LblP4Word.Text = this.SAFEWORD;

                            this.LblP1Word.ForeColor = Color.Red;

                            break;
                        case 2:

                            this.LblP1Word.Text = this.SAFEWORD;
                            this.LblP2Word.Text = this.OUTWORD;
                            this.LblP3Word.Text = this.SAFEWORD;
                            this.LblP4Word.Text = this.SAFEWORD;

                            this.LblP2Word.ForeColor = Color.Red;

                            break;
                        case 3:

                            this.LblP1Word.Text = this.SAFEWORD;
                            this.LblP2Word.Text = this.SAFEWORD;
                            this.LblP3Word.Text = this.OUTWORD;
                            this.LblP4Word.Text = this.SAFEWORD;

                            this.LblP3Word.ForeColor = Color.Red;

                            break;
                        case 4:

                            this.LblP1Word.Text = this.SAFEWORD;
                            this.LblP2Word.Text = this.SAFEWORD;
                            this.LblP3Word.Text = this.SAFEWORD;
                            this.LblP4Word.Text = this.OUTWORD;

                            this.LblP4Word.ForeColor = Color.Red;

                            break;
                    }
                }
                else if (this.PLAYERS == 5)
                {
                    this.RANDOMWOLFNUM = this.WOLFNUM.Next(1, 6);

                    switch (this.RANDOMWOLFNUM)
                    {
                        case 1:

                            this.LblP1Word.Text = this.OUTWORD;
                            this.LblP2Word.Text = this.SAFEWORD;
                            this.LblP3Word.Text = this.SAFEWORD;
                            this.LblP4Word.Text = this.SAFEWORD;
                            this.LblP5Word.Text = this.SAFEWORD;

                            this.LblP1Word.ForeColor = Color.Red;

                            break;
                        case 2:

                            this.LblP1Word.Text = this.SAFEWORD;
                            this.LblP2Word.Text = this.OUTWORD;
                            this.LblP3Word.Text = this.SAFEWORD;
                            this.LblP4Word.Text = this.SAFEWORD;
                            this.LblP5Word.Text = this.SAFEWORD;

                            this.LblP2Word.ForeColor = Color.Red;

                            break;
                        case 3:

                            this.LblP1Word.Text = this.SAFEWORD;
                            this.LblP2Word.Text = this.SAFEWORD;
                            this.LblP3Word.Text = this.OUTWORD;
                            this.LblP4Word.Text = this.SAFEWORD;
                            this.LblP5Word.Text = this.SAFEWORD;

                            this.LblP3Word.ForeColor = Color.Red;

                            break;
                        case 4:

                            this.LblP1Word.Text = this.SAFEWORD;
                            this.LblP2Word.Text = this.SAFEWORD;
                            this.LblP3Word.Text = this.SAFEWORD;
                            this.LblP4Word.Text = this.OUTWORD;
                            this.LblP5Word.Text = this.SAFEWORD;

                            this.LblP4Word.ForeColor = Color.Red;

                            break;
                        case 5:

                            this.LblP1Word.Text = this.SAFEWORD;
                            this.LblP2Word.Text = this.SAFEWORD;
                            this.LblP3Word.Text = this.SAFEWORD;
                            this.LblP4Word.Text = this.SAFEWORD;
                            this.LblP5Word.Text = this.OUTWORD;

                            this.LblP5Word.ForeColor = Color.Red;

                            break;
                    }

                }
                else if (this.PLAYERS == 6)
                {
                    this.RANDOMWOLFNUM = this.WOLFNUM.Next(1, 7);

                    switch (this.RANDOMWOLFNUM)
                    {
                        case 1:

                            this.LblP1Word.Text = this.OUTWORD;
                            this.LblP2Word.Text = this.SAFEWORD;
                            this.LblP3Word.Text = this.SAFEWORD;
                            this.LblP4Word.Text = this.SAFEWORD;
                            this.LblP5Word.Text = this.SAFEWORD;
                            this.LblP6Word.Text = this.SAFEWORD;

                            this.LblP1Word.ForeColor = Color.Red;

                            break;
                        case 2:

                            this.LblP1Word.Text = this.SAFEWORD;
                            this.LblP2Word.Text = this.OUTWORD;
                            this.LblP3Word.Text = this.SAFEWORD;
                            this.LblP4Word.Text = this.SAFEWORD;
                            this.LblP5Word.Text = this.SAFEWORD;
                            this.LblP6Word.Text = this.SAFEWORD;

                            this.LblP2Word.ForeColor = Color.Red;

                            break;
                        case 3:

                            this.LblP1Word.Text = this.SAFEWORD;
                            this.LblP2Word.Text = this.SAFEWORD;
                            this.LblP3Word.Text = this.OUTWORD;
                            this.LblP4Word.Text = this.SAFEWORD;
                            this.LblP5Word.Text = this.SAFEWORD;
                            this.LblP6Word.Text = this.SAFEWORD;

                            this.LblP3Word.ForeColor = Color.Red;

                            break;
                        case 4:

                            this.LblP1Word.Text = this.SAFEWORD;
                            this.LblP2Word.Text = this.SAFEWORD;
                            this.LblP3Word.Text = this.SAFEWORD;
                            this.LblP4Word.Text = this.OUTWORD;
                            this.LblP5Word.Text = this.SAFEWORD;
                            this.LblP6Word.Text = this.SAFEWORD;

                            this.LblP4Word.ForeColor = Color.Red;

                            break;
                        case 5:

                            this.LblP1Word.Text = this.SAFEWORD;
                            this.LblP2Word.Text = this.SAFEWORD;
                            this.LblP3Word.Text = this.SAFEWORD;
                            this.LblP4Word.Text = this.SAFEWORD;
                            this.LblP5Word.Text = this.OUTWORD;
                            this.LblP6Word.Text = this.SAFEWORD;

                            this.LblP5Word.ForeColor = Color.Red;

                            break;
                        case 6:

                            this.LblP1Word.Text = this.SAFEWORD;
                            this.LblP2Word.Text = this.SAFEWORD;
                            this.LblP3Word.Text = this.SAFEWORD;
                            this.LblP4Word.Text = this.SAFEWORD;
                            this.LblP5Word.Text = this.SAFEWORD;
                            this.LblP6Word.Text = this.OUTWORD;

                            this.LblP6Word.ForeColor = Color.Red;

                            break;
                    }



                }
            }

            
        }

        private void BtnPlayerOK_Click(object sender, EventArgs e)
        {
            this.LblPlayer1.Text = this.TbxPlayer1.Text;

            this.LblPlayer2.Text = this.TbxPlayer2.Text;

            this.LblPlayer3.Text = this.TbxPlayer3.Text;

            this.LblPlayer4.Text = this.TbxPlayer4.Text;

            this.LblPlayer5.Text = this.TbxPlayer5.Text;

            this.LblPlayer6.Text = this.TbxPlayer6.Text;
        }

        private void BtnUp1_Click(object sender, EventArgs e)
        {
            //文字数を取得
            int len = LblWin1.Text.Length;

            string strNum = this.LblWin1.Text.Substring(0, len - 1);

            int winNum = int.Parse(strNum) + 1;

            string strWin = winNum.ToString() + "勝";

            this.LblWin1.Text = strWin;

            this.WINCOUNT = true;
        }

        private void BtnUp2_Click(object sender, EventArgs e)
        {
            //文字数を取得
            int len = LblWin2.Text.Length;

            string strNum = this.LblWin2.Text.Substring(0, len - 1);

            int winNum = int.Parse(strNum) + 1;

            string strWin = winNum.ToString() + "勝";

            this.LblWin2.Text = strWin;

            this.WINCOUNT = true;
        }

        private void BtnUp3_Click(object sender, EventArgs e)
        {
            //文字数を取得
            int len = LblWin3.Text.Length;

            string strNum = this.LblWin3.Text.Substring(0, len - 1);

            int winNum = int.Parse(strNum) + 1;

            string strWin = winNum.ToString() + "勝";

            this.LblWin3.Text = strWin;

            this.WINCOUNT = true;
        }

        private void BtnUp4_Click(object sender, EventArgs e)
        {
            //文字数を取得
            int len = LblWin4.Text.Length;

            string strNum = this.LblWin4.Text.Substring(0, len - 1);

            int winNum = int.Parse(strNum) + 1;

            string strWin = winNum.ToString() + "勝";

            this.LblWin4.Text = strWin;

            this.WINCOUNT = true;
        }

        private void BtnUp5_Click(object sender, EventArgs e)
        {
            //文字数を取得
            int len = LblWin5.Text.Length;

            string strNum = this.LblWin5.Text.Substring(0, len - 1);

            int winNum = int.Parse(strNum) + 1;

            string strWin = winNum.ToString() + "勝";

            this.LblWin5.Text = strWin;

            this.WINCOUNT = true;
        }

        private void BtnUp6_Click(object sender, EventArgs e)
        {
            //文字数を取得
            int len = LblWin6.Text.Length;

            string strNum = this.LblWin6.Text.Substring(0, len - 1);

            int winNum = int.Parse(strNum) + 1;

            string strWin = winNum.ToString() + "勝";

            this.LblWin6.Text = strWin;

            this.WINCOUNT = true;
        }

        private void BtnDown1_Click(object sender, EventArgs e)
        {
            //文字数を取得
            int len = LblWin1.Text.Length;

            string strNum = this.LblWin1.Text.Substring(0, len - 1);

            int winNum = int.Parse(strNum) - 1;

            string strWin = winNum.ToString() + "勝";

            this.LblWin1.Text = strWin;

            this.WINCOUNT = true;
        }

        private void BtnDown2_Click(object sender, EventArgs e)
        {
            //文字数を取得
            int len = LblWin2.Text.Length;

            string strNum = this.LblWin2.Text.Substring(0, len - 1);

            int winNum = int.Parse(strNum) - 1;

            string strWin = winNum.ToString() + "勝";

            this.LblWin2.Text = strWin;

            this.WINCOUNT = true;
        }

        private void BtnDown3_Click(object sender, EventArgs e)
        {
            //文字数を取得
            int len = LblWin3.Text.Length;

            string strNum = this.LblWin3.Text.Substring(0, len - 1);

            int winNum = int.Parse(strNum) - 1;

            string strWin = winNum.ToString() + "勝";

            this.LblWin3.Text = strWin;

            this.WINCOUNT = true;
        }

        private void BtnDown4_Click(object sender, EventArgs e)
        {
            //文字数を取得
            int len = LblWin4.Text.Length;

            string strNum = this.LblWin4.Text.Substring(0, len - 1);

            int winNum = int.Parse(strNum) - 1;

            string strWin = winNum.ToString() + "勝";

            this.LblWin4.Text = strWin;

            this.WINCOUNT = true;
        }

        private void BtnDown5_Click(object sender, EventArgs e)
        {
            //文字数を取得
            int len = LblWin5.Text.Length;

            string strNum = this.LblWin5.Text.Substring(0, len - 1);

            int winNum = int.Parse(strNum) - 1;

            string strWin = winNum.ToString() + "勝";

            this.LblWin5.Text = strWin;

            this.WINCOUNT = true;
        }

        private void BtnDown6_Click(object sender, EventArgs e)
        {
            //文字数を取得
            int len = LblWin6.Text.Length;

            string strNum = this.LblWin6.Text.Substring(0, len - 1);

            int winNum = int.Parse(strNum) - 1;

            string strWin = winNum.ToString() + "勝";

            this.LblWin6.Text = strWin;

            this.WINCOUNT = true;
        }

        private void BtnTimerStart_Click(object sender, EventArgs e)
        {
            this.Timer = new System.Windows.Forms.Timer();
            this.Timer.Tick += new EventHandler(CountDown);

            this.COUNTTIME = this.MASTERCOUNTTIME;

            this.ProgBarTimer.Maximum = this.COUNTTIME;
            this.ProgBarTimer.Value = this.COUNTTIME;

            this.Timer.Interval = 1000;
            this.Timer.Start();

            this.WINCOUNT = false;
        }

        private void CountDown(object sender, EventArgs e)
        {

            if (this.COUNTTIME == 0)
            {
                Timer.Stop();

                this.LblTimer.ForeColor = Color.Black;

                this.LblTimer.ForeColor = Color.Red;
            }
            else if (this.COUNTTIME > 0)
            {
                this.COUNTTIME--;

                this.ProgBarTimer.Value = this.COUNTTIME;

                if (this.COUNTTIME > 60)
                {
                    this.LblTimer.ForeColor = Color.Black;

                    this.ProgBarTimer.ForeColor = Color.PaleGreen;
                }
                else
                {
                    this.LblTimer.ForeColor = Color.Red;

                    this.ProgBarTimer.ForeColor = Color.Red;
                }
                
                this.LblTimer.Text = this.COUNTTIME.ToString();
            }
        }

        private void BtnTimer3m_Click(object sender, EventArgs e)
        {
            this.MASTERCOUNTTIME = 180;
            this.COUNTTIME = this.MASTERCOUNTTIME;
            this.LblTimer.Text = this.COUNTTIME.ToString();
        }

        private void BtnTimer5m_Click(object sender, EventArgs e)
        {
            this.MASTERCOUNTTIME = 300;
            this.COUNTTIME = this.MASTERCOUNTTIME;
            this.LblTimer.Text = this.COUNTTIME.ToString();
        }

        private void BtnTimer7m_Click(object sender, EventArgs e)
        {
            this.MASTERCOUNTTIME = 420;
            this.COUNTTIME = this.MASTERCOUNTTIME;
            this.LblTimer.Text = this.COUNTTIME.ToString();
        }

        private void BtnTimer10m_Click(object sender, EventArgs e)
        {
            this.MASTERCOUNTTIME = 600;
            this.COUNTTIME = this.MASTERCOUNTTIME;
            this.LblTimer.Text = this.COUNTTIME.ToString();
        }

        private void BtnSet_Click(object sender, EventArgs e)
        {
            if(this.PSETSTATE == true)
            {
                this.PSETSTATE = false;

                this.TbxPlayer1.Visible = false;
                this.TbxPlayer2.Visible = false;
                this.TbxPlayer3.Visible = false;
                this.TbxPlayer4.Visible = false;
                this.TbxPlayer5.Visible = false;
                this.TbxPlayer6.Visible = false;

                this.TxbP1Mail.Visible = false;
                this.TxbP2Mail.Visible = false;
                this.TxbP3Mail.Visible = false;
                this.TxbP4Mail.Visible = false;
                this.TxbP5Mail.Visible = false;
                this.TxbP6Mail.Visible = false;

                this.BtnAddP3.Visible = false;
                this.BtnAddP4.Visible = false;
                this.BtnAddP5.Visible = false;
                this.BtnAddP6.Visible = false;

                this.BtnTimer3m.Visible = false;
                this.BtnTimer5m.Visible = false;
                this.BtnTimer7m.Visible = false;
                this.BtnTimer10m.Visible = false;

                this.BtnPlayerOK.Visible = false;

                this.BtnSet.Location = new Point(1227, 12);

                this.MinimumSize = new Size(1300, 350);

                this.Height = this.Height - 100;
            }
            else
            {
                this.PSETSTATE = true;

                this.TbxPlayer1.Visible = true;
                this.TbxPlayer2.Visible = true;
                this.TbxPlayer3.Visible = true;
                this.TbxPlayer4.Visible = true;
                this.TbxPlayer5.Visible = true;
                this.TbxPlayer6.Visible = true;

                this.TxbP1Mail.Visible = true;
                this.TxbP2Mail.Visible = true;
                this.TxbP3Mail.Visible = true;
                this.TxbP4Mail.Visible = true;
                this.TxbP5Mail.Visible = true;
                this.TxbP6Mail.Visible = true;

                this.BtnAddP3.Visible = true;
                this.BtnAddP4.Visible = true;
                this.BtnAddP5.Visible = true;
                this.BtnAddP6.Visible = true;

                this.BtnTimer3m.Visible = true;
                this.BtnTimer5m.Visible = true;
                this.BtnTimer7m.Visible = true;
                this.BtnTimer10m.Visible = true;

                this.BtnPlayerOK.Visible = true;

                this.BtnSet.Location = new Point(615, 12);

                this.MinimumSize = new Size(1300, 500);

                this.Height = this.Height + 100;
            }

            
        }

        private void BtnResultOn_Click(object sender, EventArgs e)
        {
            if(this.RESULT == true)
            {
                this.RESULT = false;
                this.BtnResultOn.Text = "結果OFF";
            }
            else
            {
                this.RESULT = true;
                this.BtnResultOn.Text = "結果ON";
            }
        }

        private void BtnSendMail_Click(object sender, EventArgs e)
        {
            if(this.WORDOK == true)
            {
                switch (this.PLAYERS)
                {
                    case 3:
                        if (TxbP1Mail.Text != string.Empty && TxbP2Mail.Text != string.Empty &&
                            TxbP3Mail.Text != string.Empty)
                        {
                            if (TxbP1Mail.Text.Contains("@gmail.com") && TxbP2Mail.Text.Contains("@gmail.com") &&
                                TxbP3Mail.Text.Contains("@gmail.com"))
                            {
                                var msg1 = new MimeKit.MimeMessage();
                                msg1.From.Add(new MimeKit.MailboxAddress("GM", this.USERNAME));
                                msg1.To.Add(new MimeKit.MailboxAddress("P1", TxbP1Mail.Text));
                                msg1.Subject = "ワードウルフのお題";
                                var text1 = new MimeKit.TextPart("Plain");
                                text1.Text = LblP1Word.Text;
                                msg1.Body = text1;

                                var msg2 = new MimeKit.MimeMessage();
                                msg2.From.Add(new MimeKit.MailboxAddress("GM", this.USERNAME));
                                msg2.To.Add(new MimeKit.MailboxAddress("P2", TxbP2Mail.Text));
                                msg2.Subject = "ワードウルフのお題";
                                var text2 = new MimeKit.TextPart("Plain");
                                text2.Text = LblP2Word.Text;
                                msg2.Body = text2;

                                var msg3 = new MimeKit.MimeMessage();
                                msg3.From.Add(new MimeKit.MailboxAddress("GM", this.USERNAME));
                                msg3.To.Add(new MimeKit.MailboxAddress("P3", TxbP3Mail.Text));
                                msg3.Subject = "ワードウルフのお題";
                                var text3 = new MimeKit.TextPart("Plain");
                                text3.Text = LblP3Word.Text;
                                msg3.Body = text3;

                                using (var client = new MailKit.Net.Smtp.SmtpClient())
                                {
                                    try
                                    {
                                        client.Connect(this.MAILHOST, this.MAILPORT, MailKit.Security.SecureSocketOptions.Auto);
                                        client.Authenticate(this.USERNAME, this.PASSWORD);
                                        client.Send(msg1);
                                        client.Disconnect(true);

                                        client.Connect(this.MAILHOST, this.MAILPORT, MailKit.Security.SecureSocketOptions.Auto);
                                        client.Authenticate(this.USERNAME, this.PASSWORD);
                                        client.Send(msg2);
                                        client.Disconnect(true);

                                        client.Connect(this.MAILHOST, this.MAILPORT, MailKit.Security.SecureSocketOptions.Auto);
                                        client.Authenticate(this.USERNAME, this.PASSWORD);
                                        client.Send(msg3);
                                        client.Disconnect(true);
                                    }
                                    catch
                                    {

                                    }
                                }
                            }
                            else
                            {
                                //メッセージボックスを表示する
                                MessageBox.Show("メールアドレスが間違ってるよ", "確認",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        else
                        {
                            //メッセージボックスを表示する
                            MessageBox.Show("メールアドレスが設定されてないよ", "確認",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        break;
                    case 4:
                        if (TxbP1Mail.Text != string.Empty && TxbP2Mail.Text != string.Empty &&
                            TxbP3Mail.Text != string.Empty && TxbP4Mail.Text != string.Empty)
                        {
                            if (TxbP1Mail.Text.Contains("@gmail.com") && TxbP2Mail.Text.Contains("@gmail.com") &&
                                TxbP3Mail.Text.Contains("@gmail.com") && TxbP4Mail.Text.Contains("@gmail.com"))
                            {
                                var msg1 = new MimeKit.MimeMessage();
                                msg1.From.Add(new MimeKit.MailboxAddress("GM", this.USERNAME));
                                msg1.To.Add(new MimeKit.MailboxAddress("P1", TxbP1Mail.Text));
                                msg1.Subject = "ワードウルフのお題";
                                var text1 = new MimeKit.TextPart("Plain");
                                text1.Text = LblP1Word.Text;
                                msg1.Body = text1;

                                var msg2 = new MimeKit.MimeMessage();
                                msg2.From.Add(new MimeKit.MailboxAddress("GM", this.USERNAME));
                                msg2.To.Add(new MimeKit.MailboxAddress("P2", TxbP2Mail.Text));
                                msg2.Subject = "ワードウルフのお題";
                                var text2 = new MimeKit.TextPart("Plain");
                                text2.Text = LblP2Word.Text;
                                msg2.Body = text2;

                                var msg3 = new MimeKit.MimeMessage();
                                msg3.From.Add(new MimeKit.MailboxAddress("GM", this.USERNAME));
                                msg3.To.Add(new MimeKit.MailboxAddress("P3", TxbP3Mail.Text));
                                msg3.Subject = "ワードウルフのお題";
                                var text3 = new MimeKit.TextPart("Plain");
                                text3.Text = LblP3Word.Text;
                                msg3.Body = text3;

                                var msg4 = new MimeKit.MimeMessage();
                                msg4.From.Add(new MimeKit.MailboxAddress("GM", this.USERNAME));
                                msg4.To.Add(new MimeKit.MailboxAddress("P4", TxbP4Mail.Text));
                                msg4.Subject = "ワードウルフのお題";
                                var text4 = new MimeKit.TextPart("Plain");
                                text4.Text = LblP4Word.Text;
                                msg4.Body = text4;

                                using (var client = new MailKit.Net.Smtp.SmtpClient())
                                {
                                    try
                                    {
                                        client.Connect(this.MAILHOST, this.MAILPORT, MailKit.Security.SecureSocketOptions.Auto);
                                        client.Authenticate(this.USERNAME, this.PASSWORD);
                                        client.Send(msg1);
                                        client.Disconnect(true);

                                        client.Connect(this.MAILHOST, this.MAILPORT, MailKit.Security.SecureSocketOptions.Auto);
                                        client.Authenticate(this.USERNAME, this.PASSWORD);
                                        client.Send(msg2);
                                        client.Disconnect(true);

                                        client.Connect(this.MAILHOST, this.MAILPORT, MailKit.Security.SecureSocketOptions.Auto);
                                        client.Authenticate(this.USERNAME, this.PASSWORD);
                                        client.Send(msg3);
                                        client.Disconnect(true);

                                        client.Connect(this.MAILHOST, this.MAILPORT, MailKit.Security.SecureSocketOptions.Auto);
                                        client.Authenticate(this.USERNAME, this.PASSWORD);
                                        client.Send(msg4);
                                        client.Disconnect(true);
                                    }
                                    catch
                                    {

                                    }
                                }
                            }
                            else
                            {
                                //メッセージボックスを表示する
                                MessageBox.Show("メールアドレスが間違ってるよ", "確認",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        else
                        {
                            //メッセージボックスを表示する
                            MessageBox.Show("メールアドレスが設定されてないよ", "確認",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        break;
                    case 5:
                        if (TxbP1Mail.Text != string.Empty && TxbP2Mail.Text != string.Empty &&
                            TxbP3Mail.Text != string.Empty && TxbP4Mail.Text != string.Empty &&
                            TxbP5Mail.Text != string.Empty)
                        {
                            if (TxbP1Mail.Text.Contains("@gmail.com") && TxbP2Mail.Text.Contains("@gmail.com") &&
                                TxbP3Mail.Text.Contains("@gmail.com") && TxbP4Mail.Text.Contains("@gmail.com") &&
                                TxbP5Mail.Text.Contains("@gmail.com"))
                            {
                                var msg1 = new MimeKit.MimeMessage();
                                msg1.From.Add(new MimeKit.MailboxAddress("GM", this.USERNAME));
                                msg1.To.Add(new MimeKit.MailboxAddress("P1", TxbP1Mail.Text));
                                msg1.Subject = "ワードウルフのお題";
                                var text1 = new MimeKit.TextPart("Plain");
                                text1.Text = LblP1Word.Text;
                                msg1.Body = text1;

                                var msg2 = new MimeKit.MimeMessage();
                                msg2.From.Add(new MimeKit.MailboxAddress("GM", this.USERNAME));
                                msg2.To.Add(new MimeKit.MailboxAddress("P2", TxbP2Mail.Text));
                                msg2.Subject = "ワードウルフのお題";
                                var text2 = new MimeKit.TextPart("Plain");
                                text2.Text = LblP2Word.Text;
                                msg2.Body = text2;

                                var msg3 = new MimeKit.MimeMessage();
                                msg3.From.Add(new MimeKit.MailboxAddress("GM", this.USERNAME));
                                msg3.To.Add(new MimeKit.MailboxAddress("P3", TxbP3Mail.Text));
                                msg3.Subject = "ワードウルフのお題";
                                var text3 = new MimeKit.TextPart("Plain");
                                text3.Text = LblP3Word.Text;
                                msg3.Body = text3;

                                var msg4 = new MimeKit.MimeMessage();
                                msg4.From.Add(new MimeKit.MailboxAddress("GM", this.USERNAME));
                                msg4.To.Add(new MimeKit.MailboxAddress("P4", TxbP4Mail.Text));
                                msg4.Subject = "ワードウルフのお題";
                                var text4 = new MimeKit.TextPart("Plain");
                                text4.Text = LblP4Word.Text;
                                msg4.Body = text4;

                                var msg5 = new MimeKit.MimeMessage();
                                msg5.From.Add(new MimeKit.MailboxAddress("GM", this.USERNAME));
                                msg5.To.Add(new MimeKit.MailboxAddress("P5", TxbP5Mail.Text));
                                msg5.Subject = "ワードウルフのお題";
                                var text5 = new MimeKit.TextPart("Plain");
                                text5.Text = LblP5Word.Text;
                                msg5.Body = text5;

                                using (var client = new MailKit.Net.Smtp.SmtpClient())
                                {
                                    try
                                    {
                                        client.Connect(this.MAILHOST, this.MAILPORT, MailKit.Security.SecureSocketOptions.Auto);
                                        client.Authenticate(this.USERNAME, this.PASSWORD);
                                        client.Send(msg1);
                                        client.Disconnect(true);

                                        client.Connect(this.MAILHOST, this.MAILPORT, MailKit.Security.SecureSocketOptions.Auto);
                                        client.Authenticate(this.USERNAME, this.PASSWORD);
                                        client.Send(msg2);
                                        client.Disconnect(true);

                                        client.Connect(this.MAILHOST, this.MAILPORT, MailKit.Security.SecureSocketOptions.Auto);
                                        client.Authenticate(this.USERNAME, this.PASSWORD);
                                        client.Send(msg3);
                                        client.Disconnect(true);

                                        client.Connect(this.MAILHOST, this.MAILPORT, MailKit.Security.SecureSocketOptions.Auto);
                                        client.Authenticate(this.USERNAME, this.PASSWORD);
                                        client.Send(msg4);
                                        client.Disconnect(true);

                                        client.Connect(this.MAILHOST, this.MAILPORT, MailKit.Security.SecureSocketOptions.Auto);
                                        client.Authenticate(this.USERNAME, this.PASSWORD);
                                        client.Send(msg5);
                                        client.Disconnect(true);
                                    }
                                    catch
                                    {

                                    }
                                }
                            }
                            else
                            {
                                //メッセージボックスを表示する
                                MessageBox.Show("メールアドレスが間違ってるよ", "確認",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        else
                        {
                            //メッセージボックスを表示する
                            MessageBox.Show("メールアドレスが設定されてないよ", "確認",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        break;
                    case 6:
                        if (TxbP1Mail.Text != string.Empty && TxbP2Mail.Text != string.Empty &&
                            TxbP3Mail.Text != string.Empty && TxbP4Mail.Text != string.Empty &&
                            TxbP5Mail.Text != string.Empty && TxbP6Mail.Text != string.Empty)
                        {
                            if (TxbP1Mail.Text.Contains("@gmail.com") && TxbP2Mail.Text.Contains("@gmail.com") &&
                                TxbP3Mail.Text.Contains("@gmail.com") && TxbP4Mail.Text.Contains("@gmail.com") &&
                                TxbP5Mail.Text.Contains("@gmail.com") && TxbP6Mail.Text.Contains("@gmail.com"))
                            {
                                var msg1 = new MimeKit.MimeMessage();
                                msg1.From.Add(new MimeKit.MailboxAddress("GM", this.USERNAME));
                                msg1.To.Add(new MimeKit.MailboxAddress("P1", TxbP1Mail.Text));
                                msg1.Subject = "ワードウルフのお題";
                                var text1 = new MimeKit.TextPart("Plain");
                                text1.Text = LblP1Word.Text;
                                msg1.Body = text1;

                                var msg2 = new MimeKit.MimeMessage();
                                msg2.From.Add(new MimeKit.MailboxAddress("GM", this.USERNAME));
                                msg2.To.Add(new MimeKit.MailboxAddress("P2", TxbP2Mail.Text));
                                msg2.Subject = "ワードウルフのお題";
                                var text2 = new MimeKit.TextPart("Plain");
                                text2.Text = LblP2Word.Text;
                                msg2.Body = text2;

                                var msg3 = new MimeKit.MimeMessage();
                                msg3.From.Add(new MimeKit.MailboxAddress("GM", this.USERNAME));
                                msg3.To.Add(new MimeKit.MailboxAddress("P3", TxbP3Mail.Text));
                                msg3.Subject = "ワードウルフのお題";
                                var text3 = new MimeKit.TextPart("Plain");
                                text3.Text = LblP3Word.Text;
                                msg3.Body = text3;

                                var msg4 = new MimeKit.MimeMessage();
                                msg4.From.Add(new MimeKit.MailboxAddress("GM", this.USERNAME));
                                msg4.To.Add(new MimeKit.MailboxAddress("P4", TxbP4Mail.Text));
                                msg4.Subject = "ワードウルフのお題";
                                var text4 = new MimeKit.TextPart("Plain");
                                text4.Text = LblP4Word.Text;
                                msg4.Body = text4;

                                var msg5 = new MimeKit.MimeMessage();
                                msg5.From.Add(new MimeKit.MailboxAddress("GM", this.USERNAME));
                                msg5.To.Add(new MimeKit.MailboxAddress("P5", TxbP5Mail.Text));
                                msg5.Subject = "ワードウルフのお題";
                                var text5 = new MimeKit.TextPart("Plain");
                                text5.Text = LblP5Word.Text;
                                msg5.Body = text5;

                                var msg6 = new MimeKit.MimeMessage();
                                msg6.From.Add(new MimeKit.MailboxAddress("GM", this.USERNAME));
                                msg6.To.Add(new MimeKit.MailboxAddress("P6", TxbP6Mail.Text));
                                msg6.Subject = "ワードウルフのお題";
                                var text6 = new MimeKit.TextPart("Plain");
                                text6.Text = LblP6Word.Text;
                                msg6.Body = text6;

                                using (var client = new MailKit.Net.Smtp.SmtpClient())
                                {
                                    try
                                    {
                                        client.Connect(this.MAILHOST, this.MAILPORT, MailKit.Security.SecureSocketOptions.Auto);
                                        client.Authenticate(this.USERNAME, this.PASSWORD);
                                        client.Send(msg1);
                                        client.Disconnect(true);

                                        client.Connect(this.MAILHOST, this.MAILPORT, MailKit.Security.SecureSocketOptions.Auto);
                                        client.Authenticate(this.USERNAME, this.PASSWORD);
                                        client.Send(msg2);
                                        client.Disconnect(true);

                                        client.Connect(this.MAILHOST, this.MAILPORT, MailKit.Security.SecureSocketOptions.Auto);
                                        client.Authenticate(this.USERNAME, this.PASSWORD);
                                        client.Send(msg3);
                                        client.Disconnect(true);

                                        client.Connect(this.MAILHOST, this.MAILPORT, MailKit.Security.SecureSocketOptions.Auto);
                                        client.Authenticate(this.USERNAME, this.PASSWORD);
                                        client.Send(msg4);
                                        client.Disconnect(true);

                                        client.Connect(this.MAILHOST, this.MAILPORT, MailKit.Security.SecureSocketOptions.Auto);
                                        client.Authenticate(this.USERNAME, this.PASSWORD);
                                        client.Send(msg5);
                                        client.Disconnect(true);

                                        client.Connect(this.MAILHOST, this.MAILPORT, MailKit.Security.SecureSocketOptions.Auto);
                                        client.Authenticate(this.USERNAME, this.PASSWORD);
                                        client.Send(msg6);
                                        client.Disconnect(true);
                                    }
                                    catch
                                    {

                                    }
                                }
                            }
                            else
                            {
                                //メッセージボックスを表示する
                                MessageBox.Show("メールアドレスが間違ってるよ", "確認",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        else
                        {
                            //メッセージボックスを表示する
                            MessageBox.Show("メールアドレスが設定されてないよ", "確認",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        break;
                }
            }
            else
            {
                //メッセージボックスを表示する
                MessageBox.Show("ワードがみんなに配られてないよ", "確認",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            
        }

        public void ReadSetFile()
        {
            var dirPath = Directory.GetParent(Assembly.GetExecutingAssembly().Location);

            string filePath = dirPath.ToString() + "\\SettingFile.txt";

            if (File.Exists(filePath))
            {
                string line = "";
                ArrayList al = new ArrayList();

                using (StreamReader sr = new StreamReader(
                    filePath, Encoding.GetEncoding("Shift_JIS")))
                {

                    while ((line = sr.ReadLine()) != null)
                    {
                        al.Add(line);
                    }

                    this.USERNAME = al[0].ToString();
                    this.PASSWORD = al[1].ToString();
                }
            }
        }
    }
}
