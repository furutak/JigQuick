using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Security.Permissions;
using System.Runtime.InteropServices;
using System.Linq;
using System.Drawing;
using System.IO;

namespace JigQuick
{
    public partial class frmTriple : Form
    {
        // コンフィグファイルと、出力テキストファイルは、デスクトップの指定のフォルダに保存する
        string appconfig = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\JigQuickDesk\JigQuickApp\info.ini";
        string outPath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\JigQuickDesk\ConverterTarget\";

        // 内部処理用変数
        DataTable dtPartsLot;
        DataTable dtTbi;
        DataTable dtChild;
        long cumCount;
        bool sound;
        bool dupPartsLot;
        bool dupChild;
        bool wrongScanOrder;
        int countPartsLot;
        int countChild;
        string c1c2prMatching = "OFF";
        string[] c1c2prList;
        string[] description;
        string[] partsLotBreakdown;
        string[] jigposition;
        string posOutput;
        string reverseEntry;
        string parentuse;

        // スクリーン動作調整用変数
        string autoRegister = "ON";
        int rowPartsLot;
        int rowChild;


        // コンストラクタ
        public frmTriple()
        {
            InitializeComponent();
        }

        // ロード時の処理
        private void frmInut_Load(object sender, EventArgs e)
        {
            // 当フォームの表示場所を指定
            //this.Left = 30; // 画面中央
            //this.Top = 35;  // 画面中央

            // 部品ロット保持用データテーブル、チャイルド保持用データテーブルの、行数設定を取得
            rowPartsLot = int.Parse(readIni("OTHERS", "PARTS LOT COUNT", appconfig));
            rowChild = int.Parse(readIni("OTHERS", "CHILD COUNT", appconfig));
            description = readIni("OTHERS", "DESCRIPTION", appconfig).Split(',');
            partsLotBreakdown = readIni("OTHERS", "PARTS LOT BREAKDOWN", appconfig).Split(',');
            jigposition = readIni("OTHERS", "JIG POSITION", appconfig).Split(',');
            posOutput = readIni("APPLICATION BEHAVIOR", "POSITION OUTPUT", appconfig);
            reverseEntry = readIni("APPLICATION BEHAVIOR", "REVERSE ENTRY", appconfig);
            parentuse = readIni("APPLICATION BEHAVIOR", "PARENT TEXT BOX", appconfig);
            c1c2prMatching = readIni("APPLICATION BEHAVIOR", "CH1, CH2, PR MATCHING", appconfig);
            c1c2prList = readIni("OTHERS", "CH1, CH2, PR LIST", appconfig).Split(',');

            // 自動登録モードのＯＮ・ＯＦＦ取得、登録ボタンの設定
            //autoRegister = readIni("APPLICATION BEHAVIOR", "AUTOMATIC REGISTER", appconfig);

            // 各種処理用のテーブルを生成
            dtPartsLot = new DataTable();
            dtTbi = new DataTable();
            dtChild = new DataTable();

            // テーブルの定義
            defineTables(ref dtPartsLot, ref dtTbi, ref dtChild);

            // ラベルの設定
            setLabels();

            // 累積カウント（同日のテキストファイル内のレコード数）を取得
            string outFile = outPath + DateTime.Today.ToString("yyyyMMdd") + ".txt";
            if (System.IO.File.Exists(outFile))
            {
                using (StreamReader r = new StreamReader(outFile))
                {
                    int i = 0;
                    while (r.ReadLine() != null) { i++; }
                    cumCount = (i -2) / rowPartsLot;
                }
            }

            // カウントテキストボックスの中身の表示
            txtCount.Text = cumCount.ToString();
        }

        // サブプロシージャ：ＤＴの定義
        private void defineTables(ref DataTable dt1, ref DataTable dt2, ref DataTable dt3)
        {
            // 部品ロット用グリッドビューを準備する
            dt1.Columns.Add("part_code", Type.GetType("System.String"));
            dt1.Columns.Add("part_name", Type.GetType("System.String"));
            dt1.Columns.Add("vendor", Type.GetType("System.String"));
            dt1.Columns.Add("invoice", Type.GetType("System.String"));
            dt1.Columns.Add("shipdate", Type.GetType("System.String"));
            dt1.Columns.Add("qty", Type.GetType("System.String"));

            for (int i = 0; i < rowPartsLot; i++) dt1.Rows.Add(dt1.NewRow());
            dgvPartsLot.DataSource = dt1;

            // 行ヘッダーに行番号を追加し、行ヘッダ幅を自動調節する
            for (int i = 0; i < rowPartsLot; i++) dgvPartsLot.Rows[i].HeaderCell.Value = partsLotBreakdown[i];
            dgvPartsLot.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);


            // 登録処理用ＴＢＩデータテーブルを準備する
            dt2.Columns.Add("S/N", Type.GetType("System.String"));
            dt2.Columns.Add("Lot", Type.GetType("System.String"));
            dt2.Columns.Add("ModelName", Type.GetType("System.String"));
            dt2.Columns.Add("Date", Type.GetType("System.String"));
            dt2.Columns.Add("Time", Type.GetType("System.String"));
            dt2.Columns.Add("LotInfo", Type.GetType("System.String"));


            // チャイルド用グリッドビューを準備する
            dt3.Columns.Add("child", Type.GetType("System.String"));

            for (int i = 0; i < rowChild; i++) dt3.Rows.Add(dt3.NewRow());
            dgvChild.DataSource = dt3;

            // 行ヘッダーに行番号を追加し、行ヘッダ幅を自動調節する
            for (int i = 0; i < rowChild; i++) dgvChild.Rows[i].HeaderCell.Value = (i + 1).ToString();
            dgvChild.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }

        // サブプロシージャ：ラベルの設定
        private void setLabels()
        {
            txtModel.Text = readIni("LABEL DESCRIPTION", "MODEL", appconfig);
            txtProcess.Text = readIni("LABEL DESCRIPTION", "PROCESS TEXT BOX", appconfig);
            lblPartsLot.Text = readIni("LABEL DESCRIPTION", "PARTS LOT LABEL", appconfig);
            lblChild.Text = readIni("LABEL DESCRIPTION", "CHILD LABEL", appconfig);
            lblParent.Text = readIni("LABEL DESCRIPTION", "PARENT LABEL", appconfig);
            txtParent.Enabled = parentuse == "ON" ? true : false;
        }

        // 部品ロット情報のスキャン時の処理
        private void txtPartsLot_KeyDown_1(object sender, KeyEventArgs e)
        {
            // バーコード末尾のエンターキー以外は処理しない
            if (e.KeyCode != Keys.Enter) return;

            // 空文字の場合は処理しない
            string scan = txtPartsLot.Text;
            if (scan == string.Empty) return;

            // 部品ロットグリッドビューの現在のセル行と、その次の行を格納する
            int r = dgvPartsLot.CurrentCell.RowIndex;
            int y = r < rowPartsLot - 1 ? r + 1 : rowPartsLot - 1;

            // セミコロンでＱＲ読み取り内容を分割し、グリットビューに表示する
            try
            {
                string[] split = scan.Split(';');
                dtPartsLot.Rows[r][0] = split[0];
                dtPartsLot.Rows[r][1] = split[1];
                dtPartsLot.Rows[r][2] = split[2];
                dtPartsLot.Rows[r][3] = split[3];
                dtPartsLot.Rows[r][4] = split[4];
                dtPartsLot.Rows[r][5] = split[5];

                txtPartsLot.Text = string.Empty;
                dgvPartsLot.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                dgvPartsLot.CurrentCell = dgvPartsLot[0 , y];
            }
            // 分割できない文字列の場合は、グリットビューをクリア（空表示にする）
            catch (Exception)
            {
                dtPartsLot.Rows[r][0] = string.Empty;
                dtPartsLot.Rows[r][1] = string.Empty;
                dtPartsLot.Rows[r][2] = string.Empty;
                dtPartsLot.Rows[r][3] = string.Empty;
                dtPartsLot.Rows[r][4] = string.Empty;
                dtPartsLot.Rows[r][5] = string.Empty;

                txtPartsLot.Text = string.Empty;
                dgvPartsLot.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                // MessageBox.Show(ex.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            resetViewColor(ref dgvPartsLot);
        }

        // ＣＨＩＬＤテキストボックス、スキャン時の処理
        private void txtChild_KeyDown(object sender, KeyEventArgs e)
        {
            // バーコード末尾のエンターキー以外は処理しない
            if (e.KeyCode != Keys.Enter) return;

            // 空文字の場合は処理しない
            string scan = txtChild.Text;
            if (scan == string.Empty) return;

            // チャイルドグリッドビューの現在のセル行と、その次の行を格納する
            int r = dgvChild.CurrentCell.RowIndex;
            int y = r < rowChild - 1 ? r + 1 : rowChild - 1;

            // グリットビューに表示する
            dtChild.Rows[r][0] = scan;
            txtChild.Text = string.Empty;
            dgvChild.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dgvChild.CurrentCell = dgvChild[0, y];

            resetViewColor(ref dgvChild);

            // チャイルドが存在する場合は、チャイルドスキャン用、または、ドッキングジグ下部スキャン用に、カーソルを移動
            if (r == rowChild - 1 && parentuse == "ON")
            {
                txtParent.Focus();
            }
            // チャイルドが存在しない場合は、登録を行う（ペアレントスキャン時と同様の処理）
            else if (r == rowChild - 1 && parentuse == "OFF")
            {
                txtParent_KeyDown(sender, e);
            }
        }

        // ペアレントスキャン時の処理
        private void txtParent_KeyDown(object sender, KeyEventArgs e)
        {
            // バーコード末尾のエンターキー以外は処理しない
            if (e.KeyCode != Keys.Enter) return;

            // 空文字の場合は処理しない
            if (txtParent.Text == string.Empty　&& parentuse == "ON") return;

            // 重複および空白をマーキングする
            colorViewForDuplicate(ref dgvPartsLot, "part_code");
            colorViewForDuplicate(ref dgvChild, "child");

            // 部品ロット行カウント＝ＲＯＷ設定の場合のみ、自動登録を行う
            countPartsLot = dgvPartsLot.Rows.Cast<DataGridViewRow>().Where(r => r.Cells[0].Value.ToString() != string.Empty)
                    .Select(r => r.Cells[0].Value).GroupBy(r => r).Count();
            if (dupPartsLot || countPartsLot != rowPartsLot)
            {
                MessageBox.Show("Parts lot information is not enough.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 部品ロット行カウント＝ＲＯＷ設定の場合のみ、自動登録を行う
            countChild = dgvChild.Rows.Cast<DataGridViewRow>().Where(r => r.Cells[0].Value.ToString() != string.Empty)
                    .Select(r => r.Cells[0].Value).GroupBy(r => r).Count();
            if (countChild != rowChild)　// if (dupChild || countChild != rowChild)
            {
                MessageBox.Show("Child information is not enough.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 自動登録モードＯＮ・ＯＦＦの場合分け
            if (autoRegister == "ON")
            {
                outPutTbiInfo();
            }
        }

        // 登録ボタン押下時、ＴＢＩテーブルへの登録
        private void btnRegister_Click(object sender, EventArgs e)
        {
            // 部品ロット行カウント＝ＲＯＷ設定の場合のみ、自動登録を行う
            countPartsLot = dgvPartsLot.Rows.Cast<DataGridViewRow>().Where(r => r.Cells[0].Value.ToString() != string.Empty)
                    .Select(r => r.Cells[0].Value).GroupBy(r => r).Count();
            if (dupPartsLot || countPartsLot != rowPartsLot)
            {
                MessageBox.Show("Parts lot information is not enough.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //// 部品ロット行カウント＝ＲＯＷ設定の場合のみ、自動登録を行う
            //countChild = dgvChild.Rows.Cast<DataGridViewRow>().Where(r => r.Cells[0].Value.ToString() != string.Empty)
            //        .Select(r => r.Cells[0].Value).GroupBy(r => r).Count();
            //if (countChild <= 0)　// if (dupChild || countChild != rowChild)
            //{
            //    MessageBox.Show("Child information is missing.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}

            if (parentuse == "ON" && txtParent.Text == string.Empty)　// if (dupChild || countChild != rowChild)
            {
                MessageBox.Show(lblParent.Text + " information is missing.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            outPutTbiInfo();
        }

        // サブプロシージャ：ＴＢＩテーブルへの登録
        private void outPutTbiInfo()
        {
            string sn;
            string lot = txtParent.Text.Trim() == string.Empty? "null" : txtParent.Text.Trim();
            string model = txtModel.Text.Trim();
            string date = DateTime.Today.ToString("yyyy/MM/dd");
            string time = DateTime.Now.ToString("HH:mm:ss");
            string cumRecords = string.Empty;
            int position = 1;
            int tempCount = 0;

            // チャイルド１、チャイルド２、ペアレントのマッチングを行う（当機能の設定、ＯＮの場合）
            if (c1c2prMatching == "ON")
            {
                wrongScanOrder = matchScanOrder(string.Empty, string.Empty, txtParent.Text);
                if (wrongScanOrder)
                {
                    MessageBox.Show("Case pallet barcode is wrong." + Environment.NewLine + "Please clear and re-scan all.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            // 複数の部品ロット情報に、複数のＣＨＩＬＤ情報を紐付け、部品ロット件数×ＣＨＩＬＤ件数の、レコードを作成する
            for (int i = 0; i < rowChild; i++)
            {
                sn = dgvChild[0, i].Value.ToString();
                if (sn == string.Empty) continue;
                if (sn.IndexOf("+") >= 1) sn = VBS.Left(sn, 17);

                string pos = posOutput == "OFF" ? string.Empty : position.ToString("00");
                string newRecord = string.Empty;

                for (int j = 0; j < rowPartsLot; j++)
                {
                    string lotInf;
                    // ペアレントの使用がない場合（Ｄ－４）は、ジグポジション列は空白
                    if (parentuse == "ON")
                    {
                        lotInf = dgvPartsLot[3, j].Value.ToString().Replace(":", "_") + ":" + jigposition[j] + pos + ":" + description[j] + ":" + dgvPartsLot[0, j].Value.ToString() + ":" +
                            dgvPartsLot[1, j].Value.ToString() + ":" + dgvPartsLot[2, j].Value.ToString().Replace(":", "_") + ":" + dgvPartsLot[4, j].Value.ToString();
                    }
                    else
                    {
                        lotInf = dgvPartsLot[3, j].Value.ToString().Replace(":", "_") + ":" + string.Empty + ":" + description[j] + ":" + dgvPartsLot[0, j].Value.ToString() + ":" +
                        dgvPartsLot[1, j].Value.ToString() + ":" + dgvPartsLot[2, j].Value.ToString().Replace(":", "_") + ":" + dgvPartsLot[4, j].Value.ToString();
                    }

                    lotInf = lotInf.Replace(" ", "_").Replace(",", "_").Replace("'", "_").Replace(";", "_").Replace("\"", "_");

                    // 本来、ＣＨＩＬＤを左列ＳＥＲＮＯ、ＰＡＲＥＮＴを右列ＬＯＴ、へ出力するが、ＭＡＩＮジグ等は、逆で出力する
                    if (reverseEntry == "ON")
                    {
                        newRecord = lot + "," + sn + "," + model + "," + date + "," + time + "," + lotInf;
                    }
                    else
                    {
                        newRecord = sn + "," + lot + "," + model + "," + date + "," + time + "," + lotInf;
                    }

                    cumRecords += newRecord + System.Environment.NewLine;
                }

                position += 1;
                tempCount += 1;
            }

            // 同日日付のファイルが存在する場合は追記し、存在しない場合はファイルを作成しヘッダーを書き込みの上、追記する
            try
            {
                string outFile = outPath + DateTime.Today.ToString("yyyyMMdd") + ".txt";
                if (System.IO.File.Exists(outFile))
                {
                    System.IO.File.AppendAllText(outFile, cumRecords, System.Text.Encoding.GetEncoding("UTF-8"));
                }
                else
                {
                    string header = DateTime.Today.ToString("yyyy/MM/dd") + "," + model + Environment.NewLine +
                        "SN,LOT,MODELNAME,DATE,TIME,LOTINFO" + Environment.NewLine;
                    System.IO.File.AppendAllText(outFile, header + cumRecords, System.Text.Encoding.GetEncoding("UTF-8"));
                }

                // 登録カウントの表示
                cumCount += tempCount;
                txtCount.Text = cumCount.ToString();

                // 次のチャイルドスキャンのための準備
                clerChildView();
                txtParent.Text = string.Empty;
                txtChild.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        // チャイルドクリアボタン押下時、チャイルドグリッドビューのクリア
        private void btnClerChild_Click(object sender, EventArgs e)
        {
            clerChildView();
        }

        // サブプロシージャ：チャイルドグリッドビューのクリア
        private void clerChildView()
        {
            for (int i = 0; i < rowChild; i++) dgvChild[0, i].Value = string.Empty;
            dgvChild.CurrentCell = dgvChild[0, 0];
            txtChild.Focus();
        }

        // ペアレントクリアボタン押下時、チャイルドグリッドビューのクリア
        private void button1_Click(object sender, EventArgs e)
        {
            txtParent.Text = string.Empty;
        }

        // サブプロシージャ：重複部品ロット・空レコードをマーキングする
        private void colorViewForDuplicate(ref DataGridView dgv, string field)
        {
            // 特定のフィールドが指定された場合のみ、処理を行う
            if (field != "part_code" && field != "child") return;

            try
            {
                DataTable dt = ((DataTable)dgv.DataSource).Copy();
                for (int i = 0; i < dgv.Rows.Count; i++)
                {
                    string key = adjustSelectFilterString(dgv[field, i].Value.ToString());
                    DataRow[] dr = dt.Select("[" + field + "]" + " = '" + key + "'");
                    if (key == string.Empty || dr.Length >= 2)
                    {
                        dgv.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                        soundAlarm();
                        if (field == "part_code") dupPartsLot = true;
                        else if (field == "child") dupChild = true;
                    }
                    else
                    {
                        dgv.Rows[i].DefaultCellStyle.BackColor = Color.FromKnownColor(KnownColor.Window);
                        if (field == "part_code") dupPartsLot = false;
                        else if (field == "child") dupChild = false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        // サブプロシージャ：グリットビューの色をリセットする
        private void resetViewColor(ref DataGridView dgv)
        {
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                dgv.Rows[i].DefaultCellStyle.BackColor = Color.FromKnownColor(KnownColor.Window);
            }
            dupPartsLot = false;
            dupChild = false;
        }

        // サブプロシージャ：チャイルド１、チャイルド２、ペアレントのマッチングを行う
        private bool matchScanOrder(string child1, string child2, string parent)
        {
            bool parMissing = parent.IndexOf(c1c2prList[2]) == -1 ? true : false;
            if (parMissing) soundAlarm();
            return parMissing;
        }

        // サブプロシージャ：ＳＥＬＥＣＴメソッド用キーに含まれる、ワイルドカードを調整する
        private string adjustSelectFilterString(string filterString)
        {
            string text = string.Empty;
            foreach (char c in filterString)
            {
                switch (c)
                {
                    case '*':
                    case '%':
                    case '[':
                    case ']':
                        text += "[" + c + "]";
                        break;
                    default:
                        text += c;
                        break;
                }
            }
            return text;
        }

        //MP3ファイル（今回は警告音）を再生する
        private string aliasName = "MediaFile";
        private void soundAlarm()
        {
            string currentDir = System.Environment.CurrentDirectory;
            string fileName = currentDir + @"\warning.mp3";
            string cmd;

            if (sound)
            {
                cmd = "stop " + aliasName;
                mciSendString(cmd, null, 0, IntPtr.Zero);
                cmd = "close " + aliasName;
                mciSendString(cmd, null, 0, IntPtr.Zero);
                sound = false;
            }

            cmd = "open \"" + fileName + "\" type mpegvideo alias " + aliasName;
            if (mciSendString(cmd, null, 0, IntPtr.Zero) != 0) return;
            cmd = "play " + aliasName;
            mciSendString(cmd, null, 0, IntPtr.Zero);
            sound = true;
        }
        // Windows API をインポート
        [System.Runtime.InteropServices.DllImport("winmm.dll")]
        private static extern int mciSendString(String command,StringBuilder buffer, int bufferSize, IntPtr hwndCallback);

        // 設定テキストファイルの読み込み
        private string readIni(string s, string k, string cfs)
        {
            StringBuilder retVal = new StringBuilder(255);
            string section = s;
            string key = k;
            string def = String.Empty;
            int size = 255;
            //get the value from the key in section
            int strref = GetPrivateProfileString(section, key, def, retVal, size, cfs);
            return retVal.ToString();
        }
        // Windows API をインポート
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filepath);


    }
}