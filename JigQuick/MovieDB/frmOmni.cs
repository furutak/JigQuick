using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Linq;
using System.Drawing;
using System.IO;
using System.Data.Linq;
using System.Globalization;
using System.Security.Permissions;
using System.Diagnostics;


namespace JigQuick
{
    public partial class frmOmni : Form
    {
        // コンフィグファイルと、出力テキストファイルは、デスクトップの指定のフォルダに保存する
        string appconfig = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\JigQuickDesk\JigQuickApp\info.ini";
        string outPath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\JigQuickDesk\ConverterTarget\";
        string outPath2 = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\JigQuickDesk\NtrsLog\";

        //ＪＩＧＱＵＩＣＫ用、非ローカル変数
        DataTable dtPartsLot;
        DataTable dtTbi;
        int rowPartsLot;
        int partsLotCount;
        long cumCount;
        bool sound;
        bool duplicate;
        string autoRegister = "ON";
        string parentTextBoxFunc = "OFF";
        string[] description;
        string[] partsLotBreakdown;

        //ＮＴＲＳ用、非ローカル変数
        int okCount;
        int ngCount;
        int targetProcessCount;
        string model;
        string subAssyName;
        string targetProcess;
        string headTableThisMonth = string.Empty;
        string headTableLastMonth = string.Empty;
        string dataTableThisMonth = string.Empty;
        string dataTableLastMonth = string.Empty;
        string whereProcessSql;
        string okImageFile;
        string ngImageFile;
        string standByImageFile;
        string ntrsSwitch;

        // コンストラクタ
        public frmOmni()
        {
            InitializeComponent();
        }

        // ロード時の処理
        private void frmInut_Load(object sender, EventArgs e)
        {
            // ＪＩＧＱＵＩＣＫ

            // 部品ロット保持用データテーブルの行数設定を取得
            rowPartsLot = int.Parse(readIni("OTHERS", "PARTS LOT COUNT", appconfig));
            description = readIni("OTHERS", "DESCRIPTION", appconfig).Split(',');
            partsLotBreakdown = readIni("OTHERS", "PARTS LOT BREAKDOWN", appconfig).Split(',');

            // 自動登録モードのＯＮ・ＯＦＦ取得、登録ボタンの設定
            //autoRegister = readIni("APPLICATION BEHAVIOR", "AUTOMATIC REGISTER", appconfig);

            // 各種処理用のテーブルを生成
            dtPartsLot = new DataTable();
            dtTbi = new DataTable();

            // テーブルの定義
            defineTables(ref dtPartsLot, ref dtTbi);

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

            // --------------------------------------------------------------------------------------
            // 以下、ＮＴＲＳの初期設定
            okImageFile = readIni("MODULE-DATA MATCHING", "OK IMAGE FILE", appconfig);
            ngImageFile = readIni("MODULE-DATA MATCHING", "NG IMAGE FILE", appconfig);
            standByImageFile = readIni("MODULE-DATA MATCHING", "STAND-BY IMAGE FILE", appconfig);
            string standByImagePath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\JigQuickDesk\JigQuickApp\images\" + standByImageFile;
            pnlResult.BackgroundImageLayout = ImageLayout.Zoom;
            pnlResult.BackgroundImage = System.Drawing.Image.FromFile(standByImagePath);

            model = readIni("MODULE-DATA MATCHING", "MODEL", appconfig);
            subAssyName = readIni("MODULE-DATA MATCHING", "SUB ASSY NAME", appconfig);
            targetProcess = readIni("MODULE-DATA MATCHING", "TARGET PROCESS", appconfig);
            ntrsSwitch = readIni("MODULE-DATA MATCHING", "NTRS INLINE MATCHING", appconfig);
            headTableThisMonth = model.ToLower() + DateTime.Today.ToString("yyyyMM");
            headTableLastMonth = model.ToLower() + ((VBS.Right(DateTime.Today.ToString("yyyyMM"), 2) != "01") ?
                (long.Parse(DateTime.Today.ToString("yyyyMM")) - 1).ToString() : (long.Parse(DateTime.Today.ToString("yyyy")) - 1).ToString() + "12");
            dataTableThisMonth = headTableThisMonth + "data";
            dataTableLastMonth = headTableLastMonth + "data";

            txtSubAssy.Text = model + "  " + subAssyName;
            whereProcessSql = makeSqlWhereClause(targetProcess);
            targetProcessCount = targetProcess.Where(c => c == ',').Count() + 1;

            // カウンターの表示（デフォルトはゼロ）
            txtOkCount.Text = okCount.ToString();
            txtNgCount.Text = ngCount.ToString();

            // ログ用フォルダの作成（フォルダが存在しない場合）
            if (!Directory.Exists(outPath2)) Directory.CreateDirectory(outPath2);
        }

        // サブプロシージャ：ＤＴの定義
        private void defineTables(ref DataTable dt1, ref DataTable dt2)
        {
            // 部品ロットグリッドビューを準備する
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
        }

        // サブプロシージャ：ラベルの設定
        private void setLabels()
        {
            txtModel.Text = readIni("LABEL DESCRIPTION", "MODEL", appconfig);
            txtProcess.Text = readIni("LABEL DESCRIPTION", "PROCESS TEXT BOX", appconfig);
            lblPartsLot.Text = readIni("LABEL DESCRIPTION", "PARTS LOT LABEL", appconfig);
            lblChild.Text = readIni("LABEL DESCRIPTION", "CHILD LABEL", appconfig);
            //lblParent.Text = readIni("LABEL DESCRIPTION", "PARENT LABEL", appconfig);
            parentTextBoxFunc = readIni("APPLICATION BEHAVIOR", "PARENT TEXT BOX", appconfig);
            //txtParent.Enabled = parentTextBoxFunc == "ON" ? true : false;
        }

        // 部品ロット情報リセットボタン押下時の処理
        private void btnResetParts_Click(object sender, EventArgs e)
        {
            // ＪＩＧＱＵＩＣＫ側のコントロールの更新
            int r = 4;
            for (int i = 0; i < r; i++)
            {
                dtPartsLot.Rows[i][0] = string.Empty;
                dtPartsLot.Rows[i][1] = string.Empty;
                dtPartsLot.Rows[i][2] = string.Empty;
                dtPartsLot.Rows[i][3] = string.Empty;
                dtPartsLot.Rows[i][4] = string.Empty;
                dtPartsLot.Rows[i][5] = string.Empty;
            }
            
            txtPartsLot.Text = string.Empty;
            txtPartsLot.Enabled = true;
            txtPartsLot.Focus();
            resetViewColor(ref dgvPartsLot);
 
            // ＮＴＲＳ側のコントロールの更新
            btnReset_Click(sender, e);

            // ＪＩＧＱＵＩＣＫ側のコントロールの更新
            txtChild.Text = string.Empty;
            txtChild.Enabled = false;
            txtChild.BackColor = SystemColors.Window;
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

            // ＯＫカウント＝ＲＯＷ設定の場合のみ、製品シリアル用テキストボックスを有効にする
            partsLotCount = dgvPartsLot.Rows.Cast<DataGridViewRow>().Where(a => a.Cells[0].Value.ToString() != string.Empty)
                .Select(a => a.Cells[0].Value).GroupBy(a => a).Count();
            if (!duplicate && partsLotCount == rowPartsLot)
            {
                txtPartsLot.Enabled = false;
                txtChild.Enabled = true;
            }
            else
            {
                txtPartsLot.Enabled = true;
                txtChild.Enabled = false;
            }
        }

        // ＣＨＩＬＤテキストボックス、スキャン時の処理
        private void txtChild_KeyDown(object sender, KeyEventArgs e)
        {
            // バーコード末尾のエンターキー以外は処理しない
            if (e.KeyCode != Keys.Enter) return;

            // 空文字の場合、または文字列の長さが誤っている場合は、処理しない
            if (txtChild.Text == string.Empty) return;
            if (txtChild.Text.Length != 17 && txtChild.Text.Length != 24) return;

            // テキストボックスが読み取り専用状態の場合は、処理しない
            if (txtChild.ReadOnly == true) return;
            
            // 部品ロット情報が出そろっていない場合は、処理しない
            if (duplicate || partsLotCount != rowPartsLot) return;

            // ＮＴＲＳの処理を先行して行い、その結果がＮＧの場合は、ＪＩＧＱＵＩＣＫのテキスト出力を行わない
            bool res = ntrsScanProcess(txtChild.Text);
            if (!res) return;

            // ＪＩＧＱＵＩＣＫ用テキストを出力する
            checkAndOutput();
        }

        // サブプロシージャ：テキストファイル出力前の確認
        private void checkAndOutput()
        {
            // 重複および空白をマーキングする
            colorViewForDuplicate(ref dgvPartsLot);

            // ＯＫカウント＝ＲＯＷ設定の場合のみ、自動登録を行う
            partsLotCount = dgvPartsLot.Rows.Cast<DataGridViewRow>().Where(r => r.Cells[0].Value.ToString() != string.Empty)
                    .Select(r => r.Cells[0].Value).GroupBy(r => r).Count();
            if (duplicate || partsLotCount != rowPartsLot)
            {
                MessageBox.Show("Part lot info is not enough.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 自動登録モードＯＮ・ＯＦＦの場合分け
            if (autoRegister == "ON")
            {
                outputTbiInfo();
            }
        }

        // サブプロシージャ：テキストファイル出力
        private void outputTbiInfo()
        {                           
            string sn = txtChild.Text.Trim();
            string lot = "null";
            string model = txtModel.Text.Trim();
            string date = DateTime.Today.ToString("yyyy/MM/dd");
            string time = DateTime.Now.ToString("HH:mm:ss");
            string cumRecords = string.Empty;

            if (sn == string.Empty) return;
            if (sn.IndexOf("+") >= 1) sn = VBS.Left(sn, 17);

            // 複数の部品ロット情報に、単数のＣＨＩＬＤ情報を紐付け、部品ロット件数分のレコードを作成する
            for (int i = 0; i < rowPartsLot; i++)
            {
                string lotInf = dgvPartsLot[3, i].Value.ToString().Replace(":", "_") + ":" + string.Empty + ":" + description[i] + ":" + dgvPartsLot[0, i].Value.ToString() + ":" +
                    dgvPartsLot[1, i].Value.ToString() + ":" + dgvPartsLot[2, i].Value.ToString().Replace(":", "_") + ":" + dgvPartsLot[4, i].Value.ToString();
                    lotInf = lotInf.Replace(" ", "_").Replace(",", "_").Replace("'", "_").Replace(";", "_").Replace("\"", "_");
                string newRecord = sn + "," + lot + "," + model + "," + date + "," + time + "," + lotInf;
                cumRecords += newRecord + System.Environment.NewLine;
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
                cumCount += 1;
                txtCount.Text = cumCount.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 次のスキャンに備え、テキストボックスを有効化
            txtChild.Enabled = true;
            txtChild.Focus();
            txtChild.SelectAll();
        }

        // サブプロシージャ：重複部品ロット・空レコードをマーキングする
        private void colorViewForDuplicate(ref DataGridView dgv)
        {
            DataTable dt = ((DataTable)dgv.DataSource).Copy();
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                string partCode = dgv["part_code", i].Value.ToString();
                DataRow[] dr = dt.Select("part_code = '" + partCode + "'");
                if (partCode == string.Empty || dr.Length >= 2)
                {
                    dgv.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                    soundAlarm();
                    duplicate = true;
                }
                else
                {
                    dgv.Rows[i].DefaultCellStyle.BackColor = Color.FromKnownColor(KnownColor.Window);
                    duplicate = false;
                }
            }
        }

        // サブプロシージャ：グリットビューの色をリセットする
        private void resetViewColor(ref DataGridView dgv)
        {
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                dgv.Rows[i].DefaultCellStyle.BackColor = Color.FromKnownColor(KnownColor.Window);
            }
            duplicate = false;
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


        // 以下ＮＴＲＳ


        // ＮＴＲＳサブプロシージャ： テスター用ＷＨＥＲＥ句の作成
        private string makeSqlWhereClause(string criteria)
        {
            string sql = " where (";
            foreach (string c in criteria.Split(','))
            {
                sql += "process = " + c + " or ";
            }
            sql = VBS.Left(sql, sql.Length - 3) + ") ";
            System.Diagnostics.Debug.Print(sql);
            return sql;
        }

        // リセットボタン押下時の処理： パネルとテスト結果テキストボックスをクリアし、スキャン用テキストボックスを有効にする
        private void btnReset_Click(object sender, EventArgs e)
        {
            string standByImagePath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\JigQuickDesk\JigQuickApp\images\" + standByImageFile;
            pnlResult.BackgroundImageLayout = ImageLayout.Zoom;
            pnlResult.BackgroundImage = System.Drawing.Image.FromFile(standByImagePath);

            txtResultDetail.Text = string.Empty;
            txtChild.Text = string.Empty;
            txtChild.ReadOnly = false;
            txtChild.BackColor = Color.White;
            txtChild.Focus();
        }

        // カウンターをクリアする
        private void btnSetZero_Click(object sender, EventArgs e)
        {
            okCount = 0;
            ngCount = 0;
            txtOkCount.Text = okCount.ToString();
            txtNgCount.Text = ngCount.ToString();
        }

        // テスト結果を格納するクラス
        public class TestResult
        {
            public string process { get; set; }
            public string judge { get; set; }
            public string inspectdate { get; set; }
        }

        // テスト結果のプロセスコードのみを格納するクラス
        public class ProcessList
        {
            public string process { get; set; }
        }

        // ＮＴＲＳの処理を行い、判定結果を返す
        private bool ntrsScanProcess(string id)
        {
            TfSQL tf = new TfSQL();
            DataTable dt = new DataTable();
            string log = string.Empty;
            string module = id;
            string mdlShort = VBS.Left(module, 16); // 一時的に１６ケタに設定

            string sql1 = "select process, judge, max(inspectdate) as inspectdate from (" +
                    "(select process, case when tjudge = '0' then 'PASS' else 'FAIL' end as judge, inspectdate from " + headTableThisMonth + whereProcessSql + "and serno like '" + mdlShort + "%') union all " +
                    "(select process, case when tjudge = '0' then 'PASS' else 'FAIL' end as judge, inspectdate from " + headTableLastMonth + whereProcessSql + "and serno like '" + mdlShort + "%')" +
                    ") d group by judge, process order by judge desc, process";
            System.Diagnostics.Debug.Print(sql1);
            tf.sqlDataAdapterFillDatatableFromPqmDb(sql1, ref dt);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                System.Diagnostics.Debug.Print(dt.Rows[i][0].ToString() + " " + dt.Rows[i][1].ToString() + " " + dt.Rows[i][2].ToString());
            }

            var allResults = dt.AsEnumerable().Select(r => new TestResult()
            { process = r.Field<string>("process"), judge = r.Field<string>("judge"), inspectdate = r.Field<DateTime>("inspectdate").ToString("yyyy/MM/dd HH:mm:ss"), }).ToList();
            string scanTime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            // １．パスのプロセス名を取得
            var passResults = allResults.Where(r => r.judge == "PASS").Select(r => new ProcessList() { process = r.process }).OrderBy(r => r.process).ToList();
            foreach (var p in passResults) System.Diagnostics.Debug.Print(p.process);
            // ２．１に含まれないフェイルのプロセス名を取得
            var failResults = allResults.Where(r => r.judge == "FAIL").Select(r => new ProcessList() { process = r.process }).OrderBy(r => r.process).ToList();
            List<string> process = failResults.Select(r => r.process).Except(passResults.Select(r => r.process)).ToList();
            failResults = failResults.Where(r => process.Contains(r.process)).ToList();
            foreach (var p in failResults) System.Diagnostics.Debug.Print(p.process);
            // ３．１にも２にも含まれない、テスト結果なしプロセスを取得する
            var skipResults = targetProcess.Replace("'", string.Empty).Split(',').ToList().Select(r => new ProcessList() { process = r.ToString() }).OrderBy(r => r.process).ToList();
            process = skipResults.Select(r => r.process).Except(passResults.Select(r => r.process)).ToList().Except(failResults.Select(r => r.process)).ToList();
            skipResults = skipResults.Where(r => process.Contains(r.process)).ToList();
            foreach (var p in skipResults) System.Diagnostics.Debug.Print(p.process);

            // ディスプレイ用のプロセス名リストを加工する
            string displayPass = string.Empty;
            string displayFail = string.Empty;
            string displayAll = string.Empty;   // ログ用
            List<TestResult> allLog = new List<TestResult>();
            foreach (var p in passResults)
            {
                displayPass += p.process + " ";
                allLog.Add(new TestResult { process = p.process, judge = "PASS", inspectdate = string.Empty });
            }
            displayPass = displayPass.Trim();
            foreach (var p in failResults)
            {
                displayFail += p.process + " F ";
                allLog.Add(new TestResult { process = p.process, judge = "FAIL", inspectdate = string.Empty });
            }
            foreach (var p in skipResults)
            {
                displayFail += p.process + " S ";
                allLog.Add(new TestResult { process = p.process, judge = "SKIP", inspectdate = string.Empty });
            }
            displayFail = displayFail.Trim();
            allLog = allLog.OrderBy(r => r.process).ToList();
            foreach (var p in allLog)
            {
                displayAll += (p.process + ":" + p.judge + ",");
            }
            displayAll = VBS.Left(displayAll, displayAll.Length - 1);

            bool result = false;

            // アプリケーションスクリーンに、テスト結果を表示する
            if (passResults.Count == targetProcessCount)
            {
                string okImagePass = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\JigQuickDesk\JigQuickApp\images\" + okImageFile;
                pnlResult.BackgroundImageLayout = ImageLayout.Zoom;
                pnlResult.BackgroundImage = System.Drawing.Image.FromFile(okImagePass);

                // ＯＫカウントの加算
                okCount += 1;
                txtOkCount.Text = okCount.ToString();

                // ＰＡＳＳの場合は、ＴＲＵＥを返す
                result = true;

                // 次のモジュールのスキャンにそなえ、スキャン用テキストボックスのテキストを選択し、上書き可能にする
                txtChild.SelectAll();
            }
            else
            {
                string ngImagePath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\JigQuickDesk\JigQuickApp\images\" + ngImageFile;
                pnlResult.BackgroundImageLayout = ImageLayout.Zoom;
                pnlResult.BackgroundImage = System.Drawing.Image.FromFile(ngImagePath);

                // ＮＧカウントの加算
                ngCount += 1;
                txtNgCount.Text = ngCount.ToString();

                // ＦＡＩＬの場合は、ＦＡＬＳＥを返す
                result = true;

                // 次のモジュールのスキャンをストップするめた、スキャン用テキストボックスを無効にする
                txtChild.ReadOnly = true;
                txtChild.BackColor = Color.Red;

                // アラームでの警告
                soundAlarm();
            }

            // アプリケーションスクリーンとデスクトップフォルダの両方の用途用に、日付とテスト結果詳細文字列を作成
            log = Environment.NewLine + scanTime + "," + module + "," + displayAll;

            // スクリーンへの表示
            txtResultDetail.Text = log.Replace(",", ",  ").Replace(Environment.NewLine, string.Empty);

            // ログ書込み：同日日付のファイルが存在する場合は追記し、存在しない場合はファイルを作成追記する（AppendAllText がやってくれる）
            try
            {
                string outFile = outPath2 + DateTime.Today.ToString("yyyyMMdd") + ".txt";
                if (System.IO.File.Exists(outFile))
                {
                    System.IO.File.AppendAllText(outFile, log, System.Text.Encoding.GetEncoding("UTF-8"));
                }
                else
                {
                    string header = DateTime.Today.ToString("yyyy/MM/dd") + " " + model + " " + subAssyName +
                        Environment.NewLine + "SCAN TIME,PRODUCT SERIAL,TEST DETAIL";
                    System.IO.File.AppendAllText(outFile, header + log, System.Text.Encoding.GetEncoding("UTF-8"));
                }

                return result;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return result;
            }
        }
    }
}