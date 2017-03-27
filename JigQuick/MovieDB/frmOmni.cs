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
        // �R���t�B�O�t�@�C���ƁA�o�̓e�L�X�g�t�@�C���́A�f�X�N�g�b�v�̎w��̃t�H���_�ɕۑ�����
        string appconfig = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\JigQuickDesk\JigQuickApp\info.ini";
        string outPath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\JigQuickDesk\ConverterTarget\";
        string outPath2 = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\JigQuickDesk\NtrsLog\";

        //�i�h�f�p�t�h�b�j�p�A�񃍁[�J���ϐ�
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

        //�m�s�q�r�p�A�񃍁[�J���ϐ�
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

        // �R���X�g���N�^
        public frmOmni()
        {
            InitializeComponent();
        }

        // ���[�h���̏���
        private void frmInut_Load(object sender, EventArgs e)
        {
            // �i�h�f�p�t�h�b�j

            // ���i���b�g�ێ��p�f�[�^�e�[�u���̍s���ݒ���擾
            rowPartsLot = int.Parse(readIni("OTHERS", "PARTS LOT COUNT", appconfig));
            description = readIni("OTHERS", "DESCRIPTION", appconfig).Split(',');
            partsLotBreakdown = readIni("OTHERS", "PARTS LOT BREAKDOWN", appconfig).Split(',');

            // �����o�^���[�h�̂n�m�E�n�e�e�擾�A�o�^�{�^���̐ݒ�
            //autoRegister = readIni("APPLICATION BEHAVIOR", "AUTOMATIC REGISTER", appconfig);

            // �e�폈���p�̃e�[�u���𐶐�
            dtPartsLot = new DataTable();
            dtTbi = new DataTable();

            // �e�[�u���̒�`
            defineTables(ref dtPartsLot, ref dtTbi);

            // ���x���̐ݒ�
            setLabels();

            // �ݐσJ�E���g�i�����̃e�L�X�g�t�@�C�����̃��R�[�h���j���擾
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

            // �J�E���g�e�L�X�g�{�b�N�X�̒��g�̕\��
            txtCount.Text = cumCount.ToString();

            // --------------------------------------------------------------------------------------
            // �ȉ��A�m�s�q�r�̏����ݒ�
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

            // �J�E���^�[�̕\���i�f�t�H���g�̓[���j
            txtOkCount.Text = okCount.ToString();
            txtNgCount.Text = ngCount.ToString();

            // ���O�p�t�H���_�̍쐬�i�t�H���_�����݂��Ȃ��ꍇ�j
            if (!Directory.Exists(outPath2)) Directory.CreateDirectory(outPath2);
        }

        // �T�u�v���V�[�W���F�c�s�̒�`
        private void defineTables(ref DataTable dt1, ref DataTable dt2)
        {
            // ���i���b�g�O���b�h�r���[����������
            dt1.Columns.Add("part_code", Type.GetType("System.String"));
            dt1.Columns.Add("part_name", Type.GetType("System.String"));
            dt1.Columns.Add("vendor", Type.GetType("System.String"));
            dt1.Columns.Add("invoice", Type.GetType("System.String"));
            dt1.Columns.Add("shipdate", Type.GetType("System.String"));
            dt1.Columns.Add("qty", Type.GetType("System.String"));

            for (int i = 0; i < rowPartsLot; i++) dt1.Rows.Add(dt1.NewRow());
            dgvPartsLot.DataSource = dt1;

            // �s�w�b�_�[�ɍs�ԍ���ǉ����A�s�w�b�_�����������߂���
            for (int i = 0; i < rowPartsLot; i++) dgvPartsLot.Rows[i].HeaderCell.Value = partsLotBreakdown[i];
            dgvPartsLot.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);

            // �o�^�����p�s�a�h�f�[�^�e�[�u������������
            dt2.Columns.Add("S/N", Type.GetType("System.String"));
            dt2.Columns.Add("Lot", Type.GetType("System.String"));
            dt2.Columns.Add("ModelName", Type.GetType("System.String"));
            dt2.Columns.Add("Date", Type.GetType("System.String"));
            dt2.Columns.Add("Time", Type.GetType("System.String"));
            dt2.Columns.Add("LotInfo", Type.GetType("System.String"));
        }

        // �T�u�v���V�[�W���F���x���̐ݒ�
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

        // ���i���b�g��񃊃Z�b�g�{�^���������̏���
        private void btnResetParts_Click(object sender, EventArgs e)
        {
            // �i�h�f�p�t�h�b�j���̃R���g���[���̍X�V
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
 
            // �m�s�q�r���̃R���g���[���̍X�V
            btnReset_Click(sender, e);

            // �i�h�f�p�t�h�b�j���̃R���g���[���̍X�V
            txtChild.Text = string.Empty;
            txtChild.Enabled = false;
            txtChild.BackColor = SystemColors.Window;
        }

        // ���i���b�g���̃X�L�������̏���
        private void txtPartsLot_KeyDown_1(object sender, KeyEventArgs e)
        {
            // �o�[�R�[�h�����̃G���^�[�L�[�ȊO�͏������Ȃ�
            if (e.KeyCode != Keys.Enter) return;

            // �󕶎��̏ꍇ�͏������Ȃ�
            string scan = txtPartsLot.Text;
            if (scan == string.Empty) return;

            // ���i���b�g�O���b�h�r���[�̌��݂̃Z���s�ƁA���̎��̍s���i�[����
            int r = dgvPartsLot.CurrentCell.RowIndex;
            int y = r < rowPartsLot - 1 ? r + 1 : rowPartsLot - 1;

            // �Z�~�R�����łp�q�ǂݎ����e�𕪊����A�O���b�g�r���[�ɕ\������
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
            // �����ł��Ȃ�������̏ꍇ�́A�O���b�g�r���[���N���A�i��\���ɂ���j
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

            // �n�j�J�E���g���q�n�v�ݒ�̏ꍇ�̂݁A���i�V���A���p�e�L�X�g�{�b�N�X��L���ɂ���
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

        // �b�g�h�k�c�e�L�X�g�{�b�N�X�A�X�L�������̏���
        private void txtChild_KeyDown(object sender, KeyEventArgs e)
        {
            // �o�[�R�[�h�����̃G���^�[�L�[�ȊO�͏������Ȃ�
            if (e.KeyCode != Keys.Enter) return;

            // �󕶎��̏ꍇ�A�܂��͕�����̒���������Ă���ꍇ�́A�������Ȃ�
            if (txtChild.Text == string.Empty) return;
            if (txtChild.Text.Length != 17 && txtChild.Text.Length != 24) return;

            // �e�L�X�g�{�b�N�X���ǂݎ���p��Ԃ̏ꍇ�́A�������Ȃ�
            if (txtChild.ReadOnly == true) return;
            
            // ���i���b�g��񂪏o������Ă��Ȃ��ꍇ�́A�������Ȃ�
            if (duplicate || partsLotCount != rowPartsLot) return;

            // �m�s�q�r�̏������s���čs���A���̌��ʂ��m�f�̏ꍇ�́A�i�h�f�p�t�h�b�j�̃e�L�X�g�o�͂��s��Ȃ�
            bool res = ntrsScanProcess(txtChild.Text);
            if (!res) return;

            // �i�h�f�p�t�h�b�j�p�e�L�X�g���o�͂���
            checkAndOutput();
        }

        // �T�u�v���V�[�W���F�e�L�X�g�t�@�C���o�͑O�̊m�F
        private void checkAndOutput()
        {
            // �d������ы󔒂��}�[�L���O����
            colorViewForDuplicate(ref dgvPartsLot);

            // �n�j�J�E���g���q�n�v�ݒ�̏ꍇ�̂݁A�����o�^���s��
            partsLotCount = dgvPartsLot.Rows.Cast<DataGridViewRow>().Where(r => r.Cells[0].Value.ToString() != string.Empty)
                    .Select(r => r.Cells[0].Value).GroupBy(r => r).Count();
            if (duplicate || partsLotCount != rowPartsLot)
            {
                MessageBox.Show("Part lot info is not enough.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // �����o�^���[�h�n�m�E�n�e�e�̏ꍇ����
            if (autoRegister == "ON")
            {
                outputTbiInfo();
            }
        }

        // �T�u�v���V�[�W���F�e�L�X�g�t�@�C���o��
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

            // �����̕��i���b�g���ɁA�P���̂b�g�h�k�c����R�t���A���i���b�g�������̃��R�[�h���쐬����
            for (int i = 0; i < rowPartsLot; i++)
            {
                string lotInf = dgvPartsLot[3, i].Value.ToString().Replace(":", "_") + ":" + string.Empty + ":" + description[i] + ":" + dgvPartsLot[0, i].Value.ToString() + ":" +
                    dgvPartsLot[1, i].Value.ToString() + ":" + dgvPartsLot[2, i].Value.ToString().Replace(":", "_") + ":" + dgvPartsLot[4, i].Value.ToString();
                    lotInf = lotInf.Replace(" ", "_").Replace(",", "_").Replace("'", "_").Replace(";", "_").Replace("\"", "_");
                string newRecord = sn + "," + lot + "," + model + "," + date + "," + time + "," + lotInf;
                cumRecords += newRecord + System.Environment.NewLine;
            }

            // �������t�̃t�@�C�������݂���ꍇ�͒ǋL���A���݂��Ȃ��ꍇ�̓t�@�C�����쐬���w�b�_�[���������݂̏�A�ǋL����
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

                // �o�^�J�E���g�̕\��
                cumCount += 1;
                txtCount.Text = cumCount.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ���̃X�L�����ɔ����A�e�L�X�g�{�b�N�X��L����
            txtChild.Enabled = true;
            txtChild.Focus();
            txtChild.SelectAll();
        }

        // �T�u�v���V�[�W���F�d�����i���b�g�E�󃌃R�[�h���}�[�L���O����
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

        // �T�u�v���V�[�W���F�O���b�g�r���[�̐F�����Z�b�g����
        private void resetViewColor(ref DataGridView dgv)
        {
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                dgv.Rows[i].DefaultCellStyle.BackColor = Color.FromKnownColor(KnownColor.Window);
            }
            duplicate = false;
        }

        //MP3�t�@�C���i����͌x�����j���Đ�����
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
        // Windows API ���C���|�[�g
        [System.Runtime.InteropServices.DllImport("winmm.dll")]
        private static extern int mciSendString(String command,StringBuilder buffer, int bufferSize, IntPtr hwndCallback);

        // �ݒ�e�L�X�g�t�@�C���̓ǂݍ���
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
        // Windows API ���C���|�[�g
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filepath);


        // �ȉ��m�s�q�r


        // �m�s�q�r�T�u�v���V�[�W���F �e�X�^�[�p�v�g�d�q�d��̍쐬
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

        // ���Z�b�g�{�^���������̏����F �p�l���ƃe�X�g���ʃe�L�X�g�{�b�N�X���N���A���A�X�L�����p�e�L�X�g�{�b�N�X��L���ɂ���
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

        // �J�E���^�[���N���A����
        private void btnSetZero_Click(object sender, EventArgs e)
        {
            okCount = 0;
            ngCount = 0;
            txtOkCount.Text = okCount.ToString();
            txtNgCount.Text = ngCount.ToString();
        }

        // �e�X�g���ʂ��i�[����N���X
        public class TestResult
        {
            public string process { get; set; }
            public string judge { get; set; }
            public string inspectdate { get; set; }
        }

        // �e�X�g���ʂ̃v���Z�X�R�[�h�݂̂��i�[����N���X
        public class ProcessList
        {
            public string process { get; set; }
        }

        // �m�s�q�r�̏������s���A���茋�ʂ�Ԃ�
        private bool ntrsScanProcess(string id)
        {
            TfSQL tf = new TfSQL();
            DataTable dt = new DataTable();
            string log = string.Empty;
            string module = id;
            string mdlShort = VBS.Left(module, 16); // �ꎞ�I�ɂP�U�P�^�ɐݒ�

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
            // �P�D�p�X�̃v���Z�X�����擾
            var passResults = allResults.Where(r => r.judge == "PASS").Select(r => new ProcessList() { process = r.process }).OrderBy(r => r.process).ToList();
            foreach (var p in passResults) System.Diagnostics.Debug.Print(p.process);
            // �Q�D�P�Ɋ܂܂�Ȃ��t�F�C���̃v���Z�X�����擾
            var failResults = allResults.Where(r => r.judge == "FAIL").Select(r => new ProcessList() { process = r.process }).OrderBy(r => r.process).ToList();
            List<string> process = failResults.Select(r => r.process).Except(passResults.Select(r => r.process)).ToList();
            failResults = failResults.Where(r => process.Contains(r.process)).ToList();
            foreach (var p in failResults) System.Diagnostics.Debug.Print(p.process);
            // �R�D�P�ɂ��Q�ɂ��܂܂�Ȃ��A�e�X�g���ʂȂ��v���Z�X���擾����
            var skipResults = targetProcess.Replace("'", string.Empty).Split(',').ToList().Select(r => new ProcessList() { process = r.ToString() }).OrderBy(r => r.process).ToList();
            process = skipResults.Select(r => r.process).Except(passResults.Select(r => r.process)).ToList().Except(failResults.Select(r => r.process)).ToList();
            skipResults = skipResults.Where(r => process.Contains(r.process)).ToList();
            foreach (var p in skipResults) System.Diagnostics.Debug.Print(p.process);

            // �f�B�X�v���C�p�̃v���Z�X�����X�g�����H����
            string displayPass = string.Empty;
            string displayFail = string.Empty;
            string displayAll = string.Empty;   // ���O�p
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

            // �A�v���P�[�V�����X�N���[���ɁA�e�X�g���ʂ�\������
            if (passResults.Count == targetProcessCount)
            {
                string okImagePass = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\JigQuickDesk\JigQuickApp\images\" + okImageFile;
                pnlResult.BackgroundImageLayout = ImageLayout.Zoom;
                pnlResult.BackgroundImage = System.Drawing.Image.FromFile(okImagePass);

                // �n�j�J�E���g�̉��Z
                okCount += 1;
                txtOkCount.Text = okCount.ToString();

                // �o�`�r�r�̏ꍇ�́A�s�q�t�d��Ԃ�
                result = true;

                // ���̃��W���[���̃X�L�����ɂ��Ȃ��A�X�L�����p�e�L�X�g�{�b�N�X�̃e�L�X�g��I�����A�㏑���\�ɂ���
                txtChild.SelectAll();
            }
            else
            {
                string ngImagePath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\JigQuickDesk\JigQuickApp\images\" + ngImageFile;
                pnlResult.BackgroundImageLayout = ImageLayout.Zoom;
                pnlResult.BackgroundImage = System.Drawing.Image.FromFile(ngImagePath);

                // �m�f�J�E���g�̉��Z
                ngCount += 1;
                txtNgCount.Text = ngCount.ToString();

                // �e�`�h�k�̏ꍇ�́A�e�`�k�r�d��Ԃ�
                result = true;

                // ���̃��W���[���̃X�L�������X�g�b�v����߂��A�X�L�����p�e�L�X�g�{�b�N�X�𖳌��ɂ���
                txtChild.ReadOnly = true;
                txtChild.BackColor = Color.Red;

                // �A���[���ł̌x��
                soundAlarm();
            }

            // �A�v���P�[�V�����X�N���[���ƃf�X�N�g�b�v�t�H���_�̗����̗p�r�p�ɁA���t�ƃe�X�g���ʏڍו�������쐬
            log = Environment.NewLine + scanTime + "," + module + "," + displayAll;

            // �X�N���[���ւ̕\��
            txtResultDetail.Text = log.Replace(",", ",  ").Replace(Environment.NewLine, string.Empty);

            // ���O�����݁F�������t�̃t�@�C�������݂���ꍇ�͒ǋL���A���݂��Ȃ��ꍇ�̓t�@�C�����쐬�ǋL����iAppendAllText ������Ă����j
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