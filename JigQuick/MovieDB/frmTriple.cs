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
        // �R���t�B�O�t�@�C���ƁA�o�̓e�L�X�g�t�@�C���́A�f�X�N�g�b�v�̎w��̃t�H���_�ɕۑ�����
        string appconfig = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\JigQuickDesk\JigQuickApp\info.ini";
        string outPath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\JigQuickDesk\ConverterTarget\";

        // ���������p�ϐ�
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

        // �X�N���[�����쒲���p�ϐ�
        string autoRegister = "ON";
        int rowPartsLot;
        int rowChild;


        // �R���X�g���N�^
        public frmTriple()
        {
            InitializeComponent();
        }

        // ���[�h���̏���
        private void frmInut_Load(object sender, EventArgs e)
        {
            // ���t�H�[���̕\���ꏊ���w��
            //this.Left = 30; // ��ʒ���
            //this.Top = 35;  // ��ʒ���

            // ���i���b�g�ێ��p�f�[�^�e�[�u���A�`���C���h�ێ��p�f�[�^�e�[�u���́A�s���ݒ���擾
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

            // �����o�^���[�h�̂n�m�E�n�e�e�擾�A�o�^�{�^���̐ݒ�
            //autoRegister = readIni("APPLICATION BEHAVIOR", "AUTOMATIC REGISTER", appconfig);

            // �e�폈���p�̃e�[�u���𐶐�
            dtPartsLot = new DataTable();
            dtTbi = new DataTable();
            dtChild = new DataTable();

            // �e�[�u���̒�`
            defineTables(ref dtPartsLot, ref dtTbi, ref dtChild);

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
        }

        // �T�u�v���V�[�W���F�c�s�̒�`
        private void defineTables(ref DataTable dt1, ref DataTable dt2, ref DataTable dt3)
        {
            // ���i���b�g�p�O���b�h�r���[����������
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


            // �`���C���h�p�O���b�h�r���[����������
            dt3.Columns.Add("child", Type.GetType("System.String"));

            for (int i = 0; i < rowChild; i++) dt3.Rows.Add(dt3.NewRow());
            dgvChild.DataSource = dt3;

            // �s�w�b�_�[�ɍs�ԍ���ǉ����A�s�w�b�_�����������߂���
            for (int i = 0; i < rowChild; i++) dgvChild.Rows[i].HeaderCell.Value = (i + 1).ToString();
            dgvChild.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
        }

        // �T�u�v���V�[�W���F���x���̐ݒ�
        private void setLabels()
        {
            txtModel.Text = readIni("LABEL DESCRIPTION", "MODEL", appconfig);
            txtProcess.Text = readIni("LABEL DESCRIPTION", "PROCESS TEXT BOX", appconfig);
            lblPartsLot.Text = readIni("LABEL DESCRIPTION", "PARTS LOT LABEL", appconfig);
            lblChild.Text = readIni("LABEL DESCRIPTION", "CHILD LABEL", appconfig);
            lblParent.Text = readIni("LABEL DESCRIPTION", "PARENT LABEL", appconfig);
            txtParent.Enabled = parentuse == "ON" ? true : false;
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

            resetViewColor(ref dgvPartsLot);
        }

        // �b�g�h�k�c�e�L�X�g�{�b�N�X�A�X�L�������̏���
        private void txtChild_KeyDown(object sender, KeyEventArgs e)
        {
            // �o�[�R�[�h�����̃G���^�[�L�[�ȊO�͏������Ȃ�
            if (e.KeyCode != Keys.Enter) return;

            // �󕶎��̏ꍇ�͏������Ȃ�
            string scan = txtChild.Text;
            if (scan == string.Empty) return;

            // �`���C���h�O���b�h�r���[�̌��݂̃Z���s�ƁA���̎��̍s���i�[����
            int r = dgvChild.CurrentCell.RowIndex;
            int y = r < rowChild - 1 ? r + 1 : rowChild - 1;

            // �O���b�g�r���[�ɕ\������
            dtChild.Rows[r][0] = scan;
            txtChild.Text = string.Empty;
            dgvChild.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dgvChild.CurrentCell = dgvChild[0, y];

            resetViewColor(ref dgvChild);

            // �`���C���h�����݂���ꍇ�́A�`���C���h�X�L�����p�A�܂��́A�h�b�L���O�W�O�����X�L�����p�ɁA�J�[�\�����ړ�
            if (r == rowChild - 1 && parentuse == "ON")
            {
                txtParent.Focus();
            }
            // �`���C���h�����݂��Ȃ��ꍇ�́A�o�^���s���i�y�A�����g�X�L�������Ɠ��l�̏����j
            else if (r == rowChild - 1 && parentuse == "OFF")
            {
                txtParent_KeyDown(sender, e);
            }
        }

        // �y�A�����g�X�L�������̏���
        private void txtParent_KeyDown(object sender, KeyEventArgs e)
        {
            // �o�[�R�[�h�����̃G���^�[�L�[�ȊO�͏������Ȃ�
            if (e.KeyCode != Keys.Enter) return;

            // �󕶎��̏ꍇ�͏������Ȃ�
            if (txtParent.Text == string.Empty�@&& parentuse == "ON") return;

            // �d������ы󔒂��}�[�L���O����
            colorViewForDuplicate(ref dgvPartsLot, "part_code");
            colorViewForDuplicate(ref dgvChild, "child");

            // ���i���b�g�s�J�E���g���q�n�v�ݒ�̏ꍇ�̂݁A�����o�^���s��
            countPartsLot = dgvPartsLot.Rows.Cast<DataGridViewRow>().Where(r => r.Cells[0].Value.ToString() != string.Empty)
                    .Select(r => r.Cells[0].Value).GroupBy(r => r).Count();
            if (dupPartsLot || countPartsLot != rowPartsLot)
            {
                MessageBox.Show("Parts lot information is not enough.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // ���i���b�g�s�J�E���g���q�n�v�ݒ�̏ꍇ�̂݁A�����o�^���s��
            countChild = dgvChild.Rows.Cast<DataGridViewRow>().Where(r => r.Cells[0].Value.ToString() != string.Empty)
                    .Select(r => r.Cells[0].Value).GroupBy(r => r).Count();
            if (countChild != rowChild)�@// if (dupChild || countChild != rowChild)
            {
                MessageBox.Show("Child information is not enough.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // �����o�^���[�h�n�m�E�n�e�e�̏ꍇ����
            if (autoRegister == "ON")
            {
                outPutTbiInfo();
            }
        }

        // �o�^�{�^���������A�s�a�h�e�[�u���ւ̓o�^
        private void btnRegister_Click(object sender, EventArgs e)
        {
            // ���i���b�g�s�J�E���g���q�n�v�ݒ�̏ꍇ�̂݁A�����o�^���s��
            countPartsLot = dgvPartsLot.Rows.Cast<DataGridViewRow>().Where(r => r.Cells[0].Value.ToString() != string.Empty)
                    .Select(r => r.Cells[0].Value).GroupBy(r => r).Count();
            if (dupPartsLot || countPartsLot != rowPartsLot)
            {
                MessageBox.Show("Parts lot information is not enough.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //// ���i���b�g�s�J�E���g���q�n�v�ݒ�̏ꍇ�̂݁A�����o�^���s��
            //countChild = dgvChild.Rows.Cast<DataGridViewRow>().Where(r => r.Cells[0].Value.ToString() != string.Empty)
            //        .Select(r => r.Cells[0].Value).GroupBy(r => r).Count();
            //if (countChild <= 0)�@// if (dupChild || countChild != rowChild)
            //{
            //    MessageBox.Show("Child information is missing.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}

            if (parentuse == "ON" && txtParent.Text == string.Empty)�@// if (dupChild || countChild != rowChild)
            {
                MessageBox.Show(lblParent.Text + " information is missing.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            outPutTbiInfo();
        }

        // �T�u�v���V�[�W���F�s�a�h�e�[�u���ւ̓o�^
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

            // �`���C���h�P�A�`���C���h�Q�A�y�A�����g�̃}�b�`���O���s���i���@�\�̐ݒ�A�n�m�̏ꍇ�j
            if (c1c2prMatching == "ON")
            {
                wrongScanOrder = matchScanOrder(string.Empty, string.Empty, txtParent.Text);
                if (wrongScanOrder)
                {
                    MessageBox.Show("Case pallet barcode is wrong." + Environment.NewLine + "Please clear and re-scan all.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            // �����̕��i���b�g���ɁA�����̂b�g�h�k�c����R�t���A���i���b�g�����~�b�g�h�k�c�����́A���R�[�h���쐬����
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
                    // �y�A�����g�̎g�p���Ȃ��ꍇ�i�c�|�S�j�́A�W�O�|�W�V������͋�
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

                    // �{���A�b�g�h�k�c������r�d�q�m�n�A�o�`�q�d�m�s���E��k�n�s�A�֏o�͂��邪�A�l�`�h�m�W�O���́A�t�ŏo�͂���
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
                cumCount += tempCount;
                txtCount.Text = cumCount.ToString();

                // ���̃`���C���h�X�L�����̂��߂̏���
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

        // �`���C���h�N���A�{�^���������A�`���C���h�O���b�h�r���[�̃N���A
        private void btnClerChild_Click(object sender, EventArgs e)
        {
            clerChildView();
        }

        // �T�u�v���V�[�W���F�`���C���h�O���b�h�r���[�̃N���A
        private void clerChildView()
        {
            for (int i = 0; i < rowChild; i++) dgvChild[0, i].Value = string.Empty;
            dgvChild.CurrentCell = dgvChild[0, 0];
            txtChild.Focus();
        }

        // �y�A�����g�N���A�{�^���������A�`���C���h�O���b�h�r���[�̃N���A
        private void button1_Click(object sender, EventArgs e)
        {
            txtParent.Text = string.Empty;
        }

        // �T�u�v���V�[�W���F�d�����i���b�g�E�󃌃R�[�h���}�[�L���O����
        private void colorViewForDuplicate(ref DataGridView dgv, string field)
        {
            // ����̃t�B�[���h���w�肳�ꂽ�ꍇ�̂݁A�������s��
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

        // �T�u�v���V�[�W���F�O���b�g�r���[�̐F�����Z�b�g����
        private void resetViewColor(ref DataGridView dgv)
        {
            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                dgv.Rows[i].DefaultCellStyle.BackColor = Color.FromKnownColor(KnownColor.Window);
            }
            dupPartsLot = false;
            dupChild = false;
        }

        // �T�u�v���V�[�W���F�`���C���h�P�A�`���C���h�Q�A�y�A�����g�̃}�b�`���O���s��
        private bool matchScanOrder(string child1, string child2, string parent)
        {
            bool parMissing = parent.IndexOf(c1c2prList[2]) == -1 ? true : false;
            if (parMissing) soundAlarm();
            return parMissing;
        }

        // �T�u�v���V�[�W���F�r�d�k�d�b�s���\�b�h�p�L�[�Ɋ܂܂��A���C���h�J�[�h�𒲐�����
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


    }
}