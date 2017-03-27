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
    public partial class frmTipTail : Form
    {
        // �R���t�B�O�t�@�C���ƁA�o�̓e�L�X�g�t�@�C���́A�f�X�N�g�b�v�̎w��̃t�H���_�ɕۑ�����
        string appconfig = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\JigQuickDesk\JigQuickApp\info.ini";
        string outPath = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\JigQuickDesk\ConverterTarget\";

        //���̑��A�񃍁[�J���ϐ�
        DataTable dtPartsLot;
        DataTable dtTbi;
        int rowPartsLot;
        int partsLotCount;
        long cumCount;
        bool sound;
        bool duplicate;
        bool wrongScanOrder;
        string autoRegister = "ON";
        string parentTextBoxFunc = "OFF";
        string c1c2prMatching = "ON";
        string processMode;
        string[] description;
        string[] partsLotBreakdown;
        string[] jigposition;
        string[] c1c2prList;

        // �R���X�g���N�^
        public frmTipTail()
        {
            InitializeComponent();
        }

        // ���[�h���̏���
        private void frmInut_Load(object sender, EventArgs e)
        {
            // ���t�H�[���̕\���ꏊ���w��
            //this.Left = 30; // ��ʒ���
            //this.Top = 35;  // ��ʒ���

            // ���i���b�g�ێ��p�f�[�^�e�[�u���̍s���ݒ���擾
            rowPartsLot = int.Parse(readIni("OTHERS", "PARTS LOT COUNT", appconfig));
            description = readIni("OTHERS", "DESCRIPTION", appconfig).Split(',');
            partsLotBreakdown = readIni("OTHERS", "PARTS LOT BREAKDOWN", appconfig).Split(',');
            jigposition = readIni("OTHERS", "JIG POSITION", appconfig).Split(',');
            processMode = readIni("LABEL DESCRIPTION", "PROCESS TEXT BOX", appconfig);
            c1c2prMatching = readIni("APPLICATION BEHAVIOR", "CH1, CH2, PR MATCHING", appconfig);
            c1c2prList = readIni("OTHERS", "CH1, CH2, PR LIST", appconfig).Split(','); 
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
                    if (processMode == "Docking (D-5-2)") cumCount = cumCount / 2;
                }
            }

            // �J�E���g�e�L�X�g�{�b�N�X�̒��g�̕\��
            txtCount.Text = cumCount.ToString();
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
            lblChild2.Text = readIni("LABEL DESCRIPTION", "CHILD LBL 2", appconfig);
            lblParent.Text = readIni("LABEL DESCRIPTION", "PARENT LABEL", appconfig);
            parentTextBoxFunc = readIni("APPLICATION BEHAVIOR", "PARENT TEXT BOX", appconfig);
            txtParent.Enabled = parentTextBoxFunc == "ON" ? true : false;
        }

        // �N���A�{�^���������̏���
        private void btnClerChild_Click(object sender, EventArgs e)
        {
            txtChild.Text = string.Empty;
            txtChild2.Text = string.Empty;
            txtParent.Text = string.Empty;
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
            if (txtChild.Text == string.Empty) return;

            // �H���ɂ���āA�e�L�X�g�{�b�N�X�̃J�[�\���ړ��p�^�[�����A�h�e���ŕύX����
            if (processMode == "Case Tip & Tail Holding (D-6)")
            {
                // �o�`�q�d�m�s�����܂��Ă��Ȃ��ꍇ�́A�o�`�q�d�m�s�X�L�����p�ɃJ�[�\�����ړ�
                if (txtParent.Text == string.Empty)
                {
                    txtParent.Focus();
                }
                // �b�g�h�k�c�Q�����܂��Ă��Ȃ��ꍇ�́A�b�g�h�k�c�Q�X�L�����p�ɃJ�[�\�����ړ�
                else if (txtChild2.Text == string.Empty)
                {
                    txtChild2.Focus();
                }
                // �e�A�q�Q�l�A�S���e�L�X�g����������ꍇ�̂݁A�o��
                else
                {
                    checkAndOutput();
                }
            }
            else
            {
                // �b�g�h�k�c�Q�����܂��Ă��Ȃ��ꍇ�́A�b�g�h�k�c�Q�X�L�����p�ɃJ�[�\�����ړ�
                if (txtChild2.Text == string.Empty)
                {
                    txtChild2.Focus();
                }
                // �o�`�q�d�m�s�����܂��Ă��Ȃ��ꍇ�́A�o�`�q�d�m�s�X�L�����p�ɃJ�[�\�����ړ�
                else if (txtParent.Text == string.Empty)
                {
                    txtParent.Focus();
                }
                // �e�A�q�Q�l�A�S���e�L�X�g����������ꍇ�̂݁A�o��
                else
                {
                    checkAndOutput();
                }
            }
        }

        // �b�g�h�k�c�Q�e�L�X�g�{�b�N�X�A�X�L�������̏���
        private void txtChild2_KeyDown(object sender, KeyEventArgs e)
        {
            // �o�[�R�[�h�����̃G���^�[�L�[�ȊO�͏������Ȃ�
            if (e.KeyCode != Keys.Enter) return;

            // �󕶎��̏ꍇ�͏������Ȃ�
            if (txtChild2.Text == string.Empty) return;

            // �b�g�h�k�c�����܂��Ă��Ȃ��ꍇ�́A�b�g�h�k�c�X�L�����p�ɃJ�[�\�����ړ�
            if (txtChild.Text == string.Empty)
            {
                txtChild.Focus();
            }
            // �o�`�q�d�m�s�����܂��Ă��Ȃ��ꍇ�́A�o�`�q�d�m�s�X�L�����p�ɃJ�[�\�����ړ�
            else if (txtParent.Text == string.Empty)
            {
                txtParent.Focus();
            }
            // �e�A�q�Q�l�A�S���e�L�X�g����������ꍇ�̂݁A�o��
            else
            {
                checkAndOutput();
            }
        }

        // �o�`�q�d�m�s�e�L�X�g�{�b�N�X�A�X�L�������̏���
        private void txtParent_KeyDown(object sender, KeyEventArgs e)
        {
            // �o�[�R�[�h�����̃G���^�[�L�[�ȊO�͏������Ȃ�
            if (e.KeyCode != Keys.Enter) return;

            // �󕶎��̏ꍇ�͏������Ȃ�
            if (txtParent.Text == string.Empty) return;

            // �H���ɂ���āA�e�L�X�g�{�b�N�X�̃J�[�\���ړ��p�^�[�����A�h�e���ŕύX����
            if (processMode == "Case Tip & Tail Holding (D-6)")
            {
                // �b�g�h�k�c�Q�����܂��Ă��Ȃ��ꍇ�́A�b�g�h�k�c�Q�X�L�����p�ɃJ�[�\�����ړ�
                if (txtChild2.Text == string.Empty)
                {
                    txtChild2.Focus();
                }
                // �b�g�h�k�c�����܂��Ă��Ȃ��ꍇ�́A�b�g�h�k�c�X�L�����p�ɃJ�[�\�����ړ�
                else if (txtChild.Text == string.Empty)
                {
                    txtChild.Focus();
                }
                // �e�A�q�Q�l�A�S���e�L�X�g����������ꍇ�̂݁A�o��
                else
                {
                    checkAndOutput();
                }
            }
            else
            {
                // �b�g�h�k�c�����܂��Ă��Ȃ��ꍇ�́A�b�g�h�k�c�X�L�����p�ɃJ�[�\�����ړ�
                if (txtChild.Text == string.Empty)
                {
                    txtChild.Focus();
                }
                // �b�g�h�k�c�Q�����܂��Ă��Ȃ��ꍇ�́A�b�g�h�k�c�Q�X�L�����p�ɃJ�[�\�����ړ�
                else if (txtChild2.Text == string.Empty)
                {
                    txtChild2.Focus();
                }
                // �e�A�q�Q�l�A�S���e�L�X�g����������ꍇ�̂݁A�o��
                else
                {
                    checkAndOutput();
                }
            }
         }

        // �T�u�v���V�[�W���F�e�L�X�g�t�@�C���o�͑O�̊m�F
        private void checkAndOutput()
        {
            // ���i���b�g�̏d������ы󔒂��}�[�L���O����
            colorViewForDuplicate(ref dgvPartsLot);

            // �n�j�J�E���g���q�n�v�ݒ�̏ꍇ�̂݁A�����o�^���s��
            partsLotCount = dgvPartsLot.Rows.Cast<DataGridViewRow>().Where(r => r.Cells[0].Value.ToString() != string.Empty)
                    .Select(r => r.Cells[0].Value).GroupBy(r => r).Count();

            if (duplicate || partsLotCount != rowPartsLot)
            {
                MessageBox.Show("Part lot info is not enough.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // �`���C���h�P�A�`���C���h�Q�A�y�A�����g�̃}�b�`���O���s���i���@�\�̐ݒ�A�n�m�̏ꍇ�j
            if (c1c2prMatching == "ON")
            {
                wrongScanOrder = matchScanOrder(txtChild.Text, txtChild2.Text, txtParent.Text);
                if (wrongScanOrder)
                {
                    MessageBox.Show("Scan order is wrong." + Environment.NewLine + "Please clear and re-scan all jigs.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            // �s�h�o���s�`�h�k�̏ꍇ�ƁA����ȊO�ŁA�o�͓��e���ꍇ����
            if (autoRegister == "ON")
            {
                if (processMode == "Case Tip & Tail Holding (D-6)")
                {
                    outputTbiInfoForTipTail();
                }
                else
                {
                    outputTbiInfoForDocking2();
                }
            }
        }

        // �T�u�v���V�[�W���F�e�L�X�g�t�@�C���o�́i�s�h�o���s�`�h�k�p�j
        private void outputTbiInfoForTipTail()
        {                           
            string[] sn = { txtChild.Text.Trim(), txtChild2.Text.Trim() };
            string lot = txtParent.Text.Trim() == string.Empty? "null" : txtParent.Text.Trim();
            string model = txtModel.Text.Trim();
            string date = DateTime.Today.ToString("yyyy/MM/dd");
            string time = DateTime.Now.ToString("HH:mm:ss");
            string cumRecords = string.Empty;

            // �s�h�o���s�`�h�k�̏ꍇ�A���ꂼ��̃W�O�ɑΉ����邻�ꂼ��̕����̕��i���b�g����R�t���A�Q���̃��R�[�h���쐬����
            for (int i = 0; i < sn.Length; i++)
            {
                string lotInf = (dgvPartsLot[3, i].Value.ToString().Replace(":", "_") + ":" + jigposition[i] + ":" + description[i] + ":" + dgvPartsLot[0, i].Value.ToString() + ":" +
                    dgvPartsLot[1, i].Value.ToString() + ":" + dgvPartsLot[2, i].Value.ToString().Replace(":", "_") + ":" + dgvPartsLot[4, i].Value.ToString())
                    .Replace(" ", "_").Replace(",", "_").Replace("'", "_").Replace(";", "_").Replace("\"", "_");

                // ���i�V���A���̕��i�R���t�B�O���́A�؂�̂Ă�i�P�V���̂ݕێ�����j
                if (sn[i].IndexOf("+") >= 1) sn[i] = VBS.Left(sn[i], 17);

                string newRecord = sn[i] + "," + lot + "," + model + "," + date + "," + time + "," + lotInf;
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

            // ���̃X�L�����ɔ����A
            txtChild.Text = string.Empty;
            txtChild2.Text = string.Empty;
            txtParent.Text = string.Empty;
            txtChild.Focus();
        }

        // �T�u�v���V�[�W���F�e�L�X�g�t�@�C���o�́i�c�n�b�j�h�m�f�Q�p�j
        private void outputTbiInfoForDocking2()
        {
            string[] sn = { txtParent.Text.Trim(), txtChild2.Text.Trim() };
            string lot = txtChild.Text.Trim() == string.Empty ? "null" : txtChild.Text.Trim();
            string model = txtModel.Text.Trim();
            string date = DateTime.Today.ToString("yyyy/MM/dd");
            string time = DateTime.Now.ToString("HH:mm:ss");
            string cumRecords = string.Empty;

            // �c�n�b�j�h�m�f�|�Q�ꍇ�A�e�q�`�l�d �i�h�f�ɂ͕��i���b�g���t���A�c�n�b�j�h�m�f �t�o�o�d�q�ɂ͕��i���b�g�Ȃ��ŁA�Q���̃��R�[�h���쐬����
            for (int i = 0; i < sn.Length; i++)
            {
                string lotInf = string.Empty;
                if (i == 0)
                {
                    lotInf = ("null" + ":" + jigposition[i] + ":" + description[i] + ":" + "null" + ":" + "null" + ":" + "null" + ":" + "2000/01/01")
                        .Replace(" ", "_").Replace(",", "_").Replace("'", "_").Replace(";", "_").Replace("\"", "_");
                }
                else
                {
                    lotInf = (dgvPartsLot[3, 0].Value.ToString().Replace(":", "_") + ":" + jigposition[i] + ":" + description[i] + ":" + dgvPartsLot[0, 0].Value.ToString() + ":" +
                        dgvPartsLot[1, 0].Value.ToString() + ":" + dgvPartsLot[2, 0].Value.ToString().Replace(":", "_") + ":" + dgvPartsLot[4, 0].Value.ToString())
                        .Replace(" ", "_").Replace(",", "_").Replace("'", "_").Replace(";", "_").Replace("\"", "_");
                }

                // ���i�V���A���̕��i�R���t�B�O���́A�؂�̂Ă�i�P�V���̂ݕێ�����j
                if (sn[i].IndexOf("+") >= 1) sn[i] = VBS.Left(sn[i], 17);

                string newRecord = sn[i] + "," + lot + "," + model + "," + date + "," + time + "," + lotInf;
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

            // ���̃X�L�����ɔ����A
            txtChild.Text = string.Empty;
            txtChild2.Text = string.Empty;
            txtParent.Text = string.Empty;
            txtChild.Focus();
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

        // �T�u�v���V�[�W���F�`���C���h�P�A�`���C���h�Q�A�y�A�����g�̃}�b�`���O���s��
        private bool matchScanOrder(string child1, string child2, string parent)
        {
            bool ch1Missing = child1.IndexOf(c1c2prList[0]) == -1 ? true : false;
            bool ch2Missing = child2.IndexOf(c1c2prList[1]) == -1 ? true : false;
            bool parMissing = parent.IndexOf(c1c2prList[2]) == -1 ? true : false;
            bool total = (ch1Missing | ch2Missing | parMissing);
            if (total) soundAlarm();
            return total;
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