using System;
using System.Drawing;
using System.Windows.Forms;

namespace TransportationClient
{
    public partial class Main : Form
    {
        private bool cn = false;
        public Main()
        {
            InitializeComponent();
            openCN();
            FillingTables();
        }
        private void openCN()
        {
            Lib.OpenConnect();
            connection.ForeColor = Color.BlueViolet;
            connection.Text = "Активно";
            cn = true;
            RefreshBtns();
        }
        private void closeCN()
        {
            Lib.CloseConnect();
            connection.ForeColor = Color.Red;
            connection.Text = "Неактивно";
            cn = false;
            RefreshBtns();
        }
        private void FillingTables()
        {
            foreach (var table in Lib.tables)
            {
                Table.Items.Add(table);
                tableRpt.Items.Add(table);
            }
        }
        private void openBtn_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(Table.Text))
            {
                openTable(false);
            }
            else
            {
                MessageBox.Show("Ошибка. Не задана таблица.");
            }
        }
        private void openTable(bool rpt)
        {
            if (rpt)
            {
                Lib.OpenTable(tableRpt.Text);
            }
            else
            {
                Lib.OpenTable(Table.Text);
            }
            DGVTable.DataSource = Lib.dt;

            // Autosizing by DGV
            for (int i = 0; i <= DGVTable.Columns.Count - 1; i++)
                if (i != DGVTable.Columns.Count - 1)
                    DGVTable.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                else
                    DGVTable.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            for (int i = 0; i <= DGVTable.Columns.Count - 1; i++)
            {
                int colw = DGVTable.Columns[i].Width;

                DGVTable.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;

                DGVTable.Columns[i].Width = colw;
            }
        }
        private void createRptBtn_Click(object sender, EventArgs e)
        {
            if (tableRadio.Checked)
            {
                openTable(true);
                Lib.ClientUtils.CreateRep(DGVTable);
                return;
            }
            if (queryRadio.Checked)
            {
                SwitchQ(queryRpt.Text);
                Lib.ClientUtils.CreateRep(DGVQueries);
                return;
            }
            if (customRadio.Checked)
            {
                CreateCustomQ(customRpt.Text);
                Lib.ClientUtils.CreateRep(DGVQueries);
                return;
            }
        }
        private void SwitchQ(string textboxText)
        {
            string sql;

            switch (textboxText)
            {
                case "С точным совпадением":
                    sql = Lib.exactMatchQ;
                    openQuery(sql);
                    break;
                case "С неточным совпадением":
                    sql = Lib.notExactMatchQ;
                    openQuery(sql);
                    break;
                case "С группировкой":
                    sql = Lib.groupQ;
                    openQuery(sql);
                    break;
                case "С вычисляемым полем":
                    sql = Lib.calcFieldQ;
                    openQuery(sql);
                    break;
                case "Вычисление тарифа по дате":
                    sql = Lib.ClientUtils.ModalTariffQ();
                    openQuery(sql);
                    break;
                default:
                    break;
            }
        }
        private void CreateCustomQ(string sqlText)
        {
            try
            {
                openQuery(sqlText);
            }
            catch
            {
                MessageBox.Show("Произошла ошибка, проверьте ваш запрос.");
            }
        }
        private void RefreshBtns()
        {
            createBtn.Enabled = cn;
            closeBtn.Enabled = cn;
            deleteBtn.Enabled = cn;
            openBtn.Enabled = cn;
            opencnBtn.Enabled = !cn;
            updBtn.Enabled = cn;
            createQBtn.Enabled = cn;
            createRptBtn.Enabled = cn;
        }
        private void openQuery(string sql)
        {
            Lib.CreateQuery(sql);
            DGVQueries.DataSource = Lib.dtQ;

            // Autosizing by DGV
            for (int i = 0; i <= DGVQueries.Columns.Count - 1; i++)
                if (i != DGVQueries.Columns.Count - 1)
                    DGVQueries.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                else
                    DGVQueries.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            for (int i = 0; i <= DGVQueries.Columns.Count - 1; i++)
            {
                int colw = DGVQueries.Columns[i].Width;

                DGVQueries.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;

                DGVQueries.Columns[i].Width = colw;
            }
        }
        

        private void createQBtn_Click(object sender, EventArgs e)
        {
            if (custom.Checked)
            {
                CreateCustomQ(sqlBox.Text);
                return;
            }
            SwitchQ(Query.Text);
        }

        private void поТаблицеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Lib.ClientUtils.CreateRep(DGVTable);
        }

        private void поПользовательскомуЗапросуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Lib.ClientUtils.CreateRep(DGVQueries);
        }

        private void поЗапросуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Lib.ClientUtils.CreateRep(DGVQueries);
        }

        private void создатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            createQBtn_Click(sender, e);
        }

        private void создатьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            createBtn_Click(sender, e);
        }

        private void createBtn_Click(object sender, EventArgs e)
        {
            AddModal fAdd = new AddModal();
            int split = Lib.splitter;
            TextBox[] txt = new TextBox[Lib.names.Length];
            string[] values = new string[Lib.names.Length];

            for (int i = 0; i < Lib.names.Length; i++)
            {
                Label namelabel = new Label();
                namelabel.Location = new Point(Lib.splitter * 6, split);
                namelabel.Text = Lib.names[i];
                fAdd.Controls.Add(namelabel);
                txt[i] = new TextBox();
                txt[i].Location = new Point(Lib.splitter, split);
                fAdd.Controls.Add(txt[i]);
                split += Lib.splitter * 2;
            }
            fAdd.ShowDialog();
            if (fAdd.DialogResult == DialogResult.OK)
            {
                for (int i = 0; i < txt.Length; i++)
                {
                    values[i] = txt[i].Text;
                }
                Lib.Insert(Table.Text, values);
                openTable(false);
            }
        }

        private void deleteBtn_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in DGVTable.SelectedRows)
            {
                int id = (int)row.Cells[0].Value;
                Lib.DeleteById(Table.Text, id);
                DGVTable.Rows.Remove(row);
            }
        }

        private void updBtn_Click(object sender, EventArgs e)
        {
            for (int rowI = 0; rowI < DGVTable.RowCount - 1; rowI++)
            {
                dynamic[] values = new dynamic[Lib.names.Length - 1];
                for (int i = 1; i < Lib.names.Length; i++)
                    values[i - 1] = DGVTable.Rows[rowI].Cells[i].Value;

                Lib.Update(Table.Text, values, (int)DGVTable.Rows[rowI].Cells[0].Value);
            }
        }

        private void opencnBtn_Click(object sender, EventArgs e)
        {
            openCN();
        }

        private void closeBtn_Click(object sender, EventArgs e)
        {
            closeCN();
        }
    }
}
