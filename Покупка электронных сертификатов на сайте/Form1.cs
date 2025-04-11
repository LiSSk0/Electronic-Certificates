using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Покупка_электронных_сертификатов_на_сайте
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            dataGridView1.AllowUserToAddRows = false;
            updateTable();
        }
        string tmp = "";

        //Добавление данных в таблицу
        private void button2_Click(object sender, EventArgs e)
        {
            DateTime dd;
            double sum;
            string num = "";
            string num_trans = "";

            if (dateTimePicker1.Value == null)
            {
                MessageBox.Show("Выберите дату", "Неверный формат даты", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (String.IsNullOrEmpty(textBox1.Text.Trim()))
            {
                MessageBox.Show("Введите сумму сертификата", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (String.IsNullOrEmpty(textBox3.Text.Trim()))
            {
                MessageBox.Show("Введите номер транзакции", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            try
            {
                dd = dateTimePicker1.Value;
                sum = Convert.ToDouble(textBox1.Text.Trim());
                if (!String.IsNullOrEmpty(textBox2.Text.Trim())) num = textBox2.Text;
            }
            catch (Exception exp)
            {
                MessageBox.Show($"Ввозникла ошибка преобразования данных\nПодробнее: {exp.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            num_trans = textBox3.Text.Trim();
            string sqlExpression = "";
            if (num.Trim() == "")
                sqlExpression = $"INSERT INTO Inet_zakaz.dbo.sertifikat_pokupka (dd, summa, number,inn,idapt,transaction_nomer) VALUES ('{dd.Date.ToString("yyyy-MM-dd HH:mm:ss.fff")}',{sum},null,5902882744,10735,'{num_trans}')";
            else
                sqlExpression = $"INSERT INTO Inet_zakaz.dbo.sertifikat_pokupka (dd, summa, number,inn,idapt,transaction_nomer) VALUES ('{dd.Date.ToString("yyyy-MM-dd HH:mm:ss.fff")}',{sum},'{num}',5902882744,10735,'{num_trans}')";

            using (SqlConnection connection = DBUtils.GetDBConnection())
            {
                connection.Open();
                using (SqlCommand cmd = new SqlCommand(sqlExpression, connection))
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                    }
                    catch (SqlException ex)
                    {
                        MessageBox.Show($"Возникла ошибка при добавлении данных в таблицу\nПодробнее: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                connection.Close();

            }
            MessageBox.Show($"Данные успешно добавлены в таблицу\n", "Команда выполнена", MessageBoxButtons.OK, MessageBoxIcon.Information);
            updateTable();
        }

        public void updateTable()
        {
            dataGridView1.CurrentCell = null;

            dataGridView1.Rows.Clear();

            DataTable tab = new DataTable();

            using (SqlConnection conn = DBUtils.GetDBConnection())
            {
                conn.Open();
                using (SqlCommand comm = new SqlCommand())
                {
                    comm.CommandText = "select * from Inet_zakaz.dbo.sertifikat_pokupka";
                    comm.Connection = conn;
                    tab.Load(comm.ExecuteReader());
                }
                conn.Close();
            }

            foreach (DataRow row in tab.Rows)
            {
                dataGridView1.Rows.Add(row[0], row[7], row[5], row[1], row[2]);
            }

        }
        //Удаление данных из таблицы
        private void button1_Click(object sender, EventArgs e)
        {
            delete();
        }

        public void delete()
        {
            if (dataGridView1.CurrentRow == null)
                return;

            int id = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());

            if (id == 0)
                return;

            if (MessageBox.Show($"Вы уверены, что хотите удалить строку с идентификатором {id} ?", "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            string sqlExpression = $"DELETE FROM Inet_zakaz.dbo.sertifikat_pokupka WHERE id={id}";

            using (SqlConnection connection = DBUtils.GetDBConnection())
            {
                connection.Open();
                using (SqlCommand cmd = new SqlCommand(sqlExpression, connection))
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                    }
                    catch (SqlException ex)
                    {
                        MessageBox.Show($"Ввозникла ошибка при удалении данных из таблицы\nПодробнее: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                connection.Close();

            }
            MessageBox.Show($"Данные успешно удалены из таблицы\n", "Команда выполнена", MessageBoxButtons.OK, MessageBoxIcon.Information);
            updateTable();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }


        private void dataGridView1_KeyPress(object sender, KeyPressEventArgs e)
        {
           

        }

        //редактирование данных в таблице
        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {


        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

            if (dataGridView1.CurrentCell == null)
                return;

            string sqlExpression = "";
            int id = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            string num = "";

            //редактируем номер транзакции
            if (dataGridView1.CurrentCell.ColumnIndex == 1)
            {
                sqlExpression = $"update Inet_zakaz.dbo.sertifikat_pokupka set transaction_nomer=";
            }
            //редактируем номер заказа
            if (dataGridView1.CurrentCell.ColumnIndex == 2)
            {
                sqlExpression = $"update Inet_zakaz.dbo.sertifikat_pokupka set number=";
            }

            try
            {
                if (dataGridView1.CurrentCell.Value != null)
                    num = dataGridView1.CurrentCell.Value.ToString();
            }
            catch (Exception exp)
            {
                MessageBox.Show($"Ошибка\nПодробнее: {exp.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                updateTable();
                return;
            }

            if (tmp.Trim().Equals(num.Trim()))
                return;

            if (id == 0)
                return;

            if (num.Trim() == "" && dataGridView1.CurrentCell.ColumnIndex == 2)
                sqlExpression += $"null where id = {id}";
            else
                sqlExpression += $"'{num}' where id = {id}";

            if (num.Trim() == "" && dataGridView1.CurrentCell.ColumnIndex == 1)
            {
                MessageBox.Show($"Поле \"Номер транзакции\" не должено быть пустым.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                updateTable();
                return;
            }

            using (SqlConnection connection = DBUtils.GetDBConnection())
            {
                connection.Open();
                using (SqlCommand cmd = new SqlCommand(sqlExpression, connection))
                {
                    try
                    {
                        cmd.ExecuteNonQuery();
                    }
                    catch (SqlException ex)
                    {
                        MessageBox.Show($"Ввозникла ошибка при попытки изменить данные в таблице\nПодробнее: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
                connection.Close();

            }
            MessageBox.Show($"Данные успешно изменены\n", "Команда выполнена", MessageBoxButtons.OK, MessageBoxIcon.Information);
            updateTable();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            tmp = dataGridView1.CurrentCell.Value.ToString();
        }

        // Вывод данных в Excel
        private void btnExcel_Click(object sender, EventArgs e)
        {
            DateTime ddFrom = dateTP_excel_from.Value.Date;
            DateTime ddTo = dateTP_excel_to.Value.Date;

            // Проверка корректности введённого периода
            if (ddFrom == null || ddTo == null || ddFrom > ddTo)
            {
                MessageBox.Show("Введён неверный период", "Выберите дату", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Получаем подтверждение от пользователя
            if (MessageBox.Show($"Вывести данные в Excel с {ddFrom:dd.MM.yyyy} по {ddTo:dd.MM.yyyy} ?", "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;


            DataTable tab = new DataTable();

            try
            {
                using (SqlConnection conn = DBUtils.GetDBConnection())
                {
                    conn.Open();  // закрывать необязательно, т.к. используем using
                    using (SqlCommand comm = new SqlCommand())
                    {
                        comm.CommandText = @"select id, transaction_nomer, number, dd, summa from Inet_zakaz.dbo.sertifikat_pokupka where dd between @from and @to";

                        comm.Parameters.AddWithValue("@from", ddFrom);
                        comm.Parameters.AddWithValue("@to", ddTo);
                        comm.Connection = conn;

                        tab.Load(comm.ExecuteReader());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при получении данных из БД:\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (tab.Rows.Count == 0)
            {
                MessageBox.Show("Нет данных за выбранный период", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Вывод в Excel
            Excel.Application excelApp = new Excel.Application();
            if (excelApp == null)
            {
                MessageBox.Show("Excel не установлен");
                return;
            }

            excelApp.Visible = false;
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet worksheet = workbook.Sheets[1];

            // Вывод заголовков
            for (int i = 0; i < tab.Columns.Count; i++)
                worksheet.Cells[1, i + 1] = tab.Columns[i].ColumnName;

            // Изменение ширины столбцов
            worksheet.Columns["B"].ColumnWidth = 16;
            worksheet.Columns["C"].ColumnWidth = 16;
            worksheet.Columns["D"].ColumnWidth = 20;
            worksheet.Columns["E"].ColumnWidth = 12;

            // Выравнивание столбцов по правому краю
            worksheet.Columns["A"].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            worksheet.Columns["B"].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            worksheet.Columns["C"].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            worksheet.Columns["D"].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            worksheet.Columns["E"].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;

            // Вывод данных
            for (int i = 0; i < tab.Rows.Count; i++)
            {
                for (int j = 0; j < tab.Columns.Count; j++)
                {
                    object value = tab.Rows[i][j];
                    if (value != null)
                        worksheet.Cells[i + 2, j + 1] = value.ToString();
                    else
                        worksheet.Cells[i + 2, j + 1] = "";
                }
            }

            MessageBox.Show("Данные успешно экспортированы в Excel", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            excelApp.Visible = true;
        }
    }
}
