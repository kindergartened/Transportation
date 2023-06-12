using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data;
using Microsoft.VisualBasic;
using Word = Microsoft.Office.Interop.Word;

namespace TransportationClient
{
    /// <summary>
    /// Основной класс работы с приложением и базой данных
    /// </summary>
    internal class Lib
    {
        public static DataSet ds;
        public static DataSet dsQ;
        public static DataTable dt;
        public static DataTable dtQ;
        public static OleDbConnection MyConnect;
        public static OleDbDataAdapter dataAdapter;
        public static List<string> tables;
        public static string[] names;
        public const int splitter = 30;
        const string DBPath = "D:/tstu/works/labs/malkov/Course/db.accdb;";
        public static string connectString = $"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={DBPath}";

        public static string exactMatchQ = "SELECT Должность, Клиент FROM Клиент WHERE Должность = 'Безработный'";
        public static string notExactMatchQ = "SELECT Индекс, Клиент FROM Клиент WHERE Индекс > '170001'";
        public static string groupQ = "SELECT Клиент.Город, Sum(Звонок.Длительность) AS Общая_Длительность FROM Звонок INNER JOIN Клиент ON Звонок.Код_Звонок = Клиент.Код_Звонок GROUP BY Клиент.Город HAVING Клиент.Город = 'Тверь'";
        public static string calcFieldQ = "SELECT Клиент, Страна + ', ' + Индекс + ', ' + Город + ', ' + Адрес AS Полный_адрес FROM Клиент";

        /// <summary>
        /// Открытие таблицы
        /// </summary>
        /// <param name="NameTable">Имя таблицы</param>
        public static void OpenTable(string NameTable)
        {
            dataAdapter = new OleDbDataAdapter("SELECT * FROM " + NameTable, MyConnect);
            ds = new DataSet();
            dataAdapter.Fill(ds);
            dt = ds.Tables[0];
            names = DBUtils.Caption(dt);
        }
        /// <summary>
        /// Открытие соединения
        /// </summary>
        public static void OpenConnect()
        {
            MyConnect = new OleDbConnection(connectString);
            MyConnect.Open();
            DBUtils.GetTableNames();
        }
        /// <summary>
        /// Закрытие соединения
        /// </summary>
        public static void CloseConnect()
        {
            MyConnect.Dispose();
        }
        /// <summary>
        /// Создание запроса
        /// </summary>
        /// <param name="Command">Строка команды</param>
        public static void CreateQuery(string Command)
        {
            dataAdapter = new OleDbDataAdapter(Command, MyConnect);
            dsQ = new DataSet();
            dataAdapter.Fill(dsQ);
            dtQ = dsQ.Tables[0];
            names = DBUtils.Caption(dtQ);
        }
        /// <summary>
        /// Удаление строки таблицы по первичному ключу
        /// </summary>
        /// <param name="tableName">Название таблицы</param>
        /// <param name="id">Первичный ключ</param>
        public static void DeleteById(string tableName, int id)
        {
            string sql = $"DELETE FROM {tableName} WHERE Код_{tableName} = {id}";
            OleDbCommand operation = new OleDbCommand
            {
                CommandText = sql,
                Connection = MyConnect
            };
            operation.ExecuteNonQuery();
        }
        /// <summary>
        /// Добавление строки таблицы
        /// </summary>
        /// <param name="tableName">Название таблицы</param>
        /// <param name="values">значения</param>
        public static void Insert(string tableName, string[] values)
        {
            string sql = $"INSERT INTO {tableName} ({string.Join(", ", names)}) VALUES ({DBUtils.FormatValues(values)})";
            OleDbCommand operation = new OleDbCommand
            {
                CommandText = sql,
                Connection = MyConnect
            };
            operation.ExecuteNonQuery();
        }
        /// <summary>
        /// Обновление строки таблицы по первичному ключу
        /// </summary>
        /// <param name="tableName">Название таблицы</param>
        /// <param name="values">значения</param>
        /// <param name="id">первичный ключ</param>
        public static void Update(string tableName, dynamic[] values, int id)
        {
            string sql = $"UPDATE {tableName} SET {DBUtils.FormatNamesValues(values)} WHERE {names[0]} = {id}";
            OleDbCommand operation = new OleDbCommand
            {
                CommandText = sql,
                Connection = MyConnect
            };
            operation.ExecuteNonQuery();
        }
        /// <summary>
        /// Дополнительные методы для работы с базами данных.
        /// </summary>
        private class DBUtils
        {
            /// <summary>
            /// Получение заголовков таблиц
            /// </summary>
            public static void GetTableNames()
            {
                tables = new List<string>();
                string[] restriction = new string[4];
                restriction[3] = "Table";
                DataTable schemas = MyConnect.GetSchema("Tables", restriction);
                for (int i = 0; i < schemas.Rows.Count; i++)
                {
                    tables.Add(schemas.Rows[i][2].ToString());
                }
            }
            /// <summary>
            /// Получение заголовков столбцов таблицы
            /// </summary>
            /// <param name="DT">Таблица</param>
            /// <returns>Массив строк</returns>
            public static string[] Caption(DataTable DT)
            {
                string[] StrName = new string[DT.Columns.Count];
                for (int i = 0; i < DT.Columns.Count; i++)
                {
                    StrName[i] = DT.Columns[i].Caption;
                }
                return StrName;
            }
            /// <summary>
            /// Форматирование значений для Insert запроса. Пример: value -> 'value'
            /// </summary>
            /// <param name="values">значения</param>
            /// <returns>Отформатированные значения</returns>
            public static string FormatValues(string[] values)
            {
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = "'" + values[i] + "'";
                }
                return string.Join(", ", values);
            }
            /// <summary>
            /// Форматирование значений для update запроса. Пример: column, value -> column = 'value'
            /// </summary>
            /// <param name="values">значения</param>
            /// <returns>Отформатированные значения</returns>
            public static string FormatNamesValues(dynamic[] values)
            {
                string[] result = new string[names.Length - 1];
                for (int i = 1; i < names.Length; i++)
                {
                    result[i - 1] = names[i] + " = " + $"'{values[i - 1]}'";
                }
                return string.Join(", ", result);
            }
        }
        /// <summary>
        /// Дополнительные методы для работы с клиентом приложения
        /// </summary>
        public class ClientUtils
        {
            /// <summary>
            /// Модалка для получения sql строки
            /// </summary>
            /// <returns>sql string</returns>
            public static string ModalTariffQ()
            {
                string date = Interaction.InputBox("Введите дату для определения тарифа по должностям, дата вводится в формате M/D/YYYY.", "Дата", "12/9/2022");
                int coef = int.Parse(date.Split('/')[1]) / 2 == 0 ? 2 : 3;
                return $"SELECT Клиент.Должность, Sum(Звонок.Длительность)*{coef} AS Тариф FROM Звонок INNER JOIN Клиент ON Звонок.Код_Звонок = Клиент.Код_Звонок GROUP BY Клиент.Должность, Звонок.Дата_Звонка HAVING Звонок.Дата_Звонка = #{date}#";
            }

            public static void CreateRep(DataGridView dgv)
            {
                if (dgv == null)
                {
                    MessageBox.Show("Отсутствуют данные для печати!");
                    return;
                }
                int rowc = dgv.RowCount;
                int colc = dgv.ColumnCount;
                string[,] rep = new string[rowc, colc];
                for (int i = 0; i < rowc - 1; i++)
                    for (int j = 0; j < colc; j++)
                        rep[i, j] = dgv.Rows[i].Cells[j].Value.ToString();
                Word.Application application = new Word.Application();
                Object missing = Type.Missing;
                application.Documents.Add(ref missing, ref missing, ref missing, ref missing);
                Word.Document document = application.ActiveDocument;
                Word.Range range = application.Selection.Range;
                Object behiavor = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
                Object autoFitBehiavor = Word.WdAutoFitBehavior.wdAutoFitFixed;
                document.Tables.Add(range, rowc, colc, ref behiavor, ref autoFitBehiavor);
                for (int i = 0; i < Lib.names.Length; i++)
                    document.Tables[1].Cell(1, i + 1).Range.Text = Lib.names[i].ToString();
                for (int i = 1; i < rowc; i++)
                    for (int j = 1; j < colc + 1; j++)
                        document.Tables[1].Cell(i + 1, j).Range.Text = rep[i - 1, j - 1].ToString();
                application.Visible = true;
            }
        }
    }
}
