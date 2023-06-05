using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic;
using System.Data;
using System.Data.OleDb;
using Word = Microsoft.Office.Interop.Word;

namespace TransportationClient
{
    internal class Lib
    {
        public static DataSet DS;
        public static bool F;
        public static DataTable dt;
        public static DataTable dtQ;
        public static DataRow row;
        public static OleDbConnection MyConnect;
        public static OleDbDataAdapter dataAdapter;
        public static DataSet ds;
        public static DataSet dsQ;
        public static string s;
        public static string[] Head;
        public static string[,] Add;
        public static string[] names;
        public const int splitter = 30;
        public static OleDbCommandBuilder builder;
        const string DBPath = "D:/Studies/!TSTU/malkov/access/db.accdb;";
        public static string connectString = $"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={DBPath}";

        public static string exactMatchQ = "SELECT Должность, Клиент FROM Клиент WHERE Должность = 'Безработный'";
        public static string notExactMatchQ = "SELECT Индекс, Клиент FROM Клиент WHERE Индекс > '170001'";
        public static string groupQ = "SELECT Клиент.Город, Sum(Звонок.Длительность) AS Общая_Длительность FROM Звонок INNER JOIN Клиент ON Звонок.Код_Звонок = Клиент.Код_Звонок GROUP BY Клиент.Город HAVING Клиент.Город = 'Тверь'";
        public static string calcFieldQ = "SELECT Клиент, Страна + ', ' + Индекс + ', ' + Город + ', ' + Адрес AS Полный_адрес FROM Клиент";
        public static string tariffByPosQ = "";

        public static string reportTitle = "";


        /// <summary>
        /// Открытие таблицы
        /// </summary>
        /// <param name="NameTable">Имя таблицы</param>
        public static void OpenTable(string NameTable)
        {
            dataAdapter = new OleDbDataAdapter("SELECT * FROM " + NameTable, MyConnect);
            builder = new OleDbCommandBuilder(dataAdapter);
            dataAdapter.UpdateCommand = builder.GetUpdateCommand();
            dataAdapter.DeleteCommand = builder.GetDeleteCommand();
            dataAdapter.InsertCommand = builder.GetInsertCommand();
            ds = new DataSet();
            dataAdapter.Fill(ds);
            dt = ds.Tables[0];
            names = Caption(dt);
        }
        /// <summary>
        /// Открытие соединения
        /// </summary>
        public static void OpenConnect()
        {
            MyConnect = new OleDbConnection(connectString);
            MyConnect.Open();
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
            names = Caption(dtQ);
        }
        /// <summary>
        /// Удаление строки таблицы по первичному ключу
        /// </summary>
        /// <param name="tableName">Название таблицы</param>
        /// <param name="id">Первичный ключ</param>
        public static void DeleteById(string tableName, int id)
        {
            string sql = $"DELETE FROM {tableName} WHERE Код_{tableName} = {id}";
            OleDbCommand operation = new OleDbCommand();
            operation.CommandText = sql;
            operation.Connection = MyConnect;
            operation.ExecuteNonQuery();
        }
        /// <summary>
        /// Добавление строки таблицы
        /// </summary>
        /// <param name="tableName">Название таблицы</param>
        /// <param name="values">значения</param>
        public static void Insert(string tableName, string[] values)
        {
            string sql = $"INSERT INTO {tableName} ({string.Join(", ", names)}) VALUES ({FormatValues(values)})";
            OleDbCommand operation = new OleDbCommand();
            operation.CommandText = sql;
            operation.Connection = MyConnect;
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
            string sql = $"UPDATE {tableName} SET {FormatNamesValues(values)} WHERE {names[0]} = {id}";
            OleDbCommand operation = new OleDbCommand();
            operation.CommandText = sql;
            operation.Connection = MyConnect;
            operation.ExecuteNonQuery();
        }
        /// <summary>
        /// Форматирование значений для Insert запроса. Пример: value -> 'value'
        /// </summary>
        /// <param name="values">значения</param>
        /// <returns>Отформатированные значения</returns>
        static string FormatValues(string[] values)
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
        static string FormatNamesValues(dynamic[] values)
        {
            string[] result = new string[names.Length - 1];
            for (int i = 1; i < names.Length; i++)
            {
                result[i - 1] = names[i] + " = " + $"'{values[i - 1]}'";
            }
            return string.Join(", ", result);
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
        /// Получение текущей строки 
        /// </summary>
        /// <param name="Index">Индекс текущей строки</param>
        /// <returns>Строка</returns>
        public static string CurrentRecord(int Index)
        {
            string Str = null;
            object[] items = new object[dt.Columns.Count];
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                items[i] = Lib.dt.Rows[Index][i].ToString();
            }
            string[] Caption = Lib.Caption(dt);
            for (int i = 0; i < Lib.dt.Columns.Count; i++)
            {
                Str = Str + Caption[i] + "  :  " + items[i].ToString() + "\r\n";
            }
            return Str;
        }
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
    }
}
