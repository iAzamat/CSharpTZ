using Aspose.Words;
using Microsoft.Data.Sqlite;
using System.Data;
using System.Xml.XPath;
using ClosedXML.Excel;

namespace ReadExcelFileApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            createFolders();
        }

        OpenFileDialog file = new OpenFileDialog();
        DataTable dtExcel = new DataTable();
        String pathData = "./data";
        String pathReport = "./report";
        String dbName = "userdata";
        String xmlName = "userdata";
        String reportXML = "reportXML";
        String reportSQL = "reportSQL";

        private void btnChoose_Click(object sender, EventArgs e)
        {
            if (file.ShowDialog() == DialogResult.OK)
            {
                string fileExt = Path.GetExtension(file.FileName);
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        //dtExcel = ReadExcel(file.FileName);
                        dtExcel = ExcelToDataTable(file.FileName);
                        button1.Enabled = true;
                        button2.Enabled = true;
                        button3.Enabled = true;
                        button4.Enabled = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }


    /*    private DataTable ReadExcel(string fileName)
        {
            WorkBook workbook = WorkBook.Load(fileName);
            WorkSheet sheet = workbook.DefaultWorkSheet;
            return sheet.ToDataTable(true);
        }*/

        public DataTable ExcelToDataTable(string filePath)
        {
            DataTable dt = new DataTable("Лист1");

            using (XLWorkbook workBook = new XLWorkbook(filePath))
            {
                IXLWorksheet workSheet = workBook.Worksheet(1);
                bool firstRow = true;

                foreach (IXLRow row in workSheet.Rows())
                {
                    if (firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Columns.Add(cell.Value.ToString());
                        }
                        firstRow = false;
                    }
                    else
                    {
                        dt.Rows.Add();
                        int i = 0;

                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                            i++;
                        }
                    }
                }
            }

            return dt;
        }


        private void createFolders()
        {
            if (!Directory.Exists(pathData))
            {
                Directory.CreateDirectory(pathData);
            }
            if (!Directory.Exists(pathReport))
            {
                Directory.CreateDirectory(pathReport);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                dtExcel.WriteXml($"{pathData}/{xmlName}.xml", XmlWriteMode.IgnoreSchema);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
            textBox1.AppendText("Файл xml успешно создан\r\n");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                button2.Enabled = false;
                CreateDB(dbName);

                // перебор всех строк таблицы
                foreach (DataRow row in dtExcel.Rows)
                {
                    // получаем все ячейки строки
                    var cells = row.ItemArray;
                    String firstName = cells[0].ToString();
                    String lastName = cells[1].ToString();
                    String gender = cells[2].ToString();
                    Double age = Double.Parse(cells[3].ToString());
                    String status = cells[4].ToString();
                    InsertData(dbName, firstName, lastName, gender, age, status);
                }
                textBox1.AppendText("База данных заполнена\r\n");
                button2.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void CreateDB(string dbName)
        {
            try
            {
                using (var connection = new SqliteConnection($"Data Source={pathData}/{dbName}.db"))
                {
                    connection.Open();

                    SqliteCommand command = new SqliteCommand();
                    command.Connection = connection;
                    command.CommandText = "DROP TABLE IF EXISTS Users";
                    command.ExecuteNonQuery();

                    command.CommandText = "CREATE TABLE Users(_id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT UNIQUE, " +
                        "firstName VARCHAR(50) NOT NULL, " +
                        "lastName VARCHAR(50) NOT NULL, " +
                        "gender VARCHAR(10) NOT NULL," +
                        "age INTEGER NOT NULL," +
                        "status VARCHAR(20) NOT NULL)";
                    command.ExecuteNonQuery();

                    textBox1.AppendText("Таблица Users создана\r\n");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        public void InsertData(String dbName, String firstName, String lastName, String gender, Double age, String status)
        {
            try
            {
                using (var connection = new SqliteConnection($"Data Source={pathData}/{dbName}.db"))
                {
                    connection.Open();

                    string sql = "INSERT INTO Users (firstName, lastName, gender, age, status) VALUES " +
                                "(@firstName, @lastName, @gender, @age, @status)";

                    using (SqliteCommand command = new SqliteCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@firstName", firstName);
                        command.Parameters.AddWithValue("@lastName", lastName);
                        command.Parameters.AddWithValue("@gender", gender);
                        command.Parameters.AddWithValue("@age", age);
                        command.Parameters.AddWithValue("@status", status);

                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                using (var connection = new SqliteConnection($"Data Source={pathData}/{dbName}.db"))
                {
                    connection.Open();

                    string sqlExpression = "SELECT COUNT(*) FROM Users WHERE gender='м'";
                    SqliteCommand command = new SqliteCommand(sqlExpression, connection);
                    object countMale = command.ExecuteScalar();

                    command.CommandText = "SELECT COUNT(*) FROM Users WHERE gender='ж'";
                    object countFemale = command.ExecuteScalar();

                    command.CommandText = "SELECT COUNT(*) FROM Users WHERE (gender='м' AND age>30 AND age<40)";
                    object countAgeMale = command.ExecuteScalar();


                    command.CommandText = "SELECT COUNT(*) FROM Users WHERE status='стандарт'";
                    object countStandart = command.ExecuteScalar();

                    command.CommandText = "SELECT COUNT(*) FROM Users WHERE status='премиум'";
                    object countPremium = command.ExecuteScalar();

                    command.CommandText = "SELECT COUNT(*) FROM Users WHERE (gender='ж' AND status='премиум' AND age<30)";
                    object femalePremiumAge = command.ExecuteScalar();

                    textBox1.AppendText("Отчет:\r\n");
                    textBox1.AppendText($"Мужчин: {countMale},\r\nЖенщин: {countFemale}\r\n");
                    textBox1.AppendText($"Мужчин в возрасте 30-40 лет: {countAgeMale}\r\n");
                    textBox1.AppendText($"Cтандартных: {countStandart}\r\nПремиум-аккаунтов: {countPremium}\r\n");
                    textBox1.AppendText($"Женщин с премиум-аккаунтом в возрасте до 30 лет: {femalePremiumAge}.\r\n");


                    Aspose.Words.Document doc = new Aspose.Words.Document();
                    DocumentBuilder builder = new DocumentBuilder(doc);
                    builder.Writeln("Отчет:");
                    builder.Writeln($"Мужчин: {countMale},\r\nЖенщин: {countFemale},");
                    builder.Writeln($"Мужчин в возрасте 30-40 лет: {countAgeMale},");
                    builder.Writeln($"Cтандартных: {countStandart},\r\nПремиум-аккаунтов: {countPremium},");
                    builder.Writeln($"Женщин с премиум-аккаунтом в возрасте до 30 лет: {femalePremiumAge}.");

                    doc.Save($"{pathReport}/{reportSQL}.docx");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }

        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            try
            {
                XPathNavigator nav;
                XPathDocument docNav;

                // Open the XML.
                docNav = new XPathDocument($"{pathData}/{xmlName}.xml");

                // Create a navigator to query with XPath.
                nav = docNav.CreateNavigator();

                XPathNodeIterator male = nav.Select("/DocumentElement/Лист1/Пол[../Пол='м']");
                XPathNodeIterator female = nav.Select("/DocumentElement/Лист1/Пол[../Пол='ж']");

                // Получаем количество записей
                int countMale = male.Count;
                int countFemale = female.Count;

                // Выводим количество записей на консоль
                textBox1.AppendText("Отчет:\r\n");
                textBox1.AppendText($"Количество мужчин: {countMale}\r\n");
                textBox1.AppendText($"Количество женщин: {countFemale}\r\n");

                XPathNodeIterator AgeMale = nav.Select("/DocumentElement/Лист1/Пол[../Пол='м' and ../Возраст>30 and ../Возраст<40]");

                int countAgeMale = AgeMale.Count;

                textBox1.AppendText($"Мужчин в возрасте 30-40 лет: {countAgeMale}\r\n");

                XPathNodeIterator standart = nav.Select("/DocumentElement/Лист1/Статус_x0020_аккаунта[../Статус_x0020_аккаунта='стандарт']");
                XPathNodeIterator premium = nav.Select("/DocumentElement/Лист1/Статус_x0020_аккаунта[../Статус_x0020_аккаунта='премиум']");

                int countStandart = standart.Count;
                int countPremium = premium.Count;

                textBox1.AppendText($"Cтандартных: {countStandart}\r\nПремиум-аккаунтов: {countPremium}\r\n");

                XPathNodeIterator femalePremiumAge = nav.Select("/DocumentElement/Лист1/Пол[../Пол='ж' and ../Статус_x0020_аккаунта='премиум' and ../Возраст<30]");

                int countFemalePremiumAge = femalePremiumAge.Count;

                textBox1.AppendText($"Женщин с премиум-аккаунтом в возрасте до 30 лет: {countFemalePremiumAge}.\r\n");


                Aspose.Words.Document doc = new Aspose.Words.Document();
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.Writeln("Отчет:");
                builder.Writeln($"Мужчин: {countMale},\r\nЖенщин: {countFemale},");
                builder.Writeln($"Мужчин в возрасте 30-40 лет: {countAgeMale},");
                builder.Writeln($"Cтандартных: {countStandart},\r\nПремиум-аккаунтов: {countPremium},");
                builder.Writeln($"Женщин с премиум-аккаунтом в возрасте до 30 лет: {countFemalePremiumAge}.");

                doc.Save($"{pathReport}/{reportXML}.docx");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }
    }
}