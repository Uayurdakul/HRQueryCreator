using ExcelDataReader;
using System.Data;
using System.Data.OleDb;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellQueryDemo
{
    /*
     * Dosya yolu open file dialog ile yap�labilir.
     * S�tun isimleri parametre olabilir
     * kaydederken txt dosyas� yerine xlsx dosyas�na kaydetme i�lemi yapabilir.
     * Filter de�eri default olarak and ile gelmek zorunda de�il kontrol� yap�yor.
     * excell sheet isimleri de�i�ken olaiblir. 
     * 
     */

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //public DataTable mergedData;

        public string bFilePath = "";
        public string aFilePath = "";


        private void button1_Click(object sender, EventArgs e)
        {

            string oldExcellFilePath = aFilePath;
            string excelFilePath = bFilePath;

            // SQL sorgular�n� tutacak bir StringBuilder nesnesi olu�turur
            StringBuilder sb = new StringBuilder();
            // A dosyas�n� okumak i�in bir Excel ba�lant�s� olu�turur
            string connStringA = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFilePath + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES\"";
            using (OleDbConnection connA = new OleDbConnection(connStringA))
            {
                // A dosyas�ndaki ilk sayfay� se�mek i�in bir SQL sorgusu olu�turur
                string queryA = "SELECT * FROM [SHEET$]";
                // Ba�lant�y� a�ar
                connA.Open();
                // Sorguyu �al��t�rmak i�in bir OleDbCommand nesnesi olu�turur
                using (OleDbCommand cmdA = new OleDbCommand(queryA, connA))
                {
                    // Sorgunun sonu�lar�n� okumak i�in bir OleDbDataReader nesnesi olu�turur
                    using (OleDbDataReader readerA = cmdA.ExecuteReader())
                    {
                        // A dosyas�ndaki her sat�r i�in
                        while (readerA.Read())
                        {
                            // Sat�rdaki kolon, tablo ve filtre de�erlerini al�r
                            string kolon = readerA["KOLON"].ToString();
                            string tablo = readerA["TABLE"].ToString();
                            string filtre = readerA["FILTRE"].ToString();
                            // B dosyas�n� okumak i�in bir Excel ba�lant�s� olu�turur
                            string connStringB = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + oldExcellFilePath + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES\"";
                            using (OleDbConnection connB = new OleDbConnection(connStringB))
                            {
                                // B dosyas�ndaki ilk sayfay� se�mek i�in bir SQL sorgusu olu�turur
                                string queryB = "SELECT * FROM [SHEET$]";
                                // Ba�lant�y� a�ar
                                connB.Open();
                                // Sorguyu �al��t�rmak i�in bir OleDbCommand nesnesi olu�turur
                                using (OleDbCommand cmdB = new OleDbCommand(queryB, connB))
                                {
                                    // Sorgunun sonu�lar�n� okumak i�in bir OleDbDataReader nesnesi olu�turur
                                    using (OleDbDataReader readerB = cmdB.ExecuteReader())
                                    {
                                        // b dekl her satir icin yapilacak islem
                                        while (readerB.Read())
                                        {
                                            // Sat�rdaki oldid ve newid de�erlerini ald�k
                                            string oldid = readerB["OLDID"].ToString();
                                            string newid = readerB["NEWID"].ToString();
                                            // Filtre de�eri bo�sa, UPDATE sorgusunu olu�turduk
                                            if (string.IsNullOrEmpty(filtre))
                                            {
                                                sb.AppendLine("UPDATE " + tablo + " SET " + kolon + "='" + newid + "' WHERE " + kolon + "='" + oldid + "'");
                                            }
                                            // Filtre de�eri varsa, UPDATE sorgusuna filtre ko�ulunu ekledik
                                            else
                                            {
                                                sb.AppendLine("UPDATE " + tablo + " SET " + kolon + "='" + newid + "' WHERE " + kolon + "='" + oldid + "' AND " + filtre);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            // SQL sorgular�n� metin dosyas�na kaydetmek i�in bir SaveFileDialog nesnesi olu�turdum
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            // sadece txt dosyasi
            saveFileDialog.Filter = "Text File|*.txt";
            // kaydetme tamam ise
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Olu�turulan SQL sorgular�n� se�ilen dosyaya yazar
                File.WriteAllText(saveFileDialog.FileName, sb.ToString());
                // ��lemin tamamland���n� bildirir
                MessageBox.Show("SQL sorgular� ba�ar�yla kaydedildi.");
            }
        }

        
        private void button3_Click(object sender, EventArgs e)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            

            OpenFileDialog openFileDialogB = new OpenFileDialog();
            openFileDialogB.Filter = "Excel Dosyalar�|*.xlsx;*.xls";
            if (openFileDialogB.ShowDialog() == DialogResult.OK)
            {
                bFilePath = openFileDialogB.FileName;

                // B dosyas�n� se�
                OpenFileDialog openFileDialogA = new OpenFileDialog();
                openFileDialogA.Filter = "Excel Dosyalar�|*.xlsx;*.xls";
                if (openFileDialogA.ShowDialog() == DialogResult.OK)
                {
                    aFilePath = openFileDialogA.FileName;
                }
            }

            // B.xlsx dosyas�n�n dizini
            string bExcelPath = bFilePath;// A.xlsx dosyas�n�n dizini
            string aExcelPath = aFilePath;

            // B.xlsx dosyas�ndaki t�m sat�rlar� Listbox1'e ekle
            Listbox1Doldur(bExcelPath);

            // A.xlsx dosyas�ndaki t�m sat�rlar� Listbox2'e ekle
            Listbox2Doldur(aExcelPath);

            // B.xlsx dosyas�ndaki renkli sat�rlar� Listbox3'e ekle
            // Listbox3DoldurRenkliSatirlar();
        }

        private void Listbox1Doldur(string excelPath)
        {
            // Excel dosyas�n� okuduk
            DataSet dataSet = ReadExcelFile(excelPath);

            // DataTable se�tik
            DataTable dataTable = dataSet.Tables[0];

            // DataTable'daki her sat�r� ListBox1'e ekledik
            foreach (DataRow row in dataTable.Rows)
            {
                listBox1.Items.Add(string.Join(", ", row.ItemArray));
            }
        }

        private void Listbox2Doldur(string excelPath)
        {
            // Excel dosyas�n� okuduk
            DataSet dataSet = ReadExcelFile(excelPath);

            // DataTable se�tik
            DataTable dataTable = dataSet.Tables[0];

            // DataTable'daki her sat�r� ListBox2'e ekledik
            foreach (DataRow row in dataTable.Rows)
            {
                listBox2.Items.Add(string.Join(", ", row.ItemArray));
            }
        }

        //private void Listbox3DoldurRenkliSatirlar()
        //{
        //    // Excel dosyas�n� a�
        //    Excel.Application excelApp = new Excel.Application();
        //    excelApp.Visible = false; // Excel uygulamas�n� gizle
        //    Excel.Workbook workbook = excelApp.Workbooks.Open("C:\\Users\\yunus\\Downloads\\ExcellToQuery"); // Dosya yolunu de�i�tirin
        //    Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[0]; // �lk �al��ma sayfas�n� se�

        //    // �al��ma sayfas�ndaki t�m h�creleri al
        //    Excel.Range range = worksheet.UsedRange;

        //    // H�creleri tek tek dola�
        //    for (int row = 1; row <= range.Rows.Count; row++)
        //    {
        //        for (int col = 1; col <= range.Columns.Count; col++)
        //        {
        //            // H�crenin arkaplan rengini al
        //            Excel.Range cell = (Excel.Range)range.Cells[row, col];
        //            int color = (int)cell.Interior.Color;

        //            // Arkaplan rengi mavi ise listbox'a ekle
        //            if (color == 16711680) // Mavi rengin RGB de�eri
        //            {
        //                string value = cell.Value2.ToString(); // H�crenin de�erini al
        //                listBox3.Items.Add(value); // Listbox'a ekle
        //            }
        //        }
        //    }

        //    // Excel dosyas�n� kapat
        //    workbook.Close(false);
        //    excelApp.Quit();
        //}

        private DataSet ReadExcelFile(string path)
        {
            using (FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                {
                    return reader.AsDataSet();
                }
            }
        }

        //private DataTable ReadExcelData(string filePath)
        //{
        //    DataTable dataTable = new DataTable();

        //    try
        //    {
        //        string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties=\"Excel 12.0 Xml;HDR=YES;\"";
        //        using (OleDbConnection connection = new OleDbConnection(connectionString))
        //        {
        //            connection.Open();
        //            OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM [SHEET$]", connection);
        //            adapter.Fill(dataTable);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show($"Excel dosyas� okuma hatas�: {ex.Message}");
        //    }

        //    return dataTable;
        //}

        //private DataTable MatchColumns(DataTable oldIDData, DataTable newIDData)
        //{
        //    DataTable mergedData = new DataTable();
        //    mergedData.Columns.Add("KOLON");
        //    mergedData.Columns.Add("TABLE");
        //    mergedData.Columns.Add("FILTRE");
        //    mergedData.Columns.Add("OLDID");
        //    mergedData.Columns.Add("NEWID");

        //    try
        //    {
        //        foreach (DataRow oldRow in oldIDData.Rows)
        //        {

        //            string column = oldRow["KOLON"].ToString();
        //            string table = oldRow["TABLE"].ToString();
        //            string filter = oldRow["F�LTRE"].ToString();
        //            string oldID = oldRow["OLDID"].ToString();



        //            DataRow newRow = mergedData.NewRow();

        //            DataRow matchingNewIDRow = newIDData.AsEnumerable()
        //                .FirstOrDefault(newRow => newRow["KOLON"].ToString() == column);

        //            if (matchingNewIDRow != null)
        //            {
        //                string newID = matchingNewIDRow["NEWID"].ToString();

        //                newRow["KOLON"] = column;
        //                newRow["TABLE"] = table;
        //                newRow["FILTRE"] = filter;
        //                newRow["OLDID"] = oldID;
        //                newRow["NEWID"] = newID;

        //                mergedData.Rows.Add(newRow);
        //            }

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show($"Kolonlar� e�le�tirme hatas�: {ex.Message}");
        //    }

        //    return mergedData;
        //}

        //private void SaveToExcel(DataTable data, string filePath)
        //{
        //    try
        //    {
        //        string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Extended Properties=\"Excel 12.0 Xml;HDR=YES;\"";
        //        using (OleDbConnection connection = new OleDbConnection(connectionString))
        //        {
        //            connection.Open();
        //            OleDbCommand cmd = new OleDbCommand();

        //            StringBuilder createTableQuery = new StringBuilder();
        //            createTableQuery.Append("CREATE TABLE [SHEET] (");

        //            foreach (DataColumn column in data.Columns)
        //            {
        //                createTableQuery.Append($"[{column.ColumnName}] TEXT,");
        //            }

        //            createTableQuery.Remove(createTableQuery.Length - 1, 1);
        //            createTableQuery.Append(")");
        //            cmd.Connection = connection;
        //            cmd.CommandText = createTableQuery.ToString();
        //            cmd.ExecuteNonQuery();

        //            foreach (DataRow row in data.Rows)
        //            {
        //                StringBuilder insertQuery = new StringBuilder();
        //                insertQuery.Append("INSERT INTO [SHEET$] (");

        //                foreach (DataColumn column in data.Columns)
        //                {
        //                    insertQuery.Append($"[{column.ColumnName}],");
        //                }

        //                insertQuery.Remove(insertQuery.Length - 1, 1);
        //                insertQuery.Append(") VALUES (");

        //                foreach (DataColumn column in data.Columns)
        //                {
        //                    insertQuery.Append($"'{row[column.ColumnName]}',");
        //                }

        //                insertQuery.Remove(insertQuery.Length - 1, 1);
        //                insertQuery.Append(")");

        //                cmd.CommandText = insertQuery.ToString();
        //                cmd.ExecuteNonQuery();

        //            }               
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show($"Excel dosyas�na yazma hatas�: {ex.Message}");
        //    }
        //}

        //private StringBuilder GenerateSQLQueries(DataTable excelData)
        //{
        //    StringBuilder queries = new StringBuilder();

        //    foreach (DataRow row in excelData.Rows)
        //    {
        //        string column = row["KOLON"].ToString();
        //        string table = row["TABLE"].ToString();
        //        string filter = row["FILTRE"].ToString();
        //        string oldID = row["OLDID"].ToString();
        //        string newID = row["NEWID"].ToString();

        //        string query = $"UPDATE {table} SET {column}='{newID}' WHERE {column}='{oldID}'{filter};";
        //        queries.AppendLine(query);
        //    }

        //    return queries;
        //}

        //private void WriteQueriesToFile(string queries, string filePath)
        //{
        //    try
        //    {
        //        File.WriteAllText(filePath, queries);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show($"Dosyaya yazma hatas�: {ex.Message}");
        //    }
        //}
    }
}