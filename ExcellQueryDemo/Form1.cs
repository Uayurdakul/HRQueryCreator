using ExcelDataReader;
using System.Data;
using System.Data.OleDb;
using System.Text;
using System.Windows.Forms;

namespace ExcellQueryDemo
{
    /*
     * Dosya yolu open file dialog ile yapýlabilir.
     * Sütun isimleri parametre olabilir
     * kaydederken txt dosyasý yerine xlsx dosyasýna kaydetme iþlemi yapabilir.
     * Filter deðeri default olarak and ile gelmek zorunda deðil kontrolü yapýyor.
     * excell sheet isimleri deðiþken olaiblir. 
     * 
     */

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //public DataTable mergedData;

        private void button1_Click(object sender, EventArgs e)
        {

            string oldExcellFilePath = "C:\\Users\\yunus\\Downloads\\ExcellToQueryOldID";
            string excelFilePath = "C:\\Users\\yunus\\Downloads\\ExcellToQuery";

            // SQL sorgularýný tutacak bir StringBuilder nesnesi oluþturur
            StringBuilder sb = new StringBuilder();
            // A dosyasýný okumak için bir Excel baðlantýsý oluþturur
            string connStringA = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFilePath + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES\"";
            using (OleDbConnection connA = new OleDbConnection(connStringA))
            {
                // A dosyasýndaki ilk sayfayý seçmek için bir SQL sorgusu oluþturur
                string queryA = "SELECT * FROM [SHEET$]";
                // Baðlantýyý açar
                connA.Open();
                // Sorguyu çalýþtýrmak için bir OleDbCommand nesnesi oluþturur
                using (OleDbCommand cmdA = new OleDbCommand(queryA, connA))
                {
                    // Sorgunun sonuçlarýný okumak için bir OleDbDataReader nesnesi oluþturur
                    using (OleDbDataReader readerA = cmdA.ExecuteReader())
                    {
                        // A dosyasýndaki her satýr için
                        while (readerA.Read())
                        {
                            // Satýrdaki kolon, tablo ve filtre deðerlerini alýr
                            string kolon = readerA["KOLON"].ToString();
                            string tablo = readerA["TABLE"].ToString();
                            string filtre = readerA["FILTRE"].ToString();
                            // B dosyasýný okumak için bir Excel baðlantýsý oluþturur
                            string connStringB = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + oldExcellFilePath + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES\"";
                            using (OleDbConnection connB = new OleDbConnection(connStringB))
                            {
                                // B dosyasýndaki ilk sayfayý seçmek için bir SQL sorgusu oluþturur
                                string queryB = "SELECT * FROM [SHEET$]";
                                // Baðlantýyý açar
                                connB.Open();
                                // Sorguyu çalýþtýrmak için bir OleDbCommand nesnesi oluþturur
                                using (OleDbCommand cmdB = new OleDbCommand(queryB, connB))
                                {
                                    // Sorgunun sonuçlarýný okumak için bir OleDbDataReader nesnesi oluþturur
                                    using (OleDbDataReader readerB = cmdB.ExecuteReader())
                                    {
                                        // b dekl her satir icin yapilacak islem
                                        while (readerB.Read())
                                        {
                                            // Satýrdaki oldid ve newid deðerlerini aldýk
                                            string oldid = readerB["OLDID"].ToString();
                                            string newid = readerB["NEWID"].ToString();
                                            // Filtre deðeri boþsa, UPDATE sorgusunu oluþturduk
                                            if (string.IsNullOrEmpty(filtre))
                                            {
                                                sb.AppendLine("UPDATE " + tablo + " SET " + kolon + "='" + newid + "' WHERE " + kolon + "='" + oldid + "'");
                                            }
                                            // Filtre deðeri varsa, UPDATE sorgusuna filtre koþulunu ekledik
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
            // SQL sorgularýný metin dosyasýna kaydetmek için bir SaveFileDialog nesnesi oluþturdum
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            // sadece txt dosyasi
            saveFileDialog.Filter = "Text File|*.txt";
            // kaydetme tamam ise
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Oluþturulan SQL sorgularýný seçilen dosyaya yazar
                File.WriteAllText(saveFileDialog.FileName, sb.ToString());
                // Ýþlemin tamamlandýðýný bildirir
                MessageBox.Show("SQL sorgularý baþarýyla kaydedildi.");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // B.xlsx dosyasýnýn dizini
            string bExcelPath = "C:\\Users\\yunus\\Downloads\\ExcellToQuery.xlsx";

            // A.xlsx dosyasýnýn dizini
            string aExcelPath = "C:\\Users\\yunus\\Downloads\\ExcellToQueryOldID.xlsx";

            // B.xlsx dosyasýndaki tüm satýrlarý Listbox1'e ekle
            Listbox1Doldur(bExcelPath);

            // A.xlsx dosyasýndaki tüm satýrlarý Listbox2'e ekle
            Listbox2Doldur(aExcelPath);

            // B.xlsx dosyasýndaki renkli satýrlarý Listbox3'e ekle
            Listbox3DoldurRenkliSatirlar(bExcelPath);
        }

        private void Listbox1Doldur(string excelPath)
        {
            // Excel dosyasýný okuduk
            DataSet dataSet = ReadExcelFile(excelPath);

            // DataTable seçtik
            DataTable dataTable = dataSet.Tables[0];

            // DataTable'daki her satýrý ListBox1'e ekledik
            foreach (DataRow row in dataTable.Rows)
            {
                listBox1.Items.Add(string.Join(", ", row.ItemArray));
            }
        }

        private void Listbox2Doldur(string excelPath)
        {
            // Excel dosyasýný okuduk
            DataSet dataSet = ReadExcelFile(excelPath);

            // DataTable seçtik
            DataTable dataTable = dataSet.Tables[0];

            // DataTable'daki her satýrý ListBox2'e ekledik
            foreach (DataRow row in dataTable.Rows)
            {
                listBox2.Items.Add(string.Join(", ", row.ItemArray));
            }
        }

        private void Listbox3DoldurRenkliSatirlar(string excelPath)
        {
            // Excel dosyasýný okuduk
            DataSet dataSet = ReadExcelFile(excelPath);

            // DataTable seçtik
            DataTable dataTable = dataSet.Tables[0];

            // DataTable'daki her satýrý kontrol et ve renkli olanlarý ListBox3'e ekledik
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.DefaultCellStyle.BackColor != Color.Blue) // Renk kontrolü burda yapýlýyor beyaz dýþý alsýn dedim
                {
                    listBox3.Items.Add(string.Join(", ", row.Cells.Cast<DataGridViewCell>().Select(cell => cell.Value)));
                }
            }
        }

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
        //        MessageBox.Show($"Excel dosyasý okuma hatasý: {ex.Message}");
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
        //            string filter = oldRow["FÝLTRE"].ToString();
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
        //        MessageBox.Show($"Kolonlarý eþleþtirme hatasý: {ex.Message}");
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
        //        MessageBox.Show($"Excel dosyasýna yazma hatasý: {ex.Message}");
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
        //        MessageBox.Show($"Dosyaya yazma hatasý: {ex.Message}");
        //    }
        //}
    }
}