using ExcelDataReader;
using System.Data;
using System.Data.OleDb;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

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

        public string bFilePath = "";
        public string aFilePath = "";

        public List<string> standartTablo = new List<string>
        {
            "dbo.GRIDLAYOUTSETTINGS",
            "dbo.FLOWREQUESTS",
            "dbo.FLOWREQUESTS",
            "FSFILEHISTORY",
            "dbo.DOCUMENTS",
            "dbo.DOCUMENTS",
            "dbo.FSFILEDATASET",
            "LIVEFLOWS",
            "LIVEFLOWS",
            "dbo.CHECKOUTS",
            "dbo.PROJECTVERSIONS",
            "dbo.FSFILES",
            "dbo.FSFILES",
            "dbo.OSUSERS",
            "dbo.OSPROPERTYVALUES",
            "dbo.DOCUMENTNOTES",
            "dbo.OSPOSITIONS",
            "dbo.FSFOLDERS",
            "dbo.FSFOLDERS",
            "dbo.DOCUMENTHISTORY",
            "dbo.DOCUMENTHISTORY",
            "dbo.FSFOLDERSECURITY",
            "dbo.OSMANAGERS",
            "dbo.OSMANAGERS",
            "dbo.OSGROUPCONTENT",
            "dbo.PROJECTS",
            "dbo.MESSAGES",
            "dbo.MESSAGES",
            "dbo.MESSAGEREQUESTS",
            "dbo.MESSAGEREQUESTS",
            "dbo.FAVORITES",
            "dbo.MOBILEDEVICES",
            "dbo.FSFOLDERHISTORY",
            "dbo.DASHBOARDPROFILES",
            "dbo.USERLOGINATTEMPTS",
            "dbo.PASSWORDMANAGEMENT",
            "dbo.USERPASSWORDHISTORY",
            "dbo.OSDEPARTMENTS",
            "dbo.LASTVISITEDDOCUMENTS",
            "dbo.ARCHIVEFILTERPROFILES"
        };
        public List<string> ebaTablo = new List<string>
        {
            "dbo.E_prj_Kalite_TestTalep_Form",
            "dbo.E_prj_Kalite_TestTalep_Form",
            "dbo.E_BT_YardimMasasi_Form",
            "dbo.E_eBA_AracTalep_Form",
            "dbo.[Yetki Talep MAS+ Yetkileri]",
            "dbo.E_eBA_IseAlimTalepFormu_Form",
            "dbo.[Yetki Talep Netsis Yetkileri]",
            "dbo.E_eBA_ECMFormu_Form",
            "dbo.E_eBA_ECMFormu_MDLURUN",
            "dbo.E_eBA_ECMIFormu_Form",
            "dbo.[Yetki Talep Ortak Alan Yetkileri]",
            "dbo.E_eBA_ECRFormu_MDLURUN",
            "dbo.E_eBA_EksikKart_Form",
            "dbo.E_eBA_ECMFormu_MDLAKSIYONPLANI",
            "dbo.E_eBA_ECMFormu_MDLKALITE",
            "dbo.E_eBA_ECMFormu_MDLPROSES",
            "dbo.E_eBA_ECRFormu_MDLFINANS",
            "dbo.E_prj_UrunAgaciFormu_Form",
            "dbo.E_eBA_ECMFormu_MDLPROSES2",
            "dbo.E_eBA_ECRFormu_MDLPROSES",
            "dbo.E_eBA_ECMFormu_MDLFINANS",
            "dbo.E_eBA_ECRFormu_MDLKALITE",
            "dbo.E_prj_ISG_SMATPlanlama_MDLBEYAZYAKA",
            "dbo.E_eBA_ECMFormu_MDLLOJISTIK",
            "dbo.E_eBA_ECMFormu_MDLMALZEME",
            "dbo.E_eBA_ECMFormu_MDLPLANLAMA",
            "dbo.E_eBA_ECRFormu_MDLMALZEME",
            "dbo.E_eBA_ECRFormu_MDLPROSES2",
            "dbo.[Yetki Talep Fiyat Menü Yetkileri]",
            "dbo.E_eBA_ECMFormu_MDLSTANDART",
            "dbo.E_eBA_ECRFormu_MDLLOJISTIK",
            "dbo.E_prj_ISG_SMATDenetim_Form",
            "dbo.E_eBA_IcValidasyonSureci_MDLYORUM",
            "dbo.E_eBA_ECMFormu_MDLFONKSIYON",
            "dbo.E_eBA_ECMFormu_MDLTEDARIKCI",
            "dbo.E_eBA_ECRFormu_MDLPLANLAMA",
            "dbo.E_eBA_ECRFormu_MDLSTANDART",
            "dbo.E_eBA_LPAPlanlama_MDLTUMCALISANLAR",
            "dbo.E_prj_IK_ZiyaretciTalepFormu_Form",
            "dbo.E_eBA_TaseronFaaliyetYonetimi_Form",
            "dbo.E_eBA_HarcamaDetay_Form_tblDigerEkipUyeleri",
            "dbo.E_eBA_ECRFormu_MDLFONKSIYON",
            "dbo.E_eBA_ECRFormu_MDLTEDARIKCI",
            "dbo.E_eBA_IseGirisBildirimFormu_Form",
            "dbo.E_prj_ISG_TaseronFaaliyetYonetimi_Form",
            "dbo.E_prj_ISG_SMATPlanlama_MDLTUMCALISANLAR",
            "dbo.E_eBA_ECMFormu_MDLMODIFIKASYON",
            "dbo.E_prj_Kalite_TestTalep_Form_tblBilgilendirilecekler",
            "dbo.E_eBA_ECRFormu_MDLMODIFIKASYON",
            "dbo.E_eBA_ECMFormu_MDLCEVRESARTLARI",
            "dbo.E_eBA_ECRFormu_MDLCEVRESARTLARI",
            "dbo.E_eBA_DigerSurecler_ValidasyonSureci_MDLYORUM",
            "dbo.E_eBA_ECMFormu_MDLURETILEBILIRLIK",
            "dbo.E_eBA_ECMFormu_MDLURETIMLOKASYONU",
            "dbo.E_eBA_ECRFormu_MDLURETILEBILIRLIK",
            "dbo.E_eBA_ECRFormu_MDLURETIMLOKASYONU",
            "dbo.E_eBA_TaseronFaaliyetYonetimi_Form_tblBilgilendirilecekler",
            "dbo.E_prj_Kalite_TestTalep_Form_tblArgeLabBilgilendirilecekler",
            "dbo.E_eBA_SapmaSureci_MDLKAPANMAAKSIYONU",
            "dbo.E_eBA_ECMFormu_MDLOGRENILMISDERSLER",
            "dbo.E_eBA_ECRFormu_MDLOGRENILMISDERSLER",
            "dbo.E_prj_ISG_TaseronFaaliyetYonetimi_Form_tblBilgilendirilecekler",
            "dbo.E_prj_IT_ErisimYetkilendirmeveTalep_v2_Form",
            "dbo.E_eBA_DigerSurecler_ValidasyonSureci_MDLYORUMKALITE",
            "dbo.E_eBA_LPADenetimi_ISGUygunsuzlukFormModal",
            "dbo.E_eBA_LPADenetimi_IsTalimatiFormModal",
            "dbo.E_eBA_LPADenetimi_ssssDenetimUygunsuzlukFormModal",
            "dbo.E_eBA_LPADenetimi_sssssDenetimUygunsuzlukFormModal",
            "dbo.E_eBA_SapmaSureci_MDLRISKENGELLEMEAKSIYONU",
            "dbo.E_eBA_LPADenetimi_sDenetimUygunsuzlukFormModal",
            "dbo.E_eBA_LPADenetimi_sssDenetimUygunsuzlukFormModal",
            "dbo.E_eBA_LPADenetimi_CevreDenetimUygunsuzlukFormModal",
            "dbo.E_eBA_LPADenetimi_KaliteDenetimUygunsuzlukFormModal",
            "dbo.E_eBA_LPAAksiyonTakibi_Form",
            "dbo.E_eBA_LPAAksiyonTakibi_Form",
            "dbo.E_eBA_LPADenetimi_Form",
            "dbo.E_eBA_LPADenetimi_Form",
            "dbo.E_eba_ErisimYetkilendirmeVeTalep_Form",
            "dbo.E_eBA_IcValidasyonSureci_Form",
            "dbo.E_eBA_IcValidasyonSureci_Form",
            "dbo.E_prj_OperasyonTalimati_Form",
            "dbo.E_prj_OperasyonTalimati_Form",
            "dbo.E_prj_OperasyonTalimati_Form",
            "dbo.E_eBA_LPADenetimi_ssDenetimUygunsuzlukFormModal",
            "dbo.E_prj_ProjeFazGecisSureci_Form",
            "dbo.E_prj_ProjeFazGecisSureci_Form",
            "dbo.E_prj_ProjeFazGecisSureci_Form",
            "dbo.E_eBA_DigerSurecler_ValidasyonSureci_Form",
            "dbo.E_eBA_DigerSurecler_ValidasyonSureci_Form",
            "dbo.E_eBA_DigerSurecler_ValidasyonSureci_Form",
            "dbo.E_eBA_DigerSurecler_ValidasyonSureci_Form",
            "dbo.E_eBA_DigerSurecler_ValidasyonSureci_Form",
            "dbo.E_eBA_DigerSurecler_ValidasyonSureci_Form",
            "dbo.E_eBA_DigerSurecler_ValidasyonSureci_Form",
            "dbo.E_eBA_ECRFormu_Form",
            "dbo.E_eBA_ECRFormu_Form",
            "dbo.E_eBA_ECRFormu_Form",
            "dbo.E_eBA_ECRFormu_Form",
            "dbo.E_eBA_ECRFormu_Form",
            "dbo.E_eBA_ECRFormu_Form",
            "dbo.E_eBA_ECRFormu_Form",
            "dbo.E_eBA_ECRFormu_Form"
        };



        private void button1_Click(object sender, EventArgs e)
        {

            string oldExcellFilePath = aFilePath;
            string excelFilePath = bFilePath;

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

            checkBox1.Enabled = true;
            checkBox2.Enabled = true;
            button1.Enabled = true;

            OpenFileDialog openFileDialogB = new OpenFileDialog();
            openFileDialogB.Filter = "Excel Dosyalarý|*.xlsx;*.xls";
            if (openFileDialogB.ShowDialog() == DialogResult.OK)
            {
                bFilePath = openFileDialogB.FileName;

                // B dosyasýný seç
                OpenFileDialog openFileDialogA = new OpenFileDialog();
                openFileDialogA.Filter = "Excel Dosyalarý|*.xlsx;*.xls";
                if (openFileDialogA.ShowDialog() == DialogResult.OK)
                {
                    aFilePath = openFileDialogA.FileName;
                }
            }

            // B.xlsx dosyasýnýn dizini
            string bExcelPath = bFilePath;// A.xlsx dosyasýnýn dizini
            string aExcelPath = aFilePath;

            // B.xlsx dosyasýndaki tüm satýrlarý Listbox1'e ekle
            Listbox1Doldur(bExcelPath);

            // A.xlsx dosyasýndaki tüm satýrlarý Listbox2'e ekle
            Listbox2Doldur(aExcelPath);

            // B.xlsx dosyasýndaki renkli satýrlarý Listbox3'e ekle
            // Listbox3DoldurRenkliSatirlar();
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

        //private void Listbox3DoldurRenkliSatirlar()
        //{
        //    // Excel dosyasýný aç
        //    Excel.Application excelApp = new Excel.Application();
        //    excelApp.Visible = false; // Excel uygulamasýný gizle
        //    Excel.Workbook workbook = excelApp.Workbooks.Open("C:\\Users\\yunus\\Downloads\\ExcellToQuery"); // Dosya yolunu deðiþtirin
        //    Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[0]; // Ýlk çalýþma sayfasýný seç

        //    // Çalýþma sayfasýndaki tüm hücreleri al
        //    Excel.Range range = worksheet.UsedRange;

        //    // Hücreleri tek tek dolaþ
        //    for (int row = 1; row <= range.Rows.Count; row++)
        //    {
        //        for (int col = 1; col <= range.Columns.Count; col++)
        //        {
        //            // Hücrenin arkaplan rengini al
        //            Excel.Range cell = (Excel.Range)range.Cells[row, col];
        //            int color = (int)cell.Interior.Color;

        //            // Arkaplan rengi mavi ise listbox'a ekle
        //            if (color == 16711680) // Mavi rengin RGB deðeri
        //            {
        //                string value = cell.Value2.ToString(); // Hücrenin deðerini al
        //                listBox3.Items.Add(value); // Listbox'a ekle
        //            }
        //        }
        //    }

        //    // Excel dosyasýný kapat
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

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                foreach (string item in standartTablo)
                {
                    listBox3.Items.Add(item);
                }
            }
            else
            {
                foreach (string item in standartTablo)
                {
                    listBox3.Items.Remove(item);
                }
            }

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                foreach (string item in ebaTablo)
                {
                    listBox3.Items.Add(item);
                }
            }
            else
            {
                foreach (string item in ebaTablo)
                {
                    listBox3.Items.Remove(item);
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            checkBox1.Enabled = false;
            checkBox2.Enabled = false;
            button1.Enabled = false;
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