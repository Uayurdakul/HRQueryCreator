using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
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
            List <string> sorgular=new List<string>();

            foreach (string itemLB3 in listBox3.Items)
            {
                string[] varlikTabloFiltre = itemLB3.Split(',');

                foreach (string itemLB2 in listBox2.Items)
                {
                    string[] oldIDnewID= itemLB2.Split(",");

                    sorgular.Add($"UPDATE {varlikTabloFiltre[1]} SET {varlikTabloFiltre[0]}='{oldIDnewID[1]}' WHERE {varlikTabloFiltre[0]}='{oldIDnewID[0]}' {varlikTabloFiltre[2]}");
                }               
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Text File|*.txt";
            saveFileDialog.Title = "Sorgularý Kaydet";
            saveFileDialog.ShowDialog();

            if (saveFileDialog.FileName != "")
            {
                using (StreamWriter writer = new StreamWriter(saveFileDialog.OpenFile()))
                {
                    foreach (var query in sorgular)
                    {
                        writer.WriteLine(query);
                    }
                }

                MessageBox.Show("Sorgular baþarýyla kaydedildi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            //ASAGIDAKI KISIM DIREKT EXCELLDEN VERI CEKIP SORGU OLUSTURAN KISIMDIR!!!

            //string oldExcellFilePath = aFilePath;
            //string excelFilePath = bFilePath;

            //// SQL sorgularýný tutacak bir StringBuilder nesnesi oluþturur
            //StringBuilder sb = new StringBuilder();
            //// A dosyasýný okumak için bir Excel baðlantýsý oluþturur
            //string connStringA = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelFilePath + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES\"";
            //using (OleDbConnection connA = new OleDbConnection(connStringA))
            //{
            //    // A dosyasýndaki ilk sayfayý seçmek için bir SQL sorgusu oluþturur
            //    string queryA = "SELECT * FROM [SHEET$]";
            //    // Baðlantýyý açar
            //    connA.Open();
            //    // Sorguyu çalýþtýrmak için bir OleDbCommand nesnesi oluþturur
            //    using (OleDbCommand cmdA = new OleDbCommand(queryA, connA))
            //    {
            //        // Sorgunun sonuçlarýný okumak için bir OleDbDataReader nesnesi oluþturur
            //        using (OleDbDataReader readerA = cmdA.ExecuteReader())
            //        {
            //            // A dosyasýndaki her satýr için
            //            while (readerA.Read())
            //            {
            //                // Satýrdaki kolon, tablo ve filtre deðerlerini alýr
            //                string kolon = readerA["KOLON"].ToString();
            //                string tablo = readerA["TABLE"].ToString();
            //                string filtre = readerA["FILTRE"].ToString();
            //                // B dosyasýný okumak için bir Excel baðlantýsý oluþturur
            //                string connStringB = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + oldExcellFilePath + ";Extended Properties=\"Excel 12.0 Xml;HDR=YES\"";
            //                using (OleDbConnection connB = new OleDbConnection(connStringB))
            //                {
            //                    // B dosyasýndaki ilk sayfayý seçmek için bir SQL sorgusu oluþturur
            //                    string queryB = "SELECT * FROM [SHEET$]";
            //                    // Baðlantýyý açar
            //                    connB.Open();
            //                    // Sorguyu çalýþtýrmak için bir OleDbCommand nesnesi oluþturur
            //                    using (OleDbCommand cmdB = new OleDbCommand(queryB, connB))
            //                    {
            //                        // Sorgunun sonuçlarýný okumak için bir OleDbDataReader nesnesi oluþturur
            //                        using (OleDbDataReader readerB = cmdB.ExecuteReader())
            //                        {
            //                            // b dekl her satir icin yapilacak islem
            //                            while (readerB.Read())
            //                            {
            //                                // Satýrdaki oldid ve newid deðerlerini aldýk
            //                                string oldid = readerB["OLDID"].ToString();
            //                                string newid = readerB["NEWID"].ToString();
            //                                // Filtre deðeri boþsa, UPDATE sorgusunu oluþturduk
            //                                if (string.IsNullOrEmpty(filtre))
            //                                {
            //                                    sb.AppendLine("UPDATE " + tablo + " SET " + kolon + "='" + newid + "' WHERE " + kolon + "='" + oldid + "'");
            //                                }
            //                                // Filtre deðeri varsa, UPDATE sorgusuna filtre koþulunu ekledik
            //                                else
            //                                {
            //                                    sb.AppendLine("UPDATE " + tablo + " SET " + kolon + "='" + newid + "' WHERE " + kolon + "='" + oldid + "' AND " + filtre);
            //                                }
            //                            }
            //                        }
            //                    }
            //                }
            //            }
            //        }
            //    }
            //}

            //// SQL sorgularýný metin dosyasýna kaydetmek için bir SaveFileDialog nesnesi oluþturdum
            //SaveFileDialog saveFileDialog = new SaveFileDialog();
            //// sadece txt dosyasi
            //saveFileDialog.Filter = "Text File|*.txt";
            //// kaydetme tamam ise
            //if (saveFileDialog.ShowDialog() == DialogResult.OK)
            //{
            //    // Oluþturulan SQL sorgularýný seçilen dosyaya yazar
            //    // File.WriteAllText(saveFileDialog.FileName, sb.ToString());
            //    // Ýþlemin tamamlandýðýný bildirir
            //    MessageBox.Show("SQL sorgularý baþarýyla kaydedildi.");
            //}
        }

        string bExcelPath = "";
        string aExcelPath = "";

        private void button3_Click(object sender, EventArgs e)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            checkBox1.Enabled = true;
            checkBox2.Enabled = true;
            button1.Enabled = false;

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
            bExcelPath = bFilePath;// A.xlsx dosyasýnýn dizini
            aExcelPath = aFilePath;

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
            System.Data.DataTable dataTable = dataSet.Tables[0];

            // DataTable'daki her satýrý ListBox1'e ekledik
            foreach (DataRow row in dataTable.Rows)
            {
                listBox1.Items.Add(string.Join(", ", row.ItemArray));
            }
            listBox1.Items.RemoveAt(0);
        }

        private void Listbox2Doldur(string excelPath)
        {
            // Excel dosyasýný okuduk
            DataSet dataSet = ReadExcelFile(excelPath);

            // DataTable seçtik
            System.Data.DataTable dataTable = dataSet.Tables[0];

            // DataTable'daki her satýrý ListBox2'e ekledik
            foreach (DataRow row in dataTable.Rows)
            {
                listBox2.Items.Add(string.Join(", ", row.ItemArray));
            }
            listBox2.Items.RemoveAt(0);
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

        private void EslesenTablo(List<string> tableName)
        {
            // CheckBox seçiliyse listbox elemanlarýný kontrol ettik
            List<string> secilenler = new List<string>();

            foreach (string item in listBox1.Items)
            {
                string[] kelimeler = item.Split(',');

                // tablolar eþleþiyor mu kontrol
                if (kelimeler.Length > 1 && tableName.Contains(kelimeler[1].Trim()))
                {
                    secilenler.Add(item);
                }
            }

            //Seçilen elemanlarý yeni bir listbox'a ekledik   
              listBox3.Items.AddRange(secilenler.ToArray());         
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {       
            if (checkBox1.Checked&&checkBox2.Checked)
            {
                listBox3.Items.Clear();
                listBox3.Items.AddRange(listBox1.Items);
            }
            else if (checkBox1.Checked)
            {
                button1.Enabled = true;
                EslesenTablo(standartTablo);              
            }
            else
            {
                if (checkBox2.Checked)
                { 
                    EslesenTablo(ebaTablo);
                    listBox3.Items.Clear();          
                }
                else
                {
                    listBox3.Items.Clear();
                    button1.Enabled = false;
                }
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked && checkBox2.Checked)
            {
                listBox3.Items.Clear();
                listBox3.Items.AddRange(listBox1.Items);
            }
            else if (checkBox2.Checked)
            {
                EslesenTablo(ebaTablo);
                button1.Enabled = true;
            }            
            else
            {
                if (checkBox1.Checked)
                {
                    listBox3.Items.Clear();
                    EslesenTablo(standartTablo);
                }
                else
                {
                    listBox3.Items.Clear();
                    button1.Enabled = false;
                }
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            checkBox1.Enabled = false;
            checkBox2.Enabled = false;
            button1.Enabled = false;
        }

    
    }
}