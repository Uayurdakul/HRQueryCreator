using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using System.Data;
using System.Data.OleDb;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcellQueryDemo
{
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
            "dbo.[Yetki Talep Fiyat Men� Yetkileri]",
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
            List<string> sorgular = new List<string>();

            foreach (string itemLB3 in listBox3.Items)
            {
                string[] varlikTabloFiltre = itemLB3.Split(',');

                foreach (string itemLB2 in listBox2.Items)
                {
                    string[] oldIDnewID = itemLB2.Split(",");

                    sorgular.Add($"UPDATE {varlikTabloFiltre[1].Trim(' ')} SET {varlikTabloFiltre[0].Trim(' ')}='{oldIDnewID[1].Trim(' ')}' WHERE {varlikTabloFiltre[0].Trim(' ')}='{oldIDnewID[0].Trim(' ')}' {varlikTabloFiltre[2].Trim(' ')}");
                }
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Text File|*.txt";
            saveFileDialog.Title = "Sorgular� Kaydet";
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

                MessageBox.Show("Sorgular ba�ar�yla kaydedildi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            
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
            bExcelPath = bFilePath;
            // A.xlsx dosyas�n�n dizini
            aExcelPath = aFilePath;

            // B.xlsx dosyas�ndaki t�m sat�rlar� Listbox1'e ekle
            Listbox1Doldur(bExcelPath);

            // A.xlsx dosyas�ndaki t�m sat�rlar� Listbox2'e ekle
            Listbox2Doldur(aExcelPath);

            
        }

        private void Listbox1Doldur(string excelPath)
        {
            // Excel dosyas�n� okuduk
            DataSet dataSet = ReadExcelFile(excelPath);

            // DataTable se�tik
            System.Data.DataTable dataTable = dataSet.Tables[0];

            // DataTable'daki her sat�r� ListBox1'e ekledik
            foreach (DataRow row in dataTable.Rows)
            {
                listBox1.Items.Add(string.Join(", ", row.ItemArray));
            }
            listBox1.Items.RemoveAt(0);
        }

        private void Listbox2Doldur(string excelPath)
        {
            // Excel dosyas�n� okuduk
            DataSet dataSet = ReadExcelFile(excelPath);

            // DataTable se�tik
            System.Data.DataTable dataTable = dataSet.Tables[0];

            // DataTable'daki her sat�r� ListBox2'e ekledik
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
            // CheckBox se�iliyse listbox elemanlar�n� kontrol ettik
            List<string> secilenler = new List<string>();

            foreach (string item in listBox1.Items)
            {
                string[] kelimeler = item.Split(',');

                // tablolar e�le�iyor mu kontrol
                if (kelimeler.Length > 1 && tableName.Contains(kelimeler[1].Trim()))
                {
                    secilenler.Add(item);
                }
            }

            //Se�ilen elemanlar� yeni bir listbox'a ekledik   
            listBox3.Items.AddRange(secilenler.ToArray());
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked && checkBox2.Checked)
            {
                listBox3.Items.Clear();
                label3.Text = "Standart ve Eba Tablolar�";
                listBox3.Items.AddRange(listBox1.Items);
            }
            else if (checkBox1.Checked)
            {
                button1.Enabled = true;
                EslesenTablo(standartTablo);
                label3.Text = "Standart Tablolar";
            }
            else
            {
                if (checkBox2.Checked)
                {
                    listBox3.Items.Clear();
                    EslesenTablo(ebaTablo);
                    label3.Text = "Eba Tablolar�";
                }
                else
                {
                    listBox3.Items.Clear();
                    button1.Enabled = false;
                    label3.Text = "Tablolar";
                }
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked && checkBox2.Checked)
            {
                listBox3.Items.Clear();
                label3.Text = "Standart ve Eba Tablolar�";
                listBox3.Items.AddRange(listBox1.Items);
            }
            else if (checkBox2.Checked)
            {
                EslesenTablo(ebaTablo);
                label3.Text = "Eba Tablolar�";
                button1.Enabled = true;
            }
            else
            {
                if (checkBox1.Checked)
                {
                    listBox3.Items.Clear();
                    EslesenTablo(standartTablo);
                    label3.Text = "Standart Tablolar";
                }
                else
                {
                    listBox3.Items.Clear();
                    button1.Enabled = false;
                    label3.Text = "Tablolar";
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