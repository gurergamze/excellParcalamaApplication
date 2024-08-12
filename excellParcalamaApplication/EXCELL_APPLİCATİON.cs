using ClosedXML.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;



namespace excellParcalamaApplication
{
    public partial class EXCELL_APPLİCATİON : Form
    {

        OpenFileDialog openFileDialog1 = new OpenFileDialog()
        {
            Filter = "Excel Dosyası |*.xlsx| Excel Dosyası|*.xls",
            // open file dialog açıldığında sadece excel dosyalarınu görecek
            Title = "Excel Dosyası Seçiniz..",
            // open file dialog penceresinin başlığı

        };
        OpenFileDialog openFileDialog2 = new OpenFileDialog()
        {
            Filter = "Excel Dosyası |*.xlsx| Excel Dosyası|*.xls",
            Title = "Bayi VDM Bilgileri Dosyasını Seçiniz..",
        };


        public EXCELL_APPLİCATİON()
        {

            InitializeComponent();
        }

        private void SetDataTable_To_Excel(DataTable dtTable, string PathFileName)
        {
            string folderPath = @"C:\Users\ggurer\Desktop\ExcelDeneme";
            string filePath = Path.Combine(folderPath, PathFileName + ".xlsx");


            using (var workbook = new XLWorkbook())
            {
                workbook.Worksheets.Add(dtTable, "Deneme");
                workbook.SaveAs(filePath);
            }
        }



        private void button_parcala_Click(object sender, EventArgs e)
        {
            //FileInfo fileInfo;
            //string filepath2 = AppDomain.CurrentDomain.BaseDirectory.ToString() + @"ExcelExport\";
            string folderPath = @"C:\Users\ggurer\Desktop\ExcelDeneme";


            /* foreach (string dosya in Directory.GetFiles(filepath2))
             {
                 fileInfo = new FileInfo(dosya);
                 fileInfo.Delete();
            Türkçe
             }*/
            foreach (string dosya in Directory.GetFiles(folderPath))
            {
                FileInfo fileInfo = new FileInfo(dosya);
                fileInfo.Delete();
            }
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                MessageBox.Show("Parçalama İşlemi Başlatıldı. Lütfen Bekleyiniz..");
                string DosyaYolu = openFileDialog1.FileName;// dosya yolu
                string DosyaAdi = openFileDialog1.SafeFileName; // dosya adı

                OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + DosyaYolu + "; Extended Properties='Excel 12.0 xml;HDR=YES;'");
                OleDbCommand cmd = new OleDbCommand();
                con.Open();
                cmd.Connection = con;
                cmd.CommandText = "SELECT DISTINCT " + textBox_colomn_name.Text + " FROM [VDM_Kodu$]";
                OleDbDataReader dr = cmd.ExecuteReader();


                ArrayList VDM_kodlari = new ArrayList();

                while (dr.Read())
                {
                    VDM_kodlari.Add(dr[textBox_colomn_name.Text]);
                }
                OleDbDataAdapter da = new OleDbDataAdapter("Select*from[VDM_Kodu$]", con);
                DataTable DTexcel = new DataTable();
                da.Fill(DTexcel);

                dr.Close();
                con.Close();



                int sayac = 0;
                sayac = VDM_kodlari.Count;
                for (int i = 0; i < sayac; i++)
                {

                    if (VDM_kodlari[i].ToString().Length == 0)
                        continue;

                    var results = (from myRow in DTexcel.AsEnumerable()
                                   where myRow.Field<string>(textBox_colomn_name.Text) == VDM_kodlari[i].ToString()
                                   select myRow).CopyToDataTable();
                    SetDataTable_To_Excel(results, VDM_kodlari[i].ToString());


                }


                MessageBox.Show("Parçalama İşlemi Tamamlandı.");




                if (openFileDialog2.ShowDialog() == DialogResult.OK)
                {
                    string vdmFilePath = openFileDialog2.FileName;

                    if (!File.Exists(vdmFilePath))
                    {
                        MessageBox.Show("Dosya bulunamadı.");
                        return;
                    }

                    DataTable dtVdmInfo = new DataTable();

                    try
                    {
                        
                        using (OleDbConnection con2 = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + vdmFilePath + "; Extended Properties='Excel 12.0 xml;HDR=YES;'"))
                        {
                            con2.Open();
                            OleDbCommand cmd2 = new OleDbCommand("SELECT * FROM [Sayfa1$]", con); 
                            OleDbDataAdapter da2 = new OleDbDataAdapter(cmd);
                            da2.Fill(dtVdmInfo);
                            con2.Close();
                        }

                        MessageBox.Show("Dosya başarıyla okundu.");
                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show($"Dosya okuma hatası: {ex.Message}");
                        return;
                    }




                    // Her VDM için genel toplam dosyası oluşturma
                    foreach (string vdmKod in VDM_kodlari)
                    {
                        string vdmParcaFilePath = Path.Combine(folderPath, vdmKod + ".xlsx");

                        if (File.Exists(vdmParcaFilePath))
                        {
                            DataTable vdmParcaTable = new DataTable();

                            using (OleDbConnection con3 = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + vdmParcaFilePath + "; Extended Properties='Excel 12.0 xml;HDR=YES;'"))
                            {
                                con3.Open();
                                OleDbCommand cmd3 = new OleDbCommand("SELECT * FROM [Deneme$]", con3);
                                OleDbDataAdapter da3 = new OleDbDataAdapter(cmd3);
                                da3.Fill(vdmParcaTable);
                                con3.Close();
                            }


                            // Genel toplam dosyası oluşturma
                            var genelToplamResults = (from vdmRow in vdmParcaTable.AsEnumerable()
                                                       join vdmBilgiRow in dtVdmInfo.AsEnumerable()
                                                       on vdmRow.Field<string>(textBox_colomn_name.Text) equals vdmBilgiRow.Field<string>(textBox_colomn_name.Text)
                                                       into joinedData
                                                       from vdmBilgiRow in joinedData.DefaultIfEmpty()
                                                       select new
                                                       {
                                                           vdmRow,
                                                           vdmBilgiRow
                                                       }).ToList();

                            DataTable genelToplamTable = new DataTable();
                            foreach (DataColumn column in vdmParcaTable.Columns)
                            {
                                genelToplamTable.Columns.Add(column.ColumnName);
                            }

                            foreach (DataColumn column in dtVdmInfo.Columns)
                            {
                                if (!genelToplamTable.Columns.Contains(column.ColumnName))
                                    genelToplamTable.Columns.Add(column.ColumnName);
                            }

                            foreach (var result in genelToplamResults)
                            {
                                DataRow newRow = genelToplamTable.NewRow();

                                foreach (DataColumn column in vdmParcaTable.Columns)
                                {
                                    newRow[column.ColumnName] = result.vdmRow[column];
                                }

                                if (result.vdmBilgiRow != null)
                                {
                                    foreach (DataColumn column in dtVdmInfo.Columns)
                                    {
                                        newRow[column.ColumnName] = result.vdmBilgiRow[column];
                                    }
                                }

                                genelToplamTable.Rows.Add(newRow);
                            }

                            string genelToplamFileName = "Aylık Cihaz Ücret Detayı " + vdmKod;
                        SetDataTable_To_Excel(genelToplamTable, genelToplamFileName);
                    }
                }

                MessageBox.Show("Genel Toplam Dosyaları Oluşturuldu.");
            }


            // Bayi İletişim Excel Dosyasını Seçme
            OpenFileDialog openFileDialog3 = new OpenFileDialog()
                {
                    Filter = "Excel Dosyası |*.xlsx| Excel Dosyası|*.xls",
                    Title = "Bayi İletişim Excel Dosyasını Seçiniz..",
                };

                if (openFileDialog3.ShowDialog() == DialogResult.OK)
                {
                    string mailFilePath = openFileDialog3.FileName;
                    DataTable dtMailInfo = new DataTable();

                    string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + mailFilePath + ";Extended Properties='Excel 12.0 xml;HDR=YES;'";
                    using (OleDbConnection connection = new OleDbConnection(connectionString))
                    {
                        connection.Open();

                   
                        DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                        string firstSheetName = schemaTable.Rows[0]["TABLE_NAME"].ToString();

                        
                        string query = $"SELECT * FROM [{firstSheetName}]";
                        using (OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection))
                        {
                            adapter.Fill(dtMailInfo);
                        }

                        connection.Close();
                    }

                    string emailBody = mailBody.Text;

                    foreach (DataRow row in dtMailInfo.Rows)
                    {
                        string VDMKodu = row["VDM_Kodu"].ToString();
                        string email = row["Email"].ToString();
                        // string attachmentPath = AppDomain.CurrentDomain.BaseDirectory.ToString() + @"ExcelExport\" + VDMKodu + ".xlsx";
                        string attachmentPath = Path.Combine(folderPath, VDMKodu + ".xlsx");
                        string attachmentPath2 = Path.Combine(folderPath, "Aylık Cihaz Ücret Detayı " + VDMKodu + ".xlsx");
                        if (File.Exists(attachmentPath) && File.Exists(attachmentPath2))
                        {
                            SendEmail(email, emailBody, attachmentPath, attachmentPath2);
                        }
                        else
                        {
                            MessageBox.Show($"Dosya eksik: {VDMKodu} için dosya(lar) bulunamadı.");
                        }
                    }

                    MessageBox.Show("Mailler Gönderildi.");
                }
            }
        






        }
        private void SendEmail(string toEmail, string body, string attachmentPath, string attachmentPath2)
        {
            string smtpPassword = ConfigurationManager.AppSettings["SmtpPassword"];

            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
            string fromEmail = fromMail.Text;

            mail.From = new MailAddress(fromEmail);
            mail.To.Add(toEmail);

            string currentMonthName = DateTime.Now.ToString("MMMM", new System.Globalization.CultureInfo("tr-TR"));
            string subject = $"Vodafone Silver ÖKC {currentMonthName} Ayı Mutal";

            mail.Subject = subject;
            mail.Body = body;
            Attachment attachment = new Attachment(attachmentPath);
            mail.Attachments.Add(attachment);
            Attachment attachment2 = new Attachment(attachmentPath2);
            mail.Attachments.Add(attachment2);

            SmtpServer.Port = 587;
            SmtpServer.Credentials = new NetworkCredential(fromEmail, smtpPassword);
            SmtpServer.EnableSsl = true;

            try
            {
                SmtpServer.Send(mail);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Mail gönderilemedi: " + ex.Message);
            }
        }

        private void button_exit_Click(object sender, EventArgs e)
        {
            this.Close();
            Application.Exit();
        }

        private void EXCELL_APPLİCATİON_Load(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}


