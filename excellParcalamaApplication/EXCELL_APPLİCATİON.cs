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
                        using (var workbook = new XLWorkbook(vdmFilePath))
                        {
                            var sheetNames = workbook.Worksheets.Select(ws => ws.Name).ToList();
                            MessageBox.Show($"Mevcut sayfalar: {string.Join(", ", sheetNames)}");

                            // Geçerli sayfa adıyla sayfayı seçme
                            string sheetName = "Sayfa1"; 
                            if (sheetNames.Contains(sheetName))
                            {
                                var worksheet = workbook.Worksheet(sheetName);
                                var headerRow = worksheet.Row(1);
                                if (headerRow.Cells().Count() == 0)
                                {
                                    MessageBox.Show("Başlık satırı bulunamadı.");
                                    return;
                                }

                                foreach (var cell in headerRow.Cells())
                                {
                                    dtVdmInfo.Columns.Add(cell.GetValue<string>() ?? string.Empty);
                                }

                                for (int rowIndex = 2; rowIndex <= worksheet.LastRowUsed().RowNumber(); rowIndex++)
                                {
                                    var row = worksheet.Row(rowIndex);
                                    if (row.Cells().All(cell => string.IsNullOrEmpty(cell.GetValue<string>())))
                                    {
                                        continue;
                                    }

                                    var dataRow = dtVdmInfo.NewRow();
                                    for (int colIndex = 1; colIndex <= headerRow.Cells().Count(); colIndex++)
                                    {
                                        dataRow[colIndex - 1] = row.Cell(colIndex).GetValue<string>() ?? string.Empty;
                                    }

                                    dtVdmInfo.Rows.Add(dataRow);
                                }
                            }
                            else
                            {
                                MessageBox.Show($"Sayfa '{sheetName}' bulunamadı. Lütfen geçerli bir sayfa adı girin.");
                            }
                        }

                        MessageBox.Show("Dosya başarıyla okundu.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Dosya okuma hatası: {ex.Message}");
                    }



                    // Her VDM için genel toplam dosyası oluşturma
                     foreach (string vdmKod in VDM_kodlari)
                     {
                         string vdmParcaFilePath = Path.Combine(folderPath, vdmKod + ".xlsx");
                         if (File.Exists(vdmParcaFilePath))
                         {
                             DataTable vdmParcaTable = new DataTable();

                             using (var workbook = new XLWorkbook(vdmParcaFilePath))
                             {
                                 var worksheet = workbook.Worksheet(1);
                                 bool firstRow = true;

                                 foreach (var row in worksheet.Rows())
                                 {
                                     if (firstRow)
                                     {
                                         foreach (var cell in row.Cells())
                                             vdmParcaTable.Columns.Add(cell.Value.ToString());
                                         firstRow = false;
                                     }
                                     else
                                     {
                                         vdmParcaTable.Rows.Add();
                                         int i = 0;
                                         foreach (var cell in row.Cells())
                                             vdmParcaTable.Rows[vdmParcaTable.Rows.Count - 1][i++] = cell.Value;
                                     }
                                 }
                             }

                             // Genel toplam dosyası oluşturma
                             var genelToplamResults = (from vdmRow in vdmParcaTable.AsEnumerable()
                                                       join vdmBilgiRow in dtVdmInfo.AsEnumerable()
                                                       on vdmRow.Field<string>(textBox_colomn_name.Text) equals vdmBilgiRow.Field<string>(textBox_colomn_name.Text)
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

                             DataRow newRow = genelToplamTable.NewRow();
                             foreach (DataColumn column in vdmParcaTable.Columns)
                             {
                                 newRow[column.ColumnName] = genelToplamResults[0].vdmRow[column];
                             }

                             foreach (DataColumn column in dtVdmInfo.Columns)
                             {
                                 newRow[column.ColumnName] = genelToplamResults[0].vdmBilgiRow[column];
                             }

                             genelToplamTable.Rows.Add(newRow);

                             SetDataTable_To_Excel(genelToplamTable, vdmKod + "_GenelToplam");
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

                    using (var workbook = new XLWorkbook(mailFilePath))
                    {
                        var worksheet = workbook.Worksheet(1);
                        bool firstRow = true;
                        foreach (var row in worksheet.Rows())
                        {
                            if (firstRow)
                            {
                                foreach (var cell in row.Cells())
                                    dtMailInfo.Columns.Add(cell.Value.ToString());
                                firstRow = false;
                            }
                            else
                            {
                                dtMailInfo.Rows.Add();
                                int i = 0;
                                foreach (var cell in row.Cells())
                                    dtMailInfo.Rows[dtMailInfo.Rows.Count - 1][i++] = cell.Value;
                            }
                        }
                    }
                    string emailBody = mailBody.Text;

                    foreach (DataRow row in dtMailInfo.Rows)
                    {
                        string VDMKodu = row["VDM_Kodu"].ToString();
                        string email = row["Email"].ToString();
                        // string attachmentPath = AppDomain.CurrentDomain.BaseDirectory.ToString() + @"ExcelExport\" + VDMKodu + ".xlsx";
                        string attachmentPath = Path.Combine(folderPath, VDMKodu + ".xlsx");
                        string attachmentPath2 = Path.Combine(folderPath, VDMKodu + "_GenelToplam.xlsx");
                        if (File.Exists(attachmentPath))
                        {
                            SendEmail(email, "Excel Dosyanız", emailBody, attachmentPath);
                        }
                    }

                    MessageBox.Show("Mailler Gönderildi.");
                }
            }
        






        }
        private void SendEmail(string toEmail, string subject, string body, string attachmentPath)
        {
            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");

            mail.From = new MailAddress("gurerg03@gmail.com");
            mail.To.Add(toEmail);
            mail.Subject = subject;
            mail.Body = body;
            Attachment attachment = new Attachment(attachmentPath);
            mail.Attachments.Add(attachment);

            SmtpServer.Port = 587;
            SmtpServer.Credentials = new NetworkCredential("gurerg03@gmail.com", "rnur scsl vydm czbm");
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
    }
}


