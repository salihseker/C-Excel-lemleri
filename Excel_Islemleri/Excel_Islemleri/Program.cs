using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;
using System.Drawing;
using System.Data;
using OfficeOpenXml.Drawing;

namespace Excel_Islemleri
{
    class Program
    {
        static void Main(string[] args)
        {
            //ExcelOlustur();
            //ExcelOku();
            ExcelIslemleri();

        }

        private static void ExcelOlustur() {
            string strFileName = @"D:\excelOlustur.xlsx";
            FileInfo exc_file = new FileInfo(strFileName);
            ExcelPackage pck = new ExcelPackage();
            ExcelWorkbook wb = pck.Workbook;//Excel Dosyası oluşturulur.
            ExcelWorksheet worksheet = pck.Workbook.Worksheets.Add("1.Sayfa");
            pck.SaveAs(exc_file);
            pck.Dispose();

        }

        private static void ExcelOku()
        {
            string strFileName = @"D:\excelOlustur.xlsx";//Okunacak excelin dosya yolu.
            FileInfo exc_file = new FileInfo(strFileName);
            ExcelPackage pck = new ExcelPackage(exc_file);
            ExcelWorkbook wb = pck.Workbook;//Excel Dosyası oluşturulur.
            Console.WriteLine("Excel Sayfa Sayısı : " + wb.Worksheets.Count());
            foreach (var item in wb.Worksheets)
            {
                Console.WriteLine("Sayfa İsmi : " + item.Name);
            }

            pck.Dispose();
            Console.ReadLine();

        }
        private static void ExcelIslemleri()
        {
            string strFileName = @"D:\excelIslemleri.xlsx";
            FileInfo exc_file = new FileInfo(strFileName);//Excelimizin kayıt olacağı yolu belirtiyoruz.

            //ExcelPackage pck = new ExcelPackage(@"D:\excelIslemleri.xlsx"); mevcut exceli okumak için kullanılır.
            ExcelPackage pck = new ExcelPackage();
            ExcelWorkbook wb = pck.Workbook;//Excel Dosyası oluşturulur.
            ExcelWorksheet worksheet = pck.Workbook.Worksheets.Add("ilkSayfa");//Excel için sayfa oluşturulur.
            ExcelWorksheet worksheet2 = pck.Workbook.Worksheets.Add("ikiniciSayfa");
            ExcelWorksheet worksheet3 = pck.Workbook.Worksheets.Add("ucuncuSayfa");

            worksheet2.Hidden = eWorkSheetHidden.Hidden;//eWorkSheetHidden.Visible , eWorkSheetHidden.VeryHidden // sayfayı gizleme açma işlemleri.

            //Stil işlemleri (Sayfaların stilleri için örnek ayarlar)            
            //worksheet.TabColor = Color.Blue;
            //worksheet.DefaultRowHeight = 12;
            //worksheet.HeaderFooter.FirstFooter.LeftAlignedText = string.Format("Generated: {0}", DateTime.Now.ToShortDateString());
            //worksheet.Row(1).Height = 15;
            //worksheet.Row(2).Height = 15;

            // Excel Sayfa Başlıkları Ekleme
            // Excel İlk Satırlar 
            //string deger = worksheet.Cells["A2:H1"].ToString(); //Farklı bir örnek kullanım.
            worksheet.Cells[1, 1].Value = "ID";//worksheet.Cells[satır , sütun] , excel hücresi. 
            worksheet.Cells[1, 2].Value = "Ad";
            worksheet.Cells[1, 3].Value = "Soyad";

            // Add the second row of header data
            worksheet.Cells[2, 1].Value = 100;
            worksheet.Cells[2, 2].Value = "Salih";
            worksheet.Cells[2, 3].Value = "ŞEKER";
            worksheet.Cells[2, 4].Value = 25;

            YorumEkle(worksheet , 1 ,1,"test yorum" , "Salih");
            //ResimEkle(worksheet, 3, 1, @"D:\resim1.jpg");

            worksheet.Cells[2, 5].Formula = "Sum(A2:D2)";//Excel Formulleri eklemek için 


            //ExcelRange range = Excelde belirli bir alana işlem uygulayacağımız zaman kullanılır.
            //ExcelRange range = worksheet.Cells[Baslangıç Satırı, Baslangıç Sütunu, Bitiş Satırı, Bitiş Sütunu]
            //using (ExcelRange range = worksheet.Cells[1, 1, 1, 2])
            //{
            //    range.Style.Font.Bold = true;
            //    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            //    range.Style.Fill.BackgroundColor.SetColor(Color.Black);
            //    range.Style.Font.Color.SetColor(Color.WhiteSmoke);
            //    range.Style.ShrinkToFit = false;
            //}

            // Sütunları İçeriğe göre sınırlandırmak
            worksheet.Column(1).AutoFit();
            worksheet.Column(2).AutoFit();
            worksheet.Column(3).AutoFit();

            //Excel Dökümanının Özellikleri ekelmek istersek
            pck.Workbook.Properties.Title = "Kullanıcılar";//Başlık
            pck.Workbook.Properties.Author = "Salih ŞEKER";//Yazar
            pck.Workbook.Properties.Company = "Sugar CORP";//Şirket


            pck.SaveAs(exc_file);
            pck.Dispose();
        }

        //Excel Hücresine Yorum ekleme
        private static void YorumEkle(ExcelWorksheet ws, int sutunIndex, int satirIndex, string yorum, string yazar)
        {
            //Adding a comment to a Cell
            var commentCell = ws.Cells[satirIndex, sutunIndex];
            commentCell.AddComment(yorum, yazar);
        }

        private static void ResimEkle(ExcelWorksheet ws, int sutunIndex, int satirIndex, string resimYol)
        {
            //How to Add a Image using EP Plus
            Bitmap image = new Bitmap(resimYol);
            ExcelPicture picture = null;
            if (image != null)
            {
                picture = ws.Drawings.AddPicture("pic" + satirIndex.ToString() + sutunIndex.ToString(), image);
                picture.From.Column = sutunIndex;
                picture.From.Row = satirIndex;
                picture.SetSize(100, 100);
            }
        }

        //Excelde belirtilen sayfaya datatble daki tablo yu eklemek için kullnılan method. 
        private static void SheetDoldur(ExcelWorksheet worksheet, DataTable Table, int baslangicSatiri, int baslangicSutunu, string type = "")
        {


            int clCount = Table.Columns.Count;
            int rwCount = Table.Rows.Count;
            for (int i = baslangicSatiri; i < rwCount + baslangicSatiri; i++)
            {
                for (int j = baslangicSutunu; j < clCount + baslangicSutunu; j++)
                {

                    if (!string.IsNullOrEmpty(Table.Rows[i - baslangicSatiri][j - baslangicSutunu].ToString()) && IsNumeric(Table.Rows[i - baslangicSatiri][j - baslangicSutunu].ToString()) && Table.Rows[i - baslangicSatiri][j - baslangicSutunu].ToString().Length < 10)
                        worksheet.Cells[i, j].Value = Convert.ToDouble(Table.Rows[i - baslangicSatiri][j - baslangicSutunu]);
                    else
                        worksheet.Cells[i, j].Value = Table.Rows[i - baslangicSatiri][j - baslangicSutunu];

                }


            }

        }

        //sayısal değer olup olmadığı kontrol etmek için oluşturulmuş method
        static bool IsNumeric(string text)
        {
            foreach (char chr in text)
            {
                if (!Char.IsNumber(chr) && chr != '.' && chr != ',') return false;
            }
            return true;
        }
    }
}
