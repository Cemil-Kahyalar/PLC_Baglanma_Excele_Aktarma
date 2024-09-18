using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Microsoft.Office.Interop.Excel;
using S7.Net.Types;
using S7.Net;
using DocumentFormat.OpenXml.Drawing.Diagrams;


namespace PLC_Baglanma_Excelle_Aktarma
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        Plc plc1510;
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {

            using (plc1510 = new Plc(CpuType.S71500, "192.168.0.20", 0, 1))
            {

                plc1510.Open();


                if (plc1510.IsConnected)
                {
                    int dbNumber = 101; // DB101
                    int startByte = 0; // Verinin başlangıç baytı
                    int arrayLength = 100; // Array uzunluğu

                    byte[] Data = plc1510.ReadBytes(DataType.DataBlock, dbNumber, startByte, arrayLength * 2);
                    ushort[] yanit = Word.ToArray(Data);
                    Label2.Text = "plcye bağlanıldı";
                    try
                    {
                        // Excel uygulamasını oluştur
                        Application xla = new Application();
                        xla.Visible = true;
                        Workbook wb = xla.Workbooks.Add(XlSheetType.xlWorksheet);
                        Worksheet ws = (Worksheet)xla.ActiveSheet;

                        // Başlık satırı ekle
                        ws.Cells[1, 1] = "isim";
                        ws.Cells[1, 2] = "deger";
                        ws.Cells[1, 3] = "tarih";




                        for (int i = 0; i <= 99; i++)
                        {

                            ws.Cells[i + 2, 1] = "data" + Convert.ToString(i + 1);
                            ws.Cells[i + 2, 2] = Convert.ToString(yanit[i]);
                            ws.Cells[i + 2, 3] = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        }



                        Label1.Text = $"Excel dosyası başarıyla oluşturuldu ve veriler yazıldı.";
                    }
                    catch (Exception ex)
                    {
                        Label1.Text = $"Dosya oluşturulurken bir hata oluştu: {ex.Message}";
                    }

                }
                else
                {
                    Label2.Text = "plcye bağlanılamadı";
                }

            }
        }
    }
}