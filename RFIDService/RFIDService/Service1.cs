using System;
using System.Collections;
using System.Data.SqlClient;
using System.IO;
using System.ServiceProcess;
using System.Threading;
using System.Timers;
using Symbol.RFID3;
using NiceLabel.SDK;

namespace RFIDService
{
    public partial class Service1 : ServiceBase
    {
        System.Timers.Timer timer = new System.Timers.Timer();

        public Service1()
        {
            InitializeComponent();
        }

        ///Bilgilendirme() fonksiyonuyla çıktıları gözlemleyebilmek adına bir log dosyasına yazdırıyorum.
        public void Bilgilendirme(string mesaj)
        {
            string dosyaYolu = AppDomain.CurrentDomain.BaseDirectory + "/Logs";
            if (!Directory.Exists(dosyaYolu))
            {
                Directory.CreateDirectory(dosyaYolu);
            }
            string textYolu = AppDomain.CurrentDomain.BaseDirectory + "/Logs/rfid_durum.txt";

            if (!File.Exists(textYolu))
            {
                using (StreamWriter sw = File.CreateText(textYolu))
                {
                    sw.WriteLine(mesaj);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(textYolu))
                {
                    sw.WriteLine(mesaj);
                }

            }
        }
        ///RFID Reader'a bağlanıp etiket okuma işlemlerini Baglanti() fonksiyonuyla gerçekleştiriyorum.
        public void Baglanti()
        {
            string hostname = "10.14.50.14";
            RFIDReader rfid3 = new RFIDReader(hostname, 0, 0);
            rfid3.Connect();
            //başlangıçta sensörü pasif hale alıyoruz garanti olsun diye
            //rfid3.Config.GPO[1].PortState = GPOs.GPO_PORT_STATE.FALSE;

            //etiketlerden aldığımız verilerin depolanması üzerine
            TagStorageSettings tagStorageSettings = rfid3.Config.GetTagStorageSettings();
            tagStorageSettings.EnableAccessReports = true;
            rfid3.Config.SetTagStorageSettings(tagStorageSettings);

            //anten değerlerini değiştiriyoruz
            Antennas.AntennaRfConfig antRfConfig = rfid3.Config.Antennas[1].GetRfConfig();
            antRfConfig.AntennaStopTriggerConfig.StopTriggerType =
            ANTENNA_STOP_TRIGGER_TYPE.ANTENNA_STOP_TRIGGER_TYPE_DURATION_MILLISECS;
            antRfConfig.TransmitPowerIndex = 40;  //anten gücü ayarlama
            rfid3.Config.Antennas[1].SetRfConfig(antRfConfig);

            Bilgilendirme("Bağlantı kuruldu. " + DateTime.Now);

            //okuduğum değerleri içine atmak için bir dizi tanımlıyorum
            ArrayList dizi = new ArrayList();

            rfid3.Actions.Inventory.Perform();
            if (rfid3.Config.GPI[1].PortState == GPIs.GPI_PORT_STATE.GPI_PORT_STATE_LOW)
            {
                Bilgilendirme("Tetik Bekleniyor. " + DateTime.Now);
                Thread.Sleep(1000);
            }

            Bilgilendirme("Okuma Başlatıldı. " + DateTime.Now);
            Thread.Sleep(1000);
            TagData[] remainingTags = rfid3.Actions.GetReadTags(150);
            if (remainingTags == null)
            {
                Bilgilendirme("Tag Bulunamıyor. " + DateTime.Now);
            }
            else
            {
                for (int nIndex = 0; nIndex < remainingTags.Length; nIndex++)
                {
                    int sayac = 0;
                    for (int j = 0; j < remainingTags.Length; j++)
                    {
                        if (remainingTags[nIndex].TagID == remainingTags[j].TagID)
                        {
                            sayac++;
                        }
                    }

                    //aynı etiketi her okuduğunda yazmasın ve en az 10 kez okuduysa yazsın diye
                    if (dizi.Contains(remainingTags[nIndex].TagID) == false && sayac >= 10)
                    {
                        dizi.Add(remainingTags[nIndex].TagID);
                        Bilgilendirme(nIndex + 1 + ". EPC: " + remainingTags[nIndex].TagID + " " + DateTime.Now);
                        Bilgilendirme(nIndex + 1 + ".  Tag  " + sayac + " defa okundu.  " + DateTime.Now);
                    }
                }
            }
            //ilk okunup diziye atanan epc değerini etiket oluşturmak üzere NiceLabel() fonksiyonuna gönderiyorum
            string epc = dizi[0].ToString();
            Bilgilendirme(epc);
            NiceLabel(epc);
            rfid3.Actions.Inventory.Stop();
            rfid3.Disconnect();
        }
        //NiceLabel fonksiyonunda daha önce elde ettiğim epc'nin bilgilerini databaseden çekip etiket üzerine yazdırıyorum.
        public void NiceLabel (string epc)
        {
            //PrintEngineFactory.PrintEngine.Initialize();
            IPrintEngine printEngine = PrintEngineFactory.PrintEngine;
            printEngine.Initialize();
            ILabel label = PrintEngineFactory.PrintEngine.OpenLabel("C:\\Users\\oznur.hasoglu\\Desktop\\rfid\\NL\\nicelabel_deneme\\nicelabel_deneme\\bin\\Debug\\Etiket1.nlbl");

            SqlCommand komut = new SqlCommand();
            SqlDataReader reader;
            SqlConnection sqlBaglantisi = new SqlConnection("server=.;Initial Catalog=DENEME;Integrated Security=SSPI");

            komut.Connection = sqlBaglantisi;
            komut.CommandText = "SELECT* FROM KUMAS WHERE EPC='" + epc + "'";
            sqlBaglantisi.Open();
            reader = komut.ExecuteReader();

            while (reader.Read())
            {
                label.Variables["musteri"].SetValue(reader[2].ToString());
                label.Variables["kumas_adi"].SetValue(reader[3].ToString());
                label.Variables["renk_kodu"].SetValue(reader[4].ToString());
                label.Variables["urun_no"].SetValue(reader[5].ToString());
                label.Variables["boyut"].SetValue(reader[6].ToString());
                label.Variables["tarih"].SetValue(reader[7].ToString());
            }

            sqlBaglantisi.Close();
            //label.PrintSettings.PrinterName = "Yazıcı Adı";
            //label.PrintSettings.OutputFileName = @"C:\Users\oznur.hasoglu\Desktop\rfid\NL\nicelabel_deneme\nicelabel_deneme\e.pdf";

            IPrintToGraphicsSettings pgs = new PrintToGraphicsSettings();
            pgs.PrintToFiles = true;
            pgs.Quantity = 1;
            pgs.ImageFormat = "png";
            pgs.DestinationFolder = "C:\\Users\\oznur.hasoglu\\Desktop\\rfid\\NL\\nicelabel_deneme\\nicelabel_deneme";
            pgs.PrintToFiles = true;
            pgs.PrintAll = true;
            label.PrintToGraphics(pgs);

            printEngine.Shutdown();
        }

        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            Baglanti();
            Bilgilendirme("Servis Çalışmaya Devam Ediyor. " + DateTime.Now);           
        }

        protected override void OnStart(string[] args)
        {
            Bilgilendirme("Servis Başlatıldı. " + DateTime.Now);           
            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            timer.Enabled = true;
            timer.Interval = 5000;          
        }
        protected override void OnStop()
        {
            Bilgilendirme("Servis Durduruldu. " + DateTime.Now);
        }
    }
}
