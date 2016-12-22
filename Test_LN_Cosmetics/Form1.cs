using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.IO.Compression;
using System.Net;

namespace Test_LN_Cosmetics
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public List<Price> ReadFromExcel(OleDbConnection connection, out String err)
        {
            err = "";
            List<Price> res = new List<Price>();
            OleDbCommand command;
            OleDbDataReader reader=null;
            try
            {
                command = new OleDbCommand("SELECT * FROM [Лист2$]", connection);
                reader = command.ExecuteReader();
                reader.Read();
                int j = 0;
                do
                {
                    j += 1;
                    if (j < 3) continue;
                    Price r = new Price();
                    int t = 0;
                    decimal d = 0;
                    Int64 t64 = 0;
                    DateTime dt = DateTime.MinValue;
                    if (int.TryParse(reader[0].ToString(), out t)) r.Cod=t;
                    else
                    {
                        ErrMes("Ошибка в коде товара", out err);
                        return res;
                    }
                    r.Name = reader[1].ToString();
                    r.MNF = reader[2].ToString();
                    r.CNTR = reader[3].ToString();
                    if (DateTime.TryParse(reader[4].ToString(), out dt)) r.Srok = dt;
                    else
                    {
                        ErrMes("Ошибка в дате срока годности", out err);
                        return res;
                    }
                    t = 0;
                    int.TryParse(reader[5].ToString(), out t);
                    r.Kol = t;
                    if (decimal.TryParse(reader[6].ToString(), out d)) r.Cena=d;
                    else
                    {
                        ErrMes("Ошибка в цене товара", out err);
                        return res;
                    }
                    t = 1;
                    int.TryParse(reader[7].ToString(), out t); 
                    r.Kratnost = t;
                    if (Int64.TryParse(reader[8].ToString(), out t64)) r.Barcode = t64;
                    else
                    {
                        ErrMes("Ошибка в коде", out err);
                        return res;
                    }
                    t = 18;
                    int.TryParse(reader[9].ToString(), out t);
                    r.Ratends = t;
                    res.Add(r);


                } while (reader.Read());
            }
            catch(Exception e)
            {
                err = e.Message;
            }
            finally
            {
                if(reader.IsClosed==false) reader.Close();
                reader.Dispose();
               
            }
            return res;
        }

        public void ErrMes(string err, out string em)
        {
            MessageBox.Show(err);
            em = "Ошибка в данных";
        }
        public void CreateTextFile(string fName, List<Price> PriceList, out String err)
        {
            err = "";
            try
            {
                using (System.IO.StreamWriter file = new System.IO.StreamWriter(fName, true, Encoding.UTF8))
                {
                    StringBuilder str = new StringBuilder();
                    str.Append("{\"providerName\": \"ООО НПЛ \"ЛН-Косметика\",");
                    file.WriteLine(str.ToString());
                    str = new StringBuilder();
                    str.Append("\"updateDate\": \"");
                    str.Append(DateTime.Today.ToShortDateString());
                    str.Append("\",");
                    file.WriteLine(str.ToString());


                    str = new StringBuilder();
                    str.Append("\"items\":[{");
                    file.WriteLine(str.ToString());
                    int i = 0;
                    foreach (Price item in PriceList)
                    {
                        if (i > 0)
                        {
                            str = new StringBuilder();
                            str.Append("},{");
                            file.WriteLine(str.ToString());
                        }
                        //Код препарата
                        str = new StringBuilder();
                        str.Append("\"code\": \"");
                        str.Append(item.Cod.ToString());
                        str.Append("\",");
                        file.WriteLine(str.ToString());

                        //Название препарата
                        str = new StringBuilder();
                        str.Append("\"name\": \"");
                        str.Append(item.Name.ToString());
                        str.Append("\",");
                        file.WriteLine(str.ToString());

                        //Производитель
                        str = new StringBuilder();
                        str.Append("\"manufacturer\": \"");
                        str.Append(item.MNF.ToString());
                        str.Append("\",");
                        file.WriteLine(str.ToString());


                        //Страна
                        str = new StringBuilder();
                        str.Append("\"manufacturerCountry\": \"");
                        str.Append(item.CNTR.ToString());
                        str.Append("\",");
                        file.WriteLine(str.ToString());

                        //СКол-во товара на складе поставщика
                        str = new StringBuilder();
                        str.Append("\"quantity\": ");
                        str.Append(item.Kol.ToString());
                        str.Append(",");
                        file.WriteLine(str.ToString());

                        //Штрих-код
                        str = new StringBuilder();
                        str.Append("\"barcode\": \"");
                        str.Append(item.Barcode.ToString());
                        str.Append("\",");
                        file.WriteLine(str.ToString());

                        //Цена
                        str = new StringBuilder();
                        str.Append("\"price\": ");
                        str.Append(item.Cena.ToString());
                        str.Append(",");
                        file.WriteLine(str.ToString());

                        //Крастность
                        str = new StringBuilder();
                        str.Append("\"multiplisity\": ");
                        str.Append(item.Kratnost.ToString());
                        str.Append(",");
                        file.WriteLine(str.ToString());

                        //НДС
                        str = new StringBuilder();
                        str.Append("\"ratends\": ");
                        str.Append(item.Ratends.ToString());
                        str.Append(",");
                        file.WriteLine(str.ToString());

                        //Срок годности
                        str = new StringBuilder();
                        str.Append("\"exprirationDate\": \"");
                        str.Append(item.Srok.ToShortDateString());
                        str.Append("\",");
                        file.WriteLine(str.ToString());
                        i += 1;

                    }

                    str = new StringBuilder();
                    str.Append("}]}");
                    file.WriteLine(str.ToString());
                }
            }
            catch (Exception e)
            {
                err = e.Message;
                return;
            }
            }

        
        private void Form1_Load(object sender, EventArgs e)
        {
            label2.Visible = false;
        }

         public static void SendToFTP (string ArchName, string FTPServerName)
        {

            FtpWebRequest request = (FtpWebRequest)WebRequest.Create(FTPServerName);
            request.Method = WebRequestMethods.Ftp.UploadFile;

          
            request.Credentials = new NetworkCredential ("anonymous","ponomarevase@mail.ru");

           
            StreamReader sourceStream = new StreamReader(ArchName);
            byte [] fileContents = Encoding.UTF8.GetBytes(sourceStream.ReadToEnd());
            sourceStream.Close();
            request.ContentLength = fileContents.Length;

            Stream requestStream = request.GetRequestStream();
            requestStream.Write(fileContents, 0, fileContents.Length);
            requestStream.Close();

            FtpWebResponse response = (FtpWebResponse)request.GetResponse();

            Console.WriteLine("Upload File Complete, status {0}", response.StatusDescription);

            response.Close();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
           
           
            string SourseFilePath = System.IO.Path.Combine(Application.StartupPath, "test.xls");
            OleDbConnection connection = new OleDbConnection(@"provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + SourseFilePath + "';Extended Properties=Excel 8.0;");
            try
            {
                connection.Open();
            }
            catch
            {
                MessageBox.Show("Ошибка при открытии прайс-листа!");
                return;
            }
            String err = "";
            List<Price> priceList = ReadFromExcel(connection, out err);

           if (String.IsNullOrEmpty(err) == false)
           {
               MessageBox.Show("Ошибка при чтении прайс-листа! " + err);
               return;
           }
            string path = Application.StartupPath + @"\temp";
            System.IO.Directory.CreateDirectory(path);
            DateTime todayDate = DateTime.Today;
            string date_string = todayDate.Day.ToString() + todayDate.Month.ToString() + todayDate.Year.ToString();
            string FilePath = path + @"\LNКосметика + " + date_string + ".txt";
           

            CreateTextFile(FilePath, priceList, out err);
            if (String.IsNullOrEmpty(err) == false)
            {
                MessageBox.Show("Ошибка при создании файла! " + err);
                return;
            }
            try
            {
                string pathZip = Application.StartupPath + @"\temp_zip";

                if (Directory.Exists(pathZip) == false) System.IO.Directory.CreateDirectory(pathZip);

                string zipPath = pathZip + @"\LNКосметика + " + date_string + ".zip";
                if (File.Exists(zipPath)) File.Delete(zipPath);

                ZipFile.CreateFromDirectory(path, zipPath);
                if (Directory.Exists(path)) Directory.Delete(path, true);

                string m = "Архив находится в папке " + zipPath;
                label2.Visible = true;
                label2.Text = m;

                string ftpText = ftpNameTextBox.Text;
                if (String.IsNullOrEmpty(ftpText))
                {
                    MessageBox.Show("Не указан адрес ftp-сервера");
                    return;
                }
                else
                {
                    try
                    {
                        SendToFTP(zipPath, ftpText);
                    }
                    catch
                    {
                        MessageBox.Show("Ошибка соединения с ftp-сервером. Проверьте адрес и настройки соединения. Архив создан на диске");
                        return;
                    }

                }

                MessageBox.Show("Все!");
            }
            catch
            {
                MessageBox.Show("Произошла непредвиденная ошибка");
                return;
            }
        }
    }
}
