using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;

namespace Nüfus_Projeksiyonu
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        public decimal s,s1,s2,s3,s4,s5;
        int Move;
        int Mouse_X;
        int Mouse_Y;
        private void Form1_MouseUp(object sender, MouseEventArgs e)
        {
            Move = 0;
        }

        private void Form1_MouseDown(object sender, MouseEventArgs e)
        {
            Move = 1;
            Mouse_X = e.X;
            Mouse_Y = e.Y;
        }

        private void Form1_MouseMove(object sender, MouseEventArgs e)
        {
            if (Move == 1)
            {
                this.SetDesktopLocation(MousePosition.X - Mouse_X, MousePosition.Y - Mouse_Y);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Iller ıller= new Iller();
            ıller.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Geometrik geometrik= new Geometrik();
            geometrik.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Aritmetik aritmetik = new Aritmetik();
            aritmetik.Show();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            ExcelPackage package= new ExcelPackage();
            package.Workbook.Worksheets.Add("Worksheet1");

            

            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
            worksheet.Cells[1,1].Value = "NÜFUS PROJEKSİYONU";
            worksheet.Cells[1, 1, 1, 6].Merge = true;
            worksheet.Cells[2, 1].Value = "YILLAR";
            worksheet.Cells[2, 2].Value = "NÜFUS";
            worksheet.Cells[2, 3].Value = "HEDEF YILLAR";
            worksheet.Cells[2, 4].Value = "ARİTMETİK NÜFUS";
            worksheet.Cells[2, 5].Value = "GEOMETRİK NÜFUS";
            worksheet.Cells[2, 6].Value = "İLLER BANKASI YÖNTEMİ";
            worksheet.Cells[3, 1].Value = textBox1.Text;
            worksheet.Cells[3, 2].Value = textBox6.Text;
            worksheet.Cells[3, 3].Value = textBox20.Text;
            worksheet.Cells[3, 4].Value = textBox15.Text;
            worksheet.Cells[3, 5].Value = textBox25.Text;
            worksheet.Cells[3, 6].Value = textBox30.Text;
            worksheet.Cells[4, 1].Value = textBox2.Text;
            worksheet.Cells[4, 2].Value = textBox7.Text;
            worksheet.Cells[4, 3].Value = textBox19.Text;
            worksheet.Cells[4, 4].Value = textBox14.Text;
            worksheet.Cells[4, 5].Value = textBox24.Text;
            worksheet.Cells[4, 6].Value = textBox29.Text;
            worksheet.Cells[5, 1].Value = textBox3.Text;
            worksheet.Cells[5, 2].Value = textBox8.Text;
            worksheet.Cells[5, 3].Value = textBox18.Text;
            worksheet.Cells[5, 4].Value = textBox13.Text;
            worksheet.Cells[5, 5].Value = textBox23.Text;
            worksheet.Cells[5, 6].Value = textBox28.Text;
            worksheet.Cells[6, 1].Value = textBox4.Text;
            worksheet.Cells[6, 2].Value = textBox9.Text;
            worksheet.Cells[6, 3].Value = textBox17.Text;
            worksheet.Cells[6, 4].Value = textBox12.Text;
            worksheet.Cells[6, 5].Value = textBox22.Text;
            worksheet.Cells[6, 6].Value = textBox27.Text;
            worksheet.Cells[7, 1].Value = textBox5.Text;
            worksheet.Cells[7, 2].Value = textBox10.Text;
            worksheet.Cells[7, 3].Value = textBox16.Text;
            worksheet.Cells[7, 4].Value = textBox11.Text;
            worksheet.Cells[7, 5].Value = textBox21.Text;
            worksheet.Cells[7, 6].Value = textBox26.Text;

           

            SaveFileDialog saveFileDialog= new SaveFileDialog();
            saveFileDialog.Filter = "Excel Dosyası|Nüfus Projeksiyonu.xlsx";
            saveFileDialog.ShowDialog();

            Stream stream= saveFileDialog.OpenFile();
            package.SaveAs(stream);
            stream.Close();
            MessageBox.Show("Excel Dosyası Oluşturuldu.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);



        }
        public void button4_Click(object sender, EventArgs e)
        {
            try
            {
                //Years
                int y1 = Convert.ToInt32(textBox1.Text);
                int y2 = Convert.ToInt32(textBox2.Text);
                int y3 = Convert.ToInt32(textBox3.Text);
                int y4 = Convert.ToInt32(textBox4.Text);
                int y5 = Convert.ToInt32(textBox5.Text);
                //Population
                int p1 = Convert.ToInt32(textBox6.Text);
                int p2 = Convert.ToInt32(textBox7.Text);
                int p3 = Convert.ToInt32(textBox8.Text);
                int p4 = Convert.ToInt32(textBox9.Text);
                int p5 = Convert.ToInt32(textBox10.Text);
                //Desired year
                int sy1 = Convert.ToInt32(textBox20.Text);
                int sy2 = Convert.ToInt32(textBox19.Text);
                int sy3 = Convert.ToInt32(textBox18.Text);
                int sy4 = Convert.ToInt32(textBox17.Text);
                int sy5 = Convert.ToInt32(textBox16.Text);

                //Calculated population
                //Aritmetik hesaplama
                s = (p2 - p1) / (y2 - y1);
                s1 = (p3 - p2) / (y3 - y2);
                s2 = (p4 - p3) / (y4 - y3);
                s3 = (p5 - p4) / (y5 - y4);
                s4 = (s + s1 + s2 + s3) / 4;
                
                decimal sn1 = p5 + s4 * (sy1 - y5);
                decimal sn2 = p5 + s4 * (sy2 - y5);
                decimal sn3 = p5 + s4 * (sy3 - y5);
                decimal sn4 = p5 + s4 * (sy4 - y5);
                decimal sn5 = p5 + s4 * (sy5 - y5);

                textBox15.Text = sn1.ToString("0.00");
                textBox14.Text = sn2.ToString("0.00");
                textBox13.Text = sn3.ToString("0.00");
                textBox12.Text = sn4.ToString("0.00");
                textBox11.Text = sn5.ToString("0.00");
                //Geometrik Hesaplama
                double g1 = (Math.Log(p2) - Math.Log(p1)) / (y2 - y1);
                double g2 = (Math.Log(p3) - Math.Log(p2)) / (y3 - y2);
                double g3 = (Math.Log(p4) - Math.Log(p3)) / (y4 - y3);
                double g4 = (Math.Log(p5) - Math.Log(p4)) / (y5 - y4);
                double g6 = (g1 + g2 + g3 + g4) / 4;

                double sg1 = Math.Exp(Math.Log(p5) + g6 * (sy1 - y5));
                double sg2 = Math.Exp(Math.Log(p5) + g6 * (sy2 - y5));
                double sg3 = Math.Exp(Math.Log(p5) + g6 * (sy3 - y5));
                double sg4 = Math.Exp(Math.Log(p5) + g6 * (sy4 - y5));
                double sg5 = Math.Exp(Math.Log(p5) + g6 * (sy5 - y5));

                textBox25.Text = sg1.ToString("0.00");
                textBox24.Text = sg2.ToString("0.00");
                textBox23.Text = sg3.ToString("0.00");
                textBox22.Text = sg4.ToString("0.00");
                textBox21.Text = sg5.ToString("0.00");
                //////////////////////////////////////////////////////
                ///
                double ip1 = Math.Pow(y5 - y1, 1.0 / (p1 / p5) - 1);
                if (ip1 <= 1)
                {
                    ip1 = 1;
                }
                if (ip1 >= 3)
                {
                    ip1 = 3;
                }
                if (ip1 <= 3 && ip1 >= 1)
                {
                    ip1 = ip1;
                }
                double iis1 = p5 * Math.Pow((1 + ip1 / 100), sy1 - y5);
                double iis2 = p5 * Math.Pow((1 + ip1 / 100), sy2 - y5);
                double iis3 = p5 * Math.Pow((1 + ip1 / 100), sy3 - y5);
                double iis4 = p5 * Math.Pow((1 + ip1 / 100), sy4 - y5);
                double iis5 = p5 * Math.Pow((1 + ip1 / 100), sy5 - y5);

                textBox30.Text = iis1.ToString("0.00");
                textBox29.Text = iis2.ToString("0.00");
                textBox28.Text = iis3.ToString("0.00");
                textBox27.Text = iis4.ToString("0.00");
                textBox26.Text = iis5.ToString("0.00");
                //////////    ÜSSEL YÖNTEM    ////////////
                
                
            }
            catch (Exception)
            {

                MessageBox.Show("Değerleri Kontrol Ediniz!","Information",MessageBoxButtons.OK,MessageBoxIcon.Warning);
            }
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
