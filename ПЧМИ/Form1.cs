using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using System.IO;

namespace ПЧМИ
{
    public partial class Form1 : Form
    {
        private string fio1;
        private string adr1;
        private string fio2;
        private string adr2;
        private int wei;
        private string weigth;
        private string delivery;
        private int price;
        Random rnd = new Random();
        private int tracknumber;
        private string other;
        private string pay;

        public Form1()
        {
            InitializeComponent();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Clear(); textBox2.Clear(); textBox3.Clear(); textBox4.Clear(); textBox5.Clear();
            textBox6.Clear(); textBox7.Clear(); textBox8.Clear();
            radioButton1.Checked = false; radioButton2.Checked = false;
            radioButton3.Checked = false; radioButton4.Checked = false;
            radioButton5.Checked = false; radioButton6.Checked = false;
            checkBox1.Checked = false; checkBox2.Checked = false;
            tabControl1.SelectedIndex = 1;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 5;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 0;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            fio1 = textBox1.Text;
            adr1 = textBox2.Text;
            fio2 = textBox4.Text;
            adr2 = textBox3.Text;
            weigth = textBox5.Text;
            
            if (string.IsNullOrEmpty(fio1) || string.IsNullOrEmpty(adr1) || string.IsNullOrEmpty(fio2) || string.IsNullOrEmpty(adr2) || string.IsNullOrEmpty(weigth))
            {
               MessageBox.Show("Заполните все поля", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
               return;
            }
            else
                tabControl1.SelectedIndex = 2;

            if (int.TryParse(weigth, out wei))
            {
                Console.WriteLine("Число: " + wei);
            }
            else
            {
                MessageBox.Show("Значение веса посылки не является числом", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }

        private void button5_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 1;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (!radioButton1.Checked && !radioButton2.Checked && !radioButton3.Checked)
            {
                MessageBox.Show("Выберите один из вариантов доставки", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
                tabControl1.SelectedIndex = 3;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if ((checkBox1.Checked && textBox6.Text == "") || (checkBox2.Checked && textBox7.Text == ""))
            {
                MessageBox.Show("Заполните все поля", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (!radioButton4.Checked && !radioButton5.Checked && !radioButton6.Checked)
            {
                MessageBox.Show("Выберите один из вариантов оплаты", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
                tabControl1.SelectedIndex = 4;
            if (radioButton4.Checked) pay = "Наличные";
            else if (radioButton5.Checked) pay = "Банковская карта";
            else if (radioButton6.Checked) pay = "СБП";
            tracknumber = rnd.Next(100000, 1000000);
            label41.Text = tracknumber.ToString();
            if (checkBox1.Checked && textBox6.Text != "") other += textBox6.Text + "\n";
            if (checkBox2.Checked && textBox7.Text != "") other += textBox7.Text + "\n";
            FileInfo fileInfo = new FileInfo(@"C:\Users\Windows 10\Desktop\pochta.xlsx");
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            if (!fileInfo.Exists)
            {
                // Если файл не существует, создаем новый
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    // Добавляем новый лист
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Посылки");

                    // Записываем заголовки в ячейки
                    worksheet.Cells["A1"].Value = "ID";
                    worksheet.Cells["B1"].Value = "ФИО отправителя";
                    worksheet.Cells["C1"].Value = "Адрес отправителя";
                    worksheet.Cells["D1"].Value = "ФИО получателя";
                    worksheet.Cells["E1"].Value = "Адрес получателя";
                    worksheet.Cells["F1"].Value = "Вид доставки";
                    worksheet.Cells["G1"].Value = "Дополнительные опции";
                    worksheet.Cells["H1"].Value = "Способ оплаты";
                    worksheet.Cells["I1"].Value = "Трэк-номер";
                    worksheet.Cells["J1"].Value = "Цена (руб.)";
                    worksheet.Cells["K1"].Value = "Местоположение";
                    worksheet.Cells["L1"].Value = "Статус посылки";
                    // Сохраняем файл
                    package.Save();
                }
            }

            // Открываем существующий файл и добавляем новые записи
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Посылки"];

                // Находим первую пустую строку для добавления новой записи
                int row = worksheet.Dimension.Rows + 1;

                worksheet.Cells[row, 1].Value = row-1;
                worksheet.Cells[row, 2].Value = fio1;
                worksheet.Cells[row, 3].Value = adr1;
                worksheet.Cells[row, 4].Value = fio2;
                worksheet.Cells[row, 5].Value = adr2;
                worksheet.Cells[row, 6].Value = delivery;
                worksheet.Cells[row, 7].Value = other;
                worksheet.Cells[row, 8].Value = pay;
                worksheet.Cells[row, 9].Value = tracknumber;
                worksheet.Cells[row, 10].Value = price;

                // Сохраняем файл с добавленными данными
                package.Save();
            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 2;
        }

        private void button11_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 0;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 5;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 0;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                textBox6.Enabled = true;
                price += 50;
                pricelabel.Text = price.ToString() + " руб.";
            }
            else
            {
                textBox6.Enabled = false;
                price -= 50;
                pricelabel.Text = price.ToString() + " руб.";
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked)
            {
                textBox7.Enabled = true;
                price += 30;
                pricelabel.Text = price.ToString() + " руб.";
            }
            else
            {
                textBox7.Enabled = false;
                price -= 30;
                pricelabel.Text = price.ToString() + " руб.";
            }
        }

        private void tabPage4_Enter(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                deliverylabel.Text = "Обычная";
                price = 250;
            }
            else if (radioButton2.Checked)
            {
                deliverylabel.Text = "Ускоренная";
                price = 500;
            }

            else if (radioButton3.Checked)
            {
                deliverylabel.Text = "Курьерская";
                price = 800;
            }
            delivery = deliverylabel.Text;
            wlabel.Text = weigth + "  г";
            price = (int)(price + wei * 0.3);
            pricelabel.Text = price.ToString() + "   руб.";
        }

        private void tabPage5_Enter(object sender, EventArgs e)
        {
            label43.Text = pricelabel.Text;
            label46.Text = DateTime.Now.ToString("dd.MM.yyyy");
        }

        private void tabPage6_Enter(object sender, EventArgs e)
        {
            
        }

        private void button10_Click(object sender, EventArgs e)
        {
            string searchTerm = textBox8.Text.Trim(); // Получаем значение из TextBox и удаляем лишние пробелы

            if (string.IsNullOrEmpty(searchTerm))
            {
                MessageBox.Show("Пожалуйста, введите трек-номер для поиска.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return; // Прерываем выполнение метода, если трек-номер не введен
            }
            FileInfo fileInfo = new FileInfo(@"C:\Users\Windows 10\Desktop\pochta.xlsx");
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Посылки"];

                // Перебираем ячейки в столбце с трек-номерами
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    string trackNumber = worksheet.Cells[row, 9].GetValue<string>(); // Получаем значение трек-номера из текущей строки

                    // Сравниваем трек-номер с искомым значением
                    if (trackNumber == searchTerm)
                    {
                        string fio1 = worksheet.Cells[row, 2].GetValue<string>();
                        string adr1 = worksheet.Cells[row, 3].GetValue<string>();
                        string fio2 = worksheet.Cells[row, 4].GetValue<string>();
                        string adr2 = worksheet.Cells[row, 5].GetValue<string>();
                        string delivery = worksheet.Cells[row, 6].GetValue<string>();
                        string other = worksheet.Cells[row, 7].GetValue<string>();
                        string paymentMethod = worksheet.Cells[row, 8].GetValue<string>();
                        string tracknumber = worksheet.Cells[row, 9].GetValue<string>();
                        string price = worksheet.Cells[row, 10].GetValue<string>();
                        string location = worksheet.Cells[row, 11].GetValue<string>();
                        string status = worksheet.Cells[row, 12].GetValue<string>();

                        MessageBox.Show($"Местоположение: {location}\nСтатус посылки: {status}\n ФИО отправителя: {fio1}\nАдрес отправителя: {adr1}\nФИО получателя: {fio2}\nАдрес получателя: {adr2}\nВид доставки: {delivery}\nДополнительные опции: {other}\nСпособ оплаты: {paymentMethod}\nТрек-номер: {trackNumber}\nЦена (руб.): {price}", "Результат поиска", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return; // Прерываем выполнение метода после вывода первого совпадения
                    }
                }
                
                // Если ничего не найдено
                MessageBox.Show("Трек-номер не найден.", "Результат поиска", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}
