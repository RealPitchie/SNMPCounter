using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using SnmpSharpNet;
using Excel = Microsoft.Office.Interop.Excel;

namespace CounterXML
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        // Объявление переменных
        Excel.Application xlApp = new Excel.Application();
        OpenFileDialog OPF = new OpenFileDialog();
        OpenFileDialog OPF2 = new OpenFileDialog();
        Encoding enc = Encoding.GetEncoding(1251);
        

        private void button1_Click_1(object sender, EventArgs e)
        {
            {

                //к кому будем обращаться
                string host = "192.168.0.95";

                //очищаем область ответа
                //responseTextBox.Text = "";
                string community = "public";
                //создаем запрос
                var snmp = new SimpleSnmp(host, community);
                //если запрос не валиден, то пишем об этом
                if (!snmp.Valid)
                {
                    MessageBox.Show("Запрос не валиден", "Error");
                    return;
                }
                //формируем тело запроса и отсылаем его
                Dictionary<Oid, AsnType> result = snmp.Get(SnmpVersion.Ver2, new[]
                    {

                        ".1.3.6.1.4.1.253.8.53.13.2.1.6.103.20.3", //Копии А4
                        ".1.3.6.1.4.1.253.8.53.13.2.1.6.1.20.7", //Печать А4
                        ".1.3.6.1.4.1.253.8.53.13.2.1.6.1.20.47", //Печать А3
                        ".1.3.6.1.4.1.253.8.53.13.2.1.6.103.20.6", //Копии А3
           
		            });
                //если нет ответа, то так и скажем
                if (result == null)
                {
                    MessageBox.Show("No response", "Error");
                    return;
                }
                //счетчик для форматирования

                int i = 0;

                foreach (var kvp in result)
                {
                    int n = 0;
                    //если есть ответ, то выводим
                   
                    switch (i)
                    {

                        case 0:
                            dataGridView1.Rows[0].Cells[i].Value = kvp.Value;
                            break;
                        case 1:
                            dataGridView1.Rows[0].Cells[i].Value = kvp.Value;
                            break;
                        case 2:
                            dataGridView1.Rows[0].Cells[i].Value = kvp.Value;
                            break;
                        case 3:
                            dataGridView1.Rows[0].Cells[i].Value = kvp.Value;
                            break;

                    };
                    dataGridView1.Rows[0].Cells[i].Value = kvp.Value;
                    i++;
                    n++;
                   
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //к кому будем обращаться
            string host = "192.168.0.24";

            //очищаем область ответа
            // responseTextBox.Text = "";
            string community = "public";
            //создаем запрос
            var snmp = new SimpleSnmp(host, community);
            //если запрос не валиден, то пишем об этом
            if (!snmp.Valid)
            {
                MessageBox.Show("Запрос не валиден", "Error");
                // responseTextBox.Text += Resources.Not_Valid_SNMP_HOST;
                return;
            }
            //формируем тело запроса и отсылаем его
            Dictionary<Oid, AsnType> result = snmp.Get(SnmpVersion.Ver2, new[]
                {
                //".1.3.6.1.2.1"/*,
                ".1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.1", //Черно-белые копии А4
                ".1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.2", //Черно-белая печать А4
                ".1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.7.1.2", //Черно-белая печать А3
                ".1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.7.1.1", //Черно-белые копии А3
                ".1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.1", //Цветные копии А4
                ".1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.7.2.1", //Цветные копии А3
                ".1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.2.2", //Цветная печать А4
                ".1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.7.2.2", //Цветная печать А3
                });
            //если нет ответа, то так и скажем
            if (result == null)
            {
                MessageBox.Show("No response", "Error");
                // responseTextBox.Text += Resources.NoResults;
                return;
            }
            //счетчик для форматирования
            int i = 0;
            foreach (var kvp in result)
            {
                int n = 0;

                int m = 0;
                //если есть ответ, то выводим
                switch (i)
                {
                    case 0:
                        dataGridView2.Rows[0].Cells[i].Value = kvp.Value;
                        break;
                    case 1:
                        dataGridView2.Rows[0].Cells[i].Value = kvp.Value;
                        break;
                    case 2:
                        dataGridView2.Rows[0].Cells[i].Value = kvp.Value;
                        break;
                    case 3:
                        dataGridView2.Rows[0].Cells[i].Value = kvp.Value;
                        break;
                    case 4:
                        dataGridView2.Rows[0].Cells[i].Value = kvp.Value;
                        break;
                    case 5:
                        dataGridView2.Rows[0].Cells[i].Value = kvp.Value;
                        break;
                    case 6:
                        dataGridView2.Rows[0].Cells[i].Value = kvp.Value;
                        break;
                    case 7:
                        dataGridView2.Rows[0].Cells[i].Value = kvp.Value;
                        break;

                };
                dataGridView2.Rows[0].Cells[i].Value = kvp.Value;

                i++;
                n++;
                m++;
            }




        }

        private void button3_Click(object sender, EventArgs e)
        {
            /* Пробуем запись в Excel
            //try
            //  {

                    DataSet ds = new DataSet(); // создаем пока что пустой кэш данных
                    DataTable dt = new DataTable(); // создаем пока что пустую таблицу данных
                    dt.TableName = "Xerox D95"; // название таблицы
                    dt.Columns.Add("Копия А4 чб"); // название колонок
                    dt.Columns.Add("Печать А4 чб");
                    dt.Columns.Add("Печать А3 чб");
                    dt.Columns.Add("Копия А3 чб");
                    ds.Tables.Add(dt); //в ds создается таблица, с названием и колонками, созданными выше

                    foreach (DataGridViewRow r in dataGridView1.Rows) // пока в dataGridView1 есть строки
                    {
                        DataRow row = ds.Tables["Xerox D95"].NewRow(); // создаем новую строку в таблице, занесенной в ds
                        row["Копия А4 чб"] = r.Cells[0].Value;  //в столбец этой строки заносим данные из первого столбца dataGridView1
                        row["Печать А4 чб"] = r.Cells[1].Value; // то же самое со вторыми столбцами
                        row["Печать А3 чб"] = r.Cells[2].Value; //то же самое с третьими столбцами
                        row["Печать А4 чб"] = r.Cells[3].Value;
                        ds.Tables["Xerox D95"].Rows.Add(row); //добавление всей этой строки в таблицу ds.
                    }


                    Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
                    Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
                    //Книга.
                    ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
                    //Таблица.
                    ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

                    for (int i = 1; i < dataGridView1.Rows.Count; i++)
                    {
                        for (int j = 1; j < dataGridView1.ColumnCount; j++)
                        {
                            ExcelWorkSheet.Cells[i + 1, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                        }
                    }
                    //Вызываем нашу созданную эксельку.
                    ExcelApp.Visible = true;
                    ExcelApp.UserControl = true;


                    MessageBox.Show("XML файл успешно сохранен.", "Выполнено.");
            // }
            //catch
            //{
            //  MessageBox.Show("Невозможно сохранить XML файл.", "Ошибка.");
            //}
            */

            // Запись в Excel, дубль два


            //try
            //{
            //    Excel.Application ObjExcel = new Excel.Application();
            //    Excel.Workbook ObjWorkBook;
            //    Excel.Worksheet ObjWorkSheet;
            //    ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);
            //    ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            //    Excel.Range _excelCells1 = ObjWorkSheet.get_Range("F1", "I1").Cells;
            //    _excelCells1.Columns.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
            //    // Производим объединение
            //    _excelCells1.Merge(Type.Missing);
            //    ObjWorkSheet.Cells[1, 6] = "Konica 951";

            //    // Выделяем диапазон ячеек от O1 до Q1         
            //    Excel.Range _excelCells2 = ObjWorkSheet.get_Range("B1", "E1").Cells;
            //    // Производим объединение
            //    _excelCells2.Merge(Type.Missing);
            //    _excelCells2.Columns.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
            //    ObjWorkSheet.Cells[1, 2] = "Xerox D95";

            //    ObjWorkSheet.Cells[3, 1] = "51";
            //    ObjWorkBook.SaveAs("D:\\отчет за " + DateTime.Today.Date.ToShortDateString() + ".xlsx");
            //    /**/
                
            //    ObjExcel.Visible = true;
            //   // ObjExcel.Quit();
            //}
            //catch (Exception exc)
            //{
            //    MessageBox.Show("Ошибка при составлении лога\n" + exc.Message);
            //}
            

        }


        /*
         * // Выделяем диапазон ячеек от H1 до K1         
                    Excel.Range _excelCells1 = (Excel.Range)workSheet.get_Range("H1", "K1").Cells;
                    // Производим объединение
                    _excelCells1.Merge(Type.Missing);
                    workSheet.Cells[1, 8] = "Общие";

                    // Выделяем диапазон ячеек от O1 до Q1         
                    Excel.Range _excelCells2 = (Excel.Range)workSheet.get_Range("O1", "Q1").Cells;
                    // Производим объединение
                    _excelCells2.Merge(Type.Missing);
                    workSheet.Cells[1, 15] = "Общие";
       
         * */


        private void button4_Click(object sender, EventArgs e)
        {
            //к кому будем обращаться
            string host = "192.168.0.150";

            //очищаем область ответа
            //responseTextBox.Text = "";
            string community = "public";
            //создаем запрос
            var snmp = new SimpleSnmp(host, community);
            //если запрос не валиден, то пишем об этом
            if (!snmp.Valid)
            {
                MessageBox.Show("Запрос не валиден", "Error");
                return;
            }
            //формируем тело запроса и отсылаем его
            Dictionary<Oid, AsnType> result = snmp.Get(SnmpVersion.Ver2, new[]
                {

                        ".1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.2", //Печать А4
                        ".1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.7.1.2", //Печать А3
                        ".1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.5.1.1", //Копии А4
                        ".1.3.6.1.4.1.18334.1.1.1.5.7.2.2.1.7.1.1", //Копии А3
           
		            });
            //если нет ответа, то так и скажем
            if (result == null)
            {
                MessageBox.Show("No response", "Error");
                return;
            }
            //счетчик для форматирования

            int i = 0;

            foreach (var kvp in result)
            {
                int n = 0;
                //если есть ответ, то выводим
                int m = 0;

                switch (i)
                {

                    case 0:
                        dataGridView3.Rows[0].Cells[i].Value = kvp.Value;
                        break;
                    case 1:
                        dataGridView3.Rows[0].Cells[i].Value = kvp.Value;
                        break;
                    case 2:
                        dataGridView3.Rows[0].Cells[i].Value = kvp.Value;
                        break;
                    case 3:
                        dataGridView3.Rows[0].Cells[i].Value = kvp.Value;
                        break;

                };
                dataGridView3.Rows[0].Cells[i].Value = kvp.Value;
                

                i++;
                n++;
                m++;
            }
        }

    
    }
}




