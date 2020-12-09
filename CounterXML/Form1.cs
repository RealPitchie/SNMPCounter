using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using SnmpSharpNet; 

namespace CounterXML
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        // Объявление переменных 
        OpenFileDialog OPF = new OpenFileDialog();
        OpenFileDialog OPF2 = new OpenFileDialog();
        Encoding enc = Encoding.GetEncoding(1251);
        

        private void button1_Click_1(object sender, EventArgs e)
        {
            {
                //к кому будем обращаться
                string host = "192.168.0.95"; 
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
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //к кому будем обращаться
            string host = "192.168.0.24"; 
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
             }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //к кому будем обращаться
            string host = "192.168.0.150"; 
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
            }
        }

    
    }
}
