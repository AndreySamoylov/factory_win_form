using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.ПредставлениеПриемкаСырья". При необходимости она может быть перемещена или удалена.
            this.представлениеПриемкаСырьяTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеПриемкаСырья);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.ПредставлениеВыпуск_Деталей". При необходимости она может быть перемещена или удалена.
            this.представлениеВыпуск_ДеталейTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеВыпуск_Деталей);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.ПредставлениеКонтрагенты". При необходимости она может быть перемещена или удалена.
            this.представлениеКонтрагентыTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеКонтрагенты);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.ПредставлениеЗакупки". При необходимости она может быть перемещена или удалена.
            this.представлениеЗакупкиTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеЗакупки);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.ПредставлениеОстатки_на_складах". При необходимости она может быть перемещена или удалена.
            this.представлениеОстатки_на_складахTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеОстатки_на_складах);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.ПредставлениеДоговоры_с_контрагентами". При необходимости она может быть перемещена или удалена.
            this.представлениеДоговоры_с_контрагентамиTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеДоговоры_с_контрагентами);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.ПредставлениеЖалобы_от_клиентов". При необходимости она может быть перемещена или удалена.
            this.представлениеЖалобы_от_клиентовTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеЖалобы_от_клиентов);

        }
    }
}
