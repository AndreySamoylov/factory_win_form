using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;


namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        inform_system_baseDataSetTableAdapters.QueriesTableAdapter queriesTableAdapter;
        public Form1()
        {
            InitializeComponent();
            queriesTableAdapter = new inform_system_baseDataSetTableAdapters.QueriesTableAdapter();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.ПредставлениеСырьё". При необходимости она может быть перемещена или удалена.
            this.представлениеСырьёTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеСырьё);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.ПредставлениеОплаты". При необходимости она может быть перемещена или удалена.
            this.представлениеОплатыTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеОплаты);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.ПредставлениеЗаказы". При необходимости она может быть перемещена или удалена.
            this.представлениеЗаказыTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеЗаказы);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.ПредставлениеДеталиКузовногоЦеха". При необходимости она может быть перемещена или удалена.
            this.представлениеДеталиКузовногоЦехаTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеДеталиКузовногоЦеха);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.ПредставлениеСклады". При необходимости она может быть перемещена или удалена.
            this.представлениеСкладыTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеСклады);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.ПредставлениеТовары". При необходимости она может быть перемещена или удалена.
            this.представлениеТоварыTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеТовары);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.Склады". При необходимости она может быть перемещена или удалена.
            this.складыTableAdapter.Fill(this.inform_system_baseDataSet.Склады);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.Товары". При необходимости она может быть перемещена или удалена.
            this.товарыTableAdapter.Fill(this.inform_system_baseDataSet.Товары);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.ПредставлениеБиблиотекаТипыКниг". При необходимости она может быть перемещена или удалена.
            this.представлениеБиблиотекаТипыКнигTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеБиблиотекаТипыКниг);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.ПредставлениеБиблиотекаМестаХранения". При необходимости она может быть перемещена или удалена.
            this.представлениеБиблиотекаМестаХраненияTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеБиблиотекаМестаХранения);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.ПредставлениеБиблиотекаКниги". При необходимости она может быть перемещена или удалена.
            this.представлениеБиблиотекаКнигиTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеБиблиотекаКниги);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.ПредставлениеПриемкаСырья". При необходимости она может быть перемещена или удалена.
            this.представлениеПриемкаСырьяTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеПриемкаСырья);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.ПредставлениеВыпуск_Деталей". При необходимости она может быть перемещена или удалена.
            this.представлениеВыпуск_ДеталейTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеВыпуск_Деталей);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.ПредставлениеКонтрагенты". При необходимости она может быть перемещена или удалена.
            // this.представлениеКонтрагентыBindingSource.Fill(this.inform_system_baseDataSet.ПредставлениеКонтрагенты);
            this.представлениеКонтрагентыTableAdapter1.Fill(this.inform_system_baseDataSet.ПредставлениеКонтрагенты);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.ПредставлениеЗакупки". При необходимости она может быть перемещена или удалена.
            this.представлениеЗакупкиTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеЗакупки);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.ПредставлениеОстатки_на_складах". При необходимости она может быть перемещена или удалена.
            this.представлениеОстатки_на_складахTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеОстатки_на_складах);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.ПредставлениеДоговоры_с_контрагентами". При необходимости она может быть перемещена или удалена.
            this.представлениеДоговоры_с_контрагентамиTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеДоговоры_с_контрагентами);
            // TODO: данная строка кода позволяет загрузить данные в таблицу "inform_system_baseDataSet.ПредставлениеЖалобы_от_клиентов". При необходимости она может быть перемещена или удалена.
            this.представлениеЖалобы_от_клиентовTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеЖалобы_от_клиентов);

        }

        private void toolStripButtonLibraryBookAdd_Click(object sender, EventArgs e)
        {
            String name = this.textBoxBookLibraryName.Text;
            String author = this.textBoxLibraryBookAuthor.Text;
            DateTime date = Convert.ToDateTime(this.dateTimePickerLibraryBook.Text);
            System.Data.DataRowView storage = (System.Data.DataRowView) this.comboBoxLibraryBookStorage.SelectedValue;
            Int32 id_storage = Convert.ToInt32(storage.Row[0]);
            System.Data.DataRowView bookType = (System.Data.DataRowView)this.comboBoxLubraryBookType.SelectedValue;
            Int32 id_book_type = Convert.ToInt32(bookType.Row[0]);

            queriesTableAdapter.CreateLibraryBook(
                name, author, date, id_book_type, id_storage
            );
            this.представлениеБиблиотекаКнигиTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеБиблиотекаКниги);
        }

        private void toolStripButtonLibraryBookUpdate_Click(object sender, EventArgs e)
        {
            String name = this.textBoxBookLibraryName.Text;
            String author = this.textBoxLibraryBookAuthor.Text;
            DateTime date = Convert.ToDateTime(this.dateTimePickerLibraryBook.Text);
            System.Data.DataRowView storage = (System.Data.DataRowView)this.comboBoxLibraryBookStorage.SelectedValue;
            Int32 id_storage = Convert.ToInt32(storage.Row[0]);
            System.Data.DataRowView bookType = (System.Data.DataRowView)this.comboBoxLubraryBookType.SelectedValue;
            Int32 id_book_type = Convert.ToInt32(bookType.Row[0]);

            int id = 0;
            DataRowView drv;
            drv = (DataRowView) представлениеБиблиотекаКнигиBindingSource.Current;
            id = (int)drv["код_книги"];
            queriesTableAdapter.UpdateLibraryBook(id, name, author, date, id_book_type, id_storage);
            this.представлениеБиблиотекаКнигиTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеБиблиотекаКниги);
        }

        private void toolStripButtonLibraryBookDelete_Click(object sender, EventArgs e)
        {
            DataRowView drv;
            int i = представлениеБиблиотекаКнигиBindingSource.Count;
            if (i > 0)
            {
                drv = (DataRowView) представлениеБиблиотекаКнигиBindingSource.Current;
                int id = (int)drv["код_книги"];
                queriesTableAdapter.DeleteLibraryBook(id);
                this.представлениеБиблиотекаКнигиTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеБиблиотекаКниги);
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            String name = this.textBox_book_storage.Text;

            queriesTableAdapter.CreateLibraryStorage(name);
            this.представлениеБиблиотекаМестаХраненияTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеБиблиотекаМестаХранения);
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            String name = this.textBox_book_storage.Text;

            int id = 0;
            DataRowView drv;
            drv = (DataRowView)представлениеБиблиотекаМестаХраненияBindingSource.Current;
            id = (int)drv["код_места_хранения"];
            queriesTableAdapter.UpdateLibraryStorage(id, name);

            this.представлениеБиблиотекаМестаХраненияTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеБиблиотекаМестаХранения);
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            DataRowView drv;
            int i = представлениеБиблиотекаМестаХраненияBindingSource.Count;
            if (i > 0)
            {
                drv = (DataRowView)представлениеБиблиотекаМестаХраненияBindingSource.Current;
                int id = (int)drv["код_места_хранения"];
                queriesTableAdapter.DeleteLibraryStorage(id);
                this.представлениеБиблиотекаМестаХраненияTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеБиблиотекаМестаХранения);
            }
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            String name = this.textBox_book_type.Text;

            queriesTableAdapter.CreateLibraryBookType(name);
            this.представлениеБиблиотекаТипыКнигTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеБиблиотекаТипыКниг);
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            String name = this.textBox_book_type.Text;


            int id = 0;
            DataRowView drv;
            drv = (DataRowView)представлениеБиблиотекаТипыКнигBindingSource.Current;
            id = (int)drv["код_типа_книг"];
            queriesTableAdapter.UpdateLibraryBookType(id, name);

            this.представлениеБиблиотекаТипыКнигTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеБиблиотекаТипыКниг);

        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            DataRowView drv;
            int i = представлениеБиблиотекаТипыКнигBindingSource.Count;
            if (i > 0)
            {
                drv = (DataRowView)представлениеБиблиотекаТипыКнигBindingSource.Current;
                int id = (int)drv["код_типа_книг"];
                queriesTableAdapter.DeleteLibraryBookType(id);
                this.представлениеБиблиотекаТипыКнигTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеБиблиотекаТипыКниг);
            }
        }

        private void bindingNavigatorAddNewComplaints_Click(object sender, EventArgs e)
        {
            System.Data.DataRowView contractorname = (System.Data.DataRowView)this.comboBoxContractorName.SelectedValue;
            Int32 id_contractor = Convert.ToInt32(contractorname.Row[0]);
            String complaint = this.textBoxComplaint.Text;
            System.Boolean isComplaintReviewed = this.checkBoxComplaintReviewed.Checked;
            DateTime date = Convert.ToDateTime(this.dateTimePickerComplaint.Text);
           
            queriesTableAdapter.CreateComplaint(
                 id_contractor, complaint, isComplaintReviewed, date
            );
            this.представлениеЖалобы_от_клиентовTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеЖалобы_от_клиентов);
        }

        private void bindingNavigatorDeleteComplaints_Click(object sender, EventArgs e)
        {
            DataRowView drv;
            int i = представлениеЖалобы_от_клиентовBindingSource.Count;
            if (i > 0)
            {
                drv = (DataRowView)представлениеЖалобы_от_клиентовBindingSource.Current;
                int id = (int)drv["Код"];
                queriesTableAdapter.DeleteComplaint(id);
                this.представлениеЖалобы_от_клиентовTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеЖалобы_от_клиентов);
            }

        }

        private void bindingNavigatorUpdateComplaints_Click(object sender, EventArgs e)
        {
            System.Data.DataRowView contractorname = (System.Data.DataRowView)this.comboBoxContractorName.SelectedValue;
            Int32 id_contractor = Convert.ToInt32(contractorname.Row[0]);
            String complaint = this.textBoxComplaint.Text;
            System.Boolean isComplaintReviewed = this.checkBoxComplaintReviewed.Checked;
            DateTime date = Convert.ToDateTime(this.dateTimePickerComplaint.Text);

            int id = 0;
            DataRowView drv;
            drv = (DataRowView)представлениеКонтрагентыBindingSource.Current;
            id = (int)drv["Код"];

            //// Обновляем данные через метод адаптера
            queriesTableAdapter.UpdateComplain(id, id_contractor, complaint, isComplaintReviewed, date);

            //// Перезагружаем данные для представления "Контрагенты"
            this.представлениеКонтрагентыTableAdapter1.Fill(this.inform_system_baseDataSet.ПредставлениеКонтрагенты);

            //// Перезагружаем данные для DataGridView "Жалобы от клиентов"
            this.представлениеЖалобы_от_клиентовTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеЖалобы_от_клиентов);

            //// Обновляем отображение данных в DataGridView
            this.представлениеЖалобы_от_клиентовDataGridView.Refresh();

        }

        private void bindingNavigatorAddNewItem1_Click(object sender, EventArgs e)
        {
           
        }

        private void toolStripButton8_Click(object sender, EventArgs e)
        {
            String contract_number = this.textBox2.Text;
            System.Boolean contract_signed = this.checkBox2.Checked;
            System.Data.DataRowView contractorname_1 = (System.Data.DataRowView)this.comboBox1.SelectedValue;
            Int32 id_contractor_1 = Convert.ToInt32(contractorname_1.Row[0]);

            int id = 0;
            DataRowView drv;
            drv = (DataRowView)представлениеДоговоры_с_контрагентамиBindingSource.Current;
            id = (int)drv["Код"];

            // Обновляем данные через метод адаптера
            queriesTableAdapter.UpdateContracts(id,id_contractor_1, contract_number, contract_signed);

            // Перезагружаем данные для представления "Контрагенты"
            this.представлениеКонтрагентыTableAdapter1.Fill(this.inform_system_baseDataSet.ПредставлениеКонтрагенты);

            // Перезагружаем данные для DataGridView "Договоры с контрагентами"
            this.представлениеДоговоры_с_контрагентамиTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеДоговоры_с_контрагентами);

            // Обновляем отображение данных в DataGridView
            this.представлениеДоговоры_с_контрагентамиDataGridView.Refresh();
        }

        private void bindingNavigatorDeleteItem1_Click(object sender, EventArgs e)
        {

        }

        private void bindingNavigatorAddNewItem2_Click(object sender, EventArgs e)
        {
            System.Data.DataRowView warehousename = (System.Data.DataRowView)this.comboBoxWarehouseName.SelectedValue;
            Int32 id_warehousename = Convert.ToInt32(warehousename.Row[0]);
            System.Data.DataRowView productname = (System.Data.DataRowView)this.comboBoxProductName.SelectedValue;
            Int32 id_productname = Convert.ToInt32(productname.Row[0]);
            Int32 quantitygoods = Convert.ToInt32(this.textBoxQuantityGoods.Text);
            queriesTableAdapter.CreateQuantity(
                 id_warehousename, id_productname, quantitygoods
            );
            this.представлениеОстатки_на_складахTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеОстатки_на_складах);
        }

        private void toolStripButton9_Click(object sender, EventArgs e)
        {

        }

        private void bindingNavigatorDeleteItem2_Click(object sender, EventArgs e)
        {

        }

        private void tabPage13_Click(object sender, EventArgs e)
        {
           
       
        }

        private void tabPage11_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            try
            {
                // Проверяем, что выбран подрядчик
                if (this.comboBoxContractorName.SelectedItem == null)
                {
                    MessageBox.Show("Пожалуйста, выберите контрагента.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.comboBoxContractorName.Focus();
                    return;
                }
                System.Data.DataRowView contractorname = (System.Data.DataRowView)this.comboBoxContractorName.SelectedItem;
                Int32 id_contractor = Convert.ToInt32(contractorname.Row[0]);

                // Получаем текст жалобы
                string complaint = this.textBoxComplaint.Text;
                if (string.IsNullOrWhiteSpace(complaint))
                {
                    MessageBox.Show("Пожалуйста, введите текст жалобы.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.textBoxComplaint.Focus();
                    return;
                }

                // Проверяем, выбран ли статус "Жалоба рассмотрена"
                bool isComplaintReviewed = this.checkBoxComplaintReviewed.Checked;

                // Проверяем корректность выбранной даты
                DateTime date;
                if (!DateTime.TryParse(this.dateTimePickerComplaint.Text, out date))
                {
                    MessageBox.Show("Пожалуйста, выберите корректную дату жалобы.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.dateTimePickerComplaint.Focus();
                    return;
                }

                // Выполняем добавление данных через метод адаптера
                queriesTableAdapter.CreateComplaint(id_contractor, complaint, isComplaintReviewed, date);

                // Сохраняем текущий фокус
                var previousControl = this.ActiveControl;

                // Обновляем данные для таблицы жалоб
                this.представлениеЖалобы_от_клиентовTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеЖалобы_от_клиентов);

                // Возвращаем фокус на тот элемент, который был активен до обновления
                previousControl?.Focus();
            }
            catch (Exception ex)
            {
                // Обработка исключений
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void toolStripButton10_Click(object sender, EventArgs e)
        {

            try
            {
                // Проверяем, что подрядчик выбран
                if (this.comboBoxContractorName.SelectedItem == null)
                {
                    MessageBox.Show("Пожалуйста, выберите подрядчика.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.comboBoxContractorName.Focus();
                    return;
                }
                System.Data.DataRowView contractorname = (System.Data.DataRowView)this.comboBoxContractorName.SelectedItem;
                Int32 id_contractor = Convert.ToInt32(contractorname.Row[0]);

                // Получаем текст жалобы
                string complaint = this.textBoxComplaint.Text;
                if (string.IsNullOrWhiteSpace(complaint))
                {
                    MessageBox.Show("Пожалуйста, введите текст жалобы.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.textBoxComplaint.Focus();
                    return;
                }

                // Получаем статус "Жалоба рассмотрена"
                bool isComplaintReviewed = this.checkBoxComplaintReviewed.Checked;

                // Проверяем корректность введенной даты
                DateTime date;
                if (!DateTime.TryParse(this.dateTimePickerComplaint.Text, out date))
                {
                    MessageBox.Show("Пожалуйста, выберите корректную дату жалобы.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.dateTimePickerComplaint.Focus();
                    return;
                }

                // Получаем текущую строку для обновления
                if (представлениеЖалобы_от_клиентовBindingSource.Current == null)
                {
                    MessageBox.Show("Нет выбранной строки для обновления.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DataRowView drv = (DataRowView)представлениеЖалобы_от_клиентовBindingSource.Current;
                int id = (int)drv["Код"]; // Получаем ID текущей строки

                // Обновляем данные через метод адаптера
                queriesTableAdapter.UpdateComplain(id, id_contractor, complaint, isComplaintReviewed, date);

                // Сохраняем текущий фокус
                var previousControl = this.ActiveControl;

                // Перезагружаем данные для представления "Жалобы от клиентов"
                this.представлениеЖалобы_от_клиентовTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеЖалобы_от_клиентов);

                // Возвращаем фокус на тот элемент, который был активен до обновления
                previousControl?.Focus();
            }
            catch (Exception ex)
            {
                // Обработка исключений
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void toolStripButton11_Click(object sender, EventArgs e)
        {
            DataRowView drv;
            int i = представлениеЖалобы_от_клиентовBindingSource.Count;
            if (i > 0)
            {
                drv = (DataRowView)представлениеЖалобы_от_клиентовBindingSource.Current;
                int id = (int)drv["Код"];
                queriesTableAdapter.DeleteComplaint(id);
                this.представлениеЖалобы_от_клиентовTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеЖалобы_от_клиентов);
            }
        }

        private void comboBoxContractorName_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void toolStripButton12_Click(object sender, EventArgs e)
        {
            try
            {
                // Проверяем, что номер контракта введен
                string contract_number = this.textBox2.Text;
                if (string.IsNullOrWhiteSpace(contract_number))
                {
                    MessageBox.Show("Пожалуйста, введите номер контракта.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.textBox2.Focus();
                    return;
                }

                // Проверяем, выбран ли контрагент
                if (this.comboBox1.SelectedItem == null)
                {
                    MessageBox.Show("Пожалуйста, выберите контрагент.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.comboBox1.Focus();
                    return;
                }
                System.Data.DataRowView contractorname_1 = (System.Data.DataRowView)this.comboBox1.SelectedItem;
                Int32 id_contractor_1 = Convert.ToInt32(contractorname_1.Row[0]);

                // Получаем статус подписания контракта
                bool contract_signed = this.checkBox2.Checked;

                // Выполняем добавление контракта через метод адаптера
                queriesTableAdapter.CreateContracts(id_contractor_1, contract_number, contract_signed);

                // Сохраняем текущий фокус
                var previousControl = this.ActiveControl;

                // Обновляем данные для таблицы "Договоры с контрагентами"
                this.представлениеДоговоры_с_контрагентамиTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеДоговоры_с_контрагентами);

                // Возвращаем фокус на тот элемент, который был активен до обновления
                previousControl?.Focus();
            }
            catch (Exception ex)
            {
                // Обработка исключений
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void toolStripButton13_Click(object sender, EventArgs e)
        {
            try
            {
                // Проверяем, что номер контракта введен
                string contract_number = this.textBox2.Text;
                if (string.IsNullOrWhiteSpace(contract_number))
                {
                    MessageBox.Show("Пожалуйста, введите номер контракта.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.textBox2.Focus();
                    return;
                }

                // Проверяем, что подрядчик выбран
                if (this.comboBox1.SelectedValue == null)
                {
                    MessageBox.Show("Пожалуйста, выберите подрядчика.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.comboBox1.Focus();
                    return;
                }
                System.Data.DataRowView contractorname_1 = (System.Data.DataRowView)this.comboBox1.SelectedItem;
                Int32 id_contractor_1 = Convert.ToInt32(contractorname_1.Row[0]);

                // Получаем значение из checkbox для состояния контракта (подписан/не подписан)
                System.Boolean contract_signed = this.checkBox2.Checked;

                // Проверяем, что есть выбранная строка для обновления
                if (представлениеДоговоры_с_контрагентамиBindingSource.Current == null)
                {
                    MessageBox.Show("Нет выбранной строки для обновления.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                DataRowView drv = (DataRowView)представлениеДоговоры_с_контрагентамиBindingSource.Current;
                int id = (int)drv["Код"]; // Получаем ID текущей строки для обновления

                // Выполняем обновление контракта через метод адаптера
                queriesTableAdapter.UpdateContracts(id, id_contractor_1, contract_number, contract_signed);

                // Сохраняем текущий фокус
                var previousControl = this.ActiveControl;

                // Перезагружаем данные для таблицы "Договоры с контрагентами"
                this.представлениеДоговоры_с_контрагентамиTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеДоговоры_с_контрагентами);

                // Возвращаем фокус на тот элемент, который был активен до обновления
                previousControl?.Focus();
            }
            catch (Exception ex)
            {
                // Обработка исключений
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }



        }

        private void toolStripButton14_Click(object sender, EventArgs e)
        {
            DataRowView drv;
            int i = представлениеДоговоры_с_контрагентамиBindingSource.Count;
            if (i > 0)
            {
                drv = (DataRowView)представлениеДоговоры_с_контрагентамиBindingSource.Current;
                int id = (int)drv["Код"];
                queriesTableAdapter.DeleteContracts(id);
                this.представлениеДоговоры_с_контрагентамиTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеДоговоры_с_контрагентами);
            }
        }

        private void toolStripButton8_Click_1(object sender, EventArgs e)
        {
            try
            {
                // Получаем данные из ComboBox складов
                if (this.comboBoxWarehouseName.SelectedItem == null)
                {
                    MessageBox.Show("Пожалуйста, выберите склад.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.comboBoxWarehouseName.Focus();
                    return;
                }
                System.Data.DataRowView warehousename = (System.Data.DataRowView)this.comboBoxWarehouseName.SelectedItem;
                Int32 id_warehousename = Convert.ToInt32(warehousename.Row[0]);

                // Получаем данные из ComboBox товаров
                if (this.comboBoxProductName.SelectedItem == null)
                {
                    MessageBox.Show("Пожалуйста, выберите товар.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.comboBoxProductName.Focus();
                    return;
                }
                System.Data.DataRowView productname = (System.Data.DataRowView)this.comboBoxProductName.SelectedItem;
                Int32 id_productname = Convert.ToInt32(productname.Row[0]);

                // Проверяем корректность ввода в поле количества
                if (!int.TryParse(this.textBoxQuantityGoods.Text, out Int32 quantitygoods) || quantitygoods <= 0)
                {
                    MessageBox.Show("Пожалуйста, введите корректное количество товара.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.textBoxQuantityGoods.Focus();
                    return;
                }

                // Выполняем добавление данных
                queriesTableAdapter.CreateQuantity(id_warehousename, id_productname, quantitygoods);

                // Обновляем данные в таблице
                var previousControl = this.ActiveControl; // Сохраняем текущий фокус
                this.представлениеОстатки_на_складахTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеОстатки_на_складах);
                previousControl?.Focus(); // Возвращаем фокус на предыдущий элемент
            }
            catch (Exception ex)
            {
                // Обработка исключений
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void toolStripButton15_Click(object sender, EventArgs e)
        {
            try
            {
                // Проверяем, что склад выбран
                if (this.comboBoxWarehouseName.SelectedValue == null)
                {
                    MessageBox.Show("Пожалуйста, выберите склад.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.comboBoxWarehouseName.Focus();
                    return;
                }
                System.Data.DataRowView warehousename = (System.Data.DataRowView)this.comboBoxWarehouseName.SelectedItem;
                Int32 id_warehousename = Convert.ToInt32(warehousename.Row[0]);

                // Проверяем, что товар выбран
                if (this.comboBoxProductName.SelectedValue == null)
                {
                    MessageBox.Show("Пожалуйста, выберите товар.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.comboBoxProductName.Focus();
                    return;
                }
                System.Data.DataRowView productname = (System.Data.DataRowView)this.comboBoxProductName.SelectedItem;
                Int32 id_productname = Convert.ToInt32(productname.Row[0]);

                // Проверяем корректность ввода количества товара
                if (!int.TryParse(this.textBoxQuantityGoods.Text, out Int32 quantitygoods) || quantitygoods <= 0)
                {
                    MessageBox.Show("Пожалуйста, введите корректное количество товара.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.textBoxQuantityGoods.Focus();
                    return;
                }

                // Получаем текущую строку для обновления
                if (представлениеОстатки_на_складахBindingSource.Current == null)
                {
                    MessageBox.Show("Нет выбранной строки для обновления.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DataRowView drv = (DataRowView)представлениеОстатки_на_складахBindingSource.Current;
                int id = (int)drv["Код"]; // Получаем ID текущей строки

                // Обновляем данные через метод адаптера
                queriesTableAdapter.UpdateQuantity(id, id_warehousename, id_productname, quantitygoods);

                // Сохраняем текущий фокус
                var previousControl = this.ActiveControl;

                // Перезагружаем данные для представления "Склады"
                this.представлениеОстатки_на_складахTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеОстатки_на_складах);

                // Возвращаем фокус на тот элемент, который был активен до обновления
                previousControl?.Focus();
            }
            catch (Exception ex)
            {
                // Обработка исключений
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void toolStripButton16_Click(object sender, EventArgs e)
        {
            DataRowView drv;
            int i = представлениеОстатки_на_складахBindingSource.Count;
            if (i > 0)
            {
                drv = (DataRowView)представлениеОстатки_на_складахBindingSource.Current;
                int id = (int)drv["Код"];
                queriesTableAdapter.DeleteQuantity(id);
                this.представлениеОстатки_на_складахTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеОстатки_на_складах);
            }
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void toolStripButton9_Click_1(object sender, EventArgs e)
        {

            String name_contractors = this.textBox1.Text;
            String inn = this.textBox3.Text;
            String email = this.textBox4.Text;
            String telephone = this.textBox5.Text;
            String adress = this.textBox6.Text;
            queriesTableAdapter.CreateContractors(
                 name_contractors, inn, email, telephone, adress
            );
            this.представлениеКонтрагентыTableAdapter1.Fill(this.inform_system_baseDataSet.ПредставлениеКонтрагенты);
        }

        private void toolStripButton17_Click(object sender, EventArgs e)
        {
            String name_contractors = this.textBox1.Text;
            String inn = this.textBox3.Text;
            String email = this.textBox4.Text;
            String telephone = this.textBox5.Text;
            String adress = this.textBox6.Text;

            int id = 0;
            DataRowView drv;
            drv = (DataRowView)представлениеКонтрагентыBindingSource.Current;
            id = (int)drv["Код"];

            // Обновляем данные через метод адаптера
            queriesTableAdapter.UpdateContractors(id, name_contractors, inn, email, telephone, adress);

            // Перезагружаем данные для DataGridView "Договоры с контрагентами"
            this.представлениеКонтрагентыTableAdapter1.Fill(this.inform_system_baseDataSet.ПредставлениеКонтрагенты);
        }

        private void toolStripButton18_Click(object sender, EventArgs e)
        {
            DataRowView drv;
            int i = представлениеКонтрагентыBindingSource.Count;
            if (i > 0)
            {
                drv = (DataRowView)представлениеКонтрагентыBindingSource.Current;
                int id = (int)drv["Код"];
                queriesTableAdapter.DeleteContractors(id);
                this.представлениеКонтрагентыTableAdapter1.Fill(this.inform_system_baseDataSet.ПредставлениеКонтрагенты);

            }
        }

        private void toolStripButton19_Click(object sender, EventArgs e)
        {
            System.Data.DataRowView product_name = (System.Data.DataRowView)this.comboBoxname_product.SelectedItem;
            Int32 idproductname = Convert.ToInt32(product_name.Row[0]);
            System.Data.DataRowView name_contragents = (System.Data.DataRowView)this.comboBoxname_contractor.SelectedItem;
            Int32 id_name_contragents = Convert.ToInt32(name_contragents.Row[0]);
            Int32 count = Convert.ToInt32(this.textBox9.Text);
            Decimal price = Convert.ToDecimal(this.textBox10.Text);
            //Double price = Convert.ToDouble(this.textBox10.Text);
            DateTime date = Convert.ToDateTime(this.dateTimePicker1.Text);
            queriesTableAdapter.CreatePurchases(
                 idproductname, id_name_contragents, count, price, date
            );
            this.представлениеЗакупкиTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеЗакупки);
        }

        private void toolStripButton20_Click(object sender, EventArgs e)
        {

            System.Data.DataRowView product_name = (System.Data.DataRowView)this.comboBoxname_product.SelectedItem;
            Int32 idproductname = Convert.ToInt32(product_name.Row[0]);
            System.Data.DataRowView name_contragents = (System.Data.DataRowView)this.comboBoxname_contractor.SelectedItem;
            Int32 id_name_contragents = Convert.ToInt32(name_contragents.Row[0]);
            Int32 count = Convert.ToInt32(this.textBox9.Text);
            Decimal price = Convert.ToDecimal(this.textBox10.Text);
            //Double price = Convert.ToDouble(this.textBox10.Text);
            DateTime date = Convert.ToDateTime(this.dateTimePicker1.Text);

            int id = 0;
            DataRowView drv;
            drv = (DataRowView)представлениеЗакупкиBindingSource.Current;
            id = (int)drv["Код"];

            // Обновляем данные через метод адаптера
            queriesTableAdapter.UpdatePurchases(id, idproductname, id_name_contragents, count, price, date);

            // Перезагружаем данные для представления "закупки"
            this.представлениеЗакупкиTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеЗакупки);
        }

        private void toolStripButton21_Click(object sender, EventArgs e)
        {
            DataRowView drv;
            int i = представлениеЗакупкиBindingSource.Count;
            if (i > 0)
            {
                drv = (DataRowView)представлениеЗакупкиBindingSource.Current;
                int id = (int)drv["Код"];
                queriesTableAdapter.DeletePurchases(id);
                this.представлениеЗакупкиTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеЗакупки);
            }
        }

        private void toolStripButton22_Click(object sender, EventArgs e)
        {
            
        }

        private void comboBoxWarehouseName_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void представлениеЖалобы_от_клиентовDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnExportToWord_Click(object sender, EventArgs e)
        {
           
        }

        private void представлениеДоговоры_с_контрагентамиDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

            // Создаем объект Word.Application
            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = wordApp.Documents.Add();
            try
            {
                // Добавление заголовка в документ
                Word.Paragraph paragraph = wordDoc.Content.Paragraphs.Add();
                paragraph.Range.Text = "Отчет по договорам с контрагентами";
                paragraph.Range.Font.Bold = 1;
                paragraph.Format.SpaceAfter = 10;
                paragraph.Range.InsertParagraphAfter();

                // Создание таблицы в Word
                Word.Table wordTable = wordDoc.Tables.Add(paragraph.Range, представлениеДоговоры_с_контрагентамиDataGridView.Rows.Count + 1, представлениеДоговоры_с_контрагентамиDataGridView.Columns.Count);
                wordTable.Borders.Enable = 1;

                // Добавление заголовков таблицы
                for (int i = 0; i < представлениеДоговоры_с_контрагентамиDataGridView.Columns.Count; i++)
                {
                    wordTable.Cell(1, i + 1).Range.Text = представлениеДоговоры_с_контрагентамиDataGridView.Columns[i].HeaderText;
                }

                // Добавление данных в таблицу
                for (int i = 0; i < представлениеДоговоры_с_контрагентамиDataGridView.Rows.Count; i++)
                {
                    for (int j = 0; j < представлениеДоговоры_с_контрагентамиDataGridView.Columns.Count; j++)
                    {
                        wordTable.Cell(i + 2, j + 1).Range.Text = представлениеДоговоры_с_контрагентамиDataGridView.Rows[i].Cells[j].Value?.ToString() ?? "";
                    }
                }

                // Сохранение документа
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Word Document (*.docx)|*.docx",
                    FileName = "Отчет_по_договорам"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    wordDoc.SaveAs2(saveFileDialog.FileName);
                    MessageBox.Show("Отчет успешно сохранен!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Word: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Закрываем Word
                wordDoc.Close(false);
                wordApp.Quit();
            }
        }

        private void toolStripButton22_Click_1(object sender, EventArgs e)
        {
            // Создаем объект Word.Application
            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = wordApp.Documents.Add();

            try
            {
                // Добавление заголовка в документ
                Word.Paragraph paragraph = wordDoc.Content.Paragraphs.Add();
                paragraph.Range.Text = "Отчет о договорах с контрагентами";
                paragraph.Range.Font.Bold = 1;
                paragraph.Format.SpaceAfter = 10;
                paragraph.Range.InsertParagraphAfter();

                // Вычисление количества видимых столбцов
                int visibleColumnCount = 0;
                foreach (DataGridViewColumn column in представлениеДоговоры_с_контрагентамиDataGridView.Columns)
                {
                    if (column.Visible)
                        visibleColumnCount++;
                }

                // Вычисление количества строк (исключая пустые строки для добавления)
                int rowCount = 0;
                foreach (DataGridViewRow row in представлениеДоговоры_с_контрагентамиDataGridView.Rows)
                {
                    if (!row.IsNewRow)
                        rowCount++;
                }

                // Создание таблицы в Word
                Word.Table wordTable = wordDoc.Tables.Add(paragraph.Range, rowCount + 1, visibleColumnCount);
                wordTable.Borders.Enable = 1;

                // Добавление заголовков таблицы
                int currentColumnIndex = 0;
                foreach (DataGridViewColumn column in представлениеДоговоры_с_контрагентамиDataGridView.Columns)
                {
                    if (column.Visible)
                    {
                        wordTable.Cell(1, currentColumnIndex + 1).Range.Text = column.HeaderText;
                        currentColumnIndex++;
                    }
                }

                // Добавление данных в таблицу
                int currentRowIndex = 0;
                foreach (DataGridViewRow row in представлениеДоговоры_с_контрагентамиDataGridView.Rows)
                {
                    if (row.IsNewRow) continue;

                    currentColumnIndex = 0;
                    for (int j = 0; j < представлениеДоговоры_с_контрагентамиDataGridView.Columns.Count; j++)
                    {
                        if (представлениеДоговоры_с_контрагентамиDataGridView.Columns[j].Visible)
                        {
                            var cellValue = row.Cells[j].Value;

                            // Проверка на столбец CheckBox и преобразование значений
                            if (представлениеДоговоры_с_контрагентамиDataGridView.Columns[j] is DataGridViewCheckBoxColumn)
                            {
                                wordTable.Cell(currentRowIndex + 2, currentColumnIndex + 1).Range.Text =
                                    cellValue is bool checkedValue && checkedValue ? "✔" : "✘";
                            }
                            else
                            {
                                wordTable.Cell(currentRowIndex + 2, currentColumnIndex + 1).Range.Text = cellValue?.ToString() ?? "";
                            }

                            currentColumnIndex++;
                        }
                    }

                    currentRowIndex++;
                }

                // Сохранение документа
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Word Document (*.docx)|*.docx",
                    FileName = "Отчет_о_договорах"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    wordDoc.SaveAs2(saveFileDialog.FileName);
                    MessageBox.Show("Отчет успешно сохранен!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Word: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Закрываем Word
                wordDoc.Close(false);
                wordApp.Quit();
            }
        }

        private void toolStripButton23_Click(object sender, EventArgs e)
        {
            // Создаем объект Word.Application
            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = wordApp.Documents.Add();
            try
            {
                // Добавление заголовка в документ
                Word.Paragraph paragraph = wordDoc.Content.Paragraphs.Add();
                paragraph.Range.Text = "Отчет о жалобах от клиентов";
                paragraph.Range.Font.Bold = 1;
                paragraph.Format.SpaceAfter = 10;
                paragraph.Range.InsertParagraphAfter();

                // Получаем только видимые столбцы
                var visibleColumns = представлениеЖалобы_от_клиентовDataGridView.Columns.Cast<DataGridViewColumn>()
                    .Where(col => col.Visible).ToList();

                // Подсчет строк, исключая последнюю пустую строку
                int rowCount = представлениеЖалобы_от_клиентовDataGridView.Rows
                    .Cast<DataGridViewRow>()
                    .Count(row => !row.IsNewRow);

                // Создание таблицы в Word
                Word.Table wordTable = wordDoc.Tables.Add(paragraph.Range, rowCount + 1, visibleColumns.Count);
                wordTable.Borders.Enable = 1;

                // Добавление заголовков таблицы
                for (int i = 0; i < visibleColumns.Count; i++)
                {
                    wordTable.Cell(1, i + 1).Range.Text = visibleColumns[i].HeaderText;
                }

                // Добавление данных в таблицу
                int currentRowIndex = 0;
                foreach (DataGridViewRow row in представлениеЖалобы_от_клиентовDataGridView.Rows)
                {
                    if (row.IsNewRow) continue; // Пропускаем пустую строку

                    for (int j = 0; j < visibleColumns.Count; j++)
                    {
                        var cellValue = row.Cells[visibleColumns[j].Index].Value;

                        if (cellValue is bool checkboxValue) // Если это чекбокс
                        {
                            wordTable.Cell(currentRowIndex + 2, j + 1).Range.Text = checkboxValue ? "✔" : "✘";
                        }
                        else
                        {
                            wordTable.Cell(currentRowIndex + 2, j + 1).Range.Text = cellValue?.ToString() ?? "";
                        }
                    }

                    currentRowIndex++;
                }

                // Сохранение документа
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Word Document (*.docx)|*.docx",
                    FileName = "Отчет_о_жалобах"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    wordDoc.SaveAs2(saveFileDialog.FileName);
                    MessageBox.Show("Отчет успешно сохранен!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Word: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Закрываем Word
                wordDoc.Close(false);
                wordApp.Quit();
            }
        }

        private void представлениеОстатки_на_складахDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void toolStripButton24_Click(object sender, EventArgs e)
        {
            // Создаем объект Word.Application
            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = wordApp.Documents.Add();
            try
            {
                // Добавление заголовка в документ
                Word.Paragraph paragraph = wordDoc.Content.Paragraphs.Add();
                paragraph.Range.Text = "Отчет об остатках на складах";
                paragraph.Range.Font.Bold = 1;
                paragraph.Format.SpaceAfter = 10;
                paragraph.Range.InsertParagraphAfter();

                // Получаем только видимые столбцы
                var visibleColumns = представлениеОстатки_на_складахDataGridView.Columns.Cast<DataGridViewColumn>()
                    .Where(col => col.Visible).ToList();

                // Подсчет строк, исключая последнюю пустую строку
                int rowCount = представлениеОстатки_на_складахDataGridView.Rows
                    .Cast<DataGridViewRow>()
                    .Count(row => !row.IsNewRow);

                // Создание таблицы в Word
                Word.Table wordTable = wordDoc.Tables.Add(paragraph.Range, rowCount + 1, visibleColumns.Count);
                wordTable.Borders.Enable = 1;

                // Добавление заголовков таблицы
                for (int i = 0; i < visibleColumns.Count; i++)
                {
                    wordTable.Cell(1, i + 1).Range.Text = visibleColumns[i].HeaderText;
                }

                // Добавление данных в таблицу
                int currentRowIndex = 0;
                foreach (DataGridViewRow row in представлениеОстатки_на_складахDataGridView.Rows)
                {
                    if (row.IsNewRow) continue; // Пропускаем пустую строку

                    for (int j = 0; j < visibleColumns.Count; j++)
                    {
                        var cellValue = row.Cells[visibleColumns[j].Index].Value;
                        wordTable.Cell(currentRowIndex + 2, j + 1).Range.Text = cellValue?.ToString() ?? "";
                    }

                    currentRowIndex++;
                }

                // Сохранение документа
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Word Document (*.docx)|*.docx",
                    FileName = "Отчет_об_остатках_на_складах"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    wordDoc.SaveAs2(saveFileDialog.FileName);
                    MessageBox.Show("Отчет успешно сохранен!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Word: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Закрываем Word
                wordDoc.Close(false);
                wordApp.Quit();
            }

        }

        private void toolStripButton25_Click(object sender, EventArgs e)
        {
            // Создаем объект Word.Application
            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = wordApp.Documents.Add();
            try
            {
                // Добавление заголовка в документ
                Word.Paragraph paragraph = wordDoc.Content.Paragraphs.Add();
                paragraph.Range.Text = "Отчет по контрагентам";
                paragraph.Range.Font.Bold = 1;
                paragraph.Format.SpaceAfter = 10;
                paragraph.Range.InsertParagraphAfter();

                // Получаем только видимые столбцы, кроме столбца "Код"
                var visibleColumns = представлениеКонтрагентыDataGridView.Columns.Cast<DataGridViewColumn>()
                    .Where(col => col.Visible && col.HeaderText != "Код").ToList();

                // Проверка: если видимых столбцов нет
                if (visibleColumns.Count == 0)
                {
                    MessageBox.Show("Нет видимых столбцов для отображения в отчете.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Подсчет строк, исключая последнюю пустую строку
                var nonEmptyRows = представлениеКонтрагентыDataGridView.Rows.Cast<DataGridViewRow>()
                    .Where(row => !row.IsNewRow).ToList();

                // Создание таблицы в Word
                Word.Table wordTable = wordDoc.Tables.Add(paragraph.Range, nonEmptyRows.Count + 1, visibleColumns.Count);
                wordTable.Borders.Enable = 1;

                // Добавление заголовков таблицы
                for (int i = 0; i < visibleColumns.Count; i++)
                {
                    wordTable.Cell(1, i + 1).Range.Text = visibleColumns[i].HeaderText;
                }

                // Добавление данных в таблицу
                for (int rowIndex = 0; rowIndex < nonEmptyRows.Count; rowIndex++)
                {
                    for (int colIndex = 0; colIndex < visibleColumns.Count; colIndex++)
                    {
                        var cellValue = nonEmptyRows[rowIndex].Cells[visibleColumns[colIndex].Index].Value;
                        wordTable.Cell(rowIndex + 2, colIndex + 1).Range.Text = cellValue?.ToString() ?? "";
                    }
                }

                // Сохранение документа
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Word Document (*.docx)|*.docx",
                    FileName = "Отчет_по_контрагентам"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    wordDoc.SaveAs2(saveFileDialog.FileName);
                    MessageBox.Show("Отчет успешно сохранен!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Word: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Закрываем Word
                wordDoc.Close(false);
                wordApp.Quit();
            }


        }

        private void представлениеКонтрагентыDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void представлениеЗакупкиDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void toolStripButton26_Click(object sender, EventArgs e)
        {
            // Создаем объект Word.Application
            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = wordApp.Documents.Add();
            try
            {
                // Добавление заголовка в документ
                Word.Paragraph paragraph = wordDoc.Content.Paragraphs.Add();
                paragraph.Range.Text = "Отчет по закупкам";
                paragraph.Range.Font.Bold = 1;
                paragraph.Format.SpaceAfter = 10;
                paragraph.Range.InsertParagraphAfter();

                // Получаем только видимые столбцы, кроме столбца "Код"
                var visibleColumns = представлениеЗакупкиDataGridView.Columns.Cast<DataGridViewColumn>()
                    .Where(col => col.Visible && col.HeaderText != "Код").ToList();

                // Проверка: если видимых столбцов нет
                if (visibleColumns.Count == 0)
                {
                    MessageBox.Show("Нет видимых столбцов для отображения в отчете.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Подсчет строк, исключая последнюю пустую строку
                var nonEmptyRows = представлениеЗакупкиDataGridView.Rows.Cast<DataGridViewRow>()
                    .Where(row => !row.IsNewRow).ToList();

                // Создание таблицы в Word
                Word.Table wordTable = wordDoc.Tables.Add(paragraph.Range, nonEmptyRows.Count + 1, visibleColumns.Count);
                wordTable.Borders.Enable = 1;

                // Добавление заголовков таблицы
                for (int i = 0; i < visibleColumns.Count; i++)
                {
                    wordTable.Cell(1, i + 1).Range.Text = visibleColumns[i].HeaderText;
                }

                // Добавление данных в таблицу
                for (int rowIndex = 0; rowIndex < nonEmptyRows.Count; rowIndex++)
                {
                    for (int colIndex = 0; colIndex < visibleColumns.Count; colIndex++)
                    {
                        var cellValue = nonEmptyRows[rowIndex].Cells[visibleColumns[colIndex].Index].Value;
                        wordTable.Cell(rowIndex + 2, colIndex + 1).Range.Text = cellValue?.ToString() ?? "";
                    }
                }

                // Сохранение документа
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Word Document (*.docx)|*.docx",
                    FileName = "Отчет_по_закупкам"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    wordDoc.SaveAs2(saveFileDialog.FileName);
                    MessageBox.Show("Отчет успешно сохранен!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Word: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Закрываем Word
                wordDoc.Close(false);
                wordApp.Quit();
            }

        }

        private void toolStripButtonKuzovChexAdd_Click(object sender, EventArgs e)
        {
            String name = this.textBoxKuzovChexName.Text;

            queriesTableAdapter.CreateDetail(name);

            this.представлениеДеталиКузовногоЦехаTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеДеталиКузовногоЦеха);
        }

        private void toolStripButtonKuzovChexUpdate_Click(object sender, EventArgs e)
        {
            String name = this.textBoxKuzovChexName.Text;

            int id = 0;
            DataRowView drv;
            drv = (DataRowView) this.представлениеДеталиКузовногоЦехаBindingSource.Current;
            id = (int)drv["Код детали"];
            queriesTableAdapter.UpdateDetail(id, name);
            this.представлениеДеталиКузовногоЦехаTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеДеталиКузовногоЦеха);
        }

        private void toolStripButtonKuzovChexDelete_Click(object sender, EventArgs e)
        {
            DataRowView drv;
            int i = this.представлениеДеталиКузовногоЦехаBindingSource.Count;
            if (i > 0)
            {
                drv = (DataRowView)this.представлениеДеталиКузовногоЦехаBindingSource.Current;
                int id = (int)drv["Код детали"];
                queriesTableAdapter.DeleteDetail(id);
                this.представлениеДеталиКузовногоЦехаTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеДеталиКузовногоЦеха);
            }
        }

        private void toolStripButtonDetailOutputAdd_Click(object sender, EventArgs e)
        {
            System.Data.DataRowView detail = (System.Data.DataRowView)this.comboBoxDetailOutputName.SelectedValue;
            Int32 id_detail = Convert.ToInt32(detail.Row[0]);
            DateTime date = Convert.ToDateTime(this.dateTimePickerDetailOutputDate.Text);

            queriesTableAdapter.CreateDetailOutput(id_detail, date);

            this.представлениеВыпуск_ДеталейTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеВыпуск_Деталей);
        }

        private void toolStripButtonDetailOutputUpdate_Click(object sender, EventArgs e)
        {
            System.Data.DataRowView detail = (System.Data.DataRowView)this.comboBoxDetailOutputName.SelectedValue;
            Int32 id_detail = Convert.ToInt32(detail.Row[0]);
            DateTime date = Convert.ToDateTime(this.dateTimePickerDetailOutputDate.Text);

            int id = 0;
            DataRowView drv;
            drv = (DataRowView)представлениеВыпуск_ДеталейBindingSource.Current;
            id = (int) drv["Код"];
            queriesTableAdapter.UpdateDetaiOutput(id, id_detail, date);

            this.представлениеВыпуск_ДеталейTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеВыпуск_Деталей);
        }

        private void toolStripButtonDetailOutputDelete_Click(object sender, EventArgs e)
        {
            DataRowView drv;
            int i = представлениеВыпуск_ДеталейBindingSource.Count;
            if (i > 0)
            {
                drv = (DataRowView)представлениеВыпуск_ДеталейBindingSource.Current;
                int id = (int)drv["Код"];
                queriesTableAdapter.DeleteDetaiOutput(id);
                this.представлениеВыпуск_ДеталейTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеВыпуск_Деталей);
            }
        }

        private void toolStripButtonPaymentAdd_Click(object sender, EventArgs e)
        {
            Int32 sum= Convert.ToInt32(this.textBoxPaymentSum.Text);
            DateTime date = Convert.ToDateTime(this.dateTimePickerPayment.Text);
                
            queriesTableAdapter.CreatePayment(sum, date);

            this.представлениеОплатыTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеОплаты);
        }

        private void toolStripButtonPaymentUpdate_Click(object sender, EventArgs e)
        {
            Int32 sum = Convert.ToInt32(this.textBoxPaymentSum.Text);
            DateTime date = Convert.ToDateTime(this.dateTimePickerPayment.Text);

            int id = 0;
            DataRowView drv;
            drv = (DataRowView)представлениеОплатыBindingSource.Current;
            id = (int)drv["код_оплаты"];
            queriesTableAdapter.UpdatePayment(id, sum, date);

            this.представлениеОплатыTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеОплаты);
        }

        private void toolStripButtonPaymentDelete_Click(object sender, EventArgs e)
        {
            DataRowView drv;
            int i = представлениеОплатыBindingSource.Count;
            if (i > 0)
            {
                drv = (DataRowView)представлениеОплатыBindingSource.Current;
                int id = (int)drv["код_оплаты"];
                queriesTableAdapter.DeletePayment(id);
                this.представлениеОплатыTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеОплаты);
            }
        }

        private void toolStripButtonOrdersAdd_Click(object sender, EventArgs e)
        {
            System.Data.DataRowView payment = (System.Data.DataRowView)this.comboBoxOrderPayment.SelectedValue;
            Int32 id_payment = Convert.ToInt32(payment.Row[0]);
            System.Data.DataRowView contagent = (System.Data.DataRowView)this.comboBoxOrderContragent.SelectedValue;
            Int32 id_contagent = Convert.ToInt32(contagent.Row[0]);

            queriesTableAdapter.CreateOrder(id_contagent, id_payment);

            this.представлениеЗаказыTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеЗаказы);

        }

        private void toolStripButtonOrdersUpdate_Click(object sender, EventArgs e)
        {
            System.Data.DataRowView payment = (System.Data.DataRowView)this.comboBoxOrderPayment.SelectedValue;
            Int32 id_payment = Convert.ToInt32(payment.Row[0]);
            System.Data.DataRowView contagent = (System.Data.DataRowView)this.comboBoxOrderContragent.SelectedValue;
            Int32 id_contagent = Convert.ToInt32(contagent.Row[0]);

            int id = 0;
            DataRowView drv;
            drv = (DataRowView)представлениеЗаказыBindingSource.Current;
            id = (int)drv["код_заказа"];
            queriesTableAdapter.UpdateOrder(id, id_contagent, id_payment);

            this.представлениеЗаказыTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеЗаказы);

        }

        private void toolStripButtonOrdersDelete_Click(object sender, EventArgs e)
        {
            DataRowView drv;
            int i = представлениеЗаказыBindingSource.Count;
            if (i > 0)
            {
                drv = (DataRowView)представлениеЗаказыBindingSource.Current;
                int id = (int)drv["код_заказа"];
                queriesTableAdapter.DeleteOrder(id);
                this.представлениеЗаказыTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеЗаказы);
            }
        }

        private void toolStripButtonRawAdd_Click(object sender, EventArgs e)
        {
            String name = this.textBoxRawName.Text;

            queriesTableAdapter.CreateRaw(name);

            this.представлениеСырьёTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеСырьё);

        }

        private void toolStripButtonRawUpdate_Click(object sender, EventArgs e)
        {
            String name = this.textBoxRawName.Text;

            int id = 0;
            DataRowView drv;
            drv = (DataRowView)представлениеСырьёBindingSource.Current;
            id = (int) drv["код_сырья"];

            queriesTableAdapter.UpdateRaw(id, name);

            this.представлениеСырьёTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеСырьё);

        }

        private void toolStripButtonRawDelete_Click(object sender, EventArgs e)
        {
            DataRowView drv;
            int i = представлениеСырьёBindingSource.Count;
            if (i > 0)
            {
                drv = (DataRowView)представлениеСырьёBindingSource.Current;
                int id = (int)drv["код_сырья"];
                queriesTableAdapter.DeleteRaw(id);
                this.представлениеСырьёTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеСырьё);
            }
        }

        private void toolStripButtonRawAccecingAdd_Click(object sender, EventArgs e)
        {
            System.Data.DataRowView raw = (System.Data.DataRowView)this.comboBoxRawAccecingRaw.SelectedValue;
            Int32 id_raw = Convert.ToInt32(raw.Row[0]);
            DateTime date = Convert.ToDateTime(this.dateTimePickerRawAccecing.Text);
            Int32 quantity = Convert.ToInt32(this.textBoxRawAccecingQuantity.Text);


            queriesTableAdapter.CreateRawAccepting(date, id_raw, quantity);

            this.представлениеПриемкаСырьяTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеПриемкаСырья);
        }

        private void toolStripButtonRawAccecingUpdate_Click(object sender, EventArgs e)
        {
            System.Data.DataRowView raw = (System.Data.DataRowView)this.comboBoxRawAccecingRaw.SelectedValue;
            Int32 id_raw = Convert.ToInt32(raw.Row[0]);
            DateTime date = Convert.ToDateTime(this.dateTimePickerRawAccecing.Text);
            Int32 quantity = Convert.ToInt32(this.textBoxRawAccecingQuantity.Text);

            int id = 0;
            DataRowView drv;
            drv = (DataRowView)представлениеПриемкаСырьяBindingSource.Current;
            id = (int)drv["код_приёмки"];

            queriesTableAdapter.UpdateRawAccepting(id, date, id_raw, quantity);

            this.представлениеПриемкаСырьяTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеПриемкаСырья);

        }

        private void toolStripButtonRawAccecingDelete_Click(object sender, EventArgs e)
        {

            DataRowView drv;
            int i = представлениеПриемкаСырьяBindingSource.Count;
            if (i > 0)
            {
                drv = (DataRowView)представлениеПриемкаСырьяBindingSource.Current;
                int id = (int)drv["код_приёмки"];
                queriesTableAdapter.DeleteRawAccepting(id);
                this.представлениеПриемкаСырьяTableAdapter.Fill(this.inform_system_baseDataSet.ПредставлениеПриемкаСырья);
            }
        }

        private void dateTimePickerDetailOutputDate_ValueChanged(object sender, EventArgs e)
        {

        }

        private void toolStripButton31_Click(object sender, EventArgs e)
        {
            // Функциональность создания отчета по деталям
            // Создаем объект Word.Application
            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = wordApp.Documents.Add();
            try
            {
                // Добавление заголовка в документ
                Word.Paragraph paragraph = wordDoc.Content.Paragraphs.Add();
                paragraph.Range.Text = "Отчет по деталям";
                paragraph.Range.Font.Bold = 1;
                paragraph.Format.SpaceAfter = 10;
                paragraph.Range.InsertParagraphAfter();

                // Получаем данные из TableAdapter

                var dataTable = inform_system_baseDataSet.ПредставлениеДеталиКузовногоЦеха; 

                // Проверка: если таблица пуста
                if (dataTable.Rows.Count == 0)
                {
                    MessageBox.Show("Нет данных для отображения в отчете.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Создание таблицы в Word
                Word.Table wordTable = wordDoc.Tables.Add(paragraph.Range, dataTable.Rows.Count + 1, dataTable.Columns.Count);
                wordTable.Borders.Enable = 1;

                // Добавление заголовков таблицы
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    wordTable.Cell(1, i + 1).Range.Text = dataTable.Columns[i].ColumnName;
                    wordTable.Cell(1, i + 1).Range.Font.Bold = 1; // Сделать заголовки жирными
                }

                // Добавление данных в таблицу
                for (int rowIndex = 0; rowIndex < dataTable.Rows.Count; rowIndex++)
                {
                    for (int colIndex = 0; colIndex < dataTable.Columns.Count; colIndex++)
                    {
                        var cellValue = dataTable.Rows[rowIndex][colIndex];
                        wordTable.Cell(rowIndex + 2, colIndex + 1).Range.Text = cellValue?.ToString() ?? "";
                    }
                }

                // Сохранение документа
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Word Document (*.docx)|*.docx",
                    FileName = "Отчет_по_деталям"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    wordDoc.SaveAs2(saveFileDialog.FileName);
                    MessageBox.Show("Отчет успешно сохранен!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Word: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Закрываем Word
                if (wordDoc != null)
                {
                    wordDoc.Close(false);
                    Marshal.ReleaseComObject(wordDoc); // Освобождаем объект документа
                }

                if (wordApp != null)
                {
                    wordApp.Quit();
                    Marshal.ReleaseComObject(wordApp); // Освобождаем объект приложения
                }
            }
        }

        private void toolStripButton32_Click(object sender, EventArgs e)
        {
            // Функциональность создания отчета по выпуску деталей
            // Создаем объект Word.Application
            Word.Application wordApp = new Word.Application();
            Word.Document wordDoc = wordApp.Documents.Add();
            try
            {
                // Добавление заголовка в документ
                Word.Paragraph paragraph = wordDoc.Content.Paragraphs.Add();
                paragraph.Range.Text = "Отчет по выпуску деталей";
                paragraph.Range.Font.Bold = 1;
                paragraph.Format.SpaceAfter = 10;
                paragraph.Range.InsertParagraphAfter();

                // Получаем данные из TableAdapter

                var dataTable = inform_system_baseDataSet.ПредставлениеВыпуск_Деталей;

                // Проверка: если таблица пуста
                if (dataTable.Rows.Count == 0)
                {
                    MessageBox.Show("Нет данных для отображения в отчете.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Создание таблицы в Word
                Word.Table wordTable = wordDoc.Tables.Add(paragraph.Range, dataTable.Rows.Count + 1, dataTable.Columns.Count);
                wordTable.Borders.Enable = 1;

                // Добавление заголовков таблицы
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    wordTable.Cell(1, i + 1).Range.Text = dataTable.Columns[i].ColumnName;
                    wordTable.Cell(1, i + 1).Range.Font.Bold = 1; // Сделать заголовки жирными
                }

                // Добавление данных в таблицу
                for (int rowIndex = 0; rowIndex < dataTable.Rows.Count; rowIndex++)
                {
                    for (int colIndex = 0; colIndex < dataTable.Columns.Count; colIndex++)
                    {
                        var cellValue = dataTable.Rows[rowIndex][colIndex];
                        wordTable.Cell(rowIndex + 2, colIndex + 1).Range.Text = cellValue?.ToString() ?? "";
                    }
                }

                // Сохранение документа
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Word Document (*.docx)|*.docx",
                    FileName = "Отчет_по_выпуску_деталей"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    wordDoc.SaveAs2(saveFileDialog.FileName);
                    MessageBox.Show("Отчет успешно сохранен!", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Word: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Закрываем Word
                if (wordDoc != null)
                {
                    wordDoc.Close(false);
                    Marshal.ReleaseComObject(wordDoc); // Освобождаем объект документа
                }

                if (wordApp != null)
                {
                    wordApp.Quit();
                    Marshal.ReleaseComObject(wordApp); // Освобождаем объект приложения
                }
            }
        }
    }
}
