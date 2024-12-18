﻿using System;
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
        inform_system_baseDataSetTableAdapters.QueriesTableAdapter queriesTableAdapter;
        public Form1()
        {
            InitializeComponent();
            queriesTableAdapter = new inform_system_baseDataSetTableAdapters.QueriesTableAdapter();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
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
    }
}
