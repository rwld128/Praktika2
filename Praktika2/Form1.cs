﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Praktika2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            tableLayoutPanel3.Visible = false;

            AktTable.Visible = false;
            PeremZakTable.Visible = false; 
            PerevodiTable.Visible = false;
            SmenaFamiliiTable.Visible = false;
            IzmCedTable.Visible = false;
            IzmStoimTable.Visible = false;
            MatKapitalTable.Visible = false;   
            DogovorTable.Visible = false;
            DopSoglTable.Visible = false;
        }

        private string ChangeName(string fullname)
        {
            string[] component = fullname.Split(' ');
            if (component.Length >= 3)
            {
                string sname = component[0];
                string fname = component[1];
                string dname = component[2];
                string changedname = sname + " " + fname[0] + "." + dname[0] + ".";
                return changedname;
            }
            return fullname;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                tableLayoutPanel3.Visible = true;
                tableLayoutPanel2.Visible = false;
            }
            else
            {
                tableLayoutPanel2.Visible = true;
                tableLayoutPanel3.Visible = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var origin = new Dictionary<string, string>
            {
                {"<DOG_NUM>", textBoxDOG_NUM.Text},
                {"<DOG_DATE>", dateTimePickerDOG_DATE.Value.ToString("dd.MM.yyyy")},
                {"<SGA_NUM>", textBoxSGA_NUM.Text},
                {"<SGA_DATE>", dateTimePickerSGA_DATE.Value.ToString("dd.MM.yyyy")},
                {"<SGA_UNTIL>", dateTimePickerSGA_UNTIL.Value.ToString("dd.MM.yyyy")},
                {"<DOV_DATE>", dateTimePickerDOV_DATE.Value.ToString("dd.MM.yyyy")},
                {"<DOV_NUM>", textBoxDOV_NUM.Text},
                {"<STUDENT_FIO>", textBoxSTUDENT_FIO.Text},
                {"<STUDENT_ADRES>", textBoxSTUDENT_ADRES.Text},
                {"<STUDENT_PHONE>", textBoxSTUDENT_PHONE.Text},
                {"<STUDENT_EMAIL>", textBoxSTUDENT_EMAIL.Text},
                {"<YUR_ZAK_FIO>", textBoxYUR_ZAK_FIO.Text},
                {"<YUR_ORG>", textBoxYUR_ORG.Text},
                {"<YUR_DOC>", textBoxYUR_DOC.Text},
                {"<YUR_ADRES>", textBoxYUR_ADRES.Text},
                {"<YUR_PHONE>", textBoxYUR_PHONE.Text},
                {"<YUR_BANK>", textBoxYUR_BANK.Text},
                {"<YUR_ZAK_PHONE>", textBoxYUR_ZAK_PHONE.Text},
                {"<YUR_ZAK_EMAIL>", textBoxYUR_ZAK_EMAIL.Text},
                {"<ZAK_FIO>", textBoxZAK_FIO.Text},
                {"<ZAK_ADRES>", textBoxZAK_ADRES.Text},
                {"<ZAK_INN>", textBoxZAK_INN.Text},
                {"<ZAK_PASP_SER>", textBoxZAK_PASP_SER.Text},
                {"<ZAK_PASP_NOM>", textBoxZAK_PASP_NOM.Text},
                {"<ZAK_PASP_VID>", textBoxZAK_PASP_VID.Text},
                {"<ZAK_PHONE>", textBoxZAK_PHONE.Text},
                {"<ZAK_EMAIL>", textBoxZAK_EMAIL.Text}
            };



            switch (comboBox1.SelectedItem.ToString())
            {
                case ("Акт об оказании услуг"):
                    AktTable.Visible = true;
                    PeremZakTable.Visible = false;
                    PerevodiTable.Visible = false;
                    SmenaFamiliiTable.Visible = false;
                    IzmCedTable.Visible = false;
                    IzmStoimTable.Visible = false;
                    MatKapitalTable.Visible = false;
                    DogovorTable.Visible = false;
                    DopSoglTable.Visible = false;
                    break;
                case ("Договор"):
                    AktTable.Visible = false;
                    PeremZakTable.Visible = false;
                    PerevodiTable.Visible = false;
                    SmenaFamiliiTable.Visible = false;
                    IzmCedTable.Visible = false;
                    IzmStoimTable.Visible = false;
                    MatKapitalTable.Visible = false;
                    DogovorTable.Visible = true;
                    DopSoglTable.Visible = false;
                    break;
                case ("Доп. соглашение"):
                    AktTable.Visible = false;
                    PeremZakTable.Visible = false;
                    PerevodiTable.Visible = false;
                    SmenaFamiliiTable.Visible = false;
                    IzmCedTable.Visible = false;
                    IzmStoimTable.Visible = false;
                    MatKapitalTable.Visible = false;
                    DogovorTable.Visible = false;
                    DopSoglTable.Visible = true;
                    break;
                case ("Перемена заказчика"):
                    AktTable.Visible = false;
                    PeremZakTable.Visible = true;
                    PerevodiTable.Visible = false;
                    SmenaFamiliiTable.Visible = false;
                    IzmCedTable.Visible = false;
                    IzmStoimTable.Visible = false;
                    MatKapitalTable.Visible = false;
                    DogovorTable.Visible = false;
                    DopSoglTable.Visible = false;
                    break;
                case ("Смена фамилии"):
                    AktTable.Visible = false;
                    PeremZakTable.Visible = false;
                    PerevodiTable.Visible = false;
                    SmenaFamiliiTable.Visible = true;
                    IzmCedTable.Visible = false;
                    IzmStoimTable.Visible = false;
                    MatKapitalTable.Visible = false;
                    DogovorTable.Visible = false;
                    DopSoglTable.Visible = false;
                    break;
                case ("Изменение стоимости"):
                    AktTable.Visible = false;
                    PeremZakTable.Visible = false;
                    PerevodiTable.Visible = false;
                    SmenaFamiliiTable.Visible = false;
                    IzmCedTable.Visible = false;
                    IzmStoimTable.Visible = true;
                    MatKapitalTable.Visible = false;
                    DogovorTable.Visible = false;
                    DopSoglTable.Visible = false;
                    break;
                case ("Мат. капитал"):
                    AktTable.Visible = false;
                    PeremZakTable.Visible = false;
                    PerevodiTable.Visible = false;
                    SmenaFamiliiTable.Visible = false;
                    IzmCedTable.Visible = false;
                    IzmStoimTable.Visible = false;
                    MatKapitalTable.Visible = true;
                    DogovorTable.Visible = false;
                    DopSoglTable.Visible = false;
                    break;
                case ("Переводы"):
                    AktTable.Visible = false;
                    PeremZakTable.Visible = false;
                    PerevodiTable.Visible = true;
                    SmenaFamiliiTable.Visible = false;
                    IzmCedTable.Visible = false;
                    IzmStoimTable.Visible = false;
                    MatKapitalTable.Visible = false;
                    DogovorTable.Visible = false;
                    DopSoglTable.Visible = false;
                    break;
                case ("Перемена цедента"):
                    AktTable.Visible = false;
                    PeremZakTable.Visible = false;
                    PerevodiTable.Visible = false;
                    SmenaFamiliiTable.Visible = false;
                    IzmCedTable.Visible = true;
                    IzmStoimTable.Visible = false;
                    MatKapitalTable.Visible = false;
                    DogovorTable.Visible = false;
                    DopSoglTable.Visible = false;
                    break;
                default:
                    AktTable.Visible = false;
                    PeremZakTable.Visible = false;
                    PerevodiTable.Visible = false;
                    SmenaFamiliiTable.Visible = false;
                    IzmCedTable.Visible = false;
                    IzmStoimTable.Visible = false;
                    MatKapitalTable.Visible = false;
                    DogovorTable.Visible = false;
                    DopSoglTable.Visible = false;
                    break;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}
