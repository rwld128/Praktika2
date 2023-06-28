using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowsFormsApp2;
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
            Dictionary<string, string> items = new Dictionary<string, string>(origin);
            switch (comboBox1.SelectedItem.ToString())
            {
                case ("Акт об оказании услуг"):
                    items.Add("<AKT_NUM>", textBoxAKT_NUM.Text);
                    items.Add("<AKT_DATE>", dateTimePickerAKT_DATE.Value.ToString("dd.MM.yyyy"));
                    items.Add("<YUR_ZAK_F_IO>", ChangeName(textBoxZAK_FIO.Text));
                    items.Add("<STUDENT_F_IO>", ChangeName(textBoxSTUDENT_FIO.Text));
                    if (checkBox1.Checked)
                    {
                        var paster = new WordPaster("Акт об оказании услуг юр.docx");
                        paster.Process(items);
                    }
                    else
                    {
                        var paster = new WordPaster("Акт об оказании услуг.docx");
                        paster.Process(items);
                    }
                    break;
                case ("Договор"):
                    items.Add("<NAPR>", textBoxNAPR.Text);
                    items.Add("<PROFIL>", textBoxPROFIL.Text);
                    items.Add("<LEVEL>", textBoxLEVEL.Text);
                    items.Add("<FORM>", textBoxFORM.Text);
                    items.Add("<SROK>", textBoxSROK.Text);
                    items.Add("<KURS>", textBoxKURS.Text);
                    items.Add("<YEARS>", textBoxYEARS.Text);
                    items.Add("<FULL_PRICE>", textBoxFULL_PRICE.Text);
                    items.Add("<YEARS_PRICE>", textBoxYEARS_PRICE.Text);
                    items.Add("<STUDENT_BD>", dateTimePickerSTUDENT_BD.Text);
                    items.Add("<STUD_PASP_SER>", textBoxSTUD_PASP_SER.Text);
                    items.Add("<STUD_PASP_NOM>", textBoxSTUD_PASP_NOM.Text);
                    items.Add("<STUD_PASP_VID>", textBoxSTUD_PASP_VID.Text);
                    if (checkBox1.Checked)
                    {
                        var paster = new WordPaster("Договор юр.docx");
                        paster.Process(items);
                    }
                    else
                    {
                        var paster = new WordPaster("Договор.docx");
                        paster.Process(items);
                    }
                    break;
                case ("Доп. соглашение"):
                    items.Add("DOP_SOGL_DATE", dateTimePickerDOP_SOGL_DATE.Text);
                    items.Add("DOP_SOGL_NUM", textBoxDOP_SOGL_NUM.Text);
                    items.Add("POLN_STOIM", textBoxPOLN_STOIM.Text);
                    items.Add("DSYEARS", textBoxDSYEARS.Text);
                    items.Add("DSYEARS_PRICE", textBoxDSYEARS_PRICE.Text);
                    items.Add("OSEN", textBoxOSEN.Text);
                    items.Add("VESNA", textBoxVESNA.Text);
                    items.Add("DS_EKZ", textBoxDS_EKZ.Text);
                    items.Add("<STUDENT_F_IO>", ChangeName(textBoxSTUDENT_FIO.Text));
                    if (checkBox1.Checked)
                    {
                        var paster = new WordPaster("Доп соглашение юр.docx");
                        paster.Process(items);
                    }
                    else
                    {
                        var paster = new WordPaster("Доп соглашение.docx");
                        paster.Process(items);
                    }
                    break;
                case ("Перемена заказчика"):
                    items.Add("<SOGL_ZAK_NUM>", textBoxSOGL_ZAK_NUM.Text);
                    items.Add("<SOGL_ZAK_DATE>", dateTimePickerSOGL_ZAK_DATE.Text);
                    items.Add("<NEW_ZAK_FIO>", textBoxNEW_ZAK_FIO.Text);
                    items.Add("<NEW_ZAK_ADRES>", textBoxNEW_ZAK_ADRES.Text);
                    items.Add("<NEW_ZAK_PHONE>", textBoxNEW_ZAK_PHONE.Text);
                    items.Add("<NEW_ZAK_EMAIL>", textBoxNEW_ZAK_EMAIL.Text);
                    items.Add("<NEW_INN_PASP_BANK>", textBoxNEW_INN_PASP_BANK.Text);
                    items.Add("<NEW_ZAK_EKZ>", textBoxNEW_ZAK_EKZ.Text);
                    items.Add("<STUDENT_F_IO>", ChangeName(textBoxSTUDENT_FIO.Text));
                    if (checkBox1.Checked)
                    {
                        var paster = new WordPaster("Перемена заказчика юр.docx");
                        paster.Process(items);
                    }
                    else
                    {
                        var paster = new WordPaster("Перемена заказчика.docx");
                        paster.Process(items);
                    }
                    break;
                case ("Смена фамилии"):
                    items.Add("<FAM_SOGL_DATE>", dateTimePickerFAM_SOGL_DATE.Text);
                    items.Add("<FAM_SOGL_NUM>", textBoxFAM_SOGL_NUM.Text);
                    items.Add("<ZAV_DATE>", dateTimePickerZAV_DATE.Text);
                    items.Add("<NEW_ZAK_FIO>", textBoxNEW_FAM_FIO.Text);
                    items.Add("<FAM_EKZ>", textBoxFAM_EKZ.Text);
                    if (checkBox1.Checked)
                    {
                        var paster = new WordPaster("Смена фамилии юр.docx");
                        paster.Process(items);
                    }
                    else
                    {
                        var paster = new WordPaster("Смена фамилии.docx");
                        paster.Process(items);
                    }
                    break;
                case ("Изменение стоимости"):
                    items.Add("<IZM_ST_DATE>", dateTimePickerIZM_ST_DATE.Text);
                    items.Add("<IZM_ST_NUM>", textBoxIZM_ST_NUM.Text);
                    items.Add("<NEW_FULL_PRICE>", textBoxNEW_FULL_PRICE.Text);
                    items.Add("<NYEARS>", textBoxNYEARS.Text);
                    items.Add("<NEW_OSEN>", textBoxNEW_OSEN.Text);
                    items.Add("<NEW_VESNA>", textBoxNEW_VESNA.Text);
                    items.Add("<NEW_PRICE_EKZ>", textBoxNEW_PRICE_EKZ.Text);
                    items.Add("<STUDENT_F_IO>", ChangeName(textBoxSTUDENT_FIO.Text));
                    if (checkBox1.Checked)
                    {
                        var paster = new WordPaster("Изменение стоимости юр.docx");
                        paster.Process(items);
                    }
                    else
                    {
                        var paster = new WordPaster("Изменение стоимости.docx");
                        paster.Process(items);
                    }
                    break;
                case ("Мат. капитал"):
                    items.Add("<MK_DATE>", dateTimePickerMK_DATE.Text);
                    items.Add("<MK_NUM>", textBoxMK_NUM.Text);
                    items.Add("<MK_YEARS>", textBoxMK_YEARS.Text);
                    items.Add("<MK_YERS_PRICE>", textBoxMK_YERS_PRICE.Text);
                    items.Add("<MK_OSEN>", textBoxMK_OSEN.Text);
                    items.Add("<MK_VESNA>", textBoxMK_VESNA.Text);
                    items.Add("<PLATA_DATE>", dateTimePickerPLATA_DATE.Text);
                    items.Add("<SERT_NUM>", textBoxSERT_NUM.Text);
                    items.Add("<SERT_SER>", textBoxSERT_SER.Text);
                    items.Add("<SERT_VID>", textBoxSERT_VID.Text);
                    items.Add("<SERT_NAME>", textBoxSERT_NAME.Text);
                    items.Add("<MK_EKZ>", textBoxMK_EKZ.Text);
                    if (checkBox1.Checked)
                    {
                        var paster = new WordPaster("Мат капитал юр.docx");
                        paster.Process(items);
                    }
                    else
                    {
                        var paster = new WordPaster("Мат капитал.docx");
                        paster.Process(items);
                    }
                    break;
                case ("Переводы"):
                    if (checkBox1.Checked)
                    {

                    }
                    else
                    {

                    }
                    break;
                case ("Перемена цедента"):
                    if (checkBox1.Checked)
                    {

                    }
                    else
                    {

                    }
                    break;
                default:
                    break;
            }
        }
    }
}
