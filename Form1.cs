using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Word = Microsoft.Office.Interop.Word;

using Excel = Microsoft.Office.Interop.Excel;

using System.IO;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using ComboBox = System.Windows.Forms.ComboBox;
using TextBox = System.Windows.Forms.TextBox;
using System.Windows.Forms.VisualStyles;

namespace Factura
{
    /*
    Nombre: Maria Rodríguez Iranzo
    Asignatura: M7
    Curso: 2022-2023 
    Nota esperada:7
    */


    public partial class Form1 : Form
    {
       
        // globales
        private string[] item;
        private int[] price;
        private ComboBox[] ComboBoxItem;
        private ComboBox[] ComboBoxUnit;
        private TextBox[] textBoxPriceU;
        private TextBox[] textBoxPrice;
        private int pTotal=0;
        //declaraciones globales:
        Excel.Application appExcel = null;
        Excel.Workbook workBookExcel = null;
        Excel.Worksheet workSheetExcel = null;

        private void dateclock()
        {
            string date = DateTime.Now.ToString("MM/dd/yyyy");
            textBoxDate.Text = date;

        }

        Word.Application apWord = null;
        Word.Document docWord = null;

        
        public Form1()
        {
            InitializeComponent();
            initConfig();
        }
        private void initConfig()
        {
            item = new string[] { "pelota1", "pelota2", "pelota3", "pelota4", "pelota5" };
            price = new int[] { 100, 20, 35, 49, 85 };
            ComboBoxItem = new ComboBox[] { comboBoxItem1, comboBoxItem2, comboBoxItem3, comboBoxItem4, comboBoxItem5 };
            ComboBoxUnit = new ComboBox[] { comboBoxUnit1, comboBoxUnit2, comboBoxUnit3, comboBoxUnit4, comboBoxUnit5 };
            textBoxPriceU = new TextBox[] { textBoxPUnit1, textBoxPUnit2, textBoxPUnit3, textBoxPUnit4, textBoxPUnit5 };
            textBoxPrice = new TextBox[] { textBoxImport1, textBoxImport2, textBoxImport3, textBoxImport4, textBoxImport5 };

            for (int i = 0; i < item.Length; i++)
            {

                ComboBoxItem[i].TabIndex =i;
                ComboBoxUnit[i].TabIndex = i;
                for (int j = 0; j < price.Length; j++)
                {
                    ComboBoxItem[i].Items.Add(item[j]);
                }
                for (int j = 0; j <= 5; j++) {
                    ComboBoxUnit[i].Items.Add(j);
                }

                ComboBoxUnit[i].SelectedIndex = 0;
                textBoxPriceU[i].Text = "";
                comboBoxIVA.SelectedIndex = 0;
            }

        }

        private void comboBoxItem1_SelectedIndexChanged(object sender, EventArgs e)
        {/*
            int index=comboBoxItem1.SelectedIndex;
            textBoxPriceU[index].Text = price[index].ToString();
        */
            }

        private void comboBoxItemSChange(object sender, EventArgs e)
        {
            ComboBox cb = (ComboBox)sender;
            int row=cb.TabIndex;
            int index= ComboBoxUnit[row].SelectedIndex;
            if (index >= 0)
            {
                textBoxPriceU[row].Text = price[index].ToString();

                int SelectUnitPrice = ComboBoxUnit[row].SelectedIndex;
                int total = price[index] * SelectUnitPrice;
                textBoxPrice[row].Text = total.ToString();
                int pTotal = funcionWithoutIVA();
                textBoxTotalWithoutIVA.Text=pTotal.ToString();
                
                if (comboBoxIVA.SelectedIndex >= 0)
                {
                    double IVA = Convert.ToDouble(comboBoxIVA.SelectedItem);
                    double TOTAL = pTotal*(1+IVA/100);
                    textBoxTOTAL.Text = (TOTAL.ToString());
                }

            }
        }
        private int funcionWithoutIVA()
        { 
                for (int row = 0; row < textBoxPrice.Length; row++)
                {
                    if (textBoxPrice[row].Text != "")
                    {
                        pTotal += int.Parse(textBoxPrice[row].Text);
                    }
                }
                return pTotal;
        }

        private void comboBoxIVA_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (comboBoxIVA.SelectedIndex >= 0)
            {
                double IVA = Convert.ToDouble(comboBoxIVA.SelectedItem);
                double TOTAL = pTotal * (1 + IVA / 100);
                textBoxTOTAL.Text = (TOTAL.ToString());
            }
        }
        //word
        public void sendToWord()
        {
            CrearFicheroWord();
            foreach (Control control in this.Controls)
            {
                if (control is TextBox || control is ComboBox)
                {
                    wordWriter(control);
                }
                if (control is GroupBox)
                {
                    foreach (Control c in control.Controls)
                    {
                        if (c is TextBox || c is ComboBox)
                        {
                            wordWriter(c);
                        }
                    }
                }
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
        }
        private void wordWriter(Control control)
        {
            try
            {
                // sacar el name y text
                Object bookMarcName = control.Name;
                string text = control.Text;
                docWord.Bookmarks[ref bookMarcName].Select();
                apWord.Selection.TypeText(Text: text);
            }
            catch (Exception e)
            {
            }
        }
        private void CrearFicheroWord()
        {
            Word.Application apWord0 = new Word.Application();
            Word.Document docWord0 = new Word.Document();
            string path = Directory.GetCurrentDirectory();
            path = path + "\\Facturas\\";
            string wordFileNameIn = "Plantilla.docx";
            string wordFileNameOut = "Factura_out.docx";
            docWord0 = apWord0.Documents.Open(path + wordFileNameIn);
            try { docWord0.SaveAs(path + wordFileNameOut); }
            catch (Exception e) { }
            docWord0.Close();
            apWord0.Quit();
                apWord = new Word.Application();
            docWord = new Word.Document();
            docWord = apWord.Documents.Open(path + wordFileNameOut);
        }
        private void GuardarFicheroWord()
        {
            docWord.Save();
        }
        private void CerrarWord()
        {
            apWord.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(apWord);
            apWord = null;
        }
        public void WordKiller()
        {
            System.Diagnostics.Process[] procs =
            System.Diagnostics.Process.GetProcessesByName("WINWORD");
            if (procs.Length >= 1)
            {
                for (int i = 0; i < procs.Length; i++)
                {
                    try { procs[i].Kill(); }
                    catch (Exception e) { }
                }
            }
        }

        private void guardarFicheroPDF()
        {
            Word.Application apWord = new Word.Application();
            Word.Document docWord = new Word.Document();
            string path = Directory.GetCurrentDirectory();
            path = path + "\\Facturas\\";
            string wordFileNameIn = "Factura_out.docx";
            docWord = apWord.Documents.Open(path + wordFileNameIn);
            string PdfFileNameOut = path + "Factura_out.pdf";
            docWord.SaveAs(PdfFileNameOut, Word.WdSaveFormat.wdFormatPDF);



            docWord.Close();
            apWord.Quit();
            MessageBox.Show("done");
        }


        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void button1_Click_1(object sender, EventArgs e)
        {
            WordKiller();
            sendToWord();
            MessageBox.Show("done");
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            dateclock();
        }



        private void button2_Click(object sender, EventArgs e)
        {
            guardarFicheroPDF();
            MessageBox.Show("done");
        }

        private void gestionExcel()
        {
            abrirFicheroExcel();
            EscribirEnFicheroExcel();
            guardarFicheroExcel();
            cerrarExcel();
            MessageBox.Show("Excel file created");
        }
        private void abrirFicheroExcel()
        {
            appExcel = new Excel.Application();
            String path = Directory.GetCurrentDirectory() + "\\Facturas\\";
            String ExcelFileNameIn = "Factura_plantilla";
            workBookExcel = appExcel.Workbooks.Open(path + ExcelFileNameIn);
            workSheetExcel = (Excel.Worksheet)workBookExcel.Worksheets.get_Item(1);
        }
        private void EscribirEnFicheroExcel()
        {
            int fila, col; string text;
            fila = 3; col = 3; text = textBoxSocialReason.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 3; col = 5; text = textBoxNIF1.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 4; col = 3; text = textBoxResidence1.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 6; col = 3; text = textBoxSocialReason.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 6; col = 5; text = textBoxNIF2.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 7; col = 3; text = textBoxResidence2.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 10; col = 3; text = textBoxNumber.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 10; col = 5; text = textBoxDate.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 11; col = 3; text = textBoxIOrder.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 11; col = 5; text = textBoxIReference.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 15; col = 2; text = comboBoxItem1.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 16; col = 2; text = comboBoxItem2.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 17; col = 2; text = comboBoxItem3.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 18; col = 2; text = comboBoxItem4.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 19; col = 2; text = comboBoxItem5.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 15; col = 3; text = comboBoxUnit1.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 16; col = 3; text = comboBoxUnit2.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 17; col = 3; text = comboBoxUnit3.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 18; col = 3; text = comboBoxUnit4.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 19; col = 3; text = comboBoxUnit5.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 15; col = 4; text = textBoxPUnit1.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 16; col = 4; text = textBoxPUnit2.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 17; col = 4; text = textBoxPUnit3.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 18; col = 4; text = textBoxPUnit4.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 19; col = 4; text = textBoxPUnit5.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 15; col = 5; text = textBoxImport1.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 16; col = 5; text = textBoxImport2.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 17; col = 5; text = textBoxImport3.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 18; col = 5; text = textBoxImport4.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 19; col = 5; text = textBoxImport5.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 23; col = 2; text = textBoxTotalWithoutIVA.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 23; col = 3; text = comboBoxIVA.Text;
            workSheetExcel.Cells[fila, col] = text;
            fila = 23; col = 4; text = textBoxTOTAL.Text;
            workSheetExcel.Cells[fila, col] = text;
            appExcel.DisplayAlerts = false;
            workSheetExcel.Cells[fila, col] = text;
        }
        private void guardarFicheroExcel()
        {
            String path = Directory.GetCurrentDirectory() + "\\Facturas\\";
            String ExcelFileNameOut = "Factura_out";
            workBookExcel.SaveAs(path + ExcelFileNameOut);
        }
        private void cerrarExcel()
        {
            workBookExcel.Close(true);//guardar los cambios:true
            appExcel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheetExcel);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workBookExcel);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(appExcel);
        }
        private void excel()
        {
            Excel.Application appExcel = new Excel.Application();
            Excel.Workbook workBookExcel = null;
            Excel.Worksheet workSheetExcel = null;
            workBookExcel = appExcel.Workbooks.Open(Directory.GetCurrentDirectory() +
           "\\Facturas\\" + "Factura_plantilla");
            workSheetExcel = (Excel.Worksheet)workBookExcel.Worksheets.get_Item(1);
            int fila, col; string text;
            fila = 1; col = 1; text = "Monlaujbm";
            appExcel.DisplayAlerts = false;
            workSheetExcel.Cells[fila, col] = text;
            workBookExcel.SaveAs(Directory.GetCurrentDirectory() + "\\Facturas\\" +
           "Factura_out");
            workBookExcel.Close(true);//guardar los cambios:true
            appExcel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(appExcel);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            gestionExcel();
        }

        private void textBoxSocialReason_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load_1(object sender, EventArgs e)
        {

        }
    }
}
