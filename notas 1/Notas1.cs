using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace notas_1
{
    public partial class Notas1 : Form
    {
        public Notas1()
        {
            InitializeComponent();
        }

        private void Notas1_Load(object sender, EventArgs e)
        {
            listView1.View = View.Details;

            listView1.FullRowSelect = true;

            listView1.GridLines = true;

            listView1.LabelEdit = true;

            listView1.Columns.Add("Nome", 400, HorizontalAlignment.Left);
            listView1.Columns.Add("Numero", 100, HorizontalAlignment.Left);
            listView1.Columns.Add("Nota", 100, HorizontalAlignment.Left);
            listView1.Columns.Add("sítuação", 180, HorizontalAlignment.Left);

            nomea.Select();
        }


        private void Button1_Click(object sender, EventArgs e)
        {
            //falta de informações
            if (nomea.Text.Trim().Equals(string.Empty))
            {

                MessageBox.Show("Voce deve informar o nome ", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;

            }
            else if (numeroa.Text.Trim().Equals(string.Empty))
            {

                MessageBox.Show("Voce deve informar o numero ", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;

            }
            else if (textBox1.Text.Trim().Equals(string.Empty))
            {

                MessageBox.Show("Voce deve informar ao menos duas notas ", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;

            }
            else if(textBox2.Text.Trim().Equals(string.Empty))
            {

                MessageBox.Show("Voce deve informar ao menos duas notas ", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;

            }
            else if (mu1.Text.Trim().Equals(string.Empty))
            {
                MessageBox.Show("Voce deve informar os multiplicadores ", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;


            }
            else if (mu2.Text.Trim().Equals(string.Empty))
            {
                MessageBox.Show("Voce deve informar os multiplicadores ", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;


            }
            else if (dividendo.Text.Trim().Equals(string.Empty))
            {
                MessageBox.Show("Voce deve informar o dividendo ", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;


            }
            else if (corte.Text.Trim().Equals(string.Empty))
            {
                MessageBox.Show("Voce deve informar a nota de corte ", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;


            }

            //variaveis

            decimal n1;
            decimal n2;
            decimal n3;
            decimal n4;
            decimal n5;
            decimal n6;
            decimal m1;
            decimal m2;
            decimal m3;
            decimal m4;
            decimal m5;
            decimal m6;
            decimal c;
            decimal d;
            decimal rn;
            string rf;
            string no;
            string nu;


            //calculadora


            //notas e notações
            no = nomea.Text;
            nu = numeroa.Text;
            n1 = decimal.Parse(textBox1.Text);
            n2 = decimal.Parse(textBox2.Text);
            


            //multiplicadores
            m1 = decimal.Parse(mu1.Text);
            m2 = decimal.Parse(mu2.Text);


            //dividendos e resultados
            c = decimal.Parse(corte.Text);
            d = decimal.Parse(dividendo.Text);

            //check box

            if (checkn1.Checked)
            {

                if (mu3.Text.Trim().Equals(string.Empty))
                {
                    MessageBox.Show("Voce deve informar os multiplicadores ", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;


                }

                m3 = decimal.Parse(mu3.Text);
                n3 = decimal.Parse(textBox3.Text);
                rn = ((n1 * m1) + (n2 * m2) + (n3 * m3)) / d;
                
            }
            else if (checkn2.Checked)
            {
                if (mu3.Text.Trim().Equals(string.Empty))
                {
                    MessageBox.Show("Voce deve informar os multiplicadores ", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;


                }

                if (mu4.Text.Trim().Equals(string.Empty))
                {
                    MessageBox.Show("Voce deve informar os multiplicadores ", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;


                }

                m3 = decimal.Parse(mu3.Text);
                n3 = decimal.Parse(textBox3.Text);
                m4 = decimal.Parse(mu4.Text);
                n4 = decimal.Parse(textBox4.Text);

                rn = ((n1 * m1) + (n2 * m2) + (n3 * m3) + (n4 * m4) ) / d;
               
            }
            else if (checkn3.Checked)
            {

                if (mu3.Text.Trim().Equals(string.Empty))
                {
                    MessageBox.Show("Voce deve informar os multiplicadores ", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;


                }

                if (mu4.Text.Trim().Equals(string.Empty))
                {
                    MessageBox.Show("Voce deve informar os multiplicadores ", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;


                }

                if (mu5.Text.Trim().Equals(string.Empty))
                {
                    MessageBox.Show("Voce deve informar os multiplicadores ", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;


                }


                m3 = decimal.Parse(mu3.Text);
                n3 = decimal.Parse(textBox3.Text);
                m4 = decimal.Parse(mu4.Text);
                n4 = decimal.Parse(textBox4.Text);
                m5 = decimal.Parse(mu5.Text);
                n5 = decimal.Parse(textBox5.Text);

                rn = ((n1 * m1) + (n2 * m2) + (n3 * m3) + (n4 * m4) + (n5 * m5)) / d;
               
            }
            else if (checkn3.Checked)
            {

                if (mu3.Text.Trim().Equals(string.Empty))
                {
                    MessageBox.Show("Voce deve informar os multiplicadores ", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;


                }

                if (mu4.Text.Trim().Equals(string.Empty))
                {
                    MessageBox.Show("Voce deve informar os multiplicadores ", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;


                }

                if (mu5.Text.Trim().Equals(string.Empty))
                {
                    MessageBox.Show("Voce deve informar os multiplicadores ", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;


                }

                if (mu6.Text.Trim().Equals(string.Empty))
                {
                    MessageBox.Show("Voce deve informar os multiplicadores ", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;


                }

                m3 = decimal.Parse(mu3.Text);
                n3 = decimal.Parse(textBox3.Text);
                m4 = decimal.Parse(mu4.Text);
                n4 = decimal.Parse(textBox4.Text);
                m5 = decimal.Parse(mu5.Text);
                n5 = decimal.Parse(textBox5.Text);
                m6 = decimal.Parse(mu6.Text);
                n6 = decimal.Parse(textBox6.Text);

                rn = ((n1 * m1) + (n2 * m2) + (n3 * m3) + (n4 * m4) + (n5 * m5) + (n6 * m6)) / d;
                

            }
            else
            {

                rn = ((n1 * m1) + (n2 * m2)) / d;
                


            }

            //situação
            if( rn >= c)
            {


                rf = "Aprovado";


            }
            else
            {

                rf = "Reprovado";


            }


           resul.Text = "Aluno: " + no + "  " + " Número: " + nu + "  " + "Nota: " + Convert.ToString(rn) + "  "+ rf;

            //definir itens a serem adicionados
            ListViewItem lvi = new ListViewItem(nomea.Text.Trim());
            lvi.SubItems.Add(numeroa.Text);
            lvi.SubItems.Add(Convert.ToString(rn));
            lvi.SubItems.Add(rf);
            //adicionar itens 
            listView1.Items.Add(lvi);


        }

        private void Button2_Click(object sender, EventArgs e)
        {
            //limpador
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
            textBox6.Clear();
            nomea.Clear();
            numeroa.Clear();
            resul.Clear();

        }

        private void Button3_Click(object sender, EventArgs e)
        {
            //exporta para execel
            Microsoft.Office.Interop.Excel.Application fichanotas = new Microsoft.Office.Interop.Excel.Application();
            fichanotas.Visible = true;
            Microsoft.Office.Interop.Excel.Workbook ResultadoNotas = fichanotas.Workbooks.Add(1);
            Microsoft.Office.Interop.Excel.Worksheet NotasResultados = (Microsoft.Office.Interop.Excel.Worksheet) ResultadoNotas.Worksheets[1];
            
            //montando o layout
            int linha = 1;
            int coluna = 1;

            NotasResultados.Cells[1, 1] = listView1.Columns[0].Text;
            NotasResultados.Cells[1, 2] = listView1.Columns[1].Text;
            NotasResultados.Cells[1, 3] = listView1.Columns[2].Text;
            NotasResultados.Cells[1, 4] = listView1.Columns[3].Text;

            linha++;


            foreach (ListViewItem lvi in listView1.Items)
            {
                coluna = 1;
                foreach (ListViewItem.ListViewSubItem lvs in lvi.SubItems)
                {


                    NotasResultados.Cells[linha, coluna] = lvs.Text;
                    coluna++;


                }
                linha++;


            }

            //fim
        }

        private void Button4_Click(object sender, EventArgs e)
        {

        }
    }
}
