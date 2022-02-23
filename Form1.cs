using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QueryDePara_BKO
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btOpenFile_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Selecionar Planilha Excel";
            openFileDialog1.Filter = "Excel|*.xlsx";
            openFileDialog1.ShowDialog();

            txPlanilha.Text = openFileDialog1.FileName;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string query = "";
            if (cbTruncate.Checked)
                query += $"TRUNCATE TABLE [{txDatabase.Text}].[{txSchema.Text}].[{txTable.Text}];\n";

            query += $"INSERT INTO [{txDatabase.Text}].[{txSchema.Text}].[{txTable.Text}] (nomenclatura_comercial, nomenclatura_intergrall, nomenclatura_atlys_planos, nomenclatura_atlys_bonus_franquia, preco_plano) VALUES \n";

            int nomenclatura_comercial = Convert.ToInt32(txNomenclaturaComercial.Text);
            int nomenclatura_intergrall = Convert.ToInt32(txNomenclaturaIntergrall.Text);
            int nomenclatura_atlys_planos = Convert.ToInt32(txNomenclaturaAtlysPlanos.Text);
            int nomenclatura_atlys_bonus_franquia = Convert.ToInt32(txNomenclaturaAtlysBonusFranquia.Text);
            int preco_plano = Convert.ToInt32(txPrecoPlano.Text);
            int count = 0;
            FileInfo f = new FileInfo(txPlanilha.Text);
            var package = new ExcelPackage(f);
            ExcelWorksheet ws = package.Workbook.Worksheets[0];
            for (int i = 2; i <= ws.Dimension.End.Row; i++)
            {
                if (ws.Cells[i, nomenclatura_comercial].Value == null)
                    continue;
#pragma warning disable S1643 // Strings should not be concatenated using '+' in a loop
                string auxPrecoPlano = ws.Cells[i, preco_plano].Value.ToString().Trim();
                auxPrecoPlano = Regex.Replace(auxPrecoPlano, "R\\$", "");
                double precoPlano = Convert.ToDouble(auxPrecoPlano);
                string auxAtlysBonusFranquia = ws.Cells[i, nomenclatura_atlys_bonus_franquia].Value == null ? "" : ws.Cells[i, nomenclatura_atlys_bonus_franquia].Value.ToString();

                query += $"('{ws.Cells[i, nomenclatura_comercial].Value.ToString()}', '{ws.Cells[i, nomenclatura_intergrall].Value.ToString()}', '{ws.Cells[i, nomenclatura_atlys_planos].Value.ToString()}', '{auxAtlysBonusFranquia}', '{precoPlano.ToString("0.##")}'),\n";
#pragma warning restore S1643 // Strings should not be concatenated using '+' in a loop
                count++;
                /*
                item.Tarefa = ws.Cells[i, 1].Value.ToString();
                item.Tipo = ws.Cells[i, 2].Value.ToString();
                item.LocalNumLP = ws.Cells[i, 3].Value.ToString();
                item.MesConta = ws.Cells[i, 4].Value.ToString();
                item.ValorNFOriginal = ws.Cells[i, 5].Value.ToString();
                item.ValorNFCorrigida = ws.Cells[i, 6].Value.ToString();
                item.VencimentoProrrogacao = ws.Cells[i, 7].Value.ToString();
                item.CCM = ws.Cells[i, 8].Value.ToString();
                item.Protocolo = ws.Cells[i, 9].Value.ToString();
                item.LoginCSO = ws.Cells[i, 10].Value.ToString();
                item.DataTransacao = ws.Cells[i, 11].Value.ToString();
                */
            }
            lbValoresLidos.Text = $"Valores Lidos: {count}";
            query = query.Remove(query.Length - 2, 2);
            richTextBox1.Text = query;
        }
    }
}
