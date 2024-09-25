using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using System.Xml;
using Aspose.Cells;
using System.Numerics;
using static NFe.nfeProc;
using Aspose.Cells.Tables;

namespace NFe
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string caminho;
        private void button1_Click(object sender, EventArgs e)
        {
            //Instancia o Objeto da Classe
            nfeProc nota = new nfeProc();

            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult result = fbd.ShowDialog();
            txtArquivo.Text = fbd.SelectedPath.ToString();

            //Marca o diretório a ser listado
            DirectoryInfo diretorio = new DirectoryInfo(txtArquivo.Text);
            //Executa função GetFile(Lista os arquivos desejados de acordo com o parametro)
            FileInfo[] Arquivos = diretorio.GetFiles("*.xml;");

            string[] files = System.IO.Directory.GetFiles(fbd.SelectedPath);

            // Instanciar um objeto Workbook que representa o arquivo do Excel.
            Workbook wb = new Workbook();


            

            for (int i = 2; i < files.Length+2; i++)
            {
                //Serializa o objeto
                XmlSerializer ser = new XmlSerializer(typeof(nfeProc));

                //lê o arquivo xml
                TextReader textReader = (TextReader)new StreamReader(files[i-2]);
                XmlTextReader reader = new XmlTextReader(textReader);
                reader.Read();

                //Desserializa o objeto
                nota = (nfeProc)ser.Deserialize(reader);              


                // Pegue a primeira planilha.
                Worksheet sheet = wb.Worksheets[0];

                // Obtendo a coleção de células da planilha
                Cells cells = sheet.Cells;

                cells["A1"].PutValue("Versão da Nota");
                cells["A" + i].PutValue(Convert.ToString(nota.versao));

                cells["B1"].PutValue("Código da NF");
                cells["B" + i].PutValue(Convert.ToString(nota.NFe.infNFe.ide.cNF));

                cells["C1"].PutValue("Nome do Emissor");
                cells["C" + i].PutValue(Convert.ToString(nota.NFe.infNFe.emit.xNome));

                cells["D1"].PutValue("CNPJ do Emissor");
                cells["D" + i].PutValue(Convert.ToString(nota.NFe.infNFe.emit.CNPJ));

                cells["E1"].PutValue("Data de Emissão");
                cells["E" + i].PutValue(Convert.ToString(nota.NFe.infNFe.ide.dhEmi));

                cells["F1"].PutValue("Data de Vencimento");
                cells["F" + i].PutValue(Convert.ToString(nota.NFe.infNFe.cobr.dup.dVenc));

            }

            // Salve o arquivo Excel.
            wb.Save("Excel_Table.xlsx", SaveFormat.Xlsx);


        }


    }
}
