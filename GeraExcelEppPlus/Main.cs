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

namespace GeraExcelEppPlus
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private void btnGeraExcel_Click(object sender, EventArgs e)
        {
            string nomeUnico = Guid.NewGuid().ToString() + ".xlsx";

            //Estou gerando um arquivo Excel de Exemplo na pasta de execução do Programa com um Nome GUID unico.
            string caminhoGerar = Path.Combine(Application.StartupPath, nomeUnico);

            ExcelEppPlus geraExecel = new ExcelEppPlus();
            geraExecel.GerarExcel(caminhoGerar);
        }
    }
}
