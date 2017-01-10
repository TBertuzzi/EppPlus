using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeraExcelEppPlus
{
    public class ExcelEppPlus
    {
        public void GerarExcel(string destino)
        {
            //Diretorio de Destino da Planilha
            FileInfo fi = new FileInfo(destino);

            //Declara o Objeto Principal que Gera o Excel
            using (ExcelPackage xlPackage = new ExcelPackage(fi))
            {
                //Gera a Planilha de trabalho dentro do Arquivo, pode se dar o nome que quiser
                ExcelWorksheet ws = xlPackage.Workbook.Worksheets.Add("Telefones");

                //utilizando a função é possivel definir nos cabeçalhos o tamanho e o valor
                AdicionarCabecalho(ws, 1, 1, "Nome", 20);
                AdicionarCabecalho(ws, 1, 2, "Telefone", 30);

                //É possivel Habilitar auto filtros nas Colunas
                ws.Cells[1, 1, 1, 1].AutoFilter = true;
                ws.Cells[1, 1, 1, 2].AutoFilter = true;

                //No Exemplo abaixo serão adicionadas 2 linhas, é possivel fazer com um laço.

                //A Primeira Linha (1) é o cabeçalho;

                //Segunda Linha Coluna 1
                ws.Cells[2, 1].Value = "Thiago Bertuzzi";
                ws.Cells[2, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                //Segunda Linha Coluna 2
                ws.Cells[2, 2].Value = "96666666";
                ws.Cells[2, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                //Terceira Linha Coluna 1
                ws.Cells[3, 1].Value = "Thiago Vieira";
                ws.Cells[3, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                //Terceira Linha Coluna 2
                ws.Cells[3, 2].Value = "988888";
                ws.Cells[3, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                //É possivel definir o alinhamento das celulas com ExcelHorizontalAlignment.

                //Após incluir todas as linhas desejadas, deve se fechar o arquivo e salva-lo :
                xlPackage.Save();
            }
        }

        //Função Utilizada para Gerar o Cabeçalho da Planilha
        private static void AdicionarCabecalho(ExcelWorksheet ws, int lin, int col, object valor, int largura)
        {
            //propriedades da Celula
            ws.Cells[lin, col].Value = valor;
            ws.Cells[lin, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells[lin, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells[lin, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.CornflowerBlue);
            ws.Cells[lin, col].Style.Font.Color.SetColor(System.Drawing.Color.White);
            ws.Cells[lin, col].Style.Font.Bold = true;

            ws.Column(col).Width = largura;
            ws.Column(col).Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            ws.Column(col).Style.Border.Top.Style = ExcelBorderStyle.Thin;
            ws.Column(col).Style.Border.Left.Style = ExcelBorderStyle.Thin;
            ws.Column(col).Style.Border.Right.Style = ExcelBorderStyle.Thin;
        }

    }
}
