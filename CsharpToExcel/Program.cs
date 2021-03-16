using System;
using ClosedXML.Excel;


namespace CsharpToExcel
{
    class Program
    {
        static void Main()
        {
            Console.WriteLine("Gerando arquivos Excel...!");

            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Planilha 1");

            // Título do Relatório 
            ws.Cell("B2").Value = "Exemplo de Relatório";
            var range = ws.Range("B2:I2");
            range.Merge().Style.Font.SetBold().Font.FontSize = 20;

            // Cabeçalho do Relatório
            ws.Cell("B3").Value = "Título 1";
            ws.Cell("C3").Value = "Título 2";
            ws.Cell("D3").Value = "Título 3";
            ws.Cell("E3").Value = "Título 4";
            ws.Cell("F3").Value = "Título 5";
            ws.Cell("G3").Value = "Título 6";
            ws.Cell("I3").Value = "Subtotal";

            // Corpo do Relatório 
            var linha = 4;

            for ( int i = 0; i< 20; i++ )
            {
                ws.Cell("B" + linha.ToString()).Value = "B" + i.ToString();
                ws.Cell("C" + linha.ToString()).Value = "C" + i.ToString();
                ws.Cell("D" + linha.ToString()).Value = "D" + i.ToString();
                ws.Cell("E" + linha.ToString()).Value = "E" + i.ToString();
                ws.Cell("F" + linha.ToString()).Value = "F" + i.ToString();
                ws.Cell("G" + linha.ToString()).Value = "G" + i.ToString();
                ws.Cell("H" + linha.ToString()).Value = "H" + i.ToString();
                ws.Cell("I" + linha.ToString()).Value = String.Format("{0:F2}", i * linha);
                linha++;

            }

            // Ajusto a numeração da linha
            linha--;

            // Crio a formatação do Tipo "Money" para o nosso "Subtotal"
            ws.Range("I4:I" + linha.ToString()).Style.NumberFormat.Format = "R$ #,#.##00";

            // Crio a formatação do Tipo "Money" para o nosso "Subtotal"
            range = ws.Range("B3:I" + linha.ToString());
            range.CreateTable();
            
            // Ajusto o tamanho da coluna com o coneteúdo da conluna
            ws.Columns("2-9").AdjustToContents();

            // Salvar o arquivo em Disco
            wb.SaveAs(@"C:\Users\ketlyn.a\Documents\ProjetosTeste\teste_tne.xlsx");

            // Liberar objetos
            wb.Dispose();

            Console.WriteLine("Feito!");
            Console.ReadKey();

        }
    }
}
