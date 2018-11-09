using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GerarExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            //Start - Importar um arquivo de Excel e imprimir os dados no Console.
            List<FuncionariosDTO> dadosfuncionarios = new List<FuncionariosDTO>();

            var package = new ExcelPackage(new FileInfo(@"C:\GerarExcel\GerarExcel\Cargos.xlsx"));

            ExcelWorksheet workSheet = package.Workbook.Worksheets[1];

            int colCount = workSheet.Dimension.End.Column;
            int rowCount = workSheet.Dimension.End.Row;

            //Construindo um ambiente do Console
            Console.WriteLine("".PadRight(80, '-'));
            Console.WriteLine("Nome".PadRight(15) + "funcionarios".PadRight(50) + "Idade");
            Console.WriteLine("".PadRight(80, '-'));

            //Desconsidera o cabecalho
            FuncionariosDTO funcionarios;
            for (int row = 2; row <= rowCount; row++)
            {
                dadosfuncionarios.Add(funcionarios = new FuncionariosDTO()
                {
                    Nome = workSheet.Cells[row, 1].Value.ToString(),
                    Cargo = workSheet.Cells[row, 2].Value.ToString(),
                    Idade = Convert.ToInt32(workSheet.Cells[row, 3].Value)

                });
                Console.WriteLine(funcionarios.Nome.PadRight(15) + funcionarios.Cargo.PadRight(50) + funcionarios.Idade);
            }
            Console.WriteLine("Feito");
            //End

            //Start - Exportar os dados. Nesse caso eu quero apenar exportar apenas dado "Nome".
            Console.WriteLine("Deseja exportar?");
            string resposta = Console.ReadLine();

            if ("s" == resposta || "S" == resposta || "sim" == resposta || "Sim" == resposta)
            {
                using (ExcelPackage excel = new ExcelPackage())
                {
                    excel.Workbook.Worksheets.Add("Funcionário");
                    var headerRow = new List<string[]>()
                {
                    new string[] { "Nome" }
                };
                    //Primeira Linha
                    string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";

                    //Aba que vamos trabalhar 
                    var worksheet = excel.Workbook.Worksheets["Funcionário"];
                    worksheet.Cells[headerRange].LoadFromArrays(headerRow);

                    ////Formata a coluna para texto se necessário
                    //worksheet.Column(1).Style.Numberformat.Format = "@";
                    //worksheet.Column(2).Style.Numberformat.Format = "@";

                    //Popula o excel
                    List<Cliente> clientes = Cliente.ObterClientes(dadosfuncionarios);
                    int linha = 2;
                    if (clientes.Count > 0)
                    {
                        foreach (var item in clientes)
                        {
                            worksheet.Cells[linha, 1].Value = item.Nome;
                            linha++;
                        }
                    }
                    FileInfo excelFile = new FileInfo(@"C:\Users\re029391\Desktop\Empresa.xlsx");
                    excel.SaveAs(excelFile);
                    Console.WriteLine("Exportado, vai ao Desktop!");
                }
            }
            else
            {
                Console.WriteLine("Até mais!");
            }
            //End
            Console.ReadKey();
            //Environment.Exit(0);
        }
    }

    public class Cliente
    {
        public int Idade { get; private set; }
        public string Nome { get; private set; }
        public string Cargo { get; private set; }

        public static List<Cliente> ObterClientes(List<FuncionariosDTO> dadosfuncionarios)
        {
            var clientes = new List<Cliente>();

            foreach (var item in dadosfuncionarios)
            {
                clientes.Add(new Cliente() { Idade = item.Idade, Nome = item.Nome, Cargo = item.Cargo });
            }
            return clientes;
        }
    }
}
