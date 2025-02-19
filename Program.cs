using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;

class Program
{
    static void Main()
    {
        var servicoRelatorio = new RelatorioService();
        var controladora = new RelatorioController(servicoRelatorio);
        controladora.GerarRelatorios();
    }
}

class Relatorio
{
    public string Nome { get; set; }
    public int Idade { get; set; }
    public string Cargo { get; set; }
}

class RelatorioController
{
    private readonly RelatorioService _relatorioService;

    public RelatorioController(RelatorioService relatorioService)
    {
        _relatorioService = relatorioService;
    }

    public void GerarRelatorios()
    {
        List<Relatorio> dados = _relatorioService.ObterDados();
        string pastaDestino = "./Relatorios";
        Directory.CreateDirectory(pastaDestino);

        string caminhoExcel = Path.Combine(pastaDestino, "Relatorio.xlsx");
        string caminhoPdf = Path.Combine(pastaDestino, "Relatorio.pdf");

        _relatorioService.ExportarParaExcel(dados, caminhoExcel);
        _relatorioService.ExportarParaPDF(dados, caminhoPdf);

        Console.WriteLine("Relatórios gerados com sucesso!");
    }
}

class RelatorioService
{
    public List<Relatorio> ObterDados()
    {
        return new List<Relatorio>
        {
            new Relatorio { Nome = "João", Idade = 30, Cargo = "Desenvolvedor" },
            new Relatorio { Nome = "Maria", Idade = 28, Cargo = "Analista" },
            new Relatorio { Nome = "Carlos", Idade = 35, Cargo = "Gerente" }
        };
    }

    public void ExportarParaExcel(List<Relatorio> dados, string caminho)
    {
        using (var workbook = new XLWorkbook())
        {
            var planilha = workbook.Worksheets.Add("Relatório");
            planilha.Cell(1, 1).Value = "Nome";
            planilha.Cell(1, 2).Value = "Idade";
            planilha.Cell(1, 3).Value = "Cargo";

            for (int i = 0; i < dados.Count; i++)
            {
                planilha.Cell(i + 2, 1).Value = dados[i].Nome;
                planilha.Cell(i + 2, 2).Value = dados[i].Idade;
                planilha.Cell(i + 2, 3).Value = dados[i].Cargo;
            }

            workbook.SaveAs(caminho);
        }
    }

    public void ExportarParaPDF(List<Relatorio> dados, string caminho)
    {
        using (var writer = new PdfWriter(caminho))
        using (var pdf = new PdfDocument(writer))
        {
            var doc = new Document(pdf);
            doc.Add(new Paragraph("Relatório de Funcionários"));

            foreach (var item in dados)
            {
                doc.Add(new Paragraph($"Nome: {item.Nome}, Idade: {item.Idade}, Cargo: {item.Cargo}"));
            }
        }
    }
}
