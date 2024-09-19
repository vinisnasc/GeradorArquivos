using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace CriandoArquivosTeste
{/// <summary>
 /// Classe que gera arquivos xlsx com auxilio da biblioteca ClosedXML
 /// </summary>
 /// <typeparam name="T">Deve ser informado o tipo de objeto a ser trabalhado</typeparam>
    public static class ArquivosExcel<T> where T : class
    {
        /// <summary>
        /// Cria um arquivo com uma unica planilha e tabela, a partir de um unico objeto informado.
        /// </summary>
        /// <param name="objeto"></param>
        /// <param name="path"></param>
        /// <param name="nomeArquivo"></param>
        /// <returns></returns>
        public static bool CriarArquivoExcel(T objeto, string path, string nomeArquivo)
            => CriarArquivoExcel(new List<T>() { objeto }, path, nomeArquivo);

        /// <summary>
        /// Cria um arquivo com uma unica planilha e tabela, a partir da lista informada.
        /// </summary>
        /// <param name="lista">Lista do objeto a ser criado a tabela.</param>
        /// <param name="path">Caminho a ser salvo o arquivo.</param>
        /// <param name="nomeArquivo">Nome do arquivo.</param>
        public static bool CriarArquivoExcel(List<T> lista, string path, string nomeArquivo)
        {
            if (!DadosValidos(lista, path)) return false;
            nomeArquivo = FormataNome(nomeArquivo);

            var arquivo = new XLWorkbook();

            // Cria a planilha
            var planilha = arquivo.Worksheets.Add("planilha 1");

            // Criação da tabela
            var table = new DataTable();

            // Criação do cabeçãlho
            CriaCabecalho(lista, table);

            // adiciona as linhas
            AdicionarDados(lista, table);

            planilha.Cell(1, 1).InsertTable(table);
            planilha.Table(0).Theme = XLTableTheme.TableStyleMedium26;//.Style.Fill.BackgroundColor = XLColor.Amber;
            planilha.Columns().AdjustToContents();

            arquivo.SaveAs($"{path}\\{nomeArquivo}.xlsx");
            return true;
        }

        private static bool DadosValidos(List<T> lista, string path)
        {
            // Verifica se a lista é valida
            if (lista is null || lista.Count == 0) return false;

            // Verifica se o caminho é válido
            // para .net 7 utilizar Path.Exists
            if (!Directory.Exists(path)) return false;

            return true;
        }

        private static string FormataNome(string nomeArquivo)
            => (nomeArquivo?.Trim() ?? "Novo arquivo");

        private static void CriaCabecalho(List<T> lista, DataTable table)
        {
            string json = JsonSerializer.Serialize(lista[0]);
            var values = JsonSerializer.Deserialize<Dictionary<string, object>>(json);
            foreach (var atributo in values)
            {
                var valorString = atributo.Value?.ToString();

                if (int.TryParse(valorString, out var val))
                    table.Columns.Add(atributo.Key, typeof(int));
                else if (double.TryParse(valorString, out var valor))
                    table.Columns.Add(atributo.Key, typeof(double));
                else
                    table.Columns.Add(atributo.Key, typeof(string));
            }
        }

        private static void AdicionarDados(List<T> lista, DataTable table)
        {
            for (int i = 0; i < lista.Count; i++)
            {
                var row = table.NewRow();
                row.BeginEdit();

                var jsonDado = JsonSerializer.Serialize(lista[i]);
                var valueDado = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(jsonDado);
                var listaValues = valueDado.Values.ToList();

                for (int j = 0; j < valueDado.Count; j++)
                {
                    var jsonElement = listaValues[j];
                    if (jsonElement.ValueKind == JsonValueKind.Number && jsonElement.TryGetInt32(out var intValue))
                        row[j] = intValue;
                    
                    else if (jsonElement.ValueKind == JsonValueKind.Number && jsonElement.TryGetDouble(out var doubleValue))
                        row[j] = doubleValue;
                    
                    else if (jsonElement.ValueKind == JsonValueKind.String)
                        row[j] = jsonElement.GetString();
                    
                    else if (jsonElement.ValueKind == JsonValueKind.True || jsonElement.ValueKind == JsonValueKind.False)
                        row[j] = jsonElement.GetBoolean();
                    
                    else if (jsonElement.ValueKind == JsonValueKind.Null)
                        row[j] = DBNull.Value; 
                    
                    else
                        row[j] = jsonElement.ToString();
                }

                row.EndEdit();
                table.Rows.Add(row);
            }

        }
    }
}