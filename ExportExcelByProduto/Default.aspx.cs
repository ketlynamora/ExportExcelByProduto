using System;
using System.Collections.Generic;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using ExportExcelByProduto.Modelo;


namespace ExportExcelByProduto
{
    
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void ExportExcel_Click(object remetente, EventArgs e)
        {
            /* Lista de Produtos */

            var produtos = new List<Produto>
            {
                new Produto(863, "22/02/2021 10:00", "Nordware", "Mercos", "25/02/2021 19:00", "INTEGRADO" ),
                new Produto(863, "23/02/2021 12:48", "Mercos",   "Shopee", "25/03/2021 12:09", "ERRO" ),
                new Produto(863, "23/02/2021 18:07", "Nordware", "Mercos", "15/03/2021 07:35", "INTEGRADO" ),
                new Produto(863, "25/02/2021 19:00", "Mercos",   "Mercos", "15/03/2021 07:35", "INTEGRADO" ),
                new Produto(863, "01/03/2021 07:48", "Mercos",   "Shopee", "15/03/2021 07:35", "ERRO" ),
                new Produto(863, "05/03/2021 20:50", "Nordware", "Mercos", "22/03/2021 09:00", "INTEGRADO" ),
                new Produto(863, "09/03/2021 17:15", "Mercos",   "Mercos", "22/03/2021 09:00", "INTEGRADO" ),
                new Produto(863, "15/03/2021 07:35", "Mercos",   "Shopee", "22/03/2021 09:00", "ERRO" ),
                new Produto(863, "20/03/2021 14:36", "Nordware", "Mercos", "21/03/2021 19:48", "INTEGRADO" ),
                new Produto(863, "21/03/2021 19:48", "Mercos",   "Mercos", "21/03/2021 19:48", "INTEGRADO" ),
                new Produto(863, "22/03/2021 09:00", "Nordware", "Shopee", "22/03/2021 09:00", "ERRO" ),
                new Produto(863, "22/03/2021 23:58", "Mercos",   "Mercos", "11/03/2021 12:09", "INTEGRADO" )

            };
            ExcelPackage arquivoExcel = new ExcelPackage();
            var workSheet = arquivoExcel.Workbook.Worksheets.Add(" Produtos ");

            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            //Header of table  
            workSheet.Row(1).Height = 20;
            workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Bold = true;

            workSheet.Cells[1, 1].Value = "Número";
            workSheet.Cells[1, 2].Value = "Pedido";
            workSheet.Cells[1, 3].Value = "Data de Criação";
            workSheet.Cells[1, 4].Value = "Loja";
            workSheet.Cells[1, 5].Value = "Canal";
            workSheet.Cells[1, 6].Value = "Data de Importação";
            workSheet.Cells[1, 7].Value = "Status";

            //Body of table  
            int indiceProduto = 2;
            foreach (var produto in produtos)
            {
                workSheet.Cells[indiceProduto, 1].Value = (indiceProduto - 1).ToString();
                workSheet.Cells[indiceProduto, 2].Value = produto.Pedido;
                workSheet.Cells[indiceProduto, 3].Value = produto.DataCriacao;
                workSheet.Cells[indiceProduto, 4].Value = produto.Loja;
                workSheet.Cells[indiceProduto, 5].Value = produto.CanalVenda;
                workSheet.Cells[indiceProduto, 6].Value = produto.DataImportacao;
                workSheet.Cells[indiceProduto, 7].Value = produto.Status;
                indiceProduto++;
            }

            workSheet.Column(1).AutoFit();
            workSheet.Column(2).AutoFit();
            workSheet.Column(3).AutoFit();
            workSheet.Column(4).AutoFit();
            workSheet.Column(5).AutoFit();
            workSheet.Column(6).AutoFit();
            workSheet.Column(7).AutoFit();

            string excelName = "registroProdutos";
            using (var memoryStream = new MemoryStream())
            {
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment; filename=" + excelName + ".xlsx");
                arquivoExcel.SaveAs(memoryStream);
                memoryStream.WriteTo(Response.OutputStream);
                Response.Flush();
                Response.End();
            }
        }
    }
}