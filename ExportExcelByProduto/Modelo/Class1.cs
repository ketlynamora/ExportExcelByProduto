using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportExcelByProduto.Modelo
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Web;

    namespace ExportExcelByProduto.Modelo
    {
        public class Produto
        {
            public int Pedido;
            public string DataCriacao;
            public string Loja;
            public string CanalVenda;
            public string DataImportacao;
            public string Status;

            public Produto(int pedido, string dataCriacao,
                string loja, string canalVenda, string dataImportacao, string status)
            {
                this.Pedido = pedido;
                this.DataCriacao = dataCriacao;
                this.Loja = loja;
                this.CanalVenda = canalVenda;
                this.DataImportacao = dataImportacao;
                this.Status = status;
            }

        }
    }
}