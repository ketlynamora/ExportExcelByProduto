using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportExcelByProduto.Modelo
{
    public class Produto
    {
        /*
        private int _pedido;
        private string _dataCriacao;
        private string _loja;
        private string _canalVenda;
        private string _dataImportacao;
        private string _status;
        */
        public Produto(int pedido, string dataCriacao,
            string loja, string canalVenda, string dataImportacao, string status)
        {
            Pedido = pedido;
            DataCriacao = dataCriacao;
            Loja = loja;
            CanalVenda = canalVenda;
            DataImportacao = dataImportacao;
            Status = status;
        }

        public  int Pedido { get; set; }
        public string DataCriacao { get; set; }
        public string Loja { get; set; }
        public string CanalVenda { get; set; }
        public string DataImportacao { get; set; }
        public string Status { get; set; }

    }
}
