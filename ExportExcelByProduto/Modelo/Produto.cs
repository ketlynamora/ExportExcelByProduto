using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportExcelByProduto.Modelo
{
    public class Produto
    {

        private int _pedido;
        private string _dataCriacao;
        private string _loja;
        private string _canalVenda;
        private string _dataImportacao;
        private string _status;

        public Produto(int pedido, string dataCriacao,
            string loja, string canalVenda, string dataImportacao, string status)
        {
            this._pedido = pedido;
            this._dataCriacao = dataCriacao;
            this._loja = loja;
            this._canalVenda = canalVenda;
            this._dataImportacao = dataImportacao;
            this._status = status;
        }

        public int Pedido 
        { 
            get => _pedido; 
            set => _pedido = value; 
        }
        public string DataCriacao 
        { 
            get => _dataCriacao; 
            set => _dataCriacao = value; 
        }
        public string Loja 
        { 
            get => _loja; 
            set => _loja = value; 
        }
        public string CanalVenda 
        { 
            get => _canalVenda; 
            set => _canalVenda = value; 
        }
        public string DataImportacao 
        { 
            get => _dataImportacao; 
            set => _dataImportacao = value; 
        }
        public string Status 
        { 
            get => _status; 
            set => _status = value; 
        }

        /*
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

        public int Pedido { get; private set; }
        public string DataCriacao { get; private set; }
        public string Loja { get; private set; }
        public string CanalVenda { get; private set; }
        public string DataImportacao { get; private set; }
        public string Status { get; private set; }

        */

    }
}
