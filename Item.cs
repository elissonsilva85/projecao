using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using Finisar.SQLite;

namespace Apresentacao
{
    public enum TipoItem
    {
        Hino,
        NovoHino,
        Arquivo,
        Aviso,
        Template,
        Biblia,
        ChaveValor,
        Nenhum
    }

    public class TemplateItem
    {
        public string nome;
        public string caminho;

        public TemplateItem(string n, string c)
        {
            nome = n.Substring(0, n.LastIndexOf('.')); // tira a extensão
            caminho = c;
        }

        public override String ToString()
        {
            return nome;
        }
    }

    public class HinoItem
    {
        public string nome;
        public string caminho;
        public string letra;

        public HinoItem(string n, string c)
        {
            nome = n.Substring(0, n.LastIndexOf('.')); // tira a extensão
            caminho = c;

            String linha;
            System.IO.StreamReader intro = new System.IO.StreamReader(c);
            while ( (linha = intro.ReadLine()) != null )
            {
                letra += linha + "\r\n";
            }
            intro.Close();
        }

        public override String ToString()
        {
            return nome;
        }
    }

    public class AvisoItem
    {
        public string nome;
        public string texto;
        public string caminho;

        public AvisoItem(string n, string t)
        {
            nome = n;
            texto = t;
        }

        public AvisoItem(string n, string t, string c)
        {
            nome = n.Substring(0, n.LastIndexOf('.')); // tira a extensão
            texto = t;
            caminho = c;
        }

        public override String ToString()
        {
            return nome;
        }
    }

    public class ArquivoItem
    {
        public string nome;
        public string caminho;

        public ArquivoItem(string n, string c)
        {
            nome = n.Substring(0, n.LastIndexOf('.')); // tira a extensão
            caminho = c;
        }

        public override String ToString()
        {
            return nome;
        }
    }

    public class BibliaItem
    {
        private string codTraducao, traducao;
        private string codLivro, livro;
        private int capitulo;
        private int versiculo;
        private string texto;
        private int incluir;

        private int navegarVersiculo;
        private int navegarCapitulo;

        public BibliaItem(string t, string l, int c, int v, int i = 0)
        {
            codTraducao = t;
            codLivro = l;
            capitulo = c;
            versiculo = v;
            incluir = i;

            navegarVersiculo = versiculo;
            navegarCapitulo = capitulo;

            preparaDados();
        }

        public void ReiniciarNavegacao()
        {
            navegarVersiculo = versiculo;
            navegarCapitulo = capitulo;
        }

        private static SQLiteConnection conectaBiblia()
        {
            SQLiteConnection sqlite_conn;

            sqlite_conn = new SQLiteConnection(); // Create an instance of the object
            sqlite_conn.ConnectionString = "Data Source=Biblia\\biblia.db;Version=3;New=False;Compress=False;UTF8Encoding=True"; // Set the ConnectionString
            try
            {
                sqlite_conn.Open(); // Open the connection. Now you can fire SQL-Queries
            }
            catch (Exception e)
            {
                System.Console.WriteLine(e.Message);
            }

            return sqlite_conn;
        }

        private void preparaDados()
        {
            SQLiteConnection sqlite_conn = conectaBiblia();
            SQLiteCommand sqlite_cmd;
            SQLiteDataReader sqlite_datareader;

            // -------------------
            
            sqlite_cmd = sqlite_conn.CreateCommand();
            sqlite_cmd.CommandText = "select nome_traducao from traducao where cod_traducao = '" + this.codTraducao + "';";
            sqlite_datareader = sqlite_cmd.ExecuteReader();

            if (sqlite_datareader.Read()) // Read() returns true if there is still a result line to read
                this.traducao = (string)sqlite_datareader["nome_traducao"];
            else
                this.traducao = "Tradução '"+this.codTraducao+"' não localizada.";

            // -------------------

            sqlite_cmd = sqlite_conn.CreateCommand();
            sqlite_cmd.CommandText = "select nome from livro where abrev = '" + this.codLivro + "';";
            sqlite_datareader = sqlite_cmd.ExecuteReader();

            if (sqlite_datareader.Read()) // Read() returns true if there is still a result line to read
                this.livro = (string)sqlite_datareader["nome"];
            else
                this.livro = "Livro '" + this.codLivro + "' não localizado.";

            // -------------------

            sqlite_cmd = sqlite_conn.CreateCommand();
            sqlite_cmd.CommandText = "select texto from versiculo where abrev = '" + this.codLivro + "' and cod_traducao = '" + this.codTraducao +"' and capitulo = " + this.capitulo +" and versiculo = " + this.versiculo +";";
            sqlite_datareader = sqlite_cmd.ExecuteReader();

            if (sqlite_datareader.Read()) // Read() returns true if there is still a result line to read
                this.texto = (string)sqlite_datareader["texto"];
            else
                this.texto = "Versículo não localizado.";

            // -------------------
            
            sqlite_conn.Close();
        }

        public BibliaItem ProximoVersiculo()
        {
            SQLiteConnection sqlite_conn = conectaBiblia();
            SQLiteCommand sqlite_cmd;
            SQLiteDataReader sqlite_datareader;
            BibliaItem novoVersiculo;

            // -------------------
            // Tenta o proximo verisulo

            this.navegarVersiculo += 1;

            sqlite_cmd = sqlite_conn.CreateCommand();
            sqlite_cmd.CommandText = "select texto from versiculo where abrev = '" + this.codLivro + "' and cod_traducao = '" + this.codTraducao + "' and capitulo = " + this.navegarCapitulo + " and versiculo = " + this.navegarVersiculo + ";";
            sqlite_datareader = sqlite_cmd.ExecuteReader();

            if (sqlite_datareader.Read()) // Read() returns true if there is still a result line to read
            {
                sqlite_conn.Close();

                novoVersiculo = new BibliaItem(codTraducao,codLivro,navegarCapitulo,navegarVersiculo);
                return novoVersiculo;
            }

            // -------------------
            // Se não der tenta o proximo capitulo

            this.navegarVersiculo = 1;
            this.navegarCapitulo += 1;

            sqlite_cmd = sqlite_conn.CreateCommand();
            sqlite_cmd.CommandText = "select texto from versiculo where abrev = '" + this.codLivro + "' and cod_traducao = '" + this.codTraducao + "' and capitulo = " + this.navegarCapitulo + " and versiculo = " + this.navegarVersiculo + ";";
            sqlite_datareader = sqlite_cmd.ExecuteReader();

            if (sqlite_datareader.Read()) // Read() returns true if there is still a result line to read
            {
                sqlite_conn.Close();

                novoVersiculo = new BibliaItem(codTraducao, codLivro, navegarCapitulo, navegarVersiculo);
                return novoVersiculo;
            }

            // -------------------
            // Se acabou o livro, interrompe o processo

            return null;

        }

        public static DataTable RetornaListaTraducao()
        {
            DataTable dados = new DataTable("Traducao");
            dados.Columns.Add("codigo", typeof(string));
            dados.Columns.Add("traducao", typeof(string));

            SQLiteConnection sqlite_conn = conectaBiblia();
            SQLiteCommand sqlite_cmd;
            SQLiteDataReader sqlite_datareader;

            sqlite_cmd = sqlite_conn.CreateCommand();
            sqlite_cmd.CommandText = "select cod_traducao, nome_traducao from traducao where ativo = 1 order by ordem;";
            sqlite_datareader = sqlite_cmd.ExecuteReader();

            while (sqlite_datareader.Read())
                dados.Rows.Add(new object[] { (string)sqlite_datareader["cod_traducao"], (string)sqlite_datareader["nome_traducao"] });

            return dados;
        }

        public static DataTable RetornaListaLivro()
        {
            DataTable dados = new DataTable("Traducao");
            dados.Columns.Add("codigo", typeof(string));
            dados.Columns.Add("livro", typeof(string));

            SQLiteConnection sqlite_conn = conectaBiblia();
            SQLiteCommand sqlite_cmd;
            SQLiteDataReader sqlite_datareader;

            sqlite_cmd = sqlite_conn.CreateCommand();
            sqlite_cmd.CommandText = "select abrev, nome from livro order by ordem;";
            sqlite_datareader = sqlite_cmd.ExecuteReader();

            while (sqlite_datareader.Read())
                dados.Rows.Add(new object[] { (string)sqlite_datareader["abrev"], (string)sqlite_datareader["nome"] });

            return dados;
        }

        public string Versiculo
        {
            get
            {
                return this.texto;
            }
        }

        public int Incluir
        {
            get
            {
                return this.incluir;
            }
        }

        public string Referencia
        {
            get
            {
                string referencia = livro + " " + capitulo + ":" + versiculo;

                return referencia;
            }
        }

        public string ReferenciaCompleta
        {
            get
            {
                string referencia = this.Referencia;

                if (incluir > 0)
                    referencia += " (+" + incluir + ")";

                return referencia;
            }
        }

        public override String ToString()
        {
            return "Versiculo -> " + this.Referencia;
        }
    }

    public class ChaveValorItem
    {
        public string chave;
        public string valor;

        public ChaveValorItem(string c, string v)
        {
            chave = c;
            valor = v;
        }

        public override String ToString()
        {
            return valor;
        }
    }

    public class Item
    {
        private object item;
        public TipoItem tipo;

        public Item(TipoItem t, object i)
        {
            tipo = t;
            item = i;
        }

        public override string ToString()
        {
 	        return item.ToString();
        }

        public HinoItem GetItemHino()
        {
            return (HinoItem)item;
        }

        public ArquivoItem GetItemArquivo()
        {
            return (ArquivoItem)item;
        }

        public AvisoItem GetItemAviso()
        {
            return (AvisoItem)item;
        }

        public TemplateItem GetItemTemplate()
        {        
            return (TemplateItem)item;
        }

        public BibliaItem GetItemBiblia()
        {
            return (BibliaItem)item;
        }

        public ChaveValorItem GetItemChaveValor()
        {
            return (ChaveValorItem)item;
        }

        public object GetItem()
        {
            return item;
        }
        
    }

}
