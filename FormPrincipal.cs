using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using Finisar.SQLite;
using System.Threading;
using System.Xml;
using System.Xml.Linq;

namespace Apresentacao
{
    public partial class FormPrincipal : Form
    {
        private const int TIPO_VAZIO = 0;
        private const int TIPO_AVISO = 1;
        private const int TIPO_LOUVOR = 2;
        private const int TIPO_VERSICULO = 3;

        private const string XAML_SET_IMAGE = "SetImage";
        private const string XAML_SET_TEXT = "SetText";

        private const string XAML_IMG_BIBLIA = "ImgBiblia";
        private const string XAML_IMG_REFERENCIA = "ImgReferencia";
        private const string XAML_TXT_REFERENCIA = "TxtReferencia";
        private const string XAML_TXT_VERSICULO = "TxtVersiculo";
        private const string XAML_TXT_LOUVOR = "TxtLouvor";

        private string XAML_INPUT_ID = "";
        private string XAML_IMG_BIBLIA_LOCATION = "";
        private string XAML_IMG_REFERENCIA_LOCATION = "";

        private int lastIdShowedInTooltip;
        private string groupBoxHinosName;

        String strCurrDir;
        TipoItem tipoItemSelecionado;
        List<BibliaItem> historicoBiblia;
        List<HinoItem> baseHinos;

        System.ComponentModel.DoWorkEventHandler DoWork_MostrarAviso;
        System.ComponentModel.DoWorkEventHandler DoWork_MostrarBiblia;
        System.ComponentModel.DoWorkEventHandler DoWork_MostrarSelecionados;

        public delegate void InvokeDelegate();
        private delegate int getListBoxSlidesCountCallback();
        private delegate TemplateItem getToolStripComboBoxTemplateSelectedItemCallback();
        private delegate void setObjectTextCallback(object label, string tipo, string text);
        private delegate void setListBoxSlidesAddItemCallback(string item);
        private delegate void setListBoxDisponivelAddItemCallback(HinoItem item);
        private delegate void setToolStripProgressBarMostrarStartCallback(int max); 
        
        private System.Windows.Forms.Timer myTimerAtualizaLista;

        static private bool piscaPisca = true;
        
        public FormPrincipal()
        {
            InitializeComponent();

            setObjectText(toolStripStatusLabelStatus, "toolStrip", "Inicializando ...");
            
            groupBoxHinosName = groupBoxHinos.Text;

            historicoBiblia = new List<BibliaItem>();
            baseHinos = new List<HinoItem>();

            historicoBiblia.Add(null);
            historicoBiblia.Add(null);
            historicoBiblia.Add(null);
            historicoBiblia.Add(null);
            historicoBiblia.Add(null);

            lastIdShowedInTooltip = 0;

            DoWork_MostrarAviso = new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker_DoWork_MostrarAviso);
            DoWork_MostrarBiblia = new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker_DoWork_MostrarBiblia);
            DoWork_MostrarSelecionados = new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker_DoWork_MostrarSelecionados);

            myTimerAtualizaLista = new System.Windows.Forms.Timer();
            myTimerAtualizaLista.Tick += new EventHandler(AtualizarListaHinosTimer);
            myTimerAtualizaLista.Interval = 300;

            textBoxBuscar.Focus();
        }

        private string EncodeTo64(string toEncode)
        {

            byte[] toEncodeAsBytes

                  = System.Text.Encoding.Unicode.GetBytes(toEncode);

            string returnValue

                  = System.Convert.ToBase64String(toEncodeAsBytes);

            return returnValue;

        }

        private string DecodeFrom64(string encodedData)
        {

            byte[] encodedDataAsBytes

                = System.Convert.FromBase64String(encodedData);

            string returnValue =

               System.Text.Encoding.Unicode.GetString(encodedDataAsBytes);

            return returnValue;

        }

        public static Image CreateNonIndexedImage(string path)
        {
            using (var sourceImage = Image.FromFile(path))
            {
                var targetImage = new Bitmap(sourceImage.Width, sourceImage.Height, PixelFormat.Format32bppArgb);
                using (var canvas = Graphics.FromImage(targetImage))
                {
                    canvas.DrawImageUnscaled(sourceImage, 0, 0);
                }
                return targetImage;
            }
        }

        private void FormPrincipal_Load(object sender, EventArgs e)
        {
            strCurrDir = System.IO.Directory.GetCurrentDirectory();
            tipoItemSelecionado = TipoItem.Nenhum;

            // Atualiza desenho da tela
            visaoToolStripMenuItem_CheckStateChanged(null, null);

            // Popular com os templates

            int countTemplates = 0;
            string[] pastas = System.IO.Directory.GetDirectories(strCurrDir + "\\Templates");

            for (int i = 0; i < pastas.Length; i++)
            {
                System.IO.DirectoryInfo templateFolder = new System.IO.DirectoryInfo(pastas[i]);
                string nomePasta = templateFolder.Name.Substring(2).Trim();
                
                System.IO.FileInfo[] templates = templateFolder.GetFiles("*.potx");
                for (int j = 0; j < templates.Length; j++)
                {
                    countTemplates++;
                    toolStripComboBoxTemplate.Items.Add(new TemplateItem(templates[j].Name, templates[j].FullName));
                }
            }
            if (countTemplates > 0)
                toolStripComboBoxTemplate.SelectedIndex = 0;

            // Popular com os avisos

            System.IO.DirectoryInfo avisoFolder = new System.IO.DirectoryInfo(strCurrDir + "\\Avisos");
            System.IO.FileInfo[] avisos = avisoFolder.GetFiles("*.txt");
            for (int i = 0; i < avisos.Length; i++)
                comboBoxAvisoTitulo.Items.Add(new AvisoItem(avisos[i].Name, null, avisos[i].FullName));

            comboBoxAvisoTitulo.SelectedIndex = -1;

            // Popular com os hinos

            buttonAtualizarBaseHinos_Click(null, null);

            // Popular os campos da Biblia

            AtualizarListaBibliaTraducao();

            AtualizarListaBibliaLivros();

            // Ajusta disponibilização dos botões de imagem

            checkBoxGerarImagem_CheckedChanged(null, null);
            textBoxLocalSalvarImagem.Text = System.IO.Directory.GetCurrentDirectory() + "\\vMix\\dynamicimage.png";

            // Atualiza programa para integração com vMix

            sToolStripMenuItem_Click(null, null);

            setObjectText(toolStripStatusLabelStatus, "toolStrip", "Sistema pronto para uso.");
        }

        #region EVENTOS

        private void listBoxSelecionado_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            buttonRemover.PerformClick();
        }

        private void listBoxSelecionado_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBoxSelecionado.SelectedIndex != -1)
            {
                AtivarObjetoSelecionado();
            }
        }

        private void listBoxSelecionado_KeyDown(object sender, KeyEventArgs e)
        {
            if (listBoxSelecionado.SelectedIndex != -1)
            {
                if (e.Shift && e.KeyCode == Keys.Down)
                {
                    MessageBox.Show("Desce");
                }
                else if (e.Shift && e.KeyCode == Keys.Up)
                {
                    MessageBox.Show("Sobe");
                }
            }

        }

        private void listBoxDisponivel_MouseMove(object sender, MouseEventArgs e)
        {
            int idx = listBoxDisponivel.IndexFromPoint(e.Location);
            //int idx = Convert.ToInt32(Math.Truncate(Convert.ToDouble(e.Y / listBoxDisponivel.ItemHeight)));
            if (idx >= 0 && idx < listBoxDisponivel.Items.Count)
            {
                if (idx != lastIdShowedInTooltip)
                {
                    lastIdShowedInTooltip = idx;

                    HinoItem hino = (HinoItem)listBoxDisponivel.Items[idx];
                    if (apresentaçãoSelecionadaToolStripMenuItem.Checked)
                    {
                        textBoxAtivo.Text = hino.letra;
                    }

                    if (tooltipToolStripMenuItem.Checked)
                    {
                        toolTipLetra.ToolTipTitle = hino.nome;
                        toolTipLetra.SetToolTip(this.listBoxDisponivel, hino.letra);
                        toolTipLetra.Active = true;
                        
                    }
                }
            }
            else
            {
                if (apresentaçãoSelecionadaToolStripMenuItem.Checked)
                {
                    textBoxAtivo.Text = "";
                }
                
                if(tooltipToolStripMenuItem.Checked)
                {
                    toolTipLetra.Active = false;
                }
            }
        }

        private void listBoxDisponivel_MouseLeave(object sender, EventArgs e)
        {
            toolTipLetra.Active = true;
        }

        private void listBoxDisponivel_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            SelecionarHinos();
        }

        private void comboBoxAvisoTitulo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxAvisoTitulo.SelectedIndex != -1)
            {
                String linha;
                AvisoItem aviso = (AvisoItem)comboBoxAvisoTitulo.SelectedItem;
                System.IO.StreamReader leitor = new System.IO.StreamReader(aviso.caminho, Encoding.UTF8);

                textBoxAvisos.Text = "";

                while ((linha = leitor.ReadLine()) != null)
                {
                    textBoxAvisos.Text += linha + "\r\n";
                }

                leitor.Close();

            }
        }

        private void buttonIncluirHino_Click(object sender, EventArgs e)
        {
            SelecionarHinos();
        }

        private void buttonIncluirAviso_Click(object sender, EventArgs e)
        {
            if (comboBoxAvisoTitulo.Text.Trim().Length == 0)
            {
                MessageBox.Show("Digite um título para o aviso antes de incluí-lo.", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                listBoxSelecionado.Items.Add(new Item(TipoItem.Aviso, new AvisoItem(comboBoxAvisoTitulo.Text, textBoxAvisos.Text)));
            }
        }

        private void buttonLimparAviso_Click(object sender, EventArgs e)
        {
            textBoxAvisos.Text = "";
            comboBoxAvisoTitulo.SelectedIndex = -1; ;
        }

        private void buttonRemover_Click(object sender, EventArgs e)
        {
            if (listBoxSelecionado.SelectedIndex != -1)
            {
                int selecionado = listBoxSelecionado.SelectedIndex;
                listBoxSelecionado.Items.RemoveAt(selecionado);

                if (selecionado >= listBoxSelecionado.Items.Count)
                    listBoxSelecionado.SelectedIndex = selecionado - 1;
                else
                    listBoxSelecionado.SelectedIndex = selecionado;

                if (listBoxSelecionado.Items.Count == 0)
                {
                    tipoItemSelecionado = TipoItem.Nenhum;
                    textBoxAtivo.Text = "";
                    buttonSalvarHino.Enabled = false;
                }
            }
            else if (listBoxSelecionado.Items.Count > 0)
            {
                if (MessageBox.Show("Deseja remover todos os itens da lista?", "Atenção", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                {
                    setListBoxSelecionadoClear();

                    tipoItemSelecionado = TipoItem.Nenhum;
                    textBoxAtivo.Text = "";
                    buttonSalvarHino.Enabled = false;
                }
            }
        }

        private void buttonNovoHino_Click(object sender, EventArgs e)
        {
            listBoxSelecionado.SelectedIndex = -1;
            tipoItemSelecionado = TipoItem.NovoHino;
            textBoxAtivo.Text = "";
            buttonSalvarHino.Text = "Salvar Hino";
            buttonSalvarHino.Enabled = true;
            buttonRemover.Enabled = false;
        }

        private void buttonSalvarHino_Click(object sender, EventArgs e)
        {
            switch (tipoItemSelecionado)
            {
                case TipoItem.Hino:

                    // Recuepra hino selecionado
                    HinoItem hino = ((Item)listBoxSelecionado.SelectedItem).GetItemHino();
                    int idx = listBoxDisponivel.Items.IndexOf(hino);
                    
                    // Salva alterações no disco
                    System.IO.StreamWriter escritaHino = new System.IO.StreamWriter(hino.caminho, false, Encoding.UTF8);
                    escritaHino.Write(textBoxAtivo.Text);
                    escritaHino.Close();
                    
                    // Salva alterações no hino selecionado
                    hino.letra = textBoxAtivo.Text;
                    listBoxSelecionado.Items[listBoxSelecionado.SelectedIndex] = new Item(TipoItem.Hino, hino);

                    // Salva alterações do hino na lista completa
                    listBoxDisponivel.Items[idx] = hino;
                    
                    // Ajusta botões
                    buttonRemover.Enabled = true;
                    break;

                case TipoItem.NovoHino:

                    saveFileDialog1.InitialDirectory = strCurrDir + "\\Hinos";
                    saveFileDialog1.DefaultExt = "*.txt";
                    saveFileDialog1.Filter = "Arquivo Texto|*.txt";
                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        System.IO.StreamWriter escritaNovoHino = new System.IO.StreamWriter(saveFileDialog1.FileName, false, Encoding.UTF8);
                        escritaNovoHino.Write(textBoxAtivo.Text);
                        escritaNovoHino.Close();

                        tipoItemSelecionado = TipoItem.Nenhum;
                        buttonRemover.Enabled = true;
                    }
                    break;

                case TipoItem.Aviso:
                    ((Item)listBoxSelecionado.SelectedItem).GetItemAviso().texto = textBoxAtivo.Text;
                    break;

                case TipoItem.Nenhum:
                    MessageBox.Show("Não há item selecionado", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    break;

                default:
                    MessageBox.Show("Não é possível salvar este tipo de item", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    break;
            }
        }

        private void buttonBibliaAdicionar_Click(object sender, EventArgs e)
        {
            if (validaCamposBiblia())
            {
                try
                {
                    string traducao = ((Item)comboBoxBibliaTraducao.SelectedItem).GetItemChaveValor().chave;
                    string livro = ((Item)comboBoxBibliaLivro.SelectedItem).GetItemChaveValor().chave;
                    int capitulo = Convert.ToInt32(textBoxBibliaCapitulo.Text);
                    int versiculo = Convert.ToInt32(textBoxBibliaVersiculo.Text);

                    int adicional = 0;
                    if (textBoxBibliaVersiculoAdicional.Text != "")
                    {
                        adicional = Convert.ToInt32(textBoxBibliaVersiculoAdicional.Text);
                    }

                    BibliaItem biblia = new BibliaItem(traducao, livro, capitulo, versiculo, adicional);
                    listBoxSelecionado.Items.Add(new Item(TipoItem.Biblia, biblia));

                    bibliaIncluirAdicional(biblia);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void bibliaIncluirAdicional( BibliaItem item )
        {
            if (item.Incluir > 0)
            {
                try
                {
                    item.ReiniciarNavegacao();
                    for (int i = 0; i < item.Incluir; i++)
                    {
                        BibliaItem novoVersiculo = item.ProximoVersiculo();
                        if (novoVersiculo != null)
                            listBoxSelecionado.Items.Add(new Item(TipoItem.Biblia, novoVersiculo));
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

            }
        }
        
        private void textBoxBuscar_TextChanged(object sender, EventArgs e)
        {
            if (myTimerAtualizaLista.Enabled) myTimerAtualizaLista.Stop();
            myTimerAtualizaLista.Start();
        }

        #endregion

        #region FUNÇÕES

        private void AtualizarListaHinosTimer(object sender, EventArgs e)
        {
            if (this.InvokeRequired)
            {
                /* Not on UI thread, reenter there... */
                this.BeginInvoke(new EventHandler(AtualizarListaHinosTimer), sender, e);
            }
            else
            {
                lock (myTimerAtualizaLista)
                {
                    /* only work when this is no reentry while we are already working */
                    if (this.myTimerAtualizaLista.Enabled)
                    {
                        this.myTimerAtualizaLista.Stop();
                        string titulo = "";
                        string letra = "";

                        // se os dois campos de busca estiverem vazio, então busca tudo
                        if (textBoxBuscar.Text.Length == 0 && textBoxBuscarLetra.Text.Length == 0)
                        {
                            titulo = null;
                            letra = null;
                        }
                        // Se o campo de busca por titulo estiver vazio, busca só por letra
                        else if (textBoxBuscar.Text.Length == 0)
                        {
                            titulo = null;
                            letra = textBoxBuscarLetra.Text;
                        }
                        // Se o campo de busca por letra estiver vazio, busca só por titulo
                        else if (textBoxBuscarLetra.Text.Length == 0)
                        {
                            titulo = textBoxBuscar.Text;
                            letra = null;
                        }
                        // Se os dois campos estiverem preenchidos, busca pelos dois
                        else
                        {
                            titulo = textBoxBuscar.Text;
                            letra = textBoxBuscarLetra.Text;
                        }

                        AtualizarListaHinos(titulo, letra);
                    }
                }
            }
        }

        public string RemoverAcentos(string texto)
        {
            if (string.IsNullOrEmpty(texto))
                return String.Empty;
            else
            {
                byte[] bytes = System.Text.Encoding.GetEncoding("iso-8859-8").GetBytes(texto);
                return System.Text.Encoding.UTF8.GetString(bytes);
            }
        }

        private void AtualizaBaseHinos()
        {
            setObjectText(groupBoxHinos, "groupBox", groupBoxHinosName + " (atualizando ...)");
            setObjectText(toolStripStatusLabelStatus, "toolStrip", "Carregando hinos ...");

            baseHinos.Clear();
            setListBoxDisponivelClear();

            System.IO.DirectoryInfo hinosFolder = new System.IO.DirectoryInfo(strCurrDir + "\\Hinos");
            System.IO.FileInfo[] hinos = hinosFolder.GetFiles("*.txt");

            setToolStripProgressBarMostrarStart(hinos.Length);
            for (int i = 0; i < hinos.Length; i++)
            {
                setToolStripProgressBarMostrarIncrement();

                HinoItem novo = new HinoItem(hinos[i].Name, hinos[i].FullName);
                baseHinos.Add(novo);
                setListBoxDisponivelAddItem(novo);

                setObjectText(toolStripStatusLabelStatus, "toolStrip", "Carregando hinos ... " + (i + 1) + "/" + hinos.Length);
            }
            setToolStripProgressBarMostrarStop();
            setObjectText(toolStripStatusLabelStatus, "toolStrip", "Sistema pronto para uso.");

            setObjectText(groupBoxHinos, "groupBox", groupBoxHinosName + " (" + hinos.Length + " hinos)");
            
            string titulo = textBoxBuscar.Text.Trim();
            string letra = textBoxBuscarLetra.Text.Trim();
            int buscar = titulo.Length + letra.Length;

            if (buscar > 0) AtualizarListaHinos(titulo, letra); 
        }

        private void AtualizarListaHinos(string titulo, string letra)
        {
            setObjectText(groupBoxHinos, "groupBox", groupBoxHinosName + " (filtrando ...)");

            lastIdShowedInTooltip = 0;
            setListBoxDisponivelClear();

            string letraTratada = RemoverAcentos(letra);
            string tituloTratado = RemoverAcentos(titulo);

            List<HinoItem> resultado = baseHinos.FindAll(delegate(HinoItem hino) {
                bool localizou;

                // Analisa o texto
                if (letraTratada.Length > 0)
                {
                    localizou = (RemoverAcentos(hino.letra.ToUpper()).IndexOf(letraTratada.ToUpper()) != -1);
                }
                else
                {
                    localizou = true;
                }

                // Analisa o titulo
                if (localizou && tituloTratado.Length > 0)
                {
                    localizou = (RemoverAcentos(hino.nome.ToUpper()).IndexOf(tituloTratado.ToUpper()) != -1);
                }

                return localizou;
            });

            for (int i = 0; i < resultado.Count; i++)
            {
                setListBoxDisponivelAddItem(resultado[i]);
            }

            setObjectText(groupBoxHinos, "groupBox", groupBoxHinosName + " (" + resultado.Count + " hinos)");

        }
                
        private void AtualizarListaBibliaTraducao()
        {
            DataTable dados = BibliaItem.RetornaListaTraducao();

            for (int i = 0; i < dados.Rows.Count; i++)
            {
                string codigo = (string)dados.Rows[i][0];
                string nome = (string)dados.Rows[i][1];
                Item item = new Item(TipoItem.ChaveValor, new ChaveValorItem(codigo, nome));
                comboBoxBibliaTraducao.Items.Add(item); 
            }
            comboBoxBibliaTraducao.SelectedIndex = 0;
        }

        private void AtualizarListaBibliaLivros()
        {
            DataTable dados = BibliaItem.RetornaListaLivro();

            for(int i=0; i<dados.Rows.Count; i++)
            {
                string codigo = (string)dados.Rows[i][0];
                string nome = (string)dados.Rows[i][1];
                Item item = new Item(TipoItem.ChaveValor, new ChaveValorItem(codigo, nome));
                comboBoxBibliaLivro.Items.Add(item);
            }
            comboBoxBibliaLivro.SelectedIndex = 0;
        }

        private void SelecionarHinos()
        {
            ListBox.SelectedObjectCollection hinos = listBoxDisponivel.SelectedItems;

            for (int i = 0; i < hinos.Count; i++)
            {
                listBoxSelecionado.Items.Add(new Item(TipoItem.Hino, hinos[i]));
            }
        }

        private void AtivarObjetoSelecionado()
        {
            if (listBoxSelecionado.SelectedItem != null)
            {
                Item item = (Item)listBoxSelecionado.SelectedItem;
                tipoItemSelecionado = item.tipo;

                switch (item.tipo)
                {
                    case TipoItem.Hino:
                        String linha;
                        HinoItem hino = item.GetItemHino();

                        vMixSetupXAML(TIPO_LOUVOR,hino.nome.ToUpper());
                        
                        textBoxAtivo.Text = "";

                        System.IO.StringReader leitor = new System.IO.StringReader(hino.letra);
                        while ((linha = leitor.ReadLine()) != null)
                        {
                            textBoxAtivo.Text += linha + "\r\n";
                        }

                        leitor.Close();

                        buttonSalvarHino.Text = "Salvar Hino";
                        buttonSalvarHino.Enabled = true;
                        buttonRenomear.Enabled = true;

                        setObjectText(toolStripStatusLabelStatus, "toolStrip", "Ativado Hino: " + hino.nome);

                        break;

                    case TipoItem.Aviso:
                        AvisoItem aviso = item.GetItemAviso();
                        textBoxAtivo.Text = aviso.texto;

                        vMixSetupXAML(TIPO_AVISO, aviso.texto);

                        buttonSalvarHino.Text = "Salvar Aviso";
                        buttonSalvarHino.Enabled = true;
                        buttonRenomear.Enabled = true;

                        setObjectText(toolStripStatusLabelStatus, "toolStrip", "Ativado Aviso: " + aviso.texto.Substring(0, 10) + " ...");

                        break;

                    case TipoItem.Biblia:
                        BibliaItem biblia = item.GetItemBiblia();

                        textBoxAtivo.Text = biblia.Referencia + "\r\n\r\n" + biblia.Versiculo;

                        vMixSetupXAML(TIPO_VERSICULO, biblia.Versiculo, biblia.Referencia);

                        setObjectText(toolStripStatusLabelStatus, "toolStrip", "Ativado Versículo: " + biblia.Referencia);
                        
                        break;

                    case TipoItem.Arquivo:
                        vMixSetupXAML(TIPO_VAZIO);
                        
                        textBoxAtivo.Text = "O item selecionado é do tipo Arquivo.\r\nNão há como visualizar aqui.";
                        buttonSalvarHino.Enabled = false;
                        buttonRenomear.Enabled = false;

                        setObjectText(toolStripStatusLabelStatus, "toolStrip", "Ativado Arquivo");

                        break;

                    default:
                        vMixSetupXAML(TIPO_VAZIO);
                        
                        textBoxAtivo.Text = "Item selecionado inválido.";
                        buttonSalvarHino.Enabled = false;
                        buttonRenomear.Enabled = false;

                        setObjectText(toolStripStatusLabelStatus, "toolStrip", "Erro na Ativação");

                        break;
                }
            }
        }

        private void ShowPresentation(bool salvar)
        {
            Item item;
            int count = 1;
            PowerPoint.Application objApp = null;
            PowerPoint.Slides objSlides = IniciarApresentacao(ref objApp);

            setObjectText(toolStripStatusLabelStatus, "toolStrip", "Montando slides ...");

            if (objSlides != null)
            {
                // Google it
                // PowerPoint.Application.SlideShowWindows

                // Loop nos hinos

                setToolStripProgressBarMostrarStart(listBoxSelecionado.Items.Count);

                for (int idx = 0; idx < listBoxSelecionado.Items.Count; idx++)
                {
                    item = (Item)listBoxSelecionado.Items[idx];
                    PreparaSlide(ref count, item, ref objSlides);
                    setToolStripProgressBarMostrarIncrement();
                    setObjectText(toolStripStatusLabelStatus, "toolStrip", "Montando slides ... item " + (idx + 1));
                }

                setToolStripProgressBarMostrarStop();

                FinalizarApresentacao(objApp, objSlides, salvar);

            }

            GC.Collect();
        }

        private PowerPoint.Slides IniciarApresentacao(ref PowerPoint.Application objApp)
        {            
            //Criar uma nova apresentação baseada no template
            objApp = new PowerPoint.Application();

            TemplateItem template = getToolStripComboBoxTemplateSelectedItem();

            if (template == null)
            {
                MessageBox.Show("Selecione um template para exibir.");
                return null;
            }

            String strTemplate = template.caminho;

            PowerPoint.Presentations objPresSet;
            PowerPoint._Presentation objPres;
            PowerPoint.Slides objSlides;
            //PowerPoint._Slide objSlide;
            //PowerPoint.Shapes objShapes;
            //PowerPoint.Shape objShape;

            //objApp.Assistant.On = false;
            objPresSet = objApp.Presentations;
            objPres = objPresSet.Open(strTemplate, MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoTrue);

            if (checkBoxGerarImagem.Checked)
            {
                objApp.WindowState = Microsoft.Office.Interop.PowerPoint.PpWindowState.ppWindowMinimized;
            }
            else
            {
                objApp.WindowState = Microsoft.Office.Interop.PowerPoint.PpWindowState.ppWindowMaximized;
            }

            //objPres.SlideShowSettings.AdvanceMode = Microsoft.Office.Interop.PowerPoint.PpSlideShowAdvanceMode.ppSlideShowUseSlideTimings;
            //objPres.SlideShowSettings.ShowType = Microsoft.Office.Interop.PowerPoint.PpSlideShowType.ppShowTypeKiosk;
            objSlides = objPres.Slides;

            return objSlides;
        }

        private void PreparaSlide(ref int count, Item item, ref PowerPoint.Slides objSlides)
        {
            PowerPoint.TextRange objTextRng;
            PowerPoint._Slide objSlide;

            if (count == 1) 
                objSlide = objSlides.Add(count++, PowerPoint.PpSlideLayout.ppLayoutBlank);

            switch (item.tipo)
            {
                case TipoItem.Hino:

                    if (count > 2)
                        objSlide = objSlides.Add(count++, PowerPoint.PpSlideLayout.ppLayoutBlank);

                    HinoItem hino = item.GetItemHino();

                    System.IO.StringReader strReader = new System.IO.StringReader(hino.letra);
                    String blocoHino = "";
                    String linhaHino;

                    objSlide = objSlides.Add(count++, PowerPoint.PpSlideLayout.ppLayoutTitle);
                    objTextRng = objSlide.Shapes.Title.TextFrame.TextRange;
                    objTextRng.Text = hino.nome;

                    while ((linhaHino = strReader.ReadLine()) != null)
                    {
                        if (linhaHino.Length == 0)
                        {
                            objSlide = objSlides.Add(count++, PowerPoint.PpSlideLayout.ppLayoutText);
                            objTextRng = objSlide.Shapes[2].TextFrame.TextRange;
                            objTextRng.Text = blocoHino;
                            objTextRng.Font.Size = Convert.ToInt32("32"); //textBoxTamanhoFonte.Text);
                            objTextRng.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletNone;

                            blocoHino = "";
                        }
                        else
                        {
                            blocoHino += linhaHino + "\r\n";
                        }
                    }
                    strReader.Close();

                    if (blocoHino.Length > 0)
                    {
                        objSlide = objSlides.Add(count++, PowerPoint.PpSlideLayout.ppLayoutText);
                        objTextRng = objSlide.Shapes[2].TextFrame.TextRange;
                        objTextRng.Text = blocoHino;
                        objTextRng.Font.Size = Convert.ToInt32("32"); //textBoxTamanhoFonte.Text);
                        objTextRng.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletNone;
                    }

                    break;

                case TipoItem.Aviso:

                    if (count > 2)
                        objSlide = objSlides.Add(count++, PowerPoint.PpSlideLayout.ppLayoutBlank);

                    AvisoItem aviso = item.GetItemAviso();

                    System.IO.StringReader leitorAviso = new System.IO.StringReader(aviso.texto);
                    String blocoAviso = "";
                    String linhaAviso;

                    while ((linhaAviso = leitorAviso.ReadLine()) != null)
                    {
                        if (linhaAviso.Length == 0)
                        {
                            objSlide = objSlides.Add(count++, PowerPoint.PpSlideLayout.ppLayoutContentWithCaption);
                            objTextRng = objSlide.Shapes.Title.TextFrame.TextRange;
                            objTextRng.Text = aviso.nome;

                            objTextRng = objSlide.Shapes[3].TextFrame.TextRange;
                            objTextRng.Text = blocoAviso;
                            objTextRng.Font.Size = Convert.ToInt32("32"); //textBoxTamanhoFonte.Text);
                            objTextRng.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletNone;

                            blocoAviso = "";
                        }
                        else
                        {
                            blocoAviso += linhaAviso + "\r\n";
                        }
                    }
                    leitorAviso.Close();

                    if (blocoAviso.Length > 0)
                    {
                        objSlide = objSlides.Add(count++, PowerPoint.PpSlideLayout.ppLayoutContentWithCaption);
                        objTextRng = objSlide.Shapes[1].TextFrame.TextRange;
                        objTextRng.Text = aviso.nome;

                        objTextRng = objSlide.Shapes[3].TextFrame.TextRange;
                        objTextRng.Text = blocoAviso;
                        objTextRng.Font.Size = Convert.ToInt32("32"); //textBoxTamanhoFonte.Text);
                        objTextRng.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletNone;
                    }

                    break;

                case TipoItem.Biblia:

                    BibliaItem biblia = item.GetItemBiblia();

                    objSlide = objSlides.Add(count++, PowerPoint.PpSlideLayout.ppLayoutTwoObjects);
                    objTextRng = objSlide.Shapes[2].TextFrame.TextRange;
                    objTextRng.Text = biblia.Versiculo;
                    objTextRng.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletNone;

                    objTextRng = objSlide.Shapes[3].TextFrame.TextRange;
                    objTextRng.Text = biblia.Referencia;
                    objTextRng.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletNone;

                    break;

                case TipoItem.Arquivo:

                    ArquivoItem arquivo = item.GetItemArquivo();
                    int resp = objSlides.InsertFromFile(arquivo.caminho, count++, 1, 1);

                    break;

                default:
                    break;
            }
        }

        private void FinalizarApresentacao(PowerPoint.Application objApp, PowerPoint.Slides objSlides, bool salvar)
        {
            if (objSlides.Count <= 1) return;

            PowerPoint.SlideShowSettings objSSS;
            PowerPoint._Presentation objPres = (PowerPoint._Presentation)objSlides.Parent;
            
            /*
            // Contruindo Slide #1
            // Insere um texto no slide, muda a fonte e coloca uma imagem
            objSlide = objSlides.Add(1, PowerPoint.PpSlideLayout.ppLayoutTitleOnly);
            objTextRng = objSlide.Shapes[1].TextFrame.TextRange;
            objTextRng.Text = "Elisson Santos da Silva";
            objTextRng.Font.Name = "Comic Sans MS";
            objTextRng.Font.Size = 48;
            */

            // Insere o ultimo slide com a imagem de fundo
            objSlides.Add(objSlides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);

            // Modifica a transição de Slides
            for (int i = 2; i <= objSlides.Count; i++)
            {
                objSlides[i].SlideShowTransition.AdvanceOnTime = MsoTriState.msoFalse;
                objSlides[i].SlideShowTransition.EntryEffect = PowerPoint.PpEntryEffect.ppEffectFadeSmoothly;
                objSlides[i].SlideShowTransition.Speed = PowerPoint.PpTransitionSpeed.ppTransitionSpeedFast;
            }
            
            // Roda os Slides
            objSSS = objPres.SlideShowSettings;
            objSSS.StartingSlide = 1;
            objSSS.EndingSlide = objSlides.Count;
            objSSS.LoopUntilStopped = MsoTriState.msoTrue;
            objSSS.RangeType = PowerPoint.PpSlideShowRangeType.ppShowSlideRange;

            if (salvar)
            {
                objPres.SaveAs(saveFileDialog1.FileName, PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
                setObjectText(toolStripStatusLabelStatus, "toolStrip", "Apresentação salva em " + saveFileDialog1.FileName);
            }
            else if (checkBoxGerarImagem.Checked)
            {
                setObjectText(toolStripStatusLabelStatus, "toolStrip", "Criando imagens ...");

                int size = getListBoxSlidesCount();
                for (int i = 0; i < size; i++ )
                {
                    if (i == 0) pictureBoxSlide.Image = Apresentacao.Properties.Resources.FundoSlide;
                    System.IO.File.Delete("slide" + (i + 1).ToString().PadLeft(3, '0') + ".png");
                }

                setListBoxSlidesClear();
                setToolStripProgressBarMostrarStart(objSlides.Count);
                for (int i = 1; i < objSlides.Count; i++)
                {
                    objApp.ActivePresentation.Slides[i].Export(System.IO.Directory.GetCurrentDirectory() + "\\slide" + i.ToString().PadLeft(3, '0') + ".png", "png", 768, 576);
                    setListBoxSlidesAddItem("Slide " + i);
                    setToolStripProgressBarMostrarIncrement();
                }
                setToolStripProgressBarMostrarStop();

                // Fecha a apresentação sem salvar
                objPres.Close();
                objApp.Quit();

                setListBoxSlidesSelectFirst();
                setObjectText(toolStripStatusLabelStatus, "toolStrip", "Lista de imagens criadas");
            }
            else
            {
                setObjectText(toolStripStatusLabelStatus, "toolStrip", "Apresentação em execução");
                objSSS.Run();

                // Espera até que o slideshow termine
                // Se coloca a validação pra encerrar automaticamente 
                //   depois que a apresentação terminar, 
                //   as transições não são disparadas
                //VerificaStatusApresentacao(objApp);
                 
            }
        }

        private void VerificaStatusApresentacao(PowerPoint.Application objApp)
        {
            try
            {
                PowerPoint.SlideShowWindows objSSWs = objApp.SlideShowWindows;
                if (objSSWs.Count >= 1)
                {
                    piscaPisca = !piscaPisca;
                    string pisca = piscaPisca ? "" : "(!)";
                    setObjectText(toolStripStatusLabelStatus, "toolStrip", "Apresentação em execução ... " + pisca);
                    Thread.Sleep(300);

                    Thread thread = new Thread(() => VerificaStatusApresentacao(objApp));
                    thread.IsBackground = true;
                    thread.Start();
                }
                else
                {
                    // Fecha a apresentação sem salvar
                    PowerPoint._Presentation objPres = objApp.Presentations[1];
                    objPres.Close();
                    objApp.Quit();
                    setObjectText(toolStripStatusLabelStatus, "toolStrip", "Apresentação finalizada");
                }
            }
            catch
            {
            }
        }

        #endregion

        private void buttonRenomear_Click(object sender, EventArgs e)
        {
            switch (tipoItemSelecionado)
            {
                case TipoItem.Hino:
                    HinoItem alterarHino = ((Item)listBoxSelecionado.SelectedItem).GetItemHino();
                    Renomear renomarHino = new Renomear(alterarHino.caminho);
                    if (renomarHino.ShowDialog() == DialogResult.OK)
                    {
                        int idx = listBoxDisponivel.Items.IndexOf(alterarHino);

                        // Salva alterações no disco
                        alterarHino.nome = renomarHino.NovoNome;
                        alterarHino.caminho = renomarHino.NovoCaminho;
                        listBoxSelecionado.Items[listBoxSelecionado.SelectedIndex] = new Item(TipoItem.Hino, alterarHino);

                        // Salva alterações do hino na lista completa
                        listBoxDisponivel.Items[idx] = alterarHino;
                    }
                    break;

                case TipoItem.NovoHino:
                    MessageBox.Show("Salve o hino antes de renomeá-lo", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    break;

                case TipoItem.Aviso:
                    MessageBox.Show("Não tem como renomaer um aviso ainda", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    /*
                    AvisoItem alterarAviso = ((Item)listBoxSelecionado.SelectedItem).GetItemAviso();
                    Renomear renomarAviso = new Renomear(alterarAviso.caminho);
                    if (renomarAviso.ShowDialog() == DialogResult.OK)
                    {
                        alterarAviso.nome = renomarAviso.NovoNome;
                        alterarAviso.caminho = renomarAviso.NovoCaminho;
                        listBoxSelecionado.Items[listBoxSelecionado.SelectedIndex] = new Item(TipoItem.Aviso, alterarAviso);
                    }
                    */
                    break;

                case TipoItem.Nenhum:
                    MessageBox.Show("Não há item selecionado", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    break;

                default:
                    MessageBox.Show("Não é possível renomear este tipo de item", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    break;
            }
        }

        private void textBoxBuscar_Enter(object sender, EventArgs e)
        {
            textBoxBuscar.SelectAll();
        }

        private bool validaElemento(XmlElement xmlInput)
        {
            if (xmlInput.SelectNodes("./text[@name='" + XAML_TXT_VERSICULO + "']").Count > 0)
                return true;

            if (xmlInput.SelectNodes("./text[@name='" + XAML_TXT_REFERENCIA + "']").Count > 0)
                return true;

            if (xmlInput.SelectNodes("./text[@name='" + XAML_TXT_LOUVOR + "']").Count > 0)
                return true;

            return false;

        }

        private void vMixSetupXAML(int tipo, String textoPrincipal = null, String textoReferencia = null)
        {
            if (!toolStripMenuItemAtivarSocket.Checked) return;

            String server = toolStripTextBoxIPSocket.Text;
            String url_base = string.Format("http://{0}/api/", server);

            WebRequest request;
            WebResponse response;

            try
            {
                // Carrega XML
                request = WebRequest.Create(url_base);
                response = request.GetResponse();
                String status = ((HttpWebResponse)response).StatusDescription;
                StreamReader reader = new StreamReader(response.GetResponseStream());
                string responseFromServer = reader.ReadToEnd();
                reader.Close();
                response.Close();

                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(responseFromServer);
                foreach(XmlElement xmlInput in xmlDoc.SelectNodes("//input[@type='Xaml']"))
                {
                    if(validaElemento(xmlInput))
                    {
                        string key = xmlInput.GetAttribute("key");
                        if (!XAML_INPUT_ID.Equals(key))
                        {
                            XAML_INPUT_ID = key;

                            foreach (XmlElement xmlImage in xmlInput.SelectNodes("./image"))
                            {
                                string name = xmlImage.GetAttribute("name");

                                if (XAML_IMG_BIBLIA.Equals(name))
                                    XAML_IMG_BIBLIA_LOCATION = xmlImage.InnerText.Replace("file:///", "");

                                if (XAML_IMG_REFERENCIA.Equals(name))
                                    XAML_IMG_REFERENCIA_LOCATION = xmlImage.InnerText.Replace("file:///", "");
                            }

                        }
                        break;
                    }
                }

                if (XAML_INPUT_ID.Length == 0)
                    return;

                switch (tipo)
                {
                    case TIPO_AVISO:
                    case TIPO_LOUVOR:
                        // Limpa imagem da biblia
                        vMixSendAPIMessage(url_base, XAML_SET_IMAGE, XAML_INPUT_ID, XAML_IMG_BIBLIA, "");

                        // Limpa imagem da referencia
                        vMixSendAPIMessage(url_base, XAML_SET_IMAGE, XAML_INPUT_ID, XAML_IMG_REFERENCIA, "");

                        // Limpa texto do versiculo
                        vMixSendAPIMessage(url_base, XAML_SET_TEXT, XAML_INPUT_ID, XAML_TXT_VERSICULO, "");

                        // Limpa texto da referencia
                        vMixSendAPIMessage(url_base, XAML_SET_TEXT, XAML_INPUT_ID, XAML_TXT_REFERENCIA, "");

                        // Seta texto principal
                        vMixSendAPIMessage(url_base, XAML_SET_TEXT, XAML_INPUT_ID, XAML_TXT_LOUVOR, textoPrincipal);

                        break;

                    case TIPO_VERSICULO:
                        // Limpa texto principal
                        vMixSendAPIMessage(url_base, XAML_SET_TEXT, XAML_INPUT_ID, XAML_TXT_LOUVOR, "");

                        // Seta imagem da biblia
                        vMixSendAPIMessage(url_base, XAML_SET_IMAGE, XAML_INPUT_ID, XAML_IMG_BIBLIA, XAML_IMG_BIBLIA_LOCATION);

                        // Seta imagem da referencia
                        vMixSendAPIMessage(url_base, XAML_SET_IMAGE, XAML_INPUT_ID, XAML_IMG_REFERENCIA, XAML_IMG_REFERENCIA_LOCATION);

                        // Seta texto do versiculo
                        vMixSendAPIMessage(url_base, XAML_SET_TEXT, XAML_INPUT_ID, XAML_TXT_VERSICULO, textoPrincipal);

                        // Seta texto da referencia
                        vMixSendAPIMessage(url_base, XAML_SET_TEXT, XAML_INPUT_ID, XAML_TXT_REFERENCIA, textoReferencia);

                        break;

                    case TIPO_VAZIO:
                    default:
                        // Limpa imagem da biblia
                        vMixSendAPIMessage(url_base, XAML_SET_IMAGE, XAML_INPUT_ID, XAML_IMG_BIBLIA, "");

                        // Limpa imagem da referencia
                        vMixSendAPIMessage(url_base, XAML_SET_IMAGE, XAML_INPUT_ID, XAML_IMG_REFERENCIA, "");

                        // Limpa texto do versiculo
                        vMixSendAPIMessage(url_base, XAML_SET_TEXT, XAML_INPUT_ID, XAML_TXT_VERSICULO, "");

                        // Limpa texto da referencia
                        vMixSendAPIMessage(url_base, XAML_SET_TEXT, XAML_INPUT_ID, XAML_TXT_REFERENCIA, "");

                        // Limpa texto principal
                        vMixSendAPIMessage(url_base, XAML_SET_TEXT, XAML_INPUT_ID, XAML_TXT_LOUVOR, "");

                        break;
                }

            }
            catch (Exception e)
            {
                toolStripStatusLabelServer.Text = "Exception: " + e.Message.Replace("\n", "");
            }

        }

        private void vMixSendAPIMessage(String url_base, String function, String input, String name, String value)
        {
            String url = string.Format("{0}?Function={1}&Input={2}&SelectedName={3}&Value={4}", url_base, function, input, name, value);

            WebRequest request = WebRequest.Create(url);
            WebResponse response = request.GetResponse();
            String status = ((HttpWebResponse)response).StatusDescription;
            StreamReader reader = new StreamReader(response.GetResponseStream());
            string responseFromServer = reader.ReadToEnd();
            reader.Close();
            response.Close();
        }

        private void SelecionarLinhaAndEnviar(bool selecionarLinha = false)
        {
            try
            {
                int linhaAtiva = textBoxAtivo.GetLineFromCharIndex(textBoxAtivo.SelectionStart);
                string linha = textBoxAtivo.Lines[linhaAtiva].Trim();
                selecionarLinhaAndAtualizaLabel(linhaAtiva, selecionarLinha);

                if (linha.Length > 0)
                {
                    vMixSetupXAML(TIPO_LOUVOR, linha);
                }
                else
                {
                    vMixSetupXAML(TIPO_VAZIO);
                }
            }
            catch (Exception e)
            {
                vMixSetupXAML(TIPO_VAZIO);
            }
        }

        private void textBoxAtivo_MouseClick(object sender, MouseEventArgs e)
        {
            if (textBoxAtivo.Text.Length > 0)
            {
                SelecionarLinhaAndEnviar();
            }
        }

        private void textBoxAtivo_KeyUp(object sender, KeyEventArgs e)
        {
            if (textBoxAtivo.Text.Length > 0)
            {
                SelecionarLinhaAndEnviar();
            }
        }
        
        private void textBoxSelectAll_Enter(object sender, EventArgs e)
        {
            ((TextBox)sender).SelectAll();
        }

        private void listBoxSlides_MouseClick(object sender, MouseEventArgs e)
        {
            ativarImagemSlide();
        }

        private void listBoxSlides_SelectedIndexChanged(object sender, EventArgs e)
        {
            ativarImagemSlide();
        }

        private void ativarImagemSlide()
        {
            if (listBoxSlides.SelectedIndex >= 0)
            {
                int index = listBoxSlides.SelectedIndex + 1;
                pictureBoxSlide.Image = CreateNonIndexedImage("slide" + index.ToString().PadLeft(3, '0') + ".png");

                if (checkBoxEnviarCopia.Checked)
                {
                    if (!System.IO.File.Exists(textBoxLocalSalvarImagem.Text))
                    {
                        System.IO.FileInfo info = new System.IO.FileInfo(textBoxLocalSalvarImagem.Text);

                        if (!System.IO.Directory.Exists(info.DirectoryName))
                        {
                            System.IO.Directory.CreateDirectory(info.DirectoryName);
                        }
                    }

                    System.IO.File.Copy("slide" + index.ToString().PadLeft(3, '0') + ".png", textBoxLocalSalvarImagem.Text, true);
                }
            }
        }

        private void FormPrincipal_FormClosed(object sender, FormClosedEventArgs e)
        {
            for (int i = 0; i < listBoxSlides.Items.Count; i++)
            {
                if (i == 0) pictureBoxSlide.Image = Apresentacao.Properties.Resources.FundoSlide;
                System.IO.File.Delete("slide" + (i + 1).ToString().PadLeft(3, '0') + ".png");
            }
        }

        private void checkBoxGerarImagem_CheckedChanged(object sender, EventArgs e)
        {
            checkBoxEnviarCopia.Enabled = checkBoxGerarImagem.Checked;

            checkBoxEnviarCopia_CheckedChanged(null, null);
        }

        private void checkBoxEnviarCopia_CheckedChanged(object sender, EventArgs e)
        {
            textBoxLocalSalvarImagem.Enabled = (checkBoxEnviarCopia.Enabled && checkBoxEnviarCopia.Checked);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            webBrowserVMix.Url = new Uri(textBoxVMixURL.Text);
        }

        private void powerPointToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ConfigPowerPoint config = new ConfigPowerPoint();
            config.ShowDialog();
        }

        private void salvarPorwerPointToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listBoxSelecionado.Items.Count == 0)
            {
                MessageBox.Show("Selecione pelo menos um hino.", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                saveFileDialog1.DefaultExt = ".pptx";
                saveFileDialog1.Filter = "Apresentação do Powerpoint 97-2003|*.ppt|Apresentação do Powerpoint 2007|*.pptx";
                saveFileDialog1.FilterIndex = 2;
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    ShowPresentation(true);
                    GC.Collect();
                }
            }
        }

        private void importarPowerPointToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                System.IO.FileInfo arquivo = new System.IO.FileInfo(openFileDialog1.FileName);
                String nome = arquivo.Name;
                String caminho = arquivo.FullName;
                listBoxSelecionado.Items.Add(new Item(TipoItem.Arquivo, new ArquivoItem(nome, nome)));
            }
        }

        private void sobreOProgramaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SobrePrograma s = new SobrePrograma();
            s.ShowDialog();
        }

        private void sToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (toolStripMenuItemAtivarSocket.Checked)
            {
                toolStripStatusLabelServer.Text = "Integração com vMix ATIVADO";
                toolStripTextBoxIPSocket.Enabled = true;
                buttonAvancarLinha.Enabled = true;
                buttonVoltarLinha.Enabled = true;
                SelecionarLinhaAndEnviar(false);

                tableLayoutPanelAtivo.SetRowSpan(textBoxAtivo, 1);
                tableLayoutNavegacaoLinha.Visible = true;

                if (!navegadorToolStripMenuItem.Checked)
                {
                    navegadorToolStripMenuItem.Checked = true;
                    visaoToolStripMenuItem_CheckStateChanged(null, null);
                }
            }
            else
            {
                toolStripStatusLabelServer.Text = "Integração com vMix DESATIVADO";
                toolStripTextBoxIPSocket.Enabled = false;
                buttonAvancarLinha.Enabled = false;
                buttonVoltarLinha.Enabled = false;
                labelLinhaAtiva.Text = "";

                tableLayoutPanelAtivo.SetRowSpan(textBoxAtivo, 2);
                tableLayoutNavegacaoLinha.Visible = false;

                if (navegadorToolStripMenuItem.Checked)
                {
                    navegadorToolStripMenuItem.Checked = false;
                    visaoToolStripMenuItem_CheckStateChanged(null, null);
                }
            }
        }

        private void limparTudoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            vMixSetupXAML(TIPO_VAZIO);
        }

        private void textBoxBibliaEnterKey_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                buttonBibliaAdicionar_Click(null, null);
            }
            if (e.Control && (e.KeyCode == Keys.D1 || e.KeyCode == Keys.D2 || e.KeyCode == Keys.D3 || e.KeyCode == Keys.D4 || e.KeyCode == Keys.D5))
            {
                SalvarVersiculoHistorico( e.KeyValue - 49 );
            }
        }

        private void SalvarVersiculoHistorico(int idBotao)
        {
            if (validaCamposBiblia())
            {
                string traducao = ((Item)comboBoxBibliaTraducao.SelectedItem).GetItemChaveValor().chave;
                string livro = ((Item)comboBoxBibliaLivro.SelectedItem).GetItemChaveValor().chave;
                int capitulo = Convert.ToInt32(textBoxBibliaCapitulo.Text);
                int versiculo = Convert.ToInt32(textBoxBibliaVersiculo.Text);

                int adicional = 0;
                if (textBoxBibliaVersiculoAdicional.Text != "")
                {
                    adicional = Convert.ToInt32(textBoxBibliaVersiculoAdicional.Text);
                }

                historicoBiblia[idBotao] = new BibliaItem(traducao, livro, capitulo, versiculo, adicional);

                Color cor = Color.FromArgb(170, 255, 170);
                switch (idBotao)
                {
                    case 0:
                        buttonHistoricoBiblia1.BackColor = cor;
                        break;
                    case 1:
                        buttonHistoricoBiblia2.BackColor = cor;
                        break;
                    case 2:
                        buttonHistoricoBiblia3.BackColor = cor;
                        break;
                    case 3:
                        buttonHistoricoBiblia4.BackColor = cor;
                        break;
                    case 4:
                        buttonHistoricoBiblia5.BackColor = cor;
                        break;
                    default:
                        break;
                }
            }
        }

        private void buttonHistoricoBiblia_MouseDown(object sender, MouseEventArgs e)
        {
            string nome = ((Button)sender).Name;
            int historicoIndex = Convert.ToInt32(nome.Substring(nome.Length - 1, 1)) - 1;
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                ((Button)sender).BackColor = Color.FromArgb(224, 224, 224);
                historicoBiblia[historicoIndex] = null;
            }
            else if (historicoBiblia[historicoIndex] != null)
            {
                BibliaItem item = historicoBiblia[historicoIndex];
                listBoxSelecionado.Items.Add(new Item(TipoItem.Biblia, item));

                bibliaIncluirAdicional(item);
            }
        }

        private void buttonHistoricoBiblia_MouseHover(object sender, EventArgs e)
        {
            BibliaItem item;
            string nome = ((Button)sender).Name;
            int historicoIndex = Convert.ToInt32(nome.Substring(nome.Length - 1, 1)) - 1;

            if (historicoBiblia != null && (item = historicoBiblia[historicoIndex]) != null)
            {
                toolTipBiblia.SetToolTip((Button)sender, item.ReferenciaCompleta);
                toolTipBiblia.Active = true;
            }
            else
            {
                toolTipBiblia.Active = false;
            }
        }

        private void buttonMostrarAviso_Click(object sender, EventArgs e)
        {
            if (comboBoxAvisoTitulo.Text.Trim().Length == 0)
            {
                MessageBox.Show("Digite um título para o aviso antes de mostrá-lo.", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                backgroundWorker1.DoWork += DoWork_MostrarAviso;
                backgroundWorker1.RunWorkerAsync(new string[] { comboBoxAvisoTitulo.Text.Trim(), textBoxAvisos.Text });
            }
        }

        private void buttonBibliaMostrar_Click(object sender, EventArgs e)
        {
            if (validaCamposBiblia())
            {
                string traducao = ((Item)comboBoxBibliaTraducao.SelectedItem).GetItemChaveValor().chave;
                string livro = ((Item)comboBoxBibliaLivro.SelectedItem).GetItemChaveValor().chave;
                int capitulo = Convert.ToInt32(textBoxBibliaCapitulo.Text);
                int versiculo = Convert.ToInt32(textBoxBibliaVersiculo.Text);

                int adicional = 0;
                if (textBoxBibliaVersiculoAdicional.Text != "")
                {
                    adicional = Convert.ToInt32(textBoxBibliaVersiculoAdicional.Text);
                }

                backgroundWorker1.DoWork += DoWork_MostrarBiblia;
                backgroundWorker1.RunWorkerAsync(new object[] { traducao, livro, capitulo, versiculo, adicional });
            }
        }

        private void buttonMostrarApresentacao_Click(object sender, EventArgs e)
        {
            backgroundWorker1.DoWork += DoWork_MostrarSelecionados;
            backgroundWorker1.RunWorkerAsync();
        }

        private void MostrarAviso(System.ComponentModel.DoWorkEventArgs e)
        {
            string[] args = (string[])e.Argument;

            string avisoTitulo = args[0];
            string avisoTexto = args[1];

            try
            {
                int count = 1;
                PowerPoint.Application objApp = null;
                PowerPoint.Slides objSlides = IniciarApresentacao(ref objApp);

                if (objSlides != null)
                {
                    Item item = new Item(TipoItem.Aviso, new AvisoItem(avisoTitulo, avisoTexto));
                    PreparaSlide(ref count, item, ref objSlides);

                    FinalizarApresentacao(objApp, objSlides, false);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                GC.Collect();
            }
        }

        private void MostrarBiblia(System.ComponentModel.DoWorkEventArgs e)
        {
            try
            {
                object[] args = (object[])e.Argument;

                string traducao = (string)args[0];
                string livro = (string)args[1];
                int capitulo = (int)args[2];
                int versiculo = (int)args[3];
                int adicional = (int)args[4];

                int count = 1;
                PowerPoint.Application objApp = null;
                PowerPoint.Slides objSlides = IniciarApresentacao(ref objApp);

                if (objSlides != null)
                {
                    BibliaItem biblia = new BibliaItem(traducao, livro, capitulo, versiculo, adicional);
                    PreparaSlide(ref count, new Item(TipoItem.Biblia, biblia), ref objSlides);

                    if (biblia.Incluir > 0)
                    {
                        try
                        {
                            for (int i = 0; i < biblia.Incluir; i++)
                            {
                                BibliaItem novoVersiculo = biblia.ProximoVersiculo();
                                if (novoVersiculo != null)
                                    PreparaSlide(ref count, new Item(TipoItem.Biblia, novoVersiculo), ref objSlides);
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }

                    }

                    FinalizarApresentacao(objApp, objSlides, false);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                GC.Collect();
            }
        }

        private void backgroundWorker_DoWork_MostrarAviso(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            backgroundWorker1.DoWork -= DoWork_MostrarAviso;
            MostrarAviso(e);
        }

        private void backgroundWorker_DoWork_MostrarBiblia(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            backgroundWorker1.DoWork -= DoWork_MostrarBiblia;
            MostrarBiblia(e);
        }

        private void backgroundWorker_DoWork_MostrarSelecionados(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            backgroundWorker1.DoWork -= DoWork_MostrarSelecionados;
            ShowPresentation(false);
        }

        private TemplateItem getToolStripComboBoxTemplateSelectedItem()
        {
            if (this.menuStrip1.InvokeRequired)
            {
                getToolStripComboBoxTemplateSelectedItemCallback d = getToolStripComboBoxTemplateSelectedItem;
                return (TemplateItem)this.menuStrip1.Invoke(d);
            }
            else
            {
                if (toolStripComboBoxTemplate.SelectedIndex == -1)
                    return null;

                return (TemplateItem)toolStripComboBoxTemplate.SelectedItem;

            }
        }

        private void setObjectText(object obj, string tipo, string text)
        {
            if (this.InvokeRequired)
            {
                setObjectTextCallback d = setObjectText;
                this.Invoke(d, new object[] { obj, tipo, text });
            }
            else
            {
                switch(tipo)
                {
                    case "toolStrip":
                        ((ToolStripLabel)obj).Text = text;
                        break;

                    case "textBox":
                        ((TextBox)obj).Text = text;
                        break;

                    case "label":
                        ((Label)obj).Text = text;
                        break;

                    case "groupBox":
                        ((GroupBox)obj).Text = text;
                        break;

                    default:
                        break;
                }
            }
        }

        private void setListBoxSlidesClear()
        {
            if (this.listBoxSlides.InvokeRequired)
            {
                this.listBoxSlides.Invoke(new InvokeDelegate(setListBoxSlidesClear));
            }
            else
            {
                this.listBoxSlides.Items.Clear();
            }
        }

        private void setListBoxSlidesAddItem(string item)
        {
            if (this.listBoxSlides.InvokeRequired)
            {
                setListBoxSlidesAddItemCallback d = setListBoxSlidesAddItem;
                this.listBoxSlides.Invoke(d, new object[]{ item });
            }
            else
            {
                listBoxSlides.Items.Add(item);
            }
        }

        private void setListBoxSlidesSelectFirst()
        {
            if (this.listBoxSlides.InvokeRequired)
            {
                this.listBoxSlides.Invoke(new InvokeDelegate(setListBoxSlidesSelectFirst));
            }
            else
            {
                // Se não tiver nada selecionado, então seleciona o primeiro
                if (listBoxSlides.SelectedIndex == -1)
                {
                    listBoxSlides.SelectedIndex = 0;
                    listBoxSlides_MouseClick(null, null);
                }
            }
        }

        private void setListBoxDisponivelClear()
        {
            if (this.listBoxDisponivel.InvokeRequired)
            {
                this.listBoxDisponivel.Invoke(new InvokeDelegate(setListBoxDisponivelClear));
            }
            else
            {
                this.listBoxDisponivel.Items.Clear();
            }
        }

        private void setListBoxSelecionadoClear()
        {
            if (this.listBoxSelecionado.InvokeRequired)
            {
                this.listBoxSelecionado.Invoke(new InvokeDelegate(setListBoxSelecionadoClear));
            }
            else
            {
                this.listBoxSelecionado.Items.Clear();
            }
        }

        private void setListBoxDisponivelAddItem(HinoItem item)
        {
            if (this.listBoxDisponivel.InvokeRequired)
            {
                setListBoxDisponivelAddItemCallback d = setListBoxDisponivelAddItem;
                this.listBoxDisponivel.Invoke(d, new object[] { item });
            }
            else
            {
                listBoxDisponivel.Items.Add(item);
            }
        }

        private int getListBoxSlidesCount()
        {
            if (this.listBoxSlides.InvokeRequired)
            {
                getListBoxSlidesCountCallback d = getListBoxSlidesCount;
                return (int)this.listBoxSlides.Invoke(d);
            }
            else
            {
                return this.listBoxSlides.Items.Count;
            }
        }

        private void setToolStripProgressBarMostrarStart(int max)
        {
            if (this.statusStrip1.InvokeRequired)
            {
                setToolStripProgressBarMostrarStartCallback d = setToolStripProgressBarMostrarStart;
                this.statusStrip1.Invoke(d, new object[] { max });
            }
            else
            {
                toolStripProgressBarMostrar.Visible = true;
                toolStripProgressBarMostrar.Value = 0;
                toolStripProgressBarMostrar.Minimum = 0;
                toolStripProgressBarMostrar.Maximum = max;
            }
        }

        private void setToolStripProgressBarMostrarStop()
        {
            if (this.statusStrip1.InvokeRequired)
            {
                this.statusStrip1.Invoke(new InvokeDelegate(setToolStripProgressBarMostrarStop));
            }
            else
            {
                toolStripProgressBarMostrar.Visible = false;
            }
        }

        private void setToolStripProgressBarMostrarIncrement()
        {
            if (this.statusStrip1.InvokeRequired)
            {
                this.statusStrip1.Invoke(new InvokeDelegate(setToolStripProgressBarMostrarIncrement));
            }
            else
            {
                toolStripProgressBarMostrar.Increment(1);
            }
        }

        private void visaoToolStripMenuItem_CheckStateChanged(object sender, EventArgs e)
        {
            int rowspan;

            // esquera
            rowspan = (imagemToolStripMenuItem.Checked ? 2 : 4 );
            tableLayoutPanelPrincipal.SetRowSpan(groupBoxHinos, rowspan);
            groupBoxSlides.Visible = imagemToolStripMenuItem.Checked;

            // centro
            groupBoxAvisos.Visible = avisosToolStripMenuItem.Checked;
            if (avisosToolStripMenuItem.Checked)
            {
                tableLayoutPanelPrincipal.Controls.Remove(groupBoxSelecionadas);
                tableLayoutPanelPrincipal.Controls.Add(groupBoxSelecionadas, 1, 3);
                tableLayoutPanelPrincipal.SetRowSpan(groupBoxSelecionadas, 2);
            }
            else
            {
                tableLayoutPanelPrincipal.Controls.Remove(groupBoxSelecionadas);
                tableLayoutPanelPrincipal.Controls.Add(groupBoxSelecionadas, 1, 2);
                tableLayoutPanelPrincipal.SetRowSpan(groupBoxSelecionadas, 3);
            }
            
            // direita
            rowspan = (navegadorToolStripMenuItem.Checked ? 2 : 4);
            tableLayoutPanelPrincipal.SetRowSpan(groupBoxAtiva, rowspan);
            groupBoxVMix.Visible = navegadorToolStripMenuItem.Checked;
        }

        private void buttonAtualizarBaseHinos_Click(object sender, EventArgs e)
        {
            Thread atualizar = new Thread(() => AtualizaBaseHinos());
            atualizar.IsBackground = true;
            atualizar.Start();
        }

        private bool validaCamposBiblia()
        {
            comboBoxBibliaLivro.BackColor    = System.Drawing.SystemColors.Window;
            textBoxBibliaCapitulo.BackColor  = System.Drawing.SystemColors.Window;
            textBoxBibliaVersiculo.BackColor = System.Drawing.SystemColors.Window;

            Color destaque = Color.FromArgb(255, 170, 170);

            if (comboBoxBibliaLivro.SelectedIndex == -1)
            {
                comboBoxBibliaLivro.BackColor = destaque;
                comboBoxBibliaLivro.Focus();
                return false;
            }
            
            if (textBoxBibliaCapitulo.Text.Trim().Length == 0)
            {
                textBoxBibliaCapitulo.BackColor = destaque;
                textBoxBibliaCapitulo.Focus();
                return false;
            }

            if (textBoxBibliaVersiculo.Text.Trim().Length == 0)
            {
                textBoxBibliaVersiculo.BackColor = destaque;
                textBoxBibliaVersiculo.Focus();
                return false;
            }

            return true;
        }

        private void incluirItem_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) 
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
            
        }

        private void incluirItem_DragDrop(object sender, DragEventArgs e)
        {
            string[] infoArquivo = (string[])e.Data.GetData(DataFormats.FileDrop);

            for (int i = 0; i < infoArquivo.Length; i++)
            {
                System.IO.FileInfo info = new System.IO.FileInfo(infoArquivo[i]);

                if (info.Extension != ".txt")
                {
                    MessageBox.Show("Só é aceito arquivo do tipo TXT.\nArquivo " + info.Name + " não copiado.","Atenção");
                    continue;
                }

                System.IO.DirectoryInfo hinosFolder = new System.IO.DirectoryInfo(strCurrDir + "\\Hinos");
                System.IO.FileInfo[] infoExiste = hinosFolder.GetFiles(info.Name);

                if (infoExiste.Length > 0)
                {
                    MessageBox.Show("Já existe um arquivo com esse nome\nArquivo " + info.Name + " não copiado.", "Atenção");
                    continue;
                }

                string destName = hinosFolder.FullName + "\\" + info.Name;
                System.IO.File.Copy(info.FullName, destName);

                // Analisa codificação do arquivo
                // Se não for UTF8 então faz a conversão
                Encoding originalEnconding = GetEncoding(info.FullName);
                if (originalEnconding.HeaderName != "utf-8")
                {
                    byte[] ansiBytes = System.IO.File.ReadAllBytes(destName);
                    var utf8String = Encoding.Default.GetString(ansiBytes);
                    System.IO.File.WriteAllText(destName, utf8String);
                }

                // cria novo hino nas estruturas
                HinoItem newHino = new HinoItem(info.Name, destName);
                listBoxDisponivel.Items.Add(newHino);
                baseHinos.Add(newHino);

                // inclusive nos itens selecionados
                if (((ListBox)sender).Name == "listBoxSelecionado")
                listBoxSelecionado.Items.Add(new Item(TipoItem.Hino, newHino));

            }

        }

        /// <summary>
        /// Determines a text file's encoding by analyzing its byte order mark (BOM).
        /// Defaults to ASCII when detection of the text file's endianness fails.
        /// </summary>
        /// <param name="filename">The text file to analyze.</param>
        /// <returns>The detected encoding.</returns>
        public Encoding GetEncoding(string filename)
        {
            // Read the BOM
            var bom = new byte[4];
            using (var file = new System.IO.FileStream(filename, System.IO.FileMode.Open)) file.Read(bom, 0, 4);

            // Analyze the BOM
            if (bom[0] == 0x2b && bom[1] == 0x2f && bom[2] == 0x76) return Encoding.UTF7;
            if (bom[0] == 0xef && bom[1] == 0xbb && bom[2] == 0xbf) return Encoding.UTF8;
            if (bom[0] == 0xff && bom[1] == 0xfe) return Encoding.Unicode; //UTF-16LE
            if (bom[0] == 0xfe && bom[1] == 0xff) return Encoding.BigEndianUnicode; //UTF-16BE
            if (bom[0] == 0 && bom[1] == 0 && bom[2] == 0xfe && bom[3] == 0xff) return Encoding.UTF32;
            return Encoding.ASCII;
        }

        private void buttonVoltarLinha_Click(object sender, EventArgs e)
        {
            int linhaAtiva = textBoxAtivo.GetLineFromCharIndex(textBoxAtivo.SelectionStart);

            if( linhaAtiva >= 0 )
            {
                if (linhaAtiva > 0) linhaAtiva--;

                selecionarLinhaAndAtualizaLabel(linhaAtiva);

                SelecionarLinhaAndEnviar(false);
            }

        }

        private void buttonAvancarLinha_Click(object sender, EventArgs e)
        {
            int linhaAtiva = textBoxAtivo.GetLineFromCharIndex(textBoxAtivo.SelectionStart);

            if (linhaAtiva <= textBoxAtivo.Lines.Length)
            {
                if (linhaAtiva < textBoxAtivo.Lines.Length) linhaAtiva++;

                selecionarLinhaAndAtualizaLabel(linhaAtiva);

                SelecionarLinhaAndEnviar(false);
            }

        }

        private void selecionarLinhaAndAtualizaLabel(int linhaAtiva, bool selecionarLinha = true)
        {
            if (selecionarLinha)
            {
                textBoxAtivo.SelectionStart = 0;
                textBoxAtivo.SelectionLength = 0;
                for (int i = 0; i < textBoxAtivo.Lines.Length && i <= linhaAtiva; ++i)
                {
                    if (i == linhaAtiva)
                    {
                        textBoxAtivo.SelectionLength = textBoxAtivo.Lines[i].Length;
                    }
                    else
                    {
                        textBoxAtivo.SelectionStart += textBoxAtivo.Lines[i].Length + 2;
                        //if (textBoxAtivo.Lines[i].Length == 0) textBoxAtivo.SelectionStart += 1;
                    }
                }
                textBoxAtivo.Select();
                textBoxAtivo.ScrollToCaret();
            }

            if (toolStripMenuItemAtivarSocket.Checked)
            {
                int idx = textBoxAtivo.GetLineFromCharIndex(textBoxAtivo.SelectionStart);
                if (idx < textBoxAtivo.Lines.Length)
                {
                    string linha = textBoxAtivo.Lines[idx].Trim();
                    labelLinhaAtiva.Text = linha;
                }
            }
        }
        
    }
}
