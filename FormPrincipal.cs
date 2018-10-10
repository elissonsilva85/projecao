using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using Finisar.SQLite;

namespace Apresentacao
{
    public partial class FormPrincipal : Form
    {
        String strCurrDir;
        TipoItem tipoItemSelecionado;

        public FormPrincipal()
        {
            InitializeComponent();
            filtroTituloTextBox.Focus();
        }

        private void FormPrincipal_Load(object sender, EventArgs e)
        {
            strCurrDir = System.IO.Directory.GetCurrentDirectory();
            tipoItemSelecionado = TipoItem.Nenhum;

            // Popular com os templates

            AtualizarListaTemplates();

            // Popular com os avisos

            AtualizarListaAvisos();

            // Popular os campos da Biblia

            AtualizarListaBibliaTraducao();

            AtualizarListaBibliaLivros();

            // Popular com os artistas

            AtualizarListaArtista();

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

        /*
        private void listBoxDisponivel_MouseMove(object sender, MouseEventArgs e)
        {
            int idx = listBoxDisponivel.IndexFromPoint(e.Location);
            //int idx = Convert.ToInt32(Math.Truncate(Convert.ToDouble(e.Y / listBoxDisponivel.ItemHeight)));
            if (idx >= 0 && idx < listBoxDisponivel.Items.Count)
            {
                HinoItem hino = (HinoItem)listBoxDisponivel.Items[idx];
                if (checkBoxCarregaHinoAuto.Checked)
                {
                    textBoxAtivo.Text = hino.letra;
                }
                else
                {
                    toolTip1.SetToolTip(this.listBoxDisponivel, hino.letra);
                    toolTip1.Active = true;
                }
            }
            else
            {
                if (checkBoxCarregaHinoAuto.Checked)
                {
                    textBoxAtivo.Text = "";
                }
                else
                {
                    toolTip1.Active = false;
                }
            }
        }
        */

        /*
        private void listBoxDisponivel_MouseLeave(object sender, EventArgs e)
        {
            toolTip1.Active = true;
        }
        */

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

        private void buttonSobre_Click(object sender, EventArgs e)
        {
            SobrePrograma s = new SobrePrograma();
            s.ShowDialog();
        }

        private void buttonIncluirHino_Click(object sender, EventArgs e)
        {
            SelecionarHinos();
        }

        private void buttonMostrarApresentacao_Click(object sender, EventArgs e)
        {
            ShowPresentation(false);
            GC.Collect();
        }

        private void buttonSalvarPowerpoint_Click(object sender, EventArgs e)
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
                    listBoxSelecionado.Items.Clear();

                    tipoItemSelecionado = TipoItem.Nenhum;
                    textBoxAtivo.Text = "";
                    buttonSalvarHino.Enabled = false;
                }
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
                int count = 1;
                PowerPoint.Application objApp = null;
                PowerPoint.Slides objSlides = IniciarApresentacao(ref objApp);

                if (objSlides != null)
                {
                    Item item = new Item(TipoItem.Aviso, new AvisoItem(comboBoxAvisoTitulo.Text, textBoxAvisos.Text));
                    PreparaSlide(ref count, item, ref objSlides);

                    FinalizarApresentacao(objApp, objSlides, false);
                }
            }
        }

        private void buttonIncluirPowerpoint_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                System.IO.FileInfo arquivo = new System.IO.FileInfo(openFileDialog1.FileName);
                String nome = arquivo.Name;
                String caminho = arquivo.FullName;
                listBoxSelecionado.Items.Add(new Item(TipoItem.Arquivo, new ArquivoItem(nome, nome)));
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

                    HinoItem hino = ((Item)listBoxSelecionado.SelectedItem).GetItemHino();
                    hino.AtualizarHino(textBoxTitulo.Text, textBoxAtivo.Text);
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

                        //AtualizarListaHinos();

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
            try
            {
                string traducao = ((Item)comboBoxBibliaTraducao.SelectedItem).GetItemChaveValor().chave;
                string livro = ((Item)comboBoxBibliaLivro.SelectedItem).GetItemChaveValor().chave;
                int capitulo = Convert.ToInt32(textBoxBibliaCapitulo.Text);
                int versiculo = Convert.ToInt32(textBoxBibliaVersiculo.Text);

                BibliaItem biblia = new BibliaItem(traducao, livro, capitulo, versiculo);
                listBoxSelecionado.Items.Add(new Item(TipoItem.Biblia, biblia));

                if (textBoxBibliaVersiculoAdicional.Text != "")
                {
                    try
                    {
                        int adicional = Convert.ToInt32(textBoxBibliaVersiculoAdicional.Text);
                        for(int i=0; i<adicional;i++)
                        {
                            BibliaItem novoVersiculo = biblia.ProximoVersiculo();
                            if (novoVersiculo != null)
                                listBoxSelecionado.Items.Add(new Item(TipoItem.Biblia, novoVersiculo));
                        }
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        /*
        private void textBoxBuscar_TextChanged(object sender, EventArgs e)
        {
            // se os dois campos de busca estiverem vazio, então busca tudo
            if (textBoxBuscar.Text.Length == 0 && textBoxBuscarLetra.Text.Length == 0)
            {
                AtualizarListaHinos();
            }
            // Se o campo de busca por titulo estiver vazio, busca só por letra
            else if (textBoxBuscar.Text.Length == 0)
            {
                AtualizarListaHinos(null, textBoxBuscarLetra.Text);
            }
            // Se o campo de busca por letra estiver vazio, busca só por titulo
            else if (textBoxBuscarLetra.Text.Length == 0)
            {
                AtualizarListaHinos("*" + textBoxBuscar.Text + "*.txt");
            }
            // Se os dois campos estiverem preenchidos, busca pelos dois
            else
            {
                AtualizarListaHinos("*" + textBoxBuscar.Text + "*.txt", textBoxBuscarLetra.Text);
            }
        }
        */

        #endregion

        #region FUNÇÕES

        private void AtualizarListaTemplates()
        {
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
                    comboBoxTemplate.Items.Add(new TemplateItem(templates[j].Name, templates[j].FullName));
                }
            }
            if (countTemplates > 0)
                comboBoxTemplate.SelectedIndex = 0;
        }

        private void AtualizarListaAvisos()
        {
            System.IO.DirectoryInfo avisoFolder = new System.IO.DirectoryInfo(strCurrDir + "\\Avisos");
            System.IO.FileInfo[] avisos = avisoFolder.GetFiles("*.txt");
            for (int i = 0; i < avisos.Length; i++)
                comboBoxAvisoTitulo.Items.Add(new AvisoItem(avisos[i].Name, null, avisos[i].FullName));

            comboBoxAvisoTitulo.SelectedIndex = -1;
        }

        private void AtualizarListaArtista()
        {
            DataTable artistas = HinoItem.RetornaListaArtistas();

            filtroArtistaComboBox.DataSource = artistas;
            filtroArtistaComboBox.DisplayMember = "artista";
            filtroArtistaComboBox.ValueMember = "codigo";
            filtroArtistaComboBox.SelectedIndex = -1;

            var source = new AutoCompleteStringCollection();
            for (int i = 0; i < artistas.Rows.Count; i++) source.Add(artistas.Rows[i]["artista"].ToString());
            textBoxArtista.AutoCompleteCustomSource = source;
            textBoxArtista.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            textBoxArtista.AutoCompleteSource =AutoCompleteSource.CustomSource;
            
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
            DataGridViewSelectedRowCollection linhaSelecionadas = filtroDataGridView.SelectedRows;

            for (int i = 0; i < linhaSelecionadas.Count; i++)
            {
                listBoxSelecionado.Items.Add(new Item(TipoItem.Hino, new HinoItem((int)linhaSelecionadas[i].Cells[0].Value)));
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

                        textBoxArtista.Text = hino.Artista;
                        textBoxTitulo.Text = hino.Titulo;
                        textBoxAtivo.Text = hino.Letra;

                        buttonSalvarHino.Text = "Salvar Hino";
                        buttonSalvarHino.Enabled = true;
                        
                        break;

                    case TipoItem.Aviso:
                        AvisoItem aviso = item.GetItemAviso();
                        textBoxAtivo.Text = aviso.texto;

                        buttonSalvarHino.Text = "Salvar Aviso";
                        buttonSalvarHino.Enabled = true;
                        
                        break;

                    case TipoItem.Biblia:

                        BibliaItem biblia = item.GetItemBiblia();

                        textBoxAtivo.Text = biblia.Referencia + "\r\n\r\n" + biblia.Versiculo;

                        break;

                    case TipoItem.Arquivo:
                        textBoxAtivo.Text = "O item selecionado é do tipo Arquivo.\r\nNão há como visualizar aqui.";
                        buttonSalvarHino.Enabled = false;
                        
                        break;

                    default:
                        textBoxAtivo.Text = "Item selecionado inválido.";
                        buttonSalvarHino.Enabled = false;
                        
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

            if (objSlides != null)
            {
                // Google it
                // PowerPoint.Application.SlideShowWindows

                // Loop nos hinos

                for (int idx = 0; idx < listBoxSelecionado.Items.Count; idx++)
                {
                    item = (Item)listBoxSelecionado.Items[idx];
                    PreparaSlide(ref count, item, ref objSlides);
                }

                FinalizarApresentacao(objApp, objSlides, salvar);
            }

        }

        private PowerPoint.Slides IniciarApresentacao(ref PowerPoint.Application objApp)
        {
            if (comboBoxTemplate.SelectedIndex == -1)
            {
                MessageBox.Show("Selecione um template para exibir.");
                return null;
            }

            String strTemplate = ((TemplateItem)comboBoxTemplate.SelectedItem).caminho;

            PowerPoint.Presentations objPresSet;
            PowerPoint._Presentation objPres;
            PowerPoint.Slides objSlides;
            //PowerPoint._Slide objSlide;
            //PowerPoint.Shapes objShapes;
            //PowerPoint.Shape objShape;

            //Criar uma nova apresentação baseada no template
            objApp = new PowerPoint.Application();
            objApp.Visible = MsoTriState.msoTrue;
            objApp.WindowState = Microsoft.Office.Interop.PowerPoint.PpWindowState.ppWindowMinimized;
            //objApp.Assistant.On = false;
            objPresSet = objApp.Presentations;
            objPres = objPresSet.Open(strTemplate, MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoTrue);
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

                    if (count > 1)
                        objSlide = objSlides.Add(count++, PowerPoint.PpSlideLayout.ppLayoutBlank);

                    HinoItem hino = item.GetItemHino();
                    String blocoHino = "";
                    String linhaHino;

                    objSlide = objSlides.Add(count++, PowerPoint.PpSlideLayout.ppLayoutTitle);
                    objTextRng = objSlide.Shapes.Title.TextFrame.TextRange;
                    objTextRng.Text = hino.Titulo;

                    while ((linhaHino = hino.ReadLine()) != null)
                    {
                        if (linhaHino.Length == 0)
                        {
                            objSlide = objSlides.Add(count++, PowerPoint.PpSlideLayout.ppLayoutText);
                            objTextRng = objSlide.Shapes[2].TextFrame.TextRange;
                            objTextRng.Text = blocoHino;
                            objTextRng.Font.Size = Convert.ToInt32(textBoxTamanhoFonte.Text);
                            objTextRng.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletNone;

                            blocoHino = "";
                        }
                        else
                        {
                            blocoHino += linhaHino + "\r\n";
                        }
                    }

                    if (blocoHino.Length > 0)
                    {
                        objSlide = objSlides.Add(count++, PowerPoint.PpSlideLayout.ppLayoutText);
                        objTextRng = objSlide.Shapes[2].TextFrame.TextRange;
                        objTextRng.Text = blocoHino;
                        objTextRng.Font.Size = Convert.ToInt32(textBoxTamanhoFonte.Text);
                        objTextRng.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletNone;
                    }

                    break;

                case TipoItem.Aviso:

                    if (count > 1)
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
                            objTextRng.Font.Size = Convert.ToInt32(textBoxTamanhoFonte.Text);
                            objTextRng.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletNone;

                            blocoAviso = "";
                        }
                        else
                        {
                            blocoAviso += linhaAviso + "\r\n";
                        }
                    }

                    if (blocoAviso.Length > 0)
                    {
                        objSlide = objSlides.Add(count++, PowerPoint.PpSlideLayout.ppLayoutContentWithCaption);
                        objTextRng = objSlide.Shapes[1].TextFrame.TextRange;
                        objTextRng.Text = aviso.nome;

                        objTextRng = objSlide.Shapes[3].TextFrame.TextRange;
                        objTextRng.Text = blocoAviso;
                        objTextRng.Font.Size = Convert.ToInt32(textBoxTamanhoFonte.Text);
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

            PowerPoint.SlideShowWindows objSSWs;
            PowerPoint.SlideShowTransition objSST;
            PowerPoint.SlideShowSettings objSSS;
            PowerPoint.SlideRange objSldRng;
            PowerPoint._Presentation objPres = (PowerPoint._Presentation)objSlides.Parent;
            bool bAssistantOn;

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

            int[] SlideIdx = new int[100];
            for (int i = 1; i < objSlides.Count; i++) SlideIdx[i] = i + 1;
            objSldRng = objSlides.Range(SlideIdx);
            objSST = objSldRng.SlideShowTransition;
            objSST.AdvanceOnTime = MsoTriState.msoFalse;
            objSST.EntryEffect = PowerPoint.PpEntryEffect.ppEffectFadeSmoothly;
            objSST.Speed = PowerPoint.PpTransitionSpeed.ppTransitionSpeedFast;

            // Previne que o Assistente do Officce apareca
            bAssistantOn = false; // objApp.Assistant.On;
            //objApp.Assistant.On = false;

            // Roda os Slides
            objSSS = objPres.SlideShowSettings;
            objSSS.StartingSlide = 1;
            objSSS.EndingSlide = objSlides.Count;
            objSSS.LoopUntilStopped = MsoTriState.msoTrue;
            objSSS.RangeType = PowerPoint.PpSlideShowRangeType.ppShowSlideRange;

            if (salvar)
            {
                objPres.SaveAs(saveFileDialog1.FileName, PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
            }
            else
            {
                objSSS.Run();

                // Espera até que o slideshow termine
                objSSWs = objApp.SlideShowWindows;
                try
                {
                    while (objSSWs.Count >= 1) System.Threading.Thread.Sleep(100);

                    // Reativa o Assistente do Office se ele estava ativo
                    if (bAssistantOn)
                    {
                        objApp.Assistant.On = true;
                        objApp.Assistant.Visible = false;
                    }

                    // Fecha a apresentação sem salvar
                    objPres.Close();
                    objApp.Quit();
                }
                catch
                {
                }
            }
        }

        #endregion

        private void textBoxSelecionar_Enter(object sender, EventArgs e)
        {
            ((TextBox)sender).SelectAll();
        }

        private void buttonBibliaMostrar_Click(object sender, EventArgs e)
        {
            try
            {
                int count = 1;
                PowerPoint.Application objApp = null;
                PowerPoint.Slides objSlides = IniciarApresentacao(ref objApp);

                if (objSlides != null)
                {                
                    string traducao = ((Item)comboBoxBibliaTraducao.SelectedItem).GetItemChaveValor().chave;
                    string livro = ((Item)comboBoxBibliaLivro.SelectedItem).GetItemChaveValor().chave;
                    int capitulo = Convert.ToInt32(textBoxBibliaCapitulo.Text);
                    int versiculo = Convert.ToInt32(textBoxBibliaVersiculo.Text);

                    BibliaItem biblia = new BibliaItem(traducao, livro, capitulo, versiculo);
                    PreparaSlide(ref count, new Item(TipoItem.Biblia, biblia), ref objSlides);

                    if (textBoxBibliaVersiculoAdicional.Text != "")
                    {
                        try
                        {
                            int adicional = Convert.ToInt32(textBoxBibliaVersiculoAdicional.Text);
                            for (int i = 0; i < adicional; i++)
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
        }

        private void filtroPesqHistoricoCheckBox_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void filtroConsultarButton_Click(object sender, EventArgs e)
        {
            string artista = null;
            string titulo = null;
            string letra = null;
            string inicio = null;
            string fim = null;
            bool historico = false;
        
            /*
            listBoxDisponivel.Items.Clear();

            Boolean localizou;
            HinoItem hinoTemp;
            System.IO.DirectoryInfo hinosFolder = new System.IO.DirectoryInfo(strCurrDir + "\\Hinos");
            System.IO.FileInfo[] hinos = hinosFolder.GetFiles((titulo == null) ? "*.txt" : titulo);
            for (int i = 0; i < hinos.Length; i++)
            {
                hinoTemp = new HinoItem(hinos[i].Name, hinos[i].FullName);

                if (letra == null) localizou = true;
                else localizou = (hinoTemp.letra.ToUpper().IndexOf(letra.ToUpper()) != -1);

                if (localizou) listBoxDisponivel.Items.Add(hinoTemp);
            }
            */ 
        
        }

        private void filtroLimparButton_Click(object sender, EventArgs e)
        {

        }

        private void filtroDataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void filtroDataGridView_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

        }


    }
}
