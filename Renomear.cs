using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace Apresentacao
{
    public partial class Renomear : Form
    {
        String caminho_;
        String nomeAntigo_;
        String nomeNovo_;

        public Renomear(String caminho)
        {
            InitializeComponent();

            caminho_ = System.IO.Path.GetDirectoryName(caminho) + System.IO.Path.DirectorySeparatorChar;
            nomeAntigo_ = System.IO.Path.GetFileNameWithoutExtension(caminho);

            textBoxNome.Text = nomeAntigo_;
        }

        private void Renomear_Load(object sender, EventArgs e)
        {
            textBoxNome.Focus();
        }

        private void buttonAlterar_Click(object sender, EventArgs e)
        {
            try
            {
                nomeNovo_ = textBoxNome.Text;
                System.IO.File.Move(caminho_ + nomeAntigo_ + ".txt", caminho_ + nomeNovo_ + ".txt");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public String NovoNome
        {
            get
            {
                return nomeNovo_;
            }
        }

        public String NovoCaminho
        {
            get
            {
                return caminho_ + nomeNovo_ + ".txt";
            }
        }

    }
}
