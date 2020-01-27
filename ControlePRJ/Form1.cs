using System;
using System.Collections.Generic;
//using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using MySql.Data.MySqlClient;
using System.Reflection;
using System.IO;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;


namespace ControlePRJ
{
    public partial class FormPrincipal : Form
    {
        //BANCO DE DADOS
        private MySqlConnection bdConn = new MySqlConnection(" Persist Security Info=False;server=192.168.10.6;database=controle_projeto;uid=admin;server = 192.168.10.6; database = controle_projeto; uid = admin; pwd = accenture; Allow Zero Datetime=True");
        private MySqlDataAdapter bdAdapter;
        private DataSet bdDataSet;        

        #region //****************************************** FORMs ******************************************\\

        public FormPrincipal()
        {
            InitializeComponent();
        }

        //LOAD
        private void Form1_Load(object sender, EventArgs e)
        {
            //FORM PRINCIPAL - RECUPERANDO A VERSÃO
            Version versao = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = "Controle de Projeto - Versão " + versao.ToString().Substring(0, 5);

            #region CONFIGURAÇÃO INICIAL



            /*gb_Consultar.Size = new Size(1300, 670);            
            gb_CadastroPrograma.Size = new Size(1300, 670);
            gb_CadastroRQF.Size = new Size(1300, 670);
            gb_CadastroAnalista.Size = new Size(1300, 670);
            gb_CadastroFASE.Size = new Size(1300, 670);
            gb_AcompPGM.Size = new Size(1300, 670);
            gb_Baseline.Size = new Size(1300, 670);
            gb_PesquisaBaseline.Size = new Size(1300, 670);
            gb_GrupoAnalistas.Size = new Size(1300, 670);
            gb_PesquisarProjeto.Size = new Size(1300, 670);
            gb_AcompPRJ_Contrucao.Size = new Size(1300, 670);

            gb_Consultar.Location = new Point(12, 128);            
            gb_CadastroPrograma.Location = new Point(12, 128);
            gb_CadastroRQF.Location = new Point(12, 128);
            gb_CadastroAnalista.Location = new Point(12, 128);
            gb_CadastroFASE.Location = new Point(12, 128);
            gb_AcompPGM.Location = new Point(12, 128);
            gb_Baseline.Location = new Point(12, 128);
            gb_PesquisaBaseline.Location = new Point(12, 128);
            gb_GrupoAnalistas.Location = new Point(12, 128);
            gb_PesquisarProjeto.Location = new Point(12, 128);
            gb_AcompPRJ_Contrucao.Location = new Point(12, 128);

            gb_Consultar.SendToBack();            
            gb_CadastroPrograma.SendToBack();
            gb_CadastroRQF.SendToBack();
            gb_CadastroAnalista.SendToBack();
            gb_CadastroFASE.SendToBack();
            gb_AcompPGM.SendToBack();
            gb_Baseline.SendToBack();
            gb_PesquisaBaseline.SendToBack();
            gb_GrupoAnalistas.SendToBack();
            gb_PesquisarProjeto.BringToFront();
            gb_AcompPRJ_Contrucao.SendToBack();

            /*gb_Consultar.Visible = false;            
            gb_CadastroPrograma.Visible = false;
            gb_CadastroRQF.Visible = false;
            gb_CadastroAnalista.Visible = false;
            gb_CadastroFASE.Visible = false;
            gb_AcompPGM.Visible = false;
            gb_Baseline.Visible = false;
            gb_PesquisaBaseline.Visible = false;
            gb_GrupoAnalistas.Visible = false;
            gb_PesquisarProjeto.Visible = true;
            gb_AcompPRJ_Contrucao.Visible = false;*/

            panel_AP_Programa.Size = new Size(555, 465);
            panel_AP_Construcao.Size = new Size(555, 465);
            panel_AP_Review.Size = new Size(555, 465);

            panel_AP_Programa.Location = new Point(6, 16);
            panel_AP_Construcao.Location = new Point(6, 16);
            panel_AP_Review.Location = new Point(6, 16);

            panel_AP_Programa.BringToFront();
            panel_AP_Construcao.SendToBack();
            panel_AP_Review.SendToBack();

            panel_AP_Programa.Visible = true;
            panel_AP_Construcao.Visible = false;
            panel_AP_Review.Visible = false;
            #endregion

            try
            {

                bdConn.Open();

                #region //POVOANDO COMBOBOX DE ANALISTAS
                MySqlCommand commandA = new MySqlCommand("select nome from analistas order by nome;", bdConn);
                MySqlDataReader drA = commandA.ExecuteReader();
                while (drA.Read())
                {
                    cb_respCTTU.Items.Add(drA["nome"].ToString());
                    cb_LTecnico_CPrj.Items.Add(drA["nome"].ToString());
                    cb_analistas_CR.Items.Add(drA["nome"].ToString());
                    cb_analistas_PR.Items.Add(drA["nome"].ToString());
                    cb_LGrupo.Items.Add(drA["nome"].ToString());
                    cb_AnalistaEdit.Items.Add(drA["nome"].ToString());
                    cb_LT_PesqProjeto.Items.Add(drA["nome"].ToString());
                }
                drA.Close();
                #endregion

                #region //POVOANDO COMBOBOX DE PROJETOS
                MySqlCommand commandP = new MySqlCommand("select cod_prj from prj_objeto order by cod_prj;", bdConn);
                MySqlDataReader drP = commandP.ExecuteReader();
                cb_Filtro_Projeto.Items.Add("");
                while (drP.Read())
                {
                    cb_Filtro_Projeto.Items.Add(drP["cod_prj"].ToString());
                    cb_Projeto_PBaseline.Items.Add(drP["cod_prj"].ToString());
                    cb_Projetos_PesqProjeto.Items.Add(drP["cod_prj"].ToString());
                }
                drP.Close();
                #endregion

                #region //POVOANDO COMBOBOX DE GRUPO DE ANALISTAS
                MySqlCommand commandG = new MySqlCommand("SELECT nom_grupo FROM grupo_analistas;", bdConn);
                MySqlDataReader drG = commandG.ExecuteReader();
                while (drG.Read())
                {
                    cb_GrupoAnalista.Items.Add(drG["nom_grupo"].ToString());
                    cb_NomeGrupoEdit.Items.Add(drG["nom_grupo"].ToString());
                }
                drG.Close();
                #endregion

                bdConn.Close();

            }
            catch (Exception ex)
            {
                this.Opacity = 0.9;
                ErrorForm erro = new ErrorForm(ex);
                erro.ShowDialog();
                bdConn.Close();
                this.Close();
            }

            //FIXAR TAMANHO DO FORM
            //this.FormBorderStyle = FormBorderStyle.FixedDialog;
        }

        //CHAMAR BOTÃO CONCLUÍDO
        public void botaoConcluido(string msg)
        {
            ConcluidoForm cf = new ConcluidoForm(msg);
            cf.ShowDialog();
        }

        //CHAMAR BOTÃO ALERT
        public void botaoAlert(string msg)
        {
            AlertForm af = new AlertForm(msg);
            af.ShowDialog();
        }

        //EXIBE LOAD
        void exibeLOAD()
        {
            //PROGRESS BAR                        
            progressBar1.Visible = true;
            progressBar1.Maximum = 100;
            progressBar1.Visible = true;
            for (int i = 0; i <= 100; i++)
            {
                progressBar1.Value = i;
                Thread.Sleep(TimeSpan.FromMilliseconds(50));
            }

        }

        //FECHA LOAD
        void fechaLOAD()
        {
            progressBar1.Visible = false;
            progressBar1.Value = 0;
        }

        #endregion

        #region //*********************************************** MENUS ***********************************************\\

        //MENU -  CADASTRAR NOVO PROJETO
        private void cadastrarToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            inicioCPRJ();

            abreCadastroPROJETO();
        }

        //MENU - PESQUISAR PROJETO
        private void pesquisarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            abrePesquisaPROJETO();
        }

        //MENU - ACOMPANHAMENTO DE PROJETO - STATUS DE CONSTRUÇÃO
        private void statusDeConstruçãoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            inicioAcompPRJ();

            abreAcompPRJ_Contrucao();
        }

        //MENU - PESQUISAR PROGRAMA
        private void pesquisarToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            bt_LimparConsulta.PerformClick();
            abreCONSULTA();
        }

        //MENU - CADASTRAR NOVO PROGRAMA       
        private void novoProgramaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            cadastroProgramaInicio();

            abreCadastroPROGRAMA();
        }

        //MENU - ACOMPANHAMENTO DE PROGRAMA       
        private void acompanhamentoToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            abreAcompanhamentoPGM();
        }

        //MENU - CADASTRAR RQF
        private void criarRQFToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            cadastroRQFInicio();

            abreCadastroRQF();
        }

        //MENU - CADASTRAR ANALISTA
        private void cadastrarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //LIMPA CAMPOS
            bt_Limpar_Analista.PerformClick();

            //ALTERA CAIXA DE TEXTO NOME
            tb_NomeAnalista.Visible = true;
            cb_AnalistaEdit.Visible = false;

            gb_subAnalistas.Text = "Novo Analista";

            //ALTERA BOTÕES
            bt_Salvar_Analista.Visible = true;
            bt_Limpar_Analista.BringToFront();
            bt_AtualizarAnalistas.Visible = false;
            bt_ExcluirAnalista.Visible = false;  //BOTÃO EXCLUIR

            bdConn.Open();
            cb_GrupoAnalista.Items.Clear();
            MySqlCommand commandG = new MySqlCommand("SELECT nom_grupo FROM grupo_analistas;", bdConn);
            MySqlDataReader drG = commandG.ExecuteReader();
            while (drG.Read())
                cb_GrupoAnalista.Items.Add(drG["nom_grupo"].ToString());
            drG.Close();
            bdConn.Close();

            abreCadastroAnalista();
        }

        //MENU - EDITAR ANALISTA
        private void editarExcluirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //LIMPA CAMPOS
            bt_Limpar_Analista.PerformClick();

            //ALTERA CAIXA DE TEXTO NOME
            tb_NomeAnalista.Visible = false;
            cb_AnalistaEdit.Visible = true;

            gb_subAnalistas.Text = "Editar Analista";

            //ALTERA BOTÕES
            bt_Salvar_Analista.Visible = false;
            bt_Limpar_Analista.SendToBack();
            bt_AtualizarAnalistas.Visible = true;
            bt_ExcluirAnalista.Visible = true;    //BOTÃO EXCLUIR

            abreCadastroAnalista();
        }

        //MENU - CADASTRAR FASE
        private void cadastrarToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            abreCadastroFASE();
        }

        //MENU - CADASTRAR BASELINE
        private void cadastrarBaselineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            abreCadastroBASELINE();
        }

        //MENU - PESQUISAR BASELINE
        private void pesquisarBaselineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bt_LimpaConsBaseline.PerformClick();
            abrePesquisaBASELINE();
        }

        //MENU - CADASTRAR GRUPO ANALISTA
        private void novoGrupoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //MODIFICA TEXTBOX
            tb_NomeGrupo.Visible = true;
            cb_NomeGrupoEdit.Visible = false;

            //BOTÕES
            bt_SalvarGrupo.Visible = true;
            bt_LimparGrupo.BringToFront();
            bt_EditarGrupo.Visible = false;
            bt_ExcluirGrupo.Visible = false;

            //NOME GROUPBOX
            gb_subGrupos.Text = "Novo Grupo";

            bt_LimparGrupo.PerformClick();
            atualizaCBAnalistas();
            abreCadastroGRUPO();
        }

        //MENU - EDITAR/EXCLUIR GRUPO ANALISTA
        private void editarExcluirToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //MODIFICA TEXTBOX
            tb_NomeGrupo.Visible = false;
            cb_NomeGrupoEdit.Visible = true;

            //BOTÕES
            bt_SalvarGrupo.Visible = false;
            bt_LimparGrupo.SendToBack();
            bt_EditarGrupo.Visible = true;
            bt_ExcluirGrupo.Visible = true;

            //NOME GROUPBOX
            gb_subGrupos.Text = "Editar Grupo";

            atualzaCBNomeGrupo();
            bt_LimparGrupo.PerformClick();
            abreCadastroGRUPO();
        }

        //MENU - SAIR
        private void sairToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult resultSair = MessageBox.Show("Tem certeza que deseja fechar o programa?", "Fechar o Programa", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (resultSair == DialogResult.Yes)
                this.Close();
        }

        #region MENUS ACOMPANHAMENTO
        private void programaToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            gb_Acompanhamento_Menu.Text = "Programa";

            panel_AP_Programa.BringToFront();
            panel_AP_Construcao.SendToBack();
            panel_AP_Review.SendToBack();

            panel_AP_Programa.Visible = true;
            panel_AP_Construcao.Visible = false;
            panel_AP_Review.Visible = false;
        }

        private void construçãoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            gb_Acompanhamento_Menu.Text = "Construção";

            panel_AP_Programa.SendToBack();
            panel_AP_Construcao.BringToFront();
            panel_AP_Review.SendToBack();

            panel_AP_Programa.Visible = false;
            panel_AP_Construcao.Visible = true;
            panel_AP_Review.Visible = false;
        }

        private void reviewToolStripMenuItem_Click(object sender, EventArgs e)
        {
            gb_Acompanhamento_Menu.Text = "Review";

            panel_AP_Programa.SendToBack();
            panel_AP_Construcao.SendToBack();
            panel_AP_Review.BringToFront();

            panel_AP_Programa.Visible = false;
            panel_AP_Construcao.Visible = false;
            panel_AP_Review.Visible = true;
        }
        #endregion

        //ABRE CADATRO DE PROJETO
        private void abreCadastroPROJETO()
        {
            panel_CadastroProjeto.BringToFront();
            panel_CadastroProjeto.Dock = DockStyle.Fill;

            //ATIVA PANEL
            panel_FasesProjeto_CPrj.Dock = DockStyle.Fill;
            panel_InfoProjeto_CPrj.Dock = DockStyle.Fill;            

            //ATIVA PANEL            
            panel_InfoProjeto_CPrj.Visible = true;
            panel_FasesProjeto_CPrj.Visible = false;            
        }

        //ABRE ACOMPANHAMENTO DE PROJETO - STATUS DE CONSTRUÇÃO
        private void abreAcompPRJ_Contrucao()
        {
            gb_AcompPRJ_Contrucao.BringToFront();

            gb_Consultar.Visible = false;
            //gb_CadastroProjeto.Visible = true;
            gb_CadastroPrograma.Visible = false;
            gb_CadastroRQF.Visible = false;
            gb_CadastroAnalista.Visible = false;
            gb_CadastroFASE.Visible = false;
            gb_AcompPGM.Visible = false;
            gb_Baseline.Visible = false;
            gb_PesquisaBaseline.Visible = false;
            gb_GrupoAnalistas.Visible = false;
            gb_PesquisarProjeto.Visible = false;
            gb_AcompPRJ_Contrucao.Visible = true;
        }

        //ABRE CADATRO DE PROGRAMA
        private void abreCadastroPROGRAMA()
        {
            gb_CadastroPrograma.BringToFront();

            gb_Consultar.Visible = false;
           // gb_CadastroProjeto.Visible = false;
            gb_CadastroPrograma.Visible = true;
            gb_CadastroRQF.Visible = false;
            gb_CadastroAnalista.Visible = false;
            gb_CadastroFASE.Visible = false;
            gb_AcompPGM.Visible = false;
            gb_Baseline.Visible = false;
            gb_PesquisaBaseline.Visible = false;
            gb_GrupoAnalistas.Visible = false;
            gb_PesquisarProjeto.Visible = false;
            gb_AcompPRJ_Contrucao.Visible = false;
        }

        //ABRE CONSULTA
        private void abreCONSULTA()
        {
            gb_Consultar.BringToFront();

            gb_Consultar.Visible = true;
            //gb_CadastroProjeto.Visible = false;
            gb_CadastroPrograma.Visible = false;
            gb_CadastroRQF.Visible = false;
            gb_CadastroAnalista.Visible = false;
            gb_CadastroFASE.Visible = false;
            gb_AcompPGM.Visible = false;
            gb_Baseline.Visible = false;
            gb_PesquisaBaseline.Visible = false;
            gb_GrupoAnalistas.Visible = false;
            gb_PesquisarProjeto.Visible = false;
            gb_AcompPRJ_Contrucao.Visible = false;
        }

        //ABRE CADASTRO RQF
        private void abreCadastroRQF()
        {
            panel_RQF.BringToFront();
            panel_RQF.Dock = DockStyle.Fill;
        }

        //ABRE CADASTRO ANALISTA
        private void abreCadastroAnalista()
        {
            gb_CadastroAnalista.BringToFront();

            gb_Consultar.Visible = false;
            //gb_CadastroProjeto.Visible = false;
            gb_CadastroPrograma.Visible = false;
            gb_CadastroRQF.Visible = false;
            gb_CadastroAnalista.Visible = true;
            gb_CadastroFASE.Visible = false;
            gb_AcompPGM.Visible = false;
            gb_Baseline.Visible = false;
            gb_PesquisaBaseline.Visible = false;
            gb_GrupoAnalistas.Visible = false;
            gb_PesquisarProjeto.Visible = false;
            gb_AcompPRJ_Contrucao.Visible = false;
        }

        //ABRE CADASTRO FASE
        private void abreCadastroFASE()
        {
            gb_CadastroFASE.BringToFront();

            gb_Consultar.Visible = false;
            //gb_CadastroProjeto.Visible = false;
            gb_CadastroPrograma.Visible = false;
            gb_CadastroRQF.Visible = false;
            gb_CadastroAnalista.Visible = false;
            gb_CadastroFASE.Visible = true;
            gb_AcompPGM.Visible = false;
            gb_Baseline.Visible = false;
            gb_PesquisaBaseline.Visible = false;
            gb_GrupoAnalistas.Visible = false;
            gb_PesquisarProjeto.Visible = false;
            gb_AcompPRJ_Contrucao.Visible = false;
        }

        //ABRE ACOMPANHAMENTO DE PROGRAMA
        private void abreAcompanhamentoPGM()
        {
            gb_AcompPGM.BringToFront();

            gb_Consultar.Visible = false;
            //gb_CadastroProjeto.Visible = false;
            gb_CadastroPrograma.Visible = false;
            gb_CadastroRQF.Visible = false;
            gb_CadastroAnalista.Visible = false;
            gb_CadastroFASE.Visible = false;
            gb_AcompPGM.Visible = true;
            gb_Baseline.Visible = false;
            gb_PesquisaBaseline.Visible = false;
            gb_GrupoAnalistas.Visible = false;
            gb_PesquisarProjeto.Visible = false;
            gb_AcompPRJ_Contrucao.Visible = false;
        }

        //ABRE CADASTRO BASELINE
        private void abreCadastroBASELINE()
        {
            gb_Baseline.BringToFront();

            gb_Consultar.Visible = false;
            //gb_CadastroProjeto.Visible = false;
            gb_CadastroPrograma.Visible = false;
            gb_CadastroRQF.Visible = false;
            gb_CadastroAnalista.Visible = false;
            gb_CadastroFASE.Visible = false;
            gb_AcompPGM.Visible = false;
            gb_Baseline.Visible = true;
            gb_PesquisaBaseline.Visible = false;
            gb_GrupoAnalistas.Visible = false;
            gb_PesquisarProjeto.Visible = false;
            gb_AcompPRJ_Contrucao.Visible = false;
        }

        //ABRE PESQUISA BASELINE
        private void abrePesquisaBASELINE()
        {
            gb_PesquisaBaseline.BringToFront();

            gb_Consultar.Visible = false;
           // gb_CadastroProjeto.Visible = false;
            gb_CadastroPrograma.Visible = false;
            gb_CadastroRQF.Visible = false;
            gb_CadastroAnalista.Visible = false;
            gb_CadastroFASE.Visible = false;
            gb_AcompPGM.Visible = false;
            gb_Baseline.Visible = false;
            gb_PesquisaBaseline.Visible = true;
            gb_GrupoAnalistas.Visible = false;
            gb_PesquisarProjeto.Visible = false;
            gb_AcompPRJ_Contrucao.Visible = false;
        }

        //ABRE CADASTRO DE GRUPO ANALISTA
        private void abreCadastroGRUPO()
        {
            gb_GrupoAnalistas.BringToFront();

            gb_Consultar.Visible = false;
           // gb_CadastroProjeto.Visible = false;
            gb_CadastroPrograma.Visible = false;
            gb_CadastroRQF.Visible = false;
            gb_CadastroAnalista.Visible = false;
            gb_CadastroFASE.Visible = false;
            gb_AcompPGM.Visible = false;
            gb_Baseline.Visible = false;
            gb_PesquisaBaseline.Visible = false;
            gb_GrupoAnalistas.Visible = true;
            gb_PesquisarProjeto.Visible = false;
            gb_AcompPRJ_Contrucao.Visible = false;
        }

        //ABRE PESQUISA PROJETO
        private void abrePesquisaPROJETO()
        {
            gb_PesquisarProjeto.BringToFront();

            gb_Consultar.Visible = false;
            //gb_CadastroProjeto.Visible = false;
            gb_CadastroPrograma.Visible = false;
            gb_CadastroRQF.Visible = false;
            gb_CadastroAnalista.Visible = false;
            gb_CadastroFASE.Visible = false;
            gb_AcompPGM.Visible = false;
            gb_Baseline.Visible = false;
            gb_PesquisaBaseline.Visible = false;
            gb_GrupoAnalistas.Visible = false;
            gb_PesquisarProjeto.Visible = true;
            gb_AcompPRJ_Contrucao.Visible = false;
        }

        #endregion

        #region //***************************************** PESQUISA DE PROGRAMA *****************************************\\

        //BOTÃO PESQUISAR
        private void bt_Consultar_Click(object sender, EventArgs e)
        {
            try
            {
                //CRIANDO DATASET E POVOANDO                
                bdDataSet = new DataSet();
                bdConn.Open();
                bdAdapter = new MySqlDataAdapter(cria_queryPESQUISA(), bdConn);
                bdAdapter.Fill(bdDataSet, "pgm_objeto");                              
                dataGrid_Consulta.DataSource = bdDataSet;
                dataGrid_Consulta.DataMember = "pgm_objeto";

                if (dataGrid_Consulta.RowCount == 0)
                    semRESULTADO();
                else
                {
                    //FORMATA GRIDVIEW
                    formataGRIDVIEW();

                    panel_NotFound.Visible = false;
                    bt_ComprimirGrid.Visible = true;
                    bt_ExportarExecel.Visible = true;
                }


                //FECHA CONEXÃO
                bdConn.Close();
            }
            catch (Exception ex)
            {
                this.Opacity = 0.9;
                ErrorForm erro = new ErrorForm(ex);
                erro.ShowDialog();
                bdConn.Close();
                this.Close();
            }
        }

        //BOTÃO EXPORTAR PARA EXCEL
        private void bt_ExportarExecel_Click(object sender, EventArgs e)
        {
            //OBJETO PARA SALVAR
            SaveFileDialog salvar = new SaveFileDialog();

            //CRIA PLANILHA
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
            Excel.Workbook workBook = excelApp.Workbooks.Add(); //PASTA
            Excel.Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet; //PLANILHA

            #region CRIA CABEÇALHO
            workSheet.Cells[1, "A"] = "Programa";
            workSheet.Cells[1, "B"] = "DF/RQF";
            workSheet.Cells[1, "C"] = "Projeto";
            workSheet.Cells[1, "D"] = "Sistema";
            workSheet.Cells[1, "E"] = "Responsável CCTU";
            workSheet.Cells[1, "F"] = "Responsável DSOL";
            workSheet.Cells[1, "G"] = "Duplicidade";
            workSheet.Cells[1, "H"] = "Peso";
            workSheet.Cells[1, "I"] = "Status Contrução";
            workSheet.Cells[1, "J"] = "Inicio Construçao";
            workSheet.Cells[1, "K"] = "Fim Construçao";
            workSheet.Cells[1, "L"] = "Anotações Gerais";
            workSheet.Cells[1, "M"] = "Liberado CR";
            workSheet.Cells[1, "N"] = "Responsável CR";
            workSheet.Cells[1, "O"] = "Status CR";
            workSheet.Cells[1, "P"] = "Data CR";
            workSheet.Cells[1, "Q"] = "Liberado PR";
            workSheet.Cells[1, "R"] = "Responsável PR";
            workSheet.Cells[1, "S"] = "Status PR";
            workSheet.Cells[1, "T"] = "Data PR";
            #endregion

            //PASSA VALORES DO DATAGRIDVIEW PARA PLANILHA
            int indiceCell = 2;
            for (int i = 0; i <= dataGrid_Consulta.RowCount - 1; i++)
            {
                for (int j = 0; j <= dataGrid_Consulta.ColumnCount - 1; j++)
                {
                    DataGridViewCell cell = dataGrid_Consulta[j, i];
                    if (!cell.Value.ToString().Equals("00/00/0000"))
                        workSheet.Cells[indiceCell, (j + 1)] = cell.Value.ToString().Replace("{", "");
                }
                indiceCell++;
            }

            //APLICA FORMATAÇÃO NA PLANILHA
            workSheet.Columns.AutoFit();
            Excel.Range bodyExcel = workSheet.get_Range("A1", ("T" + (dataGrid_Consulta.RowCount + 1).ToString()));
            bodyExcel.Select();
            bodyExcel.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet.ListObjects.AddEx(Excel.XlListObjectSourceType.xlSrcRange, bodyExcel, Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing).Name = "MyTableStyle";
            workSheet.ListObjects.get_Item("MyTableStyle").TableStyle = "TableStyleMedium1";

            //CONFIGURAÇÕES PARA SALVAR O ARQUIVO
            salvar.Title = "Exportar para Excel";
            salvar.Filter = "Arquivo do Excel *.xlsx | *.xlsx";
            salvar.ShowDialog(); // mostra

            //SALVA O ARQUIVO                      
            workBook.SaveAs(salvar.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workBook.Close(true, Type.Missing, Type.Missing);
            excelApp.Quit(); // ENCERRA O EXCEL

            //FINALIZA OPERAÇÃO
            exibeLOAD();
            botaoConcluido("Planilha gerada com sucesso!");
            fechaLOAD();
        }

        //CRIA QUERY PARA PESQUISAR REGISTROS
        public string cria_queryPESQUISA()
        {
            string buscar = "" +
                "P.cod_pgm," +
                " P.df_rqf," +
                " P.cod_prj," +
                " P.sistema," +
                " P.resp_cttu," +
                " P.resp_dsol," +
                " P.duplicidade," +
                " P.peso_pgm," +
                " PA.status_construcao," +
                " sum(PAD.pct_pgm)," +
                " PA.data_inicio," +
                " PA.data_fim," +
                " PA.anot_gerais," +
                " PA.lib_cr," +
                " PA.resp_cr," +
                " PA.status_cr," +
                " PA.data_cr," +
                " PA.lib_pr," +
                " PA.resp_pr," +
                " PA.status_pr," +
                " PA.data_pr";

            string consulta = "SELECT " + buscar + " FROM pgm_objeto P INNER JOIN pgm_acompanhamento PA INNER JOIN pgm_acomp_data PAD ON" +
                " P.cod_pgm = PA.cod_pgm AND" +
                " P.df_rqf = PA.df_rqf AND" +
                " P.cod_prj = PA.cod_prj AND" +
                " PA.cod_pad = PAD.cod_pad";

            if (cb_Filtro_Projeto.Text != "")
                consulta += " and P.cod_prj like '%" + cb_Filtro_Projeto.Text + "%'";

            if (tb_Filtro_DFRQF.Text != "")
                consulta += " and P.df_rqf like '%" + tb_Filtro_DFRQF.Text + "%'";

            if (tb_Filtro_Programa.Text != "")
                consulta += " and P.cod_pgm like '%" + tb_Filtro_Programa.Text + "%'";

            if (cb_Filtro_CTTU.Text != "")
                consulta += " and P.resp_cttu like '%" + cb_Filtro_CTTU.Text + "%'";

            if (cb_Filtro_Sistema.Text != "")
                consulta += " and P.sistema like '%" + cb_Filtro_Sistema.Text + "%'";

            if (checkBox_Filtro_Iniciado.Checked == true)
                consulta += " and PA.status_construcao = 'Iniciado'";

            if (checkBox_Filtro_Finalizado.Checked == true)
                consulta += " and PA.status_construcao = 'Finalizado'";

            consulta += " GROUP BY P.cod_pgm, P.df_rqf, P.cod_prj";

            return consulta;
        }

        //PESQUISA AUTOMATICA - PROJETO
        private void cb_Filtro_Projeto_TextChanged(object sender, EventArgs e)
        {
            if (cb_Filtro_Projeto.Text != "")
                bt_Consultar.PerformClick();
        }

        //PESQUISA AUTOMATICA - PROGRAMA
        private void tb_Filtro_Programa_TextChanged(object sender, EventArgs e)
        {
            if (tb_Filtro_Programa.Text != "")
                bt_Consultar.PerformClick();
        }

        //PESQUISA AUTOMATICA - DF RQF
        private void tb_Filtro_DFRQF_TextChanged(object sender, EventArgs e)
        {
            if (tb_Filtro_DFRQF.Text != "")
                bt_Consultar.PerformClick();
        }

        //PESQUISA AUTOMATICA - SISTEMA
        private void cb_Filtro_Sistema_TextChanged(object sender, EventArgs e)
        {
            if (cb_Filtro_Sistema.Text != "")
                bt_Consultar.PerformClick();
        }

        //PESQUISA AUTOMATICA - CCTU
        private void cb_Filtro_CTTU_TextChanged(object sender, EventArgs e)
        {
            if (cb_Filtro_CTTU.Text != "")
                bt_Consultar.PerformClick();
        }

        //PESQUISA AUTOMATICA - STATUS CONSTRUÇÃO - INICIADO
        private void checkBox_Filtro_Iniciado_CheckedChanged(object sender, EventArgs e)
        {
            bt_Consultar.PerformClick();
        }

        //PESQUISA AUTOMATICA - STATUS CONSTRUÇÃO - FINALIZADO
        private void checkBox_Filtro_Finalizado_CheckedChanged(object sender, EventArgs e)
        {
            bt_Consultar.PerformClick();
        }

        //BOTÃO FILTROS 
        private void bt_Filtros_Click(object sender, EventArgs e)
        {
            //POVOANDO COMBO BOX DE ANALISTAS
            if (cb_Filtro_CTTU.Items.Count == 0)
            {
                try
                {
                    bdConn.Open();
                    string query = "select nome from analistas order by nome;";
                    MySqlCommand command = new MySqlCommand(query, bdConn);
                    MySqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                    {
                        cb_Filtro_CTTU.Items.Add(dr["nome"].ToString());
                    }
                    dr.Close();
                    bdConn.Close();
                }
                catch (Exception ex)
                {
                    this.Opacity = 0.9;
                    ErrorForm erro = new ErrorForm(ex);
                    erro.ShowDialog();
                    bdConn.Close();
                    this.Close();
                }

            }

            //ATIVA E DESATIVA FILTROS
            if (panel_Filtros.Visible == true)
                panel_Filtros.Visible = false;
            else
                panel_Filtros.Visible = true;
        }

        //BOTÃO DESATIVA FILTROS
        private void bt_Sem_Filtro_Click(object sender, EventArgs e)
        {
            panel_Filtros.Visible = false;

            //LIMPAR CAMPOS
            LimparFiltros();
        }

        //BOTÃO PESQUISAR - FILTROS
        private void bt_PesquisaFiltros_Click(object sender, EventArgs e)
        {
            //PESQUISAR
            bt_Consultar.PerformClick();
        }

        //ALTERA LAYOUT GRIDVIEW
        public void formataGRIDVIEW()
        {
            try
            {
                this.dataGrid_Consulta.Columns[0].HeaderText = "Programa";
                this.dataGrid_Consulta.Columns[1].HeaderText = "DF/RQF";
                this.dataGrid_Consulta.Columns[2].HeaderText = "Projeto";
                this.dataGrid_Consulta.Columns[3].HeaderText = "Sistema";
                this.dataGrid_Consulta.Columns[4].HeaderText = "Responsável CCTU";
                this.dataGrid_Consulta.Columns[5].HeaderText = "Responsável DSOL";
                this.dataGrid_Consulta.Columns[6].HeaderText = "Duplicidade";
                this.dataGrid_Consulta.Columns[7].HeaderText = "Peso";
                this.dataGrid_Consulta.Columns[8].HeaderText = "Status Contrução";
                this.dataGrid_Consulta.Columns[9].HeaderText = "Total Contruido";
                this.dataGrid_Consulta.Columns[10].HeaderText = "Inicio Construção";
                this.dataGrid_Consulta.Columns[11].HeaderText = "Fim Construção";
                this.dataGrid_Consulta.Columns[12].HeaderText = "Anotações Gerais";
                this.dataGrid_Consulta.Columns[13].HeaderText = "Liberado CR";
                this.dataGrid_Consulta.Columns[14].HeaderText = "Responsável CR";
                this.dataGrid_Consulta.Columns[15].HeaderText = "Status CR";
                this.dataGrid_Consulta.Columns[16].HeaderText = "Data CR";
                this.dataGrid_Consulta.Columns[17].HeaderText = "Liberado PR";
                this.dataGrid_Consulta.Columns[18].HeaderText = "Responsável PR";
                this.dataGrid_Consulta.Columns[19].HeaderText = "Status PR";
                this.dataGrid_Consulta.Columns[20].HeaderText = "Data PR";
            }
            catch (Exception ex)
            {
                ErrorForm erro = new ErrorForm(ex);
                erro.ShowDialog();
            }
        }

        //LIMPAR GRIDVIEW
        public void limpaGrid()
        {
            if (this.dataGrid_Consulta.DataSource != null)
                this.dataGrid_Consulta.DataSource = null;
            else
            {
                this.dataGrid_Consulta.Rows.Clear();
                this.dataGrid_Consulta.Columns.Clear();
            }
        }

        //BOTÃO LIMPAR CONSULTA
        private void bt_LimparConsulta_Click(object sender, EventArgs e)
        {
            //LIMPAR GRIDVIEW
            limpaGrid();

            //LIMPA NOT FOUND
            panel_NotFound.Visible = false;

            //BOTÃO EXPORTAR EXCEL
            bt_ExportarExecel.Visible = false;

            //RESETA EXPAND. E COMPR.
            bt_ComprimirGrid.Visible = false;
            bt_ExpandirGrid.Visible = false;
            dataGrid_Consulta.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;

            //LIMPAR CAMPOS
            LimparFiltros();
        }

        //RESETA FILTROS
        private void LimparFiltros()
        {
            //CAMPOS DA PESQUISA
            cb_Filtro_Projeto.Text = "";
            tb_Filtro_DFRQF.Text = "";
            tb_Filtro_Programa.Text = "";
            cb_Filtro_CTTU.Text = "";
            cb_Filtro_Sistema.Text = "";
        }

        //NÃO RETORNOU RESULTADO
        public void semRESULTADO()
        {
            //LIMPA GRIDVIEW
            this.dataGrid_Consulta.Columns.Clear();

            //DESATIVA MSG NOT FOUND
            panel_NotFound.Visible = true;

            //EXPORTAR FALSE
            bt_ExportarExecel.Visible = false;

            //RESETA EXPAND. E COMPR.
            bt_ComprimirGrid.Visible = false;
            bt_ExpandirGrid.Visible = false;
            dataGrid_Consulta.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;


        }

        //EXPANDIR GRIDVIEW
        private void bt_ExpandirGrid_Click(object sender, EventArgs e)
        {
            //REDIMENSIONA TAMANHO DAS COLUNAS
            dataGrid_Consulta.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;

            bt_ComprimirGrid.Visible = true;
            bt_ExpandirGrid.Visible = false;
        }

        //COMPRIMIR GRIDVIEW
        private void bt_Comprimir_Click(object sender, EventArgs e)
        {
            //REDIMENSIONA TAMANHO DAS COLUNAS
            dataGrid_Consulta.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            bt_ComprimirGrid.Visible = false;
            bt_ExpandirGrid.Visible = true;
        }

        //FORMATA CELULAS
        private void dataGrid_Consulta_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //DUPLICIDADE
            if (e.Value != null && e.ColumnIndex == 7)
                if (e.Value.Equals("SIM"))
                    e.CellStyle.ForeColor = Color.Red;

            //TOTAL CONSTRUIDO
            if (e.ColumnIndex == 9)
                if (String.IsNullOrEmpty(e.Value.ToString()))
                    e.Value = "0 %";
                else
                    e.Value += " %";

            //DATA INICIO CONSTRUÇÃO
            if (e.Value != null && e.ColumnIndex == 10)
                if (e.Value.ToString().Equals("00/00/0000"))
                    e.Value = "-";

            //DATA FIM CONSTRUÇÃO
            if (e.Value != null && e.ColumnIndex == 11)
                if (e.Value.ToString().Equals("00/00/0000"))
                    e.Value = "-";

            //ANOTAÇÕES GERAIS
            if (e.Value != null && e.ColumnIndex == 12)
                if (e.Value.ToString().Equals(""))
                    e.Value = "-";

            //RESPONSÁVEL CODE REVIEW
            if (e.Value != null && e.ColumnIndex == 14)
                if (e.Value.ToString().Equals(""))
                    e.Value = "-";

            //DATA CODE REVIEW
            if (e.Value != null && e.ColumnIndex == 16)
                if (e.Value.ToString().Equals("00/00/0000"))
                    e.Value = "-";

            //RESPONSÁVEL PERFORMACE REVIEW
            if (e.Value != null && e.ColumnIndex == 18)
                if (e.Value.ToString().Equals(""))
                    e.Value = "-";

            //DATA PERFORMACE REVIEW
            if (e.Value != null && e.ColumnIndex == 20)
                if (e.Value.ToString().Equals("00/00/0000"))
                    e.Value = "-";
        }

        //DUPLO CLICK - DATA GRID VIEW PESQUISA - ABRE ACOMPANHAMENTO DO PRGRAMA
        private void dataGrid_Consulta_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            DialogResult resultAbreAcompanhamento = MessageBox.Show("Deseja abrir o acompanhamento do programa " + dataGrid_Consulta.CurrentRow.Cells[0].Value.ToString() + "?", "Abrir Acompanhamento", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (resultAbreAcompanhamento == DialogResult.Yes)
            {
                if (dataGrid_Consulta.CurrentRow != null)
                {
                    abreAcompanhamentoPGM();

                    panel_selecionaPrograma.Visible = false;
                    panel_acompanhamentoPGM.Visible = true;
                    bt_Salvar_Acompanhamento.Visible = true;
                    bt_Limpar_Acompanhamento.Visible = true;

                    pgm_select = dataGrid_Consulta.CurrentRow.Cells[0].Value.ToString();
                    rqf_select = dataGrid_Consulta.CurrentRow.Cells[1].Value.ToString();
                    prj_select = dataGrid_Consulta.CurrentRow.Cells[2].Value.ToString();

                    carregaAcompanhamentoPGM();
                }
            }
        }

        #endregion

        #region //**************************************** CADASTRAR PROGRAMA ****************************************\\

        //VARIAVEIS CADASTRO PROGRAMA
        RespDSOL respdsol = new RespDSOL();
        string querySPRJ;
        bool verificaCampos_CPgm = true;
        string camposInvalidos_MSG = "O preenchimento dos campos são obrigatórios!";

        //BOTÃO SALVAR
        private void bt_Salvar_Click(object sender, EventArgs e)
        {
            try
            {
                validaCAMPOS_CPgm();

                if (verificaCampos_CPgm == true)
                {

                    bdConn.Open();

                    cadastraPROGRAMA();

                    cadastraAcompanhamentoPGM();

                    cadastraDatasAcomp();

                    bdConn.Close();
                   
                    botaoConcluido("Tudo certo. O programa foi cadastrado com sucesso.");

                    tb_Programa.Text = ""; 
                }
                else
                {
                    botaoAlert(camposInvalidos_MSG);
                    verificaCampos_CPgm = true;
                }

            }
            catch (Exception ex)
            {
                ErrorForm erro = new ErrorForm(ex);
                erro.ShowDialog();
                bdConn.Close();
            }
        }

        //BOTÃO LIMPAR PESQUISA
        private void bt_Limpar_CP_Click(object sender, EventArgs e)
        {
            cb_df_rqf.Text = "";
            tb_Programa.Text = "";
            cb_respCTTU.Text = "";
            tb_RespDSOL.Text = "";
            respdsol.Nomes = "";
            cb_Sistema.SelectedIndex = 0;
            numericUp_Peso.Value = 0;
            verificaCampos_CPgm = true;

            lbl_Prj_CP.ForeColor = Color.Black;
            lbl_dfrqf_CP.ForeColor = Color.Black;
            lbl_pgm_CP.ForeColor = Color.Black;
            lbl_cttu_CP.ForeColor = Color.Black;
            lbl_dsol_CP.ForeColor = Color.Black;
            lbl_sistema_CP.ForeColor = Color.Black;
            lbl_Peso_CP.ForeColor = Color.Black;
        }

        //CRIA QUERY PARA GRAVAR DADOS
        public void cadastraPROGRAMA()
        {
            try
            {
                //VERIFICA DUPLICIDADE
                bool verificaDuplicidade = false;
                MySqlCommand verificaDuplucidade = new MySqlCommand("select * from pgm_objeto WHERE cod_pgm = '" + tb_Programa.Text + "' and cod_prj = '" + tb_Projeto.Text + "';", bdConn);
                MySqlDataReader drDuplucidade = verificaDuplucidade.ExecuteReader();
                if (drDuplucidade.Read() == true)
                    verificaDuplicidade = true;
                drDuplucidade.Close();

                //CRIA QUERY DE CADASTRO
                string cmd_insert = "INSERT INTO pgm_objeto (cod_pgm, df_rqf, cod_prj, sistema, resp_cttu, peso_pgm, resp_dsol, duplicidade) VALUES('" +
                    tb_Programa.Text.ToString() + "','" +
                    cb_df_rqf.Text.ToString() + "','" +
                    tb_Projeto.Text.ToString() + "','" +
                    cb_Sistema.Text.ToString() + "','" +
                    cb_respCTTU.Text.ToString() + "','" +
                    numericUp_Peso.Text.ToString() + "','" +
                    respdsol.Nomes.ToString() + "','" +
                    "NAO" +
                    "');";

                //CADASTRA PROGRAMA
                MySqlCommand command = new MySqlCommand(cmd_insert, bdConn);
                command.ExecuteNonQuery();

                //SE TIVER DUPLICIDADE ALTERA STATUS DE DUPLICIDADE
                if (verificaDuplicidade)
                {
                    command = new MySqlCommand("UPDATE pgm_objeto SET duplicidade = 'SIM' WHERE cod_pgm = '" + tb_Programa.Text + "';", bdConn);
                    command.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                ErrorForm erro = new ErrorForm(ex);
                erro.ShowDialog();
                bdConn.Close();
            }
        }

        //GRAVA ACOMPANHAMENTO DE PROGRAMA
        private void cadastraAcompanhamentoPGM()
        {
            string queryAcompanhamento;

            queryAcompanhamento = "INSERT INTO pgm_acompanhamento (cod_pgm, df_rqf, cod_prj, status_construcao, data_inicio, data_fim, anot_gerais, lib_cr, resp_cr, status_cr, data_cr, lib_pr, resp_pr, status_pr, data_pr) VALUES('" +
                tb_Programa.Text + "','" +
                cb_df_rqf.Text + "','" +
                tb_Projeto.Text + "','" +
                "Não Iniciado" + "','" +
                null + "','" +
                null + "','" +
                "" + "','" +
                "Nao" + "','" +
                "" + "','" +
                "Não Iniciado" + "','" +
                null + "','" +
                "Nao" + "','" +
                "" + "','" +
                "Não Iniciado" + "','" +
                null +
                "');";

            MySqlCommand command = new MySqlCommand(queryAcompanhamento, bdConn);
            command.ExecuteNonQuery();
        }

        //CADASTRA DATAS DE ACOMPANHAMENTOS
        void cadastraDatasAcomp()
        {
            List<DateTime> datasAcompanamento = new List<DateTime>();
            string codPAD = "";

            //RECUPERA CODIGO PROGRAMA ACOMPANHAMENTO DATA
            MySqlCommand command = new MySqlCommand("SELECT cod_pad FROM pgm_acompanhamento WHERE  cod_pgm = '" + tb_Programa.Text + "' and df_rqf = '" + cb_df_rqf.Text + "' and cod_prj = '" + tb_Projeto.Text + "';", bdConn);
            MySqlDataReader dr = command.ExecuteReader();
            if (dr.Read())
                codPAD = dr["cod_pad"].ToString();
            dr.Close();

            //RECUPERA DATAS DE ACOMPANHAMENTO
            command = new MySqlCommand("Select D.data from prj_objeto P natural join prj_data_acomp natural join data_acomp D where  P.cod_prj = '" + tb_Projeto.Text + "';", bdConn);
            dr = command.ExecuteReader();
            while (dr.Read())
                datasAcompanamento.Add(DateTime.Parse(dr["data"].ToString()));
            dr.Close();

            //CADASTRA DATAS DE ACOMPANHAMENTO PARA PROGRAMAS DO PROJETO
            foreach (var data in datasAcompanamento)
            {
                command = new MySqlCommand("INSERT INTO pgm_acomp_data (cod_pad, data_acomp) VALUES (" + Int16.Parse(codPAD) + ",'" + data.ToString("yyyy/MM/dd") + "');", bdConn);
                command.ExecuteNonQuery();
            }
        }

        //VALIDAÇÃO DOS CAMPOS DE CADASTRO DE PROGRAMAS
        public void validaCAMPOS_CPgm()
        {
            //VALIDA PROJETO
            if (tb_Projeto.Text == "")
            {
                verificaCampos_CPgm = false;
                lbl_Prj_CP.ForeColor = Color.Red;
            }
            else
                lbl_Prj_CP.ForeColor = Color.Black;

            //VALIDA DF/RQF
            if (cb_df_rqf.Text == "")
            {
                verificaCampos_CPgm = false;
                lbl_dfrqf_CP.ForeColor = Color.Red;
            }
            else
                lbl_dfrqf_CP.ForeColor = Color.Black;

            //VALIDA PROGRAMA
            if (tb_Programa.Text == "")
            {
                verificaCampos_CPgm = false;
                lbl_pgm_CP.ForeColor = Color.Red;
            }
            else
                lbl_pgm_CP.ForeColor = Color.Black;

            //VALIDA RESPONSÁVEL DSOL
            if (tb_RespDSOL.Text == "")
            {
                verificaCampos_CPgm = false;
                lbl_dsol_CP.ForeColor = Color.Red;
            }
            else
                lbl_dsol_CP.ForeColor = Color.Black;

            //VALIDA SISTEMA
            if (cb_Sistema.Text == "")
            {
                verificaCampos_CPgm = false;
                lbl_sistema_CP.ForeColor = Color.Red;
            }
            else
                lbl_sistema_CP.ForeColor = Color.Black;

            //VALIDA PESO
            if (numericUp_Peso.Value == 0)
            {
                verificaCampos_CPgm = false;
                lbl_Peso_CP.ForeColor = Color.Red;
            }
            else
                lbl_Peso_CP.ForeColor = Color.Black;

            //ATUALIZA FORM
            this.Refresh();
        }

        //BOTAO PARA ADICIONAR NOMES DO RESP DSOL 
        private void respDF_Click(object sender, EventArgs e)
        {
            respdsol.Nomes = "";

            respdsol.ShowDialog();

            tb_RespDSOL.Text = respdsol.Nomes;
        }

        //SELECIONAR PRJ - PESQUISA AUTOMATICA POR PROJETO
        private void tb_selecionaPRJ_TextChanged(object sender, EventArgs e)
        {

            try
            {
                if (tb_selecionaPRJ.Text != "")
                {
                    querySPRJ = "select cod_prj from prj_objeto where cod_prj like '%" + tb_selecionaPRJ.Text + "%'";
                    bdDataSet = new DataSet();
                    bdConn.Open();
                    bdAdapter = new MySqlDataAdapter(querySPRJ, bdConn);
                    bdAdapter.Fill(bdDataSet, "prj_objeto");
                    dataGrid_selecionaPRJ.DataSource = bdDataSet;

                    if (bdDataSet.Tables["prj_objeto"].Rows.Count == 0)
                        lb_ProjetoNaoEncontrado_CPgm.Visible = true;
                    else
                        lb_ProjetoNaoEncontrado_CPgm.Visible = false;

                    dataGrid_selecionaPRJ.DataMember = "prj_objeto";
                    bdConn.Close();
                }
                else
                {
                    lb_ProjetoNaoEncontrado_CPgm.Visible = false;
                    if (this.dataGrid_selecionaPRJ.DataSource != null)
                        this.dataGrid_selecionaPRJ.DataSource = null;
                    else
                    {
                        this.dataGrid_selecionaPRJ.Rows.Clear();
                        this.dataGrid_selecionaPRJ.Columns.Clear();
                    }
                }

            }
            catch (Exception ex)
            {
                this.Opacity = 0.9;
                ErrorForm erro = new ErrorForm(ex);
                erro.ShowDialog();
                bdConn.Close();
                this.Close();
            }
        }

        //DUPLO CLICK NA CELULA CELECIONADA
        private void dataGrid_selecionaPRJ_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGrid_selecionaPRJ.CurrentRow != null)
                bt_Ok_CPrj.PerformClick();
            else
                botaoAlert("Selecionar um projeto da lista antes de prosseguir.");
        }

        //BOTÃO OK - SELECIONAR PROJETO
        private void bt_Ok_CPrj_Click(object sender, EventArgs e)
        {
            if (bdDataSet.Tables["prj_objeto"].Rows.Count > 0)
            {
                try
                {
                    if (cb_df_rqf.Items.Count == 0)
                        cb_df_rqf.Items.Add("");

                    bdConn.Open();
                    string queryRQF = "select cod_rqf from df_rqf where cod_prj = '" + dataGrid_selecionaPRJ.CurrentRow.Cells[0].Value.ToString() + "';";
                    MySqlCommand commandRQF = new MySqlCommand(queryRQF, bdConn);
                    MySqlDataReader drRQF = commandRQF.ExecuteReader();

                    while (drRQF.Read())
                        cb_df_rqf.Items.Add(drRQF["cod_rqf"].ToString());

                    drRQF.Close();
                    bdConn.Close();

                    if (cb_df_rqf.Items.Count > 1)
                    {
                        panel_CP1.Visible = false;
                        panel_CP2.Visible = true;
                        bt_Salvar_CP.Visible = true;
                        bt_Limpar_CP.Visible = true;
                        tb_Projeto.Text = dataGrid_selecionaPRJ.CurrentRow.Cells[0].Value.ToString();
                    }
                    else
                        botaoAlert("O projeto selecionado não possui DF/RQF.");
                }
                catch (Exception ex)
                {
                    ErrorForm erro = new ErrorForm(ex);
                    erro.ShowDialog();
                    bdConn.Close();
                }
            }
            else
                MessageBox.Show("Nenhum projeto foi encontrado! Pesquise novamente.", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        //CADASTRO PROGRAMA INICIO
        public void cadastroProgramaInicio()
        {
            bt_Limpar_CP.PerformClick();
            tb_Projeto.Text = "";
            tb_selecionaPRJ.Text = "";
            cb_df_rqf.Items.Clear();


            panel_CP1.Visible = true;
            panel_CP2.Visible = false;
            bt_Salvar_CP.Visible = false;
            bt_Limpar_CP.Visible = false;
        }

        //PRESS KEY - ENTER
        private void tb_selecionaPRJ_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Keys)e.KeyChar == Keys.Enter)
                bt_Ok_CPrj.PerformClick();
        }

        //BOTÃO VOLTAR TELA DE SELECIONAR O PROJETO
        private void bt_Return_CPgm_Click(object sender, EventArgs e)
        {
            DialogResult resultReturnCPgm = MessageBox.Show("Deseja voltar e escolher outro Projeto?\n\nTodas alterações serão perdidas!", "Voltar", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (resultReturnCPgm == DialogResult.Yes)
                cadastroProgramaInicio();

        }

        #endregion

        #region //**************************************** CADASTRAR PROJETO ****************************************\\

        //VARIAVEIS UTILIZADAS
        MySqlCommand commandPRJ;
        string camposInvalidos_MSG_CPrj = "O preenchimento dos campos são obrigatórios!";
        int posicaoTela = 1;

        #region BOTÕES GERAIS

        //BOTÃO SALVAR CADASTRO DE PROJETO
        private void bt_Salvar_CPrj_Click(object sender, EventArgs e)
        {
            try
            {

                bdConn.Open();

                cadastrarPROJETO();
                cadastraFASE();
                cadastroDatasAcomp();

                bdConn.Close();


                exibeLOAD();

                botaoConcluido("Tudo certo. O projeto foi cadastrado com sucesso.");

                fechaLOAD();

                abreCadastroPROJETO();
                inicioCPRJ();

            }
            catch (Exception ex)
            {
                ErrorForm erro = new ErrorForm(ex);
                erro.ShowDialog();
                bdConn.Close();
            }
        }

        //BOTÃO NEXT - CADASTRO DE PROJETO - PANEL 1
        private void bt_Next_CPrj_Click(object sender, EventArgs e)
        {
            if (validaCAMPOS_CPrj())
            {
                //ATIVA PANEL
                panel_InfoProjeto_CPrj.Visible = false;
                panel_FasesProjeto_CPrj.Visible = true;
            }
            else
                botaoAlert(camposInvalidos_MSG_CPrj);            
        }        

        //BOTÃO ANTERIOR - CADASTRO DE PROJETO
        private void bt_Prev_CPrj_Click(object sender, EventArgs e)
        {                      
                //ATIVA PANEL
                panel_FasesProjeto_CPrj.Visible = false;
                panel_InfoProjeto_CPrj.Visible = true;                

            if (posicaoTela == 2)
            {
                //ATIVA PANEL
                panel_FasesProjeto_CPrj.Visible = true;
                panel_InfoProjeto_CPrj.Visible = false;                

                //HABILITA BOTÕES
                bt_Salvar_CPrj.Visible = false;
                bt_Prev_CadastroProjeto2.Visible = true;
                bt_Next_CPrj.Visible = true;
                bt_Limpar_CPrj.Visible = false;
            }            
        }

        //BOTÃO LIMPAR
        private void bt_Limpar_CPrj_Click(object sender, EventArgs e)
        {
            limpaCPrj();
        }

        //LIMPAR FORM CPrj
        public void limpaCPrj()
        {
            tb_Projeto_CPrj.Text = "00000000";
            cb_LTecnico_CPrj.Text = null;
            tb_LRequerimento_CPrj.Text = "";
            tb_RTecnico_CPrj.Text = "";
            cb_Status_CPrj.Text = null;
            rtb_DescProjeto_CPrj.Text = "";
            
            cb_F4_CPrj.Text = null;
            cb_F5_CPrj.Text = null;
            cb_F6_CPrj.Text = null;
            cb_F7_CPrj.Text = null;
            cb_F8_CPrj.Text = null;
            cb_F9_CPrj.Text = null;
            cb_F10_CPrj.Text = null;

            cb_Status_F1.Text = null;
            cb_Status_F2.Text = null;
            cb_Status_F3.Text = null;
            cb_Status_F4.Text = null;
            cb_Status_F5.Text = null;
            cb_Status_F6.Text = null;
            cb_Status_F7.Text = null;
            cb_Status_F8.Text = null;
            cb_Status_F9.Text = null;
            cb_Status_F10.Text = null;

            lb_CP_CPrj.ForeColor = Color.Black;
            lb_LT_CPrj.ForeColor = Color.Black;
            lb_LR_CPrj.ForeColor = Color.Black;
            lb_RT_CPrj.ForeColor = Color.Black;
            lb_IP_CPrj.ForeColor = Color.Black;
            lb_FP_CPrj.ForeColor = Color.Black;
            lb_R_CPrj.ForeColor = Color.Black;
            lb_Status_CPrj.ForeColor = Color.Black;
            lb_DescP_CPrj.ForeColor = Color.Black;

            lb_F1.ForeColor = Color.Black;
            lb_F2.ForeColor = Color.Black;
            lb_F3.ForeColor = Color.Black;
            lb_F4.ForeColor = Color.Black;
            lb_F5.ForeColor = Color.Black;
            lb_F6.ForeColor = Color.Black;
            lb_F7.ForeColor = Color.Black;
            lb_F8.ForeColor = Color.Black;
            lb_F9.ForeColor = Color.Black;
            lb_F10.ForeColor = Color.Black;

            lb_Status_F1.ForeColor = Color.Black;
            lb_Status_F2.ForeColor = Color.Black;
            lb_Status_F3.ForeColor = Color.Black;
            lb_Status_F4.ForeColor = Color.Black;
            lb_Status_F5.ForeColor = Color.Black;
            lb_Status_F6.ForeColor = Color.Black;
            lb_Status_F7.ForeColor = Color.Black;
            lb_Status_F8.ForeColor = Color.Black;
            lb_Status_F9.ForeColor = Color.Black;
            lb_Status_F10.ForeColor = Color.Black;
        }

        #endregion

        #region FASES DO PROJETO

        //CADASTRO FASES DO PRJ
        void cadastraFASE()
        {
            try
            {
                string queryCadastoFase;

                //FASE 1 - CONTRUÇÃO                
                string codF1 = "";
                commandPRJ = new MySqlCommand("SELECT cod_fase FROM fase WHERE nom_fase = 'Construção'", bdConn);
                MySqlDataReader f1 = commandPRJ.ExecuteReader();
                if (f1.Read())
                    codF1 = f1["cod_fase"].ToString();
                f1.Close();

                queryCadastoFase = "INSERT INTO prj_fase VALUES('" +
                tb_Projeto_CPrj.Text + "'," +
                Int16.Parse(codF1) + ",'" +
                dt_F1_Inicio.Value.ToString("yyyy/MM/dd") + "','" +
                dt_F1_Fim.Value.ToString("yyyy/MM/dd") + "','" +
                cb_Status_F1.Text +
                "');";

                commandPRJ = new MySqlCommand(queryCadastoFase, bdConn);
                commandPRJ.ExecuteNonQuery();


                //FASE 2                
                string codF2 = "";
                commandPRJ = new MySqlCommand("SELECT cod_fase FROM fase WHERE nom_fase = 'Teste Consolidado'", bdConn);
                MySqlDataReader f2 = commandPRJ.ExecuteReader();
                if (f2.Read())
                    codF2 = f2["cod_fase"].ToString();
                f2.Close();

                queryCadastoFase = "INSERT INTO prj_fase VALUES('" +
                tb_Projeto_CPrj.Text + "'," +
                Int16.Parse(codF2) + ",'" +
                dt_F2_Inicio.Value.ToString("yyyy/MM/dd") + "','" +
                dt_F2_Fim.Value.ToString("yyyy/MM/dd") + "','" +
                cb_Status_F2.Text +
                "');";

                commandPRJ = new MySqlCommand(queryCadastoFase, bdConn);
                commandPRJ.ExecuteNonQuery();

                //FASE 3                
                string codF3 = "";
                commandPRJ = new MySqlCommand("SELECT cod_fase FROM fase WHERE nom_fase = 'ET Reversa'", bdConn);
                MySqlDataReader f3 = commandPRJ.ExecuteReader();
                if (f3.Read())
                    codF3 = f3["cod_fase"].ToString();
                f3.Close();

                queryCadastoFase = "INSERT INTO prj_fase VALUES('" +
                tb_Projeto_CPrj.Text + "'," +
                Int16.Parse(codF3) + ",'" +
                dt_F3_Inicio.Value.ToString("yyyy/MM/dd") + "','" +
                dt_F3_Fim.Value.ToString("yyyy/MM/dd") + "','" +
                cb_Status_F3.Text +
                "');";

                commandPRJ = new MySqlCommand(queryCadastoFase, bdConn);
                commandPRJ.ExecuteNonQuery();


                //FASE 4
                if (panel_FaseProjeto_4.Enabled == true)
                {
                    string codF4 = "";
                    commandPRJ = new MySqlCommand("SELECT cod_fase FROM fase WHERE nom_fase = '" + cb_F4_CPrj.Text + "'", bdConn);
                    MySqlDataReader f4 = commandPRJ.ExecuteReader();
                    if (f4.Read())
                        codF4 = f4["cod_fase"].ToString();
                    f4.Close();

                    queryCadastoFase = "INSERT INTO prj_fase VALUES('" +
                    tb_Projeto_CPrj.Text + "'," +
                    Int16.Parse(codF4) + ",'" +
                    dt_F4_Inicio.Value.ToString("yyyy/MM/dd") + "','" +
                    dt_F4_Fim.Value.ToString("yyyy/MM/dd") + "','" +
                    cb_Status_F4.Text +
                    "');";

                    commandPRJ = new MySqlCommand(queryCadastoFase, bdConn);
                    commandPRJ.ExecuteNonQuery();
                }

                //FASE 5
                if (panel_FaseProjeto_5.Enabled == true)
                {
                    string codF5 = "";
                    commandPRJ = new MySqlCommand("SELECT cod_fase FROM fase WHERE nom_fase = '" + cb_F5_CPrj.Text + "'", bdConn);
                    MySqlDataReader f5 = commandPRJ.ExecuteReader();
                    if (f5.Read())
                        codF5 = f5["cod_fase"].ToString();
                    f5.Close();

                    queryCadastoFase = "INSERT INTO prj_fase VALUES('" +
                    tb_Projeto_CPrj.Text + "'," +
                    Int16.Parse(codF5) + ",'" +
                    dt_F5_Inicio.Value.ToString("yyyy/MM/dd") + "','" +
                    dt_F5_Fim.Value.ToString("yyyy/MM/dd") + "','" +
                    cb_Status_F5.Text +
                    "');";

                    commandPRJ = new MySqlCommand(queryCadastoFase, bdConn);
                    commandPRJ.ExecuteNonQuery();
                }

                //FASE 6
                if (panel_FaseProjeto_6.Enabled == true)
                {
                    string codF6 = "";
                    commandPRJ = new MySqlCommand("SELECT cod_fase FROM fase WHERE nom_fase = '" + cb_F6_CPrj.Text + "'", bdConn);
                    MySqlDataReader f6 = commandPRJ.ExecuteReader();
                    if (f6.Read())
                        codF6 = f6["cod_fase"].ToString();
                    f6.Close();

                    queryCadastoFase = "INSERT INTO prj_fase VALUES('" +
                    tb_Projeto_CPrj.Text + "'," +
                    Int16.Parse(codF6) + ",'" +
                    dt_F6_Inicio.Value.ToString("yyyy/MM/dd") + "','" +
                    dt_F6_Fim.Value.ToString("yyyy/MM/dd") + "','" +
                    cb_Status_F6.Text +
                    "');";

                    commandPRJ = new MySqlCommand(queryCadastoFase, bdConn);
                    commandPRJ.ExecuteNonQuery();
                }

                //FASE 7
                if (panel_FaseProjeto_7.Enabled == true)
                {
                    string codF7 = "";
                    commandPRJ = new MySqlCommand("SELECT cod_fase FROM fase WHERE nom_fase = '" + cb_F7_CPrj.Text + "'", bdConn);
                    MySqlDataReader f7 = commandPRJ.ExecuteReader();
                    if (f7.Read())
                        codF7 = f7["cod_fase"].ToString();
                    f7.Close();

                    queryCadastoFase = "INSERT INTO prj_fase VALUES('" +
                    tb_Projeto_CPrj.Text + "'," +
                    Int16.Parse(codF7) + ",'" +
                    dt_F7_Inicio.Value.ToString("yyyy/MM/dd") + "','" +
                    dt_F7_Fim.Value.ToString("yyyy/MM/dd") + "','" +
                    cb_Status_F7.Text +
                    "');";

                    commandPRJ = new MySqlCommand(queryCadastoFase, bdConn);
                    commandPRJ.ExecuteNonQuery();
                }

                //FASE 8
                if (panel_FaseProjeto_8.Enabled == true)
                {
                    string codF8 = "";
                    commandPRJ = new MySqlCommand("SELECT cod_fase FROM fase WHERE nom_fase = '" + cb_F8_CPrj.Text + "'", bdConn);
                    MySqlDataReader f8 = commandPRJ.ExecuteReader();
                    if (f8.Read())
                        codF8 = f8["cod_fase"].ToString();
                    f8.Close();

                    queryCadastoFase = "INSERT INTO prj_fase VALUES('" +
                    tb_Projeto_CPrj.Text + "'," +
                    Int16.Parse(codF8) + ",'" +
                    dt_F8_Inicio.Value.ToString("yyyy/MM/dd") + "','" +
                    dt_F8_Fim.Value.ToString("yyyy/MM/dd") + "','" +
                    cb_Status_F8.Text +
                    "');";

                    commandPRJ = new MySqlCommand(queryCadastoFase, bdConn);
                    commandPRJ.ExecuteNonQuery();
                }

                //FASE 9
                if (panel_FaseProjeto_9.Enabled == true)
                {
                    string codF9 = "";
                    commandPRJ = new MySqlCommand("SELECT cod_fase FROM fase WHERE nom_fase = '" + cb_F9_CPrj.Text + "'", bdConn);
                    MySqlDataReader f9 = commandPRJ.ExecuteReader();
                    if (f9.Read())
                        codF9 = f9["cod_fase"].ToString();
                    f9.Close();

                    queryCadastoFase = "INSERT INTO prj_fase VALUES('" +
                    tb_Projeto_CPrj.Text + "'," +
                    Int16.Parse(codF9) + ",'" +
                    dt_F9_Inicio.Value.ToString("yyyy/MM/dd") + "','" +
                    dt_F9_Fim.Value.ToString("yyyy/MM/dd") + "','" +
                    cb_Status_F9.Text +
                    "');";

                    commandPRJ = new MySqlCommand(queryCadastoFase, bdConn);
                    commandPRJ.ExecuteNonQuery();
                }

                //FASE 10
                if (panel_FaseProjeto_10.Enabled == true)
                {
                    string codF10 = "";
                    commandPRJ = new MySqlCommand("SELECT cod_fase FROM fase WHERE nom_fase = '" + cb_F10_CPrj.Text + "'", bdConn);
                    MySqlDataReader f10 = commandPRJ.ExecuteReader();
                    if (f10.Read())
                        codF10 = f10["cod_fase"].ToString();
                    f10.Close();


                    queryCadastoFase = "INSERT INTO prj_fase VALUES('" +
                    tb_Projeto_CPrj.Text + "'," +
                    Int16.Parse(codF10) + ",'" +
                    dt_F10_Inicio.Value.ToString("yyyy/MM/dd") + "','" +
                    dt_F10_Fim.Value.ToString("yyyy/MM/dd") + "','" +
                    cb_Status_F10.Text +
                    "');";

                    commandPRJ = new MySqlCommand(queryCadastoFase, bdConn);
                    commandPRJ.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                this.Opacity = 0.9;
                ErrorForm erro = new ErrorForm(ex);
                erro.ShowDialog();
                bdConn.Close();
                this.Close();
            }
        }
            
        #region CONTROLE DE DATAS - FASES DO PROJETO

        private void dt_F1_Inicio_ValueChanged(object sender, EventArgs e)
        {
            dt_F1_Fim.MinDate = dt_F1_Inicio.Value;
        }

        private void dt_F2_Inicio_ValueChanged(object sender, EventArgs e)
        {
            dt_F2_Fim.MinDate = dt_F2_Inicio.Value;
        }

        private void dt_F3_Inicio_ValueChanged(object sender, EventArgs e)
        {
            dt_F3_Fim.MinDate = dt_F3_Inicio.Value;
        }

        private void dt_F4_Inicio_ValueChanged(object sender, EventArgs e)
        {
            dt_F4_Fim.MinDate = dt_F4_Inicio.Value;
        }

        private void dt_F5_Inicio_ValueChanged(object sender, EventArgs e)
        {
            dt_F5_Fim.MinDate = dt_F5_Inicio.Value;
        }

        private void dt_F6_Inicio_ValueChanged(object sender, EventArgs e)
        {
            dt_F6_Fim.MinDate = dt_F6_Inicio.Value;
        }

        private void dt_F7_Inicio_ValueChanged(object sender, EventArgs e)
        {
            dt_F7_Fim.MinDate = dt_F7_Inicio.Value;
        }

        private void dt_F8_Inicio_ValueChanged(object sender, EventArgs e)
        {
            dt_F8_Fim.MinDate = dt_F8_Inicio.Value;
        }

        private void dt_F9_Inicio_ValueChanged(object sender, EventArgs e)
        {
            dt_F9_Fim.MinDate = dt_F9_Inicio.Value;
        }

        private void dt_F10_Inicio_ValueChanged(object sender, EventArgs e)
        {
            dt_F10_Fim.MinDate = dt_F10_Inicio.Value;
        }

        #endregion

        #endregion
        
        //CADASTO PROJETO
        void cadastrarPROJETO()
        {            
            string queryCadastoProjeto;           

            queryCadastoProjeto = "INSERT INTO prj_objeto (cod_prj, desc_prj, lider_tecnico, lider_requer, resp_tecnico, status_prj, release_prj, dt_ini_prj, dt_fim_prj) VALUES ('" +
                tb_Projeto_CPrj.Text + "','" +
                rtb_DescProjeto_CPrj.Text + "','" +
                cb_LTecnico_CPrj.Text + "','" +
                tb_LRequerimento_CPrj.Text + "','" +
                tb_RTecnico_CPrj.Text + "','" +
                cb_Status_CPrj.Text + "','" +
                dt_Release_CPrj.Value.ToString("yyyy/MM/dd") + "','" +
                dt_InicioP_CPrj.Value.ToString("yyyy/MM/dd") + "','" +
                dt_FimProjeto_CPrj.Value.ToString("yyyy/MM/dd") +
                "');";

            commandPRJ = new MySqlCommand(queryCadastoProjeto, bdConn);
            commandPRJ.ExecuteNonQuery();
        }

        //VALIDAÇÃO DOS CAMPOS
        bool validaCAMPOS_CPrj()
        {
            bool verificaCampos_CPrj = true;

            //DATA DE INICIO E FIM DO PROJETO
            if (dt_InicioP_CPrj.Value > dt_FimProjeto_CPrj.Value)
            {                
                lb_IP_CPrj.ForeColor = Color.Red;
                lb_FP_CPrj.ForeColor = Color.Red;
                camposInvalidos_MSG_CPrj = "Data Inicio e/ou Data Termino é inválido.";
                return false;
            }
            else
            {
                camposInvalidos_MSG_CPrj = "O preenchimento dos campos são obrigatórios!";
                lb_IP_CPrj.ForeColor = Color.Black;
                lb_FP_CPrj.ForeColor = Color.Black;
            }

            //CODIGO DO PROJETO
            if (tb_Projeto_CPrj.Text == "")
            {
                verificaCampos_CPrj = false;
                lb_CP_CPrj.ForeColor = Color.Red;
            }
            else
                lb_CP_CPrj.ForeColor = Color.Black;

            //LIDER TECNICO
            if (cb_LTecnico_CPrj.Text == "")
            {
                verificaCampos_CPrj = false;
                lb_LT_CPrj.ForeColor = Color.Red;
            }
            else
                lb_LT_CPrj.ForeColor = Color.Black;

            //LIDER REQUERIMENTO
            /*if (tb_LRequerimento_CPrj.Text == "")
            {
                verificaCampos_CPrj = false;
                lb_LR_CPrj.ForeColor = Color.Red;
            }
            else
                lb_LR_CPrj.ForeColor = Color.Black;*/

            //RESPONSAVEL TECNICO
            /*if (tb_RTecnico_CPrj.Text == "")
            {
                verificaCampos_CPrj = false;
                lb_RT_CPrj.ForeColor = Color.Red;
            }
            else
                lb_RT_CPrj.ForeColor = Color.Black;*/

            //STATUS
            if (cb_Status_CPrj.Text == "")
            {
                verificaCampos_CPrj = false;
                lb_Status_CPrj.ForeColor = Color.Red;
            }
            else
                lb_Status_CPrj.ForeColor = Color.Black;


            //DESCRIÇÃO DO PROJETO
            /*if (rtb_DescProjeto_CPrj.Text == "")
            {
                verificaCampos_CPrj = false;
                lb_DescP_CPrj.ForeColor = Color.Red;
            }
            else
                lb_DescP_CPrj.ForeColor = Color.Black;*/           

            return verificaCampos_CPrj;
        }

        //VALIDAÇÃO DOS CAMPOS
        bool validaCAMPOS_FasesPRJ()
        {
            bool verificaCampos_FasesPRJ = true;

            //FASE CONSTRUÇÃO
            if (cb_Status_F1.Text == "")
            {
                lb_IP_CPrj.ForeColor = Color.Red;
                lb_FP_CPrj.ForeColor = Color.Red;
                camposInvalidos_MSG_CPrj = "Data Inicio e/ou Data Termino é inválido.";
                return false;
            }
            else
            {
                camposInvalidos_MSG_CPrj = "O preenchimento dos campos são obrigatórios!";
                lb_IP_CPrj.ForeColor = Color.Black;
                lb_FP_CPrj.ForeColor = Color.Black;
            }

            return verificaCampos_FasesPRJ;
        }

        //CONTROLE CAMPO DATA FIM
        private void dt_InicioP_CPrj_ValueChanged(object sender, EventArgs e)
        {
            //DATA FIM DO PROJETO
            dt_FimProjeto_CPrj.MinDate = dt_InicioP_CPrj.Value + TimeSpan.FromDays(1);

            //DATAS FASES (INICIO)
            dt_F1_Inicio.MinDate = dt_InicioP_CPrj.Value + TimeSpan.FromDays(1);
            dt_F2_Inicio.MinDate = dt_InicioP_CPrj.Value + TimeSpan.FromDays(1);
            dt_F3_Inicio.MinDate = dt_InicioP_CPrj.Value + TimeSpan.FromDays(1);
            dt_F4_Inicio.MinDate = dt_InicioP_CPrj.Value + TimeSpan.FromDays(1);
            dt_F5_Inicio.MinDate = dt_InicioP_CPrj.Value + TimeSpan.FromDays(1);
            dt_F6_Inicio.MinDate = dt_InicioP_CPrj.Value + TimeSpan.FromDays(1);
            dt_F7_Inicio.MinDate = dt_InicioP_CPrj.Value + TimeSpan.FromDays(1);
            dt_F8_Inicio.MinDate = dt_InicioP_CPrj.Value + TimeSpan.FromDays(1);
            dt_F9_Inicio.MinDate = dt_InicioP_CPrj.Value + TimeSpan.FromDays(1);
            dt_F10_Inicio.MinDate = dt_InicioP_CPrj.Value + TimeSpan.FromDays(1);

            //DATAS FASES (FINAL)
            dt_F1_Fim.MinDate = dt_InicioP_CPrj.Value + TimeSpan.FromDays(1);
            dt_F2_Fim.MinDate = dt_InicioP_CPrj.Value + TimeSpan.FromDays(1);
            dt_F3_Fim.MinDate = dt_InicioP_CPrj.Value + TimeSpan.FromDays(1);
            dt_F4_Fim.MinDate = dt_InicioP_CPrj.Value + TimeSpan.FromDays(1);
            dt_F5_Fim.MinDate = dt_InicioP_CPrj.Value + TimeSpan.FromDays(1);
            dt_F6_Fim.MinDate = dt_InicioP_CPrj.Value + TimeSpan.FromDays(1);
            dt_F7_Fim.MinDate = dt_InicioP_CPrj.Value + TimeSpan.FromDays(1);
            dt_F8_Fim.MinDate = dt_InicioP_CPrj.Value + TimeSpan.FromDays(1);
            dt_F9_Fim.MinDate = dt_InicioP_CPrj.Value + TimeSpan.FromDays(1);
            dt_F10_Fim.MinDate = dt_InicioP_CPrj.Value + TimeSpan.FromDays(1);

        }

        //CONTROLE CAMPO DATA RELEASE
        private void dt_FimProjeto_CPrj_ValueChanged(object sender, EventArgs e)
        {
            //DATA RELEASE
            dt_Release_CPrj.MinDate = dt_FimProjeto_CPrj.Value;

            //DATAS FASES (INICIO)
            dt_F1_Inicio.MaxDate = dt_FimProjeto_CPrj.Value;
            dt_F2_Inicio.MaxDate = dt_FimProjeto_CPrj.Value;
            dt_F3_Inicio.MaxDate = dt_FimProjeto_CPrj.Value;
            dt_F4_Inicio.MaxDate = dt_FimProjeto_CPrj.Value;
            dt_F5_Inicio.MaxDate = dt_FimProjeto_CPrj.Value;
            dt_F6_Inicio.MaxDate = dt_FimProjeto_CPrj.Value;
            dt_F7_Inicio.MaxDate = dt_FimProjeto_CPrj.Value;
            dt_F8_Inicio.MaxDate = dt_FimProjeto_CPrj.Value;
            dt_F9_Inicio.MaxDate = dt_FimProjeto_CPrj.Value;
            dt_F10_Inicio.MaxDate = dt_FimProjeto_CPrj.Value;

            //DATAS FASES (FIM)
            dt_F1_Fim.MaxDate = dt_FimProjeto_CPrj.Value;
            dt_F2_Fim.MaxDate = dt_FimProjeto_CPrj.Value;
            dt_F3_Fim.MaxDate = dt_FimProjeto_CPrj.Value;
            dt_F4_Fim.MaxDate = dt_FimProjeto_CPrj.Value;
            dt_F5_Fim.MaxDate = dt_FimProjeto_CPrj.Value;
            dt_F6_Fim.MaxDate = dt_FimProjeto_CPrj.Value;
            dt_F7_Fim.MaxDate = dt_FimProjeto_CPrj.Value;
            dt_F8_Fim.MaxDate = dt_FimProjeto_CPrj.Value;
            dt_F9_Fim.MaxDate = dt_FimProjeto_CPrj.Value;
            dt_F10_Fim.MaxDate = dt_FimProjeto_CPrj.Value;
        }        
        
        //CADASTRO DATAS DE ACOMPANHAMENTO
        void cadastroDatasAcomp()
        {            
            //CADASTRA PROJETO NA TABELA DE ACOMPANHAMENTO
            commandPRJ = new MySqlCommand("INSERT INTO prj_data_acomp (cod_prj) VALUES ('" + tb_Projeto_CPrj.Text + "');", bdConn);
            commandPRJ.ExecuteNonQuery();

            //RECUPERA CODIGO DO PRJETO NA TABELA DE ACOMPANHAMENTO
            int codDATA = 0;
            commandPRJ = new MySqlCommand("SELECT cod_dat FROM prj_data_acomp WHERE cod_prj = '" + tb_Projeto_CPrj.Text + "';", bdConn);
            MySqlDataReader recCod = commandPRJ.ExecuteReader();
            if (recCod.Read())
                codDATA = Int16.Parse(recCod["cod_dat"].ToString());
            recCod.Close();

            //PERIDIOCIDADE PARA ACOMPANHAMENTO
            int peridiocidadeAcomp = 0;
            switch (domainUpDown_AcompanhamentoProjeto.Text)
            {                
                case "Intervalos de 1 dia":
                    peridiocidadeAcomp = 2;
                    break;
                case "Intervalos de 2 dias":
                    peridiocidadeAcomp = 3;
                    break;
                case "Intervalos de 3 dias":
                    peridiocidadeAcomp = 4;
                    break;
                case "Intervalos de 4 dias":
                    peridiocidadeAcomp = 5;
                    break;
                case "Intervalos de 5 dias":
                    peridiocidadeAcomp = 6;
                    break;
                default:
                    peridiocidadeAcomp = 1;
                    break;
            }

            //RETORNA LISTA DE FERIADOS
            FeriadosBH.Feriados feriadosBH = new FeriadosBH.Feriados();
            Dictionary<DateTime, string> feriados = new Dictionary<DateTime, string>(feriadosBH.getListaFeriados());

            List<DateTime> datas_Acompanhamento = new List<DateTime>();
            for (DateTime i = DateTime.Parse(dt_F1_Inicio.Text); i <= DateTime.Parse(dt_F1_Fim.Text); i = i.AddDays(peridiocidadeAcomp))
            {
                if (i.DayOfWeek != DayOfWeek.Saturday && i.DayOfWeek != DayOfWeek.Sunday && (!feriados.ContainsKey(i)))
                {
                    //CADASTRA DATA PARA ACOMOPANHAMENTO
                    commandPRJ = new MySqlCommand("INSERT INTO data_acomp (cod_dat, data) VALUES (" + codDATA + ", '" + i.ToString("yyyy/MM/dd") + "');", bdConn);
                    commandPRJ.ExecuteNonQuery();
                }
                else
                {
                    //PEGAR PROXIMO DIA UTIL
                    for (DateTime x = i; x <= dt_F1_Fim.Value; x = x.AddDays(1))
                    {
                        if (x.DayOfWeek != DayOfWeek.Saturday && x.DayOfWeek != DayOfWeek.Sunday && (!feriados.ContainsKey(x)))
                        {
                            commandPRJ = new MySqlCommand("INSERT INTO data_acomp (cod_dat, data) VALUES (" + codDATA + ", '" + x.ToString("yyyy/MM/dd") + "');", bdConn);
                            commandPRJ.ExecuteNonQuery();
                            i = x;
                            break;
                        }
                    }
                }

            }
        }                

        //INICIO DE CADASTRO DE PROJETOS
        void inicioCPRJ()
        {
            limpaCPrj();

            #region ADICIONA NOMES DAS FASES NO COMBOBOX
            try
            {
                cb_F4_CPrj.Items.Clear();
                cb_F5_CPrj.Items.Clear();
                cb_F6_CPrj.Items.Clear();
                cb_F7_CPrj.Items.Clear();
                cb_F8_CPrj.Items.Clear();
                cb_F9_CPrj.Items.Clear();
                cb_F10_CPrj.Items.Clear();

                    bdConn.Open();
                    string queryFASE = "select nom_fase from fase";
                    MySqlCommand commandFASE = new MySqlCommand(queryFASE, bdConn);
                    MySqlDataReader drFASE = commandFASE.ExecuteReader();

                    while (drFASE.Read())
                    {                        
                        cb_F4_CPrj.Items.Add(drFASE["nom_fase"].ToString());
                        cb_F5_CPrj.Items.Add(drFASE["nom_fase"].ToString());
                        cb_F6_CPrj.Items.Add(drFASE["nom_fase"].ToString());
                        cb_F7_CPrj.Items.Add(drFASE["nom_fase"].ToString());
                        cb_F8_CPrj.Items.Add(drFASE["nom_fase"].ToString());
                        cb_F9_CPrj.Items.Add(drFASE["nom_fase"].ToString());
                        cb_F10_CPrj.Items.Add(drFASE["nom_fase"].ToString());
                    }
                    drFASE.Close();
                    bdConn.Close();                
            }
            catch (Exception ex)
            {
                this.Opacity = 0.9;
                ErrorForm erro = new ErrorForm(ex);
                erro.ShowDialog();
                bdConn.Close();
                this.Close();
            }
            #endregion                                           
        }

        //HABILITA CADASTRO DE FASE 4
        private void cb_FaseProjeto_4_CheckedChanged(object sender, EventArgs e)
        {
            if (cb_FaseProjeto_4.Checked)
                panel_FaseProjeto_4.Enabled = true;
            else
            {
                cb_F4_CPrj.Text = null;
                dt_F4_Inicio.Value = DateTime.Today;
                dt_F4_Fim.Value = DateTime.Today;
                cb_Status_F4.Text = null;
                panel_FaseProjeto_4.Enabled = false;
            }

        }

        //HABILITA CADASTRO DE FASE 5
        private void cb_FaseProjeto_5_CheckedChanged(object sender, EventArgs e)
        {
            if (cb_FaseProjeto_5.Checked)
                panel_FaseProjeto_5.Enabled = true;
            else
            {
                cb_F5_CPrj.Text = null;
                dt_F5_Inicio.Value = DateTime.Today;
                dt_F5_Fim.Value = DateTime.Today;
                cb_Status_F5.Text = null;
                panel_FaseProjeto_5.Enabled = false;
            }
        }

        //HABILITA CADASTRO DE FASE 6
        private void cb_FaseProjeto_6_CheckedChanged(object sender, EventArgs e)
        {
            if (cb_FaseProjeto_6.Checked)
                panel_FaseProjeto_6.Enabled = true;
            else
            {
                cb_F6_CPrj.Text = null;
                dt_F6_Inicio.Value = DateTime.Today;
                dt_F6_Fim.Value = DateTime.Today;
                cb_Status_F6.Text = null;
                panel_FaseProjeto_6.Enabled = false;
            }
        }

        //HABILITA CADASTRO DE FASE 7
        private void cb_FaseProjeto_7_CheckedChanged(object sender, EventArgs e)
        {
            if (cb_FaseProjeto_7.Checked)
                panel_FaseProjeto_7.Enabled = true;
            else
            {
                cb_F7_CPrj.Text = null;
                dt_F7_Inicio.Value = DateTime.Today;
                dt_F7_Fim.Value = DateTime.Today;
                cb_Status_F7.Text = null;
                panel_FaseProjeto_7.Enabled = false;
            }
        }

        //HABILITA CADASTRO DE FASE 8
        private void cb_FaseProjeto_8_CheckedChanged(object sender, EventArgs e)
        {
            if (cb_FaseProjeto_8.Checked)
                panel_FaseProjeto_8.Enabled = true;
            else
            {
                cb_F8_CPrj.Text = null;
                dt_F8_Inicio.Value = DateTime.Today;
                dt_F8_Fim.Value = DateTime.Today;
                cb_Status_F8.Text = null;
                panel_FaseProjeto_8.Enabled = false;
            }
        }

        //HABILITA CADASTRO DE FASE 9
        private void cb_FaseProjeto_9_CheckedChanged(object sender, EventArgs e)
        {
            if (cb_FaseProjeto_9.Checked)
                panel_FaseProjeto_9.Enabled = true;
            else
            {
                cb_F9_CPrj.Text = null;
                dt_F9_Inicio.Value = DateTime.Today;
                dt_F9_Fim.Value = DateTime.Today;
                cb_Status_F9.Text = null;
                panel_FaseProjeto_9.Enabled = false;
            }
        }

        //HABILITA CADASTRO DE FASE 10
        private void cb_FaseProjeto_10_CheckedChanged(object sender, EventArgs e)
        {
            if (cb_FaseProjeto_10.Checked)
                panel_FaseProjeto_10.Enabled = true;
            else
            {
                cb_F10_CPrj.Text = null;
                dt_F10_Inicio.Value = DateTime.Today;
                dt_F10_Fim.Value = DateTime.Today;
                cb_Status_F10.Text = null;
                panel_FaseProjeto_10.Enabled = false;
            }
        }

        #endregion

        #region //**************************************** ACOMPANHAMENTO DE PROJETO ****************************************\\

        #region MENU ITENS

        //PROGRAMA
        private void menu_AP_itemPrograma_Click(object sender, EventArgs e)
        {
            panel_AP_StatusPrograma.Visible = true;
            panel_AP_StatusData.Visible = false;
            panel_AP_StatusDF.Visible = false;

            //CARREGA DADOS DO STATUS POR PROGRMA
            exibeStatusPorPGM(gb_AP_Status.Text);
        }

        //DATA
        private void menu_AP_itemData_Click(object sender, EventArgs e)
        {
            panel_AP_StatusData.Visible = true;
            panel_AP_StatusPrograma.Visible = false;
            panel_AP_StatusDF.Visible = false;

            exibeStatusPorDATA(gb_AP_Status.Text);
        }

        //DF
        private void menu_AP_itemDF_Click(object sender, EventArgs e)
        {
            panel_AP_StatusPrograma.Visible = false;
            panel_AP_StatusData.Visible = false;
            panel_AP_StatusDF.Visible = true;

            exibeStatusPorDF(gb_AP_Status.Text);
        }

        #endregion

        //SELECIONA PRJ - PESQUISA AUTOMATICA
        private void tb_selectAcompPRJ_Contrucao_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (tb_selectAcompPRJ_Contrucao.Text != "")
                {
                    bdDataSet = new DataSet();
                    bdConn.Open();
                    bdAdapter = new MySqlDataAdapter("select cod_prj from prj_objeto where cod_prj like '%" + tb_selectAcompPRJ_Contrucao.Text + "%'", bdConn);
                    bdAdapter.Fill(bdDataSet, "prj_objeto");
                    dataGrid_selectAcompPRJ_Contrucao.DataSource = bdDataSet;

                    if (bdDataSet.Tables["prj_objeto"].Rows.Count == 0)
                        lb_NF_selectAcompPRJ_Contrucao.Visible = true;
                    else
                    {
                        lb_NF_selectAcompPRJ_Contrucao.Visible = false;
                        bt_ExpGrid_AcompPRJ_ContrucaoPGM.Visible = true;
                    }

                    dataGrid_selectAcompPRJ_Contrucao.DataMember = "prj_objeto";
                    bdConn.Close();
                }
                else
                {
                    lb_NF_selectAcompPRJ_Contrucao.Visible = false;

                    if (this.dataGrid_selectAcompPRJ_Contrucao.DataSource != null)
                        this.dataGrid_selectAcompPRJ_Contrucao.DataSource = null;
                    else
                    {
                        this.dataGrid_selectAcompPRJ_Contrucao.Rows.Clear();
                        this.dataGrid_selectAcompPRJ_Contrucao.Columns.Clear();
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorForm erro = new ErrorForm(ex);
                erro.ShowDialog();
                bdConn.Close();
            }
        }

        //BOTÃO OK SELECIONA PRJ        
        private void bt_selectAcompPRJ_Contrucao_Click(object sender, EventArgs e)
        {
            if (dataGrid_selectAcompPRJ_Contrucao.CurrentRow == null)
                botaoAlert("Antes de prosseguir selecionar um projeto.");
            else
            {
                if (bdDataSet.Tables["prj_objeto"].Rows.Count > 0)
                {
                    try
                    {                       
                        //SETA TAMANHO DOS PAINEIS DE STATUS
                        panel_AP_StatusPrograma.Dock = DockStyle.Fill;
                        panel_AP_StatusData.Dock = DockStyle.Fill;
                        panel_AP_StatusDF.Dock = DockStyle.Fill;

                        //DEFINE VISIBILIDADE DOS PAINEIS DE STATUS
                        panel_AP_StatusPrograma.Visible = true;
                        panel_AP_StatusData.Visible = false;
                        panel_AP_StatusDF.Visible = false;

                        //HABILITA BOTÕES
                        bt_Atualizar_AcompPRJ.Visible = true;
                        bt_Voltar_AcompPRJ.Visible = true;
                        bt_ExportarExcel_AcompPRJ.Visible = true;
                       
                        //DEFINE NOME DO PROJETO NO GROUPBOX
                        gb_AP_Status.Text = dataGrid_selectAcompPRJ_Contrucao.CurrentRow.Cells[0].Value.ToString();                        

                        //ATIVA PAINEL DE ACOMPANHAMENTO
                        panel_selectAcompPRJ_Contrucao.Visible = false;
                        gb_AP_Status.Visible = true;

                        //CARREGA DADOS DO STATUS POR PROGRMA                        
                        menu_AP_itemPrograma.PerformClick();
                    }
                    catch (Exception ex)
                    {
                        this.Opacity = 0.9;
                        ErrorForm erro = new ErrorForm(ex);
                        erro.ShowDialog();
                        bdConn.Close();
                        this.Close();
                    }
                }
                else
                    botaoAlert("Nenhum projeto foi encontrado! Pesquise novamente.");
            }
        }

        //BOTÃO EXPORTAR EXCEL
        private void bt_ExportarExcel_AcompPRJ_Click(object sender, EventArgs e)
        {
            try
            {
                //OBJETO PARA SALVAR
                SaveFileDialog salvar = new SaveFileDialog();

                //CRIA PLANILHA
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;
                Excel.Workbook workBook = excelApp.Workbooks.Add(); //PASTA
                Excel.Worksheet ws_StatusRQF = (Excel.Worksheet)excelApp.ActiveSheet; //PLANILHA
                Excel.Worksheet ws_StatusDATA = (Excel.Worksheet)excelApp.Sheets.Add(); //PLANILHA            
                Excel.Worksheet ws_StatusPGM = (Excel.Worksheet)excelApp.Sheets.Add(); //PLANILHA            

                //RENOMEIA WORKSHEETS 
                ws_StatusRQF.Name = "Status por RQF";
                ws_StatusPGM.Name = "Status por Programa";
                ws_StatusDATA.Name = "Status por Data";

                //PLANILHA STATUS POR DATA
                excelStatusDATA(ref ws_StatusDATA);

                //PLANILHA STATUS POR PROGRAMA
                excelStatusPGM(ref ws_StatusPGM);

                //CONFIGURAÇÕES PARA SALVAR O ARQUIVO
                salvar.Title = "Exportar para Excel";
                salvar.Filter = "Arquivo do Excel *.xlsx | *.xlsx";
                salvar.ShowDialog(); // mostra

                //SALVA O ARQUIVO                      
                workBook.SaveAs(salvar.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                workBook.Close(true, Type.Missing, Type.Missing);
                excelApp.Quit(); // ENCERRA O EXCEL

                //FINALIZA OPERAÇÃO
                //exibeLOAD();
                botaoConcluido("Planilha gerada com sucesso!");
                //fechaLOAD();
            }
            catch (Exception ex)
            {
                ErrorForm erro = new ErrorForm(ex);
                erro.ShowDialog();
            }
        }

        //CALCULA E EXIBE STATUS DE CONSTRUÇÃO POR PROGRAMAS
        void exibeStatusPorPGM(string projeto)
        {
            limpaGrid_ContrucaoPGM();

            //ABRE CONEXÃO
            bdConn.Open();

            //HEADER DATAGRIDVIEW            
            dataGrid_AcompPRJ_ContrucaoPGM.Columns.Add("sistema", "Sistema");
            dataGrid_AcompPRJ_ContrucaoPGM.Columns.Add("dfrqf", "DF/RQF");
            dataGrid_AcompPRJ_ContrucaoPGM.Columns.Add("resp_cttu", "Responsável CCTU");
            dataGrid_AcompPRJ_ContrucaoPGM.Columns.Add("pgm", "Programa");
            dataGrid_AcompPRJ_ContrucaoPGM.Columns.Add("dt_ini", "DT Inicio");
            dataGrid_AcompPRJ_ContrucaoPGM.Columns.Add("dt_fim", "DT Término");
            dataGrid_AcompPRJ_ContrucaoPGM.Columns.Add("status_pgm", "Status");            
            dataGrid_AcompPRJ_ContrucaoPGM.Columns.Add("total_const", "Total Construído");

            //ESCREVE DATAS NO HEADER DO GRIDVIEW
            MySqlCommand command = new MySqlCommand("SELECT data FROM data_acomp NATURAL JOIN prj_data_acomp WHERE cod_prj = '" + projeto + "'", bdConn);
            MySqlDataReader dr = command.ExecuteReader();
            int qtdData = 0;
            while (dr.Read())
            {
                qtdData++;
                dataGrid_AcompPRJ_ContrucaoPGM.Columns.Add(("data_" + qtdData), dr["data"].ToString());
            }
            dr.Close();

            //RECUPERA QUANTIDADE DE PROGRAMAS NO PRJ
            command = new MySqlCommand("SELECT count(cod_pgm) FROM pgm_objeto WHERE cod_prj = '" + projeto + "'", bdConn);
            dr = command.ExecuteReader();
            int qtdPGM = 0;
            if (dr.Read())
                qtdPGM = Int16.Parse(dr["count(cod_pgm)"].ToString());
            dr.Close();

            //RECUPERA INFORMAÇÕES DOS PROGRAMAS
            string campos = "P.sistema, P.df_rqf, P.resp_cttu, P.cod_pgm, PA.status_construcao, PA.data_inicio, PA.data_fim, sum(PAD.pct_pgm) as pct";
            command = new MySqlCommand("SELECT " + campos + " FROM pgm_objeto P NATURAL JOIN pgm_acompanhamento PA NATURAL JOIN pgm_acomp_data PAD WHERE P.cod_prj = '" + projeto + "' GROUP BY P.df_rqf, P.cod_pgm ORDER BY P.df_rqf, P.cod_pgm;", bdConn);
            dr = command.ExecuteReader();

            string[,] linhaValor = new string[qtdPGM, (8 + qtdData)];
            int index = 0;
            while (dr.Read())
            {                
                linhaValor[index, 0] = dr["sistema"].ToString();
                linhaValor[index, 1] = dr["df_rqf"].ToString();
                linhaValor[index, 2] = dr["resp_cttu"].ToString();
                linhaValor[index, 3] = dr["cod_pgm"].ToString();
                linhaValor[index, 4] = (dr["data_inicio"].ToString() != "00/00/0000") ? dr["data_inicio"].ToString() : "-";
                linhaValor[index, 5] = (dr["data_fim"].ToString() != "00/00/0000") ? dr["data_fim"].ToString() : "-";
                linhaValor[index, 6] = dr["status_construcao"].ToString();                
                linhaValor[index, 7] = (dr["pct"].ToString() == "") ? "0%" : (dr["pct"].ToString() + "%");

                index++;
            }
            dr.Close();

            //RECUPERA PORCENTAGEM DE CONTRUÇÃO PARA CADA PROGRAMA
            int indData = 8;
            string cod_pad = "";
            for (int x = 0; x < qtdPGM; x++)
            {
                command = new MySqlCommand("SELECT cod_pad FROM pgm_acompanhamento WHERE cod_pgm = '" + linhaValor[x, 3] + "' AND df_rqf = '" + linhaValor[x, 1] + "';", bdConn);
                dr = command.ExecuteReader();
                if (dr.Read())
                    cod_pad = dr["cod_pad"].ToString();
                dr.Close();


                command = new MySqlCommand("SELECT pct_pgm FROM pgm_acomp_data WHERE cod_pad = " + cod_pad + ";", bdConn);
                dr = command.ExecuteReader();
                indData = 8;
                while (dr.Read())
                {
                    linhaValor[x, indData] = (dr["pct_pgm"].ToString() == "") ? "-" : (dr["pct_pgm"].ToString() + "%");
                    indData++;
                }
                dr.Close();
            }

            //LISTA DATAS NO GRIDVIEW
            for (int i = 0; i < qtdPGM; i++)
                switch (qtdData)
                {
                    #region VERIFICAÇÕES
                    case 1:
                        #region
                        dataGrid_AcompPRJ_ContrucaoPGM.Rows.Add(
                        linhaValor[i, 0],
                        linhaValor[i, 1],
                        linhaValor[i, 2],
                        linhaValor[i, 3],
                        linhaValor[i, 4],
                        linhaValor[i, 5],
                        linhaValor[i, 6],
                        linhaValor[i, 7],
                        linhaValor[i, 8]
                        );
                        #endregion
                        break;
                    case 2:
                        #region
                        dataGrid_AcompPRJ_ContrucaoPGM.Rows.Add(
                        linhaValor[i, 0],
                        linhaValor[i, 1],
                        linhaValor[i, 2],
                        linhaValor[i, 3],
                        linhaValor[i, 4],
                        linhaValor[i, 5],
                        linhaValor[i, 6],
                        linhaValor[i, 7],
                        linhaValor[i, 8],
                        linhaValor[i, 9]
                        );
                        #endregion
                        break;
                    case 3:
                        #region
                        dataGrid_AcompPRJ_ContrucaoPGM.Rows.Add(
                        linhaValor[i, 0],
                        linhaValor[i, 1],
                        linhaValor[i, 2],
                        linhaValor[i, 3],
                        linhaValor[i, 4],
                        linhaValor[i, 5],
                        linhaValor[i, 6],
                        linhaValor[i, 7],
                        linhaValor[i, 8],
                        linhaValor[i, 9],
                        linhaValor[i, 10]
                        );
                        #endregion
                        break;
                    case 4:
                        #region
                        dataGrid_AcompPRJ_ContrucaoPGM.Rows.Add(
                        linhaValor[i, 0],
                        linhaValor[i, 1],
                        linhaValor[i, 2],
                        linhaValor[i, 3],
                        linhaValor[i, 4],
                        linhaValor[i, 5],
                        linhaValor[i, 6],
                        linhaValor[i, 7],
                        linhaValor[i, 8],
                        linhaValor[i, 9],
                        linhaValor[i, 10],
                        linhaValor[i, 11]
                        );
                        #endregion
                        break;
                    case 5:
                        #region
                        dataGrid_AcompPRJ_ContrucaoPGM.Rows.Add(
                        linhaValor[i, 0],
                        linhaValor[i, 1],
                        linhaValor[i, 2],
                        linhaValor[i, 3],
                        linhaValor[i, 4],
                        linhaValor[i, 5],
                        linhaValor[i, 6],
                        linhaValor[i, 7],
                        linhaValor[i, 8],
                        linhaValor[i, 9],
                        linhaValor[i, 10],
                        linhaValor[i, 11],
                        linhaValor[i, 12]
                        );
                        #endregion
                        break;
                    case 6:
                        #region
                        dataGrid_AcompPRJ_ContrucaoPGM.Rows.Add(
                        linhaValor[i, 0],
                        linhaValor[i, 1],
                        linhaValor[i, 2],
                        linhaValor[i, 3],
                        linhaValor[i, 4],
                        linhaValor[i, 5],
                        linhaValor[i, 6],
                        linhaValor[i, 7],
                        linhaValor[i, 8],
                        linhaValor[i, 9],
                        linhaValor[i, 10],
                        linhaValor[i, 11],
                        linhaValor[i, 12],
                        linhaValor[i, 13]
                        );
                        #endregion
                        break;
                    case 7:
                        #region
                        dataGrid_AcompPRJ_ContrucaoPGM.Rows.Add(
                        linhaValor[i, 0],
                        linhaValor[i, 1],
                        linhaValor[i, 2],
                        linhaValor[i, 3],
                        linhaValor[i, 4],
                        linhaValor[i, 5],
                        linhaValor[i, 6],
                        linhaValor[i, 7],
                        linhaValor[i, 8],
                        linhaValor[i, 9],
                        linhaValor[i, 10],
                        linhaValor[i, 11],
                        linhaValor[i, 12],
                        linhaValor[i, 13],
                        linhaValor[i, 14]
                        );
                        #endregion
                        break;
                    case 8:
                        #region
                        dataGrid_AcompPRJ_ContrucaoPGM.Rows.Add(
                        linhaValor[i, 0],
                        linhaValor[i, 1],
                        linhaValor[i, 2],
                        linhaValor[i, 3],
                        linhaValor[i, 4],
                        linhaValor[i, 5],
                        linhaValor[i, 6],
                        linhaValor[i, 7],
                        linhaValor[i, 8],
                        linhaValor[i, 9],
                        linhaValor[i, 10],
                        linhaValor[i, 11],
                        linhaValor[i, 12],
                        linhaValor[i, 13],
                        linhaValor[i, 14],
                        linhaValor[i, 15]
                        );
                        #endregion
                        break;
                    case 9:
                        #region
                        dataGrid_AcompPRJ_ContrucaoPGM.Rows.Add(
                        linhaValor[i, 0],
                        linhaValor[i, 1],
                        linhaValor[i, 2],
                        linhaValor[i, 3],
                        linhaValor[i, 4],
                        linhaValor[i, 5],
                        linhaValor[i, 6],
                        linhaValor[i, 7],
                        linhaValor[i, 8],
                        linhaValor[i, 9],
                        linhaValor[i, 10],
                        linhaValor[i, 11],
                        linhaValor[i, 12],
                        linhaValor[i, 13],
                        linhaValor[i, 14],
                        linhaValor[i, 15],
                        linhaValor[i, 16]
                        );
                        #endregion
                        break;
                    case 10:
                        #region
                        dataGrid_AcompPRJ_ContrucaoPGM.Rows.Add(
                        linhaValor[i, 0],
                        linhaValor[i, 1],
                        linhaValor[i, 2],
                        linhaValor[i, 3],
                        linhaValor[i, 4],
                        linhaValor[i, 5],
                        linhaValor[i, 6],
                        linhaValor[i, 7],
                        linhaValor[i, 8],
                        linhaValor[i, 9],
                        linhaValor[i, 10],
                        linhaValor[i, 11],
                        linhaValor[i, 12],
                        linhaValor[i, 13],
                        linhaValor[i, 14],
                        linhaValor[i, 15],
                        linhaValor[i, 16],
                        linhaValor[i, 17]
                        );
                        #endregion
                        break;
                    case 11:
                        #region
                        dataGrid_AcompPRJ_ContrucaoPGM.Rows.Add(
                        linhaValor[i, 0],
                        linhaValor[i, 1],
                        linhaValor[i, 2],
                        linhaValor[i, 3],
                        linhaValor[i, 4],
                        linhaValor[i, 5],
                        linhaValor[i, 6],
                        linhaValor[i, 7],
                        linhaValor[i, 8],
                        linhaValor[i, 9],
                        linhaValor[i, 10],
                        linhaValor[i, 11],
                        linhaValor[i, 12],
                        linhaValor[i, 13],
                        linhaValor[i, 14],
                        linhaValor[i, 15],
                        linhaValor[i, 16],
                        linhaValor[i, 17],
                        linhaValor[i, 18]
                        );
                        #endregion
                        break;
                    case 12:
                        #region
                        dataGrid_AcompPRJ_ContrucaoPGM.Rows.Add(
                        linhaValor[i, 0],
                        linhaValor[i, 1],
                        linhaValor[i, 2],
                        linhaValor[i, 3],
                        linhaValor[i, 4],
                        linhaValor[i, 5],
                        linhaValor[i, 6],
                        linhaValor[i, 7],
                        linhaValor[i, 8],
                        linhaValor[i, 9],
                        linhaValor[i, 10],
                        linhaValor[i, 11],
                        linhaValor[i, 12],
                        linhaValor[i, 13],
                        linhaValor[i, 14],
                        linhaValor[i, 15],
                        linhaValor[i, 16],
                        linhaValor[i, 17],
                        linhaValor[i, 18],
                        linhaValor[i, 19]
                        );
                        #endregion
                        break;
                    case 13:
                        #region
                        dataGrid_AcompPRJ_ContrucaoPGM.Rows.Add(
                        linhaValor[i, 0],
                        linhaValor[i, 1],
                        linhaValor[i, 2],
                        linhaValor[i, 3],
                        linhaValor[i, 4],
                        linhaValor[i, 5],
                        linhaValor[i, 6],
                        linhaValor[i, 7],
                        linhaValor[i, 8],
                        linhaValor[i, 9],
                        linhaValor[i, 10],
                        linhaValor[i, 11],
                        linhaValor[i, 12],
                        linhaValor[i, 13],
                        linhaValor[i, 14],
                        linhaValor[i, 15],
                        linhaValor[i, 16],
                        linhaValor[i, 17],
                        linhaValor[i, 18],
                        linhaValor[i, 19],
                        linhaValor[i, 20]
                        );
                        #endregion
                        break;
                    case 14:
                        #region
                        dataGrid_AcompPRJ_ContrucaoPGM.Rows.Add(
                        linhaValor[i, 0],
                        linhaValor[i, 1],
                        linhaValor[i, 2],
                        linhaValor[i, 3],
                        linhaValor[i, 4],
                        linhaValor[i, 5],
                        linhaValor[i, 6],
                        linhaValor[i, 7],
                        linhaValor[i, 8],
                        linhaValor[i, 9],
                        linhaValor[i, 10],
                        linhaValor[i, 11],
                        linhaValor[i, 12],
                        linhaValor[i, 13],
                        linhaValor[i, 14],
                        linhaValor[i, 15],
                        linhaValor[i, 16],
                        linhaValor[i, 17],
                        linhaValor[i, 18],
                        linhaValor[i, 19],
                        linhaValor[i, 20],
                        linhaValor[i, 21]
                        );
                        #endregion
                        break;
                    default:
                        #region
                        dataGrid_AcompPRJ_ContrucaoPGM.Rows.Add(
                        linhaValor[i, 0],
                        linhaValor[i, 1],
                        linhaValor[i, 2],
                        linhaValor[i, 3],
                        linhaValor[i, 4],
                        linhaValor[i, 5],
                        linhaValor[i, 6],
                        linhaValor[i, 7],
                        linhaValor[i, 8],
                        linhaValor[i, 9],
                        linhaValor[i, 10],
                        linhaValor[i, 11],
                        linhaValor[i, 12],
                        linhaValor[i, 13],
                        linhaValor[i, 14],
                        linhaValor[i, 15],
                        linhaValor[i, 16],
                        linhaValor[i, 17],
                        linhaValor[i, 18],
                        linhaValor[i, 19],
                        linhaValor[i, 20],
                        linhaValor[i, 21],
                        linhaValor[i, 22]
                        );
                        #endregion
                        break;
                    #endregion
                }

            //FECHA CONEXÃO
            bdConn.Close();
        }

        //CALCULA E EXIBE STATUS DE CONSTRUÇÃO POR DATA
        void exibeStatusPorDATA(string projeto)
        {
            //LIMPA GRIDVIEW
            limpaGrid_ContrucaoDATA();

            #region DECLARAÇÃO DE VARIÁVEIS
            List<string> Datas = new List<string>();
            double realizadoDia = 0;
            double realizadoDiaAcumulado = 0;
            double qtdPGM = 0;
            #endregion

            //HEADER DATAGRIDVIEW
            dataGrid_AcompPRJ_ContrucaoDATA.Columns.Add("data", "Data");
            dataGrid_AcompPRJ_ContrucaoDATA.Columns.Add("atividade", "Atividade");
            dataGrid_AcompPRJ_ContrucaoDATA.Columns.Add("previsto", "% Previsto");
            dataGrid_AcompPRJ_ContrucaoDATA.Columns.Add("realizado", "% Realizado");
            dataGrid_AcompPRJ_ContrucaoDATA.Columns.Add("status", "Status");

            //ABRE CONEXÃO
            bdConn.Open();

            #region RECUPERA INFORMAÇOES GERAIS

            //DATAS DE INICIO E TERMINO CCTU
            MySqlCommand command = new MySqlCommand("SELECT dt_ini, dt_fim FROM prj_fase PF NATURAL JOIN fase F WHERE PF.cod_prj = '" + projeto + "' AND F.nom_fase = 'Construção';", bdConn);
            MySqlDataReader dr = command.ExecuteReader();
            if (dr.Read())
            {
                tb_AP_Data_dtinicio.Text = dr["dt_ini"].ToString();
                tb_AP_Data_dttermino.Text = dr["dt_fim"].ToString();

                //CALCULA TOTAL DE DIAS ÚTEIS
                tb_AP_Data_totaldias.Text = CountDiasUteis(DateTime.Parse(dr["dt_ini"].ToString()), DateTime.Parse(dr["dt_fim"].ToString())).ToString();
            }
            dr.Close();

            //LIDER TÉCNICO E LIDER DE REQUERIMENTO
            command = new MySqlCommand("SELECT lider_tecnico, lider_requer FROM prj_objeto WHERE cod_prj = '" + projeto + "';", bdConn);
            dr = command.ExecuteReader();
            if (dr.Read())
            {
                tb_AP_Data_LT.Text = dr["lider_tecnico"].ToString();
                tb_AP_Data_LR.Text = dr["lider_requer"].ToString();
            }
            dr.Close();

            #region CALCULA PORCENTAGEM DO PROJETO POR PESO

            //TOTAL DE PESO
            double totalPeso = 0;
            command = new MySqlCommand("SELECT sum(peso_pgm) FROM pgm_objeto WHERE cod_prj = '" + projeto + "';", bdConn);
            dr = command.ExecuteReader();
            if (dr.Read())
                totalPeso = ((dr["sum(peso_pgm)"].ToString() != "") ? Double.Parse(dr["sum(peso_pgm)"].ToString()) : 0);
            dr.Close();

            //VARIAVEIS
            double totalAcumProjeto = 0;
            double pesoPGM = 0;
            double totalConstPGM = 0;
            string queryPeso = "SELECT PO.peso_pgm, sum(PAD.pct_pgm) FROM " +
                                    "pgm_objeto PO INNER JOIN " +
                                    "pgm_acompanhamento PA ON " +
                                    "PO.cod_pgm = PA.cod_pgm AND " +
                                    "PO.df_rqf = PA.df_rqf " +
                                    "INNER JOIN pgm_acomp_data PAD ON " +
                                    "PA.cod_pad = PAD.cod_pad " +
                                    "WHERE PO.cod_prj = '" + projeto + "' " +
                                    "GROUP BY PO.cod_pgm, PO.df_rqf;";
            
            //CALCULA TOTAL ACUMULADO DO PROJETO
            command = new MySqlCommand(queryPeso, bdConn);
            dr = command.ExecuteReader();
            while (dr.Read())
            {
                pesoPGM = ((dr[0].ToString() != "") ? Double.Parse(dr[0].ToString()) : 0);
                totalConstPGM = ((dr[1].ToString() != "") ? Double.Parse(dr[1].ToString()) : 0);

                totalAcumProjeto += (pesoPGM / totalPeso) * totalConstPGM;
            }                
            dr.Close();

            //ALTERA LABEL E ESCREVE PORCENTAGEM NO TEXTBOX
            lb_AP_Data_Realizado.Text = "Realizado até " + DateTime.Now.ToShortDateString() + ":";
            tb_AP_Data_Realizado.Text = Math.Round(totalAcumProjeto, 2).ToString() + "%";

            #endregion

            #endregion

            #region RECUPERA DATAS DO PROJETO
            command = new MySqlCommand("SELECT data FROM data_acomp NATURAL JOIN prj_data_acomp WHERE cod_prj = '" + projeto + "'", bdConn);
            dr = command.ExecuteReader();
            while (dr.Read())
                Datas.Add(dr["data"].ToString());
            dr.Close();
            #endregion


            foreach (var data in Datas)
            {
                //CALCULA PORCENTAGEM PREVISTA PARA DIA
                double auxPrevisto = Double.Parse((CountDiasUteis(DateTime.Parse(tb_AP_Data_dtinicio.Text), DateTime.Parse(data))).ToString());
                double previsto = (auxPrevisto / Double.Parse(tb_AP_Data_totaldias.Text)) * 100;

                #region CALCULA REALIZADO DO DIA

                //SOMA TODAS AS PORCENTAGEM DO DIA
                realizadoDia = 0;
                string[] formatData = Regex.Split(data, "/");
                command = new MySqlCommand("SELECT sum(pct_pgm) FROM pgm_acomp_data WHERE data_acomp = '" + formatData[2] + "-" + formatData[1] + "-" + formatData[0] + "'", bdConn);
                dr = command.ExecuteReader();
                if (dr.Read())
                    realizadoDia = (dr["sum(pct_pgm)"].ToString() != "" ? Double.Parse(dr["sum(pct_pgm)"].ToString()) : 0);
                dr.Close();

                //QUANTIDADE DE PROGRAMAS
                qtdPGM = 0;
                command = new MySqlCommand("SELECT count(cod_pgm) FROM pgm_acompanhamento WHERE cod_prj = '" + projeto + "'", bdConn);
                dr = command.ExecuteReader();
                if (dr.Read())
                    qtdPGM = Double.Parse(dr["count(cod_pgm)"].ToString());
                dr.Close();

                //CALCULA A PORCENTAGEM REALIZADA DO DIA
                realizadoDiaAcumulado += (realizadoDia / qtdPGM);

                #endregion

                //ESCREVE LINHA NO GRIDVIEW
                dataGrid_AcompPRJ_ContrucaoDATA.Rows.Add(
                    data,
                    "CCTU",
                    ((Math.Round(previsto, 2).ToString()) + "%"),
                    (Math.Round(realizadoDiaAcumulado, 2) + "%"),
                    ((realizadoDiaAcumulado < previsto) ? "Atrasado" : "Em Dia")
                    );
            }


            //FECHA CONEXÃO
            bdConn.Close();

        }

        //CALCULA E EXIBE STATUS DE CONSTRUÇÃO POR DF
        void exibeStatusPorDF(string projeto)
        {
            //LIMPA GRIDVIEW
            limpaGrid_ContrucaoDF();

            //VARIAVEIS
            List<string> RQFs = new List<string>();
            string analista = "";
            string sistema = "";
            double pesoTotalRQF = 0;
            double realizadoRQF = 0;
            double realizadoRQFAcumulado = 0;

            //HEADER DATAGRIDVIEW
            dataGrid_AcompPRJ_ContrucaoDF.Columns.Add("analista", "Analista");
            dataGrid_AcompPRJ_ContrucaoDF.Columns.Add("sistema", "Sistema");
            dataGrid_AcompPRJ_ContrucaoDF.Columns.Add("df", "DF's");
            dataGrid_AcompPRJ_ContrucaoDF.Columns.Add("percentualR", "Percentual Realizado");

            //ABRE CONEXÃO
            bdConn.Open();            

            //RECUPERA RQF
            MySqlCommand command = new MySqlCommand("SELECT distinct df_rqf FROM pgm_objeto WHERE cod_prj = '" + projeto + "' ORDER BY df_rqf;", bdConn);
            MySqlDataReader dr = command.ExecuteReader();
            while (dr.Read())
                RQFs.Add(dr["df_rqf"].ToString());
            dr.Close();            

            //PERCORRE A LISTA DE RQF
            foreach (var rqf in RQFs)
            {
                //RECUPERA ANALISTA E SISTEMA
                analista = "";
                sistema = "";
                command = new MySqlCommand("SELECT resp_cttu, sistema FROM pgm_objeto WHERE cod_prj = '" + projeto + "' AND df_rqf = '" + rqf + "';", bdConn);
                dr = command.ExecuteReader();
                if (dr.Read())
                {
                    analista = dr["resp_cttu"].ToString();
                    sistema = dr["sistema"].ToString();
                }                    
                dr.Close();

                //RECUPERA TOTAL DE PESO DA RQF
                pesoTotalRQF = 0;
                command = new MySqlCommand("SELECT sum(peso_pgm) FROM pgm_objeto WHERE cod_prj = '" + projeto + "' AND df_rqf = '" + rqf + "';", bdConn);
                dr = command.ExecuteReader();
                if (dr.Read())
                    pesoTotalRQF = (dr[0].ToString() != "") ? Int16.Parse(dr[0].ToString()) : 0;
                dr.Close();

                //RECUPERA PESO E CONTRUÇÃO DO PROGRAMA
                realizadoRQF = 0;
                command = new MySqlCommand("SELECT PO.peso_pgm, sum(PAD.pct_pgm) FROM pgm_objeto PO NATURAL JOIN pgm_acompanhamento PA NATURAL JOIN pgm_acomp_data PAD WHERE PO.cod_prj = '" + projeto + "' AND PO.df_rqf = '" + rqf + "' GROUP BY PO.cod_pgm;", bdConn);
                dr = command.ExecuteReader();
                while (dr.Read())
                {
                    double pesoPGM = (dr[0].ToString() != "") ? Int16.Parse(dr[0].ToString()) : 0;
                    double porcentagemPGM = (dr[1].ToString() != "") ? Int16.Parse(dr[1].ToString()) : 0;

                    realizadoRQF += ((pesoPGM / pesoTotalRQF) * porcentagemPGM);
                }
                dr.Close();

                //ACUMULADO DE TOTAS RQF
                realizadoRQFAcumulado += realizadoRQF;

                //ESCREVE LINHA NO GRIDVIEW
                dataGrid_AcompPRJ_ContrucaoDF.Rows.Add(
                    analista,
                    sistema,
                    rqf,
                    Math.Round(realizadoRQF, 2).ToString() + "%"
                    );
            }

            //TOTALIZADORES
            tb_totalRQF.Text = RQFs.Count.ToString();
            tb_realizadoRQF.Text = Math.Round((realizadoRQFAcumulado/RQFs.Count), 2).ToString() + "%";

            //FECHA CONEXÃO
            bdConn.Close();

        }

        //EXCEL STATUS POR DATA
        void excelStatusDATA(ref Excel.Worksheet ws_StatusDATA)
        {
            //STATUS POR DATA
                Excel.Range selectData;

                ws_StatusDATA.Cells[1, "A"] = "Acompanhamento da Especificação Técnica";

                ws_StatusDATA.Cells[2, "A"] = "Data Inicio";
                ws_StatusDATA.Cells[2, "B"] = tb_AP_Data_dtinicio.Text;

                ws_StatusDATA.Cells[3, "A"] = "Data Final";
                ws_StatusDATA.Cells[3, "B"] = tb_AP_Data_dttermino.Text;

                ws_StatusDATA.Cells[4, "A"] = "Total de Dias";
                ws_StatusDATA.Cells[4, "B"] = tb_AP_Data_totaldias.Text;

                ws_StatusDATA.Cells[5, "A"] = "LT";
                ws_StatusDATA.Cells[5, "B"] = tb_AP_Data_LT.Text;

                ws_StatusDATA.Cells[6, "A"] = "LR";
                ws_StatusDATA.Cells[6, "B"] = tb_AP_Data_LR.Text;

                ws_StatusDATA.Cells[8, "A"] = "Realizado em:";
                ws_StatusDATA.Cells[8, "B"] = DateTime.Now.ToShortDateString();
                ws_StatusDATA.Cells[8, "C"] = tb_AP_Data_Realizado.Text;

                ws_StatusDATA.Cells[10, "A"] = "Previsto em";
                ws_StatusDATA.Cells[10, "B"] = "Atividade";
                ws_StatusDATA.Cells[10, "C"] = "% Previsto";
                ws_StatusDATA.Cells[10, "D"] = "% Realizado";
                ws_StatusDATA.Cells[10, "E"] = "Status";


                int indRows = 11;
                for (int i = 0; i <= dataGrid_AcompPRJ_ContrucaoDATA.RowCount - 1; i++)
                {
                    for (int j = 0; j <= dataGrid_AcompPRJ_ContrucaoDATA.ColumnCount - 1; j++)
                    {
                        DataGridViewCell cell = dataGrid_AcompPRJ_ContrucaoDATA[j, i];
                        ws_StatusDATA.Cells[indRows, (j + 1)] = cell.Value.ToString();

                        if (cell.Value.ToString().Equals("Atrasado"))
                        {
                            selectData = ws_StatusDATA.get_Range(("D" + indRows.ToString()), ("E" + indRows.ToString()));
                            selectData.Interior.Color = 255;
                        }

                        if (cell.Value.ToString().Equals("Em Dia"))
                        {
                            selectData = ws_StatusDATA.get_Range(("D" + indRows.ToString()), ("E" + indRows.ToString()));
                            selectData.Interior.Color = 5287936;
                        }

                    }

                    indRows++;
                }

                #region FORMATAÇÃO TABELAS STATUS POR DATA
                //-------------------------------------------------------------------------------

                //TITULO
                selectData = ws_StatusDATA.get_Range("A1", "F1");
                selectData.Merge(Type.Missing);
                selectData.Font.Size = 15;
                selectData.Font.Bold = true;
                selectData.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                selectData.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent6;

                //-------------------------------------------------------------------------------
                //DATA INCIO
                selectData = ws_StatusDATA.get_Range("B2", "F2");
                selectData.Merge(Type.Missing);
                selectData.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                //DATA FINAL
                selectData = ws_StatusDATA.get_Range("B3", "F3");
                selectData.Merge(Type.Missing);
                selectData.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                //TOTAL DIAS
                selectData = ws_StatusDATA.get_Range("B4", "F4");
                selectData.Merge(Type.Missing);
                selectData.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                //COR DIAS            
                selectData = ws_StatusDATA.get_Range("A2", "F4");
                selectData.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorDark2;

                //-------------------------------------------------------------------------------

                //LIDER TECNICO
                selectData = ws_StatusDATA.get_Range("B5", "F5");
                selectData.Merge(Type.Missing);
                selectData.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                //LIDER REQUERIMENTO
                selectData = ws_StatusDATA.get_Range("B6", "F6");
                selectData.Merge(Type.Missing);
                selectData.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                //COR LIDER
                selectData = ws_StatusDATA.get_Range("A5", "F6");
                selectData.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1;

                //-------------------------------------------------------------------------------

                //LINHA ENTRE AS COLUNAS 
                selectData = ws_StatusDATA.get_Range("A1", "F6");
                selectData.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                //-------------------------------------------------------------------------------

                //PORCENTAGEM REALIZADO ATÉ O DIA
                selectData = ws_StatusDATA.get_Range("A8", "C8");
                selectData.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                selectData.Font.Size = 15;
                selectData.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                selectData.Font.Bold = true;
                selectData.Interior.Color = 12611584;

                selectData = ws_StatusDATA.get_Range("B8", "C8");
                selectData.Font.Color = -16711681;

                //-------------------------------------------------------------------------------

                //HEADER DO DETALHAMENTO
                selectData = ws_StatusDATA.get_Range("A10", "E10");
                selectData.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                selectData.Font.Bold = true;
                selectData.Interior.Color = 12611584;
                selectData.Font.Color = -16711681;

                //DETALHAMENTO DAS DATAS
                selectData = ws_StatusDATA.get_Range("A11", ("E" + indRows.ToString()));
                selectData.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                selectData.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                //LINHA ENTRE AS COLUNAS DAS DATAS
                selectData = ws_StatusDATA.get_Range("A10", ("E" + (indRows - 1).ToString()));
                selectData.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                //-------------------------------------------------------------------------------

                //CONFIG TABELAS
                ws_StatusDATA.Columns.AutoFit();

                //-------------------------------------------------------------------------------
                #endregion
        }

        //EXCEL STATUS POR PROGRAMA
        void excelStatusPGM(ref Excel.Worksheet ws_StatusPGM)
        {
            //STATUS POR PROGRAMA
            Excel.Range selectPrograma;

            ws_StatusPGM.Cells[1, "J"] = "Controle de Status/Porcentagem prevista para as datas abaixo:";

            ws_StatusPGM.Cells[2, "J"] = "Fase->";
            ws_StatusPGM.Cells[3, "J"] = "Realizado (acumulado)->";
            ws_StatusPGM.Cells[4, "J"] = "Realizado (no dia)->";            




            //-------------------------------------------------------------------------------

            //CONFIG TABELAS
            ws_StatusPGM.Columns.AutoFit();

            //-------------------------------------------------------------------------------            
        }

        //DUPLO CLICK GRIDVIEW - SELECIONA PROJETO
        private void dataGrid_selectAcompPRJ_Contrucao_DoubleClick(object sender, EventArgs e)
        {
            bt_selectAcompPRJ_Contrucao.PerformClick();
        }

        //EXPANDIR GRIDVIEW
        private void bt_ExpGrid_AcompPRJ_ContrucaoPGM_Click(object sender, EventArgs e)
        {
            //REDIMENSIONA TAMANHO DAS COLUNAS
            dataGrid_AcompPRJ_ContrucaoPGM.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;

            bt_MinGrid_AcompPRJ_ContrucaoPGM.Visible = true;
            bt_ExpGrid_AcompPRJ_ContrucaoPGM.Visible = false;
        }

        //COMPRIMIR GRIDVIEW
        private void bt_MinGrid_AcompPRJ_ContrucaoPGM_Click(object sender, EventArgs e)
        {
            //REDIMENSIONA TAMANHO DAS COLUNAS
            dataGrid_AcompPRJ_ContrucaoPGM.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            bt_MinGrid_AcompPRJ_ContrucaoPGM.Visible = false;
            bt_ExpGrid_AcompPRJ_ContrucaoPGM.Visible = true;
        }

        //BOTÃO ATUALIZAR
        private void bt_Atualizar_AcompPRJ_Click(object sender, EventArgs e)
        {
            if (panel_AP_StatusPrograma.Visible)
                exibeStatusPorPGM(gb_AP_Status.Text);
            else if (panel_AP_StatusData.Visible)
                exibeStatusPorDATA(gb_AP_Status.Text);
            else if (panel_AP_StatusDF.Visible)
                exibeStatusPorDF(gb_AP_Status.Text);
        }

        //LIMPAR GRIDVIEW PROGRAMA
        public void limpaGrid_ContrucaoPGM()
        {
            if (this.dataGrid_AcompPRJ_ContrucaoPGM.DataSource != null)
                this.dataGrid_AcompPRJ_ContrucaoPGM.DataSource = null;
            else
            {
                this.dataGrid_AcompPRJ_ContrucaoPGM.Rows.Clear();
                this.dataGrid_AcompPRJ_ContrucaoPGM.Columns.Clear();
            }
        }

        //LIMPAR GRIDVIEW DATA
        public void limpaGrid_ContrucaoDATA()
        {
            if (this.dataGrid_AcompPRJ_ContrucaoDATA.DataSource != null)
                this.dataGrid_AcompPRJ_ContrucaoDATA.DataSource = null;
            else
            {
                this.dataGrid_AcompPRJ_ContrucaoDATA.Rows.Clear();
                this.dataGrid_AcompPRJ_ContrucaoDATA.Columns.Clear();
            }
        }

        //LIMPAR GRIDVIEW DF
        public void limpaGrid_ContrucaoDF()
        {
            if (this.dataGrid_AcompPRJ_ContrucaoDF.DataSource != null)
                this.dataGrid_AcompPRJ_ContrucaoDF.DataSource = null;
            else                                 
            {                                    
                this.dataGrid_AcompPRJ_ContrucaoDF.Rows.Clear();
                this.dataGrid_AcompPRJ_ContrucaoDF.Columns.Clear();
            }
        }

        //VOLTA PARA ESCOLHER NOVO PROJETO
        private void bt_Voltar_AcompPRJ_Click(object sender, EventArgs e)
        {
            DialogResult resultReturnAPGM = MessageBox.Show("Deseja voltar e escolher outro Projeto?", "Voltar", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (resultReturnAPGM == DialogResult.Yes)
                inicioAcompPRJ();
        }

        void inicioAcompPRJ()
        {
            limpaGrid_ContrucaoPGM();
            limpaGrid_ContrucaoDATA();
            limpaGrid_ContrucaoDF();

            panel_selectAcompPRJ_Contrucao.Visible = true;
            gb_AP_Status.Visible = false;
            gb_AP_Status.Text = "";
            bt_Atualizar_AcompPRJ.Visible = false;
            bt_Voltar_AcompPRJ.Visible = false;
            bt_ExportarExcel_AcompPRJ.Visible = false;
            tb_selectAcompPRJ_Contrucao.Text = "";
        }

        //RETORNA QUANTIDADE DE DIAS ÚTEIS
        public static int CountDiasUteis(DateTime d1, DateTime d2)
        {
            List<DateTime> datas = new List<DateTime>();          
            FeriadosBH.Feriados feriadosBH = new FeriadosBH.Feriados();
            Dictionary<DateTime, string> feriados = new Dictionary<DateTime, string>(feriadosBH.getListaFeriados());
            
            for (DateTime i = d1.Date; i <= d2.Date; i = i.AddDays(1))
                if (!feriados.ContainsKey(i))
                    datas.Add(i);
            
            return datas.Count(d => d.DayOfWeek != DayOfWeek.Saturday && d.DayOfWeek != DayOfWeek.Sunday);
        }

        //FORMATA GRIDVIEW DATA
        private void dataGrid_AcompPRJ_ContrucaoDATA_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //Status
            if (e.Value != null && e.ColumnIndex == 4)
                if (e.Value.Equals("Atrasado"))
                    e.CellStyle.BackColor = Color.IndianRed;
                else
                    e.CellStyle.BackColor = Color.LightGreen;
        }

        #endregion

        #region //**************************************** CADASTRAR RQF ****************************************\\

        //VARIAVEIS UTILIZADAS
        string querySPRJ_CadastroRQF;
        bool verificaCampos_CRQF = true;

        //BOTÃO SALVAR RQF
        private void bt_Salvar_RQF_Click(object sender, EventArgs e)
        {
            try
            {
                validaCAMPOS_CRQF();

                if (verificaCampos_CRQF == true)
                {
                    bdConn.Open();
                    cadastraRQF();
                    bdConn.Close();
                    exibeLOAD();
                    botaoConcluido("Tudo certo. A RQF foi cadastrado com sucesso.");
                    fechaLOAD();
                    bt_Limpar_RQF.PerformClick();
                }
                else
                {
                    botaoAlert("O Preenchimento dos campos são obrigatório!");
                    verificaCampos_CRQF = true;
                }
            }
            catch (Exception ex)
            {
                ErrorForm erro = new ErrorForm(ex);
                erro.ShowDialog();
                bdConn.Close();
            }
        }

        //CADASTRO RQF
        private void cadastraRQF()
        {
            string queryCadastoRQF;

            queryCadastoRQF = "INSERT INTO df_rqf VALUES('" +
               tb_RQF_CadastroRQF.Text + "','" +
               tb_Prj_CadastroRQF.Text + "','" +
               rtb_DescRQF_CadastroRQF.Text +
               "');";

            MySqlCommand commandRQF = new MySqlCommand(queryCadastoRQF, bdConn);
            commandRQF.ExecuteNonQuery();
        }

        //VALIDAÇÃO DOS CAMPOS
        public void validaCAMPOS_CRQF()
        {
            //RQF
            if (tb_RQF_CadastroRQF.Text == "RQF" || tb_RQF_CadastroRQF.Text == "RQNF")
            {
                verificaCampos_CRQF = false;
                gb_RQF_CadastroRQF.ForeColor = Color.Red;
            }
            else
                gb_RQF_CadastroRQF.ForeColor = Color.Black;

            //DESCRIÇÃO DA RQF
            if (rtb_DescRQF_CadastroRQF.Text == "")
            {
                verificaCampos_CRQF = false;
                lb_DescDF_CadastroRQF.ForeColor = Color.Red;
            }
            else
                lb_DescDF_CadastroRQF.ForeColor = Color.Black;
        }

        //SELECIONA PRJ - PESQUISA AUTOMATICA
        private void tb_SPrj_CadastroRQF_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (tb_SPrj_CadastroRQF.Text != "")
                {
                    querySPRJ_CadastroRQF = "select cod_prj from prj_objeto where cod_prj like '%" + tb_SPrj_CadastroRQF.Text + "%'";
                    bdDataSet = new DataSet();
                    bdConn.Open();
                    bdAdapter = new MySqlDataAdapter(querySPRJ_CadastroRQF, bdConn);
                    bdAdapter.Fill(bdDataSet, "prj_objeto");
                    dataGrid_SPrj_CadastroRQF.DataSource = bdDataSet;

                    if (bdDataSet.Tables["prj_objeto"].Rows.Count == 0)
                        lb_ProjetoNaoEncontrado_CRQF.Visible = true;
                    else
                        lb_ProjetoNaoEncontrado_CRQF.Visible = false;

                    dataGrid_SPrj_CadastroRQF.DataMember = "prj_objeto";
                    bdConn.Close();
                }
                else
                {
                    lb_ProjetoNaoEncontrado_CRQF.Visible = false;

                    if (this.dataGrid_SPrj_CadastroRQF.DataSource != null)
                        this.dataGrid_SPrj_CadastroRQF.DataSource = null;
                    else
                    {
                        this.dataGrid_SPrj_CadastroRQF.Rows.Clear();
                        this.dataGrid_SPrj_CadastroRQF.Columns.Clear();
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorForm erro = new ErrorForm(ex);
                erro.ShowDialog();
                bdConn.Close();
            }
        }

        //BOTÃO OK SELECIONA PRJ
        private void bt_OK_SPrj_CadastroRQF_Click(object sender, EventArgs e)
        {
            if (dataGrid_SPrj_CadastroRQF.CurrentRow == null)
                botaoAlert("Antes de prosseguir selecionar um projeto.");
            else
            {
                if (bdDataSet.Tables["prj_objeto"].Rows.Count > 0)
                {
                    try
                    {
                        if (cb_df_rqf.Items.Count == 0)
                            cb_df_rqf.Items.Add("");
                        panel_SPrj_CadastroRQF.Visible = false;
                        panel_CadastroRQF.Visible = true;
                        bt_Salvar_RQF.Visible = true;
                        bt_Limpar_RQF.Visible = true;
                        tb_Prj_CadastroRQF.Text = dataGrid_SPrj_CadastroRQF.CurrentRow.Cells[0].Value.ToString();
                    }
                    catch (Exception ex)
                    {
                        this.Opacity = 0.9;
                        ErrorForm erro = new ErrorForm(ex);
                        erro.ShowDialog();
                        bdConn.Close();
                        this.Close();
                    }
                }
                else
                    botaoAlert("Nenhum projeto foi encontrado! Pesquise novamente.");
            }

        }

        //SELECIONA PRJ - DUPLO CLICK GRIDVIEW
        private void dataGrid_SPrj_CadastroRQF_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            bt_OK_SPrj_CadastroRQF.PerformClick();
        }

        //CADASTRO RQF INICIO
        public void cadastroRQFInicio()
        {
            bt_Limpar_RQF.PerformClick();
            tb_Prj_CadastroRQF.Text = "";
            tb_SPrj_CadastroRQF.Text = "";
            panel_SPrj_CadastroRQF.Visible = true;
            panel_CadastroRQF.Visible = false;
            bt_Salvar_RQF.Visible = false;
            bt_Limpar_RQF.Visible = false;
        }

        //BOTÃO LIMPAR RQF
        private void bt_Limpar_RQF_Click(object sender, EventArgs e)
        {
            tb_RQF_CadastroRQF.Text = "";
            rtb_DescRQF_CadastroRQF.Text = "";
        }

        //BOTÃO RETURN - ESCOLHER OUTRO PRJ
        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult resultReturnCPgm = MessageBox.Show("Deseja voltar e escolher outro Projeto?\n\nTodas alterações serão perdidas!", "Voltar", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (resultReturnCPgm == DialogResult.Yes)
                cadastroRQFInicio();
        }

        //CONTROLE DO TEXTBOX DESCRIÇÃO DF       
        private void rtb_DescRQF_CadastroRQF_TextChanged(object sender, EventArgs e)
        {
            int totalValue_DF = 300 - rtb_DescRQF_CadastroRQF.TextLength;
            lb_Controle_DescDF_CadastroRQF.Text = totalValue_DF.ToString();
        }

        //KEY PRESS- ENTER - CADASTRO RQF
        private void tb_SPrj_CadastroRQF_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Keys)e.KeyChar == Keys.Enter)
                bt_OK_SPrj_CadastroRQF.PerformClick();
        }

        #endregion

        #region //**************************************** CADASTRAR ANALISTA ****************************************\\

        //BOTÃO SALVAR - ANALISTA
        private void bt_Salvar_Analista_Click(object sender, EventArgs e)
        {
            if (validaCadsAnalista())
            {
                try
                {
                    lb_Analista.ForeColor = Color.Black;
                    bdConn.Open();
                    int existeAnalista = 0;

                    MySqlCommand commandAnalista = new MySqlCommand("select count(nome) from analistas where nome = '" + tb_NomeAnalista.Text + "';", bdConn);
                    MySqlDataReader drA = commandAnalista.ExecuteReader();

                    if (drA.Read())
                        existeAnalista = Int16.Parse(drA["count(nome)"].ToString());
                    drA.Close();
                    bdConn.Close();

                    if (existeAnalista > 0)
                        botaoAlert("Este nome já foi cadastrado!");
                    else
                    {
                        int grupoAnalista = 01;
                        bdConn.Open();

                        MySqlCommand recCodGrupo = new MySqlCommand("SELECT cod_grupo FROM grupo_analistas WHERE nom_grupo = '" + cb_GrupoAnalista.Text + "'", bdConn);
                        MySqlDataReader drGrupo = recCodGrupo.ExecuteReader();

                        if (drGrupo.Read())
                            grupoAnalista = Int16.Parse(drGrupo["cod_grupo"].ToString());
                        drGrupo.Close();

                        commandAnalista = new MySqlCommand("INSERT INTO analistas (nome, email, cod_grupo) VALUES ('" + tb_NomeAnalista.Text + "', '" + tb_EmailAnalista.Text + "', " + grupoAnalista + ")", bdConn);
                        commandAnalista.ExecuteNonQuery();
                        botaoConcluido("Tudo certo. O analista foi cadastrado com sucesso.");

                        bdConn.Close();

                        atualizaCBAnalistas();

                        bt_Limpar_Analista.PerformClick();
                    }
                }
                catch (Exception ex)
                {
                    this.Opacity = 0.9;
                    ErrorForm erro = new ErrorForm(ex);
                    erro.ShowDialog();
                    bdConn.Close();
                    this.Close();
                }
            }
            else
                botaoAlert("O preenchimento dos campos são obrigatórios!");

        }

        //VALIDA CADASTRO DE ANALISTA
        public bool validaCadsAnalista()
        {
            if (String.IsNullOrEmpty(tb_NomeAnalista.Text))
            {
                lb_Analista.ForeColor = Color.Red;
                return false;
            }
            else
                lb_Analista.ForeColor = Color.Black;

            if (String.IsNullOrEmpty(tb_EmailAnalista.Text))
            {
                lb_EmailAnalista.ForeColor = Color.Red;
                return false;
            }
            else
                lb_EmailAnalista.ForeColor = Color.Black;

            if (String.IsNullOrEmpty(cb_GrupoAnalista.Text))
            {
                lb_Grupo.ForeColor = Color.Red;
                return false;
            }
            else
                lb_Grupo.ForeColor = Color.Black;

            return true;
        }

        //BOTÃO LIMPAR - ANALISTA
        private void bt_Limpar_Analista_Click(object sender, EventArgs e)
        {
            lb_Analista.ForeColor = Color.Black;
            lb_EmailAnalista.ForeColor = Color.Black;
            lb_Grupo.ForeColor = Color.Black;

            tb_NomeAnalista.Text = "";
            tb_EmailAnalista.Text = "";
            cb_GrupoAnalista.Text = null;
            tb_LiderGrupo.Text = "";
            cb_AnalistaEdit.Text = null;

            atualizaCBAnalistas();
        }

        //KEY PRESS- ENTER - ANALISTA
        private void tb_NomeAnalista_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Keys)e.KeyChar == Keys.Enter)
                bt_Salvar_Analista.PerformClick();
        }

        //SELECIONA NOME DO LIDER DO GRUPO
        private void cb_GrupoAnalista_SelectedIndexChanged(object sender, EventArgs e)
        {
            bdConn.Open();

            MySqlCommand cmd = new MySqlCommand("SELECT lider_grupo FROM grupo_analistas WHERE nom_grupo = '" + cb_GrupoAnalista.Text + "'", bdConn);
            MySqlDataReader dr = cmd.ExecuteReader();

            if (dr.Read())
                tb_LiderGrupo.Text = dr["lider_grupo"].ToString();

            bdConn.Close();
        }

        //RECUPERA INFORMAÇÕES DE ANALISTA PARA ALTERAR/EXCLUIR
        private void cb_AnalistaEdit_SelectedIndexChanged(object sender, EventArgs e)
        {
            bdConn.Open();
            string codGrupo = "";
            string nomGrupo = "";

            MySqlCommand cmd = new MySqlCommand("SELECT email,cod_grupo FROM analistas WHERE nome = '" + cb_AnalistaEdit.Text + "'", bdConn);
            MySqlDataReader dr = cmd.ExecuteReader();

            if (dr.Read())
            {
                tb_EmailAnalista.Text = dr["email"].ToString();
                codGrupo = dr["cod_grupo"].ToString();
            }
            dr.Close();

            cmd = new MySqlCommand("SELECT nom_grupo FROM grupo_analistas WHERE cod_grupo = '" + codGrupo + "'", bdConn);
            dr = cmd.ExecuteReader();

            if (dr.Read())
                nomGrupo = dr["nom_grupo"].ToString();
            dr.Close();

            bdConn.Close();

            cb_GrupoAnalista.Text = nomGrupo;
        }

        //BOTÃO EXCLUIR ANALISTA
        private void bt_ExcluirAnalista_Click(object sender, EventArgs e)
        {

            if (cb_AnalistaEdit.Text != "")
            {
                DialogResult result = MessageBox.Show("Deseja excluir " + cb_AnalistaEdit.Text + " da lista de Analistas?", "Excluir", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    try
                    {
                        bdConn.Open();
                        MySqlCommand cmd = new MySqlCommand("DELETE FROM analistas WHERE nome = '" + cb_AnalistaEdit.Text + "'", bdConn);
                        cmd.ExecuteNonQuery();
                        bdConn.Close();
                        botaoConcluido(cb_AnalistaEdit.Text + " foi removido da lista com sucesso.");
                        atualizaCBAnalistas();
                        bt_Limpar_Analista.PerformClick();
                    }
                    catch (Exception ex)
                    {
                        this.Opacity = 0.9;
                        ErrorForm erro = new ErrorForm(ex);
                        erro.ShowDialog();
                        bdConn.Close();
                        this.Close();
                    }
                }
            }
            else
                botaoAlert("Selecione um analista antes de prosseguir.");
        }

        //BOTÃO ATUALIZAR ANALISTA
        private void bt_AtualizarAnalistas_Click(object sender, EventArgs e)
        {
            try
            {
                if (cb_AnalistaEdit.Text != "")
                {
                    int grupoAnalista = 01;

                    bdConn.Open();

                    MySqlCommand recCodGrupo = new MySqlCommand("SELECT cod_grupo FROM grupo_analistas WHERE nom_grupo = '" + cb_GrupoAnalista.Text + "'", bdConn);
                    MySqlDataReader drGrupo = recCodGrupo.ExecuteReader();

                    if (drGrupo.Read())
                        grupoAnalista = Int16.Parse(drGrupo["cod_grupo"].ToString());
                    drGrupo.Close();

                    MySqlCommand cmd = new MySqlCommand("UPDATE analistas SET email = '" + tb_EmailAnalista.Text + "', cod_grupo = " + grupoAnalista + " WHERE nome = '" + cb_AnalistaEdit.Text + "'", bdConn);
                    cmd.ExecuteNonQuery();

                    bdConn.Close();

                    botaoConcluido(cb_AnalistaEdit.Text + " foi atualizado com sucesso.");

                    bt_Limpar_Analista.PerformClick();
                }
                else
                    botaoAlert("Selecione um analista antes de prosseguir.");


            }
            catch (Exception ex)
            {
                this.Opacity = 0.9;
                ErrorForm erro = new ErrorForm(ex);
                erro.ShowDialog();
                bdConn.Close();
                this.Close();
            }
        }

        //ATUALIZA CB DE ANALISTAS
        public void atualizaCBAnalistas()
        {
            cb_AnalistaEdit.Items.Clear();
            bdConn.Open();
            MySqlCommand commandA = new MySqlCommand("select nome from analistas order by nome;", bdConn);
            MySqlDataReader drA = commandA.ExecuteReader();
            while (drA.Read())
                cb_AnalistaEdit.Items.Add(drA["nome"].ToString());
            drA.Close();
            bdConn.Close();
        }

        #endregion

        #region //**************************************** CADASTRAR FASE ****************************************\\

        //BOTÃO SALVAR - FASE
        private void bt_Salvar_Fase_Click(object sender, EventArgs e)
        {
            if (tb_NomeFase.Text != "")
            {
                try
                {
                    MySqlCommand commandFase;

                    lb_NomeFase.ForeColor = Color.Black;
                    bdConn.Open();

                    commandFase = new MySqlCommand("select count(nom_fase) from fase where nom_fase = '" + tb_NomeFase.Text + "';", bdConn);
                    MySqlDataReader drF = commandFase.ExecuteReader();
                    drF.Read();

                    if (Int16.Parse(drF["count(nom_fase)"].ToString()) > 0)
                    {
                        drF.Close();
                        botaoAlert("Esta fase já existe!");
                    }
                    else
                    {
                        drF.Close();
                        commandFase = new MySqlCommand("INSERT INTO fase (nom_fase) VALUES ('" + tb_NomeFase.Text + "')", bdConn);
                        commandFase.ExecuteNonQuery();
                        botaoConcluido("Tudo certo. O registro foi cadastrado com sucesso.");
                    }

                    bdConn.Close();
                }
                catch (Exception ex)
                {
                    ErrorForm erro = new ErrorForm(ex);
                    erro.ShowDialog();
                    bdConn.Close();
                }
            }
            else
            {
                lb_NomeFase.ForeColor = Color.Red;
                botaoAlert("Digitar o nome da Fase!");
            }
        }

        //BOTÃO LIMPAR - FASE
        private void bt_Limpar_Fase_Click(object sender, EventArgs e)
        {
            lb_NomeFase.ForeColor = Color.Black;
            tb_NomeFase.Text = "";
        }

        //KEY PRESS - ENTER - FASE
        private void tb_NomeFase_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Keys)e.KeyChar == Keys.Enter)
                bt_Salvar_Fase.PerformClick();
        }

        #endregion

        #region //**************************************** ACOMPANHAMENTO DE PROGRAMA ****************************************\\

        //VARIAVEIS UTILIZADAS
        string querySPGM_AcompanhamentoPGM;
        string pgm_select;
        string prj_select;
        string rqf_select;
        static string status;
        int codPAD = 0;
        bool travaDTI = false;
        bool travaDTF = false;
        int valorAgregado = 0;
        #region oldValue
        int oldValue_DT1 = 0;
        int oldValue_DT2 = 0;
        int oldValue_DT3 = 0;
        int oldValue_DT4 = 0;
        int oldValue_DT5 = 0;
        int oldValue_DT6 = 0;
        int oldValue_DT7 = 0;
        int oldValue_DT8 = 0;
        int oldValue_DT9 = 0;
        int oldValue_DT10 = 0;
        int oldValue_DT11 = 0;
        int oldValue_DT12 = 0;
        int oldValue_DT13 = 0;
        int oldValue_DT14 = 0;
        #endregion

        //BOTÃO SALVAR - ACOMPANHAMENTO DE PROGRAMA
        private void bt_Salvar_Acompanhamento_Click(object sender, EventArgs e)
        {
            try
            {
                //if (valorAgregado == 100)
                //{
                bdConn.Open();
                atualizarAcompanhamentoPGM();
                bdConn.Close();
                botaoConcluido("Tudo certo. O programa foi atualizado com sucesso.");
                carregaAcompanhamentoPGM();
                /*}
                else
                {
                    botaoAlert("Só pode finalizar a contrução quando estiver 100% construído. Ainda restam " + (100 - valorAgregado) + "% para finalizar.");
                    carregaAcompanhamentoPGM();
                }*/


            }
            catch (Exception ex)
            {
                this.Opacity = 0.9;
                ErrorForm erro = new ErrorForm(ex);
                erro.ShowDialog();
                bdConn.Close();
                this.Close();
            }
        }

        //ATUALIZA ACOMPANHAMENTO
        private void atualizarAcompanhamentoPGM()
        {
            string queryUpdateAcompPGM;

            queryUpdateAcompPGM = "UPDATE pgm_acompanhamento SET ";

            //STATUS CONSTRUÇÃO
            queryUpdateAcompPGM += "status_construcao = '" + cb_status_acompanhamento.Text + "'";

            //DATA DE INICIO
            if (dti_acompanhamento.Enabled == true && dti_acompanhamento.Visible == true)
                queryUpdateAcompPGM += ", data_inicio = '" + dti_acompanhamento.Value.ToString("yyyy/MM/dd") + "'";

            //DATA FINAL
            if (dtf_acompanhamento.Enabled == true && dtf_acompanhamento.Visible == true)
                queryUpdateAcompPGM += ", data_fim = '" + dtf_acompanhamento.Value.ToString("yyyy/MM/dd") + "'";

            //ANOTAÇÕES GERAIS
            if (tb_AGerais_acompanhamento.Text != "")
                queryUpdateAcompPGM += ", anot_gerais = '" + tb_AGerais_acompanhamento.Text + "'";

            //LIBERA CODE REVIEW
            if (cb_status_acompanhamento.Text == "Finalizado")
                queryUpdateAcompPGM += ", lib_cr = 'Sim'";

            //LIBERA PERFORMACE REVIEW
            if (cb_status_CR.Text == "Finalizado")
                queryUpdateAcompPGM += ", lib_pr = 'Sim'";

            //CODE REVIEW
            if (panel_CR.Enabled == true)
                if (cb_status_CR.Text != "Não Iniciado" && cb_status_CR.Text != "")
                {
                    queryUpdateAcompPGM += ", status_cr = '" + cb_status_CR.Text + "'";

                    if (cb_analistas_CR.Enabled)
                        queryUpdateAcompPGM += ", resp_cr = '" + cb_analistas_CR.Text + "'";

                    //GRAVA DATA SE CAMPO ESTIVER VAZIO
                    if (data_CR.Enabled)
                        queryUpdateAcompPGM += ", data_cr = '" + data_CR.Value.ToString("yyyy/MM/dd") + "'";
                }

            //PERFORMACE REVIEW
            if (panel_PR.Enabled == true)
                if (cb_status_PR.Text != "Não Iniciado" && cb_status_PR.Text != "")
                {
                    queryUpdateAcompPGM += ", status_pr = '" + cb_status_PR.Text + "'";

                    if (cb_analistas_PR.Enabled)
                        queryUpdateAcompPGM += ", resp_pr = '" + cb_analistas_PR.Text + "'";

                    //GRAVA DATA SE CAMPO ESTIVER VAZIO
                    if (data_PR.Enabled)
                        queryUpdateAcompPGM += ", data_pr = '" + data_PR.Value.ToString("yyyy/MM/dd") + "'";
                }

            queryUpdateAcompPGM += " WHERE cod_pgm = '" + tb_programa_acompanhamento.Text + "' AND cod_prj = '" + tb_projeto_acompanhamento.Text + "' AND df_rqf = '" + tb_rqf_acompanhamento.Text + "';";

            //EXECUTA COMANDO
            MySqlCommand commandAcompanhamentoPGM = new MySqlCommand(queryUpdateAcompPGM, bdConn);
            commandAcompanhamentoPGM.ExecuteNonQuery();

            updateDatasAcompanhamento();

        }

        //ATUALIZA PERCENTUAL DE CONSTRUÇÃO
        void updateDatasAcompanhamento()
        {
            List<string> queryUpdateData = new List<string>();
            MySqlCommand command;

            if ((tb_DA_dt1.Visible == true) && (tb_DA_dt1.Enabled == true) && (tb_DA_dt1.Value > 0))
                queryUpdateData.Add("UPDATE pgm_acomp_data SET pct_pgm = " + tb_DA_dt1.Value + " WHERE cod_pad = " + codPAD + " AND data_acomp = '" + DateTime.Parse(lb_DA_dt1.Text).ToString("yyyy/MM/dd") + "';");

            if ((tb_DA_dt2.Visible == true) && (tb_DA_dt2.Enabled == true) && (tb_DA_dt2.Value > 0))
                queryUpdateData.Add("UPDATE pgm_acomp_data SET pct_pgm = " + tb_DA_dt2.Value + " WHERE cod_pad = " + codPAD + " AND data_acomp = '" + DateTime.Parse(lb_DA_dt2.Text).ToString("yyyy/MM/dd") + "';");

            if ((tb_DA_dt3.Visible == true) && (tb_DA_dt3.Enabled == true) && (tb_DA_dt3.Value > 0))
                queryUpdateData.Add("UPDATE pgm_acomp_data SET pct_pgm = " + tb_DA_dt3.Value + " WHERE cod_pad = " + codPAD + " AND data_acomp = '" + DateTime.Parse(lb_DA_dt3.Text).ToString("yyyy/MM/dd") + "';");

            if ((tb_DA_dt4.Visible == true) && (tb_DA_dt4.Enabled == true) && (tb_DA_dt4.Value > 0))
                queryUpdateData.Add("UPDATE pgm_acomp_data SET pct_pgm = " + tb_DA_dt4.Value + " WHERE cod_pad = " + codPAD + " AND data_acomp = '" + DateTime.Parse(lb_DA_dt4.Text).ToString("yyyy/MM/dd") + "';");

            if ((tb_DA_dt5.Visible == true) && (tb_DA_dt5.Enabled == true) && (tb_DA_dt5.Value > 0))
                queryUpdateData.Add("UPDATE pgm_acomp_data SET pct_pgm = " + tb_DA_dt5.Value + " WHERE cod_pad = " + codPAD + " AND data_acomp = '" + DateTime.Parse(lb_DA_dt5.Text).ToString("yyyy/MM/dd") + "';");

            if ((tb_DA_dt6.Visible == true) && (tb_DA_dt6.Enabled == true) && (tb_DA_dt6.Value > 0))
                queryUpdateData.Add("UPDATE pgm_acomp_data SET pct_pgm = " + tb_DA_dt6.Value + " WHERE cod_pad = " + codPAD + " AND data_acomp = '" + DateTime.Parse(lb_DA_dt6.Text).ToString("yyyy/MM/dd") + "';");

            if ((tb_DA_dt7.Visible == true) && (tb_DA_dt7.Enabled == true) && (tb_DA_dt7.Value > 0))
                queryUpdateData.Add("UPDATE pgm_acomp_data SET pct_pgm = " + tb_DA_dt7.Value + " WHERE cod_pad = " + codPAD + " AND data_acomp = '" + DateTime.Parse(lb_DA_dt7.Text).ToString("yyyy/MM/dd") + "';");

            if ((tb_DA_dt8.Visible == true) && (tb_DA_dt8.Enabled == true) && (tb_DA_dt8.Value > 0))
                queryUpdateData.Add("UPDATE pgm_acomp_data SET pct_pgm = " + tb_DA_dt8.Value + " WHERE cod_pad = " + codPAD + " AND data_acomp = '" + DateTime.Parse(lb_DA_dt8.Text).ToString("yyyy/MM/dd") + "';");

            if ((tb_DA_dt9.Visible == true) && (tb_DA_dt9.Enabled == true) && (tb_DA_dt9.Value > 0))
                queryUpdateData.Add("UPDATE pgm_acomp_data SET pct_pgm = " + tb_DA_dt9.Value + " WHERE cod_pad = " + codPAD + " AND data_acomp = '" + DateTime.Parse(lb_DA_dt9.Text).ToString("yyyy/MM/dd") + "';");

            if ((tb_DA_dt10.Visible == true) && (tb_DA_dt10.Enabled == true) && (tb_DA_dt10.Value > 0))
                queryUpdateData.Add("UPDATE pgm_acomp_data SET pct_pgm = " + tb_DA_dt10.Value + " WHERE cod_pad = " + codPAD + " AND data_acomp = '" + DateTime.Parse(lb_DA_dt10.Text).ToString("yyyy/MM/dd") + "';");

            if ((tb_DA_dt11.Visible == true) && (tb_DA_dt11.Enabled == true) && (tb_DA_dt11.Value > 0))
                queryUpdateData.Add("UPDATE pgm_acomp_data SET pct_pgm = " + tb_DA_dt11.Value + " WHERE cod_pad = " + codPAD + " AND data_acomp = '" + DateTime.Parse(lb_DA_dt11.Text).ToString("yyyy/MM/dd") + "';");

            if ((tb_DA_dt12.Visible == true) && (tb_DA_dt12.Enabled == true) && (tb_DA_dt12.Value > 0))
                queryUpdateData.Add("UPDATE pgm_acomp_data SET pct_pgm = " + tb_DA_dt12.Value + " WHERE cod_pad = " + codPAD + " AND data_acomp = '" + DateTime.Parse(lb_DA_dt12.Text).ToString("yyyy/MM/dd") + "';");

            if ((tb_DA_dt13.Visible == true) && (tb_DA_dt13.Enabled == true) && (tb_DA_dt13.Value > 0))
                queryUpdateData.Add("UPDATE pgm_acomp_data SET pct_pgm = " + tb_DA_dt13.Value + " WHERE cod_pad = " + codPAD + " AND data_acomp = '" + DateTime.Parse(lb_DA_dt13.Text).ToString("yyyy/MM/dd") + "';");

            if ((tb_DA_dt14.Visible == true) && (tb_DA_dt14.Enabled == true) && (tb_DA_dt14.Value > 0))
                queryUpdateData.Add("UPDATE pgm_acomp_data SET pct_pgm = " + tb_DA_dt14.Value + " WHERE cod_pad = " + codPAD + " AND data_acomp = '" + DateTime.Parse(lb_DA_dt14.Text).ToString("yyyy/MM/dd") + "';");

            if ((tb_DA_dt15.Visible == true) && (tb_DA_dt15.Enabled == true) && (tb_DA_dt15.Value > 0))
                queryUpdateData.Add("UPDATE pgm_acomp_data SET pct_pgm = " + tb_DA_dt15.Value + " WHERE cod_pad = " + codPAD + " AND data_acomp = '" + DateTime.Parse(lb_DA_dt15.Text).ToString("yyyy/MM/dd") + "';");

            foreach (var atualizaData in queryUpdateData)
            {
                command = new MySqlCommand(atualizaData, bdConn);
                command.ExecuteNonQuery();
            }
        }

        //SELECIONA PROGRAMA - PESQUISA AUTOMATICA
        private void tb_selecionaPrograma_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (tb_selecionaPrograma.Text != "")
                {
                    querySPGM_AcompanhamentoPGM = "select cod_pgm,cod_prj,df_rqf from pgm_acompanhamento where cod_pgm like '%" + tb_selecionaPrograma.Text + "%'";
                    bdDataSet = new DataSet();
                    bdConn.Open();
                    bdAdapter = new MySqlDataAdapter(querySPGM_AcompanhamentoPGM, bdConn);
                    bdAdapter.Fill(bdDataSet, "pgm_acompanhamento");
                    dataGrid_selecionaPrograma.DataSource = bdDataSet;

                    if (bdDataSet.Tables["pgm_acompanhamento"].Rows.Count == 0)
                        lb_NotFound_selecionaPrograma.Visible = true;
                    else
                        lb_NotFound_selecionaPrograma.Visible = false;

                    dataGrid_selecionaPrograma.DataMember = "pgm_acompanhamento";
                    bdConn.Close();
                }
                else
                {
                    lb_NotFound_selecionaPrograma.Visible = false;

                    if (this.dataGrid_selecionaPrograma.DataSource != null)
                        this.dataGrid_selecionaPrograma.DataSource = null;
                    else
                    {
                        this.dataGrid_selecionaPrograma.Rows.Clear();
                        this.dataGrid_selecionaPrograma.Columns.Clear();
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorForm erro = new ErrorForm(ex);
                erro.ShowDialog();
                bdConn.Close();
            }
        }

        //BOTÃO OK SELECIONA PRJ
        private void bt_Ok_selecionaPrograma_Click(object sender, EventArgs e)
        {
            if (dataGrid_selecionaPrograma.CurrentRow == null)
                botaoAlert("Selecionar um programa da lista antes de prosseguir.");
            else
            {
                if (bdDataSet.Tables["pgm_acompanhamento"].Rows.Count > 0)
                {
                    panel_selecionaPrograma.Visible = false;
                    panel_acompanhamentoPGM.Visible = true;
                    bt_Salvar_Acompanhamento.Visible = true;
                    bt_Limpar_Acompanhamento.Visible = true;

                    pgm_select = dataGrid_selecionaPrograma.CurrentRow.Cells[0].Value.ToString();
                    prj_select = dataGrid_selecionaPrograma.CurrentRow.Cells[1].Value.ToString();
                    rqf_select = dataGrid_selecionaPrograma.CurrentRow.Cells[2].Value.ToString();

                    carregaAcompanhamentoPGM();
                }
                else
                    botaoAlert("Nenhum programa foi encontrado! Pesquise novamente.");
            }
        }

        //CARREGA CAMPOS DO ACOMPANHAMENTO
        private void carregaAcompanhamentoPGM()
        {
            try
            {
                bdConn.Open();
                string camposBD = "P.cod_pgm, P.cod_prj, P.df_rqf, P.sistema, P.resp_cttu, PA.*";
                MySqlCommand commandA_PGM = new MySqlCommand("SELECT " + camposBD + " FROM pgm_objeto P natural join pgm_acompanhamento PA WHERE cod_pgm = '" + pgm_select + "' AND cod_prj = '" + prj_select + "' AND df_rqf = '" + rqf_select + "';", bdConn);
                MySqlDataReader drA_PGM = commandA_PGM.ExecuteReader();

                if (drA_PGM.Read())
                {
                    //********* INFORMAÇÕES *********\\
                    tb_programa_acompanhamento.Text = drA_PGM["cod_pgm"].ToString();
                    tb_projeto_acompanhamento.Text = drA_PGM["cod_prj"].ToString();
                    tb_rqf_acompanhamento.Text = drA_PGM["df_rqf"].ToString();
                    tb_sistema_acompanhamento.Text = drA_PGM["sistema"].ToString();
                    tb_respCTTU_acompanhamento.Text = drA_PGM["resp_cttu"].ToString();

                    codPAD = Int16.Parse(drA_PGM["cod_pad"].ToString());

                    //********* CONSTRUÇÃO *********\\
                    panel_AP_Construcao.Enabled = true;
                    cb_status_acompanhamento.Text = drA_PGM["status_construcao"].ToString();
                    status = drA_PGM["status_construcao"].ToString();

                    //STATUS CONSTRUÇÃO
                    #region STATUS CONSTRUÇÃO
                    switch (drA_PGM["status_construcao"].ToString())
                    {
                        case "Iniciado":
                            cb_status_acompanhamento.Items.Clear();
                            cb_status_acompanhamento.Items.Add("Paralisado");
                            cb_status_acompanhamento.Items.Add("Finalizado");
                            cb_status_acompanhamento.Enabled = true;
                            gb_PercentConst.Visible = true;
                            break;
                        case "Paralisado":
                            cb_status_acompanhamento.Items.Clear();
                            cb_status_acompanhamento.Items.Add("Paralisado");
                            cb_status_acompanhamento.Items.Add("Finalizado");
                            cb_status_acompanhamento.Enabled = true;
                            gb_PercentConst.Visible = true;
                            break;
                        case "Finalizado":
                            gb_PercentConst.Visible = true;
                            panel_AP_Construcao.Enabled = false;
                            break;
                        default:
                            cb_status_acompanhamento.Items.Clear();
                            cb_status_acompanhamento.Items.Add("Não Iniciado");
                            cb_status_acompanhamento.Items.Add("Iniciado");
                            cb_status_acompanhamento.Enabled = true;
                            gb_PercentConst.Visible = false;
                            break;
                    }
                    #endregion

                    //DATA DE INICIO DA CONSTRUÇÃO
                    #region DATAS DA CONTRUÇÃO
                    if (drA_PGM["data_inicio"].ToString() != "00/00/0000")
                    {
                        dti_acompanhamento.Text = drA_PGM["data_inicio"].ToString();
                        dti_acompanhamento.Enabled = false;

                        travaDTI = true;

                        dti_acompanhamento.Visible = true;
                        lb_dti_acompanhamento.Visible = true;
                    }
                    else
                    {
                        travaDTI = false;
                        dti_acompanhamento.Visible = false;
                        lb_dti_acompanhamento.Visible = false;
                        dti_acompanhamento.Enabled = true;
                    }

                    //DATA FIM DA CONSTRUÇÃO
                    if (drA_PGM["data_fim"].ToString() != "00/00/0000")
                    {
                        dtf_acompanhamento.Text = drA_PGM["data_fim"].ToString();
                        dtf_acompanhamento.Enabled = false;

                        travaDTF = true;

                        dtf_acompanhamento.Visible = true;
                        lb_dtf_acompanhamento.Visible = true;
                    }
                    else
                    {
                        travaDTF = false;
                        dtf_acompanhamento.Visible = false;
                        lb_dtf_acompanhamento.Visible = false;
                        dtf_acompanhamento.Enabled = true;
                    }
                    #endregion

                    //ANOTAÇÕES GERAIS
                    tb_AGerais_acompanhamento.Text = drA_PGM["anot_gerais"].ToString();

                    //********* CODE REVIEW *********\\                    
                    #region CARREGA CODE REVIEW
                    if (drA_PGM["lib_cr"].ToString() == "Sim")
                    {
                        //LABEL DE INFO
                        lb_liberaCR.Text = "Liberado";
                        lb_liberaCR.ForeColor = Color.Green;

                        //HABILITA CAMPOS PARA VISUALIZAÇÃO
                        panel_CR.Visible = true;
                        panel_CR.Enabled = true;

                        //STATUS DO CODE REVIEW
                        cb_status_CR.Text = drA_PGM["status_cr"].ToString();
                        switch (drA_PGM["status_cr"].ToString())
                        {
                            case "Não Iniciado":
                                cb_status_CR.Items.Clear();
                                cb_status_CR.Items.Add("Não Iniciado");
                                cb_status_CR.Items.Add("Iniciado");
                                cb_status_CR.Enabled = true;
                                break;
                            case "Iniciado":
                                cb_status_CR.Items.Clear();
                                cb_status_CR.Items.Add("Finalizado");
                                cb_status_CR.Enabled = true;
                                break;
                            case "Finalizado":
                                panel_CR.Enabled = false;
                                lb_liberaCR.Text = "Finalizado";
                                lb_liberaCR.ForeColor = Color.SteelBlue;
                                break;
                        }

                        //RESPONSÁVEL PELO CODE REVIEW
                        if ((cb_analistas_CR.Text = drA_PGM["resp_cr"].ToString()) == "")
                            cb_analistas_CR.Enabled = true;
                        else
                            cb_analistas_CR.Enabled = false;

                        //DATA DO CODE REVIEW
                        if (drA_PGM["data_cr"].ToString() == "00/00/0000")
                        {
                            data_CR.Value = DateTime.Now;
                            data_CR.Enabled = true;
                        }
                        else
                        {
                            data_CR.Value = DateTime.Parse(drA_PGM["data_cr"].ToString());
                            data_CR.Enabled = false;
                        }

                    }
                    else
                        resetaCodeReiview();

                    #endregion

                    //********* PERFOMACE REVIEW *********\\
                    #region CARREGA PERFOMACE REVIEW
                    if (drA_PGM["lib_pr"].ToString() == "Sim")
                    {
                        //LABEL E INFO 
                        lb_liberaPR.Text = "Liberado";
                        lb_liberaPR.ForeColor = Color.Green;

                        //HABILITA CAMPOS PARA VISUALIZAÇÃO
                        panel_PR.Visible = true;
                        panel_PR.Enabled = true;

                        //STATUS DO PERFORMACE REVIREW
                        cb_status_PR.Text = drA_PGM["status_pr"].ToString();
                        switch (drA_PGM["status_pr"].ToString())
                        {
                            case "Não Iniciado":
                                cb_status_PR.Items.Clear();
                                cb_status_PR.Items.Add("Não Iniciado");
                                cb_status_PR.Items.Add("Iniciado");
                                cb_status_PR.Enabled = true;
                                break;
                            case "Iniciado":
                                cb_status_PR.Items.Clear();
                                cb_status_PR.Items.Add("Finalizado");
                                cb_status_PR.Enabled = true;
                                break;
                            case "Finalizado":
                                panel_PR.Enabled = false;
                                lb_liberaPR.Text = "Finalizado";
                                lb_liberaPR.ForeColor = Color.SteelBlue;
                                break;
                        }

                        //RESPONSÁVEL PELO PERFORMACE REVIEW                         
                        if ((cb_analistas_PR.Text = drA_PGM["resp_pr"].ToString()) == "")
                            cb_analistas_PR.Enabled = true;
                        else
                            cb_analistas_PR.Enabled = false;

                        //DATA DO PERFORMACE REVIEW
                        if (drA_PGM["data_pr"].ToString() == "00/00/0000")
                        {
                            data_PR.Value = DateTime.Now;
                            data_PR.Enabled = true;
                        }
                        else
                        {
                            data_PR.Value = DateTime.Parse(drA_PGM["data_pr"].ToString());
                            data_PR.Enabled = false;
                        }

                    }
                    else
                        resetaPerformaceReiview();

                    #endregion
                }
                drA_PGM.Close();

                //DATAS DE ACOMPANHAMENTO
                #region CARREGA DATAS DE ACOMPANHAMENTO
                int posicaoData = 0;
                valorAgregado = 0;
                commandA_PGM = new MySqlCommand("SELECT data_acomp, pct_pgm FROM pgm_acomp_data WHERE cod_pad = " + codPAD + ";", bdConn);
                drA_PGM = commandA_PGM.ExecuteReader();
                while (drA_PGM.Read())
                {
                    posicaoData++;
                    carregaDataAcompanhamento(posicaoData, drA_PGM["data_acomp"].ToString(), drA_PGM["pct_pgm"].ToString());
                }
                drA_PGM.Close();

                tb_AP_P_Construcao.Text = valorAgregado.ToString() + " % Construído";

                #region oldValue
                oldValue_DT1 = 0;
                oldValue_DT2 = 0;
                oldValue_DT3 = 0;
                oldValue_DT4 = 0;
                oldValue_DT5 = 0;
                oldValue_DT6 = 0;
                oldValue_DT7 = 0;
                oldValue_DT8 = 0;
                oldValue_DT9 = 0;
                oldValue_DT10 = 0;
                oldValue_DT11 = 0;
                oldValue_DT12 = 0;
                oldValue_DT13 = 0;
                oldValue_DT14 = 0;
                #endregion

                #endregion

                bdConn.Close();
            }
            catch (Exception ex)
            {
                this.Opacity = 0.9;
                ErrorForm erro = new ErrorForm(ex);
                erro.ShowDialog();
                bdConn.Close();
                this.Close();
            }
        }

        //CARREGA DATAS DE ACOMPANHAMENTO
        void carregaDataAcompanhamento(int posicao, string data, string percentualContruido)
        {
            switch (posicao)
            {
                case 1:
                    lb_DA_dt1.Text = data;
                    lb_DA_dt1.Visible = true;
                    tb_DA_dt1.Visible = true;
                    lb_percent_1.Visible = true;

                    if (!String.IsNullOrEmpty(percentualContruido))
                    {
                        tb_DA_dt1.Value = Int16.Parse(percentualContruido);
                        tb_DA_dt1.Value = Int16.Parse(percentualContruido);
                        tb_DA_dt1.Enabled = false;
                        tb_DA_dt1.Enabled = false;
                    }
                    else
                    {
                        tb_DA_dt1.Value = 0;
                        tb_DA_dt1.Enabled = true;
                    }

                    valorAgregado += (int)tb_DA_dt1.Value;
                    break;

                case 2:
                    lb_DA_dt2.Text = data;
                    lb_DA_dt2.Visible = true;
                    tb_DA_dt2.Visible = true;
                    lb_percent_2.Visible = true;

                    if (percentualContruido != "")
                    {
                        tb_DA_dt2.Value = Int16.Parse(percentualContruido);
                        tb_DA_dt2.Enabled = false;
                        lb_percent_2.Enabled = false;
                    }
                    else
                    {
                        tb_DA_dt2.Value = 0;
                        tb_DA_dt2.Enabled = true;
                        tb_DA_dt2.Maximum = (100 - valorAgregado);
                    }

                    valorAgregado += (int)tb_DA_dt2.Value;

                    break;
                case 3:
                    lb_DA_dt3.Text = data;
                    lb_DA_dt3.Visible = true;
                    tb_DA_dt3.Visible = true;
                    lb_percent_3.Visible = true;

                    if (percentualContruido != "")
                    {
                        tb_DA_dt3.Value = Int16.Parse(percentualContruido);
                        tb_DA_dt3.Enabled = false;
                        lb_percent_3.Enabled = false;
                    }
                    else
                    {
                        tb_DA_dt3.Value = 0;
                        tb_DA_dt3.Enabled = true;
                        tb_DA_dt3.Maximum = (100 - valorAgregado);
                    }

                    valorAgregado += (int)tb_DA_dt3.Value;
                    break;
                case 4:
                    lb_DA_dt4.Text = data;
                    lb_DA_dt4.Visible = true;
                    tb_DA_dt4.Visible = true;
                    lb_percent_4.Visible = true;

                    if (percentualContruido != "")
                    {
                        tb_DA_dt4.Value = Int16.Parse(percentualContruido);
                        tb_DA_dt4.Enabled = false;
                        lb_percent_4.Enabled = false;
                    }
                    else
                    {
                        tb_DA_dt4.Value = 0;
                        tb_DA_dt4.Enabled = true;
                        tb_DA_dt4.Maximum = (100 - valorAgregado);
                    }

                    valorAgregado += (int)tb_DA_dt4.Value;
                    break;
                case 5:
                    lb_DA_dt5.Text = data;
                    lb_DA_dt5.Visible = true;
                    tb_DA_dt5.Visible = true;
                    lb_percent_5.Visible = true;

                    if (percentualContruido != "")
                    {
                        tb_DA_dt5.Value = Int16.Parse(percentualContruido);
                        tb_DA_dt5.Enabled = false;
                        lb_percent_5.Enabled = false;
                    }
                    else
                    {
                        tb_DA_dt5.Value = 0;
                        tb_DA_dt5.Enabled = true;
                        tb_DA_dt5.Maximum = (100 - valorAgregado);
                    }

                    valorAgregado += (int)tb_DA_dt5.Value;
                    break;
                case 6:
                    lb_DA_dt6.Text = data;
                    lb_DA_dt6.Visible = true;
                    tb_DA_dt6.Visible = true;
                    lb_percent_6.Visible = true;

                    if (percentualContruido != "")
                    {
                        tb_DA_dt6.Value = Int16.Parse(percentualContruido);
                        tb_DA_dt6.Enabled = false;
                        lb_percent_6.Enabled = false;
                    }
                    else
                    {
                        tb_DA_dt6.Value = 0;
                        tb_DA_dt6.Enabled = true;
                    }

                    valorAgregado += (int)tb_DA_dt6.Value;

                    break;
                case 7:
                    lb_DA_dt7.Text = data;
                    lb_DA_dt7.Visible = true;
                    tb_DA_dt7.Visible = true;
                    lb_percent_7.Visible = true;

                    if (percentualContruido != "")
                    {
                        tb_DA_dt7.Value = Int16.Parse(percentualContruido);
                        tb_DA_dt7.Enabled = false;
                        lb_percent_7.Enabled = false;
                    }
                    else
                    {
                        tb_DA_dt7.Value = 0;
                        tb_DA_dt7.Enabled = true;
                    }

                    valorAgregado += (int)tb_DA_dt7.Value;

                    break;
                case 8:
                    lb_DA_dt8.Text = data;
                    lb_DA_dt8.Visible = true;
                    tb_DA_dt8.Visible = true;
                    lb_percent_8.Visible = true;

                    if (percentualContruido != "")
                    {
                        tb_DA_dt8.Value = Int16.Parse(percentualContruido);
                        tb_DA_dt8.Enabled = false;
                        lb_percent_8.Enabled = false;
                    }
                    else
                    {
                        tb_DA_dt8.Value = 0;
                        tb_DA_dt8.Enabled = true;
                    }

                    valorAgregado += (int)tb_DA_dt8.Value;
                    break;
                case 9:
                    lb_DA_dt9.Text = data;
                    lb_DA_dt9.Visible = true;
                    tb_DA_dt9.Visible = true;
                    lb_percent_9.Visible = true;

                    if (percentualContruido != "")
                    {
                        tb_DA_dt9.Value = Int16.Parse(percentualContruido);
                        tb_DA_dt9.Enabled = false;
                        lb_percent_9.Enabled = false;
                    }
                    else
                    {
                        tb_DA_dt9.Value = 0;
                        tb_DA_dt9.Enabled = true;
                    }

                    valorAgregado += (int)tb_DA_dt9.Value;
                    break;
                case 10:
                    lb_DA_dt10.Text = data;
                    lb_DA_dt10.Visible = true;
                    tb_DA_dt10.Visible = true;
                    lb_percent_10.Visible = true;

                    if (percentualContruido != "")
                    {
                        tb_DA_dt10.Value = Int16.Parse(percentualContruido);
                        tb_DA_dt10.Enabled = false;
                        lb_percent_10.Enabled = false;
                    }
                    else
                    {
                        tb_DA_dt10.Value = 0;
                        tb_DA_dt10.Enabled = true;
                    }

                    valorAgregado += (int)tb_DA_dt10.Value;
                    break;
                case 11:
                    lb_DA_dt11.Text = data;
                    lb_DA_dt11.Visible = true;
                    tb_DA_dt11.Visible = true;
                    lb_percent_11.Visible = true;

                    if (percentualContruido != "")
                    {
                        tb_DA_dt11.Value = Int16.Parse(percentualContruido);
                        tb_DA_dt11.Enabled = false;
                        lb_percent_11.Enabled = false;
                    }
                    else
                    {
                        tb_DA_dt11.Value = 0;
                        tb_DA_dt11.Enabled = true;
                    }

                    valorAgregado += (int)tb_DA_dt11.Value;
                    break;
                case 12:
                    lb_DA_dt12.Text = data;
                    lb_DA_dt12.Visible = true;
                    tb_DA_dt12.Visible = true;
                    lb_percent_12.Visible = true;

                    if (percentualContruido != "")
                    {
                        tb_DA_dt12.Value = Int16.Parse(percentualContruido);
                        tb_DA_dt12.Enabled = false;
                        lb_percent_12.Enabled = false;
                    }
                    else
                    {
                        tb_DA_dt12.Value = 0;
                        tb_DA_dt12.Enabled = true;
                    }

                    valorAgregado += (int)tb_DA_dt12.Value;
                    break;
                case 13:
                    lb_DA_dt13.Text = data;
                    lb_DA_dt13.Visible = true;
                    tb_DA_dt13.Visible = true;
                    lb_percent_13.Visible = true;

                    if (percentualContruido != "")
                    {
                        tb_DA_dt13.Value = Int16.Parse(percentualContruido);
                        tb_DA_dt13.Enabled = false;
                        lb_percent_13.Enabled = false;
                    }
                    else
                    {
                        tb_DA_dt13.Value = 0;
                        tb_DA_dt13.Enabled = true;
                    }

                    valorAgregado += (int)tb_DA_dt13.Value;
                    break;
                case 14:
                    lb_DA_dt14.Text = data;
                    lb_DA_dt14.Visible = true;
                    tb_DA_dt14.Visible = true;
                    lb_percent_14.Visible = true;

                    if (percentualContruido != "")
                    {
                        tb_DA_dt14.Value = Int16.Parse(percentualContruido);
                        tb_DA_dt14.Enabled = false;
                        lb_percent_14.Enabled = false;
                    }
                    else
                    {
                        tb_DA_dt14.Value = 0;
                        tb_DA_dt14.Enabled = true;
                    }

                    valorAgregado += (int)tb_DA_dt14.Value;
                    break;
                case 15:
                    lb_DA_dt15.Text = data;
                    lb_DA_dt15.Visible = true;
                    tb_DA_dt15.Visible = true;
                    lb_percent_15.Visible = true;

                    if (percentualContruido != "")
                    {
                        tb_DA_dt15.Value = Int16.Parse(percentualContruido);
                        tb_DA_dt15.Enabled = false;
                        lb_percent_15.Enabled = false;
                    }
                    else
                    {
                        tb_DA_dt15.Value = 0;
                        tb_DA_dt15.Enabled = true;
                    }

                    valorAgregado += (int)tb_DA_dt15.Value;
                    break;
            }
        }

        //DUPLO CLICK GRIDVIEW - SELECIONA PROGRANA
        private void dataGrid_selecionaPrograma_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            bt_Ok_selecionaPrograma.PerformClick();
        }

        //CONTROLE DATAS DA CONSTRUÇÃO
        private void cb_status_acompanhamento_SelectedIndexChanged(object sender, EventArgs e)
        {

            //INICIADO
            if (travaDTI == false)
                if (cb_status_acompanhamento.Text == "Iniciado")
                {
                    dti_acompanhamento.Visible = true;
                    lb_dti_acompanhamento.Visible = true;
                }
                else
                {
                    dti_acompanhamento.Visible = false;
                    lb_dti_acompanhamento.Visible = false;
                }


            //FINALIZADO
            if (travaDTF == false)
                if (cb_status_acompanhamento.Text == "Finalizado")
                {
                    dtf_acompanhamento.Visible = true;
                    lb_dtf_acompanhamento.Visible = true;
                }
                else
                {
                    dtf_acompanhamento.Visible = false;
                    lb_dtf_acompanhamento.Visible = false;
                }

        }

        //BOTÃO LIMPAR - ACOMPANHAMENTO PGM
        private void bt_Limpar_Acompanhamento_Click(object sender, EventArgs e)
        {
            carregaAcompanhamentoPGM();
        }

        //RESETA CAMPOS DO CODE REVIEW 
        private void resetaCodeReiview()
        {
            lb_liberaCR.Text = "Não Liberado";
            lb_liberaCR.ForeColor = Color.Red;
            panel_CR.Visible = false;

            cb_status_CR.Text = null;
            cb_analistas_CR.Text = null;
            data_CR.Text = "";
        }

        //RESETA CAMPOS DO PERFORMACE REVIEW
        private void resetaPerformaceReiview()
        {
            lb_liberaPR.Text = "Não Liberado";
            lb_liberaPR.ForeColor = Color.Red;
            panel_PR.Visible = false;

            cb_status_PR.Text = null;
            cb_analistas_PR.Text = null;
            data_PR.Text = "";
        }

        //BOTÃO RETORNO - ACOMPANHAMENTO DE PROGRAMA
        private void bt_voltarSelecPGM_Click(object sender, EventArgs e)
        {
            DialogResult resultReturnAPGM = MessageBox.Show("Deseja voltar e escolher outro Programa?\n\nTodas alterações serão perdidas!", "Voltar", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (resultReturnAPGM == DialogResult.Yes)
            {
                tb_selecionaPrograma.Text = "";
                panel_selecionaPrograma.Visible = true;
                panel_acompanhamentoPGM.Visible = false;
                bt_Salvar_Acompanhamento.Visible = false;
                bt_Limpar_Acompanhamento.Visible = false;
            }
        }

        //KEY PRESS - ENTER - SELECIONA PROGRAMA
        private void tb_selecionaPrograma_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Keys)e.KeyChar == Keys.Enter)
                bt_Ok_selecionaPrograma.PerformClick();
        }

        //KEY PRESS - ENTER - DATA GRID - SELECIONA PROGRAMA
        private void dataGrid_selecionaPrograma_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((Keys)e.KeyChar == Keys.Enter)
                bt_Ok_selecionaPrograma.PerformClick();
        }

        #region CONTROLE DE PORCENTAGEM DE CONSTRUÇÃO
        /*
        private void tb_DA_dt1_ValueChanged(object sender, EventArgs e)
        {
            if (tb_DA_dt1.Value > oldValue_DT1)
            {
                tb_DA_dt2.Maximum--;
                tb_DA_dt3.Maximum--;
                tb_DA_dt4.Maximum--;
                tb_DA_dt5.Maximum--;
                tb_DA_dt6.Maximum--;
                tb_DA_dt7.Maximum--;
                tb_DA_dt8.Maximum--;
                tb_DA_dt9.Maximum--;
                tb_DA_dt10.Maximum--;
                tb_DA_dt11.Maximum--;
                tb_DA_dt12.Maximum--;
                tb_DA_dt13.Maximum--;
                tb_DA_dt14.Maximum--;
                tb_DA_dt15.Maximum--;
            }
            else
            {
                tb_DA_dt2.Maximum++;
                tb_DA_dt3.Maximum++;
                tb_DA_dt4.Maximum++;
                tb_DA_dt5.Maximum++;
                tb_DA_dt6.Maximum++;
                tb_DA_dt7.Maximum++;
                tb_DA_dt8.Maximum++;
                tb_DA_dt9.Maximum++;
                tb_DA_dt10.Maximum++;
                tb_DA_dt11.Maximum++;
                tb_DA_dt12.Maximum++;
                tb_DA_dt13.Maximum++;
                tb_DA_dt14.Maximum++;
                tb_DA_dt15.Maximum++;
            }

            oldValue_DT1 = (int)tb_DA_dt1.Value;
        }

        private void tb_DA_dt2_ValueChanged(object sender, EventArgs e)
        {
            if (tb_DA_dt2.Value > oldValue_DT2)
            {
                tb_DA_dt3.Maximum--;
                tb_DA_dt4.Maximum--;
                tb_DA_dt5.Maximum--;
                tb_DA_dt6.Maximum--;
                tb_DA_dt7.Maximum--;
                tb_DA_dt8.Maximum--;
                tb_DA_dt9.Maximum--;
                tb_DA_dt10.Maximum--;
                tb_DA_dt11.Maximum--;
                tb_DA_dt12.Maximum--;
                tb_DA_dt13.Maximum--;
                tb_DA_dt14.Maximum--;
                tb_DA_dt15.Maximum--;
            }
            else
            {
                tb_DA_dt3.Maximum++;
                tb_DA_dt4.Maximum++;
                tb_DA_dt5.Maximum++;
                tb_DA_dt6.Maximum++;
                tb_DA_dt7.Maximum++;
                tb_DA_dt8.Maximum++;
                tb_DA_dt9.Maximum++;
                tb_DA_dt10.Maximum++;
                tb_DA_dt11.Maximum++;
                tb_DA_dt12.Maximum++;
                tb_DA_dt13.Maximum++;
                tb_DA_dt14.Maximum++;
                tb_DA_dt15.Maximum++;
            }
            oldValue_DT2 = (int)tb_DA_dt2.Value;
        }

        private void tb_DA_dt3_ValueChanged(object sender, EventArgs e)
        {
            if (tb_DA_dt3.Value > oldValue_DT3)
            {
                tb_DA_dt4.Maximum--;
                tb_DA_dt5.Maximum--;
                tb_DA_dt6.Maximum--;
                tb_DA_dt7.Maximum--;
                tb_DA_dt8.Maximum--;
                tb_DA_dt9.Maximum--;
                tb_DA_dt10.Maximum--;
                tb_DA_dt11.Maximum--;
                tb_DA_dt12.Maximum--;
                tb_DA_dt13.Maximum--;
                tb_DA_dt14.Maximum--;
                tb_DA_dt15.Maximum--;
            }
            else
            {
                tb_DA_dt4.Maximum++;
                tb_DA_dt5.Maximum++;
                tb_DA_dt6.Maximum++;
                tb_DA_dt7.Maximum++;
                tb_DA_dt8.Maximum++;
                tb_DA_dt9.Maximum++;
                tb_DA_dt10.Maximum++;
                tb_DA_dt11.Maximum++;
                tb_DA_dt12.Maximum++;
                tb_DA_dt13.Maximum++;
                tb_DA_dt14.Maximum++;
                tb_DA_dt15.Maximum++;
            }
            oldValue_DT3 = (int)tb_DA_dt3.Value;
        }

        private void tb_DA_dt4_ValueChanged(object sender, EventArgs e)
        {
            if (tb_DA_dt4.Value > oldValue_DT4)
            {
                tb_DA_dt5.Maximum--;
                tb_DA_dt6.Maximum--;
                tb_DA_dt7.Maximum--;
                tb_DA_dt8.Maximum--;
                tb_DA_dt9.Maximum--;
                tb_DA_dt10.Maximum--;
                tb_DA_dt11.Maximum--;
                tb_DA_dt12.Maximum--;
                tb_DA_dt13.Maximum--;
                tb_DA_dt14.Maximum--;
                tb_DA_dt15.Maximum--;
            }
            else
            {
                tb_DA_dt5.Maximum++;
                tb_DA_dt6.Maximum++;
                tb_DA_dt7.Maximum++;
                tb_DA_dt8.Maximum++;
                tb_DA_dt9.Maximum++;
                tb_DA_dt10.Maximum++;
                tb_DA_dt11.Maximum++;
                tb_DA_dt12.Maximum++;
                tb_DA_dt13.Maximum++;
                tb_DA_dt14.Maximum++;
                tb_DA_dt15.Maximum++;
            }
            oldValue_DT4 = (int)tb_DA_dt4.Value;
        }

        private void tb_DA_dt5_ValueChanged(object sender, EventArgs e)
        {
            if (tb_DA_dt5.Value > oldValue_DT5)
            {
                tb_DA_dt6.Maximum--;
                tb_DA_dt7.Maximum--;
                tb_DA_dt8.Maximum--;
                tb_DA_dt9.Maximum--;
                tb_DA_dt10.Maximum--;
                tb_DA_dt11.Maximum--;
                tb_DA_dt12.Maximum--;
                tb_DA_dt13.Maximum--;
                tb_DA_dt14.Maximum--;
                tb_DA_dt15.Maximum--;
            }
            else
            {
                tb_DA_dt6.Maximum++;
                tb_DA_dt7.Maximum++;
                tb_DA_dt8.Maximum++;
                tb_DA_dt9.Maximum++;
                tb_DA_dt10.Maximum++;
                tb_DA_dt11.Maximum++;
                tb_DA_dt12.Maximum++;
                tb_DA_dt13.Maximum++;
                tb_DA_dt14.Maximum++;
                tb_DA_dt15.Maximum++;
            }
            oldValue_DT5 = (int)tb_DA_dt5.Value;
        }

        private void tb_DA_dt6_ValueChanged(object sender, EventArgs e)
        {
            if (tb_DA_dt6.Value > oldValue_DT6)
            {
                tb_DA_dt7.Maximum--;
                tb_DA_dt8.Maximum--;
                tb_DA_dt9.Maximum--;
                tb_DA_dt10.Maximum--;
                tb_DA_dt11.Maximum--;
                tb_DA_dt12.Maximum--;
                tb_DA_dt13.Maximum--;
                tb_DA_dt14.Maximum--;
                tb_DA_dt15.Maximum--;
            }
            else
            {
                tb_DA_dt7.Maximum++;
                tb_DA_dt8.Maximum++;
                tb_DA_dt9.Maximum++;
                tb_DA_dt10.Maximum++;
                tb_DA_dt11.Maximum++;
                tb_DA_dt12.Maximum++;
                tb_DA_dt13.Maximum++;
                tb_DA_dt14.Maximum++;
                tb_DA_dt15.Maximum++;
            }
            oldValue_DT6 = (int)tb_DA_dt6.Value;
        }

        private void tb_DA_dt7_ValueChanged(object sender, EventArgs e)
        {
            if (tb_DA_dt7.Value > oldValue_DT7)
            {
                tb_DA_dt8.Maximum--;
                tb_DA_dt9.Maximum--;
                tb_DA_dt10.Maximum--;
                tb_DA_dt11.Maximum--;
                tb_DA_dt12.Maximum--;
                tb_DA_dt13.Maximum--;
                tb_DA_dt14.Maximum--;
                tb_DA_dt15.Maximum--;
            }
            else
            {
                tb_DA_dt8.Maximum++;
                tb_DA_dt9.Maximum++;
                tb_DA_dt10.Maximum++;
                tb_DA_dt11.Maximum++;
                tb_DA_dt12.Maximum++;
                tb_DA_dt13.Maximum++;
                tb_DA_dt14.Maximum++;
                tb_DA_dt15.Maximum++;
            }
            oldValue_DT7 = (int)tb_DA_dt7.Value;
        }

        private void tb_DA_dt8_ValueChanged(object sender, EventArgs e)
        {
            if (tb_DA_dt8.Value > oldValue_DT8)
            {
                tb_DA_dt9.Maximum--;
                tb_DA_dt10.Maximum--;
                tb_DA_dt11.Maximum--;
                tb_DA_dt12.Maximum--;
                tb_DA_dt13.Maximum--;
                tb_DA_dt14.Maximum--;
                tb_DA_dt15.Maximum--;
            }
            else
            {
                tb_DA_dt9.Maximum++;
                tb_DA_dt10.Maximum++;
                tb_DA_dt11.Maximum++;
                tb_DA_dt12.Maximum++;
                tb_DA_dt13.Maximum++;
                tb_DA_dt14.Maximum++;
                tb_DA_dt15.Maximum++;
            }
            oldValue_DT8 = (int)tb_DA_dt8.Value;
        }

        private void tb_DA_dt9_ValueChanged(object sender, EventArgs e)
        {
            if (tb_DA_dt9.Value > oldValue_DT9)
            {
                tb_DA_dt10.Maximum--;
                tb_DA_dt11.Maximum--;
                tb_DA_dt12.Maximum--;
                tb_DA_dt13.Maximum--;
                tb_DA_dt14.Maximum--;
                tb_DA_dt15.Maximum--;
            }
            else
            {
                tb_DA_dt10.Maximum++;
                tb_DA_dt11.Maximum++;
                tb_DA_dt12.Maximum++;
                tb_DA_dt13.Maximum++;
                tb_DA_dt14.Maximum++;
                tb_DA_dt15.Maximum++;
            }
            oldValue_DT9 = (int)tb_DA_dt9.Value;
        }

        private void tb_DA_dt10_ValueChanged(object sender, EventArgs e)
        {
            if (tb_DA_dt10.Value > oldValue_DT10)
            {
                tb_DA_dt11.Maximum--;
                tb_DA_dt12.Maximum--;
                tb_DA_dt13.Maximum--;
                tb_DA_dt14.Maximum--;
                tb_DA_dt15.Maximum--;
            }
            else
            {
                tb_DA_dt11.Maximum++;
                tb_DA_dt12.Maximum++;
                tb_DA_dt13.Maximum++;
                tb_DA_dt14.Maximum++;
                tb_DA_dt15.Maximum++;
            }
            oldValue_DT10 = (int)tb_DA_dt10.Value;
        }

        private void tb_DA_dt11_ValueChanged(object sender, EventArgs e)
        {
            if (tb_DA_dt11.Value > oldValue_DT11)
            {
                tb_DA_dt12.Maximum--;
                tb_DA_dt13.Maximum--;
                tb_DA_dt14.Maximum--;
                tb_DA_dt15.Maximum--;
            }
            else
            {
                tb_DA_dt12.Maximum++;
                tb_DA_dt13.Maximum++;
                tb_DA_dt14.Maximum++;
                tb_DA_dt15.Maximum++;
            }
            oldValue_DT11 = (int)tb_DA_dt11.Value;
        }

        private void tb_DA_dt12_ValueChanged(object sender, EventArgs e)
        {
            if (tb_DA_dt12.Value > oldValue_DT12)
            {
                tb_DA_dt13.Maximum--;
                tb_DA_dt14.Maximum--;
                tb_DA_dt15.Maximum--;
            }
            else
            {
                tb_DA_dt13.Maximum++;
                tb_DA_dt14.Maximum++;
                tb_DA_dt15.Maximum++;
            }
            oldValue_DT12 = (int)tb_DA_dt12.Value;
        }

        private void tb_DA_dt13_ValueChanged(object sender, EventArgs e)
        {
            if (tb_DA_dt13.Value > oldValue_DT13)
            {
                tb_DA_dt14.Maximum--;
                tb_DA_dt15.Maximum--;
            }
            else
            {
                tb_DA_dt14.Maximum++;
                tb_DA_dt15.Maximum++;
            }
            oldValue_DT13 = (int)tb_DA_dt13.Value;
        }

        private void tb_DA_dt14_ValueChanged(object sender, EventArgs e)
        {
            if (tb_DA_dt14.Value > oldValue_DT14)
                tb_DA_dt15.Maximum--;
            else
                tb_DA_dt15.Maximum++;

            oldValue_DT14 = (int)tb_DA_dt14.Value;
        }
        */
        #endregion

        #endregion

        #region //**************************************** CADASTRO BASELINE ****************************************\\

        //VARIAVEIS UTILIZADAS
        private static string caminhoArquivoEntrada;
        private static string nomeBASELINE;
        private string scan;
        private string peekVariavel;
        private static string codPRJ_baseline;

        //BOTÃO SALVAR BASELINE
        private void bt_SalvarBaseline_Click(object sender, EventArgs e)
        {
            string codBAS = "";
            string biblioteca = "";
            string linha = "";
            string programa = "";
            string tipo_pgm = "";
            string num_linha = "";
            string result = "";
            int contLinha = 0;
            DialogResult resultNew = DialogResult.Yes;

            try
            {
                string[] Lines = File.ReadAllLines(caminhoArquivoEntrada);

                //ABRE CONEXÃO COM 
                bdConn.Open();

                using (StreamWriter arqTemp = File.CreateText(@"C:\Users\Public\Documents\basTemp.txt"))
                {
                    foreach (string line in Lines)
                    {
                        linha = line;
                        linha = linha.Replace(@"\", @"\\");

                        //CONTADOR DE LINHAS
                        contLinha++;

                        #region PRIMEIRA LINHA - (H | ...)

                        if (contLinha == 1)
                        {
                            linha = linha.Replace("'", "''");
                            int finalScan = linha.IndexOf("|", 4);
                            scan = linha.Substring(4, (finalScan - 4));
                        }

                        #endregion

                        #region SEGUNDA LINHA - (I | ...)

                        if (contLinha == 2)
                        {
                            linha = linha.Replace("'", "''");
                            //RECUPERA A STRING PESQUISADA PARA GRAVAR COMO CHAVE
                            int finalVariavel = linha.IndexOf(" |", 41);
                            peekVariavel = linha.Substring(41, (finalVariavel - 41));

                            //VERIFICA SE JÁ EXISTE ALGUM BASELINE COM A MESMA STRING ATIVADO
                            MySqlCommand verificaBaselie = new MySqlCommand("SELECT str_pes FROM prj_baseline WHERE str_pes = '" + peekVariavel + "' and cod_prj = '" + codPRJ_baseline + "' and status = 'Ativado';", bdConn);
                            MySqlDataReader dr_VB = verificaBaselie.ExecuteReader();

                            if (dr_VB.Read())
                                resultNew = MessageBox.Show("Já existe um baseline ativado com a pesquisa da string '" + peekVariavel + "'. Deseja prosseguir?", "Atenção!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            dr_VB.Close();

                            if (resultNew == DialogResult.No)
                                break;
                            else if (resultNew == DialogResult.Yes)
                            {
                                //CASO TENHA E O USUARIO PROSSIGA, ATUALIZA O STATUS PARA DESATIVADO
                                MySqlCommand updateBaseline = new MySqlCommand("UPDATE prj_baseline SET status = 'Desativado' WHERE str_pes = '" + peekVariavel + "' and cod_prj = '" + codPRJ_baseline + "';", bdConn);
                                updateBaseline.ExecuteNonQuery();
                            }

                            //RECUPERA LINHA DO COMANDO
                            int finalCmd = linha.IndexOf("|", 4);
                            linha = linha.Substring(4, (finalCmd - 4));

                            //FAZ O CADASTRO DA TABELE PRJ_BASELINE - CHAVES PRIMARIAS (STR_PES, COD_BAS)
                            MySqlCommand cmdPrj_Baseline = new MySqlCommand("INSERT INTO prj_baseline (cod_prj , str_pes, scan, info_cmd, status) VALUES ('" + codPRJ_baseline + "' ,'" + peekVariavel + "', '" + scan + "', '" + linha + "', '" + "Ativado" + "');", bdConn);
                            cmdPrj_Baseline.ExecuteNonQuery();

                            //RECUPERA CODIGO DO PROJETO
                            MySqlCommand recuperaCodBas = new MySqlCommand("SELECT cod_bas FROM prj_baseline WHERE str_pes = '" + peekVariavel + "' and cod_prj = '" + codPRJ_baseline + "' and status = 'Ativado'", bdConn);
                            MySqlDataReader dr_CodBas = recuperaCodBas.ExecuteReader();
                            if (dr_CodBas.Read())
                                codBAS = dr_CodBas["cod_bas"].ToString();
                            dr_CodBas.Close();
                        }

                        #endregion

                        #region TERCEIRA EM DIANTE - (R | ...)

                        if (contLinha >= 3)
                        {
                            biblioteca = linha.Substring(2, 8);
                            programa = linha.Substring(11, 8);
                            tipo_pgm = linha.Substring(20, 10);
                            num_linha = linha.Substring(31, 4);
                            result = linha.Substring(36, 80);
                            arqTemp.WriteLine(String.Format("{0,-12}^{1,-8}^{2,-8}^{3,-10}^{4,-4}^{5,-80}^", codBAS, biblioteca, programa, tipo_pgm, num_linha, result));
                        }

                        #endregion
                    }
                }

                //GRAVA ARQUIVO COM CAMPOS DO BASELINE
                MySqlCommand loadData = new MySqlCommand("LOAD DATA LOCAL INFILE 'C:/Users/Public/Documents/basTemp.txt' INTO TABLE baseline FIELDS TERMINATED BY '^';", bdConn);
                loadData.ExecuteNonQuery();

                //EXCLUI ARQUIVO CRIADO
                File.Delete(@"C:/Users/Public/Documents/basTemp.txt");

                //FECHA CONEXÃO COM O BD
                bdConn.Close();
                //}

                //EXIBE MENSAGEM DE CONCLUIDO
                if (!resultNew.Equals(DialogResult.No))
                    botaoConcluido("Baseline cadastrado com sucesso.");

                //LIMPA FORM
                bt_LimparBaseline.PerformClick();
            }
            catch (Exception ex)
            {
                this.Opacity = 0.9;
                ErrorForm erro = new ErrorForm(ex);
                erro.ShowDialog();
                bdConn.Close();
                this.Close();
            }
        }

        //BOTÃO PARA CARREGAR ARQUIVO DO BASELINE
        private void bt_carregaArquivo_Baseline_Click(object sender, EventArgs e)
        {
            try
            {
                bt_LimparBaseline.PerformClick();

                //CRIA OBJETO
                OpenFileDialog ofd = new OpenFileDialog();

                //DEFINE AS PROPRIEDADES DE CONTROLE           
                ofd.Multiselect = false;
                ofd.Title = "Selecionar Arquivo";
                ofd.InitialDirectory = @"C:/Users/Public/Documents";
                ofd.Filter = "All files (*.*)|*.*";
                ofd.CheckFileExists = true;
                ofd.CheckPathExists = true;
                ofd.RestoreDirectory = true;
                ofd.ReadOnlyChecked = true;
                ofd.ShowReadOnly = true;

                //ABRE CAIXA DE SELEÇAO
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    caminhoArquivoEntrada = ofd.FileName;
                    nomeBASELINE = System.IO.Path.GetFileName(caminhoArquivoEntrada);
                }
                else
                    caminhoArquivoEntrada = "";

                tb_caminhoBS.Text = caminhoArquivoEntrada;

                //VERIFICA SE CAMINHO FOI PREENCHIDO CORRETAMENTE
                if (!String.IsNullOrEmpty(caminhoArquivoEntrada))
                    if (File.Exists(caminhoArquivoEntrada))
                        carregaBaseline();

            }
            catch (Exception ex)
            {
                this.Opacity = 0.9;
                ErrorForm erro = new ErrorForm(ex);
                erro.ShowDialog();
                bdConn.Close();
                this.Close();
            }
        }

        //EXIBE O ARQUIVO CARREGADO
        public void carregaBaseline()
        {
            try
            {
                //VERIFICA ARQUIVO
                bool checkArquivo = true;

                string[] arquivo = File.ReadAllLines(caminhoArquivoEntrada);

                int tamArq = arquivo.Count();

                //PROGRESS BAR                        
                progressBar1.Visible = true;
                progressBar1.Maximum = tamArq;

                for (int i = 1; i < tamArq; i++)
                {
                    if (i == 0)
                        if (!arquivo[i].Substring(0, 26).Equals("H | RESULTADO DE SCAN PARA"))
                        {
                            botaoAlert("Arquivo inserido é invalido!");
                            checkArquivo = false;
                            break;
                        }

                    if (i == 1)
                        if (!arquivo[i].Substring(0, 20).Equals("I | COMANDO EMITIDO:"))
                        {
                            botaoAlert("Arquivo inserido é invalido!");
                            checkArquivo = false;
                            break;
                        }

                    if (i == 2)
                    {
                        if (!arquivo[i].Substring(0, 2).Equals("R|"))
                        {
                            botaoAlert("Arquivo inserido é invalido!");
                            checkArquivo = false;
                            break;
                        }

                        progressBar1.Visible = true;
                        Thread.Sleep(TimeSpan.FromSeconds(1));
                        Application.DoEvents();
                        this.Update();
                    }

                    progressBar1.Value = i + 1;
                    tb_exibeBaseline.AppendText(arquivo[i]);
                    tb_exibeBaseline.AppendText(Environment.NewLine);
                }

                if (checkArquivo)
                {
                    botaoAlert("Para salvar este baseline, pressione o botão Salvar.");

                    progressBar1.Visible = false;
                    progressBar1.Value = 0;

                    tb_exibeBaseline.Visible = true;

                    bt_SalvarBaseline.Visible = true;
                    bt_LimparBaseline.Visible = true;
                }
                else
                    tb_exibeBaseline.Text = "";
            }
            catch
            {
                botaoAlert("Arquivo inserido é invalido!");
            }
        }

        //SELECIONAR PROJETO
        private void tb_Prj_Baseline_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (tb_Prj_Baseline.Text != "")
                {
                    string querySPRJ_baseline = "select cod_prj from prj_objeto where cod_prj like '%" + tb_Prj_Baseline.Text + "%'";
                    bdDataSet = new DataSet();
                    bdConn.Open();
                    bdAdapter = new MySqlDataAdapter(querySPRJ_baseline, bdConn);
                    bdAdapter.Fill(bdDataSet, "prj_objeto");
                    dataGrid_Prj_Baseline.DataSource = bdDataSet;

                    if (bdDataSet.Tables["prj_objeto"].Rows.Count == 0)
                        lb_NotFound_Baseline.Visible = true;
                    else
                        lb_NotFound_Baseline.Visible = false;

                    dataGrid_Prj_Baseline.DataMember = "prj_objeto";
                    bdConn.Close();
                }
                else
                {
                    lb_NotFound_Baseline.Visible = false;

                    if (this.dataGrid_Prj_Baseline.DataSource != null)
                        this.dataGrid_Prj_Baseline.DataSource = null;
                    else
                    {
                        this.dataGrid_Prj_Baseline.Rows.Clear();
                        this.dataGrid_Prj_Baseline.Columns.Clear();
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorForm erro = new ErrorForm(ex);
                erro.ShowDialog();
                bdConn.Close();
            }
        }

        //BOTÃO OK PROJETO SELECIONADO
        private void bt_Ok_Prj_Baseline_Click(object sender, EventArgs e)
        {


            if (dataGrid_Prj_Baseline.CurrentRow == null)
                botaoAlert("Selecionar um programa da lista antes de prosseguir.");
            else
            {
                if (bdDataSet.Tables["prj_objeto"].Rows.Count > 0)
                {
                    panel_SP_Baseline.Visible = false;

                    codPRJ_baseline = dataGrid_Prj_Baseline.CurrentRow.Cells[0].Value.ToString();
                    tb_projetoBaseline.Text = codPRJ_baseline;
                }
                else
                    botaoAlert("Nenhum programa foi encontrado! Pesquise novamente.");
            }

        }

        //BOTÃO VOLTAR E SELECIONAR OUTRO PROJETO
        private void bt_returnBaseline_Click(object sender, EventArgs e)
        {
            panel_SP_Baseline.Visible = true;
            bt_SalvarBaseline.Visible = false;
            bt_LimparBaseline.Visible = false;
            tb_exibeBaseline.Visible = false;

            codPRJ_baseline = "";
            tb_Prj_Baseline.Text = "";

            lb_NotFound_Baseline.Visible = false;

            if (this.dataGrid_Prj_Baseline.DataSource != null)
                this.dataGrid_Prj_Baseline.DataSource = null;
            else
            {
                this.dataGrid_Prj_Baseline.Rows.Clear();
                this.dataGrid_Prj_Baseline.Columns.Clear();
            }
        }

        //DUPLO CLICK EM GRIDVIEW
        private void dataGrid_Prj_Baseline_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            bt_Ok_Prj_Baseline.PerformClick();
        }

        //BOTÃO LIMPAR CADASTRO DO BASELINE
        private void bt_LimparBaseline_Click(object sender, EventArgs e)
        {
            caminhoArquivoEntrada = "";
            tb_exibeBaseline.Text = "";
            tb_caminhoBS.Text = "";

            tb_exibeBaseline.Visible = false;

            bt_SalvarBaseline.Visible = false;
            bt_LimparBaseline.Visible = false;

            progressBar1.Visible = false;
            progressBar1.Value = 0;
        }

        #endregion

        #region //**************************************** CONSULTA BASELINE ****************************************\\

        //BOTÃO PESQUISA BASELINE
        private void bt_PesquisarBaseline_Click(object sender, EventArgs e)
        {

            try
            {
                bdConn.Open();
                bdDataSet = new DataSet();
                bdAdapter = new MySqlDataAdapter(queryPesquisaBaseline(), bdConn);
                bdAdapter.Fill(bdDataSet, "prj_baseline");
                dataGrid_PesquisaBaseline.DataSource = bdDataSet;

                if (bdDataSet.Tables["prj_baseline"].Rows.Count == 0)
                    panel_NotFound_Baseline.Visible = true;
                else
                {
                    panel_NotFound_Baseline.Visible = false;

                    dataGrid_PesquisaBaseline.DataMember = "prj_baseline";

                    this.dataGrid_PesquisaBaseline.Columns[0].HeaderText = "Projeto";
                    this.dataGrid_PesquisaBaseline.Columns[1].HeaderText = "String Pesquisada";
                    this.dataGrid_PesquisaBaseline.Columns[2].HeaderText = "Status";
                }

                bdConn.Close();
            }
            catch (Exception ex)
            {
                this.Opacity = 0.9;
                ErrorForm erro = new ErrorForm(ex);
                erro.ShowDialog();
                bdConn.Close();
                this.Close();
            }
        }

        //CRIA QUERY DE PESQUISA PARA BASELINE
        private string queryPesquisaBaseline()
        {
            string queryReturn = "SELECT cod_prj, str_pes,status FROM prj_baseline WHERE";

            if (cb_Projeto_PBaseline.Text != "")
                queryReturn += " cod_prj like '%" + cb_Projeto_PBaseline.Text + "%'";
            else
                queryReturn += " cod_prj like '%'";

            if (tb_String_PBaseline.Text != "")
                queryReturn += " and str_pes like '%" + tb_String_PBaseline.Text + "%'";

            if (checkBox_Ativo_PBaseline.Checked && checkBox_Desativado_PBaseline.Checked)
            {

            }
            else
            {
                if (checkBox_Ativo_PBaseline.Checked == false && checkBox_Desativado_PBaseline.Checked == false)
                    queryReturn += " and status = 'Nenhum'";

                if (checkBox_Ativo_PBaseline.Checked)
                    queryReturn += " and status = 'Ativado'";

                if (checkBox_Desativado_PBaseline.Checked)
                    queryReturn += " and status = 'Desativado'";
            }

            queryReturn += " order by cod_prj, str_pes;";

            return queryReturn;
        }

        //FORMATAR DRIDVIEW BASELINE
        private void dataGrid_PesquisaBaseline_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.Value != null && e.ColumnIndex == 2)
                if (e.Value.Equals("Ativado"))
                    e.CellStyle.ForeColor = Color.Green;
                else if (e.Value.Equals("Desativado"))
                    e.CellStyle.ForeColor = Color.Red;
        }

        //BOTÃO PARA LIMPAR A PESQUISA DO BASELINE
        private void bt_LimpaConsBaseline_Click(object sender, EventArgs e)
        {
            //LIMPA GRIDVIEW
            if (this.dataGrid_PesquisaBaseline.DataSource != null)
                this.dataGrid_PesquisaBaseline.DataSource = null;
            else
            {
                this.dataGrid_PesquisaBaseline.Rows.Clear();
                this.dataGrid_PesquisaBaseline.Columns.Clear();
            }

            //LIMPA NOT FOUND
            panel_NotFound_Baseline.Visible = false;

            //LIMPAR FILTROS
            cb_Projeto_PBaseline.Text = null;
            tb_String_PBaseline.Text = "";
        }

        #endregion

        #region //**************************************** CADASTRO DE GRUPO ****************************************\\

        //BOTÃO SALVAR GRUPO
        private void bt_SalvarGrupo_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(tb_NomeGrupo.Text) || cb_LGrupo.Text == "")
                botaoAlert("O preenchimento dos campos são obrigatorios!");
            else
            {

                try
                {
                    bdConn.Open();

                    MySqlCommand cmd = new MySqlCommand("INSERT INTO grupo_analistas (lider_grupo, nom_grupo) VALUES ('" + cb_LGrupo.Text + "','" + tb_NomeGrupo.Text + "');", bdConn);
                    cmd.ExecuteNonQuery();

                    botaoConcluido("O Grupo foi cadastrado com sucesso!");

                    bt_LimparGrupo.PerformClick();

                    bdConn.Close();
                }
                catch (Exception ex)
                {
                    this.Opacity = 0.9;
                    ErrorForm erro = new ErrorForm(ex);
                    erro.ShowDialog();
                    bdConn.Close();
                    this.Close();
                }
            }
        }

        //BOTÃO LIMPAR GRUPO
        private void bt_LimparGrupo_Click(object sender, EventArgs e)
        {
            tb_NomeGrupo.Text = "";
            cb_NomeGrupoEdit.Text = null;
            cb_LGrupo.Text = null;
        }

        //ATUALIZA COMBO BOX NOME GRUPOS
        public void atualzaCBNomeGrupo()
        {
            cb_NomeGrupoEdit.Items.Clear();

            bdConn.Open();
            MySqlCommand commandG = new MySqlCommand("SELECT nom_grupo FROM grupo_analistas;", bdConn);
            MySqlDataReader drG = commandG.ExecuteReader();
            while (drG.Read())
                cb_NomeGrupoEdit.Items.Add(drG["nom_grupo"].ToString());
            drG.Close();
            bdConn.Close();
        }

        //BOTÃO EDITAR GRUPO
        private void bt_EditarGrupo_Click(object sender, EventArgs e)
        {
            try
            {
                bdConn.Open();

                MySqlCommand cmd = new MySqlCommand("UPDATE grupo_analistas SET lider_grupo = '" + cb_LGrupo.Text + "' WHERE nom_grupo = '" + cb_NomeGrupoEdit.Text + "'", bdConn);
                cmd.ExecuteNonQuery();

                bdConn.Close();

                botaoConcluido("O Grupo '" + cb_NomeGrupoEdit.Text + "' foi atualizado com sucesso.");
            }
            catch (Exception ex)
            {
                this.Opacity = 0.9;
                ErrorForm erro = new ErrorForm(ex);
                erro.ShowDialog();
                bdConn.Close();
                this.Close();
            }
        }

        //RECUPERA LIDER DO GRUPO
        private void cb_NomeGrupoEdit_SelectedIndexChanged(object sender, EventArgs e)
        {
            bdConn.Open();
            MySqlCommand commandG = new MySqlCommand("SELECT lider_grupo FROM grupo_analistas WHERE nom_grupo = '" + cb_NomeGrupoEdit.Text + "';", bdConn);
            MySqlDataReader drG = commandG.ExecuteReader();
            if (drG.Read())
                cb_LGrupo.Text = drG["lider_grupo"].ToString();

            if (cb_NomeGrupoEdit.Text == "Nenhum")
                cb_LGrupo.Text = "Nenhum";

            drG.Close();
            bdConn.Close();
        }

        //BOTÃO EXCLUIR GRUPO
        private void bt_ExcluirGrupo_Click(object sender, EventArgs e)
        {
            try
            {
                bdConn.Open();
                int codGrupo = 01;

                MySqlCommand cmd = new MySqlCommand("SELECT cod_grupo FROM grupo_analistas WHERE nom_grupo = '" + cb_NomeGrupoEdit.Text + "'", bdConn);
                MySqlDataReader dr = cmd.ExecuteReader();

                if (dr.Read())
                    codGrupo = Int16.Parse(dr["cod_grupo"].ToString());
                dr.Close();

                cmd = new MySqlCommand("UPDATE analistas SET cod_grupo = 01 WHERE cod_grupo = " + codGrupo, bdConn);
                cmd.ExecuteNonQuery();

                cmd = new MySqlCommand("DELETE FROM grupo_analistas WHERE nom_grupo = '" + cb_NomeGrupoEdit.Text + "'", bdConn);
                cmd.ExecuteNonQuery();

                bdConn.Close();

                botaoConcluido("O Grupo '" + cb_NomeGrupoEdit.Text + "' foi excluido com sucesso.");

                bt_LimparGrupo.PerformClick();

                atualzaCBNomeGrupo();
            }
            catch (Exception ex)
            {
                this.Opacity = 0.9;
                ErrorForm erro = new ErrorForm(ex);
                erro.ShowDialog();
                bdConn.Close();
                this.Close();
            }
        }

        //MASCARA RQF
        private void rb_RQF_CheckedChanged(object sender, EventArgs e)
        {
            tb_RQF_CadastroRQF.Mask = "RQF99";
        }

        //MASCARA RQNF
        private void rb_RQNF_CheckedChanged(object sender, EventArgs e)
        {
            tb_RQF_CadastroRQF.Mask = "RQNF99";
        }

        #endregion

        #region //**************************************** PESQUISA DE PROJETO ****************************************\\

        //BOTÃO PESQUISAR PROJETO
        private void bt_Pesquisar_Projeto_Click(object sender, EventArgs e)
        {
            //CRIANDO DATASET E POVOANDO                
            bdDataSet = new DataSet();
            bdConn.Open();
            bdAdapter = new MySqlDataAdapter(cria_queryPESQUISA_PROJETO(), bdConn);
            bdAdapter.Fill(bdDataSet, "prj_objeto");
            dataGrid_PesquisaProjeto.DataSource = bdDataSet;
            dataGrid_PesquisaProjeto.DataMember = "prj_objeto";

            if (dataGrid_PesquisaProjeto.RowCount == 0)
                semRESULTADO_PROJETO();
            else
            {
                panel_NotFoundPrj.Visible = false;
                //FORMATA GRIDVIEW
                formataGRIDVIEWPRJ();
            }


            //FECHA CONEXÃO
            bdConn.Close();
        }

        //CRIA QUERY DE PESQUISA
        string cria_queryPESQUISA_PROJETO()
        {
            string campos = "cod_prj, " +
                "lider_tecnico, " +
                "lider_requer, " +
                "resp_tecnico, " +
                "status_prj, " +
                "dt_ini_prj, " +
                "dt_fim_prj, " +
                "release_prj " +
                "";

            string query = "SELECT " + campos + " FROM prj_objeto WHERE cod_prj like '%" + cb_Projetos_PesqProjeto.Text + "%' ";

            if (cb_LT_PesqProjeto.Text != "")
                query += "AND lider_tecnico = '" + cb_LT_PesqProjeto.Text + "' ";

            if (checkBox_EmAndamento_PesqProjeto.Checked)
                query += "AND status_prj = 'Em Andamento' ";

            if (checkBox_Finalizado_PesqProjeto.Checked)
                query += "AND status_prj = 'Finalizado' ";

            query += "ORDER BY cod_prj";

            return query;
        }

        //SEM RESULTADOS PARA A PESQUISA
        void semRESULTADO_PROJETO()
        {
            //LIMPA GRIDVIEW
            this.dataGrid_PesquisaProjeto.Columns.Clear();

            //DESATIVA MSG NOT FOUND
            panel_NotFoundPrj.Visible = true;

            //EXPORTAR FALSE
            //bt_ExportarExecel.Visible = false;

            //RESETA EXPAND. E COMPR.
            //bt_ComprimirGrid.Visible = false;
            //bt_ExpandirGrid.Visible = false;
            dataGrid_PesquisaProjeto.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        //FORMATA HEADER GRIDVIEW
        void formataGRIDVIEWPRJ()
        {
            try
            {
                this.dataGrid_PesquisaProjeto.Columns[0].HeaderText = "Projeto";
                this.dataGrid_PesquisaProjeto.Columns[1].HeaderText = "Líder Técnico";
                this.dataGrid_PesquisaProjeto.Columns[2].HeaderText = "Líder de Requerimento";
                this.dataGrid_PesquisaProjeto.Columns[3].HeaderText = "Responsável Técnico";
                this.dataGrid_PesquisaProjeto.Columns[4].HeaderText = "Status";
                this.dataGrid_PesquisaProjeto.Columns[5].HeaderText = "Inicio";
                this.dataGrid_PesquisaProjeto.Columns[6].HeaderText = "Fim";
                this.dataGrid_PesquisaProjeto.Columns[7].HeaderText = "Release";
            }
            catch (Exception ex)
            {
                ErrorForm erro = new ErrorForm(ex);
                erro.ShowDialog();
            }
        }

        //BOTÃO ATIVA FILTRO
        private void bt_FiltroPRJ_Click(object sender, EventArgs e)
        {
            if (!panel_Filtro_PesquisaPrj.Visible)
                panel_Filtro_PesquisaPrj.Visible = true;
            else
                panel_Filtro_PesquisaPrj.Visible = false;
        }

        //BOTÃO FECHA FILTRO
        private void bt_Close_FiltroPrj_Click(object sender, EventArgs e)
        {
            panel_Filtro_PesquisaPrj.Visible = false;
            LimparFiltros_PesqProjeto();
        }

        //BOTÃO PESQUISA FILTRO
        private void bt_Pesq_FiltroPrj_Click(object sender, EventArgs e)
        {
            bt_Pesquisar_Projeto.PerformClick();
        }

        //BOTÃO LIMPAR
        private void bt_Limpar_PesqProjeto_Click(object sender, EventArgs e)
        {
            //LIMPAR GRIDVIEW
            limpaGrid_PesqProjeto();

            //LIMPA NOT FOUND
            panel_NotFoundPrj.Visible = false;

            //BOTÃO EXPORTAR EXCEL
            //bt_ExportarExecel.Visible = false;

            //RESETA EXPAND. E COMPR.            
            dataGrid_PesquisaProjeto.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            //LIMPAR CAMPOS
            LimparFiltros_PesqProjeto();
        }
        
        //LIMPAR GRIDVIEW
        public void limpaGrid_PesqProjeto()
        {
            if (this.dataGrid_PesquisaProjeto.DataSource != null)
                this.dataGrid_PesquisaProjeto.DataSource = null;
            else
            {
                this.dataGrid_PesquisaProjeto.Rows.Clear();
                this.dataGrid_PesquisaProjeto.Columns.Clear();
            }
        }

        //LIMPAR CAMPOS FILTROS DA PESQUISA
        void LimparFiltros_PesqProjeto()
        {
            cb_Projetos_PesqProjeto.Text = null;
            cb_LT_PesqProjeto.Text = null;
            checkBox_EmAndamento_PesqProjeto.Checked = false;
            checkBox_Finalizado_PesqProjeto.Checked = false;
        }

        #endregion                               

      
       
       
        //**************************************** ***************************** ****************************************\\
    }
}
