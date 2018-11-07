using RightNow.AddIns.AddInViews;
using System;
using System.AddIn;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using WebServiceCall;

namespace RN_AddIn_Farmacovigilancia {
    public class WorkspaceRibbonAddIn : Panel, IWorkspaceRibbonButton {
        string sexo;
        string gravidez;
        string consequencia;
        string relacaoMedicamento;
        string evolucaoDoPaciente;
        string eventoCessouComInterrupcao;
        string houveReintroducao;
        string profissionalSaude;

        private IRecordContext RecordContext {
            get; set;
        }

        public WorkspaceRibbonAddIn(bool inDesignMode, IRecordContext RecordContext) {
            this.RecordContext = RecordContext;
        }

        #region IWorkspaceRibbonButton Members
        public new void Click() {
            gravarArquivo();
        }
        #endregion

        static string extrairValor(string xmlResposta) {
            string valor = "";
            XmlDocument document = new XmlDocument();
            document.LoadXml(xmlResposta);
            var namespaceManager = new XmlNamespaceManager(document.NameTable);
            namespaceManager.AddNamespace("soapenv", "http://schemas.xmlsoap.org/soap/envelope/");
            namespaceManager.AddNamespace("xsi", "http://www.w3.org/2001/XMLSchema-instance");
            namespaceManager.AddNamespace("patriiicia", "urn:messages.ws.rightnow.com/v1_3");
            valor = document.SelectSingleNode("//patriiicia:Row", namespaceManager).InnerText;
            valor = Regex.Replace(valor, "^[\"]+(.+?)[\"]+$", "$1");
            if(valor == null || !valor.Equals("")) {
                return valor;
            }
            else {
                return "N/I";
            }
        }

        static List<Dictionary<string, string>> extrairTabela(string xmlResposta) {
            List<Dictionary<string, string>> listaConcomitante = new List<Dictionary<string, string>>();
            XmlDocument document = new XmlDocument();
            document.LoadXml(xmlResposta);
            var namespaceManager = new XmlNamespaceManager(document.NameTable);
            namespaceManager.AddNamespace("soapenv", "http://schemas.xmlsoap.org/soap/envelope/");
            namespaceManager.AddNamespace("xsi", "http://www.w3.org/2001/XMLSchema-instance");
            namespaceManager.AddNamespace("patriiicia", "urn:messages.ws.rightnow.com/v1_3");
            XmlNodeList rowList = document.SelectNodes("//patriiicia:Row", namespaceManager);
            foreach(XmlNode node in rowList) {
                string row = node.InnerText;
                string[] separator = { ";" };
                string[] campos = row.Split(separator, StringSplitOptions.None);
                for(int i = 0; i < campos.Length; i++) {
                    campos[i] = Regex.Replace(campos[i], "^[\"]+(.+?)[\"]+$", "$1");
                    if(campos[i] == null || campos[i].Equals("")) {
                        campos[i] = "N/I";
                    }
                }
                Dictionary<string, string> usoConcomitante = new Dictionary<string, string>();
                usoConcomitante.Add("NomeConGen", campos[0]);
                usoConcomitante.Add("DoseDiaria", campos[1]);
                usoConcomitante.Add("ViaAdm", campos[2]);
                usoConcomitante.Add("IndicUso", campos[3]);
                usoConcomitante.Add("DataInic", campos[4]);
                usoConcomitante.Add("DataFim", campos[5]);
                usoConcomitante.Add("DataReintr", campos[6]);
                listaConcomitante.Add(usoConcomitante);
            }
            return listaConcomitante;
        }

        static string gerarHtmlUsoConcomitante(List<Dictionary<string, string>> listaConcomitante) {
            string html = "";
            foreach(Dictionary<string, string> usoConcomitante in listaConcomitante) {
                html += gerarLinhaConcomitante(usoConcomitante);
            }
            return html;
        }

        private static string gerarLinhaConcomitante(Dictionary<string, string> usoConcomitante) {
            string linhaConcomitante =
                "<tr>" +
                "  <td colspan=\"85\"><span>" + usoConcomitante["NomeConGen"] + "</span></td>" +
                "  <td colspan=\"20\"><span>" + usoConcomitante["DoseDiaria"] + "</span></td>" +
                "  <td colspan=\"20\"><span>" + usoConcomitante["ViaAdm"] + "</span></td>" +
                "  <td colspan=\"30\"><span>" + usoConcomitante["IndicUso"] + "</span></td>" +
                "  <td colspan=\"30\"><span>" + usoConcomitante["DataInic"] + "</span></td>" +
                "  <td colspan=\"30\"><span>" + usoConcomitante["DataFim"] + "</span></td>" +
                "  <td colspan=\"25\"><span>" + usoConcomitante["DataReintr"] + "</span></td>" +
                "</tr>";
            return linhaConcomitante;
        }

        public void gravarArquivo() {
            IIncident inc = (IIncident) RecordContext.GetWorkspaceRecord(RightNow.AddIns.Common.WorkspaceRecordType.Incident);
            IContact contato = (IContact) RecordContext.GetWorkspaceRecord(RightNow.AddIns.Common.WorkspaceRecordType.Contact);
            string contatoPrimarioId = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT Incident.PrimaryContact.Contact.id FROM Incident WHERE Incident.Id = " + inc.ID));
            string matriculaFuncionario = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.CO.PV.CustomFields.C.matricula_funcionario FROM Incident WHERE Incident.Id = " + inc.ID));
            string setor = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.CO.PV.CustomFields.C.setor FROM Incident WHERE Incident.Id = " + inc.ID));
            string telefoneComercial = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.CO.PV.phones.Number FROM Incident WHERE CustomFields.CO.PV.phones.phonetype = 0  AND Incident.Id = " + inc.ID));
            string iniciaisPaciente = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.C.iniciais_paciente FROM Incident WHERE Incident.Id = " + inc.ID));
            string dataNasc = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.CO.data_nasc_pacient FROM Incident WHERE Incident.Id = " + inc.ID));
            string idade = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.C.idade_pacient FROM Incident WHERE Incident.Id = " + inc.ID));
            string peso = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.C.peso_pacient FROM Incident WHERE Incident.id = " + inc.ID));
            string altura = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.C.alt_pacient FROM Incident WHERE Incident.Id = " + inc.ID));
            string dataInicioEvento = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.C.data_inic_evento FROM Incident WHERE Incident.Id = " + inc.ID));
            string dataFinalEvento = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.C.data_final_evento FROM Incident WHERE Incident.Id = " + inc.ID));
            string descricaoEvento = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.C.desc_event FROM Incident WHERE Incident.Id = " + inc.ID));
            string nomeMedicamentoSuspeito = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.CO.Produto.Name FROM Incident WHERE Incident.Id = " + inc.ID));
            string lote = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.CO.Lote.NumeroLote FROM incident WHERE incident.Id = " + inc.ID));
            string viaAdministracao = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.c.via_de_administracao FROM Incident WHERE Incident.Id = " + inc.ID));
            string doseDiaria = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.C.dose_diaria FROM incident WHERE incident.Id = " + inc.ID));
            string indicacaoDeUso = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.C.indicacao_de_uso FROM incident WHERE incident.Id = " + inc.ID));
            string dataInicialTratamento = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.C.data_inic_trat FROM incident WHERE incident.Id = " + inc.ID));
            string dataFinalTratamento = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.C.data_final_trat FROM incident WHERE incident.Id = " + inc.ID));
            string duracaoTratamento = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.C.duracao_do_tratamento FROM incident WHERE incident.Id = " + inc.ID));

            List<Dictionary<string, string>> usoConcomitanteLista = extrairTabela(RightnowWSSuperaTeste.realizarQueryTabela("SELECT U.NomeConGen, U.DoseDiaria, U.ViaAdm, U.IndicUso,"
                + " U.DataInic, U.DataFim, U.DataReintr FROM CO.UsoConcomitante U WHERE U.Incidente = " + inc.ID));

            string usoConcomitante = gerarHtmlUsoConcomitante(usoConcomitanteLista);
            string dadosClinicosRelevantes = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.c.dados_relevantes FROM Incident WHERE Incident.Id = " + inc.ID));
            string nomeContato = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT Contact.Name.First FROM Contact WHERE Contact.id = " + contatoPrimarioId));
            string inscricaoProfissional = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.C.insc_prof FROM Contact WHERE Contact.id = " + contatoPrimarioId));
            string ruaContato = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT Contact.Address.Street FROM Contact WHERE Contact.id = " + contatoPrimarioId));
            string cidadeContato = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT Contact.Address.City FROM Contact WHERE Contact.id = " + contatoPrimarioId));
            string cepContato = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT Contact.Address.PostalCode FROM Contact WHERE Contact.id = " + contatoPrimarioId));
            string estadoContato = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT Contact.Address.StateOrProvince.Name FROM Contact WHERE Contact.id = " + contatoPrimarioId));
            string telefoneAssistente = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT Contact.Phones.Number FROM Contact WHERE Contact.Phones.PhoneType = 3 AND Contact.Id = " + contatoPrimarioId));
            string celularContato = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.c.telefone_celular FROM Contact WHERE Contact.Id = " + contatoPrimarioId));
            string emailContato = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT Contact.Emails.Address FROM Contact WHERE Contact.Emails.AddressType = 0 AND Contact.id = " + contatoPrimarioId));
            string informacoesComplementares = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.c.informacoes_complementares FROM Incident WHERE Incident.Id = " + inc.ID));
            string especialidade = "";
            string especificacaoInterrupcao = "";
            string especificacaoReintroducao = "";
            string especObito = "";
            string dataFarmaco = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.c.data_farmaco FROM Incident WHERE Incident.Id = " + inc.ID));

            sexo = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.C.sexo_pacient.name FROM incident WHERE incident.Id = " + inc.ID));
            gravidez = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.C.gravid_pacient.name FROM Incident WHERE Incident.Id = " + inc.ID));
            consequencia = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.C.consequencia.name FROM Incident WHERE Incident.Id = " + inc.ID));
            relacaoMedicamento = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.C.relacao_medicamento.name FROM Incident WHERE Incident.Id = " + inc.ID));
            evolucaoDoPaciente = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.C.evolucao_do_paciente.name FROM Incident WHERE Incident.Id = " + inc.ID));
            eventoCessouComInterrupcao = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.C.evento_cess_interrr.name FROM incident WHERE incident.Id = " + inc.ID));
            houveReintroducao = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.C.houve_reintrod.name FROM incident WHERE incident.Id = " + inc.ID));
            profissionalSaude = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.c.profis_saude.Name FROM Contact WHERE Contact.id = " + contatoPrimarioId));

            if(telefoneComercial != null && !telefoneComercial.Equals("N/I") && !telefoneComercial.Equals("")) {
                if(telefoneComercial.Length >= 10) {
                    telefoneComercial = telefoneComercial.Insert(2, " ");
                }
                if(telefoneAssistente.Length >= 10) {
                    telefoneAssistente = telefoneAssistente.Insert(2, " ");
                }
                if(celularContato.Length >= 10) {
                    celularContato = celularContato.Insert(2, " ");
                }
            }

            if(informacoesComplementares != null && !informacoesComplementares.Equals("N/I") && !informacoesComplementares.Equals("")) {
                informacoesComplementares = informacoesComplementares.ToUpper();
            }

            if(dataNasc != null && !dataNasc.Equals("N/I") && !dataNasc.Equals("")) {
                DateTime dateTime = DateTime.ParseExact(dataNasc, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                dataNasc = dateTime.ToString("dd/MM/yyyy");
            }

            if(dataFarmaco != null && !dataFarmaco.Equals("N/I") && !dataFarmaco.Equals("")) {
                dataFarmaco = dataFarmaco.Substring(0, 10);
                DateTime dateTime = DateTime.ParseExact(dataFarmaco, "yyyy-MM-dd", CultureInfo.InvariantCulture);
                dataFarmaco = dateTime.ToString("dd/MM/yyyy");
            }

            if(informacoesComplementares != null && !informacoesComplementares.Equals("N/I") && !informacoesComplementares.Equals("") && profissionalSaude.Equals("Sim")) {
                especialidade = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.c.espec FROM Contact WHERE Contact.id = " + contatoPrimarioId));
            }

            if(eventoCessouComInterrupcao != null && !eventoCessouComInterrupcao.Equals("N/I") && !eventoCessouComInterrupcao.Equals("") && eventoCessouComInterrupcao.Equals("Sim")) {
                especificacaoInterrupcao = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.c.espec_interr FROM Incident WHERE Incident.id = " + inc.ID));
            }

            if(houveReintroducao != null && !houveReintroducao.Equals("N/I") && !houveReintroducao.Equals("") && houveReintroducao.Equals("Sim")) {
                especificacaoReintroducao = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.c.espec_reintr FROM Incident WHERE Incident.id = " + inc.ID));
            }

            if(evolucaoDoPaciente != null && !evolucaoDoPaciente.Equals("N/I") && !evolucaoDoPaciente.Equals("") && evolucaoDoPaciente.Equals("Óbito")) {
                especObito = extrairValor(RightnowWSSuperaTeste.realizarQueryCampo("SELECT CustomFields.c.espec_obito FROM Incident WHERE Incident.id = " + inc.ID));
            }

            string html1 =
                "<!DOCTYPE html>" +
                "<html>" +
                "<head>" +
                "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />" +
                "<title>Farmacovigilância</title>" +
                "<style type=\"text/css\">" +
                "p {text-align: center;}" +
                "input {vertical-align: middle;}" +
                "label {font-weight: normal;}" +
                "table {" +
                "margin: 0px;" +
                "table-layout: fixed;" +
                "border: 0;" +
                "border-right: 1px solid #000 !important;" +
                "}" +
                "body {" +
                "font-family: Arial, Helvetica, sans-serif;" +
                "font-size: 7px;" +
                "font-weight: bold;" +
                "font-style: normal;" +
                "font-variant: normal;" +
                "text-transform: uppercase;" +
                "text-align: left;" +
                "border-collapse: collapse;" +
                "border-top-color: #000;" +
                "border-right-color: #000;" +
                "border-bottom-color: #000;" +
                "border-left-color: #000;" +
                "vertical-align: top;" +
                "}" +
                "/*body c {" +
                "font-family: Arial, Helvetica, sans-serif;" +
                "font-size: 7px;" +
                "}*/" +
                "body p {" +
                "margin-left:3px;" +
                "margin-top:2px;" +
                "margin-bottom:4px;" +
                "text-align: left;" +
                "font-family: Arial, Helvetica, sans-serif;" +
                "font-size: 10px;" +
                "font-weight: bold;" +
                "}" +
                "body table tr td {" +
                "font-family: Arial, Helvetica, sans-serif;" +
                "font-size: 8px;" +
                "text-align: left;" +
                "text-transform: none;" +
                "vertical-align: top;" +
                "border: 1px solid #000;" +
                "border-top-width: 1px;" +
                "border-right-width: 1px;" +
                "border-bottom-width: 1px;" +
                "border-left-width: 1px;" +
                "/*text-align: center;" +
                "text-align: left;" +
                "font-weight: normal;*/" +
                "}" +
                "body table tr td p {" +
                "text-transform: none;" +
                "font-family: Arial, Helvetica, sans-serif;" +
                "font-weight: normal;" +
                "}" +
                "body table tr td .MsoNormal b span {" +
                "font-size: 8px;" +
                "}" +
                "body table tr td[bgcolor='#D6D6D6']{" +
                "font-size:8pt;" +
                "vertical-align:middle;" +
                "text-align: center;" +
                "font-weight: bold;" +
                "}" +
                ".style1 {font-size: 10px;}" +
                ".check-with-label:checked + .label-for-check {" +
                "font-weight: bold !important;" +
                "}" +
                "div img {" +
                "max-height: 100%;" +
                "max-width: 100%;" +
                "position: absolute;" +
                "left: 0;" +
                "bottom: 0;" +
                "margin-bottom: 5px;" +
                "}" +
                ".titulo {" +
                "text-align: center;" +
                "margin-left: 150px;" +
                "font-family: Arial, Helvetica, sans-serif;" +
                "font-size: 10pt;" +
                "margin-bottom: 5px;" +
                "}" +
                ".tabelaCheckbox {" +
                "margin-top: 5px !important;" +
                "border: 0 !important;" +
                "white-space: nowrap;" +
                "}" +
                ".tabelaCheckbox td {" +
                "border: 0 !important;" +
                "font-size: 8pt !important;" +
                "}" +
                ".dado1 {" +
                "position:absolute;" +
                "margin-top: 30px;" +
                "margin-left: 15px;" +
                "font-weight: bold;" +
                "}" +
                ".dado2 {" +
                "position:absolute;" +
                "margin-top: 15px;" +
                "font-weight: bold;" +
                "}" +
                "#tabelaConcomitante tr {" +
                "font-weight: bold;" +
                "}" +
                "#tabelaConcomitante tr:first-child {" +
                "font-weight: normal;" +
                "}" +
                "#tabelaConcomitante td {" +
                "vertical-align: middle;" +
                "padding:3px;" +
                "}" +
                "#tabelaConcomitante span {" +
                "font-size: 8pt !important;" +
                "}" +
                "td {" +
                "border-top:0 !important;" +
                "border-right:0 !important;" +
                "}" +
                "</style>" +
                "</head>" +
                "<body class=\"Farmaco\" dir=\"ltr\" lang=\"pt-BR\" style=\"position:absolute\">" +
                "  <div style=\"width:700px;position:relative;\">" +
                "    <img border=\"0\" alt=\"Image\" src=\"https://superarx.custhelp.com/euf/assets/images/Logo.png\" />" +
                "    <div class=\"titulo\">" +
                "      SISTEMA DE FARMACOVIGILÂNCIA" +
                "      <br/>" +
                "      NOTIFICAÇÃO EVENTO ADVERSO A MEDICAMENTO" +
                "    </div>" +
                "  </div>" +
                "<table width=\"700\" height=\"42\" border=\"0\" align=\"center\" cellspacing=\"0\" style=\"border-top: 1px solid #000 !important;\">" +
                "      <tr>" +
                "    <td height=\"14\" colspan=\"2\" nowrap bgcolor=\"#D6D6D6\">SAC SUPERA: 0800-708 18 18</td>" +
                "      <td colspan=\"2\" nowrap bgcolor=\"#000000\" style=\"text-align: center; color: #FFF;\">Preenchimento Exclusivo pelo depto. de Farmacovigilância</td>" +
                "    </tr>" +
                "    <tr>" +
                "      <td height=\"14\" colspan=\"2\" nowrap><p><span>PV Colaborador:  <strong>" + matriculaFuncionario + "</strong></span></p></td>" +
                "      <td width=\"177\" rowspan=\"2\" nowrap><p style=\"text-align: center\">Código relato</p><p>&nbsp;</p>" +
                "      </td>" +
                "      <td width=\"189\" rowspan=\"2\" nowrap style=\"text-align: center\"><p style=\"text-align: center\">Data de recebimento do relato</p></td>" +
                "    </tr>" +
                "    <tr>" +
                "      <td width=\"156\" height=\"14\" nowrap><p><span>Setor:  <strong>" + setor + "</strong></span></p></td>" +
                "      <td width=\"170\" nowrap nowrap><p><span>Tel:  <strong>" + telefoneComercial + "</strong></span></p></td>" +
                "    </tr>" +
                "  </table>";
            string html2 =
                "  <table width=\"700\" height=\"30\" border=\"0\" align=\"center\" cellspacing=\"0\">" +
                "    <tr>" +
                "      <td height=\"14\" bgcolor=\"#D6D6D6\">I. INFORMAÇÕES SOBRE O EVENTO ADVERSO</td></tr>" +
                "    <tr>" +
                "      <td height=\"16\" style=\"text-align: center\"><p style=\"text-align: center; font-weight: bold;\">FORMULÁRIO  MODELO CIOMS</p></td>" +
                "    </tr>" +
                "  </table>" +
                "  <table width=\"700\" height=\"277\" border=\"0\" align=\"center\" cellpadding=\"0\" cellspacing=\"0\">" +
                "    <tr>" +
                "      <td width=\"75\" height=\"72\" rowspan=\"3\" style=\"text-align: left\"><p class=\"style1\"><span ><strong>1. </strong>PACIENTE<br/>(iniciais)</span><br/><span class=\"dado1\">" + iniciaisPaciente + "</span></p></td>" +
                "      <td width=\"90\" rowspan=\"3\"><p class=\"style1\"><span><strong>2. </strong>DATA DE NASC</span><br/><br/><span class=\"dado1\">" + dataNasc + "</span></p></td>" +
                "      <td width=\"50\" rowspan=\"3\" nowrap><p class=\"style1\"><span><strong>3. </strong>IDADE</span><br/><br/><span class=\"dado1\">" + idade + "</span></p></td>" +
                "      <td width=\"67\" height=\"22\" nowrap><p class=\"style1\" style=\"margin-bottom: 0px;\"><span><strong>4. </strong>OUTROS</span><br/><br/><br/><span>Peso: <strong>" + peso + "</strong></span></p></td>" +
                "      <td width=\"125\" rowspan=\"3\" nowrap>" +
                "        <table class=\"tabelaCheckbox\">" +
                "          <tr><p><span><strong>5. </strong>SEXO</span></p></tr>" +
                "            <tr>" +
                "              <td><input type=\"checkbox\" name=\"Masculino\" id=\"Masculino\" class=\"check-with-label\" disabled=\"disabled\"/><label for=\"Masculino\" class=\"label-for-check\">Masc.</label></td>" +
                "              <td><input type=\"checkbox\" name=\"Feminino\" id=\"Feminino\" class=\"check-with-label\" disabled=\"disabled\"/><label for=\"Feminino\" class=\"label-for-check\">Fem.</label></td>" +
                "            </tr>" +
                "              <td><input type=\"checkbox\" name=\"Desc\" id=\"Desc\" class=\"check-with-label\" disabled=\"disabled\"/><label for=\"Desc.\" class=\"label-for-check\">Desc</label></td>" +
                "              <td><input type=\"checkbox\" name=\"gravidez\" id=\"gravidez\" class=\"check-with-label\" disabled=\"disabled\"/><label for=\"gravidez\" class=\"label-for-check\">Gravidez</label></td>" +
                "            <tr>" +
                "            </tr>" +
                "        </table>" +
                "      </td>" +
                "      <td width=\"105\" rowspan=\"3\"><p class=\"style1\"><span><strong>6. </strong>DATA DO EVENTO</span><br/><br/><span>Início: <strong>" + dataInicioEvento + "</strong><br/><br/><br/><span>Final: <strong>" + dataFinalEvento + "</strong></span></p></td>" +
                "      <td width=\"139\" colspan=\"2\" rowspan=\"4\" style=\"text-transform: none; text-align: left\">" +
                "      <p class=\"style1\">" +
                "          <span  style=\"font-weight: bold\">7.</span>" +
                "          <span >EM CASO DE EVENTO GRAVE, ASSINALAR CONSEQUÊNCIA(S):</span>" +
                "          <br/><br/>" +
                "          <input type=\"checkbox\" name=\"Óbito\" id=\"Óbito\" class=\"check-with-label\" disabled=\"disabled\"/>" +
                "          <label for=\"Óbito\" class=\"label-for-check\">Óbito</label>" +
                "          <br/>" +
                "          <input type=\"checkbox\" name=\"Hospitalização Necessária ou Prolongada\" id= \"Hospitalização Necessária ou Prolongada\" class=\"check-with-label\" disabled=\"disabled\"/>" +
                "          <label for=\"Hospitalização Necessária ou Prolongada\" class=\"label-for-check\">Hospitalização necessária ou prolongada</label>" +
                "          <br/>" +
                "          <input type=\"checkbox\" name=\"Incapacidade Persistente ou Permanente\" id=\"Incapacidade Persistente ou Permanente\" class=\"check-with-label\"\" disabled=\"disabled\"/>" +
                "          <label for=\"Incapacidade Persistente ou Permanente\" class=\"label-for-check\">Incapacidade persitente ou permanente</label>" +
                "          <br/>" +
                "          <input type=\"checkbox\" name=\"Ameaça à vida\" id=\"Ameaça à vida\" class=\"check-with-label\" disabled=\"disabled\"/>" +
                "          <label for=\"Ameaça à vida\" class=\"label-for-check\">Ameaça à vida</label>" +
                "          <br/>" +
                "          <input type=\"checkbox\" name=\"Anomalias Congênitas\" id=\"Anomalias Congênitas\" class=\"check-with-label\" disabled=\"disabled\"/>" +
                "          <label for=\"Anomalias Congênitas\" class=\"label-for-check\">Anomalias congenitas</label>" +
                "          <br/>" +
                "          <input type=\"checkbox\" name=\"Clinicamente Relevante\" id=\"Clinicamente Relevante\" class=\"check-with-label\" disabled=\"disabled\"/>" +
                "          <label for=\"Clinicamente Relevante\" class=\"label-for-check\">Clinicamente relevante</label>" +
                "          <br/>" +
                "          <input type=\"checkbox\" name=\"Não Grave\" id=\"Não Grave\" class=\"check-with-label\" disabled=\"disabled\"/>" +
                "          <label for=\"Não Grave\" class=\"label-for-check\">Não grave</label>" +
                "        </p>" +
                "      </td>" +
                "    </tr>" +
                "    <tr style=\"height: 0px;\"></tr>" +
                "    <tr>" +
                "      <td height=\"12\" nowrap><p class=\"style1\"><span>Alt: <strong>" + altura + "</strong></span></p></td>" +
                "    </tr>" +
                "    <tr>" +
                "      <td colspan=\"6\"><p class=\"style1\"><span><strong>8. </strong>DESCRIÇÃO DO EVENTO (incluindo dados laboratoriais relevantes)</span><br/><br/><span><strong>" + descricaoEvento + "</strong></span></p></td>" +
                "    </tr>" +
                "    <tr>" +
                "      <td colspan=\"4\">" +
                "        <p class=\"style1\"><span><strong>9. </strong>RELAÇÃO CAUSAL ENTRE O USO DO MEDICAMENTO E  O APARECIMENTO DO EVENTO ADVERSO (segundo o relator)</span></p>" +
                "        <table class=\"tabelaCheckbox\">" +
                "          <tr>" +
                "            <td><input type=\"checkbox\" name=\"Certa\" id=\"Certa\" class=\"check-with-label\" disabled=\"disabled\"/><label for=\"Certa\" class=\"label-for-check\">Certa</label></td>" +
                "            <td><input type=\"checkbox\" name=\"Possível\" id=\"Possível\" class=\"check-with-label\" disabled=\"disabled\"/><label for=\"Possível\" class=\"label-for-check\">Possível</label></td>" +
                "            <td><input type=\"checkbox\" name=\"Condicional\" id=\"Condicional\" class=\"check-with-label\"  disabled=\"disabled\"/><label for=\"Condicional\" class=\"label-for-check\">Condicional</label></td>" +
                "          </tr>" +
                "          <tr>" +
                "            <td><input type=\"checkbox\" name=\"Provável\" id=\"Provável\" class=\"check-with-label\" disabled=\"disabled\"/><label for=\"Provável\" class=\"label-for-check\">Provável</label></td>" +
                "            <td><input type=\"checkbox\" name=\"Improvável\" id=\"Improvável\" class=\"check-with-label\" disabled=\"disabled\"/><label for=\"Improvável\" class=\"label-for-check\">Improvável</label></td>" +
                "            <td><input type=\"checkbox\" name=\"Não Relacionada\" id=\"Não Relacionada\" class=\"check-with-label\" disabled=\"disabled\"/><label for=\"Não Relacionada\" class=\"label-for-check\">Não relacionada</label></td>" +
                "          </tr>" +
                "        </table>" +
                "      </td>" +
                "      <td colspan=\"4\">" +
                "          <p class=\"style1\"><span><strong>10. </strong>EVOLUÇÃO DO PACIENTE</span></p>" +
                "        <table class=\"tabelaCheckbox\" nowrap>" +
                "          <tr>" +
                "            <td><input type=\"checkbox\" name=\"Recuperado\" id=\"Recuperado\" class=\"check-with-label\" disabled=\"disabled\"/><label for=\"Recuperado\" class=\"label-for-check\">Recuperado</label></td>" +
                "            <td><input type=\"checkbox\" name=\"Desconhecido\" id=\"Desconhecido\" class=\"check-with-label\" disabled=\"disabled\"/><label for=\"Desconhecido\" class=\"label-for-check\">Desconhecido</label></td>" +
                "          </tr>" +
                "          <tr>" +
                "            <td><input type=\"checkbox\" name=\"Recuperado com sequela\" id=\"Recuperado com sequela\" class=\"check-with-label\"  disabled=\"disabled\"/><label for=\"Recuperado com sequela\" class=\"label-for-check\">Recuperado com sequela</label></td>" +
                "            <td><input type=\"checkbox\" name=\"Óbito\" id=\"Óbito\" class=\"check-with-label\" disabled=\"disabled\"/><label for=\"Óbito\" class=\"label-for-check\">Óbito. Especificar causa:</label></td>" +
                "          </tr>" +
                "          <tr>" +
                "            <td><input type=\"checkbox\" name=\"Não recuperado\" id=\"Não recuperado\" class=\"check-with-label\" disabled=\"disabled\"/><label for=\"Não recuperado\" class=\"label-for-check\">Não recuperado</label></td>" +
                "            <td><span style=\"margin-left: 20px !important; font-weight:normal;\">" + especObito + "</span></td>" +
                "          </tr>" +
                "        </table>" +
                "      </td>" +
                "    </tr>" +
                "  </table>";
            string html3 =
                "  <table width=\"700\" height=\"180\" border=\"0\" align=\"center\" cellspacing=\"0\">" +
                "    <tr><td colspan=\"100\" height=\"14\" bgcolor=\"#D6D6D6\">II.  INFORMAÇÕES SOBRE O MEDICAMENTO SUSPEITO</td></tr>" +
                "    <tr>" +
                "      <td colspan=\"53\" height=\"40\"><p class=\"style1\"><span ><strong>11. </strong>NOME DO MEDICAMENTO SUSPEITO</span><br/><span class=\"dado2\">" + nomeMedicamentoSuspeito + "</span></p></td>" +
                "      <td colspan=\"15\"><p class=\"style1\"><span ><strong>12. </strong>LOTE</span><br/><span class=\"dado2\">" + lote + "</span></p></td>" +
                "      <td colspan=\"32\" rowspan=\"2\">" +
                "        <p class=\"style1\">" +
                "          <span ><strong>13. </strong>O EVENTO ADVERSO CESSOU COM A INTERRUPÇÃO DO MEDICAMENTO?</span>" +
                "          <br/>" +
                "          <input type=\"checkbox\" name=\"Sim\" id=\"Sim\" class=\"check-with-label\" disabled=\"disabled\"/><label for=\"Sim\" class=\"label-for-check\">Sim. Especifique: " + especificacaoInterrupcao + "</label>" +
                "          <br/>" +
                "          <input type=\"checkbox\" name=\"Não\" id=\"Não\" class=\"check-with-label\" disabled=\"disabled\"/><label for=\"Não\" class=\"label-for-check\">Não</label>" +
                "          <br/>" +
                "          <input type=\"checkbox\" name=\"Não se aplica\" id=\"Não se aplica\" class=\"check-with-label\"  disabled=\"disabled\"/><label for=\"Não se aplica\" class=\"label-for-check\">Não se aplica</label>" +
                "        </p>" +
                "      </td>" +
                "      </tr>" +
                "    <tr>" +
                "      <td colspan=\"18\" height=\"40\"><p class=\"style1\"><span ><strong>14. </strong>VIA DE ADM.</span><br/><span class=\"dado2\">" + viaAdministracao + "</span></p></td>" +
                "      <td colspan=\"30\"><p class=\"style1\"><span ><strong>15. </strong>DOSE DIÁRIA</span><br/><span class=\"dado2\">" + doseDiaria + "</span></p></td>" +
                "      <td colspan=\"20\"><p class=\"style1\"><span ><strong>16. </strong>INDICAÇÃO DE USO</span><br/><br/><span class=\"dado2\"style=\"position: initial;\">" + indicacaoDeUso + "</span></p></td>" +
                "    </tr>" +
                "    <tr>" +
                "      <td colspan=\"23\" height=\"40\"><p class=\"style1\"><span ><strong>17. </strong>PERÍODO DO TRATAMENTO</span><br/><br/><span>Início: <strong>" + dataInicialTratamento + "</strong></span><br/><br/><span>Final: <strong>" + dataFinalTratamento + "</strong></span></p></td>" +
                "      <td colspan=\"45\" height=\"40\"><p class=\"style1\"><span ><strong>18. </strong>DURAÇÃO DO TRATAMENTO ATÉ INÍCIO DO EVENTO</span><br/><span class=\"dado2\">" + duracaoTratamento + "</span></p></td>" +
                "      <td colspan=\"32\">" +
                "        <p class=\"style1\">" +
                "          <span ><strong>19. </strong>HOUVE REINTRODUÇÃO DO PRODUTO?</span>" +
                "          <br/>" +
                "          <input type=\"checkbox\" name=\"Sim2\" id=\"Sim2\" class=\"check-with-label\" disabled=\"disabled\"/><label for=\"Sim2\" class=\"label-for-check\">Sim. Especifique: " + especificacaoReintroducao + "</label>" +
                "          <br/>" +
                "          <input type=\"checkbox\" name=\"Não2\" id=\"Não2\" class=\"check-with-label\" disabled=\"disabled\"/><label for=\"Não2\" class=\"label-for-check\">Não</label>" +
                "          <br/>" +
                "          <input type=\"checkbox\" name=\"Não se aplica2\" id=\"Não se aplica2\" class=\"check-with-label\" disabled=\"disabled\"/><label for=\"Não se aplica2\" class=\"label-for-check\">Não se aplica </label>" +
                "        </p>" +
                "      </td>" +
                "    </tr>" +
                "  </table>";
            string html4 =
                "  <table width=\"700\" border=\"0\" align=\"center\" cellspacing=\"0\">" +
                "    <tr>" +
                "      <td height=\"14\" bgcolor=\"#D6D6D6\">III.  MEDICAMENTO(S) CONCOMITANTE(S) E HISTÓRICO CLÍNICO</td>" +
                "    </tr>" +
                "    <tr>" +
                "      <td height=\"14\"><p class=\"style1\"><span><strong>20. </strong>MEDICAMENTO(S) CONCOMITANTE(S) (excluindo os utilizados para tratar o  evento adverso)</span></p></td>" +
                "    </tr>" +
                "  </table>" +
                "<table id=\"tabelaConcomitante\" width=\"700\" cellspacing=\"0\">" +
                "  <tr>" +
                "    <td colspan=\"85\"><span>Nome comercial e genérico</span></td>" +
                "    <td colspan=\"20\"><span>Dose</span></td>" +
                "    <td colspan=\"20\"><span>Via adm.</span></td>" +
                "    <td colspan=\"30\"><span>Indicação</span></td>" +
                "    <td colspan=\"30\"><span>Data  de início do tratamento</span></td>" +
                "    <td colspan=\"30\"><span>Data  de término do tratamento</span></td>" +
                "    <td colspan=\"25\"><span>Data  da reintrodução</span></td>" +
                "  </tr>" +
                usoConcomitante +
                "  <table width=\"700\" border=\"0\" align=\"center\" cellspacing=\"0\" cellpadding=\"0\">" +
                "    <tr>" +
                "      <td colspan=\"1\" style=\"border-top:0;\">" +
                "        <p class=\"style1\"><span><strong>21. </strong>DADOS CLÍNICOS RELEVANTES (doenças e seus tratamentos, reações a  outros medicamentos, gravidez e idade gestacional etc.)</span></p>" +
                "        <p class=\"style1\" colspan=\"1\" style=\"border-top:0;\"><span><strong>" + dadosClinicosRelevantes + "</strong></span></p>" +
                "      </td>" +
                "    </tr>" +
                "  </table>";
            string html5 =
                "  <table width=\"700\" border=\"0\" align=\"center\" cellspacing=\"0\">" +
                "    <tr>" +
                "      <td height=\"14\" bgcolor=\"#D6D6D6\">IV.  OUTRAS INFORMAÇÕES</td>" +
                "    </tr>" +
                "  </table>" +
                "  <table width=\"700\" height=\"157\" border=\"0\" align=\"center\" cellspacing=\"0\">" +
                "    <tr>" +
                "      <td width=\"357\" rowspan=\"6\"><p class=\"style1\"><br/><br/><br/><strong>22.</strong>SAC SUPERA.</p><p class=\"style1\" style=\"line-height: 15px;\">Av.  da Nações Unidas, 22532 - Bl. 1 - Vila Almeida<br/>Cep: 04795-000 - São  Paulo - SP<br/>supera.atende@superarx.com.br</p></td> " +
                "      <td colspan=\"3\" nowrap style=\"vertical-align:middle;\"><p class=\"style1\"><span><strong>25. </strong>IDENTIFICAÇÃO DO RELATOR</span></p></td>" +
                "    </tr>" +
                "    <tr>" +
                "      <td colspan=\"3\" nowrap><p>Nome:  <strong>" + nomeContato + "</strong></p></td>" +
                "    </tr>" +
                "    <tr>" +
                "      <td colspan=\"3\" nowrap><p>Inscrição  profissional: <span style=\"font-weight: bold\">" + inscricaoProfissional + "</span></p></td>" +
                "    </tr>" +
                "    <tr>" +
                "      <td height=\"14\" colspan=\"3\" nowrap><p>Endereço:  <span style=\"font-weight: bold\">" + ruaContato + "</span></p></td>" +
                "    </tr>" +
                "    <tr>" +
                "      <td width=\"129\" height=\"14\" nowrap><p>Cidade:  <span style=\"font-weight: bold\">" + cidadeContato + "</span></p></td>" +
                "      <td width=\"128\" nowrap><p>CEP:  <span style=\"font-weight: bold\">" + cepContato + "</span></p></td>" +
                "      <td width=\"78\" nowrap><p>UF:  <span style=\"font-weight: bold\">" + estadoContato + "</span></p></td>" +
                "    </tr>" +
                "    <tr>" +
                "      <td height=\"14\" nowrap><p>Tel.:  <span style=\"font-weight: bold\">" + telefoneAssistente + "</span></p></td>" +
                "      <td colspan=\"2\" nowrap><p>Cel.:  <span style=\"font-weight: bold\">" + celularContato + "</span></p></td>" +
                "    </tr>" +
                "    <tr>" +
                "      <td height=\"34\" nowrap>" +
                "          <p class=\"style1\"><span><strong>23. </strong>FONTE DO RELATO</span></p>" +
                "          <input type=\"checkbox\" name=\"Sim3\" id=\"Sim3\" class=\"check-with-label\" disabled=\"disabled\"/>" +
                "          <label for=\"Sim3\" class=\"label-for-check\" style=\"font-size: 8pt;\">Profissional saúde. Especialidade: " + especialidade + "</label>" +
                "      </td>" +
                "      <td colspan=\"3\" nowrap><p>E-mail:  <span style=\"font-weight: bold\">" + emailContato + "</span></p></td>" +
                "    </tr>" +
                "    <tr>" +
                "      <td height=\"30\" colspan=\"3\" nowrap><p class=\"style1\" style=\"font-size:9px;\"><span><strong>24. </strong>ASSINATURA E CARIMBO DO RELATOR</span></p></td>" +
                "      <td style=\"text-align: center\"><p class=\"text MsoNormal\" style=\"text-align: left; font-family: Arial, Helvetica, sans-serif;\">DATA</p>" +
                "        <p class=\"style1\" colspan=\"1\" style=\"border-top:0;\"><span><strong>" + dataFarmaco + "</strong></span></p></td>" +
                "    </tr>" +
                "  </table>" +
                "  <div style=\"width:700px\">" +
                "    <p style=\"text-align: center; margin:3px;\">POR  FAVOR, REMETER ESTE FORMULÁRIO PARA FARMACOVIGILÂNCIA SUPERA VIA  CORREIO, E-MAIL OU REPRESENTANTE</p>" +
                "    <p style=\"text-align: center; font-weight: normal;font-size: 8px; margin: 0px;\">UTILIZE  O VERSO PARA INFORMAÇÕES COMPLEMENTARES</p>" +
                "  </div>" +
                "    <div style=\" height: 80px;\"></div>";
            string html6 =
                "    <p class=\"MsoNormal\" style=\"TEXT-ALIGN: center; MARGIN: 0cm 86.45pt 0pt 73.8pt\" align=\"center\">" +
                "      <b><span style=\"FONT-SIZE: 10pt; COLOR: #181717; LINE-HEIGHT: 100%\">NOTIFICAÇÃO DE SUSPEITA DE EVENTO</span></b>" +
                "    </p>" +
                "    <p class=\"MsoNormal\" style=\"TEXT-ALIGN: center; MARGIN: 3px 86.45pt 5px 73.8pt;\" align=\"center\">" +
                "      <b><span style=\"FONT-SIZE: 10pt; COLOR: #181717; LINE-HEIGHT: 100%\">INFORMAÇÕES COMPLEMENTARES</span></b>" +
                "    </p>" +
                "  <div style=\"width:700px\" align=\"center\">" +
                "    <table class=\"TableGrid\" style=\"BORDER-COLLAPSE: collapse; border-top: 1px solid #000 !important;\" cellspacing=\"0\" cellpadding=\"0\">" +
                "      <tbody>" +
                "        <tr style=\"display:table-row-group\">" +
                "          <td style=\"BORDER-TOP: #181717 1pt solid; BORDER-RIGHT: #181717 1pt solid; WIDTH: 537.6pt; BORDER-BOTTOM: #181717 1pt solid; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 5.75pt; BORDER-LEFT: #181717 1pt solid; PADDING-RIGHT: 5.75pt; font-size: 12px; text-align: justify; font-weight: bold;\" valign=\"top\" width=\"788\">" +
                "            <p class=\"MsoNormal\" style=\"margin:20px; margin-bottom:75px; font-weight:bold; text-transform:uppercase;\">" + informacoesComplementares + "</p>" +
                "          </td>" +
                "        </tr>" +
                "        <tr style=\"HEIGHT: 24.5pt\">" +
                "          <td style=\"vertical-align:middle; BORDER-TOP: medium none; HEIGHT: 24.5pt; BORDER-RIGHT: #181717 1pt solid; WIDTH: 537.6pt; BORDER-BOTTOM: #181717 1pt solid; PADDING-BOTTOM: 0cm; PADDING-TOP: 0cm; PADDING-LEFT: 5.75pt; BORDER-LEFT: #181717 1pt solid; PADDING-RIGHT: 5.75pt\" width=\"720\">" +
                "            <p class=\"MsoNormal\" style=\"TEXT-ALIGN: center; MARGIN: 0cm 0.05pt 0pt 0cm\" align=\"center\">" +
                "              <span style=\"FONT-SIZE: 10pt; COLOR: #181717; LINE-HEIGHT: 100%\">As informações relatadas neste formulário são confidenciais e de uso exclusivo da área de Farmacovigilância Supera.</span>" +
                "            </p>" +
                "          </td>" +
                "        </tr>" +
                "      </tbody>" +
                "    </table>" +
                "  </div>" +
                "</body>" +
                "</html>";

            html1 = marcarCheckboxes(html1);
            html2 = marcarCheckboxes(html2);
            html3 = marcarCheckboxes(html3);
            html4 = marcarCheckboxes(html4);
            html5 = marcarCheckboxes(html5);
            html6 = marcarCheckboxes(html6);

            gravarArquivoHTML(inc.ID, html1, html2, html3, html4, html5, html6);

        }

        public string marcarCheckboxes(string html) {
            html = Regex.Replace(html, "(<input type=\"checkbox\" name=\"" + sexo + "\")", "$1 checked");
            html = Regex.Replace(html, "(<input type=\"checkbox\" name=\"" + gravidez + "\")", "$1 checked");
            html = Regex.Replace(html, "(<input type=\"checkbox\" name=\"" + consequencia + "\")", "$1 checked");
            html = Regex.Replace(html, "(<input type=\"checkbox\" name=\"" + relacaoMedicamento + "\")", "$1 checked");
            html = Regex.Replace(html, "(<input type=\"checkbox\" name=\"" + evolucaoDoPaciente + "\")", "$1 checked");
            html = Regex.Replace(html, "(<input type=\"checkbox\" name=\"" + eventoCessouComInterrupcao + "\")", "$1 checked");
            html = Regex.Replace(html, "(<input type=\"checkbox\" name=\"" + houveReintroducao + "2\")", "$1 checked");
            html = Regex.Replace(html, "(<input type=\"checkbox\" name=\"" + profissionalSaude + "3\")", "$1 checked");
            return html;
        }


        public static void gravarArquivoHTML(int idIncident, string html1, string html2, string html3, string html4, string html5, string html6) {
            string nomeArquivo = escolherNomeArquivo(idIncident);
            if(nomeArquivo != null) {
                StreamWriter outputFile = new StreamWriter(nomeArquivo);
                outputFile.WriteLine(html1);
                outputFile.WriteLine(html2);
                outputFile.WriteLine(html3);
                outputFile.WriteLine(html4);
                outputFile.WriteLine(html5);
                outputFile.WriteLine(html6);
                outputFile.Close();
            }
        }

        static string escolherNomeArquivo(int idIncident) {
            Stream myStream;
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.FileName = "Farmacovigilância Ocorrência " + idIncident;
            saveFileDialog1.Filter = "Arquivo HTML|*.html";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;
            try {
                if(saveFileDialog1.ShowDialog() == DialogResult.OK) {
                    if((myStream = saveFileDialog1.OpenFile()) != null) {
                        myStream.Close();
                        return saveFileDialog1.FileName;
                    }
                }
            }
            catch(ThreadStateException t) {
                MessageBox.Show("Erro ao salvar arquivo.");
                t.Message.ToString();
            }
            return null;
        }

        public string pegarValorCustomFieldPorNome(string nomeCampo) {
            IIncident inc = (IIncident) RecordContext.GetWorkspaceRecord(RightNow.AddIns.Common.WorkspaceRecordType.Incident);

            List<ICustomAttribute> i = inc.CustomAttributes.ToList();

            foreach(ICustomAttribute icustomAttribute in i) {
                IGenericField g = icustomAttribute.GenericField;
                string a = g.Name;
                if(a.Equals(nomeCampo)) {
                    return g.DataValue.Value.ToString();
                }
            }
            return "";
        }

        public string pegarValorCustomFieldPorIDString(int id, IIncident inc) {
            foreach(ICfVal icfval in inc.CustomField) {
                if(icfval.CfId == id) {
                    if(icfval.ValStr != null) {
                        return icfval.ValStr;
                    }
                }
            }
            return "";
        }

        public string pegarValorCustomFieldPorIDStringContato(int id, IContact contato) {
            foreach(ICfVal icfval in contato.CustomField) {
                if(icfval.CfId == id) {
                    if(icfval.ValStr != null) {
                        return icfval.ValStr;
                    }
                }
            }
            return "";
        }

        public string pegarValorCustomFieldPorIDInt(int id, IIncident inc) {
            foreach(ICfVal icfval in inc.CustomField) {
                if(icfval.CfId == id) {
                    if(icfval.ValInt != null) {
                        return icfval.ValInt.ToString();
                    }
                }
            }
            return "";
        }

        public DateTime pegarValorCustomFieldPorIDDate(int id, IIncident inc) {
            foreach(ICfVal icfval in inc.CustomField) {
                if(icfval.CfId == id) {
                    if(icfval.ValDate != null) {
                        return (DateTime) icfval.ValDate;
                    }
                }
            }
            return default(DateTime);
        }

    }

    [AddIn("Workspace Ribbon Button AddIn", Version = "1.0.0.0")]
    public class WorkspaceRibbonButtonFactory : IWorkspaceRibbonButtonFactory {
        #region IWorkspaceRibbonButtonFactory Members
        public IWorkspaceRibbonButton CreateControl(bool inDesignMode, IRecordContext RecordContext) {
            return new WorkspaceRibbonAddIn(inDesignMode, RecordContext);
        }

        public System.Drawing.Image Image32 {
            get {
                return Properties.Resources._32x32;
            }
        }
        #endregion

        #region IFactoryBase Members
        public System.Drawing.Image Image16 {
            get {
                return Properties.Resources._16x16;
            }
        }

        public string Text {
            get {
                return "Relatório de Farmacovigilância";
            }
        }

        public string Tooltip {
            get {
                return "Esta função gera um relatório referente aos dados de Farmacovigilância deste Incidente.";
            }
        }
        #endregion

        #region IAddInBase Members
        public bool Initialize(IGlobalContext GlobalContext) {
            return true;
        }
        #endregion
    }

}