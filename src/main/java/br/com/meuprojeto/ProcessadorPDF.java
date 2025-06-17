package br.com.meuprojeto;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ProcessadorPDF {

    public static void main(String[] args) {
        String caminhoArquivoPDF = "Arquivo de entrada PDF"; 
        String caminhoPastaSaidaExcel = "Caminho que irá o arquivo de saída";
        String nomeArquivoExcel = "Nome do Arquivo de Saída";
        String caminhoCompletoArquivoExcel = caminhoPastaSaidaExcel + nomeArquivoExcel;

        try {
            List<List<String>> dadosTabela = extrairDadosTabelaDoPDF(caminhoArquivoPDF);
            if (!dadosTabela.isEmpty()) {
                File pastaSaida = new File(caminhoPastaSaidaExcel);
                if (!pastaSaida.exists()) {
                    pastaSaida.mkdirs();
                }
                escreverParaExcel(dadosTabela, caminhoCompletoArquivoExcel);
                System.out.println("Missão concluída com sucesso! Excel salvo em: " + caminhoCompletoArquivoExcel);
            } else {
                System.out.println("Nenhum dado foi extraído do PDF. Verifique a lógica de extração.");
            }
        } catch (IOException e) {
            System.err.println("Ocorreu um erro durante a conversão: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private static List<List<String>> extrairDadosTabelaDoPDF(String caminhoArquivoPDF) throws IOException {
        List<List<String>> dadosTabela = new ArrayList<>();
        try (PDDocument documento = PDDocument.load(new FileInputStream(caminhoArquivoPDF))) {
            PDFTextStripper extratorTexto = new PDFTextStripper();
            String textoCompleto = extratorTexto.getText(documento);

            String marcadorInicioHeaderComQuebra = "Processo\nAção\nÓrgão\nAssunto\nDistribuição\nTipo\nJulgador\nParticipação";
            String marcadorInicioHeaderAchatado = "Processo Ação Órgão Julgador Assunto Distribuição Tipo Participação";

            int indiceInicio = textoCompleto.indexOf(marcadorInicioHeaderComQuebra);
            if (indiceInicio == -1) {
                indiceInicio = textoCompleto.indexOf(marcadorInicioHeaderAchatado);
                if (indiceInicio != -1) {
                    System.out.println("Aviso: Cabeçalho com quebras de linha não encontrado, usando versão 'achatada' como início.");
                }
            }

            if (indiceInicio == -1) {
                System.out.println("Marcador de início da tabela ('" + marcadorInicioHeaderComQuebra + "' ou '" + marcadorInicioHeaderAchatado + "') não encontrado.");
                return dadosTabela;
            }

            int indiceFim = textoCompleto.indexOf("Esta certidão abrange as ações das varas de família", indiceInicio);
            if (indiceFim == -1) {
                indiceFim = textoCompleto.length();
            }

            String secaoTabelaCompleta = textoCompleto.substring(indiceInicio, indiceFim);

            String dadosApenasString = secaoTabelaCompleta
                                          .replaceAll(Pattern.quote(marcadorInicioHeaderAchatado), "")
                                          .replaceAll(Pattern.quote(marcadorInicioHeaderComQuebra), "")
                                          .replaceAll("Comarca\\s+[A-Z\\s]+\\s*", "")
                                          .replaceAll("\\s{2,}", " ")
                                          .trim();

            List<String> cabecalhos = Arrays.asList(
                "Processo", "Ação", "Órgão Julgador", "Assunto", "Distribuição", "Tipo", "Participação"
            );
            dadosTabela.add(cabecalhos);

            Pattern padraoBlocoRegistro = Pattern.compile(
                "(\\d{7}-\\d{2}\\.\\d{4}\\.8\\.05\\.\\d{4}[\\s\\S]*?)(?=\\d{7}-\\d{2}\\.\\d{4}\\.8\\.05\\.\\d{4}|$)"
            );

            Matcher localizadorBloco = padraoBlocoRegistro.matcher(dadosApenasString);

            while (localizadorBloco.find()) {
                String blocoRegistroBruto = localizadorBloco.group(1).trim();
                List<String> linhaDados = parsearRegistroComplexo(blocoRegistroBruto);
                if (!linhaDados.isEmpty() && linhaDados.size() == cabecalhos.size()) {
                    dadosTabela.add(linhaDados);
                } else {
                    System.err.println("Aviso: Linha não parseada ou com número de colunas incorreto. Bloco: '" + blocoRegistroBruto + "'");
                }
            }
        }
        return dadosTabela;
    }

    private static List<String> parsearRegistroComplexo(String blocoRegistroBruto) {
        List<String> dadosLinha = new ArrayList<>();

        String blocoLimpo = blocoRegistroBruto.replaceAll("\"", "")
                                              .replaceAll("\r?\n", " ")
                                              .replaceAll("\\s+", " ")
                                              .replaceAll("├á", "á").replaceAll("├â", "â")
                                              .replaceAll("├ú", "ã").replaceAll("├ü", "Ã")
                                              .replaceAll("├®", "é").replaceAll("├¬", "ê")
                                              .replaceAll("├¡", "í")
                                              .replaceAll("├ó", "ó").replaceAll("├ô", "ô")
                                              .replaceAll("├õ", "õ")
                                              .replaceAll("├║", "ú")
                                              .replaceAll("├º", "ç").replaceAll("├ç", "Ç")
                                              .replaceAll("┬º", "º")
                                              .replaceAll("PODER JUDICIÁRIO.*?\\bTribunal de Justiça do Estado da Bahia\\b\\s*\\d+", "")
                                              .replaceAll("\\bProcesso\\s+Ação\\s+Órgão\\s+Julgador\\s+Assunto\\s+Distribuição\\s+Tipo\\s+Participação\\b", "")
                                              .replaceAll("\\bComarca\\s+[A-Z\\s]+\\b", "")
                                              .trim();

        System.out.println("DEBUG - Final Bloco Limpo para Parsing: '" + blocoLimpo + "'");

        Pattern padraoCampos = Pattern.compile(
<<<<<<< HEAD
            "^" +
            "(\\d{7}-\\d{2}\\.\\d{4}\\.8\\.05\\.\\d{4})" + // 1: Processo
            "\\s+(.+?)" + // 2: Ação 
            "\\s+((?:\\d{1,2}º\\s+VARA|VARA)(?:[^\\d]+?|.*?))" + // 3: Órgão Julgador
            "\\s+(.+?)" + // 4: Assunto
            "\\s+(\\d{2}/\\d{2}/\\d{4})" + // 5: Distribuição
            "\\s+PARTE\\s*(.*?)\\s*(ATIVA|PASSIVA)" + // 6: Tipo
            "(?:\\s*(.*))?" + // 7: Participação
            "$" // FIM
=======
            "^" + // Início
            "(\\d{7}-\\d{2}\\.\\d{4}\\.8\\.05\\.\\d{4})" + // 1: Processo 
            "\\s+(.+?)" + // 2: Ação 
            "\\s+((?:\\d{1,2}º\\s+VARA|VARA)(?:[^\\d]+?|.*?))" + // 3: Órgão Julgador 
            "\\s+(.+?)" + // 4: Assunto 
            "\\s+(\\d{2}/\\d{2}/\\d{4})" + // 5: Distribuição 
            "\\s+PARTE\\s*(.*?)\\s*(ATIVA|PASSIVA)" + // 6: Tipo 
            "(?:\\s*(.*))?" + // 7: Participação 
            "$" // Fim
>>>>>>> 429c430e96f3c8cfee066b176ce8af7f5c7d51ff
        );

        Matcher localizadorCampos = padraoCampos.matcher(blocoLimpo);

        if (localizadorCampos.find()) {
            String processo = localizadorCampos.group(1);
            String acao = localizadorCampos.group(2);
            String orgaoJulgador = localizadorCampos.group(3);
            String assunto = localizadorCampos.group(4);
            String distribuicao = localizadorCampos.group(5);
            String tipo = "PARTE " + localizadorCampos.group(7); 
<<<<<<< HEAD
            String participacao = localizadorCampos.group(8);


=======
            String participacao = localizadorCampos.group(8); 
            
            if (assunto != null && participacao != null) {
                if (assunto.equalsIgnoreCase("Acidente de") && participacao.toLowerCase().contains("trânsito")) {
                    assunto = assunto + " Trânsito";
                    participacao = participacao.replaceAll("\\bTrânsito\\b", "").trim();
                }
            }
            
            if (orgaoJulgador != null && assunto != null) {
                Pattern p = Pattern.compile("(.*?) (CONSUMO)$");
                Matcher m = p.matcher(orgaoJulgador.trim());
                if(m.matches()) {
                    orgaoJulgador = m.group(1).trim() + " " + m.group(2).trim(); 
                }
            }
            
            dadosLinha.add(processo);
            dadosLinha.add(acao);
            dadosLinha.add(orgaoJulgador.trim()); 
            dadosLinha.add(assunto.trim());       
            dadosLinha.add(distribuicao);
            dadosLinha.add(tipo.trim());          

                    orgaoJulgador = m.group(1).trim() + " " + m.group(2).trim();

                }
            }
            dadosLinha.add(processo);
            dadosLinha.add(acao);
            dadosLinha.add(orgaoJulgador.trim()); //remove espaços extras
            dadosLinha.add(assunto.trim());      
            dadosLinha.add(distribuicao);
            dadosLinha.add(tipo.trim());         
>>>>>>> 429c430e96f3c8cfee066b176ce8af7f5c7d51ff
            dadosLinha.add(participacao != null ? participacao.trim() : ""); 

        } else {
            System.err.println("Não foi possível parsear o registro completo com o padrão de campos: '" + blocoLimpo + "'");
        }

        return dadosLinha;
    }


    private static void escreverParaExcel(List<List<String>> dados, String caminhoArquivoExcel) throws IOException {
        try (Workbook planilha = new XSSFWorkbook()) {
            Sheet aba = planilha.createSheet("Dados Extraídos");

            int numeroLinha = 0;
            for (List<String> linhaDados : dados) {
                Row linha = aba.createRow(numeroLinha++);
                int numeroColuna = 0;
                for (String dadoCelula : linhaDados) {
                    Cell celula = linha.createCell(numeroColuna++);
                    celula.setCellValue(dadoCelula);
                }
            }

            if (!dados.isEmpty() && !dados.get(0).isEmpty()) {
                for (int i = 0; i < dados.get(0).size(); i++) {
                    aba.autoSizeColumn(i);
                }
            }

            try (FileOutputStream fluxoSaida = new FileOutputStream(caminhoArquivoExcel)) {
                planilha.write(fluxoSaida);
            }
        }
    }
}
