package br.eng.cvdejesuss.word_replace;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class WordReplace {
	private static final Logger logger = LogManager.getLogger(WordReplace.class);

	public static void main(String[] args) {
		// Lista de arquivos de entrada e saída
		String[] inputFilePaths = { 
				"D:\\Backup_Laptop_DELL\\Projeto_Automatizado\\Modelo\\1-Autuacao_designacao_compromisso_escrivao.docx",
				"D:\\Backup_Laptop_DELL\\Projeto_Automatizado\\Modelo\\2-Despacho_01-diligencias_necessarias.docx"
				// Adicione mais arquivos conforme necessário
		};

		String[] outputFilePaths = { 
				"D:\\Backup_Laptop_DELL\\Projeto_Automatizado\\Formatado\\1-Autuacao_designacao_compromisso_escrivao_new.docx",
				"D:\\Backup_Laptop_DELL\\Projeto_Automatizado\\Formatado\\2-Despacho_01-diligencias_necessarias_new.docx"
				// Correspondente aos arquivos de entrada
		};

		// Caminho para o arquivo de substituições
		String substitutionFilePath = "D:\\Backup_Laptop_DELL\\Projeto_Automatizado\\Modelo\\substituicoes.txt";

		// Mapeamento de palavras-chave para substituições
		Map<String, String> replacements = readReplacementsFromFile(substitutionFilePath);

		logger.info("Iniciando a substituição de palavras...");

		for (int i = 0; i < inputFilePaths.length; i++) {
			processFile(inputFilePaths[i], outputFilePaths[i], replacements);
		}

		logger.info("Substituição concluída com sucesso!");
	}

	private static Map<String, String> readReplacementsFromFile(String filePath) {
		Map<String, String> replacements = new HashMap<>();
		try (BufferedReader br = new BufferedReader(new FileReader(filePath))) {
			String line;
			while ((line = br.readLine()) != null) {
				String[] parts = line.split("=", 2);
				if (parts.length == 2) {
					replacements.put(parts[0].trim(), parts[1].trim());
				}
			}
		} catch (IOException e) {
			logger.error("Erro ao ler o arquivo de substituições: " + filePath, e);
		}
		return replacements;
	}

	private static void processFile(String inputFilePath, String outputFilePath, Map<String, String> replacements) {
		try {
			FileInputStream fis = new FileInputStream(inputFilePath);
			XWPFDocument document = new XWPFDocument(fis);

			for (XWPFParagraph paragraph : document.getParagraphs()) {
				for (XWPFRun run : paragraph.getRuns()) {
					String text = run.getText(0);
					if (text != null) {
						for (Map.Entry<String, String> entry : replacements.entrySet()) {
							String keyword = entry.getKey();
							String replacement = entry.getValue();
							if (text.contains(keyword)) {
								text = text.replace(keyword, replacement);
							}
						}
						run.setText(text, 0);
					}
				}
			}

			FileOutputStream fos = new FileOutputStream(outputFilePath);
			document.write(fos);
			fos.close();
			document.close();
			fis.close();

			logger.info("Arquivo processado com sucesso: " + inputFilePath);

		} catch (IOException e) {
			logger.error("Erro durante a substituição de palavras no arquivo: " + inputFilePath, e);
		}
	}
}
