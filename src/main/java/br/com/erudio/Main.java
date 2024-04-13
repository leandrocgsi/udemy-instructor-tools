package br.com.erudio;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.NumberFormat;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

public class Main {

    public static void main(String[] args) {
        // Caminho do arquivo HTML
        String filePath = "D:\\Udemy\\Cursos online - aprenda o que quiser, quando quiser _ Udemy.html";

        // Ler o conteúdo do arquivo HTML
        String html = "";
        try {
            File file = new File(filePath);
            Document doc = Jsoup.parse(file, "UTF-8");
            html = doc.toString();
        } catch (IOException e) {
            e.printStackTrace();
            return;
        }

        // Analisar o HTML com Jsoup
        Document doc = Jsoup.parse(html);

        // Selecionar todos os elementos com a classe "courses--card--pk1Fv"
        Elements courses = doc.select(".courses--card--pk1Fv");

        // Criar um novo arquivo Excel e escrever os dados
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Courses");

            // Escrever o cabeçalho
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Course Name");
            headerRow.createCell(1).setCellValue("Earnings This Month");
            headerRow.createCell(2).setCellValue("Total Earnings");

            // Formatar como moeda brasileira
            NumberFormat brazilianCurrencyFormat = NumberFormat.getCurrencyInstance(new Locale("pt", "BR"));

            // Escrever os dados
            int rowNum = 1;

            for (Element course : courses) {
                // Extrair o nome do curso, os ganhos do mês e o total recebido
                String courseName = course.select(".ud-heading-md").text().trim();
                String earningsThisMonth = "";
                String totalEarnings = "";

                Elements earningsElements = course.select(".ud-text-xl");
                if (earningsElements.size() >= 2) {
                    totalEarnings = earningsElements.get(1).text().replace("US$", "").trim();
                    Double earningsValue = parseCurrency(earningsElements.get(0).text().trim());
                    if (earningsValue != null && earningsValue != 0.0) {
                        earningsThisMonth = earningsElements.get(0).text().replace("US$", "").trim();
                    }
                }

                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(courseName);
                row.createCell(1).setCellValue(formatCurrency(earningsThisMonth, brazilianCurrencyFormat));
                row.createCell(2).setCellValue(formatCurrency(totalEarnings, brazilianCurrencyFormat));
            }

            // Remover a proteção da planilha
            sheet.protectSheet("");

            // Diretório de destino
            String outputDirectory = "E:\\Dropbox\\- Udemy\\UdemyTools";
            // Nome do arquivo
            String fileName = "courses.xlsx";
            // Caminho completo do arquivo de saída
            String outputFile = outputDirectory + File.separator + fileName;

            // Salvar o arquivo Excel
            try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
                workbook.write(fileOut);
                System.out.println("Arquivo Excel gerado com sucesso!");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static Double parseCurrency(String currencyText) {
        // Remover qualquer caractere não numérico, exceto o ponto decimal e a vírgula
        currencyText = currencyText.replaceAll("[^0-9.,]", "");
        // Verificar se a string está vazia
        if (currencyText.isEmpty()) {
            return null;
        }
        // Substituir vírgulas por pontos para garantir a interpretação correta como separador decimal
        currencyText = currencyText.replace(",", ".");
        // Tentar converter o texto para um double
        try {
            return Double.parseDouble(currencyText);
        } catch (NumberFormatException e) {
            return null; // Se a conversão falhar, retornar null
        }
    }

    private static String formatCurrency(String amount, NumberFormat format) {
        Double value = parseCurrency(amount);
        if (value != null) {
            return format.format(value);
        } else {
            return "";
        }
    }
}
