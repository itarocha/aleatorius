package br.itarocha.teste;

import lombok.SneakyThrows;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.Pattern;

public class TesteCsvToExcel {
    private static final String FOLDER = "/home/itamar/tmp/csv";
    private static final String EXTENSAO_CSV = ".csv";
    private static final String EXTENSAO_XLS = ".xlsx";
    private static final String CSV_DIVISOR = ";";
    private static final String SHEET_NAME = "DADOS";
    private static final String FORMATO_DATA_EXCEL = "dd/mm/yyyy";
    private static final String FORMATO_DATA_JAVA = "dd/MM/yyyy";
    private static final String REGEX_NUMERICO = "-?\\d+(\\.\\d+)?";
    private static final String REGEX_DATE = "([0-9]{2})/([0-9]{2})/([0-9]{4})";

    public static void main(String[] args) {
        run();
    }

    private static void run(){
        File folder = new File(FOLDER);
        File[] files = folder.listFiles();

        Arrays.stream(Objects.requireNonNull(files))
                .filter(File::isFile)
                .filter(n -> n.getName().endsWith(EXTENSAO_XLS))
                .forEach(f -> {
                    try {
                        FileUtils.delete(f);
                    } catch (IOException e) {
                        throw new RuntimeException(e);
                    }
                });
        Arrays.stream(Objects.requireNonNull(files))
                .filter(File::isFile)
                .filter(n -> n.getName().endsWith(EXTENSAO_CSV))
                .forEach(TesteCsvToExcel::convertToExcel);
    }

    private static void convertToExcel(File csvFile) {
        String arquivoXls = csvFile.getPath().replace(EXTENSAO_CSV, EXTENSAO_XLS);
        BufferedReader br = null;
        String linha = "";
        try {
            List<String[]> linhas = new ArrayList<>();
            br = new BufferedReader(new FileReader(csvFile));
            while ((linha = br.readLine()) != null){
                String[] pares = linha.split(CSV_DIVISOR);
                linhas.add(pares);
            }
            copyToXls(arquivoXls, linhas);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } finally {
            if (br != null){
                try {
                    br.close();
                } catch (IOException e) {
                    throw new RuntimeException(e);
                }
            }
        }
    }

    @SneakyThrows
    private static void copyToXls(String arquivoXls, List<String[]> linhas) {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet(SHEET_NAME);
        CreationHelper createHelper = wb.getCreationHelper();
        ZoneId defaultZoneId = ZoneId.systemDefault();

        int rowCount = 0;
        for (var linha: linhas){
            Row row = sheet.createRow(rowCount++);
            int colCount = 0;
            for (String conteudo: linha){
                Cell cell = row.createCell(colCount++);

                boolean num = isNumerico(conteudo);
                boolean date = isData(conteudo);
                if (num) {
                    cell.setCellValue(new BigDecimal(conteudo).doubleValue());
                } else if (date){
                    CellStyle cellStyle = wb.createCellStyle();

                    cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(FORMATO_DATA_EXCEL));
                    cell.setCellStyle(cellStyle);

                    DateTimeFormatter df = DateTimeFormatter.ofPattern(FORMATO_DATA_JAVA);
                    LocalDate ld = LocalDate.parse(conteudo, df);
                    cell.setCellValue(Date.from(ld.atStartOfDay(defaultZoneId).toInstant()));
                } else {
                    cell.setCellValue(conteudo);
                }
            }
        }

        try (FileOutputStream fout = new FileOutputStream(arquivoXls)){
            wb.write(fout);
        }
        wb.close();
    }

    private static boolean isNumerico(String s) {
        Pattern pattern = Pattern.compile(REGEX_NUMERICO);
        return (s != null && pattern.matcher(s).matches());
    }

    private static boolean isData(String s) {
        Pattern pattern = Pattern.compile(REGEX_DATE);
        return (s != null && pattern.matcher(s).matches());
    }
}
