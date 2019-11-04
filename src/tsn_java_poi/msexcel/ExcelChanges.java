package tsn_java_poi.msexcel;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.FileInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.awt.Desktop;
import java.io.FileNotFoundException;
import java.io.InputStream;
import static java.lang.System.exit;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.ss.usermodel.CellType;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;

public class ExcelChanges {

    /**
     * Чтение данных из документа MS Excel
     *
     * @param filename имя файла для чтения
     * @return
     */
    String readData(String filename) {

        String result = ""; // Строка со значениями из таблицы MS Excel
        HSSFWorkbook wb = null; // Рабочая книга MS Excel

        try {
            wb = new HSSFWorkbook(new FileInputStream(filename)); // Подключение к MS Excel
        } catch (IOException e) {
            System.err.println("File not found!");
            exit(-1); // Выход при ошибке доступа к файлу
        }

        Sheet sheet = wb.getSheetAt(0); // Лист Excel
        Iterator<Row> it = sheet.iterator(); // Итератор строк (цикл по строкам)
        while (it.hasNext()) { // Цикл по строкам текущего листа
            Row row = it.next(); // Текущая строка
            Iterator<Cell> cells = row.iterator(); // Итератор столбцов для строки (цикл по столбцам)
            while (cells.hasNext()) { // Цикл по столбцам текущей стоки
                Cell cell = cells.next(); // Текущая ячейка листа (из цикла в цикле)
                CellType cellType = cell.getCellType(); // Тип текущей ячейки 
                switch (cellType) {
                    case STRING: // Ячейка строкового типа
                        result += cell.getStringCellValue() + "=";
                        break;
                    case NUMERIC: // Ячейка числового типа
                        result += "[" + cell.getNumericCellValue() + "] ";
                        break;

                    case FORMULA: // Ячейка с формулой
                        result += "[" + cell.getNumericCellValue() + "] ";
                        break;
                    default: // Ячейка другого типа
                        result += " | ";
                        break;
                }
            }
            result += "\n";
        }

        return result;
    }

    /**
     * Запись данных в документ MS Excel
     *
     * @param filename имя файла для записи
     */
    void writeData(String filename) {
        HSSFWorkbook workbook = new HSSFWorkbook(); // Документ MS Excel
        Sheet sheet = workbook.createSheet(); // Лист MS Excel
        HSSFDataFormat df = workbook.createDataFormat(); // Формат ячейки
        HSSFCellStyle style = workbook.createCellStyle(); // Стиль ячейки
        style.setDataFormat(df.getFormat("0.000")); // Установка формата ячейки
        for (int i = 0; i < 10; i++) { // Цикл для строк 
            Row row = sheet.createRow(i); // Создание строки
            for (int j = 0; j < 5; j++) { // Цикл для столбцов 
                Cell cell = row.createCell(j); // Создание ячейки строки
                cell.setCellValue(i * j); // Установка значения ячейки
                cell.setCellStyle(style); // Установка стиля ячейки
                cell.setCellType(NUMERIC); // Установка типа ячейки
            }
        }

        try {
            FileOutputStream out = new FileOutputStream(filename); // Поток для записи данных
            workbook.write(out); // Запись данных в MS Excel
            out.close(); // Закрытие потока записи
        } catch (IOException ex) {
        }

    }

    /**
     * Изменение данных в документе MS Excel
     *
     * @param inputFileName входной файл
     * @param outputFileName выходной файл
     * @throws IOException
     */
    void modifData(String inputFileName, String outputFileName) throws IOException {
        POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(inputFileName));
        HSSFWorkbook wb = new HSSFWorkbook(fs); // Документ MS Excel
        HSSFSheet sheet = wb.getSheetAt(0); // Лист MS Excel
        HSSFRow row = null; // Строка
        HSSFCell cell = null; // Ячейка
        int rows = sheet.getPhysicalNumberOfRows(); // Получение числа строк
        for (int r = 0; r < rows; r++) { // Цикл по строкам таблицы
            row = sheet.getRow(r); // Получение строки в цикле
            if (row != null) { // Если стока не пустая
                cell = row.getCell(0); // Получение первой ячейки
                if (cell != null) { // Если ячейка не пустая
                    cell.setCellValue("Modified " + r); // Устанавливаем новое значение ячейки
                }
            }
        }
        FileOutputStream fileOut = new FileOutputStream(outputFileName); // Поток для записи в файл
        wb.write(fileOut); // Сохранение данных в документе MS Excel на диске
        fileOut.close(); // Закрытие файлового потока
    }

    /**
     * Извлечение данных из документа MS Excel
     *
     * @param fileName имя файла MS Excel
     * @throws FileNotFoundException
     * @throws IOException
     */
    void extractor(String fileName) throws FileNotFoundException, IOException {
        InputStream in = new FileInputStream(fileName); // Поток чтения из файла
        HSSFWorkbook wb = new HSSFWorkbook(in); // Документ MS Excel
        ExcelExtractor extractor = new ExcelExtractor(wb); // Извлекатель данных
        extractor.setFormulasNotResults(false); // Считать значение формул
        extractor.setIncludeSheetNames(false); // Не считывать название листов книги MS Excel
        String text = extractor.getText(); // Получить содержимое документа MS Excel
        System.out.println(text); // Вывод содержимого документа MS Excel на экран
    }

    public static void main(String... args) {
        try {
            String dir = new File(".").getAbsoluteFile().getParentFile().getAbsolutePath()
                    + System.getProperty("file.separator"); // Узнаем текущий каталог
            ExcelChanges excel = new ExcelChanges();
            excel.writeData(dir + "input.xls"); // Создание на диске документа MS Excel
            excel.modifData(dir + "input.xls", dir + "output.xls"); // Модификация данных в документе MS Excel
            System.out.println(excel.readData(dir + "output.xls")); // Вывод содержимого документа MS Excel на экран
            excel.extractor(dir + "output.xls"); // Извлечение данных из документа MS Excel
            Desktop.getDesktop().open(new File(dir + "output.xls")); // Запуск документа в MS Excel
        } catch (IOException ex) {
            System.err.println("Error!");
        }
    }

}
