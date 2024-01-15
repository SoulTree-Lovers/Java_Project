package kr.excel.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelExample {
    public static void main(String[] args) {
        try {
            FileInputStream file = new FileInputStream(new File("example.xlsx")); // 파일 읽어오기
            Workbook workbook = WorkbookFactory.create(file); // 메모리에 올라온 가상의 엑셀 시트를 워크북이라 함.
            Sheet sheet = workbook.getSheetAt(0); // 0번째 시트 가져오기

            for (Row row : sheet) {
                for (Cell cell : row) {
                    switch (cell.getCellType()) { // 셀에 있는 데이터 타입에 따라 각각 다르게 출력
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                Date date = cell.getDateCellValue();
                                DateFormat dateFormat = new SimpleDateFormat("yyyy-mm-dd");
                                String formattedDate = dateFormat.format(date);
                                System.out.print(formattedDate + "\t");
                            } else {
                                double numericValue = cell.getNumericCellValue();
                                if (numericValue == Math.floor(numericValue)) { // 정수인지 확인
                                    int intValue = (int) numericValue;
                                    System.out.print(intValue + "\t");
                                } else { // 실수라면
                                    System.out.print(numericValue + "\t");
                                }
                            }
                            break;
                        case STRING:
                            String stringValue = cell.getStringCellValue();
                            System.out.print(stringValue + "\t");
                            break;
                        case BOOLEAN:
                            Boolean boolValue = cell.getBooleanCellValue();
                            System.out.print(boolValue + "\t");
                            break;
                        case FORMULA:
                            String formulaValue = cell.getCellFormula();
                            System.out.print(formulaValue + "\t");
                            break;
                        case BLANK:
                            System.out.print("\t");
                            break;
                        default:
                            System.out.print("\t");
                            break;
                    }
                }
                System.out.println(); // 줄 바꿈
            }
            file.close();
            System.out.println("엑셀에서 데이터 읽어오기 종료");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
