package excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.*;
import java.io.*;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;


public class ExcelReader {

    // car_excel.xls 파일에서의 마지막 셀 인덱스.
    public static final int CAR_EXCEL_LAST_CELL_IDX = 5;


    public static boolean handleExcelFile(File file1) {

        if (file1 == null) {
            return false;
        }
        try (
                FileInputStream registeredExcelStream = new FileInputStream(file1);
                InputStream carExcelStream = ExcelReader.class.getResourceAsStream("/car_excel.xls");
        ) {

            // XSSFWorkbook은 .xlsx 파일 지원
            // HSSFWorkbook은 .xls 파일 지원
            Workbook registeredExcelWorkbook;
            Workbook carExcelWorkbook;

            assert carExcelStream != null; // car_excel.xls 파일은 무조건 있음.
            if (file1.getName().endsWith(".xls")) {
                registeredExcelWorkbook = new HSSFWorkbook(registeredExcelStream);
                carExcelWorkbook = new HSSFWorkbook(carExcelStream);
            } else if (file1.getName().endsWith(".xlsx")) {
                registeredExcelWorkbook = new XSSFWorkbook(registeredExcelStream);
                carExcelWorkbook = new XSSFWorkbook(carExcelStream);
            }
            else {
                return false;
            }

            // 셀 초기화에 실패했다면 return null.
            boolean result = removeCarExcelCell(carExcelWorkbook);
            if (!result) {
                return false;
            }

            // 필요한 셀의 인덱스 리스트 가져오기.
            List<Integer> needCellsIdx = getNeedCellIdxList(registeredExcelWorkbook);

            // 실제로 carExcel에 붙여넣기.
            pasteDataToWorkbook(registeredExcelWorkbook, carExcelWorkbook, needCellsIdx);

            // 최종적으로 파일 저장.
            return sendExcel(carExcelWorkbook);

        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
    }


    // car_excel.xls 파일의 기본 셀 데이터 비우기.
    private static boolean removeCarExcelCell(Workbook workbook) {

        try {
            // 첫 번째 시트 가져오기(애초에 시트 하나만 존재)
            Sheet sheet = workbook.getSheetAt(0);
            // 혹시 모를 남아 있는 데이터를 제거.
            for (int i=3; i<=sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);

                for (int col=0; col<=CAR_EXCEL_LAST_CELL_IDX; col++) {
                    Cell cell = row.getCell(col);
                    cell.setBlank();
                }
            }
            return true;
        } catch (Exception exception) {
            System.out.println("removeCellExcelFile2() 에러 발생");
            exception.printStackTrace();
            return false;
        }
    }


    // 필요한 셀의 인덱스를 가져오는 메서드.
    private static List<Integer> getNeedCellIdxList(Workbook workbook1) {

        // 필요한 셀 내용.
        List<String> needCellsContent = new ArrayList<>() {{
           add("No");
           add("예약자");
           add("상품 명");
           add("시작일자");
           add("종료일자");
           add("차량번호(비고)");
        }};

        // 필요한 셀 내용에 해당하는 인덱스를 저장할 리스트.
        List<Integer> needCellsIdx = new ArrayList<>();

        try {
            Sheet sheet = workbook1.getSheetAt(0);
            Row firstRow = sheet.getRow(0); // 첫 번째 행 가져오기
            int lastColumnIdx = firstRow.getLastCellNum(); // 마지막 열 인덱스 가져오기

            // 필요한 열의 인덱스를 찾아 저장.
            for (int i=0; i<=lastColumnIdx; i++) {
                Cell cell = firstRow.getCell(i);
                if (cell != null && needCellsContent.contains(getCellValue(cell))) {
                    needCellsIdx.add(i);
                }
            }

            return needCellsIdx;

        } catch (Exception e) {
            System.out.println("getNeedCellIdxList() 에러 발생");
            return null;
        }
    }


    // 셀 값 가져오는 메서드.
    private static String getCellValue(Cell cell) {

        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case STRING -> {
                return cell.getStringCellValue();
            }
            case NUMERIC -> {
                return String.valueOf(cell.getNumericCellValue());
            }
            case BOOLEAN -> {
                return String.valueOf(cell.getBooleanCellValue());
            }
            default -> {
                return "";
            }
        }
    }


    // registeredExcel에서 고른 셀들을 carExcel에 실제로 붙여넣기.
    private static void pasteDataToWorkbook(Workbook registeredExcelWorkbook, Workbook carExcelWorkbook, List<Integer> needCellsIdx) {

        try {
            Sheet registeredExcelSheet = registeredExcelWorkbook.getSheetAt(0);
            Sheet carExcelSheet = carExcelWorkbook.getSheetAt(0);

            int carExcelLastRowIdx = carExcelSheet.getLastRowNum();
            int startRowIdx = 3;
            List<String[]> tempRowList = new ArrayList<>();


            /* registeredExcel을 순회하면서 carExcel에 필요한 셀의 인덱스에 맞춰서
            실시간으로 값을 채워 넣는다. */
            for (int i=1; i<=carExcelLastRowIdx; i++) {
                Row row1 = registeredExcelSheet.getRow(i);
                if (row1 == null) continue;

                int columnIdx = 0;
                Row row2 = carExcelSheet.getRow(startRowIdx++);

                for (int j : needCellsIdx) {
                    Cell cell1 = row1.getCell(j);
                    row2.getCell(columnIdx++).setCellValue(getCellValue(cell1));
                }
            }


            // workbook2에서 행을 가져와 리스트에 추가
            /*
                ## 기존에 있던 Row를 이용해서 List에 추가하고 정렬한 후에 다시 덮어쓰면 안됨.
                복제본이 아니라 참조값을 가지기 때문에 새로운 행을 만들어서 해야함. -> 배열을 통해 구현.
            */
            for (int i=3; i<=carExcelLastRowIdx; i++) {
                Row row = carExcelSheet.getRow(i);
                if (row != null ) {
                    if (row.getCell(0).toString().isEmpty()) {
                        break;
                    }

                    // 현재 데이터들을 따로 옮겨서 정렬하기 위해서 임시 배열 선언.
                    String[] cellValues = new String[CAR_EXCEL_LAST_CELL_IDX+1];
                    for (int j=0; j<=CAR_EXCEL_LAST_CELL_IDX; j++) {
                        Cell oldCell = row.getCell(j); // 기존 셀 값 가져오기.
                        cellValues[j] = oldCell != null ? getCellValue(oldCell) : ""; // null 체크 후 값 가져오기
                    }
                    tempRowList.add(cellValues); // 리스트에 저장.
                }
            }
            tempRowList.sort(Comparator.comparing(row -> row[2])); // 사이트명을 기준으로 정령


            // 사이트명으로 정렬했으니 이 기준으로 다시 carExcel에 붙여넣기.
            startRowIdx = 3;
            for (String[] tempRowArr : tempRowList) {
                Row targetRow = carExcelSheet.getRow(startRowIdx++);

                for (int k=0; k<=CAR_EXCEL_LAST_CELL_IDX; k++) {

                    Cell targetCell = targetRow.getCell(k);
                    targetCell.setCellValue(tempRowArr[k]);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    // 실제로 바탕화면에 파일 저장하기.
    private static boolean sendExcel(Workbook workbook) {

        String directory = System.getProperty("user.home") + "\\Desktop"; // 바탕화면 경로로 저장경로 설정.
        String newFilePath = directory + File.separator + "modified_car.xls"; // 바탕화면에 저장할 파일명 설정.
        File excelFile = new File(newFilePath); // 바탕화면에 저장할 파일 객체 생성.

        try (FileOutputStream fos = new FileOutputStream(newFilePath)) { // 파일 출력 스트림 생성
            workbook.write(fos); // 실제로 액셀 파일 바탕화면에 생성.
            System.out.println("수정된 파일이 다음 경로에 저장되었습니다 : " + newFilePath);
        } catch (Exception exception) {
            return false;
        }

        try {
            Desktop.getDesktop().open(excelFile); // 바탕화면에 저장한 액셀 파일 열기.
        } catch (Exception ignored) { }

        return true;
    }
}
