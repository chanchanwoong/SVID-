import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.util.regex.*;

public class SVIDExtractor {

    static class SVIDData {
        private String svid;
        private String name;
        private String unit;

        public SVIDData(String svid, String name, String unit) {
            this.svid = svid;
            this.name = name;
            this.unit = unit;
        }

        public String getSvid() {
            return svid;
        }

        public String getName() {
            return name;
        }

        public String getUnit() {
            return unit;
        }
    }

    public static void main(String[] args) {
        String fileName = "LTS40_SVID LIST";
        String inputFile = "C:\\Users\\Drimsys\\Desktop\\나노팹\\SVID 관련 작업\\SVID_SSL 작업\\" + fileName + ".txt";
        String outputFile = "C:\\Users\\Drimsys\\Desktop\\나노팹\\SVID 관련 작업\\SVID_SSL 작업\\" + fileName + ".xlsx";

        try {
            List<SVIDData> svidDataList = extractSVIDData(inputFile);
            createExcelFile(svidDataList, outputFile);
            System.out.println("Excel 파일이 성공적으로 생성되었습니다: " + outputFile);
        } catch (IOException e) {
            System.err.println("오류 발생: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private static List<SVIDData> extractSVIDData(String fileName) throws IOException {
        List<SVIDData> svidDataList = new ArrayList<>();

        try (BufferedReader reader = new BufferedReader(new FileReader(fileName))) {
            String line;
            boolean foundRootList = false;

            String currentSvid = null;
            String currentName = null;
            String currentUnit = null;
            int currentSublistIndex = 0;

            while ((line = reader.readLine()) != null) {
                // 루트 리스트 찾기
                if (!foundRootList && line.trim().startsWith("L[")) {
                    Pattern pattern = Pattern.compile("L\\[(\\d+)\\]");
                    Matcher matcher = pattern.matcher(line);
                    if (matcher.find()) {
                        foundRootList = true;
                        continue;
                    }
                }

                if (foundRootList) {
                    // 서브리스트 시작 감지
                    if (line.trim().startsWith("L[3]")) {
                        currentSublistIndex = 0;
                        currentSvid = null;
                        currentName = null;
                        currentUnit = null;
                        continue;
                    }

                    currentSublistIndex++;

                    // SVID 추출 (첫 번째 항목)
                    if (currentSublistIndex == 1) {
                        Pattern svidPattern = Pattern.compile("I2\\[(\\d+)\\]|U4\\[(\\d+)\\]");
                        Matcher svidMatcher = svidPattern.matcher(line);
                        if (svidMatcher.find()) {
                            currentSvid = svidMatcher.group(1) != null ? svidMatcher.group(1) : svidMatcher.group(2);
                        }
                    }
                    // NAME 추출 (두 번째 항목)
                    else if (currentSublistIndex == 2) {
                        if (line.trim().equals("A")) {
                            // NAME이 비어있는 경우
                            currentName = "";
                        } else {
                            // A 다음에 있는 대괄호 안의 전체 내용을 추출
                            int startIdx = line.indexOf("A[") + 2;
                            if (startIdx > 1) {  // A[ 로 시작하는 경우
                                // 마지막 대괄호 위치 찾기 (라인의 마지막 문자가 ] 인 경우)
                                int endIdx = line.lastIndexOf("]");
                                if (endIdx > startIdx) {
                                    currentName = line.substring(startIdx, endIdx);
                                } else {
                                    currentName = "";
                                }
                            } else {
                                currentName = "";
                            }
                        }
                    }
                    // UNIT 추출 (세 번째 항목)
                    else if (currentSublistIndex == 3) {
                        if (line.trim().equals("A")) {
                            currentUnit = "";
                        } else {
                            // A 다음에 있는 대괄호 안의 전체 내용을 추출
                            int startIdx = line.indexOf("A[") + 2;
                            if (startIdx > 1) {  // A[ 로 시작하는 경우
                                // 마지막 대괄호 위치 찾기
                                int endIdx = line.lastIndexOf("]");
                                if (endIdx > startIdx) {
                                    currentUnit = line.substring(startIdx, endIdx);
                                } else {
                                    currentUnit = "";
                                }
                            } else {
                                currentUnit = "";
                            }
                        }

                        // 서브리스트의 마지막 항목이므로 데이터 추가
                        if (currentSvid != null) {
                            svidDataList.add(new SVIDData(currentSvid, currentName, currentUnit));
                        }
                    }
                }
            }
        }

        return svidDataList;
    }

    private static void createExcelFile(List<SVIDData> svidDataList, String outputFile) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("SVID Data");

            // 헤더 생성
            Row headerRow = sheet.createRow(0);
            Cell headerCell1 = headerRow.createCell(0);
            headerCell1.setCellValue("SVID");
            Cell headerCell2 = headerRow.createCell(1);
            headerCell2.setCellValue("NAME");
            Cell headerCell3 = headerRow.createCell(2);
            headerCell3.setCellValue("UNIT");

            // 데이터 채우기
            int rowNum = 1;
            for (SVIDData data : svidDataList) {
                Row row = sheet.createRow(rowNum++);

                Cell cell1 = row.createCell(0);
                cell1.setCellValue(data.getSvid());

                Cell cell2 = row.createCell(1);
                cell2.setCellValue(data.getName() != null ? data.getName() : "");

                Cell cell3 = row.createCell(2);
                cell3.setCellValue(data.getUnit() != null ? data.getUnit() : "");
            }

            // 열 너비 자동 조정
            for (int i = 0; i < 3; i++) {
                sheet.autoSizeColumn(i);
            }

            // 파일 저장
            try (FileOutputStream outputStream = new FileOutputStream(outputFile)) {
                workbook.write(outputStream);
            }
        }
    }
}