import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.plaf.FontUIResource;
import java.awt.*;
import java.awt.datatransfer.DataFlavor;
import java.awt.dnd.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.List;
import java.util.TooManyListenersException;

public class SVIDConverterApp extends JFrame {

    private JTextArea logArea;
    private JPanel dropPanel;
    private JButton convertButton;
    private List<File> droppedFiles = new ArrayList<>();

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

    // 글꼴 설정을 위한 메소드 추가
    private static void setUIFont() {
        // 시스템에서 사용 가능한 한글 지원 폰트 찾기
        String[] preferredFonts = {"맑은 고딕", "나눔고딕", "굴림", "돋움", "Arial Unicode MS", "Malgun Gothic"};

        Font preferredFont = null;
        for (String fontName : preferredFonts) {
            Font testFont = new Font(fontName, Font.PLAIN, 12);
            if (testFont.canDisplay('한') && testFont.canDisplay('글')) {
                preferredFont = testFont;
                break;
            }
        }

        // 적절한 폰트를 찾지 못했을 경우 기본 폰트 사용
        if (preferredFont == null) {
            preferredFont = new Font(Font.SANS_SERIF, Font.PLAIN, 12);
        }

        // 모든 UI 컴포넌트에 폰트 적용
        FontUIResource fontResource = new FontUIResource(preferredFont);
        Enumeration<Object> keys = UIManager.getDefaults().keys();
        while (keys.hasMoreElements()) {
            Object key = keys.nextElement();
            Object value = UIManager.get(key);
            if (value instanceof FontUIResource) {
                UIManager.put(key, fontResource);
            }
        }
    }

    public SVIDConverterApp() {
        super("SVID 데이터 변환기");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(600, 400);
        setLocationRelativeTo(null);

        // 메인 패널
        JPanel mainPanel = new JPanel(new BorderLayout(10, 10));
        mainPanel.setBorder(new EmptyBorder(10, 10, 10, 10));

        // 드롭 패널 만들기
        dropPanel = new JPanel(new BorderLayout());
        dropPanel.setBorder(BorderFactory.createTitledBorder("파일을 여기에 드래그 앤 드롭하세요"));
        dropPanel.setPreferredSize(new Dimension(580, 150));
        dropPanel.setBackground(new Color(240, 240, 240));

        JLabel dropLabel = new JLabel("텍스트 파일(.txt)을 여기에 드래그하세요", JLabel.CENTER);
        dropLabel.setFont(new Font("맑은 고딕", Font.BOLD, 14)); // 한글 지원 폰트로 변경
        dropPanel.add(dropLabel, BorderLayout.CENTER);

        // 드래그 앤 드롭 기능 추가
        DropTarget target = new DropTarget();
        try {
            target.addDropTargetListener(new DropTargetAdapter() {
                @Override
                public void drop(DropTargetDropEvent event) {
                    try {
                        event.acceptDrop(DnDConstants.ACTION_COPY);
                        List<File> files = (List<File>) event.getTransferable().getTransferData(DataFlavor.javaFileListFlavor);

                        droppedFiles.clear();
                        for (File file : files) {
                            if (file.getName().toLowerCase().endsWith(".txt")) {
                                droppedFiles.add(file);
                                logArea.append("파일이 추가되었습니다: " + file.getName() + "\n");
                            } else {
                                logArea.append("지원하지 않는 파일 형식입니다: " + file.getName() + "\n");
                            }
                        }

                        if (!droppedFiles.isEmpty()) {
                            convertButton.setEnabled(true);
                            dropLabel.setText(droppedFiles.size() + "개의 파일이 추가됨 - 변환 버튼을 클릭하세요");
                        } else {
                            convertButton.setEnabled(false);
                            dropLabel.setText("텍스트 파일(.txt)을 여기에 드래그하세요");
                        }

                        event.dropComplete(true);
                    } catch (Exception e) {
                        e.printStackTrace();
                        event.dropComplete(false);
                        logArea.append("오류 발생: " + e.getMessage() + "\n");
                    }
                }

                @Override
                public void dragEnter(DropTargetDragEvent event) {
                    dropPanel.setBackground(new Color(200, 230, 200));
                }

                @Override
                public void dragExit(DropTargetEvent event) {
                    dropPanel.setBackground(new Color(240, 240, 240));
                }
            });
        } catch (TooManyListenersException e) {
            e.printStackTrace();
        }

        dropPanel.setDropTarget(target);

        // 로그 영역
        logArea = new JTextArea();
        logArea.setEditable(false);
        logArea.setFont(new Font("맑은 고딕", Font.PLAIN, 12)); // 한글 지원 폰트로 변경
        JScrollPane scrollPane = new JScrollPane(logArea);
        scrollPane.setPreferredSize(new Dimension(580, 150));

        // 변환 버튼 
        convertButton = new JButton("변환");
        convertButton.setEnabled(false);
        convertButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                convertFiles();
            }
        });

        // UI 배치
        mainPanel.add(dropPanel, BorderLayout.NORTH);
        mainPanel.add(scrollPane, BorderLayout.CENTER);

        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        buttonPanel.add(convertButton);
        mainPanel.add(buttonPanel, BorderLayout.SOUTH);

        add(mainPanel);
        setVisible(true);
    }

    private void convertFiles() {
        for (File file : droppedFiles) {
            try {
                logArea.append("파일 처리 중: " + file.getName() + "\n");
                List<SVIDData> svidDataList = extractSVIDData(file.getAbsolutePath());

                // 출력 파일명 생성 (확장자 변경)
                String outputPath = file.getAbsolutePath().replace(".txt", ".xlsx");
                if (outputPath.equals(file.getAbsolutePath())) {
                    outputPath = file.getAbsolutePath() + ".xlsx";
                }

                createExcelFile(svidDataList, outputPath);
                logArea.append("변환 완료: " + new File(outputPath).getName() + "\n");
            } catch (IOException e) {
                logArea.append("오류 발생: " + e.getMessage() + "\n");
                e.printStackTrace();
            }
        }

        logArea.append("모든 파일 처리 완료!\n");
        // 처리 후 파일 목록 비우기
        droppedFiles.clear();
        convertButton.setEnabled(false);

        JLabel label = (JLabel) dropPanel.getComponent(0);
        label.setText("텍스트 파일(.txt)을 여기에 드래그하세요");
        dropPanel.setBackground(new Color(240, 240, 240));
    }

    private static List<SVIDData> extractSVIDData(String fileName) throws IOException {
        List<SVIDData> svidDataList = new ArrayList<>();

        // 파일 인코딩 설정 - UTF-8 사용
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream(fileName), "UTF-8"))) {
            String line;
            boolean foundRootList = false;

            String currentSvid = null;
            String currentName = null;
            String currentUnit = null;
            int currentSublistIndex = 0;

            while ((line = reader.readLine()) != null) {
                // 루트 리스트 찾기
                if (!foundRootList && line.trim().startsWith("L[")) {
                    foundRootList = true;
                    continue;
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
                        if (line.contains("I2[") || line.contains("U4[")) {
                            int startIdx = Math.max(line.indexOf("I2["), line.indexOf("U4["));
                            startIdx = startIdx == -1 ? (line.indexOf("I2[") != -1 ? line.indexOf("I2[") : line.indexOf("U4[")) : startIdx;
                            startIdx = startIdx + 3; // "I2[" or "U4[" length

                            int endIdx = line.indexOf("]", startIdx);
                            if (endIdx > startIdx) {
                                currentSvid = line.substring(startIdx, endIdx);
                            }
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
                                // 마지막 대괄호 위치 찾기
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

    public static void main(String[] args) {
        // Swing UI는 EDT(Event Dispatch Thread)에서 실행
        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                try {
                    // Look & Feel 설정
                    UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());

                    // 글꼴 설정 적용
                    setUIFont();
                } catch (Exception e) {
                    e.printStackTrace();
                }

                new SVIDConverterApp();
            }
        });
    }
}