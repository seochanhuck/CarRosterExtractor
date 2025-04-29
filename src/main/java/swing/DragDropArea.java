package swing;

import excel.ExcelReader;

import javax.swing.*;

import java.awt.*;
import java.awt.datatransfer.DataFlavor;
import java.awt.dnd.DropTarget;
import java.awt.dnd.DropTargetDropEvent;
import java.io.File;
import java.util.List;

public class DragDropArea extends JFrame {

    private final JTextArea textArea1; // 드래그 앤 드롭의 텍스트 영역.
    private File selectedFile1; // 드래그 앤 드롭된 파일을 저장할 변수.


    public DragDropArea() {

        selectedFile1 = null;

        // 프레임 설정
        setTitle("주차등록 액셀 추출기"); // UI 제목 설정.
        setSize(600, 400); // UI 크기 설정.
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE); // 종료 시 애플리케이션 종료.
        setLocationRelativeTo(null); // 화면 중앙에 위치.

        // 드래그 영역 설정
        textArea1 = new JTextArea("액셀 파일의 확장자인 .xls와 .xlsx 파일만 추가해주세요.");
        textArea1.setEditable(false); // 텍스트 영역 편집 불가.
        textArea1.setFocusable(false); // 포커스 비활성화.
        textArea1.setDropTarget(new DropTarget() { // 드롭 타겟 설정.
            public synchronized void drop(DropTargetDropEvent evt) { // 드롭 이벤트 처리.
                handleDrop(evt, textArea1);
            }
        });

        // 메일 패널 생성.
        JPanel panel = new JPanel(new GridLayout(2, 1));
        panel.add(new JScrollPane(textArea1));
        add(panel, BorderLayout.CENTER);

        // 버튼 생성 및 추가.
        JButton extractButton = new JButton("추출"); // 추출 버튼 추가.
        extractButton.addActionListener(e -> { // 추출 버튼 리스너 추가.
            extractData(textArea1, selectedFile1); // 액셀 파일 추출.
        });
        add(extractButton, BorderLayout.SOUTH);
    }

    private void handleDrop(DropTargetDropEvent evt, JTextArea textArea) {
        try {
            evt.acceptDrop(java.awt.dnd.DnDConstants.ACTION_COPY); // 드롭 허용.
            List<?> droppedFiles = (List<?>) evt.getTransferable().getTransferData(DataFlavor.javaFileListFlavor); // 드롭된 파일 목록 가져오기.

            if (!droppedFiles.isEmpty() && droppedFiles.get(0) instanceof File file) { // 드롭된 파일이 비어있지 않고 File 객체인지 확인.
                if (isExcelFile(file)) { // 엑셀 파일인지 확인.
                    textArea.setText("<선택한 파일 경로>\n" + "파일 : " + file.getName() + "\n" + file.getAbsolutePath());
                    selectedFile1 = file;
                    System.out.println("file1 등록 완료");
                } else { // 엑셀 파일이 아닌 경우.
                    showWarningMessage("액셀 파일만 사용 가능합니다.");
                }
            } else {
                textArea.setText("파일을 다시 선택해주세요.");
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    private void extractData(JTextArea textArea1, File file1) {

        if (file1 == null) {
            showWarningMessage("선택하지 않은 액셀파일이 있습니다.");
            return;
        }

        boolean result = ExcelReader.handleExcelFile(file1); // 드래그한 파일을 전달해서 실제로 작업 수행.
        if (result) { // 성공하면 프로그램 닫기.
            System.exit(0);
        } else { // 실패하면 경고 메시지 출력.
            String currentText = textArea1.getText();
            textArea1.setText(currentText + "\n\n@오류가 발생했습니다.");
        }
    }

    // 엑셀 파일인지 검사하는 메서드.
    private boolean isExcelFile(File file) {
        String fileName = file.getName().toLowerCase();
        return fileName.endsWith(".xlsx") || fileName.endsWith(".xls");
    }

    // 경고 메시지 출력 메서드.
    private void showWarningMessage(String message) {
        JOptionPane.showMessageDialog(this, message, "경고", JOptionPane.WARNING_MESSAGE);
    }
}
