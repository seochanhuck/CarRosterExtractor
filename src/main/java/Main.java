import swing.DragDropArea;

import javax.swing.*;

public class Main {
    public static void main(String[] args) {

        // 애플리케이션 실행
        SwingUtilities.invokeLater(() -> {
            DragDropArea frame = new DragDropArea();
            frame.setVisible(true);
        });
    }
}
