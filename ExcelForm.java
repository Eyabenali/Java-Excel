import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.io.FileOutputStream;

public class ExcelForm {
    public static void main(String[] args) {
        JFrame frame = new JFrame("Excel Form");
        frame.setSize(300, 200);
        frame.setLayout(new GridLayout(4, 2));

        JLabel nameLabel = new JLabel("Name:");
        JTextField nameField = new JTextField();
        JLabel gradeLabel = new JLabel("Grade:");
        JTextField gradeField = new JTextField();

        JButton saveButton = new JButton("Save to Excel");

        frame.add(nameLabel);
        frame.add(nameField);
        frame.add(gradeLabel);
        frame.add(gradeField);
        frame.add(new JLabel());
        frame.add(saveButton);

        saveButton.addActionListener(e -> {
            String name = nameField.getText();
            int grade = Integer.parseInt(gradeField.getText());

            try (Workbook workbook = new XSSFWorkbook()) {
                Sheet sheet = workbook.createSheet("Input");
                Row row = sheet.createRow(0);
                row.createCell(0).setCellValue("Name");
                row.createCell(1).setCellValue("Grade");

                Row dataRow = sheet.createRow(1);
                dataRow.createCell(0).setCellValue(name);
                dataRow.createCell(1).setCellValue(grade);

                FileOutputStream fos = new FileOutputStream("form_output.xlsx");
                workbook.write(fos);
                fos.close();

                JOptionPane.showMessageDialog(frame, "Saved!");
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        });

        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setVisible(true);
    }
}
