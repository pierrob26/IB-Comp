package com.company;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.util.Vector;

public class ExcelReaderApp extends JFrame {
    private JTable table;
    private DefaultTableModel tableModel;
    private JButton openButton;
    private JLabel statusLabel;

    public ExcelReaderApp() {
        setTitle("Excel Reader");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(800, 600);
        setLayout(new BorderLayout());

        // Create components
        openButton = new JButton("Open Excel File");
        statusLabel = new JLabel("No file loaded");
        tableModel = new DefaultTableModel();
        table = new JTable(tableModel);
        JScrollPane scrollPane = new JScrollPane(table);

        // Add components to frame
        JPanel topPanel = new JPanel();
        topPanel.add(openButton);
        topPanel.add(statusLabel);

        add(topPanel, BorderLayout.NORTH);
        add(scrollPane, BorderLayout.CENTER);

        // Add button listener
        openButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                openExcelFile();
            }
        });
    }

    private void openExcelFile() {
        JFileChooser fileChooser = new JFileChooser();
        int result = fileChooser.showOpenDialog(this);

        if (result == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            try {
                readExcelFile(selectedFile);
                statusLabel.setText("File loaded: " + selectedFile.getName());
            } catch (Exception ex) {
                JOptionPane.showMessageDialog(this,
                        "Error reading file: " + ex.getMessage(),
                        "Error",
                        JOptionPane.ERROR_MESSAGE);
                statusLabel.setText("Error loading file");
            }
        }
    }

    private void readExcelFile(File file) throws Exception {
        FileInputStream fis = new FileInputStream(file);
        Workbook workbook;

        // Determine the file type (XLSX or XLS)
        if (file.getName().toLowerCase().endsWith("xlsx")) {
            workbook = new XSSFWorkbook(fis);
        } else if (file.getName().toLowerCase().endsWith("xls")) {
            workbook = new HSSFWorkbook(fis);
        } else {
            throw new Exception("Not a valid Excel file");
        }

        // Get first sheet
        Sheet sheet = workbook.getSheetAt(0);

        // Clear existing table data
        tableModel.setRowCount(0);
        tableModel.setColumnCount(0);

        // Read header row
        Row headerRow = sheet.getRow(0);
        if (headerRow != null) {
            int columnCount = headerRow.getLastCellNum();
            Vector<String> columnNames = new Vector<>();

            for (int i = 0; i < columnCount; i++) {
                Cell cell = headerRow.getCell(i);
                columnNames.add(cell != null ? cell.toString() : "Column " + (i + 1));
            }
            tableModel.setColumnIdentifiers(columnNames);

            // Read data rows
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    Vector<Object> rowData = new Vector<>();
                    for (int j = 0; j < columnCount; j++) {
                        Cell cell = row.getCell(j);
                        rowData.add(cell != null ? getCellValue(cell) : "");
                    }
                    tableModel.addRow(rowData);
                }
            }
        }

        workbook.close();
        fis.close();
    }

    private Object getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue();
                }
                return cell.getNumericCellValue();
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                return cell.getCellFormula();
            default:
                return cell.toString();
        }
    }


}