package com.company;

import javax.swing.*;

public class MainOld {

    public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                new ExcelReaderApp().setVisible(true);
            }
        });
    }
}
