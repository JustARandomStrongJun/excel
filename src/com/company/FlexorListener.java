package com.company;

import jssc.SerialPortException;

import javax.swing.*;
import javax.swing.text.DefaultCaret;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.io.PrintStream;

/**
 * Created by zZzZz on 16.07.2018.
 */
public class FlexorListener {


    public JPanel getRootPanel() {
        return rootPanel;
    }
    private JFrame f; //Main frame
    private JPanel rootPanel;
    private JButton button1;
    private JButton stopButton;
    private JComboBox portSwicher;
    private JTextArea output;
    private JScrollPane scroll;
    private ComListener comListener;
    private Thread t1;
    private Main obj;
    static PrintStream printStream;


    public FlexorListener() throws IOException {

        comListener = new ComListener();
        obj = new Main();
        f = new JFrame("Flexor to Excel");
        f.getContentPane().setLayout(new FlowLayout());
        output = new JTextArea("", 5, 50);
        output.setLineWrap(true);
        scroll = new JScrollPane(output);
        scroll.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
        scroll.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);



        printStream = new PrintStream(new CustomOutputStream(output));
        System.setOut(printStream);
        System.setErr(printStream);

        for (String s : comListener.getPorts()) {
            portSwicher.addItem(s);
        }
        try {
            portSwicher.setSelectedItem(RWSetting.ReadString("Port"));
        } catch (IOException e1) {
            e1.printStackTrace();
        }

        button1.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    comListener.startPort();
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
                t1 = new Thread(obj);
                t1.start();

            }
        });
        stopButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    comListener.stopPort();

                } catch (SerialPortException e1) {
                    System.out.println(e1);
                }
            }

        });
        portSwicher.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                //System.out.println(portSwicher.getSelectedItem());

                try {
                    RWSetting.writeString("Port", String.valueOf(portSwicher.getSelectedItem()));
                } catch (IOException e1) {
                    System.out.println(e1.toString());
                }

            }
        });

    }

    private void createUIComponents() {
        // TODO: place custom component creation code here
    }
}
