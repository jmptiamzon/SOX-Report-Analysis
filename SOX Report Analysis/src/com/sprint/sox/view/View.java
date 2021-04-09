package com.sprint.sox.view;

import java.awt.BorderLayout;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.nio.file.Paths;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.UIManager;
import javax.swing.UIManager.LookAndFeelInfo;
import javax.swing.border.TitledBorder;
import javax.swing.filechooser.FileNameExtensionFilter;

import com.sprint.sox.controller.Controller;

public class View extends JFrame implements ActionListener {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private File []files;
	private JPanel paneTop, paneCenter, paneBottom;
	private GridBagConstraints gbc;
	private JTextArea fileNameContainer;
	private JLabel statusLbl;
	private JButton startBtn, browseBtn;
	private Insets inset;
	private TitledBorder titledBorder;
	private JFileChooser fileChooser;
	private JScrollPane jScrollPane;
	private Controller controller;
	private String []soxFilePath;
	
	
	public View() {
		try {
			for (LookAndFeelInfo info : UIManager.getInstalledLookAndFeels()) {
				if ("Nimbus".equals(info.getName())) {
					UIManager.setLookAndFeel(info.getClassName());
					break;
				}
			}
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
		
		setLayout(new BorderLayout());
		paneTop = new JPanel(new GridBagLayout());
		paneCenter = new JPanel(new GridBagLayout());
		paneBottom = new JPanel(new GridBagLayout());
		gbc = new GridBagConstraints();
		inset = new Insets(5, 5, 5, 5);
		
		statusLbl = new JLabel();
		startBtn = new JButton("Generate Report");
		browseBtn = new JButton("Browse Files");
		fileNameContainer = new JTextArea("No files selected.");
		fileChooser = new JFileChooser();
		
		startBtn.setEnabled(false);
		fileNameContainer.setEnabled(false);
		browseBtn.addActionListener(this);
		startBtn.addActionListener(this);

		this.setResizable(false);
		setTitle("SOX Report Analysis");
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		paneTop.setPreferredSize(new Dimension(400, 200));
		
		add(paneTop, BorderLayout.NORTH);
		add(paneCenter, BorderLayout.CENTER);
		add(paneBottom, BorderLayout.SOUTH);
		
		titledBorder = new TitledBorder("Chosen Files");
		paneTop.setBorder(titledBorder);
		jScrollPane = new JScrollPane(fileNameContainer);
		setComponent(jScrollPane, paneTop, 1, 0, 0, 100);
		
		titledBorder = new TitledBorder("Status");
		paneCenter.setBorder(titledBorder);
		statusLbl.setText("Program not running.");
		setComponent(statusLbl, paneCenter, 1, 0, 0, 2);
		
		setComponent(browseBtn, paneBottom, 0.5, 0, 0, 0);
		setComponent(startBtn, paneBottom, 0.5, 1, 0, 0);
		
		pack();
		setLocationRelativeTo(null);
		
	}
	
	public void setComponent(Component component, JPanel compPanel, double weightx, int gridx, int gridy, int ipady) {
		gbc.fill = GridBagConstraints.HORIZONTAL;
		gbc.insets = inset;
		gbc.weightx = weightx;
		gbc.gridx = gridx;
		gbc.gridy = gridy;
		gbc.ipady = ipady;
		compPanel.add(component, gbc);
	}
	
	@Override
	public void actionPerformed(ActionEvent event) {
		if (event.getSource() == browseBtn) {
			fileChooser.setAcceptAllFileFilterUsed(false);
			fileChooser.setMultiSelectionEnabled(true);
			fileChooser.addChoosableFileFilter(new FileNameExtensionFilter("Select CSV Files", "csv"));
			
			int flag = fileChooser.showOpenDialog(this);
			if (flag == JFileChooser.APPROVE_OPTION) {
				String fileNameText =  "";
				files = fileChooser.getSelectedFiles();
				Integer []fileFlag = new Integer[] {0, 0, 0, 0, 0};
				boolean dateFlag = false;
				
				if (files.length > 5 || files.length < 5) {
					JOptionPane.showMessageDialog(null, "SOX Report contains 5 files to generate.");
				
				} else {
					
					for (int ctr = 0; ctr < files.length; ctr++) {
						if (Paths.get(files[ctr].getPath()).getFileName().toString().toLowerCase().contains("moves_and_po")) {
							fileFlag[0] += 1;
						}
						
						if (Paths.get(files[ctr].getPath()).getFileName().toString().toLowerCase().contains("rma_batch")) {
							fileFlag[1] += 1;
						}
						
						if (Paths.get(files[ctr].getPath()).getFileName().toString().toLowerCase().contains("rma_receipts")) {
							fileFlag[2] += 1;
						}
						
						if (Paths.get(files[ctr].getPath()).getFileName().toString().toLowerCase().contains("rsku_report")) {
							fileFlag[3] += 1;
						}
						
						if (Paths.get(files[ctr].getPath()).getFileName().toString().toLowerCase().contains("ship_confirm")) {
							fileFlag[4] += 1;
						}
					}
					
					if (fileFlag[0] == 1 && fileFlag[1] == 1 && fileFlag[2] == 1 && fileFlag[3] == 1 && fileFlag[4] == 1) {
						String date = Paths.get(files[0].getPath()).getFileName().toString().split(",")[1] + "," + Paths.get(files[0].getPath()).getFileName().toString().split(",")[2];
						
						/* SAVED THIS CODE IF EVER THERE WILL BE CHANGES REGARDING DATES
						String date;
						LocalDate friday = LocalDate.now();
						DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MMMdd,yyyy"); 
						while (friday.getDayOfWeek() != DayOfWeek.FRIDAY) {
							friday = friday.plusDays(1);
						}
						
						date = friday.format(formatter); */
						
						for (int ctr = 0; ctr < files.length; ctr++) {
							if (!Paths.get(files[ctr].getPath()).getFileName().toString().contains(date)) {
								dateFlag = true;
								break;
							}
						}
						
						if (dateFlag) {
							JOptionPane.showMessageDialog(null, "Dates must be the same across files.");
							
						} else {
							soxFilePath = new String[5];
							
							for (int ctr = 0; ctr < files.length; ctr++) {
								fileNameText += files[ctr].getPath() + "\n";
								soxFilePath[ctr] = files[ctr].getPath();
							}
							
							fileNameContainer.setText(fileNameText);
							startBtn.setEnabled(true);
							
						}
						
					} else {
						JOptionPane.showMessageDialog(null, "Please select the 5 files from the automated email.");
					}
					
				}

			} else {
				if (fileNameContainer.getText().isEmpty()) {
					startBtn.setEnabled(false);
				}
				
			}
		}
		
		if (event.getSource() == startBtn) {	
			controller = new Controller(this);
			controller.startApp();
			startBtn.setEnabled(false);
			
		}
		
	}
	
	public JLabel getStatusLbl() {
		return statusLbl;
	}
	
	public String[] getSoxFilePath() {
		return soxFilePath;
	}
	
}