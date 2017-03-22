import java.awt.Component;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Image;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;

import javax.swing.ImageIcon;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import javax.swing.GroupLayout;
import javax.swing.GroupLayout.Alignment;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.JLabel;

import javax.swing.LayoutStyle.ComponentPlacement;
import javax.swing.border.TitledBorder;

import java.awt.Color;

import javax.swing.JButton;
import java.awt.Font;

import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;

import java.awt.event.ActionEvent;

import javax.swing.JComboBox;
import javax.swing.JSlider;

import javax.swing.UIManager;

/*
 * @author: Chenlei Zhang
 * @Data: 03/21/2017
 * @Email: ztimemakercl@gmail.com
 * 
 */
public class ImageLabelGUI extends JFrame {

	/************************Environment Setting************************/
	String pathZipcode = "E:\\ImageDownloadLess\\Los Angeles\\Los Angeles zipcode.xlsx";
	String pathHouseData = "E:\\ImageDownloadLess\\Los Angeles\\Los Angeles-ca-testless-test.xlsx";
	String pathImageRoot = "E:\\ImageDownloadLess\\Los Angeles\\";
	/*******************************************************************/
	
	private JPanel contentPane;
	private JPanel imageDisplayPanel;
	int pos = 0;
	private JLabel jLabel_Image;
	private JPanel imageInfo;
	private JPanel buttonPanel;
	private JPanel evaluatePanel;
	private JButton btnNext;
	int zipcode;
	String url;
	String[] houseInZipcode;
	int houseIndex;
	int houseImageIndex;
	private JLabel lblHouseIdTitle;
	private JLabel lblImageTypeTitle;
	private JLabel lblHouseZipId;
	private JLabel lblImageType;
	private JLabel lblInfo;
	private JLabel lblHouseOutfit;
	private JLabel lblLivingRoom;
	private JLabel lblBedroom;
	private JLabel lblKitchen;
	private JLabel lblRestroom;
	private JLabel lblGarden;
	private JLabel lblOverallEvaluation;
	
	static final int S_MIN = 0;
    static final int S_MAX = 10;
    static final int S_INIT = 0; 
    
    private JSlider sliderOutfit;
    private JSlider sliderLivingRoom;
    private JSlider sliderBedroom;
    private JSlider sliderKitchen;
    private JSlider sliderRestroom;
    private JSlider sliderGarden;
    private JSlider sliderOverall;
    private JLabel lblHOV;
    private JLabel lblLR;
    private JLabel lblB;
    private JLabel lblK;
    private JLabel lblR;
    private JLabel lblG;
    private JLabel lblO;
    private JButton btnNextHouse;
    private JButton btnPreHouse;
	

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					ImageLabelGUI frame = new ImageLabelGUI();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 * @throws IOException 
	 */
	public ImageLabelGUI() throws IOException {
		setTitle("House Evaluate");
		InitialGUI();
		//showImage(pos);
	}
	private void InitialGUI() throws IOException {
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 1060, 680);
		contentPane = new JPanel();
		contentPane.setPreferredSize(new Dimension(1060, 680));
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		
		imageDisplayPanel = new JPanel();
		imageDisplayPanel.setBackground(Color.WHITE);
		
		evaluatePanel = new JPanel();
		evaluatePanel.setBorder(new TitledBorder(null, "Step 2: Evaluate", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		
		buttonPanel = new JPanel();
		buttonPanel.setBackground(Color.WHITE);
		
		imageInfo = new JPanel();
		imageInfo.setBorder(new TitledBorder(null, "Step 1: Select a zipcode", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		
		JLabel lblHouseEvaluationBy = new JLabel("House Evaluating by Human Observation");
		lblHouseEvaluationBy.setFont(new Font("Arial", Font.BOLD, 18));
		GroupLayout gl_contentPane = new GroupLayout(contentPane);
		gl_contentPane.setHorizontalGroup(
			gl_contentPane.createParallelGroup(Alignment.TRAILING)
				.addGroup(Alignment.LEADING, gl_contentPane.createSequentialGroup()
					.addContainerGap()
					.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
						.addGroup(gl_contentPane.createSequentialGroup()
							.addGap(292)
							.addComponent(lblHouseEvaluationBy)
							.addContainerGap())
						.addGroup(gl_contentPane.createSequentialGroup()
							.addGroup(gl_contentPane.createParallelGroup(Alignment.TRAILING, false)
								.addComponent(buttonPanel, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
								.addComponent(imageDisplayPanel, GroupLayout.DEFAULT_SIZE, 587, Short.MAX_VALUE))
							.addPreferredGap(ComponentPlacement.RELATED)
							.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
								.addComponent(evaluatePanel, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
								.addComponent(imageInfo, GroupLayout.DEFAULT_SIZE, 411, Short.MAX_VALUE))
							.addGap(0))))
		);
		gl_contentPane.setVerticalGroup(
			gl_contentPane.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_contentPane.createSequentialGroup()
					.addContainerGap()
					.addComponent(lblHouseEvaluationBy, GroupLayout.DEFAULT_SIZE, 52, Short.MAX_VALUE)
					.addPreferredGap(ComponentPlacement.RELATED)
					.addGroup(gl_contentPane.createParallelGroup(Alignment.TRAILING)
						.addGroup(gl_contentPane.createSequentialGroup()
							.addGap(16)
							.addComponent(imageInfo, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
							.addPreferredGap(ComponentPlacement.RELATED)
							.addComponent(evaluatePanel, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
							.addContainerGap())
						.addGroup(gl_contentPane.createSequentialGroup()
							.addGap(32)
							.addComponent(imageDisplayPanel, GroupLayout.DEFAULT_SIZE, 455, Short.MAX_VALUE)
							.addPreferredGap(ComponentPlacement.RELATED)
							.addComponent(buttonPanel, GroupLayout.PREFERRED_SIZE, 91, Short.MAX_VALUE)
							.addGap(23))))
		);
		ArrayList<String> selectStrings = getZipcode();
		JComboBox comboBox = new JComboBox(selectStrings.toArray(new String[selectStrings.size()]));
		comboBox.setSelectedIndex(0);
		comboBox.addActionListener(new java.awt.event.ActionListener() {
	            public void actionPerformed(java.awt.event.ActionEvent evt) {
	                try {
						comboBoxActionPerformed(evt);
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
	            }

	        });
		
		JLabel lblZipCodeTitle = new JLabel("Zip Code:");
		
		lblHouseIdTitle = new JLabel("House Id:");
		
		lblImageTypeTitle = new JLabel("Image Type:");
		
		lblHouseZipId = new JLabel(" ");
		lblHouseZipId.setBackground(Color.WHITE);
		
		lblImageType = new JLabel(" ");
		lblImageType.setBackground(Color.WHITE);
		GroupLayout gl_imageInfo = new GroupLayout(imageInfo);
		gl_imageInfo.setHorizontalGroup(
			gl_imageInfo.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_imageInfo.createSequentialGroup()
					.addGap(39)
					.addGroup(gl_imageInfo.createParallelGroup(Alignment.TRAILING, false)
						.addGroup(gl_imageInfo.createSequentialGroup()
							.addComponent(lblImageTypeTitle)
							.addGap(16)
							.addComponent(lblImageType, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
						.addGroup(gl_imageInfo.createSequentialGroup()
							.addComponent(lblHouseIdTitle)
							.addGap(30)
							.addComponent(lblHouseZipId, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
						.addGroup(gl_imageInfo.createSequentialGroup()
							.addComponent(lblZipCodeTitle)
							.addGap(31)
							.addComponent(comboBox, GroupLayout.PREFERRED_SIZE, 98, GroupLayout.PREFERRED_SIZE)))
					.addContainerGap(141, Short.MAX_VALUE))
		);
		gl_imageInfo.setVerticalGroup(
			gl_imageInfo.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_imageInfo.createSequentialGroup()
					.addGroup(gl_imageInfo.createParallelGroup(Alignment.BASELINE)
						.addComponent(lblZipCodeTitle)
						.addComponent(comboBox, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
					.addPreferredGap(ComponentPlacement.RELATED)
					.addGroup(gl_imageInfo.createParallelGroup(Alignment.BASELINE)
						.addComponent(lblHouseIdTitle)
						.addComponent(lblHouseZipId))
					.addPreferredGap(ComponentPlacement.RELATED, 11, Short.MAX_VALUE)
					.addGroup(gl_imageInfo.createParallelGroup(Alignment.BASELINE)
						.addComponent(lblImageTypeTitle)
						.addComponent(lblImageType))
					.addContainerGap())
		);
		imageInfo.setLayout(gl_imageInfo);
		
		JButton btnPrevious = new JButton("Previous");
		btnPrevious.setBackground(new Color(192, 192, 192));
		btnPrevious.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                try {
                	btnPreviousActionPerformed(evt);
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
            }
        });
		
		btnNext = new JButton("Next");
		btnNext.setBackground(new Color(192, 192, 192));
		btnNext.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                try {
					btnNextActionPerformed(evt);
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
            }
        });
		GroupLayout gl_buttonPanel = new GroupLayout(buttonPanel);
		gl_buttonPanel.setHorizontalGroup(
			gl_buttonPanel.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_buttonPanel.createSequentialGroup()
					.addGap(35)
					.addComponent(btnPrevious, GroupLayout.PREFERRED_SIZE, 101, GroupLayout.PREFERRED_SIZE)
					.addPreferredGap(ComponentPlacement.RELATED, 304, Short.MAX_VALUE)
					.addComponent(btnNext, GroupLayout.PREFERRED_SIZE, 101, GroupLayout.PREFERRED_SIZE)
					.addGap(39))
		);
		gl_buttonPanel.setVerticalGroup(
			gl_buttonPanel.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_buttonPanel.createSequentialGroup()
					.addContainerGap()
					.addGroup(gl_buttonPanel.createParallelGroup(Alignment.BASELINE)
						.addComponent(btnPrevious, GroupLayout.DEFAULT_SIZE, 41, Short.MAX_VALUE)
						.addComponent(btnNext, GroupLayout.PREFERRED_SIZE, 41, GroupLayout.PREFERRED_SIZE))
					.addContainerGap())
		);
		buttonPanel.setLayout(gl_buttonPanel);
		
		JButton btnSubmit = new JButton("Submit");
		btnSubmit.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                try {
					btnSubmitActionPerformed(evt);
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (EncryptedDocumentException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (InvalidFormatException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
            }
        });
		
		lblHouseOutfit = new JLabel("1.House Outward");
		
		sliderOutfit = new JSlider(S_MIN, S_MAX, S_INIT);
		sliderOutfit.addChangeListener(new ChangeListener() {

            public void stateChanged(ChangeEvent event) {
                int currentTime = ((JSlider)event.getSource()).getValue();
                lblHOV.setText(Integer.toString(currentTime));
            }
        });
		sliderRestroom = new JSlider(S_MIN, S_MAX, S_INIT);
		sliderRestroom.addChangeListener(new ChangeListener() {

            public void stateChanged(ChangeEvent event) {
                int currentTime = ((JSlider)event.getSource()).getValue();
                lblR.setText(Integer.toString(currentTime));
            }
        });
		lblHOV = new JLabel("0");
		
		lblR = new JLabel("0");
		
		lblG = new JLabel("0");
		
		lblInfo = new JLabel("<html>Score from 1 to 10. <br>If there's no image, leave it 0.</html>");
		sliderKitchen = new JSlider(S_MIN, S_MAX, S_INIT);
		sliderKitchen.addChangeListener(new ChangeListener() {

            public void stateChanged(ChangeEvent event) {
                int currentTime = ((JSlider)event.getSource()).getValue();
                lblK.setText(Integer.toString(currentTime));
            }
        });
		
		sliderBedroom = new JSlider(S_MIN, S_MAX, S_INIT);
		sliderBedroom.addChangeListener(new ChangeListener() {

            public void stateChanged(ChangeEvent event) {
                int currentTime = ((JSlider)event.getSource()).getValue();
                lblB.setText(Integer.toString(currentTime));
            }
        });
		
		
		sliderLivingRoom = new JSlider(S_MIN, S_MAX, S_INIT);
		sliderLivingRoom.addChangeListener(new ChangeListener() {

            public void stateChanged(ChangeEvent event) {
                int currentTime = ((JSlider)event.getSource()).getValue();
                lblLR.setText(Integer.toString(currentTime));
            }
        });
		
		lblLivingRoom = new JLabel("2.Living Room");
		
		lblLR = new JLabel("0");
		
		lblBedroom = new JLabel("3.Bedroom");
		
		lblB = new JLabel("0");
		
		lblKitchen = new JLabel("4.Kitchen");
		
		lblK = new JLabel("0");
		
		lblRestroom = new JLabel("5.Restroom");
		sliderGarden = new JSlider(S_MIN, S_MAX, S_INIT);
		sliderGarden.addChangeListener(new ChangeListener() {

            public void stateChanged(ChangeEvent event) {
                int currentTime = ((JSlider)event.getSource()).getValue();
                lblG.setText(Integer.toString(currentTime));
            }
        });
		
		lblGarden = new JLabel("6.Garden");
		sliderOverall = new JSlider(S_MIN, S_MAX, S_INIT);
		sliderOverall.addChangeListener(new ChangeListener() {

            public void stateChanged(ChangeEvent event) {
                int currentTime = ((JSlider)event.getSource()).getValue();
                lblO.setText(Integer.toString(currentTime));
            }
        });
		
		lblOverallEvaluation = new JLabel("Overall Evaluate");
		lblOverallEvaluation.setFont(new Font("Tahoma", Font.BOLD, 14));
		
		lblO = new JLabel("0");
		
		btnNextHouse = new JButton("Next House");
		btnNextHouse.setBackground(UIManager.getColor("Button.background"));
		btnNextHouse.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                try {
					btnNextHouseActionPerformed(evt);
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
            }
        });
		
		btnPreHouse = new JButton("Pre House");
		btnPreHouse.setBackground(UIManager.getColor("Button.background"));
		btnPreHouse.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                try {
					btnPreHouseActionPerformed(evt);
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
            }
        });
		GroupLayout gl_evaluatePanel = new GroupLayout(evaluatePanel);
		gl_evaluatePanel.setHorizontalGroup(
			gl_evaluatePanel.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_evaluatePanel.createSequentialGroup()
					.addContainerGap()
					.addGroup(gl_evaluatePanel.createParallelGroup(Alignment.LEADING)
						.addComponent(lblInfo, Alignment.TRAILING, GroupLayout.PREFERRED_SIZE, 271, GroupLayout.PREFERRED_SIZE)
						.addGroup(Alignment.TRAILING, gl_evaluatePanel.createSequentialGroup()
							.addGroup(gl_evaluatePanel.createParallelGroup(Alignment.LEADING)
								.addComponent(lblHouseOutfit, GroupLayout.PREFERRED_SIZE, 97, GroupLayout.PREFERRED_SIZE)
								.addComponent(lblLivingRoom, GroupLayout.PREFERRED_SIZE, 97, GroupLayout.PREFERRED_SIZE)
								.addComponent(lblBedroom, GroupLayout.PREFERRED_SIZE, 97, GroupLayout.PREFERRED_SIZE)
								.addComponent(lblKitchen, GroupLayout.PREFERRED_SIZE, 97, GroupLayout.PREFERRED_SIZE)
								.addComponent(lblRestroom, GroupLayout.PREFERRED_SIZE, 97, GroupLayout.PREFERRED_SIZE)
								.addComponent(lblGarden, GroupLayout.PREFERRED_SIZE, 97, GroupLayout.PREFERRED_SIZE)
								.addComponent(lblOverallEvaluation, GroupLayout.DEFAULT_SIZE, 181, Short.MAX_VALUE))
							.addPreferredGap(ComponentPlacement.RELATED)
							.addGroup(gl_evaluatePanel.createParallelGroup(Alignment.LEADING)
								.addGroup(gl_evaluatePanel.createSequentialGroup()
									.addComponent(sliderOutfit, GroupLayout.PREFERRED_SIZE, 97, GroupLayout.PREFERRED_SIZE)
									.addComponent(lblHOV, GroupLayout.PREFERRED_SIZE, 97, GroupLayout.PREFERRED_SIZE))
								.addGroup(gl_evaluatePanel.createSequentialGroup()
									.addComponent(sliderLivingRoom, GroupLayout.PREFERRED_SIZE, 97, GroupLayout.PREFERRED_SIZE)
									.addComponent(lblLR, GroupLayout.PREFERRED_SIZE, 97, GroupLayout.PREFERRED_SIZE))
								.addGroup(gl_evaluatePanel.createSequentialGroup()
									.addComponent(sliderBedroom, GroupLayout.PREFERRED_SIZE, 97, GroupLayout.PREFERRED_SIZE)
									.addComponent(lblB, GroupLayout.PREFERRED_SIZE, 97, GroupLayout.PREFERRED_SIZE))
								.addGroup(gl_evaluatePanel.createSequentialGroup()
									.addComponent(sliderKitchen, GroupLayout.PREFERRED_SIZE, 97, GroupLayout.PREFERRED_SIZE)
									.addComponent(lblK, GroupLayout.PREFERRED_SIZE, 97, GroupLayout.PREFERRED_SIZE))
								.addGroup(gl_evaluatePanel.createSequentialGroup()
									.addComponent(sliderRestroom, GroupLayout.PREFERRED_SIZE, 97, GroupLayout.PREFERRED_SIZE)
									.addComponent(lblR, GroupLayout.PREFERRED_SIZE, 97, GroupLayout.PREFERRED_SIZE))
								.addGroup(gl_evaluatePanel.createSequentialGroup()
									.addComponent(sliderGarden, GroupLayout.PREFERRED_SIZE, 97, GroupLayout.PREFERRED_SIZE)
									.addComponent(lblG, GroupLayout.PREFERRED_SIZE, 97, GroupLayout.PREFERRED_SIZE))
								.addGroup(gl_evaluatePanel.createSequentialGroup()
									.addComponent(sliderOverall, GroupLayout.PREFERRED_SIZE, 97, GroupLayout.PREFERRED_SIZE)
									.addComponent(lblO, GroupLayout.PREFERRED_SIZE, 97, GroupLayout.PREFERRED_SIZE))))
						.addGroup(Alignment.TRAILING, gl_evaluatePanel.createSequentialGroup()
							.addComponent(btnPreHouse)
							.addGap(33)
							.addComponent(btnSubmit, GroupLayout.PREFERRED_SIZE, 135, GroupLayout.PREFERRED_SIZE)
							.addPreferredGap(ComponentPlacement.RELATED, 31, Short.MAX_VALUE)
							.addComponent(btnNextHouse, GroupLayout.PREFERRED_SIZE, 99, GroupLayout.PREFERRED_SIZE)))
					.addContainerGap())
		);
		gl_evaluatePanel.setVerticalGroup(
			gl_evaluatePanel.createParallelGroup(Alignment.TRAILING)
				.addGroup(Alignment.LEADING, gl_evaluatePanel.createSequentialGroup()
					.addComponent(lblInfo, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE)
					.addGap(18)
					.addGroup(gl_evaluatePanel.createParallelGroup(Alignment.LEADING)
						.addGroup(gl_evaluatePanel.createSequentialGroup()
							.addGroup(gl_evaluatePanel.createParallelGroup(Alignment.LEADING)
								.addComponent(sliderOutfit, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE)
								.addComponent(lblHOV, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE))
							.addGroup(gl_evaluatePanel.createParallelGroup(Alignment.LEADING)
								.addComponent(sliderLivingRoom, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE)
								.addComponent(lblLR, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE))
							.addGroup(gl_evaluatePanel.createParallelGroup(Alignment.LEADING)
								.addComponent(sliderBedroom, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE)
								.addComponent(lblB, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE))
							.addGroup(gl_evaluatePanel.createParallelGroup(Alignment.LEADING)
								.addComponent(sliderKitchen, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE)
								.addComponent(lblK, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE))
							.addGroup(gl_evaluatePanel.createParallelGroup(Alignment.LEADING)
								.addComponent(sliderRestroom, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE)
								.addComponent(lblR, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE))
							.addGroup(gl_evaluatePanel.createParallelGroup(Alignment.LEADING)
								.addComponent(sliderGarden, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE)
								.addComponent(lblG, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE)))
						.addGroup(gl_evaluatePanel.createSequentialGroup()
							.addComponent(lblHouseOutfit, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE)
							.addComponent(lblLivingRoom, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE)
							.addComponent(lblBedroom, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE)
							.addComponent(lblKitchen, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE)
							.addComponent(lblRestroom, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE)
							.addComponent(lblGarden, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE)))
					.addPreferredGap(ComponentPlacement.RELATED)
					.addGroup(gl_evaluatePanel.createParallelGroup(Alignment.LEADING)
						.addComponent(lblOverallEvaluation, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE)
						.addComponent(sliderOverall, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE)
						.addComponent(lblO, GroupLayout.PREFERRED_SIZE, 43, GroupLayout.PREFERRED_SIZE))
					.addGap(18)
					.addGroup(gl_evaluatePanel.createParallelGroup(Alignment.TRAILING)
						.addComponent(btnSubmit, GroupLayout.PREFERRED_SIZE, 32, GroupLayout.PREFERRED_SIZE)
						.addComponent(btnPreHouse)
						.addComponent(btnNextHouse))
					.addContainerGap(34, Short.MAX_VALUE))
		);
		evaluatePanel.setLayout(gl_evaluatePanel);
		
		jLabel_Image = new JLabel("");
		GroupLayout gl_imageDisplayPanel = new GroupLayout(imageDisplayPanel);
		gl_imageDisplayPanel.setHorizontalGroup(
			gl_imageDisplayPanel.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_imageDisplayPanel.createSequentialGroup()
					.addContainerGap()
					.addComponent(jLabel_Image, GroupLayout.DEFAULT_SIZE, 549, Short.MAX_VALUE)
					.addContainerGap())
		);
		gl_imageDisplayPanel.setVerticalGroup(
			gl_imageDisplayPanel.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_imageDisplayPanel.createSequentialGroup()
					.addContainerGap()
					.addComponent(jLabel_Image, GroupLayout.DEFAULT_SIZE, 400, Short.MAX_VALUE)
					.addContainerGap())
		);
		imageDisplayPanel.setLayout(gl_imageDisplayPanel);
		contentPane.setLayout(gl_contentPane);
		pack();
		setLocationRelativeTo(null);
	    setVisible(true);
	}
	
	/**************************Action Classes************************/
	
	private void comboBoxActionPerformed(ActionEvent evt) throws IOException {
		JComboBox cb = (JComboBox)evt.getSource();
		zipcode = Integer.valueOf((String)cb.getSelectedItem());
		getZipcodeFolder();
		ArrayList<String> temp = getHouseInZipcode();
		houseInZipcode = temp.toArray(new String[temp.size()]);
		houseIndex = 0;
		houseImageIndex = 0;
		String  imageUrl = houseInZipcode[houseIndex] + "\\" + houseImageIndex + ".jpg";
		File f = new File(imageUrl);
    	if(f.exists() && !f.isDirectory()) {
    		showImage();
    		changeText();
    		changeSliderValue();
		}else{
			Component frameD = null;
			JOptionPane.showMessageDialog(frameD, "No Image for this house. Please click SUBMIT button to see other houses!");
		}
	}
	

	private void btnSubmitActionPerformed(java.awt.event.ActionEvent evt) throws IOException, EncryptedDocumentException, InvalidFormatException {                                                 
		saveData();
		houseIndex ++;
		houseImageIndex = 0;
		if(houseIndex<houseInZipcode.length){
			showImage();
			changeText();
		}else{
			houseIndex--;
			Component frameD = null;
			JOptionPane.showMessageDialog(frameD, "No more houses in this zip code area, choos another zipcode!");
		}
	}  
	
	//Get the next image of the house
	private void btnNextActionPerformed(ActionEvent evt) throws IOException {
		houseImageIndex++;
		int count = 1;
		for(; houseImageIndex<6; houseImageIndex++){
			String  imageUrl = houseInZipcode[houseIndex] + "\\" + houseImageIndex + ".jpg";
			File f = new File(imageUrl);
			if(f.exists() && !f.isDirectory()) {
	    		showImage();
	    		changeText();
	    		return;
			}else {
				count++;
				continue;
			}
		}
		houseImageIndex-=count;
		Component frameD = null;
		JOptionPane.showMessageDialog(frameD, "No more Images for this house!");
	}
	
	//Get the previous image of the house
	private void btnPreviousActionPerformed(ActionEvent evt) throws IOException {
		if(houseImageIndex>0){
			houseImageIndex--;
		}else{
			Component frameD = null;
			JOptionPane.showMessageDialog(frameD, "This is the first image!");
		}
		showImage();
		changeText();
	}
	
	//Get the next house
	private void btnNextHouseActionPerformed(ActionEvent evt) throws IOException {
		houseIndex++;
		houseImageIndex = 0;
		if(houseIndex<houseInZipcode.length){
			showImage();
			changeText();
			changeSliderValue();
		}else{
			houseIndex--;
			Component frameD = null;
			JOptionPane.showMessageDialog(frameD, "No more houses in this zip code area!");
		}
	}
	
	//Get the previous house
	private void btnPreHouseActionPerformed(ActionEvent evt) throws IOException {
		houseIndex--;
		houseImageIndex = 0;
		if(houseIndex>0){
			showImage();
			changeText();
			changeSliderValue();
		}else{
			houseIndex++;
			Component frameD = null;
			JOptionPane.showMessageDialog(frameD, "This is the first house in this zip code area!");
		}
	}
	
	/*************************Private Methods************************/
	
	//Save the evaluation to the data set
	private void saveData() throws EncryptedDocumentException, InvalidFormatException, IOException {
		try{
            InputStream file = new FileInputStream(pathHouseData);
            Workbook workbook = WorkbookFactory.create(file);
            Sheet sheet = workbook.getSheetAt(0);
            Cell cellZipcode = null;
            
            //Iterator<Row> rowIterator = sheet.iterator(); // Traversing over each row of XLSX file 
        	//ArrayList<String> imageList = new ArrayList<>();
        	for(Row row:sheet) { 
        		//Row row = rowIterator.next(); // For each row, iterate through each columns 
        		Cell zipid = row.getCell(0);
        		Cell zip = row.getCell(3);
        		if (zipid == null||zip == null)
        		{
        		   //System.out.println("Cell is empty");
        		   continue;

        		}
        		
        		//System.out.print(zipid.getNumericCellValue() +"\n");
        		
        		if((int)zip.getNumericCellValue()==zipcode){
        			if((int)zipid.getNumericCellValue()==Integer.valueOf((houseInZipcode[houseIndex].substring(houseInZipcode[houseIndex].lastIndexOf("\\") + 1)))){
        				row.createCell(8).setCellValue(Integer.valueOf(lblHOV.getText()));//Outfit
        	            row.createCell(9).setCellValue(Integer.valueOf(lblLR.getText()));//Living Room
        	            row.createCell(10).setCellValue(Integer.valueOf(lblB.getText()));
        	            row.createCell(11).setCellValue(Integer.valueOf(lblK.getText()));
        	            row.createCell(12).setCellValue(Integer.valueOf(lblR.getText()));
        	            row.createCell(13).setCellValue(Integer.valueOf(lblG.getText()));
        	            row.createCell(14).setCellValue(Integer.valueOf(lblO.getText()));
        	            FileOutputStream outFile =new FileOutputStream(pathHouseData);
        	            workbook.write(outFile);
        	            outFile.close();
        	            System.out.print("The evaluation of House:"+Integer.valueOf((int) zipid.getNumericCellValue()) +" is saved.\n");
        	            break;
        			}
        		}
        		
        		    //System.out.print(zipidFolder +"\n");
        	}		
            
		} catch(Exception e){
			Component frameD = null;
			JOptionPane.showMessageDialog(frameD, "Fail to save the data!");
		}
	}
	
	//Change the show text of the house zipid label and the ImageType label
	private void changeText(){
		lblHouseZipId.setText(houseInZipcode[houseIndex].substring(houseInZipcode[houseIndex].lastIndexOf("\\") + 1));
		String houseType = null;
		switch(houseImageIndex){
		case 0: houseType = "House Outfit";
				break;
		case 1: houseType = "Living Room";
				break;
		case 2: houseType = "Bedroom";
				break;
		case 3: houseType = "Kitchen";
				break;
		case 4: houseType = "Restroom";
				break;
		case 5: houseType = "Garden";
				break;
		default:
			break;
		}
		lblImageType.setText(houseType);
	}
	
	//Change the slider value when change the house
	private void changeSliderValue() throws IOException {
		int prezipid = Integer.valueOf(lblHouseZipId.getText());
		InputStream fis = new FileInputStream(pathHouseData); // Finds the workbook instance for XLSX file 
    	XSSFWorkbook myWorkBook = new XSSFWorkbook (fis); // Return first sheet from the XLSX workbook 
    	XSSFSheet mySheet = myWorkBook.getSheetAt(0); // Get iterator to all the rows in current sheet 
    	Iterator<Row> rowIterator = mySheet.iterator(); // Traversing over each row of XLSX file 
    	ArrayList<String> zipcodes = new ArrayList<>();
    	while (rowIterator.hasNext()) { 
    		Row row = rowIterator.next(); // For each row, iterate through each columns 
    		Cell zipid = row.getCell(0);
    		if (zipid == null)
    		{
    		   continue;
    		}
    		if((int)zipid.getNumericCellValue()==prezipid){
    			
    			if(row.getCell(8)!=null)//||row.getCell(9)!=null||row.getCell(10)!=null||row.getCell(11)!=null||row.getCell(12)!=null||row.getCell(13)!=null||row.getCell(14)!=null
    			{
    				sliderOutfit.setValue((int)row.getCell(8).getNumericCellValue());
    				sliderLivingRoom.setValue((int)row.getCell(9).getNumericCellValue());
        			sliderBedroom.setValue((int)row.getCell(10).getNumericCellValue());
        			sliderKitchen.setValue((int)row.getCell(11).getNumericCellValue());
        			sliderRestroom.setValue((int)row.getCell(12).getNumericCellValue());
        			sliderGarden.setValue((int)row.getCell(13).getNumericCellValue());
        			sliderOverall.setValue((int)row.getCell(14).getNumericCellValue());
    			}else{
    				sliderOutfit.setValue(0);
    				sliderLivingRoom.setValue(0);
    				sliderBedroom.setValue(0);
    				sliderKitchen.setValue(0);
    				sliderRestroom.setValue(0);
    				sliderGarden.setValue(0);
    				sliderOverall.setValue(0);
    			}
    			break;
    		}
    	}
		
	}
	
	//Get zip codes from zip code data set.
	private ArrayList<String> getZipcode() throws IOException {
    	InputStream fis = new FileInputStream(pathZipcode); // Finds the workbook instance for XLSX file 
    	XSSFWorkbook myWorkBook = new XSSFWorkbook (fis); // Return first sheet from the XLSX workbook 
    	XSSFSheet mySheet = myWorkBook.getSheetAt(0); // Get iterator to all the rows in current sheet 
    	Iterator<Row> rowIterator = mySheet.iterator(); // Traversing over each row of XLSX file 
    	ArrayList<String> zipcodes = new ArrayList<>();
    	while (rowIterator.hasNext()) { 
    		Row row = rowIterator.next(); // For each row, iterate through each columns 
    		Cell zipcode = row.getCell(0);
    		zipcodes.add(String.valueOf((int)zipcode.getNumericCellValue()));
    	}
		return zipcodes;
	}
	
	//Get the folder of the specific zip code area, which contents the images of houses.
	public Void getZipcodeFolder() throws IOException{
		url = pathImageRoot + zipcode + "\\";
		return null;
	}
	
	//Get the list of houses in one specific zip code area
	public ArrayList<String> getHouseInZipcode() throws IOException
    {
		InputStream fis = new FileInputStream(pathHouseData); // Finds the workbook instance for XLSX file 
    	XSSFWorkbook myWorkBook = new XSSFWorkbook (fis); // Return first sheet from the XLSX workbook 
    	XSSFSheet mySheet = myWorkBook.getSheetAt(0); // Get iterator to all the rows in current sheet 
    	Iterator<Row> rowIterator = mySheet.iterator(); // Traversing over each row of XLSX file 
    	ArrayList<String> imageList = new ArrayList<>();
    	while (rowIterator.hasNext()) { 
    		Row row = rowIterator.next(); // For each row, iterate through each columns
    		
    		Cell zipid = row.getCell(0);
    		Cell zip = row.getCell(3);
    		if (zipid == null||zip == null)
    		{
    		   continue;

    		}
    		if((int)zip.getNumericCellValue()==zipcode){
    			String zipidFolder = url + String.valueOf((int)zipid.getNumericCellValue());
    		    imageList.add(zipidFolder);   
    		    //System.out.print(zipidFolder +"\n");
    		}
    	}
        return imageList;
    }

	//Display the image of one house.
    public void showImage() throws IOException
    {
    	String  imageUrl = houseInZipcode[houseIndex] + "\\" + houseImageIndex + ".jpg";
    	File f = new File(imageUrl);
    	if(f.exists() && !f.isDirectory()) { 
            ImageIcon icon = new ImageIcon(imageUrl);
            Image image = icon.getImage().getScaledInstance(jLabel_Image.getWidth(), jLabel_Image.getHeight(), Image.SCALE_SMOOTH);
            jLabel_Image.setIcon(new ImageIcon(image));
    	}
    	
    }
}
