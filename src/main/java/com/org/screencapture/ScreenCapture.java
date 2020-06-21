package com.org.screencapture;

import java.awt.EventQueue;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;

import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JButton;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.awt.event.KeyListener;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.awt.event.ActionEvent;
import javax.swing.ImageIcon;
import javax.imageio.ImageIO;
import javax.swing.GroupLayout;
import javax.swing.GroupLayout.Alignment;
import javax.swing.LayoutStyle.ComponentPlacement;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import com.itextpdf.io.image.ImageData;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.border.Border;
import com.itextpdf.layout.element.Cell;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.Table;

import java.awt.Color;

public class ScreenCapture{

	private JFrame frmScreencapture;
	private JButton play,stop,capture;
	private File createFolder;
	private Rectangle rectangle;
	private XWPFDocument document;
	private XWPFParagraph paragraph;
	private XWPFRun run;
	private BufferedImage bufferedImage;
	private String foldername=null;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					ScreenCapture window = new ScreenCapture();
					window.frmScreencapture.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public ScreenCapture() {
		initialize();

	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frmScreencapture = new JFrame();
		frmScreencapture.setTitle("ScreenCapture");
		frmScreencapture.setBounds(100, 100, 513, 198);
		frmScreencapture.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frmScreencapture.setResizable(false);
		frmScreencapture.addKeyListener(new KeyListener() {
			
			@Override
			public void keyTyped(KeyEvent e) {
				// TODO Auto-generated method stub
				
			}
			
			@Override
			public void keyReleased(KeyEvent e) {
				// TODO Auto-generated method stub
				if(e.getKeyCode()==KeyEvent.VK_PRINTSCREEN)
				{
					capture();
				}
				
			}
			
			@Override
			public void keyPressed(KeyEvent e) {
				// TODO Auto-generated method stub
				
			}
		});
		frmScreencapture.setFocusable(true);
		frmScreencapture.setFocusTraversalKeysEnabled(false);
	
		play = new JButton("");
		play.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				start();
			}
		});
		play.setIcon(new ImageIcon(ScreenCapture.class.getResource("/com/org/screencapture/play.png")));

		stop = new JButton("");
		stop.setEnabled(false);
		stop.setForeground(Color.BLACK);
		stop.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try 
				{
					createPDFDoc();
					createWordDoc();
					fileaction();
				} 
				catch (Exception e1) 
				{
					e1.printStackTrace();
				}
			}
		});
		stop.setIcon(new ImageIcon(ScreenCapture.class.getResource("/com/org/screencapture/stop.png")));

		capture = new JButton("");
		capture.setEnabled(false); 
		capture.setIcon(new ImageIcon(ScreenCapture.class.getResource("/com/org/screencapture/capture.png")));
		capture.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				capture();
			}
		});

		GroupLayout groupLayout = new GroupLayout(frmScreencapture.getContentPane());
		groupLayout.setHorizontalGroup(
				groupLayout.createParallelGroup(Alignment.LEADING)
				.addGroup(groupLayout.createSequentialGroup()
						.addGap(16)
						.addComponent(play, GroupLayout.PREFERRED_SIZE, 149, GroupLayout.PREFERRED_SIZE)
						.addPreferredGap(ComponentPlacement.RELATED)
						.addComponent(stop, GroupLayout.PREFERRED_SIZE, 149, GroupLayout.PREFERRED_SIZE)
						.addPreferredGap(ComponentPlacement.RELATED)
						.addComponent(capture)
						.addGap(34))
				);
		groupLayout.setVerticalGroup(
				groupLayout.createParallelGroup(Alignment.LEADING)
				.addGroup(groupLayout.createSequentialGroup()
						.addContainerGap()
						.addGroup(groupLayout.createParallelGroup(Alignment.LEADING)
								.addGroup(groupLayout.createSequentialGroup()
										.addComponent(play, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
										.addContainerGap())
								.addGroup(groupLayout.createSequentialGroup()
										.addComponent(stop, GroupLayout.DEFAULT_SIZE, 150, Short.MAX_VALUE)
										.addContainerGap())
								.addGroup(groupLayout.createSequentialGroup()
										.addComponent(capture, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
										.addContainerGap())))
				);
		frmScreencapture.getContentPane().setLayout(groupLayout);
	}

	private void start()
	{
		foldername="/Screenshots_"+ new SimpleDateFormat("yyyyMMdd-HHmmss").format(new Date());
		createFolder= new File(foldername);
		createFolder.mkdir();
		play.setEnabled(false);
		capture.setEnabled(true); 
		frmScreencapture.setAlwaysOnTop(true);	
		JOptionPane.showMessageDialog(null, "Use print screen key to capture screens","Screen Capture",JOptionPane.INFORMATION_MESSAGE);
	}
	private void capture()
	{
		stop.setEnabled(true);
		if(createFolder.exists() && createFolder.isDirectory())
		{
			try
			{
				frmScreencapture.setVisible(false);
				Thread.sleep(500);
				Robot robot = new Robot();
				rectangle = new Rectangle(Toolkit.getDefaultToolkit().getScreenSize());
				bufferedImage = robot.createScreenCapture(rectangle);
				String fileSuffix = new SimpleDateFormat("yyyyMMdd-HHmmss").format(new Date());
				File createImg=new File(foldername +"/" + fileSuffix + ".png");
				ImageIO.write(bufferedImage,"png",createImg);
				frmScreencapture.setVisible(true);
			}
			catch(Exception e)
			{
				System.err.println(e);
			}
		}
		else
		{
			JOptionPane.showMessageDialog(null, "Screen Capture","Warning Screenshot folder does not exist",JOptionPane.WARNING_MESSAGE);
		}

	}
	private void createWordDoc() throws IOException, InvalidFormatException
	{
		document = new XWPFDocument(); 		
		String docname = "Screenshot-" + new SimpleDateFormat("yyyyMMddHHmmss").format(new Date());
		FileOutputStream out = new FileOutputStream(new File(foldername +"/" + docname + ".docx"));
		File[] files = createFolder.listFiles();	
		paragraph=document.createParagraph();
		run = paragraph.createRun();
		int picformat=org.apache.poi.xwpf.usermodel.Document.PICTURE_TYPE_PNG;
		for (File file : files)
		{
			if(FilenameUtils.getExtension(file.getName()).equalsIgnoreCase("png"))
			{
				FileInputStream fis=new FileInputStream(file);
				run.addPicture(fis, picformat, file.getName(), Units.toEMU(470), Units.toEMU(300));	
				run.addBreak(BreakType.TEXT_WRAPPING);
				fis.close();
			}
		}
		document.write(out);
		out.close();
		document.close();		
	}
	private void createPDFDoc() throws IOException
	{
		String docname = "Screenshot-" + new SimpleDateFormat("yyyyMMddHHmmss").format(new Date());
		PdfWriter writer=new PdfWriter(foldername +"/"  + docname + ".pdf");
		PdfDocument pdfdoc=new PdfDocument(writer);
		pdfdoc.addNewPage();
		Document pdfdocument=new Document(pdfdoc);
		File[] imgfiles = createFolder.listFiles();
		@SuppressWarnings("deprecation")
		Table table= new Table(1);
		table.setBorder(Border.NO_BORDER);
		for (File file : imgfiles)
		{
			if(FilenameUtils.getExtension(file.getName()).equalsIgnoreCase("png"))
			{     
				Cell cell=new Cell();
				cell.setBorder(Border.NO_BORDER);
				ImageData data = ImageDataFactory.create(file.getPath());                     
				Image image = new Image(data);  
				image.setWidthPercent(100);
				cell.add(image);
				table.addCell(cell);
			}
		}
		pdfdocument.add(table);
		pdfdocument.close(); 	
		writer.close();
	}
	private void fileaction()
	{
		int option=JOptionPane.showConfirmDialog(null, "Do you want to keep raw image files captured?","Keep Files Warning",JOptionPane.YES_NO_OPTION);
		if(option==JOptionPane.YES_OPTION)
		{	
			JOptionPane.showMessageDialog(null, "Process Completed and Output file saved as PDF and Doc");
			frmScreencapture.dispose();
		}
		else
		{
			File[] srcimage=createFolder.listFiles();
			for(File imgfile:srcimage)
			{
				if(FilenameUtils.getExtension(imgfile.getName()).equalsIgnoreCase("png"))
				{
					String path=imgfile.getAbsolutePath();
					File delfile=new File(path);
					delfile.delete();
				}					
			}
			JOptionPane.showMessageDialog(null, "Process Completed and Output file saved as PDF and Doc with raw Image files deleted");
			frmScreencapture.dispose();
		}
	}
}
