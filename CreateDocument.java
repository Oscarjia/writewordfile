package com.writeofficeword;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import org.apache.poi.xwpf.usermodel.Borders;
import org.apache.poi.xwpf.usermodel.BreakClear;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.LineSpacingRule;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.VerticalAlign;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class CreateDocument {
	static Logger logger=LoggerFactory.getLogger(CreateDocument.class);
	static XWPFDocument doc = new XWPFDocument();	
	public static void main(String[] args) {
		try{
			Stream<Path> paths = Files.walk(Paths.get("E:\\work\\gitproject\\"));//folder that to read
			List<File> filesInFolder= paths.filter(Files::isRegularFile).map(Path::toFile).collect(Collectors.toList());
		   
		    for (File file:filesInFolder) {
		    	logger.info(file.toString());
				String line="";
		        //FileReader fileReader = 
		        //    new FileReader(file);
		        InputStreamReader isr = new InputStreamReader(new FileInputStream(file), "GBK");
		        // Always wrap FileReader in BufferedReader.
		        BufferedReader bufferedReader = 
		            new BufferedReader(isr);
		        
		        XWPFParagraph p2= doc.createParagraph();
				XWPFRun r2 = p2.createRun(); 
				r2.setText(file.toString()+";");
				r2.addCarriageReturn();
				
		        XWPFParagraph p3 = doc.createParagraph();
				XWPFRun r4 = p3.createRun(); 
		        while((line = bufferedReader.readLine()) != null) {
		        	if(line==null||line.equals("")||line.trim().equals("")){
		        		continue;
		        	}
		        	r4.setText(line.trim());
		        	r4.addCarriageReturn();
		        } 
		        bufferedReader.close();
		        isr.close();
		        
			}
		  
		    try (FileOutputStream out = new FileOutputStream("E:\\work\\gitproject\\write.docx")) {//write to word
			    doc.write(out);
			logger.info("creat document finished");
		
		} 
			
	}catch(Exception e){
		
		logger.error("",e);
		
	}
	}
	
	static void writeword(File file) throws IOException{
		

        

		/*	
		r4.addBreak(BreakType.PAGE);
		r4.setText("No more; and by a sleep to say we end "
		        + "The heart-ache and the thousand natural shocks "
		        + "That flesh is heir to, 'tis a consummation "
		        + "Devoutly to be wish'd. To die, to sleep; "
		        + "To sleep: perchance to dream: ay, there's the rub; "
		        + ".......");
        
//This would imply that this break shall be treated as a simple line break, and break the line after that word:

		XWPFRun r5 = p3.createRun();
         
		r5.setText("For in that sleep of death what dreams may come");
		r5.addCarriageReturn();//
		r5.setText("When we have shuffled off this mortal coil,"
		        + "Must give us pause: there's the respect"
		        + "That makes calamity of so long life;");
		r5.addBreak();
		r5.setText("For who would bear the whips and scorns of time,"
		        + "The oppressor's wrong, the proud man's contumely,");

		r5.addBreak(BreakClear.ALL);
		r5.setText("The pangs of despised love, the law's delay,"
		        + "The insolence of office and the spurns" + ".......");*/

		

}}
