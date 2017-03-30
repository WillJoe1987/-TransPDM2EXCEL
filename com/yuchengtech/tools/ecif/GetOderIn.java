package com.yuchengtech.tools.ecif;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.dom4j.DocumentException;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class GetOderIn {
	public void write(String xlsFile,String outFile) throws WriteException, DocumentException, IOException, BiffException {
		File infile = new File(xlsFile);
		Workbook inworkbook=Workbook.getWorkbook(infile);
		
		File outfile = new File(outFile);
		System.out.println("EXCEL File path is : [" + outfile.getAbsolutePath() + "]");
		File dir = outfile.getParentFile();
		if (!dir.exists()) {
			//�����ϲ�Ŀ¼
			dir.mkdirs();
		}
		WritableWorkbook wwb = getOdered(inworkbook,outfile);
	
	
	}
	private WritableWorkbook getOdered(Workbook inWork ,File outFile) throws RowsExceededException, WriteException, IOException{
		Sheet sh1 = inWork.getSheet(0);
		Sheet sh2 = inWork.getSheet(1);
		int rows = sh2.getRows();
		WritableWorkbook wbook = Workbook.createWorkbook(outFile);
		WritableSheet outSheet = wbook.createSheet("d", 0);
		
		List<MyRowBeen> mrbList = new ArrayList<MyRowBeen>();
		for(int i=0;i<rows;i++){
			MyRowBeen mrb = new MyRowBeen();
			mrb.setInWord(sh2.getCell(0, i).getContents());
			mrb.setOutName(sh2.getCell(4, i).getContents());
			mrbList.add(mrb);
		}
		int ListSize = mrbList.size();
		int sh1Rows = sh1.getRows();
		int pk=0;
		for(int i=0;i<ListSize;i++){
			MyRowBeen mrb = mrbList.get(i);
			for(int j=0;j<sh1Rows;j++){
				if("OPC".equalsIgnoreCase(sh1.getCell(0, j).getContents())){
					if("".equalsIgnoreCase((String)sh1.getCell(3, j).getContents())){
					if(mrb.getInWord().equalsIgnoreCase(sh1.getCell(4, j).getContents())){
						pk++;
						System.out.println(pk+"   "+sh1.getCell(0, j).getContents()+"	"+sh1.getCell(3, j).getContents()+"    "+mrb.getInWord()+"      "+mrb.getOutName());
						
						outSheet.addCell(new Label(0, i, (String)sh1.getCell(0, j).getContents()));
						outSheet.addCell(new Label(1, i, (String)sh1.getCell(3, j).getContents()));
						outSheet.addCell(new Label(2, i,mrb.getInWord()));
						outSheet.addCell(new Label(3, i, mrb.getOutName()));
					}}
				}		
			}
		}
		System.out.println(ListSize);
		wbook.write();	    
		   // wbookPermi.write();
		wbook.close();
		   // wbookPermi.close();
		wbook = null;
		return wbook;
	}
	public static void main(String[] args) throws WriteException, BiffException, DocumentException, IOException {
		String file = "C:/DALCSP1.xls";
		String file2 = "C:/DALCSP2.xls";
		GetOderIn goi = new GetOderIn();
		goi.write(file,file2);
	}
	
}
