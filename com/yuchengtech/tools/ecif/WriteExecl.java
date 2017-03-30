package com.yuchengtech.tools.ecif;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;

import org.dom4j.DocumentException;

import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableHyperlink;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
/**
 * 
 * @author WillJoe
 * @description ������д��Maping�ļ���Ȩ�޾����ļ�
 * @version 1.0
 * @date 2010-03-23
 */
public class WriteExecl {
	
	public void write(String xlsFile,String perXLS,String pdm) throws WriteException, DocumentException, IOException, BiffException {
		WritableWorkbook wbook = null;
		//WritableWorkbook wbookPermi = null;
		
		
		
		
		File file = new File(xlsFile);
		File fileOfPermi =  new File(perXLS);
		System.out.println("EXCEL File path is : [" + file.getAbsolutePath() + "]");
		File dir = file.getParentFile();
		if (!dir.exists()) {
			//�����ϲ�Ŀ¼
			dir.mkdirs();
		}
		setSheetFmt();
		ReadPDM readpdm = new ReadPDM(pdm);
		List<EcifTableBean> tables = readpdm.getTables();
		//wbookPermi = createPermissionXLS(tables, fileOfPermi);
		wbook = createWorkBook(tables,file);
		
	    wbook.write();
	    
	   // wbookPermi.write();

	    wbook.close();
	   // wbookPermi.close();
	    wbook = null;
	    System.out.println("ALL table has been written!");

	    
	   // wbookPermi = null;
	}
	  /*/ "ģ��"��ϸ��ͷ��ĵ�Ԫ���ʽ
	  WritableCellFormat fmtTitle = new WritableCellFormat();
	  

	  // "ģ��"��ϸ��ͷ��ĵ�Ԫ���ʽ(ר����TABLE CODE�ı���ɫ)
	  WritableCellFormat fmtTitle1 = new WritableCellFormat();

	  // ��ͨ����ĵ�Ԫ���ʽ(�ı�)
	  WritableCellFormat fmtBody = new WritableCellFormat();

	  // ��ͨ����ĵ�Ԫ���ʽ(����)
	  WritableCellFormat fmtBody_int = new WritableCellFormat();

	  // �����ӱ���ĵ�Ԫ���ʽ
	  WritableCellFormat fmtBody_hyper = new WritableCellFormat();

	  // ҳü
	  HeaderFooter header = new HeaderFooter();

	  // "����"��ͷ��ĵ�Ԫ���ʽ
	  WritableCellFormat fmtTitle2 = new WritableCellFormat();
	  
	  
	  */
		/**************Ŀ¼��ͷ��ʽ************/
		WritableCellFormat fmtIndexTitle = new WritableCellFormat();
		WritableCellFormat fmtModuleTitle1 = new WritableCellFormat();
		WritableCellFormat fmtModuleTitle2 = new WritableCellFormat();
		WritableCellFormat fmtModuleTable1 = new WritableCellFormat();
		WritableCellFormat fmtModuleTable2 = new WritableCellFormat();
		WritableCellFormat fmtModuleTable3 = new WritableCellFormat();
		WritableCellFormat fmtTitlePermi = new WritableCellFormat();
	  private void setSheetFmt() throws WriteException {
		    // ��ͷ��ĵ�Ԫ���ʽ
		    // ����
		    WritableFont fontTitle = new WritableFont(WritableFont.createFont("����"),
		        10, WritableFont.NO_BOLD, false);
		    WritableFont fontIndexTitle = new WritableFont(WritableFont.createFont("����"),
			        12, WritableFont.NO_BOLD, false);
		    WritableFont fmtModuleTitle = new WritableFont(WritableFont.createFont("����"),
			        10, WritableFont.BOLD, false);
		    
		    fmtIndexTitle.setFont(fontIndexTitle);
		    fmtIndexTitle.setAlignment(Alignment.CENTRE);
		    fmtIndexTitle.setBorder(Border.ALL, BorderLineStyle.THICK);
		    fmtIndexTitle.setBackground(Colour.GREEN);
		    
		    fmtModuleTitle1.setFont(fmtModuleTitle);
		    fmtModuleTitle1.setAlignment(Alignment.CENTRE);
		    fmtModuleTitle1.setBorder(Border.ALL, BorderLineStyle.THICK);
		    fmtModuleTitle1.setBackground(Colour.GOLD);
		    
		    fmtModuleTitle2.setFont(fmtModuleTitle);
		    fmtModuleTitle2.setAlignment(Alignment.LEFT);
		    fmtModuleTitle2.setBorder(Border.ALL, BorderLineStyle.THICK);
		    fmtModuleTitle2.setBackground(Colour.GOLD);
		    
		    fmtModuleTable1.setFont(fontTitle);
		    fmtModuleTable1.setAlignment(Alignment.LEFT);
		    fmtModuleTable1.setBorder(Border.BOTTOM, BorderLineStyle.THICK);
		    fmtModuleTable1.setBackground(Colour.AQUA);
		    
		    fmtModuleTable2.setFont(fontTitle);
		    fmtModuleTable2.setAlignment(Alignment.LEFT);
		    fmtModuleTable2.setBorder(Border.BOTTOM, BorderLineStyle.THICK);
		    fmtModuleTable2.setBackground(Colour.BROWN);
		    
		    fmtModuleTable3.setFont(fontTitle);
		    fmtModuleTable3.setAlignment(Alignment.LEFT);
		    fmtModuleTable3.setBorder(Border.BOTTOM, BorderLineStyle.THICK);
		    fmtModuleTable3.setBackground(Colour.RED);
		    
		    fmtTitlePermi.setFont(fmtModuleTitle);
		    fmtTitlePermi.setAlignment(Alignment.CENTRE);
		    fmtTitlePermi.setBorder(Border.ALL, BorderLineStyle.THIN);
		    fmtTitlePermi.setBackground(Colour.AQUA);
		    
		    
		    /*		    fmtTitle.setFont(fontTitle);
		    // ˮƽ����
		    fmtTitle.setAlignment(Alignment.LEFT);
		    // ��ֱ����
		    fmtTitle.setVerticalAlignment(VerticalAlignment.CENTRE);
		    // �߿�
		    fmtTitle.setBorder(Border.ALL, BorderLineStyle.THIN);
		    // �Զ�����
		    fmtTitle.setWrap(true);

		    // ��ͷ��ĵ�Ԫ���ʽ
		    // ����
		    WritableFont fontTitle1 = new WritableFont(WritableFont.createFont("����"),
		        10, WritableFont.NO_BOLD, false);
		    fmtTitle1.setFont(fontTitle1);
		    // ˮƽ����
		    fmtTitle1.setAlignment(Alignment.LEFT);
		    // ��ֱ����
		    fmtTitle1.setVerticalAlignment(VerticalAlignment.CENTRE);
		    // �߿�
		    fmtTitle1.setBorder(Border.ALL, BorderLineStyle.THIN);
		    // �Զ�����
		    fmtTitle1.setWrap(true);
		    // ����ɫ
		    fmtTitle1.setBackground(Colour.getInternalColour(35));

		    // ��ͨ����ĵ�Ԫ���ʽ(�ı�)
		    // ����
		    WritableFont fontBody = new WritableFont(WritableFont.createFont("����"), 10,
		        WritableFont.NO_BOLD, false);
		    fmtBody.setFont(fontBody);
		    // ˮƽ����
		    fmtBody.setAlignment(Alignment.LEFT);
		    // ��ֱ����
		    fmtBody.setVerticalAlignment(VerticalAlignment.CENTRE);
		    // �߿�
		    fmtBody.setBorder(Border.ALL, BorderLineStyle.THIN);
		    // �Զ�����
		    fmtBody.setWrap(true);

		    // ��ͨ����ĵ�Ԫ���ʽ(����)
		    // ����
		    WritableFont fontBody_int = new WritableFont(WritableFont.createFont("����"),
		        10, WritableFont.NO_BOLD, false);
		    fmtBody_int.setFont(fontBody_int);
		    // ˮƽ����
		    fmtBody_int.setAlignment(Alignment.RIGHT);
		    // ��ֱ����
		    fmtBody_int.setVerticalAlignment(VerticalAlignment.CENTRE);
		    // �߿�
		    fmtBody_int.setBorder(Border.ALL, BorderLineStyle.THIN);
		    // �Զ�����
		    fmtBody_int.setWrap(true);

		    // �����ӱ���ĵ�Ԫ���ʽ
		    // ����
		    WritableFont fontBody_hyper = new WritableFont(WritableFont
		        .createFont("����"), 10, WritableFont.NO_BOLD, false);
		    fontBody_hyper.setUnderlineStyle(UnderlineStyle.SINGLE);
		    fontBody_hyper.setColour(Colour.YELLOW);
		    fmtBody_hyper.setFont(fontBody_hyper);
		    // ˮƽ����
		    fmtBody_hyper.setAlignment(Alignment.LEFT);
		    // ��ֱ����
		    fmtBody_hyper.setVerticalAlignment(VerticalAlignment.CENTRE);
		    // �߿�
		    fmtBody_hyper.setBorder(Border.ALL, BorderLineStyle.THIN);
		    // �Զ�����
		    fmtBody_hyper.setWrap(true);

		    // ҳü
		    header.getCentre().setFontName("����");
		    header.getCentre().setFontSize(22);
		    header.getCentre().toggleBold();
		    header.getCentre().append("����ģ�ͱ�");
		    header.getRight().setFontName("����");
		    header.getRight().setFontSize(12);
		    header.getRight().append("\012\012��   ҳ����   ҳ");

		    // ����ҳü
		    header.getLeft().clear();
		    header.getLeft().setFontName("����");
		    header.getLeft().setFontSize(12);
		    header.getLeft().toggleBold();
		    header.getLeft().append("\012\012�����ˣ�");
		    header.getLeft().toggleBold();
		    header.getLeft().append("֣����");

		    // ��ͷ��ĵ�Ԫ���ʽ
		    // ����
		    WritableFont fontTitle2 = new WritableFont(WritableFont.createFont("����"),
		        10, WritableFont.BOLD, false);
		    fmtTitle2.setFont(fontTitle2);
		    // ˮƽ����
		    fmtTitle2.setAlignment(Alignment.LEFT);
		    // ��ֱ����
		    fmtTitle2.setVerticalAlignment(VerticalAlignment.CENTRE);
		    // �߿�
		    fmtTitle2.setBorder(Border.ALL, BorderLineStyle.THIN);
		    // �Զ�����
		    fmtTitle2.setWrap(true);
		    // ����ɫ
		    fmtTitle2.setBackground(Colour.getInternalColour(35));

		  */}
	  public WritableWorkbook createWorkBook(List<EcifTableBean> tables,File file) throws RowsExceededException, WriteException, BiffException, IOException
	  {
		  WritableWorkbook wwb = Workbook.createWorkbook(file);		  
		  WritableSheet ws_index = wwb.createSheet("Ŀ¼", 0);

		    // �����п�
		   
		  ws_index.setColumnView(0, 40);
		  ws_index.setColumnView(1, 40);
		  ws_index.setColumnView(2, 10);
		  ws_index.setColumnView(3, 10);
		  ws_index.setColumnView(4, 10);
		  ws_index.setColumnView(5, 100);
		  
		  
		  ws_index.setRowView(0,400);
		  
		  ws_index.addCell(new Label(0, 0, "ECIF���ݿ��ʵ��",fmtIndexTitle));
		  ws_index.addCell(new Label(1, 0, "ʵ������",fmtIndexTitle));
		  ws_index.addCell(new Label(2, 0, "���ά����Ա",fmtIndexTitle));
		  ws_index.addCell(new Label(3, 0, "���ά��ʱ��",fmtIndexTitle));
		  ws_index.addCell(new Label(4, 0, "Դϵͳ���",fmtIndexTitle));
		  ws_index.addCell(new Label(5, 0, "KTR·��",fmtIndexTitle));
		  
		    Object[] array = tables.toArray();
		    Comparator tabelBeanComparator = new TabelBeanComparator();
		    Arrays.sort(array, tabelBeanComparator);
		  
		  for(int i = 0;i<tables.size();i++){
			  EcifTableBean etb = (EcifTableBean)array[i];
			  //ws_index.addCell(new Label(0,i+1, etb.getCode(),fmtBody));
			  ws_index.addCell(new Label(1,i+1, etb.getName()));
			  ws_index.addCell(new Label(4,i+1, etb.getComment()));
			  ws_index.addCell(new Label(5,i+1, etb.getKTRPath()));
			 
			  
			  WritableSheet ws_module =wwb.createSheet(etb.getName()+"��ӳ���ϵ", i+1);
			  ws_index.addHyperlink(new WritableHyperlink(0,i+1,etb.getCode(),ws_module,1,0));
			  
			  
			  ws_module.setColumnView(0, 30);
			  ws_module.setColumnView(1, 30);
			  ws_module.setColumnView(2, 20);
			  ws_module.setColumnView(3, 10);
			  ws_module.setColumnView(4, 10);
			  ws_module.setColumnView(5, 10);
			  ws_module.setColumnView(6, 10);
			  ws_module.setColumnView(7, 50);
			  ws_module.setColumnView(8, 30);
			  ws_module.setColumnView(9, 30);
			  ws_module.setColumnView(10, 30);
			  ws_module.setColumnView(11, 10);
			  ws_module.setColumnView(12, 10);
			  ws_module.setColumnView(13, 10);
			  ws_module.setColumnView(14, 50);
			  ws_module.setColumnView(15, 50);
			
			  ws_module.setRowView(0,400);
			  ws_module.setRowView(1,400);
			  
			  /********��ͷ***********/
			  ws_module.addCell(new Label(0,0,"Ecif ����",fmtModuleTitle1));//  headfmt *****
			  ws_module.mergeCells(1, 0, 7, 0);
			  ws_module.addCell(new Label(1,0,etb.getName(),fmtModuleTitle2));
			  ws_module.addCell(new Label(0,1,"Ecif �����",fmtModuleTitle1));
			  ws_module.mergeCells(1, 1, 7, 1);
			  ws_module.addCell(new Label(1,1,etb.getCode(),fmtModuleTitle2));
			  			  
			  /***************����*************/
			  ArrayList<EcifColumnBean> columns = (ArrayList<EcifColumnBean>)etb.getColumns();
			 
			  
			  ws_module.addCell(new Label(0,2,"���Ա���",fmtModuleTable1));//targetfmt  *******
			  ws_module.addCell(new Label(1,2,"��������",fmtModuleTable1));
			  ws_module.addCell(new Label(2,2,"��������",fmtModuleTable1));
			  ws_module.addCell(new Label(3,2,"����",fmtModuleTable1));
			  ws_module.addCell(new Label(4,2,"����",fmtModuleTable1));
			  ws_module.addCell(new Label(5,2,"�Ƿ�����",fmtModuleTable1));
			  ws_module.addCell(new Label(6,2,"�Ƿ������ֵ",fmtModuleTable1));
			  ws_module.addCell(new Label(7,2,"˵��",fmtModuleTable1));
			  
			 /// ws_module.addCell(new Label(8,2,"Դϵͳʵ��",fmtModuleTable2));//sourcefmt    *******
			 /// ws_module.addCell(new Label(9,2,"Դϵͳʵ������",fmtModuleTable2));
			 /// ws_module.addCell(new Label(10,2,"Դϵͳ����",fmtModuleTable2));
			 /// ws_module.addCell(new Label(11,2,"��������",fmtModuleTable2));
			 /// ws_module.addCell(new Label(12,2,"����",fmtModuleTable2));
			 /// ws_module.addCell(new Label(13,2,"����",fmtModuleTable2));
			 /// ws_module.addCell(new Label(14,2,"˵��",fmtModuleTable2));
			 /// ws_module.addCell(new Label(15,2,"ת������",fmtModuleTable3));// transrule ************
			  for(int col = 0;col<columns.size();col++){
				  EcifColumnBean ecb = columns.get(col);
				  ws_module.addCell(new Label(0,col+3,ecb.getCode()));
				  ws_module.addCell(new Label(1,col+3,ecb.getName()));
				  ws_module.addCell(new Label(2,col+3,ecb.getType()));
				  ws_module.addCell(new Label(3,col+3,ecb.getLength()+""));
				  ws_module.addCell(new Label(4,col+3,ecb.getDec()+""));
				  ws_module.addCell(new Label(5,col+3,ecb.isKey()?"Y":""));
				  ws_module.addCell(new Label(6,col+3,ecb.isKey()?"Y":""));
				  ws_module.addCell(new Label(7,col+3,ecb.getComment()));				  
			  }
			  ws_module.addHyperlink(new WritableHyperlink(0,columns.size()+5,"����",ws_index,0,0));
		  }
		  return wwb;
		  		  
	  }
	  public WritableWorkbook createPermissionXLS(List<EcifTableBean> tables,File file) throws RowsExceededException, WriteException, BiffException, IOException{
		  WritableWorkbook wwb = Workbook.createWorkbook(file);	
		  WritableSheet ws_index = wwb.createSheet("ʵ�弶����", 0);
		  // �����п�		   
		  ws_index.setColumnView(0, 40);
		  ws_index.setColumnView(1, 40);
		  ws_index.setColumnView(2, 10);
		  ws_index.setColumnView(3, 10);
		  ws_index.setColumnView(4, 20);
		  ws_index.setColumnView(5, 20);
		  ws_index.setColumnView(6, 20);
		  ws_index.setColumnView(7, 10);
		  
		  ws_index.mergeCells(0, 0, 0, 1);
		  ws_index.mergeCells(1, 0, 1, 1);
		  ws_index.mergeCells(2, 0, 6, 0);
		  ws_index.mergeCells(7, 0, 7, 1);
		  
		  ws_index.addCell(new Label(0, 0, "ECIF���ݿ��ʵ��",fmtTitlePermi));
		  ws_index.addCell(new Label(1, 0, "ECIF���ݿ��ʵ��",fmtTitlePermi));
		  ws_index.addCell(new Label(2, 0, "<Դϵͳ����>",fmtTitlePermi));
		  ws_index.addCell(new Label(2, 1, "ԴϵͳID",fmtTitlePermi));
		  ws_index.addCell(new Label(3, 1, "���ȼ�",fmtTitlePermi));
		  ws_index.addCell(new Label(4, 1, "KTR�ļ���",fmtTitlePermi));
		  ws_index.addCell(new Label(5, 1, "�Ƿ���Ҫ����Դ����Ϣ",fmtTitlePermi));
		  ws_index.addCell(new Label(6, 1, "ECIF����ͬ����ϢID",fmtTitlePermi));
		  ws_index.addCell(new Label(7, 0, "���ά������",fmtTitlePermi));
		  
		    Object[] array = tables.toArray();
		    Comparator tabelBeanComparator = new TabelBeanComparator();
		    Arrays.sort(array, tabelBeanComparator);		  
		  
		  for(int i = 0;i<tables.size();i++){
			  EcifTableBean etb = (EcifTableBean)array[i];
			  //ws_index.addCell(new Label(0,i+1, etb.getCode(),fmtBody));
			  ws_index.addCell(new Label(1,i+2, etb.getName()));
			  //ws_index.addCell(new Label(4,i+2, etb.getComment()));
			 
			  
			  WritableSheet ws_module =wwb.createSheet(etb.getName()+"���Լ�����", i+1);
			  ws_index.addHyperlink(new WritableHyperlink(0,i+2,etb.getCode(),ws_module,1,0));
			  
			  
			  ws_module.setColumnView(0, 40);
			  ws_module.setColumnView(1, 40);
			  ws_module.setColumnView(2, 10);
			  ws_module.setColumnView(3, 10);
			  ws_module.setColumnView(4, 20);
			  ws_module.setColumnView(5, 20);
			  ws_module.setColumnView(6, 10);
			
			  ws_module.mergeCells(0, 0, 0, 1);
			  ws_module.mergeCells(1, 0, 1, 1);
			  ws_module.mergeCells(2, 0, 5, 0);
			  ws_module.mergeCells(6, 0, 6, 1);
			  
			  /******************��ͷ*********************/
			  ws_module.addCell(new Label(0,0,etb.getName()+"����",fmtTitlePermi));//
			  ws_module.addCell(new Label(1,0,"��������",fmtTitlePermi));//
			  ws_module.addCell(new Label(2,0,"<Դϵͳ����>",fmtTitlePermi));
			  ws_module.addCell(new Label(2,1,"ԴϵͳID",fmtTitlePermi));
			  ws_module.addCell(new Label(3,1,"���ȼ�",fmtTitlePermi));
			  ws_module.addCell(new Label(4,1,"�Ƿ���Ҫ����Դ����Ϣ",fmtTitlePermi));
			  ws_module.addCell(new Label(5,1,"ECIF����ͬ����ϢID",fmtTitlePermi));
			  ws_module.addCell(new Label(6,0,"���ά��ʱ��",fmtTitlePermi));
			  			  
			  /***************����*************/
			  ArrayList<EcifColumnBean> columns = (ArrayList<EcifColumnBean>)etb.getColumns();
			  
			  for(int col = 0;col<columns.size();col++){
				  EcifColumnBean ecb = columns.get(col);
				  ws_module.addCell(new Label(0,col+2,ecb.getCode()));
				  ws_module.addCell(new Label(1,col+2,ecb.getName()));
			  }
			  ws_module.addHyperlink(new WritableHyperlink(0,columns.size()+4,"����",ws_index,0,0));
		  }
		  
		  
		  return wwb;		  
	  }
	
}
