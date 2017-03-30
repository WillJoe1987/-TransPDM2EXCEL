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
 * @description 创建并写入Maping文件和权限矩阵文件
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
			//创建上层目录
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
	  /*/ "模型"明细表头项的单元格格式
	  WritableCellFormat fmtTitle = new WritableCellFormat();
	  

	  // "模型"明细表头项的单元格格式(专用于TABLE CODE的背景色)
	  WritableCellFormat fmtTitle1 = new WritableCellFormat();

	  // 普通表体的单元格格式(文本)
	  WritableCellFormat fmtBody = new WritableCellFormat();

	  // 普通表体的单元格格式(整数)
	  WritableCellFormat fmtBody_int = new WritableCellFormat();

	  // 超链接表体的单元格格式
	  WritableCellFormat fmtBody_hyper = new WritableCellFormat();

	  // 页眉
	  HeaderFooter header = new HeaderFooter();

	  // "索引"表头项的单元格格式
	  WritableCellFormat fmtTitle2 = new WritableCellFormat();
	  
	  
	  */
		/**************目录表头格式************/
		WritableCellFormat fmtIndexTitle = new WritableCellFormat();
		WritableCellFormat fmtModuleTitle1 = new WritableCellFormat();
		WritableCellFormat fmtModuleTitle2 = new WritableCellFormat();
		WritableCellFormat fmtModuleTable1 = new WritableCellFormat();
		WritableCellFormat fmtModuleTable2 = new WritableCellFormat();
		WritableCellFormat fmtModuleTable3 = new WritableCellFormat();
		WritableCellFormat fmtTitlePermi = new WritableCellFormat();
	  private void setSheetFmt() throws WriteException {
		    // 表头项的单元格格式
		    // 字体
		    WritableFont fontTitle = new WritableFont(WritableFont.createFont("楷体"),
		        10, WritableFont.NO_BOLD, false);
		    WritableFont fontIndexTitle = new WritableFont(WritableFont.createFont("楷体"),
			        12, WritableFont.NO_BOLD, false);
		    WritableFont fmtModuleTitle = new WritableFont(WritableFont.createFont("楷体"),
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
		    // 水平对齐
		    fmtTitle.setAlignment(Alignment.LEFT);
		    // 垂直对齐
		    fmtTitle.setVerticalAlignment(VerticalAlignment.CENTRE);
		    // 边框
		    fmtTitle.setBorder(Border.ALL, BorderLineStyle.THIN);
		    // 自动换行
		    fmtTitle.setWrap(true);

		    // 表头项的单元格格式
		    // 字体
		    WritableFont fontTitle1 = new WritableFont(WritableFont.createFont("宋体"),
		        10, WritableFont.NO_BOLD, false);
		    fmtTitle1.setFont(fontTitle1);
		    // 水平对齐
		    fmtTitle1.setAlignment(Alignment.LEFT);
		    // 垂直对齐
		    fmtTitle1.setVerticalAlignment(VerticalAlignment.CENTRE);
		    // 边框
		    fmtTitle1.setBorder(Border.ALL, BorderLineStyle.THIN);
		    // 自动换行
		    fmtTitle1.setWrap(true);
		    // 背景色
		    fmtTitle1.setBackground(Colour.getInternalColour(35));

		    // 普通表体的单元格格式(文本)
		    // 字体
		    WritableFont fontBody = new WritableFont(WritableFont.createFont("宋体"), 10,
		        WritableFont.NO_BOLD, false);
		    fmtBody.setFont(fontBody);
		    // 水平对齐
		    fmtBody.setAlignment(Alignment.LEFT);
		    // 垂直对齐
		    fmtBody.setVerticalAlignment(VerticalAlignment.CENTRE);
		    // 边框
		    fmtBody.setBorder(Border.ALL, BorderLineStyle.THIN);
		    // 自动换行
		    fmtBody.setWrap(true);

		    // 普通表体的单元格格式(整数)
		    // 字体
		    WritableFont fontBody_int = new WritableFont(WritableFont.createFont("宋体"),
		        10, WritableFont.NO_BOLD, false);
		    fmtBody_int.setFont(fontBody_int);
		    // 水平对齐
		    fmtBody_int.setAlignment(Alignment.RIGHT);
		    // 垂直对齐
		    fmtBody_int.setVerticalAlignment(VerticalAlignment.CENTRE);
		    // 边框
		    fmtBody_int.setBorder(Border.ALL, BorderLineStyle.THIN);
		    // 自动换行
		    fmtBody_int.setWrap(true);

		    // 超链接表体的单元格格式
		    // 字体
		    WritableFont fontBody_hyper = new WritableFont(WritableFont
		        .createFont("宋体"), 10, WritableFont.NO_BOLD, false);
		    fontBody_hyper.setUnderlineStyle(UnderlineStyle.SINGLE);
		    fontBody_hyper.setColour(Colour.YELLOW);
		    fmtBody_hyper.setFont(fontBody_hyper);
		    // 水平对齐
		    fmtBody_hyper.setAlignment(Alignment.LEFT);
		    // 垂直对齐
		    fmtBody_hyper.setVerticalAlignment(VerticalAlignment.CENTRE);
		    // 边框
		    fmtBody_hyper.setBorder(Border.ALL, BorderLineStyle.THIN);
		    // 自动换行
		    fmtBody_hyper.setWrap(true);

		    // 页眉
		    header.getCentre().setFontName("宋体");
		    header.getCentre().setFontSize(22);
		    header.getCentre().toggleBold();
		    header.getCentre().append("物理模型表");
		    header.getRight().setFontName("宋体");
		    header.getRight().setFontSize(12);
		    header.getRight().append("\012\012共   页，第   页");

		    // 设置页眉
		    header.getLeft().clear();
		    header.getLeft().setFontName("宋体");
		    header.getLeft().setFontSize(12);
		    header.getLeft().toggleBold();
		    header.getLeft().append("\012\012责任人：");
		    header.getLeft().toggleBold();
		    header.getLeft().append("郑靖华");

		    // 表头项的单元格格式
		    // 字体
		    WritableFont fontTitle2 = new WritableFont(WritableFont.createFont("宋体"),
		        10, WritableFont.BOLD, false);
		    fmtTitle2.setFont(fontTitle2);
		    // 水平对齐
		    fmtTitle2.setAlignment(Alignment.LEFT);
		    // 垂直对齐
		    fmtTitle2.setVerticalAlignment(VerticalAlignment.CENTRE);
		    // 边框
		    fmtTitle2.setBorder(Border.ALL, BorderLineStyle.THIN);
		    // 自动换行
		    fmtTitle2.setWrap(true);
		    // 背景色
		    fmtTitle2.setBackground(Colour.getInternalColour(35));

		  */}
	  public WritableWorkbook createWorkBook(List<EcifTableBean> tables,File file) throws RowsExceededException, WriteException, BiffException, IOException
	  {
		  WritableWorkbook wwb = Workbook.createWorkbook(file);		  
		  WritableSheet ws_index = wwb.createSheet("目录", 0);

		    // 设置列宽
		   
		  ws_index.setColumnView(0, 40);
		  ws_index.setColumnView(1, 40);
		  ws_index.setColumnView(2, 10);
		  ws_index.setColumnView(3, 10);
		  ws_index.setColumnView(4, 10);
		  ws_index.setColumnView(5, 100);
		  
		  
		  ws_index.setRowView(0,400);
		  
		  ws_index.addCell(new Label(0, 0, "ECIF数据库表实体",fmtIndexTitle));
		  ws_index.addCell(new Label(1, 0, "实体名称",fmtIndexTitle));
		  ws_index.addCell(new Label(2, 0, "最近维护人员",fmtIndexTitle));
		  ws_index.addCell(new Label(3, 0, "最近维护时间",fmtIndexTitle));
		  ws_index.addCell(new Label(4, 0, "源系统编号",fmtIndexTitle));
		  ws_index.addCell(new Label(5, 0, "KTR路径",fmtIndexTitle));
		  
		    Object[] array = tables.toArray();
		    Comparator tabelBeanComparator = new TabelBeanComparator();
		    Arrays.sort(array, tabelBeanComparator);
		  
		  for(int i = 0;i<tables.size();i++){
			  EcifTableBean etb = (EcifTableBean)array[i];
			  //ws_index.addCell(new Label(0,i+1, etb.getCode(),fmtBody));
			  ws_index.addCell(new Label(1,i+1, etb.getName()));
			  ws_index.addCell(new Label(4,i+1, etb.getComment()));
			  ws_index.addCell(new Label(5,i+1, etb.getKTRPath()));
			 
			  
			  WritableSheet ws_module =wwb.createSheet(etb.getName()+"表映射关系", i+1);
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
			  
			  /********表头***********/
			  ws_module.addCell(new Label(0,0,"Ecif 表名",fmtModuleTitle1));//  headfmt *****
			  ws_module.mergeCells(1, 0, 7, 0);
			  ws_module.addCell(new Label(1,0,etb.getName(),fmtModuleTitle2));
			  ws_module.addCell(new Label(0,1,"Ecif 表编码",fmtModuleTitle1));
			  ws_module.mergeCells(1, 1, 7, 1);
			  ws_module.addCell(new Label(1,1,etb.getCode(),fmtModuleTitle2));
			  			  
			  /***************属性*************/
			  ArrayList<EcifColumnBean> columns = (ArrayList<EcifColumnBean>)etb.getColumns();
			 
			  
			  ws_module.addCell(new Label(0,2,"属性编码",fmtModuleTable1));//targetfmt  *******
			  ws_module.addCell(new Label(1,2,"属性名称",fmtModuleTable1));
			  ws_module.addCell(new Label(2,2,"属性类型",fmtModuleTable1));
			  ws_module.addCell(new Label(3,2,"长度",fmtModuleTable1));
			  ws_module.addCell(new Label(4,2,"精度",fmtModuleTable1));
			  ws_module.addCell(new Label(5,2,"是否主键",fmtModuleTable1));
			  ws_module.addCell(new Label(6,2,"是否允许空值",fmtModuleTable1));
			  ws_module.addCell(new Label(7,2,"说明",fmtModuleTable1));
			  
			 /// ws_module.addCell(new Label(8,2,"源系统实体",fmtModuleTable2));//sourcefmt    *******
			 /// ws_module.addCell(new Label(9,2,"源系统实体名称",fmtModuleTable2));
			 /// ws_module.addCell(new Label(10,2,"源系统属性",fmtModuleTable2));
			 /// ws_module.addCell(new Label(11,2,"属性类型",fmtModuleTable2));
			 /// ws_module.addCell(new Label(12,2,"长度",fmtModuleTable2));
			 /// ws_module.addCell(new Label(13,2,"精度",fmtModuleTable2));
			 /// ws_module.addCell(new Label(14,2,"说明",fmtModuleTable2));
			 /// ws_module.addCell(new Label(15,2,"转换规则",fmtModuleTable3));// transrule ************
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
			  ws_module.addHyperlink(new WritableHyperlink(0,columns.size()+5,"返回",ws_index,0,0));
		  }
		  return wwb;
		  		  
	  }
	  public WritableWorkbook createPermissionXLS(List<EcifTableBean> tables,File file) throws RowsExceededException, WriteException, BiffException, IOException{
		  WritableWorkbook wwb = Workbook.createWorkbook(file);	
		  WritableSheet ws_index = wwb.createSheet("实体级矩阵", 0);
		  // 设置列宽		   
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
		  
		  ws_index.addCell(new Label(0, 0, "ECIF数据库表实体",fmtTitlePermi));
		  ws_index.addCell(new Label(1, 0, "ECIF数据库表实体",fmtTitlePermi));
		  ws_index.addCell(new Label(2, 0, "<源系统名称>",fmtTitlePermi));
		  ws_index.addCell(new Label(2, 1, "源系统ID",fmtTitlePermi));
		  ws_index.addCell(new Label(3, 1, "优先级",fmtTitlePermi));
		  ws_index.addCell(new Label(4, 1, "KTR文件名",fmtTitlePermi));
		  ws_index.addCell(new Label(5, 1, "是否需要覆盖源表信息",fmtTitlePermi));
		  ws_index.addCell(new Label(6, 1, "ECIF主动同步消息ID",fmtTitlePermi));
		  ws_index.addCell(new Label(7, 0, "最近维护日期",fmtTitlePermi));
		  
		    Object[] array = tables.toArray();
		    Comparator tabelBeanComparator = new TabelBeanComparator();
		    Arrays.sort(array, tabelBeanComparator);		  
		  
		  for(int i = 0;i<tables.size();i++){
			  EcifTableBean etb = (EcifTableBean)array[i];
			  //ws_index.addCell(new Label(0,i+1, etb.getCode(),fmtBody));
			  ws_index.addCell(new Label(1,i+2, etb.getName()));
			  //ws_index.addCell(new Label(4,i+2, etb.getComment()));
			 
			  
			  WritableSheet ws_module =wwb.createSheet(etb.getName()+"属性级矩阵", i+1);
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
			  
			  /******************表头*********************/
			  ws_module.addCell(new Label(0,0,etb.getName()+"属性",fmtTitlePermi));//
			  ws_module.addCell(new Label(1,0,"属性名称",fmtTitlePermi));//
			  ws_module.addCell(new Label(2,0,"<源系统名称>",fmtTitlePermi));
			  ws_module.addCell(new Label(2,1,"源系统ID",fmtTitlePermi));
			  ws_module.addCell(new Label(3,1,"优先级",fmtTitlePermi));
			  ws_module.addCell(new Label(4,1,"是否需要覆盖源表信息",fmtTitlePermi));
			  ws_module.addCell(new Label(5,1,"ECIF主动同步消息ID",fmtTitlePermi));
			  ws_module.addCell(new Label(6,0,"最近维护时间",fmtTitlePermi));
			  			  
			  /***************属性*************/
			  ArrayList<EcifColumnBean> columns = (ArrayList<EcifColumnBean>)etb.getColumns();
			  
			  for(int col = 0;col<columns.size();col++){
				  EcifColumnBean ecb = columns.get(col);
				  ws_module.addCell(new Label(0,col+2,ecb.getCode()));
				  ws_module.addCell(new Label(1,col+2,ecb.getName()));
			  }
			  ws_module.addHyperlink(new WritableHyperlink(0,columns.size()+4,"返回",ws_index,0,0));
		  }
		  
		  
		  return wwb;		  
	  }
	
}
