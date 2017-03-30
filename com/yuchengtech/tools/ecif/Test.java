package com.yuchengtech.tools.ecif;

import java.io.IOException;

import org.dom4j.DocumentException;

import jxl.read.biff.BiffException;
import jxl.write.WriteException;
/**
 * 
 * @author WillJoe
 * @description 锟斤拷锟�
 * @version 1.0
 * @date 2010-03-23
 */
public class Test {

	/**
	 * @param args
	 * @throws IOException 
	 * @throws DocumentException 
	 * @throws WriteException 
	 * @throws BiffException 
	 */
	public static void main(String[] args) throws WriteException, DocumentException, IOException, BiffException {
		// TODO Auto-generated method stub
		WriteExecl we = new WriteExecl();
		String pdm = "C:\\Users\\wlj\\Desktop\\授信系统\\通用数据采集设计.pdm"; 
	    String xsl = "C:\\Users\\wlj\\Desktop\\授信系统\\通用数据采集设计.xls";
	    String perXLS = "C:\\Users\\wlj\\Desktop\\授信系统\\通用数据采集设计-g.xls";
		we.write(xsl, perXLS ,pdm);
	}

}
