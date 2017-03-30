package com.yuchengtech.tools.ecif;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import jxl.write.WriteException;
/**
 * 
 * @author WillJoe
 * @description ��ȡPDM�б�ṹ
 * @version 1.0
 * @date 2010-03-23
 */
public class ReadPDM {

	private static List<EcifTableBean> tables = new ArrayList<EcifTableBean>();
	/*
	 * @���췽����������ָ����Ҫ��ȡ������ģ�͵��ļ�������·����
	 */
	public ReadPDM(String pdm) throws WriteException, DocumentException, IOException{
		getTables(pdm);
	}
	/*
	 * @���ã���ȡ����ģ���еı�����б�����tables��
	 * @������ָ����Ҫ��ȡ������ģ�͵��ļ�������·������
	 * 
	 */
	public void getTables(String pdm) throws DocumentException,
			IOException, WriteException {
		SAXReader saxReader = new SAXReader();
		File file = new File(pdm);
		Document doc = saxReader.read(new File(pdm));
		List<Element> list = doc.selectNodes("*//o:Table");	
		list.addAll(doc.selectNodes("*//o:table"));
		System.out.println("Total tables: " + list.size());
		for (Element table : list) {			
			String Path = "";
			Element nearestPacket = table.getParent().getParent();
			List<Element> nearPackElements = (List<Element>)nearestPacket.elements();
			for(Element element:nearPackElements){
				if(element.getQualifiedName().equalsIgnoreCase("a:Name")){
					Path = Path+element.getText();
				}
			}
			
			//Path = Path+nearestPacket.getText();
			while(null!=nearestPacket.getParent()&&null!=nearestPacket.getParent().getParent()&&null!=nearestPacket.getParent().getParent().getQualifiedName()&&!nearestPacket.getParent().getParent().getQualifiedName().equalsIgnoreCase("o:Model")){
				Element parePacket = nearestPacket.getParent().getParent();
				List<Element> packElements = (List<Element>)parePacket.elements();
				for(Element packElement:packElements){
					if(packElement.getQualifiedName().equalsIgnoreCase("a:Name")){
						
						Path = packElement.getText()+"/"+Path;							
					}
				}
				nearestPacket = parePacket;
			}
			getTableInfo(table,Path);
		}
	}

	/*
	 * @��ȡ��Ϣ�б�
	 */
	
	private void getTableInfo(Element table,String Path){
		List<Element> elements = table.elements();
		if (elements.size() < 1)
			return;

		EcifTableBean tbean = new EcifTableBean();
		tbean.setKTRPath(Path);
		for (Element element : elements) {
			String elementName = element.getQualifiedName();
			if (elementName.equalsIgnoreCase("a:Code")) {
				tbean.setCode(element.getText());
			} else if (elementName.equalsIgnoreCase("a:Name")) {
				tbean.setName(element.getText());
			} else if (elementName.equalsIgnoreCase("c:Columns")) {
				getColumnsInfo(element, tbean);
			}			
		}
		String id = table.attributeValue("Id");
		tbean.setId(id);
		tbean.setPathname(table.getPath());
		tbean.setNew_update(1);
		tables.add(tbean);
		System.out.println("��"+tbean.getName()+"�� has write , Path is ��"+tbean.getKTRPath()+"��");
	}

	/*
	 * @��ȡ�����б�
	 */
	private void getColumnsInfo(Element columns, EcifTableBean tbean) {
		List<Element> columns_elements = columns.elements();
		if (columns_elements.size() < 1)
			return;

		for (Element columns_element : columns_elements) {
			String elementName = columns_element.getQualifiedName();
			if (!elementName.equalsIgnoreCase("o:Column")) {
				continue; 
			}

			String id = columns_element.attributeValue("Id");
			List<Element> column_elements = columns_element.elements();
			EcifColumnBean cbean = new EcifColumnBean();
			for (Element column_element : column_elements) {
				elementName = column_element.getQualifiedName();
				if (elementName.equalsIgnoreCase("a:Code")) {
					cbean.setCode(column_element.getText());
				} else if (elementName.equalsIgnoreCase("a:Name")) {
					cbean.setName(column_element.getText());
				} else if (elementName.equalsIgnoreCase("a:DataType")) {
					cbean.setType(column_element.getText());
				} else if (elementName.equalsIgnoreCase("a:Length")) {
					Integer value = new Integer(column_element.getText());
					cbean.setLength(value.intValue());
				} else if(elementName.equalsIgnoreCase("a:Comment")){
					cbean.setComment(column_element.getText());
				}
			}

			cbean.setId(id);
			tbean.addColumns(cbean);
		}
	}

	public static List<EcifTableBean> getTables() {
		return tables;
	}

	public static void setTables(List<EcifTableBean> tables) {
		ReadPDM.tables = tables;
	}
	
	
	
}
