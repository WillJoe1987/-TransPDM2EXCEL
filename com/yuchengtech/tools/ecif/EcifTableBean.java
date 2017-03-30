/**
 * @�� �� ����EcifTableBean.java
 * @����������YC.ECIF��MS Excel��ʽ�����ݿ��
 * @�� �� Ա��֣����
 * @�������ڣ�2010��1��5��
 * @�޸����ڣ�
 * @�� �� �ţ�v1.0
 */
package com.yuchengtech.tools.ecif;

import java.util.ArrayList;

public class EcifTableBean {
  private String id; // PowerDesignerΪ���������ڲ�ID

  private String pathname; // PowerDesigner�ĵ��нڵ�ȫ·����

  private String code; // ���ݿ��Ĵ���

  private String name; // ���ݿ�����������

  private String parent_code; // �ϼ���Ĵ���

  private String parent_name; // �ϼ������������

  private String update_sys; // ����ϵͳ

  private String comment; // ʵ������

  private int new_update; // ����-1���޸�-2��ɾ��-3
  
  private String KTRPath;//����KTR�ļ������·��

  private ArrayList columns = new ArrayList(); // ��λ

  
  
  public String getKTRPath() {
	return KTRPath;
}

public void setKTRPath(String path) {
	KTRPath = path;
}

/**
   * @return the code
   */
  public String getCode() {
    return code;
  }

  /**
   * @param code the code to set
   */
  public void setCode(String code) {
    this.code = code;
  }

  /**
   * @return the name
   */
  public String getName() {
    return name;
  }

  /**
   * @param name the name to set
   */
  public void setName(String name) {
    this.name = name;
  }

  /**
   * @return the parent_code
   */
  public String getParent_code() {
    return parent_code;
  }

  /**
   * @param parent_code the parent_code to set
   */
  public void setParent_code(String parent_code) {
    this.parent_code = parent_code;
  }

  /**
   * @return the parent_name
   */
  public String getParent_name() {
    return parent_name;
  }

  /**
   * @param parent_name the parent_name to set
   */
  public void setParent_name(String parent_name) {
    this.parent_name = parent_name;
  }

  /**
   * @return the update_sys
   */
  public String getUpdate_sys() {
    return update_sys;
  }

  /**
   * @param update_sys the update_sys to set
   */
  public void setUpdate_sys(String update_sys) {
    this.update_sys = update_sys;
  }

  /**
   * @return the comment
   */
  public String getComment() {
    return comment;
  }

  /**
   * @param comment the comment to set
   */
  public void setComment(String comment) {
    this.comment = comment;
  }

  /**
   * @return the new_update
   */
  public int getNew_update() {
    return new_update;
  }

  /**
   * @param new_update the new_update to set
   */
  public void setNew_update(int new_update) {
    this.new_update = new_update;
  }

  public String getNew_update_Name() {
    // ����-1���޸�-2��ɾ��-3
    if (new_update == 1)
      return ("����ʵ��");
    if (new_update == 1)
      return ("�޸�ʵ��");
    if (new_update == 1)
      return ("��ɾ��ʵ��");
    return ("");
  }

  /**
   * @return the columns
   */
  public ArrayList getColumns() {
    return columns;
  }

  /**
   * @param columns the columns to set
   */
  public void setColumns(ArrayList columns) {
    this.columns = columns;
  }

  public void addColumns(EcifColumnBean column) {
    this.columns.add(column);
  }

  /**
   * @return the id
   */
  public String getId() {
    return id;
  }

  /**
   * @param id the id to set
   */
  public void setId(String id) {
    this.id = id;
  }

  public String toString() {
    StringBuffer sb = new StringBuffer();

    sb.append("Table Id: ").append(id).append(", Table name: ").append(name)
        .append(", table code: ").append(code).append(", pathname: ").append(
            pathname);
    sb.append("\n");
    for (int i = 0; i < columns.size(); i++) {
      EcifColumnBean cbean = (EcifColumnBean) columns.get(i);
      sb.append(cbean.toString());
    }
    return sb.toString();
  }

  /**
   * @return the pathname
   */
  public String getPathname() {
    return pathname;
  }

  /**
   * @param pathname the pathname to set
   */
  public void setPathname(String pathname) {
    this.pathname = pathname;
  }
}
