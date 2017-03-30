/**
 * @程 序 名：EcifTableBean.java
 * @功能描述：YC.ECIF的MS Excel格式的数据库表
 * @程 序 员：郑靖华
 * @编码日期：2010年1月5日
 * @修改日期：
 * @版 本 号：v1.0
 */
package com.yuchengtech.tools.ecif;

import java.util.ArrayList;

public class EcifTableBean {
  private String id; // PowerDesigner为对象分配的内部ID

  private String pathname; // PowerDesigner文档中节点全路径名

  private String code; // 数据库表的代码

  private String name; // 数据库表的中文名称

  private String parent_code; // 上级表的代码

  private String parent_name; // 上级表的中文名称

  private String update_sys; // 更改系统

  private String comment; // 实体描述

  private int new_update; // 新增-1，修改-2，删除-3
  
  private String KTRPath;//生成KTR文件的相对路径

  private ArrayList columns = new ArrayList(); // 栏位

  
  
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
    // 新增-1，修改-2，删除-3
    if (new_update == 1)
      return ("新增实体");
    if (new_update == 1)
      return ("修改实体");
    if (new_update == 1)
      return ("已删除实体");
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
