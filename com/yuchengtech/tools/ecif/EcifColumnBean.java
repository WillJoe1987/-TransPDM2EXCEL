/**
 * @程 序 名：EcifColumnBean.java
 * @功能描述：YC.ECIF的MS Excel格式的数据库表栏位
 * @程 序 员：郑靖华
 * @编码日期：2010年1月5日
 * @修改日期：
 * @版 本 号：v1.0
 */
package com.yuchengtech.tools.ecif;

public class EcifColumnBean {
  private String id = ""; // PowerDesigner为对象分配的内部ID

  private String code = ""; // 栏位代码

  private String name = ""; // 栏位中文名称

  private String type = ""; // 栏位类型

  private int length = 0; // 栏位长度

  private int dec = 0; // 栏位精度

  private String comment = ""; // 栏位描述

  private String status = ""; // 属性状态

  private boolean key = false; // 是否主键

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
   * @return the type
   */
  public String getType() {
    return type;
  }

  /**
   * @param type the type to set
   */
  public void setType(String type) {
    this.type = type;
  }

  /**
   * @return the length
   */
  public int getLength() {
    return length;
  }

  /**
   * @param length the length to set
   */
  public void setLength(int length) {
    this.length = length;
  }

  /**
   * @return the dec
   */
  public int getDec() {
    return dec;
  }

  /**
   * @param dec the dec to set
   */
  public void setDec(int dec) {
    this.dec = dec;
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
   * @return the status
   */
  public String getStatus() {
    return status;
  }

  /**
   * @param status the status to set
   */
  public void setStatus(String status) {
    this.status = status;
  }

  /**
   * @return the key
   */
  public boolean isKey() {
    return key;
  }

  /**
   * @param key the key to set
   */
  public void setKey(boolean key) {
    this.key = key;
  }

  public String getKeyName() {
    if (this.key)
      return ("P");

    return ("");
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

    sb.append("\tColumn Id: ").append(id).append(", column name: ")
        .append(name).append(", column code: ").append(code);
    sb.append("\n");
    return sb.toString();
  }

}
