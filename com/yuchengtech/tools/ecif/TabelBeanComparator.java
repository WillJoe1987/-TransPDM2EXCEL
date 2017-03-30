package com.yuchengtech.tools.ecif;

import java.util.Comparator;

public class TabelBeanComparator implements Comparator {
  public int compare(Object o1, Object o2) {
    EcifTableBean b1 = (EcifTableBean) o1;
    EcifTableBean b2 = (EcifTableBean) o2;

    return (b1.getCode().compareToIgnoreCase(b2.getCode()));
  }

}
