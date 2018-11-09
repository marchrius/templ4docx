package org.apache.poi.xwpf.usermodel;

import java.util.List;

/**
 * This is a utility class so that I can get access to the protected fields within XWPFNumbering.
 * Created by jklo on 11/3/16.
 */
public class NumberingUtil {

//  private final XWPFNumbering numbering;
//
//  public NumberingUtil(XWPFNumbering numbering) {
//    this.numbering = numbering;
//  }

  public static List<XWPFAbstractNum> getAbstractNums(XWPFNumbering numbering) {
    return numbering.abstractNums;
  }

  public static List<XWPFNum> getNums(XWPFNumbering numbering) {
    return numbering.nums;
  }

  public static XWPFNumbering getNumbering(XWPFNumbering numbering) {
    return numbering;
  }

}
