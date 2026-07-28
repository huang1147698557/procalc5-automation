package com.sd.discovery.single;

import cn.hutool.core.date.DateUtil;
import cn.hutool.core.thread.ThreadUtil;
import cn.hutool.core.util.NumberUtil;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import com.google.common.collect.Lists;
import java.util.Arrays;
import java.util.List;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

/**
 * BaiscApplication
 *
 * @Author: gaoweiqi
 * @CreateDate: 2021/2/3 16:26
 */
@SpringBootApplication(scanBasePackages = {"com.sd.discovery.*"})
public class BaiscApplication {

  public static void main(String[] args) {
    SpringApplication.run(BaiscApplication.class, args);
    // macOS 路径配置
    String baseDir = System.getProperty("user.dir") + "/procalc5/";
    System.setProperty("webdriver.chrome.driver", "/opt/homebrew/bin/chromedriver");
    test(baseDir);
  }

  private static void test(String baseDir) {
    StringBuilder ss = new StringBuilder();
    WebDriver driver = new ChromeDriver();
    String sheetName = DateUtil.format(DateUtil.date(), "yyyyMMddHHmmss");
    driver.get("https://procalc5.proflute.se/rotor");
    ThreadUtil.safeSleep(8000);
    WebElement username = driver.findElement(By.id("userNameInput"));
    WebElement password = driver.findElement(By.id("passwordInput"));
    username.sendKeys("EXTCNLILZHE");
    password.sendKeys("@4vxGw7L9`\"q");
    WebElement login = driver.findElement(By.xpath("//*[@id=\"submitButton\"]"));
    login.click();
    ThreadUtil.safeSleep(5000);
    ExcelWriter excelWriter = ExcelUtil.getWriter(baseDir + "result.xlsx", sheetName);
    excelWriter.writeHeadRow(Lists.newArrayList("序号", " Wet Air:", "", "", "", "", "Process left", "", "", "", "", "", "", "Process Right",
        "", "", "", "", "", "Reactivation", "", "", "", "", "", "", "RPH"));
    //登陆成功
    List<List<Object>> paraList = ExcelUtil.getReader(baseDir + "procalc5.proflute.xlsx").read();
    for (List<Object> list : paraList) {
      if (paraList.indexOf(list) == 0) {
        continue;
      }
      String linesNumber = StrUtil.toString(list.get(0));
      String UnitsofMeasure = StrUtil.toString(list.get(1));
      String RelativeHumidity = StrUtil.toString(list.get(2));
      String WetBulb = StrUtil.toString(list.get(3));
      String Pressurealtitud = StrUtil.toString(list.get(4));
      String PressurealtitudV = StrUtil.toString(list.get(5));
      String Showbypass = StrUtil.toString(list.get(6));
      String Reactivationinputtype = StrUtil.toString(list.get(7));
      String AirflowRange = StrUtil.toString(list.get(8));
      String Dewpointrange = StrUtil.toString(list.get(9));
      String Performancesafetyfactor = StrUtil.toString(list.get(10));
      String PerformancesafetyfactorV = StrUtil.toString(list.get(11));

      String ProcessAirflow = StrUtil.toString(list.get(12));
      String DesiccantNedia = StrUtil.toString(list.get(13));
      String SectorLayout = StrUtil.toString(list.get(14));
      String RotorDiameter = StrUtil.toString(list.get(16));
      String RotorDepth = StrUtil.toString(list.get(17));
      String NetFaceAreaCalculation = StrUtil.toString(list.get(18));
      String SealingArea = StrUtil.toString(list.get(19));

      String ProcessStrC = StrUtil.toString(list.get(20));
      String ProcessStrGKG = StrUtil.toString(list.get(21));
      String Rph = StrUtil.toString(list.get(22));
      Double ReactivationStart = Double.parseDouble(list.get(23).toString());
      Double ReactivationEnd = Double.parseDouble(list.get(24).toString());
      Double Reactivationbc = Double.parseDouble(list.get(25).toString());
      Double fanweiStart = Double.parseDouble(StrUtil.split(((String) list.get(26)), "~").get(0));
      Double fanweiEnd = Double.parseDouble(StrUtil.split(((String) list.get(26)), "~").get(1));

      String Reactivation1 = StrUtil.toString(list.get(27));
      String Reactivation2 = StrUtil.toString(list.get(28));
      String Reactivation3 = StrUtil.toString(list.get(29));
      //元素赋值
      //范围1赋值
      if (StrUtil.equalsIgnoreCase(UnitsofMeasure, "si")) {
        driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[1]/div/label[1]/span/input")).click();
      } else {
        driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[1]/div/label[2]/span/input")).click();
      }
      if (!StrUtil.equalsIgnoreCase(RelativeHumidity, "勾选")) {
        driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[2]/div/div/label[1]/span/input")).click();
      }
      if (!StrUtil.equalsIgnoreCase(WetBulb, "勾选")) {
        driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[2]/div/div/label[2]/span/input")).click();
      }
      if (StrUtil.equalsIgnoreCase(Pressurealtitud, "Altitude")) {
        driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[3]/div/div[1]/label[1]/span/input")).click();
      } else {
        driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[3]/div/div[1]/label[2]/span/input")).click();
      }
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[3]/div/div[2]/input"))
          .sendKeys(Keys.chord(Keys.CONTROL, "a"), PressurealtitudV);

      if (StrUtil.equalsIgnoreCase(Showbypass, "No")) {
        driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[4]/div/div/label[1]/span/input")).click();
      } else {
        driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[4]/div/div/label[2]/span/input")).click();
      }
      if (StrUtil.equalsIgnoreCase(Reactivationinputtype, "Temp")) {
        driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[5]/div/div/label[1]/span/input")).click();
      } else {
        driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[5]/div/div/label[2]/span/input")).click();
      }
      if (StrUtil.equalsIgnoreCase(AirflowRange, "Default")) {
        driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[6]/div/div/label[1]/span/input")).click();
      } else {
        driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[6]/div/div/label[2]/span/input")).click();
      }
      if (StrUtil.equalsIgnoreCase(Dewpointrange, "Default")) {
        driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[7]/div/div/label[1]/span/input")).click();
      } else {
        driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[7]/div/div/label[2]/span/input")).click();
      }
      //Performance safety factor None DeltaPercent Multiplier
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[8]/div/div/div/div")).click();
      ThreadUtil.safeSleep(500);
      List<WebElement> DesiccantMedias1 = driver.findElements(By.xpath("//*[@id=\"menu-\"]/div[3]/ul/li"));
      boolean flag = true;
      if (StrUtil.equalsIgnoreCase(Performancesafetyfactor, "None")) {
        DesiccantMedias1.get(0).click();
        flag = false;
      } else if (StrUtil.equalsIgnoreCase(Performancesafetyfactor, "+Δ% Moisture")) {
        DesiccantMedias1.get(1).click();
        flag = false;
      } else if (StrUtil.equalsIgnoreCase(Performancesafetyfactor, "x Moisture")) {
        DesiccantMedias1.get(2).click();
        flag = false;
      }
      if (flag) {
        DesiccantMedias1.get(0).click();
      }
      if (!StrUtil.equalsIgnoreCase(Performancesafetyfactor, "None")) {
        driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[1]/div/div/div/div[8]/div/div/div[2]/div/div")).click();
        List<WebElement> RotorDepths = driver.findElements(By.xpath("//*[@id=\"menu-\"]/div[3]/ul/li"));
        for (WebElement s : RotorDepths) {
          String sTemp = s.getAttribute("data-value");
          if (StrUtil.equalsAnyIgnoreCase(PerformancesafetyfactorV, sTemp)) {
            s.click();
            break;
          }
        }
      }
      //范围2赋值
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[1]/div/div[2]/input"))
          .sendKeys(Keys.chord(Keys.CONTROL, "a"), ProcessAirflow);
      //media
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[3]/div/div[2]")).click();
      ThreadUtil.safeSleep(1000);
      List<WebElement> DesiccantMedias = driver.findElements(By.xpath("//*[@id=\"menu-\"]/div[3]/ul/li"));
      flag = true;
      for (WebElement s : DesiccantMedias) {
        String sTemp = StrUtil.equalsIgnoreCase(s.getAttribute("data-value"), "1") ? "PPS" : "PPP";
        if (StrUtil.equalsIgnoreCase(DesiccantNedia, sTemp)) {
          s.click();
          flag = false;
          break;
        }
      }
      if (flag) {
        DesiccantMedias.get(0).click();
      }
      //Sector layout
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[4]/div/div[2]")).click();
      ThreadUtil.safeSleep(1000);
      List<WebElement> SectorLayouts = driver.findElements(By.xpath("//*[@id=\"menu-\"]/div[3]/ul/li"));
      flag = true;
      for (WebElement s : SectorLayouts) {
        switch (s.getAttribute("data-value")) {
          case "1":
            if (StrUtil.equalsAnyIgnoreCase(SectorLayout, "L270/90")) {
              s.click();
              flag = false;
              break;
            }
            ;
          case "2":
            if (StrUtil.equalsAnyIgnoreCase(SectorLayout, "L180/180")) {
              s.click();
              flag = false;
              break;
            }
            ;
          case "3":
            if (StrUtil.equalsAnyIgnoreCase(SectorLayout, "L240/120")) {
              s.click();
              flag = false;
              break;
            }
            ;
          case "4":
            if (StrUtil.equalsAnyIgnoreCase(SectorLayout, "L240/90/30")) {
              s.click();
              flag = false;
              break;
            }
            ;
          case "5":
            if (StrUtil.equalsAnyIgnoreCase(SectorLayout, "L180/90/90")) {
              s.click();
              flag = false;
              break;
            }
            ;
          case "6":
            if (StrUtil.equalsAnyIgnoreCase(SectorLayout, "L240/60/60")) {
              s.click();
              flag = false;
              break;
            }
            ;
          case "7":
            if (StrUtil.equalsAnyIgnoreCase(SectorLayout, "L230/65/65")) {
              s.click();
              flag = false;
              break;
            }
            ;
        }
      }
      if (flag) {
        SectorLayouts.get(0).click();
      }
/////
      WebElement RotorDiameterE = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[6]/div/div[2]"));
      RotorDiameterE.click();
      //Rotor diameter
      List<WebElement> RotorDiameters = driver.findElements(By.xpath("//*[@id=\"menu-\"]/div[3]/ul/li"));
      flag = true;
      for (WebElement s : RotorDiameters) {
        if (StrUtil.equalsIgnoreCase(RotorDiameter, s.getAttribute("data-value"))) {
          s.click();
          flag = false;
          break;
        }
      }
      if (flag) {
        RotorDiameters.get(0).click();
      }
      //Rotor depth
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[7]/div/div[2]")).click();
      ThreadUtil.safeSleep(1000);
      List<WebElement> RotorDepths = driver.findElements(By.xpath("//*[@id=\"menu-\"]/div[3]/ul/li"));
      flag = true;//*[@id="menu-"]/div[3]/ul
      for (WebElement s : RotorDepths) {
        if (StrUtil.equalsIgnoreCase(RotorDepth, s.getAttribute("data-value"))) {
          s.click();
          flag = false;
          break;
        }
      }
      if (flag) {
        RotorDepths.get(0).click();
      }
      //Net face area calculation
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[8]/div/div[2]")).click();
      ThreadUtil.safeSleep(1000);
      List<WebElement> Netfaceareacalculations = driver.findElements(By.xpath("//*[@id=\"menu-\"]/div[3]/ul/li"));
      flag = true;
      for (WebElement s : Netfaceareacalculations) {
        String sTemp = StrUtil.equalsIgnoreCase(s.getAttribute("data-value"), "0") ? "Sealing area" : "Active area";
        if (StrUtil.equalsIgnoreCase(NetFaceAreaCalculation, sTemp)) {
          s.click();
          flag = false;
          break;
        }
      }
      if (flag) {
        Netfaceareacalculations.get(0).click();
      }
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[1]/div[1]/div[2]/div/div/div/div[9]/div/div[2]/input")).sendKeys(Keys.chord(Keys.CONTROL, "a"), SealingArea);
      //按顺序放值
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[1]/div/div/input")).sendKeys(Keys.chord(Keys.CONTROL, "a"), ProcessStrC);
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[2]/div/div/input")).sendKeys(Keys.chord(Keys.CONTROL, "a"), ProcessStrGKG);
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[8]/div/div[2]/div/div/div/div/input")).sendKeys(Keys.chord(Keys.CONTROL, "a"), Rph);
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[5]/div/div[2]/div/div[1]/div/div/input")).sendKeys(Keys.chord(Keys.CONTROL, "a"), Reactivation1);
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[5]/div/div[2]/div/div[2]/div/div/input")).sendKeys(Keys.chord(Keys.CONTROL, "a"), Reactivation2);
      driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[5]/div/div[2]/div/div[5]/div/div/input")).sendKeys(Keys.chord(Keys.CONTROL, "a"), Reactivation3);
      WebElement button = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[9]/button"));
      //按照步长处理数据
      WebElement Reactivation = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[3]/div/div[2]/div/div[1]/div/div/input"));
      WebElement gkg = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div[2]/div/div/input"));
      //先算范围的边界值
      Reactivation.sendKeys(Keys.chord(Keys.CONTROL, "a"), ReactivationStart.toString());
      click(button);
      ThreadUtil.safeSleep(1500);
      if (StrUtil.isEmpty(gkg.getAttribute("value"))) {
        continue;
      }
      Double gkgLeft = Double.parseDouble(gkg.getAttribute("value"));
      Reactivation.sendKeys(Keys.chord(Keys.CONTROL, "a"), ReactivationEnd.toString());
      click(button);
      ThreadUtil.safeSleep(1500);
      if (StrUtil.isEmpty(gkg.getAttribute("value"))) {
        continue;
      }
      boolean qk1 = (fanweiStart <= gkgLeft && gkgLeft <= fanweiEnd);
      Double ReactivationStartReal = ReactivationStart;
      Double ReactivationEndReal = ReactivationEnd;
      //左侧不在范围内
      if (!qk1) {
        Reactivation.sendKeys(Keys.chord(Keys.CONTROL, "a"), StrUtil.toString(NumberUtil.add(ReactivationStart, Reactivationbc)));
        click(button);
        ThreadUtil.safeSleep(1500);
        if (StrUtil.isEmpty(gkg.getAttribute("value"))) {
          continue;
        }
        Double gkgTemp = Double.parseDouble(gkg.getAttribute("value"));
        //递增模式
        if (NumberUtil.compare(gkgTemp, gkgLeft) > 0) {
          if (NumberUtil.compare(gkgLeft, fanweiStart) < 0 && NumberUtil.compare(gkgTemp, gkgLeft) != 0) {
            Double wendus = NumberUtil.mul((Double) NumberUtil.div(NumberUtil.sub(fanweiStart, gkgLeft),
                NumberUtil.sub(gkgTemp, gkgLeft)), Reactivationbc);
            ReactivationStartReal = NumberUtil.add(ReactivationStart.doubleValue(), NumberUtils.safeMultiply(wendus, 1.0, 0).doubleValue());
          } else {
            ReactivationStartReal = null;
          }
        } else { //递减模式
          if (NumberUtil.compare(gkgLeft, fanweiEnd) > 0 && NumberUtil.compare(gkgLeft, gkgTemp) != 0) {
            Double wendus = NumberUtil.mul((Double) NumberUtil.div(NumberUtil.sub(gkgLeft, fanweiEnd),
                NumberUtil.sub(gkgLeft, gkgTemp)), Reactivationbc);
            ReactivationStartReal = NumberUtil.add(ReactivationStart.doubleValue(), NumberUtils.safeMultiply(wendus, 1.0, 0).doubleValue());
          } else {
            ReactivationStartReal = null;
          }
        }
      }
      //右侧不在范围内
      if (ReactivationEndReal == null || ReactivationStartReal == null) {
        ReactivationStartReal = ReactivationStart;
        ReactivationEndReal = ReactivationEnd;
      }
      flag = false;
      Double templeft = ReactivationStartReal.doubleValue();
      while (templeft <= ReactivationEndReal.doubleValue()) {
        Reactivation.sendKeys(Keys.chord(Keys.CONTROL, "a"), StrUtil.toString(templeft));
        try {
          click(button);
        } catch (Exception e) {
          System.out.println(e.getMessage());
        }
        ThreadUtil.safeSleep(1500);
        if (StrUtil.isEmpty(gkg.getAttribute("value"))) {
          break;
        }
        if (StrUtil.isEmpty(gkg.getAttribute("value"))) {
          continue;
        }
        Double gkgTemp = Double.parseDouble(gkg.getAttribute("value"));
        if (fanweiStart <= gkgTemp && gkgTemp <= fanweiEnd) {
          toList(driver, ss, linesNumber, excelWriter);
          flag = true;
        }
        templeft = NumberUtil.add(templeft, Reactivationbc);
        if (flag && !(fanweiStart <= gkgTemp && gkgTemp <= fanweiEnd)) {
          break;
        }
      }
      System.out.println("Datas:" + ss);
      ThreadUtil.safeSleep(1000);
    }
    excelWriter.flush();
    System.out.println("总数据:" + ss);
  }

  public static void click(WebElement button) {
    click(button, 3);
  }

  public static void click(WebElement button, int time) {
    try {
      button.click();
    } catch (Exception e) {
      System.out.println("按钮点击异常重新 点击");
      if (time > 0) {
        ThreadUtil.safeSleep(1000);
        click(button, --time);
      }

    }
  }

  public static void toList(WebDriver driver, StringBuilder ss, String lineNumber, ExcelWriter excelWriter) {
    List<String> list = Lists.newArrayList();
    String v1 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[2]/div/div[2]/div/div[1]/div/div/input")).getAttribute("value");
    String v2 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[2]/div/div[2]/div/div[2]/div/div/input")).getAttribute("value");
    String v3 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[2]/div/div[2]/div/div[3]/div/div/input")).getAttribute("value");
    String v4 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[2]/div/div[2]/div/div[4]/div/div/input")).getAttribute("value");
    String v5 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[2]/div/div[2]/div/div[5]/div/div/input")).getAttribute("value");
    ss.append(lineNumber);
    ss.append(" Wet Air:");
    ss.append(" " + v1 + " " + v2 + " " + v3 + " " + v4 + " " + v5);
    list.addAll(Arrays.asList(lineNumber, v1, v2, v3, v4, v5));
    String vv1 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[1]/div/div/input")).getAttribute("value");
    String vv2 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[2]/div/div/input")).getAttribute("value");
    String vv3 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[3]/div/div/input")).getAttribute("value");
    String vv4 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[4]/div/div/input")).getAttribute("value");
    String vv5 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[5]/div/div/input")).getAttribute("value");
    String vv6 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[6]/div/div/input")).getAttribute("value");
    String vv7 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[6]/div/div[2]/div/div[7]/div/div/input")).getAttribute("value");
//    ss.append(" Process left:");
    ss.append(" " + vv1 + " " + vv2 + " " + vv3 + " " + vv4 + " " + vv5 + " " + vv6 + " " + vv7);
    list.addAll(Arrays.asList(vv1, vv2, vv3, vv4, vv5, vv6, vv7));
    vv1 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div[1]/div/div/input")).getAttribute("value");
    vv2 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div[2]/div/div/input")).getAttribute("value");
    vv3 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div[3]/div/div/input")).getAttribute("value");
    vv4 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div[4]/div/div/input")).getAttribute("value");
    vv5 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div[5]/div/div/input")).getAttribute("value");
    vv6 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[7]/div/div[2]/div/div[6]/div/div/input")).getAttribute("value");
    ss.append(" process right:");
    ss.append(" " + vv1 + " " + vv2 + " " + vv3 + " " + vv4 + " " + vv5 + " " + vv6);
    list.addAll(Arrays.asList(vv1, vv2, vv3, vv4, vv5, vv6));
    vv1 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[3]/div/div[2]/div/div[1]/div/div/input")).getAttribute("value");
    vv2 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[3]/div/div[2]/div/div[2]/div/div/input")).getAttribute("value");
    vv3 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[3]/div/div[2]/div/div[3]/div/div/input")).getAttribute("value");
    vv4 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[3]/div/div[2]/div/div[4]/div/div/input")).getAttribute("value");
    vv5 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[3]/div/div[2]/div/div[5]/div/div/input")).getAttribute("value");
    vv6 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[3]/div/div[2]/div/div[6]/div/div/input")).getAttribute("value");
    vv7 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[3]/div/div[2]/div/div[7]/div/div/input")).getAttribute("value");
    ss.append(" Reactivation:");
    ss.append(" " + vv1 + " " + vv2 + " " + vv3 + " " + vv4 + " " + vv5 + " " + vv6 + " " + vv7);
    list.addAll(Arrays.asList(vv1, vv2, vv3, vv4, vv5, vv6, vv7));
    vv1 = driver.findElement(By.xpath("//*[@id=\"root\"]/div/div/div[1]/div[2]/div[8]/div/div[2]/div/div/div/div/input")).getAttribute("value");
    ss.append(" RPH:");
    ss.append(" " + vv1);
    ss.append("\r\n");
    list.addAll(Arrays.asList(vv1));
    excelWriter.writeRow(list);
  }
}