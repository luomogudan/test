import com.alibaba.fastjson.JSON;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.MalformedURLException;
import java.net.URL;
import java.net.URLConnection;
import java.util.Map;

public class ExcelUtil {

    /**
     * 解析excel并增添一列
     */
    public void outExcel(String path) {
        try {
            File f = new File(path);
            InputStream is = new FileInputStream(f);
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
            for (int numSheet = 0; numSheet < xssfWorkbook.getNumberOfSheets(); numSheet++) {
                XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(numSheet);

                if (xssfSheet == null) {
                    continue;
                }
                for (int rowNum = 1; rowNum <= xssfSheet.getLastRowNum(); rowNum++) {
                    XSSFRow xssfRow = xssfSheet.getRow(rowNum);
                    if (null != xssfRow) {
                        XSSFCell xssPhone = xssfRow.getCell(0);
                        String phone = xssPhone.getStringCellValue();
                        String json = load("https://tcc.taobao.com/cc/json/mobile_tel_segment.htm?tel=" + phone.trim());
                        String address = this.getAddress(json);
                        //给原有excel添加一列
                        addCell(f, xssfWorkbook, xssfSheet, rowNum, address);
                    }

                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void addCell(File f, XSSFWorkbook xssfWorkbook, XSSFSheet xssfSheet, int rowNum, String address) throws IOException {
        XSSFCell cell= xssfSheet.getRow(rowNum).createCell(3);
        xssfSheet.getRow(0).createCell(3).setCellValue("归属地");
        cell.setCellValue(address);
        OutputStream fOut = new FileOutputStream(f);
        // 把相应的Excel 工作簿存盘
        xssfWorkbook.write(fOut);
        fOut.flush();
        // 操作结束，关闭文件
        fOut.close();
    }

    /**
     * 归属地查询
     *
     * @param url
     * @return
     */
    private String load(String url) {
        StringBuilder json = new StringBuilder();
        try {
            URL oracle = new URL(url);
            URLConnection yc = oracle.openConnection();
            BufferedReader in = new BufferedReader(new InputStreamReader(
                    yc.getInputStream(), "gb2312"));
            String inputLine = null;
            while ((inputLine = in.readLine()) != null) {
                json.append(inputLine);
            }
            in.close();
        } catch (MalformedURLException e) {
        } catch (IOException e) {
        }
        return json.toString();
    }

    /**
     * 获取地址
     *
     * @param json
     * @return
     */
    private String getAddress(String json) {
        if (null == json) return null;
        String[] addresses = json.split("=");
        System.out.println(json);
        Object str = null;
        String address = addresses[1];
        Map maps = (Map) JSON.parse(address);
        for (Object o : maps.entrySet()) {
            if (o.toString().split("=")[0].equals("carrier"))
            {
                str = o.toString().split("=")[1];
            }

        }

        return str.toString();
    }





    public static void main(String[] args) {
        //13401935926 jiangsu
        //13402260873  shandong
        // 13550022805 sichuang
        //13542730641 guagndong
        // 13404804397  neimenggu
        new ExcelUtil().outExcel("E://aaa.xlsx");
        String json = new ExcelUtil().
                load("https://tcc.taobao.com/cc/json/mobile_tel_segment.htm?tel=13436002009");
        System.out.println(new ExcelUtil().getAddress(json)+"---");
    }

}
