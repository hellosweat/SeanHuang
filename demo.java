
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import com.mongodb.*;
import com.mongodb.client.FindIterable;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoDatabase;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import org.bson.Document;

public class importData {



    public static void main(String[] args) throws BiffException, IOException
    {
        MongoClient mongo = new MongoClient("localhost", 27017);
        MongoDatabase database = mongo.getDatabase("aliexpress");
        MongoCollection<Document> collection = database.getCollection("customer");



        File xlsFile = new File("/Users/seanhuang/Documents/Sean/data/list.xls");
        // 获得工作簿对象
        Workbook workbook = Workbook.getWorkbook(xlsFile);
        // 获得所有工作表
        Sheet[] sheets = workbook.getSheets();
        // 遍历工作表
        if (sheets != null)
        {
            for (Sheet sheet : sheets)
            {
                // 获得行数
                int rows = sheet.getRows();
                // 获得列数
                int cols = sheet.getColumns();
                // 读取数据
                for (int row = 0; row < rows; row++)
                {
                    List<String> data = new ArrayList<String>();
                    for (int col = 0; col < cols; col++)
                    {
                        data.add(sheet.getCell(col, row).getContents());
                    }
                    String realPhone = phoneNumberHandle(data.get(4), data.get(5));
                    Document document = new Document("_id", data.get(0))
                            .append("amount",data.get(1))
                            .append("name",data.get(2))
                            .append("country",data.get(3))
                            .append("phone",realPhone);
                    try{
                        collection.insertOne(document);
                        System.out.println("insert one ");
                    }catch (Exception e){
                        System.out.println("has existed");
                    }

                }
            }
        }
        workbook.close();
    }






    public static String deleteStartZero(String number){
        int numLen = number.length();
        char numStrs[] = number.toCharArray();
        int index = 0;
        for (int i = 0; i < numLen; i++) {
            if ('0' != numStrs[i]) {
                index = i;// 找到非零字符串并跳出
                break;
            }
        }
        String result = number.substring(index,numLen);
        return result;
    }

    public static String phoneNumberHandle(String areaCode, String phoneNum) {
        String customePhone = null;
        //首先判断区号
        if(areaCode.length() == 0){
            customePhone = "+" + deleteStartZero(phoneNum).replaceAll("(?:\\/|\\-|\\+)","");
        }
        if(areaCode.length() <= 4 && areaCode.length() > 0){
            if(!phoneNum.startsWith("00")){
                customePhone = "+" + areaCode.replaceAll("\\+","") + phoneNum.replaceAll("(?:\\/|\\-|\\+)","");
            } else {
                customePhone = "+" + deleteStartZero(phoneNum).replaceAll("(?:\\/|\\-|\\+)","");
            }
        }else{
            customePhone = "+" + deleteStartZero(areaCode).replaceAll("(?:\\/|\\-|\\+)","");
        }
        return customePhone;
    }
}
