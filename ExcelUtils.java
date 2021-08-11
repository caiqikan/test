/**
 *
 */
package org.cp;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.NumberFormat;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;
import java.util.stream.IntStream;
import java.util.stream.Stream;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.ExcelWriter;

/**
 * @author peter idolframe@gmail.com
 * @version V1.0
 * @description:
 * @date 2020年11月25日
 */
public class ExcelUtils
{
    private ExcelUtils()
    {
        throw new IllegalStateException("Utility class");
    }

    public static SimpleReadListener getReadListener(String filename)
    {
        return ExcelUtils.getReadListener(filename,0);
    }

    public static SimpleReadListener getReadListener(String filename,int sheet)
    {
        SimpleReadListener listener=new SimpleReadListener();
        try(FileInputStream fis=new FileInputStream(filename))
        {
            ExcelReader excelReader=EasyExcel.read(fis,listener).build();
            excelReader.read(EasyExcel.readSheet(sheet).build());
            excelReader.finish();
        }
        catch(IOException e)
        {
            System.out.println(filename+"..."+e);
        }
        return listener;
    }

    public static void main(String[] args)
    {
        String path="D:\\work\\项目资料\\良品铺子\\脚本开发\\报表数据核对\\RD_03021_月份商品大类销售\\202011\\";
        String a="良品.xlsx";
        String b="云徙.xlsx";
        int length=2;
        long start=System.currentTimeMillis();
        SimpleReadListener aRead=ExcelUtils.getReadListener(path+a);
        SimpleReadListener bRead=ExcelUtils.getReadListener(path+b);
        long end=System.currentTimeMillis();
        System.out.println("读取时间..."+(end-start));
        String[] keys=getKeys(length,aRead.getHeadMap());
        System.out.println("主要键值："+Arrays.toString(keys));
        doit(keys,aRead,bRead,path+"20201204.xlsx",path+"良品notIn云徙.xlsx",path+"云徙notIn良品.xlsx","良品","云徙");
    }

    public static String[] getKeys(int length,Map<Integer,String> headMap)
    {
        return headMap.values().stream().limit(length).toArray(String[]::new);
    }

    public static Map<String,Map<String,String>> copyData(String[] keys,List<Map<String,String>> datas)
    {
        LinkedHashMap<String,Map<String,String>> all=new LinkedHashMap<>();
        datas.stream().forEach(e -> all.put(Arrays.toString(getValues(keys,e)),e));
        return all;
    }

    public static void doit(String[] keys,SimpleReadListener aRead,SimpleReadListener bRead,String f1,String f2,String f3,String a,String b)
    {
        List<Map<String,String>> all=new ArrayList<>();
        List<Map<String,String>> aNotInB=new ArrayList<>();
        List<Map<String,String>> bNotInA=new ArrayList<>();
        Map<String,Map<String,String>> aMap=copyData(keys,aRead.getList());
        Map<String,Map<String,String>> bMap=copyData(keys,bRead.getList());
        String[] aKeys=aMap.keySet().toArray(new String[0]);
        for(String key:aKeys)
        {
            Map<String,String> row=aMap.get(key);
            if(bMap.containsKey(key))
            {
                all.add(getNewRow(row,bMap.get(key),keys,a,b));
                bMap.remove(key);
            }
            else
            {
                aNotInB.add(row);
            }
            aMap.remove(key);
        }
        for(Map<String,String> row:bMap.values())
        {
            bNotInA.add(row);
        }
        System.out.println(all.size());
        if(!all.isEmpty())
        {
            writeExcel(all,f1);
        }
        if(!aNotInB.isEmpty())
        {
            writeExcel(aNotInB,f2);
        }
        if(!bNotInA.isEmpty())
        {
            writeExcel(bNotInA,f3);
        }
        System.out.println(aNotInB.size());
        System.out.println(bNotInA.size());
    }

    public static Map<String,String> getNewRow(Map<String,String> arow,Map<String,String> brow,String[] keys,String a,String b)
    {
        Map<String,String> map=new LinkedHashMap<String,String>();
        // 先写主键
        Stream.of(keys).forEach(key -> map.put(key,arow.get(key)));
        //写数据
        arow.entrySet().forEach(e -> {
            String key=e.getKey();
            if (map.containsKey(key)) {

                return;
            }
            String avalue=e.getValue();
            map.put(a+"_"+key,avalue);
            // 写b表数据
            String bvalue=brow.get(key);
            map.put(b+"_"+key,bvalue);
            // 写差异
            map.put(key+"_差异--",getDif(avalue,bvalue));
        });
        // 补下b表
        brow.entrySet().forEach(e -> {
            String key=e.getKey();
            if(map.containsKey(key)||map.containsKey(b+"_"+key)) return;
            map.put(b+"_"+key,e.getValue());
        });
        return map;
    }

    public static String getDif(String avalue,String bvalue)
    {
        if(avalue==null) return bvalue==null?"0":"1";
        if(avalue.equals(bvalue)) return "0";
        if(isInteger(avalue)&&isInteger(bvalue))
        {
            return String.valueOf(sub(toDouble(avalue),toDouble(bvalue)));
        }
        return "1";
    }

    private static final Pattern PATTERN=Pattern.compile("^[+-]?[0-9]*[\\.]*[0-9]*$");
    private static final Pattern PATTERN2=Pattern.compile("^[+-]?[0-9]*[\\.]*[0-9]*%$");

    public static boolean isInteger(String str)
    {
        return str!=null&&(PATTERN.matcher(str).find()||PATTERN2.matcher(str).find());
    }

    public static Double toDouble(String str)
    {
        if(PATTERN.matcher(str).find()) return Double.parseDouble(str);
        if(PATTERN2.matcher(str).find()) try
        {
            return NumberFormat.getPercentInstance().parse(str).doubleValue();
        }
        catch(ParseException e)
        {
            return Double.parseDouble(str.trim().substring(0,str.length()-1))/100.d;
        }
        return 0d;
    }

    public static void writeExcel(List<Map<String,String>> data1,String filename)
    {
        List<List<String>> writerData=new ArrayList<>();
        // 创建第一行
        Map<String,String> title=data1.get(0);
        String[] keys=title.keySet().toArray(new String[0]);
        int c=title.size();
        List<String> head=new ArrayList<>();
        writerData.add(head);
        for(int j=0;j<c;j++)
        {
            if(keys[j].endsWith("_差异--"))
                head.add("差异");
            else
                head.add(keys[j]);
        }
        for(int i=0;i<data1.size();i++)
        {
            // 从第一行开始写入
            Map<String,String> data=data1.get(i);
            List<String> row=new ArrayList<>();
            for(int j=0;j<c;j++)
            {
                String value=data.get(keys[j]);
                row.add(value);
            }
            writerData.add(row);
        }
        writeExcel(writerData,filename,"data");
    }

    public static void writeExcel(List<List<String>> writerData,String filename,String sheet)
    {
        ExcelWriter writer=EasyExcel.write(new File(filename)).build();
        writer.write(writerData,EasyExcel.writerSheet(sheet).build());
        writer.finish();
    }

    private static String[] getValues(String[] keys,Map<String,String> row)
    {
        return IntStream.range(0,keys.length).mapToObj(i -> row.get(keys[i])).toArray(String[]::new);
    }

    /**
     * 减法
     */
    public static double sub(double d1,double d2)
    {
        return new BigDecimal(Double.toString(d1)).subtract(new BigDecimal(Double.toString(d2))).doubleValue();
    }
}
