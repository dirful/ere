package ere.ere.dirful.handle;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class JsonHandle {
	/**
	 * 将json数据导出成excel
	 * @param tempFilePath   模板路径
	 * @param jsonString     json字符串
	 * @param outputFilePath 输出文件路径
	 * @throws IOException
	 */
	public void josnToExcel(String tempFilePath, String jsonString, String outputFilePath) throws IOException {
		
		ExcelHandle handle = new  ExcelHandle();
		
	        
//	    	 String jsonExpress1 ="programmers[n].firstName"; // 解析json表达式 (phaser json express)
//	    	 String jsonExpress2 ="authors[0].firstName";
//	    	 String jsonExpress3 ="musicians[n].firstName[n].BB";
//	    	 String jsonExpress4 ="musicians[n].firstName[n].AA";
//	    	 List<String> list = JsonParser.getJsonVale(jsonExpress3, people);
	    	 handle.writeJsonData2(tempFilePath,jsonString,new ArrayList(),0);
	    	 File file = new File(outputFilePath);
	         OutputStream os = new FileOutputStream(file);
	         //写到输出流并关闭资源
	         handle.writeAndClose(tempFilePath, os);
	          
	         os.flush();
	         os.close();
		// 读取模板每一行值
		// 根据一行值判断是该进行合并还是逐行显示
	}
	
	public static void main(String[] args) {
		try {
			String tempFilePath = ExcelHandle.class.getResource("/test2.xlsx").getPath();
			String outputFilePath = "d:/data.xlsx";
			// 准备json数据
			 String people = "{ \"programmers\": [ { \"firstName\": \"Brett\", \"lastName\":\"McLaughlin\", \"email\": \"aaaa\" }," +
		    			"{ \"firstName\":\"Jason\", \"lastName\":\"Hunter\", \"email\":\"bbbb\" }," +
		    			"{ \"firstName\": \"Elliotte\", \"lastName\":\"Harold\", \"email\": \"cccc\" }]," +
		    			"\"authors\": [" +
		    			"{ \"firstName\": \"Isaac\", \"lastName\": \"Asimov\", \"genre\": \"science fiction\" }," +
		    			"{ \"firstName\": \"Tad\", \"lastName\": \"Williams\", \"genre\": \"fantasy\" }," +
		    			"{ \"firstName\": \"Frank\", \"lastName\": \"Peretti\", \"genre\": \"christian fiction\" }]," +
		    			" \"musicians\": [ " +
		    			"{ \"firstName\": [{\"AA\":\"Eric\",\"BB\":\"Eric2\"},{\"AA\":\"Fric\",\"BB\":\"Fric2\"}], \"lastName\": \"Clapton\", \"instrument\": \"guitar\" }," +
		    			"{ \"firstName\": [{\"AA\":\"Sergei\",\"BB\":\"Sergei2\"},{\"AA\":\"Tric\",\"BB\":\"Tric2\"}], \"lastName\": \"Rachmaninoff\", \"instrument\": \"piano\" }] }";
			new JsonHandle().josnToExcel(tempFilePath, people, outputFilePath);
			
//			List<Integer> list =  new ArrayList<Integer>();
//			list.add(1);
//			list.add(9);
//			list.add(8);
//			list.add(3);
//			
//			for(Integer a:list) {
//				System.out.println(a);
//			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
