package ere.ere.dirful.util;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

/**
 * 该类根据传来的json解析规则进行解析
 * 对于json字符串，只有两类情况，一种是数据组，一种是单纯key-value形式
 * 所以解析规则可以定义为key1[n].key2.key3.key4[n].key5[1].key6这种形式。
 * 其中[n]表示变量，此处要循环要循环取值；具体[1]等数字表示单取某个数组索引值
 * @author Administrator
 *
 */
public class JsonParser {
	 public static void main(String[] args) {
	    	String people = "{ \"programmers\": [ { \"firstName\": \"Brett\", \"lastName\":\"McLaughlin\", \"email\": \"aaaa\" }," +
	    			"{ \"firstName\":\"Jason\", \"lastName\":\"Hunter\", \"email\":\"bbbb\" }," +
	    			"{ \"firstName\": \"Elliotte\", \"lastName\":\"Harold\", \"email\": \"cccc\" }]," +
	    			"\"authors\": [" +
	    			"{ \"firstName\": \"Isaac\", \"lastName\": \"Asimov\", \"genre\": \"science fiction\" }," +
	    			"{ \"firstName\": \"Tad\", \"lastName\": \"Williams\", \"genre\": \"fantasy\" }," +
	    			"{ \"firstName\": \"Frank\", \"lastName\": \"Peretti\", \"genre\": \"christian fiction\" }]," +
	    			" \"musicians\": [ " +
	    			"{ \"firstName\": [\"Eric\",\"Eric2\"], \"lastName\": \"Clapton\", \"instrument\": \"guitar\" }," +
	    			"{ \"firstName\": \"Sergei\", \"lastName\": \"Rachmaninoff\", \"instrument\": \"piano\" }] }";
	    	
	    	 String jsonExpress1 ="programmers[n].firstName"; // 解析json表达式 (phaser json express)
	    	 String jsonExpress2 ="authors[0].firstName";
	    	 String jsonExpress3 ="musicians[0].firstName[n]";
	    	 JSONObject obj = JSONObject.fromObject(people); // 将String数组转化成json (String to json);
	    	 String jsonExp[] = jsonExpress3.split("\\.");   // 将解析式"."拆解
	    	// 取[]号中的数字
	    	 String arrRegexDigit = "\\[(\\d+)\\]";
	    	 Pattern patternDigit = Pattern.compile(arrRegexDigit);
//	    	// 取[]号中的变量n
//	    	 String arrRegex2="\\[(n)\\]";
//	    	 Pattern patternVariable = Pattern.compile(arrRegex2);
	    	 // 定义一个list存放get key以后的数据
	    	 List<JSONObject> objList = new ArrayList<JSONObject>();
	    	 objList.add(obj);
	    	 for(String str : jsonExp) {
	    		 // 定义一个临时list存放当前循环得到的jsonobject
	    		 List<JSONObject> tempList = new ArrayList<JSONObject>();
	    		 for(JSONObject jsonObj : objList) {
		    		 Object temp = new Object();
		    		 // 如果变量中包含"["
		    		 if(str.contains("[")) {
		    			 String key = str.substring(0,str.indexOf("["));
		    			 Matcher matcher = patternDigit.matcher(str);
		    			 if (matcher.find()) {
		    				 int index = Integer.parseInt(matcher.group(1));
		    				 JSONArray jSONArray = (JSONArray)jsonObj.get(key);
		    				 temp = jSONArray.getJSONObject(index);
		    	    		 if(temp instanceof String) {
		    	    			 System.out.println(temp.toString());
		    	    		 } else if (temp instanceof JSONObject){
		    	    			 tempList.add((JSONObject)temp);
		    	    		 }
		    	    	 } else {
		    	    		 JSONArray jSONArray = (JSONArray)jsonObj.get(key); 
		    	    		 for(int i=0 ; i < jSONArray.size() ;i++) {
		    	    			 Object myObject = jSONArray.get(i);
		    	    			 if(myObject instanceof String) {
			    	    			 System.out.println(myObject.toString());
			    	    		 } else if (myObject instanceof JSONObject){
			    	    			 tempList.add((JSONObject)myObject);
			    	    		 }
		    	    			 
		    	    		 }
		    	    	 }
		    		 } else {
		    			 temp = jsonObj.get(str);
		    			 if(temp instanceof String) {
	    	    			 System.out.println(temp.toString());
	    	    		 } else if (temp instanceof JSONObject){
	    	    			 tempList.add((JSONObject)temp);
	    	    		 }
		    		 }
	    		 }
	    		 
	    		 objList.clear();     // 清楚list存放数据
	    		 objList = tempList;  // list存放新的数据
	    		 
	    	 }
	    	 
	    	 
	    	 
	    	 
	    	 
		}
	 
	 
	    /** 
	     * 对象转换成JSON字符串 
	     *  
	     * @param obj 
	     *            需要转换的对象 
	     * @return 对象的string字符 
	     */  
	    public static String toJson(Object obj) {  
	        JSONObject jSONObject = JSONObject.fromObject(obj);  
	        return jSONObject.toString();  
	    }  
	  
	    /** 
	     * JSON字符串转换成对象 
	     *  
	     * @param jsonString 
	     *            需要转换的字符串 
	     * @param type 
	     *            需要转换的对象类型 
	     * @return 对象 
	     */  
	    @SuppressWarnings("unchecked")  
	    public static <T> T fromJson(String jsonString, Class<T> type) {  
	        JSONObject jsonObject = JSONObject.fromObject(jsonString);  
	        return (T) JSONObject.toBean(jsonObject, type);  
	    }  
	  
	    /** 
	     * 将JSONArray对象转换成list集合 
	     *  
	     * @param jsonArr 
	     * @return 
	     */  
	    public static List<Object> jsonToList(JSONArray jsonArr) {  
	        List<Object> list = new ArrayList<Object>();  
	        for (Object obj : jsonArr) {  
	            if (obj instanceof JSONArray) {  
	                list.add(jsonToList((JSONArray) obj));  
	            } else if (obj instanceof JSONObject) {  
	                list.add(jsonToMap((JSONObject) obj));  
	            } else {  
	                list.add(obj);  
	            }  
	        }  
	        return list;  
	    }  
	  
	    /** 
	     * 将json字符串转换成map对象 
	     *  
	     * @param json 
	     * @return 
	     */  
	    public static Map<String, Object> jsonToMap(String json) {  
	        JSONObject obj = JSONObject.fromObject(json);  
	        return jsonToMap(obj);  
	    }  
	  
	    /** 
	     * 将JSONObject转换成map对象 
	     *  
	     * @param json 
	     * @return 
	     */  
	    public static Map<String, Object> jsonToMap(JSONObject obj) {  
	        Set<?> set = obj.keySet();  
	        Map<String, Object> map = new HashMap<String, Object>(set.size());  
	        for (Object key : obj.keySet()) {  
	            Object value = obj.get(key);  
	            if (value instanceof JSONArray) {  
	                map.put(key.toString(), jsonToList((JSONArray) value));  
	            } else if (value instanceof JSONObject) {  
	                map.put(key.toString(), jsonToMap((JSONObject) value));  
	            } else {  
	                map.put(key.toString(), obj.get(key));  
	            }  
	  
	        }  
	        return map;  
	    }  
	
	 
}
