package com.excel.util;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.Map;

public class CopyUtil {
	
	private static Map<String,Object> filter=new HashMap<>();//过滤掉的属性
	static{
		filter.put("id", "1");
		filter.put("serialVersionUID", "1");
	}
	
	/**
	 * from将to的属性覆盖。from属性为null而to的属性不为null，复制完之后to的属性为null;
	 * @param from
	 * @param to
	 */
	@SuppressWarnings({ "rawtypes", "unchecked" })
	public static void copy(Object from,Object to) {
		Class clazz=from.getClass();
		Field[] fields=clazz.getDeclaredFields();
		try{
			if(fields != null){
				for(Field field: fields){
					String fieldName=field.getName();
					if(!filter.containsKey(fieldName)){
						String firstLetter=fieldName.charAt(0)+"";
						firstLetter=firstLetter.toUpperCase();
						String getMethodName="get"+firstLetter+fieldName.substring(1);
						String setMethodName="set"+firstLetter+fieldName.substring(1);
						
						Object param=null;
						Method getMethod=clazz.getMethod(getMethodName);
						param=getMethod.invoke(from);
						Method setMethod=to.getClass().getMethod(setMethodName,getMethod.getReturnType());
						setMethod.invoke(to, param);
					}
				}
			}
		}catch(Exception nme){
			nme.printStackTrace();
			throw new RuntimeException("复制失败");
		}
	}
	
	/**
	 * from将to的属性覆盖。from属性为null而to的属性不为null，复制完之后to的属性不变;
	 * @param from
	 * @param to
	 */
	@SuppressWarnings({ "rawtypes", "unchecked" })
	public static void copyIgnoreNull(Object from,Object to) {
		Class clazz=from.getClass();
		Field[] fields=clazz.getDeclaredFields();
		try{
			if(fields != null){
				for(Field field: fields){
					String fieldName=field.getName();
					if(!filter.containsKey(fieldName)){
						String firstLetter=fieldName.charAt(0)+"";
						firstLetter=firstLetter.toUpperCase();
						String getMethodName="get"+firstLetter+fieldName.substring(1);
						String setMethodName="set"+firstLetter+fieldName.substring(1);
						
						Object param=null;
						Method getMethod=clazz.getMethod(getMethodName);
						param=getMethod.invoke(from);
						if(param != null){
							Method setMethod=to.getClass().getMethod(setMethodName,getMethod.getReturnType());
							setMethod.invoke(to, param);
						}
					}
				}
			}
		}catch(Exception nme){
			nme.printStackTrace();
			throw new RuntimeException("复制失败");
		}
	}
	
	@SuppressWarnings("unchecked")
	public static <T> T copyFrom(T from){
		Object o=null;
		byte[] b=null;
		ByteArrayOutputStream baos=new ByteArrayOutputStream();
		ObjectOutputStream oos = null;
		ByteArrayInputStream bais=null;
		ObjectInputStream ois=null;
		try {
			oos = new ObjectOutputStream(baos);
			oos.writeObject(from);
			b=baos.toByteArray();
			
			bais=new ByteArrayInputStream(b);
			ois=new ObjectInputStream(bais);
			o=ois.readObject();
		} catch (IOException e1) {
			e1.printStackTrace();
			throw new RuntimeException("复制失败");
		} catch (ClassNotFoundException e) {
			e.printStackTrace();
			throw new RuntimeException("复制失败");
		}finally{
			if(oos != null){
				try {
					oos.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if(ois != null){
				try {
					ois.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			System.out.println(oos==null);
			System.out.println(baos==null);
		}
		return (T)o;
	}
	
}
