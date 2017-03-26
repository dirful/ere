package ere.ere.dirful.util;
import java.beans.BeanInfo;   
import java.beans.IntrospectionException;   
import java.beans.Introspector;   
import java.beans.PropertyDescriptor;   
import java.lang.reflect.InvocationTargetException;   
import java.lang.reflect.Method;   
import java.util.HashMap;   
import java.util.Map;   
  
public class BeanToMapUtils {   
     
    /**  
     * 将一个 Map 对象转化为一个 JavaBean  
     * @param type 要转化的类型  
     * @param map 包含属性值的 map  
     * @return 转化出来的 JavaBean 对象  
     * @throws IntrospectionException  
     *             如果分析类属性失败  
     * @throws IllegalAccessException  
     *             如果实例化 JavaBean 失败  
     * @throws InstantiationException  
     *             如果实例化 JavaBean 失败  
     * @throws InvocationTargetException  
     *             如果调用属性的 setter 方法失败  
     */   
    public static Object convertMap(Class type, Map map)  {   
        Object obj = null; 
       
        try {  
            // 获取类属性   
            BeanInfo beanInfo = Introspector.getBeanInfo(type);  
             obj = type.newInstance(); // 创建 JavaBean 对象   
            // 给 JavaBean 对象的属性赋值   
            PropertyDescriptor[] propertyDescriptors =  beanInfo.getPropertyDescriptors();   
            for (int i = 0; i< propertyDescriptors.length; i++) {   
                PropertyDescriptor descriptor = propertyDescriptors[i];   
                String propertyName = descriptor.getName();   
                if (map.containsKey(propertyName)) {   
                    // 下面一句可以 try 起来，这样当一个属性赋值失败的时候就不会影响其他属性赋值。   
                    Object value = map.get(propertyName);   
                    Object[] args = new Object[1];   
                    args[0] = value;   
                    descriptor.getWriteMethod().invoke(obj, args);   
                }   
            }   
        } catch (IntrospectionException e) {  
            e.printStackTrace();  
        } catch (InstantiationException e) {  
            e.printStackTrace();  
        } catch (IllegalAccessException e) {  
            e.printStackTrace();  
        } catch (IllegalArgumentException e) {  
            e.printStackTrace();  
        } catch (InvocationTargetException e) {  
            e.printStackTrace();  
        }   
        return obj;   
    }   
  
    /**  
     * 将一个 JavaBean 对象转化为一个  Map  
     * @param bean 要转化的JavaBean 对象  
     * @return 转化出来的  Map 对象  
     * @throws IntrospectionException 如果分析类属性失败  
     * @throws IllegalAccessException 如果实例化 JavaBean 失败  
     * @throws InvocationTargetException 如果调用属性的 setter 方法失败  
     */   
    @SuppressWarnings("unchecked")  
    public static Map convertBean(Object bean)   
            {   
        Class type = bean.getClass();   
        Map returnMap = new HashMap();   
        try {  
            BeanInfo    beanInfo = Introspector.getBeanInfo(type);  
            PropertyDescriptor[] propertyDescriptors =  beanInfo.getPropertyDescriptors();   
            for (int i = 0; i< propertyDescriptors.length; i++) {   
                PropertyDescriptor descriptor = propertyDescriptors[i];   
                String propertyName = descriptor.getName();   
                if (!propertyName.equals("class")) {   
                    Method readMethod = descriptor.getReadMethod();  
                    Object result = null;  
                    try{  
                         result = readMethod.invoke(bean, new Object[0]);  
                    }catch(Exception e){  
                        e.printStackTrace();  
                    }  
                    if (result != null) {  
                        returnMap.put(propertyName, result);  
                    } else {   
                        returnMap.put(propertyName, "");   
                    }   
                }   
            }   
        } catch (IntrospectionException e) {  
            e.printStackTrace();  
        }  catch (IllegalArgumentException e) {  
            e.printStackTrace();  
        }   
        return returnMap;   
    }   
}  