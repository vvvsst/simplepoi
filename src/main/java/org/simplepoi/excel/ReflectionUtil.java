/**
 * Copyright 2013-2015 JEECG (jeecgos@163.com)
 * <p>
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * <p>
 * http://www.apache.org/licenses/LICENSE-2.0
 * <p>
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.simplepoi.excel;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.*;

import org.simplepoi.excel.annotation.ExcelCollection;
import org.simplepoi.excel.annotation.ExcelEntity;
import org.simplepoi.excel.annotation.ExcelField;
import org.simplepoi.excel.constant.PoiBaseConstants;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * AutoPoi 的公共基础类
 *
 * @author JEECG
 * @date 2015年4月5日 上午12:59:22
 */
public abstract class ReflectionUtil {

    private static final Logger LOGGER = LoggerFactory.getLogger(ReflectionUtil.class);

    private ReflectionUtil() {

    }

    /**
     * 获取class的 包括父类的
     *
     * @param clazz
     * @return
     */
    public static Field[] getClassFields(Class<?> clazz) {
        List<Field> list = new ArrayList<Field>();
        Field[] fields;
        do {
            fields = clazz.getDeclaredFields();
            Collections.addAll(list, fields);
            clazz = clazz.getSuperclass();
        } while (clazz != Object.class && clazz != null);
        return list.toArray(fields);
    }


    /**
     * 获取get方法 通过EXCEL注解exportConvert判断是否支持值的转换
     */
    public static Method getGetMethod(String name, Class<?> pojoClass, boolean convert) throws Exception {
        StringBuffer getMethodName = new StringBuffer();
        if (convert) {
            getMethodName.append(PoiBaseConstants.CONVERT);
        }
        getMethodName.append(PoiBaseConstants.GET);
        getMethodName.append(name.substring(0, 1).toUpperCase());
        getMethodName.append(name.substring(1));
        Method method = null;
        try {
            method = pojoClass.getMethod(getMethodName.toString(), new Class[]{});
        } catch (Exception e) {
            method = pojoClass.getMethod(getMethodName.toString().replace(PoiBaseConstants.GET, PoiBaseConstants.IS), new Class[]{});
        }
        return method;
    }

    /**
     * 获取GET方法
     *
     * @param name
     * @param pojoClass
     * @return
     * @throws Exception
     */
    public static Method getGetMethod(String name, Class<?> pojoClass) throws Exception {
        StringBuffer getMethodName = new StringBuffer(PoiBaseConstants.GET);
        getMethodName.append(name.substring(0, 1).toUpperCase());
        getMethodName.append(name.substring(1));
        Method method = null;
        try {
            method = pojoClass.getMethod(getMethodName.toString(), new Class[]{});
        } catch (Exception e) {
            method = pojoClass.getMethod(getMethodName.toString().replace(PoiBaseConstants.GET, PoiBaseConstants.IS), new Class[]{});
        }
        return method;
    }

    /**
     * 获取SET方法
     *
     * @param name
     * @param pojoClass
     * @param type
     * @return
     * @throws Exception
     */
    public static Method getSetMethod(String name, Class<?> pojoClass, Class<?> type) throws Exception {

        String getMethodName = PoiBaseConstants.SET + name.substring(0, 1).toUpperCase() +
                name.substring(1);
//        try {
//            Method methods = ReflectionUtil.getGetMethod(field.getName(), pojoClass, field.getType());
//            newMethods.add(methods);
//        } catch (Exception e) {
//            e.printStackTrace();
//            LOGGER.error("method not found for field {}", field.getName());
//        }

        return pojoClass.getMethod(getMethodName, new Class[]{type});
    }

    public static Method getSetMethod(String name, Class<?> pojoClass, Class<?> type, boolean convert) throws Exception {
        String getMethodName;
        if (convert) {
            getMethodName = PoiBaseConstants.CONVERT + PoiBaseConstants.SET + name.substring(0, 1).toUpperCase() +
                    name.substring(1);
        } else {
            getMethodName = PoiBaseConstants.SET + name.substring(0, 1).toUpperCase() +
                    name.substring(1);
        }
        return pojoClass.getMethod(getMethodName, new Class[]{type});
    }
    //update-begin-author:taoyan date:20180615 for:TASK #2798 导入扩展方法，支持自定义导入字段转换规则


    //update-end-author:taoyan date:20180615 for:TASK #2798 导入扩展方法，支持自定义导入字段转换规则


    public static <T> T createObject(Class<T> clazz) {
        T obj = null;
//        Method setMethod;
        try {
            obj = clazz.newInstance();
            Field[] fields = getClassFields(clazz);
            for (Field field : fields) {
                if (field.getAnnotation(ExcelCollection.class) == null &&
                        field.getAnnotation(ExcelField.class) == null &&
                        field.getAnnotation(ExcelEntity.class) == null) {
                    continue;
                }
                if (isList(field.getType())) {
                    ExcelCollection collection = field.getAnnotation(ExcelCollection.class);
                    field.setAccessible(true);
                    field.set(obj, collection.type().newInstance());
//                    setMethod = getGetMethod(field.getName(), clazz, field.getType());
//                    setMethod.invoke(obj, collection.type().newInstance());
                }
            }
        } catch (Exception e) {
            LOGGER.error(e.getMessage(), e);
            throw new RuntimeException("创建对象异常");
        }
        return obj;
    }

    public static boolean isList(Class<?> clazz) {
        return List.class.isAssignableFrom(clazz);
    }


    /**
     * 判断是不是集合的实现类
     *
     * @param clazz
     * @return
     */
    public static boolean isCollection(Class<?> clazz) {
        return Collection.class.isAssignableFrom(clazz);
    }

    /**
     * 是不是java基础类
     *
     * @param field
     * @return
     */
    public static boolean isJavaClass(Field field) { // might need to be moved to ExcelImportServer todo
        Class<?> fieldType = field.getType();
        boolean isBaseClass = false;
        if (fieldType.isArray()) {
            isBaseClass = false;
        } else if (fieldType.isPrimitive()
                || fieldType.getPackage() == null
                || fieldType.getPackage().getName().equals("java.lang")
                || fieldType.getPackage().getName().equals("java.math")
                || fieldType.getPackage().getName().equals("java.sql")
                || fieldType.getPackage().getName().equals("java.util")
                || fieldType.getPackage().getName().equals("java.time")) {
            isBaseClass = true;
        }
        return isBaseClass;
    }
}
