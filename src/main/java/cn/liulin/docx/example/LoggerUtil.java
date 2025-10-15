package cn.liulin.docx.example;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 日志工具类，使用SLF4J门面模式
 * 提供统一的日志访问接口和额外的日志功能
 * 
 * @author liulin
 */
public class LoggerUtil {
    // 项目基础包名
    private static final String BASE_PACKAGE = "cn.liulin.docx.example";
    
    /**
     * 获取指定类的Logger实例
     * 
     * @param clazz 日志记录器关联的类
     * @return Logger实例
     */
    public static Logger getLogger(Class<?> clazz) {
        // 确保日志记录器名称以项目基础包名开头
        String className = clazz.getName();
        if (!className.startsWith(BASE_PACKAGE)) {
            // 如果类不在基础包下，添加基础包前缀
            return LoggerFactory.getLogger(BASE_PACKAGE + "." + className);
        }
        return LoggerFactory.getLogger(clazz);
    }
    
    /**
     * 获取指定名称的Logger实例
     * 
     * @param name 日志记录器名称
     * @return Logger实例
     */
    public static Logger getLogger(String name) {
        // 确保日志记录器名称以项目基础包名开头
        if (!name.startsWith(BASE_PACKAGE)) {
            return LoggerFactory.getLogger(BASE_PACKAGE + "." + name);
        }
        return LoggerFactory.getLogger(name);
    }
    
    /**
     * 获取格式化的日志消息
     * 
     * @param format 格式字符串
     * @param args 参数
     * @return 格式化后的消息
     */
    public static String formatMessage(String format, Object... args) {
        if (args == null || args.length == 0) {
            return format;
        }
        
        try {
            return String.format(format, args);
        } catch (Exception e) {
            // 如果格式化失败，返回原始字符串和参数
            StringBuilder sb = new StringBuilder(format);
            sb.append(" [参数: ");
            for (int i = 0; i < args.length; i++) {
                if (i > 0) sb.append(", ");
                sb.append(args[i]);
            }
            sb.append("]");
            return sb.toString();
        }
    }
    
    /**
     * 记录方法进入日志
     * 
     * @param logger 日志记录器
     * @param methodName 方法名
     * @param params 方法参数
     */
    public static void logMethodEntry(Logger logger, String methodName, Object... params) {
        if (logger.isDebugEnabled()) {
            StringBuilder sb = new StringBuilder();
            sb.append("进入方法: ").append(methodName);
            if (params != null && params.length > 0) {
                sb.append(", 参数: ");
                for (int i = 0; i < params.length; i++) {
                    if (i > 0) sb.append(", ");
                    sb.append(params[i]);
                }
            }
            logger.debug(sb.toString());
        }
    }
    
    /**
     * 记录方法退出日志
     * 
     * @param logger 日志记录器
     * @param methodName 方法名
     * @param result 返回值
     */
    public static void logMethodExit(Logger logger, String methodName, Object result) {
        if (logger.isDebugEnabled()) {
            logger.debug("退出方法: {}, 返回值: {}", methodName, result);
        }
    }
    
    /**
     * 记录方法异常日志
     * 
     * @param logger 日志记录器
     * @param methodName 方法名
     * @param e 异常
     */
    public static void logMethodException(Logger logger, String methodName, Exception e) {
        logger.error("方法 {} 执行异常: {}", methodName, e.getMessage(), e);
    }
}