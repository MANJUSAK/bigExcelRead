package com.manjusaka.test;

/**
 * description:
 * ===>获取操作系统类型工具类
 *
 * @author manjusaka[manjusakachn@gmail.com] Created on 2018-01-15 17:33
 * @version V1.1.0
 */
public class GetOsNameUtil {
    /**
     * 定义初始变量为Linux系统
     */

    private final static String OSNAME = "linux";

    /**
     * 获取操作系统类型方法
     *
     * @return 操作系统类型
     */

    public static boolean getOsName() {
        String osName = System.getProperty("os.name").toLowerCase();
        switch (osName) {
            case OSNAME:
                return true;
            default:
                return false;
        }
    }

}
