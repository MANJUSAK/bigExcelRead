package com.manjusaka.test;

import cn.afterturn.easypoi.handler.impl.ExcelDataHandlerDefaultImpl;

/**
 * description:
 * ===>
 * company
 *
 * @author manjusaka[manjusakachn@gmail.com] Created on 2018-06-14 17:05
 * @version V1.0.0
 */
public class CompanyHandler extends ExcelDataHandlerDefaultImpl<Company> {

    @Override
    public Object importHandler(Company obj, String name, Object value) {
        System.out.println(name + "-" + value);
        return super.importHandler(obj, name, value);
    }
}
