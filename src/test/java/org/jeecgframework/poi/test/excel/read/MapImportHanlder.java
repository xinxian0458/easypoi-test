package org.jeecgframework.poi.test.excel.read;

import java.util.Map;

import org.jeecgframework.poi.handler.impl.ExcelDataHandlerDefaultImpl;
import org.jeecgframework.poi.util.PoiPublicUtil;

public class MapImportHanlder extends ExcelDataHandlerDefaultImpl {

    @Override
    public void setMapValue(Map<String, Object> map, String originKey, Object value) {
        if (value instanceof Double) {
            map.put(getRealKey(originKey), PoiPublicUtil.doubleToString((Double) value));
        } else {
            map.put(getRealKey(originKey), value.toString());
        }
    }

    private String getRealKey(String originKey) {
        if (originKey.equals("交易账户")) {
            return "accountNo";
        }
        if (originKey.equals("姓名")) {
            return "name";
        }
        if (originKey.equals("客户类型")) {
            return "type";
        }
        return originKey;
    }
}
