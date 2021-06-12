package cn.afterturn.easypoi.util;

import org.junit.Test;

/**
 * @author JueYue
 * @date 2021-06-21-6-4
 * @since 1.0
 */
public class PoiFunctionUtilTest {

    @Test
    public void numformat (){
        System.out.println(PoiFunctionUtil.formatNumber(0,"###0.00"));
        System.out.println(PoiFunctionUtil.formatNumber(34320,"###0.00"));
        System.out.println(PoiFunctionUtil.formatNumber(34320.34,"###0.00"));
    }
}
