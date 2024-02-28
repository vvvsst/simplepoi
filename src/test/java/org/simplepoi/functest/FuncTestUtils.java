package org.simplepoi.functest;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.Random;

public abstract class FuncTestUtils {
    static BigDecimal low = new BigDecimal("1");
    static BigDecimal high = new BigDecimal("100");
    static Random random = new Random();
    public static BigDecimal randomBigDecimal(){
        BigDecimal range = high.subtract(low);
        return range.multiply(BigDecimal.valueOf(Math.random())).add(low).setScale(2, RoundingMode.UP);
    }

    public static Integer randomZeroOrOne(){
        return random.nextInt(2);
    }

    public static Integer randomAge(){
        return random.nextInt(20)+10;
    }
}
