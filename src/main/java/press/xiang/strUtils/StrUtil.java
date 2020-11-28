package press.xiang.strUtils;

/**
 * @author Xiang想
 * @title: StrUtil
 * @projectName POITest
 * @description: TODO
 * @date 2020/11/29  0:21
 */
public class StrUtil {
    public static void main(String[] args) {
        String s = "null";
        System.out.println(isNotBlank(s)?"非空":"是空");
    }

    public static boolean isNotBlank(String s){
        return null!=s&&!"".equals(s)&&!"undefined".equals(s)&&!"null".equals(s);
    }
}
