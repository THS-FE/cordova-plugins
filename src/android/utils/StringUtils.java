package cn.com.ths.exportdoc.thsexportdoc.utils;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class StringUtils {
    private static final String FOLDER_SEPARATOR = "/";

    private static final String WINDOWS_FOLDER_SEPARATOR = "\\";

    private static final String TOP_PATH = "..";

    private static final String CURRENT_PATH = ".";

    private static final char EXTENSION_SEPARATOR = '.';

    /**
     * Extract the filename extension from the given Java resource path,
     * e.g. "mypath/myfile.txt" -> "txt".
     * @param path the file path (may be {@code null})
     * @return the extracted filename extension, or {@code null} if none
     */
    public static String getFilenameExtension(String path) {
        if (path == null) {
            return null;
        }

        int extIndex = path.lastIndexOf(EXTENSION_SEPARATOR);
        if (extIndex == -1) {
            return null;
        }

        int folderIndex = path.lastIndexOf(FOLDER_SEPARATOR);
        if (folderIndex > extIndex) {
            return null;
        }

        return path.substring(extIndex + 1);
    }

    /**
     * Test if the given {@code String} ends with the specified suffix,
     * ignoring upper/lower case.
     * @param str the {@code String} to check
     * @param suffix the suffix to look for
     * @see java.lang.String#endsWith
     */
    public static boolean endsWithIgnoreCase(String str, String suffix) {
        return (str != null && suffix != null && str.length() >= suffix.length() &&
                str.regionMatches(true, str.length() - suffix.length(), suffix, 0, suffix.length()));
    }
    /**判断字符串是否为空
     * @param str
     * @return true:字符串为空
     * false:不为空
     */
    public static boolean isNull(String str) {
        boolean result = false;
        if (str != null && !"".equals(str) && !"null".equalsIgnoreCase(str) && !"NaN".equalsIgnoreCase(str)) {
            result = false;
        } else {
            result = true;
        }
        return result;
    }

    /**
     * 根据正则替换内容
     * @param patt
     * @param input
     * @param value
     * @return
     */
    public static String replaceContent(String patt, String input, String value){
        // 创建一个正则表达式模式，用以匹配一个单词（\b=单词边界）
        // 从正则表达式实例中运行方法并查看其如何运行
        Pattern r = Pattern.compile(patt);
        Matcher m = r.matcher(input);
        // appendReplacement方法
        m.reset();
        StringBuffer sb = new StringBuffer();
        while (m.find()) {
            // 将匹配之前的字符串复制到sb,再将匹配结果替换为："favour"，并追加到sb
            m.appendReplacement(sb, value);
        }
        m.appendTail(sb);
        return sb.toString();
    }
}
