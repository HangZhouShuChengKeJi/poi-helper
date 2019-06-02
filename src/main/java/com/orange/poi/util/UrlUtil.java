package com.orange.poi.util;

/**
 * url 工具类
 *
 * @author 小天
 * @date 2019/5/31 21:32
 */
public class UrlUtil {

    /**
     * 正斜杠
     */
    public static final char FORWARD_SLASH = '/';

    /**
     * 从 url 里获取文件扩展名
     * <pre>
     * null                                             -> null
     * ""                                               -> ""
     * "index"                                          -> ""
     * "index.jpg"                                      -> "jpg"
     * "/a/b/c/index.jpg"                               -> "jpg"
     * "/a/b/c/index.jpg?name=index.png"                -> "jpg"
     * "/a/b/c/index.jpg#name=index.png"                -> "jpg"
     * "/a/b/c/index.jpg#v=1?name=index.png"            -> "jpg"
     * "/a/b/c/index.jpg?v=1#name=index.png"            -> "jpg"
     * "www.a.com/a/b/c/index.jpg"                      -> "jpg"
     * "www.a.com/a/b/c/index.jpg?name=index.png"       -> "jpg"
     * "www.a.com/a/b/c/index.jpg#name=index.png"       -> "jpg"
     * "www.a.com/a/b/c/index.jpg#v=1?name=index.png"   -> "jpg"
     * "www.a.com/a/b/c/index.jpg?v=1#name=index.png"   -> "jpg"
     * </pre>
     *
     * @param url 原始文件名
     *
     * @return 文件扩展名
     */
    public static String getExNameFromUrl(String url) {
        if (url == null) {
            return null;
        }

        int e;
        if ((e = url.indexOf('#')) > 0) {
            int t;
            if ((t = url.indexOf('?')) > 0 && t < e) {
                e = t;
            }
        } else {
            e = url.indexOf('?');
        }

        int p;
        if (e > 0) {
            // 找最后一个 '/' 后面的，而且在 e 之前的，最后一个 '.'
            if ((p = url.lastIndexOf('.', e)) > 0 && p > url.lastIndexOf('/', e)) {
                return url.substring(p + 1, e);
            }
        } else if (e < 0) {
            // 找最后一个 '/' 后面的，最后一个 '.'
            if ((p = url.lastIndexOf('.')) > 0 && p > url.lastIndexOf('/')) {
                return url.substring(p + 1);
            }
        }
        return "";
    }
}

