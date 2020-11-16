package cn.ncodev.cache;

import org.apache.poi.util.IOUtils;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;

/**
 * 加载模板
 * @author wenbing.li
 * @date 2020/6/24 15:19
 * @since 1.0
 **/
public class FileLoader {
    private static final String HTTP = "http";
    /**
     * 连接网络文件超时时间
     */
    private static final int CONNECT_TIMEOUT = 30 * 1000;
    /**
     * 读取网络文件超时时间
     */
    private static final int READ_TIMEOUT = 60 * 1000;
    /**
     * 文件大小限制
     */
    private static final int MAX_BYTE__SIZE = 1024 * 1;

    /**
     * 加载本地 或 网络中的模板
     * @param template 网络地址 或 模板路径（先对路径 | 绝对路径）
     * @return 模板的字节码数组
     */
    byte[] loaderTemplate(String template) throws IOException {
        return loaderFile(template);
    }

    /**
     * 加载本地 或 网络中的文件
     * @param urlOrPath 网络地址 或 文件路径（先对路径 | 绝对路径）
     * @return 文件的字节码数组
     */
    public static byte[] loaderFile(String urlOrPath) throws IOException {
        InputStream inputStream = null;
        try (ByteArrayOutputStream byteArray = new ByteArrayOutputStream()){
            if (urlOrPath.toLowerCase().startsWith(HTTP)) {
                // 网络地址
                URL url = new URL(urlOrPath);
                HttpURLConnection connection = (HttpURLConnection) url.openConnection();
                connection.setConnectTimeout(CONNECT_TIMEOUT);
                connection.setReadTimeout(READ_TIMEOUT);
                connection.setDoInput(true);
                inputStream = connection.getInputStream();
            } else {
                urlOrPath = urlOrPath.replace("//", File.separator).replace("/",File.separator);
                try {
                    // 绝对路径读取文件
                    inputStream = new FileInputStream(urlOrPath);
                } catch (FileNotFoundException e){
                    // 相对路径读取文件
                    inputStream = FileLoader.class.getClassLoader().getResourceAsStream(urlOrPath);
                }
            }
            if(inputStream == null){
                throw new IOException("File read failed !");
            }
            byte[] buffer = new byte[MAX_BYTE__SIZE];
            int len;
            while ((len = inputStream.read(buffer)) > -1) {
                byteArray.write(buffer, 0, len);
            }
            byteArray.flush();
            return byteArray.toByteArray();
        } finally {
            IOUtils.closeQuietly(inputStream);
        }
    }


}
