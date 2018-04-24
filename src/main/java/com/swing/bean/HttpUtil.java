package com.swing.bean;

import org.apache.http.HttpEntity;
import org.apache.http.client.methods.CloseableHttpResponse;
import org.apache.http.client.methods.HttpPost;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.http.impl.conn.PoolingHttpClientConnectionManager;
import org.apache.http.util.EntityUtils;

import java.io.IOException;
import java.util.Map;

/***
 * http工具类
 */
public class HttpUtil {

    private static CloseableHttpClient httpClient;

    static {
        PoolingHttpClientConnectionManager cm = new PoolingHttpClientConnectionManager();
        cm.setMaxTotal(500);
        cm.setDefaultMaxPerRoute(10);
        httpClient = HttpClients.custom().setConnectionManager(cm).build();
    }

    /**
     * 发送Post请求
     *
     * @param url
     * @param httpEntity
     * @param header
     * @param json
     * @return
     */
    public static String Post(String url, HttpEntity httpEntity, Map<String, String> header, boolean json, Integer i) {
        //创建默认的 HttpClient 实例
        CloseableHttpResponse response = null;
        HttpPost httpPost = new HttpPost(url);
        String entity = "";
        if (json) {
            httpPost.setHeader("Content-Type", "application/json");
        }
        if (header != null && header.size() > 0) {
            for (String key : header.keySet()) {
                httpPost.setHeader(key, header.get(key));
            }
        }
        try {
            httpPost.setEntity(httpEntity);
            response = httpClient.execute(httpPost);
            int code = response.getStatusLine().getStatusCode();
            if (code == 200) {
                entity = EntityUtils.toString(response.getEntity(), "UTF-8");
            } else {
                System.out.println("[http-util-post]post status code err: " + response.getStatusLine().getReasonPhrase());
            }
        } catch (Exception e) {
            System.out.println("[http-util-post] " + url + " post err4: " + e.getMessage());
        } finally {
            //关闭连接，释放资源
            if (null != response) {
                try {
                    response.close();
                } catch (IOException e) {
                    System.out.println("[http-util-post]close err: " + e.getMessage());
                }
            }
        }
        return entity;

    }

}
