package org.julianjiang.javafx.processor;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;

public class InputStreamProcessor {


    public static InputStream copyStream(InputStream inputStream) throws IOException {
        // 创建一个ByteArrayOutputStream来保存复制的数据
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        byte[] buffer = new byte[1024];
        int bytesRead;
        while ((bytesRead = inputStream.read(buffer)) != -1) {
            outputStream.write(buffer, 0, bytesRead);
        }
        outputStream.flush();

        // 将ByteArrayOutputStream转换为新的InputStream
        InputStream copiedInputStream = new ByteArrayInputStream(outputStream.toByteArray());

        return copiedInputStream;
    }
}
