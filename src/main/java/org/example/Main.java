package org.example;

import java.io.*;
import java.net.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.security.CodeSource;
import java.security.ProtectionDomain;
import java.util.concurrent.*;

public class Main {
    public static final int EXIT_STATUS_NORMAL = 0;
    public static final int EXIT_STATUS_ABNORMAL = 1;
    public static final int EXIT_STATUS_TIMEOUT = 2;
    public static void main(String[] args) throws Exception {
        if (args.length != 3) {
            throw new Exception("パラメータの数が違います。[0]ファイルパス, [1]データ（json）, [2]出力ファイル削除フラグ（0 or 1）を渡してください");
        }
        Path path = getApplicationPath(Main.class).getParent().resolve(args[0]);
        String filePathString = path.toString();
        if (!Files.exists(Paths.get(filePathString))) {
            throw  new FileNotFoundException("対象のファイルが存在しません。" + filePathString);
        }
        String json = args[1];
        Boolean shouldDelete = (Integer.parseInt(args[2]) == 1);
        ThreadFactory daemon = new ThreadFactory() {
            @Override
            public Thread newThread(Runnable r) {
                Thread t = new Thread(r);
                // when main thread exit, also daemon thread exit
                t.setDaemon(true);
                return t;
            }
        };
        ExecutorService es = Executors.newSingleThreadExecutor(daemon);
        String result = "";
        try {
            Future<String> future = es.submit(new FormThread(filePathString, json, shouldDelete));
            try {
                result = future.get(5, TimeUnit.SECONDS);
            } catch (TimeoutException te) {
                System.exit(EXIT_STATUS_TIMEOUT);
            }
        } finally {
            es.shutdown();
        }
        System.out.print(result);
    }
    public static Path getApplicationPath(Class<?> cls) throws URISyntaxException {
        ProtectionDomain pd = cls.getProtectionDomain();
        CodeSource cs = pd.getCodeSource();
        URL location = cs.getLocation();
        URI uri = location.toURI();
        return Paths.get(uri);
    }
}