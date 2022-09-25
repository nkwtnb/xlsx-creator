package org.example;

import java.io.*;
import java.net.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.security.CodeSource;
import java.security.ProtectionDomain;

public class Main {
    public static void main(String[] args) throws Exception {
        if (args.length != 3) {
            throw new Exception("パラメータの数が違います。[0]ユーザーID, [1]ファイル名, [2]データ（json）を渡してください。");
        }
        Path path = getApplicationPath(Main.class).getParent().resolve("files/" + args[0] + "/" + args[1]);
        String filePathString = path.toString() + ".xlsx";
        if (!Files.exists(Paths.get(filePathString))) {
            throw  new FileNotFoundException("対象のファイルが存在しません。" + filePathString);
        }
        String json = args[2];
        Form form = new Form(filePathString, json);
        form.setValues();
//        form.setRowValues();
        String base64 = form.writeFile();
        System.out.println(base64);
    }
    public static Path getApplicationPath(Class<?> cls) throws URISyntaxException {
        ProtectionDomain pd = cls.getProtectionDomain();
        CodeSource cs = pd.getCodeSource();
        URL location = cs.getLocation();
        URI uri = location.toURI();
        return Paths.get(uri);
    }
}