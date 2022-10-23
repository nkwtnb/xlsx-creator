package org.example;

import java.io.IOException;
import java.util.concurrent.Callable;

import static java.lang.Thread.sleep;

public class FormThread implements Callable<String> {
    private final String filePath;
    private final String json;
    public FormThread(String _filePath, String _json) {
        filePath = _filePath;
        json = _json;
    }
    @Override
    public String call() throws IOException, InterruptedException{
        Form form = new Form(filePath, json);
        form.setValues();
        form.setRowValues();
        return form.writeFile();
    }
}
