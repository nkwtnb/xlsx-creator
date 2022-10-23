package org.example;

import java.io.IOException;
import java.io.InterruptedIOException;
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
    public String call() {
        try {
            Form form = new Form(filePath, json);
            form.setValues();
            form.setRowValues();
            return form.writeFile();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
