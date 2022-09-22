package org.example;

import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.core.util.JacksonFeatureSet;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.MalformedURLException;
import java.net.URI;
import java.net.URL;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

public class Main {
    public static void main(String[] args) throws IOException, InterruptedException {
        System.out.println("Hello world!");
        Form form = new Form(args[0]);
//        form.setValues();
        form.setRowValues();
        form.writeFile();

//// create a client
//        var client = HttpClient.newHttpClient();
//
//// create a request
//        var request = HttpRequest.newBuilder(
//                        URI.create("https://api.nasa.gov/planetary/apod?api_key=DEMO_KEY"))
//                .header("accept", "application/json")
//                .build();
//
//// use the client to send the request
////        var response = client.send(request, new JsonBodyHandler<>(APOD.class));
//        var response = client.send(request, HttpResponse.BodyHandlers.ofString());
//        ObjectMapper mapper = new ObjectMapper();
//        var json = "{\"key\": {\"value\": \"123\"}}";
//        Map<String, Object> map = mapper.readValue(json, HashMap.class);
//        System.out.println(map.get("key"));
//// the response:®
//        System.out.println(map);
    }
}