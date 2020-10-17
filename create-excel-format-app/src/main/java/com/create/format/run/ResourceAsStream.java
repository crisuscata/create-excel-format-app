package com.create.format.run;

import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

public class ResourceAsStream {

	public static void main(String[] args) {

		ResourceAsStream obj = new ResourceAsStream();
        System.out.println(obj.getFilePathToSave());

    }

    public String getFilePathToSave() {

        Properties prop = new Properties();

        String result = "";

        try (InputStream inputStream = getClass()
				.getClassLoader().getResourceAsStream("config.properties")) {

            prop.load(inputStream);
            result = prop.getProperty("json.filepath");

        } catch (IOException e) {
            e.printStackTrace();
        }

        return result;

    }

}
