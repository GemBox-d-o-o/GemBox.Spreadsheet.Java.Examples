package com.gembox.examples;

import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public class Util {

    public static String resourcesFolder() {
        Path folder = Paths.get(System.getProperty("user.dir"));
        while (folder != null) {
            Path potentialResourcesFolder = folder.resolve("resources");
            if (Files.exists(potentialResourcesFolder))
                return potentialResourcesFolder.toString() + "/";
            folder = folder.getParent();
        }

        throw new IllegalStateException("Resources folder was not found");
    }
}
