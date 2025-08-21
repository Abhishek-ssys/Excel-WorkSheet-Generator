package com.excel;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Random;

public class TestFileCreator {
    public static void main(String[] args) throws IOException {
        String currentDir = System.getProperty("user.dir");
        File targetdir=  new File(currentDir,"target");
        if (!targetdir.exists()) {
            targetdir.mkdirs();
        }
        Random random = new Random();

        // Sample Jira project keys to randomize
        String[] projectKeys = {"HDAG", "HCAG", "HCCUG", "HDCUG", "APIGW"};

        for (int day = 1; day <= 30; day++) {
            String fileName = String.format("aug_%02d_2025.txt", day);
            File file = new File(targetdir, fileName);

            try (FileWriter writer = new FileWriter(file)) {
                // Write a few random tasks
                for (int taskNum = 1; taskNum <= random.nextInt(4) + 1; taskNum++) {
                    writer.write("Task " + (char)('A' + day % 26) + taskNum + "\n");
                }

                // Add 1–2 Jira tickets randomly
                int jiraCount = random.nextInt(2) + 1;
                for (int j = 0; j < jiraCount; j++) {
                    String project = projectKeys[random.nextInt(projectKeys.length)];
                    int ticketNum = 1000 + random.nextInt(9000);

                    // Mix casing for "jira"
                    String jiraWord = (random.nextBoolean() ? "Jira" : (random.nextBoolean() ? "jira" : "JIRA"));

                    writer.write(jiraWord + " " + project + " " + ticketNum + "\n");
                }
            }

            System.out.println("Created file: " + file.getAbsolutePath());
        }

        System.out.println("✅ 30 test files generated successfully."+ targetdir.getAbsolutePath());
    }
}
