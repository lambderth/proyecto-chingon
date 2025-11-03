package datapull;

import com.opencsv.CSVWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.util.*;

public class ChatExtractor {
    public static void main(String[] args) {
        String inputFile = "chat_data.xlsx"; // your .xlsx file
        String outputFile = "output.csv";

        try (
            FileInputStream fis = new FileInputStream(inputFile);
            Workbook workbook = new XSSFWorkbook(fis);
            FileWriter fw = new FileWriter(outputFile);
            CSVWriter writer = new CSVWriter(fw)
        ) {
            Sheet sheet = workbook.getSheetAt(0);
            int headerRow = 0; // assumes first row is headers
            int conversationCol = 7; // Column H (0-based index)

            writer.writeNext(new String[]{"Visitor", "Associate"});

            for (int i = headerRow + 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell cell = row.getCell(conversationCol);
                if (cell == null) continue;

                String fullConversation = cell.toString().trim();
                if (fullConversation.isEmpty()) continue;

                // Split the conversation into separate lines
                String[] lines = fullConversation.split("\\r?\\n");

                boolean startCollecting = false;
                StringBuilder visitorBuffer = new StringBuilder();

                for (String rawLine : lines) {
                    String line = rawLine.trim();
                    if (line.isEmpty()) continue;

                    // Only consider lines with timestamps
                    if (!line.matches(".*\\(.*\\).*")) continue;

                    // Detect when Robot gives the signal to start collecting
                    if (line.contains("Robot:") &&
                        line.contains("Provide a quick summary of your request")) {
                        startCollecting = true;
                        continue;
                    }

                    if (!startCollecting) continue;

                    // Skip any remaining Robot lines
                    if (line.contains("Robot:")) continue;

                    // Capture Visitor messages
                    if (line.contains("Visitor:")) {
                        String message = line.substring(line.indexOf("Visitor:") + 8).trim();
                        if (visitorBuffer.length() > 0) visitorBuffer.append("\n");
                        visitorBuffer.append(message);

                    // Any other sender is treated as Associate
                    } else if (line.contains(":")) {
                        String sender = line.substring(line.indexOf(")") + 1, line.indexOf(":")).trim();
                        if (sender.equalsIgnoreCase("Visitor") || sender.equalsIgnoreCase("Robot"))
                            continue; // Skip visitors/robots here

                        String message = line.substring(line.indexOf(":") + 1).trim();

                        if (visitorBuffer.length() > 0 || !message.isEmpty()) {
                            writer.writeNext(new String[]{
                                "\"" + visitorBuffer.toString() + "\"",
                                "\"" + message + "\""
                            });
                            visitorBuffer.setLength(0);
                        }
                    }
                }

                // If visitor spoke last and no associate followed
                if (visitorBuffer.length() > 0) {
                    writer.writeNext(new String[]{
                        "\"" + visitorBuffer.toString() + "\"",
                        "\"\""
                    });
                }
            }

            System.out.println("✅ CSV file created successfully: " + outputFile);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
