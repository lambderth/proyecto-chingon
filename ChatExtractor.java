package datapull;

import com.opencsv.CSVWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.util.*;

public class ChatExtractor {
    public static void main(String[] args) {
        String inputFile = "chat_data.xlsx"; // your Excel file
        String outputFile = "output.csv";

        try (
            FileInputStream fis = new FileInputStream(inputFile);
            Workbook workbook = new XSSFWorkbook(fis);
            FileWriter fw = new FileWriter(outputFile);
            CSVWriter writer = new CSVWriter(fw)
        ) {
            Sheet sheet = workbook.getSheetAt(0);
            int headerRow = 0; // assumes first row has headers
            int conversationCol = 7; // Column H (0-based index)

            writer.writeNext(new String[]{"Visitor", "Associate"});

            for (int i = headerRow + 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell cell = row.getCell(conversationCol);
                if (cell == null) continue;

                String fullConversation = cell.toString().trim();
                if (fullConversation.isEmpty()) continue;

                // Split conversation into lines
                String[] lines = fullConversation.split("\\r?\\n");

                boolean startCollecting = false;
                StringBuilder visitorBuffer = new StringBuilder();
                StringBuilder associateBuffer = new StringBuilder();
                String lastSender = "";

                for (String rawLine : lines) {
                    String line = rawLine.trim();
                    if (line.isEmpty()) continue;

                    // Only lines with timestamps
                    if (!line.matches(".*\\(.*\\).*")) continue;

                    // Detect Robot trigger
                    if (line.contains("Robot:") &&
                        line.contains("Provide a quick summary of your request")) {
                        startCollecting = true;
                        continue;
                    }

                    if (!startCollecting) continue;

                    // Skip further Robot lines
                    if (line.contains("Robot:")) continue;

                    // Identify sender and message
                    String sender = "";
                    String message = "";

                    int senderStart = line.indexOf(")") + 1;
                    int senderEnd = line.indexOf(":");

                    if (senderStart > 0 && senderEnd > senderStart) {
                        sender = line.substring(senderStart, senderEnd).trim();
                        message = line.substring(senderEnd + 1).trim();
                    }

                    // Skip invalid lines
                    if (sender.isEmpty() || message.isEmpty()) continue;

                    // Handle Visitor messages
                    if (sender.equalsIgnoreCase("Visitor")) {
                        // If switching from Associate to Visitor, write current pair first
                        if (lastSender.equals("Associate") && associateBuffer.length() > 0) {
                            writer.writeNext(new String[]{
                                "\"" + visitorBuffer.toString() + "\"",
                                "\"" + associateBuffer.toString() + "\""
                            });
                            visitorBuffer.setLength(0);
                            associateBuffer.setLength(0);
                        }
                        if (visitorBuffer.length() > 0) visitorBuffer.append("\n");
                        visitorBuffer.append(message);
                        lastSender = "Visitor";

                    // Handle Associate (any non-Robot, non-Visitor)
                    } else {
                        if (lastSender.equals("Visitor") && visitorBuffer.length() > 0 && associateBuffer.length() > 0) {
                            writer.writeNext(new String[]{
                                "\"" + visitorBuffer.toString() + "\"",
                                "\"" + associateBuffer.toString() + "\""
                            });
                            visitorBuffer.setLength(0);
                            associateBuffer.setLength(0);
                        }

                        if (associateBuffer.length() > 0) associateBuffer.append("\n");
                        associateBuffer.append(message);
                        lastSender = "Associate";
                    }
                }

                // Write last pair (if any)
                if (visitorBuffer.length() > 0 || associateBuffer.length() > 0) {
                    writer.writeNext(new String[]{
                        "\"" + visitorBuffer.toString() + "\"",
                        "\"" + associateBuffer.toString() + "\""
                    });
                }
            }

            System.out.println("✅ CSV file created successfully: " + outputFile);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
