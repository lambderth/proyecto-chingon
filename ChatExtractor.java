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
            int headerRow = 0; // assumes row 0 has headers
            int conversationCol = 7; // Column H = index 7 (0-based)

            writer.writeNext(new String[]{"Visitor", "Associate"});

            boolean startCollecting = false;
            StringBuilder visitorBuffer = new StringBuilder();

            for (int i = headerRow + 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;

                Cell cell = row.getCell(conversationCol);
                if (cell == null) continue;

                String line = cell.toString().trim();
                if (line.isEmpty()) continue;

                // Process only lines with timestamps
                if (!line.matches(".*\\(.*\\).*")) continue;

                // Detect start after Robot message
                if (line.contains("Robot:") &&
                    line.contains("Provide a quick summary of your request")) {
                    startCollecting = true;
                    continue;
                }

                if (!startCollecting) continue;

                // Skip any remaining Robot lines
                if (line.contains("Robot:")) continue;

                // Visitor message
                if (line.contains("Visitor:")) {
                    String message = line.substring(line.indexOf("Visitor:") + 8).trim();
                    if (visitorBuffer.length() > 0) visitorBuffer.append("\n");
                    visitorBuffer.append(message);

                // Any other sender (Associate)
                } else if (line.contains(":")) {
                    String sender = line.substring(line.indexOf(")") + 1, line.indexOf(":")).trim();
                    if (sender.equalsIgnoreCase("Visitor") || sender.equalsIgnoreCase("Robot"))
                        continue; // safety skip
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

            // If visitor spoke last
            if (visitorBuffer.length() > 0) {
                writer.writeNext(new String[]{
                    "\"" + visitorBuffer.toString() + "\"",
                    "\"\""
                });
            }

            System.out.println("✅ CSV file created successfully: " + outputFile);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
