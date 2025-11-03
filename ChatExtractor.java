package datapull;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.opencsv.CSVWriter;

import java.io.*;
import java.util.*;

public class ChatExtractor {

    public static void main(String[] args) {
        String inputFile = "input.xlsx";   // Path to your Excel file
        String outputFile = "output.csv";  // Path for the generated CSV

        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook workbook = new XSSFWorkbook(fis);
             FileWriter fw = new FileWriter(outputFile);
             CSVWriter writer = new CSVWriter(fw)) {

            Sheet sheet = workbook.getSheetAt(0);
            boolean afterRobot = false;
            boolean skipHeader = true;

            List<String[]> rows = new ArrayList<>();
            StringBuilder visitorMsg = new StringBuilder();
            StringBuilder sender2Msg = new StringBuilder();
            String currentSender = "";
            String sender2Name = "Agent"; // Default until detected

            for (Row row : sheet) {
                if (skipHeader) {
                    skipHeader = false;
                    continue; // skip first header row
                }

                Cell cell = row.getCell(7); // Column H (index 7)
                if (cell == null) continue;

                String text = cell.toString().trim();
                if (text.isEmpty()) continue;

                // Wait until we pass the Robot message
                if (!afterRobot) {
                    if (text.contains("Provide a quick summary of your request")) {
                        afterRobot = true;
                    }
                    continue;
                }

                // Parse sender and message
                int colonIndex = text.indexOf(":");
                if (colonIndex == -1) continue;

                String senderPart = text.substring(0, colonIndex).trim();
                String message = text.substring(colonIndex + 1).trim();

                // Remove timestamp if present "( ... )"
                if (senderPart.contains(")")) {
                    senderPart = senderPart.substring(senderPart.indexOf(")") + 1).trim();
                }

                // Ignore robot lines
                if (senderPart.equalsIgnoreCase("Robot")) continue;

                boolean isVisitor = senderPart.equalsIgnoreCase("Visitor");

                // If we meet the first non-Visitor sender, store the name
                if (!isVisitor && sender2Name.equals("Agent")) {
                    sender2Name = senderPart;
                }

                // If sender changes, flush previous block
                if (!currentSender.isEmpty() && !currentSender.equals(senderPart)) {
                    rows.add(new String[]{visitorMsg.toString().trim(), sender2Msg.toString().trim()});
                    visitorMsg.setLength(0);
                    sender2Msg.setLength(0);
                }

                currentSender = senderPart;

                if (isVisitor) {
                    if (visitorMsg.length() > 0) visitorMsg.append("\n");
                    visitorMsg.append(message);
                } else {
                    if (sender2Msg.length() > 0) sender2Msg.append("\n");
                    sender2Msg.append(message);
                }
            }

            // Add the last set of messages
            if (visitorMsg.length() > 0 || sender2Msg.length() > 0) {
                rows.add(new String[]{visitorMsg.toString().trim(), sender2Msg.toString().trim()});
            }

            // Write CSV
            writer.writeNext(new String[]{"Visitor", sender2Name});
            writer.writeAll(rows);

            System.out.println("✅ CSV created successfully: " + outputFile);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
